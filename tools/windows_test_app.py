from __future__ import annotations

import os
import threading
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from holding79_transfer import IntegrationRunError, run_integration

APP_TITLE = "HOLDING 79 Transfer — проверка"
BUILD_MAIN_SHA = "b25da6ceb7a348187482d298f0ae38917dc7ec0e"


class App(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("760x430")
        self.minsize(700, 390)

        self.source_var = tk.StringVar()
        self.output_base_var = tk.StringVar()
        self.period_var = tk.StringVar()
        self.status_var = tk.StringVar(value="Готово к запуску")
        self.last_output_dir: Path | None = None

        root = ttk.Frame(self, padding=18)
        root.pack(fill="both", expand=True)

        ttk.Label(root, text="HOLDING 79 Transfer", font=("Segoe UI", 16, "bold")).grid(
            row=0, column=0, columnspan=3, sticky="w", pady=(0, 4)
        )
        ttk.Label(
            root,
            text="Локальная проверка ОСВ 79.2/79.3 → файлы проводок. В 1С ничего не записывает.",
        ).grid(row=1, column=0, columnspan=3, sticky="w", pady=(0, 16))

        ttk.Label(root, text="Файл ОСВ (.xlsx):").grid(row=2, column=0, sticky="w", pady=5)
        ttk.Entry(root, textvariable=self.source_var).grid(row=2, column=1, sticky="ew", padx=8)
        ttk.Button(root, text="Выбрать…", command=self.pick_source).grid(row=2, column=2, sticky="ew")

        ttk.Label(root, text="Куда сохранить результат:").grid(row=3, column=0, sticky="w", pady=5)
        ttk.Entry(root, textvariable=self.output_base_var).grid(row=3, column=1, sticky="ew", padx=8)
        ttk.Button(root, text="Выбрать…", command=self.pick_output).grid(row=3, column=2, sticky="ew")

        ttk.Label(root, text="Дата периода YYYY-MM-DD:").grid(row=4, column=0, sticky="w", pady=5)
        ttk.Entry(root, textvariable=self.period_var, width=18).grid(row=4, column=1, sticky="w", padx=8)
        ttk.Label(root, text="Можно оставить пустой — программа попробует определить из ОСВ.").grid(
            row=5, column=1, columnspan=2, sticky="w", padx=8, pady=(0, 12)
        )

        self.run_button = ttk.Button(root, text="СФОРМИРОВАТЬ ФАЙЛЫ", command=self.start_run)
        self.run_button.grid(row=6, column=0, columnspan=2, sticky="ew", pady=(8, 8))

        self.open_button = ttk.Button(root, text="Открыть папку результата", command=self.open_output)
        self.open_button.grid(row=6, column=2, sticky="ew", padx=(8, 0), pady=(8, 8))
        self.open_button.state(["disabled"])

        self.progress = ttk.Progressbar(root, mode="indeterminate")
        self.progress.grid(row=7, column=0, columnspan=3, sticky="ew", pady=(3, 10))

        ttk.Label(root, textvariable=self.status_var, wraplength=700, justify="left").grid(
            row=8, column=0, columnspan=3, sticky="nw"
        )

        ttk.Separator(root).grid(row=9, column=0, columnspan=3, sticky="ew", pady=(18, 8))
        ttk.Label(
            root,
            text=f"Тестовая сборка из принятого main: {BUILD_MAIN_SHA[:12]}…",
            foreground="#666666",
        ).grid(row=10, column=0, columnspan=3, sticky="w")

        root.columnconfigure(1, weight=1)

    def pick_source(self) -> None:
        filename = filedialog.askopenfilename(
            title="Выберите ОСВ",
            filetypes=[("Excel XLSX", "*.xlsx"), ("Все файлы", "*.*")],
        )
        if filename:
            source = Path(filename)
            self.source_var.set(str(source))
            if not self.output_base_var.get().strip():
                self.output_base_var.set(str(source.parent))

    def pick_output(self) -> None:
        folder = filedialog.askdirectory(title="Выберите папку для результата")
        if folder:
            self.output_base_var.set(folder)

    def start_run(self) -> None:
        source_text = self.source_var.get().strip()
        output_text = self.output_base_var.get().strip()
        period_text = self.period_var.get().strip()

        if not source_text:
            messagebox.showwarning(APP_TITLE, "Сначала выберите файл ОСВ .xlsx")
            return
        source = Path(source_text)
        if not source.is_file():
            messagebox.showerror(APP_TITLE, "Выбранный файл не найден.")
            return
        if source.suffix.lower() != ".xlsx":
            messagebox.showerror(APP_TITLE, "Нужен файл Excel в формате .xlsx")
            return

        output_base = Path(output_text) if output_text else source.parent
        if period_text:
            try:
                datetime.strptime(period_text, "%Y-%m-%d")
            except ValueError:
                messagebox.showerror(APP_TITLE, "Дата периода должна быть в формате YYYY-MM-DD, например 2026-03-31")
                return

        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        run_dir = output_base / f"HOLDING79_RESULT_{source.stem}_{stamp}"

        self.run_button.state(["disabled"])
        self.open_button.state(["disabled"])
        self.progress.start(12)
        self.status_var.set("Обработка ОСВ…")

        thread = threading.Thread(
            target=self._run_worker,
            args=(source, run_dir, period_text or None),
            daemon=True,
        )
        thread.start()

    def _run_worker(self, source: Path, run_dir: Path, period_end: str | None) -> None:
        try:
            result = run_integration(
                source,
                run_dir,
                period_end=period_end,
                input_name=source.name,
            )
            workbook_names = [workbook.path.name for workbook in result.exported_workbooks]
            self.after(0, self._success, run_dir, result.run_id, workbook_names, len(result.posting_rows))
        except Exception as exc:
            self.after(0, self._failure, exc)

    def _success(
        self,
        run_dir: Path,
        run_id: str,
        workbook_names: list[str],
        posting_count: int,
    ) -> None:
        self.progress.stop()
        self.run_button.state(["!disabled"])
        self.open_button.state(["!disabled"])
        self.last_output_dir = run_dir
        names = "\n".join(f"• {name}" for name in workbook_names)
        self.status_var.set(
            f"ГОТОВО. Проводок: {posting_count}. Файлов: {len(workbook_names)}.\n"
            f"Папка: {run_dir}\n"
            f"Run ID: {run_id}\n\n{names}"
        )
        messagebox.showinfo(APP_TITLE, f"Обработка завершена.\n\nРезультат:\n{run_dir}")

    def _failure(self, exc: Exception) -> None:
        self.progress.stop()
        self.run_button.state(["!disabled"])
        self.open_button.state(["disabled"])
        self.last_output_dir = None
        error_text = f"{type(exc).__name__}: {exc}"
        self.status_var.set(f"ОШИБКА / БЛОКИРОВКА:\n{error_text}")
        messagebox.showerror(
            APP_TITLE,
            "Файлы не сформированы. Программа остановилась безопасно.\n\n" + error_text,
        )

    def open_output(self) -> None:
        if self.last_output_dir is None or not self.last_output_dir.exists():
            return
        try:
            os.startfile(self.last_output_dir)  # type: ignore[attr-defined]
        except Exception as exc:
            messagebox.showerror(APP_TITLE, f"Не удалось открыть папку:\n{exc}")


if __name__ == "__main__":
    App().mainloop()
