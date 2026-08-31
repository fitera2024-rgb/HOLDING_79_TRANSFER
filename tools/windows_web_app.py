from __future__ import annotations

import json
import mimetypes
import re
import tempfile
import threading
import uuid
import webbrowser
from calendar import monthrange
from collections import defaultdict
from datetime import date
from email import policy
from email.parser import BytesParser
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import quote, urlparse
from zipfile import ZIP_DEFLATED, ZipFile

from openpyxl import load_workbook

from holding79_transfer import OUTPUT_SHEET_NAME, PostingRow, run_integration

APP_TITLE = "HOLDING 79 Transfer"
BUILD_MAIN_SHA = "b25da6ceb7a348187482d298f0ae38917dc7ec0e"
MAX_UPLOAD = 100 * 1024 * 1024
MONTH_RE = re.compile(r"^(\d{4})-(\d{2})$")
MONTHS = ("января", "февраля", "марта", "апреля", "мая", "июня", "июля", "августа", "сентября", "октября", "ноября", "декабря")

PAGE = r'''<!doctype html><html lang="ru"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>HOLDING 79 Transfer</title><style>
:root{--bg:#f3f5f8;--card:#fff;--ink:#192238;--muted:#6e7788;--line:#e0e5ed;--blue:#3157d5;--blue2:#2747b4;--green:#18794e;--red:#b42318;--warn:#9a6700;--shadow:0 20px 55px rgba(30,45,75,.10)}*{box-sizing:border-box}body{margin:0;background:radial-gradient(circle at 8% 0,rgba(49,87,213,.09),transparent 32%),var(--bg);font-family:Inter,"Segoe UI",Arial,sans-serif;color:var(--ink)}.wrap{max-width:1040px;margin:auto;padding:38px 22px 55px}.top{display:flex;justify-content:space-between;align-items:flex-start;gap:20px;margin-bottom:26px}.brand{display:flex;gap:15px;align-items:center}.logo{width:54px;height:54px;border-radius:16px;background:linear-gradient(145deg,#3d66e8,#223f9f);display:grid;place-items:center;color:#fff;font-weight:800;font-size:20px;box-shadow:var(--shadow)}h1{margin:0 0 4px;font-size:27px}.sub{font-size:14px;color:var(--muted)}.safe{padding:9px 12px;border-radius:999px;background:#edf8f1;color:var(--green);font-size:12px;font-weight:700;white-space:nowrap}.safe:before{content:"●";margin-right:7px}.card{background:var(--card);border:1px solid var(--line);border-radius:18px;box-shadow:var(--shadow);overflow:hidden}.head{padding:24px 28px 0}.head h2{margin:0;font-size:19px}.head p{margin:7px 0 0;color:var(--muted);font-size:13px;line-height:1.5}.form{padding:24px 28px 28px;display:grid;grid-template-columns:1fr 245px;gap:18px}label{display:block;font-size:13px;font-weight:700;margin-bottom:8px}input{width:100%;min-height:48px;border:1px solid #cfd6e2;border-radius:12px;background:#fff;color:var(--ink);font:inherit;outline:none}input[type=file]{padding:10px 12px}input[type=month]{padding:0 14px}input:focus{border-color:var(--blue);box-shadow:0 0 0 3px rgba(49,87,213,.12)}.hint{font-size:12px;color:var(--muted);margin-top:7px}.actions{grid-column:1/-1;border-top:1px solid var(--line);padding-top:18px;display:flex;justify-content:space-between;align-items:center;gap:15px}.status{font-size:13px;color:var(--muted)}button,a.btn{font:inherit;font-weight:700;border:0;border-radius:12px;text-decoration:none;cursor:pointer}.primary{background:var(--blue);color:#fff;padding:13px 20px;min-width:230px}.primary:hover{background:var(--blue2)}.primary:disabled{opacity:.55;cursor:default}.secondary{background:#edf1f7;color:#455066;padding:10px 15px}.progress{display:none;grid-column:1/-1;height:6px;border-radius:99px;overflow:hidden;background:#edf0f4}.progress.on{display:block}.progress span{display:block;width:36%;height:100%;background:var(--blue);border-radius:99px;animation:p 1.1s infinite ease-in-out}@keyframes p{from{transform:translateX(-110%)}to{transform:translateX(300%)}}.msg{display:none;grid-column:1/-1;padding:14px 16px;border-radius:12px;font-size:13px;line-height:1.5}.msg.err{display:block;background:#fff0ef;color:var(--red);border:1px solid #ffd3cf}.result{display:none;margin-top:24px}.result.on{display:block}.inside{padding:24px 28px 28px}.metrics{display:grid;grid-template-columns:repeat(4,1fr);gap:12px}.metric{padding:15px;background:#f8f9fc;border:1px solid var(--line);border-radius:14px}.metric b{display:block;font-size:24px;margin-bottom:3px}.metric span{font-size:11px;color:var(--muted)}.files{margin-top:18px;border:1px solid var(--line);border-radius:14px;overflow:hidden}.row{display:flex;justify-content:space-between;gap:16px;align-items:center;padding:13px 15px;border-top:1px solid var(--line)}.row:first-child{border-top:0}.name{min-width:0}.name b{display:block;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;font-size:13px}.name small{color:var(--muted)}.badge{display:inline-block;margin-left:8px;padding:3px 7px;border-radius:999px;background:#fff3cd;color:var(--warn);font-size:10px;font-weight:800}.download{color:var(--blue);background:#edf2ff;padding:8px 11px;white-space:nowrap}.all{display:flex;justify-content:flex-end;margin-top:16px}.foot{display:flex;justify-content:space-between;align-items:center;margin-top:18px;color:#8991a1;font-size:11px}@media(max-width:720px){.wrap{padding:24px 14px 40px}.top{flex-direction:column}.form{grid-template-columns:1fr;padding:20px}.actions{grid-column:1;flex-direction:column;align-items:stretch}.primary{width:100%}.metrics{grid-template-columns:repeat(2,1fr)}}
</style></head><body><div class="wrap"><div class="top"><div class="brand"><div class="logo">79</div><div><h1>HOLDING 79 Transfer</h1><div class="sub">Перенос остатков 79.2 / 79.3 на 79.1</div></div></div><div class="safe">Локально · без записи в 1С</div></div>
<section class="card"><div class="head"><h2>Сформировать файлы проводок</h2><p>Выберите ОСВ и месяц. Расчёт выполняется локально. Сформированные файлы можно скачать по организациям или одним архивом.</p></div><form class="form" id="f"><div><label>Файл ОСВ</label><input name="source" type="file" accept=".xlsx" required><div class="hint">Excel .xlsx</div></div><div><label>Период</label><input name="period" type="month" value="{{MONTH}}" required><div class="hint">Выберите месяц расчёта</div></div><div class="actions"><div class="status" id="status">Готово к запуску</div><button class="primary" id="go">Сформировать файлы</button></div><div class="progress" id="prog"><span></span></div><div class="msg" id="msg"></div></form></section>
<section class="result" id="res"><div class="card"><div class="head"><h2>Готово</h2><p id="per"></p></div><div class="inside"><div class="metrics"><div class="metric"><b id="src">0</b><span>остатков к переносу</span></div><div class="metric"><b id="post">0</b><span>проводок</span></div><div class="metric"><b id="books">0</b><span>файлов организаций</span></div><div class="metric"><b id="sp">0</b><span>файлов СПОРНО</span></div></div><div class="files" id="files"></div><div class="all"><a class="btn primary" id="zip">Скачать всё ZIP</a></div></div></div></section>
<div class="foot"><span>Тестовый веб-интерфейс · main {{SHA}}</span><button class="secondary" id="close">Закрыть программу</button></div></div><script>
const f=document.getElementById('f'),go=document.getElementById('go'),prog=document.getElementById('prog'),msg=document.getElementById('msg'),res=document.getElementById('res'),status=document.getElementById('status');function esc(s){const d=document.createElement('div');d.textContent=s??'';return d.innerHTML}f.addEventListener('submit',async e=>{e.preventDefault();msg.className='msg';msg.textContent='';res.classList.remove('on');go.disabled=true;prog.classList.add('on');status.textContent='Обработка ОСВ…';try{const r=await fetch('/api/run',{method:'POST',body:new FormData(f)}),d=await r.json();if(!r.ok||!d.ok)throw new Error(d.error||'Не удалось сформировать файлы.');document.getElementById('src').textContent=d.source_rows;document.getElementById('post').textContent=d.posting_rows;document.getElementById('books').textContent=d.workbooks.length;document.getElementById('sp').textContent=d.disputed_workbooks;document.getElementById('per').textContent='Период: '+d.period_label;document.getElementById('zip').href=d.zip_url;document.getElementById('files').innerHTML=d.workbooks.map(x=>`<div class="row"><div class="name"><b>${esc(x.name)}${x.disputed?'<span class="badge">СПОРНО</span>':''}</b><small>${x.rows} проводок</small></div><a class="btn download" href="${x.url}">Скачать</a></div>`).join('');res.classList.add('on');status.textContent='Файлы сформированы';res.scrollIntoView({behavior:'smooth'})}catch(x){msg.textContent=x.message||String(x);msg.className='msg err';status.textContent='Файлы не сформированы'}finally{prog.classList.remove('on');go.disabled=false}});document.getElementById('close').onclick=async()=>{if(!confirm('Закрыть HOLDING 79 Transfer?'))return;try{await fetch('/api/shutdown',{method:'POST'})}catch(_){ }document.body.innerHTML='<div style="font:16px Segoe UI;padding:40px">Программа закрыта. Эту вкладку можно закрыть.</div>'};
</script></body></html>'''


def period_end(text: str) -> tuple[date, str]:
    match = MONTH_RE.fullmatch(text.strip())
    if not match:
        raise ValueError("Выберите месяц расчёта.")
    year, month = int(match.group(1)), int(match.group(2))
    if not 1 <= month <= 12:
        raise ValueError("Некорректный месяц периода.")
    end = date(year, month, monthrange(year, month)[1])
    return end, f"{MONTHS[month - 1][:-1] if MONTHS[month - 1].endswith('я') else MONTHS[month - 1]} {year}"


def multipart(body: bytes, content_type: str) -> tuple[dict[str, str], dict[str, tuple[str, bytes]]]:
    prefix = f"Content-Type: {content_type}\r\nMIME-Version: 1.0\r\n\r\n".encode()
    message = BytesParser(policy=policy.default).parsebytes(prefix + body)
    fields: dict[str, str] = {}
    files: dict[str, tuple[str, bytes]] = {}
    if not message.is_multipart():
        return fields, files
    for part in message.iter_parts():
        if part.get_content_disposition() != "form-data":
            continue
        name = part.get_param("name", header="content-disposition")
        if not name:
            continue
        payload = part.get_payload(decode=True) or b""
        filename = part.get_filename()
        if filename:
            files[name] = (Path(filename).name, payload)
        else:
            fields[name] = payload.decode(part.get_content_charset() or "utf-8", errors="replace")
    return fields, files


def clean_content(row: PostingRow) -> str:
    side = "дебетового" if row.side and row.side.value == "DEBIT" else "кредитового"
    account = row.source_account.value if row.source_account else "79.x"
    period = f"{MONTHS[row.period_end.month - 1]} {row.period_end.year} года" if row.period_end else "выбранного периода"
    parts = [f"Перенос конечного {side} остатка по счету {account} на 79.1 за {period}.", f"Организация источника: {row.source_organization}."]
    if row.source_department:
        parts.append(f"ЦФО: {row.source_department}.")
    if row.source_supplier_rvp:
        parts.append(f"РВП: {row.source_supplier_rvp}.")
    if row.document_organization != row.source_organization:
        parts.append(f"Организация проводки: {row.document_organization}.")
    return " ".join(parts)


def verify_only_content_changed(original: Path, cleaned: Path, content_col: int) -> None:
    left, right = load_workbook(original, data_only=False), load_workbook(cleaned, data_only=False)
    try:
        if left.sheetnames != right.sheetnames:
            raise RuntimeError("Изменился состав листов выходного файла.")
        for name in left.sheetnames:
            a, b = left[name], right[name]
            if a.max_row != b.max_row or a.max_column != b.max_column:
                raise RuntimeError("Изменилась размерность выходного файла.")
            for r in range(1, a.max_row + 1):
                for c in range(1, a.max_column + 1):
                    if name == OUTPUT_SHEET_NAME and r >= 2 and c == content_col:
                        continue
                    if a.cell(r, c).value != b.cell(r, c).value:
                        raise RuntimeError("Изменились данные проводки вне колонки Содержание.")
    finally:
        left.close(); right.close()


def prepare_downloads(result, folder: Path):
    folder.mkdir(parents=True, exist_ok=False)
    grouped: dict[tuple[object, str], list[PostingRow]] = defaultdict(list)
    for row in result.posting_rows:
        grouped[(row.period_end, row.document_organization)].append(row)
    prepared = []
    for exported in result.exported_workbooks:
        rows = grouped[(exported.document_date, exported.document_organization)]
        if len(rows) != exported.row_count:
            raise RuntimeError("Количество проводок не совпало при подготовке файла.")
        book = load_workbook(exported.path)
        try:
            if OUTPUT_SHEET_NAME not in book.sheetnames:
                raise RuntimeError("Нет листа Загрузка_A_AA.")
            sheet = book[OUTPUT_SHEET_NAME]
            headers = [sheet.cell(1, c).value for c in range(1, sheet.max_column + 1)]
            try:
                content_col = headers.index("Содержание") + 1
            except ValueError as exc:
                raise RuntimeError("Нет колонки Содержание.") from exc
            if sheet.max_row - 1 != len(rows):
                raise RuntimeError("Количество строк файла не совпало с проводками.")
            for excel_row, posting in enumerate(rows, 2):
                sheet.cell(excel_row, content_col).value = clean_content(posting)
            target = folder / exported.path.name
            book.save(target)
        finally:
            book.close()
        verify_only_content_changed(exported.path, target, content_col)
        prepared.append((target, exported))
    return prepared


class Server(ThreadingHTTPServer):
    daemon_threads = True
    def __init__(self, address, root: Path):
        super().__init__(address, Handler); self.root = root; self.downloads = {}
    def add_download(self, path: Path) -> str:
        token = uuid.uuid4().hex; self.downloads[token] = path; return f"/download/{token}"


class Handler(BaseHTTPRequestHandler):
    server: Server
    def log_message(self, *_):
        return
    def send_bytes(self, data: bytes, content_type: str, status=200):
        self.send_response(status); self.send_header("Content-Type", content_type); self.send_header("Content-Length", str(len(data))); self.send_header("Cache-Control", "no-store"); self.end_headers(); self.wfile.write(data)
    def send_json(self, data: dict, status=200):
        self.send_bytes(json.dumps(data, ensure_ascii=False).encode(), "application/json; charset=utf-8", status)
    def do_GET(self):
        path = urlparse(self.path).path
        if path == "/":
            page = PAGE.replace("{{MONTH}}", date.today().strftime("%Y-%m")).replace("{{SHA}}", BUILD_MAIN_SHA[:12] + "…")
            return self.send_bytes(page.encode(), "text/html; charset=utf-8")
        if path.startswith("/download/"):
            item = self.server.downloads.get(path.rsplit("/", 1)[-1])
            if not item or not item.is_file():
                return self.send_json({"ok": False, "error": "Файл больше недоступен."}, HTTPStatus.NOT_FOUND)
            data = item.read_bytes(); ctype = mimetypes.guess_type(item.name)[0] or "application/octet-stream"
            self.send_response(200); self.send_header("Content-Type", ctype); self.send_header("Content-Length", str(len(data))); self.send_header("Content-Disposition", f"attachment; filename*=UTF-8''{quote(item.name)}"); self.send_header("Cache-Control", "no-store"); self.end_headers(); self.wfile.write(data); return
        self.send_json({"ok": False, "error": "Страница не найдена."}, HTTPStatus.NOT_FOUND)
    def do_POST(self):
        path = urlparse(self.path).path
        if path == "/api/shutdown":
            self.send_json({"ok": True}); threading.Thread(target=self.server.shutdown, daemon=True).start(); return
        if path != "/api/run":
            return self.send_json({"ok": False, "error": "Команда не найдена."}, HTTPStatus.NOT_FOUND)
        try:
            length = int(self.headers.get("Content-Length", "0"))
            if length <= 0 or length > MAX_UPLOAD:
                raise ValueError("Файл пустой или больше 100 МБ.")
            ctype = self.headers.get("Content-Type", "")
            if "multipart/form-data" not in ctype:
                raise ValueError("Не удалось прочитать файл.")
            fields, files = multipart(self.rfile.read(length), ctype)
            if "source" not in files:
                raise ValueError("Выберите файл ОСВ .xlsx.")
            source_name, source_bytes = files["source"]
            if Path(source_name).suffix.lower() != ".xlsx":
                raise ValueError("Нужен файл Excel .xlsx.")
            end, label = period_end(fields.get("period", ""))
            root = self.server.root / uuid.uuid4().hex; root.mkdir()
            source = root / "source.xlsx"; source.write_bytes(source_bytes)
            result = run_integration(source, root / "result", period_end=end, input_name=source_name)
            prepared = prepare_downloads(result, root / "download")
            books = []; disputed = 0
            for clean, exported in prepared:
                is_sporno = "_СПОРНО" in clean.stem; disputed += int(is_sporno)
                books.append({"name": clean.name, "rows": exported.row_count, "disputed": is_sporno, "url": self.server.add_download(clean)})
            zip_path = root / "HOLDING79_TRANSFER_RESULT.zip"
            with ZipFile(zip_path, "w", ZIP_DEFLATED) as z:
                for clean, _ in prepared:
                    z.write(clean, clean.name)
                control = root / "result" / "run_control.xlsx"
                if control.is_file(): z.write(control, "Контроль_расчета.xlsx")
            self.send_json({"ok": True, "period_label": label, "source_rows": len(result.normalized_balances), "posting_rows": len(result.posting_rows), "disputed_workbooks": disputed, "workbooks": books, "zip_url": self.server.add_download(zip_path)})
        except Exception as exc:
            self.send_json({"ok": False, "error": str(exc).strip() or type(exc).__name__}, HTTPStatus.BAD_REQUEST)


def main():
    with tempfile.TemporaryDirectory(prefix="holding79-web-") as temp:
        server = Server(("127.0.0.1", 0), Path(temp)); host, port = server.server_address
        threading.Timer(.5, lambda: webbrowser.open(f"http://{host}:{port}/", new=1)).start()
        server.serve_forever(.25); server.server_close()


if __name__ == "__main__":
    main()
