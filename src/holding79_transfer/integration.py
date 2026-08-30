"""Deterministic local orchestration for one 79.x transfer run.

The integration layer owns run artifacts and controls only.  Accounting
semantics remain in :mod:`holding79_transfer.parser`, ``transfer`` and
``exporter``; this module composes those accepted components and fails closed
before publishing a run directory when a mandatory control does not pass.
"""

from __future__ import annotations

import hashlib
import json
import re
import tempfile
from collections import Counter, defaultdict
from collections.abc import Iterable, Mapping
from dataclasses import dataclass, field, replace
from datetime import date, datetime
from decimal import Decimal
from io import BytesIO
from pathlib import Path
from typing import Any

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font
from openpyxl.worksheet.worksheet import Worksheet

from .config import TransferConfig
from .exporter import ExportedWorkbook, export_posting_rows, validate_workbook_round_trip
from .models import (
    CONTRACT_VERSION,
    RULES_VERSION,
    BalanceStatus,
    NormalizedBalance,
    PostingRow,
)
from .output import OUTPUT_HEADERS, OUTPUT_SHEET_NAME, OutputAdapterConfig
from .parser import GroupedOsvParser, ParserDiagnostic, ParseResult
from .transfer import TransferBatchResult, TransferEngine, TransferResult

RUN_ARTIFACT_NAMES: tuple[str, ...] = (
    "input_manifest.json",
    "normalized_balances.jsonl",
    "posting_rows.jsonl",
    "summary.json",
    "run_control.xlsx",
    "export_manifest.json",
)

CONTROL_SHEET_NAMES: tuple[str, ...] = (
    "Итоги",
    "Параметры_запуска",
    "Остатки_79",
    "Готовые_проводки",
    "Блокировки",
    "Контроль_до_после",
    "Исходные_строки",
    "Проверка_экспорта",
)

_SAFE_ID_RE = re.compile(r"[A-Za-z0-9][A-Za-z0-9_.-]{0,63}\Z")
_ZERO = Decimal(0)


class IntegrationRunError(RuntimeError):
    """Raised when a local run cannot be published as a valid run."""


@dataclass(frozen=True)
class IntegrationRunConfig:
    """Explicit configuration for one deterministic local run."""

    period_end: date | datetime | str | None = None
    transfer_config: TransferConfig = field(default_factory=TransferConfig)
    adapter_config: OutputAdapterConfig = field(default_factory=OutputAdapterConfig)
    input_name: str = "osv.xlsx"
    run_id: str | None = None
    sheet_names: tuple[str, ...] | None = None

    def __post_init__(self) -> None:
        if not isinstance(self.transfer_config, TransferConfig):
            raise TypeError("transfer_config must be a TransferConfig")
        if not isinstance(self.adapter_config, OutputAdapterConfig):
            raise TypeError("adapter_config must be an OutputAdapterConfig")
        if not isinstance(self.input_name, str) or not self.input_name.strip():
            raise ValueError("input_name must be a non-empty string")
        input_name = self.input_name.strip()
        if input_name in {".", ".."} or "/" in input_name or "\\" in input_name:
            raise ValueError("input_name must be a file name, not a path")
        object.__setattr__(self, "input_name", input_name)

        if self.run_id is not None:
            if not isinstance(self.run_id, str) or not _SAFE_ID_RE.fullmatch(self.run_id.strip()):
                raise ValueError("run_id must be a safe non-empty identifier")
            object.__setattr__(self, "run_id", self.run_id.strip())

        if self.sheet_names is not None:
            names = tuple(name.strip() if isinstance(name, str) else name for name in self.sheet_names)
            if not names or any(not isinstance(name, str) or not name.strip() for name in names):
                raise ValueError("sheet_names must contain non-empty names")
            if len(set(names)) != len(names):
                raise ValueError("sheet_names must not contain duplicates")
            object.__setattr__(self, "sheet_names", names)


@dataclass(frozen=True)
class IntegrationRunResult:
    """Published run metadata and the canonical financial values."""

    output_dir: Path
    run_id: str
    status: str
    normalized_balances: tuple[NormalizedBalance, ...]
    posting_rows: tuple[PostingRow, ...]
    diagnostics: tuple[ParserDiagnostic, ...]
    exported_workbooks: tuple[ExportedWorkbook, ...]

    @property
    def is_success(self) -> bool:
        return self.status == "SUCCESS"

    @property
    def artifacts(self) -> tuple[Path, ...]:
        export_paths = tuple(workbook.path for workbook in self.exported_workbooks)
        return tuple(self.output_dir / name for name in RUN_ARTIFACT_NAMES) + export_paths


def _decimal_text(value: Decimal) -> str:
    if value == _ZERO:
        return "0"
    return format(value, "f")


def _canonical_json(value: Any) -> str:
    return json.dumps(value, ensure_ascii=False, sort_keys=True, separators=(",", ":"))


def _write_json(path: Path, value: Any) -> None:
    path.write_text(
        json.dumps(value, ensure_ascii=False, sort_keys=True, indent=2) + "\n",
        encoding="utf-8",
        newline="\n",
    )


def _write_jsonl(path: Path, records: Iterable[Mapping[str, Any]]) -> None:
    with path.open("w", encoding="utf-8", newline="\n") as stream:
        for record in records:
            stream.write(_canonical_json(dict(record)))
            stream.write("\n")


def _balance_record(balance: NormalizedBalance) -> dict[str, Any]:
    return {
        "period_end": balance.period_end.isoformat() if balance.period_end else None,
        "organization": balance.organization,
        "source_account": balance.source_account.value if balance.source_account else None,
        "department": balance.department,
        "supplier_rvp": balance.supplier_rvp,
        "ending_debit": _decimal_text(balance.ending_debit),
        "ending_credit": _decimal_text(balance.ending_credit),
        "source_excel_row_ref": balance.source_excel_row_ref,
        "status": balance.status.value,
        "block_reason": balance.block_reason.value if balance.block_reason else None,
    }


def _posting_record(row: PostingRow) -> dict[str, Any]:
    return {
        "period_end": row.period_end.isoformat() if row.period_end else None,
        "document_organization": row.document_organization,
        "debit_account": row.debit_account.value,
        "debit_department": row.debit_department,
        "debit_supplier_rvp": row.debit_supplier_rvp,
        "credit_account": row.credit_account.value,
        "credit_department": row.credit_department,
        "credit_supplier_rvp": row.credit_supplier_rvp,
        "amount": _decimal_text(row.amount),
        "source_organization": row.source_organization,
        "source_account": row.source_account.value if row.source_account else None,
        "source_department": row.source_department,
        "source_supplier_rvp": row.source_supplier_rvp,
        "source_excel_row_ref": row.source_excel_row_ref,
        "financial_record_id": row.financial_record_id,
        "side": row.side.value if row.side else None,
        "rules_version": row.rules_version,
    }


def _balance_sort_key(balance: NormalizedBalance) -> tuple[str, ...]:
    return (
        balance.period_end.isoformat() if balance.period_end else "",
        balance.organization,
        balance.source_account.value if balance.source_account else "",
        balance.department,
        balance.supplier_rvp,
        balance.source_excel_row_ref or "",
        _decimal_text(balance.ending_debit),
        _decimal_text(balance.ending_credit),
    )


def _posting_sort_key(row: PostingRow) -> tuple[str, ...]:
    return (
        row.period_end.isoformat() if row.period_end else "",
        row.document_organization,
        row.financial_record_id,
        row.source_excel_row_ref or "",
        row.side.value if row.side else "",
        row.debit_account.value,
        row.credit_account.value,
        row.debit_department,
        row.credit_department,
        row.debit_supplier_rvp,
        row.credit_supplier_rvp,
        _decimal_text(row.amount),
    )


def _posting_identity(row: PostingRow) -> tuple[str, ...]:
    record = _posting_record(row)
    return tuple(str(record[key]) for key in record)


def _diagnostic_record(diagnostic: ParserDiagnostic) -> dict[str, Any]:
    return {
        "sheet_name": diagnostic.sheet_name,
        "excel_row": diagnostic.excel_row,
        "excel_column": diagnostic.excel_column,
        "source_excel_row_ref": (
            f"{diagnostic.sheet_name}!R{diagnostic.excel_row}"
            if diagnostic.sheet_name and diagnostic.excel_row is not None
            else None
        ),
        "code": diagnostic.code.value,
        "reason": diagnostic.reason.value,
        "message": diagnostic.message,
        "status": diagnostic.status.value,
    }


def _source_ref(diagnostic: ParserDiagnostic) -> str | None:
    if diagnostic.sheet_name and diagnostic.excel_row is not None:
        return f"{diagnostic.sheet_name}!R{diagnostic.excel_row}"
    return None


def _diagnostic_sort_key(diagnostic: ParserDiagnostic) -> tuple[str, int, int, str, str]:
    return (
        diagnostic.sheet_name or "",
        diagnostic.excel_row or 0,
        diagnostic.excel_column or 0,
        diagnostic.code.value,
        diagnostic.message,
    )


def _load_input_workbook(source: Any) -> tuple[Workbook, bool]:
    if isinstance(source, Workbook):
        return source, False
    if isinstance(source, Worksheet):
        return source.parent, False
    try:
        if isinstance(source, (str, Path)):
            return load_workbook(source, data_only=True), True
        if isinstance(source, (bytes, bytearray)):
            return load_workbook(BytesIO(bytes(source)), data_only=True), True
        if hasattr(source, "read"):
            stream = source
            if hasattr(stream, "seek"):
                stream.seek(0)
            return load_workbook(BytesIO(stream.read()), data_only=True), True
    except Exception as exc:  # pragma: no cover - openpyxl version-specific failures
        raise IntegrationRunError("input source is not a readable XLSX workbook") from exc
    raise TypeError(f"unsupported XLSX source type: {type(source).__name__}")


def _sheet_names(workbook: Workbook, configured: tuple[str, ...] | None) -> tuple[str, ...]:
    if configured is not None:
        missing = [name for name in configured if name not in workbook.sheetnames]
        if missing:
            raise IntegrationRunError(f"configured worksheet not found: {', '.join(missing)}")
        return configured
    if len(workbook.sheetnames) == 1:
        return (workbook.sheetnames[0],)
    preferred = tuple(
        name
        for name in workbook.sheetnames
        if "осв" in name.casefold() or "osv" in name.casefold()
    )
    return preferred or tuple(workbook.sheetnames)


def _parse_sheets(
    workbook: Workbook,
    config: IntegrationRunConfig,
) -> tuple[tuple[NormalizedBalance, ...], tuple[ParserDiagnostic, ...], list[dict[str, Any]]]:
    balances: list[NormalizedBalance] = []
    diagnostics: list[ParserDiagnostic] = []
    sheet_records: list[dict[str, Any]] = []
    for name in _sheet_names(workbook, config.sheet_names):
        result: ParseResult = GroupedOsvParser(
            config.period_end,
            sheet_name=name,
        ).parse(workbook)
        balances.extend(result.balances)
        diagnostics.extend(result.diagnostics)
        sheet_records.append(
            {
                "sheet_name": name,
                "status": result.status.value,
                "normalized_balance_count": len(result.balances),
                "diagnostic_count": len(result.diagnostics),
            }
        )
    ordered_balances = tuple(sorted(balances, key=_balance_sort_key))
    ordered_diagnostics = tuple(sorted(diagnostics, key=_diagnostic_sort_key))
    if any(diagnostic.excel_row is None for diagnostic in ordered_diagnostics):
        codes = ", ".join(diagnostic.code.value for diagnostic in ordered_diagnostics)
        raise IntegrationRunError(f"mandatory parser/source control failed: {codes}")
    if not ordered_balances:
        raise IntegrationRunError("mandatory parser/source control failed: no normalized balances")
    return ordered_balances, ordered_diagnostics, sheet_records


def _config_identity(config: IntegrationRunConfig) -> dict[str, Any]:
    adapter = config.adapter_config
    return {
        "period_end": str(config.period_end) if config.period_end is not None else None,
        "transfer": {
            "manager_organization": config.transfer_config.manager_organization,
            "manager_financial_department": config.transfer_config.manager_financial_department,
            "rules_version": config.transfer_config.rules_version,
        },
        "adapter": {
            "operation_type": adapter.operation_type,
            "currency": adapter.currency,
            "debit_activity": adapter.debit_activity,
            "credit_activity": adapter.credit_activity,
            "contract_version": adapter.contract_version,
            "rules_version": adapter.rules_version,
        },
        "input_name": config.input_name,
        "sheet_names": list(config.sheet_names) if config.sheet_names else None,
    }


def _run_id(
    config: IntegrationRunConfig,
    balances: tuple[NormalizedBalance, ...],
    diagnostics: tuple[ParserDiagnostic, ...],
) -> tuple[str, str]:
    fingerprint_payload = {
        "contract_version": CONTRACT_VERSION,
        "config": _config_identity(config),
        "normalized_balances": [_balance_record(balance) for balance in balances],
        "diagnostics": [_diagnostic_record(diagnostic) for diagnostic in diagnostics],
    }
    fingerprint = hashlib.sha256(_canonical_json(fingerprint_payload).encode("utf-8")).hexdigest()
    configured_ids = {
        value
        for value in (config.run_id, config.adapter_config.run_id)
        if value
    }
    if len(configured_ids) > 1:
        raise ValueError("run_id and adapter_config.run_id must match when both are supplied")
    explicit = next(iter(configured_ids), None)
    if explicit is not None and not _SAFE_ID_RE.fullmatch(explicit):
        raise ValueError("adapter_config.run_id must be a safe non-empty identifier")
    return explicit or f"run-{fingerprint[:16]}", fingerprint


def _control_failure(message: str) -> IntegrationRunError:
    return IntegrationRunError(f"mandatory control failed: {message}")


def _financial_controls(
    balances: tuple[NormalizedBalance, ...],
    batch: TransferBatchResult,
    diagnostics: tuple[ParserDiagnostic, ...],
    transfer_config: TransferConfig,
) -> tuple[list[dict[str, Any]], list[dict[str, Any]], dict[str, Any]]:
    if batch.status is BalanceStatus.BLOCKED:
        raise _control_failure(batch.message or "transfer batch is blocked")

    transfer_by_ref: dict[str, TransferResult] = {}
    expected_rows: list[PostingRow] = []
    control_rows: list[dict[str, Any]] = []
    source_total = _ZERO
    gk_total = _ZERO

    for balance, transfer in zip(balances, batch.transfers, strict=True):
        source_ref = balance.source_excel_row_ref or ""
        if transfer.status is BalanceStatus.NO_ACTION:
            control_rows.append(
                {
                    "source_excel_row_ref": source_ref,
                    "period_end": balance.period_end.isoformat() if balance.period_end else "",
                    "organization": balance.organization,
                    "source_account": balance.source_account.value if balance.source_account else "",
                    "department": balance.department,
                    "supplier_rvp": balance.supplier_rvp,
                    "ending_debit_before": _decimal_text(balance.ending_debit),
                    "ending_credit_before": _decimal_text(balance.ending_credit),
                    "ending_debit_after": _decimal_text(balance.ending_debit),
                    "ending_credit_after": _decimal_text(balance.ending_credit),
                    "source_posting_count": 0,
                    "source_amount": "0",
                    "gk_amount": "0",
                    "source_effect_zero": "NO_ACTION",
                    "amounts_match": "NO_ACTION",
                    "direction_correct": "NO_ACTION",
                    "source_consumed_once": "NO_ACTION",
                    "status": BalanceStatus.NO_ACTION.value,
                }
            )
            continue
        if transfer.status is not BalanceStatus.ACTIONABLE or len(transfer.rows) != 2:
            raise _control_failure(f"expected two postings for {source_ref}")
        if transfer.source_effect is None or not transfer.source_effect.passed:
            raise _control_failure(f"source-effect control failed for {source_ref}")

        transfer_by_ref[source_ref] = transfer
        source_row, gk_row = transfer.rows
        if (
            source_row.document_organization != balance.organization
            or gk_row.document_organization != transfer_config.manager_organization
        ):
            raise _control_failure(f"source/GK document organizations do not reconcile for {source_ref}")
        expected_rows.extend(transfer.rows)
        source_total += source_row.amount
        gk_total += gk_row.amount
        effect = transfer.source_effect
        control_rows.append(
            {
                "source_excel_row_ref": source_ref,
                "period_end": balance.period_end.isoformat() if balance.period_end else "",
                "organization": balance.organization,
                "source_account": balance.source_account.value if balance.source_account else "",
                "department": balance.department,
                "supplier_rvp": balance.supplier_rvp,
                "ending_debit_before": _decimal_text(effect.ending_debit_before),
                "ending_credit_before": _decimal_text(effect.ending_credit_before),
                "ending_debit_after": _decimal_text(effect.ending_debit_after),
                "ending_credit_after": _decimal_text(effect.ending_credit_after),
                "source_posting_count": effect.source_posting_count,
                "source_amount": _decimal_text(effect.source_amount or _ZERO),
                "gk_amount": _decimal_text(effect.gk_amount or _ZERO),
                "source_effect_zero": str(effect.source_effect_zero),
                "amounts_match": str(effect.amounts_match),
                "direction_correct": str(effect.direction_correct),
                "source_consumed_once": str(effect.source_consumed_once),
                "status": BalanceStatus.ACTIONABLE.value,
            }
        )

    actual_rows = tuple(batch.rows)
    expected_counts = Counter(_posting_identity(row) for row in expected_rows)
    actual_counts = Counter(_posting_identity(row) for row in actual_rows)
    if actual_counts != expected_counts:
        raise _control_failure(
            f"posting reconciliation mismatch: missing={sum((expected_counts - actual_counts).values())}, "
            f"extra={sum((actual_counts - expected_counts).values())}"
        )

    actual_refs = {row.source_excel_row_ref for row in actual_rows}
    blocked_refs = {_source_ref(diagnostic) for diagnostic in diagnostics}
    blocked_refs.discard(None)
    if actual_refs.intersection(blocked_refs):
        raise _control_failure("blocked source row produced a financial PostingRow")

    for balance in balances:
        if balance.status is not BalanceStatus.ACTIONABLE:
            continue
        source_ref = balance.source_excel_row_ref or ""
        transfer = transfer_by_ref.get(source_ref)
        matching = [row for row in actual_rows if row.source_excel_row_ref == source_ref]
        if transfer is None or len(matching) != 2:
            raise _control_failure(f"source row does not reconcile 1:1 to two postings: {source_ref}")
        if sum((row.amount for row in matching), _ZERO) != balance.amount * 2:
            raise _control_failure(f"posting amount mismatch for {source_ref}")

    if source_total != gk_total:
        raise _control_failure(
            f"source-org/GK total mismatch: {source_total} != {gk_total}"
        )
    blocked_output_passed = all(
        row.source_excel_row_ref not in blocked_refs for row in actual_rows
    )
    if not blocked_output_passed:
        raise _control_failure("blocked parser/source row has financial output")

    checks = [
        {"control": "source_row_reconciliation", "status": "PASS", "value": len(expected_rows)},
        {
            "control": "source_effect_zero",
            "status": "PASS",
            "value": sum(
                1
                for control in control_rows
                if control["status"] == BalanceStatus.ACTIONABLE.value
                and control["ending_debit_after"] == "0"
                and control["ending_credit_after"] == "0"
            ),
        },
        {
            "control": "source_org_gk_totals_match",
            "status": "PASS",
            "value": _decimal_text(source_total),
        },
        {
            "control": "blocked_rows_no_financial_output",
            "status": "PASS",
            "value": len(blocked_refs),
        },
    ]
    metrics = {
        "source_org_total": _decimal_text(source_total),
        "gk_total": _decimal_text(gk_total),
        "posting_row_count": len(actual_rows),
        "actionable_source_row_count": sum(
            balance.status is BalanceStatus.ACTIONABLE for balance in balances
        ),
        "no_action_source_row_count": sum(
            balance.status is BalanceStatus.NO_ACTION for balance in balances
        ),
        "blocked_source_row_count": len(blocked_refs),
    }
    return control_rows, checks, metrics


def _write_control_sheet(
    workbook: Workbook,
    name: str,
    headers: tuple[str, ...],
    rows: Iterable[Iterable[Any]],
) -> None:
    worksheet = workbook.create_sheet(name)
    worksheet.append(list(headers))
    for cell in worksheet[1]:
        cell.font = Font(bold=True)
    for row in rows:
        worksheet.append(["" if value is None else value for value in row])
    worksheet.freeze_panes = "A2"
    worksheet.auto_filter.ref = f"A1:{chr(64 + min(len(headers), 26))}{max(1, worksheet.max_row)}"
    worksheet.sheet_view.showGridLines = False


def _write_run_control(
    path: Path,
    *,
    run_id: str,
    config: IntegrationRunConfig,
    adapter_config: OutputAdapterConfig,
    balances: tuple[NormalizedBalance, ...],
    posting_rows: tuple[PostingRow, ...],
    diagnostics: tuple[ParserDiagnostic, ...],
    control_rows: list[dict[str, Any]],
    checks: list[dict[str, Any]],
    export_rows: list[dict[str, Any]],
    metrics: dict[str, Any],
) -> None:
    workbook = Workbook()
    workbook.remove(workbook.active)

    summary_rows = [
        ("run_id", "PASS", run_id, "non-empty deterministic run id"),
        ("contract_version", "PASS", CONTRACT_VERSION, CONTRACT_VERSION),
        ("rules_version", "PASS", config.transfer_config.rules_version, RULES_VERSION),
        ("source_rows", "PASS", len(balances) + len({_source_ref(d) for d in diagnostics}), "source rows retained"),
        ("posting_rows", "PASS", metrics["posting_row_count"], "two per actionable source row"),
        ("source_org_total", "PASS", metrics["source_org_total"], "Decimal total"),
        ("gk_total", "PASS", metrics["gk_total"], "Decimal total"),
        ("blocked_source_rows", "PASS", metrics["blocked_source_row_count"], "retained in diagnostics"),
        ("export_round_trip", "PASS", len(export_rows), "all produced workbooks reopened"),
    ]
    summary_rows.extend(
        (
            check["control"],
            check["status"],
            check["value"],
            "PASS",
        )
        for check in checks
    )
    _write_control_sheet(
        workbook,
        "Итоги",
        ("Контроль", "Статус", "Значение", "Ожидание"),
        summary_rows,
    )

    parameters = (
        ("run_id", run_id),
        ("contract_version", CONTRACT_VERSION),
        ("period_end", str(config.period_end) if config.period_end is not None else "discovered"),
        ("manager_organization", config.transfer_config.manager_organization),
        ("manager_financial_department", config.transfer_config.manager_financial_department),
        ("rules_version", config.transfer_config.rules_version),
        ("operation_type", adapter_config.operation_type),
        ("currency", adapter_config.currency),
        ("debit_activity", adapter_config.debit_activity),
        ("credit_activity", adapter_config.credit_activity),
        ("input_name", config.input_name),
    )
    _write_control_sheet(workbook, "Параметры_запуска", ("Параметр", "Значение"), parameters)

    _write_control_sheet(
        workbook,
        "Остатки_79",
        (
            "source_excel_row_ref",
            "period_end",
            "organization",
            "source_account",
            "department",
            "supplier_rvp",
            "ending_debit",
            "ending_credit",
            "status",
            "block_reason",
        ),
        (
            (
                record["source_excel_row_ref"],
                record["period_end"],
                record["organization"],
                record["source_account"],
                record["department"],
                record["supplier_rvp"],
                record["ending_debit"],
                record["ending_credit"],
                record["status"],
                record["block_reason"],
            )
            for record in sorted((_balance_record(balance) for balance in balances), key=lambda item: item["source_excel_row_ref"] or "")
        ),
    )

    posting_headers = (
        "period_end",
        "document_organization",
        "debit_account",
        "debit_department",
        "debit_supplier_rvp",
        "credit_account",
        "credit_department",
        "credit_supplier_rvp",
        "amount",
        "source_organization",
        "source_account",
        "source_department",
        "source_supplier_rvp",
        "source_excel_row_ref",
        "financial_record_id",
        "side",
        "rules_version",
    )
    _write_control_sheet(
        workbook,
        "Готовые_проводки",
        posting_headers,
        (
            tuple(
                "" if record[header] is None else record[header]
                for header in posting_headers
            )
            for record in (_posting_record(row) for row in posting_rows)
        ),
    )

    diagnostic_rows = [
        (
            _source_ref(diagnostic) or "",
            diagnostic.sheet_name or "",
            diagnostic.excel_row or "",
            "parser",
            diagnostic.code.value,
            diagnostic.reason.value,
            diagnostic.message,
            0,
            0,
            "BLOCKED",
        )
        for diagnostic in diagnostics
    ]
    _write_control_sheet(
        workbook,
        "Блокировки",
        (
            "source_excel_row_ref",
            "sheet_name",
            "excel_row",
            "stage",
            "code",
            "reason",
            "message",
            "financial_posting_count",
            "financial_export_count",
            "status",
        ),
        diagnostic_rows,
    )

    effect_headers = tuple(control_rows[0].keys()) if control_rows else (
        "source_excel_row_ref",
        "status",
    )
    _write_control_sheet(
        workbook,
        "Контроль_до_после",
        effect_headers,
        (tuple(row.get(header, "") for header in effect_headers) for row in control_rows),
    )

    diagnostic_by_ref: dict[str | None, list[ParserDiagnostic]] = defaultdict(list)
    for diagnostic in diagnostics:
        diagnostic_by_ref[_source_ref(diagnostic)].append(diagnostic)
    source_rows = []
    for balance in balances:
        source_rows.append(
            (
                balance.source_excel_row_ref or "",
                "normalized",
                balance.status.value,
                balance.organization,
                balance.source_account.value if balance.source_account else "",
                balance.department,
                balance.supplier_rvp,
                "",
                "",
            )
        )
    for source_ref, source_diagnostics in sorted(diagnostic_by_ref.items(), key=lambda item: item[0] or ""):
        for diagnostic in source_diagnostics:
            source_rows.append(
                (
                    source_ref or "",
                    "parser",
                    BalanceStatus.BLOCKED.value,
                    "",
                    "",
                    "",
                    "",
                    diagnostic.code.value,
                    diagnostic.message,
                )
            )
    _write_control_sheet(
        workbook,
        "Исходные_строки",
        (
            "source_excel_row_ref",
            "stage",
            "status",
            "organization",
            "source_account",
            "department",
            "supplier_rvp",
            "block_code",
            "message",
        ),
        source_rows,
    )

    _write_control_sheet(
        workbook,
        "Проверка_экспорта",
        (
            "path",
            "document_organization",
            "document_date",
            "sheet_name",
            "column_count",
            "financial_row_count",
            "round_trip",
            "status",
        ),
        (
            (
                record["path"],
                record["document_organization"],
                record["document_date"],
                record["sheet_name"],
                record["column_count"],
                record["financial_row_count"],
                record["round_trip"],
                record["status"],
            )
            for record in export_rows
        ),
    )

    workbook.save(path)
    workbook.close()


def _export_manifest(
    stage_root: Path,
    export_result: Any,
    rows: tuple[PostingRow, ...],
    adapter_config: OutputAdapterConfig,
) -> tuple[list[dict[str, Any]], int]:
    export_records: list[dict[str, Any]] = []
    for workbook in export_result.workbooks:
        group_rows = tuple(
            row
            for row in rows
            if row.period_end == workbook.document_date
            and row.document_organization == workbook.document_organization
        )
        validate_workbook_round_trip(workbook.path, group_rows, adapter_config)
        relative_path = workbook.path.relative_to(stage_root).as_posix()
        export_records.append(
            {
                "path": relative_path,
                "document_organization": workbook.document_organization,
                "document_date": workbook.document_date.isoformat(),
                "sheet_name": OUTPUT_SHEET_NAME,
                "column_count": len(OUTPUT_HEADERS),
                "financial_row_count": workbook.row_count,
                "round_trip": True,
                "status": "PASS",
            }
        )
    export_records.sort(key=lambda record: record["path"])
    export_row_count = sum(record["financial_row_count"] for record in export_records)
    if not export_records or export_row_count != len(rows):
        raise _control_failure("export manifest does not cover every financial PostingRow")
    return export_records, export_row_count


def _validate_artifacts(root: Path) -> None:
    for name in RUN_ARTIFACT_NAMES:
        path = root / name
        if not path.is_file():
            raise IntegrationRunError(f"mandatory artifact is missing: {name}")
    export_paths = sorted((root / "export").glob("*.xlsx"))
    if not export_paths:
        raise IntegrationRunError("mandatory artifact is missing: export/*.xlsx")

    workbook = load_workbook(root / "run_control.xlsx", read_only=True, data_only=False)
    try:
        missing = [name for name in CONTROL_SHEET_NAMES if name not in workbook.sheetnames]
        if missing:
            raise IntegrationRunError(
                "mandatory control sheets are missing: " + ", ".join(missing)
            )
    finally:
        workbook.close()

    for path in export_paths:
        workbook = load_workbook(path, read_only=True, data_only=False)
        try:
            if workbook.sheetnames != [OUTPUT_SHEET_NAME]:
                raise IntegrationRunError(f"export workbook has unexpected sheets: {path.name}")
            worksheet = workbook[OUTPUT_SHEET_NAME]
            if worksheet.max_column != len(OUTPUT_HEADERS):
                raise IntegrationRunError(f"export workbook has wrong column count: {path.name}")
        finally:
            workbook.close()


def run_integration(
    source: Any,
    output_dir: Path | str,
    *,
    config: IntegrationRunConfig | None = None,
    period_end: date | datetime | str | None = None,
    transfer_config: TransferConfig | None = None,
    adapter_config: OutputAdapterConfig | None = None,
    input_name: str | None = None,
) -> IntegrationRunResult:
    """Run parser -> transfer -> controls -> export as one local transaction.

    The destination must not already exist.  All files are built in a sibling
    staging directory and the complete directory is published only after the
    mandatory controls and workbook round trips pass.
    """

    if config is not None and any(
        value is not None for value in (period_end, transfer_config, adapter_config, input_name)
    ):
        raise ValueError("config cannot be combined with individual run options")
    run_config = config or IntegrationRunConfig(
        period_end=period_end,
        transfer_config=transfer_config or TransferConfig(),
        adapter_config=adapter_config or OutputAdapterConfig(),
        input_name=input_name or "osv.xlsx",
    )
    run_root = Path(output_dir)
    if run_root.exists():
        raise IntegrationRunError(f"run output already exists: {run_root}")
    if run_root.name in {"", ".", ".."}:
        raise ValueError("output_dir must name a dedicated run directory")
    parent = run_root.parent
    parent.mkdir(parents=True, exist_ok=True)

    workbook, close_workbook = _load_input_workbook(source)
    try:
        balances, diagnostics, sheet_records = _parse_sheets(workbook, run_config)
        run_id, fingerprint = _run_id(run_config, balances, diagnostics)
        if (
            run_config.transfer_config.rules_version != RULES_VERSION
            or run_config.adapter_config.rules_version != run_config.transfer_config.rules_version
            or run_config.adapter_config.contract_version != CONTRACT_VERSION
        ):
            raise _control_failure("transfer/output versions are not approved and identical")
        effective_adapter = replace(run_config.adapter_config, run_id=run_id)
        engine = TransferEngine(run_config.transfer_config)
        batch = engine.generate_batch(balances)
        control_rows, checks, metrics = _financial_controls(
            balances,
            batch,
            diagnostics,
            run_config.transfer_config,
        )
        if any(check["status"] != "PASS" for check in checks):
            raise _control_failure("one or more financial controls did not pass")

        with tempfile.TemporaryDirectory(prefix=".holding79-run-", dir=parent) as staging_name:
            stage_root = Path(staging_name)
            (stage_root / "export").mkdir()
            _write_json(
                stage_root / "input_manifest.json",
                {
                    "artifact_version": "H79_TRANSFER_RUN_V1",
                    "run_id": run_id,
                    "contract_version": CONTRACT_VERSION,
                    "input_name": run_config.input_name,
                    "normalized_input_sha256": fingerprint,
                    "period_end": (
                        balances[0].period_end.isoformat() if balances and balances[0].period_end else None
                    ),
                    "source_sheets": sheet_records,
                    "parser_diagnostics": [_diagnostic_record(diagnostic) for diagnostic in diagnostics],
                    "config": _config_identity(run_config),
                },
            )
            _write_jsonl(
                stage_root / "normalized_balances.jsonl",
                (_balance_record(balance) for balance in balances),
            )
            ordered_rows = tuple(sorted(batch.rows, key=_posting_sort_key))
            _write_jsonl(
                stage_root / "posting_rows.jsonl",
                (_posting_record(row) for row in ordered_rows),
            )

            export_result = export_posting_rows(
                ordered_rows,
                stage_root / "export",
                effective_adapter,
            )
            export_records, export_row_count = _export_manifest(
                stage_root,
                export_result,
                ordered_rows,
                effective_adapter,
            )
            if export_row_count != metrics["posting_row_count"]:
                raise _control_failure("export row count does not match financial PostingRows")
            checks.append({"control": "export_round_trip", "status": "PASS", "value": export_row_count})

            _write_run_control(
                stage_root / "run_control.xlsx",
                run_id=run_id,
                config=run_config,
                adapter_config=effective_adapter,
                balances=balances,
                posting_rows=ordered_rows,
                diagnostics=diagnostics,
                control_rows=control_rows,
                checks=checks,
                export_rows=export_records,
                metrics=metrics,
            )

            artifact_paths = list(RUN_ARTIFACT_NAMES)
            artifact_paths.extend(record["path"] for record in export_records)
            summary = {
                "status": "SUCCESS",
                "run_id": run_id,
                "contract_version": CONTRACT_VERSION,
                "rules_version": run_config.transfer_config.rules_version,
                "counts": {
                    "source_rows": len(balances) + len({_source_ref(d) for d in diagnostics}),
                    "normalized_balances": len(balances),
                    "actionable_source_rows": metrics["actionable_source_row_count"],
                    "no_action_source_rows": metrics["no_action_source_row_count"],
                    "blocked_source_rows": metrics["blocked_source_row_count"],
                    "posting_rows": metrics["posting_row_count"],
                    "export_rows": export_row_count,
                    "export_workbooks": len(export_records),
                    "parser_diagnostics": len(diagnostics),
                },
                "totals": {
                    "source_org": metrics["source_org_total"],
                    "gk": metrics["gk_total"],
                    "difference": _decimal_text(
                        Decimal(metrics["source_org_total"]) - Decimal(metrics["gk_total"])
                    ),
                },
                "controls": {
                    "all_passed": True,
                    "records": checks,
                },
                "artifacts": artifact_paths,
            }
            _write_json(stage_root / "export_manifest.json", {
                "run_id": run_id,
                "contract_version": CONTRACT_VERSION,
                "sheet_name": OUTPUT_SHEET_NAME,
                "headers": list(OUTPUT_HEADERS),
                "financial_row_count": export_row_count,
                "round_trip_validated": True,
                "workbooks": export_records,
            })
            _write_json(stage_root / "summary.json", summary)
            _validate_artifacts(stage_root)
            stage_root.replace(run_root)

        published_workbooks = tuple(
            ExportedWorkbook(
                path=run_root / record["path"],
                document_organization=record["document_organization"],
                document_date=date.fromisoformat(record["document_date"]),
                row_count=record["financial_row_count"],
            )
            for record in export_records
        )
        return IntegrationRunResult(
            output_dir=run_root,
            run_id=run_id,
            status="SUCCESS",
            normalized_balances=balances,
            posting_rows=ordered_rows,
            diagnostics=diagnostics,
            exported_workbooks=published_workbooks,
        )
    finally:
        if close_workbook:
            workbook.close()


def build_synthetic_osv_workbook() -> Workbook:
    """Create the in-memory OSV covering the four approved goldens plus one block."""

    cases = (
        ("DEBIT_79_2_AT", "79.2", "84272.40", "0"),
        ("CREDIT_79_2_AT", "79.2", "0", "84272.40"),
        ("DEBIT_79_3_AT", "79.3", "84272.40", "0"),
        ("CREDIT_79_3_AT", "79.3", "0", "84272.40"),
    )
    workbook = Workbook()
    for index, (name, account, debit, credit) in enumerate(cases):
        worksheet = workbook.active if index == 0 else workbook.create_sheet()
        worksheet.title = name
        _populate_synthetic_sheet(
            worksheet,
            account=account,
            organization="АТ",
            department="Б_АТ Коммерческий отдел",
            supplier="Производитель",
            debit=debit,
            credit=credit,
        )
    blocked = workbook.create_sheet("BLOCKED_SOURCE_ROW")
    _populate_synthetic_sheet(
        blocked,
        account="79.2",
        organization="АТ",
        department="Б_АТ Коммерческий отдел",
        supplier=None,
        debit="100.00",
        credit="0",
    )
    return workbook


def _populate_synthetic_sheet(
    worksheet: Worksheet,
    *,
    account: str,
    organization: str,
    department: str,
    supplier: str | None,
    debit: str,
    credit: str,
) -> None:
    worksheet.cell(1, 1).value = "Синтетическая ОСВ"
    worksheet.cell(2, 1).value = "Период: 01.12.2024 - 31.12.2024"
    worksheet.merge_cells("H4:I4")
    worksheet["H4"] = "Сальдо на конец периода"
    worksheet["H5"] = "Дебет"
    worksheet["I5"] = "Кредит"
    rows = (
        (6, account, 0),
        (7, f"Организация: {organization}", 1),
        (8, f"ЦФО: {department}", 2),
        (9, "Поставщик РВП" if supplier is None else f"Поставщик РВП: {supplier}", 3),
    )
    for row, value, indent in rows:
        worksheet.cell(row, 1).value = value
        worksheet.cell(row, 1).alignment = Alignment(indent=indent)
    worksheet.cell(9, 8).value = debit
    worksheet.cell(9, 9).value = credit


def run_synthetic_integration(output_dir: Path | str) -> IntegrationRunResult:
    """Execute the permanent local synthetic Issue #5 scenario."""

    workbook = build_synthetic_osv_workbook()
    try:
        return run_integration(
            workbook,
            output_dir,
            config=IntegrationRunConfig(
                period_end=date(2024, 12, 31),
                input_name="synthetic-integration.xlsx",
            ),
        )
    finally:
        workbook.close()


# Descriptive aliases for callers that use the run terminology from the issue.
run_end_to_end = run_integration
run_local_synthetic = run_synthetic_integration


__all__ = [
    "CONTROL_SHEET_NAMES",
    "RUN_ARTIFACT_NAMES",
    "IntegrationRunConfig",
    "IntegrationRunError",
    "IntegrationRunResult",
    "build_synthetic_osv_workbook",
    "run_end_to_end",
    "run_integration",
    "run_local_synthetic",
    "run_synthetic_integration",
]
