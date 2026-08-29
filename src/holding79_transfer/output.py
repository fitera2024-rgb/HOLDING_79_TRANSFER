"""Pure adapter from canonical PostingRow values to the 27-column loader row."""

from __future__ import annotations

from dataclasses import dataclass
from decimal import Decimal
from typing import Any

from .models import (
    CONTRACT_VERSION,
    RULES_VERSION,
    BalanceStatus,
    BlockReason,
    PostingRow,
    validate_posting_row,
)

OUTPUT_SHEET_NAME = "Загрузка_A_AA"
OUTPUT_HEADERS: tuple[str, ...] = (
    "СчетДт",
    "СчетКт",
    "ВалютаДт",
    "ВалютаКт",
    "ВидОперации",
    "ПодразделениеДт",
    "ПодразделениеКт",
    "НаправлениеДеятельностиДт",
    "НаправлениеДеятельностиКт",
    "СуммаВВалютеУчета",
    "СуммаВВалютеОтчетности",
    "СуммаВВалютеДт",
    "СуммаВВалютеКт",
    "КоличествоДт",
    "КоличествоКт",
    "Содержание",
    "СчетДтИсточник",
    "СчетКтИсточник",
    "ИдентификаторФинЗаписи",
    "ПравилоДт",
    "ПравилоКт",
    "СубконтоДт1",
    "СубконтоДт2",
    "СубконтоДт3",
    "СубконтоКт1",
    "СубконтоКт2",
    "СубконтоКт3",
)


@dataclass(frozen=True)
class OutputAdapterConfig:
    """Explicit values for adapter fields not carried by PostingRow.

    Currency and activity fields intentionally default to empty values: the
    approved contract does not provide a currency/activity reference value,
    and the adapter must not guess one.  A caller with reference data may set
    ``currency`` and the activity fields explicitly.
    """

    operation_type: str = "REPOST"
    currency: str = ""
    debit_activity: str = ""
    credit_activity: str = ""
    run_id: str = ""
    contract_version: str = CONTRACT_VERSION
    rules_version: str = RULES_VERSION

    def __post_init__(self) -> None:
        for name in (
            "operation_type",
            "currency",
            "debit_activity",
            "credit_activity",
            "run_id",
            "contract_version",
            "rules_version",
        ):
            value = getattr(self, name)
            if not isinstance(value, str):
                raise TypeError(f"{name} must be a string")
            object.__setattr__(self, name, value.strip())
        if not self.operation_type:
            raise ValueError("operation_type must not be empty")
        if not self.contract_version:
            raise ValueError("contract_version must not be empty")
        if not self.rules_version:
            raise ValueError("rules_version must not be empty")


@dataclass(frozen=True)
class OutputMappingResult:
    status: BalanceStatus
    row: dict[str, Any] | None = None
    reason: BlockReason | None = None
    message: str | None = None


class BlockedOutputMapping(ValueError):
    """Raised by the strict mapper when a PostingRow cannot be exported safely."""

    def __init__(self, reason: BlockReason, message: str) -> None:
        super().__init__(message)
        self.reason = reason
        self.message = message


def _account_value(account: Any) -> str:
    return account.value if hasattr(account, "value") else str(account)


def _content(row: PostingRow, config: OutputAdapterConfig) -> str:
    period = row.period_end.isoformat() if row.period_end else ""
    source_account = _account_value(row.source_account)
    direction = row.side.value if row.side else ""
    parts = (
        ("run_id", config.run_id),
        ("source_row_ref", row.source_excel_row_ref or ""),
        ("period_end", period),
        ("source_organization", row.source_organization),
        ("source_account", source_account),
        ("source_department", row.source_department),
        ("source_supplier_rvp", row.source_supplier_rvp),
        ("document_organization", row.document_organization),
        ("direction", direction),
        ("contract_version", config.contract_version),
        ("rules_version", config.rules_version),
    )
    return "; ".join(f"{key}={value}" for key, value in parts)


def _output_values(row: PostingRow, config: OutputAdapterConfig) -> tuple[Any, ...]:
    source_account = _account_value(row.source_account)
    rule_id = row.rule_id
    amount: Decimal = row.amount
    return (
        _account_value(row.debit_account),
        _account_value(row.credit_account),
        config.currency,
        config.currency,
        config.operation_type,
        row.debit_department,
        row.credit_department,
        config.debit_activity,
        config.credit_activity,
        amount,
        None,
        None,
        None,
        None,
        None,
        _content(row, config),
        source_account,
        source_account,
        row.financial_record_id,
        rule_id,
        rule_id,
        row.debit_supplier_rvp,
        None,
        None,
        row.credit_supplier_rvp,
        None,
        None,
    )


def output_row_values(
    row: PostingRow, config: OutputAdapterConfig | None = None
) -> tuple[Any, ...]:
    """Return one output row in exactly ``OUTPUT_HEADERS`` order."""

    result = validate_posting_row(row)
    if result.status is BalanceStatus.BLOCKED:
        reason = result.reason or BlockReason.INVALID_POSTING
        raise BlockedOutputMapping(reason, result.message or "blocked")
    adapter_config = config or OutputAdapterConfig()
    values = _output_values(row, adapter_config)
    if len(values) != len(OUTPUT_HEADERS):  # defensive contract guard
        raise AssertionError("output row does not match the 27-column contract")
    return values


def map_posting_row(
    row: PostingRow, config: OutputAdapterConfig | None = None
) -> dict[str, Any]:
    """Map a PostingRow to a header-keyed 27-column output record."""

    return dict(zip(OUTPUT_HEADERS, output_row_values(row, config), strict=True))


def try_map_posting_row(
    row: PostingRow, config: OutputAdapterConfig | None = None
) -> OutputMappingResult:
    """Return a blocked result instead of raising for pipeline diagnostics."""

    try:
        mapped = map_posting_row(row, config)
    except BlockedOutputMapping as exc:
        return OutputMappingResult(BalanceStatus.BLOCKED, reason=exc.reason, message=exc.message)
    return OutputMappingResult(BalanceStatus.ACTIONABLE, row=mapped)


HEADER = OUTPUT_HEADERS
OUTPUT_COLUMNS = OUTPUT_HEADERS
