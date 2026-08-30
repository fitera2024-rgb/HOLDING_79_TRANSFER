"""Canonical domain models for the approved 79.x transfer contract.

This module deliberately contains no file, Excel, or 1C integration.  It only
defines the values that later parser, transfer-engine, and exporter layers may
exchange.
"""

from __future__ import annotations

import hashlib
import json
from dataclasses import dataclass
from datetime import date
from decimal import Decimal, InvalidOperation
from enum import Enum
from typing import Any

from pydantic import BaseModel, ConfigDict, Field, field_validator

CONTRACT_VERSION = "0.3-approved"
RULES_VERSION = "H79_TRANSFER_V1"
DEBIT_RULE_ID = "H79_DEBIT_TRANSFER_V1"
CREDIT_RULE_ID = "H79_CREDIT_TRANSFER_V1"


class SourceAccount(str, Enum):
    """The only source accounts accepted by the month-end transfer."""

    ACCOUNT_79_2 = "79.2"
    ACCOUNT_79_3 = "79.3"


class AccountCode(str, Enum):
    """Account codes used by the approved transfer postings."""

    ACCOUNT_79_1 = "79.1"
    ACCOUNT_79_2 = "79.2"
    ACCOUNT_79_3 = "79.3"


class BalanceSide(str, Enum):
    DEBIT = "DEBIT"
    CREDIT = "CREDIT"


class BalanceStatus(str, Enum):
    ACTIONABLE = "ACTIONABLE"
    NO_ACTION = "NO_ACTION"
    BLOCKED = "BLOCKED"


class BlockReason(str, Enum):
    BLOCKED_INVALID_ENDING_BALANCE = "BLOCKED_INVALID_ENDING_BALANCE"
    BLOCKED_MISSING_SOURCE_IDENTITY = "BLOCKED_MISSING_SOURCE_IDENTITY"
    BLOCKED_INVALID_POSTING = "BLOCKED_INVALID_POSTING"

    # Short names are aliases for callers that use the domain concept rather
    # than the serialized BLOCKED_* reason code.
    INVALID_ENDING_BALANCE = BLOCKED_INVALID_ENDING_BALANCE
    MISSING_SOURCE_IDENTITY = BLOCKED_MISSING_SOURCE_IDENTITY
    INVALID_POSTING = BLOCKED_INVALID_POSTING


# Short aliases make the public vocabulary convenient without introducing a
# second set of enum values.
Status = BalanceStatus
NormalizationStatus = BalanceStatus
RowStatus = BalanceStatus
BlockedReason = BlockReason


def normalize_decimal(value: Any) -> Decimal:
    """Convert an explicit monetary representation without using binary float.

    Blank and em-dash service cells are explicit zero representations.  A
    float is rejected rather than converted, because doing so would preserve
    a binary approximation in an accounting value.
    """

    if isinstance(value, (bool, float)):
        raise TypeError("monetary values must not be bool or float")
    if value is None:
        return Decimal(0)
    if isinstance(value, Decimal):
        result = value
    elif isinstance(value, int):
        result = Decimal(value)
    elif isinstance(value, str):
        text = value.replace("\u00a0", " ").strip()
        if text in {"", "-", "—", "–"}:
            return Decimal(0)
        if text.startswith("(") and text.endswith(")"):
            text = f"-{text[1:-1].strip()}"
        if "," in text and "." in text:
            raise ValueError(f"ambiguous monetary value: {value!r}")
        if "," in text:
            text = text.replace(",", ".")
        try:
            result = Decimal(text)
        except InvalidOperation as exc:
            raise ValueError(f"invalid monetary value: {value!r}") from exc
    else:
        raise TypeError(f"unsupported monetary value type: {type(value).__name__}")

    if not result.is_finite():
        raise ValueError("monetary values must be finite")
    return result


def validate_decimal(value: Any) -> Decimal:
    try:
        return normalize_decimal(value)
    except TypeError as exc:
        raise ValueError(str(exc)) from exc


def _clean_text(value: Any) -> str:
    if value is None:
        return ""
    if not isinstance(value, str):
        raise TypeError("identity fields must be strings")
    return value.strip()


def _clean_optional_text(value: Any) -> str | None:
    if value is None:
        return None
    return _clean_text(value)


class NormalizedBalance(BaseModel):
    """One normalized source analytical row from a 79.2/79.3 balance."""

    model_config = ConfigDict(extra="forbid", frozen=True)

    period_end: date | None = None
    organization: str = ""
    source_account: SourceAccount | None = None
    department: str = ""
    supplier_rvp: str = ""
    ending_debit: Decimal = Field(default=Decimal(0))
    ending_credit: Decimal = Field(default=Decimal(0))
    source_excel_row_ref: str | None = None

    _clean_identity = field_validator("organization", "department", "supplier_rvp", mode="before")(
        _clean_text
    )
    _clean_source_ref = field_validator("source_excel_row_ref", mode="before")(
        _clean_optional_text
    )
    _clean_money = field_validator("ending_debit", "ending_credit", mode="before")(
        validate_decimal
    )

    @field_validator("ending_debit", "ending_credit")
    @classmethod
    def _ending_balance_is_non_negative(cls, value: Decimal) -> Decimal:
        if value < 0:
            raise ValueError("ending balances must be non-negative after normalization")
        return value

    @property
    def status(self) -> BalanceStatus:
        return validate_normalized_balance(self).status

    @property
    def block_reason(self) -> BlockReason | None:
        return validate_normalized_balance(self).reason

    @property
    def ending_side(self) -> BalanceSide | None:
        result = validate_normalized_balance(self)
        if result.status is not BalanceStatus.ACTIONABLE:
            return None
        return BalanceSide.DEBIT if self.ending_debit else BalanceSide.CREDIT

    @property
    def amount(self) -> Decimal:
        """Return the actionable side amount; invalid/zero rows return zero."""

        if self.ending_side is BalanceSide.DEBIT:
            return self.ending_debit
        if self.ending_side is BalanceSide.CREDIT:
            return self.ending_credit
        return Decimal(0)

    @property
    def validation(self) -> ValidationResult:
        return validate_normalized_balance(self)


@dataclass(frozen=True)
class ValidationResult:
    status: BalanceStatus
    reason: BlockReason | None = None
    message: str | None = None

    @property
    def is_actionable(self) -> bool:
        return self.status is BalanceStatus.ACTIONABLE


def validate_normalized_balance(balance: NormalizedBalance) -> ValidationResult:
    """Classify a normalized row without changing it or inventing values."""

    if balance.ending_debit and balance.ending_credit:
        return ValidationResult(
            BalanceStatus.BLOCKED,
            BlockReason.BLOCKED_INVALID_ENDING_BALANCE,
            "both ending debit and ending credit are non-zero",
        )

    missing = {
        "period_end": balance.period_end,
        "organization": balance.organization,
        "source_account": balance.source_account,
        "department": balance.department,
        "supplier_rvp": balance.supplier_rvp,
        "source_excel_row_ref": balance.source_excel_row_ref,
    }
    if any(not value for value in missing.values()):
        fields = ", ".join(name for name, value in missing.items() if not value)
        return ValidationResult(
            BalanceStatus.BLOCKED,
            BlockReason.BLOCKED_MISSING_SOURCE_IDENTITY,
            f"missing source identity: {fields}",
        )

    if not balance.ending_debit and not balance.ending_credit:
        return ValidationResult(BalanceStatus.NO_ACTION)
    return ValidationResult(BalanceStatus.ACTIONABLE)


class PostingRow(BaseModel):
    """A single balanced posting row, independent of any file format."""

    model_config = ConfigDict(extra="forbid", frozen=True)

    document_organization: str
    debit_account: AccountCode
    debit_department: str
    debit_supplier_rvp: str
    credit_account: AccountCode
    credit_department: str
    credit_supplier_rvp: str
    amount: Decimal

    # Traceability and deterministic identity are part of the canonical row,
    # but defaults keep the model able to represent the compact owner goldens.
    period_end: date | None = None
    source_organization: str = ""
    source_account: SourceAccount | None = None
    source_department: str = ""
    source_supplier_rvp: str = ""
    source_excel_row_ref: str | None = None
    financial_record_id: str = ""
    side: BalanceSide | None = None
    rules_version: str = RULES_VERSION

    _clean_posting_text = field_validator(
        "document_organization",
        "debit_department",
        "debit_supplier_rvp",
        "credit_department",
        "credit_supplier_rvp",
        "source_organization",
        "source_department",
        "source_supplier_rvp",
        "financial_record_id",
        "rules_version",
        mode="before",
    )(_clean_text)
    _clean_source_ref = field_validator("source_excel_row_ref", mode="before")(
        _clean_optional_text
    )
    _clean_posting_amount = field_validator("amount", mode="before")(validate_decimal)

    @field_validator("amount")
    @classmethod
    def _amount_is_positive(cls, value: Decimal) -> Decimal:
        if value <= 0:
            raise ValueError("a PostingRow amount must be positive")
        return value

    @property
    def rule_id(self) -> str:
        return rule_id_for_side(self.side) if self.side else ""

    @property
    def status(self) -> BalanceStatus:
        return validate_posting_row(self).status

    @property
    def validation(self) -> ValidationResult:
        return validate_posting_row(self)

    @property
    def direction(self) -> BalanceSide | None:
        return self.side


def validate_posting_row(row: PostingRow) -> ValidationResult:
    """Validate the identity needed to safely map a PostingRow to output."""

    required = {
        "period_end": row.period_end,
        "document_organization": row.document_organization,
        "source_organization": row.source_organization,
        "source_account": row.source_account,
        "source_department": row.source_department,
        "source_supplier_rvp": row.source_supplier_rvp,
        "source_excel_row_ref": row.source_excel_row_ref,
        "debit_department": row.debit_department,
        "debit_supplier_rvp": row.debit_supplier_rvp,
        "credit_department": row.credit_department,
        "credit_supplier_rvp": row.credit_supplier_rvp,
        "side": row.side,
        "financial_record_id": row.financial_record_id,
    }
    missing = [name for name, value in required.items() if not value]
    if missing:
        return ValidationResult(
            BalanceStatus.BLOCKED,
            BlockReason.BLOCKED_INVALID_POSTING,
            f"missing PostingRow identity: {', '.join(missing)}",
        )
    return ValidationResult(BalanceStatus.ACTIONABLE)


def rule_id_for_side(side: BalanceSide) -> str:
    """Return the approved deterministic rule id for an ending-balance side."""

    if side is BalanceSide.DEBIT:
        return DEBIT_RULE_ID
    if side is BalanceSide.CREDIT:
        return CREDIT_RULE_ID
    raise ValueError(f"unsupported balance side: {side!r}")


def _decimal_identity(value: Decimal) -> str:
    """Stable, exponent-free representation for hashing a monetary identity."""

    return format(value.normalize(), "f")


def normalized_business_identity(balance: NormalizedBalance) -> dict[str, str]:
    """Return the normalized business fields used by the financial-record id."""

    return {
        "period_end": balance.period_end.isoformat() if balance.period_end else "",
        "organization": balance.organization,
        "source_account": balance.source_account.value if balance.source_account else "",
        "department": balance.department,
        "supplier_rvp": balance.supplier_rvp,
        "ending_debit": _decimal_identity(balance.ending_debit),
        "ending_credit": _decimal_identity(balance.ending_credit),
    }


def financial_record_id(
    balance: NormalizedBalance, rules_version: str = RULES_VERSION
) -> str:
    """Create a stable id from normalized business identity and rules version."""

    if not isinstance(rules_version, str) or not rules_version.strip():
        raise ValueError("rules_version must be a non-empty string")
    payload = {
        "rules_version": rules_version.strip(),
        "business_identity": normalized_business_identity(balance),
    }
    encoded = json.dumps(
        payload, ensure_ascii=False, sort_keys=True, separators=(",", ":")
    ).encode("utf-8")
    return f"FR-{hashlib.sha256(encoded).hexdigest()}"


deterministic_financial_record_id = financial_record_id
