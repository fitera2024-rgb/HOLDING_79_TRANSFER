"""Pure symmetric transfer engine for exact source accounts 79.2 and 79.3."""

from __future__ import annotations

from collections.abc import Iterable
from dataclasses import dataclass
from decimal import Decimal

from .config import TransferConfig
from .models import (
    RULES_VERSION,
    AccountCode,
    BalanceSide,
    BalanceStatus,
    BlockReason,
    NormalizedBalance,
    PostingRow,
    SourceAccount,
    financial_record_id,
    validate_normalized_balance,
)


@dataclass(frozen=True)
class SourceEffectControl:
    """Evidence that one generated source row consumes the source balance."""

    source_row_ref: str | None
    source_account: SourceAccount | None
    side: BalanceSide | None
    ending_debit_before: Decimal
    ending_credit_before: Decimal
    ending_debit_after: Decimal
    ending_credit_after: Decimal
    source_posting_count: int
    source_amount: Decimal | None
    gk_amount: Decimal | None
    source_effect_zero: bool
    amounts_match: bool
    direction_correct: bool
    source_consumed_once: bool

    @property
    def passed(self) -> bool:
        """Whether every source-effect acceptance check passed."""

        return all(
            (
                self.source_effect_zero,
                self.amounts_match,
                self.direction_correct,
                self.source_consumed_once,
            )
        )

    @property
    def is_valid(self) -> bool:
        return self.passed

    @property
    def source_balance_after(self) -> Decimal:
        """Return the remaining amount on the original ending-balance side."""

        if self.side is BalanceSide.DEBIT:
            return self.ending_debit_after
        if self.side is BalanceSide.CREDIT:
            return self.ending_credit_after
        return Decimal(0)


@dataclass(frozen=True)
class TransferResult:
    """Result of classifying and, when safe, transforming one balance."""

    status: BalanceStatus
    rows: tuple[PostingRow, ...] = ()
    source_effect: SourceEffectControl | None = None
    reason: BlockReason | None = None
    message: str | None = None

    @property
    def is_actionable(self) -> bool:
        return self.status is BalanceStatus.ACTIONABLE


@dataclass(frozen=True)
class TransferBatchResult:
    """Fail-closed result for transforming a collection of source balances."""

    status: BalanceStatus
    rows: tuple[PostingRow, ...] = ()
    transfers: tuple[TransferResult, ...] = ()
    reason: BlockReason | None = None
    message: str | None = None

    @property
    def is_actionable(self) -> bool:
        return self.status is BalanceStatus.ACTIONABLE


class TransferEngine:
    """Generate the two approved postings for one normalized balance."""

    def __init__(self, config: TransferConfig | None = None) -> None:
        self.config = config or TransferConfig()

    def generate(self, balance: NormalizedBalance) -> TransferResult:
        validation = validate_normalized_balance(balance)
        if validation.status is BalanceStatus.BLOCKED:
            return TransferResult(
                status=validation.status,
                reason=validation.reason,
                message=validation.message,
            )
        if validation.status is BalanceStatus.NO_ACTION:
            return TransferResult(status=BalanceStatus.NO_ACTION)
        if self.config.rules_version != RULES_VERSION:
            return TransferResult(
                status=BalanceStatus.BLOCKED,
                reason=BlockReason.BLOCKED_INVALID_POSTING,
                message=(
                    "unsupported rules_version for transfer engine: "
                    f"{self.config.rules_version}"
                ),
            )

        rows = _build_rows(balance, self.config)
        control = validate_source_effect(balance, rows, self.config)
        if not control.passed:
            return TransferResult(
                status=BalanceStatus.BLOCKED,
                source_effect=control,
                reason=BlockReason.BLOCKED_INVALID_POSTING,
                message=_source_effect_failure_message(control),
            )
        return TransferResult(
            status=BalanceStatus.ACTIONABLE,
            rows=rows,
            source_effect=control,
        )

    def build_rows(self, balance: NormalizedBalance) -> tuple[PostingRow, ...]:
        """Return rows for an actionable balance, otherwise return no rows."""

        return self.generate(balance).rows

    def generate_batch(self, balances: Iterable[NormalizedBalance]) -> TransferBatchResult:
        """Transform balances and reject duplicate source-row consumption."""

        results: list[TransferResult] = []
        all_rows: list[PostingRow] = []
        seen_source_refs: set[str] = set()
        blocked_message: str | None = None

        for balance in balances:
            result = self.generate(balance)
            if result.is_actionable:
                source_ref = balance.source_excel_row_ref
                if source_ref in seen_source_refs:
                    result = TransferResult(
                        status=BalanceStatus.BLOCKED,
                        reason=BlockReason.BLOCKED_INVALID_POSTING,
                        message=f"source row consumed more than once: {source_ref}",
                    )
                else:
                    seen_source_refs.add(source_ref or "")
                    all_rows.extend(result.rows)
            if result.status is BalanceStatus.BLOCKED and blocked_message is None:
                blocked_message = result.message
            results.append(result)

        if any(result.status is BalanceStatus.BLOCKED for result in results):
            return TransferBatchResult(
                status=BalanceStatus.BLOCKED,
                transfers=tuple(results),
                reason=BlockReason.BLOCKED_INVALID_POSTING,
                message=blocked_message or "one or more transfers are blocked",
            )
        if any(result.status is BalanceStatus.ACTIONABLE for result in results):
            return TransferBatchResult(
                status=BalanceStatus.ACTIONABLE,
                rows=tuple(all_rows),
                transfers=tuple(results),
            )
        return TransferBatchResult(
            status=BalanceStatus.NO_ACTION,
            transfers=tuple(results),
        )


def _build_rows(balance: NormalizedBalance, config: TransferConfig) -> tuple[PostingRow, ...]:
    side = balance.ending_side
    source_account = _source_account_code(balance.source_account)
    manager_account = AccountCode.ACCOUNT_79_1
    record_id = financial_record_id(balance, config.rules_version)
    common = {
        "period_end": balance.period_end,
        "source_organization": balance.organization,
        "source_account": balance.source_account,
        "source_department": balance.department,
        "source_supplier_rvp": balance.supplier_rvp,
        "source_excel_row_ref": balance.source_excel_row_ref,
        "financial_record_id": record_id,
        "side": side,
        "rules_version": config.rules_version,
    }

    if side is BalanceSide.DEBIT:
        source_row = PostingRow(
            document_organization=balance.organization,
            debit_account=manager_account,
            debit_department=balance.department,
            debit_supplier_rvp=config.manager_organization,
            credit_account=source_account,
            credit_department=balance.department,
            credit_supplier_rvp=balance.supplier_rvp,
            amount=balance.amount,
            **common,
        )
        gk_row = PostingRow(
            document_organization=config.manager_organization,
            debit_account=manager_account,
            debit_department=config.manager_financial_department,
            debit_supplier_rvp=balance.organization,
            credit_account=manager_account,
            credit_department=config.manager_financial_department,
            credit_supplier_rvp=balance.supplier_rvp,
            amount=balance.amount,
            **common,
        )
    elif side is BalanceSide.CREDIT:
        source_row = PostingRow(
            document_organization=balance.organization,
            debit_account=source_account,
            debit_department=balance.department,
            debit_supplier_rvp=balance.supplier_rvp,
            credit_account=manager_account,
            credit_department=balance.department,
            credit_supplier_rvp=config.manager_organization,
            amount=balance.amount,
            **common,
        )
        gk_row = PostingRow(
            document_organization=config.manager_organization,
            debit_account=manager_account,
            debit_department=config.manager_financial_department,
            debit_supplier_rvp=balance.supplier_rvp,
            credit_account=manager_account,
            credit_department=config.manager_financial_department,
            credit_supplier_rvp=balance.organization,
            amount=balance.amount,
            **common,
        )
    else:  # pragma: no cover - guarded by validate_normalized_balance
        raise ValueError("an actionable balance must have a side")
    return source_row, gk_row


def _source_account_code(source_account: SourceAccount | None) -> AccountCode:
    if source_account is None:  # pragma: no cover - guarded by validation
        raise ValueError("an actionable balance must have a source account")
    return AccountCode(source_account.value)


def validate_source_effect(
    balance: NormalizedBalance,
    rows: Iterable[PostingRow],
    config: TransferConfig | None = None,
) -> SourceEffectControl:
    """Check that generated rows consume one source balance exactly once."""

    engine_config = config or TransferConfig()
    rows_tuple = tuple(rows)
    side = balance.ending_side
    source_rows = [row for row in rows_tuple if _is_source_role(balance, row)]
    gk_rows = [
        row
        for row in rows_tuple
        if _is_gk_role(balance, row, engine_config)
    ]
    source_row = source_rows[0] if len(source_rows) == 1 else None
    gk_row = gk_rows[0] if len(gk_rows) == 1 else None
    source_amount = source_row.amount if source_row else None
    gk_amount = gk_row.amount if gk_row else None

    remaining = balance.amount if source_amount is None else balance.amount - source_amount
    if side is BalanceSide.DEBIT:
        ending_debit_after = remaining
        ending_credit_after = balance.ending_credit
    elif side is BalanceSide.CREDIT:
        ending_debit_after = balance.ending_debit
        ending_credit_after = remaining
    else:
        ending_debit_after = balance.ending_debit
        ending_credit_after = balance.ending_credit

    source_consumed_once = (
        len(rows_tuple) == 2
        and len(source_rows) == 1
        and len(gk_rows) == 1
        and source_row is not None
        and source_row.source_excel_row_ref == balance.source_excel_row_ref
    )
    amounts_match = (
        source_amount is not None
        and gk_amount is not None
        and source_amount == gk_amount == balance.amount
    )
    direction_correct = _direction_is_correct(
        balance, source_row, gk_row, engine_config
    )
    source_effect_zero = (
        source_row is not None
        and direction_correct
        and ending_debit_after == Decimal(0)
        and ending_credit_after == Decimal(0)
    )

    return SourceEffectControl(
        source_row_ref=balance.source_excel_row_ref,
        source_account=balance.source_account,
        side=side,
        ending_debit_before=balance.ending_debit,
        ending_credit_before=balance.ending_credit,
        ending_debit_after=ending_debit_after,
        ending_credit_after=ending_credit_after,
        source_posting_count=len(source_rows),
        source_amount=source_amount,
        gk_amount=gk_amount,
        source_effect_zero=source_effect_zero,
        amounts_match=amounts_match,
        direction_correct=direction_correct,
        source_consumed_once=source_consumed_once,
    )


def _direction_is_correct(
    balance: NormalizedBalance,
    source_row: PostingRow | None,
    gk_row: PostingRow | None,
    config: TransferConfig,
) -> bool:
    if (
        source_row is None
        or gk_row is None
        or balance.ending_side is None
        or config.rules_version != RULES_VERSION
    ):
        return False
    source_account = _source_account_code(balance.source_account)
    manager_account = AccountCode.ACCOUNT_79_1
    source_trace_matches = all(
        (
            source_row.period_end == balance.period_end,
            source_row.source_organization == balance.organization,
            source_row.source_account == balance.source_account,
            source_row.source_department == balance.department,
            source_row.source_supplier_rvp == balance.supplier_rvp,
            source_row.source_excel_row_ref == balance.source_excel_row_ref,
            source_row.financial_record_id
            == financial_record_id(balance, config.rules_version),
            source_row.side is balance.ending_side,
            gk_row.period_end == balance.period_end,
            gk_row.source_organization == balance.organization,
            gk_row.source_account == balance.source_account,
            gk_row.source_department == balance.department,
            gk_row.source_supplier_rvp == balance.supplier_rvp,
            gk_row.source_excel_row_ref == balance.source_excel_row_ref,
            gk_row.financial_record_id
            == financial_record_id(balance, config.rules_version),
            gk_row.side is balance.ending_side,
        )
    )
    if balance.ending_side is BalanceSide.DEBIT:
        source_direction = (
            source_row.document_organization == balance.organization
            and source_row.debit_account is manager_account
            and source_row.debit_department == balance.department
            and source_row.debit_supplier_rvp == config.manager_organization
            and source_row.credit_account is source_account
            and source_row.credit_department == balance.department
            and source_row.credit_supplier_rvp == balance.supplier_rvp
        )
        gk_direction = (
            gk_row.document_organization == config.manager_organization
            and gk_row.debit_account is manager_account
            and gk_row.debit_department == config.manager_financial_department
            and gk_row.debit_supplier_rvp == balance.organization
            and gk_row.credit_account is manager_account
            and gk_row.credit_department == config.manager_financial_department
            and gk_row.credit_supplier_rvp == balance.supplier_rvp
        )
    else:
        source_direction = (
            source_row.document_organization == balance.organization
            and source_row.debit_account is source_account
            and source_row.debit_department == balance.department
            and source_row.debit_supplier_rvp == balance.supplier_rvp
            and source_row.credit_account is manager_account
            and source_row.credit_department == balance.department
            and source_row.credit_supplier_rvp == config.manager_organization
        )
        gk_direction = (
            gk_row.document_organization == config.manager_organization
            and gk_row.debit_account is manager_account
            and gk_row.debit_department == config.manager_financial_department
            and gk_row.debit_supplier_rvp == balance.supplier_rvp
            and gk_row.credit_account is manager_account
            and gk_row.credit_department == config.manager_financial_department
            and gk_row.credit_supplier_rvp == balance.organization
        )
    return source_trace_matches and source_direction and gk_direction


def _is_source_role(balance: NormalizedBalance, row: PostingRow) -> bool:
    """Identify the source-role row by its approved account direction."""

    if balance.ending_side is None:
        return False
    source_account = _source_account_code(balance.source_account)
    manager_account = AccountCode.ACCOUNT_79_1
    if balance.ending_side is BalanceSide.DEBIT:
        return (
            row.debit_account is manager_account and row.credit_account is source_account
        )
    if balance.ending_side is BalanceSide.CREDIT:
        return (
            row.debit_account is source_account and row.credit_account is manager_account
        )
    return False


def _is_gk_role(
    balance: NormalizedBalance, row: PostingRow, config: TransferConfig
) -> bool:
    """Identify the GK-role row by accounts, department, and supplier direction.

    Document organization is deliberately not used as the role discriminator:
    the source organization may itself be the configured GK organization.
    Exact document-organization identity remains part of the full direction
    validation below.
    """

    manager_account = AccountCode.ACCOUNT_79_1
    if not (
        row.debit_account is manager_account
        and row.credit_account is manager_account
        and row.debit_department == config.manager_financial_department
        and row.credit_department == config.manager_financial_department
    ):
        return False
    if balance.ending_side is BalanceSide.DEBIT:
        return (
            row.debit_supplier_rvp == balance.organization
            and row.credit_supplier_rvp == balance.supplier_rvp
        )
    if balance.ending_side is BalanceSide.CREDIT:
        return (
            row.debit_supplier_rvp == balance.supplier_rvp
            and row.credit_supplier_rvp == balance.organization
        )
    return False


def _source_effect_failure_message(control: SourceEffectControl) -> str:
    failed = [
        name
        for name, passed in (
            ("source_effect_zero", control.source_effect_zero),
            ("amounts_match", control.amounts_match),
            ("direction_correct", control.direction_correct),
            ("source_consumed_once", control.source_consumed_once),
        )
        if not passed
    ]
    return "source-effect control failed: " + ", ".join(failed)


def generate_transfer(
    balance: NormalizedBalance, config: TransferConfig | None = None
) -> TransferResult:
    """Generate the approved transfer for one normalized balance."""

    return TransferEngine(config).generate(balance)


def generate_transfer_rows(
    balance: NormalizedBalance, config: TransferConfig | None = None
) -> tuple[PostingRow, ...]:
    """Return exactly two rows for an actionable balance, otherwise no rows."""

    return TransferEngine(config).build_rows(balance)


def generate_transfers(
    balances: Iterable[NormalizedBalance], config: TransferConfig | None = None
) -> TransferBatchResult:
    """Generate transfers for many balances with duplicate-consumption control."""

    return TransferEngine(config).generate_batch(balances)


build_transfer = generate_transfer
build_transfer_rows = generate_transfer_rows
