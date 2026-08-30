from datetime import date
from decimal import Decimal
from pathlib import Path

import pytest
import yaml

from holding79_transfer import (
    CREDIT_RULE_ID,
    DEBIT_RULE_ID,
    MANAGER_FINANCIAL_DEPARTMENT,
    MANAGER_ORGANIZATION,
    RULES_VERSION,
    BalanceSide,
    BalanceStatus,
    BlockReason,
    NormalizedBalance,
    SourceAccount,
    TransferConfig,
    generate_transfer,
    generate_transfers,
    validate_source_effect,
)

GOLDENS = Path(__file__).parent / "golden"


def load(name: str) -> dict:
    return yaml.safe_load((GOLDENS / name).read_text(encoding="utf-8"))


def make_balance(case: dict) -> NormalizedBalance:
    return NormalizedBalance(
        period_end=date(2024, 12, 31),
        source_excel_row_ref=f"synthetic:{case['id']}",
        **case["input"],
    )


@pytest.mark.parametrize(
    "filename",
    [
        "debit_79_2_AT.yaml",
        "credit_79_2_AT.yaml",
        "debit_79_3_AT.yaml",
        "credit_79_3_AT.yaml",
    ],
)
def test_approved_golden_cases_generate_exactly_two_posting_rows(filename: str):
    case = load(filename)
    balance = make_balance(case)

    result = generate_transfer(balance)

    assert result.status is BalanceStatus.ACTIONABLE
    assert len(result.rows) == 2
    for row, expected in zip(result.rows, case["expected"], strict=True):
        assert row.document_organization == expected["document_organization"]
        assert row.debit_account.value == expected["debit_account"]
        assert row.debit_department == expected["debit_department"]
        assert row.debit_supplier_rvp == expected["debit_supplier_rvp"]
        assert row.credit_account.value == expected["credit_account"]
        assert row.credit_department == expected["credit_department"]
        assert row.credit_supplier_rvp == expected["credit_supplier_rvp"]
        assert row.amount == Decimal(expected["amount"])
        assert row.source_organization == balance.organization
        assert row.source_account is balance.source_account
        assert row.source_department == balance.department
        assert row.source_supplier_rvp == balance.supplier_rvp
        assert row.source_excel_row_ref == balance.source_excel_row_ref
        assert row.side is balance.ending_side
        assert row.rules_version == RULES_VERSION
        assert row.rule_id == (
            DEBIT_RULE_ID if balance.ending_side is BalanceSide.DEBIT else CREDIT_RULE_ID
        )

    assert result.rows[0].document_organization == balance.organization
    assert result.rows[1].document_organization == MANAGER_ORGANIZATION
    assert result.rows[1].debit_department == MANAGER_FINANCIAL_DEPARTMENT
    assert result.rows[0].amount == result.rows[1].amount == Decimal("84272.40")


@pytest.mark.parametrize("source_account", [SourceAccount.ACCOUNT_79_2, SourceAccount.ACCOUNT_79_3])
@pytest.mark.parametrize("side", [BalanceSide.DEBIT, BalanceSide.CREDIT])
def test_source_effect_control_proves_zero_balance_and_single_consumption(
    source_account: SourceAccount, side: BalanceSide
):
    balance = NormalizedBalance(
        period_end=date(2024, 12, 31),
        source_excel_row_ref="synthetic:effect",
        organization="АТ",
        source_account=source_account,
        department="Б_АТ Коммерческий отдел",
        supplier_rvp="Производитель",
        ending_debit="84272.40" if side is BalanceSide.DEBIT else "0",
        ending_credit="84272.40" if side is BalanceSide.CREDIT else "0",
    )

    result = generate_transfer(balance)

    assert result.source_effect is not None
    control = result.source_effect
    assert control.passed
    assert control.source_effect_zero
    assert control.ending_debit_after == Decimal(0)
    assert control.ending_credit_after == Decimal(0)
    assert control.source_balance_after == Decimal(0)
    assert control.source_posting_count == 1
    assert control.source_consumed_once
    assert control.amounts_match
    assert control.direction_correct


@pytest.mark.parametrize(
    ("source_account", "side"),
    [
        (SourceAccount.ACCOUNT_79_2, BalanceSide.DEBIT),
        (SourceAccount.ACCOUNT_79_2, BalanceSide.CREDIT),
        (SourceAccount.ACCOUNT_79_3, BalanceSide.DEBIT),
        (SourceAccount.ACCOUNT_79_3, BalanceSide.CREDIT),
    ],
)
def test_manager_organization_is_valid_source_for_both_accounts_and_sides(
    source_account: SourceAccount, side: BalanceSide
):
    amount = Decimal("1234.567")
    balance = NormalizedBalance(
        period_end=date(2024, 12, 31),
        source_excel_row_ref=f"synthetic:gk-source:{source_account.value}:{side.value}",
        organization=MANAGER_ORGANIZATION,
        source_account=source_account,
        department="Б_ГК Коммерческий отдел",
        supplier_rvp="Производитель",
        ending_debit=amount if side is BalanceSide.DEBIT else Decimal(0),
        ending_credit=amount if side is BalanceSide.CREDIT else Decimal(0),
    )

    result = generate_transfer(balance)

    assert balance.status is BalanceStatus.ACTIONABLE
    assert result.status is BalanceStatus.ACTIONABLE
    assert len(result.rows) == 2
    assert result.source_effect is not None
    assert result.source_effect.passed
    assert result.source_effect.ending_debit_after == Decimal(0)
    assert result.source_effect.ending_credit_after == Decimal(0)
    assert result.source_effect.source_balance_after == Decimal(0)
    assert result.source_effect.source_posting_count == 1
    assert result.source_effect.source_consumed_once
    assert result.source_effect.amounts_match
    assert result.source_effect.direction_correct

    source_row, gk_row = result.rows
    assert source_row.document_organization == MANAGER_ORGANIZATION
    assert gk_row.document_organization == MANAGER_ORGANIZATION
    assert source_row.amount == gk_row.amount == amount
    assert isinstance(source_row.amount, Decimal)
    assert isinstance(gk_row.amount, Decimal)
    assert gk_row.debit_account.value == "79.1"
    assert gk_row.credit_account.value == "79.1"
    assert gk_row.debit_department == MANAGER_FINANCIAL_DEPARTMENT
    assert gk_row.credit_department == MANAGER_FINANCIAL_DEPARTMENT

    if side is BalanceSide.DEBIT:
        assert source_row.debit_account.value == "79.1"
        assert source_row.credit_account.value == source_account.value
        assert source_row.debit_supplier_rvp == MANAGER_ORGANIZATION
        assert source_row.credit_supplier_rvp == balance.supplier_rvp
        assert gk_row.debit_supplier_rvp == balance.organization
        assert gk_row.credit_supplier_rvp == balance.supplier_rvp
    else:
        assert source_row.debit_account.value == source_account.value
        assert source_row.credit_account.value == "79.1"
        assert source_row.debit_supplier_rvp == balance.supplier_rvp
        assert source_row.credit_supplier_rvp == MANAGER_ORGANIZATION
        assert gk_row.debit_supplier_rvp == balance.supplier_rvp
        assert gk_row.credit_supplier_rvp == balance.organization


def test_reference_configuration_supplies_manager_identity():
    balance = NormalizedBalance(
        period_end=date(2024, 12, 31),
        source_excel_row_ref="synthetic:config",
        organization="АТ",
        source_account="79.2",
        department="Б_АТ Коммерческий отдел",
        supplier_rvp="Производитель",
        ending_debit="1.00",
    )
    config = TransferConfig(
        manager_organization="MANAGER",
        manager_financial_department="MANAGER_FINANCE",
    )

    result = generate_transfer(balance, config)

    assert result.status is BalanceStatus.ACTIONABLE
    source_row, manager_row = result.rows
    assert source_row.debit_supplier_rvp == "MANAGER"
    assert manager_row.document_organization == "MANAGER"
    assert manager_row.debit_department == manager_row.credit_department == "MANAGER_FINANCE"


def test_unsupported_rules_version_fails_closed_without_postings():
    balance = NormalizedBalance(
        period_end=date(2024, 12, 31),
        source_excel_row_ref="synthetic:unsupported-rules-version",
        organization="АТ",
        source_account=SourceAccount.ACCOUNT_79_2,
        department="Б_АТ Коммерческий отдел",
        supplier_rvp="Производитель",
        ending_debit=Decimal("1.00"),
    )

    result = generate_transfer(
        balance,
        TransferConfig(rules_version="H79_TRANSFER_V2"),
    )

    assert result.status is BalanceStatus.BLOCKED
    assert result.reason is BlockReason.INVALID_POSTING
    assert result.rows == ()
    assert result.message == (
        "unsupported rules_version for transfer engine: H79_TRANSFER_V2"
    )


def test_zero_and_invalid_balances_generate_no_rows():
    common = {
        "period_end": date(2024, 12, 31),
        "source_excel_row_ref": "synthetic:zero",
        "organization": "АТ",
        "source_account": "79.3",
        "department": "Б_АТ Коммерческий отдел",
        "supplier_rvp": "Производитель",
    }
    zero = NormalizedBalance(**common)
    invalid = NormalizedBalance(
        **(common | {"source_excel_row_ref": "synthetic:invalid"}),
        ending_debit="1.00",
        ending_credit="2.00",
    )

    zero_result = generate_transfer(zero)
    invalid_result = generate_transfer(invalid)

    assert zero_result.status is BalanceStatus.NO_ACTION
    assert zero_result.rows == ()
    assert invalid_result.status is BalanceStatus.BLOCKED
    assert invalid_result.reason is BlockReason.INVALID_ENDING_BALANCE
    assert invalid_result.rows == ()


def test_source_effect_control_fails_closed_for_wrong_amount_or_duplicate_source_row():
    case = load("debit_79_2_AT.yaml")
    balance = make_balance(case)
    result = generate_transfer(balance)
    source_row, gk_row = result.rows
    wrong_amount = source_row.model_copy(update={"amount": Decimal("1.00")})

    wrong_amount_control = validate_source_effect(balance, (wrong_amount, gk_row))
    duplicate_control = validate_source_effect(balance, (source_row, source_row, gk_row))

    assert not wrong_amount_control.passed
    assert not wrong_amount_control.source_effect_zero
    assert not wrong_amount_control.amounts_match
    assert not duplicate_control.passed
    assert duplicate_control.source_posting_count == 2
    assert not duplicate_control.source_consumed_once


def test_source_effect_control_rejects_rules_version_and_trace_mismatches():
    balance = make_balance(load("debit_79_2_AT.yaml"))
    result = generate_transfer(balance)
    source_row, gk_row = result.rows

    assert validate_source_effect(balance, (source_row, gk_row)).passed

    wrong_source_rules = source_row.model_copy(
        update={"rules_version": "WRONG_RULES"}
    )
    wrong_gk_rules = gk_row.model_copy(update={"rules_version": "WRONG_RULES"})
    mixed_rules = gk_row.model_copy(update={"rules_version": "ANOTHER_RULES"})
    wrong_record_id = source_row.model_copy(
        update={"financial_record_id": "FR-wrong"}
    )

    assert not validate_source_effect(
        balance, (wrong_source_rules, gk_row)
    ).passed
    assert not validate_source_effect(balance, (source_row, wrong_gk_rules)).passed
    assert not validate_source_effect(balance, (source_row, mixed_rules)).passed
    assert not validate_source_effect(
        balance, (wrong_record_id, gk_row)
    ).passed


def test_batch_rejects_reusing_one_source_row_reference_and_returns_no_partial_rows():
    balance = make_balance(load("debit_79_3_AT.yaml"))

    result = generate_transfers((balance, balance.model_copy()))

    assert result.status is BalanceStatus.BLOCKED
    assert result.rows == ()
    assert result.transfers[1].reason is BlockReason.INVALID_POSTING
