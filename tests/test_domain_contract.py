from datetime import date
from decimal import Decimal
from pathlib import Path

import pytest
import yaml
from pydantic import ValidationError

from holding79_transfer import (
    OUTPUT_HEADERS,
    AccountCode,
    BalanceSide,
    BalanceStatus,
    BlockedOutputMapping,
    BlockReason,
    NormalizedBalance,
    OutputAdapterConfig,
    PostingRow,
    SourceAccount,
    financial_record_id,
    map_posting_row,
    normalize_decimal,
    try_map_posting_row,
    validate_normalized_balance,
)

GOLDENS = Path(__file__).parent / "golden"


def load(name: str) -> dict:
    return yaml.safe_load((GOLDENS / name).read_text(encoding="utf-8"))


def test_source_account_is_an_exact_two_value_enum():
    assert {item.value for item in SourceAccount} == {"79.2", "79.3"}
    with pytest.raises(ValueError):
        SourceAccount("79.20")
    with pytest.raises(ValueError):
        NormalizedBalance(source_account="79.2.1")


def test_monetary_values_are_decimal_safe_and_service_values_are_explicit():
    assert normalize_decimal("84272,40") == Decimal("84272.40")
    assert normalize_decimal("—") == Decimal(0)
    assert normalize_decimal("(10.25)") == Decimal("-10.25")
    with pytest.raises(TypeError):
        normalize_decimal(0.1)
    with pytest.raises(ValidationError):
        NormalizedBalance(source_account="79.2", ending_debit=0.1)


def test_both_sided_ending_balance_is_blocked():
    balance = NormalizedBalance(
        organization="АТ",
        source_account="79.2",
        department="Б_АТ Коммерческий отдел",
        supplier_rvp="Производитель",
        ending_debit="1.00",
        ending_credit="2.00",
    )

    result = validate_normalized_balance(balance)
    assert result.status is BalanceStatus.BLOCKED
    assert result.reason is BlockReason.INVALID_ENDING_BALANCE
    assert balance.status is BalanceStatus.BLOCKED
    assert balance.block_reason is BlockReason.INVALID_ENDING_BALANCE


def test_zero_balance_is_no_action_and_missing_identity_is_blocked():
    zero = NormalizedBalance(
        period_end=date(2024, 12, 31),
        source_excel_row_ref="synthetic:zero",
        organization="АТ",
        source_account="79.3",
        department="Б_АТ Коммерческий отдел",
        supplier_rvp="Производитель",
    )
    assert zero.status is BalanceStatus.NO_ACTION

    missing = NormalizedBalance(
        period_end=date(2024, 12, 31),
        source_excel_row_ref="synthetic:missing",
        source_account="79.3",
        department="Б_АТ Коммерческий отдел",
        supplier_rvp="Производитель",
        ending_debit="1",
    )
    assert missing.status is BalanceStatus.BLOCKED
    assert missing.block_reason is BlockReason.MISSING_SOURCE_IDENTITY


@pytest.mark.parametrize(
    ("department", "supplier"),
    (("", "Производитель"), ("Б_АТ Коммерческий отдел", ""), ("", "")),
)
def test_known_organization_allows_blank_lower_analytics(department, supplier):
    balance = NormalizedBalance(
        period_end=date(2024, 12, 31),
        source_excel_row_ref="synthetic:blank-lower-analytics",
        organization="АТ",
        source_account="79.2",
        department=department,
        supplier_rvp=supplier,
        ending_debit="1.00",
    )

    assert balance.status is BalanceStatus.ACTIONABLE
    assert balance.department == department
    assert balance.supplier_rvp == supplier


@pytest.mark.parametrize(
    ("side", "account_token"),
    [
        ("debit", "79_2"),
        ("credit", "79_2"),
        ("debit", "79_3"),
        ("credit", "79_3"),
    ],
)
def test_all_four_approved_goldens_are_representable_as_canonical_models(side, account_token):
    case = load(f"{side}_{account_token}_AT.yaml")
    source = NormalizedBalance(
        period_end=date(2024, 12, 31),
        source_excel_row_ref=f"synthetic:{case['id']}",
        **case["input"],
    )
    assert source.status is BalanceStatus.ACTIONABLE
    assert source.amount == Decimal("84272.40")
    assert source.ending_side is (
        BalanceSide.DEBIT if side == "debit" else BalanceSide.CREDIT
    )

    record_id = financial_record_id(source)
    rows = []
    for expected in case["expected"]:
        rows.append(
            PostingRow(
                period_end=source.period_end,
                document_organization=expected["document_organization"],
                debit_account=expected["debit_account"],
                debit_department=expected["debit_department"],
                debit_supplier_rvp=expected["debit_supplier_rvp"],
                credit_account=expected["credit_account"],
                credit_department=expected["credit_department"],
                credit_supplier_rvp=expected["credit_supplier_rvp"],
                amount=expected["amount"],
                source_organization=source.organization,
                source_account=source.source_account,
                source_department=source.department,
                source_supplier_rvp=source.supplier_rvp,
                source_excel_row_ref=source.source_excel_row_ref,
                financial_record_id=record_id,
                side=source.ending_side,
            )
        )

    assert len(rows) == 2
    assert all(row.amount == Decimal("84272.40") for row in rows)
    assert all(row.financial_record_id == record_id for row in rows)


def test_financial_record_id_is_deterministic_and_versioned():
    balance = NormalizedBalance(
        period_end=date(2024, 12, 31),
        organization="АТ",
        source_account="79.2",
        department="Б_АТ Коммерческий отдел",
        supplier_rvp="Производитель",
        ending_debit="84272.40",
    )
    assert financial_record_id(balance) == financial_record_id(balance)
    assert financial_record_id(balance, "H79_TRANSFER_V2") != financial_record_id(
        balance, "H79_TRANSFER_V1"
    )


def test_financial_record_id_uses_normalized_decimal_identity():
    common = {
        "period_end": date(2024, 12, 31),
        "organization": "АТ",
        "source_account": "79.2",
        "department": "Б_АТ Коммерческий отдел",
        "supplier_rvp": "Производитель",
    }
    first = NormalizedBalance(**common, ending_debit="84272.40")
    second = NormalizedBalance(**common, ending_debit="84272.400")
    assert financial_record_id(first) == financial_record_id(second)


def test_output_adapter_has_exact_headers_and_explicit_mapping():
    row = PostingRow(
        period_end=date(2024, 12, 31),
        document_organization="АТ",
        debit_account=AccountCode.ACCOUNT_79_1,
        debit_department="Б_АТ Коммерческий отдел",
        debit_supplier_rvp="ГК",
        credit_account=AccountCode.ACCOUNT_79_2,
        credit_department="Б_АТ Коммерческий отдел",
        credit_supplier_rvp="Производитель",
        amount=Decimal("84272.40"),
        source_organization="АТ",
        source_account=SourceAccount.ACCOUNT_79_2,
        source_department="Б_АТ Коммерческий отдел",
        source_supplier_rvp="Производитель",
        source_excel_row_ref="synthetic:42",
        financial_record_id="FR-test",
        side=BalanceSide.DEBIT,
    )

    mapped = map_posting_row(row, OutputAdapterConfig(currency="RUB", run_id="run-1"))
    assert len(OUTPUT_HEADERS) == 27
    assert tuple(mapped) == OUTPUT_HEADERS
    assert mapped["СчетДт"] == "79.1"
    assert mapped["СчетКт"] == "79.2"
    assert mapped["ВалютаДт"] == mapped["ВалютаКт"] == "RUB"
    assert mapped["ВидОперации"] == "REPOST"
    assert mapped["СуммаВВалютеУчета"] == Decimal("84272.40")
    assert mapped["СубконтоДт1"] == "ГК"
    assert mapped["СубконтоКт1"] == "Производитель"
    assert mapped["СчетДтИсточник"] == mapped["СчетКтИсточник"] == "79.2"
    assert mapped["ПравилоДт"] == mapped["ПравилоКт"] == "H79_DEBIT_TRANSFER_V1"
    assert mapped["Содержание"].startswith("run_id=run-1; source_row_ref=synthetic:42")
    assert mapped["КоличествоДт"] is None
    assert mapped["КоличествоКт"] is None


def test_output_mapping_fails_closed_when_traceability_is_incomplete():
    row = PostingRow(
        period_end=date(2024, 12, 31),
        document_organization="АТ",
        debit_account=AccountCode.ACCOUNT_79_1,
        debit_department="Б_АТ Коммерческий отдел",
        debit_supplier_rvp="ГК",
        credit_account=AccountCode.ACCOUNT_79_2,
        credit_department="Б_АТ Коммерческий отдел",
        credit_supplier_rvp="Производитель",
        amount="1.00",
        source_organization="АТ",
        source_account=SourceAccount.ACCOUNT_79_2,
        source_department="Б_АТ Коммерческий отдел",
        source_supplier_rvp="Производитель",
        side=BalanceSide.DEBIT,
        financial_record_id="FR-test",
    )

    result = try_map_posting_row(row)
    assert result.status is BalanceStatus.BLOCKED
    assert result.reason is BlockReason.BLOCKED_INVALID_POSTING
    with pytest.raises(BlockedOutputMapping) as error:
        map_posting_row(row)
    assert error.value.reason is BlockReason.BLOCKED_INVALID_POSTING
