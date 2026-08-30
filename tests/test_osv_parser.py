from __future__ import annotations

from datetime import date
from decimal import Decimal

from openpyxl import Workbook
from openpyxl.styles import Alignment

from holding79_transfer import (
    BalanceStatus,
    ParserDiagnosticCode,
    parse_grouped_osv,
)


def make_workbook(
    rows: list[tuple[str, object, object, int | None]],
    *,
    header_row: int = 4,
    sheet_name: str = "ОСВ",
) -> Workbook:
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = sheet_name
    worksheet.cell(1, 1).value = "Синтетическая ОСВ"
    worksheet.cell(2, 1).value = "Период: 01.12.2024 - 31.12.2024"
    worksheet.merge_cells(
        start_row=header_row,
        start_column=8,
        end_row=header_row,
        end_column=9,
    )
    worksheet.cell(header_row, 8).value = "Сальдо на конец периода"
    worksheet.cell(header_row + 1, 8).value = "Дебет"
    worksheet.cell(header_row + 1, 9).value = "Кредит"

    for offset, (grouping, debit, credit, indent) in enumerate(rows, header_row + 2):
        cell = worksheet.cell(offset, 1)
        cell.value = grouping
        if indent is not None:
            cell.alignment = Alignment(indent=indent)
        worksheet.cell(offset, 8).value = debit
        worksheet.cell(offset, 9).value = credit
    return workbook


def parse(rows: list[tuple[str, object, object, int | None]], **kwargs):
    return parse_grouped_osv(make_workbook(rows, **kwargs), period_end=date(2024, 12, 31))


def leaf(
    account: str,
    organization: str,
    department: str,
    supplier: str,
    debit: object = 0,
    credit: object = 0,
) -> list[tuple[str, object, object, int | None]]:
    return [
        (account, None, None, 0),
        (f"Организация: {organization}", None, None, 1),
        (f"ЦФО: {department}", None, None, 2),
        (f"Поставщик РВП: {supplier}", debit, credit, 3),
    ]


def test_supported_accounts_and_both_ending_sides_are_normalized():
    rows = leaf("79.2", "АТ", "Б_АТ Коммерческий отдел", "Дебетовый", "84272.40")
    rows += [("Поставщик РВП: Кредитовый", 0, "84272.40", 3)]
    rows += leaf("79.3", "БТ", "Б_БТ Финансы", "Дебетовый", "1.25")
    rows += [("Поставщик РВП: Кредитовый", 0, "2.50", 3)]

    result = parse(rows)

    assert result.status is BalanceStatus.ACTIONABLE
    assert len(result.balances) == 4
    assert [balance.source_account.value for balance in result.balances] == [
        "79.2",
        "79.2",
        "79.3",
        "79.3",
    ]
    assert result.balances[0].ending_debit == Decimal("84272.40")
    assert result.balances[1].ending_credit == Decimal("84272.40")
    assert result.balances[2].ending_debit == Decimal("1.25")
    assert result.balances[3].ending_credit == Decimal("2.50")
    assert all(isinstance(balance.ending_debit, Decimal) for balance in result.balances)


def test_zero_balance_is_no_action_and_keeps_canonical_identity():
    result = parse(leaf("79.2", "АТ", "ЦФО", "Нулевой"))

    assert result.status is BalanceStatus.NO_ACTION
    assert len(result.balances) == 1
    assert result.balances[0].status is BalanceStatus.NO_ACTION


def test_technical_ov_fv_rows_and_identical_presentation_rows_do_not_duplicate_leaf():
    rows = leaf("79.2", "АТ", "ЦФО", "Производитель", "10.00")
    rows += [
        ("ОВ", "10.00", 0, 4),
        ("ФВ", "10.00", 0, 4),
        ("Итого по поставщику", "10.00", 0, 3),
        ("Поставщик РВП: Производитель", "10.00", 0, 3),
    ]

    result = parse(rows)

    assert result.status is BalanceStatus.ACTIONABLE
    assert len(result.balances) == 1
    assert result.balances[0].supplier_rvp == "Производитель"


def test_indentation_recovers_supplier_department_and_organization_boundaries():
    rows = [
        ("79.2", None, None, 0),
        ("АТ", None, None, 1),
        ("Коммерческий отдел", None, None, 2),
        ("Поставщик-1", "1.00", 0, 3),
        ("Поставщик-2", 0, "2.00", 3),
        ("Финансовый отдел", None, None, 2),
        ("Поставщик-3", "3.00", 0, 3),
        ("БТ", None, None, 1),
        ("Коммерческий отдел", None, None, 2),
        ("Поставщик-4", 0, "4.00", 3),
    ]

    result = parse(rows)

    assert result.status is BalanceStatus.ACTIONABLE
    assert [(b.organization, b.department, b.supplier_rvp) for b in result.balances] == [
        ("АТ", "Коммерческий отдел", "Поставщик-1"),
        ("АТ", "Коммерческий отдел", "Поставщик-2"),
        ("АТ", "Финансовый отдел", "Поставщик-3"),
        ("БТ", "Коммерческий отдел", "Поставщик-4"),
    ]


def test_account_boundary_resets_stale_context_and_missing_context_blocks():
    rows = leaf("79.2", "АТ", "ЦФО-1", "Поставщик-1", "1.00")
    rows += [
        ("79.3", None, None, 0),
        ("ЦФО: ЦФО-2", None, None, 2),
        ("Поставщик РВП: Поставщик-2", "2.00", 0, 3),
    ]

    result = parse(rows)

    assert result.status is BalanceStatus.BLOCKED
    assert result.balances == ()
    assert result.diagnostics[0].code is ParserDiagnosticCode.MISSING_ORGANIZATION


def test_blank_supplier_and_ambiguous_identity_are_blocked():
    blank = parse(leaf("79.2", "АТ", "ЦФО", "", "1.00"))
    assert blank.status is BalanceStatus.BLOCKED
    assert blank.diagnostics[0].code is ParserDiagnosticCode.MISSING_SUPPLIER_RVP

    ambiguous_rows = leaf("79.2", "АТ", "ЦФО", "Производитель", "1.00")
    workbook = make_workbook(ambiguous_rows)
    worksheet = workbook.active
    worksheet.cell(worksheet.max_row - 2, 1).value = "Организация"
    worksheet.cell(worksheet.max_row - 2, 2).value = "АТ"
    worksheet.cell(worksheet.max_row - 2, 3).value = "БТ"
    ambiguous = parse_grouped_osv(workbook, period_end=date(2024, 12, 31))
    assert ambiguous.status is BalanceStatus.BLOCKED
    assert ambiguous.diagnostics[0].code is ParserDiagnosticCode.AMBIGUOUS_ORGANIZATION


def test_unsupported_account_is_not_accepted_or_inherited():
    result = parse(leaf("79.20", "АТ", "ЦФО", "Поставщик", "1.00"))

    assert result.status is BalanceStatus.BLOCKED
    assert result.balances == ()
    assert result.diagnostics[0].code is ParserDiagnosticCode.UNSUPPORTED_ACCOUNT


def test_semantic_header_detection_is_independent_of_header_row_number():
    result = parse(
        leaf("79.3", "АТ", "ЦФО", "Поставщик", "12.30"),
        header_row=12,
    )

    assert result.status is BalanceStatus.ACTIONABLE
    assert result.balances[0].ending_debit == Decimal("12.30")


def test_account_label_and_exact_account_value_can_share_grouping_columns():
    workbook = make_workbook(leaf("79.2", "АТ", "ЦФО", "Поставщик", "12.30"))
    worksheet = workbook.active
    worksheet.cell(6, 1).value = "Счет"
    worksheet.cell(6, 2).value = "79.2"

    result = parse_grouped_osv(workbook, period_end=date(2024, 12, 31))

    assert result.status is BalanceStatus.ACTIONABLE
    assert result.balances[0].source_account.value == "79.2"
    assert result.balances[0].organization == "АТ"


def test_missing_or_ambiguous_ending_balance_headers_block():
    workbook = make_workbook(leaf("79.2", "АТ", "ЦФО", "Поставщик", "1.00"))
    workbook.active.cell(5, 9).value = None
    missing = parse_grouped_osv(workbook, period_end=date(2024, 12, 31))
    assert missing.status is BalanceStatus.BLOCKED
    assert missing.diagnostics[0].code is ParserDiagnosticCode.MISSING_ENDING_BALANCE_HEADERS

    ambiguous_workbook = make_workbook(leaf("79.2", "АТ", "ЦФО", "Поставщик", "1.00"))
    ambiguous_workbook.active.merge_cells("K4:L4")
    ambiguous_workbook.active["K4"] = "Сальдо на конец периода"
    ambiguous_workbook.active["K5"] = "Дебет"
    ambiguous_workbook.active["L5"] = "Кредит"
    ambiguous = parse_grouped_osv(ambiguous_workbook, period_end=date(2024, 12, 31))
    assert ambiguous.status is BalanceStatus.BLOCKED
    assert ambiguous.diagnostics[0].code is ParserDiagnosticCode.AMBIGUOUS_ENDING_BALANCE_HEADERS


def test_source_row_reference_is_stable_and_path_independent():
    first = parse(leaf("79.2", "АТ", "ЦФО", "Поставщик", "1.00"))
    second = parse(leaf("79.2", "АТ", "ЦФО", "Поставщик", "1.00"))

    assert first.balances[0].source_excel_row_ref == second.balances[0].source_excel_row_ref
    assert first.balances[0].source_excel_row_ref == "ОСВ!R9"


def test_decimal_safe_service_and_negative_representations_are_normalized():
    result = parse(leaf("79.2", "АТ", "ЦФО", "Поставщик", "1 234,50"))
    assert result.status is BalanceStatus.ACTIONABLE
    assert result.balances[0].ending_debit == Decimal("1234.50")

    negative = parse(leaf("79.3", "АТ", "ЦФО", "Поставщик", "-7.25"))
    assert negative.status is BalanceStatus.ACTIONABLE
    assert negative.balances[0].ending_debit == Decimal(0)
    assert negative.balances[0].ending_credit == Decimal("7.25")

    both_sides = parse(leaf("79.3", "АТ", "ЦФО", "Поставщик", "7.25", "1.00"))
    assert both_sides.status is BalanceStatus.BLOCKED
    assert both_sides.diagnostics[0].code is ParserDiagnosticCode.INVALID_ENDING_BALANCE
