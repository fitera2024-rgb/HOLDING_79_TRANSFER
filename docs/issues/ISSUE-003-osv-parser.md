# [CODEX] Parse grouped 1C OSV into normalized 79.x balances

```yaml
orchestration:
  base_ref: main
  depends_on: [1]
  allow_no_changes: false
```

## Goal
Implement a fail-closed parser for the grouped 1C OSV structure described in PRODUCT_CONTRACT using synthetic XLSX fixtures only.

## Acceptance criteria
- recover hierarchy context: account -> organization -> department -> supplier RVP;
- consume only leaf supplier rows, not account/org/department totals;
- distinguish and ignore technical `ОВ`/`ФВ` duplicate presentation rows correctly;
- read ending Debit/Credit columns by header semantics, not fixed magic row numbers;
- exact accounts only 79.2/79.3;
- generate stable `source_excel_row_ref`;
- missing/ambiguous mandatory context becomes explicit BLOCKED diagnostic;
- tests include debit, credit, zero, blank supplier, duplicate presentation and hierarchy boundary cases;
- no real owner XLSX committed.
