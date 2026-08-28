# [CODEX] Implement symmetric 79.2/79.3 -> 79.1 transfer engine

```yaml
orchestration:
  base_ref: main
  depends_on: [1]
  allow_no_changes: false
```

## Goal
Implement the approved generic accounting transformation for both debit and credit ending balances on exact source accounts 79.2 and 79.3.

## Acceptance criteria
- every actionable normalized balance produces exactly two PostingRows: source organization + ГК;
- debit source balance follows §5.1 of PRODUCT_CONTRACT;
- credit source balance is exactly symmetric per §5.2;
- 79.2 and 79.3 follow the same direction rules, with only the source account changing;
- all four YAML golden cases drive regression tests;
- source effect control proves the source 79.x balance becomes zero;
- amount preserved exactly as Decimal;
- configuration supplies `ГК` and `Б_ГК Финансовый отдел`; do not scatter literals through domain code;
- no Excel parser/exporter in this Issue.
