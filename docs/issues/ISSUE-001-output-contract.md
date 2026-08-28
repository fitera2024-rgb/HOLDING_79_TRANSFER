# [CODEX] Canonical domain model and 27-column output adapter contract

```yaml
orchestration:
  base_ref: main
  depends_on: []
  allow_no_changes: false
```

## Goal
Implement the canonical domain types (`NormalizedBalance`, `PostingRow`, statuses/block reasons), exact 27-column header constant/order, deterministic IDs, and pure validation/mapping helpers. Do not read Excel yet.

## Acceptance criteria
- Decimal-safe monetary values;
- exact source-account enum `79.2`/`79.3`;
- exact 27 headers from `docs/PRODUCT_CONTRACT.md`;
- deterministic financial-record id from normalized business identity + rules version;
- model-level tests can represent all four approved golden cases;
- invalid both-sided ending balance is blocked;
- no XLSX parsing, no filesystem export, no live 1C.
