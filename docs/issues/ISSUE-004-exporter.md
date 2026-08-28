# [CODEX] Export PostingRows to standard 27-column IFRS workbooks

```yaml
orchestration:
  base_ref: main
  depends_on: [1, 2]
  allow_no_changes: false
```

## Goal
Export PostingRows to one XLSX per document organization/date with sheet `Загрузка_A_AA` and exactly the accepted 27 columns.

## Acceptance criteria
- exact header order, no helper sheet inside import workbook;
- deterministic filename and row order;
- source-organization and ГК PostingRows go to their respective organization files;
- configurable operation type, pilot default `REPOST`;
- deterministic rule ids and financial ids;
- reopen every produced workbook and round-trip compare to internal PostingRows;
- round-trip mismatch blocks success;
- synthetic data only, no live write.
