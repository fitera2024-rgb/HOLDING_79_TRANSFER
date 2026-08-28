# [CODEX] End-to-end run artifacts and accounting controls

```yaml
orchestration:
  base_ref: main
  depends_on: [2, 3, 4]
  allow_no_changes: false
```

## Goal
Connect OSV parsing -> transfer engine -> controls -> 27-column export as one deterministic local run.

## Acceptance criteria
- create all run artifacts listed in PRODUCT_CONTRACT;
- `run_control.xlsx` contains the required control sheets;
- every actionable source row reconciles 1:1 to source-org + ГК PostingRows;
- source 79.2/79.3 after-effect is zero for every generated transfer;
- total source-org generated amount equals total ГК generated amount by run;
- blocked source rows produce zero financial output rows;
- idempotent rerun of same normalized input + config produces identical financial rows/ids;
- integration test uses synthetic OSV covering all four goldens and at least one blocked row;
- no UI and no live 1C in this issue.
