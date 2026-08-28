# AGENTS.md — HOLDING_79_TRANSFER

## Mission
Implement only the approved month-end 79.x transfer contract.

## Mandatory before every task
1. Read `docs/PRODUCT_CONTRACT.md`.
2. Read `governance/PROJECT_STATE.json`.
3. Read `WORKFLOW.md`.
4. Read the assigned GitHub Issue.
5. Verify current branch and exact base SHA supplied by the orchestrator.

## Accounting invariants
- No live 1C writes.
- Never invent an accounting direction outside approved contract/goldens.
- Debit and credit ending-balance logic are symmetric and owner-approved.
- The same rules apply to exact source accounts `79.2` and `79.3`.
- No prefix/fuzzy account matching.
- Preserve source organization, department and source `Поставщик РВП` traceability.
- Source org 79.x balance must become zero after the generated source posting.
- For each actionable source row exactly two PostingRows are generated: source organization + ГК.
- `ГК` and `Б_ГК Финансовый отдел` are configuration/reference data, not scattered hardcodes.
- Financial calculations use `Decimal`, never binary float.
- Every accounting rule has regression/golden coverage.
- Real accounting XLSX/MXL files must not be committed.

## Git
- Work only in orchestrator-created task branch/worktree.
- Do not checkout another branch, reset, rebase or change the exact base.
- Do not push or create/merge PRs; orchestrator owns push/PR creation.
- Do not modify unrelated files.

## Completion
- implementation + tests;
- all configured checks green;
- concise handoff in final response;
- STOP.


Output-format reference: `docs/REFERENCE_HAT_OUTPUT_CONTRACT.md`.
