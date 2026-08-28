---
orchestrator:
  python_venv: true
  test_commands:
    - "{python} -m pytest -q"
    - "{python} -m ruff check ."
  human_review_required: true
  auto_merge: false
---
You are implementing GitHub Issue #{{issue_number}} in HOLDING_79_TRANSFER.

TITLE:
{{issue_title}}

ISSUE BODY:
{{issue_body}}

EXACT BASE SHA:
{{base_sha}}

WORK BRANCH:
{{branch}}

Rules:
1. Read AGENTS.md, docs/PRODUCT_CONTRACT.md, governance/PROJECT_STATE.json first.
2. Verify the current git branch equals WORK BRANCH and treat EXACT BASE SHA as immutable accepted history.
3. Stay strictly inside Issue scope. Do not refactor unrelated code.
4. If source structure or output mapping is ambiguous, fail closed and record a BLOCKED reason; never guess accounting data.
5. Encode expected behavior in tests, including regression/golden tests for financial logic.
6. Make the smallest coherent change.
7. Run repository tests and ruff before finishing.
8. Do not push, create/merge PRs, modify GitHub Issues, release, or access live 1C. The orchestrator owns GitHub handoff.
