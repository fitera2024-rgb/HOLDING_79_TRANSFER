# Pilot sequence

The bootstrap creates all five GitHub Issues in an empty repository and labels them `codex-ready`.
Dependencies keep work from starting early.

1. Issue #1: canonical models/output contract.
2. Human reviews Draft PR #1, merges it manually, then marks Issue #1 `accepted` and closes it.
3. Issues #2 (transfer engine) and #3 (OSV parser) may then run in parallel from the updated `main`.
4. After both are manually accepted/merged, Issue #4 exporter may run (it depends on #1/#2).
5. Issue #5 starts only after #2/#3/#4 are accepted and merged.

Dependency is satisfied only when the dependency Issue is BOTH closed AND labelled `accepted`.
This prevents a task from starting merely because someone closed an Issue without accepting its code.

No agent can merge a PR. Human review is a hard gate.
