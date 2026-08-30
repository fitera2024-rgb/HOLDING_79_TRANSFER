# Pilot acceptance operator checklist

Use this checklist for the exact PR head. Check every item before sign-off.

## Build and source

- [ ] Repository: `fitera2024-rgb/HOLDING_79_TRANSFER`
- [ ] Base SHA: `296448216a62249eeeae15a40fd41234d5099c5a`
- [ ] HEAD SHA: `________________` (record the exact reviewed PR head)
- [ ] Product Contract: `0.3-approved`
- [ ] Branch/PR is the Issue #13 pilot-acceptance branch
- [ ] Synthetic source identification: `synthetic-pilot-acceptance.xlsx`
- [ ] Source sheets: four golden sheets plus `BLOCKED_SOURCE_ROW`
- [ ] `input_manifest.json` `normalized_input_sha256`: `________________`
- [ ] No client XLSX/MXL/ZIP, secrets, credentials, or sensitive paths used

## Run/config identification

- [ ] Command used exactly:
      `python scripts/run_pilot_acceptance.py --output-dir .pilot-acceptance`
- [ ] `run_id`: `________________`
- [ ] `period_end`: `2024-12-31`
- [ ] `rules_version`: `H79_TRANSFER_V1`
- [ ] Manager organization: `ГК`
- [ ] Manager financial department: `Б_ГК Финансовый отдел`
- [ ] Operation type: `REPOST`

## Artifact completeness

- [ ] `input_manifest.json`
- [ ] `normalized_balances.jsonl`
- [ ] `posting_rows.jsonl`
- [ ] `summary.json`
- [ ] `run_control.xlsx`
- [ ] `export_manifest.json`
- [ ] `export/*.xlsx`
- [ ] `run_control.xlsx` has exactly: `Итоги`, `Параметры_запуска`,
      `Остатки_79`, `Готовые_проводки`, `Блокировки`, `Контроль_до_после`,
      `Исходные_строки`, `Проверка_экспорта`

## Accounting controls

- [ ] Four approved goldens unchanged: `DEBIT_79_2_AT`, `CREDIT_79_2_AT`,
      `DEBIT_79_3_AT`, `CREDIT_79_3_AT`
- [ ] Source accounts are exactly `79.2` / `79.3`
- [ ] Every actionable source is exactly 1:2 to source-organization + `ГК`
      PostingRows
- [ ] Source-side 79.2/79.3 after-effect is zero for every actionable row
- [ ] Source-organization total equals `ГК` total; difference is `0`
- [ ] Blocked source row has zero PostingRows and zero export rows
- [ ] Decimal semantics are preserved; no float accounting values accepted

## Export validation

- [ ] Every import workbook has exactly the worksheet `Загрузка_A_AA`
- [ ] Every import workbook has exactly 27 columns in contract order
- [ ] No helper columns or sheets exist in an import workbook
- [ ] Deterministic filenames are present and manifest paths match files
- [ ] Every export workbook reopens and round-trips to PostingRows

## Regression, rerun, and atomicity

- [ ] Same synthetic input/config rerun has identical normalized financial
      meaning
- [ ] PostingRows, rule IDs, financial IDs, deterministic filenames, and export
      row order are identical
- [ ] Ordinary malformed XLSX/ZIP bytes are `BLOCKED` / `INVALID_SOURCE`
- [ ] Invalid worksheet `sheetId` type metadata is `BLOCKED` /
      `INVALID_SOURCE`; no raw workbook exception escapes
- [ ] Deliberate mandatory failure returns nonzero
- [ ] Failed run reports no success and publishes no partial financial output

## Final sign-off

- [ ] `python -m pytest -q` — PASS, count: `________`
- [ ] `python -m ruff check .` — PASS
- [ ] `git diff --check` — PASS
- [ ] `python -c "import holding79_transfer"` — PASS
- [ ] Product CI completed and succeeded on the exact HEAD
- [ ] Live 1C: not accessed
- [ ] Release/deployment: not performed
- [ ] Operator: `________________`
- [ ] Date/time: `________________`
- [ ] Final result: `READY_FOR_INDEPENDENT_ACCEPTANCE_REVIEW`
