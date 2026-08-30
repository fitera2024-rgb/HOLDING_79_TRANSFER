# Issue #13 pilot acceptance package

This is a local, synthetic dry run for the accepted pipeline:

`grouped OSV parser -> transfer engine -> accounting controls -> exact 27-column XLSX exporter`

The package is based on the exact accepted `main` SHA
`296448216a62249eeeae15a40fd41234d5099c5a` and Product Contract
`0.3-approved`.

## Safety boundary

The pilot command is permitted to read the repository, its synthetic fixtures,
and generated local output. It does not accept or require a client XLSX/MXL/ZIP,
credentials, a 1C connection, or any external accounting source. Real client
files must not be placed in the command's input or committed to Git.

Live 1C is not touched. No import into 1C, release, deployment, or live write
is performed.

## Prerequisites

- Clean checkout at the Issue #13 branch/PR head, with the accepted base SHA
  in history.
- Python 3.10 or newer.
- Project and test dependencies installed:

  `python -m pip install -e ".[test]"`

The command uses only the synthetic workbook built by
`holding79_transfer.build_synthetic_osv_workbook`. It writes only to the
dedicated output directory supplied by `--output-dir`, which must not already
exist.

## Exact dry-run command

From the repository root, use a clean output directory:

`python scripts/run_pilot_acceptance.py --output-dir .pilot-acceptance`

The command verifies that the accepted base SHA is an ancestor of the current
build HEAD when the commit is available locally. In a shallow GitHub PR
checkout, it verifies the workflow base ref is `main` without fetching or
mutating the checkout. It then returns zero and prints `ACCEPTANCE=PASS` only
after all mandatory controls, malformed-source probes, atomicity checks, and
the deterministic rerun comparison pass. Any mandatory failure returns
nonzero and no successful acceptance directory is published.

## Successful artifact set

The command publishes exactly this root set, plus the two export workbooks:

- `input_manifest.json` — synthetic source identification, source sheets,
  normalized-input SHA-256, parser diagnostics, and run configuration;
- `normalized_balances.jsonl` — canonical Decimal balance rows;
- `posting_rows.jsonl` — canonical financial PostingRows and deterministic
  rule/financial IDs;
- `summary.json` — counts, Decimal totals, and mandatory control results;
- `run_control.xlsx` — operator audit workbook;
- `export_manifest.json` — output contract, deterministic file names, and
  round-trip results;
- `export/*.xlsx` — one import workbook per document date and organization.

`run_control.xlsx` contains exactly these control sheets:

`Итоги`, `Параметры_запуска`, `Остатки_79`, `Готовые_проводки`, `Блокировки`,
`Контроль_до_после`, `Исходные_строки`, `Проверка_экспорта`.

Each import workbook contains only `Загрузка_A_AA`, exactly 27 columns in the
Product Contract order, and no helper sheets or columns.

## Mandatory accounting controls

The runner proves all of the following before publication:

- all four approved golden cases are unchanged: debit and credit for `79.2`
  and `79.3`;
- source accounts are exactly `79.2` and `79.3`;
- every actionable source balance produces exactly two PostingRows;
- source-side `79.2`/`79.3` after-effect is zero;
- source-organization and `ГК` generated totals are equal;
- the intentionally blocked source row produces zero financial output;
- financial amounts remain Decimal text and no binary-float accounting value is
  introduced;
- deterministic rule IDs and financial IDs are present;
- the exact `Загрузка_A_AA` 27-column contract is preserved;
- export workbooks have no helper sheets and reopen successfully through the
  exporter round-trip validator.

The blocked synthetic row is retained in parser diagnostics and control sheets
with zero financial posting/export counts. It is never mapped to an import
workbook.

## Deterministic rerun procedure

The command feeds the same synthetic XLSX bytes and the same configuration to
two independent local runs before publishing the first result. It compares:

- normalized balance JSONL bytes and PostingRows JSONL bytes;
- run ID and normalized-input SHA-256;
- PostingRows, rule IDs, financial IDs, Decimal amounts, and source references;
- deterministic export manifest paths and grouping;
- canonical exported row values and row order for every workbook;
- financial totals.

XLSX ZIP metadata is not treated as financial meaning. It may vary, but it must
not change financial identity, content, or export row order. The second run is
temporary and is removed after comparison.

The command also probes ordinary malformed XLSX/ZIP bytes and malformed
worksheet `sheetId` type metadata. Each must return `BLOCKED` with
`INVALID_SOURCE`, including the exact fail-closed message, without exposing a
raw workbook exception.

## STOP conditions

Stop and report `BLOCKED` if any of the following occurs:

- the base SHA, Product Contract version, or branch is not the accepted one;
- a mandatory artifact, sheet, column, row, ID, or control is missing or does
  not reconcile;
- a golden result changes;
- a blocked source produces financial output;
- a rerun changes financial meaning, IDs, filenames, or export row order;
- malformed XLSX escapes as an exception or is not classified as
  `BLOCKED`/`INVALID_SOURCE`;
- the deliberate failure probe returns zero or leaves a published output;
- repository tests, Ruff, or Product CI fails.

Do not reinterpret accounting direction or repair a business-rule ambiguity in
this package. Such a change requires `OWNER_DECISION_REQUIRED` and a new
contract decision.

## Cleanup

After recording the result, remove only the dedicated synthetic output:

PowerShell:

`Remove-Item -LiteralPath .pilot-acceptance -Recurse -Force`

The command itself cleans its private rerun and failure-probe directories.
Never point cleanup at a repository root, a home directory, or a directory
containing client data.
