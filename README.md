# Confidential Sheet Cover

`confidential-sheet-cover` is a small CLI tool that scans local `.xlsx` workbooks for sensitive payroll, salary, labor cost, personnel cost, project cost, or business expense terms, then adds a `Confidential` cover sheet as the first worksheet.

It works with any local folder, including folders synced by Google Drive for Desktop, OneDrive, Dropbox, or a normal filesystem folder.

## What It Does

- Finds `.xlsx` workbooks under one or more folders.
- Scans file paths, sheet names, and cell text.
- Uses built-in English and Korean sensitive-data keywords.
- Adds a first worksheet named `Confidential`.
- Writes a CSV manifest and JSON summary.
- Creates a backup before modifying each workbook.
- Supports `--dry-run` before making changes.

## Default Keywords

The built-in keyword list includes English and Korean terms such as:

```text
payroll, salary, salaries, wage, compensation, bonus, severance,
personnel cost, labor cost, headcount cost, project cost,
business cost, business expense, operating expense, grant budget,
인건비, 연봉, 임금, 급여, 보상, 상여금, 퇴직금, 포괄임금, 사업비, 운영비, 예산
```

You can add more terms with `--keyword` or replace the built-in list with `--no-default-keywords --keywords-file`.

## Install

```bash
python3 -m pip install .
```

Or run directly from the repository:

```bash
python3 -m confidential_sheet_cover --help
```

## Usage

Always start with a dry run.

```bash
python3 -m confidential_sheet_cover "/path/to/folder" --dry-run
```

Apply changes after reviewing `confidential-cover-results/manifest.csv`.

```bash
python3 -m confidential_sheet_cover "/path/to/folder"
```

If installed with `pip install .`, you can use the console command:

```bash
confidential-sheet-cover "/path/to/folder" --dry-run
confidential-sheet-cover "/path/to/folder"
```

## Common Options

```bash
confidential-sheet-cover "/path/to/folder" \
  --dry-run \
  --output-dir confidential-cover-results \
  --max-cells-per-sheet 5000
```

Add custom keywords:

```bash
confidential-sheet-cover "/path/to/folder" \
  --keyword "equity compensation" \
  --keyword "contractor fee"
```

Use a keyword file:

```bash
confidential-sheet-cover "/path/to/folder" \
  --keywords-file keywords.txt
```

Scan file paths and sheet names only:

```bash
confidential-sheet-cover "/path/to/folder" --no-content-scan
```

Customize the cover sheet:

```bash
confidential-sheet-cover "/path/to/folder" \
  --sheet-name "Confidential" \
  --cover-text "Confidential Document - details start on the next sheet"
```

## Outputs

By default, results are written to `confidential-cover-results/`.

```text
confidential-cover-results/
  manifest.csv
  summary.json
  backups/
```

`manifest.csv` includes:

- file path
- status
- matched keywords
- matched locations
- first sheet before and after
- backup path
- SHA-256 before and after
- error message, if any

Common statuses:

| Status | Meaning |
| --- | --- |
| `dry_run_target` | The workbook matched but was not modified because `--dry-run` was used. |
| `updated` | A new confidential cover sheet was added. |
| `moved_existing_cover` | An existing matching cover sheet was moved to the front. |
| `skipped_already_covered` | The workbook already had the expected cover as the first sheet. |
| `skipped_no_match` | No sensitive keyword was found. |
| `error` | The workbook could not be scanned or updated. |

## Safety Notes

- Run `--dry-run` first.
- Review `manifest.csv` before applying changes.
- Backups are created before each modified workbook is saved.
- This tool supports `.xlsx` files. Native Google Sheets must be exported to `.xlsx` before processing.
- Legacy `.xls` files should be converted to `.xlsx` before processing.
- `openpyxl` may not preserve every advanced Excel feature. Test on a copy before using it on critical workbooks.

## License

MIT
