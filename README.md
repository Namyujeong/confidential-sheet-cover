# Confidential Sheet Cover

[English](README.md) | [한국어](README.ko.md)

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

## Google Drive Setup

This tool does not call the Google Drive API directly. It works on local `.xlsx` files. To use it with Google Drive, first make the target workbooks available on your local filesystem.

### Option 1: Google Drive for Desktop

This is the simplest setup.

1. Install Google Drive for Desktop.
2. Sign in to the Google account that has access to the target Drive files.
3. Open the synced Drive folder on your computer.
4. Make sure the target `.xlsx` files are available offline if your Drive client is streaming files.
5. Run this tool against the local synced folder path.

Example:

```bash
confidential-sheet-cover "$HOME/Google Drive/My Drive/Finance" --dry-run
confidential-sheet-cover "$HOME/Google Drive/My Drive/Finance"
```

The exact local Drive path depends on your operating system and Google Drive for Desktop settings.

### Option 2: rclone + Google OAuth

Use this option if you need a repeatable CLI workflow for copying or mounting Drive folders.

High-level setup:

1. Create a Google Cloud project.
2. Enable the Google Drive API for that project.
3. Configure the OAuth consent screen.
4. Create an OAuth Client ID.
5. Choose **Desktop app** as the OAuth client type for local CLI authentication.
6. Install and configure `rclone`.
7. Use `rclone config` to create a Google Drive remote with the OAuth client ID and secret.
8. Copy, sync, or mount the target Drive folder locally.
9. Run this tool against the local copied or mounted folder.

Example rclone flow:

```bash
rclone config
rclone copy mydrive:"Folder/With/Workbooks" ./drive-workbooks
confidential-sheet-cover ./drive-workbooks --dry-run
confidential-sheet-cover ./drive-workbooks
```

If you are using a local OAuth redirect flow, the OAuth client type should usually be **Desktop app**, not **Web application**. A web application client can cause redirect URI mismatch errors in local CLI tools unless the exact local redirect URI is registered.

### Native Google Sheets

Native Google Sheets are not `.xlsx` files. This tool does not edit native Google Sheets directly.

Recommended options:

- Export native Google Sheets to `.xlsx`, run this tool, then upload the updated workbook if that fits your workflow.
- Use Google Drive for Desktop or rclone for Office `.xlsx` files already stored in Drive.
- Build a separate Google Sheets API workflow if you need to modify native Google Sheets in place.

### Service Accounts

A Google service account is not required for this tool.

Use a service account only if you are building an organization-managed automation with the right Google Workspace admin setup, such as domain-wide delegation. For personal or team-level local use, Google Drive for Desktop or an OAuth-based rclone remote is usually simpler.

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
