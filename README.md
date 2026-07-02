# GSS Survey Workbook Automation

Program-only repository for updating the GSS survey score trends workbook stored in Dropbox.

## Dropbox Layout

- `..\` is the live GSS Surveys folder with the main workbook and uploaded source workbooks.
- `..\Run GSS Update After Upload.cmd` is the operator launcher to run after uploading the newest GSS source workbook.
- `scripts\Update-GSS-MainWorkbook.ps1` is the updater script.
- `..\_automation_runs\` stores generated backups, logs, and copy-test workbooks.

Survey source workbooks, PDFs, screenshots, generated backups, and the main workbook are intentionally not part of this repo.

## What The Updater Does

1. Finds the newest `Sorensen FW...` rolling workbook by the `Date Range` in cell `A1`.
2. Finds the matching prior-year workbook for YoY comparison using newest week minus 364 days.
3. Reads four required rows from each needed week:
   - All Franchisees total
   - Sorensen total
   - 9354 Richmond
   - 9355 Virginia Beach
4. Adds missing rows to `Raw_Data` columns `A:Q`.
5. Rebuilds `Raw_Data` formulas in columns `R:S`.
6. Calculates and saves the workbook with Excel.

## Manual Run

1. Upload or save the newest GSS source workbook into the GSS Surveys Dropbox folder.
2. Close the main workbook if it is open in Excel.
3. Double-click `..\Run GSS Update After Upload.cmd`.

## Safety

- Running the script without `-Apply` creates a copy-only test workbook in `..\_automation_runs\test-output`.
- Running with `-Apply` updates the main workbook and first saves a timestamped backup in `..\_automation_runs\backups`.
- If the newest week and matching prior-year week are already present in `Raw_Data`, the updater skips the write.

## Requirements

- Windows
- Microsoft Excel desktop app
- PowerShell

The disabled Windows scheduled task is named `GSS Survey Main Workbook Weekly Update`. Manual execution is the intended workflow.
