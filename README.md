# GSS Survey Workbook Automation

Program-only repository for updating the GSS survey score trends workbook stored in Dropbox.

## Dropbox Layout

- `..\` is the live GSS Surveys folder with the manual launcher and organized subfolders.
- `..\01 Main Workbook\` stores the live consolidated workbook and older score-trend workbooks.
- `..\02 Weekly Rolling Source Workbooks\` stores the `Sorensen FW...` rolling source workbooks used by the updater.
- `..\03 Uploaded Survey Workbooks\` stores the uploaded `Sorensen...` survey detail workbooks and related transformed files.
- `..\04 Email Comparison PDFs\` stores the generated `GSS Email Comparison MMDDYY.pdf` files.
- `..\05 Reference Materials\` stores training/reference files.
- `..\06 Exports and Images\` stores screenshots and image exports.
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
6. Refreshes the `Email Comparison` tab with driver details for 9355 Virginia Beach and 9354 Richmond.
7. Exports `Email Comparison` to `..\04 Email Comparison PDFs\GSS Email Comparison MMDDYY.pdf`.
8. Calculates and saves the workbook with Excel.

## Manual Run

1. Upload or save the newest GSS source workbook into the GSS Surveys Dropbox folder.
2. Close the main workbook if it is open in Excel.
3. Double-click `..\Run GSS Update After Upload.cmd`.

The updater searches the organized Dropbox subfolders, so the workbook and source files do not need to stay in the folder root.

## Safety

- Running the script without `-Apply` creates a copy-only test workbook in `..\_automation_runs\test-output`.
- Running with `-Apply` updates the main workbook and first saves a timestamped backup in `..\_automation_runs\backups`.
- If the newest week and matching prior-year week are already present in `Raw_Data`, the updater skips the write.

## Requirements

- Windows
- Microsoft Excel desktop app
- PowerShell

The disabled Windows scheduled task is named `GSS Survey Main Workbook Weekly Update`. Manual execution is the intended workflow.
