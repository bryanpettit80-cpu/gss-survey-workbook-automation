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
- `Run-GSS-Workbook-Test.cmd` runs a copy-test only.
- `Run-GSS-Workbook-Update.cmd` runs the same safe copy-test-then-confirm flow as the operator launcher.
- `scripts\Update-GSS-MainWorkbook.ps1` is the updater script.
- `..\_automation_runs\` stores generated backups, logs, and copy-test workbooks.

Survey source workbooks, PDFs, screenshots, generated backups, and the main workbook are intentionally not part of this repo.

## What The Updater Does

1. Finds the newest `Sorensen FW...` rolling workbook by the `Date Range` in cell `A1`, preferring `..\02 Weekly Rolling Source Workbooks\`.
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

1. Upload or save the newest GSS source workbook into `..\02 Weekly Rolling Source Workbooks\`.
2. Close the main workbook if it is open in Excel.
3. Double-click `..\Run GSS Update After Upload.cmd`.
4. Review the copy-test summary.
5. Type `APPLY` only when the copy-test looks correct.

The updater searches the organized Dropbox subfolders, so the workbook and source files do not need to stay in the folder root.

## Safety

- The operator launcher runs without `-Apply` first, creating a copy-only test workbook in `..\_automation_runs\test-output`.
- Live apply runs only after the operator types `APPLY`.
- Running with `-Apply` updates the main workbook and first saves a timestamped backup in `..\_automation_runs\backups`.
- If the newest week and matching prior-year week are already present in `Raw_Data`, the updater skips the write.
- Each run prints a short summary and writes a JSON log to `..\_automation_runs\logs`.

## Setup Checks

- Run `scripts\Install-GSS-OperatorLauncher.ps1` to refresh the Dropbox-facing launcher from the tracked template.
- Run `scripts\Get-GSS-SetupStatus.ps1` to check the launcher, repo path, and scheduled task wiring.
- Run `scripts\Test-GSS-PowerShellSyntax.ps1` and `scripts\Test-GSS-Logic.ps1` before committing script changes.

## Requirements

- Windows
- Microsoft Excel desktop app
- PowerShell

The disabled Windows scheduled task is named `GSS Survey Main Workbook Weekly Update`. Manual execution is the intended workflow. The scheduled-task launcher runs live apply without an interactive prompt, so leave the task disabled unless scheduled live updates are explicitly requested.
