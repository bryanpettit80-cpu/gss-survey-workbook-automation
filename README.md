# GSS Survey Workbook Automation

This private repository contains the program that updates the GSS survey score-trends workbook in Dropbox.

## Weekly Update

Non-technical users should work from the parent `GSS Surveys` folder, not from this repository.

1. Save the newest `Sorensen FW...` workbook in `02 Weekly Rolling Source Workbooks`.
2. Close the main GSS workbook if it is open in Excel.
3. Double-click `Run GSS Update After Upload.cmd` in the `GSS Surveys` folder.
4. Wait for the copy-test and review to finish.
5. Type `APPLY` only when the copy-test passes and you want to update the live workbook.

The launcher always tests a copy first. It also creates a backup before changing the live workbook.

## What the Final Message Means

- `NO LIVE CHANGES MADE`: the process stopped safely without changing the live workbook.
- `UPDATE COMPLETE`: the live workbook was updated and the final review passed.
- `ATTENTION NEEDED`: stop and review the message shown on screen before using the results.

Finished comparison PDFs are saved in `04 Email Comparison PDFs`.

## Operator Folder Map

- `01 Main Workbook`: the active `GSS Score Trends - Main.xlsx` workbook, with older versions in its archive subfolder.
- `02 Weekly Rolling Source Workbooks`: weekly `Sorensen FW...` source files used by the updater.
- `03 Uploaded Survey Workbooks`: survey-detail files kept for reference; not used by the weekly updater.
- `04 Email Comparison PDFs`: finished comparison PDFs.
- `05 Reference Materials`: training and reference documents.
- `06 Exports and Images`: screenshots and image exports.
- `_automation_runs`: system-created tests, backups, logs, and QA reviews.
- `GSS Survey Workbook Automation`: this program repository.

## For Maintainers

The repository is intentionally program-only. Do not commit survey workbooks, PDFs, screenshots, generated workbooks, backups, logs, or business data.

### Main Scripts

- `scripts\Invoke-GSS-SafeWorkbookUpdate.ps1`: canonical safe workflow used by the operator launcher.
- `scripts\Update-GSS-MainWorkbook.ps1`: Excel updater.
- `scripts\Analyze-GSS-Run.ps1`: deterministic post-run QA and insight review.
- `scripts\Gss-Common.ps1`: shared conversion and path helpers.
- `scripts\Install-GSS-OperatorLauncher.ps1`: refreshes the operator launcher and plain-language guide files.
- `scripts\Get-GSS-SetupStatus.ps1`: checks operator files and scheduled-task state.

### What the Updater Does

1. Finds the newest `Sorensen FW...` rolling workbook by the `Date Range` in cell `A1`.
2. Finds the matching prior-year workbook using newest week minus 364 days.
3. Reads All Franchisees total, Sorensen total, 9354 Richmond, and 9355 Virginia Beach.
4. Adds missing rows to `Raw_Data` and refreshes its formulas.
5. Refreshes the visible `QA Checks` sheet, report status banners, and print layouts.
6. Refreshes the `Email Comparison` worksheet and exports its PDF.
7. Reapplies passwordless worksheet and workbook-structure protection, leaving only the two dashboard selectors editable.
8. Calculates and saves the workbook through Microsoft Excel.

### Safety and Output

- Copy-tests go to `..\_automation_runs\test-output`.
- Live backups go to `..\_automation_runs\backups`.
- JSON logs go to `..\_automation_runs\logs`.
- QA packages go to `..\_automation_runs\qa\run_review_YYYYMMDD_HHMMSS`.
- A blocked copy-test review prevents live apply.
- Existing current/prior-year rows are skipped rather than duplicated.

### Validation

Run these before committing script changes:

```powershell
.\scripts\Test-GSS-PowerShellSyntax.ps1
.\scripts\Test-GSS-Logic.ps1
.\scripts\Test-GSS-Analytics.ps1
```

For local Excel integration verification, run `Test-GSS-WorkbookIntegration.ps1` against a copy-test workbook. This check is intentionally not part of GitHub Actions because the hosted runner does not provide desktop Excel.

Requirements are Windows, Microsoft Excel desktop, and PowerShell.

The Windows task `GSS Survey Main Workbook Weekly Update` is intentionally disabled. Manual execution is the supported workflow unless scheduled live updates are explicitly approved.
