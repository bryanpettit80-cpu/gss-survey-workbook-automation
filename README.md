# GSS Survey Workbook Automation

GSS updates the guest-survey score workbook and comparison PDF for Richmond and
Virginia Beach. It uses the vendor's rolling population reports for restaurant
and benchmark results, and keeps comment-bearing survey analysis separate.

The program stages and checks each update, preserves the previous good report,
and can recover an interrupted workbook or PDF replacement. An independent Drive
backup has its own status, so a Drive outage does not make a valid local report fail.

This repository contains program source and synthetic tests. Survey workbooks,
guest data, reports, credentials, logs, and backups are kept outside Git.
Use the [current published release](https://github.com/bryanpettit80-cpu/gss-survey-workbook-automation/releases/latest).
The release manifest records the installed version; historical receipts retain
the version that created them.

## What it produces

- The updated score-trends Excel workbook, including dashboard, comparisons,
  report definitions, and visible validation results.
- A comparison PDF named for the reporting week.
- A plain-language result showing report, backup, and draft readiness, with
  detailed analysis in the linked review.
- A restricted review package for the existing external draft consumer when all
  package prerequisites are met. GSS does not send email.

## Setup and weekly use

Use Windows, Windows PowerShell 5.1 or later, and Microsoft Excel desktop. Download
the program and the required operational files. Online-only cloud placeholders
must be downloaded before reporting. Git and development test modules are not
ordinary runtime dependencies.

1. Keep the `GSS Survey Workbook Automation` program folder inside the GSS data
   folder, or pass the data folder explicitly when using PowerShell.
2. Install/refresh the operator launcher:

   ```powershell
   .\scripts\Test-GSS-ReleaseIntegrity.ps1 -InstallationOnly
   .\scripts\Install-GSS-OperatorLauncher.ps1 -Folder 'D:\GSS Surveys'
   ```

3. Place rolling reports in `02 Weekly Rolling Source Workbooks` and guest-detail
   exports in `03 Uploaded Survey Workbooks`.
4. Close the main workbook in Excel and double-click
   **Run GSS Update After Upload.cmd**.
5. Read the result. Ordinary complete updates do not require typing `APPLY`.
   Conflicting corrections or an unassigned detail reporting period need a
   specific decision.

Use only one computer at a time. Finish the previous run and Dropbox
synchronization before changing computers. See [setup and recovery](docs/operations.md)
for configurable paths, dependencies, and corrective actions. Reporting is
manually launched; no GSS schedule is enabled.

## How inputs are handled

The reporting period and business values establish source identity. An identical
duplicate, harmless rename, or formatting-only save does not become a new
business revision. Conflicting versions are not selected by modification time.
Corrections must explicitly supersede the selected earlier data.

Reporting checks require the expected four entities: All Franchisees, Sorensen,
9354 Richmond, and 9355 Virginia Beach. Dates, counts, percentages, and comparison
coverage are validated before publication. Missing guest detail can hold the
draft package without preventing a valid aggregate report.

Every zero-response population row currently stops workbook publication, whether
its scores are blank or numeric. Zero responses cannot establish a measured rate,
including a vendor-supplied zero percent. The existing model does not safely
present these rows as unscored, so the previous good report is preserved. The
[deferred updates](docs/future-updates.md) define the work needed to support them.

The main workbook's filename can be configured. Rolling sources are recognized
by their content, not by requiring a permanent original filename. File hashes
still protect the copies used within a transaction and recovery backup.

## Understanding the result

| Result | Meaning |
| --- | --- |
| Report complete | The local workbook/PDF revision passed validation and has a recovery copy |
| Already current | The business inputs match the existing result |
| Waiting or conflict | Required data or a specific correction decision is missing |
| Backup pending | A backup needs retry; a failed local snapshot must be recovered before another update |
| Draft package blocked | The report and draft package have different prerequisites; review the stated missing evidence |
| Recovery required | An interrupted local publication needs reconciliation before another write |

A local DriveFS copy check is not proof of completed remote cloud upload. Draft
packages contain personal guest data and require recipient/content review. The
independent report backup is a snapshot of that run; detail supplied or corrected
later is preserved in the draft package's local evidence and is not claimed as
independently backed up.
Independent backup is optional for completing a local report; the normal launcher
requires its report backup to finish before publishing a draft package.

## Reporting definitions

- Population results come from vendor aggregates. Company totals are not rebuilt
  by averaging restaurant percentages.
- Overall Experience, Service, Culinary, Value, and Pace of Meal are topbox scores:
  the percentage rated 5 (Excellent).
- Dissatisfaction is the percentage rated 1, 2, or 3. A decrease is an improvement.
- Steak Cooked Properly is the percentage of eligible respondents answering Yes.
- Recommend means ratings 9 or 10. When the vendor export supplies only its
  response count, the population percentage remains unavailable.
- Prior-year comparison uses the report Sunday minus 364 days. Adjacent 13-week
  windows overlap by 12 weeks; their changes do not establish significance.
- Commenter results describe guests who provided comments. They do not estimate
  the entire population or establish causal drivers.

## Folder guide

| Folder | Contents |
| --- | --- |
| `01 Main Workbook` | Main workbook and its established archive |
| `02 Weekly Rolling Source Workbooks` | Vendor rolling population sources |
| `03 Uploaded Survey Workbooks` | Current and archived guest detail |
| `04 Email Comparison PDFs` | Finished comparison PDFs |
| `05 Reference Materials` | Reference documents and past analysis |
| `06 Exports and Images` | Supporting exports and images |
| `_automation_runs` | Internal state, recovery copies, logs, tests, and packages |
| `GSS Survey Workbook Automation` | Installed application |

## Development and future automation

See [development and validation](docs/development.md) for repository organization,
tests, release preparation, and rollback. [Deferred updates](docs/future-updates.md)
record the zero-response model, current-format snapshot restoration, and separate
second-computer verification requirements.

Weekly scheduling and shared-coordinator integration are **not enabled**. The
[future interface specification](docs/automation-interface.md) keeps reporting
rules here and attachment delivery, notifications, and management drafts in the
shared coordinator. AOL remains authoritative for read/unread state; Gmail
mirrors AOL. Gift-card Micros intake remains separate.
