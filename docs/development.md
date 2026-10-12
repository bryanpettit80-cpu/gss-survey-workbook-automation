# Development and release

## Repository boundaries

This repository contains application source and synthetic tests. Business
workbooks, guest details, PDFs, logs, credentials, and local backups stay outside
Git. Generate synthetic Excel fixtures in a temporary directory; commit their
builder and expected business values rather than private workbook samples.

| Directory | Purpose |
| --- | --- |
| `scripts/` | Runtime, setup, analysis, Excel validation, and recovery |
| `tests/` | Behavioral tests and synthetic fixture definitions |
| `tools/` | Repository validation runners |
| `config/` | Reporting policy and static-analysis configuration |
| `templates/` | Installed launcher and operator guides |
| `release/` | Version and installation/release metadata |
| `docs/` | Operations, maintenance, and future integration specifications |

Prefer small changes to the existing PowerShell and Excel implementation. Keep
business parsing and comparisons separate from file publication. Share pure
helpers where the behavior is identical; do not create a framework merely to
move functions into more files.

## Validation

Use Windows PowerShell 5.1 for compatibility with the deployed Excel application.
Development checks use Pester 5.8.0 and PSScriptAnalyzer 1.25.0. Runtime users do
not need these modules or Git.

```powershell
.\tools\Test-GSS-PowerShellSyntax.ps1
.\tools\Test-GSS-PSScriptAnalyzer.ps1
.\tools\Test-GSS-Pester.ps1
.\tests\Test-GSS-Logic.ps1
.\tests\Test-GSS-Analytics.ps1
.\tests\Test-GSS-EmailPackage.ps1
.\tests\Test-GSS-DriveBackup.ps1
.\tests\Test-GSS-HistoricalRecovery.ps1
```

When the separate AOL consumer source is available, run the package suite with
`-AolSourceRoot 'D:\Projects\aol-to-gmail-sync\src'` to exercise its parser against
synthetic packages. This reads the consumer code; it does not create drafts or
contact mail providers.

The Excel integration validator belongs in `scripts/` because restore tooling
also uses it. Run it only against an isolated test/restore workbook. GitHub-hosted
validation does not establish desktop Excel rendering or a second-PC test.

Regression expectations must independently specify dates, entities, counts,
scores, changes, and report contents. For example, an overall score of 70 versus
75 and 80 has changes of -5 and -10 percentage points. Dissatisfaction of 5 versus
8 and 9 represents improvements of +3 and +4 points. Do not average restaurant
percentages to recreate the vendor's population totals.
See [regression scenarios](regression-coverage.md) for representative fixtures,
expected outcomes, and the isolated desktop Excel command.

Test negative/invalid counts, out-of-range percentages, wrong rolling periods,
missing and duplicate entities, harmless renames/saves, conflicting corrections,
and missed-week catch-up. Test interruptions around workbook/PDF publication,
local-backup failure, independent-backup retry, and recovery under another root.
Review PDF labels, visible values, page breaks, and clipping. File-byte equality
is useful for verifying copied artifacts, not for proving business equivalence.

The established missing-history interval between July 6 and October 5, 2025 is
retained as a compatibility rule. Tests must not fabricate weekly observations
in that interval. A real second-computer run remains an additional acceptance
check; changing a test root on the same computer does not establish it.

## Installation and release controls

Validate source and packaging during release preparation. Ordinary reporting
must not require a clean Git checkout, an exact Git tag, or an old test workbook
and log to remain present. Record the installed version in run evidence.

Keep release changes reviewable and preserve the previous installed package for
rollback. Run non-Excel CI and isolated desktop Excel checks before deploying.
Publish a versioned program-only archive and concise release notes describing
behavior, validation, and any unverified platform coverage. Keep installation
instructions aligned with the published artifact.

Existing legacy backup verification helpers may still need their historical
release evidence when inspecting old snapshots. Preserve those original records;
do not rewrite them to manufacture compatibility or current verification.
`Invoke-GSS-ReleaseBackup.ps1` and the commissioned backup format are historical
tools for the pre-1.2 release contract. They do not create a new 1.2 release backup.
New releases use the program-only archive; normal report recovery uses the
versioned run snapshots described in [operations](operations.md).
Verify a new snapshot from the program folder:

```powershell
.\scripts\Invoke-GSS-DriveBackup.ps1 -Operation VerifyRun -RunSnapshotPath '<snapshot-folder>\run-snapshot.json'
```

The manifest records each artifact's original relative path. Restoring an entire data folder from
that format remains a reviewed manual operation; the quarterly full-folder
restore drill applies to the legacy format. `Resume` handles current transaction
rollback and independent backup retry.

## Local cleanup and rollback

Organize historical analysis under `05 Reference Materials/Past Analysis` and old
program archives under `_automation_runs/archive/program-releases`. Record every
move's original path, destination, size, and SHA-256; reverse that manifest if a
path dependency is found. Include program archives in the independent backup
scope. Do not delete source-hold, quarantine, historical recovery, or transaction
artifacts based only on age.

The existing AOL consumer binds source paths to the GSS data root. Keep the data
root, published packages, and `_automation_runs/email_outbox` stable during this
refactor. A future path/schema migration must be tested with that consumer.

See [operations](operations.md) for operator actions and
[future integration](automation-interface.md) for the coordinator boundary.
