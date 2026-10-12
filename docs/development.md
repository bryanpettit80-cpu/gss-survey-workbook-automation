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
| `tools/` | Repository validation and release tools |
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

Run commands from the repository root. Install the pinned modules once with
`./tools/Initialize-GSS-Development.ps1`; use `-CheckOnly` to verify them without
installing anything. The initializer uses the current user's module directory
and does not change reporting configuration.

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

CI keeps the complete Windows PowerShell job named `powershell`. Pull requests,
pushes to `main`, version tags, and manual dispatch run validation. Feature
branches are checked through their PR instead of running the same suite for
both the branch push and its PR. The `cloud-development` job separately runs the
portable suite on Ubuntu 24.04. A passing Linux job does not establish Codex
account access or environment publication.

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

Validate every required reporting period against the available sources and
workbook rows. No calendar interval, including July through September 2025, is
exempt from completeness checks. Synthetic fixtures may cover those dates;
missing observations must remain missing rather than being invented.
A real second-computer run remains a separate acceptance check; changing a test
root on the same computer does not establish it. Record that evidence separately
using the [deferred-update criteria](future-updates.md#second-computer-verification).

## Codex Cloud development

Use Cloud for source changes and synthetic checks. Desktop Excel COM automation,
workbook recalculation/PDF rendering, Windows task scheduling, local protected
credentials, and Dropbox/Drive reporting remain Windows operations. The Cloud
suite does not run the operator launcher, copy operational data, create mail
drafts, or send messages. There is no server to start and no business credential
required.

The development bootstrap supports Ubuntu 22.04, 24.04, and 26.04 on x64; CI
exercises 24.04. It installs PowerShell 7.6.6 from Microsoft's official release
with a pinned SHA-256, then Pester 5.8.0 and PSScriptAnalyzer 1.25.0. Use it only in
the disposable development environment; it may upgrade an older PowerShell
package but refuses to downgrade a newer one. Root or passwordless `sudo` is needed when the
pinned package is absent. See the [official Ubuntu installation guide](https://learn.microsoft.com/en-us/powershell/scripting/install/install-ubuntu)
and [pinned release](https://github.com/PowerShell/PowerShell/releases/tag/v7.6.6).

From the repository root:

```bash
bash ./tools/Setup-GSS-Cloud.sh
pwsh -NoProfile -File ./tools/Test-GSS-Cloud.ps1
```

If a suitable PowerShell 7.2+ is already available and a runtime installation is
not wanted, run `pwsh -NoProfile -File ./tools/Initialize-GSS-Development.ps1`,
then the portable test command. Record the actual PowerShell version; this is
not the pinned bootstrap configuration.

The portable suite runs syntax and static analysis over all PowerShell source,
then reporting, accepted source-choice, workbook-discovery, and analytics
checks. On Linux, analytics explicitly excludes its single Windows Dropbox-path assertion; all pure analytics assertions still run. Windows mutex subprocess, transaction, backup, email-package,
historical-recovery, and desktop Excel suites remain outside that Linux scope.
Use the complete Windows commands above before releasing changes that affect
those paths. Do not convert excluded checks into passing results.

Inspect existing environments before creating one. The
[current Cloud environment workflow](https://learn.chatgpt.com/docs/environments/cloud-environments)
uses a prepared, published environment for new tasks; the
[legacy workflow](https://learn.chatgpt.com/docs/environments/cloud-environment)
still serves Code Review and GitHub/Linear integrations. A legacy repository
association does not prove that a current published environment exists.

For this repository, prefer a private, single-repository development environment
unless sharing has been explicitly approved. Keep guest data, Dropbox mounts,
mail credentials, and live backup paths out of its prepared filesystem. Use the
bootstrap as the install command and the portable test command as the readiness
check. No environment variables or secrets are required by the synthetic tests.
Do not copy a developer's local credentials into Cloud.

Setup needs Ubuntu package mirrors, `github.com` and
`release-assets.githubusercontent.com` for PowerShell, and `www.powershellgallery.com`
plus its package download hosts for the pinned modules. Confirm redirects and
actual allowed hosts during setup. Keep network access limited to those needs;
do not assume the package-manager preset includes PowerShell Gallery.

Review setup results before saving and publishing/republishing. Start a fresh
task from that published environment and record:

- Environment name and identifier, repository association, access scope, and
  publication state.
- `pwd`, `git rev-parse HEAD`, and `git status --short`, confirming the intended
  repository/revision without unrelated changes.
- PowerShell/Pester/PSScriptAnalyzer versions and the portable command's actual
  exit status, pass count, and any failures or excluded coverage.
- Whether setup used cached state, what network access was verified, and which
  Windows-only checks still need a local run.

Saving configuration or passing local/Linux CI tests is not a fresh Cloud-task
result. If account access is unavailable, keep environment identity, publication,
and fresh-task validation explicitly unverified.

## Installation and release controls

Validate source and packaging during release preparation. Ordinary reporting
must not require a clean Git checkout, an exact Git tag, or an old test workbook
and log to remain present. Record the installed version in run evidence.

Keep release changes reviewable and preserve the previous installed package for
rollback. Run non-Excel CI and isolated desktop Excel checks before deploying.
Publish a versioned program-only archive and concise release notes describing
behavior, validation, and any unverified platform coverage. Keep installation
instructions aligned with the published artifact.

Review `git status --short` and stage each intended changed or new file by its
exact path (for example, `git add -- docs/development.md`). Do not stage the entire
worktree or whole directories: untracked operational files may be present there.
The manifest generator inventories tracked files, including newly staged paths.
Review the staged file list and diff before regenerating it:

```powershell
git diff --cached --name-only
git diff --cached
.\tools\New-GSS-ReleaseManifest.ps1 -Version '<next version>'
git add release/release-manifest.json
```

Review the staged paths so unrelated local work or business data never enters
the program archive. Validate the regenerated
manifest and record the actual publication and installation outcomes separately.

Existing legacy backup verification helpers may still need their historical
release evidence when inspecting old snapshots. Preserve those original records;
do not rewrite them to manufacture compatibility or current verification.
`Invoke-GSS-ReleaseBackup.ps1` and the commissioned backup format are historical
tools for the pre-1.2 release contract. They do not create a current release backup.
Current releases use the program-only archive; normal report recovery uses the
versioned run snapshots described in [operations](operations.md).
Verify a new snapshot from the program folder:

```powershell
.\scripts\Invoke-GSS-DriveBackup.ps1 -Operation VerifyRun -RunSnapshotPath '<snapshot-folder>\run-snapshot.json'
```

The manifest records each included artifact's original relative path. It covers
a selected run inventory, not every file in the GSS folder. `VerifyRun` checks
the stored bytes; it does not restore files or establish full-folder coverage.
Restoring manifest-listed files to another folder remains a reviewed manual
operation. `Resume-GSS-PendingFinalize.ps1` handles current transaction rollback
and independent backup retry. `Invoke-GSS-QuarterlyRestoreDrill.ps1` exercises the
legacy commissioned snapshot format, not `gss-run-snapshot/v1`.
See [deferred snapshot restoration](future-updates.md#current-format-snapshot-restoration)
for the tooling, coverage checks, and drill required before making broader
recovery claims.

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
[Deferred updates](future-updates.md) retain the unimplemented zero-response and
snapshot-restoration work with acceptance criteria.
