# Contributor guidance

Keep this repository program-only. Never commit operational workbooks, guest
details, PDFs, reports, screenshots, credentials, logs, backups, or real-business
test fixtures. Synthetic fixture builders and expected results belong in tests.

Use the existing PowerShell/desktop Excel pipeline. Separate reporting business
rules from file publication and keep changes focused. Ordinary reporting must
not require Git, clean checkout status, original-machine eligibility, or retained
historical test workbooks. Release validation belongs in development/install.

Preserve report-period/entity/value checks, explicit correction decisions,
within-transaction hashes, write exclusion, verified local recovery copies, and
conflict-aware rollback. Independent backup status is separate from report
success. Do not confuse a DriveFS copy with verified remote upload.

Run tests against temporary synthetic fixtures or isolated copies. Never use a
production apply to validate a code change. Keep the live source data and
published reports untouched unless the user explicitly requests a report run.

Keep `templates/Run GSS Update After Upload.cmd` canonical and refresh operator
files through `scripts/Install-GSS-OperatorLauncher.ps1`. Preserve the GSS data
root and `_automation_runs/email_outbox` compatibility with the external consumer.
Do not move immutable package evidence without a validated migration.

Only one computer may write at a time. Sequential use after synchronization is
supported; Dropbox is not a distributed lock. Keep schedules disabled unless
the user explicitly requests activation. Do not add mail scanning or sending.

Run the validation commands in [docs/development.md](docs/development.md). Update
tests, CI paths, operator guidance, and release metadata with behavioral changes.
Do not freeze obsolete architecture with source-text assertions when behavioral
checks can verify the required protection. Preserve historical receipts and
reversible cleanup evidence.

For Ubuntu x64 Cloud development, run `bash ./tools/Setup-GSS-Cloud.sh` during
environment setup, then `pwsh -NoProfile -File ./tools/Test-GSS-Cloud.ps1` from
the repository root. The latter deliberately runs the portable synthetic
subset; it does not replace Windows transaction/recovery tests or desktop Excel.
Cloud tasks must not run the operator launcher, configure live backup paths,
mount Dropbox, activate schedules, or access guest data or mail credentials.
No long-running application service or startup server is required.

Before claiming Codex Cloud is ready, record the actual environment identifier,
privacy/publication state, tested commit, runtime versions, commands, and results
from a fresh task started from the published environment. Local or GitHub Linux
CI results alone are not that evidence. Current Cloud environments and legacy
Code Review/integration environments are separate; inspect each when relevant.
