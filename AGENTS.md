# Codex Notes

This repository contains only the GSS workbook automation program. Do not commit survey source workbooks, PDFs, images, generated workbooks, backups, logs, or customer/business data.

The live Dropbox data folder is the parent of this repository:

`C:\Users\bryan\Dropbox\Automations\GSS Surveys`

The operator entry point is outside the repo:

`Run GSS Update After Upload.cmd`

Use `scripts\Update-GSS-MainWorkbook.ps1` for code changes. Run it without `-Apply` first for a copy-only test. Live runs require Microsoft Excel on Windows and write backups/logs to `..\_automation_runs`.

Do not bind release eligibility to a named workstation. A passed local Excel release receipt is portable when its exact tag, commit, workbook hash, source log, and run fingerprint all validate. Keep each live copy-test and `APPLY` transaction in the same launcher session and retain the transaction hostname as audit evidence.

Keep `templates\Run GSS Update After Upload.cmd` as the canonical Dropbox-facing launcher and refresh the parent launcher with `scripts\Install-GSS-OperatorLauncher.ps1`. Leave the scheduled task disabled unless the user explicitly requests scheduled live updates.

`scripts\Analyze-GSS-Run.ps1` is the deterministic post-run QA and insight analyzer. Generated review files belong under `..\_automation_runs\qa` and must not be committed.
