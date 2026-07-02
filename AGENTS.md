# Codex Notes

This repository contains only the GSS workbook automation program. Do not commit survey source workbooks, PDFs, images, generated workbooks, backups, logs, or customer/business data.

The live Dropbox data folder is the parent of this repository:

`C:\Users\bryan\Dropbox\Marketing\GSS Surveys`

The operator entry point is outside the repo:

`Run GSS Update After Upload.cmd`

Use `scripts\Update-GSS-MainWorkbook.ps1` for code changes. Run it without `-Apply` first for a copy-only test. Live runs require Microsoft Excel on Windows and write backups/logs to `..\_automation_runs`.
