# GSS Survey Workbook Automation

This private repository contains the program that updates the GSS survey score-trends workbook in Dropbox.

## Weekly Update

Non-technical users should work from the parent `GSS Surveys` folder, not from this repository.

1. Save the newest `Sorensen FW...` workbook in `02 Weekly Rolling Source Workbooks`.
2. Close the main GSS workbook if it is open in Excel.
3. Confirm Google Drive for desktop is running and this workstation's commissioned private GSS backup folder is available.
4. Double-click `Run GSS Update After Upload.cmd` in the `GSS Surveys` folder.
5. Wait for the copy-test and review to finish.
6. Type `APPLY` only when the tested run fingerprint is shown and you want to update the live workbook.

The launcher can be started on any properly configured Windows workstation with the synchronized GSS folder, Microsoft Excel desktop, and the commissioned private Google Drive backup available. Each invocation stages and tests the workbook once, then binds `APPLY` to that exact tested fingerprint. Keep the copy-test and `APPLY` in the same launcher session. The workflow must prepare and hash-verify the Drive backup before promoting the staged workbook.

Run GSS on only one computer at a time. Do not open a second GSS update while another workstation's launcher is still running; Dropbox synchronization is not a cross-workstation lock.

## What the Final Message Means

- `NO LIVE CHANGES MADE`: the process stopped safely without changing the live workbook.
- `UPDATE COMPLETE`: the live workbook was updated, the final review passed, and the Drive snapshot was committed.
- `BACKUP FINALIZATION PENDING`: the workbook is valid and stays in place, but no READY email package is published until the idempotent Drive-finalization retry succeeds.
- `ATTENTION NEEDED`: stop and review the message shown on screen before using the results.

Finished comparison PDFs are saved in `04 Email Comparison PDFs`.
They are named for the report week, not the file-modification date. After a successful live apply and committed Drive snapshot, the launcher publishes a portable package for the AOL draft writer. A copy-test never publishes an email package.

Every READY package includes three reviewed attachments, including an unchanged raw guest-detail workbook that contains personal data. READY means the package passed the workflow controls; it does not mean anonymous or PII-free. Treat the package as **CONTAINS PERSONAL DATA — RESTRICTED**. Human recipient/content review is required, and automatic sending is permanently disabled.

Portable-package path screening treats a complete scheme-qualified HTTP(S) token, including RFC URL punctuation and path-shaped query or fragment values, as URL content. It preserves a possible local-path suffix for rejection only after the reviewed prose delimiters comma, semicolon, `)`, `]`, or `}`, or at an unambiguous direct backslash boundary; opening or URL-valid punctuation is not broadened into a path boundary.

## Operator Folder Map

- `01 Main Workbook`: the active `GSS Score Trends - Main.xlsx` workbook, with older versions in its archive subfolder.
- `02 Weekly Rolling Source Workbooks`: weekly `Sorensen FW...` source files used by the updater.
- `03 Uploaded Survey Workbooks`: current and archived guest-detail exports used only for the post-apply email package; missing or corrupt detail data does not block an otherwise safe workbook update, but it does block drafting.
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
- `scripts\Import-GSS-HistoricalGuestDetail.ps1`: one-time, manifest-bound historical guest-detail recovery.
- `scripts\Gss-EmailPackage.ps1`: restricted email-package publisher with risk-reduced analysis text and an explicitly labeled raw personal-data attachment.
- `scripts\Gss-Common.ps1`: shared conversion and path helpers.
- `scripts\Invoke-GSS-DriveBackup.ps1`: private Drive backup, finalization, retention reporting, capacity projection, and verify-only restore entry point.
- `scripts\Invoke-GSS-ReleaseBackup.ps1`: closed three-file Drive backup for an exact tagged release, manifest, and local Excel receipt.
- `scripts\Gss-DriveBackup.ps1`: hash-verified DriveFS backup functions.
- `scripts\Resume-GSS-PendingFinalize.ps1`: idempotent backup/package continuation that never reapplies the workbook.
- `scripts\Invoke-GSS-QuarterlyRestoreDrill.ps1`: isolated Drive restore plus desktop Excel validation.
- `scripts\Test-GSS-ReleaseIntegrity.ps1`: tagged-release, clean-tree, manifest, and executable-file guard.
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

- Copy-tests and transaction staging go to `..\_automation_runs\test-output`.
- Live backups go to `..\_automation_runs\backups`.
- JSON logs go to `..\_automation_runs\logs`.
- QA packages go to `..\_automation_runs\qa\run_review_YYYYMMDD_HHMMSS`.
- Ready email packages go to `..\_automation_runs\email_outbox\<package-id>` only after a successful live apply, final workbook validation, and committed Drive snapshot.
- The post-run result reports `WorkbookStatus`, `AnalysisStatus`, `EmailReadiness`, `CommenterLens.Status`, and `Modeling.Status` separately. Only workbook blockers prevent apply; detail-data or attachment-evidence blockers prevent drafting. The structural `PopulationRawDataUnavailable` modeling status never blocks the workbook, Drive backup, or package.
- A blocked copy-test review prevents live apply.
- Exact entity-set checks block partial, duplicate, or unexpected current/prior-year week rows before mutation.
- Live promotion rechecks the starting workbook and source hashes, uses a same-volume atomic replace, and rolls back only when the live file still matches this run's promoted hash.
- Logs, QA evidence, manifests, and receipts use atomic writes.

### Independent Google Drive Backup

The supported destination is the owner-only `GSS Survey Backups` folder in `bryan.pettit80@gmail.com` My Drive. Runtime access uses the exact Drive-for-desktop path stored in `%LOCALAPPDATA%\GSSSurveyWorkbookAutomation\settings.json`; the workflow never falls back to Dropbox or another folder.

The curated recovery set includes operator guides, folders `01` through `06`, raw guest-detail workbooks, QA/log/state evidence, finalized READY packages, current transaction artifacts, and exact program-release archives. It excludes test output, old local runtime backups, quarantine, staging/temp files, and the repository working tree/`.git`. Because the recovery set contains raw guest data, the Drive root is classified **CONTAINS PERSONAL DATA — RESTRICTED**.

`Prepare` writes and hash-verifies a `.partial-<run-id>` snapshot and writes `prepared-manifest.json` last. `Finalize` refreshes the post-apply artifacts, promotes the snapshot, and atomically writes `backup-manifest.json` plus `commit-receipt.json`. Runtime verification is recorded as `drivefs_hash_verified`; commissioning and restore drills add Google Drive metadata readback evidence.

Historical recovery uses a separate `RecoveryOnly` snapshot purpose. Its inventory contains only the explicitly reviewed recovered archive files, first-seen ledger, QA, manifest, and transaction receipt. It does not sweep the live workbook or unrelated operational files into the recovery snapshot.

Tagged program delivery uses a separate `ReleaseOnly` snapshot purpose. It accepts exactly three hash-bound, non-row-level artifacts: the tagged ZIP archive, `release-manifest.json`, and the passed local Excel validation receipt. It validates the clean tag, archive contents, manifest controls, receipt, copy-test workbook hash, and source-run fingerprint before preparing or committing the Drive snapshot. It never applies or replaces the live workbook.

Retention is report-only: keep the newest completed snapshot in each of the latest 13 ISO weeks plus the latest snapshot in each of the prior 12 calendar months. The automation reports older candidates but never deletes them. Quarterly restore drills extract to `%LOCALAPPDATA%` and verify hashes without overwriting the live Dropbox tree.

Useful commands:

```powershell
.\scripts\Invoke-GSS-DriveBackup.ps1 -Operation Inventory -RunSummaryPath <run-summary.json> -OutputObject
.\scripts\Invoke-GSS-DriveBackup.ps1 -Operation CapacityProjection -RunSummaryPath <run-summary.json> -OutputObject
.\scripts\Invoke-GSS-DriveBackup.ps1 -Operation RetentionReport -OutputObject
.\scripts\Invoke-GSS-DriveBackup.ps1 -Operation VerifyRestore -RunId <committed-run-id> -OutputObject
.\scripts\Resume-GSS-PendingFinalize.ps1 -LiveRunLogPath <exact-apply-log.json>
.\scripts\Invoke-GSS-QuarterlyRestoreDrill.ps1 -RunId <committed-run-id>
.\scripts\Invoke-GSS-ReleaseBackup.ps1 -Operation Inventory -OutputObject
.\scripts\Invoke-GSS-ReleaseBackup.ps1 -Operation Prepare -OutputObject
.\scripts\Invoke-GSS-ReleaseBackup.ps1 -Operation Finalize -OutputObject
.\scripts\Invoke-GSS-ReleaseBackup.ps1 -Operation VerifyRestore -OutputObject
```

The resume command never reapplies the workbook. For a run created by the current release, it revalidates the exact run and live hashes, retries the Drive commit idempotently, and publishes the READY package only after the backup is committed. Cross-release recovery is allowed only for a code-reviewed release pair (currently `v1.1.8` to `v1.1.9`), not merely because two versions are adjacent. That package-only path requires an unpublished package-blocked receipt, the annotated historical tag, a hash-verified committed workbook snapshot, an exact match for the run log, every detail workbook, and the feedback ledger, plus separately committed `RecoveryOnly` evidence for historical-detail metadata. Undeclared release pairs or changed inputs are rejected. Every publisher enters the shared workstation transaction mutex, rechecks the exact source evidence during capture and immediately before promotion, and restores the feedback ledger to its exact prior bytes if a late publication step fails. The quarterly drill adds desktop Excel validation to the hash-verified restore and writes a combined receipt inside the isolated restore folder.

### Historical Guest-Detail Recovery

Historical imports require a closed `gss-historical-recovery/v1` manifest and exact staged XLSX hashes. Report-week assignment comes from chronological evidence and the paired rolling workbook's terminal fiscal week; a repeated email subject is not authoritative.

`Plan` is read-only. `Apply` takes the global GSS transaction mutex, validates file bytes, row/date bounds, response identities, and destination paths, then baselines recovered response hashes in the first-seen ledger before any historical XLSX becomes visible. A manifest-bound receipt supports idempotent resume and conservative rollback. The live score-trends workbook, current email package, scheduled task, and automatic-sending setting are outside the recovery transaction.

When commenter analytics sees XLSX files in `Recovered Historical Detail`, it requires the committed recovery receipt and sibling manifest, reconstructs every destination from the current GSS root, and rechecks the exact file set, sizes, SHA-256 values, row counts, and response-set hashes before and after parsing. A missing, added, substituted, or uncommitted recovered file produces a commenter-lens data-quality review; it does not block or modify the live workbook. Current and ordinary archived uploads outside that recovered subtree retain the normal weekly workflow.

```powershell
.\scripts\Import-GSS-HistoricalGuestDetail.ps1 -Operation Plan -FolderPath <gss-root> -ManifestPath <manifest.json> -SourcePath <staged-xlsx[]> -OutputObject
.\scripts\Import-GSS-HistoricalGuestDetail.ps1 -Operation Apply -FolderPath <gss-root> -ManifestPath <manifest.json> -SourcePath <staged-xlsx[]> -OutputObject
```

`source_report_week` identifies the file/report assignment. Per-response fiscal week remains based on the local visit date, so late or boundary responses can fall in a different response week without changing the source file's assignment.

### Analytics Source Contract

Analysis policy v4 reflects the source design actually delivered by the survey system:

- The weekly rolling workbook is the authoritative population aggregate for all submitted surveys.
- Row-level ratings exist only for surveys that include comments.
- Row-level ratings for non-comment surveys are not provided.
- Comment-bearing surveys are a structurally selected subset and are not assumed to represent the population.

The review therefore keeps population results descriptive: current 13-week score and response count, movement versus the adjacent rolling window, prior-year movement, and benchmark gaps. The existing response-count tier is exposed as `ResponseVolumeTier`; it is not a statistical confidence interval. Adjacent 13-week windows overlap for 12 weeks, and restaurant results are nested in the broader benchmarks, so those comparisons do not support significance claims.

The separate commenter lens is always labeled **Among guests who provided comments**. For each restaurant and supported score it reports the separately labeled population aggregate, unique comment-survey count, and commenter event count and rate. It reports cross-source comment coverage and commenter-versus-population gaps only when the commenter rows and population aggregate are explicitly verified as the same reporting partition. When that same alignment is verified, the population aggregate reconstructs to an exact event count, and commenter scores are complete, it may also derive the non-comment survey count and event rate and report a descriptive commenter-versus-non-commenter selection gap.

The reporting date must be a Sunday, making the commenter window an explicit inclusive Monday-through-Sunday 13-week period. Stable vendor response IDs are not consistently available across the exports, so deduplication is deterministic and content-based; the reviewer output discloses the resulting low residual risk of a false or missed match.

Selection diagnostics are suppressed for the affected metric when report-week alignment, restaurant mapping, deduplication, score completeness, population numerator reconstruction, or `commenter count <= population count` does not validate. `Recommend` detractor results remain commenter-only unless an exact matching population aggregate becomes available. Themes are shares among comment-bearing surveys, use unique survey denominators, and may overlap.

The commenter lens writes aggregate-only `commenter_lens.json` and `commenter_lens.csv` files into the restricted QA package and is also summarized in `review.json`, `review.md`, the restricted analysis document, and the email preview. The restricted email-package manifest hash-binds those reviewer files, the analysis, both previews, and the classification notice as six non-attachment artifacts. They do not add a fourth email attachment: the package still contains exactly three reviewed attachments. Raw rows, comments, contacts, and individual predictions are never written to an analytics artifact.

Population driver modeling is unsupported under this source contract and always reports the nonblocking status `PopulationRawDataUnavailable` with reason `non_comment_survey_rows_not_provided`. The program does not produce logistic driver models, causal or significance claims, population prevalence estimates from commenter rows, or individual predictions. Forecasting remains deferred and may use only sufficiently long aggregate population history in a separately reviewed future change.

### Validation

Run these before committing script changes:

```powershell
.\scripts\Test-GSS-PowerShellSyntax.ps1
.\scripts\Test-GSS-PSScriptAnalyzer.ps1
.\scripts\Test-GSS-Pester.ps1
.\scripts\Test-GSS-Logic.ps1
.\scripts\Test-GSS-Analytics.ps1
.\scripts\Test-GSS-EmailPackage.ps1
.\scripts\Test-GSS-DriveBackup.ps1
.\scripts\Test-GSS-HistoricalRecovery.ps1
.\scripts\Test-GSS-ReleaseIntegrity.ps1 -SkipTagCheck
```

PSScriptAnalyzer `1.25.0` and Pester `5.8.0` are pinned in CI. The analyzer baseline records reviewed legacy warnings and fails on any new warning fingerprint; Pester covers new modules while the established regression harnesses remain in place.

For local Excel integration verification, run `Test-GSS-WorkbookIntegration.ps1` against the staged copy-test workbook and write its machine-readable receipt to `..\_automation_runs\state\release\local-excel-validation-receipt.json`. Production validates that the receipt is a passed copy-only test bound to the exact HEAD/tag, workbook hash, Excel version, certifying workstation, and run fingerprint. The validated receipt is portable across configured workstations; its recorded source hostname remains hash-bound audit evidence but is not an execution allowlist. This check is intentionally not part of GitHub Actions because the hosted runner does not provide desktop Excel.

### Release Contract

Production execution requires a clean, exact tagged release at `HEAD`, a matching `release\release-manifest.json`, and no untracked executable files. Generate the reviewed manifest only after code and policy changes are complete:

```powershell
.\scripts\New-GSS-ReleaseManifest.ps1 -Version <major.minor.patch>
.\scripts\Test-GSS-ReleaseIntegrity.ps1 -SkipTagCheck
```

After the manifest commit is merged and CI is green, create the annotated release tag on that exact `main` commit, run the local Excel validation, archive the exact tag, and back up the release manifest, archive, and validation receipt with the closed `ReleaseOnly` Drive workflow. The production launcher performs the full tag and receipt checks and refuses mutable or unapproved code.

Requirements are Windows, Microsoft Excel desktop, and PowerShell.

The Windows task `GSS Survey Main Workbook Weekly Update` is intentionally disabled. Manual execution is the supported workflow unless scheduled live updates are explicitly approved.
