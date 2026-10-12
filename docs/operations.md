# Setup, weekly reporting, and recovery

## Set up a Windows computer

1. Install Microsoft Excel desktop and use Windows PowerShell 5.1 or later.
2. Download the GSS operational folders. Ensure the workbooks are available
   offline, not only cloud placeholders. Sign in to Dropbox/Drive through their
   normal desktop applications when those services are used.
3. Place the downloaded program in `GSS Survey Workbook Automation` inside the
   data folder. The launcher finds this location relative to itself; it does not
   require a particular username or drive letter.
4. From the program folder, install the operator files and check setup:

   ```powershell
   .\scripts\Test-GSS-ReleaseIntegrity.ps1 -InstallationOnly
   .\scripts\Install-GSS-OperatorLauncher.ps1 -Folder 'D:\GSS Surveys'
   .\scripts\Get-GSS-SetupStatus.ps1 -Folder 'D:\GSS Surveys'
   ```

5. Optionally configure an existing private independent backup folder:

   ```powershell
   .\scripts\Set-GSS-Setup.ps1 -BackupFolder 'G:\My Drive\GSS Backups'
   ```

   This stores the destination in this Windows user's local settings. Another
   computer can use a different drive letter. An existing installation's Drive
   destination is recognized without rebuilding its commissioning history.
   `-ClearBackup` removes the optional destination setting; local recovery copies
   remain mandatory, and results show that independent backup is pending.
   Backup is optional for generating the local report. The normal launcher still
   waits for report-backup completion before publishing a draft package.

Git, Pester, and historical Excel test receipts are not runtime prerequisites.
The one-time installation check verifies the downloaded program files without
Git; ordinary reporting does not repeat it.
Excel registration is a setup check, not proof that an actual report has passed
recalculation and rendering.

## Run the weekly update

Save vendor population reports in `02 Weekly Rolling Source Workbooks` and detail
exports in `03 Uploaded Survey Workbooks`. Close the main workbook in Excel and
double-click **Run GSS Update After Upload.cmd**. Launching authorizes a normal
validated update; there is no additional `APPLY` prompt.

Read the final message and `Last GSS Result.txt`. The program checks a staged
copy, preserves the previous good report, and verifies publication. Missing or
invalid aggregate inputs leave the existing report unchanged. A missing or
unassigned detail export can leave draft preparation waiting while a valid
aggregate report completes.

Every zero-response row stops publication when it would remain in the workbook,
including historical rows and rows whose vendor scores are numeric zeros. Zero
responses cannot establish a measured rate. The original vendor values remain
unchanged for content identification, but the existing workbook model cannot
present these rows safely as unscored. The message identifies the entity and
reporting week, and the last good report is preserved. A workbook-model change is
required before these results can be published safely. A zero-percent score with
a positive response count remains valid.

The default main workbook is `GSS Score Trends - Main.xlsx`. If it is renamed,
the program can recognize a unique score-trends workbook directly inside
`01 Main Workbook`. It never silently substitutes an archived workbook. With
multiple candidates, select the intended filename explicitly:

```powershell
.\scripts\Invoke-GSS-SafeWorkbookUpdate.ps1 -Folder 'D:\GSS Surveys' -MainWorkbookName 'Guest Survey Trends.xlsx'
```

## Duplicates, corrections, and detail assignment

Identical business content is a duplicate even when the file has another name or
different Excel serialization. Keep original sources for traceability, but their
exact filenames are not permanent reporting identity.

If same-period sources contain different values, the launcher identifies the
conflict and asks which revision to use. Canceling leaves the report unchanged.
For an explicit maintenance invocation, use `-ReportingWeekEnding yyyy-MM-dd`
and `-CorrectionSourcePath` with the selected source file. The date must be a
Sunday. Selecting a correction replaces that period's business rows and refreshes
the affected comparisons; it must not append another copy of the same entities.
After successful validation, the accepted content choice is retained. Renaming
that file or retaining a known older version does not require choosing again;
an unfamiliar revision requires a new decision.

Guest-detail filenames often have no reporting date. The program uses explicit
assignments and existing package evidence rather than guessing from the latest
visit date. A first assignment may require a file/week decision. Corrected detail
must identify which previous export it supersedes so both revisions are not
counted as independent surveys.

Exports without reliable response identifiers can remain ambiguous and require
review. Existing packages remain immutable; a legacy package whose original
external files changed may need a one-time revision review before a new package
can be prepared.

## Backups and interrupted work

There are two separate protections:

- A local previous-good recovery copy is verified before report replacement.
- An independent copy is attempted separately and may remain pending without
  invalidating a completed local report.

If the local recovery snapshot itself cannot be completed, the program retains
`LocalSnapshotPending` and requires recovery before the next report update. An
unavailable independent backup destination alone does not hold that write lock.

Backup verification through Drive for desktop means the copied bytes were read
and checked locally. It does not assert that cloud upload is complete. The backup
destination contains guest data and should have the intended restricted access.

The report snapshot captures files and assignments available when that report's
snapshot was created. It remains immutable. Detail arriving or corrected later
is frozen in the new draft package locally; its status is `LocalOnly`, not a claim
of independent backup. Report-backup and draft-evidence status are separate.

Use setup/status to see pending run identifiers. Retry an existing run using the
identifier shown in its result:

```powershell
.\scripts\Resume-GSS-PendingFinalize.ps1 -Folder 'D:\GSS Surveys' -RunId '<run-id>'
```

Resume never reapplies a committed report. It retries its snapshot/backup and
eligible package work. An interrupted local publication is reconciled from its
recorded evidence. Restoration must refuse to overwrite a file changed by an
unrelated later action.

For a deliberate restoration of the previous report from a selected transaction:

```powershell
.\scripts\Resume-GSS-PendingFinalize.ps1 -Folder 'D:\GSS Surveys' -RunId '<run-id>' -RestorePreviousReport
```

Preserve the identified transaction files if recovery reports missing or conflicting
evidence. Do not delete the active-transaction marker to force another write.
Original legacy receipts remain unchanged; new recovery records describe any
continuation. Old backup inspection tools remain available for their historical
snapshot format.

## Use another computer

Finish the current run, close Excel, and let Dropbox synchronize before switching.
Download the program, required sources, main workbook, and durable internal state
on the other computer. Run setup/status there and configure its backup path if
needed. No original computer name or historical certification workbook is needed
for ordinary reporting.

Run on only one computer at a time. The local lock prevents overlapping processes
on one computer; it cannot make Dropbox a cross-computer lock.

## Internal files and mail boundaries

Leave logs, staging, snapshots, correction assignments, history, and packages
under `_automation_runs`. Past analysis belongs under `05 Reference Materials`;
old program archives belong in the internal archive. Do not delete unresolved
holds or backup evidence based solely on file age.

GSS prepares restricted packages for review. The external consumer owns draft
creation. New package source snapshots preserve the evidence actually published;
historical packages are not rewritten. Management messages remain drafts, and
automatic sending and weekly scheduling remain disabled.

See [future integration](automation-interface.md) for the proposed shared
coordinator and [development](development.md) for validation and release work.
