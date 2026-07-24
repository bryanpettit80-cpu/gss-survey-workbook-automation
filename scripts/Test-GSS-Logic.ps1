[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'

$scriptRoot = Split-Path -Parent $PSCommandPath
. (Join-Path $scriptRoot 'Update-GSS-MainWorkbook.ps1')

function Assert-Equal {
    param(
        [object]$Actual,
        [object]$Expected,
        [string]$Name
    )

    if ($Actual -ne $Expected) {
        throw "Assertion failed for $Name. Expected '$Expected'; actual '$Actual'."
    }
}

function Assert-True {
    param([bool]$Value, [string]$Name)
    if (-not $Value) {
        throw "Assertion failed for $Name."
    }
}

function Assert-ThrowsLike {
    param(
        [scriptblock]$Action,
        [string]$ExpectedMessage,
        [string]$Name
    )

    try {
        & $Action
    }
    catch {
        if ($_.Exception.Message -like $ExpectedMessage) {
            return
        }
        throw "Assertion failed for $Name. Expected error like '$ExpectedMessage'; actual '$($_.Exception.Message)'."
    }
    throw "Assertion failed for $Name. Expected an error, but no error was thrown."
}

function New-TestSource {
    param([string]$Path, [datetime]$WeekEnding, [datetime]$LastWriteTime)

    [pscustomobject]@{
        WeekEnding = $WeekEnding
        File = [pscustomobject]@{
            FullName = $Path
            LastWriteTime = $LastWriteTime
        }
    }
}

Assert-Equal (Normalize-Header 'Pace Of Meal Dissat (lower is better)') 'paceofmealdissatlowerisbetter' 'Normalize-Header punctuation removal'
Assert-Equal (Normalize-Header ' Restaurant ') 'restaurant' 'Normalize-Header whitespace'

$folder = 'C:\GSS Surveys'
Assert-Equal (Test-ExcludedGssPath 'C:\GSS Surveys\GSS Survey Workbook Automation\scripts\Update.ps1' $folder) $true 'Repo folder excluded'
Assert-Equal (Test-ExcludedGssPath 'C:\GSS Surveys\_automation_runs\logs\run.json' $folder) $true 'Automation output excluded'
Assert-Equal (Test-ExcludedGssPath 'C:\GSS Surveys\02 Weekly Rolling Source Workbooks\Sorensen FW01.xlsx' $folder) $false 'Source folder included'

$sources = @(
    (New-TestSource 'older-same-week.xlsx' ([datetime]'2026-06-28') ([datetime]'2026-07-01T08:00:00')),
    (New-TestSource 'newer-same-week.xlsx' ([datetime]'2026-06-28') ([datetime]'2026-07-01T09:00:00')),
    (New-TestSource 'prior-week.xlsx' ([datetime]'2026-06-21') ([datetime]'2026-07-02T09:00:00'))
)

$latest = Select-LatestRollingSource $sources
Assert-Equal $latest.File.FullName 'newer-same-week.xlsx' 'Latest source uses week then last write time'

$priorYear = Select-RollingSourceByWeek @(
    (New-TestSource 'wrong.xlsx' ([datetime]'2025-06-22') ([datetime]'2025-06-22')),
    (New-TestSource 'right.xlsx' ([datetime]'2025-06-29') ([datetime]'2025-06-29'))
) ([datetime]'2025-06-29')
Assert-Equal $priorYear.File.FullName 'right.xlsx' 'Prior-year source selection'

$layout = @(Get-GssWorkbookLayoutPlan 4 13)
Assert-Equal (($layout | Where-Object Sheet -eq 'README').PrintArea) '$A$1:$A$34' 'README print area'
Assert-Equal (($layout | Where-Object Sheet -eq 'README').FitTall) 2 'README two-page maximum'
Assert-Equal (($layout | Where-Object Sheet -eq 'Quick_Read_WoW').PrintArea) '$A$1:$Q$9' 'WoW dynamic print area'
Assert-Equal (($layout | Where-Object Sheet -eq 'Quick_Read_YoY').PrintArea) '$A$1:$Q$9' 'YoY dynamic print area'
Assert-Equal (($layout | Where-Object Sheet -eq 'Drivers_Detail').PrintArea) '$A$1:$AB$18' 'Drivers dynamic print area'
Assert-Equal (($layout | Where-Object Sheet -eq 'Exec_Dashboard').PrintArea) '$A$1:$AD$57' 'Dashboard print area includes chart'
Assert-Equal (($layout | Where-Object Sheet -eq 'Email Comparison').PrintArea) '$A$1:$AB$38' 'Email comparison print area'

$emailComparisonLayout = Get-GssEmailComparisonLayoutPlan
Assert-Equal ($emailComparisonLayout.WrappedTableHeaderRanges -join ',') 'C5:D5,C25:D25' 'Email comparison table headers wrap'
Assert-Equal (($emailComparisonLayout.VisibleColumnWidths | ForEach-Object { '{0}:{1}' -f $_.Column, $_.Width }) -join ',') '2:22,3:18.5,4:22.5' 'Email comparison visible column widths'

$guardrails = Get-GssGuardrailPlan
Assert-Equal $guardrails.ProtectWorkbookStructure $true 'Workbook structure protection plan'
Assert-Equal $guardrails.ProtectAllWorksheets $true 'Worksheet protection plan'
Assert-Equal $guardrails.LockScope 'AllCells' 'Worksheet lock scope'
Assert-Equal ($guardrails.UnlockedCells -join ',') 'Exec_Dashboard!C3,Exec_Dashboard!F3' 'Only dashboard selectors unlocked'
Assert-Equal $guardrails.ValidationShowError $true 'Validation errors enabled'
Assert-Equal $guardrails.IntentionalGapIsBlocker $false 'Intentional 2025 gap is informational'

$qaPlan = @(Get-GssQaCheckPlan 5 14 5096)
Assert-Equal $qaPlan.Count 7 'QA check count'
Assert-True (($qaPlan | Where-Object Row -eq 7).Formula -like '*UPPER*Entities*TRUE*') 'Active entity text TRUE normalization'
Assert-True (($qaPlan | Where-Object Row -eq 7).Formula.Contains('="1"')) 'Active entity numeric 1 normalization'
Assert-True (($qaPlan | Where-Object Row -eq 11).Formula.Contains('="1"')) 'Dashboard selector numeric 1 normalization'
Assert-True (($qaPlan | Where-Object Row -eq 7).Formula -like '*MMULT*') 'Current-period blank metric detection'
Assert-True (($qaPlan | Where-Object Row -eq 8).Formula -like '*$B$3-7>DATE(2025,7,6)*$B$3-7<DATE(2025,10,5)*') 'Prior-week intentional gap exemption'
Assert-True (($qaPlan | Where-Object Row -eq 8).DetailFormula -like '*Intentional 2025 history gap*') 'Prior-week intentional gap detail'
Assert-True (-not (($qaPlan | Where-Object Row -eq 7).Formula -like '*DATE(2025,7,6)*')) 'Current week still requires complete data'
Assert-True (($qaPlan | Where-Object Row -eq 12).Formula -like '*>=8*') 'Capacity reserves two weekly loads'

$capacityAtLimit = Get-GssRawDataCapacityPlan 5088 8 5096
Assert-Equal $capacityAtLimit.AvailableRows 8 'Capacity rows available'
Assert-Equal $capacityAtLimit.ProjectedLastRow 5096 'Capacity projected last row'
Assert-Equal $capacityAtLimit.Fits $true 'Capacity accepts exact formula bound'
$capacityOverLimit = Get-GssRawDataCapacityPlan 5089 8 5096
Assert-Equal $capacityOverLimit.Fits $false 'Capacity rejects rows beyond formula bound'
Assert-ThrowsLike { Assert-GssRawDataCapacity 5089 8 5096 } '*only 7 row(s) remain through row 5096*No rows were appended.*' 'Capacity failure occurs before append'

$entitySet = Get-GssEntitySetValidation @(
    'All Franchisees|(TOTAL)',
    'Sorensen|(TOTAL)',
    'Sorensen|9354 Richmond',
    'Sorensen|9355 Virginia Beach'
)
Assert-Equal $entitySet.IsValid $true 'Exact entity set passes'
Assert-Equal $entitySet.ActualCount 4 'Exact entity set row count'
Assert-ThrowsLike {
    Assert-GssExactEntitySet @(
        'All Franchisees|(TOTAL)',
        'Sorensen|(TOTAL)',
        'Sorensen|9354 Richmond'
    ) 'partial week'
} '*missing: Sorensen|9355 Virginia Beach*No workbook mutation was performed.*' 'Partial entity week fails before mutation'
Assert-ThrowsLike {
    Assert-GssExactEntitySet @(
        'All Franchisees|(TOTAL)',
        'Sorensen|(TOTAL)',
        'Sorensen|9354 Richmond',
        'Sorensen|9354 Richmond',
        'Sorensen|9355 Virginia Beach'
    ) 'duplicate week'
} '*duplicates: Sorensen|9354 Richmond (2 rows)*No workbook mutation was performed.*' 'Duplicate entity week fails before mutation'
Assert-ThrowsLike {
    Assert-GssExactEntitySet @(
        'All Franchisees|(TOTAL)',
        'Sorensen|(TOTAL)',
        'Sorensen|9354 Richmond',
        'Sorensen|9355 Virginia Beach',
        'Other|9999'
    ) 'unexpected week'
} '*unexpected: Other|9999*No workbook mutation was performed.*' 'Unexpected entity week fails before mutation'

$fingerprintArguments = @{
    RunId = '11111111-1111-1111-1111-111111111111'
    HostName = 'GSS-HOST'
    CurrentWeekEnding = '2026-07-19'
    StartingWorkbookSha256 = ('a' * 64)
    CurrentSourceSha256 = ('b' * 64)
    PriorYearSourceSha256 = ('c' * 64)
    StagedWorkbookSha256 = ('d' * 64)
    StagedPdfSha256 = ('e' * 64)
    ProgramRelease = 'v1.0.0'
}
$fingerprint = Get-GssPreparedRunFingerprint @fingerprintArguments
Assert-Equal $fingerprint.Length 64 'Prepared fingerprint SHA256 length'
Assert-Equal (Get-GssPreparedRunFingerprint @fingerprintArguments) $fingerprint 'Prepared fingerprint deterministic'
$fingerprintArguments.StagedPdfSha256 = ('f' * 64)
Assert-True ((Get-GssPreparedRunFingerprint @fingerprintArguments) -ne $fingerprint) 'Prepared fingerprint binds staged PDF'

$updaterSource = Get-Content -LiteralPath (Join-Path $scriptRoot 'Update-GSS-MainWorkbook.ps1') -Raw
$capacityCallIndex = $updaterSource.IndexOf('$capacityPlan = Assert-GssRawDataCapacity')
$guardrailRemovalIndex = $updaterSource.IndexOf('Remove-GssWorkbookGuardrails $targetWb')
$appendCallIndex = $updaterSource.IndexOf('$written = Add-RawDataRows')
Assert-True ($capacityCallIndex -ge 0 -and $capacityCallIndex -lt $guardrailRemovalIndex) 'Capacity preflight precedes workbook mutation'
Assert-True ($capacityCallIndex -lt $appendCallIndex) 'Capacity preflight precedes Raw_Data append'

$allCellsAcquireIndex = $updaterSource.IndexOf('$allCells = $worksheet.Cells')
$allCellsLockIndex = $updaterSource.IndexOf('$allCells.Locked = $true')
$dashboardUnlockIndex = $updaterSource.IndexOf('Set-GssDashboardValidation $Workbook $Configuration')
Assert-True ($allCellsAcquireIndex -ge 0 -and $allCellsAcquireIndex -lt $allCellsLockIndex) 'Guardrails acquire all worksheet cells before locking'
Assert-True ($allCellsLockIndex -lt $dashboardUnlockIndex) 'All worksheet cells are locked before dashboard selectors are unlocked'

$targetSaveIndex = $updaterSource.IndexOf('$targetWb.Save()')
$targetCloseIndex = $updaterSource.IndexOf('$targetWb.Close($false)', $targetSaveIndex)
$fileEvidenceIndex = $updaterSource.IndexOf('$fileEvidence = @(', $targetSaveIndex)
Assert-True ($targetSaveIndex -ge 0 -and $targetSaveIndex -lt $targetCloseIndex) 'Target workbook is saved before it is closed for hashing'
Assert-True ($targetCloseIndex -lt $fileEvidenceIndex) 'Target workbook is closed before file evidence is hashed'

$safeLauncherSource = Get-Content -LiteralPath (Join-Path $scriptRoot 'Invoke-GSS-SafeWorkbookUpdate.ps1') -Raw
Assert-True ($safeLauncherSource.Contains('$copyRun = & $updater')) 'Safe launcher captures copy-test run object'
Assert-True ($safeLauncherSource.Contains('-RunId $runId')) 'Safe launcher binds copy test to unique run ID'
Assert-True (
    $safeLauncherSource.Contains('-RunLogPath $copyRun.LogPath') -and
    $safeLauncherSource.Contains('LogPath = $RunLogPath') -and
    $safeLauncherSource.Contains('OutputObject = $true') -and
    $safeLauncherSource.Contains('return & $AnalyzerPath @arguments')
) 'Safe launcher pairs copy analysis to exact log'
Assert-True ($safeLauncherSource.Contains('$liveRun = & $updater')) 'Safe launcher captures live promotion object'
Assert-True ($safeLauncherSource.Contains('-PreparedRunLogPath $copyRun.LogPath')) 'Live promotion reuses exact staged run'
Assert-True ($safeLauncherSource.Contains('-ExpectedFingerprint $copyRun.RunFingerprint')) 'Live promotion requires exact reviewed fingerprint'
Assert-True ($safeLauncherSource.Contains('-DrivePreparedManifestPath $backupPrepare.PreparedManifestPath')) 'Live promotion requires exact Drive prepared manifest'
Assert-True ($safeLauncherSource.Contains('-ExpectedDrivePreparedManifestSha256 $backupPrepare.PreparedManifestSha256')) 'Live promotion binds Drive prepared manifest hash'
Assert-True ($safeLauncherSource.Contains('Global\GSSSurveyWorkbookAutomationTransaction')) 'Safe launcher owns named workstation mutex'
Assert-True ($safeLauncherSource.Contains('& $releaseValidator -RepoRoot $repoRoot')) 'Safe launcher enforces release integrity first'
Assert-True (-not $safeLauncherSource.Contains('SkipReleaseIntegrityCheck')) 'Safe launcher exposes no release-integrity bypass'
Assert-True (-not $safeLauncherSource.Contains('[string]$DriveBackupScript')) 'Safe launcher exposes no injectable Drive coordinator'
Assert-True ($safeLauncherSource.Contains("Unresolved GSS transaction")) 'Safe launcher blocks a new run while an unresolved receipt exists'
Assert-True ($safeLauncherSource.Contains("'RollbackAttempting'")) 'Safe launcher records rollback intent before invoking recovery'
Assert-True ($safeLauncherSource.Contains("'RollbackBlocked'")) 'Safe launcher preserves rollback failure as a distinct state'
$prepareIndex = $safeLauncherSource.IndexOf('-Operation Prepare')
$liveApplyIndex = $safeLauncherSource.IndexOf('$liveRun = & $updater')
$finalizeIndex = $safeLauncherSource.IndexOf('-Operation Finalize')
$packageIndex = $safeLauncherSource.LastIndexOf('-PublishPackage')
Assert-True ($prepareIndex -ge 0 -and $prepareIndex -lt $liveApplyIndex) 'Drive preparation precedes live promotion'
Assert-True ($finalizeIndex -gt $liveApplyIndex -and $finalizeIndex -lt $packageIndex) 'Drive finalization precedes package publication'
Assert-True ($safeLauncherSource.Contains('$arguments.PublishEmailPackage = $true')) 'Safe analysis helper maps package publication to the analyzer switch'
Assert-True ($safeLauncherSource.Contains("if (`$confirmation -cne 'APPLY')")) 'Literal case-sensitive APPLY remains required'
Assert-True ($updaterSource.Contains('Direct live mutation is disabled.')) 'Updater rejects unbound direct live apply'
Assert-True ($updaterSource.Contains('Live promotion requires the hash-verified Drive prepared manifest.')) 'Updater rejects live promotion without prepared Drive evidence'
Assert-True ($updaterSource.Contains('Drive prepared manifest is outside the exact commissioned private Drive run directory.')) 'Updater binds prepared manifest to commissioned Drive root'
Assert-True ($updaterSource.Contains("Apply release")) 'Updater independently enforces the approved release on direct apply'
Assert-True ($updaterSource.Contains('[System.IO.File]::Replace($workbookCandidate, $MainWorkbookPath, $backupPath, $true)')) 'Live workbook promotion is atomic with run-owned backup'
Assert-True ($updaterSource.Contains('Rollback blocked because the live workbook changed after this run.')) 'Rollback is conflict aware'
Assert-True ($updaterSource.Contains('Write-GssAtomicJson -Path $logPath')) 'Run receipt is written atomically'
Assert-True ($updaterSource.Contains('Role = ''reviewed_staged_workbook''')) 'Prepared summary identifies staged workbook backup artifact'
Assert-True ($updaterSource.Contains('Role = ''run_owned_pre_apply_workbook_backup''')) 'Committed summary identifies run-owned pre-apply backup artifact'
Assert-True ($updaterSource.Contains('ReleaseArchives = $releaseArchives')) 'Prepared summary includes exact release archives'
Assert-True ($updaterSource.Contains('$pdfName = ''GSS Email Comparison {0}.pdf'' -f $LatestSource.WeekEnding.ToString(''MMddyy'')')) 'PDF filename uses reporting week'
Assert-True (-not $updaterSource.Contains('$LatestSource.File.LastWriteTime.ToString(''MMddyy'')')) 'PDF filename does not use source modification date'

$resumeSource = Get-Content -LiteralPath (Join-Path $scriptRoot 'Resume-GSS-PendingFinalize.ps1') -Raw
$retryIndex = $resumeSource.IndexOf('-Operation RetryFinalize')
$resumePublishIndex = $resumeSource.IndexOf('-PublishEmailPackage')
Assert-True ($retryIndex -ge 0 -and $retryIndex -lt $resumePublishIndex) 'Resume commits Drive before package publication'
Assert-True (-not $resumeSource.Contains(' -Apply ')) 'Resume path never applies a workbook'
Assert-True (-not $resumeSource.Contains('SkipReleaseIntegrityCheck')) 'Resume exposes no release-integrity bypass'
Assert-True (-not $resumeSource.Contains('[string]$DriveBackupScript')) 'Resume exposes no injectable Drive coordinator'
Assert-True ($resumeSource.Contains('active-transaction.json')) 'Resume preserves and owns the unresolved transaction marker'
Assert-True ($resumeSource.Contains("'BackupCommitted'")) 'Resume records Drive commit before package work'
Assert-True ($resumeSource.Contains("'PackageBlocked'")) 'Resume distinguishes package failure from Drive finalization'
Assert-True ($resumeSource.Contains("'BackupBlocked'")) 'Resume preserves non-retryable Drive blocked status'
Assert-True ($safeLauncherSource.Contains("if (`$finalizeStatus -eq 'Blocked')")) 'Safe launcher handles Drive blocked separately from pending finalize'
Assert-True ($safeLauncherSource.Contains("'BackupBlocked'")) 'Safe launcher records non-retryable Drive conflict for manual review'
Assert-True ($safeLauncherSource.Contains('-Operation Abort')) 'Prepared Drive snapshot is explicitly aborted after safe rollback or pre-commit failure'

$portableRoot = 'C:\Users\bryan\Dropbox\Marketing\GSS Surveys'
$otherProfilePath = "C:\Users\Other User\Dropbox\Marketing\GSS Surveys\04 Email Comparison PDFs\report.pdf"
Assert-Equal (ConvertTo-GssDropboxRelativePath -Path $otherProfilePath -FolderPath $portableRoot) '04 Email Comparison PDFs/report.pdf' 'Cross-profile portable path recovery'
Assert-ThrowsLike { ConvertTo-GssDropboxRelativePath -Path '..\escape.txt' -FolderPath $portableRoot } '*traversal segment*' 'Portable path traversal rejection'

Write-Host 'GSS non-Excel logic tests passed.'
