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

$atomicTestRoot = Join-Path ([System.IO.Path]::GetTempPath()) ('gss-atomic-replace-' + [guid]::NewGuid().ToString('N'))
$lockProcess = $null
New-Item -ItemType Directory -Path $atomicTestRoot -Force | Out-Null
try {
    $atomicTestPath = Join-Path $atomicTestRoot 'receipt.json'
    Write-GssAtomicText -Path $atomicTestPath -Content '{"status":"first"}'
    Write-GssAtomicText -Path $atomicTestPath -Content '{"status":"replaced"}'
    Assert-Equal (Get-Content -LiteralPath $atomicTestPath -Raw) '{"status":"replaced"}' 'Atomic text replaces an existing receipt'
    Assert-Equal @(
        Get-ChildItem -LiteralPath $atomicTestRoot -File |
            Where-Object { $_.Name -like '.t-*' -or $_.Name -like '.b-*' }
    ).Count 0 'Atomic replacement cleans temporary and backup files'

    $lockedPath = Join-Path $atomicTestRoot 'locked-receipt.json'
    $lockMarkerPath = Join-Path $atomicTestRoot 'lock-acquired.txt'
    Write-GssAtomicText -Path $lockedPath -Content '{"status":"before-lock"}'
    $escapedLockedPath = $lockedPath.Replace("'", "''")
    $escapedMarkerPath = $lockMarkerPath.Replace("'", "''")
    $lockCommand = @"
`$stream = [System.IO.File]::Open(
    '$escapedLockedPath',
    [System.IO.FileMode]::Open,
    [System.IO.FileAccess]::Read,
    [System.IO.FileShare]::None
)
try {
    [System.IO.File]::WriteAllText('$escapedMarkerPath', 'locked')
    Start-Sleep -Milliseconds 1000
}
finally {
    `$stream.Dispose()
}
"@
    $encodedLockCommand = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($lockCommand))
    $lockProcess = Start-Process -FilePath 'powershell.exe' -ArgumentList @(
        '-NoLogo',
        '-NoProfile',
        '-NonInteractive',
        '-EncodedCommand',
        $encodedLockCommand
    ) -WindowStyle Hidden -PassThru
    for ($waitAttempt = 1; $waitAttempt -le 200 -and -not (Test-Path -LiteralPath $lockMarkerPath -PathType Leaf); $waitAttempt++) {
        Start-Sleep -Milliseconds 10
    }
    Assert-True (Test-Path -LiteralPath $lockMarkerPath -PathType Leaf) 'Atomic replacement lock test acquired the destination'
    Write-GssAtomicText -Path $lockedPath -Content '{"status":"after-lock"}'
    $lockProcess.WaitForExit()
    Assert-Equal (Get-Content -LiteralPath $lockedPath -Raw) '{"status":"after-lock"}' 'Atomic text retries a transient sharing violation'
}
finally {
    if ($null -ne $lockProcess -and -not $lockProcess.HasExited) {
        $lockProcess.Kill()
        $lockProcess.WaitForExit()
    }
    if ($null -ne $lockProcess) {
        $lockProcess.Dispose()
    }
    if (Test-Path -LiteralPath $atomicTestRoot -PathType Container) {
        Remove-Item -LiteralPath $atomicTestRoot -Recurse -Force
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
$operatorLauncherTemplate = Get-Content -LiteralPath (Join-Path (Split-Path -Parent $scriptRoot) 'templates\Run GSS Update After Upload.cmd') -Raw
$operatorGuideTemplate = Get-Content -LiteralPath (Join-Path (Split-Path -Parent $scriptRoot) 'templates\00 START HERE - GSS Survey Updates.txt') -Raw
Assert-True ($operatorLauncherTemplate.Contains('Only one computer may run GSS at a time.')) 'Operator launcher warns against concurrent cross-workstation runs'
Assert-True ($operatorGuideTemplate.Contains('Run GSS on only one computer at a time.')) 'Operator guide requires one workstation run at a time'
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
$resumeCurrentReleasePublishIndex = $resumeSource.LastIndexOf('-PublishEmailPackage')
$resumeExpectedEvidenceIndex = $resumeSource.IndexOf('-ExpectedPackageInputEvidence $priorReleasePackageInputEvidence')
Assert-True ($retryIndex -ge 0 -and $retryIndex -lt $resumePublishIndex) 'Resume commits Drive before package publication'
Assert-True (-not $resumeSource.Contains(' -Apply ')) 'Resume path never applies a workbook'
Assert-True (-not $resumeSource.Contains('SkipReleaseIntegrityCheck')) 'Resume exposes no release-integrity bypass'
Assert-True (-not $resumeSource.Contains('[string]$DriveBackupScript')) 'Resume exposes no injectable Drive coordinator'
Assert-True ($resumeSource.Contains('active-transaction.json')) 'Resume preserves and owns the unresolved transaction marker'
Assert-True ($resumeSource.Contains("'BackupCommitted'")) 'Resume records Drive commit before package work'
Assert-True ($resumeSource.Contains("'PackageBlocked'")) 'Resume distinguishes package failure from Drive finalization'
Assert-True ($resumeSource.Contains("'BackupBlocked'")) 'Resume preserves non-retryable Drive blocked status'
Assert-True ($resumeSource.Contains('Test-GssCommittedBackupSnapshot')) 'Prior-release resume cryptographically validates the committed Drive snapshot'
Assert-True ($resumeSource.Contains("RunRelease = 'v1.1.8'") -and $resumeSource.Contains("CurrentRelease = 'v1.1.9'")) 'Prior-release resume declares only the reviewed v1.1.8 to v1.1.9 compatibility bridge'
Assert-True (-not $resumeSource.Contains('$runVersion.Patch + 1')) 'Adjacent patch releases do not gain package recovery automatically'
Assert-True ($resumeSource.Contains("[string]`$Receipt.TransactionStatus -cne 'PackageBlocked'")) 'Prior-release resume requires the exact package-blocked transaction state'
Assert-True ($resumeSource.Contains("[string]`$Receipt.BackupStatus -cne 'Committed'")) 'Prior-release resume requires a committed local backup state'
Assert-True ($resumeSource.Contains("[string]`$document.snapshot_purpose -cne 'WorkbookTransaction'")) 'Prior-release resume rejects non-workbook Drive snapshots'
Assert-True ($resumeSource.Contains("[string]`$manifest.host -cne [Environment]::MachineName")) 'Prior-release Drive evidence is bound to this workstation'
Assert-True ($resumeSource.Contains('cat-file -t $tagReference')) 'Prior-release recovery requires an annotated historical release tag object'
Assert-True ($resumeSource.Contains("[string]`$Receipt.ReleaseCommit -cne `$ExpectedReleaseCommit")) 'Prior-release receipt commit is bound to the peeled annotated tag commit'
Assert-True ($resumeSource.Contains("Add-GssResumeReceiptValue `$receipt 'PackageRecoveryBackupEvidence' `$finalize")) 'Prior-release package recovery records its verified snapshot evidence separately'
Assert-True ($resumeSource.Contains('PackageOnlyRecovery = $true')) 'Prior-release recovery is explicitly identified as package-only'
Assert-True ($resumeSource.Contains('Assert-GssResumePackageInputSet')) 'Prior-release recovery binds live package inputs to the committed workbook snapshot'
Assert-True ($resumeSource.Contains('Assert-GssResumeHistoricalRecoveryEvidence')) 'Prior-release recovery validates historical metadata through committed RecoveryOnly snapshots'
Assert-True (
    $resumeExpectedEvidenceIndex -gt $resumePublishIndex -and
    $resumeExpectedEvidenceIndex -lt $resumeCurrentReleasePublishIndex
) 'Prior-release publisher receives the validated package-input evidence while current-release publication keeps its existing call shape'

$resumeTokens = $null
$resumeParseErrors = $null
$resumeAst = [System.Management.Automation.Language.Parser]::ParseInput(
    $resumeSource,
    [ref]$resumeTokens,
    [ref]$resumeParseErrors
)
Assert-Equal $resumeParseErrors.Count 0 'Resume script parses without errors'
foreach ($resumeFunctionName in @(
    'Get-GssResumeReleaseVersion',
    'Get-GssResumeReleaseMode',
    'Get-GssResumeAnnotatedReleaseCommit',
    'Get-GssResumeHash',
    'Get-GssResumeTextHash',
    'Test-GssResumeSamePath',
    'ConvertTo-GssResumeManifestPortablePath',
    'ConvertTo-GssResumePortablePath',
    'Get-GssResumePackageInputDescriptor',
    'Get-GssResumeTrustedBackupManifest',
    'Assert-GssResumePackageInputSet',
    'Assert-GssResumeHistoricalRecoveryEvidence',
    'Assert-GssResumePriorReleaseReceipt',
    'Get-GssResumeCommittedSnapshotEvidence'
)) {
    $resumeFunction = $resumeAst.Find({
        param($node)
        $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and
        $node.Name -eq $resumeFunctionName
    }, $true)
    Assert-True ($null -ne $resumeFunction) "Resume helper exists: $resumeFunctionName"
    Set-Item -Path "Function:\$resumeFunctionName" -Value $resumeFunction.Body.GetScriptBlock()
}

Assert-Equal (Get-GssResumeReleaseMode -RunRelease 'v1.1.9' -CurrentRelease 'v1.1.9') 'CurrentRelease' 'Resume accepts the exact current release'
Assert-Equal (Get-GssResumeReleaseMode -RunRelease 'v1.1.10' -CurrentRelease 'v1.1.10') 'CurrentRelease' 'Resume preserves exact current-release behavior for future canonical releases'
Assert-Equal (Get-GssResumeReleaseMode -RunRelease 'v1.1.8' -CurrentRelease 'v1.1.9') 'DeclaredPriorReleasePackageRecovery' 'Resume accepts the explicitly reviewed v1.1.8 to v1.1.9 package recovery bridge'
Assert-ThrowsLike {
    Get-GssResumeReleaseMode -RunRelease 'v1.1.7' -CurrentRelease 'v1.1.9'
} '*does not have a code-reviewed package-recovery compatibility declaration*' 'Resume rejects undeclared older releases'
Assert-ThrowsLike {
    Get-GssResumeReleaseMode -RunRelease 'v1.1.9' -CurrentRelease 'v1.1.10'
} '*does not have a code-reviewed package-recovery compatibility declaration*' 'Resume rejects a future adjacent patch without a reviewed declaration'
Assert-ThrowsLike {
    Get-GssResumeReleaseMode -RunRelease 'v1.1.9' -CurrentRelease 'v1.2.0'
} '*does not have a code-reviewed package-recovery compatibility declaration*' 'Resume rejects a minor-version predecessor'
Assert-ThrowsLike {
    Get-GssResumeReleaseMode -RunRelease '1.1.9' -CurrentRelease '1.1.9'
} '*not a canonical vMAJOR.MINOR.PATCH tag*' 'Resume rejects malformed release tags even when equal'

$resumeFixtureRun = [pscustomobject]@{
    RunId = '11111111-1111-1111-1111-111111111111'
    RunFingerprint = ('b' * 64)
    HostName = [Environment]::MachineName
    ProgramRelease = 'v1.1.8'
}
$resumeFixtureReceipt = [pscustomobject]@{
    ReceiptSchemaVersion = 1
    RunId = $resumeFixtureRun.RunId
    RunFingerprint = $resumeFixtureRun.RunFingerprint
    HostName = $resumeFixtureRun.HostName
    ProgramRelease = $resumeFixtureRun.ProgramRelease
    ReleaseIntegrityStatus = 'Passed'
    ReleaseCommit = ('a' * 40)
    LiveRunLogPath = 'C:\GSS\_automation_runs\logs\committed-apply.json'
    TransactionStatus = 'PackageBlocked'
    BackupStatus = 'Committed'
    PackagePublished = $false
    BackupPrepare = [pscustomobject]@{ PreparedManifestSha256 = ('c' * 64) }
    BackupFinalize = [pscustomobject]@{
        Status = 'Committed'
        BackupStatus = 'Committed'
        RunId = $resumeFixtureRun.RunId
        Fingerprint = $resumeFixtureRun.RunFingerprint
    }
}
Assert-GssResumePriorReleaseReceipt `
    -Run $resumeFixtureRun `
    -Receipt $resumeFixtureReceipt `
    -LiveRunLogPath 'C:\GSS\_automation_runs\logs\committed-apply.json' `
    -ExpectedReleaseCommit ('a' * 40)
$resumeFixtureReceipt.ReleaseCommit = ('d' * 40)
Assert-ThrowsLike {
    Assert-GssResumePriorReleaseReceipt `
        -Run $resumeFixtureRun `
        -Receipt $resumeFixtureReceipt `
        -LiveRunLogPath 'C:\GSS\_automation_runs\logs\committed-apply.json' `
        -ExpectedReleaseCommit ('a' * 40)
} '*does not match the committed run, release-integrity evidence, log, and workstation*' 'Prior-release recovery rejects a receipt commit that differs from the annotated tag commit'
$resumeFixtureReceipt.ReleaseCommit = ('a' * 40)
$resumeFixtureReceipt.TransactionStatus = 'Committed'
Assert-ThrowsLike {
    Assert-GssResumePriorReleaseReceipt `
        -Run $resumeFixtureRun `
        -Receipt $resumeFixtureReceipt `
        -LiveRunLogPath 'C:\GSS\_automation_runs\logs\committed-apply.json' `
        -ExpectedReleaseCommit ('a' * 40)
} '*requires an unpublished PackageBlocked receipt with BackupStatus Committed*' 'Prior-release recovery rejects an already committed package state'

$resumeSnapshotTestRoot = Join-Path ([System.IO.Path]::GetTempPath()) ('gss-resume-snapshot-' + [guid]::NewGuid().ToString('N'))
$resumeSnapshotPath = Join-Path $resumeSnapshotTestRoot 'committed-snapshot'
$originalDriveRootContext = (Get-Item -LiteralPath Function:Get-GssDriveBackupRootContext).ScriptBlock
$originalDriveSnapshotFinder = (Get-Item -LiteralPath Function:Find-GssDriveBackupSnapshot).ScriptBlock
$originalCommittedSnapshotValidator = (Get-Item -LiteralPath Function:Test-GssCommittedBackupSnapshot).ScriptBlock
New-Item -ItemType Directory -Path $resumeSnapshotPath -Force | Out-Null
try {
    $resumeSnapshotRunId = '22222222-2222-2222-2222-222222222222'
    $resumeSnapshotFingerprint = ('1' * 64)
    $resumeSnapshotDrive = [pscustomobject]@{ verification_level = 'drivefs_hash_verified' }
    $resumeSnapshotPreparedManifest = [pscustomobject]@{
        status = 'Prepared'
        run_id = $resumeSnapshotRunId
        fingerprint = $resumeSnapshotFingerprint
        snapshot_purpose = 'WorkbookTransaction'
        host = [Environment]::MachineName
        release = 'v1.1.8'
        report_week = '2026-08-23'
        drive = $resumeSnapshotDrive
    }
    $resumeSnapshotManifest = [pscustomobject]@{
        status = 'Committed'
        run_id = $resumeSnapshotRunId
        fingerprint = $resumeSnapshotFingerprint
        snapshot_purpose = 'WorkbookTransaction'
        host = [Environment]::MachineName
        release = 'v1.1.8'
        report_week = '2026-08-23'
        drive = $resumeSnapshotDrive
    }
    $resumePreparedManifestPath = Join-Path $resumeSnapshotPath 'prepared-manifest.json'
    $resumeManifestPath = Join-Path $resumeSnapshotPath 'backup-manifest.json'
    $resumeCommitReceiptPath = Join-Path $resumeSnapshotPath 'commit-receipt.json'
    $utf8WithoutBom = New-Object System.Text.UTF8Encoding($false)
    [System.IO.File]::WriteAllText(
        $resumePreparedManifestPath,
        ($resumeSnapshotPreparedManifest | ConvertTo-Json -Depth 5 -Compress),
        $utf8WithoutBom
    )
    [System.IO.File]::WriteAllText(
        $resumeManifestPath,
        ($resumeSnapshotManifest | ConvertTo-Json -Depth 5 -Compress),
        $utf8WithoutBom
    )
    $resumePreparedManifestHash = (Get-FileHash -LiteralPath $resumePreparedManifestPath -Algorithm SHA256).Hash.ToLowerInvariant()
    $resumeManifestHash = (Get-FileHash -LiteralPath $resumeManifestPath -Algorithm SHA256).Hash.ToLowerInvariant()
    $resumeSnapshotCommitReceipt = [pscustomobject]@{
        status = 'Committed'
        run_id = $resumeSnapshotRunId
        fingerprint = $resumeSnapshotFingerprint
        snapshot_purpose = 'WorkbookTransaction'
        verification_level = 'drivefs_hash_verified'
        backup_manifest_sha256 = $resumeManifestHash
        prepared_manifest_sha256 = $resumePreparedManifestHash
    }
    [System.IO.File]::WriteAllText(
        $resumeCommitReceiptPath,
        ($resumeSnapshotCommitReceipt | ConvertTo-Json -Depth 5 -Compress),
        $utf8WithoutBom
    )

    $resumeSnapshotRun = [pscustomobject]@{
        RunId = $resumeSnapshotRunId
        RunFingerprint = $resumeSnapshotFingerprint
        HostName = [Environment]::MachineName
        ProgramRelease = 'v1.1.8'
        CurrentWeekEnding = '2026-08-23'
        DrivePreparedManifestSha256 = $resumePreparedManifestHash
    }
    $resumeSnapshotTransactionReceipt = [pscustomobject]@{
        BackupPrepare = [pscustomobject]@{
            PreparedManifestSha256 = $resumePreparedManifestHash
        }
        BackupFinalize = [pscustomobject]@{
            BackupManifestSha256 = $resumeManifestHash
            SnapshotPath = $resumeSnapshotPath
            BackupManifestPath = $resumeManifestPath
            CommitReceiptPath = $resumeCommitReceiptPath
        }
    }
    $script:resumeCommittedSnapshotFixture = [pscustomobject]@{
        RootPath = $resumeSnapshotTestRoot
        RunId = $resumeSnapshotRunId
        SnapshotPath = $resumeSnapshotPath
        Validation = [pscustomobject]@{
            Receipt = $resumeSnapshotCommitReceipt
            Manifest = $resumeSnapshotManifest
            PreparedManifest = $resumeSnapshotPreparedManifest
            ReceiptPath = $resumeCommitReceiptPath
            ManifestPath = $resumeManifestPath
            PreparedManifestPath = $resumePreparedManifestPath
            ManifestSha256 = $resumeManifestHash
            PreparedManifestSha256 = $resumePreparedManifestHash
            ValidatedFileCount = 3
            ValidatedPreparedFileCount = 2
        }
    }

    Set-Item -LiteralPath Function:Get-GssDriveBackupRootContext -Value {
        [pscustomobject]@{ RootPath = $script:resumeCommittedSnapshotFixture.RootPath }
    }
    Set-Item -LiteralPath Function:Find-GssDriveBackupSnapshot -Value {
        param([string]$RootPath, [string]$RunId)
        if (-not (Test-GssResumeSamePath -Left $RootPath -Right $script:resumeCommittedSnapshotFixture.RootPath) -or
            $RunId -cne $script:resumeCommittedSnapshotFixture.RunId) {
            throw 'Snapshot finder received unexpected fixture identity.'
        }
        [pscustomobject]@{
            Path = $script:resumeCommittedSnapshotFixture.SnapshotPath
            IsPartial = $false
        }
    }
    Set-Item -LiteralPath Function:Test-GssCommittedBackupSnapshot -Value {
        param([string]$SnapshotPath)
        if (-not (Test-GssResumeSamePath -Left $SnapshotPath -Right $script:resumeCommittedSnapshotFixture.SnapshotPath)) {
            throw 'Snapshot validator received an unexpected fixture path.'
        }
        $script:resumeCommittedSnapshotFixture.Validation
    }

    $resumeSnapshotEvidence = Get-GssResumeCommittedSnapshotEvidence `
        -Run $resumeSnapshotRun `
        -Receipt $resumeSnapshotTransactionReceipt
    Assert-Equal $resumeSnapshotEvidence.Status 'Committed' 'Committed snapshot fixture returns committed status'
    Assert-Equal $resumeSnapshotEvidence.RunId $resumeSnapshotRunId 'Committed snapshot fixture preserves run identity'
    Assert-Equal $resumeSnapshotEvidence.Fingerprint $resumeSnapshotFingerprint 'Committed snapshot fixture preserves fingerprint'
    Assert-Equal $resumeSnapshotEvidence.BackupManifestSha256 $resumeManifestHash 'Committed snapshot fixture returns the real final manifest hash'
    Assert-Equal $resumeSnapshotEvidence.PreparedManifestSha256 $resumePreparedManifestHash 'Committed snapshot fixture returns the real prepared manifest hash'
    Assert-Equal $resumeSnapshotEvidence.ValidatedFileCount 3 'Committed snapshot fixture returns final validation count'
    Assert-Equal $resumeSnapshotEvidence.ValidatedPreparedFileCount 2 'Committed snapshot fixture returns prepared validation count'
    Assert-Equal $resumeSnapshotEvidence.PackageOnlyRecovery $true 'Committed snapshot fixture is package-only evidence'

    $resumeSnapshotManifest.fingerprint = ('9' * 64)
    Assert-ThrowsLike {
        Get-GssResumeCommittedSnapshotEvidence `
            -Run $resumeSnapshotRun `
            -Receipt $resumeSnapshotTransactionReceipt
    } '*snapshot identity or purpose does not match*' 'Committed snapshot fixture rejects a mismatched manifest fingerprint'
    $resumeSnapshotManifest.fingerprint = $resumeSnapshotFingerprint

    $resumeSnapshotTransactionReceipt.BackupFinalize.BackupManifestSha256 = ('f' * 64)
    Assert-ThrowsLike {
        Get-GssResumeCommittedSnapshotEvidence `
            -Run $resumeSnapshotRun `
            -Receipt $resumeSnapshotTransactionReceipt
    } '*manifest hashes do not match*' 'Committed snapshot fixture rejects a mismatched final manifest hash'
}
finally {
    Set-Item -LiteralPath Function:Get-GssDriveBackupRootContext -Value $originalDriveRootContext
    Set-Item -LiteralPath Function:Find-GssDriveBackupSnapshot -Value $originalDriveSnapshotFinder
    Set-Item -LiteralPath Function:Test-GssCommittedBackupSnapshot -Value $originalCommittedSnapshotValidator
    Remove-Variable -Name resumeCommittedSnapshotFixture -Scope Script -ErrorAction SilentlyContinue
    if (Test-Path -LiteralPath $resumeSnapshotTestRoot -PathType Container) {
        Remove-Item -LiteralPath $resumeSnapshotTestRoot -Recurse -Force
    }
}

$resumeInputTestRoot = Join-Path ([System.IO.Path]::GetTempPath()) ('gss-resume-inputs-' + [guid]::NewGuid().ToString('N'))
$resumeInputMainDirectory = Join-Path $resumeInputTestRoot '01 Main Workbook'
$resumeInputPdfDirectory = Join-Path $resumeInputTestRoot '04 Email Comparison PDFs'
$resumeInputRollingDirectory = Join-Path $resumeInputTestRoot '02 Weekly Rolling Source Workbooks'
$resumeInputLogDirectory = Join-Path $resumeInputTestRoot '_automation_runs\logs'
$resumeInputStateDirectory = Join-Path $resumeInputTestRoot '_automation_runs\state'
$resumeInputDetailDirectory = Join-Path $resumeInputTestRoot '03 Uploaded Survey Workbooks'
$resumeInputArchiveDirectory = Join-Path $resumeInputDetailDirectory 'Archive - Previous Uploads'
$resumeInputHistoricalDirectory = Join-Path $resumeInputTestRoot '_automation_runs\historical-recovery\fixture-transaction'
foreach ($directory in @(
    $resumeInputMainDirectory,
    $resumeInputPdfDirectory,
    $resumeInputRollingDirectory,
    $resumeInputLogDirectory,
    $resumeInputStateDirectory,
    $resumeInputDetailDirectory,
    $resumeInputArchiveDirectory,
    $resumeInputHistoricalDirectory
)) {
    New-Item -ItemType Directory -Path $directory -Force | Out-Null
}
try {
    $resumeInputTargetWorkbookPath = Join-Path $resumeInputMainDirectory 'main.xlsx'
    $resumeInputComparisonPdfPath = Join-Path $resumeInputPdfDirectory 'comparison.pdf'
    $resumeInputCurrentSourcePath = Join-Path $resumeInputRollingDirectory 'current.xlsx'
    $resumeInputPriorYearSourcePath = Join-Path $resumeInputRollingDirectory 'prior-year.xlsx'
    $resumeInputLogPath = Join-Path $resumeInputLogDirectory 'gss_update_fixture_apply.json'
    $resumeInputCurrentDetailPath = Join-Path $resumeInputDetailDirectory 'current-detail.xlsx'
    $resumeInputArchiveDetailPath = Join-Path $resumeInputArchiveDirectory 'archive-detail.xlsx'
    $resumeInputLedgerPath = Join-Path $resumeInputStateDirectory 'gss_feedback_first_seen.json'
    $resumeInputHistoricalReceiptPath = Join-Path $resumeInputHistoricalDirectory 'transaction-receipt.json'
    $resumeInputHistoricalManifestPath = Join-Path $resumeInputHistoricalDirectory 'recovery-manifest.json'
    $resumeInputHistoricalSummaryPath = Join-Path $resumeInputHistoricalDirectory 'drive-recovery-summary.json'
    $resumeInputLogText = '{"Mode":"ApplyToMainWorkbook","RunId":"fixture"}'
    $resumeInputLedgerText = '{"schema_version":"gss-feedback-first-seen/v1","entries":[]}'
    [System.IO.File]::WriteAllText($resumeInputTargetWorkbookPath, 'target-workbook-bytes', $utf8WithoutBom)
    [System.IO.File]::WriteAllText($resumeInputComparisonPdfPath, 'comparison-pdf-bytes', $utf8WithoutBom)
    [System.IO.File]::WriteAllText($resumeInputCurrentSourcePath, 'current-source-bytes', $utf8WithoutBom)
    [System.IO.File]::WriteAllText($resumeInputPriorYearSourcePath, 'prior-year-source-bytes', $utf8WithoutBom)
    [System.IO.File]::WriteAllText($resumeInputLogPath, $resumeInputLogText, $utf8WithoutBom)
    [System.IO.File]::WriteAllText($resumeInputCurrentDetailPath, 'current-detail-bytes', $utf8WithoutBom)
    [System.IO.File]::WriteAllText($resumeInputArchiveDetailPath, 'archive-detail-bytes', $utf8WithoutBom)
    [System.IO.File]::WriteAllText($resumeInputLedgerPath, $resumeInputLedgerText, $utf8WithoutBom)
    [System.IO.File]::WriteAllText($resumeInputHistoricalReceiptPath, '{"state":"Committed"}', $utf8WithoutBom)
    [System.IO.File]::WriteAllText($resumeInputHistoricalManifestPath, '{"schema_version":"fixture"}', $utf8WithoutBom)
    [System.IO.File]::WriteAllText($resumeInputHistoricalSummaryPath, '{"schema_version":"fixture"}', $utf8WithoutBom)

    $resumeInputManifestRecords = @(
        $resumeInputTargetWorkbookPath,
        $resumeInputComparisonPdfPath,
        $resumeInputCurrentSourcePath,
        $resumeInputPriorYearSourcePath,
        $resumeInputLogPath,
        $resumeInputCurrentDetailPath,
        $resumeInputArchiveDetailPath,
        $resumeInputLedgerPath
    ) | ForEach-Object {
        $item = Get-Item -LiteralPath $_
        [pscustomobject]@{
            portable_path = ConvertTo-GssResumePortablePath -Path $item.FullName -FolderPath $resumeInputTestRoot
            byte_size = [long]$item.Length
            sha256 = Get-GssResumeHash -Path $item.FullName
        }
    }
    $resumeInputManifestRecords += [pscustomobject]@{
        portable_path = 'release/fixture.zip'
        byte_size = 1
        sha256 = ('a' * 64)
    }
    $resumeInputCommittedManifest = [pscustomobject]@{ files = $resumeInputManifestRecords }
    $resumeInputRun = [pscustomobject]@{
        TargetWorkbook = $resumeInputTargetWorkbookPath
        EmailComparisonPdf = $resumeInputComparisonPdfPath
        CurrentSourceWorkbook = $resumeInputCurrentSourcePath
        PriorYearSourceWorkbook = $resumeInputPriorYearSourcePath
    }

    $resumeInputEvidence = Assert-GssResumePackageInputSet `
        -FolderPath $resumeInputTestRoot `
        -LiveRunLogPath $resumeInputLogPath `
        -Run $resumeInputRun `
        -CommittedManifest $resumeInputCommittedManifest
    Assert-Equal $resumeInputEvidence.InputCount 8 'Package-input fixture binds four fixed sources, log, current detail, archive detail, and ledger'
    Assert-Equal $resumeInputEvidence.Inputs.Count 8 'Package-input evidence returns every validated source'
    Assert-Equal (($resumeInputEvidence.Inputs[0].PSObject.Properties.Name | Sort-Object) -join ',') 'ByteSize,PortablePath,Sha256' 'Package-input evidence exposes only portable path, size, and hash'
    Assert-Equal @($resumeInputEvidence.Inputs | Where-Object { $_.PSObject.Properties.Name -contains 'FullPath' }).Count 0 'Package-input evidence does not disclose full local paths'
    Assert-Equal $resumeInputEvidence.DetailWorkbookCount 2 'Package-input fixture binds every enumerated detail workbook'
    Assert-Equal $resumeInputEvidence.FeedbackLedgerIncluded $true 'Package-input fixture binds ledger presence and bytes'
    Assert-Equal $resumeInputEvidence.HistoricalRecoveryMetadataValidation 'SeparateCommittedRecoveryOnlySnapshots' 'Historical metadata is validated through its separate RecoveryOnly contract'

    $resumeInputExtraDetailPath = Join-Path $resumeInputDetailDirectory 'extra-detail.xlsx'
    [System.IO.File]::WriteAllText($resumeInputExtraDetailPath, 'extra-detail-bytes', $utf8WithoutBom)
    Assert-ThrowsLike {
        Assert-GssResumePackageInputSet `
            -FolderPath $resumeInputTestRoot `
            -LiveRunLogPath $resumeInputLogPath `
            -Run $resumeInputRun `
            -CommittedManifest $resumeInputCommittedManifest
    } '*package input is absent from the committed snapshot*extra-detail.xlsx*' 'Package-input fixture rejects an added detail workbook'
    Remove-Item -LiteralPath $resumeInputExtraDetailPath -Force

    $resumeInputRemovedDetailPath = "$resumeInputArchiveDetailPath.removed"
    Move-Item -LiteralPath $resumeInputArchiveDetailPath -Destination $resumeInputRemovedDetailPath
    try {
        Assert-ThrowsLike {
            Assert-GssResumePackageInputSet `
                -FolderPath $resumeInputTestRoot `
                -LiveRunLogPath $resumeInputLogPath `
                -Run $resumeInputRun `
                -CommittedManifest $resumeInputCommittedManifest
        } '*package input is missing from the live source set*archive-detail.xlsx*' 'Package-input fixture rejects a deleted detail workbook'
    }
    finally {
        Move-Item -LiteralPath $resumeInputRemovedDetailPath -Destination $resumeInputArchiveDetailPath
    }

    [System.IO.File]::WriteAllText($resumeInputLedgerPath, '{"changed":true}', $utf8WithoutBom)
    Assert-ThrowsLike {
        Assert-GssResumePackageInputSet `
            -FolderPath $resumeInputTestRoot `
            -LiveRunLogPath $resumeInputLogPath `
            -Run $resumeInputRun `
            -CommittedManifest $resumeInputCommittedManifest
    } '*package input changed after the committed snapshot*gss_feedback_first_seen.json*' 'Package-input fixture rejects changed ledger bytes'
    [System.IO.File]::WriteAllText($resumeInputLedgerPath, $resumeInputLedgerText, $utf8WithoutBom)

    [System.IO.File]::WriteAllText($resumeInputLogPath, '{"changed":true}', $utf8WithoutBom)
    Assert-ThrowsLike {
        Assert-GssResumePackageInputSet `
            -FolderPath $resumeInputTestRoot `
            -LiveRunLogPath $resumeInputLogPath `
            -Run $resumeInputRun `
            -CommittedManifest $resumeInputCommittedManifest
    } '*package input changed after the committed snapshot*gss_update_fixture_apply.json*' 'Package-input fixture rejects changed live run-log bytes'
    [System.IO.File]::WriteAllText($resumeInputLogPath, $resumeInputLogText, $utf8WithoutBom)

    [System.IO.File]::WriteAllText($resumeInputComparisonPdfPath, 'changed-comparison-pdf-bytes', $utf8WithoutBom)
    Assert-ThrowsLike {
        Assert-GssResumePackageInputSet `
            -FolderPath $resumeInputTestRoot `
            -LiveRunLogPath $resumeInputLogPath `
            -Run $resumeInputRun `
            -CommittedManifest $resumeInputCommittedManifest
    } '*package input changed after the committed snapshot*comparison.pdf*' 'Package-input fixture rejects changed promoted comparison-PDF bytes'
    [System.IO.File]::WriteAllText($resumeInputComparisonPdfPath, 'comparison-pdf-bytes', $utf8WithoutBom)

    [System.IO.File]::WriteAllText($resumeInputHistoricalReceiptPath, '{"state":"changed-verification-only"}', $utf8WithoutBom)
    $resumeInputEvidenceAfterMetadataChange = Assert-GssResumePackageInputSet `
        -FolderPath $resumeInputTestRoot `
        -LiveRunLogPath $resumeInputLogPath `
        -Run $resumeInputRun `
        -CommittedManifest $resumeInputCommittedManifest
    Assert-Equal $resumeInputEvidenceAfterMetadataChange.SourceSetSha256 $resumeInputEvidence.SourceSetSha256 'Workbook snapshot input set excludes separately attested recovery metadata'
}
finally {
    if (Test-Path -LiteralPath $resumeInputTestRoot -PathType Container) {
        Remove-Item -LiteralPath $resumeInputTestRoot -Recurse -Force
    }
}

$resumeRecoveryTestRoot = Join-Path ([System.IO.Path]::GetTempPath()) ('gss-resume-recovery-' + [guid]::NewGuid().ToString('N'))
$resumeRecoveryTransactionPath = Join-Path $resumeRecoveryTestRoot '_automation_runs\historical-recovery\fixture-transaction'
$originalRecoveryRootContext = (Get-Item -LiteralPath Function:Get-GssDriveBackupRootContext).ScriptBlock
$originalRecoverySnapshotFinder = (Get-Item -LiteralPath Function:Find-GssDriveBackupSnapshot).ScriptBlock
$originalRecoverySnapshotValidator = (Get-Item -LiteralPath Function:Test-GssCommittedBackupSnapshot).ScriptBlock
New-Item -ItemType Directory -Path $resumeRecoveryTransactionPath -Force | Out-Null
try {
    $resumeRecoveryRunId = 'gss-recovery-fixture-fy26'
    $resumeRecoveryFingerprint = 'sha256:' + ('7' * 64)
    $resumeRecoverySummaryPath = Join-Path $resumeRecoveryTransactionPath 'drive-recovery-summary.json'
    $resumeRecoveryManifestPath = Join-Path $resumeRecoveryTransactionPath 'recovery-manifest.json'
    $resumeRecoveryReceiptPath = Join-Path $resumeRecoveryTransactionPath 'transaction-receipt.json'
    $resumeRecoverySummary = [pscustomobject]@{
        schema_version = 'gss-recovery-drive-summary/v1'
        RunId = $resumeRecoveryRunId
        RunFingerprint = $resumeRecoveryFingerprint
        SnapshotPurpose = 'RecoveryOnly'
    }
    [System.IO.File]::WriteAllText(
        $resumeRecoverySummaryPath,
        ($resumeRecoverySummary | ConvertTo-Json -Compress),
        $utf8WithoutBom
    )
    [System.IO.File]::WriteAllText($resumeRecoveryManifestPath, '{"recovery":"manifest"}', $utf8WithoutBom)
    [System.IO.File]::WriteAllText($resumeRecoveryReceiptPath, '{"state":"Committed"}', $utf8WithoutBom)
    foreach ($terminalState in @('RolledBack', 'RolledBackConservative')) {
        $terminalManifest = [pscustomobject][ordered]@{
            schema_version = 'gss-historical-recovery/v1'
            fiscal_year = if ($terminalState -ceq 'RolledBack') { 'FY24' } else { 'FY25' }
            sources = @()
        }
        $terminalManifestText = $terminalManifest | ConvertTo-Json -Depth 4 -Compress
        $terminalManifestSha256 = Get-GssResumeTextHash -Text $terminalManifestText
        $terminalDirectory = Join-Path `
            (Split-Path -Parent $resumeRecoveryTransactionPath) `
            $terminalManifestSha256
        New-Item -ItemType Directory -Path $terminalDirectory -Force | Out-Null
        $terminalManifestPath = Join-Path $terminalDirectory 'recovery-manifest.json'
        [System.IO.File]::WriteAllText($terminalManifestPath, $terminalManifestText, $utf8WithoutBom)
        $terminalReceipt = [pscustomobject][ordered]@{
            schema_version = 'gss-historical-recovery-receipt/v1'
            classification = 'CONTAINS PERSONAL DATA - RESTRICTED'
            contains_personal_data = $true
            manifest_sha256 = $terminalManifestSha256
            manifest_snapshot_path = $terminalManifestPath
            transaction_id = "historical-recovery:$terminalManifestSha256"
            state = $terminalState
            files = @()
        }
        [System.IO.File]::WriteAllText(
            (Join-Path $terminalDirectory 'transaction-receipt.json'),
            ($terminalReceipt | ConvertTo-Json -Depth 4 -Compress),
            $utf8WithoutBom
        )
    }

    $resumeRecoveryRoleBindings = @(
        [pscustomobject]@{ Role = 'recovery_run_summary'; PortablePath = 'recovery/evidence/drive-recovery-summary.json'; LivePath = $resumeRecoverySummaryPath },
        [pscustomobject]@{ Role = 'recovery_manifest'; PortablePath = 'recovery/evidence/recovery-manifest.json'; LivePath = $resumeRecoveryManifestPath },
        [pscustomobject]@{ Role = 'recovery_receipt'; PortablePath = 'recovery/evidence/transaction-receipt.json'; LivePath = $resumeRecoveryReceiptPath }
    )
    $resumeRecoveryFiles = @($resumeRecoveryRoleBindings | ForEach-Object {
        $item = Get-Item -LiteralPath $_.LivePath
        [pscustomobject]@{
            role = $_.Role
            portable_path = $_.PortablePath
            byte_size = [long]$item.Length
            sha256 = Get-GssResumeHash -Path $item.FullName
        }
    })
    $resumeRecoveryDrive = [pscustomobject]@{ verification_level = 'drivefs_hash_verified' }
    $resumeRecoveryCommitReceipt = [pscustomobject]@{
        status = 'Committed'
        run_id = $resumeRecoveryRunId
        fingerprint = $resumeRecoveryFingerprint
        snapshot_purpose = 'RecoveryOnly'
        verification_level = 'drivefs_hash_verified'
    }
    $resumeRecoveryManifest = [pscustomobject]@{
        status = 'Committed'
        run_id = $resumeRecoveryRunId
        fingerprint = $resumeRecoveryFingerprint
        snapshot_purpose = 'RecoveryOnly'
        drive = $resumeRecoveryDrive
        files = $resumeRecoveryFiles
    }
    $resumeRecoveryPreparedManifest = [pscustomobject]@{
        status = 'Prepared'
        run_id = $resumeRecoveryRunId
        fingerprint = $resumeRecoveryFingerprint
        snapshot_purpose = 'RecoveryOnly'
        drive = $resumeRecoveryDrive
    }
    $script:resumeRecoveryFixture = [pscustomobject]@{
        RootPath = Join-Path $resumeRecoveryTestRoot 'drive-root'
        RunId = $resumeRecoveryRunId
        SnapshotPath = Join-Path $resumeRecoveryTestRoot 'drive-root\committed-snapshot'
        Validation = [pscustomobject]@{
            Receipt = $resumeRecoveryCommitReceipt
            Manifest = $resumeRecoveryManifest
            PreparedManifest = $resumeRecoveryPreparedManifest
        }
    }

    Set-Item -LiteralPath Function:Get-GssDriveBackupRootContext -Value {
        [pscustomobject]@{ RootPath = $script:resumeRecoveryFixture.RootPath }
    }
    Set-Item -LiteralPath Function:Find-GssDriveBackupSnapshot -Value {
        param([string]$RootPath, [string]$RunId)
        if (-not (Test-GssResumeSamePath -Left $RootPath -Right $script:resumeRecoveryFixture.RootPath) -or
            $RunId -cne $script:resumeRecoveryFixture.RunId) {
            throw 'Recovery snapshot finder received unexpected fixture identity.'
        }
        [pscustomobject]@{ Path = $script:resumeRecoveryFixture.SnapshotPath; IsPartial = $false }
    }
    Set-Item -LiteralPath Function:Test-GssCommittedBackupSnapshot -Value {
        param([string]$SnapshotPath)
        if (-not (Test-GssResumeSamePath -Left $SnapshotPath -Right $script:resumeRecoveryFixture.SnapshotPath)) {
            throw 'Recovery snapshot validator received an unexpected fixture path.'
        }
        $script:resumeRecoveryFixture.Validation
    }

    $resumeRecoveryEvidence = Assert-GssResumeHistoricalRecoveryEvidence -FolderPath $resumeRecoveryTestRoot
    Assert-Equal $resumeRecoveryEvidence.TransactionCount 1 'Historical recovery fixture ignores terminal rollback receipts and validates one consulted transaction'
    Assert-Equal $resumeRecoveryEvidence.EvidenceFileCount 3 'Historical recovery fixture binds summary, manifest, and receipt'
    Assert-Equal $resumeRecoveryEvidence.Verification 'CommittedRecoveryOnlySnapshots' 'Historical recovery fixture requires committed RecoveryOnly evidence'

    $forgedTerminalDirectory = Join-Path (Split-Path -Parent $resumeRecoveryTransactionPath) ('f' * 64)
    New-Item -ItemType Directory -Path $forgedTerminalDirectory -Force | Out-Null
    [System.IO.File]::WriteAllText(
        (Join-Path $forgedTerminalDirectory 'transaction-receipt.json'),
        '{"state":"RolledBack"}',
        $utf8WithoutBom
    )
    Assert-ThrowsLike {
        Assert-GssResumeHistoricalRecoveryEvidence -FolderPath $resumeRecoveryTestRoot
    } '*Terminal historical recovery receipt identity is invalid*' 'Historical recovery fixture rejects an unbound state-only terminal receipt'
    Remove-Item -LiteralPath $forgedTerminalDirectory -Recurse -Force

    [System.IO.File]::WriteAllText($resumeRecoveryReceiptPath, '{"state":"Changed"}', $utf8WithoutBom)
    Assert-ThrowsLike {
        Assert-GssResumeHistoricalRecoveryEvidence -FolderPath $resumeRecoveryTestRoot
    } '*recovery evidence changed after its committed RecoveryOnly snapshot*transaction-receipt.json*' 'Historical recovery fixture rejects changed receipt bytes'
    [System.IO.File]::WriteAllText($resumeRecoveryReceiptPath, '{"state":"Committed"}', $utf8WithoutBom)

    $resumeRecoveryCommitReceipt.verification_level = 'unverified'
    Assert-ThrowsLike {
        Assert-GssResumeHistoricalRecoveryEvidence -FolderPath $resumeRecoveryTestRoot
    } '*recovery snapshot is not committed*' 'Historical recovery fixture requires DriveFS hash verification'
}
finally {
    Set-Item -LiteralPath Function:Get-GssDriveBackupRootContext -Value $originalRecoveryRootContext
    Set-Item -LiteralPath Function:Find-GssDriveBackupSnapshot -Value $originalRecoverySnapshotFinder
    Set-Item -LiteralPath Function:Test-GssCommittedBackupSnapshot -Value $originalRecoverySnapshotValidator
    Remove-Variable -Name resumeRecoveryFixture -Scope Script -ErrorAction SilentlyContinue
    if (Test-Path -LiteralPath $resumeRecoveryTestRoot -PathType Container) {
        Remove-Item -LiteralPath $resumeRecoveryTestRoot -Recurse -Force
    }
}
Assert-True ($safeLauncherSource.Contains("if (`$finalizeStatus -eq 'Blocked')")) 'Safe launcher handles Drive blocked separately from pending finalize'
Assert-True ($safeLauncherSource.Contains("'BackupBlocked'")) 'Safe launcher records non-retryable Drive conflict for manual review'
Assert-True ($safeLauncherSource.Contains('-Operation Abort')) 'Prepared Drive snapshot is explicitly aborted after safe rollback or pre-commit failure'

$quarterlyDrillSource = Get-Content -LiteralPath (Join-Path $scriptRoot 'Invoke-GSS-QuarterlyRestoreDrill.ps1') -Raw
$quarterlyTryIndex = $quarterlyDrillSource.IndexOf('try {')
$quarterlyResolverIndex = $quarterlyDrillSource.IndexOf('$restoredWorkbookMapping = Resolve-GssDriveBackupRestoredFile')
$quarterlyFinallyIndex = $quarterlyDrillSource.IndexOf('finally {')
$quarterlyReceiptIndex = $quarterlyDrillSource.IndexOf('Write-GssDriveBackupAtomicJson -Path $drillReceipt')
Assert-True (
    $quarterlyTryIndex -ge 0 -and
    $quarterlyResolverIndex -gt $quarterlyTryIndex -and
    $quarterlyResolverIndex -lt $quarterlyFinallyIndex
) 'Quarterly drill resolves restored workbook mapping inside the receipt-producing guard'
Assert-True (
    $quarterlyFinallyIndex -gt $quarterlyResolverIndex -and
    $quarterlyReceiptIndex -gt $quarterlyFinallyIndex
) 'Quarterly drill writes combined failure evidence from its finally block'
Assert-True (
    $quarterlyDrillSource.Contains('$restoredWorkbookMapping = $null') -and
    $quarterlyDrillSource.Contains('-not [string]::IsNullOrWhiteSpace($restoredWorkbook)')
) 'Quarterly drill safely records missing-workbook mapping failures'

$portableRoot = 'C:\Users\bryan\Dropbox\Automations\GSS Surveys'
$otherProfilePath = "C:\Users\Other User\Dropbox\Automations\GSS Surveys\04 Email Comparison PDFs\report.pdf"
Assert-Equal (ConvertTo-GssDropboxRelativePath -Path $otherProfilePath -FolderPath $portableRoot) '04 Email Comparison PDFs/report.pdf' 'Cross-profile portable path recovery'
Assert-ThrowsLike { ConvertTo-GssDropboxRelativePath -Path '..\escape.txt' -FolderPath $portableRoot } '*traversal segment*' 'Portable path traversal rejection'

Write-Host 'GSS non-Excel logic tests passed.'
