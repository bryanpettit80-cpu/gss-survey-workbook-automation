[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'

$scriptRoot = Join-Path (Split-Path -Parent (Split-Path -Parent $PSCommandPath)) 'scripts'
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
        SemanticSha256 = ('a' * 64)
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
        $expectedTemporaryRoot = [IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd('\') + '\'
        if (-not [IO.Path]::GetFullPath($atomicTestRoot).StartsWith($expectedTemporaryRoot, [StringComparison]::OrdinalIgnoreCase) -or
            [IO.Path]::GetFileName($atomicTestRoot) -notmatch '^gss-atomic-replace-[a-f0-9]{32}$') { throw 'Unexpected atomic-test cleanup target.' }
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
    (New-TestSource 'a-duplicate.xlsx' ([datetime]'2026-06-28') ([datetime]'2026-07-01T08:00:00')),
    (New-TestSource 'z-duplicate.xlsx' ([datetime]'2026-06-28') ([datetime]'2026-07-01T09:00:00')),
    (New-TestSource 'prior-week.xlsx' ([datetime]'2026-06-21') ([datetime]'2026-07-02T09:00:00'))
)

$latest = Select-LatestRollingSource $sources
Assert-Equal $latest.File.FullName 'a-duplicate.xlsx' 'Equal same-week content uses stable path order and ignores last write time'

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

# Runtime behavior is covered by Gss.Operations.Tests.ps1 and
# Gss.Transactions.Tests.ps1. The retired source-text assertions froze the old
# Git/Drive/APPLY approval chain and its release-pair allowlist; they were not
# business-correctness checks. Source-content regression cases live in
# Gss.Reporting.Tests.ps1; legacy recovery/import tests remain separate.

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

$portableRoot = 'C:\Users\Operator\Dropbox\Automations\GSS Surveys'
$otherProfilePath = "C:\Users\Other User\Dropbox\Automations\GSS Surveys\04 Email Comparison PDFs\report.pdf"
Assert-Equal (ConvertTo-GssDropboxRelativePath -Path $otherProfilePath -FolderPath $portableRoot) '04 Email Comparison PDFs/report.pdf' 'Cross-profile portable path recovery'
Assert-ThrowsLike { ConvertTo-GssDropboxRelativePath -Path '..\escape.txt' -FolderPath $portableRoot } '*traversal segment*' 'Portable path traversal rejection'

Write-Host 'GSS non-Excel logic tests passed.'
