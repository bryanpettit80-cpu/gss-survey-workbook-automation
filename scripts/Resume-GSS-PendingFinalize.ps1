[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$LiveRunLogPath,
    [string]$Folder,
    [switch]$OutputObject
)

$ErrorActionPreference = 'Stop'

$scriptRoot = Split-Path -Parent $PSCommandPath
$repoRoot = Split-Path -Parent $scriptRoot
$releaseValidator = Join-Path $scriptRoot 'Test-GSS-ReleaseIntegrity.ps1'
$releaseIntegrity = & $releaseValidator -RepoRoot $repoRoot

function Write-GssResumeAtomicJson {
    param([Parameter(Mandatory)][string]$Path, [Parameter(Mandatory)][object]$InputObject)

    $directory = Split-Path -Parent $Path
    New-Item -ItemType Directory -Path $directory -Force | Out-Null
    $temporaryPath = Join-Path $directory ('.t-{0}' -f ([guid]::NewGuid().ToString('N').Substring(0, 8)))
    try {
        [System.IO.File]::WriteAllText(
            $temporaryPath,
            ($InputObject | ConvertTo-Json -Depth 10),
            (New-Object System.Text.UTF8Encoding($false))
        )
        if (Test-Path -LiteralPath $Path -PathType Leaf) {
            [System.IO.File]::Replace($temporaryPath, $Path, $null, $true)
        }
        else {
            [System.IO.File]::Move($temporaryPath, $Path)
        }
    }
    finally {
        if (Test-Path -LiteralPath $temporaryPath -PathType Leaf) {
            Remove-Item -LiteralPath $temporaryPath -Force
        }
    }
}

function Get-GssResumeHash {
    param([Parameter(Mandatory)][string]$Path)
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        throw "Required run evidence is missing: $Path"
    }
    return (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash.ToLowerInvariant()
}

function Get-GssResumeTextHash {
    param([Parameter(Mandatory)][AllowEmptyString()][string]$Text)

    $sha = [System.Security.Cryptography.SHA256]::Create()
    try {
        $bytes = (New-Object System.Text.UTF8Encoding($false)).GetBytes($Text)
        return ([System.BitConverter]::ToString($sha.ComputeHash($bytes))).Replace('-', '').ToLowerInvariant()
    }
    finally {
        $sha.Dispose()
    }
}

function Get-GssResumeRunFingerprint {
    param([Parameter(Mandatory)][object]$Run)

    return Get-GssResumeTextHash (@(
        'gss-transaction-v1',
        ([string]$Run.RunId).Trim().ToLowerInvariant(),
        ([string]$Run.HostName).Trim().ToLowerInvariant(),
        ([string]$Run.CurrentWeekEnding).Trim(),
        ([string]$Run.StartingWorkbookSha256).Trim().ToLowerInvariant(),
        ([string]$Run.CurrentSourceSha256).Trim().ToLowerInvariant(),
        ([string]$Run.PriorYearSourceSha256).Trim().ToLowerInvariant(),
        ([string]$Run.StagedWorkbookSha256).Trim().ToLowerInvariant(),
        ([string]$Run.StagedPdfSha256).Trim().ToLowerInvariant(),
        ([string]$Run.ProgramRelease).Trim().ToLowerInvariant()
    ) -join "`n")
}

function Add-GssResumeReceiptValue {
    param([Parameter(Mandatory)][object]$Receipt, [Parameter(Mandatory)][string]$Name, [object]$Value)
    if ($Receipt.PSObject.Properties.Name -contains $Name) {
        $Receipt.$Name = $Value
    }
    else {
        $Receipt | Add-Member -NotePropertyName $Name -NotePropertyValue $Value
    }
}

$LiveRunLogPath = (Resolve-Path -LiteralPath $LiveRunLogPath).Path
$run = Get-Content -LiteralPath $LiveRunLogPath -Raw | ConvertFrom-Json
if ([string]$run.Mode -ne 'ApplyToMainWorkbook' -or [string]$run.TransactionStatus -ne 'Committed') {
    throw 'ResumeFinalize requires a committed live workbook run log.'
}
if ([string]::IsNullOrWhiteSpace([string]$run.RunId) -or [string]::IsNullOrWhiteSpace([string]$run.RunFingerprint)) {
    throw 'Live run log is missing RunId or RunFingerprint.'
}
if ([string]$run.HostName -cne [Environment]::MachineName) {
    throw "Live run belongs to host '$($run.HostName)', not this workstation."
}
if ((Get-GssResumeRunFingerprint $run) -ne ([string]$run.RunFingerprint).ToLowerInvariant()) {
    throw 'Live run fingerprint no longer matches its immutable transaction evidence.'
}
if ([string]$run.ProgramRelease -ne [string]$releaseIntegrity.ReleaseTag) {
    throw "Live run release '$($run.ProgramRelease)' does not match current approved release '$($releaseIntegrity.ReleaseTag)'."
}

if ([string]::IsNullOrWhiteSpace($Folder)) {
    $Folder = [string]$run.Folder
}
$Folder = (Resolve-Path -LiteralPath $Folder).Path
$allowedLogRoot = [System.IO.Path]::GetFullPath((Join-Path $Folder '_automation_runs\logs')).TrimEnd('\') + '\'
if (-not $LiveRunLogPath.StartsWith($allowedLogRoot, [System.StringComparison]::OrdinalIgnoreCase)) {
    throw 'Live run log is outside the configured GSS automation log directory.'
}
$DriveBackupScript = Join-Path $scriptRoot 'Invoke-GSS-DriveBackup.ps1'
if (-not (Test-Path -LiteralPath $DriveBackupScript -PathType Leaf)) {
    throw "Drive backup coordinator is unavailable: $DriveBackupScript"
}

foreach ($check in @(
    [pscustomobject]@{ Name = 'live workbook'; Path = [string]$run.TargetWorkbook; Expected = [string]$run.PromotedWorkbookSha256 },
    [pscustomobject]@{ Name = 'staged workbook'; Path = [string]$run.StagedWorkbook; Expected = [string]$run.StagedWorkbookSha256 },
    [pscustomobject]@{ Name = 'staged PDF'; Path = [string]$run.StagedPdf; Expected = [string]$run.StagedPdfSha256 },
    [pscustomobject]@{ Name = 'current source'; Path = [string]$run.CurrentSourceWorkbook; Expected = [string]$run.CurrentSourceSha256 },
    [pscustomobject]@{ Name = 'prior-year source'; Path = [string]$run.PriorYearSourceWorkbook; Expected = [string]$run.PriorYearSourceSha256 }
)) {
    if ((Get-GssResumeHash $check.Path) -ne $check.Expected.ToLowerInvariant()) {
        throw "ResumeFinalize blocked because the $($check.Name) no longer matches the committed run."
    }
}

$analyzer = Join-Path $scriptRoot 'Analyze-GSS-Run.ps1'
$stateDirectory = Join-Path $Folder '_automation_runs\state'
$activeRunPath = Join-Path $stateDirectory 'active-transaction.json'
$receiptPath = Join-Path $stateDirectory "transaction-$($run.RunId).json"
$receipt = if (Test-Path -LiteralPath $receiptPath -PathType Leaf) {
    Get-Content -LiteralPath $receiptPath -Raw | ConvertFrom-Json
}
else {
    [pscustomobject]@{
        ReceiptSchemaVersion = 1
        RunId = [string]$run.RunId
        HostName = [Environment]::MachineName
        ProgramRelease = [string]$run.ProgramRelease
        LiveRunLogPath = $LiveRunLogPath
        AutomaticSending = 'Disabled'
    }
}

if (Test-Path -LiteralPath $activeRunPath -PathType Leaf) {
    $activePreflight = Get-Content -LiteralPath $activeRunPath -Raw | ConvertFrom-Json
    if ([string]$activePreflight.RunId -ne [string]$run.RunId) {
        throw "A different unresolved GSS transaction is active: $($activePreflight.RunId)"
    }
    if ([string]$activePreflight.TransactionStatus -in @('RollbackAttempting', 'RollbackBlocked')) {
        throw "Transaction '$($run.RunId)' requires rollback review and cannot be resumed as a Drive finalization."
    }
}

$mutex = New-Object System.Threading.Mutex($false, 'Global\GSSSurveyWorkbookAutomationTransaction')
$ownsMutex = $false
$ownsActiveMarker = $false
$driveCommitted = $false
$driveBlocked = $false
try {
    try {
        $ownsMutex = $mutex.WaitOne(0)
    }
    catch [System.Threading.AbandonedMutexException] {
        $ownsMutex = $true
    }
    if (-not $ownsMutex) {
        throw 'Another GSS workbook transaction is already active on this workstation.'
    }

    if (Test-Path -LiteralPath $activeRunPath -PathType Leaf) {
        $active = Get-Content -LiteralPath $activeRunPath -Raw | ConvertFrom-Json
        if ([string]$active.RunId -ne [string]$run.RunId) {
            throw "A different unresolved GSS transaction is active: $($active.RunId)"
        }
    }
    Add-GssResumeReceiptValue $receipt 'ReceiptPath' $receiptPath
    Add-GssResumeReceiptValue $receipt 'TransactionStatus' 'RetryingFinalize'
    Add-GssResumeReceiptValue $receipt 'UpdatedUtc' ([datetime]::UtcNow.ToString('o'))
    Write-GssResumeAtomicJson -Path $activeRunPath -InputObject $receipt
    $ownsActiveMarker = $true
    Write-GssResumeAtomicJson -Path $receiptPath -InputObject $receipt

    $finalize = & $DriveBackupScript -Operation RetryFinalize -RunSummaryPath $LiveRunLogPath -OutputObject
    if ([string]$finalize.Status -eq 'Blocked') {
        $driveBlocked = $true
        throw 'Drive retry is blocked by a non-retryable backup-chain or identity conflict. Manual review is required.'
    }
    if ([string]$finalize.Status -ne 'Committed' -or
        [string]$finalize.RunId -ne [string]$run.RunId -or
        [string]$finalize.Fingerprint -ne [string]$run.RunFingerprint) {
        throw "Drive retry did not commit the exact run (status: '$($finalize.Status)')."
    }
    $driveCommitted = $true
    Add-GssResumeReceiptValue $receipt 'BackupStatus' 'Committed'
    Add-GssResumeReceiptValue $receipt 'BackupFinalize' $finalize
    Add-GssResumeReceiptValue $receipt 'TransactionStatus' 'BackupCommitted'
    Add-GssResumeReceiptValue $receipt 'UpdatedUtc' ([datetime]::UtcNow.ToString('o'))
    Write-GssResumeAtomicJson -Path $activeRunPath -InputObject $receipt
    Write-GssResumeAtomicJson -Path $receiptPath -InputObject $receipt

    $review = & $analyzer -Folder $Folder -LogPath $LiveRunLogPath -OutputObject
    if ($review.WorkbookStatus -eq 'Blocked') {
        throw 'The committed live workbook no longer passes final review; package publication remains blocked.'
    }
    $published = & $analyzer -Folder $Folder -LogPath $LiveRunLogPath -OutputObject -PublishEmailPackage
    if ($published.EmailReadiness -ne 'Ready' -or -not $published.EmailPackage) {
        Add-GssResumeReceiptValue $receipt 'TransactionStatus' 'PackageBlocked'
        Add-GssResumeReceiptValue $receipt 'PackagePublished' $false
        Add-GssResumeReceiptValue $receipt 'Error' (@($published.Qa.EmailBlockers) -join '; ')
        Add-GssResumeReceiptValue $receipt 'UpdatedUtc' ([datetime]::UtcNow.ToString('o'))
        Write-GssResumeAtomicJson -Path $activeRunPath -InputObject $receipt
        Write-GssResumeAtomicJson -Path $receiptPath -InputObject $receipt
        throw 'Drive is committed, but email package publication remains blocked.'
    }

    Add-GssResumeReceiptValue $receipt 'TransactionStatus' 'Committed'
    Add-GssResumeReceiptValue $receipt 'PackagePublished' $true
    Add-GssResumeReceiptValue $receipt 'PackagePath' ([string]$published.EmailPackage.PackagePath)
    Add-GssResumeReceiptValue $receipt 'Error' $null
    Add-GssResumeReceiptValue $receipt 'CompletedUtc' ([datetime]::UtcNow.ToString('o'))
    Add-GssResumeReceiptValue $receipt 'UpdatedUtc' ([datetime]::UtcNow.ToString('o'))
    Write-GssResumeAtomicJson -Path $activeRunPath -InputObject $receipt
    Write-GssResumeAtomicJson -Path $receiptPath -InputObject $receipt

    Write-Information 'GSS Drive finalization and package publication completed.' -InformationAction Continue
    Write-Information "  Run: $($run.RunId)" -InformationAction Continue
    Write-Information "  Package: $($published.EmailPackage.PackagePath)" -InformationAction Continue
    Write-Information '  Automatic sending: disabled' -InformationAction Continue
    if ($OutputObject) { return $receipt }
}
catch {
    Add-GssResumeReceiptValue $receipt 'TransactionStatus' $(if ($driveCommitted) { 'PackageBlocked' } elseif ($driveBlocked) { 'BackupBlocked' } else { 'PendingFinalize' })
    Add-GssResumeReceiptValue $receipt 'BackupStatus' $(if ($driveCommitted) { 'Committed' } elseif ($driveBlocked) { 'Blocked' } else { 'PendingFinalize' })
    Add-GssResumeReceiptValue $receipt 'Error' $_.Exception.Message
    Add-GssResumeReceiptValue $receipt 'UpdatedUtc' ([datetime]::UtcNow.ToString('o'))
    try {
        if ($ownsActiveMarker) {
            Write-GssResumeAtomicJson -Path $activeRunPath -InputObject $receipt
        }
        Write-GssResumeAtomicJson -Path $receiptPath -InputObject $receipt
    }
    catch { Write-Verbose "The resume transaction receipt could not be updated: $($_.Exception.Message)" }
    throw
}
finally {
    if ($ownsActiveMarker -and
        $receipt.TransactionStatus -in @('Committed', 'PackageBlocked') -and
        (Test-Path -LiteralPath $activeRunPath -PathType Leaf)) {
        try {
            $active = Get-Content -LiteralPath $activeRunPath -Raw | ConvertFrom-Json
            if ([string]$active.RunId -eq [string]$run.RunId) {
                Remove-Item -LiteralPath $activeRunPath -Force
            }
        }
        catch { Write-Verbose "The resumed active transaction marker could not be removed: $($_.Exception.Message)" }
    }
    if ($ownsMutex) {
        try { $mutex.ReleaseMutex() }
        catch { Write-Verbose "The GSS transaction mutex could not be released cleanly: $($_.Exception.Message)" }
    }
    $mutex.Dispose()
}
