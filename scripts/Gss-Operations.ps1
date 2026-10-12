# Shared, local reporting operations. Release certification is an installation
# concern; these helpers do not require Git or historical test workbooks.
function Get-GssRuntimeInfo {
    param([Parameter(Mandatory)][string]$RepoRoot)
    $manifestPath = Join-Path $RepoRoot 'release\release-manifest.json'
    $version = 'development'
    if (Test-Path -LiteralPath $manifestPath -PathType Leaf) {
        $manifest = Get-Content -LiteralPath $manifestPath -Raw | ConvertFrom-Json
        if (-not [string]::IsNullOrWhiteSpace([string]$manifest.release_tag)) { $version = [string]$manifest.release_tag }
    }
    [pscustomobject]@{ ReleaseTag = $version; HeadCommit = $null; Status = 'Runtime'; TransactionSchemaVersion = 2 }
}

function Write-GssOperationJson {
    param([Parameter(Mandatory)][string]$Path, [Parameter(Mandatory)][object]$Value)
    $directory = Split-Path -Parent $Path
    [void][IO.Directory]::CreateDirectory($directory)
    $temporary = Join-Path $directory ('.json-' + [guid]::NewGuid().ToString('N'))
    $previous = Join-Path $directory ('.previous-' + [guid]::NewGuid().ToString('N'))
    try {
        [IO.File]::WriteAllText($temporary, ($Value | ConvertTo-Json -Depth 30), (New-Object Text.UTF8Encoding($false)))
        if (Test-Path -LiteralPath $Path -PathType Leaf) { [IO.File]::Replace($temporary, $Path, $previous, $true) }
        else { [IO.File]::Move($temporary, $Path) }
    }
    finally {
        if (Test-Path -LiteralPath $temporary) { Remove-Item -LiteralPath $temporary -Force }
        if (Test-Path -LiteralPath $previous) { Remove-Item -LiteralPath $previous -Force }
    }
}

function Get-GssOperationHash {
    param([Parameter(Mandatory)][string]$Path)
    (Get-FileHash -LiteralPath $Path -Algorithm SHA256 -ErrorAction Stop).Hash.ToLowerInvariant()
}

function Resolve-GssOperationPath {
    param([Parameter(Mandatory)][string]$Root, [Parameter(Mandatory)][string]$RelativePath, [switch]$RequireFile)
    if ([IO.Path]::IsPathRooted($RelativePath) -or $RelativePath -match '(^|[\\/])\.\.([\\/]|$)' -or $RelativePath.Contains(':')) {
        throw "Expected a contained relative GSS path: $RelativePath"
    }
    $rootPath = [IO.Path]::GetFullPath($Root).TrimEnd('\', '/')
    $path = [IO.Path]::GetFullPath((Join-Path $rootPath $RelativePath.Replace('/', '\')))
    if (-not $path.StartsWith($rootPath + '\', [StringComparison]::OrdinalIgnoreCase)) { throw "Path escaped the GSS root: $RelativePath" }
    if ($RequireFile -and -not (Test-Path -LiteralPath $path -PathType Leaf)) { throw "Required GSS file is unavailable: $path" }
    return $path
}

function Get-GssOperationRelativePath {
    param([Parameter(Mandatory)][string]$Root, [Parameter(Mandatory)][string]$Path)
    $rootPath = [IO.Path]::GetFullPath($Root).TrimEnd('\', '/')
    $fullPath = [IO.Path]::GetFullPath($Path)
    if (-not $fullPath.StartsWith($rootPath + '\', [StringComparison]::OrdinalIgnoreCase)) { throw "GSS artifact is outside its data folder: $Path" }
    $fullPath.Substring($rootPath.Length + 1).Replace('\', '/')
}

function ConvertTo-GssPortableRun {
    param([Parameter(Mandatory)][object]$Run, [Parameter(Mandatory)][string]$Folder)
    # Work on a copy. An old log is immutable evidence, never rewritten by resume.
    $copy = $Run | ConvertTo-Json -Depth 30 | ConvertFrom-Json
    if ([int]$copy.ReceiptSchemaVersion -notin @(1, 2)) { throw "Unsupported transaction schema: $($copy.ReceiptSchemaVersion)" }
    $originalRoot = [string]$copy.Folder
    foreach ($binding in @(
        @('TargetWorkbook', 'TargetWorkbookRelativePath'), @('EmailComparisonPdf', 'EmailComparisonPdfRelativePath'), @('StartingWorkbook', 'StartingWorkbookRelativePath'),
        @('CurrentSourceWorkbook', 'CurrentSourceRelativePath'), @('PriorYearSourceWorkbook', 'PriorYearSourceRelativePath'),
        @('BackupWorkbook', 'BackupWorkbookRelativePath'), @('PdfBackup', 'PdfBackupRelativePath'),
        @('StagedWorkbook', 'StagedWorkbookRelativePath'), @('StagedPdf', 'StagedPdfRelativePath'),
        @('LogPath', 'LogRelativePath'), @('PreparedRunLogPath', 'PreparedRunLogRelativePath')
    )) {
        $value = [string]$copy.($binding[0])
        if ([string]::IsNullOrWhiteSpace($value)) { continue }
        $relative = [string]$copy.($binding[1])
        if ([string]::IsNullOrWhiteSpace($relative)) {
            $relative = if ([IO.Path]::IsPathRooted($value)) { Get-GssOperationRelativePath -Root $originalRoot -Path $value } else { $value }
        }
        $copy | Add-Member -NotePropertyName $binding[0] -NotePropertyValue (Resolve-GssOperationPath -Root $Folder -RelativePath $relative) -Force
        $copy | Add-Member -NotePropertyName $binding[1] -NotePropertyValue $relative -Force
    }
    $copy.Folder = [IO.Path]::GetFullPath($Folder)
    return $copy
}

function Get-GssBackupFolder {
    param([string]$BackupFolder)
    if (-not [string]::IsNullOrWhiteSpace($BackupFolder)) { return [IO.Path]::GetFullPath($BackupFolder) }
    $settingsRoot = Join-Path $env:LOCALAPPDATA 'GSSSurveyWorkbookAutomation'
    $runtimePath = Join-Path $settingsRoot 'runtime-settings.json'
    if (Test-Path -LiteralPath $runtimePath -PathType Leaf) {
        $settings = Get-Content -LiteralPath $runtimePath -Raw | ConvertFrom-Json
        if ([int]$settings.schema_version -ne 1) { throw 'Unsupported GSS runtime settings schema.' }
        if (-not [string]::IsNullOrWhiteSpace([string]$settings.backup_folder)) { return [IO.Path]::GetFullPath([string]$settings.backup_folder) }
        return $null
    }
    # Existing installations retain their destination, without historical trust prerequisites.
    $legacyPath = Join-Path $settingsRoot 'settings.json'
    if (Test-Path -LiteralPath $legacyPath -PathType Leaf) {
        $legacy = Get-Content -LiteralPath $legacyPath -Raw | ConvertFrom-Json
        if (-not [string]::IsNullOrWhiteSpace([string]$legacy.drive_root_path)) { return [IO.Path]::GetFullPath([string]$legacy.drive_root_path) }
    }
    return $null
}

function Find-GssExistingReport {
    param([Parameter(Mandatory)][object]$PreparedRun, [Parameter(Mandatory)][string]$Folder)
    if (-not [bool]$PreparedRun.NoBusinessChanges) { return $null }
    if ([string]::IsNullOrWhiteSpace([string]$PreparedRun.ProgramRelease)) { return $null }
    $logs = Join-Path $Folder '_automation_runs\logs'
    if (-not (Test-Path -LiteralPath $logs -PathType Container)) { return $null }
    foreach ($file in Get-ChildItem -LiteralPath $logs -File -Filter '*_apply.json' | Sort-Object LastWriteTimeUtc -Descending) {
        try {
            $original = Get-Content -LiteralPath $file.FullName -Raw | ConvertFrom-Json
            if ([string]$original.TransactionStatus -ne 'Committed' -or
                [string]$original.ProgramRelease -ne [string]$PreparedRun.ProgramRelease -or
                [string]$original.CurrentWeekEnding -ne [string]$PreparedRun.CurrentWeekEnding -or
                [string]$original.PromotedWorkbookSha256 -ne [string]$PreparedRun.StartingWorkbookSha256) { continue }
            $run = ConvertTo-GssPortableRun -Run $original -Folder $Folder
            if (-not (Test-Path -LiteralPath $run.TargetWorkbook -PathType Leaf) -or
                (Get-GssOperationHash $run.TargetWorkbook) -ne [string]$PreparedRun.StartingWorkbookSha256 -or
                -not (Test-Path -LiteralPath $run.EmailComparisonPdf -PathType Leaf) -or
                (Get-GssOperationHash $run.EmailComparisonPdf) -ne [string]$run.PromotedPdfSha256) { continue }
            $run | Add-Member -NotePropertyName LogPath -NotePropertyValue $file.FullName -Force
            return $run
        }
        catch { Write-Verbose "Existing report receipt was not reusable: $($_.Exception.Message)" }
    }
    return $null
}

function ConvertTo-GssRefreshedAnalysisRun {
    param([Parameter(Mandatory)][object]$ReportRun, [Parameter(Mandatory)][object]$PreparedRun, [Parameter(Mandatory)][string]$Folder)
    if ([string]$PreparedRun.Mode -ne 'CopyTestOnly' -or -not [bool]$PreparedRun.NoBusinessChanges -or
        [string]$PreparedRun.CurrentWeekEnding -ne [string]$ReportRun.CurrentWeekEnding -or
        [string]$PreparedRun.ProgramRelease -ne [string]$ReportRun.ProgramRelease -or
        [string]$PreparedRun.StartingWorkbookSha256 -cne [string]$ReportRun.PromotedWorkbookSha256) {
        throw 'The fresh evidence check does not identify the unchanged current report.'
    }
    $copy = ConvertTo-GssPortableRun -Run $ReportRun -Folder $Folder
    $fresh = ConvertTo-GssPortableRun -Run $PreparedRun -Folder $Folder
    foreach ($binding in @(
        @('TargetWorkbook', 'PromotedWorkbookSha256'), @('EmailComparisonPdf', 'PromotedPdfSha256')
    )) {
        if ((Get-GssOperationHash $copy.($binding[0])) -cne [string]$copy.($binding[1])) { throw 'The current report changed after the fresh evidence check.' }
    }
    foreach ($binding in @(@('CurrentSourceWorkbook', 'CurrentSourceSha256'), @('PriorYearSourceWorkbook', 'PriorYearSourceSha256'))) {
        if ((Get-GssOperationHash $fresh.($binding[0])) -cne [string]$fresh.($binding[1])) { throw 'A fresh source changed after report validation.' }
    }
    foreach ($source in @($fresh.SourceEvidence)) {
        if ($null -eq $source) { continue }
        $path = Resolve-GssOperationPath -Root $Folder -RelativePath ([string]$source.RelativePath) -RequireFile
        if ((Get-GssOperationHash $path) -cne [string]$source.Sha256) { throw 'A selected fresh source changed after report validation.' }
    }
    foreach ($name in @('CurrentSourceWorkbook', 'CurrentSourceRelativePath', 'CurrentSourceSha256', 'CurrentSourceSemanticSha256', 'PriorYearSourceWorkbook', 'PriorYearSourceRelativePath', 'PriorYearSourceSha256', 'PriorYearSourceSemanticSha256', 'SourceSetSha256', 'SourceEvidence', 'SourceChoiceSetSha256', 'AcceptedSourceChoices')) {
        $copy | Add-Member -NotePropertyName $name -NotePropertyValue $fresh.$name -Force
    }
    $evidence = foreach ($binding in @(
        @('live_workbook', 'TargetWorkbook', 'TargetWorkbookRelativePath', 'PromotedWorkbookSha256'),
        @('comparison_pdf', 'EmailComparisonPdf', 'EmailComparisonPdfRelativePath', 'PromotedPdfSha256'),
        @('rolling_workbook', 'CurrentSourceWorkbook', 'CurrentSourceRelativePath', 'CurrentSourceSha256'),
        @('prior_year_rolling_workbook', 'PriorYearSourceWorkbook', 'PriorYearSourceRelativePath', 'PriorYearSourceSha256')
    )) {
        $hash = Get-GssOperationHash $copy.($binding[1])
        if ($hash -cne [string]$copy.($binding[3])) { throw 'Report evidence changed while preparing its portable receipt.' }
        [pscustomobject]@{ Role = $binding[0]; RelativePath = $copy.($binding[2]); ByteSize = (Get-Item -LiteralPath $copy.($binding[1])).Length; Sha256 = $hash }
    }
    $copy | Add-Member -NotePropertyName FileEvidence -NotePropertyValue @($evidence) -Force
    $copy | Add-Member -NotePropertyName EvidenceKind -NotePropertyValue 'ValidatedSourceRefreshForExistingReport' -Force
    $copy | Add-Member -NotePropertyName OriginalRunLogSha256 -NotePropertyValue (Get-GssOperationHash $ReportRun.LogPath) -Force
    $copy | Add-Member -NotePropertyName OriginalRunFingerprint -NotePropertyValue $ReportRun.RunFingerprint -Force
    $copy | Add-Member -NotePropertyName FreshCheckLogRelativePath -NotePropertyValue (Get-GssOperationRelativePath -Root $Folder -Path $PreparedRun.LogPath) -Force
    $copy | Add-Member -NotePropertyName FreshCheckLogSha256 -NotePropertyValue (Get-GssOperationHash $PreparedRun.LogPath) -Force
    Import-Module (Join-Path $PSScriptRoot 'Gss-ReleaseIntegrity.psm1') -Scope Local
    $copy.RunFingerprint = Get-GssReleaseRunFingerprint -Run $copy
    return $copy
}

function Write-GssOperationSummary {
    param([Parameter(Mandatory)][string]$Folder, [Parameter(Mandatory)][object]$Receipt, [Parameter(Mandatory)][string]$ResumeScript)
    Write-GssOperationJson -Path (Join-Path $Folder '_automation_runs\state\latest-result.json') -Value $Receipt
    $summary = "GSS report: $($Receipt.ReportStatus)`r`nWeek ending: $($Receipt.ReportingWeekEnding)`r`nRun: $($Receipt.RunId)`r`nReport backup: $($Receipt.BackupStatus)`r`nDraft package: $($Receipt.PackageStatus)`r`nDraft evidence backup: $($Receipt.DraftEvidenceBackupStatus)`r`nUpdated: $($Receipt.UpdatedUtc)`r`n"
    if ($Receipt.PSObject.Properties['CurrentReportMatches'] -and -not $Receipt.CurrentReportMatches) { $summary += "This result concerns a historical report; current report files differ.`r`n" }
    if ($Receipt.Error) { $summary += "Details: $($Receipt.Error)`r`n" }
    if ($Receipt.BackupError) { $summary += "Backup details: $($Receipt.BackupError)`r`n" }
    elseif ($Receipt.Backup -and $Receipt.Backup.Error) { $summary += "Backup details: $($Receipt.Backup.Error)`r`n" }
    if ($Receipt.PackageError) { $summary += "Draft details: $($Receipt.PackageError)`r`n" }
    if ($Receipt.PackageBlockers) { $summary += "Draft details: $($Receipt.PackageBlockers -join '; ')`r`n" }
    if ($Receipt.DraftEvidenceBackupStatus -eq 'LocalOnly') { $summary += "Draft sources are frozen in the local package. The report backup does not confirm backup of later draft revisions.`r`n" }
    if ($Receipt.ReportStatus -in @('Ready', 'AlreadyCurrent', 'RecoveryRequired') -and ($Receipt.BackupStatus -ne 'Committed' -or $Receipt.PackageStatus -ne 'Ready')) {
        $retryRunId = if ($Receipt.ReportRunId) { $Receipt.ReportRunId } else { $Receipt.RunId }
        $summary += ('Retry: powershell.exe -NoProfile -ExecutionPolicy Bypass -File "{0}" -Folder "{1}" -RunId "{2}"' -f $ResumeScript, $Folder, $retryRunId) + "`r`n"
    }
    [IO.File]::WriteAllText((Join-Path $Folder 'Last GSS Result.txt'), $summary, (New-Object Text.UTF8Encoding($false)))
}

function Complete-GssUnstartedOperation {
    param([Parameter(Mandatory)][string]$Folder, [Parameter(Mandatory)][string]$RunId, [Parameter(Mandatory)][string]$ResumeScript)
    if ($RunId -notmatch '^[a-fA-F0-9-]{36}$') { throw 'RunId must be a GUID.' }
    $mutex = New-Object Threading.Mutex($false, 'Global\GSSSurveyWorkbookAutomationTransaction')
    $ownsMutex = $false
    try {
        try { $ownsMutex = $mutex.WaitOne(0) } catch [Threading.AbandonedMutexException] { $ownsMutex = $true }
        if (-not $ownsMutex) { throw 'Another GSS update is running on this computer.' }
        $state = Join-Path $Folder '_automation_runs\state'
        $activePath = Join-Path $state 'active-transaction.json'
        $receiptPath = Join-Path $state "operations\transaction-$RunId.json"
        $active = if (Test-Path -LiteralPath $activePath -PathType Leaf) { Get-Content -LiteralPath $activePath -Raw | ConvertFrom-Json } else { $null }
        if ($active -and [string]$active.RunId -ne $RunId) { throw 'A different report update must be resolved first.' }
        $receipt = if (Test-Path -LiteralPath $receiptPath -PathType Leaf) { Get-Content -LiteralPath $receiptPath -Raw | ConvertFrom-Json } else { $active }
        if ($null -eq $receipt -or [int]$receipt.ReceiptSchemaVersion -ne 2 -or [string]$receipt.RunId -ne $RunId -or
            [string]$receipt.TransactionStatus -notin @('Preparing', 'Promoting', 'RecoveryRequired')) { throw 'No supported unstarted operation receipt was found for this run.' }
        $logs = Join-Path $Folder '_automation_runs\logs'
        if (Test-Path -LiteralPath $logs -PathType Container) {
            if (@(Get-ChildItem -LiteralPath $logs -File -Filter "*_$RunId`_apply.json").Count -gt 0) { throw 'This run has an apply journal and requires transaction recovery.' }
        }
        if ([string]$receipt.TransactionStatus -ne 'Preparing') {
            if (-not $receipt.CopyRunLogPath) { throw 'The prepared report identity is missing; review this operation before clearing it.' }
            $preparedPath = Resolve-GssOperationPath -Root $Folder -RelativePath ([string]$receipt.CopyRunLogPath) -RequireFile
            $prepared = Get-Content -LiteralPath $preparedPath -Raw | ConvertFrom-Json
            if ([string]$prepared.RunId -ne $RunId -or [string]$prepared.Mode -ne 'CopyTestOnly') { throw 'The prepared report identity does not match this operation.' }
            $startingRelative = if ($prepared.StartingWorkbookRelativePath) { [string]$prepared.StartingWorkbookRelativePath } else { [string]$receipt.StartingWorkbookRelativePath }
            if (-not $startingRelative) { throw 'The previous report path is unavailable; review the unstarted operation.' }
            $startingPath = Resolve-GssOperationPath -Root $Folder -RelativePath $startingRelative -RequireFile
            if ((Get-GssOperationHash $startingPath) -cne [string]$prepared.StartingWorkbookSha256) { throw 'The live workbook changed; the missing journal cannot establish an unchanged report.' }
            if ($receipt.StartingPdfRelativePath) {
                $pdfPath = Resolve-GssOperationPath -Root $Folder -RelativePath ([string]$receipt.StartingPdfRelativePath)
                $pdfExists = Test-Path -LiteralPath $pdfPath -PathType Leaf
                if ($pdfExists -ne [bool]$receipt.StartingPdfExisted -or ($pdfExists -and (Get-GssOperationHash $pdfPath) -cne [string]$receipt.StartingPdfSha256)) { throw 'The report PDF changed; the missing journal cannot establish an unchanged report.' }
            }
        }
        # No apply journal means this update never crossed the first live-write
        # boundary. Do not alter any live file or any staged evidence.
        foreach ($entry in @{
            TransactionStatus = 'Aborted'; ReportStatus = 'PreviousReportUnchanged'; BackupStatus = 'NotStarted'; ReportBackupStatus = 'NotStarted';
            PackageStatus = 'NotPublished'; PackagePublished = $false; Error = $null; UpdatedUtc = [datetime]::UtcNow.ToString('o');
            Recovery = [pscustomobject]@{ Status = 'AbortedBeforePublication'; Reason = 'The update stopped before report publication; existing report files were not changed by recovery.' }
        }.GetEnumerator()) { $receipt | Add-Member -NotePropertyName $entry.Key -NotePropertyValue $entry.Value -Force }
        Write-GssOperationJson -Path $receiptPath -Value $receipt
        $latestPath = Join-Path $state 'latest-result.json'
        $latest = if (Test-Path -LiteralPath $latestPath -PathType Leaf) { Get-Content -LiteralPath $latestPath -Raw | ConvertFrom-Json } else { $null }
        if ($null -eq $latest -or [string]$latest.RunId -eq $RunId -or [string]$latest.ReportRunId -eq $RunId) { Write-GssOperationSummary -Folder $Folder -Receipt $receipt -ResumeScript $ResumeScript }
        if ($active) { Remove-Item -LiteralPath $activePath -Force }
        return $receipt
    }
    finally { if ($ownsMutex) { $mutex.ReleaseMutex() }; $mutex.Dispose() }
}

function Test-GssRunSnapshot {
    param([Parameter(Mandatory)][string]$ManifestPath)
    $manifest = Get-Content -LiteralPath $ManifestPath -Raw | ConvertFrom-Json
    if ([string]$manifest.schema_version -ne 'gss-run-snapshot/v1' -or [string]$manifest.run_id -notmatch '^[a-fA-F0-9-]{36}$') { throw 'Invalid local recovery snapshot identity.' }
    $reportDate = [datetime]::MinValue
    if (-not [datetime]::TryParseExact([string]$manifest.report_week, 'yyyy-MM-dd', [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::None, [ref]$reportDate)) { throw 'Invalid recovery snapshot report week.' }
    $root = Split-Path -Parent $ManifestPath
    $seen = @{}
    foreach ($entry in @($manifest.files)) {
        $key = ([string]$entry.path).ToLowerInvariant()
        if ($seen.ContainsKey($key)) { throw 'Duplicate recovery snapshot file.' }
        $seen[$key] = $true
        $path = Resolve-GssOperationPath -Root $root -RelativePath ([string]$entry.path) -RequireFile
        if ((Get-Item -LiteralPath $path).Length -ne [long]$entry.bytes -or (Get-GssOperationHash $path) -cne [string]$entry.sha256) { throw "Recovery snapshot file changed: $($entry.path)" }
    }
    foreach ($role in @('report_workbook', 'report_pdf', 'previous_workbook', 'run_log')) {
        if (@($manifest.files | Where-Object role -eq $role).Count -ne 1) { throw "Recovery snapshot requires exactly one $role artifact." }
    }
    return $manifest
}

function New-GssRunSnapshot {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Internal transaction backup called only by an explicitly launched update or recovery; adding an interactive confirmation would make recovery unreliable.')]
    [CmdletBinding()]
    param([Parameter(Mandatory)][object]$Run, [Parameter(Mandatory)][string]$Folder)
    $run = ConvertTo-GssPortableRun -Run $Run -Folder $Folder
    if ([string]$run.TransactionStatus -ne 'Committed' -or [string]$run.RunId -notmatch '^[a-fA-F0-9-]{36}$') { throw 'A committed reporting transaction is required for a recovery snapshot.' }
    $parent = Join-Path $Folder '_automation_runs\backups\runs'
    $destination = Resolve-GssOperationPath -Root $parent -RelativePath ([string]$run.RunId)
    $manifestPath = Join-Path $destination 'run-snapshot.json'
    if (Test-Path -LiteralPath $manifestPath -PathType Leaf) {
        $existing = Test-GssRunSnapshot $manifestPath
        if ([string]$existing.fingerprint -cne [string]$run.RunFingerprint) { throw 'Recovery snapshot run ID conflicts with its fingerprint.' }
        return $manifestPath
    }
    if (Test-Path -LiteralPath $destination) { throw "Incomplete local recovery snapshot requires review: $destination" }
    [void][IO.Directory]::CreateDirectory($parent)
    $staging = Resolve-GssOperationPath -Root $parent -RelativePath ('.staging-' + [guid]::NewGuid().ToString('N'))
    [void][IO.Directory]::CreateDirectory((Join-Path $staging 'files'))
    $entries = New-Object Collections.Generic.List[object]
    $sources = @(
        @{ Role = 'report_workbook'; Path = $run.StagedWorkbook; Expected = $run.StagedWorkbookSha256; Original = $run.TargetWorkbookRelativePath },
        @{ Role = 'report_pdf'; Path = $run.StagedPdf; Expected = $run.StagedPdfSha256; Original = $run.EmailComparisonPdfRelativePath },
        @{ Role = 'previous_workbook'; Path = $run.BackupWorkbook; Expected = $run.StartingWorkbookSha256; Original = $run.BackupWorkbookRelativePath },
        @{ Role = 'current_source'; Path = $run.CurrentSourceWorkbook; Expected = $run.CurrentSourceSha256; Original = $run.CurrentSourceRelativePath },
        @{ Role = 'prior_source'; Path = $run.PriorYearSourceWorkbook; Expected = $run.PriorYearSourceSha256; Original = $run.PriorYearSourceRelativePath },
        @{ Role = 'run_log'; Path = $run.LogPath; Expected = ''; Original = $run.LogRelativePath }
    )
    $includedSourcePaths = @{}
    foreach ($source in $sources) { if ($source.Path) { $includedSourcePaths[([string]$source.Path).ToLowerInvariant()] = $true } }
    foreach ($sourceEvidence in @($run.SourceEvidence)) {
        if ($null -eq $sourceEvidence) { continue }
        $sourcePath = Resolve-GssOperationPath -Root $Folder -RelativePath ([string]$sourceEvidence.RelativePath) -RequireFile
        if ($includedSourcePaths.ContainsKey($sourcePath.ToLowerInvariant())) { continue }
        $sources += @{ Role = 'applied_source'; Path = $sourcePath; Expected = [string]$sourceEvidence.Sha256; Original = [string]$sourceEvidence.RelativePath }
        $includedSourcePaths[$sourcePath.ToLowerInvariant()] = $true
    }
    if ($run.PdfBackup) { $sources += @{ Role = 'previous_pdf'; Path = $run.PdfBackup; Expected = [string]$run.PdfBackupSha256; Original = $run.PdfBackupRelativePath } }
    foreach ($relative in @('_automation_runs/state/detail-reporting.json', '_automation_runs/state/gss_feedback_first_seen.json')) {
        $path = Resolve-GssOperationPath -Root $Folder -RelativePath $relative
        if (Test-Path -LiteralPath $path -PathType Leaf) { $sources += @{ Role = 'reporting_state'; Path = $path; Expected = ''; Original = $relative } }
    }
    # Keep survey evidence and historical assignments recoverable without copying
    # accumulated QA folders, past draft packages, program trees, or old backups.
    foreach ($selection in @(
        @{ Folder = '03 Uploaded Survey Workbooks'; Filter = '*.xlsx'; Role = 'detail_source' },
        @{ Folder = '_automation_runs/historical-recovery'; Filter = '*.json'; Role = 'historical_evidence' },
        @{ Folder = '_automation_runs/archive/program-releases'; Filter = '*'; Role = 'archived_program' }
    )) {
        $sourceRoot = Resolve-GssOperationPath -Root $Folder -RelativePath $selection.Folder
        if (Test-Path -LiteralPath $sourceRoot -PathType Container) {
            foreach ($file in Get-ChildItem -LiteralPath $sourceRoot -File -Recurse -Filter $selection.Filter) {
                if ($file.Name.StartsWith('~$')) { continue }
                $sources += @{ Role = $selection.Role; Path = $file.FullName; Expected = ''; Original = (Get-GssOperationRelativePath -Root $Folder -Path $file.FullName) }
            }
        }
    }
    try {
        foreach ($source in $sources) {
            $hash = Get-GssOperationHash ([string]$source.Path)
            if ($source.Expected -and $hash -cne [string]$source.Expected) { throw "Recovery input no longer matches its transaction: $($source.Role)" }
            $relative = 'files/{0:d3}{1}' -f $entries.Count, [IO.Path]::GetExtension([string]$source.Path)
            $target = Resolve-GssOperationPath -Root $staging -RelativePath $relative
            Copy-Item -LiteralPath $source.Path -Destination $target -ErrorAction Stop
            if ((Get-GssOperationHash $target) -cne $hash -or (Get-GssOperationHash $source.Path) -cne $hash) { throw "Recovery input changed while copying: $($source.Role)" }
            $entries.Add([pscustomobject]@{ role = $source.Role; path = $relative; original_path = $source.Original; bytes = (Get-Item -LiteralPath $target).Length; sha256 = $hash })
        }
        $manifest = [ordered]@{ schema_version = 'gss-run-snapshot/v1'; run_id = [string]$run.RunId; fingerprint = [string]$run.RunFingerprint; report_week = [string]$run.CurrentWeekEnding; program_version = [string]$run.ProgramRelease; created_utc = [datetime]::UtcNow.ToString('o'); files = $entries.ToArray() }
        Write-GssOperationJson -Path (Join-Path $staging 'run-snapshot.json') -Value $manifest
        [void](Test-GssRunSnapshot (Join-Path $staging 'run-snapshot.json'))
        [IO.Directory]::Move($staging, $destination)
        return $manifestPath
    }
    finally { if (Test-Path -LiteralPath $staging) { Remove-Item -LiteralPath $staging -Recurse -Force } }
}

function Sync-GssRunSnapshot {
    param([Parameter(Mandatory)][string]$ManifestPath, [string]$BackupFolder)
    $manifest = Test-GssRunSnapshot $ManifestPath
    $backupRoot = Get-GssBackupFolder -BackupFolder $BackupFolder
    if ([string]::IsNullOrWhiteSpace($backupRoot) -or -not (Test-Path -LiteralPath $backupRoot -PathType Container)) {
        return [pscustomobject]@{ Status = 'Pending'; BackupStatus = 'Pending'; RunId = $manifest.run_id; Error = 'Backup folder is not configured or is unavailable.'; VerificationLevel = 'local_snapshot_verified'; CloudUploadConfirmed = $false }
    }
    $destination = Join-Path $backupRoot ('reports\{0}\{1}' -f $manifest.report_week, $manifest.run_id)
    $targetManifest = Join-Path $destination 'run-snapshot.json'
    $commitPath = Join-Path $destination 'COMMITTED.json'
    $manifestHash = Get-GssOperationHash $ManifestPath
    if (Test-Path -LiteralPath $commitPath) {
        $commit = Get-Content -LiteralPath $commitPath -Raw | ConvertFrom-Json
        if ([string]$commit.manifest_sha256 -cne $manifestHash) { throw 'Backup run ID already identifies a different snapshot.' }
        if ((Get-GssOperationHash $targetManifest) -cne $manifestHash) { throw 'Committed backup manifest no longer matches its receipt.' }
        [void](Test-GssRunSnapshot $targetManifest)
        return [pscustomobject]@{ Status = 'Committed'; BackupStatus = 'Committed'; RunId = $manifest.run_id; Fingerprint = $manifest.fingerprint; SnapshotPath = $destination; VerificationLevel = 'drivefs_hash_verified'; CloudUploadConfirmed = $false; Idempotent = $true }
    }
    [void][IO.Directory]::CreateDirectory($destination)
    foreach ($entry in @($manifest.files)) {
        $source = Resolve-GssOperationPath -Root (Split-Path -Parent $ManifestPath) -RelativePath $entry.path -RequireFile
        $target = Resolve-GssOperationPath -Root $destination -RelativePath $entry.path
        [void][IO.Directory]::CreateDirectory((Split-Path -Parent $target))
        Copy-Item -LiteralPath $source -Destination $target -Force -ErrorAction Stop
        if ((Get-GssOperationHash $target) -cne [string]$entry.sha256) { throw "Backup copy verification failed: $($entry.path)" }
    }
    Copy-Item -LiteralPath $ManifestPath -Destination $targetManifest -Force -ErrorAction Stop
    if ((Get-GssOperationHash $targetManifest) -cne $manifestHash) { throw 'Backup manifest copy verification failed.' }
    [void](Test-GssRunSnapshot $targetManifest)
    Write-GssOperationJson -Path $commitPath -Value ([ordered]@{ schema_version = 1; run_id = $manifest.run_id; fingerprint = $manifest.fingerprint; manifest_sha256 = $manifestHash; committed_utc = [datetime]::UtcNow.ToString('o'); verification_level = 'drivefs_hash_verified'; cloud_upload_confirmed = $false })
    [pscustomobject]@{ Status = 'Committed'; BackupStatus = 'Committed'; RunId = $manifest.run_id; Fingerprint = $manifest.fingerprint; SnapshotPath = $destination; VerificationLevel = 'drivefs_hash_verified'; CloudUploadConfirmed = $false; Idempotent = $false }
}
