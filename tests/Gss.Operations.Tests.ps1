BeforeAll {
    $operationsRepo = Split-Path -Parent $PSScriptRoot
    . (Join-Path $operationsRepo 'scripts\Gss-Operations.ps1')

    function New-GssOperationsFixture {
        [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Creates isolated test fixtures only under Pester TestDrive.')]
        [CmdletBinding()]
        param([string]$Root)
        $id = [guid]::NewGuid().ToString('D')
        $files = @{
            '01 Main Workbook/main.xlsx' = 'new report'
            '04 Email Comparison PDFs/report.pdf' = 'new PDF'
            '02 Weekly Rolling Source Workbooks/current.xlsx' = 'current source'
            '02 Weekly Rolling Source Workbooks/prior.xlsx' = 'prior source'
            '_automation_runs/test-output/stage/main.xlsx' = 'new report'
            '_automation_runs/test-output/stage/report.pdf' = 'new PDF'
            '_automation_runs/backups/previous.xlsx' = 'previous good report'
        }
        foreach ($pair in $files.GetEnumerator()) {
            $path = Join-Path $Root $pair.Key
            [void][IO.Directory]::CreateDirectory((Split-Path -Parent $path))
            [IO.File]::WriteAllText($path, $pair.Value)
        }
        $run = [pscustomobject]@{
            ReceiptSchemaVersion = 2; RunId = $id; RunFingerprint = ('a' * 64)
            HostName = 'ORIGINAL-COMPUTER'; ProgramRelease = 'v1.1.12'; Folder = $Root
            Mode = 'ApplyToMainWorkbook'; TransactionStatus = 'Committed'; CurrentWeekEnding = '2026-10-04'
            TargetWorkbookRelativePath = '01 Main Workbook/main.xlsx'
            StartingWorkbookRelativePath = '01 Main Workbook/main.xlsx'
            EmailComparisonPdfRelativePath = '04 Email Comparison PDFs/report.pdf'
            CurrentSourceRelativePath = '02 Weekly Rolling Source Workbooks/current.xlsx'
            PriorYearSourceRelativePath = '02 Weekly Rolling Source Workbooks/prior.xlsx'
            StagedWorkbookRelativePath = '_automation_runs/test-output/stage/main.xlsx'
            StagedPdfRelativePath = '_automation_runs/test-output/stage/report.pdf'
            BackupWorkbookRelativePath = '_automation_runs/backups/previous.xlsx'
            PdfBackup = $null
            LogRelativePath = "_automation_runs/logs/gss_update_fixture_$id`_apply.json"
        }
        foreach ($binding in @(
            @('TargetWorkbook', 'TargetWorkbookRelativePath'), @('EmailComparisonPdf', 'EmailComparisonPdfRelativePath'),
            @('CurrentSourceWorkbook', 'CurrentSourceRelativePath'), @('PriorYearSourceWorkbook', 'PriorYearSourceRelativePath'),
            @('StagedWorkbook', 'StagedWorkbookRelativePath'), @('StagedPdf', 'StagedPdfRelativePath'),
            @('BackupWorkbook', 'BackupWorkbookRelativePath'), @('LogPath', 'LogRelativePath')
        )) { $run | Add-Member -NotePropertyName $binding[0] -NotePropertyValue (Join-Path $Root $run.($binding[1])) }
        foreach ($binding in @(
            @('StartingWorkbookSha256', 'BackupWorkbook'), @('PromotedWorkbookSha256', 'TargetWorkbook'),
            @('PromotedPdfSha256', 'EmailComparisonPdf'), @('StagedWorkbookSha256', 'StagedWorkbook'),
            @('StagedPdfSha256', 'StagedPdf'), @('CurrentSourceSha256', 'CurrentSourceWorkbook'),
            @('PriorYearSourceSha256', 'PriorYearSourceWorkbook')
        )) { $run | Add-Member -NotePropertyName $binding[0] -NotePropertyValue (Get-GssOperationHash $run.($binding[1])) }
        Write-GssOperationJson -Path $run.LogPath -Value $run
        return $run
    }

    function New-GssLauncherSandbox {
        [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Creates isolated mock program files only under Pester TestDrive.')]
        [CmdletBinding()]
        param([string]$Root, [string]$Fault = '')
        $program = Join-Path $Root 'program'
        $data = Join-Path $Root 'data'
        [void][IO.Directory]::CreateDirectory((Join-Path $program 'scripts'))
        foreach ($name in @('Gss-Operations.ps1', 'Gss-ReleaseIntegrity.psm1', 'Invoke-GSS-SafeWorkbookUpdate.ps1')) {
            Copy-Item -LiteralPath (Join-Path $operationsRepo "scripts\$name") -Destination (Join-Path $program "scripts\$name")
        }
        $run = New-GssOperationsFixture -Root $data
        [void][IO.Directory]::CreateDirectory((Join-Path $data '.fixture'))
        Write-GssOperationJson -Path (Join-Path $data '.fixture\run.json') -Value $run
        [IO.File]::WriteAllText((Join-Path $data '.fixture\fault.txt'), $Fault)
        Copy-Item -LiteralPath $run.BackupWorkbook -Destination $run.TargetWorkbook -Force
        $updater = @'
[CmdletBinding()]
param($Folder, $MainWorkbookName, $RunId, $ProgramRelease, [switch]$MutexAlreadyHeld, [switch]$OutputObject, [switch]$Apply, $PreparedRunLogPath, $ExpectedFingerprint, $RollbackRunLogPath, $ReportingWeekEnding, $CorrectionSourcePath)
function Write-GssAcceptedSourceChoiceReceipt {
    param($PreparedRun, $ExistingRun, $FolderPath, $LogDirectory)
    $choicePath = Join-Path $LogDirectory 'accepted-source-choices.json'
    [IO.File]::WriteAllText($choicePath, ($PreparedRun.AcceptedSourceChoices | ConvertTo-Json -Depth 6))
    return $choicePath
}
if ($MyInvocation.InvocationName -eq '.') { return }
$run = Get-Content -LiteralPath (Join-Path $Folder '.fixture\run.json') -Raw | ConvertFrom-Json
$fault = Get-Content -LiteralPath (Join-Path $Folder '.fixture\fault.txt') -Raw
if ($RollbackRunLogPath) {
    Copy-Item -LiteralPath $run.BackupWorkbook -Destination $run.TargetWorkbook -Force
    return [pscustomobject]@{ Status = 'RolledBack' }
}
if ($Apply) {
    $prepared = Get-Content -LiteralPath $PreparedRunLogPath -Raw | ConvertFrom-Json
    $run.RunId = $prepared.RunId
    $run.LogPath = Join-Path $Folder "_automation_runs\logs\gss_update_fixture_$($run.RunId)_apply.json"
    $run.LogRelativePath = "_automation_runs/logs/gss_update_fixture_$($run.RunId)_apply.json"
    if ($fault -eq 'blocked-before-journal') { throw 'Simulated pre-journal validation failure' }
    if ($fault -eq 'blocked') {
        $run.TransactionStatus = 'Blocked'
        $run | Add-Member WorkbookWasPromoted $false
        $run | Add-Member PdfWasPromoted $false
        $run | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $run.LogPath
        throw 'Simulated blocked apply with neither artifact promoted'
    }
    if ($fault -eq 'interrupted') {
        $run.TransactionStatus = 'Promoting'
        $run | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $run.LogPath
        Copy-Item -LiteralPath $run.StagedWorkbook -Destination $run.TargetWorkbook -Force
        throw 'Simulated process interruption after workbook replacement'
    }
    Copy-Item -LiteralPath $run.StagedWorkbook -Destination $run.TargetWorkbook -Force
    if ($fault -eq 'snapshot-fail') { [IO.File]::WriteAllText($run.PriorYearSourceWorkbook, 'source changed before local snapshot') }
}
else {
    $run.RunId = $RunId
    $run.Mode = 'CopyTestOnly'
    $run.TransactionStatus = 'Prepared'
    $run.LogPath = Join-Path $Folder "_automation_runs\logs\prepared_$RunId.json"
    if ($fault -like 'noop*') {
        $run | Add-Member NoBusinessChanges $true
        $run.StartingWorkbookSha256 = (Get-FileHash -LiteralPath $run.TargetWorkbook -Algorithm SHA256).Hash.ToLowerInvariant()
        if ($fault -eq 'noop-choice') { $run | Add-Member AcceptedSourceChoices @([pscustomobject]@{ WeekEnding = $run.CurrentWeekEnding; SemanticSha256 = ('2' * 64) }) }
    }
}
$run | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $run.LogPath
$run
'@
        [IO.File]::WriteAllText((Join-Path $program 'scripts\Update-GSS-MainWorkbook.ps1'), $updater)
        $analyzer = @'
[CmdletBinding()]
param($Folder, $LogPath, [switch]$OutputObject, [switch]$PublishEmailPackage)
$run = Get-Content -LiteralPath $LogPath -Raw | ConvertFrom-Json
$fault = Get-Content -LiteralPath (Join-Path $Folder '.fixture\fault.txt') -Raw
if ($PublishEmailPackage) { 'published' | Set-Content -LiteralPath (Join-Path $Folder '.fixture\published.txt') }
[pscustomobject]@{
    WorkbookStatus = if ($fault -eq 'invalid' -and $run.Mode -eq 'ApplyToMainWorkbook') { 'Blocked' } else { 'Ready' }
    MarkdownPath = Join-Path $Folder '_automation_runs\qa\review.md'
    DetailReadiness = if ($fault -eq 'noop-details') { [pscustomobject]@{ Status = 'RequiresAssignment'; Candidates = @([pscustomobject]@{ Path = 'fixture.xlsx'; RelativePath = 'fixture.xlsx'; ContentSha256 = ('1' * 64); ResponseCount = 10; VisitDateStart = '2026-10-01'; VisitDateEnd = '2026-10-04' }) } } else { $null }
    EmailReadiness = 'WaitingForInput'
    EmailPackage = $null
    Qa = [pscustomobject]@{ EmailBlockers = @('No guest-detail fixture') }
}
'@
        [IO.File]::WriteAllText((Join-Path $program 'scripts\Analyze-GSS-Run.ps1'), $analyzer)
        $resume = @'
param($Folder, $LiveRunLogPath, $FreshPreparedRunLogPath, $BackupFolder, [switch]$OutputObject)
$run = Get-Content -LiteralPath $LiveRunLogPath -Raw | ConvertFrom-Json
[IO.File]::WriteAllText((Join-Path $Folder '.fixture\resumed-id.txt'), [string]$run.RunId)
[pscustomobject]@{ ReportStatus = 'Ready'; BackupStatus = 'Pending'; PackageStatus = 'WaitingForBackup'; PackagePublished = $false }
'@
        [IO.File]::WriteAllText((Join-Path $program 'scripts\Resume-GSS-PendingFinalize.ps1'), $resume)
        $detailSetter = @'
function Set-GssDetailReportingAssignment {
    param($FolderPath, $DetailPath, $ReportingDate, $SupersedesContentSha256)
    [IO.File]::WriteAllText((Join-Path $FolderPath '.fixture\assigned-week.txt'), [string]$ReportingDate)
}
'@
        [IO.File]::WriteAllText((Join-Path $program 'scripts\Gss-EmailPackage.ps1'), $detailSetter)
        [pscustomobject]@{ Program = $program; Data = $data; Run = $run; Launcher = (Join-Path $program 'scripts\Invoke-GSS-SafeWorkbookUpdate.ps1') }
    }
}

Describe 'Portable, independent reporting recovery' {
    It 'reads program version without Git or historical certification files' {
        $root = Join-Path $TestDrive 'version'
        Write-GssOperationJson -Path (Join-Path $root 'release\release-manifest.json') -Value @{ release_tag = 'v2.0.0' }
        (Get-GssRuntimeInfo -RepoRoot $root).ReleaseTag | Should -Be 'v2.0.0'
        (Test-Path -LiteralPath (Join-Path $root '.git')) | Should -BeFalse
    }

    It 'captures the previous good report and every selected catch-up source' {
        $root = Join-Path $TestDrive 'snapshot'
        $run = New-GssOperationsFixture $root
        $extra = Join-Path $root '02 Weekly Rolling Source Workbooks\catchup.xlsx'
        [IO.File]::WriteAllText($extra, 'catch-up source')
        $run | Add-Member SourceEvidence @([pscustomobject]@{ RelativePath = '02 Weekly Rolling Source Workbooks/catchup.xlsx'; Sha256 = (Get-GssOperationHash $extra) })
        $registry = Join-Path $root '_automation_runs\state\detail-reporting.json'
        Write-GssOperationJson $registry @{ schema_version = 1 }
        $archive = Join-Path $root '_automation_runs\archive\program-releases\v1\program.zip'
        [void][IO.Directory]::CreateDirectory((Split-Path -Parent $archive))
        [IO.File]::WriteAllText($archive, 'archived release')
        $path = New-GssRunSnapshot -Run $run -Folder $root
        $manifest = Test-GssRunSnapshot $path
        @($manifest.files | Where-Object role -eq 'previous_workbook').Count | Should -Be 1
        @($manifest.files | Where-Object role -eq 'applied_source').Count | Should -Be 1
        @($manifest.files | Where-Object role -eq 'reporting_state').Count | Should -Be 1
        @($manifest.files | Where-Object role -eq 'archived_program').Count | Should -Be 1
    }

    It 'rebinds legacy log paths to another root without changing old evidence or host metadata' {
        $first = Join-Path $TestDrive 'first'
        $second = Join-Path $TestDrive 'second'
        $run = New-GssOperationsFixture $first
        $run.ReceiptSchemaVersion = 1
        $run.PSObject.Properties.Remove('StagedWorkbookRelativePath')
        $before = $run | ConvertTo-Json -Depth 10
        $portable = ConvertTo-GssPortableRun -Run $run -Folder $second
        $portable.StagedWorkbook | Should -Be (Join-Path $second '_automation_runs\test-output\stage\main.xlsx')
        $portable.HostName | Should -Be 'ORIGINAL-COMPUTER'
        ($run | ConvertTo-Json -Depth 10) | Should -Be $before
    }

    It 'refreshes validated source evidence for a no-op while preserving the original report log' {
        $root = Join-Path $TestDrive 'fresh-evidence'
        $run = New-GssOperationsFixture $root
        $originalHash = Get-GssOperationHash $run.LogPath
        $fresh = $run | ConvertTo-Json -Depth 15 | ConvertFrom-Json
        $fresh.RunId = [guid]::NewGuid().ToString('D')
        $fresh.Mode = 'CopyTestOnly'
        $fresh | Add-Member NoBusinessChanges $true
        $fresh.StartingWorkbookSha256 = $run.PromotedWorkbookSha256
        [IO.File]::WriteAllText($fresh.CurrentSourceWorkbook, 'same business data, different workbook metadata')
        $fresh.CurrentSourceSha256 = Get-GssOperationHash $fresh.CurrentSourceWorkbook
        $fresh.LogPath = Join-Path $root '_automation_runs\logs\fresh.json'
        Write-GssOperationJson -Path $fresh.LogPath -Value $fresh
        $derived = ConvertTo-GssRefreshedAnalysisRun -ReportRun $run -PreparedRun $fresh -Folder $root
        $derived.CurrentSourceSha256 | Should -Be $fresh.CurrentSourceSha256
        $derived.RunFingerprint | Should -Not -Be $run.RunFingerprint
        $derived.OriginalRunFingerprint | Should -Be $run.RunFingerprint
        ($derived.FileEvidence | Where-Object Role -eq 'rolling_workbook').Sha256 | Should -Be $fresh.CurrentSourceSha256
        (Get-GssOperationHash $run.LogPath) | Should -Be $originalHash
        [IO.File]::WriteAllText($fresh.CurrentSourceWorkbook, 'changed after review')
        { ConvertTo-GssRefreshedAnalysisRun -ReportRun $run -PreparedRun $fresh -Folder $root } | Should -Throw '*fresh source changed*'
    }

    It 'requires one new validated report when the reporting release changes' {
        $root = Join-Path $TestDrive 'release-refresh'
        $run = New-GssOperationsFixture $root
        $prepared = $run | ConvertTo-Json -Depth 15 | ConvertFrom-Json
        $prepared | Add-Member NoBusinessChanges $true
        $prepared.StartingWorkbookSha256 = $run.PromotedWorkbookSha256
        (Find-GssExistingReport -PreparedRun $prepared -Folder $root).RunId | Should -Be $run.RunId
        $prepared.ProgramRelease = 'v2.0.0'
        Find-GssExistingReport -PreparedRun $prepared -Folder $root | Should -BeNullOrEmpty
    }

    It 'closes an interrupted preparation without an apply journal or a workbook restore' {
        $root = Join-Path $TestDrive 'unstarted-preparation'
        $id = [guid]::NewGuid().ToString('D')
        $operation = [pscustomobject]@{ ReceiptSchemaVersion = 2; RunId = $id; TransactionStatus = 'Preparing' }
        Write-GssOperationJson -Path (Join-Path $root '_automation_runs\state\active-transaction.json') -Value $operation
        $result = & (Join-Path $operationsRepo 'scripts\Resume-GSS-PendingFinalize.ps1') -Folder $root -RunId $id -OutputObject
        $result.Recovery.Status | Should -Be 'AbortedBeforePublication'
        $result.ReportStatus | Should -Be 'PreviousReportUnchanged'
        (Test-Path -LiteralPath (Join-Path $root '_automation_runs\state\active-transaction.json')) | Should -BeFalse
    }

    It 'checks starting workbook and PDF before closing a promotion with a missing journal' {
        $root = Join-Path $TestDrive 'unstarted-promotion'
        $run = New-GssOperationsFixture $root
        Remove-Item -LiteralPath $run.LogPath -Force
        Copy-Item -LiteralPath $run.BackupWorkbook -Destination $run.TargetWorkbook -Force
        $run.Mode = 'CopyTestOnly'
        $run.LogPath = Join-Path $root '_automation_runs\logs\prepared.json'
        Write-GssOperationJson -Path $run.LogPath -Value $run
        $operation = [pscustomobject]@{
            ReceiptSchemaVersion = 2; RunId = $run.RunId; TransactionStatus = 'Promoting'; CopyRunLogPath = '_automation_runs/logs/prepared.json'
            StartingPdfRelativePath = $run.EmailComparisonPdfRelativePath; StartingPdfExisted = $true; StartingPdfSha256 = $run.PromotedPdfSha256
        }
        $active = Join-Path $root '_automation_runs\state\active-transaction.json'
        Write-GssOperationJson -Path $active -Value $operation
        [IO.File]::WriteAllText($run.TargetWorkbook, 'unrelated newer workbook')
        { & (Join-Path $operationsRepo 'scripts\Resume-GSS-PendingFinalize.ps1') -Folder $root -RunId $run.RunId -OutputObject } | Should -Throw '*live workbook changed*'
        (Test-Path -LiteralPath $active) | Should -BeTrue
        Copy-Item -LiteralPath $run.BackupWorkbook -Destination $run.TargetWorkbook -Force
        [IO.File]::WriteAllText($run.EmailComparisonPdf, 'unrelated newer PDF')
        { & (Join-Path $operationsRepo 'scripts\Resume-GSS-PendingFinalize.ps1') -Folder $root -RunId $run.RunId -OutputObject } | Should -Throw '*report PDF changed*'
        [IO.File]::WriteAllText($run.EmailComparisonPdf, 'new PDF')
        $result = & (Join-Path $operationsRepo 'scripts\Resume-GSS-PendingFinalize.ps1') -Folder $root -RunId $run.RunId -OutputObject
        $result.Recovery.Status | Should -Be 'AbortedBeforePublication'
        (Test-Path -LiteralPath $active) | Should -BeFalse
    }

    It 'handles a durable <Status> apply receipt without an unrecoverable state gate' -TestCases @(@{ Status = 'Aborted' }, @{ Status = 'Blocked' }, @{ Status = 'Conflict' }) {
        param($Status)
        $root = Join-Path $TestDrive "terminal-$Status"
        $run = New-GssOperationsFixture $root
        $run.TransactionStatus = $Status
        $run | Add-Member WorkbookWasPromoted ($Status -ne 'Blocked')
        $run | Add-Member PdfWasPromoted ($Status -ne 'Blocked')
        if ($Status -ne 'Conflict') {
            Copy-Item -LiteralPath $run.BackupWorkbook -Destination $run.TargetWorkbook -Force
            Remove-Item -LiteralPath $run.EmailComparisonPdf -Force
        }
        Write-GssOperationJson -Path $run.LogPath -Value $run
        $hash = Get-GssOperationHash $run.LogPath
        $active = Join-Path $root '_automation_runs\state\active-transaction.json'
        Write-GssOperationJson -Path $active -Value @{ ReceiptSchemaVersion = 2; RunId = $run.RunId; TransactionStatus = 'RecoveryRequired' }
        $result = & (Join-Path $operationsRepo 'scripts\Resume-GSS-PendingFinalize.ps1') -Folder $root -RunId $run.RunId -RestorePreviousReport -OutputObject
        $result.TransactionStatus | Should -Be 'Aborted'
        (Get-GssOperationHash $run.TargetWorkbook) | Should -Be $run.StartingWorkbookSha256
        (Get-GssOperationHash $run.LogPath) | Should -Be $hash
        (Test-Path -LiteralPath $active) | Should -BeFalse
    }

    It 'keeps a conflicted apply blocked when current files are unrelated' {
        $root = Join-Path $TestDrive 'conflict-unrelated'
        $run = New-GssOperationsFixture $root
        $run.TransactionStatus = 'Conflict'
        Write-GssOperationJson -Path $run.LogPath -Value $run
        [IO.File]::WriteAllText($run.TargetWorkbook, 'unrelated current workbook')
        $active = Join-Path $root '_automation_runs\state\active-transaction.json'
        Write-GssOperationJson -Path $active -Value @{ ReceiptSchemaVersion = 2; RunId = $run.RunId; TransactionStatus = 'RecoveryRequired' }
        { & (Join-Path $operationsRepo 'scripts\Resume-GSS-PendingFinalize.ps1') -Folder $root -RunId $run.RunId -RestorePreviousReport -OutputObject } | Should -Throw '*newer or unrelated*'
        (Test-Path -LiteralPath $active) | Should -BeTrue
        (Get-Content -LiteralPath $run.TargetWorkbook -Raw) | Should -Be 'unrelated current workbook'
    }

    It 'keeps a verified local snapshot when the cloud folder is unavailable' {
        $root = Join-Path $TestDrive 'offline'
        $run = New-GssOperationsFixture $root
        $snapshot = New-GssRunSnapshot -Run $run -Folder $root
        $result = Sync-GssRunSnapshot -ManifestPath $snapshot -BackupFolder (Join-Path $TestDrive 'unavailable')
        $result.BackupStatus | Should -Be 'Pending'
        $result.CloudUploadConfirmed | Should -BeFalse
        (Test-GssRunSnapshot $snapshot).run_id | Should -Be $run.RunId
    }

    It 'retries an immutable snapshot after later reporting changed live files, without duplicating the backup' {
        $root = Join-Path $TestDrive 'retry'
        $backup = Join-Path $TestDrive 'backup'
        [void][IO.Directory]::CreateDirectory($backup)
        $run = New-GssOperationsFixture $root
        $snapshot = New-GssRunSnapshot -Run $run -Folder $root
        [IO.File]::WriteAllText($run.TargetWorkbook, 'next week report')
        [IO.File]::WriteAllText($run.CurrentSourceWorkbook, 'next week source')
        (New-GssRunSnapshot -Run $run -Folder $root) | Should -Be $snapshot
        $first = Sync-GssRunSnapshot -ManifestPath $snapshot -BackupFolder $backup
        $second = Sync-GssRunSnapshot -ManifestPath $snapshot -BackupFolder $backup
        $first.BackupStatus | Should -Be 'Committed'
        $second.Idempotent | Should -BeTrue
        (Get-ChildItem -LiteralPath (Join-Path $backup 'reports') -Recurse -Filter 'COMMITTED.json').Count | Should -Be 1
        (Get-Content -LiteralPath $run.TargetWorkbook -Raw) | Should -Be 'next week report'
    }

    It 'finishes a partially copied cloud snapshot but refuses corruption of a committed one' {
        $root = Join-Path $TestDrive 'partial'
        $backup = Join-Path $TestDrive 'partial-backup'
        [void][IO.Directory]::CreateDirectory($backup)
        $run = New-GssOperationsFixture $root
        $snapshot = New-GssRunSnapshot -Run $run -Folder $root
        $partial = Join-Path $backup "reports\2026-10-04\$($run.RunId)\files\000.xlsx"
        [void][IO.Directory]::CreateDirectory((Split-Path -Parent $partial))
        [IO.File]::WriteAllText($partial, 'interrupted transfer')
        (Sync-GssRunSnapshot -ManifestPath $snapshot -BackupFolder $backup).Status | Should -Be 'Committed'
        [IO.File]::WriteAllText($partial, 'changed after commit')
        { Sync-GssRunSnapshot -ManifestPath $snapshot -BackupFolder $backup } | Should -Throw '*snapshot file changed*'
    }

    It 'resumes backup on a relocated folder for an older release and host without changing its original log' {
        $first = Join-Path $TestDrive 'resume-original'
        $moved = Join-Path $TestDrive 'resume-moved'
        $backup = Join-Path $TestDrive 'resume-backup'
        [void][IO.Directory]::CreateDirectory($backup)
        $run = New-GssOperationsFixture $first
        [void](New-GssRunSnapshot -Run $run -Folder $first)
        Copy-Item -LiteralPath $first -Destination $moved -Recurse
        $movedLog = Join-Path $moved $run.LogRelativePath
        $logHash = Get-GssOperationHash $movedLog
        $newerSummary = [pscustomobject]@{ RunId = [guid]::NewGuid().ToString('D'); ReportStatus = 'Ready'; BackupStatus = 'Pending' }
        Write-GssOperationJson -Path (Join-Path $moved '_automation_runs\state\latest-result.json') -Value $newerSummary
        [IO.File]::WriteAllText((Join-Path $moved $run.TargetWorkbookRelativePath), 'a later report')
        $result = & (Join-Path $operationsRepo 'scripts\Resume-GSS-PendingFinalize.ps1') -Folder $moved -RunId $run.RunId -BackupFolder $backup -OutputObject
        $result.BackupStatus | Should -Be 'Committed'
        $result.PackageStatus | Should -Be 'SupersededOrChanged'
        $result.CurrentReportMatches | Should -BeFalse
        (Get-GssOperationHash $movedLog) | Should -Be $logHash
        (Get-Content -LiteralPath (Join-Path $moved $run.TargetWorkbookRelativePath) -Raw) | Should -Be 'a later report'
        (Get-Content -LiteralPath (Join-Path $moved '_automation_runs\state\latest-result.json') -Raw | ConvertFrom-Json).RunId | Should -Be $newerSummary.RunId
    }

    It 'leaves a failed historical backup pending without locking the newer current report' {
        $root = Join-Path $TestDrive 'historical-missing-snapshot'
        $run = New-GssOperationsFixture $root
        [IO.File]::WriteAllText($run.TargetWorkbook, 'newer report')
        [IO.File]::WriteAllText($run.CurrentSourceWorkbook, 'newer source')
        $latest = [pscustomobject]@{ RunId = [guid]::NewGuid().ToString('D'); ReportStatus = 'Ready'; BackupStatus = 'Committed' }
        Write-GssOperationJson -Path (Join-Path $root '_automation_runs\state\latest-result.json') -Value $latest
        $result = & (Join-Path $operationsRepo 'scripts\Resume-GSS-PendingFinalize.ps1') -Folder $root -RunId $run.RunId -BackupFolder (Join-Path $TestDrive 'missing-history-backup') -OutputObject
        $result.BackupStatus | Should -Be 'Pending'
        $result.LocalSnapshotStatus | Should -Be 'Failed'
        $result.CurrentReportMatches | Should -BeFalse
        (Test-Path -LiteralPath (Join-Path $root '_automation_runs\state\active-transaction.json')) | Should -BeFalse
        (Get-Content -LiteralPath (Join-Path $root '_automation_runs\state\latest-result.json') -Raw | ConvertFrom-Json).RunId | Should -Be $latest.RunId
        (Get-Content -LiteralPath $run.TargetWorkbook -Raw) | Should -Be 'newer report'
    }
}

Describe 'Launcher transaction outcomes' {
    It 'closes a failed apply without a stale marker when <Fault>' -TestCases @(@{ Fault = 'blocked-before-journal' }, @{ Fault = 'blocked' }) {
        param($Fault)
        $sandbox = New-GssLauncherSandbox -Root (Join-Path $TestDrive "launcher-$Fault") -Fault $Fault
        Copy-Item -LiteralPath (Join-Path $operationsRepo 'scripts\Resume-GSS-PendingFinalize.ps1') -Destination (Join-Path $sandbox.Program 'scripts\Resume-GSS-PendingFinalize.ps1') -Force
        { & $sandbox.Launcher -Folder $sandbox.Data -NonInteractive } | Should -Throw '*Simulated*'
        $latest = Get-Content -LiteralPath (Join-Path $sandbox.Data '_automation_runs\state\latest-result.json') -Raw | ConvertFrom-Json
        $latest.ReportStatus | Should -Be 'PreviousReportUnchanged'
        (Get-GssOperationHash $sandbox.Run.TargetWorkbook) | Should -Be $sandbox.Run.StartingWorkbookSha256
        (Test-Path -LiteralPath (Join-Path $sandbox.Data '_automation_runs\state\active-transaction.json')) | Should -BeFalse
    }
    It 'commits with a local snapshot offline, publishes no draft, and releases the write marker' {
        $sandbox = New-GssLauncherSandbox -Root (Join-Path $TestDrive 'offline-launch')
        $result = & $sandbox.Launcher -Folder $sandbox.Data -BackupFolder (Join-Path $TestDrive 'offline-drive') -NonInteractive -OutputObject
        $result.ReportStatus | Should -Be 'Ready'
        $result.BackupStatus | Should -Be 'Pending'
        $result.PackageStatus | Should -Be 'WaitingForBackup'
        (Test-Path -LiteralPath (Join-Path $sandbox.Data 'Last GSS Result.txt')) | Should -BeTrue
        (Test-Path -LiteralPath (Join-Path $sandbox.Data '_automation_runs\state\active-transaction.json')) | Should -BeFalse
        (Test-Path -LiteralPath (Join-Path $sandbox.Data '.fixture\published.txt')) | Should -BeFalse
        # A second independent report is permitted while the first backup waits.
        $next = & $sandbox.Launcher -Folder $sandbox.Data -BackupFolder (Join-Path $TestDrive 'offline-drive') -NonInteractive -OutputObject
        $next.RunId | Should -Not -Be $result.RunId
        $next.ReportStatus | Should -Be 'Ready'
    }

    It 'restores the previous good report when post-promotion validation fails' {
        $sandbox = New-GssLauncherSandbox -Root (Join-Path $TestDrive 'bad-validation') -Fault 'invalid'
        { & $sandbox.Launcher -Folder $sandbox.Data -NonInteractive -OutputObject } | Should -Throw '*completed report failed validation*'
        (Get-GssOperationHash $sandbox.Run.TargetWorkbook) | Should -Be $sandbox.Run.StartingWorkbookSha256
        $status = Get-Content -LiteralPath (Join-Path $sandbox.Data '_automation_runs\state\latest-result.json') -Raw | ConvertFrom-Json
        $status.ReportStatus | Should -Be 'PreviousReportRestored'
        (Test-Path -LiteralPath (Join-Path $sandbox.Data '_automation_runs\state\active-transaction.json')) | Should -BeFalse
    }

    It 'retains a write marker until failed local snapshot creation is repaired and refreshes the last result' {
        $sandbox = New-GssLauncherSandbox -Root (Join-Path $TestDrive 'local-snapshot-failure') -Fault 'snapshot-fail'
        $first = & $sandbox.Launcher -Folder $sandbox.Data -BackupFolder (Join-Path $TestDrive 'snapshot-offline') -NonInteractive -OutputObject
        $first.TransactionStatus | Should -Be 'LocalSnapshotPending'
        $activePath = Join-Path $sandbox.Data '_automation_runs\state\active-transaction.json'
        (Get-Content -LiteralPath $activePath -Raw | ConvertFrom-Json).RunId | Should -Be $first.RunId
        { & $sandbox.Launcher -Folder $sandbox.Data -NonInteractive } | Should -Throw '*Recovery command:*'
        $interruptedStatus = Get-Content -LiteralPath $activePath -Raw | ConvertFrom-Json
        $interruptedStatus.TransactionStatus = 'ReportReady'
        Write-GssOperationJson -Path $activePath -Value $interruptedStatus
        $interruptedError = try { & $sandbox.Launcher -Folder $sandbox.Data -NonInteractive } catch { $_.Exception.Message }
        $interruptedError | Should -Match 'Recovery command:'
        $interruptedError | Should -Not -Match 'RestorePreviousReport'
        [IO.File]::WriteAllText($sandbox.Run.PriorYearSourceWorkbook, 'prior source')
        Copy-Item -LiteralPath (Join-Path $operationsRepo 'scripts\Resume-GSS-PendingFinalize.ps1') -Destination (Join-Path $sandbox.Program 'scripts\Resume-GSS-PendingFinalize.ps1') -Force
        $backup = Join-Path $TestDrive 'snapshot-online'
        [void][IO.Directory]::CreateDirectory($backup)
        $resumed = & (Join-Path $sandbox.Program 'scripts\Resume-GSS-PendingFinalize.ps1') -Folder $sandbox.Data -RunId $first.RunId -BackupFolder $backup -OutputObject
        $resumed.LocalSnapshotStatus | Should -Be 'Verified'
        $resumed.TransactionStatus | Should -Be 'Committed'
        $resumed.BackupError | Should -BeNullOrEmpty
        (Test-Path -LiteralPath $activePath) | Should -BeFalse
        $latest = Get-Content -LiteralPath (Join-Path $sandbox.Data '_automation_runs\state\latest-result.json') -Raw | ConvertFrom-Json
        $latest.LocalSnapshotStatus | Should -Be 'Verified'
        $latest.ReportBackupStatus | Should -Be 'Committed'
        (Get-Content -LiteralPath (Join-Path $sandbox.Data 'Last GSS Result.txt') -Raw) | Should -Match 'Report backup: Committed'
        (Get-Content -LiteralPath (Join-Path $sandbox.Data 'Last GSS Result.txt') -Raw) | Should -Not -Match 'Recovery input no longer matches'
    }

    It 'offers a late detail assignment even when the report itself is already current' {
        $sandbox = New-GssLauncherSandbox -Root (Join-Path $TestDrive 'late-detail-assignment')
        $first = & $sandbox.Launcher -Folder $sandbox.Data -BackupFolder (Join-Path $TestDrive 'detail-offline') -NonInteractive -OutputObject
        [IO.File]::WriteAllText((Join-Path $sandbox.Data '.fixture\fault.txt'), 'noop-details')
        Mock Read-Host { '1' }
        $second = & $sandbox.Launcher -Folder $sandbox.Data -BackupFolder (Join-Path $TestDrive 'detail-offline') -OutputObject
        $second.ReportRunId | Should -Be $first.RunId
        (Get-Content -LiteralPath (Join-Path $sandbox.Data '.fixture\assigned-week.txt') -Raw) | Should -Be '2026-10-04'
    }

    It 'persists accepted source choices from a successful no-op without another report revision' {
        $sandbox = New-GssLauncherSandbox -Root (Join-Path $TestDrive 'no-op-source-choice')
        $first = & $sandbox.Launcher -Folder $sandbox.Data -BackupFolder (Join-Path $TestDrive 'choice-offline') -NonInteractive -OutputObject
        [IO.File]::WriteAllText((Join-Path $sandbox.Data '.fixture\fault.txt'), 'noop-choice')
        $before = @(Get-ChildItem -LiteralPath (Join-Path $sandbox.Data '_automation_runs\logs') -Filter '*_apply.json').Count
        $second = & $sandbox.Launcher -Folder $sandbox.Data -BackupFolder (Join-Path $TestDrive 'choice-offline') -NonInteractive -OutputObject
        $second.ReportStatus | Should -Be 'AlreadyCurrent'
        $second.ReportRunId | Should -Be $first.RunId
        $second.AcceptedSourceChoiceLogPath | Should -Be '_automation_runs/logs/accepted-source-choices.json'
        @(Get-ChildItem -LiteralPath (Join-Path $sandbox.Data '_automation_runs\logs') -Filter '*_apply.json').Count | Should -Be $before
    }

    It 'leaves an already current report untouched and retries its existing backup without a new business revision' {
        $sandbox = New-GssLauncherSandbox -Root (Join-Path $TestDrive 'already-current')
        $first = & $sandbox.Launcher -Folder $sandbox.Data -BackupFolder (Join-Path $TestDrive 'noop-offline') -NonInteractive -OutputObject
        $workbookHash = Get-GssOperationHash $sandbox.Run.TargetWorkbook
        $before = @(Get-ChildItem -LiteralPath (Join-Path $sandbox.Data '_automation_runs\logs') -Filter '*_apply.json').Count
        [IO.File]::WriteAllText((Join-Path $sandbox.Data '.fixture\fault.txt'), 'noop')
        $second = & $sandbox.Launcher -Folder $sandbox.Data -BackupFolder (Join-Path $TestDrive 'noop-offline') -NonInteractive -OutputObject
        $second.ReportStatus | Should -Be 'AlreadyCurrent'
        $second.ReportRunId | Should -Be $first.RunId
        (Get-Content -LiteralPath (Join-Path $sandbox.Data '.fixture\resumed-id.txt') -Raw) | Should -Be $first.RunId
        (Get-GssOperationHash $sandbox.Run.TargetWorkbook) | Should -Be $workbookHash
        @(Get-ChildItem -LiteralPath (Join-Path $sandbox.Data '_automation_runs\logs') -Filter '*_apply.json').Count | Should -Be $before
    }

    It 'retains the journal and blocks the next write after an interrupted workbook replacement' {
        $sandbox = New-GssLauncherSandbox -Root (Join-Path $TestDrive 'interrupted') -Fault 'interrupted'
        { & $sandbox.Launcher -Folder $sandbox.Data -NonInteractive } | Should -Throw '*Simulated process interruption*'
        $activePath = Join-Path $sandbox.Data '_automation_runs\state\active-transaction.json'
        $active = Get-Content -LiteralPath $activePath -Raw | ConvertFrom-Json
        $active.TransactionStatus | Should -Be 'RecoveryRequired'
        (Get-ChildItem -LiteralPath (Join-Path $sandbox.Data '_automation_runs\logs') -Filter "*_$($active.RunId)_apply.json").Count | Should -Be 1
        $before = Get-GssOperationHash $sandbox.Run.TargetWorkbook
        { & $sandbox.Launcher -Folder $sandbox.Data -NonInteractive } | Should -Throw '*Recovery command:*Resume-GSS-PendingFinalize.ps1*'
        (Get-GssOperationHash $sandbox.Run.TargetWorkbook) | Should -Be $before
        (Get-Content -LiteralPath $activePath -Raw | ConvertFrom-Json).RunId | Should -Be $active.RunId
    }
}
