BeforeAll {
    . (Join-Path (Split-Path -Parent $PSScriptRoot) 'scripts\Update-GSS-MainWorkbook.ps1')

    function New-TestTransaction {
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Creates an isolated disposable Pester TestDrive fixture.')]
        [CmdletBinding()]
        param([string]$Root, [bool]$ExistingPdf = $true)
        $folder = Join-Path $Root ([guid]::NewGuid().ToString('N'))
        $main = Join-Path $folder '01 Main Workbook\GSS.xlsx'
        $stage = Join-Path $folder '03 Automation\Tests\stage.xlsx'
        $stagePdf = Join-Path $folder '03 Automation\Tests\comparison.pdf'
        $livePdf = Join-Path $folder '04 Email Comparison PDFs\comparison.pdf'
        $backup = Join-Path $folder '03 Automation\Backups'
        $logs = Join-Path $folder '03 Automation\Logs'
        foreach ($directory in @((Split-Path $main), (Split-Path $stage), (Split-Path $livePdf), $backup, $logs, (Join-Path $folder '02 Rolling Reports'))) {
            New-Item -ItemType Directory -Path $directory -Force | Out-Null
        }
        [IO.File]::WriteAllText($main, 'previous verified workbook')
        [IO.File]::WriteAllText($stage, 'new validated workbook')
        [IO.File]::WriteAllText($stagePdf, 'new validated PDF')
        if ($ExistingPdf) { [IO.File]::WriteAllText($livePdf, 'previous verified PDF') }
        $evidence = @()
        foreach ($entry in @(
            @('current.xlsx', '2026-10-04', 'CurrentWeek'),
            @('prior.xlsx', '2025-10-05', 'PriorYearWeek'),
            @('catchup.xlsx', '2026-09-27', 'CatchUpWeek')
        )) {
            $source = Join-Path $folder ('02 Rolling Reports\' + $entry[0])
            [IO.File]::WriteAllText($source, ('validated ' + $entry[0]))
            $evidence += [pscustomobject]@{
                Role = $entry[2]; WeekEnding = $entry[1]; PeriodStart = ([datetime]$entry[1]).AddDays(-90).ToString('yyyy-MM-dd')
                RelativePath = ConvertTo-GssDropboxRelativePath $source $folder
                Sha256 = Get-GssSha256 $source
                SemanticSha256 = Get-GssTextSha256 ('content ' + $entry[0])
            }
        }
        $prepared = [pscustomobject]@{
            ReceiptSchemaVersion = 2; RunId = [guid]::NewGuid().ToString('N'); HostName = [Environment]::MachineName
            Mode = 'CopyTestOnly'; TransactionStatus = 'Prepared'; WorkbookStatus = 'READY'; GuardrailsApplied = $true
            CurrentWeekEnding = '2026-10-04'; PriorYearWeekEnding = '2025-10-05'; ProgramRelease = 'test'
            StartingWorkbookSha256 = Get-GssSha256 $main; CurrentSourceSha256 = $evidence[0].Sha256; PriorYearSourceSha256 = $evidence[1].Sha256
            CurrentSourceRelativePath = $evidence[0].RelativePath; PriorYearSourceRelativePath = $evidence[1].RelativePath
            TargetWorkbookRelativePath = ConvertTo-GssDropboxRelativePath $stage $folder
            EmailComparisonPdfRelativePath = ConvertTo-GssDropboxRelativePath $stagePdf $folder
            StagedWorkbookSha256 = Get-GssSha256 $stage; StagedPdfSha256 = Get-GssSha256 $stagePdf
            SourceEvidence = $evidence; SourceSetSha256 = Get-GssSourceEvidenceFingerprint $evidence
            RowsAppended = 8; RowsCorrected = 0; BusinessDataChanged = $true; NoBusinessChanges = $false
        }
        $prepared | Add-Member RunFingerprint (Get-GssFingerprintFromPreparedRun $prepared)
        return [pscustomobject]@{ Folder = $folder; Main = $main; Stage = $stage; StagePdf = $stagePdf; LivePdf = $livePdf; Backup = $backup; Logs = $logs; Prepared = $prepared }
    }

    function Invoke-TestPromotion {
        param([object]$Fixture)
        Invoke-GssPreparedRunPromotion -PreparedRun $Fixture.Prepared -ExpectedFingerprint $Fixture.Prepared.RunFingerprint -FolderPath $Fixture.Folder -MainWorkbookPath $Fixture.Main -BackupDirectory $Fixture.Backup -LogDirectory $Fixture.Logs
    }

    function New-TestInterruptedJournal {
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Injects interruption states only in an isolated disposable Pester TestDrive fixture.')]
        [CmdletBinding()]
        param([object]$Fixture, [string]$Point)
        $workbookBackup = Join-Path $Fixture.Backup 'prior.xlsx'
        Copy-Item -LiteralPath $Fixture.Main -Destination $workbookBackup
        $pdfExists = Test-Path -LiteralPath $Fixture.LivePdf -PathType Leaf
        $pdfBackup = Join-Path $Fixture.Backup 'prior.pdf'
        if ($pdfExists) { Copy-Item -LiteralPath $Fixture.LivePdf -Destination $pdfBackup }
        $journal = [pscustomobject]@{
            ReceiptSchemaVersion = 2; RunId = $Fixture.Prepared.RunId; TransactionStatus = $Point
            TargetWorkbookRelativePath = ConvertTo-GssDropboxRelativePath $Fixture.Main $Fixture.Folder
            BackupWorkbookRelativePath = ConvertTo-GssDropboxRelativePath $workbookBackup $Fixture.Folder
            StartingWorkbookSha256 = Get-GssSha256 $Fixture.Main; PromotedWorkbookSha256 = Get-GssSha256 $Fixture.Stage
            EmailComparisonPdfRelativePath = ConvertTo-GssDropboxRelativePath $Fixture.LivePdf $Fixture.Folder
            PromotedPdfSha256 = Get-GssSha256 $Fixture.StagePdf; PdfExistedBefore = $pdfExists
            StartingPdfSha256 = if ($pdfExists) { Get-GssSha256 $Fixture.LivePdf } else { $null }
            PdfBackupSha256 = if ($pdfExists) { Get-GssSha256 $pdfBackup } else { $null }
            PdfBackupRelativePath = if ($pdfExists) { ConvertTo-GssDropboxRelativePath $pdfBackup $Fixture.Folder } else { $null }
        }
        if ($Point -in @('WorkbookPromoted', 'PdfPromoted', 'Committed')) { Copy-Item -LiteralPath $Fixture.Stage -Destination $Fixture.Main -Force }
        if ($Point -in @('PdfPromoted', 'Committed')) { Copy-Item -LiteralPath $Fixture.StagePdf -Destination $Fixture.LivePdf -Force }
        return $journal
    }
}

Describe 'Prepared promotion evidence and local backup' {
    BeforeEach {
        $fixture = New-TestTransaction $TestDrive
        $null = $fixture # Consumed by each Pester test in this Describe block.
        Mock Test-GssPersistedWorkbook {
            [pscustomobject]@{ Status = 'Ready'; WorkbookPath = $WorkbookPath; WorstCaseWeeklyLoadsRemaining = 1000 }
        }
    }

    It 'promotes validated bytes using a verified local backup without requiring Drive' {
        $receipt = Invoke-TestPromotion $fixture
        $receipt.TransactionStatus | Should -Be 'Committed'
        $receipt.LocalBackupStatus | Should -Be 'Verified'
        (Get-GssSha256 $fixture.Main) | Should -Be $fixture.Prepared.StagedWorkbookSha256
        (Get-GssSha256 $receipt.BackupWorkbook) | Should -Be $fixture.Prepared.StartingWorkbookSha256
        (Get-GssSha256 $receipt.PdfBackup) | Should -Be $receipt.StartingPdfSha256
        $receipt.LivePostSaveValidation.ValidationMethod | Should -Be 'MatchesValidatedStagedWorkbook'
        $saved = Get-Content -LiteralPath $receipt.LogPath -Raw | ConvertFrom-Json
        $saved.SourceEvidence.Count | Should -Be 3
        $saved.ReceiptSchemaVersion | Should -Be 2
        Should -Invoke Test-GssPersistedWorkbook -Times 1 -Exactly
    }

    It 'leaves both live outputs untouched when stage validation fails' {
        Mock Test-GssPersistedWorkbook { throw 'injected persisted formula error' }
        $pdfHash = Get-GssSha256 $fixture.LivePdf
        { Invoke-TestPromotion $fixture } | Should -Throw '*formula error*'
        (Get-GssSha256 $fixture.Main) | Should -Be $fixture.Prepared.StartingWorkbookSha256
        (Get-GssSha256 $fixture.LivePdf) | Should -Be $pdfHash
        @(Get-ChildItem -LiteralPath $fixture.Backup -File).Count | Should -Be 0
    }

    It 'never overwrites an existing backup when a generated name collides' {
        Mock Get-Date { '20000101_000000' } -ParameterFilter { $Format -eq 'yyyyMMdd_HHmmss' }
        $key = (Get-GssTextSha256 $fixture.Prepared.RunId).Substring(0, 12)
        $backup = Join-Path $fixture.Backup ("main_20000101_000000_$key.xlsx")
        [IO.File]::WriteAllText($backup, 'existing backup must survive')
        $oldPdfHash = Get-GssSha256 $fixture.LivePdf
        { Invoke-TestPromotion $fixture } | Should -Throw '*already exists*'
        [IO.File]::ReadAllText($backup) | Should -Be 'existing backup must survive'
        (Get-GssSha256 $fixture.Main) | Should -Be $fixture.Prepared.StartingWorkbookSha256
        (Get-GssSha256 $fixture.LivePdf) | Should -Be $oldPdfHash
    }

    It 'rejects changed catch-up evidence before any live mutation' {
        $source = Resolve-GssDropboxPath $fixture.Prepared.SourceEvidence[2].RelativePath $fixture.Folder
        [IO.File]::WriteAllText($source, 'changed after prepare')
        { Invoke-TestPromotion $fixture } | Should -Throw '*Rolling source changed*'
        (Get-GssSha256 $fixture.Main) | Should -Be $fixture.Prepared.StartingWorkbookSha256
        Should -Invoke Test-GssPersistedWorkbook -Times 0 -Exactly
    }

    It 'binds the entire source set into the transaction fingerprint' {
        $original = $fixture.Prepared.RunFingerprint
        $fixture.Prepared.SourceEvidence[2].SemanticSha256 = Get-GssTextSha256 'changed content'
        { Assert-GssSourceEvidence $fixture.Prepared $fixture.Folder } | Should -Throw '*source-set evidence changed*'
        $fixture.Prepared.SourceSetSha256 = Get-GssSourceEvidenceFingerprint $fixture.Prepared.SourceEvidence
        (Get-GssFingerprintFromPreparedRun $fixture.Prepared) | Should -Not -Be $original
        { Invoke-TestPromotion $fixture } | Should -Throw '*fingerprint verification failed*'
    }

    It 'requires a supplied optional Drive manifest to match the run' {
        $manifest = Join-Path $fixture.Logs 'drive.json'
        [IO.File]::WriteAllText($manifest, (@{ status = 'Prepared'; run_id = 'different-run'; fingerprint = $fixture.Prepared.RunFingerprint } | ConvertTo-Json))
        { Invoke-GssPreparedRunPromotion -PreparedRun $fixture.Prepared -ExpectedFingerprint $fixture.Prepared.RunFingerprint -FolderPath $fixture.Folder -MainWorkbookPath $fixture.Main -BackupDirectory $fixture.Backup -LogDirectory $fixture.Logs -PreparedDriveManifestPath $manifest -PreparedDriveManifestSha256 (Get-GssSha256 $manifest) } | Should -Throw '*not bound*'
        (Get-GssSha256 $fixture.Main) | Should -Be $fixture.Prepared.StartingWorkbookSha256
    }

    It 'preserves or restores both outputs if journal write fails at <Point>' -ForEach @(
        @{ Point = 'Promoting' }, @{ Point = 'WorkbookPromoted' }, @{ Point = 'PdfPromoted' }
    ) {
        $script:injectedAt = $Point
        $script:injectionFired = $false
        Mock Write-GssAtomicJson {
            param($Path, $InputObject, $Depth)
            $null = $Depth
            Write-GssAtomicText -Path $Path -Content ($InputObject | ConvertTo-Json -Depth 8)
            if (-not $script:injectionFired -and $InputObject.TransactionStatus -eq $script:injectedAt) {
                $script:injectionFired = $true
                throw 'injected journal checkpoint failure'
            }
        }
        $oldPdfHash = Get-GssSha256 $fixture.LivePdf
        { Invoke-TestPromotion $fixture } | Should -Throw '*journal checkpoint failure*'
        $script:injectionFired | Should -BeTrue
        (Get-GssSha256 $fixture.Main) | Should -Be $fixture.Prepared.StartingWorkbookSha256
        (Get-GssSha256 $fixture.LivePdf) | Should -Be $oldPdfHash
        $receiptFile = Get-ChildItem -LiteralPath $fixture.Logs -Filter '*_apply.json' | Select-Object -First 1
        $saved = Get-Content -LiteralPath $receiptFile.FullName -Raw | ConvertFrom-Json
        $saved.Status | Should -BeIn @('Blocked', 'RolledBack')
        $saved.BackupWorkbookRelativePath | Should -Not -BeNullOrEmpty
    }
}

Describe 'Interrupted transaction recovery' {
    It 'recovers interruption at <Point> and supports a second harmless recovery' -ForEach @(
        @{ Point = 'Promoting' }, @{ Point = 'WorkbookPromoted' }, @{ Point = 'PdfPromoted' }, @{ Point = 'Committed' }
    ) {
        $fixture = New-TestTransaction $TestDrive
        $journal = New-TestInterruptedJournal $fixture $Point
        $result = Repair-GssInterruptedTransaction $journal $fixture.Folder
        $result.Status | Should -BeIn @('RolledBack', 'Unchanged')
        (Get-GssSha256 $fixture.Main) | Should -Be $journal.StartingWorkbookSha256
        (Get-GssSha256 $fixture.LivePdf) | Should -Be $journal.StartingPdfSha256
        (Repair-GssInterruptedTransaction $journal $fixture.Folder).Status | Should -Be 'Unchanged'
    }

    It 'removes only the PDF created by the interrupted run when there was no older PDF' {
        $fixture = New-TestTransaction $TestDrive -ExistingPdf $false
        $journal = New-TestInterruptedJournal $fixture 'PdfPromoted'
        (Repair-GssInterruptedTransaction $journal $fixture.Folder).PdfRollback | Should -Be 'RemovedRunOwnedFile'
        Test-Path -LiteralPath $fixture.LivePdf | Should -BeFalse
        (Repair-GssInterruptedTransaction $journal $fixture.Folder).Status | Should -Be 'Unchanged'
    }

    It 'refuses an unrelated workbook change without touching the run PDF' {
        $fixture = New-TestTransaction $TestDrive
        $journal = New-TestInterruptedJournal $fixture 'PdfPromoted'
        [IO.File]::WriteAllText($fixture.Main, 'newer user work')
        { Repair-GssInterruptedTransaction $journal $fixture.Folder } | Should -Throw '*newer or unrelated change*'
        [IO.File]::ReadAllText($fixture.Main) | Should -Be 'newer user work'
        (Get-GssSha256 $fixture.LivePdf) | Should -Be $journal.PromotedPdfSha256
    }

    It 'checks an unrelated PDF before restoring the workbook' {
        $fixture = New-TestTransaction $TestDrive
        $journal = New-TestInterruptedJournal $fixture 'PdfPromoted'
        [IO.File]::WriteAllText($fixture.LivePdf, 'newer user PDF')
        { Repair-GssInterruptedTransaction $journal $fixture.Folder } | Should -Throw '*PDF contains a newer*'
        (Get-GssSha256 $fixture.Main) | Should -Be $journal.PromotedWorkbookSha256
        [IO.File]::ReadAllText($fixture.LivePdf) | Should -Be 'newer user PDF'
    }

    It 'refuses a corrupted workbook backup without changing either output' {
        $fixture = New-TestTransaction $TestDrive
        $journal = New-TestInterruptedJournal $fixture 'PdfPromoted'
        $backup = Resolve-GssDropboxPath $journal.BackupWorkbookRelativePath $fixture.Folder
        [IO.File]::WriteAllText($backup, 'corrupted backup')
        { Repair-GssInterruptedTransaction $journal $fixture.Folder } | Should -Throw '*backup failed verification*'
        (Get-GssSha256 $fixture.Main) | Should -Be $journal.PromotedWorkbookSha256
        (Get-GssSha256 $fixture.LivePdf) | Should -Be $journal.PromotedPdfSha256
    }
}

Describe 'Durable choices on a successful unchanged report' {
    BeforeEach {
        $script:fixture = New-TestTransaction $TestDrive
        $prepared = $script:fixture.Prepared
        $prepared.NoBusinessChanges = $true
        $prepared.BusinessDataChanged = $false
        $prepared.RowsAppended = 0
        $prepared | Add-Member PostSaveValidation ([pscustomobject]@{ Status = 'Ready' })
        $choice = [pscustomobject]@{
            WeekEnding = '2026-10-04'; SelectedSemanticSha256 = $prepared.SourceEvidence[0].SemanticSha256
            KnownSemanticSha256 = @($prepared.SourceEvidence[0].SemanticSha256, (Get-GssTextSha256 'rejected newer upload'))
        }
        $prepared | Add-Member AcceptedSourceChoices @($choice)
        $prepared | Add-Member SourceChoiceSetSha256 (Get-GssSourceChoiceFingerprint @($choice))
        $prepared.RunFingerprint = Get-GssFingerprintFromPreparedRun $prepared
        $script:existingReport = [pscustomobject]@{
            TransactionStatus = 'Committed'; RunId = 'previous-report'
            TargetWorkbookRelativePath = ConvertTo-GssDropboxRelativePath $script:fixture.Main $script:fixture.Folder
            PromotedWorkbookSha256 = $prepared.StartingWorkbookSha256
        }
    }

    It 'does not treat an ordinary prepared copy test as an accepted choice' {
        Write-GssAtomicJson -Path (Join-Path $script:fixture.Logs 'prepared.json') -InputObject $script:fixture.Prepared
        @(Get-GssCommittedSourceChoice $script:fixture.Logs).Count | Should -Be 0
    }

    It 'persists a separate accepted decision without changing workbook or PDF bytes' {
        $pdfHash = Get-GssSha256 $script:fixture.LivePdf
        $path = Write-GssAcceptedSourceChoiceReceipt -PreparedRun $script:fixture.Prepared -ExistingRun $script:existingReport -FolderPath $script:fixture.Folder -LogDirectory $script:fixture.Logs
        $saved = Get-Content -LiteralPath $path -Raw | ConvertFrom-Json
        $saved.Mode | Should -Be 'SourceChoiceOnly'
        $saved.NoWorkbookOrPdfChange | Should -BeTrue
        @(Get-GssCommittedSourceChoice $script:fixture.Logs).Count | Should -Be 1
        (Get-GssSha256 $script:fixture.Main) | Should -Be $script:fixture.Prepared.StartingWorkbookSha256
        (Get-GssSha256 $script:fixture.LivePdf) | Should -Be $pdfHash
    }

    It 'rejects a no-op choice after the existing workbook changes' {
        [IO.File]::WriteAllText($script:fixture.Main, 'changed by another user')
        { Write-GssAcceptedSourceChoiceReceipt -PreparedRun $script:fixture.Prepared -ExistingRun $script:existingReport -FolderPath $script:fixture.Folder -LogDirectory $script:fixture.Logs } | Should -Throw '*workbook changed*'
        @(Get-GssCommittedSourceChoice $script:fixture.Logs).Count | Should -Be 0
    }

    It 'rejects changed data instead of recording it as an unchanged choice' {
        $script:fixture.Prepared.RowsCorrected = 4
        { Write-GssAcceptedSourceChoiceReceipt -PreparedRun $script:fixture.Prepared -ExistingRun $script:existingReport -FolderPath $script:fixture.Folder -LogDirectory $script:fixture.Logs } | Should -Throw '*validated unchanged workbook*'
    }
}
