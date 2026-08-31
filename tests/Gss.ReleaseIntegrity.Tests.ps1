$modulePath = Join-Path (Split-Path -Parent $PSScriptRoot) 'scripts\Gss-ReleaseIntegrity.psm1'
Import-Module $modulePath -Force

Describe 'GSS release integrity module' {
    It 'selects program and policy files but excludes data and the manifest itself' {
        (Test-GssReleaseTrackedPath -Path 'scripts/Update-GSS-MainWorkbook.ps1') | Should -BeTrue
        (Test-GssReleaseTrackedPath -Path 'scripts/future_analytics.py') | Should -BeTrue
        (Test-GssReleaseTrackedPath -Path 'requirements-future-analytics.lock') | Should -BeTrue
        (Test-GssReleaseTrackedPath -Path 'config/analysis-policy.json') | Should -BeTrue
        (Test-GssReleaseTrackedPath -Path 'release/release-manifest.json') | Should -BeFalse
        (Test-GssReleaseTrackedPath -Path 'fixtures/GSS Score Trends - Main.xlsx') | Should -BeFalse
    }

    It 'detects a changed file hash without opening workbook content' {
        $filePath = Join-Path $TestDrive 'scripts\sample.ps1'
        New-Item -ItemType Directory -Path (Split-Path -Parent $filePath) -Force | Out-Null
        Set-Content -LiteralPath $filePath -Value 'Write-Output original' -Encoding UTF8
        $evidence = Get-GssReleaseFileEvidence -Path $filePath
        $manifest = [pscustomobject]@{
            files = @(
                [pscustomobject]@{
                    path = 'scripts/sample.ps1'
                    hash_mode = $evidence.HashMode
                    canonical_size_bytes = $evidence.CanonicalSizeBytes
                    sha256 = $evidence.Sha256
                }
            )
        }

        @(Test-GssReleaseManifestFile -Manifest $manifest -RepoRoot $TestDrive).Count | Should -Be 0
        Add-Content -LiteralPath $filePath -Value 'Write-Output changed'
        @(Test-GssReleaseManifestFile -Manifest $manifest -RepoRoot $TestDrive) | Should -Contain 'Release file canonical size changed: scripts/sample.ps1'
        @(Test-GssReleaseManifestFile -Manifest $manifest -RepoRoot $TestDrive) | Should -Contain 'Release file hash changed: scripts/sample.ps1'
    }

    It 'treats CRLF and LF as the same reviewed text content' {
        $lfPath = Join-Path $TestDrive 'lf.ps1'
        $crlfPath = Join-Path $TestDrive 'crlf.ps1'
        [System.IO.File]::WriteAllText($lfPath, "one`ntwo`n", (New-Object System.Text.UTF8Encoding($false)))
        [System.IO.File]::WriteAllText($crlfPath, "one`r`ntwo`r`n", (New-Object System.Text.UTF8Encoding($false)))

        (Get-GssReleaseFileEvidence $lfPath).Sha256 | Should -Be (Get-GssReleaseFileEvidence $crlfPath).Sha256
    }

    It 'accepts exact hashed Excel evidence certified on another workstation' {
        $head = 'a' * 40
        $tag = 'v1.0.0'
        $dataRoot = Join-Path $TestDrive 'GSS Surveys'
        $workbookRelativePath = '_automation_runs/test-output/run-id/GSS Score Trends - Main.xlsx'
        $sourceLogRelativePath = '_automation_runs/logs/copy-run.json'
        $workbookPath = Join-Path $dataRoot $workbookRelativePath.Replace('/', '\')
        $sourceLogPath = Join-Path $dataRoot $sourceLogRelativePath.Replace('/', '\')
        New-Item -ItemType Directory -Path (Split-Path -Parent $workbookPath), (Split-Path -Parent $sourceLogPath) -Force | Out-Null
        Set-Content -LiteralPath $workbookPath -Value 'copy-only workbook evidence' -Encoding UTF8
        $workbookHash = (Get-FileHash -LiteralPath $workbookPath -Algorithm SHA256).Hash.ToLowerInvariant()
        $sourceRun = [pscustomobject]@{
            Mode = 'CopyTestOnly'
            TransactionStatus = 'Prepared'
            ProgramRelease = $tag
            HostName = 'RELEASE-CERTIFICATION-HOST'
            RunId = '11111111-1111-1111-1111-111111111111'
            CurrentWeekEnding = '2026-07-19'
            StartingWorkbookSha256 = '1' * 64
            CurrentSourceSha256 = '2' * 64
            PriorYearSourceSha256 = '3' * 64
            StagedWorkbookRelativePath = $workbookRelativePath
            StagedWorkbookSha256 = $workbookHash
            StagedPdfSha256 = '4' * 64
            RunFingerprint = $null
        }
        $sourceRun.RunFingerprint = Get-GssReleaseRunFingerprint -Run $sourceRun
        $sourceRun | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $sourceLogPath -Encoding UTF8
        $receipt = [pscustomobject]@{
            ReceiptSchemaVersion = 1
            Status = 'Passed'
            GitHead = $head
            ReleaseTag = $tag
            ExcelVersion = '16.0'
            WorkbookPath = $workbookRelativePath
            WorkbookSha256 = $workbookHash
            SourceRunFingerprint = $sourceRun.RunFingerprint
            SourceRunLogPath = $sourceLogRelativePath
            FormulaErrors = 0
            ConstantErrors = 0
        }

        $sourceRun.HostName | Should -Not -Be ([Environment]::MachineName)
        @(Test-GssExcelValidationReceipt -Receipt $receipt -ExpectedHead $head -ExpectedTag $tag -DataRoot $dataRoot).Count | Should -Be 0
        $sourceRun.HostName = ''
        $sourceRun.RunFingerprint = Get-GssReleaseRunFingerprint -Run $sourceRun
        $sourceRun | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $sourceLogPath -Encoding UTF8
        $receipt.SourceRunFingerprint = $sourceRun.RunFingerprint
        @(Test-GssExcelValidationReceipt -Receipt $receipt -ExpectedHead $head -ExpectedTag $tag -DataRoot $dataRoot) | Should -Contain 'Receipt source run has no audit workstation.'
        $sourceRun.HostName = 'RELEASE-CERTIFICATION-HOST'
        $sourceRun.RunFingerprint = Get-GssReleaseRunFingerprint -Run $sourceRun
        $sourceRun | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $sourceLogPath -Encoding UTF8
        $receipt.SourceRunFingerprint = $sourceRun.RunFingerprint
        $receipt.GitHead = 'd' * 40
        @(Test-GssExcelValidationReceipt -Receipt $receipt -ExpectedHead $head -ExpectedTag $tag -DataRoot $dataRoot) | Should -Contain 'Receipt Git HEAD does not match the release.'
        $receipt.GitHead = $head
        $receipt.WorkbookPath = '01 Main Workbook/GSS Score Trends - Main.xlsx'
        @(Test-GssExcelValidationReceipt -Receipt $receipt -ExpectedHead $head -ExpectedTag $tag -DataRoot $dataRoot) | Should -Contain 'Receipt workbook is not an isolated copy-test artifact.'
        $receipt.WorkbookPath = $workbookRelativePath
        Add-Content -LiteralPath $workbookPath -Value 'tampered'
        @(Test-GssExcelValidationReceipt -Receipt $receipt -ExpectedHead $head -ExpectedTag $tag -DataRoot $dataRoot) | Should -Contain 'Receipt workbook hash does not match the copy-test artifact.'
    }
}
