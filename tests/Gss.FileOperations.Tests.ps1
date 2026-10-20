BeforeAll {
    $fileOperationsRepo = Split-Path -Parent $PSScriptRoot
    . (Join-Path $fileOperationsRepo 'scripts\Gss-Operations.ps1')
    if (-not ('GssSyntheticDirectoryLock' -as [type])) {
        Add-Type -TypeDefinition @'
using System;
using System.ComponentModel;
using System.Runtime.InteropServices;
using Microsoft.Win32.SafeHandles;
public static class GssSyntheticDirectoryLock {
    [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern SafeFileHandle CreateFile(string path, uint access, uint share, IntPtr security, uint disposition, uint flags, IntPtr template);
    public static SafeFileHandle Open(string path) {
        // Allow reads and writes, but deny deletion/rename until the test releases it.
        SafeFileHandle handle = CreateFile(path, 0x80000000, 3, IntPtr.Zero, 3, 0x02000000, IntPtr.Zero);
        if (handle.IsInvalid) { int error = Marshal.GetLastWin32Error(); handle.Dispose(); throw new Win32Exception(error); }
        return handle;
    }
}
'@
    }
    $script:realSnapshotCheck = (Get-Command Test-GssRunSnapshot).ScriptBlock

    function New-GssFileOperationFixture {
        [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Creates synthetic files under Pester TestDrive only.')]
        [CmdletBinding()]
        param([string]$Root)
        $id = [guid]::NewGuid().ToString('D')
        $run = [pscustomobject]@{
            ReceiptSchemaVersion = 2; RunId = $id; RunFingerprint = ('b' * 64)
            Folder = $Root; TransactionStatus = 'Committed'; CurrentWeekEnding = '2026-10-11'; ProgramRelease = 'test'
            TargetWorkbookRelativePath = 'reports/main.xlsx'; EmailComparisonPdfRelativePath = 'reports/report.pdf'
            StagedWorkbookRelativePath = 'stage/main.xlsx'; StagedPdfRelativePath = 'stage/report.pdf'
            BackupWorkbookRelativePath = 'previous/main.xlsx'; CurrentSourceRelativePath = 'sources/current.xlsx'
            PriorYearSourceRelativePath = 'sources/prior.xlsx'; LogRelativePath = 'logs/run.json'
        }
        foreach ($binding in @(
            @('TargetWorkbook', 'TargetWorkbookRelativePath'), @('EmailComparisonPdf', 'EmailComparisonPdfRelativePath'),
            @('StagedWorkbook', 'StagedWorkbookRelativePath'), @('StagedPdf', 'StagedPdfRelativePath'),
            @('BackupWorkbook', 'BackupWorkbookRelativePath'), @('CurrentSourceWorkbook', 'CurrentSourceRelativePath'),
            @('PriorYearSourceWorkbook', 'PriorYearSourceRelativePath'), @('LogPath', 'LogRelativePath')
        )) {
            $path = Join-Path $Root $run.($binding[1])
            [void][IO.Directory]::CreateDirectory((Split-Path -Parent $path))
            [IO.File]::WriteAllText($path, ('synthetic ' + $binding[0]))
            $run | Add-Member -NotePropertyName $binding[0] -NotePropertyValue $path
        }
        foreach ($binding in @(
            @('StagedWorkbookSha256', 'StagedWorkbook'), @('StagedPdfSha256', 'StagedPdf'),
            @('StartingWorkbookSha256', 'BackupWorkbook'), @('CurrentSourceSha256', 'CurrentSourceWorkbook'),
            @('PriorYearSourceSha256', 'PriorYearSourceWorkbook')
        )) { $run | Add-Member -NotePropertyName $binding[0] -NotePropertyValue (Get-GssOperationHash $run.($binding[1])) }
        return $run
    }
}

Describe 'Bounded Windows file-lock retries' {
    It 'retries a wrapped Windows lock error <Code> and returns the successful value' -TestCases @(@{ Code = -2147024864 }, @{ Code = -2147024863 }) {
        param($Code)
        $state = [pscustomobject]@{ Attempts = 0; Code = $Code }
        Mock Start-Sleep {}
        $result = Invoke-GssFileOperationRetry -Operation {
            $state.Attempts++
            if ($state.Attempts -lt 3) {
                throw [Management.Automation.MethodInvocationException]::new('wrapped file operation', [IO.IOException]::new('temporary lock', $state.Code))
            }
            'complete'
        }
        $result | Should -Be 'complete'
        $state.Attempts | Should -Be 3
        Should -Invoke Start-Sleep -Times 2 -Exactly
    }

    It 'does not retry a non-lock IO error <Code>' -TestCases @(@{ Code = -2147024891 }, @{ Code = -2147024894 }, @{ Code = -2147024784 }) {
        param($Code)
        $state = [pscustomobject]@{ Attempts = 0; Code = $Code }
        Mock Start-Sleep {}
        { Invoke-GssFileOperationRetry {
            $state.Attempts++
            throw [IO.IOException]::new('permanent failure', $state.Code)
        } } | Should -Throw '*permanent failure*'
        $state.Attempts | Should -Be 1
        Should -Invoke Start-Sleep -Times 0 -Exactly
    }

    It 'stops after its bounded number of attempts and preserves the error' {
        $state = [pscustomobject]@{ Attempts = 0 }
        Mock Start-Sleep {}
        { Invoke-GssFileOperationRetry -MaxAttempts 3 -Operation {
            $state.Attempts++
            throw [IO.IOException]::new('persistent sharing violation', -2147024864)
        } } | Should -Throw '*persistent sharing violation*'
        $state.Attempts | Should -Be 3
        Should -Invoke Start-Sleep -Times 2 -Exactly
    }

    It 'bounds the explicit directory-rename access-denied retry and preserves permanent denial' {
        $state = [pscustomobject]@{ Attempts = 0 }
        Mock Start-Sleep {}
        $failure = $null
        try {
            Invoke-GssFileOperationRetry -RetryDirectoryRenameAccessDenied -MaxAttempts 20 -RetryDelayMilliseconds 1000 -Operation {
                $state.Attempts++
                throw [Management.Automation.MethodInvocationException]::new('wrapped rename', [IO.IOException]::new('persistent rename denial', -2147024891))
            }
        }
        catch { $failure = $_ }
        $failure.Exception.Message | Should -Be 'wrapped rename'
        $failure.Exception.InnerException.Message | Should -Be 'persistent rename denial'
        $failure.Exception.InnerException.HResult | Should -Be -2147024891
        $state.Attempts | Should -Be 20
        Should -Invoke Start-Sleep -Times 19 -Exactly -ParameterFilter { $Milliseconds -eq 1000 }
    }
}

Describe 'Atomic receipt publication and temporary cleanup' {
    It 'publishes a complete new JSON file without temporary leftovers' {
        $root = Join-Path $TestDrive 'new-json'
        $path = Join-Path $root 'result.json'
        Write-GssOperationJson $path @{ status = 'Ready'; values = @(1, 2, 3) }
        $result = Get-Content -LiteralPath $path -Raw | ConvertFrom-Json
        $result.status | Should -Be 'Ready'
        @($result.values).Count | Should -Be 3
        @(Get-ChildItem -LiteralPath $root -Force).Count | Should -Be 1
    }

    It 'retries an actual sharing violation before atomically replacing an existing receipt' {
        $path = Join-Path $TestDrive 'replace-json\result.json'
        Write-GssOperationJson $path @{ revision = 1 }
        $script:gssFileLock = [IO.File]::Open($path, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::Read)
        Mock Start-Sleep { $script:gssFileLock.Dispose() }
        try { Write-GssOperationJson $path @{ revision = 2 } }
        finally { $script:gssFileLock.Dispose() }
        (Get-Content -LiteralPath $path -Raw | ConvertFrom-Json).revision | Should -Be 2
        @(Get-ChildItem -LiteralPath (Split-Path -Parent $path) -Force).Count | Should -Be 1
        Should -Invoke Start-Sleep -Times 1 -Exactly
    }

    It 'retries an actual lock on the operator status file' {
        $root = Join-Path $TestDrive 'locked-summary'
        $receipt = [pscustomobject]@{ RunId = [guid]::NewGuid().ToString('D'); ReportStatus = 'Ready'; BackupStatus = 'Committed'; PackageStatus = 'WaitingForInput' }
        Write-GssOperationSummary -Folder $root -Receipt $receipt -ResumeScript 'resume.ps1'
        $path = Join-Path $root 'Last GSS Result.txt'
        $script:gssSummaryLock = [IO.File]::Open($path, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::Read)
        Mock Start-Sleep { $script:gssSummaryLock.Dispose() }
        try { Write-GssOperationSummary -Folder $root -Receipt $receipt -ResumeScript 'resume.ps1' }
        finally { $script:gssSummaryLock.Dispose() }
        (Get-Content -LiteralPath $path -Raw) | Should -Match 'GSS report: Ready'
        Should -Invoke Start-Sleep -Times 1 -Exactly
    }

    It 'keeps a successful publication successful when previous-file cleanup remains locked' {
        $path = Join-Path $TestDrive 'cleanup-after-success\result.json'
        Write-GssOperationJson $path @{ revision = 1 }
        Mock Remove-Item { throw [IO.IOException]::new('cleanup-only sharing violation', -2147024864) }
        Mock Start-Sleep {}
        $warnings = @()
        Write-GssOperationJson $path @{ revision = 2 } -WarningVariable warnings
        (Get-Content -LiteralPath $path -Raw | ConvertFrom-Json).revision | Should -Be 2
        ($warnings -join ' ') | Should -Match 'retained for review'
        @(Get-ChildItem -LiteralPath (Split-Path -Parent $path) -Force -Filter '.previous-*').Count | Should -Be 1
        Should -Invoke Remove-Item -Times 8 -Exactly
        Should -Invoke Start-Sleep -Times 7 -Exactly
    }

    It 'preserves a failed atomic replacement when temporary cleanup also fails' {
        $path = Join-Path $TestDrive 'cleanup-after-failure\result.json'
        Write-GssOperationJson $path @{ revision = 1 }
        $fileLock = [IO.File]::Open($path, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::Read)
        Mock Remove-Item { throw [IO.IOException]::new('cleanup-only sharing violation', -2147024864) }
        Mock Start-Sleep {}
        $failure = $null
        try { Write-GssOperationJson $path @{ revision = 2 } }
        catch { $failure = $_ }
        finally { $fileLock.Dispose() }
        $failure | Should -Not -BeNullOrEmpty
        $failure.Exception | Should -BeOfType ([Management.Automation.MethodInvocationException])
        $failure.Exception.InnerException.HResult | Should -Be -2147024864
        $failure.Exception.Message | Should -Not -Match 'cleanup-only'
        (Get-Content -LiteralPath $path -Raw | ConvertFrom-Json).revision | Should -Be 1
        Should -Invoke Start-Sleep -Times 14 -Exactly
    }

    It 'refuses recursive cleanup outside the exact temporary parent or of the parent itself' {
        $parent = Join-Path $TestDrive 'contained'
        $outside = Join-Path $TestDrive ('.staging-' + [guid]::NewGuid().ToString('N'))
        [void][IO.Directory]::CreateDirectory($parent)
        [void][IO.Directory]::CreateDirectory($outside)
        Mock Remove-Item { throw 'unsafe deletion was attempted' }
        Remove-GssOperationTemporaryPath -Path $outside -Parent $parent -Recurse
        Remove-GssOperationTemporaryPath -Path $parent -Parent $parent -Recurse
        Test-Path -LiteralPath $outside | Should -BeTrue
        Test-Path -LiteralPath $parent | Should -BeTrue
        Should -Invoke Remove-Item -Times 0 -Exactly
    }
}

Describe 'Recovery snapshot directory publication' {
    It 'retries directory access denied caused by an actual child file without delete sharing' {
        $root = Join-Path $TestDrive 'snapshot-child-lock'
        $run = New-GssFileOperationFixture $root
        Mock Test-GssRunSnapshot {
            param($ManifestPath)
            $checked = & $realSnapshotCheck -ManifestPath $ManifestPath
            $child = Join-Path (Split-Path -Parent $ManifestPath) ([string]$checked.files[0].path)
            $script:gssSnapshotChildLock = [IO.File]::Open($child, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::Read)
            return $checked
        }
        Mock Start-Sleep { $script:gssSnapshotChildLock.Dispose() }
        try { $manifestPath = New-GssRunSnapshot -Run $run -Folder $root }
        finally { if ($script:gssSnapshotChildLock) { $script:gssSnapshotChildLock.Dispose() } }
        $manifest = & $realSnapshotCheck -ManifestPath $manifestPath
        $manifest.run_id | Should -Be $run.RunId
        @($manifest.files).Count | Should -Be 6
        @(Get-ChildItem -LiteralPath (Split-Path -Parent (Split-Path -Parent $manifestPath)) -Force -Filter '.staging-*').Count | Should -Be 0
        Should -Invoke Start-Sleep -Times 1 -Exactly -ParameterFilter { $Milliseconds -eq 1000 }
    }

    It 'retries an actual directory sharing violation and publishes the verified snapshot' {
        $root = Join-Path $TestDrive 'snapshot-lock'
        $run = New-GssFileOperationFixture $root
        Mock Test-GssRunSnapshot {
            param($ManifestPath)
            $checked = & $script:realSnapshotCheck -ManifestPath $ManifestPath
            $script:gssDirectoryLock = [GssSyntheticDirectoryLock]::Open((Split-Path -Parent $ManifestPath))
            return $checked
        }
        Mock Start-Sleep { $script:gssDirectoryLock.Dispose() }
        try { $manifestPath = New-GssRunSnapshot -Run $run -Folder $root }
        finally { if ($script:gssDirectoryLock) { $script:gssDirectoryLock.Dispose() } }
        $manifest = & $script:realSnapshotCheck -ManifestPath $manifestPath
        $manifest.run_id | Should -Be $run.RunId
        @($manifest.files).Count | Should -Be 6
        @(Get-ChildItem -LiteralPath (Split-Path -Parent (Split-Path -Parent $manifestPath)) -Force -Filter '.staging-*').Count | Should -Be 0
        Should -Invoke Start-Sleep -Times 1 -Exactly
    }

    It 'preserves a snapshot validation failure when staging cleanup also fails' {
        $root = Join-Path $TestDrive 'snapshot-cleanup-failure'
        $run = New-GssFileOperationFixture $root
        [IO.File]::WriteAllText($run.PriorYearSourceWorkbook, 'changed source')
        Mock Remove-Item { throw [IO.IOException]::new('cleanup-only sharing violation', -2147024864) }
        Mock Start-Sleep {}
        { New-GssRunSnapshot -Run $run -Folder $root } | Should -Throw '*Recovery input no longer matches its transaction: prior_source*'
        Test-Path -LiteralPath (Join-Path $root "_automation_runs\backups\runs\$($run.RunId)") | Should -BeFalse
        @(Get-ChildItem -LiteralPath (Join-Path $root '_automation_runs\backups\runs') -Force -Filter '.staging-*').Count | Should -Be 1
        Should -Invoke Start-Sleep -Times 7 -Exactly
    }
}
