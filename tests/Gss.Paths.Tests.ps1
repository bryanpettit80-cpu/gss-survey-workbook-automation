BeforeAll {
    . (Join-Path (Split-Path -Parent $PSScriptRoot) 'scripts\Gss-Common.ps1')
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    Add-Type -AssemblyName System.IO.Compression

    function New-TestMainWorkbook {
        [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Creates synthetic files only under the isolated Pester TestDrive.')]
        [CmdletBinding()]
        param([string]$Path, [string[]]$Sheets = @('Raw_Data', 'QA Checks', 'Email Comparison'))
        $archive = [System.IO.Compression.ZipFile]::Open($Path, [System.IO.Compression.ZipArchiveMode]::Create)
        try {
            $entry = $archive.CreateEntry('xl/workbook.xml')
            $writer = New-Object System.IO.StreamWriter($entry.Open())
            try {
                $nodes = @($Sheets | ForEach-Object { '<sheet name="' + $_ + '" />' }) -join ''
                $writer.Write('<workbook><sheets>' + $nodes + '</sheets></workbook>')
            }
            finally { $writer.Dispose() }
        }
        finally { $archive.Dispose() }
    }
}

Describe 'Main workbook discovery' {
    BeforeEach {
        $root = Join-Path $TestDrive ([guid]::NewGuid().ToString('N'))
        $main = Join-Path $root '01 Main Workbook'
        New-Item -ItemType Directory -Path $main -Force | Out-Null
    }

    It 'recognizes a harmless rename from workbook structure' {
        $renamed = Join-Path $main 'Guest Survey Trends.xlsx'
        New-TestMainWorkbook -Path $renamed
        Resolve-MainWorkbookPath $root 'GSS Score Trends - Main.xlsx' | Should -Be $renamed
    }

    It 'uses normal discovery for an omitted name and never returns a directory' {
        $renamed = Join-Path $main 'Guest Survey Trends.xlsx'
        New-TestMainWorkbook -Path $renamed
        Resolve-MainWorkbookPath $root '' | Should -Be $renamed
        New-Item -ItemType Directory -Path (Join-Path $main 'Wrong.xlsx') | Out-Null
        { Resolve-MainWorkbookPath $root 'Wrong.xlsx' } | Should -Throw '*Main workbook not found*'
    }

    It 'requires a choice between multiple renamed main workbooks' {
        New-TestMainWorkbook -Path (Join-Path $main 'First.xlsx')
        New-TestMainWorkbook -Path (Join-Path $main 'Second.xlsx')
        { Resolve-MainWorkbookPath $root 'GSS Score Trends - Main.xlsx' } | Should -Throw '*Multiple score-trends*'
        Resolve-MainWorkbookPath $root 'Second.xlsx' | Should -Be (Join-Path $main 'Second.xlsx')
    }

    It 'does not silently use an archived report when the active workbook is missing' {
        $archive = Join-Path $main 'Archive'
        New-Item -ItemType Directory -Path $archive | Out-Null
        New-TestMainWorkbook -Path (Join-Path $archive 'GSS Score Trends - Main.xlsx')
        { Resolve-MainWorkbookPath $root 'GSS Score Trends - Main.xlsx' } | Should -Throw '*Archived copies are not selected*'
    }

    It 'does not mistake an unrelated Excel file for the main report' {
        New-TestMainWorkbook -Path (Join-Path $main 'Notes.xlsx') -Sheets @('Notes')
        { Resolve-MainWorkbookPath $root 'GSS Score Trends - Main.xlsx' } | Should -Throw '*Main workbook not found*'
    }

    It 'does not search other archives when the entire main-workbook folder is absent' {
        $legacyRoot = Join-Path $TestDrive ('legacy-' + [guid]::NewGuid().ToString('N'))
        $archive = Join-Path $legacyRoot '05 Reference Materials\Archive'
        New-Item -ItemType Directory -Path $archive -Force | Out-Null
        New-TestMainWorkbook -Path (Join-Path $archive 'GSS Score Trends - Main.xlsx')
        { Resolve-MainWorkbookPath $legacyRoot 'GSS Score Trends - Main.xlsx' } | Should -Throw '*Archived copies are not selected*'
    }

    It 'keeps technical archives out of recursive business-input discovery' {
        Test-ExcludedGssPath (Join-Path $root '_automation_runs\archive\program-releases\old.xlsx') $root | Should -BeTrue
    }
}
