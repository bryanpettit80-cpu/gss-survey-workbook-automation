[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSCommandPath)
$scripts = Get-ChildItem -LiteralPath $repoRoot -Recurse -File |
    Where-Object { $_.Extension -ieq '.ps1' }
$failures = @()

foreach ($script in $scripts) {
    $tokens = $null
    $errors = $null
    [System.Management.Automation.Language.Parser]::ParseFile($script.FullName, [ref]$tokens, [ref]$errors) | Out-Null
    foreach ($errorRecord in $errors) {
        $failures += [pscustomobject]@{
            File = $script.FullName
            Line = $errorRecord.Extent.StartLineNumber
            Column = $errorRecord.Extent.StartColumnNumber
            Message = $errorRecord.Message
        }
    }
}

if ($failures.Count -gt 0) {
    $failures | Format-Table -AutoSize
    throw "PowerShell parser check failed for $($failures.Count) error(s)."
}

Write-Host "PowerShell parser check passed for $($scripts.Count) script(s)."
