[CmdletBinding()]
param(
    [string]$Folder
)

$ErrorActionPreference = 'Stop'

$scriptRoot = Split-Path -Parent $PSCommandPath
$projectRoot = Split-Path -Parent $scriptRoot

if ([string]::IsNullOrWhiteSpace($Folder)) {
    $Folder = Split-Path -Parent $projectRoot
}

$Folder = (Resolve-Path -LiteralPath $Folder).Path
$template = Join-Path $projectRoot 'templates\Run GSS Update After Upload.cmd'
$target = Join-Path $Folder 'Run GSS Update After Upload.cmd'

if (-not (Test-Path -LiteralPath $template)) {
    throw "Launcher template not found: $template"
}

Copy-Item -LiteralPath $template -Destination $target -Force

$templateHash = (Get-FileHash -LiteralPath $template -Algorithm SHA256).Hash
$targetHash = (Get-FileHash -LiteralPath $target -Algorithm SHA256).Hash
if ($templateHash -ne $targetHash) {
    throw "Launcher refresh failed; target does not match template: $target"
}

Write-Host "Operator launcher refreshed: $target"
