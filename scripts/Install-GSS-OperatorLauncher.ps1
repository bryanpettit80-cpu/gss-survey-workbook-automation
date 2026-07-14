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
$assets = @(
    [pscustomobject]@{
        Name = 'Run GSS Update After Upload.cmd'
        Template = Join-Path $projectRoot 'templates\Run GSS Update After Upload.cmd'
        Target = Join-Path $Folder 'Run GSS Update After Upload.cmd'
    },
    [pscustomobject]@{
        Name = '00 START HERE - GSS Survey Updates.txt'
        Template = Join-Path $projectRoot 'templates\00 START HERE - GSS Survey Updates.txt'
        Target = Join-Path $Folder '00 START HERE - GSS Survey Updates.txt'
    },
    [pscustomobject]@{
        Name = '00 DROP NEW WEEKLY WORKBOOK HERE.txt'
        Template = Join-Path $projectRoot 'templates\00 DROP NEW WEEKLY WORKBOOK HERE.txt'
        Target = Join-Path $Folder '02 Weekly Rolling Source Workbooks\00 DROP NEW WEEKLY WORKBOOK HERE.txt'
    },
    [pscustomobject]@{
        Name = '00 ABOUT THIS FOLDER.txt'
        Template = Join-Path $projectRoot 'templates\00 ABOUT THIS FOLDER.txt'
        Target = Join-Path $Folder '03 Uploaded Survey Workbooks\00 ABOUT THIS FOLDER.txt'
    }
)

foreach ($asset in $assets) {
    if (-not (Test-Path -LiteralPath $asset.Template)) {
        throw "Operator file template not found: $($asset.Template)"
    }

    Copy-Item -LiteralPath $asset.Template -Destination $asset.Target -Force

    $templateHash = (Get-FileHash -LiteralPath $asset.Template -Algorithm SHA256).Hash
    $targetHash = (Get-FileHash -LiteralPath $asset.Target -Algorithm SHA256).Hash
    if ($templateHash -ne $targetHash) {
        throw "Operator file refresh failed; target does not match template: $($asset.Target)"
    }

    Write-Host "Operator file refreshed: $($asset.Target)"
}
