function Normalize-Header {
    param([object]$Value)
    if ($null -eq $Value) { return '' }
    return ([regex]::Replace(([string]$Value).ToLowerInvariant(), '[^a-z0-9]+', '')).Trim()
}

function Convert-ExcelDate {
    param([object]$Value)
    if ($null -eq $Value) { return $null }
    if ($Value -is [datetime]) { return $Value.Date }
    if ($Value -is [double] -or $Value -is [int]) { return ([datetime]::FromOADate([double]$Value)).Date }
    return ([datetime]::Parse([string]$Value)).Date
}

function Convert-ToNullableDouble {
    param([object]$Value)
    if ($null -eq $Value) { return $null }
    if ($Value -is [string] -and [string]::IsNullOrWhiteSpace($Value)) { return $null }
    return [double]$Value
}

function Convert-ToBool {
    param([object]$Value)
    if ($null -eq $Value) { return $false }
    if ($Value -is [bool]) { return $Value }
    if ($Value -is [double] -or $Value -is [int]) { return ([double]$Value) -ne 0 }
    return ([string]$Value).Trim() -match '^(true|yes|1)$'
}

function Test-ExcludedGssPath {
    param([string]$Path, [string]$FolderPath)

    $resolvedPath = [System.IO.Path]::GetFullPath($Path).TrimEnd('\')
    $excludedFolders = @(
        (Join-Path $FolderPath 'GSS Survey Workbook Automation'),
        (Join-Path $FolderPath '_automation_runs')
    )

    foreach ($excludedFolder in $excludedFolders) {
        $resolvedExcluded = [System.IO.Path]::GetFullPath($excludedFolder).TrimEnd('\')
        if ($resolvedPath.Equals($resolvedExcluded, [System.StringComparison]::OrdinalIgnoreCase) -or
            $resolvedPath.StartsWith("$resolvedExcluded\", [System.StringComparison]::OrdinalIgnoreCase)) {
            return $true
        }
    }

    return $false
}

function Resolve-GssChildFolder {
    param([string]$FolderPath, [string]$ChildFolderName)

    $candidate = Join-Path $FolderPath $ChildFolderName
    if (Test-Path -LiteralPath $candidate -PathType Container) {
        return (Resolve-Path -LiteralPath $candidate).Path
    }

    return $null
}

function Get-GssFiles {
    param([string]$FolderPath, [string]$Filter)

    Get-ChildItem -LiteralPath $FolderPath -File -Filter $Filter -Recurse |
        Where-Object { -not (Test-ExcludedGssPath $_.FullName $FolderPath) }
}

function Test-GssMainWorkbookStructure {
    param([Parameter(Mandatory)][string]$Path)

    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $archive = $null
    $reader = $null
    try {
        $archive = [System.IO.Compression.ZipFile]::OpenRead($Path)
        $entry = $archive.GetEntry('xl/workbook.xml')
        if (-not $entry) { return $false }
        $reader = New-Object System.IO.StreamReader($entry.Open())
        $document = New-Object System.Xml.XmlDocument
        $document.XmlResolver = $null
        $document.LoadXml($reader.ReadToEnd())
        $names = @($document.SelectNodes('//*[local-name()="sheet"]') | ForEach-Object { $_.GetAttribute('name') })
        return @('Raw_Data', 'QA Checks', 'Email Comparison' | Where-Object { $_ -notin $names }).Count -eq 0
    }
    catch {
        return $false
    }
    finally {
        if ($reader) { $reader.Dispose() }
        if ($archive) { $archive.Dispose() }
    }
}

function Resolve-MainWorkbookPath {
    param([string]$FolderPath, [string]$WorkbookName)

    $mainWorkbookFolder = Resolve-GssChildFolder $FolderPath '01 Main Workbook'
    if ($mainWorkbookFolder) {
        $preferredPath = Join-Path $mainWorkbookFolder $WorkbookName
        if (Test-Path -LiteralPath $preferredPath) {
            return (Resolve-Path -LiteralPath $preferredPath).Path
        }

        if ($WorkbookName -eq 'GSS Score Trends - Main.xlsx') {
            $candidates = @(Get-ChildItem -LiteralPath $mainWorkbookFolder -File -Filter '*.xlsx' |
                Where-Object { -not $_.Name.StartsWith('~$') -and (Test-GssMainWorkbookStructure -Path $_.FullName) })
            if ($candidates.Count -eq 1) { return $candidates[0].FullName }
            if ($candidates.Count -gt 1) {
                throw "Multiple score-trends workbooks are present. Select one with -MainWorkbookName. Candidates: $(($candidates.Name) -join '; ')"
            }
        }
        throw "Main workbook not found in '$mainWorkbookFolder': $WorkbookName. Restore the workbook or select its filename with -MainWorkbookName. Archived copies are not selected automatically."
    }

    $directPath = Join-Path $FolderPath $WorkbookName
    if (Test-Path -LiteralPath $directPath) {
        return (Resolve-Path -LiteralPath $directPath).Path
    }

    Write-Warning "Main workbook was not found in '01 Main Workbook'. Falling back to a recursive compatibility search under $FolderPath."
    $matches = @(Get-GssFiles $FolderPath $WorkbookName)
    if ($matches.Count -eq 1) {
        return $matches[0].FullName
    }
    if ($matches.Count -gt 1) {
        $paths = ($matches | ForEach-Object { $_.FullName }) -join '; '
        throw "Multiple main workbook matches were found. Move old copies out of the GSS Surveys folder or pass -MainWorkbookName. Matches: $paths"
    }

    throw "Main workbook not found under ${FolderPath}: $WorkbookName"
}

function ConvertTo-GssDropboxRelativePath {
    param(
        [Parameter(Mandatory)]
        [string]$Path,
        [Parameter(Mandatory)]
        [string]$FolderPath
    )

    $root = [System.IO.Path]::GetFullPath($FolderPath).TrimEnd('\', '/')
    $candidate = $Path.Trim().Trim('"')
    if ([string]::IsNullOrWhiteSpace($candidate)) {
        throw 'A portable GSS path cannot be blank.'
    }

    if (-not [System.IO.Path]::IsPathRooted($candidate)) {
        $relative = $candidate.Replace('/', '\').TrimStart('\')
    }
    else {
        $fullCandidate = [System.IO.Path]::GetFullPath($candidate)
        if ($fullCandidate.Equals($root, [System.StringComparison]::OrdinalIgnoreCase)) {
            $relative = ''
        }
        elseif ($fullCandidate.StartsWith("$root\", [System.StringComparison]::OrdinalIgnoreCase)) {
            $relative = $fullCandidate.Substring($root.Length + 1)
        }
        else {
            # Logs can be produced on another Windows profile and then synced by Dropbox.
            # Recover only the suffix below the canonical GSS Surveys folder.
            $marker = '\GSS Surveys\'
            $markerIndex = $fullCandidate.IndexOf($marker, [System.StringComparison]::OrdinalIgnoreCase)
            if ($markerIndex -lt 0) {
                throw "Path is outside the GSS Surveys Dropbox tree: $candidate"
            }
            $relative = $fullCandidate.Substring($markerIndex + $marker.Length)
        }
    }

    $segments = @($relative -split '[\\/]' | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    if (@($segments | Where-Object { $_ -eq '.' -or $_ -eq '..' }).Count -gt 0) {
        throw "Portable GSS path contains a traversal segment: $candidate"
    }
    return ($segments -join '/')
}

function Resolve-GssDropboxPath {
    param(
        [Parameter(Mandatory)]
        [string]$Path,
        [Parameter(Mandatory)]
        [string]$FolderPath,
        [switch]$RequireFile
    )

    $root = [System.IO.Path]::GetFullPath($FolderPath).TrimEnd('\', '/')
    $relative = ConvertTo-GssDropboxRelativePath -Path $Path -FolderPath $root
    $resolved = [System.IO.Path]::GetFullPath((Join-Path $root ($relative.Replace('/', '\'))))
    if (-not $resolved.StartsWith("$root\", [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "Resolved GSS path escaped the Dropbox root: $Path"
    }
    if ($RequireFile -and -not (Test-Path -LiteralPath $resolved -PathType Leaf)) {
        throw "GSS file is missing or not fully available: $relative"
    }
    return $resolved
}

function Get-GssSha256 {
    param([Parameter(Mandatory)][string]$Path)

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        throw "Cannot hash missing GSS file: $Path"
    }
    return (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash.ToLowerInvariant()
}
