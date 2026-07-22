$ErrorActionPreference = 'Stop'

$scriptRoot = Split-Path -Parent $PSCommandPath
if (-not (Get-Command ConvertTo-GssDropboxRelativePath -ErrorAction SilentlyContinue)) {
    . (Join-Path $scriptRoot 'Gss-Common.ps1')
}

$script:GssEmailPackageSchemaVersion = 'gss-email-package/v1'
$script:GssAnalysisPolicyVersion = 'gss-analysis-policy/v1'
$script:GssFeedbackLedgerVersion = 'gss-feedback-first-seen/v1'
$script:GssThemeNames = @('service', 'culinary', 'pace', 'value', 'hospitality/recovery', 'recognition')

function Get-GssStringSha256 {
    param([Parameter(Mandatory)][string]$Value)

    $sha = [System.Security.Cryptography.SHA256]::Create()
    try {
        $bytes = [System.Text.Encoding]::UTF8.GetBytes($Value)
        return (($sha.ComputeHash($bytes) | ForEach-Object { $_.ToString('x2') }) -join '')
    }
    finally {
        $sha.Dispose()
    }
}

function Get-GssZipEntryText {
    param([object]$Archive, [string]$EntryName)

    $entry = $Archive.GetEntry($EntryName)
    if (-not $entry) { return $null }
    $stream = $entry.Open()
    try {
        $reader = New-Object System.IO.StreamReader($stream)
        try { return $reader.ReadToEnd() }
        finally { $reader.Dispose() }
    }
    finally { $stream.Dispose() }
}

function ConvertFrom-GssCellReference {
    param([string]$Reference)

    $letters = ([regex]::Match($Reference, '^[A-Z]+')).Value
    $column = 0
    foreach ($character in $letters.ToCharArray()) {
        $column = ($column * 26) + ([int][char]$character - [int][char]'A' + 1)
    }
    return ($column - 1)
}

function Normalize-GssFeedbackHeader {
    param([object]$Value)
    if ($null -eq $Value) { return '' }
    return [regex]::Replace(([string]$Value).ToLowerInvariant(), '[^a-z0-9]+', '')
}

function Normalize-GssFeedbackText {
    param([object]$Value)
    if ($null -eq $Value) { return '' }
    return ([regex]::Replace(([string]$Value).Trim().ToLowerInvariant(), '\s+', ' '))
}

function Get-GssRestaurantDisplayName {
    param([string]$Value)
    if ([string]::IsNullOrWhiteSpace($Value)) { return '' }
    return ([regex]::Replace($Value.Trim(), '^\d{4}\s+', '')).Trim()
}

function Read-GssXlsxFirstWorksheet {
    param([Parameter(Mandatory)][string]$Path)

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        throw "Detail workbook is missing: $Path"
    }
    if ((Get-Item -LiteralPath $Path).Length -le 0) {
        throw "Detail workbook is empty or not fully synced: $Path"
    }

    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $archive = $null
    try {
        $archive = [System.IO.Compression.ZipFile]::OpenRead($Path)
        [xml]$workbookXml = Get-GssZipEntryText $archive 'xl/workbook.xml'
        [xml]$relationshipXml = Get-GssZipEntryText $archive 'xl/_rels/workbook.xml.rels'
        if (-not $workbookXml -or -not $relationshipXml) {
            throw 'The workbook package is missing its workbook metadata.'
        }

        $workbookNs = New-Object System.Xml.XmlNamespaceManager($workbookXml.NameTable)
        $workbookNs.AddNamespace('m', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main')
        $workbookNs.AddNamespace('r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')
        $sheet = $workbookXml.SelectSingleNode('//m:sheets/m:sheet[1]', $workbookNs)
        if (-not $sheet) { throw 'The detail workbook has no worksheet.' }

        $relationshipMap = @{}
        foreach ($relationship in $relationshipXml.Relationships.Relationship) {
            $relationshipMap[[string]$relationship.Id] = [string]$relationship.Target
        }
        $relationshipId = $sheet.GetAttribute('id', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')
        $target = $relationshipMap[$relationshipId]
        if ([string]::IsNullOrWhiteSpace($target)) { throw 'The first worksheet relationship is missing.' }
        $entryName = if ($target.StartsWith('/')) { $target.TrimStart('/') } else { 'xl/' + $target.TrimStart('/') }

        $sharedStrings = @()
        $sharedText = Get-GssZipEntryText $archive 'xl/sharedStrings.xml'
        if ($sharedText) {
            [xml]$sharedXml = $sharedText
            $sharedNs = New-Object System.Xml.XmlNamespaceManager($sharedXml.NameTable)
            $sharedNs.AddNamespace('m', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main')
            foreach ($item in $sharedXml.SelectNodes('//m:si', $sharedNs)) {
                $parts = @($item.SelectNodes('.//m:t', $sharedNs) | ForEach-Object { $_.InnerText })
                $sharedStrings += ($parts -join '')
            }
        }

        [xml]$worksheetXml = Get-GssZipEntryText $archive $entryName
        if (-not $worksheetXml) { throw 'The first worksheet XML is missing.' }
        $worksheetNs = New-Object System.Xml.XmlNamespaceManager($worksheetXml.NameTable)
        $worksheetNs.AddNamespace('m', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main')
        $rowNodes = @($worksheetXml.SelectNodes('//m:sheetData/m:row', $worksheetNs))
        if ($rowNodes.Count -lt 2) { throw 'The detail workbook has no response rows.' }

        $parsedRows = @()
        foreach ($rowNode in $rowNodes) {
            $values = @{}
            foreach ($cell in $rowNode.SelectNodes('./m:c', $worksheetNs)) {
                $columnIndex = ConvertFrom-GssCellReference ([string]$cell.r)
                $valueNode = $cell.SelectSingleNode('./m:v', $worksheetNs)
                $inlineNode = $cell.SelectSingleNode('./m:is', $worksheetNs)
                $value = ''
                if ([string]$cell.t -eq 's' -and $valueNode) {
                    $value = $sharedStrings[[int]$valueNode.InnerText]
                }
                elseif ([string]$cell.t -eq 'inlineStr' -and $inlineNode) {
                    $value = (@($inlineNode.SelectNodes('.//m:t', $worksheetNs) | ForEach-Object { $_.InnerText }) -join '')
                }
                elseif ($valueNode) {
                    $value = $valueNode.InnerText
                }
                $values[$columnIndex] = [string]$value
            }
            $parsedRows += ,$values
        }

        $headers = @()
        $headerMax = ($parsedRows[0].Keys | Measure-Object -Maximum).Maximum
        for ($columnIndex = 0; $columnIndex -le $headerMax; $columnIndex++) {
            $headers += if ($parsedRows[0].ContainsKey($columnIndex)) { [string]$parsedRows[0][$columnIndex] } else { '' }
        }
        $normalizedHeaders = @($headers | ForEach-Object { Normalize-GssFeedbackHeader $_ })
        $duplicateHeaders = @($normalizedHeaders | Where-Object { $_ } | Group-Object | Where-Object { $_.Count -gt 1 })
        if ($duplicateHeaders.Count -gt 0) {
            throw "Detail workbook contains duplicate normalized header: $($duplicateHeaders[0].Name)"
        }

        $records = @()
        for ($rowIndex = 1; $rowIndex -lt $parsedRows.Count; $rowIndex++) {
            $sourceRow = $parsedRows[$rowIndex]
            $record = [ordered]@{}
            $hasContent = $false
            for ($columnIndex = 0; $columnIndex -lt $headers.Count; $columnIndex++) {
                $normalizedHeader = $normalizedHeaders[$columnIndex]
                if ([string]::IsNullOrWhiteSpace($normalizedHeader)) { continue }
                $cellValue = if ($sourceRow.ContainsKey($columnIndex)) { [string]$sourceRow[$columnIndex] } else { '' }
                if (-not [string]::IsNullOrWhiteSpace($cellValue)) { $hasContent = $true }
                $record[$normalizedHeader] = $cellValue
            }
            if ($hasContent) { $records += [pscustomobject]$record }
        }

        return [pscustomobject]@{
            SheetName = [string]$sheet.name
            Headers = $headers
            NormalizedHeaders = $normalizedHeaders
            Records = $records
        }
    }
    catch {
        throw "Detail workbook is corrupt or unsupported ($Path): $($_.Exception.Message)"
    }
    finally {
        if ($archive) { $archive.Dispose() }
    }
}

function ConvertTo-GssFeedbackDate {
    param([object]$Value)

    if ($null -eq $Value -or [string]::IsNullOrWhiteSpace([string]$Value)) { return $null }
    $number = 0.0
    if ([double]::TryParse([string]$Value, [ref]$number) -and $number -gt 1000) {
        return [datetime]::FromOADate($number).Date
    }
    $date = [datetime]::MinValue
    if ([datetime]::TryParse([string]$Value, [ref]$date)) { return $date.Date }
    return $null
}

function Get-GssFeedbackProperty {
    param([object]$Record, [string]$Name)
    $property = $Record.PSObject.Properties[$Name]
    if ($property) { return [string]$property.Value }
    return ''
}

function Test-GssFeedbackAnswers {
    param([object]$Record, [string]$SourceLabel)

    $ranges = [ordered]@{
        overall = @(1, 5)
        service = @(1, 5)
        culinary = @(1, 5)
        value = @(1, 5)
        paceofmeal = @(1, 5)
        recommend = @(0, 10)
        eventbookingprocess = @(1, 5)
    }
    foreach ($name in $ranges.Keys) {
        $raw = Get-GssFeedbackProperty $Record $name
        if ([string]::IsNullOrWhiteSpace($raw)) { continue }
        $number = 0.0
        if (-not [double]::TryParse($raw, [ref]$number) -or $number -lt $ranges[$name][0] -or $number -gt $ranges[$name][1]) {
            throw "Invalid $name answer '$raw' in $SourceLabel."
        }
    }
}

function Read-GssDetailWorkbook {
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$FolderPath
    )

    $table = Read-GssXlsxFirstWorksheet $Path
    foreach ($required in @('restaurantname', 'reservationdate', 'text')) {
        if (-not $table.NormalizedHeaders.Contains($required)) {
            throw "Detail workbook is missing required header '$required': $Path"
        }
    }

    $portablePath = ConvertTo-GssDropboxRelativePath -Path $Path -FolderPath $FolderPath
    $responses = @()
    $rowNumber = 1
    foreach ($record in $table.Records) {
        $rowNumber++
        $restaurant = (Get-GssFeedbackProperty $record 'restaurantname').Trim()
        $visitDate = ConvertTo-GssFeedbackDate (Get-GssFeedbackProperty $record 'reservationdate')
        if ([string]::IsNullOrWhiteSpace($restaurant) -and $null -eq $visitDate -and [string]::IsNullOrWhiteSpace((Get-GssFeedbackProperty $record 'text'))) {
            continue
        }
        if ([string]::IsNullOrWhiteSpace($restaurant) -or $null -eq $visitDate) {
            throw "Detail workbook contains an incomplete response at row ${rowNumber}: $portablePath"
        }
        Test-GssFeedbackAnswers $record "$portablePath row $rowNumber"

        $restaurantIdMatch = [regex]::Match($restaurant, '^\s*(\d{4})\b')
        $restaurantId = if ($restaurantIdMatch.Success) { $restaurantIdMatch.Groups[1].Value } else { Normalize-GssFeedbackHeader $restaurant }
        $time = (Get-GssFeedbackProperty $record 'reservationtime').Trim()
        $text = (Get-GssFeedbackProperty $record 'text').Trim()
        $answerNames = @('overall', 'service', 'culinary', 'value', 'paceofmeal', 'recommend', 'managervisit', 'steakcookedcorrectly', 'eventbookingprocess', 'firstvisit')
        $answers = [ordered]@{}
        foreach ($answerName in $answerNames) { $answers[$answerName] = (Get-GssFeedbackProperty $record $answerName).Trim() }
        $timestampKey = "$($visitDate.ToString('yyyy-MM-dd')) $((Normalize-GssFeedbackText $time))".Trim()
        $identity = @(
            (Normalize-GssFeedbackText $restaurant),
            $timestampKey,
            (Normalize-GssFeedbackText $text),
            (($answerNames | ForEach-Object { Normalize-GssFeedbackText $answers[$_] }) -join '|')
        ) -join "`n"
        $dncValue = (Get-GssFeedbackProperty $record 'alertguestsdonotcontact').Trim()
        $doNotContact = (-not [string]::IsNullOrWhiteSpace($dncValue)) -and ($dncValue -notmatch '^(0|n|no|false)$')

        $responses += [pscustomobject]@{
            ResponseHash = Get-GssStringSha256 $identity
            RestaurantId = $restaurantId
            Restaurant = $restaurant
            VisitDate = $visitDate
            ReservationTime = $time
            Text = $text
            Answers = [pscustomobject]$answers
            DoNotContact = $doNotContact
            GuestFirstName = (Get-GssFeedbackProperty $record 'guestfirstname').Trim()
            GuestLastName = (Get-GssFeedbackProperty $record 'guestlastname').Trim()
            SourcePath = $portablePath
            SourceRow = $rowNumber
        }
    }
    if ($responses.Count -eq 0) { throw "Detail workbook contains no usable responses: $portablePath" }

    return [pscustomobject]@{
        Path = $Path
        PortablePath = $portablePath
        HeaderCount = @($table.NormalizedHeaders | Where-Object { $_ }).Count
        Headers = $table.NormalizedHeaders
        Responses = $responses
        VisitDateStart = ($responses.VisitDate | Measure-Object -Minimum).Minimum
        VisitDateEnd = ($responses.VisitDate | Measure-Object -Maximum).Maximum
    }
}

function Get-GssKnownGuestNames {
    param([object[]]$Responses)

    $names = @()
    foreach ($response in $Responses) {
        foreach ($value in @($response.GuestFirstName, $response.GuestLastName)) {
            if ([string]::IsNullOrWhiteSpace($value)) { continue }
            $names += $value.Trim()
            $names += @([regex]::Matches($value, "[A-Za-z][A-Za-z'-]{1,}") | ForEach-Object { $_.Value })
        }
    }
    $uniqueNames = @($names | Where-Object { $_.Length -ge 2 } | Sort-Object -Unique)
    return @($uniqueNames | Sort-Object `
        @{ Expression = { $_.Length }; Descending = $true }, `
        @{ Expression = { $_ }; Descending = $false })
}

function Get-GssPiiRedactionRules {
    return @(
        [pscustomobject]@{ Label = 'email'; Pattern = '(?i)\b[A-Z0-9._%+\-]+@[A-Z0-9.\-]+\.[A-Z]{2,}\b'; Replacement = '[REDACTED EMAIL]' },
        [pscustomobject]@{ Label = 'url'; Pattern = '(?i)(?<![@\w])(?<url>(?:(?:https?://|www\.)[^\s<>"\x27]*?|(?:[a-z0-9](?:[a-z0-9-]{0,61}[a-z0-9])?\.)+[a-z]{2,63}(?:/[^\s<>"\x27]*?)?))(?<punct>[.,;:!?)}\]]*)(?=\s|$)'; Replacement = '[REDACTED URL]${punct}' },
        [pscustomobject]@{ Label = 'phone'; Pattern = '(?<![\w\d.])(?:(?:\+?1[\s.-]+)?[2-9]\d{2}\.[2-9]\d{6}|(?!\d+\.\d+(?![\w\d]|\.\d))(?:\+?\d{10,15}|(?:\+?\d{1,3}[\s.-]+)?(?:\(\d{2,4}\)[\s.-]*|\d{2,4}[\s.-]+)\d{2,4}[\s.-]*\d{4}|\d{3}[- ]\d{4}))(?![\w\d]|\.\d)'; Replacement = '[REDACTED PHONE]' },
        [pscustomobject]@{ Label = 'booking_identifier'; Pattern = '(?i)\b(?:check|confirmation|booking|reservation|resy)\s*(?:(?:id|number|no\.?)\s*)?(?:#|:)?\s*(?=[A-Z0-9-]{4,}\b)(?=[A-Z0-9-]*\d)[A-Z0-9-]{4,}\b'; Replacement = '[REDACTED BOOKING ID]' }
    )
}

function Get-GssUnsafeControlPattern {
    # Permit normal horizontal tab and line endings while rejecting other C0/C1
    # controls and Unicode bidi controls that can disguise package text.
    return '[\x00-\x08\x0B\x0C\x0E-\x1F\x7F-\x9F\u061C\u200E\u200F\u202A-\u202E\u2066-\u2069]'
}

function Get-GssRemainingPiiTypes {
    param([string]$Text, [string[]]$KnownNames)

    $remainingPii = @()
    foreach ($rule in @(Get-GssPiiRedactionRules)) {
        if ([regex]::IsMatch([string]$Text, $rule.Pattern)) { $remainingPii += $rule.Label }
    }
    if ([regex]::IsMatch([string]$Text, (Get-GssUnsafeControlPattern))) { $remainingPii += 'unsafe_control' }
    foreach ($name in @($KnownNames)) {
        if ([string]::IsNullOrWhiteSpace($name)) { continue }
        if ([regex]::IsMatch([string]$Text, '(?i)(?<![A-Za-z])' + [regex]::Escape($name) + '(?![A-Za-z])')) {
            $remainingPii += 'known_name'
        }
    }
    return @($remainingPii | Select-Object -Unique)
}

function Protect-GssFeedbackText {
    param([string]$Text, [string[]]$KnownNames)

    $sanitized = if ($null -eq $Text) { '' } else { [string]$Text }
    $redactionCount = 0
    $controlMatches = [regex]::Matches($sanitized, (Get-GssUnsafeControlPattern))
    $redactionCount += $controlMatches.Count
    $sanitized = [regex]::Replace($sanitized, (Get-GssUnsafeControlPattern), '[REDACTED CONTROL]')
    $rules = @(Get-GssPiiRedactionRules)
    foreach ($rule in $rules) {
        $matches = [regex]::Matches($sanitized, $rule.Pattern)
        $redactionCount += $matches.Count
        $sanitized = [regex]::Replace($sanitized, $rule.Pattern, $rule.Replacement)
    }
    foreach ($name in @($KnownNames)) {
        if ([string]::IsNullOrWhiteSpace($name)) { continue }
        $pattern = '(?i)(?<![A-Za-z])' + [regex]::Escape($name) + '(?![A-Za-z])'
        $matches = [regex]::Matches($sanitized, $pattern)
        $redactionCount += $matches.Count
        $sanitized = [regex]::Replace($sanitized, $pattern, '[REDACTED NAME]')
    }
    $sanitized = [regex]::Replace($sanitized, '\s+', ' ').Trim()

    $remainingPii = @(Get-GssRemainingPiiTypes -Text $sanitized -KnownNames $KnownNames)

    return [pscustomobject]@{
        Text = $sanitized
        RedactionCount = $redactionCount
        PiiScanPassed = ($remainingPii.Count -eq 0)
        RemainingPiiTypes = @($remainingPii | Select-Object -Unique)
    }
}

function Get-GssFeedbackThemes {
    param([object]$Response, [string]$SanitizedText)

    $text = Normalize-GssFeedbackText $SanitizedText
    $themes = @()
    $patterns = [ordered]@{
        service = '\b(service|server|waiter|waitress|staff|attentive|courteous|professional)\b'
        culinary = '\b(food|steak|meal|dish|salad|potato|crab|lamb|cook|cooked|temperature|delicious)\b'
        pace = '\b(wait|waiting|slow|rushed|timing|delay|minutes|pace)\b'
        value = '\b(value|price|priced|cost|expensive|worth|portion|portions)\b'
        'hospitality/recovery' = '\b(manager|management|apology|apologize|resolve|resolved|recovery|follow.?up|care|complaint)\b'
        recognition = '\b(birthday|anniversary|celebration|occasion|recognition|recognize|acknowledge|memorable)\b'
    }
    foreach ($theme in $patterns.Keys) {
        if ($text -match $patterns[$theme]) { $themes += $theme }
    }
    $answerThemeMap = [ordered]@{ service = 'service'; culinary = 'culinary'; paceofmeal = 'pace'; value = 'value' }
    foreach ($answerName in $answerThemeMap.Keys) {
        $raw = Get-GssFeedbackProperty $Response.Answers $answerName
        $number = 0.0
        if ([double]::TryParse($raw, [ref]$number) -and $number -le 3) { $themes += $answerThemeMap[$answerName] }
    }
    $overall = 0.0
    $manager = Get-GssFeedbackProperty $Response.Answers 'managervisit'
    if ([double]::TryParse((Get-GssFeedbackProperty $Response.Answers 'overall'), [ref]$overall) -and $overall -le 3 -and $manager -match '^(?i:n|no)$') {
        $themes += 'hospitality/recovery'
    }
    return @($themes | Select-Object -Unique)
}

function Get-GssFeedbackSentiment {
    param([object]$Response, [string]$SanitizedText)

    $text = Normalize-GssFeedbackText $SanitizedText
    $concern = $text -match '\b(disappoint|disappointed|poor|cold|overcooked|undercooked|rude|slow|expensive|complaint|worst|not good|burnt|scorched)\b'
    $positive = $text -match '\b(excellent|wonderful|great|outstanding|delicious|fantastic|perfect|memorable|exceptional)\b'
    foreach ($name in @('overall', 'service', 'culinary', 'value', 'paceofmeal')) {
        $number = 0.0
        if ([double]::TryParse((Get-GssFeedbackProperty $Response.Answers $name), [ref]$number)) {
            if ($number -le 3) { $concern = $true }
            if ($number -ge 4) { $positive = $true }
        }
    }
    if ($concern -and $positive) { return 'mixed' }
    if ($concern) { return 'concern' }
    if ($positive) { return 'positive' }
    return 'neutral'
}

function Read-GssFeedbackLedger {
    param([string]$Path)

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        return [pscustomobject]@{ schema_version = $script:GssFeedbackLedgerVersion; entries = @() }
    }
    try {
        $ledger = Get-Content -Raw -LiteralPath $Path | ConvertFrom-Json
        if ($ledger.schema_version -ne $script:GssFeedbackLedgerVersion) {
            throw "Unsupported feedback ledger version: $($ledger.schema_version)"
        }
        return $ledger
    }
    catch {
        throw "Feedback ledger is corrupt: $($_.Exception.Message)"
    }
}

function Invoke-GssFileSystemRetry {
    param(
        [Parameter(Mandatory)][scriptblock]$Operation,
        [int]$MaxAttempts = 12,
        [int]$DelayMilliseconds = 500
    )

    if ($MaxAttempts -lt 1) { throw 'MaxAttempts must be at least 1.' }
    if ($DelayMilliseconds -lt 0) { throw 'DelayMilliseconds cannot be negative.' }

    $lastError = $null
    for ($attempt = 1; $attempt -le $MaxAttempts; $attempt++) {
        try {
            return (& $Operation)
        }
        catch {
            $lastError = $_
            if ($attempt -lt $MaxAttempts -and $DelayMilliseconds -gt 0) {
                Start-Sleep -Milliseconds $DelayMilliseconds
            }
        }
    }
    throw $lastError
}

function Publish-GssStagedEmailPackage {
    param(
        [Parameter(Mandatory)][string]$StagingPath,
        [Parameter(Mandatory)][string]$PackagePath,
        [Parameter(Mandatory)][scriptblock]$ValidationOperation,
        [int]$MaxAttempts = 12,
        [int]$DelayMilliseconds = 500
    )

    $resolvedStaging = [System.IO.Path]::GetFullPath($StagingPath)
    $resolvedPackage = [System.IO.Path]::GetFullPath($PackagePath)
    if ([string]::Equals($resolvedStaging, $resolvedPackage, [System.StringComparison]::OrdinalIgnoreCase)) {
        throw 'Staging and package directories must be different.'
    }
    $stagingParent = [System.IO.Path]::GetDirectoryName($resolvedStaging)
    $packageParent = [System.IO.Path]::GetDirectoryName($resolvedPackage)
    if (-not [string]::Equals($stagingParent, $packageParent, [System.StringComparison]::OrdinalIgnoreCase)) {
        throw 'Staging and package directories must share the same outbox parent.'
    }

    $promoted = $false
    try {
        Invoke-GssFileSystemRetry -MaxAttempts $MaxAttempts -DelayMilliseconds $DelayMilliseconds -Operation {
            # Same-parent paths guarantee a same-volume rename. Directory.Move
            # fails if another publisher created the destination first instead
            # of nesting this staging directory inside that publisher's package.
            [System.IO.Directory]::Move($resolvedStaging, $resolvedPackage)
        }
        $promoted = $true
        return (Invoke-GssFileSystemRetry -MaxAttempts $MaxAttempts -DelayMilliseconds $DelayMilliseconds -Operation $ValidationOperation)
    }
    catch {
        $publicationError = $_
        $cleanupPath = if ($promoted) { $resolvedPackage } else { $resolvedStaging }
        if (Test-Path -LiteralPath $cleanupPath -PathType Container) {
            try {
                Invoke-GssFileSystemRetry -MaxAttempts $MaxAttempts -DelayMilliseconds $DelayMilliseconds -Operation {
                    Remove-Item -LiteralPath $cleanupPath -Recurse -Force
                }
            }
            catch {
                $key = if ($promoted) { 'GssPromotedPackageCleanupError' } else { 'GssStagingCleanupError' }
                $publicationError.Exception.Data[$key] = $_.Exception.Message
            }
        }
        throw
    }
}

function Write-GssFeedbackLedger {
    param([object]$Ledger, [string]$Path)

    $directory = Split-Path -Parent $Path
    New-Item -ItemType Directory -Path $directory -Force | Out-Null
    $temporary = "$Path.tmp.$([guid]::NewGuid().ToString('N'))"
    try {
        $Ledger | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $temporary -Encoding UTF8
        Move-Item -LiteralPath $temporary -Destination $Path -Force
    }
    finally {
        if (Test-Path -LiteralPath $temporary) { Remove-Item -LiteralPath $temporary -Force }
    }
}

function Get-GssDetailInventory {
    param(
        [Parameter(Mandatory)][string]$FolderPath,
        [Parameter(Mandatory)][datetime]$ReportingDate
    )

    $detailFolder = Join-Path $FolderPath '03 Uploaded Survey Workbooks'
    if (-not (Test-Path -LiteralPath $detailFolder -PathType Container)) {
        throw "Guest-detail folder is missing: 03 Uploaded Survey Workbooks"
    }
    $files = @(Get-ChildItem -LiteralPath $detailFolder -File -Filter '*.xlsx' -Recurse |
        Where-Object { $_.Name -notlike '~$*' -and $_.Length -gt 0 } |
        Sort-Object FullName)
    if ($files.Count -eq 0) { throw 'No guest-detail workbooks were found.' }

    $workbooks = @()
    foreach ($file in $files) {
        $workbooks += Read-GssDetailWorkbook -Path $file.FullName -FolderPath $FolderPath
    }

    $directPrefix = [System.IO.Path]::GetFullPath($detailFolder).TrimEnd('\') + '\'
    $direct = @($workbooks | Where-Object {
        $parent = [System.IO.Path]::GetDirectoryName([System.IO.Path]::GetFullPath($_.Path)).TrimEnd('\') + '\'
        $parent.Equals($directPrefix, [System.StringComparison]::OrdinalIgnoreCase)
    })
    if ($direct.Count -eq 0) { throw 'No current guest-detail workbook exists directly in 03 Uploaded Survey Workbooks.' }
    $eligible = @($direct | Where-Object { $_.VisitDateEnd.Date -le $ReportingDate.Date })
    if ($eligible.Count -eq 0) {
        throw "No current guest-detail workbook has a visit date on or before reporting date $($ReportingDate.ToString('yyyy-MM-dd'))."
    }
    $current = $eligible |
        Sort-Object @{ Expression = { $_.VisitDateEnd }; Descending = $true }, @{ Expression = { $_.PortablePath }; Descending = $false } |
        Select-Object -First 1

    $groups = @{}
    foreach ($workbook in $workbooks) {
        foreach ($response in $workbook.Responses) {
            if (-not $groups.ContainsKey($response.ResponseHash)) { $groups[$response.ResponseHash] = @() }
            $groups[$response.ResponseHash] += $response
        }
    }
    $uniqueResponses = @()
    foreach ($hash in @($groups.Keys | Sort-Object)) {
        $instances = @($groups[$hash] | Sort-Object SourcePath, SourceRow)
        $canonical = $instances[0]
        $uniqueResponses += [pscustomobject]@{
            ResponseHash = $hash
            RestaurantId = $canonical.RestaurantId
            Restaurant = $canonical.Restaurant
            VisitDate = $canonical.VisitDate
            Text = $canonical.Text
            Answers = $canonical.Answers
            DoNotContact = [bool](@($instances | Where-Object { $_.DoNotContact }).Count -gt 0)
            GuestFirstName = $canonical.GuestFirstName
            GuestLastName = $canonical.GuestLastName
            SourcePaths = @($instances.SourcePath | Sort-Object -Unique)
        }
    }

    return [pscustomobject]@{
        CurrentWorkbook = $current
        Workbooks = $workbooks
        AllResponseInstances = @($workbooks.Responses)
        UniqueResponses = $uniqueResponses
        DuplicateResponseCount = (@($workbooks.Responses).Count - $uniqueResponses.Count)
    }
}

function Get-GssFeedbackSelection {
    param(
        [Parameter(Mandatory)][object]$Inventory,
        [Parameter(Mandatory)][object]$Ledger,
        [Parameter(Mandatory)][datetime]$ReportingDate
    )

    $reportingDateText = $ReportingDate.ToString('yyyy-MM-dd')
    $ledgerByHash = @{}
    foreach ($entry in @($Ledger.entries)) { $ledgerByHash[[string]$entry.response_hash] = $entry }
    $isInitialLedger = ($ledgerByHash.Count -eq 0)
    $currentPath = $Inventory.CurrentWorkbook.PortablePath
    $selected = @()
    foreach ($response in @($Inventory.UniqueResponses | Sort-Object ResponseHash)) {
        if ($response.VisitDate.Date -gt $ReportingDate.Date) { continue }
        $existing = $ledgerByHash[$response.ResponseHash]
        $inCurrentAttachment = @($response.SourcePaths | Where-Object { $_ -eq $currentPath }).Count -gt 0
        if ($existing) {
            if ($existing.first_seen_package_id -ne 'baseline' -and $existing.first_seen_reporting_date -eq $reportingDateText) { $selected += $response }
        }
        elseif (-not $isInitialLedger -or $inCurrentAttachment) {
            $selected += $response
        }
    }
    $canonicalHashes = @($selected.ResponseHash | Sort-Object -Unique)
    $fingerprintMaterial = "gss-feedback-selection/v1`n" + ($canonicalHashes -join "`n")
    return [pscustomobject]@{
        Responses = $selected
        ResponseHashes = $canonicalHashes
        Fingerprint = Get-GssStringSha256 $fingerprintMaterial
        LedgerByHash = $ledgerByHash
        IsInitialLedger = $isInitialLedger
    }
}

function New-GssSanitizedFeedback {
    param(
        [Parameter(Mandatory)][object]$Inventory,
        [Parameter(Mandatory)][object]$Ledger,
        [Parameter(Mandatory)][string]$PackageId,
        [Parameter(Mandatory)][datetime]$ReportingDate,
        [object]$FeedbackSelection
    )

    $reportingDateText = $ReportingDate.ToString('yyyy-MM-dd')
    $selection = if ($FeedbackSelection) { $FeedbackSelection } else { Get-GssFeedbackSelection -Inventory $Inventory -Ledger $Ledger -ReportingDate $ReportingDate }
    $ledgerByHash = $selection.LedgerByHash
    $isInitialLedger = [bool]$selection.IsInitialLedger
    $currentPath = $Inventory.CurrentWorkbook.PortablePath
    $selected = @($selection.Responses)

    $knownNames = @(Get-GssKnownGuestNames $Inventory.AllResponseInstances)
    $internalCards = @()
    $piiRedactions = 0
    $piiFailures = @()
    foreach ($response in $selected) {
        $protected = Protect-GssFeedbackText -Text $response.Text -KnownNames $knownNames
        $piiRedactions += $protected.RedactionCount
        if (-not $protected.PiiScanPassed) { $piiFailures += $protected.RemainingPiiTypes }
        $themes = @(Get-GssFeedbackThemes -Response $response -SanitizedText $protected.Text)
        $unknownThemes = @($themes | Where-Object { $_ -notin $script:GssThemeNames })
        if ($unknownThemes.Count -gt 0) { throw "Unknown guest-feedback theme: $($unknownThemes[0])" }
        $internalCards += [pscustomobject]@{
            EvidenceId = 'feedback-' + $response.ResponseHash.Substring(0, 16)
            ResponseHash = $response.ResponseHash
            RestaurantId = $response.RestaurantId
            Restaurant = $response.Restaurant
            VisitDate = $response.VisitDate.ToString('yyyy-MM-dd')
            ThemeIds = $themes
            Sentiment = Get-GssFeedbackSentiment -Response $response -SanitizedText $protected.Text
            DoNotContact = $response.DoNotContact
            SanitizedText = $protected.Text
        }
    }

    $themes = @()
    foreach ($restaurantGroup in @($internalCards | Group-Object RestaurantId)) {
        foreach ($themeName in $script:GssThemeNames) {
            $matching = @($restaurantGroup.Group | Where-Object { $_.ThemeIds -contains $themeName } | Sort-Object ResponseHash -Unique)
            if ($matching.Count -lt 2) { continue }
            $slug = $themeName.Replace('/', '-').Replace(' ', '-')
            $themes += [pscustomobject][ordered]@{
                theme_id = "theme-$($restaurantGroup.Name)-$slug"
                restaurant_id = $restaurantGroup.Name
                category = $themeName
                unique_response_count = $matching.Count
                concern_count = @($matching | Where-Object { $_.Sentiment -in @('concern', 'mixed') }).Count
                positive_count = @($matching | Where-Object { $_.Sentiment -in @('positive', 'mixed') }).Count
                do_not_contact_count = @($matching | Where-Object { $_.DoNotContact }).Count
                quotable_evidence_ids = @($matching | Where-Object { -not $_.DoNotContact -and $_.SanitizedText } | Select-Object -ExpandProperty EvidenceId)
            }
        }
    }

    $publicCards = @()
    foreach ($card in @($internalCards | Where-Object { -not $_.DoNotContact })) {
        $restaurantDisplayName = Get-GssRestaurantDisplayName $card.Restaurant
        $publicCards += [pscustomobject][ordered]@{
            evidence_id = $card.EvidenceId
            restaurant_id = $card.RestaurantId
            restaurant = $restaurantDisplayName
            restaurant_name = $restaurantDisplayName
            source_entity = $card.Restaurant
            kind = 'guest_feedback_response'
            finding_type = 'response'
            visit_date = $card.VisitDate
            theme_ids = @($themes | Where-Object { $_.restaurant_id -eq $card.RestaurantId -and $card.ThemeIds -contains $_.category } | Select-Object -ExpandProperty theme_id)
            sentiment = $card.Sentiment
            quote_allowed = $true
            outreach_allowed = $true
            sanitized_text = $card.SanitizedText
            display_text = $card.SanitizedText
        }
    }

    $newEntries = @($Ledger.entries)
    foreach ($response in $Inventory.UniqueResponses) {
        if ($response.VisitDate.Date -gt $ReportingDate.Date) { continue }
        if ($ledgerByHash.ContainsKey($response.ResponseHash)) { continue }
        $inCurrentAttachment = @($response.SourcePaths | Where-Object { $_ -eq $currentPath }).Count -gt 0
        $firstPackage = if ($isInitialLedger -and -not $inCurrentAttachment) { 'baseline' } else { $PackageId }
        $firstDate = if ($isInitialLedger -and -not $inCurrentAttachment) { $response.VisitDate.ToString('yyyy-MM-dd') } else { $reportingDateText }
        $newEntries += [pscustomobject][ordered]@{
            response_hash = $response.ResponseHash
            first_seen_reporting_date = $firstDate
            first_seen_package_id = $firstPackage
        }
    }
    $nextLedger = [pscustomobject][ordered]@{
        schema_version = $script:GssFeedbackLedgerVersion
        entries = @($newEntries | Sort-Object response_hash)
    }

    $dateStart = if ($selected.Count -gt 0) { ($selected.VisitDate | Measure-Object -Minimum).Minimum.ToString('yyyy-MM-dd') } else { $null }
    $dateEnd = if ($selected.Count -gt 0) { ($selected.VisitDate | Measure-Object -Maximum).Maximum.ToString('yyyy-MM-dd') } else { $null }
    return [pscustomobject]@{
        ResponseCount = $selected.Count
        VisitDateStart = $dateStart
        VisitDateEnd = $dateEnd
        Cards = $publicCards
        InternalCards = $internalCards
        Themes = $themes
        SelectionFingerprint = [string]$selection.Fingerprint
        NextLedger = $nextLedger
        Privacy = [pscustomobject][ordered]@{
            pii_scan_passed = ($piiFailures.Count -eq 0)
            pii_redaction_count = $piiRedactions
            do_not_contact_response_count = @($internalCards | Where-Object { $_.DoNotContact }).Count
            do_not_contact_text_excluded = $true
            guest_name_fields_excluded = $true
            raw_detail_contains_personal_data = $true
            remaining_pii_types = @($piiFailures | Select-Object -Unique)
        }
    }
}

function Get-GssDeterministicPackageId {
    param(
        [string]$ReportingDate,
        [object[]]$SourceDescriptors,
        [Parameter(Mandatory)][string]$FeedbackSelectionFingerprint,
        [string]$SchemaVersion = $script:GssEmailPackageSchemaVersion,
        [string]$PolicyVersion = $script:GssAnalysisPolicyVersion
    )

    $normalizedSources = @($SourceDescriptors | ForEach-Object {
        $portablePath = if (-not [string]::IsNullOrWhiteSpace([string]$_.source_path)) { [string]$_.source_path } else { [string]$_.path }
        if ([string]::IsNullOrWhiteSpace($portablePath)) { throw "Package source descriptor is missing its portable path: $($_.role)" }
        if ($null -eq $_.byte_size) { throw "Package source descriptor is missing its byte size: $($_.role)" }
        [pscustomobject]@{
            role = [string]$_.role
            path = $portablePath.Replace('\', '/')
            byte_size = [long]$_.byte_size
            sha256 = ([string]$_.sha256).ToLowerInvariant()
        }
    })
    $sourceText = @($normalizedSources |
        Sort-Object @{ Expression = { $_.role.ToLowerInvariant() } }, @{ Expression = { $_.path.ToLowerInvariant() } } |
        ForEach-Object { "$($_.role):$($_.path):$($_.byte_size):$($_.sha256)" }) -join "`n"
    if ($FeedbackSelectionFingerprint -notmatch '^[a-f0-9]{64}$') { throw 'Feedback selection fingerprint must be a lowercase SHA-256 hash.' }
    $material = "$SchemaVersion`n$PolicyVersion`n$ReportingDate`nfeedback_selection:$FeedbackSelectionFingerprint`n$sourceText"
    return 'gss-' + $ReportingDate.Replace('-', '') + '-' + (Get-GssStringSha256 $material).Substring(0, 20)
}

function Get-GssSourceDescriptor {
    param([string]$Role, [string]$Path, [string]$FolderPath)

    $resolved = Resolve-GssDropboxPath -Path $Path -FolderPath $FolderPath -RequireFile
    $item = Get-Item -LiteralPath $resolved
    if ($item.Length -le 0) { throw "Attachment is empty or not fully synced: $Role" }
    return [pscustomobject][ordered]@{
        role = $Role
        source_path = ConvertTo-GssDropboxRelativePath -Path $resolved -FolderPath $FolderPath
        source_full_path = $resolved
        byte_size = [long]$item.Length
        sha256 = Get-GssSha256 $resolved
    }
}

function Assert-GssLoggedFileEvidence {
    param([object]$RunLog, [object[]]$Descriptors)

    $logged = @($RunLog.FileEvidence)
    if ($logged.Count -eq 0) {
        throw 'Run log does not contain the required file size and SHA-256 evidence.'
    }
    $loggedRoles = @('comparison_pdf', 'rolling_workbook', 'prior_year_rolling_workbook', 'live_workbook')
    foreach ($descriptor in @($Descriptors | Where-Object { $_.role -in $loggedRoles })) {
        $entry = $logged | Where-Object { ([string]$_.Role) -eq $descriptor.role } | Select-Object -First 1
        if (-not $entry) { throw "Run log is missing file evidence for $($descriptor.role)." }
        if ([long]$entry.ByteSize -ne [long]$descriptor.byte_size -or ([string]$entry.Sha256).ToLowerInvariant() -ne $descriptor.sha256) {
            throw "File changed after the successful workbook run: $($descriptor.role)"
        }
        $loggedRelative = ([string]$entry.RelativePath).Replace('\', '/')
        if ($loggedRelative -ne $descriptor.source_path) {
            throw "Run log path does not match the current file for $($descriptor.role)."
        }
    }
}

function Format-GssEvidenceNumber {
    param([object]$Value, [switch]$Signed)
    if ($null -eq $Value) { return 'n/a' }
    $number = [math]::Round([double]$Value, 1, [System.MidpointRounding]::AwayFromZero)
    $format = if ($Signed) { '+0.0;-0.0;0.0' } else { '0.0' }
    return [string]::Format([System.Globalization.CultureInfo]::InvariantCulture, "{0:$format}", $number)
}

function ConvertTo-GssMetricEvidenceCard {
    param([object]$Item, [string]$Kind)

    $restaurantDisplayName = Get-GssRestaurantDisplayName $Item.Entity
    $displayText = '{0}: 13-week rolling result {1}; change versus previous rolling window {2} points; change versus prior-year rolling window {3} points; versus all franchisees {4} points.' -f `
        $Item.Metric, (Format-GssEvidenceNumber $Item.Current), (Format-GssEvidenceNumber $Item.ChangeVsPreviousRollingWindow -Signed), (Format-GssEvidenceNumber $Item.YoYImprovement -Signed), (Format-GssEvidenceNumber $Item.VsAllFranchisees -Signed)
    return [pscustomobject][ordered]@{
        evidence_id = $Item.EvidenceId
        restaurant_id = $Item.RestaurantId
        restaurant = $restaurantDisplayName
        restaurant_name = $restaurantDisplayName
        source_entity = $Item.Entity
        kind = $Kind
        finding_type = $Kind
        metric_key = $Item.RawMetric
        metric = $Item.Metric
        direction = if ($Item.LowerIsBetter) { 'lower_is_better' } else { 'higher_is_better' }
        finding_direction = $Item.CandidateDirection
        lower_is_better = [bool]$Item.LowerIsBetter
        rolling_value = $Item.Current
        change_vs_previous_window = $Item.ChangeVsPreviousRollingWindow
        change_vs_prior_year = $Item.YoYImprovement
        vs_franchise = $Item.VsAllFranchisees
        persistent = [bool]$Item.PersistentMovement
        corroboration = @($Item.Corroboration)
        display_text = $displayText
    }
}

function Test-GssNativeFiniteNumber {
    param([object]$Value)

    $isNativeNumber =
        $Value -is [byte] -or
        $Value -is [sbyte] -or
        $Value -is [int16] -or
        $Value -is [uint16] -or
        $Value -is [int32] -or
        $Value -is [uint32] -or
        $Value -is [int64] -or
        $Value -is [uint64] -or
        $Value -is [single] -or
        $Value -is [double] -or
        $Value -is [decimal]
    if (-not $isNativeNumber) { return $false }

    $number = [double]$Value
    return -not [double]::IsNaN($number) -and -not [double]::IsInfinity($number)
}

function Get-GssRequiredEvidenceProperty {
    param(
        [Parameter(Mandatory)][object]$Card,
        [Parameter(Mandatory)][string]$PropertyName,
        [Parameter(Mandatory)][string]$CardLabel
    )

    $property = $Card.PSObject.Properties[$PropertyName]
    if ($null -eq $property) {
        throw "GSS analysis evidence contract violation: $CardLabel is missing required property '$PropertyName'."
    }
    return $property.Value
}

function Assert-GssMetricEvidenceCardContract {
    param(
        [Parameter(Mandatory)][object]$Card,
        [Parameter(Mandatory)][string]$CardLabel
    )

    foreach ($propertyName in @('evidence_id', 'restaurant_id', 'metric_key', 'metric')) {
        $value = Get-GssRequiredEvidenceProperty -Card $Card -PropertyName $propertyName -CardLabel $CardLabel
        if ([string]::IsNullOrWhiteSpace([string]$value)) {
            throw "GSS analysis evidence contract violation: $CardLabel has an empty '$propertyName'."
        }
    }

    $direction = Get-GssRequiredEvidenceProperty -Card $Card -PropertyName 'direction' -CardLabel $CardLabel
    if ($direction -notin @('higher_is_better', 'lower_is_better')) {
        throw "GSS analysis evidence contract violation: $CardLabel has invalid direction '$direction'."
    }
    $lowerIsBetter = Get-GssRequiredEvidenceProperty -Card $Card -PropertyName 'lower_is_better' -CardLabel $CardLabel
    if ($lowerIsBetter -isnot [bool]) {
        throw "GSS analysis evidence contract violation: $CardLabel property 'lower_is_better' must be Boolean."
    }
    if ([bool]$lowerIsBetter -ne ($direction -eq 'lower_is_better')) {
        throw "GSS analysis evidence contract violation: $CardLabel direction conflicts with 'lower_is_better'."
    }

    $rollingValue = Get-GssRequiredEvidenceProperty -Card $Card -PropertyName 'rolling_value' -CardLabel $CardLabel
    if (-not (Test-GssNativeFiniteNumber $rollingValue)) {
        throw "GSS analysis evidence contract violation: $CardLabel property 'rolling_value' must be a native finite number."
    }
    foreach ($propertyName in @('change_vs_previous_window', 'change_vs_prior_year', 'vs_franchise')) {
        $value = Get-GssRequiredEvidenceProperty -Card $Card -PropertyName $propertyName -CardLabel $CardLabel
        if ($null -ne $value -and -not (Test-GssNativeFiniteNumber $value)) {
            throw "GSS analysis evidence contract violation: $CardLabel property '$propertyName' must be null or a native finite number."
        }
    }
}

function Assert-GssThemeEvidenceCardContract {
    param(
        [Parameter(Mandatory)][object]$Card,
        [Parameter(Mandatory)][string]$CardLabel
    )

    foreach ($propertyName in @('theme_id', 'restaurant_id', 'category')) {
        $value = Get-GssRequiredEvidenceProperty -Card $Card -PropertyName $propertyName -CardLabel $CardLabel
        if ([string]::IsNullOrWhiteSpace([string]$value)) {
            throw "GSS analysis evidence contract violation: $CardLabel has an empty '$propertyName'."
        }
    }

    foreach ($propertyName in @('unique_response_count', 'concern_count', 'positive_count', 'do_not_contact_count')) {
        $value = Get-GssRequiredEvidenceProperty -Card $Card -PropertyName $propertyName -CardLabel $CardLabel
        if (-not (Test-GssNativeFiniteNumber $value) -or [double]$value -lt 0 -or [double]$value -ne [math]::Truncate([double]$value)) {
            throw "GSS analysis evidence contract violation: $CardLabel property '$propertyName' must be a nonnegative native integer."
        }
    }
}

function Test-GssAnalysisEvidenceContract {
    param([Parameter(Mandatory)][object]$Analysis)

    foreach ($propertyName in @('metric_evidence', 'theme_evidence', 'evidence_cards')) {
        if ($null -eq $Analysis.PSObject.Properties[$propertyName]) {
            throw "GSS analysis evidence contract violation: analysis is missing required property '$propertyName'."
        }
    }

    $unifiedCards = @($Analysis.evidence_cards)
    foreach ($card in @($Analysis.metric_evidence)) {
        $evidenceId = [string](Get-GssRequiredEvidenceProperty -Card $card -PropertyName 'evidence_id' -CardLabel 'metric evidence card')
        Assert-GssMetricEvidenceCardContract -Card $card -CardLabel "metric evidence '$evidenceId'"
        $matches = @($unifiedCards | Where-Object { [string]$_.evidence_id -eq $evidenceId })
        if ($matches.Count -ne 1) {
            throw "GSS analysis evidence contract violation: metric evidence '$evidenceId' must appear exactly once in evidence_cards."
        }
        Assert-GssMetricEvidenceCardContract -Card $matches[0] -CardLabel "evidence_cards metric '$evidenceId'"
    }

    foreach ($card in @($Analysis.theme_evidence)) {
        $themeId = [string](Get-GssRequiredEvidenceProperty -Card $card -PropertyName 'theme_id' -CardLabel 'theme evidence card')
        Assert-GssThemeEvidenceCardContract -Card $card -CardLabel "theme evidence '$themeId'"
        $matches = @($unifiedCards | Where-Object { [string]$_.theme_id -eq $themeId })
        if ($matches.Count -ne 1) {
            throw "GSS analysis evidence contract violation: theme evidence '$themeId' must appear exactly once in evidence_cards."
        }
        Assert-GssThemeEvidenceCardContract -Card $matches[0] -CardLabel "evidence_cards theme '$themeId'"
    }

    return $true
}

function New-GssEvidencePreviewText {
    param([object]$Analysis, [object]$Feedback, [datetime]$ReportingDate)

    $lines = @()
    $lines += "Subject: GSS Survey Report - Week Ending $($ReportingDate.ToString('MMddyy'))"
    $lines += ''
    $lines += 'All,'
    $lines += ''
    $reportingDateDisplay = $ReportingDate.ToString('MMMM d, yyyy', [System.Globalization.CultureInfo]::InvariantCulture)
    $lines += "The attached report covers 13-week rolling results through $reportingDateDisplay. Movements below compare that result with the previous rolling window and the matching prior-year rolling window."
    foreach ($restaurant in @($Analysis.RestaurantFindings)) {
        $lines += ''
        $lines += "$($restaurant.Restaurant):"
        if (@($restaurant.Strengths).Count -eq 0 -and @($restaurant.Opportunities).Count -eq 0) {
            $lines += 'No movement met the reporting thresholds.'
        }
        foreach ($item in @($restaurant.Strengths)) {
            $lines += "Strength: $($item.Metric) was $(Format-GssEvidenceNumber $item.Current), changing $(Format-GssEvidenceNumber $item.ChangeVsPreviousRollingWindow -Signed) points versus the previous rolling window and $(Format-GssEvidenceNumber $item.YoYImprovement -Signed) points versus prior year."
        }
        foreach ($item in @($restaurant.Opportunities)) {
            $lines += "Opportunity: $($item.Metric) was $(Format-GssEvidenceNumber $item.Current), changing $(Format-GssEvidenceNumber $item.ChangeVsPreviousRollingWindow -Signed) points versus the previous rolling window and $(Format-GssEvidenceNumber $item.YoYImprovement -Signed) points versus prior year."
        }
    }
    $lines += ''
    if ($Feedback.ResponseCount -gt 0) {
        $lines += "New guest feedback: $($Feedback.ResponseCount) unique responses with visit dates from $($Feedback.VisitDateStart) through $($Feedback.VisitDateEnd). Themes are reported only when supported by at least two unique responses."
        foreach ($theme in @($Feedback.Themes)) {
            $lines += "- $($theme.restaurant_id) $($theme.category): $($theme.unique_response_count) unique responses."
        }
    }
    else {
        $lines += 'New guest feedback: no previously unseen responses were found in the validated detail exports.'
    }
    $lines += ''
    $lines += 'Recommended follow-up: review the evidence-backed opportunities with the appropriate restaurant leaders and confirm local context before choosing an action.'
    $lines += ''
    $lines += 'Bottom line: these directional comparisons identify items for review; they do not establish statistical significance or causation.'
    $lines += ''
    $lines += 'Methodology: score results are 13-week rolling aggregates. Guest feedback is deduplicated across current and archived exports, anonymized in this package, and described using its actual visit-date range rather than as an exact seven-day sample.'
    return ($lines -join [Environment]::NewLine)
}

function ConvertTo-GssSimpleHtml {
    param([string]$Text)

    $encodedLines = @($Text -split '\r?\n' | ForEach-Object { [System.Net.WebUtility]::HtmlEncode($_) })
    $body = @()
    foreach ($line in $encodedLines) {
        if ([string]::IsNullOrWhiteSpace($line)) { $body += '<br>' }
        elseif ($line.StartsWith('Subject:')) { $body += "<p><strong>$line</strong></p>" }
        elseif ($line.EndsWith(':')) { $body += "<h2>$line</h2>" }
        elseif ($line.StartsWith('- ')) { $body += "<p>&bull; $($line.Substring(2))</p>" }
        else { $body += "<p>$line</p>" }
    }
    return '<!doctype html><html><head><meta charset="utf-8"></head><body style="font-family:Arial,sans-serif;color:#222;line-height:1.4">' + ($body -join '') + '</body></html>'
}

function Test-GssExistingEmailPackage {
    param(
        [string]$PackagePath,
        [string]$PackageId,
        [object[]]$ExpectedSourceDescriptors,
        [Parameter(Mandatory)][string]$ExpectedFeedbackSelectionFingerprint
    )

    $readyPath = Join-Path $PackagePath 'READY'
    $manifestPath = Join-Path $PackagePath 'email_manifest.json'
    if (-not (Test-Path -LiteralPath $readyPath -PathType Leaf) -or -not (Test-Path -LiteralPath $manifestPath -PathType Leaf)) {
        throw "Existing deterministic package is incomplete and was not overwritten: $PackageId"
    }
    $readyLines = @(Get-Content -LiteralPath $readyPath)
    if ($readyLines.Count -ne 2 -or $readyLines[0] -ne $script:GssEmailPackageSchemaVersion -or $readyLines[1] -ne $PackageId) {
        throw "Existing deterministic package READY marker does not match package ID: $PackageId"
    }
    $manifest = Get-Content -Raw -LiteralPath $manifestPath | ConvertFrom-Json
    if ($manifest.schema_version -ne $script:GssEmailPackageSchemaVersion -or $manifest.package_id -ne $PackageId) {
        throw "Existing deterministic package manifest does not match package ID: $PackageId"
    }
    if ([string]$manifest.feedback_selection_sha256 -notmatch '^[a-f0-9]{64}$' -or [string]$manifest.feedback_selection_sha256 -ne $ExpectedFeedbackSelectionFingerprint) {
        throw "Existing deterministic package feedback selection does not match current package inputs: $PackageId"
    }
    $recomputedPackageId = Get-GssDeterministicPackageId `
        -ReportingDate ([string]$manifest.reporting.reporting_date) `
        -SourceDescriptors @($manifest.sources) `
        -FeedbackSelectionFingerprint ([string]$manifest.feedback_selection_sha256) `
        -SchemaVersion ([string]$manifest.schema_version) `
        -PolicyVersion ([string]$manifest.policy_version)
    if ($recomputedPackageId -ne $PackageId) {
        throw "Existing deterministic package ID does not match its attested source inventory: $PackageId"
    }
    foreach ($requiredName in @('analysis.json', 'email_preview.txt', 'email_preview.html', 'READY')) {
        $requiredPath = Join-Path $PackagePath $requiredName
        if (-not (Test-Path -LiteralPath $requiredPath -PathType Leaf) -or (Get-Item -LiteralPath $requiredPath).Length -le 0) {
            throw "Existing deterministic package is missing required output: $requiredName"
        }
    }
    $analysis = Get-Content -Raw -LiteralPath (Join-Path $PackagePath 'analysis.json') | ConvertFrom-Json
    if ($analysis.schema_version -ne $script:GssEmailPackageSchemaVersion -or $analysis.package_id -ne $PackageId) {
        throw "Existing deterministic package analysis does not match package ID: $PackageId"
    }
    if ([string]$analysis.feedback_selection_sha256 -ne [string]$manifest.feedback_selection_sha256) {
        throw "Existing deterministic package analysis feedback selection does not match its manifest: $PackageId"
    }
    $null = Test-GssAnalysisEvidenceContract -Analysis $analysis
    $availableEvidenceIds = @(
        @($analysis.portfolio_evidence) + @($analysis.metric_evidence) + @($analysis.sanitized_feedback) |
            ForEach-Object { [string]$_.evidence_id } |
            Where-Object { $_ } |
            Sort-Object -Unique
    )
    foreach ($evidenceId in @($manifest.evidence_ids)) {
        if ($evidenceId -notin $availableEvidenceIds) { throw "Existing package references unknown evidence ID: $evidenceId" }
    }
    $availableThemeIds = @($analysis.feedback_themes.theme_id | Where-Object { $_ } | Sort-Object -Unique)
    foreach ($themeId in @($manifest.theme_ids)) {
        if ($themeId -notin $availableThemeIds) { throw "Existing package references unknown theme ID: $themeId" }
    }
    $expectedSources = @($ExpectedSourceDescriptors | ForEach-Object {
        [pscustomobject]@{
            role = [string]$_.role
            path = ([string]$_.source_path).Replace('\', '/')
            byte_size = [long]$_.byte_size
            sha256 = ([string]$_.sha256).ToLowerInvariant()
        }
    })
    $actualSources = @($manifest.sources)
    if ($actualSources.Count -ne $expectedSources.Count) {
        throw "Existing package source inventory does not match current package inputs: $PackageId"
    }
    foreach ($expectedSource in $expectedSources) {
        $matches = @($actualSources | Where-Object {
            ([string]$_.role) -eq $expectedSource.role -and ([string]$_.path).Replace('\', '/') -eq $expectedSource.path
        })
        if ($matches.Count -ne 1) { throw "Existing package source inventory is missing or duplicated: $($expectedSource.role) $($expectedSource.path)" }
        $actualSource = $matches[0]
        if ([long]$actualSource.byte_size -ne $expectedSource.byte_size -or ([string]$actualSource.sha256).ToLowerInvariant() -ne $expectedSource.sha256) {
            throw "Existing package source evidence does not match current input: $($expectedSource.role) $($expectedSource.path)"
        }
    }
    $runLogSource = @($expectedSources | Where-Object role -eq 'run_log')
    if ($runLogSource.Count -ne 1 -or ([string]$manifest.source_log_path).Replace('\', '/') -ne $runLogSource[0].path) {
        throw "Existing package source_log_path does not match its run_log source: $PackageId"
    }
    foreach ($attachment in @($manifest.attachments)) {
        $path = Join-Path $PackagePath ([string]$attachment.path).Replace('/', '\')
        if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { throw "Existing package attachment is missing: $($attachment.role)" }
        $item = Get-Item -LiteralPath $path
        if ([long]$item.Length -ne [long]$attachment.byte_size -or (Get-GssSha256 $path) -ne [string]$attachment.sha256) {
            throw "Existing package attachment does not match its manifest: $($attachment.role)"
        }
    }
    return $manifest
}

function New-GssEmailPackage {
    param(
        [Parameter(Mandatory)][string]$FolderPath,
        [Parameter(Mandatory)][object]$RunLog,
        [Parameter(Mandatory)][object]$AnalysisResult,
        [string]$LedgerPath
    )

    if ($RunLog.Mode -ne 'ApplyToMainWorkbook') {
        throw 'Email packages may be published only from a successful live apply log.'
    }
    if ($AnalysisResult.WorkbookStatus -eq 'Blocked') {
        throw 'Email package publication is blocked because workbook validation failed.'
    }
    if ($AnalysisResult.AnalysisStatus -eq 'Blocked') {
        throw 'Email package publication is blocked because analysis validation failed.'
    }

    $reportingDate = [datetime]::Parse([string]$RunLog.CurrentWeekEnding).Date
    $inventory = Get-GssDetailInventory -FolderPath $FolderPath -ReportingDate $reportingDate
    $sourceDescriptors = @(
        (Get-GssSourceDescriptor -Role 'comparison_pdf' -Path ([string]$RunLog.EmailComparisonPdf) -FolderPath $FolderPath),
        (Get-GssSourceDescriptor -Role 'rolling_workbook' -Path ([string]$RunLog.CurrentSourceWorkbook) -FolderPath $FolderPath),
        (Get-GssSourceDescriptor -Role 'prior_year_rolling_workbook' -Path ([string]$RunLog.PriorYearSourceWorkbook) -FolderPath $FolderPath),
        (Get-GssSourceDescriptor -Role 'live_workbook' -Path ([string]$RunLog.TargetWorkbook) -FolderPath $FolderPath),
        (Get-GssSourceDescriptor -Role 'detail_workbook' -Path $inventory.CurrentWorkbook.Path -FolderPath $FolderPath),
        (Get-GssSourceDescriptor -Role 'run_log' -Path ([string]$AnalysisResult.LogPath) -FolderPath $FolderPath)
    )
    $sourceDescriptors += @($inventory.Workbooks |
        Where-Object { $_.PortablePath -ne $inventory.CurrentWorkbook.PortablePath } |
        Sort-Object PortablePath |
        ForEach-Object { Get-GssSourceDescriptor -Role 'detail_archive_workbook' -Path $_.Path -FolderPath $FolderPath })
    Assert-GssLoggedFileEvidence -RunLog $RunLog -Descriptors $sourceDescriptors
    if ([string]::IsNullOrWhiteSpace($LedgerPath)) {
        $LedgerPath = Join-Path (Join-Path (Join-Path $FolderPath '_automation_runs') 'state') 'gss_feedback_first_seen.json'
    }
    $ledger = Read-GssFeedbackLedger $LedgerPath
    $feedbackSelection = Get-GssFeedbackSelection -Inventory $inventory -Ledger $ledger -ReportingDate $reportingDate
    $packageId = Get-GssDeterministicPackageId `
        -ReportingDate $reportingDate.ToString('yyyy-MM-dd') `
        -SourceDescriptors $sourceDescriptors `
        -FeedbackSelectionFingerprint $feedbackSelection.Fingerprint
    $feedback = New-GssSanitizedFeedback `
        -Inventory $inventory `
        -Ledger $ledger `
        -PackageId $packageId `
        -ReportingDate $reportingDate `
        -FeedbackSelection $feedbackSelection
    if (-not $feedback.Privacy.pii_scan_passed) {
        throw "PII remained after guest-feedback sanitization: $($feedback.Privacy.remaining_pii_types -join ', ')"
    }

    if (Get-Command Select-GssRestaurantFindings -ErrorAction SilentlyContinue) {
        $AnalysisResult.RestaurantFindings = @(Select-GssRestaurantFindings -MetricDetail $AnalysisResult.MetricDetail -GuestThemes $feedback.Themes)
    }

    $rollingStart = $reportingDate.AddDays(-90)
    $previousWindowEnd = $reportingDate.AddDays(-7)
    $previousWindowStart = $rollingStart.AddDays(-7)
    $priorYearEnd = [datetime]::Parse([string]$RunLog.PriorYearWeekEnding).Date
    $reporting = [pscustomobject][ordered]@{
        reporting_date = $reportingDate.ToString('yyyy-MM-dd')
        rolling_start_date = $rollingStart.ToString('yyyy-MM-dd')
        rolling_end_date = $reportingDate.ToString('yyyy-MM-dd')
        previous_window_start_date = $previousWindowStart.ToString('yyyy-MM-dd')
        previous_window_end_date = $previousWindowEnd.ToString('yyyy-MM-dd')
        prior_year_end_date = $priorYearEnd.ToString('yyyy-MM-dd')
        result_label = "13-week rolling through $($reportingDate.ToString('yyyy-MM-dd'))"
        comparison_label = 'change versus previous rolling window'
    }
    $statuses = [pscustomobject][ordered]@{
        workbook_status = [string]$AnalysisResult.WorkbookStatus
        analysis_status = [string]$AnalysisResult.AnalysisStatus
        email_readiness = 'Ready'
    }

    $metricEvidence = @()
    foreach ($restaurant in @($AnalysisResult.RestaurantFindings)) {
        foreach ($item in @($restaurant.Strengths)) { $metricEvidence += ConvertTo-GssMetricEvidenceCard $item 'strength' }
        foreach ($item in @($restaurant.Opportunities)) { $metricEvidence += ConvertTo-GssMetricEvidenceCard $item 'opportunity' }
    }
    $themeEvidence = @()
    foreach ($theme in @($feedback.Themes)) {
        $restaurantFinding = $AnalysisResult.RestaurantFindings | Where-Object { $_.RestaurantId -eq $theme.restaurant_id } | Select-Object -First 1
        $sourceEntity = if ($restaurantFinding) { $restaurantFinding.Restaurant } else { [string]$theme.restaurant_id }
        $restaurantName = if ($restaurantFinding -and $restaurantFinding.Name) { $restaurantFinding.Name } else { Get-GssRestaurantDisplayName $sourceEntity }
        $findingType = if ([int]$theme.concern_count -gt [int]$theme.positive_count) { 'opportunity' } elseif ([int]$theme.positive_count -gt [int]$theme.concern_count) { 'strength' } else { 'mixed' }
        $dncNote = if ([int]$theme.do_not_contact_count -gt 0) { "; $($theme.do_not_contact_count) anonymous do-not-contact response(s) included in counts" } else { '' }
        $displayText = "${restaurantName}: new guest feedback $($theme.category) theme in $($theme.unique_response_count) unique responses; $($theme.concern_count) concern and $($theme.positive_count) positive$dncNote."
        $themeEvidence += [pscustomobject][ordered]@{
            theme_id = $theme.theme_id
            restaurant_id = $theme.restaurant_id
            restaurant = $restaurantName
            restaurant_name = $restaurantName
            source_entity = $sourceEntity
            kind = 'guest_feedback_theme'
            finding_type = $findingType
            category = $theme.category
            unique_response_count = [int]$theme.unique_response_count
            concern_count = [int]$theme.concern_count
            positive_count = [int]$theme.positive_count
            do_not_contact_count = [int]$theme.do_not_contact_count
            quote_allowed = $false
            outreach_allowed = $false
            display_text = $displayText
        }
    }
    $feedbackBasis = if ($feedback.ResponseCount -gt 0) {
        "$($feedback.ResponseCount) unique new guest-feedback responses with visit dates $($feedback.VisitDateStart) through $($feedback.VisitDateEnd)"
    }
    else {
        'no previously unseen guest-feedback responses in the validated exports'
    }
    $portfolioDisplayText = "Portfolio reporting basis: 13-week rolling results through $($reporting.reporting_date); changes compare with the previous rolling window and matching prior-year rolling window; $feedbackBasis. These comparisons are directional and do not establish statistical significance or causation."
    $portfolioEvidence = [pscustomobject][ordered]@{
        evidence_id = 'portfolio-' + (Get-GssStringSha256 ("$packageId|reporting-basis")).Substring(0, 16)
        kind = 'reporting_basis'
        finding_type = 'context'
        restaurant_id = 'portfolio'
        restaurant = 'Portfolio'
        restaurant_name = 'Portfolio'
        display_text = $portfolioDisplayText
    }
    $restaurantSections = @()
    foreach ($restaurant in @($AnalysisResult.RestaurantFindings)) {
        $restaurantThemes = @($feedback.Themes | Where-Object { $_.restaurant_id -eq $restaurant.RestaurantId })
        $restaurantFeedbackCards = @($feedback.InternalCards | Where-Object { $_.RestaurantId -eq $restaurant.RestaurantId })
        $restaurantVisitStart = if ($restaurantFeedbackCards.Count -gt 0) { ($restaurantFeedbackCards.VisitDate | Sort-Object | Select-Object -First 1) } else { $null }
        $restaurantVisitEnd = if ($restaurantFeedbackCards.Count -gt 0) { ($restaurantFeedbackCards.VisitDate | Sort-Object -Descending | Select-Object -First 1) } else { $null }
        $restaurantSections += [pscustomobject][ordered]@{
            restaurant_id = $restaurant.RestaurantId
            restaurant = $restaurant.Name
            name = $restaurant.Name
            source_entity = $restaurant.Restaurant
            strengths = @($metricEvidence | Where-Object { $_.restaurant_id -eq $restaurant.RestaurantId -and $_.kind -eq 'strength' })
            opportunities = @($metricEvidence | Where-Object { $_.restaurant_id -eq $restaurant.RestaurantId -and $_.kind -eq 'opportunity' })
            guest_feedback = [pscustomobject][ordered]@{
                response_count = $restaurantFeedbackCards.Count
                visit_date_start = $restaurantVisitStart
                visit_date_end = $restaurantVisitEnd
                themes = $restaurantThemes
                evidence = @($themeEvidence | Where-Object { $_.restaurant_id -eq $restaurant.RestaurantId })
            }
        }
    }

    $analysisDocument = [pscustomobject][ordered]@{
        schema_version = $script:GssEmailPackageSchemaVersion
        policy_version = $script:GssAnalysisPolicyVersion
        package_id = $packageId
        feedback_selection_sha256 = $feedback.SelectionFingerprint
        reporting = $reporting
        statuses = $statuses
        methodology = [pscustomobject][ordered]@{
            score_basis = 'Each score is a 13-week rolling result; rolling rows are not averaged together.'
            comparison_basis = 'Direction-adjusted change versus the previous rolling window and matching prior-year rolling window.'
            significance_claimed = $false
            causation_claimed = $false
            feedback_label = 'new guest feedback'
            feedback_response_count = $feedback.ResponseCount
            feedback_visit_date_start = $feedback.VisitDateStart
            feedback_visit_date_end = $feedback.VisitDateEnd
        }
        portfolio = [pscustomobject][ordered]@{
            evidence = @($portfolioEvidence)
        }
        restaurants = $restaurantSections
        portfolio_evidence = @($portfolioEvidence)
        metric_evidence = $metricEvidence
        theme_evidence = $themeEvidence
        feedback_themes = $feedback.Themes
        sanitized_feedback = $feedback.Cards
        evidence_cards = @($portfolioEvidence) + @($metricEvidence) + @($themeEvidence) + @($feedback.Cards)
        privacy = $feedback.Privacy
    }
    $null = Test-GssAnalysisEvidenceContract -Analysis $analysisDocument

    $outbox = Join-Path (Join-Path $FolderPath '_automation_runs') 'email_outbox'
    New-Item -ItemType Directory -Path $outbox -Force | Out-Null
    $packagePath = Join-Path $outbox $packageId
    if (Test-Path -LiteralPath $packagePath -PathType Container) {
        $existingManifest = Test-GssExistingEmailPackage `
            -PackagePath $packagePath `
            -PackageId $packageId `
            -ExpectedSourceDescriptors $sourceDescriptors `
            -ExpectedFeedbackSelectionFingerprint $feedback.SelectionFingerprint
        Write-GssFeedbackLedger -Ledger $feedback.NextLedger -Path $LedgerPath
        return [pscustomobject]@{
            EmailReadiness = 'Ready'
            PackageId = $packageId
            PackagePath = $packagePath
            ManifestPath = Join-Path $packagePath 'email_manifest.json'
            ReadyMarkerPath = Join-Path $packagePath 'READY'
            ExistingPackage = $true
        }
    }

    $stagingPath = Join-Path $outbox ('.staging-' + $packageId + '-' + [guid]::NewGuid().ToString('N'))
    $attachmentDirectory = Join-Path $stagingPath 'attachments'
    New-Item -ItemType Directory -Path $attachmentDirectory -Force | Out-Null
    try {
        $dateToken = $reportingDate.ToString('MMddyy')
        $attachmentNames = @{
            comparison_pdf = "GSS Email Comparison $dateToken.pdf"
            rolling_workbook = "GSS 13-Week Rolling $dateToken.xlsx"
            detail_workbook = "GSS Guest Detail $dateToken.xlsx"
        }
        $attachments = @()
        foreach ($source in @($sourceDescriptors | Where-Object { $_.role -in @('comparison_pdf', 'rolling_workbook', 'detail_workbook') })) {
            $fileName = $attachmentNames[$source.role]
            $destination = Join-Path $attachmentDirectory $fileName
            Copy-Item -LiteralPath $source.source_full_path -Destination $destination
            $copiedItem = Get-Item -LiteralPath $destination
            $copiedHash = Get-GssSha256 $destination
            if ([long]$copiedItem.Length -ne [long]$source.byte_size -or $copiedHash -ne $source.sha256) {
                throw "Attachment changed while the package was being created: $($source.role)"
            }
            $attachments += [pscustomobject][ordered]@{
                role = $source.role
                path = 'attachments/' + $fileName
                source_path = $source.source_path
                byte_size = [long]$copiedItem.Length
                sha256 = $copiedHash
            }
        }

        $previewText = New-GssEvidencePreviewText -Analysis $AnalysisResult -Feedback $feedback -ReportingDate $reportingDate
        $previewHtml = ConvertTo-GssSimpleHtml $previewText
        $analysisPath = Join-Path $stagingPath 'analysis.json'
        $textPath = Join-Path $stagingPath 'email_preview.txt'
        $htmlPath = Join-Path $stagingPath 'email_preview.html'
        $manifestPath = Join-Path $stagingPath 'email_manifest.json'
        $analysisDocument | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath $analysisPath -Encoding UTF8
        $previewText | Set-Content -LiteralPath $textPath -Encoding UTF8
        $previewHtml | Set-Content -LiteralPath $htmlPath -Encoding UTF8

        $sourceLogDescriptor = @($sourceDescriptors | Where-Object role -eq 'run_log')
        if ($sourceLogDescriptor.Count -ne 1) { throw 'Package source inventory must contain exactly one run_log descriptor.' }
        $sourceLogPath = [string]$sourceLogDescriptor[0].source_path
        $manifest = [pscustomobject][ordered]@{
            schema_version = $script:GssEmailPackageSchemaVersion
            policy_version = $script:GssAnalysisPolicyVersion
            package_id = $packageId
            feedback_selection_sha256 = $feedback.SelectionFingerprint
            reporting = $reporting
            statuses = $statuses
            source_log_path = $sourceLogPath
            analysis_path = 'analysis.json'
            text_preview_path = 'email_preview.txt'
            html_preview_path = 'email_preview.html'
            ready_marker_path = 'READY'
            sources = @($sourceDescriptors | ForEach-Object {
                [pscustomobject][ordered]@{
                    role = $_.role
                    path = $_.source_path
                    byte_size = $_.byte_size
                    sha256 = $_.sha256
                }
            })
            attachments = $attachments
            evidence_ids = @(
                @($portfolioEvidence) + @($metricEvidence) + @($feedback.Cards) |
                    ForEach-Object { [string]$_.evidence_id } |
                    Where-Object { $_ } |
                    Sort-Object -Unique
            )
            theme_ids = @($feedback.Themes.theme_id | Sort-Object -Unique)
            privacy = $feedback.Privacy
        }
        $manifest | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath $manifestPath -Encoding UTF8

        foreach ($requiredName in @('email_manifest.json', 'analysis.json', 'email_preview.txt', 'email_preview.html')) {
            $requiredPath = Join-Path $stagingPath $requiredName
            if (-not (Test-Path -LiteralPath $requiredPath -PathType Leaf) -or (Get-Item -LiteralPath $requiredPath).Length -le 0) {
                throw "Package output is empty or missing: $requiredName"
            }
        }
        foreach ($attachment in $attachments) {
            $attachmentPath = Join-Path $stagingPath $attachment.path.Replace('/', '\')
            if ((Get-GssSha256 $attachmentPath) -ne $attachment.sha256) { throw "Final attachment validation failed: $($attachment.role)" }
        }
        foreach ($source in $sourceDescriptors) {
            $sourceItem = Get-Item -LiteralPath $source.source_full_path
            if ([long]$sourceItem.Length -ne [long]$source.byte_size -or (Get-GssSha256 $source.source_full_path) -ne $source.sha256) {
                throw "Source changed while the package was being created: $($source.role) $($source.source_path)"
            }
        }
        $analysisText = Get-Content -Raw -LiteralPath $analysisPath
        $manifestText = Get-Content -Raw -LiteralPath $manifestPath
        $aiFacingText = $analysisText + $previewText + $previewHtml
        $nonRawText = $aiFacingText + $manifestText
        if ($nonRawText -match '(?i)(?:[A-Z]:[\\/]|\\\\[^\\])') { throw 'A machine-specific path leaked into a portable package file.' }
        # Deterministic methodology and metadata can legitimately contain an
        # ordinary word that is also present in a guest-name field (for example,
        # "sample"). Scan all portable text for generic PII, but scope the
        # known-name check to guest-derived text that was actually published.
        $sanitizedFeedbackText = @($feedback.Cards | ForEach-Object {
            [string]$_.sanitized_text
            [string]$_.display_text
        }) -join "`n"
        $remainingPii = @(
            @(Get-GssRemainingPiiTypes -Text $nonRawText -KnownNames @()) +
            @(Get-GssRemainingPiiTypes -Text $sanitizedFeedbackText -KnownNames (Get-GssKnownGuestNames $inventory.AllResponseInstances)) |
                Sort-Object -Unique
        )
        if ($remainingPii.Count -gt 0) { throw "PII remained in portable package output: $($remainingPii -join ', ')" }

        # Persist first-seen state before making the ready package visible. If the
        # final promotion fails, the next identical run reselects this package ID.
        Write-GssFeedbackLedger -Ledger $feedback.NextLedger -Path $LedgerPath
        # This marker is the final package-file write; the directory is then promoted atomically.
        "$($script:GssEmailPackageSchemaVersion)`n$packageId" | Set-Content -LiteralPath (Join-Path $stagingPath 'READY') -Encoding ASCII
        $null = Publish-GssStagedEmailPackage `
            -StagingPath $stagingPath `
            -PackagePath $packagePath `
            -ValidationOperation {
                Test-GssExistingEmailPackage `
                    -PackagePath $packagePath `
                    -PackageId $packageId `
                    -ExpectedSourceDescriptors $sourceDescriptors `
                    -ExpectedFeedbackSelectionFingerprint $feedback.SelectionFingerprint
            }
        return [pscustomobject]@{
            EmailReadiness = 'Ready'
            PackageId = $packageId
            PackagePath = $packagePath
            ManifestPath = Join-Path $packagePath 'email_manifest.json'
            ReadyMarkerPath = Join-Path $packagePath 'READY'
            ExistingPackage = $false
        }
    }
    catch {
        $packageError = $_
        if (Test-Path -LiteralPath $stagingPath -PathType Container) {
            $resolvedOutbox = [System.IO.Path]::GetFullPath($outbox).TrimEnd('\') + '\'
            $resolvedStaging = [System.IO.Path]::GetFullPath($stagingPath)
            if ($resolvedStaging.StartsWith($resolvedOutbox, [System.StringComparison]::OrdinalIgnoreCase)) {
                try {
                    Invoke-GssFileSystemRetry -Operation {
                        Remove-Item -LiteralPath $resolvedStaging -Recurse -Force
                    }
                }
                catch {
                    $packageError.Exception.Data['GssStagingCleanupError'] = $_.Exception.Message
                }
            }
        }
        throw $packageError
    }
}
