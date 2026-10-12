# Pure reporting rules. These functions do not open Excel or change files.

function Get-GssSourceHeader {
    return @(
        'Restaurant', 'Ownership', 'Count', 'Overall Experience', 'Service', 'Culinary',
        'Value', 'Pace of Meal', 'Steak Cooked Properly (Y)', 'Mgr Asked About Exp (Y)',
        'Overall Dissat (lower is better)', 'Service Dissat (lower is better)',
        'Culinary Dissat (lower is better)', 'Value Dissat (lower is better)',
        'Pace Of Meal Dissat (lower is better)'
    )
}

function ConvertTo-GssSourceEntityKey {
    param([string]$Ownership, [AllowNull()][string]$Restaurant)

    $name = if ([string]::IsNullOrWhiteSpace($Restaurant)) { '(TOTAL)' } else { $Restaurant.Trim() }
    return "$($Ownership.Trim())|$name"
}

function ConvertTo-GssHeaderDate {
    param([Parameter(Mandatory)][string]$Value)

    return [datetime]::ParseExact(
        $Value,
        [string[]]@('M/d/yyyy', 'MM/dd/yyyy', 'M/dd/yyyy', 'MM/d/yyyy'),
        [Globalization.CultureInfo]::InvariantCulture,
        [Globalization.DateTimeStyles]::None
    ).Date
}

function Get-GssReportingPeriod {
    param(
        [Parameter(Mandatory)][datetime]$PeriodStart,
        [Parameter(Mandatory)][datetime]$PeriodEnd,
        [datetime]$AsOfDate = (Get-Date).Date,
        [string]$Context = 'Rolling source'
    )

    $start = $PeriodStart.Date
    $end = $PeriodEnd.Date
    if ($end.DayOfWeek -ne [DayOfWeek]::Sunday -or ($end - $start).Days -ne 90) {
        throw "$Context must cover exactly 13 complete weeks, Monday through Sunday (91 inclusive days). Found $($start.ToString('yyyy-MM-dd')) through $($end.ToString('yyyy-MM-dd'))."
    }
    if ($end -gt $AsOfDate.Date) {
        throw "$Context ends in the future: $($end.ToString('yyyy-MM-dd'))."
    }
    return [pscustomobject]@{ Start = $start; End = $end; InclusiveDays = 91 }
}

function ConvertTo-GssMetricValue {
    param(
        [Parameter(Mandatory)][AllowEmptyCollection()][AllowNull()][object[]]$Values,
        [string]$Context = 'Source row'
    )

    $headers = @(Get-GssSourceHeader)[2..14]
    if ($Values.Count -ne $headers.Count) {
        throw "$Context must contain a response count and all 12 reporting metrics."
    }
    $result = New-Object 'object[]' $headers.Count
    for ($index = 0; $index -lt $headers.Count; $index++) {
        $value = $Values[$index]
        if ($null -eq $value -or ($value -is [string] -and [string]::IsNullOrWhiteSpace($value))) {
            if ($index -gt 0 -and $result[0] -eq 0) {
                $result[$index] = $null
                continue
            }
            throw "$Context is missing '$($headers[$index])'."
        }
        $number = 0.0
        if ($value -is [bool] -or $value -is [datetime] -or
            -not [double]::TryParse([string]$value, [Globalization.NumberStyles]::Float,
                [Globalization.CultureInfo]::InvariantCulture, [ref]$number) -or
            [double]::IsNaN($number) -or [double]::IsInfinity($number)) {
            throw "$Context has a non-numeric or non-finite '$($headers[$index])' value."
        }
        if ($index -eq 0) {
            if ($number -lt 0 -or $number -ne [math]::Truncate($number)) {
                throw "$Context response count must be a nonnegative whole number."
            }
        }
        else {
            # Eight decimal places are substantially finer than the delivered
            # report precision while ignoring harmless Excel serialization noise.
            $number = [math]::Round($number, 8, [MidpointRounding]::AwayFromZero)
            if ($number -lt 0 -or $number -gt 100) {
                throw "$Context '$($headers[$index])' must be between 0 and 100 percentage points."
            }
        }
        $result[$index] = $number
    }
    # Preserve null metrics for a genuine zero-response row; never coerce them to zero.
    return ,$result
}

function Get-GssSourceSemanticHash {
    param(
        [Parameter(Mandatory)][object[]]$Rows,
        [Parameter(Mandatory)][datetime]$PeriodStart,
        [Parameter(Mandatory)][datetime]$PeriodEnd
    )

    $parts = @('gss-rolling-content/v1', $PeriodStart.ToString('yyyy-MM-dd'), $PeriodEnd.ToString('yyyy-MM-dd'))
    foreach ($row in @($Rows | Sort-Object { ConvertTo-GssSourceEntityKey $_.Ownership $_.Restaurant })) {
        if ([datetime]$row.Week -ne $PeriodEnd.Date) { throw 'Semantic source rows do not match their reporting period.' }
        $key = ConvertTo-GssSourceEntityKey $row.Ownership $row.Restaurant
        $values = ConvertTo-GssMetricValue -Values $row.Metrics -Context $key
        $numbers = @($values | ForEach-Object {
            if ($null -eq $_) { 'null' } else { ([double]$_).ToString('R', [Globalization.CultureInfo]::InvariantCulture) }
        })
        $parts += "$key|$($numbers -join '|')"
    }
    $sha = [System.Security.Cryptography.SHA256]::Create()
    try {
        return ([BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes(($parts -join "`n"))))).Replace('-', '').ToLowerInvariant()
    }
    finally { $sha.Dispose() }
}

function Stop-GssSourceChoice {
    [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Creates and throws an in-memory choice error; does not change files or application state.')]
    [CmdletBinding()]
    param([Parameter(Mandatory)][object[]]$Sources, [Parameter(Mandatory)][string]$Reason)

    $choiceError = New-Object System.InvalidOperationException($Reason)
    $choiceError.Data['GssChoice'] = 'CorrectionSource'
    $choiceError.Data['Message'] = $Reason
    $choiceError.Data['Candidates'] = @($Sources | Sort-Object { $_.File.FullName } | ForEach-Object {
        [pscustomobject]@{ Path = [string]$_.File.FullName; WeekEnding = $_.WeekEnding.ToString('yyyy-MM-dd') }
    })
    throw $choiceError
}

function Select-GssSourceContent {
    param(
        [Parameter(Mandatory)][object[]]$Sources,
        [string[]]$CorrectionSourcePaths = @()
    )

    if ($Sources.Count -eq 0) { throw 'No rolling source is available for the requested week.' }
    if (@($Sources | ForEach-Object { $_.WeekEnding.ToString('yyyy-MM-dd') } | Sort-Object -Unique).Count -ne 1) {
        throw 'Content selection requires candidates for one reporting week.'
    }
    $chosen = @($Sources | Where-Object { $_.File.FullName -in $CorrectionSourcePaths })
    if ($chosen.Count -gt 1) { throw 'Choose only one correction source for each reporting week.' }
    if ($chosen.Count -eq 1) { return $chosen[0] }
    if (@($Sources | Select-Object -ExpandProperty SemanticSha256 -Unique).Count -ne 1) {
        Stop-GssSourceChoice -Sources $Sources -Reason "Different rolling reports exist for $($Sources[0].WeekEnding.ToString('yyyy-MM-dd')). Choose the intended revision; file timestamps are not authoritative."
    }
    # Equal content is one logical input. A stable path order avoids filesystem-time dependence.
    return $Sources | Sort-Object { $_.File.FullName.ToLowerInvariant() }, { $_.File.FullName } | Select-Object -First 1
}

function Assert-GssWorkbookScoreCompatibility {
    param([AllowEmptyCollection()][object[]]$Rows = @())
    foreach ($row in $Rows) {
        if ($row.Metrics.Count -lt 13 -or [double]$row.Metrics[0] -ne 0) { continue }
        $entity = ConvertTo-GssSourceEntityKey $row.Ownership $row.Restaurant
        throw "The workbook model cannot report scores safely for zero-response row $entity, week $($row.Week.ToString('yyyy-MM-dd')). Zero responses cannot establish a measured score, including a vendor-supplied numeric zero. The last good report is preserved."
    }
}

function Get-GssRequiredSourceWeek {
    param(
        [Parameter(Mandatory)][datetime]$ReportingWeek,
        [AllowEmptyCollection()][datetime[]]$ExistingWeeks = @(),
        [AllowEmptyCollection()][datetime[]]$CorrectionWeeks = @()
    )

    $week = $ReportingWeek.Date
    $existing = @($ExistingWeeks | ForEach-Object { $_.Date } | Sort-Object -Unique)
    if (@($existing | Where-Object { $_ -gt $week }).Count -gt 0) {
        throw 'The main workbook already contains a later reporting week. Run that latest week and specify the older correction source instead of replacing a newer report with an older one.'
    }
    $required = @{}
    $required[$week.ToString('yyyy-MM-dd')] = 'CurrentWeek'
    $required[$week.AddDays(-364).ToString('yyyy-MM-dd')] = 'PriorYearWeek'
    $previous = $week.AddDays(-7)
    if ($previous -notin $existing) {
        $required[$previous.ToString('yyyy-MM-dd')] = 'CatchUpWeek'
    }
    if ($existing.Count -gt 0) {
        $lastExisting = $existing[-1]
        for ($cursor = $lastExisting.AddDays(7); $cursor -le $week; $cursor = $cursor.AddDays(7)) {
            if (-not $required.ContainsKey($cursor.ToString('yyyy-MM-dd'))) {
                $required[$cursor.ToString('yyyy-MM-dd')] = 'CatchUpWeek'
            }
            $priorYear = $cursor.AddDays(-364)
            if ($priorYear -notin $existing -and -not $required.ContainsKey($priorYear.ToString('yyyy-MM-dd'))) {
                $required[$priorYear.ToString('yyyy-MM-dd')] = 'CatchUpPriorYearWeek'
            }
        }
    }
    foreach ($correction in $CorrectionWeeks) {
        if ($correction.Date -gt $week) { throw 'A correction source cannot be later than the selected reporting week.' }
        if (-not $required.ContainsKey($correction.ToString('yyyy-MM-dd'))) {
            $required[$correction.ToString('yyyy-MM-dd')] = 'HistoricalCorrection'
        }
    }
    foreach ($key in @($required.Keys | Sort-Object)) {
        [pscustomobject]@{ WeekEnding = [datetime]::ParseExact($key, 'yyyy-MM-dd', [Globalization.CultureInfo]::InvariantCulture); Role = $required[$key] }
    }
}

function New-GssWeekUpdatePlan {
    [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Returns an in-memory plan only; does not apply the planned changes.')]
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][object[]]$SelectedSources,
        [AllowEmptyCollection()][object[]]$ExistingRows = @(),
        [string[]]$CorrectionSourcePaths = @()
    )

    foreach ($requested in @($SelectedSources | Sort-Object { $_.Source.WeekEnding })) {
        $source = $requested.Source
        $existing = @($ExistingRows | Where-Object { $_.Week.Date -eq $source.WeekEnding.Date })
        $oldHash = $null
        $existingValid = $true
        if ($existing.Count -gt 0) {
            $keys = @($existing | ForEach-Object { ConvertTo-GssSourceEntityKey $_.Ownership $_.Restaurant })
            $expected = @('All Franchisees|(TOTAL)', 'Sorensen|(TOTAL)', 'Sorensen|9354 Richmond', 'Sorensen|9355 Virginia Beach')
            $existingValid = $keys.Count -eq 4 -and @($keys | Sort-Object -Unique).Count -eq 4 -and @($expected | Where-Object { $_ -notin $keys }).Count -eq 0
            if ($existingValid) {
                try { $oldHash = Get-GssSourceSemanticHash $existing $source.PeriodStart $source.WeekEnding }
                catch { $existingValid = $false }
            }
        }
        $action = if ($existing.Count -eq 0) { 'Append' } elseif ($existingValid -and $oldHash -eq $source.SemanticSha256) { 'Unchanged' } else { 'Correct' }
        if ($action -eq 'Correct' -and $source.File.FullName -notin $CorrectionSourcePaths) {
            Stop-GssSourceChoice -Sources @($source) -Reason "The uploaded report changes existing values for $($source.WeekEnding.ToString('yyyy-MM-dd')). Choose this source to replace that reporting period; the previous good workbook will be backed up."
        }
        [pscustomobject]@{
            Role = $requested.Role
            Source = $source
            Action = $action
            ExistingRows = $existing.Count
            Rows = @($source.Rows)
            PreviousSemanticSha256 = $oldHash
            SemanticSha256 = $source.SemanticSha256
        }
    }
}
