BeforeAll {
    . (Join-Path (Split-Path -Parent $PSScriptRoot) 'scripts\Update-GSS-MainWorkbook.ps1')

    function New-TestRollingSource {
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Returns only an in-memory test fixture.')]
        [CmdletBinding()]
        param([string]$Name = 'current.xlsx', [datetime]$Week = [datetime]'2026-10-04', [double]$Overall = 70, [double]$Dissat = 5)
        $rows = @(
            [pscustomobject]@{ SourceFile = $Name; Week = $Week; Ownership = 'All Franchisees'; Restaurant = $null; Metrics = @(1000, 68, 80, 85, 76, 78, 92, 66, 7, 6, 7, 8, 9) },
            [pscustomobject]@{ SourceFile = $Name; Week = $Week; Ownership = 'Sorensen'; Restaurant = $null; Metrics = @(300, 73, 80, 85, 76, 78, 92, 66, 5, 6, 7, 8, 9) },
            [pscustomobject]@{ SourceFile = $Name; Week = $Week; Ownership = 'Sorensen'; Restaurant = '9354 Richmond'; Metrics = @(150, $Overall, 80, 85, 76, 78, 92, 66, $Dissat, 6, 7, 8, 9) },
            [pscustomobject]@{ SourceFile = $Name; Week = $Week; Ownership = 'Sorensen'; Restaurant = '9355 Virginia Beach'; Metrics = @(150, 82, 80, 85, 76, 78, 92, 66, 4, 6, 7, 8, 9) }
        )
        [pscustomobject]@{
            File = [pscustomobject]@{ FullName = "C:\GSS Test\$Name"; Name = $Name; LastWriteTime = [datetime]'2026-10-05' }
            WeekEnding = $Week
            PeriodStart = $Week.AddDays(-90)
            Rows = $rows
            SemanticSha256 = Get-GssSourceSemanticHash $rows $Week.AddDays(-90) $Week
        }
    }
}

Describe 'Reporting period and number validation' {
    It 'requires the actual inclusive 13-week Monday-Sunday period' {
        (Get-GssReportingPeriod ([datetime]'2026-07-06') ([datetime]'2026-10-04')).InclusiveDays | Should -Be 91
        { Get-GssReportingPeriod ([datetime]'2026-01-01') ([datetime]'2026-10-04') } | Should -Throw '*13 complete weeks*'
        { Get-GssReportingPeriod ([datetime]'2026-07-07') ([datetime]'2026-10-05') } | Should -Throw '*Monday through Sunday*'
    }

    It 'rejects a future report even if its dates are internally consistent' {
        { Get-GssReportingPeriod ([datetime]'2026-07-13') ([datetime]'2026-10-11') -AsOfDate ([datetime]'2026-10-04') } | Should -Throw '*future*'
    }

    It 'rejects negative and fractional response counts' {
        { ConvertTo-GssMetricValue (@(-1) + @(75) * 12) } | Should -Throw '*whole number*'
        { ConvertTo-GssMetricValue (@(1.5) + @(75) * 12) } | Should -Throw '*whole number*'
    }

    It 'rejects missing, out-of-range, nonnumeric and nonfinite scores' {
        foreach ($invalid in @($null, '', -1, 101, 'incorrect', [double]::NaN, [double]::PositiveInfinity)) {
            { ConvertTo-GssMetricValue (@(150, $invalid) + @(75) * 11) } | Should -Throw
        }
    }

    It 'keeps zero-response scores as unscored instead of manufacturing zero percentages' {
        $values = ConvertTo-GssMetricValue (@(0) + @($null) * 12)
        $values.Count | Should -Be 13
        $values[0] | Should -Be 0
        $values[1] | Should -BeNullOrEmpty
    }

    It 'stops publication when the legacy workbook would turn an unscored result into false zeros' {
        $source = New-TestRollingSource
        $source.Rows[2].Metrics = ConvertTo-GssMetricValue (@(0) + @($null) * 12)
        { Assert-GssWorkbookScoreCompatibility $source.Rows } | Should -Throw '*cannot display blank scores safely*last good report is preserved*'
        $source.Rows[2].Metrics = ConvertTo-GssMetricValue (@(150, 0) + @(75) * 11)
        # A measured zero percent with nonzero responses remains legitimate.
        { Assert-GssWorkbookScoreCompatibility $source.Rows } | Should -Not -Throw
    }

    It 'accepts invariant numeric data on a computer using a different locale' {
        $previousCulture = [Threading.Thread]::CurrentThread.CurrentCulture
        try {
            [Threading.Thread]::CurrentThread.CurrentCulture = [Globalization.CultureInfo]'fr-FR'
            $values = ConvertTo-GssMetricValue (@(150, 75.5) + @(75) * 11)
            $values[1] | Should -Be 75.5
        }
        finally { [Threading.Thread]::CurrentThread.CurrentCulture = $previousCulture }
    }
}

Describe 'Content identity and corrections' {
    It 'ignores filename, row order and harmless numerical serialization noise' {
        $original = New-TestRollingSource -Name 'original.xlsx'
        $renamed = New-TestRollingSource -Name 'Weekly survey results.xlsx' -Overall 70.000000000001
        [array]::Reverse($renamed.Rows)
        (Get-GssSourceSemanticHash $renamed.Rows $renamed.PeriodStart $renamed.WeekEnding) | Should -Be $original.SemanticSha256
        (New-TestRollingSource -Overall 70.0001).SemanticSha256 | Should -Not -Be $original.SemanticSha256
    }

    It 'chooses equal-content aliases without using file timestamps' {
        $first = New-TestRollingSource -Name 'a.xlsx'
        $second = New-TestRollingSource -Name 'z.xlsx'
        $second.File.LastWriteTime = [datetime]'2099-01-01'
        (Select-GssSourceContent @($second, $first)).File.FullName | Should -Be $first.File.FullName
    }

    It 'reports a structured choice for different versions of the same week' {
        $first = New-TestRollingSource -Name 'original.xlsx'
        $second = New-TestRollingSource -Name 'corrected.xlsx' -Overall 72
        $caught = $null
        try { Select-GssSourceContent @($first, $second) }
        catch { $caught = $_.Exception }
        $caught | Should -Not -BeNullOrEmpty
        $caught.Data['GssChoice'] | Should -Be 'CorrectionSource'
        @($caught.Data['Candidates']).Count | Should -Be 2
        (Select-GssSourceContent @($first, $second) @($second.File.FullName)).SemanticSha256 | Should -Be $second.SemanticSha256
        { Select-GssSourceContent @($first, $second) @($first.File.FullName, $second.File.FullName) } | Should -Throw '*only one*'
    }

    It 'does not append an existing report after a harmless rename' {
        $old = New-TestRollingSource -Name 'original.xlsx'
        $new = New-TestRollingSource -Name 'renamed.xlsx'
        $plan = @(New-GssWeekUpdatePlan @([pscustomobject]@{ Role = 'CurrentWeek'; Source = $new }) $old.Rows)
        $plan[0].Action | Should -Be 'Unchanged'
    }

    It 'requires an explicit correction choice even when only one updated file is present' {
        $old = New-TestRollingSource -Overall 70
        $new = New-TestRollingSource -Overall 72
        $selected = @([pscustomobject]@{ Role = 'CurrentWeek'; Source = $new })
        { New-GssWeekUpdatePlan $selected $old.Rows } | Should -Throw '*changes existing values*'
        $plan = @(New-GssWeekUpdatePlan $selected $old.Rows @($new.File.FullName))
        $plan[0].Action | Should -Be 'Correct'
        $plan[0].Rows[2].Metrics[1] | Should -Be 72
        $plan[0].PreviousSemanticSha256 | Should -Be $old.SemanticSha256
    }

    It 'repairs partial or duplicate existing periods only through an explicit correction' {
        $source = New-TestRollingSource
        $selected = @([pscustomobject]@{ Role = 'CurrentWeek'; Source = $source })
        foreach ($existing in @(@($source.Rows[0..2]), @($source.Rows + $source.Rows[2]))) {
            { New-GssWeekUpdatePlan $selected $existing } | Should -Throw '*changes existing values*'
            $plan = @(New-GssWeekUpdatePlan $selected $existing @($source.File.FullName))
            $plan[0].Action | Should -Be 'Correct'
            $plan[0].Rows.Count | Should -Be 4
        }
    }

    It 'appends all four business entities exactly once for a new period' {
        $source = New-TestRollingSource
        $plan = @(New-GssWeekUpdatePlan @([pscustomobject]@{ Role = 'CurrentWeek'; Source = $source }) @())
        $plan[0].Action | Should -Be 'Append'
        $plan[0].Rows.Count | Should -Be 4
    }
}

Describe 'Chronological catch-up' {
    It 'requests a missed week and its missing prior-year comparison before the current week' {
        $plan = @(Get-GssRequiredSourceWeek ([datetime]'2026-10-04') @([datetime]'2026-09-20', [datetime]'2025-09-28'))
        @($plan | ForEach-Object { $_.WeekEnding.ToString('yyyy-MM-dd') }) -join ',' | Should -Be '2025-10-05,2026-09-27,2026-10-04'
    }

    It 'keeps the declared 2025 history gap as a compatibility exception' {
        $plan = @(Get-GssRequiredSourceWeek ([datetime]'2025-10-05') @([datetime]'2025-07-06'))
        @($plan | Where-Object { Test-GssIntentionalHistoryGap $_.WeekEnding }).Count | Should -Be 0
        @($plan | Where-Object Role -eq 'CurrentWeek').Count | Should -Be 1
        $catchup = @(Get-GssRequiredSourceWeek ([datetime]'2026-10-04') @([datetime]'2026-09-20'))
        @($catchup | Where-Object { $_.WeekEnding -eq [datetime]'2025-09-28' }).Count | Should -Be 0
    }

    It 'includes an explicitly selected historical correction and does not regress a newer report' {
        $plan = @(Get-GssRequiredSourceWeek ([datetime]'2026-10-04') @([datetime]'2026-10-04') @([datetime]'2026-09-06'))
        @($plan | Where-Object Role -eq 'HistoricalCorrection').Count | Should -Be 1
        { Get-GssRequiredSourceWeek ([datetime]'2026-09-27') @([datetime]'2026-10-04') } | Should -Throw '*later reporting week*'
    }
}

Describe 'Business results in the existing renderer' {
    BeforeEach {
        $script:renderedCell = @{}
        $script:businessValue = @{}
        $metricNames = @(Get-GssSourceHeader)[2..14]
        foreach ($source in @((New-TestRollingSource), (New-TestRollingSource -Week ([datetime]'2026-09-27') -Overall 75 -Dissat 8), (New-TestRollingSource -Week ([datetime]'2025-10-05') -Overall 80 -Dissat 9))) {
            foreach ($row in $source.Rows) {
                for ($index = 0; $index -lt $metricNames.Count; $index++) {
                    $key = '{0}|{1}|{2}' -f (ConvertTo-GssSourceEntityKey $row.Ownership $row.Restaurant), $row.Week.ToString('yyyyMMdd'), $metricNames[$index]
                    $script:businessValue[$key] = $row.Metrics[$index]
                }
            }
        }
        Mock Set-CellValue { param($Worksheet, $Row, $Column, $Value) $null = $Worksheet; $script:renderedCell["$Row,$Column"] = $Value }
        Mock Get-RawMetricValue {
            param($RawDataWorksheet, $RawIndex, $EntityKey, $WeekEnding, $MetricKey)
            $null = $RawDataWorksheet; $null = $RawIndex
            $script:businessValue[('{0}|{1}|{2}' -f $EntityKey, $WeekEnding.ToString('yyyyMMdd'), $MetricKey)]
        }
    }

    It 'renders the independently specified score movements and units' {
        $cells = [pscustomobject]@{}
        $cells | Add-Member ScriptMethod Item { [pscustomobject]@{ NumberFormat = '' } }
        $sheet = [pscustomobject]@{ Cells = $cells }
        $metrics = @(
            [pscustomobject]@{ DisplayName = 'Overall Experience'; RawKey = 'Overall Experience'; LowerIsBetter = $false; IncludeInMovers = $true; Category = 'Overall'; Owner = 'GM' },
            [pscustomobject]@{ DisplayName = 'Overall Dissat'; RawKey = 'Overall Dissat (lower is better)'; LowerIsBetter = $true; IncludeInMovers = $true; Category = 'Overall'; Owner = 'GM' },
            [pscustomobject]@{ DisplayName = 'Count'; RawKey = 'Count'; LowerIsBetter = $false; IncludeInMovers = $false; Category = 'Sampling'; Owner = 'Ops' }
        )
        Update-EmailComparisonSection $sheet $sheet ([pscustomobject]@{}) $metrics ([datetime]'2026-10-04') 'Sorensen|9354 Richmond' 'Richmond' 3 6 'All Franchisees|(TOTAL)' 'All Franchisees'
        $script:renderedCell['6,2'] | Should -Be 70
        $script:renderedCell['6,3'] | Should -Be -5
        $script:renderedCell['6,4'] | Should -Be -10
        $script:renderedCell['7,2'] | Should -Be 5
        $script:renderedCell['7,3'] | Should -Be 3
        $script:renderedCell['7,4'] | Should -Be 4
        $script:renderedCell['8,2'] | Should -Be 150
        $script:renderedCell['8,3'] | Should -BeNullOrEmpty
        $script:renderedCell['6,5'] | Should -BeNullOrEmpty
    }

    It 'retains the vendor company score instead of averaging the two restaurants' {
        $cells = [pscustomobject]@{}
        $cells | Add-Member ScriptMethod Item { [pscustomobject]@{ NumberFormat = '' } }
        $sheet = [pscustomobject]@{ Cells = $cells }
        $metric = [pscustomobject]@{ DisplayName = 'Overall Experience'; RawKey = 'Overall Experience'; LowerIsBetter = $false; IncludeInMovers = $true; Category = 'Overall'; Owner = 'GM' }
        Update-EmailComparisonSection $sheet $sheet ([pscustomobject]@{}) @($metric) ([datetime]'2026-10-04') 'Sorensen|(TOTAL)' 'Company' 3 6 'All Franchisees|(TOTAL)' 'All Franchisees'
        # Two equally sampled restaurant scores are 70 and 82 (mean 76).
        # The vendor's company aggregate is independently supplied as 73.
        $script:renderedCell['6,2'] | Should -Be 73
        $script:renderedCell['6,6'] | Should -Be 300
    }
}

Describe 'Excel error inspection' {
    BeforeEach { Mock Release-ComObject {} }

    It 'independently confirms no errors when SpecialCells reports no matching cells' {
        $range = [pscustomobject]@{}
        $range | Add-Member ScriptMethod SpecialCells { throw 'No matching cells' }
        $range | Add-Member ScriptMethod Address { 'A1:B2' }
        $sheet = [pscustomobject]@{ UsedRange = $range }
        $sheet | Add-Member ScriptMethod Evaluate { 0 }
        Get-GssExcelErrorCellCount $sheet -4123 | Should -Be 0
    }

    It 'does not report zero errors when the independent Excel inspection fails too' {
        $range = [pscustomobject]@{}
        $range | Add-Member ScriptMethod SpecialCells { throw 'COM inspection failed' }
        $range | Add-Member ScriptMethod Address { 'A1:B2' }
        $sheet = [pscustomobject]@{ UsedRange = $range }
        $sheet | Add-Member ScriptMethod Evaluate { throw 'RPC server unavailable' }
        { Get-GssExcelErrorCellCount $sheet -4123 } | Should -Throw '*RPC server unavailable*'
    }
}
