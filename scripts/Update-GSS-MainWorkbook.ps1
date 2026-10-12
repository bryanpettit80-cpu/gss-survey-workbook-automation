[CmdletBinding()]
param(
    [string]$Folder,
    [string]$MainWorkbookName = 'GSS Score Trends - Main.xlsx',
    [switch]$Apply,
    [switch]$OutputObject,
    [string]$RunId,
    [string]$PreparedRunLogPath,
    [string]$ExpectedFingerprint,
    [string]$RollbackRunLogPath,
    [switch]$MutexAlreadyHeld,
    [string]$ProgramRelease = 'unreleased',
    [string]$DrivePreparedManifestPath,
    [string]$ExpectedDrivePreparedManifestSha256,
    [string]$ReportingWeekEnding,
    [string[]]$CorrectionSourcePath = @()
)

$ErrorActionPreference = 'Stop'

$scriptRoot = Split-Path -Parent $PSCommandPath
. (Join-Path $scriptRoot 'Gss-Common.ps1')
. (Join-Path $scriptRoot 'Gss-Reporting.ps1')
. (Join-Path $scriptRoot 'Gss-SourceChoices.ps1')
. (Join-Path $scriptRoot 'Gss-DriveBackup.ps1')
$script:GssProgramRepoRoot = Split-Path -Parent $scriptRoot

$xlUp = -4162
$xlToLeft = -4159
$xlPasteFormats = -4122
$xlCalculationAutomatic = -4105
$xlCalculationManual = -4135
$xlTypePDF = 0
$xlPortrait = 1
$xlLandscape = 2
$xlSheetVisible = -1
$xlValidateList = 3
$xlValidAlertStop = 1
$xlBetween = 1
$xlUnlockedCells = 1
$xlCellTypeFormulas = -4123
$xlCellTypeConstants = 2
$xlErrors = 16
$xlLinkTypeExcelLinks = 1
$gssRawDataLastFormulaRow = 5096
$gssExpectedEntityKeys = @(
    'All Franchisees|(TOTAL)',
    'Sorensen|(TOTAL)',
    'Sorensen|9354 Richmond',
    'Sorensen|9355 Virginia Beach'
)

function Release-ComObject {
    param([object]$ComObject)
    if ($null -ne $ComObject) {
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($ComObject)
    }
}

function Write-GssAtomicText {
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][AllowEmptyString()][string]$Content
    )

    $directory = Split-Path -Parent $Path
    if ([string]::IsNullOrWhiteSpace($directory)) {
        throw "An absolute or directory-qualified output path is required: $Path"
    }
    New-Item -ItemType Directory -Path $directory -Force | Out-Null
    $temporaryPath = Join-Path $directory ('.t-{0}' -f ([guid]::NewGuid().ToString('N').Substring(0, 8)))
    $replacementBackupPath = Join-Path $directory ('.b-{0}' -f ([guid]::NewGuid().ToString('N').Substring(0, 8)))
    try {
        [System.IO.File]::WriteAllText($temporaryPath, $Content, (New-Object System.Text.UTF8Encoding($false)))
        if (Test-Path -LiteralPath $Path -PathType Leaf) {
            for ($attempt = 1; $attempt -le 20; $attempt++) {
                try {
                    [System.IO.File]::Replace($temporaryPath, $Path, $replacementBackupPath, $true)
                    break
                }
                catch [System.IO.IOException] {
                    if ($attempt -eq 20) { throw }
                    if (Test-Path -LiteralPath $replacementBackupPath -PathType Leaf) {
                        Remove-Item -LiteralPath $replacementBackupPath -Force
                    }
                    Start-Sleep -Milliseconds 250
                }
            }
        }
        else {
            [System.IO.File]::Move($temporaryPath, $Path)
        }
    }
    finally {
        if (Test-Path -LiteralPath $temporaryPath -PathType Leaf) {
            Remove-Item -LiteralPath $temporaryPath -Force
        }
        if (Test-Path -LiteralPath $replacementBackupPath -PathType Leaf) {
            Remove-Item -LiteralPath $replacementBackupPath -Force
        }
    }
}

function Write-GssAtomicJson {
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][object]$InputObject,
        [int]$Depth = 8
    )

    Write-GssAtomicText -Path $Path -Content ($InputObject | ConvertTo-Json -Depth $Depth)
}

function Get-GssTextSha256 {
    param([Parameter(Mandatory)][AllowEmptyString()][string]$Text)

    $sha = [System.Security.Cryptography.SHA256]::Create()
    try {
        $bytes = (New-Object System.Text.UTF8Encoding($false)).GetBytes($Text)
        return ([System.BitConverter]::ToString($sha.ComputeHash($bytes))).Replace('-', '').ToLowerInvariant()
    }
    finally {
        $sha.Dispose()
    }
}

function Get-GssPreparedRunFingerprint {
    param(
        [Parameter(Mandatory)][string]$RunId,
        [Parameter(Mandatory)][string]$HostName,
        [Parameter(Mandatory)][string]$CurrentWeekEnding,
        [Parameter(Mandatory)][string]$StartingWorkbookSha256,
        [Parameter(Mandatory)][string]$CurrentSourceSha256,
        [Parameter(Mandatory)][string]$PriorYearSourceSha256,
        [Parameter(Mandatory)][string]$StagedWorkbookSha256,
        [Parameter(Mandatory)][string]$StagedPdfSha256,
        [string]$ProgramRelease = 'unreleased',
        [string]$SourceSetSha256,
        [string]$SourceChoiceSetSha256
    )

    $parts = @(
        'gss-transaction-v1',
        $RunId.Trim().ToLowerInvariant(),
        $HostName.Trim().ToLowerInvariant(),
        $CurrentWeekEnding.Trim(),
        $StartingWorkbookSha256.Trim().ToLowerInvariant(),
        $CurrentSourceSha256.Trim().ToLowerInvariant(),
        $PriorYearSourceSha256.Trim().ToLowerInvariant(),
        $StagedWorkbookSha256.Trim().ToLowerInvariant(),
        $StagedPdfSha256.Trim().ToLowerInvariant(),
        $ProgramRelease.Trim().ToLowerInvariant()
    )
    if (-not [string]::IsNullOrWhiteSpace($SourceSetSha256)) { $parts += $SourceSetSha256.Trim().ToLowerInvariant() }
    if (-not [string]::IsNullOrWhiteSpace($SourceChoiceSetSha256)) { $parts += $SourceChoiceSetSha256.Trim().ToLowerInvariant() }
    return Get-GssTextSha256 ($parts -join "`n")
}

function Get-GssFingerprintFromPreparedRun {
    param([Parameter(Mandatory)][object]$PreparedRun)

    return Get-GssPreparedRunFingerprint `
        -RunId ([string]$PreparedRun.RunId) `
        -HostName ([string]$PreparedRun.HostName) `
        -CurrentWeekEnding ([string]$PreparedRun.CurrentWeekEnding) `
        -StartingWorkbookSha256 ([string]$PreparedRun.StartingWorkbookSha256) `
        -CurrentSourceSha256 ([string]$PreparedRun.CurrentSourceSha256) `
        -PriorYearSourceSha256 ([string]$PreparedRun.PriorYearSourceSha256) `
        -StagedWorkbookSha256 ([string]$PreparedRun.StagedWorkbookSha256) `
        -StagedPdfSha256 ([string]$PreparedRun.StagedPdfSha256) `
        -ProgramRelease $(if ($PreparedRun.ProgramRelease) { [string]$PreparedRun.ProgramRelease } else { 'unreleased' }) `
        -SourceSetSha256 ([string]$PreparedRun.SourceSetSha256) `
        -SourceChoiceSetSha256 ([string]$PreparedRun.SourceChoiceSetSha256)
}

function Get-GssSourceChoiceFingerprint {
    param([AllowEmptyCollection()][object[]]$Choices = @())
    $material = @($Choices | Sort-Object WeekEnding | ForEach-Object {
        @($_.WeekEnding, $_.SelectedSemanticSha256, (@($_.KnownSemanticSha256 | Sort-Object -Unique) -join ',')) -join '|'
    }) -join "`n"
    return Get-GssTextSha256 $material
}

function Get-GssCommittedSourceChoice {
    param([Parameter(Mandatory)][string]$LogDirectory)
    $receipts = @()
    foreach ($file in @(Get-ChildItem -LiteralPath $LogDirectory -Filter '*.json' -File -ErrorAction SilentlyContinue | Where-Object { $_.Name -match '_(apply|source_choices)\.json$' })) {
        try {
            $receipt = Get-Content -LiteralPath $file.FullName -Raw | ConvertFrom-Json
            if ($receipt.TransactionStatus -ne 'Committed' -or -not $receipt.AcceptedSourceChoices) { continue }
            if ($file.Name -like '*_source_choices.json' -and ($receipt.Mode -ne 'SourceChoiceOnly' -or $receipt.Status -ne 'Accepted')) { continue }
            if ((Get-GssSourceChoiceFingerprint @($receipt.AcceptedSourceChoices)) -ne [string]$receipt.SourceChoiceSetSha256) { continue }
            if ((Get-GssFingerprintFromPreparedRun $receipt) -ne [string]$receipt.RunFingerprint) { continue }
            $receipts += $receipt
        }
        catch { Write-Verbose "Ignoring unreadable correction evidence in $($file.Name): $($_.Exception.Message)" }
    }
    foreach ($receipt in @($receipts | Sort-Object { [datetime]$_.TimestampUtc } -Descending)) {
        foreach ($choice in @($receipt.AcceptedSourceChoices)) {
            $matchedEvidence = @($receipt.SourceEvidence | Where-Object { $_.WeekEnding -eq $choice.WeekEnding -and $_.SemanticSha256 -eq $choice.SelectedSemanticSha256 })
            if ($matchedEvidence.Count -eq 1) { $choice }
        }
    }
}

function Write-GssAcceptedSourceChoiceReceipt {
    param(
        [Parameter(Mandatory)][object]$PreparedRun,
        [Parameter(Mandatory)][object]$ExistingRun,
        [Parameter(Mandatory)][string]$FolderPath,
        [Parameter(Mandatory)][string]$LogDirectory
    )
    if (-not $PreparedRun.AcceptedSourceChoices) { return $null }
    if ($PreparedRun.Mode -ne 'CopyTestOnly' -or $PreparedRun.TransactionStatus -ne 'Prepared' -or
        -not [bool]$PreparedRun.NoBusinessChanges -or [int]$PreparedRun.RowsAppended -ne 0 -or [int]$PreparedRun.RowsCorrected -ne 0 -or
        $PreparedRun.PostSaveValidation.Status -ne 'Ready' -or $ExistingRun.TransactionStatus -ne 'Committed') {
        throw 'Source-choice-only evidence requires a validated unchanged workbook and a committed existing report.'
    }
    if ((Get-GssFingerprintFromPreparedRun $PreparedRun) -ne [string]$PreparedRun.RunFingerprint) { throw 'Source-choice preparation fingerprint verification failed.' }
    Assert-GssSourceEvidence -Run $PreparedRun -FolderPath $FolderPath
    $mainPath = Resolve-GssDropboxPath -Path ([string]$ExistingRun.TargetWorkbookRelativePath) -FolderPath $FolderPath -RequireFile
    $mainHash = Get-GssSha256 $mainPath
    if ($mainHash -ne [string]$PreparedRun.StartingWorkbookSha256 -or $mainHash -ne [string]$ExistingRun.PromotedWorkbookSha256) {
        throw 'The existing report workbook changed before recording the source choice.'
    }
    $receipt = $PreparedRun | Select-Object *
    $receiptFields = @{
        Mode = 'SourceChoiceOnly'; Status = 'Accepted'; TransactionStatus = 'Committed'
        TimestampUtc = [datetime]::UtcNow.ToString('o'); ExistingReportRunId = [string]$ExistingRun.RunId
        NoWorkbookOrPdfChange = $true; DecisionBasis = 'ValidatedUnchangedWorkbookAndExistingCommittedReport'
    }
    foreach ($name in $receiptFields.Keys) { $receipt | Add-Member -NotePropertyName $name -NotePropertyValue $receiptFields[$name] -Force }
    $path = Join-Path $LogDirectory ("gss_choice_$($PreparedRun.RunId)_source_choices.json")
    $receipt | Add-Member -NotePropertyName LogPath -NotePropertyValue $path -Force
    Write-GssAtomicJson -Path $path -InputObject $receipt
    return $path
}

function Get-GssSourceEvidenceFingerprint {
    param([Parameter(Mandatory)][object[]]$Evidence)

    $material = @($Evidence | Sort-Object WeekEnding, RelativePath | ForEach-Object {
        @($_.WeekEnding, $_.PeriodStart, $_.RelativePath, $_.Sha256, $_.SemanticSha256) -join '|'
    }) -join "`n"
    return Get-GssTextSha256 $material
}

function Assert-GssSourceEvidence {
    param([Parameter(Mandatory)][object]$Run, [Parameter(Mandatory)][string]$FolderPath)

    if (-not $Run.SourceEvidence) {
        if ([int]$Run.ReceiptSchemaVersion -ge 2) { throw 'Prepared source-set evidence is missing.' }
        return
    }
    if ($Run.SourceChoiceSetSha256 -and (Get-GssSourceChoiceFingerprint @($Run.AcceptedSourceChoices)) -ne [string]$Run.SourceChoiceSetSha256) {
        throw 'Prepared source-choice evidence changed. Prepare a fresh run.'
    }
    if ((Get-GssSourceEvidenceFingerprint @($Run.SourceEvidence)) -ne [string]$Run.SourceSetSha256) {
        throw 'Prepared source-set evidence changed. Prepare a fresh run.'
    }
    foreach ($source in @($Run.SourceEvidence)) {
        $path = Resolve-GssDropboxPath -Path $source.RelativePath -FolderPath $FolderPath -RequireFile
        if ((Get-GssSha256 $path) -ne [string]$source.Sha256) {
            throw "Rolling source changed during this transaction: $($source.RelativePath). Prepare a fresh run."
        }
    }
}

function Get-CellString {
    param([object]$Worksheet, [int]$Row, [int]$Column)
    $value = $Worksheet.Cells.Item($Row, $Column).Value2
    if ($null -eq $value) { return '' }
    return ([string]$value).Trim()
}

function Set-CellValue {
    param([object]$Worksheet, [int]$Row, [int]$Column, [object]$Value)
    $cell = $null
    try {
        $cell = $Worksheet.Cells.Item($Row, $Column)
        if ($null -eq $Value) {
            $cell.Value2 = ''
            return
        }
        if ($Value -is [bool]) {
            $Value = if ($Value) { 'TRUE' } else { 'FALSE' }
        }
        elseif ($Value -is [ValueType] -and $Value -isnot [datetime]) {
            # Excel Value2 accepts numbers as COM doubles. PowerShell can retain
            # boxed integer/decimal wrappers from pipeline fixture or cell data.
            $Value = [double]$Value
        }

        try {
            $cell.Value2 = $Value
        }
        catch [System.InvalidCastException] {
            try { $cell.Value = $Value }
            catch { throw "Excel rejected cell ($Row,$Column) value type $($Value.GetType().FullName): $($_.Exception.Message)" }
        }
    }
    finally {
        Release-ComObject $cell
    }
}

function Set-CellFormula {
    param([object]$Worksheet, [int]$Row, [int]$Column, [string]$Formula)

    $cell = $null
    try {
        $cell = $Worksheet.Cells.Item($Row, $Column)
        try {
            $cell.Formula2 = $Formula
        }
        catch {
            $cell.Formula = $Formula
        }
    }
    finally {
        Release-ComObject $cell
    }
}

function Get-GssExcelColor {
    param([int]$Red, [int]$Green, [int]$Blue)
    return ($Red + (256 * $Green) + (65536 * $Blue))
}

function Get-GssQaCheckPlan {
    param(
        [int]$EntityLastRow = 5,
        [int]$MetricLastRow = 14,
        [int]$RawDataLastRow = $gssRawDataLastFormulaRow
    )

    $activeEntityRange = '''Entities''!$E$2:$E${ENTITY_LAST}'.Replace('${ENTITY_LAST}', [string]$EntityLastRow)
    $activeFlagPredicate = '((UPPER(${ACTIVE_RANGE}&"")="TRUE")+(UPPER(${ACTIVE_RANGE}&"")="YES")+(UPPER(${ACTIVE_RANGE}&"")="1")>0)'.Replace('${ACTIVE_RANGE}', $activeEntityRange)
    $activeCount = 'SUMPRODUCT(--${ACTIVE_FLAG})'.Replace('${ACTIVE_FLAG}', $activeFlagPredicate)
    $entityRows = 'SUMPRODUCT(--${ACTIVE_FLAG},--(COUNTIFS(''Raw_Data''!$R$2:$R${RAW_LAST},''Entities''!$D$2:$D${ENTITY_LAST},''Raw_Data''!$B$2:$B${RAW_LAST},${WEEK})=1))'.Replace('${ACTIVE_FLAG}', $activeFlagPredicate).Replace('${ENTITY_LAST}', [string]$EntityLastRow).Replace('${RAW_LAST}', [string]$RawDataLastRow)
    $blankMetrics = 'SUMPRODUCT(--(''Raw_Data''!$B$2:$B${RAW_LAST}=${WEEK}),--(''Raw_Data''!$E$2:$E${RAW_LAST}=""))+SUMPRODUCT(--(''Raw_Data''!$B$2:$B${RAW_LAST}=${WEEK}),--(''Raw_Data''!$E$2:$E${RAW_LAST}>0),MMULT(--(''Raw_Data''!$F$2:$Q${RAW_LAST}=""),TRANSPOSE(COLUMN(''Raw_Data''!$F$1:$Q$1)^0)))'.Replace('${RAW_LAST}', [string]$RawDataLastRow)

    $checks = @()
    $checks += [pscustomobject]@{
        Row = 6
        Name = 'Duplicate Raw_Data RowKeys'
        Formula = '=IF(SUMPRODUCT(--(''Raw_Data''!$S$2:$S${RAW_LAST}<>""),--(COUNTIF(''Raw_Data''!$S$2:$S${RAW_LAST},''Raw_Data''!$S$2:$S${RAW_LAST})>1))=0,"READY","ATTENTION")'.Replace('${RAW_LAST}', [string]$RawDataLastRow)
        DetailFormula = '=SUMPRODUCT(--(''Raw_Data''!$S$2:$S${RAW_LAST}<>""),--(COUNTIF(''Raw_Data''!$S$2:$S${RAW_LAST},''Raw_Data''!$S$2:$S${RAW_LAST})>1))&" duplicate data row(s)"'.Replace('${RAW_LAST}', [string]$RawDataLastRow)
    }

    $periods = @(
        [pscustomobject]@{ Row = 7; Name = 'Current rolling-window entity and metric completeness'; Week = '$B$3'; AllowIntentionalGap = $false },
        [pscustomobject]@{ Row = 8; Name = 'Previous rolling-window entity and metric completeness'; Week = '$B$3-7'; AllowIntentionalGap = $true },
        [pscustomobject]@{ Row = 9; Name = 'Prior-year rolling-window entity and metric completeness'; Week = '$B$3-364'; AllowIntentionalGap = $false }
    )
    foreach ($period in $periods) {
        $rowCheck = $entityRows.Replace('${WEEK}', $period.Week)
        $blankCheck = $blankMetrics.Replace('${WEEK}', $period.Week)
        $formula = '=IF(AND(${ROW_CHECK}=${ACTIVE_COUNT},${BLANK_CHECK}=0),"READY","ATTENTION")'.Replace('${ROW_CHECK}', $rowCheck).Replace('${ACTIVE_COUNT}', $activeCount).Replace('${BLANK_CHECK}', $blankCheck)
        $detailFormula = '=${ROW_CHECK}&" of "&${ACTIVE_COUNT}&" active entities; "&${BLANK_CHECK}&" blank required metric value(s)"'.Replace('${ROW_CHECK}', $rowCheck).Replace('${ACTIVE_COUNT}', $activeCount).Replace('${BLANK_CHECK}', $blankCheck)
        if ($period.AllowIntentionalGap) {
            $intentionalGap = 'AND(${WEEK}>DATE(2025,7,6),${WEEK}<DATE(2025,10,5))'.Replace('${WEEK}', $period.Week)
            $formula = '=IF(${INTENTIONAL_GAP},"READY",IF(AND(${ROW_CHECK}=${ACTIVE_COUNT},${BLANK_CHECK}=0),"READY","ATTENTION"))'.Replace('${INTENTIONAL_GAP}', $intentionalGap).Replace('${ROW_CHECK}', $rowCheck).Replace('${ACTIVE_COUNT}', $activeCount).Replace('${BLANK_CHECK}', $blankCheck)
            $detailFormula = '=IF(${INTENTIONAL_GAP},"Intentional 2025 history gap",${ROW_CHECK}&" of "&${ACTIVE_COUNT}&" active entities; "&${BLANK_CHECK}&" blank required metric value(s)")'.Replace('${INTENTIONAL_GAP}', $intentionalGap).Replace('${ROW_CHECK}', $rowCheck).Replace('${ACTIVE_COUNT}', $activeCount).Replace('${BLANK_CHECK}', $blankCheck)
        }
        $checks += [pscustomobject]@{
            Row = $period.Row
            Name = $period.Name
            Formula = $formula
            DetailFormula = $detailFormula
        }
    }

    $checks += [pscustomobject]@{
        Row = 10
        Name = 'Required Raw_Data metric headers'
        Formula = '=IF(SUMPRODUCT(--(''Metric_Catalog''!$B$2:$B${METRIC_LAST}<>""),--ISNUMBER(MATCH(''Metric_Catalog''!$B$2:$B${METRIC_LAST},''Raw_Data''!$A$1:$S$1,0)))=COUNTA(''Metric_Catalog''!$B$2:$B${METRIC_LAST}),"READY","ATTENTION")'.Replace('${METRIC_LAST}', [string]$MetricLastRow)
        DetailFormula = '=(COUNTA(''Metric_Catalog''!$B$2:$B${METRIC_LAST})-SUMPRODUCT(--(''Metric_Catalog''!$B$2:$B${METRIC_LAST}<>""),--ISNUMBER(MATCH(''Metric_Catalog''!$B$2:$B${METRIC_LAST},''Raw_Data''!$A$1:$S$1,0))))&" missing required header(s)"'.Replace('${METRIC_LAST}', [string]$MetricLastRow)
    }
    $checks += [pscustomobject]@{
        Row = 11
        Name = 'Dashboard selector validity'
        Formula = '=IF(AND(SUMPRODUCT(--(''Entities''!$A$2:$A${ENTITY_LAST}=''Exec_Dashboard''!$C$3),--${ACTIVE_FLAG})=1,COUNTIF(''Metric_Catalog''!$A$2:$A${METRIC_LAST},''Exec_Dashboard''!$F$3)=1),"READY","ATTENTION")'.Replace('${ENTITY_LAST}', [string]$EntityLastRow).Replace('${METRIC_LAST}', [string]$MetricLastRow).Replace('${ACTIVE_FLAG}', $activeFlagPredicate)
        DetailFormula = '="Entity: "&''Exec_Dashboard''!$C$3&"; Metric: "&''Exec_Dashboard''!$F$3'
    }
    $checks += [pscustomobject]@{
        Row = 12
        Name = 'Remaining Raw_Data capacity'
        Formula = '=IF(ROWS(''Raw_Data''!$B$2:$B${RAW_LAST})-COUNT(''Raw_Data''!$B$2:$B${RAW_LAST})>=8,"READY","ATTENTION")'.Replace('${RAW_LAST}', [string]$RawDataLastRow)
        DetailFormula = '=(ROWS(''Raw_Data''!$B$2:$B${RAW_LAST})-COUNT(''Raw_Data''!$B$2:$B${RAW_LAST}))&" rows remain before the formula bound"'.Replace('${RAW_LAST}', [string]$RawDataLastRow)
    }

    return $checks
}

function Get-GssWorkbookLayoutPlan {
    param([int]$ActiveEntityCount, [int]$MetricCount)

    return @(
        [pscustomobject]@{ Sheet = 'README'; PrintArea = '$A$1:$A$34'; Orientation = $xlPortrait; FitWide = 1; FitTall = 2 },
        [pscustomobject]@{ Sheet = 'QA Checks'; PrintArea = '$A$1:$C$12'; Orientation = $xlPortrait; FitWide = 1; FitTall = 1 },
        [pscustomobject]@{ Sheet = 'Quick_Read_WoW'; PrintArea = ('$A$1:$Q${0}' -f (5 + $ActiveEntityCount)); Orientation = $xlLandscape; FitWide = 1; FitTall = 1 },
        [pscustomobject]@{ Sheet = 'Quick_Read_YoY'; PrintArea = ('$A$1:$Q${0}' -f (5 + $ActiveEntityCount)); Orientation = $xlLandscape; FitWide = 1; FitTall = 1 },
        [pscustomobject]@{ Sheet = 'Exec_Dashboard'; PrintArea = '$A$1:$AD$57'; Orientation = $xlLandscape; FitWide = 1; FitTall = 1 },
        [pscustomobject]@{ Sheet = 'Drivers_Detail'; PrintArea = ('$A$1:$AB${0}' -f (5 + $MetricCount)); Orientation = $xlLandscape; FitWide = 1; FitTall = 1 },
        [pscustomobject]@{ Sheet = 'Email Comparison'; PrintArea = '$A$1:$AB$38'; Orientation = $xlLandscape; FitWide = 1; FitTall = 1 }
    )
}

function Get-GssEmailComparisonLayoutPlan {
    return [pscustomobject]@{
        WrappedTableHeaderRanges = @('C5:D5', 'C25:D25')
        VisibleColumnWidths = @(
            [pscustomobject]@{ Column = 2; Width = 22.0 }
            [pscustomobject]@{ Column = 3; Width = 18.5 }
            [pscustomobject]@{ Column = 4; Width = 22.5 }
        )
    }
}

function Get-GssGuardrailPlan {
    return [pscustomobject]@{
        ProtectWorkbookStructure = $true
        ProtectAllWorksheets = $true
        LockScope = 'AllCells'
        UnlockedCells = @('Exec_Dashboard!C3', 'Exec_Dashboard!F3')
        ValidationShowError = $true
        IntentionalGapIsBlocker = $false
    }
}

function Get-HeaderMap {
    param([object]$Worksheet, [int]$HeaderRow, [switch]$AllowDuplicateHeaders)

    $lastCol = $Worksheet.Cells.Item($HeaderRow, $Worksheet.Columns.Count).End($xlToLeft).Column
    $map = @{}
    for ($col = 1; $col -le $lastCol; $col++) {
        $key = Normalize-Header (Get-CellString $Worksheet $HeaderRow $col)
        if ($key) {
            if ($map.ContainsKey($key) -and -not $AllowDuplicateHeaders) { throw "Duplicate header '$key' at row $HeaderRow, columns $($map[$key]) and $col." }
            $map[$key] = $col
        }
    }
    return $map
}

function Get-NormalizedSourceName {
    param([System.IO.FileInfo]$File)
    # Filename is provenance only. Reporting identity is the parsed period and content.
    return $File.Name
}

function Get-RollingSources {
    param([object]$Excel, [string]$FolderPath)

    $sourceFolder = Resolve-GssChildFolder $FolderPath '02 Weekly Rolling Source Workbooks'
    $searchRoot = if ($sourceFolder) { $sourceFolder } else { $FolderPath }
    $candidates = @(Get-GssFiles $searchRoot '*.xlsx' |
        Where-Object { $_.Name -notlike '~$*' } |
        Sort-Object FullName)

    if ($candidates.Count -eq 0 -and $sourceFolder) {
        Write-Warning "No rolling Sorensen source workbooks were found in '02 Weekly Rolling Source Workbooks'. Falling back to a recursive compatibility search under $FolderPath."
        $candidates = @(Get-GssFiles $FolderPath '*.xlsx' |
            Where-Object { $_.Name -notlike '~$*' } |
            Sort-Object FullName)
    }

    $sources = @()
    foreach ($candidate in $candidates) {
        $sourceWb = $null
        $sourceWs = $null
        try {
            $sourceWb = $Excel.Workbooks.Open($candidate.FullName, 0, $true)
            $sourceWs = $sourceWb.Worksheets.Item(1)
            $a1 = Get-CellString $sourceWs 1 1

            $headerRow = $null
            for ($row = 1; $row -le 10; $row++) {
                $rowHeaders = Get-HeaderMap $sourceWs $row -AllowDuplicateHeaders
                if ($rowHeaders.ContainsKey('restaurant') -and $rowHeaders.ContainsKey('ownership')) {
                    $headerRow = $row
                    break
                }
            }

            if ($null -eq $headerRow) {
                continue
            }

            if ($a1 -match 'Date Range:\s*(?<start>\d{1,2}/\d{1,2}/\d{4})\s+to\s+(?<end>\d{1,2}/\d{1,2}/\d{4})') {
                $startDate = ConvertTo-GssHeaderDate $Matches.start
                $endDate = ConvertTo-GssHeaderDate $Matches.end

                $sources += [pscustomobject]@{
                    File = $candidate
                    WeekEnding = $endDate
                    PeriodStart = $startDate
                    SourceFileName = Get-NormalizedSourceName $candidate
                    HeaderRow = $headerRow
                }
            }
        }
        finally {
            if ($sourceWb) { $sourceWb.Close($false) }
            Release-ComObject $sourceWs
            Release-ComObject $sourceWb
        }
    }

    if ($sources.Count -eq 0) {
        throw "No rolling Sorensen source workbook with a Date Range header was found in $FolderPath."
    }

    return $sources
}

function Select-RollingSourceByWeek {
    param(
        [object[]]$Sources,
        [datetime]$WeekEnding
    )

    $weekCandidates = @($Sources | Where-Object { $_.WeekEnding.Date -eq $WeekEnding.Date })
    if ($weekCandidates.Count -eq 0) { return $null }
    return Select-GssSourceContent -Sources $weekCandidates
}

function Select-LatestRollingSource {
    param([object[]]$Sources)

    if ($Sources.Count -eq 0) { return $null }
    $latestWeek = ($Sources | Sort-Object WeekEnding -Descending | Select-Object -First 1).WeekEnding
    return Select-RollingSourceByWeek -Sources $Sources -WeekEnding $latestWeek
}

function Get-GssEntitySetValidation {
    param(
        [Parameter(Mandatory)][AllowEmptyCollection()][string[]]$ActualKeys,
        [string[]]$ExpectedKeys = $gssExpectedEntityKeys
    )

    $actual = @($ActualKeys | ForEach-Object { ([string]$_).Trim() } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    $expected = @($ExpectedKeys | ForEach-Object { ([string]$_).Trim() })
    $missing = @($expected | Where-Object { $_ -notin $actual })
    $unexpected = @($actual | Where-Object { $_ -notin $expected } | Sort-Object -Unique)
    $duplicates = @(
        $actual |
            Group-Object |
            Where-Object Count -gt 1 |
            ForEach-Object { '{0} ({1} rows)' -f $_.Name, $_.Count }
    )

    return [pscustomobject]@{
        IsValid = ($missing.Count -eq 0 -and $unexpected.Count -eq 0 -and $duplicates.Count -eq 0 -and $actual.Count -eq $expected.Count)
        ExpectedCount = $expected.Count
        ActualCount = $actual.Count
        Missing = $missing
        Unexpected = $unexpected
        Duplicates = $duplicates
    }
}

function Assert-GssExactEntitySet {
    param(
        [Parameter(Mandatory)][AllowEmptyCollection()][string[]]$ActualKeys,
        [Parameter(Mandatory)][string]$Context,
        [string[]]$ExpectedKeys = $gssExpectedEntityKeys
    )

    $validation = Get-GssEntitySetValidation -ActualKeys $ActualKeys -ExpectedKeys $ExpectedKeys
    if (-not $validation.IsValid) {
        $details = @()
        if ($validation.Missing.Count -gt 0) { $details += "missing: $($validation.Missing -join ', ')" }
        if ($validation.Unexpected.Count -gt 0) { $details += "unexpected: $($validation.Unexpected -join ', ')" }
        if ($validation.Duplicates.Count -gt 0) { $details += "duplicates: $($validation.Duplicates -join ', ')" }
        if ($details.Count -eq 0) { $details += "expected $($validation.ExpectedCount) rows but found $($validation.ActualCount)" }
        throw "$Context does not contain the exact required entity set ($($details -join '; ')). No workbook mutation was performed."
    }
    return $validation
}

function Get-SourceRows {
    param(
        [object]$SourceWorksheet,
        [hashtable]$HeaderMap,
        [int]$HeaderRow,
        [datetime]$WeekEnding,
        [string]$SourceFileName
    )

    $requiredSourceHeaders = @(Get-GssSourceHeader)

    foreach ($header in $requiredSourceHeaders) {
        $key = Normalize-Header $header
        if (-not $HeaderMap.ContainsKey($key)) {
            throw "Source workbook is missing required header '$header'."
        }
    }

    $restaurantCol = $HeaderMap[(Normalize-Header 'Restaurant')]
    $ownershipCol = $HeaderMap[(Normalize-Header 'Ownership')]
    $lastRow = $SourceWorksheet.Cells.Item($SourceWorksheet.Rows.Count, $ownershipCol).End($xlUp).Row

    $rows = @()
    for ($row = ($HeaderRow + 1); $row -le $lastRow; $row++) {
        $restaurant = Get-CellString $SourceWorksheet $row $restaurantCol
        $ownership = Get-CellString $SourceWorksheet $row $ownershipCol

        $wanted =
            ($ownership -eq 'All Franchisees' -and [string]::IsNullOrWhiteSpace($restaurant)) -or
            ($ownership -eq 'Sorensen' -and [string]::IsNullOrWhiteSpace($restaurant)) -or
            ($ownership -eq 'Sorensen' -and $restaurant -in @('9354 Richmond', '9355 Virginia Beach'))

        if (-not $wanted) { continue }

        $metricValues = @()
        foreach ($header in $requiredSourceHeaders[2..($requiredSourceHeaders.Count - 1)]) {
            $metricValues += $SourceWorksheet.Cells.Item($row, $HeaderMap[(Normalize-Header $header)]).Value2
        }

        $rows += [pscustomobject]@{
            SourceFile = $SourceFileName
            Week = $WeekEnding
            Restaurant = if ([string]::IsNullOrWhiteSpace($restaurant)) { $null } else { $restaurant }
            Ownership = $ownership
            Metrics = (ConvertTo-GssMetricValue -Values $metricValues -Context "$SourceFileName row $row")
        }
    }

    $actualKeys = @($rows | ForEach-Object {
        if ($null -eq $_.Restaurant) { "$($_.Ownership)|(TOTAL)" } else { "$($_.Ownership)|$($_.Restaurant)" }
    })
    $null = Assert-GssExactEntitySet -ActualKeys $actualKeys -Context "Source workbook $SourceFileName for week ending $($WeekEnding.ToString('yyyy-MM-dd'))"

    return $rows
}

function Get-SourceRowsFromFile {
    param(
        [object]$Excel,
        [object]$Source
    )

    $sourceWb = $null
    $sourceWs = $null
    $beforeHash = Get-GssSha256 $Source.File.FullName
    try {
        $null = Get-GssReportingPeriod -PeriodStart $Source.PeriodStart -PeriodEnd $Source.WeekEnding -Context $Source.File.Name
        $sourceWb = $Excel.Workbooks.Open($Source.File.FullName, 0, $true)
        $sourceWs = $sourceWb.Worksheets.Item(1)
        $title = Get-CellString $sourceWs 1 1
        if ($title -notmatch 'Date Range:\s*(?<start>\d{1,2}/\d{1,2}/\d{4})\s+to\s+(?<end>\d{1,2}/\d{1,2}/\d{4})' -or
            (ConvertTo-GssHeaderDate $Matches.start) -ne $Source.PeriodStart -or
            (ConvertTo-GssHeaderDate $Matches.end) -ne $Source.WeekEnding) {
            throw 'A rolling source reporting period changed during preparation. Run again after uploads finish.'
        }
        $sourceHeaders = Get-HeaderMap $sourceWs $Source.HeaderRow
        $rows = @(Get-SourceRows $sourceWs $sourceHeaders $Source.HeaderRow $Source.WeekEnding $Source.SourceFileName)
    }
    finally {
        if ($sourceWb) { $sourceWb.Close($false) }
        Release-ComObject $sourceWs
        Release-ComObject $sourceWb
    }
    if ((Get-GssSha256 $Source.File.FullName) -ne $beforeHash) {
        throw 'A rolling source changed while it was read. Run again after uploads finish.'
    }
    $Source | Add-Member -NotePropertyName Sha256 -NotePropertyValue $beforeHash -Force
    $Source | Add-Member -NotePropertyName Rows -NotePropertyValue $rows -Force
    $Source | Add-Member -NotePropertyName SemanticSha256 -NotePropertyValue (Get-GssSourceSemanticHash $rows $Source.PeriodStart $Source.WeekEnding) -Force
    return $rows
}

function Get-WeekRowCount {
    param([object]$RawDataWorksheet, [datetime]$WeekEnding)

    $lastRow = $RawDataWorksheet.Cells.Item($RawDataWorksheet.Rows.Count, 1).End($xlUp).Row
    $count = 0
    for ($row = 2; $row -le $lastRow; $row++) {
        $week = Convert-ExcelDate $RawDataWorksheet.Cells.Item($row, 2).Value2
        if ($week -and $week.Date -eq $WeekEnding.Date) {
            $count++
        }
    }
    return $count
}

function Read-GssRawDataRow {
    param([Parameter(Mandatory)][object]$Worksheet)

    $headers = Get-HeaderMap $Worksheet 1
    foreach ($name in (@('Week', 'SourceFile') + @(Get-GssSourceHeader))) {
        if (-not $headers.ContainsKey((Normalize-Header $name))) { throw "Raw_Data is missing required header '$name'." }
    }
    $weekColumn = $headers[(Normalize-Header 'Week')]
    $lastRow = $Worksheet.Cells.Item($Worksheet.Rows.Count, $weekColumn).End($xlUp).Row
    $metricHeaders = @(Get-GssSourceHeader)[2..14]
    for ($row = 2; $row -le $lastRow; $row++) {
        $value = $Worksheet.Cells.Item($row, $weekColumn).Value2
        if ($null -eq $value -or ($value -is [string] -and [string]::IsNullOrWhiteSpace($value))) { continue }
        $week = Convert-ExcelDate $value
        $metrics = @()
        foreach ($metricHeader in $metricHeaders) {
            $metrics += $Worksheet.Cells.Item($row, $headers[(Normalize-Header $metricHeader)]).Value2
        }
        [pscustomobject]@{
            Row = $row
            Week = $week
            SourceFile = Get-CellString $Worksheet $row $headers[(Normalize-Header 'SourceFile')]
            Ownership = Get-CellString $Worksheet $row $headers[(Normalize-Header 'Ownership')]
            Restaurant = Get-CellString $Worksheet $row $headers[(Normalize-Header 'Restaurant')]
            Metrics = $metrics
        }
    }
}

function Set-GssRawDataWeek {
    [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Edits only the caller-owned staged workbook; persistence and verified live promotion are controlled by the transaction entrypoint.')]
    [CmdletBinding()]
    param([Parameter(Mandatory)][object]$Worksheet, [Parameter(Mandatory)][object[]]$Rows)

    if ($Rows.Count -ne 4) { throw 'A corrected reporting week must contain exactly four validated entity rows.' }
    $week = $Rows[0].Week
    $allRows = @(Read-GssRawDataRow $Worksheet)
    $existing = @($allRows | Where-Object { $_.Week.Date -eq $week.Date } | Sort-Object Row)
    # Never delete worksheet rows for corrections: Excel would rewrite formula
    # references elsewhere in the workbook and permanently shrink modeled ranges.
    if ($existing.Count -eq 4) {
        $targetRows = @($existing | ForEach-Object { $_.Row })
        $replacementRows = $Rows
    }
    else {
        $replacementRows = @(@($allRows | Where-Object { $_.Week.Date -ne $week.Date }) + $Rows | Sort-Object Week, Ownership, Restaurant)
        $lastRow = $Worksheet.Cells.Item($Worksheet.Rows.Count, 1).End($xlUp).Row
        $Worksheet.Range("A2:S$([math]::Max($lastRow, $replacementRows.Count + 1))").ClearContents() | Out-Null
        $targetRows = @(2..($replacementRows.Count + 1))
    }
    for ($index = 0; $index -lt $replacementRows.Count; $index++) {
        $sourceRow = $replacementRows[$index]
        $targetRow = $targetRows[$index]
        $values = @($sourceRow.SourceFile, $sourceRow.Week.ToOADate(), $sourceRow.Restaurant, $sourceRow.Ownership) + $sourceRow.Metrics
        for ($col = 0; $col -lt 17; $col++) { Set-CellValue $Worksheet $targetRow ($col + 1) $values[$col] }
        $Worksheet.Cells.Item($targetRow, 2).NumberFormat = 'm/d/yyyy'
        $Worksheet.Cells.Item($targetRow, 18).Formula = '=D' + $targetRow + '&"|"&IF(C' + $targetRow + '="","(TOTAL)",TRIM(C' + $targetRow + '))'
        $Worksheet.Cells.Item($targetRow, 19).Formula = '=R' + $targetRow + '&"|"&TEXT(B' + $targetRow + ',"yyyymmdd")'
    }
    return $Rows.Count
}

function Get-GssWeekEntitySet {
    param([object]$RawDataWorksheet, [datetime]$WeekEnding)

    $headers = Get-HeaderMap $RawDataWorksheet 1
    foreach ($header in @('Week', 'Restaurant', 'Ownership')) {
        if (-not $headers.ContainsKey((Normalize-Header $header))) {
            throw "Raw_Data is missing required header '$header'."
        }
    }
    $weekCol = $headers[(Normalize-Header 'Week')]
    $restaurantCol = $headers[(Normalize-Header 'Restaurant')]
    $ownershipCol = $headers[(Normalize-Header 'Ownership')]
    $lastRow = $RawDataWorksheet.Cells.Item($RawDataWorksheet.Rows.Count, 1).End($xlUp).Row
    $keys = @()
    for ($row = 2; $row -le $lastRow; $row++) {
        $week = Convert-ExcelDate $RawDataWorksheet.Cells.Item($row, $weekCol).Value2
        if ($week -and $week.Date -eq $WeekEnding.Date) {
            $keys += (Get-EntityKey (Get-CellString $RawDataWorksheet $row $ownershipCol) (Get-CellString $RawDataWorksheet $row $restaurantCol))
        }
    }
    return $keys
}

function Test-WeekExists {
    param([object]$RawDataWorksheet, [datetime]$WeekEnding)
    return (Get-WeekRowCount $RawDataWorksheet $WeekEnding) -gt 0
}

function Get-InsertRowForWeek {
    param([object]$RawDataWorksheet, [datetime]$WeekEnding)

    $lastRow = $RawDataWorksheet.Cells.Item($RawDataWorksheet.Rows.Count, 1).End($xlUp).Row
    for ($row = 2; $row -le $lastRow; $row++) {
        $week = Convert-ExcelDate $RawDataWorksheet.Cells.Item($row, 2).Value2
        if ($week -and $week.Date -gt $WeekEnding.Date) {
            return $row
        }
    }
    return ($lastRow + 1)
}

function Get-GssRawDataCapacityPlan {
    param(
        [int]$LastRawDataRow,
        [int]$RowsToAppend,
        [int]$LastModeledRow = $gssRawDataLastFormulaRow
    )

    if ($LastRawDataRow -lt 1) { throw 'LastRawDataRow must include at least the Raw_Data header row.' }
    if ($RowsToAppend -lt 0) { throw 'RowsToAppend cannot be negative.' }
    if ($LastModeledRow -lt 1) { throw 'LastModeledRow must be positive.' }

    $availableRows = [Math]::Max(0, $LastModeledRow - $LastRawDataRow)
    $projectedLastRow = $LastRawDataRow + $RowsToAppend
    return [pscustomobject]@{
        LastRawDataRow = $LastRawDataRow
        RowsToAppend = $RowsToAppend
        AvailableRows = $availableRows
        ProjectedLastRow = $projectedLastRow
        LastModeledRow = $LastModeledRow
        Fits = $projectedLastRow -le $LastModeledRow
    }
}

function Assert-GssRawDataCapacity {
    param(
        [int]$LastRawDataRow,
        [int]$RowsToAppend,
        [int]$LastModeledRow = $gssRawDataLastFormulaRow
    )

    $plan = Get-GssRawDataCapacityPlan $LastRawDataRow $RowsToAppend $LastModeledRow
    if (-not $plan.Fits) {
        throw "Raw_Data modeled capacity is insufficient. $RowsToAppend row(s) are ready to append, but only $($plan.AvailableRows) row(s) remain through row $LastModeledRow. No rows were appended."
    }
    return $plan
}

function Add-RawDataRows {
    param(
        [object]$RawDataWorksheet,
        [object[]]$RowsToAppend
    )

    $rows = @($RowsToAppend)
    if ($rows.Count -eq 0) { return 0 }

    $weekEnding = $rows[0].Week
    $insertRow = Get-InsertRowForWeek $RawDataWorksheet $weekEnding
    $lastRawRow = $RawDataWorksheet.Cells.Item($RawDataWorksheet.Rows.Count, 1).End($xlUp).Row
    $rowCount = $rows.Count
    $endInsertRow = $insertRow + $rowCount - 1

    if ($insertRow -le $lastRawRow) {
        # Move values inside the fixed model. Inserting worksheet rows rewrites
        # dashboard range starts and can exclude the newly inserted older week.
        $shiftedFirstRow = $insertRow + $rowCount
        $shiftedLastRow = $lastRawRow + $rowCount
        $existingValues = $RawDataWorksheet.Range("A$insertRow`:Q$lastRawRow").Value2
        $RawDataWorksheet.Range("A$shiftedFirstRow`:Q$shiftedLastRow").Value2 = $existingValues
        $RawDataWorksheet.Range("B$shiftedFirstRow`:B$shiftedLastRow").NumberFormat = 'm/d/yyyy'
        $RawDataWorksheet.Range("R$shiftedFirstRow`:R$shiftedLastRow").FormulaR1C1 = '=RC[-14]&"|"&IF(RC[-15]="","(TOTAL)",TRIM(RC[-15]))'
        $RawDataWorksheet.Range("S$shiftedFirstRow`:S$shiftedLastRow").FormulaR1C1 = '=RC[-1]&"|"&TEXT(RC[-17],"yyyymmdd")'
    }

    $templateRow = if ($insertRow -gt 2) { $insertRow - 1 } else { $endInsertRow + 1 }
    $rowsWritten = 0
    foreach ($sourceRow in $rows) {
        $targetRow = $insertRow + $rowsWritten

        $RawDataWorksheet.Range("A$templateRow`:S$templateRow").Copy() | Out-Null
        $RawDataWorksheet.Range("A$targetRow`:S$targetRow").PasteSpecial($xlPasteFormats) | Out-Null

        $values = @(
            $sourceRow.SourceFile,
            $sourceRow.Week.ToOADate(),
            $sourceRow.Restaurant,
            $sourceRow.Ownership
        ) + $sourceRow.Metrics

        $rowMatrix = New-Object 'object[,]' 1, 17
        for ($col = 1; $col -le 17; $col++) {
            $rowMatrix[0, ($col - 1)] = $values[$col - 1]
        }
        $RawDataWorksheet.Range("A$targetRow`:Q$targetRow").Value2 = $rowMatrix

        $RawDataWorksheet.Cells.Item($targetRow, 2).NumberFormat = 'm/d/yyyy'
        $RawDataWorksheet.Cells.Item($targetRow, 18).Formula = '=D' + $targetRow + '&"|"&IF(C' + $targetRow + '="","(TOTAL)",TRIM(C' + $targetRow + '))'
        $RawDataWorksheet.Cells.Item($targetRow, 19).Formula = '=R' + $targetRow + '&"|"&TEXT(B' + $targetRow + ',"yyyymmdd")'

        $rowsWritten++
    }

    return $rowsWritten
}

function Get-EntityKey {
    param([string]$Ownership, [string]$Restaurant)

    $restaurantName = $Restaurant
    if ([string]::IsNullOrWhiteSpace($restaurantName)) {
        $restaurantName = '(TOTAL)'
    }
    else {
        $restaurantName = $restaurantName.Trim()
    }

    return "$($Ownership.Trim())|$restaurantName"
}

function Get-EntityWeekKey {
    param([string]$EntityKey, [datetime]$WeekEnding)
    return "$EntityKey|$($WeekEnding.ToString('yyyyMMdd'))"
}

function Get-MetricCatalogRows {
    param([object]$MetricCatalogWorksheet)

    $lastRow = $MetricCatalogWorksheet.Cells.Item($MetricCatalogWorksheet.Rows.Count, 1).End($xlUp).Row
    $metrics = @()
    for ($row = 2; $row -le $lastRow; $row++) {
        $displayName = Get-CellString $MetricCatalogWorksheet $row 1
        $rawKey = Get-CellString $MetricCatalogWorksheet $row 2
        if ([string]::IsNullOrWhiteSpace($displayName) -or [string]::IsNullOrWhiteSpace($rawKey)) {
            continue
        }

        $metrics += [pscustomobject]@{
            DisplayName = $displayName
            RawKey = $rawKey
            LowerIsBetter = Convert-ToBool $MetricCatalogWorksheet.Cells.Item($row, 3).Value2
            IncludeInMovers = Convert-ToBool $MetricCatalogWorksheet.Cells.Item($row, 4).Value2
            Category = Get-CellString $MetricCatalogWorksheet $row 5
            Owner = Get-CellString $MetricCatalogWorksheet $row 6
        }
    }

    return $metrics
}

function Get-GssConfigurationCounts {
    param([object]$Workbook)

    $entityWs = $null
    $metricWs = $null
    try {
        $entityWs = $Workbook.Worksheets.Item('Entities')
        $metricWs = $Workbook.Worksheets.Item('Metric_Catalog')
        $entityLastRow = $entityWs.Cells.Item($entityWs.Rows.Count, 1).End($xlUp).Row
        $metricLastRow = $metricWs.Cells.Item($metricWs.Rows.Count, 1).End($xlUp).Row
        $activeEntityCount = 0
        for ($row = 2; $row -le $entityLastRow; $row++) {
            if ((Get-CellString $entityWs $row 1) -and (Convert-ToBool $entityWs.Cells.Item($row, 5).Value2)) {
                $activeEntityCount++
            }
        }

        $metricCount = 0
        for ($row = 2; $row -le $metricLastRow; $row++) {
            if ((Get-CellString $metricWs $row 1) -and (Get-CellString $metricWs $row 2)) {
                $metricCount++
            }
        }

        return [pscustomobject]@{
            EntityLastRow = $entityLastRow
            ActiveEntityCount = $activeEntityCount
            MetricLastRow = $metricLastRow
            MetricCount = $metricCount
        }
    }
    finally {
        Release-ComObject $metricWs
        Release-ComObject $entityWs
    }
}

function Remove-GssWorkbookGuardrails {
    param([object]$Workbook)

    if ($Workbook.ProtectStructure) {
        $Workbook.Unprotect('')
    }

    for ($index = 1; $index -le $Workbook.Worksheets.Count; $index++) {
        $worksheet = $null
        try {
            $worksheet = $Workbook.Worksheets.Item($index)
            if ($worksheet.ProtectContents -or $worksheet.ProtectDrawingObjects -or $worksheet.ProtectScenarios) {
                $worksheet.Unprotect('')
            }
        }
        finally {
            Release-ComObject $worksheet
        }
    }
}

function Set-GssReadmeContent {
    param([object]$Workbook)

    $worksheet = $null
    $sectionRange = $null
    $contentRange = $null
    try {
        $worksheet = $Workbook.Worksheets.Item('README')
        Set-CellValue $worksheet 5 1 '1) Save the newest Sorensen rolling survey workbook in 02 Weekly Rolling Source Workbooks.'
        Set-CellValue $worksheet 6 1 '2) Close this workbook, then double-click Run GSS Update After Upload.cmd in the GSS Surveys folder.'
        Set-CellValue $worksheet 7 1 '3) The program validates a staged report and backs up the previous report before updating. Choose a revision only if it finds conflicting or corrected inputs.'
        Set-CellValue $worksheet 8 1 '4) Reopen this workbook and check QA Checks. Use the reports when the overall status is READY.'
        Set-CellValue $worksheet 9 1 'The automation recalculates the workbook, refreshes the email PDF, protects formulas, and creates a backup before live changes. Do not paste into Raw_Data manually.'
        Set-CellValue $worksheet 12 1 '1) Previous rolling-window view (legacy tab name: Quick_Read_WoW): compare each 13-week rolling result with the previous rolling window across all active entities.'
        Set-CellValue $worksheet 13 1 '2) Quick_Read_YoY: compare the same 13-week rolling result with the matching prior-year rolling window.'
        Set-CellValue $worksheet 14 1 '3) Exec_Dashboard: use the Entity and Metric dropdowns to drill into one story.'
        Set-CellValue $worksheet 15 1 '4) Drivers_Detail: review every metric for the entity selected on Exec_Dashboard.'
        Set-CellValue $worksheet 18 1 '• Previous rolling-window view (legacy tab name: Quick_Read_WoW): positive change versus the previous rolling window means directional improvement. A blank is surfaced by QA Checks when required source data is missing.'
        Set-CellValue $worksheet 19 1 '• Quick_Read_YoY: compares the reporting-date rolling result with the rolling result ending 364 days earlier. The confirmed 2025 history break is informational, not an error.'
        Set-CellValue $worksheet 20 1 '• Exec_Dashboard: selector-driven view. The same selections drive Drivers_Detail and the movers table.'
        Set-CellValue $worksheet 21 1 '• QA Checks: the visible READY/ATTENTION gate checks row keys, period completeness, headers, selectors, and remaining capacity.'
        Set-CellValue $worksheet 22 1 'Metric definitions:'
        Set-CellValue $worksheet 23 1 '• Overall Experience, Service, Culinary, Value, and Pace of Meal are vendor topbox scores: the percentage rated 5 (Excellent). Dissat is the percentage rated 1 (Poor), 2 (Fair), or 3 (Good).'
        Set-CellValue $worksheet 24 1 '• Recommend means the percentage rating 9 or 10. The current rolling export supplies its response count but not its percentage, so this workbook does not present Recommend as a population score.'
        Set-CellValue $worksheet 25 1 '• Steak Cooked Properly is the percentage of steak-ordering guests who answered Yes when asked whether the steak was cooked properly the first time.'
        Set-CellValue $worksheet 26 1 'Data hygiene rules (checked automatically):'
        Set-CellValue $worksheet 27 1 '• Weekly comment ratings: 5=Excellent, 4=Very Good, 3=Good, 2=Fair, 1=Poor; steak answers use Y/N.'
        Set-CellValue $worksheet 28 1 '• Alert Guests GR (Guest Relations already contacted) and NC (no contact requested) are excluded from individual outreach suggestions. Restaurant Manager Follow Up remains a manual notes field.'
        Set-CellValue $worksheet 29 1 '• The break between the 7/6/2025 and 10/5/2025 data points is intentional and does not block READY status.'
        Set-CellValue $worksheet 31 1 '• ATTENTION status: stop using the reports and share the newest file from _automation_runs\logs with the workbook maintainer.'
        Set-CellValue $worksheet 32 1 '• Dropdown rejects a choice: select an active entity or a metric shown in the dropdown list.'
        Set-CellValue $worksheet 34 1 '• Current model capacity is Raw_Data rows 2:5096. QA Checks warns before fewer than two weekly loads (8 rows) remain.'

        Set-CellFormula $worksheet 3 1 '="Workbook status: "&''QA Checks''!$B$2&" | 13-week rolling through: "&TEXT(''QA Checks''!$B$3,"m/d/yyyy")'

        $contentRange = $worksheet.Range('A1:A34')
        $contentRange.WrapText = $true
        $contentRange.VerticalAlignment = -4160
        $contentRange.Font.Name = 'Aptos'
        $contentRange.Font.Size = 10
        $worksheet.Columns.Item(1).ColumnWidth = 112
        $contentRange.Rows.AutoFit() | Out-Null

        $worksheet.Range('A1').Font.Size = 16
        $worksheet.Range('A1').Font.Bold = $true
        $worksheet.Range('A1').Font.Color = (Get-GssExcelColor 255 255 255)
        $worksheet.Range('A1').Interior.Color = (Get-GssExcelColor 31 78 121)
        foreach ($row in @(4, 11, 17, 23, 26, 30)) {
            $sectionRange = $worksheet.Range("A$row")
            $sectionRange.Font.Bold = $true
            $sectionRange.Interior.Color = (Get-GssExcelColor 221 235 247)
            Release-ComObject $sectionRange
            $sectionRange = $null
        }
    }
    finally {
        Release-ComObject $sectionRange
        Release-ComObject $contentRange
        Release-ComObject $worksheet
    }
}

function Set-GssBanner {
    param(
        [object]$Worksheet,
        [string]$Address,
        [string]$Formula,
        [switch]$Merge
    )

    $range = $null
    try {
        $range = $Worksheet.Range($Address)
        if ($Merge) {
            if ($range.MergeCells) { $range.UnMerge() }
            $range.ClearContents() | Out-Null
            $range.Merge() | Out-Null
        }
        try {
            $range.Cells.Item(1, 1).Formula2 = $Formula
        }
        catch {
            $range.Cells.Item(1, 1).Formula = $Formula
        }
        $range.Font.Name = 'Aptos'
        $range.Font.Bold = $true
        $range.Font.Size = 10
        $range.HorizontalAlignment = -4131
        $range.VerticalAlignment = -4108
        $range.WrapText = $true
        $range.RowHeight = 24
    }
    finally {
        Release-ComObject $range
    }
}

function Set-GssWorkbookBanners {
    param([object]$Workbook)

    $worksheet = $null
    try {
        $worksheet = $Workbook.Worksheets.Item('Quick_Read_WoW')
        Set-GssBanner $worksheet 'A4:Q4' '="Workbook status: "&''QA Checks''!$B$2&" | Positive = improvement (direction-adjusted for dissatisfaction). Blank = missing data."' -Merge
        Release-ComObject $worksheet
        $worksheet = $Workbook.Worksheets.Item('Quick_Read_YoY')
        Set-GssBanner $worksheet 'A4:Q4' '="Workbook status: "&''QA Checks''!$B$2&" | Positive = directional improvement versus the matching prior-year rolling window. Blank = missing data."' -Merge
        Release-ComObject $worksheet
        $worksheet = $Workbook.Worksheets.Item('Exec_Dashboard')
        Set-GssBanner $worksheet 'A2:F2' '="Workbook status: "&''QA Checks''!$B$2&" | Change the two dropdowns below to explore the reports."' -Merge
        Release-ComObject $worksheet
        $worksheet = $Workbook.Worksheets.Item('Drivers_Detail')
        Set-GssBanner $worksheet 'A2:D2' '="Workbook status: "&''QA Checks''!$B$2&" | This table follows the Executive Dashboard entity selection."' -Merge
        Release-ComObject $worksheet
        $worksheet = $Workbook.Worksheets.Item('Email Comparison')
        Set-GssBanner $worksheet 'A2:AB2' '="Workbook status: "&''QA Checks''!$B$2&" | Ready-to-share comparison output."' -Merge
    }
    finally {
        Release-ComObject $worksheet
    }
}

function Set-GssEmailComparisonLayout {
    param([object]$Workbook)

    $worksheet = $null
    $range = $null
    $plan = Get-GssEmailComparisonLayoutPlan
    try {
        $worksheet = $Workbook.Worksheets.Item('Email Comparison')

        foreach ($address in $plan.WrappedTableHeaderRanges) {
            $range = $worksheet.Range($address)
            $range.WrapText = $true
            Release-ComObject $range
            $range = $null
        }

        foreach ($column in $plan.VisibleColumnWidths) {
            $range = $worksheet.Columns.Item($column.Column)
            $range.ColumnWidth = $column.Width
            Release-ComObject $range
            $range = $null
        }
    }
    finally {
        Release-ComObject $range
        Release-ComObject $worksheet
    }
}

function Update-GssQaSheet {
    param(
        [object]$Workbook,
        [object]$Configuration
    )

    $qaWs = $null
    $readmeWs = $null
    $range = $null
    try {
        $readmeWs = $Workbook.Worksheets.Item('README')
        try {
            $qaWs = $Workbook.Worksheets.Item('QA Checks')
        }
        catch {
            $qaWs = $null
        }

        if ($qaWs -and $qaWs.Index -ne ($readmeWs.Index + 1)) {
            $qaWs.Delete()
            Release-ComObject $qaWs
            $qaWs = $null
        }
        if (-not $qaWs) {
            $qaWs = $Workbook.Worksheets.Add([System.Type]::Missing, $readmeWs)
            $qaWs.Name = 'QA Checks'
        }
        $qaWs.Visible = $xlSheetVisible
        $qaWs.Cells.Clear() | Out-Null

        Set-CellValue $qaWs 1 1 'GSS Workbook QA Checks'
        Set-CellValue $qaWs 2 1 'Overall workbook status'
        Set-CellFormula $qaWs 2 2 '=IFERROR(IF(COUNTIF($B$6:$B$12,"READY")=ROWS($B$6:$B$12),"READY","ATTENTION"),"ATTENTION")'
        Set-CellValue $qaWs 3 1 '13-week rolling through'
        Set-CellFormula $qaWs 3 2 "=MAX('Raw_Data'!`$B`$2:`$B`$$gssRawDataLastFormulaRow)"
        $qaWs.Cells.Item(3, 2).NumberFormat = 'm/d/yyyy'
        Set-CellValue $qaWs 4 1 'Confirmed history note'
        Set-CellValue $qaWs 4 2 'The break between the 7/6/2025 and 10/5/2025 data points is intentional and is not a blocker.'
        $qaWs.Range('B4:C4').Merge() | Out-Null
        Set-CellValue $qaWs 5 1 'Check'
        Set-CellValue $qaWs 5 2 'Status'
        Set-CellValue $qaWs 5 3 'Detail'

        foreach ($check in @(Get-GssQaCheckPlan $Configuration.EntityLastRow $Configuration.MetricLastRow $gssRawDataLastFormulaRow)) {
            Set-CellValue $qaWs $check.Row 1 $check.Name
            Set-CellFormula $qaWs $check.Row 2 $check.Formula
            Set-CellFormula $qaWs $check.Row 3 $check.DetailFormula
        }

        $range = $qaWs.Range('A1:C12')
        $range.Font.Name = 'Aptos'
        $range.Font.Size = 10
        $range.VerticalAlignment = -4108
        $range.WrapText = $true
        Release-ComObject $range
        $range = $qaWs.Range('A1:C1')
        $range.Merge() | Out-Null
        $range.Font.Size = 16
        $range.Font.Bold = $true
        $range.Font.Color = (Get-GssExcelColor 255 255 255)
        $range.Interior.Color = (Get-GssExcelColor 31 78 121)
        Release-ComObject $range
        $range = $qaWs.Range('A5:C5')
        $range.Font.Bold = $true
        $range.Font.Color = (Get-GssExcelColor 255 255 255)
        $range.Interior.Color = (Get-GssExcelColor 68 114 196)
        Release-ComObject $range
        $range = $qaWs.Range('A2:B2')
        $range.Font.Bold = $true
        $range.Font.Size = 12
        Release-ComObject $range
        $range = $qaWs.Range('A4:C4')
        $range.Interior.Color = (Get-GssExcelColor 255 242 204)
        Release-ComObject $range
        $range = $qaWs.Range('A6:C12')
        $range.Borders.LineStyle = 1
        Release-ComObject $range
        $qaWs.Columns.Item(1).ColumnWidth = 42
        $qaWs.Columns.Item(2).ColumnWidth = 18
        $qaWs.Columns.Item(3).ColumnWidth = 76
        $qaWs.Rows.Item('1:12').AutoFit() | Out-Null
        $qaWs.Tab.Color = (Get-GssExcelColor 112 173 71)
    }
    finally {
        Release-ComObject $range
        Release-ComObject $readmeWs
        Release-ComObject $qaWs
    }
}

function Set-GssWorkbookPrintLayouts {
    param([object]$Workbook, [object]$Configuration)

    foreach ($layout in @(Get-GssWorkbookLayoutPlan $Configuration.ActiveEntityCount $Configuration.MetricCount)) {
        $worksheet = $null
        $pageSetup = $null
        try {
            $worksheet = $Workbook.Worksheets.Item($layout.Sheet)
            $pageSetup = $worksheet.PageSetup
            $pageSetup.PrintArea = $layout.PrintArea
            $pageSetup.Orientation = $layout.Orientation
            $pageSetup.Zoom = $false
            $pageSetup.FitToPagesWide = $layout.FitWide
            $pageSetup.FitToPagesTall = $layout.FitTall
            $pageSetup.PrintGridlines = $false
            $pageSetup.CenterHorizontally = $true
            $pageSetup.CenterVertically = $false
        }
        finally {
            Release-ComObject $pageSetup
            Release-ComObject $worksheet
        }
    }
}

function Set-GssDashboardValidation {
    param([object]$Workbook, [object]$Configuration)

    $worksheet = $null
    $cell = $null
    $validation = $null
    try {
        $worksheet = $Workbook.Worksheets.Item('Exec_Dashboard')
        foreach ($definition in @(
            [pscustomobject]@{ Address = 'C3'; Formula = "=Entities!`$A`$2:`$A`$$($Configuration.EntityLastRow)"; Message = 'Choose an active entity from the dropdown.' },
            [pscustomobject]@{ Address = 'F3'; Formula = "=Metric_Catalog!`$A`$2:`$A`$$($Configuration.MetricLastRow)"; Message = 'Choose a metric from the dropdown.' }
        )) {
            $cell = $worksheet.Range($definition.Address)
            $cell.Locked = $false
            $validation = $cell.Validation
            try { $validation.Delete() } catch { }
            $validation.Add($xlValidateList, $xlValidAlertStop, $xlBetween, $definition.Formula)
            $validation.IgnoreBlank = $false
            $validation.InCellDropdown = $true
            $validation.ShowInput = $true
            $validation.ShowError = $true
            $validation.InputTitle = 'Dashboard selection'
            $validation.InputMessage = $definition.Message
            $validation.ErrorTitle = 'Invalid dashboard selection'
            $validation.ErrorMessage = $definition.Message
            Release-ComObject $validation
            $validation = $null
            Release-ComObject $cell
            $cell = $null
        }
    }
    finally {
        Release-ComObject $validation
        Release-ComObject $cell
        Release-ComObject $worksheet
    }
}

function Set-GssStatusStyles {
    param([object]$Workbook, [string]$WorkbookStatus)

    $isReady = $WorkbookStatus -eq 'READY'
    $fillColor = if ($isReady) { Get-GssExcelColor 226 239 218 } else { Get-GssExcelColor 252 228 214 }
    $fontColor = if ($isReady) { Get-GssExcelColor 0 97 0 } else { Get-GssExcelColor 156 0 6 }
    foreach ($definition in @(
        [pscustomobject]@{ Sheet = 'README'; Address = 'A3' },
        [pscustomobject]@{ Sheet = 'QA Checks'; Address = 'A2:B2' },
        [pscustomobject]@{ Sheet = 'Quick_Read_WoW'; Address = 'A4:Q4' },
        [pscustomobject]@{ Sheet = 'Quick_Read_YoY'; Address = 'A4:Q4' },
        [pscustomobject]@{ Sheet = 'Exec_Dashboard'; Address = 'A2:F2' },
        [pscustomobject]@{ Sheet = 'Drivers_Detail'; Address = 'A2:D2' },
        [pscustomobject]@{ Sheet = 'Email Comparison'; Address = 'A2:AB2' }
    )) {
        $worksheet = $null
        $range = $null
        try {
            $worksheet = $Workbook.Worksheets.Item($definition.Sheet)
            $range = $worksheet.Range($definition.Address)
            $range.Interior.Color = $fillColor
            $range.Font.Color = $fontColor
            $range.Font.Bold = $true
        }
        finally {
            Release-ComObject $range
            Release-ComObject $worksheet
        }
    }
}

function Get-GssWorkbookStatus {
    param([object]$Workbook)

    $worksheet = $null
    $cell = $null
    try {
        $worksheet = $Workbook.Worksheets.Item('QA Checks')
        $cell = $worksheet.Range('B2')
        return ([string]$cell.Value2).Trim().ToUpperInvariant()
    }
    finally {
        Release-ComObject $cell
        Release-ComObject $worksheet
    }
}

function Set-GssWorkbookGuardrails {
    param([object]$Workbook, [object]$Configuration)

    for ($index = 1; $index -le $Workbook.Worksheets.Count; $index++) {
        $worksheet = $null
        $allCells = $null
        try {
            $worksheet = $Workbook.Worksheets.Item($index)
            $allCells = $worksheet.Cells
            $allCells.Locked = $true
        }
        finally {
            Release-ComObject $allCells
            Release-ComObject $worksheet
        }
    }

    Set-GssDashboardValidation $Workbook $Configuration

    for ($index = 1; $index -le $Workbook.Worksheets.Count; $index++) {
        $worksheet = $null
        try {
            $worksheet = $Workbook.Worksheets.Item($index)
            $worksheet.Protect('', $true, $true, $true, $true, $false, $false, $false, $false, $false, $false, $false, $false, $false, $true, $false)
            $worksheet.EnableSelection = if ($worksheet.Name -eq 'Exec_Dashboard') { $xlUnlockedCells } else { 0 }
        }
        finally {
            Release-ComObject $worksheet
        }
    }

    $Workbook.Protect('', $true, $false)
}

function Test-GssWorkbookGuardrailsApplied {
    param([object]$Workbook)

    if (-not $Workbook.ProtectStructure) { return $false }
    for ($index = 1; $index -le $Workbook.Worksheets.Count; $index++) {
        $worksheet = $null
        try {
            $worksheet = $Workbook.Worksheets.Item($index)
            if (-not $worksheet.ProtectContents) { return $false }
        }
        finally {
            Release-ComObject $worksheet
        }
    }

    $execWs = $null
    $entityCell = $null
    $metricCell = $null
    try {
        $execWs = $Workbook.Worksheets.Item('Exec_Dashboard')
        $entityCell = $execWs.Range('C3')
        $metricCell = $execWs.Range('F3')
        return (-not [bool]$entityCell.Locked) -and (-not [bool]$metricCell.Locked)
    }
    finally {
        Release-ComObject $metricCell
        Release-ComObject $entityCell
        Release-ComObject $execWs
    }
}

function Get-RawDataIndex {
    param([object]$RawDataWorksheet)

    $headers = Get-HeaderMap $RawDataWorksheet 1
    $requiredHeaders = @('Week', 'Restaurant', 'Ownership', 'Count')
    foreach ($header in $requiredHeaders) {
        $key = Normalize-Header $header
        if (-not $headers.ContainsKey($key)) {
            throw "Raw_Data is missing required header '$header'."
        }
    }

    $weekCol = $headers[(Normalize-Header 'Week')]
    $restaurantCol = $headers[(Normalize-Header 'Restaurant')]
    $ownershipCol = $headers[(Normalize-Header 'Ownership')]
    $lastRow = $RawDataWorksheet.Cells.Item($RawDataWorksheet.Rows.Count, 1).End($xlUp).Row
    $rowsByKey = @{}

    for ($row = 2; $row -le $lastRow; $row++) {
        $week = Convert-ExcelDate $RawDataWorksheet.Cells.Item($row, $weekCol).Value2
        if ($null -eq $week) { continue }

        $ownership = Get-CellString $RawDataWorksheet $row $ownershipCol
        if ([string]::IsNullOrWhiteSpace($ownership)) { continue }

        $restaurant = Get-CellString $RawDataWorksheet $row $restaurantCol
        $entityKey = Get-EntityKey $ownership $restaurant
        $rowsByKey[(Get-EntityWeekKey $entityKey $week)] = $row
    }

    return [pscustomobject]@{
        Headers = $headers
        RowsByKey = $rowsByKey
    }
}

function Get-RawMetricValue {
    param(
        [object]$RawDataWorksheet,
        [object]$RawIndex,
        [string]$EntityKey,
        [datetime]$WeekEnding,
        [string]$MetricKey
    )

    $rowKey = Get-EntityWeekKey $EntityKey $WeekEnding
    if (-not $RawIndex.RowsByKey.ContainsKey($rowKey)) {
        return $null
    }

    $headerKey = Normalize-Header $MetricKey
    if (-not $RawIndex.Headers.ContainsKey($headerKey)) {
        throw "Raw_Data is missing metric column '$MetricKey'."
    }

    $row = $RawIndex.RowsByKey[$rowKey]
    $col = $RawIndex.Headers[$headerKey]
    return Convert-ToNullableDouble $RawDataWorksheet.Cells.Item($row, $col).Value2
}

function Get-AverageValue {
    param([object[]]$Values)

    $usable = @($Values | Where-Object { $null -ne $_ })
    if ($usable.Count -eq 0) { return $null }
    return ($usable | Measure-Object -Average).Average
}

function Get-Improvement {
    param([object]$CurrentValue, [object]$PriorValue, [bool]$LowerIsBetter)

    if ($null -eq $CurrentValue -or $null -eq $PriorValue) { return $null }
    if ($LowerIsBetter) {
        return ([double]$PriorValue - [double]$CurrentValue)
    }
    return ([double]$CurrentValue - [double]$PriorValue)
}

function Get-ThreeWeekTrend {
    param(
        [object]$RawDataWorksheet,
        [object]$RawIndex,
        [string]$EntityKey,
        [datetime]$LatestWeek,
        [string]$MetricKey,
        [bool]$LowerIsBetter,
        [bool]$IncludeInMovers
    )

    if (-not $IncludeInMovers) {
        return [pscustomobject]@{ Arrow = ''; Average = $null; Persistent = '' }
    }

    $improvements = @()
    foreach ($weekOffset in @(0, -7, -14)) {
        $week = $LatestWeek.AddDays($weekOffset)
        $current = Get-RawMetricValue $RawDataWorksheet $RawIndex $EntityKey $week $MetricKey
        $prior = Get-RawMetricValue $RawDataWorksheet $RawIndex $EntityKey $week.AddDays(-7) $MetricKey
        $improvements += Get-Improvement $current $prior $LowerIsBetter
    }

    $average = Get-AverageValue $improvements
    if ($null -eq $average -or [math]::Abs($average) -lt 0.0000001) {
        return [pscustomobject]@{ Arrow = ''; Average = $average; Persistent = '' }
    }

    $arrow = if ($average -gt 0) { [string][char]0x25B2 } else { [string][char]0x25BC }
    $usable = @($improvements | Where-Object { $null -ne $_ })
    $persistent = ''
    if ($usable.Count -eq 3) {
        $positiveCount = @($usable | Where-Object { $_ -gt 0 }).Count
        $negativeCount = @($usable | Where-Object { $_ -lt 0 }).Count
        if ($positiveCount -eq 3 -or $negativeCount -eq 3) {
            $persistent = $average
        }
    }

    return [pscustomobject]@{ Arrow = $arrow; Average = $average; Persistent = $persistent }
}

function Update-EmailComparisonSection {
    param(
        [object]$Worksheet,
        [object]$RawDataWorksheet,
        [object]$RawIndex,
        [object[]]$Metrics,
        [datetime]$LatestWeek,
        [string]$EntityKey,
        [string]$EntityName,
        [int]$HeaderRow,
        [int]$FirstMetricRow,
        [string]$CompareEntityKey,
        [string]$CompareEntityName
    )

    Set-CellValue $Worksheet $HeaderRow 1 'Entity'
    Set-CellValue $Worksheet $HeaderRow 2 $EntityName
    Set-CellValue $Worksheet $HeaderRow 4 '13-Week Rolling Through'
    Set-CellValue $Worksheet $HeaderRow 5 $LatestWeek.ToOADate()
    $Worksheet.Cells.Item($HeaderRow, 5).NumberFormat = 'm/d/yyyy'
    Set-CellValue $Worksheet $HeaderRow 8 $CompareEntityKey
    Set-CellValue $Worksheet $HeaderRow 18 $LatestWeek.ToOADate()
    $Worksheet.Cells.Item($HeaderRow, 18).NumberFormat = 'm/d/yyyy'
    Set-CellValue $Worksheet ($HeaderRow + 1) 1 'Compare To'
    Set-CellValue $Worksheet ($HeaderRow + 1) 2 $CompareEntityName
    Set-CellValue $Worksheet ($HeaderRow + 2) 2 '13-Week Rolling Result'
    Set-CellValue $Worksheet ($HeaderRow + 2) 3 'Change vs Previous Rolling Window'
    Set-CellValue $Worksheet ($HeaderRow + 2) 4 'Change vs Prior-Year Rolling Window'
    Set-CellValue $Worksheet ($HeaderRow + 2) 5 ''

    $latestMinus7 = $LatestWeek.AddDays(-7)
    $latestMinus14 = $LatestWeek.AddDays(-14)
    $latestMinus21 = $LatestWeek.AddDays(-21)
    $priorYearWeek = $LatestWeek.AddDays(-364)
    $currentCount = Get-RawMetricValue $RawDataWorksheet $RawIndex $EntityKey $LatestWeek 'Count'

    for ($i = 0; $i -lt $Metrics.Count; $i++) {
        $metric = $Metrics[$i]
        $row = $FirstMetricRow + $i
        $current = Get-RawMetricValue $RawDataWorksheet $RawIndex $EntityKey $LatestWeek $metric.RawKey
        $priorWeek = Get-RawMetricValue $RawDataWorksheet $RawIndex $EntityKey $latestMinus7 $metric.RawKey
        $priorYear = Get-RawMetricValue $RawDataWorksheet $RawIndex $EntityKey $priorYearWeek $metric.RawKey
        $twoWeeksAgo = Get-RawMetricValue $RawDataWorksheet $RawIndex $EntityKey $latestMinus14 $metric.RawKey
        $threeWeeksAgo = Get-RawMetricValue $RawDataWorksheet $RawIndex $EntityKey $latestMinus21 $metric.RawKey
        $rollingAverage = $null
        $wowImprovement = if ($metric.DisplayName -eq 'Count') { $null } else { Get-Improvement $current $priorWeek $metric.LowerIsBetter }
        $yoyImprovement = Get-Improvement $current $priorYear $metric.LowerIsBetter

        $compareCurrent = Get-RawMetricValue $RawDataWorksheet $RawIndex $CompareEntityKey $LatestWeek $metric.RawKey
        $comparePriorYear = Get-RawMetricValue $RawDataWorksheet $RawIndex $CompareEntityKey $priorYearWeek $metric.RawKey
        $compareYoyImprovement = Get-Improvement $compareCurrent $comparePriorYear $metric.LowerIsBetter
        $vsCompareCurrent = Get-Improvement $current $compareCurrent $metric.LowerIsBetter
        $vsCompareYoy = if ($null -eq $yoyImprovement -or $null -eq $compareYoyImprovement) { $null } else { $yoyImprovement - $compareYoyImprovement }
        $trend = Get-ThreeWeekTrend $RawDataWorksheet $RawIndex $EntityKey $LatestWeek $metric.RawKey $metric.LowerIsBetter $metric.IncludeInMovers
        $currentBadness = if ($null -eq $current -or $metric.DisplayName -eq 'Count') { $null } elseif ($metric.LowerIsBetter) { $current } else { 100 - $current }
        $yoyBadness = if ($null -eq $yoyImprovement -or $yoyImprovement -ge 0) { $null } else { -1 * $yoyImprovement }

        Set-CellValue $Worksheet $row 1 $metric.DisplayName
        Set-CellValue $Worksheet $row 2 $current
        Set-CellValue $Worksheet $row 3 $wowImprovement
        Set-CellValue $Worksheet $row 4 $yoyImprovement
        Set-CellValue $Worksheet $row 5 $rollingAverage
        Set-CellValue $Worksheet $row 6 $currentCount
        Set-CellValue $Worksheet $row 7 $metric.RawKey
        Set-CellValue $Worksheet $row 8 $metric.LowerIsBetter
        Set-CellValue $Worksheet $row 9 $metric.IncludeInMovers
        Set-CellValue $Worksheet $row 10 $priorWeek
        Set-CellValue $Worksheet $row 11 $priorYear
        Set-CellValue $Worksheet $row 12 $twoWeeksAgo
        Set-CellValue $Worksheet $row 13 $threeWeeksAgo
        Set-CellValue $Worksheet $row 14 $metric.Category
        Set-CellValue $Worksheet $row 15 $metric.Owner
        Set-CellValue $Worksheet $row 16 ''
        Set-CellValue $Worksheet $row 17 ''
        Set-CellValue $Worksheet $row 18 $trend.Arrow
        Set-CellValue $Worksheet $row 19 $yoyBadness
        Set-CellValue $Worksheet $row 20 $currentBadness
        Set-CellValue $Worksheet $row 21 ''
        Set-CellValue $Worksheet $row 22 ''
        Set-CellValue $Worksheet $row 23 $trend.Average
        Set-CellValue $Worksheet $row 24 $trend.Persistent

        if ($metric.DisplayName -eq 'Count') {
            Set-CellValue $Worksheet $row 25 ''
            Set-CellValue $Worksheet $row 26 ''
            Set-CellValue $Worksheet $row 27 ''
            Set-CellValue $Worksheet $row 28 ''
        }
        else {
            Set-CellValue $Worksheet $row 25 $compareCurrent
            Set-CellValue $Worksheet $row 26 $compareYoyImprovement
            Set-CellValue $Worksheet $row 27 $vsCompareCurrent
            Set-CellValue $Worksheet $row 28 $vsCompareYoy
        }
    }
}

function Update-EmailComparisonData {
    param(
        [object]$Workbook,
        [object]$RawDataWorksheet,
        [object]$LatestSource
    )

    $emailWs = $null
    $metricWs = $null
    try {
        $emailWs = $Workbook.Worksheets.Item('Email Comparison')
        $metricWs = $Workbook.Worksheets.Item('Metric_Catalog')
        $metrics = @(Get-MetricCatalogRows $metricWs)
        if ($metrics.Count -eq 0) {
            throw 'Metric_Catalog did not contain any metrics for the email comparison sheet.'
        }

        $rawIndex = Get-RawDataIndex $RawDataWorksheet
        $compareEntityKey = 'All Franchisees|(TOTAL)'
        $compareEntityName = 'TOTAL - All Franchisees'
        $latestWeek = $LatestSource.WeekEnding

        Update-EmailComparisonSection $emailWs $RawDataWorksheet $rawIndex $metrics $latestWeek 'Sorensen|9355 Virginia Beach' '9355 Virginia Beach (Sorensen)' 3 6 $compareEntityKey $compareEntityName
        Update-EmailComparisonSection $emailWs $RawDataWorksheet $rawIndex $metrics $latestWeek 'Sorensen|9354 Richmond' '9354 Richmond (Sorensen)' 23 26 $compareEntityKey $compareEntityName
    }
    finally {
        Release-ComObject $metricWs
        Release-ComObject $emailWs
    }
}

function Export-EmailComparisonPdf {
    param(
        [object]$Workbook,
        [object]$LatestSource,
        [string]$FolderPath,
        [string]$TestDir,
        [bool]$ApplyMode
    )

    $emailWs = $null
    try {
        $emailWs = $Workbook.Worksheets.Item('Email Comparison')
        $pdfDir = if ($ApplyMode) { Join-Path $FolderPath '04 Email Comparison PDFs' } else { $TestDir }
        New-Item -ItemType Directory -Path $pdfDir -Force | Out-Null
        $pdfName = 'GSS Email Comparison {0}.pdf' -f $LatestSource.WeekEnding.ToString('MMddyy')
        $pdfPath = Join-Path $pdfDir $pdfName
        $emailWs.ExportAsFixedFormat($xlTypePDF, $pdfPath)
        return $pdfPath
    }
    finally {
        Release-ComObject $emailWs
    }
}

function Get-GssExcelErrorCellCount {
    param([object]$Worksheet, [int]$CellType)

    $usedRange = $null
    $errorCells = $null
    try {
        $usedRange = $Worksheet.UsedRange
        try {
            $errorCells = $usedRange.SpecialCells($CellType, $xlErrors)
            return [int]$errorCells.Count
        }
        catch {
            # Excel reports "no cells found" with a generic COM error, which is
            # also used for real automation failures. Independently count errors
            # instead of treating every inspection failure as a clean workbook.
            $address = [string]$usedRange.Address($false, $false)
            $formulaPredicate = if ($CellType -eq $xlCellTypeFormulas) {
                "--ISFORMULA($address)"
            }
            elseif ($CellType -eq $xlCellTypeConstants) {
                "--(ISFORMULA($address)=FALSE)"
            }
            else { throw "Unsupported Excel error inspection cell type: $CellType" }
            $count = $Worksheet.Evaluate("SUMPRODUCT(--ISERROR($address),$formulaPredicate)")
            if ($null -eq $count -or $count -is [System.Runtime.InteropServices.ErrorWrapper] -or
                [double]::IsNaN([double]$count) -or [double]::IsInfinity([double]$count) -or
                [double]$count -lt 0 -or [double]$count -ne [math]::Truncate([double]$count)) {
                throw 'Excel could not independently verify the workbook error count.'
            }
            return [int]$count
        }
    }
    finally {
        Release-ComObject $errorCells
        Release-ComObject $usedRange
    }
}

function Test-GssPersistedWorkbook {
    param(
        [Parameter(Mandatory)][string]$WorkbookPath,
        [Parameter(Mandatory)][datetime]$CurrentWeek,
        [Parameter(Mandatory)][datetime]$PriorYearWeek,
        [string]$CurrentSourceFileName,
        [string]$PriorYearSourceFileName,
        [object[]]$ExpectedSourceEvidence = @()
    )

    $ExpectedSourceEvidence = @($ExpectedSourceEvidence | Where-Object { $null -ne $_ })
    $excel = $null
    $workbook = $null
    $rawWs = $null
    $qaWs = $null
    $qaStatusCell = $null
    $qaWeekCell = $null
    $formulaErrorCount = 0
    $constantErrorCount = 0
    $externalLinks = @()
    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $excel.EnableEvents = $false
        $excel.AskToUpdateLinks = $false
        $workbook = $excel.Workbooks.Open($WorkbookPath, 0, $true)
        $excel.CalculateFullRebuild()

        if (-not [bool]$workbook.ProtectStructure) {
            throw 'Workbook structure is not protected after persistence.'
        }
        for ($index = 1; $index -le $workbook.Worksheets.Count; $index++) {
            $worksheet = $null
            try {
                $worksheet = $workbook.Worksheets.Item($index)
                if (-not [bool]$worksheet.ProtectContents) {
                    throw "Worksheet $($worksheet.Name) is not protected after persistence."
                }
                $formulaErrorCount += Get-GssExcelErrorCellCount $worksheet $xlCellTypeFormulas
                $constantErrorCount += Get-GssExcelErrorCellCount $worksheet $xlCellTypeConstants
            }
            finally {
                Release-ComObject $worksheet
            }
        }
        if (-not (Test-GssWorkbookGuardrailsApplied $workbook)) {
            throw 'Persisted workbook guardrails failed, including the required unlocked Exec_Dashboard selectors.'
        }
        if ($formulaErrorCount -gt 0 -or $constantErrorCount -gt 0) {
            throw "Persisted workbook contains formula or constant errors (formula=$formulaErrorCount; constant=$constantErrorCount)."
        }

        try {
            $linkResult = $workbook.LinkSources($xlLinkTypeExcelLinks)
            if ($null -ne $linkResult) {
                $externalLinks = @($linkResult | ForEach-Object { [string]$_ } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
            }
        }
        catch {
            throw "Excel could not inspect external workbook links: $($_.Exception.Message)"
        }
        if ($externalLinks.Count -gt 0) {
            throw "Persisted workbook contains external formula link(s): $($externalLinks -join ', ')"
        }

        $qaWs = $workbook.Worksheets.Item('QA Checks')
        $qaStatusCell = $qaWs.Range('B2')
        $qaWeekCell = $qaWs.Range('B3')
        if (([string]$qaStatusCell.Value2).Trim().ToUpperInvariant() -ne 'READY') {
            throw "Persisted workbook QA status is '$([string]$qaStatusCell.Value2)' instead of READY."
        }
        $persistedQaWeek = Convert-ExcelDate $qaWeekCell.Value2
        if ($null -eq $persistedQaWeek -or $persistedQaWeek.Date -ne $CurrentWeek.Date) {
            throw "Persisted workbook QA reporting date does not match $($CurrentWeek.ToString('yyyy-MM-dd'))."
        }

        $rawWs = $workbook.Worksheets.Item('Raw_Data')
        $headers = Get-HeaderMap $rawWs 1
        foreach ($header in @('SourceFile', 'Week', 'Restaurant', 'Ownership')) {
            if (-not $headers.ContainsKey((Normalize-Header $header))) {
                throw "Raw_Data is missing required persisted validation header '$header'."
            }
        }
        $sourceCol = $headers[(Normalize-Header 'SourceFile')]
        $weekCol = $headers[(Normalize-Header 'Week')]
        $lastRawRow = $rawWs.Cells.Item($rawWs.Rows.Count, 1).End($xlUp).Row
        $capacity = Get-GssRawDataCapacityPlan $lastRawRow 0 $gssRawDataLastFormulaRow
        if (-not $capacity.Fits) {
            throw "Persisted Raw_Data extends beyond modeled row $gssRawDataLastFormulaRow."
        }

        foreach ($expected in @(
            [pscustomobject]@{ Week = $CurrentWeek; SourceFile = $CurrentSourceFileName; Role = 'current' },
            [pscustomobject]@{ Week = $PriorYearWeek; SourceFile = $PriorYearSourceFileName; Role = 'prior-year' }
        )) {
            $keys = @(Get-GssWeekEntitySet $rawWs $expected.Week)
            $null = Assert-GssExactEntitySet -ActualKeys $keys -Context "Persisted $($expected.Role) week $($expected.Week.ToString('yyyy-MM-dd'))"
            $sourceNames = @()
            for ($row = 2; $row -le $lastRawRow; $row++) {
                $week = Convert-ExcelDate $rawWs.Cells.Item($row, $weekCol).Value2
                if ($week -and $week.Date -eq $expected.Week.Date) {
                    $sourceNames += (Get-CellString $rawWs $row $sourceCol)
                    $entityFormula = [string]$rawWs.Cells.Item($row, 18).Formula
                    $rowKeyFormula = [string]$rawWs.Cells.Item($row, 19).Formula
                    if ([string]::IsNullOrWhiteSpace($entityFormula) -or [string]::IsNullOrWhiteSpace($rowKeyFormula)) {
                        throw "Persisted $($expected.Role) week is missing entity/row-key formulas at Raw_Data row $row."
                    }
                }
            }
            # SourceFile remains useful provenance but renaming an input does not
            # invalidate an already imported, semantically identical report.
        }

        $rawRows = @(Read-GssRawDataRow $rawWs)
        $expectedContent = if ($ExpectedSourceEvidence.Count -gt 0) { $ExpectedSourceEvidence } else {
            @($CurrentWeek, $PriorYearWeek, $CurrentWeek.AddDays(-7)) |
                Where-Object { -not (Test-GssIntentionalHistoryGap $_) } |
                ForEach-Object { [pscustomobject]@{ WeekEnding = $_.ToString('yyyy-MM-dd'); PeriodStart = $_.AddDays(-90).ToString('yyyy-MM-dd'); SemanticSha256 = $null } }
        }
        foreach ($source in @($expectedContent)) {
            $week = [datetime]::ParseExact([string]$source.WeekEnding, 'yyyy-MM-dd', [Globalization.CultureInfo]::InvariantCulture)
            $periodRows = @($rawRows | Where-Object { $_.Week.Date -eq $week.Date })
            $null = Assert-GssExactEntitySet -ActualKeys @($periodRows | ForEach-Object { ConvertTo-GssSourceEntityKey $_.Ownership $_.Restaurant }) -Context "Persisted reporting week $($source.WeekEnding)"
            $semanticHash = Get-GssSourceSemanticHash $periodRows ([datetime]$source.PeriodStart) $week
            if ($source.SemanticSha256 -and $semanticHash -ne [string]$source.SemanticSha256) {
                throw "Persisted reporting values differ from the validated source for $($source.WeekEnding)."
            }
        }

        return [pscustomobject]@{
            Status = 'Ready'
            WorkbookPath = $WorkbookPath
            WorkbookSha256 = Get-GssSha256 $WorkbookPath
            WorkbookStructureProtected = $true
            WorksheetProtectionVerified = $true
            FormulaErrorCount = $formulaErrorCount
            ConstantErrorCount = $constantErrorCount
            ExternalFormulaLinkCount = $externalLinks.Count
            CurrentWeekEnding = $CurrentWeek.ToString('yyyy-MM-dd')
            PriorYearWeekEnding = $PriorYearWeek.ToString('yyyy-MM-dd')
            EntitySetVerified = $true
            SourcePeriodAndContentVerified = $true
            SourceFilenameIsProvenanceOnly = $true
            LastRawDataRow = $lastRawRow
            RemainingModeledRows = $capacity.AvailableRows
            WorstCaseRowsPerWeeklyLoad = $gssExpectedEntityKeys.Count
            WorstCaseWeeklyLoadsRemaining = [math]::Floor($capacity.AvailableRows / $gssExpectedEntityKeys.Count)
            CapacityRedesignReviewThreshold = 104
        }
    }
    finally {
        Release-ComObject $qaWeekCell
        Release-ComObject $qaStatusCell
        Release-ComObject $qaWs
        Release-ComObject $rawWs
        if ($workbook) { $workbook.Close($false) }
        Release-ComObject $workbook
        if ($excel) {
            $excel.Quit()
            Release-ComObject $excel
        }
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

function Resolve-GssTransactionLogPath {
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$AutomationDirectory
    )

    $resolved = (Resolve-Path -LiteralPath $Path).Path
    $allowedRoot = [System.IO.Path]::GetFullPath($AutomationDirectory).TrimEnd('\') + '\'
    if (-not $resolved.StartsWith($allowedRoot, [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "Transaction log is outside the GSS automation directory: $resolved"
    }
    return $resolved
}

function Repair-GssInterruptedTransaction {
    param(
        [Parameter(Mandatory)][object]$Run,
        [Parameter(Mandatory)][string]$FolderPath
    )

    $schema = if ($Run.ReceiptSchemaVersion) { [int]$Run.ReceiptSchemaVersion } else { 1 }
    if ($schema -notin @(1, 2)) { throw "Unsupported local transaction receipt schema: $schema" }
    $workbookPath = Resolve-GssDropboxPath -Path ([string]$Run.TargetWorkbookRelativePath) -FolderPath $FolderPath -RequireFile
    $backupPath = Resolve-GssDropboxPath -Path ([string]$Run.BackupWorkbookRelativePath) -FolderPath $FolderPath -RequireFile
    $startingHash = ([string]$Run.StartingWorkbookSha256).ToLowerInvariant()
    $promotedHash = ([string]$Run.PromotedWorkbookSha256).ToLowerInvariant()
    if ($startingHash -notmatch '^[a-f0-9]{64}$' -or $promotedHash -notmatch '^[a-f0-9]{64}$') {
        throw 'The transaction receipt does not contain valid old and new workbook hashes.'
    }
    if ((Get-GssSha256 $backupPath) -ne $startingHash) { throw 'The previous workbook backup failed verification.' }
    $currentHash = Get-GssSha256 $workbookPath
    if ($currentHash -ne $startingHash -and $currentHash -ne $promotedHash) {
        throw 'Recovery stopped because the live workbook contains a newer or unrelated change.'
    }
    $restoreWorkbook = $currentHash -eq $promotedHash -and $currentHash -ne $startingHash

    $pdfPath = $null
    $pdfBackupPath = $null
    $pdfAction = 'Unchanged'
    $currentPdfHash = $null
    $startingPdfHash = [string]$Run.StartingPdfSha256
    if ($Run.EmailComparisonPdfRelativePath) {
        $pdfPath = Resolve-GssDropboxPath -Path ([string]$Run.EmailComparisonPdfRelativePath) -FolderPath $FolderPath
        if ($Run.PdfBackupRelativePath) {
            $pdfBackupPath = Resolve-GssDropboxPath -Path ([string]$Run.PdfBackupRelativePath) -FolderPath $FolderPath -RequireFile
            $backupHash = Get-GssSha256 $pdfBackupPath
            $recordedBackupHash = [string]$Run.PdfBackupSha256
            if ([string]::IsNullOrWhiteSpace($recordedBackupHash)) {
                $recordedBackupHash = [string](@($Run.TransactionArtifacts | Where-Object Role -eq 'run_owned_pre_apply_pdf_backup' | Select-Object -First 1).Sha256)
            }
            if ($recordedBackupHash -and $backupHash -ne $recordedBackupHash) { throw 'The previous PDF backup failed verification.' }
            if ($startingPdfHash -and $backupHash -ne $startingPdfHash) { throw 'The PDF backup does not match the previous report.' }
            $startingPdfHash = $backupHash
        }
        if (Test-Path -LiteralPath $pdfPath -PathType Leaf) {
            $currentPdfHash = Get-GssSha256 $pdfPath
            if ($startingPdfHash -and $currentPdfHash -eq $startingPdfHash) {
                $pdfAction = 'Unchanged'
            }
            elseif ($currentPdfHash -eq [string]$Run.PromotedPdfSha256) {
                if ([bool]$Run.PdfExistedBefore) {
                    if (-not $pdfBackupPath) { throw 'Recovery needs the previous PDF backup, but it is missing.' }
                    $pdfAction = 'RestoreBackup'
                }
                else { $pdfAction = 'RemoveRunOwnedFile' }
            }
            else { throw 'Recovery stopped because the report PDF contains a newer or unrelated change.' }
        }
        elseif ([bool]$Run.PdfExistedBefore) {
            throw 'Recovery stopped because the previous report PDF is missing. Restore it explicitly from the verified backup.'
        }
    }

    # Check both targets and both backups before changing either. A retry can
    # finish a partially completed rollback without overwriting unrelated work.
    $candidates = @()
    $changed = $false
    $workbookRollback = 'Unchanged'
    $pdfRollback = 'Unchanged'
    try {
        if ($restoreWorkbook) {
            $candidate = Join-Path (Split-Path -Parent $workbookPath) ('.rollback-{0}-{1}' -f [string]$Run.RunId, [IO.Path]::GetFileName($workbookPath))
            Copy-Item -LiteralPath $backupPath -Destination $candidate -Force
            $candidates += $candidate
            if ((Get-GssSha256 $candidate) -ne $startingHash) { throw 'Workbook restore copy verification failed.' }
            if ((Get-GssSha256 $workbookPath) -ne $currentHash) { throw 'The live workbook changed during recovery.' }
            if ($pdfPath -and $currentPdfHash -and (Get-GssSha256 $pdfPath) -ne $currentPdfHash) { throw 'The PDF changed during recovery.' }
            [IO.File]::Replace($candidate, $workbookPath, [System.Management.Automation.Language.NullString]::Value, $true)
            $changed = $true
            $workbookRollback = 'RestoredBackup'
        }
        if ($pdfAction -eq 'RestoreBackup') {
            $candidate = Join-Path (Split-Path -Parent $pdfPath) ('.rollback-{0}-{1}' -f [string]$Run.RunId, [IO.Path]::GetFileName($pdfPath))
            Copy-Item -LiteralPath $pdfBackupPath -Destination $candidate -Force
            $candidates += $candidate
            if ((Get-GssSha256 $candidate) -ne $startingPdfHash) { throw 'PDF restore copy verification failed.' }
            if ((Get-GssSha256 $pdfPath) -ne $currentPdfHash) { throw 'The PDF changed during recovery.' }
            [IO.File]::Replace($candidate, $pdfPath, [System.Management.Automation.Language.NullString]::Value, $true)
            $changed = $true
            $pdfRollback = 'RestoredBackup'
        }
        elseif ($pdfAction -eq 'RemoveRunOwnedFile') {
            if ((Get-GssSha256 $pdfPath) -ne $currentPdfHash) { throw 'The PDF changed during recovery.' }
            Remove-Item -LiteralPath $pdfPath -Force
            $changed = $true
            $pdfRollback = 'RemovedRunOwnedFile'
        }
        if ((Get-GssSha256 $workbookPath) -ne $startingHash) { throw 'Workbook recovery verification failed.' }
        if ($startingPdfHash -and $pdfPath -and (Get-GssSha256 $pdfPath) -ne $startingPdfHash) { throw 'PDF recovery verification failed.' }
        return [pscustomobject]@{
            Status = if ($changed) { 'RolledBack' } else { 'Unchanged' }
            RunId = [string]$Run.RunId
            WorkbookPath = $workbookPath
            RestoredWorkbookSha256 = $startingHash
            WorkbookRollback = $workbookRollback
            PdfRollback = $pdfRollback
            TimestampUtc = [datetime]::UtcNow.ToString('o')
        }
    }
    finally {
        foreach ($candidate in $candidates) {
            if (Test-Path -LiteralPath $candidate -PathType Leaf) { Remove-Item -LiteralPath $candidate -Force }
        }
    }
}

function Restore-GssPromotedTransaction {
    param(
        [Parameter(Mandatory)][object]$Run,
        [Parameter(Mandatory)][string]$FolderPath
    )

    return Repair-GssInterruptedTransaction -Run $Run -FolderPath $FolderPath
}

function Invoke-GssPreparedRunPromotion {
    param(
        [Parameter(Mandatory)][object]$PreparedRun,
        [Parameter(Mandatory)][string]$ExpectedFingerprint,
        [Parameter(Mandatory)][string]$FolderPath,
        [Parameter(Mandatory)][string]$MainWorkbookPath,
        [Parameter(Mandatory)][string]$BackupDirectory,
        [Parameter(Mandatory)][string]$LogDirectory,
        [string]$PreparedDriveManifestPath,
        [string]$PreparedDriveManifestSha256
    )

    if ([string]::IsNullOrWhiteSpace($ExpectedFingerprint)) {
        throw 'Live promotion requires the exact reviewed copy-test fingerprint.'
    }
    if ([string]$PreparedRun.Mode -ne 'CopyTestOnly' -or [string]$PreparedRun.TransactionStatus -ne 'Prepared') {
        throw 'Live promotion requires a Prepared copy-test run.'
    }
    if ([string]$PreparedRun.HostName -cne [Environment]::MachineName) {
        throw "Prepared run belongs to host '$($PreparedRun.HostName)', not this workstation '$([Environment]::MachineName)'."
    }

    $resolvedDriveManifestPath = $null
    $actualDriveManifestSha256 = $null
    if ($PreparedDriveManifestPath -or $PreparedDriveManifestSha256) {
        if (-not $PreparedDriveManifestPath -or -not $PreparedDriveManifestSha256) {
            throw 'An optional Drive manifest must include both its path and SHA-256.'
        }
        $resolvedDriveManifestPath = (Resolve-Path -LiteralPath $PreparedDriveManifestPath).Path
        $actualDriveManifestSha256 = Get-GssSha256 $resolvedDriveManifestPath
        if ($actualDriveManifestSha256 -ne $PreparedDriveManifestSha256.Trim().ToLowerInvariant()) {
            throw 'Drive prepared manifest hash changed before live promotion.'
        }
        $driveManifest = Get-Content -LiteralPath $resolvedDriveManifestPath -Raw | ConvertFrom-Json
        if ([string]$driveManifest.status -ne 'Prepared' -or
            [string]$driveManifest.run_id -ne [string]$PreparedRun.RunId -or
            [string]$driveManifest.fingerprint -ne [string]$PreparedRun.RunFingerprint) {
            throw 'Drive prepared manifest is not bound to this transaction.'
        }
    }

    $calculatedFingerprint = Get-GssFingerprintFromPreparedRun $PreparedRun
    if ($calculatedFingerprint -ne ([string]$PreparedRun.RunFingerprint).ToLowerInvariant() -or
        $calculatedFingerprint -ne $ExpectedFingerprint.Trim().ToLowerInvariant()) {
        throw 'Prepared run fingerprint verification failed. Run a fresh copy-test.'
    }

    $stagedWorkbookPath = Resolve-GssDropboxPath -Path ([string]$PreparedRun.TargetWorkbookRelativePath) -FolderPath $FolderPath -RequireFile
    $stagedPdfPath = Resolve-GssDropboxPath -Path ([string]$PreparedRun.EmailComparisonPdfRelativePath) -FolderPath $FolderPath -RequireFile
    $currentSourcePath = Resolve-GssDropboxPath -Path ([string]$PreparedRun.CurrentSourceRelativePath) -FolderPath $FolderPath -RequireFile
    $priorYearSourcePath = Resolve-GssDropboxPath -Path ([string]$PreparedRun.PriorYearSourceRelativePath) -FolderPath $FolderPath -RequireFile

    $preflightEvidence = [ordered]@{
        StartingWorkbookSha256 = Get-GssSha256 $MainWorkbookPath
        CurrentSourceSha256 = Get-GssSha256 $currentSourcePath
        PriorYearSourceSha256 = Get-GssSha256 $priorYearSourcePath
        StagedWorkbookSha256 = Get-GssSha256 $stagedWorkbookPath
        StagedPdfSha256 = Get-GssSha256 $stagedPdfPath
    }
    foreach ($check in @(
        [pscustomobject]@{ Name = 'starting live workbook'; Actual = $preflightEvidence.StartingWorkbookSha256; Expected = [string]$PreparedRun.StartingWorkbookSha256 },
        [pscustomobject]@{ Name = 'current source'; Actual = $preflightEvidence.CurrentSourceSha256; Expected = [string]$PreparedRun.CurrentSourceSha256 },
        [pscustomobject]@{ Name = 'prior-year source'; Actual = $preflightEvidence.PriorYearSourceSha256; Expected = [string]$PreparedRun.PriorYearSourceSha256 },
        [pscustomobject]@{ Name = 'staged workbook'; Actual = $preflightEvidence.StagedWorkbookSha256; Expected = [string]$PreparedRun.StagedWorkbookSha256 },
        [pscustomobject]@{ Name = 'staged PDF'; Actual = $preflightEvidence.StagedPdfSha256; Expected = [string]$PreparedRun.StagedPdfSha256 }
    )) {
        if ($check.Actual -ne $check.Expected.ToLowerInvariant()) {
            throw "Prepared transaction is stale because the $($check.Name) hash changed. Run a fresh copy-test."
        }
    }
    Assert-GssSourceEvidence -Run $PreparedRun -FolderPath $FolderPath

    $currentWeek = [datetime]::ParseExact([string]$PreparedRun.CurrentWeekEnding, 'yyyy-MM-dd', [Globalization.CultureInfo]::InvariantCulture)
    $priorYearWeek = [datetime]::ParseExact([string]$PreparedRun.PriorYearWeekEnding, 'yyyy-MM-dd', [Globalization.CultureInfo]::InvariantCulture)
    $currentSourceFileNameInRawData = if ($PreparedRun.CurrentSourceFileNameInRawData) {
        [string]$PreparedRun.CurrentSourceFileNameInRawData
    }
    else {
        [System.IO.Path]::GetFileName([string]$PreparedRun.CurrentSourceWorkbook)
    }
    $priorYearSourceFileNameInRawData = if ($PreparedRun.PriorYearSourceFileNameInRawData) {
        [string]$PreparedRun.PriorYearSourceFileNameInRawData
    }
    else {
        [System.IO.Path]::GetFileName([string]$PreparedRun.PriorYearSourceWorkbook)
    }
    $stageValidation = Test-GssPersistedWorkbook `
        -WorkbookPath $stagedWorkbookPath `
        -CurrentWeek $currentWeek `
        -PriorYearWeek $priorYearWeek `
        -CurrentSourceFileName $currentSourceFileNameInRawData `
        -PriorYearSourceFileName $priorYearSourceFileNameInRawData `
        -ExpectedSourceEvidence @($PreparedRun.SourceEvidence)

    $runId = [string]$PreparedRun.RunId
    $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $workbookDirectory = Split-Path -Parent $MainWorkbookPath
    $workbookCandidate = Join-Path $workbookDirectory ('.partial-{0}-{1}' -f $runId, [System.IO.Path]::GetFileName($MainWorkbookPath))
    $pdfDirectory = Join-Path $FolderPath '04 Email Comparison PDFs'
    New-Item -ItemType Directory -Path $pdfDirectory, $BackupDirectory, $LogDirectory -Force | Out-Null
    $productionPdfPath = Join-Path $pdfDirectory ([System.IO.Path]::GetFileName($stagedPdfPath))
    $pdfCandidate = Join-Path $pdfDirectory ('.partial-{0}-{1}' -f $runId, [System.IO.Path]::GetFileName($productionPdfPath))
    $runFileKey = (Get-GssTextSha256 $runId).Substring(0, 12)
    $backupPath = Join-Path $BackupDirectory ("main_{0}_{1}.xlsx" -f $timestamp, $runFileKey)
    $pdfExistedBefore = Test-Path -LiteralPath $productionPdfPath -PathType Leaf
    $startingPdfSha256 = if ($pdfExistedBefore) { Get-GssSha256 $productionPdfPath } else { $null }
    $pdfBackupPath = if ($pdfExistedBefore) {
        Join-Path $BackupDirectory ("report_{0}_{1}.pdf" -f $timestamp, $runFileKey)
    }
    else { $null }
    $workbookPromoted = $false
    $pdfPromoted = $false
    $promotionLogPath = Join-Path $LogDirectory "gss_update_$timestamp`_$runId`_apply.json"
    $journal = $null

    try {
        Copy-Item -LiteralPath $stagedWorkbookPath -Destination $workbookCandidate -Force
        Copy-Item -LiteralPath $stagedPdfPath -Destination $pdfCandidate -Force
        if ((Get-GssSha256 $workbookCandidate) -ne $preflightEvidence.StagedWorkbookSha256 -or
            (Get-GssSha256 $pdfCandidate) -ne $preflightEvidence.StagedPdfSha256) {
            throw 'Promotion candidate copy verification failed.'
        }

        [IO.File]::Copy($MainWorkbookPath, $backupPath, $false)
        if ((Get-GssSha256 $backupPath) -ne $preflightEvidence.StartingWorkbookSha256) {
            throw 'The previous workbook backup failed verification. No live change was made.'
        }
        if ($pdfExistedBefore) {
            [IO.File]::Copy($productionPdfPath, $pdfBackupPath, $false)
            if ((Get-GssSha256 $pdfBackupPath) -ne $startingPdfSha256) {
                throw 'The previous PDF backup failed verification. No live change was made.'
            }
        }
        $journal = $PreparedRun | Select-Object *
        $journalFields = @{
            ReceiptSchemaVersion = 2
            TimestampUtc = [datetime]::UtcNow.ToString('o')
            Mode = 'ApplyToMainWorkbook'
            Status = 'PreparedForPromotion'
            TransactionStatus = 'Promoting'
            PreparedRunLogPath = [string]$PreparedRun.LogPath
            TargetWorkbook = $MainWorkbookPath
            TargetWorkbookRelativePath = ConvertTo-GssDropboxRelativePath $MainWorkbookPath $FolderPath
            PromotedWorkbookSha256 = $preflightEvidence.StagedWorkbookSha256
            BackupWorkbook = $backupPath
            BackupWorkbookRelativePath = ConvertTo-GssDropboxRelativePath $backupPath $FolderPath
            StagedWorkbookRelativePath = ConvertTo-GssDropboxRelativePath $stagedWorkbookPath $FolderPath
            StagedPdfRelativePath = ConvertTo-GssDropboxRelativePath $stagedPdfPath $FolderPath
            EmailComparisonPdf = $productionPdfPath
            EmailComparisonPdfRelativePath = ConvertTo-GssDropboxRelativePath $productionPdfPath $FolderPath
            PromotedPdfSha256 = $preflightEvidence.StagedPdfSha256
            PdfExistedBefore = $pdfExistedBefore
            StartingPdfSha256 = $startingPdfSha256
            PdfBackupSha256 = $startingPdfSha256
            PdfBackup = $pdfBackupPath
            PdfBackupRelativePath = if ($pdfBackupPath) { ConvertTo-GssDropboxRelativePath $pdfBackupPath $FolderPath } else { $null }
            LocalBackupStatus = 'Verified'
            LogPath = $promotionLogPath
        }
        foreach ($name in $journalFields.Keys) { $journal | Add-Member -NotePropertyName $name -NotePropertyValue $journalFields[$name] -Force }
        # A process interruption after either replacement leaves enough evidence
        # to restore the previous report safely on this or another computer.
        Write-GssAtomicJson -Path $promotionLogPath -InputObject $journal

        # This is the final optimistic-concurrency check immediately before the
        # only mutation of the live workbook.
        if ((Get-GssSha256 $MainWorkbookPath) -ne $preflightEvidence.StartingWorkbookSha256 -or
            (Get-GssSha256 $currentSourcePath) -ne $preflightEvidence.CurrentSourceSha256 -or
            (Get-GssSha256 $priorYearSourcePath) -ne $preflightEvidence.PriorYearSourceSha256) {
            throw 'Live workbook or source evidence changed immediately before promotion. No live change was made.'
        }
        Assert-GssSourceEvidence -Run $PreparedRun -FolderPath $FolderPath
        if ($pdfExistedBefore) {
            if ((Get-GssSha256 $productionPdfPath) -ne $startingPdfSha256) { throw 'The previous PDF changed before promotion. No live change was made.' }
        }
        elseif (Test-Path -LiteralPath $productionPdfPath -PathType Leaf) {
            throw 'Another report PDF appeared before promotion. No live change was made.'
        }

        if ([System.IO.Path]::GetPathRoot($workbookCandidate) -cne [System.IO.Path]::GetPathRoot($MainWorkbookPath)) {
            throw 'Atomic workbook replacement requires a same-volume staging candidate.'
        }
        [System.IO.File]::Replace($workbookCandidate, $MainWorkbookPath, [System.Management.Automation.Language.NullString]::Value, $true)
        $workbookPromoted = $true
        $promotedWorkbookHash = Get-GssSha256 $MainWorkbookPath
        if ($promotedWorkbookHash -ne $preflightEvidence.StagedWorkbookSha256) {
            throw 'Promoted workbook hash does not equal the reviewed staged workbook hash.'
        }
        if ((Get-GssSha256 $backupPath) -ne $preflightEvidence.StartingWorkbookSha256) {
            throw 'Run-owned atomic workbook backup does not equal the starting workbook hash.'
        }

        $journal.Status = 'WorkbookPromoted'
        $journal.TransactionStatus = 'WorkbookPromoted'
        Write-GssAtomicJson -Path $promotionLogPath -InputObject $journal
        $liveValidation = $stageValidation | Select-Object *
        $liveValidation.WorkbookPath = $MainWorkbookPath
        $liveValidation | Add-Member -NotePropertyName ValidationMethod -NotePropertyValue 'MatchesValidatedStagedWorkbook' -Force

        if ($pdfExistedBefore) {
            if ((Get-GssSha256 $productionPdfPath) -ne $startingPdfSha256) { throw 'The previous PDF changed before its replacement.' }
            [System.IO.File]::Replace($pdfCandidate, $productionPdfPath, [System.Management.Automation.Language.NullString]::Value, $true)
        }
        else {
            [System.IO.File]::Move($pdfCandidate, $productionPdfPath)
        }
        $pdfPromoted = $true
        $promotedPdfHash = Get-GssSha256 $productionPdfPath
        if ($promotedPdfHash -ne $preflightEvidence.StagedPdfSha256) {
            throw 'Promoted PDF hash does not equal the reviewed staged PDF hash.'
        }
        $journal.Status = 'PdfPromoted'
        $journal.TransactionStatus = 'PdfPromoted'
        Write-GssAtomicJson -Path $promotionLogPath -InputObject $journal

        $fileEvidence = @(
            [pscustomobject]@{
                Role = 'rolling_workbook'
                RelativePath = [string]$PreparedRun.CurrentSourceRelativePath
                ByteSize = [long](Get-Item -LiteralPath $currentSourcePath).Length
                Sha256 = $preflightEvidence.CurrentSourceSha256
            },
            [pscustomobject]@{
                Role = 'prior_year_rolling_workbook'
                RelativePath = [string]$PreparedRun.PriorYearSourceRelativePath
                ByteSize = [long](Get-Item -LiteralPath $priorYearSourcePath).Length
                Sha256 = $preflightEvidence.PriorYearSourceSha256
            },
            [pscustomobject]@{
                Role = 'live_workbook'
                RelativePath = ConvertTo-GssDropboxRelativePath -Path $MainWorkbookPath -FolderPath $FolderPath
                ByteSize = [long](Get-Item -LiteralPath $MainWorkbookPath).Length
                Sha256 = $promotedWorkbookHash
            },
            [pscustomobject]@{
                Role = 'comparison_pdf'
                RelativePath = ConvertTo-GssDropboxRelativePath -Path $productionPdfPath -FolderPath $FolderPath
                ByteSize = [long](Get-Item -LiteralPath $productionPdfPath).Length
                Sha256 = $promotedPdfHash
            }
        )

        $summary = [pscustomobject]@{
            ReceiptSchemaVersion = 2
            Timestamp = (Get-Date).ToString('s')
            TimestampUtc = [datetime]::UtcNow.ToString('o')
            RunId = $runId
            HostName = [Environment]::MachineName
            ProgramRelease = if ($PreparedRun.ProgramRelease) { [string]$PreparedRun.ProgramRelease } else { 'unreleased' }
            Mode = 'ApplyToMainWorkbook'
            Status = 'Updated'
            TransactionStatus = 'Committed'
            WorkbookStatus = [string]$PreparedRun.WorkbookStatus
            GuardrailsApplied = [bool]$PreparedRun.GuardrailsApplied
            RunFingerprint = $calculatedFingerprint
            DrivePreparedManifestPath = $resolvedDriveManifestPath
            DrivePreparedManifestSha256 = $actualDriveManifestSha256
            PreparedRunLogPath = [string]$PreparedRun.LogPath
            Folder = $FolderPath
            CurrentWeekEnding = [string]$PreparedRun.CurrentWeekEnding
            PriorYearWeekEnding = [string]$PreparedRun.PriorYearWeekEnding
            ReportingPeriodLabel = [string]$PreparedRun.ReportingPeriodLabel
            RollingPeriodStart = [string]$PreparedRun.RollingPeriodStart
            RollingPeriodEnd = [string]$PreparedRun.RollingPeriodEnd
            PreviousRollingWindowEnding = [string]$PreparedRun.PreviousRollingWindowEnding
            ComparisonLabel = [string]$PreparedRun.ComparisonLabel
            CurrentSourceWorkbook = $currentSourcePath
            CurrentSourceRelativePath = [string]$PreparedRun.CurrentSourceRelativePath
            CurrentSourceFileNameInRawData = $currentSourceFileNameInRawData
            PriorYearSourceWorkbook = $priorYearSourcePath
            PriorYearSourceRelativePath = [string]$PreparedRun.PriorYearSourceRelativePath
            PriorYearSourceFileNameInRawData = $priorYearSourceFileNameInRawData
            StartingWorkbookSha256 = $preflightEvidence.StartingWorkbookSha256
            CurrentSourceSha256 = $preflightEvidence.CurrentSourceSha256
            PriorYearSourceSha256 = $preflightEvidence.PriorYearSourceSha256
            CurrentSourceSemanticSha256 = [string]$PreparedRun.CurrentSourceSemanticSha256
            PriorYearSourceSemanticSha256 = [string]$PreparedRun.PriorYearSourceSemanticSha256
            SourceSetSha256 = [string]$PreparedRun.SourceSetSha256
            SourceEvidence = @($PreparedRun.SourceEvidence)
            SourceChoiceSetSha256 = [string]$PreparedRun.SourceChoiceSetSha256
            AcceptedSourceChoices = @($PreparedRun.AcceptedSourceChoices)
            StagedWorkbook = $stagedWorkbookPath
            StagedWorkbookRelativePath = ConvertTo-GssDropboxRelativePath $stagedWorkbookPath $FolderPath
            StagedWorkbookSha256 = $preflightEvidence.StagedWorkbookSha256
            StagedPdf = $stagedPdfPath
            StagedPdfRelativePath = ConvertTo-GssDropboxRelativePath $stagedPdfPath $FolderPath
            StagedPdfSha256 = $preflightEvidence.StagedPdfSha256
            TargetWorkbook = $MainWorkbookPath
            TargetWorkbookRelativePath = ConvertTo-GssDropboxRelativePath -Path $MainWorkbookPath -FolderPath $FolderPath
            PromotedWorkbookSha256 = $promotedWorkbookHash
            BackupWorkbook = $backupPath
            BackupWorkbookRelativePath = ConvertTo-GssDropboxRelativePath -Path $backupPath -FolderPath $FolderPath
            EmailComparisonPdf = $productionPdfPath
            EmailComparisonPdfRelativePath = ConvertTo-GssDropboxRelativePath -Path $productionPdfPath -FolderPath $FolderPath
            PromotedPdfSha256 = $promotedPdfHash
            PdfExistedBefore = $pdfExistedBefore
            StartingPdfSha256 = $startingPdfSha256
            PdfBackupSha256 = $startingPdfSha256
            PdfBackup = $pdfBackupPath
            PdfBackupRelativePath = if ($pdfBackupPath) { ConvertTo-GssDropboxRelativePath -Path $pdfBackupPath -FolderPath $FolderPath } else { $null }
            TransactionArtifacts = @(
                [pscustomobject]@{
                    SourcePath = $stagedWorkbookPath
                    Role = 'reviewed_staged_workbook'
                    Classification = 'restricted_operational'
                    Sha256 = $preflightEvidence.StagedWorkbookSha256
                },
                [pscustomobject]@{
                    SourcePath = $stagedPdfPath
                    Role = 'reviewed_staged_comparison_pdf'
                    Classification = 'restricted_operational'
                    Sha256 = $preflightEvidence.StagedPdfSha256
                },
                [pscustomobject]@{
                    SourcePath = $backupPath
                    Role = 'run_owned_pre_apply_workbook_backup'
                    Classification = 'restricted_operational'
                    Sha256 = $preflightEvidence.StartingWorkbookSha256
                }
            ) + @(
                if ($pdfBackupPath) {
                    [pscustomobject]@{
                        SourcePath = $pdfBackupPath
                        Role = 'run_owned_pre_apply_pdf_backup'
                        Classification = 'restricted_operational'
                        Sha256 = Get-GssSha256 $pdfBackupPath
                    }
                }
            )
            ReleaseArchives = @($PreparedRun.ReleaseArchives)
            FileEvidence = $fileEvidence
            RowsAppended = [int]$PreparedRun.RowsAppended
            RowsCorrected = [int]$PreparedRun.RowsCorrected
            BusinessDataChanged = [bool]$PreparedRun.BusinessDataChanged
            NoBusinessChanges = [bool]$PreparedRun.NoBusinessChanges
            WeeksAppended = @($PreparedRun.WeeksAppended)
            WeeksCorrected = @($PreparedRun.WeeksCorrected)
            WeeksSkipped = @($PreparedRun.WeeksSkipped)
            PreparedPostSaveValidation = $stageValidation
            LivePostSaveValidation = $liveValidation
            WorstCaseWeeklyLoadsRemaining = [int]$liveValidation.WorstCaseWeeklyLoadsRemaining
            LocalBackupStatus = 'Verified'
            RollbackPolicy = 'OnlyIfLiveHashEqualsRunOwnedPromotedHash'
        }
        $summary | Add-Member -NotePropertyName LogPath -NotePropertyValue $promotionLogPath
        Write-GssAtomicJson -Path $promotionLogPath -InputObject $summary
        return $summary
    }
    catch {
        $failureMessage = $_.Exception.Message
        $rollback = $null
        if ($workbookPromoted -or $pdfPromoted) {
            try { $rollback = Repair-GssInterruptedTransaction -Run $journal -FolderPath $FolderPath }
            catch {
                $rollback = [pscustomobject]@{
                    Status = 'RollbackBlocked'
                    Error = $_.Exception.Message
                    TimestampUtc = [datetime]::UtcNow.ToString('o')
                }
            }
        }
        $restored = $rollback -and $rollback.Status -in @('RolledBack', 'Unchanged')
        $failureReceipt = if ($journal) { $journal | Select-Object * } else { $PreparedRun | Select-Object * }
        $failureFields = @{
            ReceiptSchemaVersion = 2
            TimestampUtc = [datetime]::UtcNow.ToString('o')
            RunId = $runId
            HostName = [Environment]::MachineName
            Mode = 'ApplyToMainWorkbook'
            Status = if ($restored) { 'RolledBack' } elseif ($workbookPromoted -or $pdfPromoted) { 'RollbackBlocked' } else { 'Blocked' }
            TransactionStatus = if ($restored) { 'Aborted' } elseif ($workbookPromoted -or $pdfPromoted) { 'Conflict' } else { 'Blocked' }
            Error = $failureMessage
            RunFingerprint = $calculatedFingerprint
            WorkbookWasPromoted = $workbookPromoted
            PdfWasPromoted = $pdfPromoted
            Rollback = $rollback
            LogPath = $promotionLogPath
        }
        foreach ($name in $failureFields.Keys) { $failureReceipt | Add-Member -NotePropertyName $name -NotePropertyValue $failureFields[$name] -Force }
        Write-GssAtomicJson -Path $promotionLogPath -InputObject $failureReceipt
        throw "$failureMessage Transaction receipt: $promotionLogPath"
    }
    finally {
        foreach ($candidate in @($workbookCandidate, $pdfCandidate)) {
            if ($candidate -and (Test-Path -LiteralPath $candidate -PathType Leaf)) {
                Remove-Item -LiteralPath $candidate -Force
            }
        }
    }
}

function Write-RunSummary {
    param([object]$Summary)

    Write-Host ''
    Write-Host 'GSS workbook update summary'
    Write-Host ('  Mode: {0}' -f $Summary.Mode)
    Write-Host ('  Status: {0}' -f $Summary.Status)
    Write-Host ('  Workbook QA: {0}' -f $Summary.WorkbookStatus)
    Write-Host ('  Guardrails applied: {0}' -f $Summary.GuardrailsApplied)
    Write-Host ('  13-week rolling through: {0}' -f $Summary.CurrentWeekEnding)
    Write-Host ('  Prior-year rolling through: {0}' -f $Summary.PriorYearWeekEnding)
    Write-Host ('  Rows appended: {0}; rows corrected: {1}' -f $Summary.RowsAppended, [int]$Summary.RowsCorrected)
    Write-Host ('  Current source: {0}' -f $Summary.CurrentSourceWorkbook)
    Write-Host ('  Prior-year source: {0}' -f $Summary.PriorYearSourceWorkbook)
    Write-Host ('  Target workbook: {0}' -f $Summary.TargetWorkbook)
    if ($Summary.BackupWorkbook) {
        Write-Host ('  Backup workbook: {0}' -f $Summary.BackupWorkbook)
    }
    Write-Host ('  Email PDF: {0}' -f $Summary.EmailComparisonPdf)
    Write-Host ('  Log: {0}' -f $Summary.LogPath)
    Write-Host ''
}

function Invoke-GssWorkbookUpdate {
    param(
        [string]$FolderPath,
        [string]$WorkbookName = 'GSS Score Trends - Main.xlsx',
        [switch]$ApplyMode,
        [switch]$ReturnObject,
        [string]$RunIdentifier,
        [string]$PreparedLogPath,
        [string]$ReviewedFingerprint,
        [string]$RollbackLogPath,
        [switch]$CallerOwnsMutex,
        [string]$ReleaseIdentifier = 'unreleased',
        [string]$PreparedDriveManifestPath,
        [string]$PreparedDriveManifestSha256,
        [string]$ReportingWeek,
        [string[]]$CorrectionSourcePaths = @()
    )

    if ([string]::IsNullOrWhiteSpace($FolderPath)) {
        $scriptRoot = Split-Path -Parent $PSCommandPath
        $projectRoot = Split-Path -Parent $scriptRoot
        $FolderPath = Split-Path -Parent $projectRoot
    }

    $FolderPath = $FolderPath.Trim('"')
    $FolderPath = (Resolve-Path -LiteralPath $FolderPath).Path
    $mainPath = Resolve-MainWorkbookPath $FolderPath $WorkbookName

    $automationDir = Join-Path $FolderPath '_automation_runs'
    $logDir = Join-Path $automationDir 'logs'
    $backupDir = Join-Path $automationDir 'backups'
    $testDir = Join-Path $automationDir 'test-output'
    New-Item -ItemType Directory -Path $logDir, $backupDir, $testDir -Force | Out-Null

    $transactionId = if ([string]::IsNullOrWhiteSpace($RunIdentifier)) {
        [guid]::NewGuid().ToString('D')
    }
    else {
        try { ([guid]$RunIdentifier).ToString('D') }
        catch { throw "RunId must be a GUID: $RunIdentifier" }
    }
    $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss_fff'
    $transactionDir = Join-Path $testDir $transactionId
    $logPath = Join-Path $logDir "gss_update_$timestamp`_$transactionId.json"
    $excel = $null
    $sourceWb = $null
    $sourceWs = $null
    $targetWb = $null
    $rawWs = $null
    $targetPath = $null
    $backupPath = $null
    $emailComparisonPdfPath = $null
    $workbookStatus = $null
    $guardrailsApplied = $false
    $mutex = $null
    $ownsMutex = $false

    try {
        if (-not $CallerOwnsMutex) {
            $mutex = New-Object System.Threading.Mutex($false, 'Global\GSSSurveyWorkbookAutomationTransaction')
            try {
                $ownsMutex = $mutex.WaitOne(0)
            }
            catch [System.Threading.AbandonedMutexException] {
                $ownsMutex = $true
            }
            if (-not $ownsMutex) {
                throw 'Another GSS workbook transaction is already active on this workstation.'
            }
        }

        if (-not [string]::IsNullOrWhiteSpace($RollbackLogPath)) {
            if ($ApplyMode -or -not [string]::IsNullOrWhiteSpace($PreparedLogPath)) {
                throw 'Rollback cannot be combined with prepare or apply parameters.'
            }
            $resolvedRollbackLog = Resolve-GssTransactionLogPath -Path $RollbackLogPath -AutomationDirectory $automationDir
            $rollbackRun = Get-Content -LiteralPath $resolvedRollbackLog -Raw | ConvertFrom-Json
            $rollback = Restore-GssPromotedTransaction -Run $rollbackRun -FolderPath $FolderPath
            $rollbackReceiptPath = Join-Path $logDir "gss_rollback_$timestamp`_$($rollback.RunId).json"
            $rollback | Add-Member -NotePropertyName LogPath -NotePropertyValue $rollbackReceiptPath
            Write-GssAtomicJson -Path $rollbackReceiptPath -InputObject $rollback
            if ($ReturnObject) { return $rollback }
            Write-Information "GSS rollback completed: $rollbackReceiptPath" -InformationAction Continue
            return
        }

        if ($ApplyMode) {
            if ([string]::IsNullOrWhiteSpace($PreparedLogPath)) {
                throw 'Direct live mutation is disabled. Supply the reviewed PreparedRunLogPath and fingerprint from a copy-test.'
            }
            $resolvedPreparedLog = Resolve-GssTransactionLogPath -Path $PreparedLogPath -AutomationDirectory $automationDir
            $preparedRun = Get-Content -LiteralPath $resolvedPreparedLog -Raw | ConvertFrom-Json
            $summary = Invoke-GssPreparedRunPromotion `
                -PreparedRun $preparedRun `
                -ExpectedFingerprint $ReviewedFingerprint `
                -FolderPath $FolderPath `
                -MainWorkbookPath $mainPath `
                -BackupDirectory $backupDir `
                -LogDirectory $logDir `
                -PreparedDriveManifestPath $PreparedDriveManifestPath `
                -PreparedDriveManifestSha256 $PreparedDriveManifestSha256
            Write-RunSummary $summary
            if ($ReturnObject) { return $summary }
            return
        }

        if (-not [string]::IsNullOrWhiteSpace($PreparedLogPath) -or
            -not [string]::IsNullOrWhiteSpace($ReviewedFingerprint)) {
            throw 'PreparedRunLogPath and ExpectedFingerprint are valid only for -Apply.'
        }

        New-Item -ItemType Directory -Path $transactionDir -Force | Out-Null
        $targetPath = Join-Path $transactionDir $WorkbookName
        $startingWorkbookSha256 = Get-GssSha256 $mainPath

        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $excel.EnableEvents = $false
        $excel.AskToUpdateLinks = $false
        try {
            $excel.Calculation = $xlCalculationAutomatic
        }
        catch {
            Write-Warning "Excel did not accept the automatic calculation setting before opening a workbook. Continuing; the updater will still calculate before saving."
        }

        $allSources = @(Get-RollingSources $excel $FolderPath)
        $reportingDate = if ([string]::IsNullOrWhiteSpace($ReportingWeek)) {
            ($allSources | Sort-Object WeekEnding -Descending | Select-Object -First 1).WeekEnding
        }
        else {
            try { [datetime]::ParseExact($ReportingWeek, 'yyyy-MM-dd', [Globalization.CultureInfo]::InvariantCulture) }
            catch { throw 'ReportingWeekEnding must be an ISO date such as 2026-10-04.' }
        }
        $null = Get-GssReportingPeriod -PeriodStart $reportingDate.AddDays(-90) -PeriodEnd $reportingDate -Context 'Requested reporting week'
        $approvedSourcePaths = @($CorrectionSourcePaths | ForEach-Object {
            Resolve-GssDropboxPath -Path $_ -FolderPath $FolderPath -RequireFile
        } | Select-Object -Unique)
        $correctionSources = @()
        foreach ($path in $approvedSourcePaths) {
            $pathCandidates = @($allSources | Where-Object { $_.File.FullName -eq $path })
            if ($pathCandidates.Count -ne 1) { throw "Correction source is not a recognized rolling report: $path" }
            $correctionSources += $pathCandidates[0]
        }
        Copy-Item -LiteralPath $mainPath -Destination $targetPath -Force
        if ((Get-GssSha256 $targetPath) -ne $startingWorkbookSha256) {
            throw 'Staged workbook copy does not match the starting live workbook hash.'
        }

        $targetWb = $excel.Workbooks.Open($targetPath, 0, $false)
        if ($targetWb.ReadOnly) {
            throw "Staged workbook opened read-only and cannot be prepared: $targetPath"
        }

        $rawWs = $targetWb.Worksheets.Item('Raw_Data')

        $existingRows = @(Read-GssRawDataRow $rawWs)
        if (@($existingRows | Where-Object { $_.Week -gt $reportingDate }).Count -gt 0) {
            throw 'The main workbook already contains a later reporting week. Run that latest week and choose the older correction source instead.'
        }
        $completeExistingWeeks = @()
        foreach ($group in @($existingRows | Group-Object { $_.Week.ToString('yyyy-MM-dd') })) {
            $keys = @($group.Group | ForEach-Object { ConvertTo-GssSourceEntityKey $_.Ownership $_.Restaurant })
            $valid = (Get-GssEntitySetValidation $keys).IsValid
            if ($valid) {
                foreach ($row in $group.Group) {
                    try { $null = ConvertTo-GssMetricValue $row.Metrics "Existing Raw_Data row $($row.Row)" }
                    catch { $valid = $false; break }
                }
            }
            if ($valid) { $completeExistingWeeks += $group.Group[0].Week }
        }
        $requiredWeeks = @(Get-GssRequiredSourceWeek -ReportingWeek $reportingDate -ExistingWeeks $completeExistingWeeks -CorrectionWeeks @($correctionSources | ForEach-Object { $_.WeekEnding }))
        $committedChoices = @(Get-GssCommittedSourceChoice -LogDirectory $logDir)
        $acceptedSourceChoices = @()
        $requestedSources = @()
        foreach ($required in $requiredWeeks) {
            $acceptedPath = $null
            $candidates = @($allSources | Where-Object { $_.WeekEnding.Date -eq $required.WeekEnding.Date })
            if ($candidates.Count -eq 0) {
                throw "Missing rolling report for $($required.WeekEnding.ToString('yyyy-MM-dd')) ($($required.Role)). The previous good report was not changed."
            }
            foreach ($candidate in $candidates) { $null = @(Get-SourceRowsFromFile $excel $candidate) }
            $selectionPaths = $approvedSourcePaths
            $explicitWeekChoice = @($candidates | Where-Object { $_.File.FullName -in $approvedSourcePaths }).Count -gt 0
            if (-not $explicitWeekChoice) {
                $acceptedPath = Get-GssAcceptedSourceChoicePath -Sources $candidates -AcceptedChoices $committedChoices
                if ($acceptedPath) { $selectionPaths = @($selectionPaths) + $acceptedPath }
            }
            $selected = Select-GssSourceContent -Sources $candidates -CorrectionSourcePaths $selectionPaths
            if ($explicitWeekChoice -or $acceptedPath) {
                $choice = New-GssSourceChoiceEvidence -Sources $candidates -SelectedSourcePath $selected.File.FullName
                $earlierChoice = $committedChoices | Where-Object WeekEnding -eq $choice.WeekEnding | Select-Object -First 1
                if ($earlierChoice) { $choice.KnownSemanticSha256 = @(@($choice.KnownSemanticSha256) + @($earlierChoice.KnownSemanticSha256) | Sort-Object -Unique) }
                $acceptedSourceChoices += $choice
            }
            $requestedSources += [pscustomobject]@{ Role = $required.Role; Source = $selected }
        }
        $latestSource = ($requestedSources | Where-Object Role -eq 'CurrentWeek').Source
        $priorYearSource = ($requestedSources | Where-Object Role -eq 'PriorYearWeek').Source
        $priorYearWeek = $reportingDate.AddDays(-364)
        $currentSourceSha256 = $latestSource.Sha256
        $priorYearSourceSha256 = $priorYearSource.Sha256
        $updatePlan = @(New-GssWeekUpdatePlan -SelectedSources $requestedSources -ExistingRows $existingRows -CorrectionSourcePaths $approvedSourcePaths)
        $replacedWeeks = @($updatePlan | Where-Object Action -ne 'Unchanged' | ForEach-Object { $_.Source.WeekEnding.Date })
        $projectedRows = @($existingRows | Where-Object { $_.Week.Date -notin $replacedWeeks })
        foreach ($change in @($updatePlan | Where-Object Action -ne 'Unchanged')) { $projectedRows += @($change.Rows) }
        Assert-GssWorkbookScoreCompatibility -Rows $projectedRows
        $sourceEvidence = @($updatePlan | ForEach-Object {
            [pscustomobject]@{
                Role = $_.Role
                WeekEnding = $_.Source.WeekEnding.ToString('yyyy-MM-dd')
                PeriodStart = $_.Source.PeriodStart.ToString('yyyy-MM-dd')
                RelativePath = ConvertTo-GssDropboxRelativePath -Path $_.Source.File.FullName -FolderPath $FolderPath
                Sha256 = $_.Source.Sha256
                SemanticSha256 = $_.Source.SemanticSha256
            }
        })
        $sourceSetSha256 = Get-GssSourceEvidenceFingerprint $sourceEvidence
        $sourceChoiceSetSha256 = if ($acceptedSourceChoices.Count -gt 0) { Get-GssSourceChoiceFingerprint $acceptedSourceChoices } else { $null }
        $rowsRequired = 0
        foreach ($item in $updatePlan) {
            if ($item.Action -eq 'Append') { $rowsRequired += $item.Rows.Count }
            elseif ($item.Action -eq 'Correct') { $rowsRequired += [math]::Max(0, $item.Rows.Count - $item.ExistingRows) }
        }
        $lastRawDataRow = $rawWs.Cells.Item($rawWs.Rows.Count, 1).End($xlUp).Row
        $capacityPlan = Assert-GssRawDataCapacity $lastRawDataRow $rowsRequired $gssRawDataLastFormulaRow
        Remove-GssWorkbookGuardrails $targetWb
        # Validate all sources first, then batch staged cell edits. Recalculating
        # the entire model after each repaired cell is both slow and unnecessary.
        $excel.Calculation = $xlCalculationManual

        $rowsAppended = 0
        $rowsCorrected = 0
        $weeksAppended = @()
        $weeksCorrected = @()
        $weeksSkipped = @()
        foreach ($item in $updatePlan) {
            $source = $item.Source
            $record = [pscustomobject]@{
                Role = $item.Role
                WeekEnding = $source.WeekEnding.ToString('yyyy-MM-dd')
                ExistingRows = $item.ExistingRows
                SourceWorkbook = $source.File.FullName
                SourceFileNameInRawData = $source.SourceFileName
                SemanticSha256 = $source.SemanticSha256
                PreviousSemanticSha256 = $item.PreviousSemanticSha256
            }
            switch ($item.Action) {
                'Append' {
                    $written = Add-RawDataRows $rawWs $item.Rows
                    $rowsAppended += $written
                    $record | Add-Member -NotePropertyName RowsAppended -NotePropertyValue $written
                    $weeksAppended += $record
                }
                'Correct' {
                    $written = Set-GssRawDataWeek $rawWs $item.Rows
                    $rowsCorrected += $written
                    $record | Add-Member -NotePropertyName RowsCorrected -NotePropertyValue $written
                    $weeksCorrected += $record
                }
                default { $weeksSkipped += $record }
            }
        }
        $status = if ($rowsAppended -eq 0 -and $rowsCorrected -eq 0) { 'PreparedAlreadyCurrent' } else { 'Prepared' }

        $configuration = Get-GssConfigurationCounts $targetWb
        Update-EmailComparisonData $targetWb $rawWs $latestSource
        Update-GssQaSheet $targetWb $configuration
        Set-GssReadmeContent $targetWb
        Set-GssWorkbookBanners $targetWb
        Set-GssEmailComparisonLayout $targetWb
        Set-GssWorkbookPrintLayouts $targetWb $configuration
        $excel.Calculation = $xlCalculationAutomatic
        $excel.CalculateFullRebuild()
        $workbookStatus = Get-GssWorkbookStatus $targetWb
        Set-GssStatusStyles $targetWb $workbookStatus
        $emailComparisonPdfPath = Export-EmailComparisonPdf $targetWb $latestSource $FolderPath $transactionDir $false
        Set-GssWorkbookGuardrails $targetWb $configuration
        $guardrailsApplied = Test-GssWorkbookGuardrailsApplied $targetWb
        if (-not $guardrailsApplied) {
            throw 'Workbook guardrail verification failed before save.'
        }
        $targetWb.Save()

        # Excel keeps the saved workbook exclusively open until the COM workbook
        # is closed. Release it before calculating file evidence so a copy-test
        # and a live run hash the exact bytes that were persisted.
        $targetWb.Close($false)
        Release-ComObject $rawWs
        $rawWs = $null
        Release-ComObject $targetWb
        $targetWb = $null

        $postSaveValidation = Test-GssPersistedWorkbook `
            -WorkbookPath $targetPath `
            -CurrentWeek $latestSource.WeekEnding `
            -PriorYearWeek $priorYearWeek `
            -CurrentSourceFileName $latestSource.SourceFileName `
            -PriorYearSourceFileName $priorYearSource.SourceFileName `
            -ExpectedSourceEvidence $sourceEvidence

        $stagedWorkbookSha256 = Get-GssSha256 $targetPath
        $stagedPdfSha256 = Get-GssSha256 $emailComparisonPdfPath
        $fingerprint = Get-GssPreparedRunFingerprint `
            -RunId $transactionId `
            -HostName ([Environment]::MachineName) `
            -CurrentWeekEnding $latestSource.WeekEnding.ToString('yyyy-MM-dd') `
            -StartingWorkbookSha256 $startingWorkbookSha256 `
            -CurrentSourceSha256 $currentSourceSha256 `
            -PriorYearSourceSha256 $priorYearSourceSha256 `
            -StagedWorkbookSha256 $stagedWorkbookSha256 `
            -StagedPdfSha256 $stagedPdfSha256 `
            -ProgramRelease $ReleaseIdentifier `
            -SourceSetSha256 $sourceSetSha256 `
            -SourceChoiceSetSha256 $sourceChoiceSetSha256

        $fileEvidence = @(
            [pscustomobject]@{
                Role = 'rolling_workbook'
                RelativePath = ConvertTo-GssDropboxRelativePath -Path $latestSource.File.FullName -FolderPath $FolderPath
                ByteSize = [long](Get-Item -LiteralPath $latestSource.File.FullName).Length
                Sha256 = $currentSourceSha256
            },
            [pscustomobject]@{
                Role = 'prior_year_rolling_workbook'
                RelativePath = ConvertTo-GssDropboxRelativePath -Path $priorYearSource.File.FullName -FolderPath $FolderPath
                ByteSize = [long](Get-Item -LiteralPath $priorYearSource.File.FullName).Length
                Sha256 = $priorYearSourceSha256
            },
            [pscustomobject]@{
                Role = 'live_workbook'
                RelativePath = ConvertTo-GssDropboxRelativePath -Path $targetPath -FolderPath $FolderPath
                ByteSize = [long](Get-Item -LiteralPath $targetPath).Length
                Sha256 = $stagedWorkbookSha256
            },
            [pscustomobject]@{
                Role = 'comparison_pdf'
                RelativePath = ConvertTo-GssDropboxRelativePath -Path $emailComparisonPdfPath -FolderPath $FolderPath
                ByteSize = [long](Get-Item -LiteralPath $emailComparisonPdfPath).Length
                Sha256 = $stagedPdfSha256
            }
        )
        $releaseArchiveDirectory = Join-Path $automationDir 'state\release'
        $releaseArchives = if (Test-Path -LiteralPath $releaseArchiveDirectory -PathType Container) {
            @(Get-ChildItem -LiteralPath $releaseArchiveDirectory -File -Filter '*.zip' | Sort-Object Name | ForEach-Object { $_.FullName })
        }
        else {
            @()
        }

        $summary = [pscustomobject]@{
            ReceiptSchemaVersion = 2
            Timestamp = (Get-Date).ToString('s')
            TimestampUtc = [datetime]::UtcNow.ToString('o')
            RunId = $transactionId
            HostName = [Environment]::MachineName
            ProgramRelease = $ReleaseIdentifier
            Mode = 'CopyTestOnly'
            Status = $status
            TransactionStatus = 'Prepared'
            WorkbookStatus = $workbookStatus
            GuardrailsApplied = $guardrailsApplied
            RunFingerprint = $fingerprint
            Folder = $FolderPath
            CurrentWeekEnding = $latestSource.WeekEnding.ToString('yyyy-MM-dd')
            PriorYearWeekEnding = $priorYearWeek.ToString('yyyy-MM-dd')
            ReportingPeriodLabel = "13-week rolling through $($latestSource.WeekEnding.ToString('yyyy-MM-dd'))"
            RollingPeriodStart = $latestSource.PeriodStart.ToString('yyyy-MM-dd')
            RollingPeriodEnd = $latestSource.WeekEnding.ToString('yyyy-MM-dd')
            PreviousRollingWindowEnding = $latestSource.WeekEnding.AddDays(-7).ToString('yyyy-MM-dd')
            ComparisonLabel = 'change versus previous rolling window'
            CurrentSourceWorkbook = $latestSource.File.FullName
            CurrentSourceRelativePath = ConvertTo-GssDropboxRelativePath -Path $latestSource.File.FullName -FolderPath $FolderPath
            CurrentSourceFileNameInRawData = $latestSource.SourceFileName
            CurrentSourceSha256 = $currentSourceSha256
            CurrentSourceSemanticSha256 = $latestSource.SemanticSha256
            PriorYearSourceWorkbook = $priorYearSource.File.FullName
            PriorYearSourceRelativePath = ConvertTo-GssDropboxRelativePath -Path $priorYearSource.File.FullName -FolderPath $FolderPath
            PriorYearSourceFileNameInRawData = $priorYearSource.SourceFileName
            PriorYearSourceSha256 = $priorYearSourceSha256
            PriorYearSourceSemanticSha256 = $priorYearSource.SemanticSha256
            SourceSetSha256 = $sourceSetSha256
            SourceEvidence = $sourceEvidence
            SourceChoiceSetSha256 = $sourceChoiceSetSha256
            AcceptedSourceChoices = $acceptedSourceChoices
            StartingWorkbook = $mainPath
            StartingWorkbookRelativePath = ConvertTo-GssDropboxRelativePath -Path $mainPath -FolderPath $FolderPath
            StartingWorkbookSha256 = $startingWorkbookSha256
            TargetWorkbook = $targetPath
            TargetWorkbookRelativePath = ConvertTo-GssDropboxRelativePath -Path $targetPath -FolderPath $FolderPath
            StagedWorkbook = $targetPath
            StagedWorkbookRelativePath = ConvertTo-GssDropboxRelativePath -Path $targetPath -FolderPath $FolderPath
            StagedWorkbookSha256 = $stagedWorkbookSha256
            BackupWorkbook = $backupPath
            EmailComparisonPdf = $emailComparisonPdfPath
            EmailComparisonPdfRelativePath = ConvertTo-GssDropboxRelativePath -Path $emailComparisonPdfPath -FolderPath $FolderPath
            StagedPdf = $emailComparisonPdfPath
            StagedPdfRelativePath = ConvertTo-GssDropboxRelativePath -Path $emailComparisonPdfPath -FolderPath $FolderPath
            StagedPdfSha256 = $stagedPdfSha256
            TransactionArtifacts = @(
                [pscustomobject]@{
                    SourcePath = $targetPath
                    Role = 'reviewed_staged_workbook'
                    Classification = 'restricted_operational'
                    Sha256 = $stagedWorkbookSha256
                },
                [pscustomobject]@{
                    SourcePath = $emailComparisonPdfPath
                    Role = 'reviewed_staged_comparison_pdf'
                    Classification = 'restricted_operational'
                    Sha256 = $stagedPdfSha256
                }
            )
            ReleaseArchives = $releaseArchives
            FileEvidence = $fileEvidence
            RowsAppended = $rowsAppended
            RowsCorrected = $rowsCorrected
            BusinessDataChanged = ($rowsAppended -gt 0 -or $rowsCorrected -gt 0)
            NoBusinessChanges = ($rowsAppended -eq 0 -and $rowsCorrected -eq 0)
            WeeksAppended = $weeksAppended
            WeeksCorrected = $weeksCorrected
            WeeksSkipped = $weeksSkipped
            PostSaveValidation = $postSaveValidation
            WorstCaseWeeklyLoadsRemaining = [int]$postSaveValidation.WorstCaseWeeklyLoadsRemaining
            PromotionPolicy = 'ExactFingerprintAndStartingHashesRequired'
        }

        $summary | Add-Member -NotePropertyName LogPath -NotePropertyValue $logPath
        Write-GssAtomicJson -Path $logPath -InputObject $summary
        Write-RunSummary $summary
        if ($ReturnObject) {
            return $summary
        }
    }
    catch {
        $failure = $_
        if (-not (Test-Path -LiteralPath $logPath -PathType Leaf)) {
            $failureReceipt = [pscustomobject]@{
                ReceiptSchemaVersion = 2
                TimestampUtc = [datetime]::UtcNow.ToString('o')
                RunId = $transactionId
                HostName = [Environment]::MachineName
                Mode = if ($ApplyMode) { 'ApplyToMainWorkbook' } elseif ($RollbackLogPath) { 'Rollback' } else { 'CopyTestOnly' }
                Status = 'Blocked'
                TransactionStatus = 'Blocked'
                Error = $failure.Exception.Message
            }
            Write-GssAtomicJson -Path $logPath -InputObject $failureReceipt
        }
        throw
    }
    finally {
        if ($targetWb) { $targetWb.Close($false) }
        if ($sourceWb) { $sourceWb.Close($false) }
        Release-ComObject $rawWs
        Release-ComObject $targetWb
        Release-ComObject $sourceWs
        Release-ComObject $sourceWb
        if ($excel) {
            $excel.Quit()
            Release-ComObject $excel
        }
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
        if ($ownsMutex -and $mutex) {
            try { $mutex.ReleaseMutex() }
            catch { Write-Verbose "The GSS transaction mutex could not be released cleanly: $($_.Exception.Message)" }
        }
        if ($mutex) { $mutex.Dispose() }
    }
}

if ($MyInvocation.InvocationName -ne '.') {
    Invoke-GssWorkbookUpdate `
        -FolderPath $Folder `
        -WorkbookName $MainWorkbookName `
        -ApplyMode:$Apply `
        -ReturnObject:$OutputObject `
        -RunIdentifier $RunId `
        -PreparedLogPath $PreparedRunLogPath `
        -ReviewedFingerprint $ExpectedFingerprint `
        -RollbackLogPath $RollbackRunLogPath `
        -CallerOwnsMutex:$MutexAlreadyHeld `
        -ReleaseIdentifier $ProgramRelease `
        -PreparedDriveManifestPath $DrivePreparedManifestPath `
        -PreparedDriveManifestSha256 $ExpectedDrivePreparedManifestSha256 `
        -ReportingWeek $ReportingWeekEnding `
        -CorrectionSourcePaths $CorrectionSourcePath
}
