# Accepted rolling-report revisions are content decisions, independent of filenames.
# Pass only decisions from committed runs, newest first. Prepared choices are not authority.

function Get-GssAcceptedSourceChoicePath {
    param(
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Sources,
        [AllowEmptyCollection()][object[]]$AcceptedChoices = @()
    )

    if ($Sources.Count -eq 0) { return $null }
    $weeks = @($Sources | ForEach-Object { ([datetime]$_.WeekEnding).ToString('yyyy-MM-dd') } | Sort-Object -Unique)
    if ($weeks.Count -ne 1) { throw 'An accepted source choice must concern exactly one reporting week.' }
    foreach ($source in $Sources) {
        if ([string]$source.SemanticSha256 -cnotmatch '^[a-f0-9]{64}$') { throw 'Source choice candidates require validated content fingerprints.' }
    }

    # Never fall back to an older choice when the latest chosen revision is absent.
    # Removing a corrected file must not silently reinstate its superseded version.
    $latest = @($AcceptedChoices | Where-Object { $_ -and [string]$_.WeekEnding -eq $weeks[0] } | Select-Object -First 1)
    if ($latest.Count -eq 0) { return $null }
    $choice = $latest[0]
    $selected = [string]$choice.SelectedSemanticSha256
    $known = @($choice.KnownSemanticSha256 | ForEach-Object { [string]$_ } | Sort-Object -Unique)
    if ($selected -cnotmatch '^[a-f0-9]{64}$' -or $selected -cnotin $known -or
        @($known | Where-Object { $_ -cnotmatch '^[a-f0-9]{64}$' }).Count -gt 0) { return $null }
    if (@($Sources | Where-Object { [string]$_.SemanticSha256 -cnotin $known }).Count -gt 0) { return $null }
    $matchingSources = @($Sources | Where-Object { [string]$_.SemanticSha256 -ceq $selected })
    if ($matchingSources.Count -eq 0) { return $null }
    return [string](($matchingSources | Sort-Object { $_.File.FullName.ToLowerInvariant() }, { $_.File.FullName } | Select-Object -First 1).File.FullName)
}

function New-GssSourceChoiceEvidence {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Returns an in-memory decision record; the caller persists it only with a reporting receipt.')]
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][object[]]$Sources,
        [Parameter(Mandatory)][string]$SelectedSourcePath
    )

    if ($Sources.Count -eq 0) { throw 'A source choice requires at least one candidate.' }
    $weeks = @($Sources | ForEach-Object { ([datetime]$_.WeekEnding).ToString('yyyy-MM-dd') } | Sort-Object -Unique)
    if ($weeks.Count -ne 1) { throw 'A source choice must concern exactly one reporting week.' }
    foreach ($source in $Sources) {
        if ([string]$source.SemanticSha256 -cnotmatch '^[a-f0-9]{64}$') { throw 'Source choice candidates require validated content fingerprints.' }
    }
    $selected = @($Sources | Where-Object { [string]$_.File.FullName -eq $SelectedSourcePath })
    if ($selected.Count -ne 1) { throw 'The selected source path must identify exactly one validated candidate.' }
    return [pscustomobject]@{
        WeekEnding = $weeks[0]
        SelectedSemanticSha256 = [string]$selected[0].SemanticSha256
        KnownSemanticSha256 = @($Sources | ForEach-Object { [string]$_.SemanticSha256 } | Sort-Object -Unique)
    }
}
