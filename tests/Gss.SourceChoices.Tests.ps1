BeforeAll {
    . (Join-Path (Split-Path -Parent $PSScriptRoot) 'scripts\Gss-SourceChoices.ps1')

    function Get-ChoiceCandidate {
        param([string]$Name, [string]$Digest, [datetime]$Week = [datetime]'2026-10-04')
        [pscustomobject]@{
            WeekEnding = $Week
            SemanticSha256 = $Digest
            File = [pscustomobject]@{ FullName = "C:\GSS Test\$Name" }
        }
    }
}

Describe 'Accepted rolling source revisions' {
    BeforeEach {
        $original = Get-ChoiceCandidate 'original.xlsx' ('a' * 64)
        $correction = Get-ChoiceCandidate 'corrected.xlsx' ('b' * 64)
        $sources = @($original, $correction)
        $script:accepted = New-GssSourceChoiceEvidence -Sources $sources -SelectedSourcePath $correction.File.FullName
    }

    It 'remembers the chosen content and known alternatives without storing paths' {
        $script:accepted.WeekEnding | Should -Be '2026-10-04'
        $script:accepted.SelectedSemanticSha256 | Should -Be ('b' * 64)
        @($script:accepted.KnownSemanticSha256).Count | Should -Be 2
        ($script:accepted | ConvertTo-Json) | Should -Not -Match '\.xlsx'
        (Get-GssAcceptedSourceChoicePath -Sources $sources -AcceptedChoices @($script:accepted)) | Should -Be $correction.File.FullName
    }

    It 'reuses a committed choice after rename, relocation, duplicate aliases, or removal of the old version' {
        $renamed = Get-ChoiceCandidate 'renamed.xlsx' ('b' * 64)
        $renamed.File.FullName = 'D:\Shared GSS\renamed.xlsx'
        (Get-GssAcceptedSourceChoicePath -Sources @($renamed) -AcceptedChoices @($script:accepted)) | Should -Be $renamed.File.FullName
        $alias = Get-ChoiceCandidate 'a-alias.xlsx' ('b' * 64)
        (Get-GssAcceptedSourceChoicePath -Sources @($original, $correction, $alias) -AcceptedChoices @($script:accepted)) | Should -Be $alias.File.FullName
    }

    It 'requires a new choice when an unknown business revision appears' {
        $newRevision = Get-ChoiceCandidate 'new-correction.xlsx' ('c' * 64)
        (Get-GssAcceptedSourceChoicePath -Sources ($sources + $newRevision) -AcceptedChoices @($script:accepted)) | Should -BeNullOrEmpty
        $reaccepted = New-GssSourceChoiceEvidence -Sources ($sources + $newRevision) -SelectedSourcePath $newRevision.File.FullName
        (Get-GssAcceptedSourceChoicePath -Sources ($sources + $newRevision) -AcceptedChoices @($reaccepted, $script:accepted)) | Should -Be $newRevision.File.FullName
    }

    It 'does not silently return to an older revision when the latest chosen content is missing' {
        $earlier = New-GssSourceChoiceEvidence -Sources @($original) -SelectedSourcePath $original.File.FullName
        (Get-GssAcceptedSourceChoicePath -Sources @($original) -AcceptedChoices @($script:accepted, $earlier)) | Should -BeNullOrEmpty
    }

    It 'isolates decisions by reporting period and makes no choice without accepted evidence' {
        $other = Get-ChoiceCandidate 'other-week.xlsx' ('b' * 64) ([datetime]'2026-09-27')
        (Get-GssAcceptedSourceChoicePath -Sources @($other) -AcceptedChoices @($script:accepted)) | Should -BeNullOrEmpty
        (Get-GssAcceptedSourceChoicePath -Sources $sources -AcceptedChoices @()) | Should -BeNullOrEmpty
        (Get-GssAcceptedSourceChoicePath -Sources @() -AcceptedChoices @($script:accepted)) | Should -BeNullOrEmpty
    }

    It 'fails closed for malformed evidence and mixed reporting periods' {
        $script:accepted.KnownSemanticSha256 = @('not-a-fingerprint')
        (Get-GssAcceptedSourceChoicePath -Sources $sources -AcceptedChoices @($script:accepted)) | Should -BeNullOrEmpty
        $other = Get-ChoiceCandidate 'other-week.xlsx' ('b' * 64) ([datetime]'2026-09-27')
        { Get-GssAcceptedSourceChoicePath -Sources @($original, $other) } | Should -Throw '*exactly one reporting week*'
        { New-GssSourceChoiceEvidence -Sources @($original, $other) -SelectedSourcePath $original.File.FullName } | Should -Throw '*exactly one reporting week*'
        { New-GssSourceChoiceEvidence -Sources $sources -SelectedSourcePath 'missing.xlsx' } | Should -Throw '*exactly one validated candidate*'
        $correction.SemanticSha256 = 'invalid'
        { Get-GssAcceptedSourceChoicePath -Sources $sources } | Should -Throw '*validated content fingerprints*'
    }
}
