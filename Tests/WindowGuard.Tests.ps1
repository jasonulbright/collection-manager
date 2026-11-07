#Requires -Modules Pester
# Tests of the window-test guard (Test-WindowTestsAllowed in WindowHarness.ps1).
# The guard keeps window tests off the dev PC (AYW) and runs them only when
# SUITE_ALLOW_WINDOW_TESTS is 1. This file injects the computer name and the
# opt-in value, so it never reads the real environment and never opens a window.

BeforeDiscovery {
    # Every other test file that can open a window must carry the guard.
    $script:WindowTestFiles = @(Get-ChildItem -LiteralPath $PSScriptRoot -Filter '*.Tests.ps1' -File |
            Where-Object { $_.Name -ne 'WindowGuard.Tests.ps1' -and (Get-Content -LiteralPath $_.FullName -Raw) -match 'Invoke-WpfSta|New-WpfWindow|ShowDialog|\.Show\(\)' } |
            ForEach-Object { @{ Name = $_.Name; Path = $_.FullName } })
}

BeforeAll {
    . "$PSScriptRoot\WindowHarness.ps1"
    # Pass the computer name as a variable, so the linter rule against a hard-coded ComputerName stays quiet.
    function Get-Guard { param([string]$Computer, $Value) Test-WindowTestsAllowed -ComputerName $Computer -OptIn $Value }
}

Describe 'Test-WindowTestsAllowed' {
    It 'refuses on AYW even with the opt-in' {
        $r = Get-Guard 'AYW' '1'
        $r.Allowed | Should -BeFalse
        $r.Reason | Should -Match 'AYW'
    }

    It 'refuses on ayw in lower case' {
        (Get-Guard 'ayw' '1').Allowed | Should -BeFalse
    }

    It 'refuses on AYW without the opt-in' {
        (Get-Guard 'AYW' '').Allowed | Should -BeFalse
    }

    It 'allows another computer with the opt-in 1' {
        $r = Get-Guard 'CM01' '1'
        $r.Allowed | Should -BeTrue
        $r.Reason | Should -BeNullOrEmpty
    }

    It 'refuses another computer when the opt-in is <Label>, and names the opt-in' -ForEach @(
        @{ Label = 'empty'; Value = '' }
        @{ Label = 'null'; Value = $null }
        @{ Label = '0'; Value = '0' }
        @{ Label = 'true'; Value = 'true' }
    ) {
        $r = Get-Guard 'CLIENT01' $Value
        $r.Allowed | Should -BeFalse
        $r.Reason | Should -Match 'SUITE_ALLOW_WINDOW_TESTS'
    }

    It 'gives a reason whenever it refuses (<Name>, opt-in "<Value>")' -ForEach @(
        @{ Name = 'AYW'; Value = '1' }
        @{ Name = 'CM01'; Value = '' }
        @{ Name = 'CM01'; Value = '0' }
    ) {
        (Get-Guard $Name $Value).Reason | Should -Not -BeNullOrEmpty
    }
}

Describe 'Window test files carry the guard' {
    It 'finds at least one test file that can open a window' -ForEach @(@{ Found = @($script:WindowTestFiles).Count }) {
        $Found | Should -BeGreaterThan 0
    }

    It '<Name> calls Test-WindowTestsAllowed and applies -Skip' -ForEach $script:WindowTestFiles {
        $text = Get-Content -LiteralPath $Path -Raw
        $text | Should -Match 'Test-WindowTestsAllowed'
        $text | Should -Match '-Skip:'
    }
}
