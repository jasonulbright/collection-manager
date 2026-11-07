#Requires -Modules Pester
# Tests of the window-test guard (Test-WindowTestsAllowed in WindowHarness.ps1).
# The guard runs window tests only when SUITE_ALLOW_WINDOW_TESTS is 1. This file
# injects the opt-in value, so it never reads the real environment and never
# opens a window.

BeforeDiscovery {
    # Every other test file that can open a window must carry the guard.
    # Same scope as Invoke-Checks.ps1: nested folders count, the lab scripts do not run as tests.
    $script:WindowTestFiles = @(Get-ChildItem -LiteralPath $PSScriptRoot -Filter '*.Tests.ps1' -File -Recurse |
            Where-Object { $_.Name -ne 'WindowGuard.Tests.ps1' -and $_.FullName -notlike '*\Lab\*' -and (Get-Content -LiteralPath $_.FullName -Raw) -match 'Invoke-WpfSta|New-WpfWindow|ShowDialog|\.Show\(\)' } |
            ForEach-Object { @{ Name = $_.FullName.Substring($PSScriptRoot.Length + 1); Path = $_.FullName } })
}

BeforeAll {
    . "$PSScriptRoot\WindowHarness.ps1"
    function Get-Guard { param($Value) Test-WindowTestsAllowed -OptIn $Value }
}

Describe 'Test-WindowTestsAllowed' {
    It 'allows when the opt-in is 1' {
        $r = Get-Guard '1'
        $r.Allowed | Should -BeTrue
        $r.Reason | Should -BeNullOrEmpty
    }

    It 'refuses when the opt-in is <Label>, and names the opt-in' -ForEach @(
        @{ Label = 'empty'; Value = '' }
        @{ Label = 'null'; Value = $null }
        @{ Label = '0'; Value = '0' }
        @{ Label = 'true'; Value = 'true' }
    ) {
        $r = Get-Guard $Value
        $r.Allowed | Should -BeFalse
        $r.Reason | Should -Be 'SUITE_ALLOW_WINDOW_TESTS is not 1'
    }
}

Describe 'Test-WindowTestsAllowed decides on the opt-in only' {
    It 'has no ComputerName parameter' {
        (Get-Command Test-WindowTestsAllowed).Parameters.Keys | Should -Not -Contain 'ComputerName'
    }

    It 'allows when the opt-in is 1, whatever the machine' {
        (Test-WindowTestsAllowed -OptIn '1').Allowed | Should -BeTrue
    }
}

Describe 'Test-WindowTestsAllowed reads SUITE_ALLOW_WINDOW_TESTS when called without -OptIn' {
    BeforeAll { $script:SavedOptIn = $env:SUITE_ALLOW_WINDOW_TESTS }
    AfterAll {
        if ($null -eq $script:SavedOptIn) { Remove-Item Env:\SUITE_ALLOW_WINDOW_TESTS -ErrorAction SilentlyContinue }
        else { $env:SUITE_ALLOW_WINDOW_TESTS = $script:SavedOptIn }
    }

    It 'allows when the variable is 1' {
        $env:SUITE_ALLOW_WINDOW_TESTS = '1'
        $r = Test-WindowTestsAllowed
        $r.Allowed | Should -BeTrue
        $r.Reason | Should -BeNullOrEmpty
    }

    It 'refuses when the variable is <Label>' -ForEach @(
        @{ Label = '0'; Value = '0' }
        @{ Label = 'true'; Value = 'true' }
        @{ Label = 'empty'; Value = '' }
    ) {
        $env:SUITE_ALLOW_WINDOW_TESTS = $Value
        $r = Test-WindowTestsAllowed
        $r.Allowed | Should -BeFalse
        $r.Reason | Should -Be 'SUITE_ALLOW_WINDOW_TESTS is not 1'
    }

    It 'refuses when the variable is not set' {
        Remove-Item Env:\SUITE_ALLOW_WINDOW_TESTS -ErrorAction SilentlyContinue
        (Test-WindowTestsAllowed).Allowed | Should -BeFalse
    }
}

Describe 'Window test files carry the guard' {
    It 'finds at least one test file that can open a window' -ForEach @(@{ Found = @($script:WindowTestFiles).Count }) {
        $Found | Should -BeGreaterThan 0
    }

    It 'prints the skip reason once: Invoke-Checks sets the marker and Window.Tests reads it' {
        (Get-Content -LiteralPath (Join-Path (Split-Path -Parent $PSScriptRoot) 'Invoke-Checks.ps1') -Raw) | Should -Match '\$env:SUITE_WINDOW_GUARD_REPORTED\s*=\s*''1'''
        (Get-Content -LiteralPath (Join-Path $PSScriptRoot 'Window.Tests.ps1') -Raw) | Should -Match '-not \$env:SUITE_WINDOW_GUARD_REPORTED'
    }

    It '<Name> calls Test-WindowTestsAllowed and applies -Skip' -ForEach $script:WindowTestFiles {
        $text = Get-Content -LiteralPath $Path -Raw
        $text | Should -Match 'Test-WindowTestsAllowed'
        $text | Should -Match '-Skip:'
    }

    It '<Name> applies the guard to every top-level Describe' -ForEach $script:WindowTestFiles {
        $describes = @(Get-Content -LiteralPath $Path | Where-Object { $_ -match '^\s*Describe\b' })
        $describes.Count | Should -BeGreaterThan 0
        @($describes | Where-Object { $_ -notmatch '-Skip:\(-not \$script:WindowGuard\.Allowed\)' }) | Should -BeNullOrEmpty
    }
}

Describe 'Get-WpfSession refuses where the guard says no' {
    BeforeAll {
        Mock Test-WindowTestsAllowed { [pscustomobject]@{ Allowed = $false; Reason = 'test' } }
    }
    AfterEach { $script:HarnessSession = $null }

    It 'throws and creates no session' {
        $script:HarnessSession = $null
        { Get-WpfSession } | Should -Throw '*not allowed*'
        $script:HarnessSession | Should -BeNullOrEmpty
    }

    It 'does not hand out a cached session' {
        $script:HarnessSession = [pscustomobject]@{ RunspaceStateInfo = [pscustomobject]@{ State = 'Opened' } }
        { Get-WpfSession } | Should -Throw '*not allowed*'
    }
}
