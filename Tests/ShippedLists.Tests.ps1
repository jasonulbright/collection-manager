#Requires -Modules Pester
# The shipped template lists against the reference values in ShippedVersions.psd1.
# A template with a missing field, an OS build that the reference does not know, a client or
# Microsoft 365 Apps build that does not match its version, or a missing newest entry fails here.
# Refresh steps: docs\VERSION-REFRESH.md.

BeforeAll {
    . "$PSScriptRoot\TestHelpers.ps1"
    Import-AppModule
    $script:Root = Split-Path -Parent $PSScriptRoot
    $script:Ref = Import-PowerShellDataFile -LiteralPath (Join-Path $PSScriptRoot 'ShippedVersions.psd1')
    # Windows PowerShell 5.1 returns a JSON array as one object. The loop unrolls it.
    function script:Read-TemplateFile {
        param([string]$Path)
        $items = ConvertFrom-Json -InputObject (Get-Content -LiteralPath $Path -Raw)
        foreach ($i in $items) { $i }
    }
    $script:Operational = @(Read-TemplateFile (Join-Path $script:Root 'Templates\operational-collections.json'))
    $script:Parameterized = @(Read-TemplateFile (Join-Path $script:Root 'Templates\parameterized-templates.json'))

    function script:Get-Clause {
        param([string]$Query, [string]$Pattern)
        $m = [regex]::Match($Query, $Pattern)
        if ($m.Success) { $m.Groups[1].Value } else { $null }
    }
}

Describe 'Operational templates' {
    It 'has the field <_> on every template' -ForEach 'Name', 'Category', 'LimitingCollection', 'Query', 'RefreshType' {
        $field = $_
        $missing = foreach ($t in $script:Operational) { if ([string]::IsNullOrWhiteSpace([string]$t.$field)) { "$($t.Name)" } }
        @($missing) | Should -BeNullOrEmpty
    }

    It 'has no duplicate template name' {
        $dups = @($script:Operational | Group-Object Name | Where-Object Count -gt 1 | ForEach-Object Name)
        $dups | Should -BeNullOrEmpty
    }

    It 'uses a known RefreshType' {
        @($script:Operational | Where-Object { $_.RefreshType -notin 'Manual', 'Periodic', 'Continuous', 'Both' } | ForEach-Object Name) | Should -BeNullOrEmpty
    }
}

Describe 'Parameterized templates' {
    It 'has the field <_> on every template' -ForEach 'Name', 'Category', 'Description', 'Query', 'LimitingCollection' {
        $field = $_
        $missing = foreach ($t in $script:Parameterized) { if ([string]::IsNullOrWhiteSpace([string]$t.$field)) { "$($t.Name)" } }
        @($missing) | Should -BeNullOrEmpty
    }

    It 'declares each placeholder with a label and uses it in the query' {
        $bad = foreach ($t in $script:Parameterized) {
            foreach ($p in @($t.Parameters)) {
                if (-not $p.Placeholder -or -not $p.Label -or ($t.Query -notlike "*$($p.Placeholder)*")) { "$($t.Name): $($p.Placeholder)" }
            }
        }
        @($bad) | Should -BeNullOrEmpty
    }

    It 'has no placeholder in the query that the parameters do not declare' {
        $bad = foreach ($t in $script:Parameterized) {
            $declared = @($t.Parameters | ForEach-Object Placeholder)
            foreach ($m in [regex]::Matches($t.Query, '\{[A-Za-z]+\}')) { if ($m.Value -notin $declared) { "$($t.Name): $($m.Value)" } }
        }
        @($bad) | Should -BeNullOrEmpty
    }
}

Describe 'ConfigMgr client versions' {
    It 'matches the reference build for every version template' {
        $bad = foreach ($t in $script:Operational) {
            if ($t.Name -notmatch '^Clients Version \| (\d{4})$') { continue }
            $ver = $Matches[1]
            $build = Get-Clause $t.Query "ClientVersion (?:like|=) '5\.00\.(\d+)\."
            if (-not $script:Ref.ConfigMgrClient.Versions.ContainsKey($ver)) { "$($t.Name): unknown ConfigMgr version $ver" }
            elseif ($build -ne $script:Ref.ConfigMgrClient.Versions[$ver]) { "$($t.Name): build $build, reference $($script:Ref.ConfigMgrClient.Versions[$ver])" }
        }
        @($bad) | Should -BeNullOrEmpty
    }

    It 'has a template for the newest version' {
        $names = @($script:Operational | ForEach-Object Name)
        $names | Should -Contain "Clients Version | $($script:Ref.Newest.ConfigMgrClient)"
    }

    It 'names the newest version in the Not Latest template and uses its build' {
        $newest = $script:Ref.Newest.ConfigMgrClient
        $t = @($script:Operational | Where-Object { $_.Name -like 'Clients Version | Not Latest (*' })
        $t.Count | Should -Be 1
        $t[0].Name | Should -Be "Clients Version | Not Latest ($newest)"
        $t[0].Query | Should -BeLike "*ClientVersion not like '5.00.$($script:Ref.ConfigMgrClient.Versions[$newest]).10%'*"
    }

    It 'has no Clients Version template with a name that the checks cannot read' {
        $bad = foreach ($t in $script:Operational) {
            if ($t.Category -eq 'Clients Version' -and $t.Name -notmatch '^Clients Version \| (\d{4}|R2 .+|Not Latest \(\d{4}\))$') { $t.Name }
        }
        @($bad) | Should -BeNullOrEmpty
    }
}

Describe 'Microsoft 365 Apps' {
    It 'matches the reference build for every build template' {
        $bad = foreach ($t in $script:Operational) {
            if ($t.Name -notmatch '^(?:Microsoft 365 Apps Build|Office 365 Build Version) \| (\d{4})$') { continue }
            $ver = $Matches[1]
            $build = Get-Clause $t.Query "VersionToReport like '16\.0\.(\d+)\."
            $expected = if ($script:Ref.M365AppsBuild.Versions.ContainsKey($ver)) { $script:Ref.M365AppsBuild.Versions[$ver] } else { $script:Ref.M365AppsBuild.Unverified[$ver] }
            if (-not $expected) { "$($t.Name): unknown Microsoft 365 Apps version $ver" }
            elseif ($build -ne $expected) { "$($t.Name): build $build, reference $expected" }
        }
        @($bad) | Should -BeNullOrEmpty
    }

    It 'has a template for every reference version from 2208 on' {
        $have = foreach ($t in $script:Operational) { if ($t.Name -match '^Microsoft 365 Apps Build \| (\d{4})$') { $Matches[1] } }
        $want = @($script:Ref.M365AppsBuild.Versions.Keys | Where-Object { [int]$_ -ge 2208 })
        @($want | Where-Object { $_ -notin $have } | Sort-Object) | Should -BeNullOrEmpty
    }

    It 'has a template for the newest version' {
        @($script:Operational | ForEach-Object Name) | Should -Contain "Microsoft 365 Apps Build | $($script:Ref.Newest.M365AppsBuild)"
    }

    It 'matches the reference GUID for every channel template' {
        $bad = foreach ($t in $script:Operational) {
            if ($t.Name -notmatch '^Microsoft 365 Apps Channel \| (.+)$') { continue }
            $channel = $Matches[1]
            $guid = Get-Clause $t.Query 'officecdn\.microsoft\.com/pr/([0-9a-f-]{36})'
            if (-not $script:Ref.M365AppsChannel.Channels.ContainsKey($channel)) { "$($t.Name): unknown channel" }
            elseif ($guid -ne $script:Ref.M365AppsChannel.Channels[$channel]) { "$($t.Name): GUID $guid" }
        }
        @($bad) | Should -BeNullOrEmpty
    }
}

Describe 'Windows builds' {
    It 'matches the reference build for every Windows client template' {
        $bad = foreach ($t in $script:Operational) {
            if ($t.Name -notmatch '^Workstations \| (Windows (?:10|11) v\w+)$') { continue }
            $label = $Matches[1]
            $build = Get-Clause $t.Query "SMS_R_System\.Build = '10\.0\.(\d+)'"
            if (-not $script:Ref.WindowsClientBuild.Builds.ContainsKey($label)) { "$($t.Name): unknown OS build name" }
            elseif ($build -ne $script:Ref.WindowsClientBuild.Builds[$label]) { "$($t.Name): build $build, reference $($script:Ref.WindowsClientBuild.Builds[$label])" }
        }
        @($bad) | Should -BeNullOrEmpty
    }

    It 'has no OS build in a Build clause that the reference does not know' {
        $known = @($script:Ref.WindowsClientBuild.Builds.Values)
        $bad = foreach ($t in $script:Operational) {
            foreach ($m in [regex]::Matches($t.Query, "SMS_R_System\.Build = '10\.0\.(\d+)'")) { if ($m.Groups[1].Value -notin $known) { "$($t.Name): $($m.Groups[1].Value)" } }
        }
        @($bad) | Should -BeNullOrEmpty
    }

    It 'has a template for every reference Windows client build' {
        $names = @($script:Operational | ForEach-Object Name)
        $bad = foreach ($label in $script:Ref.WindowsClientBuild.Builds.Keys) { if ("Workstations | $label" -notin $names) { $label } }
        @($bad | Sort-Object) | Should -BeNullOrEmpty
    }

    It 'matches the reference build for every Windows Server template' {
        $bad = foreach ($t in $script:Operational) {
            if ($t.Name -notmatch '^Servers \| (Windows 20(?:16|19|22|25))$') { continue }
            $label = $Matches[1]
            $build = Get-Clause $t.Query "OPERATING_SYSTEM\.BuildNumber = '(\d+)'"
            if (-not $script:Ref.WindowsServerBuild.Builds.ContainsKey($label)) { "$($t.Name): unknown server name" }
            elseif ($build -ne $script:Ref.WindowsServerBuild.Builds[$label]) { "$($t.Name): build $build, reference $($script:Ref.WindowsServerBuild.Builds[$label])" }
        }
        @($bad) | Should -BeNullOrEmpty
    }

    It 'has a template for every reference Windows Server build' {
        $names = @($script:Operational | ForEach-Object Name)
        $bad = foreach ($label in $script:Ref.WindowsServerBuild.Builds.Keys) { if ("Servers | $label" -notin $names) { $label } }
        @($bad | Sort-Object) | Should -BeNullOrEmpty
    }

    It 'has a template for the newest Windows 10 build and Windows Server build' {
        $names = @($script:Operational | ForEach-Object Name)
        $names | Should -Contain "Workstations | $($script:Ref.Newest.Windows10Build)"
        $names | Should -Contain "Servers | $($script:Ref.Newest.ServerBuild)"
    }

    It 'has a template for the newest Windows 11 build' {
        @($script:Operational | ForEach-Object Name) | Should -Contain "Workstations | $($script:Ref.Newest.Windows11Build)"
    }
}

