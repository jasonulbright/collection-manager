#Requires -Modules Pester
# Logic of start-collectionmanager.ps1. The script cannot be dot-sourced (it opens
# the window), so each function is cut out of its syntax tree and defined here,
# with plain objects in place of the window controls.

BeforeAll {
    . "$PSScriptRoot\TestHelpers.ps1"
    Import-AppModule
    $script:ShellPath = Join-Path $script:AppRoot 'start-collectionmanager.ps1'
    $script:Ast = [System.Management.Automation.Language.Parser]::ParseFile($script:ShellPath, [ref]$null, [ref]$null)
    $wanted = 'Get-CmPreferences', 'Save-CmPreferences', 'Get-CollectionTypeGlyph', 'Get-CollectionFilterValue', 'Test-CollectionFilterMatch',
        'Update-Filter', 'Set-StatusText', 'Update-StatusBarSummary', 'Get-ActiveExportInfo'
    foreach ($n in $wanted) {
        $def = $script:Ast.Find({ param($x) $x -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $x.Name -eq $n }, $true)
        if (-not $def) { throw "function $n not found in the script" }
        . ([scriptblock]::Create($def.Extent.Text))
    }
}

Describe 'Get-CollectionTypeGlyph' {
    It 'marks a built-in collection and leaves a custom one blank' {
        [int](Get-CollectionTypeGlyph -Collection ([PSCustomObject]@{ IsBuiltIn = $true })) | Should -Be 0x22EF
        Get-CollectionTypeGlyph -Collection ([PSCustomObject]@{ IsBuiltIn = $false }) | Should -Be ''
        Get-CollectionTypeGlyph -Collection ([PSCustomObject]@{ IsBuiltIn = $null }) | Should -Be ''
    }
}

Describe 'Test-CollectionFilterMatch' {
    It 'filter "<Filter>" on <Case> gives <Expected>' -ForEach @(
        @{ Filter = 'All';                  Case = 'a custom empty row';     Row = @{ IsBuiltIn = $false; MemberCount = 0; DirectCount = 0; QueryCount = 0 }; Expected = $true }
        @{ Filter = 'Built-in only';        Case = 'a built-in row';         Row = @{ IsBuiltIn = $true;  MemberCount = 5; DirectCount = 0; QueryCount = 0 }; Expected = $true }
        @{ Filter = 'Built-in only';        Case = 'a custom row';           Row = @{ IsBuiltIn = $false; MemberCount = 5; DirectCount = 0; QueryCount = 0 }; Expected = $false }
        @{ Filter = 'Custom only';          Case = 'a custom row';           Row = @{ IsBuiltIn = $false; MemberCount = 5; DirectCount = 0; QueryCount = 0 }; Expected = $true }
        @{ Filter = 'Custom only';          Case = 'a built-in row';         Row = @{ IsBuiltIn = $true;  MemberCount = 5; DirectCount = 0; QueryCount = 0 }; Expected = $false }
        @{ Filter = 'Empty (zero members)'; Case = 'zero members';           Row = @{ IsBuiltIn = $false; MemberCount = 0; DirectCount = 0; QueryCount = 0 }; Expected = $true }
        @{ Filter = 'Empty (zero members)'; Case = 'one member';             Row = @{ IsBuiltIn = $false; MemberCount = 1; DirectCount = 0; QueryCount = 0 }; Expected = $false }
        @{ Filter = 'Has direct members';   Case = 'a direct rule';          Row = @{ IsBuiltIn = $false; MemberCount = 1; DirectCount = 2; QueryCount = 0 }; Expected = $true }
        @{ Filter = 'Has direct members';   Case = 'no direct rule';         Row = @{ IsBuiltIn = $false; MemberCount = 1; DirectCount = 0; QueryCount = 3 }; Expected = $false }
        @{ Filter = 'Has query rules';      Case = 'a query rule';           Row = @{ IsBuiltIn = $false; MemberCount = 1; DirectCount = 0; QueryCount = 3 }; Expected = $true }
        @{ Filter = 'Has query rules';      Case = 'no query rule';          Row = @{ IsBuiltIn = $false; MemberCount = 1; DirectCount = 2; QueryCount = 0 }; Expected = $false }
        @{ Filter = 'Something else';       Case = 'any row';                Row = @{ IsBuiltIn = $false; MemberCount = 1; DirectCount = 0; QueryCount = 0 }; Expected = $true }
        @{ Filter = '';                     Case = 'any row';                Row = @{ IsBuiltIn = $true;  MemberCount = 0; DirectCount = 0; QueryCount = 0 }; Expected = $true }
    ) {
        Test-CollectionFilterMatch -Row ([PSCustomObject]$Row) -Filter $Filter | Should -Be $Expected
    }

    It 'treats null counts as zero' {
        Test-CollectionFilterMatch -Row ([PSCustomObject]@{ IsBuiltIn = $null; MemberCount = $null; DirectCount = $null; QueryCount = $null }) -Filter 'Empty (zero members)' | Should -BeTrue
    }
}

Describe 'Get-CollectionFilterValue' {
    It 'returns All when nothing is selected, the item content for a combo box item, and a string as it is' {
        Add-Type -AssemblyName PresentationFramework
        $cboCollectionFilter = [PSCustomObject]@{ SelectedItem = $null }
        Get-CollectionFilterValue | Should -Be 'All'
        $item = New-Object System.Windows.Controls.ComboBoxItem
        $item.Content = 'Custom only'
        $cboCollectionFilter = [PSCustomObject]@{ SelectedItem = $item }
        Get-CollectionFilterValue | Should -Be 'Custom only'
        $cboCollectionFilter = [PSCustomObject]@{ SelectedItem = 'Has query rules' }
        Get-CollectionFilterValue | Should -Be 'Has query rules'
    }
}

Describe 'Update-Filter' {
    BeforeEach {
        $script:ActiveView = 'Collections'
        $txtFilter = [PSCustomObject]@{ Text = '' }
        $cboCollectionFilter = [PSCustomObject]@{ SelectedItem = 'All' }
        $gridCollections = [PSCustomObject]@{ ItemsSource = $null }
        $gridOperationalTemplates = [PSCustomObject]@{ ItemsSource = $null }
        $gridParameterizedTemplates = [PSCustomObject]@{ ItemsSource = $null }
        $script:Collections = @(
            [PSCustomObject]@{ Name = 'All Servers'; CollectionID = 'MCM00001'; Comment = 'prod'; IsBuiltIn = $false; MemberCount = 4; DirectCount = 1; QueryCount = 0 },
            [PSCustomObject]@{ Name = 'All Systems'; CollectionID = 'SMS00001'; Comment = $null;  IsBuiltIn = $true;  MemberCount = 9; DirectCount = 0; QueryCount = 0 },
            [PSCustomObject]@{ Name = $null;         CollectionID = 'MCM00003'; Comment = 'Empty one'; IsBuiltIn = $false; MemberCount = 0; DirectCount = 0; QueryCount = 2 })
        $script:OperationalTemplates = @([PSCustomObject]@{ Name = 'Clients | All'; Category = 'Clients' }, [PSCustomObject]@{ Name = 'Disk low'; Category = 'Hardware' })
        $script:ParameterizedTemplates = @([PSCustomObject]@{ Name = 'Software by name'; Category = 'Software' })
    }

    It 'Collections view: text "<Text>" and filter "<Filter>" show <Expected>' -ForEach @(
        @{ Text = '';          Filter = 'All';                  Expected = 'MCM00001,SMS00001,MCM00003' }
        @{ Text = 'SERVERS';   Filter = 'All';                  Expected = 'MCM00001' }
        @{ Text = 'sms0';      Filter = 'All';                  Expected = 'SMS00001' }
        @{ Text = 'prod';      Filter = 'All';                  Expected = 'MCM00001' }
        @{ Text = '  empty ';  Filter = 'All';                  Expected = 'MCM00003' }
        @{ Text = 'zzz';       Filter = 'All';                  Expected = '' }
        @{ Text = '';          Filter = 'Built-in only';        Expected = 'SMS00001' }
        @{ Text = '';          Filter = 'Custom only';          Expected = 'MCM00001,MCM00003' }
        @{ Text = '';          Filter = 'Empty (zero members)'; Expected = 'MCM00003' }
        @{ Text = 'all';       Filter = 'Custom only';          Expected = 'MCM00001' }
    ) {
        $txtFilter.Text = $Text
        $cboCollectionFilter.SelectedItem = $Filter
        Update-Filter
        (@($gridCollections.ItemsSource) | ForEach-Object CollectionID) -join ',' | Should -Be $Expected
    }

    It 'treats the filter text <Text> as plain text' -ForEach @(@{ Text = '.*' }, @{ Text = '[' }, @{ Text = '(' }, @{ Text = '%' }, @{ Text = "'" }) {
        $txtFilter.Text = $Text
        { Update-Filter } | Should -Not -Throw
        @($gridCollections.ItemsSource).Count | Should -Be 0
    }

    It 'Templates view: text "<Text>" shows <Op> operational and <Pa> parameterized templates' -ForEach @(
        @{ Text = '';         Op = 2; Pa = 1 }
        @{ Text = 'clients';  Op = 1; Pa = 0 }
        @{ Text = 'SOFTWARE'; Op = 0; Pa = 1 }
        @{ Text = 'disk';     Op = 1; Pa = 0 }
        @{ Text = 'nothing';  Op = 0; Pa = 0 }
    ) {
        $script:ActiveView = 'Templates'
        $txtFilter.Text = $Text
        Update-Filter
        @($gridOperationalTemplates.ItemsSource).Count | Should -Be $Op
        @($gridParameterizedTemplates.ItemsSource).Count | Should -Be $Pa
    }

    It 'WQL Editor view: leaves the grids alone' {
        $script:ActiveView = 'WQL Editor'
        $gridCollections.ItemsSource = 'unchanged'
        Update-Filter
        $gridCollections.ItemsSource | Should -Be 'unchanged'
    }
}

Describe 'Update-StatusBarSummary' {
    BeforeEach {
        function Set-StatusText { param([string]$Text) $script:StatusText = $Text }
        $global:Prefs = @{ SiteCode = 'MCM'; SMSProvider = 'cm01' }
        $script:IsConnected = $false
        $script:CollectionInventoryLoaded = $false
        $script:RawCollections = @(); $script:Collections = @(); $script:OperationalTemplates = @(); $script:ParameterizedTemplates = @()
        $script:LastRefreshTime = $null
        $script:StatusText = $null
    }

    AfterAll { Remove-Variable -Name Prefs -Scope Global -ErrorAction SilentlyContinue }

    It 'says what to do when <Case>' -ForEach @(
        @{ Case = 'the site code is missing'; Site = ''; Provider = 'cm01'; Expected = 'Open Options to configure site code and SMS provider' }
        @{ Case = 'the provider is missing'; Site = 'MCM'; Provider = ''; Expected = 'Open Options to configure site code and SMS provider' }
        @{ Case = 'both are set'; Site = 'MCM'; Provider = 'cm01'; Expected = 'Connect to the site, or refresh to load the collection inventory.' }
    ) {
        $global:Prefs = @{ SiteCode = $Site; SMSProvider = $Provider }
        Update-StatusBarSummary
        $script:StatusText | Should -Be $Expected
    }

    It 'names the site when connected' {
        $script:IsConnected = $true
        Update-StatusBarSummary
        $script:StatusText | Should -Be 'Connected to MCM'
    }

    It 'shows the inventory, the templates and the refresh time' {
        $script:IsConnected = $true
        $script:CollectionInventoryLoaded = $true
        $script:RawCollections = @(1, 2, 3)
        $script:OperationalTemplates = @(1, 2); $script:ParameterizedTemplates = @(1)
        $script:LastRefreshTime = [datetime]'2026-11-06T09:08:07'
        Update-StatusBarSummary
        $script:StatusText | Should -Be 'Connected to MCM   |   3 collections   |   2 op templates   |   1 param templates   |   last refresh 09:08:07'
    }

    It 'counts the targeted collections when the inventory is not loaded' {
        $script:Collections = @(1, 2)
        Update-StatusBarSummary
        $script:StatusText | Should -Match '2 targeted collection\(s\)'
    }
}

Describe 'Get-ActiveExportInfo' {
    It 'exports the grid on the Collections view with its columns, and nothing elsewhere' {
        $gridCollections = [PSCustomObject]@{ ItemsSource = 'rows' }
        $script:ActiveView = 'Collections'
        $i = Get-ActiveExportInfo
        $i.Name | Should -Be 'Collections'
        $i.Rows | Should -Be 'rows'
        $i.Columns | Should -Be @('Name', 'CollectionID', 'MemberCount', 'LimitingCollection', 'RefreshType', 'Comment')
        $script:ActiveView = 'Templates'
        Get-ActiveExportInfo | Should -BeNullOrEmpty
    }
}

Describe 'Preferences file' {
    BeforeEach { $global:PrefsPath = Join-Path $TestDrive ('prefs-{0}.json' -f [guid]::NewGuid().ToString('N')) }
    AfterAll { Remove-Variable -Name PrefsPath -Scope Global -ErrorAction SilentlyContinue }

    It 'gives the defaults on first run' {
        $p = Get-CmPreferences
        $p.DarkMode | Should -BeTrue
        $p.SiteCode | Should -Be ''
        $p.SMSProvider | Should -Be ''
    }

    It 'gives the defaults for a <Case> file' -ForEach @(
        @{ Case = 'corrupt'; Content = '{ not json' }
        @{ Case = 'empty'; Content = '' }
        @{ Case = 'array'; Content = '[1,2,3]' }
    ) {
        Set-Content -LiteralPath $global:PrefsPath -Value $Content
        (Get-CmPreferences).SiteCode | Should -Be ''
    }

    It 'ignores a key it does not know' {
        Set-Content -LiteralPath $global:PrefsPath -Value '{ "SiteCode": "MCM", "Evil": "x" }'
        $p = Get-CmPreferences
        $p.SiteCode | Should -Be 'MCM'
        $p.ContainsKey('Evil') | Should -BeFalse
    }

    It 'round-trips the value <Value>' -ForEach @(@{ Value = 'MCM' }, @{ Value = 'Ünï-日本' }, @{ Value = 'a "q" \ b' }, @{ Value = ('x' * 5000) }) {
        Save-CmPreferences -Prefs @{ DarkMode = $false; SiteCode = $Value; SMSProvider = 'cm01' }
        $p = Get-CmPreferences
        $p.SiteCode | Should -Be $Value
        $p.DarkMode | Should -BeFalse
    }

    It 'does not throw when the file cannot be written' {
        $global:PrefsPath = Join-Path $TestDrive 'dir-in-the-way'
        New-Item -ItemType Directory -Path $global:PrefsPath | Out-Null
        { Save-CmPreferences -Prefs @{ DarkMode = $true; SiteCode = 'X'; SMSProvider = '' } } | Should -Not -Throw
    }
}

Describe 'Script contract' {
    It 'calls only commands that exist' {
        $own = @($script:Ast.FindAll({ param($x) $x -is [System.Management.Automation.Language.FunctionDefinitionAst] }, $true) | ForEach-Object Name)
        $called = @($script:Ast.FindAll({ param($x) $x -is [System.Management.Automation.Language.CommandAst] }, $true) |
                ForEach-Object { $_.GetCommandName() } | Where-Object { $_ } | Sort-Object -Unique)
        Import-Module (Join-Path $script:AppRoot 'Lib\SuiteCommon\SuiteCommon.psd1') -Global -DisableNameChecking -ErrorAction SilentlyContinue
        $missing = @($called | Where-Object { $_ -notin $own -and -not (Get-Command $_ -ErrorAction SilentlyContinue) })
        $missing | Where-Object { $_ -notmatch '^(Get|Set|New|Remove|Start|Invoke|Add|Copy)-CM' } | Should -BeNullOrEmpty
    }
}
