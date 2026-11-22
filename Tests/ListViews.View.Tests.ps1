#Requires -Modules Pester
# The saved site queries list and the template lists reach the screen. Each test runs the load
# function of the script with real WPF controls (not shown) and reads the number of rows that
# the control holds. A test that checks only that a control exists, or that a function was
# called, passes while the list is empty.
#
# The script cannot be dot-sourced (it opens the window), so each function is cut out of its
# syntax tree and defined here.

BeforeAll {
    . "$PSScriptRoot\TestHelpers.ps1"
    Import-AppModule
    Add-Type -AssemblyName PresentationFramework
    $script:Stubs = New-CmStub -Name Get-CMQuery
    Mock Write-Log {} -ModuleName CollectionManagerCommon

    $script:ShellPath = Join-Path $script:AppRoot 'start-collectionmanager.ps1'
    $ast = [System.Management.Automation.Language.Parser]::ParseFile($script:ShellPath, [ref]$null, [ref]$null)
    foreach ($n in 'Sync-SiteSavedQueryCache', 'Update-SiteSavedQueryList', 'Invoke-LoadTemplates') {
        $def = $ast.Find({ param($x) $x -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $x.Name -eq $n }, $true)
        if (-not $def) { throw "function $n not found in the script" }
        # $PSScriptRoot in the script is the app folder, not the Tests folder.
        $text = $def.Extent.Text.Replace('$PSScriptRoot', ("'{0}'" -f $script:AppRoot))
        . ([scriptblock]::Create($text))
    }
    function Add-LogLine { param([string]$Text) $script:LogLines += $Text }
    function Update-StatusBarSummary { }
}

AfterAll {
    Remove-CmStub -Name $script:Stubs
}

Describe 'Saved site queries fill the list' {
    BeforeEach {
        $script:LogLines = @()
        $script:IsConnected = $true
        $script:SelectedSiteSavedQueryId = ''
        $script:SiteSavedQueries = @()
        $txtSavedQueryFilter = New-Object System.Windows.Controls.TextBox
        $gridSavedQueries = New-Object System.Windows.Controls.DataGrid
        $txtWqlValidation = New-Object System.Windows.Controls.TextBox
        Mock Get-CMQuery {
            [PSCustomObject]@{ QueryID = 'MCM00003'; Name = 'Windows 11 devices'; Expression = 'select * from SMS_R_System'; Comments = 'win11'; TargetClassName = 'SMS_R_System'; LimitToCollectionID = 'SMS00001' }
            [PSCustomObject]@{ QueryID = 'MCM00002'; Name = 'Servers'; Expression = 'select * from SMS_R_System'; Comments = ''; TargetClassName = 'SMS_R_System'; LimitToCollectionID = '' }
            [PSCustomObject]@{ QueryID = 'MCM00001'; Name = 'Laptops'; Expression = 'select * from SMS_R_System'; Comments = 'chassis'; TargetClassName = 'SMS_R_System'; LimitToCollectionID = '' }
        } -ModuleName CollectionManagerCommon
    }

    It 'puts every saved query of the site into the grid, in name order, with the name and id the user reads' {
        Sync-SiteSavedQueryCache | Should -BeTrue
        $gridSavedQueries.Items.Count | Should -Be 3
        (@($gridSavedQueries.Items) | ForEach-Object Name) -join ',' | Should -Be 'Laptops,Servers,Windows 11 devices'
        $gridSavedQueries.Items[0].QueryID | Should -Be 'MCM00001'
        $txtWqlValidation.Text | Should -Be 'Loaded 3 saved site queries.'
    }

    It 'shows only the queries that match the filter text' {
        $txtSavedQueryFilter.Text = 'serv'
        $ok = Sync-SiteSavedQueryCache
        $ok | Should -BeTrue -Because $txtWqlValidation.Text
        $gridSavedQueries.Items.Count | Should -Be 1
        $gridSavedQueries.Items[0].Name | Should -Be 'Servers'
    }

    It 'shows the one saved query of a site that has exactly one (a single object is not a list)' {
        Mock Get-CMQuery { [PSCustomObject]@{ QueryID = 'MCM00001'; Name = 'Laptops'; Expression = 'select 1'; Comments = ''; TargetClassName = 'SMS_R_System'; LimitToCollectionID = '' } } -ModuleName CollectionManagerCommon
        $ok = Sync-SiteSavedQueryCache
        $ok | Should -BeTrue -Because $txtWqlValidation.Text
        $gridSavedQueries.Items.Count | Should -Be 1
        $gridSavedQueries.Items[0].Name | Should -Be 'Laptops'
    }

    It 'shows an empty grid and a message when the site has no saved query' {
        Mock Get-CMQuery { } -ModuleName CollectionManagerCommon
        Sync-SiteSavedQueryCache | Should -BeTrue
        $gridSavedQueries.Items.Count | Should -Be 0
        $txtWqlValidation.Text | Should -Be 'Loaded 0 saved site queries.'
    }

    It 'says why the grid is empty when the read fails' {
        Mock Get-CMQuery { throw 'the provider did not answer' } -ModuleName CollectionManagerCommon
        Sync-SiteSavedQueryCache | Should -BeFalse
        $gridSavedQueries.Items.Count | Should -Be 0
        $txtWqlValidation.Text | Should -Match 'did not answer'
    }
}

Describe 'The shipped templates fill the template grids' {
    BeforeEach {
        $script:LogLines = @()
        $gridOperationalTemplates = New-Object System.Windows.Controls.DataGrid
        $gridParameterizedTemplates = New-Object System.Windows.Controls.DataGrid
    }

    It 'loads every template of the two shipped files into its grid' {
        $op = Get-Content -LiteralPath (Join-Path $script:AppRoot 'Templates\operational-collections.json') -Raw | ConvertFrom-Json
        $pa = Get-Content -LiteralPath (Join-Path $script:AppRoot 'Templates\parameterized-templates.json') -Raw | ConvertFrom-Json
        $opCount = @($op).Count
        $paCount = @($pa).Count
        $opCount | Should -BeGreaterThan 0
        $paCount | Should -BeGreaterThan 0

        Invoke-LoadTemplates
        $gridOperationalTemplates.Items.Count | Should -Be $opCount
        $gridParameterizedTemplates.Items.Count | Should -Be $paCount
        foreach ($t in @($gridOperationalTemplates.Items) + @($gridParameterizedTemplates.Items)) {
            [string]$t.Name | Should -Not -BeNullOrEmpty
        }
        @($gridParameterizedTemplates.Items | Where-Object { $null -eq $_.ParameterCount }).Count | Should -Be 0
    }
}
