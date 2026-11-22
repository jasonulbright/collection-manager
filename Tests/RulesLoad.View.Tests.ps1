#Requires -Modules Pester
# The rules of a collection reach the screen. The rules pass reads each collection with
# Get-CollectionRuleSet, stores the result in the rules index, and the window copies it into the
# rule grids, the rule counts and the WQL Editor rule list. These tests run that whole path
# with real WPF controls and read what the user would see: the number of rows in each control.
#
# The site fake behaves like the real site: a read with the lazy option returns the collection
# with an EMPTY CollectionRules property (checked on the lab). A test that only checks the call
# shape passed while the window showed no rules. These tests check the rows.
#
# The script cannot be dot-sourced (it opens the window), so each function and event handler is
# cut out of its syntax tree and defined here. Controls are real WPF controls that are not shown.

BeforeAll {
    . "$PSScriptRoot\TestHelpers.ps1"
    Import-AppModule
    Add-Type -AssemblyName PresentationFramework
    $script:Stubs = New-CmStub -Name Invoke-CMWmiQuery
    Mock Write-Log {} -ModuleName CollectionManagerCommon

    $script:ShellPath = Join-Path $script:AppRoot 'start-collectionmanager.ps1'
    $script:Ast = [System.Management.Automation.Language.Parser]::ParseFile($script:ShellPath, [ref]$null, [ref]$null)
    foreach ($n in 'Update-RuleCountsFromIndex', 'Apply-WqlCollectionSelection', 'Update-Filter', 'Get-CollectionFilterValue', 'Test-CollectionFilterMatch') {
        $def = $script:Ast.Find({ param($x) $x -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $x.Name -eq $n }, $true)
        if (-not $def) { throw "function $n not found in the script" }
        . ([scriptblock]::Create($def.Extent.Text))
    }
    # The handler of the collection grid: $gridCollections.Add_SelectionChanged({ ... })
    $call = $script:Ast.Find({
            param($x)
            $x -is [System.Management.Automation.Language.InvokeMemberExpressionAst] -and $x.Member.Value -eq 'Add_SelectionChanged' -and
            $x.Expression.Extent.Text -eq '$gridCollections'
        }, $true)
    if (-not $call) { throw 'the SelectionChanged handler of gridCollections was not found in the script' }
    $script:SelectionHandler = $call.Arguments[0].ScriptBlock.GetScriptBlock()

    # What the site holds, by collection id.
    function New-SiteRule {
        param([string]$Kind, [string]$Name, $Value)
        switch ($Kind) {
            'Direct'  { [PSCustomObject]@{ SmsProviderObjectPath = 'SMS_CollectionRuleDirect.ResourceID=16777220'; RuleName = $Name; ResourceID = [int]$Value } }
            'Query'   { [PSCustomObject]@{ SmsProviderObjectPath = 'SMS_CollectionRuleQuery.RuleName=x'; RuleName = $Name; QueryExpression = $Value } }
            'Include' { [PSCustomObject]@{ SmsProviderObjectPath = 'SMS_CollectionRuleIncludeCollection.x'; RuleName = $Name; IncludeCollectionID = $Value } }
            'Exclude' { [PSCustomObject]@{ SmsProviderObjectPath = 'SMS_CollectionRuleExcludeCollection.x'; RuleName = $Name; ExcludeCollectionID = $Value } }
        }
    }
    $global:FxSite = @{
        'MCM00001' = @(
            (New-SiteRule Direct 'PC1' 16777220), (New-SiteRule Direct 'PC2' 16777221), (New-SiteRule Direct 'PC3' 16777222),
            (New-SiteRule Query 'Servers' 'select * from SMS_R_System where OperatingSystemNameandVersion like ''%Server%'''),
            (New-SiteRule Query 'Laptops' 'select * from SMS_R_System where Name like ''LT%'''),
            (New-SiteRule Include 'Inc A' 'MCM00002'),
            (New-SiteRule Exclude 'Exc A' 'MCM00003'))
        'MCM00002' = @()
    }
    # The fake site: a lazy read gives no rules, as the real provider does.
    Mock Invoke-CMWmiQuery {
        $id = [regex]::Match([string]$Query, "CollectionID = '([^']+)'").Groups[1].Value
        $rules = if ($Option -eq 'Lazy') { @() } else { @($global:FxSite[$id]) }
        [PSCustomObject]@{ CollectionID = $id; CollectionRules = $rules }
    } -ModuleName CollectionManagerCommon
}

AfterAll {
    Remove-CmStub -Name $script:Stubs
    Remove-Variable -Name FxSite -Scope Global -ErrorAction SilentlyContinue
}

Describe 'The rules pass fills the rule controls' {
    BeforeEach {
        # The rules pass: one Get-CollectionRuleSet call per collection, the result stored in the index by id.
        $script:CollectionRulesIndex = [hashtable]::Synchronized(@{})
        foreach ($id in 'MCM00001', 'MCM00002') { $script:CollectionRulesIndex[$id] = Get-CollectionRuleSet -CollectionId $id }

        $script:ActiveView = 'Collections'
        $script:Collections = @(
            [PSCustomObject]@{ Name = 'All Servers'; CollectionID = 'MCM00001'; MemberCount = 4; LimitingCollection = 'All Systems'; RefreshType = 'Manual'; Comment = ''; IsBuiltIn = $false; DirectCount = $null; QueryCount = $null }
            [PSCustomObject]@{ Name = 'Empty'; CollectionID = 'MCM00002'; MemberCount = 0; LimitingCollection = 'All Systems'; RefreshType = 'Manual'; Comment = ''; IsBuiltIn = $false; DirectCount = $null; QueryCount = $null })
        $script:WqlSelectedCollection = $null

        $txtFilter = New-Object System.Windows.Controls.TextBox
        $cboCollectionFilter = New-Object System.Windows.Controls.ComboBox
        $gridCollections = New-Object System.Windows.Controls.DataGrid
        $gridQueryRules = New-Object System.Windows.Controls.DataGrid
        $gridIncludeRules = New-Object System.Windows.Controls.DataGrid
        $gridExcludeRules = New-Object System.Windows.Controls.DataGrid
        $gridDirectMembers = New-Object System.Windows.Controls.DataGrid
        $txtCollProperties = New-Object System.Windows.Controls.TextBox
        $lstWqlRules = New-Object System.Windows.Controls.ListBox
        $txtSelectedColl = New-Object System.Windows.Controls.TextBlock
        $txtRuleName = New-Object System.Windows.Controls.TextBox
        $txtWqlEditor = New-Object System.Windows.Controls.TextBox
        $txtWqlValidation = New-Object System.Windows.Controls.TextBox
        $gridCollections.ItemsSource = $script:Collections
    }

    It 'reads the rules of a collection with a faithful lazy-empty site: the index holds all four kinds' {
        $idx = $script:CollectionRulesIndex['MCM00001']
        @($idx.Direct).Count | Should -Be 3
        @($idx.Query).Count | Should -Be 2
        @($idx.Include).Count | Should -Be 1
        @($idx.Exclude).Count | Should -Be 1
    }

    It 'shows the rules of the selected collection in the three rule grids when the pass ends' {
        $gridCollections.SelectedItem = $script:Collections[0]
        Update-RuleCountsFromIndex -Final
        $gridQueryRules.Items.Count | Should -Be 2
        $gridIncludeRules.Items.Count | Should -Be 1
        $gridExcludeRules.Items.Count | Should -Be 1
        $gridQueryRules.Items[0].RuleName | Should -Be 'Servers'
        $gridQueryRules.Items[0].QueryExpression | Should -Match 'Server'
        $gridIncludeRules.Items[0].IncludeCollectionID | Should -Be 'MCM00002'
        $gridExcludeRules.Items[0].ExcludeCollectionID | Should -Be 'MCM00003'
    }

    It 'puts the rule counts into the collection rows that the grid shows' {
        Update-RuleCountsFromIndex -Final
        $gridCollections.Items.Count | Should -Be 2
        $gridCollections.Items[0].DirectCount | Should -Be 3
        $gridCollections.Items[0].QueryCount | Should -Be 2
        $gridCollections.Items[1].QueryCount | Should -Be 0
    }

    It 'fills the rule list of the WQL Editor when the rules arrive after the collection was picked' {
        $script:ActiveView = 'WQL Editor'
        Apply-WqlCollectionSelection -Collection $script:Collections[0]
        $lstWqlRules.Items.Count | Should -Be 2 -Because 'the rules are in the index'
        @($lstWqlRules.Items) | Should -Be @('Servers', 'Laptops')
        $txtSelectedColl.Text | Should -Match 'MCM00001'
    }

    It 'fills the rule list of the WQL Editor late: a pick before the pass ends shows the rules when the pass ends' {
        $script:ActiveView = 'WQL Editor'
        $full = $script:CollectionRulesIndex
        $script:CollectionRulesIndex = [hashtable]::Synchronized(@{})
        Apply-WqlCollectionSelection -Collection $script:Collections[0]
        $lstWqlRules.Items.Count | Should -Be 0
        $script:CollectionRulesIndex = $full
        Update-RuleCountsFromIndex -Final
        $lstWqlRules.Items.Count | Should -Be 2
    }

    It 'shows the rules of a collection that the user selects after the pass ended' {
        Update-RuleCountsFromIndex -Final
        $gridCollections.SelectedItem = $script:Collections[0]
        & $script:SelectionHandler
        $gridQueryRules.Items.Count | Should -Be 2
        $gridIncludeRules.Items.Count | Should -Be 1
        $gridExcludeRules.Items.Count | Should -Be 1
        $txtCollProperties.Text | Should -Match 'All Servers'
    }

    It 'shows no rule rows for a collection without rules, and does not fail' {
        $gridCollections.SelectedItem = $script:Collections[1]
        & $script:SelectionHandler
        $gridQueryRules.Items.Count | Should -Be 0
        $gridIncludeRules.Items.Count | Should -Be 0
        $gridExcludeRules.Items.Count | Should -Be 0
    }
}
