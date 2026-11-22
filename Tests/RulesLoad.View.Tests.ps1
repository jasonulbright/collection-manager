#Requires -Modules Pester
# The rules of a collection reach the screen. The rules pass reads each collection with
# Get-CollectionRuleSet, stores the result in the rules index, and the window copies it into the
# rule grids, the rule counts and the WQL Editor rule list. These tests run that whole path
# with real WPF controls and read what the user would see: the number of rows in each control.
#
# The site fake behaves like the real site: a read with the lazy or the fast option returns the collection
# with an EMPTY CollectionRules property (lazy: checked on the lab; fast: skips lazy properties). A test that only checks the call
# shape passed while the window showed no rules. These tests check the rows.
#
# The script cannot be dot-sourced (it opens the window), so each function and event handler is
# cut out of its syntax tree and defined here. Controls are real WPF controls that are not shown.

BeforeAll {
    . "$PSScriptRoot\TestHelpers.ps1"
    Import-AppModule
    Add-Type -AssemblyName PresentationFramework
    # WPF controls need an STA thread. Windows PowerShell and PowerShell 7 start in STA; a harness that starts MTA must add -STA.
    if ([Threading.Thread]::CurrentThread.GetApartmentState() -ne 'STA') { throw 'These tests create WPF controls and need an STA thread. Start the host with -STA.' }
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
    # The fake site: a lazy or fast read gives no rules, as the real provider does.
    Mock Invoke-CMWmiQuery {
        if ([string]$Query -match 'CollectionType = 2') { return @($global:FxList) }
        $id = [regex]::Match([string]$Query, "CollectionID = '([^']+)'").Groups[1].Value
        $rules = if ($Option -in 'Lazy', 'Fast') { @() } else { @($global:FxSite[$id]) }
        [PSCustomObject]@{ CollectionID = $id; CollectionRules = $rules }
    } -ModuleName CollectionManagerCommon
}

AfterAll {
    Remove-CmStub -Name $script:Stubs
    Remove-Variable -Name FxSite, FxList, Prefs -Scope Global -ErrorAction SilentlyContinue
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

Describe 'The collection list reaches the grid' {
    BeforeAll {
        # The timer handler that ends the background refresh and fills the collection grid.
        $calls = $script:Ast.FindAll({
                param($x)
                $x -is [System.Management.Automation.Language.InvokeMemberExpressionAst] -and $x.Member.Value -eq 'Add_Tick' -and
                $x.Expression.Extent.Text -eq '$script:BgTimer' -and $x.Extent.Text.Contains('Loading the rules of each collection in the background')
            }, $true)
        if (@($calls).Count -ne 1) { throw "expected one refresh timer handler, found $(@($calls).Count)" }
        $script:RefreshTick = $calls[0].Arguments[0].ScriptBlock.GetScriptBlock()
        $def = $script:Ast.Find({ param($x) $x -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $x.Name -eq 'Get-CollectionTypeGlyph' }, $true)
        . ([scriptblock]::Create($def.Extent.Text))
        function Add-LogLine { param([string]$Text) $script:LogLines += $Text }
        function Set-StatusText { param([string]$Text) $script:StatusText = $Text }
        function Update-WqlTreeFromState { param([string]$Needle) }
        function Update-StatusBarSummary { }
        function Sync-FooterSite { }
        function Start-RulesLoad { $script:RulesLoadStarted++ }
        Set-Item -Path 'function:script:Connect-CMSite' -Value { param($SiteCode, $SMSProvider) $true }
        Set-Item -Path 'function:script:Write-Log' -Value { param($Message, $Level, [switch]$Quiet) }
        # The site: the list query answers with plain columns, as the real provider does.
        $global:FxList = @(
            [PSCustomObject]@{ CollectionID = 'SMS00001'; Name = 'All Systems'; MemberCount = 12; RefreshType = 4; LimitToCollectionName = ''; LimitToCollectionID = ''; Comment = 'Built in' }
            [PSCustomObject]@{ CollectionID = 'MCM00001'; Name = 'All Servers'; MemberCount = 4; RefreshType = 2; LimitToCollectionName = 'All Systems'; LimitToCollectionID = 'SMS00001'; Comment = 'prod' }
            [PSCustomObject]@{ CollectionID = 'MCM00002'; Name = 'Empty'; MemberCount = 0; RefreshType = 1; LimitToCollectionName = 'All Systems'; LimitToCollectionID = 'SMS00001'; Comment = '' })
    }

    BeforeEach {
        $script:LogLines = @(); $script:StatusText = ''; $script:RulesLoadStarted = 0
        $global:Prefs = @{ SiteCode = 'MCM'; SMSProvider = 'cm01' }
        $script:ActiveView = 'Collections'
        $script:IsConnected = $false
        $script:LoadStarted = Get-Date
        $script:WqlSelectedCollection = $null
        $script:Collections = @(); $script:RawCollections = @()
        $script:CollectionRulesIndex = [hashtable]::Synchronized(@{})
        $txtFilter = New-Object System.Windows.Controls.TextBox
        $cboCollectionFilter = New-Object System.Windows.Controls.ComboBox
        $gridCollections = New-Object System.Windows.Controls.DataGrid
        $gridCollections.CanUserAddRows = $false
        $gridQueryRules = New-Object System.Windows.Controls.DataGrid
        $gridIncludeRules = New-Object System.Windows.Controls.DataGrid
        $gridExcludeRules = New-Object System.Windows.Controls.DataGrid
        $lstWqlRules = New-Object System.Windows.Controls.ListBox
        $txtSelectedColl = New-Object System.Windows.Controls.TextBlock
        $txtRuleName = New-Object System.Windows.Controls.TextBox
        $txtWqlEditor = New-Object System.Windows.Controls.TextBox
        $txtWqlValidation = New-Object System.Windows.Controls.TextBox
        $txtWqlTreeFilter = New-Object System.Windows.Controls.TextBox
        $txtProgressStep = New-Object System.Windows.Controls.TextBox
        $btnConnect = New-Object System.Windows.Controls.Button
        $btnRefresh = [PSCustomObject]@{ IsEnabled = $false }
        $progressOverlay = [PSCustomObject]@{ Visibility = [System.Windows.Visibility]::Visible }
        $script:BgTimer = [PSCustomObject]@{} | Add-Member -MemberType ScriptMethod -Name Stop -Value { } -PassThru
        $script:BgPowerShell = [PSCustomObject]@{} | Add-Member -MemberType ScriptMethod -Name EndInvoke -Value { param($h) } -PassThru | Add-Member -MemberType ScriptMethod -Name Dispose -Value { } -PassThru
        $script:BgInvokeHandle = $null
        # The refresh worker: the bulk list function of the module, then the folder id of each collection.
        $listed = @(Get-AllDeviceCollections | ForEach-Object { $_ | Add-Member -MemberType NoteProperty -Name FolderID -Value 0 -Force -PassThru })
        $script:BgState = @{ Done = $true; ErrorMsg = $null; Connected = $true; Step = 'x'; Result = [PSCustomObject]@{ Collections = $listed; Folders = @() } }
    }

    It 'shows every collection of the site in the grid, with the cells the user reads' {
        & $script:RefreshTick
        $gridCollections.Items.Count | Should -Be 3
        $gridCollections.Items[1].Name | Should -Be 'All Servers'
        $gridCollections.Items[1].CollectionID | Should -Be 'MCM00001'
        $gridCollections.Items[1].MemberCount | Should -Be 4
        $gridCollections.Items[1].RefreshType | Should -Be 'Periodic'
        $gridCollections.Items[1].LimitingCollection | Should -Be 'All Systems'
        $gridCollections.Items[1].Comment | Should -Be 'prod'
        $gridCollections.Items[0].IsBuiltIn | Should -BeTrue
        $gridCollections.Items[1].IsBuiltIn | Should -BeFalse
        $gridCollections.Items[2].RefreshType | Should -Be 'Manual'
    }

    It 'starts the rules pass, gives the window back and says how many collections loaded' {
        & $script:RefreshTick
        $script:RulesLoadStarted | Should -Be 1
        [string]$progressOverlay.Visibility | Should -Be 'Collapsed'
        $btnRefresh.IsEnabled | Should -BeTrue
        ($script:LogLines -join ' | ') | Should -Match 'Loaded 3 collections'
    }

    It 'shows the rule counts and the rules of the selected collection when the rules pass ends after the list' {
        & $script:RefreshTick
        foreach ($id in 'SMS00001', 'MCM00001', 'MCM00002') { $script:CollectionRulesIndex[$id] = Get-CollectionRuleSet -CollectionId $id }
        $gridCollections.SelectedItem = $script:Collections[1]
        Update-RuleCountsFromIndex -Final
        $gridCollections.Items[1].DirectCount | Should -Be 3
        $gridCollections.Items[1].QueryCount | Should -Be 2
        $gridQueryRules.Items.Count | Should -Be 2
    }

    It 'shows the one collection of a site that has one (a single object is not a list)' {
        $script:BgState.Result.Collections = $script:BgState.Result.Collections[1]
        & $script:RefreshTick
        $gridCollections.Items.Count | Should -Be 1
        $gridCollections.Items[0].Name | Should -Be 'All Servers'
    }

    It 'shows an empty grid for a site without device collections' {
        $script:BgState.Result.Collections = @()
        & $script:RefreshTick
        $gridCollections.Items.Count | Should -Be 0
    }

    It 'keeps the grid as it was and says the refresh failed when the background step failed' {
        $gridCollections.ItemsSource = @([PSCustomObject]@{ Name = 'KEEP' })
        $script:BgState.ErrorMsg = 'boom'
        $script:BgState.Connected = $false
        & $script:RefreshTick
        $gridCollections.Items.Count | Should -Be 1
        $script:StatusText | Should -Be 'Refresh failed.'
        $script:RulesLoadStarted | Should -Be 0
        $btnRefresh.IsEnabled | Should -BeTrue
    }
}

Describe 'The WQL Editor tree shows the folders and the collections' {
    BeforeAll {
        $def = $script:Ast.Find({ param($x) $x -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $x.Name -eq 'Update-WqlTreeFromState' }, $true)
        . ([scriptblock]::Create($def.Extent.Text))
    }

    BeforeEach {
        $script:ActiveView = 'WQL Editor'
        $treeWqlCollections = New-Object System.Windows.Controls.TreeView
        $script:Folders = @(
            [PSCustomObject]@{ FolderID = 10; ParentID = 0; Name = 'Servers' }
            [PSCustomObject]@{ FolderID = 11; ParentID = 10; Name = 'Prod' })
        $script:Collections = @(
            [PSCustomObject]@{ Name = 'All Laptops'; CollectionID = 'MCM00001'; MemberCount = 5; FolderID = 0 }
            [PSCustomObject]@{ Name = 'Web Servers'; CollectionID = 'MCM00002'; MemberCount = 2; FolderID = 11 }
            [PSCustomObject]@{ Name = 'DB Servers'; CollectionID = 'MCM00003'; MemberCount = 1; FolderID = 10 })
    }

    It 'shows the folder tree with each collection under its folder, and counts the collections' {
        Update-WqlTreeFromState | Should -Be 3
        $treeWqlCollections.Items.Count | Should -Be 2
        $treeWqlCollections.Items[0].Header | Should -Be '[+] Servers'
        $treeWqlCollections.Items[1].Header | Should -Be 'All Laptops  (MCM00001, 5 members)'
        $servers = $treeWqlCollections.Items[0]
        $servers.Items.Count | Should -Be 2
        $servers.Items[0].Header | Should -Be '[+] Prod'
        $servers.Items[0].Items.Count | Should -Be 1
        $servers.Items[0].Items[0].Header | Should -Be 'Web Servers  (MCM00002, 2 members)'
    }

    It 'shows only the folders that hold a match when the filter text is set' {
        Update-WqlTreeFromState -Needle 'web' | Should -Be 1
        $treeWqlCollections.Items.Count | Should -Be 1
        $treeWqlCollections.Items[0].Items[0].Items[0].Header | Should -Match 'Web Servers'
    }

    It 'shows one hint line instead of the tree for a site with more than 1000 collections and no filter' {
        $script:Collections = @(1..1001 | ForEach-Object { [PSCustomObject]@{ Name = "C$_"; CollectionID = "MCM$_"; MemberCount = 0; FolderID = 0 } })
        Update-WqlTreeFromState | Should -Be 0
        $treeWqlCollections.Items.Count | Should -Be 1
        $treeWqlCollections.Items[0].Header | Should -Match '1001 collections loaded'
    }

    It 'shows an empty tree for a site without collections' {
        $script:Collections = @()
        Update-WqlTreeFromState
        $treeWqlCollections.Items.Count | Should -Be 0
    }
}
