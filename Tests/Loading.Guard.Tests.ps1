#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.1.0' }
# Guard for the member loads (Q-47). It parses the start script and opens no window.
#   - The load of the direct members and the Edit Membership list make no Configuration Manager call on the UI thread:
#     the code of both functions has none, and each starts its work through Start-SuiteBgWork.
#   - A new load, a close of the dialog and the end of the program stop the load (Stop-SuiteBgWork -Discard).

BeforeAll {
    $script:Root = Split-Path -Parent $PSScriptRoot
    $script:ScriptPath = Join-Path $script:Root 'start-collectionmanager.ps1'
    $script:Ast = [System.Management.Automation.Language.Parser]::ParseFile($script:ScriptPath, [ref]$null, [ref]$null)
    function Get-FunctionText {
        param([string]$Name)
        $fn = $script:Ast.Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $Name }, $true)
        if (-not $fn) { throw "Function $Name not found" }
        return $fn.Extent.Text
    }
    $script:CmCall = 'Add-DirectMember|Remove-DirectMember|Get-CollectionMembers|Get-CMCollectionMember|Get-CM[A-Z]\w*|Invoke-CM\w+|Get-CimInstance|Get-WmiObject|Connect-CMSite|Set-Location\s+\$'
}

Describe 'Member loads stay off the UI thread' {
    It 'makes no Configuration Manager call in <Name>' -ForEach @(@{ Name = 'Invoke-LoadDirectMembers' }, @{ Name = 'Show-EditMembershipDialog' }) {
        $text = Get-FunctionText -Name $Name
        # The reload script block of the dialog and the callbacks hold no call; the worker is Read-SuiteCollectionMember.
        $text | Should -Not -Match $script:CmCall
        $text | Should -Match 'Start-SuiteBgWork'
        $text | Should -Match 'Read-SuiteCollectionMember'
    }

    It 'stops the running load when a new one starts, the dialog closes or the program ends' {
        (Get-FunctionText -Name 'Invoke-LoadDirectMembers') | Should -Match 'Stop-SuiteBgWork -Handle \$script:MemberLoad -Discard'
        $dialog = Get-FunctionText -Name 'Show-EditMembershipDialog'
        $dialog | Should -Match 'Stop-SuiteBgWork -Handle \$dlgState\.Load -Discard'
        $dialog | Should -Match 'Add_Closed'
        $text = Get-Content -LiteralPath $script:ScriptPath -Raw
        $text | Should -Match 'Close-SuiteBgRunspace -Runspace \$script:MemberRunspace'
    }

    It 'fills the grid from a collection that grows, not from one array' {
        $text = Get-FunctionText -Name 'Invoke-LoadDirectMembers'
        $text | Should -Match 'ObservableCollection'
        $text | Should -Not -Match 'ItemsSource\s*=\s*@\('
    }

    It 'uses the vendored loading helpers' {
        $suite = Get-Content -LiteralPath (Join-Path $script:Root 'Lib\SuiteCommon\SuiteCommon.psm1') -Raw
        foreach ($name in 'Invoke-SuiteCMRead', 'Get-SuiteCollectionMemberChunk', 'Start-SuiteBgWork', 'Read-SuiteCollectionMember') { $suite | Should -Match "function $name" }
        $suite | Should -Match '\[switch\]\$Discard'
    }
}

Describe 'The load can be cancelled' {
    It 'has a Cancel button on the progress overlay and the overlay takes the click' {
        $script:XamlText = Get-Content -LiteralPath (Join-Path $script:Root 'MainWindow.xaml') -Raw
        $script:XamlText | Should -Match '<Button x:Name="btnCancelWork" Content="Cancel" Style="\{StaticResource SquareInlineButtonStyle\}"'
        $script:XamlText | Should -Match '<Border x:Name="progressOverlay"\s+Visibility="Collapsed"\s+IsHitTestVisible="True"'
    }

    It 'stops the work, drops the runspace and gives the window back when the button is clicked' {
        $text = Get-Content -LiteralPath $script:ScriptPath -Raw
        $text | Should -Match '\$btnCancelWork\s+= \$window\.FindName\(''btnCancelWork''\)'
        $text | Should -Match '\$btnCancelWork\.Add_Click\(\{ Invoke-CancelWork \}\)'
        $fn = Get-FunctionText -Name 'Invoke-CancelWork'
        $fn | Should -Match "\`$script:BgState\['Cancel'\] = \`$true"
        $fn | Should -Match 'Dispose-BgWork'
        $fn | Should -Match 'Close-SuiteBgRunspace -Runspace \$script:BgRunspace'
        $fn | Should -Match '\$progressOverlay\.Visibility = \[System\.Windows\.Visibility\]::Collapsed'
        $fn | Should -Not -Match $script:CmCall
    }
}

Describe 'Add and remove of direct members stay off the UI thread' {
    It 'starts both writes through Start-DirectMemberWrite and calls no write in the dialog itself' {
        $dialog = Get-FunctionText -Name 'Show-EditMembershipDialog'
        $dialog | Should -Match "Start-DirectMemberWrite -DlgState \`$dlgState -Kind 'Add'"
        $dialog | Should -Match "Start-DirectMemberWrite -DlgState \`$dlgState -Kind 'Remove'"
        $dialog | Should -Not -Match 'Add-DirectMember|Remove-DirectMember'
    }

    It 'runs the write in a runspace of its own with Start-SuiteBgWork, connects there and checks the cancel flag' {
        $helper = Get-FunctionText -Name 'Start-DirectMemberWrite'
        $helper | Should -Match 'New-SuiteBgRunspace'
        $helper | Should -Match 'Start-SuiteBgWork -Runspace \$DlgState\.WriteRunspace'
        $helper | Should -Match 'Connect-CMSite -SiteCode \$Site -SMSProvider \$Provider'
        $helper | Should -Match "\`$State\['Cancel'\]"
        $helper | Should -Match 'IsEnabled = \$false'
    }

    It 'turns the buttons on again and reloads the list when the write ends, and stops the write when the dialog closes' {
        $dialog = Get-FunctionText -Name 'Show-EditMembershipDialog'
        ([regex]::Matches($dialog, 'IsEnabled = \$true')).Count | Should -BeGreaterOrEqual 2
        $dialog | Should -Match '& \$ctx\.Reload'
        $dialog | Should -Match 'Stop-SuiteBgWork -Handle \$dlgState\.Write -Discard'
        $dialog | Should -Match 'Close-SuiteBgRunspace -Runspace \$dlgState\.WriteRunspace'
    }
}

Describe 'The collection list shows at once and the rules follow in their own runspace (Q-80)' {
    BeforeAll {
        $script:Text = Get-Content -LiteralPath $script:ScriptPath -Raw
        $script:Refresh = [regex]::Match($script:Text, '(?s)function Invoke-Refresh \{.*?\r?\n\}\r?\n').Value
        $script:RulesLoad = [regex]::Match($script:Text, '(?s)function Start-RulesLoad \{.*?\r?\n\}\r?\n').Value
    }

    It 'reads the list with the bulk function: no per-collection cmdlet and no rules parsing in the refresh worker' {
        $script:Refresh | Should -Match 'Get-AllDeviceCollections'
        $script:Refresh | Should -Not -Match 'Get-CMDeviceCollection'
        $script:Refresh | Should -Not -Match '\.CollectionRules'
        $script:Refresh | Should -Not -Match 'Parsing rules'
    }

    It 'starts the rules pass after the list is applied, and a new refresh or the end of the program stops it' {
        $script:Refresh.LastIndexOf('Start-RulesLoad') | Should -BeGreaterThan $script:Refresh.IndexOf('$gridCollections.ItemsSource = $script:Collections')
        $script:Refresh | Should -Match 'Stop-RulesLoad'
        $script:Text | Should -Match '(?s)\$window\.Add_Closing\(\{.*?Stop-RulesLoad'
    }

    It 'reads the rules of one collection at a time in a runspace of its own, stops on Cancel and counts failures' {
        $script:RulesLoad | Should -Match 'New-SuiteBgRunspace'
        $script:RulesLoad | Should -Match 'Get-CollectionRuleSet -CollectionId \$id'
        $script:RulesLoad | Should -Match 'if \(\$State\[''Cancel''\]\) \{ break \}'
        $script:RulesLoad | Should -Match '\$State\.FailedN\[\$W\] = \[int\]\$State\.FailedN\[\$W\] \+ 1'
        $script:RulesLoad | Should -Match 'Loading rules: \{1\} of \{2\}'
    }

    It 'shows blank rule counts until a collection is read, and has a time limit with a clear message' {
        $script:Refresh | Should -Match '\$directCount = if \(\$idx\) \{ @\(\$idx\.Direct\)\.Count \}\s+else \{ \$null \}'
        $script:Refresh | Should -Match 'Loading the collection list took longer than'
        $script:Text | Should -Match 'function Get-LoadTimeoutSeconds'
    }
}
