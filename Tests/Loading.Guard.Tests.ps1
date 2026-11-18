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
    $script:CmCall = 'Get-CollectionMembers|Get-CMCollectionMember|Get-CM[A-Z]\w*|Invoke-CM\w+|Get-CimInstance|Get-WmiObject|Connect-CMSite|Set-Location\s+\$'
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