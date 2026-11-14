#Requires -Modules Pester

<#
.SYNOPSIS
    Pester 5.x test for the WQL Editor layout (XAML only, no window is opened).
#>

BeforeAll {
    $script:xmlWindow = [xml](Get-Content -LiteralPath (Join-Path -Path (Split-Path -Path $PSScriptRoot -Parent) -ChildPath 'MainWindow.xaml') -Raw)
    $script:nsWpf = New-Object System.Xml.XmlNamespaceManager($script:xmlWindow.NameTable)
    $script:nsWpf.AddNamespace('w', 'http://schemas.microsoft.com/winfx/2006/xaml/presentation')
    $script:nsWpf.AddNamespace('x', 'http://schemas.microsoft.com/winfx/2006/xaml')
    $script:view = $script:xmlWindow.SelectSingleNode("//w:Grid[@x:Name='viewWqlEditor']", $script:nsWpf)
    $script:pane = $script:view.SelectSingleNode("w:Grid[@Grid.Column='2']", $script:nsWpf)
}

Describe 'WQL Editor layout (XAML)' {
    It 'finds the view' { $script:view | Should -Not -BeNullOrEmpty }

    It 'gives the right pane a minimum width and star-sized list and editor rows with a minimum height' {
        $cols = @($script:view.SelectSingleNode('w:Grid.ColumnDefinitions', $script:nsWpf).ChildNodes)
        [int]$cols[2].GetAttribute('MinWidth') | Should -BeGreaterOrEqual 360
        $rowsNode = $script:pane.SelectSingleNode('.//w:Grid.RowDefinitions', $script:nsWpf)
        $starRows = @($rowsNode.ChildNodes | Where-Object { $_.GetAttribute('Height') -match '\*$' })
        $starRows.Count | Should -BeGreaterOrEqual 2
        foreach ($r in $starRows) { [int]$r.GetAttribute('MinHeight') | Should -BeGreaterOrEqual 72 }
    }

    It 'puts every margin of the right pane on the 4 px grid' {
        foreach ($node in $script:pane.SelectNodes('.//*[@Margin]')) {
            foreach ($n in ($node.GetAttribute('Margin') -split '[,\s]+' | Where-Object { $_ })) {
                ([double]$n % 4) | Should -Be 0 -Because ('Margin="{0}" on {1}' -f $node.GetAttribute('Margin'), $node.LocalName)
            }
        }
    }

    It 'labels the query box' {
        $script:pane.SelectSingleNode(".//w:TextBlock[@Text='Query']", $script:nsWpf) | Should -Not -BeNullOrEmpty
    }
}
