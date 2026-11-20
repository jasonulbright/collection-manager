#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.1.0' }
# Guard tests for control tooltips. They parse the XAML and the XAML here-strings of the script.
# They create no WPF control and open no window, so they run on Windows PowerShell 5.1 and on
# PowerShell 7.
#
# Rule: every control in the audited list has a ToolTip that is not empty, is ASCII, ends in a
# period and is not longer than 200 characters.
#
# A control that is not in the list needs no tooltip: OK, Cancel, Close and the other dialog buttons
# whose label says it all, read-only detail and log boxes, grid columns, and the other controls of
# the audit. To audit a new control, add a row to $script:Audited.

BeforeDiscovery {
    # File: path under the repository root. Name: the x:Name of the control, or Tab:<header text>
    # for a TabItem that has no name.
    $script:Audited = @(
        @{ File = 'MainWindow.xaml'; Name = 'txtFilter' }
        @{ File = 'MainWindow.xaml'; Name = 'Tab:Collections' }
        @{ File = 'MainWindow.xaml'; Name = 'txtWqlCollectionId' }
        @{ File = 'MainWindow.xaml'; Name = 'txtWqlTreeFilter' }
        @{ File = 'MainWindow.xaml'; Name = 'tabSiteQueries' }
        @{ File = 'MainWindow.xaml'; Name = 'txtSavedQueryFilter' }
        @{ File = 'MainWindow.xaml'; Name = 'txtSavedQueryName' }
        @{ File = 'MainWindow.xaml'; Name = 'txtSavedQueryComment' }
        @{ File = 'MainWindow.xaml'; Name = 'txtRuleName' }
        @{ File = 'MainWindow.xaml'; Name = 'txtWqlEditor' }
        @{ File = 'MainWindow.xaml'; Name = 'Tab:Operational' }
        @{ File = 'MainWindow.xaml'; Name = 'Tab:Parameterized' }
        @{ File = 'MainWindow.xaml'; Name = 'btnCancelWork' }
        @{ File = 'start-collectionmanager.ps1'; Name = 'txtName' }
        @{ File = 'start-collectionmanager.ps1'; Name = 'txtComment' }
        @{ File = 'start-collectionmanager.ps1'; Name = 'cboRefresh' }
        @{ File = 'start-collectionmanager.ps1'; Name = 'txtNewName' }
        @{ File = 'start-collectionmanager.ps1'; Name = 'txtAddDevice' }
        @{ File = 'start-collectionmanager.ps1'; Name = 'btnAddDevice' }
        @{ File = 'start-collectionmanager.ps1'; Name = 'btnRemoveSel' }
        @{ File = 'start-collectionmanager.ps1'; Name = 'btnReloadMembers' }
        @{ File = 'start-collectionmanager.ps1'; Name = 'txtRuleName' }
        @{ File = 'start-collectionmanager.ps1'; Name = 'btnCatConnection' }
        @{ File = 'start-collectionmanager.ps1'; Name = 'btnCatAbout' }
        @{ File = 'start-collectionmanager.ps1'; Name = 'txtSiteCode' }
        @{ File = 'start-collectionmanager.ps1'; Name = 'txtSmsProvider' }
    )
}

BeforeAll {
    $script:Root = Split-Path -Path $PSScriptRoot -Parent
    $script:XamlNs = 'http://schemas.microsoft.com/winfx/2006/xaml'

    # The XAML documents of one file: the file itself for a .xaml file, each XAML here-string for a script.
    function Get-XamlDocument {
        param([string]$File)
        $path = Join-Path $script:Root $File
        $text = [IO.File]::ReadAllText($path)
        $chunks = New-Object System.Collections.Generic.List[string]
        if ($File -like '*.xaml') { $chunks.Add($text) }
        else {
            foreach ($m in [regex]::Matches($text, "(?ms)@(?:'|`")[ \t]*\r?\n(\s*<.*?)\r?\n(?:'|`")@")) { $chunks.Add($m.Groups[1].Value) }
        }
        $docs = New-Object System.Collections.Generic.List[object]
        foreach ($c in $chunks) {
            $x = New-Object System.Xml.XmlDocument
            $err = ''
            try { $x.LoadXml($c) }
            catch {
                $wrapped = '<Wrap xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml" xmlns:Controls="http://metro.mahapps.com/winfx/xaml/controls">' + ($c -replace '^\s*<\?xml[^>]*\?>', '') + '</Wrap>'
                try { $x.LoadXml($wrapped) } catch { $err = $_.Exception.Message }
            }
            $docs.Add([pscustomobject]@{ Xml = $x; ParseError = $err })
        }
        $docs.ToArray()
    }

    # The elements of a file that carry the name. Tab:<text> finds the TabItem whose header text is <text>.
    function Get-NamedElement {
        param([string]$File, [string]$Name)
        $found = New-Object System.Collections.Generic.List[object]
        foreach ($doc in @(Get-XamlDocument -File $File)) {
            if ($doc.ParseError) { continue }
            if ($Name -like 'Tab:*') {
                $header = $Name.Substring(4)
                foreach ($tab in @($doc.Xml.SelectNodes('//*[local-name()="TabItem"]'))) {
                    if (@($tab.SelectNodes('./*[local-name()="TabItem.Header"]/*[local-name()="TextBlock"][@Text="' + $header + '"]')).Count -gt 0) { $found.Add($tab) }
                }
            }
            else {
                foreach ($el in @($doc.Xml.SelectNodes('//*[@*[local-name()="Name"]="' + $Name + '"]'))) {
                    if ($el.GetAttribute('Name', $script:XamlNs) -eq $Name) { $found.Add($el) }
                }
            }
        }
        $found.ToArray()
    }
}

Describe 'Tooltips' {
    It 'parses the XAML of every audited file' {
        foreach ($file in @('MainWindow.xaml', 'start-collectionmanager.ps1')) {
            $docs = @(Get-XamlDocument -File $file)
            $docs.Count | Should -BeGreaterThan 0 -Because ('{0} holds XAML' -f $file)
            foreach ($d in $docs) { $d.ParseError | Should -BeNullOrEmpty -Because ('the XAML of {0} parses' -f $file) }
        }
    }

    It 'has a tooltip on <Name> in <File>' -ForEach $script:Audited {
        $elements = @(Get-NamedElement -File $File -Name $Name)
        $elements.Count | Should -BeGreaterThan 0 -Because ('{0} exists in {1}' -f $Name, $File)
        foreach ($el in $elements) {
            $tip = $el.GetAttribute('ToolTip')
            $tip | Should -Not -BeNullOrEmpty -Because ('{0} needs a tooltip' -f $Name)
            $tip | Should -Match '^[\x20-\x7E]+$' -Because ('the tooltip of {0} is ASCII' -f $Name)
            $tip | Should -Match '\.$' -Because ('the tooltip of {0} ends in a period' -f $Name)
            $tip.Length | Should -BeLessOrEqual 200 -Because ('the tooltip of {0} is short' -f $Name)
        }
    }
}
