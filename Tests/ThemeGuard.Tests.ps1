#Requires -Modules Pester
# Guard tests for the light and dark themes. They parse the files. They create no
# WPF control and open no window, so they run on Windows PowerShell 5.1 and on PowerShell 7.
#   a. A blue background always comes with a white foreground.
#   b. The light theme and the dark theme differ in colors only.
#   c. The UI strings, the README and the user guide hold none of the banned words.

BeforeDiscovery {
    $script:XamlSourceNames = @('MainWindow.xaml', 'start-collectionmanager.ps1')
}

BeforeAll {
    $script:Root = Split-Path -Path $PSScriptRoot -Parent
    $script:XamlSourceNames = @('MainWindow.xaml', 'start-collectionmanager.ps1')

    # The window XAML of the app: the main window file and every XAML here-string of the start script.
    function Get-XamlDocument {
        param([string]$Name)
        $path = Join-Path $script:Root $Name
        $text = [IO.File]::ReadAllText($path)
        $docs = New-Object System.Collections.Generic.List[object]
        if ($Name -like '*.xaml') {
            $x = New-Object System.Xml.XmlDocument
            $x.LoadXml($text)
            $docs.Add([pscustomobject]@{ Label = $Name; Xml = $x })
        }
        else {
            $i = 0
            foreach ($m in [regex]::Matches($text, "(?ms)@'\s*\r?\n(<Controls:MetroWindow.*?)\r?\n'@")) {
                $i++
                $x = New-Object System.Xml.XmlDocument
                $x.LoadXml($m.Groups[1].Value)
                $docs.Add([pscustomobject]@{ Label = ('{0} window {1}' -f $Name, $i); Xml = $x })
            }
        }
        , $docs.ToArray()
    }

    function Select-XamlNode {
        param($Node, [string]$Local)
        @($Node.SelectNodes(('descendant-or-self::*[local-name()="{0}"]' -f $Local)))
    }

    # True when the brush text names a blue or an accent color.
    function Test-BlueBrush {
        param([string]$Value)
        if ([string]::IsNullOrWhiteSpace($Value)) { return $false }
        if ($Value -match 'MahApps\.Brushes\.(Accent|AccentBase|Highlight|Selected|WindowTitle|TextBox\.Focus)' -and $Value -notmatch 'IdealForeground') { return $true }
        if ($Value -match '\b(Blue|DodgerBlue|RoyalBlue|SteelBlue|CornflowerBlue|DeepSkyBlue)\b') { return $true }
        if ($Value -match '#(?:[0-9A-Fa-f]{2})?([0-9A-Fa-f]{2})([0-9A-Fa-f]{2})([0-9A-Fa-f]{2})\b') {
            $r = [Convert]::ToInt32($Matches[1], 16); $g = [Convert]::ToInt32($Matches[2], 16); $b = [Convert]::ToInt32($Matches[3], 16)
            if ($b -ge 0x80 -and $b -gt ($r + 40) -and $b -ge $g) { return $true }
        }
        $false
    }

    function Test-WhiteBrush {
        param([string]$Value)
        if ([string]::IsNullOrWhiteSpace($Value)) { return $false }
        [bool]($Value -match 'IdealForeground|Brushes\.White|^\s*White\s*$|^\s*#(FF)?FFFFFF\s*$')
    }

    # The Background value of a Setter: an attribute or a nested SolidColorBrush.
    function Get-SetterValue {
        param($Setter)
        $v = $Setter.GetAttribute('Value')
        if ($v) { return $v }
        $c = $Setter.SelectSingleNode('descendant::*[@Color]')
        if ($c) { return $c.GetAttribute('Color') }
        ''
    }

    # Returns a list of "<where>: <what>" for every blue background without a white foreground.
    function Get-BlueWithoutWhite {
        param($Document)
        $found = New-Object System.Collections.Generic.List[string]
        $xml = $Document.Xml
        # A style, and each trigger of a style, that sets a blue Background or is based on an Accent style.
        foreach ($style in (Select-XamlNode $xml.DocumentElement 'Style')) {
            $key = $style.GetAttribute('Key', 'http://schemas.microsoft.com/winfx/2006/xaml')
            $setters = @($style.SelectNodes('*[local-name()="Setter"]'))
            $bg = @($setters | Where-Object { $_.GetAttribute('Property') -eq 'Background' })
            $fg = @($setters | Where-Object { $_.GetAttribute('Property') -eq 'Foreground' })
            $blue = ($style.GetAttribute('BasedOn') -match 'Accent') -or (@($bg | Where-Object { Test-BlueBrush (Get-SetterValue $_) }).Count -gt 0)
            if ($blue -and (@($fg | Where-Object { Test-WhiteBrush (Get-SetterValue $_) }).Count -eq 0)) {
                $found.Add(('{0}: style {1} has a blue background and no white foreground' -f $Document.Label, $key))
            }
            foreach ($trigger in @($style.SelectNodes('descendant::*[local-name()="Trigger" or local-name()="DataTrigger" or local-name()="MultiTrigger"]'))) {
                $tbg = @($trigger.SelectNodes('*[local-name()="Setter"]') | Where-Object { $_.GetAttribute('Property') -eq 'Background' -and (Test-BlueBrush (Get-SetterValue $_)) })
                $tfg = @($trigger.SelectNodes('*[local-name()="Setter"]') | Where-Object { $_.GetAttribute('Property') -eq 'Foreground' -and (Test-WhiteBrush (Get-SetterValue $_)) })
                if ($tbg.Count -gt 0 -and $tfg.Count -eq 0) {
                    $found.Add(('{0}: a trigger of style {1} has a blue background and no white foreground' -f $Document.Label, $key))
                }
            }
        }
        # A control with a blue Background attribute. Containers have no text.
        $noText = 'Border|Grid|StackPanel|DockPanel|WrapPanel|Canvas|Rectangle|Ellipse|Path|ScrollViewer|ResourceDictionary|Style|ColumnDefinition|RowDefinition|MetroWindow|Setter'
        foreach ($el in @($xml.SelectNodes('//*[@Background]'))) {
            if ($el.LocalName -match ('^(' + $noText + ')$')) { continue }
            if (-not (Test-BlueBrush $el.GetAttribute('Background'))) { continue }
            if (-not (Test-WhiteBrush $el.GetAttribute('Foreground'))) {
                $found.Add(('{0}: {1} {2} has a blue background and no white foreground' -f $Document.Label, $el.LocalName, $el.GetAttribute('Name', 'http://schemas.microsoft.com/winfx/2006/xaml')))
            }
        }
        , $found.ToArray()
    }

    function Get-ScriptAst {
        param([string]$Path)
        $tokens = $null; $errors = $null
        [System.Management.Automation.Language.Parser]::ParseFile($Path, [ref]$tokens, [ref]$errors)
    }
}

Describe 'Blue backgrounds carry white text' {
    It 'finds the XAML of <_>' -ForEach $script:XamlSourceNames {
        @(Get-XamlDocument -Name $_).Count | Should -BeGreaterThan 0
    }

    It 'has no blue background without a white foreground in <_>' -ForEach $script:XamlSourceNames {
        $problems = foreach ($doc in @(Get-XamlDocument -Name $_)) { Get-BlueWithoutWhite -Document $doc }
        @($problems) | Should -BeNullOrEmpty
    }

    It 'detects a blue style without white text (the check itself)' {
        $x = New-Object System.Xml.XmlDocument
        $x.LoadXml('<ResourceDictionary xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"><Style x:Key="Bad" TargetType="Button"><Setter Property="Background" Value="#0078D4"/></Style><Style x:Key="Good" TargetType="Button"><Setter Property="Background" Value="#0078D4"/><Setter Property="Foreground" Value="White"/></Style></ResourceDictionary>')
        $r = @(Get-BlueWithoutWhite -Document ([pscustomobject]@{ Label = 'sample'; Xml = $x }))
        $r.Count | Should -Be 1
        $r[0] | Should -Match 'Bad'
    }

    It 'sets white text on the accent selection of every grid row' {
        $x = (Get-XamlDocument -Name 'MainWindow.xaml')[0].Xml
        $row = Select-XamlNode $x.DocumentElement 'Style' | Where-Object { $_.GetAttribute('Key', 'http://schemas.microsoft.com/winfx/2006/xaml') -eq 'ViewDataGridRow' }
        $row | Should -Not -BeNullOrEmpty
        $cell = Select-XamlNode $x.DocumentElement 'Style' | Where-Object { $_.GetAttribute('Key', 'http://schemas.microsoft.com/winfx/2006/xaml') -eq 'GridCellStyle' }
        $cell.InnerXml | Should -Match 'IsSelected'
        $cell.InnerXml | Should -Match 'Foreground'
    }

    It 'vendors a SuiteCommon that paints white text on blue buttons' {
        $text = [IO.File]::ReadAllText((Join-Path $script:Root 'Lib\SuiteCommon\SuiteCommon.psm1'))
        $text | Should -Match 'OnBlue\s*=\s*\$bc\.ConvertFrom\(''#FFFFFF''\)'
        $text | Should -Match 'function Set-ThemeForeground'
        $text | Should -Match 'FieldTitleFontSize\s*=\s*14'
        $text | Should -Match 'function Find-FieldTitle'
    }

    It 'does not paint a blue background in the start script without a white foreground nearby' {
        $ast = Get-ScriptAst -Path (Join-Path $script:Root 'start-collectionmanager.ps1')
        $assign = $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.AssignmentStatementAst] -and $n.Left.Extent.Text -match '\.Background$' }, $true)
        $bad = foreach ($a in $assign) {
            if (Test-BlueBrush $a.Right.Extent.Text) {
                $scope = $a.Parent.Extent.Text
                if ($scope -notmatch 'Foreground') { 'line {0}: {1}' -f $a.Extent.StartLineNumber, $a.Extent.Text }
            }
        }
        @($bad) | Should -BeNullOrEmpty
    }
}

Describe 'Site code and SMS Provider titles and values' {
    It 'sets the title <Label> to 14 pt or more and bold' -ForEach @(
        @{ Label = 'Site Code'; Name = 'lblSiteCode' }
        @{ Label = 'SMS Provider FQDN'; Name = 'lblSmsProvider' }
    ) {
        $hits = foreach ($src in $script:XamlSourceNames) {
            foreach ($doc in @(Get-XamlDocument -Name $src)) {
                foreach ($el in @($doc.Xml.SelectNodes('//*[@*[local-name()="Name"]="' + $Name + '"]'))) { $el }
            }
        }
        @($hits).Count | Should -BeGreaterThan 0
        foreach ($el in $hits) {
            $el.GetAttribute('Text') | Should -Be $Label
            [double]$el.GetAttribute('FontSize') | Should -BeGreaterOrEqual 14
            $el.GetAttribute('FontWeight') | Should -Be 'Bold'
        }
    }

    It 'keeps the value box <_> in the body style: no bold, no size above 12' -ForEach @('txtSiteCode', 'txtSmsProvider') {
        $name = $_
        $hits = foreach ($src in $script:XamlSourceNames) {
            foreach ($doc in @(Get-XamlDocument -Name $src)) {
                foreach ($box in @($doc.Xml.SelectNodes('//*[@*[local-name()="Name"]="' + $name + '"]'))) { $box }
            }
        }
        @($hits).Count | Should -BeGreaterThan 0
        foreach ($box in $hits) {
            $box.GetAttribute('FontWeight') | Should -BeIn @('', 'Normal')
            if ($box.GetAttribute('FontSize')) { [double]$box.GetAttribute('FontSize') | Should -BeLessOrEqual 12 }
        }
    }
}

Describe 'Every button takes a shared themed style' {
    It 'gives every Button in <_> a Style that is based on a MahApps button style' -ForEach $script:XamlSourceNames {
        $problems = foreach ($doc in @(Get-XamlDocument -Name $_)) {
            $styles = @{}
            foreach ($s in (Select-XamlNode $doc.Xml.DocumentElement 'Style')) { $styles[$s.GetAttribute('Key', 'http://schemas.microsoft.com/winfx/2006/xaml')] = $s }
            foreach ($b in (Select-XamlNode $doc.Xml.DocumentElement 'Button')) {
                $id = $b.GetAttribute('Name', 'http://schemas.microsoft.com/winfx/2006/xaml')
                $ref = $b.GetAttribute('Style')
                if ($ref -notmatch '^\{StaticResource\s+([\w\.]+)\}$') { '{0}: button {1} has no Style' -f $doc.Label, $id; continue }
                $key = $Matches[1]
                if ($key -like 'MahApps.Styles.Button*') { continue }
                $seen = @{}
                while ($key -and $styles.ContainsKey($key) -and -not $seen.ContainsKey($key)) {
                    $seen[$key] = $true
                    $based = $styles[$key].GetAttribute('BasedOn')
                    $key = if ($based -match '^\{StaticResource\s+([\w\.]+)\}$') { $Matches[1] } else { '' }
                }
                if ($key -notlike 'MahApps.Styles.Button*') { '{0}: button {1} uses a style that is not based on a MahApps button style' -f $doc.Label, $id }
            }
        }
        @($problems) | Should -BeNullOrEmpty
    }

    It 'sets no hard-coded Background, Foreground or BorderBrush on a Button' -ForEach $script:XamlSourceNames {
        $problems = foreach ($doc in @(Get-XamlDocument -Name $_)) {
            foreach ($b in (Select-XamlNode $doc.Xml.DocumentElement 'Button')) {
                foreach ($attr in 'Background', 'Foreground', 'BorderBrush') {
                    $v = $b.GetAttribute($attr)
                    if ($v -and $v -notmatch 'DynamicResource') { '{0}: button {1} sets {2}="{3}"' -f $doc.Label, $b.GetAttribute('Name', 'http://schemas.microsoft.com/winfx/2006/xaml'), $attr, $v }
                }
            }
        }
        @($problems) | Should -BeNullOrEmpty
    }

    It 'detects a button without a style (the check itself)' {
        $x = New-Object System.Xml.XmlDocument
        $x.LoadXml('<Grid xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"><Button x:Name="bare" Content="x"/></Grid>')
        $b = Select-XamlNode $x.DocumentElement 'Button'
        $b.GetAttribute('Style') | Should -BeNullOrEmpty
    }
}
Describe 'Light and dark differ in colors only' {
    It 'sets no font property differently in the dark and the light branch of <_>' -ForEach @('start-collectionmanager.ps1', 'Lib\SuiteCommon\SuiteCommon.psm1') {
        $ast = Get-ScriptAst -Path (Join-Path $script:Root $_)
        $fontPattern = '\.(FontSize|FontWeight|FontStyle|FontFamily|TextDecorations)\s*='
        $themeIf = $ast.FindAll({ param($n)
                $n -is [System.Management.Automation.Language.IfStatementAst] -and $n.Clauses[0].Item1.Extent.Text -match 'IsDark|isDark|DarkMode|IsOn'
            }, $true)
        $bad = foreach ($if in $themeIf) {
            $thenText = $if.Clauses[0].Item2.Extent.Text
            $elseText = if ($if.ElseClause) { $if.ElseClause.Extent.Text } else { '' }
            $thenFonts = @([regex]::Matches($thenText, $fontPattern + '[^\r\n]*') | ForEach-Object { $_.Value.Trim() } | Sort-Object)
            $elseFonts = @([regex]::Matches($elseText, $fontPattern + '[^\r\n]*') | ForEach-Object { $_.Value.Trim() } | Sort-Object)
            if (($thenFonts -join "`n") -ne ($elseFonts -join "`n")) { 'line {0}: font assignments differ' -f $if.Extent.StartLineNumber }
        }
        @($bad) | Should -BeNullOrEmpty
    }

    It 'binds no font setter in <_> to the theme' -ForEach $script:XamlSourceNames {
        $problems = foreach ($doc in @(Get-XamlDocument -Name $_)) {
            foreach ($trigger in @($doc.Xml.SelectNodes('//*[local-name()="Trigger" or local-name()="DataTrigger" or local-name()="MultiTrigger"]'))) {
                $text = $trigger.OuterXml
                if ($text -match 'Dark|Light|Theme(?!Background|Foreground)' -and $text -match 'Property="(FontSize|FontWeight|FontStyle|TextDecorations)"') {
                    '{0}: a trigger sets a font property for a theme' -f $doc.Label
                }
            }
        }
        @($problems) | Should -BeNullOrEmpty
    }

    It 'keeps FontSize, FontWeight, FontStyle and TextDecorations out of the light and dark theme dictionaries' {
        $themeFiles = @(Get-ChildItem -LiteralPath $script:Root -Recurse -File -Include '*.xaml' | Where-Object { $_.FullName -notlike '*\Tests\*' -and $_.Name -match 'Dark|Light' })
        $problems = foreach ($f in $themeFiles) {
            if ([IO.File]::ReadAllText($f.FullName) -match 'Font(Size|Weight|Style)|TextDecorations') { $f.Name }
        }
        @($problems) | Should -BeNullOrEmpty
    }
}

Describe 'Banned words' {
    BeforeAll {
        $script:Banned = '\b(easily|simply|seamless(ly)?|powerful|robust|streamline[ds]?|quickly|just|effortless(ly)?|intuitive(ly)?|handy)\b|helps you|allows you to'

        function Get-UiString {
            $out = New-Object System.Collections.Generic.List[object]
            foreach ($src in $script:XamlSourceNames) {
                foreach ($doc in @(Get-XamlDocument -Name $src)) {
                    foreach ($attr in @($doc.Xml.SelectNodes('//@*[local-name()="ToolTip" or local-name()="Content" or local-name()="Text" or local-name()="Header" or local-name()="Title" or local-name()="Watermark" or local-name()="HelpText"]'))) {
                        $out.Add([pscustomobject]@{ Where = $doc.Label; Text = $attr.Value })
                    }
                }
            }
            foreach ($file in 'start-collectionmanager.ps1', 'Module\CollectionManagerCommon.psm1') {
                $ast = Get-ScriptAst -Path (Join-Path $script:Root $file)
                $strings = $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.StringConstantExpressionAst] -or $n -is [System.Management.Automation.Language.ExpandableStringExpressionAst] }, $true)
                foreach ($s in $strings) {
                    # An attribute such as SuppressMessage holds a justification for the analyzer, not a UI string.
                    $p = $s.Parent; $inAttribute = $false
                    while ($p) { if ($p -is [System.Management.Automation.Language.AttributeAst]) { $inAttribute = $true; break }; $p = $p.Parent }
                    if ($inAttribute) { continue }
                    $out.Add([pscustomobject]@{ Where = ('{0}:{1}' -f $file, $s.Extent.StartLineNumber); Text = [string]$s.Value })
                }
            }
            foreach ($f in @(Get-ChildItem -LiteralPath (Join-Path $script:Root 'Templates') -Filter '*.json' -File)) {
                $out.Add([pscustomobject]@{ Where = ('Templates\{0}' -f $f.Name); Text = [IO.File]::ReadAllText($f.FullName) })
            }
            , $out.ToArray()
        }
    }

    It 'finds no banned word in a tooltip or UI string' {
        $hits = foreach ($s in @(Get-UiString)) {
            if ($s.Text -match $script:Banned) { '{0}: {1}' -f $s.Where, $Matches[0] }
        }
        @($hits) | Should -BeNullOrEmpty
    }

    It 'finds no banned word in the README or the user guide' {
        $files = @(Get-Item -LiteralPath (Join-Path $script:Root 'README.md'))
        $files += @(Get-ChildItem -LiteralPath (Join-Path $script:Root 'docs') -Recurse -File -Filter '*.md')
        $hits = foreach ($f in $files) {
            $n = 0
            foreach ($line in [IO.File]::ReadAllLines($f.FullName)) {
                $n++
                if ($line -match $script:Banned) { '{0}:{1}: {2}' -f $f.Name, $n, $Matches[0] }
            }
        }
        @($hits) | Should -BeNullOrEmpty
    }

    It 'detects a banned word (the check itself)' {
        'This helps you to quickly pick a site.' | Should -Match $script:Banned
        'Select a collection.' | Should -Not -Match $script:Banned
    }
}
