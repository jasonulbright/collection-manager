#Requires -Modules Pester
# Window test for the chrome and the buttons of the main window and of every dialog, in both themes.
# Every dialog is built from its XAML text, given the theme the way the script gives it (Set-DialogTheme),
# and shown. Rules:
#   - the dialog title bar brush (active and inactive) is the title bar brush of the main window
#   - the title bar text of the dialog is readable (4.5:1) on that brush
#   - every button has bold text (SemiBold or heavier, 600; a normal weight is a defect)
#   - no button has a border thicker than 1 px, and no button border has the text color of the theme

BeforeDiscovery {
    . "$PSScriptRoot\WindowHarness.ps1"
    $script:WindowGuard = Test-WindowTestsAllowed
    if (-not $script:WindowGuard.Allowed -and -not $env:SUITE_WINDOW_GUARD_REPORTED) { Write-Host ('Window tests skipped: {0}.' -f $script:WindowGuard.Reason) }
    $script:ThemeCases = @(@{ Dark = $true; Theme = 'dark' }, @{ Dark = $false; Theme = 'light' })
}

BeforeAll {
    . "$PSScriptRoot\WindowHarness.ps1"
    $script:WindowGuard = Test-WindowTestsAllowed
    if (-not $script:WindowGuard.Allowed) { return }
    . "$PSScriptRoot\WindowConfig.ps1"
    $script:Root = Split-Path -Parent $PSScriptRoot
    $ast = [System.Management.Automation.Language.Parser]::ParseFile((Join-Path $script:Root $script:AppConfig.Shell), [ref]$null, [ref]$null)
    $dialogXaml = @($ast.FindAll({
                param($n)
                $n -is [System.Management.Automation.Language.StringConstantExpressionAst] -and $n.Value -match '^\s*<Controls:MetroWindow'
            }, $true) | ForEach-Object { $_.Value })
    $styleAst = $ast.Find({ param($n) $n -is [System.Management.Automation.Language.AssignmentStatementAst] -and $n.Left.Extent.Text -eq '$script:DialogButtonStyleXaml' }, $true)
    if ($styleAst) {
        $styleXaml = [string]$styleAst.Right.Expression.Value
        $dialogXaml = @($dialogXaml | ForEach-Object { $_.Replace('<!--DIALOG_BUTTON_STYLES-->', $styleXaml) })
    }
    $script:Measured = @{}
    foreach ($dark in $true, $false) {
        $key = if ($dark) { 'dark' } else { 'light' }
        $script:Measured[$key] = Invoke-WpfSta -Arguments @{ Dark = $dark; Config = $script:AppConfig; Dialogs = $dialogXaml } -Script {
            $cfg = $Arguments.Config
            $main = Open-TestWindow -XamlPath (Join-Path $Root $cfg.Xaml) -Dark $Arguments.Dark
            Initialize-ShellTheme -Window $main -Dark $Arguments.Dark -Views $cfg.Views -OptionsButtonName $cfg.OptionsButton -LogLabelName $cfg.LogLabel -ActiveView $cfg.Views[0].Name
            $hex = { param($b) if ($b -is [System.Windows.Media.SolidColorBrush]) { $b.Color.ToString() } else { '' } }
            $rel = {
                param($c)
                $f = { param($v) $v = $v / 255.0; if ($v -le 0.03928) { $v / 12.92 } else { [math]::Pow(($v + 0.055) / 1.055, 2.4) } }
                0.2126 * (& $f $c.R) + 0.7152 * (& $f $c.G) + 0.0722 * (& $f $c.B)
            }
            $ratio = {
                param($a, $b)
                $l1 = & $rel $a; $l2 = & $rel $b
                if ($l1 -lt $l2) { $t = $l1; $l1 = $l2; $l2 = $t }
                ($l1 + 0.05) / ($l2 + 0.05)
            }
            $buttonRows = {
                param($win, $label)
                $rows = @()
                foreach ($d in (Get-VisualDescendants $win)) {
                    if ($d -is [System.Windows.Controls.Primitives.ButtonBase] -and $d -isnot [System.Windows.Controls.Primitives.ToggleButton] -and (Test-VisibleInTree $d)) {
                        # Skip the window command buttons (minimize, maximize, close) and the template buttons of other controls.
                        if ($d.GetType().Name -ne 'Button') { continue }
                        if ($d.TemplatedParent -and $d.TemplatedParent.GetType().Name -ne 'Button') { continue }
                        $content = if ($d.Content -is [string]) { $d.Content } else { $d.Name }
                        if (-not $content) { continue }
                        $fgBrush = $d.Foreground
                        $rows += [pscustomobject]@{
                            Window = $label; Name = $d.Name; Text = [string]$content; Weight = $d.FontWeight.ToOpenTypeWeight()
                            BorderThickness = [math]::Max([math]::Max($d.BorderThickness.Left, $d.BorderThickness.Right), [math]::Max($d.BorderThickness.Top, $d.BorderThickness.Bottom))
                            BorderColor = (& $hex $d.BorderBrush); ForegroundColor = (& $hex $fgBrush)
                        }
                    }
                }
                $rows
            }
            $themeText = & $hex $main.FindResource('MahApps.Brushes.ThemeForeground')
            $mainRows = @(& $buttonRows $main 'main')
            $out = [pscustomobject]@{
                ThemeText = $themeText
                MainTitle = (& $hex $main.WindowTitleBrush); MainTitleInactive = (& $hex $main.NonActiveWindowTitleBrush)
                Main = $mainRows; Dialogs = @()
            }
            $i = 0
            foreach ($xaml in $Arguments.Dialogs) {
                $i++
                [xml]$x = $xaml
                $d = [System.Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader $x))
                $d.WindowStartupLocation = 'Manual'
                $d.Left = $(if ($env:SUITE_VISIBLE_WINDOWS -eq '1') { 120 } else { -32000 }); $d.Top = $(if ($env:SUITE_VISIBLE_WINDOWS -eq '1') { 120 } else { -32000 })
                $d.ShowInTaskbar = $false
                Set-DialogTheme -Dialog $d -IsDark $Arguments.Dark
                $d.Show(); Wait-Layout $d
                $title = $d.Title
                $tbText = if ($d.TitleForeground -is [System.Windows.Media.SolidColorBrush]) { $d.TitleForeground.Color } else { [System.Windows.Media.Colors]::White }
                $act = $d.WindowTitleBrush
                $out.Dialogs += [pscustomobject]@{
                    Title = $title; Index = $i
                    Active = (& $hex $act); Inactive = (& $hex $d.NonActiveWindowTitleBrush)
                    TitleRatio = $(if ($act -is [System.Windows.Media.SolidColorBrush]) { [math]::Round((& $ratio $tbText $act.Color), 2) } else { 0 })
                    Buttons = @(& $buttonRows $d $title)
                }
                $d.Close()
            }
            $main.Close()
            $out
        }
    }
}

AfterAll {
    if ($script:WindowGuard.Allowed) { Close-WpfSession }
}

Describe 'Chrome <Theme>' -ForEach $script:ThemeCases -Skip:(-not $script:WindowGuard.Allowed) {
    BeforeAll { $script:M = $script:Measured[$Theme] }

    It 'finds every dialog of the script' {
        $funcs = @([regex]::Matches((Get-Content -LiteralPath (Join-Path $script:Root $script:AppConfig.Shell) -Raw), '(?m)^function Show-\w*Dialog\b'))
        $funcs.Count | Should -BeGreaterThan 0
        @($script:M.Dialogs).Count | Should -BeGreaterOrEqual $funcs.Count
    }

    It 'gives every dialog the title bar brush of the main window' {
        $bad = @($script:M.Dialogs | Where-Object { $_.Active -ne $script:M.MainTitle -or $_.Inactive -ne $script:M.MainTitleInactive } |
                ForEach-Object { '{0}: active {1} inactive {2}, main {3} {4}' -f $_.Title, $_.Active, $_.Inactive, $script:M.MainTitle, $script:M.MainTitleInactive })
        $bad | Should -BeNullOrEmpty
    }

    It 'keeps the title text of every dialog readable at 4.5:1' {
        $bad = @($script:M.Dialogs | Where-Object { $_.TitleRatio -lt 4.5 } | ForEach-Object { '{0}: {1}' -f $_.Title, $_.TitleRatio })
        $bad | Should -BeNullOrEmpty
    }

    It 'shows bold text on every button of the main window and of every dialog' {
        $all = @($script:M.Main) + @($script:M.Dialogs | ForEach-Object { $_.Buttons })
        $all.Count | Should -BeGreaterThan 5
        $bad = @($all | Where-Object { $_.Weight -lt 600 } | ForEach-Object { '{0} "{1}" weight {2}' -f $_.Window, $_.Text, $_.Weight })
        $bad | Should -BeNullOrEmpty
    }

    It 'gives no button a border thicker than 1 px or in the text color of the theme' {
        $all = @($script:M.Main) + @($script:M.Dialogs | ForEach-Object { $_.Buttons })
        $bad = @($all | Where-Object { $_.BorderThickness -gt 1 -or ($_.BorderThickness -gt 0 -and $_.BorderColor -eq $script:M.ThemeText) } |
                ForEach-Object { '{0} "{1}" border {2}px {3}' -f $_.Window, $_.Text, $_.BorderThickness, $_.BorderColor })
        $bad | Should -BeNullOrEmpty
    }
}
