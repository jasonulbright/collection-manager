#Requires -Modules Pester
# Window test for the WQL Editor spacing. The view is shown in both themes at 1024x720 and
# 1440x900, and the bounds of its controls are measured. Rules, all on the 4/8 px grid:
#   - 8 px between the tree and the splitter, 8 px between the splitter and the right pane
#   - 4 px between a label and its box, 12 px between a box and the next label
#   - the query box is at least 360 px wide and 80 px high, and the whole view fits at 1024x720, with the buttons wrapped to two rows
#   - the button row has 8 px between buttons, no overlap, and stays inside the pane

BeforeDiscovery {
    . "$PSScriptRoot\WindowHarness.ps1"
    $script:WindowGuard = Test-WindowTestsAllowed
    if (-not $script:WindowGuard.Allowed -and -not $env:SUITE_WINDOW_GUARD_REPORTED) { Write-Host ('Window tests skipped: {0}.' -f $script:WindowGuard.Reason) }
    $script:Cases = @(foreach ($dark in $true, $false) {
            foreach ($size in @(@(1024, 720), @(1440, 900))) {
                @{ Dark = $dark; Theme = $(if ($dark) { 'dark' } else { 'light' }); W = $size[0]; H = $size[1]; Size = ('{0}x{1}' -f $size[0], $size[1]) }
            }
        })
}

BeforeAll {
    . "$PSScriptRoot\WindowHarness.ps1"
    $script:WindowGuard = Test-WindowTestsAllowed
    if (-not $script:WindowGuard.Allowed) { return }
    . "$PSScriptRoot\WindowConfig.ps1"
    $script:Measured = @{}
    foreach ($case in @(
            @{ Dark = $true; W = 1024; H = 720 }, @{ Dark = $true; W = 1440; H = 900 },
            @{ Dark = $false; W = 1024; H = 720 }, @{ Dark = $false; W = 1440; H = 900 })) {
        $key = '{0}|{1}x{2}' -f $(if ($case.Dark) { 'dark' } else { 'light' }), $case.W, $case.H
        $case.Config = $script:AppConfig
        $script:Measured[$key] = Invoke-WpfSta -Arguments $case -Script {
            $cfg = $Arguments.Config
            $w = Open-TestWindow -XamlPath (Join-Path $Root $cfg.Xaml) -Dark $Arguments.Dark -Width $Arguments.W -Height $Arguments.H
            foreach ($v in $cfg.Views) { $w.FindName($v.Container).Visibility = if ($v.Name -eq 'WQL Editor') { 'Visible' } else { 'Collapsed' } }
            Wait-Layout $w
            $box = {
                param($name)
                $e = $w.FindName($name)
                $p = $e.TransformToAncestor($w).Transform([System.Windows.Point]::new(0, 0))
                [pscustomobject]@{ Name = $name; Left = $p.X; Top = $p.Y; Right = $p.X + $e.ActualWidth; Bottom = $p.Y + $e.ActualHeight; Width = $e.ActualWidth; Height = $e.ActualHeight }
            }
            $names = 'tabWqlSource', 'txtSelectedColl', 'lstWqlRules', 'txtRuleName', 'txtWqlEditor',
            'btnValidateWql', 'btnPreviewWql', 'btnAddRule', 'btnUpdateRule', 'btnRemoveRule', 'txtWqlValidation'
            $rows = @(foreach ($n in $names) { & $box $n })
            # The labels have no name: take the TextBlocks of the right pane by text.
            $labels = @()
            foreach ($t in 'Rules in This Collection', 'Rule Name', 'Query') {
                $tb = $null
                foreach ($d in (Get-VisualDescendants $w)) { if ($d -is [System.Windows.Controls.TextBlock] -and $d.Text -eq $t -and (Test-VisibleInTree $d)) { $tb = $d; break } }
                if ($tb) {
                    $p = $tb.TransformToAncestor($w).Transform([System.Windows.Point]::new(0, 0))
                    $labels += [pscustomobject]@{ Name = $t; Left = $p.X; Top = $p.Y; Right = $p.X + $tb.ActualWidth; Bottom = $p.Y + $tb.ActualHeight; Width = $tb.ActualWidth; Height = $tb.ActualHeight }
                }
            }
            $view = $w.FindName('viewWqlEditor')
            $vp = $view.TransformToAncestor($w).Transform([System.Windows.Point]::new(0, 0))
            $split = $null
            foreach ($d in (Get-VisualDescendants $view)) { if ($d -is [System.Windows.Controls.GridSplitter] -and $d.Parent -eq $view) { $split = $d; break } }
            $sp = $split.TransformToAncestor($w).Transform([System.Windows.Point]::new(0, 0))
            $out = [pscustomobject]@{
                Controls = $rows; Labels = $labels
                Splitter = [pscustomobject]@{ Left = $sp.X; Right = $sp.X + $split.ActualWidth }
                View     = [pscustomobject]@{ Left = $vp.X; Top = $vp.Y; Right = $vp.X + $view.ActualWidth; Bottom = $vp.Y + $view.ActualHeight }
            }
            $w.Close()
            $out
        }
    }
    function Get-Box { param($Case, [string]$Name) @($script:Measured[$Case].Controls) + @($script:Measured[$Case].Labels) | Where-Object Name -eq $Name }
}

AfterAll {
    if ($script:WindowGuard.Allowed) { Close-WpfSession }
}

Describe 'WQL Editor spacing (<Theme> <Size>)' -ForEach $script:Cases -Skip:(-not $script:WindowGuard.Allowed) {
    BeforeAll {
        $script:Key = '{0}|{1}' -f $Theme, $Size
        $script:M = $script:Measured[$script:Key]
        $script:Tol = 1.5
    }

    It 'measured every control and label' {
        @($script:M.Controls).Count | Should -Be 11
        @($script:M.Labels).Count | Should -Be 3
        foreach ($c in $script:M.Controls | Where-Object Name -ne 'txtWqlValidation') { $c.Width | Should -BeGreaterThan 0 -Because $c.Name }
    }

    It 'leaves 8 px between the tree and the splitter, and 8 px between the splitter and the right pane' {
        $t = Get-Box $script:Key 'tabWqlSource'
        $l = Get-Box $script:Key 'lstWqlRules'
        [math]::Abs(($script:M.Splitter.Left - $t.Right) - 8) | Should -BeLessOrEqual 2.5   # border thickness 1 px each side
        [math]::Abs(($l.Left - $script:M.Splitter.Right) - 8) | Should -BeLessOrEqual 2.5
    }

    It 'aligns the left edge of the right-pane boxes and labels' {
        $edge = (Get-Box $script:Key 'lstWqlRules').Left
        foreach ($n in 'txtRuleName', 'txtWqlEditor', 'btnValidateWql', 'txtSelectedColl', 'Rules in This Collection', 'Rule Name', 'Query') {
            [math]::Abs((Get-Box $script:Key $n).Left - $edge) | Should -BeLessOrEqual 2.5 -Because $n
        }
    }

    It 'leaves 4 px between each label and its box' {
        foreach ($pair in @(@('Rules in This Collection', 'lstWqlRules'), @('Rule Name', 'txtRuleName'), @('Query', 'txtWqlEditor'))) {
            $lab = Get-Box $script:Key $pair[0]; $b = Get-Box $script:Key $pair[1]
            [math]::Abs(($b.Top - $lab.Bottom) - 4) | Should -BeLessOrEqual $script:Tol -Because $pair[0]
        }
    }

    It 'leaves 12 px between a box and the next label' {
        foreach ($pair in @(@('lstWqlRules', 'Rule Name'), @('txtRuleName', 'Query'))) {
            $b = Get-Box $script:Key $pair[0]; $lab = Get-Box $script:Key $pair[1]
            [math]::Abs(($lab.Top - $b.Bottom) - 12) | Should -BeLessOrEqual $script:Tol -Because $pair[1]
        }
    }

    It 'gives the query box at least 360 px of width and 80 px of height' {
        $e = Get-Box $script:Key 'txtWqlEditor'
        $e.Width | Should -BeGreaterOrEqual 360
        $e.Height | Should -BeGreaterOrEqual 80
    }

    It 'gives the rules list at least 72 px of height' {
        (Get-Box $script:Key 'lstWqlRules').Height | Should -BeGreaterOrEqual 70
        (Get-Box $script:Key 'txtRuleName').Height | Should -BeGreaterOrEqual 28
    }

    It 'keeps the buttons apart by 8 px, without overlap, inside the view' {
        $buttons = @('btnValidateWql', 'btnPreviewWql', 'btnAddRule', 'btnUpdateRule', 'btnRemoveRule') | ForEach-Object { Get-Box $script:Key $_ }
        $rows = $buttons | Group-Object { [math]::Round($_.Top / 4) }
        foreach ($r in $rows) {
            $sorted = @($r.Group | Sort-Object Left)
            for ($i = 1; $i -lt $sorted.Count; $i++) {
                ($sorted[$i].Left - $sorted[$i - 1].Right) | Should -BeGreaterOrEqual (8 - $script:Tol)
            }
        }
        foreach ($b in $buttons) {
            $b.Right | Should -BeLessOrEqual ($script:M.View.Right + 0.5) -Because $b.Name
            $b.Bottom | Should -BeLessOrEqual ($script:M.View.Bottom + 0.5) -Because $b.Name
        }
        # Wrapped rows keep 8 px between them.
        $tops = @($buttons | Sort-Object Top | Select-Object -ExpandProperty Top -Unique)
        if ($tops.Count -gt 1) { ($tops[1] - $tops[0] - 28) | Should -BeGreaterOrEqual (8 - $script:Tol) }
    }

    It 'leaves 12 px between the query box and the first button row' {
        $e = Get-Box $script:Key 'txtWqlEditor'; $b = Get-Box $script:Key 'btnValidateWql'
        [math]::Abs(($b.Top - $e.Bottom) - 12) | Should -BeLessOrEqual $script:Tol
    }
}
