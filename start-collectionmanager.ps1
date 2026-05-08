<#
.SYNOPSIS
    MahApps.Metro WPF shell for the MECM Collection Manager.

.DESCRIPTION
    Sidebar navigation across three views (Collections, WQL Editor, Templates),
    inline action bar (Refresh, filter, status filter, exports), modal dialogs
    for Options and per-collection actions, log drawer, and status bar. Status
    conveyed via glyph at ThemeForeground (no row-color coloring). Site code
    and SMS provider configured from the Options sidebar button.

    Requirements:
      - PowerShell 5.1
      - .NET Framework 4.7.2+
      - MahApps.Metro DLLs in .\Lib\
      - CollectionManagerCommon module under .\Module\
      - ConfigurationManager console (provides Get-CMDeviceCollection family)

.NOTES
    ScriptName : start-collectionmanager.ps1
    Version    : 1.0.1
    Updated    : 2026-05-04
#>

[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidGlobalVars', '', Justification='Per feedback_ps_wpf_handler_rules.md and PS51-WPF-001..003: flat-.ps1 GetNewClosure strips $script: scope. $global: survives closure scope-strip and keeps shared mutable state reachable from closure-captured handlers.')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter', '', Justification='WPF event handler scriptblocks bind positional sender/args ($s, $e). The sender is required to fulfill the signature even when the handler body does not read it.')]
[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'

# =============================================================================
# Startup transcript (best-effort).
# =============================================================================
$__txDir = Join-Path $PSScriptRoot 'Logs'
try {
    if (-not (Test-Path -LiteralPath $__txDir)) { New-Item -ItemType Directory -Path $__txDir -Force | Out-Null }
    $__tx = Join-Path $__txDir ('CollectionManager-startup-{0}.txt' -f (Get-Date -Format 'yyyyMMdd-HHmmss'))
    Start-Transcript -LiteralPath $__tx -Force | Out-Null
} catch { $null = $_ }

# =============================================================================
# STA guard. WPF requires STA. PS51-WPF-009.
# =============================================================================
if ([System.Threading.Thread]::CurrentThread.GetApartmentState() -ne 'STA') {
    $psExe = (Get-Process -Id $PID).Path
    $fwd   = @('-NoProfile','-ExecutionPolicy','Bypass','-STA','-File',$PSCommandPath)
    Start-Process -FilePath $psExe -ArgumentList $fwd | Out-Null
    try { Stop-Transcript | Out-Null } catch { $null = $_ }
    exit 0
}

# =============================================================================
# Assemblies.
# =============================================================================
Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Windows.Forms

$libDir = Join-Path $PSScriptRoot 'Lib'
if (-not (Test-Path -LiteralPath $libDir)) {
    throw "Lib/ directory not found at: $libDir. Re-extract the release zip."
}

Get-ChildItem -LiteralPath $libDir -File -ErrorAction SilentlyContinue |
    Unblock-File -ErrorAction SilentlyContinue

[void][System.Reflection.Assembly]::LoadFrom((Join-Path $libDir 'Microsoft.Xaml.Behaviors.dll'))
[void][System.Reflection.Assembly]::LoadFrom((Join-Path $libDir 'ControlzEx.dll'))
[void][System.Reflection.Assembly]::LoadFrom((Join-Path $libDir 'MahApps.Metro.dll'))

# =============================================================================
# Module import.
# =============================================================================
$__modulePath = Join-Path $PSScriptRoot 'Module\CollectionManagerCommon.psd1'
if (-not (Test-Path -LiteralPath $__modulePath)) {
    throw "Shared module not found at: $__modulePath"
}
Import-Module -Name $__modulePath -Force -DisableNameChecking
if (-not (Get-Command Initialize-Logging -ErrorAction SilentlyContinue)) {
    throw "CollectionManagerCommon imported but Initialize-Logging is not exported."
}

# =============================================================================
# Preferences (CollectionManager.prefs.json next to the script).
# Closure-safe via $global:.
# =============================================================================
$global:PrefsPath = Join-Path $PSScriptRoot 'CollectionManager.prefs.json'

function Get-CmPreferences {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseSingularNouns', '', Justification='Returns the full preferences hashtable by design.')]
    param()
    $defaults = @{
        DarkMode    = $true
        SiteCode    = ''
        SMSProvider = ''
    }
    if (Test-Path -LiteralPath $global:PrefsPath) {
        try {
            $loaded = Get-Content -LiteralPath $global:PrefsPath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
            foreach ($k in @($defaults.Keys)) {
                $val = $loaded.$k
                if ($null -ne $val) { $defaults[$k] = $val }
            }
        } catch { $null = $_ }
    }
    return $defaults
}

function Save-CmPreferences {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseSingularNouns', '', Justification='Writes the full preferences hashtable by design.')]
    param([Parameter(Mandatory)][hashtable]$Prefs)
    try {
        $Prefs | ConvertTo-Json | Set-Content -LiteralPath $global:PrefsPath -Encoding UTF8
    } catch { $null = $_ }
}

$global:Prefs = Get-CmPreferences

# =============================================================================
# Tool log.
# =============================================================================
$script:ToolLogPath = Join-Path $__txDir ('CollectionManager-{0}.log' -f (Get-Date -Format 'yyyyMMdd-HHmmss'))
Initialize-Logging -LogPath $script:ToolLogPath

# =============================================================================
# Load XAML and resolve named elements.
# =============================================================================
$xamlPath = Join-Path $PSScriptRoot 'MainWindow.xaml'
if (-not (Test-Path -LiteralPath $xamlPath)) {
    throw "MainWindow.xaml not found at: $xamlPath"
}
[xml]$xaml = Get-Content -LiteralPath $xamlPath -Raw
$reader = New-Object System.Xml.XmlNodeReader $xaml
$window = [System.Windows.Markup.XamlReader]::Load($reader)

$txtAppTitle        = $window.FindName('txtAppTitle')
$txtVersion         = $window.FindName('txtVersion')
$txtThemeLabel      = $window.FindName('txtThemeLabel')
$toggleTheme        = $window.FindName('toggleTheme')

$btnViewCollections = $window.FindName('btnViewCollections')
$btnViewWqlEditor   = $window.FindName('btnViewWqlEditor')
$btnViewTemplates   = $window.FindName('btnViewTemplates')
$btnOptions         = $window.FindName('btnOptions')

$txtModuleTitle    = $window.FindName('txtModuleTitle')
$txtModuleSubtitle = $window.FindName('txtModuleSubtitle')

$btnRefresh           = $window.FindName('btnRefresh')
$txtFilter            = $window.FindName('txtFilter')
$cboCollectionFilter  = $window.FindName('cboCollectionFilter')
$btnExportCsv         = $window.FindName('btnExportCsv')
$btnExportHtml        = $window.FindName('btnExportHtml')

$viewCollections = $window.FindName('viewCollections')
$viewWqlEditor   = $window.FindName('viewWqlEditor')
$viewTemplates   = $window.FindName('viewTemplates')

$gridCollections     = $window.FindName('gridCollections')
$tabCollDetail       = $window.FindName('tabCollDetail')
$txtCollProperties   = $window.FindName('txtCollProperties')
$gridDirectMembers   = $window.FindName('gridDirectMembers')
$gridQueryRules      = $window.FindName('gridQueryRules')
$gridIncludeRules    = $window.FindName('gridIncludeRules')
$gridExcludeRules    = $window.FindName('gridExcludeRules')

$btnNewCollection  = $window.FindName('btnNewCollection')
$btnCopyCollection = $window.FindName('btnCopyCollection')
$btnEditMembership = $window.FindName('btnEditMembership')
$btnEvaluateColl   = $window.FindName('btnEvaluateColl')
$btnRemoveColl     = $window.FindName('btnRemoveColl')

$txtWqlTreeFilter      = $window.FindName('txtWqlTreeFilter')
$treeWqlCollections    = $window.FindName('treeWqlCollections')
$txtSelectedColl       = $window.FindName('txtSelectedColl')
$lstWqlRules           = $window.FindName('lstWqlRules')
$btnAddRule        = $window.FindName('btnAddRule')
$btnUpdateRule     = $window.FindName('btnUpdateRule')
$btnRemoveRule     = $window.FindName('btnRemoveRule')
$txtRuleName       = $window.FindName('txtRuleName')
$txtWqlEditor      = $window.FindName('txtWqlEditor')
$btnValidateWql    = $window.FindName('btnValidateWql')
$btnPreviewWql     = $window.FindName('btnPreviewWql')
$txtWqlValidation  = $window.FindName('txtWqlValidation')

$tabTemplateKind             = $window.FindName('tabTemplateKind')
$gridOperationalTemplates    = $window.FindName('gridOperationalTemplates')
$gridParameterizedTemplates  = $window.FindName('gridParameterizedTemplates')
$txtTemplateDescription      = $window.FindName('txtTemplateDescription')
$txtTemplateLimiting         = $window.FindName('txtTemplateLimiting')
$txtTemplateParamsHeader     = $window.FindName('txtTemplateParamsHeader')
$itemsTemplateParams         = $window.FindName('itemsTemplateParams')
$txtTemplateExpandedQuery    = $window.FindName('txtTemplateExpandedQuery')
$btnCopyToWqlEditor          = $window.FindName('btnCopyToWqlEditor')
$btnApplyTemplate            = $window.FindName('btnApplyTemplate')

$progressOverlay  = $window.FindName('progressOverlay')
$txtProgressTitle = $window.FindName('txtProgressTitle')
$txtProgressStep  = $window.FindName('txtProgressStep')

$lblLogOutput = $window.FindName('lblLogOutput')
$txtLog       = $window.FindName('txtLog')
$txtStatus    = $window.FindName('txtStatus')

$null = $txtAppTitle, $txtVersion, $tabTemplateKind

# =============================================================================
# Log drawer + status bar helpers.
# =============================================================================
function Add-LogLine {
    param([Parameter(Mandatory)][string]$Message)
    $ts = (Get-Date).ToString('HH:mm:ss')
    $line = '{0}  {1}' -f $ts, $Message
    if ([string]::IsNullOrWhiteSpace($txtLog.Text)) {
        $txtLog.Text = $line
    } else {
        $txtLog.AppendText([Environment]::NewLine + $line)
    }
    $txtLog.ScrollToEnd()
}

function Set-StatusText {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification='Updates an in-window TextBlock only; no external state.')]
    param([Parameter(Mandatory)][string]$Text)
    $txtStatus.Text = $Text
}

# =============================================================================
# Title-bar drag fallback. PS51-WPF-033.
# Some PowerShell launch contexts can leave MahApps' custom title thumb unable
# to initiate native window move. Install a WM_NCHITTEST hook returning
# HTCAPTION for the title band, plus a managed DragMove fallback for hosts
# where HwndSource cannot be hooked. Wire on every MetroWindow.
# =============================================================================
$script:TitleBarHitTestWindows = @{}
$script:TitleBarHitTestHooks   = @{}

function Get-TitleBarDragHeight {
    param([MahApps.Metro.Controls.MetroWindow]$Window)
    try {
        $h = [double]$Window.TitleBarHeight
        if ($h -gt 0 -and -not [double]::IsNaN($h)) { return $h }
    } catch { $null = $_ }
    return 30.0
}

function Get-InputAncestors {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseSingularNouns', '', Justification='Private visual-tree helper yields an ancestor chain.')]
    param([System.Windows.DependencyObject]$Start)
    $cur = $Start
    while ($cur) {
        $cur
        $parent = $null
        if ($cur -is [System.Windows.Media.Visual] -or $cur -is [System.Windows.Media.Media3D.Visual3D]) {
            try { $parent = [System.Windows.Media.VisualTreeHelper]::GetParent($cur) } catch { $parent = $null }
        }
        if (-not $parent -and $cur -is [System.Windows.FrameworkElement]) { $parent = $cur.Parent }
        if (-not $parent -and $cur -is [System.Windows.FrameworkContentElement]) { $parent = $cur.Parent }
        if (-not $parent -and $cur -is [System.Windows.ContentElement]) {
            try { $parent = [System.Windows.ContentOperations]::GetParent($cur) } catch { $parent = $null }
        }
        $cur = $parent
    }
}

function Test-IsWindowCommandPoint {
    param([MahApps.Metro.Controls.MetroWindow]$Window, [System.Windows.Point]$Point)
    try {
        [void]$Window.ApplyTemplate()
        $commands = $Window.Template.FindName('PART_WindowButtonCommands', $Window)
        if ($commands -and $commands.IsVisible -and $commands.ActualWidth -gt 0 -and $commands.ActualHeight -gt 0) {
            $origin = $commands.TransformToAncestor($Window).Transform([System.Windows.Point]::new(0, 0))
            if ($Point.X -ge $origin.X -and $Point.X -le ($origin.X + $commands.ActualWidth) -and
                $Point.Y -ge $origin.Y -and $Point.Y -le ($origin.Y + $commands.ActualHeight)) {
                return $true
            }
        }
    } catch { $null = $_ }
    return ($Window.ActualWidth -gt 150 -and $Point.X -ge ($Window.ActualWidth - 150))
}

function Add-NativeTitleBarHitTestHook {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification='Installs an in-process HWND hook for this WPF window only.')]
    param([MahApps.Metro.Controls.MetroWindow]$Window)
    try {
        $helper = [System.Windows.Interop.WindowInteropHelper]::new($Window)
        $source = [System.Windows.Interop.HwndSource]::FromHwnd($helper.Handle)
        if (-not $source) { return }
        $key = $helper.Handle.ToInt64().ToString()
        if ($script:TitleBarHitTestHooks.ContainsKey($key)) { return }
        $script:TitleBarHitTestWindows[$key] = $Window
        $hook = [System.Windows.Interop.HwndSourceHook]{
            param([IntPtr]$hwnd, [int]$msg, [IntPtr]$wParam, [IntPtr]$lParam, [ref]$handled)
            $WM_NCHITTEST = 0x0084; $HTCAPTION = 2
            if ($msg -ne $WM_NCHITTEST) { return [IntPtr]::Zero }
            try {
                $target = $script:TitleBarHitTestWindows[$hwnd.ToInt64().ToString()]
                if (-not $target) { return [IntPtr]::Zero }
                $raw = $lParam.ToInt64()
                $screenX = [int]($raw -band 0xffff); if ($screenX -ge 0x8000) { $screenX -= 0x10000 }
                $screenY = [int](($raw -shr 16) -band 0xffff); if ($screenY -ge 0x8000) { $screenY -= 0x10000 }
                $pt = $target.PointFromScreen([System.Windows.Point]::new($screenX, $screenY))
                $titleBarH = Get-TitleBarDragHeight -Window $target
                if ($pt.X -lt 0 -or $pt.X -gt $target.ActualWidth) { return [IntPtr]::Zero }
                if ($pt.Y -lt 4 -or $pt.Y -gt $titleBarH) { return [IntPtr]::Zero }
                if (Test-IsWindowCommandPoint -Window $target -Point $pt) { return [IntPtr]::Zero }
                $handled.Value = $true
                return [IntPtr]$HTCAPTION
            } catch { return [IntPtr]::Zero }
        }
        $script:TitleBarHitTestHooks[$key] = $hook
        $source.AddHook($hook)
    } catch { $null = $_ }
}

function Remove-NativeTitleBarHitTestHook {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification='Removes an in-process HWND hook for this WPF window only.')]
    param([MahApps.Metro.Controls.MetroWindow]$Window)
    try {
        $helper = [System.Windows.Interop.WindowInteropHelper]::new($Window)
        $key = $helper.Handle.ToInt64().ToString()
        if ($script:TitleBarHitTestHooks.ContainsKey($key)) {
            $source = [System.Windows.Interop.HwndSource]::FromHwnd($helper.Handle)
            if ($source) { $source.RemoveHook($script:TitleBarHitTestHooks[$key]) }
            $script:TitleBarHitTestHooks.Remove($key)
        }
        if ($script:TitleBarHitTestWindows.ContainsKey($key)) {
            $script:TitleBarHitTestWindows.Remove($key)
        }
    } catch { $null = $_ }
}

function Install-TitleBarDragFallback {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification='Registers window-local WPF event handlers for title-bar drag fallback.')]
    param([MahApps.Metro.Controls.MetroWindow]$Window)
    $Window.Add_SourceInitialized({ param($s, $e) Add-NativeTitleBarHitTestHook -Window $s })
    $Window.Add_Closed({ param($s, $e) Remove-NativeTitleBarHitTestHook -Window $s })
    $Window.Add_PreviewMouseLeftButtonDown({
        param($s, $e)
        try {
            if ($s.WindowState -eq [System.Windows.WindowState]::Maximized) { return }
            $titleBarH = Get-TitleBarDragHeight -Window $s
            $pos = $e.GetPosition($s)
            if ($pos.Y -lt 4 -or $pos.Y -gt $titleBarH) { return }
            if (Test-IsWindowCommandPoint -Window $s -Point $pos) { return }
            foreach ($ancestor in Get-InputAncestors -Start ($e.OriginalSource -as [System.Windows.DependencyObject])) {
                if ($ancestor -is [System.Windows.Controls.Primitives.ButtonBase]) { return }
            }
            $s.DragMove()
            $e.Handled = $true
        } catch { $null = $_ }
    })
}

Install-TitleBarDragFallback -Window $window

# =============================================================================
# Theme setup and toggle.
# =============================================================================
[void][ControlzEx.Theming.ThemeManager]::Current.ChangeTheme($window, 'Dark.Steel')

$script:DarkButtonBg      = [System.Windows.Media.BrushConverter]::new().ConvertFrom('#1E1E1E')
$script:DarkButtonBorder  = [System.Windows.Media.BrushConverter]::new().ConvertFrom('#555555')
$script:DarkActiveBg      = [System.Windows.Media.BrushConverter]::new().ConvertFrom('#3A3A3A')
$script:LightWfBg         = [System.Windows.Media.BrushConverter]::new().ConvertFrom('#0078D4')
$script:LightWfBorder     = [System.Windows.Media.BrushConverter]::new().ConvertFrom('#006CBE')
$script:LightActiveBg     = [System.Windows.Media.BrushConverter]::new().ConvertFrom('#005A9E')

$script:TitleBarBlue         = [System.Windows.Media.BrushConverter]::new().ConvertFrom('#0078D4')
$script:TitleBarBlueInactive = [System.Windows.Media.BrushConverter]::new().ConvertFrom('#4BA3E0')

$script:LogLabelDark  = [System.Windows.Media.BrushConverter]::new().ConvertFrom('#B0B0B0')
$script:LogLabelLight = [System.Windows.Media.BrushConverter]::new().ConvertFrom('#595959')

$script:ViewButtons = @(
    @{ Name = 'Collections'; Button = $btnViewCollections },
    @{ Name = 'WQL Editor';  Button = $btnViewWqlEditor   },
    @{ Name = 'Templates';   Button = $btnViewTemplates   }
)
$script:ActiveView = 'Collections'

function Update-SidebarButtonTheme {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification='Mutates in-window brush properties only.')]
    param()
    $isDark   = [bool]$global:Prefs['DarkMode']
    $idleBg   = if ($isDark) { $script:DarkButtonBg }     else { $script:LightWfBg }
    $activeBg = if ($isDark) { $script:DarkActiveBg }     else { $script:LightActiveBg }
    $border   = if ($isDark) { $script:DarkButtonBorder } else { $script:LightWfBorder }
    $thickness = [System.Windows.Thickness]::new(1)

    foreach ($v in $script:ViewButtons) {
        if (-not $v.Button) { continue }
        $isActive = ($v.Name -eq $script:ActiveView)
        $v.Button.Background      = if ($isActive) { $activeBg } else { $idleBg }
        $v.Button.BorderBrush     = $border
        $v.Button.BorderThickness = $thickness
    }
    if ($btnOptions) {
        $btnOptions.Background      = $idleBg
        $btnOptions.BorderBrush     = $border
        $btnOptions.BorderThickness = $thickness
    }
    if ($lblLogOutput) {
        $lblLogOutput.Foreground = if ($isDark) { $script:LogLabelDark } else { $script:LogLabelLight }
    }
}

function Update-TitleBarBrushes {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification='Mutates in-window brush properties only.')]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseSingularNouns', '', Justification='Sets both active and non-active title brushes per theme.')]
    param()
    $isDark = [bool]$global:Prefs['DarkMode']
    if ($isDark) {
        $window.ClearValue([MahApps.Metro.Controls.MetroWindow]::WindowTitleBrushProperty)
        $window.ClearValue([MahApps.Metro.Controls.MetroWindow]::NonActiveWindowTitleBrushProperty)
    } else {
        $window.WindowTitleBrush          = $script:TitleBarBlue
        $window.NonActiveWindowTitleBrush = $script:TitleBarBlueInactive
    }
}

$__startIsDark = [bool]$global:Prefs['DarkMode']
$toggleTheme.IsOn = $__startIsDark
$txtThemeLabel.Text = if ($__startIsDark) { 'Dark Theme' } else { 'Light Theme' }
Update-SidebarButtonTheme
# NOTE: ChangeTheme to a non-default theme + WindowTitleBrush mutation are
# DEFERRED to $window.Add_Loaded. Calling them at script-top breaks the
# title bar's NCHITTEST routing.

$toggleTheme.Add_Toggled({
    $isDark = [bool]$toggleTheme.IsOn
    if ($isDark) {
        [void][ControlzEx.Theming.ThemeManager]::Current.ChangeTheme($window, 'Dark.Steel')
        $txtThemeLabel.Text = 'Dark Theme'
    } else {
        [void][ControlzEx.Theming.ThemeManager]::Current.ChangeTheme($window, 'Light.Blue')
        $txtThemeLabel.Text = 'Light Theme'
    }
    $global:Prefs['DarkMode'] = $isDark
    Save-CmPreferences -Prefs $global:Prefs
    Update-SidebarButtonTheme
    Update-TitleBarBrushes
    Add-LogLine ('Theme: {0}' -f $(if ($isDark) { 'dark' } else { 'light' }))
})

# =============================================================================
# View switching.
# =============================================================================
$script:ViewMeta = @{
    'Collections' = @{ Title = 'Collections'; Subtitle = 'All device collections in the site. Configure Site / Provider in Options, then click Refresh.' }
    'WQL Editor'  = @{ Title = 'WQL Editor';  Subtitle = 'Edit query rules on a selected collection. Validate syntax and preview matches before committing.' }
    'Templates'   = @{ Title = 'Templates';   Subtitle = '157 ready-made operational queries plus 20 parameterized templates. Apply to a collection or copy into the WQL Editor.' }
}

function Set-ActiveView {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification='Updates in-window Visibility + header text only.')]
    param([Parameter(Mandatory)][ValidateSet('Collections','WQL Editor','Templates')][string]$View)

    $script:ActiveView = $View

    $viewCollections.Visibility = if ($View -eq 'Collections') { [System.Windows.Visibility]::Visible } else { [System.Windows.Visibility]::Collapsed }
    $viewWqlEditor.Visibility   = if ($View -eq 'WQL Editor')  { [System.Windows.Visibility]::Visible } else { [System.Windows.Visibility]::Collapsed }
    $viewTemplates.Visibility   = if ($View -eq 'Templates')   { [System.Windows.Visibility]::Visible } else { [System.Windows.Visibility]::Collapsed }

    $meta = $script:ViewMeta[$View]
    if ($meta) {
        $txtModuleTitle.Text    = $meta.Title
        $txtModuleSubtitle.Text = $meta.Subtitle
    }

    Update-SidebarButtonTheme
    Update-ActionBarVisibility
    Update-Filter
    Update-StatusBarSummary
}

$btnViewCollections.Add_Click({ Set-ActiveView -View 'Collections' })
$btnViewWqlEditor.Add_Click({   Set-ActiveView -View 'WQL Editor'   })
$btnViewTemplates.Add_Click({   Set-ActiveView -View 'Templates'    })

# =============================================================================
# Crash handlers (PS51-WPF-010, PS51-WPF-011, PS51-WPF-025).
# =============================================================================
$global:__crashLog = Join-Path $__txDir ('CollectionManager-crash-{0}.txt' -f (Get-Date -Format 'yyyyMMdd-HHmmss'))

$global:__writeCrash = {
    param($Source, $Exception)
    try {
        $lines = @()
        $lines += ('=== ' + $Source + ' @ ' + (Get-Date -Format 'o') + ' ===')
        $lines += ('Type   : ' + $Exception.GetType().FullName)
        $lines += ('Message: ' + $Exception.Message)
        $lines += ('Stack  :')
        $lines += ([string]$Exception.StackTrace).Split([Environment]::NewLine)
        $inner = $Exception.InnerException
        $depth = 1
        while ($inner) {
            $lines += ('--- InnerException depth ' + $depth + ' ---')
            $lines += ('Type   : ' + $inner.GetType().FullName)
            $lines += ('Message: ' + $inner.Message)
            $lines += ('Stack  :')
            $lines += ([string]$inner.StackTrace).Split([Environment]::NewLine)
            $inner = $inner.InnerException
            $depth++
        }
        [System.IO.File]::AppendAllText($global:__crashLog, (($lines -join [Environment]::NewLine) + [Environment]::NewLine))
    } catch { $null = $_ }
}

$window.Dispatcher.Add_UnhandledException({
    param($s, $e)
    & $global:__writeCrash 'DispatcherUnhandledException' $e.Exception
    $e.Handled = $false
})

[AppDomain]::CurrentDomain.Add_UnhandledException({
    param($s, $e)
    & $global:__writeCrash 'AppDomainUnhandledException' ([Exception]$e.ExceptionObject)
})

# =============================================================================
# Glyph mapping per brand: ✓/✗/⋯/⚠ at ThemeForeground (no colored fills).
# =============================================================================
function Get-CollectionTypeGlyph {
    param($Collection)
    if ($Collection.IsBuiltIn) { return [char]0x22EF }   # ⋯ system-managed, read-only
    return ''
}

# =============================================================================
# State for the collections / templates / WQL editor views.
# =============================================================================
$script:Collections          = @()         # decorated for grid (with TypeGlyph)
$script:RawCollections       = @()         # raw module output
$script:CollectionRulesIndex = @{}         # CollectionID -> @{ Direct=[]; Query=[]; Include=[]; Exclude=[] }
$script:LastRefreshTime      = $null
$script:IsConnectedFromBg    = $false

$script:OperationalTemplates    = @()
$script:ParameterizedTemplates  = @()

$script:CurrentTemplate    = $null         # selected template (operational or parameterized)
$script:CurrentTemplateRow = $null         # bound parameter rows for ItemsControl
$script:CurrentTemplateExpanded = ''       # most recent expanded query

# =============================================================================
# Action bar visibility and status bar summary.
# =============================================================================
function Update-ActionBarVisibility {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification='Toggles in-window Visibility only.')]
    param()

    switch ($script:ActiveView) {
        'Collections' {
            $cboCollectionFilter.Visibility = [System.Windows.Visibility]::Visible
            $btnExportCsv.Visibility        = [System.Windows.Visibility]::Visible
            $btnExportHtml.Visibility       = [System.Windows.Visibility]::Visible
            $txtFilter.Visibility           = [System.Windows.Visibility]::Visible
            $btnRefresh.Content             = 'Refresh'
            $txtFilter.Tag                  = 'Filter by name, ID, or comment...'
        }
        'WQL Editor' {
            # Hide every action-bar control except Refresh: the in-view tree's
            # own filter textbox handles collection filtering on this view, so
            # the wide global filter would just steal real estate at the top.
            $cboCollectionFilter.Visibility = [System.Windows.Visibility]::Collapsed
            $btnExportCsv.Visibility        = [System.Windows.Visibility]::Collapsed
            $btnExportHtml.Visibility       = [System.Windows.Visibility]::Collapsed
            $txtFilter.Visibility           = [System.Windows.Visibility]::Collapsed
            $btnRefresh.Content             = 'Refresh'
        }
        'Templates' {
            $cboCollectionFilter.Visibility = [System.Windows.Visibility]::Collapsed
            $btnExportCsv.Visibility        = [System.Windows.Visibility]::Collapsed
            $btnExportHtml.Visibility       = [System.Windows.Visibility]::Collapsed
            $txtFilter.Visibility           = [System.Windows.Visibility]::Visible
            $btnRefresh.Content             = 'Reload Templates'
            $txtFilter.Tag                  = 'Filter by category or name...'
        }
    }
    if ($txtFilter.Visibility -eq [System.Windows.Visibility]::Visible) {
        [MahApps.Metro.Controls.TextBoxHelper]::SetWatermark($txtFilter, [string]$txtFilter.Tag)
    }
}

function Update-StatusBarSummary {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification='Updates an in-window TextBlock only.')]
    param()

    $parts = @()
    if ($script:IsConnectedFromBg -and $global:Prefs.SiteCode) {
        $parts += "Connected to $($global:Prefs.SiteCode)"
    } elseif (-not $global:Prefs.SiteCode -or -not $global:Prefs.SMSProvider) {
        $parts += 'Open Options to configure site code and SMS provider'
    } else {
        $parts += 'Ready. Click Refresh.'
    }
    if ($script:RawCollections -and @($script:RawCollections).Count -gt 0) {
        $parts += ('{0} collections' -f @($script:RawCollections).Count)
    }
    if ($script:OperationalTemplates -and @($script:OperationalTemplates).Count -gt 0) {
        $parts += ('{0} op templates' -f @($script:OperationalTemplates).Count)
    }
    if ($script:ParameterizedTemplates -and @($script:ParameterizedTemplates).Count -gt 0) {
        $parts += ('{0} param templates' -f @($script:ParameterizedTemplates).Count)
    }
    if ($script:LastRefreshTime) {
        $parts += ('last refresh {0}' -f $script:LastRefreshTime.ToString('HH:mm:ss'))
    }
    Set-StatusText ($parts -join '   |   ')
}

# =============================================================================
# Filter and detail panels.
# =============================================================================
function Get-CollectionFilterValue {
    if (-not $cboCollectionFilter.SelectedItem) { return 'All' }
    $item = $cboCollectionFilter.SelectedItem
    if ($item -is [System.Windows.Controls.ComboBoxItem]) { return [string]$item.Content }
    return [string]$item
}

function Test-CollectionFilterMatch {
    param($Row, [string]$Filter)
    switch ($Filter) {
        'All'                  { return $true }
        'Built-in only'        { return [bool]$Row.IsBuiltIn }
        'Custom only'          { return -not [bool]$Row.IsBuiltIn }
        'Empty (zero members)' { return ([int]$Row.MemberCount -eq 0) }
        'Has direct members'   { return ([int]$Row.DirectCount -gt 0) }
        'Has query rules'      { return ([int]$Row.QueryCount -gt 0) }
        default                { return $true }
    }
}

function Update-Filter {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification='Recomputes ItemsSource on the active grid only.')]
    param()

    $needle = ([string]$txtFilter.Text).Trim().ToLowerInvariant()

    switch ($script:ActiveView) {
        'Collections' {
            $rows = $script:Collections
            if ($needle) {
                $rows = @($rows | Where-Object {
                    ([string]$_.Name).ToLowerInvariant().Contains($needle) -or
                    ([string]$_.CollectionID).ToLowerInvariant().Contains($needle) -or
                    ([string]$_.Comment).ToLowerInvariant().Contains($needle)
                })
            }
            $filterValue = Get-CollectionFilterValue
            if ($filterValue -ne 'All') {
                $rows = @($rows | Where-Object { Test-CollectionFilterMatch -Row $_ -Filter $filterValue })
            }
            $gridCollections.ItemsSource = $rows
        }
        'WQL Editor' {
            # No-op: the inline tree's own filter (txtWqlTreeFilter) drives
            # filtering on this view. The global filter textbox is hidden.
        }
        'Templates' {
            $opRows = $script:OperationalTemplates
            $paRows = $script:ParameterizedTemplates
            if ($needle) {
                $opRows = @($opRows | Where-Object {
                    ([string]$_.Name).ToLowerInvariant().Contains($needle) -or
                    ([string]$_.Category).ToLowerInvariant().Contains($needle)
                })
                $paRows = @($paRows | Where-Object {
                    ([string]$_.Name).ToLowerInvariant().Contains($needle) -or
                    ([string]$_.Category).ToLowerInvariant().Contains($needle)
                })
            }
            $gridOperationalTemplates.ItemsSource    = $opRows
            $gridParameterizedTemplates.ItemsSource  = $paRows
        }
    }
}

$txtFilter.Add_TextChanged({ Update-Filter })
$cboCollectionFilter.Add_SelectionChanged({ Update-Filter })

# Collection grid -> detail panel
$gridCollections.Add_SelectionChanged({
    $row = $gridCollections.SelectedItem
    if (-not $row) {
        $txtCollProperties.Text = 'Select a collection to see its properties.'
        $gridDirectMembers.ItemsSource = $null
        $gridQueryRules.ItemsSource    = $null
        $gridIncludeRules.ItemsSource  = $null
        $gridExcludeRules.ItemsSource  = $null
        return
    }

    $lines = @(
        ('Name:                 {0}' -f $row.Name),
        ('Collection ID:        {0}' -f $row.CollectionID),
        ('Member Count:         {0}' -f $row.MemberCount),
        ('Limiting Collection:  {0}' -f $row.LimitingCollection),
        ('Refresh Type:         {0}' -f $row.RefreshType),
        ('Built-in:             {0}' -f $row.IsBuiltIn),
        '',
        'Comment:',
        $row.Comment
    )
    $txtCollProperties.Text = $lines -join [Environment]::NewLine

    $idx = $script:CollectionRulesIndex[$row.CollectionID]
    if ($idx) {
        $gridQueryRules.ItemsSource    = $idx.Query
        $gridIncludeRules.ItemsSource  = $idx.Include
        $gridExcludeRules.ItemsSource  = $idx.Exclude
    } else {
        $gridQueryRules.ItemsSource    = $null
        $gridIncludeRules.ItemsSource  = $null
        $gridExcludeRules.ItemsSource  = $null
    }

    # Direct members are loaded lazily to keep refresh times bounded for
    # large environments. Triggered when the user picks the Direct Members tab.
    $gridDirectMembers.ItemsSource = $null
    $gridDirectMembers.Tag = $row.CollectionID
})

$tabCollDetail.Add_SelectionChanged({
    if ($tabCollDetail.SelectedIndex -ne 1) { return }   # Direct Members tab
    $collId = [string]$gridDirectMembers.Tag
    if (-not $collId) { return }
    if ($gridDirectMembers.ItemsSource) { return }       # already loaded

    Invoke-LoadDirectMembers -CollectionID $collId
})

# WQL Editor: inline TreeView (left pane) drives rule list + editor on the right.
$script:WqlSelectedCollection = $null

function Apply-WqlCollectionSelection {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification='Updates in-window TextBlock + ListBox only.')]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseApprovedVerbs', '', Justification='Apply reads naturally as the action verb here.')]
    param($Collection)
    $script:WqlSelectedCollection = $Collection
    if (-not $Collection) {
        $txtSelectedColl.Text = 'Pick a collection from the tree on the left.'
        $lstWqlRules.ItemsSource = @()
        $lstWqlRules.Tag = @()
        return
    }
    $txtSelectedColl.Text = ('Editing rules in: {0}  ({1}, {2} members)' -f $Collection.Name, $Collection.CollectionID, $Collection.MemberCount)
    $idx = $script:CollectionRulesIndex[$Collection.CollectionID]
    if ($idx) {
        $names = @($idx.Query | ForEach-Object { $_.RuleName })
        $lstWqlRules.Tag = $names
        $lstWqlRules.ItemsSource = $names
    } else {
        $lstWqlRules.Tag = @()
        $lstWqlRules.ItemsSource = @()
    }
    $txtRuleName.Text = ''
    $txtWqlEditor.Text = ''
    $txtWqlValidation.Text = ''
}

function Update-WqlTreeFromState {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification='Rebuilds the in-window WQL TreeView only.')]
    param([string]$Needle = '')
    if (-not $treeWqlCollections) { return }
    if (-not $script:Collections -or @($script:Collections).Count -eq 0) {
        $treeWqlCollections.Items.Clear()
        return
    }
    Build-CollectionTree -TreeView $treeWqlCollections `
        -AllCollections $script:Collections `
        -AllFolders     $script:Folders `
        -Needle         $Needle
}

$txtWqlTreeFilter.Add_TextChanged({
    Update-WqlTreeFromState -Needle ([string]$txtWqlTreeFilter.Text)
})

$treeWqlCollections.Add_SelectedItemChanged({
    $node = $treeWqlCollections.SelectedItem
    if (-not $node -or -not $node.Tag) { return }
    if ($node.Tag.Type -eq 'Collection') {
        Apply-WqlCollectionSelection -Collection $node.Tag.Object
    }
})

$lstWqlRules.Add_SelectionChanged({
    $name = [string]$lstWqlRules.SelectedItem
    if (-not $name) { return }
    $coll = $script:WqlSelectedCollection
    if (-not $coll) { return }
    $idx = $script:CollectionRulesIndex[$coll.CollectionID]
    if (-not $idx) { return }
    $rule = $idx.Query | Where-Object { $_.RuleName -eq $name } | Select-Object -First 1
    if ($rule) {
        $txtRuleName.Text  = $rule.RuleName
        $txtWqlEditor.Text = $rule.QueryExpression
        $txtWqlValidation.Text = ''
    }
})

# Templates: selection -> populate detail + parameter form
$gridOperationalTemplates.Add_SelectionChanged({
    $row = $gridOperationalTemplates.SelectedItem
    if (-not $row) { return }
    Show-TemplateDetail -Template $row -IsParameterized $false
})

$gridParameterizedTemplates.Add_SelectionChanged({
    $row = $gridParameterizedTemplates.SelectedItem
    if (-not $row) { return }
    Show-TemplateDetail -Template $row -IsParameterized $true
})

function Show-TemplateDetail {
    param([Parameter(Mandatory)]$Template, [bool]$IsParameterized)

    $script:CurrentTemplate = $Template
    $txtTemplateDescription.Text = [string]$Template.Description
    $txtTemplateLimiting.Text    = [string]$Template.LimitingCollection

    if ($IsParameterized -and $Template.Parameters -and @($Template.Parameters).Count -gt 0) {
        $rows = New-Object System.Collections.ObjectModel.ObservableCollection[psobject]
        foreach ($p in @($Template.Parameters)) {
            $obj = [PSCustomObject]@{
                Placeholder = [string]$p.Placeholder
                Label       = [string]$p.Label
                HelpText    = [string]$p.HelpText
                Value       = [string]$p.DefaultValue
            }
            $rows.Add($obj)
        }
        $script:CurrentTemplateRow = $rows
        $itemsTemplateParams.ItemsSource = $rows
        $txtTemplateParamsHeader.Visibility = [System.Windows.Visibility]::Visible

        # Live re-expand when any value changes. ItemsControl bind is two-way;
        # use a CollectionChanged + property-change loop on each row to refresh.
        foreach ($r in $rows) {
            $r | Add-Member -MemberType NoteProperty -Name '__pcb' -Value $null -Force -ErrorAction SilentlyContinue
        }

        Update-TemplateExpandedQuery
    } else {
        $script:CurrentTemplateRow = $null
        $itemsTemplateParams.ItemsSource = $null
        $txtTemplateParamsHeader.Visibility = [System.Windows.Visibility]::Collapsed
        $script:CurrentTemplateExpanded = [string]$Template.Query
        $txtTemplateExpandedQuery.Text = $script:CurrentTemplateExpanded
    }
}

function Update-TemplateExpandedQuery {
    if (-not $script:CurrentTemplate) { return }
    $tpl = $script:CurrentTemplate
    if (-not $script:CurrentTemplateRow) {
        $script:CurrentTemplateExpanded = [string]$tpl.Query
    } else {
        $values = @{}
        foreach ($r in $script:CurrentTemplateRow) {
            $values[[string]$r.Placeholder] = [string]$r.Value
        }
        $expanded = [string]$tpl.Query
        foreach ($k in $values.Keys) {
            $expanded = $expanded.Replace($k, $values[$k])
        }
        $script:CurrentTemplateExpanded = $expanded
    }
    $txtTemplateExpandedQuery.Text = $script:CurrentTemplateExpanded
}

# Re-expand on textbox edits inside the parameter ItemsControl. The
# ItemsControl binds Value Mode=TwoWay UpdateSourceTrigger=PropertyChanged,
# so each keystroke pushes the new value into the row. We just need to
# trigger Update-TemplateExpandedQuery on any descendant TextBox change.
$itemsTemplateParams.AddHandler(
    [System.Windows.Controls.TextBox]::TextChangedEvent,
    [System.Windows.RoutedEventHandler]{ Update-TemplateExpandedQuery }
)

# =============================================================================
# Background runspace for refresh and lazy member fetch.
# =============================================================================
$script:BgRunspace     = $null
$script:BgPowerShell   = $null
$script:BgInvokeHandle = $null
$script:BgState        = $null
$script:BgTimer        = $null

function Initialize-BgRunspace {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification='Lazy-init of the background runspace; idempotent.')]
    param()
    if ($script:BgRunspace -and $script:BgRunspace.RunspaceStateInfo.State -eq 'Opened') { return }

    $script:BgRunspace = [runspacefactory]::CreateRunspace()
    $script:BgRunspace.ApartmentState = 'STA'
    $script:BgRunspace.ThreadOptions  = 'ReuseThread'
    $script:BgRunspace.Open()

    $modulePath = Join-Path $PSScriptRoot 'Module\CollectionManagerCommon.psd1'

    $initPS = [powershell]::Create()
    $initPS.Runspace = $script:BgRunspace
    [void]$initPS.AddScript({
        param($ModulePath, $LogPath)
        Import-Module -Name $ModulePath -Force -DisableNameChecking
        if ($LogPath) { Initialize-Logging -LogPath $LogPath -Attach }
    }).AddArgument($modulePath).AddArgument($script:ToolLogPath)
    [void]$initPS.Invoke()
    $initPS.Dispose()
}

function Dispose-BgWork {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseApprovedVerbs', '', Justification='Dispose semantics intentional and reads as a single action.')]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification='Tears down ephemeral runspace plumbing only.')]
    param()
    if ($script:BgTimer)      { try { $script:BgTimer.Stop() }      catch { $null = $_ } ; $script:BgTimer = $null }
    if ($script:BgPowerShell) {
        try { [void]$script:BgPowerShell.Stop() } catch { $null = $_ }
        try { $script:BgPowerShell.Dispose() }   catch { $null = $_ }
        $script:BgPowerShell = $null
    }
    $script:BgInvokeHandle = $null
}

function Invoke-Refresh {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification='Posts work to the background runspace and arms a DispatcherTimer.')]
    param()

    if ($script:ActiveView -eq 'Templates') {
        # Templates load locally from disk; no MECM dependency.
        Invoke-LoadTemplates
        return
    }

    if (-not $global:Prefs.SiteCode -or -not $global:Prefs.SMSProvider) {
        Add-LogLine 'Refresh: site code and SMS provider must be set in Options first.'
        Set-StatusText 'Open Options to configure site code and SMS provider, then refresh.'
        return
    }

    Initialize-BgRunspace
    Dispose-BgWork

    $script:BgState = [hashtable]::Synchronized(@{
        Step     = 'Connecting...'
        Done     = $false
        Result   = $null
        ErrorMsg = $null
    })

    $btnRefresh.IsEnabled = $false
    $txtProgressTitle.Text = 'Loading collections'
    $txtProgressStep.Text  = 'Connecting...'
    $progressOverlay.Visibility = [System.Windows.Visibility]::Visible
    Add-LogLine ('Refresh: site={0} provider={1}' -f $global:Prefs.SiteCode, $global:Prefs.SMSProvider)
    Set-StatusText 'Refreshing...'

    $siteCode    = [string]$global:Prefs.SiteCode
    $smsProvider = [string]$global:Prefs.SMSProvider

    $script:BgPowerShell = [powershell]::Create()
    $script:BgPowerShell.Runspace = $script:BgRunspace
    [void]$script:BgPowerShell.AddScript({
        param($SiteCode, $SMSProvider, $State)
        try {
            if (-not (Test-CMConnection)) {
                $State.Step = "Connecting to $SiteCode..."
                $ok = Connect-CMSite -SiteCode $SiteCode -SMSProvider $SMSProvider
                if (-not $ok) {
                    $State.ErrorMsg = "Failed to connect to site $SiteCode (provider $SMSProvider)."
                    return
                }
            }

            $State.Step = 'Loading folder hierarchy...'
            $folders   = @()
            $folderMap = @{}
            try {
                $folders   = @(Get-CMCollectionFolderTree -SMSProvider $SMSProvider -SiteCode $SiteCode)
                $folderMap = Get-CMCollectionFolderMap -SMSProvider $SMSProvider -SiteCode $SiteCode
            } catch {
                # Non-fatal: if CIM is blocked, picker degrades to a flat root list.
                $folders = @()
                $folderMap = @{}
            }

            $State.Step = 'Loading device collections...'
            $collections = @(Get-AllDeviceCollections)

            # Annotate each collection with its FolderID (0 = tree root).
            $collections = @($collections | ForEach-Object {
                $fid = if ($folderMap.ContainsKey([string]$_.CollectionID)) { [int]$folderMap[[string]$_.CollectionID] } else { 0 }
                $_ | Add-Member -MemberType NoteProperty -Name FolderID -Value $fid -Force -PassThru
            })

            $State.Step = ('Parsing rules for {0} collections...' -f $collections.Count)
            $rulesIndex = @{}
            foreach ($c in $collections) {
                $direct  = @()
                $query   = @()
                $include = @()
                $exclude = @()

                foreach ($r in @($c.CollectionRules)) {
                    if ($null -eq $r) { continue }
                    $typeName = $null
                    try { $typeName = [string]$r.SmsProviderObjectPath } catch { $typeName = $null }
                    if (-not $typeName) {
                        try { $typeName = $r.GetType().Name } catch { continue }
                    }
                    switch -Wildcard ($typeName) {
                        '*RuleDirect*' {
                            $direct += [PSCustomObject]@{
                                Name       = $r.RuleName
                                ResourceID = $r.ResourceID
                                Domain     = ''
                            }
                        }
                        '*RuleQuery*' {
                            $query += [PSCustomObject]@{
                                RuleName        = $r.RuleName
                                QueryExpression = $r.QueryExpression
                            }
                        }
                        '*RuleIncludeCollection*' {
                            $include += [PSCustomObject]@{
                                IncludeCollectionName = $r.RuleName
                                IncludeCollectionID   = $r.IncludeCollectionID
                            }
                        }
                        '*RuleExcludeCollection*' {
                            $exclude += [PSCustomObject]@{
                                ExcludeCollectionName = $r.RuleName
                                ExcludeCollectionID   = $r.ExcludeCollectionID
                            }
                        }
                    }
                }

                $rulesIndex[$c.CollectionID] = @{
                    Direct  = $direct
                    Query   = $query
                    Include = $include
                    Exclude = $exclude
                }
            }

            $State.Result = [PSCustomObject]@{
                Collections = $collections
                RulesIndex  = $rulesIndex
                Folders     = $folders
            }
        }
        catch {
            $State.ErrorMsg = $_.Exception.Message
        }
        finally {
            $State.Done = $true
        }
    }).AddArgument($siteCode).AddArgument($smsProvider).AddArgument($script:BgState)

    $script:BgInvokeHandle = $script:BgPowerShell.BeginInvoke()

    $script:BgTimer = New-Object System.Windows.Threading.DispatcherTimer
    $script:BgTimer.Interval = [TimeSpan]::FromMilliseconds(150)
    $script:BgTimer.Add_Tick({
        if ($script:BgState) {
            $current = [string]$script:BgState.Step
            if ($txtProgressStep.Text -ne $current) { $txtProgressStep.Text = $current }
        }
        if ($script:BgState -and $script:BgState.Done) {
            $script:BgTimer.Stop()
            try { [void]$script:BgPowerShell.EndInvoke($script:BgInvokeHandle) } catch { $null = $_ }
            try { $script:BgPowerShell.Dispose() } catch { $null = $_ }
            $script:BgPowerShell   = $null
            $script:BgInvokeHandle = $null

            if ($script:BgState.ErrorMsg) {
                $progressOverlay.Visibility = [System.Windows.Visibility]::Collapsed
                $btnRefresh.IsEnabled = $true
                $script:IsConnectedFromBg = $false
                Add-LogLine ('Refresh failed: {0}' -f $script:BgState.ErrorMsg)
                Set-StatusText 'Refresh failed.'
                return
            }

            $script:IsConnectedFromBg = $true
            $r = $script:BgState.Result
            $script:RawCollections       = @($r.Collections)
            $script:CollectionRulesIndex = $r.RulesIndex
            $script:Folders              = @($r.Folders)
            $script:FolderById           = @{}
            foreach ($f in $script:Folders) { $script:FolderById[[int]$f.FolderID] = $f }
            $script:LastRefreshTime      = Get-Date

            $script:Collections = @($script:RawCollections | ForEach-Object {
                $idx = $script:CollectionRulesIndex[$_.CollectionID]
                $directCount = if ($idx) { @($idx.Direct).Count }  else { 0 }
                $queryCount  = if ($idx) { @($idx.Query).Count }   else { 0 }
                [PSCustomObject]@{
                    TypeGlyph          = Get-CollectionTypeGlyph -Collection $_
                    Name               = $_.Name
                    CollectionID       = $_.CollectionID
                    MemberCount        = $_.MemberCount
                    LimitingCollection = $_.LimitToCollectionName
                    RefreshType        = $_.RefreshType
                    Comment            = $_.Comment
                    IsBuiltIn          = $_.IsBuiltIn
                    DirectCount        = $directCount
                    QueryCount         = $queryCount
                    FolderID           = [int]$_.FolderID
                }
            })

            $gridCollections.ItemsSource = $script:Collections

            # Repaint the inline WQL Editor tree from the freshly-loaded state.
            Update-WqlTreeFromState -Needle ([string]$txtWqlTreeFilter.Text)
            # Re-apply the previously-selected collection's rule list, if any.
            if ($script:WqlSelectedCollection) {
                $reselected = $script:Collections | Where-Object { $_.CollectionID -eq $script:WqlSelectedCollection.CollectionID } | Select-Object -First 1
                if ($reselected) { Apply-WqlCollectionSelection -Collection $reselected }
                else { Apply-WqlCollectionSelection -Collection $null }
            }

            Update-Filter
            Update-StatusBarSummary

            $progressOverlay.Visibility = [System.Windows.Visibility]::Collapsed
            $btnRefresh.IsEnabled = $true

            # Mirror the bg runspace's CM connection into the UI thread so per-
            # action CM cmdlet calls from Add_Click handlers (New Collection,
            # Add/Update/Remove rule, Apply Template, Remove, Evaluate, Edit
            # Membership) resolve. CM cmdlets require the current location to
            # be the site PSDrive, and Set-Location is per-runspace.
            try {
                [void](Connect-CMSite -SiteCode $global:Prefs.SiteCode -SMSProvider $global:Prefs.SMSProvider)
            } catch {
                Add-LogLine ('UI-thread CM connect: {0}' -f $_.Exception.Message)
            }

            Add-LogLine ('Refresh complete: {0} collections.' -f @($script:RawCollections).Count)
        }
    })
    $script:BgTimer.Start()
}

function Invoke-LoadTemplates {
    Add-LogLine 'Loading templates from disk...'
    try {
        $opPath = Join-Path $PSScriptRoot 'Templates\operational-collections.json'
        $paPath = Join-Path $PSScriptRoot 'Templates\parameterized-templates.json'
        $script:OperationalTemplates = @(Get-OperationalTemplates  -Path $opPath)
        $script:ParameterizedTemplates = @(Get-ParameterizedTemplates -Path $paPath | ForEach-Object {
            $_ | Add-Member -MemberType NoteProperty -Name 'ParameterCount' -Value (@($_.Parameters).Count) -Force -PassThru
        })
        Add-LogLine ('Loaded {0} operational + {1} parameterized templates.' -f @($script:OperationalTemplates).Count, @($script:ParameterizedTemplates).Count)
        $gridOperationalTemplates.ItemsSource    = $script:OperationalTemplates
        $gridParameterizedTemplates.ItemsSource  = $script:ParameterizedTemplates
        Update-StatusBarSummary
    } catch {
        Add-LogLine ('Template load failed: {0}' -f $_.Exception.Message)
    }
}

function Invoke-LoadDirectMembers {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification='Posts a one-shot member fetch to the bg runspace.')]
    param([Parameter(Mandatory)][string]$CollectionID)

    if (-not $script:IsConnectedFromBg) {
        Add-LogLine 'Direct members: refresh first to establish a CM connection.'
        return
    }

    Initialize-BgRunspace

    $local = [hashtable]::Synchronized(@{ Done = $false; Result = $null; ErrorMsg = $null })
    $ps = [powershell]::Create()
    $ps.Runspace = $script:BgRunspace
    [void]$ps.AddScript({
        param($CollId, $State)
        try {
            $State.Result = @(Get-CollectionMembers -CollectionId $CollId)
        } catch {
            $State.ErrorMsg = $_.Exception.Message
        } finally {
            $State.Done = $true
        }
    }).AddArgument($CollectionID).AddArgument($local)

    $handle = $ps.BeginInvoke()

    $timer = New-Object System.Windows.Threading.DispatcherTimer
    $timer.Interval = [TimeSpan]::FromMilliseconds(120)
    $timer.Tag = @{ ps = $ps; handle = $handle; state = $local; collId = $CollectionID }
    $timer.Add_Tick({
        $info = $this.Tag
        if ($info.state.Done) {
            $this.Stop()
            try { [void]$info.ps.EndInvoke($info.handle) } catch { $null = $_ }
            try { $info.ps.Dispose() } catch { $null = $_ }
            if ($info.state.ErrorMsg) {
                Add-LogLine ('Direct members load failed: {0}' -f $info.state.ErrorMsg)
                return
            }
            # Only apply if the user hasn't moved on to a different collection.
            if ([string]$gridDirectMembers.Tag -eq [string]$info.collId) {
                $gridDirectMembers.ItemsSource = @($info.state.Result)
                Add-LogLine ('Loaded {0} direct members for {1}.' -f @($info.state.Result).Count, $info.collId)
            }
        }
    })
    $timer.Start()
}

$btnRefresh.Add_Click({ Invoke-Refresh })

# =============================================================================
# WQL editor wiring (validate / preview / add / update / remove).
# =============================================================================
$btnValidateWql.Add_Click({
    $wql = ([string]$txtWqlEditor.Text).Trim()
    if (-not $wql) {
        $txtWqlValidation.Text = 'Enter a WQL query first.'
        return
    }
    $result = Test-WqlQuery -QueryExpression $wql
    if ($result.IsValid) {
        $txtWqlValidation.Text = ('OK: query parses ({0} {1} returned).' -f $result.ResultCount, $(if ($result.ResultCount -eq 1) { 'row' } else { 'rows' }))
    } else {
        $txtWqlValidation.Text = ('Invalid: {0}' -f $result.ErrorMessage)
    }
    Add-LogLine ('WQL validate: {0}' -f $txtWqlValidation.Text)
})

$btnPreviewWql.Add_Click({
    $wql = ([string]$txtWqlEditor.Text).Trim()
    if (-not $wql) {
        $txtWqlValidation.Text = 'Enter a WQL query first.'
        return
    }
    if (-not $script:IsConnectedFromBg) {
        $txtWqlValidation.Text = 'Refresh first to establish a CM connection.'
        return
    }
    $txtWqlValidation.Text = 'Previewing...'
    Initialize-BgRunspace
    $local = [hashtable]::Synchronized(@{ Done = $false; Result = $null; ErrorMsg = $null })
    $ps = [powershell]::Create()
    $ps.Runspace = $script:BgRunspace
    [void]$ps.AddScript({
        param($Wql, $State)
        try { $State.Result = Invoke-WqlPreview -QueryExpression $Wql }
        catch { $State.ErrorMsg = $_.Exception.Message }
        finally { $State.Done = $true }
    }).AddArgument($wql).AddArgument($local)
    $handle = $ps.BeginInvoke()
    $timer = New-Object System.Windows.Threading.DispatcherTimer
    $timer.Interval = [TimeSpan]::FromMilliseconds(120)
    $timer.Tag = @{ ps = $ps; handle = $handle; state = $local }
    $timer.Add_Tick({
        $info = $this.Tag
        if ($info.state.Done) {
            $this.Stop()
            try { [void]$info.ps.EndInvoke($info.handle) } catch { $null = $_ }
            try { $info.ps.Dispose() } catch { $null = $_ }
            if ($info.state.ErrorMsg) {
                $txtWqlValidation.Text = ('Preview failed: {0}' -f $info.state.ErrorMsg)
                Add-LogLine $txtWqlValidation.Text
                return
            }
            $r = $info.state.Result
            $count = if ($r) { [int]$r.TotalCount } else { 0 }
            $txtWqlValidation.Text = ('Preview: {0} {1} matched.' -f $count, $(if ($count -eq 1) { 'row' } else { 'rows' }))
            Add-LogLine $txtWqlValidation.Text
        }
    })
    $timer.Start()
})

$btnAddRule.Add_Click({
    $coll = $script:WqlSelectedCollection
    $name = ([string]$txtRuleName.Text).Trim()
    $wql  = ([string]$txtWqlEditor.Text).Trim()
    if (-not $coll -or -not $name -or -not $wql) {
        $txtWqlValidation.Text = 'Select a collection and enter rule name + WQL.'
        return
    }
    try {
        Add-QueryRule -CollectionId $coll.CollectionID -RuleName $name -QueryExpression $wql
        Add-LogLine ('Added rule "{0}" to collection {1}.' -f $name, $coll.Name)
        Invoke-Refresh
    } catch {
        $txtWqlValidation.Text = ('Add rule failed: {0}' -f $_.Exception.Message)
        Add-LogLine $txtWqlValidation.Text
    }
})

$btnUpdateRule.Add_Click({
    $coll = $script:WqlSelectedCollection
    $name = ([string]$txtRuleName.Text).Trim()
    $wql  = ([string]$txtWqlEditor.Text).Trim()
    if (-not $coll -or -not $name -or -not $wql) {
        $txtWqlValidation.Text = 'Select a collection and a rule, and enter updated WQL.'
        return
    }
    $oldName = [string]$lstWqlRules.SelectedItem
    if (-not $oldName) { $oldName = $name }
    try {
        Update-QueryRule -CollectionId $coll.CollectionID -OldRuleName $oldName -NewRuleName $name -NewQueryExpression $wql
        Add-LogLine ('Updated rule "{0}" -> "{1}" on collection {2}.' -f $oldName, $name, $coll.Name)
        Invoke-Refresh
    } catch {
        $txtWqlValidation.Text = ('Update rule failed: {0}' -f $_.Exception.Message)
        Add-LogLine $txtWqlValidation.Text
    }
})

$btnRemoveRule.Add_Click({
    $coll = $script:WqlSelectedCollection
    $name = [string]$lstWqlRules.SelectedItem
    if (-not $coll -or -not $name) {
        $txtWqlValidation.Text = 'Select a collection and a rule to remove.'
        return
    }
    if (-not (Show-ConfirmDialog -Title 'Remove Rule' -Message ("Remove query rule '$name' from '$($coll.Name)'?"))) {
        return
    }
    try {
        Remove-QueryRule -CollectionId $coll.CollectionID -RuleName $name
        Add-LogLine ('Removed rule "{0}" from collection {1}.' -f $name, $coll.Name)
        Invoke-Refresh
    } catch {
        $txtWqlValidation.Text = ('Remove rule failed: {0}' -f $_.Exception.Message)
        Add-LogLine $txtWqlValidation.Text
    }
})

# =============================================================================
# Templates wiring (Copy to WQL Editor / Apply to Collection).
# =============================================================================
$btnCopyToWqlEditor.Add_Click({
    if (-not $script:CurrentTemplate -or -not $script:CurrentTemplateExpanded) {
        Add-LogLine 'Copy to WQL Editor: select a template first.'
        return
    }
    Set-ActiveView -View 'WQL Editor'
    $txtRuleName.Text  = [string]$script:CurrentTemplate.Name
    $txtWqlEditor.Text = [string]$script:CurrentTemplateExpanded
    $txtWqlValidation.Text = 'Template loaded into editor. Pick target collection and Validate / Add Rule.'
    Add-LogLine ('Loaded template "{0}" into the WQL Editor.' -f $script:CurrentTemplate.Name)
})

$btnApplyTemplate.Add_Click({
    if (Show-ApplyTemplateDialog) { Invoke-Refresh }
})

# =============================================================================
# Per-action modal dialogs (theme-honoring, drag-fallback installed).
# =============================================================================
function Set-DialogTheme {
    param([Parameter(Mandatory)][System.Windows.Window]$Dialog)
    $isDark = [bool]$global:Prefs['DarkMode']
    if ($isDark) {
        [void][ControlzEx.Theming.ThemeManager]::Current.ChangeTheme($Dialog, 'Dark.Steel')
    } else {
        [void][ControlzEx.Theming.ThemeManager]::Current.ChangeTheme($Dialog, 'Light.Blue')
        $Dialog.WindowTitleBrush          = $script:TitleBarBlue
        $Dialog.NonActiveWindowTitleBrush = $script:TitleBarBlueInactive
    }
}

function Show-NewCollectionDialog {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseSingularNouns', '', Justification='Modal dialog show / dispose; reads as a single action.')]
    param()

    if (-not $script:IsConnectedFromBg) {
        Add-LogLine 'New Collection: refresh first to establish a CM connection.'
        return $false
    }

    $dlgXaml = @'
<Controls:MetroWindow
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    xmlns:Controls="clr-namespace:MahApps.Metro.Controls;assembly=MahApps.Metro"
    Title="New Collection"
    Width="540" Height="380"
    WindowStartupLocation="CenterOwner"
    TitleCharacterCasing="Normal"
    GlowBrush="{DynamicResource MahApps.Brushes.Accent}"
    NonActiveGlowBrush="{DynamicResource MahApps.Brushes.Accent}"
    BorderThickness="1"
    ResizeMode="NoResize"
    ShowIconOnTitleBar="False">
    <Window.Resources>
        <ResourceDictionary>
            <ResourceDictionary.MergedDictionaries>
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Controls.xaml" />
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Fonts.xaml" />
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Themes/Dark.Steel.xaml" />
            </ResourceDictionary.MergedDictionaries>
            <Style x:Key="DialogButton" TargetType="Button" BasedOn="{StaticResource MahApps.Styles.Button.Square}">
                <Setter Property="MinWidth" Value="90"/><Setter Property="Height" Value="32"/>
                <Setter Property="Margin" Value="0,0,8,0"/>
                <Setter Property="Controls:ControlsHelper.ContentCharacterCasing" Value="Normal"/>
            </Style>
            <Style x:Key="DialogAccentButton" TargetType="Button" BasedOn="{StaticResource MahApps.Styles.Button.Square.Accent}">
                <Setter Property="MinWidth" Value="90"/><Setter Property="Height" Value="32"/>
                <Setter Property="Margin" Value="0,0,8,0"/>
                <Setter Property="Controls:ControlsHelper.ContentCharacterCasing" Value="Normal"/>
            </Style>
        </ResourceDictionary>
    </Window.Resources>
    <Grid Margin="20,16,20,12">
        <Grid.RowDefinitions>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        <StackPanel Grid.Row="0">
            <TextBlock Text="Name" FontSize="11" Margin="0,0,0,2" Foreground="{DynamicResource MahApps.Brushes.Gray1}"/>
            <TextBox x:Name="txtName" FontSize="12" Padding="6,4,6,4"
                     Controls:TextBoxHelper.Watermark="e.g. Workstations - Pilot Ring"/>

            <TextBlock Text="Limiting Collection" FontSize="11" Margin="0,12,0,2" Foreground="{DynamicResource MahApps.Brushes.Gray1}"/>
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <Border Grid.Column="0" BorderThickness="1"
                        BorderBrush="{DynamicResource MahApps.Brushes.Gray8}"
                        Background="{DynamicResource MahApps.Brushes.ThemeBackground}"
                        Padding="8,5,8,5" Margin="0,0,8,0">
                    <TextBlock x:Name="txtLimiting" FontSize="12" Text="(none)"
                               Foreground="{DynamicResource MahApps.Brushes.Gray1}"
                               VerticalAlignment="Center" TextTrimming="CharacterEllipsis"/>
                </Border>
                <Button x:Name="btnPickLimiting" Grid.Column="1" Content="Browse..."
                        Style="{StaticResource DialogButton}" MinWidth="90" Margin="0"
                        ToolTip="Browse the folder tree to pick the limiting collection"/>
            </Grid>

            <TextBlock Text="Comment" FontSize="11" Margin="0,12,0,2" Foreground="{DynamicResource MahApps.Brushes.Gray1}"/>
            <TextBox x:Name="txtComment" FontSize="12" Padding="6,4,6,4"
                     Controls:TextBoxHelper.Watermark="Optional"/>

            <TextBlock Text="Refresh Type" FontSize="11" Margin="0,12,0,2" Foreground="{DynamicResource MahApps.Brushes.Gray1}"/>
            <ComboBox x:Name="cboRefresh" FontSize="12">
                <ComboBoxItem Content="Manual"/>
                <ComboBoxItem Content="Periodic"/>
                <ComboBoxItem Content="Continuous"/>
                <ComboBoxItem Content="Both" IsSelected="True"/>
            </ComboBox>
        </StackPanel>
        <StackPanel Grid.Row="1" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,16,0,0">
            <Button x:Name="btnOk"     Content="Create" Style="{StaticResource DialogAccentButton}" IsDefault="True"/>
            <Button x:Name="btnCancel" Content="Cancel" Style="{StaticResource DialogButton}"        IsCancel="True"/>
        </StackPanel>
    </Grid>
</Controls:MetroWindow>
'@
    [xml]$dx = $dlgXaml
    $reader2 = New-Object System.Xml.XmlNodeReader $dx
    $dlg = [System.Windows.Markup.XamlReader]::Load($reader2)
    $dlg.Owner = $window
    Install-TitleBarDragFallback -Window $dlg
    Set-DialogTheme -Dialog $dlg

    $txtName        = $dlg.FindName('txtName')
    $txtLimiting    = $dlg.FindName('txtLimiting')
    $btnPickLimiting= $dlg.FindName('btnPickLimiting')
    $txtComment     = $dlg.FindName('txtComment')
    $cboRefresh     = $dlg.FindName('cboRefresh')
    $btnOk          = $dlg.FindName('btnOk')
    $btnCancel      = $dlg.FindName('btnCancel')

    $script:NewLimiting = $script:Collections | Where-Object { $_.Name -eq 'All Systems' } | Select-Object -First 1
    if ($script:NewLimiting) {
        $txtLimiting.Text = ('{0}  ({1})' -f $script:NewLimiting.Name, $script:NewLimiting.CollectionID)
    }
    $btnPickLimiting.Add_Click({
        $picked = Show-CollectionPickerDialog -Title 'Pick Limiting Collection'
        if ($picked) {
            $script:NewLimiting = $picked
            $txtLimiting.Text = ('{0}  ({1})' -f $picked.Name, $picked.CollectionID)
        }
    })

    $script:NewCollResult = $false
    $btnOk.Add_Click({
        $name = ([string]$txtName.Text).Trim()
        $lim  = $script:NewLimiting
        if (-not $name -or -not $lim) {
            Add-LogLine 'New Collection: name and limiting collection are required.'
            return
        }
        $rt = if ($cboRefresh.SelectedItem) { [string]$cboRefresh.SelectedItem.Content } else { 'Both' }
        try {
            $params = @{
                Name                   = $name
                LimitingCollectionName = $lim.Name
                RefreshType            = $rt
            }
            if ($txtComment.Text) { $params['Comment'] = ([string]$txtComment.Text) }
            New-ManagedCollection @params | Out-Null
            Add-LogLine ('Created collection "{0}" (limiting: {1}, refresh: {2}).' -f $name, $lim.Name, $rt)
            $script:NewCollResult = $true
            $dlg.DialogResult = $true
            $dlg.Close()
        } catch {
            Add-LogLine ('New Collection failed: {0}' -f $_.Exception.Message)
        }
    })
    $btnCancel.Add_Click({ $dlg.DialogResult = $false; $dlg.Close() })

    [void]$dlg.ShowDialog()
    return $script:NewCollResult
}

function Show-CopyCollectionDialog {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseSingularNouns', '', Justification='Modal dialog show / dispose; reads as a single action.')]
    param([Parameter(Mandatory)]$Source)

    $dlgXaml = @'
<Controls:MetroWindow
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    xmlns:Controls="clr-namespace:MahApps.Metro.Controls;assembly=MahApps.Metro"
    Title="Copy Collection"
    Width="520" SizeToContent="Height"
    WindowStartupLocation="CenterOwner"
    TitleCharacterCasing="Normal"
    GlowBrush="{DynamicResource MahApps.Brushes.Accent}"
    NonActiveGlowBrush="{DynamicResource MahApps.Brushes.Accent}"
    BorderThickness="1"
    ResizeMode="NoResize"
    ShowIconOnTitleBar="False">
    <Window.Resources>
        <ResourceDictionary>
            <ResourceDictionary.MergedDictionaries>
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Controls.xaml" />
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Fonts.xaml" />
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Themes/Dark.Steel.xaml" />
            </ResourceDictionary.MergedDictionaries>
            <Style x:Key="DialogButton" TargetType="Button" BasedOn="{StaticResource MahApps.Styles.Button.Square}">
                <Setter Property="MinWidth" Value="90"/><Setter Property="Height" Value="32"/>
                <Setter Property="Margin" Value="0,0,8,0"/>
                <Setter Property="Controls:ControlsHelper.ContentCharacterCasing" Value="Normal"/>
            </Style>
            <Style x:Key="DialogAccentButton" TargetType="Button" BasedOn="{StaticResource MahApps.Styles.Button.Square.Accent}">
                <Setter Property="MinWidth" Value="90"/><Setter Property="Height" Value="32"/>
                <Setter Property="Margin" Value="0,0,8,0"/>
                <Setter Property="Controls:ControlsHelper.ContentCharacterCasing" Value="Normal"/>
            </Style>
        </ResourceDictionary>
    </Window.Resources>
    <Grid Margin="20,16,20,12">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        <StackPanel Grid.Row="0">
            <TextBlock Text="Source" FontSize="11" Margin="0,0,0,2" Foreground="{DynamicResource MahApps.Brushes.Gray1}"/>
            <TextBlock x:Name="txtSource" FontSize="13" FontWeight="SemiBold"/>
        </StackPanel>
        <StackPanel Grid.Row="1" Margin="0,16,0,0">
            <TextBlock Text="New Name" FontSize="11" Margin="0,0,0,2" Foreground="{DynamicResource MahApps.Brushes.Gray1}"/>
            <TextBox x:Name="txtNewName" FontSize="12" Padding="6,4,6,4"
                     Controls:TextBoxHelper.Watermark="Name of the new collection"/>
        </StackPanel>
        <StackPanel Grid.Row="2" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,16,0,0">
            <Button x:Name="btnOk"     Content="Copy"   Style="{StaticResource DialogAccentButton}" IsDefault="True"/>
            <Button x:Name="btnCancel" Content="Cancel" Style="{StaticResource DialogButton}"        IsCancel="True"/>
        </StackPanel>
    </Grid>
</Controls:MetroWindow>
'@
    [xml]$dx = $dlgXaml
    $reader2 = New-Object System.Xml.XmlNodeReader $dx
    $dlg = [System.Windows.Markup.XamlReader]::Load($reader2)
    $dlg.Owner = $window
    Install-TitleBarDragFallback -Window $dlg
    Set-DialogTheme -Dialog $dlg

    $txtSource  = $dlg.FindName('txtSource')
    $txtNewName = $dlg.FindName('txtNewName')
    $btnOk      = $dlg.FindName('btnOk')
    $btnCancel  = $dlg.FindName('btnCancel')

    $txtSource.Text  = ('{0} ({1})' -f $Source.Name, $Source.CollectionID)
    $txtNewName.Text = ('Copy of ' + $Source.Name)

    $script:CopyCollResult = $false
    $btnOk.Add_Click({
        $newName = ([string]$txtNewName.Text).Trim()
        if (-not $newName) { Add-LogLine 'Copy: new name is required.'; return }
        try {
            Copy-ManagedCollection -SourceCollectionId $Source.CollectionID -NewName $newName | Out-Null
            Add-LogLine ('Copied "{0}" -> "{1}".' -f $Source.Name, $newName)
            $script:CopyCollResult = $true
            $dlg.DialogResult = $true
            $dlg.Close()
        } catch {
            Add-LogLine ('Copy failed: {0}' -f $_.Exception.Message)
        }
    })
    $btnCancel.Add_Click({ $dlg.DialogResult = $false; $dlg.Close() })

    [void]$dlg.ShowDialog()
    return $script:CopyCollResult
}

function Show-EditMembershipDialog {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseSingularNouns', '', Justification='Modal dialog show / dispose; reads as a single action.')]
    param([Parameter(Mandatory)]$Collection)

    if (-not $script:IsConnectedFromBg) {
        Add-LogLine 'Edit Membership: refresh first to establish a CM connection.'
        return $false
    }

    $dlgXaml = @'
<Controls:MetroWindow
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    xmlns:Controls="clr-namespace:MahApps.Metro.Controls;assembly=MahApps.Metro"
    Title="Edit Membership"
    Width="720" Height="540"
    MinWidth="640" MinHeight="480"
    WindowStartupLocation="CenterOwner"
    TitleCharacterCasing="Normal"
    GlowBrush="{DynamicResource MahApps.Brushes.Accent}"
    NonActiveGlowBrush="{DynamicResource MahApps.Brushes.Accent}"
    BorderThickness="1"
    ShowIconOnTitleBar="False">
    <Window.Resources>
        <ResourceDictionary>
            <ResourceDictionary.MergedDictionaries>
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Controls.xaml" />
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Fonts.xaml" />
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Themes/Dark.Steel.xaml" />
            </ResourceDictionary.MergedDictionaries>
            <Style x:Key="DialogButton" TargetType="Button" BasedOn="{StaticResource MahApps.Styles.Button.Square}">
                <Setter Property="MinWidth" Value="90"/><Setter Property="Height" Value="28"/>
                <Setter Property="Margin" Value="0,0,8,0"/>
                <Setter Property="Controls:ControlsHelper.ContentCharacterCasing" Value="Normal"/>
            </Style>
        </ResourceDictionary>
    </Window.Resources>
    <Grid Margin="16,12,16,12">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        <TextBlock x:Name="txtCollName" FontSize="14" FontWeight="SemiBold" Margin="0,0,0,12"/>

        <DataGrid x:Name="gridMembers" Grid.Row="1" AutoGenerateColumns="False"
                  CanUserAddRows="False" CanUserDeleteRows="False" IsReadOnly="True"
                  SelectionMode="Extended" SelectionUnit="FullRow"
                  Background="{DynamicResource MahApps.Brushes.ThemeBackground}"
                  Foreground="{DynamicResource MahApps.Brushes.ThemeForeground}"
                  RowHeaderWidth="0" BorderThickness="0" FontSize="11">
            <DataGrid.Columns>
                <DataGridTextColumn Header="Name"        Width="2*" Binding="{Binding Name}"/>
                <DataGridTextColumn Header="ResourceID"  Width="120" Binding="{Binding ResourceID}"/>
                <DataGridTextColumn Header="Domain"      Width="*"  Binding="{Binding Domain}"/>
            </DataGrid.Columns>
        </DataGrid>

        <StackPanel Grid.Row="2" Orientation="Horizontal" Margin="0,12,0,0">
            <TextBox x:Name="txtAddDevice" Width="240" FontSize="12" Padding="6,4,6,4"
                     Controls:TextBoxHelper.Watermark="Device name"/>
            <Button x:Name="btnAddDevice"    Content="Add Direct Member"  Style="{StaticResource DialogButton}" Margin="8,0,0,0" MinWidth="160"/>
            <Button x:Name="btnRemoveSel"    Content="Remove Selected"    Style="{StaticResource DialogButton}" Margin="8,0,0,0" MinWidth="140"/>
            <Button x:Name="btnReloadMembers" Content="Reload"            Style="{StaticResource DialogButton}" Margin="8,0,0,0" MinWidth="100"/>
        </StackPanel>

        <StackPanel Grid.Row="3" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,16,0,0">
            <Button x:Name="btnClose" Content="Close" Style="{StaticResource DialogButton}" IsCancel="True"/>
        </StackPanel>
    </Grid>
</Controls:MetroWindow>
'@
    [xml]$dx = $dlgXaml
    $reader2 = New-Object System.Xml.XmlNodeReader $dx
    $dlg = [System.Windows.Markup.XamlReader]::Load($reader2)
    $dlg.Owner = $window
    Install-TitleBarDragFallback -Window $dlg
    Set-DialogTheme -Dialog $dlg

    $txtCollName     = $dlg.FindName('txtCollName')
    $gridMembers     = $dlg.FindName('gridMembers')
    $txtAddDevice    = $dlg.FindName('txtAddDevice')
    $btnAddDevice    = $dlg.FindName('btnAddDevice')
    $btnRemoveSel    = $dlg.FindName('btnRemoveSel')
    $btnReloadMembers = $dlg.FindName('btnReloadMembers')
    $btnClose        = $dlg.FindName('btnClose')

    $txtCollName.Text = ('Direct members: {0} ({1})' -f $Collection.Name, $Collection.CollectionID)

    $reload = {
        try {
            $members = @(Get-CollectionMembers -CollectionId $Collection.CollectionID)
            $gridMembers.ItemsSource = $members
        } catch {
            Add-LogLine ('Edit Membership: load failed: {0}' -f $_.Exception.Message)
        }
    }
    & $reload

    $btnAddDevice.Add_Click({
        $name = ([string]$txtAddDevice.Text).Trim()
        if (-not $name) { return }
        try {
            Add-DirectMember -CollectionId $Collection.CollectionID -DeviceName $name
            Add-LogLine ('Added "{0}" to {1}.' -f $name, $Collection.Name)
            $txtAddDevice.Text = ''
            & $reload
        } catch {
            Add-LogLine ('Add direct member failed: {0}' -f $_.Exception.Message)
        }
    })

    $btnRemoveSel.Add_Click({
        $rows = @($gridMembers.SelectedItems)
        if (@($rows).Count -eq 0) { return }
        if (-not (Show-ConfirmDialog -Title 'Remove Members' -Message ("Remove $(@($rows).Count) selected member(s) from $($Collection.Name)?"))) {
            return
        }
        foreach ($r in $rows) {
            try {
                Remove-DirectMember -CollectionId $Collection.CollectionID -ResourceId ([int]$r.ResourceID)
                Add-LogLine ('Removed "{0}" (ResourceID {1}) from {2}.' -f $r.Name, $r.ResourceID, $Collection.Name)
            } catch {
                Add-LogLine ('Remove failed for ResourceID {0}: {1}' -f $r.ResourceID, $_.Exception.Message)
            }
        }
        & $reload
    })

    $btnReloadMembers.Add_Click({ & $reload })
    $btnClose.Add_Click({ $dlg.DialogResult = $false; $dlg.Close() })

    [void]$dlg.ShowDialog()
    return $true
}

function Show-ApplyTemplateDialog {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseSingularNouns', '', Justification='Modal dialog show / dispose; reads as a single action.')]
    param()

    if (-not $script:CurrentTemplate -or -not $script:CurrentTemplateExpanded) {
        Add-LogLine 'Apply Template: select a template first.'
        return $false
    }
    if (-not $script:IsConnectedFromBg) {
        Add-LogLine 'Apply Template: refresh first to establish a CM connection.'
        return $false
    }

    $dlgXaml = @'
<Controls:MetroWindow
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    xmlns:Controls="clr-namespace:MahApps.Metro.Controls;assembly=MahApps.Metro"
    Title="Apply Template to Collection"
    Width="640" Height="500"
    MinWidth="540" MinHeight="440"
    WindowStartupLocation="CenterOwner"
    TitleCharacterCasing="Normal"
    GlowBrush="{DynamicResource MahApps.Brushes.Accent}"
    NonActiveGlowBrush="{DynamicResource MahApps.Brushes.Accent}"
    BorderThickness="1"
    ShowIconOnTitleBar="False">
    <Window.Resources>
        <ResourceDictionary>
            <ResourceDictionary.MergedDictionaries>
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Controls.xaml" />
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Fonts.xaml" />
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Themes/Dark.Steel.xaml" />
            </ResourceDictionary.MergedDictionaries>
            <Style x:Key="DialogButton" TargetType="Button" BasedOn="{StaticResource MahApps.Styles.Button.Square}">
                <Setter Property="MinWidth" Value="90"/><Setter Property="Height" Value="32"/>
                <Setter Property="Margin" Value="0,0,8,0"/>
                <Setter Property="Controls:ControlsHelper.ContentCharacterCasing" Value="Normal"/>
            </Style>
            <Style x:Key="DialogAccentButton" TargetType="Button" BasedOn="{StaticResource MahApps.Styles.Button.Square.Accent}">
                <Setter Property="MinWidth" Value="90"/><Setter Property="Height" Value="32"/>
                <Setter Property="Margin" Value="0,0,8,0"/>
                <Setter Property="Controls:ControlsHelper.ContentCharacterCasing" Value="Normal"/>
            </Style>
        </ResourceDictionary>
    </Window.Resources>
    <Grid Margin="20,16,20,12">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        <StackPanel Grid.Row="0">
            <TextBlock Text="Template" FontSize="11" Margin="0,0,0,2" Foreground="{DynamicResource MahApps.Brushes.Gray1}"/>
            <TextBlock x:Name="txtTemplateName" FontSize="13" FontWeight="SemiBold"/>
        </StackPanel>
        <StackPanel Grid.Row="1" Margin="0,12,0,0">
            <TextBlock Text="Target Collection" FontSize="11" Margin="0,0,0,2" Foreground="{DynamicResource MahApps.Brushes.Gray1}"/>
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <Border Grid.Column="0" BorderThickness="1"
                        BorderBrush="{DynamicResource MahApps.Brushes.Gray8}"
                        Background="{DynamicResource MahApps.Brushes.ThemeBackground}"
                        Padding="8,5,8,5" Margin="0,0,8,0">
                    <TextBlock x:Name="txtTarget" FontSize="12" Text="(no target selected)"
                               Foreground="{DynamicResource MahApps.Brushes.Gray1}"
                               VerticalAlignment="Center" TextTrimming="CharacterEllipsis"/>
                </Border>
                <Button x:Name="btnPickTarget" Grid.Column="1" Content="Browse..."
                        Style="{StaticResource DialogButton}" MinWidth="90" Margin="0"
                        ToolTip="Browse the folder tree to pick the target collection"/>
            </Grid>
        </StackPanel>
        <StackPanel Grid.Row="2" Margin="0,12,0,0">
            <TextBlock Text="Rule Name" FontSize="11" Margin="0,0,0,2" Foreground="{DynamicResource MahApps.Brushes.Gray1}"/>
            <TextBox x:Name="txtRuleName" FontSize="12" Padding="6,4,6,4"/>
        </StackPanel>
        <TextBlock Grid.Row="3" Text="Expanded WQL" FontSize="11" Margin="0,12,0,2" Foreground="{DynamicResource MahApps.Brushes.Gray1}"/>
        <TextBox x:Name="txtExpanded" Grid.Row="4" IsReadOnly="True"
                 AcceptsReturn="True" TextWrapping="NoWrap"
                 VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Auto"
                 FontFamily="Cascadia Code, Consolas, Courier New" FontSize="12" Padding="8"
                 Background="{DynamicResource MahApps.Brushes.ThemeBackground}"
                 Foreground="{DynamicResource MahApps.Brushes.ThemeForeground}"
                 BorderThickness="1"
                 BorderBrush="{DynamicResource MahApps.Brushes.Gray8}"/>
        <StackPanel Grid.Row="5" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,12,0,0">
            <Button x:Name="btnOk"     Content="Apply"  Style="{StaticResource DialogAccentButton}" IsDefault="True"/>
            <Button x:Name="btnCancel" Content="Cancel" Style="{StaticResource DialogButton}"        IsCancel="True"/>
        </StackPanel>
    </Grid>
</Controls:MetroWindow>
'@
    [xml]$dx = $dlgXaml
    $reader2 = New-Object System.Xml.XmlNodeReader $dx
    $dlg = [System.Windows.Markup.XamlReader]::Load($reader2)
    $dlg.Owner = $window
    Install-TitleBarDragFallback -Window $dlg
    Set-DialogTheme -Dialog $dlg

    $txtTemplateName = $dlg.FindName('txtTemplateName')
    $txtTarget       = $dlg.FindName('txtTarget')
    $btnPickTarget   = $dlg.FindName('btnPickTarget')
    $txtRuleName2    = $dlg.FindName('txtRuleName')
    $txtExpanded     = $dlg.FindName('txtExpanded')
    $btnOk           = $dlg.FindName('btnOk')
    $btnCancel       = $dlg.FindName('btnCancel')

    $script:ApplyTarget = $null
    $btnPickTarget.Add_Click({
        # Built-ins (SMS prefix) can't host new query rules; filter them out.
        $picked = Show-CollectionPickerDialog -Title 'Pick Target Collection' -IncludeBuiltIn $false
        if ($picked) {
            $script:ApplyTarget = $picked
            $txtTarget.Text = ('{0}  ({1})' -f $picked.Name, $picked.CollectionID)
        }
    })

    $txtTemplateName.Text = [string]$script:CurrentTemplate.Name
    $txtRuleName2.Text    = [string]$script:CurrentTemplate.Name
    $txtExpanded.Text     = [string]$script:CurrentTemplateExpanded

    $script:ApplyTplResult = $false
    $btnOk.Add_Click({
        $target = $script:ApplyTarget
        $ruleName = ([string]$txtRuleName2.Text).Trim()
        if (-not $target -or -not $ruleName) {
            Add-LogLine 'Apply Template: pick a target collection and rule name.'
            return
        }
        try {
            Add-QueryRule -CollectionId $target.CollectionID -RuleName $ruleName -QueryExpression ([string]$txtExpanded.Text)
            Add-LogLine ('Applied template "{0}" as rule "{1}" on {2}.' -f $script:CurrentTemplate.Name, $ruleName, $target.Name)
            $script:ApplyTplResult = $true
            $dlg.DialogResult = $true
            $dlg.Close()
        } catch {
            Add-LogLine ('Apply Template failed: {0}' -f $_.Exception.Message)
        }
    })
    $btnCancel.Add_Click({ $dlg.DialogResult = $false; $dlg.Close() })

    [void]$dlg.ShowDialog()
    return $script:ApplyTplResult
}

# =============================================================================
# Wired action button handlers.
# =============================================================================
$btnNewCollection.Add_Click({
    if (Show-NewCollectionDialog) { Invoke-Refresh }
})
$btnCopyCollection.Add_Click({
    $row = $gridCollections.SelectedItem
    if (-not $row) { Add-LogLine 'Copy: pick a source collection first.'; return }
    if (Show-CopyCollectionDialog -Source $row) { Invoke-Refresh }
})
$btnEditMembership.Add_Click({
    $row = $gridCollections.SelectedItem
    if (-not $row) { Add-LogLine 'Edit Membership: pick a collection first.'; return }
    if ($row.IsBuiltIn) {
        Add-LogLine 'Edit Membership: built-in collections are read-only.'
        return
    }
    [void](Show-EditMembershipDialog -Collection $row)
    Invoke-Refresh
})
$btnEvaluateColl.Add_Click({
    $row = $gridCollections.SelectedItem
    if (-not $row) {
        Add-LogLine 'Evaluate: pick a collection first.'
        return
    }
    if (-not $script:IsConnectedFromBg) {
        Add-LogLine 'Evaluate: refresh first to establish a CM connection.'
        return
    }
    try {
        Invoke-CollectionEvaluation -CollectionId $row.CollectionID
        Add-LogLine ('Triggered membership evaluation for {0}.' -f $row.Name)
    } catch {
        Add-LogLine ('Evaluate failed: {0}' -f $_.Exception.Message)
    }
})
$btnRemoveColl.Add_Click({
    $row = $gridCollections.SelectedItem
    if (-not $row) { Add-LogLine 'Remove: pick a collection first.'; return }
    if ($row.IsBuiltIn) {
        Add-LogLine ('Remove: built-in collection {0} cannot be deleted.' -f $row.CollectionID)
        return
    }
    if (-not (Show-ConfirmDialog -Title 'Remove Collection' -Message ("Permanently delete collection '$($row.Name)' ($($row.CollectionID))?"))) {
        return
    }
    try {
        $ok = Remove-ManagedCollection -CollectionId $row.CollectionID
        if ($ok) {
            Add-LogLine ('Removed collection "{0}" ({1}).' -f $row.Name, $row.CollectionID)
            Invoke-Refresh
        }
    } catch {
        Add-LogLine ('Remove failed: {0}' -f $_.Exception.Message)
    }
})

# =============================================================================
# Export buttons.
# =============================================================================
function Get-ActiveExportInfo {
    switch ($script:ActiveView) {
        'Collections' {
            return @{
                Name    = 'Collections'
                Columns = @('Name','CollectionID','MemberCount','LimitingCollection','RefreshType','Comment')
                Rows    = $gridCollections.ItemsSource
            }
        }
        default { return $null }
    }
}

$btnExportCsv.Add_Click({
    $info = Get-ActiveExportInfo
    if (-not $info -or -not $info.Rows -or @($info.Rows).Count -eq 0) {
        Add-LogLine 'Export CSV: nothing to export.'
        return
    }
    $sfd = New-Object Microsoft.Win32.SaveFileDialog
    $sfd.Filter = 'CSV files (*.csv)|*.csv'
    $sfd.FileName = ('CollectionManager-{0}-{1}.csv' -f $info.Name, (Get-Date -Format 'yyyyMMdd-HHmmss'))
    $reportsDir = Join-Path $PSScriptRoot 'Reports'
    if (-not (Test-Path -LiteralPath $reportsDir)) { New-Item -ItemType Directory -Path $reportsDir -Force | Out-Null }
    $sfd.InitialDirectory = $reportsDir
    if ($sfd.ShowDialog() -eq $true) {
        Export-CollectionCsv -Collections $info.Rows -OutputPath $sfd.FileName
        Add-LogLine ('Exported CSV: {0}' -f $sfd.FileName)
    }
})

$btnExportHtml.Add_Click({
    $info = Get-ActiveExportInfo
    if (-not $info -or -not $info.Rows -or @($info.Rows).Count -eq 0) {
        Add-LogLine 'Export HTML: nothing to export.'
        return
    }
    $sfd = New-Object Microsoft.Win32.SaveFileDialog
    $sfd.Filter = 'HTML files (*.html)|*.html'
    $sfd.FileName = ('CollectionManager-{0}-{1}.html' -f $info.Name, (Get-Date -Format 'yyyyMMdd-HHmmss'))
    $reportsDir = Join-Path $PSScriptRoot 'Reports'
    if (-not (Test-Path -LiteralPath $reportsDir)) { New-Item -ItemType Directory -Path $reportsDir -Force | Out-Null }
    $sfd.InitialDirectory = $reportsDir
    if ($sfd.ShowDialog() -eq $true) {
        Export-CollectionHtml -Collections $info.Rows -OutputPath $sfd.FileName
        Add-LogLine ('Exported HTML: {0}' -f $sfd.FileName)
    }
})

# =============================================================================
# Tree builder used by both the inline WQL Editor TreeView and the modal
# picker dialog. Mirrors the CM console's left-pane folder hierarchy via
# SMS_ObjectContainerNode/Item. Returns the count of collection leaves that
# survived the filter.
# =============================================================================
function Build-CollectionTree {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification='Mutates the in-window TreeView only.')]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseApprovedVerbs', '', Justification='Build is the natural verb for tree assembly.')]
    param(
        [Parameter(Mandatory)][System.Windows.Controls.TreeView]$TreeView,
        [Parameter(Mandatory)]$AllCollections,
        $AllFolders,
        [string]$Needle = ''
    )

    $colls   = @($AllCollections)
    $folders = @($AllFolders)

    $needleLower = ([string]$Needle).Trim().ToLowerInvariant()
    $hasFilter   = -not [string]::IsNullOrWhiteSpace($needleLower)

    $foldersByParent = @{}
    foreach ($f in $folders) {
        $parentId = [int]$f.ParentID
        if (-not $foldersByParent.ContainsKey($parentId)) { $foldersByParent[$parentId] = @() }
        $foldersByParent[$parentId] += $f
    }
    $collectionsByFolder = @{}
    foreach ($c in $colls) {
        $fid = [int]$c.FolderID
        if (-not $collectionsByFolder.ContainsKey($fid)) { $collectionsByFolder[$fid] = @() }
        $collectionsByFolder[$fid] += $c
    }

    $TreeView.Items.Clear()
    $script:__TreeLeafCount = 0

    $matchesNeedle = {
        param($Coll)
        if (-not $hasFilter) { return $true }
        $name = ([string]$Coll.Name).ToLowerInvariant()
        $id   = ([string]$Coll.CollectionID).ToLowerInvariant()
        return ($name.Contains($needleLower) -or $id.Contains($needleLower))
    }

    $populate = {
        param($ParentNode, [int]$FolderID)
        $any = $false

        $childFolders = if ($foldersByParent.ContainsKey($FolderID)) { @($foldersByParent[$FolderID] | Sort-Object Name) } else { @() }
        foreach ($f in $childFolders) {
            $folderNode = New-Object System.Windows.Controls.TreeViewItem
            $folderNode.Header = ('[+] {0}' -f $f.Name)
            $folderNode.Tag = @{ Type = 'Folder'; Object = $f }
            $folderNode.FontWeight = [System.Windows.FontWeights]::SemiBold
            if ($hasFilter) { $folderNode.IsExpanded = $true }
            $hadAny = & $populate $folderNode ([int]$f.FolderID)
            if ($hadAny -or -not $hasFilter) {
                [void]$ParentNode.Items.Add($folderNode)
                $any = $true
            }
        }

        $childColls = if ($collectionsByFolder.ContainsKey($FolderID)) { @($collectionsByFolder[$FolderID] | Sort-Object Name) } else { @() }
        foreach ($c in $childColls) {
            if (-not (& $matchesNeedle $c)) { continue }
            $collNode = New-Object System.Windows.Controls.TreeViewItem
            $collNode.Header = ('{0}  ({1}, {2} members)' -f $c.Name, $c.CollectionID, $c.MemberCount)
            $collNode.Tag = @{ Type = 'Collection'; Object = $c }
            [void]$ParentNode.Items.Add($collNode)
            $script:__TreeLeafCount++
            $any = $true
        }
        return $any
    }

    & $populate $TreeView 0
    return $script:__TreeLeafCount
}

# =============================================================================
# Tree picker dialog. Used by the New Collection (limiting) and Apply Template
# (target) modals; the WQL Editor view uses an inline tree instead.
# =============================================================================
function Show-CollectionPickerDialog {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseSingularNouns', '', Justification='Modal dialog show / dispose; reads as a single action.')]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter', 'IncludeBuiltIn', Justification='Surfaced as a flag now; consumed by the build closure later.')]
    param(
        [string]$Title = 'Pick Collection',
        [bool]$IncludeBuiltIn = $true
    )

    if (-not $script:Collections -or @($script:Collections).Count -eq 0) {
        Add-LogLine 'Picker: refresh first to load collections.'
        return $null
    }

    $dlgXaml = @'
<Controls:MetroWindow
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    xmlns:Controls="clr-namespace:MahApps.Metro.Controls;assembly=MahApps.Metro"
    Title=""
    Width="720" Height="640"
    MinWidth="540" MinHeight="420"
    WindowStartupLocation="CenterOwner"
    TitleCharacterCasing="Normal"
    GlowBrush="{DynamicResource MahApps.Brushes.Accent}"
    NonActiveGlowBrush="{DynamicResource MahApps.Brushes.Accent}"
    BorderThickness="1"
    ShowIconOnTitleBar="False">
    <Window.Resources>
        <ResourceDictionary>
            <ResourceDictionary.MergedDictionaries>
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Controls.xaml" />
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Fonts.xaml" />
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Themes/Dark.Steel.xaml" />
            </ResourceDictionary.MergedDictionaries>
            <Style x:Key="DialogButton" TargetType="Button" BasedOn="{StaticResource MahApps.Styles.Button.Square}">
                <Setter Property="MinWidth" Value="90"/><Setter Property="Height" Value="32"/>
                <Setter Property="Margin" Value="0,0,8,0"/>
                <Setter Property="Controls:ControlsHelper.ContentCharacterCasing" Value="Normal"/>
            </Style>
            <Style x:Key="DialogAccentButton" TargetType="Button" BasedOn="{StaticResource MahApps.Styles.Button.Square.Accent}">
                <Setter Property="MinWidth" Value="90"/><Setter Property="Height" Value="32"/>
                <Setter Property="Margin" Value="0,0,8,0"/>
                <Setter Property="Controls:ControlsHelper.ContentCharacterCasing" Value="Normal"/>
            </Style>
        </ResourceDictionary>
    </Window.Resources>
    <Grid Margin="16,12,16,12">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        <TextBox x:Name="txtPickerFilter" Grid.Row="0" FontSize="12" Padding="6,4,6,4" Margin="0,4,0,8"
                 Controls:TextBoxHelper.Watermark="Filter by collection name or ID..."/>
        <Border Grid.Row="1" BorderThickness="1"
                BorderBrush="{DynamicResource MahApps.Brushes.Gray8}"
                Background="{DynamicResource MahApps.Brushes.ThemeBackground}">
            <TreeView x:Name="treePicker" FontSize="12"
                      VirtualizingStackPanel.IsVirtualizing="True"
                      VirtualizingStackPanel.VirtualizationMode="Recycling"
                      Background="{DynamicResource MahApps.Brushes.ThemeBackground}"
                      Foreground="{DynamicResource MahApps.Brushes.ThemeForeground}"
                      BorderThickness="0"/>
        </Border>
        <TextBlock x:Name="txtPickerStatus" Grid.Row="2" FontSize="11" Margin="0,8,0,0"
                   Foreground="{DynamicResource MahApps.Brushes.Gray1}"
                   Text="Pick a collection (folders are not selectable)."/>
        <StackPanel Grid.Row="3" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,12,0,0">
            <Button x:Name="btnOk"     Content="OK"     Style="{StaticResource DialogAccentButton}" IsDefault="True" IsEnabled="False"/>
            <Button x:Name="btnCancel" Content="Cancel" Style="{StaticResource DialogButton}"        IsCancel="True"/>
        </StackPanel>
    </Grid>
</Controls:MetroWindow>
'@
    [xml]$dx = $dlgXaml
    $reader2 = New-Object System.Xml.XmlNodeReader $dx
    $dlg = [System.Windows.Markup.XamlReader]::Load($reader2)
    $dlg.Owner = $window
    $dlg.Title = $Title
    Install-TitleBarDragFallback -Window $dlg
    Set-DialogTheme -Dialog $dlg

    $txtPickerFilter = $dlg.FindName('txtPickerFilter')
    $treePicker      = $dlg.FindName('treePicker')
    $txtPickerStatus = $dlg.FindName('txtPickerStatus')
    $btnOk           = $dlg.FindName('btnOk')
    $btnCancel       = $dlg.FindName('btnCancel')

    $allCollections = if ($IncludeBuiltIn) { $script:Collections } else { $script:Collections | Where-Object { -not $_.IsBuiltIn } }
    $allCollections = @($allCollections)

    $rebuildTree = {
        param([string]$Needle)
        $count = Build-CollectionTree -TreeView $treePicker `
            -AllCollections $allCollections `
            -AllFolders     $script:Folders `
            -Needle         $Needle
        if ([string]::IsNullOrWhiteSpace($Needle)) {
            $totalColls = @($allCollections).Count
            $totalFolders = @($script:Folders).Count
            $txtPickerStatus.Text = ('{0} collections across {1} folders. Pick one (folders are not selectable).' -f $totalColls, $totalFolders)
        } else {
            $txtPickerStatus.Text = ('{0} collections match "{1}".' -f $count, $Needle.Trim())
        }
    }

    & $rebuildTree ''

    $script:PickerResult = $null
    $treePicker.Add_SelectedItemChanged({
        $node = $treePicker.SelectedItem
        if (-not $node -or -not $node.Tag -or $node.Tag.Type -ne 'Collection') {
            $btnOk.IsEnabled = $false
            $script:PickerResult = $null
            return
        }
        $btnOk.IsEnabled = $true
        $script:PickerResult = $node.Tag.Object
    })

    $txtPickerFilter.Add_TextChanged({ & $rebuildTree ([string]$txtPickerFilter.Text) })

    $btnOk.Add_Click({
        if ($script:PickerResult) {
            $dlg.DialogResult = $true
            $dlg.Close()
        }
    })
    $btnCancel.Add_Click({
        $script:PickerResult = $null
        $dlg.DialogResult = $false
        $dlg.Close()
    })

    [void]$dlg.ShowDialog()
    return $script:PickerResult
}

# =============================================================================
# Themed Yes/No confirmation dialog (shared by remove flows).
# =============================================================================
function Show-ConfirmDialog {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseSingularNouns', '', Justification='Modal dialog show / dispose; reads as a single action.')]
    param(
        [Parameter(Mandatory)][string]$Title,
        [Parameter(Mandatory)][string]$Message
    )

    $dlgXaml = @'
<Controls:MetroWindow
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    xmlns:Controls="clr-namespace:MahApps.Metro.Controls;assembly=MahApps.Metro"
    Title=""
    Width="480" SizeToContent="Height"
    MinWidth="380"
    WindowStartupLocation="CenterOwner"
    TitleCharacterCasing="Normal"
    GlowBrush="{DynamicResource MahApps.Brushes.Accent}"
    NonActiveGlowBrush="{DynamicResource MahApps.Brushes.Accent}"
    BorderThickness="1"
    ResizeMode="NoResize"
    ShowIconOnTitleBar="False">
    <Window.Resources>
        <ResourceDictionary>
            <ResourceDictionary.MergedDictionaries>
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Controls.xaml" />
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Fonts.xaml" />
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Themes/Dark.Steel.xaml" />
            </ResourceDictionary.MergedDictionaries>
            <Style x:Key="DialogButton" TargetType="Button" BasedOn="{StaticResource MahApps.Styles.Button.Square}">
                <Setter Property="MinWidth" Value="90"/>
                <Setter Property="Height"   Value="32"/>
                <Setter Property="Margin"   Value="0,0,8,0"/>
                <Setter Property="Controls:ControlsHelper.ContentCharacterCasing" Value="Normal"/>
            </Style>
            <Style x:Key="DialogAccentButton" TargetType="Button" BasedOn="{StaticResource MahApps.Styles.Button.Square.Accent}">
                <Setter Property="MinWidth" Value="90"/>
                <Setter Property="Height"   Value="32"/>
                <Setter Property="Margin"   Value="0,0,8,0"/>
                <Setter Property="Controls:ControlsHelper.ContentCharacterCasing" Value="Normal"/>
            </Style>
        </ResourceDictionary>
    </Window.Resources>
    <Grid Margin="16,12,16,12">
        <Grid.RowDefinitions>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        <TextBlock x:Name="txtMsg" Grid.Row="0" TextWrapping="Wrap" FontSize="13" Margin="0,8,0,16"/>
        <StackPanel Grid.Row="1" Orientation="Horizontal" HorizontalAlignment="Right">
            <Button x:Name="btnYes" Content="Yes" Style="{StaticResource DialogAccentButton}" IsDefault="True"/>
            <Button x:Name="btnNo"  Content="No"  Style="{StaticResource DialogButton}"        IsCancel="True"/>
        </StackPanel>
    </Grid>
</Controls:MetroWindow>
'@
    [xml]$dx = $dlgXaml
    $reader2 = New-Object System.Xml.XmlNodeReader $dx
    $dlg = [System.Windows.Markup.XamlReader]::Load($reader2)
    $dlg.Owner = $window
    $dlg.Title = $Title
    Install-TitleBarDragFallback -Window $dlg

    $isDark = [bool]$global:Prefs['DarkMode']
    if ($isDark) {
        [void][ControlzEx.Theming.ThemeManager]::Current.ChangeTheme($dlg, 'Dark.Steel')
    } else {
        [void][ControlzEx.Theming.ThemeManager]::Current.ChangeTheme($dlg, 'Light.Blue')
        $dlg.WindowTitleBrush          = $script:TitleBarBlue
        $dlg.NonActiveWindowTitleBrush = $script:TitleBarBlueInactive
    }

    $txtMsg = $dlg.FindName('txtMsg')
    $btnYes = $dlg.FindName('btnYes')
    $btnNo  = $dlg.FindName('btnNo')
    $txtMsg.Text = $Message

    $btnYes.Add_Click({ $dlg.DialogResult = $true;  $dlg.Close() })
    $btnNo.Add_Click({  $dlg.DialogResult = $false; $dlg.Close() })

    return [bool]$dlg.ShowDialog()
}

# =============================================================================
# Options dialog -- Connection / About. Phase 6.
# =============================================================================
function Show-OptionsDialog {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseSingularNouns', '', Justification='Modal dialog show / dispose; reads as a single action.')]
    param()

    $dlgXaml = @'
<Controls:MetroWindow
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    xmlns:Controls="clr-namespace:MahApps.Metro.Controls;assembly=MahApps.Metro"
    Title="Options"
    Width="640" Height="380"
    MinWidth="560" MinHeight="380"
    WindowStartupLocation="CenterOwner"
    TitleCharacterCasing="Normal"
    GlowBrush="{DynamicResource MahApps.Brushes.Accent}"
    NonActiveGlowBrush="{DynamicResource MahApps.Brushes.Accent}"
    BorderThickness="1"
    ShowIconOnTitleBar="False">
    <Window.Resources>
        <ResourceDictionary>
            <ResourceDictionary.MergedDictionaries>
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Controls.xaml" />
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Fonts.xaml" />
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Themes/Dark.Steel.xaml" />
            </ResourceDictionary.MergedDictionaries>
            <Style x:Key="CategoryRowStyle" TargetType="Button" BasedOn="{StaticResource MahApps.Styles.Button.Square}">
                <Setter Property="Height" Value="36"/>
                <Setter Property="HorizontalContentAlignment" Value="Left"/>
                <Setter Property="Padding" Value="14,0,14,0"/>
                <Setter Property="FontSize" Value="13"/>
                <Setter Property="Controls:ControlsHelper.ContentCharacterCasing" Value="Normal"/>
                <Setter Property="Margin" Value="0"/>
            </Style>
            <Style x:Key="DialogButton" TargetType="Button" BasedOn="{StaticResource MahApps.Styles.Button.Square}">
                <Setter Property="MinWidth" Value="90"/>
                <Setter Property="Height" Value="32"/>
                <Setter Property="Margin" Value="0,0,8,0"/>
                <Setter Property="Controls:ControlsHelper.ContentCharacterCasing" Value="Normal"/>
            </Style>
            <Style x:Key="DialogAccentButton" TargetType="Button" BasedOn="{StaticResource MahApps.Styles.Button.Square.Accent}">
                <Setter Property="MinWidth" Value="90"/>
                <Setter Property="Height" Value="32"/>
                <Setter Property="Margin" Value="0,0,8,0"/>
                <Setter Property="Controls:ControlsHelper.ContentCharacterCasing" Value="Normal"/>
            </Style>
        </ResourceDictionary>
    </Window.Resources>
    <Grid>
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="180"/>
            <ColumnDefinition Width="1"/>
            <ColumnDefinition Width="*"/>
        </Grid.ColumnDefinitions>
        <Grid.RowDefinitions>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <Border Grid.Column="0" Grid.Row="0" Padding="6,12,0,12">
            <StackPanel>
                <Button x:Name="btnCatConnection" Content="Connection" Style="{StaticResource CategoryRowStyle}"/>
                <Button x:Name="btnCatAbout"      Content="About"      Style="{StaticResource CategoryRowStyle}"/>
            </StackPanel>
        </Border>

        <Border Grid.Column="1" Grid.Row="0" Background="{DynamicResource MahApps.Brushes.Gray8}"/>

        <Grid Grid.Column="2" Grid.Row="0" Margin="20,16,20,16">
            <StackPanel x:Name="paneConnection" Visibility="Visible">
                <TextBlock Text="MECM Connection" FontSize="13" FontWeight="SemiBold" Margin="0,0,0,10"/>
                <TextBlock Text="Site Code" FontSize="11" Margin="0,4,0,2"
                           Foreground="{DynamicResource MahApps.Brushes.Gray1}"/>
                <TextBox x:Name="txtSiteCode" FontSize="12" Padding="6,4,6,4"
                         Controls:TextBoxHelper.Watermark="e.g. P01" Width="120" HorizontalAlignment="Left"/>
                <TextBlock Text="SMS Provider FQDN" FontSize="11" Margin="0,12,0,2"
                           Foreground="{DynamicResource MahApps.Brushes.Gray1}"/>
                <TextBox x:Name="txtSmsProvider" FontSize="12" Padding="6,4,6,4"
                         Controls:TextBoxHelper.Watermark="e.g. cm01.contoso.com"/>
                <TextBlock Text="Used for the CM PSDrive root. Collection Manager performs both reads (Get-CMDeviceCollection) and writes (New / Set / Remove) -- the account running this app needs the matching CM permissions."
                           FontSize="11" TextWrapping="Wrap" Margin="0,16,0,0"
                           Foreground="{DynamicResource MahApps.Brushes.Gray1}"/>
            </StackPanel>

            <StackPanel x:Name="paneAbout" Visibility="Collapsed">
                <TextBlock Text="About" FontSize="13" FontWeight="SemiBold" Margin="0,0,0,10"/>
                <TextBlock x:Name="txtAboutVersion" Text="Collection Manager v1.0.1"
                           FontSize="13" FontWeight="SemiBold"/>
                <TextBlock Text="Browse, create, copy, and remove MECM device collections. Edit query rules with offline WQL validation and a 1-shot result preview. Apply ready-made operational queries or fill out parameterized templates and add the resulting rule to a target collection."
                           FontSize="12" TextWrapping="Wrap" Margin="0,8,0,0"/>
                <TextBlock Text="157 ready-made operational queries plus 20 parameterized templates ship with the app."
                           FontSize="12" TextWrapping="Wrap" Margin="0,12,0,0"/>
                <TextBlock Text="Author: Jason Ulbright. License: MIT."
                           FontSize="11" Margin="0,16,0,0"
                           Foreground="{DynamicResource MahApps.Brushes.Gray1}"/>
            </StackPanel>
        </Grid>

        <Border Grid.Row="1" Grid.ColumnSpan="3" Padding="16,12,16,12">
            <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
                <Button x:Name="btnOk"     Content="OK"     Style="{StaticResource DialogAccentButton}" IsDefault="True"/>
                <Button x:Name="btnCancel" Content="Cancel" Style="{StaticResource DialogButton}"        IsCancel="True"/>
            </StackPanel>
        </Border>
    </Grid>
</Controls:MetroWindow>
'@

    [xml]$dx = $dlgXaml
    $reader2 = New-Object System.Xml.XmlNodeReader $dx
    $dlg = [System.Windows.Markup.XamlReader]::Load($reader2)
    $dlg.Owner = $window
    Install-TitleBarDragFallback -Window $dlg

    $isDark = [bool]$global:Prefs['DarkMode']
    if ($isDark) {
        [void][ControlzEx.Theming.ThemeManager]::Current.ChangeTheme($dlg, 'Dark.Steel')
    } else {
        [void][ControlzEx.Theming.ThemeManager]::Current.ChangeTheme($dlg, 'Light.Blue')
        $dlg.WindowTitleBrush          = $script:TitleBarBlue
        $dlg.NonActiveWindowTitleBrush = $script:TitleBarBlueInactive
    }

    $btnCatConnection = $dlg.FindName('btnCatConnection')
    $btnCatAbout      = $dlg.FindName('btnCatAbout')
    $paneConnection   = $dlg.FindName('paneConnection')
    $paneAbout        = $dlg.FindName('paneAbout')
    $txtSiteCode      = $dlg.FindName('txtSiteCode')
    $txtSmsProvider   = $dlg.FindName('txtSmsProvider')
    $btnOk            = $dlg.FindName('btnOk')
    $btnCancel        = $dlg.FindName('btnCancel')

    $txtSiteCode.Text    = [string]$global:Prefs.SiteCode
    $txtSmsProvider.Text = [string]$global:Prefs.SMSProvider

    $btnCatConnection.Add_Click({
        $paneConnection.Visibility = [System.Windows.Visibility]::Visible
        $paneAbout.Visibility      = [System.Windows.Visibility]::Collapsed
    })
    $btnCatAbout.Add_Click({
        $paneConnection.Visibility = [System.Windows.Visibility]::Collapsed
        $paneAbout.Visibility      = [System.Windows.Visibility]::Visible
    })

    $btnOk.Add_Click({
        $newSite     = ([string]$txtSiteCode.Text).Trim()
        $newProvider = ([string]$txtSmsProvider.Text).Trim()
        $connectionChanged = ($newSite -ne [string]$global:Prefs.SiteCode) -or
                             ($newProvider -ne [string]$global:Prefs.SMSProvider)

        $global:Prefs.SiteCode    = $newSite
        $global:Prefs.SMSProvider = $newProvider
        Save-CmPreferences -Prefs $global:Prefs

        if ($connectionChanged) {
            Dispose-BgWork
            if ($script:BgRunspace) {
                try { $script:BgRunspace.Close() }   catch { $null = $_ }
                try { $script:BgRunspace.Dispose() } catch { $null = $_ }
                $script:BgRunspace = $null
            }
            $script:BgState           = $null
            $script:IsConnectedFromBg = $false
            $progressOverlay.Visibility = [System.Windows.Visibility]::Collapsed
            $btnRefresh.IsEnabled       = $true
        }

        $dlg.DialogResult = $true
        $dlg.Close()
    })
    $btnCancel.Add_Click({
        $dlg.DialogResult = $false
        $dlg.Close()
    })

    [void]$dlg.ShowDialog()
    Update-StatusBarSummary
}

$btnOptions.Add_Click({ Show-OptionsDialog })

# =============================================================================
# Window state persistence (with WinForms->WPF schema bridge).
# =============================================================================
$global:WindowStatePath = Join-Path $PSScriptRoot 'CollectionManager.windowstate.json'

function Save-WindowState {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification='Writes a small JSON state file; idempotent.')]
    param()
    try {
        $state = @{
            Left       = [int]$window.Left
            Top        = [int]$window.Top
            Width      = [int]$window.Width
            Height     = [int]$window.Height
            Maximized  = ($window.WindowState -eq [System.Windows.WindowState]::Maximized)
            ActiveView = $script:ActiveView
        }
        $state | ConvertTo-Json | Set-Content -LiteralPath $global:WindowStatePath -Encoding UTF8
    } catch { $null = $_ }
}

function Restore-WindowState {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification='Reads the JSON state file and applies geometry; idempotent.')]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseApprovedVerbs', '', Justification='Restore is intentional and reads as a single action.')]
    param()
    if (-not (Test-Path -LiteralPath $global:WindowStatePath)) { return }
    try {
        $s = Get-Content -LiteralPath $global:WindowStatePath -Raw | ConvertFrom-Json -ErrorAction Stop

        # Schema bridge: WinForms 1.0 used X/Y/ActiveTab; WPF refresh uses
        # Left/Top/ActiveView. Read both so legacy files don't snap the window
        # to (0,0) at MinSize after the upgrade.
        $left = if ($null -ne $s.Left) { [int]$s.Left } elseif ($null -ne $s.X) { [int]$s.X } else { $null }
        $top  = if ($null -ne $s.Top)  { [int]$s.Top  } elseif ($null -ne $s.Y) { [int]$s.Y } else { $null }
        $w    = if ($null -ne $s.Width)  { [int]$s.Width  } else { $null }
        $h    = if ($null -ne $s.Height) { [int]$s.Height } else { $null }

        if ($s.Maximized) {
            $window.WindowState = [System.Windows.WindowState]::Maximized
        } elseif ($null -ne $left -and $null -ne $top -and $null -ne $w -and $null -ne $h) {
            $screen = [System.Windows.Forms.Screen]::FromPoint([System.Drawing.Point]::new($left, $top))
            $bounds = $screen.WorkingArea
            $left = [Math]::Max($bounds.X, [Math]::Min($left, $bounds.Right - 200))
            $top  = [Math]::Max($bounds.Y, [Math]::Min($top,  $bounds.Bottom - 100))
            $window.Left   = $left
            $window.Top    = $top
            $window.Width  = [Math]::Max($window.MinWidth,  $w)
            $window.Height = [Math]::Max($window.MinHeight, $h)
        }

        if ($s.ActiveView -in @('Collections','WQL Editor','Templates')) {
            Set-ActiveView -View ([string]$s.ActiveView)
        }
    } catch { $null = $_ }
}

$window.Add_Closing({
    Save-WindowState
    Dispose-BgWork
    if ($script:BgRunspace) {
        try { $script:BgRunspace.Close() }  catch { $null = $_ }
        try { $script:BgRunspace.Dispose() } catch { $null = $_ }
    }
})

$window.Add_Loaded({
    Restore-WindowState

    # Apply user theme prefs AFTER the chrome has fully attached.
    $isDark = [bool]$global:Prefs['DarkMode']
    if (-not $isDark) {
        [void][ControlzEx.Theming.ThemeManager]::Current.ChangeTheme($window, 'Light.Blue')
    }
    Update-TitleBarBrushes

    Update-ActionBarVisibility
    Update-StatusBarSummary
    Add-LogLine 'Collection Manager ready. Configure Site / Provider in Options, then click Refresh.'

    # Auto-load templates so the Templates view is populated on first switch
    # without requiring an MECM connection.
    Invoke-LoadTemplates
})

# =============================================================================
# Run.
# =============================================================================
[void]$window.ShowDialog()
try { Stop-Transcript | Out-Null } catch { $null = $_ }
