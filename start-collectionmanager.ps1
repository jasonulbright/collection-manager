<#
.SYNOPSIS
    WinForms front-end for MECM Collection Manager with Offline WQL Editor.

.DESCRIPTION
    Provides a GUI for managing MECM device collections outside the console.
    Features offline WQL query editing, reusable templates (157 ready-made +
    20 parameterized), collection CRUD, direct member management, and bulk
    query rule editing.

.EXAMPLE
    .\start-collectionmanager.ps1

.NOTES
    Requirements:
      - PowerShell 5.1
      - .NET Framework 4.8+
      - Windows Forms (System.Windows.Forms)
      - Configuration Manager console installed

    ScriptName : start-collectionmanager.ps1
    Purpose    : WinForms front-end for MECM collection management
    Version    : 1.0.0
    Updated    : 2026-03-03
#>

param()

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

[System.Windows.Forms.Application]::EnableVisualStyles()
try { [System.Windows.Forms.Application]::SetCompatibleTextRenderingDefault($false) } catch { }

$moduleRoot = Join-Path $PSScriptRoot "Module"
Import-Module (Join-Path $moduleRoot "CollectionManagerCommon.psd1") -Force -DisableNameChecking

# Initialize tool logging
$toolLogFolder = Join-Path $PSScriptRoot "Logs"
if (-not (Test-Path -LiteralPath $toolLogFolder)) {
    New-Item -ItemType Directory -Path $toolLogFolder -Force | Out-Null
}
$toolLogPath = Join-Path $toolLogFolder ("CollMgr-{0}.log" -f (Get-Date -Format 'yyyyMMdd-HHmmss'))
Initialize-Logging -LogPath $toolLogPath

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

function Set-ModernButtonStyle {
    param(
        [Parameter(Mandatory)][System.Windows.Forms.Button]$Button,
        [Parameter(Mandatory)][System.Drawing.Color]$BackColor
    )
    $Button.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $Button.FlatAppearance.BorderSize = 0
    $Button.BackColor = $BackColor
    $Button.ForeColor = [System.Drawing.Color]::White
    $Button.UseVisualStyleBackColor = $false
    $Button.Cursor = [System.Windows.Forms.Cursors]::Hand
    $hover = [System.Drawing.Color]::FromArgb([Math]::Max(0, $BackColor.R - 18), [Math]::Max(0, $BackColor.G - 18), [Math]::Max(0, $BackColor.B - 18))
    $down  = [System.Drawing.Color]::FromArgb([Math]::Max(0, $BackColor.R - 36), [Math]::Max(0, $BackColor.G - 36), [Math]::Max(0, $BackColor.B - 36))
    $Button.FlatAppearance.MouseOverBackColor = $hover
    $Button.FlatAppearance.MouseDownBackColor = $down
}

function Enable-DoubleBuffer {
    param([Parameter(Mandatory)][System.Windows.Forms.Control]$Control)
    $prop = $Control.GetType().GetProperty("DoubleBuffered", [System.Reflection.BindingFlags] "Instance,NonPublic")
    if ($prop) { $prop.SetValue($Control, $true, $null) | Out-Null }
}

function Add-LogLine {
    param([Parameter(Mandatory)][System.Windows.Forms.TextBox]$TextBox, [Parameter(Mandatory)][string]$Message)
    $ts = (Get-Date).ToString("HH:mm:ss")
    $line = "{0}  {1}" -f $ts, $Message
    if ([string]::IsNullOrWhiteSpace($TextBox.Text)) { $TextBox.Text = $line }
    else { $TextBox.AppendText([Environment]::NewLine + $line) }
    $TextBox.SelectionStart = $TextBox.TextLength
    $TextBox.ScrollToCaret()
}

function Save-WindowState {
    $statePath = Join-Path $PSScriptRoot "CollectionManager.windowstate.json"
    $state = @{
        X = $form.Location.X; Y = $form.Location.Y
        Width = $form.Size.Width; Height = $form.Size.Height
        Maximized = ($form.WindowState -eq [System.Windows.Forms.FormWindowState]::Maximized)
        ActiveTab = $tabMain.SelectedIndex
    }
    $state | ConvertTo-Json | Set-Content -LiteralPath $statePath -Encoding UTF8
}

function Restore-WindowState {
    $statePath = Join-Path $PSScriptRoot "CollectionManager.windowstate.json"
    if (-not (Test-Path -LiteralPath $statePath)) { return }
    try {
        $state = Get-Content -LiteralPath $statePath -Raw | ConvertFrom-Json
        if ($state.Maximized) { $form.WindowState = [System.Windows.Forms.FormWindowState]::Maximized }
        else {
            $screen = [System.Windows.Forms.Screen]::FromPoint((New-Object System.Drawing.Point($state.X, $state.Y)))
            $bounds = $screen.WorkingArea
            $x = [Math]::Max($bounds.X, [Math]::Min($state.X, $bounds.Right - 200))
            $y = [Math]::Max($bounds.Y, [Math]::Min($state.Y, $bounds.Bottom - 100))
            $form.Location = New-Object System.Drawing.Point($x, $y)
            $form.Size = New-Object System.Drawing.Size([Math]::Max($form.MinimumSize.Width, $state.Width), [Math]::Max($form.MinimumSize.Height, $state.Height))
        }
        if ($null -ne $state.ActiveTab -and $state.ActiveTab -ge 0 -and $state.ActiveTab -lt $tabMain.TabCount) {
            $tabMain.SelectedIndex = [int]$state.ActiveTab
        }
    } catch { }
}

# ---------------------------------------------------------------------------
# Preferences
# ---------------------------------------------------------------------------

function Get-CmPreferences {
    $prefsPath = Join-Path $PSScriptRoot "CollectionManager.prefs.json"
    $defaults = @{ DarkMode = $false; SiteCode = ''; SMSProvider = '' }
    if (Test-Path -LiteralPath $prefsPath) {
        try {
            $loaded = Get-Content -LiteralPath $prefsPath -Raw | ConvertFrom-Json
            if ($null -ne $loaded.DarkMode) { $defaults.DarkMode = [bool]$loaded.DarkMode }
            if ($loaded.SiteCode)           { $defaults.SiteCode = $loaded.SiteCode }
            if ($loaded.SMSProvider)         { $defaults.SMSProvider = $loaded.SMSProvider }
        } catch { }
    }
    return $defaults
}

function Save-CmPreferences {
    param([hashtable]$Prefs)
    $prefsPath = Join-Path $PSScriptRoot "CollectionManager.prefs.json"
    $Prefs | ConvertTo-Json | Set-Content -LiteralPath $prefsPath -Encoding UTF8
}

$script:Prefs = Get-CmPreferences

# ---------------------------------------------------------------------------
# Colors (theme-aware)
# ---------------------------------------------------------------------------

$clrAccent = [System.Drawing.Color]::FromArgb(0, 120, 212)

if ($script:Prefs.DarkMode) {
    $clrFormBg   = [System.Drawing.Color]::FromArgb(30, 30, 30)
    $clrPanelBg  = [System.Drawing.Color]::FromArgb(40, 40, 40)
    $clrHint     = [System.Drawing.Color]::FromArgb(140, 140, 140)
    $clrSubtitle = [System.Drawing.Color]::FromArgb(180, 200, 220)
    $clrGridAlt  = [System.Drawing.Color]::FromArgb(48, 48, 48)
    $clrGridLine = [System.Drawing.Color]::FromArgb(60, 60, 60)
    $clrDetailBg = [System.Drawing.Color]::FromArgb(45, 45, 45)
    $clrSepLine  = [System.Drawing.Color]::FromArgb(55, 55, 55)
    $clrLogBg    = [System.Drawing.Color]::FromArgb(35, 35, 35)
    $clrLogFg    = [System.Drawing.Color]::FromArgb(200, 200, 200)
    $clrText     = [System.Drawing.Color]::FromArgb(220, 220, 220)
    $clrGridText = [System.Drawing.Color]::FromArgb(220, 220, 220)
    $clrErrText  = [System.Drawing.Color]::FromArgb(255, 100, 100)
    $clrWarnText = [System.Drawing.Color]::FromArgb(255, 200, 80)
    $clrOkText   = [System.Drawing.Color]::FromArgb(80, 200, 80)
    $clrInfoText = [System.Drawing.Color]::FromArgb(100, 180, 255)
    $clrCardBlue = [System.Drawing.Color]::FromArgb(25, 40, 60)
    $clrTreeBg   = [System.Drawing.Color]::FromArgb(38, 38, 38)
} else {
    $clrFormBg   = [System.Drawing.Color]::FromArgb(245, 246, 248)
    $clrPanelBg  = [System.Drawing.Color]::White
    $clrHint     = [System.Drawing.Color]::FromArgb(140, 140, 140)
    $clrSubtitle = [System.Drawing.Color]::FromArgb(220, 230, 245)
    $clrGridAlt  = [System.Drawing.Color]::FromArgb(248, 250, 252)
    $clrGridLine = [System.Drawing.Color]::FromArgb(230, 230, 230)
    $clrDetailBg = [System.Drawing.Color]::FromArgb(250, 250, 250)
    $clrSepLine  = [System.Drawing.Color]::FromArgb(218, 220, 224)
    $clrLogBg    = [System.Drawing.Color]::White
    $clrLogFg    = [System.Drawing.Color]::Black
    $clrText     = [System.Drawing.Color]::Black
    $clrGridText = [System.Drawing.Color]::Black
    $clrErrText  = [System.Drawing.Color]::FromArgb(180, 0, 0)
    $clrWarnText = [System.Drawing.Color]::FromArgb(180, 120, 0)
    $clrOkText   = [System.Drawing.Color]::FromArgb(34, 139, 34)
    $clrInfoText = [System.Drawing.Color]::FromArgb(0, 100, 180)
    $clrCardBlue = [System.Drawing.Color]::FromArgb(220, 235, 255)
    $clrTreeBg   = [System.Drawing.Color]::White
}

# Dark mode ToolStrip renderer
if ($script:Prefs.DarkMode) {
    if (-not ('DarkToolStripRenderer' -as [type])) {
        $rendererCs = @(
            'using System.Drawing;', 'using System.Windows.Forms;',
            'public class DarkToolStripRenderer : ToolStripProfessionalRenderer {',
            '    private Color _bg;',
            '    public DarkToolStripRenderer(Color bg) : base() { _bg = bg; }',
            '    protected override void OnRenderToolStripBorder(ToolStripRenderEventArgs e) { }',
            '    protected override void OnRenderToolStripBackground(ToolStripRenderEventArgs e) { using (var b = new SolidBrush(_bg)) { e.Graphics.FillRectangle(b, e.AffectedBounds); } }',
            '    protected override void OnRenderMenuItemBackground(ToolStripItemRenderEventArgs e) { if (e.Item.Selected || e.Item.Pressed) { using (var b = new SolidBrush(Color.FromArgb(60, 60, 60))) { e.Graphics.FillRectangle(b, new Rectangle(Point.Empty, e.Item.Size)); } } }',
            '    protected override void OnRenderSeparator(ToolStripSeparatorRenderEventArgs e) { int y = e.Item.Height / 2; using (var p = new Pen(Color.FromArgb(70, 70, 70))) { e.Graphics.DrawLine(p, 0, y, e.Item.Width, y); } }',
            '    protected override void OnRenderImageMargin(ToolStripRenderEventArgs e) { using (var b = new SolidBrush(_bg)) { e.Graphics.FillRectangle(b, e.AffectedBounds); } }',
            '}'
        ) -join "`r`n"
        Add-Type -ReferencedAssemblies System.Windows.Forms, System.Drawing -TypeDefinition $rendererCs
    }
    $script:DarkRenderer = New-Object DarkToolStripRenderer($clrPanelBg)
}

# ---------------------------------------------------------------------------
# Preferences dialog
# ---------------------------------------------------------------------------

function Show-PreferencesDialog {
    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = "Preferences"; $dlg.Size = New-Object System.Drawing.Size(440, 300)
    $dlg.MinimumSize = $dlg.Size; $dlg.MaximumSize = $dlg.Size
    $dlg.StartPosition = "CenterParent"; $dlg.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $dlg.MaximizeBox = $false; $dlg.MinimizeBox = $false; $dlg.ShowInTaskbar = $false
    $dlg.Font = New-Object System.Drawing.Font("Segoe UI", 9.5); $dlg.BackColor = $clrFormBg

    $grpApp = New-Object System.Windows.Forms.GroupBox
    $grpApp.Text = "Appearance"; $grpApp.SetBounds(16, 12, 392, 60)
    $grpApp.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $grpApp.ForeColor = $clrText; $grpApp.BackColor = $clrFormBg
    if ($script:Prefs.DarkMode) { $grpApp.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat; $grpApp.ForeColor = $clrSepLine }
    $dlg.Controls.Add($grpApp)

    $chkDark = New-Object System.Windows.Forms.CheckBox
    $chkDark.Text = "Enable dark mode (requires restart)"; $chkDark.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $chkDark.AutoSize = $true; $chkDark.Location = New-Object System.Drawing.Point(14, 24)
    $chkDark.Checked = $script:Prefs.DarkMode; $chkDark.ForeColor = $clrText; $chkDark.BackColor = $clrFormBg
    if ($script:Prefs.DarkMode) { $chkDark.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat; $chkDark.ForeColor = [System.Drawing.Color]::FromArgb(170, 170, 170) }
    $grpApp.Controls.Add($chkDark)

    $grpConn = New-Object System.Windows.Forms.GroupBox
    $grpConn.Text = "MECM Connection"; $grpConn.SetBounds(16, 82, 392, 110)
    $grpConn.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $grpConn.ForeColor = $clrText; $grpConn.BackColor = $clrFormBg
    if ($script:Prefs.DarkMode) { $grpConn.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat; $grpConn.ForeColor = $clrSepLine }
    $dlg.Controls.Add($grpConn)

    $lblSC = New-Object System.Windows.Forms.Label; $lblSC.Text = "Site Code:"; $lblSC.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $lblSC.Location = New-Object System.Drawing.Point(14, 30); $lblSC.AutoSize = $true; $lblSC.ForeColor = $clrText
    $grpConn.Controls.Add($lblSC)
    $txtSC = New-Object System.Windows.Forms.TextBox; $txtSC.SetBounds(130, 27, 80, 24); $txtSC.Text = $script:Prefs.SiteCode
    $txtSC.Font = New-Object System.Drawing.Font("Segoe UI", 9); $txtSC.BackColor = $clrDetailBg; $txtSC.ForeColor = $clrText
    $grpConn.Controls.Add($txtSC)

    $lblSP = New-Object System.Windows.Forms.Label; $lblSP.Text = "SMS Provider:"; $lblSP.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $lblSP.Location = New-Object System.Drawing.Point(14, 64); $lblSP.AutoSize = $true; $lblSP.ForeColor = $clrText
    $grpConn.Controls.Add($lblSP)
    $txtSP = New-Object System.Windows.Forms.TextBox; $txtSP.SetBounds(130, 61, 240, 24); $txtSP.Text = $script:Prefs.SMSProvider
    $txtSP.Font = New-Object System.Drawing.Font("Segoe UI", 9); $txtSP.BackColor = $clrDetailBg; $txtSP.ForeColor = $clrText
    $grpConn.Controls.Add($txtSP)

    $btnSave = New-Object System.Windows.Forms.Button; $btnSave.Text = "Save"; $btnSave.SetBounds(220, 210, 90, 32)
    $btnSave.Font = New-Object System.Drawing.Font("Segoe UI", 9); Set-ModernButtonStyle -Button $btnSave -BackColor $clrAccent
    $dlg.Controls.Add($btnSave)
    $btnCancel = New-Object System.Windows.Forms.Button; $btnCancel.Text = "Cancel"; $btnCancel.SetBounds(318, 210, 90, 32)
    $btnCancel.Font = New-Object System.Drawing.Font("Segoe UI", 9); $btnCancel.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnCancel.ForeColor = $clrText; $btnCancel.BackColor = $clrFormBg
    $dlg.Controls.Add($btnCancel)

    $btnSave.Add_Click({
        $script:Prefs.DarkMode = $chkDark.Checked; $script:Prefs.SiteCode = $txtSC.Text.Trim(); $script:Prefs.SMSProvider = $txtSP.Text.Trim()
        Save-CmPreferences -Prefs $script:Prefs
        $lblSiteVal.Text = if ($script:Prefs.SiteCode) { $script:Prefs.SiteCode } else { "(not set)" }
        $lblProviderVal.Text = if ($script:Prefs.SMSProvider) { $script:Prefs.SMSProvider } else { "(not set)" }
        $dlg.DialogResult = [System.Windows.Forms.DialogResult]::OK; $dlg.Close()
    })
    $btnCancel.Add_Click({ $dlg.DialogResult = [System.Windows.Forms.DialogResult]::Cancel; $dlg.Close() })
    $dlg.AcceptButton = $btnSave; $dlg.CancelButton = $btnCancel
    $dlg.ShowDialog($form) | Out-Null; $dlg.Dispose()
}

# ---------------------------------------------------------------------------
# Form
# ---------------------------------------------------------------------------

$form = New-Object System.Windows.Forms.Form
$form.Text = "Collection Manager"; $form.Size = New-Object System.Drawing.Size(1360, 900)
$form.MinimumSize = New-Object System.Drawing.Size(1000, 650); $form.StartPosition = "CenterScreen"
$form.Font = New-Object System.Drawing.Font("Segoe UI", 9.5); $form.BackColor = $clrFormBg

# ---------------------------------------------------------------------------
# StatusStrip (Dock:Bottom)
# ---------------------------------------------------------------------------

$statusStrip = New-Object System.Windows.Forms.StatusStrip
$statusStrip.BackColor = $clrPanelBg; $statusStrip.ForeColor = $clrText; $statusStrip.SizingGrip = $false
if ($script:Prefs.DarkMode -and $script:DarkRenderer) { $statusStrip.Renderer = $script:DarkRenderer }
$statusLabel = New-Object System.Windows.Forms.ToolStripStatusLabel; $statusLabel.Text = "Disconnected"
$statusLabel.Spring = $true; $statusLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft; $statusLabel.ForeColor = $clrText
$statusStrip.Items.Add($statusLabel) | Out-Null
$statusRowCount = New-Object System.Windows.Forms.ToolStripStatusLabel; $statusRowCount.Text = ""; $statusRowCount.ForeColor = $clrHint
$statusStrip.Items.Add($statusRowCount) | Out-Null
$form.Controls.Add($statusStrip)

# ---------------------------------------------------------------------------
# Log console (Dock:Bottom)
# ---------------------------------------------------------------------------

$pnlLog = New-Object System.Windows.Forms.Panel; $pnlLog.Dock = [System.Windows.Forms.DockStyle]::Bottom
$pnlLog.Height = 95; $pnlLog.BackColor = $clrLogBg; $form.Controls.Add($pnlLog)
$txtLog = New-Object System.Windows.Forms.TextBox; $txtLog.Dock = [System.Windows.Forms.DockStyle]::Fill
$txtLog.Multiline = $true; $txtLog.ReadOnly = $true; $txtLog.WordWrap = $true
$txtLog.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical
$txtLog.Font = New-Object System.Drawing.Font("Consolas", 9); $txtLog.BackColor = $clrLogBg; $txtLog.ForeColor = $clrLogFg
$txtLog.BorderStyle = [System.Windows.Forms.BorderStyle]::None; $pnlLog.Controls.Add($txtLog)
$pnlLogSep = New-Object System.Windows.Forms.Panel; $pnlLogSep.Dock = [System.Windows.Forms.DockStyle]::Bottom
$pnlLogSep.Height = 1; $pnlLogSep.BackColor = $clrSepLine; $form.Controls.Add($pnlLogSep)

# ---------------------------------------------------------------------------
# Button panel (Dock:Bottom)
# ---------------------------------------------------------------------------

$pnlButtons = New-Object System.Windows.Forms.Panel; $pnlButtons.Dock = [System.Windows.Forms.DockStyle]::Bottom
$pnlButtons.Height = 56; $pnlButtons.BackColor = $clrPanelBg
$pnlButtons.Padding = New-Object System.Windows.Forms.Padding(12, 10, 12, 10); $form.Controls.Add($pnlButtons)
$flowButtons = New-Object System.Windows.Forms.FlowLayoutPanel; $flowButtons.Dock = [System.Windows.Forms.DockStyle]::Fill
$flowButtons.FlowDirection = [System.Windows.Forms.FlowDirection]::LeftToRight; $flowButtons.WrapContents = $false
$flowButtons.BackColor = $clrPanelBg; $pnlButtons.Controls.Add($flowButtons)

$btnExportCsv = New-Object System.Windows.Forms.Button; $btnExportCsv.Text = "Export CSV"
$btnExportCsv.Size = New-Object System.Drawing.Size(120, 34); $btnExportCsv.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$btnExportCsv.Margin = New-Object System.Windows.Forms.Padding(0, 0, 8, 0)
Set-ModernButtonStyle -Button $btnExportCsv -BackColor ([System.Drawing.Color]::FromArgb(50, 130, 50))
$btnExportHtml = New-Object System.Windows.Forms.Button; $btnExportHtml.Text = "Export HTML"
$btnExportHtml.Size = New-Object System.Drawing.Size(120, 34); $btnExportHtml.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$btnExportHtml.Margin = New-Object System.Windows.Forms.Padding(0, 0, 8, 0)
Set-ModernButtonStyle -Button $btnExportHtml -BackColor ([System.Drawing.Color]::FromArgb(50, 130, 50))
$flowButtons.Controls.Add($btnExportCsv); $flowButtons.Controls.Add($btnExportHtml)
$pnlBtnSep = New-Object System.Windows.Forms.Panel; $pnlBtnSep.Dock = [System.Windows.Forms.DockStyle]::Bottom
$pnlBtnSep.Height = 1; $pnlBtnSep.BackColor = $clrSepLine; $form.Controls.Add($pnlBtnSep)

# ---------------------------------------------------------------------------
# MenuStrip
# ---------------------------------------------------------------------------

$menuStrip = New-Object System.Windows.Forms.MenuStrip; $menuStrip.BackColor = $clrPanelBg; $menuStrip.ForeColor = $clrText
if ($script:Prefs.DarkMode -and $script:DarkRenderer) { $menuStrip.Renderer = $script:DarkRenderer }

$mnuFile = New-Object System.Windows.Forms.ToolStripMenuItem("&File"); $mnuFile.ForeColor = $clrText
$mnuPrefs = New-Object System.Windows.Forms.ToolStripMenuItem("&Preferences...")
$mnuPrefs.ShortcutKeys = [System.Windows.Forms.Keys]::Control -bor [System.Windows.Forms.Keys]::Oemcomma
$mnuPrefs.ForeColor = $clrText; $mnuPrefs.Add_Click({ Show-PreferencesDialog })
$mnuFile.DropDownItems.Add($mnuPrefs) | Out-Null
$mnuFile.DropDownItems.Add((New-Object System.Windows.Forms.ToolStripSeparator)) | Out-Null
$mnuExit = New-Object System.Windows.Forms.ToolStripMenuItem("E&xit"); $mnuExit.ForeColor = $clrText
$mnuExit.Add_Click({ $form.Close() }); $mnuFile.DropDownItems.Add($mnuExit) | Out-Null
$menuStrip.Items.Add($mnuFile) | Out-Null

$mnuView = New-Object System.Windows.Forms.ToolStripMenuItem("&View"); $mnuView.ForeColor = $clrText
$tabNames = @('Collections', 'WQL Editor', 'Create / Clone', 'Direct Members', 'Templates')
for ($idx = 0; $idx -lt $tabNames.Count; $idx++) {
    $mi = New-Object System.Windows.Forms.ToolStripMenuItem($tabNames[$idx]); $mi.ForeColor = $clrText; $mi.Tag = $idx
    $mi.Add_Click({ $tabMain.SelectedIndex = [int]$this.Tag }.GetNewClosure())
    $mnuView.DropDownItems.Add($mi) | Out-Null
}
$menuStrip.Items.Add($mnuView) | Out-Null

$mnuHelp = New-Object System.Windows.Forms.ToolStripMenuItem("&Help"); $mnuHelp.ForeColor = $clrText
$mnuAbout = New-Object System.Windows.Forms.ToolStripMenuItem("&About"); $mnuAbout.ForeColor = $clrText
$mnuAbout.Add_Click({
    [System.Windows.Forms.MessageBox]::Show("Collection Manager v1.0.0`r`n`r`nOffline WQL editor and collection management for MECM device collections.`r`n157 ready-made + 20 parameterized query templates.`r`n`r`nRequires: ConfigMgr console.", "About", "OK", "Information") | Out-Null
})
$mnuHelp.DropDownItems.Add($mnuAbout) | Out-Null; $menuStrip.Items.Add($mnuHelp) | Out-Null

# ---------------------------------------------------------------------------
# Header panel (Dock:Top)
# ---------------------------------------------------------------------------

$pnlHeader = New-Object System.Windows.Forms.Panel; $pnlHeader.Dock = [System.Windows.Forms.DockStyle]::Top
$pnlHeader.Height = 60; $pnlHeader.BackColor = $clrAccent
$lblTitle = New-Object System.Windows.Forms.Label; $lblTitle.Text = "Collection Manager"
$lblTitle.Font = New-Object System.Drawing.Font("Segoe UI", 16, [System.Drawing.FontStyle]::Bold)
$lblTitle.ForeColor = [System.Drawing.Color]::White; $lblTitle.AutoSize = $true
$lblTitle.Location = New-Object System.Drawing.Point(16, 6); $lblTitle.BackColor = [System.Drawing.Color]::Transparent
$pnlHeader.Controls.Add($lblTitle)
$lblSubtitle = New-Object System.Windows.Forms.Label; $lblSubtitle.Text = "Offline WQL Editor & Collection Management"
$lblSubtitle.Font = New-Object System.Drawing.Font("Segoe UI", 9); $lblSubtitle.ForeColor = $clrSubtitle
$lblSubtitle.AutoSize = $true; $lblSubtitle.Location = New-Object System.Drawing.Point(18, 36)
$lblSubtitle.BackColor = [System.Drawing.Color]::Transparent; $pnlHeader.Controls.Add($lblSubtitle)
$form.Controls.Add($pnlHeader)

# ---------------------------------------------------------------------------
# Connection bar (Dock:Top)
# ---------------------------------------------------------------------------

$pnlConnBar = New-Object System.Windows.Forms.Panel; $pnlConnBar.Dock = [System.Windows.Forms.DockStyle]::Top
$pnlConnBar.Height = 40; $pnlConnBar.BackColor = $clrPanelBg
$pnlConnBar.Padding = New-Object System.Windows.Forms.Padding(12, 6, 12, 6); $form.Controls.Add($pnlConnBar)
$flowConn = New-Object System.Windows.Forms.FlowLayoutPanel; $flowConn.Dock = [System.Windows.Forms.DockStyle]::Fill
$flowConn.FlowDirection = [System.Windows.Forms.FlowDirection]::LeftToRight; $flowConn.WrapContents = $false
$flowConn.BackColor = $clrPanelBg; $pnlConnBar.Controls.Add($flowConn)

$lblSite = New-Object System.Windows.Forms.Label; $lblSite.Text = "Site:"; $lblSite.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$lblSite.AutoSize = $true; $lblSite.Margin = New-Object System.Windows.Forms.Padding(0, 5, 2, 0); $lblSite.ForeColor = $clrText; $lblSite.BackColor = $clrPanelBg
$flowConn.Controls.Add($lblSite)
$lblSiteVal = New-Object System.Windows.Forms.Label; $lblSiteVal.Text = if ($script:Prefs.SiteCode) { $script:Prefs.SiteCode } else { "(not set)" }
$lblSiteVal.Font = New-Object System.Drawing.Font("Segoe UI", 9); $lblSiteVal.AutoSize = $true
$lblSiteVal.Margin = New-Object System.Windows.Forms.Padding(0, 5, 16, 0); $lblSiteVal.ForeColor = $clrHint; $lblSiteVal.BackColor = $clrPanelBg
$flowConn.Controls.Add($lblSiteVal)
$lblProvider = New-Object System.Windows.Forms.Label; $lblProvider.Text = "Provider:"; $lblProvider.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$lblProvider.AutoSize = $true; $lblProvider.Margin = New-Object System.Windows.Forms.Padding(0, 5, 2, 0); $lblProvider.ForeColor = $clrText; $lblProvider.BackColor = $clrPanelBg
$flowConn.Controls.Add($lblProvider)
$lblProviderVal = New-Object System.Windows.Forms.Label; $lblProviderVal.Text = if ($script:Prefs.SMSProvider) { $script:Prefs.SMSProvider } else { "(not set)" }
$lblProviderVal.Font = New-Object System.Drawing.Font("Segoe UI", 9); $lblProviderVal.AutoSize = $true
$lblProviderVal.Margin = New-Object System.Windows.Forms.Padding(0, 5, 24, 0); $lblProviderVal.ForeColor = $clrHint; $lblProviderVal.BackColor = $clrPanelBg
$flowConn.Controls.Add($lblProviderVal)

$btnLoad = New-Object System.Windows.Forms.Button; $btnLoad.Text = "Load Collections"
$btnLoad.Size = New-Object System.Drawing.Size(150, 26); $btnLoad.Font = New-Object System.Drawing.Font("Segoe UI", 9)
Set-ModernButtonStyle -Button $btnLoad -BackColor $clrAccent; $flowConn.Controls.Add($btnLoad)

$pnlSep1 = New-Object System.Windows.Forms.Panel; $pnlSep1.Dock = [System.Windows.Forms.DockStyle]::Top
$pnlSep1.Height = 1; $pnlSep1.BackColor = $clrSepLine; $form.Controls.Add($pnlSep1)

# ---------------------------------------------------------------------------
# Summary cards (Dock:Top)
# ---------------------------------------------------------------------------

$pnlCards = New-Object System.Windows.Forms.Panel; $pnlCards.Dock = [System.Windows.Forms.DockStyle]::Top
$pnlCards.Height = 56; $pnlCards.BackColor = $clrFormBg
$pnlCards.Padding = New-Object System.Windows.Forms.Padding(8, 4, 8, 4); $form.Controls.Add($pnlCards)
$flowCards = New-Object System.Windows.Forms.FlowLayoutPanel; $flowCards.Dock = [System.Windows.Forms.DockStyle]::Fill
$flowCards.FlowDirection = [System.Windows.Forms.FlowDirection]::LeftToRight; $flowCards.WrapContents = $false
$flowCards.BackColor = $clrFormBg; $pnlCards.Controls.Add($flowCards)

function New-SummaryCard { param([string]$Title, [int]$TabIndex)
    $card = New-Object System.Windows.Forms.Panel; $card.Size = New-Object System.Drawing.Size(220, 44)
    $card.Margin = New-Object System.Windows.Forms.Padding(4, 2, 4, 2); $card.BackColor = $clrPanelBg
    $card.BorderStyle = [System.Windows.Forms.BorderStyle]::None; $card.Cursor = [System.Windows.Forms.Cursors]::Hand; $card.Tag = $TabIndex
    $bar = New-Object System.Windows.Forms.Panel; $bar.Dock = [System.Windows.Forms.DockStyle]::Left; $bar.Width = 4; $bar.BackColor = $clrHint; $card.Controls.Add($bar)
    $lt = New-Object System.Windows.Forms.Label; $lt.Text = $Title; $lt.Font = New-Object System.Drawing.Font("Segoe UI", 8.5, [System.Drawing.FontStyle]::Bold)
    $lt.ForeColor = $clrText; $lt.AutoSize = $true; $lt.Location = New-Object System.Drawing.Point(10, 4); $lt.BackColor = [System.Drawing.Color]::Transparent; $card.Controls.Add($lt)
    $lv = New-Object System.Windows.Forms.Label; $lv.Text = "--"; $lv.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $lv.ForeColor = $clrHint; $lv.AutoSize = $true; $lv.Location = New-Object System.Drawing.Point(10, 22); $lv.BackColor = [System.Drawing.Color]::Transparent; $lv.Tag = "value"; $card.Controls.Add($lv)
    $ch = { $tabMain.SelectedIndex = [int]$this.Parent.Tag }; $cch = { $tabMain.SelectedIndex = [int]$this.Tag }
    $card.Add_Click($cch); $lt.Add_Click($ch); $lv.Add_Click($ch)
    return $card
}

$cardCollections = New-SummaryCard -Title "Total Collections"  -TabIndex 0
$cardQueryRules  = New-SummaryCard -Title "Query Rules"        -TabIndex 1
$cardDirect      = New-SummaryCard -Title "Direct Members"     -TabIndex 3
$cardInclExcl    = New-SummaryCard -Title "Include/Exclude"    -TabIndex 0
$flowCards.Controls.Add($cardCollections); $flowCards.Controls.Add($cardQueryRules)
$flowCards.Controls.Add($cardDirect); $flowCards.Controls.Add($cardInclExcl)

function Update-Card { param([System.Windows.Forms.Panel]$Card, [string]$ValueText, [string]$Severity)
    $bar = $Card.Controls[0]; $vl = $Card.Controls | Where-Object { $_.Tag -eq 'value' }
    switch ($Severity) {
        'info'  { $bar.BackColor = $clrInfoText; $Card.BackColor = $clrCardBlue; if ($vl) { $vl.ForeColor = $clrInfoText } }
        'ok'    { $bar.BackColor = $clrOkText;   $Card.BackColor = $clrFormBg;   if ($vl) { $vl.ForeColor = $clrOkText } }
        default { $bar.BackColor = $clrHint;     $Card.BackColor = $clrPanelBg;  if ($vl) { $vl.ForeColor = $clrHint } }
    }
    if ($vl) { $vl.Text = $ValueText }
}

$pnlSep2 = New-Object System.Windows.Forms.Panel; $pnlSep2.Dock = [System.Windows.Forms.DockStyle]::Top
$pnlSep2.Height = 1; $pnlSep2.BackColor = $clrSepLine; $form.Controls.Add($pnlSep2)

# ---------------------------------------------------------------------------
# Filter bar (Dock:Top)
# ---------------------------------------------------------------------------

$pnlFilter = New-Object System.Windows.Forms.Panel; $pnlFilter.Dock = [System.Windows.Forms.DockStyle]::Top
$pnlFilter.Height = 44; $pnlFilter.BackColor = $clrPanelBg
$pnlFilter.Padding = New-Object System.Windows.Forms.Padding(12, 6, 12, 6); $form.Controls.Add($pnlFilter)
$flowFilter = New-Object System.Windows.Forms.FlowLayoutPanel; $flowFilter.Dock = [System.Windows.Forms.DockStyle]::Fill
$flowFilter.FlowDirection = [System.Windows.Forms.FlowDirection]::LeftToRight; $flowFilter.WrapContents = $false
$flowFilter.BackColor = $clrPanelBg; $pnlFilter.Controls.Add($flowFilter)

$lblFilt = New-Object System.Windows.Forms.Label; $lblFilt.Text = "Filter:"; $lblFilt.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$lblFilt.AutoSize = $true; $lblFilt.Margin = New-Object System.Windows.Forms.Padding(0, 6, 4, 0); $lblFilt.ForeColor = $clrText; $lblFilt.BackColor = $clrPanelBg
$flowFilter.Controls.Add($lblFilt)
$txtFilter = New-Object System.Windows.Forms.TextBox; $txtFilter.Font = New-Object System.Drawing.Font("Segoe UI", 9.5)
$txtFilter.Width = 300; $txtFilter.Margin = New-Object System.Windows.Forms.Padding(0, 2, 16, 0)
$txtFilter.BorderStyle = if ($script:Prefs.DarkMode) { [System.Windows.Forms.BorderStyle]::None } else { [System.Windows.Forms.BorderStyle]::FixedSingle }
$txtFilter.BackColor = $clrDetailBg; $txtFilter.ForeColor = $clrText; $flowFilter.Controls.Add($txtFilter)

$pnlSep3 = New-Object System.Windows.Forms.Panel; $pnlSep3.Dock = [System.Windows.Forms.DockStyle]::Top
$pnlSep3.Height = 1; $pnlSep3.BackColor = $clrSepLine; $form.Controls.Add($pnlSep3)

# ---------------------------------------------------------------------------
# Themed DataGridView helper
# ---------------------------------------------------------------------------

function New-ThemedGrid { param([switch]$MultiSelect)
    $g = New-Object System.Windows.Forms.DataGridView; $g.Dock = [System.Windows.Forms.DockStyle]::Fill
    $g.ReadOnly = $true; $g.AllowUserToAddRows = $false; $g.AllowUserToDeleteRows = $false; $g.AllowUserToResizeRows = $false
    $g.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect; $g.MultiSelect = [bool]$MultiSelect
    $g.AutoGenerateColumns = $false; $g.RowHeadersVisible = $false; $g.BackgroundColor = $clrPanelBg
    $g.BorderStyle = [System.Windows.Forms.BorderStyle]::None; $g.CellBorderStyle = [System.Windows.Forms.DataGridViewCellBorderStyle]::SingleHorizontal
    $g.GridColor = $clrGridLine; $g.ColumnHeadersDefaultCellStyle.BackColor = $clrAccent
    $g.ColumnHeadersDefaultCellStyle.ForeColor = [System.Drawing.Color]::White
    $g.ColumnHeadersDefaultCellStyle.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Bold)
    $g.ColumnHeadersDefaultCellStyle.Padding = New-Object System.Windows.Forms.Padding(4)
    $g.ColumnHeadersDefaultCellStyle.WrapMode = [System.Windows.Forms.DataGridViewTriState]::False
    $g.ColumnHeadersHeightSizeMode = [System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode]::DisableResizing
    $g.ColumnHeadersHeight = 32; $g.ColumnHeadersBorderStyle = [System.Windows.Forms.DataGridViewHeaderBorderStyle]::Single
    $g.EnableHeadersVisualStyles = $false; $g.DefaultCellStyle.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $g.DefaultCellStyle.ForeColor = $clrGridText; $g.DefaultCellStyle.BackColor = $clrPanelBg
    $g.DefaultCellStyle.Padding = New-Object System.Windows.Forms.Padding(2)
    $g.DefaultCellStyle.SelectionBackColor = if ($script:Prefs.DarkMode) { [System.Drawing.Color]::FromArgb(38, 79, 120) } else { [System.Drawing.Color]::FromArgb(0, 120, 215) }
    $g.DefaultCellStyle.SelectionForeColor = [System.Drawing.Color]::White
    $g.RowTemplate.Height = 26; $g.AlternatingRowsDefaultCellStyle.BackColor = $clrGridAlt
    Enable-DoubleBuffer -Control $g; return $g
}

# ---------------------------------------------------------------------------
# TabControl (Fill)
# ---------------------------------------------------------------------------

$tabMain = New-Object System.Windows.Forms.TabControl; $tabMain.Dock = [System.Windows.Forms.DockStyle]::Fill
$tabMain.Font = New-Object System.Drawing.Font("Segoe UI", 9.5, [System.Drawing.FontStyle]::Bold)
$tabMain.SizeMode = [System.Windows.Forms.TabSizeMode]::Fixed
$tabMain.ItemSize = New-Object System.Drawing.Size(130, 30)
$tabMain.DrawMode = [System.Windows.Forms.TabDrawMode]::OwnerDrawFixed
$tabMain.Add_DrawItem({
    param($s, $e)
    $e.Graphics.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::ClearTypeGridFit
    $tab = $s.TabPages[$e.Index]; $sel = ($s.SelectedIndex -eq $e.Index)
    $bg = if ($script:Prefs.DarkMode) { if ($sel) { $clrAccent } else { $clrPanelBg } } else { if ($sel) { $clrAccent } else { [System.Drawing.Color]::FromArgb(240, 240, 240) } }
    $fg = if ($sel) { [System.Drawing.Color]::White } else { $clrText }
    $bb = New-Object System.Drawing.SolidBrush($bg); $e.Graphics.FillRectangle($bb, $e.Bounds)
    $ft = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Bold)
    $sf = New-Object System.Drawing.StringFormat; $sf.Alignment = [System.Drawing.StringAlignment]::Near
    $sf.LineAlignment = [System.Drawing.StringAlignment]::Far; $sf.FormatFlags = [System.Drawing.StringFormatFlags]::NoWrap
    $tr = New-Object System.Drawing.RectangleF(($e.Bounds.X + 8), $e.Bounds.Y, ($e.Bounds.Width - 12), ($e.Bounds.Height - 3))
    $tb = New-Object System.Drawing.SolidBrush($fg); $e.Graphics.DrawString($tab.Text, $ft, $tb, $tr, $sf)
    $bb.Dispose(); $tb.Dispose(); $ft.Dispose(); $sf.Dispose()
})
$form.Controls.Add($tabMain)

# ===================== TAB 0: Collections =====================

$tabCollections = New-Object System.Windows.Forms.TabPage; $tabCollections.Text = "Collections"
$tabCollections.BackColor = $clrFormBg; $tabMain.TabPages.Add($tabCollections)

$splitColl = New-Object System.Windows.Forms.SplitContainer; $splitColl.Dock = [System.Windows.Forms.DockStyle]::Fill
$splitColl.Orientation = [System.Windows.Forms.Orientation]::Horizontal; $splitColl.SplitterDistance = 400
$splitColl.SplitterWidth = 6; $splitColl.BackColor = $clrSepLine
$splitColl.Panel1.BackColor = $clrPanelBg; $splitColl.Panel2.BackColor = $clrPanelBg
$splitColl.Panel1MinSize = 100; $splitColl.Panel2MinSize = 80; $tabCollections.Controls.Add($splitColl)

$gridCollections = New-ThemedGrid
$colCName = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colCName.HeaderText = "Name"; $colCName.DataPropertyName = "Name"; $colCName.Width = 250
$colCID   = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colCID.HeaderText = "Collection ID"; $colCID.DataPropertyName = "CollectionID"; $colCID.Width = 100
$colCMem  = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colCMem.HeaderText = "Members"; $colCMem.DataPropertyName = "MemberCount"; $colCMem.Width = 70
$colCLim  = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colCLim.HeaderText = "Limiting Collection"; $colCLim.DataPropertyName = "LimitToCollectionName"; $colCLim.Width = 180
$colCRef  = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colCRef.HeaderText = "Refresh"; $colCRef.DataPropertyName = "RefreshType"; $colCRef.Width = 80
$colCCmt  = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colCCmt.HeaderText = "Comment"; $colCCmt.DataPropertyName = "Comment"; $colCCmt.AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::Fill
$gridCollections.Columns.AddRange([System.Windows.Forms.DataGridViewColumn[]]@($colCName, $colCID, $colCMem, $colCLim, $colCRef, $colCCmt))
$splitColl.Panel1.Controls.Add($gridCollections)

$dtCollections = New-Object System.Data.DataTable
[void]$dtCollections.Columns.Add("Name", [string]); [void]$dtCollections.Columns.Add("CollectionID", [string])
[void]$dtCollections.Columns.Add("MemberCount", [int]); [void]$dtCollections.Columns.Add("LimitToCollectionName", [string])
[void]$dtCollections.Columns.Add("RefreshType", [string]); [void]$dtCollections.Columns.Add("Comment", [string])
[void]$dtCollections.Columns.Add("IsBuiltIn", [bool])
$gridCollections.DataSource = $dtCollections

$txtCollDetail = New-Object System.Windows.Forms.RichTextBox; $txtCollDetail.Dock = [System.Windows.Forms.DockStyle]::Fill
$txtCollDetail.ReadOnly = $true; $txtCollDetail.BackColor = $clrDetailBg; $txtCollDetail.ForeColor = $clrText
$txtCollDetail.Font = New-Object System.Drawing.Font("Consolas", 9.5); $txtCollDetail.BorderStyle = [System.Windows.Forms.BorderStyle]::None
$splitColl.Panel2.Controls.Add($txtCollDetail)

# ===================== TAB 1: WQL Editor =====================

$tabWqlEditor = New-Object System.Windows.Forms.TabPage; $tabWqlEditor.Text = "WQL Editor"
$tabWqlEditor.BackColor = $clrFormBg; $tabMain.TabPages.Add($tabWqlEditor)

$splitWql = New-Object System.Windows.Forms.SplitContainer; $splitWql.Dock = [System.Windows.Forms.DockStyle]::Fill
$splitWql.Orientation = [System.Windows.Forms.Orientation]::Vertical; $splitWql.SplitterWidth = 6
$splitWql.BackColor = $clrSepLine; $splitWql.Panel1.BackColor = $clrPanelBg; $splitWql.Panel2.BackColor = $clrPanelBg
$splitWql.Panel1MinSize = 100; $splitWql.Panel2MinSize = 100
$tabWqlEditor.Controls.Add($splitWql)
# SplitterDistance deferred to Shown event (Width is 0 before layout)

# Left panel: collection selector + rule list + buttons
$pnlWqlLeft = New-Object System.Windows.Forms.Panel; $pnlWqlLeft.Dock = [System.Windows.Forms.DockStyle]::Fill; $pnlWqlLeft.BackColor = $clrPanelBg
$splitWql.Panel1.Controls.Add($pnlWqlLeft)

$lblWqlColl = New-Object System.Windows.Forms.Label; $lblWqlColl.Text = "Collection:"; $lblWqlColl.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$lblWqlColl.AutoSize = $true; $lblWqlColl.Location = New-Object System.Drawing.Point(8, 8); $lblWqlColl.ForeColor = $clrText; $lblWqlColl.BackColor = $clrPanelBg
$pnlWqlLeft.Controls.Add($lblWqlColl)

$cboWqlColl = New-Object System.Windows.Forms.ComboBox; $cboWqlColl.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDown
$cboWqlColl.Font = New-Object System.Drawing.Font("Segoe UI", 9); $cboWqlColl.SetBounds(8, 28, 280, 24)
$cboWqlColl.BackColor = $clrDetailBg; $cboWqlColl.ForeColor = $clrText; $cboWqlColl.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$pnlWqlLeft.Controls.Add($cboWqlColl)

$lblWqlRules = New-Object System.Windows.Forms.Label; $lblWqlRules.Text = "Query Rules:"; $lblWqlRules.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$lblWqlRules.AutoSize = $true; $lblWqlRules.Location = New-Object System.Drawing.Point(8, 60); $lblWqlRules.ForeColor = $clrText; $lblWqlRules.BackColor = $clrPanelBg
$pnlWqlLeft.Controls.Add($lblWqlRules)

$lstWqlRules = New-Object System.Windows.Forms.ListBox; $lstWqlRules.SetBounds(8, 80, 280, 200)
$lstWqlRules.Font = New-Object System.Drawing.Font("Segoe UI", 9); $lstWqlRules.BackColor = $clrDetailBg; $lstWqlRules.ForeColor = $clrText
$lstWqlRules.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$lstWqlRules.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right -bor [System.Windows.Forms.AnchorStyles]::Bottom
$pnlWqlLeft.Controls.Add($lstWqlRules)

$pnlWqlLeftBtns = New-Object System.Windows.Forms.FlowLayoutPanel; $pnlWqlLeftBtns.Dock = [System.Windows.Forms.DockStyle]::Bottom
$pnlWqlLeftBtns.Height = 40; $pnlWqlLeftBtns.BackColor = $clrPanelBg; $pnlWqlLeftBtns.Padding = New-Object System.Windows.Forms.Padding(4, 4, 4, 4)
$btnAddRule = New-Object System.Windows.Forms.Button; $btnAddRule.Text = "Add Rule"; $btnAddRule.Size = New-Object System.Drawing.Size(90, 30); $btnAddRule.Font = New-Object System.Drawing.Font("Segoe UI", 8.5)
Set-ModernButtonStyle -Button $btnAddRule -BackColor $clrAccent
$btnRemoveRule = New-Object System.Windows.Forms.Button; $btnRemoveRule.Text = "Remove"; $btnRemoveRule.Size = New-Object System.Drawing.Size(80, 30); $btnRemoveRule.Font = New-Object System.Drawing.Font("Segoe UI", 8.5)
Set-ModernButtonStyle -Button $btnRemoveRule -BackColor ([System.Drawing.Color]::FromArgb(180, 60, 60))
$btnUpdateRule = New-Object System.Windows.Forms.Button; $btnUpdateRule.Text = "Update Rule"; $btnUpdateRule.Size = New-Object System.Drawing.Size(100, 30); $btnUpdateRule.Font = New-Object System.Drawing.Font("Segoe UI", 8.5)
Set-ModernButtonStyle -Button $btnUpdateRule -BackColor ([System.Drawing.Color]::FromArgb(50, 130, 50))
$pnlWqlLeftBtns.Controls.Add($btnAddRule); $pnlWqlLeftBtns.Controls.Add($btnRemoveRule); $pnlWqlLeftBtns.Controls.Add($btnUpdateRule)
$pnlWqlLeft.Controls.Add($pnlWqlLeftBtns)

# Right panel: rule name + WQL editor + validate/preview
$pnlWqlRight = New-Object System.Windows.Forms.Panel; $pnlWqlRight.Dock = [System.Windows.Forms.DockStyle]::Fill; $pnlWqlRight.BackColor = $clrPanelBg
$splitWql.Panel2.Controls.Add($pnlWqlRight)

$lblRuleName = New-Object System.Windows.Forms.Label; $lblRuleName.Text = "Rule Name:"; $lblRuleName.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$lblRuleName.AutoSize = $true; $lblRuleName.Location = New-Object System.Drawing.Point(8, 8); $lblRuleName.ForeColor = $clrText; $lblRuleName.BackColor = $clrPanelBg
$pnlWqlRight.Controls.Add($lblRuleName)
$txtRuleName = New-Object System.Windows.Forms.TextBox; $txtRuleName.SetBounds(8, 28, 400, 24)
$txtRuleName.Font = New-Object System.Drawing.Font("Segoe UI", 9); $txtRuleName.BackColor = $clrDetailBg; $txtRuleName.ForeColor = $clrText
$txtRuleName.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$pnlWqlRight.Controls.Add($txtRuleName)

$lblWqlQuery = New-Object System.Windows.Forms.Label; $lblWqlQuery.Text = "WQL Query:"; $lblWqlQuery.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$lblWqlQuery.AutoSize = $true; $lblWqlQuery.Location = New-Object System.Drawing.Point(8, 60); $lblWqlQuery.ForeColor = $clrText; $lblWqlQuery.BackColor = $clrPanelBg
$pnlWqlRight.Controls.Add($lblWqlQuery)
$txtWqlEditor = New-Object System.Windows.Forms.TextBox; $txtWqlEditor.SetBounds(8, 80, 600, 200); $txtWqlEditor.Multiline = $true
$txtWqlEditor.ScrollBars = [System.Windows.Forms.ScrollBars]::Both; $txtWqlEditor.WordWrap = $false; $txtWqlEditor.AcceptsReturn = $true
$txtWqlEditor.Font = New-Object System.Drawing.Font("Consolas", 10); $txtWqlEditor.BackColor = $clrDetailBg; $txtWqlEditor.ForeColor = $clrText
$txtWqlEditor.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$txtWqlEditor.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right -bor [System.Windows.Forms.AnchorStyles]::Bottom
$pnlWqlRight.Controls.Add($txtWqlEditor)

$pnlWqlActions = New-Object System.Windows.Forms.FlowLayoutPanel; $pnlWqlActions.Dock = [System.Windows.Forms.DockStyle]::Bottom
$pnlWqlActions.Height = 40; $pnlWqlActions.BackColor = $clrPanelBg; $pnlWqlActions.Padding = New-Object System.Windows.Forms.Padding(4, 4, 4, 4)
$btnValidate = New-Object System.Windows.Forms.Button; $btnValidate.Text = "Validate Query"; $btnValidate.Size = New-Object System.Drawing.Size(130, 30); $btnValidate.Font = New-Object System.Drawing.Font("Segoe UI", 8.5)
Set-ModernButtonStyle -Button $btnValidate -BackColor $clrAccent
$btnPreview = New-Object System.Windows.Forms.Button; $btnPreview.Text = "Preview Results"; $btnPreview.Size = New-Object System.Drawing.Size(130, 30); $btnPreview.Font = New-Object System.Drawing.Font("Segoe UI", 8.5)
Set-ModernButtonStyle -Button $btnPreview -BackColor ([System.Drawing.Color]::FromArgb(100, 100, 100))
$lblValidation = New-Object System.Windows.Forms.Label; $lblValidation.Text = ""; $lblValidation.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$lblValidation.AutoSize = $true; $lblValidation.Margin = New-Object System.Windows.Forms.Padding(8, 8, 0, 0); $lblValidation.ForeColor = $clrHint; $lblValidation.BackColor = $clrPanelBg
$pnlWqlActions.Controls.Add($btnValidate); $pnlWqlActions.Controls.Add($btnPreview); $pnlWqlActions.Controls.Add($lblValidation)
$pnlWqlRight.Controls.Add($pnlWqlActions)

# ===================== TAB 2: Create / Clone =====================

$tabCreate = New-Object System.Windows.Forms.TabPage; $tabCreate.Text = "Create / Clone"
$tabCreate.BackColor = $clrFormBg; $tabMain.TabPages.Add($tabCreate)

$pnlCreate = New-Object System.Windows.Forms.Panel; $pnlCreate.Dock = [System.Windows.Forms.DockStyle]::Fill
$pnlCreate.BackColor = $clrFormBg; $pnlCreate.Padding = New-Object System.Windows.Forms.Padding(20, 16, 20, 16)
$tabCreate.Controls.Add($pnlCreate)

# Create section
$grpCreate = New-Object System.Windows.Forms.GroupBox; $grpCreate.Text = "Create New Collection"
$grpCreate.SetBounds(0, 0, 600, 220); $grpCreate.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$grpCreate.ForeColor = $clrText; $grpCreate.BackColor = $clrFormBg
if ($script:Prefs.DarkMode) { $grpCreate.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat; $grpCreate.ForeColor = $clrSepLine }
$pnlCreate.Controls.Add($grpCreate)

$lblCrName = New-Object System.Windows.Forms.Label; $lblCrName.Text = "Name:"; $lblCrName.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$lblCrName.Location = New-Object System.Drawing.Point(14, 30); $lblCrName.AutoSize = $true; $lblCrName.ForeColor = $clrText
$grpCreate.Controls.Add($lblCrName)
$txtCrName = New-Object System.Windows.Forms.TextBox; $txtCrName.SetBounds(140, 27, 300, 24)
$txtCrName.Font = New-Object System.Drawing.Font("Segoe UI", 9); $txtCrName.BackColor = $clrDetailBg; $txtCrName.ForeColor = $clrText
$grpCreate.Controls.Add($txtCrName)

$lblCrLimit = New-Object System.Windows.Forms.Label; $lblCrLimit.Text = "Limiting Collection:"; $lblCrLimit.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$lblCrLimit.Location = New-Object System.Drawing.Point(14, 62); $lblCrLimit.AutoSize = $true; $lblCrLimit.ForeColor = $clrText
$grpCreate.Controls.Add($lblCrLimit)
$cboCrLimit = New-Object System.Windows.Forms.ComboBox; $cboCrLimit.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDown
$cboCrLimit.SetBounds(140, 59, 300, 24); $cboCrLimit.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$cboCrLimit.BackColor = $clrDetailBg; $cboCrLimit.ForeColor = $clrText; $cboCrLimit.Text = "All Systems"
$grpCreate.Controls.Add($cboCrLimit)

$lblCrComment = New-Object System.Windows.Forms.Label; $lblCrComment.Text = "Comment:"; $lblCrComment.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$lblCrComment.Location = New-Object System.Drawing.Point(14, 94); $lblCrComment.AutoSize = $true; $lblCrComment.ForeColor = $clrText
$grpCreate.Controls.Add($lblCrComment)
$txtCrComment = New-Object System.Windows.Forms.TextBox; $txtCrComment.SetBounds(140, 91, 300, 24)
$txtCrComment.Font = New-Object System.Drawing.Font("Segoe UI", 9); $txtCrComment.BackColor = $clrDetailBg; $txtCrComment.ForeColor = $clrText
$grpCreate.Controls.Add($txtCrComment)

$lblCrRefresh = New-Object System.Windows.Forms.Label; $lblCrRefresh.Text = "Refresh Type:"; $lblCrRefresh.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$lblCrRefresh.Location = New-Object System.Drawing.Point(14, 126); $lblCrRefresh.AutoSize = $true; $lblCrRefresh.ForeColor = $clrText
$grpCreate.Controls.Add($lblCrRefresh)
$cboCrRefresh = New-Object System.Windows.Forms.ComboBox; $cboCrRefresh.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$cboCrRefresh.SetBounds(140, 123, 160, 24); $cboCrRefresh.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$cboCrRefresh.BackColor = $clrDetailBg; $cboCrRefresh.ForeColor = $clrText
[void]$cboCrRefresh.Items.AddRange(@('Both', 'Periodic', 'Continuous', 'Manual')); $cboCrRefresh.SelectedIndex = 0
$grpCreate.Controls.Add($cboCrRefresh)

$btnCreateColl = New-Object System.Windows.Forms.Button; $btnCreateColl.Text = "Create Collection"; $btnCreateColl.SetBounds(140, 165, 150, 34)
$btnCreateColl.Font = New-Object System.Drawing.Font("Segoe UI", 9); Set-ModernButtonStyle -Button $btnCreateColl -BackColor $clrAccent
$grpCreate.Controls.Add($btnCreateColl)

# Clone section
$grpClone = New-Object System.Windows.Forms.GroupBox; $grpClone.Text = "Clone Existing Collection"
$grpClone.SetBounds(0, 236, 600, 130); $grpClone.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$grpClone.ForeColor = $clrText; $grpClone.BackColor = $clrFormBg
if ($script:Prefs.DarkMode) { $grpClone.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat; $grpClone.ForeColor = $clrSepLine }
$pnlCreate.Controls.Add($grpClone)

$lblClSrc = New-Object System.Windows.Forms.Label; $lblClSrc.Text = "Source Collection:"; $lblClSrc.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$lblClSrc.Location = New-Object System.Drawing.Point(14, 30); $lblClSrc.AutoSize = $true; $lblClSrc.ForeColor = $clrText
$grpClone.Controls.Add($lblClSrc)
$cboClSrc = New-Object System.Windows.Forms.ComboBox; $cboClSrc.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDown
$cboClSrc.SetBounds(140, 27, 300, 24); $cboClSrc.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$cboClSrc.BackColor = $clrDetailBg; $cboClSrc.ForeColor = $clrText
$grpClone.Controls.Add($cboClSrc)

$lblClName = New-Object System.Windows.Forms.Label; $lblClName.Text = "New Name:"; $lblClName.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$lblClName.Location = New-Object System.Drawing.Point(14, 62); $lblClName.AutoSize = $true; $lblClName.ForeColor = $clrText
$grpClone.Controls.Add($lblClName)
$txtClName = New-Object System.Windows.Forms.TextBox; $txtClName.SetBounds(140, 59, 300, 24)
$txtClName.Font = New-Object System.Drawing.Font("Segoe UI", 9); $txtClName.BackColor = $clrDetailBg; $txtClName.ForeColor = $clrText
$grpClone.Controls.Add($txtClName)

$btnClone = New-Object System.Windows.Forms.Button; $btnClone.Text = "Clone Collection"; $btnClone.SetBounds(140, 92, 150, 34)
$btnClone.Font = New-Object System.Drawing.Font("Segoe UI", 9); Set-ModernButtonStyle -Button $btnClone -BackColor ([System.Drawing.Color]::FromArgb(130, 80, 180))
$grpClone.Controls.Add($btnClone)

# ===================== TAB 3: Direct Members =====================

$tabMembers = New-Object System.Windows.Forms.TabPage; $tabMembers.Text = "Direct Members"
$tabMembers.BackColor = $clrFormBg; $tabMain.TabPages.Add($tabMembers)

$pnlMembers = New-Object System.Windows.Forms.Panel; $pnlMembers.Dock = [System.Windows.Forms.DockStyle]::Fill
$pnlMembers.BackColor = $clrFormBg; $pnlMembers.Padding = New-Object System.Windows.Forms.Padding(8, 8, 8, 8)
$tabMembers.Controls.Add($pnlMembers)

$lblMemColl = New-Object System.Windows.Forms.Label; $lblMemColl.Text = "Collection:"; $lblMemColl.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$lblMemColl.AutoSize = $true; $lblMemColl.Location = New-Object System.Drawing.Point(8, 8); $lblMemColl.ForeColor = $clrText; $lblMemColl.BackColor = $clrFormBg
$pnlMembers.Controls.Add($lblMemColl)
$cboMemColl = New-Object System.Windows.Forms.ComboBox; $cboMemColl.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDown
$cboMemColl.SetBounds(100, 5, 350, 24); $cboMemColl.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$cboMemColl.BackColor = $clrDetailBg; $cboMemColl.ForeColor = $clrText
$pnlMembers.Controls.Add($cboMemColl)

$btnLoadMembers = New-Object System.Windows.Forms.Button; $btnLoadMembers.Text = "Load Members"; $btnLoadMembers.SetBounds(460, 3, 130, 26)
$btnLoadMembers.Font = New-Object System.Drawing.Font("Segoe UI", 8.5); Set-ModernButtonStyle -Button $btnLoadMembers -BackColor $clrAccent
$pnlMembers.Controls.Add($btnLoadMembers)

$gridMembers = New-ThemedGrid -MultiSelect; $gridMembers.SetBounds(8, 38, 800, 300)
$gridMembers.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right -bor [System.Windows.Forms.AnchorStyles]::Bottom
$colMName = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colMName.HeaderText = "Device Name"; $colMName.DataPropertyName = "Name"; $colMName.Width = 200
$colMRid  = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colMRid.HeaderText = "Resource ID"; $colMRid.DataPropertyName = "ResourceID"; $colMRid.Width = 100
$colMDom  = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colMDom.HeaderText = "Domain"; $colMDom.DataPropertyName = "Domain"; $colMDom.Width = 120
$colMCli  = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colMCli.HeaderText = "Client"; $colMCli.DataPropertyName = "IsClient"; $colMCli.Width = 60
$colMAct  = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colMAct.HeaderText = "Active"; $colMAct.DataPropertyName = "IsActive"; $colMAct.AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::Fill
$gridMembers.Columns.AddRange([System.Windows.Forms.DataGridViewColumn[]]@($colMName, $colMRid, $colMDom, $colMCli, $colMAct))
$pnlMembers.Controls.Add($gridMembers)

$dtMembers = New-Object System.Data.DataTable
[void]$dtMembers.Columns.Add("Name", [string]); [void]$dtMembers.Columns.Add("ResourceID", [int])
[void]$dtMembers.Columns.Add("Domain", [string]); [void]$dtMembers.Columns.Add("IsClient", [bool]); [void]$dtMembers.Columns.Add("IsActive", [bool])
$gridMembers.DataSource = $dtMembers

$lblAddDevices = New-Object System.Windows.Forms.Label; $lblAddDevices.Text = "Add devices (one per line):"
$lblAddDevices.Font = New-Object System.Drawing.Font("Segoe UI", 9); $lblAddDevices.AutoSize = $true
$lblAddDevices.ForeColor = $clrText; $lblAddDevices.BackColor = $clrFormBg; $lblAddDevices.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
$pnlMembers.Controls.Add($lblAddDevices)

$txtAddDevices = New-Object System.Windows.Forms.TextBox; $txtAddDevices.Multiline = $true
$txtAddDevices.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical; $txtAddDevices.Font = New-Object System.Drawing.Font("Consolas", 9)
$txtAddDevices.BackColor = $clrDetailBg; $txtAddDevices.ForeColor = $clrText
$txtAddDevices.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$pnlMembers.Controls.Add($txtAddDevices)

$btnAddMembers = New-Object System.Windows.Forms.Button; $btnAddMembers.Text = "Add Members"
$btnAddMembers.Size = New-Object System.Drawing.Size(130, 30); $btnAddMembers.Font = New-Object System.Drawing.Font("Segoe UI", 8.5)
$btnAddMembers.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
Set-ModernButtonStyle -Button $btnAddMembers -BackColor $clrAccent; $pnlMembers.Controls.Add($btnAddMembers)

$btnRemoveMembers = New-Object System.Windows.Forms.Button; $btnRemoveMembers.Text = "Remove Selected"
$btnRemoveMembers.Size = New-Object System.Drawing.Size(140, 30); $btnRemoveMembers.Font = New-Object System.Drawing.Font("Segoe UI", 8.5)
$btnRemoveMembers.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
Set-ModernButtonStyle -Button $btnRemoveMembers -BackColor ([System.Drawing.Color]::FromArgb(180, 60, 60)); $pnlMembers.Controls.Add($btnRemoveMembers)

# Position member add controls on Resize
$pnlMembers.Add_Resize({
    $bottom = $pnlMembers.ClientSize.Height
    $lblAddDevices.Location = New-Object System.Drawing.Point(8, ($bottom - 135))
    $txtAddDevices.SetBounds(8, ($bottom - 115), ($pnlMembers.ClientSize.Width - 180), 70)
    $btnAddMembers.Location = New-Object System.Drawing.Point(($pnlMembers.ClientSize.Width - 160), ($bottom - 115))
    $btnRemoveMembers.Location = New-Object System.Drawing.Point(($pnlMembers.ClientSize.Width - 160), ($bottom - 75))
})

# ===================== TAB 4: Templates =====================

$tabTemplates = New-Object System.Windows.Forms.TabPage; $tabTemplates.Text = "Templates"
$tabTemplates.BackColor = $clrFormBg; $tabMain.TabPages.Add($tabTemplates)

$splitTemplates = New-Object System.Windows.Forms.SplitContainer; $splitTemplates.Dock = [System.Windows.Forms.DockStyle]::Fill
$splitTemplates.Orientation = [System.Windows.Forms.Orientation]::Vertical; $splitTemplates.SplitterWidth = 6
$splitTemplates.BackColor = $clrSepLine; $splitTemplates.Panel1.BackColor = $clrPanelBg; $splitTemplates.Panel2.BackColor = $clrPanelBg
$splitTemplates.Panel1MinSize = 100; $splitTemplates.Panel2MinSize = 100
$tabTemplates.Controls.Add($splitTemplates)
# SplitterDistance deferred to Shown event (Width is 0 before layout)

# Left: template tree
$treeTemplates = New-Object System.Windows.Forms.TreeView; $treeTemplates.Dock = [System.Windows.Forms.DockStyle]::Fill
$treeTemplates.Font = New-Object System.Drawing.Font("Segoe UI", 9.5); $treeTemplates.BackColor = $clrTreeBg; $treeTemplates.ForeColor = $clrText
$treeTemplates.BorderStyle = [System.Windows.Forms.BorderStyle]::None; $treeTemplates.FullRowSelect = $true
$treeTemplates.HideSelection = $false; $treeTemplates.ShowLines = $true; $treeTemplates.ShowPlusMinus = $true
if ($script:Prefs.DarkMode) { $treeTemplates.LineColor = [System.Drawing.Color]::FromArgb(80, 80, 80) }
Enable-DoubleBuffer -Control $treeTemplates; $splitTemplates.Panel1.Controls.Add($treeTemplates)

# Right: template detail + parameter inputs + WQL preview + action buttons
$pnlTplRight = New-Object System.Windows.Forms.Panel; $pnlTplRight.Dock = [System.Windows.Forms.DockStyle]::Fill; $pnlTplRight.BackColor = $clrPanelBg
$pnlTplRight.Padding = New-Object System.Windows.Forms.Padding(12, 8, 12, 8); $splitTemplates.Panel2.Controls.Add($pnlTplRight)

$lblTplName = New-Object System.Windows.Forms.Label; $lblTplName.Text = "Select a template"; $lblTplName.Font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
$lblTplName.AutoSize = $true; $lblTplName.Location = New-Object System.Drawing.Point(12, 8); $lblTplName.ForeColor = $clrText; $lblTplName.BackColor = $clrPanelBg
$pnlTplRight.Controls.Add($lblTplName)
$lblTplDesc = New-Object System.Windows.Forms.Label; $lblTplDesc.Text = ""; $lblTplDesc.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$lblTplDesc.AutoSize = $true; $lblTplDesc.Location = New-Object System.Drawing.Point(12, 32); $lblTplDesc.ForeColor = $clrHint; $lblTplDesc.BackColor = $clrPanelBg
$pnlTplRight.Controls.Add($lblTplDesc)

# Dynamic parameter panel (populated when a parameterized template is selected)
$pnlTplParams = New-Object System.Windows.Forms.FlowLayoutPanel; $pnlTplParams.SetBounds(12, 60, 500, 120)
$pnlTplParams.FlowDirection = [System.Windows.Forms.FlowDirection]::TopDown; $pnlTplParams.WrapContents = $false
$pnlTplParams.BackColor = $clrPanelBg; $pnlTplParams.AutoScroll = $true
$pnlTplParams.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$pnlTplRight.Controls.Add($pnlTplParams)

$lblTplWql = New-Object System.Windows.Forms.Label; $lblTplWql.Text = "Generated WQL:"; $lblTplWql.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$lblTplWql.AutoSize = $true; $lblTplWql.Location = New-Object System.Drawing.Point(12, 188); $lblTplWql.ForeColor = $clrText; $lblTplWql.BackColor = $clrPanelBg
$pnlTplRight.Controls.Add($lblTplWql)
$txtTplWql = New-Object System.Windows.Forms.TextBox; $txtTplWql.SetBounds(12, 208, 580, 150); $txtTplWql.Multiline = $true; $txtTplWql.ReadOnly = $true
$txtTplWql.ScrollBars = [System.Windows.Forms.ScrollBars]::Both; $txtTplWql.WordWrap = $false
$txtTplWql.Font = New-Object System.Drawing.Font("Consolas", 9.5); $txtTplWql.BackColor = $clrDetailBg; $txtTplWql.ForeColor = $clrText
$txtTplWql.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right -bor [System.Windows.Forms.AnchorStyles]::Bottom
$pnlTplRight.Controls.Add($txtTplWql)

$pnlTplActions = New-Object System.Windows.Forms.FlowLayoutPanel; $pnlTplActions.Dock = [System.Windows.Forms.DockStyle]::Bottom
$pnlTplActions.Height = 44; $pnlTplActions.BackColor = $clrPanelBg; $pnlTplActions.Padding = New-Object System.Windows.Forms.Padding(8, 6, 8, 6)
$btnTplCopyToEditor = New-Object System.Windows.Forms.Button; $btnTplCopyToEditor.Text = "Copy to WQL Editor"; $btnTplCopyToEditor.Size = New-Object System.Drawing.Size(160, 30)
$btnTplCopyToEditor.Font = New-Object System.Drawing.Font("Segoe UI", 8.5); Set-ModernButtonStyle -Button $btnTplCopyToEditor -BackColor $clrAccent
$btnTplCreateColl = New-Object System.Windows.Forms.Button; $btnTplCreateColl.Text = "Create Collection from Template"; $btnTplCreateColl.Size = New-Object System.Drawing.Size(240, 30)
$btnTplCreateColl.Font = New-Object System.Drawing.Font("Segoe UI", 8.5); Set-ModernButtonStyle -Button $btnTplCreateColl -BackColor ([System.Drawing.Color]::FromArgb(50, 130, 50))
$pnlTplActions.Controls.Add($btnTplCopyToEditor); $pnlTplActions.Controls.Add($btnTplCreateColl)
$pnlTplRight.Controls.Add($pnlTplActions)

# ---------------------------------------------------------------------------
# Finalize dock Z-order
# ---------------------------------------------------------------------------

$form.Controls.Add($menuStrip); $menuStrip.SendToBack()
$pnlSep3.BringToFront(); $pnlFilter.BringToFront(); $pnlSep2.BringToFront()
$pnlCards.BringToFront(); $pnlSep1.BringToFront(); $pnlConnBar.BringToFront(); $pnlHeader.BringToFront()
$tabMain.BringToFront()

# ---------------------------------------------------------------------------
# Module-scoped data
# ---------------------------------------------------------------------------

$script:CollectionData    = @()
$script:OperationalTpls   = @()
$script:ParameterizedTpls = @()
$script:SelectedTemplate  = $null

# ---------------------------------------------------------------------------
# Status bar helper
# ---------------------------------------------------------------------------

function Update-StatusBar {
    $parts = @()
    if (Test-CMConnection) { $parts += "Connected to $($script:Prefs.SiteCode)" } else { $parts += "Disconnected" }
    $statusLabel.Text = $parts -join " | "
    $tabIdx = $tabMain.SelectedIndex
    $count = switch ($tabIdx) {
        0 { $dtCollections.DefaultView.Count }
        3 { $dtMembers.DefaultView.Count }
        default { 0 }
    }
    if ($count -gt 0) { $statusRowCount.Text = "$count rows" } else { $statusRowCount.Text = "" }
}

# ---------------------------------------------------------------------------
# Filter logic
# ---------------------------------------------------------------------------

function Invoke-ApplyFilter {
    $filterText = $txtFilter.Text.Trim()
    if ($tabMain.SelectedIndex -eq 0 -and $filterText) {
        $escaped = $filterText.Replace("'", "''")
        $dtCollections.DefaultView.RowFilter = "Name LIKE '%$escaped%' OR CollectionID LIKE '%$escaped%' OR Comment LIKE '%$escaped%'"
    } elseif ($tabMain.SelectedIndex -eq 0) {
        $dtCollections.DefaultView.RowFilter = ''
    }
    Update-StatusBar
}

$txtFilter.Add_TextChanged({ Invoke-ApplyFilter })
$tabMain.Add_SelectedIndexChanged({ Invoke-ApplyFilter; Update-StatusBar })

# ---------------------------------------------------------------------------
# Load Collections workflow
# ---------------------------------------------------------------------------

function Invoke-LoadCollections {
    if (-not $script:Prefs.SiteCode -or -not $script:Prefs.SMSProvider) {
        [System.Windows.Forms.MessageBox]::Show("Site Code and SMS Provider must be configured in File > Preferences.", "Configuration Required", "OK", "Warning") | Out-Null
        return
    }

    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor; $btnLoad.Enabled = $false

    try {
        if (-not (Test-CMConnection)) {
            Add-LogLine -TextBox $txtLog -Message "Connecting to $($script:Prefs.SiteCode)..."
            [System.Windows.Forms.Application]::DoEvents()
            $connected = Connect-CMSite -SiteCode $script:Prefs.SiteCode -SMSProvider $script:Prefs.SMSProvider
            if (-not $connected) {
                Add-LogLine -TextBox $txtLog -Message "ERROR: Failed to connect"; return
            }
            Add-LogLine -TextBox $txtLog -Message "Connected to $($script:Prefs.SiteCode)"
        }

        Add-LogLine -TextBox $txtLog -Message "Loading device collections..."
        [System.Windows.Forms.Application]::DoEvents()
        $script:CollectionData = @(Get-AllDeviceCollections)
        Add-LogLine -TextBox $txtLog -Message "Loaded $($script:CollectionData.Count) collections"
        [System.Windows.Forms.Application]::DoEvents()

        # Populate collections DataTable
        $dtCollections.Clear()
        $dtCollections.BeginLoadData()
        foreach ($c in $script:CollectionData) {
            [void]$dtCollections.Rows.Add($c.Name, $c.CollectionID, $c.MemberCount, $c.LimitToCollectionName, $c.RefreshType, $c.Comment, $c.IsBuiltIn)
        }
        $dtCollections.EndLoadData()

        # Populate collection dropdowns (WQL editor, Create, Clone, Members)
        $collNames = @($script:CollectionData | Sort-Object Name | ForEach-Object { $_.Name })
        $cboWqlColl.Items.Clear(); $cboWqlColl.Items.AddRange($collNames)
        $cboCrLimit.Items.Clear(); $cboCrLimit.Items.AddRange(@('All Systems') + $collNames)
        $cboClSrc.Items.Clear(); $cboClSrc.Items.AddRange($collNames)
        $cboMemColl.Items.Clear(); $cboMemColl.Items.AddRange($collNames)

        # Update summary cards
        $totalQueryRules = ($script:CollectionData | Where-Object { $_.CollectionRules } | ForEach-Object { @($_.CollectionRules | Where-Object { $_.SmsProviderObjectPath -match 'QueryRule' }).Count } | Measure-Object -Sum).Sum
        Update-Card -Card $cardCollections -ValueText "$($script:CollectionData.Count)" -Severity 'info'
        Update-Card -Card $cardQueryRules -ValueText "$([int]$totalQueryRules)" -Severity 'info'

        Add-LogLine -TextBox $txtLog -Message "Collections loaded. Use tabs to manage."
    }
    catch {
        Add-LogLine -TextBox $txtLog -Message "ERROR: $($_.Exception.Message)"
        Write-Log "Load failed: $_" -Level ERROR
    }
    finally {
        $form.Cursor = [System.Windows.Forms.Cursors]::Default; $btnLoad.Enabled = $true; Update-StatusBar
    }
}

$btnLoad.Add_Click({ Invoke-LoadCollections })

# ---------------------------------------------------------------------------
# Collection detail on selection
# ---------------------------------------------------------------------------

$gridCollections.Add_SelectionChanged({
    if ($gridCollections.SelectedRows.Count -eq 0) { $txtCollDetail.Text = ''; return }
    $rowIdx = $gridCollections.SelectedRows[0].Index
    if ($rowIdx -lt 0 -or $rowIdx -ge $dtCollections.DefaultView.Count) { return }
    $row = $dtCollections.DefaultView[$rowIdx]
    $lines = @(
        "COLLECTION DETAILS", ("-" * 40), "",
        "Name:          $($row['Name'])",
        "Collection ID: $($row['CollectionID'])",
        "Members:       $($row['MemberCount'])",
        "Limiting:      $($row['LimitToCollectionName'])",
        "Refresh:       $($row['RefreshType'])",
        "Comment:       $($row['Comment'])",
        "Built-in:      $($row['IsBuiltIn'])"
    )
    $txtCollDetail.Text = $lines -join "`r`n"
})

# ---------------------------------------------------------------------------
# WQL Editor: load rules when collection selected
# ---------------------------------------------------------------------------

$cboWqlColl.Add_SelectedIndexChanged({
    $collName = $cboWqlColl.SelectedItem
    if (-not $collName) { return }
    $lstWqlRules.Items.Clear(); $txtRuleName.Text = ''; $txtWqlEditor.Text = ''; $lblValidation.Text = ''

    $coll = $script:CollectionData | Where-Object { $_.Name -eq $collName } | Select-Object -First 1
    if (-not $coll) { return }

    try {
        Add-LogLine -TextBox $txtLog -Message "Loading query rules for '$collName'..."
        [System.Windows.Forms.Application]::DoEvents()
        $rules = @(Get-CollectionQueryRules -CollectionId $coll.CollectionID)
        foreach ($r in $rules) { [void]$lstWqlRules.Items.Add($r.RuleName) }
        Add-LogLine -TextBox $txtLog -Message "Loaded $($rules.Count) query rules"
        $script:CurrentQueryRules = $rules
    }
    catch {
        Add-LogLine -TextBox $txtLog -Message "ERROR loading rules: $($_.Exception.Message)"
    }
})

$lstWqlRules.Add_SelectedIndexChanged({
    $ruleName = $lstWqlRules.SelectedItem
    if (-not $ruleName -or -not $script:CurrentQueryRules) { return }
    $rule = $script:CurrentQueryRules | Where-Object { $_.RuleName -eq $ruleName } | Select-Object -First 1
    if ($rule) {
        $txtRuleName.Text = $rule.RuleName
        $txtWqlEditor.Text = $rule.QueryExpression
        $lblValidation.Text = ''
    }
})

# ---------------------------------------------------------------------------
# WQL Editor: Validate, Preview, Add/Remove/Update
# ---------------------------------------------------------------------------

$btnValidate.Add_Click({
    $query = $txtWqlEditor.Text.Trim()
    if (-not $query) { $lblValidation.Text = "No query to validate"; $lblValidation.ForeColor = $clrWarnText; return }
    Add-LogLine -TextBox $txtLog -Message "Validating WQL query..."
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    [System.Windows.Forms.Application]::DoEvents()
    try {
        $result = Test-WqlQuery -QueryExpression $query
        if ($result.IsValid) {
            $lblValidation.Text = "VALID ($($result.ResultCount) results)"; $lblValidation.ForeColor = $clrOkText
            Add-LogLine -TextBox $txtLog -Message "Query valid: $($result.ResultCount) results"
        } else {
            $lblValidation.Text = "INVALID: $($result.ErrorMessage)"; $lblValidation.ForeColor = $clrErrText
            Add-LogLine -TextBox $txtLog -Message "Query invalid: $($result.ErrorMessage)"
        }
    }
    catch { $lblValidation.Text = "Error: $_"; $lblValidation.ForeColor = $clrErrText }
    finally { $form.Cursor = [System.Windows.Forms.Cursors]::Default }
})

$btnPreview.Add_Click({
    $query = $txtWqlEditor.Text.Trim()
    if (-not $query) { return }
    Add-LogLine -TextBox $txtLog -Message "Running WQL preview..."
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    [System.Windows.Forms.Application]::DoEvents()
    try {
        $preview = Invoke-WqlPreview -QueryExpression $query -MaxResults 50
        $msg = "Preview: $($preview.TotalCount) total results"
        if ($preview.Results.Count -gt 0) {
            $msg += "`r`n`r`n" + ($preview.Results | Format-Table -AutoSize | Out-String)
        }
        [System.Windows.Forms.MessageBox]::Show($msg, "WQL Preview ($($preview.TotalCount) results)", "OK", "Information") | Out-Null
        Add-LogLine -TextBox $txtLog -Message "Preview complete: $($preview.TotalCount) results"
    }
    catch { Add-LogLine -TextBox $txtLog -Message "Preview error: $_" }
    finally { $form.Cursor = [System.Windows.Forms.Cursors]::Default }
})

$btnAddRule.Add_Click({
    $collName = $cboWqlColl.SelectedItem; $ruleName = $txtRuleName.Text.Trim(); $query = $txtWqlEditor.Text.Trim()
    if (-not $collName -or -not $ruleName -or -not $query) {
        [System.Windows.Forms.MessageBox]::Show("Select a collection and enter a rule name and query.", "Missing Input", "OK", "Warning") | Out-Null; return
    }
    $coll = $script:CollectionData | Where-Object { $_.Name -eq $collName } | Select-Object -First 1
    if (-not $coll) { return }
    try {
        Add-QueryRule -CollectionId $coll.CollectionID -RuleName $ruleName -QueryExpression $query
        Add-LogLine -TextBox $txtLog -Message "Added query rule '$ruleName' to '$collName'"
        $cboWqlColl.SelectedIndex = $cboWqlColl.SelectedIndex  # Refresh rules list
    }
    catch { Add-LogLine -TextBox $txtLog -Message "ERROR adding rule: $($_.Exception.Message)" }
})

$btnRemoveRule.Add_Click({
    $collName = $cboWqlColl.SelectedItem; $ruleName = $lstWqlRules.SelectedItem
    if (-not $collName -or -not $ruleName) { return }
    $confirm = [System.Windows.Forms.MessageBox]::Show("Remove query rule '$ruleName' from '$collName'?", "Confirm Remove", "YesNo", "Question")
    if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { return }
    $coll = $script:CollectionData | Where-Object { $_.Name -eq $collName } | Select-Object -First 1
    if (-not $coll) { return }
    try {
        Remove-QueryRule -CollectionId $coll.CollectionID -RuleName $ruleName
        Add-LogLine -TextBox $txtLog -Message "Removed query rule '$ruleName'"
        $cboWqlColl.SelectedIndex = $cboWqlColl.SelectedIndex
    }
    catch { Add-LogLine -TextBox $txtLog -Message "ERROR removing rule: $($_.Exception.Message)" }
})

$btnUpdateRule.Add_Click({
    $collName = $cboWqlColl.SelectedItem; $oldRuleName = $lstWqlRules.SelectedItem
    $newRuleName = $txtRuleName.Text.Trim(); $query = $txtWqlEditor.Text.Trim()
    if (-not $collName -or -not $oldRuleName -or -not $newRuleName -or -not $query) {
        [System.Windows.Forms.MessageBox]::Show("Select a collection, rule, and enter updated query.", "Missing Input", "OK", "Warning") | Out-Null; return
    }
    $coll = $script:CollectionData | Where-Object { $_.Name -eq $collName } | Select-Object -First 1
    if (-not $coll) { return }
    try {
        Update-QueryRule -CollectionId $coll.CollectionID -OldRuleName $oldRuleName -NewRuleName $newRuleName -NewQueryExpression $query
        Add-LogLine -TextBox $txtLog -Message "Updated query rule '$oldRuleName' on '$collName'"
        $cboWqlColl.SelectedIndex = $cboWqlColl.SelectedIndex
    }
    catch { Add-LogLine -TextBox $txtLog -Message "ERROR updating rule: $($_.Exception.Message)" }
})

# ---------------------------------------------------------------------------
# Create / Clone handlers
# ---------------------------------------------------------------------------

$btnCreateColl.Add_Click({
    $name = $txtCrName.Text.Trim(); $limit = $cboCrLimit.Text.Trim(); $comment = $txtCrComment.Text.Trim()
    $refresh = $cboCrRefresh.SelectedItem
    if (-not $name -or -not $limit) {
        [System.Windows.Forms.MessageBox]::Show("Name and Limiting Collection are required.", "Missing Input", "OK", "Warning") | Out-Null; return
    }
    try {
        $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        $result = New-ManagedCollection -Name $name -LimitingCollectionName $limit -Comment $comment -RefreshType $refresh
        Add-LogLine -TextBox $txtLog -Message "Created collection '$name' (ID: $($result.CollectionID))"
        $txtCrName.Text = ''; $txtCrComment.Text = ''
    }
    catch { Add-LogLine -TextBox $txtLog -Message "ERROR creating collection: $($_.Exception.Message)" }
    finally { $form.Cursor = [System.Windows.Forms.Cursors]::Default }
})

$btnClone.Add_Click({
    $srcName = $cboClSrc.Text.Trim(); $newName = $txtClName.Text.Trim()
    if (-not $srcName -or -not $newName) {
        [System.Windows.Forms.MessageBox]::Show("Source collection and new name are required.", "Missing Input", "OK", "Warning") | Out-Null; return
    }
    $coll = $script:CollectionData | Where-Object { $_.Name -eq $srcName } | Select-Object -First 1
    if (-not $coll) { Add-LogLine -TextBox $txtLog -Message "Source collection not found"; return }
    try {
        $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        $result = Copy-ManagedCollection -SourceCollectionId $coll.CollectionID -NewName $newName
        Add-LogLine -TextBox $txtLog -Message "Cloned '$srcName' -> '$newName' (ID: $($result.CollectionID))"
        $txtClName.Text = ''
    }
    catch { Add-LogLine -TextBox $txtLog -Message "ERROR cloning: $($_.Exception.Message)" }
    finally { $form.Cursor = [System.Windows.Forms.Cursors]::Default }
})

# ---------------------------------------------------------------------------
# Direct Members handlers
# ---------------------------------------------------------------------------

$btnLoadMembers.Add_Click({
    $collName = $cboMemColl.Text.Trim()
    if (-not $collName) { return }
    $coll = $script:CollectionData | Where-Object { $_.Name -eq $collName } | Select-Object -First 1
    if (-not $coll) { Add-LogLine -TextBox $txtLog -Message "Collection not found"; return }
    try {
        $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        Add-LogLine -TextBox $txtLog -Message "Loading members of '$collName'..."
        [System.Windows.Forms.Application]::DoEvents()
        $members = @(Get-CollectionMembers -CollectionId $coll.CollectionID)
        $dtMembers.Clear(); $dtMembers.BeginLoadData()
        foreach ($m in $members) { [void]$dtMembers.Rows.Add($m.Name, $m.ResourceID, $m.Domain, $m.IsClient, $m.IsActive) }
        $dtMembers.EndLoadData()
        Add-LogLine -TextBox $txtLog -Message "Loaded $($members.Count) members"
    }
    catch { Add-LogLine -TextBox $txtLog -Message "ERROR: $($_.Exception.Message)" }
    finally { $form.Cursor = [System.Windows.Forms.Cursors]::Default; Update-StatusBar }
})

$btnAddMembers.Add_Click({
    $collName = $cboMemColl.Text.Trim()
    if (-not $collName) { return }
    $coll = $script:CollectionData | Where-Object { $_.Name -eq $collName } | Select-Object -First 1
    if (-not $coll) { return }
    $devices = $txtAddDevices.Text -split "`r`n|`n|," | ForEach-Object { $_.Trim() } | Where-Object { $_ }
    if ($devices.Count -eq 0) { return }
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    foreach ($dev in $devices) {
        try {
            $result = Add-DirectMember -CollectionId $coll.CollectionID -DeviceName $dev
            if ($result) { Add-LogLine -TextBox $txtLog -Message "Added '$dev' to '$collName'" }
        }
        catch { Add-LogLine -TextBox $txtLog -Message "ERROR adding '$dev': $($_.Exception.Message)" }
        [System.Windows.Forms.Application]::DoEvents()
    }
    $txtAddDevices.Text = ''
    $form.Cursor = [System.Windows.Forms.Cursors]::Default
})

$btnRemoveMembers.Add_Click({
    $collName = $cboMemColl.Text.Trim()
    if (-not $collName -or $gridMembers.SelectedRows.Count -eq 0) { return }
    $coll = $script:CollectionData | Where-Object { $_.Name -eq $collName } | Select-Object -First 1
    if (-not $coll) { return }
    $confirm = [System.Windows.Forms.MessageBox]::Show("Remove $($gridMembers.SelectedRows.Count) selected member(s)?", "Confirm Remove", "YesNo", "Question")
    if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { return }
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    foreach ($row in $gridMembers.SelectedRows) {
        $rid = [int]$dtMembers.DefaultView[$row.Index]["ResourceID"]
        try { Remove-DirectMember -CollectionId $coll.CollectionID -ResourceId $rid; Add-LogLine -TextBox $txtLog -Message "Removed ResourceId $rid" }
        catch { Add-LogLine -TextBox $txtLog -Message "ERROR removing $rid : $($_.Exception.Message)" }
    }
    $form.Cursor = [System.Windows.Forms.Cursors]::Default
    $btnLoadMembers.PerformClick()  # Refresh
})

# ---------------------------------------------------------------------------
# Templates: Load and populate tree
# ---------------------------------------------------------------------------

function Invoke-LoadTemplates {
    $script:OperationalTpls = @(Get-OperationalTemplates)
    $script:ParameterizedTpls = @(Get-ParameterizedTemplates)

    $treeTemplates.BeginUpdate(); $treeTemplates.Nodes.Clear()

    # Ready-Made
    $nodeOp = New-Object System.Windows.Forms.TreeNode("Ready-Made ($($script:OperationalTpls.Count))"); $nodeOp.ForeColor = $clrText
    $categories = $script:OperationalTpls | ForEach-Object { $_.Category } | Sort-Object -Unique
    foreach ($cat in $categories) {
        $catItems = @($script:OperationalTpls | Where-Object { $_.Category -eq $cat })
        $catNode = New-Object System.Windows.Forms.TreeNode("$cat ($($catItems.Count))"); $catNode.ForeColor = $clrText
        foreach ($tpl in $catItems) {
            $leaf = New-Object System.Windows.Forms.TreeNode($tpl.Name); $leaf.Tag = @{ Type = 'operational'; Template = $tpl }; $leaf.ForeColor = $clrText
            $catNode.Nodes.Add($leaf) | Out-Null
        }
        $nodeOp.Nodes.Add($catNode) | Out-Null
    }
    $treeTemplates.Nodes.Add($nodeOp) | Out-Null

    # Parameterized
    $nodeParam = New-Object System.Windows.Forms.TreeNode("Parameterized ($($script:ParameterizedTpls.Count))"); $nodeParam.ForeColor = $clrText
    $pCats = $script:ParameterizedTpls | ForEach-Object { $_.Category } | Sort-Object -Unique
    foreach ($cat in $pCats) {
        $catItems = @($script:ParameterizedTpls | Where-Object { $_.Category -eq $cat })
        $catNode = New-Object System.Windows.Forms.TreeNode("$cat ($($catItems.Count))"); $catNode.ForeColor = $clrText
        foreach ($tpl in $catItems) {
            $leaf = New-Object System.Windows.Forms.TreeNode($tpl.Name); $leaf.Tag = @{ Type = 'parameterized'; Template = $tpl }; $leaf.ForeColor = $clrText
            $catNode.Nodes.Add($leaf) | Out-Null
        }
        $nodeParam.Nodes.Add($catNode) | Out-Null
    }
    $treeTemplates.Nodes.Add($nodeParam) | Out-Null

    $nodeOp.Expand(); $nodeParam.Expand()
    $treeTemplates.EndUpdate()
    Add-LogLine -TextBox $txtLog -Message "Loaded $($script:OperationalTpls.Count) operational + $($script:ParameterizedTpls.Count) parameterized templates"
}

# Template selection handler
$treeTemplates.Add_AfterSelect({
    param($s, $e)
    $node = $e.Node; if (-not $node -or -not $node.Tag) { return }
    $info = $node.Tag; $tpl = $info.Template; $script:SelectedTemplate = $info

    $lblTplName.Text = $tpl.Name
    $lblTplDesc.Text = if ($tpl.Description) { $tpl.Description } else { "Category: $($tpl.Category) | Limiting: $($tpl.LimitingCollection)" }

    # Clear parameter panel
    $pnlTplParams.Controls.Clear()

    if ($info.Type -eq 'parameterized' -and $tpl.Parameters -and $tpl.Parameters.Count -gt 0) {
        foreach ($param in $tpl.Parameters) {
            $lbl = New-Object System.Windows.Forms.Label; $lbl.Text = "$($param.Label):"; $lbl.Font = New-Object System.Drawing.Font("Segoe UI", 9)
            $lbl.AutoSize = $true; $lbl.ForeColor = $clrText; $lbl.BackColor = $clrPanelBg; $lbl.Margin = New-Object System.Windows.Forms.Padding(0, 4, 0, 0)
            $pnlTplParams.Controls.Add($lbl)
            $txt = New-Object System.Windows.Forms.TextBox; $txt.Width = 350; $txt.Font = New-Object System.Drawing.Font("Segoe UI", 9)
            $txt.BackColor = $clrDetailBg; $txt.ForeColor = $clrText; $txt.Tag = $param.Placeholder
            if ($param.DefaultValue) { $txt.Text = $param.DefaultValue }
            if ($param.HelpText) {
                $tip = New-Object System.Windows.Forms.ToolTip; $tip.SetToolTip($txt, $param.HelpText)
            }
            $txt.Add_TextChanged({ Update-TemplatePreview }.GetNewClosure())
            $pnlTplParams.Controls.Add($txt)
        }
        Update-TemplatePreview
    } else {
        $txtTplWql.Text = $tpl.Query
    }
})

function Update-TemplatePreview {
    if (-not $script:SelectedTemplate) { return }
    $tpl = $script:SelectedTemplate.Template
    $values = @{}
    foreach ($ctrl in $pnlTplParams.Controls) {
        if ($ctrl -is [System.Windows.Forms.TextBox] -and $ctrl.Tag) {
            $values[$ctrl.Tag] = $ctrl.Text
        }
    }
    if ($values.Count -gt 0) {
        $txtTplWql.Text = Expand-TemplateParameters -QueryTemplate $tpl.Query -ParameterValues $values
    } else {
        $txtTplWql.Text = $tpl.Query
    }
}

# Template action buttons
$btnTplCopyToEditor.Add_Click({
    $query = $txtTplWql.Text.Trim()
    if (-not $query) { return }
    $txtWqlEditor.Text = $query
    if ($script:SelectedTemplate) { $txtRuleName.Text = $script:SelectedTemplate.Template.Name }
    $tabMain.SelectedIndex = 1  # Switch to WQL Editor tab
    Add-LogLine -TextBox $txtLog -Message "Template query copied to WQL Editor"
})

$btnTplCreateColl.Add_Click({
    if (-not $script:SelectedTemplate) { return }
    $tpl = $script:SelectedTemplate.Template; $query = $txtTplWql.Text.Trim()
    if (-not $query) { return }
    if (-not (Test-CMConnection)) {
        [System.Windows.Forms.MessageBox]::Show("Connect to MECM first (Load Collections).", "Not Connected", "OK", "Warning") | Out-Null; return
    }
    try {
        $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        $limitColl = if ($tpl.LimitingCollection) { $tpl.LimitingCollection } else { 'All Systems' }
        $refreshType = if ($tpl.RefreshType) { $tpl.RefreshType } else { 'Both' }
        $coll = New-ManagedCollection -Name $tpl.Name -LimitingCollectionName $limitColl -RefreshType $refreshType
        Add-QueryRule -CollectionId $coll.CollectionID -RuleName $tpl.Name -QueryExpression $query
        Add-LogLine -TextBox $txtLog -Message "Created collection '$($tpl.Name)' with query rule from template"
    }
    catch { Add-LogLine -TextBox $txtLog -Message "ERROR: $($_.Exception.Message)" }
    finally { $form.Cursor = [System.Windows.Forms.Cursors]::Default }
})

# ---------------------------------------------------------------------------
# Export handlers
# ---------------------------------------------------------------------------

$btnExportCsv.Add_Click({
    if ($dtCollections.Rows.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("No data to export.", "Export", "OK", "Information") | Out-Null; return
    }
    $sfd = New-Object System.Windows.Forms.SaveFileDialog; $sfd.Filter = "CSV files (*.csv)|*.csv"
    $sfd.FileName = "Collections-$(Get-Date -Format 'yyyyMMdd-HHmmss').csv"
    $sfd.InitialDirectory = Join-Path $PSScriptRoot "Reports"
    if ($sfd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        Export-CollectionCsv -DataTable $dtCollections -OutputPath $sfd.FileName
        Add-LogLine -TextBox $txtLog -Message "Exported CSV: $($sfd.FileName)"
    }
    $sfd.Dispose()
})

$btnExportHtml.Add_Click({
    if ($dtCollections.Rows.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("No data to export.", "Export", "OK", "Information") | Out-Null; return
    }
    $sfd = New-Object System.Windows.Forms.SaveFileDialog; $sfd.Filter = "HTML files (*.html)|*.html"
    $sfd.FileName = "Collections-$(Get-Date -Format 'yyyyMMdd-HHmmss').html"
    $sfd.InitialDirectory = Join-Path $PSScriptRoot "Reports"
    if ($sfd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        Export-CollectionHtml -DataTable $dtCollections -OutputPath $sfd.FileName
        Add-LogLine -TextBox $txtLog -Message "Exported HTML: $($sfd.FileName)"
    }
    $sfd.Dispose()
})

# ---------------------------------------------------------------------------
# Form events
# ---------------------------------------------------------------------------

$form.Add_FormClosing({
    Save-WindowState
    if (Test-CMConnection) { Disconnect-CMSite }
})

$form.Add_Shown({
    Restore-WindowState; Update-StatusBar
    # Set SplitContainer distances now that layout is complete and Width is real
    $splitWql.SplitterDistance = [Math]::Max($splitWql.Panel1MinSize, [int]($splitWql.Width * 0.3))
    $splitTemplates.SplitterDistance = [Math]::Max($splitTemplates.Panel1MinSize, [int]($splitTemplates.Width * 0.35))
    Invoke-LoadTemplates
    Add-LogLine -TextBox $txtLog -Message "Collection Manager ready. Configure Site/Provider in Preferences, then click Load Collections."
})

# ---------------------------------------------------------------------------
# Run
# ---------------------------------------------------------------------------

[System.Windows.Forms.Application]::Run($form)
