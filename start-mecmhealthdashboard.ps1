<#
.SYNOPSIS
    WinForms front-end for MECM Environment Health Dashboard.

.DESCRIPTION
    Provides a GUI for monitoring MECM environment health across six areas:
    deployments, content distribution, distribution points, client health,
    inactive devices, and site component/system health.

    Features:
      - Summary cards with at-a-glance status for all health areas
      - Tabbed detail views with drill-down
      - Auto-refresh with configurable interval
      - Export to CSV or HTML
      - Dark mode / light mode

.EXAMPLE
    .\start-mecmhealthdashboard.ps1

.NOTES
    Requirements:
      - PowerShell 5.1
      - .NET Framework 4.8+
      - Windows Forms (System.Windows.Forms)
      - Configuration Manager console installed
      - SQL Server access (Invoke-Sqlcmd) for client health views

    ScriptName : start-mecmhealthdashboard.ps1
    Purpose    : WinForms front-end for MECM environment health monitoring
    Version    : 1.0.0
    Updated    : 2026-02-26
#>

param()

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

[System.Windows.Forms.Application]::EnableVisualStyles()
try { [System.Windows.Forms.Application]::SetCompatibleTextRenderingDefault($false) } catch { }

$moduleRoot = Join-Path $PSScriptRoot "Module"
Import-Module (Join-Path $moduleRoot "MECMHealthDashCommon.psd1") -Force

# Initialize tool logging
$toolLogFolder = Join-Path $PSScriptRoot "Logs"
if (-not (Test-Path -LiteralPath $toolLogFolder)) {
    New-Item -ItemType Directory -Path $toolLogFolder -Force | Out-Null
}
$toolLogPath = Join-Path $toolLogFolder ("HealthDash-{0}.log" -f (Get-Date -Format 'yyyyMMdd-HHmmss'))
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

    $hover = [System.Drawing.Color]::FromArgb(
        [Math]::Max(0, $BackColor.R - 18),
        [Math]::Max(0, $BackColor.G - 18),
        [Math]::Max(0, $BackColor.B - 18)
    )
    $down = [System.Drawing.Color]::FromArgb(
        [Math]::Max(0, $BackColor.R - 36),
        [Math]::Max(0, $BackColor.G - 36),
        [Math]::Max(0, $BackColor.B - 36)
    )

    $Button.FlatAppearance.MouseOverBackColor = $hover
    $Button.FlatAppearance.MouseDownBackColor = $down
}

function Enable-DoubleBuffer {
    param([Parameter(Mandatory)][System.Windows.Forms.Control]$Control)
    $prop = $Control.GetType().GetProperty("DoubleBuffered", [System.Reflection.BindingFlags] "Instance,NonPublic")
    if ($prop) { $prop.SetValue($Control, $true, $null) | Out-Null }
}

function Add-LogLine {
    param(
        [Parameter(Mandatory)][System.Windows.Forms.TextBox]$TextBox,
        [Parameter(Mandatory)][string]$Message
    )
    $ts = (Get-Date).ToString("HH:mm:ss")
    $line = "{0}  {1}" -f $ts, $Message

    if ([string]::IsNullOrWhiteSpace($TextBox.Text)) {
        $TextBox.Text = $line
    }
    else {
        $TextBox.AppendText([Environment]::NewLine + $line)
    }

    $TextBox.SelectionStart = $TextBox.TextLength
    $TextBox.ScrollToCaret()
}

function Save-WindowState {
    $statePath = Join-Path $PSScriptRoot "MECMHealthDash.windowstate.json"
    $state = @{
        X                = $form.Location.X
        Y                = $form.Location.Y
        Width            = $form.Size.Width
        Height           = $form.Size.Height
        Maximized        = ($form.WindowState -eq [System.Windows.Forms.FormWindowState]::Maximized)
        ActiveTab        = $tabMain.SelectedIndex
    }
    $state | ConvertTo-Json | Set-Content -LiteralPath $statePath -Encoding UTF8
}

function Restore-WindowState {
    $statePath = Join-Path $PSScriptRoot "MECMHealthDash.windowstate.json"
    if (-not (Test-Path -LiteralPath $statePath)) { return }

    try {
        $state = Get-Content -LiteralPath $statePath -Raw | ConvertFrom-Json
        if ($state.Maximized) {
            $form.WindowState = [System.Windows.Forms.FormWindowState]::Maximized
        } else {
            $form.Location = New-Object System.Drawing.Point($state.X, $state.Y)
            $form.Size = New-Object System.Drawing.Size($state.Width, $state.Height)
        }
        if ($null -ne $state.ActiveTab) {
            $tabMain.SelectedIndex = [int]$state.ActiveTab
        }
    } catch { }
}

# ---------------------------------------------------------------------------
# Preferences
# ---------------------------------------------------------------------------

function Get-HdPreferences {
    $prefsPath = Join-Path $PSScriptRoot "MECMHealthDash.prefs.json"
    $defaults = @{
        DarkMode             = $false
        SiteCode             = ''
        SMSProvider          = ''
        SQLServer            = ''
        AutoRefreshMinutes   = 15
        InactiveThresholdDays = 14
    }

    if (Test-Path -LiteralPath $prefsPath) {
        try {
            $loaded = Get-Content -LiteralPath $prefsPath -Raw | ConvertFrom-Json
            if ($null -ne $loaded.DarkMode)              { $defaults.DarkMode              = [bool]$loaded.DarkMode }
            if ($loaded.SiteCode)                        { $defaults.SiteCode              = $loaded.SiteCode }
            if ($loaded.SMSProvider)                     { $defaults.SMSProvider            = $loaded.SMSProvider }
            if ($loaded.SQLServer)                       { $defaults.SQLServer              = $loaded.SQLServer }
            if ($null -ne $loaded.AutoRefreshMinutes)    { $defaults.AutoRefreshMinutes     = [int]$loaded.AutoRefreshMinutes }
            if ($null -ne $loaded.InactiveThresholdDays) { $defaults.InactiveThresholdDays  = [int]$loaded.InactiveThresholdDays }
        } catch { }
    }

    return $defaults
}

function Save-HdPreferences {
    param([hashtable]$Prefs)
    $prefsPath = Join-Path $PSScriptRoot "MECMHealthDash.prefs.json"
    $Prefs | ConvertTo-Json | Set-Content -LiteralPath $prefsPath -Encoding UTF8
}

$script:Prefs = Get-HdPreferences

# ---------------------------------------------------------------------------
# Colors (theme-aware)
# ---------------------------------------------------------------------------

$clrAccent = [System.Drawing.Color]::FromArgb(0, 120, 212)

if ($script:Prefs.DarkMode) {
    $clrFormBg     = [System.Drawing.Color]::FromArgb(30, 30, 30)
    $clrPanelBg    = [System.Drawing.Color]::FromArgb(40, 40, 40)
    $clrHint       = [System.Drawing.Color]::FromArgb(140, 140, 140)
    $clrSubtitle   = [System.Drawing.Color]::FromArgb(180, 200, 220)
    $clrGridAlt    = [System.Drawing.Color]::FromArgb(48, 48, 48)
    $clrGridLine   = [System.Drawing.Color]::FromArgb(60, 60, 60)
    $clrDetailBg   = [System.Drawing.Color]::FromArgb(45, 45, 45)
    $clrSepLine    = [System.Drawing.Color]::FromArgb(55, 55, 55)
    $clrInputBdr   = [System.Drawing.Color]::FromArgb(70, 70, 70)
    $clrLogBg      = [System.Drawing.Color]::FromArgb(35, 35, 35)
    $clrLogFg      = [System.Drawing.Color]::FromArgb(200, 200, 200)
    $clrText       = [System.Drawing.Color]::FromArgb(220, 220, 220)
    $clrGridText   = [System.Drawing.Color]::FromArgb(220, 220, 220)
    $clrErrText    = [System.Drawing.Color]::FromArgb(255, 100, 100)
    $clrWarnText   = [System.Drawing.Color]::FromArgb(255, 200, 80)
    $clrOkText     = [System.Drawing.Color]::FromArgb(80, 200, 80)
    $clrCardGreen  = [System.Drawing.Color]::FromArgb(30, 60, 30)
    $clrCardYellow = [System.Drawing.Color]::FromArgb(60, 50, 20)
    $clrCardRed    = [System.Drawing.Color]::FromArgb(60, 25, 25)
} else {
    $clrFormBg     = [System.Drawing.Color]::FromArgb(245, 246, 248)
    $clrPanelBg    = [System.Drawing.Color]::White
    $clrHint       = [System.Drawing.Color]::FromArgb(140, 140, 140)
    $clrSubtitle   = [System.Drawing.Color]::FromArgb(220, 230, 245)
    $clrGridAlt    = [System.Drawing.Color]::FromArgb(248, 250, 252)
    $clrGridLine   = [System.Drawing.Color]::FromArgb(230, 230, 230)
    $clrDetailBg   = [System.Drawing.Color]::FromArgb(250, 250, 250)
    $clrSepLine    = [System.Drawing.Color]::FromArgb(218, 220, 224)
    $clrInputBdr   = [System.Drawing.Color]::FromArgb(200, 200, 200)
    $clrLogBg      = [System.Drawing.Color]::White
    $clrLogFg      = [System.Drawing.Color]::Black
    $clrText       = [System.Drawing.Color]::Black
    $clrGridText   = [System.Drawing.Color]::Black
    $clrErrText    = [System.Drawing.Color]::FromArgb(180, 0, 0)
    $clrWarnText   = [System.Drawing.Color]::FromArgb(180, 120, 0)
    $clrOkText     = [System.Drawing.Color]::FromArgb(34, 139, 34)
    $clrCardGreen  = [System.Drawing.Color]::FromArgb(220, 245, 220)
    $clrCardYellow = [System.Drawing.Color]::FromArgb(255, 248, 220)
    $clrCardRed    = [System.Drawing.Color]::FromArgb(255, 225, 225)
}

# Custom dark mode ToolStrip renderer
if ($script:Prefs.DarkMode) {
    if (-not ('DarkToolStripRenderer' -as [type])) {
        $rendererCs = (
            'using System.Drawing;',
            'using System.Windows.Forms;',
            'public class DarkToolStripRenderer : ToolStripProfessionalRenderer {',
            '    private Color _bg;',
            '    public DarkToolStripRenderer(Color bg) : base() { _bg = bg; }',
            '    protected override void OnRenderToolStripBorder(ToolStripRenderEventArgs e) { }',
            '    protected override void OnRenderToolStripBackground(ToolStripRenderEventArgs e) {',
            '        using (var b = new SolidBrush(_bg)) { e.Graphics.FillRectangle(b, e.AffectedBounds); }',
            '    }',
            '    protected override void OnRenderMenuItemBackground(ToolStripItemRenderEventArgs e) {',
            '        if (e.Item.Selected || e.Item.Pressed) {',
            '            using (var b = new SolidBrush(Color.FromArgb(60, 60, 60)))',
            '            { e.Graphics.FillRectangle(b, new Rectangle(Point.Empty, e.Item.Size)); }',
            '        }',
            '    }',
            '    protected override void OnRenderSeparator(ToolStripSeparatorRenderEventArgs e) {',
            '        int y = e.Item.Height / 2;',
            '        using (var p = new Pen(Color.FromArgb(70, 70, 70)))',
            '        { e.Graphics.DrawLine(p, 0, y, e.Item.Width, y); }',
            '    }',
            '    protected override void OnRenderImageMargin(ToolStripRenderEventArgs e) {',
            '        using (var b = new SolidBrush(_bg)) { e.Graphics.FillRectangle(b, e.AffectedBounds); }',
            '    }',
            '}'
        ) -join "`r`n"
        Add-Type -ReferencedAssemblies System.Windows.Forms, System.Drawing -TypeDefinition $rendererCs
    }
    $script:DarkRenderer = New-Object DarkToolStripRenderer($clrPanelBg)
}

# ---------------------------------------------------------------------------
# Dialogs
# ---------------------------------------------------------------------------

function Show-PreferencesDialog {
    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = "Preferences"
    $dlg.Size = New-Object System.Drawing.Size(440, 420)
    $dlg.MinimumSize = $dlg.Size
    $dlg.MaximumSize = $dlg.Size
    $dlg.StartPosition = "CenterParent"
    $dlg.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $dlg.MaximizeBox = $false
    $dlg.MinimizeBox = $false
    $dlg.ShowInTaskbar = $false
    $dlg.Font = New-Object System.Drawing.Font("Segoe UI", 9.5)
    $dlg.BackColor = $clrFormBg

    # Appearance
    $grpAppearance = New-Object System.Windows.Forms.GroupBox
    $grpAppearance.Text = "Appearance"
    $grpAppearance.SetBounds(16, 12, 392, 60)
    $grpAppearance.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $grpAppearance.ForeColor = $clrText
    $grpAppearance.BackColor = $clrFormBg
    if ($script:Prefs.DarkMode) { $grpAppearance.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat; $grpAppearance.ForeColor = $clrSepLine }
    $dlg.Controls.Add($grpAppearance)

    $chkDark = New-Object System.Windows.Forms.CheckBox
    $chkDark.Text = "Enable dark mode (requires restart)"
    $chkDark.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $chkDark.AutoSize = $true
    $chkDark.Location = New-Object System.Drawing.Point(14, 24)
    $chkDark.Checked = $script:Prefs.DarkMode
    $chkDark.ForeColor = $clrText
    $chkDark.BackColor = $clrFormBg
    if ($script:Prefs.DarkMode) { $chkDark.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat; $chkDark.ForeColor = [System.Drawing.Color]::FromArgb(170, 170, 170) }
    $grpAppearance.Controls.Add($chkDark)

    # MECM Connection
    $grpConn = New-Object System.Windows.Forms.GroupBox
    $grpConn.Text = "MECM Connection"
    $grpConn.SetBounds(16, 82, 392, 146)
    $grpConn.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $grpConn.ForeColor = $clrText
    $grpConn.BackColor = $clrFormBg
    if ($script:Prefs.DarkMode) { $grpConn.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat; $grpConn.ForeColor = $clrSepLine }
    $dlg.Controls.Add($grpConn)

    $lblSiteCode = New-Object System.Windows.Forms.Label
    $lblSiteCode.Text = "Site Code:"
    $lblSiteCode.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $lblSiteCode.Location = New-Object System.Drawing.Point(14, 30)
    $lblSiteCode.AutoSize = $true
    $lblSiteCode.ForeColor = $clrText
    $grpConn.Controls.Add($lblSiteCode)

    $txtSiteCodePref = New-Object System.Windows.Forms.TextBox
    $txtSiteCodePref.SetBounds(130, 27, 80, 24)
    $txtSiteCodePref.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $txtSiteCodePref.MaxLength = 3
    $txtSiteCodePref.Text = $script:Prefs.SiteCode
    $txtSiteCodePref.BackColor = $clrDetailBg
    $txtSiteCodePref.ForeColor = $clrText
    $txtSiteCodePref.BorderStyle = if ($script:Prefs.DarkMode) { [System.Windows.Forms.BorderStyle]::None } else { [System.Windows.Forms.BorderStyle]::FixedSingle }
    $grpConn.Controls.Add($txtSiteCodePref)

    $lblServer = New-Object System.Windows.Forms.Label
    $lblServer.Text = "SMS Provider:"
    $lblServer.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $lblServer.Location = New-Object System.Drawing.Point(14, 64)
    $lblServer.AutoSize = $true
    $lblServer.ForeColor = $clrText
    $grpConn.Controls.Add($lblServer)

    $txtServerPref = New-Object System.Windows.Forms.TextBox
    $txtServerPref.SetBounds(130, 61, 240, 24)
    $txtServerPref.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $txtServerPref.Text = $script:Prefs.SMSProvider
    $txtServerPref.BackColor = $clrDetailBg
    $txtServerPref.ForeColor = $clrText
    $txtServerPref.BorderStyle = if ($script:Prefs.DarkMode) { [System.Windows.Forms.BorderStyle]::None } else { [System.Windows.Forms.BorderStyle]::FixedSingle }
    $grpConn.Controls.Add($txtServerPref)

    $lblSQL = New-Object System.Windows.Forms.Label
    $lblSQL.Text = "SQL Server:"
    $lblSQL.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $lblSQL.Location = New-Object System.Drawing.Point(14, 98)
    $lblSQL.AutoSize = $true
    $lblSQL.ForeColor = $clrText
    $grpConn.Controls.Add($lblSQL)

    $txtSQLPref = New-Object System.Windows.Forms.TextBox
    $txtSQLPref.SetBounds(130, 95, 240, 24)
    $txtSQLPref.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $txtSQLPref.Text = $script:Prefs.SQLServer
    $txtSQLPref.BackColor = $clrDetailBg
    $txtSQLPref.ForeColor = $clrText
    $txtSQLPref.BorderStyle = if ($script:Prefs.DarkMode) { [System.Windows.Forms.BorderStyle]::None } else { [System.Windows.Forms.BorderStyle]::FixedSingle }
    $grpConn.Controls.Add($txtSQLPref)

    # Auto-Refresh
    $grpRefresh = New-Object System.Windows.Forms.GroupBox
    $grpRefresh.Text = "Auto-Refresh"
    $grpRefresh.SetBounds(16, 238, 392, 60)
    $grpRefresh.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $grpRefresh.ForeColor = $clrText
    $grpRefresh.BackColor = $clrFormBg
    if ($script:Prefs.DarkMode) { $grpRefresh.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat; $grpRefresh.ForeColor = $clrSepLine }
    $dlg.Controls.Add($grpRefresh)

    $lblInterval = New-Object System.Windows.Forms.Label
    $lblInterval.Text = "Interval (minutes):"
    $lblInterval.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $lblInterval.Location = New-Object System.Drawing.Point(14, 26)
    $lblInterval.AutoSize = $true
    $lblInterval.ForeColor = $clrText
    $grpRefresh.Controls.Add($lblInterval)

    $cboInterval = New-Object System.Windows.Forms.ComboBox
    $cboInterval.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    $cboInterval.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $cboInterval.SetBounds(150, 23, 70, 24)
    $cboInterval.BackColor = $clrDetailBg
    $cboInterval.ForeColor = $clrText
    [void]$cboInterval.Items.AddRange(@('5', '10', '15', '30', '60'))
    $currentIdx = $cboInterval.Items.IndexOf([string]$script:Prefs.AutoRefreshMinutes)
    $cboInterval.SelectedIndex = if ($currentIdx -ge 0) { $currentIdx } else { 2 }
    $grpRefresh.Controls.Add($cboInterval)

    # OK / Cancel
    $btnOK = New-Object System.Windows.Forms.Button
    $btnOK.Text = "OK"
    $btnOK.Size = New-Object System.Drawing.Size(90, 32)
    $btnOK.Location = New-Object System.Drawing.Point(228, 320)
    $btnOK.DialogResult = [System.Windows.Forms.DialogResult]::OK
    Set-ModernButtonStyle -Button $btnOK -BackColor $clrAccent
    $dlg.Controls.Add($btnOK)
    $dlg.AcceptButton = $btnOK

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = "Cancel"
    $btnCancel.Size = New-Object System.Drawing.Size(90, 32)
    $btnCancel.Location = New-Object System.Drawing.Point(326, 320)
    $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $btnCancel.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnCancel.FlatAppearance.BorderColor = $clrSepLine
    $btnCancel.ForeColor = $clrText
    $btnCancel.BackColor = $clrFormBg
    $dlg.Controls.Add($btnCancel)
    $dlg.CancelButton = $btnCancel

    if ($dlg.ShowDialog($form) -eq [System.Windows.Forms.DialogResult]::OK) {
        $darkChanged = ($chkDark.Checked -ne $script:Prefs.DarkMode)
        $script:Prefs.DarkMode              = $chkDark.Checked
        $script:Prefs.SiteCode              = $txtSiteCodePref.Text.Trim().ToUpper()
        $script:Prefs.SMSProvider           = $txtServerPref.Text.Trim()
        $script:Prefs.SQLServer             = $txtSQLPref.Text.Trim()
        $script:Prefs.AutoRefreshMinutes    = [int]$cboInterval.SelectedItem
        Save-HdPreferences -Prefs $script:Prefs

        # Update connection bar labels
        $lblSiteVal.Text   = if ($script:Prefs.SiteCode)    { $script:Prefs.SiteCode }    else { '(not set)' }
        $lblServerVal.Text = if ($script:Prefs.SMSProvider)  { $script:Prefs.SMSProvider }  else { '(not set)' }
        $lblSQLVal.Text    = if ($script:Prefs.SQLServer)    { $script:Prefs.SQLServer }    else { '(not set)' }

        # Update auto-refresh timer interval
        $script:RefreshTimer.Interval = $script:Prefs.AutoRefreshMinutes * 60 * 1000

        if ($darkChanged) {
            $restart = [System.Windows.Forms.MessageBox]::Show(
                "Theme change requires a restart. Restart now?",
                "Restart Required",
                [System.Windows.Forms.MessageBoxButtons]::YesNo,
                [System.Windows.Forms.MessageBoxIcon]::Question
            )
            if ($restart -eq [System.Windows.Forms.DialogResult]::Yes) {
                Start-Process powershell -ArgumentList @('-ExecutionPolicy', 'Bypass', '-File', $PSCommandPath)
                $form.Close()
            }
        }
    }

    $dlg.Dispose()
}

function Show-AboutDialog {
    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = "About MECM Health Dashboard"
    $dlg.Size = New-Object System.Drawing.Size(460, 320)
    $dlg.MinimumSize = $dlg.Size
    $dlg.MaximumSize = $dlg.Size
    $dlg.StartPosition = "CenterParent"
    $dlg.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $dlg.MaximizeBox = $false
    $dlg.MinimizeBox = $false
    $dlg.ShowInTaskbar = $false
    $dlg.Font = New-Object System.Drawing.Font("Segoe UI", 9.5)
    $dlg.BackColor = $clrFormBg

    $lblAboutTitle = New-Object System.Windows.Forms.Label
    $lblAboutTitle.Text = "MECM Health Dashboard"
    $lblAboutTitle.Font = New-Object System.Drawing.Font("Segoe UI", 14, [System.Drawing.FontStyle]::Bold)
    $lblAboutTitle.ForeColor = $clrAccent
    $lblAboutTitle.AutoSize = $true
    $lblAboutTitle.BackColor = $clrFormBg
    $lblAboutTitle.Location = New-Object System.Drawing.Point(90, 30)
    $dlg.Controls.Add($lblAboutTitle)

    $lblVersion = New-Object System.Windows.Forms.Label
    $lblVersion.Text = "MECM Health Dashboard v1.0.0"
    $lblVersion.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $lblVersion.ForeColor = $clrText
    $lblVersion.AutoSize = $true
    $lblVersion.BackColor = $clrFormBg
    $lblVersion.Location = New-Object System.Drawing.Point(110, 60)
    $dlg.Controls.Add($lblVersion)

    $lblDesc = New-Object System.Windows.Forms.Label
    $lblDesc.Text = ("Unified MECM environment health monitoring." +
        " View deployment status, content distribution, DP availability," +
        " client health, inactive devices, and site component status" +
        " in a single dashboard with auto-refresh.")
    $lblDesc.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $lblDesc.ForeColor = $clrText
    $lblDesc.SetBounds(30, 100, 390, 80)
    $lblDesc.BackColor = $clrFormBg
    $lblDesc.TextAlign = [System.Drawing.ContentAlignment]::TopCenter
    $dlg.Controls.Add($lblDesc)

    $lblCopyright = New-Object System.Windows.Forms.Label
    $lblCopyright.Text = "(c) 2026 - All rights reserved"
    $lblCopyright.Font = New-Object System.Drawing.Font("Segoe UI", 8.5, [System.Drawing.FontStyle]::Italic)
    $lblCopyright.ForeColor = $clrHint
    $lblCopyright.AutoSize = $true
    $lblCopyright.BackColor = $clrFormBg
    $lblCopyright.Location = New-Object System.Drawing.Point(142, 200)
    $dlg.Controls.Add($lblCopyright)

    $btnClose = New-Object System.Windows.Forms.Button
    $btnClose.Text = "OK"
    $btnClose.Size = New-Object System.Drawing.Size(90, 32)
    $btnClose.Location = New-Object System.Drawing.Point(175, 240)
    $btnClose.DialogResult = [System.Windows.Forms.DialogResult]::OK
    Set-ModernButtonStyle -Button $btnClose -BackColor $clrAccent
    $dlg.Controls.Add($btnClose)
    $dlg.AcceptButton = $btnClose

    [void]$dlg.ShowDialog($form)
    $dlg.Dispose()
}

# ---------------------------------------------------------------------------
# Form
# ---------------------------------------------------------------------------

$form = New-Object System.Windows.Forms.Form
$form.Text = "MECM Health Dashboard"
$form.StartPosition = "CenterScreen"
$form.Size = New-Object System.Drawing.Size(1440, 950)
$form.MinimumSize = New-Object System.Drawing.Size(1100, 800)
$form.AutoScaleMode = [System.Windows.Forms.AutoScaleMode]::Dpi
$form.Font = New-Object System.Drawing.Font("Segoe UI", 9.5)
$form.BackColor = $clrFormBg
$form.Icon = [System.Drawing.SystemIcons]::Application

# ---------------------------------------------------------------------------
# Auto-refresh timer
# ---------------------------------------------------------------------------

$script:RefreshTimer = New-Object System.Windows.Forms.Timer
$script:RefreshTimer.Interval = $script:Prefs.AutoRefreshMinutes * 60 * 1000
$script:RefreshTimer.Enabled = $false
$script:NextRefreshTime = $null
$script:AutoRefreshPaused = $false

# Countdown timer (updates status bar every second)
$script:CountdownTimer = New-Object System.Windows.Forms.Timer
$script:CountdownTimer.Interval = 1000
$script:CountdownTimer.Enabled = $false

# ---------------------------------------------------------------------------
# Menu bar
# ---------------------------------------------------------------------------

$menuStrip = New-Object System.Windows.Forms.MenuStrip
$menuStrip.Dock = [System.Windows.Forms.DockStyle]::Top
$menuStrip.BackColor = $clrPanelBg
$menuStrip.ForeColor = $clrText
if ($script:DarkRenderer) {
    $menuStrip.Renderer = $script:DarkRenderer
} else {
    $menuStrip.RenderMode = [System.Windows.Forms.ToolStripRenderMode]::System
}
$menuStrip.Padding = New-Object System.Windows.Forms.Padding(4, 2, 0, 0)

# File menu
$mnuFile = New-Object System.Windows.Forms.ToolStripMenuItem("&File")
$mnuFilePrefs = New-Object System.Windows.Forms.ToolStripMenuItem("&Preferences...")
$mnuFilePrefs.Add_Click({ Show-PreferencesDialog })
$mnuFileSep = New-Object System.Windows.Forms.ToolStripSeparator
$mnuFileExit = New-Object System.Windows.Forms.ToolStripMenuItem("E&xit")
$mnuFileExit.Add_Click({ $form.Close() })
[void]$mnuFile.DropDownItems.Add($mnuFilePrefs)
[void]$mnuFile.DropDownItems.Add($mnuFileSep)
[void]$mnuFile.DropDownItems.Add($mnuFileExit)

# Actions menu
$mnuActions = New-Object System.Windows.Forms.ToolStripMenuItem("&Actions")
$mnuActRefresh = New-Object System.Windows.Forms.ToolStripMenuItem("&Refresh Now")
$mnuActRefresh.Add_Click({ Invoke-RefreshAll })
$mnuActPause = New-Object System.Windows.Forms.ToolStripMenuItem("&Pause Auto-Refresh")
$mnuActPause.Add_Click({
    $script:AutoRefreshPaused = $true
    $script:RefreshTimer.Stop()
    $script:CountdownTimer.Stop()
    $mnuActPause.Enabled = $false
    $mnuActResume.Enabled = $true
    Add-LogLine -TextBox $txtLog -Message "Auto-refresh paused"
    Update-StatusBar
})
$mnuActResume = New-Object System.Windows.Forms.ToolStripMenuItem("&Resume Auto-Refresh")
$mnuActResume.Enabled = $false
$mnuActResume.Add_Click({
    $script:AutoRefreshPaused = $false
    $script:NextRefreshTime = (Get-Date).AddMilliseconds($script:RefreshTimer.Interval)
    $script:RefreshTimer.Start()
    $script:CountdownTimer.Start()
    $mnuActPause.Enabled = $true
    $mnuActResume.Enabled = $false
    Add-LogLine -TextBox $txtLog -Message "Auto-refresh resumed"
    Update-StatusBar
})
[void]$mnuActions.DropDownItems.Add($mnuActRefresh)
[void]$mnuActions.DropDownItems.Add((New-Object System.Windows.Forms.ToolStripSeparator))
[void]$mnuActions.DropDownItems.Add($mnuActPause)
[void]$mnuActions.DropDownItems.Add($mnuActResume)

# View menu
$mnuView = New-Object System.Windows.Forms.ToolStripMenuItem("&View")
$tabNames = @('&Deployments', '&Content', 'D&Ps', 'C&lient Health', '&Inactive Devices', '&Site Health')
for ($i = 0; $i -lt $tabNames.Count; $i++) {
    $idx = $i
    $mnuItem = New-Object System.Windows.Forms.ToolStripMenuItem($tabNames[$i])
    $mnuItem.Add_Click({ $tabMain.SelectedIndex = $idx }.GetNewClosure())
    [void]$mnuView.DropDownItems.Add($mnuItem)
}

# Help menu
$mnuHelp = New-Object System.Windows.Forms.ToolStripMenuItem("&Help")
$mnuHelpAbout = New-Object System.Windows.Forms.ToolStripMenuItem("&About...")
$mnuHelpAbout.Add_Click({ Show-AboutDialog })
[void]$mnuHelp.DropDownItems.Add($mnuHelpAbout)

[void]$menuStrip.Items.Add($mnuFile)
[void]$menuStrip.Items.Add($mnuActions)
[void]$menuStrip.Items.Add($mnuView)
[void]$menuStrip.Items.Add($mnuHelp)
$form.MainMenuStrip = $menuStrip

# ---------------------------------------------------------------------------
# StatusStrip (Dock:Bottom - add FIRST so it stays at very bottom)
# ---------------------------------------------------------------------------

$status = New-Object System.Windows.Forms.StatusStrip
$status.BackColor = if ($script:Prefs.DarkMode) { [System.Drawing.Color]::FromArgb(45, 45, 45) } else { [System.Drawing.Color]::FromArgb(240, 240, 240) }
$status.ForeColor = $clrText
$status.Dock = [System.Windows.Forms.DockStyle]::Bottom
if ($script:DarkRenderer) {
    $status.Renderer = $script:DarkRenderer
} else {
    $status.RenderMode = [System.Windows.Forms.ToolStripRenderMode]::System
}
$status.SizingGrip = $false
$statusLabel = New-Object System.Windows.Forms.ToolStripStatusLabel
$statusLabel.Text = "Disconnected. Configure site in File > Preferences, then click Refresh Now."
$statusLabel.ForeColor = $clrText
$status.Items.Add($statusLabel) | Out-Null
$form.Controls.Add($status)

# ---------------------------------------------------------------------------
# Log console panel (Dock:Bottom)
# ---------------------------------------------------------------------------

$pnlLog = New-Object System.Windows.Forms.Panel
$pnlLog.Dock = [System.Windows.Forms.DockStyle]::Bottom
$pnlLog.Height = 95
$pnlLog.Padding = New-Object System.Windows.Forms.Padding(12, 4, 12, 6)
$pnlLog.BackColor = $clrFormBg
$form.Controls.Add($pnlLog)

$txtLog = New-Object System.Windows.Forms.TextBox
$txtLog.Multiline = $true
$txtLog.ReadOnly = $true
$txtLog.ScrollBars = if ($script:Prefs.DarkMode) { [System.Windows.Forms.ScrollBars]::None } else { [System.Windows.Forms.ScrollBars]::Vertical }
$txtLog.Font = New-Object System.Drawing.Font("Consolas", 9)
$txtLog.BackColor = $clrLogBg
$txtLog.ForeColor = $clrLogFg
$txtLog.WordWrap = $true
$txtLog.Dock = [System.Windows.Forms.DockStyle]::Fill
$txtLog.BorderStyle = [System.Windows.Forms.BorderStyle]::None
$pnlLog.Controls.Add($txtLog)

# ---------------------------------------------------------------------------
# Button panel (Dock:Bottom)
# ---------------------------------------------------------------------------

$pnlButtons = New-Object System.Windows.Forms.Panel
$pnlButtons.Dock = [System.Windows.Forms.DockStyle]::Bottom
$pnlButtons.Height = 56
$pnlButtons.Padding = New-Object System.Windows.Forms.Padding(12, 10, 12, 4)
$pnlButtons.BackColor = $clrFormBg
$form.Controls.Add($pnlButtons)

$pnlSepButtons = New-Object System.Windows.Forms.Panel
$pnlSepButtons.Dock = [System.Windows.Forms.DockStyle]::Top
$pnlSepButtons.Height = 1
$pnlSepButtons.BackColor = $clrSepLine
$pnlButtons.Controls.Add($pnlSepButtons)

$flowButtons = New-Object System.Windows.Forms.FlowLayoutPanel
$flowButtons.Dock = [System.Windows.Forms.DockStyle]::Fill
$flowButtons.FlowDirection = [System.Windows.Forms.FlowDirection]::LeftToRight
$flowButtons.WrapContents = $false
$flowButtons.BackColor = $clrFormBg
$pnlButtons.Controls.Add($flowButtons)

$btnExportCsv = New-Object System.Windows.Forms.Button
$btnExportCsv.Text = "Export CSV"
$btnExportCsv.Font = New-Object System.Drawing.Font("Segoe UI", 9.5, [System.Drawing.FontStyle]::Bold)
$btnExportCsv.Size = New-Object System.Drawing.Size(120, 38)
$btnExportCsv.Margin = New-Object System.Windows.Forms.Padding(0, 0, 10, 0)
Set-ModernButtonStyle -Button $btnExportCsv -BackColor ([System.Drawing.Color]::FromArgb(34, 139, 34))
$flowButtons.Controls.Add($btnExportCsv)

$btnExportHtml = New-Object System.Windows.Forms.Button
$btnExportHtml.Text = "Export HTML"
$btnExportHtml.Font = New-Object System.Drawing.Font("Segoe UI", 9.5, [System.Drawing.FontStyle]::Bold)
$btnExportHtml.Size = New-Object System.Drawing.Size(120, 38)
$btnExportHtml.Margin = New-Object System.Windows.Forms.Padding(0, 0, 10, 0)
Set-ModernButtonStyle -Button $btnExportHtml -BackColor ([System.Drawing.Color]::FromArgb(34, 139, 34))
$flowButtons.Controls.Add($btnExportHtml)

# ---------------------------------------------------------------------------
# Header panel (Dock:Top)
# ---------------------------------------------------------------------------

$pnlHeader = New-Object System.Windows.Forms.Panel
$pnlHeader.Dock = [System.Windows.Forms.DockStyle]::Top
$pnlHeader.Height = 60
$pnlHeader.BackColor = $clrAccent
$pnlHeader.Padding = New-Object System.Windows.Forms.Padding(16, 0, 16, 0)
$form.Controls.Add($pnlHeader)

$lblTitle = New-Object System.Windows.Forms.Label
$lblTitle.Text = "MECM Health Dashboard"
$lblTitle.Font = New-Object System.Drawing.Font("Segoe UI", 17, [System.Drawing.FontStyle]::Bold)
$lblTitle.ForeColor = [System.Drawing.Color]::White
$lblTitle.AutoSize = $true
$lblTitle.BackColor = [System.Drawing.Color]::Transparent
$lblTitle.Location = New-Object System.Drawing.Point(16, 8)
$pnlHeader.Controls.Add($lblTitle)

$lblSubtitle = New-Object System.Windows.Forms.Label
$lblSubtitle.Text = "Environment Health Overview"
$lblSubtitle.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$lblSubtitle.ForeColor = $clrSubtitle
$lblSubtitle.AutoSize = $true
$lblSubtitle.BackColor = [System.Drawing.Color]::Transparent
$lblSubtitle.Location = New-Object System.Drawing.Point(18, 36)
$pnlHeader.Controls.Add($lblSubtitle)

# ---------------------------------------------------------------------------
# Connection bar (Dock:Top)
# ---------------------------------------------------------------------------

$pnlConnBar = New-Object System.Windows.Forms.Panel
$pnlConnBar.Dock = [System.Windows.Forms.DockStyle]::Top
$pnlConnBar.Height = 36
$pnlConnBar.BackColor = $clrPanelBg
$pnlConnBar.Padding = New-Object System.Windows.Forms.Padding(12, 6, 12, 6)
$form.Controls.Add($pnlConnBar)

$flowConn = New-Object System.Windows.Forms.FlowLayoutPanel
$flowConn.Dock = [System.Windows.Forms.DockStyle]::Fill
$flowConn.FlowDirection = [System.Windows.Forms.FlowDirection]::LeftToRight
$flowConn.WrapContents = $false
$flowConn.BackColor = $clrPanelBg
$pnlConnBar.Controls.Add($flowConn)

$lblSiteLabel = New-Object System.Windows.Forms.Label
$lblSiteLabel.Text = "Site:"
$lblSiteLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$lblSiteLabel.AutoSize = $true
$lblSiteLabel.Margin = New-Object System.Windows.Forms.Padding(0, 3, 2, 0)
$lblSiteLabel.ForeColor = $clrText
$lblSiteLabel.BackColor = $clrPanelBg
$flowConn.Controls.Add($lblSiteLabel)

$lblSiteVal = New-Object System.Windows.Forms.Label
$lblSiteVal.Text = if ($script:Prefs.SiteCode) { $script:Prefs.SiteCode } else { '(not set)' }
$lblSiteVal.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$lblSiteVal.AutoSize = $true
$lblSiteVal.Margin = New-Object System.Windows.Forms.Padding(0, 3, 16, 0)
$lblSiteVal.ForeColor = if ($script:Prefs.SiteCode) { $clrAccent } else { $clrHint }
$lblSiteVal.BackColor = $clrPanelBg
$flowConn.Controls.Add($lblSiteVal)

$lblServerLabel = New-Object System.Windows.Forms.Label
$lblServerLabel.Text = "Server:"
$lblServerLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$lblServerLabel.AutoSize = $true
$lblServerLabel.Margin = New-Object System.Windows.Forms.Padding(0, 3, 2, 0)
$lblServerLabel.ForeColor = $clrText
$lblServerLabel.BackColor = $clrPanelBg
$flowConn.Controls.Add($lblServerLabel)

$lblServerVal = New-Object System.Windows.Forms.Label
$lblServerVal.Text = if ($script:Prefs.SMSProvider) { $script:Prefs.SMSProvider } else { '(not set)' }
$lblServerVal.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$lblServerVal.AutoSize = $true
$lblServerVal.Margin = New-Object System.Windows.Forms.Padding(0, 3, 16, 0)
$lblServerVal.ForeColor = if ($script:Prefs.SMSProvider) { $clrAccent } else { $clrHint }
$lblServerVal.BackColor = $clrPanelBg
$flowConn.Controls.Add($lblServerVal)

$lblSQLLabel = New-Object System.Windows.Forms.Label
$lblSQLLabel.Text = "SQL:"
$lblSQLLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$lblSQLLabel.AutoSize = $true
$lblSQLLabel.Margin = New-Object System.Windows.Forms.Padding(0, 3, 2, 0)
$lblSQLLabel.ForeColor = $clrText
$lblSQLLabel.BackColor = $clrPanelBg
$flowConn.Controls.Add($lblSQLLabel)

$lblSQLVal = New-Object System.Windows.Forms.Label
$lblSQLVal.Text = if ($script:Prefs.SQLServer) { $script:Prefs.SQLServer } else { '(not set)' }
$lblSQLVal.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$lblSQLVal.AutoSize = $true
$lblSQLVal.Margin = New-Object System.Windows.Forms.Padding(0, 3, 16, 0)
$lblSQLVal.ForeColor = if ($script:Prefs.SQLServer) { $clrAccent } else { $clrHint }
$lblSQLVal.BackColor = $clrPanelBg
$flowConn.Controls.Add($lblSQLVal)

$lblConnStatus = New-Object System.Windows.Forms.Label
$lblConnStatus.Text = "Disconnected"
$lblConnStatus.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Italic)
$lblConnStatus.AutoSize = $true
$lblConnStatus.Margin = New-Object System.Windows.Forms.Padding(0, 3, 20, 0)
$lblConnStatus.ForeColor = $clrHint
$lblConnStatus.BackColor = $clrPanelBg
$flowConn.Controls.Add($lblConnStatus)

$btnRefreshAll = New-Object System.Windows.Forms.Button
$btnRefreshAll.Text = "Refresh Now"
$btnRefreshAll.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$btnRefreshAll.Size = New-Object System.Drawing.Size(110, 24)
$btnRefreshAll.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 0)
Set-ModernButtonStyle -Button $btnRefreshAll -BackColor $clrAccent
$flowConn.Controls.Add($btnRefreshAll)

# Separator below connection bar
$pnlSep1 = New-Object System.Windows.Forms.Panel
$pnlSep1.Dock = [System.Windows.Forms.DockStyle]::Top
$pnlSep1.Height = 1
$pnlSep1.BackColor = $clrSepLine
$form.Controls.Add($pnlSep1)

# ---------------------------------------------------------------------------
# Summary cards panel (Dock:Top)
# ---------------------------------------------------------------------------

$pnlCards = New-Object System.Windows.Forms.Panel
$pnlCards.Dock = [System.Windows.Forms.DockStyle]::Top
$pnlCards.Height = 110
$pnlCards.BackColor = $clrFormBg
$pnlCards.Padding = New-Object System.Windows.Forms.Padding(12, 8, 12, 4)
$form.Controls.Add($pnlCards)

$flowCards = New-Object System.Windows.Forms.FlowLayoutPanel
$flowCards.Dock = [System.Windows.Forms.DockStyle]::Fill
$flowCards.FlowDirection = [System.Windows.Forms.FlowDirection]::LeftToRight
$flowCards.WrapContents = $true
$flowCards.BackColor = $clrFormBg
$pnlCards.Controls.Add($flowCards)

# Card creation helper
function New-SummaryCard {
    param(
        [string]$Title,
        [int]$TabIndex
    )

    $card = New-Object System.Windows.Forms.Panel
    $card.Size = New-Object System.Drawing.Size(200, 44)
    $card.Margin = New-Object System.Windows.Forms.Padding(4, 2, 4, 2)
    $card.BackColor = $clrPanelBg
    $card.BorderStyle = [System.Windows.Forms.BorderStyle]::None
    $card.Cursor = [System.Windows.Forms.Cursors]::Hand
    $card.Tag = $TabIndex

    # Left color indicator bar
    $bar = New-Object System.Windows.Forms.Panel
    $bar.Dock = [System.Windows.Forms.DockStyle]::Left
    $bar.Width = 4
    $bar.BackColor = $clrHint
    $card.Controls.Add($bar)

    $lblCardTitle = New-Object System.Windows.Forms.Label
    $lblCardTitle.Text = $Title
    $lblCardTitle.Font = New-Object System.Drawing.Font("Segoe UI", 8.5, [System.Drawing.FontStyle]::Bold)
    $lblCardTitle.ForeColor = $clrText
    $lblCardTitle.AutoSize = $true
    $lblCardTitle.Location = New-Object System.Drawing.Point(10, 4)
    $lblCardTitle.BackColor = [System.Drawing.Color]::Transparent
    $card.Controls.Add($lblCardTitle)

    $lblCardValue = New-Object System.Windows.Forms.Label
    $lblCardValue.Text = "--"
    $lblCardValue.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $lblCardValue.ForeColor = $clrHint
    $lblCardValue.AutoSize = $true
    $lblCardValue.Location = New-Object System.Drawing.Point(10, 22)
    $lblCardValue.BackColor = [System.Drawing.Color]::Transparent
    $lblCardValue.Tag = "value"
    $card.Controls.Add($lblCardValue)

    # Click on card or its children switches tab
    # Do NOT use GetNewClosure() here - $tabMain is created later and must resolve at click time
    $clickHandler = { $tabMain.SelectedIndex = [int]$this.Parent.Tag }
    $cardClickHandler = { $tabMain.SelectedIndex = [int]$this.Tag }
    $card.Add_Click($cardClickHandler)
    $lblCardTitle.Add_Click($clickHandler)
    $lblCardValue.Add_Click($clickHandler)

    return $card
}

$cardDeploy   = New-SummaryCard -Title "Deployments"      -TabIndex 0
$cardContent  = New-SummaryCard -Title "Content"           -TabIndex 1
$cardDPs      = New-SummaryCard -Title "Distribution Points" -TabIndex 2
$cardClients  = New-SummaryCard -Title "Client Health"     -TabIndex 3
$cardDevices  = New-SummaryCard -Title "Inactive Devices"  -TabIndex 4
$cardSite     = New-SummaryCard -Title "Site Health"       -TabIndex 5

$flowCards.Controls.Add($cardDeploy)
$flowCards.Controls.Add($cardContent)
$flowCards.Controls.Add($cardDPs)
$flowCards.Controls.Add($cardClients)
$flowCards.Controls.Add($cardDevices)
$flowCards.Controls.Add($cardSite)

# Card update helper
function Update-Card {
    param(
        [System.Windows.Forms.Panel]$Card,
        [string]$ValueText,
        [string]$Severity   # 'ok', 'warn', 'critical'
    )

    $bar = $Card.Controls[0]
    $valLabel = $Card.Controls | Where-Object { $_.Tag -eq 'value' }

    switch ($Severity) {
        'ok'       { $bar.BackColor = $clrOkText;   $Card.BackColor = $clrCardGreen;  if ($valLabel) { $valLabel.ForeColor = $clrOkText } }
        'warn'     { $bar.BackColor = $clrWarnText;  $Card.BackColor = $clrCardYellow; if ($valLabel) { $valLabel.ForeColor = $clrWarnText } }
        'critical' { $bar.BackColor = $clrErrText;   $Card.BackColor = $clrCardRed;    if ($valLabel) { $valLabel.ForeColor = $clrErrText } }
        default    { $bar.BackColor = $clrHint;      $Card.BackColor = $clrPanelBg;    if ($valLabel) { $valLabel.ForeColor = $clrHint } }
    }

    if ($valLabel) { $valLabel.Text = $ValueText }
}

# Separator below cards
$pnlSep2 = New-Object System.Windows.Forms.Panel
$pnlSep2.Dock = [System.Windows.Forms.DockStyle]::Top
$pnlSep2.Height = 1
$pnlSep2.BackColor = $clrSepLine
$form.Controls.Add($pnlSep2)

# ---------------------------------------------------------------------------
# Filter bar (Dock:Top)
# ---------------------------------------------------------------------------

$pnlFilter = New-Object System.Windows.Forms.Panel
$pnlFilter.Dock = [System.Windows.Forms.DockStyle]::Top
$pnlFilter.Height = 44
$pnlFilter.BackColor = $clrPanelBg
$pnlFilter.Padding = New-Object System.Windows.Forms.Padding(12, 6, 12, 6)
$form.Controls.Add($pnlFilter)

$flowFilter = New-Object System.Windows.Forms.FlowLayoutPanel
$flowFilter.Dock = [System.Windows.Forms.DockStyle]::Fill
$flowFilter.FlowDirection = [System.Windows.Forms.FlowDirection]::LeftToRight
$flowFilter.WrapContents = $false
$flowFilter.BackColor = $clrPanelBg
$pnlFilter.Controls.Add($flowFilter)

$lblStatusFilter = New-Object System.Windows.Forms.Label
$lblStatusFilter.Text = "Status:"
$lblStatusFilter.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$lblStatusFilter.AutoSize = $true
$lblStatusFilter.Margin = New-Object System.Windows.Forms.Padding(0, 6, 4, 0)
$lblStatusFilter.ForeColor = $clrText
$lblStatusFilter.BackColor = $clrPanelBg
$flowFilter.Controls.Add($lblStatusFilter)

$cboStatus = New-Object System.Windows.Forms.ComboBox
$cboStatus.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$cboStatus.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$cboStatus.Width = 140
$cboStatus.Margin = New-Object System.Windows.Forms.Padding(0, 2, 16, 0)
$cboStatus.BackColor = $clrDetailBg
$cboStatus.ForeColor = $clrText
$cboStatus.FlatStyle = if ($script:Prefs.DarkMode) { [System.Windows.Forms.FlatStyle]::Flat } else { [System.Windows.Forms.FlatStyle]::Standard }
[void]$cboStatus.Items.AddRange(@('All', 'Failed/Error', 'Warning/InProgress', 'OK/Healthy'))
$cboStatus.SelectedIndex = 0
$flowFilter.Controls.Add($cboStatus)

$lblFilter = New-Object System.Windows.Forms.Label
$lblFilter.Text = "Filter:"
$lblFilter.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$lblFilter.AutoSize = $true
$lblFilter.Margin = New-Object System.Windows.Forms.Padding(0, 6, 4, 0)
$lblFilter.ForeColor = $clrText
$lblFilter.BackColor = $clrPanelBg
$flowFilter.Controls.Add($lblFilter)

$txtFilter = New-Object System.Windows.Forms.TextBox
$txtFilter.Font = New-Object System.Drawing.Font("Segoe UI", 9.5)
$txtFilter.Width = 260
$txtFilter.Margin = New-Object System.Windows.Forms.Padding(0, 2, 0, 0)
$txtFilter.BorderStyle = if ($script:Prefs.DarkMode) { [System.Windows.Forms.BorderStyle]::None } else { [System.Windows.Forms.BorderStyle]::FixedSingle }
$txtFilter.BackColor = $clrDetailBg
$txtFilter.ForeColor = $clrText
$flowFilter.Controls.Add($txtFilter)

# Separator below filter bar
$pnlSep3 = New-Object System.Windows.Forms.Panel
$pnlSep3.Dock = [System.Windows.Forms.DockStyle]::Top
$pnlSep3.Height = 1
$pnlSep3.BackColor = $clrSepLine
$form.Controls.Add($pnlSep3)

# ---------------------------------------------------------------------------
# Helper: Create a themed DataGridView
# ---------------------------------------------------------------------------

function New-ThemedGrid {
    param([switch]$MultiSelect)

    $g = New-Object System.Windows.Forms.DataGridView
    $g.Dock = [System.Windows.Forms.DockStyle]::Fill
    $g.ReadOnly = $true
    $g.AllowUserToAddRows = $false
    $g.AllowUserToDeleteRows = $false
    $g.AllowUserToResizeRows = $false
    $g.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
    $g.MultiSelect = [bool]$MultiSelect
    $g.AutoGenerateColumns = $false
    $g.RowHeadersVisible = $false
    $g.BackgroundColor = $clrPanelBg
    $g.BorderStyle = [System.Windows.Forms.BorderStyle]::None
    $g.CellBorderStyle = [System.Windows.Forms.DataGridViewCellBorderStyle]::SingleHorizontal
    $g.GridColor = $clrGridLine
    $g.ColumnHeadersDefaultCellStyle.BackColor = $clrAccent
    $g.ColumnHeadersDefaultCellStyle.ForeColor = [System.Drawing.Color]::White
    $g.ColumnHeadersDefaultCellStyle.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Bold)
    $g.ColumnHeadersDefaultCellStyle.Padding = New-Object System.Windows.Forms.Padding(4)
    $g.ColumnHeadersDefaultCellStyle.WrapMode = [System.Windows.Forms.DataGridViewTriState]::False
    $g.ColumnHeadersHeightSizeMode = [System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode]::DisableResizing
    $g.ColumnHeadersHeight = 32
    $g.ColumnHeadersBorderStyle = [System.Windows.Forms.DataGridViewHeaderBorderStyle]::Single
    $g.EnableHeadersVisualStyles = $false
    $g.DefaultCellStyle.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $g.DefaultCellStyle.ForeColor = $clrGridText
    $g.DefaultCellStyle.BackColor = $clrPanelBg
    $g.DefaultCellStyle.Padding = New-Object System.Windows.Forms.Padding(2)
    $g.DefaultCellStyle.SelectionBackColor = if ($script:Prefs.DarkMode) { [System.Drawing.Color]::FromArgb(38, 79, 120) } else { [System.Drawing.Color]::FromArgb(0, 120, 215) }
    $g.DefaultCellStyle.SelectionForeColor = [System.Drawing.Color]::White
    $g.RowTemplate.Height = 26
    $g.AlternatingRowsDefaultCellStyle.BackColor = $clrGridAlt

    Enable-DoubleBuffer -Control $g
    return $g
}

# ---------------------------------------------------------------------------
# TabControl (Fill)
# ---------------------------------------------------------------------------

$tabMain = New-Object System.Windows.Forms.TabControl
$tabMain.Dock = [System.Windows.Forms.DockStyle]::Fill
$tabMain.Font = New-Object System.Drawing.Font("Segoe UI", 9.5, [System.Drawing.FontStyle]::Bold)
$tabMain.SizeMode = [System.Windows.Forms.TabSizeMode]::Fixed
$tabMain.ItemSize = New-Object System.Drawing.Size(140, 30)
$tabMain.DrawMode = [System.Windows.Forms.TabDrawMode]::OwnerDrawFixed
$tabMain.Add_DrawItem({
    param($s, $e)
    $e.Graphics.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::ClearTypeGridFit
    $tab = $s.TabPages[$e.Index]
    $isSelected = ($s.SelectedIndex -eq $e.Index)

    # Background
    $bgColor = if ($script:Prefs.DarkMode) {
        if ($isSelected) { $clrAccent } else { $clrPanelBg }
    } else {
        if ($isSelected) { $clrAccent } else { [System.Drawing.Color]::FromArgb(240, 240, 240) }
    }
    $fgColor = if ($isSelected) { [System.Drawing.Color]::White } else { $clrText }

    $bgBrush = New-Object System.Drawing.SolidBrush($bgColor)
    $e.Graphics.FillRectangle($bgBrush, $e.Bounds)

    # Font: bold always, 8pt
    $font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Bold)

    # Bottom-left anchor: Near horizontal, Far vertical with left padding
    $sf = New-Object System.Drawing.StringFormat
    $sf.Alignment = [System.Drawing.StringAlignment]::Near
    $sf.LineAlignment = [System.Drawing.StringAlignment]::Far
    $sf.FormatFlags = [System.Drawing.StringFormatFlags]::NoWrap

    # Inset rect for left padding and bottom padding
    $textRect = New-Object System.Drawing.RectangleF(
        ($e.Bounds.X + 8),
        $e.Bounds.Y,
        ($e.Bounds.Width - 12),
        ($e.Bounds.Height - 3)
    )

    $textBrush = New-Object System.Drawing.SolidBrush($fgColor)
    $e.Graphics.DrawString($tab.Text, $font, $textBrush, $textRect, $sf)

    $bgBrush.Dispose(); $textBrush.Dispose(); $font.Dispose(); $sf.Dispose()
})

$form.Controls.Add($tabMain)

# ===================== TAB 0: Deployments =====================

$tabDeploy = New-Object System.Windows.Forms.TabPage
$tabDeploy.Text = "Deployments"
$tabDeploy.BackColor = $clrFormBg
$tabMain.TabPages.Add($tabDeploy)

$splitDeploy = New-Object System.Windows.Forms.SplitContainer
$splitDeploy.Dock = [System.Windows.Forms.DockStyle]::Fill
$splitDeploy.Orientation = [System.Windows.Forms.Orientation]::Horizontal
$splitDeploy.SplitterDistance = 350
$splitDeploy.SplitterWidth = 6
$splitDeploy.BackColor = $clrSepLine
$splitDeploy.Panel1.BackColor = $clrPanelBg
$splitDeploy.Panel2.BackColor = $clrPanelBg
$splitDeploy.Panel1MinSize = 100
$splitDeploy.Panel2MinSize = 80
$tabDeploy.Controls.Add($splitDeploy)

$gridDeploy = New-ThemedGrid

$colDName    = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colDName.HeaderText = "Name";         $colDName.DataPropertyName = "DeploymentName";  $colDName.Width = 220
$colDType    = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colDType.HeaderText = "Type";         $colDType.DataPropertyName = "DeploymentType";  $colDType.Width = 110
$colDColl    = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colDColl.HeaderText = "Collection";   $colDColl.DataPropertyName = "CollectionName";  $colDColl.Width = 160
$colDPurpose = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colDPurpose.HeaderText = "Purpose";   $colDPurpose.DataPropertyName = "Purpose";      $colDPurpose.Width = 70
$colDTarget  = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colDTarget.HeaderText = "Targeted";   $colDTarget.DataPropertyName = "NumberTargeted"; $colDTarget.Width = 70
$colDSuccess = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colDSuccess.HeaderText = "Success";   $colDSuccess.DataPropertyName = "NumberSuccess"; $colDSuccess.Width = 65
$colDFailed  = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colDFailed.HeaderText = "Failed";     $colDFailed.DataPropertyName = "NumberErrors";  $colDFailed.Width = 55
$colDInProg  = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colDInProg.HeaderText = "In Progress"; $colDInProg.DataPropertyName = "NumberInProgress"; $colDInProg.Width = 80
$colDUnknown = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colDUnknown.HeaderText = "Unknown";   $colDUnknown.DataPropertyName = "NumberUnknown"; $colDUnknown.Width = 65
$colDPct     = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colDPct.HeaderText = "% Compliant";   $colDPct.DataPropertyName = "PercentCompliant"; $colDPct.AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::Fill

$gridDeploy.Columns.AddRange([System.Windows.Forms.DataGridViewColumn[]]@($colDName, $colDType, $colDColl, $colDPurpose, $colDTarget, $colDSuccess, $colDFailed, $colDInProg, $colDUnknown, $colDPct))
$splitDeploy.Panel1.Controls.Add($gridDeploy)

$dtDeploy = New-Object System.Data.DataTable
[void]$dtDeploy.Columns.Add("DeploymentId", [string])
[void]$dtDeploy.Columns.Add("DeploymentName", [string])
[void]$dtDeploy.Columns.Add("DeploymentType", [string])
[void]$dtDeploy.Columns.Add("CollectionName", [string])
[void]$dtDeploy.Columns.Add("Purpose", [string])
[void]$dtDeploy.Columns.Add("NumberTargeted", [int])
[void]$dtDeploy.Columns.Add("NumberSuccess", [int])
[void]$dtDeploy.Columns.Add("NumberErrors", [int])
[void]$dtDeploy.Columns.Add("NumberInProgress", [int])
[void]$dtDeploy.Columns.Add("NumberUnknown", [int])
[void]$dtDeploy.Columns.Add("PercentCompliant", [double])
$gridDeploy.DataSource = $dtDeploy

# Deployment detail panel
$txtDeployInfo = New-Object System.Windows.Forms.RichTextBox
$txtDeployInfo.Dock = [System.Windows.Forms.DockStyle]::Fill
$txtDeployInfo.ReadOnly = $true
$txtDeployInfo.BackColor = $clrDetailBg
$txtDeployInfo.ForeColor = $clrText
$txtDeployInfo.Font = New-Object System.Drawing.Font("Consolas", 9.5)
$txtDeployInfo.BorderStyle = [System.Windows.Forms.BorderStyle]::None
$splitDeploy.Panel2.Controls.Add($txtDeployInfo)

# Row color coding for deployments
$gridDeploy.Add_RowPrePaint({
    param($s, $e)
    try {
        if ($e.RowIndex -ge 0 -and $e.RowIndex -lt $dtDeploy.DefaultView.Count) {
            $rowView = $dtDeploy.DefaultView[$e.RowIndex]
            $failed = [int]$rowView["NumberErrors"]
            $inProg = [int]$rowView["NumberInProgress"]
            if ($failed -gt 0) {
                $s.Rows[$e.RowIndex].DefaultCellStyle.ForeColor = $clrErrText
            }
            elseif ($inProg -gt 0) {
                $s.Rows[$e.RowIndex].DefaultCellStyle.ForeColor = $clrWarnText
            }
            else {
                $s.Rows[$e.RowIndex].DefaultCellStyle.ForeColor = $clrGridText
            }
        }
    } catch {}
})

# ===================== TAB 1: Content =====================

$tabContent = New-Object System.Windows.Forms.TabPage
$tabContent.Text = "Content"
$tabContent.BackColor = $clrFormBg
$tabMain.TabPages.Add($tabContent)

$gridContent = New-ThemedGrid

$colCName    = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colCName.HeaderText = "Content Name";  $colCName.DataPropertyName = "ContentName";     $colCName.Width = 280
$colCType    = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colCType.HeaderText = "Type";          $colCType.DataPropertyName = "ContentType";     $colCType.Width = 120
$colCPkgId   = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colCPkgId.HeaderText = "Package ID";   $colCPkgId.DataPropertyName = "PackageID";      $colCPkgId.Width = 90
$colCTotalDP = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colCTotalDP.HeaderText = "Total DPs";   $colCTotalDP.DataPropertyName = "TotalDPs";     $colCTotalDP.Width = 70
$colCInstall = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colCInstall.HeaderText = "Installed";   $colCInstall.DataPropertyName = "InstalledCount"; $colCInstall.Width = 70
$colCFailed  = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colCFailed.HeaderText = "Failed";       $colCFailed.DataPropertyName = "FailedCount";   $colCFailed.Width = 60
$colCInProg  = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colCInProg.HeaderText = "In Progress";  $colCInProg.DataPropertyName = "InProgressCount"; $colCInProg.AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::Fill

$gridContent.Columns.AddRange([System.Windows.Forms.DataGridViewColumn[]]@($colCName, $colCType, $colCPkgId, $colCTotalDP, $colCInstall, $colCFailed, $colCInProg))
$tabContent.Controls.Add($gridContent)

$dtContent = New-Object System.Data.DataTable
[void]$dtContent.Columns.Add("ContentName", [string])
[void]$dtContent.Columns.Add("ContentType", [string])
[void]$dtContent.Columns.Add("PackageID", [string])
[void]$dtContent.Columns.Add("TotalDPs", [int])
[void]$dtContent.Columns.Add("InstalledCount", [int])
[void]$dtContent.Columns.Add("FailedCount", [int])
[void]$dtContent.Columns.Add("InProgressCount", [int])
$gridContent.DataSource = $dtContent

$gridContent.Add_RowPrePaint({
    param($s, $e)
    try {
        if ($e.RowIndex -ge 0 -and $e.RowIndex -lt $dtContent.DefaultView.Count) {
            $rowView = $dtContent.DefaultView[$e.RowIndex]
            $failed = [int]$rowView["FailedCount"]
            $inProg = [int]$rowView["InProgressCount"]
            if ($failed -gt 0) { $s.Rows[$e.RowIndex].DefaultCellStyle.ForeColor = $clrErrText }
            elseif ($inProg -gt 0) { $s.Rows[$e.RowIndex].DefaultCellStyle.ForeColor = $clrWarnText }
            else { $s.Rows[$e.RowIndex].DefaultCellStyle.ForeColor = $clrGridText }
        }
    } catch {}
})

# ===================== TAB 2: Distribution Points =====================

$tabDPs = New-Object System.Windows.Forms.TabPage
$tabDPs.Text = "Distribution Points"
$tabDPs.BackColor = $clrFormBg
$tabMain.TabPages.Add($tabDPs)

$splitDPs = New-Object System.Windows.Forms.SplitContainer
$splitDPs.Dock = [System.Windows.Forms.DockStyle]::Fill
$splitDPs.Orientation = [System.Windows.Forms.Orientation]::Horizontal
$splitDPs.SplitterDistance = 350
$splitDPs.SplitterWidth = 6
$splitDPs.BackColor = $clrSepLine
$splitDPs.Panel1.BackColor = $clrPanelBg
$splitDPs.Panel2.BackColor = $clrPanelBg
$splitDPs.Panel1MinSize = 100
$splitDPs.Panel2MinSize = 80
$tabDPs.Controls.Add($splitDPs)

$gridDPs = New-ThemedGrid

$colPName    = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colPName.HeaderText = "DP Name";      $colPName.DataPropertyName = "DPName";         $colPName.Width = 250
$colPSite    = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colPSite.HeaderText = "Site";         $colPSite.DataPropertyName = "SiteCode";       $colPSite.Width = 60
$colPStatus  = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colPStatus.HeaderText = "Status";     $colPStatus.DataPropertyName = "Status";       $colPStatus.Width = 80
$colPTotal   = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colPTotal.HeaderText = "Total Content"; $colPTotal.DataPropertyName = "TotalContent"; $colPTotal.Width = 100
$colPFailed  = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colPFailed.HeaderText = "Failed";     $colPFailed.DataPropertyName = "FailedContent"; $colPFailed.Width = 60
$colPPull    = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colPPull.HeaderText = "Pull DP";      $colPPull.DataPropertyName = "IsPullDP";       $colPPull.AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::Fill

$gridDPs.Columns.AddRange([System.Windows.Forms.DataGridViewColumn[]]@($colPName, $colPSite, $colPStatus, $colPTotal, $colPFailed, $colPPull))
$splitDPs.Panel1.Controls.Add($gridDPs)

$dtDPs = New-Object System.Data.DataTable
[void]$dtDPs.Columns.Add("DPName", [string])
[void]$dtDPs.Columns.Add("SiteCode", [string])
[void]$dtDPs.Columns.Add("Status", [string])
[void]$dtDPs.Columns.Add("TotalContent", [int])
[void]$dtDPs.Columns.Add("FailedContent", [int])
[void]$dtDPs.Columns.Add("IsPullDP", [string])
$gridDPs.DataSource = $dtDPs

$txtDPInfo = New-Object System.Windows.Forms.RichTextBox
$txtDPInfo.Dock = [System.Windows.Forms.DockStyle]::Fill
$txtDPInfo.ReadOnly = $true
$txtDPInfo.BackColor = $clrDetailBg
$txtDPInfo.ForeColor = $clrText
$txtDPInfo.Font = New-Object System.Drawing.Font("Consolas", 9.5)
$txtDPInfo.BorderStyle = [System.Windows.Forms.BorderStyle]::None
$splitDPs.Panel2.Controls.Add($txtDPInfo)

$gridDPs.Add_RowPrePaint({
    param($s, $e)
    try {
        if ($e.RowIndex -ge 0 -and $e.RowIndex -lt $dtDPs.DefaultView.Count) {
            $rowView = $dtDPs.DefaultView[$e.RowIndex]
            $st = [string]$rowView["Status"]
            if ($st -eq 'Critical') { $s.Rows[$e.RowIndex].DefaultCellStyle.ForeColor = $clrErrText }
            elseif ($st -eq 'Warning') { $s.Rows[$e.RowIndex].DefaultCellStyle.ForeColor = $clrWarnText }
            else { $s.Rows[$e.RowIndex].DefaultCellStyle.ForeColor = $clrGridText }
        }
    } catch {}
})

# ===================== TAB 3: Client Health =====================

$tabClients = New-Object System.Windows.Forms.TabPage
$tabClients.Text = "Client Health"
$tabClients.BackColor = $clrFormBg
$tabMain.TabPages.Add($tabClients)

$splitClients = New-Object System.Windows.Forms.SplitContainer
$splitClients.Dock = [System.Windows.Forms.DockStyle]::Fill
$splitClients.Orientation = [System.Windows.Forms.Orientation]::Horizontal
$splitClients.SplitterDistance = 350
$splitClients.SplitterWidth = 6
$splitClients.BackColor = $clrSepLine
$splitClients.Panel1.BackColor = $clrPanelBg
$splitClients.Panel2.BackColor = $clrPanelBg
$splitClients.Panel1MinSize = 100
$splitClients.Panel2MinSize = 80
$tabClients.Controls.Add($splitClients)

$gridClients = New-ThemedGrid

$colClName   = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colClName.HeaderText = "Device";       $colClName.DataPropertyName = "DeviceName";       $colClName.Width = 180
$colClHealth = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colClHealth.HeaderText = "Health";     $colClHealth.DataPropertyName = "HealthState";    $colClHealth.Width = 80
$colClActive = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colClActive.HeaderText = "Active";     $colClActive.DataPropertyName = "ActiveStatus";   $colClActive.Width = 70
$colClOnline = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colClOnline.HeaderText = "Last Online"; $colClOnline.DataPropertyName = "LastOnlineTime"; $colClOnline.Width = 130
$colClDDR    = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colClDDR.HeaderText = "Last DDR";      $colClDDR.DataPropertyName = "LastDDR";           $colClDDR.Width = 130
$colClPolicy = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colClPolicy.HeaderText = "Last Policy"; $colClPolicy.DataPropertyName = "LastPolicyRequest"; $colClPolicy.Width = 130
$colClHW     = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colClHW.HeaderText = "Last HW Inv";    $colClHW.DataPropertyName = "LastHWInventory";    $colClHW.Width = 130
$colClVer    = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colClVer.HeaderText = "Version";       $colClVer.DataPropertyName = "ClientVersion";     $colClVer.AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::Fill

$gridClients.Columns.AddRange([System.Windows.Forms.DataGridViewColumn[]]@($colClName, $colClHealth, $colClActive, $colClOnline, $colClDDR, $colClPolicy, $colClHW, $colClVer))
$splitClients.Panel1.Controls.Add($gridClients)

$dtClients = New-Object System.Data.DataTable
[void]$dtClients.Columns.Add("DeviceName", [string])
[void]$dtClients.Columns.Add("HealthState", [string])
[void]$dtClients.Columns.Add("ActiveStatus", [string])
[void]$dtClients.Columns.Add("LastOnlineTime", [string])
[void]$dtClients.Columns.Add("LastDDR", [string])
[void]$dtClients.Columns.Add("LastPolicyRequest", [string])
[void]$dtClients.Columns.Add("LastHWInventory", [string])
[void]$dtClients.Columns.Add("ClientVersion", [string])
$gridClients.DataSource = $dtClients

$txtClientInfo = New-Object System.Windows.Forms.RichTextBox
$txtClientInfo.Dock = [System.Windows.Forms.DockStyle]::Fill
$txtClientInfo.ReadOnly = $true
$txtClientInfo.BackColor = $clrDetailBg
$txtClientInfo.ForeColor = $clrText
$txtClientInfo.Font = New-Object System.Drawing.Font("Consolas", 9.5)
$txtClientInfo.BorderStyle = [System.Windows.Forms.BorderStyle]::None
$splitClients.Panel2.Controls.Add($txtClientInfo)

$gridClients.Add_RowPrePaint({
    param($s, $e)
    try {
        if ($e.RowIndex -ge 0 -and $e.RowIndex -lt $dtClients.DefaultView.Count) {
            $rowView = $dtClients.DefaultView[$e.RowIndex]
            $health = [string]$rowView["HealthState"]
            if ($health -eq 'Unhealthy') { $s.Rows[$e.RowIndex].DefaultCellStyle.ForeColor = $clrErrText }
            elseif ($health -eq 'Unknown') { $s.Rows[$e.RowIndex].DefaultCellStyle.ForeColor = $clrWarnText }
            else { $s.Rows[$e.RowIndex].DefaultCellStyle.ForeColor = $clrGridText }
        }
    } catch {}
})

# ===================== TAB 4: Inactive Devices =====================

$tabInactive = New-Object System.Windows.Forms.TabPage
$tabInactive.Text = "Inactive Devices"
$tabInactive.BackColor = $clrFormBg
$tabMain.TabPages.Add($tabInactive)

$gridInactive = New-ThemedGrid

$colIDName  = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colIDName.HeaderText = "Device";        $colIDName.DataPropertyName = "DeviceName";      $colIDName.Width = 200
$colIDOnline = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colIDOnline.HeaderText = "Last Online"; $colIDOnline.DataPropertyName = "LastOnlineTime"; $colIDOnline.Width = 140
$colIDDDR   = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colIDDDR.HeaderText = "Last DDR";       $colIDDDR.DataPropertyName = "LastDDR";          $colIDDDR.Width = 140
$colIDDays  = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colIDDays.HeaderText = "Days Inactive"; $colIDDays.DataPropertyName = "DaysSinceContact"; $colIDDays.Width = 100
$colIDOS    = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colIDOS.HeaderText = "OS";              $colIDOS.DataPropertyName = "OperatingSystem";   $colIDOS.Width = 200
$colIDVer   = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colIDVer.HeaderText = "Version";        $colIDVer.DataPropertyName = "ClientVersion";    $colIDVer.AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::Fill

$gridInactive.Columns.AddRange([System.Windows.Forms.DataGridViewColumn[]]@($colIDName, $colIDOnline, $colIDDDR, $colIDDays, $colIDOS, $colIDVer))
$tabInactive.Controls.Add($gridInactive)

$dtInactive = New-Object System.Data.DataTable
[void]$dtInactive.Columns.Add("DeviceName", [string])
[void]$dtInactive.Columns.Add("LastOnlineTime", [string])
[void]$dtInactive.Columns.Add("LastDDR", [string])
[void]$dtInactive.Columns.Add("DaysSinceContact", [int])
[void]$dtInactive.Columns.Add("OperatingSystem", [string])
[void]$dtInactive.Columns.Add("ClientVersion", [string])
$gridInactive.DataSource = $dtInactive

# ===================== TAB 5: Site Health =====================

$tabSite = New-Object System.Windows.Forms.TabPage
$tabSite.Text = "Site Health"
$tabSite.BackColor = $clrFormBg
$tabMain.TabPages.Add($tabSite)

$splitSite = New-Object System.Windows.Forms.SplitContainer
$splitSite.Dock = [System.Windows.Forms.DockStyle]::Fill
$splitSite.Orientation = [System.Windows.Forms.Orientation]::Horizontal
$splitSite.SplitterDistance = 350
$splitSite.SplitterWidth = 6
$splitSite.BackColor = $clrSepLine
$splitSite.Panel1.BackColor = $clrPanelBg
$splitSite.Panel2.BackColor = $clrPanelBg
$splitSite.Panel1MinSize = 100
$splitSite.Panel2MinSize = 80
$tabSite.Controls.Add($splitSite)

$gridSite = New-ThemedGrid

$colSName    = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colSName.HeaderText = "Name";          $colSName.DataPropertyName = "Name";            $colSName.Width = 250
$colSType    = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colSType.HeaderText = "Type";          $colSType.DataPropertyName = "ItemType";        $colSType.Width = 100
$colSMachine = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colSMachine.HeaderText = "Server";     $colSMachine.DataPropertyName = "MachineName";  $colSMachine.Width = 180
$colSStatus  = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colSStatus.HeaderText = "Status";      $colSStatus.DataPropertyName = "Status";        $colSStatus.Width = 80
$colSState   = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colSState.HeaderText = "State";        $colSState.DataPropertyName = "State";          $colSState.Width = 80
$colSStarted = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colSStarted.HeaderText = "Last Started"; $colSStarted.DataPropertyName = "LastStarted"; $colSStarted.AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::Fill

$gridSite.Columns.AddRange([System.Windows.Forms.DataGridViewColumn[]]@($colSName, $colSType, $colSMachine, $colSStatus, $colSState, $colSStarted))
$splitSite.Panel1.Controls.Add($gridSite)

$dtSite = New-Object System.Data.DataTable
[void]$dtSite.Columns.Add("Name", [string])
[void]$dtSite.Columns.Add("ItemType", [string])
[void]$dtSite.Columns.Add("MachineName", [string])
[void]$dtSite.Columns.Add("Status", [string])
[void]$dtSite.Columns.Add("State", [string])
[void]$dtSite.Columns.Add("LastStarted", [string])
$gridSite.DataSource = $dtSite

$txtSiteInfo = New-Object System.Windows.Forms.RichTextBox
$txtSiteInfo.Dock = [System.Windows.Forms.DockStyle]::Fill
$txtSiteInfo.ReadOnly = $true
$txtSiteInfo.BackColor = $clrDetailBg
$txtSiteInfo.ForeColor = $clrText
$txtSiteInfo.Font = New-Object System.Drawing.Font("Consolas", 9.5)
$txtSiteInfo.BorderStyle = [System.Windows.Forms.BorderStyle]::None
$splitSite.Panel2.Controls.Add($txtSiteInfo)

$gridSite.Add_RowPrePaint({
    param($s, $e)
    try {
        if ($e.RowIndex -ge 0 -and $e.RowIndex -lt $dtSite.DefaultView.Count) {
            $rowView = $dtSite.DefaultView[$e.RowIndex]
            $st = [string]$rowView["Status"]
            if ($st -eq 'Critical') { $s.Rows[$e.RowIndex].DefaultCellStyle.ForeColor = $clrErrText }
            elseif ($st -eq 'Warning') { $s.Rows[$e.RowIndex].DefaultCellStyle.ForeColor = $clrWarnText }
            else { $s.Rows[$e.RowIndex].DefaultCellStyle.ForeColor = $clrGridText }
        }
    } catch {}
})

# ---------------------------------------------------------------------------
# Finalize dock Z-order
# ---------------------------------------------------------------------------

$form.Controls.Add($menuStrip)
$menuStrip.SendToBack()

$pnlSep3.BringToFront()
$pnlFilter.BringToFront()
$pnlSep2.BringToFront()
$pnlCards.BringToFront()
$pnlSep1.BringToFront()
$pnlConnBar.BringToFront()
$pnlHeader.BringToFront()

$tabMain.BringToFront()

# ---------------------------------------------------------------------------
# Module-scoped data (populated by Refresh All)
# ---------------------------------------------------------------------------

$script:DeploymentData  = @()
$script:ContentData     = @()
$script:DPData          = @()
$script:ClientData      = @()
$script:InactiveData    = @()
$script:ComponentData   = @()
$script:SystemData      = @()
$script:SQLConnected    = $false

# ---------------------------------------------------------------------------
# Status bar update helper
# ---------------------------------------------------------------------------

function Update-StatusBar {
    $parts = @()

    if (Test-CMConnection) {
        $parts += "Connected to $($script:Prefs.SiteCode)"
    } else {
        $parts += "Disconnected"
    }

    if ($script:AutoRefreshPaused) {
        $parts += "Auto-refresh: Paused"
    }
    elseif ($script:NextRefreshTime) {
        $remaining = ($script:NextRefreshTime - (Get-Date))
        if ($remaining.TotalSeconds -gt 0) {
            $parts += "Next refresh: {0}:{1:D2}" -f [int]$remaining.TotalMinutes, $remaining.Seconds
        }
    }

    if ($script:LastRefreshTime) {
        $parts += "Last refresh: $($script:LastRefreshTime.ToString('HH:mm:ss'))"
    }

    $statusLabel.Text = $parts -join " | "
}

$script:LastRefreshTime = $null

# ---------------------------------------------------------------------------
# Core operations
# ---------------------------------------------------------------------------

function Invoke-RefreshAll {
    if (-not $script:Prefs.SiteCode -or -not $script:Prefs.SMSProvider) {
        [System.Windows.Forms.MessageBox]::Show(
            "Site Code and SMS Provider must be configured in File > Preferences.",
            "Configuration Required", "OK", "Warning") | Out-Null
        return
    }

    # Pause auto-refresh during manual refresh
    $script:RefreshTimer.Stop()
    $script:CountdownTimer.Stop()

    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    $btnRefreshAll.Enabled = $false

    try {
        # Step 1: Connect if needed
        if (-not (Test-CMConnection)) {
            Add-LogLine -TextBox $txtLog -Message "Connecting to site $($script:Prefs.SiteCode)..."
            [System.Windows.Forms.Application]::DoEvents()

            $connected = Connect-CMSite -SiteCode $script:Prefs.SiteCode -SMSProvider $script:Prefs.SMSProvider
            if (-not $connected) {
                Add-LogLine -TextBox $txtLog -Message "Connection failed. Check preferences and try again."
                $statusLabel.Text = "Connection failed."
                return
            }
            $lblConnStatus.Text = "Connected"
            $lblConnStatus.ForeColor = $clrOkText
        }

        # Step 2: Test SQL if configured
        if ($script:Prefs.SQLServer -and -not $script:SQLConnected) {
            Add-LogLine -TextBox $txtLog -Message "Testing SQL connection to $($script:Prefs.SQLServer)..."
            [System.Windows.Forms.Application]::DoEvents()
            $script:SQLConnected = Test-SQLConnection -SQLServer $script:Prefs.SQLServer -SiteCode $script:Prefs.SiteCode
            if ($script:SQLConnected) {
                Add-LogLine -TextBox $txtLog -Message "SQL connection verified"
            } else {
                Add-LogLine -TextBox $txtLog -Message "SQL connection failed - client health and inactive device tabs will be empty"
            }
        }

        # Step 3: Deployment health
        Add-LogLine -TextBox $txtLog -Message "Querying deployment health..."
        [System.Windows.Forms.Application]::DoEvents()
        $script:DeploymentData = @(Get-DeploymentHealth)
        $deployCounts = Get-DeploymentHealthCounts -DeploymentData $script:DeploymentData
        Add-LogLine -TextBox $txtLog -Message "Found $($script:DeploymentData.Count) deployments ($($deployCounts.FailedDeployments) with errors)"
        [System.Windows.Forms.Application]::DoEvents()

        $dtDeploy.Clear()
        foreach ($d in $script:DeploymentData) {
            [void]$dtDeploy.Rows.Add(
                $d.DeploymentId, $d.DeploymentName, $d.DeploymentType,
                $d.CollectionName, $d.Purpose, $d.NumberTargeted,
                $d.NumberSuccess, $d.NumberErrors, $d.NumberInProgress,
                $d.NumberUnknown, $d.PercentCompliant
            )
        }

        $deploySeverity = if ($deployCounts.FailedDeployments -gt 0) { 'critical' } elseif ($deployCounts.OverallCompliance -lt 95) { 'warn' } else { 'ok' }
        Update-Card -Card $cardDeploy -ValueText "$($deployCounts.FailedDeployments) failed" -Severity $deploySeverity
        [System.Windows.Forms.Application]::DoEvents()

        # Step 4: Content distribution health
        Add-LogLine -TextBox $txtLog -Message "Querying content distribution health..."
        [System.Windows.Forms.Application]::DoEvents()
        $script:ContentData = @(Get-ContentDistributionHealth -SMSProvider $script:Prefs.SMSProvider -SiteCode $script:Prefs.SiteCode)
        $contentCounts = Get-ContentHealthCounts -ContentData $script:ContentData
        Add-LogLine -TextBox $txtLog -Message "Content: $($contentCounts.TotalContentWithIssues) items with issues"
        [System.Windows.Forms.Application]::DoEvents()

        $dtContent.Clear()
        foreach ($c in $script:ContentData) {
            [void]$dtContent.Rows.Add($c.PackageID, '', $c.PackageID, $c.TotalDPs, $c.InstalledCount, $c.FailedCount, $c.InProgressCount)
        }

        $contentSeverity = if ($contentCounts.TotalFailedPairs -gt 0) { 'critical' } elseif ($contentCounts.TotalContentWithIssues -gt 0) { 'warn' } else { 'ok' }
        Update-Card -Card $cardContent -ValueText "$($contentCounts.TotalFailedPairs) failed" -Severity $contentSeverity
        [System.Windows.Forms.Application]::DoEvents()

        # Step 5: DP health
        Add-LogLine -TextBox $txtLog -Message "Querying distribution point health..."
        [System.Windows.Forms.Application]::DoEvents()
        $script:DPData = @(Get-DPHealth -SMSProvider $script:Prefs.SMSProvider -SiteCode $script:Prefs.SiteCode)
        $dpCounts = Get-DPHealthCounts -DPData $script:DPData
        Add-LogLine -TextBox $txtLog -Message "DPs: $($dpCounts.TotalDPs) total, $($dpCounts.OfflineCount) offline, $($dpCounts.DegradedCount) degraded"
        [System.Windows.Forms.Application]::DoEvents()

        $dtDPs.Clear()
        foreach ($dp in $script:DPData) {
            [void]$dtDPs.Rows.Add($dp.DPName, $dp.SiteCode, $dp.Status, $dp.TotalContent, $dp.FailedContent, $dp.IsPullDP)
        }

        $dpSeverity = if ($dpCounts.OfflineCount -gt 0) { 'critical' } elseif ($dpCounts.DegradedCount -gt 0) { 'warn' } else { 'ok' }
        Update-Card -Card $cardDPs -ValueText "$($dpCounts.OfflineCount) offline" -Severity $dpSeverity
        [System.Windows.Forms.Application]::DoEvents()

        # Step 6: Client health (SQL)
        if ($script:SQLConnected) {
            Add-LogLine -TextBox $txtLog -Message "Querying client health (SQL)..."
            [System.Windows.Forms.Application]::DoEvents()
            $script:ClientData = @(Get-ClientHealthSummary -SQLServer $script:Prefs.SQLServer -SiteCode $script:Prefs.SiteCode)
            $clientCounts = Get-ClientHealthCounts -ClientData $script:ClientData
            Add-LogLine -TextBox $txtLog -Message "Clients: $($clientCounts.HealthyCount) healthy, $($clientCounts.UnhealthyCount) unhealthy, $($clientCounts.InactiveCount) inactive"
            [System.Windows.Forms.Application]::DoEvents()

            $dtClients.Clear()
            foreach ($cl in $script:ClientData) {
                [void]$dtClients.Rows.Add(
                    $cl.DeviceName, $cl.HealthState, $cl.ActiveStatus,
                    $(if ($cl.LastOnlineTime) { $cl.LastOnlineTime.ToString('yyyy-MM-dd HH:mm') } else { '' }),
                    $(if ($cl.LastDDR) { $cl.LastDDR.ToString('yyyy-MM-dd HH:mm') } else { '' }),
                    $(if ($cl.LastPolicyRequest) { $cl.LastPolicyRequest.ToString('yyyy-MM-dd HH:mm') } else { '' }),
                    $(if ($cl.LastHWInventory) { $cl.LastHWInventory.ToString('yyyy-MM-dd HH:mm') } else { '' }),
                    $cl.ClientVersion
                )
            }

            $clientSeverity = if ($clientCounts.UnhealthyCount -gt 0) { 'critical' } elseif ($clientCounts.InactiveCount -gt 0) { 'warn' } else { 'ok' }
            Update-Card -Card $cardClients -ValueText "$($clientCounts.UnhealthyCount) unhealthy" -Severity $clientSeverity
        } else {
            Update-Card -Card $cardClients -ValueText "No SQL" -Severity 'default'
        }
        [System.Windows.Forms.Application]::DoEvents()

        # Step 7: Inactive devices (SQL)
        if ($script:SQLConnected) {
            Add-LogLine -TextBox $txtLog -Message "Querying inactive devices (threshold: $($script:Prefs.InactiveThresholdDays) days)..."
            [System.Windows.Forms.Application]::DoEvents()
            $script:InactiveData = @(Get-InactiveDevices -SQLServer $script:Prefs.SQLServer -SiteCode $script:Prefs.SiteCode -ThresholdDays $script:Prefs.InactiveThresholdDays)
            $inactiveCounts = Get-InactiveDeviceCounts -DeviceData $script:InactiveData
            Add-LogLine -TextBox $txtLog -Message "Inactive devices: $($inactiveCounts.InactiveCount)"
            [System.Windows.Forms.Application]::DoEvents()

            $dtInactive.Clear()
            foreach ($id in $script:InactiveData) {
                [void]$dtInactive.Rows.Add(
                    $id.DeviceName,
                    $(if ($id.LastOnlineTime) { $id.LastOnlineTime.ToString('yyyy-MM-dd HH:mm') } else { '' }),
                    $(if ($id.LastDDR) { $id.LastDDR.ToString('yyyy-MM-dd HH:mm') } else { '' }),
                    $id.DaysSinceContact,
                    $id.OperatingSystem,
                    $id.ClientVersion
                )
            }

            $inactiveSeverity = if ($inactiveCounts.InactiveCount -gt 50) { 'critical' } elseif ($inactiveCounts.InactiveCount -gt 0) { 'warn' } else { 'ok' }
            Update-Card -Card $cardDevices -ValueText "$($inactiveCounts.InactiveCount) stale" -Severity $inactiveSeverity
        } else {
            Update-Card -Card $cardDevices -ValueText "No SQL" -Severity 'default'
        }
        [System.Windows.Forms.Application]::DoEvents()

        # Step 8: Site health (WMI)
        Add-LogLine -TextBox $txtLog -Message "Querying site component health..."
        [System.Windows.Forms.Application]::DoEvents()
        $script:ComponentData = @(Get-SiteComponentHealth -SMSProvider $script:Prefs.SMSProvider -SiteCode $script:Prefs.SiteCode)
        $script:SystemData = @(Get-SiteSystemHealth -SMSProvider $script:Prefs.SMSProvider -SiteCode $script:Prefs.SiteCode)
        $siteCounts = Get-SiteHealthCounts -ComponentData $script:ComponentData -SystemData $script:SystemData
        Add-LogLine -TextBox $txtLog -Message "Site health: $($siteCounts.OKCount) OK, $($siteCounts.WarningCount) warning, $($siteCounts.CriticalCount) critical"
        [System.Windows.Forms.Application]::DoEvents()

        $dtSite.Clear()
        foreach ($comp in $script:ComponentData) {
            [void]$dtSite.Rows.Add(
                $comp.ComponentName, $comp.ItemType, $comp.MachineName,
                $comp.Status, $comp.State,
                $(if ($comp.LastStarted) { try { $comp.LastStarted.ToString('yyyy-MM-dd HH:mm') } catch { '' } } else { '' })
            )
        }
        foreach ($sys in $script:SystemData) {
            [void]$dtSite.Rows.Add(
                $sys.RoleName, $sys.ItemType, $sys.ServerName,
                $sys.Status, '', ''
            )
        }

        $siteSeverity = if ($siteCounts.CriticalCount -gt 0) { 'critical' } elseif ($siteCounts.WarningCount -gt 0) { 'warn' } else { 'ok' }
        Update-Card -Card $cardSite -ValueText "$($siteCounts.CriticalCount) critical" -Severity $siteSeverity
        [System.Windows.Forms.Application]::DoEvents()

        # Done
        $script:LastRefreshTime = Get-Date
        Add-LogLine -TextBox $txtLog -Message "Refresh complete."
    }
    catch {
        Add-LogLine -TextBox $txtLog -Message "ERROR: $_"
    }
    finally {
        $form.Cursor = [System.Windows.Forms.Cursors]::Default
        $btnRefreshAll.Enabled = $true

        # Restart auto-refresh timer
        if (-not $script:AutoRefreshPaused) {
            $script:NextRefreshTime = (Get-Date).AddMilliseconds($script:RefreshTimer.Interval)
            $script:RefreshTimer.Start()
            $script:CountdownTimer.Start()
        }
        Update-StatusBar
    }
}

# ---------------------------------------------------------------------------
# Drill-down handlers
# ---------------------------------------------------------------------------

$gridDeploy.Add_SelectionChanged({
    if ($gridDeploy.SelectedRows.Count -eq 0) { $txtDeployInfo.Text = ''; return }
    $rowIdx = $gridDeploy.SelectedRows[0].Index
    if ($rowIdx -lt 0 -or $rowIdx -ge $dtDeploy.DefaultView.Count) { return }

    $row = $dtDeploy.DefaultView[$rowIdx]
    $lines = @()
    $lines += "Deployment:   $($row['DeploymentName'])"
    $lines += "Type:         $($row['DeploymentType'])"
    $lines += "Collection:   $($row['CollectionName'])"
    $lines += "Purpose:      $($row['Purpose'])"
    $lines += "Targeted:     $($row['NumberTargeted'])"
    $lines += "Success:      $($row['NumberSuccess'])"
    $lines += "Failed:       $($row['NumberErrors'])"
    $lines += "In Progress:  $($row['NumberInProgress'])"
    $lines += "Unknown:      $($row['NumberUnknown'])"
    $lines += "% Compliant:  $($row['PercentCompliant'])%"
    $txtDeployInfo.Text = $lines -join "`r`n"
})

$gridDPs.Add_SelectionChanged({
    if ($gridDPs.SelectedRows.Count -eq 0) { $txtDPInfo.Text = ''; return }
    $rowIdx = $gridDPs.SelectedRows[0].Index
    if ($rowIdx -lt 0 -or $rowIdx -ge $dtDPs.DefaultView.Count) { return }

    $row = $dtDPs.DefaultView[$rowIdx]
    $lines = @()
    $lines += "DP Name:      $($row['DPName'])"
    $lines += "Site Code:    $($row['SiteCode'])"
    $lines += "Status:       $($row['Status'])"
    $lines += "Pull DP:      $($row['IsPullDP'])"
    $lines += "Total Content: $($row['TotalContent'])"
    $lines += "Failed:       $($row['FailedContent'])"
    $txtDPInfo.Text = $lines -join "`r`n"
})

$gridClients.Add_SelectionChanged({
    if ($gridClients.SelectedRows.Count -eq 0) { $txtClientInfo.Text = ''; return }
    $rowIdx = $gridClients.SelectedRows[0].Index
    if ($rowIdx -lt 0 -or $rowIdx -ge $dtClients.DefaultView.Count) { return }

    $row = $dtClients.DefaultView[$rowIdx]
    $lines = @()
    $lines += "Device:       $($row['DeviceName'])"
    $lines += "Health:       $($row['HealthState'])"
    $lines += "Active:       $($row['ActiveStatus'])"
    $lines += "Last Online:  $($row['LastOnlineTime'])"
    $lines += "Last DDR:     $($row['LastDDR'])"
    $lines += "Last Policy:  $($row['LastPolicyRequest'])"
    $lines += "Last HW Inv:  $($row['LastHWInventory'])"
    $lines += "Version:      $($row['ClientVersion'])"
    $txtClientInfo.Text = $lines -join "`r`n"
})

$gridSite.Add_SelectionChanged({
    if ($gridSite.SelectedRows.Count -eq 0) { $txtSiteInfo.Text = ''; return }
    $rowIdx = $gridSite.SelectedRows[0].Index
    if ($rowIdx -lt 0 -or $rowIdx -ge $dtSite.DefaultView.Count) { return }

    $row = $dtSite.DefaultView[$rowIdx]
    $lines = @()
    $lines += "Name:         $($row['Name'])"
    $lines += "Type:         $($row['ItemType'])"
    $lines += "Server:       $($row['MachineName'])"
    $lines += "Status:       $($row['Status'])"
    $lines += "State:        $($row['State'])"
    $lines += "Last Started: $($row['LastStarted'])"
    $txtSiteInfo.Text = $lines -join "`r`n"
})

# ---------------------------------------------------------------------------
# Filter logic
# ---------------------------------------------------------------------------

function Apply-ActiveFilter {
    $tabIdx = $tabMain.SelectedIndex
    $statusVal = $cboStatus.SelectedItem
    $textVal = $txtFilter.Text.Trim()

    # Determine which DataTable to filter
    $dt = switch ($tabIdx) {
        0 { $dtDeploy }
        1 { $dtContent }
        2 { $dtDPs }
        3 { $dtClients }
        4 { $dtInactive }
        5 { $dtSite }
    }

    if (-not $dt) { return }

    $parts = @()

    # Status filter
    if ($statusVal -and $statusVal -ne 'All') {
        switch ($tabIdx) {
            0 { # Deployments
                switch ($statusVal) {
                    'Failed/Error'       { $parts += "NumberErrors > 0" }
                    'Warning/InProgress'  { $parts += "NumberInProgress > 0" }
                    'OK/Healthy'          { $parts += "NumberErrors = 0 AND NumberInProgress = 0" }
                }
            }
            1 { # Content
                switch ($statusVal) {
                    'Failed/Error'       { $parts += "FailedCount > 0" }
                    'Warning/InProgress'  { $parts += "InProgressCount > 0" }
                    'OK/Healthy'          { $parts += "FailedCount = 0 AND InProgressCount = 0" }
                }
            }
            2 { # DPs
                switch ($statusVal) {
                    'Failed/Error'       { $parts += "Status = 'Critical'" }
                    'Warning/InProgress'  { $parts += "Status = 'Warning'" }
                    'OK/Healthy'          { $parts += "Status = 'OK'" }
                }
            }
            3 { # Client Health
                switch ($statusVal) {
                    'Failed/Error'       { $parts += "HealthState = 'Unhealthy'" }
                    'Warning/InProgress'  { $parts += "ActiveStatus = 'Inactive'" }
                    'OK/Healthy'          { $parts += "HealthState = 'Healthy'" }
                }
            }
            5 { # Site Health
                switch ($statusVal) {
                    'Failed/Error'       { $parts += "Status = 'Critical'" }
                    'Warning/InProgress'  { $parts += "Status = 'Warning'" }
                    'OK/Healthy'          { $parts += "Status = 'OK'" }
                }
            }
        }
    }

    # Text filter
    if ($textVal) {
        $escaped = $textVal.Replace("'", "''")
        $textPart = switch ($tabIdx) {
            0 { "(DeploymentName LIKE '*$escaped*' OR CollectionName LIKE '*$escaped*')" }
            1 { "(ContentName LIKE '*$escaped*' OR PackageID LIKE '*$escaped*')" }
            2 { "DPName LIKE '*$escaped*'" }
            3 { "DeviceName LIKE '*$escaped*'" }
            4 { "DeviceName LIKE '*$escaped*'" }
            5 { "(Name LIKE '*$escaped*' OR MachineName LIKE '*$escaped*')" }
        }
        if ($textPart) { $parts += $textPart }
    }

    $dt.DefaultView.RowFilter = if ($parts.Count -gt 0) { $parts -join ' AND ' } else { '' }
}

$cboStatus.Add_SelectedIndexChanged({ Apply-ActiveFilter })
$txtFilter.Add_TextChanged({ Apply-ActiveFilter })
$tabMain.Add_SelectedIndexChanged({ Apply-ActiveFilter })

# ---------------------------------------------------------------------------
# Export handlers
# ---------------------------------------------------------------------------

$btnExportCsv.Add_Click({
    $tabIdx = $tabMain.SelectedIndex
    $dt = switch ($tabIdx) {
        0 { $dtDeploy }
        1 { $dtContent }
        2 { $dtDPs }
        3 { $dtClients }
        4 { $dtInactive }
        5 { $dtSite }
    }
    $tabName = $tabMain.TabPages[$tabIdx].Text -replace ' ', ''

    $reportsDir = Join-Path $PSScriptRoot "Reports"
    if (-not (Test-Path -LiteralPath $reportsDir)) { New-Item -ItemType Directory -Path $reportsDir -Force | Out-Null }

    $sfd = New-Object System.Windows.Forms.SaveFileDialog
    $sfd.Filter = "CSV Files (*.csv)|*.csv"
    $sfd.FileName = "HealthDash-$tabName-$(Get-Date -Format 'yyyyMMdd-HHmmss').csv"
    $sfd.InitialDirectory = $reportsDir
    if ($sfd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        Export-HealthStatusCsv -DataTable $dt -OutputPath $sfd.FileName
        Add-LogLine -TextBox $txtLog -Message "Exported CSV: $($sfd.FileName)"
    }
})

$btnExportHtml.Add_Click({
    $tabIdx = $tabMain.SelectedIndex
    $dt = switch ($tabIdx) {
        0 { $dtDeploy }
        1 { $dtContent }
        2 { $dtDPs }
        3 { $dtClients }
        4 { $dtInactive }
        5 { $dtSite }
    }
    $tabName = $tabMain.TabPages[$tabIdx].Text

    $reportsDir = Join-Path $PSScriptRoot "Reports"
    if (-not (Test-Path -LiteralPath $reportsDir)) { New-Item -ItemType Directory -Path $reportsDir -Force | Out-Null }

    $sfd = New-Object System.Windows.Forms.SaveFileDialog
    $sfd.Filter = "HTML Files (*.html)|*.html"
    $sfd.FileName = "HealthDash-$($tabName -replace ' ', '')-$(Get-Date -Format 'yyyyMMdd-HHmmss').html"
    $sfd.InitialDirectory = $reportsDir
    if ($sfd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        Export-HealthStatusHtml -DataTable $dt -OutputPath $sfd.FileName -ReportTitle "MECM Health - $tabName"
        Add-LogLine -TextBox $txtLog -Message "Exported HTML: $($sfd.FileName)"
    }
})

# ---------------------------------------------------------------------------
# Button / timer wiring
# ---------------------------------------------------------------------------

$btnRefreshAll.Add_Click({ Invoke-RefreshAll })

$script:RefreshTimer.Add_Tick({ Invoke-RefreshAll })

$script:CountdownTimer.Add_Tick({ Update-StatusBar })

# ---------------------------------------------------------------------------
# Form events
# ---------------------------------------------------------------------------

$form.Add_Shown({
    Restore-WindowState
    Add-LogLine -TextBox $txtLog -Message "MECM Health Dashboard ready. Configure site in File > Preferences, then click Refresh Now."
})

$form.Add_FormClosing({
    $script:RefreshTimer.Stop()
    $script:CountdownTimer.Stop()
    Save-WindowState
    if (Test-CMConnection) {
        Disconnect-CMSite
    }
})

# ---------------------------------------------------------------------------
# Launch
# ---------------------------------------------------------------------------

[void]$form.ShowDialog()
$form.Dispose()
