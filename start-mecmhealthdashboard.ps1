<#
.SYNOPSIS
    MahApps.Metro WPF shell for the MECM Health Dashboard.

.DESCRIPTION
    Replaces the v1.0.x WinForms shell with a brand-aligned WPF UI: sidebar
    navigation across six health views (Deployments, Content, Distribution
    Points, Client Health, Inactive Devices, Site Health) plus an Options
    pane, inline action bar (Refresh All, Pause / Resume, filter, status
    filter, exports, copy summary), log drawer, and status bar. Status
    conveyed via per-row glyph (no red / yellow / green row coloring).
    Site / SMS Provider / SQL configured from the Options sidebar (no File
    menu).

    Requirements:
      - PowerShell 5.1
      - .NET Framework 4.7.2+
      - MahApps.Metro DLLs in .\Lib\
      - MECMHealthDashCommon module under .\Module\
      - ConfigurationManager console (provides Get-CMDeployment et al.)
      - SqlServer module (Invoke-Sqlcmd) for client health / inactive devices

.NOTES
    ScriptName : start-mecmhealthdashboard.ps1
    Version    : 1.0.0
    Updated    : 2026-05-01
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
    $__tx = Join-Path $__txDir ('HealthDash-startup-{0}.txt' -f (Get-Date -Format 'yyyyMMdd-HHmmss'))
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
$__modulePath = Join-Path $PSScriptRoot 'Module\MECMHealthDashCommon.psd1'
if (-not (Test-Path -LiteralPath $__modulePath)) {
    throw "Shared module not found at: $__modulePath"
}
Import-Module -Name $__modulePath -Force -DisableNameChecking
if (-not (Get-Command Initialize-Logging -ErrorAction SilentlyContinue)) {
    throw "MECMHealthDashCommon imported but Initialize-Logging is not exported."
}

# =============================================================================
# Preferences (MECMHealthDash.prefs.json next to the script).
# Closure-safe via $global:.
# =============================================================================
$global:PrefsPath = Join-Path $PSScriptRoot 'MECMHealthDash.prefs.json'

function Get-HdPreferences {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseSingularNouns', '', Justification='Returns the full preferences hashtable by design.')]
    param()
    $defaults = @{
        DarkMode              = $true
        SiteCode              = ''
        SMSProvider           = ''
        SQLServer             = ''
        AutoRefreshMinutes    = 15
        InactiveThresholdDays = 14
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

function Save-HdPreferences {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseSingularNouns', '', Justification='Writes the full preferences hashtable by design.')]
    param([Parameter(Mandatory)][hashtable]$Prefs)
    try {
        $Prefs | ConvertTo-Json | Set-Content -LiteralPath $global:PrefsPath -Encoding UTF8
    } catch { $null = $_ }
}

$global:Prefs = Get-HdPreferences

# =============================================================================
# Tool log.
# =============================================================================
$script:toolLogPath = Join-Path $__txDir ('HealthDash-{0}.log' -f (Get-Date -Format 'yyyyMMdd-HHmmss'))
Initialize-Logging -LogPath $script:toolLogPath

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

$txtAppTitle   = $window.FindName('txtAppTitle')
$txtVersion    = $window.FindName('txtVersion')
$txtThemeLabel = $window.FindName('txtThemeLabel')
$toggleTheme   = $window.FindName('toggleTheme')

$btnViewDeployments = $window.FindName('btnViewDeployments')
$btnViewContent     = $window.FindName('btnViewContent')
$btnViewDPs         = $window.FindName('btnViewDPs')
$btnViewClients     = $window.FindName('btnViewClients')
$btnViewInactive    = $window.FindName('btnViewInactive')
$btnViewSite        = $window.FindName('btnViewSite')
$btnOptions         = $window.FindName('btnOptions')

$txtModuleTitle    = $window.FindName('txtModuleTitle')
$txtModuleSubtitle = $window.FindName('txtModuleSubtitle')

$btnRefreshAll  = $window.FindName('btnRefreshAll')
$btnPauseResume = $window.FindName('btnPauseResume')
$txtFilter      = $window.FindName('txtFilter')
$cboStatus      = $window.FindName('cboStatus')
$btnExportCsv   = $window.FindName('btnExportCsv')
$btnExportHtml  = $window.FindName('btnExportHtml')
$btnCopySummary = $window.FindName('btnCopySummary')

$viewDeployments = $window.FindName('viewDeployments')
$viewContent     = $window.FindName('viewContent')
$viewDPs         = $window.FindName('viewDPs')
$viewClients     = $window.FindName('viewClients')
$viewInactive    = $window.FindName('viewInactive')
$viewSite        = $window.FindName('viewSite')

$gridDeploy   = $window.FindName('gridDeploy');   $txtDeployDetail   = $window.FindName('txtDeployDetail')
$gridContent  = $window.FindName('gridContent');  $txtContentDetail  = $window.FindName('txtContentDetail')
$gridDPs      = $window.FindName('gridDPs');      $txtDPDetail       = $window.FindName('txtDPDetail')
$gridClients  = $window.FindName('gridClients');  $txtClientDetail   = $window.FindName('txtClientDetail')
$gridInactive = $window.FindName('gridInactive'); $txtInactiveDetail = $window.FindName('txtInactiveDetail')
$gridSite     = $window.FindName('gridSite');     $txtSiteDetail     = $window.FindName('txtSiteDetail')

# Options is a modal dialog (Show-OptionsDialog), not an inline view -- see
# the comment in MainWindow.xaml where the inline ScrollViewer used to sit.

$progressOverlay  = $window.FindName('progressOverlay')
$txtProgressTitle = $window.FindName('txtProgressTitle')
$txtProgressStep  = $window.FindName('txtProgressStep')

$lblLogOutput = $window.FindName('lblLogOutput')
$txtLog       = $window.FindName('txtLog')
$txtStatus    = $window.FindName('txtStatus')

$null = $txtAppTitle, $txtVersion

# =============================================================================
# Log drawer + status bar helpers.
# Both pipe to the file log so post-close diagnostics survive.
# =============================================================================
function Add-LogLine {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification='Appends one line to the in-window TextBox + the on-disk log; idempotent by design.')]
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('INFO','WARN','ERROR')][string]$Level = 'INFO'
    )
    $ts = (Get-Date).ToString('HH:mm:ss')
    $line = '{0}  {1}' -f $ts, $Message
    if ([string]::IsNullOrWhiteSpace($txtLog.Text)) {
        $txtLog.Text = $line
    } else {
        $txtLog.AppendText([Environment]::NewLine + $line)
    }
    $txtLog.ScrollToEnd()
    try { Write-Log -Message $Message -Level $Level -Quiet } catch { $null = $_ }
}

function Set-StatusText {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification='Updates an in-window TextBlock only; no external state.')]
    param([Parameter(Mandatory)][string]$Text)
    $txtStatus.Text = $Text
}

# =============================================================================
# Title-bar drag fallback.
# Some VS Code PowerShell launch contexts can leave MahApps' custom title thumb
# unable to initiate native window move. Match the proven mecm-homelab fix:
# first return HTCAPTION from WM_NCHITTEST for the title band, then keep a
# managed DragMove fallback for hosts where HwndSource cannot be hooked.
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
        if (-not $parent -and $cur -is [System.Windows.FrameworkElement]) {
            $parent = $cur.Parent
        }
        if (-not $parent -and $cur -is [System.Windows.FrameworkContentElement]) {
            $parent = $cur.Parent
        }
        if (-not $parent -and $cur -is [System.Windows.ContentElement]) {
            try { $parent = [System.Windows.ContentOperations]::GetParent($cur) } catch { $parent = $null }
        }

        $cur = $parent
    }
}

function Test-IsWindowCommandPoint {
    param(
        [MahApps.Metro.Controls.MetroWindow]$Window,
        [System.Windows.Point]$Point
    )

    try {
        [void]$Window.ApplyTemplate()
        $commands = $Window.Template.FindName('PART_WindowButtonCommands', $Window)
        if ($commands -and $commands.IsVisible -and $commands.ActualWidth -gt 0 -and $commands.ActualHeight -gt 0) {
            $origin = $commands.TransformToAncestor($Window).Transform([System.Windows.Point]::new(0, 0))
            if ($Point.X -ge $origin.X -and
                $Point.X -le ($origin.X + $commands.ActualWidth) -and
                $Point.Y -ge $origin.Y -and
                $Point.Y -le ($origin.Y + $commands.ActualHeight)) {
                return $true
            }
        }
    } catch { $null = $_ }

    # Template lookup can fail before the first layout pass. Keep the
    # right-side caption buttons available with a conservative fallback.
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
            param(
                [IntPtr]$hwnd,
                [int]$msg,
                [IntPtr]$wParam,
                [IntPtr]$lParam,
                [ref]$handled
            )

            $WM_NCHITTEST = 0x0084
            $HTCAPTION = 2
            if ($msg -ne $WM_NCHITTEST) { return [IntPtr]::Zero }

            try {
                $windowKey = $hwnd.ToInt64().ToString()
                $target = $script:TitleBarHitTestWindows[$windowKey]
                if (-not $target) { return [IntPtr]::Zero }

                $raw = $lParam.ToInt64()
                $screenX = [int]($raw -band 0xffff)
                if ($screenX -ge 0x8000) { $screenX -= 0x10000 }
                $screenY = [int](($raw -shr 16) -band 0xffff)
                if ($screenY -ge 0x8000) { $screenY -= 0x10000 }
                $pt = $target.PointFromScreen([System.Windows.Point]::new($screenX, $screenY))
                $titleBarH = Get-TitleBarDragHeight -Window $target

                if ($pt.X -lt 0 -or $pt.X -gt $target.ActualWidth) { return [IntPtr]::Zero }
                if ($pt.Y -lt 4 -or $pt.Y -gt $titleBarH) { return [IntPtr]::Zero }
                if (Test-IsWindowCommandPoint -Window $target -Point $pt) { return [IntPtr]::Zero }

                $handled.Value = $true
                return [IntPtr]$HTCAPTION
            } catch {
                return [IntPtr]::Zero
            }
        }

        $script:TitleBarHitTestHooks[$key] = $hook
        $source.AddHook($hook)
    } catch {
        Add-LogLine ("Title-bar native hit-test fallback failed: {0}" -f $_.Exception.Message) -Level WARN
    }
}

function Install-TitleBarDragFallback {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification='Registers window-local WPF event handlers for title-bar drag fallback.')]
    param([MahApps.Metro.Controls.MetroWindow]$Window)

    $Window.Add_SourceInitialized({
        param($s, $e)
        Add-NativeTitleBarHitTestHook -Window $s
    })

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
        } catch {
            # DragMove throws if the mouse button is released before it starts.
            $null = $_
        }
    })
}

Install-TitleBarDragFallback -Window $window

# =============================================================================
# Theme setup and toggle.
# =============================================================================
[void][ControlzEx.Theming.ThemeManager]::Current.ChangeTheme($window, 'Dark.Steel')

$script:DarkButtonBg     = [System.Windows.Media.BrushConverter]::new().ConvertFrom('#1E1E1E')
$script:DarkButtonBorder = [System.Windows.Media.BrushConverter]::new().ConvertFrom('#555555')
$script:DarkActiveBg     = [System.Windows.Media.BrushConverter]::new().ConvertFrom('#3A3A3A')
$script:LightWfBg        = [System.Windows.Media.BrushConverter]::new().ConvertFrom('#0078D4')
$script:LightWfBorder    = [System.Windows.Media.BrushConverter]::new().ConvertFrom('#006CBE')
$script:LightActiveBg    = [System.Windows.Media.BrushConverter]::new().ConvertFrom('#005A9E')

$script:TitleBarBlue         = [System.Windows.Media.BrushConverter]::new().ConvertFrom('#0078D4')
$script:TitleBarBlueInactive = [System.Windows.Media.BrushConverter]::new().ConvertFrom('#4BA3E0')

$script:LogLabelDark  = [System.Windows.Media.BrushConverter]::new().ConvertFrom('#B0B0B0')
$script:LogLabelLight = [System.Windows.Media.BrushConverter]::new().ConvertFrom('#595959')

$script:ViewButtons = @(
    @{ Name = 'Deployments';       Button = $btnViewDeployments },
    @{ Name = 'Content';           Button = $btnViewContent     },
    @{ Name = 'DPs';               Button = $btnViewDPs         },
    @{ Name = 'Clients';           Button = $btnViewClients     },
    @{ Name = 'Inactive';          Button = $btnViewInactive    },
    @{ Name = 'Site';              Button = $btnViewSite        }
)
$script:ActiveView = 'Deployments'

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
        # Options is a modal popup, never an active view -- always render idle.
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
# DEFERRED to $window.Add_Loaded. See the supersedence-auditor note about
# WindowChromeBehavior and NCHITTEST routing on title-bar drag.

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
    Save-HdPreferences -Prefs $global:Prefs
    Update-SidebarButtonTheme
    Update-TitleBarBrushes
    Add-LogLine ('Theme: {0}' -f $(if ($isDark) { 'dark' } else { 'light' }))
})

# =============================================================================
# View switching.
# =============================================================================
$script:ViewMeta = @{
    'Deployments' = @{ Title = 'Deployments';         Subtitle = 'Application, package, software-update, and task-sequence deployments. Refresh to populate.' }
    'Content'     = @{ Title = 'Content Distribution'; Subtitle = 'Only content with failed or in-progress DP-content pairs is shown. Healthy items are filtered out at the source.' }
    'DPs'         = @{ Title = 'Distribution Points';  Subtitle = 'DP roster with site assignment, status from SMS_SiteSystemSummarizer, and pull-DP flag.' }
    'Clients'     = @{ Title = 'Client Health';        Subtitle = 'Per-device CCM health and active status. Requires SQL Server access (CM_<site> database).' }
    'Inactive'    = @{ Title = 'Inactive Devices';     Subtitle = 'Devices exceeding the inactivity threshold (configured in Options). SQL-backed.' }
    'Site'        = @{ Title = 'Site Health';          Subtitle = 'Site components (SMS_ComponentSummarizer) + site-system roles (SMS_SiteSystemSummarizer) in one rollup.' }
}

function Set-ActiveView {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification='Updates in-window Visibility + header text only.')]
    param([Parameter(Mandatory)][ValidateSet('Deployments','Content','DPs','Clients','Inactive','Site')][string]$View)

    $script:ActiveView = $View

    $viewDeployments.Visibility = if ($View -eq 'Deployments') { [System.Windows.Visibility]::Visible } else { [System.Windows.Visibility]::Collapsed }
    $viewContent.Visibility     = if ($View -eq 'Content')     { [System.Windows.Visibility]::Visible } else { [System.Windows.Visibility]::Collapsed }
    $viewDPs.Visibility         = if ($View -eq 'DPs')         { [System.Windows.Visibility]::Visible } else { [System.Windows.Visibility]::Collapsed }
    $viewClients.Visibility     = if ($View -eq 'Clients')     { [System.Windows.Visibility]::Visible } else { [System.Windows.Visibility]::Collapsed }
    $viewInactive.Visibility    = if ($View -eq 'Inactive')    { [System.Windows.Visibility]::Visible } else { [System.Windows.Visibility]::Collapsed }
    $viewSite.Visibility        = if ($View -eq 'Site')        { [System.Windows.Visibility]::Visible } else { [System.Windows.Visibility]::Collapsed }

    $meta = $script:ViewMeta[$View]
    if ($meta) {
        $txtModuleTitle.Text    = $meta.Title
        $txtModuleSubtitle.Text = $meta.Subtitle
    }

    Update-SidebarButtonTheme
    Update-Filter
    Update-StatusBarSummary
}

$btnViewDeployments.Add_Click({ Set-ActiveView -View 'Deployments' })
$btnViewContent.Add_Click({     Set-ActiveView -View 'Content' })
$btnViewDPs.Add_Click({         Set-ActiveView -View 'DPs' })
$btnViewClients.Add_Click({     Set-ActiveView -View 'Clients' })
$btnViewInactive.Add_Click({    Set-ActiveView -View 'Inactive' })
$btnViewSite.Add_Click({        Set-ActiveView -View 'Site' })
# btnOptions handler is wired below where Show-OptionsDialog is defined.

# =============================================================================
# Crash handlers (PS51-WPF-010, PS51-WPF-011, PS51-WPF-025).
# =============================================================================
$global:__crashLog = Join-Path $__txDir ('HealthDash-crash-{0}.txt' -f (Get-Date -Format 'yyyyMMdd-HHmmss'))

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
# Glyphs (per brand: check, x, warn, ellipsis at ThemeForeground).
# 0x2713=check, 0x2717=x, 0x26A0=warn, 0x22EF=ellipsis.
# =============================================================================
$script:Glyph = @{
    OK      = [char]0x2713
    Error   = [char]0x2717
    Warn    = [char]0x26A0
    Unknown = [char]0x22EF
}

function Get-DeploymentGlyph {
    param($Row)
    if ([int]$Row.NumberErrors -gt 0)     { return $script:Glyph.Error }
    # Zero-target check must come BEFORE the compliance check; otherwise an
    # untargeted deployment (Targeted=0, Compliance=0) trips the "<95%" warn
    # branch and never reaches the Unknown fallback.
    if ([int]$Row.NumberTargeted -eq 0)   { return $script:Glyph.Unknown }
    if ([int]$Row.NumberInProgress -gt 0) { return $script:Glyph.Warn }
    if ([double]$Row.PercentCompliant -lt 95) { return $script:Glyph.Warn }
    return $script:Glyph.OK
}

function Get-ContentGlyph {
    param($Row)
    if ([int]$Row.FailedCount -gt 0)     { return $script:Glyph.Error }
    if ([int]$Row.InProgressCount -gt 0) { return $script:Glyph.Warn }
    return $script:Glyph.OK
}

function Get-DPGlyph {
    param($Row)
    switch ([string]$Row.Status) {
        'OK'       { return $script:Glyph.OK }
        'Warning'  { return $script:Glyph.Warn }
        'Critical' { return $script:Glyph.Error }
        default    { return $script:Glyph.Unknown }
    }
}

function Get-ClientGlyph {
    param($Row)
    if ([string]$Row.HealthState  -eq 'Unhealthy') { return $script:Glyph.Error }
    if ([string]$Row.ActiveStatus -eq 'Inactive')  { return $script:Glyph.Warn }
    if ([string]$Row.HealthState  -eq 'Healthy')   { return $script:Glyph.OK }
    return $script:Glyph.Unknown
}

function Get-InactiveGlyph {
    param($Row)
    # Every row in the Inactive view is already past the user's configured
    # threshold (the SQL query filters on DaysSinceContact > threshold), so
    # the floor is Warn -- never Unknown. The 60-day boundary escalates to
    # Error.
    $days = [int]$Row.DaysSinceContact
    if ($days -gt 60) { return $script:Glyph.Error }
    return $script:Glyph.Warn
}

function Get-SiteGlyph {
    param($Row)
    switch ([string]$Row.Status) {
        'OK'       { return $script:Glyph.OK }
        'Warning'  { return $script:Glyph.Warn }
        'Critical' { return $script:Glyph.Error }
        default    { return $script:Glyph.Unknown }
    }
}

# =============================================================================
# Refresh state.
# =============================================================================
$script:DeploymentRows = @()
$script:ContentRows    = @()
$script:DPRows         = @()
$script:ClientRows     = @()
$script:InactiveRows   = @()
$script:SiteRows       = @()

$script:DeploymentCounts = $null
$script:ContentCounts    = $null
$script:DPCounts         = $null
$script:ClientCounts     = $null
$script:InactiveCounts   = $null
$script:SiteCounts       = $null

$script:LastRefreshTime     = $null
$script:IsConnectedFromBg   = $false
$script:SQLConnectedFromBg  = $false

# =============================================================================
# Decorators: enrich raw module rows with StatusGlyph + display fields.
# =============================================================================
function Format-DateOrEmpty {
    param($Value)
    if ($null -eq $Value) { return '' }
    $dt = $Value -as [datetime]
    if ($dt) { return $dt.ToString('yyyy-MM-dd HH:mm') }
    return [string]$Value
}

function ConvertTo-DeploymentGridRows {
    param($Rows)
    $out = @()
    foreach ($r in @($Rows)) {
        $out += [PSCustomObject]@{
            StatusGlyph      = Get-DeploymentGlyph -Row $r
            DeploymentId     = $r.DeploymentId
            DeploymentName   = $r.DeploymentName
            DeploymentType   = $r.DeploymentType
            CollectionName   = $r.CollectionName
            Purpose          = $r.Purpose
            NumberTargeted   = $r.NumberTargeted
            NumberSuccess    = $r.NumberSuccess
            NumberErrors     = $r.NumberErrors
            NumberInProgress = $r.NumberInProgress
            NumberUnknown    = $r.NumberUnknown
            PercentCompliant = $r.PercentCompliant
        }
    }
    return ,$out
}

function ConvertTo-ContentGridRows {
    param($Rows, $NameMap)
    $out = @()
    foreach ($r in @($Rows)) {
        $info = if ($NameMap) { $NameMap[$r.PackageID] } else { $null }
        $name = if ($info) { $info.Name } else { $r.PackageID }
        $type = if ($info) { $info.Type } else { 'Application' }
        $out += [PSCustomObject]@{
            StatusGlyph     = Get-ContentGlyph -Row $r
            ContentName     = $name
            ContentType     = $type
            PackageID       = $r.PackageID
            TotalDPs        = $r.TotalDPs
            InstalledCount  = $r.InstalledCount
            FailedCount     = $r.FailedCount
            InProgressCount = $r.InProgressCount
        }
    }
    return ,$out
}

function ConvertTo-DPGridRows {
    param($Rows)
    $out = @()
    foreach ($r in @($Rows)) {
        $out += [PSCustomObject]@{
            StatusGlyph   = Get-DPGlyph -Row $r
            DPName        = $r.DPName
            SiteCode      = $r.SiteCode
            Status        = $r.Status
            TotalContent  = $r.TotalContent
            FailedContent = $r.FailedContent
            IsPullDP      = $r.IsPullDP
        }
    }
    return ,$out
}

function ConvertTo-ClientGridRows {
    param($Rows)
    $out = @()
    foreach ($r in @($Rows)) {
        $out += [PSCustomObject]@{
            StatusGlyph        = Get-ClientGlyph -Row $r
            DeviceName         = $r.DeviceName
            HealthState        = $r.HealthState
            ActiveStatus       = $r.ActiveStatus
            LastOnlineDisplay  = Format-DateOrEmpty $r.LastOnlineTime
            LastDDRDisplay     = Format-DateOrEmpty $r.LastDDR
            LastPolicyDisplay  = Format-DateOrEmpty $r.LastPolicyRequest
            LastHWDisplay      = Format-DateOrEmpty $r.LastHWInventory
            ClientVersion      = $r.ClientVersion
            OperatingSystem    = $r.OperatingSystem
        }
    }
    return ,$out
}

function ConvertTo-InactiveGridRows {
    param($Rows)
    $out = @()
    foreach ($r in @($Rows)) {
        $out += [PSCustomObject]@{
            StatusGlyph       = Get-InactiveGlyph -Row $r
            DeviceName        = $r.DeviceName
            LastOnlineDisplay = Format-DateOrEmpty $r.LastOnlineTime
            LastDDRDisplay    = Format-DateOrEmpty $r.LastDDR
            DaysSinceContact  = $r.DaysSinceContact
            OperatingSystem   = $r.OperatingSystem
            ClientVersion     = $r.ClientVersion
        }
    }
    return ,$out
}

function ConvertTo-SiteGridRows {
    param($ComponentRows, $SystemRows)
    $out = @()
    foreach ($r in @($ComponentRows)) {
        $out += [PSCustomObject]@{
            StatusGlyph        = Get-SiteGlyph -Row $r
            Name               = $r.ComponentName
            ItemType           = $r.ItemType
            MachineName        = $r.MachineName
            Status             = $r.Status
            State              = $r.State
            LastStartedDisplay = Format-DateOrEmpty $r.LastStarted
        }
    }
    foreach ($r in @($SystemRows)) {
        $out += [PSCustomObject]@{
            StatusGlyph        = Get-SiteGlyph -Row $r
            Name               = $r.RoleName
            ItemType           = $r.ItemType
            MachineName        = $r.ServerName
            Status             = $r.Status
            State              = ''
            LastStartedDisplay = ''
        }
    }
    return ,$out
}

# =============================================================================
# Filter / status combobox / detail panel wiring.
# =============================================================================
function Get-StatusFilterValue {
    if (-not $cboStatus.SelectedItem) { return 'All' }
    $item = $cboStatus.SelectedItem
    if ($item -is [System.Windows.Controls.ComboBoxItem]) { return [string]$item.Content }
    return [string]$item
}

function Test-RowMatches {
    param($Row, [string]$Filter)
    switch ($Filter) {
        'All' { return $true }
        'OK / Healthy'        { return ($Row.StatusGlyph -eq $script:Glyph.OK) }
        'Warning / In Progress' { return ($Row.StatusGlyph -eq $script:Glyph.Warn) }
        'Failed / Error'      { return ($Row.StatusGlyph -eq $script:Glyph.Error) }
        default { return $true }
    }
}

function Update-Filter {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification='Recomputes ItemsSource on the active grid only; no external state.')]
    param()

    $needle = ([string]$txtFilter.Text).Trim().ToLowerInvariant()
    $statusFilter = Get-StatusFilterValue

    function Test-Needle {
        param($Row, [string[]]$Fields)
        if (-not $needle) { return $true }
        foreach ($f in $Fields) {
            $val = $Row.PSObject.Properties[$f].Value
            if ($null -eq $val) { continue }
            if (([string]$val).ToLowerInvariant().Contains($needle)) { return $true }
        }
        return $false
    }

    switch ($script:ActiveView) {
        'Deployments' {
            $rows = $script:DeploymentRows |
                Where-Object { Test-Needle -Row $_ -Fields @('DeploymentName','CollectionName','DeploymentType') } |
                Where-Object { Test-RowMatches -Row $_ -Filter $statusFilter }
            $gridDeploy.ItemsSource = @($rows)
        }
        'Content' {
            $rows = $script:ContentRows |
                Where-Object { Test-Needle -Row $_ -Fields @('ContentName','PackageID','ContentType') } |
                Where-Object { Test-RowMatches -Row $_ -Filter $statusFilter }
            $gridContent.ItemsSource = @($rows)
        }
        'DPs' {
            $rows = $script:DPRows |
                Where-Object { Test-Needle -Row $_ -Fields @('DPName','SiteCode') } |
                Where-Object { Test-RowMatches -Row $_ -Filter $statusFilter }
            $gridDPs.ItemsSource = @($rows)
        }
        'Clients' {
            $rows = $script:ClientRows |
                Where-Object { Test-Needle -Row $_ -Fields @('DeviceName','HealthState','ActiveStatus','OperatingSystem') } |
                Where-Object { Test-RowMatches -Row $_ -Filter $statusFilter }
            $gridClients.ItemsSource = @($rows)
        }
        'Inactive' {
            $rows = $script:InactiveRows |
                Where-Object { Test-Needle -Row $_ -Fields @('DeviceName','OperatingSystem') } |
                Where-Object { Test-RowMatches -Row $_ -Filter $statusFilter }
            $gridInactive.ItemsSource = @($rows)
        }
        'Site' {
            $rows = $script:SiteRows |
                Where-Object { Test-Needle -Row $_ -Fields @('Name','MachineName','ItemType') } |
                Where-Object { Test-RowMatches -Row $_ -Filter $statusFilter }
            $gridSite.ItemsSource = @($rows)
        }
    }
}

$txtFilter.Add_TextChanged({ Update-Filter })
$cboStatus.Add_SelectionChanged({ Update-Filter })

# Detail panels.
$gridDeploy.Add_SelectionChanged({
    $row = $gridDeploy.SelectedItem
    if (-not $row) { $txtDeployDetail.Text = 'Select a deployment to see status breakdown.'; return }
    $lines = @(
        'DEPLOYMENT',
        ('-' * 40),
        ('Name:         {0}' -f $row.DeploymentName),
        ('Type:         {0}' -f $row.DeploymentType),
        ('Collection:   {0}' -f $row.CollectionName),
        ('Purpose:      {0}' -f $row.Purpose),
        '',
        ('Targeted:     {0}' -f $row.NumberTargeted),
        ('Success:      {0}' -f $row.NumberSuccess),
        ('Errors:       {0}' -f $row.NumberErrors),
        ('In Progress:  {0}' -f $row.NumberInProgress),
        ('Unknown:      {0}' -f $row.NumberUnknown),
        ('% Compliant:  {0}%' -f $row.PercentCompliant),
        '',
        ('Deployment ID: {0}' -f $row.DeploymentId)
    )
    $txtDeployDetail.Text = $lines -join [Environment]::NewLine
})

$gridContent.Add_SelectionChanged({
    $row = $gridContent.SelectedItem
    if (-not $row) { $txtContentDetail.Text = 'Select a content item to see DP breakdown.'; return }
    $lines = @(
        'CONTENT',
        ('-' * 40),
        ('Name:        {0}' -f $row.ContentName),
        ('Type:        {0}' -f $row.ContentType),
        ('Package ID:  {0}' -f $row.PackageID),
        '',
        ('Total DPs:   {0}' -f $row.TotalDPs),
        ('Installed:   {0}' -f $row.InstalledCount),
        ('Failed:      {0}' -f $row.FailedCount),
        ('In Progress: {0}' -f $row.InProgressCount)
    )
    $txtContentDetail.Text = $lines -join [Environment]::NewLine
})

$gridDPs.Add_SelectionChanged({
    $row = $gridDPs.SelectedItem
    if (-not $row) { $txtDPDetail.Text = 'Select a distribution point to see details.'; return }
    $lines = @(
        'DISTRIBUTION POINT',
        ('-' * 40),
        ('Name:          {0}' -f $row.DPName),
        ('Site:          {0}' -f $row.SiteCode),
        ('Status:        {0}' -f $row.Status),
        ('Pull DP:       {0}' -f $row.IsPullDP),
        ('Total Content: {0}' -f $row.TotalContent),
        ('Failed:        {0}' -f $row.FailedContent)
    )
    $txtDPDetail.Text = $lines -join [Environment]::NewLine
})

$gridClients.Add_SelectionChanged({
    $row = $gridClients.SelectedItem
    if (-not $row) { $txtClientDetail.Text = 'Select a device to see client health details.'; return }
    $lines = @(
        'CLIENT HEALTH',
        ('-' * 40),
        ('Device:        {0}' -f $row.DeviceName),
        ('Health State:  {0}' -f $row.HealthState),
        ('Active:        {0}' -f $row.ActiveStatus),
        ('Client:        {0}' -f $row.ClientVersion),
        ('OS:            {0}' -f $row.OperatingSystem),
        '',
        ('Last Online:   {0}' -f $row.LastOnlineDisplay),
        ('Last DDR:      {0}' -f $row.LastDDRDisplay),
        ('Last Policy:   {0}' -f $row.LastPolicyDisplay),
        ('Last HW Inv:   {0}' -f $row.LastHWDisplay)
    )
    $txtClientDetail.Text = $lines -join [Environment]::NewLine
})

$gridInactive.Add_SelectionChanged({
    $row = $gridInactive.SelectedItem
    if (-not $row) { $txtInactiveDetail.Text = 'Select a device to see contact history.'; return }
    $lines = @(
        'INACTIVE DEVICE',
        ('-' * 40),
        ('Device:           {0}' -f $row.DeviceName),
        ('Days Since:       {0}' -f $row.DaysSinceContact),
        ('Operating System: {0}' -f $row.OperatingSystem),
        ('Client Version:   {0}' -f $row.ClientVersion),
        '',
        ('Last Online:      {0}' -f $row.LastOnlineDisplay),
        ('Last DDR:         {0}' -f $row.LastDDRDisplay)
    )
    $txtInactiveDetail.Text = $lines -join [Environment]::NewLine
})

$gridSite.Add_SelectionChanged({
    $row = $gridSite.SelectedItem
    if (-not $row) { $txtSiteDetail.Text = 'Select a component or site system for details.'; return }
    $lines = @(
        'SITE HEALTH ITEM',
        ('-' * 40),
        ('Name:         {0}' -f $row.Name),
        ('Type:         {0}' -f $row.ItemType),
        ('Server:       {0}' -f $row.MachineName),
        ('Status:       {0}' -f $row.Status),
        ('State:        {0}' -f $row.State),
        ('Last Started: {0}' -f $row.LastStartedDisplay)
    )
    $txtSiteDetail.Text = $lines -join [Environment]::NewLine
})

# =============================================================================
# Status bar summary.
# =============================================================================
function Update-StatusBarSummary {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification='Updates an in-window TextBlock only.')]
    param()

    $parts = @()
    if (-not $global:Prefs.SiteCode -or -not $global:Prefs.SMSProvider) {
        $parts += 'Open Options to configure site code and SMS provider'
    } elseif ($script:IsConnectedFromBg) {
        $parts += "Connected to $($global:Prefs.SiteCode)"
    } else {
        $parts += 'Ready. Click Refresh All.'
    }

    if ($script:DeploymentCounts) { $parts += ('{0} deployments / {1} failed' -f $script:DeploymentCounts.TotalDeployments, $script:DeploymentCounts.FailedDeployments) }
    if ($script:ContentCounts)    { $parts += ('{0} content issues' -f $script:ContentCounts.TotalContentWithIssues) }
    if ($script:DPCounts)         { $parts += ('{0} DPs / {1} offline' -f $script:DPCounts.TotalDPs, $script:DPCounts.OfflineCount) }
    if ($script:ClientCounts)     { $parts += ('{0} unhealthy / {1} inactive' -f $script:ClientCounts.UnhealthyCount, $script:ClientCounts.InactiveCount) }
    if ($script:InactiveCounts)   { $parts += ('{0} stale devices' -f $script:InactiveCounts.InactiveCount) }
    if ($script:SiteCounts)       { $parts += ('site {0} OK / {1} crit' -f $script:SiteCounts.OKCount, $script:SiteCounts.CriticalCount) }

    if ($script:LastRefreshTime) {
        $parts += ('last refresh {0}' -f $script:LastRefreshTime.ToString('HH:mm:ss'))
    }
    if ($script:AutoRefreshPaused) {
        $parts += 'auto-refresh paused'
    }

    Set-StatusText ($parts -join '   |   ')
}

# =============================================================================
# Background refresh runspace. The six health queries run sequentially in a
# STA runspace so the UI stays responsive and the progress overlay animates.
# =============================================================================
$script:BgRunspace     = $null
$script:BgPowerShell   = $null
$script:BgInvokeHandle = $null
$script:RefreshState   = $null
$script:RefreshTimer   = $null

function Initialize-RefreshRunspace {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification='Lazy-init of the background refresh runspace; idempotent.')]
    param()
    if ($script:BgRunspace -and $script:BgRunspace.RunspaceStateInfo.State -eq 'Opened') { return }

    $script:BgRunspace = [runspacefactory]::CreateRunspace()
    $script:BgRunspace.ApartmentState = 'STA'
    $script:BgRunspace.ThreadOptions  = 'ReuseThread'
    $script:BgRunspace.Open()

    $modulePath = Join-Path $PSScriptRoot 'Module\MECMHealthDashCommon.psd1'

    # The bg runspace gets its own module instance with its own module-scoped
    # state, so the UI thread's Initialize-Logging call doesn't reach it.
    # Without the second Initialize-Logging -Attach, every Write-Log inside
    # the bg refresh path is silently dropped (module-scoped log path stays
    # $null) -- which is exactly the GUI-only logging gap the v1.0.x WinForms
    # shell suffered from.
    $initPS = [powershell]::Create()
    $initPS.Runspace = $script:BgRunspace
    [void]$initPS.AddScript({
        param($ModulePath, $LogPath)
        Import-Module -Name $ModulePath -Force -DisableNameChecking
        if ($LogPath) { Initialize-Logging -LogPath $LogPath -Attach }
    }).AddArgument($modulePath).AddArgument($script:toolLogPath)
    [void]$initPS.Invoke()
    $initPS.Dispose()
}

function Invoke-RefreshAll {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification='Posts work to the background runspace and arms a DispatcherTimer.')]
    param()

    if (-not $global:Prefs.SiteCode -or -not $global:Prefs.SMSProvider) {
        Add-LogLine 'Refresh: site code and SMS provider must be set in Options first.' -Level WARN
        Set-StatusText 'Open Options to configure site code and SMS provider, then refresh.'
        return
    }

    Initialize-RefreshRunspace

    if ($script:RefreshTimer) { try { $script:RefreshTimer.Stop() } catch { $null = $_ } }
    if ($script:BgPowerShell) {
        try { [void]$script:BgPowerShell.Stop() } catch { $null = $_ }
        try { $script:BgPowerShell.Dispose() }   catch { $null = $_ }
        $script:BgPowerShell = $null
    }

    # Pause auto-refresh ticking while a refresh runs.
    if ($script:AutoTimer) { $script:AutoTimer.Stop() }

    $script:RefreshState = [hashtable]::Synchronized(@{
        Step     = 'Connecting...'
        Done     = $false
        Result   = $null
        ErrorMsg = $null
    })

    $btnRefreshAll.IsEnabled = $false
    $txtProgressTitle.Text = 'Refreshing health data'
    $txtProgressStep.Text  = 'Connecting...'
    $progressOverlay.Visibility = [System.Windows.Visibility]::Visible
    Add-LogLine ('Refresh: site={0} provider={1} sql={2}' -f $global:Prefs.SiteCode, $global:Prefs.SMSProvider, $(if ($global:Prefs.SQLServer) { $global:Prefs.SQLServer } else { '(none)' }))
    Set-StatusText 'Refreshing...'

    $siteCode    = [string]$global:Prefs.SiteCode
    $smsProvider = [string]$global:Prefs.SMSProvider
    $sqlServer   = [string]$global:Prefs.SQLServer
    $threshold   = [int]$global:Prefs.InactiveThresholdDays

    $script:BgPowerShell = [powershell]::Create()
    $script:BgPowerShell.Runspace = $script:BgRunspace
    [void]$script:BgPowerShell.AddScript({
        param($SiteCode, $SMSProvider, $SQLServer, $ThresholdDays, $State)
        try {
            if (-not (Test-CMConnection)) {
                $State.Step = "Connecting to $SiteCode..."
                $ok = Connect-CMSite -SiteCode $SiteCode -SMSProvider $SMSProvider
                if (-not $ok) {
                    $State.ErrorMsg = "Failed to connect to site $SiteCode (provider $SMSProvider)."
                    return
                }
            }

            $sqlOk = $false
            if ($SQLServer) {
                $State.Step = "Testing SQL connection ($SQLServer)..."
                $sqlOk = Test-SQLConnection -SQLServer $SQLServer -SiteCode $SiteCode
            }

            $State.Step = 'Querying deployment health...'
            $deployData = @(Get-DeploymentHealth)
            $deployCounts = Get-DeploymentHealthCounts -DeploymentData $deployData

            $State.Step = 'Querying content distribution health...'
            $contentData = @(Get-ContentDistributionHealth -SMSProvider $SMSProvider -SiteCode $SiteCode)
            $contentCounts = Get-ContentHealthCounts -ContentData $contentData

            $State.Step = 'Resolving content names...'
            $nameMap = Get-ContentNameMap -SMSProvider $SMSProvider -SiteCode $SiteCode

            $State.Step = 'Querying distribution point health...'
            $dpData = @(Get-DPHealth -SMSProvider $SMSProvider -SiteCode $SiteCode)
            $dpCounts = Get-DPHealthCounts -DPData $dpData

            $clientData = @()
            $clientCounts = $null
            $inactiveData = @()
            $inactiveCounts = $null
            if ($sqlOk) {
                $State.Step = 'Querying client health (SQL)...'
                $clientData = @(Get-ClientHealthSummary -SQLServer $SQLServer -SiteCode $SiteCode)
                $clientCounts = Get-ClientHealthCounts -ClientData $clientData

                $State.Step = "Querying inactive devices (>$ThresholdDays days)..."
                $inactiveData = @(Get-InactiveDevices -SQLServer $SQLServer -SiteCode $SiteCode -ThresholdDays $ThresholdDays)
                $inactiveCounts = Get-InactiveDeviceCounts -DeviceData $inactiveData
            }

            $State.Step = 'Querying site component health...'
            $componentData = @(Get-SiteComponentHealth -SMSProvider $SMSProvider -SiteCode $SiteCode)
            $State.Step = 'Querying site system health...'
            $systemData = @(Get-SiteSystemHealth -SMSProvider $SMSProvider -SiteCode $SiteCode)
            $siteCounts = Get-SiteHealthCounts -ComponentData $componentData -SystemData $systemData

            $State.Result = [PSCustomObject]@{
                SqlOk           = $sqlOk
                DeploymentData  = $deployData
                DeploymentCounts = $deployCounts
                ContentData     = $contentData
                ContentCounts   = $contentCounts
                ContentNameMap  = $nameMap
                DPData          = $dpData
                DPCounts        = $dpCounts
                ClientData      = $clientData
                ClientCounts    = $clientCounts
                InactiveData    = $inactiveData
                InactiveCounts  = $inactiveCounts
                ComponentData   = $componentData
                SystemData      = $systemData
                SiteCounts      = $siteCounts
            }
        }
        catch {
            $State.ErrorMsg = $_.Exception.Message
        }
        finally {
            $State.Done = $true
        }
    }).AddArgument($siteCode).AddArgument($smsProvider).AddArgument($sqlServer).AddArgument($threshold).AddArgument($script:RefreshState)

    $script:BgInvokeHandle = $script:BgPowerShell.BeginInvoke()

    $script:RefreshTimer = New-Object System.Windows.Threading.DispatcherTimer
    $script:RefreshTimer.Interval = [TimeSpan]::FromMilliseconds(150)
    $script:RefreshTimer.Add_Tick({
        if ($script:RefreshState) {
            $current = [string]$script:RefreshState.Step
            if ($txtProgressStep.Text -ne $current) { $txtProgressStep.Text = $current }
        }
        if ($script:RefreshState -and $script:RefreshState.Done) {
            $script:RefreshTimer.Stop()
            try { [void]$script:BgPowerShell.EndInvoke($script:BgInvokeHandle) } catch { $null = $_ }
            try { $script:BgPowerShell.Dispose() } catch { $null = $_ }
            $script:BgPowerShell   = $null
            $script:BgInvokeHandle = $null

            if ($script:RefreshState.ErrorMsg) {
                $progressOverlay.Visibility = [System.Windows.Visibility]::Collapsed
                $btnRefreshAll.IsEnabled = $true
                $script:IsConnectedFromBg = $false
                Add-LogLine ('Refresh failed: {0}' -f $script:RefreshState.ErrorMsg) -Level ERROR
                Set-StatusText 'Refresh failed.'
                # Restart auto-refresh even on failure so a transient blip doesn't kill the cadence.
                Start-AutoRefreshIfEnabled
                return
            }

            $script:IsConnectedFromBg  = $true
            $r = $script:RefreshState.Result
            $script:SQLConnectedFromBg = [bool]$r.SqlOk

            $script:DeploymentCounts = $r.DeploymentCounts
            $script:ContentCounts    = $r.ContentCounts
            $script:DPCounts         = $r.DPCounts
            $script:ClientCounts     = $r.ClientCounts
            $script:InactiveCounts   = $r.InactiveCounts
            $script:SiteCounts       = $r.SiteCounts

            $script:DeploymentRows = ConvertTo-DeploymentGridRows -Rows $r.DeploymentData
            $script:ContentRows    = ConvertTo-ContentGridRows    -Rows $r.ContentData -NameMap $r.ContentNameMap
            $script:DPRows         = ConvertTo-DPGridRows         -Rows $r.DPData
            $script:ClientRows     = ConvertTo-ClientGridRows     -Rows $r.ClientData
            $script:InactiveRows   = ConvertTo-InactiveGridRows   -Rows $r.InactiveData
            $script:SiteRows       = ConvertTo-SiteGridRows       -ComponentRows $r.ComponentData -SystemRows $r.SystemData

            $script:LastRefreshTime = Get-Date

            Update-Filter
            Update-StatusBarSummary

            $progressOverlay.Visibility = [System.Windows.Visibility]::Collapsed
            $btnRefreshAll.IsEnabled = $true

            $sqlNote = if ($r.SqlOk) { 'SQL ok' } else { 'SQL skipped' }
            Add-LogLine ('Refresh complete: {0} deployments, {1} content issues, {2} DPs, {3} clients, {4} inactive, {5} site items ({6})' -f $script:DeploymentRows.Count, $script:ContentRows.Count, $script:DPRows.Count, $script:ClientRows.Count, $script:InactiveRows.Count, $script:SiteRows.Count, $sqlNote)

            Start-AutoRefreshIfEnabled
        }
    })
    $script:RefreshTimer.Start()
}

$btnRefreshAll.Add_Click({ Invoke-RefreshAll })

# =============================================================================
# Auto-refresh timer.
# =============================================================================
$script:AutoTimer         = $null
$script:AutoRefreshPaused = $false

function Start-AutoRefreshIfEnabled {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification='Mutates the in-process DispatcherTimer; idempotent.')]
    param()
    if ($script:AutoRefreshPaused) { return }
    $minutes = [int]$global:Prefs.AutoRefreshMinutes
    if ($minutes -le 0) { return }
    if (-not $script:AutoTimer) {
        $script:AutoTimer = New-Object System.Windows.Threading.DispatcherTimer
        $script:AutoTimer.Add_Tick({ Invoke-RefreshAll })
    }
    $script:AutoTimer.Interval = [TimeSpan]::FromMinutes($minutes)
    $script:AutoTimer.Start()
}

function Stop-AutoRefresh {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification='Stops the in-process DispatcherTimer; idempotent.')]
    param()
    if ($script:AutoTimer) { $script:AutoTimer.Stop() }
}

$btnPauseResume.Add_Click({
    $script:AutoRefreshPaused = -not $script:AutoRefreshPaused
    if ($script:AutoRefreshPaused) {
        Stop-AutoRefresh
        $btnPauseResume.Content = 'Resume Auto-Refresh'
        Add-LogLine 'Auto-refresh paused.'
    } else {
        $btnPauseResume.Content = 'Pause Auto-Refresh'
        Start-AutoRefreshIfEnabled
        Add-LogLine 'Auto-refresh resumed.'
    }
    Update-StatusBarSummary
})

# =============================================================================
# Options modal dialog. Replaces an earlier inline-view implementation that
# correlated with a title-bar drag failure under non-elevated VS Code launches
# (per Jason 2026-05-01). The supersedence-auditor / app-packager brand
# pattern is a separate MetroWindow popup, which keeps the main window's
# visual tree free of inline option controls.
# =============================================================================
function Show-OptionsDialog {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseSingularNouns', '', Justification='Modal dialog show / dispose; verb-noun reads as a single action.')]
    param()

    $dlgXaml = @'
<Controls:MetroWindow
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    xmlns:Controls="clr-namespace:MahApps.Metro.Controls;assembly=MahApps.Metro"
    Title="Options"
    Width="700" Height="560"
    MinWidth="600" MinHeight="540"
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
            <Style x:Key="CategoryRowStyle" TargetType="Button"
                   BasedOn="{StaticResource MahApps.Styles.Button.Square}">
                <Setter Property="Height" Value="36"/>
                <Setter Property="HorizontalContentAlignment" Value="Left"/>
                <Setter Property="Padding" Value="14,0,14,0"/>
                <Setter Property="FontSize" Value="13"/>
                <Setter Property="Controls:ControlsHelper.ContentCharacterCasing" Value="Normal"/>
                <Setter Property="Margin" Value="0"/>
            </Style>
            <Style x:Key="DialogButton" TargetType="Button"
                   BasedOn="{StaticResource MahApps.Styles.Button.Square}">
                <Setter Property="MinWidth" Value="90"/>
                <Setter Property="Height" Value="32"/>
                <Setter Property="Margin" Value="0,0,8,0"/>
                <Setter Property="Controls:ControlsHelper.ContentCharacterCasing" Value="Normal"/>
            </Style>
            <Style x:Key="DialogAccentButton" TargetType="Button"
                   BasedOn="{StaticResource MahApps.Styles.Button.Square.Accent}">
                <Setter Property="MinWidth" Value="90"/>
                <Setter Property="Height" Value="32"/>
                <Setter Property="Margin" Value="0,0,8,0"/>
                <Setter Property="Controls:ControlsHelper.ContentCharacterCasing" Value="Normal"/>
            </Style>
            <Style x:Key="OptLabel" TargetType="TextBlock">
                <Setter Property="FontSize" Value="11"/>
                <Setter Property="Foreground" Value="{DynamicResource MahApps.Brushes.Gray1}"/>
                <Setter Property="Margin" Value="0,12,0,2"/>
            </Style>
            <Style x:Key="OptHint" TargetType="TextBlock">
                <Setter Property="FontSize" Value="10"/>
                <Setter Property="Foreground" Value="{DynamicResource MahApps.Brushes.Gray3}"/>
                <Setter Property="TextWrapping" Value="Wrap"/>
                <Setter Property="Margin" Value="0,2,0,0"/>
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
                <Button x:Name="btnCatRefresh"    Content="Refresh"    Style="{StaticResource CategoryRowStyle}"/>
                <Button x:Name="btnCatAbout"      Content="About"      Style="{StaticResource CategoryRowStyle}"/>
            </StackPanel>
        </Border>

        <Border Grid.Column="1" Grid.Row="0" Background="{DynamicResource MahApps.Brushes.Gray8}"/>

        <Grid Grid.Column="2" Grid.Row="0" Margin="20,16,20,16">

            <StackPanel x:Name="paneConnection" Visibility="Visible">
                <TextBlock Text="MECM Connection" FontSize="13" FontWeight="SemiBold" Margin="0,0,0,10"/>

                <TextBlock Style="{StaticResource OptLabel}" Text="Site Code"/>
                <TextBox x:Name="txtSiteCode" FontSize="12" Padding="6,4,6,4" MaxLength="3"
                         Width="120" HorizontalAlignment="Left"
                         Controls:TextBoxHelper.Watermark="e.g. MCM"/>
                <TextBlock Style="{StaticResource OptHint}" Text="Three-character primary site code."/>

                <TextBlock Style="{StaticResource OptLabel}" Text="SMS Provider"/>
                <TextBox x:Name="txtSmsProvider" FontSize="12" Padding="6,4,6,4"
                         Controls:TextBoxHelper.Watermark="e.g. cm01.contoso.com"/>
                <TextBlock Style="{StaticResource OptHint}" Text="FQDN of the SMS Provider host (typically the primary site server)."/>

                <TextBlock Style="{StaticResource OptLabel}" Text="SQL Server"/>
                <TextBox x:Name="txtSqlServer" FontSize="12" Padding="6,4,6,4"
                         Controls:TextBoxHelper.Watermark="e.g. cm01.contoso.com (blank to skip SQL views)"/>
                <TextBlock Style="{StaticResource OptHint}"
                           Text="SQL instance hosting CM_&lt;site&gt; (FQDN or FQDN\InstanceName). Leave blank to skip the Client Health and Inactive Devices views."/>
            </StackPanel>

            <StackPanel x:Name="paneRefresh" Visibility="Collapsed">
                <TextBlock Text="Refresh" FontSize="13" FontWeight="SemiBold" Margin="0,0,0,10"/>

                <TextBlock Style="{StaticResource OptLabel}" Text="Auto-refresh interval (minutes)"/>
                <ComboBox x:Name="cboRefreshInterval" Width="160" HorizontalAlignment="Left" FontSize="12">
                    <ComboBoxItem Content="5"/>
                    <ComboBoxItem Content="10"/>
                    <ComboBoxItem Content="15"/>
                    <ComboBoxItem Content="30"/>
                    <ComboBoxItem Content="60"/>
                </ComboBox>
                <TextBlock Style="{StaticResource OptHint}" Text="Manual refresh resets the timer. Pause / Resume from the action bar."/>

                <TextBlock Style="{StaticResource OptLabel}" Text="Inactivity threshold (days)"/>
                <ComboBox x:Name="cboInactiveDays" Width="160" HorizontalAlignment="Left" FontSize="12">
                    <ComboBoxItem Content="7"/>
                    <ComboBoxItem Content="14"/>
                    <ComboBoxItem Content="30"/>
                    <ComboBoxItem Content="60"/>
                    <ComboBoxItem Content="90"/>
                </ComboBox>
                <TextBlock Style="{StaticResource OptHint}" Text="Devices with no DDR contact in this many days appear in the Inactive Devices view."/>
            </StackPanel>

            <StackPanel x:Name="paneAbout" Visibility="Collapsed">
                <TextBlock Text="About" FontSize="13" FontWeight="SemiBold" Margin="0,0,0,10"/>
                <TextBlock Text="MECM Health Dashboard v1.0.0" FontSize="13" FontWeight="SemiBold"/>
                <TextBlock Text="Single-pane environmental health for MECM sites: deployments, content distribution, distribution points, client health, inactive devices, and site components / systems. Glyph-only status indicators; no red / yellow / green coloring."
                           FontSize="12" TextWrapping="Wrap" Margin="0,8,0,0"/>
                <TextBlock Text="PowerShell 5.1 + WPF (MahApps.Metro). Data layer: ConfigurationManager cmdlets + WMI summarizers + Invoke-Sqlcmd against the CM_&lt;site&gt; database."
                           FontSize="12" TextWrapping="Wrap" Margin="0,12,0,0"/>
                <TextBlock Text="Author: Jason Ulbright. License: MIT." FontSize="11" Margin="0,16,0,0"
                           Foreground="{DynamicResource MahApps.Brushes.Gray1}"/>
            </StackPanel>
        </Grid>

        <Border Grid.Row="1" Grid.ColumnSpan="3" Padding="16,12,16,12">
            <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
                <Button x:Name="btnOk"     Content="OK"     Style="{StaticResource DialogAccentButton}" IsDefault="True"/>
                <Button x:Name="btnCancel" Content="Cancel" Style="{StaticResource DialogButton}" IsCancel="True"/>
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
    $btnCatRefresh    = $dlg.FindName('btnCatRefresh')
    $btnCatAbout      = $dlg.FindName('btnCatAbout')
    $paneConnection   = $dlg.FindName('paneConnection')
    $paneRefresh      = $dlg.FindName('paneRefresh')
    $paneAbout        = $dlg.FindName('paneAbout')
    $txtSiteCode      = $dlg.FindName('txtSiteCode')
    $txtSmsProvider   = $dlg.FindName('txtSmsProvider')
    $txtSqlServer     = $dlg.FindName('txtSqlServer')
    $cboRefreshInterval = $dlg.FindName('cboRefreshInterval')
    $cboInactiveDays    = $dlg.FindName('cboInactiveDays')
    $btnOk            = $dlg.FindName('btnOk')
    $btnCancel        = $dlg.FindName('btnCancel')

    # Seed controls from prefs.
    $txtSiteCode.Text    = [string]$global:Prefs.SiteCode
    $txtSmsProvider.Text = [string]$global:Prefs.SMSProvider
    $txtSqlServer.Text   = [string]$global:Prefs.SQLServer

    $minutes = [string][int]$global:Prefs.AutoRefreshMinutes
    foreach ($i in $cboRefreshInterval.Items) {
        if ([string]$i.Content -eq $minutes) { $cboRefreshInterval.SelectedItem = $i; break }
    }
    if (-not $cboRefreshInterval.SelectedItem) { $cboRefreshInterval.SelectedIndex = 2 }

    $days = [string][int]$global:Prefs.InactiveThresholdDays
    foreach ($i in $cboInactiveDays.Items) {
        if ([string]$i.Content -eq $days) { $cboInactiveDays.SelectedItem = $i; break }
    }
    if (-not $cboInactiveDays.SelectedItem) { $cboInactiveDays.SelectedIndex = 1 }

    $btnCatConnection.Add_Click({
        $paneConnection.Visibility = [System.Windows.Visibility]::Visible
        $paneRefresh.Visibility    = [System.Windows.Visibility]::Collapsed
        $paneAbout.Visibility      = [System.Windows.Visibility]::Collapsed
    })
    $btnCatRefresh.Add_Click({
        $paneConnection.Visibility = [System.Windows.Visibility]::Collapsed
        $paneRefresh.Visibility    = [System.Windows.Visibility]::Visible
        $paneAbout.Visibility      = [System.Windows.Visibility]::Collapsed
    })
    $btnCatAbout.Add_Click({
        $paneConnection.Visibility = [System.Windows.Visibility]::Collapsed
        $paneRefresh.Visibility    = [System.Windows.Visibility]::Collapsed
        $paneAbout.Visibility      = [System.Windows.Visibility]::Visible
    })

    $btnOk.Add_Click({
        $newSite     = ([string]$txtSiteCode.Text).Trim()
        $newProvider = ([string]$txtSmsProvider.Text).Trim()
        $newSql      = ([string]$txtSqlServer.Text).Trim()
        $newMinutes  = if ($cboRefreshInterval.SelectedItem) { [int]([string]$cboRefreshInterval.SelectedItem.Content) } else { 15 }
        $newDays     = if ($cboInactiveDays.SelectedItem)    { [int]([string]$cboInactiveDays.SelectedItem.Content) }    else { 14 }

        $connectionChanged = ($newSite -ne [string]$global:Prefs.SiteCode) -or
                             ($newProvider -ne [string]$global:Prefs.SMSProvider) -or
                             ($newSql -ne [string]$global:Prefs.SQLServer)

        $global:Prefs.SiteCode              = $newSite
        $global:Prefs.SMSProvider           = $newProvider
        $global:Prefs.SQLServer             = $newSql
        $global:Prefs.AutoRefreshMinutes    = $newMinutes
        $global:Prefs.InactiveThresholdDays = $newDays
        Save-HdPreferences -Prefs $global:Prefs
        Add-LogLine ('Options saved: site={0} provider={1} sql={2} interval={3}m threshold={4}d' -f $newSite, $newProvider, $(if ($newSql) { $newSql } else { '(none)' }), $newMinutes, $newDays)

        if ($connectionChanged) {
            # Bg runspace caches the prior CM connection. Recycle it so the
            # next refresh reconnects with the new site / provider / SQL
            # values.
            if ($script:RefreshTimer) {
                try { $script:RefreshTimer.Stop() } catch { $null = $_ }
                $script:RefreshTimer = $null
            }
            if ($script:BgPowerShell) {
                try { [void]$script:BgPowerShell.Stop() } catch { $null = $_ }
                try { $script:BgPowerShell.Dispose() }   catch { $null = $_ }
                $script:BgPowerShell = $null
            }
            if ($script:BgRunspace) {
                try { $script:BgRunspace.Close() }  catch { $null = $_ }
                try { $script:BgRunspace.Dispose() } catch { $null = $_ }
                $script:BgRunspace = $null
            }
            $script:BgInvokeHandle      = $null
            $script:RefreshState        = $null
            $script:IsConnectedFromBg   = $false
            $script:SQLConnectedFromBg  = $false

            $progressOverlay.Visibility = [System.Windows.Visibility]::Collapsed
            $btnRefreshAll.IsEnabled    = $true
        }

        # Reset auto-refresh cadence with the new interval.
        Stop-AutoRefresh
        if (-not $script:AutoRefreshPaused) { Start-AutoRefreshIfEnabled }

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
# Export and copy summary.
# =============================================================================
function ConvertTo-DataTableForExport {
    param([Parameter(Mandatory)]$Rows, [string[]]$Columns)
    $dt = New-Object System.Data.DataTable
    if ($Columns -and $Columns.Count -gt 0) {
        foreach ($c in $Columns) { [void]$dt.Columns.Add($c, [string]) }
    } elseif (@($Rows).Count -gt 0) {
        $first = @($Rows)[0]
        foreach ($p in $first.PSObject.Properties) {
            if ($p.Name -ne 'StatusGlyph') { [void]$dt.Columns.Add($p.Name, [string]) }
        }
    } else {
        return $dt
    }
    foreach ($row in @($Rows)) {
        $vals = @()
        foreach ($col in $dt.Columns) {
            $val = $row.PSObject.Properties[$col.ColumnName].Value
            $vals += [string]$val
        }
        [void]$dt.Rows.Add($vals)
    }
    return $dt
}

function Get-ActiveExportInfo {
    switch ($script:ActiveView) {
        'Deployments' {
            return @{
                Name    = 'Deployments'
                Columns = @('DeploymentName','DeploymentType','CollectionName','Purpose','NumberTargeted','NumberSuccess','NumberErrors','NumberInProgress','NumberUnknown','PercentCompliant')
                Rows    = $gridDeploy.ItemsSource
            }
        }
        'Content' {
            return @{
                Name    = 'Content'
                Columns = @('ContentName','ContentType','PackageID','TotalDPs','InstalledCount','FailedCount','InProgressCount')
                Rows    = $gridContent.ItemsSource
            }
        }
        'DPs' {
            return @{
                Name    = 'DistributionPoints'
                Columns = @('DPName','SiteCode','Status','TotalContent','FailedContent','IsPullDP')
                Rows    = $gridDPs.ItemsSource
            }
        }
        'Clients' {
            return @{
                Name    = 'ClientHealth'
                Columns = @('DeviceName','HealthState','ActiveStatus','LastOnlineDisplay','LastDDRDisplay','LastPolicyDisplay','LastHWDisplay','ClientVersion','OperatingSystem')
                Rows    = $gridClients.ItemsSource
            }
        }
        'Inactive' {
            return @{
                Name    = 'InactiveDevices'
                Columns = @('DeviceName','LastOnlineDisplay','LastDDRDisplay','DaysSinceContact','OperatingSystem','ClientVersion')
                Rows    = $gridInactive.ItemsSource
            }
        }
        'Site' {
            return @{
                Name    = 'SiteHealth'
                Columns = @('Name','ItemType','MachineName','Status','State','LastStartedDisplay')
                Rows    = $gridSite.ItemsSource
            }
        }
        default { return $null }
    }
}

$btnExportCsv.Add_Click({
    $info = Get-ActiveExportInfo
    if (-not $info -or -not $info.Rows -or @($info.Rows).Count -eq 0) {
        Add-LogLine 'Export CSV: nothing to export.' -Level WARN
        return
    }
    $sfd = New-Object Microsoft.Win32.SaveFileDialog
    $sfd.Filter = 'CSV files (*.csv)|*.csv'
    $sfd.FileName = ('HealthDash-{0}-{1}.csv' -f $info.Name, (Get-Date -Format 'yyyyMMdd-HHmmss'))
    $reportsDir = Join-Path $PSScriptRoot 'Reports'
    if (-not (Test-Path -LiteralPath $reportsDir)) { New-Item -ItemType Directory -Path $reportsDir -Force | Out-Null }
    $sfd.InitialDirectory = $reportsDir
    if ($sfd.ShowDialog() -eq $true) {
        $dt = ConvertTo-DataTableForExport -Rows $info.Rows -Columns $info.Columns
        Export-HealthStatusCsv -DataTable $dt -OutputPath $sfd.FileName
        Add-LogLine ('Exported CSV: {0}' -f $sfd.FileName)
    }
})

$btnExportHtml.Add_Click({
    $info = Get-ActiveExportInfo
    if (-not $info -or -not $info.Rows -or @($info.Rows).Count -eq 0) {
        Add-LogLine 'Export HTML: nothing to export.' -Level WARN
        return
    }
    $sfd = New-Object Microsoft.Win32.SaveFileDialog
    $sfd.Filter = 'HTML files (*.html)|*.html'
    $sfd.FileName = ('HealthDash-{0}-{1}.html' -f $info.Name, (Get-Date -Format 'yyyyMMdd-HHmmss'))
    $reportsDir = Join-Path $PSScriptRoot 'Reports'
    if (-not (Test-Path -LiteralPath $reportsDir)) { New-Item -ItemType Directory -Path $reportsDir -Force | Out-Null }
    $sfd.InitialDirectory = $reportsDir
    if ($sfd.ShowDialog() -eq $true) {
        $dt = ConvertTo-DataTableForExport -Rows $info.Rows -Columns $info.Columns
        Export-HealthStatusHtml -DataTable $dt -OutputPath $sfd.FileName -ReportTitle ('MECM Health - {0}' -f $info.Name)
        Add-LogLine ('Exported HTML: {0}' -f $sfd.FileName)
    }
})

$btnCopySummary.Add_Click({
    if (-not $script:DeploymentCounts) {
        Add-LogLine 'Copy Summary: no refresh data yet.' -Level WARN
        return
    }
    $summary = New-HealthSummaryText `
        -DeploymentCounts $script:DeploymentCounts `
        -ContentCounts    $script:ContentCounts `
        -DPCounts         $script:DPCounts `
        -ClientCounts     $script:ClientCounts `
        -InactiveCounts   $script:InactiveCounts `
        -SiteCounts       $script:SiteCounts
    [System.Windows.Clipboard]::SetText($summary)
    Add-LogLine 'Summary copied to clipboard.'
})

# =============================================================================
# Window state persistence.
# =============================================================================
$global:WindowStatePath = Join-Path $PSScriptRoot 'MECMHealthDash.windowstate.json'

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

        # Schema bridge: WinForms 1.0.x used X / Y / ActiveTab; WPF refresh uses
        # Left / Top / ActiveView. Legacy files would otherwise snap the window
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

        if ($s.ActiveView -in @('Deployments','Content','DPs','Clients','Inactive','Site')) {
            Set-ActiveView -View ([string]$s.ActiveView)
        } elseif ($null -ne $s.ActiveTab) {
            # WinForms 1.0.x stored the tab index. Map it onto the WPF view
            # name so an upgrade preserves the user's last-active tab instead
            # of defaulting to Deployments.
            $tabMap = @('Deployments','Content','DPs','Clients','Inactive','Site')
            $idx = [int]$s.ActiveTab
            if ($idx -ge 0 -and $idx -lt $tabMap.Count) {
                Set-ActiveView -View $tabMap[$idx]
            }
        }
    } catch { $null = $_ }
}

$window.Add_Closing({
    Save-WindowState
    Stop-AutoRefresh
    if ($script:RefreshTimer) { try { $script:RefreshTimer.Stop() } catch { $null = $_ } }
    if ($script:BgPowerShell) {
        try { [void]$script:BgPowerShell.Stop() } catch { $null = $_ }
        try { $script:BgPowerShell.Dispose() }   catch { $null = $_ }
    }
    if ($script:BgRunspace) {
        try { $script:BgRunspace.Close() }  catch { $null = $_ }
        try { $script:BgRunspace.Dispose() } catch { $null = $_ }
    }
    if (Test-CMConnection) { try { Disconnect-CMSite } catch { $null = $_ } }
})

$window.Add_Loaded({
    Restore-WindowState

    # Apply user theme prefs AFTER the chrome has fully attached. Calling
    # ChangeTheme + WindowTitleBrush mutation at script-top breaks the title
    # bar's NCHITTEST routing (drag becomes a no-op). See the supersedence-
    # auditor note where Update-TitleBarBrushes is defined.
    $isDark = [bool]$global:Prefs['DarkMode']
    if (-not $isDark) {
        [void][ControlzEx.Theming.ThemeManager]::Current.ChangeTheme($window, 'Light.Blue')
    }
    Update-TitleBarBrushes

    Update-StatusBarSummary
    Add-LogLine 'MECM Health Dashboard ready. Configure Site / Provider in Options, then click Refresh All.'
})

# =============================================================================
# Run.
# =============================================================================
[void]$window.ShowDialog()
try { Stop-Transcript | Out-Null } catch { $null = $_ }
