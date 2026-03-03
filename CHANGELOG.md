# Changelog

All notable changes to the MECM Health Dashboard are documented in this file.

## [1.0.1] - 2026-03-03

### Fixed

- **Client health SQL query** -- Fixed invalid column name `LastOnlineTime` to `LastOnline` in `v_CH_ClientSummary` queries; affected both Client Health and Inactive Devices tabs (query failed silently, returning 0 records)
- **SQL error visibility** -- Client health and inactive device query failures now log a warning in the UI log panel (previously errors were only written to the log file)
- **SQL connection state** -- Changing the SQL Server in Preferences now resets the connection state so the new server is tested on the next refresh

## [1.0.0] - 2026-02-26

### Added

- **GUI application** (`start-mecmhealthdashboard.ps1`) with WinForms interface
  - Header panel with title and subtitle
  - Connection bar showing site code, SMS provider, SQL server, and connection status
  - Six color-coded summary cards (Deployments, Content, DPs, Clients, Devices, Site)
  - TabControl with 6 detail tabs
  - Adaptive filter bar (controls change per active tab)
  - DataGridView with color-coded rows (red = critical, orange = warning)
  - Detail panels with RichTextBox info and sub-grids
  - Live log console with timestamped progress messages
  - Status bar with connection, auto-refresh countdown, last refresh, and issue count

- **Summary cards panel**
  - Six clickable cards in a 2x3 FlowLayoutPanel
  - Color-coded borders and backgrounds (green/yellow/red)
  - Click any card to navigate to its corresponding tab

- **Deployments tab**
  - Columns: Name, Type, Collection, Purpose, Targeted, Success, Failed, InProgress, Unknown, % Compliant
  - Per-device drill-down for selected deployment via `Get-CMDeploymentStatusDetails`
  - Filter by deployment type (Application, SoftwareUpdate, Package, TaskSequence)

- **Content tab**
  - Columns: ContentName, Type, PackageID, TotalDPs, Installed, Failed, InProgress
  - Shows only content with distribution failures or in-progress items

- **Distribution Points tab**
  - Columns: DPName, SiteCode, Status, TotalContent, FailedContent, IsPullDP
  - Detail panel with DP info

- **Client Health tab** (SQL-backed)
  - Columns: DeviceName, HealthState, ActiveStatus, LastOnline, LastDDR, LastPolicyRequest, LastHWInventory, LastHealthEval, ClientVersion
  - Queries `v_CH_ClientSummary` joined with `v_R_System`
  - Filter by health state (Healthy, Unhealthy, Inactive)

- **Inactive Devices tab** (SQL-backed)
  - Columns: DeviceName, LastOnline, LastDDR, DaysSinceContact, OperatingSystem, ClientVersion
  - Configurable inactivity threshold (7, 14, 30, 60, 90 days)

- **Site Health tab** (WMI-backed)
  - Combined view of site components (`SMS_ComponentSummarizer`) and site systems (`SMS_SiteSystemSummarizer`)
  - Columns: Name, Type, MachineName, Status, State, AvailabilityState, LastStarted

- **Auto-refresh timer**
  - Configurable interval (5, 10, 15, 30, 60 minutes; default 15)
  - Status bar countdown display
  - Pause/resume via Actions menu
  - Manual refresh resets the timer

- **Dark mode** with full theme support
  - Custom `DarkToolStripRenderer` for MenuStrip and StatusStrip
  - Owner-draw TabControl headers with ClearType anti-aliasing
  - Configurable via File > Preferences
  - Persisted in `MECMHealthDash.prefs.json`

- **Export**
  - CSV export of active tab grid data
  - HTML export with color-coded status cells and embedded CSS

- **Menu bar** with File (Preferences, Exit), Actions (Refresh Now, Pause/Resume Auto-Refresh), View (tab shortcuts), Help (About)

- **Window state persistence** across sessions (`MECMHealthDash.windowstate.json`)

- **Core module** (`MECMHealthDashCommon.psm1`) with 24 exported functions
  - CM site connection management with PSDrive
  - SQL connection testing via Invoke-Sqlcmd
  - Deployment health queries via CM cmdlets
  - Content distribution health via bulk WMI (`SMS_PackageStatusDistPointsSummarizer`)
  - DP health via CM cmdlets + WMI (`SMS_SiteSystemSummarizer`)
  - Client health and inactive device queries via SQL
  - Site component and system health via WMI
  - CSV and HTML export with color-coded status cells
