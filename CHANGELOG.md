# Changelog

All notable changes to MECM Health Dashboard are documented in this
file.

## [1.0.0] - 2026-05-02

MECM Health Dashboard is a single-pane environmental health tool for
MECM (Configuration Manager) sites. Six health views (deployments,
content distribution, distribution points, client health, inactive
devices, and site components / systems) with auto-refresh,
glyph-based status indicators, and per-view export. Read-only by
design: only `Get-CM*` cmdlets, `SMS_*Summarizer` WMI classes, and
SQL `SELECT` against the `CM_<site>` database are issued.

Ships as a zip + `install.ps1` wrapper; no MSI, no code signing
required.

### Shell

- **MahApps.Metro WPF shell** with sidebar navigation across six
  data views plus an Options modal. No menu bar.
- **Glyph status** — per-row `✓ ⚠ ✗ ⋯` glyph in the first column of
  every grid at `ThemeForeground`. No red / yellow / green coloring;
  status communicated by glyph shape (passes WCAG SC 1.4.1).
- **Action bar** (per data view) — Refresh All, Pause / Resume
  Auto-Refresh, filter, status filter, Export CSV, Export HTML, Copy
  Summary.
- **Options modal** — Site Code, SMS Provider, SQL Server, refresh
  interval, inactivity threshold, About. Sidebar category list
  (Connection / Refresh / About) with OK / Cancel footer.
- **Theme toggle** in the sidebar bottom (Dark.Steel / Light.Blue);
  no restart required.
- **Background STA runspace** for the six health queries with an
  animated MahApps `ProgressRing` overlay. The UI stays responsive
  while large environments load.
- **Crash logs** — Dispatcher and AppDomain unhandled exceptions are
  written to `Logs/HealthDash-crash-*.txt` for post-mortem.
- **File logging** — `Add-LogLine` pipes to both the in-window log
  drawer and the on-disk session log so refresh failures persist
  after the window closes.
- **Window-state persistence** — position, size, maximized, and
  active view survive across sessions.

### Data views

- **Deployments** — Name, Type, Collection, Purpose, Targeted,
  Success, Errors, In Progress, % Compliant. Per-deployment status
  breakdown in the detail panel.
  `Set-CMQueryResultMaximum -Maximum 0` removes the default 1000-row
  cap so large deployment lists populate fully.
- **Content** — Only content with failed or in-progress DP-content
  pairs is shown (healthy content filtered at source). Bulk
  `PackageID` → `ContentName` resolution via `SMS_PackageBaseclass`.
- **Distribution Points** — DP roster with site, status (from
  `SMS_SiteSystemSummarizer`), pull-DP flag.
- **Client Health** (SQL-backed) — Per-device CCM health, active
  status, last online / DDR / policy / HW inventory, client version.
  Queries `v_CH_ClientSummary` joined with `v_R_System`.
- **Inactive Devices** (SQL-backed) — Devices exceeding the
  configurable inactivity threshold (7 / 14 / 30 / 60 / 90 days);
  glyph escalates to Error past 60 days.
- **Site Health** (WMI-backed) — Combined view of site components
  (`SMS_ComponentSummarizer`) and site-system roles
  (`SMS_SiteSystemSummarizer`).
- All SQL and WMI queries use unlimited timeouts
  (`-QueryTimeout 0`, `-OperationTimeoutSec 0`) to support large
  environments on slow links.

### Auto-refresh

- Configurable interval: 5 / 10 / 15 / 30 / 60 minutes (default 15).
- Pause / Resume from the action bar. Manual refresh resets the
  timer.

### Export

- **CSV** — full active-view grid data, no glyph column.
- **HTML** — self-contained styled report with glyph status column +
  bold weight for non-zero failure / in-progress counts. No color
  coding (brand parity with the in-app glyph rule).

### Stack

- PowerShell 5.1 + WPF (MahApps.Metro). MahApps, ControlzEx, and
  Microsoft.Xaml.Behaviors ship vendored in `Lib/`.
- `ConfigurationManager` PowerShell module (CM cmdlets).
- `SqlServer` PowerShell module (`Invoke-Sqlcmd`).
- WMI via `Get-CimInstance` against `root\SMS\site_<code>`.
