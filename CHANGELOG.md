# Changelog

All notable changes to MECM Health Dashboard are documented in this
file.

## [1.3.0] - 2026-08-16

### Changed

- **Window chrome, theming, and window-state persistence now come from
  the vendored `SuiteCommon` module** (0.3.0). The legacy tab-index
  bridge for pre-WPF state files is preserved. Fixed by the shared
  layer: this tool's drag fallback never unhooked closed windows, so
  every window (including each modal dialog) leaked its hook and window
  reference for the process lifetime.

## [1.2.0] - 2026-08-14

### Changed

- **Shared plumbing moved to the vendored `SuiteCommon` module.** Logging
  (`Initialize-Logging`, `Write-Log`) and CM site connection
  (`Connect-CMSite`, `Disconnect-CMSite`, `Test-CMConnection`) now load
  from `Lib\SuiteCommon\`, shared across the tool suite and synced from
  the suite-core repository instead of hand-edited per repo. The
  connection additionally gains behavior this tool's own copy lacked: a
  globally scoped CMSite PSDrive, normalized ConfigurationManager module
  path resolution with known-install-path fallback, provider rebind when
  the configured SMS Provider changes, and rebuild of a stale drive whose
  provider connection died.
- **Module manifest GUID corrected** to a unique value (it previously
  duplicated another tool's manifest GUID).

## [1.1.0] - 2026-07-17

Doc-verified data-layer corrections (every ConfigMgr WMI class, SQL
view, and cmdlet call checked against Microsoft Learn), close-crash
hardening, and two Phase 4 features: Trends and Alerts.

### Fixed (data layer)

- **Client Health / Inactive Devices SQL** — queries referenced
  `v_CH_ClientSummary.HealthState` and `.LastOnline`, which do not
  exist; both views failed on every refresh. Now selects
  `LastEvaluationHealthy` (1 pass / 2 fail / 3 unknown),
  `LastActiveTime`, and `ClientStateDescription`. Inactive-device
  aging uses `COALESCE(LastDDR, LastActiveTime)` so DDR-less clients
  still surface.
- **Deployment type labels** — `FeatureType` map was off by one:
  task sequences displayed as "Baseline" and baselines as "Other (6)".
  Corrected to the documented values (1 Application, 2 Program,
  5 Software Update, 6 Baseline, 7 Task Sequence, ...).
- **Deployment names** — grid used `ApplicationName`, which is only
  populated for application deployments; package / update / TS rows
  showed blank names. Now prefers `SoftwareName`.
- **Pull DP flag** — `IsPullDP` is not a top-level property of
  `SMS_SCI_SysResUse`; every DP showed "No". Now read from the
  embedded `Props` array.
- **Get-DeploymentDetails** — called `Get-CMDeploymentStatusDetails
  -DeploymentID`, a parameter that does not exist (the cmdlet only
  accepts `-InputObject`). Rewired through the documented chain
  (Get-CMDeployment -> per-feature-type status cmdlet -> details) and
  added the missing `Requirements Not Met` (3) status.
- **Content distribution states** — DP-content pairs in
  `REMOVAL_PENDING` (4) / `REMOVAL_RETRYING` (5) were counted in no
  bucket; now tallied as in-progress. Added Device Setting / Virtual
  App / Application labels to the content name map.
- **DP status rollup** — `SMS_SiteSystemSummarizer` has one instance
  per storage object, so a DP with several drives appeared with
  whichever status happened to come last; now the worst status wins.
- Removed `Set-CMQueryResultMaximum -Maximum 0` (current cmdlet
  library is unbounded by default; the 0 semantics are undocumented).

### Fixed (shell)

- **Crash on close** — the `Closing` handler could throw
  `PipelineStoppedException` during host teardown (see
  `Logs/HealthDash-crash-20260501-052310.txt`). The handler body is
  now fully guarded, background pipeline shutdown is asynchronous
  (`BeginStop` / `CloseAsync` instead of blocking `Stop` on the UI
  thread), and shutdown-time `PipelineStoppedException` is logged but
  no longer fatal.
- **Refresh re-entrancy** — a second refresh no longer stops the
  in-flight pipeline (which could freeze the UI mid-CIM-abort); it is
  ignored with a log line instead.
- **Auto-refresh now arms at launch** when a site is configured,
  instead of only after the first manual refresh.
- Client grid: "Last Online" column renamed "Last Active"
  (`LastActiveTime` semantics); detail panel gains Client State
  (`ClientStateDescription`).

### Added

- **Trends view** — per-metric rolling history charts (7 / 30 / 90
  days) drawn in pure WPF. A snapshot of all health counts is appended
  to `History/metrics-history.csv` at every completed refresh
  (180-day retention). Nine metrics: compliance %, deployments with
  errors, content issues, failed pairs, DPs critical / degraded,
  unhealthy clients, inactive devices, site criticals. Export CSV
  exports the charted series.
- **Alerts** — transition-based threshold alerts evaluated after each
  refresh: any critical site component / site system, any critical DP,
  any failed DP-content pair, or overall compliance below a
  configurable floor (70 / 80 / 90 / 95%). Local-only delivery per the
  roadmap: Windows toast + `Logs/HealthDash-alerts.log` + log drawer.
  Configured from the new **Options > Alerts** pane; alerts re-fire
  only after a metric recovers and breaches again.

## [1.0.0] - 2026-05-02

MECM Health Dashboard is a single-pane environmental health tool for
MECM (Configuration Manager) sites. Six health views (deployments,
content distribution, distribution points, client health, inactive
devices, and site components / systems) with auto-refresh,
glyph-based status indicators, and per-view export. Read-only by
design: only `Get-CM*` cmdlets, `SMS_*Summarizer` WMI classes, and
SQL `SELECT` against the `CM_<site>` database are issued.

Extract the zip and run `start-mecmhealthdashboard.ps1`.

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
