<#
.SYNOPSIS
    Core module for MECM Environment Health Dashboard.

.DESCRIPTION
    Import this module to get:
      - Structured logging (Initialize-Logging, Write-Log)
      - CM site connection management (Connect-CMSite, Disconnect-CMSite, Test-CMConnection)
      - SQL connection testing (Test-SQLConnection)
      - Deployment health queries via CM cmdlets
      - Content distribution health via bulk WMI
      - Distribution point health via CM cmdlets + WMI
      - Client health and inactive device queries via SQL views
      - Site component and system health via WMI summarizers
      - Export to CSV, HTML, and clipboard summary

.EXAMPLE
    Import-Module "$PSScriptRoot\Module\MECMHealthDashCommon.psd1" -Force
    Initialize-Logging -LogPath "C:\temp\healthdash.log"
    Connect-CMSite -SiteCode 'MCM' -SMSProvider 'sccm.domain.com'
#>

# ---------------------------------------------------------------------------
# Module-scoped state
# ---------------------------------------------------------------------------

$script:__HDLogPath            = $null
$script:OriginalLocation       = $null
$script:ConnectedSiteCode      = $null
$script:ConnectedSMSProvider   = $null

# ---------------------------------------------------------------------------
# Logging
# ---------------------------------------------------------------------------

function Initialize-Logging {
    param([string]$LogPath)

    $script:__HDLogPath = $LogPath

    if ($LogPath) {
        $parentDir = Split-Path -Path $LogPath -Parent
        if ($parentDir -and -not (Test-Path -LiteralPath $parentDir)) {
            New-Item -ItemType Directory -Path $parentDir -Force | Out-Null
        }

        $header = "[{0}] [INFO ] === Log initialized ===" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
        Set-Content -LiteralPath $LogPath -Value $header -Encoding UTF8
    }
}

function Write-Log {
    <#
    .SYNOPSIS
        Writes a timestamped, severity-tagged log message.
    #>
    param(
        [AllowEmptyString()]
        [Parameter(Mandatory, Position = 0)]
        [string]$Message,

        [ValidateSet('INFO', 'WARN', 'ERROR')]
        [string]$Level = 'INFO',

        [switch]$Quiet
    )

    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $formatted = "[{0}] [{1,-5}] {2}" -f $timestamp, $Level, $Message

    if (-not $Quiet) {
        Write-Host $formatted
        if ($Level -eq 'ERROR') {
            $host.UI.WriteErrorLine($formatted)
        }
    }

    if ($script:__HDLogPath) {
        Add-Content -LiteralPath $script:__HDLogPath -Value $formatted -Encoding UTF8 -ErrorAction SilentlyContinue
    }
}

# ---------------------------------------------------------------------------
# CM Connection
# ---------------------------------------------------------------------------

function Connect-CMSite {
    <#
    .SYNOPSIS
        Imports the ConfigurationManager module, creates a PSDrive, and sets location.
    .DESCRIPTION
        Returns $true on success, $false on failure.
    #>
    param(
        [Parameter(Mandatory)][string]$SiteCode,
        [Parameter(Mandatory)][string]$SMSProvider
    )

    $script:OriginalLocation = Get-Location

    # Import CM module if not already loaded
    if (-not (Get-Module ConfigurationManager -ErrorAction SilentlyContinue)) {
        $cmModulePath = $null
        if ($env:SMS_ADMIN_UI_PATH) {
            $cmModulePath = Join-Path $env:SMS_ADMIN_UI_PATH '..\ConfigurationManager.psd1'
        }

        if (-not $cmModulePath -or -not (Test-Path -LiteralPath $cmModulePath)) {
            Write-Log "ConfigurationManager module not found. Ensure the CM console is installed." -Level ERROR
            return $false
        }

        try {
            Import-Module $cmModulePath -ErrorAction Stop
            Write-Log "Imported ConfigurationManager module"
        }
        catch {
            Write-Log "Failed to import ConfigurationManager module: $_" -Level ERROR
            return $false
        }
    }

    # Create PSDrive if needed
    if (-not (Get-PSDrive -Name $SiteCode -PSProvider CMSite -ErrorAction SilentlyContinue)) {
        try {
            New-PSDrive -Name $SiteCode -PSProvider CMSite -Root $SMSProvider -ErrorAction Stop | Out-Null
            Write-Log "Created PSDrive for site $SiteCode"
        }
        catch {
            Write-Log "Failed to create PSDrive for site $SiteCode : $_" -Level ERROR
            return $false
        }
    }

    try {
        Set-Location "${SiteCode}:" -ErrorAction Stop
        $site = Get-CMSite -SiteCode $SiteCode -ErrorAction Stop
        Write-Log "Connected to site $SiteCode ($($site.SiteName))"
        $script:ConnectedSiteCode    = $SiteCode
        $script:ConnectedSMSProvider = $SMSProvider
        return $true
    }
    catch {
        Write-Log "Failed to connect to site $SiteCode : $_" -Level ERROR
        return $false
    }
}

function Disconnect-CMSite {
    <#
    .SYNOPSIS
        Restores the original location before CM connection.
    #>
    if ($script:OriginalLocation) {
        try { Set-Location $script:OriginalLocation -ErrorAction SilentlyContinue } catch { }
    }
    $script:ConnectedSiteCode    = $null
    $script:ConnectedSMSProvider = $null
    Write-Log "Disconnected from CM site"
}

function Test-CMConnection {
    <#
    .SYNOPSIS
        Returns $true if currently connected to a CM site.
    #>
    if (-not $script:ConnectedSiteCode) { return $false }

    try {
        $drive = Get-PSDrive -Name $script:ConnectedSiteCode -PSProvider CMSite -ErrorAction Stop
        return ($null -ne $drive)
    }
    catch {
        return $false
    }
}

function Test-SQLConnection {
    <#
    .SYNOPSIS
        Tests SQL connectivity to the CM site database.
    .DESCRIPTION
        Returns $true if Invoke-Sqlcmd can reach CM_<SiteCode> on the specified SQL server.
    #>
    param(
        [Parameter(Mandatory)][string]$SQLServer,
        [Parameter(Mandatory)][string]$SiteCode
    )

    $dbName = "CM_$SiteCode"

    try {
        $result = Invoke-Sqlcmd -ServerInstance $SQLServer -Database $dbName -Query "SELECT 1 AS Test" -ErrorAction Stop
        if ($result.Test -eq 1) {
            Write-Log "SQL connection verified: $SQLServer / $dbName"
            return $true
        }
        return $false
    }
    catch {
        Write-Log "SQL connection failed to $SQLServer / $dbName : $_" -Level ERROR
        return $false
    }
}

# ---------------------------------------------------------------------------
# Deployment Health
# ---------------------------------------------------------------------------

function Get-DeploymentHealth {
    <#
    .SYNOPSIS
        Returns all deployments with status counts.
    #>
    Write-Log "Querying deployment health..."

    $deployments = Get-CMDeployment -ErrorAction Stop

    $results = foreach ($d in $deployments) {
        $targeted   = [int]$d.NumberTargeted
        $success    = [int]$d.NumberSuccess
        $errors     = [int]$d.NumberErrors
        $inProg     = [int]$d.NumberInProgress
        $unknown    = [int]$d.NumberUnknown
        $pctCompliant = if ($targeted -gt 0) { [math]::Round(($success / $targeted) * 100, 1) } else { 0 }

        $deployType = switch ($d.FeatureType) {
            1 { 'Application' }
            2 { 'Package' }
            5 { 'Software Update' }
            7 { 'Baseline' }
            8 { 'Task Sequence' }
            default { "Other ($($d.FeatureType))" }
        }

        $purpose = switch ($d.DeploymentIntent) {
            1 { 'Required' }
            2 { 'Available' }
            default { 'Unknown' }
        }

        [PSCustomObject]@{
            DeploymentId    = $d.DeploymentID
            DeploymentName  = $d.ApplicationName
            DeploymentType  = $deployType
            CollectionName  = $d.CollectionName
            Purpose         = $purpose
            NumberTargeted  = $targeted
            NumberSuccess   = $success
            NumberErrors    = $errors
            NumberInProgress = $inProg
            NumberUnknown   = $unknown
            PercentCompliant = $pctCompliant
        }
    }

    Write-Log "Found $($results.Count) deployments"
    return $results
}

function Get-DeploymentDetails {
    <#
    .SYNOPSIS
        Returns per-device status for a specific deployment.
    #>
    param(
        [Parameter(Mandatory)][string]$DeploymentId
    )

    Write-Log "Querying deployment details for $DeploymentId..."

    try {
        $details = Get-CMDeploymentStatusDetails -DeploymentID $DeploymentId -ErrorAction Stop

        $results = foreach ($d in $details) {
            [PSCustomObject]@{
                DeviceName     = $d.DeviceName
                StatusType     = $d.StatusType
                StatusDescription = switch ([int]$d.StatusType) {
                    1 { 'Success' }
                    2 { 'In Progress' }
                    4 { 'Unknown' }
                    5 { 'Error' }
                    default { "Status $($d.StatusType)" }
                }
                LastStatusTime = $d.SummarizationTime
            }
        }

        Write-Log "Retrieved $($results.Count) device status records"
        return $results
    }
    catch {
        Write-Log "Failed to get deployment details: $_" -Level WARN
        return @()
    }
}

function Get-DeploymentHealthCounts {
    <#
    .SYNOPSIS
        Returns aggregate deployment health counts.
    #>
    param(
        [Parameter(Mandatory)][PSCustomObject[]]$DeploymentData
    )

    $total = $DeploymentData.Count
    $withErrors = ($DeploymentData | Where-Object { $_.NumberErrors -gt 0 }).Count
    $totalTargeted = ($DeploymentData | Measure-Object -Property NumberTargeted -Sum).Sum
    $totalSuccess  = ($DeploymentData | Measure-Object -Property NumberSuccess -Sum).Sum
    $overallPct = if ($totalTargeted -gt 0) { [math]::Round(($totalSuccess / $totalTargeted) * 100, 1) } else { 0 }

    return [PSCustomObject]@{
        TotalDeployments  = $total
        FailedDeployments = $withErrors
        OverallCompliance = $overallPct
    }
}

# ---------------------------------------------------------------------------
# Content Distribution Health
# ---------------------------------------------------------------------------

function Get-ContentDistributionHealth {
    <#
    .SYNOPSIS
        Bulk WMI query for content distribution status, returns only items with failures or in-progress.
    #>
    param(
        [Parameter(Mandatory)][string]$SMSProvider,
        [Parameter(Mandatory)][string]$SiteCode
    )

    Write-Log "Running bulk content distribution status query..."

    $raw = Get-CimInstance -ComputerName $SMSProvider `
        -Namespace "root\SMS\site_$SiteCode" `
        -ClassName SMS_PackageStatusDistPointsSummarizer -ErrorAction Stop

    # Aggregate per PackageID using hashtable
    $byPackage = @{}
    foreach ($row in $raw) {
        $pkgId = $row.PackageID
        if (-not $byPackage.ContainsKey($pkgId)) {
            $byPackage[$pkgId] = @{ TotalDPs = 0; Installed = 0; InProgress = 0; Failed = 0 }
        }
        $byPackage[$pkgId].TotalDPs++

        switch ($row.State) {
            0       { $byPackage[$pkgId].Installed++ }
            8       { $byPackage[$pkgId].Installed++ }
            { $_ -in 1, 2, 7 } { $byPackage[$pkgId].InProgress++ }
            { $_ -in 3, 6 }    { $byPackage[$pkgId].Failed++ }
        }
    }

    # Filter to only items with failures or in-progress
    $results = foreach ($pkgId in $byPackage.Keys) {
        $s = $byPackage[$pkgId]
        if ($s.Failed -gt 0 -or $s.InProgress -gt 0) {
            $pct = if ($s.TotalDPs -gt 0) { [math]::Round(($s.Installed / $s.TotalDPs) * 100, 1) } else { 0 }

            [PSCustomObject]@{
                PackageID       = $pkgId
                TotalDPs        = $s.TotalDPs
                InstalledCount  = $s.Installed
                InProgressCount = $s.InProgress
                FailedCount     = $s.Failed
                PctComplete     = "$pct%"
            }
        }
    }

    Write-Log "Content health: $($results.Count) items with failures or in-progress out of $($byPackage.Count) total"
    return $results
}

function Get-ContentHealthCounts {
    <#
    .SYNOPSIS
        Returns aggregate content distribution health counts.
    #>
    param(
        [Parameter(Mandatory)][AllowEmptyCollection()][PSCustomObject[]]$ContentData
    )

    $totalWithIssues = $ContentData.Count
    $totalFailedPairs = ($ContentData | Measure-Object -Property FailedCount -Sum).Sum

    return [PSCustomObject]@{
        TotalContentWithIssues = $totalWithIssues
        TotalFailedPairs       = [int]$totalFailedPairs
    }
}

# ---------------------------------------------------------------------------
# Distribution Point Health
# ---------------------------------------------------------------------------

function Get-DPHealth {
    <#
    .SYNOPSIS
        Returns DP health by combining CM cmdlets with WMI site system summarizer.
    #>
    param(
        [Parameter(Mandatory)][string]$SMSProvider,
        [Parameter(Mandatory)][string]$SiteCode
    )

    Write-Log "Querying distribution point health..."

    # Get DPs from CM
    $dps = Get-CMDistributionPoint -ErrorAction Stop

    # Get site system status via WMI
    $sysStatus = Get-CimInstance -ComputerName $SMSProvider `
        -Namespace "root\SMS\site_$SiteCode" `
        -ClassName SMS_SiteSystemSummarizer `
        -Filter "Role = 'SMS Distribution Point'" -ErrorAction SilentlyContinue

    $statusLookup = @{}
    if ($sysStatus) {
        foreach ($ss in $sysStatus) {
            $name = ''
            if ($ss.SiteSystem -match '\\\\([^\\]+)\\?') {
                $name = $Matches[1].ToUpper()
            }
            if ($name) {
                $statusLookup[$name] = [int]$ss.Status
            }
        }
    }

    $results = foreach ($dp in $dps) {
        $serverName = ''
        if ($dp.NetworkOSPath -match '\\\\(.+)') {
            $serverName = $Matches[1].TrimEnd('\').ToUpper()
        }

        $siteCode = $dp.SiteCode
        $isPullDP = if ($dp.IsPullDP) { 'Yes' } else { 'No' }

        $statusVal = if ($statusLookup.ContainsKey($serverName)) { $statusLookup[$serverName] } else { -1 }
        $statusText = switch ($statusVal) {
            0 { 'OK' }
            1 { 'Warning' }
            2 { 'Critical' }
            default { 'Unknown' }
        }

        [PSCustomObject]@{
            DPName        = $serverName
            SiteCode      = $siteCode
            Status        = $statusText
            StatusValue   = $statusVal
            IsPullDP      = $isPullDP
            TotalContent  = 0
            FailedContent = 0
        }
    }

    Write-Log "Found $($results.Count) distribution points"
    return $results
}

function Get-DPDetails {
    <#
    .SYNOPSIS
        Returns content status for a specific DP from cached bulk status data.
    #>
    param(
        [Parameter(Mandatory)][string]$DPName,
        [Parameter(Mandatory)][AllowEmptyCollection()][PSCustomObject[]]$StatusRows
    )

    $dpRows = $StatusRows | Where-Object { $_.DPName -eq $DPName }
    return $dpRows
}

function Get-DPHealthCounts {
    <#
    .SYNOPSIS
        Returns aggregate DP health counts.
    #>
    param(
        [Parameter(Mandatory)][AllowEmptyCollection()][PSCustomObject[]]$DPData
    )

    $total   = $DPData.Count
    $offline = ($DPData | Where-Object { $_.Status -eq 'Critical' }).Count
    $warning = ($DPData | Where-Object { $_.Status -eq 'Warning' }).Count

    return [PSCustomObject]@{
        TotalDPs      = $total
        OfflineCount  = $offline
        DegradedCount = $warning
    }
}

# ---------------------------------------------------------------------------
# Client Health (SQL)
# ---------------------------------------------------------------------------

function Get-ClientHealthSummary {
    <#
    .SYNOPSIS
        Queries CM database for client health summary data.
    #>
    param(
        [Parameter(Mandatory)][string]$SQLServer,
        [Parameter(Mandatory)][string]$SiteCode
    )

    Write-Log "Querying client health from SQL ($SQLServer)..."

    $dbName = "CM_$SiteCode"
    $query = @(
        "SELECT",
        "    sys.Name0 AS DeviceName,",
        "    ISNULL(ch.HealthState, 0) AS HealthState,",
        "    ISNULL(ch.ClientActiveStatus, 0) AS ClientActiveStatus,",
        "    ch.LastOnlineTime,",
        "    ch.LastDDR,",
        "    ch.LastPolicyRequest,",
        "    ch.LastHW AS LastHWInventory,",
        "    ch.LastHealthEvaluation,",
        "    sys.Client_Version0 AS ClientVersion,",
        "    sys.Operating_System_Name_and0 AS OperatingSystem",
        "FROM v_CH_ClientSummary ch",
        "JOIN v_R_System sys ON ch.ResourceID = sys.ResourceID",
        "WHERE sys.Client0 = 1"
    ) -join "`r`n"

    try {
        $rows = Invoke-Sqlcmd -ServerInstance $SQLServer -Database $dbName -Query $query -ErrorAction Stop

        $results = foreach ($r in $rows) {
            $healthText = switch ([int]$r.HealthState) {
                1 { 'Healthy' }
                2 { 'Unhealthy' }
                default { 'Unknown' }
            }
            $activeText = switch ([int]$r.ClientActiveStatus) {
                1 { 'Active' }
                0 { 'Inactive' }
                default { 'Unknown' }
            }

            [PSCustomObject]@{
                DeviceName          = $r.DeviceName
                HealthState         = $healthText
                HealthStateValue    = [int]$r.HealthState
                ActiveStatus        = $activeText
                ActiveStatusValue   = [int]$r.ClientActiveStatus
                LastOnlineTime      = $r.LastOnlineTime
                LastDDR             = $r.LastDDR
                LastPolicyRequest   = $r.LastPolicyRequest
                LastHWInventory     = $r.LastHWInventory
                LastHealthEvaluation = $r.LastHealthEvaluation
                ClientVersion       = $r.ClientVersion
                OperatingSystem     = $r.OperatingSystem
            }
        }

        Write-Log "Retrieved $($results.Count) client health records"
        return $results
    }
    catch {
        Write-Log "Client health SQL query failed: $_" -Level ERROR
        return @()
    }
}

function Get-ClientHealthCounts {
    <#
    .SYNOPSIS
        Returns aggregate client health counts from cached data.
    #>
    param(
        [Parameter(Mandatory)][AllowEmptyCollection()][PSCustomObject[]]$ClientData
    )

    $healthy   = ($ClientData | Where-Object { $_.HealthStateValue -eq 1 }).Count
    $unhealthy = ($ClientData | Where-Object { $_.HealthStateValue -eq 2 }).Count
    $inactive  = ($ClientData | Where-Object { $_.ActiveStatusValue -eq 0 }).Count

    return [PSCustomObject]@{
        HealthyCount   = $healthy
        UnhealthyCount = $unhealthy
        InactiveCount  = $inactive
    }
}

# ---------------------------------------------------------------------------
# Inactive Devices (SQL)
# ---------------------------------------------------------------------------

function Get-InactiveDevices {
    <#
    .SYNOPSIS
        Queries CM database for devices exceeding the inactivity threshold.
    #>
    param(
        [Parameter(Mandatory)][string]$SQLServer,
        [Parameter(Mandatory)][string]$SiteCode,
        [int]$ThresholdDays = 14
    )

    Write-Log "Querying inactive devices (threshold: $ThresholdDays days)..."

    $dbName = "CM_$SiteCode"
    $query = @(
        "SELECT",
        "    sys.Name0 AS DeviceName,",
        "    ch.LastOnlineTime,",
        "    ch.LastDDR,",
        "    DATEDIFF(day, ch.LastDDR, GETDATE()) AS DaysSinceContact,",
        "    sys.Operating_System_Name_and0 AS OperatingSystem,",
        "    sys.Client_Version0 AS ClientVersion",
        "FROM v_CH_ClientSummary ch",
        "JOIN v_R_System sys ON ch.ResourceID = sys.ResourceID",
        "WHERE sys.Client0 = 1",
        "  AND DATEDIFF(day, ch.LastDDR, GETDATE()) > $ThresholdDays",
        "ORDER BY DaysSinceContact DESC"
    ) -join "`r`n"

    try {
        $rows = Invoke-Sqlcmd -ServerInstance $SQLServer -Database $dbName -Query $query -ErrorAction Stop

        $results = foreach ($r in $rows) {
            [PSCustomObject]@{
                DeviceName       = $r.DeviceName
                LastOnlineTime   = $r.LastOnlineTime
                LastDDR          = $r.LastDDR
                DaysSinceContact = [int]$r.DaysSinceContact
                OperatingSystem  = $r.OperatingSystem
                ClientVersion    = $r.ClientVersion
            }
        }

        Write-Log "Found $($results.Count) inactive devices (>$ThresholdDays days)"
        return $results
    }
    catch {
        Write-Log "Inactive devices SQL query failed: $_" -Level ERROR
        return @()
    }
}

function Get-InactiveDeviceCounts {
    <#
    .SYNOPSIS
        Returns inactive device count from cached data.
    #>
    param(
        [Parameter(Mandatory)][AllowEmptyCollection()][PSCustomObject[]]$DeviceData
    )

    return [PSCustomObject]@{
        InactiveCount = $DeviceData.Count
    }
}

# ---------------------------------------------------------------------------
# Site Health (WMI)
# ---------------------------------------------------------------------------

function Get-SiteComponentHealth {
    <#
    .SYNOPSIS
        Queries SMS_ComponentSummarizer for component health status.
    #>
    param(
        [Parameter(Mandatory)][string]$SMSProvider,
        [Parameter(Mandatory)][string]$SiteCode
    )

    Write-Log "Querying site component health..."

    try {
        $raw = Get-CimInstance -ComputerName $SMSProvider `
            -Namespace "root\SMS\site_$SiteCode" `
            -Query "SELECT ComponentName, MachineName, Status, State, AvailabilityState, NextScheduledTime, LastStarted, TallyInterval FROM SMS_ComponentSummarizer WHERE TallyInterval = '0001128000100008'" `
            -ErrorAction Stop

        $results = foreach ($c in $raw) {
            $statusText = switch ([int]$c.Status) {
                0 { 'OK' }
                1 { 'Warning' }
                2 { 'Critical' }
                default { "Unknown ($($c.Status))" }
            }

            $stateText = switch ([int]$c.State) {
                0 { 'Stopped' }
                1 { 'Started' }
                2 { 'Paused' }
                3 { 'Installing' }
                4 { 'Re-installing' }
                5 { 'De-installing' }
                default { "Unknown ($($c.State))" }
            }

            [PSCustomObject]@{
                ComponentName    = $c.ComponentName
                MachineName      = $c.MachineName
                Status           = $statusText
                StatusValue      = [int]$c.Status
                State            = $stateText
                AvailabilityState = [int]$c.AvailabilityState
                LastStarted      = $c.LastStarted
                ItemType         = 'Component'
            }
        }

        Write-Log "Retrieved $($results.Count) component status records"
        return $results
    }
    catch {
        Write-Log "Site component health query failed: $_" -Level ERROR
        return @()
    }
}

function Get-SiteSystemHealth {
    <#
    .SYNOPSIS
        Queries SMS_SiteSystemSummarizer for site system role health.
    #>
    param(
        [Parameter(Mandatory)][string]$SMSProvider,
        [Parameter(Mandatory)][string]$SiteCode
    )

    Write-Log "Querying site system health..."

    try {
        $raw = Get-CimInstance -ComputerName $SMSProvider `
            -Namespace "root\SMS\site_$SiteCode" `
            -ClassName SMS_SiteSystemSummarizer -ErrorAction Stop

        $results = foreach ($s in $raw) {
            $serverName = ''
            if ($s.SiteSystem -match '\\\\([^\\]+)\\?') {
                $serverName = $Matches[1].ToUpper()
            }

            $statusText = switch ([int]$s.Status) {
                0 { 'OK' }
                1 { 'Warning' }
                2 { 'Critical' }
                default { "Unknown ($($s.Status))" }
            }

            [PSCustomObject]@{
                ServerName        = $serverName
                SiteCode          = $s.SiteCode
                RoleName          = $s.Role
                Status            = $statusText
                StatusValue       = [int]$s.Status
                AvailabilityState = [int]$s.AvailabilityState
                ItemType          = 'Site System'
            }
        }

        Write-Log "Retrieved $($results.Count) site system status records"
        return $results
    }
    catch {
        Write-Log "Site system health query failed: $_" -Level ERROR
        return @()
    }
}

function Get-SiteHealthCounts {
    <#
    .SYNOPSIS
        Returns aggregate site health counts.
    #>
    param(
        [AllowEmptyCollection()][PSCustomObject[]]$ComponentData = @(),
        [AllowEmptyCollection()][PSCustomObject[]]$SystemData = @()
    )

    $allItems = @($ComponentData) + @($SystemData)
    $ok       = ($allItems | Where-Object { $_.StatusValue -eq 0 }).Count
    $warning  = ($allItems | Where-Object { $_.StatusValue -eq 1 }).Count
    $critical = ($allItems | Where-Object { $_.StatusValue -eq 2 }).Count

    return [PSCustomObject]@{
        OKCount       = $ok
        WarningCount  = $warning
        CriticalCount = $critical
    }
}

# ---------------------------------------------------------------------------
# Export
# ---------------------------------------------------------------------------

function Export-HealthStatusCsv {
    <#
    .SYNOPSIS
        Exports a DataTable to CSV.
    #>
    param(
        [Parameter(Mandatory)][System.Data.DataTable]$DataTable,
        [Parameter(Mandatory)][string]$OutputPath
    )

    $parentDir = Split-Path -Path $OutputPath -Parent
    if ($parentDir -and -not (Test-Path -LiteralPath $parentDir)) {
        New-Item -ItemType Directory -Path $parentDir -Force | Out-Null
    }

    $rows = @()
    foreach ($row in $DataTable.Rows) {
        $obj = [ordered]@{}
        foreach ($col in $DataTable.Columns) {
            $obj[$col.ColumnName] = $row[$col.ColumnName]
        }
        $rows += [PSCustomObject]$obj
    }

    $rows | Export-Csv -LiteralPath $OutputPath -NoTypeInformation -Encoding UTF8
    Write-Log "Exported CSV to $OutputPath"
}

function Export-HealthStatusHtml {
    <#
    .SYNOPSIS
        Exports a DataTable to a self-contained HTML report with color-coded status cells.
    #>
    param(
        [Parameter(Mandatory)][System.Data.DataTable]$DataTable,
        [Parameter(Mandatory)][string]$OutputPath,
        [string]$ReportTitle = 'Health Status Report'
    )

    $parentDir = Split-Path -Path $OutputPath -Parent
    if ($parentDir -and -not (Test-Path -LiteralPath $parentDir)) {
        New-Item -ItemType Directory -Path $parentDir -Force | Out-Null
    }

    $css = @(
        '<style>',
        'body { font-family: "Segoe UI", Arial, sans-serif; margin: 20px; background: #fafafa; }',
        'h1 { color: #0078D4; margin-bottom: 4px; }',
        '.summary { color: #666; margin-bottom: 12px; font-size: 0.9em; }',
        'table { border-collapse: collapse; width: 100%; margin-top: 12px; }',
        'th { background: #0078D4; color: #fff; padding: 8px 12px; text-align: left; }',
        'td { padding: 6px 12px; border-bottom: 1px solid #e0e0e0; }',
        'tr:nth-child(even) { background: #f5f5f5; }',
        '.failed { color: #c00; font-weight: bold; }',
        '.warning { color: #b87800; }',
        '.ok { color: #228b22; }',
        '</style>'
    ) -join "`r`n"

    $headerRow = ($DataTable.Columns | ForEach-Object { "<th>$($_.ColumnName)</th>" }) -join ''
    $bodyRows = foreach ($row in $DataTable.Rows) {
        $cells = foreach ($col in $DataTable.Columns) {
            $val = [string]$row[$col.ColumnName]
            $cssClass = ''
            if ($col.ColumnName -match 'Failed|Error|Critical|Unhealthy' -and $val -match '^\d+$' -and [int]$val -gt 0) {
                $cssClass = ' class="failed"'
            }
            elseif ($col.ColumnName -match 'InProgress|Warning|Degraded' -and $val -match '^\d+$' -and [int]$val -gt 0) {
                $cssClass = ' class="warning"'
            }
            elseif ($val -in 'OK', 'Healthy', 'Active') {
                $cssClass = ' class="ok"'
            }
            elseif ($val -in 'Critical', 'Error', 'Unhealthy', 'Failed') {
                $cssClass = ' class="failed"'
            }
            elseif ($val -in 'Warning', 'Degraded', 'In Progress') {
                $cssClass = ' class="warning"'
            }
            "<td$cssClass>$val</td>"
        }
        "<tr>$($cells -join '')</tr>"
    }

    $html = @(
        '<!DOCTYPE html>',
        '<html><head><meta charset="utf-8"><title>' + $ReportTitle + '</title>',
        $css,
        '</head><body>',
        "<h1>$ReportTitle</h1>",
        "<div class='summary'>Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') | Rows: $($DataTable.Rows.Count)</div>",
        "<table><thead><tr>$headerRow</tr></thead>",
        "<tbody>$($bodyRows -join "`r`n")</tbody></table>",
        '</body></html>'
    ) -join "`r`n"

    Set-Content -LiteralPath $OutputPath -Value $html -Encoding UTF8
    Write-Log "Exported HTML to $OutputPath"
}

function New-HealthSummaryText {
    <#
    .SYNOPSIS
        Returns a plain text summary of all health counts for clipboard/log.
    #>
    param(
        [PSCustomObject]$DeploymentCounts,
        [PSCustomObject]$ContentCounts,
        [PSCustomObject]$DPCounts,
        [PSCustomObject]$ClientCounts,
        [PSCustomObject]$InactiveCounts,
        [PSCustomObject]$SiteCounts
    )

    $lines = @(
        "MECM Environment Health Summary - $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')",
        ("-" * 60),
        "Deployments:  $($DeploymentCounts.TotalDeployments) total, $($DeploymentCounts.FailedDeployments) with errors ($($DeploymentCounts.OverallCompliance)% compliant)",
        "Content:      $($ContentCounts.TotalContentWithIssues) items with issues, $($ContentCounts.TotalFailedPairs) failed DP-content pairs",
        "DPs:          $($DPCounts.TotalDPs) total, $($DPCounts.OfflineCount) offline, $($DPCounts.DegradedCount) degraded",
        "Clients:      $($ClientCounts.HealthyCount) healthy, $($ClientCounts.UnhealthyCount) unhealthy, $($ClientCounts.InactiveCount) inactive",
        "Devices:      $($InactiveCounts.InactiveCount) inactive (exceeding threshold)",
        "Site Health:  $($SiteCounts.OKCount) OK, $($SiteCounts.WarningCount) warning, $($SiteCounts.CriticalCount) critical"
    )

    return ($lines -join "`r`n")
}
