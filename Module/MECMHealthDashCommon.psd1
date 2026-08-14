@{
    RootModule        = 'MECMHealthDashCommon.psm1'
    ModuleVersion     = '1.2.0'
    GUID              = '8d2a6f4e-1c7b-4e93-a5d8-0b6e3f9c2a17'
    Author            = 'Jason Ulbright'
    Description       = 'MECM environment health dashboard - deployment, content, DP, client, and site health queries.'
    PowerShellVersion = '5.1'

    FunctionsToExport = @(
        # Logging and CM connection come from the vendored SuiteCommon
        # module (Lib\SuiteCommon), imported globally by the root module.
        'Test-SQLConnection'

        # Deployment Health
        'Get-DeploymentHealth'
        'Get-DeploymentDetails'
        'Get-DeploymentHealthCounts'

        # Content Distribution Health
        'Get-ContentDistributionHealth'
        'Get-ContentHealthCounts'
        'Get-ContentNameMap'

        # Distribution Point Health
        'Get-DPHealth'
        'Get-DPDetails'
        'Get-DPHealthCounts'

        # Client Health (SQL)
        'Get-ClientHealthSummary'
        'Get-ClientHealthCounts'

        # Inactive Devices (SQL)
        'Get-InactiveDevices'
        'Get-InactiveDeviceCounts'

        # Site Health (WMI)
        'Get-SiteComponentHealth'
        'Get-SiteSystemHealth'
        'Get-SiteHealthCounts'

        # Metrics history (Trends view)
        'Add-MetricsHistoryEntry'
        'Get-MetricsHistory'

        # Export
        'Export-HealthStatusCsv'
        'Export-HealthStatusHtml'
        'New-HealthSummaryText'
    )

    CmdletsToExport   = @()
    VariablesToExport  = @()
    AliasesToExport    = @()
}
