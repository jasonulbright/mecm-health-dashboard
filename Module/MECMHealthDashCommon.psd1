@{
    RootModule        = 'MECMHealthDashCommon.psm1'
    ModuleVersion     = '1.0.0'
    GUID              = 'c3d4e5f6-a7b8-9012-cdef-345678901234'
    Author            = 'Jason Ulbright'
    Description       = 'MECM environment health dashboard - deployment, content, DP, client, and site health queries.'
    PowerShellVersion = '5.1'

    FunctionsToExport = @(
        # Logging
        'Initialize-Logging'
        'Write-Log'

        # CM Connection
        'Connect-CMSite'
        'Disconnect-CMSite'
        'Test-CMConnection'
        'Test-SQLConnection'

        # Deployment Health
        'Get-DeploymentHealth'
        'Get-DeploymentDetails'
        'Get-DeploymentHealthCounts'

        # Content Distribution Health
        'Get-ContentDistributionHealth'
        'Get-ContentHealthCounts'

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

        # Export
        'Export-HealthStatusCsv'
        'Export-HealthStatusHtml'
        'New-HealthSummaryText'
    )

    CmdletsToExport   = @()
    VariablesToExport  = @()
    AliasesToExport    = @()
}
