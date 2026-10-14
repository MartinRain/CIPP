function Set-CIPPDBCacheSharePointSiteUsage {
    <#
    .SYNOPSIS
        Caches SharePoint site listing and site usage details for a tenant

    .DESCRIPTION
        Active sites + usage from SPO admin RenderAdminListData (same source as the site browser),
        Graph getAllSites for Graph ids / sharepointIds, Get-CIPPSPOSite for file-level archive
        metrics, and a Graph lists bulk pass for AutoMapUrl. Writes the same SharePointSiteListing
        and SharePointSiteUsage property shapes consumed by Get-CIPPSharePointSiteUsageReport.

    .PARAMETER TenantFilter
        The tenant to cache SharePoint site usage for

    .PARAMETER QueueId
        The queue ID to update with total tasks (optional)
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$TenantFilter,
        [string]$QueueId
    )

    try {
        Write-LogMessage -API 'CIPPDBCache' -tenant $TenantFilter -message 'Caching SharePoint site listing and usage' -sev Debug

        $Built = Get-CIPPSharePointSiteUsageRows -TenantFilter $TenantFilter -IncludeArchive -LogApi 'CIPPDBCache'

        # Resolve inherited modes from the existing SPOTenant reporting cache, collected before
        # SharePointSiteUsage in the normal SharePoint collection. No extra SharePoint/Graph call.
        $GlobalAutoTrim = $null
        if ($Built.SiteListing | Where-Object { $_.versionPolicyInherited -eq $true } | Select-Object -First 1) {
            try {
                $GlobalRow = Get-CIPPDbItem -TenantFilter $TenantFilter -Type 'SPOTenant' -ErrorAction Stop |
                    Where-Object { $_.RowKey -ne 'SPOTenant-Count' } | Select-Object -First 1
                if ($GlobalRow -and $GlobalRow.Data) {
                    $GlobalSettings = $GlobalRow.Data | ConvertFrom-Json -Depth 10
                    if ($null -ne $GlobalSettings.EnableAutoExpirationVersionTrim) {
                        $GlobalAutoTrim = [System.Convert]::ToBoolean($GlobalSettings.EnableAutoExpirationVersionTrim)
                    }
                }
            } catch {
                Write-LogMessage -API 'CIPPDBCache' -tenant $TenantFilter -message "Could not read cached SPO tenant version policy: $($_.Exception.Message)" -sev Warning
            }
        }

        foreach ($Site in $Built.SiteListing) {
            $Inherited = $Site.versionPolicyInherited
            $EffectiveAutoTrim = if ($Inherited -eq $true) { $GlobalAutoTrim } else { $Site.versionPolicyAutoTrim }
            $Mode = if ($null -eq $Inherited -or $null -eq $EffectiveAutoTrim) {
                'Unknown'
            } elseif ($EffectiveAutoTrim) {
                'Automatic'
            } else {
                'Manual'
            }
            if ($Inherited -eq $true) {
                $Mode = if ($Mode -eq 'Unknown') { 'Inherited (Unknown)' } else { "$Mode (Inherited)" }
            }
            $Site | Add-Member -NotePropertyName 'versionPolicy' -NotePropertyValue $Mode -Force
        }

        Add-CIPPDbItem -TenantFilter $TenantFilter -Type 'SharePointSiteListing' -Data @($Built.SiteListing) -AddCount
        Add-CIPPDbItem -TenantFilter $TenantFilter -Type 'SharePointSiteUsage' -Data @($Built.UsageRows) -AddCount

        Write-LogMessage -API 'CIPPDBCache' -tenant $TenantFilter -message 'Cached SharePoint site listing and usage successfully' -sev Debug

    } catch {
        Write-LogMessage -API 'CIPPDBCache' -tenant $TenantFilter -message "Failed to cache SharePoint site usage: $($_.Exception.Message)" -sev Error -LogData (Get-CippException -Exception $_)
    }
}
