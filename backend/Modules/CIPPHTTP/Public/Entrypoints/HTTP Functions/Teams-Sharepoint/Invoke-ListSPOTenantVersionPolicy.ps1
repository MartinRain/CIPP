function Invoke-ListSPOTenantVersionPolicy {
    <#
    .FUNCTIONALITY
        Entrypoint
    .ROLE
        Sharepoint.Admin.Read
    .DESCRIPTION
        Reads the SharePoint tenant's default file version policy. No site or library settings are changed.
    #>
    [CmdletBinding()]
    param($Request, $TriggerMetadata)

    $APIName = $Request.Params.CIPPEndpoint
    $TenantFilter = $Request.Query.tenantFilter

    if ([string]::IsNullOrWhiteSpace([string]$TenantFilter) -or $TenantFilter -eq 'AllTenants') {
        return [HttpResponseContext]@{
            StatusCode = [HttpStatusCode]::BadRequest
            Body       = @{ Results = 'Select one tenant to view its global SharePoint version policy.' }
        }
    }

    try {
        # The modal is an administrative editor: never prefill it with a potentially stale 1-hour cache.
        $Tenant = Get-CIPPSPOTenant -TenantFilter $TenantFilter -SkipCache -UseCertificate -ErrorAction Stop
        if (-not $Tenant -or $null -eq $Tenant.EnableAutoExpirationVersionTrim) {
            throw 'SharePoint did not return a valid tenant version policy.'
        }

        $Policy = [PSCustomObject]@{
            EnableAutoTrim         = [System.Convert]::ToBoolean($Tenant.EnableAutoExpirationVersionTrim)
            MajorVersionLimit      = if ($null -ne $Tenant.MajorVersionLimit) { [int]$Tenant.MajorVersionLimit } else { $null }
            ExpireVersionsAfterDays = if ($null -ne $Tenant.ExpireVersionsAfterDays) { [int]$Tenant.ExpireVersionsAfterDays } else { $null }
        }

        return [HttpResponseContext]@{
            StatusCode = [HttpStatusCode]::OK
            Body       = @{ Results = $Policy }
        }
    } catch {
        $ErrorMessage = Get-CippException -Exception $_
        Write-LogMessage -Headers $Request.Headers -API $APIName -tenant $TenantFilter -message "Failed to read the global SharePoint version policy: $($ErrorMessage.NormalizedError)" -sev Error -LogData $ErrorMessage
        $Code = if ($_.Exception.Data['SPOAccessDenied']) { [HttpStatusCode]::Forbidden } else { [HttpStatusCode]::InternalServerError }
        return [HttpResponseContext]@{
            StatusCode = $Code
            Body       = @{ Results = "Failed to read the global SharePoint version policy: $($ErrorMessage.NormalizedError)" }
        }
    }
}
