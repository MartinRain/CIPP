function Invoke-ExecSPOTenantVersionPolicy {
    <#
    .FUNCTIONALITY
        Entrypoint
    .ROLE
        Sharepoint.Admin.ReadWrite
    .DESCRIPTION
        Updates only the SharePoint tenant's default file version policy, with concurrency checks.
        Does not propagate to existing libraries or launch version cleanup.
    #>
    [CmdletBinding()]
    param($Request, $TriggerMetadata)

    $APIName = $Request.Params.CIPPEndpoint
    $TenantFilter = $Request.Body.tenantFilter
    $InputPolicy = $Request.Body.Policy
    $ExpectedPolicy = $Request.Body.ExpectedPolicy

    $BadRequest = {
        param([string]$Message)
        return [HttpResponseContext]@{
            StatusCode = [HttpStatusCode]::BadRequest
            Body       = @{ Results = $Message }
        }
    }

    if ([string]::IsNullOrWhiteSpace([string]$TenantFilter) -or $TenantFilter -eq 'AllTenants') {
        return (& $BadRequest 'Select one tenant to modify its global SharePoint version policy.')
    }
    if ($null -eq $InputPolicy -or $null -eq $ExpectedPolicy -or
        $null -eq $InputPolicy.EnableAutoTrim -or $null -eq $ExpectedPolicy.EnableAutoTrim) {
        return (& $BadRequest 'Both Policy and ExpectedPolicy with an explicit EnableAutoTrim value are required.')
    }

    # Normalize numeric values without relying on PowerShell's permissive boolean conversion:
    # [bool]'false' is $true, so strings or other coercions must never be accepted as flags.
    $ReadPolicy = {
        param($Value, [bool]$ValidateDesired)
        $RawAuto = $Value.EnableAutoTrim
        if ($RawAuto -is [bool]) {
            $Auto = $RawAuto
        } elseif ($RawAuto -is [string] -and $RawAuto -in @('true', 'false')) {
            $Auto = [System.Convert]::ToBoolean($RawAuto)
        } else {
            throw 'EnableAutoTrim must be a boolean.'
        }

        $Major = $null
        $Days = $null
        if (-not $Auto) {
            foreach ($Field in @('MajorVersionLimit', 'ExpireVersionsAfterDays')) {
                $Raw = $Value.$Field
                if ($null -eq $Raw -or "$Raw" -notmatch '^\d+$') { throw "$Field must be a whole number." }
                $Parsed = 0
                if (-not [int]::TryParse("$Raw", [ref]$Parsed)) { throw "$Field is out of range." }
                if ($Field -eq 'MajorVersionLimit') { $Major = $Parsed } else { $Days = $Parsed }
            }
            if ($ValidateDesired) {
                if ($Major -lt 100 -or $Major -gt 50000) { throw 'MajorVersionLimit must be between 100 and 50000.' }
                if ($Days -ne 0 -and ($Days -lt 30 -or $Days -gt 36500)) {
                    throw 'ExpireVersionsAfterDays must be 0 or between 30 and 36500.'
                }
            }
        }
        return [PSCustomObject]@{
            EnableAutoTrim = $Auto
            MajorVersionLimit = $Major
            ExpireVersionsAfterDays = $Days
        }
    }

    try {
        $Desired = & $ReadPolicy $InputPolicy $true
        $Expected = & $ReadPolicy $ExpectedPolicy $false
    } catch {
        return (& $BadRequest $_.Exception.Message)
    }

    $Equals = {
        param($A, $B)
        if ([bool]$A.EnableAutoTrim -ne [bool]$B.EnableAutoTrim) { return $false }
        if ($A.EnableAutoTrim) { return $true }
        return ([int]$A.MajorVersionLimit -eq [int]$B.MajorVersionLimit -and
            [int]$A.ExpireVersionsAfterDays -eq [int]$B.ExpireVersionsAfterDays)
    }

    try {
        $Tenant = Get-CIPPSPOTenant -TenantFilter $TenantFilter -SkipCache -UseCertificate -ErrorAction Stop
        if (-not $Tenant -or $null -eq $Tenant.EnableAutoExpirationVersionTrim -or
            [string]::IsNullOrWhiteSpace([string]$Tenant._ObjectIdentity_)) {
            throw 'SharePoint did not return a usable tenant version policy or CSOM identity.'
        }

        $Current = [PSCustomObject]@{
            EnableAutoTrim = [System.Convert]::ToBoolean($Tenant.EnableAutoExpirationVersionTrim)
            MajorVersionLimit = if ($null -ne $Tenant.MajorVersionLimit) { [int]$Tenant.MajorVersionLimit } else { $null }
            ExpireVersionsAfterDays = if ($null -ne $Tenant.ExpireVersionsAfterDays) { [int]$Tenant.ExpireVersionsAfterDays } else { $null }
        }

        if (-not (& $Equals $Current $Expected)) {
            return [HttpResponseContext]@{
                StatusCode = [HttpStatusCode]::Conflict
                Body       = @{
                    Results = 'The tenant policy changed since the editor was opened. Refresh the policy and review it before saving.'
                    CurrentPolicy = $Current
                }
            }
        }

        if (& $Equals $Current $Desired) {
            return [HttpResponseContext]@{
                StatusCode = [HttpStatusCode]::OK
                Body       = @{ Results = 'The global SharePoint version policy is already configured as requested.'; Changed = $false; Policy = $Current }
            }
        }

        $Parameters = if ($Desired.EnableAutoTrim) {
            @(
                @{ Type = 'Boolean'; Value = $true }
                @{ Type = 'Int32'; Value = -1 }
                @{ Type = 'Int32'; Value = -1 }
            )
        } else {
            @(
                @{ Type = 'Boolean'; Value = $false }
                @{ Type = 'Int32'; Value = [int]$Desired.MajorVersionLimit }
                @{ Type = 'Int32'; Value = [int]$Desired.ExpireVersionsAfterDays }
            )
        }

        $null = $Tenant | Set-CIPPSPOTenant -MethodName 'SetFileVersionPolicy' -MethodParameters $Parameters -UseCertificate -ErrorAction Stop
        Write-LogMessage -Headers $Request.Headers -API $APIName -tenant $TenantFilter -message "Updated global SharePoint version policy: AutoTrim=$($Desired.EnableAutoTrim), MajorVersionLimit=$($Desired.MajorVersionLimit), ExpireVersionsAfterDays=$($Desired.ExpireVersionsAfterDays). Previous: AutoTrim=$($Current.EnableAutoTrim), MajorVersionLimit=$($Current.MajorVersionLimit), ExpireVersionsAfterDays=$($Current.ExpireVersionsAfterDays). No existing-library propagation or cleanup was requested." -sev Info

        # A SharePoint write may take time to become visible; never silently retry a write.
        try {
            $VerifiedTenant = Get-CIPPSPOTenant -TenantFilter $TenantFilter -SkipCache -UseCertificate -ErrorAction Stop
            $VerifiedPolicy = [PSCustomObject]@{
                EnableAutoTrim = [System.Convert]::ToBoolean($VerifiedTenant.EnableAutoExpirationVersionTrim)
                MajorVersionLimit = if ($null -ne $VerifiedTenant.MajorVersionLimit) { [int]$VerifiedTenant.MajorVersionLimit } else { $null }
                ExpireVersionsAfterDays = if ($null -ne $VerifiedTenant.ExpireVersionsAfterDays) { [int]$VerifiedTenant.ExpireVersionsAfterDays } else { $null }
            }
            $Verified = & $Equals $VerifiedPolicy $Desired
        } catch {
            $Verified = $false
            $VerifiedPolicy = $null
            Write-LogMessage -Headers $Request.Headers -API $APIName -tenant $TenantFilter -message "Global version policy write accepted, but verification failed: $($_.Exception.Message)" -sev Warning
        }

        $Message = if ($Verified) {
            'Global SharePoint version policy updated and verified.'
        } else {
            'SharePoint accepted the policy update, but the new values could not yet be verified. Refresh the policy before retrying.'
        }
        return [HttpResponseContext]@{
            StatusCode = [HttpStatusCode]::OK
            Body       = @{ Results = $Message; Changed = $true; Verified = $Verified; Policy = $VerifiedPolicy }
        }
    } catch {
        $ErrorMessage = Get-CippException -Exception $_
        Write-LogMessage -Headers $Request.Headers -API $APIName -tenant $TenantFilter -message "Failed to update global SharePoint version policy: $($ErrorMessage.NormalizedError)" -sev Error -LogData $ErrorMessage
        $Code = if ($_.Exception.Data['SPOAccessDenied']) { [HttpStatusCode]::Forbidden } else { [HttpStatusCode]::InternalServerError }
        return [HttpResponseContext]@{
            StatusCode = $Code
            Body       = @{ Results = "Failed to update global SharePoint version policy: $($ErrorMessage.NormalizedError)" }
        }
    }
}
