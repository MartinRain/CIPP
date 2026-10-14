function Resolve-CIPPIntuneMamAllUsersGroupId {
    <#
    .SYNOPSIS
        Resolves the tenant-specific Entra group Intune uses for MAM "All Users".

    .DESCRIPTION
        Direct App Protection / MAM assignments accept groupAssignmentTarget and
        exclusionGroupAssignmentTarget only. The Intune portal expresses "All users" by assigning
        the tenant's dynamic security group whose membershipRule is exactly "All Users".

        This resolver is intentionally narrow: it performs a directory lookup only for MAM policy
        types when the requested broad target contains All Users. Other policy types and target
        choices return $null without issuing a Graph request.

        The display name is not used because Microsoft localizes it. Resolution is based on the
        stable directory properties observed on the portal-created target.

    .FUNCTIONALITY
        Internal
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$TenantFilter,
        [Parameter(Mandatory = $true)]
        [string]$PolicyType,
        [Parameter(Mandatory = $true)]
        [string]$AssignTo
    )

    $TargetMetadata = Get-CIPPIntuneAssignmentTarget -AssignTo $AssignTo -PolicyType $PolicyType
    $NeedsAllUsers = $AssignTo -in @('allLicensedUsers', 'AllDevicesAndUsers')
    if (-not $TargetMetadata.IsMam -or -not $NeedsAllUsers) {
        return $null
    }

    $Uri = "https://graph.microsoft.com/v1.0/groups?`$filter=membershipRule eq 'All Users'&`$select=id,displayName,groupTypes,membershipRule,membershipRuleProcessingState,securityEnabled,mailEnabled&`$count=true"
    $Candidates = @(
        New-GraphGetRequest -uri $Uri -tenantid $TenantFilter -ComplexFilter |
            Where-Object {
                $_.membershipRule -eq 'All Users' -and
                $_.securityEnabled -eq $true -and
                $_.groupTypes -contains 'DynamicMembership'
            }
    )

    if ($Candidates.Count -eq 0) {
        throw "Could not resolve the tenant-specific All Users dynamic security group required for MAM assignments in tenant '$TenantFilter'."
    }

    if ($Candidates.Count -gt 1) {
        throw "Resolved multiple tenant-specific All Users dynamic security groups for MAM assignments in tenant '$TenantFilter'. Refusing to choose one automatically."
    }

    return [string]$Candidates[0].id
}
