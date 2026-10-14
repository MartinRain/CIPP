function Get-CIPPIntuneAssignmentTarget {
    <#
    .SYNOPSIS
        The include targets one of the broad AssignTo options produces for a given policy type.

    .DESCRIPTION
        Remediation (Set-CIPPAssignedPolicy) and the assignment comparison
        (Compare-CIPPIntuneAssignments) both read the broad targets through this function, so what
        the comparison asserts is what remediation writes. A comparison expecting a target
        remediation never produces is a deviation no run can clear - the same invariant
        Get-CIPPIntuneAssignTarget exists to protect, one level further down.

        App Protection and managed app configuration are served by Intune's MAM service rather than
        the device management service, and it accepts only groupAssignmentTarget and
        exclusionGroupAssignmentTarget - an assign carrying allLicensedUsersAssignmentTarget or
        allDevicesAssignmentTarget is rejected outright. "All users" is expressed there as a group
        assignment. For direct MAM assignments the portal targets the tenant-specific Entra dynamic
        security group whose membershipRule is "All Users"; callers pass that resolved group id into
        this helper. A MAM policy protects a user's apps and has no device audience at
        all, so All Devices has no equivalent and is reported as unsupported rather than guessed at.

        Device Preparation profiles (Autopilot device preparation) have the same shape for a
        different reason: the deployment starts when an assigned user signs in during OOBE, so the
        assignment surface is user groups only - the portal picker offers the same well-known All
        Users virtual group and no All Devices equivalent. The broad virtual targets are not what
        the portal writes for them and leave the profile without an effective assignment.

        'customGroup' and 'On' produce no broad target: the caller resolves group names itself.

    .PARAMETER AssignTo
        The canonical AssignTo value from Get-CIPPIntuneAssignTarget.

    .PARAMETER PolicyType
        The policy's type. Accepts either the CIPP template type ('AppProtection') or the Graph
        collection segment ('iosManagedAppProtections'), because the two callers speak in different
        vocabularies.

    .OUTPUTS
        PSCustomObject with:
            Targets     - the target objects to assign, in order
            Unsupported - a message when the requested target cannot be expressed for this policy
                          type, otherwise $null
            Dropped     - targets that were requested but cannot be expressed, where the rest of
                          the request still can
            GroupNames  - display names for any virtual group a broad target resolved to, keyed by
                          id, so the comparison can name it instead of printing a bare GUID
            Equivalents - assignment types that express the same audience as a broad group target,
                          keyed by that target's group id. Remediation keeps writing the group
                          target; the comparison accepts either shape.

    .EXAMPLE
        Get-CIPPIntuneAssignmentTarget -AssignTo 'allLicensedUsers' -PolicyType 'iosManagedAppProtections'

    .FUNCTIONALITY
        Internal
    #>
    [CmdletBinding()]
    param(
        [string]$AssignTo,
        [string]$PolicyType,
        [string]$MamAllUsersGroupId
    )

    # Device Preparation and Apple enrollment still use Intune's well-known All Users virtual
    # group. Direct MAM/App Protection assignments are different: callers resolve the tenant's
    # dynamic All Users security group and pass its id through MamAllUsersGroupId.
    $VirtualAllUsersGroupId = 'acacacac-9df4-4c7d-9d50-4ef0226f57a9'

    $MamPolicyTypes = @(
        'AppProtection'
        'managedAppPolicies'
        'iosManagedAppProtections'
        'androidManagedAppProtections'
        'windowsManagedAppProtections'
        'targetedManagedAppConfigurations'
    )
    $IsMam = $MamPolicyTypes -contains $PolicyType

    # Policy types whose assignment surface is user groups only. Device Preparation deployments
    # trigger on the enrolling user, so a device audience cannot be expressed for them at all.
    # Apple enrollment type profiles apply to the enrolling user the same way - the portal's
    # picker offers user groups and nothing else.
    $UserGroupOnlyTypes = @(
        'DevicePrepProfile'
        'AppleEnrollmentTypeProfile'
    )
    $IsUserGroupOnly = $IsMam -or ($UserGroupOnlyTypes -contains $PolicyType)

    $EffectiveAllUsersGroupId = if ($IsMam -and $MamAllUsersGroupId) {
        $MamAllUsersGroupId
    } else {
        $VirtualAllUsersGroupId
    }

    $AllUsersTarget = if ($IsUserGroupOnly) {
        @{ '@odata.type' = '#microsoft.graph.groupAssignmentTarget'; groupId = $EffectiveAllUsersGroupId }
    } else {
        @{ '@odata.type' = '#microsoft.graph.allLicensedUsersAssignmentTarget' }
    }
    $AllDevicesTarget = @{ '@odata.type' = '#microsoft.graph.allDevicesAssignmentTarget' }

    $Targets = [System.Collections.Generic.List[object]]::new()
    $Dropped = [System.Collections.Generic.List[string]]::new()
    $Unsupported = $null

    switch ($AssignTo) {
        'allLicensedUsers' {
            $Targets.Add($AllUsersTarget)
        }
        'AllDevices' {
            if ($IsMam) {
                $Unsupported = "'$PolicyType' policies protect a user's apps and have no device audience, so they cannot be assigned to All Devices. Assign to all users or to a group instead."
            } elseif ($IsUserGroupOnly) {
                $Unsupported = "'$PolicyType' deployments start when an assigned user signs in, so they cannot be assigned to All Devices. Assign to all users instead."
            } else {
                $Targets.Add($AllDevicesTarget)
            }
        }
        'AllDevicesAndUsers' {
            if ($IsUserGroupOnly) {
                # The users half is still expressible, so honour it rather than failing the whole
                # assignment over a target this policy type has no concept of.
                $Dropped.Add('All Devices')
            } else {
                $Targets.Add($AllDevicesTarget)
            }
            $Targets.Add($AllUsersTarget)
        }
    }

    # Give broad group targets a stable friendly name for drift output. MAM may point at a real,
    # localized Entra group, so do not derive this label from the directory displayName.
    $GroupNames = @{}
    if ($Targets.groupId -contains $EffectiveAllUsersGroupId) {
        $GroupNames[$EffectiveAllUsersGroupId] = 'All Users'
    }

    # Outside MAM, the device management service reports an All Users group assignment back as
    # allLicensedUsersAssignmentTarget, so both shapes describe the same working assignment and the
    # comparison must accept either. The MAM service rejects that type outright, so it stays out.
    $Equivalents = @{}
    if (-not $IsMam -and $Targets.groupId -contains $VirtualAllUsersGroupId) {
        $Equivalents[$VirtualAllUsersGroupId] = @('#microsoft.graph.allLicensedUsersAssignmentTarget')
    }

    [PSCustomObject]@{
        Targets     = @($Targets)
        Unsupported = $Unsupported
        Dropped     = @($Dropped)
        GroupNames  = $GroupNames
        Equivalents = $Equivalents
        IsMam       = $IsMam
    }
}
