# Pester tests for tenant-specific MAM All Users group resolution.

BeforeAll {
    $RepoRoot = Split-Path -Parent (Split-Path -Parent (Split-Path -Parent $PSCommandPath))
    function New-GraphGetRequest { [CmdletBinding()] param($uri, $tenantid, [switch]$ComplexFilter) }
    . (Join-Path $RepoRoot 'Modules/CIPPCore/Public/Get-CIPPIntuneAssignmentTarget.ps1')
    . (Join-Path $RepoRoot 'Modules/CIPPCore/Public/Resolve-CIPPIntuneMamAllUsersGroupId.ps1')
    $script:Tenant = 'contoso.onmicrosoft.com'
    $script:MamAllUsersId = '55555555-5555-5555-5555-555555555555'
}

Describe 'Resolve-CIPPIntuneMamAllUsersGroupId' {
    BeforeEach {
        Mock New-GraphGetRequest {
            @([PSCustomObject]@{
                id = $script:MamAllUsersId
                displayName = 'Tous les utilisateurs'
                groupTypes = @('DynamicMembership')
                membershipRule = 'All Users'
                membershipRuleProcessingState = 'On'
                securityEnabled = $true
                mailEnabled = $false
            })
        }
    }

    It 'resolves by membershipRule rather than localized display name' {
        $Result = Resolve-CIPPIntuneMamAllUsersGroupId -TenantFilter $script:Tenant -PolicyType 'iosManagedAppProtections' -AssignTo 'allLicensedUsers'
        $Result | Should -Be $script:MamAllUsersId
        Should -Invoke New-GraphGetRequest -Times 1 -Exactly -ParameterFilter {
            $uri -like '*membershipRule eq ''All Users''*' -and $ComplexFilter
        }
    }

    It 'does not query Graph for non-MAM policy types' {
        $Result = Resolve-CIPPIntuneMamAllUsersGroupId -TenantFilter $script:Tenant -PolicyType 'deviceConfigurations' -AssignTo 'allLicensedUsers'
        $Result | Should -BeNullOrEmpty
        Should -Invoke New-GraphGetRequest -Times 0 -Exactly
    }

    It 'does not query Graph for the unsupported All Devices MAM target' {
        $Result = Resolve-CIPPIntuneMamAllUsersGroupId -TenantFilter $script:Tenant -PolicyType 'iosManagedAppProtections' -AssignTo 'AllDevices'
        $Result | Should -BeNullOrEmpty
        Should -Invoke New-GraphGetRequest -Times 0 -Exactly
    }

    It 'fails explicitly when no matching dynamic security group exists' {
        Mock New-GraphGetRequest { @() }
        { Resolve-CIPPIntuneMamAllUsersGroupId -TenantFilter $script:Tenant -PolicyType 'iosManagedAppProtections' -AssignTo 'allLicensedUsers' } |
            Should -Throw -ExpectedMessage '*Could not resolve*All Users*'
    }

    It 'fails explicitly when multiple matching groups exist' {
        Mock New-GraphGetRequest {
            @(
                [PSCustomObject]@{ id='1'; groupTypes=@('DynamicMembership'); membershipRule='All Users'; securityEnabled=$true },
                [PSCustomObject]@{ id='2'; groupTypes=@('DynamicMembership'); membershipRule='All Users'; securityEnabled=$true }
            )
        }
        { Resolve-CIPPIntuneMamAllUsersGroupId -TenantFilter $script:Tenant -PolicyType 'iosManagedAppProtections' -AssignTo 'allLicensedUsers' } |
            Should -Throw -ExpectedMessage '*multiple*All Users*'
    }

    It 'ignores a normal group merely named All Users' {
        Mock New-GraphGetRequest {
            @([PSCustomObject]@{ id='fake'; displayName='All Users'; groupTypes=@(); membershipRule=$null; securityEnabled=$true })
        }
        { Resolve-CIPPIntuneMamAllUsersGroupId -TenantFilter $script:Tenant -PolicyType 'iosManagedAppProtections' -AssignTo 'allLicensedUsers' } |
            Should -Throw -ExpectedMessage '*Could not resolve*'
    }
}
