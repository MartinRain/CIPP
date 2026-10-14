BeforeAll {
    $RepoRoot = Split-Path -Parent (Split-Path -Parent (Split-Path -Parent $PSCommandPath))

    class HttpResponseContext { [int]$StatusCode; [object]$Body }
    $TypeAccelerators = [PowerShell].Assembly.GetType('System.Management.Automation.TypeAccelerators')
    if (-not ([System.Management.Automation.PSTypeName]'HttpStatusCode').Type) {
        $TypeAccelerators::Add('HttpStatusCode', [System.Net.HttpStatusCode])
    }

    function Get-CIPPSPOTenant {
        [CmdletBinding()]
        param($TenantFilter, [switch]$SkipCache, [switch]$UseCertificate)
    }
    function Set-CIPPSPOTenant {
        [CmdletBinding()]
        param(
            [Parameter(ValueFromPipeline = $true)]$InputObject,
            $MethodName,
            $MethodParameters,
            [switch]$UseCertificate
        )
        process {}
    }
    function Get-CippException {
        param($Exception)
        [pscustomobject]@{ NormalizedError = $Exception.Exception.Message }
    }
    function Write-LogMessage {
        param($Headers, $API, $tenant, $message, $sev, $LogData)
    }

    . (Join-Path $RepoRoot 'Modules/CIPPHTTP/Public/Entrypoints/HTTP Functions/Teams-Sharepoint/Invoke-ListSPOTenantVersionPolicy.ps1')
    . (Join-Path $RepoRoot 'Modules/CIPPHTTP/Public/Entrypoints/HTTP Functions/Teams-Sharepoint/Invoke-ExecSPOTenantVersionPolicy.ps1')

    function New-VersionPolicyRequest {
        param(
            $Desired = @{ EnableAutoTrim = $true },
            $Expected = @{ EnableAutoTrim = $false; MajorVersionLimit = 500; ExpireVersionsAfterDays = 0 },
            $Tenant = 'contoso.onmicrosoft.com'
        )
        return [pscustomobject]@{
            Params = @{ CIPPEndpoint = 'ExecSPOTenantVersionPolicy' }
            Headers = @{}
            Query = @{ tenantFilter = $Tenant }
            Body = [pscustomobject]@{
                tenantFilter = $Tenant
                Policy = [pscustomobject]$Desired
                ExpectedPolicy = [pscustomobject]$Expected
            }
        }
    }

    function New-VersionPolicyTenant {
        param([bool]$Auto = $false, [int]$Major = 500, [int]$Days = 0)
        return [pscustomobject]@{
            _ObjectIdentity_ = 'CSOM-id'
            TenantFilter = 'contoso.onmicrosoft.com'
            EnableAutoExpirationVersionTrim = $Auto
            MajorVersionLimit = $Major
            ExpireVersionsAfterDays = $Days
        }
    }
}

Describe 'Tenant global SharePoint version policy endpoints' {
    BeforeEach {
        $script:TenantRead = New-VersionPolicyTenant
        Mock Get-CIPPSPOTenant { $script:TenantRead }
        Mock Set-CIPPSPOTenant { }
        Mock Write-LogMessage { }
    }

    It 'GET returns only normalized policy fields and uses a fresh certificate read' {
        $Response = Invoke-ListSPOTenantVersionPolicy -Request (New-VersionPolicyRequest) -TriggerMetadata $null
        $Response.StatusCode | Should -Be 200
        $Response.Body.Results.EnableAutoTrim | Should -BeFalse
        $Response.Body.Results.MajorVersionLimit | Should -Be 500
        Should -Invoke Get-CIPPSPOTenant -Times 1 -ParameterFilter { $SkipCache -and $UseCertificate }
    }

    It 'GET refuses AllTenants without a SharePoint read' {
        $Response = Invoke-ListSPOTenantVersionPolicy -Request (New-VersionPolicyRequest -Tenant 'AllTenants')
        $Response.StatusCode | Should -Be 400
        Should -Invoke Get-CIPPSPOTenant -Times 0
    }

    It 'POST refuses AllTenants without a SharePoint write' {
        $Response = Invoke-ExecSPOTenantVersionPolicy -Request (New-VersionPolicyRequest -Tenant 'AllTenants')
        $Response.StatusCode | Should -Be 400
        Should -Invoke Get-CIPPSPOTenant -Times 0
        Should -Invoke Set-CIPPSPOTenant -Times 0
    }

    It 'POST rejects invalid manual limits before any SharePoint read' {
        $Response = Invoke-ExecSPOTenantVersionPolicy -Request (New-VersionPolicyRequest -Desired @{
            EnableAutoTrim = $false; MajorVersionLimit = 50; ExpireVersionsAfterDays = 10
        })
        $Response.StatusCode | Should -Be 400
        Should -Invoke Get-CIPPSPOTenant -Times 0
        Should -Invoke Set-CIPPSPOTenant -Times 0
    }

    It 'POST rejects false strings that would become true with a PowerShell bool cast' {
        $Response = Invoke-ExecSPOTenantVersionPolicy -Request (New-VersionPolicyRequest -Desired @{
            EnableAutoTrim = 'not-a-boolean'
        })
        $Response.StatusCode | Should -Be 400
        Should -Invoke Set-CIPPSPOTenant -Times 0
    }

    It 'POST refuses a stale expected policy' {
        $Response = Invoke-ExecSPOTenantVersionPolicy -Request (New-VersionPolicyRequest -Expected @{
            EnableAutoTrim = $false; MajorVersionLimit = 300; ExpireVersionsAfterDays = 0
        })
        $Response.StatusCode | Should -Be 409
        $Response.Body.CurrentPolicy.MajorVersionLimit | Should -Be 500
        Should -Invoke Set-CIPPSPOTenant -Times 0
    }

    It 'POST skips the write when the policy already matches' {
        $Response = Invoke-ExecSPOTenantVersionPolicy -Request (New-VersionPolicyRequest -Desired @{
            EnableAutoTrim = $false; MajorVersionLimit = 500; ExpireVersionsAfterDays = 0
        })
        $Response.StatusCode | Should -Be 200
        $Response.Body.Changed | Should -BeFalse
        Should -Invoke Set-CIPPSPOTenant -Times 0
    }

    It 'POST switches to automatic with the standard CSOM sentinels and verifies' {
        $script:ReadCount = 0
        Mock Get-CIPPSPOTenant {
            $script:ReadCount++
            if ($script:ReadCount -eq 1) { New-VersionPolicyTenant }
            else { New-VersionPolicyTenant -Auto $true -Major -1 -Days -1 }
        }
        $Response = Invoke-ExecSPOTenantVersionPolicy -Request (New-VersionPolicyRequest)
        $Response.StatusCode | Should -Be 200
        $Response.Body.Verified | Should -BeTrue
        Should -Invoke Set-CIPPSPOTenant -Times 1 -ParameterFilter {
            $MethodName -eq 'SetFileVersionPolicy' -and $UseCertificate -and
            $MethodParameters[0].Value -eq $true -and
            $MethodParameters[1].Value -eq -1 -and $MethodParameters[2].Value -eq -1
        }
        Should -Invoke Get-CIPPSPOTenant -Times 2 -ParameterFilter { $SkipCache -and $UseCertificate }
    }

    It 'POST applies the explicit manual policy without changing existing sites' {
        $script:ReadCount = 0
        Mock Get-CIPPSPOTenant {
            $script:ReadCount++
            if ($script:ReadCount -eq 1) { New-VersionPolicyTenant -Auto $true -Major -1 -Days -1 }
            else { New-VersionPolicyTenant -Auto $false -Major 1000 -Days 365 }
        }
        $Response = Invoke-ExecSPOTenantVersionPolicy -Request (New-VersionPolicyRequest -Expected @{
            EnableAutoTrim = $true
        } -Desired @{
            EnableAutoTrim = $false; MajorVersionLimit = 1000; ExpireVersionsAfterDays = 365
        })
        $Response.StatusCode | Should -Be 200
        $Response.Body.Verified | Should -BeTrue
        Should -Invoke Set-CIPPSPOTenant -Times 1 -ParameterFilter {
            $MethodParameters[0].Value -eq $false -and
            $MethodParameters[1].Value -eq 1000 -and
            $MethodParameters[2].Value -eq 365
        }
    }

    It 'POST does not retry when verification is delayed' {
        Mock Get-CIPPSPOTenant { New-VersionPolicyTenant }
        $Response = Invoke-ExecSPOTenantVersionPolicy -Request (New-VersionPolicyRequest)
        $Response.StatusCode | Should -Be 200
        $Response.Body.Changed | Should -BeTrue
        $Response.Body.Verified | Should -BeFalse
        Should -Invoke Set-CIPPSPOTenant -Times 1
    }

    It 'POST surfaces a SharePoint write exception instead of claiming success' {
        Mock Set-CIPPSPOTenant { throw 'CSOM write rejected' }
        $Response = Invoke-ExecSPOTenantVersionPolicy -Request (New-VersionPolicyRequest)
        $Response.StatusCode | Should -Be 500
        $Response.Body.Results | Should -Match 'CSOM write rejected'
        Should -Invoke Set-CIPPSPOTenant -Times 1
    }
}
