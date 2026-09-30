# Regression tests for Exchange organization hydration during Unified Audit Log remediation.
# New tenants can remain dehydrated briefly after Enable-OrganizationCustomization returns.

BeforeAll {
    $RepoRoot = Split-Path -Parent (Split-Path -Parent (Split-Path -Parent $PSCommandPath))

    . (Join-Path $RepoRoot 'Modules/CIPPStandards/Public/Standards/Invoke-CIPPStandardAuditLog.ps1')
    . (Join-Path $RepoRoot 'Modules/CIPPCore/Public/Baselines/Invoke-CIPPBaselineExoRequest.ps1')

    function Test-CIPPStandardLicense { param($StandardName, $TenantFilter, $Preset) $true }
    function New-ExoRequest {
        param($tenantid, $cmdlet, $cmdParams, $useSystemMailbox, $Select, [switch]$Compliance)
    }
    function Write-LogMessage { param($API, $tenant, $message, $sev, $LogData) }
    function Write-StandardsAlert { param($message, $object, $tenant, $standardName, $standardId) }
    function Set-CIPPStandardsCompareField { param($FieldName, $CurrentValue, $ExpectedValue, $TenantFilter) }
    function Get-NormalizedError { param($Message) "$Message" }
    function Get-CippException { param($Exception) [pscustomobject]@{ NormalizedError = "$($Exception.Message)" } }

    $script:Tenant = 'contoso.onmicrosoft.com'
}

Describe 'Invoke-CIPPStandardAuditLog organization customization wait' {
    BeforeEach {
        $script:OrgReads = 0
        $script:Logs = [System.Collections.Generic.List[object]]::new()
        Mock Start-Sleep {}
        Mock Write-LogMessage {
            param($API, $tenant, $message, $sev, $LogData)
            $script:Logs.Add([pscustomobject]@{ Message = $message; Sev = "$sev" })
        }
        Mock Write-StandardsAlert {}
        Mock Set-CIPPStandardsCompareField {}
    }

    It 'does not wait when the tenant is already hydrated' {
        Mock New-ExoRequest {
            param($tenantid, $cmdlet, $cmdParams, $useSystemMailbox, $Select)
            switch ($cmdlet) {
                'Get-AdminAuditLogConfig' { [pscustomobject]@{ UnifiedAuditLogIngestionEnabled = $false } }
                'Get-OrganizationConfig' { [pscustomobject]@{ IsDehydrated = $false } }
            }
        }

        Invoke-CIPPStandardAuditLog -Tenant $script:Tenant -Settings ([pscustomobject]@{
            remediate = $true
            alert = $false
            report = $false
            standardId = 'AuditLog'
        })

        Should -Invoke New-ExoRequest -Times 0 -ParameterFilter { $cmdlet -eq 'Enable-OrganizationCustomization' }
        Should -Invoke New-ExoRequest -Times 1 -Exactly -ParameterFilter { $cmdlet -eq 'Set-AdminAuditLogConfig' }
        Should -Invoke Start-Sleep -Times 0
    }

    It 'waits for a newly onboarded tenant to hydrate before enabling the audit log' {
        Mock New-ExoRequest {
            param($tenantid, $cmdlet, $cmdParams, $useSystemMailbox, $Select)
            switch ($cmdlet) {
                'Get-AdminAuditLogConfig' { return [pscustomobject]@{ UnifiedAuditLogIngestionEnabled = $false } }
                'Get-OrganizationConfig' {
                    $script:OrgReads++
                    # Initial read = dehydrated; first poll = still dehydrated; second poll = ready.
                    return [pscustomobject]@{ IsDehydrated = ($script:OrgReads -lt 3) }
                }
            }
        }

        Invoke-CIPPStandardAuditLog -Tenant $script:Tenant -Settings ([pscustomobject]@{
            remediate = $true
            alert = $false
            report = $false
            standardId = 'AuditLog'
        })

        Should -Invoke New-ExoRequest -Times 1 -Exactly -ParameterFilter { $cmdlet -eq 'Enable-OrganizationCustomization' }
        Should -Invoke New-ExoRequest -Times 1 -Exactly -ParameterFilter { $cmdlet -eq 'Set-AdminAuditLogConfig' }
        Should -Invoke Start-Sleep -Times 1 -Exactly -ParameterFilter { $Seconds -eq 5 }
    }

    It 'logs an error and does not run Set-AdminAuditLogConfig when hydration times out' {
        Mock New-ExoRequest {
            param($tenantid, $cmdlet, $cmdParams, $useSystemMailbox, $Select)
            switch ($cmdlet) {
                'Get-AdminAuditLogConfig' { [pscustomobject]@{ UnifiedAuditLogIngestionEnabled = $false } }
                'Get-OrganizationConfig' { [pscustomobject]@{ IsDehydrated = $true } }
            }
        }

        Invoke-CIPPStandardAuditLog -Tenant $script:Tenant -Settings ([pscustomobject]@{
            remediate = $true
            alert = $false
            report = $false
            standardId = 'AuditLog'
        })

        Should -Invoke New-ExoRequest -Times 1 -Exactly -ParameterFilter { $cmdlet -eq 'Enable-OrganizationCustomization' }
        Should -Invoke New-ExoRequest -Times 0 -ParameterFilter { $cmdlet -eq 'Set-AdminAuditLogConfig' }
        Should -Invoke Start-Sleep -Times 12 -Exactly -ParameterFilter { $Seconds -eq 5 }
        @($script:Logs | Where-Object {
            $_.Sev -eq 'Error' -and $_.Message -match 'still provisioning after 60 seconds'
        }).Count | Should -Be 1
    }
}

Describe 'Invoke-CIPPBaselineExoRequest organization customization wait' {
    BeforeEach {
        $script:OrgReads = 0
        Mock Start-Sleep {}
    }

    It 'waits for hydration before continuing to the next Exchange cmdlet' {
        Mock New-ExoRequest {
            param($tenantid, $cmdlet, $cmdParams, $useSystemMailbox, $Select, [switch]$Compliance)
            if ($cmdlet -eq 'Get-OrganizationConfig') {
                $script:OrgReads++
                return [pscustomobject]@{ IsDehydrated = ($script:OrgReads -lt 2) }
            }
        }

        $Remediate = [pscustomobject]@{
            cmdlets = @(
                [pscustomobject]@{ cmdlet = 'Enable-OrganizationCustomization'; params = [pscustomobject]@{}; continueOnError = $true }
                [pscustomobject]@{ cmdlet = 'Set-AdminAuditLogConfig'; params = [pscustomobject]@{ UnifiedAuditLogIngestionEnabled = $true } }
            )
        }

        Invoke-CIPPBaselineExoRequest -Remediate $Remediate -TenantFilter $script:Tenant

        Should -Invoke New-ExoRequest -Times 1 -Exactly -ParameterFilter { $cmdlet -eq 'Enable-OrganizationCustomization' }
        Should -Invoke New-ExoRequest -Times 1 -Exactly -ParameterFilter { $cmdlet -eq 'Set-AdminAuditLogConfig' }
        Should -Invoke Start-Sleep -Times 1 -Exactly -ParameterFilter { $Seconds -eq 5 }
    }

    It 'throws a clear error and stops the remediation when hydration times out' {
        Mock New-ExoRequest {
            param($tenantid, $cmdlet, $cmdParams, $useSystemMailbox, $Select, [switch]$Compliance)
            if ($cmdlet -eq 'Get-OrganizationConfig') {
                return [pscustomobject]@{ IsDehydrated = $true }
            }
        }

        $Remediate = [pscustomobject]@{
            cmdlets = @(
                [pscustomobject]@{ cmdlet = 'Enable-OrganizationCustomization'; params = [pscustomobject]@{}; continueOnError = $true }
                [pscustomobject]@{ cmdlet = 'Set-AdminAuditLogConfig'; params = [pscustomobject]@{ UnifiedAuditLogIngestionEnabled = $true } }
            )
        }

        { Invoke-CIPPBaselineExoRequest -Remediate $Remediate -TenantFilter $script:Tenant } |
            Should -Throw '*still provisioning*after 60 seconds*'

        Should -Invoke New-ExoRequest -Times 0 -ParameterFilter { $cmdlet -eq 'Set-AdminAuditLogConfig' }
        Should -Invoke Start-Sleep -Times 12 -Exactly -ParameterFilter { $Seconds -eq 5 }
    }
}
