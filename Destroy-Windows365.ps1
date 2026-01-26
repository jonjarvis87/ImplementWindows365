<#
.SYNOPSIS
    Removes Windows 365 artifacts created by Deploy-Windows365.ps1 (groups, user settings, provisioning policies).

.DESCRIPTION
    - Connects to Microsoft Graph (beta profile) with required scopes.
    - Unassigns Cloud PC provisioning policies, then deletes them.
    - Unassigns Cloud PC user settings (W365_AdminSettings, W365_UserSettings, AI_Enabled_Cloud_PC), then deletes them.
    - Deletes Entra ID security groups created by the deployment script:
        * SG-W365CloudPC-* (licensing)
        * SG-W365ENT-*-User/Admin
        * SG-W365FL-*-User/Admin

.NOTES
    Destructive operation. Ensure you target the correct tenant. No soft-delete for user settings or policies.
#>

[CmdletBinding()]
param(
    [switch]$RemoveProvisioningPolicies = $true,
    [switch]$RemoveUserSettings = $true,
    [switch]$RemoveGroups = $true,

    [string[]]$KeepPolicies = @(),
    [string[]]$KeepUserSettings = @(),
    [string[]]$KeepGroups = @()
)

function Install-GraphModuleIfNeeded {
    if (-not (Get-Module -ListAvailable -Name Microsoft.Graph)) {
        Write-Host "Installing Microsoft.Graph..." -ForegroundColor Yellow
        Install-Module -Name Microsoft.Graph -AllowClobber -Force -ErrorAction Stop
    }
    if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Beta.DeviceManagement.Administration)) {
        Write-Host "Installing Microsoft.Graph.Beta.DeviceManagement.Administration..." -ForegroundColor Yellow
        Install-Module -Name Microsoft.Graph.Beta.DeviceManagement.Administration -MinimumVersion 2.23.0 -AllowClobber -Force -ErrorAction Stop
    }
}

function Ensure-BetaProfile {
    $selectProfileCmd = Get-Command Select-MgProfile -ErrorAction SilentlyContinue
    if ($selectProfileCmd) {
        try { Select-MgProfile -Name beta -ErrorAction Stop } catch { Write-Warning "Could not select beta profile: $_" }
    }
}

function Invoke-GraphAssignClear {
    param(
        [Parameter(Mandatory)] [string]$Uri
    )
    $payload = @{ assignments = @() } | ConvertTo-Json -Depth 3
    Invoke-MgGraphRequest -Method POST -Uri $Uri -Body $payload -ContentType "application/json" -ErrorAction Stop | Out-Null
}

function Remove-ProvisioningPolicies {
    Write-Host "Removing provisioning policies..." -ForegroundColor Cyan
    $policies = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/beta/deviceManagement/virtualEndpoint/provisioningPolicies" -ErrorAction Stop
    $targets = @($policies.value | Where-Object { $_.displayName -like '*-W365-*-Provisioning Policy' })

    if ($KeepPolicies.Count -gt 0) {
        $targets = $targets | Where-Object { $_.displayName -notin $KeepPolicies }
    }

    foreach ($p in $targets) {
        Write-Host "Processing policy: $($p.displayName)" -ForegroundColor Yellow
        try {
            $assignUri = "https://graph.microsoft.com/beta/deviceManagement/virtualEndpoint/provisioningPolicies/$($p.id)/assign"
            Invoke-GraphAssignClear -Uri $assignUri
            Write-Host "  Assignments cleared" -ForegroundColor Green
        }
        catch { Write-Warning "  Failed to clear assignments: $_" }

        try {
            Invoke-MgGraphRequest -Method DELETE -Uri "https://graph.microsoft.com/beta/deviceManagement/virtualEndpoint/provisioningPolicies/$($p.id)" -ErrorAction Stop
            Write-Host "  Deleted" -ForegroundColor Green
        }
        catch { Write-Warning "  Failed to delete: $_" }
    }
}

function Remove-UserSettings {
    Write-Host "Removing user settings..." -ForegroundColor Cyan
    $settings = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/beta/deviceManagement/virtualEndpoint/userSettings" -ErrorAction Stop
    $targets = @($settings.value | Where-Object { $_.displayName -in @('W365_AdminSettings','W365_UserSettings','AI_Enabled_Cloud_PC') })

    if ($KeepUserSettings.Count -gt 0) {
        $targets = $targets | Where-Object { $_.displayName -notin $KeepUserSettings }
    }

    foreach ($s in $targets) {
        Write-Host "Processing user setting: $($s.displayName)" -ForegroundColor Yellow
        try {
            $assignUri = "https://graph.microsoft.com/beta/deviceManagement/virtualEndpoint/userSettings/$($s.id)/assign"
            Invoke-GraphAssignClear -Uri $assignUri
            Write-Host "  Assignments cleared" -ForegroundColor Green
        }
        catch { Write-Warning "  Failed to clear assignments: $_" }

        try {
            Invoke-MgGraphRequest -Method DELETE -Uri "https://graph.microsoft.com/beta/deviceManagement/virtualEndpoint/userSettings/$($s.id)" -ErrorAction Stop
            Write-Host "  Deleted" -ForegroundColor Green
        }
        catch { Write-Warning "  Failed to delete: $_" }
    }
}

function Remove-Groups {
    Write-Host "Removing Entra ID groups..." -ForegroundColor Cyan
    $filters = @(
        "startswith(displayName,'SG-W365CloudPC-')",
        "startswith(displayName,'SG-W365ENT-')",
        "startswith(displayName,'SG-W365FL-')"
    )

    foreach ($f in $filters) {
        try {
            $resp = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/groups?`$filter=$f&`$top=999" -ErrorAction Stop
            foreach ($g in $resp.value) {
                if ($KeepGroups -and ($g.displayName -in $KeepGroups)) {
                    Write-Host "Skipping kept group: $($g.displayName)" -ForegroundColor Yellow
                    continue
                }
                Write-Host "Deleting group: $($g.displayName)" -ForegroundColor Yellow
                try {
                    Invoke-MgGraphRequest -Method DELETE -Uri "https://graph.microsoft.com/v1.0/groups/$($g.id)" -ErrorAction Stop
                    Write-Host "  Deleted" -ForegroundColor Green
                }
                catch { Write-Warning "  Failed to delete group $($g.displayName): $_" }
            }
        }
        catch { Write-Warning "Failed to query groups with filter $f : $_" }
    }
}

Install-GraphModuleIfNeeded
Ensure-BetaProfile

Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan
Connect-MgGraph -Scopes "User.ReadWrite.All","Application.ReadWrite.All","CloudPC.ReadWrite.All","Group.ReadWrite.All" -ErrorAction Stop
Write-Host "Connected." -ForegroundColor Green

if ($RemoveProvisioningPolicies) { Remove-ProvisioningPolicies }
if ($RemoveUserSettings) { Remove-UserSettings }
if ($RemoveGroups) { Remove-Groups }

Write-Host "Cleanup complete." -ForegroundColor Green