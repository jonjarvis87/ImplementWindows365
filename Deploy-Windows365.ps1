<#
.SYNOPSIS
    Create a Windows 365 environment (Groups, User Settings, Provisioning Policy)
.DESCRIPTION
    Prompts for a Cloud PC SKU, then:
    - Ensures Microsoft.Graph.Authentication is installed (lightweight module)
    - Connects to Microsoft Graph via REST API
    - Creates (or reuses) Entra ID security groups with customizable naming
    - Creates (or reuses) Cloud PC User Settings policies + assigns them
    - Creates (or reuses) a Cloud PC Provisioning Policy + assigns it
    - Preserves existing assignments by merging current + new groups (Graph /assign is replace-all)
.PARAMETER CloudPCTypeChoice
    The Cloud PC SKU choice (dynamic from tenant). If not provided, will prompt interactively.
.PARAMETER RegionChoice
    The region choice (dynamic from tenant). Regions are always prompted for selection.
.PARAMETER Language
    Optional language code for Windows 11 provisioning policy (default: en-GB). If not provided, will prompt via interactive grid.
.PARAMETER GroupPrefix
    Customize the security group prefix (default: "SG-W365"). Use organizational naming standards if needed.
    Examples: "SG-W365", "ACME-W365", "IT-CloudPC"
.PARAMETER ProvisioningPolicySuffix
    Customize the provisioning policy suffix (default: "Provisioning Policy"). Use organizational naming standards if needed.
    Examples: "Provisioning Policy", "Policy", "Config"
    Final format: <Region>-W365-<LicenseType>-<ProvisioningPolicySuffix>
.NOTES
    Script name: Deploy-Windows365.ps1
    Author:      Jon Jarvis
    Required scopes: User.ReadWrite.All, Application.ReadWrite.All, CloudPC.ReadWrite.All, Group.ReadWrite.All
    Requires:    PowerShell 7.0 or higher
    Naming Convention Best Practices:
    - Groups: <GroupPrefix>-<LicenseType>-<Region>-<Role> (e.g., SG-W365ENT-EastAsia-User)
    - Policies: <Region>-W365-<LicenseType>-<Suffix> (e.g., EastAsia-W365-Enterprise-Provisioning Policy)
    - Use descriptive prefixes; avoid single letters
    - Include product identifier; distinguish by role, scope, and type
    - Keep names under 64 characters for Azure compatibility
#>

#Requires -Version 7.0

[CmdletBinding()]
param(
    [Parameter()]
    [int]$CloudPCTypeChoice,
    
    [Parameter()]
    [int]$RegionChoice,
    
    [Parameter()]
    [string]$Language = "",
    
    [Parameter()]
    [string]$GroupPrefix = "SG-W365",
    
    [Parameter()]
    [string]$ProvisioningPolicySuffix = "Provisioning Policy"
)

# ----------------------------
# Helpers
# ----------------------------

function Get-ValidChoice {
    param(
        [Parameter(Mandatory)] [int]$Min,
        [Parameter(Mandatory)] [int]$Max
    )
    
    do {
        try {
            $choice = Read-Host "Enter your choice"
            $choiceInt = [int]$choice
            if ($choiceInt -ge $Min -and $choiceInt -le $Max) {
                return $choiceInt
            }
            else {
                Write-Host "Invalid choice. Please enter a number between $Min and $Max." -ForegroundColor Red
            }
        }
        catch {
            Write-Host "Invalid input. Please enter a valid number." -ForegroundColor Red
        }
    } while ($true)
}

function Install-GraphModuleIfNeeded {
    # Only install Microsoft.Graph.Authentication (~2MB lightweight module)
    # This provides Connect-MgGraph and Invoke-MgGraphRequest
    # All other operations use direct REST API calls (no heavy modules needed)
    if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication)) {
        Write-Host "Installing Microsoft.Graph.Authentication (lightweight, ~2MB)..." -ForegroundColor Yellow
        try {
            Install-Module -Name Microsoft.Graph.Authentication -AllowClobber -Force -ErrorAction Stop
            Write-Host "Microsoft.Graph.Authentication module installed successfully" -ForegroundColor Green
        }
        catch {
            Write-Error "Failed to install Microsoft.Graph.Authentication module: $_"
            throw
        }
    }
    else {
        Write-Host "Microsoft.Graph.Authentication module already installed." -ForegroundColor Green
    }
}

function Get-AllGraphItems {
    param(
        [Parameter(Mandatory)] [string]$Uri
    )

    $items = @()
    $nextLink = $Uri

    while ($nextLink) {
        $response = Invoke-MgGraphRequest -Method GET -Uri $nextLink -ErrorAction Stop
        if ($response.value) { $items += $response.value }
        $nextLink = $response.'@odata.nextLink'
    }

    return $items
}

function Get-SkuMetrics {
    param(
        [Parameter(Mandatory)] [string]$DisplayName
    )

    # Parse values like "8vCPU/32GB/256GB" or "16vCPU/64GB/1TB" from the plan display name
    if ($DisplayName -match '(?<vcpu>\d+)vCPU/(?<ram>\d+)GB/(?<storage>[\d\.]+)(?<unit>TB|GB)') {
        $vcpu    = [int]$Matches['vcpu']
        $ramGb   = [int]$Matches['ram']
        $storage = [double]$Matches['storage']
        $unit    = $Matches['unit']
        $storageGb = if ($unit -eq 'TB') { $storage * 1024 } else { $storage }

        return [pscustomobject]@{
            Vcpu      = $vcpu
            RamGb     = $ramGb
            StorageGb = [int][math]::Round($storageGb,0)
        }
    }

    return $null
}

function Test-IsCopilotEligibleSku {
    param(
        [Parameter(Mandatory)] [string]$DisplayName
    )

    $metrics = Get-SkuMetrics -DisplayName $DisplayName
    if (-not $metrics) { return $false }

    return ($metrics.Vcpu -ge 8 -and $metrics.RamGb -ge 32 -and $metrics.StorageGb -ge 256)
}

function Get-OrCreateGroup {
    param(
        [Parameter(Mandatory)] [string]$DisplayName,
        [Parameter(Mandatory)] [string]$Description
    )

    try {
        # Use REST API directly (no Graph module cmdlets needed)
        $response = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/groups?`$filter=displayName eq '$DisplayName'" -ErrorAction Stop
        $existing = $response.value | Select-Object -First 1

        if ($existing) {
            Write-Host "Group already exists: $DisplayName" -ForegroundColor Green
            return $existing.Id
        }

        # MailNickname must be unique, even for security groups
        $mailNick = ("grp-" + ([guid]::NewGuid().ToString("N").Substring(0,10)))

        $params = @{
            DisplayName     = $DisplayName
            MailEnabled     = $false
            MailNickname    = $mailNick
            SecurityEnabled = $true
            Description     = $Description
        }

        Write-Host "Creating group: $DisplayName" -ForegroundColor Yellow
        
        # Use REST API directly (no Graph module cmdlets needed)
        $createResponse = Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/v1.0/groups" -Body ($params | ConvertTo-Json) -ContentType "application/json" -ErrorAction Stop
        Write-Host "Group created successfully with ID: $($createResponse.id)" -ForegroundColor Green
        return $createResponse.id
    }
    catch {
        Write-Error "Failed to create or retrieve group '$DisplayName': $_"
        throw
    }
}

function Get-OrCreateDynamicDeviceGroup {
    param(
        [Parameter(Mandatory)] [string]$DisplayName,
        [Parameter(Mandatory)] [string]$Description,
        [Parameter()] [string]$MembershipRule
    )

    # Default membership rule for Cloud PCs based on device model
    # This uses the deviceModel property which is more reliable than display name
    # Customize the rule based on your needs:
    # - For device model containing "Cloud PC": (device.deviceModel -contains "Cloud PC")
    # - For devices starting with "CPC-": (device.displayName -startsWith "CPC-")
    # - For devices containing "365": (device.displayName -contains "365")
    # - For device category: (device.deviceCategory -eq "CloudPC")
    if ([string]::IsNullOrWhiteSpace($MembershipRule)) {
        $MembershipRule = '(device.deviceModel -contains "Cloud PC")'
    }

    try {
        # Check if group already exists
        $response = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/groups?`$filter=displayName eq '$DisplayName'" -ErrorAction Stop
        $existing = $response.value | Select-Object -First 1

        if ($existing) {
            Write-Host "Dynamic device group already exists: $DisplayName" -ForegroundColor Green
            return $existing.Id
        }

        # MailNickname must be unique, even for security groups
        $mailNick = ("grp-" + ([guid]::NewGuid().ToString("N").Substring(0,10)))

        $params = @{
            displayName              = $DisplayName
            mailEnabled              = $false
            mailNickname             = $mailNick
            securityEnabled          = $true
            description              = $Description
            groupTypes               = @("DynamicMembership")
            membershipRuleProcessingState = "On"
            membershipRule           = $MembershipRule
        }

        Write-Host "Creating dynamic device group: $DisplayName" -ForegroundColor Yellow
        Write-Host "  Membership rule: $MembershipRule" -ForegroundColor Cyan
        
        # Create dynamic group via REST API
        $createResponse = Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/v1.0/groups" -Body ($params | ConvertTo-Json) -ContentType "application/json" -ErrorAction Stop
        Write-Host "Dynamic device group created successfully with ID: $($createResponse.id)" -ForegroundColor Green
        return $createResponse.id
    }
    catch {
        Write-Error "Failed to create or retrieve dynamic device group '$DisplayName': $_"
        throw
    }
}

function Get-OrCreateCloudPcUserSetting {
    param(
        [Parameter(Mandatory)] [string]$DisplayName,
        [Parameter(Mandatory)] [bool]$LocalAdminEnabled,
        [Parameter(Mandatory)] [string]$TargetGroupId
    )

    try {
        # Use REST API directly (no Graph module cmdlets needed)
        $response = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/beta/deviceManagement/virtualEndpoint/userSettings?`$filter=displayName eq '$DisplayName'" -ErrorAction Stop
        $existing = $response.value | Select-Object -First 1
        $settingAlreadyExisted = $null -ne $existing

        if (-not $existing) {
            Write-Host "Creating Cloud PC User Setting: $DisplayName" -ForegroundColor Yellow

            $params = @{
                displayName       = $DisplayName
                localAdminEnabled = $LocalAdminEnabled
                resetEnabled      = $true
                restorePointSetting = @{
                    userRestoreEnabled = $true
                    frequencyInHours   = 6
                }
                crossRegionDisasterRecoverySetting = @{
                    crossRegionDisasterRecoveryEnabled         = $false
                    maintainCrossRegionRestorePointEnabled     = $true
                    disasterRecoveryNetworkSetting             = $null
                    disasterRecoveryType                       = "notConfigured"
                    userInitiatedDisasterRecoveryAllowed       = $false
                }
                notificationSetting = @{
                    restartPromptsDisabled = $false
                }
            }

            # Create user setting via REST API
            $createResponse = Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/beta/deviceManagement/virtualEndpoint/userSettings" -Body ($params | ConvertTo-Json -Depth 10) -ContentType "application/json" -ErrorAction Stop
            $existing = $createResponse
        }
        else {
            Write-Host "Cloud PC User Setting already exists: $DisplayName" -ForegroundColor Green
        }

    # Graph /assign replaces all assignments; always merge existing + new before sending
    # Get existing assignments (expand on the resource, then fallback to /assignments)
    $existingGroupIds = @()
    $gotAssignments = $false
    try {
        # Escape $ in query string so PowerShell does not expand it
        $expanded = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/beta/deviceManagement/virtualEndpoint/userSettings/$($existing.Id)?`$expand=assignments" -ErrorAction SilentlyContinue
        if ($expanded -and $expanded.assignments) {
            $existingGroupIds = @($expanded.assignments | ForEach-Object { $_.target.groupId })
            $gotAssignments = $true
        }
    }
    catch {
        Write-Verbose "Could not retrieve expanded assignments: $_"
    }

    if (-not $gotAssignments) {
        try {
            $assignmentResponse = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/beta/deviceManagement/virtualEndpoint/userSettings/$($existing.Id)/assignments" -ErrorAction SilentlyContinue
            if ($assignmentResponse -and $assignmentResponse.value) {
                $existingGroupIds = @($assignmentResponse.value | ForEach-Object { $_.target.groupId })
                $gotAssignments = $true
            }
        }
        catch {
            Write-Verbose "Could not retrieve assignments via /assignments: $_"
        }
    }

    # Only skip assignment if setting already existed AND we couldn't read assignments (avoid wiping unknown assignments)
    # For newly created settings, empty assignments are expected, so proceed with assignment
    if (-not $gotAssignments -and $settingAlreadyExisted) {
        Write-Warning "Skipping assignment update for '$DisplayName' because existing assignments could not be read."
        return $existing.Id
    }

    # Build complete assignment list: existing + new (avoiding duplicates) – optimized for PS7
    $allGroupIds = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    
    # Add existing groups
    foreach ($gid in $existingGroupIds) {
        if ($gid) { [void]$allGroupIds.Add($gid) }
    }
    
    # Add target group and detect if it's new
    $isNew = $allGroupIds.Add($TargetGroupId)
    
    # Build assignments array
    $assignments = @()
    foreach ($gid in $allGroupIds) {
        $assignments += @{
            id     = $null
            target = @{ groupId = $gid }
        }
    }

    $assignParams = @{
        Assignments = $assignments
    }

    if ($isNew) {
        Write-Host "Assigning '$DisplayName' to group $TargetGroupId" -ForegroundColor Cyan
    }
    else {
        Write-Host "Group $TargetGroupId already assigned to '$DisplayName', preserving all assignments" -ForegroundColor Green
    }
    
    # Assign via REST API /assign endpoint (PS7-only simplified path)
    $uri = "https://graph.microsoft.com/beta/deviceManagement/virtualEndpoint/userSettings/$($existing.Id)/assign"
    Invoke-MgGraphRequest -Method POST -Uri $uri -Body ($assignParams | ConvertTo-Json -Depth 6) -ContentType "application/json" -ErrorAction Stop
    
    Write-Verbose "User setting assigned successfully with $($assignments.Count) total assignments"
    return $existing.Id
    }
    catch {
        Write-Error "Failed to create or assign Cloud PC User Setting '$DisplayName': $_"
        throw
    }
}

function Get-OrCreateProvisioningPolicy {
    param(
        [Parameter(Mandatory)] [string]$DisplayName,
        [Parameter(Mandatory)] [string[]]$AssignGroupIds,
        [Parameter(Mandatory)] [string]$RegionGroup,
        [Parameter(Mandatory)] [string]$CountryRegion,
        [Parameter(Mandatory)] [string]$ImageId,
        [Parameter(Mandatory)] [string]$ImageDisplayName,
        [Parameter()] [string]$Language = "en-GB",
        [Parameter()] [string]$ProvisioningType = "dedicated",
        [Parameter()] [bool]$UserSettingsPersistence = $false,
        [Parameter()] [string]$ServicePlanId = $null
    )

    # Validate required parameters
    if ([string]::IsNullOrWhiteSpace($RegionGroup)) {
        throw "RegionGroup parameter cannot be null or empty"
    }
    if ([string]::IsNullOrWhiteSpace($CountryRegion)) {
        throw "CountryRegion parameter cannot be null or empty"
    }

    try {
        # Retrieve existing policy via REST API (PS7-only)
        $existing = $null
        $policyAlreadyExisted = $false
        try {
            $response = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/beta/deviceManagement/virtualEndpoint/provisioningPolicies?`$filter=displayName eq '$DisplayName'" -ErrorAction Stop
            $existing = $response.value | Select-Object -First 1
        }
        catch {
            Write-Verbose "Failed to retrieve existing policy via Graph API"
        }

        if ($existing) {
            $policyAlreadyExisted = $true
        }

        if (-not $existing) {
            Write-Host "Creating Provisioning Policy: $DisplayName" -ForegroundColor Yellow
            Write-Verbose "RegionGroup parameter: $RegionGroup | CountryRegion parameter: $CountryRegion"

            $params = @{
                displayName        = $DisplayName
                description        = ""
                provisioningType   = $ProvisioningType
                userExperienceType = "cloudPc"
                managedBy          = "windows365"
                imageId            = $ImageId
                imageDisplayName   = $ImageDisplayName
                imageType          = "gallery"
                microsoftManagedDesktop = @{
                    type    = "notManaged"
                    profile = ""
                }
                enableSingleSignOn = $true
                domainJoinConfigurations = @(
                    @{
                        type        = "azureADJoin"
                        regionGroup = $RegionGroup
                        regionName  = $CountryRegion
                    }
                )
                windowsSettings = @{
                    language = $Language
                }
                cloudPcNamingTemplate = $null
                scopeIds = @("0")
                autopatch = @{
                    autopatchGroupId = $null
                }
                userSettingsPersistenceEnabled = $UserSettingsPersistence
                userSettingsPersistenceConfiguration = @{
                    userSettingsPersistenceEnabled        = $UserSettingsPersistence
                    userSettingsPersistenceStorageSizeCategory = "sixteenGB"
                }
                autopilotConfiguration = $null
            }

            # Create via REST API with language fallback
            Write-Verbose "Provisioning payload: $($params | ConvertTo-Json -Depth 6)"
            $attemptedFallback = $false
            $languageFallbackUsed = $false
            try {
                $createResponse = Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/beta/deviceManagement/virtualEndpoint/provisioningPolicies" -Body ($params | ConvertTo-Json) -ContentType "application/json" -ErrorAction Stop
                $existing = $createResponse
            }
            catch {
                if (-not $attemptedFallback -and $params.windowsSettings.language -ne "en-GB") {
                    Write-Warning "Provisioning policy creation failed (language validation). Retrying with en-GB."
                    $params.windowsSettings.language = "en-GB"
                    $attemptedFallback = $true
                    $languageFallbackUsed = $true
                    $createResponse = Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/beta/deviceManagement/virtualEndpoint/provisioningPolicies" -Body ($params | ConvertTo-Json) -ContentType "application/json" -ErrorAction Stop
                    $existing = $createResponse
                }
                else {
                    throw
                }
            }
            
            if ($languageFallbackUsed) {
                Write-Host "Provisioning policy was created using language fallback (en-GB). Please update the language manually in the Microsoft Intune admin center if required." -ForegroundColor Red
            }
            Write-Verbose "Provisioning policy created: $DisplayName with ID: $($existing.id)"
        }
        else {
            Write-Host "Provisioning Policy already exists: $DisplayName" -ForegroundColor Green
        }

        # For Frontline policies, skip group assignment as they are automatically assigned via license
        if ($ServicePlanId) {
            Write-Host "Frontline provisioning policy created successfully." -ForegroundColor Green
            return $existing.id
        }

        # Assign via /assign endpoint with pre-validation of group IDs
        # Graph /assign replaces all assignments; always merge existing + new before sending
        $validGroupIds = @()
        foreach ($gid in $AssignGroupIds | Where-Object { $_ -and (-not [string]::IsNullOrWhiteSpace($_)) }) {
            try {
                $groupResponse = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/groups/$gid" -ErrorAction Stop
                $validGroupIds += $groupResponse.id
            }
            catch {
                Write-Warning "Skipping assignment: group ID not found or inaccessible: $gid"
            }
        }

        if (-not $validGroupIds -or $validGroupIds.Count -eq 0) {
            Write-Warning "No valid group IDs found for assignment. Skipping /assign for '$DisplayName'."
            return $existing.id
        }

        # Get existing assignments and merge with new ones
        $existingGroupIds = @()
        $gotAssignments = $false
        try {
            # Escape $ in query string so PowerShell does not expand it
            $expanded = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/beta/deviceManagement/virtualEndpoint/provisioningPolicies/$($existing.id)?`$expand=assignments" -ErrorAction Stop
            if ($expanded) {
                # Check if assignments property exists and has items
                if ($expanded.PSObject.Properties['assignments'] -and $expanded.assignments) {
                    $existingGroupIds = @($expanded.assignments | ForEach-Object { $_.target.groupId })
                }
                $gotAssignments = $true
                Write-Verbose "Successfully retrieved assignments for provisioning policy. Found $($existingGroupIds.Count) existing assignments."
            }
        }
        catch {
            Write-Verbose "Could not retrieve expanded assignments: $_"
        }

        # For newly created policies or if we successfully read assignments (even if empty), proceed
        # Only skip if policy already existed AND we couldn't read assignments at all
        if (-not $gotAssignments -and $policyAlreadyExisted) {
            Write-Warning "Skipping assignment update for '$DisplayName' because existing assignments could not be read."
            return $existing.id
        }

        # Build complete assignment list: existing + new (avoiding duplicates) – optimized for PS7
        $allGroupIds = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
        foreach ($gid in $existingGroupIds) { if ($gid) { [void]$allGroupIds.Add($gid) } }
        foreach ($gid in $validGroupIds) { if ($gid) { [void]$allGroupIds.Add($gid) } }

        # Build assignments array with complete list
        # For Frontline, use servicePlanId instead of groupId
        $assignments = @()
        if ($ServicePlanId) {
            # Frontline: assign via servicePlanId
            $assignments += @{
                id     = $null
                target = @{ servicePlanId = $ServicePlanId }
            }
        } else {
            # Enterprise: assign via groupId
            foreach ($gid in $allGroupIds) {
                $assignments += @{
                    id     = $null
                    target = @{ groupId = $gid }
                }
            }
        }

        $assignParams = @{ assignments = $assignments }
        $assignJson = $assignParams | ConvertTo-Json -Depth 10 -Compress:$false

        Write-Host "Assigning Provisioning Policy '$DisplayName' to groups..." -ForegroundColor Cyan
        Write-Verbose "Assign payload: $assignJson"

        try {
            Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/beta/deviceManagement/virtualEndpoint/provisioningPolicies/$($existing.id)/assign" -Body $assignJson -ContentType "application/json" -ErrorAction Stop | Out-Null
            Write-Verbose "Provisioning policy assigned successfully via /assign with $($assignments.Count) total assignments"
        }
        catch {
            Write-Error "Failed to assign Provisioning Policy '$DisplayName': $_"
            Write-Error "Payload used: $assignJson"
            throw
        }
        return $existing.id
    }
    catch {
        Write-Error "Failed to create or assign Provisioning Policy '$DisplayName': $_"
        throw
    }
}

# ----------------------------
# Main
# ----------------------------

Install-GraphModuleIfNeeded

# All Graph operations use REST via Invoke-MgGraphRequest; no extra module imports required

# Connect to Graph
Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan
try {
    Connect-MgGraph -Scopes "User.ReadWrite.All","Application.ReadWrite.All","CloudPC.ReadWrite.All","Group.ReadWrite.All" -ErrorAction Stop
    Write-Verbose "Successfully connected to Microsoft Graph"
}
catch {
    Write-Error "Failed to connect to Microsoft Graph: $_"
    throw
}

# Choose Windows 365 License Type (Enterprise or Frontline)
Write-Host "`nChoose your Windows 365 license type:" -ForegroundColor Green
Write-Host " 1. Enterprise" -ForegroundColor White
Write-Host " 2. Frontline" -ForegroundColor White
Write-Host ""
$licenseTypeChoice = Get-ValidChoice -Min 1 -Max 2
$LicenseType = if ($licenseTypeChoice -eq 1) { "Enterprise" } else { "Frontline" }
Write-Host "Selected: Windows 365 $LicenseType" -ForegroundColor Cyan

# For Frontline, choose Dedicated or Shared
if ($LicenseType -eq "Frontline") {
    Write-Host "`nChoose your Frontline provisioning type:" -ForegroundColor Green
    Write-Host " 1. Dedicated (Shared by User)" -ForegroundColor White
    Write-Host " 2. Shared (Shared by Entra Group)" -ForegroundColor White
    Write-Host ""
    $frontlineTypeChoice = Get-ValidChoice -Min 1 -Max 2
    $FrontlineProvisioningType = if ($frontlineTypeChoice -eq 1) { "sharedByUser" } else { "sharedByEntraGroup" }
    $FrontlineDisplayType = if ($frontlineTypeChoice -eq 1) { "Dedicated" } else { "Shared" }
    Write-Host "Selected: Frontline $FrontlineDisplayType" -ForegroundColor Cyan
}
else {
    $FrontlineProvisioningType = "dedicated"  # Enterprise default
    $FrontlineDisplayType = "Enterprise"
}

# Get available Cloud PC service plans from Graph
Write-Host "`nRetrieving available Windows 365 Cloud PC service plans..." -ForegroundColor Cyan
try {
    $servicePlans = @()

    # Use direct Graph REST API call for better reliability
    try {
        $servicePlans = Get-AllGraphItems -Uri "https://graph.microsoft.com/beta/deviceManagement/virtualEndpoint/servicePlans"
    }
    catch {
        Write-Error "Failed to retrieve Cloud PC service plans via REST API: $_"
        throw
    }
    
    if (-not $servicePlans -or $servicePlans.Count -eq 0) {
        Write-Warning "No Cloud PC service plans found. This may indicate insufficient permissions or no available plans."
        throw "Unable to retrieve Cloud PC service plans"
    }
    
    # Normalize objects to ensure DisplayName and Id properties are accessible
    $servicePlans = $servicePlans | ForEach-Object {
        if ($_ -is [string]) { 
            [pscustomobject]@{ DisplayName = $_; Id = $_ } 
        }
        else { 
            # Extract properties from hashtable or object
            $displayName = if ($_.displayName) { $_.displayName } elseif ($_.DisplayName) { $_.DisplayName } else { "Unknown" }
            $id = if ($_.id) { $_.id } elseif ($_.Id) { $_.Id } else { $null }
            [pscustomobject]@{ DisplayName = $displayName; Id = $id; OriginalObject = $_ }
        }
    }

    # Filter based on license type selection
    if ($LicenseType -eq "Enterprise") {
        # Filter out Business and Frontline SKUs for Enterprise
        $servicePlans = $servicePlans | Where-Object { 
            $_.DisplayName -notmatch 'Business' -and $_.DisplayName -notmatch 'Frontline' 
        }
    }
    else {
        # For Frontline, only show Frontline SKUs
        $servicePlans = $servicePlans | Where-Object { 
            $_.DisplayName -match 'Frontline' 
        }
    }

    if (-not $servicePlans -or $servicePlans.Count -eq 0) {
        Write-Warning "No Cloud PC service plans found after filtering. Check available SKUs."
        throw "No $LicenseType Cloud PC service plans available"
    }

    # Sort by display name for consistent ordering
    $servicePlans = $servicePlans | Sort-Object DisplayName
    $CloudPCType = $servicePlans
    
    Write-Verbose "Found $($CloudPCType.Count) Cloud PC service plans (Business/Frontline filtered out)"
}
catch {
    Write-Error "Failed to retrieve Cloud PC service plans: $_"
    throw
}

# Pick Cloud PC type
if (-not $CloudPCTypeChoice) {
    Write-Host "`nChoose your Windows 365 Cloud PC by selecting its corresponding number:" -ForegroundColor Green
    for ($i = 0; $i -lt $CloudPCType.Count; $i++) {
        Write-Host ("{0,2}. {1}" -f ($i + 1), $CloudPCType[$i].DisplayName)
    }
    Write-Host ""
    
    $Windows365CloudPCTypeVariable = Get-ValidChoice -Min 1 -Max $CloudPCType.Count
}
else {
    if ($CloudPCTypeChoice -lt 1 -or $CloudPCTypeChoice -gt $CloudPCType.Count) {
        Write-Error "Invalid CloudPCTypeChoice parameter. Must be between 1 and $($CloudPCType.Count)."
        Write-Host "Available plans:" -ForegroundColor Yellow
        for ($i = 0; $i -lt $CloudPCType.Count; $i++) {
            Write-Host ("{0,2}. {1}" -f ($i + 1), $CloudPCType[$i].DisplayName)
        }
        throw "CloudPCTypeChoice out of range"
    }
    $Windows365CloudPCTypeVariable = $CloudPCTypeChoice
    Write-Verbose "Using parameter-specified Cloud PC type choice: $CloudPCTypeChoice"
}

$SelectedServicePlan = $CloudPCType[$Windows365CloudPCTypeVariable - 1]
$Windows365CloudPCType = $SelectedServicePlan.DisplayName
$IsCopilotEligible = Test-IsCopilotEligibleSku -DisplayName $Windows365CloudPCType

# Get supported region groups for Cloud PC
Write-Host "`nRetrieving supported Windows 365 region groups..." -ForegroundColor Cyan
try {
    $supportedRegions = @()

    # Use direct Graph REST API call for better reliability
    try {
        # Filter for Windows 365 supported regions and select relevant fields
        $uri = "https://graph.microsoft.com/beta/deviceManagement/virtualEndpoint/supportedRegions?`$filter=supportedSolution eq 'windows365'&`$select=id,displayName,regionStatus,supportedSolution,regionGroup,cloudDevicePlatformSupported,geographicLocationType"
        $supportedRegions = Get-AllGraphItems -Uri $uri
    }
    catch {
        Write-Error "Failed to retrieve supported regions via REST API: $_"
        throw
    }

    if (-not $supportedRegions -or $supportedRegions.Count -eq 0) {
        Write-Warning "No supported region groups returned. Check your permissions or tenant availability."
        throw "Unable to retrieve supported region groups"
    }
    
    # Filter out duplicate entries - keep only regions where geographicLocationType is null
    # (Some regions have duplicate entries with different regionGroup values; we want the canonical ones)
    $supportedRegions = $supportedRegions | Where-Object {
        $geoType = $_.geographicLocationType
        if (-not $geoType) { $geoType = $_.GeographicLocationType }
        $null -eq $geoType
    }
    
    Write-Verbose "Filtered to $($supportedRegions.Count) unique regions (removed duplicates)"

    # Normalize objects to ensure RegionGroup, RegionName, and DisplayName are present (PS7 null-coalescing)
    $supportedRegions = $supportedRegions | ForEach-Object {
        if ($_ -is [string]) {
            [pscustomobject]@{ RegionGroup = $_; RegionName = $_; DisplayName = $_ }
        }
        else {
            $rg = ($_.regionGroup) ?? ($_.RegionGroup)
            $dn = ($_.displayName) ?? ($_.DisplayName)
            $rn = $dn ?? $rg
            [pscustomobject]@{ RegionGroup = $rg; RegionName = $rn; DisplayName = $dn }
        }
    }

    Write-Verbose "Supported regions raw data (first 3):"
    $supportedRegions | Select-Object -First 3 | ForEach-Object { Write-Verbose ($_ | ConvertTo-Json -Depth 2) }

    # Use individual regions (regionName) so we can target a specific location when geographicLocationType="region"
    $regionOptions = $supportedRegions |
        Sort-Object DisplayName, RegionName |
        Select-Object -Property RegionGroup, RegionName, DisplayName -Unique

    Write-Verbose "Found $($regionOptions.Count) supported region groups"
}
catch {
    Write-Error "Failed to retrieve supported region groups: $_"
    throw
}

if (-not $RegionChoice) {
    # Step 1: Get unique regionGroup values from normalized objects
    $uniqueGroups = @()
    foreach ($region in $supportedRegions) {
        $rg = $region.RegionGroup
        if ($rg -and $rg -notin $uniqueGroups) {
            $uniqueGroups += $rg
        }
    }
    $uniqueGroups = $uniqueGroups | Sort-Object
    
    Write-Verbose "Found $($uniqueGroups.Count) unique region groups: $($uniqueGroups -join ', ')"

    Write-Host "`nChoose your Windows 365 region GROUP:" -ForegroundColor Green
    for ($i = 0; $i -lt $uniqueGroups.Count; $i++) {
        $groupName = $uniqueGroups[$i]
        # Convert camelCase to Title Case (e.g., "southAmerica" -> "South America")
        $words = [regex]::Split($groupName, '(?<=[a-z])(?=[A-Z])')
        $friendlyName = ($words | ForEach-Object { $_.Substring(0,1).ToUpper() + $_.Substring(1).ToLower() }) -join ' '
        Write-Host ("{0,2}. {1}" -f ($i + 1), $friendlyName)
    }
    Write-Host ""

    $groupChoice = Get-ValidChoice -Min 1 -Max $uniqueGroups.Count
    $selectedGroupValue = $uniqueGroups[$groupChoice - 1]

    # Step 2: Get all displayName values for the selected regionGroup
    Write-Verbose "Selected region group: '$selectedGroupValue'"
    $regionsInGroup = $supportedRegions | Where-Object { $_.RegionGroup -eq $selectedGroupValue } | Sort-Object DisplayName
    
    Write-Verbose "Found $($regionsInGroup.Count) regions in group '$selectedGroupValue'"
    
    if (-not $regionsInGroup -or $regionsInGroup.Count -eq 0) {
        Write-Warning "No regions found for group '$selectedGroupValue'. This may indicate a data consistency issue."
        Write-Host "`nAvailable regions data:" -ForegroundColor Yellow
        $supportedRegions | ForEach-Object {
            Write-Host "  DisplayName: $($_.DisplayName), RegionGroup: '$($_.RegionGroup)'" -ForegroundColor Gray
        }
        throw "No regions available for the selected region group"
    }
    
    Write-Host "`nChoose your Windows 365 REGION:" -ForegroundColor Green
    for ($i = 0; $i -lt $regionsInGroup.Count; $i++) {
        $r = $regionsInGroup[$i]
        Write-Host ("{0,2}. {1}" -f ($i + 1), $r.DisplayName)
    }
    Write-Host ""

    $regionChoice = Get-ValidChoice -Min 1 -Max $regionsInGroup.Count
    $selectedRegion = $regionsInGroup[$regionChoice - 1]

    $SelectedRegionGroup = $selectedRegion.RegionGroup
    $SelectedRegionDisplayName = $selectedRegion.DisplayName
    $SelectedRegionName = $selectedRegion.RegionName
    
    # No transformation; use regionGroup as-is for the policy
    $SelectedCountryRegion = $SelectedRegionName
}
else {
    # Legacy flat selection using RegionChoice parameter
    if ($RegionChoice -lt 1 -or $RegionChoice -gt $regionOptions.Count) {
        Write-Error "Invalid RegionChoice parameter. Must be between 1 and $($regionOptions.Count)."
        Write-Host "Available regions:" -ForegroundColor Yellow
        for ($i = 0; $i -lt $regionOptions.Count; $i++) {
            $option = $regionOptions[$i]
            Write-Host ("{0,2}. {1} ({2})" -f ($i + 1), $option.DisplayName, $option.RegionGroup)
        }
        throw "RegionChoice out of range"
    }
    $SelectedRegionGroup = $regionOptions[$RegionChoice - 1].RegionGroup
    $SelectedRegionDisplayName = $regionOptions[$RegionChoice - 1].DisplayName
    $SelectedRegionName = $regionOptions[$RegionChoice - 1].RegionName
    
    # Use regionGroup and region name as-is from the API
    $SelectedCountryRegion = $SelectedRegionName
    Write-Verbose "Using parameter-specified region choice: $RegionChoice"
}

# Fallbacks to avoid binding errors when API data is sparse
if ([string]::IsNullOrWhiteSpace($SelectedRegionName)) {
    $SelectedRegionName = $SelectedRegionGroup
}
if ([string]::IsNullOrWhiteSpace($SelectedCountryRegion)) {
    $SelectedCountryRegion = $SelectedRegionName
}

Write-Host "`nCreating/reusing Windows 365 objects for: $Windows365CloudPCType" -ForegroundColor Green

# Get available images for Cloud PC
Write-Host "`nRetrieving available Windows 365 images..." -ForegroundColor Cyan
try {
    $images = @()

    # Try gallery images endpoint first (more reliable for Microsoft gallery images)
    try {
        Write-Verbose "Attempting to retrieve gallery images..."
        $images = Get-AllGraphItems -Uri "https://graph.microsoft.com/beta/deviceManagement/virtualEndpoint/galleryImages"
        Write-Verbose "Retrieved $($images.Count) images from gallery endpoint"
    }
    catch {
        Write-Verbose "Gallery images endpoint failed, trying device images endpoint: $_"
        
        # Fallback to device images endpoint using REST API
        try {
            $images = Get-AllGraphItems -Uri "https://graph.microsoft.com/beta/deviceManagement/virtualEndpoint/deviceImages"
        }
        catch {
            Write-Verbose "Beta endpoint failed, trying v1 endpoint..."
            try {
                $images = Get-AllGraphItems -Uri "https://graph.microsoft.com/v1.0/deviceManagement/virtualEndpoint/deviceImages"
            }
            catch {
                Write-Error "Failed to retrieve device images: $_"
                throw
            }
        }
    }

    Write-Verbose "Retrieved $($images.Count) total images from Graph"

    if (-not $images -or $images.Count -eq 0) {
        Write-Warning "No device images found from Graph. Using a known Microsoft gallery image ID."
        $SelectedImage = @{
            Id          = "microsoftwindowsdesktop_windows-ent-cpc_win11-25h2-ent-cpc-m365"
            DisplayName = "Windows 11 Enterprise + Microsoft 365 Apps 25H2"
        }
    }
    else {
        # Filter out unsupported images - keep supported, supportedWithWarning, or images without status property
        $availableImages = $images | Where-Object { 
            -not $_.Status -or $_.Status -ne "notSupported"
        }

        if (-not $availableImages -or $availableImages.Count -eq 0) {
            Write-Warning "No available device images found. All images may be in processing or error state."
            Write-Host "Showing all images (including unavailable ones):" -ForegroundColor Yellow
            $availableImages = $images | Sort-Object DisplayName
        }
        else {
            # Sort by display name
            $availableImages = $availableImages | Sort-Object DisplayName
        }

        Write-Verbose "Found $($availableImages.Count) avdailable device images"

        # Select image
        Write-Host "`nChoose a Windows 11 image by selecting its corresponding number:" -ForegroundColor Green
        for ($i = 0; $i -lt $availableImages.Count; $i++) {
            $status = if ($availableImages[$i].Status -eq "available") { "" } elseif ($availableImages[$i].Status) { " [Status: $($availableImages[$i].Status)]" } else { "" }
            Write-Host ("{0,2}. {1}{2}" -f ($i + 1), $availableImages[$i].DisplayName, $status)
        }
        Write-Host ""

        $imageChoice = Get-ValidChoice -Min 1 -Max $availableImages.Count
        $SelectedImage = $availableImages[$imageChoice - 1]
    }

    Write-Host "Selected image: $($SelectedImage.DisplayName)" -ForegroundColor Cyan

    # Fallback if selected image is placeholder
    if (-not $SelectedImage.Id -or $SelectedImage.Id -eq "default") {
        Write-Verbose "Selected image was default/empty. Applying known gallery image fallback."
        $SelectedImage = @{
            Id          = "microsoftwindowsdesktop_windows-ent-cpc_win11-25h2-ent-cpc-m365"
            DisplayName = "Windows 11 Enterprise + Microsoft 365 Apps 25H2"
        }
    }
}
catch {
    Write-Error "Failed to retrieve device images: $_"
    Write-Verbose "Error details: $($_ | Out-String)"
    throw
}

# Language selection
$SupportedLanguages = @(
    @{ DisplayName = "Arabic (Saudi Arabia)"; Code = "ar-SA" },
    @{ DisplayName = "Bulgarian (Bulgaria)"; Code = "bg-BG" },
    @{ DisplayName = "Chinese (Simplified)"; Code = "zh-CN" },
    @{ DisplayName = "Chinese (Traditional)"; Code = "zh-TW" },
    @{ DisplayName = "Croatian (Croatia)"; Code = "hr-HR" },
    @{ DisplayName = "Czech (Czech Republic)"; Code = "cs-CZ" },
    @{ DisplayName = "Danish (Denmark)"; Code = "da-DK" },
    @{ DisplayName = "Dutch (Netherlands)"; Code = "nl-NL" },
    @{ DisplayName = "English (Australia)"; Code = "en-AU" },
    @{ DisplayName = "English (Ireland)"; Code = "en-IE" },
    @{ DisplayName = "English (New Zealand)"; Code = "en-NZ" },
    @{ DisplayName = "English (United Kingdom)"; Code = "en-GB" },
    @{ DisplayName = "English (United States)"; Code = "en-US" },
    @{ DisplayName = "Estonian (Estonia)"; Code = "et-EE" },
    @{ DisplayName = "Finnish (Finland)"; Code = "fi-FI" },
    @{ DisplayName = "French (Canada)"; Code = "fr-CA" },
    @{ DisplayName = "French (France)"; Code = "fr-FR" },
    @{ DisplayName = "German (Germany)"; Code = "de-DE" },
    @{ DisplayName = "Greek (Greece)"; Code = "el-GR" },
    @{ DisplayName = "Hebrew (Israel)"; Code = "he-IL" },
    @{ DisplayName = "Hindi (India)"; Code = "hi-IN" },
    @{ DisplayName = "Hungarian (Hungary)"; Code = "hu-HU" },
    @{ DisplayName = "Italian (Italy)"; Code = "it-IT" },
    @{ DisplayName = "Japanese (Japan)"; Code = "ja-JP" },
    @{ DisplayName = "Korean (Korea)"; Code = "ko-KR" },
    @{ DisplayName = "Latvian (Latvia)"; Code = "lv-LV" },
    @{ DisplayName = "Lithuanian (Lithuania)"; Code = "lt-LT" },
    @{ DisplayName = "Norwegian (Bokmål)"; Code = "nb-NO" },
    @{ DisplayName = "Polish (Poland)"; Code = "pl-PL" },
    @{ DisplayName = "Portuguese (Brazil)"; Code = "pt-BR" },
    @{ DisplayName = "Portuguese (Portugal)"; Code = "pt-PT" },
    @{ DisplayName = "Romanian (Romania)"; Code = "ro-RO" },
    @{ DisplayName = "Russian (Russia)"; Code = "ru-RU" },
    @{ DisplayName = "Serbian (Latin)"; Code = "sr-Latn-RS" },
    @{ DisplayName = "Slovak (Slovakia)"; Code = "sk-SK" },
    @{ DisplayName = "Slovenian (Slovenia)"; Code = "sl-SI" },
    @{ DisplayName = "Spanish (Mexico)"; Code = "es-MX" },
    @{ DisplayName = "Spanish (Spain)"; Code = "es-ES" },
    @{ DisplayName = "Swedish (Sweden)"; Code = "sv-SE" },
    @{ DisplayName = "Thai (Thailand)"; Code = "th-TH" },
    @{ DisplayName = "Turkish (Turkey)"; Code = "tr-TR" },
    @{ DisplayName = "Ukrainian (Ukraine)"; Code = "uk-UA" }
)

if ([string]::IsNullOrWhiteSpace($Language)) {
    Write-Host "`nSelect your Windows 11 language from the grid below:" -ForegroundColor Green
    
    # Add a custom property to display default indicator
    $languagesForGrid = $SupportedLanguages | ForEach-Object {
        $default = if ($_.Code -eq "en-GB") { "[Default]" } else { "" }
        [PSCustomObject]@{
            "Language" = "$($_.DisplayName) $default"
            "Code" = $_.Code
            "DisplayName" = $_.DisplayName
        }
    }
    
    $selectedFromGrid = $languagesForGrid | Out-GridView -Title "Select Windows 11 Language" -OutputMode Single
    
    if ($null -eq $selectedFromGrid) {
        Write-Host "No language selected. Using default (en-GB)" -ForegroundColor Yellow
        $SelectedLanguage = "en-GB"
    }
    else {
        $SelectedLanguage = $selectedFromGrid.Code
        Write-Host "Selected language: $($selectedFromGrid.DisplayName)" -ForegroundColor Cyan
    }
}
else {
    # Validate provided language code
    $languageFound = $SupportedLanguages | Where-Object { $_.Code -eq $Language }
    if ($languageFound) {
        $SelectedLanguage = $Language
        Write-Host "Using language: $($languageFound.DisplayName)" -ForegroundColor Cyan
    }
    else {
        Write-Warning "Language code '$Language' not recognized. Using default en-GB"
        $SelectedLanguage = "en-GB"
    }
}

# Calculate friendly region name for use in group names
$policyRegionNameRaw  = if ($SelectedRegionDisplayName) { $SelectedRegionDisplayName } else { $SelectedRegionName -replace '[_-]', ' ' -replace '(?<=.)([A-Z])',' $1' }
$policyRegionName     = (Get-Culture).TextInfo.ToTitleCase($policyRegionNameRaw.ToLower().Trim())
$regionLabel          = (Get-Culture).TextInfo.ToTitleCase(($SelectedRegionName -replace '[_-]', ' ').ToLower().Trim())

# Groups - Create licensing group and location-based groups with naming convention
# Naming convention: <GroupPrefix>-<Type>-<Region>-<Role>
# Example: SG-W365ENT-EastAsia-User
# The prefix can be customized via -GroupPrefix parameter; default is "SG-W365"
# Best Practice Tips:
# - Use descriptive prefixes (avoid single letters)
# - Include product identifier (W365, CloudPC, etc.)
# - Distinguish by role (User/Admin) and scope (Region/License)
# - Keep under 64 characters for Azure compatibility

# Licensing group (merges all users and admins for license assignment) - based on Cloud PC type
$LicensingGroupName = "${GroupPrefix}CloudPC_${Windows365CloudPCType}"

# Create description with service plan info for easy license matching
$LicensingGroupDescription = "All Windows 365 users and admins for license assignment`nService Plan ID: $($SelectedServicePlan.Id)"

# Location-based groups for user/admin settings
$licenseTypeInfix = if ($LicenseType -eq "Frontline") { "FL" } else { "ENT" }
$groupBase = "${GroupPrefix}-${licenseTypeInfix}"
$UserGroupName  = "${groupBase}-${regionLabel}-User"
$AdminGroupName = "${groupBase}-${regionLabel}-Admin"

Write-Verbose "Creating/retrieving groups..."
Write-Verbose "Using group prefix: $GroupPrefix (customize with -GroupPrefix parameter if needed)"
Write-Verbose "Service Plan ID: $($SelectedServicePlan.Id)"
$GroupIDLicensing = Get-OrCreateGroup -DisplayName $LicensingGroupName -Description $LicensingGroupDescription
$GroupIDUser      = Get-OrCreateGroup -DisplayName $UserGroupName  -Description "Windows 365 users in $LocationName"
$GroupIDAdmin     = Get-OrCreateGroup -DisplayName $AdminGroupName -Description "Windows 365 local admins in $LocationName"

# Dynamic device group for Cloud PCs
$DynamicDeviceGroupName = "${GroupPrefix}CloudPC-Devices"
$DynamicDeviceGroupDescription = "Dynamic group that automatically includes all Windows 365 Cloud PC devices based on naming convention (CPC-*). Customize the membership rule if your Cloud PCs use a different naming convention."
$GroupIDCloudPCDevices = Get-OrCreateDynamicDeviceGroup -DisplayName $DynamicDeviceGroupName -Description $DynamicDeviceGroupDescription

# Allow time for group replication before policy assignment (prov policy /assign is more eventual)
Write-Verbose "Waiting for group replication to complete..."
Start-Sleep -Seconds 10

# User Settings
Write-Verbose "Creating/retrieving Cloud PC User Settings..."
Get-OrCreateCloudPcUserSetting -DisplayName "W365_AdminSettings" -LocalAdminEnabled $true  -TargetGroupId $GroupIDAdmin
Get-OrCreateCloudPcUserSetting -DisplayName "W365_UserSettings"  -LocalAdminEnabled $false -TargetGroupId $GroupIDUser

if ($IsCopilotEligible) {
    Write-Host "`n⚠️ AI Enabled Cloud PC detected: creating AI_Enabled_Cloud_PC user setting and assigning $LicensingGroupName." -ForegroundColor Green
    Get-OrCreateCloudPcUserSetting -DisplayName "AI_Enabled_Cloud_PC" -LocalAdminEnabled $false -TargetGroupId $GroupIDLicensing | Out-Null
}

# Provisioning Policy - per Region
Write-Verbose "Creating/retrieving Provisioning Policy for region..."

# Set provisioning type and user settings persistence based on license type
$provType = $FrontlineProvisioningType
$userPersistence = if ($LicenseType -eq "Frontline" -and $FrontlineProvisioningType -eq "sharedByEntraGroup") { $true } else { $false }
$servicePlanId = if ($LicenseType -eq "Frontline") { $SelectedServicePlan.Id } else { $null }

# Provisioning Policy naming convention: <Region>-W365-<LicenseType>-<Suffix>
# Example: EastAsia-W365-Enterprise-Provisioning Policy
# Customize suffix via -ProvisioningPolicySuffix parameter if needed
$ProvisioningPolicyName = "$policyRegionName-W365-$LicenseType-$ProvisioningPolicySuffix"
Write-Verbose "Provisioning Policy naming convention: Region-W365-LicenseType-Suffix (customize suffix with -ProvisioningPolicySuffix parameter)"
$cloudPcProvisioningPolicyId = Get-OrCreateProvisioningPolicy -DisplayName $ProvisioningPolicyName -AssignGroupIds @($GroupIDAdmin, $GroupIDUser) -RegionGroup $SelectedRegionGroup -CountryRegion $SelectedCountryRegion -ImageId $SelectedImage.Id -ImageDisplayName $SelectedImage.DisplayName -Language $SelectedLanguage -ProvisioningType $provType -UserSettingsPersistence $userPersistence -ServicePlanId $servicePlanId

Write-Host "`nDone ✅" -ForegroundColor Green

# Display consolidated manual steps
Write-Host "`n$('=' * 80)" -ForegroundColor Cyan
Write-Host "PLEASE NOTE - Manual Steps Required" -ForegroundColor Cyan
Write-Host "$('=' * 80)" -ForegroundColor Cyan

Write-Host "`n1. License Assignment:" -ForegroundColor Yellow
Write-Host "   Assign the correct Windows 365 license to the licensing group:" -ForegroundColor White
Write-Host "   → $LicensingGroupName" -ForegroundColor Green

if ($LicenseType -eq "Frontline") {
    Write-Host "`n2. Frontline Policy Assignment:" -ForegroundColor Yellow
    Write-Host "   ⚠️  Frontline provisioning policies require manual assignment after licenses are purchased." -ForegroundColor White
    Write-Host "   Once users receive Windows 365 Frontline licenses, the provisioning policy will" -ForegroundColor White
    Write-Host "   automatically apply based on the service plan selected." -ForegroundColor White
    Write-Host "   → Policy: $ProvisioningPolicyName" -ForegroundColor Green
}

Write-Host "`n$( if ($LicenseType -eq 'Frontline') { '3' } else { '2' } ). User and Admin Group Assignments:" -ForegroundColor Yellow
Write-Host "   The following location-based groups have been assigned to settings policies:" -ForegroundColor White
Write-Host "   → $UserGroupName (user settings)" -ForegroundColor Green
Write-Host "   → $AdminGroupName (admin settings)" -ForegroundColor Green

Write-Host "`n$( if ($LicenseType -eq 'Frontline') { '4' } else { '3' } ). Cloud PC Devices Group:" -ForegroundColor Yellow
Write-Host "   A dynamic device security group has been created to automatically include all Cloud PC devices:" -ForegroundColor White
Write-Host "   → $DynamicDeviceGroupName" -ForegroundColor Green
Write-Host "   Current membership rule: (device.displayName -startsWith `"CPC-`")" -ForegroundColor White
Write-Host "   To customize the rule (e.g., for different naming conventions), update the group membership rule in:" -ForegroundColor White
Write-Host "   Entra ID > Groups > $DynamicDeviceGroupName > Membership rules" -ForegroundColor Cyan
Write-Host "   Examples of other membership rules:" -ForegroundColor White
Write-Host "     • For devices starting with 'CloudPC-': (device.displayName -startsWith `"CloudPC-`")" -ForegroundColor Gray
Write-Host "     • For devices containing '365': (device.displayName -contains `"365`")" -ForegroundColor Gray
Write-Host "     • For a custom category: (device.deviceCategory -eq `"CloudPC`")" -ForegroundColor Gray

Write-Host "`n$( if ($LicenseType -eq 'Frontline') { '5' } else { '4' } ). Cross-Region Disaster Recovery (Optional):" -ForegroundColor Yellow
Write-Host "   DR settings have been created with defaults disabled." -ForegroundColor White
Write-Host "   If you need to configure DR for your Cloud PCs, manually update the settings" -ForegroundColor White
Write-Host "   in the Microsoft Intune admin center under Devices > Cloud PCs > User settings." -ForegroundColor White

Write-Host "`n$('=' * 80)" -ForegroundColor Cyan

# Cleanup
Write-Verbose "Disconnecting from Microsoft Graph..."
#Disconnect-MgGraph | Out-Null
Write-Host "`nDisconnected from Microsoft Graph." -ForegroundColor Cyan