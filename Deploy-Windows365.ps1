<#
.SYNOPSIS
    Create a Windows 365 environment (Groups, User Settings, Provisioning Policy)
.DESCRIPTION
    Prompts for a Cloud PC SKU, then:
    - Ensures Microsoft.Graph is installed
    - Connects to Microsoft Graph (Beta)
    - Creates (or reuses) two Azure AD security groups
    - Creates (or reuses) Cloud PC User Settings policies + assigns them
    - Creates (or reuses) a Cloud PC Provisioning Policy + assigns it
    - Preserves existing assignments by merging current + new groups (Graph /assign is replace-all)
.PARAMETER CloudPCTypeChoice
    The Cloud PC SKU choice (dynamic from tenant). If not provided, will prompt interactively.
.PARAMETER RegionChoice
    The region choice (dynamic from tenant). Regions are always prompted for selection.
.NOTES
    Script name: Deploy-Windows365.ps1
    Author:      Jon Jarvis
    Required scopes: User.ReadWrite.All, Application.ReadWrite.All, CloudPC.ReadWrite.All, Group.ReadWrite.All
#>

[CmdletBinding()]
param(
    [Parameter()]
    [int]$CloudPCTypeChoice,
    
    [Parameter()]
    [int]$RegionChoice,
    
    [Parameter()]
    [string]$Language = ""
)

# ----------------------------
# Helpers
# ----------------------------

function Install-GraphModuleIfNeeded {
    if (-not (Get-Module -ListAvailable -Name Microsoft.Graph)) {
        Write-Host "Microsoft.Graph module not found. Installing..." -ForegroundColor Yellow
        try {
            Install-Module -Name Microsoft.Graph -AllowClobber -Force -ErrorAction Stop
            Write-Verbose "Microsoft.Graph module installed successfully"
        }
        catch {
            Write-Error "Failed to install Microsoft.Graph module: $_"
            throw
        }
    }
    else {
        Write-Host "Microsoft.Graph module already installed. Skipping install." -ForegroundColor Green
    }
    # Install required Beta DeviceManagement.Administration module
    if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Beta.DeviceManagement.Administration)) {
        Write-Host "Microsoft.Graph.Beta.DeviceManagement.Administration module not found. Installing..." -ForegroundColor Yellow
        try {
            Install-Module -Name Microsoft.Graph.Beta.DeviceManagement.Administration -MinimumVersion 2.23.0 -AllowClobber -Force -ErrorAction Stop
            Write-Verbose "Microsoft.Graph.Beta.DeviceManagement.Administration module installed successfully"
        }
        catch {
            Write-Warning "Failed to install Microsoft.Graph.Beta.DeviceManagement.Administration module: $_. Will use REST API fallback."
        }
    }
    else {
        Write-Host "Microsoft.Graph.Beta.DeviceManagement.Administration module already installed." -ForegroundColor Green
    }}

function Get-ValidChoice {
    param(
        [int]$Min = 1,
        [int]$Max
    )

    $choice = $null

    do {
        $choiceRaw = Read-Host "Enter a number ($Min-$Max)"
        $choiceOk = [int]::TryParse($choiceRaw, [ref]$choice)

        if (-not $choiceOk -or $choice -lt $Min -or $choice -gt $Max) {
            Write-Host "Invalid option. Please enter a number between $Min and $Max." -ForegroundColor Red
            $choice = $null
        }
    } while ($null -eq $choice)

    return $choice
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

function Get-OrCreateGroup {
    param(
        [Parameter(Mandatory)] [string]$DisplayName,
        [Parameter(Mandatory)] [string]$Description
    )

    try {
        $existing = $null
        
        # Try Get-MgGroup first, but fall back to direct Graph API if module is not available
        if (Get-Command Get-MgGroup -ErrorAction SilentlyContinue) {
            try {
                $existing = Get-MgGroup -Filter "displayName eq '$DisplayName'" -ConsistencyLevel eventual -CountVariable c -ErrorAction Stop
            }
            catch {
                Write-Verbose "Get-MgGroup failed, falling back to direct Graph API: $_"
                # Fall through to direct API call
            }
        }
        
        # Fallback to direct Graph API if cmdlet failed or not available
        if (-not $existing) {
            $response = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/groups?`$filter=displayName eq '$DisplayName'" -ErrorAction Stop
            $existing = $response.value | Select-Object -First 1
        }

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
        
        # Try New-MgGroup first, then fallback to direct Graph API
        if (Get-Command New-MgGroup -ErrorAction SilentlyContinue) {
            try {
                $new = New-MgGroup -BodyParameter $params -ErrorAction Stop
                Write-Verbose "Group created successfully with ID: $($new.Id)"
                return $new.Id
            }
            catch {
                Write-Verbose "New-MgGroup failed, falling back to direct Graph API: $_"
                # Fall through to direct API call
            }
        }
        
        # Fallback to direct Graph API for group creation
        $createResponse = Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/v1.0/groups" -Body ($params | ConvertTo-Json) -ContentType "application/json" -ErrorAction Stop
        Write-Verbose "Group created successfully with ID: $($createResponse.id)"
        return $createResponse.id
    }
    catch {
        Write-Error "Failed to create or retrieve group '$DisplayName': $_"
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
        $existing = $null
        $settingAlreadyExisted = $false
        
        # Try cmdlet first, but fall back to direct Graph API if module is not available
        if (Get-Command Get-MgDeviceManagementVirtualEndpointUserSetting -ErrorAction SilentlyContinue) {
            try {
                $existing = Get-MgDeviceManagementVirtualEndpointUserSetting -Filter "displayName eq '$DisplayName'" -ErrorAction Stop
                if ($existing) { $settingAlreadyExisted = $true }
            }
            catch {
                Write-Verbose "Get-MgDeviceManagementVirtualEndpointUserSetting failed, falling back to direct Graph API: $_"
                # Fall through to direct API call
            }
        }
        
        # Fallback to direct Graph API if cmdlet failed or not available
        if (-not $existing) {
            $response = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/beta/deviceManagement/virtualEndpoint/userSettings?`$filter=displayName eq '$DisplayName'" -ErrorAction Stop
            $existing = $response.value | Select-Object -First 1
            if ($existing) { $settingAlreadyExisted = $true }
        }

        if (-not $existing) {
            Write-Host "Creating Cloud PC User Setting: $DisplayName" -ForegroundColor Yellow

            $params = @{
                DisplayName       = $DisplayName
                LocalAdminEnabled = $LocalAdminEnabled
                RestorePointSetting = @{
                    UserRestoreEnabled = $true
                    FrequencyInHours   = 12
                }
            }

            # Try New-MgDeviceManagementVirtualEndpointUserSetting first, then fallback
            if (Get-Command New-MgDeviceManagementVirtualEndpointUserSetting -ErrorAction SilentlyContinue) {
                try {
                    $existing = New-MgDeviceManagementVirtualEndpointUserSetting -BodyParameter $params -ErrorAction Stop
                }
                catch {
                    Write-Verbose "New-MgDeviceManagementVirtualEndpointUserSetting failed, falling back to direct Graph API: $_"
                    # Fall through to direct API call
                }
            }
            
            # Fallback to direct Graph API for user setting creation
            if (-not $existing) {
                $createResponse = Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/beta/deviceManagement/virtualEndpoint/userSettings" -Body ($params | ConvertTo-Json -Depth 3) -ContentType "application/json" -ErrorAction Stop
                $existing = $createResponse
            }
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

    # Build complete assignment list: existing + new (avoiding duplicates)
    $allGroupIds = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    
    # Add existing groups
    foreach ($gid in $existingGroupIds) {
        $allGroupIds.Add($gid) | Out-Null
    }
    
    # Add target group if not already present
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
    
    # Try Set-MgDeviceManagementVirtualEndpointUserSetting first, then fallback to direct Graph API
    $assignSuccess = $false
    if (Get-Command Set-MgDeviceManagementVirtualEndpointUserSetting -ErrorAction SilentlyContinue) {
        try {
            Set-MgDeviceManagementVirtualEndpointUserSetting -CloudPcUserSettingId $existing.Id -BodyParameter $assignParams -ErrorAction Stop
            $assignSuccess = $true
        }
        catch {
            Write-Verbose "Set-MgDeviceManagementVirtualEndpointUserSetting failed, falling back to direct Graph API: $_"
            # Fall through to direct API call
        }
    }
    
    if (-not $assignSuccess) {
        # Direct Graph API fallback using /assign endpoint (POST with assignments array)
        $uri = "https://graph.microsoft.com/beta/deviceManagement/virtualEndpoint/userSettings/$($existing.Id)/assign"
        Invoke-MgGraphRequest -Method POST -Uri $uri -Body ($assignParams | ConvertTo-Json -Depth 5) -ContentType "application/json" -ErrorAction Stop
    }
    
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
        [Parameter()] [string]$Language = "en-GB"
    )

    # Validate required parameters
    if ([string]::IsNullOrWhiteSpace($RegionGroup)) {
        throw "RegionGroup parameter cannot be null or empty"
    }
    if ([string]::IsNullOrWhiteSpace($CountryRegion)) {
        throw "CountryRegion parameter cannot be null or empty"
    }

    try {
        # Try to get existing policy - first try cmdlet, then direct API
        $existing = $null
        $policyAlreadyExisted = $false
        
        if (Get-Command Get-MgBetaDeviceManagementVirtualEndpointProvisioningPolicy -ErrorAction SilentlyContinue) {
            try {
                $existing = Get-MgBetaDeviceManagementVirtualEndpointProvisioningPolicy -Filter "displayName eq '$DisplayName'" -ErrorAction Stop
            }
            catch {
                Write-Verbose "Failed to retrieve with cmdlet, falling back to direct Graph API..."
            }
        }
        
        if (-not $existing) {
            try {
                $response = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/beta/deviceManagement/virtualEndpoint/provisioningPolicies?`$filter=displayName eq '$DisplayName'" -ErrorAction Stop
                $existing = $response.value | Select-Object -First 1
            }
            catch {
                Write-Verbose "Failed to retrieve existing policy via Graph API"
            }
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
                provisioningType   = "dedicated"
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
                userSettingsPersistenceEnabled = $false
                userSettingsPersistenceConfiguration = @{
                    userSettingsPersistenceEnabled        = $false
                    userSettingsPersistenceStorageSizeCategory = "sixteenGB"
                }
                autopilotConfiguration = $null
            }

            # Try cmdlet first, then direct API
            if (Get-Command New-MgBetaDeviceManagementVirtualEndpointProvisioningPolicy -ErrorAction SilentlyContinue) {
                try {
                    $existing = New-MgBetaDeviceManagementVirtualEndpointProvisioningPolicy -BodyParameter $params -ErrorAction Stop
                }
                catch {
                    Write-Verbose "Failed to create with cmdlet, trying direct Graph API..."
                    Write-Verbose "Provisioning payload: $($params | ConvertTo-Json -Depth 6)"
                    $createResponse = Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/beta/deviceManagement/virtualEndpoint/provisioningPolicies" -Body ($params | ConvertTo-Json) -ContentType "application/json" -ErrorAction Stop
                    $existing = $createResponse
                }
            }
            else {
                Write-Verbose "Provisioning payload: $($params | ConvertTo-Json -Depth 6)"
                $createResponse = Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/beta/deviceManagement/virtualEndpoint/provisioningPolicies" -Body ($params | ConvertTo-Json) -ContentType "application/json" -ErrorAction Stop
                $existing = $createResponse
            }
            
            Write-Verbose "Provisioning policy created: $DisplayName with ID: $($existing.id)"
        }
        else {
            Write-Host "Provisioning Policy already exists: $DisplayName" -ForegroundColor Green
        }

        # Assign via /assign endpoint with pre-validation of group IDs
        # Graph /assign replaces all assignments; always merge existing + new before sending
        $validGroupIds = @()
        foreach ($gid in $AssignGroupIds | Where-Object { $_ -and (-not [string]::IsNullOrWhiteSpace($_)) }) {
            try {
                # Try Get-MgGroup first, but fall back to direct Graph API if module is not available
                $groupFound = $false
                if (Get-Command Get-MgGroup -ErrorAction SilentlyContinue) {
                    try {
                        $g = Get-MgGroup -GroupId $gid -ErrorAction Stop
                        $validGroupIds += $g.Id
                        $groupFound = $true
                    }
                    catch {
                        Write-Verbose "Get-MgGroup failed for $gid, trying direct Graph API: $_"
                    }
                }
                
                # Fallback to direct Graph API
                if (-not $groupFound) {
                    $groupResponse = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/groups/$gid" -ErrorAction Stop
                    $validGroupIds += $groupResponse.id
                }
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

        # Build complete assignment list: existing + new (avoiding duplicates)
        $allGroupIds = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
        
        foreach ($gid in $existingGroupIds) {
            $allGroupIds.Add($gid) | Out-Null
        }
        
        foreach ($gid in $validGroupIds) {
            $allGroupIds.Add($gid) | Out-Null
        }

        # Build assignments array with complete list
        $assignments = @()
        foreach ($gid in $allGroupIds) {
            $assignments += @{
                id     = $null
                target = @{ groupId = $gid }
            }
        }

        $assignParams = @{ assignments = $assignments }
        $assignJson = $assignParams | ConvertTo-Json -Depth 4

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

# Import required Graph modules
Write-Verbose "Importing Microsoft Graph modules..."
try {
    Import-Module Microsoft.Graph.Beta.DeviceManagement.Administration -ErrorAction Stop
}
catch {
    Write-Verbose "Beta DeviceManagement module not found, trying standard module..."
    try {
        Import-Module Microsoft.Graph.DeviceManagement.Administration -ErrorAction Stop
    }
    catch {
        Write-Warning "DeviceManagement.Administration module not available. Using direct Graph API calls only."
    }
}

# Note: Do not explicitly import Microsoft.Graph.Groups or Microsoft.Graph.Authentication here
# as Connect-MgGraph already loads them and causes assembly conflicts

# Use Beta profile when available
$selectProfileCmd = Get-Command Select-MgProfile -ErrorAction SilentlyContinue
if ($selectProfileCmd) {
    try {
        Select-MgProfile -Name beta -ErrorAction Stop
    }
    catch {
        Write-Warning "Could not select beta profile: $_"
    }
}
else {
    Write-Verbose "Select-MgProfile not found. Using beta endpoints directly."
}

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

# Get available Cloud PC service plans from Graph
Write-Host "`nRetrieving available Windows 365 Cloud PC service plans..." -ForegroundColor Cyan
try {
    $servicePlans = @()

    if (Get-Command Get-MgDeviceManagementVirtualEndpointServicePlan -ErrorAction SilentlyContinue) {
        try {
            $servicePlans = Get-MgDeviceManagementVirtualEndpointServicePlan -All -ErrorAction Stop
        }
        catch {
            # Fallback without -All if not supported
            $servicePlans = Get-MgDeviceManagementVirtualEndpointServicePlan -ErrorAction Stop
        }
    }
    else {
        # Fallback to direct Graph call (beta) if cmdlet is unavailable, with paging
        $servicePlans = Get-AllGraphItems -Uri "https://graph.microsoft.com/beta/deviceManagement/virtualEndpoint/servicePlans"
    }
    
    if (-not $servicePlans -or $servicePlans.Count -eq 0) {
        Write-Warning "No Cloud PC service plans found. This may indicate insufficient permissions or no available plans."
        throw "Unable to retrieve Cloud PC service plans"
    }
    
    # Normalize objects to have DisplayName
    $servicePlans = $servicePlans | ForEach-Object {
        if ($_ -is [string]) { [pscustomobject]@{ DisplayName = $_ } }
        else { $_ }
    }

    # Filter out Business and Frontline SKUs
    $servicePlans = $servicePlans | Where-Object { 
        $_.DisplayName -notmatch 'Business' -and $_.DisplayName -notmatch 'Frontline' 
    }

    if (-not $servicePlans -or $servicePlans.Count -eq 0) {
        Write-Warning "No Cloud PC service plans found after filtering. Check available SKUs."
        throw "No Enterprise Cloud PC service plans available"
    }

    # Sort by display name for consistent ordering
    $servicePlans = $servicePlans | Sort-Object DisplayName
    $CloudPCType = $servicePlans.DisplayName
    
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
        Write-Host ("{0,2}. {1}" -f ($i + 1), $CloudPCType[$i])
    }
    Write-Host ""
    
    $Windows365CloudPCTypeVariable = Get-ValidChoice -Min 1 -Max $CloudPCType.Count
}
else {
    if ($CloudPCTypeChoice -lt 1 -or $CloudPCTypeChoice -gt $CloudPCType.Count) {
        Write-Error "Invalid CloudPCTypeChoice parameter. Must be between 1 and $($CloudPCType.Count)."
        Write-Host "Available plans:" -ForegroundColor Yellow
        for ($i = 0; $i -lt $CloudPCType.Count; $i++) {
            Write-Host ("{0,2}. {1}" -f ($i + 1), $CloudPCType[$i])
        }
        throw "CloudPCTypeChoice out of range"
    }
    $Windows365CloudPCTypeVariable = $CloudPCTypeChoice
    Write-Verbose "Using parameter-specified Cloud PC type choice: $CloudPCTypeChoice"
}

$Windows365CloudPCType = $CloudPCType[$Windows365CloudPCTypeVariable - 1]

# Get supported region groups for Cloud PC
Write-Host "`nRetrieving supported Windows 365 region groups..." -ForegroundColor Cyan
try {
    $supportedRegions = @()

    if (Get-Command Get-MgDeviceManagementVirtualEndpointSupportedRegion -ErrorAction SilentlyContinue) {
        try {
            $supportedRegions = Get-MgDeviceManagementVirtualEndpointSupportedRegion -All -ErrorAction Stop
        }
        catch {
            # Fallback without -All if not supported
            $supportedRegions = Get-MgDeviceManagementVirtualEndpointSupportedRegion -ErrorAction Stop
        }
    }
    else {
        # Fallback to direct Graph call (beta) if cmdlet is unavailable, with paging
        # Filter for Windows 365 supported regions and select relevant fields
        $uri = "https://graph.microsoft.com/beta/deviceManagement/virtualEndpoint/supportedRegions?`$filter=supportedSolution eq 'windows365'&`$select=id,displayName,regionStatus,supportedSolution,regionGroup,cloudDevicePlatformSupported,geographicLocationType"
        $supportedRegions = Get-AllGraphItems -Uri $uri
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

    # Normalize objects to ensure RegionGroup, RegionName, and DisplayName are present (use Graph field names when available)
    $supportedRegions = $supportedRegions | ForEach-Object {
        if ($_ -is [string]) {
            [pscustomobject]@{
                RegionGroup = $_
                RegionName  = $_
                DisplayName = $_
            }
        }
        else {
            # Graph API returns: regionGroup, displayName (no regionName field exists)
            # Try to get values from properties (lowercase from Graph API, or PascalCase as fallback)
            $rg = $_.regionGroup
            if (-not $rg) { $rg = $_.RegionGroup }
            
            $dn = $_.displayName
            if (-not $dn) { $dn = $_.DisplayName }
            
            # RegionName should be the same as DisplayName (the actual region like "brazilsouth")
            # There is no separate regionName field in the Graph API
            $rn = $dn
            
            # Final fallbacks
            if (-not $rn) { $rn = $rg }
            if (-not $dn) { $dn = $rg }
            
            [pscustomobject]@{
                RegionGroup = $rg
                RegionName  = $rn
                DisplayName = $dn
            }
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
    # Step 1: Get unique regionGroup values
    $uniqueGroups = @()
    foreach ($region in $supportedRegions) {
        if ($region.RegionGroup -and $region.RegionGroup -notin $uniqueGroups) {
            $uniqueGroups += $region.RegionGroup
        }
    }
    $uniqueGroups = $uniqueGroups | Sort-Object

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
    $regionsInGroup = $supportedRegions | Where-Object { $_.RegionGroup -eq $selectedGroupValue } | Sort-Object DisplayName
    
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
        
        # Fallback to device images endpoint
        if (Get-Command Get-MgDeviceManagementVirtualEndpointDeviceImage -ErrorAction SilentlyContinue) {
            try {
                $images = Get-MgDeviceManagementVirtualEndpointDeviceImage -All -ErrorAction Stop
            }
            catch {
                # Fallback without -All if not supported
                $images = Get-MgDeviceManagementVirtualEndpointDeviceImage -ErrorAction Stop
            }
        }
        else {
            # Fallback to direct Graph call (beta) if cmdlet is unavailable, with paging
            try {
                $images = Get-AllGraphItems -Uri "https://graph.microsoft.com/beta/deviceManagement/virtualEndpoint/deviceImages"
            }
            catch {
                Write-Verbose "Beta endpoint failed, trying v1 endpoint..."
                $images = Get-AllGraphItems -Uri "https://graph.microsoft.com/v1.0/deviceManagement/virtualEndpoint/deviceImages"
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

        Write-Verbose "Found $($availableImages.Count) available device images"

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
    @{ DisplayName = "English (United States)"; Code = "en-US" },
    @{ DisplayName = "English (United Kingdom)"; Code = "en-GB" },
    @{ DisplayName = "Spanish (Spain)"; Code = "es-ES" },
    @{ DisplayName = "Spanish (Mexico)"; Code = "es-MX" },
    @{ DisplayName = "French (France)"; Code = "fr-FR" },
    @{ DisplayName = "French (Canada)"; Code = "fr-CA" },
    @{ DisplayName = "German (Germany)"; Code = "de-DE" },
    @{ DisplayName = "Italian (Italy)"; Code = "it-IT" },
    @{ DisplayName = "Portuguese (Brazil)"; Code = "pt-BR" },
    @{ DisplayName = "Portuguese (Portugal)"; Code = "pt-PT" },
    @{ DisplayName = "Dutch (Netherlands)"; Code = "nl-NL" },
    @{ DisplayName = "Russian (Russia)"; Code = "ru-RU" },
    @{ DisplayName = "Japanese (Japan)"; Code = "ja-JP" },
    @{ DisplayName = "Korean (Korea)"; Code = "ko-KR" },
    @{ DisplayName = "Chinese (Simplified)"; Code = "zh-CN" },
    @{ DisplayName = "Chinese (Traditional)"; Code = "zh-TW" },
    @{ DisplayName = "Arabic (Saudi Arabia)"; Code = "ar-SA" },
    @{ DisplayName = "Hindi (India)"; Code = "hi-IN" },
    @{ DisplayName = "Polish (Poland)"; Code = "pl-PL" },
    @{ DisplayName = "Turkish (Turkey)"; Code = "tr-TR" }
)

if ([string]::IsNullOrWhiteSpace($Language)) {
    Write-Host "`nChoose your Windows 11 language by selecting its corresponding number:" -ForegroundColor Green
    for ($i = 0; $i -lt $SupportedLanguages.Count; $i++) {
        $selected = if ($SupportedLanguages[$i].Code -eq "en-GB") { " [Default]" } else { "" }
        Write-Host ("{0,2}. {1}{2}" -f ($i + 1), $SupportedLanguages[$i].DisplayName, $selected)
    }
    Write-Host ""

    $languageChoice = Get-ValidChoice -Min 1 -Max $SupportedLanguages.Count
    $SelectedLanguage = $SupportedLanguages[$languageChoice - 1].Code
    Write-Host "Selected language: $($SupportedLanguages[$languageChoice - 1].DisplayName)" -ForegroundColor Cyan
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

# Groups - per Cloud PC Type
$UserGroupName  = "GRP_Users_$Windows365CloudPCType"
$AdminGroupName = "GRP_Admins_$Windows365CloudPCType"

Write-Verbose "Creating/retrieving groups..."
$GroupIDUser  = Get-OrCreateGroup -DisplayName $UserGroupName  -Description "Contains $Windows365CloudPCType Users"
$GroupIDAdmin = Get-OrCreateGroup -DisplayName $AdminGroupName -Description "Contains $Windows365CloudPCType Admins"

# Allow time for group replication before policy assignment (prov policy /assign is more eventual)
Write-Verbose "Waiting for group replication to complete..."
Start-Sleep -Seconds 10

# User Settings
Write-Verbose "Creating/retrieving Cloud PC User Settings..."
$cloudPcAdminSettingId = Get-OrCreateCloudPcUserSetting -DisplayName "W365_AdminSettings" -LocalAdminEnabled $true  -TargetGroupId $GroupIDAdmin
$cloudPcUserSettingId  = Get-OrCreateCloudPcUserSetting -DisplayName "W365_UserSettings"  -LocalAdminEnabled $false -TargetGroupId $GroupIDUser

# Provisioning Policy - per Region
Write-Verbose "Creating/retrieving Provisioning Policy for region..."

# Human-friendly region name only (skip group prefix to avoid awkward strings)
$policyRegionNameRaw  = if ($SelectedRegionDisplayName) { $SelectedRegionDisplayName } else { $SelectedRegionName -replace '[_-]', ' ' -replace '(?<=.)([A-Z])',' $1' }
$policyRegionName     = (Get-Culture).TextInfo.ToTitleCase($policyRegionNameRaw.ToLower().Trim())

$ProvisioningPolicyName = "$policyRegionName-W365-Enterprise-Provisioning Policy"
$cloudPcProvisioningPolicyId = Get-OrCreateProvisioningPolicy -DisplayName $ProvisioningPolicyName -AssignGroupIds @($GroupIDAdmin, $GroupIDUser) -RegionGroup $SelectedRegionGroup -CountryRegion $SelectedCountryRegion -ImageId $SelectedImage.Id -ImageDisplayName $SelectedImage.DisplayName -Language $SelectedLanguage

Write-Host "`nDone ✅" -ForegroundColor Green
Write-Host "Remember to assign the correct Windows 365 license to the groups created:" -ForegroundColor Yellow
Write-Host " - $UserGroupName" -ForegroundColor Yellow
Write-Host " - $AdminGroupName" -ForegroundColor Yellow

# Cleanup
Write-Verbose "Disconnecting from Microsoft Graph..."
#Disconnect-MgGraph | Out-Null
Write-Host "`nDisconnected from Microsoft Graph." -ForegroundColor Cyan
