#Requires -Version 7.0
<#
.SYNOPSIS
    Windows 365 Deployment Wizard — WPF graphical front-end.
.DESCRIPTION
    Step-by-step wizard (7 pages):
      Connect → SKU → Region → Image → Language → Windows Update → Summary → Results
    New in this version:
      - Automatic Windows 365 licence assignment to the licensing group
      - Windows Update for Business ring creation with three profile presets
      - All list boxes use proper Grid layouts so they are fully scrollable
.NOTES
    Requires Microsoft.Graph.Authentication (auto-installed if missing).
    Must run on Windows (WPF). STA thread is handled automatically.
    Additional Graph scopes required:
      LicenseAssignment.ReadWrite.All   — group-based licence assignment
      DeviceManagementConfiguration.ReadWrite.All — Windows Update rings
#>

# ── STA thread required for WPF ──────────────────────────────────────────────
if ([System.Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {
    $rs = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace()
    $rs.ApartmentState = [System.Threading.ApartmentState]::STA
    $rs.ThreadOptions  = [System.Management.Automation.Runspaces.PSThreadOptions]::ReuseThread
    $rs.Open()
    $ps = [System.Management.Automation.PowerShell]::Create()
    $ps.Runspace = $rs
    [void]$ps.AddScript(". '$($MyInvocation.MyCommand.Path)'")
    [void]$ps.Invoke()
    $ps.Dispose(); $rs.Close()
    return
}

Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase

# ════════════════════════════════════════════════════════════════════════════
#  GRAPH HELPER FUNCTIONS
# ════════════════════════════════════════════════════════════════════════════

function Install-GraphModuleIfNeeded {
    if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication)) {
        Install-Module -Name Microsoft.Graph.Authentication -AllowClobber -Force -ErrorAction Stop
    }
}

function Get-AllGraphItems {
    param([Parameter(Mandatory)][string]$Uri)
    $items    = @()
    $nextLink = $Uri
    while ($nextLink) {
        $response = Invoke-MgGraphRequest -Method GET -Uri $nextLink -ErrorAction Stop
        if ($response.value) { $items += $response.value }
        $nextLink = $response.'@odata.nextLink'
    }
    return $items
}

function Get-SkuMetrics {
    param([Parameter(Mandatory)][string]$DisplayName)
    if ($DisplayName -match '(?<vcpu>\d+)vCPU/(?<ram>\d+)GB/(?<storage>[\d\.]+)(?<unit>TB|GB)') {
        $storageGb = if ($Matches['unit'] -eq 'TB') { [double]$Matches['storage'] * 1024 } else { [double]$Matches['storage'] }
        return [pscustomobject]@{ Vcpu = [int]$Matches['vcpu']; RamGb = [int]$Matches['ram']; StorageGb = [int][math]::Round($storageGb, 0) }
    }
    return $null
}

function Test-IsCopilotEligibleSku {
    param([Parameter(Mandatory)][string]$DisplayName)
    $m = Get-SkuMetrics -DisplayName $DisplayName
    return ($m -and $m.Vcpu -ge 8 -and $m.RamGb -ge 32 -and $m.StorageGb -ge 256)
}

function Get-OrCreateGroup {
    param([Parameter(Mandatory)][string]$DisplayName, [Parameter(Mandatory)][string]$Description)
    $response = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/groups?`$filter=displayName eq '$DisplayName'" -ErrorAction Stop
    $existing = $response.value | Select-Object -First 1
    if ($existing) { return @{ Id = $existing.Id; Created = $false } }
    $mailNick = "grp-" + [guid]::NewGuid().ToString("N").Substring(0, 10)
    $params   = @{ DisplayName = $DisplayName; MailEnabled = $false; MailNickname = $mailNick; SecurityEnabled = $true; Description = $Description }
    $created  = Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/v1.0/groups" -Body ($params | ConvertTo-Json) -ContentType "application/json" -ErrorAction Stop
    return @{ Id = $created.id; Created = $true }
}

function Get-OrCreateDynamicDeviceGroup {
    param(
        [Parameter(Mandatory)][string]$DisplayName,
        [Parameter(Mandatory)][string]$Description,
        [string]$MembershipRule = '(device.displayName -startsWith "CPC-")'
    )
    $response = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/groups?`$filter=displayName eq '$DisplayName'" -ErrorAction Stop
    $existing = $response.value | Select-Object -First 1
    if ($existing) { return @{ Id = $existing.Id; Created = $false } }
    $mailNick = "grp-" + [guid]::NewGuid().ToString("N").Substring(0, 10)
    $params   = @{
        displayName = $DisplayName; mailEnabled = $false; mailNickname = $mailNick; securityEnabled = $true
        description = $Description; groupTypes = @("DynamicMembership"); membershipRuleProcessingState = "On"
        membershipRule = $MembershipRule
    }
    $created = Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/v1.0/groups" -Body ($params | ConvertTo-Json) -ContentType "application/json" -ErrorAction Stop
    return @{ Id = $created.id; Created = $true }
}

function Get-OrCreateCloudPcUserSetting {
    param(
        [Parameter(Mandatory)][string]$DisplayName,
        [Parameter(Mandatory)][bool]$LocalAdminEnabled,
        [Parameter(Mandatory)][string]$TargetGroupId
    )
    $response  = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/beta/deviceManagement/virtualEndpoint/userSettings?`$filter=displayName eq '$DisplayName'" -ErrorAction Stop
    $existing  = $response.value | Select-Object -First 1
    $wasCreated = $false
    $settingAlreadyExisted = $null -ne $existing

    if (-not $existing) {
        $params = @{
            displayName = $DisplayName; localAdminEnabled = $LocalAdminEnabled; resetEnabled = $true
            restorePointSetting = @{ userRestoreEnabled = $true; frequencyInHours = 6 }
            crossRegionDisasterRecoverySetting = @{
                crossRegionDisasterRecoveryEnabled     = $false
                maintainCrossRegionRestorePointEnabled = $true
                disasterRecoveryNetworkSetting         = $null
                disasterRecoveryType                   = "notConfigured"
                userInitiatedDisasterRecoveryAllowed   = $false
            }
            notificationSetting = @{ restartPromptsDisabled = $false }
        }
        $existing   = Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/beta/deviceManagement/virtualEndpoint/userSettings" -Body ($params | ConvertTo-Json -Depth 10) -ContentType "application/json" -ErrorAction Stop
        $wasCreated = $true
    }

    $existingGroupIds = @()
    $gotAssignments   = $false
    try {
        $expanded = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/beta/deviceManagement/virtualEndpoint/userSettings/$($existing.Id)?`$expand=assignments" -ErrorAction SilentlyContinue
        if ($expanded -and $expanded.assignments) { $existingGroupIds = @($expanded.assignments | ForEach-Object { $_.target.groupId }); $gotAssignments = $true }
    } catch {}

    if (-not $gotAssignments -and $settingAlreadyExisted) { return @{ Id = $existing.Id; Created = $wasCreated } }

    $allGroupIds = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($gid in $existingGroupIds) { if ($gid) { [void]$allGroupIds.Add($gid) } }
    [void]$allGroupIds.Add($TargetGroupId)

    $assignments = @($allGroupIds | ForEach-Object { @{ id = $null; target = @{ groupId = $_ } } })
    Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/beta/deviceManagement/virtualEndpoint/userSettings/$($existing.Id)/assign" -Body (@{ Assignments = $assignments } | ConvertTo-Json -Depth 6) -ContentType "application/json" -ErrorAction Stop | Out-Null
    return @{ Id = $existing.Id; Created = $wasCreated }
}

function Get-OrCreateProvisioningPolicy {
    param(
        [Parameter(Mandatory)][string]$DisplayName,
        [Parameter(Mandatory)][string[]]$AssignGroupIds,
        [Parameter(Mandatory)][string]$RegionGroup,
        [Parameter(Mandatory)][string]$CountryRegion,
        [Parameter(Mandatory)][string]$ImageId,
        [Parameter(Mandatory)][string]$ImageDisplayName,
        [string]$Language = "en-GB",
        [string]$ProvisioningType = "dedicated",
        [bool]$UserSettingsPersistence = $false,
        [string]$ServicePlanId = $null
    )
    $response  = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/beta/deviceManagement/virtualEndpoint/provisioningPolicies?`$filter=displayName eq '$DisplayName'" -ErrorAction Stop
    $existing  = $response.value | Select-Object -First 1
    $wasCreated = $false
    $policyAlreadyExisted = $null -ne $existing

    if (-not $existing) {
        $params = @{
            displayName = $DisplayName; description = ""; provisioningType = $ProvisioningType
            userExperienceType = "cloudPc"; managedBy = "windows365"
            imageId = $ImageId; imageDisplayName = $ImageDisplayName; imageType = "gallery"
            microsoftManagedDesktop = @{ type = "notManaged"; profile = "" }
            enableSingleSignOn = $true
            domainJoinConfigurations = @(@{ type = "azureADJoin"; regionGroup = $RegionGroup; regionName = $CountryRegion })
            windowsSettings = @{ language = $Language }
            cloudPcNamingTemplate = $null; scopeIds = @("0")
            autopatch = @{ autopatchGroupId = $null }
            userSettingsPersistenceEnabled = $UserSettingsPersistence
            userSettingsPersistenceConfiguration = @{ userSettingsPersistenceEnabled = $UserSettingsPersistence; userSettingsPersistenceStorageSizeCategory = "sixteenGB" }
            autopilotConfiguration = $null
        }
        try {
            $existing = Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/beta/deviceManagement/virtualEndpoint/provisioningPolicies" -Body ($params | ConvertTo-Json -Depth 6) -ContentType "application/json" -ErrorAction Stop
        } catch {
            if ($params.windowsSettings.language -ne "en-GB") {
                $params.windowsSettings.language = "en-GB"
                $existing = Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/beta/deviceManagement/virtualEndpoint/provisioningPolicies" -Body ($params | ConvertTo-Json -Depth 6) -ContentType "application/json" -ErrorAction Stop
            } else { throw }
        }
        $wasCreated = $true
    }

    if ($ServicePlanId) { return @{ Id = $existing.id; Created = $wasCreated } }

    $validGroupIds = @()
    foreach ($gid in ($AssignGroupIds | Where-Object { $_ })) {
        try { $g = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/groups/$gid" -ErrorAction Stop; $validGroupIds += $g.id } catch {}
    }
    if ($validGroupIds.Count -eq 0) { return @{ Id = $existing.id; Created = $wasCreated } }

    $existingGroupIds = @()
    $gotAssignments   = $false
    try {
        $expanded = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/beta/deviceManagement/virtualEndpoint/provisioningPolicies/$($existing.id)?`$expand=assignments" -ErrorAction Stop
        if ($expanded.PSObject.Properties['assignments'] -and $expanded.assignments) { $existingGroupIds = @($expanded.assignments | ForEach-Object { $_.target.groupId }) }
        $gotAssignments = $true
    } catch {}

    if (-not $gotAssignments -and $policyAlreadyExisted) { return @{ Id = $existing.id; Created = $wasCreated } }

    $allGroupIds = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($gid in $existingGroupIds) { if ($gid) { [void]$allGroupIds.Add($gid) } }
    foreach ($gid in $validGroupIds)    { if ($gid) { [void]$allGroupIds.Add($gid) } }

    $assignments = @($allGroupIds | ForEach-Object { @{ id = $null; target = @{ groupId = $_ } } })
    Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/beta/deviceManagement/virtualEndpoint/provisioningPolicies/$($existing.id)/assign" -Body (@{ assignments = $assignments } | ConvertTo-Json -Depth 10) -ContentType "application/json" -ErrorAction Stop | Out-Null
    return @{ Id = $existing.id; Created = $wasCreated }
}

# ── Automatic licence assignment ──────────────────────────────────────────────
# Finds the subscribed SKU that contains the selected service plan and assigns
# it to the licensing group via group-based licensing.
# Requires: LicenseAssignment.ReadWrite.All + Azure AD P1 on the tenant.

function Set-GroupLicense {
    param(
        [Parameter(Mandatory)][string]$GroupId,
        [Parameter(Mandatory)][string]$ServicePlanId
    )
    $skus = Get-AllGraphItems -Uri "https://graph.microsoft.com/v1.0/subscribedSkus?`$select=skuId,skuPartNumber,servicePlans,consumedUnits,prepaidUnits"

    $matchingSku = $skus | Where-Object {
        $_.servicePlans | Where-Object { $_.servicePlanId -eq $ServicePlanId }
    } | Select-Object -First 1

    if (-not $matchingSku) {
        throw "No subscribed SKU found containing service plan '$ServicePlanId'. Ensure the Windows 365 licence is purchased in this tenant."
    }

    $available = $matchingSku.prepaidUnits.enabled - $matchingSku.consumedUnits
    $warning   = if ($available -le 0) { " ⚠️ No available units — assignment may fail for users." } else { "" }

    $payload = @{
        addLicenses    = @(@{ skuId = $matchingSku.skuId; disabledPlans = @() })
        removeLicenses = @()
    }
    Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/v1.0/groups/$GroupId/assignLicense" `
        -Body ($payload | ConvertTo-Json -Depth 5) -ContentType "application/json" -ErrorAction Stop | Out-Null

    return @{ SkuPartNumber = $matchingSku.skuPartNumber; Warning = $warning }
}

# ── Windows Update for Business ring ─────────────────────────────────────────
# Creates (or reuses) a WUfB configuration policy assigned to the devices group.
# Requires: DeviceManagementConfiguration.ReadWrite.All

function Get-OrCreateUpdateRing {
    param(
        [Parameter(Mandatory)][string]$DisplayName,
        [Parameter(Mandatory)][string]$DeviceGroupId,
        [string]$Profile = 'standard'   # standard | recommended | deferred
    )

    $qualityDays, $featureDays = switch ($Profile) {
        'recommended' { 7,  30  }
        'deferred'    { 14, 180 }
        default       { 7,  0   }   # standard — feature updates follow Windows as a Service cadence
    }

    $response = Invoke-MgGraphRequest -Method GET `
        -Uri "https://graph.microsoft.com/beta/deviceManagement/deviceConfigurations?`$filter=displayName eq '$DisplayName'" `
        -ErrorAction Stop
    $existing = $response.value | Select-Object -First 1
    if ($existing) { return @{ Id = $existing.id; Created = $false } }

    $params = @{
        "@odata.type"                      = "#microsoft.graph.windowsUpdateForBusinessConfiguration"
        displayName                        = $DisplayName
        description                        = "Windows Update ring for Windows 365 Cloud PCs — created by Deploy-Windows365-GUI"
        microsoftUpdateServiceAllowed      = $true
        driversExcluded                    = $false
        qualityUpdatesDeferralPeriodInDays = $qualityDays
        featureUpdatesDeferralPeriodInDays = $featureDays
        featureUpdatesRollbackWindowInDays = 10
        automaticUpdateMode                = "autoInstallAtMaintenanceTime"
        businessReadyUpdatesOnly           = "userDefined"
        skipChecksBeforeRestart            = $false
        userPauseAccess                    = "enabled"
        userWindowsUpdateScanAccess        = "enabled"
        updateNotificationLevel            = "defaultNotifications"
        deliveryOptimizationMode           = "httpOnly"
        prereleaseFeatures                 = "settingsOnly"
        roleScopeTagIds                    = @()
    }

    $created = Invoke-MgGraphRequest -Method POST `
        -Uri "https://graph.microsoft.com/beta/deviceManagement/deviceConfigurations" `
        -Body ($params | ConvertTo-Json -Depth 5) -ContentType "application/json" -ErrorAction Stop

    $assignPayload = @{
        assignments = @(@{
            target = @{
                "@odata.type" = "#microsoft.graph.groupAssignmentTarget"
                groupId       = $DeviceGroupId
            }
        })
    }
    Invoke-MgGraphRequest -Method POST `
        -Uri "https://graph.microsoft.com/beta/deviceManagement/deviceConfigurations/$($created.id)/assign" `
        -Body ($assignPayload | ConvertTo-Json -Depth 5) -ContentType "application/json" -ErrorAction Stop | Out-Null

    return @{ Id = $created.id; Created = $true }
}

# ════════════════════════════════════════════════════════════════════════════
#  SUPPORTED LANGUAGES
# ════════════════════════════════════════════════════════════════════════════

$script:SupportedLanguages = @(
    [pscustomobject]@{ DisplayName = "Arabic (Saudi Arabia)";    Code = "ar-SA" }
    [pscustomobject]@{ DisplayName = "Bulgarian (Bulgaria)";     Code = "bg-BG" }
    [pscustomobject]@{ DisplayName = "Chinese (Simplified)";     Code = "zh-CN" }
    [pscustomobject]@{ DisplayName = "Chinese (Traditional)";    Code = "zh-TW" }
    [pscustomobject]@{ DisplayName = "Croatian (Croatia)";       Code = "hr-HR" }
    [pscustomobject]@{ DisplayName = "Czech (Czech Republic)";   Code = "cs-CZ" }
    [pscustomobject]@{ DisplayName = "Danish (Denmark)";         Code = "da-DK" }
    [pscustomobject]@{ DisplayName = "Dutch (Netherlands)";      Code = "nl-NL" }
    [pscustomobject]@{ DisplayName = "English (Australia)";      Code = "en-AU" }
    [pscustomobject]@{ DisplayName = "English (Ireland)";        Code = "en-IE" }
    [pscustomobject]@{ DisplayName = "English (New Zealand)";    Code = "en-NZ" }
    [pscustomobject]@{ DisplayName = "English (United Kingdom)"; Code = "en-GB" }
    [pscustomobject]@{ DisplayName = "English (United States)";  Code = "en-US" }
    [pscustomobject]@{ DisplayName = "Estonian (Estonia)";       Code = "et-EE" }
    [pscustomobject]@{ DisplayName = "Finnish (Finland)";        Code = "fi-FI" }
    [pscustomobject]@{ DisplayName = "French (Canada)";          Code = "fr-CA" }
    [pscustomobject]@{ DisplayName = "French (France)";          Code = "fr-FR" }
    [pscustomobject]@{ DisplayName = "German (Germany)";         Code = "de-DE" }
    [pscustomobject]@{ DisplayName = "Greek (Greece)";           Code = "el-GR" }
    [pscustomobject]@{ DisplayName = "Hebrew (Israel)";          Code = "he-IL" }
    [pscustomobject]@{ DisplayName = "Hindi (India)";            Code = "hi-IN" }
    [pscustomobject]@{ DisplayName = "Hungarian (Hungary)";      Code = "hu-HU" }
    [pscustomobject]@{ DisplayName = "Italian (Italy)";          Code = "it-IT" }
    [pscustomobject]@{ DisplayName = "Japanese (Japan)";         Code = "ja-JP" }
    [pscustomobject]@{ DisplayName = "Korean (Korea)";           Code = "ko-KR" }
    [pscustomobject]@{ DisplayName = "Latvian (Latvia)";         Code = "lv-LV" }
    [pscustomobject]@{ DisplayName = "Lithuanian (Lithuania)";   Code = "lt-LT" }
    [pscustomobject]@{ DisplayName = "Norwegian (Bokmal)";       Code = "nb-NO" }
    [pscustomobject]@{ DisplayName = "Polish (Poland)";          Code = "pl-PL" }
    [pscustomobject]@{ DisplayName = "Portuguese (Brazil)";      Code = "pt-BR" }
    [pscustomobject]@{ DisplayName = "Portuguese (Portugal)";    Code = "pt-PT" }
    [pscustomobject]@{ DisplayName = "Romanian (Romania)";       Code = "ro-RO" }
    [pscustomobject]@{ DisplayName = "Russian (Russia)";         Code = "ru-RU" }
    [pscustomobject]@{ DisplayName = "Serbian (Latin)";          Code = "sr-Latn-RS" }
    [pscustomobject]@{ DisplayName = "Slovak (Slovakia)";        Code = "sk-SK" }
    [pscustomobject]@{ DisplayName = "Slovenian (Slovenia)";     Code = "sl-SI" }
    [pscustomobject]@{ DisplayName = "Spanish (Mexico)";         Code = "es-MX" }
    [pscustomobject]@{ DisplayName = "Spanish (Spain)";          Code = "es-ES" }
    [pscustomobject]@{ DisplayName = "Swedish (Sweden)";         Code = "sv-SE" }
    [pscustomobject]@{ DisplayName = "Thai (Thailand)";          Code = "th-TH" }
    [pscustomobject]@{ DisplayName = "Turkish (Turkey)";         Code = "tr-TR" }
    [pscustomobject]@{ DisplayName = "Ukrainian (Ukraine)";      Code = "uk-UA" }
)

# ════════════════════════════════════════════════════════════════════════════
#  XAML  (8 pages — index 0-7)
#  0 Connect   1 SKU   2 Region   3 Image   4 Language
#  5 Windows Update   6 Summary   7 Results
# ════════════════════════════════════════════════════════════════════════════

[xml]$xaml = @'
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="Windows 365 Deployment Wizard"
    Width="800" Height="620"
    MinWidth="800" MinHeight="620"
    WindowStartupLocation="CenterScreen"
    ResizeMode="CanMinimize"
    FontFamily="Segoe UI"
    FontSize="13"
    Background="White">

    <Window.Resources>

        <Style x:Key="PrimaryBtn" TargetType="Button">
            <Setter Property="Background"      Value="#0078D4"/>
            <Setter Property="Foreground"      Value="White"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Padding"         Value="18,7"/>
            <Setter Property="Cursor"          Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border Background="{TemplateBinding Background}" CornerRadius="3" Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True"><Setter Property="Background" Value="#106EBE"/></Trigger>
                            <Trigger Property="IsEnabled"   Value="False"><Setter Property="Background" Value="#C8C8C8"/><Setter Property="Foreground" Value="#999"/></Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <Style x:Key="SecondaryBtn" TargetType="Button">
            <Setter Property="Background"      Value="White"/>
            <Setter Property="Foreground"      Value="#0078D4"/>
            <Setter Property="BorderBrush"     Value="#0078D4"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Padding"         Value="18,7"/>
            <Setter Property="Cursor"          Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}" CornerRadius="3" Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True"><Setter Property="Background" Value="#EBF3FB"/></Trigger>
                            <Trigger Property="IsEnabled"   Value="False"><Setter Property="BorderBrush" Value="#CCC"/><Setter Property="Foreground" Value="#CCC"/></Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <Style x:Key="DeployBtn" TargetType="Button">
            <Setter Property="Background"      Value="#107C10"/>
            <Setter Property="Foreground"      Value="White"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Padding"         Value="20,7"/>
            <Setter Property="FontWeight"      Value="SemiBold"/>
            <Setter Property="Cursor"          Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border Background="{TemplateBinding Background}" CornerRadius="3" Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True"><Setter Property="Background" Value="#0B6A0B"/></Trigger>
                            <Trigger Property="IsEnabled"   Value="False"><Setter Property="Background" Value="#C8C8C8"/><Setter Property="Foreground" Value="#999"/></Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <Style TargetType="ListBoxItem">
            <Setter Property="Padding" Value="10,6"/>
        </Style>

        <Style TargetType="TextBox">
            <Setter Property="Padding"     Value="6,5"/>
            <Setter Property="BorderBrush" Value="#D0D0D0"/>
        </Style>

        <Style TargetType="RadioButton">
            <Setter Property="Margin" Value="0,4,0,4"/>
        </Style>

    </Window.Resources>

    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="72"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="56"/>
        </Grid.RowDefinitions>

        <!-- ═══ HEADER ═══ -->
        <Border Grid.Row="0" Background="#0078D4">
            <Grid Margin="28,0">
                <StackPanel VerticalAlignment="Center">
                    <TextBlock Text="Windows 365 Deployment Wizard" Foreground="White" FontSize="17" FontWeight="SemiBold"/>
                    <TextBlock x:Name="TxtPageTitle" Foreground="#B3D9F5" FontSize="12" Margin="0,3,0,0"/>
                </StackPanel>
                <TextBlock x:Name="TxtStepCount" HorizontalAlignment="Right" VerticalAlignment="Center" Foreground="#B3D9F5" FontSize="12"/>
            </Grid>
        </Border>

        <!-- ═══ PAGES ═══ -->
        <Grid Grid.Row="1">

            <TabControl x:Name="WizardTabs" BorderThickness="0" Background="White">
                <TabControl.Resources>
                    <Style TargetType="TabItem"><Setter Property="Visibility" Value="Collapsed"/></Style>
                </TabControl.Resources>

                <!-- ─── PAGE 0 · Connect + Licence Type ─── -->
                <TabItem>
                    <ScrollViewer VerticalScrollBarVisibility="Auto">
                        <StackPanel Margin="32,28,32,16" MaxWidth="520" HorizontalAlignment="Left">

                            <TextBlock Text="Sign in to Microsoft Graph" FontSize="15" FontWeight="SemiBold" Margin="0,0,0,6"/>
                            <TextBlock TextWrapping="Wrap" Foreground="#666" Margin="0,0,0,18"
                                Text="Sign in with an account that has CloudPC.ReadWrite.All, Group.ReadWrite.All, LicenseAssignment.ReadWrite.All and DeviceManagementConfiguration.ReadWrite.All delegated permissions."/>

                            <Button x:Name="BtnConnect" Content="Sign in to Microsoft Graph" Style="{StaticResource PrimaryBtn}" HorizontalAlignment="Left"/>
                            <TextBlock x:Name="TxtConnectionStatus" Margin="0,10,0,0" FontSize="13" TextWrapping="Wrap"/>

                            <StackPanel x:Name="PanelLicenseType" Visibility="Collapsed" Margin="0,22,0,0">
                                <Separator Margin="0,0,0,20"/>
                                <TextBlock Text="Licence Type" FontWeight="SemiBold" Margin="0,0,0,10"/>
                                <RadioButton x:Name="RbEnterprise" Content="Enterprise"  GroupName="LicenseType"/>
                                <RadioButton x:Name="RbFrontline"  Content="Frontline"   GroupName="LicenseType"/>

                                <StackPanel x:Name="PanelFrontlineType" Visibility="Collapsed" Margin="22,10,0,0">
                                    <TextBlock Text="Provisioning Type" FontWeight="SemiBold" Margin="0,0,0,8"/>
                                    <RadioButton x:Name="RbFLDedicated" Content="Dedicated — shared by user"     GroupName="FLType"/>
                                    <RadioButton x:Name="RbFLShared"    Content="Shared — shared by Entra group" GroupName="FLType"/>
                                </StackPanel>
                            </StackPanel>

                        </StackPanel>
                    </ScrollViewer>
                </TabItem>

                <!-- ─── PAGE 1 · Cloud PC SKU ─── -->
                <TabItem>
                    <Grid Margin="28,20">
                        <Grid.RowDefinitions>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="*"/>
                        </Grid.RowDefinitions>
                        <TextBlock Grid.Row="0" Text="Select Cloud PC SKU" FontSize="15" FontWeight="SemiBold" Margin="0,0,0,12"/>
                        <ListBox x:Name="LbSKU" Grid.Row="1" BorderBrush="#D0D0D0"/>
                    </Grid>
                </TabItem>

                <!-- ─── PAGE 2 · Region ─── -->
                <!-- Both list boxes sit in Grid rows with Height="*" so they fill and scroll correctly -->
                <TabItem>
                    <Grid Margin="28,20">
                        <Grid.RowDefinitions>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="*"/>
                        </Grid.RowDefinitions>
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="*"/>
                            <ColumnDefinition Width="14"/>
                            <ColumnDefinition Width="*"/>
                        </Grid.ColumnDefinitions>

                        <TextBlock Grid.Row="0" Grid.ColumnSpan="3" Text="Select Region" FontSize="15" FontWeight="SemiBold" Margin="0,0,0,12"/>

                        <!-- Left panel — region group -->
                        <Grid Grid.Row="1" Grid.Column="0">
                            <Grid.RowDefinitions>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="*"/>
                            </Grid.RowDefinitions>
                            <TextBlock Grid.Row="0" Text="Region Group" FontWeight="SemiBold" Foreground="#444" Margin="0,0,0,6"/>
                            <ListBox x:Name="LbRegionGroup" Grid.Row="1" BorderBrush="#D0D0D0" HorizontalContentAlignment="Stretch"/>
                        </Grid>

                        <!-- Right panel — specific region -->
                        <Grid Grid.Row="1" Grid.Column="2">
                            <Grid.RowDefinitions>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="*"/>
                            </Grid.RowDefinitions>
                            <TextBlock Grid.Row="0" Text="Specific Region" FontWeight="SemiBold" Foreground="#444" Margin="0,0,0,6"/>
                            <ListBox x:Name="LbRegion" Grid.Row="1" BorderBrush="#D0D0D0" HorizontalContentAlignment="Stretch"/>
                        </Grid>

                    </Grid>
                </TabItem>

                <!-- ─── PAGE 3 · Image ─── -->
                <TabItem>
                    <Grid Margin="28,20">
                        <Grid.RowDefinitions>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="*"/>
                        </Grid.RowDefinitions>
                        <TextBlock Grid.Row="0" Text="Select Windows 11 Image" FontSize="15" FontWeight="SemiBold" Margin="0,0,0,12"/>
                        <ListBox x:Name="LbImage" Grid.Row="1" BorderBrush="#D0D0D0"/>
                    </Grid>
                </TabItem>

                <!-- ─── PAGE 4 · Language ─── -->
                <TabItem>
                    <Grid Margin="28,20">
                        <Grid.RowDefinitions>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="*"/>
                        </Grid.RowDefinitions>
                        <TextBlock Grid.Row="0" Text="Select Language" FontSize="15" FontWeight="SemiBold" Margin="0,0,0,12"/>
                        <TextBox   x:Name="TxtLanguageSearch" Grid.Row="1" Margin="0,0,0,8"/>
                        <ListBox   x:Name="LbLanguage" Grid.Row="2" BorderBrush="#D0D0D0"/>
                    </Grid>
                </TabItem>

                <!-- ─── PAGE 5 · Windows Update ─── -->
                <TabItem>
                    <ScrollViewer Margin="28,20" VerticalScrollBarVisibility="Auto">
                        <StackPanel MaxWidth="560" HorizontalAlignment="Left">

                            <TextBlock Text="Windows Update Settings" FontSize="15" FontWeight="SemiBold" Margin="0,0,0,6"/>
                            <TextBlock TextWrapping="Wrap" Foreground="#666" Margin="0,0,0,18"
                                Text="Optionally create a Windows Update for Business ring and assign it to your Cloud PC devices group. You can also enable Microsoft Autopatch on the provisioning policy."/>

                            <!-- Update ring -->
                            <TextBlock Text="Windows Update for Business Ring" FontWeight="SemiBold" Margin="0,0,0,8"/>
                            <CheckBox x:Name="ChkCreateUpdateRing" Content="Create a Windows Update for Business ring" IsChecked="True"/>

                            <StackPanel x:Name="PanelUpdateRing" Margin="22,10,0,0">
                                <TextBlock Text="Ring Profile" Foreground="#555" Margin="0,0,0,6"/>
                                <RadioButton x:Name="RbRingStandard"     Content="Standard — quality updates deferred 7 days, feature updates on release" GroupName="UpdateRing" IsChecked="True"/>
                                <RadioButton x:Name="RbRingRecommended"  Content="Recommended — quality deferred 7 days, feature updates deferred 30 days"  GroupName="UpdateRing"/>
                                <RadioButton x:Name="RbRingDeferred"     Content="Deferred — quality deferred 14 days, feature updates deferred 180 days"    GroupName="UpdateRing"/>

                                <TextBlock Text="Ring Name" Foreground="#555" Margin="0,14,0,4"/>
                                <TextBox x:Name="TxtUpdateRingName" Text="W365-CloudPC-UpdateRing"/>
                            </StackPanel>

                            <Separator Margin="0,20,0,20"/>

                            <!-- Autopatch -->
                            <TextBlock Text="Microsoft Autopatch" FontWeight="SemiBold" Margin="0,0,0,8"/>
                            <TextBlock TextWrapping="Wrap" Foreground="#666" Margin="0,0,0,10"
                                Text="Autopatch fully automates Windows and Microsoft 365 update management. Enabling this sets the provisioning policy to use Autopatch. Requires Autopatch to be set up in your tenant (Intune > Tenant administration > Windows Autopatch)."/>
                            <CheckBox x:Name="ChkAutopatch" Content="Enable Microsoft Autopatch on the provisioning policy" IsChecked="False"/>

                            <Border x:Name="PanelAutopatch" BorderBrush="#FFF4CE" BorderThickness="1" CornerRadius="4"
                                    Background="#FFFBF0" Padding="12,10" Margin="0,10,0,0" Visibility="Collapsed">
                                <StackPanel>
                                    <TextBlock Text="⚠️  Prerequisites" FontWeight="SemiBold" Margin="0,0,0,6"/>
                                    <TextBlock TextWrapping="Wrap" FontSize="12" Foreground="#555"
                                        Text="• Autopatch must be activated in your tenant before deploying.&#x0a;• The provisioning policy will reference your Default Autopatch group.&#x0a;• If Autopatch is not configured the policy creation will fall back gracefully."/>
                                </StackPanel>
                            </Border>

                        </StackPanel>
                    </ScrollViewer>
                </TabItem>

                <!-- ─── PAGE 6 · Summary ─── -->
                <TabItem>
                    <ScrollViewer Margin="28,20" VerticalScrollBarVisibility="Auto">
                        <StackPanel>

                            <TextBlock Text="Review &amp; Confirm" FontSize="15" FontWeight="SemiBold" Margin="0,0,0,16"/>

                            <!-- Selections -->
                            <TextBlock Text="Your Selections" FontWeight="SemiBold" Foreground="#444" Margin="0,0,0,8"/>
                            <Border BorderBrush="#E0E0E0" BorderThickness="1" CornerRadius="4" Padding="14,8">
                                <Grid>
                                    <Grid.ColumnDefinitions>
                                        <ColumnDefinition Width="150"/>
                                        <ColumnDefinition Width="*"/>
                                    </Grid.ColumnDefinitions>
                                    <Grid.RowDefinitions>
                                        <RowDefinition Height="Auto"/>
                                        <RowDefinition Height="Auto"/>
                                        <RowDefinition Height="Auto"/>
                                        <RowDefinition Height="Auto"/>
                                        <RowDefinition Height="Auto"/>
                                        <RowDefinition Height="Auto"/>
                                        <RowDefinition Height="Auto"/>
                                    </Grid.RowDefinitions>
                                    <TextBlock Grid.Row="0" Grid.Column="0" Text="Licence Type"    Foreground="#777" Margin="0,5"/>
                                    <TextBlock Grid.Row="0" Grid.Column="1" x:Name="SumLicenseType"  FontWeight="SemiBold" Margin="0,5" TextWrapping="Wrap"/>
                                    <TextBlock Grid.Row="1" Grid.Column="0" Text="Cloud PC SKU"    Foreground="#777" Margin="0,5"/>
                                    <TextBlock Grid.Row="1" Grid.Column="1" x:Name="SumSKU"          FontWeight="SemiBold" Margin="0,5" TextWrapping="Wrap"/>
                                    <TextBlock Grid.Row="2" Grid.Column="0" Text="Region"          Foreground="#777" Margin="0,5"/>
                                    <TextBlock Grid.Row="2" Grid.Column="1" x:Name="SumRegion"       FontWeight="SemiBold" Margin="0,5" TextWrapping="Wrap"/>
                                    <TextBlock Grid.Row="3" Grid.Column="0" Text="Image"           Foreground="#777" Margin="0,5"/>
                                    <TextBlock Grid.Row="3" Grid.Column="1" x:Name="SumImage"        FontWeight="SemiBold" Margin="0,5" TextWrapping="Wrap"/>
                                    <TextBlock Grid.Row="4" Grid.Column="0" Text="Language"        Foreground="#777" Margin="0,5"/>
                                    <TextBlock Grid.Row="4" Grid.Column="1" x:Name="SumLanguage"     FontWeight="SemiBold" Margin="0,5"/>
                                    <TextBlock Grid.Row="5" Grid.Column="0" Text="Licence Assign"  Foreground="#777" Margin="0,5"/>
                                    <TextBlock Grid.Row="5" Grid.Column="1" x:Name="SumLicAssign"   FontWeight="SemiBold" Margin="0,5" TextWrapping="Wrap"/>
                                    <TextBlock Grid.Row="6" Grid.Column="0" Text="Windows Update"  Foreground="#777" Margin="0,5"/>
                                    <TextBlock Grid.Row="6" Grid.Column="1" x:Name="SumUpdateRing"  FontWeight="SemiBold" Margin="0,5" TextWrapping="Wrap"/>
                                </Grid>
                            </Border>

                            <!-- Objects to create -->
                            <TextBlock Text="Objects to Create / Reuse" FontWeight="SemiBold" Foreground="#444" Margin="0,16,0,8"/>
                            <Border BorderBrush="#E0E0E0" BorderThickness="1" CornerRadius="4" Padding="14,8">
                                <Grid>
                                    <Grid.ColumnDefinitions>
                                        <ColumnDefinition Width="150"/>
                                        <ColumnDefinition Width="*"/>
                                    </Grid.ColumnDefinitions>
                                    <Grid.RowDefinitions>
                                        <RowDefinition Height="Auto"/>
                                        <RowDefinition Height="Auto"/>
                                        <RowDefinition Height="Auto"/>
                                        <RowDefinition Height="Auto"/>
                                        <RowDefinition Height="Auto"/>
                                    </Grid.RowDefinitions>
                                    <TextBlock Grid.Row="0" Grid.Column="0" Text="Licensing Group"  Foreground="#777" Margin="0,5"/>
                                    <TextBlock Grid.Row="0" Grid.Column="1" x:Name="SumLicGroup"     FontFamily="Consolas" FontSize="12" Margin="0,5" TextWrapping="Wrap"/>
                                    <TextBlock Grid.Row="1" Grid.Column="0" Text="User Group"        Foreground="#777" Margin="0,5"/>
                                    <TextBlock Grid.Row="1" Grid.Column="1" x:Name="SumUserGroup"    FontFamily="Consolas" FontSize="12" Margin="0,5" TextWrapping="Wrap"/>
                                    <TextBlock Grid.Row="2" Grid.Column="0" Text="Admin Group"       Foreground="#777" Margin="0,5"/>
                                    <TextBlock Grid.Row="2" Grid.Column="1" x:Name="SumAdminGroup"   FontFamily="Consolas" FontSize="12" Margin="0,5" TextWrapping="Wrap"/>
                                    <TextBlock Grid.Row="3" Grid.Column="0" Text="Devices Group"     Foreground="#777" Margin="0,5"/>
                                    <TextBlock Grid.Row="3" Grid.Column="1" x:Name="SumDevicesGroup" FontFamily="Consolas" FontSize="12" Margin="0,5" TextWrapping="Wrap"/>
                                    <TextBlock Grid.Row="4" Grid.Column="0" Text="Policy"            Foreground="#777" Margin="0,5"/>
                                    <TextBlock Grid.Row="4" Grid.Column="1" x:Name="SumPolicy"       FontFamily="Consolas" FontSize="12" Margin="0,5" TextWrapping="Wrap"/>
                                </Grid>
                            </Border>

                            <!-- Advanced options -->
                            <Expander Header="Advanced Options" Margin="0,14,0,4" FontSize="12">
                                <Border BorderBrush="#E8E8E8" BorderThickness="1" CornerRadius="4" Margin="0,6,0,0" Padding="14,10">
                                    <StackPanel>
                                        <TextBlock Text="Group Prefix" Foreground="#555" Margin="0,0,0,4"/>
                                        <TextBox x:Name="TxtGroupPrefix" Text="SG-W365" Margin="0,0,0,10"/>
                                        <TextBlock Text="Provisioning Policy Suffix" Foreground="#555" Margin="0,0,0,4"/>
                                        <TextBox x:Name="TxtPolicySuffix" Text="Provisioning Policy" Margin="0,0,0,10"/>
                                        <Button x:Name="BtnRecalc" Content="Recalculate Names" Style="{StaticResource SecondaryBtn}" HorizontalAlignment="Left" Padding="12,5"/>
                                    </StackPanel>
                                </Border>
                            </Expander>

                        </StackPanel>
                    </ScrollViewer>
                </TabItem>

                <!-- ─── PAGE 7 · Results ─── -->
                <TabItem>
                    <Grid Margin="28,20">
                        <Grid.RowDefinitions>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="*"/>
                            <RowDefinition Height="Auto"/>
                        </Grid.RowDefinitions>
                        <TextBlock x:Name="TxtResultHeading" Grid.Row="0" FontSize="15" FontWeight="SemiBold" Margin="0,0,0,14"/>
                        <ScrollViewer Grid.Row="1" VerticalScrollBarVisibility="Auto">
                            <StackPanel x:Name="PanelResults"/>
                        </ScrollViewer>
                        <Button x:Name="BtnCopySteps" Grid.Row="2" Content="Copy Manual Steps to Clipboard"
                                Style="{StaticResource SecondaryBtn}" HorizontalAlignment="Left" Margin="0,10,0,0"/>
                    </Grid>
                </TabItem>

            </TabControl>

            <!-- Loading overlay -->
            <Grid x:Name="LoadingOverlay" Visibility="Collapsed" Background="#CC1A1A1A" Panel.ZIndex="10">
                <StackPanel VerticalAlignment="Center" HorizontalAlignment="Center">
                    <ProgressBar IsIndeterminate="True" Width="280" Height="3" Foreground="White" Background="#444"/>
                    <TextBlock x:Name="TxtLoading" Foreground="White" HorizontalAlignment="Center" FontSize="13" Margin="0,14,0,0"/>
                </StackPanel>
            </Grid>

        </Grid>

        <!-- ═══ FOOTER ═══ -->
        <Border Grid.Row="2" BorderBrush="#E0E0E0" BorderThickness="0,1,0,0" Background="White">
            <Grid Margin="28,0">
                <Button x:Name="BtnBack" Content="&#8592; Back" Style="{StaticResource SecondaryBtn}"
                        HorizontalAlignment="Left" VerticalAlignment="Center" IsEnabled="False"/>
                <StackPanel Orientation="Horizontal" HorizontalAlignment="Right" VerticalAlignment="Center">
                    <Button x:Name="BtnNext"   Content="Next &#8594;"     Style="{StaticResource PrimaryBtn}"  Margin="0,0,8,0"/>
                    <Button x:Name="BtnDeploy" Content="Deploy  &#10003;" Style="{StaticResource DeployBtn}"   Visibility="Collapsed"/>
                    <Button x:Name="BtnClose"  Content="Close"            Style="{StaticResource SecondaryBtn}" Visibility="Collapsed"/>
                </StackPanel>
            </Grid>
        </Border>

    </Grid>
</Window>
'@

# ════════════════════════════════════════════════════════════════════════════
#  LOAD WINDOW
# ════════════════════════════════════════════════════════════════════════════

$reader = [System.Xml.XmlNodeReader]::new($xaml)
$window = [System.Windows.Markup.XamlReader]::Load($reader)

function ctrl { param($n) $window.FindName($n) }

# ════════════════════════════════════════════════════════════════════════════
#  STATE
# ════════════════════════════════════════════════════════════════════════════

$script:state = @{
    IsConnected               = $false
    LicenseType               = $null
    FrontlineType             = 'sharedByUser'
    SelectedServicePlan       = $null
    SelectedRegionGroup       = $null
    SelectedRegionName        = $null
    SelectedRegionDisplayName = $null
    SelectedImage             = $null
    SelectedLanguage          = $null
    CreateUpdateRing          = $true
    UpdateRingProfile         = 'standard'
    UpdateRingName            = 'W365-CloudPC-UpdateRing'
    EnableAutopatch           = $false
    CalculatedNames           = $null
    ServicePlans              = @()
    AllRegions                = @()
    Images                    = @()
    FilteredLanguages         = $script:SupportedLanguages
    ManualStepsText           = ""
}

$script:currentPage = 0
$script:totalPages  = 7   # pages 0-6, results = page 7
$script:pageTitles  = @(
    'Connect to Microsoft Graph'
    'Select Cloud PC SKU'
    'Select Region'
    'Select Windows 11 Image'
    'Select Language'
    'Windows Update Settings'
    'Review & Confirm'
    'Deployment Results'
)

# ════════════════════════════════════════════════════════════════════════════
#  UI HELPERS
# ════════════════════════════════════════════════════════════════════════════

function Set-Page {
    param([int]$index)
    (ctrl 'WizardTabs').SelectedIndex = $index
    $script:currentPage = $index
    (ctrl 'TxtPageTitle').Text = $script:pageTitles[$index]
    # Results page (7) has no step count or Back button
    (ctrl 'TxtStepCount').Text    = if ($index -lt $script:totalPages) { "Step $($index + 1) of $script:totalPages" } else { "" }
    (ctrl 'BtnBack').IsEnabled    = ($index -gt 0 -and $index -lt $script:totalPages)
    (ctrl 'BtnBack').Visibility   = if ($index -eq $script:totalPages) { 'Collapsed' } else { 'Visible' }
    (ctrl 'BtnNext').Visibility   = if ($index -ge ($script:totalPages - 1)) { 'Collapsed' } else { 'Visible' }
    (ctrl 'BtnDeploy').Visibility = if ($index -eq ($script:totalPages - 1)) { 'Visible' } else { 'Collapsed' }
    (ctrl 'BtnClose').Visibility  = if ($index -eq $script:totalPages) { 'Visible' } else { 'Collapsed' }
}

function Show-Loading {
    param([string]$Message = "Please wait...")
    (ctrl 'TxtLoading').Text = $Message
    (ctrl 'LoadingOverlay').Visibility = 'Visible'
    (ctrl 'BtnNext').IsEnabled   = $false
    (ctrl 'BtnBack').IsEnabled   = $false
    (ctrl 'BtnDeploy').IsEnabled = $false
    $window.Dispatcher.Invoke([System.Windows.Threading.DispatcherPriority]::Render, [Action]{})
}

function Hide-Loading {
    (ctrl 'LoadingOverlay').Visibility = 'Collapsed'
    (ctrl 'BtnNext').IsEnabled   = $true
    (ctrl 'BtnDeploy').IsEnabled = $true
    (ctrl 'BtnBack').IsEnabled   = ($script:currentPage -gt 0 -and $script:currentPage -lt $script:totalPages)
}

function Show-Alert {
    param([string]$Message, [string]$Title = "Notice", [string]$Icon = "Warning")
    [System.Windows.MessageBox]::Show($Message, $Title, 'OK', $Icon) | Out-Null
}

# ════════════════════════════════════════════════════════════════════════════
#  SUMMARY CALCULATION
# ════════════════════════════════════════════════════════════════════════════

function Update-Summary {
    $prefix = (ctrl 'TxtGroupPrefix').Text.Trim(); if (-not $prefix) { $prefix = "SG-W365" }
    $suffix = (ctrl 'TxtPolicySuffix').Text.Trim(); if (-not $suffix) { $suffix = "Provisioning Policy" }

    $regionLabel      = (Get-Culture).TextInfo.ToTitleCase(($script:state.SelectedRegionName -replace '[_-]',' ').ToLower().Trim())
    $policyRegionRaw  = $script:state.SelectedRegionDisplayName ?? $script:state.SelectedRegionName
    $policyRegionName = (Get-Culture).TextInfo.ToTitleCase($policyRegionRaw.ToLower().Trim())
    $licInfix         = if ($script:state.LicenseType -eq "Frontline") { "FL" } else { "ENT" }

    $script:state.CalculatedNames = @{
        GroupPrefix    = $prefix
        PolicySuffix   = $suffix
        LicensingGroup = "${prefix}CloudPC_$($script:state.SelectedServicePlan.DisplayName)"
        UserGroup      = "${prefix}-${licInfix}-${regionLabel}-User"
        AdminGroup     = "${prefix}-${licInfix}-${regionLabel}-Admin"
        DevicesGroup   = "${prefix}CloudPC-Devices"
        PolicyName     = "${policyRegionName}-W365-$($script:state.LicenseType)-${suffix}"
    }

    $licAssignText   = if ($script:state.LicenseType -eq 'Enterprise') { "Automatic (group-based licensing)" } else { "Manual — Frontline licences assigned by service plan" }
    $updateRingText  = if ($script:state.CreateUpdateRing) { "$($script:state.UpdateRingProfile) profile  [$($script:state.UpdateRingName)]" } else { "Skipped" }
    if ($script:state.EnableAutopatch) { $updateRingText += "  +  Autopatch enabled" }

    (ctrl 'SumLicenseType').Text  = $script:state.LicenseType
    (ctrl 'SumSKU').Text          = $script:state.SelectedServicePlan.DisplayName
    (ctrl 'SumRegion').Text       = "$($script:state.SelectedRegionDisplayName) ($($script:state.SelectedRegionGroup))"
    (ctrl 'SumImage').Text        = $script:state.SelectedImage.displayName
    (ctrl 'SumLanguage').Text     = $script:state.SelectedLanguage.DisplayName
    (ctrl 'SumLicAssign').Text    = $licAssignText
    (ctrl 'SumUpdateRing').Text   = $updateRingText
    (ctrl 'SumLicGroup').Text     = $script:state.CalculatedNames.LicensingGroup
    (ctrl 'SumUserGroup').Text    = $script:state.CalculatedNames.UserGroup
    (ctrl 'SumAdminGroup').Text   = $script:state.CalculatedNames.AdminGroup
    (ctrl 'SumDevicesGroup').Text = $script:state.CalculatedNames.DevicesGroup
    (ctrl 'SumPolicy').Text       = $script:state.CalculatedNames.PolicyName
}

# ════════════════════════════════════════════════════════════════════════════
#  RESULTS PAGE HELPERS
# ════════════════════════════════════════════════════════════════════════════

function Add-ResultSection {
    param([string]$Title)
    $tb = New-Object System.Windows.Controls.TextBlock
    $tb.Text       = $Title
    $tb.FontWeight = 'SemiBold'
    $tb.Foreground = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromRgb(0,0x78,0xD4))
    $tb.Margin     = [System.Windows.Thickness]::new(0,10,0,4)
    (ctrl 'PanelResults').Children.Add($tb) | Out-Null
}

function Add-ResultRow {
    param([string]$Label, [string]$Value, [bool]$Created = $true, [string]$Note = "")
    $row = New-Object System.Windows.Controls.Grid
    $c0  = New-Object System.Windows.Controls.ColumnDefinition; $c0.Width = [System.Windows.GridLength]::new(140)
    $c1  = New-Object System.Windows.Controls.ColumnDefinition; $c1.Width = [System.Windows.GridLength]::new(1,[System.Windows.GridUnitType]::Star)
    $row.ColumnDefinitions.Add($c0); $row.ColumnDefinitions.Add($c1)
    $row.Margin = [System.Windows.Thickness]::new(0,2,0,2)

    $lbl = New-Object System.Windows.Controls.TextBlock
    $lbl.Text = $Label; $lbl.Foreground = [System.Windows.Media.Brushes]::Gray; $lbl.VerticalAlignment = 'Top'
    [System.Windows.Controls.Grid]::SetColumn($lbl, 0)

    $icon = if ($Created) { "✅ Created" } else { "⚡ Already existed" }
    $val  = New-Object System.Windows.Controls.TextBlock
    $val.Text         = "$icon  —  $Value$(if ($Note) { "  $Note" })"
    $val.TextWrapping = 'Wrap'
    $val.FontFamily   = [System.Windows.Media.FontFamily]::new("Consolas")
    $val.FontSize     = 11
    $val.VerticalAlignment = 'Top'
    [System.Windows.Controls.Grid]::SetColumn($val, 1)

    $row.Children.Add($lbl) | Out-Null
    $row.Children.Add($val) | Out-Null
    (ctrl 'PanelResults').Children.Add($row) | Out-Null
}

function Add-ResultNote {
    param([string]$Text, [string]$Color = "#666666")
    $tb = New-Object System.Windows.Controls.TextBlock
    $r, $g, $b = [Convert]::ToByte($Color.Substring(1,2),16), [Convert]::ToByte($Color.Substring(3,2),16), [Convert]::ToByte($Color.Substring(5,2),16)
    $tb.Text         = $Text
    $tb.TextWrapping = 'Wrap'
    $tb.Foreground   = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromRgb($r,$g,$b))
    $tb.FontSize     = 12
    $tb.Margin       = [System.Windows.Thickness]::new(0,4,0,0)
    (ctrl 'PanelResults').Children.Add($tb) | Out-Null
}

# ════════════════════════════════════════════════════════════════════════════
#  DEPLOYMENT
# ════════════════════════════════════════════════════════════════════════════

function Start-Deployment {
    (ctrl 'BtnDeploy').IsEnabled = $false
    (ctrl 'BtnBack').IsEnabled   = $false
    (ctrl 'PanelResults').Children.Clear()

    $names = $script:state.CalculatedNames

    try {
        # ── Groups ────────────────────────────────────────────────────────
        Show-Loading "Creating Entra ID groups..."
        $rLic     = Get-OrCreateGroup -DisplayName $names.LicensingGroup -Description "Windows 365 licensing group — service plan: $($script:state.SelectedServicePlan.Id)"
        $rUser    = Get-OrCreateGroup -DisplayName $names.UserGroup       -Description "Windows 365 users in $($script:state.SelectedRegionDisplayName)"
        $rAdmin   = Get-OrCreateGroup -DisplayName $names.AdminGroup      -Description "Windows 365 local admins in $($script:state.SelectedRegionDisplayName)"
        $rDevices = Get-OrCreateDynamicDeviceGroup -DisplayName $names.DevicesGroup -Description "Dynamic group — all Windows 365 Cloud PC devices (CPC-* prefix)"

        # Group replication wait
        for ($i = 10; $i -gt 0; $i--) {
            (ctrl 'TxtLoading').Text = "Waiting for group replication ($i s)..."
            $window.Dispatcher.Invoke([System.Windows.Threading.DispatcherPriority]::Render, [Action]{})
            Start-Sleep -Seconds 1
        }

        # ── Licence assignment ────────────────────────────────────────────
        $licResult = $null
        if ($script:state.LicenseType -eq 'Enterprise') {
            Show-Loading "Assigning Windows 365 licence to licensing group..."
            try {
                $licResult = Set-GroupLicense -GroupId $rLic.Id -ServicePlanId $script:state.SelectedServicePlan.Id
            } catch {
                # Non-fatal — tenant may not have Azure AD P1 or licence not purchased yet
                $licResult = @{ SkuPartNumber = "⚠️ Could not auto-assign: $_"; Warning = "" }
            }
        }

        # ── Cloud PC user settings ────────────────────────────────────────
        Show-Loading "Creating Cloud PC user settings..."
        $rAdminSettings = Get-OrCreateCloudPcUserSetting -DisplayName "W365_AdminSettings" -LocalAdminEnabled $true  -TargetGroupId $rAdmin.Id
        $rUserSettings  = Get-OrCreateCloudPcUserSetting -DisplayName "W365_UserSettings"  -LocalAdminEnabled $false -TargetGroupId $rUser.Id

        $isCopilot   = Test-IsCopilotEligibleSku -DisplayName $script:state.SelectedServicePlan.DisplayName
        $rAiSettings = $null
        if ($isCopilot) {
            $rAiSettings = Get-OrCreateCloudPcUserSetting -DisplayName "AI_Enabled_Cloud_PC" -LocalAdminEnabled $false -TargetGroupId $rLic.Id
        }

        # ── Provisioning policy ───────────────────────────────────────────
        Show-Loading "Creating provisioning policy..."
        $provType        = $script:state.FrontlineType ?? 'dedicated'
        $userPersistence = ($script:state.LicenseType -eq 'Frontline' -and $script:state.FrontlineType -eq 'sharedByEntraGroup')
        $servicePlanId   = if ($script:state.LicenseType -eq 'Frontline') { $script:state.SelectedServicePlan.Id } else { $null }

        $rPolicy = Get-OrCreateProvisioningPolicy `
            -DisplayName             $names.PolicyName `
            -AssignGroupIds          @($rAdmin.Id, $rUser.Id) `
            -RegionGroup             $script:state.SelectedRegionGroup `
            -CountryRegion           $script:state.SelectedRegionName `
            -ImageId                 $script:state.SelectedImage.id `
            -ImageDisplayName        $script:state.SelectedImage.displayName `
            -Language                $script:state.SelectedLanguage.Code `
            -ProvisioningType        $provType `
            -UserSettingsPersistence $userPersistence `
            -ServicePlanId           $servicePlanId

        # ── Windows Update for Business ring ─────────────────────────────
        $rUpdateRing = $null
        if ($script:state.CreateUpdateRing) {
            Show-Loading "Creating Windows Update for Business ring..."
            try {
                $rUpdateRing = Get-OrCreateUpdateRing `
                    -DisplayName   $script:state.UpdateRingName `
                    -DeviceGroupId $rDevices.Id `
                    -Profile       $script:state.UpdateRingProfile
            } catch {
                $rUpdateRing = @{ Id = $null; Created = $false; Error = "$_" }
            }
        }

        # ── Autopatch ─────────────────────────────────────────────────────
        $autopatchNote = ""
        if ($script:state.EnableAutopatch) {
            Show-Loading "Enabling Autopatch on provisioning policy..."
            try {
                # Fetch Autopatch groups to get the default group ID
                $apGroups = Invoke-MgGraphRequest -Method GET `
                    -Uri "https://graph.microsoft.com/beta/deviceManagement/autopatchGroups" `
                    -ErrorAction Stop
                $defaultGroup = $apGroups.value | Where-Object { $_.isDefaultGroup -eq $true } | Select-Object -First 1
                if ($defaultGroup) {
                    $apPatch = @{ autopatch = @{ autopatchGroupId = $defaultGroup.id } }
                    Invoke-MgGraphRequest -Method PATCH `
                        -Uri "https://graph.microsoft.com/beta/deviceManagement/virtualEndpoint/provisioningPolicies/$($rPolicy.Id)" `
                        -Body ($apPatch | ConvertTo-Json -Depth 4) -ContentType "application/json" -ErrorAction Stop | Out-Null
                    $autopatchNote = "Autopatch group: $($defaultGroup.displayName)"
                } else {
                    $autopatchNote = "⚠️ No default Autopatch group found — configure Autopatch in Intune first"
                }
            } catch {
                $autopatchNote = "⚠️ Autopatch could not be applied: $_"
            }
        }

        Hide-Loading

        # ── Build results page ────────────────────────────────────────────
        (ctrl 'TxtResultHeading').Text       = "Deployment Complete  ✅"
        (ctrl 'TxtResultHeading').Foreground = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromRgb(16,124,16))

        Add-ResultSection "Entra ID Groups"
        Add-ResultRow "Licensing" $names.LicensingGroup -Created $rLic.Created
        Add-ResultRow "Users"     $names.UserGroup      -Created $rUser.Created
        Add-ResultRow "Admins"    $names.AdminGroup     -Created $rAdmin.Created
        Add-ResultRow "Devices"   $names.DevicesGroup   -Created $rDevices.Created

        Add-ResultSection "Licence Assignment"
        if ($licResult) {
            Add-ResultRow "SKU" $licResult.SkuPartNumber -Created $true -Note $licResult.Warning
        } else {
            Add-ResultNote "Skipped — Frontline licences are assigned per service plan automatically."
        }

        Add-ResultSection "Cloud PC User Settings"
        Add-ResultRow "Admin Settings" "W365_AdminSettings" -Created $rAdminSettings.Created
        Add-ResultRow "User Settings"  "W365_UserSettings"  -Created $rUserSettings.Created
        if ($isCopilot -and $rAiSettings) {
            Add-ResultRow "AI Settings" "AI_Enabled_Cloud_PC" -Created $rAiSettings.Created
        }

        Add-ResultSection "Provisioning Policy"
        Add-ResultRow "Policy" $names.PolicyName -Created $rPolicy.Created
        if ($autopatchNote) { Add-ResultNote $autopatchNote }

        Add-ResultSection "Windows Update"
        if ($rUpdateRing) {
            if ($rUpdateRing.Error) {
                Add-ResultNote "⚠️ Update ring could not be created: $($rUpdateRing.Error)" "#C50F1F"
            } else {
                Add-ResultRow "Update Ring" $script:state.UpdateRingName -Created $rUpdateRing.Created
            }
        } else {
            Add-ResultNote "Skipped."
        }

        # ── Manual steps text for clipboard ──────────────────────────────
        $step = 1
        $script:state.ManualStepsText = "MANUAL STEPS REQUIRED`n$('=' * 55)`n`n"
        if ($script:state.LicenseType -eq 'Enterprise') {
            if ($licResult -and $licResult.SkuPartNumber -notlike '⚠️*') {
                $script:state.ManualStepsText += "$step. Licence Assignment (DONE — auto-assigned via group-based licensing)`n   Group: $($names.LicensingGroup)`n`n"
            } else {
                $script:state.ManualStepsText += "$step. Licence Assignment (MANUAL)`n   Assign '$($script:state.SelectedServicePlan.DisplayName)' to:`n   $($names.LicensingGroup)`n`n"
            }
        } else {
            $script:state.ManualStepsText += "$step. Frontline Licence`n   Assign policy '$($names.PolicyName)' once Frontline licences are purchased.`n`n"
        }
        $step++
        $script:state.ManualStepsText += "$step. User & Admin Groups`n   User  : $($names.UserGroup)`n   Admin : $($names.AdminGroup)`n`n"; $step++
        $script:state.ManualStepsText += "$step. Devices Group (auto-populates from CPC-* naming)`n   $($names.DevicesGroup)`n`n"; $step++
        if ($script:state.EnableAutopatch) {
            $script:state.ManualStepsText += "$step. Autopatch`n   Verify Autopatch activation in Intune > Tenant administration > Windows Autopatch.`n`n"; $step++
        }
        $script:state.ManualStepsText += "$step. Cross-Region DR (optional)`n   Intune admin centre > Devices > Cloud PCs > User settings.`n"

        Add-ResultSection "Next Steps"
        Add-ResultNote "Click 'Copy Manual Steps to Clipboard' for the full post-deployment checklist."

        Disconnect-MgGraph | Out-Null
        Set-Page 7
    }
    catch {
        Hide-Loading
        (ctrl 'TxtResultHeading').Text       = "Deployment Failed  ✗"
        (ctrl 'TxtResultHeading').Foreground = [System.Windows.Media.Brushes]::Crimson
        $errTb = New-Object System.Windows.Controls.TextBlock
        $errTb.Text        = "Error: $_`n`nCheck that your account has the required permissions and that the Windows 365 licence is purchased in this tenant."
        $errTb.TextWrapping = 'Wrap'; $errTb.Foreground = [System.Windows.Media.Brushes]::Crimson
        (ctrl 'PanelResults').Children.Add($errTb) | Out-Null
        (ctrl 'BtnDeploy').IsEnabled = $true
        (ctrl 'BtnBack').IsEnabled   = $true
        Set-Page 7
    }
}

# ════════════════════════════════════════════════════════════════════════════
#  NAVIGATION
# ════════════════════════════════════════════════════════════════════════════

function Move-Next {
    switch ($script:currentPage) {

        0 { # Connect / Licence Type
            if (-not $script:state.IsConnected)  { Show-Alert "Please sign in to Microsoft Graph first." "Not Signed In"; return }
            if (-not $script:state.LicenseType)  { Show-Alert "Please select a licence type." "Selection Required"; return }
            if ($script:state.LicenseType -eq "Frontline" -and -not $script:state.FrontlineType) {
                Show-Alert "Please select a Frontline provisioning type." "Selection Required"; return
            }
            Show-Loading "Retrieving Cloud PC service plans..."
            try {
                $all   = Get-AllGraphItems -Uri "https://graph.microsoft.com/beta/deviceManagement/virtualEndpoint/servicePlans"
                $plans = @($all | ForEach-Object {
                    $dn = ($_.displayName ?? $_.DisplayName) ?? "Unknown"
                    $id = ($_.id ?? $_.Id)
                    [pscustomobject]@{ DisplayName = $dn; Id = $id }
                })
                $plans = if ($script:state.LicenseType -eq "Enterprise") {
                    $plans | Where-Object { $_.DisplayName -notmatch 'Business|Frontline' }
                } else {
                    $plans | Where-Object { $_.DisplayName -match 'Frontline' }
                }
                $script:state.ServicePlans = @($plans | Sort-Object DisplayName)
                $lb = ctrl 'LbSKU'; $lb.Items.Clear()
                $script:state.ServicePlans | ForEach-Object { $lb.Items.Add($_.DisplayName) }
            } catch { Hide-Loading; Show-Alert "Failed to load service plans:`n$_" "Error" "Error"; return }
            Hide-Loading; Set-Page 1
        }

        1 { # SKU
            $lb = ctrl 'LbSKU'
            if ($lb.SelectedIndex -lt 0) { Show-Alert "Please select a Cloud PC SKU." "Selection Required"; return }
            $script:state.SelectedServicePlan = $script:state.ServicePlans[$lb.SelectedIndex]
            Show-Loading "Retrieving supported regions..."
            try {
                $uri = "https://graph.microsoft.com/beta/deviceManagement/virtualEndpoint/supportedRegions?`$filter=supportedSolution eq 'windows365'&`$select=id,displayName,regionGroup,geographicLocationType"
                $raw = Get-AllGraphItems -Uri $uri
                $raw = $raw | Where-Object { $null -eq ($_.geographicLocationType ?? $_.GeographicLocationType) }
                $script:state.AllRegions = @($raw | ForEach-Object {
                    $rg = $_.regionGroup ?? $_.RegionGroup
                    $dn = $_.displayName ?? $_.DisplayName
                    [pscustomobject]@{ RegionGroup = $rg; DisplayName = $dn; RegionName = $dn }
                } | Sort-Object RegionGroup, DisplayName)

                $uniqueGroups = @($script:state.AllRegions.RegionGroup | Sort-Object -Unique)
                $lbGroup = ctrl 'LbRegionGroup'; $lbGroup.Items.Clear()
                foreach ($g in $uniqueGroups) {
                    $words    = [regex]::Split($g, '(?<=[a-z])(?=[A-Z])')
                    $friendly = ($words | ForEach-Object { $_.Substring(0,1).ToUpper() + $_.Substring(1).ToLower() }) -join ' '
                    $lbGroup.Items.Add([pscustomobject]@{ Display = $friendly; Value = $g })
                }
                $lbGroup.DisplayMemberPath = 'Display'
            } catch { Hide-Loading; Show-Alert "Failed to load regions:`n$_" "Error" "Error"; return }
            Hide-Loading; Set-Page 2
        }

        2 { # Region
            if ((ctrl 'LbRegionGroup').SelectedIndex -lt 0 -or (ctrl 'LbRegion').SelectedIndex -lt 0) {
                Show-Alert "Please select both a region group and a specific region." "Selection Required"; return
            }
            $selGroup  = (ctrl 'LbRegionGroup').SelectedItem
            $selRegion = (ctrl 'LbRegion').SelectedItem
            $script:state.SelectedRegionGroup        = $selGroup.Value
            $script:state.SelectedRegionDisplayName  = $selRegion.Display
            $script:state.SelectedRegionName         = $selRegion.Value
            Show-Loading "Retrieving Windows 11 gallery images..."
            try {
                $imgs = Get-AllGraphItems -Uri "https://graph.microsoft.com/beta/deviceManagement/virtualEndpoint/galleryImages"
                $script:state.Images = @($imgs | Where-Object { -not $_.status -or $_.status -ne "notSupported" } | Sort-Object displayName)
                $lb = ctrl 'LbImage'; $lb.Items.Clear()
                $script:state.Images | ForEach-Object { $lb.Items.Add($_.displayName) }
            } catch { Hide-Loading; Show-Alert "Failed to load images:`n$_" "Error" "Error"; return }
            Hide-Loading; Set-Page 3
        }

        3 { # Image
            $lb = ctrl 'LbImage'
            if ($lb.SelectedIndex -lt 0) { Show-Alert "Please select a Windows 11 image." "Selection Required"; return }
            $script:state.SelectedImage = $script:state.Images[$lb.SelectedIndex]
            Set-Page 4
        }

        4 { # Language
            $lb = ctrl 'LbLanguage'
            if ($lb.SelectedIndex -lt 0) { Show-Alert "Please select a language." "Selection Required"; return }
            $script:state.SelectedLanguage = $script:state.FilteredLanguages[$lb.SelectedIndex]
            Set-Page 5
        }

        5 { # Windows Update — read selections then go to summary
            $script:state.CreateUpdateRing = (ctrl 'ChkCreateUpdateRing').IsChecked
            $script:state.UpdateRingName   = (ctrl 'TxtUpdateRingName').Text.Trim()
            if (-not $script:state.UpdateRingName) { $script:state.UpdateRingName = "W365-CloudPC-UpdateRing" }
            $script:state.UpdateRingProfile = switch ($true) {
                (ctrl 'RbRingRecommended').IsChecked { 'recommended' }
                (ctrl 'RbRingDeferred').IsChecked    { 'deferred' }
                default                              { 'standard' }
            }
            $script:state.EnableAutopatch = (ctrl 'ChkAutopatch').IsChecked
            Update-Summary
            Set-Page 6
        }
    }
}

# ════════════════════════════════════════════════════════════════════════════
#  EVENT HANDLERS
# ════════════════════════════════════════════════════════════════════════════

(ctrl 'BtnConnect').Add_Click({
    Show-Loading "Connecting to Microsoft Graph..."
    try {
        Install-GraphModuleIfNeeded
        Connect-MgGraph -Scopes "User.ReadWrite.All","Application.ReadWrite.All","CloudPC.ReadWrite.All","Group.ReadWrite.All","LicenseAssignment.ReadWrite.All","DeviceManagementConfiguration.ReadWrite.All" -ErrorAction Stop
        $script:state.IsConnected = $true
        (ctrl 'TxtConnectionStatus').Text       = [char]0x2714 + "  Connected to Microsoft Graph"
        (ctrl 'TxtConnectionStatus').Foreground = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromRgb(16,124,16))
        (ctrl 'PanelLicenseType').Visibility    = 'Visible'
        (ctrl 'BtnConnect').IsEnabled           = $false
    } catch {
        (ctrl 'TxtConnectionStatus').Text       = [char]0x2718 + "  Connection failed — $_"
        (ctrl 'TxtConnectionStatus').Foreground = [System.Windows.Media.Brushes]::Crimson
    }
    Hide-Loading
})

(ctrl 'RbEnterprise').Add_Checked({ $script:state.LicenseType = "Enterprise"; (ctrl 'PanelFrontlineType').Visibility = 'Collapsed' })
(ctrl 'RbFrontline').Add_Checked({  $script:state.LicenseType = "Frontline";  (ctrl 'PanelFrontlineType').Visibility = 'Visible' })
(ctrl 'RbFLDedicated').Add_Checked({ $script:state.FrontlineType = "sharedByUser" })
(ctrl 'RbFLShared').Add_Checked({    $script:state.FrontlineType = "sharedByEntraGroup" })

(ctrl 'LbRegionGroup').Add_SelectionChanged({
    $selGroup = (ctrl 'LbRegionGroup').SelectedItem
    if (-not $selGroup) { return }
    $regions  = @($script:state.AllRegions | Where-Object { $_.RegionGroup -eq $selGroup.Value } | Sort-Object DisplayName)
    $lbRegion = ctrl 'LbRegion'; $lbRegion.Items.Clear()
    $regions | ForEach-Object { $lbRegion.Items.Add([pscustomobject]@{ Display = $_.DisplayName; Value = $_.RegionName }) }
    $lbRegion.DisplayMemberPath = 'Display'
})

(ctrl 'ChkCreateUpdateRing').Add_Checked({   (ctrl 'PanelUpdateRing').Visibility = 'Visible' })
(ctrl 'ChkCreateUpdateRing').Add_Unchecked({ (ctrl 'PanelUpdateRing').Visibility = 'Collapsed' })
(ctrl 'ChkAutopatch').Add_Checked({          (ctrl 'PanelAutopatch').Visibility  = 'Visible' })
(ctrl 'ChkAutopatch').Add_Unchecked({        (ctrl 'PanelAutopatch').Visibility  = 'Collapsed' })

(ctrl 'TxtLanguageSearch').Add_GotFocus({
    if ((ctrl 'TxtLanguageSearch').Tag -eq 'placeholder') {
        (ctrl 'TxtLanguageSearch').Text       = ""
        (ctrl 'TxtLanguageSearch').Foreground = [System.Windows.Media.Brushes]::Black
        (ctrl 'TxtLanguageSearch').Tag        = ""
    }
})
(ctrl 'TxtLanguageSearch').Add_LostFocus({
    if ((ctrl 'TxtLanguageSearch').Text -eq "") {
        (ctrl 'TxtLanguageSearch').Text       = "Search languages..."
        (ctrl 'TxtLanguageSearch').Foreground = [System.Windows.Media.Brushes]::Gray
        (ctrl 'TxtLanguageSearch').Tag        = "placeholder"
    }
})
(ctrl 'TxtLanguageSearch').Add_TextChanged({
    if ((ctrl 'TxtLanguageSearch').Tag -eq 'placeholder') { return }
    $search   = (ctrl 'TxtLanguageSearch').Text.ToLower()
    $filtered = @($script:SupportedLanguages | Where-Object { $_.DisplayName.ToLower().Contains($search) })
    $script:state.FilteredLanguages = $filtered
    $lb = ctrl 'LbLanguage'; $lb.Items.Clear()
    $filtered | ForEach-Object { $lb.Items.Add($_.DisplayName) }
})

(ctrl 'BtnRecalc').Add_Click({ Update-Summary })
(ctrl 'BtnNext').Add_Click({   Move-Next })
(ctrl 'BtnBack').Add_Click({   if ($script:currentPage -gt 0 -and $script:currentPage -lt $script:totalPages) { Set-Page ($script:currentPage - 1) } })
(ctrl 'BtnDeploy').Add_Click({ Start-Deployment })
(ctrl 'BtnClose').Add_Click({  $window.Close() })
(ctrl 'BtnCopySteps').Add_Click({
    if ($script:state.ManualStepsText) {
        [System.Windows.Clipboard]::SetText($script:state.ManualStepsText)
        Show-Alert "Manual steps copied to clipboard." "Copied" "Information"
    }
})

# Populate language list on startup with placeholder
$script:SupportedLanguages | ForEach-Object { (ctrl 'LbLanguage').Items.Add($_.DisplayName) }
(ctrl 'TxtLanguageSearch').Text       = "Search languages..."
(ctrl 'TxtLanguageSearch').Foreground = [System.Windows.Media.Brushes]::Gray
(ctrl 'TxtLanguageSearch').Tag        = "placeholder"

# ════════════════════════════════════════════════════════════════════════════
#  LAUNCH
# ════════════════════════════════════════════════════════════════════════════
Set-Page 0
$window.ShowDialog() | Out-Null
