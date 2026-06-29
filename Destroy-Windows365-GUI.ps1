#Requires -Version 7.0
<#
.SYNOPSIS
    Windows 365 Cleanup Tool — WPF graphical front-end.
.DESCRIPTION
    Connects to Microsoft Graph, scans for Windows 365 objects, lets you
    individually check/uncheck each item, then deletes only what you selected.
.NOTES
    Requires Microsoft.Graph.Authentication (auto-installed if missing).
    Must run on Windows (WPF). STA thread is handled automatically.
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

# ── Graph helpers ─────────────────────────────────────────────────────────────

function Install-GraphModuleIfNeeded {
    if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication)) {
        Install-Module -Name Microsoft.Graph.Authentication -AllowClobber -Force -ErrorAction Stop
    }
}

function Get-AllGraphItems {
    param([Parameter(Mandatory)][string]$Uri)
    $items = @()
    $nextLink = $Uri
    while ($nextLink) {
        $response = Invoke-MgGraphRequest -Method GET -Uri $nextLink -ErrorAction Stop
        if ($response.value) { $items += $response.value }
        $nextLink = $response.'@odata.nextLink'
    }
    return $items
}

function Clear-PolicyAssignments {
    param([Parameter(Mandatory)][string]$Uri)
    $payload = @{ assignments = @() } | ConvertTo-Json -Depth 3
    Invoke-MgGraphRequest -Method POST -Uri $Uri -Body $payload -ContentType "application/json" -ErrorAction Stop | Out-Null
}

# ── XAML ─────────────────────────────────────────────────────────────────────

[xml]$xaml = @'
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="CloudEndpoint.AI — Windows 365 Cleanup Tool"
    Width="700" Height="660"
    MinWidth="700" MinHeight="600"
    WindowStartupLocation="CenterScreen"
    ResizeMode="CanResize"
    FontFamily="Segoe UI"
    FontSize="13"
    Background="White">

    <Window.Resources>

        <Style x:Key="PrimaryBtn" TargetType="Button">
            <Setter Property="Background"      Value="#2BC0B8"/>
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
                            <Trigger Property="IsMouseOver" Value="True"><Setter Property="Background" Value="#1A9E98"/></Trigger>
                            <Trigger Property="IsEnabled"   Value="False"><Setter Property="Background" Value="#C8C8C8"/><Setter Property="Foreground" Value="#999"/></Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <Style x:Key="SecondaryBtn" TargetType="Button">
            <Setter Property="Background"      Value="White"/>
            <Setter Property="Foreground"      Value="#2BC0B8"/>
            <Setter Property="BorderBrush"     Value="#2BC0B8"/>
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
                            <Trigger Property="IsMouseOver" Value="True"><Setter Property="Background" Value="#E6F9F8"/></Trigger>
                            <Trigger Property="IsEnabled"   Value="False"><Setter Property="BorderBrush" Value="#CCC"/><Setter Property="Foreground" Value="#CCC"/></Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <Style x:Key="SmallBtn" TargetType="Button">
            <Setter Property="Background"      Value="White"/>
            <Setter Property="Foreground"      Value="#555"/>
            <Setter Property="BorderBrush"     Value="#CCC"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Padding"         Value="10,4"/>
            <Setter Property="FontSize"        Value="11"/>
            <Setter Property="Cursor"          Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}" CornerRadius="3" Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True"><Setter Property="Background" Value="#F5F5F5"/></Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <Style x:Key="DangerBtn" TargetType="Button">
            <Setter Property="Background"      Value="#C50F1F"/>
            <Setter Property="Foreground"      Value="White"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Padding"         Value="18,7"/>
            <Setter Property="FontWeight"      Value="SemiBold"/>
            <Setter Property="Cursor"          Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border Background="{TemplateBinding Background}" CornerRadius="3" Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True"><Setter Property="Background" Value="#A00E1A"/></Trigger>
                            <Trigger Property="IsEnabled"   Value="False"><Setter Property="Background" Value="#C8C8C8"/><Setter Property="Foreground" Value="#999"/></Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

    </Window.Resources>

    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="72"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="56"/>
        </Grid.RowDefinitions>

        <!-- ═══ HEADER ═══ -->
        <Border Grid.Row="0" Background="#111827">
            <Grid Margin="28,0">
                <StackPanel VerticalAlignment="Center">
                    <TextBlock Text="CloudEndpoint.AI  ·  Windows 365 Cleanup Tool" Foreground="White" FontSize="17" FontWeight="SemiBold"/>
                    <TextBlock Text="Scan your tenant, then choose exactly which objects to remove."
                               Foreground="#7DD8D4" FontSize="12" Margin="0,3,0,0" TextWrapping="Wrap"/>
                </StackPanel>
            </Grid>
        </Border>

        <!-- ═══ CONTENT ═══ -->
        <Grid Grid.Row="1">

            <TabControl x:Name="MainTabs" BorderThickness="0" Background="White">
                <TabControl.Resources>
                    <Style TargetType="TabItem"><Setter Property="Visibility" Value="Collapsed"/></Style>
                </TabControl.Resources>

                <!-- PAGE 0 · Connect -->
                <TabItem>
                    <StackPanel Margin="32,28,32,16" MaxWidth="480" HorizontalAlignment="Left">
                        <TextBlock Text="Sign in to Microsoft Graph" FontSize="15" FontWeight="SemiBold" Margin="0,0,0,6"/>
                        <TextBlock TextWrapping="Wrap" Foreground="#666" Margin="0,0,0,18"
                            Text="Sign in with an account that has CloudPC.ReadWrite.All, Group.ReadWrite.All and DeviceManagementConfiguration.ReadWrite.All permissions."/>
                        <Button x:Name="BtnConnect" Content="Sign in to Microsoft Graph" Style="{StaticResource PrimaryBtn}" HorizontalAlignment="Left"/>
                        <TextBlock x:Name="TxtConnectionStatus" Margin="0,10,0,0" FontSize="13" TextWrapping="Wrap"/>
                    </StackPanel>
                </TabItem>

                <!-- PAGE 1 · Scan + Select -->
                <TabItem>
                    <Grid Margin="28,18,28,8">
                        <Grid.RowDefinitions>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="*"/>
                        </Grid.RowDefinitions>

                        <!-- Scan filters -->
                        <Grid Grid.Row="0" Margin="0,0,0,10">
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="Auto"/>
                            </Grid.ColumnDefinitions>
                            <WrapPanel Grid.Column="0" VerticalAlignment="Center">
                                <TextBlock Text="Scan for:" VerticalAlignment="Center" Margin="0,0,12,4" Foreground="#555"/>
                                <CheckBox x:Name="ChkPolicies"     Content="Provisioning Policies" IsChecked="True" Margin="0,0,14,4" VerticalAlignment="Center"/>
                                <CheckBox x:Name="ChkUserSettings" Content="User Settings"          IsChecked="True" Margin="0,0,14,4" VerticalAlignment="Center"/>
                                <CheckBox x:Name="ChkGroups"       Content="Entra ID Groups"        IsChecked="True" Margin="0,0,14,4" VerticalAlignment="Center"/>
                                <CheckBox x:Name="ChkUpdateRings"  Content="Update Rings"           IsChecked="True" Margin="0,0,14,4" VerticalAlignment="Center"/>
                                <CheckBox x:Name="ChkAiConfigs"    Content="Setting Profiles"       IsChecked="True" Margin="0,0,14,4" VerticalAlignment="Center"/>
                            </WrapPanel>
                            <Button x:Name="BtnScan" Grid.Column="1" Content="Scan Tenant" Style="{StaticResource PrimaryBtn}" VerticalAlignment="Center"/>
                        </Grid>

                        <!-- Status / count line -->
                        <TextBlock x:Name="TxtScanStatus" Grid.Row="1" Margin="0,0,0,8"
                                   Foreground="#555" FontSize="12" Visibility="Collapsed"/>

                        <!-- Check-all / none toolbar (hidden until scan completes) -->
                        <StackPanel x:Name="PanelSelectionBar" Grid.Row="2" Orientation="Horizontal"
                                    Margin="0,0,0,8" Visibility="Collapsed">
                            <TextBlock Text="Select:" VerticalAlignment="Center" Foreground="#555" Margin="0,0,8,0" FontSize="12"/>
                            <Button x:Name="BtnCheckAll"   Content="All"   Style="{StaticResource SmallBtn}" Margin="0,0,6,0"/>
                            <Button x:Name="BtnUncheckAll" Content="None"  Style="{StaticResource SmallBtn}"/>
                        </StackPanel>

                        <!-- Scrollable item list -->
                        <Border Grid.Row="3" BorderBrush="#E0E0E0" BorderThickness="1" CornerRadius="4">
                            <ScrollViewer VerticalScrollBarVisibility="Auto" Padding="4">
                                <StackPanel x:Name="PanelItems"/>
                            </ScrollViewer>
                        </Border>

                    </Grid>
                </TabItem>

                <!-- PAGE 2 · Results -->
                <TabItem>
                    <Grid Margin="28,20">
                        <Grid.RowDefinitions>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="*"/>
                        </Grid.RowDefinitions>
                        <TextBlock x:Name="TxtResultHeading" Grid.Row="0" FontSize="15" FontWeight="SemiBold" Margin="0,0,0,14"/>
                        <ScrollViewer Grid.Row="1" VerticalScrollBarVisibility="Auto">
                            <StackPanel x:Name="PanelResults"/>
                        </ScrollViewer>
                    </Grid>
                </TabItem>

            </TabControl>

            <!-- Loading overlay -->
            <Grid x:Name="LoadingOverlay" Visibility="Collapsed" Background="#CC1A1A1A" Panel.ZIndex="10">
                <StackPanel VerticalAlignment="Center" HorizontalAlignment="Center">
                    <ProgressBar IsIndeterminate="True" Width="260" Height="3" Foreground="White" Background="#444"/>
                    <TextBlock x:Name="TxtLoading" Foreground="White" HorizontalAlignment="Center" FontSize="13" Margin="0,14,0,0"/>
                </StackPanel>
            </Grid>

        </Grid>

        <!-- ═══ FOOTER ═══ -->
        <Border Grid.Row="2" BorderBrush="#E0E0E0" BorderThickness="0,1,0,0" Background="White">
            <Grid Margin="28,0">
                <StackPanel Orientation="Horizontal" HorizontalAlignment="Right" VerticalAlignment="Center">
                    <Button x:Name="BtnDelete" Content="Delete Selected"
                            Style="{StaticResource DangerBtn}" Visibility="Collapsed" Margin="0,0,8,0"/>
                    <Button x:Name="BtnClose"  Content="Close"
                            Style="{StaticResource SecondaryBtn}"/>
                </StackPanel>
            </Grid>
        </Border>

    </Grid>
</Window>
'@

# ── Load window ───────────────────────────────────────────────────────────────

$reader = [System.Xml.XmlNodeReader]::new($xaml)
$window = [System.Windows.Markup.XamlReader]::Load($reader)

function ctrl { param($n) $window.FindName($n) }

# ── State ─────────────────────────────────────────────────────────────────────

# Each entry: @{ Type='Policy'|'UserSetting'|'Group'; Id='...'; Name='...'; Checkbox=$chkControl }
$script:foundItems = [System.Collections.Generic.List[hashtable]]::new()

# ── UI helpers ────────────────────────────────────────────────────────────────

function Show-Loading {
    param([string]$Message = "Please wait...")
    (ctrl 'TxtLoading').Text = $Message
    (ctrl 'LoadingOverlay').Visibility = 'Visible'
    (ctrl 'BtnScan').IsEnabled   = $false
    (ctrl 'BtnDelete').IsEnabled = $false
    $window.Dispatcher.Invoke([System.Windows.Threading.DispatcherPriority]::Render, [Action]{})
}

function Hide-Loading {
    (ctrl 'LoadingOverlay').Visibility = 'Collapsed'
    (ctrl 'BtnScan').IsEnabled   = $true
    (ctrl 'BtnDelete').IsEnabled = $true
}

function Show-Alert {
    param([string]$Message, [string]$Title = "Notice", [string]$Icon = "Warning")
    [System.Windows.MessageBox]::Show($Message, $Title, 'OK', $Icon) | Out-Null
}

function Update-DeleteButton {
    $selected = @($script:foundItems | Where-Object { $_.Checkbox.IsChecked -eq $true })
    $btn = ctrl 'BtnDelete'
    if ($selected.Count -gt 0) {
        $btn.Content    = "Delete Selected ($($selected.Count))"
        $btn.Visibility = 'Visible'
    } else {
        $btn.Visibility = 'Collapsed'
    }
}

# Adds a bold section header to the items panel
function Add-SectionHeader {
    param([string]$Title, [string]$Color = "#333")
    $panel = ctrl 'PanelItems'

    $border = New-Object System.Windows.Controls.Border
    $border.Background = [System.Windows.Media.SolidColorBrush]::new(
        [System.Windows.Media.Color]::FromRgb(0xF3, 0xF3, 0xF3))
    $border.Margin  = [System.Windows.Thickness]::new(0, 4, 0, 2)
    $border.Padding = [System.Windows.Thickness]::new(10, 5, 10, 5)

    $tb = New-Object System.Windows.Controls.TextBlock
    $tb.Text       = $Title
    $tb.FontWeight = 'SemiBold'
    $tb.FontSize   = 12
    $r, $g, $b = [System.Convert]::ToByte($Color.Substring(1,2),16),
                 [System.Convert]::ToByte($Color.Substring(3,2),16),
                 [System.Convert]::ToByte($Color.Substring(5,2),16)
    $tb.Foreground = [System.Windows.Media.SolidColorBrush]::new(
        [System.Windows.Media.Color]::FromRgb($r, $g, $b))

    $border.Child = $tb
    $panel.Children.Add($border) | Out-Null
}

# Adds a checkbox row for one item; returns the CheckBox control
function Add-ItemCheckbox {
    param([string]$Name, [string]$TypeTag)

    $panel = ctrl 'PanelItems'

    $chk = New-Object System.Windows.Controls.CheckBox
    $chk.IsChecked = $true
    $chk.Margin    = [System.Windows.Thickness]::new(10, 3, 10, 3)
    $chk.Tag       = $TypeTag

    $sp = New-Object System.Windows.Controls.StackPanel
    $sp.Orientation = 'Horizontal'

    $nameTb = New-Object System.Windows.Controls.TextBlock
    $nameTb.Text       = $Name
    $nameTb.FontFamily = [System.Windows.Media.FontFamily]::new("Consolas")
    $nameTb.FontSize   = 12
    $nameTb.VerticalAlignment = 'Center'

    $sp.Children.Add($nameTb) | Out-Null
    $chk.Content = $sp

    $chk.Add_Checked({   Update-DeleteButton })
    $chk.Add_Unchecked({ Update-DeleteButton })

    $panel.Children.Add($chk) | Out-Null
    return $chk
}

function Add-ResultRow {
    param([string]$Label, [bool]$Success = $true)
    $panel = ctrl 'PanelResults'
    $tb    = New-Object System.Windows.Controls.TextBlock
    $icon  = if ($Success) { "✅" } else { "❌" }
    $tb.Text         = "$icon  $Label"
    $tb.TextWrapping = 'Wrap'
    $tb.FontFamily   = [System.Windows.Media.FontFamily]::new("Consolas")
    $tb.FontSize     = 11
    $tb.Margin       = [System.Windows.Thickness]::new(0, 2, 0, 2)
    $panel.Children.Add($tb) | Out-Null
}

# ── Scan ─────────────────────────────────────────────────────────────────────

function Invoke-Scan {
    (ctrl 'PanelItems').Children.Clear()
    (ctrl 'BtnDelete').Visibility      = 'Collapsed'
    (ctrl 'PanelSelectionBar').Visibility = 'Collapsed'
    (ctrl 'TxtScanStatus').Visibility  = 'Collapsed'
    $script:foundItems.Clear()

    Show-Loading "Scanning tenant for Windows 365 objects..."

    try {
        if ((ctrl 'ChkPolicies').IsChecked) {
            (ctrl 'TxtLoading').Text = "Scanning provisioning policies..."
            $window.Dispatcher.Invoke([System.Windows.Threading.DispatcherPriority]::Render, [Action]{})

            $policies = Invoke-MgGraphRequest -Method GET `
                -Uri "https://graph.microsoft.com/beta/deviceManagement/virtualEndpoint/provisioningPolicies" `
                -ErrorAction Stop
            $matched = @($policies.value | Where-Object { $_.displayName -like '*-W365-*-*' })

            if ($matched.Count -gt 0) {
                Add-SectionHeader "Provisioning Policies" "#2BC0B8"
                foreach ($p in $matched) {
                    $chk = Add-ItemCheckbox -Name $p.displayName -TypeTag "Policy"
                    $script:foundItems.Add(@{ Type='Policy'; Id=$p.id; Name=$p.displayName; Checkbox=$chk })
                }
            }
        }

        if ((ctrl 'ChkUserSettings').IsChecked) {
            (ctrl 'TxtLoading').Text = "Scanning user settings..."
            $window.Dispatcher.Invoke([System.Windows.Threading.DispatcherPriority]::Render, [Action]{})

            $settings = Invoke-MgGraphRequest -Method GET `
                -Uri "https://graph.microsoft.com/beta/deviceManagement/virtualEndpoint/userSettings" `
                -ErrorAction Stop
            $matched = @($settings.value | Where-Object {
                $_.displayName -in @('W365_AdminSettings','W365_UserSettings','AI_Enabled_Cloud_PC')
            })

            if ($matched.Count -gt 0) {
                Add-SectionHeader "Cloud PC User Settings" "#8764B8"
                foreach ($s in $matched) {
                    $chk = Add-ItemCheckbox -Name $s.displayName -TypeTag "UserSetting"
                    $script:foundItems.Add(@{ Type='UserSetting'; Id=$s.id; Name=$s.displayName; Checkbox=$chk })
                }
            }
        }

        if ((ctrl 'ChkGroups').IsChecked) {
            (ctrl 'TxtLoading').Text = "Scanning Entra ID groups..."
            $window.Dispatcher.Invoke([System.Windows.Threading.DispatcherPriority]::Render, [Action]{})

            $allGroups = Get-AllGraphItems -Uri "https://graph.microsoft.com/v1.0/groups?`$select=id,displayName"
            $matched   = @($allGroups | Where-Object { $_.displayName -like '*W365*' })

            if ($matched.Count -gt 0) {
                Add-SectionHeader "Entra ID Groups" "#107C10"
                foreach ($g in ($matched | Sort-Object displayName)) {
                    $chk = Add-ItemCheckbox -Name $g.displayName -TypeTag "Group"
                    $script:foundItems.Add(@{ Type='Group'; Id=$g.id; Name=$g.displayName; Checkbox=$chk })
                }
            }
        }

        if ((ctrl 'ChkUpdateRings').IsChecked) {
            (ctrl 'TxtLoading').Text = "Scanning Windows Update rings..."
            $window.Dispatcher.Invoke([System.Windows.Threading.DispatcherPriority]::Render, [Action]{})

            # Match rings created by the deploy wizard (description marker) plus the
            # default ring name, in case the description was edited in the portal.
            $rings   = Get-AllGraphItems -Uri "https://graph.microsoft.com/beta/deviceManagement/deviceConfigurations?`$select=id,displayName,description"
            $matched = @($rings | Where-Object {
                $_.description -like '*created by Deploy-Windows365-GUI*' -or
                $_.displayName -eq 'W365-CloudPC-UpdateRing'
            })

            if ($matched.Count -gt 0) {
                Add-SectionHeader "Windows Update Rings" "#B7950B"
                foreach ($r in ($matched | Sort-Object displayName)) {
                    $chk = Add-ItemCheckbox -Name $r.displayName -TypeTag "UpdateRing"
                    $script:foundItems.Add(@{ Type='UpdateRing'; Id=$r.id; Name=$r.displayName; Checkbox=$chk })
                }
            }
        }

        if ((ctrl 'ChkAiConfigs').IsChecked) {
            (ctrl 'TxtLoading').Text = "Scanning Cloud PC setting profiles..."
            $window.Dispatcher.Invoke([System.Windows.Threading.DispatcherPriority]::Render, [Action]{})

            try {
                $profiles = Get-AllGraphItems -Uri "https://graph.microsoft.com/beta/deviceManagement/virtualEndpoint/settingProfiles"
                $matched  = @($profiles | Where-Object { $_.displayName -in @('W365_Frontier_AIEnabled','W365R-Settings') })

                if ($matched.Count -gt 0) {
                    Add-SectionHeader "Cloud PC Setting Profiles" "#C77DBA"
                    foreach ($p in ($matched | Sort-Object displayName)) {
                        $chk = Add-ItemCheckbox -Name $p.displayName -TypeTag "AiConfig"
                        $script:foundItems.Add(@{ Type='AiConfig'; Id=$p.id; Name=$p.displayName; Checkbox=$chk })
                    }
                }
            } catch {
                # settingProfiles endpoint is not available in all tenants — skip quietly
            }
        }

        Hide-Loading

        if ($script:foundItems.Count -eq 0) {
            (ctrl 'TxtScanStatus').Text       = "No matching Windows 365 objects found in this tenant."
            (ctrl 'TxtScanStatus').Foreground = [System.Windows.Media.Brushes]::DimGray
            (ctrl 'TxtScanStatus').Visibility = 'Visible'
            return
        }

        (ctrl 'TxtScanStatus').Text       = "Found $($script:foundItems.Count) object(s). Tick the ones you want to delete, then click Delete Selected."
        (ctrl 'TxtScanStatus').Foreground = [System.Windows.Media.Brushes]::DimGray
        (ctrl 'TxtScanStatus').Visibility = 'Visible'
        (ctrl 'PanelSelectionBar').Visibility = 'Visible'
        Update-DeleteButton
    }
    catch {
        Hide-Loading
        Show-Alert "Failed during scan:`n$_" "Scan Error" "Error"
    }
}

# ── Delete ────────────────────────────────────────────────────────────────────

function Invoke-Delete {
    $toDelete = @($script:foundItems | Where-Object { $_.Checkbox.IsChecked -eq $true })
    if ($toDelete.Count -eq 0) { Show-Alert "No items selected." "Nothing Selected" "Information"; return }

    $names = ($toDelete | ForEach-Object { "  • $($_.Name)" }) -join "`n"
    $confirm = [System.Windows.MessageBox]::Show(
        "Permanently delete these $($toDelete.Count) object(s)?`n`n$names`n`nThis cannot be undone.",
        "Confirm Deletion", 'OKCancel', 'Warning')
    if ($confirm -ne 'OK') { return }

    (ctrl 'PanelResults').Children.Clear()
    (ctrl 'BtnDelete').IsEnabled = $false
    (ctrl 'BtnScan').IsEnabled   = $false

    try {
        foreach ($item in $toDelete) {
            Show-Loading "Deleting: $($item.Name)..."

            switch ($item.Type) {
                'Policy' {
                    try {
                        $assignUri = "https://graph.microsoft.com/beta/deviceManagement/virtualEndpoint/provisioningPolicies/$($item.Id)/assign"
                        Clear-PolicyAssignments -Uri $assignUri
                        Invoke-MgGraphRequest -Method DELETE `
                            -Uri "https://graph.microsoft.com/beta/deviceManagement/virtualEndpoint/provisioningPolicies/$($item.Id)" `
                            -ErrorAction Stop
                        Add-ResultRow "Provisioning Policy deleted: $($item.Name)" -Success $true
                    } catch {
                        Add-ResultRow "FAILED — $($item.Name): $_" -Success $false
                    }
                }
                'UserSetting' {
                    try {
                        $assignUri = "https://graph.microsoft.com/beta/deviceManagement/virtualEndpoint/userSettings/$($item.Id)/assign"
                        Clear-PolicyAssignments -Uri $assignUri
                        Invoke-MgGraphRequest -Method DELETE `
                            -Uri "https://graph.microsoft.com/beta/deviceManagement/virtualEndpoint/userSettings/$($item.Id)" `
                            -ErrorAction Stop
                        Add-ResultRow "User Setting deleted: $($item.Name)" -Success $true
                    } catch {
                        Add-ResultRow "FAILED — $($item.Name): $_" -Success $false
                    }
                }
                'Group' {
                    try {
                        Invoke-MgGraphRequest -Method DELETE `
                            -Uri "https://graph.microsoft.com/v1.0/groups/$($item.Id)" `
                            -ErrorAction Stop
                        Add-ResultRow "Group deleted: $($item.Name)" -Success $true
                    } catch {
                        Add-ResultRow "FAILED — $($item.Name): $_" -Success $false
                    }
                }
                'UpdateRing' {
                    try {
                        Invoke-MgGraphRequest -Method DELETE `
                            -Uri "https://graph.microsoft.com/beta/deviceManagement/deviceConfigurations/$($item.Id)" `
                            -ErrorAction Stop
                        Add-ResultRow "Update Ring deleted: $($item.Name)" -Success $true
                    } catch {
                        Add-ResultRow "FAILED — $($item.Name): $_" -Success $false
                    }
                }
                'AiConfig' {
                    try {
                        $assignUri = "https://graph.microsoft.com/beta/deviceManagement/virtualEndpoint/settingProfiles/$($item.Id)/assign"
                        Clear-PolicyAssignments -Uri $assignUri
                        Invoke-MgGraphRequest -Method DELETE `
                            -Uri "https://graph.microsoft.com/beta/deviceManagement/virtualEndpoint/settingProfiles/$($item.Id)" `
                            -ErrorAction Stop
                        Add-ResultRow "AI Config deleted: $($item.Name)" -Success $true
                    } catch {
                        Add-ResultRow "FAILED — $($item.Name): $_" -Success $false
                    }
                }
            }
        }

        Hide-Loading
        (ctrl 'TxtResultHeading').Text       = "Cleanup Complete"
        (ctrl 'TxtResultHeading').Foreground = [System.Windows.Media.Brushes]::DimGray
        (ctrl 'MainTabs').SelectedIndex      = 2
        Disconnect-MgGraph | Out-Null
    }
    catch {
        Hide-Loading
        (ctrl 'TxtResultHeading').Text       = "Cleanup encountered errors"
        (ctrl 'TxtResultHeading').Foreground = [System.Windows.Media.Brushes]::Crimson
        $errTb = New-Object System.Windows.Controls.TextBlock
        $errTb.Text         = "Unexpected error: $_"
        $errTb.TextWrapping = 'Wrap'
        $errTb.Foreground   = [System.Windows.Media.Brushes]::Crimson
        (ctrl 'PanelResults').Children.Add($errTb) | Out-Null
        (ctrl 'MainTabs').SelectedIndex = 2
    }
}

# ── Event handlers ────────────────────────────────────────────────────────────

(ctrl 'BtnConnect').Add_Click({
    Show-Loading "Connecting to Microsoft Graph..."
    try {
        Install-GraphModuleIfNeeded
        Connect-MgGraph -Scopes "CloudPC.ReadWrite.All","Group.ReadWrite.All","DeviceManagementConfiguration.ReadWrite.All" -ErrorAction Stop
        (ctrl 'TxtConnectionStatus').Text       = [char]0x2714 + "  Connected to Microsoft Graph"
        (ctrl 'TxtConnectionStatus').Foreground = [System.Windows.Media.SolidColorBrush]::new(
            [System.Windows.Media.Color]::FromRgb(16, 124, 16))
        (ctrl 'BtnConnect').IsEnabled   = $false
        Hide-Loading
        (ctrl 'MainTabs').SelectedIndex = 1
    } catch {
        (ctrl 'TxtConnectionStatus').Text       = [char]0x2718 + "  Connection failed — $_"
        (ctrl 'TxtConnectionStatus').Foreground = [System.Windows.Media.Brushes]::Crimson
        Hide-Loading
    }
})

(ctrl 'BtnScan').Add_Click({ Invoke-Scan })

(ctrl 'BtnCheckAll').Add_Click({
    $script:foundItems | ForEach-Object { $_.Checkbox.IsChecked = $true }
    Update-DeleteButton
})

(ctrl 'BtnUncheckAll').Add_Click({
    $script:foundItems | ForEach-Object { $_.Checkbox.IsChecked = $false }
    Update-DeleteButton
})

(ctrl 'BtnDelete').Add_Click({ Invoke-Delete })
(ctrl 'BtnClose').Add_Click({  $window.Close() })

# ── Launch ────────────────────────────────────────────────────────────────────
(ctrl 'MainTabs').SelectedIndex = 0
$window.ShowDialog() | Out-Null
