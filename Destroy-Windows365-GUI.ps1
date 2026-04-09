#Requires -Version 7.0
<#
.SYNOPSIS
    Windows 365 Cleanup Tool — WPF graphical front-end.
.DESCRIPTION
    GUI wrapper for Destroy-Windows365.ps1. Connects to Microsoft Graph,
    previews what will be deleted, then deletes on explicit confirmation.
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
    Title="Windows 365 Cleanup Tool"
    Width="700" Height="620"
    MinWidth="700" MinHeight="620"
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

        <Style TargetType="CheckBox">
            <Setter Property="Margin" Value="0,4,0,4"/>
        </Style>

        <Style TargetType="TextBox">
            <Setter Property="Padding"     Value="6,5"/>
            <Setter Property="BorderBrush" Value="#D0D0D0"/>
        </Style>

    </Window.Resources>

    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="72"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="56"/>
        </Grid.RowDefinitions>

        <!-- ═══ HEADER ═══ -->
        <Border Grid.Row="0" Background="#C50F1F">
            <Grid Margin="28,0">
                <StackPanel VerticalAlignment="Center">
                    <TextBlock Text="Windows 365 Cleanup Tool" Foreground="White" FontSize="17" FontWeight="SemiBold"/>
                    <TextBlock Text="Permanently removes provisioning policies, user settings, and groups from this tenant."
                               Foreground="#F7C0C5" FontSize="12" Margin="0,3,0,0" TextWrapping="Wrap"/>
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
                            Text="Sign in with an account that has CloudPC.ReadWrite.All and Group.ReadWrite.All permissions."/>
                        <Button x:Name="BtnConnect" Content="Sign in to Microsoft Graph" Style="{StaticResource PrimaryBtn}" HorizontalAlignment="Left"/>
                        <TextBlock x:Name="TxtConnectionStatus" Margin="0,10,0,0" FontSize="13" TextWrapping="Wrap"/>
                    </StackPanel>
                </TabItem>

                <!-- PAGE 1 · Configure + Preview -->
                <TabItem>
                    <Grid Margin="28,20">
                        <Grid.RowDefinitions>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="*"/>
                        </Grid.RowDefinitions>

                        <!-- What to remove -->
                        <StackPanel Grid.Row="0" Margin="0,0,0,16">
                            <TextBlock Text="What to Remove" FontSize="15" FontWeight="SemiBold" Margin="0,0,0,12"/>
                            <Border BorderBrush="#E0E0E0" BorderThickness="1" CornerRadius="4" Padding="16,12">
                                <StackPanel>
                                    <CheckBox x:Name="ChkPolicies"     Content="Provisioning Policies  (names matching  *-W365-*-*)"  IsChecked="True"/>
                                    <CheckBox x:Name="ChkUserSettings" Content="Cloud PC User Settings  (W365_AdminSettings, W365_UserSettings, AI_Enabled_Cloud_PC)" IsChecked="True"/>
                                    <CheckBox x:Name="ChkGroups"       Content="Entra ID Security Groups  (display names containing  W365)" IsChecked="True"/>
                                </StackPanel>
                            </Border>

                            <!-- Exclusions -->
                            <Expander Header="Exclusions — skip specific items" Margin="0,10,0,0" FontSize="12">
                                <Border BorderBrush="#E8E8E8" BorderThickness="1" CornerRadius="4" Margin="0,6,0,0" Padding="14,10">
                                    <StackPanel>
                                        <TextBlock Text="Policies to keep (comma-separated display names)" Foreground="#555" Margin="0,0,0,4"/>
                                        <TextBox x:Name="TxtKeepPolicies"      Margin="0,0,0,10"/>
                                        <TextBlock Text="User settings to keep (comma-separated display names)" Foreground="#555" Margin="0,0,0,4"/>
                                        <TextBox x:Name="TxtKeepUserSettings"  Margin="0,0,0,10"/>
                                        <TextBlock Text="Groups to keep (comma-separated display names)" Foreground="#555" Margin="0,0,0,4"/>
                                        <TextBox x:Name="TxtKeepGroups"/>
                                    </StackPanel>
                                </Border>
                            </Expander>

                            <Button x:Name="BtnPreview" Content="Preview Changes" Style="{StaticResource SecondaryBtn}"
                                    HorizontalAlignment="Left" Margin="0,14,0,0"/>
                        </StackPanel>

                        <!-- Preview results -->
                        <StackPanel Grid.Row="1" Margin="0,0,0,8">
                            <TextBlock x:Name="TxtPreviewHeading" FontWeight="SemiBold" Foreground="#444" Visibility="Collapsed"/>
                        </StackPanel>

                        <Border Grid.Row="2" BorderBrush="#E0E0E0" BorderThickness="1" CornerRadius="4" Visibility="{Binding ElementName=TxtPreviewHeading, Path=Visibility}">
                            <ScrollViewer VerticalScrollBarVisibility="Auto" Padding="4">
                                <StackPanel x:Name="PanelPreview"/>
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
                    <Button x:Name="BtnDelete" Content="Delete Selected  &#10007;"
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

$script:isConnected = $false

$script:previewData = @{
    Policies     = @()
    UserSettings = @()
    Groups       = @()
}

# ── UI helpers ────────────────────────────────────────────────────────────────

function Show-Loading {
    param([string]$Message = "Please wait...")
    (ctrl 'TxtLoading').Text = $Message
    (ctrl 'LoadingOverlay').Visibility = 'Visible'
    (ctrl 'BtnPreview').IsEnabled = $false
    (ctrl 'BtnDelete').IsEnabled  = $false
    $window.Dispatcher.Invoke([System.Windows.Threading.DispatcherPriority]::Render, [Action]{})
}

function Hide-Loading {
    (ctrl 'LoadingOverlay').Visibility = 'Collapsed'
    (ctrl 'BtnPreview').IsEnabled = $true
    (ctrl 'BtnDelete').IsEnabled  = $true
}

function Show-Alert {
    param([string]$Message, [string]$Title = "Notice", [string]$Icon = "Warning")
    [System.Windows.MessageBox]::Show($Message, $Title, 'OK', $Icon) | Out-Null
}

function Add-PreviewItem {
    param([string]$Category, [string]$Name, [string]$CategoryColor = "#0078D4")
    $panel = ctrl 'PanelPreview'
    $row   = New-Object System.Windows.Controls.Grid
    $c0    = New-Object System.Windows.Controls.ColumnDefinition; $c0.Width = [System.Windows.GridLength]::new(160)
    $c1    = New-Object System.Windows.Controls.ColumnDefinition; $c1.Width = [System.Windows.GridLength]::new(1, [System.Windows.GridUnitType]::Star)
    $row.ColumnDefinitions.Add($c0); $row.ColumnDefinitions.Add($c1)
    $row.Margin = [System.Windows.Thickness]::new(4, 2, 4, 2)

    $catTb = New-Object System.Windows.Controls.TextBlock
    $catTb.Text       = $Category
    $catTb.FontSize   = 11
    $r, $g, $b = [System.Convert]::ToByte($CategoryColor.Substring(1,2), 16), [System.Convert]::ToByte($CategoryColor.Substring(3,2), 16), [System.Convert]::ToByte($CategoryColor.Substring(5,2), 16)
    $catTb.Foreground = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromRgb($r, $g, $b))
    $catTb.VerticalAlignment = 'Top'; $catTb.Margin = [System.Windows.Thickness]::new(0,2,0,0)
    [System.Windows.Controls.Grid]::SetColumn($catTb, 0)

    $nameTb = New-Object System.Windows.Controls.TextBlock
    $nameTb.Text         = $Name
    $nameTb.FontFamily   = [System.Windows.Media.FontFamily]::new("Consolas")
    $nameTb.FontSize     = 11
    $nameTb.TextWrapping = 'Wrap'
    $nameTb.VerticalAlignment = 'Top'; $nameTb.Margin = [System.Windows.Thickness]::new(0,2,0,0)
    [System.Windows.Controls.Grid]::SetColumn($nameTb, 1)

    $row.Children.Add($catTb)  | Out-Null
    $row.Children.Add($nameTb) | Out-Null
    $panel.Children.Add($row)  | Out-Null
}

function Add-ResultRow {
    param([string]$Label, [bool]$Success = $true)
    $panel = ctrl 'PanelResults'
    $tb    = New-Object System.Windows.Controls.TextBlock
    $icon  = if ($Success) { "✅" } else { "❌" }
    $tb.Text        = "$icon  $Label"
    $tb.TextWrapping = 'Wrap'
    $tb.FontFamily  = [System.Windows.Media.FontFamily]::new("Consolas")
    $tb.FontSize    = 11
    $tb.Margin      = [System.Windows.Thickness]::new(0, 2, 0, 2)
    $panel.Children.Add($tb) | Out-Null
}

# ── Get exclusion lists ───────────────────────────────────────────────────────

function Get-ExclusionList {
    param([string]$Raw)
    return @($Raw -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
}

# ── Preview ───────────────────────────────────────────────────────────────────

function Invoke-Preview {
    (ctrl 'PanelPreview').Children.Clear()
    (ctrl 'BtnDelete').Visibility = 'Collapsed'
    $script:previewData = @{ Policies = @(); UserSettings = @(); Groups = @() }

    $keepPolicies     = Get-ExclusionList (ctrl 'TxtKeepPolicies').Text
    $keepUserSettings = Get-ExclusionList (ctrl 'TxtKeepUserSettings').Text
    $keepGroups       = Get-ExclusionList (ctrl 'TxtKeepGroups').Text

    $totalFound = 0

    Show-Loading "Scanning tenant for Windows 365 objects..."

    try {
        if ((ctrl 'ChkPolicies').IsChecked) {
            (ctrl 'TxtLoading').Text = "Scanning provisioning policies..."
            $window.Dispatcher.Invoke([System.Windows.Threading.DispatcherPriority]::Render, [Action]{})
            $policies = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/beta/deviceManagement/virtualEndpoint/provisioningPolicies" -ErrorAction Stop
            $targets  = @($policies.value | Where-Object { $_.displayName -like '*-W365-*-*' })
            if ($keepPolicies.Count -gt 0) { $targets = $targets | Where-Object { $_.displayName -notin $keepPolicies } }
            $script:previewData.Policies = @($targets)
            $totalFound += $targets.Count
        }

        if ((ctrl 'ChkUserSettings').IsChecked) {
            (ctrl 'TxtLoading').Text = "Scanning user settings..."
            $window.Dispatcher.Invoke([System.Windows.Threading.DispatcherPriority]::Render, [Action]{})
            $settings = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/beta/deviceManagement/virtualEndpoint/userSettings" -ErrorAction Stop
            $targets  = @($settings.value | Where-Object { $_.displayName -in @('W365_AdminSettings','W365_UserSettings','AI_Enabled_Cloud_PC') })
            if ($keepUserSettings.Count -gt 0) { $targets = $targets | Where-Object { $_.displayName -notin $keepUserSettings } }
            $script:previewData.UserSettings = @($targets)
            $totalFound += $targets.Count
        }

        if ((ctrl 'ChkGroups').IsChecked) {
            (ctrl 'TxtLoading').Text = "Scanning Entra ID groups..."
            $window.Dispatcher.Invoke([System.Windows.Threading.DispatcherPriority]::Render, [Action]{})
            $allGroups = Get-AllGraphItems -Uri "https://graph.microsoft.com/v1.0/groups?`$select=id,displayName"
            $targets   = @($allGroups | Where-Object { $_.displayName -like '*W365*' })
            if ($keepGroups.Count -gt 0) { $targets = $targets | Where-Object { $_.displayName -notin $keepGroups } }
            $script:previewData.Groups = @($targets)
            $totalFound += $targets.Count
        }

        Hide-Loading

        if ($totalFound -eq 0) {
            (ctrl 'TxtPreviewHeading').Text       = "Nothing to delete — no matching objects found."
            (ctrl 'TxtPreviewHeading').Foreground = [System.Windows.Media.Brushes]::DimGray
            (ctrl 'TxtPreviewHeading').Visibility = 'Visible'
            return
        }

        (ctrl 'TxtPreviewHeading').Text       = "The following $totalFound object(s) will be permanently deleted:"
        (ctrl 'TxtPreviewHeading').Foreground = [System.Windows.Media.Brushes]::Crimson
        (ctrl 'TxtPreviewHeading').Visibility = 'Visible'

        foreach ($p in $script:previewData.Policies) {
            Add-PreviewItem "Provisioning Policy" $p.displayName "#0078D4"
        }
        foreach ($s in $script:previewData.UserSettings) {
            Add-PreviewItem "User Setting" $s.displayName "#8764B8"
        }
        foreach ($g in $script:previewData.Groups) {
            Add-PreviewItem "Entra ID Group" $g.displayName "#107C10"
        }

        (ctrl 'BtnDelete').Visibility = 'Visible'
    }
    catch {
        Hide-Loading
        Show-Alert "Failed to load preview:`n$_" "Error" "Error"
    }
}

# ── Delete ────────────────────────────────────────────────────────────────────

function Invoke-Delete {
    $totalItems = $script:previewData.Policies.Count + $script:previewData.UserSettings.Count + $script:previewData.Groups.Count
    if ($totalItems -eq 0) { Show-Alert "Nothing to delete." "Nothing to Delete" "Information"; return }

    $confirm = [System.Windows.MessageBox]::Show(
        "You are about to permanently delete $totalItems object(s) from this tenant.`n`nThis cannot be undone.`n`nType OK to proceed.",
        "Confirm Deletion", 'OKCancel', 'Warning')
    if ($confirm -ne 'OK') { return }

    (ctrl 'PanelResults').Children.Clear()
    (ctrl 'BtnDelete').IsEnabled  = $false
    (ctrl 'BtnPreview').IsEnabled = $false

    try {
        # Provisioning policies
        foreach ($p in $script:previewData.Policies) {
            Show-Loading "Removing policy: $($p.displayName)..."
            try {
                $assignUri = "https://graph.microsoft.com/beta/deviceManagement/virtualEndpoint/provisioningPolicies/$($p.id)/assign"
                Clear-PolicyAssignments -Uri $assignUri
                Invoke-MgGraphRequest -Method DELETE -Uri "https://graph.microsoft.com/beta/deviceManagement/virtualEndpoint/provisioningPolicies/$($p.id)" -ErrorAction Stop
                Add-ResultRow "Policy deleted: $($p.displayName)" -Success $true
            } catch {
                Add-ResultRow "FAILED to delete policy: $($p.displayName) — $_" -Success $false
            }
        }

        # User settings
        foreach ($s in $script:previewData.UserSettings) {
            Show-Loading "Removing user setting: $($s.displayName)..."
            try {
                $assignUri = "https://graph.microsoft.com/beta/deviceManagement/virtualEndpoint/userSettings/$($s.id)/assign"
                Clear-PolicyAssignments -Uri $assignUri
                Invoke-MgGraphRequest -Method DELETE -Uri "https://graph.microsoft.com/beta/deviceManagement/virtualEndpoint/userSettings/$($s.id)" -ErrorAction Stop
                Add-ResultRow "User setting deleted: $($s.displayName)" -Success $true
            } catch {
                Add-ResultRow "FAILED to delete user setting: $($s.displayName) — $_" -Success $false
            }
        }

        # Groups
        foreach ($g in $script:previewData.Groups) {
            Show-Loading "Removing group: $($g.displayName)..."
            try {
                Invoke-MgGraphRequest -Method DELETE -Uri "https://graph.microsoft.com/v1.0/groups/$($g.id)" -ErrorAction Stop
                Add-ResultRow "Group deleted: $($g.displayName)" -Success $true
            } catch {
                Add-ResultRow "FAILED to delete group: $($g.displayName) — $_" -Success $false
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
        $errTb.Text        = "Unexpected error: $_"
        $errTb.TextWrapping = 'Wrap'; $errTb.Foreground = [System.Windows.Media.Brushes]::Crimson
        (ctrl 'PanelResults').Children.Add($errTb) | Out-Null
        (ctrl 'MainTabs').SelectedIndex = 2
    }
}

# ── Event handlers ────────────────────────────────────────────────────────────

(ctrl 'BtnConnect').Add_Click({
    Show-Loading "Connecting to Microsoft Graph..."
    try {
        Install-GraphModuleIfNeeded
        Connect-MgGraph -Scopes "CloudPC.ReadWrite.All","Group.ReadWrite.All" -ErrorAction Stop
        $script:isConnected = $true
        (ctrl 'TxtConnectionStatus').Text       = [char]0x2714 + "  Connected to Microsoft Graph"
        (ctrl 'TxtConnectionStatus').Foreground = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromRgb(16, 124, 16))
        (ctrl 'BtnConnect').IsEnabled           = $false
        Hide-Loading
        (ctrl 'MainTabs').SelectedIndex         = 1
    } catch {
        (ctrl 'TxtConnectionStatus').Text       = [char]0x2718 + "  Connection failed — $_"
        (ctrl 'TxtConnectionStatus').Foreground = [System.Windows.Media.Brushes]::Crimson
        Hide-Loading
    }
})

(ctrl 'BtnPreview').Add_Click({ Invoke-Preview })
(ctrl 'BtnDelete').Add_Click({  Invoke-Delete })
(ctrl 'BtnClose').Add_Click({   $window.Close() })

# ── Launch ────────────────────────────────────────────────────────────────────
(ctrl 'MainTabs').SelectedIndex = 0
$window.ShowDialog() | Out-Null
