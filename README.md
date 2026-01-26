# ImplementWindows365

A comprehensive PowerShell script to deploy and configure Windows 365 Cloud PC environments in Microsoft Azure, including Entra ID security groups, user settings policies, and provisioning policies.

## Overview

`Deploy-Windows365.ps1` automates the complete setup of Windows 365 Cloud PC infrastructure for both **Enterprise** and **Frontline**. It handles:

- Entra ID security group creation/reuse for licensing, user, and admin roles (with new standardized naming)
- Cloud PC user settings policies with local admin configuration and persistence options for Frontline Shared
- Cloud PC provisioning policies with regional deployment (Enterprise assigned to groups; Frontline auto-assigned via license)
- Intelligent assignment preservation for Enterprise to avoid overwriting existing configurations

## Prerequisites

### Required Permissions

Your Microsoft Entra ID account must have the following Microsoft Graph API scopes:
- `User.ReadWrite.All` - For user management
- `Application.ReadWrite.All` - For application management
- `CloudPC.ReadWrite.All` - For Windows 365 Cloud PC management
- `Group.ReadWrite.All` - For Entra ID group management

### Required Software

- **PowerShell 5.0+** (tested on PowerShell 7.x)
- **Microsoft.Graph** module (script will auto-install if missing)
- **Microsoft.Graph.Beta.DeviceManagement.Administration** module (v2.23.0+, recommended but not required - script has fallback)

### Tenant Requirements

- Active Windows 365 subscription
- At least one available Cloud PC SKU/service plan
- Supported regions with Windows 365 availability
- At least one device image available in your tenant

## Installation

1. Clone or download this repository
2. Open PowerShell as Administrator
3. Navigate to the script directory
4. (Optional) Set execution policy if needed:
   ```powershell
   Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
   ```

## Usage

### Basic Usage (Interactive)

```powershell
.\Deploy-Windows365.ps1
```

The script will prompt you to:
1. Choose license type: **Enterprise** or **Frontline** (then Dedicated vs Shared for Frontline)
2. Select a Windows 365 Cloud PC SKU (filtered to Enterprise or Frontline plans based on your choice)
3. Select a region group (two-step: region group, then specific region)
4. Select a Windows 11 device image (unsupported images are filtered; warnings are allowed)
5. Select a language (20 options; falls back to en-GB with a warning if the selection is rejected by Graph)

### Usage with Parameters

```powershell
.\Deploy-Windows365.ps1 -CloudPCTypeChoice 1 -RegionChoice 1
```

**Parameters:**
- `-CloudPCTypeChoice` (int): SKU selection (1-based index from available plans)
- `-RegionChoice` (int): Region selection (1-based index from available regions)
   - Note: Region selection remains interactive (two-step prompt) when not provided

### With Verbose Output

```powershell
.\Deploy-Windows365.ps1 -Verbose
```

## What the Script Creates

### Entra ID Security Groups

The script creates/reuses three security groups with standardized names:

1. **Licensing (all users/admins for SKU)**
   - `SG-W365CloudPC-<Cloud PC Type>`
   - Assign Windows 365 licenses to this group (required)

2. **User settings group (per region)**
   - Enterprise: `SG-W365ENT-<Region>-User`
   - Frontline: `SG-W365FL-<Region>-User`
   - Assigned to the standard user settings (no local admin rights)

3. **Admin settings group (per region)**
   - Enterprise: `SG-W365ENT-<Region>-Admin`
   - Frontline: `SG-W365FL-<Region>-Admin`
   - Assigned to the admin settings (with local admin rights)

**Important:** You must assign the Windows 365 license to `SG-W365CloudPC-<Cloud PC Type>` after creation. Frontline provisioning policies are applied via license assignment (no group /assign step).

### Cloud PC User Settings

Creates two user setting policies:

- **`W365_AdminSettings`** - Assigned to `[Location]_Windows365_LocalAdmin` group with local admin enabled
- **`W365_UserSettings`** - Assigned to `[Location]_Windows365_User` group with local admin disabled

**Default Settings:**
- Reset enabled; restore point frequency 6 hours
- DR settings created but disabled by default (configure manually if needed)
- SSO enabled where available
- Language fallback: if the chosen language is rejected, the policy is created with en-GB and a warning is shown
- AI option: if the selected SKU is Copilot-eligible (≥8 vCPU/32GB/256GB), `AI_Enabled_Cloud_PC` is created and assigned to the licensing group

### Cloud PC Provisioning Policy

Creates a provisioning policy named:
```
<RegionName>-W365-<LicenseType>-Provisioning Policy
```

**Configuration:**
- Provisioning type: 
   - Enterprise: dedicated
   - Frontline Dedicated: sharedByUser
   - Frontline Shared: sharedByEntraGroup (with user settings persistence enabled)
- User experience: Cloud PC (full desktop)
- Domain join: Entra ID join
- Image: Selected Windows 11 enterprise image (supported or supportedWithWarning)
- Windows language: Configurable (en-GB default, 20+ languages supported; fallback to en-GB on validation failure)
- Assignments:
   - Enterprise: assigned to user/admin groups via /assign (merged to preserve existing)
   - Frontline: no /assign; policy applies when licenses are assigned

## Key Features

### Intelligent Assignment Management

- **Enterprise:** Retrieves existing assignments, merges, and applies to avoid overwriting.
- **Frontline:** Skips `/assign`; policies apply automatically when Frontline licenses are assigned.

### Automatic Module Installation

If Microsoft Graph modules are not found, the script automatically installs them with administrator confirmation.

### Robust Error Handling

- Graceful fallback from PowerShell cmdlets to direct Microsoft Graph API calls
- Detailed verbose logging for troubleshooting
- Pre-validation of group IDs before assignment
- Safe handling of API response variations

### Group Replication Management

The script includes built-in delays to account for Entra ID group replication latency before assigning policies.

## Output

The script displays:
- Status messages (creation, reuse, assignment)
- Selected configuration (Cloud PC type, region, image)
- Group IDs for reference
- Completion confirmation with next steps

### Important Reminder

After successful script execution, remember to assign Windows 365 licenses to:

- `SG-W365CloudPC-<Cloud PC Type>` (licensing group)
- Frontline: ensure users receive Frontline licenses; the provisioning policy applies via license

## Troubleshooting

### "No Cloud PC service plans found"

- Verify your Windows 365 subscription is active
- Check that you have the `CloudPC.ReadWrite.All` scope
- Ensure your tenant has at least one Cloud PC SKU available

### "No supported region groups returned"

- Verify Windows 365 is available in your region
- Check Microsoft Graph permissions
- Some regions may not support Windows 365

### "Only unsupported images are available"

- This typically means no compatible images are available in your tenant
- Check the Microsoft 365 admin center for image status
- Wait for image processing to complete if recently uploaded
- Contact Microsoft support if images remain unavailable

### "Failed to install Microsoft.Graph module"

- Run PowerShell as Administrator
- Check your internet connection
- Verify you have permission to install modules
- Check proxy/firewall settings if behind a corporate proxy

### "Group creation failed"

- Verify `Group.ReadWrite.All` permission
- Ensure the group names are not already in use
- Check if you have permissions in the tenant

### "Script hangs on region selection"

- Ensure you're entering numeric choices (not letters)
- Try running with `-Verbose` for more details
- Press `Ctrl+C` to cancel and try again

## Advanced Usage

### Verbose Logging

For detailed troubleshooting, enable PowerShell verbose output:

```powershell
$VerbosePreference = "Continue"
.\Deploy-Windows365.ps1
```

### Direct Graph API Fallback

The script automatically falls back to direct Microsoft Graph REST API calls if PowerShell cmdlets are unavailable. This ensures maximum compatibility across different module versions.

### Idempotent Execution

The script is safe to run multiple times:
- If groups exist, they're reused
- If policies exist, they're reused and assignments are merged
- No duplicate objects are created

## Requirements Summary

| Requirement | Version/Details |
|---|---|
| PowerShell | 5.0+ (7.x recommended) |
| Microsoft.Graph | Latest stable |
| Entra ID Role | At minimum: Groups Admin, Cloud PC Admin |
| Windows 365 License | Required for group assignment |
| API Scopes | User.ReadWrite.All, Application.ReadWrite.All, CloudPC.ReadWrite.All, Group.ReadWrite.All |

## Notes

- The script uses the Microsoft Graph Beta API for Windows 365 resources
- Policies are created as "Gallery Image" type
- Single sign-on (SSO) is enabled by default
- User settings persistence is disabled by default
- The script preserves existing assignments to prevent data loss

## Support

For issues or questions:
1. Run with `-Verbose` flag for detailed logging
2. Check Microsoft Graph health dashboard
3. Verify all prerequisites are met
4. Review required permissions in Entra ID admin center

## Author

Jon Jarvis

## License

[Add your license information here]
