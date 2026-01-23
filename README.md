# ImplementWindows365

A comprehensive PowerShell script to deploy and configure Windows 365 Cloud PC environments in Microsoft Azure, including Azure AD security groups, user settings policies, and provisioning policies.

## Overview

`Deploy-Windows365.ps1` automates the complete setup of Windows 365 Cloud PC infrastructure. It handles:

- Azure AD security group creation/reuse for user and admin roles
- Cloud PC user settings policies with local admin configuration
- Cloud PC provisioning policies with regional deployment
- Intelligent assignment preservation to avoid overwriting existing configurations

## Prerequisites

### Required Permissions

Your Microsoft Entra ID (Azure AD) account must have the following Microsoft Graph API scopes:
- `User.ReadWrite.All` - For user management
- `Application.ReadWrite.All` - For application management
- `CloudPC.ReadWrite.All` - For Windows 365 Cloud PC management
- `Group.ReadWrite.All` - For Azure AD group management

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
1. Select a Windows 365 Cloud PC SKU (Enterprise plan options)
2. Select a region group (Americas, Asia, Europe, etc.)
3. Select a specific region within that group
4. Select a Windows 11 device image

### Usage with Parameters

```powershell
.\Deploy-Windows365.ps1 -CloudPCTypeChoice 1 -RegionChoice 1
```

**Parameters:**
- `-CloudPCTypeChoice` (int): SKU selection (1-based index from available plans)
- `-RegionChoice` (int): Region selection (1-based index from available regions)
  - Note: Region selection is always interactive - use the two-step prompt for specific region selection

### With Verbose Output

```powershell
.\Deploy-Windows365.ps1 -Verbose
```

## What the Script Creates

### Azure AD Security Groups

The script creates two security groups per Cloud PC type:

- **`GRP_Users_[CloudPCType]`** - Users who need standard Cloud PC access (no local admin)
- **`GRP_Admins_[CloudPCType]`** - Admin users with local administrator rights on Cloud PCs

**Note:** You must assign the appropriate Windows 365 licenses to these groups after creation.

### Cloud PC User Settings

Creates two user setting policies:

- **`W365_AdminSettings`** - Assigned to admin group with local admin enabled
- **`W365_UserSettings`** - Assigned to user group with local admin disabled

**Default Settings:**
- Restore point enabled with 12-hour frequency
- SSO enabled where available

### Cloud PC Provisioning Policy

Creates a provisioning policy named:
```
[RegionName]-W365-Enterprise-Provisioning Policy
```

**Configuration:**
- Provisioning type: Dedicated
- User experience: Cloud PC (full desktop)
- Domain join: Azure AD join
- Image: Selected Windows 11 enterprise image
- Windows language: English (US)
- Assigned to: Both user and admin groups

## Key Features

### Intelligent Assignment Management

The script preserves existing assignments when updating policies. Instead of replacing all assignments (as the Graph API's `/assign` endpoint does), the script:

1. Retrieves existing assignments
2. Merges new groups with existing assignments
3. Applies the complete merged list
4. Prevents accidental removal of user access

### Automatic Module Installation

If Microsoft Graph modules are not found, the script automatically installs them with administrator confirmation.

### Robust Error Handling

- Graceful fallback from PowerShell cmdlets to direct Microsoft Graph API calls
- Detailed verbose logging for troubleshooting
- Pre-validation of group IDs before assignment
- Safe handling of API response variations

### Group Replication Management

The script includes built-in delays to account for Azure AD group replication latency before assigning policies.

## Output

The script displays:
- Status messages (creation, reuse, assignment)
- Selected configuration (Cloud PC type, region, image)
- Group IDs for reference
- Completion confirmation with next steps

### Important Reminder

After successful script execution, remember to assign Windows 365 licenses:
```
License Assignment Required
- GRP_Users_[CloudPCType]
- GRP_Admins_[CloudPCType]
```

## Troubleshooting

### "No Cloud PC service plans found"

- Verify your Windows 365 subscription is active
- Check that you have the `CloudPC.ReadWrite.All` scope
- Ensure your tenant has at least one Cloud PC SKU available

### "No supported region groups returned"

- Verify Windows 365 is available in your region
- Check Microsoft Graph permissions
- Some regions may not support Windows 365

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
| Azure AD Role | At minimum: Groups Admin, Cloud PC Admin |
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
4. Review required permissions in Azure AD admin center

## Author

Jon Jarvis

## License

[Add your license information here]
