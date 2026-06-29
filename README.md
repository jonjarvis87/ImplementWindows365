# Windows 365 Toolkit

**by [CloudEndpoint.AI](https://www.cloudendpoint.ai)**

A pair of WPF GUI wizards for deploying and cleaning up Windows 365 Cloud PC environments — no PowerShell experience required. Built for IT professionals managing Microsoft Intune and Windows 365.

---

## Tools

| Tool | Script | Purpose |
|---|---|---|
| Deployment Wizard | `Deploy-Windows365-GUI.ps1` | Step-by-step wizard to build a complete Windows 365 environment |
| Cleanup Tool | `Destroy-Windows365-GUI.ps1` | Scan your tenant and selectively remove Windows 365 objects |

---

## Deployment Wizard

`Deploy-Windows365-GUI.ps1` walks you through an 8-step wizard to configure a production-ready Windows 365 environment. It handles both **Enterprise** and **Frontline** licence types.

### What it creates

| Object | Details |
|---|---|
| Entra ID security groups | Licensing, User, Admin, and dynamic Devices groups |
| Cloud PC user settings | `W365_AdminSettings` (local admin on) and `W365_UserSettings` (local admin off) |
| Cloud PC provisioning policy | Named `<Region>-W365-<Type>-Policy`, assigned to user/admin groups |
| Windows Update for Business ring | Optional — Broad, Balanced, or Targeted preset profiles |
| AI Cloud PC config | Created automatically if the selected SKU is Copilot-eligible (≥ 8 vCPU / 32 GB / 256 GB) and the region supports it |
| Autopilot device preparation profile | Optional — linked to the provisioning policy if selected |

### Wizard steps

1. **Sign in** — Connect to Microsoft Graph and choose Enterprise or Frontline
2. **Cloud PC SKU** — Select the service plan (filtered to your licence type)
3. **Region** — Pick region group and specific region
4. **Image** — Select a Windows 11 gallery image
5. **Language** — Choose the Windows language (42 options; falls back to en-GB if rejected)
6. **Windows Update** — Configure a WUfB ring or skip
7. **Autopilot & Naming** — Optional Autopilot profile and device naming template
8. **Review & Deploy** — Summary of all selections before committing

### Licence types

**Enterprise**
Each user gets their own Cloud PC without restrictions on when they can connect.
- Group-based licence assignment is configured automatically
- Provisioning policy assigned to user and admin groups

**Frontline — Dedicated**
Recommended for part-time or shift workers. A single licence provisions up to three Cloud PCs used non-concurrently, each assigned to a single user.

**Frontline — Shared**
Recommended for short-session users who do not require data persistence. A single licence provisions one Cloud PC shared non-concurrently among a group.

For Frontline, session assignment to the provisioning policy is **optional** in the wizard. When enabled, you set:
- **Assignment name** — label shown in the Intune portal for the allotment
- **Number of sessions** — sessions to reserve from your Frontline licence pool

**Reserve**
Short-term, dedicated Cloud PCs for business continuity — when a primary device is lost, broken, or unavailable. Reserve differs from the other types:
- **Fixed size** — 4 vCPU / 16 GB / 128 GB. The SKU page is skipped.
- **Automatic image** — Reserve provisions the latest gallery image (`imageId: automatic`), so the Image page is skipped too. Pick a specific image later in Intune if needed.
- **Geography only** — you pick a geography, and the service auto-selects the region within it (the wizard's region page collapses to a single geography list). The geography is written to `domainJoinConfigurations.geographicLocationType`.
- **On-demand provisioning** — Cloud PCs are *not* created automatically. The wizard sets up the policy, groups, and licence assignment; you then provision per-user from Intune when cover is needed.
- **Two groups, no extras** — Reserve creates just a User and an Admin group (`SG-W365R-<geography>-User` / `-Admin`), assigns both to the policy, and licenses both. There is no separate licensing or devices group, and the **Windows Update ring and Autopatch pages are skipped** (not applicable to Reserve's ephemeral Cloud PCs).
- Up to 10 days of Cloud PC access per user per year. A user's Cloud PC is eligible to provision **7 days after** their Reserve licence is assigned.

The wizard creates and assigns the Reserve provisioning policy (`provisioningType: reserve`) and attempts group-based licensing of the Reserve SKU. The post-deploy checklist (Copy Manual Steps) covers assigning licences to users and provisioning on demand.

### Policy naming convention

```
<Region>-W365-<LicenceType>-<Policy>
```

| Example | Meaning |
|---|---|
| `Uksouth-W365-Enterprise-Policy` | Enterprise, UK South |
| `Uksouth-W365-Frontline-Dedicated-Policy` | Frontline Dedicated, UK South |
| `Uksouth-W365-Frontline-Shared-Policy` | Frontline Shared, UK South |
| `Europe-W365-Reserve-Policy` | Reserve, Europe geography |

The suffix (`Policy` by default) and group prefix (`SG-W365` by default) can be customised in the Advanced Options expander on the Review page.

### Group naming convention

```
<Prefix>-<LicenceInfix>-<Region>-<Role>
```

| Group | Example |
|---|---|
| Licensing | `SG-W365CloudPC_Cloud PC Enterprise 4vCPU/16GB/256GB` |
| User | `SG-W365-ENT-Uksouth-User` |
| Admin | `SG-W365-ENT-Uksouth-Admin` |
| Devices (dynamic) | `SG-W365CloudPC-Devices` |

`ENT` is used for Enterprise deployments and `FL` for Frontline. Reserve uses a distinct `SG-W365R-<geography>-User` / `-Admin` pair (e.g. `SG-W365R-Europe-User`) and creates no licensing or devices group.

---

## Cleanup Tool

`Destroy-Windows365-GUI.ps1` connects to your tenant, scans for Windows 365 objects, and lets you choose exactly what to remove before deleting anything.

### What it can remove

- Provisioning policies
- Cloud PC user settings
- Entra ID security groups

Each item is listed with a checkbox — nothing is deleted until you review the selection and confirm. Scan filters let you include or exclude each object type before scanning.

---

## Prerequisites

### Required permissions

Your account must have the following Microsoft Graph delegated scopes:

| Scope | Used for |
|---|---|
| `CloudPC.ReadWrite.All` | Provisioning policies, user settings, AI config |
| `Group.ReadWrite.All` | Creating and managing Entra ID security groups |
| `LicenseAssignment.ReadWrite.All` | Group-based licence assignment (Enterprise) |
| `DeviceManagementConfiguration.ReadWrite.All` | Windows Update for Business rings |

The wizard requests all scopes automatically at sign-in.

### Software requirements

- **Windows** — WPF is Windows-only; these scripts will not run on macOS or Linux
- **PowerShell 7.0 or higher**
  ```powershell
  winget install --id Microsoft.PowerShell --source winget
  ```
- **Microsoft.Graph.Authentication** — lightweight (~2 MB); auto-installed by the script if missing

### Tenant requirements

- Active Windows 365 subscription (Enterprise and/or Frontline)
- At least one available Cloud PC SKU
- At least one supported Windows 11 gallery image

---

## Running the scripts

```powershell
# Deployment wizard
.\Deploy-Windows365-GUI.ps1

# Cleanup tool
.\Destroy-Windows365-GUI.ps1
```

If prompted about execution policy:

```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

Both scripts detect if they are not running on an STA thread and relaunch automatically — you do not need to do anything special.

---

## Troubleshooting

### "No Cloud PC service plans found"
- Verify your Windows 365 subscription is active in this tenant
- Confirm you have `CloudPC.ReadWrite.All` scope
- Ensure at least one Cloud PC SKU is available

### "No supported region groups returned"
- Check that Windows 365 is available in your target region
- Verify Microsoft Graph permissions

### "Only unsupported images are available"
- No compatible gallery images are in your tenant; check the Microsoft 365 admin centre
- Wait for any recently uploaded images to finish processing

### Frontline assignment fails
- The Cloud PC backend can take longer than the Graph directory to replicate new policies
- The wizard retries the `/assign` call up to 6 times with a 10-second gap — wait for it to complete
- If it still fails, the results page will show the exact error and the manual steps to complete in the Intune portal

### Group replication delays
- The wizard waits 15 seconds after group creation and verifies each group is accessible before proceeding
- If groups are still not accessible after retries, check Entra ID health

---

## Idempotency

Both tools are safe to re-run:
- If a group or policy already exists it is reused, not duplicated
- Existing policy assignments are merged rather than overwritten

---

## Author

**Jon Jarvis** — Microsoft MVP (Intune & Windows 365)

- Website: [cloudendpoint.ai](https://www.cloudendpoint.ai)
- Email: [jon@jonjarvis.co.uk](mailto:jon@jonjarvis.co.uk)
- LinkedIn: [linkedin.com/in/jonjarvisuk](https://www.linkedin.com/in/jonjarvisuk/)

---

## License

MIT License — free to use, modify, and share within the Microsoft community.

Copyright (c) 2026 Jon Jarvis

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
