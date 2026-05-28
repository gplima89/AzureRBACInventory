# Azure RBAC Inventory

PowerShell script that inventories **Azure RBAC role assignments granted to users** and exports a polished Excel workbook.

Service principals, managed identities, and group assignments are intentionally skipped — only `ObjectType = User` assignments are included.

## What you get

A single `.xlsx` file with these sheets:

| Sheet | Contents |
|---|---|
| **Role Assignments** | One row per assignment, sorted by user. Includes user (display name, UPN, objectId), role name + role GUID, scope type (Management Group / Subscription / Resource Group / Resource), scope name, and the full scope path. Privileged roles (`Owner`, `User Access Administrator`) and scope types are color-coded. |
| **User Summary** | Per-user totals: assignments, distinct roles, distinct scopes, and a flag for privileged roles. |
| **Role Summary** | Per-role totals: assignment count and distinct users. |
| **Role Definitions** | All Azure role definitions in tenant — name, GUID, custom/built-in, description, assignable scopes. Use this to map role names ↔ GUIDs. |
| **About** | Run metadata (who ran it, when, subscriptions scanned, totals). |

## Prerequisites

- PowerShell 7+ (recommended) or Windows PowerShell 5.1
- Modules (auto-installed on first run for the current user):
  - `Az.Accounts`
  - `Az.Resources`
  - `ImportExcel`
- An Azure sign-in with **Reader** (or higher) on the subscriptions / management groups you want to inventory.

## Usage

```powershell
# 1. Sign in (once per session)
Connect-AzAccount

# 2. Default: scan every enabled subscription you can see
./Get-AzureUserRBACInventory.ps1

# Specific subscriptions
./Get-AzureUserRBACInventory.ps1 -SubscriptionIds '00000000-0000-0000-0000-000000000000','11111111-1111-1111-1111-111111111111'

# Specific Management Groups
./Get-AzureUserRBACInventory.ps1 -ManagementGroupIds 'mg-finance','mg-platform'

# All Management Groups the signed-in user can see
./Get-AzureUserRBACInventory.ps1 -ScanAllManagementGroups

# Specific Resource Groups (format: '<subId>/<rgName>' or a full scope path)
./Get-AzureUserRBACInventory.ps1 -ResourceGroupScopes '00000000-0000-0000-0000-000000000000/rg-prod','00000000-0000-0000-0000-000000000000/rg-dev'

# Filter by role name (wildcards supported)
./Get-AzureUserRBACInventory.ps1 -RoleNames 'Owner','*Contributor*','User Access Administrator'

# Combine scopes + role filter + custom output
./Get-AzureUserRBACInventory.ps1 `
    -ManagementGroupIds 'mg-platform' `
    -SubscriptionIds '00000000-0000-0000-0000-000000000000' `
    -ResourceGroupScopes '00000000-0000-0000-0000-000000000000/rg-prod' `
    -RoleNames 'Owner','*Admin*' `
    -OutputPath C:\Reports\RBAC-2026-05.xlsx
```

### Parameters

| Parameter | Purpose |
|---|---|
| `-ManagementGroupIds` | One or more MG IDs (name, not full path) to query at the MG scope. |
| `-SubscriptionIds` | One or more subscription IDs to query. |
| `-ResourceGroupScopes` | One or more RG scopes. Format: `<subId>/<rgName>` or full `/subscriptions/.../resourceGroups/...` path. |
| `-ScanAllManagementGroups` | Enumerate every MG the signed-in account can see. |
| `-RoleNames` | Filter assignments by role name (case-insensitive, wildcards `*` supported). |
| `-OutputPath` | Excel output path. If omitted, defaults to `./AzureUserRBACInventory_<yyyyMMdd-HHmmss>.xlsx`. Pass a directory to drop a timestamped file inside it, or a full file path to control the name. |

> If no MG / Sub / RG parameter is supplied, the script defaults to scanning every enabled subscription the signed-in account can see.
>
> `Get-AzRoleAssignment` at a subscription scope already returns assignments inherited from parent MGs plus those at child RGs and resources. Use the explicit scope parameters when you want to **narrow** the scan (faster, less noise) rather than broaden it.

## Output location

Defaults to `./AzureUserRBACInventory.xlsx` in the current directory.
