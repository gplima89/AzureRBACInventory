<#
.SYNOPSIS
    Builds an Excel inventory of Azure RBAC role assignments granted to USERS.

.DESCRIPTION
    Collects all Azure role definitions (name + GUID) and all role assignments
    where the principal is a User (Service Principals, Managed Identities and
    Groups are skipped). Output is a styled multi-sheet Excel workbook sorted
    by user, including the resolved scope type (Management Group, Subscription,
    Resource Group, or Resource).

.PARAMETER ManagementGroupIds
    Optional list of Management Group IDs (the MG name/id, not the full path)
    to scan at the MG scope. Inherited assignments visible through child
    subscriptions are NOT duplicated.

.PARAMETER SubscriptionIds
    Optional list of subscription IDs to scan.

.PARAMETER ResourceGroupScopes
    Optional list of resource group scopes to scan. Accepts either:
      - 'subId/rgName'                 (e.g. '0000.../my-rg')
      - a full scope path beginning with '/subscriptions/.../resourceGroups/...'

.PARAMETER RoleNames
    Optional list of role names (case-insensitive) to include. Wildcards (*)
    are supported. If omitted, all roles are included.
    Examples: 'Owner', 'Contributor', '*Reader*'

.PARAMETER ScanAllManagementGroups
    Switch. Enumerate every Management Group the signed-in account can see
    and collect MG-scoped assignments. Useful when you don't know the MG IDs.

    Default behavior (no MG/Sub/RG parameter supplied): scan every enabled
    subscription the signed-in account can see.

.PARAMETER OutputPath
    Path of the Excel file to create.
      - If omitted, defaults to ./AzureUserRBACInventory_<yyyyMMdd-HHmmss>.xlsx
        in the current directory (timestamped so successive runs don't
        overwrite each other).
      - If a directory is supplied, a timestamped file is created inside it.
      - If a full file path is supplied, it is used as-is.

.EXAMPLE
    ./Get-AzureUserRBACInventory.ps1

.EXAMPLE
    ./Get-AzureUserRBACInventory.ps1 -SubscriptionIds 'xxxx-xxxx','yyyy-yyyy' -ScanAllManagementGroups

.EXAMPLE
    ./Get-AzureUserRBACInventory.ps1 -ResourceGroupScopes 'xxxx-xxxx/rg-prod','xxxx-xxxx/rg-dev' -RoleNames 'Owner','*Contributor*'

.EXAMPLE
    ./Get-AzureUserRBACInventory.ps1 -ManagementGroupIds 'mg-finance','mg-platform'

.EXAMPLE
    # Drop a timestamped file into C:\Reports
    ./Get-AzureUserRBACInventory.ps1 -OutputPath C:\Reports

.EXAMPLE
    # Use an explicit file name
    ./Get-AzureUserRBACInventory.ps1 -OutputPath C:\Reports\RBAC-2026-05.xlsx

.NOTES
    Requires modules: Az.Accounts, Az.Resources, ImportExcel.
    Run Connect-AzAccount first.
#>

[CmdletBinding()]
param(
    [string[]] $ManagementGroupIds,
    [string[]] $SubscriptionIds,
    [string[]] $ResourceGroupScopes,
    [string[]] $RoleNames,
    [switch]   $ScanAllManagementGroups,
    [string]   $OutputPath
)

# Default output path includes a timestamp so successive runs don't overwrite.
if (-not $OutputPath) {
    $timestamp  = Get-Date -Format 'yyyyMMdd-HHmmss'
    $OutputPath = Join-Path (Get-Location) "AzureUserRBACInventory_$timestamp.xlsx"
} elseif (Test-Path $OutputPath -PathType Container) {
    # If a directory was passed, drop a timestamped file into it.
    $timestamp  = Get-Date -Format 'yyyyMMdd-HHmmss'
    $OutputPath = Join-Path $OutputPath "AzureUserRBACInventory_$timestamp.xlsx"
}

# --- Module checks -----------------------------------------------------------
function Ensure-Module {
    param([string] $Name)
    if (-not (Get-Module -ListAvailable -Name $Name)) {
        Write-Host "Installing module '$Name'..." -ForegroundColor Yellow
        Install-Module $Name -Scope CurrentUser -Force -AllowClobber -Repository PSGallery
    }
    Import-Module $Name -ErrorAction Stop
}

Ensure-Module Az.Accounts
Ensure-Module Az.Resources
Ensure-Module ImportExcel

# --- Auth check --------------------------------------------------------------
# Microsoft Graph access is required so Get-AzRoleAssignment / Get-AzRoleDefinition
# can resolve user display names, UPNs, and role metadata. Conditional Access /
# MFA policies often require an explicit -AuthScope on first sign-in.
$ctx = Get-AzContext
$needGraphReauth = $false

function Test-GraphAccess {
    try {
        Get-AzRoleDefinition -Name 'Reader' -ErrorAction Stop | Out-Null
        return $true
    } catch {
        if ($_.Exception.Message -match 'MicrosoftGraphEndpointResourceId' -or
            $_.Exception.Message -match 'Authentication failed' -or
            $_.Exception.Message -match 'expired') {
            return $false
        }
        throw
    }
}

if (-not $ctx) {
    Write-Host 'No Azure context found. Launching Connect-AzAccount (with Graph scope)...' -ForegroundColor Yellow
    Connect-AzAccount -AuthScope MicrosoftGraphEndpointResourceId | Out-Null
    $ctx = Get-AzContext
} else {
    if (-not (Test-GraphAccess)) {
        Write-Host 'Microsoft Graph token missing or expired. Re-authenticating with Graph scope...' -ForegroundColor Yellow
        Connect-AzAccount -AuthScope MicrosoftGraphEndpointResourceId -TenantId $ctx.Tenant.Id | Out-Null
        $ctx = Get-AzContext
    }
}
Write-Host "Signed in as: $($ctx.Account.Id)" -ForegroundColor Cyan

# --- Scope parsing -----------------------------------------------------------
function Resolve-Scope {
    param([string] $Scope)

    if ([string]::IsNullOrWhiteSpace($Scope) -or $Scope -eq '/') {
        return [pscustomobject]@{
            ScopeType   = 'Root (Tenant)'
            ScopeName   = '/'
            MGOrSub     = ''
            ResourceGroup = ''
            ResourceType  = ''
            ResourceName  = ''
        }
    }

    # Management Group
    if ($Scope -match '^/providers/Microsoft\.Management/managementGroups/([^/]+)$') {
        return [pscustomobject]@{
            ScopeType     = 'Management Group'
            ScopeName     = $Matches[1]
            MGOrSub       = $Matches[1]
            ResourceGroup = ''
            ResourceType  = ''
            ResourceName  = ''
        }
    }

    # Subscription
    if ($Scope -match '^/subscriptions/([^/]+)$') {
        return [pscustomobject]@{
            ScopeType     = 'Subscription'
            ScopeName     = $Matches[1]
            MGOrSub       = $Matches[1]
            ResourceGroup = ''
            ResourceType  = ''
            ResourceName  = ''
        }
    }

    # Resource Group
    if ($Scope -match '^/subscriptions/([^/]+)/resourceGroups/([^/]+)$') {
        return [pscustomobject]@{
            ScopeType     = 'Resource Group'
            ScopeName     = $Matches[2]
            MGOrSub       = $Matches[1]
            ResourceGroup = $Matches[2]
            ResourceType  = ''
            ResourceName  = ''
        }
    }

    # Resource (anything deeper)
    if ($Scope -match '^/subscriptions/([^/]+)/resourceGroups/([^/]+)/providers/(.+)$') {
        $sub  = $Matches[1]
        $rg   = $Matches[2]
        $rest = $Matches[3]
        $parts = $rest -split '/'
        # parts looks like: <RP>, <type>, <name>[, <subtype>, <subname>, ...]
        $resourceType = if ($parts.Count -ge 2) { "$($parts[0])/$($parts[1])" } else { $parts[0] }
        $resourceName = ($parts | Select-Object -Skip 2) -join '/'
        return [pscustomobject]@{
            ScopeType     = 'Resource'
            ScopeName     = $resourceName
            MGOrSub       = $sub
            ResourceGroup = $rg
            ResourceType  = $resourceType
            ResourceName  = $resourceName
        }
    }

    return [pscustomobject]@{
        ScopeType     = 'Unknown'
        ScopeName     = $Scope
        MGOrSub       = ''
        ResourceGroup = ''
        ResourceType  = ''
        ResourceName  = ''
    }
}

# --- Collect role definitions ------------------------------------------------
Write-Host 'Collecting role definitions...' -ForegroundColor Cyan
$roleDefs = Get-AzRoleDefinition | Sort-Object Name
$roleDefRows = $roleDefs | ForEach-Object {
    [pscustomobject]@{
        RoleName       = $_.Name
        RoleId         = $_.Id
        IsCustom       = $_.IsCustom
        Description    = $_.Description
        AssignableScopes = ($_.AssignableScopes -join '; ')
    }
}
Write-Host ("  Found {0} role definitions." -f $roleDefRows.Count) -ForegroundColor Green

# --- Resolve scope set to scan ----------------------------------------------
$explicitScopeProvided = ($ManagementGroupIds -and $ManagementGroupIds.Count -gt 0) -or
                         ($SubscriptionIds   -and $SubscriptionIds.Count   -gt 0) -or
                         ($ResourceGroupScopes -and $ResourceGroupScopes.Count -gt 0) -or
                         $ScanAllManagementGroups

if (-not $explicitScopeProvided) {
    Write-Host 'No scope parameters specified — defaulting to all enabled subscriptions.' -ForegroundColor Yellow
    $SubscriptionIds = (Get-AzSubscription | Where-Object { $_.State -eq 'Enabled' }).Id
}

# Build list of (Label, Scope) tuples to query Get-AzRoleAssignment with.
$scopeQueries = New-Object System.Collections.Generic.List[object]

foreach ($mgId in ($ManagementGroupIds | Where-Object { $_ })) {
    $scopeQueries.Add([pscustomobject]@{
        Label = "Management Group '$mgId'"
        Scope = "/providers/Microsoft.Management/managementGroups/$mgId"
        SetContextSub = $null
    }) | Out-Null
}

if ($ScanAllManagementGroups) {
    try {
        $mgs = Get-AzManagementGroup -ErrorAction Stop
        foreach ($mg in $mgs) {
            if ($ManagementGroupIds -and ($ManagementGroupIds -contains $mg.Name)) { continue }
            $scopeQueries.Add([pscustomobject]@{
                Label = "Management Group '$($mg.Name)'"
                Scope = "/providers/Microsoft.Management/managementGroups/$($mg.Name)"
                SetContextSub = $null
            }) | Out-Null
        }
    } catch {
        Write-Warning "Could not enumerate management groups: $($_.Exception.Message)"
    }
}

foreach ($subId in ($SubscriptionIds | Where-Object { $_ })) {
    $scopeQueries.Add([pscustomobject]@{
        Label = "Subscription '$subId'"
        Scope = "/subscriptions/$subId"
        SetContextSub = $subId
    }) | Out-Null
}

foreach ($rg in ($ResourceGroupScopes | Where-Object { $_ })) {
    if ($rg.StartsWith('/')) {
        $scopePath = $rg.TrimEnd('/')
    } elseif ($rg -match '^([^/]+)/([^/]+)$') {
        $scopePath = "/subscriptions/$($Matches[1])/resourceGroups/$($Matches[2])"
    } else {
        Write-Warning "Skipping resource group scope '$rg' — expected 'subId/rgName' or a full '/subscriptions/.../resourceGroups/...' path."
        continue
    }
    $subForContext = $null
    if ($scopePath -match '^/subscriptions/([^/]+)/resourceGroups/[^/]+$') {
        $subForContext = $Matches[1]
    }
    $scopeQueries.Add([pscustomobject]@{
        Label = "Resource Group '$scopePath'"
        Scope = $scopePath
        SetContextSub = $subForContext
    }) | Out-Null
}

# --- Collect role assignments -----------------------------------------------
$allAssignments = New-Object System.Collections.Generic.List[object]
$seen = New-Object 'System.Collections.Generic.HashSet[string]'

foreach ($q in $scopeQueries) {
    try {
        if ($q.SetContextSub) {
            Set-AzContext -SubscriptionId $q.SetContextSub -ErrorAction Stop | Out-Null
        }
        Write-Host ("Scanning {0}" -f $q.Label) -ForegroundColor Cyan
        $assignments = Get-AzRoleAssignment -Scope $q.Scope -ErrorAction Stop
        foreach ($a in $assignments) {
            if ($a.ObjectType -ne 'User') { continue }
            $key = "$($a.RoleAssignmentId)"
            if ([string]::IsNullOrEmpty($key)) {
                $key = "$($a.Scope)|$($a.RoleDefinitionId)|$($a.ObjectId)"
            }
            if ($seen.Add($key)) {
                $allAssignments.Add($a) | Out-Null
            }
        }
    } catch {
        Write-Warning ("Failed to scan {0}: {1}" -f $q.Label, $_.Exception.Message)
    }
}

# --- Role name filter (supports wildcards) ----------------------------------
if ($RoleNames -and $RoleNames.Count -gt 0) {
    $patterns = @($RoleNames)
    $before = $allAssignments.Count
    $filtered = $allAssignments | Where-Object {
        $rn = $_.RoleDefinitionName
        foreach ($p in $patterns) { if ($rn -like $p) { return $true } }
        return $false
    }
    $allAssignments = [System.Collections.Generic.List[object]]::new()
    foreach ($a in $filtered) { $allAssignments.Add($a) | Out-Null }
    Write-Host ("Role name filter: kept {0} of {1} assignments." -f $allAssignments.Count, $before) -ForegroundColor Yellow
}

Write-Host ("Collected {0} unique user role assignments." -f $allAssignments.Count) -ForegroundColor Green

# --- Shape rows --------------------------------------------------------------
$rows = foreach ($a in $allAssignments) {
    $s = Resolve-Scope -Scope $a.Scope
    [pscustomobject]@{
        UserDisplayName   = $a.DisplayName
        UserPrincipalName = $a.SignInName
        UserObjectId      = $a.ObjectId
        RoleName          = $a.RoleDefinitionName
        RoleId            = $a.RoleDefinitionId
        ScopeType         = $s.ScopeType
        ScopeName         = $s.ScopeName
        SubscriptionOrMG  = $s.MGOrSub
        ResourceGroup     = $s.ResourceGroup
        ResourceType      = $s.ResourceType
        FullScope         = $a.Scope
        RoleAssignmentId  = $a.RoleAssignmentId
    }
}

$rows = @($rows | Sort-Object UserDisplayName, ScopeType, RoleName)

if ($rows.Count -eq 0) {
    Write-Warning 'No user role assignments were found in the scanned scopes.'
    # Add a placeholder so the workbook still has the sheet structure.
    $rows = @([pscustomobject]@{
        UserDisplayName   = '(no user assignments found)'
        UserPrincipalName = ''
        UserObjectId      = ''
        RoleName          = ''
        RoleId            = ''
        ScopeType         = ''
        ScopeName         = ''
        SubscriptionOrMG  = ''
        ResourceGroup     = ''
        ResourceType      = ''
        FullScope         = ''
        RoleAssignmentId  = ''
    })
}

# --- Summary -----------------------------------------------------------------
$userSummary = $rows | Group-Object UserDisplayName, UserPrincipalName | ForEach-Object {
    $parts = $_.Name -split ', '
    [pscustomobject]@{
        UserDisplayName     = $parts[0]
        UserPrincipalName   = $parts[1]
        TotalAssignments    = $_.Count
        DistinctRoles       = ($_.Group.RoleName | Sort-Object -Unique).Count
        DistinctScopes      = ($_.Group.FullScope | Sort-Object -Unique).Count
        HasOwnerOrUserAccessAdmin = (($_.Group.RoleName | Where-Object { $_ -in 'Owner','User Access Administrator' }).Count -gt 0)
    }
} | Sort-Object -Property @{Expression='TotalAssignments';Descending=$true}, UserDisplayName

$roleSummary = $rows | Group-Object RoleName | ForEach-Object {
    [pscustomobject]@{
        RoleName        = $_.Name
        AssignmentCount = $_.Count
        DistinctUsers   = ($_.Group.UserObjectId | Sort-Object -Unique).Count
    }
} | Sort-Object -Property @{Expression='AssignmentCount';Descending=$true}, RoleName

# --- Export to Excel ---------------------------------------------------------
if (Test-Path $OutputPath) { Remove-Item $OutputPath -Force }

Write-Host "Writing Excel: $OutputPath" -ForegroundColor Cyan

# Sheet 1: Role Assignments (main)
$xl = $rows | Export-Excel -Path $OutputPath -WorksheetName 'Role Assignments' `
    -AutoSize -FreezeTopRow -BoldTopRow -AutoFilter `
    -TableName 'tblRoleAssignments' -TableStyle 'Medium2' -PassThru

$ws = $xl.Workbook.Worksheets['Role Assignments']

$colNames = @($rows[0].PSObject.Properties.Name)

# Color-code ScopeType column
$scopeColIndex = [array]::IndexOf($colNames, 'ScopeType') + 1
if ($scopeColIndex -gt 0 -and $rows.Count -gt 0) {
    $lastRow = $rows.Count + 1
    $colLetter = [OfficeOpenXml.ExcelCellAddress]::GetColumnLetter($scopeColIndex)
    $range = "$colLetter`2:$colLetter$lastRow"
    Add-ConditionalFormatting -Worksheet $ws -Address $range -RuleType Equal -ConditionValue '"Management Group"' -BackgroundColor ([System.Drawing.Color]::FromArgb(214,234,248)) -ForegroundColor ([System.Drawing.Color]::FromArgb(21,67,96))
    Add-ConditionalFormatting -Worksheet $ws -Address $range -RuleType Equal -ConditionValue '"Subscription"'     -BackgroundColor ([System.Drawing.Color]::FromArgb(212,239,223)) -ForegroundColor ([System.Drawing.Color]::FromArgb(20,90,50))
    Add-ConditionalFormatting -Worksheet $ws -Address $range -RuleType Equal -ConditionValue '"Resource Group"'   -BackgroundColor ([System.Drawing.Color]::FromArgb(253,235,208)) -ForegroundColor ([System.Drawing.Color]::FromArgb(125,79,16))
    Add-ConditionalFormatting -Worksheet $ws -Address $range -RuleType Equal -ConditionValue '"Resource"'         -BackgroundColor ([System.Drawing.Color]::FromArgb(245,215,215)) -ForegroundColor ([System.Drawing.Color]::FromArgb(120,40,40))
}

# Highlight privileged roles
$roleColIndex = [array]::IndexOf($colNames, 'RoleName') + 1
if ($roleColIndex -gt 0 -and $rows.Count -gt 0) {
    $lastRow = $rows.Count + 1
    $colLetter = [OfficeOpenXml.ExcelCellAddress]::GetColumnLetter($roleColIndex)
    $range = "$colLetter`2:$colLetter$lastRow"
    Add-ConditionalFormatting -Worksheet $ws -Address $range -RuleType ContainsText -ConditionValue 'Owner' -BackgroundColor ([System.Drawing.Color]::FromArgb(248,215,218)) -ForegroundColor ([System.Drawing.Color]::FromArgb(120,30,40)) -Bold
    Add-ConditionalFormatting -Worksheet $ws -Address $range -RuleType ContainsText -ConditionValue 'User Access Administrator' -BackgroundColor ([System.Drawing.Color]::FromArgb(248,215,218)) -ForegroundColor ([System.Drawing.Color]::FromArgb(120,30,40)) -Bold
}

Close-ExcelPackage $xl

# Sheet 2: User Summary
$userSummary | Export-Excel -Path $OutputPath -WorksheetName 'User Summary' `
    -AutoSize -FreezeTopRow -BoldTopRow -AutoFilter `
    -TableName 'tblUserSummary' -TableStyle 'Medium6'

# Sheet 3: Role Summary
$roleSummary | Export-Excel -Path $OutputPath -WorksheetName 'Role Summary' `
    -AutoSize -FreezeTopRow -BoldTopRow -AutoFilter `
    -TableName 'tblRoleSummary' -TableStyle 'Medium4'

# Sheet 4: Role Definitions
$roleDefRows | Export-Excel -Path $OutputPath -WorksheetName 'Role Definitions' `
    -AutoSize -FreezeTopRow -BoldTopRow -AutoFilter `
    -TableName 'tblRoleDefinitions' -TableStyle 'Medium9'

# Sheet 5: About
$about = [pscustomobject]@{
    GeneratedOn         = (Get-Date).ToString('u')
    GeneratedBy         = $ctx.Account.Id
    Tenant              = $ctx.Tenant.Id
    ManagementGroups    = (($ManagementGroupIds  | Where-Object { $_ }) -join ', ')
    Subscriptions       = (($SubscriptionIds    | Where-Object { $_ }) -join ', ')
    ResourceGroupScopes = (($ResourceGroupScopes | Where-Object { $_ }) -join ', ')
    RoleNameFilter      = (($RoleNames           | Where-Object { $_ }) -join ', ')
    ScopesQueried       = $scopeQueries.Count
    AssignmentsTotal    = $rows.Count
    UsersTotal          = ($rows.UserObjectId | Sort-Object -Unique).Count
    RolesUsed           = ($rows.RoleName | Sort-Object -Unique).Count
}
$about | Export-Excel -Path $OutputPath -WorksheetName 'About' -AutoSize -BoldTopRow

# Reorder sheets: 'Role Assignments' first, 'About' last.
# MoveToStart moves the named sheet to index 0, so iterate the desired order
# in REVERSE — the first item in $order ends up at the front.
$pkg = Open-ExcelPackage -Path $OutputPath
$order = @('Role Assignments','User Summary','Role Summary','Role Definitions','About')
for ($i = $order.Count - 1; $i -ge 0; $i--) {
    $sheet = $pkg.Workbook.Worksheets[$order[$i]]
    if ($sheet) { $pkg.Workbook.Worksheets.MoveToStart($sheet.Name) }
}
Close-ExcelPackage $pkg

Write-Host "Done. Output: $OutputPath" -ForegroundColor Green
