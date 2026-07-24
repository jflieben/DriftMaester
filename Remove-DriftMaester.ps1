<#
.SYNOPSIS
Removes DriftMaester resources, schedules, and permissions.

.DESCRIPTION
Reverses the core installer actions for DriftMaester. Supports either full resource-group removal
for installer-managed resource groups or scoped cleanup of DriftMaester assets inside a shared group.

.NOTES
Author: Jos Lieben / Lieben Consultancy
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string] $Subscription,
    [Parameter(Mandatory = $true)][string] $ResourceGroup,
    [Parameter(Mandatory = $false)][switch] $KeepData,
    [Parameter(Mandatory = $false)][switch] $Force
)

$ErrorActionPreference = 'Stop'

function Write-UninstallLog {
    param(
        [Parameter(Mandatory = $true)][string] $Message,
        [Parameter(Mandatory = $false)][ValidateSet('Info','Warning','Error','Success')][string] $Level = 'Info'
    )

    $prefix = "[{0:u}] [{1}]" -f (Get-Date), $Level.ToUpperInvariant()
    switch ($Level) {
        'Warning' { Write-Warning "$prefix $Message" }
        'Error' { Write-Error "$prefix $Message" -ErrorAction Continue }
        'Success' { Write-Host "$prefix $Message" -ForegroundColor Green }
        default { Write-Host "$prefix $Message" }
    }
}

function ConvertTo-SafeToken {
    param([string] $Value, [int] $MaxLength = 10)

    if ([string]::IsNullOrWhiteSpace($Value)) { return 'default' }
    $safe = ($Value.ToLowerInvariant() -replace '[^a-z0-9]', '')
    if ([string]::IsNullOrWhiteSpace($safe)) { $safe = 'default' }
    if ($safe.Length -gt $MaxLength) { $safe = $safe.Substring(0, $MaxLength) }
    return $safe
}

function New-DeterministicSuffix {
    param([Parameter(Mandatory = $true)][string] $Seed)

    $bytes = [Text.Encoding]::UTF8.GetBytes($Seed.ToLowerInvariant())
    $sha = [System.Security.Cryptography.SHA256]::Create()
    try {
        $hash = $sha.ComputeHash($bytes)
    } finally {
        $sha.Dispose()
    }

    return (($hash | Select-Object -First 5 | ForEach-Object { $_.ToString('x2') }) -join '')
}

function Get-DriftNames {
    param(
        [Parameter(Mandatory = $true)][string] $SelectedSubscriptionId,
        [Parameter(Mandatory = $true)][string] $SelectedResourceGroupName
    )

    $suffix = New-DeterministicSuffix -Seed "$SelectedSubscriptionId/$SelectedResourceGroupName"
    $safeGroup = ConvertTo-SafeToken -Value $SelectedResourceGroupName -MaxLength 10
    $storageRaw = "maester$safeGroup$suffix"
    $storageName = ($storageRaw -replace '[^a-z0-9]', '')
    if ($storageName.Length -gt 24) { $storageName = $storageName.Substring(0,24) }

    [PSCustomObject]@{
        AutomationAccount = 'driftmaester'
        StorageAccount = $storageName
        InvokeRunbook = 'Invoke-DriftMaester'
        UpdateRunbook = 'Update-DriftMaester'
        InvokeSchedule = 'driftmaester-invoke'
        UpdateSchedule = 'driftmaester-update'
    }
}

function Confirm-Action {
    param([Parameter(Mandatory = $true)][string] $Prompt)

    if ($Force) { return $true }
    $answer = Read-Host "$Prompt [y/N]"
    return $answer -match '^(y|yes)$'
}

if (-not (Get-AzContext -ErrorAction SilentlyContinue)) {
    Connect-AzAccount | Out-Null
}

$sub = Get-AzSubscription -SubscriptionId $Subscription -ErrorAction SilentlyContinue
if (-not $sub) {
    $sub = Get-AzSubscription -SubscriptionName $Subscription -ErrorAction SilentlyContinue | Select-Object -First 1
}
if (-not $sub) {
    throw "Could not resolve subscription '$Subscription'."
}

Set-AzContext -SubscriptionId $sub.Id -Tenant $sub.TenantId | Out-Null

$rg = Get-AzResourceGroup -Name $ResourceGroup -ErrorAction SilentlyContinue
if (-not $rg) {
    throw "Resource group '$ResourceGroup' was not found."
}

$names = Get-DriftNames -SelectedSubscriptionId $sub.Id -SelectedResourceGroupName $ResourceGroup
$isManagedRg = $false
if ($rg.Tags -and $rg.Tags.ContainsKey('DriftMaesterManaged') -and [string] $rg.Tags['DriftMaesterManaged'] -eq 'true') {
    $isManagedRg = $true
}

Write-UninstallLog "Resolved DriftMaester assets in resource group '$ResourceGroup'."
Write-UninstallLog "Detected automation account '$($names.AutomationAccount)' and storage account '$($names.StorageAccount)'."

$fullDelete = $isManagedRg -and (-not $KeepData)
if ($fullDelete) {
    if (-not (Confirm-Action -Prompt "Resource group '$ResourceGroup' is tagged as DriftMaester-managed. Remove entire resource group?")) {
        Write-UninstallLog 'Uninstall cancelled.' -Level Warning
        exit 0
    }

    Write-UninstallLog "Removing resource group '$ResourceGroup' (full uninstall)." -Level Warning
    Remove-AzResourceGroup -Name $ResourceGroup -Force -AsJob | Out-Null
    Write-UninstallLog 'Resource group deletion started. Uninstall complete.' -Level Success
    exit 0
}

if (-not (Confirm-Action -Prompt 'Proceed with scoped DriftMaester cleanup in this shared resource group?')) {
    Write-UninstallLog 'Uninstall cancelled.' -Level Warning
    exit 0
}

$automation = Get-AzAutomationAccount -ResourceGroupName $ResourceGroup -Name $names.AutomationAccount -ErrorAction SilentlyContinue
$managedIdentityObjectId = $null
$managedIdentityAppId = $null
if ($automation -and $automation.Identity -and $automation.Identity.PrincipalId) {
    $managedIdentityObjectId = [string] $automation.Identity.PrincipalId
    $sp = Get-AzADServicePrincipal -ObjectId $managedIdentityObjectId -ErrorAction SilentlyContinue
    if ($sp -and $sp.AppId) {
        $managedIdentityAppId = [string] $sp.AppId
    }
}

if ($automation) {
    try {
        $scheduledRunbooks = @(Get-AzAutomationScheduledRunbook -ResourceGroupName $ResourceGroup -AutomationAccountName $names.AutomationAccount -ErrorAction SilentlyContinue)
        foreach ($link in $scheduledRunbooks) {
            if ($link.RunbookName -in @($names.InvokeRunbook, $names.UpdateRunbook) -or $link.ScheduleName -in @($names.InvokeSchedule, $names.UpdateSchedule)) {
                Unregister-AzAutomationScheduledRunbook -ResourceGroupName $ResourceGroup -AutomationAccountName $names.AutomationAccount -RunbookName $link.RunbookName -ScheduleName $link.ScheduleName -Force -ErrorAction SilentlyContinue | Out-Null
            }
        }

        foreach ($scheduleName in @($names.InvokeSchedule, $names.UpdateSchedule)) {
            Remove-AzAutomationSchedule -ResourceGroupName $ResourceGroup -AutomationAccountName $names.AutomationAccount -Name $scheduleName -Force -ErrorAction SilentlyContinue | Out-Null
        }

        foreach ($runbookName in @($names.InvokeRunbook, $names.UpdateRunbook)) {
            Remove-AzAutomationRunbook -ResourceGroupName $ResourceGroup -AutomationAccountName $names.AutomationAccount -Name $runbookName -Force -ErrorAction SilentlyContinue | Out-Null
        }

        if (-not $KeepData) {
            Remove-AzAutomationAccount -ResourceGroupName $ResourceGroup -Name $names.AutomationAccount -Force -ErrorAction SilentlyContinue | Out-Null
            Write-UninstallLog "Removed automation account '$($names.AutomationAccount)'."
        } else {
            Write-UninstallLog "KeepData is enabled; automation account '$($names.AutomationAccount)' was left in place." -Level Warning
        }
    } catch {
        Write-UninstallLog "Automation cleanup encountered an issue: $($_.Exception.Message)" -Level Warning
    }
}

if ($managedIdentityObjectId) {
    $roleAssignments = @(Get-AzRoleAssignment -ObjectId $managedIdentityObjectId -ErrorAction SilentlyContinue)
    foreach ($assignment in $roleAssignments) {
        $roleName = [string] $assignment.RoleDefinitionName
        $scope = [string] $assignment.Scope
        if ($roleName -in @('Reader', 'Storage Blob Data Contributor', 'Automation Contributor', 'User Access Administrator')) {
            try {
                Remove-AzRoleAssignment -ObjectId $managedIdentityObjectId -RoleDefinitionName $roleName -Scope $scope -ErrorAction SilentlyContinue | Out-Null
            } catch {
                Write-UninstallLog "Could not remove role '$roleName' at '$scope': $($_.Exception.Message)" -Level Warning
            }
        }
    }
}

try {
    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
    Import-Module Microsoft.Graph.Applications -ErrorAction Stop
    Import-Module Microsoft.Graph.Identity.DirectoryManagement -ErrorAction Stop

    Connect-MgGraph -Scopes 'Application.Read.All','AppRoleAssignment.ReadWrite.All','RoleManagement.ReadWrite.Directory' -TenantId $sub.TenantId -NoWelcome | Out-Null

    if ($managedIdentityAppId) {
        $miSp = Get-MgServicePrincipal -Filter "appId eq '$managedIdentityAppId'" -All | Select-Object -First 1
        if ($miSp) {
            $assignments = @(Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $miSp.Id -All)
            foreach ($appAssignment in $assignments) {
                try {
                    Remove-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $miSp.Id -AppRoleAssignmentId $appAssignment.Id -ErrorAction SilentlyContinue
                } catch {
                    Write-UninstallLog "Could not remove app-role assignment '$($appAssignment.Id)': $($_.Exception.Message)" -Level Warning
                }
            }

            $roles = @(Get-MgDirectoryRole -All)
            foreach ($role in $roles) {
                $members = @(Get-MgDirectoryRoleMember -DirectoryRoleId $role.Id -All)
                if ($members | Where-Object { $_.Id -eq $miSp.Id }) {
                    try {
                        Remove-MgDirectoryRoleMemberByRef -DirectoryRoleId $role.Id -DirectoryObjectId $miSp.Id -ErrorAction SilentlyContinue
                    } catch {
                        Write-UninstallLog "Could not remove directory role member from '$($role.DisplayName)': $($_.Exception.Message)" -Level Warning
                    }
                }
            }
        }
    }
} catch {
    Write-UninstallLog "Graph cleanup skipped or partially failed: $($_.Exception.Message)" -Level Warning
}

if (-not $KeepData) {
    try {
        Remove-AzStorageAccount -ResourceGroupName $ResourceGroup -Name $names.StorageAccount -Force -ErrorAction SilentlyContinue | Out-Null
        Write-UninstallLog "Removed storage account '$($names.StorageAccount)'."
    } catch {
        Write-UninstallLog "Could not remove storage account '$($names.StorageAccount)': $($_.Exception.Message)" -Level Warning
    }
} else {
    Write-UninstallLog 'KeepData is enabled; storage account was preserved.'
}

Write-UninstallLog 'Scoped DriftMaester uninstall completed.' -Level Success
