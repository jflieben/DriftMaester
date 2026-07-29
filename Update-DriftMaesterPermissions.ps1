<#
.SYNOPSIS
Reconciles the Entra ID directory roles and API permissions of the DriftMaester managed identity.

.DESCRIPTION
A lightweight companion to Install-DriftMaester.ps1. It does NOT touch Azure resources, runbooks, schedules,
storage, Azure RBAC, or Exchange Online mail-send RBAC. It only:

- Resolves the DriftMaester Automation Account's system-assigned managed identity in the target resource group.
- Ensures the Microsoft Graph application permissions are assigned.
- Ensures the SharePoint Online application permissions are assigned.
- Ensures the Exchange Online 'Exchange.ManageAsApp' application permission is assigned.
- Ensures the Entra ID directory roles are assigned.

The required permissions and roles come from the shared DriftMaester.Permissions.ps1 definitions file, so this
script and the installer never drift apart. Existing grants are skipped, so the script is idempotent and safe to
re-run. Use it after adding new Maester tests that require extra permissions, without re-running the full
installer.

.PARAMETER Subscription
Subscription id or name that hosts the DriftMaester Automation Account.

.PARAMETER ResourceGroup
Resource group that contains the DriftMaester Automation Account.

.PARAMETER TenantId
Optional Entra ID tenant id to force for the Azure and Graph sign-ins.

.PARAMETER AzDevOps
Optional. When omitted, the 'Azure DevOps Administrator' directory role is
skipped (the DriftMaester DevOps tests are only relevant when a DevOps organization is configured).

.EXAMPLE
./Update-DriftMaesterPermissions.ps1 -Subscription "sub-id-or-name" -ResourceGroup "rg-driftmaester"

.EXAMPLE
iex ((Invoke-WebRequest -UseBasicParsing 'https://raw.githubusercontent.com/jflieben/DriftMaester/main/Update-DriftMaesterPermissions.ps1').Content)

.NOTES
Author: Jos Lieben / Lieben Consultancy
Website: https://www.lieben.nu
Blog: https://www.lieben.nu/liebensraum/
Free for non-commercial use. Commercial use requires a license:
https://www.lieben.nu/liebensraum/commercial-use/
#>

[CmdletBinding()]
param(
	[Parameter(Mandatory = $false)][string] $Subscription,
	[Parameter(Mandatory = $false)][string] $ResourceGroup,
	[Parameter(Mandatory = $false)][string] $TenantId,
	[switch] $AzDevOps
)

$ErrorActionPreference = 'Stop'
$script:AutomationAccountName = 'driftmaester'
$script:DefinitionsUrl = 'https://raw.githubusercontent.com/jflieben/DriftMaester/main/DriftMaester.Permissions.ps1'

function Write-UpdateLog {
	param(
		[Parameter(Mandatory = $true)][string] $Message,
		[Parameter(Mandatory = $false)][ValidateSet('Info', 'Warning', 'Error', 'Success')][string] $Level = 'Info'
	)

	$prefix = "[{0:u}] [{1}]" -f (Get-Date), $Level.ToUpperInvariant()
	switch ($Level) {
		'Warning' { Write-Warning "$prefix $Message" }
		'Error' { Write-Error "$prefix $Message" -ErrorAction Continue }
		'Success' { Write-Host "$prefix $Message" -ForegroundColor Green }
		default { Write-Host "$prefix $Message" }
	}
}

function Assert-RequiredModule {
	param([Parameter(Mandatory = $true)][string] $Name)

	if (-not (Get-Module -ListAvailable -Name $Name | Select-Object -First 1)) {
		throw "Required module '$Name' is not installed. Install it with: Install-Module $Name -Scope CurrentUser"
	}

	Import-Module $Name -ErrorAction Stop
}

function Get-DriftMaesterDefinitions {
	# Loads the shared permission/role definitions. Uses a local copy when the script runs from disk, and falls
	# back to fetching it from GitHub when the script is piped straight into the shell (iex).
	param(
		[Parameter(Mandatory = $false)][string] $DefinitionsUrl = $script:DefinitionsUrl
	)

	$localPath = if ($PSScriptRoot) { Join-Path -Path $PSScriptRoot -ChildPath 'DriftMaester.Permissions.ps1' } else { $null }
	if ($localPath -and (Test-Path -LiteralPath $localPath)) {
		$content = Get-Content -LiteralPath $localPath -Raw
	} else {
		$content = (Invoke-WebRequest -UseBasicParsing -Uri $DefinitionsUrl -ErrorAction Stop).Content
	}

	$definitions = & ([scriptblock]::Create($content))
	if (-not $definitions -or -not $definitions.GraphApplicationPermissions) {
		throw "Failed to load DriftMaester permission definitions from '$(if ($localPath -and (Test-Path -LiteralPath $localPath)) { $localPath } else { $DefinitionsUrl })'."
	}

	return $definitions
}

function Connect-ToAzureForUpdate {
	param([Parameter(Mandatory = $false)][string] $RequestedTenantId)

	Assert-RequiredModule -Name Az.Accounts

	$context = Get-AzContext -ErrorAction SilentlyContinue
	if (-not $context) {
		Write-UpdateLog 'Connecting to Azure. Use an account that can read the Automation Account and assign directory roles / API permissions.'
		$connectParams = @{ ErrorAction = 'Stop' }
		if (-not [string]::IsNullOrWhiteSpace($RequestedTenantId)) { $connectParams['Tenant'] = $RequestedTenantId }
		Connect-AzAccount @connectParams | Out-Null
	}
}

function Resolve-DriftSubscription {
	param(
		[Parameter(Mandatory = $false)][string] $SubscriptionInput,
		[Parameter(Mandatory = $false)][string] $RequestedTenantId
	)

	if (-not [string]::IsNullOrWhiteSpace($SubscriptionInput)) {
		$subscription = Get-AzSubscription -SubscriptionId $SubscriptionInput -ErrorAction SilentlyContinue
		if (-not $subscription) {
			$subscription = Get-AzSubscription -SubscriptionName $SubscriptionInput -ErrorAction SilentlyContinue | Select-Object -First 1
		}
		if (-not $subscription) {
			throw "Could not resolve subscription '$SubscriptionInput'."
		}
		return $subscription
	}

	$listParams = @{ ErrorAction = 'Stop' }
	if (-not [string]::IsNullOrWhiteSpace($RequestedTenantId)) { $listParams['TenantId'] = $RequestedTenantId }
	$subscriptions = @(Get-AzSubscription @listParams | Where-Object { $_.State -eq 'Enabled' } | Sort-Object Name)
	if ($subscriptions.Count -eq 0) {
		throw 'No enabled subscriptions were found for the signed-in account.'
	}
	if ($subscriptions.Count -eq 1) {
		return $subscriptions[0]
	}

	Write-Host 'Select the subscription that hosts DriftMaester:'
	for ($i = 0; $i -lt $subscriptions.Count; $i++) {
		Write-Host ("  [{0}] {1} ({2})" -f ($i + 1), $subscriptions[$i].Name, $subscriptions[$i].Id)
	}
	while ($true) {
		$choice = Read-Host 'Enter the number of the subscription'
		if ($choice -match '^\d+$' -and [int] $choice -ge 1 -and [int] $choice -le $subscriptions.Count) {
			return $subscriptions[[int] $choice - 1]
		}
		Write-Host 'Invalid selection. Try again.'
	}
}

function Get-DriftManagedIdentityObjectId {
	param(
		[Parameter(Mandatory = $true)][string] $ResourceGroupName,
		[Parameter(Mandatory = $true)][string] $AutomationAccountName
	)

	$automationAccount = Get-AzAutomationAccount -ResourceGroupName $ResourceGroupName -Name $AutomationAccountName -ErrorAction SilentlyContinue
	if (-not $automationAccount) {
		throw "Automation Account '$AutomationAccountName' was not found in resource group '$ResourceGroupName'. Run Install-DriftMaester.ps1 first."
	}

	$objectId = $null
	if ($automationAccount.Identity -and $automationAccount.Identity.PrincipalId) {
		$objectId = [string] $automationAccount.Identity.PrincipalId
	}

	if ([string]::IsNullOrWhiteSpace($objectId)) {
		$accountResource = Get-AzResource -ResourceGroupName $ResourceGroupName -ResourceType 'Microsoft.Automation/automationAccounts' -Name $AutomationAccountName -ErrorAction SilentlyContinue
		if ($accountResource -and $accountResource.Identity -and $accountResource.Identity.PrincipalId) {
			$objectId = [string] $accountResource.Identity.PrincipalId
		}
	}

	if ([string]::IsNullOrWhiteSpace($objectId)) {
		throw "Automation Account '$AutomationAccountName' has no system-assigned managed identity. Re-run Install-DriftMaester.ps1."
	}

	return $objectId
}

function Connect-ToGraphWithAzureToken {
	# Preferred path: reuse the Azure sign-in to mint a Graph token, so the operator signs in once and MSAL
	# version conflicts in the session are avoided (an -AccessToken login never touches MSAL).
	param([Parameter(Mandatory = $false)][string] $RequestedTenantId)

	try {
		$tokenParams = @{ ResourceUrl = 'https://graph.microsoft.com'; ErrorAction = 'Stop' }
		if ((Get-Command Get-AzAccessToken).Parameters.ContainsKey('AsSecureString')) {
			$tokenParams['AsSecureString'] = $true
		}
		if (-not [string]::IsNullOrWhiteSpace($RequestedTenantId)) { $tokenParams['TenantId'] = $RequestedTenantId }

		$token = Get-AzAccessToken @tokenParams
		$secureToken = if ($token.Token -is [System.Security.SecureString]) { $token.Token } else { ConvertTo-SecureString -String ([string] $token.Token) -AsPlainText -Force }

		Connect-MgGraph -AccessToken $secureToken -NoWelcome -ErrorAction Stop | Out-Null
		$null = Invoke-MgGraphRequest -Method GET -Uri 'https://graph.microsoft.com/v1.0/servicePrincipals?$top=1&$select=id' -ErrorAction Stop

		$graphContext = Get-MgContext -ErrorAction SilentlyContinue
		Write-UpdateLog "Reusing the Azure sign-in for Microsoft Graph (tenant $($graphContext.TenantId)). No second sign-in needed."
		return $true
	} catch {
		Write-UpdateLog "Could not reuse the Azure sign-in for Microsoft Graph: $($_.Exception.Message)" -Level Warning
		Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
		return $false
	}
}

function Connect-ToGraphForUpdate {
	param([Parameter(Mandatory = $false)][string] $RequestedTenantId)

	Assert-RequiredModule -Name Microsoft.Graph.Authentication
	$requiredScopes = @(
		'Application.Read.All',
		'AppRoleAssignment.ReadWrite.All',
		'Directory.ReadWrite.All',
		'RoleManagement.ReadWrite.Directory'
	)

	$azureTokenAttempted = $false
	while ($true) {
		$context = Get-MgContext -ErrorAction SilentlyContinue
		if ($context -and (-not $RequestedTenantId -or $context.TenantId -eq $RequestedTenantId)) {
			return
		}

		if ($context) {
			if ($RequestedTenantId -and $context.TenantId -ne $RequestedTenantId) {
				Write-UpdateLog "The current Graph tenant does not match the selected tenant '$RequestedTenantId'." -Level Warning
			} else {
				return
			}
			Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
		}

		if (-not $azureTokenAttempted) {
			$azureTokenAttempted = $true
			if (Connect-ToGraphWithAzureToken -RequestedTenantId $RequestedTenantId) {
				continue
			}
			Write-UpdateLog 'Falling back to an interactive Microsoft Graph sign-in.'
		}

		$params = @{ Scopes = $requiredScopes; NoWelcome = $true }
		if ($RequestedTenantId) { $params['TenantId'] = $RequestedTenantId }
		Write-UpdateLog 'Connecting to Microsoft Graph. Use a Global Administrator or Privileged Role Administrator account.'
		Connect-MgGraph @params -ErrorAction Stop | Out-Null
	}
}

function Invoke-GraphRequestAllPages {
	param([Parameter(Mandatory = $true)][string] $Uri)

	$items = [System.Collections.Generic.List[object]]::new()
	$next = $Uri
	while (-not [string]::IsNullOrWhiteSpace($next)) {
		$response = Invoke-MgGraphRequest -Method GET -Uri $next
		foreach ($item in @($response.value)) { $items.Add($item) }
		$next = if ($response.'@odata.nextLink') { [string] $response.'@odata.nextLink' } else { $null }
	}

	$items.ToArray()
}

function Get-ServicePrincipalByAppId {
	param([Parameter(Mandatory = $true)][string] $AppId)

	$filter = [Uri]::EscapeDataString("appId eq '$AppId'")
	$servicePrincipals = @(Invoke-GraphRequestAllPages -Uri "https://graph.microsoft.com/v1.0/servicePrincipals?`$filter=$filter&`$select=id,appId,displayName,appRoles")
	$servicePrincipal = $servicePrincipals | Select-Object -First 1
	if (-not $servicePrincipal) {
		throw "Could not find service principal with appId '$AppId'."
	}

	return $servicePrincipal
}

function Set-DriftAppRoleAssignment {
	param(
		[Parameter(Mandatory = $true)][object] $ManagedIdentityServicePrincipal,
		[Parameter(Mandatory = $true)][object] $ResourceServicePrincipal,
		[Parameter(Mandatory = $true)][string] $AppRoleValue
	)

	$appRole = @($ResourceServicePrincipal.appRoles | Where-Object { $_.value -eq $AppRoleValue -and $_.allowedMemberTypes -contains 'Application' } | Select-Object -First 1)
	if (-not $appRole) {
		Write-UpdateLog "App role '$AppRoleValue' was not found on '$($ResourceServicePrincipal.displayName)'. Skipping." -Level Warning
		return
	}

	$existingUri = "https://graph.microsoft.com/v1.0/servicePrincipals/$($ManagedIdentityServicePrincipal.id)/appRoleAssignments"
	$existingAssignments = @(Invoke-GraphRequestAllPages -Uri $existingUri)
	$matchingAssignment = $existingAssignments | Where-Object { $_.appRoleId -eq $appRole.id -and $_.resourceId -eq $ResourceServicePrincipal.id } | Select-Object -First 1
	if ($matchingAssignment) {
		Write-UpdateLog "API permission '$AppRoleValue' already assigned on '$($ResourceServicePrincipal.displayName)'."
		return
	}

	Write-UpdateLog "Assigning API permission '$AppRoleValue' on '$($ResourceServicePrincipal.displayName)'." -Level Success
	$body = @{
		principalId = $ManagedIdentityServicePrincipal.id
		resourceId  = $ResourceServicePrincipal.id
		appRoleId   = $appRole.id
	}
	Invoke-MgGraphRequest -Method POST -Uri $existingUri -Body ($body | ConvertTo-Json) -ContentType 'application/json' | Out-Null
}

function Set-DriftDirectoryRoleAssignment {
	param(
		[Parameter(Mandatory = $true)][string] $ManagedIdentityServicePrincipalId,
		[Parameter(Mandatory = $true)][string] $RoleDisplayName
	)

	$roleFilter = [Uri]::EscapeDataString("displayName eq '$RoleDisplayName'")
	$role = @(Invoke-GraphRequestAllPages -Uri "https://graph.microsoft.com/v1.0/directoryRoles?`$filter=$roleFilter") | Select-Object -First 1
	if (-not $role) {
		$template = @(Invoke-GraphRequestAllPages -Uri 'https://graph.microsoft.com/v1.0/directoryRoleTemplates') | Where-Object { $_.displayName -eq $RoleDisplayName } | Select-Object -First 1
		if (-not $template) {
			Write-UpdateLog "Directory role '$RoleDisplayName' was not found. Skipping." -Level Warning
			return
		}

		Write-UpdateLog "Activating directory role '$RoleDisplayName'."
		Invoke-MgGraphRequest -Method POST -Uri 'https://graph.microsoft.com/v1.0/directoryRoles' -Body (@{ roleTemplateId = $template.id } | ConvertTo-Json) -ContentType 'application/json' | Out-Null
		$role = @(Invoke-GraphRequestAllPages -Uri "https://graph.microsoft.com/v1.0/directoryRoles?`$filter=$roleFilter") | Select-Object -First 1
	}

	if (-not $role) {
		Write-UpdateLog "Directory role '$RoleDisplayName' could not be activated. Skipping." -Level Warning
		return
	}

	$membersUri = "https://graph.microsoft.com/v1.0/directoryRoles/$($role.id)/members?`$select=id"
	$members = @(Invoke-GraphRequestAllPages -Uri $membersUri)
	$matchingMember = $members | Where-Object { $_.id -eq $ManagedIdentityServicePrincipalId } | Select-Object -First 1
	if ($matchingMember) {
		Write-UpdateLog "Directory role '$RoleDisplayName' already assigned to the managed identity."
		return
	}

	Write-UpdateLog "Assigning directory role '$RoleDisplayName' to the managed identity." -Level Success
	$body = @{
		'@odata.id' = "https://graph.microsoft.com/v1.0/directoryObjects/$ManagedIdentityServicePrincipalId"
	}

	try {
		Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/v1.0/directoryRoles/$($role.id)/members/`$ref" -Body ($body | ConvertTo-Json) -ContentType 'application/json' -ErrorAction Stop | Out-Null
	} catch {
		$postError = $_
		$memberExistsAfterPost = $false
		try {
			$membersAfterPost = @(Invoke-GraphRequestAllPages -Uri $membersUri)
			$memberExistsAfterPost = [bool] ($membersAfterPost | Where-Object { $_.id -eq $ManagedIdentityServicePrincipalId } | Select-Object -First 1)
		} catch {
			$memberExistsAfterPost = $false
		}

		$errorText = @(
			$postError.Exception.Message
			$postError.ErrorDetails.Message
			($postError | Out-String)
		) -join [Environment]::NewLine

		if ($memberExistsAfterPost -or $errorText -match 'already exist|object references already exist') {
			Write-UpdateLog "Directory role '$RoleDisplayName' already assigned to the managed identity."
			return
		}

		throw $postError
	}
}

# ---------------------------------------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------------------------------------

Write-UpdateLog 'DriftMaester permission update started.'

$definitions = Get-DriftMaesterDefinitions

if ([string]::IsNullOrWhiteSpace($ResourceGroup)) {
	$ResourceGroup = Read-Host 'Enter the resource group that contains the DriftMaester Automation Account'
	if ([string]::IsNullOrWhiteSpace($ResourceGroup)) {
		throw 'A resource group is required.'
	}
}

Connect-ToAzureForUpdate -RequestedTenantId $TenantId
$selectedSubscription = Resolve-DriftSubscription -SubscriptionInput $Subscription -RequestedTenantId $TenantId
Set-AzContext -SubscriptionId $selectedSubscription.Id -Tenant $selectedSubscription.TenantId | Out-Null
Write-UpdateLog "Using subscription '$($selectedSubscription.Name)' ($($selectedSubscription.Id))."

$managedIdentityObjectId = Get-DriftManagedIdentityObjectId -ResourceGroupName $ResourceGroup -AutomationAccountName $script:AutomationAccountName
Write-UpdateLog "Resolved DriftMaester managed identity object id '$managedIdentityObjectId'."

Connect-ToGraphForUpdate -RequestedTenantId $TenantId

# The managed identity's service principal object id is the Automation Account's PrincipalId.
$managedIdentityServicePrincipal = [PSCustomObject]@{ id = $managedIdentityObjectId }

$graphServicePrincipal = Get-ServicePrincipalByAppId -AppId $definitions.GraphAppId
foreach ($permission in (@($definitions.GraphApplicationPermissions) | Sort-Object -Unique)) {
	Set-DriftAppRoleAssignment -ManagedIdentityServicePrincipal $managedIdentityServicePrincipal -ResourceServicePrincipal $graphServicePrincipal -AppRoleValue $permission
}

$exchangeServicePrincipal = Get-ServicePrincipalByAppId -AppId $definitions.ExchangeOnlineAppId
foreach ($permission in (@($definitions.ExchangeApplicationPermissions) | Sort-Object -Unique)) {
	Set-DriftAppRoleAssignment -ManagedIdentityServicePrincipal $managedIdentityServicePrincipal -ResourceServicePrincipal $exchangeServicePrincipal -AppRoleValue $permission
}

try {
	$sharePointServicePrincipal = Get-ServicePrincipalByAppId -AppId $definitions.SharePointOnlineAppId
	foreach ($permission in (@($definitions.SharePointApplicationPermissions) | Sort-Object -Unique)) {
		Set-DriftAppRoleAssignment -ManagedIdentityServicePrincipal $managedIdentityServicePrincipal -ResourceServicePrincipal $sharePointServicePrincipal -AppRoleValue $permission
	}
} catch {
	Write-UpdateLog "Could not assign SharePoint Online application permissions: $($_.Exception.Message). The Maester SharePoint Online tests will be skipped until this is resolved." -Level Warning
}

foreach ($roleName in @($definitions.DirectoryRoles)) {
	if ($roleName -eq 'Azure DevOps Administrator' -and !$AzDevOps) {
		Write-UpdateLog "Skipping directory role '$roleName' because AzDevOps was not provided (use -AzDevOps to include it)."
		continue
	}
	Set-DriftDirectoryRoleAssignment -ManagedIdentityServicePrincipalId $managedIdentityObjectId -RoleDisplayName $roleName
}

Write-UpdateLog 'DriftMaester permissions and directory roles reconciled.' -Level Success
Write-UpdateLog 'Note: this script does not configure Exchange Online mail-send RBAC or Azure RBAC. Re-run Install-DriftMaester.ps1 if those need to change.'
