<#
.SYNOPSIS
Installs or updates DriftMaester in Azure Automation.

.DESCRIPTION
Creates the resource group, Automation Account, PowerShell 7.6 runtime environment, storage account, runbooks,
schedules, Azure RBAC assignments, and managed identity API permissions needed by the DriftMaester runbooks.

The script is idempotent. Re-running it updates runbook content, recreates deterministic schedules/job schedules,
keeps existing Azure resources, and skips API permissions and role assignments that already exist.

.PARAMETER AlwaysSendReport
When true, configures the scheduled Invoke-DriftMaester runbook to send a report after every run.
When false, reports are only sent on the first run or when drift is detected.

.PARAMETER IncludeCopilotAndDataverse
When true, configures the scheduled Invoke-DriftMaester runbook to include Power Platform,
Copilot, Dynamics, and Dataverse-backed Maester tests.
When false, those checks are skipped and Dataverse connection warnings are suppressed.

.PARAMETER IncludeMaesterReport
When true, configures the scheduled Invoke-DriftMaester runbook to attach the most recent original
Maester HTML report (zipped) to the report email, in addition to the drift report.
When false (default), only the drift report is attached. In large tenants the Maester report can be
big, so the attachment may be rejected by the recipient mail system.

.PARAMETER TimeZone
Time zone id for the Azure Automation schedules. Defaults to the local system time zone.

.EXAMPLE
./Install-DriftMaester.ps1

.NOTES
Author: Jos Lieben / Lieben Consultancy
Website: https://www.lieben.nu
Blog: https://www.lieben.nu/liebensraum/
Free for non-commercial use. Commercial use requires a license:
https://www.lieben.nu/liebensraum/commercial-use/
#>

[CmdletBinding()]
param(
	[Parameter(Mandatory = $false)][switch] $GuiMode,
	[Parameter(Mandatory = $false)][string] $Subscription,
	[Parameter(Mandatory = $false)][string] $ResourceGroup,
	[Parameter(Mandatory = $false)][string] $Location,
	[Parameter(Mandatory = $false)][ArgumentCompleter({ 'daily', 'weekly', 'monthly' })][string] $Frequency,
	[Parameter(Mandatory = $false)][string] $TimeOfDay,
	[Parameter(Mandatory = $false)][ArgumentCompleter({ 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday' })][string] $DayOfWeek,
	[Parameter(Mandatory = $false)][ValidateRange(-1, 31)][int] $DayOfMonth,
	[Parameter(Mandatory = $false)][string] $Recipients,
	[Parameter(Mandatory = $false)][string] $SenderUserId,
	[Parameter(Mandatory = $false)][string] $DevOpsOrg,
	[Parameter(Mandatory = $false)][string] $TenantId,
	[Parameter(Mandatory = $false)][string] $MailSubject,
	[Parameter(Mandatory = $false)][string] $TimeZone,
	[Parameter(Mandatory = $false)][ValidateSet('auto', 'attach', 'link')][string] $ReportDelivery = 'auto',
	[Parameter(Mandatory = $false)][ValidateSet('none', 'low', 'medium', 'high', 'critical')][string] $AlertMinimumSeverity = 'none',
	[Parameter(Mandatory = $false)][string] $AlertRecipients,
	[Parameter(Mandatory = $false)][string] $TeamsWebhookUrl,
	[Parameter(Mandatory = $false)][ValidateRange(30, 3650)][int] $RetentionDays = 180,
	[Parameter(Mandatory = $false)][switch] $RunNow,
	[Parameter(Mandatory = $false)][bool] $AlwaysSendReport = $false,
	[Parameter(Mandatory = $false)][bool] $IncludeCopilotAndDataverse = $false,
	[Parameter(Mandatory = $false)][bool] $IncludeMaesterReport = $false
)

$ErrorActionPreference = 'Stop'
$script:AutomationApiVersion = '2024-10-23'
$script:Location = 'westeurope'
$script:PreferredRuntimeName = 'driftmaester'
$script:RuntimeVersion = '7.6'
$script:InvokeRunbookName = 'Invoke-DriftMaester'
$script:UpdateRunbookName = 'Update-DriftMaester'
$script:InvokeScheduleName = 'driftmaester-invoke'
$script:UpdateScheduleName = 'driftmaester-update'
$script:DriftMaesterVersion = '1.3.0'
$script:GithubRawBase = 'https://raw.githubusercontent.com/jflieben/DriftMaester/main/Runbooks'
$script:GraphAppId = '00000003-0000-0000-c000-000000000000'
$script:ExchangeOnlineAppId = '00000002-0000-0ff1-ce00-000000000000'
$script:SharePointOnlineAppId = '00000003-0000-0ff1-ce00-000000000000'
$script:DefaultRuntimePackages = @{
	'az'          = '15.1.0'
	'azure cli' = '2.77.0'
}

$RequiredGraphApplicationPermissions = @(
	'Policy.Read.ConditionalAccess',
    'DeviceManagementManagedDevices.Read.All',
    'UserAuthenticationMethod.Read.All',
    'OnPremDirectorySynchronization.Read.All',
    'SharePointTenantSettings.Read.All',
    'ReportSettings.ReadWrite.All',
	'ReportSettings.Read.All',
    'PrivilegedAccess.Read.AzureAD',
    'OrgSettings-Forms.Read.All',
    'DeviceManagementServiceConfig.Read.All',
    'SecurityIdentitiesHealth.Read.All',
    'DirectoryRecommendations.Read.All',
    'Directory.Read.All',
    'RoleManagement.Read.All',
    'DeviceManagementRBAC.Read.All',
    'SecurityIdentitiesSensors.Read.All',
    'DeviceManagementConfiguration.Read.All',
    'OrgSettings-AppsAndServices.Read.All',
    'IdentityRiskEvent.Read.All',
    'Policy.Read.All',
    'Reports.Read.All',
    'ThreatHunting.Read.All',
    'AuditLog.Read.All',
	'EntitlementManagement.Read.All',
	'NetworkAccess.Read.All',
	'RoleManagementAlert.Read.Directory'
)

# Granted on the 'Office 365 SharePoint Online' service principal (not Graph). PnP.PowerShell connects to the
# SharePoint tenant admin endpoint app-only with the managed identity, and Get-PnPTenant (used by the Maester
# SharePoint Online tests) is only authorised by Sites.FullControl.All. Lower roles return 401 on that endpoint.
$RequiredSharePointApplicationPermissions = @(
	'Sites.FullControl.All'
)

$DirectoryRolesForManagedIdentity = @(
	'Global Reader',
	'Teams Reader',
    'Azure DevOps Administrator'
)

function Write-InstallLog {
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
		throw "Required module '$Name' is not installed."
	}

	Import-Module $Name -ErrorAction Stop
}

function Get-MsalConflictHint {
	# Only called after an interactive Graph sign-in has already failed, to turn a .NET type load error into
	# something the operator can act on. Microsoft.Graph.Authentication binds to the Microsoft.Identity.Client
	# (MSAL) build it ships with, .NET keeps a single version of an assembly per process, and it cannot be
	# unloaded. So when another module got its own older MSAL in first, Connect-MgGraph fails before any browser
	# opens. This is only a hint: an assembly loaded into a module's private AssemblyLoadContext also shows up
	# here, and those copies are harmless, which is why this must never gate the install on its own.
	$loadedMsal = @([System.AppDomain]::CurrentDomain.GetAssemblies() | Where-Object { $_.GetName().Name -eq 'Microsoft.Identity.Client' }) | Select-Object -First 1
	if (-not $loadedMsal) {
		return ''
	}

	$graphModule = Get-Module -Name Microsoft.Graph.Authentication | Select-Object -First 1
	if (-not $graphModule) {
		$graphModule = Get-Module -ListAvailable -Name Microsoft.Graph.Authentication | Sort-Object Version -Descending | Select-Object -First 1
	}
	if (-not $graphModule) {
		return ''
	}

	$shippedMsal = Get-ChildItem -Path $graphModule.ModuleBase -Recurse -Filter 'Microsoft.Identity.Client.dll' -ErrorAction SilentlyContinue | Select-Object -First 1
	if (-not $shippedMsal) {
		return ''
	}

	$requiredVersion = [System.Reflection.AssemblyName]::GetAssemblyName($shippedMsal.FullName).Version
	$loadedVersion = $loadedMsal.GetName().Version
	if ($loadedVersion -ge $requiredVersion) {
		return ''
	}

	$origin = if ([string]::IsNullOrWhiteSpace($loadedMsal.Location)) { 'a module that was already imported' } else { $loadedMsal.Location }
	return @"

This session loaded Microsoft.Identity.Client $loadedVersion while Microsoft.Graph.Authentication $($graphModule.Version) expects $requiredVersion, which is the usual cause of this failure.
Loaded from: $origin
.NET cannot load two versions of the same assembly, so no sign-in prompt can appear. Close this window, open a new PowerShell session, and run the installer before importing other modules (PnP.PowerShell, MicrosoftTeams and ExchangeOnlineManagement each ship their own copy).
"@
}

function Read-RequiredValue {
	param(
		[Parameter(Mandatory = $true)][string] $Prompt,
		[Parameter(Mandatory = $false)][string] $DefaultValue
	)

	while ($true) {
		$fullPrompt = if ([string]::IsNullOrWhiteSpace($DefaultValue)) { $Prompt } else { "$Prompt [$DefaultValue]" }
		$value = Read-Host $fullPrompt
		if ([string]::IsNullOrWhiteSpace($value)) {
			if (-not [string]::IsNullOrWhiteSpace($DefaultValue)) { return $DefaultValue }
			continue
		}

		return $value.Trim()
	}
}

function Read-OptionalValue {
	param(
		[Parameter(Mandatory = $true)][string] $Prompt,
		[Parameter(Mandatory = $false)][string] $DefaultValue
	)

	$fullPrompt = if ([string]::IsNullOrWhiteSpace($DefaultValue)) { $Prompt } else { "$Prompt [$DefaultValue]" }
	$value = Read-Host $fullPrompt
	if ([string]::IsNullOrWhiteSpace($value)) { return $DefaultValue }
	return $value.Trim()
}

function Read-YesNo {
	param(
		[Parameter(Mandatory = $true)][string] $Prompt,
		[Parameter(Mandatory = $false)][switch] $DefaultNo
	)

	$defaultText = if ($DefaultNo) { 'n' } else { 'y' }
	while ($true) {
		$value = Read-Host "$Prompt [y/n, default $defaultText]"
		if ([string]::IsNullOrWhiteSpace($value)) { return -not $DefaultNo }
		switch ($value.Trim().ToLowerInvariant()) {
			'y' { return $true }
			'yes' { return $true }
			'n' { return $false }
			'no' { return $false }
			default { Write-InstallLog 'Please answer y or n.' -Level Warning }
		}
	}
}

function ConvertTo-RecipientArray {
	param([Parameter(Mandatory = $false)][string] $RecipientText)

	if ([string]::IsNullOrWhiteSpace($RecipientText)) {
		return @()
	}

	@($RecipientText -split ',' | ForEach-Object { $_.Trim() } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
}

function Test-IsValidEmailAddress {
	param([Parameter(Mandatory = $true)][string] $Address)

	try {
		$null = [System.Net.Mail.MailAddress]::new($Address)
		return $true
	} catch {
		return $false
	}
}

function Assert-ValidRecipients {
	param(
		[Parameter(Mandatory = $true)][string[]] $Addresses,
		[Parameter(Mandatory = $false)][string] $Label = 'Recipients'
	)

	if ($Addresses.Count -eq 0) {
		throw "$Label is required and must contain at least one e-mail address."
	}

	$invalid = @($Addresses | Where-Object { -not (Test-IsValidEmailAddress -Address $_) })
	if ($invalid.Count -gt 0) {
		throw "$Label contains invalid e-mail address(es): $($invalid -join ', ')"
	}
}

function ConvertTo-TimeOfDaySpan {
	param([Parameter(Mandatory = $true)][string] $Value)

	$timeParts = $Value -split ':'
	if ($timeParts.Count -ne 2) {
		throw 'Time value must be in HH:mm format.'
	}

	$hours = 0
	$minutes = 0
	if (-not [int]::TryParse($timeParts[0], [ref] $hours) -or -not [int]::TryParse($timeParts[1], [ref] $minutes)) {
		throw 'Time value must be numeric HH:mm format.'
	}

	if ($hours -lt 0 -or $hours -gt 23 -or $minutes -lt 0 -or $minutes -gt 59) {
		throw 'Time value must be a valid 24-hour clock time in HH:mm format.'
	}

	[timespan]::new($hours, $minutes, 0)
}

function Select-AzureSubscription {
	param([Parameter(Mandatory = $false)][string] $RequestedSubscriptionId)

	$subscriptions = @(Get-AzSubscription | Sort-Object Name)
	if ($subscriptions.Count -eq 0) {
		throw 'No Azure subscriptions are visible for the current account.'
	}

	if (-not [string]::IsNullOrWhiteSpace($RequestedSubscriptionId)) {
		$match = $subscriptions | Where-Object { $_.Id -eq $RequestedSubscriptionId -or $_.Name -eq $RequestedSubscriptionId } | Select-Object -First 1
		if (-not $match) {
			throw "Subscription '$RequestedSubscriptionId' was not found or is not visible."
		}

		Set-AzContext -SubscriptionId $match.Id | Out-Null
		return $match
	}

	Write-Output 'Available subscriptions:'
	for ($index = 0; $index -lt $subscriptions.Count; $index++) {
		Write-Output ("  {0}. {1} ({2})" -f ($index + 1), $subscriptions[$index].Name, $subscriptions[$index].Id)
	}

	while ($true) {
		$choice = Read-Host 'Select subscription number'
		$number = 0
		if ([int]::TryParse($choice, [ref] $number) -and $number -ge 1 -and $number -le $subscriptions.Count) {
			$selected = $subscriptions[$number - 1]
			Set-AzContext -SubscriptionId $selected.Id | Out-Null
			return $selected
		}
	}
}

function Set-DriftSubscriptionContext {
	param(
		[Parameter(Mandatory = $true)][string] $TargetSubscriptionId,
		[Parameter(Mandatory = $false)][string] $TargetTenantId
	)

	$currentContext = Get-AzContext -ErrorAction SilentlyContinue
	if ($currentContext -and $currentContext.Subscription -and $currentContext.Subscription.Id -eq $TargetSubscriptionId) {
		return
	}

	Write-InstallLog "Setting Azure context to subscription '$TargetSubscriptionId'."
	if ([string]::IsNullOrWhiteSpace($TargetTenantId)) {
		Set-AzContext -SubscriptionId $TargetSubscriptionId -ErrorAction Stop | Out-Null
	} else {
		Set-AzContext -SubscriptionId $TargetSubscriptionId -Tenant $TargetTenantId -ErrorAction Stop | Out-Null
	}
}

function ConvertTo-SafeToken {
	param(
		[Parameter(Mandatory = $true)][string] $Value,
		[Parameter(Mandatory = $false)][int] $MaxLength = 18
	)

	$safe = ($Value.ToLowerInvariant() -replace '[^a-z0-9]', '')
	if ([string]::IsNullOrWhiteSpace($safe)) { $safe = 'dm' }
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

function Get-DriftMaesterNames {
	param(
		[Parameter(Mandatory = $true)][string] $SelectedSubscriptionId,
		[Parameter(Mandatory = $true)][string] $SelectedResourceGroupName
	)

	$suffix = New-DeterministicSuffix -Seed "$SelectedSubscriptionId/$SelectedResourceGroupName"
	$safeGroup = ConvertTo-SafeToken -Value $SelectedResourceGroupName -MaxLength 10
	$automationName = 'driftmaester'
	$storageRaw = "maester$safeGroup$suffix"
	$storageName = ($storageRaw -replace '[^a-z0-9]', '')
	if ($storageName.Length -gt 24) { $storageName = $storageName.Substring(0, 24) }

	[PSCustomObject]@{
		AutomationAccountName = $automationName
		StorageAccountName    = $storageName
	}
}

function Invoke-AzureRest {
	param(
		[Parameter(Mandatory = $true)][ValidateSet('GET', 'PUT', 'POST', 'DELETE')][string] $Method,
		[Parameter(Mandatory = $true)][string] $Path,
		[Parameter(Mandatory = $false)][object] $Body
	)

	$params = @{
		Method      = $Method
		Path        = $Path
		ErrorAction = 'Stop'
	}
	if ($PSBoundParameters.ContainsKey('Body')) {
		$params['Payload'] = ($Body | ConvertTo-Json -Depth 30)
	}

	$response = Invoke-AzRestMethod @params
	if ([string]::IsNullOrWhiteSpace($response.Content)) {
		return $null
	}

	return ($response.Content | ConvertFrom-Json)
}

function Test-IsAzureAuthorizationFailure {
	param([Parameter(Mandatory = $true)][System.Management.Automation.ErrorRecord] $ErrorRecord)

	if (-not $ErrorRecord) { return $false }
	$message = [string] $ErrorRecord.Exception.Message
	if ($message -match 'AuthorizationFailed|Forbidden|does not have authorization|insufficient privileges|status code 403|status code 401') {
		return $true
	}

	return $false
}

function Register-DriftAzureProvider {
	param([Parameter(Mandatory = $true)][string] $ProviderNamespace)

	$provider = Get-AzResourceProvider -ProviderNamespace $ProviderNamespace -ErrorAction SilentlyContinue
	if ($provider -and $provider.RegistrationState -eq 'Registered') {
		return
	}

	Write-InstallLog "Registering Azure resource provider '$ProviderNamespace'."
	Register-AzResourceProvider -ProviderNamespace $ProviderNamespace | Out-Null
}

function Set-DriftResourceGroup {
	param(
		[Parameter(Mandatory = $true)][string] $Name,
		[Parameter(Mandatory = $true)][string] $TargetLocation
	)

	$resourceGroup = Get-AzResourceGroup -Name $Name -ErrorAction SilentlyContinue
	if ($resourceGroup) {
		Write-InstallLog "Resource group '$Name' already exists."
		return $resourceGroup
	}

	Write-InstallLog "Creating resource group '$Name' in '$TargetLocation'."
	return (New-AzResourceGroup -Name $Name -Location $TargetLocation -Tag @{ DriftMaesterManaged = 'true' })
}

function Set-DriftAutomationAccount {
	param(
		[Parameter(Mandatory = $true)][string] $Name,
		[Parameter(Mandatory = $true)][string] $TargetResourceGroupName,
		[Parameter(Mandatory = $true)][string] $TargetLocation
	)

	$automationAccount = Get-AzAutomationAccount -ResourceGroupName $TargetResourceGroupName -Name $Name -ErrorAction SilentlyContinue
	if (-not $automationAccount) {
		Write-InstallLog "Creating Automation Account '$Name' with system-assigned managed identity."
		$automationAccount = New-AzAutomationAccount -ResourceGroupName $TargetResourceGroupName -Name $Name -Location $TargetLocation -Plan Basic -AssignSystemIdentity
	} else {
		Write-InstallLog "Automation Account '$Name' already exists."
		if (-not $automationAccount.Identity -or -not $automationAccount.Identity.PrincipalId) {
			Write-InstallLog "Enabling system-assigned managed identity on Automation Account '$Name'."
			Set-AzAutomationAccount -ResourceGroupName $TargetResourceGroupName -Name $Name -AssignSystemIdentity | Out-Null
			$automationAccount = Get-AzAutomationAccount -ResourceGroupName $TargetResourceGroupName -Name $Name
		}
	}

	return $automationAccount
}

function Set-DriftStorageAccount {
	param(
		[Parameter(Mandatory = $true)][string] $Name,
		[Parameter(Mandatory = $true)][string] $TargetResourceGroupName,
		[Parameter(Mandatory = $true)][string] $TargetLocation
	)

	$storageAccount = Get-AzStorageAccount -ResourceGroupName $TargetResourceGroupName -Name $Name -ErrorAction SilentlyContinue
	if (-not $storageAccount) {
		Write-InstallLog "Creating storage account '$Name'."
		$storageAccount = New-AzStorageAccount -ResourceGroupName $TargetResourceGroupName -Name $Name -Location $TargetLocation -SkuName Standard_LRS -Kind StorageV2 -EnableHttpsTrafficOnly $true -MinimumTlsVersion TLS1_2 -AllowBlobPublicAccess $false -AllowSharedKeyAccess $false
	} else {
		Write-InstallLog "Storage account '$Name' already exists."
		Write-InstallLog "Applying storage hardening baseline to '$Name' (HTTPS only, TLS1_2, no public blob access, no shared key auth)."
		$null = Set-AzStorageAccount -ResourceGroupName $TargetResourceGroupName -Name $Name -EnableHttpsTrafficOnly $true -MinimumTlsVersion TLS1_2 -AllowBlobPublicAccess $false -AllowSharedKeyAccess $false
		$storageAccount = Get-AzStorageAccount -ResourceGroupName $TargetResourceGroupName -Name $Name
	}

	$context = $storageAccount.Context
	if (-not (Get-AzStorageContainer -Context $context -Name 'maester' -ErrorAction SilentlyContinue)) {
		Write-InstallLog "Creating blob container 'maester'."
		New-AzStorageContainer -Context $context -Name 'maester' -Permission Off | Out-Null
	}

	return $storageAccount
}

function Set-DriftStorageSecurityPosture {
	param(
		[Parameter(Mandatory = $true)][string] $ResourceGroupName,
		[Parameter(Mandatory = $true)][string] $StorageAccountName
	)

	Write-InstallLog "Enabling blob and container soft delete plus blob versioning on '$StorageAccountName'."
	$null = Enable-AzStorageBlobDeleteRetentionPolicy -ResourceGroupName $ResourceGroupName -StorageAccountName $StorageAccountName -RetentionDays 30
	$null = Enable-AzStorageContainerDeleteRetentionPolicy -ResourceGroupName $ResourceGroupName -StorageAccountName $StorageAccountName -RetentionDays 30
	$null = Update-AzStorageBlobServiceProperty -ResourceGroupName $ResourceGroupName -StorageAccountName $StorageAccountName -IsVersioningEnabled $true
}

function Set-DriftStorageLifecyclePolicy {
	param(
		[Parameter(Mandatory = $true)][string] $ResourceGroupName,
		[Parameter(Mandatory = $true)][string] $StorageAccountName,
		[Parameter(Mandatory = $true)][ValidateRange(30, 3650)][int] $RetentionDays
	)

	$filter = New-AzStorageAccountManagementPolicyFilter -BlobType blockBlob -PrefixMatch @('maester/')
	if (-not $filter) { throw 'Failed to build the storage management policy filter (New-AzStorageAccountManagementPolicyFilter returned null).' }

	$baseAction = Add-AzStorageAccountManagementPolicyAction -BaseBlobAction Delete -DaysAfterModificationGreaterThan $RetentionDays
	if (-not $baseAction) { throw 'Failed to build the storage management policy action (Add-AzStorageAccountManagementPolicyAction returned null).' }

	$rule = New-AzStorageAccountManagementPolicyRule -Name 'driftmaester-retention' -Action $baseAction -Filter $filter
	if (-not $rule) { throw 'Failed to build the storage management policy rule (New-AzStorageAccountManagementPolicyRule returned null).' }

	$existingPolicy = Get-AzStorageAccountManagementPolicy -ResourceGroupName $ResourceGroupName -StorageAccountName $StorageAccountName -ErrorAction SilentlyContinue
	if ($existingPolicy) {
		Write-InstallLog "Updating storage lifecycle policy 'driftmaester-retention' to keep DriftMaester artifacts for $RetentionDays day(s)."
		$remainingRules = @($existingPolicy.Policy.Rules | Where-Object { $_ -and $_.Name -ne 'driftmaester-retention' })
		$ruleSet = @(@($remainingRules) + $rule | Where-Object { $_ })
		Set-AzStorageAccountManagementPolicy -ResourceGroupName $ResourceGroupName -StorageAccountName $StorageAccountName -Rule $ruleSet | Out-Null
	} else {
		Write-InstallLog "Creating storage lifecycle policy 'driftmaester-retention' with retention of $RetentionDays day(s)."
		Set-AzStorageAccountManagementPolicy -ResourceGroupName $ResourceGroupName -StorageAccountName $StorageAccountName -Rule @($rule) | Out-Null
	}
}

function Set-DriftAzRoleAssignment {
	param(
		[Parameter(Mandatory = $true)][string] $ObjectId,
		[Parameter(Mandatory = $true)][string] $RoleDefinitionName,
		[Parameter(Mandatory = $true)][string] $Scope
	)

	$existing = Get-AzRoleAssignment -ObjectId $ObjectId -RoleDefinitionName $RoleDefinitionName -Scope $Scope -ErrorAction SilentlyContinue | Select-Object -First 1
	if ($existing) {
		Write-InstallLog "Azure RBAC '$RoleDefinitionName' already assigned at '$Scope'."
		return
	}

	Write-InstallLog "Assigning Azure RBAC '$RoleDefinitionName' at '$Scope'."
	New-AzRoleAssignment -ObjectId $ObjectId -RoleDefinitionName $RoleDefinitionName -Scope $Scope | Out-Null
}

function Get-DriftDefaultTimeZoneId {
	return [System.TimeZoneInfo]::Local.Id
}

function Resolve-DriftTimeZoneId {
	param([Parameter(Mandatory = $false)][string] $TimeZoneId)

	if ([string]::IsNullOrWhiteSpace($TimeZoneId)) {
		return Get-DriftDefaultTimeZoneId
	}

	try {
		return [System.TimeZoneInfo]::FindSystemTimeZoneById($TimeZoneId.Trim()).Id
	} catch {
		Write-InstallLog "Time zone '$TimeZoneId' could not be validated on this system. Passing it to Azure Automation as provided." -Level Warning
		return $TimeZoneId.Trim()
	}
}

function Get-DriftScheduleNow {
	param([Parameter(Mandatory = $true)][string] $TimeZoneId)

	try {
		$timeZoneInfo = [System.TimeZoneInfo]::FindSystemTimeZoneById($TimeZoneId)
		return [System.TimeZoneInfo]::ConvertTimeFromUtc([datetime]::UtcNow, $timeZoneInfo)
	} catch {
		return Get-Date
	}
}

function Get-NextDailyOccurrence {
	param(
		[Parameter(Mandatory = $true)][timespan] $TimeOfDay,
		[Parameter(Mandatory = $true)][string] $TimeZoneId
	)

	$now = Get-DriftScheduleNow -TimeZoneId $TimeZoneId
	$candidate = [datetime]::new($now.Year, $now.Month, $now.Day, 0, 0, 0).Add($TimeOfDay)
	if ($candidate -le $now.AddMinutes(10)) {
		$candidate = $candidate.AddDays(1)
	}

	return $candidate
}

function Get-NextWeeklyOccurrence {
	param(
		[Parameter(Mandatory = $true)][string] $DayOfWeek,
		[Parameter(Mandatory = $true)][timespan] $TimeOfDay,
		[Parameter(Mandatory = $true)][string] $TimeZoneId
	)

	$day = [System.Enum]::Parse([System.DayOfWeek], $DayOfWeek, $true)
	$now = Get-DriftScheduleNow -TimeZoneId $TimeZoneId
	$candidate = [datetime]::new($now.Year, $now.Month, $now.Day, 0, 0, 0).Add($TimeOfDay)
	$delta = (([int] $day) - ([int] $candidate.DayOfWeek) + 7) % 7
	$candidate = $candidate.AddDays($delta)
	if ($candidate -le $now.AddMinutes(10)) {
		$candidate = $candidate.AddDays(7)
	}

	return $candidate
}

function Get-NextMonthlyOccurrence {
	param(
		[Parameter(Mandatory = $true)][ValidateSet('First', 'DayOfMonth')][string] $DayMode,
		[Parameter(Mandatory = $false)][ValidateRange(-1, 31)][int] $DayOfMonth = 1,
		[Parameter(Mandatory = $true)][timespan] $TimeOfDay,
		[Parameter(Mandatory = $true)][string] $TimeZoneId
	)

	$now = Get-DriftScheduleNow -TimeZoneId $TimeZoneId
	$targetDay = if ($DayMode -eq 'DayOfMonth') {
		if ($DayOfMonth -eq -1) {
			[DateTime]::DaysInMonth($now.Year, $now.Month)
		} else {
			[Math]::Min([Math]::Max(1, $DayOfMonth), [DateTime]::DaysInMonth($now.Year, $now.Month))
		}
	} else {
		1
	}

	$candidate = [datetime]::new($now.Year, $now.Month, $targetDay, $TimeOfDay.Hours, $TimeOfDay.Minutes, 0)
	if ($candidate -le $now.AddMinutes(10)) {
		$nextMonth = $now.AddMonths(1)
		$nextDay = if ($DayMode -eq 'DayOfMonth') {
			if ($DayOfMonth -eq -1) {
				[DateTime]::DaysInMonth($nextMonth.Year, $nextMonth.Month)
			} else {
				[Math]::Min([Math]::Max(1, $DayOfMonth), [DateTime]::DaysInMonth($nextMonth.Year, $nextMonth.Month))
			}
		} else {
			1
		}
		$candidate = [datetime]::new($nextMonth.Year, $nextMonth.Month, $nextDay, $TimeOfDay.Hours, $TimeOfDay.Minutes, 0)
	}

	return $candidate
}

function New-DriftScheduleSelection {
	$frequencyInput = Read-RequiredValue -Prompt 'Run frequency (daily, weekly, monthly)' -DefaultValue 'daily'
	$frequencyInput = $frequencyInput.Trim().ToLowerInvariant()
	$timeZoneId = Resolve-DriftTimeZoneId -TimeZoneId (Read-OptionalValue -Prompt 'Schedule time zone id, press enter to use the local time zone' -DefaultValue (Get-DriftDefaultTimeZoneId))
	$timeInput = Read-RequiredValue -Prompt "Run time (HH:mm in $timeZoneId)" -DefaultValue '02:00'
	$timeOfDay = ConvertTo-TimeOfDaySpan -Value $timeInput

	switch ($frequencyInput) {
		'weekly' {
			$dayOfWeek = Read-RequiredValue -Prompt 'Day of week for weekly schedule (Monday..Sunday)' -DefaultValue 'Monday'
			return [PSCustomObject]@{
				Frequency       = 'Week'
				StartTime       = Get-NextWeeklyOccurrence -DayOfWeek $dayOfWeek -TimeOfDay $timeOfDay -TimeZoneId $timeZoneId
				TimeOfDay       = $timeOfDay
				TimeZoneId      = $timeZoneId
				WeekDays        = @($dayOfWeek)
				MonthDays       = @()
				MonthlyDayMode  = $null
				DescriptionText = "weekly on $dayOfWeek"
			}
		}
		'monthly' {
			$monthDayInput = Read-OptionalValue -Prompt 'Day of month for monthly schedule (1-31 or -1 for last day)' -DefaultValue '1'
			$monthDay = 1
			if (-not [int]::TryParse($monthDayInput, [ref] $monthDay) -or ($monthDay -lt 1 -or $monthDay -gt 31) -and $monthDay -ne -1) {
				throw 'Monthly day must be 1-31 or -1 (last day of month).'
			}
			$monthDayText = if ($monthDay -eq -1) { 'last day' } else { "$monthDay" }
			return [PSCustomObject]@{
				Frequency       = 'Month'
				StartTime       = Get-NextMonthlyOccurrence -DayMode 'DayOfMonth' -DayOfMonth $monthDay -TimeOfDay $timeOfDay -TimeZoneId $timeZoneId
				TimeOfDay       = $timeOfDay
				TimeZoneId      = $timeZoneId
				WeekDays        = @()
				MonthDays       = @($monthDay)
				MonthlyDayMode  = 'DayOfMonth'
				DescriptionText = "monthly on day $monthDayText"
			}
		}
		default {
			return [PSCustomObject]@{
				Frequency       = 'Day'
				StartTime       = Get-NextDailyOccurrence -TimeOfDay $timeOfDay -TimeZoneId $timeZoneId
				TimeOfDay       = $timeOfDay
				TimeZoneId      = $timeZoneId
				WeekDays        = @()
				MonthDays       = @()
				MonthlyDayMode  = $null
				DescriptionText = 'daily'
			}
		}
	}
}

function New-UpdateScheduleSelection {
	param([Parameter(Mandatory = $true)][pscustomobject] $InvokeSchedule)

	$scheduleTimeZone = if ($InvokeSchedule.TimeZoneId) { [string] $InvokeSchedule.TimeZoneId } else { Get-DriftDefaultTimeZoneId }
	$invokeStart = [datetime] $InvokeSchedule.StartTime
	$start = $invokeStart.AddHours(-1)
	$crossedDayBoundary = $start.Date -ne $invokeStart.Date
	$now = Get-DriftScheduleNow -TimeZoneId $scheduleTimeZone
	$updateWeekDays = @($InvokeSchedule.WeekDays)
	$updateMonthDays = @($InvokeSchedule.MonthDays)

	if ($crossedDayBoundary) {
		switch ($InvokeSchedule.Frequency) {
			'Week' {
				$shiftedDays = foreach ($weekDay in @($InvokeSchedule.WeekDays)) {
					try {
						$parsed = [System.Enum]::Parse([System.DayOfWeek], [string] $weekDay, $true)
						([System.DayOfWeek](((7 + [int]$parsed - 1) % 7))).ToString()
					} catch {
						$weekDay
					}
				}
				$updateWeekDays = @($shiftedDays)
			}
			'Month' {
				$shiftedMonthDays = foreach ($day in @($InvokeSchedule.MonthDays)) {
					$parsedDay = 0
					if ([int]::TryParse([string] $day, [ref] $parsedDay)) {
						if ($parsedDay -eq 1) { -1 } elseif ($parsedDay -gt 1) { $parsedDay - 1 } else { $parsedDay }
					} else {
						$day
					}
				}
				$updateMonthDays = @($shiftedMonthDays)
			}
		}
	}

	if ($start -le $now.AddMinutes(10)) {
		switch ($InvokeSchedule.Frequency) {
			'Week' { $start = $start.AddDays(7) }
			'Month' { $start = $start.AddMonths(1) }
			default { $start = $start.AddDays(1) }
		}
	}

	return [PSCustomObject]@{
		Frequency       = $InvokeSchedule.Frequency
		StartTime       = $start
		TimeOfDay       = $start.TimeOfDay
		TimeZoneId      = $scheduleTimeZone
		WeekDays        = $updateWeekDays
		MonthDays       = $updateMonthDays
		MonthlyDayMode  = $InvokeSchedule.MonthlyDayMode
		DescriptionText = "$($InvokeSchedule.DescriptionText), one hour before invoke"
	}
}

function Enable-DriftAzureRootAccess {
	Write-InstallLog 'Attempting to elevate Azure root access for the current account. This only works for a Global Administrator with access management elevation allowed.' -Level Warning
	Invoke-AzRestMethod -Path '/providers/Microsoft.Authorization/elevateAccess?api-version=2015-07-01' -Method POST | Out-Null
}

function Get-CurrentPrincipalObjectIdFromArmToken {
	$tokenResponse = Get-AzAccessToken -ResourceUrl 'https://management.azure.com/' -AsSecureString
	$token = [System.Net.NetworkCredential]::new('', $tokenResponse.Token).Password
	$parts = $token.Split('.')
	if ($parts.Count -lt 2) {
		return $null
	}

	$payload = $parts[1].Replace('-', '+').Replace('_', '/')
	switch ($payload.Length % 4) {
		2 { $payload += '==' }
		3 { $payload += '=' }
		1 { $payload += '===' }
	}

	try {
		$json = [Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($payload)) | ConvertFrom-Json
		return [string] $json.oid
	} catch {
		return $null
	}
}

function Disable-DriftAzureRootAccess {
	$currentContext = Get-AzContext
	$currentAccountId = [string] $currentContext.Account.Id
	$currentPrincipalObjectId = Get-CurrentPrincipalObjectIdFromArmToken
	if ([string]::IsNullOrWhiteSpace($currentAccountId) -and [string]::IsNullOrWhiteSpace($currentPrincipalObjectId)) {
		Write-InstallLog 'Could not determine the current Azure account principal for root access cleanup.' -Level Warning
		return
	}

	$assignments = @(Get-AzRoleAssignment -RoleDefinitionId '18d7d88d-d35e-4fb5-a5c3-7773c20a72d9' -ErrorAction SilentlyContinue | Where-Object {
		$matchesSignInName = -not [string]::IsNullOrWhiteSpace($currentAccountId) -and $_.SignInName -eq $currentAccountId
		$matchesObjectId = -not [string]::IsNullOrWhiteSpace($currentPrincipalObjectId) -and [string] $_.ObjectId -eq $currentPrincipalObjectId
		$_.Scope -eq '/' -and ($matchesSignInName -or $matchesObjectId)
	})

	foreach ($assignment in $assignments) {
		$assignmentPath = [string] $assignment.RoleAssignmentId
		if ([string]::IsNullOrWhiteSpace($assignmentPath)) { continue }
		if (-not $assignmentPath.StartsWith('/')) { $assignmentPath = "/$assignmentPath" }

		Write-InstallLog "Removing temporary Azure root User Access Administrator elevation for '$currentAccountId' ($currentPrincipalObjectId)."
		Invoke-AzRestMethod -Path "$assignmentPath?api-version=2018-07-01" -Method DELETE | Out-Null
	}

	if (-not [string]::IsNullOrWhiteSpace($currentPrincipalObjectId)) {
		$leftover = @(Get-AzRoleAssignment -RoleDefinitionId '18d7d88d-d35e-4fb5-a5c3-7773c20a72d9' -ObjectId $currentPrincipalObjectId -Scope '/' -ErrorAction SilentlyContinue)
		if ($leftover.Count -gt 0) {
			Write-InstallLog "Warning: $($leftover.Count) root User Access Administrator assignment(s) remain for principal '$currentPrincipalObjectId'. Review and remove manually if these were temporary elevations." -Level Warning
		}
	}
}

function Set-DriftAzureRootReaderAssignment {
	param(
		[Parameter(Mandatory = $true)][string] $ServicePrincipalObjectId,
		[Parameter(Mandatory = $true)][string] $Scope
	)

	$existing = Get-AzRoleAssignment -ObjectId $ServicePrincipalObjectId -RoleDefinitionName 'Reader' -Scope $Scope -ErrorAction Stop | Select-Object -First 1
	if ($existing) {
		Write-InstallLog "Azure root RBAC 'Reader' already assigned at '$Scope'."
		return
	}

	Write-InstallLog "Assigning Azure root RBAC 'Reader' at '$Scope'."
	New-AzRoleAssignment -ObjectId $ServicePrincipalObjectId -Scope $Scope -RoleDefinitionName 'Reader' -ObjectType 'ServicePrincipal' -ErrorAction Stop | Out-Null
}

function Set-DriftAzureRootReaderAssignments {
	param([Parameter(Mandatory = $true)][string] $ServicePrincipalObjectId)

	$scopes = @('/', '/providers/Microsoft.aadiam')
	$elevated = $false

	try {
		foreach ($scope in $scopes) {
			Set-DriftAzureRootReaderAssignment -ServicePrincipalObjectId $ServicePrincipalObjectId -Scope $scope
		}
		return
	} catch {
		if (-not (Test-IsAzureAuthorizationFailure -ErrorRecord $_)) { throw }
		Write-InstallLog "Azure root Reader assignment failed due to insufficient access: $($_.Exception.Message)" -Level Warning
	}

	try {
		Enable-DriftAzureRootAccess
		$elevated = $true

		foreach ($scope in $scopes) {
			Set-DriftAzureRootReaderAssignment -ServicePrincipalObjectId $ServicePrincipalObjectId -Scope $scope
		}
	} finally {
		if ($elevated) {
			Disable-DriftAzureRootAccess
		}
	}
}

function Set-DriftRuntimeEnvironment {
	param(
		[Parameter(Mandatory = $true)][string] $SelectedSubscriptionId,
		[Parameter(Mandatory = $true)][string] $TargetResourceGroupName,
		[Parameter(Mandatory = $true)][string] $AutomationAccountName,
		[Parameter(Mandatory = $true)][string] $TargetLocation
	)

	$runtimePath = "/subscriptions/$SelectedSubscriptionId/resourceGroups/$TargetResourceGroupName/providers/Microsoft.Automation/automationAccounts/$AutomationAccountName/runtimeEnvironments/$($script:PreferredRuntimeName)?api-version=$($script:AutomationApiVersion)"
	$existing = $null
	try {
		$existing = Invoke-AzureRest -Method GET -Path $runtimePath
	} catch {
		$existing = $null
	}

	$runtimeVersionMatches = $existing -and [string] $existing.properties.runtime.version -like "$($script:RuntimeVersion)*"
	$defaultPackagesConfigured = $false
	if ($runtimeVersionMatches -and $existing.properties.defaultPackages) {
		$defaultPackagesConfigured = $true
		foreach ($packageName in $script:DefaultRuntimePackages.Keys) {
			$currentVersion = [string] $existing.properties.defaultPackages.$packageName
			if ($currentVersion -ne $script:DefaultRuntimePackages[$packageName]) {
				$defaultPackagesConfigured = $false
				break
			}
		}
	}

	if ($runtimeVersionMatches -and $defaultPackagesConfigured) {
		Write-InstallLog "Runtime environment '$($script:PreferredRuntimeName)' already exists."
		return $existing
	}

	Write-InstallLog "Creating/updating runtime environment '$($script:PreferredRuntimeName)' with PowerShell $($script:RuntimeVersion) and default packages."
	$body = @{
		location   = $TargetLocation
		properties = @{
			defaultPackages = $script:DefaultRuntimePackages
			runtime     = @{
				language = 'PowerShell'
				version  = $script:RuntimeVersion
			}
			description = 'Runtime environment for DriftMaester automation.'
		}
	}
	return Invoke-AzureRest -Method PUT -Path $runtimePath -Body $body
}

function Set-DriftRunbookFromGithub {
	param(
		[Parameter(Mandatory = $true)][string] $SelectedSubscriptionId,
		[Parameter(Mandatory = $true)][string] $TargetResourceGroupName,
		[Parameter(Mandatory = $true)][string] $AutomationAccountName,
		[Parameter(Mandatory = $true)][string] $TargetLocation,
		[Parameter(Mandatory = $true)][string] $RunbookName,
		[Parameter(Mandatory = $true)][string] $SourceFileName
	)

	$contentUri = "$($script:GithubRawBase)/$SourceFileName"
	$escapedRunbookName = [Uri]::EscapeDataString($RunbookName)
	$path = "/subscriptions/$SelectedSubscriptionId/resourceGroups/$TargetResourceGroupName/providers/Microsoft.Automation/automationAccounts/$AutomationAccountName/runbooks/${escapedRunbookName}?api-version=$($script:AutomationApiVersion)"
	$body = @{
		location   = $TargetLocation
		properties = @{
			runbookType        = 'PowerShell'
			runtimeEnvironment = $script:PreferredRuntimeName
			logVerbose         = $false
			logProgress        = $false
			description        = "DriftMaester $RunbookName runbook."
			publishContentLink = @{
				uri = $contentUri
			}
		}
	}

	Write-InstallLog "Creating/updating runbook '$RunbookName' from '$contentUri'."
	Invoke-AzureRest -Method PUT -Path $path -Body $body | Out-Null
}

function Set-DriftAutomationSchedule {
	param(
		[Parameter(Mandatory = $true)][string] $SelectedSubscriptionId,
		[Parameter(Mandatory = $true)][string] $TargetResourceGroupName,
		[Parameter(Mandatory = $true)][string] $AutomationAccountName,
		[Parameter(Mandatory = $true)][string] $ScheduleName,
		[Parameter(Mandatory = $true)][pscustomobject] $ScheduleSelection,
		[Parameter(Mandatory = $true)][string] $Description
	)

	$escapedScheduleName = [Uri]::EscapeDataString($ScheduleName)
	$path = "/subscriptions/$SelectedSubscriptionId/resourceGroups/$TargetResourceGroupName/providers/Microsoft.Automation/automationAccounts/$AutomationAccountName/schedules/${escapedScheduleName}?api-version=2023-11-01"
	$scheduleTimeZone = if ($ScheduleSelection.TimeZoneId) { [string] $ScheduleSelection.TimeZoneId } else { Get-DriftDefaultTimeZoneId }
	$body = @{
		properties = @{
			description = $Description
			startTime   = $ScheduleSelection.StartTime.ToString('o')
			expiryTime  = '9999-12-31T23:59:59Z'
			interval    = 1
			frequency   = $ScheduleSelection.Frequency
			timeZone    = $scheduleTimeZone
		}
	}

	if ($ScheduleSelection.Frequency -eq 'Week') {
		$body.properties['advancedSchedule'] = @{ weekDays = @($ScheduleSelection.WeekDays) }
	} elseif ($ScheduleSelection.Frequency -eq 'Month') {
		$body.properties['advancedSchedule'] = @{ monthDays = @($ScheduleSelection.MonthDays) }
	}

	Write-InstallLog "Creating/updating schedule '$ScheduleName' starting '$($ScheduleSelection.StartTime.ToString('u'))' in time zone '$scheduleTimeZone' ($($ScheduleSelection.DescriptionText))."
	Invoke-AzureRest -Method PUT -Path $path -Body $body | Out-Null
}

function Remove-ExistingJobSchedules {
	param(
		[Parameter(Mandatory = $true)][string] $SelectedSubscriptionId,
		[Parameter(Mandatory = $true)][string] $TargetResourceGroupName,
		[Parameter(Mandatory = $true)][string] $AutomationAccountName,
		[Parameter(Mandatory = $true)][string] $RunbookName,
		[Parameter(Mandatory = $true)][string] $ScheduleName
	)

	$path = "/subscriptions/$SelectedSubscriptionId/resourceGroups/$TargetResourceGroupName/providers/Microsoft.Automation/automationAccounts/$AutomationAccountName/jobSchedules?api-version=2023-11-01"
	$jobSchedules = @((Invoke-AzureRest -Method GET -Path $path).value)
	foreach ($jobSchedule in $jobSchedules) {
		if ([string] $jobSchedule.properties.runbook.name -ieq $RunbookName -and [string] $jobSchedule.properties.schedule.name -ieq $ScheduleName) {
			$deletePath = "/subscriptions/$SelectedSubscriptionId/resourceGroups/$TargetResourceGroupName/providers/Microsoft.Automation/automationAccounts/$AutomationAccountName/jobSchedules/$($jobSchedule.name)?api-version=2023-11-01"
			Write-InstallLog "Removing existing job schedule '$($jobSchedule.name)' for '$RunbookName' / '$ScheduleName'."
			Invoke-AzureRest -Method DELETE -Path $deletePath | Out-Null
		}
	}
}

function Set-DriftJobSchedule {
	param(
		[Parameter(Mandatory = $true)][string] $SelectedSubscriptionId,
		[Parameter(Mandatory = $true)][string] $TargetResourceGroupName,
		[Parameter(Mandatory = $true)][string] $AutomationAccountName,
		[Parameter(Mandatory = $true)][string] $RunbookName,
		[Parameter(Mandatory = $true)][string] $ScheduleName,
		[Parameter(Mandatory = $false)][hashtable] $Parameters = @{}
	)

	Remove-ExistingJobSchedules -SelectedSubscriptionId $SelectedSubscriptionId -TargetResourceGroupName $TargetResourceGroupName -AutomationAccountName $AutomationAccountName -RunbookName $RunbookName -ScheduleName $ScheduleName

	$jobScheduleId = [guid]::NewGuid().ToString()
	$path = "/subscriptions/$SelectedSubscriptionId/resourceGroups/$TargetResourceGroupName/providers/Microsoft.Automation/automationAccounts/$AutomationAccountName/jobSchedules/${jobScheduleId}?api-version=2023-11-01"
	$body = @{
		properties = @{
			schedule   = @{ name = $ScheduleName }
			runbook    = @{ name = $RunbookName }
			parameters = $Parameters
		}
	}

	Write-InstallLog "Linking runbook '$RunbookName' to schedule '$ScheduleName'."
	Invoke-AzureRest -Method PUT -Path $path -Body $body | Out-Null
}

function Start-UpdateRunbook {
	param(
		[Parameter(Mandatory = $true)][string] $SelectedSubscriptionId,
		[Parameter(Mandatory = $true)][string] $TargetResourceGroupName,
		[Parameter(Mandatory = $true)][string] $AutomationAccountName
	)

	$jobName = [guid]::NewGuid().ToString()
	$path = "/subscriptions/$SelectedSubscriptionId/resourceGroups/$TargetResourceGroupName/providers/Microsoft.Automation/automationAccounts/$($AutomationAccountName)/jobs/$($jobName)?api-version=2023-11-01"
	$body = @{
		properties = @{
			runbook = @{ name = $script:UpdateRunbookName }
		}
	}

	Write-InstallLog "Starting '$($script:UpdateRunbookName)' immediately so the runtime modules are installed or updated."
	Invoke-AzureRest -Method PUT -Path $path -Body $body | Out-Null
	return $jobName
}

function Start-InvokeRunbookNow {
	param(
		[Parameter(Mandatory = $true)][string] $SelectedSubscriptionId,
		[Parameter(Mandatory = $true)][string] $TargetResourceGroupName,
		[Parameter(Mandatory = $true)][string] $AutomationAccountName,
		[Parameter(Mandatory = $true)][hashtable] $Parameters
	)

	$jobName = [guid]::NewGuid().ToString()
	$path = "/subscriptions/$SelectedSubscriptionId/resourceGroups/$TargetResourceGroupName/providers/Microsoft.Automation/automationAccounts/$($AutomationAccountName)/jobs/$($jobName)?api-version=2023-11-01"
	$body = @{
		properties = @{
			runbook    = @{ name = $script:InvokeRunbookName }
			parameters = $Parameters
		}
	}

	Write-InstallLog "Starting '$($script:InvokeRunbookName)' immediately with configured parameters."
	Invoke-AzureRest -Method PUT -Path $path -Body $body | Out-Null
	return $jobName
}

function Wait-ForAutomationJobCompletion {
	param(
		[Parameter(Mandatory = $true)][string] $ResourceGroupName,
		[Parameter(Mandatory = $true)][string] $AutomationAccountName,
		[Parameter(Mandatory = $true)][string] $JobId,
		[Parameter(Mandatory = $false)][int] $TimeoutMinutes = 30,
		[Parameter(Mandatory = $false)][int] $PollSeconds = 20
	)

	$deadline = (Get-Date).AddMinutes($TimeoutMinutes)
	$lastStatus = $null
	while ((Get-Date) -lt $deadline) {
		$job = Get-AzAutomationJob -ResourceGroupName $ResourceGroupName -AutomationAccountName $AutomationAccountName -Id $JobId -ErrorAction SilentlyContinue
		if (-not $job) {
			Start-Sleep -Seconds $PollSeconds
			continue
		}

		$status = [string] $job.Status
		if ($status -ne $lastStatus) {
			Write-InstallLog "Automation job '$JobId' status: $status"
			$lastStatus = $status
		}

		if ($status -in @('Completed', 'Stopped', 'Failed', 'Suspended')) {
			return $job
		}

		Start-Sleep -Seconds $PollSeconds
	}

	throw "Timed out waiting for Automation job '$JobId' after $TimeoutMinutes minute(s)."
}

function Write-DriftAccessReport {
	param(
		[Parameter(Mandatory = $true)][string] $ManagedIdentityObjectId,
		[Parameter(Mandatory = $true)][string] $ManagedIdentityClientId,
		[Parameter(Mandatory = $true)][string] $SubscriptionId,
		[Parameter(Mandatory = $true)][string] $ResourceGroupName,
		[Parameter(Mandatory = $true)][string] $AutomationAccountName,
		[Parameter(Mandatory = $true)][string[]] $GraphPermissions,
		[Parameter(Mandatory = $true)][string[]] $DirectoryRoles,
		[Parameter(Mandatory = $false)][string[]] $SharePointPermissions = @(),
		[Parameter(Mandatory = $false)][string] $ExchangeOrganization
	)

	Write-InstallLog 'Access report (effective grants configured by installer):' -Level Success
	Write-InstallLog "  Managed identity object id: $ManagedIdentityObjectId"
	Write-InstallLog "  Managed identity client id: $ManagedIdentityClientId"
	Write-InstallLog "  Subscription scope reader: /subscriptions/$SubscriptionId/resourceGroups/$ResourceGroupName"
	Write-InstallLog "  Storage data role: Storage Blob Data Contributor on storage account in '$ResourceGroupName'"
	Write-InstallLog "  Automation role: Automation Contributor on '$AutomationAccountName'"
	Write-InstallLog '  Root scopes: Reader on / and /providers/Microsoft.aadiam'
	Write-InstallLog ("  Graph application permissions: {0}" -f (($GraphPermissions | Sort-Object -Unique) -join ', '))
	Write-InstallLog '  Exchange application permission: Exchange.ManageAsApp'
	if ($SharePointPermissions.Count -gt 0) {
		Write-InstallLog ("  SharePoint Online application permissions: {0}" -f (($SharePointPermissions | Sort-Object -Unique) -join ', '))
	}
	Write-InstallLog ("  Directory roles: {0}" -f (($DirectoryRoles | Sort-Object -Unique) -join ', '))
	if (-not [string]::IsNullOrWhiteSpace($ExchangeOrganization)) {
		Write-InstallLog "  Exchange organization used for RBAC setup: $ExchangeOrganization"
	}
}

function Connect-ToGraphWithAzureToken {
	# Preferred path: hand Connect-MgGraph a Graph token minted from the Azure sign-in that already happened.
	# The operator signs in once instead of twice, and because -AccessToken never touches MSAL this also works in
	# sessions where another module has already pinned an incompatible Microsoft.Identity.Client.
	# The Azure PowerShell client is pre-consented for the delegated scopes this installer needs, including
	# AppRoleAssignment.ReadWrite.All, Application.ReadWrite.All and Directory.AccessAsUser.All.
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

		# Prove the token is actually usable for the directory reads and writes that follow, rather than
		# discovering half way through the install that it is not.
		$null = Invoke-MgGraphRequest -Method GET -Uri 'https://graph.microsoft.com/v1.0/servicePrincipals?$top=1&$select=id' -ErrorAction Stop

		$graphContext = Get-MgContext -ErrorAction SilentlyContinue
		Write-InstallLog "Reusing the Azure sign-in for Microsoft Graph (tenant $($graphContext.TenantId)). No second sign-in needed."
		return $true
	} catch {
		Write-InstallLog "Could not reuse the Azure sign-in for Microsoft Graph: $($_.Exception.Message)" -Level Warning
		Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
		return $false
	}
}

function Connect-ToGraphForInstall {
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
			Write-Output 'Current Microsoft Graph login:'
			Write-Output "  Account:   $($context.Account)"
			Write-Output "  Tenant:    $($context.TenantId)"
			Write-Output "  Client ID: $($context.ClientId)"
			if ($RequestedTenantId -and $context.TenantId -ne $RequestedTenantId) {
				Write-InstallLog "The current Graph tenant does not match the selected tenant '$RequestedTenantId'." -Level Warning
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
			Write-InstallLog 'Falling back to an interactive Microsoft Graph sign-in.'
		}

		$params = @{ Scopes = $requiredScopes; NoWelcome = $true }
		if ($RequestedTenantId) { $params['TenantId'] = $RequestedTenantId }
		Write-InstallLog 'Connecting to Microsoft Graph. Use a Global Administrator or Privileged Role Administrator account.'
		try {
			Connect-MgGraph @params -ErrorAction Stop | Out-Null
		} catch {
			# An incompatible MSAL already in the session fails here, before any browser opens, with a type load
			# error that says nothing about the real cause. Add the diagnosis rather than let it surface raw.
			throw ("Microsoft Graph sign-in failed: {0}{1}" -f $_.Exception.Message, (Get-MsalConflictHint))
		}
	}
}

function Connect-ToExchangeOnlineForInstall {
	param([Parameter(Mandatory = $false)][string] $Organization)

	Assert-RequiredModule -Name ExchangeOnlineManagement

	$connection = Get-ConnectionInformation -ErrorAction SilentlyContinue | Where-Object { $_.State -eq 'Connected' } | Select-Object -First 1
	if ($connection) {
		Write-Host 'Current Exchange Online login:'
		Write-Host "  User:         $($connection.UserPrincipalName)"
		Write-Host "  Organization: $($connection.Organization)"
		return
	}

	$connectParams = @{ ShowBanner = $false }
	if (-not [string]::IsNullOrWhiteSpace($Organization)) {
		$connectParams['Organization'] = $Organization
	}

	Write-InstallLog 'Connecting to Exchange Online. Use an account that can create service principals and management role assignments.'
	Connect-ExchangeOnline @connectParams | Out-Null
}

function Get-ExchangeServicePrincipalByAppId {
	param([Parameter(Mandatory = $true)][string] $AppId)

	try {
		$servicePrincipal = Get-ServicePrincipal -AppId $AppId -ErrorAction Stop
		return $servicePrincipal | Select-Object -First 1
	} catch {
		$allServicePrincipals = @(Get-ServicePrincipal -ErrorAction Stop)
		return $allServicePrincipals | Where-Object { $_.AppId -eq $AppId } | Select-Object -First 1
	}
}

function Set-DriftExchangeManagementRoleAssignment {
	param(
		[Parameter(Mandatory = $true)][string] $Role,
		[Parameter(Mandatory = $true)][string] $ExchangeAppName,
		[Parameter(Mandatory = $false)][string] $CustomResourceScope
	)

	$existingAssignments = @(Get-ManagementRoleAssignment -Role $Role -ErrorAction Stop | Where-Object {
		$_.RoleAssigneeName -eq $ExchangeAppName -or
		$_.RoleAssignee -eq $ExchangeAppName -or
		$_.App -eq $ExchangeAppName
	})

	if ($existingAssignments.Count -gt 0) {
		Write-InstallLog "Exchange Online role assignment '$Role' already exists for '$ExchangeAppName'."
		return
	}

	$assignmentParams = @{ Role = $Role; App = $ExchangeAppName }
	if (-not [string]::IsNullOrWhiteSpace($CustomResourceScope)) {
		$assignmentParams['CustomResourceScope'] = $CustomResourceScope
		Write-InstallLog "Assigning Exchange Online role '$Role' to '$ExchangeAppName' scoped to '$CustomResourceScope'."
	} else {
		Write-InstallLog "Assigning Exchange Online role '$Role' to '$ExchangeAppName'."
	}

	New-ManagementRoleAssignment @assignmentParams | Out-Null
}

function Set-DriftMailSendScope {
	param([Parameter(Mandatory = $true)][string] $SenderMailbox)

	$mailbox = Get-Mailbox -Identity $SenderMailbox -ErrorAction SilentlyContinue | Select-Object -First 1
	if (-not $mailbox) {
		Write-InstallLog "Mail sender '$SenderMailbox' could not be resolved to an Exchange Online mailbox. The 'Application Mail.Send' role will be assigned without a mailbox scope (the managed identity will be able to send as any mailbox in the tenant)." -Level Warning
		return $null
	}

	$primarySmtp = [string] $mailbox.PrimarySmtpAddress
	$scopeName = 'DriftMaester-MailSend'
	$recipientFilter = "PrimarySmtpAddress -eq '$primarySmtp'"

	$existingScope = Get-ManagementScope -Identity $scopeName -ErrorAction SilentlyContinue | Select-Object -First 1
	if ($existingScope) {
		if ([string] $existingScope.RecipientFilter -ne $recipientFilter) {
			Write-InstallLog "Updating management scope '$scopeName' to target mailbox '$primarySmtp'."
			Set-ManagementScope -Identity $scopeName -RecipientRestrictionFilter $recipientFilter | Out-Null
		} else {
			Write-InstallLog "Management scope '$scopeName' already targets mailbox '$primarySmtp'."
		}
	} else {
		Write-InstallLog "Creating management scope '$scopeName' restricted to mailbox '$primarySmtp'."
		New-ManagementScope -Name $scopeName -RecipientRestrictionFilter $recipientFilter | Out-Null
	}

	return $scopeName
}

function Set-DriftExchangeOnlineRbac {
	param(
		[Parameter(Mandatory = $true)][string] $ManagedIdentityClientId,
		[Parameter(Mandatory = $true)][string] $ManagedIdentityObjectId,
		[Parameter(Mandatory = $true)][string] $DisplayName,
		[Parameter(Mandatory = $false)][string] $Organization,
		[Parameter(Mandatory = $false)][string] $SenderMailbox
	)

	try {
		Connect-ToExchangeOnlineForInstall -Organization $Organization

		foreach ($commandName in @('Get-ServicePrincipal', 'New-ServicePrincipal', 'Get-ManagementRoleAssignment', 'New-ManagementRoleAssignment', 'Get-ManagementScope', 'New-ManagementScope', 'Set-ManagementScope', 'Get-Mailbox')) {
			if (-not (Get-Command -Name $commandName -ErrorAction SilentlyContinue)) {
				throw "Exchange Online command '$commandName' is not available after connecting. Update ExchangeOnlineManagement and make sure the account has Exchange RBAC permissions."
			}
		}

		$servicePrincipal = Get-ExchangeServicePrincipalByAppId -AppId $ManagedIdentityClientId
		if ($servicePrincipal) {
			Write-InstallLog "Exchange Online service principal for app '$ManagedIdentityClientId' already exists."
			$exchangeAppName = if ($servicePrincipal.DisplayName) { [string] $servicePrincipal.DisplayName } elseif ($servicePrincipal.Name) { [string] $servicePrincipal.Name } elseif ($servicePrincipal.Identity) { [string] $servicePrincipal.Identity } else { $DisplayName }
		} else {
			Write-InstallLog "Creating Exchange Online service principal '$DisplayName'."
			New-ServicePrincipal -AppId $ManagedIdentityClientId -ObjectId $ManagedIdentityObjectId -DisplayName $DisplayName | Out-Null
			$exchangeAppName = $DisplayName
		}

		# Read-only Exchange configuration access used by the Maester checks.
		Set-DriftExchangeManagementRoleAssignment -Role 'View-Only Configuration' -ExchangeAppName $exchangeAppName

		# Least-privilege mail sending: grant 'Application Mail.Send' through Exchange RBAC instead of the
		# tenant-wide Graph Mail.Send application permission, scoped to the sender mailbox when it can be resolved.
		$mailSendScope = $null
		if (-not [string]::IsNullOrWhiteSpace($SenderMailbox)) {
			$mailSendScope = Set-DriftMailSendScope -SenderMailbox $SenderMailbox
		} else {
			Write-InstallLog "No mail sender mailbox was provided. The 'Application Mail.Send' role will be assigned without a mailbox scope (the managed identity will be able to send as any mailbox in the tenant)." -Level Warning
		}

		Set-DriftExchangeManagementRoleAssignment -Role 'Application Mail.Send' -ExchangeAppName $exchangeAppName -CustomResourceScope $mailSendScope
		return $true
	} catch {
		Write-InstallLog "Exchange Online RBAC mail sending could not be configured: $($_.Exception.Message)" -Level Warning
		return $false
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

function Get-InitialTenantDomainFromGraph {
	try {
		$response = Invoke-MgGraphRequest -Method GET -Uri 'https://graph.microsoft.com/v1.0/domains?$select=id,isInitial'
		$initialDomain = @($response.value | Where-Object { $_.isInitial } | Select-Object -First 1).id
		if (-not [string]::IsNullOrWhiteSpace($initialDomain)) {
			return [string] $initialDomain
		}
	} catch {
		Write-InstallLog "Could not resolve tenant initial domain from Graph for Exchange Online connection: $($_.Exception.Message)" -Level Warning
	}

	return $null
}

function Set-DriftAppRoleAssignment {
	param(
		[Parameter(Mandatory = $true)][object] $ManagedIdentityServicePrincipal,
		[Parameter(Mandatory = $true)][object] $ResourceServicePrincipal,
		[Parameter(Mandatory = $true)][string] $AppRoleValue
	)

	$appRole = @($ResourceServicePrincipal.appRoles | Where-Object { $_.value -eq $AppRoleValue -and $_.allowedMemberTypes -contains 'Application' } | Select-Object -First 1)
	if (-not $appRole) {
		Write-InstallLog "App role '$AppRoleValue' was not found on '$($ResourceServicePrincipal.displayName)'. Skipping." -Level Warning
		return
	}

	$existingUri = "https://graph.microsoft.com/v1.0/servicePrincipals/$($ManagedIdentityServicePrincipal.id)/appRoleAssignments"
	$existingAssignments = @(Invoke-GraphRequestAllPages -Uri $existingUri)
	$matchingAssignment = $existingAssignments | Where-Object { $_.appRoleId -eq $appRole.id -and $_.resourceId -eq $ResourceServicePrincipal.id } | Select-Object -First 1
	if ($matchingAssignment) {
		Write-InstallLog "API permission '$AppRoleValue' already assigned on '$($ResourceServicePrincipal.displayName)'."
		return
	}

	Write-InstallLog "Assigning API permission '$AppRoleValue' on '$($ResourceServicePrincipal.displayName)'."
	$body = @{
		principalId = $ManagedIdentityServicePrincipal.id
		resourceId  = $ResourceServicePrincipal.id
		appRoleId   = $appRole.id
	}
	Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/v1.0/servicePrincipals/$($ManagedIdentityServicePrincipal.id)/appRoleAssignments" -Body ($body | ConvertTo-Json) -ContentType 'application/json' | Out-Null
}

function Set-DriftGraphMailSendFallback {
	param(
		[Parameter(Mandatory = $true)][string] $ManagedIdentityClientId,
		[Parameter(Mandatory = $true)][bool] $Enabled
	)

	$managedIdentityServicePrincipal = Get-ServicePrincipalByAppId -AppId $ManagedIdentityClientId
	$graphServicePrincipal = Get-ServicePrincipalByAppId -AppId $script:GraphAppId

	$appRole = @($graphServicePrincipal.appRoles | Where-Object { $_.value -eq 'Mail.Send' -and $_.allowedMemberTypes -contains 'Application' } | Select-Object -First 1)
	if (-not $appRole) {
		Write-InstallLog "Graph application role 'Mail.Send' was not found on '$($graphServicePrincipal.displayName)'." -Level Warning
		return
	}

	$assignmentsUri = "https://graph.microsoft.com/v1.0/servicePrincipals/$($managedIdentityServicePrincipal.id)/appRoleAssignments"
	$existingAssignment = @(Invoke-GraphRequestAllPages -Uri $assignmentsUri) | Where-Object { $_.appRoleId -eq $appRole.id -and $_.resourceId -eq $graphServicePrincipal.id } | Select-Object -First 1

	if ($Enabled) {
		if ($existingAssignment) {
			Write-InstallLog "Graph 'Mail.Send' fallback permission is already assigned to the managed identity."
			return
		}

		Write-InstallLog "Assigning the broad Graph 'Mail.Send' application permission as a fallback for mail sending." -Level Warning
		$body = @{
			principalId = $managedIdentityServicePrincipal.id
			resourceId  = $graphServicePrincipal.id
			appRoleId   = $appRole.id
		}
		Invoke-MgGraphRequest -Method POST -Uri $assignmentsUri -Body ($body | ConvertTo-Json) -ContentType 'application/json' | Out-Null
	} else {
		if (-not $existingAssignment) {
			return
		}

		Write-InstallLog "Removing the broad Graph 'Mail.Send' permission because Exchange RBAC mail sending is now configured."
		Invoke-MgGraphRequest -Method DELETE -Uri "$assignmentsUri/$($existingAssignment.id)" | Out-Null
	}
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
			Write-InstallLog "Directory role '$RoleDisplayName' was not found. Skipping." -Level Warning
			return
		}

		Write-InstallLog "Activating directory role '$RoleDisplayName'."
		Invoke-MgGraphRequest -Method POST -Uri 'https://graph.microsoft.com/v1.0/directoryRoles' -Body (@{ roleTemplateId = $template.id } | ConvertTo-Json) -ContentType 'application/json' | Out-Null
		$role = @(Invoke-GraphRequestAllPages -Uri "https://graph.microsoft.com/v1.0/directoryRoles?`$filter=$roleFilter") | Select-Object -First 1
	}

	if (-not $role) {
		Write-InstallLog "Directory role '$RoleDisplayName' could not be activated. Skipping." -Level Warning
		return
	}

	$membersUri = "https://graph.microsoft.com/v1.0/directoryRoles/$($role.id)/members?`$select=id"
	$members = @(Invoke-GraphRequestAllPages -Uri $membersUri)
	$matchingMember = $members | Where-Object { $_.id -eq $ManagedIdentityServicePrincipalId } | Select-Object -First 1
	if ($matchingMember) {
		Write-InstallLog "Directory role '$RoleDisplayName' already assigned to the managed identity."
		return
	}

	Write-InstallLog "Assigning directory role '$RoleDisplayName' to the managed identity."
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
			Write-InstallLog "Directory role '$RoleDisplayName' already assigned to the managed identity."
			return
		}

		throw $postError
	}
}

function Set-DriftManagedIdentityApiPermissions {
	param(
		[Parameter(Mandatory = $true)][string] $ManagedIdentityClientId,
		[Parameter(Mandatory = $false)][string] $RequestedTenantId,
		[Parameter(Mandatory = $false)][string] $DevOpsOrganization
	)

	Connect-ToGraphForInstall -RequestedTenantId $RequestedTenantId

	$managedIdentityServicePrincipal = Get-ServicePrincipalByAppId -AppId $ManagedIdentityClientId
	$graphServicePrincipal = Get-ServicePrincipalByAppId -AppId $script:GraphAppId
	$exchangeServicePrincipal = Get-ServicePrincipalByAppId -AppId $script:ExchangeOnlineAppId

	foreach ($permission in ($RequiredGraphApplicationPermissions | Sort-Object -Unique)) {
		Set-DriftAppRoleAssignment -ManagedIdentityServicePrincipal $managedIdentityServicePrincipal -ResourceServicePrincipal $graphServicePrincipal -AppRoleValue $permission
	}

	Set-DriftAppRoleAssignment -ManagedIdentityServicePrincipal $managedIdentityServicePrincipal -ResourceServicePrincipal $exchangeServicePrincipal -AppRoleValue 'Exchange.ManageAsApp'

	# SharePoint Online is a separate resource from Graph, so the SPO tests need app roles on its own service principal.
	try {
		$sharePointServicePrincipal = Get-ServicePrincipalByAppId -AppId $script:SharePointOnlineAppId
		foreach ($permission in ($RequiredSharePointApplicationPermissions | Sort-Object -Unique)) {
			Set-DriftAppRoleAssignment -ManagedIdentityServicePrincipal $managedIdentityServicePrincipal -ResourceServicePrincipal $sharePointServicePrincipal -AppRoleValue $permission
		}
	} catch {
		Write-InstallLog "Could not assign SharePoint Online application permissions: $($_.Exception.Message). The Maester SharePoint Online tests will be skipped until this is resolved." -Level Warning
	}

	foreach ($roleName in $DirectoryRolesForManagedIdentity) {
		if ($roleName -eq 'Azure DevOps Administrator' -and [string]::IsNullOrWhiteSpace($DevOpsOrganization)) {
			Write-InstallLog "Skipping directory role '$roleName' because no Azure DevOps organization was configured."
			continue
		}
		Set-DriftDirectoryRoleAssignment -ManagedIdentityServicePrincipalId $managedIdentityServicePrincipal.id -RoleDisplayName $roleName
	}
}

function Get-InvokeParameters {
	param(
		[Parameter(Mandatory = $true)][string] $Recipients,
		[Parameter(Mandatory = $false)][string] $SenderUserId,
		[Parameter(Mandatory = $false)][string] $DevOpsOrg,
		[Parameter(Mandatory = $false)][string] $TargetTenantId,
		[Parameter(Mandatory = $false)][string] $SubjectPrefix,
		[Parameter(Mandatory = $false)][ValidateSet('auto', 'attach', 'link')][string] $ReportDelivery = 'auto',
		[Parameter(Mandatory = $false)][ValidateSet('none', 'low', 'medium', 'high', 'critical')][string] $AlertMinimumSeverity = 'none',
		[Parameter(Mandatory = $false)][string] $AlertRecipients,
		[Parameter(Mandatory = $false)][string] $TeamsWebhookUrl,
		[Parameter(Mandatory = $false)][ValidateRange(30, 3650)][int] $RetentionDays = 180,
		[Parameter(Mandatory = $false)][bool] $AlwaysSendReport = $false,
		[Parameter(Mandatory = $false)][bool] $IncludeCopilotAndDataverse = $false,
		[Parameter(Mandatory = $false)][bool] $IncludeMaesterReport = $false
	)

	$parameters = @{
		reportrecipient = $Recipients
		reportdelivery = $ReportDelivery
		alertminimumseverity = $AlertMinimumSeverity
		retentiondays = $RetentionDays
		alwayssendreport = $AlwaysSendReport
		includeCopilotAndDataverse = $IncludeCopilotAndDataverse
		includeMaesterReport = $IncludeMaesterReport
	}
	if (-not [string]::IsNullOrWhiteSpace($SubjectPrefix)) { $parameters['mailsubjectprefix'] = $SubjectPrefix } else{ $parameters['mailsubjectprefix'] = 'DriftMaester Report' }
	if (-not [string]::IsNullOrWhiteSpace($SenderUserId)) { $parameters['mailsenderuserid'] = $SenderUserId }
	if (-not [string]::IsNullOrWhiteSpace($DevOpsOrg)) { $parameters['devopsorganization'] = $DevOpsOrg }
	if (-not [string]::IsNullOrWhiteSpace($TargetTenantId)) { $parameters['tenantid'] = $TargetTenantId }
	if (-not [string]::IsNullOrWhiteSpace($AlertRecipients)) { $parameters['alertrecipient'] = $AlertRecipients }
	if (-not [string]::IsNullOrWhiteSpace($TeamsWebhookUrl)) { $parameters['teamswebhookurl'] = $TeamsWebhookUrl }

	return $parameters
}

# GUI Functions
function Show-DriftMaesterGui {
	param(
		[Parameter(Mandatory = $false)][string] $PrefilledSubscription,
		[Parameter(Mandatory = $false)][string] $PrefilledResourceGroup,
		[Parameter(Mandatory = $false)][string] $PrefilledLocation,
		[Parameter(Mandatory = $false)][string] $PrefilledFrequency,
		[Parameter(Mandatory = $false)][string] $PrefilledDayOfWeek,
		[Parameter(Mandatory = $false)][int] $PrefilledDayOfMonth = 1,
		[Parameter(Mandatory = $false)][string] $PrefilledTimeOfDay,
		[Parameter(Mandatory = $false)][string] $PrefilledRecipients,
		[Parameter(Mandatory = $false)][string] $PrefilledAlertRecipients,
		[Parameter(Mandatory = $false)][string] $PrefilledSenderUserId,
		[Parameter(Mandatory = $false)][string] $PrefilledDevOpsOrg,
		[Parameter(Mandatory = $false)][string] $PrefilledTenantId,
		[Parameter(Mandatory = $false)][string] $PrefilledMailSubject,
		[Parameter(Mandatory = $false)][ValidateSet('auto', 'attach', 'link')][string] $PrefilledReportDelivery = 'auto',
		[Parameter(Mandatory = $false)][ValidateSet('none', 'low', 'medium', 'high', 'critical')][string] $PrefilledAlertMinimumSeverity = 'none',
		[Parameter(Mandatory = $false)][string] $PrefilledTeamsWebhookUrl,
		[Parameter(Mandatory = $false)][int] $PrefilledRetentionDays = 180,
		[Parameter(Mandatory = $false)][bool] $PrefilledRunNow = $false,
		[Parameter(Mandatory = $false)][string] $PrefilledTimeZone,
		[Parameter(Mandatory = $false)][bool] $PrefilledAlwaysSendReport = $false,
		[Parameter(Mandatory = $false)][bool] $PrefilledIncludeCopilotAndDataverse = $false,
		[Parameter(Mandatory = $false)][bool] $PrefilledIncludeMaesterReport = $false
	)

	if (-not [System.Runtime.InteropServices.RuntimeInformation]::IsOSPlatform([System.Runtime.InteropServices.OSPlatform]::Windows)) {
		throw 'GUI mode is only supported on Windows. On non-Windows platforms, provide all required parameters or use the console prompts.'
	}

	$subscriptions = @(Get-AzSubscription | Sort-Object Name)
	if ($subscriptions.Count -eq 0) {
		throw 'No Azure subscriptions are visible for the current account.'
	}

	$currentAzContext = Get-AzContext -ErrorAction SilentlyContinue
	$defaultRecipient = if ($currentAzContext -and -not [string]::IsNullOrWhiteSpace([string] $currentAzContext.Account.Id)) { [string] $currentAzContext.Account.Id } else { '' }
	$prefilled = @{
		subscription = $PrefilledSubscription
		resourceGroup = if ([string]::IsNullOrWhiteSpace($PrefilledResourceGroup)) { 'driftmaester' } else { $PrefilledResourceGroup }
		location = if ([string]::IsNullOrWhiteSpace($PrefilledLocation)) { $script:Location } else { $PrefilledLocation }
		frequency = if ([string]::IsNullOrWhiteSpace($PrefilledFrequency)) { 'daily' } else { $PrefilledFrequency }
		dayOfWeek = if ([string]::IsNullOrWhiteSpace($PrefilledDayOfWeek)) { 'Monday' } else { $PrefilledDayOfWeek }
		dayOfMonth = if ($PrefilledDayOfMonth -eq 0) { 1 } else { $PrefilledDayOfMonth }
		timeOfDay = if ([string]::IsNullOrWhiteSpace($PrefilledTimeOfDay)) { '02:00' } else { $PrefilledTimeOfDay }
		timeZone = Resolve-DriftTimeZoneId -TimeZoneId $PrefilledTimeZone
		recipients = if ([string]::IsNullOrWhiteSpace($PrefilledRecipients)) { $defaultRecipient } else { $PrefilledRecipients }
		alertRecipients = $PrefilledAlertRecipients
		senderUserId = $PrefilledSenderUserId
		devopsOrg = $PrefilledDevOpsOrg
		tenantId = $PrefilledTenantId
		mailSubject = if ([string]::IsNullOrWhiteSpace($PrefilledMailSubject)) { 'DriftMaester Report' } else { $PrefilledMailSubject }
		reportDelivery = $PrefilledReportDelivery
		alertMinimumSeverity = $PrefilledAlertMinimumSeverity
		teamsWebhookUrl = $PrefilledTeamsWebhookUrl
		retentionDays = $PrefilledRetentionDays
		runNow = $PrefilledRunNow
		alwaysSendReport = $PrefilledAlwaysSendReport
		includeCopilotAndDataverse = $PrefilledIncludeCopilotAndDataverse
		includeMaesterReport = $PrefilledIncludeMaesterReport
	}

	$uiScript = {
		param([hashtable] $Prefilled, [object[]] $Subscriptions)

		Add-Type -AssemblyName System.Windows.Forms
		Add-Type -AssemblyName System.Drawing
		[System.Windows.Forms.Application]::EnableVisualStyles()

		$form = [System.Windows.Forms.Form]::new()
		$form.Text = 'DriftMaester Installer'
		$form.StartPosition = 'CenterScreen'
		$form.AutoScaleMode = [System.Windows.Forms.AutoScaleMode]::Dpi
		$form.Size = [System.Drawing.Size]::new(820, 640)
		$form.MinimumSize = [System.Drawing.Size]::new(780, 600)
		$form.Font = [System.Drawing.Font]::new('Segoe UI', 9)
		$form.MaximizeBox = $true
		$form.BackColor = [System.Drawing.Color]::White

		$tooltip = [System.Windows.Forms.ToolTip]::new()
		$tooltip.AutoPopDelay = 20000
		$tooltip.InitialDelay = 300
		$tooltip.ReshowDelay = 100

		$accent = [System.Drawing.Color]::FromArgb(37, 99, 235)
		$muted = [System.Drawing.Color]::FromArgb(100, 116, 139)
		$fieldWidth = 640

		$root = [System.Windows.Forms.TableLayoutPanel]::new()
		$root.Dock = 'Fill'
		$root.ColumnCount = 1
		$root.RowCount = 3
		[void] $root.RowStyles.Add([System.Windows.Forms.RowStyle]::new([System.Windows.Forms.SizeType]::Absolute, 86))
		[void] $root.RowStyles.Add([System.Windows.Forms.RowStyle]::new([System.Windows.Forms.SizeType]::Percent, 100))
		[void] $root.RowStyles.Add([System.Windows.Forms.RowStyle]::new([System.Windows.Forms.SizeType]::Absolute, 66))
		$form.Controls.Add($root)

		$headerPanel = [System.Windows.Forms.Panel]::new()
		$headerPanel.Dock = 'Fill'
		$headerPanel.BackColor = [System.Drawing.Color]::FromArgb(243, 246, 251)
		$root.Controls.Add($headerPanel, 0, 0)

		$header = [System.Windows.Forms.Label]::new()
		$header.Text = 'Configure DriftMaester'
		$header.Font = [System.Drawing.Font]::new('Segoe UI', 14, [System.Drawing.FontStyle]::Bold)
		$header.ForeColor = [System.Drawing.Color]::FromArgb(23, 32, 51)
		$header.Location = [System.Drawing.Point]::new(24, 14)
		$header.AutoSize = $true
		$headerPanel.Controls.Add($header)

		$stepLabel = [System.Windows.Forms.Label]::new()
		$stepLabel.Font = [System.Drawing.Font]::new('Segoe UI', 9)
		$stepLabel.ForeColor = $accent
		$stepLabel.Location = [System.Drawing.Point]::new(26, 52)
		$stepLabel.AutoSize = $true
		$headerPanel.Controls.Add($stepLabel)

		$contentPanel = [System.Windows.Forms.Panel]::new()
		$contentPanel.Dock = 'Fill'
		$root.Controls.Add($contentPanel, 0, 1)

		$footerPanel = [System.Windows.Forms.Panel]::new()
		$footerPanel.Dock = 'Fill'
		$footerPanel.BackColor = [System.Drawing.Color]::FromArgb(243, 246, 251)
		$root.Controls.Add($footerPanel, 0, 2)

		$buttonRow = [System.Windows.Forms.FlowLayoutPanel]::new()
		$buttonRow.Dock = 'Fill'
		$buttonRow.FlowDirection = 'RightToLeft'
		$buttonRow.Padding = [System.Windows.Forms.Padding]::new(0, 14, 18, 0)
		$footerPanel.Controls.Add($buttonRow)

		$installButton = [System.Windows.Forms.Button]::new()
		$installButton.Text = 'Install'
		$installButton.Size = [System.Drawing.Size]::new(150, 34)
		$installButton.BackColor = $accent
		$installButton.ForeColor = [System.Drawing.Color]::White
		$installButton.FlatStyle = 'Flat'
		$installButton.FlatAppearance.BorderSize = 0
		$buttonRow.Controls.Add($installButton)

		$nextButton = [System.Windows.Forms.Button]::new()
		$nextButton.Text = 'Next >'
		$nextButton.Size = [System.Drawing.Size]::new(110, 34)
		$nextButton.BackColor = $accent
		$nextButton.ForeColor = [System.Drawing.Color]::White
		$nextButton.FlatStyle = 'Flat'
		$nextButton.FlatAppearance.BorderSize = 0
		$buttonRow.Controls.Add($nextButton)

		$backButton = [System.Windows.Forms.Button]::new()
		$backButton.Text = '< Back'
		$backButton.Size = [System.Drawing.Size]::new(110, 34)
		$backButton.FlatStyle = 'Flat'
		$buttonRow.Controls.Add($backButton)

		$cancelButton = [System.Windows.Forms.Button]::new()
		$cancelButton.Text = 'Cancel'
		$cancelButton.Size = [System.Drawing.Size]::new(110, 34)
		$cancelButton.FlatStyle = 'Flat'
		$cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
		$buttonRow.Controls.Add($cancelButton)
		$form.CancelButton = $cancelButton

		function New-StepPanel {
			$panel = [System.Windows.Forms.FlowLayoutPanel]::new()
			$panel.Dock = 'Fill'
			$panel.FlowDirection = 'TopDown'
			$panel.WrapContents = $false
			$panel.AutoScroll = $true
			$panel.Padding = [System.Windows.Forms.Padding]::new(26, 14, 26, 14)
			$panel.Visible = $false
			$contentPanel.Controls.Add($panel)
			return $panel
		}

		function Add-Field {
			param(
				[System.Windows.Forms.FlowLayoutPanel] $Panel,
				[string] $Title,
				[System.Windows.Forms.Control] $Control,
				[string] $Description
			)

			$group = [System.Windows.Forms.Panel]::new()
			$group.Width = $fieldWidth
			$group.Margin = [System.Windows.Forms.Padding]::new(0, 0, 0, 10)

			$controlTop = 0
			if (-not [string]::IsNullOrWhiteSpace($Title)) {
				$titleLabel = [System.Windows.Forms.Label]::new()
				$titleLabel.Text = $Title
				$titleLabel.Font = [System.Drawing.Font]::new('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
				$titleLabel.ForeColor = [System.Drawing.Color]::FromArgb(23, 32, 51)
				$titleLabel.Location = [System.Drawing.Point]::new(0, 0)
				$titleLabel.Size = [System.Drawing.Size]::new($fieldWidth, 20)
				$group.Controls.Add($titleLabel)
				$controlTop = 22
			}

			$Control.Location = [System.Drawing.Point]::new(0, $controlTop)
			$group.Controls.Add($Control)

			$descTop = $controlTop + $Control.Height + 3
			if (-not [string]::IsNullOrWhiteSpace($Description)) {
				$descLabel = [System.Windows.Forms.Label]::new()
				$descLabel.Text = $Description
				$descLabel.Font = [System.Drawing.Font]::new('Segoe UI', 8)
				$descLabel.ForeColor = $muted
				$descLabel.Location = [System.Drawing.Point]::new(0, $descTop)
				$descLabel.Size = [System.Drawing.Size]::new($fieldWidth, 30)
				$group.Controls.Add($descLabel)
				$tooltip.SetToolTip($Control, $Description)
				$group.Height = $descTop + 32
			} else {
				$group.Height = $descTop
			}

			[void] $Panel.Controls.Add($group)
			return $group
		}

		function New-Combo {
			param([int] $Width = 320, [switch] $ReadOnly)
			$combo = [System.Windows.Forms.ComboBox]::new()
			if ($ReadOnly) { $combo.DropDownStyle = 'DropDownList' }
			$combo.Size = [System.Drawing.Size]::new($Width, 26)
			return $combo
		}

		function New-Text {
			param([string] $Text, [int] $Width = 640)
			$box = [System.Windows.Forms.TextBox]::new()
			$box.Size = [System.Drawing.Size]::new($Width, 26)
			$box.Text = [string] $Text
			return $box
		}

		function New-Check {
			param([string] $Text, [bool] $Checked)
			$check = [System.Windows.Forms.CheckBox]::new()
			$check.Text = $Text
			$check.Size = [System.Drawing.Size]::new($fieldWidth, 24)
			$check.Checked = $Checked
			return $check
		}

		$stepTitles = @('Azure target', 'Schedule', 'Reporting', 'Options', 'Review & install')
		$steps = @()

		# Step 1 - Azure target
		$step0 = New-StepPanel
		$steps += $step0

		$subscriptionCombo = New-Combo -Width 640 -ReadOnly
		$subscriptionCombo.DisplayMember = 'DisplayName'
		$selectedSubscriptionIndex = -1
		for ($index = 0; $index -lt $Subscriptions.Count; $index++) {
			$subscription = $Subscriptions[$index]
			$item = [PSCustomObject]@{
				DisplayName = ('{0} ({1})' -f $subscription.Name, $subscription.Id)
				Id = [string] $subscription.Id
				Name = [string] $subscription.Name
				TenantId = [string] $subscription.TenantId
			}
			[void] $subscriptionCombo.Items.Add($item)
			if ($Prefilled.subscription -and ($Prefilled.subscription -eq $item.Id -or $Prefilled.subscription -eq $item.Name)) {
				$selectedSubscriptionIndex = $index
			}
		}
		if ($subscriptionCombo.Items.Count -gt 0) {
			$subscriptionCombo.SelectedIndex = $(if ($selectedSubscriptionIndex -ge 0) { $selectedSubscriptionIndex } else { 0 })
		}
		Add-Field -Panel $step0 -Title 'Azure subscription *' -Control $subscriptionCombo -Description 'The subscription that will host the DriftMaester resource group, automation account and storage.' | Out-Null

		$resourceGroupBox = New-Text -Text $Prefilled.resourceGroup -Width 640
		Add-Field -Panel $step0 -Title 'Resource group *' -Control $resourceGroupBox -Description 'Created if it does not exist. All DriftMaester resources live here so you can remove everything in one go.' | Out-Null

		$locationCombo = New-Combo -Width 320
		$locationOptions = @()
		try {
			$locationOptions = @((Get-AzLocation | Sort-Object Location | Select-Object -ExpandProperty Location -Unique))
		} catch {
			$locationOptions = @()
		}
		if ($locationOptions.Count -eq 0) {
			$locationOptions = @('westeurope', 'northeurope', 'uksouth', 'eastus', 'westus2')
		}
		foreach ($locationName in $locationOptions) { [void] $locationCombo.Items.Add($locationName) }
		$locationCombo.Text = [string] $Prefilled.location
		Add-Field -Panel $step0 -Title 'Location' -Control $locationCombo -Description 'Azure region for the resource group and automation account. Pick one close to your admins.' | Out-Null

		# Step 2 - Schedule
		$step1 = New-StepPanel
		$steps += $step1

		$frequencyCombo = New-Combo -Width 220 -ReadOnly
		foreach ($frequencyName in @('daily', 'weekly', 'monthly')) { [void] $frequencyCombo.Items.Add($frequencyName) }
		$frequencyCombo.SelectedItem = [string] $Prefilled.frequency
		if (-not $frequencyCombo.SelectedItem) { $frequencyCombo.SelectedIndex = 0 }
		Add-Field -Panel $step1 -Title 'Frequency *' -Control $frequencyCombo -Description 'How often the drift scan runs.' | Out-Null

		$dayOfWeekCombo = New-Combo -Width 220 -ReadOnly
		foreach ($dayName in @('Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday')) { [void] $dayOfWeekCombo.Items.Add($dayName) }
		$dayOfWeekCombo.SelectedItem = [string] $Prefilled.dayOfWeek
		if (-not $dayOfWeekCombo.SelectedItem) { $dayOfWeekCombo.SelectedItem = 'Monday' }
		$dayOfWeekField = Add-Field -Panel $step1 -Title 'Day of week' -Control $dayOfWeekCombo -Description 'Only used for weekly schedules.'

		$dayOfMonthBox = New-Text -Text ([string] $Prefilled.dayOfMonth) -Width 120
		$dayOfMonthField = Add-Field -Panel $step1 -Title 'Day of month' -Control $dayOfMonthBox -Description '1-31, or -1 for the last day of the month. Only used for monthly schedules.'

		$timeBox = New-Text -Text $Prefilled.timeOfDay -Width 120
		Add-Field -Panel $step1 -Title 'Time of day (HH:mm) *' -Control $timeBox -Description 'Start time in 24-hour format, interpreted in the time zone below.' | Out-Null

		$timeZoneCombo = New-Combo -Width 640
		foreach ($timeZoneInfo in ([System.TimeZoneInfo]::GetSystemTimeZones() | Sort-Object Id)) { [void] $timeZoneCombo.Items.Add($timeZoneInfo.Id) }
		$timeZoneCombo.Text = [string] $Prefilled.timeZone
		Add-Field -Panel $step1 -Title 'Time zone *' -Control $timeZoneCombo -Description 'The schedule fires according to this time zone.' | Out-Null

		# Step 3 - Reporting
		$step2 = New-StepPanel
		$steps += $step2

		$recipientsBox = [System.Windows.Forms.TextBox]::new()
		$recipientsBox.Size = [System.Drawing.Size]::new(640, 56)
		$recipientsBox.Multiline = $true
		$recipientsBox.ScrollBars = 'Vertical'
		$recipientsBox.Text = [string] $Prefilled.recipients
		Add-Field -Panel $step2 -Title 'Report recipients *' -Control $recipientsBox -Description 'Comma-separated. Everyone listed here receives every report that is sent.' | Out-Null

		$alertRecipientsBox = New-Text -Text $Prefilled.alertRecipients -Width 640
		Add-Field -Panel $step2 -Title 'Alert recipients' -Control $alertRecipientsBox -Description 'Optional. When drift meets the alert severity below, the report goes to these people instead of the normal recipients.' | Out-Null

		$senderBox = New-Text -Text $Prefilled.senderUserId -Width 640
		Add-Field -Panel $step2 -Title 'Mail sender user id' -Control $senderBox -Description 'Optional mailbox or UPN used as the sender. Defaults to the first report recipient.' | Out-Null

		$mailSubjectBox = New-Text -Text $Prefilled.mailSubject -Width 640
		Add-Field -Panel $step2 -Title 'Mail subject prefix' -Control $mailSubjectBox -Description 'Prefix added to every email subject, handy for inbox rules or multi-tenant setups.' | Out-Null

		$reportDeliveryCombo = New-Combo -Width 220 -ReadOnly
		foreach ($delivery in @('auto', 'attach', 'link')) { [void] $reportDeliveryCombo.Items.Add($delivery) }
		$reportDeliveryCombo.SelectedItem = [string] $Prefilled.reportDelivery
		if (-not $reportDeliveryCombo.SelectedItem) { $reportDeliveryCombo.SelectedItem = 'auto' }
		Add-Field -Panel $step2 -Title 'Report delivery mode' -Control $reportDeliveryCombo -Description 'auto attaches when small and links when large. attach always attaches. link always points to blob storage.' | Out-Null

		$alertSeverityCombo = New-Combo -Width 220 -ReadOnly
		foreach ($severity in @('none', 'low', 'medium', 'high', 'critical')) { [void] $alertSeverityCombo.Items.Add($severity) }
		$alertSeverityCombo.SelectedItem = [string] $Prefilled.alertMinimumSeverity
		if (-not $alertSeverityCombo.SelectedItem) { $alertSeverityCombo.SelectedItem = 'none' }
		Add-Field -Panel $step2 -Title 'Alert minimum severity' -Control $alertSeverityCombo -Description 'Minimum severity of a drift change before the alert recipients are notified.' | Out-Null

		$teamsWebhookBox = New-Text -Text $Prefilled.teamsWebhookUrl -Width 640
		Add-Field -Panel $step2 -Title 'Teams webhook URL' -Control $teamsWebhookBox -Description 'Optional. Posts a short summary to a Teams channel using an incoming webhook.' | Out-Null

		# Step 4 - Options
		$step3 = New-StepPanel
		$steps += $step3

		$tenantBox = New-Text -Text $Prefilled.tenantId -Width 640
		Add-Field -Panel $step3 -Title 'Tenant id' -Control $tenantBox -Description 'Optional. Target tenant for token acquisition in multi-tenant scenarios. Auto-filled from the subscription.' | Out-Null

		$devOpsBox = New-Text -Text $Prefilled.devopsOrg -Width 640
		Add-Field -Panel $step3 -Title 'Azure DevOps organization' -Control $devOpsBox -Description 'Optional. Organization name to include Azure DevOps pipeline drift checks.' | Out-Null

		$retentionDaysBox = New-Text -Text ([string] $Prefilled.retentionDays) -Width 120
		Add-Field -Panel $step3 -Title 'Retention days' -Control $retentionDaysBox -Description 'How long stored results are kept before cleanup (30-3650 days).' | Out-Null

		$alwaysSendReportBox = New-Check -Text 'Always send report, even when no drift is detected' -Checked ([bool] $Prefilled.alwaysSendReport)
		Add-Field -Panel $step3 -Title '' -Control $alwaysSendReportBox -Description 'When off, a report is only sent on the first run or when drift is detected.' | Out-Null

		$runNowBox = New-Check -Text 'Run update + invoke immediately after install' -Checked ([bool] $Prefilled.runNow)
		Add-Field -Panel $step3 -Title '' -Control $runNowBox -Description 'Kicks off a first scan right after installation so you get a baseline report.' | Out-Null

		$includeCopilotAndDataverseBox = New-Check -Text 'Include Copilot, Power Platform, Dynamics, and Dataverse checks' -Checked ([bool] $Prefilled.includeCopilotAndDataverse)
		Add-Field -Panel $step3 -Title '' -Control $includeCopilotAndDataverseBox -Description 'Adds optional workload scanning. Requires the managed identity to have Dataverse access.' | Out-Null

		$includeMaesterReportBox = New-Check -Text 'Attach the full original Maester report (zipped) to the email' -Checked ([bool] $Prefilled.includeMaesterReport)
		Add-Field -Panel $step3 -Title '' -Control $includeMaesterReportBox -Description 'Heads up: in larger tenants this report can be big. Even zipped it may exceed mail size limits and get rejected.' | Out-Null

		# Step 5 - Review
		$step4 = New-StepPanel
		$steps += $step4

		$summaryBox = [System.Windows.Forms.TextBox]::new()
		$summaryBox.Size = [System.Drawing.Size]::new(640, 360)
		$summaryBox.Multiline = $true
		$summaryBox.ReadOnly = $true
		$summaryBox.ScrollBars = 'Vertical'
		$summaryBox.BackColor = [System.Drawing.Color]::FromArgb(248, 250, 252)
		$summaryBox.Font = [System.Drawing.Font]::new('Consolas', 9)
		Add-Field -Panel $step4 -Title 'Review your settings' -Control $summaryBox -Description 'Check everything below, then click Install. Use Back to change anything.' | Out-Null

		# Frequency-dependent visibility
		$updateFrequencyVisibility = {
			$frequencyValue = [string] $frequencyCombo.Text
			$dayOfWeekField.Visible = ($frequencyValue -eq 'weekly')
			$dayOfMonthField.Visible = ($frequencyValue -eq 'monthly')
		}
		$null = $frequencyCombo.Add_SelectedIndexChanged({ & $updateFrequencyVisibility })
		& $updateFrequencyVisibility

		# Subscription -> tenant autofill
		$subscriptionCombo.Add_SelectedIndexChanged({
			if ([string]::IsNullOrWhiteSpace($tenantBox.Text) -and $subscriptionCombo.SelectedItem) {
				$tenantBox.Text = [string] $subscriptionCombo.SelectedItem.TenantId
			}
		})
		if ([string]::IsNullOrWhiteSpace($tenantBox.Text) -and $subscriptionCombo.SelectedItem) {
			$tenantBox.Text = [string] $subscriptionCombo.SelectedItem.TenantId
		}

		$nav = @{ Step = 0 }

		$buildSummary = {
			$lines = @()
			$lines += 'Subscription:     ' + $(if ($subscriptionCombo.SelectedItem) { [string] $subscriptionCombo.SelectedItem.DisplayName } else { '' })
			$lines += 'Resource group:   ' + [string] $resourceGroupBox.Text.Trim()
			$lines += 'Location:         ' + [string] $locationCombo.Text.Trim()
			$lines += ''
			$lines += 'Frequency:        ' + [string] $frequencyCombo.Text.Trim()
			if ([string] $frequencyCombo.Text -eq 'weekly') { $lines += 'Day of week:      ' + [string] $dayOfWeekCombo.Text.Trim() }
			if ([string] $frequencyCombo.Text -eq 'monthly') { $lines += 'Day of month:     ' + [string] $dayOfMonthBox.Text.Trim() }
			$lines += 'Time of day:      ' + [string] $timeBox.Text.Trim()
			$lines += 'Time zone:        ' + [string] $timeZoneCombo.Text.Trim()
			$lines += ''
			$lines += 'Report to:        ' + (([string] $recipientsBox.Text) -replace '\s+', ' ').Trim()
			$lines += 'Alert to:         ' + [string] $alertRecipientsBox.Text.Trim()
			$lines += 'Sender:           ' + [string] $senderBox.Text.Trim()
			$lines += 'Subject prefix:   ' + [string] $mailSubjectBox.Text.Trim()
			$lines += 'Delivery mode:    ' + [string] $reportDeliveryCombo.Text.Trim()
			$lines += 'Alert severity:   ' + [string] $alertSeverityCombo.Text.Trim()
			$lines += 'Teams webhook:    ' + $(if ([string]::IsNullOrWhiteSpace($teamsWebhookBox.Text)) { '(none)' } else { 'configured' })
			$lines += ''
			$lines += 'Tenant id:        ' + [string] $tenantBox.Text.Trim()
			$lines += 'DevOps org:       ' + [string] $devOpsBox.Text.Trim()
			$lines += 'Retention days:   ' + [string] $retentionDaysBox.Text.Trim()
			$lines += 'Always send:      ' + [string] $alwaysSendReportBox.Checked
			$lines += 'Run now:          ' + [string] $runNowBox.Checked
			$lines += 'Copilot/Dataverse:' + ' ' + [string] $includeCopilotAndDataverseBox.Checked
			$lines += 'Attach Maester:   ' + [string] $includeMaesterReportBox.Checked
			$summaryBox.Text = ($lines -join [Environment]::NewLine)
		}

		$showStep = {
			param([int] $Index)
			for ($i = 0; $i -lt $steps.Count; $i++) { $steps[$i].Visible = ($i -eq $Index) }
			$nav.Step = $Index
			$stepLabel.Text = ('Step {0} of {1}   -   {2}' -f ($Index + 1), $steps.Count, $stepTitles[$Index])
			$backButton.Enabled = ($Index -gt 0)
			$isLast = ($Index -eq ($steps.Count - 1))
			$nextButton.Visible = -not $isLast
			$installButton.Visible = $isLast
			if ($isLast) { & $buildSummary }
		}

		$validateStep = {
			param([int] $Index)
			$issues = @()
			switch ($Index) {
				0 {
					if (-not $subscriptionCombo.SelectedItem) { $issues += 'Select an Azure subscription.' }
					if ([string]::IsNullOrWhiteSpace($resourceGroupBox.Text)) { $issues += 'Enter a resource group name.' }
				}
				1 {
					if ([string]::IsNullOrWhiteSpace($frequencyCombo.Text)) { $issues += 'Select a frequency.' }
					if (-not [regex]::IsMatch([string] $timeBox.Text.Trim(), '^([01]?\d|2[0-3]):[0-5]\d$')) { $issues += 'Time of day must be HH:mm in 24-hour format.' }
					if ([string] $frequencyCombo.Text -eq 'monthly') {
						$dayNumber = 0
						if (-not [int]::TryParse([string] $dayOfMonthBox.Text.Trim(), [ref] $dayNumber) -or (($dayNumber -lt 1 -or $dayNumber -gt 31) -and $dayNumber -ne -1)) { $issues += 'Day of month must be 1-31 or -1.' }
					}
					try { $null = [System.TimeZoneInfo]::FindSystemTimeZoneById([string] $timeZoneCombo.Text.Trim()) } catch { $issues += 'Select a valid time zone.' }
				}
				2 {
					if ([string]::IsNullOrWhiteSpace($recipientsBox.Text)) { $issues += 'Enter at least one report recipient.' }
					$allRecipientCandidates = @(([string] $recipientsBox.Text).Split(',') + ([string] $alertRecipientsBox.Text).Split(',')) | ForEach-Object { $_.Trim() } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
					$invalidEmails = @()
					foreach ($candidate in $allRecipientCandidates) {
						try { $null = [System.Net.Mail.MailAddress]::new($candidate) } catch { $invalidEmails += $candidate }
					}
					if ($invalidEmails.Count -gt 0) { $issues += ('Invalid e-mail address(es): {0}' -f ($invalidEmails -join ', ')) }
				}
				3 {
					$retentionValue = 0
					if (-not [int]::TryParse([string] $retentionDaysBox.Text.Trim(), [ref] $retentionValue) -or $retentionValue -lt 30 -or $retentionValue -gt 3650) { $issues += 'Retention days must be between 30 and 3650.' }
				}
			}
			return $issues
		}

		$nextButton.Add_Click({
			$issues = @(& $validateStep $nav.Step)
			if ($issues.Count -gt 0) {
				[System.Windows.Forms.MessageBox]::Show(($issues -join [Environment]::NewLine), 'DriftMaester Installer', 'OK', 'Warning') | Out-Null
				return
			}
			& $showStep ([Math]::Min($nav.Step + 1, $steps.Count - 1))
		})

		$backButton.Add_Click({ & $showStep ([Math]::Max($nav.Step - 1, 0)) })

		$installButton.Add_Click({
			$allIssues = @()
			foreach ($stepIndex in 0, 1, 2, 3) { $allIssues += @(& $validateStep $stepIndex) }
			if ($allIssues.Count -gt 0) {
				[System.Windows.Forms.MessageBox]::Show(($allIssues -join [Environment]::NewLine), 'DriftMaester Installer Validation', 'OK', 'Warning') | Out-Null
				return
			}

			$form.Tag = @{
				subscription = [string] $subscriptionCombo.SelectedItem.Id
				resourceGroup = [string] $resourceGroupBox.Text.Trim()
				location = [string] $locationCombo.Text.Trim()
				frequency = [string] $frequencyCombo.Text.Trim()
				dayOfWeek = [string] $dayOfWeekCombo.Text.Trim()
				dayOfMonth = [int] $dayOfMonthBox.Text.Trim()
				timeOfDay = [string] $timeBox.Text.Trim()
				timeZone = [string] $timeZoneCombo.Text.Trim()
				recipients = [string] $recipientsBox.Text.Trim()
				alertRecipients = [string] $alertRecipientsBox.Text.Trim()
				senderUserId = [string] $senderBox.Text.Trim()
				devopsOrg = [string] $devOpsBox.Text.Trim()
				tenantId = [string] $tenantBox.Text.Trim()
				mailSubject = [string] $mailSubjectBox.Text.Trim()
				reportDelivery = [string] $reportDeliveryCombo.Text.Trim()
				alertMinimumSeverity = [string] $alertSeverityCombo.Text.Trim()
				teamsWebhookUrl = [string] $teamsWebhookBox.Text.Trim()
				retentionDays = [int] $retentionDaysBox.Text.Trim()
				runNow = [bool] $runNowBox.Checked
				alwaysSendReport = [bool] $alwaysSendReportBox.Checked
				includeCopilotAndDataverse = [bool] $includeCopilotAndDataverseBox.Checked
				includeMaesterReport = [bool] $includeMaesterReportBox.Checked
			}
			$form.DialogResult = [System.Windows.Forms.DialogResult]::OK
			$form.Close()
		})

		& $showStep 0

		$dialogResult = $form.ShowDialog()
		if ($dialogResult -eq [System.Windows.Forms.DialogResult]::OK) {
			return $form.Tag
		}

		return $null
	}

	if ([System.Threading.Thread]::CurrentThread.GetApartmentState() -eq [System.Threading.ApartmentState]::STA) {
		return (& $uiScript $prefilled $subscriptions)
	}

	$runspace = [runspacefactory]::CreateRunspace()
	$runspace.ApartmentState = [System.Threading.ApartmentState]::STA
	$runspace.ThreadOptions = 'ReuseThread'
	$runspace.Open()
	$powershell = [powershell]::Create()
	try {
		$powershell.Runspace = $runspace
		[void] $powershell.AddScript($uiScript).AddArgument($prefilled).AddArgument($subscriptions)
		$result = $powershell.Invoke()
		if ($powershell.Streams.Error.Count -gt 0) {
			throw $powershell.Streams.Error[0]
		}

		return ($result | Select-Object -First 1)
	} finally {
		$powershell.Dispose()
		$runspace.Dispose()
	}
}

Assert-RequiredModule -Name Az.Accounts
Assert-RequiredModule -Name Az.Resources
Assert-RequiredModule -Name Az.Automation
Assert-RequiredModule -Name Az.Storage
Write-InstallLog "Starting DriftMaester installer version $($script:DriftMaesterVersion)."

if (-not (Get-AzContext -ErrorAction SilentlyContinue)) {
	Write-InstallLog 'Connecting to Azure.'
	Connect-AzAccount | Out-Null
}

# Determine if we need GUI or CLI mode
$needsGuiInput = $GuiMode -or [string]::IsNullOrWhiteSpace($Subscription) -or [string]::IsNullOrWhiteSpace($ResourceGroup) -or [string]::IsNullOrWhiteSpace($Recipients) -or [string]::IsNullOrWhiteSpace($Frequency) -or [string]::IsNullOrWhiteSpace($TimeOfDay)

if ($needsGuiInput -and [System.Runtime.InteropServices.RuntimeInformation]::IsOSPlatform([System.Runtime.InteropServices.OSPlatform]::Windows)) {
	Write-InstallLog 'Launching GUI installer...'
	$collectedParams = Show-DriftMaesterGui -PrefilledSubscription $Subscription -PrefilledResourceGroup $ResourceGroup -PrefilledLocation $Location -PrefilledFrequency $Frequency -PrefilledDayOfWeek $DayOfWeek -PrefilledDayOfMonth $DayOfMonth -PrefilledTimeOfDay $TimeOfDay -PrefilledTimeZone $TimeZone -PrefilledRecipients $Recipients -PrefilledAlertRecipients $AlertRecipients -PrefilledSenderUserId $SenderUserId -PrefilledDevOpsOrg $DevOpsOrg -PrefilledTenantId $TenantId -PrefilledMailSubject $MailSubject -PrefilledReportDelivery $ReportDelivery -PrefilledAlertMinimumSeverity $AlertMinimumSeverity -PrefilledTeamsWebhookUrl $TeamsWebhookUrl -PrefilledRetentionDays $RetentionDays -PrefilledRunNow ([bool]$RunNow) -PrefilledAlwaysSendReport $AlwaysSendReport -PrefilledIncludeCopilotAndDataverse $IncludeCopilotAndDataverse -PrefilledIncludeMaesterReport $IncludeMaesterReport

	if (-not $collectedParams) {
		Write-InstallLog 'Installation cancelled.' -Level Warning
		# return, not exit: the documented install pipes this script into Invoke-Expression, where exit would
		# terminate the operator's whole PowerShell session instead of just ending the installer.
		return
	}
	
	# Use GUI-collected parameters
	$Subscription = $collectedParams.subscription
	$ResourceGroup = $collectedParams.resourceGroup
	$Location = $collectedParams.location
	$Frequency = $collectedParams.frequency
	$DayOfWeek = $collectedParams.dayOfWeek
	$DayOfMonth = [int] $collectedParams.dayOfMonth
	$TimeOfDay = $collectedParams.timeOfDay
	$TimeZone = $collectedParams.timeZone
	$Recipients = $collectedParams.recipients
	$AlertRecipients = $collectedParams.alertRecipients
	$SenderUserId = $collectedParams.senderUserId
	$DevOpsOrg = $collectedParams.devopsOrg
	$TenantId = $collectedParams.tenantId
	$MailSubject = $collectedParams.mailSubject
	$ReportDelivery = $collectedParams.reportDelivery
	$AlertMinimumSeverity = $collectedParams.alertMinimumSeverity
	$TeamsWebhookUrl = $collectedParams.teamsWebhookUrl
	$RetentionDays = [int] $collectedParams.retentionDays
	$RunNow = [bool] $collectedParams.runNow
	$AlwaysSendReport = [bool] $collectedParams.alwaysSendReport
	$IncludeCopilotAndDataverse = [bool] $collectedParams.includeCopilotAndDataverse
	$IncludeMaesterReport = [bool] $collectedParams.includeMaesterReport
}

# Handle both headless mode (all parameters provided) and CLI mode (interactive fallback)
if (-not [string]::IsNullOrWhiteSpace($Subscription)) {
	Write-InstallLog 'Running installer with provided parameters.'
	
	# Validate required headless parameters
	if ([string]::IsNullOrWhiteSpace($ResourceGroup)) { throw 'ResourceGroup parameter is required in headless mode.' }
	if ([string]::IsNullOrWhiteSpace($Recipients)) { throw 'Recipients parameter is required in headless mode.' }
	if ([string]::IsNullOrWhiteSpace($Frequency)) { throw 'Frequency parameter is required in headless mode.' }
	if ([string]::IsNullOrWhiteSpace($TimeOfDay)) { throw 'TimeOfDay parameter is required in headless mode.' }

	$selectedSubscription = Get-AzSubscription -SubscriptionId $Subscription -ErrorAction Stop
	Set-DriftSubscriptionContext -TargetSubscriptionId $selectedSubscription.Id -TargetTenantId $selectedSubscription.TenantId
	$SubscriptionId = $selectedSubscription.Id
	$ResourceGroupName = $ResourceGroup
	$DeploymentLocation = if ([string]::IsNullOrWhiteSpace($Location)) { $script:Location } else { $Location }

	# Parse recipients
	$ReportRecipient = @(ConvertTo-RecipientArray -RecipientText $Recipients)
	Assert-ValidRecipients -Addresses $ReportRecipient -Label 'Recipients'
	$AlertRecipientValues = @(ConvertTo-RecipientArray -RecipientText $AlertRecipients)
	if ($AlertRecipientValues.Count -gt 0) {
		Assert-ValidRecipients -Addresses $AlertRecipientValues -Label 'AlertRecipients'
	}

	# Parse schedule - GUI provides lowercase values
	$frequencyMap = @{ 'daily' = 'Day'; 'weekly' = 'Week'; 'monthly' = 'Month' }
	$scheduleFrequency = $frequencyMap[$Frequency.ToLower()]

	if (-not $scheduleFrequency) {
		throw "Unsupported Frequency value '$Frequency'. Allowed values: daily, weekly, monthly."
	}

	$timeOfDaySpan = ConvertTo-TimeOfDaySpan -Value $TimeOfDay
	$ScheduleTimeZone = Resolve-DriftTimeZoneId -TimeZoneId $TimeZone

	# Build schedule selection based on frequency
	$invokeSchedule = switch ($scheduleFrequency) {
		'Day' {
			[PSCustomObject]@{
				Frequency       = 'Day'
				StartTime       = Get-NextDailyOccurrence -TimeOfDay $timeOfDaySpan -TimeZoneId $ScheduleTimeZone
				TimeOfDay       = $timeOfDaySpan
				TimeZoneId      = $ScheduleTimeZone
				WeekDays        = @()
				MonthDays       = @()
				MonthlyDayMode  = $null
				DescriptionText = 'daily'
			}
		}
		'Week' {
			$selectedDay = if ([string]::IsNullOrWhiteSpace($DayOfWeek)) { 'Monday' } else { [string] $DayOfWeek }
			$allowedDays = @('Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday')
			if ($allowedDays -notcontains $selectedDay) {
				throw "Unsupported DayOfWeek value '$selectedDay'. Allowed values: $($allowedDays -join ', ')."
			}
			[PSCustomObject]@{
				Frequency       = 'Week'
				StartTime       = Get-NextWeeklyOccurrence -DayOfWeek $selectedDay -TimeOfDay $timeOfDaySpan -TimeZoneId $ScheduleTimeZone
				TimeOfDay       = $timeOfDaySpan
				TimeZoneId      = $ScheduleTimeZone
				WeekDays        = @($selectedDay)
				MonthDays       = @()
				MonthlyDayMode  = $null
				DescriptionText = "weekly on $selectedDay"
			}
		}
		'Month' {
			$selectedMonthDay = if ($PSBoundParameters.ContainsKey('DayOfMonth') -and $DayOfMonth -ne 0) { $DayOfMonth } else { 1 }
			if (($selectedMonthDay -lt 1 -or $selectedMonthDay -gt 31) -and $selectedMonthDay -ne -1) {
				throw "DayOfMonth must be between 1 and 31, or -1 for the last day. Received '$selectedMonthDay'."
			}
			$monthDescription = if ($selectedMonthDay -eq -1) { 'last day' } else { "day $selectedMonthDay" }
			[PSCustomObject]@{
				Frequency       = 'Month'
				StartTime       = Get-NextMonthlyOccurrence -DayMode 'DayOfMonth' -DayOfMonth $selectedMonthDay -TimeOfDay $timeOfDaySpan -TimeZoneId $ScheduleTimeZone
				TimeOfDay       = $timeOfDaySpan
				TimeZoneId      = $ScheduleTimeZone
				WeekDays        = @()
				MonthDays       = @($selectedMonthDay)
				MonthlyDayMode  = 'DayOfMonth'
				DescriptionText = "monthly on $monthDescription"
			}
		}
	}

	$updateSchedule = New-UpdateScheduleSelection -InvokeSchedule $invokeSchedule

	$MailSenderUserId = $SenderUserId
	$DevOpsOrganization = $DevOpsOrg
	if ([string]::IsNullOrWhiteSpace($TenantId)) {
		$TenantId = $selectedSubscription.TenantId
	}
	$MailSubjectPrefix = if ([string]::IsNullOrWhiteSpace($MailSubject)) { 'DriftMaester Report' } else { $MailSubject }
	$AlwaysSendReportEnabled = [bool] $AlwaysSendReport
	$IncludeCopilotAndDataverseEnabled = [bool] $IncludeCopilotAndDataverse
	$IncludeMaesterReportEnabled = [bool] $IncludeMaesterReport
	$ReportDeliveryMode = if ([string]::IsNullOrWhiteSpace($ReportDelivery)) { 'auto' } else { $ReportDelivery.ToLowerInvariant() }
	$AlertSeverityMode = if ([string]::IsNullOrWhiteSpace($AlertMinimumSeverity)) { 'none' } else { $AlertMinimumSeverity.ToLowerInvariant() }
	$AlertRecipientsText = ($AlertRecipientValues -join ',')
	if ([string]::IsNullOrWhiteSpace($AlertRecipientsText)) {
		$AlertRecipientsText = $recipientParameter
	}

	Connect-ToGraphForInstall -RequestedTenantId $TenantId
} else {
	# CLI mode - interactive prompts (original behavior)
	$selectedSubscription = Select-AzureSubscription
	Set-DriftSubscriptionContext -TargetSubscriptionId $selectedSubscription.Id -TargetTenantId $selectedSubscription.TenantId
	$SubscriptionId = $selectedSubscription.Id
	$ResourceGroupName = Read-RequiredValue -Prompt 'Desired resource group name'
	$DeploymentLocation = Read-OptionalValue -Prompt 'Azure deployment location, press enter to use westeurope' -DefaultValue $script:Location
	$invokeSchedule = New-DriftScheduleSelection
	$updateSchedule = New-UpdateScheduleSelection -InvokeSchedule $invokeSchedule

	$recipientText = Read-RequiredValue -Prompt 'Which email addresses should receive reports, comma-separated'
	$ReportRecipient = @(ConvertTo-RecipientArray -RecipientText $recipientText)
	Assert-ValidRecipients -Addresses $ReportRecipient -Label 'Recipients'
	$AlertRecipientsText = Read-OptionalValue -Prompt '[OPTIONAL] Alert-only recipients, comma-separated (uses main recipients if omitted)'
	$alertRecipientsArray = @(ConvertTo-RecipientArray -RecipientText $AlertRecipientsText)
	if ($alertRecipientsArray.Count -gt 0) {
		Assert-ValidRecipients -Addresses $alertRecipientsArray -Label 'AlertRecipients'
		$AlertRecipientsText = ($alertRecipientsArray -join ',')
	}
	$MailSenderUserId = Read-OptionalValue -Prompt '[OPTIONAL] Mail sender UPN or user id, press enter to use first recipient'
	$DevOpsOrganization = Read-OptionalValue -Prompt 'Azure DevOps organization for Maester checks, optional'
	$TenantId = Read-OptionalValue -Prompt '[OPTIONAL] Tenant id to pass to the runbooks, press enter to use the tenant of the selected subscription'
	if ([string]::IsNullOrWhiteSpace($TenantId)) {
		$TenantId = $selectedSubscription.TenantId
	}
	$MailSubjectPrefix = Read-OptionalValue -Prompt '[OPTIONAL] Mail subject prefix for report emails, press enter to use default'
	$ReportDeliveryMode = (Read-OptionalValue -Prompt '[OPTIONAL] Report delivery mode (auto, attach, link)' -DefaultValue 'auto').ToLowerInvariant()
	$AlertSeverityMode = (Read-OptionalValue -Prompt '[OPTIONAL] Alert minimum severity (none, low, medium, high, critical)' -DefaultValue 'none').ToLowerInvariant()
	$TeamsWebhookUrl = Read-OptionalValue -Prompt '[OPTIONAL] Teams incoming webhook URL for summary notifications'
	$retentionInput = Read-OptionalValue -Prompt '[OPTIONAL] Result retention days (30-3650)' -DefaultValue '180'
	$retentionCandidate = 0
	if ([int]::TryParse($retentionInput, [ref] $retentionCandidate) -and $retentionCandidate -ge 30 -and $retentionCandidate -le 3650) {
		$RetentionDays = $retentionCandidate
	} else {
		throw 'Retention days must be an integer between 30 and 3650.'
	}
	$AlwaysSendReportEnabled = Read-YesNo -Prompt 'Always send report emails, even when no drift is detected?' -DefaultNo
	$IncludeCopilotAndDataverseEnabled = Read-YesNo -Prompt 'Include Copilot, Power Platform, Dynamics, and Dataverse checks?' -DefaultNo
	$IncludeMaesterReportEnabled = Read-YesNo -Prompt 'Attach the full original Maester report (zipped) to the report email? Note: in larger tenants this attachment can become big and may be rejected by the recipient mail system.' -DefaultNo
	$RunNow = Read-YesNo -Prompt 'Run update and invoke immediately after installation?' -DefaultNo
	Connect-ToGraphForInstall -RequestedTenantId $TenantId
}

if ([string]::IsNullOrWhiteSpace($ReportDeliveryMode)) {
	$ReportDeliveryMode = 'auto'
}
if ([string]::IsNullOrWhiteSpace($AlertSeverityMode)) {
	$AlertSeverityMode = 'none'
}

$recipientParameter = ($ReportRecipient | ForEach-Object { $_.Trim() } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }) -join ','
if ([string]::IsNullOrWhiteSpace($recipientParameter)) {
	throw 'At least one report recipient is required.'
}

Set-DriftSubscriptionContext -TargetSubscriptionId $SubscriptionId -TargetTenantId $selectedSubscription.TenantId

Register-DriftAzureProvider -ProviderNamespace 'Microsoft.Automation'
Register-DriftAzureProvider -ProviderNamespace 'Microsoft.Storage'
Register-DriftAzureProvider -ProviderNamespace 'Microsoft.Authorization'
Write-InstallLog 'Step 1/8: Azure providers ready.'

$names = Get-DriftMaesterNames -SelectedSubscriptionId $SubscriptionId -SelectedResourceGroupName $ResourceGroupName
$resourceGroup = Set-DriftResourceGroup -Name $ResourceGroupName -TargetLocation $DeploymentLocation
$resourceLocation = if ($resourceGroup -and -not [string]::IsNullOrWhiteSpace([string] $resourceGroup.Location)) { [string] $resourceGroup.Location } else { $DeploymentLocation }
$automationAccount = Set-DriftAutomationAccount -Name $names.AutomationAccountName -TargetResourceGroupName $ResourceGroupName -TargetLocation $resourceLocation
$storageAccount = Set-DriftStorageAccount -Name $names.StorageAccountName -TargetResourceGroupName $ResourceGroupName -TargetLocation $resourceLocation
Set-DriftStorageSecurityPosture -ResourceGroupName $ResourceGroupName -StorageAccountName $names.StorageAccountName
Set-DriftStorageLifecyclePolicy -ResourceGroupName $ResourceGroupName -StorageAccountName $names.StorageAccountName -RetentionDays $RetentionDays
Write-InstallLog 'Step 2/8: Core resources created or reused.'

$automationAccount = Get-AzAutomationAccount -ResourceGroupName $ResourceGroupName -Name $names.AutomationAccountName
if (-not $automationAccount.Identity -or -not $automationAccount.Identity.PrincipalId) {
	throw "Automation Account '$($names.AutomationAccountName)' does not have a system-assigned managed identity after creation."
}

$managedIdentityPrincipalId = [string] $automationAccount.Identity.PrincipalId
$managedIdentityClientId = $null
if ($automationAccount.Identity.UserAssignedIdentities) {
	Write-InstallLog 'Automation Account has user-assigned identities, but DriftMaester uses the system-assigned identity for this install.' -Level Warning
}

$accountResource = Get-AzResource -ResourceGroupName $ResourceGroupName -ResourceType 'Microsoft.Automation/automationAccounts' -Name $names.AutomationAccountName
if ($accountResource.Identity -and $accountResource.Identity.PrincipalId) {
	$managedIdentityPrincipalId = [string] $accountResource.Identity.PrincipalId
	if ($accountResource.Identity.ClientId) {
		$managedIdentityClientId = [string] $accountResource.Identity.ClientId
	}
}

if ([string]::IsNullOrWhiteSpace($managedIdentityClientId)) {
	Write-InstallLog 'ClientId for the managed identity was not available immediately after Automation Account creation. This is expected, and the installer will wait for it to become available.' -Level Warning
	$stopTime = (Get-Date).AddMinutes(5)
	while ([string]::IsNullOrWhiteSpace($managedIdentityClientId) -and (Get-Date) -lt $stopTime) {
		$principal = Get-AzADServicePrincipal -ObjectId $managedIdentityPrincipalId -ErrorAction SilentlyContinue
		if ($principal -and $principal.AppId) {
			$managedIdentityClientId = [string] $principal.AppId
		} else {
			Start-Sleep -Seconds 15
		}
	}
}

if ([string]::IsNullOrWhiteSpace($managedIdentityClientId)) {
	throw 'Could not resolve the managed identity client id needed for API permission assignment.'
}

$subscriptionScope = "/subscriptions/$SubscriptionId"
$resourceGroupScope = "$subscriptionScope/resourceGroups/$ResourceGroupName"
$storageScope = $storageAccount.Id
$automationScope = $accountResource.ResourceId

Set-DriftAzRoleAssignment -ObjectId $managedIdentityPrincipalId -RoleDefinitionName 'Reader' -Scope $resourceGroupScope
Set-DriftAzRoleAssignment -ObjectId $managedIdentityPrincipalId -RoleDefinitionName 'Storage Blob Data Contributor' -Scope $storageScope
Set-DriftAzRoleAssignment -ObjectId $managedIdentityPrincipalId -RoleDefinitionName 'Automation Contributor' -Scope $automationScope
Set-DriftAzureRootReaderAssignments -ServicePrincipalObjectId $managedIdentityPrincipalId
Write-InstallLog 'Step 3/8: Azure RBAC and root reader assignments reconciled.'

Set-DriftRuntimeEnvironment -SelectedSubscriptionId $SubscriptionId -TargetResourceGroupName $ResourceGroupName -AutomationAccountName $names.AutomationAccountName -TargetLocation $resourceLocation | Out-Null
Set-DriftRunbookFromGithub -SelectedSubscriptionId $SubscriptionId -TargetResourceGroupName $ResourceGroupName -AutomationAccountName $names.AutomationAccountName -TargetLocation $resourceLocation -RunbookName $script:UpdateRunbookName -SourceFileName 'Update-DriftMaester.ps1'
Set-DriftRunbookFromGithub -SelectedSubscriptionId $SubscriptionId -TargetResourceGroupName $ResourceGroupName -AutomationAccountName $names.AutomationAccountName -TargetLocation $resourceLocation -RunbookName $script:InvokeRunbookName -SourceFileName 'Invoke-DriftMaester.ps1'
Write-InstallLog 'Step 4/8: Runtime and runbooks reconciled.'

Set-DriftAutomationSchedule -SelectedSubscriptionId $SubscriptionId -TargetResourceGroupName $ResourceGroupName -AutomationAccountName $names.AutomationAccountName -ScheduleName $script:InvokeScheduleName -ScheduleSelection $invokeSchedule -Description 'Runs DriftMaester tenant drift detection.'
Set-DriftAutomationSchedule -SelectedSubscriptionId $SubscriptionId -TargetResourceGroupName $ResourceGroupName -AutomationAccountName $names.AutomationAccountName -ScheduleName $script:UpdateScheduleName -ScheduleSelection $updateSchedule -Description 'Updates DriftMaester runtime modules one hour before the invoke schedule.'
Write-InstallLog 'Step 5/8: Schedules reconciled.'

$invokeParameters = Get-InvokeParameters -Recipients $recipientParameter -SenderUserId $MailSenderUserId -DevOpsOrg $DevOpsOrganization -TargetTenantId $TenantId -SubjectPrefix $MailSubjectPrefix -ReportDelivery $ReportDeliveryMode -AlertMinimumSeverity $AlertSeverityMode -AlertRecipients $AlertRecipientsText -TeamsWebhookUrl $TeamsWebhookUrl -RetentionDays $RetentionDays -AlwaysSendReport $AlwaysSendReportEnabled -IncludeCopilotAndDataverse $IncludeCopilotAndDataverseEnabled -IncludeMaesterReport $IncludeMaesterReportEnabled
Set-DriftJobSchedule -SelectedSubscriptionId $SubscriptionId -TargetResourceGroupName $ResourceGroupName -AutomationAccountName $names.AutomationAccountName -RunbookName $script:InvokeRunbookName -ScheduleName $script:InvokeScheduleName -Parameters $invokeParameters
Set-DriftJobSchedule -SelectedSubscriptionId $SubscriptionId -TargetResourceGroupName $ResourceGroupName -AutomationAccountName $names.AutomationAccountName -RunbookName $script:UpdateRunbookName -ScheduleName $script:UpdateScheduleName
Write-InstallLog 'Step 6/8: Job schedules updated.'

Set-DriftManagedIdentityApiPermissions -ManagedIdentityClientId $managedIdentityClientId -RequestedTenantId $TenantId -DevOpsOrganization $DevOpsOrganization
Write-InstallLog 'Step 7/8: Graph and SharePoint API permissions and directory roles reconciled.'
$exchangeOrganization = Get-InitialTenantDomainFromGraph
# Mirror the runbook's sender logic (mailsenderuserid, else first recipient) so the Mail.Send RBAC scope targets the mailbox that actually sends the report.
$effectiveMailSender = if (-not [string]::IsNullOrWhiteSpace($MailSenderUserId)) { $MailSenderUserId } else { $ReportRecipient | Select-Object -First 1 }
$mailSendRbacConfigured = Set-DriftExchangeOnlineRbac -ManagedIdentityClientId $managedIdentityClientId -ManagedIdentityObjectId $managedIdentityPrincipalId -DisplayName $names.AutomationAccountName -Organization $exchangeOrganization -SenderMailbox $effectiveMailSender
if ($mailSendRbacConfigured) {
	# RBAC mail sending works; remove any broad Graph Mail.Send permission left over from a previous fallback run.
	Set-DriftGraphMailSendFallback -ManagedIdentityClientId $managedIdentityClientId -Enabled $false
} else {
	Write-InstallLog "Falling back to the broad Microsoft Graph 'Mail.Send' application permission so DriftMaester can still send reports. This lets the managed identity send mail as ANY mailbox in the tenant, which is more permissive than the Exchange RBAC model. If you would rather avoid this, resolve the Exchange Online RBAC issue reported above and re-run this installer; it will then switch to scoped mail sending and remove this broad permission automatically." -Level Warning
	Set-DriftGraphMailSendFallback -ManagedIdentityClientId $managedIdentityClientId -Enabled $true
}
Write-InstallLog 'Step 8/8: Exchange Online mail permissions reconciled.'

Write-DriftAccessReport -ManagedIdentityObjectId $managedIdentityPrincipalId -ManagedIdentityClientId $managedIdentityClientId -SubscriptionId $SubscriptionId -ResourceGroupName $ResourceGroupName -AutomationAccountName $names.AutomationAccountName -GraphPermissions $RequiredGraphApplicationPermissions -DirectoryRoles $DirectoryRolesForManagedIdentity -SharePointPermissions $RequiredSharePointApplicationPermissions -ExchangeOrganization $exchangeOrganization

$updateJobId = Start-UpdateRunbook -SelectedSubscriptionId $SubscriptionId -TargetResourceGroupName $ResourceGroupName -AutomationAccountName $names.AutomationAccountName
if ($RunNow) {
	Write-InstallLog 'RunNow selected. Waiting for update runbook to complete before starting invoke runbook.'
	$updateJobResult = Wait-ForAutomationJobCompletion -ResourceGroupName $ResourceGroupName -AutomationAccountName $names.AutomationAccountName -JobId $updateJobId -TimeoutMinutes 40 -PollSeconds 20
	if ($updateJobResult.Status -ne 'Completed') {
		throw "RunNow aborted because update runbook ended in status '$($updateJobResult.Status)'."
	}

	$invokeJobId = Start-InvokeRunbookNow -SelectedSubscriptionId $SubscriptionId -TargetResourceGroupName $ResourceGroupName -AutomationAccountName $names.AutomationAccountName -Parameters $invokeParameters
	Write-InstallLog "RunNow started invoke runbook job '$invokeJobId'."
}

Write-InstallLog ("Install summary: AutomationAccount='{0}', ResourceGroup='{1}', StorageAccount='{2}', InvokeSchedule='{3}', UpdateSchedule='{4}', ReportDelivery='{5}', AlertMinimumSeverity='{6}', RetentionDays='{7}'" -f $names.AutomationAccountName, $ResourceGroupName, $names.StorageAccountName, $invokeSchedule.DescriptionText, $updateSchedule.DescriptionText, $ReportDeliveryMode, $AlertSeverityMode, $RetentionDays)
if ($RunNow) {
	Write-InstallLog 'RunNow workflow completed. The first invoke runbook execution has started.' -Level Success
} else {
	Write-InstallLog 'After the Update-DriftMaester run is complete, you can manually run the Invoke-DriftMaester runbook, or wait for the next scheduled run.' -Level Info
}
Write-InstallLog 'DriftMaester installation completed.' -Level Success