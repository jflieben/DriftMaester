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
	[Parameter(Mandatory = $false)][string] $Frequency,
	[Parameter(Mandatory = $false)][string] $TimeOfDay,
	[Parameter(Mandatory = $false)][string] $Recipients,
	[Parameter(Mandatory = $false)][string] $SenderUserId,
	[Parameter(Mandatory = $false)][string] $DevOpsOrg,
	[Parameter(Mandatory = $false)][string] $TenantId,
	[Parameter(Mandatory = $false)][string] $MailSubject,
	[Parameter(Mandatory = $false)][bool] $AlwaysSendReport = $false,
	[Parameter(Mandatory = $false)][bool] $IncludeCopilotAndDataverse = $false
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
$script:GithubRawBase = 'https://raw.githubusercontent.com/jflieben/DriftMaester/main/Runbooks'
$script:GraphAppId = '00000003-0000-0000-c000-000000000000'
$script:ExchangeOnlineAppId = '00000002-0000-0ff1-ce00-000000000000'
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
    'PrivilegedAccess.Read.AzureAD',
    'OrgSettings-Forms.Read.All',
    'DeviceManagementServiceConfig.Read.All',
    'SecurityIdentitiesHealth.Read.All',
    'DirectoryRecommendations.Read.All',
    'Directory.Read.All',
    'ReportSettings.Read.All',
    'RoleManagement.Read.All',
    'DeviceManagementRBAC.Read.All',
    'SecurityIdentitiesSensors.Read.All',
    'DeviceManagementConfiguration.Read.All',
    'OrgSettings-AppsAndServices.Read.All',
    'Mail.Send',
    'IdentityRiskEvent.Read.All',
    'Policy.Read.All',
    'Reports.Read.All',
    'ThreatHunting.Read.All',
    'AuditLog.Read.All'

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
	return (New-AzResourceGroup -Name $Name -Location $TargetLocation)
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
		$storageAccount = New-AzStorageAccount -ResourceGroupName $TargetResourceGroupName -Name $Name -Location $TargetLocation -SkuName Standard_LRS -Kind StorageV2 -EnableHttpsTrafficOnly $true -MinimumTlsVersion TLS1_2 -AllowBlobPublicAccess $false
	} else {
		Write-InstallLog "Storage account '$Name' already exists."
	}

	$context = $storageAccount.Context
	if (-not (Get-AzStorageContainer -Context $context -Name 'maester' -ErrorAction SilentlyContinue)) {
		Write-InstallLog "Creating blob container 'maester'."
		New-AzStorageContainer -Context $context -Name 'maester' -Permission Off | Out-Null
	}

	return $storageAccount
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

function Get-NextDailyOccurrence {
	param([Parameter(Mandatory = $true)][timespan] $TimeOfDay)

	$now = Get-Date
	$candidate = $now.Date.Add($TimeOfDay)
	if ($candidate -le $now.AddMinutes(10)) {
		$candidate = $candidate.AddDays(1)
	}

	return $candidate
}

function Get-NextWeeklyOccurrence {
	param(
		[Parameter(Mandatory = $true)][string] $DayOfWeek,
		[Parameter(Mandatory = $true)][timespan] $TimeOfDay
	)

	$day = [System.Enum]::Parse([System.DayOfWeek], $DayOfWeek, $true)
	$now = Get-Date
	$candidate = $now.Date.Add($TimeOfDay)
	$delta = (([int] $day) - ([int] $candidate.DayOfWeek) + 7) % 7
	$candidate = $candidate.AddDays($delta)
	if ($candidate -le $now.AddMinutes(10)) {
		$candidate = $candidate.AddDays(7)
	}

	return $candidate
}

function Get-NextMonthlyOccurrence {
	param(
		[Parameter(Mandatory = $true)][ValidateSet('First')][string] $DayMode,
		[Parameter(Mandatory = $true)][timespan] $TimeOfDay
	)

	$now = Get-Date
	$candidate = (Get-Date -Year $now.Year -Month $now.Month -Day 1 -Hour $TimeOfDay.Hours -Minute $TimeOfDay.Minutes -Second 0)
	if ($candidate -le $now.AddMinutes(10)) {
		$nextMonth = $now.AddMonths(1)
		$candidate = (Get-Date -Year $nextMonth.Year -Month $nextMonth.Month -Day 1 -Hour $TimeOfDay.Hours -Minute $TimeOfDay.Minutes -Second 0)
	}

	return $candidate
}

function New-DriftScheduleSelection {
	$frequencyInput = Read-RequiredValue -Prompt 'Run frequency (daily, weekly, monthly)' -DefaultValue 'daily'
	$frequencyInput = $frequencyInput.Trim().ToLowerInvariant()
	$timeInput = Read-RequiredValue -Prompt 'Run time (HH:mm local)' -DefaultValue '02:00'
	$timeParts = $timeInput -split ':'
	if ($timeParts.Count -ne 2) { throw 'Time must be in HH:mm format.' }
	$timeOfDay = [timespan]::new([int]$timeParts[0], [int]$timeParts[1], 0)

	switch ($frequencyInput) {
		'weekly' {
			$dayOfWeek = Read-RequiredValue -Prompt 'Day of week for weekly schedule (Monday..Sunday)' -DefaultValue 'Monday'
			return [PSCustomObject]@{
				Frequency       = 'Week'
				StartTime       = Get-NextWeeklyOccurrence -DayOfWeek $dayOfWeek -TimeOfDay $timeOfDay
				TimeOfDay       = $timeOfDay
				WeekDays        = @($dayOfWeek)
				MonthDays       = @()
				MonthlyDayMode  = $null
				DescriptionText = "weekly on $dayOfWeek"
			}
		}
		'monthly' {
			return [PSCustomObject]@{
				Frequency       = 'Month'
				StartTime       = Get-NextMonthlyOccurrence -DayMode 'First' -TimeOfDay $timeOfDay
				TimeOfDay       = $timeOfDay
				WeekDays        = @()
				MonthDays       = @(1)
				MonthlyDayMode  = 'First'
				DescriptionText = 'monthly on the 1st day'
			}
		}
		default {
			return [PSCustomObject]@{
				Frequency       = 'Day'
				StartTime       = Get-NextDailyOccurrence -TimeOfDay $timeOfDay
				TimeOfDay       = $timeOfDay
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

	$start = $InvokeSchedule.StartTime.AddHours(-1)
	$now = Get-Date
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
		WeekDays        = @($InvokeSchedule.WeekDays)
		MonthDays       = @($InvokeSchedule.MonthDays)
		MonthlyDayMode  = $InvokeSchedule.MonthlyDayMode
		DescriptionText = "$($InvokeSchedule.DescriptionText), one hour before invoke"
	}
}

function Enable-DriftAzureRootAccess {
	Write-InstallLog 'Attempting to elevate Azure root access for the current account. This only works for a Global Administrator with access management elevation allowed.' -Level Warning
	Invoke-AzRestMethod -Path '/providers/Microsoft.Authorization/elevateAccess?api-version=2015-07-01' -Method POST | Out-Null
}

function Disable-DriftAzureRootAccess {
	$currentContext = Get-AzContext
	$currentAccountId = [string] $currentContext.Account.Id
	if ([string]::IsNullOrWhiteSpace($currentAccountId)) {
		Write-InstallLog 'Could not determine the current Azure account for root access cleanup.' -Level Warning
		return
	}

	$assignments = @(Get-AzRoleAssignment -RoleDefinitionId '18d7d88d-d35e-4fb5-a5c3-7773c20a72d9' -ErrorAction SilentlyContinue | Where-Object {
		$_.Scope -eq '/' -and $_.SignInName -eq $currentAccountId
	})

	foreach ($assignment in $assignments) {
		$assignmentPath = [string] $assignment.RoleAssignmentId
		if ([string]::IsNullOrWhiteSpace($assignmentPath)) { continue }
		if (-not $assignmentPath.StartsWith('/')) { $assignmentPath = "/$assignmentPath" }

		Write-InstallLog "Removing temporary Azure root User Access Administrator elevation for '$currentAccountId'."
		Invoke-AzRestMethod -Path "$assignmentPath?api-version=2018-07-01" -Method DELETE | Out-Null
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
			logProgress        = $true
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
	$body = @{
		properties = @{
			description = $Description
			startTime   = $ScheduleSelection.StartTime.ToString('o')
			expiryTime  = '9999-12-31T23:59:59Z'
			interval    = 1
			frequency   = $ScheduleSelection.Frequency
			timeZone    = [System.TimeZoneInfo]::Local.Id
		}
	}

	if ($ScheduleSelection.Frequency -eq 'Week') {
		$body.properties['advancedSchedule'] = @{ weekDays = @($ScheduleSelection.WeekDays) }
	} elseif ($ScheduleSelection.Frequency -eq 'Month') {
		$body.properties['advancedSchedule'] = @{ monthDays = @($ScheduleSelection.MonthDays) }
	}

	Write-InstallLog "Creating/updating schedule '$ScheduleName' starting '$($ScheduleSelection.StartTime.ToString('u'))' ($($ScheduleSelection.DescriptionText))."
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

function Connect-ToGraphForInstall {
	param([Parameter(Mandatory = $false)][string] $RequestedTenantId)

	Assert-RequiredModule -Name Microsoft.Graph.Authentication
	$requiredScopes = @(
		'Application.Read.All',
		'AppRoleAssignment.ReadWrite.All',
		'Directory.ReadWrite.All',
		'RoleManagement.ReadWrite.Directory'
	)

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

		$params = @{ Scopes = $requiredScopes; NoWelcome = $true }
		if ($RequestedTenantId) { $params['TenantId'] = $RequestedTenantId }
		Write-InstallLog 'Connecting to Microsoft Graph. Use a Global Administrator or Privileged Role Administrator account.'
		Connect-MgGraph @params | Out-Null
	}
}

function Connect-ToExchangeOnlineForInstall {
	param([Parameter(Mandatory = $false)][string] $Organization)

	Assert-RequiredModule -Name ExchangeOnlineManagement

	$connection = Get-ConnectionInformation -ErrorAction SilentlyContinue | Where-Object { $_.State -eq 'Connected' } | Select-Object -First 1
	if ($connection) {
		Write-Output 'Current Exchange Online login:'
		Write-Output "  User:         $($connection.UserPrincipalName)"
		Write-Output "  Organization: $($connection.Organization)"
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

function Set-DriftExchangeOnlineRbac {
	param(
		[Parameter(Mandatory = $true)][string] $ManagedIdentityClientId,
		[Parameter(Mandatory = $true)][string] $ManagedIdentityObjectId,
		[Parameter(Mandatory = $true)][string] $DisplayName,
		[Parameter(Mandatory = $false)][string] $Organization
	)

	Connect-ToExchangeOnlineForInstall -Organization $Organization

	foreach ($commandName in @('Get-ServicePrincipal', 'New-ServicePrincipal', 'Get-ManagementRoleAssignment', 'New-ManagementRoleAssignment')) {
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

	$existingAssignments = @(Get-ManagementRoleAssignment -Role 'View-Only Configuration' -ErrorAction Stop | Where-Object {
		$_.RoleAssigneeName -eq $exchangeAppName -or
		$_.RoleAssignee -eq $exchangeAppName -or
		$_.App -eq $exchangeAppName
	})

	if ($existingAssignments.Count -gt 0) {
		Write-InstallLog "Exchange Online role assignment 'View-Only Configuration' already exists for '$exchangeAppName'."
		return
	}

	Write-InstallLog "Assigning Exchange Online role 'View-Only Configuration' to '$exchangeAppName'."
	New-ManagementRoleAssignment -Role 'View-Only Configuration' -App $exchangeAppName | Out-Null
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

function Set-DriftDirectoryRoleAssignment {
	param(
		[Parameter(Mandatory = $true)][string] $ManagedIdentityServicePrincipalId,
		[Parameter(Mandatory = $true)][string] $RoleDisplayName
	)

	$roleFilter = [Uri]::EscapeDataString("displayName eq '$RoleDisplayName'")
	$role = @(Invoke-GraphRequestAllPages -Uri "https://graph.microsoft.com/v1.0/directoryRoles?`$filter=$roleFilter") | Select-Object -First 1
	if (-not $role) {
		$template = @(Invoke-GraphRequestAllPages -Uri "https://graph.microsoft.com/v1.0/directoryRoleTemplates?`$filter=$roleFilter") | Select-Object -First 1
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
		Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/v1.0/directoryRoles/$($role.id)/members/`$ref" -Body ($body | ConvertTo-Json) -ContentType 'application/json' | Out-Null
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

		throw
	}
}

function Set-DriftManagedIdentityApiPermissions {
	param(
		[Parameter(Mandatory = $true)][string] $ManagedIdentityClientId,
		[Parameter(Mandatory = $false)][string] $RequestedTenantId
	)

	Connect-ToGraphForInstall -RequestedTenantId $RequestedTenantId

	$managedIdentityServicePrincipal = Get-ServicePrincipalByAppId -AppId $ManagedIdentityClientId
	$graphServicePrincipal = Get-ServicePrincipalByAppId -AppId $script:GraphAppId
	$exchangeServicePrincipal = Get-ServicePrincipalByAppId -AppId $script:ExchangeOnlineAppId

	foreach ($permission in ($RequiredGraphApplicationPermissions | Sort-Object -Unique)) {
		Set-DriftAppRoleAssignment -ManagedIdentityServicePrincipal $managedIdentityServicePrincipal -ResourceServicePrincipal $graphServicePrincipal -AppRoleValue $permission
	}

	Set-DriftAppRoleAssignment -ManagedIdentityServicePrincipal $managedIdentityServicePrincipal -ResourceServicePrincipal $exchangeServicePrincipal -AppRoleValue 'Exchange.ManageAsApp'

	foreach ($roleName in $DirectoryRolesForManagedIdentity) {
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
		[Parameter(Mandatory = $false)][bool] $AlwaysSendReport = $false,
		[Parameter(Mandatory = $false)][bool] $IncludeCopilotAndDataverse = $false
	)

	$parameters = @{
		reportrecipient = $Recipients
		alwayssendreport = $AlwaysSendReport
		includeCopilotAndDataverse = $IncludeCopilotAndDataverse
	}
	if (-not [string]::IsNullOrWhiteSpace($SubjectPrefix)) { $parameters['mailsubjectprefix'] = $SubjectPrefix } else{ $parameters['mailsubjectprefix'] = 'DriftMaester Report' }
	if (-not [string]::IsNullOrWhiteSpace($SenderUserId)) { $parameters['mailsenderuserid'] = $SenderUserId }
	if (-not [string]::IsNullOrWhiteSpace($DevOpsOrg)) { $parameters['devopsorganization'] = $DevOpsOrg }
	if (-not [string]::IsNullOrWhiteSpace($TargetTenantId)) { $parameters['tenantid'] = $TargetTenantId }

	return $parameters
}

# GUI Functions
function Show-DriftMaesterGui {
	param(
		[Parameter(Mandatory = $false)][string] $PrefilledSubscription,
		[Parameter(Mandatory = $false)][string] $PrefilledResourceGroup,
		[Parameter(Mandatory = $false)][string] $PrefilledLocation,
		[Parameter(Mandatory = $false)][string] $PrefilledFrequency,
		[Parameter(Mandatory = $false)][string] $PrefilledTimeOfDay,
		[Parameter(Mandatory = $false)][string] $PrefilledRecipients,
		[Parameter(Mandatory = $false)][string] $PrefilledSenderUserId,
		[Parameter(Mandatory = $false)][string] $PrefilledDevOpsOrg,
		[Parameter(Mandatory = $false)][string] $PrefilledTenantId,
		[Parameter(Mandatory = $false)][string] $PrefilledMailSubject,
		[Parameter(Mandatory = $false)][bool] $PrefilledAlwaysSendReport = $false,
		[Parameter(Mandatory = $false)][bool] $PrefilledIncludeCopilotAndDataverse = $false
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
		timeOfDay = if ([string]::IsNullOrWhiteSpace($PrefilledTimeOfDay)) { '02:00' } else { $PrefilledTimeOfDay }
		recipients = if ([string]::IsNullOrWhiteSpace($PrefilledRecipients)) { $defaultRecipient } else { $PrefilledRecipients }
		senderUserId = $PrefilledSenderUserId
		devopsOrg = $PrefilledDevOpsOrg
		tenantId = $PrefilledTenantId
		mailSubject = if ([string]::IsNullOrWhiteSpace($PrefilledMailSubject)) { 'DriftMaester Report' } else { $PrefilledMailSubject }
		alwaysSendReport = $PrefilledAlwaysSendReport
		includeCopilotAndDataverse = $PrefilledIncludeCopilotAndDataverse
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
		$form.Size = [System.Drawing.Size]::new(900, 720)
		$form.MinimumSize = [System.Drawing.Size]::new(860, 680)
		$form.Font = [System.Drawing.Font]::new('Segoe UI', 9)
		$form.MaximizeBox = $false

		function New-Label {
			param([string] $Text, [int] $X, [int] $Y)
			$label = [System.Windows.Forms.Label]::new()
			$label.Text = $Text
			$label.Location = [System.Drawing.Point]::new($X, $Y)
			$label.Size = [System.Drawing.Size]::new(230, 26)
			$label.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
			$form.Controls.Add($label)
			return $label
		}

		function New-TextBox {
			param([int] $X, [int] $Y, [string] $Text, [int] $Width = 560)
			$textBox = [System.Windows.Forms.TextBox]::new()
			$textBox.Location = [System.Drawing.Point]::new($X, $Y)
			$textBox.Size = [System.Drawing.Size]::new($Width, 28)
			$textBox.Text = [string] $Text
			$form.Controls.Add($textBox)
			return $textBox
		}

		$header = [System.Windows.Forms.Label]::new()
		$header.Text = 'Configure DriftMaester Azure Automation deployment'
		$header.Font = [System.Drawing.Font]::new('Segoe UI', 12, [System.Drawing.FontStyle]::Bold)
		$header.Location = [System.Drawing.Point]::new(22, 18)
		$header.Size = [System.Drawing.Size]::new(820, 30)
		$form.Controls.Add($header)

		$left = 28
		$top = 64
		$inputLeft = 270
		$row = 40
		$inputWidth = 560

		New-Label -Text 'Azure subscription *' -X $left -Y $top | Out-Null
		$subscriptionCombo = [System.Windows.Forms.ComboBox]::new()
		$subscriptionCombo.DropDownStyle = 'DropDownList'
		$subscriptionCombo.Location = [System.Drawing.Point]::new($inputLeft, $top)
		$subscriptionCombo.Size = [System.Drawing.Size]::new($inputWidth, 28)
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
		$form.Controls.Add($subscriptionCombo)

		New-Label -Text 'Resource group *' -X $left -Y ($top + $row) | Out-Null
		$resourceGroupBox = New-TextBox -X $inputLeft -Y ($top + $row) -Text $Prefilled.resourceGroup

		New-Label -Text 'Location' -X $left -Y ($top + ($row * 2)) | Out-Null
		$locationCombo = [System.Windows.Forms.ComboBox]::new()
		$locationCombo.Location = [System.Drawing.Point]::new($inputLeft, ($top + ($row * 2)))
		$locationCombo.Size = [System.Drawing.Size]::new(260, 28)
		foreach ($locationName in @('westeurope', 'northeurope', 'uksouth', 'eastus', 'westus2')) { [void] $locationCombo.Items.Add($locationName) }
		$locationCombo.Text = [string] $Prefilled.location
		$form.Controls.Add($locationCombo)

		New-Label -Text 'Frequency *' -X $left -Y ($top + ($row * 3)) | Out-Null
		$frequencyCombo = [System.Windows.Forms.ComboBox]::new()
		$frequencyCombo.DropDownStyle = 'DropDownList'
		$frequencyCombo.Location = [System.Drawing.Point]::new($inputLeft, ($top + ($row * 3)))
		$frequencyCombo.Size = [System.Drawing.Size]::new(260, 28)
		foreach ($frequencyName in @('daily', 'weekly', 'monthly')) { [void] $frequencyCombo.Items.Add($frequencyName) }
		$frequencyCombo.SelectedItem = [string] $Prefilled.frequency
		if (-not $frequencyCombo.SelectedItem) { $frequencyCombo.SelectedIndex = 0 }
		$form.Controls.Add($frequencyCombo)

		New-Label -Text 'Time of day (HH:mm) *' -X $left -Y ($top + ($row * 4)) | Out-Null
		$timeBox = New-TextBox -X $inputLeft -Y ($top + ($row * 4)) -Text $Prefilled.timeOfDay -Width 120

		New-Label -Text 'Report recipients *' -X $left -Y ($top + ($row * 5)) | Out-Null
		$recipientsBox = [System.Windows.Forms.TextBox]::new()
		$recipientsBox.Location = [System.Drawing.Point]::new($inputLeft, ($top + ($row * 5)))
		$recipientsBox.Size = [System.Drawing.Size]::new($inputWidth, 68)
		$recipientsBox.Multiline = $true
		$recipientsBox.ScrollBars = 'Vertical'
		$recipientsBox.Text = [string] $Prefilled.recipients
		$form.Controls.Add($recipientsBox)

		New-Label -Text 'Mail sender user id' -X $left -Y ($top + ($row * 7)) | Out-Null
		$senderBox = New-TextBox -X $inputLeft -Y ($top + ($row * 7)) -Text $Prefilled.senderUserId

		New-Label -Text 'Mail subject prefix' -X $left -Y ($top + ($row * 8)) | Out-Null
		$mailSubjectBox = New-TextBox -X $inputLeft -Y ($top + ($row * 8)) -Text $Prefilled.mailSubject

		New-Label -Text 'Azure DevOps organization' -X $left -Y ($top + ($row * 9)) | Out-Null
		$devOpsBox = New-TextBox -X $inputLeft -Y ($top + ($row * 9)) -Text $Prefilled.devopsOrg

		New-Label -Text 'Tenant id' -X $left -Y ($top + ($row * 10)) | Out-Null
		$tenantBox = New-TextBox -X $inputLeft -Y ($top + ($row * 10)) -Text $Prefilled.tenantId

		New-Label -Text 'Report behavior' -X $left -Y ($top + ($row * 11)) | Out-Null
		$alwaysSendReportBox = [System.Windows.Forms.CheckBox]::new()
		$alwaysSendReportBox.Text = 'Always send report, even when no drift is detected'
		$alwaysSendReportBox.Location = [System.Drawing.Point]::new($inputLeft, ($top + ($row * 11)))
		$alwaysSendReportBox.Size = [System.Drawing.Size]::new($inputWidth, 28)
		$alwaysSendReportBox.Checked = [bool] $Prefilled.alwaysSendReport
		$form.Controls.Add($alwaysSendReportBox)

		New-Label -Text 'Optional workloads' -X $left -Y ($top + ($row * 12)) | Out-Null
		$includeCopilotAndDataverseBox = [System.Windows.Forms.CheckBox]::new()
		$includeCopilotAndDataverseBox.Text = 'Include Copilot, Power Platform, Dynamics, and Dataverse checks'
		$includeCopilotAndDataverseBox.Location = [System.Drawing.Point]::new($inputLeft, ($top + ($row * 12)))
		$includeCopilotAndDataverseBox.Size = [System.Drawing.Size]::new($inputWidth, 28)
		$includeCopilotAndDataverseBox.Checked = [bool] $Prefilled.includeCopilotAndDataverse
		$form.Controls.Add($includeCopilotAndDataverseBox)

		$subscriptionCombo.Add_SelectedIndexChanged({
			if ([string]::IsNullOrWhiteSpace($tenantBox.Text) -and $subscriptionCombo.SelectedItem) {
				$tenantBox.Text = [string] $subscriptionCombo.SelectedItem.TenantId
			}
		})
		if ([string]::IsNullOrWhiteSpace($tenantBox.Text) -and $subscriptionCombo.SelectedItem) {
			$tenantBox.Text = [string] $subscriptionCombo.SelectedItem.TenantId
		}

		$cancelButton = [System.Windows.Forms.Button]::new()
		$cancelButton.Text = 'Cancel'
		$cancelButton.Location = [System.Drawing.Point]::new(590, 630)
		$cancelButton.Size = [System.Drawing.Size]::new(105, 34)
		$cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
		$form.CancelButton = $cancelButton
		$form.Controls.Add($cancelButton)

		$continueButton = [System.Windows.Forms.Button]::new()
		$continueButton.Text = 'Continue installation'
		$continueButton.Location = [System.Drawing.Point]::new(705, 630)
		$continueButton.Size = [System.Drawing.Size]::new(135, 34)
		$form.AcceptButton = $continueButton
		$form.Controls.Add($continueButton)

		$continueButton.Add_Click({
			$missing = @()
			if (-not $subscriptionCombo.SelectedItem) { $missing += 'Azure subscription' }
			if ([string]::IsNullOrWhiteSpace($resourceGroupBox.Text)) { $missing += 'Resource group' }
			if ([string]::IsNullOrWhiteSpace($frequencyCombo.Text)) { $missing += 'Frequency' }
			if ([string]::IsNullOrWhiteSpace($timeBox.Text)) { $missing += 'Time of day' }
			if ([string]::IsNullOrWhiteSpace($recipientsBox.Text)) { $missing += 'Report recipients' }
			if ($missing.Count -gt 0) {
				[System.Windows.Forms.MessageBox]::Show(('Please fill in: {0}' -f ($missing -join ', ')), 'DriftMaester Installer', 'OK', 'Warning') | Out-Null
				return
			}

			$form.Tag = @{
				subscription = [string] $subscriptionCombo.SelectedItem.Id
				resourceGroup = [string] $resourceGroupBox.Text.Trim()
				location = [string] $locationCombo.Text.Trim()
				frequency = [string] $frequencyCombo.Text.Trim()
				timeOfDay = [string] $timeBox.Text.Trim()
				recipients = [string] $recipientsBox.Text.Trim()
				senderUserId = [string] $senderBox.Text.Trim()
				devopsOrg = [string] $devOpsBox.Text.Trim()
				tenantId = [string] $tenantBox.Text.Trim()
				mailSubject = [string] $mailSubjectBox.Text.Trim()
				alwaysSendReport = [bool] $alwaysSendReportBox.Checked
				includeCopilotAndDataverse = [bool] $includeCopilotAndDataverseBox.Checked
			}
			$form.DialogResult = [System.Windows.Forms.DialogResult]::OK
			$form.Close()
		})

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

if (-not (Get-AzContext -ErrorAction SilentlyContinue)) {
	Write-InstallLog 'Connecting to Azure.'
	Connect-AzAccount | Out-Null
}

# Determine if we need GUI or CLI mode
$needsGuiInput = $GuiMode -or [string]::IsNullOrWhiteSpace($Subscription) -or [string]::IsNullOrWhiteSpace($ResourceGroup) -or [string]::IsNullOrWhiteSpace($Recipients) -or [string]::IsNullOrWhiteSpace($Frequency) -or [string]::IsNullOrWhiteSpace($TimeOfDay)

if ($needsGuiInput) {
	Write-InstallLog 'Launching GUI installer...'
	$collectedParams = Show-DriftMaesterGui -PrefilledSubscription $Subscription -PrefilledResourceGroup $ResourceGroup -PrefilledLocation $Location -PrefilledFrequency $Frequency -PrefilledTimeOfDay $TimeOfDay -PrefilledRecipients $Recipients -PrefilledSenderUserId $SenderUserId -PrefilledDevOpsOrg $DevOpsOrg -PrefilledTenantId $TenantId -PrefilledMailSubject $MailSubject -PrefilledAlwaysSendReport $AlwaysSendReport -PrefilledIncludeCopilotAndDataverse $IncludeCopilotAndDataverse

	if (-not $collectedParams) {
		Write-InstallLog 'Installation cancelled.' -Level Warning
		exit 0
	}
	
	# Use GUI-collected parameters
	$Subscription = $collectedParams.subscription
	$ResourceGroup = $collectedParams.resourceGroup
	$Location = $collectedParams.location
	$Frequency = $collectedParams.frequency
	$TimeOfDay = $collectedParams.timeOfDay
	$Recipients = $collectedParams.recipients
	$SenderUserId = $collectedParams.senderUserId
	$DevOpsOrg = $collectedParams.devopsOrg
	$TenantId = $collectedParams.tenantId
	$MailSubject = $collectedParams.mailSubject
	$AlwaysSendReport = [bool] $collectedParams.alwaysSendReport
	$IncludeCopilotAndDataverse = [bool] $collectedParams.includeCopilotAndDataverse
}

# Handle both headless mode (all parameters provided) and CLI mode (interactive fallback)
if (-not [string]::IsNullOrWhiteSpace($Subscription)) {
	Write-InstallLog 'Running installer with provided parameters.'
	
	# Validate required GUI parameters
	if ([string]::IsNullOrWhiteSpace($Subscription)) { throw 'Subscription parameter is required in GUI mode.' }
	if ([string]::IsNullOrWhiteSpace($ResourceGroup)) { throw 'ResourceGroup parameter is required in GUI mode.' }
	if ([string]::IsNullOrWhiteSpace($Recipients)) { throw 'Recipients parameter is required in GUI mode.' }
	if ([string]::IsNullOrWhiteSpace($Frequency)) { throw 'Frequency parameter is required in GUI mode.' }
	if ([string]::IsNullOrWhiteSpace($TimeOfDay)) { throw 'TimeOfDay parameter is required in GUI mode.' }

	$selectedSubscription = Get-AzSubscription -SubscriptionId $Subscription -ErrorAction Stop
	$SubscriptionId = $selectedSubscription.Id
	$ResourceGroupName = $ResourceGroup
	$DeploymentLocation = if ([string]::IsNullOrWhiteSpace($Location)) { $script:Location } else { $Location }

	# Parse recipients
	$ReportRecipient = @($Recipients -split ',' | ForEach-Object { $_.Trim() } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })

	# Parse schedule - GUI provides lowercase values
	$frequencyMap = @{ 'daily' = 'Day'; 'weekly' = 'Week'; 'monthly' = 'Month' }
	$scheduleFrequency = $frequencyMap[$Frequency.ToLower()]

	$timeParts = $TimeOfDay -split ':'
	if ($timeParts.Count -ne 2) { throw 'TimeOfDay must be in HH:mm format' }
	$timeOfDaySpan = [timespan]::new([int]$timeParts[0], [int]$timeParts[1], 0)

	# Build schedule selection based on frequency
	$invokeSchedule = switch ($scheduleFrequency) {
		'Day' {
			[PSCustomObject]@{
				Frequency       = 'Day'
				StartTime       = Get-NextDailyOccurrence -TimeOfDay $timeOfDaySpan
				TimeOfDay       = $timeOfDaySpan
				WeekDays        = @()
				MonthDays       = @()
				MonthlyDayMode  = $null
				DescriptionText = 'daily'
			}
		}
		'Week' {
			# Default to Monday if not specified
			$dayOfWeek = 'Monday'
			[PSCustomObject]@{
				Frequency       = 'Week'
				StartTime       = Get-NextWeeklyOccurrence -DayOfWeek $dayOfWeek -TimeOfDay $timeOfDaySpan
				TimeOfDay       = $timeOfDaySpan
				WeekDays        = @($dayOfWeek)
				MonthDays       = @()
				MonthlyDayMode  = $null
				DescriptionText = "weekly on $dayOfWeek"
			}
		}
		'Month' {
			[PSCustomObject]@{
				Frequency       = 'Month'
				StartTime       = Get-NextMonthlyOccurrence -DayMode 'First' -TimeOfDay $timeOfDaySpan
				TimeOfDay       = $timeOfDaySpan
				WeekDays        = @()
				MonthDays       = @(1)
				MonthlyDayMode  = 'First'
				DescriptionText = 'monthly on the 1st day'
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

	Connect-ToGraphForInstall -RequestedTenantId $TenantId
} else {
	# CLI mode - interactive prompts (original behavior)
	$selectedSubscription = Select-AzureSubscription
	$SubscriptionId = $selectedSubscription.Id
	$ResourceGroupName = Read-RequiredValue -Prompt 'Desired resource group name'
	$DeploymentLocation = Read-OptionalValue -Prompt 'Azure deployment location, press enter to use westeurope' -DefaultValue $script:Location
	$invokeSchedule = New-DriftScheduleSelection
	$updateSchedule = New-UpdateScheduleSelection -InvokeSchedule $invokeSchedule

	$recipientText = Read-RequiredValue -Prompt 'Which email addresses should receive reports, comma-separated'
	$ReportRecipient = @($recipientText -split ',' | ForEach-Object { $_.Trim() } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
	$MailSenderUserId = Read-OptionalValue -Prompt '[OPTIONAL] Mail sender UPN or user id, press enter to use first recipient'
	$DevOpsOrganization = Read-OptionalValue -Prompt 'Azure DevOps organization for Maester checks, optional'
	$TenantId = Read-OptionalValue -Prompt '[OPTIONAL] Tenant id to pass to the runbooks, press enter to use the tenant of the selected subscription'
	if ([string]::IsNullOrWhiteSpace($TenantId)) {
		$TenantId = $selectedSubscription.TenantId
	}
	$MailSubjectPrefix = Read-OptionalValue -Prompt '[OPTIONAL] Mail subject prefix for report emails, press enter to use default'
	$AlwaysSendReportEnabled = Read-YesNo -Prompt 'Always send report emails, even when no drift is detected?' -DefaultNo
	$IncludeCopilotAndDataverseEnabled = Read-YesNo -Prompt 'Include Copilot, Power Platform, Dynamics, and Dataverse checks?' -DefaultNo
	Connect-ToGraphForInstall -RequestedTenantId $TenantId
}

$recipientParameter = ($ReportRecipient | ForEach-Object { $_.Trim() } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }) -join ','
if ([string]::IsNullOrWhiteSpace($recipientParameter)) {
	throw 'At least one report recipient is required.'
}

Register-DriftAzureProvider -ProviderNamespace 'Microsoft.Automation'
Register-DriftAzureProvider -ProviderNamespace 'Microsoft.Storage'
Register-DriftAzureProvider -ProviderNamespace 'Microsoft.Authorization'

$names = Get-DriftMaesterNames -SelectedSubscriptionId $SubscriptionId -SelectedResourceGroupName $ResourceGroupName
$resourceGroup = Set-DriftResourceGroup -Name $ResourceGroupName -TargetLocation $DeploymentLocation
$resourceLocation = if ($resourceGroup -and -not [string]::IsNullOrWhiteSpace([string] $resourceGroup.Location)) { [string] $resourceGroup.Location } else { $DeploymentLocation }
$automationAccount = Set-DriftAutomationAccount -Name $names.AutomationAccountName -TargetResourceGroupName $ResourceGroupName -TargetLocation $resourceLocation
$storageAccount = Set-DriftStorageAccount -Name $names.StorageAccountName -TargetResourceGroupName $ResourceGroupName -TargetLocation $resourceLocation

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
	$principal = Get-AzADServicePrincipal -ObjectId $managedIdentityPrincipalId -ErrorAction SilentlyContinue
	if ($principal -and $principal.AppId) {
		$managedIdentityClientId = [string] $principal.AppId
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

Set-DriftRuntimeEnvironment -SelectedSubscriptionId $SubscriptionId -TargetResourceGroupName $ResourceGroupName -AutomationAccountName $names.AutomationAccountName -TargetLocation $resourceLocation | Out-Null
Set-DriftRunbookFromGithub -SelectedSubscriptionId $SubscriptionId -TargetResourceGroupName $ResourceGroupName -AutomationAccountName $names.AutomationAccountName -TargetLocation $resourceLocation -RunbookName $script:UpdateRunbookName -SourceFileName 'Update-DriftMaester.ps1'
Set-DriftRunbookFromGithub -SelectedSubscriptionId $SubscriptionId -TargetResourceGroupName $ResourceGroupName -AutomationAccountName $names.AutomationAccountName -TargetLocation $resourceLocation -RunbookName $script:InvokeRunbookName -SourceFileName 'Invoke-DriftMaester.ps1'

Set-DriftAutomationSchedule -SelectedSubscriptionId $SubscriptionId -TargetResourceGroupName $ResourceGroupName -AutomationAccountName $names.AutomationAccountName -ScheduleName $script:InvokeScheduleName -ScheduleSelection $invokeSchedule -Description 'Runs DriftMaester tenant drift detection.'
Set-DriftAutomationSchedule -SelectedSubscriptionId $SubscriptionId -TargetResourceGroupName $ResourceGroupName -AutomationAccountName $names.AutomationAccountName -ScheduleName $script:UpdateScheduleName -ScheduleSelection $updateSchedule -Description 'Updates DriftMaester runtime modules one hour before the invoke schedule.'

$invokeParameters = Get-InvokeParameters -Recipients $recipientParameter -SenderUserId $MailSenderUserId -DevOpsOrg $DevOpsOrganization -TargetTenantId $TenantId -SubjectPrefix $MailSubjectPrefix -AlwaysSendReport $AlwaysSendReportEnabled -IncludeCopilotAndDataverse $IncludeCopilotAndDataverseEnabled
Set-DriftJobSchedule -SelectedSubscriptionId $SubscriptionId -TargetResourceGroupName $ResourceGroupName -AutomationAccountName $names.AutomationAccountName -RunbookName $script:InvokeRunbookName -ScheduleName $script:InvokeScheduleName -Parameters $invokeParameters
Set-DriftJobSchedule -SelectedSubscriptionId $SubscriptionId -TargetResourceGroupName $ResourceGroupName -AutomationAccountName $names.AutomationAccountName -RunbookName $script:UpdateRunbookName -ScheduleName $script:UpdateScheduleName

Set-DriftManagedIdentityApiPermissions -ManagedIdentityClientId $managedIdentityClientId -RequestedTenantId $TenantId
$exchangeOrganization = Get-InitialTenantDomainFromGraph
Set-DriftExchangeOnlineRbac -ManagedIdentityClientId $managedIdentityClientId -ManagedIdentityObjectId $managedIdentityPrincipalId -DisplayName $names.AutomationAccountName -Organization $exchangeOrganization
$Null = Start-UpdateRunbook -SelectedSubscriptionId $SubscriptionId -TargetResourceGroupName $ResourceGroupName -AutomationAccountName $names.AutomationAccountName
Write-InstallLog 'After the Update-DriftMaester run is complete, you can manually run the Invoke-DriftMaester runbook, or wait for the next scheduled run' -Level Info
Write-InstallLog 'DriftMaester installation completed.' -Level Success