<#
.SYNOPSIS
Shared DriftMaester permission and role definitions.

.DESCRIPTION
Single source of truth for the API permissions and Entra ID directory roles that the DriftMaester
managed identity needs:

- Microsoft Graph application permissions (app roles granted on the Graph service principal).
- SharePoint Online application permissions (app roles on the 'Office 365 SharePoint Online' SP).
- Exchange Online application permission (app role on the 'Office 365 Exchange Online' SP).
- Entra ID directory roles assigned to the managed identity's service principal.

Both Install-DriftMaester.ps1 and Update-DriftMaesterPermissions.ps1 load these definitions so the two
scripts can never drift apart. When this script is executed (dot-sourced, invoked, or fetched from GitHub
and run through Invoke-Expression / a scriptblock) it emits a single hashtable describing every requirement.

.NOTES
Author: Jos Lieben / Lieben Consultancy
Website: https://www.lieben.nu
Free for non-commercial use. Commercial use requires a license:
https://www.lieben.nu/liebensraum/commercial-use/
#>

@{
	# Well-known first-party service principals the managed identity receives app roles on.
	GraphAppId            = '00000003-0000-0000-c000-000000000000'
	ExchangeOnlineAppId   = '00000002-0000-0ff1-ce00-000000000000'
	SharePointOnlineAppId = '00000003-0000-0ff1-ce00-000000000000'

	# Microsoft Graph application permissions (app roles) granted on the Graph service principal.
	GraphApplicationPermissions = @(
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
		'RoleEligibilitySchedule.Read.Directory',
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

	# Granted on the 'Office 365 SharePoint Online' service principal (not Graph). PnP.PowerShell connects to
	# the SharePoint tenant admin endpoint app-only with the managed identity, and Get-PnPTenant (used by the
	# Maester SharePoint Online tests) is only authorised by Sites.FullControl.All. Lower roles return 401.
	SharePointApplicationPermissions = @(
		'Sites.FullControl.All'
	)

	# Granted on the 'Office 365 Exchange Online' service principal so the managed identity can run the Maester
	# Exchange configuration tests app-only.
	ExchangeApplicationPermissions = @(
		'Exchange.ManageAsApp'
	)

	# Entra ID directory roles assigned to the managed identity's service principal. 'Azure DevOps
	# Administrator' is only assigned when an Azure DevOps organization is configured.
	DirectoryRoles = @(
		'Global Reader',
		'Teams Reader',
		'Azure DevOps Administrator'
	)
}
