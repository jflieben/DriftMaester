# Changelog

## 1.3.0
- SharePoint Online support: the installer grants the managed identity the SharePoint `Sites.FullControl.All` application role, and the invoke runbook connects PnP.PowerShell to the tenant admin endpoint with that managed identity so the Maester SharePoint Online tests run unattended.
- Reports now lead with the number of passed tests instead of the score percentage as this is more valueable when new tests are added
- Fixed the ORCA tests failing with "Cannot find type [PolicyInfo]". Maester is now imported at script scope in the invoke runbook, because its manifest loads the ORCA class definitions through `ScriptsToProcess`, which only defines them in the scope that called `Import-Module`.
- The installer now detects an incompatible `Microsoft.Identity.Client` (MSAL) already loaded in the session by PnP.PowerShell, MicrosoftTeams or ExchangeOnlineManagement, and reports which module loaded it, instead of failing in `Connect-MgGraph` with a type load error and no sign-in prompt.

## 1.2.0
- Installer hardening: storage security baseline, lifecycle retention policy, root-elevation cleanup by object id, access report output, and optional RunNow workflow.
- Invoke runbook reliability: token refresh before post-processing, report-delivery modes, severity gating, failure notifications, retention cleanup, summary blobs for trend reads, and optional Teams webhook notifications.
- Update runbook reliability: target runtime resolution now stays scoped to driftmaester runtime and adds Maester dependency compatibility auto-bump behavior.
- Added Remove-DriftMaester uninstaller script for full or scoped cleanup.
