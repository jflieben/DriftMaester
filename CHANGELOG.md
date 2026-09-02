# Changelog

## 1.4.0
- The invoke runbook now writes a machine-readable NUnit test-results XML to blob storage on every run, alongside the existing JSON/HTML/Markdown reports. The newest results are always available at a fixed path (`<prefix>/<tenantId>/latest/TestResults.xml`) so CI systems can fetch them reliably.
- Added an Azure DevOps integration: a sample pipeline (`Pipelines/azure-pipelines-driftmaester.yml`) downloads the latest results XML from blob, fails if it is older than a configurable freshness threshold, and publishes it to the Azure DevOps Tests tab. Drift surfaces through Azure DevOps' own run-to-run test history. See `docs/azure-devops-pipeline.md`.
- No extra cloud resources or state stores are needed: the runbook keeps producing and storing everything in Azure Automation + blob, the pipeline only presents the results.
- Pinned Maester module at version 2.2.0 to prevent breakage by major updates

## 1.3.0
- SharePoint Online support: the installer grants the managed identity the SharePoint `Sites.FullControl.All` application role, and the invoke runbook connects PnP.PowerShell to the tenant admin endpoint with that managed identity so the Maester SharePoint Online tests run unattended.
- Reports now lead with the number of passed tests instead of the score percentage as this is more valueable when new tests are added
- Fixed the ORCA tests failing with "Cannot find type [PolicyInfo]". Maester is now imported at script scope in the invoke runbook, because its manifest loads the ORCA class definitions through `ScriptsToProcess`, which only defines them in the scope that called `Import-Module`.
- The installer now reuses the Azure sign-in for Microsoft Graph by passing an Az-issued Graph token to `Connect-MgGraph`, so the admin signs in once instead of twice. 
- Centralized permissions and added a permissions reconciliation script (Update-DriftMaesterPermissions.ps1) that automatically adds new permissions and can be executed lightweight in Azure Shell

## 1.2.0
- Installer hardening: storage security baseline, lifecycle retention policy, root-elevation cleanup by object id, access report output, and optional RunNow workflow.
- Invoke runbook reliability: token refresh before post-processing, report-delivery modes, severity gating, failure notifications, retention cleanup, summary blobs for trend reads, and optional Teams webhook notifications.
- Update runbook reliability: target runtime resolution now stays scoped to driftmaester runtime and adds Maester dependency compatibility auto-bump behavior.
- Added Remove-DriftMaester uninstaller script for full or scoped cleanup.
