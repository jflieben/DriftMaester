# DriftMaester

DriftMaester is a plug-and-play Azure Automation implementation of Maester with drift detection, trend tracking, and email reporting.

It helps the (slightly?) less tech savvy or time-constrained IT admins to not only run Maester, but to run it regularly and get alerted when their compliance drifts.

![DriftMaester HTML report preview](Media/htmlreport.png)

## Quick install:

```powershell
iex ((Invoke-WebRequest -UseBasicParsing 'https://raw.githubusercontent.com/jflieben/DriftMaester/main/Install-DriftMaester.ps1').Content)
```

## What DriftMaester does

- Automatically deploys an Azure Automation setup (resource group, automation account, runtime, runbooks, storage).
- Sets correct authorizations in Graph, Azure and Exchange
- Runs Maester checks on a schedule (daily, weekly, or monthly).
- Stores result history in secure Azure Blob Storage.
- Detects configuration drift and highlights regressions each run
- Sends HTML reports by email.
- Keeps its runtime modules and package set updated using the update runbook, so each run has the latest tests / Maester version.

## Repository structure

- `Install-DriftMaester.ps1`: interactive installer and updater for Azure resources and permissions.
- `Runbooks/Invoke-DriftMaester.ps1`: main runbook that executes Maester, compares results, and mails reports.
- `Runbooks/Update-DriftMaester.ps1`: runtime/package maintenance runbook.

## High-level architecture

```mermaid
flowchart LR
		A[Operator runs Install-DriftMaester.ps1] --> B[Azure Resource Group]
		B --> C[Automation Account + Managed Identity]
		C --> D[Runtime Environment driftmaester]
		C --> E[Runbooks: Update + Invoke]
		C --> F[Schedules + Job Schedules]
		C --> G[Storage Account + maester container]
		E --> H[Update runbook installs required modules/packages]
		E --> I[Invoke runbook executes Maester checks]
		I --> G
		I --> J[HTML mail report to recipients]
```

## Processing flow

```mermaid
flowchart LR
    A[Install-DriftMaester.ps1] --> B[Automation Account + Managed Identity]
    B --> C[Update Runbook]
    B --> D[Invoke Runbook]
    D --> E[Blob Storage History]
    D --> F[Email Full & Drift Report]

    classDef setup fill:#dae8fc,stroke:#6c8ebf,color:#102a43,stroke-width:1px
    classDef runbooks fill:#fff2cc,stroke:#d6b656,color:#5f4b1b,stroke-width:1px
    classDef output fill:#f8cecc,stroke:#b85450,color:#4a1c1c,stroke-width:1px
    classDef report fill:#e1d5e7,stroke:#9673a6,color:#3f2a4d,stroke-width:1px

    class A,B setup
    class C,D runbooks
    class E output
    class F report
```

## Requirements

### Local requirements (for running installer)

- PowerShell 7+
- Az PowerShell modules installed locally:
	- `Az.Accounts`
	- `Az.Resources`
	- `Az.Automation`
	- `Az.Storage`
- Microsoft Graph PowerShell module:
	- `Microsoft.Graph.Authentication`
- Exchange Online module:
	- `ExchangeOnlineManagement`

Local requirements are automatically handled by the installer. If a required module is missing, `Install-DriftMaester.ps1` installs it automatically.

### Azure and Entra permissions

You should run install with a Global Administator OR an account that can:

- Create/update resource groups and Azure Automation resources.
- Assign RBAC roles in the target resource group.
- Assign root-scope Reader roles to the managed identity (`/` and `/providers/Microsoft.aadiam`).
- Grant Microsoft Graph app roles and directory roles to the managed identity.
- Create Exchange Online service principal and management role assignment.

For root-scope assignment, the installer attempts normal assignment first. If denied, it temporarily elevates access using:

- `POST /providers/Microsoft.Authorization/elevateAccess?api-version=2015-07-01`

and automatically removes the temporary elevation assignment afterward.

### Runtime requirements (inside Azure Automation)

The runtime environment is configured with default packages and then maintained by the update runbook.

### Optional workload prerequisites

Some checks need additional access outside the Azure resources and Entra permissions configured by the installer. Configure these before enabling the related scans if you wish to include them

**Copilot Studio, Power Platform, Dynamics, and Dataverse**

Dataverse access is required for the Copilot Studio security tests that evaluate Copilot Studio agent configurations.

Create an application user in Power Platform:

1. Go to the Power Platform Admin Center.
2. Select the environment you want DriftMaester to evaluate.
3. Go to **Settings** > **Users + permissions** > **Application users**.
4. Click **New app user** > **Add an app**.
5. Search for `driftmaester` and select it.
6. Select the correct business unit.
7. Assign a security role with read access:
	- `Basic User` for simplicity, or
	- A custom role, for example `Maester Security Reader`, with organization-level read access on `Agent (bot)`, `Agent component (botcomponent)`, `User (systemuser)`, and `Connection Reference (connectionreference)`.
8. Click **Create**.

After this is configured, enable the workload in the installer with `-IncludeCopilotAndDataverse $true`, or tick the matching checkbox in the GUI. If you've already installed you can rerun the installer or you can modify the scheduled job in Azure directly.

**Azure DevOps**

For Azure DevOps scans, the DriftMaester managed identity must be added to the Azure DevOps organization at `https://dev.azure.com/ORGNAMEHERE/_settings/users` and granted access to every project that should be scanned.

For maximum test coverage, also add the DriftMaester managed identity to **Security** > **Permissions** > **Project Collection Administrators**.

## Installation

### Direct from GitHub

Copy/paste ready command to run directly from GitHub:

```powershell
iex ((Invoke-WebRequest -UseBasicParsing 'https://raw.githubusercontent.com/jflieben/DriftMaester/main/Install-DriftMaester.ps1').Content)
```

3. Follow prompts in order:
	 - Select Azure subscription.
	 - Enter resource group name.
	 - Enter deployment location (press Enter for `westeurope`).
	 - Choose run frequency, run time, and time zone.
	 - Enter report recipients and optional sender/org/tenant/subject/report behavior/workload settings.
4. Approve sign-in prompts for Azure, Graph, and Exchange Online when requested.


### Manual, GUI mode (Windows, recommended)

Runs automatically if any required parameter is missing:

```powershell
.\Install-DriftMaester.ps1
```

The native PowerShell form opens on Windows:
- Select your Azure subscription
- Enter resource group name, location
- Choose run frequency (daily, weekly, monthly), time, and time zone
- Enter email recipients and optional settings
- Choose whether reports are sent only when drift is detected, or after every run
- Choose whether Copilot, Power Platform, Dynamics, and Dataverse checks are included

Installer preview:

![DriftMaester installer](Media/installer.png)

On non-Windows platforms, provide all required parameters or use the console prompts.

### CLI/Headless Mode (Scripting & Automation)

For scripting or when you have all parameters ready, supply them all:

```powershell
.\Install-DriftMaester.ps1 -Subscription "sub-id" -ResourceGroup "rg-drift" `
	-Recipients "admin@company.com" -Frequency "daily" -TimeOfDay "02:00" -TimeZone "W. Europe Standard Time" `
	-AlwaysSendReport $true -IncludeCopilotAndDataverse $true
```

The installer runs mostly silently except for logins to Graph and Exchange if you're not already logged in.

`-AlwaysSendReport $true` passes `AlwaysSendReport` to the scheduled `Invoke-DriftMaester` runbook. By default this is `$false`, which means DriftMaester sends mail on the first run and when drift is detected. Set it to `$true` when you want a report after every scheduled run, even if nothing changed.

`-TimeZone "W. Europe Standard Time"` sets the time zone used by the Azure Automation schedules. If omitted, the installer uses the local system time zone. Use the time zone id expected by Azure Automation, for example `W. Europe Standard Time`.

`-IncludeCopilotAndDataverse $true` passes `includeCopilotAndDataverse` to the scheduled `Invoke-DriftMaester` runbook. By default this is `$false`, which skips Copilot Studio, Power Platform, Dynamics, and Dataverse-backed Maester checks and suppresses Dataverse connection warnings. Set it to `$true` when the managed identity has access to the Dataverse environment and you want those checks included.

**Mixed Mode** (Partial Parameters)

Provide some parameters and use GUI for the rest:

```powershell
.\Install-DriftMaester.ps1 -Subscription "sub-id" -Recipients "admin@company.com"
```

The GUI launches with those fields pre-populated.

### What the installer creates/configures

- Resource group (if missing)
- Automation account with system-assigned managed identity
- Storage account and `maester` blob container
- Runtime environment `driftmaester`
- Runbooks `Invoke-DriftMaester` and `Update-DriftMaester`
- Schedules and job schedules
- Azure RBAC role assignments for managed identity
- Root-scope Reader role assignments for managed identity
- Graph API app-role assignments and directory roles
- Exchange Online service principal + `View-Only Configuration` assignment

This follows official documentation at [Maester](https://maester.dev/docs/connect-maester/connect-maester-advanced/)

## Usage

### Normal operations

- DriftMaester runs automatically on its configured schedule.
- Update runbook runs one hour before invoke runbook.
- Invoke runbook stores results and sends report emails.

Report preview:

Email summary received after a run:

![DriftMaester email report preview](Media/emailreport.png)

### Manual run

You can start either runbook manually from Azure Automation:

- `Update-DriftMaester` first (optional but recommended)
- `Invoke-DriftMaester`

### Re-running installation

Re-running `Install-DriftMaester.ps1` is supported and intended.

Idempotent behavior includes:

- Reuse existing resources where possible.
- Reapply runbook content from this repository.
- Recreate deterministic schedules/job schedules.
- Skip role assignments and permissions already present.

## Troubleshooting

### Runtime shows packages as Not Configured

- Re-run installer to reset everything
- Run `Update-DriftMaester` manually and check runbook output.

### Root RBAC assignment fails

- Ensure account is Global Administrator and elevation is allowed.

### Graph/Exchange permission or auth failures

- Use an account with enough admin rights for consent, app role assignment, and EXO RBAC.
- Verify tenant context in prompts before confirming.

## Security notes

- Managed identity is used for runbook runtime auth.
- Results are stored in your own Azure Storage account.
- Email is sent through Graph API using configured sender behavior.
- Review role assignments and granted app permissions before production use.

## Licensing and links

Free for non-commercial use.

Although I've contributed to the open source Maester project, I am not directly affiliated with it and this tool only serves as a wrapper / feature add to make your life easier :)

Commercial use requires a license: https://www.lieben.nu/liebensraum/commercial-use/

- Site: https://www.lieben.nu
- Blog: https://www.lieben.nu/liebensraum/
- Author: Jos Lieben (Lieben Consultancy)
