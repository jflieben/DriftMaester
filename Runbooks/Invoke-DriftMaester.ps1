<#
.SYNOPSIS
Runs Maester from Azure Automation with managed identity, stores results in Azure Blob Storage, detects drift, and emails an HTML report.

.DESCRIPTION
This runbook is designed for Azure Automation PowerShell 7.6+ It authenticates with the Automation Account managed identity,
updates the Maester test set, validates required Microsoft Graph application permissions, connects to the services Maester can
use from automation, runs the full Maester test suite, stores native Maester result files in the configured storage account,
compares the current run with the previous stored run, builds a trend over previous runs, and sends a single HTML report.

It also automatically updates Maester and all Maester tests before each run.

Use Install-DriftMaester to automatically install Maester into a resource group of your choosing.

.PARAMETER ReportRecipient
Required. One or more recipient email addresses for run results and failure notifications.
Use a comma-separated string when you want to notify multiple people.

.PARAMETER MailSenderUserId
Optional. Mailbox UPN or user id used as the Graph sender in /users/{id}/sendMail.
Set this when the report must come from a dedicated shared mailbox; otherwise the first ReportRecipient is used.

.PARAMETER devOpsOrganization
Optional. Azure DevOps organization name used for Maester Azure DevOps connection checks.
Set this only when you want Azure DevOps drift and permission coverage in the report.

.PARAMETER TenantId
Optional. Tenant id to target explicitly for Azure and Graph token acquisition.
Set this in multi-tenant or cross-tenant automation scenarios to avoid ambiguity.

.PARAMETER MailSubjectPrefix
Optional. Prefix added to all mail subjects (default: Maester report).
Set this when you want environment or customer context in inboxes, such as PROD or a tenant short name.

.PARAMETER AlwaysSendReport
Optional. Boolean flag (default: $false). When $true, sends a report even if no drift is detected.
When $false, sends a report only on first run (no prior history) or when drift is detected.

.PARAMETER includeCopilotAndDataverse
Optional. Boolean flag (default: $false). When $true, includes Power Platform / Copilot / Dynamics scanning in Maester tests.
When $false, Dataverse connection failures are not reported as warnings since the services are not being scanned.

.EXAMPLE
./Invoke-DriftMaester.ps1 -ReportRecipient "security@contoso.com,platform@contoso.com" -MailSenderUserId "maester-reports@contoso.com" -MailSubjectPrefix "PROD Maester" -AlwaysSendReport $true -includeCopilotAndDataverse $true

.NOTES
Author: Jos Lieben / Lieben Consultancy
Website: https://www.lieben.nu
Blog: https://www.lieben.nu/liebensraum/
Free for non-commercial use. Commercial use requires a license:
https://www.lieben.nu/liebensraum/commercial-use/

#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string] $ReportRecipient = "",

    [Parameter(Mandatory = $false)]
    [string] $MailSenderUserId,

    [Parameter(Mandatory = $false)]
    [string] $devOpsOrganization,    

    [Parameter(Mandatory = $false)]
    [string] $TenantId,

    [Parameter(Mandatory = $false)]
    [string] $MailSubjectPrefix = 'Maester report',

    [Parameter(Mandatory = $false)]
    [bool] $AlwaysSendReport = $false,

    [Parameter(Mandatory = $false)]
    [bool] $includeCopilotAndDataverse = $false
)

[string] $AzureEnvironment = 'AzureCloud'
[string] $GraphEnvironment = 'Global'
[string] $ExchangeEnvironmentName = 'O365Default'
[string] $ResultsContainerName = 'maester'
[string] $BlobPrefix = 'maester'
[int] $TrendRunCount = 10
$ErrorActionPreference = 'Stop'
$script:DetectedStorageAccountName = $null
$script:DetectedStorageResourceGroupName = $null
$script:DetectedStorageSubscriptionId = $null
$script:GraphAppRoles = @()
$script:ConnectedManagedIdentityClientId = $null
$script:ConnectedTenantId = $null

function Import-RequiredModule {
    param([Parameter(Mandatory = $true)][string] $Name)

    Write-RunLog "Importing module '$Name'."
    $Null = Import-Module $Name 4>&1 | Where-Object { $_ -isnot [System.Management.Automation.VerboseRecord] }
}

function Get-InstalledMaesterModuleVersion {
    $module = Get-Module -Name Maester | Sort-Object Version -Descending | Select-Object -First 1
    if (-not $module) {
        $module = Get-Module -ListAvailable -Name Maester | Sort-Object Version -Descending | Select-Object -First 1
    }

    if (-not $module) {
        Write-RunLog "Could not determine the installed Maester module version." -Level Warning
        return 'Unknown'
    }

    $version = if ($module.Version) { $module.Version.ToString() } else { 'Unknown' }
    $prerelease = $module.PrivateData.PSData.Prerelease
    if (-not [string]::IsNullOrWhiteSpace($prerelease)) {
        return "$version-$prerelease"
    }

    return $version
}

function ConvertTo-PlainTextFromSecureString {
    param([Parameter(Mandatory = $true)][securestring] $SecureString)

    try {
        return ConvertFrom-SecureString -SecureString $SecureString -AsPlainText -Force
    } catch {
        return ConvertFrom-SecureString -SecureString $SecureString -AsPlainText
    }
}

function Get-MaesterCloudResourceUrls {
    $graphResource = switch ($GraphEnvironment) {
        'China' { 'https://microsoftgraph.chinacloudapi.cn' }
        'USGov' { 'https://graph.microsoft.us' }
        'USGovDoD' { 'https://graph.microsoft.us' }
        'Germany' { 'https://graph.microsoft.de' }
        default { 'https://graph.microsoft.com' }
    }

    $exchangeResource = switch ($ExchangeEnvironmentName) {
        'O365China' { 'https://partner.outlook.cn' }
        'O365USGovDoD' { 'https://outlook.office365.us' }
        'O365USGovGCCHigh' { 'https://outlook.office365.us' }
        'O365GermanyCloud' { 'https://outlook.office.de' }
        default { 'https://outlook.office365.com' }
    }

    $ippsResource = switch ($ExchangeEnvironmentName) {
        'O365China' { 'https://ps.compliance.protection.partner.outlook.cn' }
        'O365USGovDoD' { 'https://ps.compliance.protection.office365.us' }
        'O365USGovGCCHigh' { 'https://ps.compliance.protection.office365.us' }
        'O365GermanyCloud' { 'https://ps.compliance.protection.outlook.de' }
        default { 'https://ps.compliance.protection.outlook.com' }
    }

    [PSCustomObject]@{
        Graph          = $graphResource
        ExchangeOnline = $exchangeResource
        IPPS           = $ippsResource
        Teams          = '48ac35b8-9aa8-4d74-927d-1f4a14a0b239'
    }
}

function Get-AzAccessTokenForResource {
    param([Parameter(Mandatory = $true)][string] $ResourceUrl)

    $tokenParams = @{
        ResourceUrl    = $ResourceUrl
        AsSecureString = $true
    }
    if ($TenantId) { $tokenParams['TenantId'] = $TenantId }
    Get-AzAccessToken @tokenParams
}

function Get-JwtPayload {
    param([Parameter(Mandatory = $true)][string] $Token)

    $parts = $Token.Split('.')
    if ($parts.Count -lt 2) { return $null }

    $payload = $parts[1].Replace('-', '+').Replace('_', '/')
    switch ($payload.Length % 4) {
        2 { $payload += '==' }
        3 { $payload += '=' }
        1 { $payload += '===' }
    }

    [Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($payload)) | ConvertFrom-Json
}

function Write-RunLog {
    param(
        [Parameter(Mandatory = $true)]
        [string] $Message,

        [Parameter(Mandatory = $false)]
        [ValidateSet('Info', 'Warning', 'Error', 'Success')]
        [string] $Level = 'Info'
    )

    $prefix = "[{0:u}] [{1}]" -f (Get-Date), $Level.ToUpperInvariant()
    switch ($Level) {
        'Warning' { Write-Warning "$prefix $Message" }
        'Error' { Write-Error "$prefix $Message" -ErrorAction Continue }
        default { Write-Verbose "$prefix $Message" }
    }
}

function ConvertTo-HtmlEncodedText {
    param([AllowNull()][object] $Value)

    if ($null -eq $Value) {
        return ''
    }

    return [System.Net.WebUtility]::HtmlEncode([string] $Value)
}

function Connect-RunbookIdentity {
    Write-RunLog "Connecting to Azure with managed identity."
    $azParams = @{
        Identity              = $true
        Environment           = $AzureEnvironment
        SkipContextPopulation = $true
    }
    if ($TenantId) { $azParams['Tenant'] = $TenantId }
    Connect-AzAccount @azParams | Out-Null

    Write-RunLog "Connecting to Microsoft Graph with an Az-issued managed identity access token."
    $resourceUrls = Get-MaesterCloudResourceUrls
    $graphToken = Get-AzAccessTokenForResource -ResourceUrl $resourceUrls.Graph
    $graphTokenPlain = ConvertTo-PlainTextFromSecureString -SecureString $graphToken.Token
    $graphPayload = Get-JwtPayload -Token $graphTokenPlain
    $script:GraphAppRoles = @($graphPayload.roles) | Sort-Object -Unique
    $script:ConnectedManagedIdentityClientId = if ($graphPayload.appid) { $graphPayload.appid } else { $graphPayload.azp }
    $script:ConnectedTenantId = if ($TenantId) { $TenantId } else { $graphPayload.tid }

    $graphParams = @{ NoWelcome = $true }
    if ($GraphEnvironment -ne 'Global') { $graphParams['Environment'] = $GraphEnvironment }
    Connect-MgGraph -AccessToken $graphToken.Token @graphParams | Out-Null

    $context = Get-MgContext
    if (-not $context) {
        throw 'Microsoft Graph context was not created after Connect-MgGraph -AccessToken.'
    }

    if ([string]::IsNullOrWhiteSpace($script:ConnectedManagedIdentityClientId)) { $script:ConnectedManagedIdentityClientId = $context.ClientId }
    if ([string]::IsNullOrWhiteSpace($script:ConnectedTenantId)) { $script:ConnectedTenantId = $context.TenantId }

    Write-RunLog "Graph connected as app '$($script:ConnectedManagedIdentityClientId)' in tenant '$($script:ConnectedTenantId)'."
    return $context
}

function Get-InitialTenantDomain {
    $graphRoot = (Get-MaesterCloudResourceUrls).Graph.TrimEnd('/')
    $domains = Invoke-MgGraphRequest -Method GET -Uri "$graphRoot/v1.0/domains?`$select=id,isInitial"
    $initialDomain = @($domains.value | Where-Object { $_.isInitial } | Select-Object -First 1).id
    if ([string]::IsNullOrWhiteSpace($initialDomain)) {
        throw 'Could not detect the tenant initial domain from Microsoft Graph /domains.'
    }

    return $initialDomain
}

function Get-RequiredGraphPermissions {
    $scopeParams = @{}
    $scopeParams['Privileged'] = $true
    $required = @(Get-MtGraphScope @scopeParams)
    $required += 'Mail.Send'
    $required | Sort-Object -Unique
}

function Test-GraphPermissions {
    param(
        [Parameter(Mandatory = $true)]
        [string[]] $RequiredPermissions
    )

    $context = Get-MgContext
    if (-not $context) {
        throw 'Cannot validate Graph permissions because there is no active Graph context.'
    }

    $currentPermissions = @($script:GraphAppRoles + $context.Scopes) | Sort-Object -Unique
    $missing = foreach ($permission in $RequiredPermissions) {
        $readWriteEquivalent = $permission -replace '\.Read\.', '.ReadWrite.'
        if ($currentPermissions -notcontains $permission -and $currentPermissions -notcontains $readWriteEquivalent) {
            [PSCustomObject]@{
                Service    = 'Microsoft Graph'
                Permission = $permission
                Type       = 'Application'
                Reason     = 'Required by Maester tests or by this runbook to send the report.'
            }
        }
    }

    return @($missing)
}

function Connect-OptionalMaesterServices {
    $missingServices = [System.Collections.Generic.List[object]]::new()
    $resourceUrls = Get-MaesterCloudResourceUrls
    $graphContext = Get-MgContext
    $tenantForExchange = if ($TenantId) { $TenantId } elseif ($script:ConnectedTenantId) { $script:ConnectedTenantId } else { $graphContext.TenantId }
    $appId = if ($script:ConnectedManagedIdentityClientId) { $script:ConnectedManagedIdentityClientId } else { $graphContext.ClientId }
    $initialDomain = $null

    try {
        $initialDomain = Get-InitialTenantDomain
    } catch {
        $missingServices.Add([PSCustomObject]@{
                Service    = 'Exchange Online'
                Permission = 'Tenant initial domain / MOERA'
                Type       = 'Configuration'
                Reason     = $_.Exception.Message
            })
    }

    if (-not [string]::IsNullOrWhiteSpace($tenantForExchange)) {
        try {
            Write-RunLog "Connecting to Exchange Online with an Az-issued access token."
            $outlookToken = ConvertTo-PlainTextFromSecureString -SecureString ((Get-AzAccessTokenForResource -ResourceUrl $resourceUrls.ExchangeOnline).Token)
            $exoParams = @{
                AccessToken             = $outlookToken
                AppId                   = $appId
                Organization            = $tenantForExchange
                ExchangeEnvironmentName = $ExchangeEnvironmentName
                ShowBanner              = $false
            }
            Connect-ExchangeOnline @exoParams | Out-Null
        } catch {
            $missingServices.Add([PSCustomObject]@{
                    Service    = 'Exchange Online'
                    Permission = 'Exchange.ManageAsApp plus an Exchange admin role assignment'
                    Type       = 'Application/RBAC'
                    Reason     = "Exchange Online managed identity connection failed: $($_.Exception.Message)"
                })
        }
    }

    if (-not [string]::IsNullOrWhiteSpace($initialDomain)) {
        try {
            Write-RunLog "Connecting to Security & Compliance PowerShell with an Az-issued access token."
            $ippsToken = ConvertTo-PlainTextFromSecureString -SecureString ((Get-AzAccessTokenForResource -ResourceUrl $resourceUrls.IPPS).Token)
            $ippsParams = @{
                AccessToken  = $ippsToken
                Organization = $initialDomain
                ShowBanner   = $false
            }
            Connect-IPPSSession @ippsParams | Out-Null
        } catch {
            $missingServices.Add([PSCustomObject]@{
                    Service    = 'Security & Compliance PowerShell'
                    Permission = 'Exchange.ManageAsApp plus the required compliance/exchange role assignments'
                    Type       = 'Application/RBAC'
                    Reason     = "Security & Compliance managed identity connection failed: $($_.Exception.Message)"
                })
        }
    }

    try {
        Write-RunLog "Connecting to Microsoft Teams with Az-issued Graph and Teams access tokens."
        $teamsGraphToken = ConvertTo-PlainTextFromSecureString -SecureString ((Get-AzAccessTokenForResource -ResourceUrl $resourceUrls.Graph).Token)
        $teamsToken = ConvertTo-PlainTextFromSecureString -SecureString ((Get-AzAccessTokenForResource -ResourceUrl $resourceUrls.Teams).Token)
        $teamsParams = @{
            AccessTokens = @($teamsGraphToken, $teamsToken)
        }
        if ($TeamsEnvironmentName) { $teamsParams['TeamsEnvironmentName'] = $TeamsEnvironmentName }
        Connect-MicrosoftTeams @teamsParams | Out-Null
    } catch {
        $missingServices.Add([PSCustomObject]@{
                Service    = 'Microsoft Teams PowerShell'
                Permission = 'Teams PowerShell application permissions and role assignments supported by the MicrosoftTeams module'
                Type       = 'Application/RBAC'
                Reason     = "Teams managed identity connection failed: $($_.Exception.Message)"
            })
    }

    if ($includeCopilotAndDataverse) {
        try {
            Write-RunLog "Connecting Maester to Dataverse for Copilot Studio agent security tests."
            $dataverseParams = @{
                Service              = 'Dataverse'
                AzureEnvironment     = $AzureEnvironment
                Environment          = $GraphEnvironment
                ExchangeEnvironmentName = $ExchangeEnvironmentName
            }
            if ($TenantId) { $dataverseParams['TenantId'] = $TenantId }
            Connect-Maester @dataverseParams | Out-Null
            $maesterSession = Get-MtSession
            if (-not $maesterSession -or [string]::IsNullOrWhiteSpace([string] $maesterSession.DataverseApiBase)) {
                $missingServices.Add([PSCustomObject]@{
                        Service    = 'Dataverse / Copilot Studio'
                        Permission = 'Dataverse environment access for the managed identity'
                        Type       = 'Application/RBAC'
                        Reason     = 'Dataverse environment was not resolved or no Dataverse token could be acquired. Copilot Studio agent security tests will be skipped or report no agent data.'
                    })
            }
        } catch {
            $missingServices.Add([PSCustomObject]@{
                    Service    = 'Dataverse / Copilot Studio'
                    Permission = 'Dataverse environment access for the managed identity'
                    Type       = 'Application/RBAC'
                    Reason     = "Dataverse managed identity connection failed: $($_.Exception.Message)"
                })
        }
    } else {
        Write-RunLog "Copilot Studio and Dataverse scanning not enabled (includeCopilotAndDataverse is false), skipping Dataverse connection and tests."
    }

    try {
        if($devOpsOrganization){
            Write-RunLog "Connecting Maester to Azure DevOps for pipeline drift tests."
            Connect-ADOPS -Organization $devOpsOrganization -ManagedIdentity
            if (-not (Test-MtConnection -Service AzureDevOps -ErrorAction SilentlyContinue)) {
                $missingServices.Add([PSCustomObject]@{
                        Service    = 'Azure DevOps'
                        Permission = 'Azure DevOps connection'
                        Type       = 'Delegated/External'
                        Reason     = 'Azure DevOps is not connected. Maester Azure DevOps tests require an Azure DevOps connection and will be skipped.'
                    })
            }
        } else {
            Write-RunLog "Azure DevOps organization not specified, skipping Azure DevOps connection and tests."
        }
    } catch {
        $missingServices.Add([PSCustomObject]@{
                Service    = 'Azure DevOps'
                Permission = 'Azure DevOps connection'
                Type       = 'Delegated/External'
                Reason     = "Azure DevOps connection check failed: $($_.Exception.Message)"
            })
    }

    return @($missingServices)
}

function Send-DriftMail {
    param(
        [Parameter(Mandatory = $true)]
        [string] $Subject,

        [Parameter(Mandatory = $true)]
        [string] $HtmlBody,

        [Parameter(Mandatory = $false)]
        [string] $AttachmentPath
    )

    $reportRecipients = @($ReportRecipient -split ',' | ForEach-Object { $_.Trim() } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    if ($reportRecipients.Count -eq 0) {
        throw 'Cannot send mail because ReportRecipient is empty. Provide one or more comma-separated email addresses.'
    }

    $senderUserId = if ($MailSenderUserId) { $MailSenderUserId } else { $reportRecipients[0] }

    $attachments = @()
    if ($AttachmentPath -and (Test-Path -Path $AttachmentPath -PathType Leaf)) {
        $bytes = [System.IO.File]::ReadAllBytes($AttachmentPath)
        $attachments += @{
            '@odata.type' = '#microsoft.graph.fileAttachment'
            name          = [System.IO.Path]::GetFileName($AttachmentPath)
            contentType   = 'text/html'
            contentBytes  = [System.Convert]::ToBase64String($bytes)
        }
    }

    $message = @{
        message         = @{
            subject      = $Subject
            body         = @{
                contentType = 'HTML'
                content     = $HtmlBody
            }
            toRecipients = @($reportRecipients | ForEach-Object {
                    @{
                        emailAddress = @{
                            address = $_
                        }
                    }
                })
        }
        saveToSentItems = $false
    }

    if ($attachments.Count -gt 0) {
        $message.message['attachments'] = $attachments
    }

    $graphRoot = (Get-MaesterCloudResourceUrls).Graph.TrimEnd('/')
    $sendMailUri = "$graphRoot/v1.0/users/$senderUserId/sendMail"
    Invoke-MgGraphRequest -Method POST -Uri $sendMailUri -Body ($message | ConvertTo-Json -Depth 12) -ContentType 'application/json' | Out-Null
}

function New-MissingPermissionReportHtml {
    param(
        [Parameter(Mandatory = $true)]
        [object[]] $MissingItems,

        [Parameter(Mandatory = $false)]
        [object] $GraphContext
    )

    $rows = foreach ($item in $MissingItems) {
        '<tr><td>{0}</td><td>{1}</td><td>{2}</td><td>{3}</td></tr>' -f `
            (ConvertTo-HtmlEncodedText $item.Service),
            (ConvertTo-HtmlEncodedText $item.Permission),
            (ConvertTo-HtmlEncodedText $item.Type),
            (ConvertTo-HtmlEncodedText $item.Reason)
    }

    $clientId = if ($GraphContext) { $GraphContext.ClientId } else { '' }
    $tenant = if ($GraphContext) { $GraphContext.TenantId } else { $TenantId }
    $permissionList = ($MissingItems | ForEach-Object { "<li><strong>$(ConvertTo-HtmlEncodedText $_.Permission)</strong> <span>($(ConvertTo-HtmlEncodedText $_.Service))</span></li>" }) -join [Environment]::NewLine

    return @"
<!doctype html>
<html>
<head>
<meta charset="utf-8">
<style>
body{margin:0;background:#f6f8fb;color:#172033;font-family:Segoe UI,Arial,sans-serif;line-height:1.45}.wrap{max-width:980px;margin:0 auto;padding:28px}.hero{background:#fff;border:1px solid #dbe3ef;border-radius:8px;padding:24px}.badge{display:inline-block;background:#fee2e2;color:#991b1b;border:1px solid #fecaca;border-radius:999px;padding:5px 10px;font-size:12px;font-weight:700;text-transform:uppercase}h1{font-size:24px;margin:14px 0 8px}p{margin:8px 0;color:#42526a}.meta{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:10px;margin:18px 0}.meta div{background:#f8fafc;border:1px solid #e5eaf2;border-radius:6px;padding:10px}table{border-collapse:collapse;width:100%;background:#fff;margin-top:18px;border:1px solid #dbe3ef}th,td{text-align:left;padding:10px;border-bottom:1px solid #edf1f7;vertical-align:top}th{background:#f1f5f9;color:#334155;font-size:12px;text-transform:uppercase}ul{background:#fff;border:1px solid #dbe3ef;border-radius:8px;padding:16px 16px 16px 34px}.foot{font-size:12px;color:#64748b;margin-top:20px}
</style>
</head>
<body><div class="wrap"><div class="hero"><span class="badge">Run aborted</span><h1>Maester drift detection could not start</h1><p>The managed identity is missing permissions or service connectivity required before running Maester. Re-run Install-DriftMaester.ps1 as Global Admin or manually add the items below and rerun the automation job.</p><div class="meta"><div><strong>Managed identity client id</strong><br>$(ConvertTo-HtmlEncodedText $clientId)</div><div><strong>Tenant id</strong><br>$(ConvertTo-HtmlEncodedText $tenant)</div></div></div><h2>Missing items</h2><ul>$permissionList</ul><table><thead><tr><th>Service</th><th>Permission or setting</th><th>Type</th><th>Details</th></tr></thead><tbody>$($rows -join [Environment]::NewLine)</tbody></table><p class="foot">Generated by Invoke-MaesterDriftDetection.ps1 on $(ConvertTo-HtmlEncodedText (Get-Date -Format 'yyyy-MM-dd HH:mm:ss K')).</p></div></body></html>
"@
}

function New-OptionalConnectionWarningEmailHtml {
    param(
        [Parameter(Mandatory = $false)]
        [object[]] $OptionalWarnings
    )

    if (-not $OptionalWarnings -or $OptionalWarnings.Count -eq 0) {
        return ''
    }

    $warningRows = foreach ($warning in $OptionalWarnings) {
        '<tr><td style="padding:8px;border-bottom:1px solid #fde68a;vertical-align:top;"><strong>{0}</strong><br><span style="color:#92400e;font-size:12px;">{1}</span></td><td style="padding:8px;border-bottom:1px solid #fde68a;vertical-align:top;">{2}</td></tr>' -f `
            (ConvertTo-HtmlEncodedText $warning.Service),
            (ConvertTo-HtmlEncodedText $warning.Type),
            (ConvertTo-HtmlEncodedText $warning.Reason)
    }

    return @"
<tr><td style="padding:0 24px 12px 24px;"><table role="presentation" width="100%" cellspacing="0" cellpadding="0" style="border-collapse:collapse;background:#fffbeb;border:1px solid #fcd34d;border-radius:6px;"><tr><td style="padding:12px 12px 4px 12px;"><div style="font-size:12px;font-weight:700;color:#92400e;text-transform:uppercase;">Optional service warning</div><p style="margin:6px 0 8px 0;color:#78350f;">Some optional workload connections were not available. The run continued, but tests for those services may be skipped or have less detail.</p></td></tr><tr><td style="padding:0 12px 12px 12px;"><table width="100%" cellspacing="0" cellpadding="0" style="border-collapse:collapse;border:1px solid #fde68a;background:#fffaf0;"><thead><tr style="background:#fef3c7;"><th align="left" style="padding:8px;color:#78350f;font-size:12px;">Service</th><th align="left" style="padding:8px;color:#78350f;font-size:12px;">Details</th></tr></thead><tbody>$($warningRows -join [Environment]::NewLine)</tbody></table></td></tr></table></td></tr>
"@
}

function Get-StorageContextForResults {
    $automationAccount = Find-CurrentAutomationAccount
    $storagePath = "/subscriptions/$($automationAccount.SubscriptionId)/resourceGroups/$($automationAccount.ResourceGroupName)/providers/Microsoft.Storage/storageAccounts?api-version=2023-01-01"
    $storageAccounts = @((Invoke-AzRestJson -Path $storagePath).value)
    if ($storageAccounts.Count -eq 0) {
        throw "No storage account was found in the Automation Account resource group '$($automationAccount.ResourceGroupName)'. Create one there and grant the managed identity Storage Blob Data Contributor."
    }

    $selectedStorage = @($storageAccounts | Where-Object { $_.name -like '*maester*' } | Sort-Object name | Select-Object -First 1)
    if (-not $selectedStorage) {
        $selectedStorage = $storageAccounts | Sort-Object name | Select-Object -First 1
    }

    if ($storageAccounts.Count -gt 1) {
        Write-RunLog "Multiple storage accounts found in '$($automationAccount.ResourceGroupName)'. Using '$($selectedStorage.name)'." -Level Warning
    }

    $script:DetectedStorageAccountName = $selectedStorage.name
    $script:DetectedStorageResourceGroupName = $automationAccount.ResourceGroupName
    $script:DetectedStorageSubscriptionId = $automationAccount.SubscriptionId

    Write-RunLog "Using storage account '$($script:DetectedStorageAccountName)' in resource group '$($script:DetectedStorageResourceGroupName)' for Maester results."

    $context = New-AzStorageContext -StorageAccountName $script:DetectedStorageAccountName -UseConnectedAccount
    $container = Get-AzStorageContainer -Context $context -Name $ResultsContainerName -ErrorAction SilentlyContinue
    if (-not $container) {
        Write-RunLog "Creating blob container '$ResultsContainerName'."
        New-AzStorageContainer -Context $context -Name $ResultsContainerName -Permission Off | Out-Null
    }

    return $context
}

function Invoke-AzRestJson {
    param([Parameter(Mandatory = $true)][string] $Path)

    $response = Invoke-AzRestMethod -Method GET -Path $Path
    if ([string]::IsNullOrWhiteSpace($response.Content)) {
        return $null
    }

    $response.Content | ConvertFrom-Json
}

function Get-ResourceGroupNameFromResourceId {
    param([Parameter(Mandatory = $true)][string] $ResourceId)

    if ($ResourceId -match '/resourceGroups/([^/]+)/') {
        return $Matches[1]
    }

    throw "Could not parse resource group from resource id '$ResourceId'."
}

function Get-CurrentManagedIdentityPrincipalId {
    $context = Get-MgContext
    $clientId = if ($script:ConnectedManagedIdentityClientId) { $script:ConnectedManagedIdentityClientId } elseif ($context) { $context.ClientId } else { $null }
    if ([string]::IsNullOrWhiteSpace($clientId)) {
        throw 'Cannot resolve the managed identity service principal because Graph context does not include a client id.'
    }

    $graphRoot = (Get-MaesterCloudResourceUrls).Graph.TrimEnd('/')
    $filter = [Uri]::EscapeDataString("appId eq '$clientId'")
    $servicePrincipals = Invoke-MgGraphRequest -Method GET -Uri "$graphRoot/v1.0/servicePrincipals?`$filter=$filter&`$select=id,appId"
    $servicePrincipal = @($servicePrincipals.value | Select-Object -First 1)
    if (-not $servicePrincipal) {
        throw "Could not find a service principal for managed identity client id '$clientId'."
    }

    return $servicePrincipal.id
}

function Find-CurrentAutomationAccount {
    $context = Get-MgContext
    $clientId = if ($script:ConnectedManagedIdentityClientId) { $script:ConnectedManagedIdentityClientId } else { $context.ClientId }
    $principalId = Get-CurrentManagedIdentityPrincipalId
    $subscriptions = @((Invoke-AzRestJson -Path '/subscriptions?api-version=2020-01-01').value)

    foreach ($subscription in $subscriptions) {
        $subscriptionId = $subscription.subscriptionId
        $automationAccounts = @((Invoke-AzRestJson -Path "/subscriptions/$subscriptionId/providers/Microsoft.Automation/automationAccounts?api-version=2023-11-01").value)
        foreach ($account in $automationAccounts) {
            $identity = $account.identity
            if (-not $identity) { continue }

            $matchesIdentity = $false
            if ($identity.principalId -eq $principalId) {
                $matchesIdentity = $true
            }

            if (-not $matchesIdentity -and $identity.userAssignedIdentities) {
                foreach ($userAssignedIdentity in $identity.userAssignedIdentities.PSObject.Properties) {
                    if ($userAssignedIdentity.Value.clientId -eq $clientId -or $userAssignedIdentity.Value.principalId -eq $principalId) {
                        $matchesIdentity = $true
                        break
                    }
                }
            }

            if ($matchesIdentity) {
                return [PSCustomObject]@{
                    Name              = $account.name
                    ResourceGroupName = Get-ResourceGroupNameFromResourceId -ResourceId $account.id
                    SubscriptionId    = $subscriptionId
                    ResourceId        = $account.id
                }
            }
        }
    }

    throw 'Could not find an Azure Automation Account in visible subscriptions with the current managed identity assigned. Grant the identity Reader on the Automation Account/resource group so it can detect its storage account.'
}

function Save-BlobFile {
    param(
        [Parameter(Mandatory = $true)]
        [object] $StorageContext,

        [Parameter(Mandatory = $true)]
        [string] $FilePath,

        [Parameter(Mandatory = $true)]
        [string] $BlobName
    )

    Set-AzStorageBlobContent -Context $StorageContext -Container $ResultsContainerName -File $FilePath -Blob $BlobName -Force | Out-Null
}

function Get-HistoricalResultBlobs {
    param(
        [Parameter(Mandatory = $true)]
        [object] $StorageContext,

        [Parameter(Mandatory = $true)]
        [string] $TenantResultPrefix
    )

    @(Get-AzStorageBlob -Context $StorageContext -Container $ResultsContainerName -Prefix $TenantResultPrefix -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -like '*.json' -and $_.Name -match '/TestResults-[0-9]{8}-[0-9]{6}\.json$' } |
        Sort-Object @{ Expression = { $_.LastModified.UtcDateTime }; Descending = $true }, Name)
}

function Read-ResultBlob {
    param(
        [Parameter(Mandatory = $true)]
        [object] $StorageContext,

        [Parameter(Mandatory = $true)]
        [string] $BlobName,

        [Parameter(Mandatory = $true)]
        [string] $DestinationFolder
    )

    $fileName = Split-Path -Path $BlobName -Leaf
    $destination = Join-Path -Path $DestinationFolder -ChildPath $fileName
    Get-AzStorageBlobContent -Context $StorageContext -Container $ResultsContainerName -Blob $BlobName -Destination $destination -Force | Out-Null
    Get-Content -Path $destination -Raw | ConvertFrom-Json
}

function Get-TestIdentityKey {
    param([Parameter(Mandatory = $true)][object] $Test)

    if ($Test.PSObject.Properties.Name -contains 'Id' -and -not [string]::IsNullOrWhiteSpace([string] $Test.Id)) {
        return [string] $Test.Id
    }

    return [string] $Test.Name
}

function Get-ResultWeight {
    param([AllowNull()][string] $Result)

    switch ($Result) {
        'Passed' { return 50 }
        'Skipped' { return 30 }
        'NotRun' { return 30 }
        'Investigate' { return 20 }
        'Failed' { return 10 }
        'Error' { return 0 }
        default { return 0 }
    }
}

function Compare-MaesterRunResults {
    param(
        [Parameter(Mandatory = $true)]
        [object] $Current,

        [Parameter(Mandatory = $false)]
        [object] $Previous
    )

    if (-not $Previous) {
        return [PSCustomObject]@{
            HasPrevious = $false
            Items       = @()
            Summary     = [PSCustomObject]@{
                Regressed = 0
                Improved  = 0
                Changed   = 0
                Added     = 0
                Removed   = 0
            }
        }
    }

    $currentByKey = @{}
    foreach ($test in @($Current.Tests)) {
        $currentByKey[(Get-TestIdentityKey -Test $test)] = $test
    }

    $previousByKey = @{}
    foreach ($test in @($Previous.Tests)) {
        $previousByKey[(Get-TestIdentityKey -Test $test)] = $test
    }

    $allKeys = @($currentByKey.Keys + $previousByKey.Keys) | Sort-Object -Unique
    $items = foreach ($key in $allKeys) {
        $currentTest = $currentByKey[$key]
        $previousTest = $previousByKey[$key]

        if ($null -eq $previousTest -and $null -ne $currentTest) {
            $classification = if (@('Failed', 'Error', 'Investigate') -contains $currentTest.Result) { 'New finding' } else { 'Added test' }
        } elseif ($null -ne $previousTest -and $null -eq $currentTest) {
            $classification = 'Removed test'
        } elseif ($previousTest.Result -ne $currentTest.Result) {
            $previousWeight = Get-ResultWeight -Result $previousTest.Result
            $currentWeight = Get-ResultWeight -Result $currentTest.Result
            if ($currentWeight -lt $previousWeight) {
                $classification = 'Regressed'
            } elseif ($currentWeight -gt $previousWeight) {
                $classification = 'Improved'
            } else {
                $classification = 'Changed'
            }
        } else {
            $classification = 'Unchanged'
        }

        if ($classification -ne 'Unchanged') {
            [PSCustomObject]@{
                Key            = $key
                Id             = if ($currentTest) { $currentTest.Id } else { $previousTest.Id }
                Title          = if ($currentTest) { $currentTest.Title } else { $previousTest.Title }
                Service        = if ($currentTest -and $currentTest.ResultDetail) { $currentTest.ResultDetail.Service } elseif ($previousTest -and $previousTest.ResultDetail) { $previousTest.ResultDetail.Service } else { '' }
                Severity       = if ($currentTest) { $currentTest.Severity } else { $previousTest.Severity }
                PreviousResult = if ($previousTest) { $previousTest.Result } else { '' }
                CurrentResult  = if ($currentTest) { $currentTest.Result } else { '' }
                Classification = $classification
            }
        }
    }

    $items = @($items)
    return [PSCustomObject]@{
        HasPrevious = $true
        Items       = $items
        Summary     = [PSCustomObject]@{
            Regressed = @($items | Where-Object { $_.Classification -eq 'Regressed' }).Count
            Improved  = @($items | Where-Object { $_.Classification -eq 'Improved' }).Count
            Changed   = @($items | Where-Object { $_.Classification -eq 'Changed' }).Count
            Added     = @($items | Where-Object { $_.Classification -in @('Added test', 'New finding') }).Count
            Removed   = @($items | Where-Object { $_.Classification -eq 'Removed test' }).Count
        }
    }
}

function Get-ScoreFromResult {
    param([Parameter(Mandatory = $true)][object] $Result)

    if ([int] $Result.TotalCount -le 0) { return 0 }
    $healthy = [int] $Result.PassedCount
    $needsReview = [int] $Result.InvestigateCount
    $score = (($healthy + ($needsReview * 0.5)) / [double] $Result.TotalCount) * 100
    return [Math]::Round($score, 1)
}

function New-TrendPoint {
    param([Parameter(Mandatory = $true)][object] $Result)

    [PSCustomObject]@{
        ExecutedAt        = [datetime] $Result.ExecutedAt
        Label             = ([datetime] $Result.ExecutedAt).ToString('dd MMM HH:mm')
        Score             = Get-ScoreFromResult -Result $Result
        PassedCount       = [int] $Result.PassedCount
        FailedCount       = [int] $Result.FailedCount
        ErrorCount        = [int] $Result.ErrorCount
        InvestigateCount  = [int] $Result.InvestigateCount
        SkippedCount      = [int] $Result.SkippedCount
        NotRunCount       = [int] $Result.NotRunCount
        TotalCount        = [int] $Result.TotalCount
    }
}

function New-TrendSvg {
    param([Parameter(Mandatory = $true)][object[]] $Trend)

    if ($Trend.Count -lt 2) {
        return '<div class="empty">Not enough history for a trend yet.</div>'
    }

    $width = 760
    $height = 210
    $paddingLeft = 46
    $paddingRight = 18
    $paddingTop = 20
    $paddingBottom = 38
    $plotWidth = $width - $paddingLeft - $paddingRight
    $plotHeight = $height - $paddingTop - $paddingBottom
    $maxIndex = [Math]::Max(1, $Trend.Count - 1)

    $points = for ($index = 0; $index -lt $Trend.Count; $index++) {
        $x = $paddingLeft + (($plotWidth / $maxIndex) * $index)
        $y = $paddingTop + ($plotHeight - (($Trend[$index].Score / 100) * $plotHeight))
        '{0},{1}' -f ([Math]::Round($x, 1)), ([Math]::Round($y, 1))
    }

    $circles = for ($index = 0; $index -lt $Trend.Count; $index++) {
        $x = $paddingLeft + (($plotWidth / $maxIndex) * $index)
        $y = $paddingTop + ($plotHeight - (($Trend[$index].Score / 100) * $plotHeight))
        '<circle cx="{0}" cy="{1}" r="4"><title>{2}: {3}%</title></circle>' -f ([Math]::Round($x, 1)), ([Math]::Round($y, 1)), (ConvertTo-HtmlEncodedText $Trend[$index].Label), $Trend[$index].Score
    }

    $labels = for ($index = 0; $index -lt $Trend.Count; $index++) {
        if ($index -eq 0 -or $index -eq ($Trend.Count - 1) -or $index % 2 -eq 0) {
            $x = $paddingLeft + (($plotWidth / $maxIndex) * $index)
            '<text x="{0}" y="198" text-anchor="middle">{1}</text>' -f ([Math]::Round($x, 1)), (ConvertTo-HtmlEncodedText $Trend[$index].Label)
        }
    }

    return @"
<svg class="trend" viewBox="0 0 $width $height" role="img" aria-label="Score trend over previous Maester runs">
<line x1="$paddingLeft" y1="$paddingTop" x2="$paddingLeft" y2="$($height - $paddingBottom)" />
<line x1="$paddingLeft" y1="$($height - $paddingBottom)" x2="$($width - $paddingRight)" y2="$($height - $paddingBottom)" />
<text x="8" y="28">100%</text><text x="15" y="$($height - $paddingBottom)">0%</text>
<polyline points="$($points -join ' ')" />
$($circles -join [Environment]::NewLine)
$($labels -join [Environment]::NewLine)
</svg>
"@
}

function Get-MaesterTestService {
    param([AllowNull()][object] $Test)

    if ($Test -and $Test.ResultDetail -and -not [string]::IsNullOrWhiteSpace([string] $Test.ResultDetail.Service)) {
        return [string] $Test.ResultDetail.Service
    }

    if ($Test -and -not [string]::IsNullOrWhiteSpace([string] $Test.Block)) {
        return [string] $Test.Block
    }

    return ''
}

function Get-MaesterTestHelpUrl {
    param([AllowNull()][object] $Test)

    if ($Test -and -not [string]::IsNullOrWhiteSpace([string] $Test.HelpUrl) -and [string] $Test.HelpUrl -match '^https?://') {
        return [string] $Test.HelpUrl
    }

    return ''
}

function Get-MaesterTestErrorText {
    param([AllowNull()][object] $Test)

    if (-not $Test -or -not $Test.ErrorRecord) {
        return ''
    }

    $errorRecords = @($Test.ErrorRecord | Where-Object { $null -ne $_ })
    if ($errorRecords.Count -eq 0) {
        return ''
    }

    $sections = foreach ($errorRecord in $errorRecords) {
        $lines = [System.Collections.Generic.List[string]]::new()

        if ($errorRecord.Exception -and -not [string]::IsNullOrWhiteSpace([string] $errorRecord.Exception.Message)) {
            $lines.Add("Exception: $($errorRecord.Exception.Message)")
        } elseif (-not [string]::IsNullOrWhiteSpace([string] $errorRecord)) {
            $lines.Add("Error: $errorRecord")
        }

        if ($errorRecord.CategoryInfo) {
            if (-not [string]::IsNullOrWhiteSpace([string] $errorRecord.CategoryInfo.Reason)) {
                $lines.Add("Reason: $($errorRecord.CategoryInfo.Reason)")
            }
            if (-not [string]::IsNullOrWhiteSpace([string] $errorRecord.CategoryInfo.Category)) {
                $lines.Add("Category: $($errorRecord.CategoryInfo.Category)")
            }
            if (-not [string]::IsNullOrWhiteSpace([string] $errorRecord.CategoryInfo.TargetName)) {
                $lines.Add("Target: $($errorRecord.CategoryInfo.TargetName)")
            }
        }

        if (-not [string]::IsNullOrWhiteSpace([string] $errorRecord.FullyQualifiedErrorId)) {
            $lines.Add("Fully qualified error id: $($errorRecord.FullyQualifiedErrorId)")
        }

        if ($errorRecord.InvocationInfo -and -not [string]::IsNullOrWhiteSpace([string] $errorRecord.InvocationInfo.PositionMessage)) {
            $lines.Add("Position:$([Environment]::NewLine)$($errorRecord.InvocationInfo.PositionMessage)")
        }

        if (-not [string]::IsNullOrWhiteSpace([string] $errorRecord.ScriptStackTrace)) {
            $lines.Add("Stack trace:$([Environment]::NewLine)$($errorRecord.ScriptStackTrace)")
        }

        if ($lines.Count -gt 0) {
            $lines -join [Environment]::NewLine
        }
    }

    return @($sections | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }) -join "$([Environment]::NewLine)$([Environment]::NewLine)---$([Environment]::NewLine)$([Environment]::NewLine)"
}

function ConvertTo-MaesterDetailValueText {
    param([AllowNull()][object] $Value)

    if ($null -eq $Value) {
        return ''
    }

    if ($Value -is [string]) {
        return [string] $Value
    }

    if ($Value -is [System.Collections.IEnumerable] -and $Value -isnot [string]) {
        $items = @($Value | Where-Object { $null -ne $_ })
        if ($items.Count -eq 0) {
            return ''
        }

        if ($items | Where-Object { $_ -isnot [string] -and $_.PSObject.Properties.Count -gt 0 } | Select-Object -First 1) {
            try {
                return ($items | ConvertTo-Json -Depth 8 -Compress:$false)
            } catch {
                return ($items | Out-String).Trim()
            }
        }

        return ($items | ForEach-Object { [string] $_ }) -join ', '
    }

    if ($Value.PSObject.Properties.Count -gt 0 -and $Value.GetType().Namespace -ne 'System') {
        try {
            return ($Value | ConvertTo-Json -Depth 8 -Compress:$false)
        } catch {
            return ($Value | Out-String).Trim()
        }
    }

    return [string] $Value
}

function New-MaesterDetailBlockHtml {
    param(
        [Parameter(Mandatory = $true)][string] $Label,
        [AllowNull()][object] $Value
    )

    $text = ConvertTo-MaesterDetailValueText -Value $Value
    if ([string]::IsNullOrWhiteSpace($text)) {
        return ''
    }

    return '<div class="detail-block"><div class="detail-label">{0}</div><pre>{1}</pre></div>' -f `
        (ConvertTo-HtmlEncodedText $Label),
        (ConvertTo-HtmlEncodedText $text.Trim())
}

function Get-MaesterSkippedReasonFallback {
    param([AllowNull()][object] $Test)

    if (-not $Test -or [string] $Test.Result -ne 'Skipped') {
        return ''
    }

    $detail = $Test.ResultDetail
    if ($detail -and (-not [string]::IsNullOrWhiteSpace([string] $detail.SkippedReason) -or -not [string]::IsNullOrWhiteSpace([string] $detail.TestSkipped))) {
        return ''
    }

    $tags = (@($Test.Tag) -join ', ')
    $sourceFile = [string] $Test.ScriptBlockFile
    $block = [string] $Test.Block
    $reasonSource = "Tags: $tags`nPester block: $block`nSource file: $sourceFile"

    if ($tags -match '(^|,\s*)XSPM(,|$)' -or $sourceFile -match 'Test-Xspm') {
        return "Skipped by a Pester Describe -Skip condition before the individual Maester test could add a skipped reason. The source test skips Exposure Management/XSPM checks when Get-MtLicenseInformation -Product DefenderXDR does not return DefenderXDR. This usually means Microsoft Defender XDR licensing is not present or Maester could not detect it in this automation context.`n`n$reasonSource"
    }

    if ($tags -match 'AIAgent|CopilotStudio' -or $sourceFile -match 'Test-AIAgentSecurity') {
        return "Skipped by a Pester Describe -Skip condition before the individual Maester test could add a skipped reason. The source test skips Copilot Studio agent security checks when Maester is not connected to Dataverse. Configure DataverseEnvironmentUrl in maester-config.json if auto-discovery cannot find the environment, and make sure the managed identity can acquire a Dataverse token and read the selected environment.`n`n$reasonSource"
    }

    if ($sourceFile -match 'ConditionalAccessWhatIf') {
        return "Skipped by a Pester -Skip condition in the Conditional Access What If tests. These checks are skipped by Maester when the required Entra ID licensing is not detected.`n`n$reasonSource"
    }

    return ''
}

function Get-MaesterTestDetailHtml {
    param([AllowNull()][object] $Test)

    if (-not $Test) {
        return ''
    }

    $blocks = [System.Collections.Generic.List[string]]::new()

    if ($Test.ResultDetail) {
        $detail = $Test.ResultDetail
        $knownProperties = @('TestDescription', 'TestResult', 'SkippedReason', 'TestSkipped', 'TestInvestigate', 'Service', 'Severity', 'TestTitle')

        $blocks.Add((New-MaesterDetailBlockHtml -Label 'Description' -Value $detail.TestDescription))
        $blocks.Add((New-MaesterDetailBlockHtml -Label 'Evidence and checked objects' -Value $detail.TestResult))
        $blocks.Add((New-MaesterDetailBlockHtml -Label 'Skipped reason' -Value $detail.SkippedReason))

        $state = [System.Collections.Generic.List[string]]::new()
        if (-not [string]::IsNullOrWhiteSpace([string] $detail.TestSkipped)) {
            $state.Add("Skipped because: $($detail.TestSkipped)")
        }
        if ($detail.TestInvestigate -eq $true) {
            $state.Add('Requires investigation: true')
        }
        if (-not [string]::IsNullOrWhiteSpace([string] $detail.TestTitle) -and [string] $detail.TestTitle -ne [string] $Test.Name) {
            $state.Add("Maester title: $($detail.TestTitle)")
        }
        $blocks.Add((New-MaesterDetailBlockHtml -Label 'Maester state' -Value $state))

        foreach ($property in @($detail.PSObject.Properties | Where-Object { $_.Name -notin $knownProperties })) {
            $blocks.Add((New-MaesterDetailBlockHtml -Label $property.Name -Value $property.Value))
        }
    }

    $blocks.Add((New-MaesterDetailBlockHtml -Label 'Likely skipped reason' -Value (Get-MaesterSkippedReasonFallback -Test $Test)))

    $blocks.Add((New-MaesterDetailBlockHtml -Label 'Tags' -Value $Test.Tag))
    $blocks.Add((New-MaesterDetailBlockHtml -Label 'Pester block' -Value $Test.Block))
    $blocks.Add((New-MaesterDetailBlockHtml -Label 'Source file' -Value $Test.ScriptBlockFile))

    $content = @($blocks | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }) -join [Environment]::NewLine
    if ([string]::IsNullOrWhiteSpace($content)) {
        return ''
    }

    return '<div class="detail-content">{0}</div>' -f $content
}

function Get-MaesterTestDetailSearchText {
    param([AllowNull()][object] $Test)

    if (-not $Test) {
        return ''
    }

    $parts = [System.Collections.Generic.List[string]]::new()
    if ($Test.ResultDetail) {
        foreach ($property in @($Test.ResultDetail.PSObject.Properties)) {
            $parts.Add((ConvertTo-MaesterDetailValueText -Value $property.Value))
        }
    }
    $parts.Add((Get-MaesterSkippedReasonFallback -Test $Test))
    $parts.Add((ConvertTo-MaesterDetailValueText -Value $Test.Tag))
    $parts.Add((ConvertTo-MaesterDetailValueText -Value $Test.Block))
    $parts.Add((ConvertTo-MaesterDetailValueText -Value $Test.ScriptBlockFile))

    return @($parts | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }) -join ' '
}

function New-MaesterTestTableRows {
    param([Parameter(Mandatory = $true)][object[]] $Tests)

    $rowIndex = 0
    $rows = foreach ($test in @($Tests | Sort-Object Result, Severity, Id)) {
        $rowIndex++
        $result = [string] $test.Result
        $service = Get-MaesterTestService -Test $test
        $helpUrl = Get-MaesterTestHelpUrl -Test $test
        $docCell = if ($helpUrl) {
            '<a href="{0}" target="_blank" rel="noopener">Docs</a>' -f (ConvertTo-HtmlEncodedText $helpUrl)
        } else {
            '<span class="muted">-</span>'
        }
        $errorText = Get-MaesterTestErrorText -Test $test
        $detailHtml = Get-MaesterTestDetailHtml -Test $test
        $detailText = Get-MaesterTestDetailSearchText -Test $test
        $hasDetails = -not [string]::IsNullOrWhiteSpace($detailHtml)
        $hasError = -not [string]::IsNullOrWhiteSpace($errorText)
        $detailId = 'test-detail-{0}' -f $rowIndex
        $reviewCell = if ($hasDetails -or $hasError) {
            $detailPanel = if ($hasDetails) {
                '<div class="modal-panel" data-panel="details">{0}</div>' -f $detailHtml
            } else {
                ''
            }
            $errorPanel = if ($hasError) {
                '<div class="modal-panel" data-panel="error"><pre>{0}</pre></div>' -f (ConvertTo-HtmlEncodedText $errorText)
            } else {
                ''
            }
            '<button type="button" class="detail-open" data-detail-id="{0}">Review</button><div id="{0}" class="modal-source" hidden data-title="{1}" data-id="{2}" data-has-details="{3}" data-has-error="{4}">{5}{6}</div>' -f `
                (ConvertTo-HtmlEncodedText $detailId),
                (ConvertTo-HtmlEncodedText $test.Title),
                (ConvertTo-HtmlEncodedText $test.Id),
                ([string] $hasDetails).ToLowerInvariant(),
                ([string] $hasError).ToLowerInvariant(),
                $detailPanel,
                $errorPanel
        } else {
            '<span class="muted">-</span>'
        }
        $searchText = '{0} {1} {2} {3} {4} {5} {6}' -f $result, $test.Id, $test.Title, $test.Severity, $service, $errorText, $detailText

        '<tr data-result="{0}" data-search="{1}"><td><span class="pill result-{2}">{3}</span></td><td>{4}</td><td>{5}</td><td>{6}</td><td>{7}</td><td>{8}</td><td>{9}</td><td>{10}</td></tr>' -f `
            (ConvertTo-HtmlEncodedText $result),
            (ConvertTo-HtmlEncodedText $searchText.ToLowerInvariant()),
            (ConvertTo-HtmlEncodedText $result),
            (ConvertTo-HtmlEncodedText $result),
            (ConvertTo-HtmlEncodedText $test.Id),
            (ConvertTo-HtmlEncodedText $test.Title),
            (ConvertTo-HtmlEncodedText $test.Severity),
            (ConvertTo-HtmlEncodedText $service),
            (ConvertTo-HtmlEncodedText $test.Duration),
            $docCell,
            $reviewCell
    }

    if ($rows) {
        return @($rows) -join [Environment]::NewLine
    }

    return '<tr><td colspan="8" class="empty">No test results were present in the Maester result.</td></tr>'
}

function New-MaesterDriftReportHtml {
    param(
        [Parameter(Mandatory = $true)]
        [object] $CurrentResult,

        [Parameter(Mandatory = $false)]
        [object] $PreviousResult,

        [Parameter(Mandatory = $true)]
        [object] $Diff,

        [Parameter(Mandatory = $true)]
        [object[]] $Trend,

        [Parameter(Mandatory = $true)]
        [object] $UploadedBlobs,

        [Parameter(Mandatory = $true)]
        [string] $ModuleVersion
    )

    $score = Get-ScoreFromResult -Result $CurrentResult
    $previousScore = if ($PreviousResult) { Get-ScoreFromResult -Result $PreviousResult } else { $null }
    $scoreDelta = if ($null -ne $previousScore) { [Math]::Round($score - $previousScore, 1) } else { $null }
    $scoreDeltaText = if ($null -ne $scoreDelta) { if ($scoreDelta -gt 0) { "+$scoreDelta" } else { [string] $scoreDelta } } else { 'No previous run' }
    $scoreClass = if ($null -eq $scoreDelta) { 'neutral' } elseif ($scoreDelta -lt 0) { 'bad' } elseif ($scoreDelta -gt 0) { 'good' } else { 'neutral' }

    $findings = @($CurrentResult.Tests | Where-Object { $_.Result -in @('Failed', 'Error', 'Investigate') } | Sort-Object Result, Severity, Id)
    $allTestRows = New-MaesterTestTableRows -Tests @($CurrentResult.Tests)

    $diffRows = if ($Diff.HasPrevious -and $Diff.Items.Count -gt 0) {
        foreach ($item in @($Diff.Items | Sort-Object Classification, Id)) {
            '<tr><td><span class="pill drift-{0}">{1}</span></td><td>{2}</td><td>{3}</td><td>{4}</td><td>{5}</td><td>{6}</td></tr>' -f `
                (([string] $item.Classification).ToLowerInvariant().Replace(' ', '-')),
                (ConvertTo-HtmlEncodedText $item.Classification),
                (ConvertTo-HtmlEncodedText $item.Id),
                (ConvertTo-HtmlEncodedText $item.Title),
                (ConvertTo-HtmlEncodedText $item.PreviousResult),
                (ConvertTo-HtmlEncodedText $item.CurrentResult),
                (ConvertTo-HtmlEncodedText $item.Severity)
        }
    } elseif ($Diff.HasPrevious) {
        '<tr><td colspan="7" class="empty">No drift detected compared with the previous run.</td></tr>'
    } else {
        '<tr><td colspan="7" class="empty">No previous result was present in blob storage, so no diff was calculated for this first run.</td></tr>'
    }

    $trendRows = foreach ($point in $Trend) {
        '<tr><td>{0}</td><td>{1}%</td><td>{2}</td><td>{3}</td><td>{4}</td><td>{5}</td><td>{6}</td></tr>' -f `
            (ConvertTo-HtmlEncodedText $point.Label),
            (ConvertTo-HtmlEncodedText $point.Score),
            (ConvertTo-HtmlEncodedText $point.PassedCount),
            (ConvertTo-HtmlEncodedText $point.FailedCount),
            (ConvertTo-HtmlEncodedText $point.ErrorCount),
            (ConvertTo-HtmlEncodedText $point.InvestigateCount),
            (ConvertTo-HtmlEncodedText $point.TotalCount)
    }

    $trendSvg = New-TrendSvg -Trend $Trend
    $previousRunText = if ($PreviousResult) { ([datetime] $PreviousResult.ExecutedAt).ToString('yyyy-MM-dd HH:mm:ss K') } else { 'No previous run found' }
    $blobList = @($UploadedBlobs.PSObject.Properties | ForEach-Object { '<li><strong>{0}</strong><br><span>{1}</span></li>' -f (ConvertTo-HtmlEncodedText $_.Name), (ConvertTo-HtmlEncodedText $_.Value) }) -join [Environment]::NewLine

    return @"
<!doctype html>
<html>
<head>
<meta charset="utf-8">
<style>
body{margin:0;background:#f3f6fb;color:#172033;font-family:Segoe UI,Arial,sans-serif;line-height:1.45}.wrap{max-width:1280px;margin:0 auto;padding:28px}.hero,.card,table,ul.blobs,.trend,.filters{background:#fff;border:1px solid #d9e2ef;border-radius:8px}.hero{padding:24px}.eyebrow{font-size:12px;font-weight:700;color:#2563eb;text-transform:uppercase;letter-spacing:.04em}h1{font-size:26px;margin:8px 0 10px}h2{font-size:18px;margin:28px 0 10px}.grid{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:12px;margin-top:18px}.card{padding:16px}.card .label{font-size:12px;color:#64748b;text-transform:uppercase;font-weight:700}.card .value{font-size:28px;font-weight:700;margin-top:4px}.good{color:#047857}.bad{color:#b91c1c}.neutral{color:#475569}.meta{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:10px;margin-top:18px}.meta div{background:#f8fafc;border:1px solid #e5eaf2;border-radius:6px;padding:10px}.pill{display:inline-block;border-radius:999px;padding:4px 8px;font-size:12px;font-weight:700}.result-Passed{background:#dcfce7;color:#166534}.result-Failed,.result-Error,.drift-regressed,.drift-new-finding{background:#fee2e2;color:#991b1b}.result-Investigate,.result-Skipped,.result-NotRun,.drift-changed,.drift-added-test{background:#fef3c7;color:#92400e}.drift-improved{background:#dcfce7;color:#166534}.drift-removed-test{background:#e0f2fe;color:#075985}table{border-collapse:separate;border-spacing:0;width:100%;overflow:hidden}th,td{text-align:left;padding:10px;border-bottom:1px solid #edf1f7;vertical-align:top}tr:last-child td{border-bottom:0}th{background:#eef4fb;color:#334155;font-size:12px;text-transform:uppercase}.empty{color:#64748b;text-align:center;padding:18px}.muted{color:#94a3b8}.trend{width:100%;height:auto}.trend line{stroke:#cbd5e1;stroke-width:1}.trend polyline{fill:none;stroke:#2563eb;stroke-width:3}.trend circle{fill:#2563eb}.trend text{fill:#64748b;font-size:11px}ul.blobs{padding:14px 14px 14px 30px}.blobs span{color:#64748b;font-size:12px}.foot{font-size:12px;color:#64748b;margin-top:20px}.filters{display:flex;gap:10px;align-items:center;flex-wrap:wrap;padding:12px;margin-bottom:10px}.filters input,.filters select{border:1px solid #cbd5e1;border-radius:6px;padding:8px 10px;font:inherit}.filters input{min-width:280px;flex:1}.filters label{font-size:12px;font-weight:700;color:#475569;text-transform:uppercase}#allTests{table-layout:fixed}#allTests th:nth-child(1),#allTests td:nth-child(1){width:86px}#allTests th:nth-child(2),#allTests td:nth-child(2){width:76px;word-break:break-word;font-size:12px}#allTests th:nth-child(4),#allTests td:nth-child(4){width:82px}#allTests th:nth-child(5),#allTests td:nth-child(5){width:126px}#allTests th:nth-child(6),#allTests td:nth-child(6){width:82px}#allTests th:nth-child(7),#allTests td:nth-child(7){width:56px}#allTests th:nth-child(8),#allTests td:nth-child(8){width:92px}.detail-open{border:1px solid #bfdbfe;background:#eff6ff;color:#1d4ed8;border-radius:6px;padding:6px 10px;font:inherit;font-weight:700;cursor:pointer}.detail-open:hover{background:#dbeafe}.modal-backdrop{position:fixed;inset:0;background:rgba(15,23,42,.62);display:none;align-items:center;justify-content:center;padding:24px;z-index:999}.modal-backdrop.open{display:flex}.modal{background:#fff;border-radius:8px;box-shadow:0 24px 80px rgba(15,23,42,.35);width:min(1120px,96vw);max-height:88vh;display:flex;flex-direction:column;overflow:hidden}.modal-head{display:flex;justify-content:space-between;gap:14px;align-items:flex-start;padding:18px 20px;border-bottom:1px solid #e5eaf2}.modal-title{font-size:18px;font-weight:700}.modal-subtitle{font-size:12px;color:#64748b;margin-top:4px}.modal-close{border:0;background:#f1f5f9;color:#334155;border-radius:6px;padding:7px 10px;cursor:pointer;font:inherit}.modal-tabs{display:flex;gap:8px;padding:12px 20px 0;border-bottom:1px solid #e5eaf2}.modal-tab{border:1px solid #cbd5e1;border-bottom:0;background:#f8fafc;color:#475569;border-radius:6px 6px 0 0;padding:8px 12px;cursor:pointer;font:inherit;font-weight:700}.modal-tab.active{background:#fff;color:#1d4ed8}.modal-body{padding:18px 20px;overflow:auto}.modal-panel{display:none}.modal-panel.active{display:block}.detail-block{margin:0 0 14px}.detail-label{color:#475569;font-size:12px;font-weight:700;text-transform:uppercase}.modal pre{white-space:pre-wrap;word-break:break-word;border-radius:6px;margin:6px 0 0 0;padding:10px;background:#f8fafc;border:1px solid #dbe6f3;color:#172033}.modal-panel[data-panel="error"] pre{background:#fff7ed;border-color:#fed7aa;color:#7c2d12}a{color:#2563eb;text-decoration:none}a:hover{text-decoration:underline}@media(max-width:860px){.grid,.meta{grid-template-columns:1fr 1fr}#allTests{table-layout:auto}.modal{width:98vw;max-height:92vh}}@media(max-width:560px){.grid,.meta{grid-template-columns:1fr}.wrap{padding:16px}.filters input{min-width:0;width:100%}.modal-backdrop{padding:10px}.modal-head{padding:14px}.modal-tabs{padding-left:14px}.modal-body{padding:14px}}
</style>
</head>
<body><div class="wrap"><div class="hero"><div class="eyebrow">Maester $(ConvertTo-HtmlEncodedText $ModuleVersion) report</div><h1>$(ConvertTo-HtmlEncodedText $CurrentResult.TenantName)</h1><p>Executed at $(ConvertTo-HtmlEncodedText ([datetime] $CurrentResult.ExecutedAt).ToString('yyyy-MM-dd HH:mm:ss K')). Previous run: $(ConvertTo-HtmlEncodedText $previousRunText).</p><div class="meta"><div><strong>Tenant id</strong><br>$(ConvertTo-HtmlEncodedText $CurrentResult.TenantId)</div><div><strong>Maester version</strong><br>$(ConvertTo-HtmlEncodedText $ModuleVersion)</div><div><strong>Storage account</strong><br>$(ConvertTo-HtmlEncodedText $script:DetectedStorageAccountName)</div></div></div><div class="grid"><div class="card"><div class="label">Score</div><div class="value">$score%</div></div><div class="card"><div class="label">Score delta</div><div class="value $scoreClass">$scoreDeltaText</div></div><div class="card"><div class="label">Findings</div><div class="value bad">$($findings.Count)</div></div><div class="card"><div class="label">Total tests</div><div class="value">$($CurrentResult.TotalCount)</div></div><div class="card"><div class="label">Passed</div><div class="value good">$($CurrentResult.PassedCount)</div></div><div class="card"><div class="label">Failed</div><div class="value bad">$($CurrentResult.FailedCount)</div></div><div class="card"><div class="label">Errors</div><div class="value bad">$($CurrentResult.ErrorCount)</div></div><div class="card"><div class="label">Investigate</div><div class="value neutral">$($CurrentResult.InvestigateCount)</div></div></div><h2>Drift since previous run</h2><table><thead><tr><th>Status</th><th>Id</th><th>Title</th><th>Previous</th><th>Current</th><th>Severity</th></tr></thead><tbody>$($diffRows -join [Environment]::NewLine)</tbody></table><h2>Score trend</h2>$trendSvg<table><thead><tr><th>Run</th><th>Score</th><th>Passed</th><th>Failed</th><th>Errors</th><th>Investigate</th><th>Total</th></tr></thead><tbody>$($trendRows -join [Environment]::NewLine)</tbody></table><h2>All Maester tests</h2><div class="filters"><label for="resultFilter">Result</label><select id="resultFilter"><option value="">All</option><option>Passed</option><option>Failed</option><option>Error</option><option>Investigate</option><option>Skipped</option><option>NotRun</option></select><label for="testSearch">Search</label><input id="testSearch" type="search" placeholder="Search id, title, severity, service, evidence, or error"></div><table id="allTests"><thead><tr><th>Result</th><th>Id</th><th>Title</th><th>Severity</th><th>Service</th><th>Duration</th><th>Fix</th><th>Review</th></tr></thead><tbody>$allTestRows</tbody></table><h2>Stored artifacts</h2><ul class="blobs">$blobList</ul><p class="foot">Generated by Invoke-MaesterDriftDetection.ps1. Native Maester JSON, HTML and Markdown are stored unchanged in the '$ResultsContainerName' blob container.</p></div><div id="detailModal" class="modal-backdrop" role="dialog" aria-modal="true" aria-hidden="true"><div class="modal"><div class="modal-head"><div><div id="modalTitle" class="modal-title">Test details</div><div id="modalSubtitle" class="modal-subtitle"></div></div><button type="button" class="modal-close">Close</button></div><div id="modalTabs" class="modal-tabs"></div><div id="modalBody" class="modal-body"></div></div></div><script>(function(){var result=document.getElementById('resultFilter');var search=document.getElementById('testSearch');var rows=[].slice.call(document.querySelectorAll('#allTests tbody tr[data-result]'));function apply(){var selected=(result.value||'').toLowerCase();var query=(search.value||'').toLowerCase();rows.forEach(function(row){var okResult=!selected||row.getAttribute('data-result').toLowerCase()===selected;var okSearch=!query||(row.getAttribute('data-search')||'').indexOf(query)>=0;row.style.display=okResult&&okSearch?'':'none';});}if(result){result.addEventListener('change',apply);}if(search){search.addEventListener('input',apply);}var modal=document.getElementById('detailModal');var modalTitle=document.getElementById('modalTitle');var modalSubtitle=document.getElementById('modalSubtitle');var modalTabs=document.getElementById('modalTabs');var modalBody=document.getElementById('modalBody');function selectTab(name){[].slice.call(modalTabs.querySelectorAll('.modal-tab')).forEach(function(tab){tab.classList.toggle('active',tab.getAttribute('data-tab')===name);});[].slice.call(modalBody.querySelectorAll('.modal-panel')).forEach(function(panel){panel.classList.toggle('active',panel.getAttribute('data-panel')===name);});modalBody.scrollTop=0;}function closeModal(){modal.classList.remove('open');modal.setAttribute('aria-hidden','true');modalBody.innerHTML='';modalTabs.innerHTML='';}function openModal(source){modalTitle.textContent=source.getAttribute('data-title')||'Test details';modalSubtitle.textContent=source.getAttribute('data-id')?'Id: '+source.getAttribute('data-id'):'';modalBody.innerHTML=source.innerHTML;modalTabs.innerHTML='';var panels=[].slice.call(modalBody.querySelectorAll('.modal-panel'));panels.forEach(function(panel,index){var name=panel.getAttribute('data-panel');var label=name==='error'?'Technical error':'Details';var tab=document.createElement('button');tab.type='button';tab.className='modal-tab';tab.setAttribute('data-tab',name);tab.textContent=label;tab.addEventListener('click',function(){selectTab(name);});modalTabs.appendChild(tab);if(index===0){selectTab(name);}});modal.classList.add('open');modal.setAttribute('aria-hidden','false');}document.addEventListener('click',function(event){var opener=event.target.closest('.detail-open');if(opener){var source=document.getElementById(opener.getAttribute('data-detail-id'));if(source){openModal(source);}return;}if(event.target.classList.contains('modal-close')||event.target===modal){closeModal();}});document.addEventListener('keydown',function(event){if(event.key==='Escape'&&modal.classList.contains('open')){closeModal();}});}());</script></body></html>
"@
}

function New-MaesterDriftEmailHtml {
    param(
        [Parameter(Mandatory = $true)]
        [object] $CurrentResult,

        [Parameter(Mandatory = $false)]
        [object] $PreviousResult,

        [Parameter(Mandatory = $true)]
        [object] $Diff,

        [Parameter(Mandatory = $true)]
        [object[]] $Trend,

        [Parameter(Mandatory = $true)]
        [string] $ModuleVersion,

        [Parameter(Mandatory = $false)]
        [object[]] $OptionalWarnings = @()
    )

    $score = Get-ScoreFromResult -Result $CurrentResult
    $previousScore = if ($PreviousResult) { Get-ScoreFromResult -Result $PreviousResult } else { $null }
    $scoreDelta = if ($null -ne $previousScore) { [Math]::Round($score - $previousScore, 1) } else { $null }
    $scoreDeltaText = if ($null -ne $scoreDelta) { if ($scoreDelta -gt 0) { "+$scoreDelta" } else { [string] $scoreDelta } } else { 'No previous run' }
    $scoreDeltaColor = if ($null -eq $scoreDelta) { '#475569' } elseif ($scoreDelta -lt 0) { '#b91c1c' } elseif ($scoreDelta -gt 0) { '#047857' } else { '#475569' }
    $findingCount = @($CurrentResult.Tests | Where-Object { $_.Result -in @('Failed', 'Error', 'Investigate') }).Count
    $previousRunText = if ($PreviousResult) { ([datetime] $PreviousResult.ExecutedAt).ToString('yyyy-MM-dd HH:mm:ss K') } else { 'No previous run found' }

    $diffRows = if ($Diff.HasPrevious -and $Diff.Items.Count -gt 0) {
        foreach ($item in @($Diff.Items | Sort-Object Classification, Id | Select-Object -First 25)) {
            '<tr><td style="padding:8px;border-bottom:1px solid #e5eaf2;">{0}</td><td style="padding:8px;border-bottom:1px solid #e5eaf2;">{1}</td><td style="padding:8px;border-bottom:1px solid #e5eaf2;">{2}</td><td style="padding:8px;border-bottom:1px solid #e5eaf2;">{3}</td><td style="padding:8px;border-bottom:1px solid #e5eaf2;">{4}</td></tr>' -f `
                (ConvertTo-HtmlEncodedText $item.Classification),
                (ConvertTo-HtmlEncodedText $item.Id),
                (ConvertTo-HtmlEncodedText $item.Title),
                (ConvertTo-HtmlEncodedText $item.PreviousResult),
                (ConvertTo-HtmlEncodedText $item.CurrentResult)
        }
    } elseif ($Diff.HasPrevious) {
        '<tr><td colspan="5" style="padding:14px;color:#64748b;text-align:center;">No drift detected compared with the previous run.</td></tr>'
    } else {
        '<tr><td colspan="5" style="padding:14px;color:#64748b;text-align:center;">No previous result was present in blob storage, so no diff was calculated.</td></tr>'
    }

    $trendRows = foreach ($point in $Trend) {
        '<tr><td style="padding:8px;border-bottom:1px solid #e5eaf2;">{0}</td><td style="padding:8px;border-bottom:1px solid #e5eaf2;">{1}%</td><td style="padding:8px;border-bottom:1px solid #e5eaf2;">{2}</td><td style="padding:8px;border-bottom:1px solid #e5eaf2;">{3}</td><td style="padding:8px;border-bottom:1px solid #e5eaf2;">{4}</td><td style="padding:8px;border-bottom:1px solid #e5eaf2;">{5}</td></tr>' -f `
            (ConvertTo-HtmlEncodedText $point.Label),
            (ConvertTo-HtmlEncodedText $point.Score),
            (ConvertTo-HtmlEncodedText $point.PassedCount),
            (ConvertTo-HtmlEncodedText $point.FailedCount),
            (ConvertTo-HtmlEncodedText $point.ErrorCount),
            (ConvertTo-HtmlEncodedText $point.InvestigateCount)
    }

    $optionalWarningHtml = New-OptionalConnectionWarningEmailHtml -OptionalWarnings $OptionalWarnings

    return @"
<!doctype html>
<html><body style="margin:0;background:#f5f7fb;color:#172033;font-family:Segoe UI,Arial,sans-serif;line-height:1.45;">
<table role="presentation" width="100%" cellspacing="0" cellpadding="0" style="background:#f5f7fb;"><tr><td align="center" style="padding:24px;">
<table role="presentation" width="760" cellspacing="0" cellpadding="0" style="width:760px;max-width:100%;background:#ffffff;border:1px solid #d9e2ef;border-radius:8px;">
<tr><td style="padding:24px 24px 8px 24px;"><div style="font-size:12px;font-weight:700;color:#2563eb;text-transform:uppercase;">Maester drift report</div><h1 style="font-size:24px;margin:8px 0 8px 0;color:#172033;">$(ConvertTo-HtmlEncodedText $CurrentResult.TenantName)</h1><p style="margin:0;color:#475569;">Executed at $(ConvertTo-HtmlEncodedText ([datetime] $CurrentResult.ExecutedAt).ToString('yyyy-MM-dd HH:mm:ss K')). Previous run: $(ConvertTo-HtmlEncodedText $previousRunText).</p></td></tr>
<tr><td style="padding:12px 24px;"><table role="presentation" width="100%" cellspacing="0" cellpadding="0"><tr><td style="padding:12px;background:#f8fafc;border:1px solid #e5eaf2;"><div style="font-size:12px;color:#64748b;font-weight:700;text-transform:uppercase;">Score</div><div style="font-size:28px;font-weight:700;">$score%</div></td><td style="padding:12px;background:#f8fafc;border:1px solid #e5eaf2;"><div style="font-size:12px;color:#64748b;font-weight:700;text-transform:uppercase;">Score delta</div><div style="font-size:28px;font-weight:700;color:$scoreDeltaColor;">$scoreDeltaText</div></td><td style="padding:12px;background:#f8fafc;border:1px solid #e5eaf2;"><div style="font-size:12px;color:#64748b;font-weight:700;text-transform:uppercase;">Findings</div><div style="font-size:28px;font-weight:700;color:#b91c1c;">$findingCount</div></td><td style="padding:12px;background:#f8fafc;border:1px solid #e5eaf2;"><div style="font-size:12px;color:#64748b;font-weight:700;text-transform:uppercase;">Passed</div><div style="font-size:28px;font-weight:700;color:#047857;">$($CurrentResult.PassedCount)</div></td></tr></table></td></tr>
<tr><td style="padding:8px 24px;"><p style="margin:0;color:#475569;">The attached HTML report contains all tests, passed results, documentation links, and browser filtering.</p></td></tr>
<tr><td style="padding:16px 24px 8px 24px;"><h2 style="font-size:17px;margin:0 0 8px 0;">Drift since previous run</h2><table width="100%" cellspacing="0" cellpadding="0" style="border-collapse:collapse;border:1px solid #d9e2ef;"><thead><tr style="background:#eef4fb;"><th align="left" style="padding:8px;">Status</th><th align="left" style="padding:8px;">Id</th><th align="left" style="padding:8px;">Title</th><th align="left" style="padding:8px;">Previous</th><th align="left" style="padding:8px;">Current</th></tr></thead><tbody>$($diffRows -join [Environment]::NewLine)</tbody></table></td></tr>
<tr><td style="padding:16px 24px 24px 24px;"><h2 style="font-size:17px;margin:0 0 8px 0;">Score trend</h2><table width="100%" cellspacing="0" cellpadding="0" style="border-collapse:collapse;border:1px solid #d9e2ef;"><thead><tr style="background:#eef4fb;"><th align="left" style="padding:8px;">Run</th><th align="left" style="padding:8px;">Score</th><th align="left" style="padding:8px;">Passed</th><th align="left" style="padding:8px;">Failed</th><th align="left" style="padding:8px;">Errors</th><th align="left" style="padding:8px;">Investigate</th></tr></thead><tbody>$($trendRows -join [Environment]::NewLine)</tbody></table><p style="font-size:12px;color:#64748b;margin:16px 0 0 0;">Maester version: $(ConvertTo-HtmlEncodedText $ModuleVersion). Storage account: $(ConvertTo-HtmlEncodedText $script:DetectedStorageAccountName).</p></td></tr>
$optionalWarningHtml
</table></td></tr></table>
</body></html>
"@
}

function Initialize-WorkingTests {
    param([Parameter(Mandatory = $true)][string] $WorkingRoot)

    $testsWorkingPath = Join-Path -Path $WorkingRoot -ChildPath 'maester-tests'
    New-Item -Path $testsWorkingPath -ItemType Directory -Force | Out-Null

    Write-RunLog "Updating Maester tests in working folder '$testsWorkingPath'."
    Update-MaesterTests -Path $testsWorkingPath -Force | Out-Null

    return $testsWorkingPath
}

function Invoke-FullMaesterRun {
    param(
        [Parameter(Mandatory = $true)]
        [string] $TestsPath,

        [Parameter(Mandatory = $true)]
        [string] $OutputFolder,

        [Parameter(Mandatory = $true)]
        [string] $OutputFileName
    )

    $invokeParams = @{
        Path                 = $TestsPath
        OutputFolder         = $OutputFolder
        OutputFolderFileName = $OutputFileName
        NonInteractive       = $true
        PassThru             = $true
        Verbosity            = 'None'
        SkipVersionCheck     = $true
        SkipGraphConnect       = $true
        IncludePreview       = $false
        IncludeLongRunning   = $true
    }

    Write-RunLog "Running Maester tests from '$TestsPath'."
    Invoke-Maester @invokeParams
}

Write-Output "Starting Maester drift detection run."
$workingRoot = Join-Path -Path ([System.IO.Path]::GetTempPath()) -ChildPath "MaesterDrift-$([Guid]::NewGuid().ToString('N'))"
New-Item -Path $workingRoot -ItemType Directory -Force | Out-Null
Write-Output "Working root: $workingRoot"

try {
    Write-Output "Detecting Azure Automation Account and Storage Account for Maester results."
    $parsedReportRecipients = @($ReportRecipient -split ',' | ForEach-Object { $_.Trim() } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    if ($parsedReportRecipients.Count -eq 0) {
        throw 'ReportRecipient is required. Add one or more comma-separated email addresses as a runbook parameter.'
    }

    $global:VerbosePreference = 'SilentlyContinue'
    Import-RequiredModule -Name DLLPickle
    Import-DPLibrary
    Import-RequiredModule -Name ExchangeOnlineManagement
    Import-RequiredModule -Name MicrosoftTeams
    Import-RequiredModule -Name Az.Accounts
    Import-RequiredModule -Name Az.Storage
    Import-RequiredModule -Name Microsoft.Graph.Authentication
    Import-RequiredModule -Name Maester
    $global:VerbosePreference = 'Continue'

    $graphContext = Connect-RunbookIdentity
    $requiredGraphPermissions = Get-RequiredGraphPermissions
    $missingGraphPermissions = [System.Collections.Generic.List[object]]::new()
    foreach ($missing in (Test-GraphPermissions -RequiredPermissions $requiredGraphPermissions)) {
        $missingGraphPermissions.Add($missing)
    }

    if ($missingGraphPermissions.Count -ge @($requiredGraphPermissions).Count) {
        $missingHtml = New-MissingPermissionReportHtml -MissingItems $missingGraphPermissions.ToArray() -GraphContext $graphContext
        $subject = "$MailSubjectPrefix aborted: no required Graph permissions available"
        try {
            Send-DriftMail -Subject $subject -HtmlBody $missingHtml
            Write-RunLog "Missing required permission report sent to $($parsedReportRecipients -join ', ')" -Level Warning
        } catch {
            Write-RunLog "Could not send missing permission report. This usually means Mail.Send is also missing or the sender mailbox is not allowed. Error: $($_.Exception.Message)" -Level Warning
        }

        throw "None of the required Microsoft Graph permissions are available for the managed identity."
    }

    Write-Output "Testing connections to optional services used by Maester for richer reporting and drift detection."

    $optionalConnectionWarnings = [System.Collections.Generic.List[object]]::new()
    foreach ($missing in $missingGraphPermissions) {
        $optionalConnectionWarnings.Add($missing)
        Write-RunLog "Graph permission warning: $($missing.Permission) is not available. Related Maester tests may be skipped or have less detail." -Level Warning
    }

    foreach ($missing in (Connect-OptionalMaesterServices)) {
        $optionalConnectionWarnings.Add($missing)
        Write-RunLog "Optional service warning: $($missing.Service) - $($missing.Reason)" -Level Warning
    }

    $testsWorkingPath = Initialize-WorkingTests -WorkingRoot $workingRoot
    $runOutputFolder = Join-Path -Path $workingRoot -ChildPath 'test-results'
    New-Item -Path $runOutputFolder -ItemType Directory -Force | Out-Null

    Write-Output "Invoking Maester run and generating results, this can take a while depending on tenant size."

    $timestamp = Get-Date -Format 'yyyyMMdd-HHmmss'
    $outputFileName = "TestResults-$timestamp"
    $maesterResult = Invoke-FullMaesterRun -TestsPath $testsWorkingPath -OutputFolder $runOutputFolder -OutputFileName $outputFileName

    if (-not $maesterResult) {
        throw 'Invoke-Maester did not return a result object.'
    }else{
        Write-Output "Done! Comparing with previous runs (if any)...."
    }

    $jsonPath = Join-Path -Path $runOutputFolder -ChildPath "$outputFileName.json"
    $htmlPath = Join-Path -Path $runOutputFolder -ChildPath "$outputFileName.html"
    $markdownPath = Join-Path -Path $runOutputFolder -ChildPath "$outputFileName.md"

    foreach ($path in @($jsonPath, $htmlPath, $markdownPath)) {
        if (-not (Test-Path -Path $path -PathType Leaf)) {
            throw "Expected Maester output file was not created: $path"
        }
    }

    $storageContext = Get-StorageContextForResults
    $tenantSafe = ([string] $maesterResult.TenantId) -replace '[^a-zA-Z0-9-]', '_'
    if ([string]::IsNullOrWhiteSpace($tenantSafe)) { $tenantSafe = 'unknown-tenant' }
    $tenantPrefix = (($BlobPrefix.Trim('/'), $tenantSafe) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }) -join '/'
    $resultPrefix = "$tenantPrefix/results"
    $reportPrefix = "$tenantPrefix/reports"

    $existingBlobs = Get-HistoricalResultBlobs -StorageContext $storageContext -TenantResultPrefix $resultPrefix
    $previousBlob = $existingBlobs | Select-Object -First 1
    $previousResult = $null
    if ($previousBlob) {
        Write-RunLog "Previous result found: $($previousBlob.Name)."
        $previousResult = Read-ResultBlob -StorageContext $storageContext -BlobName $previousBlob.Name -DestinationFolder $workingRoot
    } else {
        Write-RunLog "No previous result found for tenant '$tenantSafe'. Diff will be skipped."
    }

    $jsonBlob = "$resultPrefix/$outputFileName.json"
    $htmlBlob = "$resultPrefix/$outputFileName.html"
    $markdownBlob = "$resultPrefix/$outputFileName.md"
    Save-BlobFile -StorageContext $storageContext -FilePath $jsonPath -BlobName $jsonBlob
    Save-BlobFile -StorageContext $storageContext -FilePath $htmlPath -BlobName $htmlBlob
    Save-BlobFile -StorageContext $storageContext -FilePath $markdownPath -BlobName $markdownBlob

    $allBlobs = Get-HistoricalResultBlobs -StorageContext $storageContext -TenantResultPrefix $resultPrefix
    $trendResults = [System.Collections.Generic.List[object]]::new()
    foreach ($blob in @($allBlobs | Select-Object -First $TrendRunCount)) {
        try {
            $trendResult = Read-ResultBlob -StorageContext $storageContext -BlobName $blob.Name -DestinationFolder $workingRoot
            $trendResults.Add((New-TrendPoint -Result $trendResult))
        } catch {
            Write-RunLog "Could not read trend result '$($blob.Name)': $($_.Exception.Message)" -Level Warning
        }
    }
    $trend = @($trendResults.ToArray() | Sort-Object ExecutedAt)

    $diff = Compare-MaesterRunResults -Current $maesterResult -Previous $previousResult
    $uploadedBlobs = [PSCustomObject]@{
        Json           = $jsonBlob
        Html           = $htmlBlob
        Markdown       = $markdownBlob
        PreviousResult = if ($previousBlob) { $previousBlob.Name } else { 'None' }
    }

    $driftBlob = "$reportPrefix/DriftReport-$timestamp.html"
    $uploadedBlobs | Add-Member -MemberType NoteProperty -Name DriftReport -Value $driftBlob

    $moduleVersion = Get-InstalledMaesterModuleVersion
    $driftHtml = New-MaesterDriftReportHtml -CurrentResult $maesterResult -PreviousResult $previousResult -Diff $diff -Trend $trend -UploadedBlobs $uploadedBlobs -ModuleVersion $moduleVersion
    $emailHtml = New-MaesterDriftEmailHtml -CurrentResult $maesterResult -PreviousResult $previousResult -Diff $diff -Trend $trend -ModuleVersion $moduleVersion -OptionalWarnings $optionalConnectionWarnings.ToArray()
    $driftReportPath = Join-Path -Path $runOutputFolder -ChildPath "DriftReport-$timestamp.html"
    $driftHtml | Out-File -FilePath $driftReportPath -Encoding UTF8

    Save-BlobFile -StorageContext $storageContext -FilePath $driftReportPath -BlobName $driftBlob

    $subjectScore = Get-ScoreFromResult -Result $maesterResult
    $subject = "${MailSubjectPrefix}: $($maesterResult.TenantName) score $subjectScore%, findings $($maesterResult.FailedCount + $maesterResult.ErrorCount + $maesterResult.InvestigateCount)"
    
    # Determine whether to send the report based on AlwaysSendReport flag, first run detection, and drift presence
    $isFirstRun = -not $previousResult
    $hasDiff = $diff.Summary.Regressed -gt 0 -or $diff.Summary.Improved -gt 0 -or $diff.Summary.Changed -gt 0 -or $diff.Summary.Added -gt 0 -or $diff.Summary.Removed -gt 0
    $shouldSendReport = $AlwaysSendReport -or $isFirstRun -or $hasDiff
    
    if ($shouldSendReport) {
        Send-DriftMail -Subject $subject -HtmlBody $emailHtml -AttachmentPath $driftReportPath
        Write-RunLog "Maester drift detection completed. Report sent to $($parsedReportRecipients -join ', ')" -Level Success
    } else {
        Write-RunLog "Maester drift detection completed. No changes detected and AlwaysSendReport is false, so no report was sent." -Level Info
    }
} catch {
    Write-RunLog "Unhandled runbook exception: $($_.Exception.Message)" -Level Error
    Write-RunLog "Exception type: $($_.Exception.GetType().FullName)" -Level Error

    if ($_.ScriptStackTrace) {
        Write-RunLog "Script stack trace:$([Environment]::NewLine)$($_.ScriptStackTrace)" -Level Error
    }

    if ($_.Exception.InnerException) {
        Write-RunLog "Inner exception: $($_.Exception.InnerException.Message)" -Level Error
    }

    if ($_.InvocationInfo) {
        Write-RunLog "Failed command: $($_.InvocationInfo.Line)" -Level Error
        Write-RunLog "Position: $($_.InvocationInfo.PositionMessage)" -Level Error
    }

    throw
} finally {
    if ((Test-Path -Path $workingRoot)) {
        Remove-Item -Path $workingRoot -Recurse -Force -ErrorAction SilentlyContinue
    }
}