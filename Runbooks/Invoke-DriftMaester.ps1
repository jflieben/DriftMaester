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

.PARAMETER IncludeMaesterReport
Optional. Boolean flag (default: $false). When $true, the most recent original Maester HTML report is attached to the
report email alongside the DriftMaester drift report. When $false (default), only the drift report is attached.

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
    [string] $MailSubjectPrefix = 'DriftMaester report',

    [Parameter(Mandatory = $false)]
    [ValidateSet('auto', 'attach', 'link')]
    [string] $ReportDelivery = 'auto',

    [Parameter(Mandatory = $false)]
    [ValidateSet('none', 'low', 'medium', 'high', 'critical')]
    [string] $AlertMinimumSeverity = 'none',

    [Parameter(Mandatory = $false)]
    [string] $AlertRecipient,

    [Parameter(Mandatory = $false)]
    [string] $TeamsWebhookUrl,

    [Parameter(Mandatory = $false)]
    [ValidateRange(30, 3650)]
    [int] $RetentionDays = 180,

    [Parameter(Mandatory = $false)]
    [bool] $AlwaysSendReport = $false,

    [Parameter(Mandatory = $false)]
    [bool] $includeCopilotAndDataverse = $false,

    [Parameter(Mandatory = $false)]
    [bool] $IncludeMaesterReport = $false
)

[string] $AzureEnvironment = 'AzureCloud'
[string] $GraphEnvironment = 'Global'
[string] $ExchangeEnvironmentName = 'O365Default'
[string] $TeamsEnvironmentName = ''
[string] $ResultsContainerName = 'maester'
[string] $BlobPrefix = 'maester'
[int] $TrendRunCount = 10
$script:DriftMaesterVersion = '1.3.0'
$script:GraphRequestMaxAttempts = 5
$script:GraphRetryBaseSeconds = 5
$script:GraphRequestBodySoftLimitBytes = 3000000
$script:ReportArtifactRoot = $null
$ErrorActionPreference = 'Stop'
$script:DetectedStorageAccountName = $null
$script:DetectedStorageResourceGroupName = $null
$script:DetectedStorageSubscriptionId = $null
$script:GraphAppRoles = @()
$script:ConnectedManagedIdentityClientId = $null
$script:ConnectedTenantId = $null
$script:RunLogBuffer = [System.Collections.Generic.List[string]]::new()
$script:RunLogFlushed = $false

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

    $sharePointSuffix = switch ($GraphEnvironment) {
        'China' { 'sharepoint.cn' }
        'USGov' { 'sharepoint.us' }
        'USGovDoD' { 'sharepoint-mil.us' }
        'Germany' { 'sharepoint.de' }
        default { 'sharepoint.com' }
    }

    [PSCustomObject]@{
        Graph            = $graphResource
        ExchangeOnline   = $exchangeResource
        IPPS             = $ippsResource
        Teams            = '48ac35b8-9aa8-4d74-927d-1f4a14a0b239'
        SharePointSuffix = $sharePointSuffix
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

    $line = "[{0:u}] [{1}] {2}" -f (Get-Date), $Level.ToUpperInvariant(), $Message

    if (-not $script:RunLogBuffer) {
        $script:RunLogBuffer = [System.Collections.Generic.List[string]]::new()
    }

    try {
        $script:RunLogBuffer.Add($line)
    } catch {
        # Last-resort fallback so logging never disappears entirely.
        Write-Output $line
    }
}

function Flush-RunLog {
    if ($script:RunLogFlushed) {
        return
    }

    if ($script:RunLogBuffer -and $script:RunLogBuffer.Count -gt 0) {
        foreach ($line in $script:RunLogBuffer) {
            Write-Output $line
        }
    }

    $script:RunLogFlushed = $true
}

function Get-RecipientList {
    param(
        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [AllowEmptyString()]
        [string] $Value
    )

    if ([string]::IsNullOrWhiteSpace($Value)) {
        return @()
    }

    @($Value -split ',' | ForEach-Object { $_.Trim() } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
}

function Get-AlertRecipientList {
    $alerts = Get-RecipientList -Value $AlertRecipient
    if ($alerts.Count -gt 0) {
        return $alerts
    }

    return (Get-RecipientList -Value $ReportRecipient)
}

function Invoke-GraphRequestWithRetry {
    param(
        [Parameter(Mandatory = $true)][ValidateSet('GET', 'POST', 'PUT', 'PATCH', 'DELETE')][string] $Method,
        [Parameter(Mandatory = $true)][string] $Uri,
        [Parameter(Mandatory = $false)][AllowNull()][object] $Body,
        [Parameter(Mandatory = $false)][string] $ContentType = 'application/json'
    )

    for ($attempt = 1; $attempt -le $script:GraphRequestMaxAttempts; $attempt++) {
        try {
            $params = @{ Method = $Method; Uri = $Uri; ErrorAction = 'Stop' }
            if ($PSBoundParameters.ContainsKey('Body')) { $params['Body'] = $Body }
            if (-not [string]::IsNullOrWhiteSpace($ContentType)) { $params['ContentType'] = $ContentType }
            return Invoke-MgGraphRequest @params
        } catch {
            $statusCode = $null
            if ($_.Exception.Response) {
                try { $statusCode = [int] $_.Exception.Response.StatusCode } catch { $statusCode = $null }
            }

            $isTransient = $statusCode -eq 408 -or $statusCode -eq 429 -or ($statusCode -ge 500 -and $statusCode -le 599)
            if (-not $isTransient -or $attempt -ge $script:GraphRequestMaxAttempts) {
                throw
            }

            $delaySeconds = [Math]::Min(60, [int]($script:GraphRetryBaseSeconds * [Math]::Pow(2, ($attempt - 1))))
            Write-RunLog "Graph request to '$Uri' failed with HTTP $statusCode. Retrying attempt $($attempt + 1)/$($script:GraphRequestMaxAttempts) in $delaySeconds second(s)." -Level Warning
            Start-Sleep -Seconds $delaySeconds
        }
    }
}

function Refresh-GraphConnection {
    Write-RunLog 'Refreshing Microsoft Graph token for post-scan operations.'
    $resourceUrls = Get-MaesterCloudResourceUrls
    $graphToken = Get-AzAccessTokenForResource -ResourceUrl $resourceUrls.Graph
    $graphTokenPlain = ConvertTo-PlainTextFromSecureString -SecureString $graphToken.Token
    $graphPayload = Get-JwtPayload -Token $graphTokenPlain
    $script:GraphAppRoles = @($graphPayload.roles) | Sort-Object -Unique

    $graphParams = @{ NoWelcome = $true }
    if ($GraphEnvironment -ne 'Global') { $graphParams['Environment'] = $GraphEnvironment }
    Connect-MgGraph -AccessToken $graphToken.Token @graphParams | Out-Null
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
    $domains = Invoke-GraphRequestWithRetry -Method GET -Uri "$graphRoot/v1.0/domains?`$select=id,isInitial"
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

function Get-SharePointAdminUrl {
    param([Parameter(Mandatory = $true)][string] $InitialDomain)

    $tenantPrefix = ($InitialDomain -split '\.')[0]
    if ([string]::IsNullOrWhiteSpace($tenantPrefix)) {
        throw "Could not derive the SharePoint admin URL from initial domain '$InitialDomain'."
    }

    return "https://$tenantPrefix-admin.$((Get-MaesterCloudResourceUrls).SharePointSuffix)"
}

# Maester connects to SharePoint Online through Connect-Maester -Service SharePointOnline, which only supports
# interactive, device code, or certificate thumbprint authentication. None of those work from an Automation Account,
# so the connection is made here instead: PnP.PowerShell is pointed at the tenant admin endpoint with the managed
# identity. Maester's Test-MtConnection and Get-MtSpo call Get-PnPConnection and Get-PnPTenant against whatever PnP
# connection exists in the session, so the SharePoint tests light up without any change to the Maester module.
# Both authentication paths need the SharePoint (not Graph) 'Sites.FullControl.All' application role on the managed
# identity, which Install-DriftMaester grants.
function Connect-SharePointOnlineForMaester {
    param([Parameter(Mandatory = $true)][string] $AdminUrl)

    Import-RequiredModule -Name PnP.PowerShell

    try {
        Write-RunLog "Connecting to SharePoint Online at '$AdminUrl' with the managed identity."
        Connect-PnPOnline -Url $AdminUrl -ManagedIdentity -ErrorAction Stop
    } catch {
        # PnP resolves the managed identity endpoint itself. When that is unavailable, fall back to the same
        # Az-issued token mechanism the Exchange and Teams connections use.
        Write-RunLog "PnP managed identity connection failed: $($_.Exception.Message). Retrying with an Az-issued SharePoint access token." -Level Warning
        $sharePointToken = ConvertTo-PlainTextFromSecureString -SecureString ((Get-AzAccessTokenForResource -ResourceUrl $AdminUrl).Token)
        Connect-PnPOnline -Url $AdminUrl -AccessToken $sharePointToken -ErrorAction Stop
    }

    if (-not (Test-MtConnection -Service SharePointOnline -ErrorAction SilentlyContinue)) {
        throw 'PnP reported a connection but Maester does not see it, so the SharePoint Online tests would be skipped.'
    }

    # Get-PnPTenant is what every Maester SharePoint test ends up calling, and it is the call that fails when the
    # Sites.FullControl.All application role is missing. Probing it here turns a silent test skip into a warning.
    Get-PnPTenant -ErrorAction Stop | Out-Null
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
                Service    = 'Exchange Online, Security & Compliance and SharePoint Online'
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

    if (-not [string]::IsNullOrWhiteSpace($initialDomain)) {
        try {
            Connect-SharePointOnlineForMaester -AdminUrl (Get-SharePointAdminUrl -InitialDomain $initialDomain)
        } catch {
            $missingServices.Add([PSCustomObject]@{
                    Service    = 'SharePoint Online'
                    Permission = "Sites.FullControl.All application role on the 'Office 365 SharePoint Online' service principal"
                    Type       = 'Application'
                    Reason     = "SharePoint Online managed identity connection failed: $($_.Exception.Message). Maester SharePoint Online tests will be skipped."
                })
        }
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
            Connect-ADOPS -Organization $devOpsOrganization -ManagedIdentity | Out-Null
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
        [string[]] $AttachmentPath,

        [Parameter(Mandatory = $false)]
        [string[]] $RecipientsOverride
    )

    $reportRecipients = if ($RecipientsOverride -and $RecipientsOverride.Count -gt 0) {
        @($RecipientsOverride | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | ForEach-Object { $_.Trim() } | Sort-Object -Unique)
    } else {
        @(Get-RecipientList -Value $ReportRecipient)
    }
    if ($reportRecipients.Count -eq 0) {
        throw 'Cannot send mail because ReportRecipient is empty. Provide one or more comma-separated email addresses.'
    }

    $senderUserId = if ($MailSenderUserId) { $MailSenderUserId } else { $reportRecipients[0] }

    $attachments = @()
    $attachedFileNames = [System.Collections.Generic.List[string]]::new()
    $bodySizeEstimate = 0
    $maxBodyEstimate = $script:GraphRequestBodySoftLimitBytes
    foreach ($path in @($AttachmentPath | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })) {
        if (-not (Test-Path -Path $path -PathType Leaf)) { continue }
        $bytes = [System.IO.File]::ReadAllBytes($path)
        $encodedSize = [Math]::Ceiling(([double]$bytes.Length / 3.0)) * 4
        if (($bodySizeEstimate + $encodedSize) -gt $maxBodyEstimate) {
            Write-RunLog "Skipping attachment '$([System.IO.Path]::GetFileName($path))' because Graph sendMail request size would exceed the safety limit." -Level Warning
            continue
        }

        $bodySizeEstimate += [int] $encodedSize
        $contentType = switch ([System.IO.Path]::GetExtension($path).ToLowerInvariant()) {
            '.zip' { 'application/zip' }
            '.html' { 'text/html' }
            default { 'application/octet-stream' }
        }
        $attachments += @{
            '@odata.type' = '#microsoft.graph.fileAttachment'
            name          = [System.IO.Path]::GetFileName($path)
            contentType   = $contentType
            contentBytes  = [System.Convert]::ToBase64String($bytes)
        }
        $attachedFileNames.Add([System.IO.Path]::GetFileName($path))
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
    Invoke-GraphRequestWithRetry -Method POST -Uri $sendMailUri -Body ($message | ConvertTo-Json -Depth 12) -ContentType 'application/json' | Out-Null

    return [PSCustomObject]@{
        AttachedFiles = @($attachedFileNames.ToArray())
        HasAttachments = $attachments.Count -gt 0
    }
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
<body><div class="wrap"><div class="hero"><span class="badge">Run aborted</span><h1>Maester drift detection could not start</h1><p>The managed identity is missing permissions or service connectivity required before running Maester. Re-run Install-DriftMaester.ps1 as Global Admin or manually add the items below and rerun the automation job.</p><div class="meta"><div><strong>Managed identity client id</strong><br>$(ConvertTo-HtmlEncodedText $clientId)</div><div><strong>Tenant id</strong><br>$(ConvertTo-HtmlEncodedText $tenant)</div></div></div><h2>Missing items</h2><ul>$permissionList</ul><table><thead><tr><th>Service</th><th>Permission or setting</th><th>Type</th><th>Details</th></tr></thead><tbody>$($rows -join [Environment]::NewLine)</tbody></table><p class="foot">Generated by Invoke-DriftMaester.ps1 on $(ConvertTo-HtmlEncodedText (Get-Date -Format 'yyyy-MM-dd HH:mm:ss K')).</p></div></body></html>
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
        $detailBits = @()
        if (-not [string]::IsNullOrWhiteSpace($warning.Permission)) { $detailBits += (ConvertTo-HtmlEncodedText $warning.Permission) }
        if (-not [string]::IsNullOrWhiteSpace($warning.Type)) { $detailBits += (ConvertTo-HtmlEncodedText $warning.Type) }
        $detailLine = if ($detailBits.Count -gt 0) { "<br><span style=`"color:#92400e;font-size:12px;`">$($detailBits -join ' &middot; ')</span>" } else { '' }
        '<tr><td style="padding:8px;border-bottom:1px solid #fde68a;vertical-align:top;"><strong>{0}</strong>{1}</td><td style="padding:8px;border-bottom:1px solid #fde68a;vertical-align:top;">{2}</td></tr>' -f `
            (ConvertTo-HtmlEncodedText $warning.Service),
            $detailLine,
            (ConvertTo-HtmlEncodedText $warning.Reason)
    }

    return @"
<tr><td style="padding:0 24px 12px 24px;"><table role="presentation" width="100%" cellspacing="0" cellpadding="0" style="border-collapse:collapse;background:#fffbeb;border:1px solid #fcd34d;border-radius:6px;"><tr><td style="padding:12px 12px 4px 12px;"><div style="font-size:12px;font-weight:700;color:#92400e;text-transform:uppercase;">Optional service warning</div><p style="margin:6px 0 8px 0;color:#78350f;">Some optional workload connections were not available. The run continued, but tests for those services may be skipped or have less detail.</p></td></tr><tr><td style="padding:0 12px 12px 12px;"><table width="100%" cellspacing="0" cellpadding="0" style="border-collapse:collapse;border:1px solid #fde68a;background:#fffaf0;"><thead><tr style="background:#fef3c7;"><th align="left" style="padding:8px;color:#78350f;font-size:12px;">Service &amp; missing permission or setting</th><th align="left" style="padding:8px;color:#78350f;font-size:12px;">Details</th></tr></thead><tbody>$($warningRows -join [Environment]::NewLine)</tbody></table></td></tr></table></td></tr>
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

function Resolve-StorageContext {
    param([Parameter(Mandatory = $true)][AllowNull()][object] $Value)

    if ($null -eq $Value) {
        throw 'Storage context resolution failed because no context object was returned.'
    }

    $candidates = @($Value)
    if ($candidates.Count -eq 1) {
        return $candidates[0]
    }

    foreach ($candidate in $candidates) {
        if ($null -eq $candidate) { continue }

        $typeName = $candidate.GetType().FullName
        if ($typeName -like '*StorageContext*' -or ($candidate.PSObject.Properties.Name -contains 'StorageAccountName')) {
            Write-RunLog "Storage context normalization selected '$typeName' from multi-object pipeline output." -Level Warning
            return $candidate
        }
    }

    $typeSummary = @($candidates | ForEach-Object {
            if ($null -eq $_) { 'null' } else { $_.GetType().FullName }
        }) -join ', '
    throw "Storage context resolution failed. Expected a storage context object but received: $typeSummary"
}

function Get-AutomationAccountFromEnvironment {
    if ([string]::IsNullOrWhiteSpace($env:AUTOMATION_ACCOUNT_ID)) {
        return $null
    }

    $resourceId = [string] $env:AUTOMATION_ACCOUNT_ID
    if ($resourceId -notmatch '^/subscriptions/([^/]+)/resourceGroups/([^/]+)/providers/Microsoft\.Automation/automationAccounts/([^/]+)$') {
        Write-RunLog "AUTOMATION_ACCOUNT_ID has unexpected format '$resourceId'. Falling back to discovery." -Level Warning
        return $null
    }

    return [PSCustomObject]@{
        Name              = $Matches[3]
        ResourceGroupName = $Matches[2]
        SubscriptionId    = $Matches[1]
        ResourceId        = $resourceId
    }
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
    $servicePrincipals = Invoke-GraphRequestWithRetry -Method GET -Uri "$graphRoot/v1.0/servicePrincipals?`$filter=$filter&`$select=id,appId"
    $servicePrincipal = @($servicePrincipals.value | Select-Object -First 1)
    if (-not $servicePrincipal) {
        throw "Could not find a service principal for managed identity client id '$clientId'."
    }

    return $servicePrincipal.id
}

function Find-CurrentAutomationAccount {
    $fromEnvironment = Get-AutomationAccountFromEnvironment
    if ($fromEnvironment) {
        return $fromEnvironment
    }

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

function Save-RunSummaryBlob {
    param(
        [Parameter(Mandatory = $true)][object] $StorageContext,
        [Parameter(Mandatory = $true)][string] $BlobName,
        [Parameter(Mandatory = $true)][object] $Result,
        [Parameter(Mandatory = $true)][string] $ModuleVersion
    )

    $summary = [PSCustomObject]@{
        TenantId         = $Result.TenantId
        TenantName       = $Result.TenantName
        ExecutedAt       = $Result.ExecutedAt
        PassedCount      = [int] $Result.PassedCount
        FailedCount      = [int] $Result.FailedCount
        ErrorCount       = [int] $Result.ErrorCount
        InvestigateCount = [int] $Result.InvestigateCount
        SkippedCount     = [int] $Result.SkippedCount
        NotRunCount      = [int] $Result.NotRunCount
        TotalCount       = [int] $Result.TotalCount
        Score            = Get-ScoreFromResult -Result $Result
        ModuleVersion    = $ModuleVersion
    }

    $tempPath = Join-Path -Path ([System.IO.Path]::GetTempPath()) -ChildPath ("DriftSummary-{0}.json" -f [guid]::NewGuid().ToString('N'))
    try {
        $summary | ConvertTo-Json -Depth 8 | Out-File -FilePath $tempPath -Encoding UTF8
        Save-BlobFile -StorageContext $StorageContext -FilePath $tempPath -BlobName $BlobName
    } finally {
        if (Test-Path -Path $tempPath) {
            Remove-Item -Path $tempPath -Force -ErrorAction SilentlyContinue
        }
    }
}

function Get-HistoricalSummaryBlobs {
    param(
        [Parameter(Mandatory = $true)][object] $StorageContext,
        [Parameter(Mandatory = $true)][string] $TenantSummaryPrefix
    )

    @(Get-AzStorageBlob -Context $StorageContext -Container $ResultsContainerName -Prefix $TenantSummaryPrefix -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -match '/summary-[0-9]{8}-[0-9]{6}\.json$' } |
        Sort-Object @{ Expression = { $_.LastModified.UtcDateTime }; Descending = $true }, Name)
}

function Read-SummaryBlob {
    param(
        [Parameter(Mandatory = $true)][object] $StorageContext,
        [Parameter(Mandatory = $true)][string] $BlobName,
        [Parameter(Mandatory = $true)][string] $DestinationFolder
    )

    $fileName = Split-Path -Path $BlobName -Leaf
    $destination = Join-Path -Path $DestinationFolder -ChildPath $fileName
    Get-AzStorageBlobContent -Context $StorageContext -Container $ResultsContainerName -Blob $BlobName -Destination $destination -Force | Out-Null
    Get-Content -Path $destination -Raw | ConvertFrom-Json
}

function Send-TeamsNotification {
    param(
        [Parameter(Mandatory = $true)][string] $Subject,
        [Parameter(Mandatory = $true)][object] $CurrentResult,
        [Parameter(Mandatory = $false)][object] $PreviousResult,
        [Parameter(Mandatory = $true)][object] $Diff,
        [Parameter(Mandatory = $true)][string] $DriftBlobName,
        [Parameter(Mandatory = $true)][string] $RecipientSummary
    )

    if ([string]::IsNullOrWhiteSpace($TeamsWebhookUrl)) {
        return
    }

    $passedDelta = if ($PreviousResult) { [int] $CurrentResult.PassedCount - [int] $PreviousResult.PassedCount } else { 0 }
    $passedDeltaSuffix = if ($passedDelta -gt 0) { " (+$passedDelta)" } elseif ($passedDelta -lt 0) { " ($passedDelta)" } else { '' }
    $findingCount = [int] $CurrentResult.FailedCount + [int] $CurrentResult.ErrorCount + [int] $CurrentResult.InvestigateCount
    $reportUrl = "https://$($script:DetectedStorageAccountName).blob.core.windows.net/$ResultsContainerName/$DriftBlobName"

    $payload = @{
        Report = @{
            Subject   = $Subject
            Tenant    = $($CurrentResult.TenantName)
            Passed    = "$($CurrentResult.PassedCount / $CurrentResult.TotalCount)$($passedDeltaSuffix)"
            Findings  = $findingCount
            Regressed = $($Diff.Summary.Regressed)
            ReportUrl = $reportUrl
        }
    
        Recipients = @(
            $RecipientSummary
        )
    } | ConvertTo-Json -Depth 6

    try {
        Invoke-RestMethod -Method POST -Uri $TeamsWebhookUrl -Body $payload -ContentType 'application/json' | Out-Null
    } catch {
        Write-RunLog "Teams webhook notification failed: $($_.Exception.Message)" -Level Warning
    }
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

function Enforce-ResultRetention {
    param(
        [Parameter(Mandatory = $true)][object] $StorageContext,
        [Parameter(Mandatory = $true)][string] $TenantPrefix,
        [Parameter(Mandatory = $true)][ValidateRange(30, 3650)][int] $RetentionDays
    )

    $cutoff = [datetime]::UtcNow.AddDays(-1 * $RetentionDays)
    $blobs = @(Get-AzStorageBlob -Context $StorageContext -Container $ResultsContainerName -Prefix $TenantPrefix -ErrorAction SilentlyContinue)
    $deleted = 0
    foreach ($blob in $blobs) {
        if (-not $blob.LastModified) { continue }
        if ($blob.LastModified.UtcDateTime -ge $cutoff) { continue }

        try {
            Remove-AzStorageBlob -Context $StorageContext -Container $ResultsContainerName -Blob $blob.Name -Force | Out-Null
            $deleted++
        } catch {
            Write-RunLog "Retention cleanup could not delete blob '$($blob.Name)': $($_.Exception.Message)" -Level Warning
        }
    }

    if ($deleted -gt 0) {
        Write-RunLog "Retention cleanup removed $deleted blob(s) older than $RetentionDays day(s) under '$TenantPrefix'."
    }
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
    $height = 240
    $paddingLeft = 54
    $paddingRight = 20
    $paddingTop = 26
    $paddingBottom = 44
    $plotWidth = $width - $paddingLeft - $paddingRight
    $plotHeight = $height - $paddingTop - $paddingBottom
    $maxIndex = [Math]::Max(1, $Trend.Count - 1)

    # The trend follows the number of passed tests, not the score percentage. Maester keeps adding tests for
    # workloads a given tenant does not use; those land in Skipped/NotRun, inflate TotalCount and drag the
    # percentage down without anything in the tenant having changed. A count of passing tests only moves when
    # the tenant's posture or the test set actually changes.
    $passedCounts = @($Trend | ForEach-Object { [int] $_.PassedCount })
    $minPassed = ($passedCounts | Measure-Object -Minimum).Minimum
    $maxPassed = ($passedCounts | Measure-Object -Maximum).Maximum

    # Zoom the y-axis to the observed range so small changes are clearly visible, while keeping a sensible
    # minimum span so a near-flat line is not over-amplified.
    $tick = 5
    $axisMin = [Math]::Max(0, [Math]::Floor(($minPassed - 2) / $tick) * $tick)
    $axisMax = [Math]::Ceiling(($maxPassed + 2) / $tick) * $tick
    if (($axisMax - $axisMin) -lt 10) {
        $axisMin = [Math]::Max(0, $axisMax - 10)
        if (($axisMax - $axisMin) -lt 10) { $axisMax = $axisMin + 10 }
    }
    $axisSpan = [double]($axisMax - $axisMin)
    if ($axisSpan -le 0) { $axisSpan = 10 }

    $yFor = { param($passed) $paddingTop + ($plotHeight - ((([double] $passed - $axisMin) / $axisSpan) * $plotHeight)) }
    $xFor = { param($index) $paddingLeft + (($plotWidth / $maxIndex) * $index) }

    $coords = for ($index = 0; $index -lt $Trend.Count; $index++) {
        [PSCustomObject]@{
            X      = [Math]::Round((& $xFor $index), 1)
            Y      = [Math]::Round((& $yFor $Trend[$index].PassedCount), 1)
            Passed = [int] $Trend[$index].PassedCount
            Total  = [int] $Trend[$index].TotalCount
            Label  = $Trend[$index].Label
            Delta  = if ($index -eq 0) { $null } else { [int] $Trend[$index].PassedCount - [int] $Trend[$index - 1].PassedCount }
        }
    }

    $points = ($coords | ForEach-Object { '{0},{1}' -f $_.X, $_.Y }) -join ' '
    $baselineY = $height - $paddingBottom
    $areaPoints = '{0},{1} {2} {3},{1}' -f $paddingLeft, $baselineY, $points, ($width - $paddingRight)

    $midPassed = [Math]::Round(($axisMin + $axisMax) / 2, 0)
    $midY = [Math]::Round((& $yFor $midPassed), 1)

    $circles = foreach ($c in $coords) {
        $fill = if ($null -eq $c.Delta -or $c.Delta -eq 0) { '#2563eb' } elseif ($c.Delta -gt 0) { '#059669' } else { '#dc2626' }
        $deltaText = if ($null -eq $c.Delta) { 'baseline' } elseif ($c.Delta -gt 0) { "+$($c.Delta)" } else { [string] $c.Delta }
        '<circle cx="{0}" cy="{1}" r="4.5" style="fill:{2}"><title>{3}: {4} of {5} tests passed ({6})</title></circle>' -f $c.X, $c.Y, $fill, (ConvertTo-HtmlEncodedText $c.Label), $c.Passed, $c.Total, $deltaText
    }

    # Numeric value above each point (nudged below the point when it sits near the top edge).
    $valueLabels = foreach ($c in $coords) {
        $ty = if ($c.Y -lt ($paddingTop + 14)) { $c.Y + 17 } else { $c.Y - 9 }
        '<text class="pointval" x="{0}" y="{1}" text-anchor="middle">{2}</text>' -f $c.X, $ty, $c.Passed
    }

    $xLabels = for ($index = 0; $index -lt $Trend.Count; $index++) {
        if ($index -eq 0 -or $index -eq ($Trend.Count - 1) -or $index % 2 -eq 0) {
            '<text x="{0}" y="{1}" text-anchor="middle">{2}</text>' -f [Math]::Round((& $xFor $index), 1), ($baselineY + 17), (ConvertTo-HtmlEncodedText $Trend[$index].Label)
        }
    }

    return @"
<svg class="trend" viewBox="0 0 $width $height" role="img" aria-label="Passed test trend over previous Maester runs">
<defs><linearGradient id="trendfill" x1="0" y1="0" x2="0" y2="1"><stop offset="0%" stop-color="#2563eb" stop-opacity="0.18"/><stop offset="100%" stop-color="#2563eb" stop-opacity="0"/></linearGradient></defs>
<line class="grid" x1="$paddingLeft" y1="$paddingTop" x2="$($width - $paddingRight)" y2="$paddingTop" />
<line class="grid" x1="$paddingLeft" y1="$midY" x2="$($width - $paddingRight)" y2="$midY" />
<line x1="$paddingLeft" y1="$paddingTop" x2="$paddingLeft" y2="$baselineY" />
<line x1="$paddingLeft" y1="$baselineY" x2="$($width - $paddingRight)" y2="$baselineY" />
<text class="axis" x="$($paddingLeft - 8)" y="$($paddingTop + 4)" text-anchor="end">$axisMax</text>
<text class="axis" x="$($paddingLeft - 8)" y="$($midY + 4)" text-anchor="end">$midPassed</text>
<text class="axis" x="$($paddingLeft - 8)" y="$($baselineY + 4)" text-anchor="end">$axisMin</text>
<polygon points="$areaPoints" style="fill:url(#trendfill);stroke:none" />
<polyline points="$points" />
$($circles -join [Environment]::NewLine)
$($valueLabels -join [Environment]::NewLine)
$($xLabels -join [Environment]::NewLine)
</svg>
<div class="trendnote">Passed tests per run. Y-axis is zoomed to $axisMin&ndash;$axisMax so small changes stay visible. Point colour shows improvement (green) or regression (red) versus the prior run; the number on each point is that run's passed test count.</div>
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

function New-MaesterReviewCellHtml {
    param(
        [AllowNull()][object] $Test,
        [Parameter(Mandatory = $true)][string] $DetailId
    )

    if (-not $Test) {
        return '<span class="muted">-</span>'
    }

    $errorText = Get-MaesterTestErrorText -Test $Test
    $detailHtml = Get-MaesterTestDetailHtml -Test $Test
    $hasDetails = -not [string]::IsNullOrWhiteSpace($detailHtml)
    $hasError = -not [string]::IsNullOrWhiteSpace($errorText)

    if (-not ($hasDetails -or $hasError)) {
        return '<span class="muted">-</span>'
    }

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

    return '<button type="button" class="detail-open" data-detail-id="{0}">Review</button><div id="{0}" class="modal-source" hidden data-title="{1}" data-id="{2}" data-has-details="{3}" data-has-error="{4}">{5}{6}</div>' -f `
        (ConvertTo-HtmlEncodedText $DetailId),
        (ConvertTo-HtmlEncodedText $Test.Title),
        (ConvertTo-HtmlEncodedText $Test.Id),
        ([string] $hasDetails).ToLowerInvariant(),
        ([string] $hasError).ToLowerInvariant(),
        $detailPanel,
        $errorPanel
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
        $detailText = Get-MaesterTestDetailSearchText -Test $test
        $reviewCell = New-MaesterReviewCellHtml -Test $test -DetailId ('test-detail-{0}' -f $rowIndex)
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
    $passed = [int] $CurrentResult.PassedCount
    $passedDelta = if ($PreviousResult) { $passed - [int] $PreviousResult.PassedCount } else { $null }
    $passedDeltaText = if ($null -ne $passedDelta) { if ($passedDelta -gt 0) { "+$passedDelta" } else { [string] $passedDelta } } else { 'No previous run' }
    $passedClass = if ($null -eq $passedDelta) { 'neutral' } elseif ($passedDelta -lt 0) { 'bad' } elseif ($passedDelta -gt 0) { 'good' } else { 'neutral' }

    $findings = @($CurrentResult.Tests | Where-Object { $_.Result -in @('Failed', 'Error', 'Investigate') } | Sort-Object Result, Severity, Id)
    $allTestRows = New-MaesterTestTableRows -Tests @($CurrentResult.Tests)

    # Lookups so each drift row can offer the same Review popup as the all-tests table.
    # Prefer the current run's test detail; fall back to the previous run for removed tests.
    $currentTestByKey = @{}
    foreach ($test in @($CurrentResult.Tests)) { $currentTestByKey[(Get-TestIdentityKey -Test $test)] = $test }
    $previousTestByKey = @{}
    if ($PreviousResult) {
        foreach ($test in @($PreviousResult.Tests)) { $previousTestByKey[(Get-TestIdentityKey -Test $test)] = $test }
    }

    $diffRows = if ($Diff.HasPrevious -and $Diff.Items.Count -gt 0) {
        $diffIndex = 0
        foreach ($item in @($Diff.Items | Sort-Object Classification, Id)) {
            $diffIndex++
            $reviewTest = if ($currentTestByKey.ContainsKey($item.Key)) { $currentTestByKey[$item.Key] } elseif ($previousTestByKey.ContainsKey($item.Key)) { $previousTestByKey[$item.Key] } else { $null }
            $reviewCell = New-MaesterReviewCellHtml -Test $reviewTest -DetailId ('drift-detail-{0}' -f $diffIndex)
            '<tr><td><span class="pill drift-{0}">{1}</span></td><td>{2}</td><td>{3}</td><td>{4}</td><td>{5}</td><td>{6}</td><td>{7}</td></tr>' -f `
                (([string] $item.Classification).ToLowerInvariant().Replace(' ', '-')),
                (ConvertTo-HtmlEncodedText $item.Classification),
                (ConvertTo-HtmlEncodedText $item.Id),
                (ConvertTo-HtmlEncodedText $item.Title),
                (ConvertTo-HtmlEncodedText $item.PreviousResult),
                (ConvertTo-HtmlEncodedText $item.CurrentResult),
                (ConvertTo-HtmlEncodedText $item.Severity),
                $reviewCell
        }
    } elseif ($Diff.HasPrevious) {
        '<tr><td colspan="7" class="empty">No drift detected compared with the previous run.</td></tr>'
    } else {
        '<tr><td colspan="7" class="empty">No previous result was present in blob storage, so no diff was calculated for this first run.</td></tr>'
    }

    $trendRows = foreach ($point in $Trend) {
        '<tr><td>{0}</td><td><strong>{1}</strong></td><td>{2}</td><td>{3}</td><td>{4}</td><td>{5}</td><td class="muted">{6}%</td></tr>' -f `
            (ConvertTo-HtmlEncodedText $point.Label),
            (ConvertTo-HtmlEncodedText $point.PassedCount),
            (ConvertTo-HtmlEncodedText $point.FailedCount),
            (ConvertTo-HtmlEncodedText $point.ErrorCount),
            (ConvertTo-HtmlEncodedText $point.InvestigateCount),
            (ConvertTo-HtmlEncodedText $point.TotalCount),
            (ConvertTo-HtmlEncodedText $point.Score)
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
body{margin:0;background:#f3f6fb;color:#172033;font-family:Segoe UI,Arial,sans-serif;line-height:1.45}.wrap{max-width:1280px;margin:0 auto;padding:28px}.hero,.card,table,ul.blobs,.trend,.filters{background:#fff;border:1px solid #d9e2ef;border-radius:8px}.hero{padding:24px}.eyebrow{font-size:12px;font-weight:700;color:#2563eb;text-transform:uppercase;letter-spacing:.04em}h1{font-size:26px;margin:8px 0 10px}h2{font-size:18px;margin:28px 0 10px}.grid{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:12px;margin-top:18px}.card{padding:16px}.card .label{font-size:12px;color:#64748b;text-transform:uppercase;font-weight:700}.card .value{font-size:28px;font-weight:700;margin-top:4px}.good{color:#047857}.bad{color:#b91c1c}.neutral{color:#475569}.meta{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:10px;margin-top:18px}.meta div{background:#f8fafc;border:1px solid #e5eaf2;border-radius:6px;padding:10px}.pill{display:inline-block;border-radius:999px;padding:4px 8px;font-size:12px;font-weight:700}.result-Passed{background:#dcfce7;color:#166534}.result-Failed,.result-Error,.drift-regressed,.drift-new-finding{background:#fee2e2;color:#991b1b}.result-Investigate,.result-Skipped,.result-NotRun,.drift-changed,.drift-added-test{background:#fef3c7;color:#92400e}.drift-improved{background:#dcfce7;color:#166534}.drift-removed-test{background:#e0f2fe;color:#075985}table{border-collapse:separate;border-spacing:0;width:100%;overflow:hidden}th,td{text-align:left;padding:10px;border-bottom:1px solid #edf1f7;vertical-align:top}tr:last-child td{border-bottom:0}th{background:#eef4fb;color:#334155;font-size:12px;text-transform:uppercase}.empty{color:#64748b;text-align:center;padding:18px}.muted{color:#94a3b8}.trend{width:100%;height:auto}.trend line{stroke:#cbd5e1;stroke-width:1}.trend line.grid{stroke:#e2e8f0;stroke-width:1;stroke-dasharray:3 3}.trend polyline{fill:none;stroke:#2563eb;stroke-width:3}.trend circle{fill:#2563eb}.trend text{fill:#64748b;font-size:11px}.trend text.pointval{fill:#1e293b;font-size:11px;font-weight:700}.trend text.axis{fill:#94a3b8;font-size:10px}.trendnote{font-size:12px;color:#64748b;margin:8px 0 0 0}ul.blobs{padding:14px 14px 14px 30px}.blobs span{color:#64748b;font-size:12px}.foot{font-size:12px;color:#64748b;margin-top:20px}.filters{display:flex;gap:10px;align-items:center;flex-wrap:wrap;padding:12px;margin-bottom:10px}.filters input,.filters select{border:1px solid #cbd5e1;border-radius:6px;padding:8px 10px;font:inherit}.filters input{min-width:280px;flex:1}.filters label{font-size:12px;font-weight:700;color:#475569;text-transform:uppercase}#allTests{table-layout:fixed}#allTests th:nth-child(1),#allTests td:nth-child(1){width:86px}#allTests th:nth-child(2),#allTests td:nth-child(2){width:76px;word-break:break-word;font-size:12px}#allTests th:nth-child(4),#allTests td:nth-child(4){width:82px}#allTests th:nth-child(5),#allTests td:nth-child(5){width:126px}#allTests th:nth-child(6),#allTests td:nth-child(6){width:82px}#allTests th:nth-child(7),#allTests td:nth-child(7){width:56px}#allTests th:nth-child(8),#allTests td:nth-child(8){width:92px}.detail-open{border:1px solid #bfdbfe;background:#eff6ff;color:#1d4ed8;border-radius:6px;padding:6px 10px;font:inherit;font-weight:700;cursor:pointer}.detail-open:hover{background:#dbeafe}.modal-backdrop{position:fixed;inset:0;background:rgba(15,23,42,.62);display:none;align-items:center;justify-content:center;padding:24px;z-index:999}.modal-backdrop.open{display:flex}.modal{background:#fff;border-radius:8px;box-shadow:0 24px 80px rgba(15,23,42,.35);width:min(1120px,96vw);max-height:88vh;display:flex;flex-direction:column;overflow:hidden}.modal-head{display:flex;justify-content:space-between;gap:14px;align-items:flex-start;padding:18px 20px;border-bottom:1px solid #e5eaf2}.modal-title{font-size:18px;font-weight:700}.modal-subtitle{font-size:12px;color:#64748b;margin-top:4px}.modal-close{border:0;background:#f1f5f9;color:#334155;border-radius:6px;padding:7px 10px;cursor:pointer;font:inherit}.modal-tabs{display:flex;gap:8px;padding:12px 20px 0;border-bottom:1px solid #e5eaf2}.modal-tab{border:1px solid #cbd5e1;border-bottom:0;background:#f8fafc;color:#475569;border-radius:6px 6px 0 0;padding:8px 12px;cursor:pointer;font:inherit;font-weight:700}.modal-tab.active{background:#fff;color:#1d4ed8}.modal-body{padding:18px 20px;overflow:auto}.modal-panel{display:none}.modal-panel.active{display:block}.detail-block{margin:0 0 14px}.detail-label{color:#475569;font-size:12px;font-weight:700;text-transform:uppercase}.modal pre{white-space:pre-wrap;word-break:break-word;border-radius:6px;margin:6px 0 0 0;padding:10px;background:#f8fafc;border:1px solid #dbe6f3;color:#172033}.modal-panel[data-panel="error"] pre{background:#fff7ed;border-color:#fed7aa;color:#7c2d12}a{color:#2563eb;text-decoration:none}a:hover{text-decoration:underline}@media(max-width:860px){.grid,.meta{grid-template-columns:1fr 1fr}#allTests{table-layout:auto}.modal{width:98vw;max-height:92vh}}@media(max-width:560px){.grid,.meta{grid-template-columns:1fr}.wrap{padding:16px}.filters input{min-width:0;width:100%}.modal-backdrop{padding:10px}.modal-head{padding:14px}.modal-tabs{padding-left:14px}.modal-body{padding:14px}}
</style>
</head>
<body><div class="wrap"><div class="hero"><div class="eyebrow">Maester $(ConvertTo-HtmlEncodedText $ModuleVersion) report</div><h1>$(ConvertTo-HtmlEncodedText $CurrentResult.TenantName)</h1><p>Executed at $(ConvertTo-HtmlEncodedText ([datetime] $CurrentResult.ExecutedAt).ToString('yyyy-MM-dd HH:mm:ss K')). Previous run: $(ConvertTo-HtmlEncodedText $previousRunText).</p><div class="meta"><div><strong>Tenant id</strong><br>$(ConvertTo-HtmlEncodedText $CurrentResult.TenantId)</div><div><strong>Maester version</strong><br>$(ConvertTo-HtmlEncodedText $ModuleVersion)</div><div><strong>Storage account</strong><br>$(ConvertTo-HtmlEncodedText $script:DetectedStorageAccountName)</div></div></div><div class="grid"><div class="card"><div class="label">Passed tests</div><div class="value good">$passed</div></div><div class="card"><div class="label">Passed delta</div><div class="value $passedClass">$passedDeltaText</div></div><div class="card"><div class="label">Findings</div><div class="value bad">$($findings.Count)</div></div><div class="card"><div class="label">Total tests</div><div class="value">$($CurrentResult.TotalCount)</div></div><div class="card"><div class="label">Failed</div><div class="value bad">$($CurrentResult.FailedCount)</div></div><div class="card"><div class="label">Errors</div><div class="value bad">$($CurrentResult.ErrorCount)</div></div><div class="card"><div class="label">Investigate</div><div class="value neutral">$($CurrentResult.InvestigateCount)</div></div><div class="card"><div class="label">Score</div><div class="value neutral">$score%</div></div></div><h2>Drift since previous run</h2><table><thead><tr><th>Status</th><th>Id</th><th>Title</th><th>Previous</th><th>Current</th><th>Severity</th><th>Review</th></tr></thead><tbody>$($diffRows -join [Environment]::NewLine)</tbody></table><h2>Passed test trend</h2>$trendSvg<table><thead><tr><th>Run</th><th>Passed</th><th>Failed</th><th>Errors</th><th>Investigate</th><th>Total</th><th>Score</th></tr></thead><tbody>$($trendRows -join [Environment]::NewLine)</tbody></table><h2>All Maester tests</h2><div class="filters"><label for="resultFilter">Result</label><select id="resultFilter"><option value="">All</option><option>Passed</option><option>Failed</option><option>Error</option><option>Investigate</option><option>Skipped</option><option>NotRun</option></select><label for="testSearch">Search</label><input id="testSearch" type="search" placeholder="Search id, title, severity, service, evidence, or error"></div><table id="allTests"><thead><tr><th>Result</th><th>Id</th><th>Title</th><th>Severity</th><th>Service</th><th>Duration</th><th>Fix</th><th>Review</th></tr></thead><tbody>$allTestRows</tbody></table><h2>Stored artifacts</h2><ul class="blobs">$blobList</ul><p class="foot">Generated by Invoke-DriftMaester.ps1. Native Maester JSON, HTML and Markdown are stored unchanged in the '$ResultsContainerName' blob container.</p></div><div id="detailModal" class="modal-backdrop" role="dialog" aria-modal="true" aria-hidden="true"><div class="modal"><div class="modal-head"><div><div id="modalTitle" class="modal-title">Test details</div><div id="modalSubtitle" class="modal-subtitle"></div></div><button type="button" class="modal-close">Close</button></div><div id="modalTabs" class="modal-tabs"></div><div id="modalBody" class="modal-body"></div></div></div><script>(function(){var result=document.getElementById('resultFilter');var search=document.getElementById('testSearch');var rows=[].slice.call(document.querySelectorAll('#allTests tbody tr[data-result]'));function apply(){var selected=(result.value||'').toLowerCase();var query=(search.value||'').toLowerCase();rows.forEach(function(row){var okResult=!selected||row.getAttribute('data-result').toLowerCase()===selected;var okSearch=!query||(row.getAttribute('data-search')||'').indexOf(query)>=0;row.style.display=okResult&&okSearch?'':'none';});}if(result){result.addEventListener('change',apply);}if(search){search.addEventListener('input',apply);}var modal=document.getElementById('detailModal');var modalTitle=document.getElementById('modalTitle');var modalSubtitle=document.getElementById('modalSubtitle');var modalTabs=document.getElementById('modalTabs');var modalBody=document.getElementById('modalBody');function selectTab(name){[].slice.call(modalTabs.querySelectorAll('.modal-tab')).forEach(function(tab){tab.classList.toggle('active',tab.getAttribute('data-tab')===name);});[].slice.call(modalBody.querySelectorAll('.modal-panel')).forEach(function(panel){panel.classList.toggle('active',panel.getAttribute('data-panel')===name);});modalBody.scrollTop=0;}function closeModal(){modal.classList.remove('open');modal.setAttribute('aria-hidden','true');modalBody.innerHTML='';modalTabs.innerHTML='';}function openModal(source){modalTitle.textContent=source.getAttribute('data-title')||'Test details';modalSubtitle.textContent=source.getAttribute('data-id')?'Id: '+source.getAttribute('data-id'):'';modalBody.innerHTML=source.innerHTML;modalTabs.innerHTML='';var panels=[].slice.call(modalBody.querySelectorAll('.modal-panel'));panels.forEach(function(panel,index){var name=panel.getAttribute('data-panel');var label=name==='error'?'Technical error':'Details';var tab=document.createElement('button');tab.type='button';tab.className='modal-tab';tab.setAttribute('data-tab',name);tab.textContent=label;tab.addEventListener('click',function(){selectTab(name);});modalTabs.appendChild(tab);if(index===0){selectTab(name);}});modal.classList.add('open');modal.setAttribute('aria-hidden','false');}document.addEventListener('click',function(event){var opener=event.target.closest('.detail-open');if(opener){var source=document.getElementById(opener.getAttribute('data-detail-id'));if(source){openModal(source);}return;}if(event.target.classList.contains('modal-close')||event.target===modal){closeModal();}});document.addEventListener('keydown',function(event){if(event.key==='Escape'&&modal.classList.contains('open')){closeModal();}});}());</script></body></html>
"@
}

function New-TrendEmailChartRows {
    param([Parameter(Mandatory = $true)][object[]] $Trend)

    # Outlook (Word engine) cannot render SVG, so the trend is drawn as nested tables with
    # bgcolor attributes and fixed width/height cells. This is the most compatible, A/V-safe
    # way to show a bar chart in mail: no images, no external resources, no inline CSS tricks.
    $barMax = 280

    # Bars show passed test counts, scaled against the best run in this window. Bars start at zero, so without
    # that scaling every run would look identical once a tenant passes a few hundred tests.
    $peakPassed = [Math]::Max(1, (@($Trend | ForEach-Object { [int] $_.PassedCount }) | Measure-Object -Maximum).Maximum)

    $rows = for ($i = 0; $i -lt $Trend.Count; $i++) {
        $point = $Trend[$i]
        $passed = [int] $point.PassedCount
        $barWidth = [int][Math]::Round(($passed / [double] $peakPassed) * $barMax)
        if ($barWidth -lt 2) { $barWidth = 2 }
        if ($barWidth -gt $barMax) { $barWidth = $barMax }
        $restWidth = $barMax - $barWidth
        $delta = if ($i -eq 0) { $null } else { $passed - [int] $Trend[$i - 1].PassedCount }

        if ($null -eq $delta) {
            $barColor = '#2563eb'
            $deltaChip = '<span style="color:#64748b;font-size:12px;">baseline</span>'
        } elseif ($delta -gt 0) {
            $barColor = '#059669'
            $deltaChip = "<span style=`"color:#047857;font-weight:700;font-size:12px;`">&#9650; +$delta</span>"
        } elseif ($delta -lt 0) {
            $barColor = '#dc2626'
            $deltaChip = "<span style=`"color:#b91c1c;font-weight:700;font-size:12px;`">&#9660; $delta</span>"
        } else {
            $barColor = '#94a3b8'
            $deltaChip = '<span style="color:#64748b;font-size:12px;">&#9644; 0</span>'
        }

        $restCell = if ($restWidth -gt 0) {
            '<td height="14" width="{0}" bgcolor="#eef2f7" style="width:{0}px;height:14px;font-size:1px;line-height:14px;">&nbsp;</td>' -f $restWidth
        } else {
            ''
        }

        @"
<tr>
<td style="padding:5px 10px 5px 0;white-space:nowrap;color:#475569;font-size:12px;">$(ConvertTo-HtmlEncodedText $point.Label)</td>
<td style="padding:5px 0;">
<table role="presentation" cellspacing="0" cellpadding="0" border="0" style="border-collapse:collapse;border-radius:3px;overflow:hidden;"><tr><td height="14" width="$barWidth" bgcolor="$barColor" style="width:${barWidth}px;height:14px;background:$barColor;font-size:1px;line-height:14px;">&nbsp;</td>$restCell</tr></table>
</td>
<td style="padding:5px 0 5px 12px;white-space:nowrap;font-weight:700;color:#172033;font-size:13px;">$($point.PassedCount)</td>
<td style="padding:5px 0 5px 12px;white-space:nowrap;">$deltaChip</td>
</tr>
"@
    }

    return ($rows -join [Environment]::NewLine)
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

    $passed = [int] $CurrentResult.PassedCount
    $passedDelta = if ($PreviousResult) { $passed - [int] $PreviousResult.PassedCount } else { $null }
    $passedDeltaText = if ($null -ne $passedDelta) { if ($passedDelta -gt 0) { "+$passedDelta" } else { [string] $passedDelta } } else { 'No previous run' }
    $passedDeltaColor = if ($null -eq $passedDelta) { '#475569' } elseif ($passedDelta -lt 0) { '#b91c1c' } elseif ($passedDelta -gt 0) { '#047857' } else { '#475569' }
    $findingCount = @($CurrentResult.Tests | Where-Object { $_.Result -in @('Failed', 'Error', 'Investigate') }).Count
    $previousRunText = if ($PreviousResult) { ([datetime] $PreviousResult.ExecutedAt).ToString('yyyy-MM-dd HH:mm:ss K') } else { 'No previous run found' }

    $sortedDiffItems = @($Diff.Items | Sort-Object Classification, Id)
    $truncatedDiffItems = @($sortedDiffItems | Select-Object -First 25)
    $diffRows = if ($Diff.HasPrevious -and $sortedDiffItems.Count -gt 0) {
        foreach ($item in $truncatedDiffItems) {
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
        '<tr><td style="padding:8px;border-bottom:1px solid #e5eaf2;">{0}</td><td style="padding:8px;border-bottom:1px solid #e5eaf2;font-weight:700;">{1}</td><td style="padding:8px;border-bottom:1px solid #e5eaf2;">{2}</td><td style="padding:8px;border-bottom:1px solid #e5eaf2;">{3}</td><td style="padding:8px;border-bottom:1px solid #e5eaf2;">{4}</td><td style="padding:8px;border-bottom:1px solid #e5eaf2;color:#64748b;">{5}</td></tr>' -f `
            (ConvertTo-HtmlEncodedText $point.Label),
            (ConvertTo-HtmlEncodedText $point.PassedCount),
            (ConvertTo-HtmlEncodedText $point.FailedCount),
            (ConvertTo-HtmlEncodedText $point.ErrorCount),
            (ConvertTo-HtmlEncodedText $point.InvestigateCount),
            (ConvertTo-HtmlEncodedText $point.TotalCount)
    }

    $trendChartRows = if ($Trend -and $Trend.Count -gt 0) { New-TrendEmailChartRows -Trend $Trend } else { '' }
    $optionalWarningHtml = New-OptionalConnectionWarningEmailHtml -OptionalWarnings $OptionalWarnings

    $truncationNote = ''
    if ($Diff.HasPrevious -and $sortedDiffItems.Count -gt $truncatedDiffItems.Count) {
        $truncationNote = '<p style="margin:8px 24px;color:#92400e;">Showing first {0} of {1} drift items. See the full HTML report for all changes.</p>' -f $truncatedDiffItems.Count, $sortedDiffItems.Count
    }

    return @"
<!doctype html>
<html><body style="margin:0;background:#f5f7fb;color:#172033;font-family:Segoe UI,Arial,sans-serif;line-height:1.45;">
<table role="presentation" width="100%" cellspacing="0" cellpadding="0" style="background:#f5f7fb;"><tr><td align="center" style="padding:24px;">
<table role="presentation" width="760" cellspacing="0" cellpadding="0" style="width:760px;max-width:100%;background:#ffffff;border:1px solid #d9e2ef;border-radius:8px;">
<tr><td style="padding:24px 24px 8px 24px;"><div style="font-size:12px;font-weight:700;color:#2563eb;text-transform:uppercase;">Maester drift report</div><h1 style="font-size:24px;margin:8px 0 8px 0;color:#172033;">$(ConvertTo-HtmlEncodedText $CurrentResult.TenantName)</h1><p style="margin:0;color:#475569;">Executed at $(ConvertTo-HtmlEncodedText ([datetime] $CurrentResult.ExecutedAt).ToString('yyyy-MM-dd HH:mm:ss K')). Previous run: $(ConvertTo-HtmlEncodedText $previousRunText).</p></td></tr>
<tr><td style="padding:12px 24px;"><table role="presentation" width="100%" cellspacing="0" cellpadding="0"><tr><td style="padding:12px;background:#f8fafc;border:1px solid #e5eaf2;"><div style="font-size:12px;color:#64748b;font-weight:700;text-transform:uppercase;">Passed tests</div><div style="font-size:28px;font-weight:700;color:#047857;">$passed</div></td><td style="padding:12px;background:#f8fafc;border:1px solid #e5eaf2;"><div style="font-size:12px;color:#64748b;font-weight:700;text-transform:uppercase;">Passed delta</div><div style="font-size:28px;font-weight:700;color:$passedDeltaColor;">$passedDeltaText</div></td><td style="padding:12px;background:#f8fafc;border:1px solid #e5eaf2;"><div style="font-size:12px;color:#64748b;font-weight:700;text-transform:uppercase;">Findings</div><div style="font-size:28px;font-weight:700;color:#b91c1c;">$findingCount</div></td><td style="padding:12px;background:#f8fafc;border:1px solid #e5eaf2;"><div style="font-size:12px;color:#64748b;font-weight:700;text-transform:uppercase;">Total tests</div><div style="font-size:28px;font-weight:700;">$($CurrentResult.TotalCount)</div></td></tr></table></td></tr>
<tr><td style="padding:8px 24px;"><p style="margin:0;color:#475569;">The attached report (a zipped HTML file) contains all tests, passed results, documentation links, per-test review details, and browser filtering.</p></td></tr>
<tr><td style="padding:16px 24px 8px 24px;"><h2 style="font-size:17px;margin:0 0 8px 0;">Drift since previous run</h2><table width="100%" cellspacing="0" cellpadding="0" style="border-collapse:collapse;border:1px solid #d9e2ef;"><thead><tr style="background:#eef4fb;"><th align="left" style="padding:8px;">Status</th><th align="left" style="padding:8px;">Id</th><th align="left" style="padding:8px;">Title</th><th align="left" style="padding:8px;">Previous</th><th align="left" style="padding:8px;">Current</th></tr></thead><tbody>$($diffRows -join [Environment]::NewLine)</tbody></table></td></tr>
$truncationNote
<tr><td style="padding:16px 24px 24px 24px;"><h2 style="font-size:17px;margin:0 0 12px 0;">Passed test trend</h2><table role="presentation" width="100%" cellspacing="0" cellpadding="0" border="0" style="border-collapse:collapse;background:#f8fafc;border:1px solid #e5eaf2;border-radius:8px;"><tr><td style="padding:14px 18px;"><table role="presentation" cellspacing="0" cellpadding="0" border="0" style="border-collapse:collapse;">$trendChartRows</table></td></tr></table><table width="100%" cellspacing="0" cellpadding="0" style="border-collapse:collapse;border:1px solid #d9e2ef;margin-top:14px;"><thead><tr style="background:#eef4fb;"><th align="left" style="padding:8px;">Run</th><th align="left" style="padding:8px;">Passed</th><th align="left" style="padding:8px;">Failed</th><th align="left" style="padding:8px;">Errors</th><th align="left" style="padding:8px;">Investigate</th><th align="left" style="padding:8px;">Total</th></tr></thead><tbody>$($trendRows -join [Environment]::NewLine)</tbody></table><p style="font-size:12px;color:#64748b;margin:16px 0 0 0;">Bars show passed tests per run, scaled against the best run shown. Green bars/arrows mark runs that improved versus the prior run, red mark regressions. Maester version: $(ConvertTo-HtmlEncodedText $ModuleVersion). Storage account: $(ConvertTo-HtmlEncodedText $script:DetectedStorageAccountName).</p></td></tr>
$optionalWarningHtml
</table></td></tr></table>
</body></html>
"@
}

function Get-ReportLinkHtml {
    param(
        [Parameter(Mandatory = $true)][string] $StorageAccountName,
        [Parameter(Mandatory = $true)][string] $ContainerName,
        [Parameter(Mandatory = $true)][string] $BlobName
    )

    $encodedBlob = [System.Uri]::EscapeDataString($BlobName) -replace '%2F', '/'
    $url = "https://$StorageAccountName.blob.core.windows.net/$ContainerName/$encodedBlob"
    return '<p style="margin:10px 0;color:#334155;">Full report: <a href="{0}" target="_blank" rel="noopener">{1}</a></p>' -f (ConvertTo-HtmlEncodedText $url), (ConvertTo-HtmlEncodedText $BlobName)
}

function Get-SeverityWeight {
    param([Parameter(Mandatory = $true)][string] $Severity)

    switch ($Severity.ToLowerInvariant()) {
        'critical' { return 5 }
        'high' { return 4 }
        'medium' { return 3 }
        'low' { return 2 }
        default { return 1 }
    }
}

function Test-DiffMeetsSeverityThreshold {
    param(
        [Parameter(Mandatory = $true)][object] $Diff,
        [Parameter(Mandatory = $true)][string] $MinimumSeverity
    )

    if ([string]::IsNullOrWhiteSpace($MinimumSeverity) -or $MinimumSeverity -eq 'none') {
        return $true
    }

    $threshold = Get-SeverityWeight -Severity $MinimumSeverity
    foreach ($item in @($Diff.Items)) {
        if ($item.Classification -in @('Regressed', 'New finding', 'Changed')) {
            if (Get-SeverityWeight -Severity ([string]$item.Severity) -ge $threshold) {
                return $true
            }
        }
    }

    return $false
}

function Send-FailureNotification {
    param(
        [Parameter(Mandatory = $true)][string] $ErrorMessage,
        [Parameter(Mandatory = $false)][string] $StackTrace
    )

    $recipients = @(Get-AlertRecipientList)
    if ($recipients.Count -eq 0) {
        Write-RunLog 'Failure notification skipped because no alert recipients were configured.' -Level Warning
        return
    }

    $jobId = $env:AUTOMATION_JOB_ID
    $subject = "$MailSubjectPrefix FAILED: DriftMaester run error"
    $body = @"
<!doctype html>
<html><body style="font-family:Segoe UI,Arial,sans-serif;background:#f8fafc;color:#0f172a;">
<h2>DriftMaester run failed</h2>
<p>Tenant: $(ConvertTo-HtmlEncodedText $script:ConnectedTenantId)</p>
<p>Automation job id: $(ConvertTo-HtmlEncodedText $jobId)</p>
<p>Error: $(ConvertTo-HtmlEncodedText $ErrorMessage)</p>
<pre style="background:#111827;color:#e5e7eb;padding:12px;border-radius:6px;overflow:auto;">$(ConvertTo-HtmlEncodedText $StackTrace)</pre>
</body></html>
"@

    try {
        $existingReportRecipient = $ReportRecipient
        $ReportRecipient = ($recipients -join ',')
        $null = Send-DriftMail -Subject $subject -HtmlBody $body
        $ReportRecipient = $existingReportRecipient
    } catch {
        Write-RunLog "Failure notification mail attempt failed: $($_.Exception.Message)" -Level Warning
        $ReportRecipient = $existingReportRecipient
    }
}

function Initialize-WorkingTests {
    param([Parameter(Mandatory = $true)][string] $WorkingRoot)

    $testsWorkingPath = Join-Path -Path $WorkingRoot -ChildPath 'maester-tests'
    New-Item -Path $testsWorkingPath -ItemType Directory -Force | Out-Null

    Write-RunLog "Updating Maester tests in working folder '$testsWorkingPath'."
    Update-MaesterTests -Path $testsWorkingPath -Force | Out-Null

    return $testsWorkingPath
}

function Try-RestoreTestsCache {
    param(
        [Parameter(Mandatory = $true)][object] $StorageContext,
        [Parameter(Mandatory = $true)][string] $CacheBlobName,
        [Parameter(Mandatory = $true)][string] $TargetTestsPath
    )

    $cacheArchive = Join-Path -Path ([System.IO.Path]::GetTempPath()) -ChildPath ("DriftTestsCache-{0}.zip" -f [guid]::NewGuid().ToString('N'))
    try {
        Get-AzStorageBlobContent -Context $StorageContext -Container $ResultsContainerName -Blob $CacheBlobName -Destination $cacheArchive -Force -ErrorAction Stop | Out-Null
        Expand-Archive -Path $cacheArchive -DestinationPath $TargetTestsPath -Force
        Write-RunLog "Restored Maester test cache from '$CacheBlobName'." -Level Warning
        return $true
    } catch {
        return $false
    } finally {
        if (Test-Path -Path $cacheArchive) {
            Remove-Item -Path $cacheArchive -Force -ErrorAction SilentlyContinue
        }
    }
}

function Save-TestsCache {
    param(
        [Parameter(Mandatory = $true)][object] $StorageContext,
        [Parameter(Mandatory = $true)][string] $CacheBlobName,
        [Parameter(Mandatory = $true)][string] $TestsPath
    )

    $cacheArchive = Join-Path -Path ([System.IO.Path]::GetTempPath()) -ChildPath ("DriftTestsCache-{0}.zip" -f [guid]::NewGuid().ToString('N'))
    try {
        Compress-Archive -Path (Join-Path -Path $TestsPath -ChildPath '*') -DestinationPath $cacheArchive -Force
        Save-BlobFile -StorageContext $StorageContext -FilePath $cacheArchive -BlobName $CacheBlobName
    } catch {
        Write-RunLog "Could not persist Maester test cache: $($_.Exception.Message)" -Level Warning
    } finally {
        if (Test-Path -Path $cacheArchive) {
            Remove-Item -Path $cacheArchive -Force -ErrorAction SilentlyContinue
        }
    }
}

function Apply-TenantMaesterConfig {
    param(
        [Parameter(Mandatory = $true)][object] $StorageContext,
        [Parameter(Mandatory = $true)][string] $TenantPrefix,
        [Parameter(Mandatory = $true)][string] $TestsPath
    )

    $configBlob = "$TenantPrefix/config/maester-config.json"
    $configPath = Join-Path -Path $TestsPath -ChildPath 'maester-config.json'
    try {
        Get-AzStorageBlobContent -Context $StorageContext -Container $ResultsContainerName -Blob $configBlob -Destination $configPath -Force -ErrorAction Stop | Out-Null
        Write-RunLog "Applied tenant Maester config from '$configBlob'."
    } catch {
        Write-RunLog "No tenant Maester config found at '$configBlob'. Default Maester config will be used."
    }

    $customPrefix = "$TenantPrefix/config/custom-tests/"
    $customBlobs = @(Get-AzStorageBlob -Context $StorageContext -Container $ResultsContainerName -Prefix $customPrefix -ErrorAction SilentlyContinue)
    if ($customBlobs.Count -gt 0) {
        foreach ($blob in $customBlobs) {
            $relativePath = $blob.Name.Substring($customPrefix.Length)
            if ([string]::IsNullOrWhiteSpace($relativePath)) { continue }
            $destination = Join-Path -Path $TestsPath -ChildPath $relativePath
            $destinationFolder = Split-Path -Path $destination -Parent
            if (-not (Test-Path -Path $destinationFolder)) {
                New-Item -Path $destinationFolder -ItemType Directory -Force | Out-Null
            }
            Get-AzStorageBlobContent -Context $StorageContext -Container $ResultsContainerName -Blob $blob.Name -Destination $destination -Force | Out-Null
        }
        Write-RunLog "Applied $($customBlobs.Count) custom test file(s) from '$customPrefix'."
    }
}

function Get-SuppressionMap {
    param(
        [Parameter(Mandatory = $true)][object] $StorageContext,
        [Parameter(Mandatory = $true)][string] $TenantPrefix,
        [Parameter(Mandatory = $true)][string] $WorkingRoot
    )

    $suppressionBlob = "$TenantPrefix/config/suppressions.json"
    $suppressionPath = Join-Path -Path $WorkingRoot -ChildPath 'suppressions.json'
    try {
        Get-AzStorageBlobContent -Context $StorageContext -Container $ResultsContainerName -Blob $suppressionBlob -Destination $suppressionPath -Force -ErrorAction Stop | Out-Null
        $suppression = Get-Content -Path $suppressionPath -Raw | ConvertFrom-Json
        $map = @{}
        foreach ($entry in @($suppression.items)) {
            if (-not [string]::IsNullOrWhiteSpace([string] $entry.id)) {
                $map[[string] $entry.id] = $entry
            }
        }
        return $map
    } catch {
        return @{}
    }
}

function Apply-SuppressionsToResult {
    param(
        [Parameter(Mandatory = $true)][object] $Result,
        [Parameter(Mandatory = $true)][hashtable] $SuppressionMap
    )

    if (-not $SuppressionMap -or $SuppressionMap.Keys.Count -eq 0) {
        return 0
    }

    $suppressedCount = 0
    foreach ($test in @($Result.Tests)) {
        $testId = [string] $test.Id
        if ([string]::IsNullOrWhiteSpace($testId)) { continue }
        if (-not $SuppressionMap.ContainsKey($testId)) { continue }
        if ($test.Result -notin @('Failed', 'Error', 'Investigate')) { continue }

        $test.Result = 'Skipped'
        if (-not $test.SkippedReason) {
            $reason = [string] $SuppressionMap[$testId].reason
            if ([string]::IsNullOrWhiteSpace($reason)) { $reason = 'Suppressed by baseline policy' }
            $test | Add-Member -MemberType NoteProperty -Name SkippedReason -Value $reason -Force
        }
        $suppressedCount++
    }

    if ($suppressedCount -gt 0) {
        $Result.FailedCount = @($Result.Tests | Where-Object { $_.Result -eq 'Failed' }).Count
        $Result.ErrorCount = @($Result.Tests | Where-Object { $_.Result -eq 'Error' }).Count
        $Result.InvestigateCount = @($Result.Tests | Where-Object { $_.Result -eq 'Investigate' }).Count
        $Result.SkippedCount = @($Result.Tests | Where-Object { $_.Result -eq 'Skipped' }).Count
    }

    return $suppressedCount
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

Write-RunLog "Starting DriftMaester invoke run version $($script:DriftMaesterVersion)."
$workingRoot = Join-Path -Path ([System.IO.Path]::GetTempPath()) -ChildPath "MaesterDrift-$([Guid]::NewGuid().ToString('N'))"
New-Item -Path $workingRoot -ItemType Directory -Force | Out-Null
Write-RunLog "Working root: $workingRoot"

try {
    Write-RunLog "Detecting Azure Automation Account and Storage Account for Maester results."
    $parsedReportRecipients = @(Get-RecipientList -Value $ReportRecipient)
    if ($parsedReportRecipients.Count -eq 0) {
        throw 'ReportRecipient is required. Add one or more comma-separated email addresses as a runbook parameter.'
    }

    $global:VerbosePreference = 'SilentlyContinue'
    Import-RequiredModule -Name DLLPickle
    Import-DPLibrary
    Import-Module PSPublishModule -Force
    Import-RequiredModule -Name MicrosoftTeams
    Import-RequiredModule -Name Az.Accounts
    Import-RequiredModule -Name Az.Storage
    Import-IsolatedModule -Profile ExchangeOnlineManagement | Out-Null
    Import-RequiredModule -Name Microsoft.Graph.Authentication
    # Maester is imported here at script scope instead of through Import-RequiredModule on purpose. Its manifest
    # declares ScriptsToProcess = 'OrcaClasses.ps1', and ScriptsToProcess runs in the scope that called
    # Import-Module. Importing from inside a function therefore defines PolicyInfo, PolicyType and the other ORCA
    # classes in that function scope, and they are gone by the time the tests run, which makes every ORCA test fail
    # with "Cannot find type [PolicyInfo]". Keep this call at script scope (try/catch does not introduce a scope).
    Write-RunLog "Importing module 'Maester'."
    $Null = Import-Module Maester 4>&1 | Where-Object { $_ -isnot [System.Management.Automation.VerboseRecord] }
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

    Write-RunLog "Testing connections to optional services used by Maester for richer reporting and drift detection."

    $optionalConnectionWarnings = [System.Collections.Generic.List[object]]::new()
    foreach ($missing in $missingGraphPermissions) {
        $optionalConnectionWarnings.Add($missing)
        Write-RunLog "Graph permission warning: $($missing.Permission) is not available. Related Maester tests may be skipped or have less detail." -Level Warning
    }

    foreach ($missing in (Connect-OptionalMaesterServices)) {
        $optionalConnectionWarnings.Add($missing)
        Write-RunLog "Optional service warning: $($missing.Service) - $($missing.Reason)" -Level Warning
    }

    $storageContext = Resolve-StorageContext -Value (Get-StorageContextForResults)
    Write-RunLog "Resolved storage context type: $($storageContext.GetType().FullName)"
    $resolvedTenantId = if (-not [string]::IsNullOrWhiteSpace($script:ConnectedTenantId)) { $script:ConnectedTenantId } elseif (-not [string]::IsNullOrWhiteSpace($TenantId)) { $TenantId } else { 'unknown-tenant' }
    $tenantSafe = ([string] $resolvedTenantId) -replace '[^a-zA-Z0-9-]', '_'
    if ([string]::IsNullOrWhiteSpace($tenantSafe)) { $tenantSafe = 'unknown-tenant' }
    $tenantPrefix = (($BlobPrefix.Trim('/'), $tenantSafe) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }) -join '/'
    $resultPrefix = "$tenantPrefix/results"
    $reportPrefix = "$tenantPrefix/reports"
    $summaryPrefix = "$tenantPrefix/summary"
    $cacheBlob = "$tenantPrefix/cache/maester-tests.zip"

    $testsWorkingPath = Join-Path -Path $workingRoot -ChildPath 'maester-tests'
    New-Item -Path $testsWorkingPath -ItemType Directory -Force | Out-Null

    try {
        Write-RunLog "Updating Maester tests in working folder '$testsWorkingPath'."
        Update-MaesterTests -Path $testsWorkingPath -Force | Out-Null
        Save-TestsCache -StorageContext $storageContext -CacheBlobName $cacheBlob -TestsPath $testsWorkingPath
    } catch {
        Write-RunLog "Could not update Maester tests from GitHub: $($_.Exception.Message). Trying cached test set." -Level Warning
        if (-not (Try-RestoreTestsCache -StorageContext $storageContext -CacheBlobName $cacheBlob -TargetTestsPath $testsWorkingPath)) {
            throw 'Maester test update failed and no cached tests were available to continue safely.'
        }
    }

    Apply-TenantMaesterConfig -StorageContext $storageContext -TenantPrefix $tenantPrefix -TestsPath $testsWorkingPath

    $runOutputFolder = Join-Path -Path $workingRoot -ChildPath 'test-results'
    New-Item -Path $runOutputFolder -ItemType Directory -Force | Out-Null

    Write-RunLog "Invoking Maester run and generating results, this can take a while depending on tenant size."

    $timestamp = Get-Date -Format 'yyyyMMdd-HHmmss'
    $outputFileName = "TestResults-$timestamp"
    $maesterResult = Invoke-FullMaesterRun -TestsPath $testsWorkingPath -OutputFolder $runOutputFolder -OutputFileName $outputFileName

    if (-not $maesterResult) {
        throw 'Invoke-Maester did not return a result object.'
    } else {
        Write-RunLog "Done! Comparing with previous runs (if any)."
    }

    $suppressionMap = Get-SuppressionMap -StorageContext $storageContext -TenantPrefix $tenantPrefix -WorkingRoot $workingRoot
    $suppressedCount = Apply-SuppressionsToResult -Result $maesterResult -SuppressionMap $suppressionMap
    if ($suppressedCount -gt 0) {
        Write-RunLog "Applied suppression policy to $suppressedCount finding(s)."
    }

    Refresh-GraphConnection

    $jsonPath = Join-Path -Path $runOutputFolder -ChildPath "$outputFileName.json"
    $htmlPath = Join-Path -Path $runOutputFolder -ChildPath "$outputFileName.html"
    $markdownPath = Join-Path -Path $runOutputFolder -ChildPath "$outputFileName.md"

    foreach ($path in @($jsonPath, $htmlPath, $markdownPath)) {
        if (-not (Test-Path -Path $path -PathType Leaf)) {
            throw "Expected Maester output file was not created: $path"
        }
    }

    $tenantSafe = ([string] $maesterResult.TenantId) -replace '[^a-zA-Z0-9-]', '_'
    if ([string]::IsNullOrWhiteSpace($tenantSafe)) { $tenantSafe = 'unknown-tenant' }
    $tenantPrefix = (($BlobPrefix.Trim('/'), $tenantSafe) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }) -join '/'
    $resultPrefix = "$tenantPrefix/results"
    $reportPrefix = "$tenantPrefix/reports"
    $summaryPrefix = "$tenantPrefix/summary"

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
    $summaryBlob = "$summaryPrefix/summary-$timestamp.json"
    Save-BlobFile -StorageContext $storageContext -FilePath $jsonPath -BlobName $jsonBlob
    Save-BlobFile -StorageContext $storageContext -FilePath $htmlPath -BlobName $htmlBlob
    Save-BlobFile -StorageContext $storageContext -FilePath $markdownPath -BlobName $markdownBlob
    Save-RunSummaryBlob -StorageContext $storageContext -BlobName $summaryBlob -Result $maesterResult -ModuleVersion (Get-InstalledMaesterModuleVersion)

    $summaryBlobs = Get-HistoricalSummaryBlobs -StorageContext $storageContext -TenantSummaryPrefix $summaryPrefix
    $trendResults = [System.Collections.Generic.List[object]]::new()
    foreach ($blob in @($summaryBlobs | Select-Object -First $TrendRunCount)) {
        try {
            $trendResult = Read-SummaryBlob -StorageContext $storageContext -BlobName $blob.Name -DestinationFolder $workingRoot
            $trendResults.Add([PSCustomObject]@{
                ExecutedAt       = [datetime] $trendResult.ExecutedAt
                Label            = ([datetime] $trendResult.ExecutedAt).ToString('dd MMM HH:mm')
                Score            = [double] $trendResult.Score
                PassedCount      = [int] $trendResult.PassedCount
                FailedCount      = [int] $trendResult.FailedCount
                ErrorCount       = [int] $trendResult.ErrorCount
                InvestigateCount = [int] $trendResult.InvestigateCount
                SkippedCount     = [int] $trendResult.SkippedCount
                NotRunCount      = [int] $trendResult.NotRunCount
                TotalCount       = [int] $trendResult.TotalCount
            })
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
        Summary        = $summaryBlob
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
    Enforce-ResultRetention -StorageContext $storageContext -TenantPrefix $tenantPrefix -RetentionDays $RetentionDays

    $subjectPassedDelta = if ($previousResult) { [int] $maesterResult.PassedCount - [int] $previousResult.PassedCount } else { 0 }
    $subjectDeltaSuffix = if ($subjectPassedDelta -gt 0) { " (+$subjectPassedDelta)" } elseif ($subjectPassedDelta -lt 0) { " ($subjectPassedDelta)" } else { '' }
    $subject = "${MailSubjectPrefix}: $($maesterResult.TenantName) passed $($maesterResult.PassedCount)/$($maesterResult.TotalCount)$subjectDeltaSuffix, findings $($maesterResult.FailedCount + $maesterResult.ErrorCount + $maesterResult.InvestigateCount)"
    
    # Determine whether to send the report based on AlwaysSendReport flag, first run detection, and drift presence
    $isFirstRun = -not $previousResult
    $hasDiff = $diff.Summary.Regressed -gt 0 -or $diff.Summary.Improved -gt 0 -or $diff.Summary.Changed -gt 0 -or $diff.Summary.Added -gt 0 -or $diff.Summary.Removed -gt 0
    $meetsSeverityPolicy = Test-DiffMeetsSeverityThreshold -Diff $diff -MinimumSeverity $AlertMinimumSeverity
    $shouldSendReport = $AlwaysSendReport -or $isFirstRun -or ($hasDiff -and $meetsSeverityPolicy)
    $alertRecipients = @(Get-AlertRecipientList)
    $mailRecipientsToUse = if ($hasDiff -and $meetsSeverityPolicy -and $alertRecipients.Count -gt 0) { $alertRecipients } else { $parsedReportRecipients }
    $recipientSummary = $mailRecipientsToUse -join ', '
    
    if ($shouldSendReport) {
        $driftReportZipPath = Join-Path -Path $runOutputFolder -ChildPath "DriftReport-$timestamp.zip"
        Write-RunLog "Zipping the drift report '$([System.IO.Path]::GetFileName($driftReportPath))' for attachment to keep the mail small even for large tenants."
        Compress-Archive -Path $driftReportPath -DestinationPath $driftReportZipPath -Force
        $attachmentPaths = @()
        if ($ReportDelivery -ne 'link') {
            $attachmentPaths += $driftReportZipPath
        }
        if ($IncludeMaesterReport -and $ReportDelivery -ne 'link') {
            $maesterReportZipPath = Join-Path -Path $runOutputFolder -ChildPath "$outputFileName.zip"
            Write-RunLog "IncludeMaesterReport is enabled. Zipping the original Maester report '$([System.IO.Path]::GetFileName($htmlPath))' for attachment."
            Compress-Archive -Path $htmlPath -DestinationPath $maesterReportZipPath -Force
            $attachmentPaths += $maesterReportZipPath
        }

        $emailBodyToSend = $emailHtml
        if ($ReportDelivery -eq 'link') {
            $emailBodyToSend += (Get-ReportLinkHtml -StorageAccountName $script:DetectedStorageAccountName -ContainerName $ResultsContainerName -BlobName $driftBlob)
            $attachmentPaths = @()
        }

        $mailResult = Send-DriftMail -Subject $subject -HtmlBody $emailBodyToSend -AttachmentPath $attachmentPaths -RecipientsOverride $mailRecipientsToUse
        Send-TeamsNotification -Subject $subject -CurrentResult $maesterResult -PreviousResult $previousResult -Diff $diff -DriftBlobName $driftBlob -RecipientSummary $recipientSummary
        Write-RunLog "Maester drift detection completed. Report sent to $recipientSummary. Attachments included: $($mailResult.AttachedFiles -join ', ')"
    } else {
        Write-RunLog "Maester drift detection completed. No changes met alert policy and AlwaysSendReport is false, so no report was sent."
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

    Send-FailureNotification -ErrorMessage $_.Exception.Message -StackTrace $_.ScriptStackTrace

    throw
} finally {
    Flush-RunLog

    if ((Test-Path -Path $workingRoot)) {
        Remove-Item -Path $workingRoot -Recurse -Force -ErrorAction SilentlyContinue
    }
}
