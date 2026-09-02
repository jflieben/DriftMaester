<#
.SYNOPSIS
Checks for the DriftMaester PowerShell runtime environment in the current Azure Automation Account and enforces the required module versions.

.DESCRIPTION
This runbook connects with the Automation Account managed identity, resolves the current Automation Account context
from the job environment or by finding the Automation Account that has the current managed identity assigned,
checks for a runtime environment named driftmaester or a PowerShell 7.6 runtime environment, creates the driftmaester runtime environment
when no match exists, and installs or updates the required package list using the Azure Automation Runtime Environment REST API.

It also keeps the DriftMaester runbooks themselves current: it downloads the latest Invoke-DriftMaester and
Update-DriftMaester source from GitHub, compares it with the content published in the Automation Account, and
re-imports (and publishes) any runbook whose content differs or is missing. Updating the running Update-DriftMaester
runbook does not affect the currently executing job; the new version is used on the next run.

Fixed package versions are pinned in this script. A required package can use version descriptor latest to resolve the latest stable
non-preview version from PSGallery directly, without requiring PowerShellGet or the NuGet package provider.

.EXAMPLE
./Update-DriftMaester.ps1

.NOTES
Author: Jos Lieben / JSolve B.V.
https://jsolve.nl
Blog: https://www.lieben.nu/liebensraum/
Free for non-commercial use. Commercial use requires a license:
https://jsolve.nl/commercial-use.html
#>

$ErrorActionPreference = 'Stop'
$RuntimeVersion      = '7.6'
$PreferredRuntimeName = 'driftmaester'
$PollIntervalSeconds = 15
$TimeoutMinutes      = 20
$RunbookSyncTimeoutMinutes = 5
$WebRequestMaxAttempts = 5
$WebRequestRetryBaseSeconds = 5
$PackageImportMaxAttempts = 3
$PackageImportRetryBaseSeconds = 30
$script:AutomationApiVersion = '2024-10-23'
$script:ArmResourceUrl = 'https://management.azure.com/'
$script:ArmAccessTokenPayload = $null
$script:RunLogDeferDepth = 0
$script:DeferredRunLogs = [System.Collections.Generic.List[string]]::new()
$GithubRawBase = 'https://raw.githubusercontent.com/jflieben/DriftMaester/main/Runbooks'
$ManagedRunbooks = @(
    [PSCustomObject]@{ RunbookName = 'Invoke-DriftMaester'; SourceFileName = 'Invoke-DriftMaester.ps1' }
    [PSCustomObject]@{ RunbookName = 'Update-DriftMaester'; SourceFileName = 'Update-DriftMaester.ps1' }
)
$DefaultRuntimePackages = @{
	'az'          = '15.1.0'
	'azure cli' = '2.77.0'
}
$RequiredModules = @(
    [PSCustomObject]@{ PackageName = 'ADOPS'; Version = '2.4.2' }
    [PSCustomObject]@{ PackageName = 'Az.Accounts'; Version = '5.4.0' }
    [PSCustomObject]@{ PackageName = 'Az.Storage'; Version = '9.6.1' }
    [PSCustomObject]@{ PackageName = 'AzAuth'; Version = '2.9.0' }
    [PSCustomObject]@{ PackageName = 'DLLPickle'; Version = '1.3.1' }
    [PSCustomObject]@{ PackageName = 'ExchangeOnlineManagement'; Version = '3.9.2' }
    [PSCustomObject]@{ PackageName = 'Maester'; Version = '2.2.0' }
    [PSCustomObject]@{ PackageName = 'Microsoft.Graph.Authentication'; Version = '2.37.0' }
    [PSCustomObject]@{ PackageName = 'MicrosoftTeams'; Version = '7.7.0' }
    [PSCustomObject]@{ PackageName = 'Pester'; Version = '5.7.1' }
    [PSCustomObject]@{ PackageName = 'PnP.PowerShell'; Version = '3.3.0' }
    [PSCustomObject]@{ PackageName = 'PSPublishModule'; Version = '3.0.66' }
)

function Write-RunLog {
    param(
        [Parameter(Mandatory = $true)][string] $Message,
        [Parameter(Mandatory = $false)][ValidateSet('Info', 'Warning', 'Error', 'Success')][string] $Level = 'Info'
    )

    $prefix = "[{0:u}] [{1}]" -f (Get-Date), $Level.ToUpperInvariant()
    $logLine = "$prefix $Message"

    if ($script:RunLogDeferDepth -gt 0) {
        $script:DeferredRunLogs.Add($logLine)
        return
    }

    Write-Output $logLine
}

function Write-DeferredRunLog {
    foreach ($logLine in $script:DeferredRunLogs.ToArray()) {
        Write-Output $logLine
    }

    $script:DeferredRunLogs.Clear()
}

function Invoke-WithDeferredRunLog {
    param([Parameter(Mandatory = $true)][scriptblock] $ScriptBlock)

    $script:RunLogDeferDepth++
    try {
        [PSCustomObject]@{
            Output      = @(& $ScriptBlock)
            ErrorRecord = $null
        }
    } catch {
        [PSCustomObject]@{
            Output      = @()
            ErrorRecord = $_
        }
    } finally {
        $script:RunLogDeferDepth--
    }
}

function Invoke-WebRequestWithRetry {
    param(
        [Parameter(Mandatory = $true)][string] $Uri,
        [Parameter(Mandatory = $false)][hashtable] $Headers,
        [Parameter(Mandatory = $false)][int] $MaximumRedirection = -1
    )

    $requestParams = @{
        Uri             = $Uri
        UseBasicParsing = $true
        ErrorAction     = 'Stop'
    }

    if ($Headers) { $requestParams['Headers'] = $Headers }
    if ($MaximumRedirection -ge 0) { $requestParams['MaximumRedirection'] = $MaximumRedirection }

    for ($attempt = 1; $attempt -le $WebRequestMaxAttempts; $attempt++) {
        try {
            return Invoke-WebRequest @requestParams
        } catch {
            $statusCode = $null
            if ($_.Exception.Response) {
                try { $statusCode = [int] $_.Exception.Response.StatusCode } catch { $statusCode = $null }
            }

            if ($MaximumRedirection -eq 0 -and $statusCode -in @(301, 302, 303, 307, 308)) {
                throw
            }

            $isTransient = $statusCode -eq 408 -or $statusCode -eq 429 -or ($statusCode -ge 500 -and $statusCode -le 599)
            if (-not $isTransient -or $attempt -ge $WebRequestMaxAttempts) {
                throw
            }

            $delaySeconds = [Math]::Min(60, [int]($WebRequestRetryBaseSeconds * [Math]::Pow(2, ($attempt - 1))))
            $retryAfterValues = $null
            try {
                if ($_.Exception.Response.Headers -and $_.Exception.Response.Headers.TryGetValues('Retry-After', [ref] $retryAfterValues)) {
                    $retryAfterValue = @($retryAfterValues | Select-Object -First 1)[0]
                    $retryAfterSeconds = 0
                    if ([int]::TryParse([string] $retryAfterValue, [ref] $retryAfterSeconds) -and $retryAfterSeconds -gt 0) {
                        $delaySeconds = [Math]::Min(120, $retryAfterSeconds)
                    }
                }
            } catch {
                $delaySeconds = $delaySeconds
            }

            Write-RunLog "Web request to '$Uri' failed with HTTP $statusCode. Retrying attempt $($attempt + 1)/$WebRequestMaxAttempts in $delaySeconds second(s)." -Level Warning
            Start-Sleep -Seconds $delaySeconds
        }
    }
}

function Connect-RunbookIdentity {
    Write-RunLog "Connecting to Azure with managed identity."
    Connect-AzAccount -Identity -SkipContextPopulation | Out-Null
}

function Get-ArmAccessTokenPlainText {
    $tokenResponse = Get-AzAccessToken -ResourceUrl $script:ArmResourceUrl -AsSecureString
    [System.Net.NetworkCredential]::new('', $tokenResponse.Token).Password
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

function Get-CurrentManagedIdentityFromArmToken {
    if (-not $script:ArmAccessTokenPayload) {
        $script:ArmAccessTokenPayload = Get-JwtPayload -Token (Get-ArmAccessTokenPlainText)
    }

    if (-not $script:ArmAccessTokenPayload) {
        throw 'Could not read the managed identity claims from the ARM access token.'
    }

    [PSCustomObject]@{
        ClientId    = if ($script:ArmAccessTokenPayload.appid) { $script:ArmAccessTokenPayload.appid } else { $script:ArmAccessTokenPayload.azp }
        PrincipalId = $script:ArmAccessTokenPayload.oid
        TenantId    = $script:ArmAccessTokenPayload.tid
    }
}

function Get-CurrentAutomationAccountFromEnvironment {
    if ([string]::IsNullOrWhiteSpace($env:AUTOMATION_ACCOUNT_ID)) {
        return $null
    }

    $resourceId = $env:AUTOMATION_ACCOUNT_ID
    if ($resourceId -notmatch '^/subscriptions/([^/]+)/resourceGroups/([^/]+)/providers/Microsoft\.Automation/automationAccounts/([^/]+)$') {
        throw "AUTOMATION_ACCOUNT_ID has an unexpected format: '$resourceId'."
    }

    [PSCustomObject]@{
        SubscriptionId       = $Matches[1]
        ResourceGroupName    = $Matches[2]
        AutomationAccountName = $Matches[3]
    }
}

function Get-ResourceGroupNameFromResourceId {
    param([Parameter(Mandatory = $true)][string] $ResourceId)

    if ($ResourceId -match '/resourceGroups/([^/]+)/') {
        return $Matches[1]
    }

    throw "Could not parse resource group from resource id '$ResourceId'."
}

function Find-CurrentAutomationAccount {
    $identity = Get-CurrentManagedIdentityFromArmToken
    if ([string]::IsNullOrWhiteSpace($identity.PrincipalId) -and [string]::IsNullOrWhiteSpace($identity.ClientId)) {
        throw 'Could not resolve the current managed identity principal id or client id from the ARM token.'
    }

    Write-RunLog "AUTOMATION_ACCOUNT_ID was not available. Searching visible subscriptions for an Automation Account assigned to the current managed identity."

    $subscriptionsUri = "https://management.azure.com/subscriptions?api-version=2020-01-01"
    $subscriptions = @(Get-AllPagesFromArmList -InitialUri $subscriptionsUri)

    foreach ($subscription in $subscriptions) {
        $subscriptionId = $subscription.subscriptionId
        if ([string]::IsNullOrWhiteSpace($subscriptionId)) { continue }

        Write-RunLog "Checking Automation Accounts in subscription '$subscriptionId'."
        $accountsUri = "https://management.azure.com/subscriptions/$subscriptionId/providers/Microsoft.Automation/automationAccounts?api-version=2023-11-01"

        try {
            $automationAccounts = @(Get-AllPagesFromArmList -InitialUri $accountsUri)
        } catch {
            Write-RunLog "Skipping subscription '$subscriptionId': $_" -Level Warning
            continue
        }

        foreach ($account in $automationAccounts) {
            $accountIdentity = $account.identity
            if (-not $accountIdentity) { continue }

            $matchesIdentity = $false
            if (-not [string]::IsNullOrWhiteSpace($identity.PrincipalId) -and $accountIdentity.principalId -eq $identity.PrincipalId) {
                $matchesIdentity = $true
            }

            if (-not $matchesIdentity -and $accountIdentity.userAssignedIdentities) {
                foreach ($userAssignedIdentity in $accountIdentity.userAssignedIdentities.PSObject.Properties) {
                    if ((-not [string]::IsNullOrWhiteSpace($identity.ClientId) -and $userAssignedIdentity.Value.clientId -eq $identity.ClientId) -or
                        (-not [string]::IsNullOrWhiteSpace($identity.PrincipalId) -and $userAssignedIdentity.Value.principalId -eq $identity.PrincipalId)) {
                        $matchesIdentity = $true
                        break
                    }
                }
            }

            if ($matchesIdentity) {
                return [PSCustomObject]@{
                    SubscriptionId        = $subscriptionId
                    ResourceGroupName     = Get-ResourceGroupNameFromResourceId -ResourceId $account.id
                    AutomationAccountName = $account.name
                }
            }
        }
    }

    throw 'Could not find an Azure Automation Account in visible subscriptions with the current managed identity assigned. Grant the identity Reader on the Automation Account or resource group so it can detect itself.'
}

function Resolve-AutomationAccountContext {
    $detected = Get-CurrentAutomationAccountFromEnvironment

    if (-not $detected) {
        $detected = Find-CurrentAutomationAccount
    }

    $SubscriptionId       = if ($detected) { $detected.SubscriptionId } else { $null }
    $ResourceGroupName    = if ($detected) { $detected.ResourceGroupName } else { $null }
    $AutomationAccountName = if ($detected) { $detected.AutomationAccountName } else { $null }

    if (-not $SubscriptionId) {
        $ctx = Get-AzContext
        if ($ctx -and $ctx.Subscription -and $ctx.Subscription.Id) {
            $SubscriptionId = $ctx.Subscription.Id
        }
    }

    if ([string]::IsNullOrWhiteSpace($SubscriptionId) -or [string]::IsNullOrWhiteSpace($ResourceGroupName) -or [string]::IsNullOrWhiteSpace($AutomationAccountName)) {
        throw 'Unable to resolve Automation Account context from AUTOMATION_ACCOUNT_ID or managed identity assignment.'
    }

    [PSCustomObject]@{
        SubscriptionId        = $SubscriptionId
        ResourceGroupName     = $ResourceGroupName
        AutomationAccountName = $AutomationAccountName
    }
}

function Invoke-ArmRequest {
    param(
        [Parameter(Mandatory = $true)][ValidateSet('GET', 'PUT')][string] $Method,
        [Parameter(Mandatory = $true)][string] $Uri,
        [Parameter(Mandatory = $false)][object] $Body
    )

    $accessToken = Get-ArmAccessTokenPlainText

    $requestParams = @{
        Method      = $Method
        Uri         = $Uri
        Headers     = @{ Authorization = "Bearer $accessToken" }
        ErrorAction = 'Stop'
    }

    if ($PSBoundParameters.ContainsKey('Body')) {
        $requestParams['Body'] = ($Body | ConvertTo-Json -Depth 20)
        $requestParams['ContentType'] = 'application/json'
    }

    try {
        Invoke-RestMethod @requestParams
    } catch {
        $details = $null
        if ($_.ErrorDetails -and -not [string]::IsNullOrWhiteSpace($_.ErrorDetails.Message)) {
            $details = $_.ErrorDetails.Message
        } elseif ($_.Exception.Response) {
            try {
                $stream = $_.Exception.Response.GetResponseStream()
                if ($stream) {
                    $reader = [System.IO.StreamReader]::new($stream)
                    $details = $reader.ReadToEnd()
                }
            } catch {
                $details = $null
            }
        }

        if ([string]::IsNullOrWhiteSpace($details)) {
            throw
        }

        throw "ARM $Method request failed for '$Uri'. $($_.Exception.Message) Details: $details"
    }
}

function Get-AllPagesFromArmList {
    param([Parameter(Mandatory = $true)][string] $InitialUri)

    $items = New-Object System.Collections.Generic.List[object]
    $nextLink = $InitialUri

    while (-not [string]::IsNullOrWhiteSpace($nextLink)) {
        $response = Invoke-ArmRequest -Method GET -Uri $nextLink
        if ($response.value) {
            foreach ($item in $response.value) {
                $items.Add($item)
            }
        }

        if ($response.nextLink) {
            $nextLink = [string] $response.nextLink
        } else {
            $nextLink = $null
        }
    }

    return $items
}

function Get-LatestStablePackageVersion {
    param([Parameter(Mandatory = $true)][string] $PackageName)

    Write-RunLog "Checking PSGallery for the latest stable $PackageName version."

    $escapedPackageName = [System.Uri]::EscapeDataString($PackageName.Replace("'", "''"))
    $galleryUri = "https://www.powershellgallery.com/api/v2/FindPackagesById()?id='$escapedPackageName'"
    $nextUri = $galleryUri
    $versions = [System.Collections.Generic.List[object]]::new()

    while (-not [string]::IsNullOrWhiteSpace($nextUri)) {
        $response = Invoke-WebRequestWithRetry -Uri $nextUri -Headers @{ Accept = 'application/atom+xml' }
        [xml] $feed = $response.Content

        $namespaceManager = [System.Xml.XmlNamespaceManager]::new($feed.NameTable)
        $namespaceManager.AddNamespace('atom', 'http://www.w3.org/2005/Atom')
        $namespaceManager.AddNamespace('m', 'http://schemas.microsoft.com/ado/2007/08/dataservices/metadata')
        $namespaceManager.AddNamespace('d', 'http://schemas.microsoft.com/ado/2007/08/dataservices')

        foreach ($entry in $feed.SelectNodes('//atom:entry', $namespaceManager)) {
            $properties = $entry.SelectSingleNode('m:properties', $namespaceManager)
            if (-not $properties) { continue }

            $id = $properties.SelectSingleNode('d:Id', $namespaceManager).InnerText
            $versionText = $properties.SelectSingleNode('d:Version', $namespaceManager).InnerText
            $isPrereleaseNode = $properties.SelectSingleNode('d:IsPrerelease', $namespaceManager)
            $isPrerelease = $isPrereleaseNode -and [System.Convert]::ToBoolean($isPrereleaseNode.InnerText)

            if ($id -ine $PackageName -or $isPrerelease) { continue }

            try {
                $version = [version] $versionText
            } catch {
                Write-RunLog "Skipping $PackageName version '$versionText' because it is not a parseable stable version." -Level Warning
                continue
            }

            $versions.Add([PSCustomObject]@{
                    VersionText = $versionText
                    Version     = $version
                })
        }

        $nextLink = $feed.SelectSingleNode('//atom:link[@rel="next"]', $namespaceManager)
        $nextUri = if ($nextLink) { [string] $nextLink.href } else { $null }
    }

    $stable = @($versions.ToArray() | Sort-Object Version -Descending | Select-Object -First 1)

    if (-not $stable) {
        throw "No stable $PackageName version was found in PSGallery."
    }

    [PSCustomObject]@{
        Name    = $PackageName
        Version = $stable.VersionText
        Uri     = "https://www.powershellgallery.com/api/v2/package/$PackageName/$($stable.VersionText)"
    }
}

function Get-PSGalleryPackageContentUri {
    param(
        [Parameter(Mandatory = $true)][string] $PackageName,
        [Parameter(Mandatory = $true)][string] $Version
    )

    $packageUri = "https://www.powershellgallery.com/api/v2/package/$PackageName/$Version"

    try {
        $response = Invoke-WebRequestWithRetry -Uri $packageUri -MaximumRedirection 0
        if ($response.BaseResponse -and $response.BaseResponse.ResponseUri) {
            return [string] $response.BaseResponse.ResponseUri.AbsoluteUri
        }

        return $packageUri
    } catch {
        if ($_.Exception.Response -and [int] $_.Exception.Response.StatusCode -in @(301, 302, 303, 307, 308)) {
            $location = $_.Exception.Response.Headers.Location
            if (-not [string]::IsNullOrWhiteSpace([string] $location)) {
                return [string] $location
            }
        }

        throw "Could not resolve PSGallery package content URI for $PackageName $Version. $($_.Exception.Message)"
    }
}

function Resolve-RequiredPackageVersion {
    param(
        [Parameter(Mandatory = $true)][string] $PackageName,
        [Parameter(Mandatory = $true)][string] $VersionDescriptor
    )

    if ($VersionDescriptor -ieq 'latest') {
        $latest = Get-LatestStablePackageVersion -PackageName $PackageName
        $latest.Uri = Get-PSGalleryPackageContentUri -PackageName $PackageName -Version $latest.Version
        return $latest
    }

    if ([string]::IsNullOrWhiteSpace($VersionDescriptor)) {
        throw "Required package '$PackageName' has no version descriptor. Use a fixed version or 'latest'."
    }

    [PSCustomObject]@{
        Name    = $PackageName
        Version = $VersionDescriptor
        Uri     = Get-PSGalleryPackageContentUri -PackageName $PackageName -Version $VersionDescriptor
    }
}

function Get-MinimumVersionFromNuGetRange {
    param([Parameter(Mandatory = $true)][string] $Range)

    if ([string]::IsNullOrWhiteSpace($Range)) {
        return $null
    }

    $trimmed = $Range.Trim()
    if ($trimmed -match '^[\[\(]\s*([^,\]\)\s]+)') {
        return $Matches[1]
    }

    if ($trimmed -match '^\d+(\.\d+){0,3}([\-\+][0-9A-Za-z\.-]+)?$') {
        return $trimmed
    }

    return $null
}

function Get-NuGetDependencyHints {
    param(
        [Parameter(Mandatory = $true)][string] $PackageName,
        [Parameter(Mandatory = $true)][string] $PackageVersion,
        [Parameter(Mandatory = $true)][string] $PackageUri
    )

    $tempFile = Join-Path -Path ([System.IO.Path]::GetTempPath()) -ChildPath ("$PackageName-$PackageVersion-$([Guid]::NewGuid().ToString('N')).nupkg")
    try {
        Write-RunLog "Downloading package metadata for compatibility hints: $PackageName $PackageVersion."
        $response = Invoke-WebRequestWithRetry -Uri $PackageUri
        [System.IO.File]::WriteAllBytes($tempFile, $response.Content)

        Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction SilentlyContinue
        $archive = [System.IO.Compression.ZipFile]::OpenRead($tempFile)
        try {
            $nuspecEntry = $archive.Entries | Where-Object { $_.FullName -like '*.nuspec' } | Select-Object -First 1
            if (-not $nuspecEntry) {
                Write-RunLog "No nuspec entry found in package '$PackageName'. Compatibility hints were skipped." -Level Warning
                return @{}
            }

            $reader = [System.IO.StreamReader]::new($nuspecEntry.Open())
            try {
                [xml] $nuspecXml = $reader.ReadToEnd()
            } finally {
                $reader.Dispose()
            }

            $dependencyNodes = @($nuspecXml.SelectNodes('//*[local-name()="dependencies"]//*[local-name()="dependency"]'))
            if (-not $dependencyNodes -or $dependencyNodes.Count -eq 0) {
                return @{}
            }

            $hints = @{}
            foreach ($dep in $dependencyNodes) {
                $depName = [string] $dep.id
                $versionRange = [string] $dep.version
                if ([string]::IsNullOrWhiteSpace($depName)) { continue }

                $minimumVersion = Get-MinimumVersionFromNuGetRange -Range $versionRange
                if ([string]::IsNullOrWhiteSpace($minimumVersion)) { continue }

                $hints[$depName] = $minimumVersion
            }

            return $hints
        } finally {
            $archive.Dispose()
        }
    } catch {
        Write-RunLog "Could not parse dependency metadata for '$PackageName' $PackageVersion. Continuing with pinned versions. Error: $($_.Exception.Message)" -Level Warning
        return @{}
    } finally {
        if (Test-Path -Path $tempFile) {
            Remove-Item -Path $tempFile -Force -ErrorAction SilentlyContinue
        }
    }
}

function Get-EffectiveRequiredModules {
    param([Parameter(Mandatory = $true)][object[]] $BaseRequiredModules)

    $effective = [System.Collections.Generic.List[object]]::new()
    foreach ($module in $BaseRequiredModules) {
        $effective.Add([PSCustomObject]@{
                PackageName = [string] $module.PackageName
                Version     = [string] $module.Version
            })
    }

    $maesterRequirement = $effective | Where-Object { $_.PackageName -ieq 'Maester' } | Select-Object -First 1
    if (-not $maesterRequirement) {
        return $effective.ToArray()
    }

    if ([string]::IsNullOrWhiteSpace([string] $maesterRequirement.Version)) {
        return $effective.ToArray()
    }

    $resolvedMaester = Resolve-RequiredPackageVersion -PackageName 'Maester' -VersionDescriptor ([string] $maesterRequirement.Version)
    $maesterRequirement.Version = [string] $resolvedMaester.Version
    Write-RunLog "Resolved Maester target version '$($resolvedMaester.Version)'. Evaluating dependency compatibility."

    $dependencyHints = Get-NuGetDependencyHints -PackageName 'Maester' -PackageVersion ([string] $resolvedMaester.Version) -PackageUri ([string] $resolvedMaester.Uri)
    if (-not $dependencyHints.Keys -or $dependencyHints.Keys.Count -eq 0) {
        return $effective.ToArray()
    }

    foreach ($module in $effective) {
        $packageName = [string] $module.PackageName
        if ([string]::IsNullOrWhiteSpace($packageName)) { continue }
        if ($packageName -ieq 'Maester') { continue }
        if (-not $dependencyHints.ContainsKey($packageName)) { continue }

        $minimumVersion = [string] $dependencyHints[$packageName]
        $currentDescriptor = [string] $module.Version
        if ([string]::IsNullOrWhiteSpace($minimumVersion) -or [string]::IsNullOrWhiteSpace($currentDescriptor)) { continue }
        if ($currentDescriptor -ieq 'latest') { continue }

        try {
            $currentVersion = [version] $currentDescriptor
            $requiredVersion = [version] $minimumVersion
            if ($currentVersion -lt $requiredVersion) {
                Write-RunLog "Auto-bumping '$packageName' from pinned version '$currentDescriptor' to '$minimumVersion' to stay compatible with Maester $($resolvedMaester.Version)." -Level Warning
                $module.Version = $minimumVersion
            }
        } catch {
            Write-RunLog "Could not compare version for dependency '$packageName' (current='$currentDescriptor', required='$minimumVersion'). Keeping current descriptor." -Level Warning
        }
    }

    return $effective.ToArray()
}

function Get-RuntimeEnvironments {
    param([Parameter(Mandatory = $true)][pscustomobject] $AutomationContext)

    $uri = "https://management.azure.com/subscriptions/$($AutomationContext.SubscriptionId)/resourceGroups/$($AutomationContext.ResourceGroupName)/providers/Microsoft.Automation/automationAccounts/$($AutomationContext.AutomationAccountName)/runtimeEnvironments?api-version=$($script:AutomationApiVersion)"
    Get-AllPagesFromArmList -InitialUri $uri
}

function Get-AutomationAccount {
    param([Parameter(Mandatory = $true)][pscustomobject] $AutomationContext)

    $uri = "https://management.azure.com/subscriptions/$($AutomationContext.SubscriptionId)/resourceGroups/$($AutomationContext.ResourceGroupName)/providers/Microsoft.Automation/automationAccounts/$($AutomationContext.AutomationAccountName)?api-version=2023-11-01"
    Invoke-ArmRequest -Method GET -Uri $uri
}

function Get-RuntimeEnvironment {
    param(
        [Parameter(Mandatory = $true)][pscustomobject] $AutomationContext,
        [Parameter(Mandatory = $true)][string] $RuntimeEnvironmentName
    )

    $escapedRuntime = [System.Uri]::EscapeDataString($RuntimeEnvironmentName)
    $uri = "https://management.azure.com/subscriptions/$($AutomationContext.SubscriptionId)/resourceGroups/$($AutomationContext.ResourceGroupName)/providers/Microsoft.Automation/automationAccounts/$($AutomationContext.AutomationAccountName)/runtimeEnvironments/$($escapedRuntime)?api-version=$($script:AutomationApiVersion)"
    Invoke-ArmRequest -Method GET -Uri $uri
}

function New-RuntimeEnvironment {
    param(
        [Parameter(Mandatory = $true)][pscustomobject] $AutomationContext,
        [Parameter(Mandatory = $true)][string] $RuntimeEnvironmentName,
        [Parameter(Mandatory = $true)][string] $RuntimeVersion
    )

    $automationAccount = Get-AutomationAccount -AutomationContext $AutomationContext
    $location = [string] $automationAccount.location
    if ([string]::IsNullOrWhiteSpace($location)) {
        throw "Could not determine location for Automation Account '$($AutomationContext.AutomationAccountName)'."
    }

    Write-RunLog "Creating runtime environment '$RuntimeEnvironmentName' with PowerShell $RuntimeVersion in location '$location'."

    $escapedRuntime = [System.Uri]::EscapeDataString($RuntimeEnvironmentName)
    $uri = "https://management.azure.com/subscriptions/$($AutomationContext.SubscriptionId)/resourceGroups/$($AutomationContext.ResourceGroupName)/providers/Microsoft.Automation/automationAccounts/$($AutomationContext.AutomationAccountName)/runtimeEnvironments/$($escapedRuntime)?api-version=$($script:AutomationApiVersion)"
    $body = @{
        location   = $location
        properties = @{
            defaultPackages = $DefaultRuntimePackages
            runtime = @{
                language = 'PowerShell'
                version  = $RuntimeVersion
            }
            description = 'Runtime environment for DriftMaester automation.'
        }
    }

    Invoke-ArmRequest -Method PUT -Uri $uri -Body $body
}

function Wait-ForRuntimeEnvironmentProvisioning {
    param(
        [Parameter(Mandatory = $true)][pscustomobject] $AutomationContext,
        [Parameter(Mandatory = $true)][string] $RuntimeEnvironmentName,
        [Parameter(Mandatory = $true)][int] $TimeoutMinutes,
        [Parameter(Mandatory = $true)][int] $PollIntervalSeconds
    )

    $deadline = (Get-Date).AddMinutes($TimeoutMinutes)

    while ((Get-Date) -lt $deadline) {
        $runtimeEnvironment = Get-RuntimeEnvironment -AutomationContext $AutomationContext -RuntimeEnvironmentName $RuntimeEnvironmentName
        $state = [string] $runtimeEnvironment.properties.provisioningState
        $version = [string] $runtimeEnvironment.properties.runtime.version
        Write-RunLog "Runtime environment '$RuntimeEnvironmentName' provisioning state: $state (PowerShell $version)."

        if ($state -eq 'Succeeded' -or $state -eq 'Created') {
            return $runtimeEnvironment
        }

        if ($state -in @('Failed', 'Canceled')) {
            $errorMessage = $runtimeEnvironment.properties.error.message
            if ([string]::IsNullOrWhiteSpace($errorMessage)) { $errorMessage = 'No runtime environment error details were returned.' }
            throw "Runtime environment creation failed with state '$state'. $errorMessage"
        }

        Start-Sleep -Seconds $PollIntervalSeconds
    }

    throw "Timed out waiting for runtime environment '$RuntimeEnvironmentName' to finish provisioning after $TimeoutMinutes minute(s)."
}

function Resolve-TargetRuntimeEnvironment {
    param([Parameter(Mandatory = $true)][pscustomobject] $AutomationContext)

    try {
        $runtime = Get-RuntimeEnvironment -AutomationContext $AutomationContext -RuntimeEnvironmentName $PreferredRuntimeName
        $runtimeLanguage = [string] $runtime.properties.runtime.language
        $runtimeVersion = [string] $runtime.properties.runtime.version
        if ($runtimeLanguage -ne 'PowerShell' -or $runtimeVersion -notmatch "^$([regex]::Escape($RuntimeVersion))(\.|$)") {
            Write-RunLog "Runtime environment '$PreferredRuntimeName' exists but is configured as '$runtimeLanguage $runtimeVersion'. DriftMaester expects PowerShell $RuntimeVersion and will attempt package updates anyway." -Level Warning
        }

        return $runtime
    } catch {
        $statusCode = $null
        if ($_.Exception.Response) {
            try { $statusCode = [int] $_.Exception.Response.StatusCode } catch { $statusCode = $null }
        }

        if ($statusCode -ne 404) {
            throw
        }
    }

    Write-RunLog "No runtime environment named '$PreferredRuntimeName' was found. Creating '$PreferredRuntimeName'." -Level Warning
    $null = New-RuntimeEnvironment -AutomationContext $AutomationContext -RuntimeEnvironmentName $PreferredRuntimeName -RuntimeVersion $RuntimeVersion
    Wait-ForRuntimeEnvironmentProvisioning -AutomationContext $AutomationContext -RuntimeEnvironmentName $PreferredRuntimeName -TimeoutMinutes $TimeoutMinutes -PollIntervalSeconds $PollIntervalSeconds
}

function Get-RuntimeEnvironmentPackages {
    param(
        [Parameter(Mandatory = $true)][pscustomobject] $AutomationContext,
        [Parameter(Mandatory = $true)][string] $RuntimeEnvironmentName
    )

    $escapedRuntime = [System.Uri]::EscapeDataString($RuntimeEnvironmentName)
    $uri = "https://management.azure.com/subscriptions/$($AutomationContext.SubscriptionId)/resourceGroups/$($AutomationContext.ResourceGroupName)/providers/Microsoft.Automation/automationAccounts/$($AutomationContext.AutomationAccountName)/runtimeEnvironments/$escapedRuntime/packages?api-version=$($script:AutomationApiVersion)"
    Get-AllPagesFromArmList -InitialUri $uri
}

function Get-PackageState {
    param(
        [Parameter(Mandatory = $true)][pscustomobject] $AutomationContext,
        [Parameter(Mandatory = $true)][string] $RuntimeEnvironmentName,
        [Parameter(Mandatory = $true)][string] $PackageName
    )

    $escapedRuntime = [System.Uri]::EscapeDataString($RuntimeEnvironmentName)
    $escapedPackage = [System.Uri]::EscapeDataString($PackageName)
    $uri = "https://management.azure.com/subscriptions/$($AutomationContext.SubscriptionId)/resourceGroups/$($AutomationContext.ResourceGroupName)/providers/Microsoft.Automation/automationAccounts/$($AutomationContext.AutomationAccountName)/runtimeEnvironments/$escapedRuntime/packages/$($escapedPackage)?api-version=$($script:AutomationApiVersion)"

    Invoke-ArmRequest -Method GET -Uri $uri
}

function Wait-ForPackageProvisioning {
    param(
        [Parameter(Mandatory = $true)][pscustomobject] $AutomationContext,
        [Parameter(Mandatory = $true)][string] $RuntimeEnvironmentName,
        [Parameter(Mandatory = $true)][string] $PackageName,
        [Parameter(Mandatory = $true)][int] $TimeoutMinutes,
        [Parameter(Mandatory = $true)][int] $PollIntervalSeconds
    )

    $deadline = (Get-Date).AddMinutes($TimeoutMinutes)

    while ((Get-Date) -lt $deadline) {
        $package = Get-PackageState -AutomationContext $AutomationContext -RuntimeEnvironmentName $RuntimeEnvironmentName -PackageName $PackageName
        $state = [string] $package.properties.provisioningState
        $version = [string] $package.properties.version
        Write-RunLog "Package '$PackageName' provisioning state: $state (version: $version)."

        if ($state -eq 'Succeeded' -or $state -eq 'Created') {
            return $package
        }

        if ($state -in @('Failed', 'Canceled')) {
            $errorMessage = $package.properties.error.message
            if ([string]::IsNullOrWhiteSpace($errorMessage)) { $errorMessage = 'No package error details were returned.' }
            throw "Package update failed with state '$state'. $errorMessage"
        }

        Start-Sleep -Seconds $PollIntervalSeconds
    }

    throw "Timed out waiting for package '$PackageName' to finish provisioning after $TimeoutMinutes minute(s)."
}

function Test-TransientPackageImportFailure {
    param([AllowNull()][string] $Message)

    if ([string]::IsNullOrWhiteSpace($Message)) {
        return $false
    }

    return $Message -match 'Internal Server Error|Response status code does not indicate success:\s*5\d\d|status code.*5\d\d|internal error occurred during module import|UploadDependencyContentAsync|SetModuleContent|temporar|transient|timeout|timed out'
}

function Update-PackageInRuntimeEnvironment {
    param(
        [Parameter(Mandatory = $true)][pscustomobject] $AutomationContext,
        [Parameter(Mandatory = $true)][pscustomobject] $RuntimeEnvironment,
        [Parameter(Mandatory = $true)][pscustomobject] $RequiredPackage,
        [Parameter(Mandatory = $false)][AllowNull()][pscustomobject] $ExistingPackage,
        [Parameter(Mandatory = $true)][pscustomobject] $TargetPackage
    )

    $runtimeName = [string] $RuntimeEnvironment.name
    $packageName = [string] $RequiredPackage.PackageName
    $existingVersion = if ($ExistingPackage) { [string] $ExistingPackage.properties.version } else { $null }
    $existingState = if ($ExistingPackage) { [string] $ExistingPackage.properties.provisioningState } else { 'NotInstalled' }

    if ($existingVersion -eq $TargetPackage.Version -and $existingState -in @('Succeeded', 'Created')) {
        Write-RunLog "$packageName is already up to date in runtime '$runtimeName' at version $existingVersion." -Level Success
        return [PSCustomObject]@{
            RuntimeEnvironment = $runtimeName
            PackageName        = $packageName
            PreviousVersion    = $existingVersion
            TargetVersion      = $TargetPackage.Version
            Changed            = $false
            State              = $existingState
        }
    }

    Write-RunLog "Installing/updating $packageName in runtime '$runtimeName' from '$existingVersion' to '$($TargetPackage.Version)'."

    $escapedRuntime = [System.Uri]::EscapeDataString($runtimeName)
    $escapedPackage = [System.Uri]::EscapeDataString($packageName)
    $uri = "https://management.azure.com/subscriptions/$($AutomationContext.SubscriptionId)/resourceGroups/$($AutomationContext.ResourceGroupName)/providers/Microsoft.Automation/automationAccounts/$($AutomationContext.AutomationAccountName)/runtimeEnvironments/$escapedRuntime/packages/$($escapedPackage)?api-version=$($script:AutomationApiVersion)"

    $body = @{
        location   = $RuntimeEnvironment.location
        properties = @{
            contentLink = @{
                uri     = $TargetPackage.Uri
                version = $TargetPackage.Version
            }
        }
    }

    $package = $null
    for ($attempt = 1; $attempt -le $PackageImportMaxAttempts; $attempt++) {
        try {
            if ($attempt -gt 1) {
                Write-RunLog "Retrying import for package '$packageName' in runtime '$runtimeName' (attempt $attempt/$PackageImportMaxAttempts)." -Level Warning
            }

            $null = Invoke-ArmRequest -Method PUT -Uri $uri -Body $body
            $package = Wait-ForPackageProvisioning -AutomationContext $AutomationContext -RuntimeEnvironmentName $runtimeName -PackageName $packageName -TimeoutMinutes $TimeoutMinutes -PollIntervalSeconds $PollIntervalSeconds
            break
        } catch {
            $message = $_.Exception.Message
            $isTransient = Test-TransientPackageImportFailure -Message $message
            if (-not $isTransient -or $attempt -ge $PackageImportMaxAttempts) {
                throw
            }

            $delaySeconds = [Math]::Min(180, [int]($PackageImportRetryBaseSeconds * [Math]::Pow(2, ($attempt - 1))))
            Write-RunLog "Package '$packageName' import failed with a transient Azure Automation error. Retrying attempt $($attempt + 1)/$PackageImportMaxAttempts in $delaySeconds second(s). Error: $message" -Level Warning
            Start-Sleep -Seconds $delaySeconds
        }
    }

    if (-not $package) {
        throw "Package '$packageName' did not return a provisioning result after $PackageImportMaxAttempts attempt(s)."
    }

    [PSCustomObject]@{
        RuntimeEnvironment = $runtimeName
        PackageName        = $packageName
        PreviousVersion    = $existingVersion
        TargetVersion      = $TargetPackage.Version
        Changed            = $true
        State              = [string] $package.properties.provisioningState
        ActualVersion      = [string] $package.properties.version
    }
}

function Update-RequiredPackagesInRuntimeEnvironment {
    param(
        [Parameter(Mandatory = $true)][pscustomobject] $AutomationContext,
        [Parameter(Mandatory = $true)][pscustomobject] $RuntimeEnvironment
    )

    $runtimeName = [string] $RuntimeEnvironment.name
    Write-RunLog "Inspecting packages in runtime environment '$runtimeName'."

    if (-not $RequiredModules -or $RequiredModules.Count -eq 0) {
        throw 'No required modules were configured in the runbook.'
    }

    $effectiveRequiredModules = @(Get-EffectiveRequiredModules -BaseRequiredModules $RequiredModules)
    $existingPackages = @(Get-RuntimeEnvironmentPackages -AutomationContext $AutomationContext -RuntimeEnvironmentName $runtimeName | Sort-Object name)
    $existingPackagesByName = @{}
    foreach ($existingPackage in $existingPackages) {
        $existingPackagesByName[[string] $existingPackage.name] = $existingPackage
    }

    $results = [System.Collections.Generic.List[object]]::new()
    foreach ($requiredPackage in @($effectiveRequiredModules | Sort-Object PackageName)) {
        $packageName = [string] $requiredPackage.PackageName
        if ([string]::IsNullOrWhiteSpace($packageName)) { continue }

        try {
            $targetPackage = Resolve-RequiredPackageVersion -PackageName $packageName -VersionDescriptor ([string] $requiredPackage.Version)
            $existingPackage = $existingPackagesByName[$packageName]
            $results.Add((Update-PackageInRuntimeEnvironment -AutomationContext $AutomationContext -RuntimeEnvironment $RuntimeEnvironment -RequiredPackage $requiredPackage -ExistingPackage $existingPackage -TargetPackage $targetPackage))
        } catch {
            Write-RunLog "Skipping package '$packageName': $($_.Exception.Message)" -Level Warning
            $results.Add([PSCustomObject]@{
                    RuntimeEnvironment = $runtimeName
                    PackageName        = $packageName
                    PreviousVersion    = if ($existingPackagesByName.ContainsKey($packageName)) { [string] $existingPackagesByName[$packageName].properties.version } else { $null }
                    TargetVersion      = $null
                    Changed            = $false
                    State              = if ($existingPackagesByName.ContainsKey($packageName)) { [string] $existingPackagesByName[$packageName].properties.provisioningState } else { 'NotInstalled' }
                    Skipped            = $true
                    Reason             = $_.Exception.Message
                })
        }
    }

    $results.ToArray()
}

function Get-NormalizedScriptText {
    param([Parameter(Mandatory = $false)][AllowNull()][string] $Text)

    if ([string]::IsNullOrEmpty($Text)) { return '' }

    $bom = [char]0xFEFF
    if ($Text[0] -eq $bom) { $Text = $Text.Substring(1) }

    $Text = $Text -replace "`r`n", "`n" -replace "`r", "`n"
    return $Text.Trim()
}

function Get-GithubRunbookContent {
    param([Parameter(Mandatory = $true)][string] $SourceFileName)

    $uri = "$GithubRawBase/$SourceFileName"
    Write-RunLog "Downloading latest runbook content from '$uri'."
    $response = Invoke-WebRequestWithRetry -Uri $uri
    return [string] $response.Content
}

function Get-RunbookContentByState {
    param(
        [Parameter(Mandatory = $true)][pscustomobject] $AutomationContext,
        [Parameter(Mandatory = $true)][string] $RunbookName,
        [Parameter(Mandatory = $true)][ValidateSet('Published', 'Draft')][string] $State
    )

    $escapedRunbook = [System.Uri]::EscapeDataString($RunbookName)
    $contentPath = if ($State -eq 'Draft') { 'draft/content' } else { 'content' }
    $uri = "https://management.azure.com/subscriptions/$($AutomationContext.SubscriptionId)/resourceGroups/$($AutomationContext.ResourceGroupName)/providers/Microsoft.Automation/automationAccounts/$($AutomationContext.AutomationAccountName)/runbooks/$escapedRunbook/$contentPath`?api-version=$($script:AutomationApiVersion)"
    $accessToken = Get-ArmAccessTokenPlainText

    try {
        $response = Invoke-WebRequest -Uri $uri -Headers @{ Authorization = "Bearer $accessToken" } -UseBasicParsing -TimeoutSec 120 -ErrorAction Stop
        $raw = $response.Content
        # The runbook content endpoint is served with a non-text content type, so Invoke-WebRequest exposes
        # the body as a byte[]. Casting that to [string] yields space-joined decimal codes, not the script,
        # so decode it as UTF-8 explicitly. A leading BOM (if any) is stripped later during normalization.
        if ($raw -is [byte[]]) {
            return [System.Text.Encoding]::UTF8.GetString($raw)
        }

        return [string] $raw
    } catch {
        $statusCode = $null
        if ($_.Exception.Response) {
            try { $statusCode = [int] $_.Exception.Response.StatusCode } catch { $statusCode = $null }
        }

        if ($statusCode -eq 404) { return $null }
        throw
    }
}

function Wait-ForRunbookContent {
    param(
        [Parameter(Mandatory = $true)][pscustomobject] $AutomationContext,
        [Parameter(Mandatory = $true)][string] $RunbookName,
        [Parameter(Mandatory = $true)][ValidateSet('Published', 'Draft')][string] $State,
        [Parameter(Mandatory = $true)][string] $ExpectedNormalizedContent,
        [Parameter(Mandatory = $true)][int] $TimeoutMinutes,
        [Parameter(Mandatory = $true)][int] $PollIntervalSeconds
    )

    $deadline = (Get-Date).AddMinutes($TimeoutMinutes)
    while ($true) {
        $content = Get-RunbookContentByState -AutomationContext $AutomationContext -RunbookName $RunbookName -State $State
        if ($null -ne $content -and (Get-NormalizedScriptText -Text $content) -ceq $ExpectedNormalizedContent) {
            return $true
        }

        if ((Get-Date) -ge $deadline) {
            return $false
        }

        Start-Sleep -Seconds $PollIntervalSeconds
    }
}

function Invoke-ArmWebRequest {
    param(
        [Parameter(Mandatory = $true)][ValidateSet('GET', 'PUT', 'POST')][string] $Method,
        [Parameter(Mandatory = $true)][string] $Uri,
        [Parameter(Mandatory = $false)][AllowNull()][object] $Body,
        [Parameter(Mandatory = $false)][string] $ContentType
    )

    $accessToken = Get-ArmAccessTokenPlainText
    $requestParams = @{
        Method          = $Method
        Uri             = $Uri
        Headers         = @{ Authorization = "Bearer $accessToken" }
        UseBasicParsing = $true
        TimeoutSec      = 120
        ErrorAction     = 'Stop'
    }

    if ($null -ne $Body) { $requestParams['Body'] = $Body }
    if (-not [string]::IsNullOrWhiteSpace($ContentType)) { $requestParams['ContentType'] = $ContentType }

    try {
        Invoke-WebRequest @requestParams
    } catch {
        $details = $null
        if ($_.ErrorDetails -and -not [string]::IsNullOrWhiteSpace($_.ErrorDetails.Message)) {
            $details = $_.ErrorDetails.Message
        }

        if ([string]::IsNullOrWhiteSpace($details)) { throw }

        throw "ARM $Method request failed for '$Uri'. $($_.Exception.Message) Details: $details"
    }
}

function Test-RunbookExists {
    param(
        [Parameter(Mandatory = $true)][pscustomobject] $AutomationContext,
        [Parameter(Mandatory = $true)][string] $RunbookName
    )

    $escapedRunbook = [System.Uri]::EscapeDataString($RunbookName)
    $uri = "https://management.azure.com/subscriptions/$($AutomationContext.SubscriptionId)/resourceGroups/$($AutomationContext.ResourceGroupName)/providers/Microsoft.Automation/automationAccounts/$($AutomationContext.AutomationAccountName)/runbooks/$($escapedRunbook)?api-version=$($script:AutomationApiVersion)"
    $accessToken = Get-ArmAccessTokenPlainText

    try {
        $null = Invoke-WebRequest -Uri $uri -Headers @{ Authorization = "Bearer $accessToken" } -UseBasicParsing -TimeoutSec 120 -ErrorAction Stop
        return $true
    } catch {
        $statusCode = $null
        if ($_.Exception.Response) {
            try { $statusCode = [int] $_.Exception.Response.StatusCode } catch { $statusCode = $null }
        }

        if ($statusCode -eq 404) { return $false }
        throw
    }
}

function New-RunbookShell {
    param(
        [Parameter(Mandatory = $true)][pscustomobject] $AutomationContext,
        [Parameter(Mandatory = $true)][string] $RunbookName,
        [Parameter(Mandatory = $true)][string] $RuntimeEnvironmentName,
        [Parameter(Mandatory = $true)][string] $Location
    )

    $escapedRunbook = [System.Uri]::EscapeDataString($RunbookName)
    $uri = "https://management.azure.com/subscriptions/$($AutomationContext.SubscriptionId)/resourceGroups/$($AutomationContext.ResourceGroupName)/providers/Microsoft.Automation/automationAccounts/$($AutomationContext.AutomationAccountName)/runbooks/$($escapedRunbook)?api-version=$($script:AutomationApiVersion)"
    $body = @{
        location   = $Location
        name       = $RunbookName
        properties = @{
            runbookType        = 'PowerShell'
            runtimeEnvironment = $RuntimeEnvironmentName
            logVerbose         = $false
            logProgress        = $false
            description        = "DriftMaester $RunbookName runbook."
        }
    }

    $null = Invoke-ArmRequest -Method PUT -Uri $uri -Body $body
}

function Import-RunbookDraftContent {
    param(
        [Parameter(Mandatory = $true)][pscustomobject] $AutomationContext,
        [Parameter(Mandatory = $true)][string] $RunbookName,
        [Parameter(Mandatory = $true)][string] $Content
    )

    $escapedRunbook = [System.Uri]::EscapeDataString($RunbookName)
    $uri = "https://management.azure.com/subscriptions/$($AutomationContext.SubscriptionId)/resourceGroups/$($AutomationContext.ResourceGroupName)/providers/Microsoft.Automation/automationAccounts/$($AutomationContext.AutomationAccountName)/runbooks/$($escapedRunbook)/draft/content?api-version=$($script:AutomationApiVersion)"

    $null = Invoke-ArmWebRequest -Method PUT -Uri $uri -Body $Content -ContentType 'text/plain'
}

function Publish-RunbookDraft {
    param(
        [Parameter(Mandatory = $true)][pscustomobject] $AutomationContext,
        [Parameter(Mandatory = $true)][string] $RunbookName
    )

    $escapedRunbook = [System.Uri]::EscapeDataString($RunbookName)
    $uri = "https://management.azure.com/subscriptions/$($AutomationContext.SubscriptionId)/resourceGroups/$($AutomationContext.ResourceGroupName)/providers/Microsoft.Automation/automationAccounts/$($AutomationContext.AutomationAccountName)/runbooks/$($escapedRunbook)/publish?api-version=$($script:AutomationApiVersion)"

    $null = Invoke-ArmWebRequest -Method POST -Uri $uri
}

function Update-ManagedRunbook {
    param(
        [Parameter(Mandatory = $true)][pscustomobject] $AutomationContext,
        [Parameter(Mandatory = $true)][string] $RuntimeEnvironmentName,
        [Parameter(Mandatory = $true)][string] $Location,
        [Parameter(Mandatory = $true)][pscustomobject] $Runbook
    )

    $runbookName = [string] $Runbook.RunbookName
    $sourceFileName = [string] $Runbook.SourceFileName
    $contentUri = "$GithubRawBase/$sourceFileName"

    $githubContent = Get-GithubRunbookContent -SourceFileName $sourceFileName
    $normalizedGithub = Get-NormalizedScriptText -Text $githubContent
    if ([string]::IsNullOrWhiteSpace($normalizedGithub)) {
        throw "Downloaded content for runbook '$runbookName' from '$contentUri' was empty."
    }

    $currentContent = Get-RunbookContentByState -AutomationContext $AutomationContext -RunbookName $runbookName -State Published
    if ($null -ne $currentContent -and (Get-NormalizedScriptText -Text $currentContent) -ceq $normalizedGithub) {
        Write-RunLog "Runbook '$runbookName' is already up to date with GitHub." -Level Success
        return [PSCustomObject]@{ RunbookName = $runbookName; Changed = $false; State = 'Current' }
    }

    if ($null -eq $currentContent) {
        Write-RunLog "Runbook '$runbookName' has no published content. Importing from '$contentUri'."
    } else {
        Write-RunLog "Runbook '$runbookName' is out of date with GitHub. Importing the new content from '$contentUri'."
    }

    # Only create the runbook when it is genuinely missing; re-creating an existing runbook risks changing
    # its runbook type or runtime environment binding.
    if (-not (Test-RunbookExists -AutomationContext $AutomationContext -RunbookName $runbookName)) {
        Write-RunLog "Runbook '$runbookName' does not exist yet. Creating it in runtime environment '$RuntimeEnvironmentName'."
        New-RunbookShell -AutomationContext $AutomationContext -RunbookName $runbookName -RuntimeEnvironmentName $RuntimeEnvironmentName -Location $Location
    }

    Import-RunbookDraftContent -AutomationContext $AutomationContext -RunbookName $runbookName -Content $githubContent

    # The draft content import is asynchronous. Publishing before it commits would promote the previous
    # draft and leave the published content stale, so wait for the draft to reflect the new content first.
    if (-not (Wait-ForRunbookContent -AutomationContext $AutomationContext -RunbookName $runbookName -State Draft -ExpectedNormalizedContent $normalizedGithub -TimeoutMinutes $RunbookSyncTimeoutMinutes -PollIntervalSeconds $PollIntervalSeconds)) {
        throw "Runbook '$runbookName' draft did not reflect the new GitHub content within $RunbookSyncTimeoutMinutes minute(s). The content import did not take effect."
    }

    Publish-RunbookDraft -AutomationContext $AutomationContext -RunbookName $runbookName

    if (-not (Wait-ForRunbookContent -AutomationContext $AutomationContext -RunbookName $runbookName -State Published -ExpectedNormalizedContent $normalizedGithub -TimeoutMinutes $RunbookSyncTimeoutMinutes -PollIntervalSeconds $PollIntervalSeconds)) {
        throw "Runbook '$runbookName' was published but its content did not match the GitHub source within $RunbookSyncTimeoutMinutes minute(s). The publish did not take effect."
    }

    Write-RunLog "Runbook '$runbookName' was imported from GitHub and published successfully. Updates to the currently running runbook take effect on its next run." -Level Success

    return [PSCustomObject]@{ RunbookName = $runbookName; Changed = $true; State = 'Published' }
}

function Update-ManagedRunbooks {
    param([Parameter(Mandatory = $true)][pscustomobject] $AutomationContext, [Parameter(Mandatory = $true)][string] $RuntimeEnvironmentName)

    $automationAccount = Get-AutomationAccount -AutomationContext $AutomationContext
    $location = [string] $automationAccount.location
    if ([string]::IsNullOrWhiteSpace($location)) {
        throw "Could not determine location for Automation Account '$($AutomationContext.AutomationAccountName)'."
    }

    $results = [System.Collections.Generic.List[object]]::new()
    foreach ($runbook in $ManagedRunbooks) {
        try {
            $results.Add((Update-ManagedRunbook -AutomationContext $AutomationContext -RuntimeEnvironmentName $RuntimeEnvironmentName -Location $location -Runbook $runbook))
        } catch {
            Write-RunLog "Skipping runbook '$($runbook.RunbookName)': $($_.Exception.Message)" -Level Warning
            $results.Add([PSCustomObject]@{
                    RunbookName = [string] $runbook.RunbookName
                    Changed     = $false
                    State       = 'Skipped'
                    Skipped     = $true
                    Reason      = $_.Exception.Message
                })
        }
    }

    $results.ToArray()
}

try {
    Import-Module Az.Accounts -ErrorAction Stop

    Connect-RunbookIdentity

    $automationContextCall = Invoke-WithDeferredRunLog -ScriptBlock { Resolve-AutomationAccountContext }
    Write-DeferredRunLog
    if ($automationContextCall.ErrorRecord) { throw $automationContextCall.ErrorRecord }
    $automationContext = $automationContextCall.Output | Select-Object -First 1
    Write-RunLog "Resolved Automation Account: $($automationContext.AutomationAccountName) in resource group '$($automationContext.ResourceGroupName)' (subscription $($automationContext.SubscriptionId))."

    $targetRuntimeCall = Invoke-WithDeferredRunLog -ScriptBlock { Resolve-TargetRuntimeEnvironment -AutomationContext $automationContext }
    Write-DeferredRunLog
    if ($targetRuntimeCall.ErrorRecord) { throw $targetRuntimeCall.ErrorRecord }
    $targetRuntime = $targetRuntimeCall.Output | Select-Object -First 1
    Write-RunLog "Selected runtime environment '$($targetRuntime.name)' (PowerShell $($targetRuntime.properties.runtime.version))."

    $resultsCall = Invoke-WithDeferredRunLog -ScriptBlock { Update-RequiredPackagesInRuntimeEnvironment -AutomationContext $automationContext -RuntimeEnvironment $targetRuntime }
    Write-DeferredRunLog
    if ($resultsCall.ErrorRecord) { throw $resultsCall.ErrorRecord }
    $results = @($resultsCall.Output)
    $changedCount = @($results | Where-Object { $_.Changed }).Count
    $skippedCount = @($results | Where-Object { $_.Skipped }).Count
    $unchangedCount = $results.Count - $changedCount - $skippedCount

    Write-RunLog "Required package update completed. Runtime='$($targetRuntime.name)', Total='$($results.Count)', Updated='$changedCount', Current='$unchangedCount', Skipped='$skippedCount'." -Level Success

    $runbookResultsCall = Invoke-WithDeferredRunLog -ScriptBlock { Update-ManagedRunbooks -AutomationContext $automationContext -RuntimeEnvironmentName $targetRuntime.name }
    Write-DeferredRunLog
    if ($runbookResultsCall.ErrorRecord) { throw $runbookResultsCall.ErrorRecord }
    $runbookResults = @($runbookResultsCall.Output)
    $runbookChanged = @($runbookResults | Where-Object { $_.Changed }).Count
    $runbookSkipped = @($runbookResults | Where-Object { $_.Skipped }).Count
    $runbookCurrent = $runbookResults.Count - $runbookChanged - $runbookSkipped

    Write-RunLog "Runbook sync from GitHub completed. Total='$($runbookResults.Count)', Updated='$runbookChanged', Current='$runbookCurrent', Skipped='$runbookSkipped'." -Level Success
} catch {
    Write-RunLog "Update-DriftMaester runbook failed: $($_.Exception.Message)" -Level Error
    throw
}
