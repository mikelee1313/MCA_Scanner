<#
.SYNOPSIS
    MCA Data Collection Scanner — collects Microsoft 365 workload data for
    Microsoft Cloud Adoption assessments.

.DESCRIPTION
    Interactive tool that presents a workload menu, then collects configuration
    and usage data for the selected workload(s).  Each workload's data is
    exported to timestamped CSV files in $OutputFolder.

    AUTHENTICATION:
      SharePoint Online [1] and OneDrive for Business [4] use the
      Microsoft.Online.SharePoint.PowerShell module with interactive
      (delegated) authentication via Connect-SPOService.  No Entra ID app
      registration is required — any SharePoint Administrator can run these
      workloads.

      Each workload prefers its native administration module.  SharePoint Online
      and OneDrive use Microsoft.Online.SharePoint.PowerShell via
      Connect-SPOService; Exchange and Security/Compliance use
      ExchangeOnlineManagement via Connect-ExchangeOnline / Connect-IPPSSession;
      Teams uses MicrosoftTeams via Connect-MicrosoftTeams; Power BI and Power
      Platform use their native PowerShell modules when present.  Microsoft Graph
      is retained for Entra ID and for report/security data where Microsoft 365
      has no equivalent workload-specific PowerShell cmdlet.

    SharePoint Online [1] & OneDrive for Business [4] require:
      Microsoft.Online.SharePoint.PowerShell module
        Install-Module Microsoft.Online.SharePoint.PowerShell -Scope CurrentUser
      SharePoint Administrator role (delegated, interactive login)
      $tenantUrl configured in the Configuration section

    Exchange Online [2] additionally requires:
      ExchangeOnlineManagement module  (Install-Module ExchangeOnlineManagement -Scope CurrentUser)
      Azure AD application permission 'Exchange.ManageAsApp' (grant + admin consent)
      Service principal assigned 'Exchange Administrator' or 'Exchange Recipient Administrator'
      role in Entra ID > Roles and administrators.

    Microsoft Teams [3] additionally requires:
      MicrosoftTeams module  (Install-Module MicrosoftTeams -Scope CurrentUser)
      Service principal assigned 'Teams Administrator' or 'Global Reader' role
      in Entra ID > Roles and administrators.

    Security & Compliance [6] additionally requires:
      ExchangeOnlineManagement module  (Install-Module ExchangeOnlineManagement -Scope CurrentUser)
      Service principal assigned 'Compliance Administrator' or equivalent role
      for Connect-IPPSSession based collection.

    Entra ID [5] / Graph-backed report and security fallback data require:
      Directory.Read.All, Reports.Read.All, Policy.Read.All,
            SecurityEvents.Read.All, InformationProtectionPolicy.Read,
      AuditLog.Read.All

    Power Platform [7] additionally requires:
      'Power BI Service' Tenant.Read.All application permission
      'Power Platform Administrator' role on the service principal

.NOTES
    Author : Mike Lee
    Version: 3.0
#>

function Start-ScannerInCleanPwshIfNeeded {
    <#
    .SYNOPSIS
        Relaunches this script in a clean PowerShell 7 host when running from
        Windows PowerShell or VS Code-integrated hosts, which often preload
        assemblies that conflict with MicrosoftTeams/Microsoft.Graph auth stacks.
    #>
    if ($env:MCA_SCANNER_ISOLATED -eq '1') { return }

    $needsPwsh7 = $PSVersionTable.PSEdition -ne 'Core' -or $PSVersionTable.PSVersion.Major -lt 7
    $isVsCodeHost = ($Host.Name -match 'Visual Studio Code') -or ($env:TERM_PROGRAM -eq 'vscode')

    if (-not ($needsPwsh7 -or $isVsCodeHost)) { return }

    $pwsh = Get-Command -Name 'pwsh' -ErrorAction SilentlyContinue
    if (-not $pwsh) {
        Write-Warning 'PowerShell 7 (pwsh) was not found. Continue in current host; auth module conflicts may occur.'
        return
    }

    if (-not $PSCommandPath) { return }

    Write-Host 'Relaunching MCA scanner in clean PowerShell 7 process (-NoProfile) to avoid module assembly conflicts...' -ForegroundColor Yellow

    $env:MCA_SCANNER_ISOLATED = '1'
    & $pwsh.Source -NoProfile -ExecutionPolicy Bypass -File $PSCommandPath
    exit $LASTEXITCODE
}

Start-ScannerInCleanPwshIfNeeded

#region Configuration
##############################################################
#                    CONFIGURATION SECTION                   #
##############################################################

# ---- Debug output ($true = verbose Graph call tracing) ----
$debug = $true

# ---- Tenant ID ----
$tenantId = '9cfc42cb-51da-4055-87e9-b20a170b6ba3'   # Tenant ID or verified domain  e.g. 'contoso.onmicrosoft.com'

# ---- Tenant root SharePoint URL (required for SharePoint workload PnP calls) ----
$tenantUrl = 'https://m365cpi13246019.sharepoint.com'   # e.g. 'https://contoso.sharepoint.com'

# ---- Output folder for exported CSV files ----
$OutputFolder = "$env:USERPROFILE\Documents\MCA_Assessment"

# ---- Throttle / retry settings ----
$MaxRetries = 15
$InitialBackoffSec = 3
$RequestTimeoutSec = 300

# ---- Unlicensed OneDrive report polling ----
$SPOExportPollIntervalSec = 5    # seconds between readiness polls
$SPOExportMaxWaitSec = 300  # max seconds to wait for export file

##############################################################
#                  END CONFIGURATION SECTION                 #
##############################################################
#endregion Configuration

#region Initialization
$date = Get-Date -Format 'yyyyMMdd_HHmmss'
$today = (Get-Date).Date
$script:assemblyConflictHintShown = $false

$global:token = $null
$global:tokenExpiry = $null

# Ensure output directory exists
if (-not (Test-Path $OutputFolder)) {
    New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null
}

$RunLog = Join-Path $OutputFolder "MCA_Scan_$date.log"

function Write-Log {
    param(
        [string] $Message,
        [ValidateSet('INFO', 'SUCCESS', 'WARN', 'ERROR', 'DEBUG')]
        [string] $Level = 'INFO'
    )
    $ts = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $entry = "$ts [$Level] $Message"
    Add-Content -Path $RunLog -Value $entry -Encoding UTF8
    switch ($Level) {
        'SUCCESS' { Write-Host $Message -ForegroundColor Green }
        'WARN' { Write-Host $Message -ForegroundColor Yellow }
        'ERROR' { Write-Host $Message -ForegroundColor Red }
        'DEBUG' { if ($debug) { Write-Host $Message -ForegroundColor DarkGray } }
        default { Write-Host $Message -ForegroundColor Cyan }
    }
}

function Initialize-ModuleCompatibility {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [bool] $NeedsTeamsCompatibility
    )

    if (-not $NeedsTeamsCompatibility) { return }

    # Pre-import Teams only when needed. This keeps Entra-only runs from loading
    # auth assemblies that can conflict with Microsoft.Graph authentication.
    if (Get-Module -ListAvailable -Name 'MicrosoftTeams') {
        try {
            Import-Module MicrosoftTeams -ErrorAction Stop
            Write-Log '  Module compatibility: preloaded MicrosoftTeams for Teams workload.' 'DEBUG'
        }
        catch {
            Write-Log "  Module compatibility preload warning (MicrosoftTeams): $($_.Exception.Message)" 'WARN'
        }
    }
}
#endregion Initialization

#region Graph Helper Functions

function Invoke-GraphRequestWithThrottleHandling {
    <#
    .SYNOPSIS
        Wrapper around Invoke-MgGraphRequest.
        Throttle handling and token refresh are provided automatically by the Microsoft.Graph SDK.
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)] [string]    $Uri,
        [Parameter(Mandatory)] [string]    $Method,
        [Parameter()]          [hashtable] $Headers = @{},
        [Parameter()]          [string]    $Body = $null,
        [Parameter()]          [string]    $ContentType = 'application/json',
        [Parameter()]          [int]       $MaxRetries = $script:MaxRetries,
        [Parameter()]          [int]       $InitialBackoffSeconds = $script:InitialBackoffSec,
        [Parameter()]          [int]       $TimeoutSeconds = $script:RequestTimeoutSec
    )

    Test-ValidToken
    Write-Log "  Graph -> $Method $Uri" 'DEBUG'

    $invokeParams = @{
        Uri         = $Uri
        Method      = $Method
        OutputType  = 'PSObject'
        ErrorAction = 'Stop'
    }
    if ($Body) {
        $invokeParams['Body'] = $Body
        $invokeParams['ContentType'] = $ContentType
    }
    return Invoke-MgGraphRequest @invokeParams
}

function Invoke-GraphPagedRequest {
    <#
    .SYNOPSIS
        Executes a Graph GET and automatically follows @odata.nextLink pages,
        returning all results as a single flat list.
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)] [string] $Uri
    )

    $results = [System.Collections.Generic.List[object]]::new()
    $nextLink = $Uri

    do {
        Test-ValidToken
        $response = Invoke-MgGraphRequest -Uri $nextLink -Method GET -OutputType PSObject

        if ($null -ne $response.value) {
            $results.AddRange([object[]]$response.value)
        }
        else {
            return $response
        }

        $nextLink = $response.'@odata.nextLink'
        if ($nextLink) { Write-Log '  Fetching next page...' 'DEBUG' }
    } while ($nextLink)

    return $results
}

function Invoke-GraphReportRequest {
    <#
    .SYNOPSIS
        Downloads a Graph Reports API CSV endpoint and saves it directly to disk.
        Returns the saved file path, or $null on failure.
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)] [string] $Uri,
        [Parameter(Mandatory)] [string] $FileName
    )

    $outPath = Join-Path $OutputFolder $FileName
    try {
        Test-ValidToken
        Invoke-MgGraphRequest -Uri $Uri -Method GET -OutputFilePath $outPath -ErrorAction Stop
        Write-Log "  Report saved -> $outPath" 'SUCCESS'
        return $outPath
    }
    catch {
        Write-Log "  Report download failed [$FileName]: $($_.Exception.Message)" 'WARN'
        if (-not $script:assemblyConflictHintShown -and $_.Exception.Message -match 'Method not found' -and $_.Exception.Message -match '(Azure\.Identity|Microsoft\.Identity\.Client)') {
            $script:assemblyConflictHintShown = $true
            Write-Log '  Detected auth assembly conflict in the current host. Re-run with: pwsh -NoProfile -File "<path-to-this-script>"' 'WARN'
            Write-Log '  If needed, update modules: MicrosoftTeams and Microsoft.Graph (CurrentUser scope).' 'WARN'
        }
        return $null
    }
}

#endregion Graph Helper Functions

#region Authentication Functions

$global:mgConnected = $false

function Connect-ToMicrosoftGraph {
    <#
    .SYNOPSIS
        Connects to Microsoft Graph using interactive (delegated) authentication.
        Requires: Install-Module Microsoft.Graph -Scope CurrentUser
    #>
    Write-Log "Connecting to Microsoft Graph (interactive)..."

    if (-not (Get-Module -ListAvailable -Name 'Microsoft.Graph.Authentication')) {
        throw "Microsoft.Graph.Authentication module not found.`n  Install with: Install-Module Microsoft.Graph -Scope CurrentUser"
    }
    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop

    $scopes = @(
        'Directory.Read.All',
        'Reports.Read.All',
        'Policy.Read.All',
        'SecurityEvents.Read.All',
        'InformationProtectionPolicy.Read',
        'AuditLog.Read.All',
        'UserAuthenticationMethod.Read.All',
        'RoleManagement.Read.Directory',
        'PrivilegedAccess.Read.AzureAD',
        'AccessReview.Read.All',
        'Application.Read.All',
        'DelegatedPermissionGrant.Read.All',
        'DeviceManagementManagedDevices.Read.All',
        'DeviceManagementConfiguration.Read.All',
        'ServiceMessage.Read.All',
        'ExternalConnection.Read.All',
        'Organization.Read.All'
    )

    try {
        # Preferred path: browser-based delegated auth.
        Connect-MgGraph -TenantId $tenantId -Scopes $scopes -NoWelcome -ErrorAction Stop
    }
    catch {
        $msg = $_.Exception.Message
        $isKnownMsalAssemblyConflict = $msg -match 'Method not found' -and
        $msg -match 'BaseAbstractApplicationBuilder`1\.WithLogging' -and
        $msg -match '(Microsoft\.Identity\.Client|Microsoft\.IdentityModel\.Abstractions)'

        if ($isKnownMsalAssemblyConflict) {
            Write-Log '  Graph interactive browser auth failed due to a known MSAL assembly conflict in the current host.' 'WARN'
            Write-Log '  Retrying Graph auth using device code flow (bypasses browser-credential path)...' 'WARN'

            try {
                Connect-MgGraph -TenantId $tenantId -Scopes $scopes -UseDeviceAuthentication -NoWelcome -ErrorAction Stop
            }
            catch {
                Write-Log '  Device code fallback also failed. Recommended remediation:' 'WARN'
                Write-Log '    1) Close all PowerShell hosts and re-run using: pwsh -NoProfile -File "<path-to-this-script>"' 'WARN'
                Write-Log '    2) Update modules (CurrentUser): Microsoft.Graph and MicrosoftTeams' 'WARN'
                throw
            }
        }
        else {
            throw
        }
    }

    $global:mgConnected = $true
    Write-Log "  Connected to Microsoft Graph (interactive delegated auth)" 'SUCCESS'
}

function Test-ValidToken {
    if (-not $global:mgConnected) {
        Connect-ToMicrosoftGraph
        return
    }
    $ctx = Get-MgContext -ErrorAction SilentlyContinue
    if (-not $ctx) {
        Write-Log 'Microsoft Graph session expired — reconnecting...' 'WARN'
        $global:mgConnected = $false
        Connect-ToMicrosoftGraph
    }
}

function Get-GraphAuthHeaders {
    # Authentication is handled automatically by Invoke-MgGraphRequest / Microsoft.Graph SDK.
    # Retained for call-site compatibility — no manual bearer token required.
    Test-ValidToken
    return @{}
}

function Get-AppTokenForScope {
    <#
    .SYNOPSIS
        Not used in interactive auth mode.  App-only token acquisition requires an
        Entra ID app registration, which is not configured for interactive auth.
        Native module auth (Connect-ToPowerBIService / Connect-ToPowerPlatformAdmin)
        is used instead; REST-only fallbacks for Power Platform are skipped.
    #>
    param([Parameter(Mandatory)][string]$Scope)
    return $null
}

#endregion Authentication Functions

#region PnP Helper Functions

function Connect-ToPnPSite {
    <#
    .SYNOPSIS
        Connects PnP.PowerShell to the specified site URL using the configured
        authentication type.  Defaults to $tenantUrl if $Url is omitted.
    #>
    [CmdletBinding()]
    param (
        [Parameter()] [string] $Url = $script:tenantUrl
    )

    if ([string]::IsNullOrWhiteSpace($Url)) {
        throw 'No site URL provided and $tenantUrl is empty. Set $tenantUrl in the Configuration section.'
    }

    if (-not (Get-Module -ListAvailable -Name 'PnP.PowerShell')) {
        throw "PnP.PowerShell module not found. Install it with: Install-Module PnP.PowerShell -Scope CurrentUser"
    }
    Import-Module PnP.PowerShell -ErrorAction Stop
}

function Invoke-PnPWithRetry {
    <#
    .SYNOPSIS
        Executes a script block with Retry-After / exponential-backoff handling.
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)] [scriptblock] $ScriptBlock,
        [Parameter()]          [int]         $MaxRetries = $script:MaxRetries,
        [Parameter()]          [int]         $InitialBackoffSeconds = $script:InitialBackoffSec
    )

    $retryCount = 0
    $backoffSec = $InitialBackoffSeconds

    while ($retryCount -le $MaxRetries) {
        try {
            return & $ScriptBlock
        }
        catch {
            $statusCode = $null
            if ($_.Exception.Response) { $statusCode = [int]$_.Exception.Response.StatusCode }
            if (-not $statusCode -and $_.Exception.Message -match '(429|502|503|504)') {
                $statusCode = [int]$Matches[1]
            }

            $isRetryable = $statusCode -in @(429, 502, 503, 504) -or
            ($_.Exception -is [System.Net.WebException] -and
            $_.Exception.Status -in @(
                [System.Net.WebExceptionStatus]::Timeout,
                [System.Net.WebExceptionStatus]::ConnectionClosed))

            if (-not $isRetryable) { throw $_ }
            if ($retryCount -ge $MaxRetries) { Write-Warning "Max retries ($MaxRetries) reached."; throw $_ }

            $waitSec = $backoffSec
            if ($statusCode -in @(429, 503)) {
                try {
                    $ra = $_.Exception.Response.Headers['Retry-After']
                    if ($ra) { $waitSec = [int]$ra }
                }
                catch {}
            }

            $retryCount++
            Write-Log "    Throttled ($statusCode). Waiting ${waitSec}s (attempt $retryCount/$MaxRetries)..." 'WARN'
            Start-Sleep -Seconds $waitSec
            $backoffSec = [Math]::Min($backoffSec * 2, 300)
        }
    }
}

#endregion PnP Helper Functions

#region MicrosoftTeams Helper Functions

function Connect-ToMicrosoftTeams {
    <#
    .SYNOPSIS
        Connects the MicrosoftTeams PowerShell module using interactive (delegated) authentication.
    #>
    if (-not (Get-Module -ListAvailable -Name 'MicrosoftTeams')) {
        throw "MicrosoftTeams module not found. Install with: Install-Module MicrosoftTeams -Scope CurrentUser"
    }
    Import-Module MicrosoftTeams -ErrorAction Stop

    # Force a delegated user sign-in with the native Teams module.
    # Do not use app-only/certificate parameters for this workload.
    try { Disconnect-MicrosoftTeams -Confirm:$false -ErrorAction SilentlyContinue | Out-Null } catch {}
    Connect-MicrosoftTeams -TenantId $tenantId -ErrorAction Stop
    Write-Log "  Connected to MicrosoftTeams module (interactive delegated auth)" 'SUCCESS'
}

function ConvertTo-FlatCsvRow {
    <#
    .SYNOPSIS
        Flattens a PSObject into an [ordered] hashtable suitable for Export-Csv.
        Nested objects and arrays are JSON-encoded.  Prepends a CollectDate stamp.
    #>
    param([Parameter(Mandatory)] [object] $InputObject)
    $row = [ordered]@{ CollectDate = (Get-Date -Format 'yyyy-MM-dd HH:mm') }
    foreach ($prop in $InputObject.PSObject.Properties) {
        if ($prop.Name -like '@*') { continue }
        $val = $prop.Value
        if ($null -eq $val) {
            $val = ''
        }
        elseif ($val -is [System.Collections.IEnumerable] -and $val -isnot [string]) {
            $val = ($val | ConvertTo-Json -Compress -Depth 3 -ErrorAction SilentlyContinue)
        }
        elseif ($val -is [PSCustomObject]) {
            $val = ($val | ConvertTo-Json -Compress -Depth 3 -ErrorAction SilentlyContinue)
        }
        $row[$prop.Name] = $val
    }
    return [PSCustomObject]$row
}

function ConvertTo-ByteCount {
    <#
    .SYNOPSIS
        Converts Exchange/PowerShell size values to bytes for aggregate-only reporting.
    #>
    param([Parameter()] [object] $Value)

    if ($null -eq $Value) { return 0L }

    try {
        if ($Value.PSObject.Methods.Name -contains 'ToBytes') {
            return [int64]$Value.ToBytes()
        }
    }
    catch {}

    $text = [string]$Value
    if ([string]::IsNullOrWhiteSpace($text)) { return 0L }

    if ($text -match '\(([\d,]+)\s+bytes\)') {
        return [int64](($Matches[1] -replace ',', ''))
    }

    if ($text -match '([\d,.]+)\s*(TB|GB|MB|KB|B|bytes)') {
        $number = [double](($Matches[1] -replace ',', ''))
        switch -Regex ($Matches[2]) {
            '^TB$' { return [int64]($number * 1TB) }
            '^GB$' { return [int64]($number * 1GB) }
            '^MB$' { return [int64]($number * 1MB) }
            '^KB$' { return [int64]($number * 1KB) }
            default { return [int64]$number }
        }
    }

    return 0L
}

function New-ObjectCountSummary {
    param(
        [Parameter(Mandatory)] [string] $ObjectType,
        [Parameter(Mandatory)] [int] $Count,
        [Parameter()] [hashtable] $AdditionalProperties = @{}
    )

    $row = [ordered]@{
        ObjectType  = $ObjectType
        ObjectCount = $Count
        CollectDate = (Get-Date -Format 'yyyy-MM-dd HH:mm')
    }
    foreach ($key in $AdditionalProperties.Keys) {
        $row[$key] = $AdditionalProperties[$key]
    }
    return [PSCustomObject]$row
}

function Export-GroupedCountSummary {
    param(
        [Parameter(Mandatory)] [object[]] $Data,
        [Parameter(Mandatory)] [string] $FileName,
        [Parameter(Mandatory)] [string] $GroupProperty,
        [Parameter(Mandatory)] [string] $ObjectType
    )

    $rows = @($Data | Group-Object -Property $GroupProperty | ForEach-Object {
            New-ObjectCountSummary -ObjectType $ObjectType -Count $_.Count -AdditionalProperties @{
                GroupingProperty = $GroupProperty
                GroupingValue    = if ([string]::IsNullOrWhiteSpace([string]$_.Name)) { 'Unspecified' } else { $_.Name }
            }
        })

    if ($rows.Count -eq 0) {
        $rows = @(New-ObjectCountSummary -ObjectType $ObjectType -Count 0 -AdditionalProperties @{
                GroupingProperty = $GroupProperty
                GroupingValue    = 'None'
            })
    }

    Export-ToCsv -Data $rows -FileName $FileName
}

#endregion MicrosoftTeams Helper Functions

#region ExchangeOnline Helper Functions

function Connect-ToExchangeOnline {
    <#
    .SYNOPSIS
        Connects the ExchangeOnlineManagement module using interactive (delegated) authentication.
    #>
    if (-not (Get-Module -ListAvailable -Name 'ExchangeOnlineManagement')) {
        throw "ExchangeOnlineManagement module not found. Install with: Install-Module ExchangeOnlineManagement -Scope CurrentUser"
    }
    Import-Module ExchangeOnlineManagement -ErrorAction Stop

    Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
    Write-Log "  Connected to Exchange Online (interactive delegated auth)" 'SUCCESS'
}

#endregion ExchangeOnline Helper Functions

#region Security & Compliance Helper Functions

function Connect-ToSecurityCompliance {
    <#
    .SYNOPSIS
        Connects the Security & Compliance PowerShell session using
        ExchangeOnlineManagement's Connect-IPPSSession cmdlet with interactive auth.
    #>
    if (-not (Get-Module -ListAvailable -Name 'ExchangeOnlineManagement')) {
        throw "ExchangeOnlineManagement module not found. Install with: Install-Module ExchangeOnlineManagement -Scope CurrentUser"
    }
    Import-Module ExchangeOnlineManagement -ErrorAction Stop

    $connectParams = @{
        ShowBanner  = $false
        ErrorAction = 'Stop'
    }
    $ippsCommand = Get-Command -Name Connect-IPPSSession -ErrorAction Stop
    if ($ippsCommand.Parameters.ContainsKey('CommandName')) {
        $connectParams['CommandName'] = @(
            'Get-Label',
            'Get-DlpCompliancePolicy',
            'Get-DlpComplianceRule',
            'Get-AdminAuditLogConfig',
            'Get-UnifiedAuditLogRetentionPolicy',
            'Get-ProtectionAlert',
            'Search-UnifiedAuditLog'
        )
    }

    Connect-IPPSSession @connectParams
    Write-Log "  Connected to Security & Compliance PowerShell (interactive delegated auth)" 'SUCCESS'
}

#endregion Security & Compliance Helper Functions

#region Power Platform Helper Functions

function Connect-ToPowerBIService {
    <#
    .SYNOPSIS
        Connects MicrosoftPowerBIMgmt.Profile using interactive (delegated) authentication.
    #>
    if (-not (Get-Module -ListAvailable -Name 'MicrosoftPowerBIMgmt.Profile')) {
        throw "MicrosoftPowerBIMgmt module not found. Install with: Install-Module MicrosoftPowerBIMgmt -Scope CurrentUser"
    }
    Import-Module MicrosoftPowerBIMgmt.Profile -ErrorAction Stop

    Connect-PowerBIServiceAccount -ErrorAction Stop | Out-Null
    Write-Log "  Connected to Power BI PowerShell module (interactive delegated auth)" 'SUCCESS'
}

function Connect-ToPowerPlatformAdmin {
    <#
    .SYNOPSIS
        Connects Microsoft.PowerApps.Administration.PowerShell using interactive (delegated) authentication.
    #>
    if (-not (Get-Module -ListAvailable -Name 'Microsoft.PowerApps.Administration.PowerShell')) {
        throw "Microsoft.PowerApps.Administration.PowerShell module not found. Install with: Install-Module Microsoft.PowerApps.Administration.PowerShell -Scope CurrentUser"
    }

    $imported = $false
    try {
        Import-Module Microsoft.PowerApps.Administration.PowerShell -ErrorAction Stop
        $imported = $true
    }
    catch {
        $msg = $_.Exception.Message
        $isAssemblyConflict = $msg -match 'Microsoft\.Identity\.Client' -or
        $msg -match 'already loaded' -or
        $msg -match 'Could not load file or assembly'

        $canUseWinPSCompat = $PSVersionTable.PSVersion.Major -gt 5 -and
        (Get-Command Import-Module).Parameters.ContainsKey('UseWindowsPowerShell')

        if ($isAssemblyConflict -and $canUseWinPSCompat) {
            Write-Log '  Power Platform module import hit auth assembly conflict; retrying via Windows PowerShell compatibility session...' 'WARN'
            try {
                Remove-Module Microsoft.PowerApps.Administration.PowerShell -ErrorAction SilentlyContinue
                Remove-Module Microsoft.PowerApps.AuthModule -ErrorAction SilentlyContinue
                Import-Module Microsoft.PowerApps.Administration.PowerShell -UseWindowsPowerShell -DisableNameChecking -ErrorAction Stop
                $imported = $true
            }
            catch {
                throw "Power Platform module import failed after compatibility retry: $($_.Exception.Message)"
            }
        }
        else {
            throw
        }
    }

    if (-not $imported -or -not (Get-Command -Name 'Add-PowerAppsAccount' -ErrorAction SilentlyContinue)) {
        throw "Power Platform admin cmdlets are unavailable in this session after module import."
    }

    Add-PowerAppsAccount -Endpoint 'prod' -TenantID $tenantId -ErrorAction Stop | Out-Null
    Write-Log "  Connected to Power Platform admin PowerShell module (interactive delegated auth)" 'SUCCESS'
}

function Get-PowerPlatformEnvironmentsIsolated {
    <#
    .SYNOPSIS
        Collects Power Platform environments in a separate pwsh -NoProfile process
        to avoid in-process Microsoft.Identity.Client assembly conflicts.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string] $TenantId
    )

    $pwsh = Get-Command -Name 'pwsh' -ErrorAction SilentlyContinue
    if (-not $pwsh) {
        throw 'PowerShell 7 (pwsh) was not found for isolated Power Platform collection.'
    }

    $tmpPath = Join-Path ([System.IO.Path]::GetTempPath()) ("mca_pp_env_{0}_{1}.json" -f $PID, [Guid]::NewGuid().ToString('N'))
    $scriptText = @"
`$ErrorActionPreference = 'Stop'
`$ProgressPreference = 'SilentlyContinue'
Import-Module Microsoft.PowerApps.Administration.PowerShell -ErrorAction Stop
Add-PowerAppsAccount -Endpoint 'prod' -TenantID '$TenantId' -ErrorAction Stop | Out-Null
`$envs = @(Get-AdminPowerAppEnvironment -ErrorAction Stop)
`$rows = foreach (`$env in `$envs) {
    [PSCustomObject]@{
        Name              = `$env.name
        DisplayName       = if (`$env.DisplayName) { `$env.DisplayName } else { `$env.properties.displayName }
        Location          = `$env.location
        Type              = if (`$env.EnvironmentType) { `$env.EnvironmentType } else { `$env.properties.environmentSku }
        ProvisioningState = if (`$env.ProvisioningState) { `$env.ProvisioningState } else { `$env.properties.provisioningState }
        IsDefault         = if (`$null -ne `$env.IsDefault) { `$env.IsDefault } else { `$env.properties.isDefault }
        CreatedTime       = if (`$env.CreatedTime) { `$env.CreatedTime } else { `$env.properties.createdTime }
    }
}
`$rows | ConvertTo-Json -Depth 6 -Compress
"@

    try {
        $cmd = "& { $scriptText } | Set-Content -Path '$tmpPath' -Encoding UTF8"
        & $pwsh.Source -NoProfile -Command $cmd
        if (-not (Test-Path $tmpPath)) {
            throw 'Isolated Power Platform collection did not produce output.'
        }

        $json = Get-Content -Path $tmpPath -Raw -ErrorAction Stop
        if ([string]::IsNullOrWhiteSpace($json)) { return @() }

        $parsed = $json | ConvertFrom-Json -ErrorAction Stop
        if ($parsed -is [System.Array]) { return @($parsed) }
        return @($parsed)
    }
    finally {
        Remove-Item -Path $tmpPath -Force -ErrorAction SilentlyContinue
    }
}

#endregion Power Platform Helper Functions

#region Menu & UI

function Show-Banner {
    Clear-Host
    Write-Host ""
    Write-Host "  ============================================================" -ForegroundColor Cyan
    Write-Host "     MCA Data Collection Scanner" -ForegroundColor White
    Write-Host "     Microsoft Cloud Adoption Assessment" -ForegroundColor White
    Write-Host "     $(Get-Date -Format 'dddd, MMMM dd yyyy   HH:mm')" -ForegroundColor DarkGray
    Write-Host "  ============================================================" -ForegroundColor Cyan
    Write-Host ""
}

function Show-WorkloadMenu {
    Show-Banner
    Write-Host "  Select a workload to scan:" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "    [1]  SharePoint Online" -ForegroundColor White
    Write-Host "    [2]  Exchange Online" -ForegroundColor White
    Write-Host "    [3]  Microsoft Teams" -ForegroundColor White
    Write-Host "    [4]  OneDrive for Business" -ForegroundColor White
    Write-Host "    [5]  Entra ID  (Users / Groups / Licenses)" -ForegroundColor White
    Write-Host "    [6]  Security & Compliance" -ForegroundColor White
    Write-Host "    [7]  Power Platform" -ForegroundColor White
    Write-Host "    [8]  Information Barriers" -ForegroundColor White
    Write-Host "    [A]  All Workloads" -ForegroundColor Green
    Write-Host "    [Q]  Quit" -ForegroundColor Red
    Write-Host ""
    Write-Host "  Output folder: $OutputFolder" -ForegroundColor DarkGray
    Write-Host ""

    $valid = @('1', '2', '3', '4', '5', '6', '7', '8', 'A', 'Q')
    do {
        $choice = (Read-Host "  Enter choice").Trim().ToUpper()
    } while ($choice -notin $valid)

    return $choice
}

function Export-ToCsv {
    param(
        [Parameter(Mandatory)] [AllowEmptyCollection()] [object[]] $Data,
        [Parameter(Mandatory)] [string]   $FileName
    )
    if ($null -eq $Data -or $Data.Count -eq 0) {
        Write-Log "  No data returned for: $FileName" 'WARN'
        return
    }
    $path = Join-Path $OutputFolder $FileName
    $Data | Export-Csv -Path $path -NoTypeInformation -Encoding UTF8 -Force
    Write-Log "  Exported $($Data.Count) record(s) -> $path" 'SUCCESS'
}

#endregion Menu & UI

#region Workload: SharePoint Online

function Collect-SharePointData {
    Write-Log ""
    Write-Log "==  SharePoint Online  ==" 'SUCCESS'

    if ([string]::IsNullOrWhiteSpace($tenantUrl)) {
        Write-Log "  Skipping SharePoint Online — `$tenantUrl not configured in the Configuration section." 'WARN'
        return
    }

    $adminUrl = $tenantUrl.TrimEnd('/') -replace 'https://([^.]+)\.sharepoint\.com.*', 'https://$1-admin.sharepoint.com'
    $tenantNamePart = (($tenantUrl -replace '^https?://', '') -split '\.')[0]
    $adminUrlsToScan = [System.Collections.Generic.List[string]]::new()
    $adminUrlsToScan.Add($adminUrl)
    Write-Log "  Connecting to SharePoint Online admin ($adminUrl)..."

    try {
        if ($PSVersionTable.PSVersion.Major -gt 5) {
            Import-Module Microsoft.Online.SharePoint.PowerShell `
                -UseWindowsPowerShell -DisableNameChecking -ErrorAction Stop
        }
        else {
            Import-Module Microsoft.Online.SharePoint.PowerShell `
                -DisableNameChecking -ErrorAction Stop
        }

        Connect-SPOService -Url $adminUrl -ErrorAction Stop
        Write-Log "  Connected to SharePoint Online admin." 'SUCCESS'
        
        # Track that we're already connected to the primary admin URL
        $activeGeoAdmins = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
        $null = $activeGeoAdmins.Add($adminUrl)

        # Detect multi-geo using SPO module native cmdlets and build GEO admin URL list.
        try {
            $geoQuotaRaw = @(Get-SPOGeoStorageQuota -AllLocations -ErrorAction Stop)
            if ($geoQuotaRaw.Count -gt 0) {
                try {
                    $geoQuotaRows = foreach ($g in $geoQuotaRaw) {
                        $row = ConvertTo-FlatCsvRow -InputObject $g
                        $row | Add-Member -NotePropertyName CollectDate -NotePropertyValue (Get-Date -Format 'yyyy-MM-dd HH:mm') -Force
                        $row
                    }
                    Export-ToCsv -Data @($geoQuotaRows) -FileName "SPO_GeoStorageQuota_$date.csv"
                    Write-Log "  SPO geo storage quota exported: $($geoQuotaRows.Count) row(s)." 'SUCCESS'
                }
                catch { Write-Log "  SPO geo storage quota export unavailable: $($_.Exception.Message)" 'WARN' }

                $discoveredAdminUrls = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
                $geoCodeMap = @{}

                foreach ($geo in $geoQuotaRaw) {
                    foreach ($prop in $geo.PSObject.Properties) {
                        if ($prop.Value -is [string] -and $prop.Value -match '^https://[^/]+-admin\.sharepoint\.') {
                            $null = $discoveredAdminUrls.Add($prop.Value.TrimEnd('/'))
                        }
                    }

                    $geoCode = $null
                    foreach ($locProp in @('GeoLocation', 'Location', 'DataLocation', 'PreferredDataLocation', 'AllowedDataLocation')) {
                        if ($geo.PSObject.Properties.Name -contains $locProp) {
                            $candidate = [string]$geo.$locProp
                            if (-not [string]::IsNullOrWhiteSpace($candidate)) {
                                $geoCode = $candidate.Trim()
                                break
                            }
                        }
                    }

                    if (-not [string]::IsNullOrWhiteSpace($geoCode)) {
                        if ($geoCode -notmatch '^(NAM|NA|DEFAULT|PRIMARY|HOME)$') {
                            $satUrl1 = "https://{0}{1}-admin.sharepoint.com" -f $tenantNamePart, $geoCode.ToLowerInvariant()
                            $satUrl2 = "https://{0}-{1}-admin.sharepoint.com" -f $tenantNamePart, $geoCode.ToLowerInvariant()
                            $null = $discoveredAdminUrls.Add($satUrl1)
                            $null = $discoveredAdminUrls.Add($satUrl2)
                            $geoCodeMap[$satUrl1] = $geoCode.ToUpperInvariant()
                            $geoCodeMap[$satUrl2] = $geoCode.ToUpperInvariant()
                        }
                        else {
                            $geoCodeMap[$adminUrl] = 'PRIMARY'
                        }
                    }
                }

                foreach ($u in $discoveredAdminUrls) {
                    if ($u -and -not ($adminUrlsToScan -contains $u)) {
                        $adminUrlsToScan.Add($u)
                    }
                }

                if ($adminUrlsToScan.Count -gt 1) {
                    Write-Log "  Multi-geo detected. GEO admin endpoints to scan: $($adminUrlsToScan.Count)" 'SUCCESS'
                }
                else {
                    Write-Log "  Multi-geo endpoints not detected; scanning primary GEO only." 'DEBUG'
                }
            }
        }
        catch {
            Write-Log "  Multi-geo discovery unavailable. Scanning primary GEO only: $($_.Exception.Message)" 'WARN'
            $geoCodeMap = @{ $adminUrl = 'PRIMARY' }
        }

        # --- SharePoint site totals (single-geo or multi-geo via admin endpoint loop, per-geo output files) ---
        Write-Log "  SharePoint site totals (Get-SPOSite)..."
        try {
            $allSiteList = [System.Collections.Generic.List[object]]::new()
            $multiGeoDetected = ($adminUrlsToScan.Count -gt 1)

            foreach ($geoAdminUrl in $adminUrlsToScan) {
                try {
                    if (-not $activeGeoAdmins.Contains($geoAdminUrl)) {
                        try { Disconnect-SPOService -ErrorAction SilentlyContinue } catch {}
                        Connect-SPOService -Url $geoAdminUrl -ErrorAction Stop
                        $null = $activeGeoAdmins.Add($geoAdminUrl)
                    }

                    $geoSites = @(Get-SPOSite -Limit All)
                    $geoSiteCount = @($geoSites).Count
                    $geoUsedMB = (@($geoSites) | Measure-Object -Property StorageUsageCurrent -Sum).Sum

                    # Determine geo suffix for filename
                    $geoSuffix = if ($geoCodeMap.ContainsKey($geoAdminUrl)) { $geoCodeMap[$geoAdminUrl] } else { 'PRIMARY' }
                    if ($adminUrlsToScan.Count -eq 1) {
                        $geoSuffix = ''
                    }
                    else {
                        $geoSuffix = "_${geoSuffix}"
                    }

                    # Export per-geo summary
                    if ($geoSiteCount -gt 0) {
                        $tenantPoolMB = if ($adminUrlsToScan.Count -eq 1) { [double]$spoTenant.StorageQuota } else { [double]$geoUsedMB * 2 }
                        $perGeoSummary = @([PSCustomObject]@{
                                TotalSites              = $geoSiteCount
                                TotalStorageUsedMB      = [math]::Round($geoUsedMB, 2)
                                TotalStorageUsedGB      = [math]::Round(($geoUsedMB / 1024), 2)
                                AvgStorageUsedPerSiteMB = if ($geoSiteCount -gt 0) { [math]::Round(($geoUsedMB / $geoSiteCount), 2) } else { 0 }
                                AvgStorageUsedPerSiteGB = if ($geoSiteCount -gt 0) { [math]::Round((($geoUsedMB / $geoSiteCount) / 1024), 4) } else { 0 }
                                PercentGeoStorageUsed   = if ($tenantPoolMB -gt 0) { [math]::Round((($geoUsedMB / $tenantPoolMB) * 100), 4) } else { 0 }
                                GeoAdminUrl             = $geoAdminUrl
                                CollectDate             = (Get-Date -Format 'yyyy-MM-dd HH:mm')
                            })
                        Export-ToCsv -Data $perGeoSummary -FileName "SPO_SiteSummary${geoSuffix}_$date.csv"
                        Write-Log "  GEO scan complete (${geoSuffix}): $geoSiteCount sites | $([math]::Round(($geoUsedMB / 1024), 2)) GB used" 'SUCCESS'
                    }
                    else {
                        Write-Log "  GEO scan (${geoSuffix}): no sites found." 'DEBUG'
                    }

                    foreach ($site in $geoSites) {
                        $allSiteList.Add($site)
                    }
                }
                catch {
                    Write-Log "  GEO scan failed for ${geoAdminUrl}: $($_.Exception.Message)" 'WARN'
                }
            }

            $allSites = @($allSiteList)
            $siteCount = @($allSites).Count
            $totalUsedMB = (@($allSites) | Measure-Object -Property StorageUsageCurrent -Sum).Sum
            Write-Log "  SPO total sites across all GEOs: $siteCount" 'SUCCESS'

            # Reconnect to the primary admin endpoint for tenant-level settings cmdlets only if multi-geo was detected.
            if ($multiGeoDetected) {
                try {
                    Disconnect-SPOService -ErrorAction SilentlyContinue
                    Connect-SPOService -Url $adminUrl -ErrorAction Stop
                }
                catch {
                    Write-Log "  Could not reconnect to primary admin endpoint before tenant settings collection: $($_.Exception.Message)" 'WARN'
                }
            }
        }
        catch {
            $allSites = @()
            $siteCount = 0
            $totalUsedMB = 0
            Write-Log "  SPO site totals failed: $($_.Exception.Message)" 'WARN'
        }

        # --- Full SPO tenant settings via Get-SPOTenant ---
        Write-Log "  SPO tenant settings (Get-SPOTenant)..."
        try {
            $spoTenant = Get-SPOTenant

            # Refresh SPO storage summary with tenant pool metrics (aligns with admin center banner)
            try {
                Write-Log "  SPO tenant storage pool summary..."

                $sitesForSummary = if ($allSites) { @($allSites) } else { @(Get-SPOSite -Limit All) }
                $siteCount = $sitesForSummary.Count
                $totalUsedMB = ($sitesForSummary | Measure-Object -Property StorageUsageCurrent -Sum).Sum
                $tenantPoolMB = [double]$spoTenant.StorageQuota
                $allocatedMB = [double]$spoTenant.StorageQuotaAllocated
                $allocatedIsReported = ($allocatedMB -gt 0)

                $spoSummary = @([PSCustomObject]@{
                        TotalSites                = $siteCount
                        TotalStorageUsedMB        = [math]::Round($totalUsedMB, 2)
                        TotalStorageUsedGB        = [math]::Round(($totalUsedMB / 1024), 2)
                        AvgStorageUsedPerSiteMB   = if ($siteCount -gt 0) { [math]::Round(($totalUsedMB / $siteCount), 2) } else { 0 }
                        TenantStoragePoolTB       = [math]::Round(($tenantPoolMB / 1024 / 1024), 2)
                        TenantStoragePoolMB       = [math]::Round($tenantPoolMB, 2)
                        AllocatedToSitesReported  = $allocatedIsReported
                        AllocatedToSitesMB        = if ($allocatedIsReported) { [math]::Round($allocatedMB, 2) } else { '' }
                        AllocatedToSitesGB        = if ($allocatedIsReported) { [math]::Round(($allocatedMB / 1024), 2) } else { '' }
                        AllocatedToSitesGBPrecise = if ($allocatedIsReported) { [math]::Round(($allocatedMB / 1024), 4) } else { '' }
                        PercentTenantPoolUsed     = if ($tenantPoolMB -gt 0) { [math]::Round((($totalUsedMB / $tenantPoolMB) * 100), 4) } else { 0 }
                        AvgStorageUsedPerSiteGB   = if ($siteCount -gt 0) { [math]::Round((($totalUsedMB / $siteCount) / 1024), 4) } else { 0 }
                        UsedOfTenantPoolDisplay   = ("{0} MB used of {1} TB" -f [math]::Round($totalUsedMB, 2), [math]::Round(($tenantPoolMB / 1024 / 1024), 2))
                        StorageAllocationModeNote = if (-not $allocatedIsReported) { 'StorageQuotaAllocated returned 0 (common with automatic storage management).' } else { '' }
                        CollectDate               = (Get-Date -Format 'yyyy-MM-dd HH:mm')
                    })

                Export-ToCsv -Data $spoSummary -FileName "SPO_SiteSummary_$date.csv"
                Write-Log ("  SPO storage: " +
                    "$($spoSummary[0].TotalStorageUsedMB) MB used of $($spoSummary[0].TenantStoragePoolTB) TB " +
                    "($($spoSummary[0].PercentTenantPoolUsed)% used)") 'SUCCESS'
            }
            catch { Write-Log "  SPO tenant storage summary unavailable: $($_.Exception.Message)" 'WARN' }

            # Tenant-level sharing policy summary (matches core controls in the SPO admin center)
            try {
                Write-Log "  Tenant sharing policy summary..."
                $sharingSummary = @([PSCustomObject]@{
                        SharePointSharingCapability                   = $spoTenant.CoreSharingCapability
                        OneDriveSharingCapability                     = $spoTenant.OneDriveSharingCapability
                        DefaultSharingLinkType                        = $spoTenant.DefaultSharingLinkType
                        DefaultLinkPermission                         = $spoTenant.DefaultLinkPermission
                        SharePointDefaultShareLinkScope               = $spoTenant.CoreDefaultShareLinkScope
                        SharePointDefaultShareLinkRole                = $spoTenant.CoreDefaultShareLinkRole
                        OneDriveDefaultShareLinkScope                 = $spoTenant.OneDriveDefaultShareLinkScope
                        OneDriveDefaultShareLinkRole                  = $spoTenant.OneDriveDefaultShareLinkRole
                        FileAnonymousLinkType                         = $spoTenant.FileAnonymousLinkType
                        FolderAnonymousLinkType                       = $spoTenant.FolderAnonymousLinkType
                        RequireAnonymousLinksExpireInDays             = $spoTenant.RequireAnonymousLinksExpireInDays
                        PreventExternalUsersFromResharing             = $spoTenant.PreventExternalUsersFromResharing
                        AllowGuestUserShareToUsersNotInSiteCollection = $spoTenant.AllowGuestUserShareToUsersNotInSiteCollection
                        AnyoneLinkTrackUsers                          = $spoTenant.AnyoneLinkTrackUsers
                        BccExternalSharingInvitations                 = $spoTenant.BccExternalSharingInvitations
                        BccExternalSharingInvitationsList             = $spoTenant.BccExternalSharingInvitationsList
                        SharingAllowedDomainList                      = $spoTenant.SharingAllowedDomainList
                        SharingBlockedDomainList                      = $spoTenant.SharingBlockedDomainList
                        SharingDomainRestrictionMode                  = $spoTenant.SharingDomainRestrictionMode
                        CollectDate                                   = (Get-Date -Format 'yyyy-MM-dd HH:mm')
                    })
                Export-ToCsv -Data $sharingSummary -FileName "SPO_SharingPolicySummary_$date.csv"
                Write-Log "  Tenant sharing policy summary collected." 'SUCCESS'
            }
            catch { Write-Log "  Tenant sharing policy summary unavailable: $($_.Exception.Message)" 'WARN' }

            # Site lifecycle management policies (inactive sites / ownership attestation)
            try {
                Write-Log "  Site lifecycle management policies..."

                $lifecycleCmdCandidates = @(
                    'Get-SPOSiteLifecycleManagementPolicy',
                    'Get-SPOSiteOwnershipPolicy',
                    'Get-SPOInactiveSitePolicy',
                    'Get-SPOSiteAttestationPolicy'
                )

                $lifecycleData = [System.Collections.Generic.List[object]]::new()
                foreach ($cmdName in $lifecycleCmdCandidates) {
                    $cmd = Get-Command -Name $cmdName -ErrorAction SilentlyContinue
                    if ($cmd) {
                        try {
                            $rows = @(& $cmdName)
                            foreach ($r in $rows) {
                                if ($null -ne $r) {
                                    $obj = ConvertTo-FlatCsvRow -InputObject $r
                                    $obj | Add-Member -NotePropertyName SourceCmdlet -NotePropertyValue $cmdName -Force
                                    $lifecycleData.Add($obj)
                                }
                            }
                        }
                        catch {
                            Write-Log "  Lifecycle cmdlet $cmdName failed: $($_.Exception.Message)" 'WARN'
                        }
                    }
                }

                if ($lifecycleData.Count -gt 0) {
                    Export-ToCsv -Data @($lifecycleData) -FileName "SPO_SiteLifecyclePolicies_$date.csv"
                    Write-Log "  Site lifecycle policies collected via lifecycle cmdlets: $($lifecycleData.Count)" 'SUCCESS'
                }
                else {
                    Write-Log "  Detailed lifecycle policies unavailable: no lifecycle cmdlets were exposed by this SPO module/session." 'WARN'
                }
            }
            catch { Write-Log "  Site lifecycle policy collection failed: $($_.Exception.Message)" 'WARN' }

            # Versioning policy from Get-SPOTenant
            $vpRow = [ordered]@{
                VersionHistoryMode           = ''
                MajorVersionLimit            = ''
                MinorVersionLimit            = ''
                ExpiresAfterDays             = ''
                DeleteOldVersionsWhenExpired = ''
                SiteDefaultStorageLimitMB    = ''
                DeletedUserODBRetentionDays  = ''
                CollectDate                  = (Get-Date -Format 'yyyy-MM-dd HH:mm')
                Source                       = 'SPO Get-SPOTenant'
            }
            try {
                $autoTrim = $spoTenant.EnableAutoExpirationVersionTrim
                if ($null -ne $autoTrim) { $vpRow['VersionHistoryMode'] = if ($autoTrim) { 'Automatic' } else { 'Manual' } }
                if ($null -ne $spoTenant.MajorVersionLimit) { $vpRow['MajorVersionLimit'] = $spoTenant.MajorVersionLimit }
                if ($null -ne $spoTenant.MinorVersionLimit) { $vpRow['MinorVersionLimit'] = $spoTenant.MinorVersionLimit }
                $expDays = $spoTenant.ExpireVersionsAfterDays
                if ($null -ne $expDays) { $vpRow['ExpiresAfterDays'] = if ([int]$expDays -eq 0) { 'Never' } else { $expDays } }
                if ($null -ne $spoTenant.DeleteOldVersionsWhenExpired) { $vpRow['DeleteOldVersionsWhenExpired'] = $spoTenant.DeleteOldVersionsWhenExpired }
                if ($null -ne $spoTenant.StorageQuotaAllocated) { $vpRow['SiteDefaultStorageLimitMB'] = $spoTenant.StorageQuotaAllocated }
                if ($null -ne $spoTenant.OrphanedPersonalSitesRetentionPeriod) { $vpRow['DeletedUserODBRetentionDays'] = $spoTenant.OrphanedPersonalSitesRetentionPeriod }
                Export-ToCsv -Data @([PSCustomObject]$vpRow) -FileName "SPO_VersioningPolicy_$date.csv"
                Write-Log "  Versioning policy collected from Get-SPOTenant." 'SUCCESS'
            }
            catch { Write-Log "  Could not collect versioning policy: $($_.Exception.Message)" 'WARN' }

            # ---- Descriptions for all known Get-SPOTenant properties ----
            $spoDescriptions = @{
                AllowCommentsTextOnEmailEnabled                                  = "Determines whether email notifications about documents include the actual text of comments made in the document."
                AllowDownloadingNonWebViewableFiles                              = "Allows users to download files that cannot be previewed in a browser when access from unmanaged devices is restricted."
                AllowEditing                                                     = "Controls whether users can edit Office files in the browser when the unmanaged-devices policy is set to allow limited web-only access."
                AllowGuestUserShareToUsersNotInSiteCollection                    = "Determines whether guest users can share files and folders with people who are not members of the site collection."
                AllowLimitedAccessOnUnmanagedDevices                             = "Allows limited, web-only access from unmanaged devices that are not compliant or domain-joined."
                AllowOverrideForBlockUserInfoVisibility                          = "Allows site and group admins to override the tenant-level setting that controls visibility of user profile information."
                AllowSelectSecurityGroups                                        = "Specifies security groups whose members are permitted to perform sharing actions that may otherwise be restricted."
                AllowSelectiveSideLoading                                        = "Allows selective sideloading of SharePoint Framework (SPFx) solutions in specific site collections."
                AnyoneLinkTrackUsers                                             = "When enabled, SharePoint logs the identity of users who access content via 'Anyone' (anonymous) sharing links."
                ApplyAppEnforcedRestrictionsToAdHocRecipients                    = "Applies app-enforced conditional access restrictions to recipients of ad-hoc (direct) sharing invitations."
                AuthContextResilienceMode                                        = "Configures how SharePoint handles authentication context resilience when Azure AD conditional access policies are applied."
                BccExternalSharingInvitations                                    = "When enabled, all external sharing invitations send a BCC copy to the addresses in BccExternalSharingInvitationsList."
                BccExternalSharingInvitationsList                                = "Comma-separated list of email addresses that receive BCC copies of all external sharing invitations."
                BlockAccessOnUnmanagedDevices                                    = "Blocks all access to SharePoint Online and OneDrive for Business from devices that are not compliant or domain-joined."
                BlockDownloadFileTypeIds                                         = "File type IDs whose files cannot be downloaded when accessed from an unmanaged or restricted device."
                BlockDownloadLinksFileType                                       = "Category of files (WebPreviewableFiles or ServerRenderedFilesOnly) for which download links are blocked in limited-access scenarios."
                BlockDownloadOfAllFilesForGuests                                 = "Prevents guest (external) users from downloading any files from SharePoint Online and OneDrive for Business."
                BlockDownloadOfAllFilesOnUnmanagedDevices                        = "Prevents users on unmanaged devices from downloading any SharePoint Online or OneDrive for Business files."
                BlockDownloadOfViewableFilesForGuests                            = "Prevents guest users from downloading files that can be previewed in the browser (e.g., Office documents, PDFs)."
                BlockDownloadOfViewableFilesOnUnmanagedDevices                   = "Prevents users on unmanaged devices from downloading files that can be previewed in the browser."
                BlockMacSync                                                     = "Blocks the OneDrive for Business Mac sync client from syncing content with the tenant."
                BlockSendLabelMismatchEmail                                      = "Suppresses the email notification sent to users when a document's sensitivity label does not match the site's label policy."
                CommentsOnFilesDisabled                                          = "Disables the ability for users to add comments to SharePoint files across the entire tenant."
                CommentsOnListItemsDisabled                                      = "Disables the ability for users to add comments to SharePoint list items across the tenant."
                CommentsOnSitePagesDisabled                                      = "Disables the comments section on modern SharePoint site pages (news, pages) across the tenant."
                CompatibilityRange                                               = "Defines the minimum and maximum compatibility levels allowed for new site collections created in the tenant."
                ConditionalAccessPolicy                                          = "Conditional access policy applied to unmanaged devices: AllowFullAccess, AllowLimitedAccess, or BlockAccess."
                ConditionalAccessPolicyErrorHelpLink                             = "Custom URL shown as a help link on the error page displayed to users blocked by conditional access policies."
                ContentTypeSyncSiteTemplatesList                                 = "Site templates for which content type syndication (push) from the Content Type Gallery is enabled."
                CoreBlockGuestsAsSiteAdmin                                       = "Prevents guest (external) users from being assigned as site collection administrators on SharePoint sites."
                CoreDefaultLinkToExistingAccess                                  = "When enabled, the default sharing link for SharePoint (non-OneDrive) sites grants existing permissions rather than new access."
                CoreDefaultShareLinkRole                                         = "Default permission level (View or Edit) for sharing links created on SharePoint (non-OneDrive) sites."
                CoreDefaultShareLinkScope                                        = "Default audience scope (Anyone, Organization, SpecificPeople) for sharing links on SharePoint (non-OneDrive) sites."
                CoreRequestFilesLinkEnabled                                      = "Enables the 'Request files' feature on SharePoint document libraries, letting owners generate a link for others to upload files."
                CoreRequestFilesLinkExpirationInDays                             = "Number of days after which 'Request files' links on SharePoint sites automatically expire. 0 means no expiration."
                CoreSharingCapability                                            = "Sharing capability for SharePoint sites independent of OneDrive: Disabled, ExistingExternalUserSharingOnly, ExternalUserSharingOnly, or ExternalUserAndGuestSharing."
                CustomizedExternalSharingServiceUrl                              = "Custom URL used to replace the default SharePoint external sharing service endpoint."
                DefaultContentCenterSite                                         = "URL of the default SharePoint Syntex / Purview content center site for the tenant."
                DefaultLinkPermission                                            = "Default permission level (View or Edit) for all new sharing links created across the tenant."
                DefaultODBMode                                                   = "Default sharing mode for OneDrive for Business sites: Direct, Private, or Off."
                DefaultSharingLinkType                                           = "Default type of sharing link: None, Direct (specific people), Internal (org-wide), or AnonymousAccess (Anyone)."
                DenySelectSecurityGroups                                         = "Security groups whose members are explicitly prevented from sharing content externally."
                DisableAddToOneDrive                                             = "Disables the 'Add shortcut to OneDrive' feature so users cannot create OneDrive shortcuts to shared libraries."
                DisableBackToClassic                                             = "Removes the 'Back to classic SharePoint' link from the modern UI, preventing users from switching to the classic experience."
                DisableCustomAppAuthentication                                   = "Disables SharePoint app-only authentication using legacy SharePoint App Principals (ACS-based apps)."
                DisableDocumentLibraryDefaultLabeling                            = "Prevents default sensitivity labels from being automatically applied to documents in SharePoint document libraries."
                DisabledModernListTemplateIds                                    = "Collection of GUID-based template IDs that are disabled and hidden from users when creating modern SharePoint lists."
                DisabledWebPartIds                                               = "Collection of GUIDs representing web parts that are blocked from being added to pages across the tenant."
                DisableListSync                                                  = "Prevents SharePoint lists from being synced to the OneDrive desktop sync client, blocking offline list access."
                DisableOutlookPSTVersionTrimming                                 = "Disables automatic PST file version trimming within the Outlook integration with SharePoint Online."
                DisablePersonalListCreation                                      = "Prevents users from creating personal SharePoint lists that are visible only to themselves."
                DisableReportProblemDialog                                       = "Hides the 'Report a problem' feedback dialog from the SharePoint Online user interface."
                DisableSpacesActivities                                          = "Disables activity tracking and the activity feed within SharePoint Spaces (3D / mixed-reality content)."
                DisallowInfectedFileDownload                                     = "Blocks users from downloading files detected as containing malware by virus scanning."
                DisplayNamesOfFileViewers                                        = "Shows the names and profile pictures of users who have recently viewed a file."
                DisplayNamesOfFileViewersInSpo                                   = "Controls whether viewer names are shown in SharePoint Online file cards specifically."
                DisplayStartASiteOption                                          = "Controls whether the 'Create site' option is visible to users on the SharePoint home page."
                EmailAttestationEnabled                                          = "Enables email-based attestation, requiring external users to periodically verify their identity via email."
                EmailAttestationReAuthDays                                       = "Number of days an external user's email attestation remains valid before re-verification is required."
                EmailAttestationRequired                                         = "Requires all external (guest) users to complete email attestation before accessing shared content."
                EnableAIPIntegration                                             = "Enables integration between SharePoint Online and Azure Information Protection (Microsoft Purview) for automatic sensitivity labeling."
                EnableAutoNewsDigest                                             = "Enables automatic news digest emails, sending users periodic summaries of SharePoint news they may have missed."
                EnableGuestSignInAcceleration                                    = "Enables home realm discovery acceleration for guest sign-ins by redirecting guests directly to their identity provider."
                EnableMinimumVersionRequirement                                  = "Enforces a minimum file version requirement, ensuring a base number of versions is always retained before cleanup."
                EnableRestrictedAccessControl                                    = "Enables restricted access control (site-level access policies) that can be applied to individual SharePoint sites."
                EnableSensitivityLabelForPDF                                     = "Enables sensitivity labels to be applied to and read from PDF files in SharePoint Online and OneDrive."
                EnableVersionExpirationSetting                                   = "Enables the version expiration settings UI and functionality in SharePoint Online document libraries."
                ExcludedFileExtensionsForSyncClient                              = "List of file extensions (e.g., .exe, .ps1) excluded from syncing by the OneDrive desktop sync client."
                ExternalServicesEnabled                                          = "Enables or disables external services integration in SharePoint, such as external workflows or BCS connections."
                ExternalUserExpirationRequired                                   = "Requires external user accounts to automatically expire after a specified number of days."
                ExternalUserExpireInDays                                         = "Number of days after which external user access automatically expires when ExternalUserExpirationRequired is enabled."
                FileAnonymousLinkType                                            = "Default permission type for anonymous 'Anyone' file sharing links: View or Edit."
                FilePickerExternalImageSearchEnabled                             = "Enables users to search for and insert external (internet) images via the file picker in SharePoint Online."
                FolderAnonymousLinkType                                          = "Default permission type for anonymous 'Anyone' folder sharing links: View or Edit."
                GuestSharingGroupAllowListInTenant                               = "Comma-separated list of security groups whose members are specifically allowed to share content with guests."
                HideDefaultThemes                                                = "Hides the built-in SharePoint color themes from the 'Change the look' panel, enforcing custom branding."
                HideSyncButtonOnDocLib                                           = "Hides the 'Sync' button on SharePoint document library toolbars to discourage or prevent desktop sync."
                HideSyncButtonOnODB                                              = "Hides the 'Sync' button on OneDrive for Business to discourage or prevent desktop sync."
                IBImplicitGroupBased                                             = "Enables implicit group-based information barrier mode, automatically applying barriers based on group membership."
                IBMode                                                           = "Information barriers mode for the tenant: Off, Open, Explicit, OwnerModerated, or Implicit."
                ImageTaggingOption                                               = "Controls AI-powered automatic image tagging in SharePoint photo libraries: Disabled, BasicImageTaggingDisabled, or AllImageTaggingDisabled."
                IncludeAtAGlanceInShareEmails                                    = "Includes an 'At a glance' preview showing recent activity in email notifications sent when content is shared."
                IPAddressAllowList                                               = "Comma-separated list of IP address ranges (CIDR) from which access to SharePoint Online is permitted when enforcement is on."
                IPAddressEnforcement                                             = "Enables IP-address-based access control, restricting SharePoint Online access to the ranges in IPAddressAllowList."
                IPAddressWACTokenLifetime                                        = "Token lifetime (in minutes) for Office Online Server (WAC) sessions when IP address enforcement is enabled."
                IsCollabMeetingNotesFluidEnabled                                 = "Enables Fluid-based collaborative meeting notes (Loop components) within SharePoint and Teams meetings."
                IsEnableAppAuthPopUpEnabled                                      = "Enables a pop-up window flow for SharePoint app authentication when iframe-based auth is blocked by the browser."
                IsFluidEnabled                                                   = "Enables Microsoft Loop (Fluid Framework) components in SharePoint Online and other Microsoft 365 surfaces."
                IsHubSitesMultiGeoFlightEnabled                                  = "Enables multi-geo support for hub sites, allowing hub sites to span multiple geographic locations."
                IsLoopEnabled                                                    = "Enables Microsoft Loop workspaces and components to be created and shared within the tenant."
                IsMnAFlightEnabled                                               = "Indicates whether the Meetings and Annotations (M&A) feature flight is enabled for the tenant."
                IsMultipleHomeSitesFlightEnabled                                 = "Indicates whether the multiple Viva Connections home sites feature flight is enabled for the tenant."
                IsVivaHomeFlightEnabled                                          = "Indicates whether the Viva Connections home experience flight is enabled for the tenant."
                IsWBFlightEnabled                                                = "Indicates whether the Microsoft Whiteboard integration flight is enabled for the tenant."
                LabelMismatchEmailHelpLink                                       = "Custom help URL included in emails sent when a document's sensitivity label conflicts with the library's required label."
                LegacyAuthProtocolsEnabled                                       = "Allows legacy authentication protocols (e.g., Basic, Digest) to connect to SharePoint Online."
                LimitedAccessFileType                                            = "File types accessible in browser-only (limited access) mode: OfficeOnlineFilesOnly, WebPreviewableFiles, or OtherFiles."
                MachineLearningCaptureEnabled                                    = "Enables machine learning-based document capture and extraction features (SharePoint Syntex / Microsoft Purview)."
                MajorVersionLimit                                                = "Maximum number of major versions retained per document when manual versioning trim is used."
                MarkNewFilesSensitiveByDefault                                   = "Automatically marks newly uploaded files as sensitive until a sensitivity label is explicitly applied."
                MediaTranscription                                               = "Enables or disables automatic transcription of video and audio files uploaded to SharePoint Online and OneDrive."
                MinCompatibilityLevel                                            = "Minimum site collection compatibility level permitted in the tenant (used together with MaxCompatibilityLevel)."
                MinimumBytesForMac                                               = "Minimum file size (in bytes) that triggers file processing by the OneDrive Mac sync client."
                NoAccessRedirectUrl                                              = "Custom URL to redirect users to when they are denied access to a SharePoint Online site."
                NotificationsInOneDriveForBusinessEnabled                        = "Enables in-app and push activity notifications for OneDrive for Business events (file shared, comment added, etc.)."
                NotificationsInSharePointEnabled                                 = "Enables in-app activity notifications within SharePoint Online for comments, shares, and @mentions."
                NotifyOwnersWhenInvitationsAccepted                              = "Sends email notifications to site owners when external (guest) sharing invitations are accepted."
                NotifyOwnersWhenItemsReshared                                    = "Sends email notifications to content owners when another user reshares their content with additional people."
                ODBAccessRequests                                                = "Controls access request behavior in OneDrive for Business: On, Off, or Unspecified."
                ODBMembersCanShare                                               = "Controls whether members of a OneDrive owner's site can share content with others: On, Off, or Unspecified."
                OfficeClientADALDisabled                                         = "Disables modern authentication (ADAL/OAuth) for Office desktop client applications, forcing legacy authentication."
                OneDriveBlockGuestsAsSiteAdmin                                   = "Prevents guest users from being added as site collection administrators to OneDrive for Business sites."
                OneDriveDefaultLinkToExistingAccess                              = "When enabled, the default sharing link for OneDrive grants existing permissions rather than creating new access."
                OneDriveDefaultShareLinkRole                                     = "Default permission role (View or Edit) for sharing links created in OneDrive for Business."
                OneDriveDefaultShareLinkScope                                    = "Default audience scope (Anyone, Organization, SpecificPeople) for new sharing links in OneDrive for Business."
                OneDriveForGuestsEnabled                                         = "Enables provisioning of OneDrive for Business sites for guest (external) users added to the tenant."
                OneDriveLoopDefaultSharingLinkRole                               = "Default sharing permission role for Microsoft Loop components stored in OneDrive."
                OneDriveLoopDefaultSharingLinkScope                              = "Default sharing scope (audience) for Microsoft Loop components stored in OneDrive."
                OneDriveLoopSharingCapability                                    = "Sharing capability for Microsoft Loop content stored in OneDrive."
                OneDriveRequestFilesLinkEnabled                                  = "Enables the 'Request files' feature on OneDrive, letting users generate an upload link to receive files from others."
                OneDriveRequestFilesLinkExpirationInDays                         = "Number of days after which 'Request files' upload links in OneDrive automatically expire. 0 means no expiration."
                OneDriveSharingCapability                                        = "Sharing capability specifically for OneDrive for Business sites, independent of SharePoint site sharing settings."
                OneDriveStorageQuota                                             = "Default storage quota (in MB) allocated to each user's OneDrive for Business."
                OptOutOfGrooveBlock                                              = "Opts the tenant out of blocking the legacy Groove (OneDrive for Business) sync client."
                OptOutOfGrooveSoftBlock                                          = "Opts the tenant out of soft-blocking (warning) the legacy Groove sync client."
                OrphanedPersonalSitesRetentionPeriod                             = "Days to retain a deleted user's OneDrive for Business before permanent deletion (default is 30 days)."
                OwnerAnonymousNotification                                       = "Notifies content owners via email when an anonymous ('Anyone') link to their content is accessed."
                PermissiveBrowserFileHandlingOverride                            = "Overrides the default restrictive browser file handling, allowing certain file types to open directly in the browser."
                PreventExternalUsersFromResharing                                = "Prevents external users from resharing files and folders they have access to but do not own."
                ProvisionSharedWithEveryoneFolder                                = "Automatically creates a 'Shared with Everyone' folder in each user's OneDrive when it is first provisioned."
                PublicCdnAllowedFileTypes                                        = "Comma-separated file extensions (e.g., CSS, EOT, GIF, PNG) permitted to be hosted in the Office 365 public CDN."
                PublicCdnEnabled                                                 = "Enables the Office 365 public CDN, allowing tenant assets to be served from Microsoft's global CDN network."
                PublicCdnOrigins                                                 = "Configured origins (SharePoint library paths) published to the Office 365 public CDN."
                ReduceTempTokenLifetimeEnabled                                   = "Enables the reduction of temporary access token lifetimes for improved security in time-sensitive scenarios."
                ReduceTempTokenLifetimeValue                                     = "Custom duration (in minutes) for reduced-lifetime temporary tokens when ReduceTempTokenLifetimeEnabled is true."
                RequireAcceptingAccountMatchInvitedAccount                       = "Requires that a guest user accept a sharing invitation using the same email address to which it was sent."
                RequireAnonymousLinksExpireInDays                                = "Requires all 'Anyone' (anonymous) sharing links to expire after the specified number of days."
                ResourceQuota                                                    = "Total resource quota (sandboxed solution execution points) available to the SharePoint Online tenant."
                ResourceQuotaAllocated                                           = "Amount of resource quota already consumed / allocated across all site collections in the tenant."
                RootSiteUrl                                                      = "URL of the root (top-level) SharePoint site collection for the tenant (e.g., https://contoso.sharepoint.com)."
                SearchResolveExactEmailOrUPNResults                              = "When enabled, an exact email address or UPN match in the people picker is always returned first."
                SelfServiceSiteCreationDisabled                                  = "Disables the ability for end users to create new SharePoint sites (team sites, communication sites) via self-service."
                SetDefaultLinkToExistingAccess                                   = "Sets the default sharing link type to 'People with existing access' across all content in the tenant."
                SharedWithEverybodyStatus                                        = "Indicates the current status of the 'Shared with Everyone' folder feature in OneDrive provisioning."
                ShowAllUsersClaim                                                = "Shows the 'All Users (windows)' claim in the SharePoint people picker, allowing it to be used in permissions."
                ShowEveryoneClaim                                                = "Shows the 'Everyone' claim in the SharePoint people picker for use in permissions and sharing."
                ShowEveryoneExceptExternalUsersClaim                             = "Shows the 'Everyone except external users' claim in the SharePoint people picker for use in permissions."
                ShowNGSCDialogForSyncOnODB                                       = "Shows the next-generation sync client (OneDrive) dialog when users click Sync on a OneDrive for Business library."
                ShowPeoplePickerGroupSuggestionsForGuestUsers                    = "Shows security group name suggestions in the people picker when searching for guest users."
                ShowPeoplePickerSuggestionsForGuestUsers                         = "Shows user name and email suggestions for guest users in the people picker (requires Azure AD lookup)."
                SignInAccelerationDomain                                         = "ADFS or identity provider domain pre-populated for sign-in acceleration (home realm discovery)."
                SocialBarOnSitePagesDisabled                                     = "Disables the social bar (Like, Follow, Share, Views) displayed at the top of modern SharePoint site pages."
                SpecialCharactersStateInFileFolderNames                          = "Controls whether special characters (# %) are allowed in SharePoint file and folder names: Allowed or Disallowed."
                StartASiteFormUrl                                                = "Custom URL to redirect users to when they click 'Create site', replacing the default SharePoint site creation form."
                StorageQuota                                                     = "Total storage quota (in MB) provisioned for the entire SharePoint Online tenant."
                StorageQuotaAllocated                                            = "Amount of storage quota (in MB) already allocated across all site collections in the tenant."
                StreamLaunchConfig                                               = "Microsoft Stream migration / launch configuration: 0 = Classic Stream, 1 = Migrating, 2 = New Stream."
                StreamLaunchConfigLastUpdated                                    = "Timestamp of when the StreamLaunchConfig setting was last modified."
                StreamLaunchConfigUpdateInProgress                               = "Indicates whether a Stream launch configuration change is currently being processed."
                SyncAadB2BManagementPolicy                                       = "Enables synchronization of Azure AD B2B collaboration management policies with SharePoint's sharing framework."
                TaxonomyTaggingEnabled                                           = "Enables automatic taxonomy (managed metadata) tagging of documents using AI in SharePoint Syntex."
                TlsTokenBindingPolicyValue                                       = "TLS token binding enforcement level for enhanced session security in SharePoint Online."
                UseFindPeopleInPeoplePicker                                      = "Uses the Azure AD 'Find People' search service in the people picker instead of relying solely on the local GAL."
                UsePersistentCookiesForExplorerView                              = "Uses persistent (long-lived) cookies to maintain authentication when users open libraries with 'View in File Explorer'."
                ViewersCanCommentOnMediaDisabled                                 = "Disables the ability for file viewers to add comments on video and audio media files in SharePoint and OneDrive."
                ViewInFileExplorerEnabled                                        = "Enables the 'View in File Explorer' option in modern SharePoint document libraries for Windows Explorer-style file access."
                # --- Additional properties from Get-SPOTenant ---
                PSComputerName                                                   = "PowerShell remoting metadata: name of the computer the command ran on (typically 'localhost')."
                PSShowComputerName                                               = "PowerShell remoting metadata: indicates whether the computer name is shown in output (True/False)."
                RunspaceId                                                       = "Identifier of the PowerShell runspace/session used to execute the command."
                BonusStorageQuotaMB                                              = "Additional storage quota (MB) allocated to the tenant beyond the base subscription quota."
                ArchiveRedirectUrl                                               = "URL used to redirect users when they access content that has been archived."
                SharingCapability                                                = "Overall tenant-level external sharing setting: Disabled, ExistingExternalUserSharingOnly, ExternalUserSharingOnly, or ExternalUserAndGuestSharing."
                IsSharePointAddInsDisabled                                       = "Disables SharePoint Add-ins (legacy add-in model) across the entire tenant."
                IsSharePointAddInsBlocked                                        = "Blocks SharePoint Add-ins from running even if they were previously installed."
                DisableSharePointStoreAccess                                     = "Prevents users from accessing the SharePoint Store to acquire new add-ins."
                SiteOwnerManageLegacyServicePrincipalEnabled                     = "Allows site collection owners to manage legacy ACS-based service principals for their sites."
                AllowEveryoneExceptExternalUsersClaimInPrivateSite               = "Allows the 'Everyone except external users' claim to be used in permissions on private (non-public) sites."
                SearchResolveExactEmailOrUPN                                     = "Requires an exact email address or UPN match when resolving users in the people picker search."
                SharingAllowedDomainList                                         = "Comma-separated list of domains explicitly allowed for external sharing invitations."
                SharingBlockedDomainList                                         = "Comma-separated list of domains explicitly blocked from receiving external sharing invitations."
                SharingDomainRestrictionMode                                     = "Mode for domain-based sharing restrictions: None, AllowList, or BlockList."
                EnableTenantRestrictionsInsights                                 = "Enables insights and monitoring for tenant restriction policies applied to the tenant."
                EnablePromotedFileHandlers                                       = "Enables promoted (custom) file handlers that override default file-open behavior for specific file types."
                AppOnlyBypassPeoplePickerPolicies                                = "Allows app-only service principals to bypass people picker restriction policies."
                EnableDiscoverableByOrganizationForVideos                        = "Makes video files stored in SharePoint/OneDrive discoverable by all users in the organization."
                SiteOwnersCanAccessMissingContent                                = "Allows site collection owners to access and recover content that is otherwise inaccessible."
                AllowAppsBypassOfUnmanagedDevicePolicy                           = "Allows specific applications to bypass the unmanaged device conditional access policy."
                DisabledAdaptiveCardExtensionIds                                 = "Collection of GUIDs for Adaptive Card Extensions (Viva Connections) that are disabled tenant-wide."
                RestrictedAccessControlforSitesErrorHelpLink                     = "Custom help URL shown to users blocked by restricted access control policies on SharePoint sites."
                EnableAzureADB2BIntegration                                      = "Enables Azure AD B2B collaboration integration so external guest invitations use the AAD B2B framework."
                RestrictedAccessControlForOneDriveErrorHelpLink                  = "Custom help URL shown to users blocked by restricted access control policies on OneDrive."
                ResyncContentSecurityPolicyConfigurationEntries                  = "Triggers a resync of Content Security Policy (CSP) configuration entries across SharePoint."
                ContentSecurityPolicyEnforcement                                 = "Enables enforcement of Content Security Policy headers on SharePoint Online pages."
                DelayContentSecurityPolicyEnforcement                            = "Delays enforcement of Content Security Policy to allow time for compatibility remediation."
                OneDriveOrganizationSharingLinkMaxExpirationInDays               = "Maximum number of days before organization-scoped sharing links in OneDrive automatically expire."
                OneDriveOrganizationSharingLinkRecommendedExpirationInDays       = "Recommended expiration duration (days) for organization-scoped sharing links in OneDrive."
                CoreLoopDefaultSharingLinkScope                                  = "Default sharing scope (Anyone, Organization, SpecificPeople) for Microsoft Loop components in SharePoint."
                CoreLoopDefaultSharingLinkRole                                   = "Default sharing permission role (View or Edit) for Microsoft Loop components in SharePoint."
                BlockAppAccessWithAuthenticationContext                          = "Blocks application access when a specific Azure AD authentication context policy is required but not satisfied."
                CoreOrganizationSharingLinkMaxExpirationInDays                   = "Maximum number of days before organization-scoped sharing links on SharePoint sites automatically expire."
                CoreOrganizationSharingLinkRecommendedExpirationInDays           = "Recommended expiration duration (days) for organization-scoped sharing links on SharePoint sites."
                AllowAnonymousMeetingParticipantsToAccessWhiteboards             = "Allows anonymous (unauthenticated) meeting participants to access Microsoft Whiteboards shared in meetings."
                Workflows2013State                                               = "Current state of SharePoint 2013 workflows in the tenant: Disabled, Allowed, or Blocked."
                IsWBFluidEnabled                                                 = "Enables the Fluid Framework for Microsoft Whiteboard, allowing real-time collaborative whiteboard content."
                ExtendPermissionsToUnprotectedFiles                              = "Extends sensitivity-label-based permissions to files that do not have a sensitivity label applied."
                EnableSensitivityLabelForOneNote                                 = "Enables sensitivity labels to be applied to and enforced on OneNote notebooks in SharePoint and OneDrive."
                EnableSensitivityLabelForVideoFiles                              = "Enables sensitivity labels to be applied to and enforced on video files in SharePoint and OneDrive."
                DisableAddShortcutsToOneDrive                                    = "Disables the 'Add shortcut to OneDrive' menu option across SharePoint, preventing users from creating OneDrive shortcuts to shared libraries."
                Workflow2010Disabled                                             = "Indicates whether SharePoint 2010 workflows are disabled for the tenant."
                StopNew2010Workflows                                             = "Prevents creation of new SharePoint 2010 workflows while allowing existing ones to continue running."
                AllowSharingOutsideRestrictedAccessControlGroups                 = "Allows content sharing with users outside the security groups configured in Restricted Access Control policies."
                StopNew2013Workflows                                             = "Prevents creation of new SharePoint 2013 workflows while allowing existing ones to continue running."
                StopAlerts                                                       = "Disables SharePoint alert emails and notifications for all users across the tenant."
                BlockUserInfoVisibility                                          = "Restricts visibility of user profile information (name, photo, presence) to users within the organization."
                BlockUserInfoVisibilityInOneDrive                                = "Restricts visibility of user profile information specifically within OneDrive for Business contexts."
                BlockUserInfoVisibilityInSharePoint                              = "Restricts visibility of user profile information specifically within SharePoint Online contexts."
                InformationBarriersSuspension                                    = "Temporarily suspends Information Barriers policy enforcement across the tenant."
                AppBypassInformationBarriers                                     = "Allows specific applications to bypass Information Barriers policies when accessing SharePoint content."
                DefaultOneDriveInformationBarrierMode                            = "Default Information Barriers enforcement mode for OneDrive for Business sites: Open, Explicit, or Implicit."
                AppAccessInformationBarriersAllowList                            = "List of application IDs permitted to access content across Information Barriers segments."
                AllOrganizationSecurityGroupId                                   = "Object ID of the security group representing all users in the organization, used in sharing and IB policies."
                DisableSpacesActivation                                          = "Disables SharePoint Spaces (immersive 3D/mixed-reality experiences) from being activated in the tenant."
                DisableVivaConnectionsAnalytics                                  = "Disables analytics data collection and reporting for the Viva Connections experience."
                HideSyncButtonOnTeamSite                                         = "Hides the 'Sync' button on SharePoint team site document libraries."
                EnableAutoExpirationVersionTrim                                  = "Enables automatic trimming of old file versions based on expiration age rather than a fixed count."
                EnableMediaReactions                                             = "Enables emoji reactions on media files (video, audio) stored in SharePoint and OneDrive."
                ExpireVersionsAfterDays                                          = "Number of days after which old document versions are automatically deleted. 0 means versions never expire."
                VersionPolicyFileTypeOverride                                    = "Overrides the default version policy for specific file types, applying different retention rules."
                MediaTranscriptionAutomaticFeatures                              = "Enables automatic feature processing (chapters, highlights) on top of base media transcription."
                ShowOpenInDesktopOptionForSyncedFiles                            = "Shows an 'Open in desktop app' option for files synced via the OneDrive sync client."
                ShowPeoplePickerGroupSuggestionsForIB                            = "Shows Information Barriers-compatible group suggestions in the SharePoint people picker."
                DelegateRestrictedContentDiscoverabilityManagement               = "Allows delegated administrators to manage restricted content discoverability settings on behalf of site owners."
                DelegateRestrictedAccessControlManagement                        = "Allows delegated administrators to manage Restricted Access Control policies on behalf of site owners."
                BlockDownloadFileTypePolicy                                      = "Enables the policy that blocks downloads of specific file types from SharePoint and OneDrive."
                ExcludedBlockDownloadGroupIds                                    = "Security group IDs whose members are excluded from file-type download blocking policies."
                LegacyBrowserAuthProtocolsEnabled                                = "Enables legacy authentication protocols for older browser clients connecting to SharePoint Online."
                AllowLegacyBrowserAuthProtocolsEnabledSetting                    = "Controls whether the LegacyBrowserAuthProtocolsEnabled setting can be modified by tenant administrators."
                AllowLegacyAuthProtocolsEnabledSetting                           = "Controls whether the LegacyAuthProtocolsEnabled setting can be toggled by tenant administrators."
                RecycleBinRetentionPeriod                                        = "Number of days deleted items are retained in the SharePoint Recycle Bin before permanent deletion."
                IsDataAccessInCardDesignerEnabled                                = "Enables data source connections and data access features within the SharePoint Card Designer."
                MassDeleteNotificationDisabled                                   = "Disables email notifications sent to site owners when a large number of items are deleted."
                MassDeleteNotificationDisabledForODB                             = "Disables mass-delete email notifications specifically for OneDrive for Business."
                MassDeleteNotificationDisabledForSPO                             = "Disables mass-delete email notifications specifically for SharePoint Online sites."
                BusinessConnectivityServiceDisabled                              = "Disables the Business Connectivity Services (BCS) feature that connects SharePoint to external data sources."
                AllowSensitivityLabelOnRecords                                   = "Allows sensitivity labels to be applied to items that have been declared as records in SharePoint."
                DelayDenyAddAndCustomizePagesEnforcement                         = "Delays enforcement of the DenyAddAndCustomizePages setting, giving time for custom page migration."
                DelayDenyAddAndCustomizePagesEnforcementOnClassicPublishingSites = "Delays DenyAddAndCustomizePages enforcement specifically for classic SharePoint publishing sites."
                AllowClassicPublishingSiteCreation                               = "Allows creation of classic SharePoint publishing sites (WCM publishing feature) in the tenant."
                EsignatureEnabled                                                = "Enables the Microsoft eSignature service for requesting and applying electronic signatures to documents."
                ESignatureSiteList                                               = "List of SharePoint site URLs where the eSignature feature is enabled."
                ESignatureThirdPartyProviderInfoList                             = "Configuration details for third-party eSignature providers integrated with SharePoint."
                ESignatureAppList                                                = "List of application IDs authorized to use the SharePoint eSignature API."
                WhoCanShareAnonymousAllowList                                    = "Security group IDs whose members are permitted to create anonymous ('Anyone') sharing links."
                WhoCanShareAuthenticatedGuestAllowList                           = "Security group IDs whose members are permitted to share content with authenticated external guest users."
                DocumentUnderstandingModelScope                                  = "Scope for SharePoint Syntex document understanding models: Tenant, SelectedSites, or Off."
                DocumentUnderstandingModelSelectedSitesList                      = "List of site URLs where SharePoint Syntex document understanding models are enabled."
                AIBuilderModelScope                                              = "Scope for AI Builder models in SharePoint Syntex: Tenant, SelectedSites, or Off."
                AIBuilderModelSelectedSitesList                                  = "List of site URLs where AI Builder model processing is enabled."
                AIBuilderSelectedSitesIncludesContentCenters                     = "Indicates whether content center sites are included in the AI Builder selected sites list."
                PrebuiltModelScope                                               = "Scope for SharePoint Syntex prebuilt (out-of-box) AI models: Tenant, SelectedSites, or Off."
                PrebuiltModelSelectedSitesList                                   = "List of site URLs where prebuilt SharePoint Syntex models are enabled."
                DocumentTranslationScope                                         = "Scope for SharePoint Syntex automatic document translation: Tenant, SelectedSites, or Off."
                DocumentTranslationSelectedSitesList                             = "List of site URLs where automatic document translation is enabled."
                AutofillColumnsScope                                             = "Scope for SharePoint Syntex autofill columns (AI-powered metadata extraction): Tenant, SelectedSites, or Off."
                AutofillColumnsSelectedSitesList                                 = "List of site URLs where autofill columns (AI metadata extraction) are enabled."
                KnowledgeAgentScope                                              = "Scope for knowledge agents (Microsoft 365 Copilot / SharePoint agents): Tenant, SelectedSites, or Off."
                KnowledgeAgentSelectedSitesList                                  = "List of site URLs whose content is indexed and available to knowledge agents."
                OpticalCharacterRecognitionScope                                 = "Scope for SharePoint Syntex OCR (optical character recognition) processing: Tenant, SelectedSites, or Off."
                OpticalCharacterRecognitionSelectedSitesList                     = "List of site URLs where OCR processing is enabled for image and scanned document files."
                AllowWebPropertyBagUpdateWhenDenyAddAndCustomizePagesIsEnabled   = "Allows updates to the SharePoint web property bag even when DenyAddAndCustomizePages is enforced."
                AllowSelectSecurityGroupsInSPSitesList                           = "Security groups explicitly allowed to apply site-level sharing restrictions on SharePoint sites."
                DenySelectSecurityGroupsInSPSitesList                            = "Security groups explicitly denied from applying site-level sharing restrictions on SharePoint sites."
                ExemptNativeUsersFromTenantLevelRestricedAccessControl           = "Exempts internal (native) users from tenant-level Restricted Access Control policies."
                AllowSelectSGsInODBListInTenant                                  = "Security groups allowed to configure Restricted Access Control on OneDrive for Business sites."
                DenySelectSGsInODBListInTenant                                   = "Security groups denied from configuring Restricted Access Control on OneDrive for Business sites."
                EnforceRequestDigest                                             = "Enforces request digest validation on SharePoint REST API calls to prevent cross-site request forgery."
                RestrictResourceAccountAccess                                    = "Restricts access to SharePoint content for resource accounts (e.g., meeting room, equipment accounts)."
                RestrictExternalSharingForAgents                                 = "Restricts Microsoft 365 Copilot agents from sharing content externally outside the organization."
                RestrictExternalSharing                                          = "Applies additional restrictions on external sharing beyond the standard SharingCapability setting."
                AllowFileArchive                                                 = "Enables the file archive feature that allows documents to be moved to a long-term archive storage tier."
                AllowFileArchiveOnNewSitesByDefault                              = "Enables file archive on all newly created SharePoint sites by default."
                M365AdditionalStorageSPOEnabled                                  = "Enables Microsoft 365 additional storage capacity that can be allocated to SharePoint Online."
            }

            # Build one row per property: Property | Value | Description | CollectDate
            $spoRows = $spoTenant.PSObject.Properties |
            Where-Object { $_.Name -notlike '@*' } |
            ForEach-Object {
                $propName = $_.Name
                $val = $_.Value
                if ($null -eq $val) { $val = '' }
                elseif ($val -is [System.Collections.IEnumerable] -and $val -isnot [string]) {
                    $val = ($val -join ', ')
                }
                [PSCustomObject]@{
                    Property    = $propName
                    Value       = $val
                    Description = if ($spoDescriptions.ContainsKey($propName)) { $spoDescriptions[$propName] } else { '(no description — verify property name)' }
                    CollectDate = (Get-Date -Format 'yyyy-MM-dd HH:mm')
                }
            }

            Export-ToCsv -Data $spoRows -FileName "SPO_GetSPOTenant_Full_$date.csv"
            Write-Log "  SPO tenant properties collected via Get-SPOTenant: $($spoRows.Count)" 'SUCCESS'
        }
        catch {
            Write-Log "  Get-SPOTenant collection failed: $($_.Exception.Message)" 'WARN'
        }
    }
    catch {
        Write-Log "  SharePoint Online connection failed: $($_.Exception.Message)" 'ERROR'
    }
    finally {
        try { Disconnect-SPOService -ErrorAction SilentlyContinue } catch {}
    }

    Write-Log "  SharePoint Online collection complete." 'SUCCESS'
}

#endregion

#region Workload: Exchange Online

function Collect-ExchangeData {
    Write-Log ""
    Write-Log "==  Exchange Online  ==" 'SUCCESS'

    #---------------------------------------------------------------------------
    # PART 1 — ExchangeOnlineManagement PS module
    #---------------------------------------------------------------------------
    Write-Log "  Connecting ExchangeOnlineManagement PS module for detailed config..."
    $exoModuleConnected = $false
    try {
        Connect-ToExchangeOnline
        $exoModuleConnected = $true
    }
    catch {
        Write-Log "  ExchangeOnlineManagement module unavailable: $($_.Exception.Message)" 'WARN'
        Write-Log "  --> FIX CHECKLIST:" 'WARN'
        Write-Log "      1. Install module (if missing):" 'WARN'
        Write-Log "            Install-Module ExchangeOnlineManagement -Scope CurrentUser" 'WARN'
        Write-Log "      2. Ensure your account has 'Exchange Administrator' role:" 'WARN'
        Write-Log "            Entra ID > Roles and administrators > Exchange Administrator" 'WARN'
        Write-Log "            > Add assignments > select your account > Add" 'WARN'
    }

    if ($exoModuleConnected) {
        try {
            # Inline helper — same *>&1 noise-suppression pattern as the Teams $Collect.
            $ExCollect = {
                param([string]$Label, [string]$FileName, [scriptblock]$Cmd)
                Write-Log "  $Label..." 'DEBUG'
                try {
                    $raw = @(& $Cmd *>&1)
                    $data = @($raw | Where-Object {
                            $_ -isnot [System.Management.Automation.ErrorRecord] -and
                            $_ -isnot [System.Management.Automation.WarningRecord] -and
                            $_ -isnot [System.Management.Automation.InformationRecord] -and
                            $_ -isnot [System.Management.Automation.VerboseRecord] -and
                            $_ -isnot [System.Management.Automation.DebugRecord] -and
                            $_ -isnot [string] -and
                            $null -ne $_
                        })
                    $errs = @($raw | Where-Object { $_ -is [System.Management.Automation.ErrorRecord] })
                    if ($data.Count -gt 0) {
                        $rows = @($data | ForEach-Object { ConvertTo-FlatCsvRow -InputObject $_ })
                        Export-ToCsv -Data $rows -FileName $FileName
                        Write-Log "  ${Label}: $($rows.Count) record(s)" 'SUCCESS'
                    }
                    elseif ($errs.Count -gt 0) {
                        $errMsg = $errs[0].Exception.Message
                        $hint = if ($errMsg -match 'Access Denied|AccessDenied|Forbidden|Unauthorized|insufficient') {
                            "`n  --> FIX: Assign 'Exchange Administrator' or 'Exchange Recipient Administrator'" +
                            "`n       role to your account in Entra ID > Roles and administrators."
                        }
                        else { '' }
                        Write-Log "  ${Label} unavailable: $errMsg$hint" 'WARN'
                    }
                    else { Write-Log "  ${Label}: no data returned" 'WARN' }
                }
                catch {
                    $hint = if ($_.Exception.Message -match 'Access Denied|AccessDenied|Forbidden|Unauthorized|insufficient') {
                        "`n  --> FIX: Assign 'Exchange Administrator' or 'Exchange Recipient Administrator'" +
                        "`n       role to your account in Entra ID > Roles and administrators."
                    }
                    else { '' }
                    Write-Log "  ${Label} unavailable: $($_.Exception.Message)$hint" 'WARN'
                }
            }

            # ---- Mailbox and group inventory ---------------------------------------------
            try {
                Write-Log "  Mailbox usage detail (Get-EXOMailboxStatistics)..."
                $mailboxes = @(Get-EXOMailbox -ResultSize Unlimited -PropertySets Minimum *>&1 |
                    Where-Object { $_ -isnot [System.Management.Automation.ErrorRecord] -and $null -ne $_ -and $_ -isnot [string] })
                $mailboxCount = 0
                $mailboxStatsFailed = 0
                $totalMailboxBytes = 0L
                $totalMailboxItems = 0L
                $totalDeletedItems = 0L
                foreach ($mbx in $mailboxes) {
                    try {
                        $stats = Get-EXOMailboxStatistics -Identity $mbx.UserPrincipalName -ErrorAction Stop
                        $mailboxCount++
                        $totalMailboxBytes += ConvertTo-ByteCount -Value $stats.TotalItemSize
                        $totalMailboxItems += [int64]$stats.ItemCount
                        $totalDeletedItems += [int64]$stats.DeletedItemCount
                    }
                    catch {
                        $mailboxStatsFailed++
                        Write-Log "  Mailbox statistics unavailable for one mailbox: $($_.Exception.Message)" 'DEBUG'
                    }
                }
                $mailboxSummary = @([PSCustomObject]@{
                        ObjectType                 = 'Mailbox'
                        ObjectCount                = $mailboxCount
                        StatisticsFailures         = $mailboxStatsFailed
                        TotalItemCount             = $totalMailboxItems
                        TotalDeletedItemCount      = $totalDeletedItems
                        TotalMailboxSizeBytes      = $totalMailboxBytes
                        TotalMailboxSizeGB         = [math]::Round(($totalMailboxBytes / 1GB), 2)
                        AverageMailboxSizeBytes    = if ($mailboxCount -gt 0) { [math]::Round(($totalMailboxBytes / $mailboxCount), 0) } else { 0 }
                        AverageMailboxSizeGB       = if ($mailboxCount -gt 0) { [math]::Round((($totalMailboxBytes / $mailboxCount) / 1GB), 4) } else { 0 }
                        AverageItemCountPerMailbox = if ($mailboxCount -gt 0) { [math]::Round(($totalMailboxItems / $mailboxCount), 2) } else { 0 }
                        CollectDate                = (Get-Date -Format 'yyyy-MM-dd HH:mm')
                    })
                Export-ToCsv -Data $mailboxSummary -FileName "EXO_MailboxUsageSummary_$date.csv"
                Write-Log "  Mailboxes: $mailboxCount | Total size: $($mailboxSummary[0].TotalMailboxSizeGB) GB | Avg: $($mailboxSummary[0].AverageMailboxSizeGB) GB" 'SUCCESS'
            }
            catch {
                Write-Log "  Native mailbox usage detail unavailable: $($_.Exception.Message)" 'WARN'
                $null = Invoke-GraphReportRequest `
                    -Uri "https://graph.microsoft.com/v1.0/reports/getMailboxUsageMailboxCounts(period='D30')" `
                    -FileName "EXO_MailboxUsageMailboxCounts_$date.csv"
                $null = Invoke-GraphReportRequest `
                    -Uri "https://graph.microsoft.com/v1.0/reports/getMailboxUsageStorage(period='D30')" `
                    -FileName "EXO_MailboxUsageStorage_$date.csv"
            }

            try {
                Write-Log "  Distribution groups summary..."
                $dgCount = @(Get-DistributionGroup -RecipientTypeDetails MailUniversalDistributionGroup -ResultSize Unlimited).Count
                Export-ToCsv -Data @(New-ObjectCountSummary -ObjectType 'DistributionGroup' -Count $dgCount) -FileName "EXO_DistributionGroupsSummary_$date.csv"
                Write-Log "  Distribution groups: $dgCount" 'SUCCESS'
            }
            catch { Write-Log "  Distribution groups summary unavailable: $($_.Exception.Message)" 'WARN' }

            try {
                Write-Log "  Mail-enabled security groups summary..."
                $mesgCount = @(Get-DistributionGroup -RecipientTypeDetails MailUniversalSecurityGroup -ResultSize Unlimited).Count
                Export-ToCsv -Data @(New-ObjectCountSummary -ObjectType 'MailEnabledSecurityGroup' -Count $mesgCount) -FileName "EXO_MailEnabledSecurityGroupsSummary_$date.csv"
                Write-Log "  Mail-enabled security groups: $mesgCount" 'SUCCESS'
            }
            catch { Write-Log "  Mail-enabled security groups summary unavailable: $($_.Exception.Message)" 'WARN' }

            # Microsoft 365 usage report data has no reliable ExchangeOnlineManagement
            # cmdlet equivalent; use aggregate-only Graph Reports to avoid user PII.
            $null = Invoke-GraphReportRequest `
                -Uri "https://graph.microsoft.com/v1.0/reports/getEmailActivityCounts(period='D30')" `
                -FileName "EXO_EmailActivityCounts_$date.csv"

            # ---- Organization & authentication ----------------------------------------
            & $ExCollect 'OrgConfig'       "EXO_OrgConfig_$date.csv" { Get-OrganizationConfig }
            & $ExCollect 'TransportConfig' "EXO_TransportConfig_$date.csv" { Get-TransportConfig }
            # Basic auth: authentication policies list each blocked legacy-auth protocol
            & $ExCollect 'AuthPolicies'    "EXO_AuthPolicies_$date.csv" { Get-AuthenticationPolicy }
            # POP3/IMAP per-plan defaults
            & $ExCollect 'CASMailboxPlans' "EXO_CASMailboxPlans_$date.csv" { Get-CASMailboxPlan }

            # ---- Mail flow --------------------------------------------------------------
            # Accepted domains — also used below for DNS checks and vanity-namespace flag
            & $ExCollect 'AcceptedDomains'    "EXO_AcceptedDomains_$date.csv" { Get-AcceptedDomain }
            & $ExCollect 'RemoteDomains'      "EXO_RemoteDomains_$date.csv" { Get-RemoteDomain }
            # Connectors — inbound/outbound; UseMXRecord=False outbound = centralized mail flow
            & $ExCollect 'InboundConnectors'  "EXO_InboundConnectors_$date.csv" { Get-InboundConnector }
            & $ExCollect 'OutboundConnectors' "EXO_OutboundConnectors_$date.csv" { Get-OutboundConnector }
            & $ExCollect 'TransportRules'     "EXO_TransportRules_$date.csv" { Get-TransportRule }

            # ---- Security & message hygiene --------------------------------------------
            # Defender for Office 365
            & $ExCollect 'AntiPhishPolicies'      "EXO_AntiPhishPolicies_$date.csv" { Get-AntiPhishPolicy }
            & $ExCollect 'SafeLinksPolicies'      "EXO_SafeLinksPolicies_$date.csv" { Get-SafeLinksPolicy }
            & $ExCollect 'SafeAttachmentPolicies' "EXO_SafeAttachmentPolicies_$date.csv" { Get-SafeAttachmentPolicy }
            # EOP anti-spam / malware / connection filters
            & $ExCollect 'InboundSpamFilter'      "EXO_InboundSpamFilter_$date.csv" { Get-HostedContentFilterPolicy }
            # OutboundSpamFilter — AutoForwardingMode column answers "Is auto-forwarding allowed?"
            & $ExCollect 'OutboundSpamFilter'     "EXO_OutboundSpamFilter_$date.csv" { Get-HostedOutboundSpamFilterPolicy }
            & $ExCollect 'MalwareFilter'          "EXO_MalwareFilter_$date.csv" { Get-MalwareFilterPolicy }
            & $ExCollect 'ConnectionFilter'       "EXO_ConnectionFilter_$date.csv" { Get-HostedConnectionFilterPolicy }
            # DKIM signing config — Enabled/Status per domain
            & $ExCollect 'DkimSigningConfig'      "EXO_DkimSigningConfig_$date.csv" { Get-DkimSigningConfig }

            # ---- Message encryption & compliance ---------------------------------------
            & $ExCollect 'IRMConfig'  "EXO_IRMConfig_$date.csv" { Get-IRMConfiguration }
            & $ExCollect 'OMEConfig'  "EXO_OMEConfig_$date.csv" { Get-OMEConfiguration }

            # ---- Collaboration ---------------------------------------------------------
            # SharingPolicy — answers "Is calendar sharing enabled / what's the scope?"
            & $ExCollect 'SharingPolicies' "EXO_SharingPolicies_$date.csv" { Get-SharingPolicy }

            # ---- Compliance / retention / hold -----------------------------------------
            & $ExCollect 'JournalRules'      "EXO_JournalRules_$date.csv" { Get-JournalRule }
            & $ExCollect 'RetentionPolicies' "EXO_RetentionPolicies_$date.csv" { Get-RetentionPolicy }
            & $ExCollect 'RetentionTags'     "EXO_RetentionTags_$date.csv" { Get-RetentionPolicyTag }

            # ---- Litigation hold mailboxes ---------------------------------------------
            try {
                Write-Log "  Litigation hold mailboxes..."
                $holdRaw = @(Get-EXOMailbox -ResultSize 5000 `
                        -Filter "LitigationHoldEnabled -eq `$true" `
                        -PropertySets Minimum, Hold *>&1)
                $holdMbx = @($holdRaw | Where-Object {
                        $_ -isnot [System.Management.Automation.ErrorRecord] -and $null -ne $_ -and $_ -isnot [string]
                    })
                $holdErrs = @($holdRaw | Where-Object { $_ -is [System.Management.Automation.ErrorRecord] })
                if ($holdMbx.Count -gt 0) {
                    $holdDurations = @($holdMbx | Group-Object -Property LitigationHoldDuration | ForEach-Object {
                            New-ObjectCountSummary -ObjectType 'LitigationHoldMailbox' -Count $_.Count -AdditionalProperties @{
                                GroupingProperty = 'LitigationHoldDuration'
                                GroupingValue    = if ([string]::IsNullOrWhiteSpace([string]$_.Name)) { 'Unspecified' } else { $_.Name }
                            }
                        })
                    Export-ToCsv -Data $holdDurations -FileName "EXO_LitigationHoldMailboxesSummary_$date.csv"
                    Write-Log "  Litigation hold mailboxes: $($holdMbx.Count)" 'SUCCESS'
                }
                elseif ($holdErrs.Count -gt 0) {
                    Write-Log "  Litigation hold query unavailable: $($holdErrs[0].Exception.Message)" 'WARN'
                }
                else {
                    Write-Log "  Litigation hold mailboxes: 0 found" 'SUCCESS'
                    Export-ToCsv -Data @([PSCustomObject]@{ Count = 0; Note = 'No mailboxes with LitigationHoldEnabled'; CollectDate = (Get-Date -Format 'yyyy-MM-dd HH:mm') }) `
                        -FileName "EXO_LitigationHoldMailboxesSummary_$date.csv"
                }
            }
            catch { Write-Log "  Litigation hold query failed: $($_.Exception.Message)" 'WARN' }

            # ---- Archive mailbox summary -----------------------------------------------
            try {
                Write-Log "  Archive mailbox summary..."
                $archRaw = @(Get-EXOMailbox -ResultSize 5000 -Filter "ArchiveStatus -ne 'None'" `
                        -PropertySets Minimum, Archive *>&1)
                $archMbx = @($archRaw | Where-Object {
                        $_ -isnot [System.Management.Automation.ErrorRecord] -and $null -ne $_ -and $_ -isnot [string]
                    })
                $archErrs = @($archRaw | Where-Object { $_ -is [System.Management.Automation.ErrorRecord] })
                if ($archMbx.Count -gt 0) {
                    $archData = @($archMbx | Group-Object -Property ArchiveStatus, ArchiveState, AutoExpandingArchiveEnabled | ForEach-Object {
                            New-ObjectCountSummary -ObjectType 'ArchiveMailbox' -Count $_.Count -AdditionalProperties @{
                                GroupingProperty = 'ArchiveStatus|ArchiveState|AutoExpandingArchiveEnabled'
                                GroupingValue    = if ([string]::IsNullOrWhiteSpace([string]$_.Name)) { 'Unspecified' } else { $_.Name }
                            }
                        })
                    Export-ToCsv -Data $archData -FileName "EXO_ArchiveMailboxesSummary_$date.csv"
                    Write-Log "  Archive mailboxes: $($archMbx.Count)" 'SUCCESS'
                }
                elseif ($archErrs.Count -gt 0) {
                    Write-Log "  Archive mailbox query unavailable: $($archErrs[0].Exception.Message)" 'WARN'
                }
                else {
                    Write-Log "  Archive mailboxes: 0 found" 'SUCCESS'
                }
            }
            catch { Write-Log "  Archive mailbox query failed: $($_.Exception.Message)" 'WARN' }

            # ---- Public folders --------------------------------------------------------
            try {
                Write-Log "  Public folders..."
                $pfRaw = @(Get-PublicFolder -Identity '\' *>&1)
                $pfErrs = @($pfRaw | Where-Object { $_ -is [System.Management.Automation.ErrorRecord] })
                $pfRoot = @($pfRaw | Where-Object { $_ -isnot [System.Management.Automation.ErrorRecord] -and $null -ne $_ })
                if ($pfErrs.Count -gt 0 -or $pfRoot.Count -eq 0) {
                    $errDetail = if ($pfErrs.Count -gt 0) { $pfErrs[0].Exception.Message } else { 'Root public folder not found' }
                    Write-Log "  Public folders: not configured or unavailable ($errDetail)" 'SUCCESS'
                    Export-ToCsv -Data @([PSCustomObject]@{
                            PublicFoldersPresent = $false
                            DetailAvailable      = $false
                            CollectDate          = (Get-Date -Format 'yyyy-MM-dd HH:mm')
                        }) -FileName "EXO_PublicFoldersSummary_$date.csv"
                }
                else {
                    Write-Log "  Public folders detected — collecting top-level tree..."
                    $pfData = @(Get-PublicFolder -Recurse -ResultSize 500 *>&1) |
                    Where-Object { $_ -isnot [System.Management.Automation.ErrorRecord] -and $null -ne $_ -and $_ -isnot [string] }
                    if ($pfData.Count -gt 0) {
                        $pfRows = @($pfData | Group-Object -Property FolderType, MailEnabled | ForEach-Object {
                                New-ObjectCountSummary -ObjectType 'PublicFolder' -Count $_.Count -AdditionalProperties @{
                                    GroupingProperty = 'FolderType|MailEnabled'
                                    GroupingValue    = if ([string]::IsNullOrWhiteSpace([string]$_.Name)) { 'Unspecified' } else { $_.Name }
                                }
                            }
                        )
                        Export-ToCsv -Data $pfRows -FileName "EXO_PublicFoldersSummary_$date.csv"
                        Write-Log "  Public folders: $($pfData.Count) collected" 'SUCCESS'
                    }
                }
            }
            catch { Write-Log "  Public folder query failed: $($_.Exception.Message)" 'WARN' }
        }
        finally {
            try { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue | Out-Null } catch {}
            Write-Log "  Disconnected from ExchangeOnlineManagement module" 'DEBUG'
        }
    }

    #---------------------------------------------------------------------------
    # PART 2 — DNS security checks: MX, SPF, DKIM, DMARC
    # Domain list sourced from the native Get-AcceptedDomain export.
    # Skipped silently if Resolve-DnsName is unavailable (non-Windows platform).
    #---------------------------------------------------------------------------
    try {
        Write-Log "  DNS security checks (MX / SPF / DKIM / DMARC)..."
        $checkDomains = @()
        $accDomPath = Join-Path $OutputFolder "EXO_AcceptedDomains_$date.csv"
        if (Test-Path $accDomPath) {
            $checkDomains = @(Import-Csv $accDomPath -Encoding UTF8 |
                Where-Object { $_.DomainName -notmatch '\.onmicrosoft\.com$' } |
                Select-Object -ExpandProperty DomainName)
        }
        if ($checkDomains.Count -eq 0) {
            Write-Log "  DNS checks skipped — no non-onmicrosoft.com accepted domains found from Exchange Online" 'WARN'
        }
        elseif (-not (Get-Command Resolve-DnsName -ErrorAction SilentlyContinue)) {
            Write-Log "  DNS checks skipped — Resolve-DnsName not available on this platform" 'WARN'
        }
        else {
            $dnsRows = foreach ($dom in $checkDomains) {
                $mx = $spf = $dmarc = $dkimSel1 = $dkimSel2 = ''
                try {
                    $mx = (@(Resolve-DnsName -Name $dom -Type MX -ErrorAction Stop |
                            Where-Object { $_.Type -eq 'MX' } | Sort-Object Preference |
                            ForEach-Object { "$($_.NameExchange) (pref $($_.Preference))" }) -join '; ')
                }
                catch {}
                try {
                    $spf = (@(Resolve-DnsName -Name $dom -Type TXT -ErrorAction Stop |
                            Where-Object { $_.Strings -like 'v=spf1*' } |
                            ForEach-Object { $_.Strings -join ' ' }) -join ' ')
                }
                catch {}
                try {
                    $dmarc = (@(Resolve-DnsName -Name "_dmarc.$dom" -Type TXT -ErrorAction Stop |
                            ForEach-Object { $_.Strings -join ' ' }) -join ' ')
                }
                catch {}
                try {
                    $dkimSel1 = (@(Resolve-DnsName -Name "selector1._domainkey.$dom" -Type CNAME -ErrorAction Stop |
                            Select-Object -ExpandProperty NameHost) -join '; ')
                }
                catch {}
                try {
                    $dkimSel2 = (@(Resolve-DnsName -Name "selector2._domainkey.$dom" -Type CNAME -ErrorAction Stop |
                            Select-Object -ExpandProperty NameHost) -join '; ')
                }
                catch {}
                [PSCustomObject]@{
                    Domain         = $dom
                    MX_Record      = $mx
                    SPF_Record     = $spf
                    SPF_Present    = ($spf -ne '')
                    DMARC_Record   = $dmarc
                    DMARC_Present  = ($dmarc -ne '')
                    DKIM_Selector1 = $dkimSel1
                    DKIM_Selector2 = $dkimSel2
                    DKIM_Present   = ($dkimSel1 -ne '' -or $dkimSel2 -ne '')
                    CollectDate    = (Get-Date -Format 'yyyy-MM-dd HH:mm')
                }
            }
            Export-ToCsv -Data $dnsRows -FileName "EXO_DNSSecurityChecks_$date.csv"
            Write-Log "  DNS checks completed for $($dnsRows.Count) domain(s)" 'SUCCESS'
        }
    }
    catch { Write-Log "  DNS security checks failed: $($_.Exception.Message)" 'WARN' }

    Write-Log "  Exchange Online collection complete." 'SUCCESS'
}

#endregion

#region Workload: Microsoft Teams

function Collect-TeamsData {
    Write-Log ""
    Write-Log "==  Microsoft Teams  ==" 'SUCCESS'

    Write-Log "  Connecting MicrosoftTeams PowerShell module..."
    $teamsModuleConnected = $false
    try {
        Connect-ToMicrosoftTeams
        $teamsModuleConnected = $true
    }
    catch {
        Write-Log "  MicrosoftTeams module unavailable: $($_.Exception.Message)" 'WARN'
        if (-not $script:assemblyConflictHintShown -and $_.Exception.Message -match 'Method not found' -and $_.Exception.Message -match 'Microsoft\.Identity\.Client') {
            $script:assemblyConflictHintShown = $true
            Write-Log '  Detected auth assembly conflict in the current host. Re-run with: pwsh -NoProfile -File "<path-to-this-script>"' 'WARN'
            Write-Log '  If needed, update modules: MicrosoftTeams and Microsoft.Graph (CurrentUser scope).' 'WARN'
        }
        Write-Log "  --> Install-Module MicrosoftTeams -Scope CurrentUser" 'WARN'
        Write-Log "      Assign 'Teams Administrator' or 'Global Reader' role to your account" 'WARN'
        Write-Log "      in Entra ID > Roles and administrators, then re-run." 'WARN'
    }

    if ($teamsModuleConnected) {
        try {
            Write-Log "  Teams summary (Get-Team)..."
            $teams = @(Get-Team)
            $teamData = @($teams | Group-Object -Property Visibility, Archived | ForEach-Object {
                    New-ObjectCountSummary -ObjectType 'Team' -Count $_.Count -AdditionalProperties @{
                        GroupingProperty = 'Visibility|Archived'
                        GroupingValue    = if ([string]::IsNullOrWhiteSpace([string]$_.Name)) { 'Unspecified' } else { $_.Name }
                    }
                }
            )
            if ($teamData.Count -eq 0) {
                $teamData = @(New-ObjectCountSummary -ObjectType 'Team' -Count 0)
            }
            Export-ToCsv -Data $teamData -FileName "Teams_Summary_$date.csv"
            Write-Log "  Total Teams: $($teams.Count)" 'SUCCESS'
        }
        catch { Write-Log "  Teams list unavailable: $($_.Exception.Message)" 'WARN' }

        try {
            Write-Log "  Teams tenant settings (Get-CsTenant)..."
            $tenantSettings = @(Get-CsTenant)
            if ($tenantSettings.Count -gt 0) {
                $rows = @($tenantSettings | ForEach-Object { ConvertTo-FlatCsvRow -InputObject $_ })
                Export-ToCsv -Data $rows -FileName "Teams_TenantSettings_$date.csv"
            }
        }
        catch { Write-Log "  Teams tenant settings unavailable: $($_.Exception.Message)" 'WARN' }

        try {
            Write-Log "  Teams app settings..."
            $appSettingsCmd = @('Get-CsTeamsAppSettings', 'Get-TeamsAppSettings') |
            Where-Object { Get-Command -Name $_ -ErrorAction SilentlyContinue } |
            Select-Object -First 1
            if ($appSettingsCmd) {
                $appSettings = @(& $appSettingsCmd)
                if ($appSettings.Count -gt 0) {
                    $rows = @($appSettings | ForEach-Object { ConvertTo-FlatCsvRow -InputObject $_ })
                    Export-ToCsv -Data $rows -FileName "Teams_AppSettings_$date.csv"
                }
            }
            else {
                Write-Log "  Teams app settings unavailable: no Teams app settings cmdlet is exposed by this MicrosoftTeams module version." 'WARN'
            }
        }
        catch { Write-Log "  Teams app settings unavailable: $($_.Exception.Message)" 'WARN' }
    }

    # MicrosoftTeams does not expose these M365 usage telemetry reports as native
    # cmdlets; keep Graph Reports API calls so the prior report CSVs are not lost.
    # --- Teams user activity counts (30 days) ---
    $null = Invoke-GraphReportRequest `
        -Uri "https://graph.microsoft.com/v1.0/reports/getTeamsUserActivityCounts(period='D30')" `
        -FileName "Teams_ActivityCounts_$date.csv"

    # --- Teams active user counts (30 days) ---
    $null = Invoke-GraphReportRequest `
        -Uri "https://graph.microsoft.com/v1.0/reports/getTeamsUserActivityUserCounts(period='D30')" `
        -FileName "Teams_ActivityUserCounts_$date.csv"

    # --- Teams device usage distribution (30 days) ---
    $null = Invoke-GraphReportRequest `
        -Uri "https://graph.microsoft.com/v1.0/reports/getTeamsDeviceUsageDistributionUserCounts(period='D30')" `
        -FileName "Teams_DeviceUsageDistributionUserCounts_$date.csv"

    # --- Teams meeting activity counts (30 days) ---
    $meetingCountsPath = Invoke-GraphReportRequest `
        -Uri "https://graph.microsoft.com/v1.0/reports/getTeamsMeetingActivityCounts(period='D30')" `
        -FileName "Teams_MeetingActivityCounts_$date.csv"

    if (-not $meetingCountsPath) {
        Write-Log "  NOTE: Meeting activity report(s) failed with 400 Bad Request." 'WARN'
        Write-Log "        This is a tenant/license issue, NOT a permissions issue." 'WARN'
        Write-Log "        These reports require Teams meeting telemetry to be available in the tenant." 'WARN'
        Write-Log "        Tenants without Teams Essentials/M365 Business Basic or higher, or with" 'WARN'
        Write-Log "        meeting recording/telemetry disabled, will see 400 regardless of API permissions." 'WARN'
    }

    # --- Teams team activity counts (30 days) ---
    $null = Invoke-GraphReportRequest `
        -Uri "https://graph.microsoft.com/v1.0/reports/getTeamsTeamActivityCounts(period='D30')" `
        -FileName "Teams_TeamActivityCounts_$date.csv"

    # --- MicrosoftTeams PowerShell module: policies and configuration ---

    if ($teamsModuleConnected) {
        try {
            # Inline helper: invoke a CS cmdlet, flatten each result row, export to CSV.
            # Uses *>&1 to redirect ALL PowerShell streams (errors, warnings, verbose,
            # debug, information/Write-Host) into the pipeline.  This suppresses the
            # raw "Access Denied / Correlation id" console spam that some CS cmdlets
            # emit as non-terminating errors or via Write-Host, regardless of
            # $ErrorActionPreference.  We then split the captured objects by type:
            #   - ErrorRecord  → treated as an error; shows FIX hint if Access Denied
            #   - *Record types / strings → discarded (banners, verbose noise)
            #   - everything else → legitimate cmdlet output, exported to CSV
            $Collect = {
                param([string]$Label, [string]$FileName, [scriptblock]$Cmd)
                Write-Log "  $Label..." 'DEBUG'
                try {
                    $raw = @(& $Cmd *>&1)
                    $data = @($raw | Where-Object {
                            $_ -isnot [System.Management.Automation.ErrorRecord] -and
                            $_ -isnot [System.Management.Automation.WarningRecord] -and
                            $_ -isnot [System.Management.Automation.InformationRecord] -and
                            $_ -isnot [System.Management.Automation.VerboseRecord] -and
                            $_ -isnot [System.Management.Automation.DebugRecord] -and
                            $_ -isnot [string] -and
                            $null -ne $_
                        })
                    $errs = @($raw | Where-Object { $_ -is [System.Management.Automation.ErrorRecord] })

                    if ($data.Count -gt 0) {
                        $rows = @($data | ForEach-Object { ConvertTo-FlatCsvRow -InputObject $_ })
                        Export-ToCsv -Data $rows -FileName $FileName
                        Write-Log "  ${Label}: $($rows.Count) record(s)" 'SUCCESS'
                    }
                    elseif ($errs.Count -gt 0) {
                        $errMsg = $errs[0].Exception.Message
                        $hint = if ($errMsg -match 'Access Denied|AccessDenied|access denied|Forbidden') {
                            "`n  --> FIX: In Entra ID > Roles and administrators, assign 'Teams Administrator'" +
                            "`n       role to your account, then re-run."
                        }
                        else { '' }
                        Write-Log "  ${Label} unavailable: $errMsg$hint" 'WARN'
                    }
                    else { Write-Log "  ${Label}: no data returned" 'WARN' }
                }
                catch {
                    $hint = if ($_.Exception.Message -match 'Access Denied|AccessDenied|access denied|Forbidden') {
                        "`n  --> FIX: In Entra ID > Roles and administrators, assign 'Teams Administrator'" +
                        "`n       role to your account, then re-run."
                    }
                    else { '' }
                    Write-Log "  ${Label} unavailable: $($_.Exception.Message)$hint" 'WARN'
                }
            }

            # ---- Tenant / global configuration ----
            & $Collect 'CsTenant'                     "Teams_CsTenant_$date.csv" { Get-CsTenant }
            & $Collect 'TeamsClientConfiguration'     "Teams_ClientConfig_$date.csv" { Get-CsTeamsClientConfiguration }
            & $Collect 'FederationConfiguration'      "Teams_FederationConfig_$date.csv" { Get-CsTenantFederationConfiguration }
            & $Collect 'MeetingConfiguration'         "Teams_MeetingConfig_$date.csv" { Get-CsTeamsMeetingConfiguration }
            & $Collect 'TeamsUpgradeConfiguration'    "Teams_UpgradeConfig_$date.csv" { Get-CsTeamsUpgradeConfiguration }
            & $Collect 'GuestCallingConfiguration'    "Teams_GuestCallingConfig_$date.csv" { Get-CsTeamsGuestCallingConfiguration }
            & $Collect 'GuestMeetingConfiguration'    "Teams_GuestMeetingConfig_$date.csv" { Get-CsTeamsGuestMeetingConfiguration }
            & $Collect 'GuestMessagingConfiguration'  "Teams_GuestMessagingConfig_$date.csv" { Get-CsTeamsGuestMessagingConfiguration }
            & $Collect 'LiveEventsConfiguration'      "Teams_LiveEventsConfig_$date.csv" { Get-CsTeamsMeetingBroadcastConfiguration }

            # ---- Policies (Global policy + any custom instances per type) ----
            & $Collect 'MeetingPolicies'              "Teams_Policy_Meeting_$date.csv" { Get-CsTeamsMeetingPolicy }
            & $Collect 'MessagingPolicies'            "Teams_Policy_Messaging_$date.csv" { Get-CsTeamsMessagingPolicy }
            & $Collect 'CallingPolicies'              "Teams_Policy_Calling_$date.csv" { Get-CsTeamsCallingPolicy }
            & $Collect 'ChannelsPolicies'             "Teams_Policy_Channels_$date.csv" { Get-CsTeamsChannelsPolicy }
            & $Collect 'AppSetupPolicies'             "Teams_Policy_AppSetup_$date.csv" { Get-CsTeamsAppSetupPolicy }
            & $Collect 'AppPermissionPolicies'        "Teams_Policy_AppPermission_$date.csv" { Get-CsTeamsAppPermissionPolicy }
            & $Collect 'LiveEventsPolicies'           "Teams_Policy_LiveEvents_$date.csv" { Get-CsTeamsMeetingBroadcastPolicy }
            & $Collect 'TeamsUpgradePolicies'         "Teams_Policy_Upgrade_$date.csv" { Get-CsTeamsUpgradePolicy }
            & $Collect 'ShiftsPolicies'               "Teams_Policy_Shifts_$date.csv" { Get-CsTeamsShiftsPolicy }
            & $Collect 'AudioConferencingPolicies'    "Teams_Policy_AudioConferencing_$date.csv" { Get-CsTeamsAudioConferencingPolicy }
            & $Collect 'FilesPolicies'                "Teams_Policy_Files_$date.csv" { Get-CsTeamsFilesPolicy }
            & $Collect 'EmergencyCallingPolicies'     "Teams_Policy_EmergencyCalling_$date.csv" { Get-CsTeamsEmergencyCallingPolicy }
            & $Collect 'EmergencyCallRoutingPolicies' "Teams_Policy_EmergencyCallRouting_$date.csv" { Get-CsTeamsEmergencyCallRoutingPolicy }
            & $Collect 'CallParkPolicies'             "Teams_Policy_CallPark_$date.csv" { Get-CsTeamsCallParkPolicy }
            & $Collect 'UpdateManagementPolicies'     "Teams_Policy_UpdateMgmt_$date.csv" { Get-CsTeamsUpdateManagementPolicy }
            & $Collect 'VoiceAppsPolicies'            "Teams_Policy_VoiceApps_$date.csv" { Get-CsTeamsVoiceApplicationsPolicy }
            & $Collect 'ComplianceRecordingPolicies'  "Teams_Policy_ComplianceRecording_$date.csv" { Get-CsTeamsComplianceRecordingPolicy }
            & $Collect 'IPPhonePolicies'              "Teams_Policy_IPPhone_$date.csv" { Get-CsTeamsIPPhonePolicy }

            # ---- Voice / telephony ----
            & $Collect 'OnlineVoiceRoutingPolicies'   "Teams_Voice_RoutingPolicies_$date.csv" { Get-CsOnlineVoiceRoutingPolicy }
            & $Collect 'OnlineDialPlans'              "Teams_Voice_DialPlans_$date.csv" { Get-CsTenantDialPlan }
            & $Collect 'OnlinePstnUsages'             "Teams_Voice_PstnUsages_$date.csv" { Get-CsOnlinePstnUsage }
            & $Collect 'AutoAttendants'               "Teams_Voice_AutoAttendants_$date.csv" { Get-CsAutoAttendant }
            & $Collect 'CallQueues'                   "Teams_Voice_CallQueues_$date.csv" { Get-CsCallQueue }
        }
        finally {
            try { Disconnect-MicrosoftTeams -Confirm:$false -ErrorAction SilentlyContinue | Out-Null } catch {}
            Write-Log "  Disconnected from MicrosoftTeams module" 'DEBUG'
        }
    }

    Write-Log "  Microsoft Teams collection complete." 'SUCCESS'
}

#endregion

#region Workload: OneDrive for Business

function Collect-OneDriveData {
    Write-Log ""
    Write-Log "==  OneDrive for Business  ==" 'SUCCESS'

    if ([string]::IsNullOrWhiteSpace($tenantUrl)) {
        Write-Log "  Skipping OneDrive for Business — `$tenantUrl not configured in the Configuration section." 'WARN'
        return
    }

    $adminUrl = $tenantUrl.TrimEnd('/') -replace 'https://([^.]+)\.sharepoint\.com.*', 'https://$1-admin.sharepoint.com'
    Write-Log "  Connecting to SharePoint Online admin ($adminUrl) for OneDrive collection..."

    try {
        if ($PSVersionTable.PSVersion.Major -gt 5) {
            Import-Module Microsoft.Online.SharePoint.PowerShell `
                -UseWindowsPowerShell -DisableNameChecking -ErrorAction Stop
        }
        else {
            Import-Module Microsoft.Online.SharePoint.PowerShell `
                -DisableNameChecking -ErrorAction Stop
        }
        Connect-SPOService -Url $adminUrl -ErrorAction Stop
        Write-Log "  Connected to SharePoint Online admin." 'SUCCESS'

        # --- OneDrive site summary (count + storage totals) ---
        Write-Log "  OneDrive site summary (Get-SPOSite -IncludePersonalSite)..."
        try {
            $allSitesWithPersonal = Get-SPOSite -IncludePersonalSite $true -Limit All
            $odbSites = $allSitesWithPersonal | Where-Object { $_.Url -like '*-my.sharepoint.com/personal/*' }

            $odbCount = @($odbSites).Count
            $totalUsedMB = (@($odbSites) | Measure-Object -Property StorageUsageCurrent -Sum).Sum
            $totalQuotaMB = (@($odbSites) | Measure-Object -Property StorageQuota -Sum).Sum

            $summaryData = @([PSCustomObject]@{
                    TotalSites          = $odbCount
                    TotalStorageUsedGB  = [math]::Round(($totalUsedMB / 1024), 2)
                    TotalStorageQuotaGB = [math]::Round(($totalQuotaMB / 1024), 2)
                    PercentQuotaUsed    = if ($totalQuotaMB -gt 0) { [math]::Round((($totalUsedMB / $totalQuotaMB) * 100), 2) } else { 0 }
                    AvgStorageUsedGB    = if ($odbCount -gt 0) { [math]::Round((($totalUsedMB / $odbCount) / 1024), 2) } else { 0 }
                    CollectDate         = (Get-Date -Format 'yyyy-MM-dd HH:mm')
                })
            Export-ToCsv -Data $summaryData -FileName "ODB_StorageSummary_$date.csv"
            Write-Log ("  ODB summary: $odbCount sites | " +
                "$($summaryData[0].TotalStorageUsedGB) GB used of $($summaryData[0].TotalStorageQuotaGB) GB quota") 'SUCCESS'
        }
        catch { Write-Log "  OneDrive site summary failed: $($_.Exception.Message)" 'WARN' }

        # --- OneDrive tenant quota settings ---
        try {
            Write-Log "  OneDrive quota settings (Get-SPOTenant)..."
            $spoTenant = Get-SPOTenant
            $odbQuota = @([PSCustomObject]@{
                    OneDriveStorageQuotaMB               = $spoTenant.OneDriveStorageQuota
                    OrphanedPersonalSitesRetentionDays   = $spoTenant.OrphanedPersonalSitesRetentionPeriod
                    OneDriveSharingCapability            = $spoTenant.OneDriveSharingCapability
                    OneDriveDefaultShareLinkScope        = $spoTenant.OneDriveDefaultShareLinkScope
                    OneDriveDefaultShareLinkRole         = $spoTenant.OneDriveDefaultShareLinkRole
                    OneDriveRequestFilesLinkEnabled      = $spoTenant.OneDriveRequestFilesLinkEnabled
                    OneDriveBlockGuestsAsSiteAdmin       = $spoTenant.OneDriveBlockGuestsAsSiteAdmin
                    OneDriveForGuestsEnabled             = $spoTenant.OneDriveForGuestsEnabled
                    MassDeleteNotificationDisabledForODB = $spoTenant.MassDeleteNotificationDisabledForODB
                    HideSyncButtonOnODB                  = $spoTenant.HideSyncButtonOnODB
                    CollectDate                          = (Get-Date -Format 'yyyy-MM-dd HH:mm')
                })
            Export-ToCsv -Data $odbQuota -FileName "ODB_TenantSettings_$date.csv"
            Write-Log "  OneDrive tenant settings collected." 'SUCCESS'
        }
        catch { Write-Log "  OneDrive tenant settings unavailable: $($_.Exception.Message)" 'WARN' }
    }
    catch {
        Write-Log "  OneDrive for Business connection failed: $($_.Exception.Message)" 'ERROR'
    }
    finally {
        try { Disconnect-SPOService -ErrorAction SilentlyContinue } catch {}
    }

    Write-Log "  OneDrive for Business collection complete." 'SUCCESS'
}

#endregion

#region Workload: Entra ID

function Collect-EntraIDData {
    Write-Log ""
    Write-Log "==  Entra ID  ==" 'SUCCESS'

    # --- All users ---
    try {
        Write-Log "  Users..."
        $users = Invoke-GraphPagedRequest -Uri `
            "https://graph.microsoft.com/v1.0/users?`$select=id,userType,accountEnabled,assignedLicenses,onPremisesSyncEnabled&`$top=999"
        $totalUsers = @($users).Count
        $licensed = @($users | Where-Object { $_.assignedLicenses.Count -gt 0 }).Count
        $guests = @($users | Where-Object { $_.userType -eq 'Guest' }).Count
        $disabled = @($users | Where-Object { $_.accountEnabled -eq $false }).Count
        $synced = @($users | Where-Object { $_.onPremisesSyncEnabled -eq $true }).Count
        $userData = @(
            [PSCustomObject]@{
                ObjectType            = 'User'
                TotalUsers            = $totalUsers
                MemberUsers           = $totalUsers - $guests
                GuestUsers            = $guests
                LicensedUsers         = $licensed
                UnlicensedUsers       = $totalUsers - $licensed
                EnabledUsers          = $totalUsers - $disabled
                DisabledUsers         = $disabled
                OnPremisesSyncedUsers = $synced
                CloudOnlyUsers        = $totalUsers - $synced
                CollectDate           = (Get-Date -Format 'yyyy-MM-dd HH:mm')
            }
        )
        Export-ToCsv -Data $userData -FileName "EntraID_UsersSummary_$date.csv"
        Write-Log "  Users: $totalUsers total | $licensed licensed | $guests guests" 'SUCCESS'
    }
    catch { Write-Log "  Users unavailable: $($_.Exception.Message)" 'WARN' }

    # --- All groups ---
    try {
        Write-Log "  Groups..."
        $groups = Invoke-GraphPagedRequest -Uri `
            "https://graph.microsoft.com/v1.0/groups?`$select=id,groupTypes,mailEnabled,securityEnabled,membershipRule&`$top=999"
        $typedGroups = foreach ($g in $groups) {
            $gType = if ($g.groupTypes -contains 'Unified') { 'Microsoft 365' }
            elseif ($g.mailEnabled -and -not $g.securityEnabled) { 'Distribution' }
            elseif ($g.mailEnabled -and $g.securityEnabled) { 'Mail-Enabled Security' }
            else { 'Security' }
            [PSCustomObject]@{
                GroupType       = $gType
                IsDynamic       = [bool]$g.membershipRule
                MailEnabled     = $g.mailEnabled
                SecurityEnabled = $g.securityEnabled
            }
        }
        $groupData = @($typedGroups | Group-Object -Property GroupType, IsDynamic, MailEnabled, SecurityEnabled | ForEach-Object {
                New-ObjectCountSummary -ObjectType 'Group' -Count $_.Count -AdditionalProperties @{
                    GroupingProperty = 'GroupType|IsDynamic|MailEnabled|SecurityEnabled'
                    GroupingValue    = if ([string]::IsNullOrWhiteSpace([string]$_.Name)) { 'Unspecified' } else { $_.Name }
                }
            })
        Export-ToCsv -Data $groupData -FileName "EntraID_GroupsSummary_$date.csv"
        Write-Log "  Groups: $(@($groups).Count)" 'SUCCESS'
    }
    catch { Write-Log "  Groups unavailable: $($_.Exception.Message)" 'WARN' }

    # --- License SKUs ---
    try {
        Write-Log "  License SKUs..."
        $skus = Invoke-GraphPagedRequest -Uri `
            'https://graph.microsoft.com/v1.0/subscribedSkus?$select=skuPartNumber,skuId,capabilityStatus,consumedUnits,prepaidUnits'
        $skuData = foreach ($s in $skus) {
            [PSCustomObject]@{
                SkuPartNumber    = $s.skuPartNumber
                SkuId            = $s.skuId
                CapabilityStatus = $s.capabilityStatus
                ConsumedUnits    = $s.consumedUnits
                EnabledUnits     = $s.prepaidUnits.enabled
                WarningUnits     = $s.prepaidUnits.warning
                SuspendedUnits   = $s.prepaidUnits.suspended
                CollectDate      = (Get-Date -Format 'yyyy-MM-dd HH:mm')
            }
        }
        Export-ToCsv -Data $skuData -FileName "EntraID_LicenseSKUs_$date.csv"
        Write-Log "  License SKUs: $($skuData.Count)" 'SUCCESS'
    }
    catch { Write-Log "  License SKUs unavailable: $($_.Exception.Message)" 'WARN' }

    # --- Conditional Access policies ---
    try {
        Write-Log "  Conditional Access policies..."
        $policies = Invoke-GraphPagedRequest -Uri `
            'https://graph.microsoft.com/v1.0/identity/conditionalAccess/policies?$select=id,displayName,state,conditions,grantControls,createdDateTime,modifiedDateTime'
        $caData = foreach ($p in $policies) {
            [PSCustomObject]@{
                DisplayName      = $p.displayName
                State            = $p.state
                IncludeUsers     = ($p.conditions.users.includeUsers -join '; ')
                IncludeGroups    = ($p.conditions.users.includeGroups -join '; ')
                IncludeApps      = ($p.conditions.applications.includeApplications -join '; ')
                GrantControls    = ($p.grantControls.builtInControls -join '; ')
                CreatedDateTime  = $p.createdDateTime
                ModifiedDateTime = $p.modifiedDateTime
                CollectDate      = (Get-Date -Format 'yyyy-MM-dd HH:mm')
            }
        }
        Export-ToCsv -Data $caData -FileName "EntraID_ConditionalAccess_$date.csv"
        Write-Log "  CA policies: $($caData.Count)" 'SUCCESS'
    }
    catch { Write-Log "  CA policies unavailable: $($_.Exception.Message)" 'WARN' }

    # --- Auth methods registration summary ---
    try {
        Write-Log "  Authentication method registration summary..."
        $headers = Get-GraphAuthHeaders
        $authReg = Invoke-GraphRequestWithThrottleHandling `
            -Uri 'https://graph.microsoft.com/v1.0/reports/authenticationMethods/usersRegisteredByFeature' `
            -Method 'GET' -Headers $headers
        if ($authReg.userRegistrationFeatureCounts) {
            $authData = foreach ($item in $authReg.userRegistrationFeatureCounts) {
                [PSCustomObject]@{
                    Feature     = $item.feature
                    UserCount   = $item.userCount
                    CollectDate = (Get-Date -Format 'yyyy-MM-dd HH:mm')
                }
            }
            Export-ToCsv -Data $authData -FileName "EntraID_AuthMethodsSummary_$date.csv"
        }
    }
    catch { Write-Log "  Auth method summary unavailable: $($_.Exception.Message)" 'WARN' }

    # --- M365 Apps aggregate usage counts (30 days) ---
    $null = Invoke-GraphReportRequest `
        -Uri "https://graph.microsoft.com/v1.0/reports/getM365AppUserCounts(period='D30')" `
        -FileName "EntraID_M365AppUserCounts_$date.csv"

    Write-Log "  Entra ID collection complete." 'SUCCESS'
}

#endregion

#region Workload: Security & Compliance

function New-CollectorStatusRow {
    param(
        [Parameter(Mandatory)] [string] $ObjectType,
        [Parameter(Mandatory)] [string] $Status,
        [Parameter()] [string] $Note = ''
    )

    return [PSCustomObject]@{
        ObjectType  = $ObjectType
        Status      = $Status
        Note        = $Note
        CollectDate = (Get-Date -Format 'yyyy-MM-dd HH:mm')
    }
}

function Invoke-GraphPagedRequestWithFallback {
    param(
        [Parameter(Mandatory)] [string[]] $Uris,
        [Parameter(Mandatory)] [string] $Description
    )

    $lastError = $null
    foreach ($uri in $Uris) {
        try {
            Write-Log "    Trying $Description endpoint: $uri" 'DEBUG'
            return @(Invoke-GraphPagedRequest -Uri $uri)
        }
        catch {
            $lastError = $_
            Write-Log "    $Description endpoint failed: $($_.Exception.Message)" 'DEBUG'
        }
    }

    if ($lastError) { throw $lastError }
    throw "No $Description endpoint returned data."
}

function Collect-PurviewDlpPolicies {
    try {
        Write-Log "  Purview DLP policies..."
        $policies = @(Get-DlpCompliancePolicy -ErrorAction Stop)
        $policyRows = foreach ($policy in $policies) {
            $row = ConvertTo-FlatCsvRow -InputObject $policy
            if (-not ($row.PSObject.Properties.Name -contains 'Workload')) {
                $row | Add-Member -NotePropertyName Workload -NotePropertyValue (($policy.Workload -join '; ')) -Force
            }
            if (-not ($row.PSObject.Properties.Name -contains 'Mode')) {
                $modeValue = if ($policy.Mode) { $policy.Mode } elseif ($policy.State) { $policy.State } else { '' }
                $row | Add-Member -NotePropertyName Mode -NotePropertyValue $modeValue -Force
            }
            $row
        }
        Export-ToCsv -Data $policyRows -FileName "Purview_DLPPolicies_$date.csv"

        $rules = @(Get-DlpComplianceRule -ErrorAction Stop)
        $ruleRows = foreach ($rule in $rules) {
            $row = ConvertTo-FlatCsvRow -InputObject $rule
            if (-not ($row.PSObject.Properties.Name -contains 'Policy')) {
                $row | Add-Member -NotePropertyName Policy -NotePropertyValue $rule.ParentPolicyName -Force
            }
            $row
        }
        Export-ToCsv -Data $ruleRows -FileName "Purview_DLPRules_$date.csv"

        $coverageRows = @($policies | ForEach-Object {
                $workloads = @($_.Workload)
                if ($workloads.Count -eq 0 -or [string]::IsNullOrWhiteSpace(($workloads -join ''))) {
                    $workloads = @('Unspecified')
                }
                foreach ($workload in $workloads) {
                    [PSCustomObject]@{
                        Workload    = $workload
                        Mode        = if ($_.Mode) { $_.Mode } elseif ($_.State) { $_.State } else { 'Unspecified' }
                        PolicyName  = $_.Name
                        Enabled     = $_.Enabled
                        CollectDate = (Get-Date -Format 'yyyy-MM-dd HH:mm')
                    }
                }
            })
        Export-ToCsv -Data $coverageRows -FileName "Purview_DLPCoverageSummary_$date.csv"
        Write-Log "  DLP policies: $($policies.Count); rules: $($rules.Count)" 'SUCCESS'
    }
    catch {
        Write-Log "  Purview DLP policies unavailable: $($_.Exception.Message)" 'WARN'
    }
}

function Collect-PurviewAuditConfigAndChangeEvents {
    try {
        Write-Log "  Purview audit configuration..."

        if (Get-Command -Name Get-AdminAuditLogConfig -ErrorAction SilentlyContinue) {
            $auditConfig = Get-AdminAuditLogConfig -ErrorAction Stop
            Export-ToCsv -Data @(ConvertTo-FlatCsvRow -InputObject $auditConfig) -FileName "Purview_AuditConfig_$date.csv"
        }
        else {
            Write-Log "  Get-AdminAuditLogConfig is unavailable in this session." 'WARN'
        }

        if (Get-Command -Name Get-UnifiedAuditLogRetentionPolicy -ErrorAction SilentlyContinue) {
            $retentionPolicies = @(Get-UnifiedAuditLogRetentionPolicy -ErrorAction Stop)
            $retentionRows = @(foreach ($policy in $retentionPolicies) { ConvertTo-FlatCsvRow -InputObject $policy })
            if ($retentionRows.Count -eq 0) {
                $retentionRows = @(New-CollectorStatusRow -ObjectType 'UnifiedAuditLogRetentionPolicy' -Status 'NoneReturned' -Note 'No explicit Unified Audit Log retention policies were returned.')
            }
            Export-ToCsv -Data $retentionRows -FileName "Purview_AuditRetentionPolicies_$date.csv"
        }
        else {
            Export-ToCsv -Data @(New-CollectorStatusRow -ObjectType 'UnifiedAuditLogRetentionPolicy' -Status 'CmdletUnavailable' -Note 'Get-UnifiedAuditLogRetentionPolicy is unavailable in this session.') -FileName "Purview_AuditRetentionPolicies_$date.csv"
        }

        if (Get-Command -Name Get-ProtectionAlert -ErrorAction SilentlyContinue) {
            $alerts = @(Get-ProtectionAlert -ErrorAction Stop)
            $alertRows = @(foreach ($alert in $alerts) { ConvertTo-FlatCsvRow -InputObject $alert })
            if ($alertRows.Count -eq 0) {
                $alertRows = @(New-CollectorStatusRow -ObjectType 'ProtectionAlert' -Status 'NoneReturned' -Note 'No Purview alert policies were returned.')
            }
            Export-ToCsv -Data $alertRows -FileName "Purview_AlertPolicies_$date.csv"
        }
        else {
            Export-ToCsv -Data @(New-CollectorStatusRow -ObjectType 'ProtectionAlert' -Status 'CmdletUnavailable' -Note 'Get-ProtectionAlert is unavailable in this session.') -FileName "Purview_AlertPolicies_$date.csv"
        }

        if (Get-Command -Name Search-UnifiedAuditLog -ErrorAction SilentlyContinue) {
            $startDate = (Get-Date).AddDays(-30)
            $endDate = Get-Date
            $changeOperations = @(
                'Add member to role',
                'Remove member from role',
                'Add eligible member to role',
                'Remove eligible member from role',
                'Add conditional access policy',
                'Update conditional access policy',
                'Delete conditional access policy',
                'Set-DlpCompliancePolicy',
                'Set-DlpComplianceRule',
                'Set-SPOTenant',
                'SharingSet',
                'SiteCollectionAdminAdded',
                'SiteCollectionAdminRemoved'
            )

            $changeEvents = @(Search-UnifiedAuditLog -StartDate $startDate -EndDate $endDate -Operations $changeOperations -ResultSize 5000 -ErrorAction Stop)
            $changeRows = @(foreach ($event in $changeEvents) {
                    $auditData = $null
                    try { $auditData = $event.AuditData | ConvertFrom-Json -ErrorAction Stop } catch {}
                    [PSCustomObject]@{
                        CreationDate = $event.CreationDate
                        Operation    = $event.Operations
                        Workload     = $event.Workload
                        UserIds      = $event.UserIds
                        ObjectId     = if ($auditData) { $auditData.ObjectId } else { '' }
                        ResultStatus = if ($auditData) { $auditData.ResultStatus } else { '' }
                        CollectDate  = (Get-Date -Format 'yyyy-MM-dd HH:mm')
                    }
                })
            if ($changeRows.Count -eq 0) {
                $changeRows = @(New-CollectorStatusRow -ObjectType 'UnifiedAuditLogChangeEvent' -Status 'NoneReturned' -Note 'No targeted configuration-change events were returned for the last 30 days.')
            }
            Export-ToCsv -Data $changeRows -FileName "Purview_AuditChangeEvents_$date.csv"

            $copilotEvents = @(Search-UnifiedAuditLog -StartDate $startDate -EndDate $endDate -FreeText 'Copilot' -ResultSize 5000 -ErrorAction SilentlyContinue)
            if ($copilotEvents.Count -gt 0) {
                $copilotRows = foreach ($event in $copilotEvents) { ConvertTo-FlatCsvRow -InputObject $event }
                Export-ToCsv -Data $copilotRows -FileName "Purview_CopilotAuditEvents_$date.csv"
            }
            else {
                Export-ToCsv -Data @([PSCustomObject]@{
                        EventType   = 'CopilotAuditEvents'
                        EventCount  = 0
                        StartDate   = $startDate
                        EndDate     = $endDate
                        CollectDate = (Get-Date -Format 'yyyy-MM-dd HH:mm')
                    }) -FileName "Purview_CopilotAuditEvents_$date.csv"
            }
        }
        else {
            Export-ToCsv -Data @(New-CollectorStatusRow -ObjectType 'UnifiedAuditLogChangeEvent' -Status 'CmdletUnavailable' -Note 'Search-UnifiedAuditLog is unavailable in this session.') -FileName "Purview_AuditChangeEvents_$date.csv"
            Export-ToCsv -Data @(New-CollectorStatusRow -ObjectType 'CopilotAuditEvents' -Status 'CmdletUnavailable' -Note 'Search-UnifiedAuditLog is unavailable in this session.') -FileName "Purview_CopilotAuditEvents_$date.csv"
        }
    }
    catch {
        Write-Log "  Purview audit configuration/change telemetry unavailable: $($_.Exception.Message)" 'WARN'
    }
}

function Collect-EntraPrivilegedAccess {
    try {
        Write-Log "  Entra privileged access / PIM..."

        $activeAssignments = Invoke-GraphPagedRequest -Uri "https://graph.microsoft.com/beta/roleManagement/directory/roleAssignmentScheduleInstances?`$expand=principal,roleDefinition&`$top=999"
        $eligibleAssignments = Invoke-GraphPagedRequest -Uri "https://graph.microsoft.com/beta/roleManagement/directory/roleEligibilityScheduleInstances?`$expand=principal,roleDefinition&`$top=999"

        $assignmentRows = @(
            foreach ($assignment in $activeAssignments) {
                [PSCustomObject]@{
                    AssignmentType = 'Active'
                    RoleName       = $assignment.roleDefinition.displayName
                    RoleId         = $assignment.roleDefinitionId
                    PrincipalId    = $assignment.principalId
                    PrincipalName  = $assignment.principal.displayName
                    PrincipalType  = $assignment.principal.'@odata.type'
                    MemberType     = $assignment.memberType
                    StartDateTime  = $assignment.startDateTime
                    EndDateTime    = $assignment.endDateTime
                    CollectDate    = (Get-Date -Format 'yyyy-MM-dd HH:mm')
                }
            }
            foreach ($assignment in $eligibleAssignments) {
                [PSCustomObject]@{
                    AssignmentType = 'Eligible'
                    RoleName       = $assignment.roleDefinition.displayName
                    RoleId         = $assignment.roleDefinitionId
                    PrincipalId    = $assignment.principalId
                    PrincipalName  = $assignment.principal.displayName
                    PrincipalType  = $assignment.principal.'@odata.type'
                    MemberType     = $assignment.memberType
                    StartDateTime  = $assignment.startDateTime
                    EndDateTime    = $assignment.endDateTime
                    CollectDate    = (Get-Date -Format 'yyyy-MM-dd HH:mm')
                }
            }
        )
        Export-ToCsv -Data $assignmentRows -FileName "EntraID_PIMRoleAssignments_$date.csv"

        $copilotAdminRows = @($assignmentRows | Where-Object { $_.RoleName -match '^Copilot Administrator$|Microsoft 365 Copilot' })
        if ($copilotAdminRows.Count -eq 0) {
            $copilotAdminRows = @([PSCustomObject]@{
                    AssignmentType = 'CopilotAdministrator'
                    RoleName       = 'Copilot Administrator'
                    PrincipalId    = ''
                    PrincipalName  = ''
                    PrincipalType  = ''
                    MemberType     = ''
                    StartDateTime  = ''
                    EndDateTime    = ''
                    Note           = 'No Copilot Administrator role assignments found in PIM active/eligible schedule instances.'
                    CollectDate    = (Get-Date -Format 'yyyy-MM-dd HH:mm')
                })
        }
        Export-ToCsv -Data $copilotAdminRows -FileName "EntraID_CopilotAdminAssignments_$date.csv"

        $startDate = (Get-Date).AddDays(-30).ToString('o')
        $roleAudits = Invoke-GraphPagedRequest -Uri "https://graph.microsoft.com/v1.0/auditLogs/directoryAudits?`$filter=activityDateTime ge $startDate and category eq 'RoleManagement'&`$top=999"
        $roleAuditRows = foreach ($audit in $roleAudits) { ConvertTo-FlatCsvRow -InputObject $audit }
        Export-ToCsv -Data $roleAuditRows -FileName "EntraID_RoleManagementAuditEvents_$date.csv"

        try {
            $accessReviews = Invoke-GraphPagedRequestWithFallback -Description 'access review definitions' -Uris @(
                "https://graph.microsoft.com/v1.0/identityGovernance/accessReviews/definitions?`$top=100",
                "https://graph.microsoft.com/beta/identityGovernance/accessReviews/definitions?`$top=100"
            )
            $accessReviewRows = @(foreach ($review in $accessReviews) { ConvertTo-FlatCsvRow -InputObject $review })
            if ($accessReviewRows.Count -eq 0) {
                $accessReviewRows = @(New-CollectorStatusRow -ObjectType 'AccessReviewDefinition' -Status 'NoneReturned' -Note 'No access review definitions were returned.')
            }
            Export-ToCsv -Data $accessReviewRows -FileName "EntraID_AccessReviewDefinitions_$date.csv"
        }
        catch {
            Export-ToCsv -Data @(New-CollectorStatusRow -ObjectType 'AccessReviewDefinition' -Status 'Unavailable' -Note $_.Exception.Message) -FileName "EntraID_AccessReviewDefinitions_$date.csv"
            Write-Log "  Access review definitions unavailable: $($_.Exception.Message)" 'WARN'
        }

        Write-Log "  PIM assignments collected: active=$($activeAssignments.Count), eligible=$($eligibleAssignments.Count)" 'SUCCESS'
    }
    catch {
        Write-Log "  Entra PIM/role governance unavailable: $($_.Exception.Message)" 'WARN'
    }
}

function Collect-EntraAppConsents {
    try {
        Write-Log "  Enterprise apps and OAuth consents..."

        $servicePrincipals = Invoke-GraphPagedRequest -Uri "https://graph.microsoft.com/v1.0/servicePrincipals?`$select=id,appId,displayName,accountEnabled,verifiedPublisher,appOwnerOrganizationId,signInAudience&`$top=999"
        $servicePrincipalById = @{}
        foreach ($sp in $servicePrincipals) { $servicePrincipalById[$sp.id] = $sp }

        $spRows = foreach ($sp in $servicePrincipals) {
            [PSCustomObject]@{
                Id                     = $sp.id
                AppId                  = $sp.appId
                DisplayName            = $sp.displayName
                AccountEnabled         = $sp.accountEnabled
                VerifiedPublisher      = $sp.verifiedPublisher.displayName
                AppOwnerOrganizationId = $sp.appOwnerOrganizationId
                SignInAudience         = $sp.signInAudience
                HasVerifiedPublisher   = -not [string]::IsNullOrWhiteSpace([string]$sp.verifiedPublisher.displayName)
                CollectDate            = (Get-Date -Format 'yyyy-MM-dd HH:mm')
            }
        }
        Export-ToCsv -Data $spRows -FileName "EntraID_ServicePrincipals_$date.csv"

        $highRiskScopes = @(
            'Mail.ReadWrite',
            'Mail.ReadWrite.Shared',
            'Files.ReadWrite.All',
            'Sites.ReadWrite.All',
            'Directory.ReadWrite.All',
            'User.ReadWrite.All',
            'Group.ReadWrite.All',
            'Application.ReadWrite.All',
            'RoleManagement.ReadWrite.Directory'
        )

        $delegatedGrants = Invoke-GraphPagedRequest -Uri "https://graph.microsoft.com/v1.0/oauth2PermissionGrants?`$top=999"
        $delegatedRows = foreach ($grant in $delegatedGrants) {
            $client = $servicePrincipalById[$grant.clientId]
            $resource = $servicePrincipalById[$grant.resourceId]
            $scopes = @($grant.scope -split ' ' | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
            [PSCustomObject]@{
                GrantType      = 'Delegated'
                ClientId       = $grant.clientId
                ClientApp      = $client.displayName
                ResourceId     = $grant.resourceId
                ResourceApp    = $resource.displayName
                ConsentType    = $grant.consentType
                PrincipalId    = $grant.principalId
                Scopes         = ($scopes -join '; ')
                HighRiskScopes = (($scopes | Where-Object { $_ -in $highRiskScopes }) -join '; ')
                CollectDate    = (Get-Date -Format 'yyyy-MM-dd HH:mm')
            }
        }
        Export-ToCsv -Data $delegatedRows -FileName "EntraID_OAuth2PermissionGrants_$date.csv"

        $appRoleRows = [System.Collections.Generic.List[object]]::new()
        foreach ($sp in $servicePrincipals) {
            try {
                $assignments = Invoke-GraphPagedRequest -Uri "https://graph.microsoft.com/v1.0/servicePrincipals/$($sp.id)/appRoleAssignments?`$top=999"
                foreach ($assignment in $assignments) {
                    $resource = $servicePrincipalById[$assignment.resourceId]
                    $appRoleRows.Add([PSCustomObject]@{
                            GrantType            = 'Application'
                            ClientId             = $sp.id
                            ClientApp            = $sp.displayName
                            ResourceId           = $assignment.resourceId
                            ResourceApp          = if ($resource) { $resource.displayName } else { $assignment.resourceDisplayName }
                            AppRoleId            = $assignment.appRoleId
                            PrincipalDisplayName = $assignment.principalDisplayName
                            CreatedDateTime      = $assignment.createdDateTime
                            CollectDate          = (Get-Date -Format 'yyyy-MM-dd HH:mm')
                        })
                }
            }
            catch {
                Write-Log "    App-role assignment lookup skipped for $($sp.displayName): $($_.Exception.Message)" 'DEBUG'
            }
        }
        Export-ToCsv -Data @($appRoleRows) -FileName "EntraID_AppRoleAssignments_$date.csv"
        Write-Log "  Service principals: $($servicePrincipals.Count); delegated grants: $($delegatedRows.Count); app-role grants: $($appRoleRows.Count)" 'SUCCESS'
    }
    catch {
        Write-Log "  App consent/OAuth inventory unavailable: $($_.Exception.Message)" 'WARN'
    }
}

function Collect-IntuneDeviceCompliance {
    try {
        Write-Log "  Intune device compliance..."

        $devices = Invoke-GraphPagedRequestWithFallback -Description 'Intune managed devices' -Uris @(
            "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices?`$top=999",
            "https://graph.microsoft.com/beta/deviceManagement/managedDevices?`$top=999",
            "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices?`$select=id,deviceName,operatingSystem,complianceState,managedDeviceOwnerType,userPrincipalName,lastSyncDateTime&`$top=999"
        )
        $deviceRows = foreach ($device in $devices) {
            [PSCustomObject]@{
                DeviceId          = $device.id
                DeviceName        = $device.deviceName
                OperatingSystem   = $device.operatingSystem
                ComplianceState   = $device.complianceState
                ManagementAgent   = if ($device.PSObject.Properties.Name -contains 'managementAgent') { $device.managementAgent } else { '' }
                OwnerType         = if ($device.PSObject.Properties.Name -contains 'managedDeviceOwnerType') { $device.managedDeviceOwnerType } elseif ($device.PSObject.Properties.Name -contains 'ownerType') { $device.ownerType } else { '' }
                UserPrincipalName = $device.userPrincipalName
                LastSyncDateTime  = $device.lastSyncDateTime
                CollectDate       = (Get-Date -Format 'yyyy-MM-dd HH:mm')
            }
        }
        if ($deviceRows.Count -eq 0) {
            $deviceRows = @(New-CollectorStatusRow -ObjectType 'ManagedDevice' -Status 'NoneReturned' -Note 'No Intune managed devices were returned.')
        }
        Export-ToCsv -Data $deviceRows -FileName "Intune_ManagedDevices_$date.csv"

        $summaryRows = @($devices | Group-Object -Property complianceState | ForEach-Object {
                New-ObjectCountSummary -ObjectType 'ManagedDevice' -Count $_.Count -AdditionalProperties @{
                    GroupingProperty = 'ComplianceState'
                    GroupingValue    = if ([string]::IsNullOrWhiteSpace([string]$_.Name)) { 'Unspecified' } else { $_.Name }
                }
            })
        if ($summaryRows.Count -eq 0) {
            $summaryRows = @(New-ObjectCountSummary -ObjectType 'ManagedDevice' -Count 0 -AdditionalProperties @{
                    GroupingProperty = 'ComplianceState'
                    GroupingValue    = 'None'
                })
        }
        Export-ToCsv -Data $summaryRows -FileName "Intune_DeviceComplianceSummary_$date.csv"

        $policies = Invoke-GraphPagedRequestWithFallback -Description 'Intune compliance policies' -Uris @(
            "https://graph.microsoft.com/v1.0/deviceManagement/deviceCompliancePolicies?`$top=999",
            "https://graph.microsoft.com/beta/deviceManagement/deviceCompliancePolicies?`$top=999"
        )
        $policyRows = @(foreach ($policy in $policies) { ConvertTo-FlatCsvRow -InputObject $policy })
        if ($policyRows.Count -eq 0) {
            $policyRows = @(New-CollectorStatusRow -ObjectType 'DeviceCompliancePolicy' -Status 'NoneReturned' -Note 'No Intune device compliance policies were returned.')
        }
        Export-ToCsv -Data $policyRows -FileName "Intune_DeviceCompliancePolicies_$date.csv"

        try {
            $stateSummary = $null
            foreach ($uri in @(
                    "https://graph.microsoft.com/v1.0/deviceManagement/deviceCompliancePolicyDeviceStateSummary",
                    "https://graph.microsoft.com/beta/deviceManagement/deviceCompliancePolicyDeviceStateSummary"
                )) {
                try {
                    $stateSummary = Invoke-GraphRequestWithThrottleHandling -Uri $uri -Method 'GET'
                    break
                }
                catch {
                    Write-Log "    Intune compliance state summary endpoint failed: $($_.Exception.Message)" 'DEBUG'
                }
            }
            if (-not $stateSummary) { throw 'No Intune compliance state summary endpoint returned data.' }
            Export-ToCsv -Data @(ConvertTo-FlatCsvRow -InputObject $stateSummary) -FileName "Intune_DeviceCompliancePolicyStateSummary_$date.csv"
        }
        catch {
            Export-ToCsv -Data @(New-CollectorStatusRow -ObjectType 'DeviceCompliancePolicyStateSummary' -Status 'Unavailable' -Note $_.Exception.Message) -FileName "Intune_DeviceCompliancePolicyStateSummary_$date.csv"
            Write-Log "  Intune compliance state summary unavailable: $($_.Exception.Message)" 'WARN'
        }

        try {
            $configProfiles = Invoke-GraphPagedRequestWithFallback -Description 'Intune device configuration profiles' -Uris @(
                "https://graph.microsoft.com/v1.0/deviceManagement/deviceConfigurations?`$top=999",
                "https://graph.microsoft.com/beta/deviceManagement/deviceConfigurations?`$top=999"
            )
            $configRows = @(foreach ($profile in $configProfiles) { ConvertTo-FlatCsvRow -InputObject $profile })
            if ($configRows.Count -eq 0) {
                $configRows = @(New-CollectorStatusRow -ObjectType 'DeviceConfigurationProfile' -Status 'NoneReturned' -Note 'No Intune device configuration profiles were returned.')
            }
            Export-ToCsv -Data $configRows -FileName "Intune_DeviceConfigurationProfiles_$date.csv"
        }
        catch {
            Export-ToCsv -Data @(New-CollectorStatusRow -ObjectType 'DeviceConfigurationProfile' -Status 'Unavailable' -Note $_.Exception.Message) -FileName "Intune_DeviceConfigurationProfiles_$date.csv"
            Write-Log "  Intune device configuration profiles unavailable: $($_.Exception.Message)" 'WARN'
        }

        Write-Log "  Intune managed devices: $($devices.Count); compliance policies: $($policies.Count)" 'SUCCESS'
    }
    catch {
        Write-Log "  Intune device compliance unavailable: $($_.Exception.Message)" 'WARN'
    }
}

function Collect-EntraGroupGovernance {
    try {
        Write-Log "  Ownerless M365 groups and naming policy..."

        $groups = Invoke-GraphPagedRequest -Uri "https://graph.microsoft.com/v1.0/groups?`$select=id,displayName,mail,groupTypes,createdDateTime&`$top=999"
        $m365Groups = @($groups | Where-Object { $_.groupTypes -contains 'Unified' })

        $ownerRows = foreach ($group in $m365Groups) {
            $owners = @(Invoke-GraphPagedRequest -Uri "https://graph.microsoft.com/v1.0/groups/$($group.id)/owners?`$select=id,displayName,userPrincipalName&`$top=999")
            [PSCustomObject]@{
                OwnerCount = $owners.Count
            }
        }

        $atRiskGroups = @($ownerRows | Where-Object { $_.OwnerCount -le 1 })
        $zeroOwnerCount = @($ownerRows | Where-Object { $_.OwnerCount -eq 0 }).Count
        $singleOwnerCount = @($ownerRows | Where-Object { $_.OwnerCount -eq 1 }).Count
        $ownerlessSummary = @([PSCustomObject]@{
                ObjectType                   = 'Microsoft365Group'
                TotalMicrosoft365Groups      = $m365Groups.Count
                GroupsWithZeroOwners         = $zeroOwnerCount
                GroupsWithOneOwner           = $singleOwnerCount
                GroupsWithOneOrFewerOwners   = $atRiskGroups.Count
                PercentWithOneOrFewerOwners  = if ($m365Groups.Count -gt 0) { [math]::Round(($atRiskGroups.Count / $m365Groups.Count) * 100, 2) } else { 0 }
                MinimumRecommendedOwnerCount = 2
                CollectDate                  = (Get-Date -Format 'yyyy-MM-dd HH:mm')
            })
        Export-ToCsv -Data $ownerlessSummary -FileName "EntraID_OwnerlessOrSingleOwnerGroups_$date.csv"

        try {
            $settings = Invoke-GraphPagedRequestWithFallback -Description 'group naming settings' -Uris @(
                "https://graph.microsoft.com/v1.0/groupSettings?`$top=999",
                "https://graph.microsoft.com/beta/groupSettings?`$top=999",
                "https://graph.microsoft.com/beta/settings?`$top=999",
                "https://graph.microsoft.com/beta/directory/settings?`$top=999"
            )
            $groupSettings = @($settings | Where-Object {
                    $_.displayName -eq 'Group.Unified' -or
                    $_.templateId -eq '62375ab9-6b52-47ed-826b-58e47e0e304b'
                })
            $settingRows = @(foreach ($setting in $groupSettings) { ConvertTo-FlatCsvRow -InputObject $setting })
            if ($settingRows.Count -eq 0) {
                $settingRows = @(New-CollectorStatusRow -ObjectType 'GroupNamingPolicy' -Status 'NoneReturned' -Note 'No Group.Unified tenant group settings were returned.')
            }
            Export-ToCsv -Data $settingRows -FileName "EntraID_GroupNamingPolicy_$date.csv"
            Write-Log "  M365 groups: $($m365Groups.Count); groups with <=1 owner: $($atRiskGroups.Count); group naming settings: $($settingRows.Count)" 'SUCCESS'
        }
        catch {
            Export-ToCsv -Data @(New-CollectorStatusRow -ObjectType 'GroupNamingPolicy' -Status 'Unavailable' -Note $_.Exception.Message) -FileName "EntraID_GroupNamingPolicy_$date.csv"
            Write-Log "  Group naming policy unavailable: $($_.Exception.Message)" 'WARN'
            Write-Log "  M365 groups: $($m365Groups.Count); groups with <=1 owner: $($atRiskGroups.Count)" 'SUCCESS'
        }
    }
    catch {
        Write-Log "  Group governance unavailable: $($_.Exception.Message)" 'WARN'
    }
}

function Collect-ServiceCommsAndGraphConnectors {
    try {
        Write-Log "  Message Center posts..."
        $messages = Invoke-GraphPagedRequest -Uri "https://graph.microsoft.com/v1.0/admin/serviceAnnouncement/messages?`$top=100"
        $messageRows = foreach ($message in $messages) {
            [PSCustomObject]@{
                Id            = $message.id
                Title         = $message.title
                Category      = $message.category
                Severity      = $message.severity
                StartDateTime = $message.startDateTime
                EndDateTime   = $message.endDateTime
                IsMajorChange = $message.isMajorChange
                Services      = ($message.services -join '; ')
                Tags          = ($message.tags -join '; ')
                CollectDate   = (Get-Date -Format 'yyyy-MM-dd HH:mm')
            }
        }
        Export-ToCsv -Data $messageRows -FileName "M365_MessageCenterPosts_$date.csv"
        Write-Log "  Message Center posts: $($messageRows.Count)" 'SUCCESS'
    }
    catch {
        Write-Log "  Message Center posts unavailable: $($_.Exception.Message)" 'WARN'
    }

    try {
        Write-Log "  Microsoft Graph connectors..."
        $connections = Invoke-GraphPagedRequestWithFallback -Description 'Graph connector inventory' -Uris @(
            "https://graph.microsoft.com/v1.0/external/connections?`$top=999",
            "https://graph.microsoft.com/beta/external/connections?`$top=999"
        )
        $connectionRows = @(foreach ($connection in $connections) { ConvertTo-FlatCsvRow -InputObject $connection })
        if ($connectionRows.Count -eq 0) {
            $connectionRows = @(New-CollectorStatusRow -ObjectType 'ExternalConnection' -Status 'NoneReturned' -Note 'No Microsoft Graph connectors were returned.')
        }
        Export-ToCsv -Data $connectionRows -FileName "MicrosoftSearch_GraphConnectors_$date.csv"
        Write-Log "  Graph connectors: $($connectionRows.Count)" 'SUCCESS'
    }
    catch {
        Export-ToCsv -Data @(New-CollectorStatusRow -ObjectType 'ExternalConnection' -Status 'Unavailable' -Note $_.Exception.Message) -FileName "MicrosoftSearch_GraphConnectors_$date.csv"
        Write-Log "  Graph connector inventory unavailable: $($_.Exception.Message)" 'WARN'
    }
}

function Collect-SecurityData {
    Write-Log ""
    Write-Log "==  Security & Compliance  ==" 'SUCCESS'
    $ippsConnected = $false

    # --- Secure Score ---
    try {
        Write-Log "  Secure Score..."
        $scoreResponse = Invoke-GraphRequestWithThrottleHandling -Uri 'https://graph.microsoft.com/v1.0/security/secureScores?$top=1' -Method 'GET'
        $scores = @($scoreResponse.value)
        if ($scores -and $scores.Count -gt 0) {
            $ss = $scores[0]
            $avgScore = ($ss.averageComparativeScores | Where-Object { $_.basis -eq 'AllTenants' } |
                Select-Object -ExpandProperty averageScore -ErrorAction SilentlyContinue)
            Export-ToCsv -Data @([PSCustomObject]@{
                    CurrentScore       = $ss.currentScore
                    MaxScore           = $ss.maxScore
                    PercentScore       = if ($ss.maxScore -gt 0) { [math]::Round($ss.currentScore / $ss.maxScore * 100, 1) } else { 0 }
                    AvgAllTenantsScore = $avgScore
                    ActiveUserCount    = $ss.activeUserCount
                    EnabledServices    = ($ss.enabledServices -join '; ')
                    CreatedDateTime    = $ss.createdDateTime
                    CollectDate        = (Get-Date -Format 'yyyy-MM-dd HH:mm')
                }) -FileName "Security_SecureScore_$date.csv"
            Write-Log "  Secure Score: $($ss.currentScore) / $($ss.maxScore)  ($([math]::Round($ss.currentScore/$ss.maxScore*100,1))%)" 'SUCCESS'
        }
    }
    catch { Write-Log "  Secure Score unavailable: $($_.Exception.Message)" 'WARN' }

    # --- Secure Score control profiles ---
    try {
        Write-Log "  Secure Score control profiles..."
        $controls = Invoke-GraphPagedRequest -Uri `
            'https://graph.microsoft.com/v1.0/security/secureScoreControlProfiles?$select=id,title,controlCategory,implementationStatus,maxScore,rank,threats,tier,remediation'
        $controlData = foreach ($c in $controls) {
            [PSCustomObject]@{
                Title                = $c.title
                ControlCategory      = $c.controlCategory
                ImplementationStatus = $c.implementationStatus
                MaxScore             = $c.maxScore
                Rank                 = $c.rank
                Tier                 = $c.tier
                Threats              = ($c.threats -join '; ')
                Remediation          = $c.remediation
                CollectDate          = (Get-Date -Format 'yyyy-MM-dd HH:mm')
            }
        }
        Export-ToCsv -Data $controlData -FileName "Security_SecureScoreControls_$date.csv"
        Write-Log "  Control profiles: $($controlData.Count)" 'SUCCESS'
    }
    catch { Write-Log "  Secure Score controls unavailable: $($_.Exception.Message)" 'WARN' }

    try {
        Write-Log "  Connecting Security & Compliance PowerShell module..."
        Connect-ToSecurityCompliance
        $ippsConnected = $true
    }
    catch {
        Write-Log "  Security & Compliance PowerShell unavailable: $($_.Exception.Message)" 'WARN'
    }

    # --- Sensitivity labels ---
    try {
        Write-Log "  Sensitivity labels (Get-Label)..."
        if (-not $ippsConnected) {
            throw 'Security & Compliance PowerShell is not connected.'
        }
        $labels = @(Get-Label)
        $labelData = foreach ($l in $labels) {
            $flat = ConvertTo-FlatCsvRow -InputObject $l
            if (-not ($flat.PSObject.Properties.Name -contains 'Name')) {
                $flat | Add-Member -NotePropertyName Name -NotePropertyValue $l.Name -Force
            }
            $flat
        }
        Export-ToCsv -Data $labelData -FileName "Security_SensitivityLabels_$date.csv"
        Write-Log "  Sensitivity labels: $($labelData.Count)" 'SUCCESS'
    }
    catch {
        Write-Log "  Native sensitivity labels unavailable: $($_.Exception.Message)" 'WARN'
        try {
            $labels = Invoke-GraphPagedRequest -Uri `
                'https://graph.microsoft.com/v1.0/security/informationProtection/sensitivityLabels?$select=id,name,description,isActive,sensitivity,color'
            $labelData = foreach ($l in $labels) {
                [PSCustomObject]@{
                    Name        = $l.name
                    Id          = $l.id
                    Description = $l.description
                    IsActive    = $l.isActive
                    Sensitivity = $l.sensitivity
                    Color       = $l.color
                    CollectDate = (Get-Date -Format 'yyyy-MM-dd HH:mm')
                }
            }
            Export-ToCsv -Data $labelData -FileName "Security_SensitivityLabels_$date.csv"
            Write-Log "  Sensitivity labels via Graph fallback: $($labelData.Count)" 'SUCCESS'
        }
        catch { Write-Log "  Sensitivity labels unavailable: $($_.Exception.Message)" 'WARN' }
    }

    # --- Named locations (Conditional Access) ---
    try {
        Write-Log "  Named locations..."
        $locations = Invoke-GraphPagedRequest -Uri `
            'https://graph.microsoft.com/v1.0/identity/conditionalAccess/namedLocations?$select=id,displayName,createdDateTime,modifiedDateTime'
        $locData = foreach ($l in $locations) {
            [PSCustomObject]@{
                DisplayName      = $l.displayName
                Id               = $l.id
                CreatedDateTime  = $l.createdDateTime
                ModifiedDateTime = $l.modifiedDateTime
                CollectDate      = (Get-Date -Format 'yyyy-MM-dd HH:mm')
            }
        }
        Export-ToCsv -Data $locData -FileName "Security_NamedLocations_$date.csv"
        Write-Log "  Named locations: $($locData.Count)" 'SUCCESS'
    }
    catch { Write-Log "  Named locations unavailable: $($_.Exception.Message)" 'WARN' }

    if ($ippsConnected) {
        Collect-PurviewDlpPolicies
        Collect-PurviewAuditConfigAndChangeEvents
    }
    else {
        Write-Log "  Skipping Purview DLP and audit collectors because Security & Compliance PowerShell is not connected." 'WARN'
    }

    Collect-EntraPrivilegedAccess
    Collect-EntraAppConsents
    Collect-IntuneDeviceCompliance
    Collect-EntraGroupGovernance
    Collect-ServiceCommsAndGraphConnectors

    if ($ippsConnected) {
        try { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue | Out-Null } catch {}
        Write-Log "  Disconnected from Security & Compliance PowerShell" 'DEBUG'
    }

    Write-Log "  Security & Compliance collection complete." 'SUCCESS'
}

#endregion

#region Workload: Power Platform

function Collect-PowerPlatformData {
    Write-Log ""
    Write-Log "==  Power Platform  ==" 'SUCCESS'

    Write-Log "  Power BI workspaces..."
    try {
        Connect-ToPowerBIService
        Import-Module MicrosoftPowerBIMgmt.Workspaces -ErrorAction Stop
        $pbiWorkspaces = @(Get-PowerBIWorkspace -Scope Organization -All -ErrorAction Stop)

        $pbiWsData = foreach ($ws in $pbiWorkspaces) {
            [PSCustomObject]@{
                Id                    = $ws.id
                Name                  = $ws.name
                Type                  = $ws.type
                State                 = $ws.state
                IsOnDedicatedCapacity = $ws.isOnDedicatedCapacity
                CapacityId            = $ws.capacityId
                CollectDate           = (Get-Date -Format 'yyyy-MM-dd HH:mm')
            }
        }
        Export-ToCsv -Data $pbiWsData -FileName "PowerPlatform_PowerBI_Workspaces_$date.csv"
        Write-Log "  Power BI workspaces: $($pbiWorkspaces.Count)" 'SUCCESS'
    }
    catch {
        Write-Log "  Power BI native module collection unavailable: $($_.Exception.Message)" 'WARN'
        try {
            $pbiToken = Get-AppTokenForScope -Scope 'https://analysis.windows.net/powerbi/api/.default'
            if ([string]::IsNullOrWhiteSpace($pbiToken)) {
                Write-Log "  Power BI REST fallback skipped: app-only token acquisition is not configured for interactive auth mode." 'WARN'
                throw 'SkipPbiRestFallback'
            }
            $pbiHeader = @{ Authorization = "Bearer $pbiToken" }
            $pbiWorkspaces = [System.Collections.Generic.List[object]]::new()
            $pbiNextLink = 'https://api.powerbi.com/v1.0/myorg/admin/groups?$top=5000'
            do {
                $pbiResp = Invoke-RestMethod -Uri $pbiNextLink -Method 'GET' `
                    -Headers $pbiHeader -ErrorAction Stop -Verbose:$false
                if ($pbiResp.value) { $pbiWorkspaces.AddRange([object[]]$pbiResp.value) }
                $pbiNextLink = $pbiResp.'@odata.nextLink'
            } while ($pbiNextLink)

            $pbiWsData = foreach ($ws in $pbiWorkspaces) {
                [PSCustomObject]@{
                    Id                    = $ws.id
                    Name                  = $ws.name
                    Type                  = $ws.type
                    State                 = $ws.state
                    IsOnDedicatedCapacity = $ws.isOnDedicatedCapacity
                    CapacityId            = $ws.capacityId
                    CollectDate           = (Get-Date -Format 'yyyy-MM-dd HH:mm')
                }
            }
            Export-ToCsv -Data $pbiWsData -FileName "PowerPlatform_PowerBI_Workspaces_$date.csv"
            Write-Log "  Power BI workspaces via REST fallback: $($pbiWorkspaces.Count)" 'SUCCESS'
        }
        catch {
            if ($_.Exception.Message -eq 'SkipPbiRestFallback') { }
            else {
                $sc = $null
                if ($_.Exception.Response) { $sc = [int]$_.Exception.Response.StatusCode }
                if ($sc -in @(401, 403)) {
                    Write-Log "  Power BI workspaces unavailable ($sc — access denied)." 'WARN'
                    Write-Log "    Ensure your account has 'Power BI Administrator' or 'Global Administrator' role." 'WARN'
                    Write-Log "    In Power BI Admin portal: Tenant settings -> Allow users to use Power BI APIs." 'WARN'
                }
                else {
                    Write-Log "  Power BI workspaces unavailable: $($_.Exception.Message)" 'WARN'
                }
            }
        }
    }

    # --- Power BI capacities ---
    try {
        Import-Module MicrosoftPowerBIMgmt.Capacities -ErrorAction Stop
        $capacities = @(Get-PowerBICapacity -Scope Organization -ErrorAction Stop)
        if ($capacities.Count -gt 0) {
            $capData = foreach ($c in $capacities) {
                [PSCustomObject]@{
                    Id          = $c.id
                    DisplayName = $c.displayName
                    Sku         = $c.sku
                    State       = $c.state
                    Region      = $c.region
                    CollectDate = (Get-Date -Format 'yyyy-MM-dd HH:mm')
                }
            }
            Export-ToCsv -Data $capData -FileName "PowerPlatform_PowerBI_Capacities_$date.csv"
            Write-Log "  Power BI capacities: $($capacities.Count)" 'SUCCESS'
        }
        else {
            Write-Log "  Power BI capacities: none (no Premium/Fabric capacity in tenant)" 'SUCCESS'
        }
    }
    catch {
        try {
            $pbiToken = Get-AppTokenForScope -Scope 'https://analysis.windows.net/powerbi/api/.default'
            if ([string]::IsNullOrWhiteSpace($pbiToken)) {
                Write-Log "  Power BI capacities REST fallback skipped: app-only token acquisition is not configured for interactive auth mode." 'WARN'
                throw 'SkipPbiCapacityRestFallback'
            }
            $pbiHeader = @{ Authorization = "Bearer $pbiToken" }
            $capResp = Invoke-RestMethod -Uri 'https://api.powerbi.com/v1.0/myorg/admin/capacities' `
                -Method 'GET' -Headers $pbiHeader -ErrorAction Stop -Verbose:$false
            if ($capResp.value -and $capResp.value.Count -gt 0) {
                $capData = foreach ($c in $capResp.value) {
                    [PSCustomObject]@{
                        Id          = $c.id
                        DisplayName = $c.displayName
                        Sku         = $c.sku
                        State       = $c.state
                        Region      = $c.region
                        CollectDate = (Get-Date -Format 'yyyy-MM-dd HH:mm')
                    }
                }
                Export-ToCsv -Data $capData -FileName "PowerPlatform_PowerBI_Capacities_$date.csv"
                Write-Log "  Power BI capacities via REST fallback: $($capResp.value.Count)" 'SUCCESS'
            }
        }
        catch { } # advisory already printed by the workspaces block above
    }

    # --- Power Platform environments (isolated first; native + REST fallback) ---
    # Requires: service principal assigned 'Power Platform Administrator' role in Azure AD
    Write-Log "  Power Platform environments..."

    $envCollected = $false

    # Prefer an isolated pwsh process first to avoid in-session MSAL assembly conflicts
    # with modules already loaded for other workloads (for example Power BI modules).
    try {
        Write-Log '  Collecting Power Platform environments in isolated pwsh process...' 'DEBUG'
        $isolatedEnvs = @(Get-PowerPlatformEnvironmentsIsolated -TenantId $tenantId)
        $isolatedEnvData = foreach ($env in $isolatedEnvs) {
            [PSCustomObject]@{
                Name              = $env.Name
                DisplayName       = $env.DisplayName
                Location          = $env.Location
                Type              = $env.Type
                ProvisioningState = $env.ProvisioningState
                IsDefault         = $env.IsDefault
                CreatedTime       = $env.CreatedTime
                CollectDate       = (Get-Date -Format 'yyyy-MM-dd HH:mm')
            }
        }
        if ($isolatedEnvData.Count -gt 0) {
            Export-ToCsv -Data $isolatedEnvData -FileName "PowerPlatform_Environments_$date.csv"
            Write-Log "  Power Platform environments (isolated): $($isolatedEnvData.Count)" 'SUCCESS'
            $envCollected = $true
        }
        else {
            Write-Log '  Isolated Power Platform collection returned no environments.' 'WARN'
        }
    }
    catch {
        Write-Log "  Isolated Power Platform collection unavailable: $($_.Exception.Message)" 'WARN'
    }

    if (-not $envCollected) {
        try {
            Connect-ToPowerPlatformAdmin
            $envs = @(Get-AdminPowerAppEnvironment -ErrorAction Stop)
            $envData = foreach ($env in $envs) {
                [PSCustomObject]@{
                    Name              = $env.name
                    DisplayName       = if ($env.DisplayName) { $env.DisplayName } else { $env.properties.displayName }
                    Location          = $env.location
                    Type              = if ($env.EnvironmentType) { $env.EnvironmentType } else { $env.properties.environmentSku }
                    ProvisioningState = if ($env.ProvisioningState) { $env.ProvisioningState } else { $env.properties.provisioningState }
                    IsDefault         = if ($null -ne $env.IsDefault) { $env.IsDefault } else { $env.properties.isDefault }
                    CreatedTime       = if ($env.CreatedTime) { $env.CreatedTime } else { $env.properties.createdTime }
                    CollectDate       = (Get-Date -Format 'yyyy-MM-dd HH:mm')
                }
            }
            Export-ToCsv -Data $envData -FileName "PowerPlatform_Environments_$date.csv"
            Write-Log "  Power Platform environments: $($envs.Count)" 'SUCCESS'
            $envCollected = $true
        }
        catch {
            Write-Log "  Power Platform native module collection unavailable: $($_.Exception.Message)" 'WARN'
        }
    }

    if (-not $envCollected) {
        try {
            $ppToken = Get-AppTokenForScope -Scope 'https://service.powerapps.com/.default'
            if ([string]::IsNullOrWhiteSpace($ppToken)) {
                Write-Log "  Power Platform REST fallback skipped: app-only token acquisition is not configured for interactive auth mode." 'WARN'
                throw 'SkipPowerPlatformRestFallback'
            }
            $ppHeader = @{ Authorization = "Bearer $ppToken" }
            $ppResp = Invoke-RestMethod `
                -Uri 'https://api.bap.microsoft.com/providers/Microsoft.BusinessAppPlatform/scopes/admin/environments?api-version=2021-04-01' `
                -Method 'GET' -Headers $ppHeader -ErrorAction Stop -Verbose:$false
            $envData = foreach ($env in $ppResp.value) {
                [PSCustomObject]@{
                    Name              = $env.name
                    DisplayName       = $env.properties.displayName
                    Location          = $env.location
                    Type              = $env.properties.environmentSku
                    ProvisioningState = $env.properties.provisioningState
                    IsDefault         = $env.properties.isDefault
                    CreatedTime       = $env.properties.createdTime
                    CollectDate       = (Get-Date -Format 'yyyy-MM-dd HH:mm')
                }
            }
            Export-ToCsv -Data $envData -FileName "PowerPlatform_Environments_$date.csv"
            Write-Log "  Power Platform environments via REST fallback: $($ppResp.value.Count)" 'SUCCESS'
        }
        catch {
            if ($_.Exception.Message -eq 'SkipPowerPlatformRestFallback') { }
            else {
                $sc = $null
                if ($_.Exception.Response) { $sc = [int]$_.Exception.Response.StatusCode }
                if ($sc -in @(401, 403)) {
                    Write-Log "  Power Platform environments unavailable ($sc — access denied)." 'WARN'
                    Write-Log "    Ensure your account has 'Power Platform Administrator' role in Entra ID > Roles." 'WARN'
                }
                else {
                    Write-Log "  Power Platform environments unavailable: $($_.Exception.Message)" 'WARN'
                }
            }
        }
    }

    Write-Log "  Power Platform collection complete." 'SUCCESS'
}

#endregion

#region Workload: Information Barriers

function Collect-InformationBarriersData {
    Write-Log ""
    Write-Log "==  Information Barriers  ==" 'SUCCESS'
    Write-Log "  Collecting Information Barriers configuration and SharePoint IB settings..."

    if ([string]::IsNullOrWhiteSpace($tenantUrl)) {
        Write-Log "  Skipping Information Barriers - `$tenantUrl not configured in the Configuration section." 'WARN'
        return
    }

    $adminUrl = $tenantUrl.TrimEnd('/') -replace 'https://([^.]+)\.sharepoint\.com.*', 'https://$1-admin.sharepoint.com'
    $ippsConnected = $false
    $spoConnected = $false
    $segmentNameByGuid = @{}

    function Get-IbGuidFromValue {
        param([Parameter()] [object] $Value)

        $text = [string]$Value
        if ([string]::IsNullOrWhiteSpace($text)) { return '' }

        # Accept either raw GUID or identity forms such as tenant\<guid>.
        if ($text -match '([0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12})') {
            return $Matches[1]
        }

        return ''
    }

    try {
        # Step 4: Org-level IB state from Exchange Online + Security & Compliance PowerShell.
        $exoConnected = $false
        try {
            Write-Log "  Connecting to Exchange Online (interactive)..."
            Connect-ToExchangeOnline
            $exoConnected = $true

            try {
                $orgConfigRaw = Get-OrganizationConfig -ErrorAction Stop
                $orgConfig = @([PSCustomObject]@{
                        DefaultPublicFolderProhibitPostQuota    = $orgConfigRaw.DefaultPublicFolderProhibitPostQuota
                        DistributionGroupDefaultOU              = $orgConfigRaw.DistributionGroupDefaultOU
                        DistributionGroupNameBlockedWordsList   = $orgConfigRaw.DistributionGroupNameBlockedWordsList
                        DistributionGroupNamingPolicy           = $orgConfigRaw.DistributionGroupNamingPolicy
                        UpgradeIBInProgress                     = $orgConfigRaw.UpgradeIBInProgress
                        VisibleMeetingUpdateProperties          = $orgConfigRaw.VisibleMeetingUpdateProperties
                        MaxInformationBarrierSegmentsLegacy     = $orgConfigRaw.MaxInformationBarrierSegmentsLegacy
                        MaxInformationBarrierSegments           = $orgConfigRaw.MaxInformationBarrierSegments
                        MaxInformationBarrierBridges            = $orgConfigRaw.MaxInformationBarrierBridges
                        InformationBarriersManagementEnabled    = $orgConfigRaw.InformationBarriersManagementEnabled
                        InformationBarriersEnforcementEnabled   = $orgConfigRaw.InformationBarriersEnforcementEnabled
                        InformationBarriersRestrictPeopleSearch = $orgConfigRaw.InformationBarriersRestrictPeopleSearch
                        InformationBarrierMode                  = $orgConfigRaw.InformationBarrierMode
                    })
                Export-ToCsv -Data $orgConfig -FileName "IB_OrgConfig_$date.csv"
                Write-Log "  IB org configuration collected." 'SUCCESS'
            }
            catch { Write-Log "  IB org configuration unavailable: $($_.Exception.Message)" 'WARN' }

            try {
                $policyConfigRaw = Get-PolicyConfig -ErrorAction Stop
                $policyConfig = @([PSCustomObject]@{
                        SensitiveInformationScanTimeWindowExo     = $policyConfigRaw.SensitiveInformationScanTimeWindowExo
                        InformationBarrierMode                    = $policyConfigRaw.InformationBarrierMode
                        InformationBarrierPeopleSearchRestriction = $policyConfigRaw.InformationBarrierPeopleSearchRestriction
                    })
                Export-ToCsv -Data $policyConfig -FileName "IB_PolicyConfig_$date.csv"
                Write-Log "  IB policy configuration collected." 'SUCCESS'
            }
            catch { Write-Log "  IB policy configuration unavailable: $($_.Exception.Message)" 'WARN' }
        }
        catch {
            Write-Log "  IB org/policy configuration unavailable: $($_.Exception.Message)" 'WARN'
        }

        # Now connect to Security & Compliance for segments and policies
        try {
            Write-Log "  Connecting to Security & Compliance PowerShell (interactive)..."
            Connect-ToSecurityCompliance
            $ippsConnected = $true

            try {
                Write-Log "  Retrieving IB segments (Get-ExoInformationBarrierSegment)..."
                $segmentsRaw = @(Get-ExoInformationBarrierSegment -ErrorAction Stop)
                $segments = @($segmentsRaw | ForEach-Object {
                        $segmentGuid = Get-IbGuidFromValue -Value $_.Name
                        $segmentName = if ([string]::IsNullOrWhiteSpace([string]$_.DisplayName)) { $segmentGuid } else { [string]$_.DisplayName }
                        if (-not [string]::IsNullOrWhiteSpace($segmentGuid)) {
                            $segmentNameByGuid[$segmentGuid] = $segmentName
                        }
                        [PSCustomObject]@{
                            Name            = $segmentName
                            UserGroupFilter = $_.MembershipFilter
                            ExoSegmentId    = $segmentGuid
                        }
                    })
                Export-ToCsv -Data $segments -FileName "IB_Segments_$date.csv"
                Write-Log "  IB segments: $($segments.Count)" 'SUCCESS'
            }
            catch { Write-Log "  IB segments unavailable: $($_.Exception.Message)" 'WARN' }

            try {
                Write-Log "  Retrieving IB policies (Get-ExoInformationBarrierPolicy)..."
                $policiesRaw = @(Get-ExoInformationBarrierPolicy -ErrorAction Stop)
                $policies = @($policiesRaw | ForEach-Object {
                        $assignedGuid = Get-IbGuidFromValue -Value $_.SegmentId
                        $relatedGuid = Get-IbGuidFromValue -Value $_.SegmentRelationship

                        $assignedSegmentName = if ($segmentNameByGuid.ContainsKey($assignedGuid)) { $segmentNameByGuid[$assignedGuid] } else { $assignedGuid }
                        $relatedSegmentName = if ($segmentNameByGuid.ContainsKey($relatedGuid)) { $segmentNameByGuid[$relatedGuid] } else { $relatedGuid }

                        $segmentsBlocked = ''
                        $segmentsAllowed = ''
                        if ([string]$_.RelationshipType -eq 'Block') {
                            $segmentsBlocked = $relatedSegmentName
                        }
                        elseif ([string]$_.RelationshipType -eq 'Allow') {
                            $segmentsAllowed = $relatedSegmentName
                        }

                        $policyName = [string]$_.DisplayName
                        if ([string]::IsNullOrWhiteSpace($policyName)) {
                            if ([string]$_.RelationshipType -eq 'Block') {
                                $policyName = "$assignedSegmentName - Blocks - $relatedSegmentName"
                            }
                            elseif ([string]$_.RelationshipType -eq 'Allow') {
                                $policyName = "$assignedSegmentName - Allows - $relatedSegmentName"
                            }
                            else {
                                $policyName = [string]$_.Id
                            }
                        }

                        [PSCustomObject]@{
                            Name            = $policyName
                            AssignedSegment = $assignedSegmentName
                            SegmentsBlocked = $segmentsBlocked
                            SegmentsAllowed = $segmentsAllowed
                            ExoPolicyId     = [string]$_.Id
                            State           = $_.State
                            Guid            = $_.Guid
                            BlockVisibility = $_.BlockVisibility
                        }
                    })
                Export-ToCsv -Data $policies -FileName "IB_Policies_$date.csv"
                Write-Log "  IB policies: $($policies.Count)" 'SUCCESS'
            }
            catch { Write-Log "  IB policies unavailable: $($_.Exception.Message)" 'WARN' }
        }
        catch {
            Write-Log "  IB segments and policies collection unavailable: $($_.Exception.Message)" 'WARN'
        }

        # Step 5: SharePoint Online IB-related settings
        try {
            if ($PSVersionTable.PSVersion.Major -gt 5) {
                Import-Module Microsoft.Online.SharePoint.PowerShell `
                    -UseWindowsPowerShell -DisableNameChecking -ErrorAction Stop
            }
            else {
                Import-Module Microsoft.Online.SharePoint.PowerShell `
                    -DisableNameChecking -ErrorAction Stop
            }

            Write-Log "  Connecting to SharePoint Online admin ($adminUrl)..."
            Connect-SPOService -Url $adminUrl -ErrorAction Stop
            $spoConnected = $true
            Write-Log "  Connected to SharePoint Online admin." 'SUCCESS'

            try {
                $spoTenant = Get-SPOTenant
                $spoIbSettings = @([PSCustomObject]@{
                        DefaultOneDriveInformationBarrierMode     = $spoTenant.DefaultOneDriveInformationBarrierMode
                        InformationBarriersSuspension             = $spoTenant.InformationBarriersSuspension
                        IBImplicitGroupBased                      = $spoTenant.IBImplicitGroupBased
                        ShowPeoplePickerGroupSuggestionsForIB     = $spoTenant.ShowPeoplePickerGroupSuggestionsForIB
                        BypassDomainMatching                      = $spoTenant.BypassDomainMatching
                        AppBypassInformationBarriers              = $spoTenant.AppBypassInformationBarriers
                        InformationBarriersInvitationRestrictions = $spoTenant.InformationBarriersInvitationRestrictions
                        CollectDate                               = (Get-Date -Format 'yyyy-MM-dd HH:mm')
                    })
                Export-ToCsv -Data $spoIbSettings -FileName "IB_SPOSettings_$date.csv"
                Write-Log "  IB SharePoint settings collected." 'SUCCESS'
            }
            catch { Write-Log "  IB SharePoint settings unavailable: $($_.Exception.Message)" 'WARN' }
        }
        catch {
            Write-Log "  IB SharePoint settings collection unavailable: $($_.Exception.Message)" 'WARN'
        }
    }
    finally {
        if ($exoConnected) {
            try { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue } catch {}
        }
        if ($ippsConnected) {
            try { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue } catch {}
        }
        if ($spoConnected) {
            try { Disconnect-SPOService -ErrorAction SilentlyContinue } catch {}
        }
    }

    Write-Log "  Information Barriers collection complete." 'SUCCESS'
}

#endregion

#region Summary Report

function Write-SummaryReport {
    param(
        [string[]] $CompletedWorkloads
    )

    Write-Host ""
    Write-Host "  ============================================================" -ForegroundColor Cyan
    Write-Host "   MCA Assessment Scan — Summary" -ForegroundColor White
    Write-Host "   Completed: $(Get-Date -Format 'yyyy-MM-dd HH:mm')" -ForegroundColor DarkGray
    Write-Host "  ============================================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "  Workloads scanned:" -ForegroundColor Yellow
    foreach ($wl in $CompletedWorkloads) {
        Write-Host "    [+] $wl" -ForegroundColor Green
    }
    Write-Host ""

    $files = Get-ChildItem -Path $OutputFolder -Filter "*.csv" -ErrorAction SilentlyContinue |
    Sort-Object Name
    if ($files) {
        Write-Host "  Output files ($($files.Count)):" -ForegroundColor Yellow
        foreach ($f in $files) {
            $size = "{0:N1} KB" -f ($f.Length / 1KB)
            Write-Host "    $($f.Name.PadRight(65)) $size" -ForegroundColor White
        }
    }
    Write-Host ""
    Write-Host "  Output folder : $OutputFolder" -ForegroundColor Cyan
    Write-Host "  Run log       : $RunLog" -ForegroundColor Cyan
    Write-Host ""
}

#endregion Summary Report

#region Main
##############################################################
#                      MAIN EXECUTION                        #
##############################################################

try {
    # Always require tenantId and tenantUrl
    if ([string]::IsNullOrWhiteSpace($tenantId)) { throw 'TenantId is empty. Fill in the Configuration section.' }
    if ([string]::IsNullOrWhiteSpace($tenantUrl)) { throw 'tenantUrl is empty. Fill in the Configuration section.' }

    # Show workload menu — auth validation happens after so SPO/ODB can run without app credentials
    $choice = Show-WorkloadMenu

    if ($choice -eq 'Q') {
        Write-Host "`n  Exiting." -ForegroundColor Red
        exit 0
    }

    Show-Banner
    Write-Log "MCA Data Collection Scanner starting."
    Write-Log "Workload selection : $choice"
    Write-Log "Tenant             : $tenantId"
    Write-Log "Output folder      : $OutputFolder"

    $runAll = ($choice -eq 'A')
    $needsTeamsCompatibility = ($runAll -or $choice -eq '3')
    Initialize-ModuleCompatibility -NeedsTeamsCompatibility $needsTeamsCompatibility

    $completed = [System.Collections.Generic.List[string]]::new()

    # Dispatch to workload collector(s)
    if ($runAll -or $choice -eq '1') { Collect-SharePointData  ; $completed.Add('SharePoint Online') }
    if ($runAll -or $choice -eq '2') { Collect-ExchangeData    ; $completed.Add('Exchange Online') }
    if ($runAll -or $choice -eq '3') { Collect-TeamsData       ; $completed.Add('Microsoft Teams') }
    if ($runAll -or $choice -eq '4') { Collect-OneDriveData    ; $completed.Add('OneDrive for Business') }
    if ($runAll -or $choice -eq '5') { Collect-EntraIDData     ; $completed.Add('Entra ID') }
    if ($runAll -or $choice -eq '6') { Collect-SecurityData    ; $completed.Add('Security & Compliance') }
    if ($runAll -or $choice -eq '7') { Collect-PowerPlatformData; $completed.Add('Power Platform') }
    if ($runAll -or $choice -eq '8') { Collect-InformationBarriersData; $completed.Add('Information Barriers') }

    Write-SummaryReport -CompletedWorkloads $completed
    Write-Log "MCA Data Collection Scanner completed successfully." 'SUCCESS'
}
catch {
    Write-Log "Scanner failed: $($_.Exception.Message)" 'ERROR'
    Write-Log $_.ScriptStackTrace 'ERROR'
    throw
}

#endregion Main
