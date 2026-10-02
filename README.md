# MCA Data Collection Scanner

An interactive PowerShell scanner for collecting Microsoft 365 configuration, usage, and governance data for Microsoft Cloud Adoption assessments.

The scanner exports timestamped CSV files and a run log. It favors aggregate summaries and applies privacy filtering to exported rows. It does not remediate tenant settings or produce a compliance certification.

**Script documented here:** `MCA-assessment-scanner.ps1`  
**Author:** Mike Lee  
**Script header version:** 4.0

> This README describes the current local script, including its module preflight, authentication changes, and privacy filtering. If your repository uses a different script filename, substitute that filename in the commands below. Keep the published script and README together.

## Contents

- [What it collects](#what-it-collects)
- [Requirements and supported hosts](#requirements-and-supported-hosts)
- [Required PowerShell modules](#required-powershell-modules)
- [Permissions and authentication](#permissions-and-authentication)
- [Quick start](#quick-start)
- [Configuration](#configuration)
- [What happens during a scan](#what-happens-during-a-scan)
- [Output and interpretation](#output-and-interpretation)
- [Mailbox usage collection](#mailbox-usage-collection)
- [Privacy and sensitive-data handling](#privacy-and-sensitive-data-handling)
- [Troubleshooting](#troubleshooting)
- [Known limitations](#known-limitations)
- [Publishing and reporting issues](#publishing-and-reporting-issues)

## What it collects

Choose one workload or all workloads from the interactive menu.

| Choice | Workload | Collection areas |
| --- | --- | --- |
| `1` | SharePoint Online | Site/storage summaries, geography storage quotas, tenant settings, sharing settings, versioning settings, and site lifecycle settings. |
| `2` | Exchange Online | Mailbox usage aggregates; distribution-group, mail-enabled security-group, archive, hold, and public-folder summaries; organization, transport, retention, connector, and protection settings; DNS checks. |
| `3` | Microsoft Teams | Team summaries, tenant configuration, meeting/messaging/calling/app policies, guest and federation settings, voice configuration, auto attendants, call queues, and aggregate usage reports. |
| `4` | OneDrive for Business | Personal-site storage summary and OneDrive tenant settings. |
| `5` | Entra ID | User and group summaries, license SKUs, Conditional Access settings, authentication-method registration counts, and Microsoft 365 Apps usage counts. |
| `6` | Security & Compliance | Secure Score, control profiles, sensitivity labels, named locations, Purview DLP/audit configuration and selected events, privileged-access governance, application consents, Intune compliance, group governance, Message Center posts, and Graph connectors. |
| `7` | Power Platform | Power BI workspaces and capacities, plus Power Platform environments. |
| `8` | Information Barriers | Exchange organization/policy settings, segments, Information Barriers policies, and SharePoint Information Barriers settings. |
| `A` | All workloads | Runs workloads `1` through `8` sequentially. |
| `Q` | Quit | Exits without starting workload collection. |

Availability varies by permissions, licensing, service configuration, installed module version, and API availability. Selecting a workload does not guarantee that every collector will return data.

The Security & Compliance workload also creates files prefixed `EntraID_`, `Intune_`, `Purview_`, `M365_`, and `MicrosoftSearch_`. These additional governance collectors do not all run when choosing Entra ID alone.

## Requirements and supported hosts

The complete scanner is intended for **Windows**, including Windows PowerShell compatibility sessions used by the SharePoint module.

| Host | Behavior |
| --- | --- |
| PowerShell 7 on Windows | Recommended for all-workload runs. Start in a fresh `pwsh -NoProfile` session. |
| Windows PowerShell 5.1 console | Native collectors are supported, subject to each module's requirements. A compatibility notice identifies Graph-dependent collection areas. |
| Windows PowerShell 5.1 ISE | Remains in ISE; it is not automatically moved to PowerShell 7. Native collectors can run, but authentication-library conflicts can prevent Graph-backed collection. |
| PowerShell 5.0 and earlier | Not supported; the scanner requires at least 5.1. |
| PowerShell 7 integrated hosts or sessions with authentication libraries already loaded | The scanner attempts a clean `pwsh -NoProfile` child process to avoid inherited assembly conflicts. |
| Linux/macOS | Full scanner operation is not supported by this documented Windows-oriented workflow. |

Additional requirements:

- Network access to Microsoft 365 sign-in, Microsoft Graph, the selected workload services, and PowerShell Gallery if installing modules.
- An account authorized to read the selected tenant configuration and reports.
- A writable output folder.
- For Microsoft Graph on Windows PowerShell 5.1, .NET Framework 4.7.2 or later and the SDK's other prerequisites.
- Windows PowerShell availability for SharePoint compatibility imports from PowerShell 7.

Local administrator elevation is not normally needed for `CurrentUser` module installation. **Running PowerShell as a local administrator does not grant Microsoft 365 administrator roles.**

### What can be lost if Graph fails in ISE or PowerShell 5.1?

These are conditional limitations, not automatic exclusions.

| Selected workload | Unavailable if Graph authentication/access preflight fails | Native collection that can still be attempted |
| --- | --- | --- |
| Exchange Online | Graph mailbox usage summary and aggregate Graph reports. | Exchange configuration and native summaries. |
| Teams | Graph activity, device, meeting, and team usage reports. | Teams policies, configuration, and voice collectors. |
| Entra ID | The Entra ID workload is Graph-backed. | No native Entra ID fallback in this scanner. |
| Security & Compliance | Graph Secure Score, named locations, privileged access, app-consent, Intune, group governance, Message Center, connectors, and Graph label fallback. | Purview/Compliance PowerShell collectors, if sign-in and RBAC succeed. |

SharePoint, OneDrive, Information Barriers, and Power Platform have their own native-module authentication dependencies and may fail independently.

## Required PowerShell modules

The scanner checks only the modules relevant to the selected workloads.

| Module | Used for |
| --- | --- |
| `Microsoft.Online.SharePoint.PowerShell` | SharePoint, OneDrive, and SharePoint Information Barriers settings. |
| `ExchangeOnlineManagement` | Exchange, Security & Compliance PowerShell, and Information Barriers. |
| `MicrosoftTeams` | Teams policies, configuration, team summaries, and voice settings. |
| `Microsoft.Graph.Authentication` | Graph requests, report downloads, Entra ID, and Graph-backed security/governance collectors. |
| `Microsoft.Graph.Reports` | Optional enhancement for the D30 mailbox usage detail aggregation. |
| `MicrosoftPowerBIMgmt.Profile` | Power BI authentication. |
| `MicrosoftPowerBIMgmt.Workspaces` | Power BI workspace collection. |
| `MicrosoftPowerBIMgmt.Capacities` | Power BI capacity collection. |
| `Microsoft.PowerApps.Administration.PowerShell` | Power Platform environments. |

The Power BI modules are installed through the `MicrosoftPowerBIMgmt` package.

### Automatic module preflight

If modules are missing, the scanner displays their names, collection impact, and installation commands, then offers:

- `I`: Install missing packages with `Install-Module -Scope CurrentUser`.
- `C`: Continue with available modules; affected collectors may fail or use an available fallback.
- `Q`: Quit before workload collection.

If `Install-Module` is unavailable, only continue/quit is offered. Installation failures are displayed, and remaining missing modules are checked again.

Modules are **not** installed without an explicit choice. PowerShell Gallery, repository trust, or package-provider prompts can also appear.

### Manual installation

Install only the packages needed for your selected workloads, in the PowerShell host you intend to use:

```powershell
Install-Module Microsoft.Online.SharePoint.PowerShell -Scope CurrentUser
Install-Module ExchangeOnlineManagement -Scope CurrentUser
Install-Module MicrosoftTeams -Scope CurrentUser
Install-Module Microsoft.Graph.Authentication -Scope CurrentUser
Install-Module Microsoft.Graph.Reports -Scope CurrentUser
Install-Module MicrosoftPowerBIMgmt -Scope CurrentUser
Install-Module Microsoft.PowerApps.Administration.PowerShell -Scope CurrentUser
```

Windows PowerShell and PowerShell 7 can use different module installation folders. Confirm that a package is discoverable in the host where the scanner will run:

```powershell
Get-Module -ListAvailable -Name Microsoft.Graph.Authentication, MicrosoftTeams,
    ExchangeOnlineManagement, Microsoft.Online.SharePoint.PowerShell |
    Select-Object Name, Version, ModuleBase
```

The preflight detects missing modules; it is not a full version, dependency, or authentication health check.

## Permissions and authentication

### Delegated sign-in

The implemented collectors use interactive **delegated user authentication**:

| Service | Connection command |
| --- | --- |
| Microsoft Graph | `Connect-MgGraph` |
| SharePoint / OneDrive | `Connect-SPOService` |
| Exchange Online | `Connect-ExchangeOnline` |
| Security & Compliance | `Connect-IPPSSession` |
| Teams | `Connect-MicrosoftTeams` |
| Power BI | `Connect-PowerBIServiceAccount` |
| Power Platform | `Add-PowerAppsAccount` |

This workflow does not require you to configure a client secret, certificate, or custom application registration. It does require appropriate delegated consent and workload roles. Multiple services can prompt separately; the Power Platform environment child process can require another sign-in.

Where the installed cmdlets support the parameters, the scanner uses `DisableWAM` for Exchange, Compliance, and Teams, and `UseSystemBrowser` for SharePoint. Older module versions may not expose these options.

The script header contains some legacy app-only/service-principal guidance, especially for Power Platform. The current connection implementations above are delegated; do not treat those legacy comments as an app-only setup procedure.

### Workload roles

These are starting points, not an exhaustive RBAC entitlement list:

| Workload | Typical access requirement |
| --- | --- |
| SharePoint / OneDrive | SharePoint Administrator or equivalent SharePoint administration access. |
| Exchange | Exchange Administrator, Exchange Recipient Administrator, or equivalent Exchange RBAC for the requested cmdlets. |
| Teams | Teams Administrator or suitable read access; Global Reader coverage depends on the specific cmdlet. |
| Entra / Graph reports | Consented delegated scopes and appropriate directory/report-reader roles for each endpoint. |
| Security & Compliance | Appropriate Purview/Compliance role groups, audit permissions where applicable, and Graph endpoint-specific access. |
| Power BI | Tenant-wide workspace/capacity read access, generally through an appropriate Fabric/Power BI administrator role and tenant settings. |
| Power Platform | Power Platform Administrator or equivalent environment administration access. |
| Information Barriers | Applicable Exchange, Purview Information Barriers, and SharePoint permissions. |

Use the least privilege permitted by your organization's assessment process. Activate eligible PIM roles before starting, where applicable. A successful sign-in does not prove authorization for every command.

### Microsoft Graph scopes

Whenever Graph authentication is needed, the current script requests this fixed set of delegated scopes, rather than a smaller set tailored to the workload:

```text
Directory.Read.All
Reports.Read.All
Policy.Read.All
SecurityEvents.Read.All
InformationProtectionPolicy.Read
AuditLog.Read.All
UserAuthenticationMethod.Read.All
RoleManagement.Read.Directory
PrivilegedAccess.Read.AzureAD
AccessReview.Read.All
Application.Read.All
DelegatedPermissionGrant.Read.All
DeviceManagementManagedDevices.Read.All
DeviceManagementConfiguration.Read.All
ServiceMessage.Read.All
ExternalConnection.Read.All
Organization.Read.All
```

Review this list with your tenant administrator before granting consent. The Graph preflight signs in and reads the organization endpoint; it does **not** validate every endpoint, role assignment, license, or native workload permission.

## Quick start

1. Download the script from the repository and inspect it before execution.
2. Edit the Configuration section with your tenant ID/domain and SharePoint root URL.
3. Start a fresh PowerShell 7 console without preloading Teams, Exchange, Graph, or other authentication modules.
4. Run the saved script, select the workload, and respond to the preflight and sign-in prompts.
5. Review the timestamped CSV files and run log before drawing conclusions or sharing data.

Recommended command:

```powershell
pwsh -NoProfile -File "C:\Assessment\365-Assessment-Scanner.ps1"
```

Windows PowerShell 5.1:

```powershell
powershell.exe -NoProfile -File "C:\Assessment\365-Assessment-Scanner.ps1"
```

For ISE, open the saved `.ps1` file and run the complete script with **F5**. Running isolated selections can omit initialization or preflight logic.

If Windows marks the reviewed download as blocked, unblock that specific file:

```powershell
Unblock-File -LiteralPath "C:\Assessment\365-Assessment-Scanner.ps1"
```

Follow your organization's execution-policy requirements. Do not disable organizational security controls to run an assessment.

## Configuration

Configuration is edited inside the script; it does not currently expose a command-line parameter interface.

Replace tenant-specific values before running or publishing a reusable copy:

```powershell
$tenantId = 'contoso.onmicrosoft.com'
$tenantUrl = 'https://contoso.sharepoint.com'
$spoGeoAdminUrls = @{}
$OutputFolder = "$env:USERPROFILE\Documents\MCA_Assessment"
$debug = $false
```

The script currently validates that both `$tenantId` and `$tenantUrl` are populated, even for workloads that do not use SharePoint.

| Setting | Purpose |
| --- | --- |
| `$tenantId` | Target tenant GUID or verified domain used by supported connection helpers. |
| `$tenantUrl` | SharePoint tenant root URL; an admin URL is derived from it. |
| `$spoGeoAdminUrls` | Explicit geography-to-admin-endpoint mapping for SharePoint multi-geo scans. |
| `$OutputFolder` | Folder for CSV exports and the run log. |
| `$debug` | Controls display of DEBUG messages; DEBUG entries are still written to the run log. |
| `$MaxRetries`, `$InitialBackoffSec` | Settings for the explicit retry helper; not a uniform retry guarantee for every collector. |
| `$RequestTimeoutSec` | Declared timeout setting; do not assume every service command uses it. |
| `$SPOExportPollIntervalSec`, `$SPOExportMaxWaitSec` | Declared report-poll settings; their presence does not mean an unlicensed OneDrive report is currently collected. |

### SharePoint multi-geo

For a multi-geo tenant, supply actual admin endpoints for every expected geography, including the geography represented by `$tenantUrl`:

```powershell
$tenantUrl = 'https://contoso.sharepoint.com'
$spoGeoAdminUrls = @{
    NAM = 'https://contoso-admin.sharepoint.com'
    GBR = 'https://contosogbr-admin.sharepoint.com'
}
```

These values are illustrative. Use the verified endpoints and geography codes for your tenant.

The scanner uses geography quota data to identify expected locations; it does not infer admin hostnames from geography codes. Unresolved or failed geographies are reported, and consolidated site summaries can reflect incomplete coverage.

## What happens during a scan

1. Host handling runs; certain PowerShell 7 sessions are relaunched with `-NoProfile`. ISE stays in place.
2. The script prepares the output folder and presents the workload menu.
3. Windows PowerShell 5.1 shows the relevant compatibility notice.
4. Module preflight offers installation, partial continuation, or quit if packages are missing.
5. The operator confirms that the intended account has the required tenant access.
6. For workloads `2`, `3`, `5`, or `6`, Graph sign-in and basic tenant-read preflight run **before native modules are loaded**.
7. On Graph preflight failure, the operator chooses to quit or continue with Graph-backed requests disabled for the run.
8. Selected collectors execute, export available results, and log unavailable collectors.
9. A summary lists workloads traversed and CSV files in the output folder.

Graph authentication failures are latched to prevent repeated sign-in attempts from subsequent collectors. Individual collectors can still log unavailable-data warnings.

The scanner retains selected Exchange/Compliance/SharePoint sessions during all-workload runs and attempts cleanup of tracked sessions at the end. Do not assume that all SDK token caches or all service sessions are cleared.

## Output and interpretation

Default output folder:

```text
%USERPROFILE%\Documents\MCA_Assessment
```

Filenames use a run timestamp in the form `yyyyMMdd_HHmmss`.

| Example | Content |
| --- | --- |
| `MCA_Scan_<timestamp>.log` | Timestamped INFO, SUCCESS, WARN, ERROR, and DEBUG entries. |
| `SPO_SiteSummary_<timestamp>.csv` | SharePoint site/storage summary. |
| `SPO_GetSPOTenant_Full_<timestamp>.csv` | Tenant properties and their exported values, with targeted privacy handling. |
| `EXO_MailboxUsageSummary_<timestamp>.csv` | Aggregated D30 mailbox usage detail. |
| `EXO_DNSSecurityChecks_<timestamp>.csv` | DNS security checks for discovered non-onmicrosoft.com accepted domains. |
| `Teams_Policy_Meeting_<timestamp>.csv` | Teams meeting policy settings. |
| `Teams_Voice_AutoAttendants_<timestamp>.csv` | Sanitized auto-attendant configuration. |
| `ODB_StorageSummary_<timestamp>.csv` | OneDrive storage aggregates. |
| `EntraID_UsersSummary_<timestamp>.csv` | User-population summary subject to export filtering. |
| `Security_SecureScore_<timestamp>.csv` | Available Secure Score data. |
| `Purview_DLPPolicies_<timestamp>.csv` | DLP policy configuration. |
| `Intune_DeviceComplianceSummary_<timestamp>.csv` | Device compliance summary. |
| `PowerPlatform_Environments_<timestamp>.csv` | Environment configuration. |
| `IB_Policies_<timestamp>.csv` | Information Barriers policy data subject to identifier filtering. |

Exports use UTF-8 CSV. Some nested configuration values are serialized as JSON inside cells.

### Important interpretation rules

- A missing file can mean no returned data, a failed connection, missing permissions, an unsupported command, or an unavailable feature.
- A blank column can be intentional privacy filtering, not a missing source value.
- Some collectors write explicit status rows such as `Unavailable` or `NoneReturned`; read the status/note and log together.
- A workload marked complete means its function finished, not that all subcollectors succeeded.
- The final summary lists all CSV files in the output folder, including files from earlier runs.
- Use the filename timestamp and matching log to identify a specific run, or configure a separate folder for each assessment.
- Do not use a fixed expected file count as proof of completeness.

## Mailbox usage collection

The scanner uses:

```powershell
Get-MgReportMailboxUsageDetail -Period 'D30' -OutFile <temporary-file>
```

It aggregates the downloaded report into mailbox-count, storage, item-count, deleted-item-count, average-size, and report-refresh metadata fields. The common export sanitizer is then applied, so some summary columns can be blank.

This replaces slow per-mailbox `Get-EXOMailboxStatistics` calls. The report is a service-generated reporting snapshot, not live mailbox statistics:

- Values can lag the current tenant state.
- The report population can differ from a live Exchange mailbox inventory.
- It does not provide live per-mailbox troubleshooting details or recreate every field from `Get-EXOMailboxStatistics`.
- Use the report refresh date when interpreting results.

If the detail report/module fails, the collector attempts Graph aggregate mailbox-count and storage report endpoints. Those fallbacks still require working Graph authentication and report access.

In the current implementation, mailbox detail and its fallback are inside the successfully connected native Exchange collector block. A failed Exchange connection can therefore prevent these mailbox reports from being attempted.

The raw mailbox detail CSV is temporarily written under the user's temporary directory. Cleanup is attempted in `finally`; a deletion failure is logged for manual cleanup.

## Privacy and sensitive-data handling

**Privacy filtering reduces exposure; it is not a guarantee that every output is anonymous or free of secrets.**

### Implemented protections

- Common export processing blanks columns whose names match broad identity-related patterns, including user, principal, name, mail, address, owner, actor, tenant, domain, URL, GUID, and ID patterns.
- Remaining text is filtered for email addresses, hyphenated GUIDs, IPv4-like strings, selected SAS URL patterns, and `SharedAccessSignature=` values.
- SharePoint full tenant-property values receive additional handling, including HTTP/HTTPS URL removal.
- Several identity-heavy collectors produce summaries instead of exporting full per-user records.
- Graph report downloads are processed through the export sanitizer.
- Persisted log messages receive the common text filter.

### Important boundaries

- The field-name filter is deliberately broad and can remove useful, non-personal values, including some summary metrics. Do not assume all calculated fields survive export.
- The scanner can read identifiable source data into memory before aggregation or sanitization.
- Raw report CSVs and the isolated Power Platform result file can exist temporarily on disk.
- Cleanup attempts are not secure erasure and cannot guarantee removal after a crash, forced termination, or filesystem failure.
- Console messages can display tenant information, local paths, and error details that are filtered differently in the saved log.
- Nested JSON, arbitrary free text, unusual credential formats, webhook tokens, storage account keys, and other unrecognized secrets are not comprehensively guaranteed to be removed by the pattern filters.
- Configuration and governance information can remain sensitive even without direct personal identifiers.
- Existing exports are not retroactively sanitized when the script changes.
- Outputs are not automatically encrypted, sensitivity-labeled, or approved for external distribution.

Before sharing, inspect all CSVs and logs under your organization's data-handling rules. Do not post raw assessment output, sign-in diagnostics, tokens, or customer-specific configuration to GitHub.

The assessment collectors are read-oriented. Module installation changes the local environment; authentication and consent can create local token-cache or tenant consent state. The script does not intentionally remediate Microsoft 365 policies.

## Troubleshooting

### Graph `WithLogging` / `Method not found` errors

Example:

```text
InteractiveBrowserCredential authentication failed:
Method not found: ... BaseAbstractApplicationBuilder`1.WithLogging(...)
```

This indicates an authentication-library binary mismatch, not proof of missing tenant permissions.

The scanner initializes Graph before loading native workload modules. On PowerShell 7, it also attempts a clean relaunch if relevant authentication assemblies are already loaded.

Start a new console and run the scanner directly:

```powershell
pwsh -NoProfile -File "C:\Assessment\365-Assessment-Scanner.ps1"
```

Do not import Teams, Exchange, Power BI, Azure, or other authentication modules first. `Remove-Module` does not reliably unload .NET assemblies already loaded into the process.

Changing from browser authentication to device code does not repair a binary mismatch. The scanner does not retry device code for recognized `WithLogging` or `RefreshCacheAsync` conflicts.

If a clean run still fails, inspect installed module versions and paths. Repair/update affected packages deliberately, then restart the host; do not delete every installed module as a blanket workaround.

### `RefreshCacheAsync` / method implementation errors

These can indicate mixed or incompatible Graph authentication components. Check the selected Graph module versions and start a new host after any update. Module presence alone is not evidence that its dependencies can load correctly.

### Missing `msalruntime` / WAM errors

Where supported, native connection helpers disable WAM or use the system browser. Graph can retry once with device-code authentication for a recognized missing-WAM-runtime error.

Device-code authentication can itself be blocked by tenant policy. Do not weaken Conditional Access to make the scan run.

### Access denied, 403, or insufficient privileges

Confirm the correct tenant/account, active workload roles, delegated consent, service licensing, and endpoint-specific access. The operator access confirmation is not an automated role audit.

### SharePoint: no valid OAuth session

Check the configured tenant root/admin endpoint, SharePoint Administrator access, installed SharePoint module, and successful sign-in. In PowerShell 7, SharePoint normally uses a Windows PowerShell compatibility session.

### Missing mailbox usage output

Review both Exchange connection and Graph report errors. Verify `Microsoft.Graph.Reports`, report access, and the report refresh date. Native mailbox configuration can still be collected when Graph reports are unavailable.

### Teams reports return 400 or no data

Inspect the exact endpoint and service error. Licensing, telemetry availability, tenant/report support, request validity, and permissions can affect the outcome. A 400 response alone does not establish one specific cause.

### Power Platform asks for another sign-in

The environment collector tries a separate clean PowerShell process to avoid assembly conflicts. That process has its own authentication context. If isolated collection fails, the scanner attempts an in-process native fallback, which can also prompt.

### A workload completed but expected files are absent

Read WARN/ERROR entries and any status rows. Empty results are not always exported, and collector-level failures do not necessarily stop the whole scan.

## Known limitations

- Interactive operation; this is not an unattended scheduled-task or app-only automation workflow.
- Native module behavior and compatibility vary by installed version.
- The Graph scope list is fixed and broader than some individual workloads need.
- Basic Graph preflight does not validate every collector's authorization.
- Sanitization intentionally sacrifices identity detail and can also remove non-sensitive fields.
- Temporary files and console output require appropriate local handling.
- Reporting data can be delayed, and some governance/audit collectors use bounded windows or result limits rather than exhaustive historical exports.
- Some governance endpoints use beta or alternative endpoint fallbacks; service contracts can change.
- A successful run message does not certify complete collection, tenant security, or regulatory compliance.
- Commercial-cloud endpoints and SharePoint hostname assumptions are used; sovereign-cloud support is not established by this README.
- No built-in upload to GitHub or assessment portal; review and sharing are separate operator responsibilities.

## Publishing and reporting issues

### Before publishing to GitHub

1. Replace tenant-specific configuration in the published script with clearly documented example values.
2. Publish this README alongside the matching script; update command examples if the repository filename differs.
3. Exclude assessment CSVs, logs, temporary files, and local customer configuration.
4. Review the actual repository diff for identifiers and credentials before committing.
5. If distributing under a license, include the license chosen by the repository owner; this README does not assign one.

Suggested `.gitignore` additions when using a repository-local assessment folder:

```gitignore
MCA_Assessment/
MCA_Scan_*.log
mca_report_*.csv
mca_mailbox_usage_*.csv
mca_pp_env_*.json
```

For an issue report, include the script revision, PowerShell version/edition, selected workload, relevant module versions, and a sanitized error excerpt. Do not include tokens, raw report files, customer tenant URLs, or identifiable audit events.

Useful local diagnostic commands:

```powershell
$PSVersionTable

Get-Module -ListAvailable -Name Microsoft.Graph.Authentication,
    Microsoft.Graph.Reports, MicrosoftTeams, ExchangeOnlineManagement,
    Microsoft.Online.SharePoint.PowerShell |
    Select-Object Name, Version, ModuleBase
```

Review and redact diagnostic output before posting it publicly.

### References

- [Install Microsoft Graph PowerShell](https://learn.microsoft.com/en-us/powershell/microsoftgraph/installation)
- [Microsoft Graph authentication](https://learn.microsoft.com/en-us/powershell/microsoftgraph/authentication-commands)
- [Exchange Online PowerShell](https://learn.microsoft.com/en-us/powershell/exchange/exchange-online-powershell)
- [SharePoint Online Management Shell](https://learn.microsoft.com/en-us/powershell/sharepoint/sharepoint-online/connect-sharepoint-online)
- [Microsoft Teams PowerShell](https://learn.microsoft.com/en-us/microsoftteams/teams-powershell-overview)
