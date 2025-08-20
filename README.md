# Get-M365CopilotReadiness.ps1

**Checks a tenant’s technical readiness for Copilot f### SharePoint Online / OneDrive

* Derives SPO Admin URL from the tenant's `*.onmicrosoft.com` domain when not provided, or uses `-SPOAdminUrl` if passed.
* Connects to SPO Admin and retrieves:

  * `Get-SPOTenant` **tenant properties** (e.g., OneDrive storage quota, sharing/CA policy indicators, etc.)
  * **Enhanced sharing settings with descriptions**: Sharing capabilities, link types, expiration policies with clear explanations
  * Total SPO sites & **OneDrive** site count (filtered by URL/template)
  * **Content type synchronization settings**
  * **Teams integration settings**
  * **Search and language configurations**
  * **Site-level restrictions and policies**
* Emits:

  * `SharePointOnline.Connected`, `AdminUrl`, `TenantProperties`, `TotalSites`, `OneDriveSites`, `SharingSettings`, `RacPolicySites`, `RestrictedSites`, `SearchSettings`, `LanguageSettings`, `ContentTypeSync`, `TeamsSettings`, `Notes`

> Readiness flag derived: **`OneDrive_Provisioned`** (Pass if any OneDrive sites found). 365** by connecting (best effort) to Microsoft Graph, Exchange Online, SharePoint Online, and Microsoft Teams, then outputting a JSON report and a human-readable HTML summary.

* **Outputs:**

  * `copi> **Notes for contributors**
>
> * This script currently performs **readiness heuristics** with a "best effort" connection model and light signals. Pull requests that add deeper checks (e.g., detailed Teams/EXO settings, network egress validation, Purview/Defender signals, richer licensing mapping) are welcome—please keep output backward-compatible or gate behind a switch.
> * **Recent additions:** The script now includes comprehensive Entra ID external sharing and guest user policy collection to provide insights into tenant-level external collaboration settings that may impact Copilot usage and security posture.
> * **Enhanced reporting format (August 2025):** All configuration settings now include descriptive explanations alongside their values, making reports more administrator-friendly. The HTML output features three-column tables (Setting, Value, Description) and contextual guidance about Copilot impact.
> * **Improved error handling:** Enhanced graceful handling of policy endpoints that may not be available in all tenant configurations, with clear informational messaging instead of confusing warnings.
> * If you contribute app-only authentication support, add a separate section in this README with Azure AD app registration steps and exact Graph permissions (Application) required.readiness.json`
  * `copilot-readiness.html`

> References used in the script:
>
> * Requirements: [https://learn.microsoft.com/en-us/copilot/microsoft-365/microsoft-365-copilot-requirements](https://learn.microsoft.com/en-us/copilot/microsoft-365/microsoft-365-copilot-requirements)
> * Enablement:  [https://learn.microsoft.com/en-us/copilot/microsoft-365/microsoft-365-copilot-enablement-resources](https://learn.microsoft.com/en-us/copilot/microsoft-365/microsoft-365-copilot-enablement-resources)
> * Licensing:   [https://learn.microsoft.com/en-us/copilot/microsoft-365/microsoft-365-copilot-licensing](https://learn.microsoft.com/en-us/copilot/microsoft-365/microsoft-365-copilot-licensing)

---

## Table of Contents

* [What the script does](#what-the-script-does)
* [Per-service details](#per-service-details)
* [Prerequisites](#prerequisites)
* [Required permissions / roles](#required-permissions--roles)
* [Parameters](#parameters)
* [Installation & usage](#installation--usage)
* [Output files & schema](#output-files--schema)
* [Troubleshooting](#troubleshooting)
* [Author & date](#author--date)

---

## What the script does

1. **Module bootstrap (optional):** Installs and imports required PowerShell modules for the current user scope unless `-SkipModuleInstall` is specified.
2. **Service connections (best effort):**

   * Microsoft Graph (with read-only scopes)
   * Exchange Online
   * SharePoint Online (SPO Admin)
   * Microsoft Teams
3. **Signals collection aligned to Microsoft Learn guidance:**

   * **Licensing & eligibility** for Copilot (base license patterns + Copilot SKUs)
   * **Exchange Online** mailbox presence (basic signal for “hosted in EXO”)
   * **SharePoint/OneDrive** tenant properties & counts (incl. OneDrive site count)
   * **Teams** connectivity check
   * **Tenant/Org** basics from Graph (display name, domains, etc.)
4. **Report generation:** Writes a **structured JSON** payload and an **HTML** summary dashboard with readiness flags, captured errors, and basic guidance links.

---

## Per-service details

### Microsoft Graph

* Connects with read-only scopes to query **organization** details, **subscribed SKUs**, and **Entra ID policies**.
* **Scopes requested:**
  `Organization.Read.All`, `Directory.Read.All`, `User.Read.All`, `ExternalItem.Read.All`, `Sites.Read.All`, `ExternalConnection.Read.All`, `Policy.Read.All`

### Licensing signals

* Pulls **Subscribed SKUs** and identifies:

  * **Eligible base licenses** (e.g., Microsoft 365 E3/E5, Office 365 E3/E5; pattern-based detection)
  * Presence of **Copilot-related SKUs** (inspects service plans that match `COPILOT` / `CPI`)
* Emits:

  * `Licensing.CopilotSkuPresent` (boolean)
  * `Licensing.CopilotSkus` (subset of SKUs relevant to Copilot)
  * `Licensing.EligibleBaseLicenses` (subset of base SKUs)
  * `Licensing.AllRelevantSkus` (merged view of base + Copilot SKUs with counts)

### Exchange Online

* Attempts EXO connection and queries user mailboxes (lightweight check).
* Emits:

  * `ExchangeOnline.Connected` (boolean)
  * `ExchangeOnline.UserMailboxCount` (int, if accessible)
  * `ExchangeOnline.Notes` (errors or connection notes)

> Readiness flag derived: **`EXO_PrimaryMailboxHostedInEXO`** (heuristic: at least one user mailbox returned).

### SharePoint Online / OneDrive

* Derives SPO Admin URL from the tenant’s `*.onmicrosoft.com` domain when not provided, or uses `-SPOAdminUrl` if passed.
* Connects to SPO Admin and retrieves:

  * `Get-SPOTenant` **tenant properties** (e.g., OneDrive storage quota, sharing/CA policy indicators, etc.)
  * Total SPO sites & **OneDrive** site count (filtered by URL/template)
* Emits:

  * `SharePointOnline.Connected`, `AdminUrl`, `TenantProperties`, `TotalSites`, `OneDriveSites`, `SharingSettings`, `RacPolicySites`, `RestrictedSites`, `Notes`

> Readiness flag derived: **`OneDrive_Provisioned`** (Pass if any OneDrive sites found).

### Microsoft Teams

* Attempts a Teams connection to validate access (no heavy data collection).
* Emits:

  * `Teams.Connected` (boolean)

### Entra ID External Sharing & Guest Settings

* Retrieves tenant-level external collaboration and guest user policies that may impact Copilot usage and security.
* **Enhanced with descriptive explanations**: Each setting includes both the value and a clear description of what it means and its impact on Copilot.
* Collects:

  * **Authorization Policy:** Guest invitation settings, default user permissions, email verification settings
  * **External Identities Policy:** External identity management and data removal policies
  * **Guest User Settings:** Guest user role configurations and statistics
  * **Cross-Tenant Access Policy:** Cross-tenant collaboration settings and allowed cloud endpoints
  * **Organization Settings:** Privacy profile, notification email configurations
  * **Guest User Statistics:** Count and impact assessment of guest users in the tenant
* **Admin-friendly format**: All settings include contextual descriptions explaining their significance for Copilot readiness and security
* Emits:

  * `EntraId.Connected`, `AuthorizationPolicy`, `ExternalSharingSettings`, `GuestUserSettings`, `CrossTenantAccessPolicy`, `GuestUserStatistics`, `OrganizationSettings`, `Notes`

> Readiness flag derived: **`EntraId_ExternalSharingConfigured`** (Pass if authorization policies are accessible and configured).

---

## Enhanced Reporting Features

### Administrator-Friendly Output

The script now provides enhanced, administrator-friendly reporting with the following improvements:

* **Descriptive Format**: All configuration settings include both the actual value and a clear description of what it means
* **Copilot Context**: Explanations of how each setting impacts Copilot usage and security
* **Three-Column Tables**: HTML output shows Setting, Value, and Description for maximum clarity
* **Impact Assessment**: Clear indicators of security considerations and data governance implications

### Value-Description Structure

Settings are now reported in an enhanced format:

```json
"AllowInvitesFrom": {
  "Value": "adminsAndGuestInviters",
  "Description": "Who can invite external users (none, adminsAndGuestInviters, adminsGuestInvitersAndAllMembers, everyone)"
}
```

### Key Benefits

* **No Technical Guessing**: Every setting includes plain-language explanations
* **Decision Support**: Descriptions help administrators make informed policy decisions
* **Security Awareness**: Clear identification of settings that affect external access
* **Copilot Readiness**: Understand how configurations impact Copilot functionality

---

## Prerequisites

* **PowerShell:** Windows PowerShell 5.1 or PowerShell 7.x
* **Network access** to Microsoft 365 endpoints for Graph, Exchange Online, SharePoint Online, and Teams.
* **Interactive sign-in** capability (unless you adapt for app-only auth; this script is written for interactive admin usage).

### PowerShell modules (auto-installed unless `-SkipModuleInstall`)

| Module                                       | Minimum Version  |
| -------------------------------------------- | ---------------- |
| Microsoft.Graph.Authentication               | 2.8.0            |
| Microsoft.Graph.Identity.DirectoryManagement | 2.8.0            |
| Microsoft.Graph.Search                       | 2.8.0            |
| ExchangeOnlineManagement                     | 3.4.0            |
| MicrosoftTeams                               | 5.6.0            |
| Microsoft.Online.SharePoint.PowerShell       | 16.0.24908.12000 |

> The script imports these modules and will attempt installation for CurrentUser scope if they’re missing.

---

## Required permissions / roles

> The script connects **interactively** and uses **read-only** Graph scopes.

* **Graph scopes:**
  `Organization.Read.All`, `Directory.Read.All`, `User.Read.All`, `ExternalItem.Read.All`, `Sites.Read.All`, `ExternalConnection.Read.All`, `Policy.Read.All`
* **Exchange Online:** permissions sufficient to run **`Get-EXOMailbox`** (e.g., **View-Only Recipients** or higher; many tenants grant this via EXO/Exchange admin roles).
* **SharePoint Online:** **SharePoint Administrator** (or **Global Administrator**) is typically required for **`Get-SPOTenant`**.
* **Microsoft Teams:** Teams admin-level read permissions are safest; this script only verifies connection.

> If your account lacks a given role, the script continues where possible and records a connection failure and/or data errors in the final report.

---

## Parameters

```powershell
PARAMETERS
----------
-OutputPath <String>
    Directory for report outputs. Defaults to current directory (".").
    Example: -OutputPath "C:\Temp\M365Readiness"

-SkipModuleInstall [Switch]
    If specified, the script will NOT attempt to install missing modules.

-SPOAdminUrl <String>
    Optional explicit SharePoint Admin URL (e.g., https://contoso-admin.sharepoint.com).
    If omitted, the script tries to derive it from your tenant’s *.onmicrosoft.com domain.
```

---

## Installation & usage

### 1) Clone or download

```powershell
git clone https://github.com/<your-org>/<your-repo>.git
cd <your-repo>
```

### 2) Run the script (interactive)

```powershell
# Basic run, outputs to current directory
.\Get-M365CopilotReadiness.ps1

# Specify output folder
.\Get-M365CopilotReadiness.ps1 -OutputPath "C:\Temp\M365Readiness"

# Provide SPO Admin URL explicitly (if your tenant derivation is special)
.\Get-M365CopilotReadiness.ps1 -SPOAdminUrl "https://contoso-admin.sharepoint.com"

# Skip auto-install of modules (if you pre-installed everything)
.\Get-M365CopilotReadiness.ps1 -SkipModuleInstall
```

> You’ll be prompted to sign in for each service connection (Graph, EXO, SPO, Teams). If a connection fails, the script logs the error and continues.

---

## Output files & schema

### `copilot-readiness.json`

High-level structure:

```json
{
  "GeneratedAtUtc": "2025-08-19T14:12:34Z",
  "ScriptDurationSec": 12.34,
  "LearnReferences": [...],
  "Connections": { "Graph": true, "EXO": true, "Teams": true, "SPO": true },
  "Tenant": {
    "DisplayName": "Contoso Ltd",
    "Id": "...",
    "DefaultDomain": "contoso.onmicrosoft.com",
    "VerifiedDomains": [ "contoso.com", "contoso.onmicrosoft.com", ... ],
    "CountryLetterCode": "...",
    "TenantType": "..."
  },
  "Licensing": {
    "CopilotSkuPresent": true,
    "CopilotSkus": [ { "SkuPartNumber": "...", "ConsumedUnits": 123, ... } ],
    "EligibleBaseLicenses": [ { "SkuPartNumber": "MICROSOFT_365_E5", ... } ],
    "AllRelevantSkus": [ ... ]
  },
  "Services": {
    "ExchangeOnline": { "Connected": true, "UserMailboxCount": 123, "Notes": "" },
    "SharePointOnline": {
      "Connected": true,
      "AdminUrl": "https://contoso-admin.sharepoint.com",
      "TenantProperties": { "OneDriveStorageQuota": "...", "SharingCapability": "...", ... },
      "TotalSites": 456,
      "OneDriveSites": 123,
      "SharingSettings": {
        "TenantSharingLevel": {
          "Value": "ExternalUserAndGuestSharing",
          "Description": "Tenant-wide sharing level (Disabled, ExternalUserSharingOnly, ExistingExternalUserSharingOnly, ExternalUserAndGuestSharing)"
        },
        "DefaultSharingLinkType": {
          "Value": "Internal",
          "Description": "Default sharing link type for new sharing links (None, Direct, Internal, AnonymousAccess)"
        },
        "FileAnonymousLinkType": {
          "Value": "Edit",
          "Description": "Anonymous link permissions for files (None, View, Edit)"
        }
      },
      "RacPolicySites": [ ... ],
      "RestrictedSites": [ ... ],
      "Notes": ""
    },
    "Teams": { "Connected": true },
    "EntraId": {
      "Connected": true,
      "AuthorizationPolicy": {
        "AllowInvitesFrom": {
          "Value": "adminsAndGuestInviters",
          "Description": "Who can invite external users (none, adminsAndGuestInviters, adminsGuestInvitersAndAllMembers, everyone)"
        },
        "AllowEmailVerifiedUsersToJoinOrganization": {
          "Value": true,
          "Description": "Whether email-verified users can join the organization without invitation"
        },
        "DefaultUserRolePermissions": { ... }
      },
      "ExternalSharingSettings": {
        "AllowExternalIdentitiesToLeave": {
          "Value": true,
          "Description": "Whether external users can leave the organization on their own"
        },
        "AllowDeletedIdentitiesDataRemoval": {
          "Value": true,
          "Description": "Whether data is automatically removed when external identities are deleted"
        }
      },
      "GuestUserSettings": { ... },
      "CrossTenantAccessPolicy": {
        "IsServiceDefault": {
          "Value": true,
          "Description": "Whether this policy uses service defaults (true) or has custom configuration (false)"
        }
      },
      "GuestUserStatistics": {
        "TotalGuestUsers": {
          "Value": 42,
          "Description": "Number of guest users currently in the tenant"
        },
        "Impact": {
          "Value": "External users present",
          "Description": "Potential impact on Copilot data access and security considerations"
        }
      },
      "OrganizationSettings": { ... },
      "Notes": ""
    },
    "Graph": { "Connected": true, "Scopes": [ "Organization.Read.All", ... ] }
  }
}
```

### `copilot-readiness.html`

* A compact, readable dashboard summarizing:

  * Connection status per service
  * Licensing highlights (base/Copilot)
  * Basic Exchange / SPO / OneDrive signals
  * **Enhanced Entra ID external sharing and guest user configuration with descriptions**
  * **Improved SharePoint sharing settings with contextual explanations**
  * **Three-column tables showing Setting, Value, and Description for better admin understanding**
  * **Contextual guidance about how settings impact Copilot usage and security**
  * Links to Microsoft Learn references
  * Error/notes section (if any)

---

## Troubleshooting

* **Graph connection fails / consent prompts:**
  Ensure your account can grant or has admin consent for the read-only scopes listed above. Retry after consent or run as an admin with sufficient rights.

* **`Get-SPOTenant` access denied:**
  You likely need the **SharePoint Administrator** or **Global Administrator** role.

* **`Get-EXOMailbox` access denied / throttled:**
  Ensure you have Exchange permissions (View-Only Recipients or above). Large tenants may throttle; rerun or scope with your own adaptation if needed.

* **Entra ID policy access denied:**
  The script requires `Policy.Read.All` permissions to retrieve authorization and external identity policies. Ensure your account has sufficient permissions or admin consent has been granted for this scope.

* **External Identities Policy warnings:**
  If you see informational messages about External Identities Policy not being available, this is expected behavior in many tenant configurations. This policy is only available in tenants with specific licensing (like Azure AD Premium P2) or certain configurations.

* **Teams connection fails but others succeed:**
  This does not block report generation; it will be recorded as not connected. Verify the **MicrosoftTeams** module is current and that your account has Teams admin access.

* **Module import/installation errors:**
  Use `-SkipModuleInstall` if your environment restricts `Install-Module`, and pre-install the required modules with your standard process (e.g., internal repository).

---

## Author & date

* **Author:** John Cummings ([john@jcummings.net](mailto:john@jcummings.net))
* **Date:** August 20, 2025

---

> **Notes for contributors**
>
> * This script currently performs **readiness heuristics** with a “best effort” connection model and light signals. Pull requests that add deeper checks (e.g., detailed Teams/EXO settings, network egress validation, Purview/Defender signals, richer licensing mapping) are welcome—please keep output backward-compatible or gate behind a switch.
> * If you contribute app-only authentication support, add a separate section in this README with Azure AD app registration steps and exact Graph permissions (Application) required.
