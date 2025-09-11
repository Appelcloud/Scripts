# Entra Authentication Methods Migration Audit

This repository contains a PowerShell script that audits your Microsoft Entra ID tenant for the migration from legacy MFA/SSPR to the unified Authentication Methods policy (MC678069). It produces a tenant‑named Excel workbook and an HTML report with clear guidance and counts.

## What It Does

- Reads user accounts and their registered authentication methods
- Assesses MFA coverage and “strong vs. weak” methods
- Highlights users who need action, including users who must “Register email or phone for SSPR”
- Optionally reads your current Authentication Methods policy status
- Optionally discovers resource accounts (rooms/workspaces) and excludes them
- Exports:
  - Excel workbook with multiple sheets (Detailed, Action Required, Summary)
  - HTML dashboard report
 
## DEMO

HTML report

<img width="2551" height="1220" alt="CleanShot 2025-09-11 at 19 32 10" src="https://github.com/user-attachments/assets/e914e1cf-3772-43e2-b922-6d02ee648d3e" />


EXCEL FILE

<img width="2556" height="1151" alt="CleanShot 2025-09-11 at 19 31 28" src="https://github.com/user-attachments/assets/292a07c9-de3a-4681-bf4b-70b1f4bf69a6" />


## Requirements

- PowerShell 7+ (recommended) or Windows PowerShell 5.1
- Microsoft Entra ID P1 or P2 license in the tenant
- Modules (installed automatically unless `-SkipModuleCheck` is used):
  - Microsoft.Graph PowerShell SDK
  - ImportExcel (only if Excel export is enabled)
- Directory role: Global Administrator or Authentication Policy Administrator

## Least‑Privilege Graph Permissions

The script connects to Microsoft Graph with minimal delegated scopes by default:

- `User.Read.All`, `UserAuthenticationMethod.Read.All`, `Reports.Read.All`, `Organization.Read.All`

Optional scopes are only requested when you opt into related features:

- `Policy.Read.All` when using `-IncludePolicyStatus`
- `Place.Read.All` when excluding resources (default behavior; used to discover rooms/workspaces). Pass `-IncludeResources` to avoid requesting this scope.

Tip: If your consent screen shows a very large list, someone may have previously consented to extra scopes for the shared Graph CLI app. You can leave “Consent on behalf of your organization” unchecked for ad‑hoc runs.

## Installation

Clone or download this repository, then open a PowerShell terminal in the repo root.

## Quick Start

Run the audit with defaults (Excel + HTML, exclude resources):

```
Entra-AuthMethods-MigrationAudit.ps1
```

Outputs are saved to the current folder as:

- `AuthMethods_MigrationReport_<TenantName>.xlsx`
- `AuthMethods_MigrationReport_<TenantName>.html`

## Parameters

- `-OutputPath <path>`: Where to save reports. Default: current directory
- `-SkipModuleCheck`: Skip auto‑install/import of required modules
- `-ExportHTML`: Generate HTML report (default: on)
- `-ExportExcel`: Generate Excel workbook (default: on)
- `-ExportCSV`: Also produce CSV files (default: off)
- `-IncludeResources`: Include resource accounts (shared mailboxes/rooms). If omitted, the script discovers rooms/sharedmailboxes and excludes them, and also excludes unlicensed mailboxes.
- `-IncludePolicyStatus`: Include current Authentication Methods policy state (adds `Policy.Read.All`)

Examples:

- Default minimal scopes (exclude resources, no policy status):
  - ``Entra-AuthMethods-MigrationAudit.ps1``
- Include resources (no Places permission requested):
  - ``Entra-AuthMethods-MigrationAudit.ps1 -IncludeResources``
- Include policy status section:
  - ``Entra-AuthMethods-MigrationAudit.ps1 -IncludePolicyStatus``
- Change output directory and enable CSV too:
  - ``Entra-AuthMethods-MigrationAudit.ps1 -OutputPath "C:\Reports" -ExportCSV``

## Output Details

- Excel workbook `AuthMethods_MigrationReport_<TenantName>.xlsx` with sheets:
  - **Detailed Report**: All users and detected methods
  - **Action Required**: Users where `NeedsAction` is true OR the only required action is “Register email or phone for SSPR”
  - **Summary**: One row with aggregate statistics
- HTML dashboard `AuthMethods_MigrationReport_<TenantName>.html` with:
  - Summary KPIs, readiness score, and risk categories
  - Authentication Methods distribution
  - Prominent note pointing to Excel for the full user lists
  - “Users Requiring Immediate Action” table (top 50; see Excel for full set)

## Resource Filtering Behavior

- By default, resource accounts are excluded:
  - Uses Graph Places (rooms/workspaces) to identify resource email addresses
  - Excludes unlicensed accounts to catch most shared mailboxes
- Pass `-IncludeResources` to include these accounts and avoid requesting Places permissions

## License Check (P1/P2)

The script verifies Entra ID P1/P2 via `subscribedSkus`. If not found, it stops with an error. If you prefer a warning‑only mode, you can adapt the function `Test-EntraP1Requirement` accordingly.

## Troubleshooting

- Excel auto‑fit warnings on macOS/Linux: harmless; ImportExcel cannot auto‑size without certain OS components.
- No users in HTML “Action Required”, but you see them in Excel: The HTML shows the first 50; the full list is in the Excel sheet.

## Security Notes

- All Graph calls are read‑only GET operations
- No write or configuration changes are performed

## Feedback / Contributions

Issues and PRs are welcome for improvements, bug fixes, or additional report sections.

