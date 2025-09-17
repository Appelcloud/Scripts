# Entra Authentication Methods Migration Audit

This repository contains a PowerShell script that audits your Microsoft Entra ID tenant for the migration from legacy MFA/SSPR to the unified Authentication Methods policy (MC678069). It produces a tenant‑named Excel workbook and an HTML report with clear guidance, a combined legacy→modern mapping, and actionable recommendations.

## What It Does

- Enumerates user accounts and their registered authentication methods (Graph).
- Assesses MFA coverage and “strong vs. weak” methods.
- Highlights users who need action.
- Reads the Authentication Methods policy configuration (Graph beta)
- Compares legacy usage (from registration reports) with modern configuration to show alignment vs. gaps
- Exports:
  - Excel workbook with multiple sheets (Detailed, Action Required, Auth Methods Status, Legacy vs Modern)
  - HTML dashboard report with Migration Readiness at the top and a combined “Authentication Methods Status” table
 
## DEMO

HTML report

<img width="2535" height="1222" alt="CleanShot 2025-09-17 at 13 17 35" src="https://github.com/user-attachments/assets/121a248a-97c3-4ed0-9911-87a664c3708a" />

EXCEL FILE

<img width="2556" height="1151" alt="CleanShot 2025-09-11 at 19 31 28" src="https://github.com/user-attachments/assets/292a07c9-de3a-4681-bf4b-70b1f4bf69a6" />


## Requirements

- PowerShell 7+ (recommended) or Windows PowerShell 5.1
- Microsoft Entra ID P1 or P2 license in the tenant
- Modules (installed automatically unless `-SkipModuleCheck` is used):
  - Microsoft.Graph PowerShell SDK
  - ImportExcel (only if Excel export is enabled)
- Directory role: Global Administrator or Authentication Policy Administrator

## Graph Permissions Requested

The script connects to Microsoft Graph with delegated scopes:

- Always: `User.Read.All`, `UserAuthenticationMethod.Read.All`, `Reports.Read.All`, `AuditLog.Read.All`, `Organization.Read.All`, `Policy.Read.All`
- Optional: `Place.Read.All` when excluding resources (default); don't pass `-IncludeResources` to avoid requesting Places.

Notes:
- Modern policy read uses the beta endpoint with `$expand=authenticationMethodConfigurations`.
- Legacy usage is taken from the authentication methods registration reports.

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
- `-IncludePolicyStatus`: Show current policy migration state (preMigration/migrationInProgress/etc.)
- `-Quiet`: Reduce console spam. Default: off (verbose)
- `-ShowDetails`: Print detailed progress such as user retrieval counts and resource discovery (default: off)

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
  - **Auth Methods Status**: Combined comparison of legacy usage and modern policy, including the same alignment Status and Recommendation as the HTML
  - **Legacy vs Modern**: Underlying comparison rows (one per method), for further filtering/pivoting
- HTML dashboard `AuthMethods_MigrationReport_<TenantName>.html` with:
  - Migration Readiness Assessment at the top (overall status and score)
  - Authentication Methods Status (combined legacy vs modern) table
  - Summary cards, risk categories, and methods distribution
  - “Users Requiring Action” table (top 50), with a note that full lists are in the Excel export

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
- Policy read 400/badRequest: Ensure `Policy.Read.All` is consented; this script requests it by default. The policy read uses `$expand` with the beta endpoint.

## Security Notes

- All Graph calls are read‑only GET operations
- No write or configuration changes are performed

## Feedback / Contributions

Issues and PRs are welcome for improvements, bug fixes, or additional report sections.
