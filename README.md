# Microsoft 365 Tenant-to-Tenant (T2T) Migration Automation Framework

![PowerShell 7+](https://img.shields.io/badge/PowerShell-7%2B-5391FE?logo=powershell&logoColor=white)
![Microsoft Graph SDK](https://img.shields.io/badge/Microsoft_Graph_SDK-2.x-0078D4?logo=microsoftazure&logoColor=white)
![Exchange Online v3](https://img.shields.io/badge/Exchange_Online_Management-v3-0078D4?logo=microsoftexchange&logoColor=white)

A scripted, auditable framework for moving mailboxes and identities between two Microsoft 365 tenants using the **native cross-tenant mailbox migration** capability of Exchange Online, orchestrated with PowerShell 7, the Microsoft Graph SDK, and Exchange Online Management v3.

---

## Executive Summary

Tenant-to-tenant migrations are high-risk, low-tolerance engagements: a wrong-tenant write or a premature domain release can cause mail-flow outages. This framework decomposes the migration into four gated phases — **Pre-flight, Discovery, Staging/Sync, and Cutover** — each with explicit validation, transcript logging, and reversible, idempotent operations.

The framework was developed using an **AI-accelerated engineering workflow**, combining rapid, machine-assisted code generation with human architectural review, and was subjected to **rigorous lab testing against a disposable tenant pair before production deployment**. Every destructive or tenant-mutating operation supports `-WhatIf` / `-Confirm`, and every run produces a full transcript for post-migration audit.

## Repository Layout

```
.
├── README.md
├── .gitignore
├── scripts/
│   ├── Test-MigrationPreflight.ps1        # Reusable environment/connection gate
│   ├── Step1_SourceExport.ps1             # Phase 1: source tenant discovery
│   ├── Step2_TargetStagingAndBatch.ps1    # Phase 2: target staging + migration batch
│   └── Step3_SourceDomainCleanup.ps1      # Phase 3: source domain release preparation
└── docs/
    └── ARCHITECTURE.md                    # Cross-tenant workflow diagram + design notes
```

## Prerequisites & Module Setup

| Requirement | Detail |
|---|---|
| PowerShell | **7.0+** (`pwsh`) — scripts declare `#Requires -Version 7.0` |
| Microsoft.Graph | **>= 2.0.0** (Graph SDK v2) |
| ExchangeOnlineManagement | **>= 3.0.0** (REST-based EXO v3 module) |
| Source tenant rights | Global Reader or equivalent for discovery; User Admin + Exchange Admin for cutover |
| Target tenant rights | User Admin (Graph `User.ReadWrite.All`) + Exchange Admin |
| One-time cross-tenant setup | App registration, organization relationship, and migration endpoint (below) |

Install the modules:

```powershell
Install-Module Microsoft.Graph -MinimumVersion 2.0.0 -Scope CurrentUser -Force
Install-Module ExchangeOnlineManagement -MinimumVersion 3.0.0 -Scope CurrentUser -Force
```

### One-time cross-tenant plumbing (target tenant)

Native cross-tenant moves require an Exchange migration endpoint backed by an app-only OAuth trust. Complete these steps once, following the official Microsoft procedure — [Cross-tenant mailbox migration](https://learn.microsoft.com/microsoft-365/enterprise/cross-tenant-mailbox-migration):

1. Create an app registration in the **target** tenant with the Exchange `full_access_as_app` application permission, and have a **source** tenant admin grant consent.
2. Create a mail-enabled security group in the **source** tenant containing the mailboxes to migrate.
3. Create the organization relationship in the **target** tenant (`New-OrganizationRelationship` with mailbox-move publishing enabled).
4. Create the migration endpoint in the **target** tenant:

```powershell
$credential = Get-Credential   # target tenant Exchange admin
New-MigrationEndpoint -ExchangeRemoteMove `
    -Name 'T2T-SourceEndpoint' `
    -RemoteServer 'outlook.office365.com' `
    -ApplicationId '<target app registration application (client) id>' `
    -AppSettingsTenantId '<source tenant id>' `
    -Credentials $credential
```

## Execution Workflow

> Run each phase from the repository root in `pwsh`. Phases 0–1 connect to the **source** tenant; phase 2 connects to the **target** tenant; phase 3 returns to the **source** tenant.

### Phase 0 — Pre-flight

```powershell
Connect-MgGraph -Scopes 'User.Read.All', 'Group.Read.All', 'Domain.Read.All'
Connect-ExchangeOnline

./scripts/Test-MigrationPreflight.ps1
```

Validates the PowerShell engine, module versions, Graph context/scopes, and the Exchange Online session. Every step script re-runs this gate internally and refuses to proceed on failure.

### Phase 1 — Discovery (Source Export)

```powershell
./scripts/Step1_SourceExport.ps1
```

Exports enabled users (UPN, primary SMTP, proxies), distribution groups, and memberships to `C:\MigrationData\SourceTenant\*.csv`. Review the `ActiveUsers` CSV before proceeding — it is the input contract for phase 2.

### Phase 2 — Staging / Sync (Target)

Reconnect to the **target** tenant, then stage and submit:

```powershell
Disconnect-MgGraph
Disconnect-ExchangeOnline -Confirm:$false
Connect-MgGraph -Scopes 'User.Read.All', 'User.ReadWrite.All'
Connect-ExchangeOnline

./scripts/Step2_TargetStagingAndBatch.ps1 `
    -SourceCsvPath 'C:\MigrationData\SourceTenant\ActiveUsers_<timestamp>.csv' `
    -SourceOnMicrosoftDomain 'sourcetenant.onmicrosoft.com' `
    -TargetOnMicrosoftDomain 'targettenant.onmicrosoft.com' `
    -MigrationEndpointName 'T2T-SourceEndpoint'
```

The script creates target cloud users via Graph, mail-enables them as MailUsers pointing at the source mailboxes, writes the mapping CSV to `C:\MigrationData\TargetTenant\`, and submits the batch **without auto-starting it**. Review, then start and monitor:

```powershell
Start-MigrationBatch -Identity 'T2T-sourcetenant-<timestamp>'
Get-MigrationBatch -Identity 'T2T-sourcetenant-<timestamp>' | Format-List Status, *Count*
Get-MigrationUser -BatchId 'T2T-sourcetenant-<timestamp>' | Get-MigrationUserStatistics
```

### Phase 3 — Cutover

When every migration user reports healthy sync, complete the batch and license the target users:

```powershell
Set-MigrationBatch -Identity 'T2T-sourcetenant-<timestamp>' -CompleteAfter (Get-Date)
```

Then reconnect to the **source** tenant and strip the custom domain from all source identities:

```powershell
./scripts/Step3_SourceDomainCleanup.ps1 `
    -CustomDomain 'contoso.com' `
    -SourceOnMicrosoftDomain 'sourcetenant.onmicrosoft.com'
```

The script rewrites UPNs and primary SMTP addresses to the source `.onmicrosoft.com` domain across mailboxes, mail users, contacts, distribution groups, and Microsoft 365 groups, then prints the remaining manual checklist: set the initial domain as default, remove the custom domain from the source tenant, verify it in the target tenant, and cut DNS over.

## Safety & Verification

- **Pre-flight validation** — engine, module versions, Graph scopes, and EXO connectivity are checked before any tenant operation; failures are terminating.
- **Transcript logging** — every step writes a full transcript to `C:\MigrationData\Logs\` for audit and rollback analysis.
- **Tenant confirmation gates** — Steps 2 and 3 display the connected tenant and require explicit operator confirmation (`-Force` to bypass for automation), preventing wrong-tenant writes.
- **`-WhatIf` / `-Confirm` support** — all mutating steps support dry runs; migration batches are never auto-started.
- **Idempotent re-runs** — existing target users are reused, ExternalEmailAddress drift is corrected, and Step 3 isolates failures per object and reports them in a CSV.
- **Lab testing protocol** — the full sequence was rehearsed against a disposable lab tenant pair, including a pilot batch, sync-health verification via `Get-MigrationUserStatistics`, and a domain-release dry run, before production deployment. Re-run this protocol for every new tenant pair.
- **Data hygiene** — tenant exports contain PII; `*.csv`, logs, and `MigrationData/` are git-ignored and must never be committed.

## Author

**Elmar America** — Senior Microsoft Cloud Engineer