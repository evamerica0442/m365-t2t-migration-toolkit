# Architecture — T2T Migration Automation Framework

## Overview

The framework orchestrates a **native cross-tenant mailbox migration** between two
independent Microsoft 365 tenants. Control-plane operations (identity discovery,
provisioning, address rewrites) run over **Microsoft Graph** and **Exchange Online
Management v3** from the operator's workstation; the data-plane mailbox move runs
server-side inside Exchange Online's Migration Service (MRS) over an app-only
OAuth trust established by a cross-tenant migration endpoint.

No migration data ever passes through the operator machine except CSV exports
and transcripts. No credentials are stored by the scripts.

## Cross-Tenant Workflow

```
      SOURCE TENANT A                                  TARGET TENANT B
      (contoso.com)                                    (target.onmicrosoft.com)
+--------------------------------------+   +--------------------------------------+
|  Entra ID                            |   |  Entra ID                            |
|    users / groups @contoso.com       |   |    staged users @target.onmicrosoft  |
|  Exchange Online                     |   |  Exchange Online                     |
|    live user mailboxes               |   |    MailUsers (pre-migration)         |
+------------------+-------------------+   +------------------+-------------------+
                   |                                          ^
  [1] DISCOVERY    |  Graph + EXO, read-only                  |
      Step1        v                                          |
        +------------------------------+                     |
        |  C:\MigrationData\SourceTenant |    [2] STAGING      |
        |    ActiveUsers_*.csv           +--------------------->  Step2
        |    DistributionGroups_*.csv    |                     |  Graph: New-MgUser
        |    GroupMemberships_*.csv      |                     |  EXO: Enable-MailUser
        +------------------------------+                     |       (ExternalEmailAddress
                                                             |        -> source mailbox)
                                                             |
                                                             |  mapping CSV (EmailAddress)
                                                             |  New-MigrationBatch
                   +-----------------------------------------+
                   |   [2] SYNC — Exchange MRS cross-tenant move
                   |         app-only OAuth via migration endpoint
                   |         (New-MigrationEndpoint -ExchangeRemoteMove)
                   v
+--------------------------------------+   +--------------------------------------+
|  Source mailboxes drain; after batch |   |  Target MailUsers convert to live    |
|  completion they become MailUsers    |   |  mailboxes; assign licenses, verify  |
+------------------+-------------------+   +--------------------------------------+
                   |
  [3] CUTOVER      |  Step3: UPN + primary SMTP rewritten
      Step3        |         -> sourcetenant.onmicrosoft.com
                   v
        Release contoso.com from Tenant A
                   |
                   +------------------------->  Verify contoso.com in Tenant B,
                                                set target UPN/SMTP, cut DNS over
```

## Phase / Script Responsibility Matrix

| Phase | Script | Connected tenant | Plane | Mutates? |
|---|---|---|---|---|
| 0 — Pre-flight | `Test-MigrationPreflight.ps1` | Either | Local + auth context | No |
| 1 — Discovery | `Step1_SourceExport.ps1` | Source A | Graph (read) + EXO (read) | No — CSV exports only |
| 2 — Staging/Sync | `Step2_TargetStagingAndBatch.ps1` | Target B | Graph (write) + EXO (write) + MRS | Yes — users, MailUsers, batch |
| 3 — Cutover | `Step3_SourceDomainCleanup.ps1` | Source A | Graph (write) + EXO (write) | Yes — UPNs, proxy addresses |

## Data Artifacts

| Path | Produced by | Consumed by | Contents |
|---|---|---|---|
| `C:\MigrationData\SourceTenant\ActiveUsers_*.csv` | Step 1 | Step 2 (`-SourceCsvPath`) | DisplayName, UPN, MailNickname, primary SMTP, proxies, mailbox flag |
| `C:\MigrationData\SourceTenant\DistributionGroups_*.csv` | Step 1 | Operator review | Group identity, primary SMTP, type |
| `C:\MigrationData\SourceTenant\GroupMemberships_*.csv` | Step 1 | Operator review | Flattened group → member edges |
| `C:\MigrationData\TargetTenant\MigrationMapping_*.csv` | Step 2 | `New-MigrationBatch -CSVData` | Single `EmailAddress` column: target MailUser addresses |
| `C:\MigrationData\TargetTenant\ProvisioningReport_*.csv` | Step 2 | Audit | Source → target mapping with per-user action/status |
| `C:\MigrationData\Logs\*.txt` | All steps | Audit | Full transcripts |
| `C:\MigrationData\Logs\Step3_Failures_*.csv` | Step 3 (on error) | Operator remediation | Objects that resisted address rewrite |

All artifacts are git-ignored: they contain tenant PII and must never enter source control.

## Authentication Model

- **Control plane (operator):** delegated Microsoft Graph scopes, least-privilege per phase —
  read-only (`User.Read.All`, `Group.Read.All`, `Domain.Read.All`) for discovery;
  `User.ReadWrite.All` for staging and cutover. Exchange Online sessions use the
  operator's admin account via the EXO v3 module.
- **Data plane (MRS):** app-only OAuth. The target tenant's app registration is
  consented in the source tenant with Exchange `full_access_as_app`; the migration
  endpoint (`New-MigrationEndpoint -ExchangeRemoteMove`) binds the two. The scripts
  never handle the app secret — it lives only in the endpoint configuration.
- Sessions opened by a script are closed in its `finally` block; pre-existing
  operator sessions are left intact.

## Key Design Decisions

1. **MailUser staging, not licensed mailboxes.** Native cross-tenant moves target
   MailUser objects whose `ExternalEmailAddress` identifies the source mailbox.
   Step 2 creates unlicensed cloud users via Graph, then mail-enables them in EXO.
   Licensing happens only after cutover, avoiding double-paid mailboxes.
2. **Batches are never auto-started.** Submission and sync start are separate
   operator decisions, giving a review checkpoint between staging and data movement.
3. **Tenant confirmation gates.** Steps 2 and 3 print the connected tenant and
   require explicit confirmation; Step 3 additionally fails fast if the custom
   domain is not verified in the connected tenant (wrong-tenant protection).
4. **Idempotency.** Re-running Step 2 reuses existing users and corrects
   `ExternalEmailAddress` drift; re-running Step 3 only touches objects that still
   reference the custom domain.
5. **Failure isolation.** Step 3 catches per object, exports a failure CSV, and
   throws only at the end — domain release is blocked until the failure list is empty.
6. **Domain release stays manual.** Setting the default domain and removing the
   custom domain are tenant-wide, hard-to-reverse actions; the script prepares the
   directory and prints the checklist instead of executing them.

## Recovery Notes

- A failed Step 2 run can be re-run safely; already-created users are reused and
  the mapping CSV is regenerated with a fresh timestamp.
- If a batch misbehaves, remove it with `Remove-MigrationBatch` before resubmitting;
  target MailUsers remain valid staging objects.
- Step 3 failures are almost always address collisions (`alias@onmicrosoft` already
  in use). Resolve the collision, then re-run — completed objects are skipped.