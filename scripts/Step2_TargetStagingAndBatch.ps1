#Requires -Version 7.0

<#
.SYNOPSIS
    Phase 2 - Staging/Sync: provisions target identities and submits a cross-tenant migration batch.

.DESCRIPTION
    Run this script while connected to the TARGET tenant. For each row of the
    Step 1 ActiveUsers export it:

      1. Creates a cloud user in the target tenant via Microsoft Graph
         (<alias>@<TargetOnMicrosoftDomain>, random temporary password, forced
         password change at next sign-in). Existing users are reused.
      2. Mail-enables the user in Exchange Online as a MailUser whose
         ExternalEmailAddress points at the source mailbox - the object type
         required by native cross-tenant mailbox moves.
      3. Writes the migration mapping CSV (EmailAddress column) plus an
         auditable provisioning report to C:\MigrationData\TargetTenant.
      4. Submits the batch with New-MigrationBatch against the pre-created
         cross-tenant migration endpoint. The batch is NOT auto-started;
         review it, then run Start-MigrationBatch.

    Prerequisites in the target tenant (one-time, see README): app registration
    consented in the source tenant, organization relationship, and a migration
    endpoint created with New-MigrationEndpoint -ExchangeRemoteMove.

.PARAMETER SourceCsvPath
    Path to the ActiveUsers CSV produced by Step1_SourceExport.ps1.

.PARAMETER SourceOnMicrosoftDomain
    Source tenant initial domain, e.g. sourcetenant.onmicrosoft.com.

.PARAMETER TargetOnMicrosoftDomain
    Target tenant initial domain, e.g. targettenant.onmicrosoft.com.

.PARAMETER MigrationEndpointName
    Name of the existing cross-tenant migration endpoint in the target tenant.

.PARAMETER LogDirPath
    Directory for the transcript log. Default: C:\MigrationData\Logs

.PARAMETER Force
    Bypass the interactive target-tenant confirmation prompt.

.EXAMPLE
    ./Step2_TargetStagingAndBatch.ps1 -SourceCsvPath 'C:\MigrationData\SourceTenant\ActiveUsers_20261109_120000.csv' `
        -SourceOnMicrosoftDomain 'sourcetenant.onmicrosoft.com' `
        -TargetOnMicrosoftDomain 'targettenant.onmicrosoft.com' `
        -MigrationEndpointName 'T2T-SourceEndpoint'
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory = $true)]
    [ValidateScript({ Test-Path $_ -PathType Leaf })]
    [string]$SourceCsvPath,

    [Parameter(Mandatory = $true)]
    [ValidatePattern('\.onmicrosoft\.com$')]
    [string]$SourceOnMicrosoftDomain,

    [Parameter(Mandatory = $true)]
    [ValidatePattern('\.onmicrosoft\.com$')]
    [string]$TargetOnMicrosoftDomain,

    [Parameter(Mandatory = $true)]
    [string]$MigrationEndpointName,

    [string]$LogDirPath = 'C:\MigrationData\Logs',

    [switch]$Force
)

function New-TemporaryPassword {
    $chars = [char[]]((48..57) + (65..90) + (97..122) + @(33, 35, 36, 37, 38, 42, 43, 45, 61, 63, 64))
    $bytes = [byte[]]::new(24)
    [System.Security.Cryptography.RandomNumberGenerator]::Fill($bytes)
    -join ($bytes | ForEach-Object { $chars[$_ % $chars.Length] })
}

$ErrorActionPreference = 'Stop'
. "$PSScriptRoot\Test-MigrationPreflight.ps1"

$timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$stagingDir = 'C:\MigrationData\TargetTenant'
New-Item -ItemType Directory -Path $stagingDir, $LogDirPath -Force | Out-Null
Start-Transcript -Path (Join-Path $LogDirPath "Step2_TargetStagingAndBatch_$timestamp.txt") | Out-Null

$graphOpenedHere = $false
$exoOpenedHere = $false

try {
    # --- Input validation -------------------------------------------------------
    $sourceUsers = @(Import-Csv -Path $SourceCsvPath)
    if ($sourceUsers.Count -eq 0) { throw "Source CSV '$SourceCsvPath' contains no rows." }

    $requiredColumns = @('DisplayName', 'UserPrincipalName', 'PrimarySmtpAddress')
    $actualColumns = $sourceUsers[0].PSObject.Properties.Name
    foreach ($column in $requiredColumns) {
        if ($column -notin $actualColumns) {
            throw "Source CSV is missing required column '$column'. Re-run Step1_SourceExport.ps1."
        }
    }

    Import-Module Microsoft.Graph.Authentication, Microsoft.Graph.Users -ErrorAction Stop
    Import-Module ExchangeOnlineManagement -ErrorAction Stop

    if (-not (Get-MgContext)) {
        Connect-MgGraph -Scopes 'User.Read.All', 'User.ReadWrite.All' -NoWelcome
        $graphOpenedHere = $true
    }

    $exoConnected = $false
    try {
        $exoConnected = [bool](Get-ConnectionInformation -ErrorAction Stop |
            Where-Object { $_.State -eq 'Connected' })
    }
    catch {
        $exoConnected = $false
    }
    if (-not $exoConnected) {
        Connect-ExchangeOnline -ShowBanner:$false
        $exoOpenedHere = $true
    }

    Test-MigrationPreflight -RequiredGraphScopes @('User.ReadWrite.All')

    # --- Tenant safety gate ------------------------------------------------------
    $graphContext = Get-MgContext
    $exoSession = Get-ConnectionInformation | Where-Object { $_.State -eq 'Connected' } | Select-Object -First 1
    Write-Host "Graph tenant : $($graphContext.TenantId)"
    Write-Host "EXO account  : $($exoSession.UserPrincipalName)"
    if (-not $Force) {
        $confirmed = $PSCmdlet.ShouldContinue(
            "Staging will create users and submit a migration batch in the tenant shown above. Is this the TARGET tenant ($TargetOnMicrosoftDomain)?",
            'Confirm target tenant')
        if (-not $confirmed) { throw 'Operator aborted: verify you are connected to the target tenant before re-running.' }
    }

    # Fail fast if the endpoint does not exist, before provisioning anything.
    $endpoint = Get-MigrationEndpoint -Identity $MigrationEndpointName

    # --- Provision target identities ----------------------------------------------
    $mapping = [System.Collections.Generic.List[pscustomobject]]::new()
    $report = [System.Collections.Generic.List[pscustomobject]]::new()

    foreach ($row in $sourceUsers) {
        $sourceUpn = $row.UserPrincipalName
        $sourceSmtp = $row.PrimarySmtpAddress

        if ([string]::IsNullOrWhiteSpace($sourceSmtp)) {
            Write-Warning "Skipping '$sourceUpn': no primary SMTP address in source export (no mailbox to migrate)."
            $report.Add([pscustomobject]@{ SourceUserPrincipalName = $sourceUpn; SourcePrimarySmtp = ''; TargetUserPrincipalName = ''; Action = 'Skipped'; Status = 'No source mailbox' })
            continue
        }

        $aliasSeed = if ($row.MailNickname) { $row.MailNickname } else { ($sourceUpn -split '@')[0] }
        $alias = ($aliasSeed -replace '[^a-zA-Z0-9._-]', '').ToLowerInvariant()
        if ($alias.Length -gt 64) { $alias = $alias.Substring(0, 64) }
        if (-not $alias) {
            Write-Warning "Skipping '$sourceUpn': could not derive a valid target alias."
            $report.Add([pscustomobject]@{ SourceUserPrincipalName = $sourceUpn; SourcePrimarySmtp = $sourceSmtp; TargetUserPrincipalName = ''; Action = 'Skipped'; Status = 'Invalid alias' })
            continue
        }
        $targetUpn = '{0}@{1}' -f $alias, $TargetOnMicrosoftDomain

        $action = 'Reused existing user'
        $status = 'Staged'

        if ($PSCmdlet.ShouldProcess($targetUpn, "Stage target identity for source mailbox '$sourceSmtp'")) {
            $existingUser = Get-MgUser -Filter "userPrincipalName eq '$targetUpn'" | Select-Object -First 1
            if (-not $existingUser) {
                $passwordProfile = @{
                    Password                      = (New-TemporaryPassword)
                    ForceChangePasswordNextSignIn = $true
                }
                $existingUser = New-MgUser -DisplayName $row.DisplayName `
                    -UserPrincipalName $targetUpn `
                    -MailNickname $alias `
                    -AccountEnabled:$true `
                    -PasswordProfile $passwordProfile
                $action = 'Created user'
                Write-Host "Created target user $targetUpn"
            }

            # New Graph users take a moment to surface in Exchange Online.
            $exoUser = $null
            for ($attempt = 1; $attempt -le 12; $attempt++) {
                $exoUser = Get-User -Identity $targetUpn -ErrorAction SilentlyContinue
                if ($exoUser) { break }
                if ($attempt -eq 1) { Write-Host "  Waiting for $targetUpn to synchronize into Exchange Online..." }
                Start-Sleep -Seconds 10
            }
            if (-not $exoUser) {
                throw "Target user '$targetUpn' did not appear in Exchange Online after 120 seconds. Re-run the script; already-created users are reused."
            }

            $recipient = Get-Recipient -Identity $targetUpn -ErrorAction SilentlyContinue
            if ($recipient -and $recipient.RecipientTypeDetails -eq 'UserMailbox') {
                Write-Warning "Skipping '$targetUpn': target already has a mailbox. Cross-tenant moves require a MailUser target."
                $report.Add([pscustomobject]@{ SourceUserPrincipalName = $sourceUpn; SourcePrimarySmtp = $sourceSmtp; TargetUserPrincipalName = $targetUpn; Action = 'Skipped'; Status = 'Target already has mailbox' })
                continue
            }

            if (-not $recipient) {
                Enable-MailUser -Identity $targetUpn -ExternalEmailAddress $sourceSmtp | Out-Null
                $action += '; mail-enabled'
            }
            elseif ($recipient.RecipientTypeDetails -eq 'MailUser') {
                $mailUser = Get-MailUser -Identity $targetUpn
                $currentExternal = $mailUser.ExternalEmailAddress -replace '^SMTP:', ''
                if ($currentExternal -ne $sourceSmtp) {
                    Set-MailUser -Identity $targetUpn -ExternalEmailAddress $sourceSmtp | Out-Null
                    $action += '; corrected ExternalEmailAddress'
                }
            }
            else {
                Write-Warning "Skipping '$targetUpn': unexpected recipient type '$($recipient.RecipientTypeDetails)'."
                $report.Add([pscustomobject]@{ SourceUserPrincipalName = $sourceUpn; SourcePrimarySmtp = $sourceSmtp; TargetUserPrincipalName = $targetUpn; Action = 'Skipped'; Status = "Unexpected recipient type $($recipient.RecipientTypeDetails)" })
                continue
            }

            $mapping.Add([pscustomobject]@{ EmailAddress = $targetUpn })
            Write-Host "Staged $targetUpn <- $sourceSmtp"
        }
        else {
            $action = 'WhatIf - no changes made'
            $status = 'Not staged'
            $mapping.Add([pscustomobject]@{ EmailAddress = $targetUpn })
        }

        $report.Add([pscustomobject]@{
            SourceUserPrincipalName = $sourceUpn
            SourcePrimarySmtp       = $sourceSmtp
            TargetUserPrincipalName = $targetUpn
            Action                  = $action
            Status                  = $status
        })
    }

    if ($mapping.Count -eq 0) { throw 'No eligible mailboxes were staged; migration batch was not created.' }

    # --- Mapping + report CSVs -----------------------------------------------------
    $mappingCsv = Join-Path $stagingDir "MigrationMapping_$timestamp.csv"
    $mapping | Export-Csv -Path $mappingCsv -NoTypeInformation -Encoding UTF8

    $reportCsv = Join-Path $stagingDir "ProvisioningReport_$timestamp.csv"
    $report | Export-Csv -Path $reportCsv -NoTypeInformation -Encoding UTF8

    Write-Host "Mapping CSV : $mappingCsv ($($mapping.Count) entries)"
    Write-Host "Audit report: $reportCsv"

    # --- Submit migration batch (deliberately not auto-started) ---------------------
    $batchName = 'T2T-{0}-{1}' -f $SourceOnMicrosoftDomain.Split('.')[0], $timestamp
    if ($PSCmdlet.ShouldProcess($batchName, "Submit cross-tenant migration batch with $($mapping.Count) mailbox(es) via endpoint '$MigrationEndpointName'")) {
        $csvBytes = [System.IO.File]::ReadAllBytes($mappingCsv)
        $batch = New-MigrationBatch -Name $batchName `
            -SourceEndpoint $endpoint.Identity `
            -TargetDeliveryDomain $TargetOnMicrosoftDomain `
            -CSVData $csvBytes
        Write-Host "Migration batch submitted: $($batch.Identity) (Status: $($batch.Status))" -ForegroundColor Green
    }

    Write-Host ''
    Write-Host 'Next steps:'
    Write-Host "  1. Review the batch, then start sync : Start-MigrationBatch -Identity '$batchName'"
    Write-Host "  2. Monitor progress                 : Get-MigrationBatch -Identity '$batchName' | Format-List"
    Write-Host "                                        Get-MigrationUser -BatchId '$batchName' | Get-MigrationUserStatistics"
    Write-Host "  3. When sync health is verified     : Set-MigrationBatch -Identity '$batchName' -CompleteAfter (Get-Date)"
    Write-Host '  4. After completion, license the target users, then run Step3_SourceDomainCleanup.ps1 in the SOURCE tenant.'
}
catch {
    Write-Error "Target staging failed: $($_.Exception.Message)"
    throw
}
finally {
    if ($exoOpenedHere) { Disconnect-ExchangeOnline -Confirm:$false }
    if ($graphOpenedHere) { Disconnect-MgGraph | Out-Null }
    Stop-Transcript | Out-Null
}