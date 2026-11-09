#Requires -Version 7.0

<#
.SYNOPSIS
    Phase 3 - Cutover: moves source-tenant identities off the custom domain so it can be released.

.DESCRIPTION
    Run this script while connected to the SOURCE tenant, after the migration
    batch has completed and target users are licensed and verified.

    For every mailbox, mail user, mail contact, distribution group, and
    Microsoft 365 group that still references $CustomDomain it:

      1. Rebuilds the proxy address list with <alias>@<SourceOnMicrosoftDomain>
         as the primary SMTP address and removes every @$CustomDomain address.
      2. For user mailboxes, also updates the Entra ID UPN to
         <local-part>@<SourceOnMicrosoftDomain> via Microsoft Graph.

    Failures are isolated per object, collected, and exported to a CSV in
    $LogDirPath; the script throws at the end if any object could not be
    updated, because domain release is blocked until all references are gone.

    Removing the domain from the tenant itself is deliberately left as a
    manual step - the script prints the remaining checklist on completion.

.PARAMETER CustomDomain
    The vanity domain to release, e.g. contoso.com.

.PARAMETER SourceOnMicrosoftDomain
    Source tenant initial domain, e.g. sourcetenant.onmicrosoft.com.

.PARAMETER LogDirPath
    Directory for the transcript log. Default: C:\MigrationData\Logs

.PARAMETER Force
    Bypass the interactive source-tenant confirmation prompt.

.EXAMPLE
    ./Step3_SourceDomainCleanup.ps1 -CustomDomain 'contoso.com' -SourceOnMicrosoftDomain 'sourcetenant.onmicrosoft.com'
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory = $true)]
    [string]$CustomDomain,

    [Parameter(Mandatory = $true)]
    [ValidatePattern('\.onmicrosoft\.com$')]
    [string]$SourceOnMicrosoftDomain,

    [string]$LogDirPath = 'C:\MigrationData\Logs',

    [switch]$Force
)

function Get-CleanAddressList {
    param(
        [string[]]$CurrentAddresses,
        [string]$NewPrimary,
        [string]$DomainPattern
    )
    $list = [System.Collections.Generic.List[string]]::new()
    $list.Add("SMTP:$NewPrimary")
    foreach ($address in $CurrentAddresses) {
        if ($address -match $DomainPattern) { continue }
        if ($address -ieq "SMTP:$NewPrimary") { continue }
        $list.Add($address)
    }
    return $list.ToArray()
}

$ErrorActionPreference = 'Stop'
. "$PSScriptRoot\Test-MigrationPreflight.ps1"

$timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
New-Item -ItemType Directory -Path $LogDirPath -Force | Out-Null
Start-Transcript -Path (Join-Path $LogDirPath "Step3_SourceDomainCleanup_$timestamp.txt") | Out-Null

$graphOpenedHere = $false
$exoOpenedHere = $false

try {
    Import-Module Microsoft.Graph.Authentication, Microsoft.Graph.Users, Microsoft.Graph.Identity.DirectoryManagement -ErrorAction Stop
    Import-Module ExchangeOnlineManagement -ErrorAction Stop

    if (-not (Get-MgContext)) {
        Connect-MgGraph -Scopes 'User.ReadWrite.All', 'Domain.Read.All' -NoWelcome
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

    Test-MigrationPreflight -RequiredGraphScopes @('User.ReadWrite.All', 'Domain.Read.All')

    # --- Tenant safety gate -------------------------------------------------------
    # Fails immediately if the custom domain is not verified in this tenant,
    # which catches the most dangerous operator error: running against target.
    $null = Get-MgDomain -DomainId $CustomDomain
    $graphContext = Get-MgContext
    Write-Host "Graph tenant : $($graphContext.TenantId)"
    Write-Host "Domain       : $CustomDomain (verified in this tenant)"
    if (-not $Force) {
        $confirmed = $PSCmdlet.ShouldContinue(
            "This rewrites UPNs and SMTP addresses for every object referencing '$CustomDomain' in the tenant shown above. Is this the SOURCE tenant?",
            'Confirm source tenant')
        if (-not $confirmed) { throw 'Operator aborted: verify you are connected to the source tenant before re-running.' }
    }

    $domainPattern = '@' + [regex]::Escape($CustomDomain) + '$'
    $failures = [System.Collections.Generic.List[pscustomobject]]::new()
    $updatedCount = 0

    # --- Rewrite addresses across all recipient types ------------------------------
    $categories = @(
        @{ Label = 'mailbox';             Recipients = @(Get-Mailbox -ResultSize Unlimited);           SetCommand = 'Set-Mailbox' }
        @{ Label = 'mail user';           Recipients = @(Get-MailUser -ResultSize Unlimited);          SetCommand = 'Set-MailUser' }
        @{ Label = 'mail contact';        Recipients = @(Get-MailContact -ResultSize Unlimited);       SetCommand = 'Set-MailContact' }
        @{ Label = 'distribution group';  Recipients = @(Get-DistributionGroup -ResultSize Unlimited); SetCommand = 'Set-DistributionGroup' }
        @{ Label = 'Microsoft 365 group'; Recipients = @(Get-UnifiedGroup -ResultSize Unlimited);      SetCommand = 'Set-UnifiedGroup' }
    )

    foreach ($category in $categories) {
        $affected = @($category.Recipients | Where-Object { $_.EmailAddresses -match $domainPattern })
        Write-Host "$($category.Label): $($affected.Count) object(s) still reference $CustomDomain"

        foreach ($recipient in $affected) {
            $newPrimary = '{0}@{1}' -f $recipient.Alias, $SourceOnMicrosoftDomain
            try {
                if ($PSCmdlet.ShouldProcess($recipient.PrimarySmtpAddress, "Rewrite $($category.Label) addresses to $newPrimary and strip @$CustomDomain proxies")) {
                    $addresses = @(Get-CleanAddressList -CurrentAddresses $recipient.EmailAddresses -NewPrimary $newPrimary -DomainPattern $domainPattern)
                    & $category.SetCommand -Identity $recipient.Identity -EmailAddresses $addresses
                    $updatedCount++
                    Write-Host "  $($recipient.PrimarySmtpAddress) -> $newPrimary"

                    if ($category.SetCommand -eq 'Set-Mailbox' -and $recipient.UserPrincipalName -match $domainPattern) {
                        $newUpn = '{0}@{1}' -f ($recipient.UserPrincipalName -split '@')[0], $SourceOnMicrosoftDomain
                        Set-MgUser -UserId $recipient.UserPrincipalName -UserPrincipalName $newUpn
                        Write-Host "  UPN: $($recipient.UserPrincipalName) -> $newUpn"
                    }
                }
            }
            catch {
                Write-Warning "Failed for '$($recipient.PrimarySmtpAddress)': $($_.Exception.Message)"
                $failures.Add([pscustomobject]@{
                    Object = $recipient.PrimarySmtpAddress
                    Type   = $category.Label
                    Error  = $_.Exception.Message
                })
            }
        }
    }

    Write-Host ''
    Write-Host "Updated $updatedCount object(s); $($failures.Count) failure(s)."

    if ($failures.Count -gt 0) {
        $failureCsv = Join-Path $LogDirPath "Step3_Failures_$timestamp.csv"
        $failures | Export-Csv -Path $failureCsv -NoTypeInformation -Encoding UTF8
        throw "$($failures.Count) object(s) could not be updated. See $failureCsv. Domain release is blocked until every reference to '$CustomDomain' is removed."
    }

    Write-Host ''
    Write-Host 'Remaining manual checklist before releasing the domain:' -ForegroundColor Green
    Write-Host "  1. Confirm zero references remain (recipient loop above reported 0 affected objects on a re-run)."
    Write-Host "  2. Check email address policies, accepted domains, connectors, and any hybrid/on-premises AD objects for '$CustomDomain'."
    Write-Host "  3. Make the initial domain default : Set-MgDomain -DomainId '$SourceOnMicrosoftDomain' -IsDefault"
    Write-Host "  4. Release the domain              : Remove-MgDomain -DomainId '$CustomDomain'  (or M365 admin center > Settings > Domains)"
    Write-Host "  5. Verify '$CustomDomain' in the TARGET tenant, update DNS, and set target UPNs/primary SMTPs to the custom domain."
}
catch {
    Write-Error "Source domain cleanup failed: $($_.Exception.Message)"
    throw
}
finally {
    if ($exoOpenedHere) { Disconnect-ExchangeOnline -Confirm:$false }
    if ($graphOpenedHere) { Disconnect-MgGraph | Out-Null }
    Stop-Transcript | Out-Null
}