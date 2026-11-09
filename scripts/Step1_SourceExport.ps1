#Requires -Version 7.0

<#
.SYNOPSIS
    Phase 1 - Discovery: exports users, mailboxes, distribution groups, and memberships from the SOURCE tenant.

.DESCRIPTION
    Connects to Microsoft Graph and Exchange Online in the source tenant (reusing
    existing sessions when present), validates the environment with
    Test-MigrationPreflight, and exports three CSV files to $ExportDir:

      ActiveUsers_<timestamp>.csv         - enabled member accounts joined to mailbox data
      DistributionGroups_<timestamp>.csv  - all distribution / mail-enabled security groups
      GroupMemberships_<timestamp>.csv    - flattened group membership

    All activity is transcribed to $LogDirPath. Connections opened by this script
    are closed in the finally block; pre-existing sessions are left untouched.

.PARAMETER ExportDir
    Destination directory for the CSV exports. Default: C:\MigrationData\SourceTenant

.PARAMETER LogDirPath
    Directory for the transcript log. Default: C:\MigrationData\Logs

.EXAMPLE
    ./Step1_SourceExport.ps1
#>

[CmdletBinding()]
param(
    [string]$ExportDir = 'C:\MigrationData\SourceTenant',
    [string]$LogDirPath = 'C:\MigrationData\Logs'
)

$ErrorActionPreference = 'Stop'
. "$PSScriptRoot\Test-MigrationPreflight.ps1"

$timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
New-Item -ItemType Directory -Path $ExportDir, $LogDirPath -Force | Out-Null
Start-Transcript -Path (Join-Path $LogDirPath "Step1_SourceExport_$timestamp.txt") | Out-Null

$graphOpenedHere = $false
$exoOpenedHere = $false

try {
    Import-Module Microsoft.Graph.Authentication, Microsoft.Graph.Users, Microsoft.Graph.Groups -ErrorAction Stop
    Import-Module ExchangeOnlineManagement -ErrorAction Stop

    if (-not (Get-MgContext)) {
        Connect-MgGraph -Scopes 'User.Read.All', 'Group.Read.All', 'Domain.Read.All' -NoWelcome
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

    Test-MigrationPreflight -RequiredGraphScopes @('User.Read.All', 'Group.Read.All', 'Domain.Read.All')

    # --- Active users joined to mailbox data -----------------------------------
    Write-Host 'Enumerating enabled member accounts (Microsoft Graph)...'
    $users = Get-MgUser -Filter "accountEnabled eq true and userType eq 'Member'" -All `
        -Property Id, DisplayName, UserPrincipalName, Mail, MailNickname, AccountEnabled

    Write-Host 'Enumerating mailboxes (Exchange Online)...'
    $mailboxes = Get-EXOMailbox -ResultSize Unlimited `
        -Properties UserPrincipalName, PrimarySmtpAddress, WindowsEmailAddress, EmailAddresses, RecipientTypeDetails

    $mailboxByUpn = @{}
    foreach ($mailbox in $mailboxes) {
        $mailboxByUpn[$mailbox.UserPrincipalName.ToLowerInvariant()] = $mailbox
    }

    $userRows = foreach ($user in $users) {
        $mailbox = $mailboxByUpn[$user.UserPrincipalName.ToLowerInvariant()]
        [pscustomobject]@{
            DisplayName         = $user.DisplayName
            UserPrincipalName   = $user.UserPrincipalName
            MailNickname        = $user.MailNickname
            PrimarySmtpAddress  = if ($mailbox) { $mailbox.PrimarySmtpAddress } else { $user.Mail }
            ProxyAddresses      = if ($mailbox) { @($mailbox.EmailAddresses | Where-Object { $_ -match '^smtp:' }) -join ';' } else { '' }
            RecipientTypeDetails = if ($mailbox) { $mailbox.RecipientTypeDetails } else { 'UserWithoutMailbox' }
            HasMailbox          = [bool]$mailbox
        }
    }
    $usersCsv = Join-Path $ExportDir "ActiveUsers_$timestamp.csv"
    $userRows | Export-Csv -Path $usersCsv -NoTypeInformation -Encoding UTF8

    # --- Distribution groups ----------------------------------------------------
    Write-Host 'Enumerating distribution groups...'
    $groups = Get-DistributionGroup -ResultSize Unlimited
    $groupsCsv = Join-Path $ExportDir "DistributionGroups_$timestamp.csv"
    $groups |
        Select-Object DisplayName, PrimarySmtpAddress, Alias, RecipientTypeDetails, GroupType |
        Export-Csv -Path $groupsCsv -NoTypeInformation -Encoding UTF8

    # --- Group memberships -------------------------------------------------------
    Write-Host 'Enumerating group memberships...'
    $membershipRows = foreach ($group in $groups) {
        try {
            Get-DistributionGroupMember -Identity $group.Identity -ResultSize Unlimited | ForEach-Object {
                [pscustomobject]@{
                    GroupDisplayName  = $group.DisplayName
                    GroupPrimarySmtp  = $group.PrimarySmtpAddress
                    MemberDisplayName = $_.DisplayName
                    MemberPrimarySmtp = $_.PrimarySmtpAddress
                    MemberType        = $_.RecipientTypeDetails
                }
            }
        }
        catch {
            Write-Warning "Could not read members of '$($group.DisplayName)': $($_.Exception.Message)"
        }
    }
    $membersCsv = Join-Path $ExportDir "GroupMemberships_$timestamp.csv"
    $membershipRows | Export-Csv -Path $membersCsv -NoTypeInformation -Encoding UTF8

    Write-Host ''
    Write-Host "Export complete: $(@($userRows).Count) users, $(@($groups).Count) groups, $(@($membershipRows).Count) memberships." -ForegroundColor Green
    Write-Host "  $usersCsv"
    Write-Host "  $groupsCsv"
    Write-Host "  $membersCsv"
    Write-Host ''
    Write-Host 'Next: review ActiveUsers CSV, then run Step2_TargetStagingAndBatch.ps1 against the TARGET tenant.'
}
catch {
    Write-Error "Source export failed: $($_.Exception.Message)"
    throw
}
finally {
    if ($exoOpenedHere) { Disconnect-ExchangeOnline -Confirm:$false }
    if ($graphOpenedHere) { Disconnect-MgGraph | Out-Null }
    Stop-Transcript | Out-Null
}