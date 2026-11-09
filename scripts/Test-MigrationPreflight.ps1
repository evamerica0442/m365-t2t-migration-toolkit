#Requires -Version 7.0

function Test-MigrationPreflight {
    <#
    .SYNOPSIS
        Validates the engine, modules, and service connections required by the T2T migration scripts.

    .DESCRIPTION
        Runs a fixed set of pre-flight checks and reports each as Passed/Failed:
          1. PowerShell engine is 7.0 or later.
          2. Microsoft.Graph module (>= 2.0.0) is installed.
          3. ExchangeOnlineManagement module (>= 3.0.0) is installed.
          4. An active Microsoft Graph context exists with the required delegated scopes.
          5. An active Exchange Online session exists.

        Throws a terminating error if any check fails, so calling scripts stop
        before touching tenant data. Dot-source this file, then call the function:

            . "$PSScriptRoot\Test-MigrationPreflight.ps1"
            Test-MigrationPreflight -RequiredGraphScopes @('User.ReadWrite.All')

        Running the file directly executes the function with default read-only scopes.

    .PARAMETER RequiredGraphScopes
        Delegated scopes that must be present in the current Graph context.

    .PARAMETER SkipGraph
        Skip the Graph context/scope check (module version is still validated).

    .PARAMETER SkipExchange
        Skip the Exchange Online connection check (module version is still validated).

    .OUTPUTS
        PSCustomObject[] - one record per check: Check, Passed, Detail.
    #>
    [CmdletBinding()]
    param(
        [string[]]$RequiredGraphScopes = @('User.Read.All', 'Group.Read.All', 'Domain.Read.All'),
        [switch]$SkipGraph,
        [switch]$SkipExchange
    )

    $checks = [System.Collections.Generic.List[pscustomobject]]::new()

    # 1. PowerShell engine
    $engineOk = $PSVersionTable.PSVersion -ge [version]'7.0'
    $checks.Add([pscustomobject]@{
        Check  = 'PowerShell engine >= 7.0'
        Passed = $engineOk
        Detail = "Running $($PSVersionTable.PSVersion) on $($PSVersionTable.PSEdition)"
    })

    # 2/3. Module availability and minimum versions
    $moduleRequirements = @(
        @{ Name = 'Microsoft.Graph';          Minimum = [version]'2.0.0' }
        @{ Name = 'ExchangeOnlineManagement'; Minimum = [version]'3.0.0' }
    )
    foreach ($req in $moduleRequirements) {
        $installed = Get-Module -ListAvailable -Name $req.Name |
            Sort-Object -Property Version -Descending |
            Select-Object -First 1

        if (-not $installed) {
            $checks.Add([pscustomobject]@{
                Check  = "$($req.Name) >= $($req.Minimum)"
                Passed = $false
                Detail = "Not installed. Run: Install-Module $($req.Name) -MinimumVersion $($req.Minimum) -Scope CurrentUser"
            })
        }
        else {
            $checks.Add([pscustomobject]@{
                Check  = "$($req.Name) >= $($req.Minimum)"
                Passed = ($installed.Version -ge $req.Minimum)
                Detail = "Installed version $($installed.Version)"
            })
        }
    }

    # 4. Graph context and delegated scopes
    if (-not $SkipGraph) {
        $passed = $false
        $detail = 'No active Graph context. Run Connect-MgGraph with the required scopes.'
        try {
            Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
            $context = Get-MgContext
            if ($context) {
                $missingScopes = @($RequiredGraphScopes | Where-Object { $_ -notin $context.Scopes })
                if ($missingScopes.Count -gt 0) {
                    $detail = "Context present but missing scopes: $($missingScopes -join ', '). Re-run Connect-MgGraph."
                }
                else {
                    $passed = $true
                    $detail = "Tenant $($context.TenantId), account $($context.Account); all required scopes present."
                }
            }
        }
        catch {
            $detail = "Graph context check failed: $($_.Exception.Message)"
        }
        $checks.Add([pscustomobject]@{
            Check  = 'Microsoft Graph context + scopes'
            Passed = $passed
            Detail = $detail
        })
    }

    # 5. Exchange Online session
    if (-not $SkipExchange) {
        $passed = $false
        $detail = 'No active Exchange Online session. Run Connect-ExchangeOnline.'
        try {
            Import-Module ExchangeOnlineManagement -ErrorAction Stop
            $session = Get-ConnectionInformation |
                Where-Object { $_.State -eq 'Connected' } |
                Select-Object -First 1
            if ($session) {
                $passed = $true
                $detail = "Connected as $($session.UserPrincipalName)."
            }
        }
        catch {
            $detail = "Exchange Online check failed: $($_.Exception.Message)"
        }
        $checks.Add([pscustomobject]@{
            Check  = 'Exchange Online connection'
            Passed = $passed
            Detail = $detail
        })
    }

    $checks | Format-Table -AutoSize -Wrap | Out-String -Width 200 | Write-Host

    $failed = @($checks | Where-Object { -not $_.Passed })
    if ($failed.Count -gt 0) {
        throw "Pre-flight validation failed: $($failed.Count) check(s) did not pass. Resolve the items above before running migration steps."
    }

    Write-Host 'Pre-flight validation passed.' -ForegroundColor Green
    return $checks
}

# Execute with defaults when run directly; stay silent when dot-sourced.
if ($MyInvocation.InvocationName -ne '.') {
    Test-MigrationPreflight
}