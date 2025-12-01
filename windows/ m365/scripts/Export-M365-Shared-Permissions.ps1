<#
.SYNOPSIS
    Export FullAccess permissions to shared mailboxes.
    (EN) Exports mapping between shared mailboxes and users with FullAccess.
    (PL) Eksportuje powiązania skrzynek współdzielonych z użytkownikami mającymi FullAccess.

.DESCRIPTION
    (EN) This script exports mapping between shared mailboxes and users
         who have explicit FullAccess permissions (non-inherited).

    (PL) Skrypt eksportuje mapowanie pomiędzy skrzynkami współdzielonymi a użytkownikami,
         którzy mają przyznane jawne uprawnienia FullAccess (nie dziedziczone).

.REQUIREMENTS
    - PowerShell 5.1+ (or PowerShell 7+)
    - ExchangeOnlineManagement module
    - Active Exchange Online session (Connect-ExchangeOnline)

.EXAMPLE
    .\Export-M365-Shared-Permissions.ps1
    .\Export-M365-Shared-Permissions.ps1 -OutputPath "C:\Reports\M365"
#>

param (
    [string]$OutputPath = ".\output"
)

Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force

$IsAdmin = ([Security.Principal.WindowsPrincipal] `
    [Security.Principal.WindowsIdentity]::GetCurrent() `
).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

if (-not $IsAdmin) {
    Write-Warning "This script is not running as Administrator."
    Write-Warning "Some actions (e.g. installing modules for all users) may fail."
    $choice = Read-Host "Continue anyway? (Y/N)"
    if ($choice -ne "Y") { exit }
}

if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {

    Write-Host "The required module 'ExchangeOnlineManagement' is not installed." -ForegroundColor Yellow
    $confirm = Read-Host "Do you want to install it now? (Y/N)"

    if ($confirm -eq "Y") {
        try {
            Write-Host "Installing module ExchangeOnlineManagement..." -ForegroundColor Cyan
            if ($IsAdmin) {
                Install-Module ExchangeOnlineManagement -Force -Scope AllUsers
            } else {
                Install-Module ExchangeOnlineManagement -Force -Scope CurrentUser
            }
            Write-Host "Module installed successfully." -ForegroundColor Green
        }
        catch {
            Write-Error "Module installation failed. Reason: $($_.Exception.Message)"
            exit
        }
    }
    else {
        Write-Host "Cannot continue without the module. Exiting." -ForegroundColor Red
        exit
    }
}

Import-Module ExchangeOnlineManagement -ErrorAction Stop

$exoSessionOk = $true
try {
    Get-EXOMailbox -ResultSize 1 -ErrorAction Stop | Out-Null
}
catch {
    $exoSessionOk = $false
}

if (-not $exoSessionOk) {
    Write-Warning "Not connected to Exchange Online."
    Write-Host "You need an active EXO session to run this script." -ForegroundColor Yellow
    $go = Read-Host "Do you want to connect now with Connect-ExchangeOnline? (Y/N)"
    if ($go -eq "Y") {
        try {
            Connect-ExchangeOnline
        } catch {
            Write-Error "Failed to connect to Exchange Online: $($_.Exception.Message)"
            exit
        }
    } else {
        Write-Host "Cannot continue without an active EXO session. Exiting." -ForegroundColor Red
        exit
    }
}

New-Item -ItemType Directory -Path $OutputPath -ErrorAction SilentlyContinue | Out-Null

$timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
$outFile   = Join-Path $OutputPath "M365_Shared-Mailbox-Permissions_$timestamp.csv"

Write-Host "Exporting shared mailbox permissions to: $outFile" -ForegroundColor Cyan

$results = @()

Get-EXOMailbox -RecipientTypeDetails SharedMailbox -ResultSize Unlimited | ForEach-Object {
    $mb = $_
    Write-Host "  Processing shared mailbox: $($mb.PrimarySmtpAddress)" -ForegroundColor Yellow

    $perms = Get-EXOMailboxPermission -Identity $mb.Identity -ErrorAction SilentlyContinue | Where-Object {
        -not $_.IsInherited -and
        $_.User -notlike "NT AUTHORITY\SELF" -and
        $_.User -notlike "S-1-5-*" -and
        $_.AccessRights -contains "FullAccess"
    }

    foreach ($p in $perms) {
        $results += [PSCustomObject]@{
            SharedMailbox      = $mb.PrimarySmtpAddress
            SharedDisplayName  = $mb.DisplayName
            User               = $p.User.ToString()
            AccessRights       = ($p.AccessRights -join ",")
        }
    }
}

$results |
    Export-Csv -Path $outFile -NoTypeInformation -Encoding UTF8

Write-Host "Done." -ForegroundColor Green
