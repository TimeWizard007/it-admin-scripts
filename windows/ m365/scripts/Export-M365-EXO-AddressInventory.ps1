<#
.SYNOPSIS
    Export Exchange Online email address inventory in one run (All-In-One).
    (EN) Exports Exchange Online email address inventory (users, shared mailboxes, permissions, recipients).
    (PL) Eksportuje inwentaryzację adresów e-mail w Exchange Online (użytkownicy, skrzynki współdzielone, uprawnienia, odbiorcy) w jednym przebiegu.

.DESCRIPTION
    (EN)
    This script collects a focused EMAIL INVENTORY for Exchange Online:
      - User mailboxes (with all SMTP addresses)
      - Shared mailboxes (with all SMTP addresses)
      - FullAccess permissions to shared mailboxes (who has access to what)
      - All recipients from Get-Recipient (users, groups, shared, contacts, resources)

    It does NOT export all possible Microsoft 365 configuration (e.g. Teams, SharePoint, licenses).
    Output is intended for:
      - address clean-up projects,
      - documentation for the customer,
      - audits of Exchange Online addressing.

    (PL)
    Skrypt zbiera ukierunkowaną INWENTARYZACJĘ ADRESÓW E-MAIL w Exchange Online:
      - skrzynki użytkowników (ze wszystkimi adresami SMTP),
      - skrzynki współdzielone (ze wszystkimi adresami SMTP),
      - uprawnienia FullAccess do skrzynek współdzielonych (kto do czego ma dostęp),
      - wszystkich odbiorców z Get-Recipient (użytkownicy, grupy, skrzynki współdzielone, kontakty, zasoby).

    Skrypt NIE eksportuje pełnej konfiguracji całego Microsoft 365 (np. Teams, SharePoint, licencje).
    Wynik jest przeznaczony do:
      - projektów porządkowania adresacji,
      - dokumentowania adresów e-mail dla klienta,
      - audytów adresacji Exchange Online.

.REQUIREMENTS
    - PowerShell 5.1+ (or PowerShell 7+)
    - ExchangeOnlineManagement module
    - Active Exchange Online session (Connect-ExchangeOnline)

.EXAMPLE
    .\Export-M365-AIO.ps1
    .\Export-M365-AIO.ps1 -OutputPath "C:\Reports\M365"
#>

param (
    [string]$OutputPath = ".\output"
)

# ===== Universal bootstrap =====

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

Write-Host "Starting full M365 email inventory export..." -ForegroundColor Cyan
Write-Host "Output folder: $OutputPath" -ForegroundColor Cyan
Write-Host "Timestamp:     $timestamp" -ForegroundColor Cyan

### 1. Users + emails
$outUsers = Join-Path $OutputPath "M365_Users-Emails_$timestamp.csv"
Write-Host "`n[1/4] Exporting user mailboxes -> $outUsers" -ForegroundColor Yellow

Get-EXOMailbox -RecipientTypeDetails UserMailbox -ResultSize Unlimited |
    Select-Object DisplayName,
                  UserPrincipalName,
                  PrimarySmtpAddress,
                  EmailAddresses |
    Export-Csv -Path $outUsers -NoTypeInformation -Encoding UTF8

### 2. Shared mailboxes
$outShared = Join-Path $OutputPath "M365_Shared-Mailboxes_$timestamp.csv"
Write-Host "`n[2/4] Exporting shared mailboxes -> $outShared" -ForegroundColor Yellow

Get-EXOMailbox -RecipientTypeDetails SharedMailbox -ResultSize Unlimited |
    Select-Object DisplayName,
                  PrimarySmtpAddress,
                  EmailAddresses |
    Export-Csv -Path $outShared -NoTypeInformation -Encoding UTF8

### 3. Shared mailbox permissions
$outPerms = Join-Path $OutputPath "M365_Shared-Mailbox-Permissions_$timestamp.csv"
Write-Host "`n[3/4] Exporting shared mailbox permissions -> $outPerms" -ForegroundColor Yellow

$results = @()

Get-EXOMailbox -RecipientTypeDetails SharedMailbox -ResultSize Unlimited | ForEach-Object {
    $mb = $_
    Write-Host "   Processing shared mailbox: $($mb.PrimarySmtpAddress)" -ForegroundColor DarkYellow

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
    Export-Csv -Path $outPerms -NoTypeInformation -Encoding UTF8

### 4. All recipients
$outRecipients = Join-Path $OutputPath "M365_All-Recipients_$timestamp.csv"
Write-Host "`n[4/4] Exporting all recipients -> $outRecipients" -ForegroundColor Yellow

Get-Recipient -ResultSize Unlimited |
    Select-Object Name,
                  RecipientType,
                  PrimarySmtpAddress,
                  EmailAddresses |
    Export-Csv -Path $outRecipients -NoTypeInformation -Encoding UTF8

Write-Host "`nFull M365 email inventory export completed." -ForegroundColor Green
Write-Host "Files created:" -ForegroundColor Green
Write-Host " - $outUsers"
Write-Host " - $outShared"
Write-Host " - $outPerms"
Write-Host " - $outRecipients"
