<#
.SYNOPSIS
    Export Exchange Online user mailboxes (basic email inventory).
    (EN) Exports user mailboxes with primary SMTP address and all SMTP addresses.
    (PL) Eksportuje skrzynki użytkowników w Exchange Online (adres główny + wszystkie adresy SMTP).

.DESCRIPTION
    (EN) This script exports all user mailboxes (RecipientTypeDetails = UserMailbox)
         from Exchange Online to a CSV file. It includes:
           - DisplayName
           - UserPrincipalName
           - PrimarySmtpAddress
           - EmailAddresses (all SMTP addresses, including aliases and technical addresses)

    (PL) Skrypt eksportuje wszystkie skrzynki użytkowników (RecipientTypeDetails = UserMailbox)
         z Exchange Online do pliku CSV, zawierając:
           - DisplayName
           - UserPrincipalName
           - PrimarySmtpAddress
           - EmailAddresses (wszystkie adresy SMTP, w tym aliasy oraz adresy techniczne).

.REQUIREMENTS
    - PowerShell 5.1+ (or PowerShell 7+)
    - ExchangeOnlineManagement module
    - Active Exchange Online session (Connect-ExchangeOnline)

.NOTES
    (EN) Recommended to run as Administrator (to allow module installation for all users).
    (PL) Zalecane uruchamianie jako Administrator (aby móc instalować moduł dla wszystkich użytkowników).

.EXAMPLE
    .\Export-M365-Users-Emails.ps1
    .\Export-M365-Users-Emails.ps1 -OutputPath "C:\Reports\M365"
#>

param (
    [string]$OutputPath = ".\output"
)

# =========================
# Universal bootstrap block
# =========================

# (EN) Allow script execution in current PowerShell session only.
# (PL) Umożliwiamy wykonywanie skryptów tylko w bieżącej sesji PowerShell.
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force

# (EN) Check if script is running as Administrator.
# (PL) Sprawdzamy, czy skrypt uruchomiony jest jako Administrator.
$IsAdmin = ([Security.Principal.WindowsPrincipal] `
    [Security.Principal.WindowsIdentity]::GetCurrent() `
).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

if (-not $IsAdmin) {
    Write-Warning "This script is not running as Administrator."
    Write-Warning "Some actions (e.g. installing modules for all users) may fail."
    $choice = Read-Host "Continue anyway? (Y/N)"
    if ($choice -ne "Y") { exit }
}

# (EN) Check if ExchangeOnlineManagement module is available; if not, offer to install it.
# (PL) Sprawdzamy, czy moduł ExchangeOnlineManagement jest dostępny; jeśli nie, proponuje instalację.
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

# (EN) Import the module.
# (PL) Import modułu.
Import-Module ExchangeOnlineManagement -ErrorAction Stop

# (EN) Check if there is an active Exchange Online session; if not, offer to connect.
# (PL) Sprawdzamy, czy istnieje aktywna sesja Exchange Online; jeśli nie, proponujemy połączenie.
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

# =========================
# Main script logic
# =========================

# (EN) Ensure output directory exists.
# (PL) Sprawdzamy czy katalog wyjściowy istnieje.
New-Item -ItemType Directory -Path $OutputPath -ErrorAction SilentlyContinue | Out-Null

$timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
$outFile   = Join-Path $OutputPath "M365_Users-Emails_$timestamp.csv"

Write-Host "Exporting user mailboxes to: $outFile" -ForegroundColor Cyan

Get-EXOMailbox -RecipientTypeDetails UserMailbox -ResultSize Unlimited |
    Select-Object DisplayName,
                  UserPrincipalName,
                  PrimarySmtpAddress,
                  EmailAddresses |
    Export-Csv -Path $outFile -NoTypeInformation -Encoding UTF8

Write-Host "Done." -ForegroundColor Green
