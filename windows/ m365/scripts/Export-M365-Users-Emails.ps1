<#
.SYNOPSIS
    Export Microsoft 365 user mailboxes with primary SMTP address and all email addresses.

.DESCRIPTION
    This script exports all user mailboxes (RecipientTypeDetails = UserMailbox)
    from Exchange Online to a CSV file. It includes:
    - DisplayName
    - UserPrincipalName
    - PrimarySmtpAddress
    - EmailAddresses (all SMTP addresses, including aliases and technical addresses)

.REQUIREMENTS
    - PowerShell 5.1+ (or PowerShell 7+)
    - ExchangeOnlineManagement module
    - Active Exchange Online session (Connect-ExchangeOnline)

.EXAMPLE
    .\Export-M365-Users-Emails.ps1
    .\Export-M365-Users-Emails.ps1 -OutputPath "C:\Reports\M365"
#>

param (
    [string]$OutputPath = ".\output"
)

# Ensure output directory exists
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
