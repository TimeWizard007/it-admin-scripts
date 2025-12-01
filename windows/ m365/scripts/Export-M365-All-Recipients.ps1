<#
.SYNOPSIS
    Export all Exchange Online recipients.

.DESCRIPTION
    This script exports all recipients from Exchange Online using Get-Recipient:
    - User mailboxes
    - Shared mailboxes
    - Groups
    - Contacts
    - Resources, etc.

.REQUIREMENTS
    - PowerShell 5.1+ (or PowerShell 7+)
    - ExchangeOnlineManagement module
    - Active Exchange Online session (Connect-ExchangeOnline)

.EXAMPLE
    .\Export-M365-All-Recipients.ps1
    .\Export-M365-All-Recipients.ps1 -OutputPath "C:\Reports\M365"
#>

param (
    [string]$OutputPath = ".\output"
)

New-Item -ItemType Directory -Path $OutputPath -ErrorAction SilentlyContinue | Out-Null

$timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
$outFile   = Join-Path $OutputPath "M365_All-Recipients_$timestamp.csv"

Write-Host "Exporting all recipients to: $outFile" -ForegroundColor Cyan

Get-Recipient -ResultSize Unlimited |
    Select-Object Name,
                  RecipientType,
                  PrimarySmtpAddress,
                  EmailAddresses |
    Export-Csv -Path $outFile -NoTypeInformation -Encoding UTF8

Write-Host "Done." -ForegroundColor Green
