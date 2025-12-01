<#
.SYNOPSIS
    Export Microsoft 365 shared mailboxes with primary SMTP and all email addresses.

.DESCRIPTION
    This script exports all shared mailboxes (RecipientTypeDetails = SharedMailbox)
    from Exchange Online to a CSV file.

.REQUIREMENTS
    - PowerShell 5.1+ (or PowerShell 7+)
    - ExchangeOnlineManagement module
    - Active Exchange Online session (Connect-ExchangeOnline)

.EXAMPLE
    .\Export-M365-Shared-Mailboxes.ps1
    .\Export-M365-Shared-Mailboxes.ps1 -OutputPath "C:\Reports\M365"
#>

param (
    [string]$OutputPath = ".\output"
)

New-Item -ItemType Directory -Path $OutputPath -ErrorAction SilentlyContinue | Out-Null

$timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
$outFile   = Join-Path $OutputPath "M365_Shared-Mailboxes_$timestamp.csv"

Write-Host "Exporting shared mailboxes to: $outFile" -ForegroundColor Cyan

Get-EXOMailbox -RecipientTypeDetails SharedMailbox -ResultSize Unlimited |
    Select-Object DisplayName,
                  PrimarySmtpAddress,
                  EmailAddresses |
    Export-Csv -Path $outFile -NoTypeInformation -Encoding UTF8

Write-Host "Done." -ForegroundColor Green
