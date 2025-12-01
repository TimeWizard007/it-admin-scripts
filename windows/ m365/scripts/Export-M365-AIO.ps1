<#
.SYNOPSIS
    Generate full Microsoft 365 email inventory in one run (All-In-One).

.DESCRIPTION
    This script runs a complete set of exports for Exchange Online:
    - User mailboxes with all email addresses
    - Shared mailboxes
    - FullAccess permissions to shared mailboxes
    - All recipients (Get-Recipient)

    All files are stored in a single output folder with timestamped names.
    It is intended for audits, clean-up projects and documentation
    (e.g. providing full address inventory to the customer).

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
    Write-Host "   Processing: $($mb.PrimarySmtpAddress)" -ForegroundColor DarkYellow

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

$results | Export-Csv -Path $outPerms -NoTypeInformation -Encoding UTF8

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
