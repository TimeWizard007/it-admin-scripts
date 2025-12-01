<#
.SYNOPSIS
    Export FullAccess permissions to shared mailboxes.

.DESCRIPTION
    This script exports mapping between shared mailboxes and users
    who have explicit FullAccess permissions (non-inherited).

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

$results | Export-Csv -Path $outFile -NoTypeInformation -Encoding UTF8

Write-Host "Done." -ForegroundColor Green
