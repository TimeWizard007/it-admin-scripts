# 🇵🇱 / 🇬🇧 README – Polish (PL) first, English (EN) second
---

# 🇵🇱 Exchange Online – Zestaw skryptów do inwentaryzacji adresów e-mail

## 1. Opis projektu
Ten zestaw skryptów PowerShell umożliwia wykonanie **szczegółowej inwentaryzacji adresów e-mail w Exchange Online (Microsoft 365)**.  
Skrypty są przeznaczone do:

- audytów adresacji,
- porządkowania środowiska e-mail,
- raportowania aliasów i skrzynek współdzielonych,
- dokumentowania uprawnień FullAccess.

Każdy skrypt posiada:
- sprawdzanie modułu ExchangeOnlineManagement (z auto-instalacją),
- obsługę uruchomienia bez uprawnień administratora,
- ustawianie ExecutionPolicy tylko dla bieżącej sesji,
- komentarze PL+EN,
- komunikaty wyłącznie po angielsku (standard administracyjny).

---

## 2. Struktura repozytorium
scripts/
│ ├── Export-M365-Users-Emails.ps1
│ ├── Export-M365-Shared-Mailboxes.ps1
│ ├── Export-M365-Shared-Permissions.ps1
│ ├── Export-M365-All-Recipients.ps1
│ └── Export-M365-AIO.ps1

---

## 3. Opis skryptów

### 3.1 Export-M365-Users-Emails.ps1
Eksportuje skrzynki użytkowników:
- Primary SMTP,
- wszystkie adresy SMTP (aliasy, techniczne).

### 3.2 Export-M365-Shared-Mailboxes.ps1
Eksportuje skrzynki współdzielone:
- Primary SMTP,
- wszystkie przypisane adresy SMTP.

### 3.3 Export-M365-Shared-Permissions.ps1
Eksportuje uprawnienia FullAccess:
- kto ma dostęp do której shared mailbox.

### 3.4 Export-M365-All-Recipients.ps1
Eksportuje wszystkie obiekty zwracane przez `Get-Recipient`:
- użytkownicy,
- grupy,
- shared mailboxy,
- zasoby,
- kontakty.

### 3.5 Export-M365-AIO.ps1
Skrypt „All-In-One”:
- uruchamia wszystkie cztery eksporty,
- zapisuje cztery pliki CSV,
- idealny do audytów i raportów dla klientów.

---

## 4. Wymagania
- Windows PowerShell 5.1+ lub PowerShell 7+
- Moduł ExchangeOnlineManagement:

```powershell
Install-Module ExchangeOnlineManagement
Connect-ExchangeOnline
```
————————————————————————————————————————

🇬🇧 Exchange Online – Email Address Inventory Scripts

## 1. Project Description
This PowerShell script set enables generating a detailed email address inventory for Exchange Online (Microsoft 365).
The scripts are intended for:

- address audits,
- cleanup of the email environment,
- reporting aliases and shared mailboxes,
- documenting FullAccess permissions.

Each script includes:
- validation of the ExchangeOnlineManagement module (with auto-install),
- execution without administrator rights support,
- setting ExecutionPolicy for the current session only,
- dual PL+EN inline comments,
- English-only console output (administrative standard).

---
## 2. Repository Structure
scripts/
│   ├── Export-M365-Users-Emails.ps1
│   ├── Export-M365-Shared-Mailboxes.ps1
│   ├── Export-M365-Shared-Permissions.ps1
│   ├── Export-M365-All-Recipients.ps1
│   └── Export-M365-AIO.ps1

## 3. Script Descriptions

### 3.1 Export-M365-Users-Emails.ps1
Exports user mailboxes:
- Primary SMTP,
- all SMTP addresses (aliases, technical).

### 3.2 Export-M365-Shared-Mailboxes.ps1
Exports shared mailboxes:
- Primary SMTP,
- all assigned SMTP addresses.

### 3.3 Export-M365-Shared-Permissions.ps1
Exports FullAccess permissions:
- who has access to which shared mailbox.

### 3.4 Export-M365-All-Recipients.ps1
Exports all objects returned by Get-Recipient:
- users,
- groups,
- shared mailboxes,
- resources,
- contacts.

### 3.5 Export-M365-AIO.ps1
“All-In-One” script:
- runs all four exports,
- outputs four CSV files,
- ideal for audits and customer reporting.

---

## 4.Requirements
- Windows PowerShell 5.1+ lub PowerShell 7+
- Moduł ExchangeOnlineManagement:

```powershell
Install-Module ExchangeOnlineManagement
Connect-ExchangeOnline
```
————————————————————————————————————————
