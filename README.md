# PowerShell Commands

This repository contains PowerShell automation scripts for Microsoft 365 and SharePoint testing scenarios.

## Repository contents

- `SharePoint/Create-SharePointBulkTestUsers.ps1` - Creates a SharePoint Online list and fills it with large volumes of fictional employee records for pagination and performance testing.

## SharePoint bulk test user generator

The script in `SharePoint/Create-SharePointBulkTestUsers.ps1` is designed for SharePoint Online scenarios where you need to generate a large list of demo records quickly.

### What it does

- Connects to a SharePoint site using PnP.PowerShell
- Creates a new generic list if one does not already exist
- Adds custom fields such as Employee ID, first name, last name, department, job title, office location, hire date, active flag, and processing status
- Creates indexes on commonly filtered fields before the list grows beyond SharePoint thresholds
- Inserts records in batches to improve reliability and performance
- Generates realistic-looking test data for pagination, filtering, and Power Automate flow validation

### Default behavior

By default, the script creates a list called `Demo Users - Pagination Test` and inserts 12,000 records in batches of 100.

The generated records include:

- Unique employee IDs like `EMP000001`
- Repeating first/last name combinations
- Department and job title values
- Office location values
- Active/inactive flags
- `Completed` and `Pending` processing statuses to emulate real-world filter scenarios

## Prerequisites

Before running the script, make sure you have:

- PowerShell 7.4 or later
- The `PnP.PowerShell` module installed
- A SharePoint Online site URL
- An Entra ID app registration with delegated SharePoint permissions

Install the required module if needed:

```powershell
Install-Module PnP.PowerShell -Scope CurrentUser
```

## Usage

```powershell
./SharePoint/Create-SharePointBulkTestUsers.ps1 `
  -SiteUrl 'https://contoso.sharepoint.com/sites/AutomationLab' `
  -ClientId '00000000-0000-0000-0000-000000000000' `
  -RecordCount 12000
```

Optional parameters:

```powershell
./SharePoint/Create-SharePointBulkTestUsers.ps1 `
  -SiteUrl 'https://contoso.sharepoint.com/sites/AutomationLab' `
  -ClientId '00000000-0000-0000-0000-000000000000' `
  -RecordCount 25000 `
  -BatchSize 250 `
  -ListName 'Demo Users - Custom Test Data'
```

## Notes

- The script does not create real Microsoft 365 user accounts.
- It only creates fictional list items for testing.
- It refuses to overwrite an existing list with the same name to avoid accidental data loss.
- This is especially useful for testing SharePoint list view thresholds, filter performance, and pagination in Power Automate.

## Example app registration flow

This script works with the SharePoint delegated permission model described in the PnP PowerShell docs. A typical setup is:

```powershell
Register-PnPEntraIDAppForInteractiveLogin `
  -ApplicationName "SharePoint Pagination Testing" `
  -Tenant "contoso.onmicrosoft.com" `
  -SharePointDelegatePermissions "AllSites.FullControl"
```

## Related resources

- [PnP PowerShell documentation](https://pnp.github.io/powershell/)
- [SharePoint Online app registration guidance](https://pnp.github.io/powershell/articles/registerapplication.html)
