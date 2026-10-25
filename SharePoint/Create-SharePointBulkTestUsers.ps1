#Requires -Version 7.4
# Requires PowerShell version 7.4 or later so the script can use the latest runtime features and compatibility.
<#
# Starts a PowerShell comment-based help block so the script can include documentation metadata.
# This help block begins the script's synopsis section.
.SYNOPSIS
# Describes the script purpose: create a SharePoint Online sample list and add fake users in batches.
Creates a SharePoint Online sample Users list and inserts fictional users in batches.
# Begins a sample execution example for users who want to run the script.
.EXAMPLE
# Shows a command example with the SharePoint site URL, app client ID, and total number of records to insert.
./Create-SharePointBulkTestUsers.ps1 -SiteUrl 'https://contoso.sharepoint.com/sites/AutomationLab' -ClientId '00000000-0000-0000-0000-000000000000' -RecordCount 12000
# Starts the notes section, which documents requirements and scope for the script.
.NOTES
# States that the script requires PnP.PowerShell and a delegated SharePoint permission configuration.
Requires PnP.PowerShell and an Entra ID app registration with appropriate delegated SharePoint permissions.
# Clarifies that the script does not create real Microsoft 365 accounts and only adds fictional list items.
This script does not create real Microsoft 365 user accounts. It creates fictional list items only.
# Ends the comment-based help block so PowerShell can continue with executable code.

#>

# Follow the link step by step in order to register application on azure portal and get a client ID and set permissions as a delegate sharepoint allsites.fullcontrol
# https://pnp.github.io/powershell/articles/registerapplication.html

# OR by registring 
# Register-PnPEntraIDAppForInteractiveLogin `
#     -ApplicationName "SharePoint Pagination Testing" `
#     -Tenant "contoso.onmicrosoft.com" `
#     -SharePointDelegatePermissions "AllSites.FullControl"

# Enables advanced PowerShell parameter binding features like validation and support for common cmdlet behaviors.
[CmdletBinding()]
# Starts the script parameter block where runtime settings are defined.
param(
# Requires the SharePoint site URL parameter and marks it as mandatory.
    [Parameter(Mandatory = $true)]
# Validates that the URL is in the expected SharePoint Online format, such as /sites or /teams.
    [ValidatePattern('^https://[^/]+\.sharepoint\.com/(sites|teams)/[^/]+/?$')]
# Stores the target SharePoint site URL as a string value.
    [string] $SiteUrl,

# Requires the Entra app client ID and marks it as mandatory.
    [Parameter(Mandatory = $true)]
# Ensures the client ID is not empty before the script continues.
    [ValidateNotNullOrEmpty()]
# Stores the SharePoint app registration client ID used for authentication.
    [string] $ClientId,

# Restricts the total number of items to a range between 5001 and 100000.
    [ValidateRange(5001, 100000)]
# Sets the default record count to 12,000 if the caller does not provide one.
    [int] $RecordCount = 12000,

# Restricts the batch size to a range between 1 and 500 records.
    [ValidateRange(1, 500)]
# Sets the default number of records inserted per batch to 100.
    [int] $BatchSize = 100,

# Ensures the list name is not empty when provided by the user.
    [ValidateNotNullOrEmpty()]
# Defines the default SharePoint list name used for the generated demo list.
    [string] $ListName = 'Demo Users - Pagination Test'
)

# Stops script execution immediately on non-terminating errors so failure conditions are visible and explicit.
$ErrorActionPreference = 'Stop'

# Checks whether the required PnP.PowerShell module is installed on the machine.
if (-not (Get-Module -ListAvailable -Name PnP.PowerShell)) {
# Throws a clear message if the module is missing so the user knows how to install it.
    throw 'PnP.PowerShell is not installed. Run: Install-Module PnP.PowerShell -Scope CurrentUser'
}
# Imports the PnP.PowerShell module so SharePoint cmdlets become available in this script.
Import-Module PnP.PowerShell -ErrorAction Stop

# Displays a status message before connecting to the target SharePoint site.
Write-Host "Connecting to $SiteUrl ..." -ForegroundColor Cyan
# Connects to the SharePoint site using the app client ID and interactive sign-in, returning the active connection object.
$connection = Connect-PnPOnline -Url $SiteUrl -ClientId $ClientId -Interactive -ReturnConnection -ErrorAction Stop

# Safety: never overwrite or insert duplicates into an existing list.
# Queries the SharePoint site for a list with the chosen name to prevent accidental collisions.
$existingList = Get-PnPList -Identity $ListName -Connection $connection -ErrorAction SilentlyContinue
# If a list with the same name already exists, the script stops before making any changes.
if ($null -ne $existingList) {
# Raises a descriptive error to preserve the existing list and prevent accidental data changes.
    throw "A list named '$ListName' already exists. Choose another -ListName. No data was changed."
}

# Displays a message before creating the new SharePoint list.
Write-Host "Creating list '$ListName' ..." -ForegroundColor Cyan
# Creates a new generic SharePoint list with the desired name and adds it to quick launch.
New-PnPList -Title $ListName -Template GenericList -OnQuickLaunch -Connection $connection | Out-Null

# Title is the built-in SharePoint column. Other field names are stable internal names.
# Defines the custom metadata fields used to store employee details for each fictional record.
$fields = @(
# Creates the EmployeeId column as a text field containing the unique employee code.
    @{ Name = 'EmployeeId';      Label = 'Employee ID';       Type = 'Text' },
# Creates the FirstName column for the employee's first name.
    @{ Name = 'FirstName';       Label = 'First Name';        Type = 'Text' },
# Creates the LastName column for the employee's last name.
    @{ Name = 'LastName';        Label = 'Last Name';         Type = 'Text' },
# Creates the Email column for the user email address.
    @{ Name = 'Email';           Label = 'Email';             Type = 'Text' },
# Creates the Department column for the employee's organization unit.
    @{ Name = 'Department';      Label = 'Department';        Type = 'Text' },
# Creates the JobTitle column for the employee's role.
    @{ Name = 'JobTitle';        Label = 'Job Title';         Type = 'Text' },
# Creates the OfficeLocation column for the employee's office site.
    @{ Name = 'OfficeLocation';  Label = 'Office Location';   Type = 'Text' },
# Creates the HireDate column using a DateTime field to store employment dates.
    @{ Name = 'HireDate';        Label = 'Hire Date';         Type = 'DateTime' },
# Creates the IsActive column as a Boolean flag to show whether the record is active.
    @{ Name = 'IsActive';        Label = 'Is Active';         Type = 'Boolean' },
# Creates the ProcessingStatus column to support filtering and pagination testing.
    @{ Name = 'ProcessingStatus';Label = 'Processing Status'; Type = 'Text' }
)
# Loops through each field definition and adds the field to the SharePoint list.
foreach ($field in $fields) {
# Adds the new field to the list, setting its internal name, display name, and data type.
    Add-PnPField -List $ListName -InternalName $field.Name -DisplayName $field.Label `
# Applies the field type and adds it to the default list view.
        -Type $field.Type -AddToDefaultView -Connection $connection | Out-Null
}

# Index filter columns BEFORE crossing the SharePoint list view threshold.
# Iterates through the key filter columns that should be indexed for better query performance.
foreach ($internalName in @('EmployeeId', 'ProcessingStatus', 'Department')) {
# Marks the selected field as indexed to improve filter and query efficiency.
    Set-PnPField -List $ListName -Identity $internalName -Values @{ Indexed = $true } `
# Uses the active SharePoint connection to apply the index change.
        -Connection $connection -ErrorAction Stop
# Writes a console message confirming which field was indexed.
    Write-Host "Indexed: $internalName"
}

# Defines the list of first names used to build fictional employee identities.
$firstNames = @('Alex','Jamie','Taylor','Jordan','Morgan','Casey','Riley','Avery','Quinn','Cameron')
# Defines the list of last names used to create realistic full names.
$lastNames  = @('Smith','Brown','Johnson','Lee','Martin','Wilson','Clark','Hall','Young','Walker')
# Defines departments used to simulate different organizational groups.
$departments = @('IT','HR','Finance','Sales','Operations','Marketing')
# Defines job titles used to simulate employee roles.
$jobs = @('Engineer','Analyst','Specialist','Coordinator','Administrator','Manager')
# Defines office locations used to simulate a distributed workforce.
$offices = @('Beirut','Dubai','Doha','London')

# The first 6,000 records are Completed; the remaining records are Pending.
# This creates a useful Power Automate filter/pagination test on a >5,000-item list.
# Creates a batch object for submitting multiple list items efficiently.
$batch = New-PnPBatch -Connection $connection
# Tracks how many items are queued in the current batch before submission.
$queued = 0
# Tracks how many records have already been successfully submitted.
$submitted = 0
# Stores the time when the bulk insert begins for elapsed-time reporting.
$started = Get-Date

# Begins the loop that generates records from 1 up to the requested record count.
for ($i = 1; $i -le $RecordCount; $i++) {
# Selects a first name based on the current record number in a repeating pattern.
    $first = $firstNames[($i - 1) % $firstNames.Count]
# Selects a last name using a calculated offset so names repeat in groups.
    $last  = $lastNames[[int][math]::Floor(($i - 1) / $firstNames.Count) % $lastNames.Count]
# Generates a unique employee ID in the format EMP000001, EMP000002, and so on.
    $employeeId = 'EMP{0:D6}' -f $i
# Marks the first 6000 records as Completed and the rest as Pending.
    $status = if ($i -le 6000) { 'Completed' } else { 'Pending' }

# Builds a hashtable of field values for a single fictional employee record.
    $values = @{
# Sets the SharePoint list item title to "First Last (EMP000001)".
        Title            = "$first $last ($employeeId)"
# Stores the unique employee ID in the EmployeeId column.
        EmployeeId       = $employeeId
# Stores the employee's first name in the FirstName column.
        FirstName        = $first
# Stores the employee's last name in the LastName column.
        LastName         = $last
# Creates a unique email address using the current loop index.
        Email            = ('user{0:D6}@example.test' -f $i)
# Selects a department from the predefined list in rotation.
        Department       = $departments[($i - 1) % $departments.Count]
# Selects a job title from the predefined list in rotation.
        JobTitle         = $jobs[($i - 1) % $jobs.Count]
# Selects an office location from the predefined list in rotation.
        OfficeLocation   = $offices[($i - 1) % $offices.Count]
# Sets the hire date based on a base date plus a calculated number of days.
        HireDate         = (Get-Date '2020-01-01').AddDays($i % 1800)
# Marks every tenth record as inactive, leaving the rest active.
        IsActive         = ($i % 10 -ne 0)
# Stores the record status as Completed or Pending for testing scenarios.
        ProcessingStatus = $status
    }

# Adds the current employee record to the batch queue without committing immediately.
    Add-PnPListItem -List $ListName -Values $values -Batch $batch -Connection $connection | Out-Null
# Increments the count of records queued in the current batch.
    $queued++

# Checks whether the current batch is full or whether the final record has been reached.
    if ($queued -ge $BatchSize -or $i -eq $RecordCount) {
# Deliberately fail instead of blindly replaying a potentially partially committed batch.
# Submits the current batch to SharePoint and stops the script on failure.
        Invoke-PnPBatch -Batch $batch -Connection $connection -ErrorAction Stop
# Adds the number of records in the just-submitted batch to the total submitted count.
        $submitted += $queued
# Writes progress with the current number submitted and percentage completion.
        Write-Progress -Activity 'Creating fictional SharePoint users' `
# Shows the current submitted count and total requested count in the status text.
            -Status "$submitted of $RecordCount submitted" `
# Calculates the completion percentage for the progress bar.
            -PercentComplete (100 * $submitted / $RecordCount)
# Prints a simple status line showing the submitted count.
        Write-Host "Submitted $submitted / $RecordCount items"
# Creates a fresh batch object for the next group of items.
        $batch = New-PnPBatch -Connection $connection
# Resets the queued counter to zero for the next batch.
        $queued = 0
    }
}
# Marks the progress bar as complete after the loop finishes.
Write-Progress -Activity 'Creating fictional SharePoint users' -Completed

# Retrieves the list metadata, including item count and root folder information after insertion.
$list = Get-PnPList -Identity $ListName -Includes ItemCount,RootFolder -Connection $connection
# Prints the final item count reported by SharePoint in green to confirm success.
Write-Host "Done. SharePoint reports $($list.ItemCount) items." -ForegroundColor Green
# Builds and prints the full SharePoint list URL for direct access.
Write-Host "List URL: $(([uri]$SiteUrl).GetLeftPart([System.UriPartial]::Authority))$($list.RootFolder.ServerRelativeUrl)"
# Shows how long the script took to create all records in minutes.
Write-Host "Elapsed: $([math]::Round(((Get-Date) - $started).TotalMinutes, 2)) minutes"
# Prints a sample Power Automate filter expression for Pending items.
Write-Host "Power Automate filter: ProcessingStatus eq 'Pending'"
# Prints a useful tip for handling large SharePoint lists in Power Automate pagination flows.
Write-Host 'Tip: enable Get items pagination, and use Top Count 500 for your update-in-batches flow.'
