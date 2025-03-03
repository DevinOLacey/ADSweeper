<#
    Script Name: AD_SweeperV2.ps1
    Purpose    : Clean and update Active Directory (AD) attributes based on SQL data
    Author     : Devin Lacey
    Date       : 01/17/2025

    Description:
    This script synchronizes and maintains Active Directory user attributes with data from a SQL database.
    It performs the following main functions:
    - Updates user attributes (Department, Title, Office, Description)
    - Manages employee numbers
    - Handles terminated employees
    - Moves users to appropriate Organizational Units (OUs)
    - Generates detailed Excel reports of all changes and discrepancies

    Requirements:
    - Active Directory Module
    - ImportExcel Module
    - SQL Server access
    - Appropriate AD permissions
#>

# Initialize script and verify running status
Write-Host "AD_SweeperV2 is running..." -ForegroundColor Green

# Import Required Modules             
Import-Module ActiveDirectory

# Check and install ImportExcel module if not present
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Write-Host "ImportExcel module not found. Installing it now..." -ForegroundColor Yellow
    Install-Module ImportExcel -Scope CurrentUser -Force -SkipPublisherCheck
}

# Define departments to exclude from processing
# These departments are typically managed separately or require special handling
$ExcludedDepartments = @(
    "Information Technology",
    "IT",
    "IT Department",
    "Tribal Council",
    "Gaming Commission",
    "Seceon",
    "Tribal Government",
    "SVC",
    "HealthMailbox"
    "Vendor",
    "Deputy",
    "Shared",
    "Insights",
    "Test",
    "UKG"
    )

# Configure output file paths and ensure they exist
$BaseExportPath = "\\jamfs01\Shared\Information_Technology\AD_Sweeper\Output"

# Create export directory if it doesn't exist
if (-not (Test-Path -Path $BaseExportPath)) {
    try {
        New-Item -ItemType Directory -Path $BaseExportPath -Force | Out-Null
        Write-Host "Created base export directory at '$BaseExportPath'." -ForegroundColor Green
    }
    catch {
        Write-Host "Failed to create base export directory: $($_.Exception.Message)" -ForegroundColor Red
    }
}

# Generate timestamp for unique file naming
$RunTimestamp = Get-Date -Format 'yyyyMMddHHmmss'

# Define consolidated output file path
$ConsolidatedOutputFile = Join-Path $BaseExportPath "AD_Sweeper-$RunTimestamp.xlsx"

# Define AD attributes to retrieve for processing
$AdFields = @("Office", "Department", "Description", "Title", "EmployeeNumber", "Name", "SamAccountName", "DistinguishedName")

# Initialize collections for tracking different user categories
$AllNoSqlData             = @() # Users without matching SQL data
$TerminatedActiveAdUsers  = @() # Users marked as terminated but still active in AD
$AllUpdatedUsers          = @() # Users who had attributes updated
$Admins                   = @() # Users requiring elevated admin privileges
$ProcessedUsers           = @{} # Track processed users to prevent duplicates
$DisabledUsers            = @() # Users who were disabled during this run

try {
    # Establish SQL Connection
    $conn = Connect-SqlServer
    if ($conn) {
        # Filter AD Users - Exclude specified departments and specific users
        $AdUsers = Get-ADUser -Filter * -Properties $AdFields | Where-Object {
            $ExcludedDepartments -notcontains $_.Department -and
            $ExcludedDepartments -notcontains $_.Office -and
            $_.Name -ne "Josh Ford"
        }

        Write-Host "Filtered AD user count: $($AdUsers.Count)"

        if (!$AdUsers) { 
            Write-Warning "No AD users found outside the excluded departments."
        }

        # Main processing loop - Iterate through each AD user
        foreach ($AdUser in $AdUsers) {
            $FullName       = $AdUser.Name
            $EmployeeNumber = $AdUser.EmployeeNumber
            $Changes        = @{ }  # Track changes needed for this user
            $SqlResult      = $null

            # Employee Number Processing
            # Clean up existing employee numbers or attempt to find and add missing ones
            if ($EmployeeNumber) {
                # Clean up existing employee number by removing whitespace
                $CleanEmployeeNumber = $EmployeeNumber.Trim()
                if ($CleanEmployeeNumber -ne $EmployeeNumber) {
                    $Changes["EmployeeNumber"] = $CleanEmployeeNumber
                    Set-ADUser -Identity $AdUser.DistinguishedName -EmployeeNumber $CleanEmployeeNumber
                    $AdUser.EmployeeNumber = $CleanEmployeeNumber
                }
                if ($CleanEmployeeNumber) {
                    # Query SQL database for user information using employee number
                    $Query = @"
SELECT Department, JobTitle, EmployeeNumber, EmployeeStatus
FROM UT_HCJ_IDWORKS 
WHERE EmployeeNumber = '$CleanEmployeeNumber'
"@
                    $SqlResult = Execute-SqlQuery -Connection $conn -SqlQuery $Query
                }
            }
            else {
                # If no employee number exists, attempt to find user by name
                $NameParts = $FullName -split ', ' -or $FullName -split ' '
                if ($NameParts.Count -ge 2) {
                    $FirstName = $NameParts[0]
                    $LastName  = $NameParts[-1]

                    # Query SQL database using first and last name
                    $Query = @"
SELECT Department, JobTitle, EmployeeNumber, EmployeeStatus
FROM UT_HCJ_IDWORKS 
WHERE FirstName = '$FirstName' AND LastName = '$LastName'
"@
                    $SqlResult = Execute-SqlQuery -Connection $conn -SqlQuery $Query
                }
            }

            # Process SQL results and update AD attributes
            if ($SqlResult) {
                foreach ($Row in $SqlResult) {
                    # Add missing employee number if found in SQL
                    if ([string]::IsNullOrEmpty($AdUser.EmployeeNumber) -and $Row.EmployeeNumber -and $Row.EmployeeStatus -eq 'Active') {
                        $Changes["EmployeeNumber"] = $Row.EmployeeNumber
                        Set-ADUser -Identity $AdUser.DistinguishedName -EmployeeNumber $Row.EmployeeNumber
                        $AdUser.EmployeeNumber = $Row.EmployeeNumber
                        Write-Host "Added EmployeeNumber '$($Row.EmployeeNumber)' to '$FullName'." -ForegroundColor DarkGray
                    }

                    # Department and Title Processing
                    # Convert SQL department names to standardized format
                    if ($Row.JobTitle -like "*Compliance*") {
                        $SqlDepartment = "Compliance"
                    } elseif ($Row.Department -like "*EVS*") {
                        $SqlDepartment = "EVS"
                    } else {
                        $SqlDepartment = ConvertDeptName -department $Row.Department
                    }
                    $SqlTitle = MapJobTitles -adTitle $Row.JobTitle
                    
                    # Compare and update Department and Office if different
                    if ($SqlDepartment -and (($null -eq $AdUser.Department) -or ($SqlDepartment -cne $AdUser.Department))) {
                        $Changes["Department"] = $SqlDepartment
                    }
                    if ($SqlDepartment -and (($null -eq $AdUser.Office) -or ($SqlDepartment -cne $AdUser.Office))) {
                        $Changes["physicalDeliveryOfficeName"] = $SqlDepartment
                    }

                    # Compare and update Title and Description if different
                    if ($SqlTitle -and (($null -eq $AdUser.Title) -or ($SqlTitle -cne $AdUser.Title))) {
                        $Changes["Title"] = $SqlTitle
                    }
                    if ($SqlTitle -and (($null -eq $AdUser.Description) -or ($SqlTitle -cne $AdUser.Description))) {
                        $Changes["Description"] = $SqlTitle
                    }

                    # Handle terminated or disabled users
                    if ($Row.EmployeeStatus -eq 'Terminated' -or -not $AdUser.Enabled) {
                        $CorrectOU = "OU=Disabled,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
                        $ChangesMade = $false

                        # Remove terminated/disabled users from all groups
                        $Groups = Get-ADUser -Identity $AdUser.DistinguishedName -Properties MemberOf | Select-Object -ExpandProperty MemberOf
                        if ($Groups.Count -gt 0) {
                            foreach ($Group in $Groups) {
                                Remove-ADGroupMember -Identity $Group -Members $AdUser -Confirm:$false
                            }
                            $ChangesMade = $true
                        }

                        # Move terminated/disabled users to correct OU if needed
                        if ($AdUser.DistinguishedName -notmatch $CorrectOU) {
                            UpdateUserOU -AdUser $AdUser -correctOU $CorrectOU
                            $ChangesMade = $true
                        }

                        # Track disabled user changes for reporting
                        if ($ChangesMade) {
                            $DisabledUsers += [PSCustomObject]@{
                                Name              = $AdUser.Name
                                SamAccountName    = $AdUser.SamAccountName
                                Department        = $AdUser.Department
                                Title            = $AdUser.Title
                                EmployeeNumber    = $AdUser.EmployeeNumber
                                DistinguishedName = $AdUser.DistinguishedName
                                Groups           = $Groups -join ", "
                            }
                        }
                    } else {
                        # Determine correct OU for active users based on job title and department
                        $CorrectOU = MapOU -jobTitle $AdUser.Title -department $AdUser.Department
                    }

                    # Process OU changes for active users
                    if (-not [string]::IsNullOrEmpty($CorrectOU)) {
                        # Special handling for Food and Beverage supervisors/managers
                        if ($AdUser.DistinguishedName -match "OU=supervisors,OU=Food and Beverage,OU=users,OU=Jamul,DC=jamulcasinosd,DC=com") {
                            # Keep users already in supervisors OU
                        } elseif ($AdUser.Title -match "(?i)supervisor|manager" -and $AdUser.DistinguishedName -match "OU=Food and Beverage,OU=users,OU=Jamul,DC=jamulcasinosd,DC=com") {
                            # Move matching titles to supervisors OU
                            $CorrectOU = "OU=supervisors,OU=Food and Beverage,OU=users,OU=Jamul,DC=jamulcasinosd,DC=com"
                            UpdateUserOU -AdUser $AdUser -correctOU $CorrectOU
                            $Changes["OU"] = $CorrectOU
                            Write-Host "Moving $($AdUser.Name) to supervisors OU due to title." -ForegroundColor Cyan
                        } else {
                            # Standard OU update for other users
                            UpdateUserOU -AdUser $AdUser -correctOU $CorrectOU
                            $Changes["OU"] = $CorrectOU
                            Write-Host "Moving $($AdUser.Name) to $CorrectOU" -ForegroundColor Cyan
                        }
                    } else {
                        Write-Host "Failed to determine CorrectOU for $($AdUser.Name). Skipping OU update." -ForegroundColor Red
                    }

                    # Track users with attribute changes
                    if ($Changes.Count -gt 0) {
                        $AllUpdatedUsers += [PSCustomObject]@{
                            Name           = $FullName
                            Office         = $AdUser.Office
                            Department     = $AdUser.Department
                            SQLDepartment  = $SqlDepartment
                            ' '            = ''
                            Description    = $AdUser.Description
                            Title          = $AdUser.Title
                            SQLTitle       = $SqlTitle
                            EmployeeNumber = if ($Changes["EmployeeNumber"]) { $Changes["EmployeeNumber"] } else { $AdUser.EmployeeNumber }
                            OU             = if ($Changes["OU"]) { $Changes["OU"] } else { $AdUser.DistinguishedName }
                        }
                    }

                    # Apply accumulated changes to AD
                    if ($Changes.Count -gt 0) {
                        foreach ($key in $Changes.Keys) {
                            # Prevent duplicate processing using unique key
                            $uniqueKey = "$($AdUser.SamAccountName)-$key"

                            if (-not $ProcessedUsers.ContainsKey($uniqueKey)) {
                                try {
                                    # Apply specific attribute changes
                                    switch ($key) {
                                        "Department" {
                                            Set-ADUser -Identity $AdUser.DistinguishedName -Department $Changes[$key]
                                        }
                                        "Title" {
                                            Set-ADUser -Identity $AdUser.DistinguishedName -Title $Changes[$key]
                                        }
                                        "Description" {
                                            Set-ADUser -Identity $AdUser.DistinguishedName -Description $Changes[$key]
                                        }
                                        "physicalDeliveryOfficeName" {
                                            Set-ADUser -Identity $AdUser.DistinguishedName -Replace @{ physicalDeliveryOfficeName = $Changes[$key] }
                                        }
                                        "EmployeeNumber" {
                                            Set-ADUser -Identity $AdUser.DistinguishedName -EmployeeNumber $Changes[$key]
                                        }
                                    }
                                    # Mark change as processed
                                    $ProcessedUsers[$uniqueKey] = $true
                                }
                                catch {
                                    Write-Host "Insufficient access rights for user: $($AdUser.Name)" -ForegroundColor Red
                                    # Track users requiring admin privileges
                                    if (-not $Admins | Where-Object { $_.SamAccountName -eq $AdUser.SamAccountName }) {
                                        $Admins += [PSCustomObject]@{
                                            Name              = $AdUser.Name
                                            SamAccountName    = $AdUser.SamAccountName
                                            Department        = $AdUser.Department
                                            Title            = $AdUser.Title
                                            EmployeeNumber    = $AdUser.EmployeeNumber
                                            DistinguishedName = $AdUser.DistinguishedName
                                        }
                                    }
                                }
                            }
                        }
                    }

                    # Track terminated users who are still active in AD
                    if ($AdUser.Enabled -and $Row.EmployeeStatus -eq 'Terminated') {
                        $TerminatedActiveAdUsers += [PSCustomObject]@{
                            Name              = $AdUser.Name
                            SamAccountName    = $AdUser.SamAccountName
                            Department        = $AdUser.Department
                            Title            = $AdUser.Title
                            EmployeeNumber    = $AdUser.EmployeeNumber
                            EmployeeStatus    = $Row.EmployeeStatus
                            DistinguishedName = $AdUser.DistinguishedName
                        }
                    }

                    # Track users without matching SQL data
                    if ($SqlResult -eq $null -and $AdUser.Enabled) {
                        $AllNoSqlData += [PSCustomObject]@{
                            Name           = $FullName
                            Office         = $AdUser.Office
                            Department     = $AdUser.Department
                            SQLDepartment  = "N/A"
                            ' '            = ''
                            Description    = $AdUser.Description
                            Title          = $AdUser.Title
                            SQLTitle       = "N/A"
                            EmployeeNumber = $AdUser.EmployeeNumber
                        }
                    }
                }
            }
            else {
                # Track users without matching SQL data
                if ($SqlResult -eq $null -and $AdUser.Enabled) {
                    $AllNoSqlData += [PSCustomObject]@{
                        Name           = $FullName
                        Office         = $AdUser.Office
                        Department     = $AdUser.Department
                        SQLDepartment  = "N/A"
                        ' '            = ''
                        Description    = $AdUser.Description
                        Title          = $AdUser.Title
                        SQLTitle       = "N/A"
                        EmployeeNumber = $AdUser.EmployeeNumber
                    }
                }
            }
        }

        # Report generation section
        # Report on terminated users still active in AD
        if ($TerminatedActiveAdUsers.Count -gt 0) {
            Write-Host "Accumulated $($TerminatedActiveAdUsers.Count) terminated active AD users." -ForegroundColor Magenta
        }
        else {
            Write-Host "No terminated active AD users found." -ForegroundColor Yellow
        }
        
        # Report on users without SQL data matches
        if ($AllNoSqlData.Count -gt 0) {
            Write-Host "Accumulated $($AllNoSqlData.Count) unmatched users." -ForegroundColor Cyan
        }
        else {
            Write-Host "No unmatched users found." -ForegroundColor Yellow
        }

        # Export reports to Excel

        # Export unmatched users report
        if ($AllNoSqlData.Count -gt 0) {
            try {
                $AllNoSqlData |
                    Export-Excel -Path $ConsolidatedOutputFile `
                                 -WorksheetName "Unmatched Users" `
                                 -AutoSize
                Write-Host "All unmatched users have been exported to 'Unmatched Users' worksheet." -ForegroundColor Green
            }
            catch {
                Write-Host "Failed to export Unmatched Users: $($_.Exception.Message)" -ForegroundColor Red
            }
        }
        else {
            Write-Warning "No unmatched users found to export."
        }

        # Export terminated users report
        if ($TerminatedActiveAdUsers.Count -gt 0) {
            try {
                $TerminatedActiveAdUsers |
                    Export-Excel -Path $ConsolidatedOutputFile `
                                 -WorksheetName "Terminated Users" `
                                 -AutoSize -Append
                Write-Host "Terminated active AD users have been exported to 'Terminated Users' worksheet." -ForegroundColor Yellow
            }
            catch {
                Write-Host "Failed to export Terminated Active Users: $($_.Exception.Message)" -ForegroundColor Red
            }
        }
        else {
            Write-Warning "No terminated active AD users found to export."
        }

        # Export updated users report with formatting
        if ($AllUpdatedUsers.Count -gt 0) {
            $AllUpdatedUsers | Export-Excel -Path $ConsolidatedOutputFile -WorksheetName "Updated Users" -AutoSize -Append

            # Apply Excel formatting to highlight changes
            $excelPackage = Open-ExcelPackage -Path $ConsolidatedOutputFile
            $ws          = $excelPackage.Workbook.Worksheets["Updated Users"]

            $rowCount = $ws.Dimension.Rows

            # Format each row to highlight changes
            for ($row = 2; $row -le $rowCount; $row++) {
                $officeCell = $ws.Cells[$row,2].Value
                $deptCell   = $ws.Cells[$row,3].Value
                $sqlDeptCell = $ws.Cells[$row,4].Value
                $descCell   = $ws.Cells[$row,6].Value
                $titleCell  = $ws.Cells[$row,7].Value
                $sqlTitleCell = $ws.Cells[$row,8].Value
                $empNumCell = $ws.Cells[$row,9].Value
                $ouCell = $ws.Cells[$row,10].Value

                $hasIssue = $false
                $formattedCells = @()

                # Highlight Office vs SQL Department differences
                if ($null -eq $officeCell -or $null -eq $sqlDeptCell -or ($officeCell -cne $sqlDeptCell)) {
                    Set-ExcelRange -Worksheet $ws -Range "B${row}:B${row}" -BackgroundColor LightBlue
                    Set-ExcelRange -Worksheet $ws -Range "D${row}:D${row}" -BackgroundColor LightGreen
                    $formattedCells += "B${row}", "D${row}"
                    $hasIssue = $true
                }

                # Highlight Department vs SQL Department differences
                if ($null -eq $deptCell -or $null -eq $sqlDeptCell -or ($deptCell -cne $sqlDeptCell)) {
                    Set-ExcelRange -Worksheet $ws -Range "C${row}:C${row}" -BackgroundColor LightBlue
                    Set-ExcelRange -Worksheet $ws -Range "D${row}:D${row}" -BackgroundColor LightGreen
                    $formattedCells += "C${row}", "D${row}"
                    $hasIssue = $true
                }

                # Highlight Title vs SQL Title differences
                if ($null -eq $titleCell -or $null -eq $sqlTitleCell -or ($titleCell -cne $sqlTitleCell)) {
                    Set-ExcelRange -Worksheet $ws -Range "G${row}:G${row}" -BackgroundColor LightBlue
                    Set-ExcelRange -Worksheet $ws -Range "H${row}:H${row}" -BackgroundColor LightGreen
                    $formattedCells += "G${row}", "H${row}"
                    $hasIssue = $true
                }

                # Highlight Description vs SQL Title differences
                if ($null -eq $descCell -or $null -eq $sqlTitleCell -or ($descCell -cne $sqlTitleCell)) {
                    Set-ExcelRange -Worksheet $ws -Range "F${row}:F${row}" -BackgroundColor LightBlue
                    Set-ExcelRange -Worksheet $ws -Range "H${row}:H${row}" -BackgroundColor LightGreen
                    $formattedCells += "F${row}", "H${row}"
                    $hasIssue = $true
                }

                # Highlight changed Employee Number
                if ("EmployeeNumber" -in $Changes.Keys) {
                    Set-ExcelRange -Worksheet $ws -Range "I${row}:I${row}" -BackgroundColor Green 
                    $formattedCells += "I${row}"
                    $hasIssue = $true
                }

                # Highlight changed OU
                if ("OU" -in $Changes.Keys) {
                    Set-ExcelRange -Worksheet $ws -Range "J${row}:J${row}" -BackgroundColor Green
                    $formattedCells += "J${row}"
                    $hasIssue = $true
                }

                # Highlight entire row for any changes
                if ($hasIssue) {
                    $columns = 1..9
                    foreach ($col in $columns) {
                        $cellAddress = "$([char](64 + $col))$row"
                        if ($formattedCells -notcontains $cellAddress) {
                            Set-ExcelRange -Worksheet $ws -Range "$cellAddress`:$cellAddress" -BackgroundColor LightYellow
                        }
                    }
                }
            }

            # Save Excel formatting changes
            Close-ExcelPackage $excelPackage
            Write-Host "All updated users have been exported to 'Updated Users' worksheet." -ForegroundColor Green
        }

        # Export admin access required report
        if ($Admins.Count -gt 0) {
            try {
                $Admins |
                    Export-Excel -Path $ConsolidatedOutputFile `
                                 -WorksheetName "Admins" `
                                 -AutoSize -Append
                Write-Host "Users with insufficient access rights have been exported to 'Admins' worksheet." -ForegroundColor Green
            }
            catch {
                Write-Host "Failed to export Admins: $($_.Exception.Message)" -ForegroundColor Red
            }
        }

        # Export disabled users report
        if ($DisabledUsers.Count -gt 0) {
            try {
                $DisabledUsers |
                    Export-Excel -Path $ConsolidatedOutputFile `
                                 -WorksheetName "Disabled Users" `
                                 -AutoSize -Append
                Write-Host "Disabled users with changes have been exported to 'Disabled Users' worksheet." -ForegroundColor Green
            }
            catch {
                Write-Host "Failed to export Disabled Users: $($_.Exception.Message)" -ForegroundColor Red
            }
        }
        else {
            Write-Warning "No disabled users with changes found to export."
        }

        # Close SQL Connection
        Close-SqlConnection -Connection $conn
    }
}
catch {
    # Error handling for the entire script
    Write-Host "An error occurred: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Line: $($_.InvocationInfo.ScriptLineNumber) | Command: $($_.InvocationInfo.MyCommand)" -ForegroundColor Red
}
