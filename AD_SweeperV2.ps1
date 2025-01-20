<#
    Script Name: AD_SweeperV2.ps1
    Purpose    : Clean and update Active Directory (AD) attributes based on SQL data
    Author     : Devin Lacey
    Date       : 01/17/2025
#>

# Import Required Modules
Import-Module "C:\PowerShellModules\SqlUtils.ps1"      # Custom SQL utility functions
. "C:\PowerShellModules\Functions.ps1"                # Custom department & title mapping functions
Import-Module ActiveDirectory
Import-Module ImportExcel

# Departments to exclude
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

# Set Up Base Excel Export Path
$BaseExportPath = "C:\ADSweeper"

# Ensure Base Export Path exists
if (-not (Test-Path -Path $BaseExportPath)) {
    try {
        New-Item -ItemType Directory -Path $BaseExportPath -Force | Out-Null
        Write-Host "Created base export directory at '$BaseExportPath'." -ForegroundColor Green
    }
    catch {
        Write-Host "Failed to create base export directory: $($_.Exception.Message)" -ForegroundColor Red
    }
}

# Generate a unique timestamp for this run
$RunTimestamp = Get-Date -Format 'yyyyMMddHHmmss'

# This will be our single consolidated Excel file
$ConsolidatedOutputFile = Join-Path $BaseExportPath "AD_Sweeper-$RunTimestamp.xlsx"

# Define AD Fields to Retrieve
$AdFields = @("Office", "Department", "Description", "Title", "EmployeeNumber", "Name", "SamAccountName", "DistinguishedName")

# Initialize Master Collections
$AllNoSqlData             = @()
$TerminatedActiveAdUsers  = @()
$AllUpdatedUsers          = @()
$Admins = @()

# Initialize a hash table to track processed users
$ProcessedUsers = @{}

try {
    # Establish SQL Connection
    $conn = Connect-SqlServer
    if ($conn) {
        $AllAdUsers = Get-ADUser -Filter "enabled -eq 'true'" -Properties $AdFields

        $AdUsers = foreach ($AdUser in $AllAdUsers) {
            if (
                $ExcludedDepartments -notcontains $AdUser.Department -and
                $ExcludedDepartments -notcontains $AdUser.Office -and
                $AdUser.Name -notmatch "Josh Ford"
            ) {
                $AdUser
            }
        }

        Write-Host "Filtered AD user count: $($AdUsers.Count)"

        if (!$AdUsers) { 
            Write-Warning "No AD users found outside the excluded departments."
        }

        foreach ($AdUser in $AdUsers) {
            $FullName       = $AdUser.Name
            $EmployeeNumber = $AdUser.EmployeeNumber
            $Changes        = @{ }
            $SqlResult      = $null

            # Clean Up or Add an EmployeeNumber if Missing
            if ($EmployeeNumber) {
                $CleanEmployeeNumber = $EmployeeNumber.Trim()
                if ($CleanEmployeeNumber -ne $EmployeeNumber) {
                    $Changes["EmployeeNumber"] = $CleanEmployeeNumber
                    Set-ADUser -Identity $AdUser.DistinguishedName -EmployeeNumber $CleanEmployeeNumber
                    $AdUser.EmployeeNumber = $CleanEmployeeNumber
                }
                if ($CleanEmployeeNumber) {
                    $Query = @"
SELECT Department, JobTitle, EmployeeNumber, EmployeeStatus
FROM UT_HCJ_IDWORKS 
WHERE EmployeeNumber = '$CleanEmployeeNumber'
"@
                    $SqlResult = Execute-SqlQuery -Connection $conn -SqlQuery $Query
                }
            }
            else {
                # Look Up by Name if No EmployeeNumber
                $FirstName = $null
                $LastName  = $null

                if ($FullName -match ',') {
                    # "Last, First" Format
                    $NameParts = $FullName -split ', '
                    if ($NameParts.Count -ge 2) {
                        $LastName  = $NameParts[0]
                        $FirstName = $NameParts[1]
                    }
                }
                else {
                    # "First Last" Format
                    $NameParts = $FullName -split ' '
                    if ($NameParts.Count -ge 2) {
                        $FirstName = $NameParts[0]
                        $LastName  = $NameParts[-1]
                    }
                }

                if ($FirstName -and $LastName) {
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
                    # If Missing EmployeeNumber, Add It
                    if ([string]::IsNullOrEmpty($AdUser.EmployeeNumber) -and $Row.EmployeeNumber -and $Row.EmployeeStatus -eq 'Active') {
                        $Changes["EmployeeNumber"] = $Row.EmployeeNumber
                        Set-ADUser -Identity $AdUser.DistinguishedName -EmployeeNumber $Row.EmployeeNumber
                        $AdUser.EmployeeNumber = $Row.EmployeeNumber
                        Write-Host "Added EmployeeNumber '$($Row.EmployeeNumber)' to '$FullName'." -ForegroundColor DarkGray
                    }

                        # Convert SQL Department & Job Title
                        if ($Row.JobTitle -like "*Compliance*") {
                            $SqlDepartment = "Compliance"
                        }
                        else {
                            $SqlDepartment = ConvertDeptName -department $Row.Department
                        }
                        $SqlTitle = MapJobTitles -adTitle $Row.JobTitle
                        
                        # Compare and Update Department and Office
                        if ($SqlDepartment -and (($null -eq $AdUser.Department) -or ($SqlDepartment -cne $AdUser.Department))) {
                            $Changes["Department"] = $SqlDepartment
                        }
                        if ($SqlDepartment -and (($null -eq $AdUser.Office) -or ($SqlDepartment -cne $AdUser.Office))) {
                            $Changes["physicalDeliveryOfficeName"] = $SqlDepartment
                        }

                    # Compare and Update Title and Description
                    if ($SqlTitle -and (($null -eq $AdUser.Title) -or ($SqlTitle -cne $AdUser.Title))) {
                        $Changes["Title"] = $SqlTitle
                    }
                    if ($SqlTitle -and (($null -eq $AdUser.Description) -or ($SqlTitle -cne $AdUser.Description))) {
                        $Changes["Description"] = $SqlTitle
                    }

                    # Track in Updated Users if There's at Least One Difference
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
                            DN             = $AdUser.DistinguishedName
                        }
                    }

                        # Apply Changes to AD
                        if ($Changes.Count -gt 0) {
                            foreach ($key in $Changes.Keys) {
                                # Create a unique key for the user and change type
                                $uniqueKey = "$($AdUser.SamAccountName)-$key"

                                # Check if the user has already been processed for this change
                                if (-not $ProcessedUsers.ContainsKey($uniqueKey)) {
                                    try {
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
                                        # Mark this user as processed for this change
                                        $ProcessedUsers[$uniqueKey] = $true
                                    }
                                    catch {
                                        Write-Host "Insufficient access rights for user: $($AdUser.Name)" -ForegroundColor Red
                                        # Add user to Admins collection if not already added
                                        if (-not $Admins | Where-Object { $_.SamAccountName -eq $AdUser.SamAccountName }) {
                                            $Admins += [PSCustomObject]@{
                                                Name              = $AdUser.Name
                                                SamAccountName    = $AdUser.SamAccountName
                                                Department        = $AdUser.Department
                                                Title             = $AdUser.Title
                                                EmployeeNumber    = $AdUser.EmployeeNumber
                                                DistinguishedName = $AdUser.DistinguishedName
                                            }
                                        }
                                    }
                                }
                            }
                        }

                        # Identify Terminated Active AD Users
                        if ($AdUser.Enabled -and $Row.EmployeeStatus -eq 'Terminated') {
                            $TerminatedActiveAdUsers += [PSCustomObject]@{
                                Name              = $AdUser.Name
                                SamAccountName    = $AdUser.SamAccountName
                                Department        = $AdUser.Department
                                Title             = $AdUser.Title
                                EmployeeNumber    = $AdUser.EmployeeNumber
                                EmployeeStatus    = $Row.EmployeeStatus
                                DistinguishedName = $AdUser.DistinguishedName
                            }
                        }
                    }
                }
                else {
                    # No Matching SQL Data
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

        if ($TerminatedActiveAdUsers.Count -gt 0) {
            Write-Host "Accumulated $($TerminatedActiveAdUsers.Count) terminated active AD users." -ForegroundColor Magenta
        }
        else {
            Write-Host "No terminated active AD users found." -ForegroundColor Yellow
        }
        
        if ($AllNoSqlData.Count -gt 0) {
            Write-Host "Accumulated $($AllNoSqlData.Count) unmatched users." -ForegroundColor Cyan
        }
        else {
            Write-Host "No unmatched users found." -ForegroundColor Yellow
        }

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

        if ($AllUpdatedUsers.Count -gt 0) {
            $AllUpdatedUsers | Export-Excel -Path $ConsolidatedOutputFile -WorksheetName "Updated Users" -AutoSize -Append

            # Excel formatting for updated users
            $excelPackage = Open-ExcelPackage -Path $ConsolidatedOutputFile
            $ws          = $excelPackage.Workbook.Worksheets["Updated Users"]

            $rowCount = $ws.Dimension.Rows

            for ($row = 2; $row -le $rowCount; $row++) {
                $officeCell = $ws.Cells[$row,2].Value
                $deptCell   = $ws.Cells[$row,3].Value
                $sqlDeptCell = $ws.Cells[$row,4].Value
                $descCell   = $ws.Cells[$row,6].Value
                $titleCell  = $ws.Cells[$row,7].Value
                $sqlTitleCell = $ws.Cells[$row,8].Value
                $empNumCell = $ws.Cells[$row,9].Value

                $hasIssue = $false
                $formattedCells = @()

                # Office vs SQL Dept
                if ($null -eq $officeCell -or $null -eq $sqlDeptCell -or ($officeCell -cne $sqlDeptCell)) {
                    Set-ExcelRange -Worksheet $ws -Range "B${row}:B${row}" -BackgroundColor LightBlue
                    Set-ExcelRange -Worksheet $ws -Range "D${row}:D${row}" -BackgroundColor LightGreen
                    $formattedCells += "B${row}", "D${row}"
                    $hasIssue = $true
                }

                # Department vs SQL Dept
                if ($null -eq $deptCell -or $null -eq $sqlDeptCell -or ($deptCell -cne $sqlDeptCell)) {
                    Set-ExcelRange -Worksheet $ws -Range "C${row}:C${row}" -BackgroundColor LightBlue
                    Set-ExcelRange -Worksheet $ws -Range "D${row}:D${row}" -BackgroundColor LightGreen
                    $formattedCells += "C${row}", "D${row}"
                    $hasIssue = $true
                }

                # Title vs SQL Title
                if ($null -eq $titleCell -or $null -eq $sqlTitleCell -or ($titleCell -cne $sqlTitleCell)) {
                    Set-ExcelRange -Worksheet $ws -Range "G${row}:G${row}" -BackgroundColor LightBlue
                    Set-ExcelRange -Worksheet $ws -Range "H${row}:H${row}" -BackgroundColor LightGreen
                    $formattedCells += "G${row}", "H${row}"
                    $hasIssue = $true
                }

                # Description vs SQL Title
                if ($null -eq $descCell -or $null -eq $sqlTitleCell -or ($descCell -cne $sqlTitleCell)) {
                    Set-ExcelRange -Worksheet $ws -Range "F${row}:F${row}" -BackgroundColor LightBlue
                    Set-ExcelRange -Worksheet $ws -Range "H${row}:H${row}" -BackgroundColor LightGreen
                    $formattedCells += "F${row}", "H${row}"
                    $hasIssue = $true
                }

                # Highlight EmployeeNumber if it's changed
                if ("EmployeeNumber" -in $Changes.Keys) {
                    Set-ExcelRange -Worksheet $ws -Range "I${row}:I${row}" -BackgroundColor '#FFCCCC' 
                    $formattedCells += "I${row}"
                    $hasIssue = $true
                }

                # If there was any difference, highlight the row in Yellow (except already formatted cells)
                if ($hasIssue) {
                    # Example: highlight columns A..I
                    $columns = 1..9
                    foreach ($col in $columns) {
                        $cellAddress = "$([char](64 + $col))$row"
                        if ($formattedCells -notcontains $cellAddress) {
                            Set-ExcelRange -Worksheet $ws -Range "$cellAddress`:$cellAddress" -BackgroundColor LightYellow
                        }
                    }
                }
            }

            # Save changes
            Close-ExcelPackage $excelPackage
            Write-Host "All updated users have been exported to 'Updated Users' worksheet." -ForegroundColor Green
        }

        # Export Admins to a new worksheet
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
        else {
            Write-Warning "No users with insufficient access rights found to export."
        }

        # Close SQL Connection
        Close-SqlConnection -Connection $conn
    }
catch {
    Write-Host "An error occurred: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Line: $($_.InvocationInfo.ScriptLineNumber) | Command: $($_.InvocationInfo.MyCommand)" -ForegroundColor Red
}

