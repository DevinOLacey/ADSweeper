<#
    Script Name: Add_Employee_Number.ps1
    Purpose    : Add or update employee numbers for Active Directory users
    Author     : Devin Lacey
    Date       : 01/17/2025

    Description:
    This script provides an interactive interface to add or update employee numbers
    in Active Directory. It allows administrators to:
    - Search for users by their full name
    - View current employee numbers
    - Update employee numbers for existing users
    - Process multiple users in succession

    Requirements:
    - Active Directory PowerShell module (RSAT tools)
    - Appropriate AD permissions to modify user attributes
    - User must have rights to modify employeeNumber attribute

    Usage:
    1. Run the script in PowerShell
    2. Enter the user's full name when prompted
    3. Enter the new employee number
    4. Confirm the changes
    5. Choose to continue with another user or exit
#>

# Import the Active Directory module (ensure RSAT tools are installed)
Import-Module ActiveDirectory

do {
    # Prompt for user input
    $FullName = Read-Host "Enter the Full Name of the user (or type 'exit' to quit)"
    if ($FullName -eq 'exit') { break }

    $NewEmployeeNumber = Read-Host "Enter the new EmployeeNumber (or type 'exit' to quit)"
    if ($NewEmployeeNumber -eq 'exit') { break }

    # Search for the user in Active Directory and attempt to update their employee number
    try {
        # Search for user by full name
        $User = Get-ADUser -Filter "Name -eq '$FullName'" -Properties EmployeeNumber

        if ($User) {
            # Display current user information
            Write-Host "User found: $($User.Name)"
            Write-Host "Current EmployeeNumber: $($User.EmployeeNumber)"

            # Update the EmployeeNumber attribute
            Set-ADUser -Identity $User.DistinguishedName -EmployeeNumber $NewEmployeeNumber

            # Confirm successful update
            Write-Host "EmployeeNumber updated successfully to $NewEmployeeNumber" -ForegroundColor Green
        } else {
            # User not found notification
            Write-Host "No user found with the name $FullName" -ForegroundColor Red
        }
    } catch {
        # Error handling for AD operations
        Write-Host "An error occurred: $_" -ForegroundColor Red
    }

    # Separator for visual clarity
    Write-Host "-----------------------------"
    
    # Prompt to continue or exit
    $continue = Read-Host "Do you want to update another user? (Y/N)"
} while ($continue -match '^(Y|y)$')
