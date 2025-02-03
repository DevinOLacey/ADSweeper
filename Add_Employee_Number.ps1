# Import the Active Directory module (ensure RSAT tools are installed)
Import-Module ActiveDirectory

do {
    # Prompt for FullName and EmployeeNumber
    $FullName = Read-Host "Enter the Full Name of the user (or type 'exit' to quit)"
    if ($FullName -eq 'exit') { break }

    $NewEmployeeNumber = Read-Host "Enter the new EmployeeNumber (or type 'exit' to quit)"
    if ($NewEmployeeNumber -eq 'exit') { break }

    # Search for the user in Active Directory
    try {
        $User = Get-ADUser -Filter "Name -eq '$FullName'" -Properties EmployeeNumber

        if ($User) {
            Write-Host "User found: $($User.Name)"
            Write-Host "Current EmployeeNumber: $($User.EmployeeNumber)"

            # Update the EmployeeNumber attribute
            Set-ADUser -Identity $User.DistinguishedName -EmployeeNumber $NewEmployeeNumber

            Write-Host "EmployeeNumber updated successfully to $NewEmployeeNumber" -ForegroundColor Green
        } else {
            Write-Host "No user found with the name $FullName" -ForegroundColor Red
        }
    } catch {
        Write-Host "An error occurred: $_" -ForegroundColor Red
    }

    Write-Host "-----------------------------"
    $continue = Read-Host "Do you want to update another user? (Y/N)"
} while ($continue -match '^(Y|y)$')
