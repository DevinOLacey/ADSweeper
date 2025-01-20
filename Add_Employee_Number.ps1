# Import the Active Directory module (ensure RSAT tools are installed)
Import-Module ActiveDirectory

# Prompt for FullName and EmployeeNumber
$FullName = Read-Host "Enter the Full Name of the user"
$NewEmployeeNumber = Read-Host "Enter the new EmployeeNumber"

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
