# ADSweeper.psm1

<#
    Module Name: ADSweeper.psm1
    Purpose    : PowerShell module implementation for ADSweeper
    Author     : Devin Lacey
    Date       : 01/17/2025

    Description:
    This PowerShell module serves as the main entry point for the ADSweeper tool.
    It provides:
    - Function exports for the ADSweeper toolset
    - Integration of helper functions from Functions.ps1 and SqlUtils.ps1
    - Main execution function (Invoke-ADSweeper)

    Dependencies:
    - Functions.ps1: Contains AD attribute mapping and conversion functions
    - SqlUtils.ps1: Provides SQL Server connectivity functions
    - AD_SweeperV2.ps1: Contains main AD synchronization logic
#>

# Dot-source helper scripts so their functions (including Connect-SqlServer) become available.
. "$PSScriptRoot\Functions.ps1"
. "$PSScriptRoot\SqlUtils.ps1"

Write-Verbose "Helper functions loaded from $PSScriptRoot"

<#
.SYNOPSIS
    Main entry point for the ADSweeper tool.

.DESCRIPTION
    Executes the AD Sweeper synchronization process by loading and running
    the main script logic from AD_SweeperV2.ps1.

.NOTES
    This function requires appropriate AD permissions and SQL Server access.
    It should be run with administrative privileges.

.EXAMPLE
    Invoke-ADSweeper
    Runs the AD synchronization process.
#>
function Invoke-ADSweeper {
    [CmdletBinding()]
    param()

    Write-Host ">>> Starting Invoke-ADSweeper..." -ForegroundColor Yellow

    # Dot-source the main AD Sweeper logic file.
    . "$PSScriptRoot\AD_SweeperV2.ps1"
    
    Write-Host ">>> Finished processing in Invoke-ADSweeper." -ForegroundColor Yellow
}

# Export the public functions (for testing, we export the helper functions as well).
Export-ModuleMember -Function Invoke-ADSweeper
