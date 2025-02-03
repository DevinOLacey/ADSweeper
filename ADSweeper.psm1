# ADSweeper.psm1

# Dot-source helper scripts so their functions (including Connect-SqlServer) become available.
. "$PSScriptRoot\Functions.ps1"
. "$PSScriptRoot\SqlUtils.ps1"

Write-Verbose "Helper functions loaded from $PSScriptRoot"

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
