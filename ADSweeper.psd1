<#
    File Name: ADSweeper.psd1
    Purpose  : PowerShell module manifest for ADSweeper
    Author   : Devin Lacey
    Date     : 01/17/2025

    Description:
    This is the module manifest file for the ADSweeper PowerShell module.
    It defines:
    - Module metadata (version, author, etc.)
    - Module dependencies
    - Exported functions and aliases
    - Required assemblies and files

    This manifest is required for proper module loading and versioning.
#>

@{
    # Module metadata
    ModuleVersion = '1.0.0'
    GUID          = '01234567-89ab-cdef-0123-456789abcdef'
    Author        = 'Devin Lacey'
    Description   = 'Clean and update AD attributes based on SQL data.'
    RootModule    = 'ADSweeper.psm1'

    # Additional metadata (optional but recommended)
    CompanyName   = 'Jamul Casino'
    Copyright     = 'Devin Lacey. All rights reserved.'
    
    # Minimum PowerShell version required
    PowerShellVersion = '5.1'

    # Required modules
    RequiredModules = @(
        'ActiveDirectory'
    )

    # Functions to export
    FunctionsToExport = @('Invoke-ADSweeper')
}
