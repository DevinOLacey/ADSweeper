# AD Sweeper Script

## Overview
This script manages Active Directory (AD) users by updating employee attributes based on SQL data. Users are categorized into three groups:

1. **Unmatched Users:** No corresponding SQL data found.
   - **Action Required:** Run `Add Employee Number.bat` to generate data, then manually update users in AD.

2. **Terminated Users:** Marked as "Terminated" in SQL but still active in AD.
   - **Action Required:** Disable accounts in AD and remove any active Microsoft 365 licenses.

3. **Updated Users:** Attributes corrected based on SQL data.
   - **No manual action required.** This is for documentation purposes only.

## Setup Instructions

### 1. Download and Extract
1. Download the repository as a ZIP from GitHub.
2. Extract the contents to a convenient location.

### 2. Create Directory and Move Files
1. Create the required directory:
   ```
   mkdir -p C:\ADSweeper\Script
   ```

2. Move all extracted files into `C:\ADSweeper\Script` maintaining this structure:
   ```
   C:\ADSweeper\Script\
   ├── ADSweeper.psd1
   ├── ADSweeper.psm1
   ├── AD_SweeperV2.ps1
   ├── Add_Employee_Number.ps1
   ├── Add Employee Number.bat
   ├── Functions.ps1
   ├── SqlUtils.ps1
   ├── encrypted_password.txt
   └── Dependencies\
       └── Microsoft.Data.SqlClient.5.0.1\
           └── lib\
               └── netstandard2.0\
                   └── Microsoft.Data.SqlClient.dll
   ```

### 3. Run the Script
Navigate to `C:\ADSweeper\Script` and run:
```
C:\ADSweeper\Script\AD SweeperV2.bat
```
> Note: This batch file requires elevated privileges.

### 4. Additional Tools
To update employee numbers for unmatched users, run:
```
C:\ADSweeper\Script\Add Employee Number.bat
```

## Requirements
- **Permissions:** Ensure you have the necessary AD and script execution permissions
- **Dependencies:** 
  - RSAT tools
  - ImportExcel PowerShell module
- **Review Process:** Always verify changes in the generated Excel files before applying manual updates

## Support
For support, contact Devin Lacey at Devin.Lacey@JamulCasino.com

### Key Improvements:
- **Concise formatting:** Uses bullet points and sections for clarity.
- **Minimal but essential details:** Removed redundant wording while preserving all necessary setup steps.
- **Proper Markdown formatting:** Ensures compatibility with GitHub and easy readability.

This **README.md** can be placed in your GitHub repository, providing a simple, structured guide for users to download, install, and execute the AD Sweeper script. 🚀