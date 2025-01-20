# AD Sweeper Script

## Overview
This script is designed to streamline the management of Active Directory (AD) users by updating employee attributes and correcting any inaccurate fields using SQL data. The script categorizes users into three groups:

### 1. **Unmatched Users**
   - These users do not have corresponding SQL data or failed to retrieve SQL data.
   - **Action Required:**
     - Run the 'Add Employee Number.bat' tool to generate the necessary data.
     - Manually update these users by:
       1. Inputting their full AD name as it appears in the generated Excel sheet.
       2. Adding their employee number, obtained from JAM reports.

### 2. **Terminated Users**
   - Active users in AD whose SQL data is marked as "Terminated."
   - **Action Required:**
     - Manually disable these accounts in AD.
     - Log in to the Microsoft 365 Admin Center and:
       - Verify if the user has any active licenses.
       - Remove any active licenses if they exist.

### 3. **Updated Users**
   - These are users whose attributes were corrected or updated based on SQL data.
   - **No Manual Action Required:**
     - This sheet is for documentation purposes only to track the changes made by the script.

## Setup Instructions

### Step 1: Create the Required Directory
To set up this script on your computer, create the required directory structure. Open your terminal and run the following command:

```bash
mkdir -p C:\ADSweeper\Script
```

### Step 2: Move Script Files
Move all the files provided in this folder to the directory created above (`C:\ADSweeper\Script`).

### Step 3: Run the Script
To execute the script, run the batch file `AD SweeperV2.bat` by:
1. Navigating to the `C:\ADSweeper\Script` folder in your terminal or file explorer.
2. Double-clicking on the `AD SweeperV2.bat` file, or running the following command in your terminal:

```bash
C:\ADSweeper\Script\AD SweeperV2.bat
```

## Notes
- Ensure you have the necessary permissions to access Active Directory and execute the script.
- Always review the generated sheets to verify changes before performing manual updates.


For additional support, contact Devin Lacey at Devin.Lacey@JamulCasino.com

---

This README is intended to guide users in setting up and running the AD Sweeper script effectively. Follow the steps carefully to ensure proper execution and results.

