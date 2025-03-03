@REM =================================================================
@REM File Name: Add Employee Number.bat
@REM Purpose  : Launcher for Add_Employee_Number.ps1 with elevated privileges
@REM Author   : Devin Lacey
@REM Date     : 01/17/2025
@REM 
@REM Description:
@REM This batch file provides a secure way to run the Add_Employee_Number.ps1
@REM PowerShell script with proper domain credentials. It:
@REM - Prompts for domain username
@REM - Runs the script with elevated privileges
@REM - Uses bypass execution policy for this session only
@REM =================================================================

@echo off
echo Using DOMAIN: Jamulcasinosd.com
set /p USERNAME=Enter the username (without domain): 
runas /user:Jamulcasinosd.com\%USERNAME% "powershell.exe -NoProfile -ExecutionPolicy Bypass -File \"C:\ADSweeper\Script\Add_Employee_Number.ps1\""
pause
