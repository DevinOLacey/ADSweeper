@REM =================================================================
@REM File Name: AD SweeperV2.bat
@REM Purpose  : Launcher for ADSweeper PowerShell module with elevated privileges
@REM Author   : Devin Lacey
@REM Date     : 01/17/2025
@REM 
@REM Description:
@REM This batch file provides a secure way to run the ADSweeper module with
@REM proper domain credentials. It:
@REM - Prompts for domain username
@REM - Imports the ADSweeper module
@REM - Runs Invoke-ADSweeper with elevated privileges
@REM - Uses bypass execution policy for this session only
@REM - Keeps the PowerShell window open after execution (-NoExit)
@REM =================================================================

@echo off
echo Using DOMAIN: Jamulcasinosd.com
set /p USERNAME=Enter the username (without domain): 
runas /user:Jamulcasinosd.com\%USERNAME% "powershell.exe -NoProfile -ExecutionPolicy Bypass -NoExit -Command \"Import-Module 'C:\ADSweeper\Script\ADSweeper.psd1'; Invoke-ADSweeper\""
pause
