@echo off
echo Using DOMAIN: Jamulcasinosd.com
set /p USERNAME=Enter the username (without domain): 
runas /user:Jamulcasinosd.com\%USERNAME% "powershell.exe -NoProfile -ExecutionPolicy Bypass -NoExit -Command \"Import-Module 'C:\ADSweeper\Script\ADSweeper.psd1'; Invoke-ADSweeper\""
pause
