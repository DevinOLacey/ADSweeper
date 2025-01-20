@echo off
echo Using DOMAIN: Jamulcasinosd.com
set /p USERNAME=Enter the username (without domain): 
runas /user:Jamulcasinosd.com\%USERNAME% "powershell.exe -NoProfile -ExecutionPolicy Bypass -File \"C:\ADSweeper\Script\Add_Employee_Number.ps1""
pause
