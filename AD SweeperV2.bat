@echo off
echo Using DOMAIN: Jamulcasinosd.com
set /p USERNAME=Enter the username (without domain): 
runas /user:Jamulcasinosd.com\%USERNAME% "powershell.exe -NoProfile -ExecutionPolicy Bypass -File \"C:\ADSweeper\Script\AD_SweeperV2.ps1""
pause
