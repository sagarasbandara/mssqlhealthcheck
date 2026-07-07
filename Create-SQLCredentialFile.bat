@echo off
setlocal

cd /d "%~dp0"
powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "$cred = Get-Credential; $cred | Export-Clixml -Path '.\sql-prod-credential.xml'"

endlocal
