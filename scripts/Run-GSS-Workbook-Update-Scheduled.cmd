@echo off
setlocal
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Update-GSS-MainWorkbook.ps1" -Folder "%~dp0..\.." -Apply
exit /b %ERRORLEVEL%
