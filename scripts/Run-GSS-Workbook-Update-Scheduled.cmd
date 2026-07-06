@echo off
setlocal
rem Task Scheduler launcher only. This runs live apply without an interactive prompt.
rem The scheduled task is intentionally disabled unless manual scheduling is requested.
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Update-GSS-MainWorkbook.ps1" -Folder "%~dp0..\.." -Apply
exit /b %ERRORLEVEL%
