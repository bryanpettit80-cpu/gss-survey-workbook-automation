@echo off
setlocal
rem Operator launcher. It always runs a copy-test before asking for live apply.
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0GSS Survey Workbook Automation\scripts\Invoke-GSS-SafeWorkbookUpdate.ps1" -Folder "%~dp0."
set EXITCODE=%ERRORLEVEL%
echo.
if "%EXITCODE%"=="0" (
  echo GSS workbook launcher finished.
) else (
  echo GSS workbook launcher failed with exit code %EXITCODE%.
)
echo.
pause
exit /b %EXITCODE%
