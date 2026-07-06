@echo off
setlocal
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0scripts\Invoke-GSS-SafeWorkbookUpdate.ps1" -Folder "%~dp0.."
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
