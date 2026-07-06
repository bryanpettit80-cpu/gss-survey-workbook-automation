@echo off
setlocal
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0scripts\Update-GSS-MainWorkbook.ps1" -Folder "%~dp0.."
set EXITCODE=%ERRORLEVEL%
echo.
if "%EXITCODE%"=="0" (
  echo GSS workbook copy-test finished.
) else (
  echo GSS workbook copy-test failed with exit code %EXITCODE%.
)
echo.
pause
exit /b %EXITCODE%
