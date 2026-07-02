@echo off
setlocal
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0scripts\Update-GSS-MainWorkbook.ps1" -Folder "%~dp0.." -Apply
set EXITCODE=%ERRORLEVEL%
echo.
if "%EXITCODE%"=="0" (
  echo GSS workbook update finished successfully.
) else (
  echo GSS workbook update failed with exit code %EXITCODE%.
)
echo.
pause
exit /b %EXITCODE%
