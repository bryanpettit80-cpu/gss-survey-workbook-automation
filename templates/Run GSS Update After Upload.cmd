@echo off
setlocal
title GSS Survey Workbook Update
echo.
echo GSS SURVEY WORKBOOK UPDATE
echo ==========================
echo This will validate the update and preserve the previous good report.
echo Complete, valid inputs will be processed without another confirmation.
echo Conflicting corrections require a specific choice.
echo Google Drive backup has a separate status and can be retried.
echo Automatic email sending is permanently disabled.
echo Only one computer may run GSS at a time.
echo Close this window if a GSS update is already running on another computer.
echo.
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0GSS Survey Workbook Automation\scripts\Invoke-GSS-SafeWorkbookUpdate.ps1" -Folder "%~dp0."
set EXITCODE=%ERRORLEVEL%
echo.
if "%EXITCODE%"=="0" (
  echo Finished. Review the outcome shown above.
) else (
  echo ATTENTION NEEDED: The update did not finish normally.
  echo No further action should be taken until the message above is reviewed.
)
echo.
echo You may close this window after reviewing the result.
pause
exit /b %EXITCODE%
