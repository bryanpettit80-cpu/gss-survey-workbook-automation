@echo off
setlocal
title GSS Survey Workbook Update
echo.
echo GSS SURVEY WORKBOOK UPDATE
echo ==========================
echo This will test the update first. The live workbook will not change
echo unless the test passes and you type APPLY when prompted.
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
