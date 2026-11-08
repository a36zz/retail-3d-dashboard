@echo off
title Update dashboard data
echo Reading data\sales_data.xlsx and updating index.html. Please wait...
echo.
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0scripts\Update-Data.ps1"
echo.
if errorlevel 1 (
  echo FAILED - see the message above.
) else (
  echo Done - index.html is updated. Open it in your browser.
)
pause
