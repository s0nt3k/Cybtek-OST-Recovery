@echo off
rem Double-click to save Outlook data files for db4gtzehner@yahoo.com.
rem Runs as administrator so restore-point copies are included (Windows asks for approval).
rem Copies go to a MailRescue-<computer>-<time> folder next to this file.
setlocal
set "PS=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"
set "SCRIPT=%~dp0Save-OutlookMailCache.ps1"
net session >nul 2>&1
if errorlevel 1 (
  "%PS%" -NoProfile -Command "Start-Process -FilePath '%~f0' -Verb RunAs"
  exit /b
)
"%PS%" -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT%" %*
echo.
pause
