@echo off
setlocal enabledelayedexpansion

set "CHROME_PROFILE=%LOCALAPPDATA%\sam_automation_chrome_profile"

set "CHROME_EXE="
if exist "%ProgramFiles%\Google\Chrome\Application\chrome.exe" set "CHROME_EXE=%ProgramFiles%\Google\Chrome\Application\chrome.exe"
if not defined CHROME_EXE if exist "%ProgramFiles(x86)%\Google\Chrome\Application\chrome.exe" set "CHROME_EXE=%ProgramFiles(x86)%\Google\Chrome\Application\chrome.exe"
if not defined CHROME_EXE if exist "%LOCALAPPDATA%\Google\Chrome\Application\chrome.exe" set "CHROME_EXE=%LOCALAPPDATA%\Google\Chrome\Application\chrome.exe"

if not defined CHROME_EXE (
    echo Could not find chrome.exe in the usual install locations.
    echo Edit this file and set CHROME_EXE to your Chrome path manually.
    pause
    exit /b 1
)

echo Launching Chrome with remote debugging on port 9222...
echo Profile: %CHROME_PROFILE%
echo.
echo Log into SAM.gov in the window that opens ^(first time only - the
echo session is saved in this profile for next time^), then leave it open
echo and run sam_automation.py from a separate terminal.
echo.

start "" "%CHROME_EXE%" --remote-debugging-port=9222 --user-data-dir="%CHROME_PROFILE%" "https://sam.gov/search"
