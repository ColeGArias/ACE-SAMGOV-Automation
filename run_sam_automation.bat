@echo off
setlocal

cd /d "%~dp0"
set "WRAPPER_LOG=%~dp0scheduler_wrapper.log"

call "%~dp0venv\Scripts\activate.bat"
python sam_automation.py
set "EXIT_CODE=%ERRORLEVEL%"

echo %date% %time% - sam_automation.py exited with code %EXIT_CODE% >> "%WRAPPER_LOG%"

exit /b %EXIT_CODE%
