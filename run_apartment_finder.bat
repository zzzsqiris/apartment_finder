@echo off
setlocal
cd /d "%~dp0"

where py >nul 2>nul
if %errorlevel%==0 (
  py -3 run_apartment_finder.py
  goto :eof
)

where python >nul 2>nul
if %errorlevel%==0 (
  python run_apartment_finder.py
  goto :eof
)

echo Python was not found. Install Python 3 from https://www.python.org/downloads/windows/
echo During install, check "Add python.exe to PATH", then reopen this window.
pause
exit /b 1
