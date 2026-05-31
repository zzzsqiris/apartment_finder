@echo off
setlocal
cd /d "%~dp0"

where py >nul 2>nul
if %errorlevel%==0 (
  py -3 -m pip install -r requirements.txt
  goto :done
)

where python >nul 2>nul
if %errorlevel%==0 (
  python -m pip install -r requirements.txt
  goto :done
)

echo Python was not found. Install Python 3 from https://www.python.org/downloads/windows/
echo During install, check "Add python.exe to PATH", then reopen this window.
pause
exit /b 1

:done
echo.
echo Dependencies installed. You can now run run_apartment_finder.bat
pause
