@echo off
rem One-time setup: creates a private Python environment and installs what the scanner needs.
cd /d "%~dp0"
where py >nul 2>nul
if errorlevel 1 (
  echo Python 3.10 or newer is required. Install it from https://www.python.org/downloads/
  echo and tick "Add python.exe to PATH" in the installer, then run this file again.
  pause
  exit /b 1
)
if not exist ".venv\Scripts\python.exe" py -3 -m venv .venv
call ".venv\Scripts\activate.bat"
python -m pip install --upgrade pip
pip install -r requirements.txt
if errorlevel 1 (
  echo.
  echo Install failed - see the messages above.
  pause
  exit /b 1
)
echo.
echo Setup finished. Double-click run.bat to start the scanner.
pause
