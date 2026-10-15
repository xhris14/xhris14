@echo off
rem One-time setup: creates a private Python environment and installs what the scanner needs.
cd /d "%~dp0"

rem Prefer a well-supported Python (3.12, then 3.13, 3.11, 3.10); fall back to whatever is newest.
set PY=
for %%V in (3.12 3.13 3.11 3.10) do (
  if not defined PY (
    py -%%V --version >nul 2>nul
    if not errorlevel 1 set PY=py -%%V
  )
)
if not defined PY (
  py -3 --version >nul 2>nul
  if not errorlevel 1 set PY=py -3
)
if not defined PY (
  python --version >nul 2>nul
  if not errorlevel 1 set PY=python
)
if not defined PY (
  echo Python was not found.
  echo Install Python 3.12 from https://www.python.org/downloads/ then run this file again.
  pause
  exit /b 1
)
echo Using: %PY%
%PY% --version

if not exist ".venv\Scripts\python.exe" %PY% -m venv .venv
if not exist ".venv\Scripts\python.exe" (
  echo Could not create the Python environment.
  pause
  exit /b 1
)
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
