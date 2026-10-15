@echo off
rem Starts the scanner. Extra arguments are passed through, e.g.:  run.bat --demo slip.jpg
cd /d "%~dp0"
if not exist ".venv\Scripts\python.exe" call setup.bat
".venv\Scripts\python.exe" -m quiniela_scanner %*
if errorlevel 1 pause
