@echo off
cd /d "%~dp0"
".venv\Scripts\python.exe" -c "import streamlit, playwright" >nul 2>&1
if errorlevel 1 (
  echo The tool is not set up on this PC yet. Running setup first ...
  call "Setup (first time).bat"
)
rem Each double-click starts the next free copy, so several websites can be filled at the same time.
".venv\Scripts\python.exe" launcher.py
pause
