@echo off
cd /d "%~dp0"
echo === Smart Form Filler - first-time setup ===
echo.

python -c "import sys; assert sys.version_info >= (3, 9)" >nul 2>&1
if errorlevel 1 (
  echo Python was not found. Installing Python 3.12 ...
  winget install -e --id Python.Python.3.12 --scope user --accept-package-agreements --accept-source-agreements
  if errorlevel 1 (
    echo.
    echo Could not install Python automatically.
    echo Download it from https://www.python.org/downloads/ and tick "Add python.exe to PATH" when installing.
    echo Then double-click this file again.
    pause
    exit /b 1
  )
  echo.
  echo Python is installed. Close this window and double-click this file again to finish setup.
  pause
  exit /b 0
)

rem A .venv copied from another PC does not work here, so rebuild it.
if exist ".venv\Scripts\python.exe" (
  ".venv\Scripts\python.exe" -c "import sys" >nul 2>&1 || rmdir /s /q .venv
)
if not exist ".venv\Scripts\python.exe" (
  echo Creating the tool's Python environment ...
  python -m venv .venv
)

echo Installing the tool's libraries - this needs internet and takes a few minutes ...
".venv\Scripts\python.exe" -m pip install --quiet --upgrade pip
".venv\Scripts\python.exe" -m pip install --timeout 120 --retries 5 -r requirements.txt
if errorlevel 1 (
  echo.
  echo Installing failed. Check the internet connection and double-click this file again.
  pause
  exit /b 1
)

echo.
echo Setup complete. Now double-click "Start Smart Form Filler.bat".
pause
