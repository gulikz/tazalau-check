@echo off
rem Tazalau Check - run from source code (developer mode).
rem First run creates a private Python environment in .venv and installs the requirements.
setlocal
cd /d "%~dp0"

where py >nul 2>nul
if %errorlevel%==0 (set "PY=py -3") else (set "PY=python")

if not exist ".venv\Scripts\python.exe" (
    echo First run: installing components, this takes a few minutes...
    %PY% -m venv .venv
    if errorlevel 1 goto :nopython
    ".venv\Scripts\python.exe" -m pip install --upgrade pip
    ".venv\Scripts\python.exe" -m pip install -r requirements.txt
    if errorlevel 1 goto :pipfail
)

start "" ".venv\Scripts\pythonw.exe" tazalau_app.py %*
exit /b 0

:nopython
echo.
echo ERROR: Python 3.9 or newer was not found. Install it from https://www.python.org/downloads/
echo        ^(tick "Add python.exe to PATH" in the installer^) and run this file again.
pause
exit /b 1

:pipfail
echo.
echo ERROR: could not install the components ^(no internet access or a proxy is required^).
echo        Ask IT for the pip proxy settings, then delete the .venv folder and run this file again.
rmdir /s /q .venv 2>nul
pause
exit /b 1
