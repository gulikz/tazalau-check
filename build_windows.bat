@echo off
rem Tazalau Check - build a standalone program for colleagues (no Python needed on their PCs).
rem Result: dist\TazalauCheck\TazalauCheck.exe  - give colleagues the whole dist\TazalauCheck folder.
setlocal
cd /d "%~dp0"

where py >nul 2>nul
if %errorlevel%==0 (set "PY=py -3") else (set "PY=python")

if not exist ".venv\Scripts\python.exe" (
    %PY% -m venv .venv
    if errorlevel 1 goto :fail
)
".venv\Scripts\python.exe" -m pip install --upgrade pip
".venv\Scripts\python.exe" -m pip install -r requirements-dev.txt
if errorlevel 1 goto :fail

echo.
echo === 1/4 Automatic tests (no internet needed) ===
".venv\Scripts\python.exe" -m pytest -q
if errorlevel 1 goto :fail

echo.
echo === 2/4 Tests against the real site tazalau.qoldau.kz ===
".venv\Scripts\python.exe" -m pytest -q -m live
if errorlevel 1 goto :fail

echo.
echo === 3/4 Building the program ===
".venv\Scripts\python.exe" -m PyInstaller --noconfirm --clean tazalau_check.spec
if errorlevel 1 goto :fail

echo.
echo === 4/4 Self-test of the built program (browser + site) ===
start /wait "" "dist\TazalauCheck\TazalauCheck.exe" --selftest
if errorlevel 1 goto :selftestfail

echo.
echo DONE. Give colleagues the whole folder:  dist\TazalauCheck
echo They start the program with:            TazalauCheck.exe
pause
exit /b 0

:selftestfail
echo.
echo ERROR: the built program could not start the browser or reach the site.
pause
exit /b 1

:fail
echo.
echo ERROR: the build was stopped, see the messages above.
pause
exit /b 1
