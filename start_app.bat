@echo off
setlocal enabledelayedexpansion
cd /d "%~dp0"

echo ===================================================
echo     Master Inspection Plan — Starting Server
echo ===================================================
echo.

:: 1. Check if a local embeddable python executable exists in python/ folder
if exist "python\python.exe" (
    echo Using Portable Embeddable Python...
    set "PY_CMD=python\python.exe"
    goto RUN
)

:: 2. Check if python command is available in PATH
where python >nul 2>nul
if %errorlevel% equ 0 (
    echo Using System Python...
    set "PY_CMD=python"
    goto RUN
)

:: 3. Check if python3 command is available
where python3 >nul 2>nul
if %errorlevel% equ 0 (
    echo Using System Python 3...
    set "PY_CMD=python3"
    goto RUN
)

:: 4. Check standard Windows default install locations if not in PATH
if exist "%LocalAppData%\Programs\Python\Python311\python.exe" (
    set "PY_CMD=%LocalAppData%\Programs\Python\Python311\python.exe"
    goto RUN
)
if exist "%LocalAppData%\Programs\Python\Python312\python.exe" (
    set "PY_CMD=%LocalAppData%\Programs\Python\Python312\python.exe"
    goto RUN
)
if exist "%LocalAppData%\Programs\Python\Python313\python.exe" (
    set "PY_CMD=%LocalAppData%\Programs\Python\Python313\python.exe"
    goto RUN
)

echo ERROR: Python executable not found.
echo.
echo To run this application without installing Python:
echo Download Python Embeddable zip from python.org, extract it into a folder named "python" in this directory.
echo.
pause
exit /b 1

:RUN
echo Launching Master Inspection Plan web app...
"%PY_CMD%" app.py inspection_plan.db
pause
