@echo off
REM ============================================================
REM  inforcer SOP Generator - One-time Setup
REM  Run this ONCE. It will:
REM    1. Create a Python virtual environment
REM    2. Install required packages
REM    3. Prompt for your API key and save it permanently
REM ============================================================

setlocal

echo.
echo ============================================================
echo   inforcer SOP Generator - Setup
echo ============================================================
echo.

REM --- Move into the folder this script lives in ---
cd /d "%~dp0"

REM --- Check Python is available ---
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python is not installed or not on PATH.
    echo.
    echo Install Python from the Microsoft Store ^(search "Python 3.12"^)
    echo or from https://www.python.org/downloads/
    echo.
    pause
    exit /b 1
)

echo [1/3] Creating virtual environment...
if not exist ".venv\Scripts\activate.bat" (
    python -m venv .venv
    if errorlevel 1 (
        echo ERROR: failed to create virtual environment.
        pause
        exit /b 1
    )
) else (
    echo       .venv already exists, skipping.
)

echo.
echo [2/3] Installing required packages...
call ".venv\Scripts\activate.bat"
python -m pip install --upgrade pip >nul
pip install -r requirements.txt
if errorlevel 1 (
    echo ERROR: failed to install packages.
    pause
    exit /b 1
)

echo.
echo [3/3] Saving your inforcer API key...
echo.
echo   Paste your API key below ^(it will be saved permanently
echo   as a Windows user environment variable named
echo   INFORCER_API_KEY^).
echo.
set /p APIKEY="   API key: "

if "%APIKEY%"=="" (
    echo.
    echo ERROR: no key entered. Re-run setup.bat to try again.
    pause
    exit /b 1
)

setx INFORCER_API_KEY "%APIKEY%" >nul
if errorlevel 1 (
    echo ERROR: failed to save the API key.
    pause
    exit /b 1
)

echo.
echo ============================================================
echo   Setup complete.
echo ============================================================
echo.
echo   Your API key is saved permanently for your Windows user.
echo   You can now double-click run.bat any time to generate SOPs.
echo.
echo   NOTE: If run.bat is already open, close and reopen it
echo         so it picks up the newly-saved API key.
echo.
pause
endlocal
