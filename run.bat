@echo off
setlocal

cd /d "%~dp0"

if not exist ".venv\Scripts\python.exe" (
    echo.
    echo ERROR: virtual environment not found.
    echo Run setup.bat first to install dependencies.
    echo.
    pause
    exit /b 1
)

if "%INFORCER_API_KEY%"=="" (
    echo.
    echo WARNING: INFORCER_API_KEY is not set in this session.
    echo If setup.bat was run in another window, close this window and
    echo re-open run.bat so the key is picked up.
    echo.
)

".venv\Scripts\python.exe" "inforcer_sop_generator.py"
set EXITCODE=%ERRORLEVEL%

if exist "output" (
    start "" "output"
)

echo.
pause
exit /b %EXITCODE%
