@echo off
title Oceaneering QC Comparison Tool
color 0B

echo ================================================
echo    Oceaneering QC Comparison Tool - Launcher
echo ================================================
echo.

:: ---- Check Python is installed ----
python --version >nul 2>&1
if %errorlevel% neq 0 (
    color 0C
    echo  ERROR: Python was not found on this machine.
    echo.
    echo  Please download and install Python 3.x from:
    echo     https://www.python.org/downloads/
    echo.
    echo  Make sure to check "Add Python to PATH" during installation.
    echo.
    pause
    exit /b 1
)

:: ---- Set working directory to script location ----
cd /d "%~dp0"

:: ---- Create virtual environment on first run ----
if not exist ".venv" (
    echo  First-time setup: creating virtual environment...
    python -m venv .venv
    if %errorlevel% neq 0 (
        color 0C
        echo  ERROR: Failed to create virtual environment.
        pause
        exit /b 1
    )
    echo  Done.
    echo.
)

:: ---- Activate virtual environment ----
call "%~dp0.venv\Scripts\activate.bat"

:: ---- Install / update dependencies on first run ----
if not exist ".venv\installed.flag" (
    echo  Installing required packages (this only happens once)...
    echo  Please wait...
    echo.
    pip install --quiet -r requirements.txt
    if %errorlevel% neq 0 (
        color 0C
        echo  ERROR: Failed to install packages.
        echo  Check your internet connection and try again.
        pause
        exit /b 1
    )
    echo. > .venv\installed.flag
    echo  Packages installed successfully.
    echo.
)

:: ---- Launch the app ----
echo  Starting QC Tool... your browser will open automatically.
echo  To stop the tool, close this window or press Ctrl+C.
echo.

python -m streamlit run app.py --server.headless false

:: ---- If streamlit exits unexpectedly ----
echo.
echo  The application has closed.
pause
