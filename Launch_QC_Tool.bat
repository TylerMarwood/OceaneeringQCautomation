@echo off
title Oceaneering QC Comparison Tool
color 0B

echo ================================================
echo    Oceaneering QC Comparison Tool - Launcher
echo ================================================
echo.

:: ---- Check Python is installed ----
echo  [1/5] Checking Python installation...
python --version
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
echo  [1/5] Python OK.
echo.

:: ---- Set working directory to script location ----
echo  [2/5] Setting working directory...
echo  Script location: %~dp0
cd /d "%~dp0"
echo  Current directory: %CD%
echo  [2/5] Working directory OK.
echo.

:: ---- Create virtual environment on first run ----
echo  [3/5] Checking virtual environment...
if not exist ".venv" (
    echo  Creating virtual environment for the first time...
    python -m venv .venv
    if %errorlevel% neq 0 (
        color 0C
        echo  ERROR: Failed to create virtual environment. Error code: %errorlevel%
        pause
        exit /b 1
    )
    echo  Virtual environment created successfully.
) else (
    echo  Virtual environment already exists, skipping.
)
echo  [3/5] Virtual environment OK.
echo.

:: ---- Activate virtual environment ----
echo  [4/5] Activating virtual environment...
echo  Looking for: %~dp0.venv\Scripts\activate.bat
if not exist "%~dp0.venv\Scripts\activate.bat" (
    color 0C
    echo  ERROR: activate.bat not found. The virtual environment may be corrupted.
    echo  Delete the .venv folder and try again.
    pause
    exit /b 1
)
call "%~dp0.venv\Scripts\activate.bat"
if %errorlevel% neq 0 (
    color 0C
    echo  ERROR: Failed to activate virtual environment. Error code: %errorlevel%
    pause
    exit /b 1
)
echo  [4/5] Virtual environment activated OK.
echo.

:: ---- Install / update dependencies on first run ----
echo  [5/5] Checking dependencies...
if not exist ".venv\installed.flag" (
    echo  Installing required packages (this only happens once)...
    echo  Please wait...
    echo.
    pip install -r requirements.txt
    if %errorlevel% neq 0 (
        color 0C
        echo  ERROR: Failed to install packages. Error code: %errorlevel%
        echo  Check your internet connection and try again.
        pause
        exit /b 1
    )
    echo. > .venv\installed.flag
    echo  Packages installed successfully.
) else (
    echo  Dependencies already installed, skipping.
)
echo  [5/5] Dependencies OK.
echo.

:: ---- Launch the app ----
echo  All checks passed. Launching QC Tool...
echo  Your browser will open automatically to http://localhost:8501
echo  To stop the tool, close this window or press Ctrl+C.
echo.

python -m streamlit run app.py --server.headless false

:: ---- If streamlit exits unexpectedly ----
echo.
echo  The application has closed.
pause
