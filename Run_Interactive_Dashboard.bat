@echo off
title COPQ Interactive Dashboard
color 0B
setlocal
cd /d "%~dp0"

echo ===============================================================================
echo     COPQ INTERACTIVE DASHBOARD - FILTERS, CHECKBOXES ^& LIVE CHARTS
echo ===============================================================================
echo.
echo  [+] Checking Python environment...
python --version >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo  [!] Python is not found in PATH. Please install Python 3.9+ first.
    echo.
    exit /b 1
)

:: Ensure required packages
echo  [+] Verifying required Python packages...
python -c "import pandas, openpyxl, numpy, matplotlib, pptx" >NUL 2>&1
if %errorlevel% neq 0 (
    echo  [!] Installing missing packages: pandas, openpyxl, numpy, matplotlib, python-pptx...
    python -m pip install pandas openpyxl numpy matplotlib python-pptx
    echo.
)

echo  [+] Starting Interactive Dashboard...
echo.

python Interactive_Dashboard.py

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo  [!] Application exited with code %ERRORLEVEL%.
)

exit /b %ERRORLEVEL%
