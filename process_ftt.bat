@echo off
title FTT B/C Report Generator
color 0B
setlocal
cd /d "%~dp0"

echo ===============================================================================
echo     FTT B/C GRADE REPORT GENERATOR (process_ftt.py)
echo ===============================================================================
echo.
echo  [+] Checking Python environment...
python --version >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo  [!] Python is not found in PATH. Please install Python 3.9+ first.
    echo.
    pause
    exit /b 1
)

:: Ensure required packages
echo  [+] Verifying required Python packages...
python -c "import pandas, openpyxl, numpy" >NUL 2>&1
if %errorlevel% neq 0 (
    echo  [!] Installing missing packages: pandas, openpyxl, numpy...
    python -m pip install pandas openpyxl numpy
    echo.
)

:: Check input files exist
if not exist "FTT_Input" (
    echo  [!] Folder "FTT_Input" does not exist. Create it and drop FTT export files inside.
    echo.
    pause
    exit /b 1
)

dir /b "FTT_Input\*.xls*" >NUL 2>&1
if %ERRORLEVEL% NEQ 0 (
    dir /b "FTT_Input\*.csv" >NUL 2>&1
    if %ERRORLEVEL% NEQ 0 (
        echo  [!] No FTT files found in "FTT_Input". Drop your FTT export files there first.
        echo.
        pause
        exit /b 1
    )
)

if not exist "Output" mkdir "Output"

echo  [+] Running FTT processor (this may take a minute for large files)...
echo.
python process_ftt.py --ftt-dir "FTT_Input" --output "Output\FTT_Combined_Report.xlsx" --color-map "BC_Color\Color_Defect.xlsx"

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo  [!] Report generation FAILED with code %ERRORLEVEL%. See errors above.
    echo.
    pause
    exit /b %ERRORLEVEL%
)

echo.
echo  [+] Done! Report saved to: Output\FTT_Combined_Report.xlsx
echo.
choice /c YN /m "Open output folder now"
if %ERRORLEVEL% EQU 1 start "" "Output"
