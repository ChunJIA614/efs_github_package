@echo off
setlocal EnableExtensions
cd /d "%~dp0"

title Malaysia Court E-Filing (EFS) Scraper - Final Edition

echo ============================================================================
echo   MALAYSIA COURT E-FILING (EFS) SCRAPER - FINAL EDITION
echo ============================================================================
echo   Working Folder : %CD%
echo   Output Folder  : %CD%\efs_output (.xlsx only)
echo ============================================================================
echo.

where python >nul 2>nul
if errorlevel 1 (
    echo [ERROR] Python was not found in your system PATH.
    echo Please install Python 3 and ensure "Add Python to PATH" is checked.
    echo.
    pause
    exit /b 1
)

echo [1/2] Verifying required Python packages...
python -c "import playwright, bs4, openpyxl" 2>nul
if errorlevel 1 (
    echo Installing required packages (playwright, beautifulsoup4, openpyxl)...
    python -m pip install playwright beautifulsoup4 openpyxl
    if errorlevel 1 (
        echo.
        echo [ERROR] Package installation failed.
        pause
        exit /b 1
    )
)
echo       Packages verified successfully.
echo.
echo [2/2] Launching EFS Scraper Application...
echo       (If using the Desktop UI window, you can control everything from there.)
echo.
python efs_scraper_reverse_v9.py

set ERR=%ERRORLEVEL%
if not "%ERR%"=="0" (
    echo.
    echo ============================================================================
    echo   Program exited with code %ERR%
    echo ============================================================================
    pause
)
exit /b %ERR%
