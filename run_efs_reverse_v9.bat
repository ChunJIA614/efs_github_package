@echo off
setlocal EnableExtensions
:: Use pushd instead of cd /d so running from UNC network shares (\\server\Public\...) never fails!
pushd "%~dp0"

title Malaysia Court E-Filing (EFS) Scraper - Final Edition

echo ============================================================================
echo   MALAYSIA COURT E-FILING (EFS) SCRAPER - FINAL EDITION
echo ============================================================================
echo   Working Folder : %CD%
echo   Output Folder  : %CD%\efs_output (.xlsx only)
echo ============================================================================
echo.

:: Detect working Python 3 interpreter (handles both 'python' and 'py -3' and avoids Microsoft Store stub)
set "PY_CMD="
python -c "import sys; assert sys.version_info >= (3, 8)" >nul 2>nul
if not errorlevel 1 (
    set "PY_CMD=python"
) else (
    py -3 -c "import sys; assert sys.version_info >= (3, 8)" >nul 2>nul
    if not errorlevel 1 (
        set "PY_CMD=py -3"
    )
)

if "%PY_CMD%"=="" (
    echo [ERROR] Python 3.8+ was not found on this computer!
    echo.
    echo Reason for crash on personal PC:
    echo   1. Python is not installed on this PC yet, OR
    echo   2. "Add python.exe to PATH" was not checked during installation.
    echo.
    echo How to fix:
    echo   Download and install Python 3 from https://www.python.org/downloads/
    echo   ^(Make sure to check the box "Add python.exe to PATH" at the bottom of the installer!^)
    echo.
    pause
    popd
    exit /b 1
)

echo [1/2] Verifying required Python packages using: %PY_CMD%
%PY_CMD% -c "import playwright, bs4, openpyxl" >nul 2>nul
if errorlevel 1 (
    echo       Installing required packages on this PC ^(playwright, beautifulsoup4, openpyxl^)...
    %PY_CMD% -m pip install --upgrade pip
    %PY_CMD% -m pip install playwright beautifulsoup4 openpyxl
    if errorlevel 1 (
        echo.
        echo [ERROR] Package installation failed. Please check your internet connection.
        pause
        popd
        exit /b 1
    )
)
echo       Packages verified successfully.
echo.
echo [2/2] Launching EFS Scraper Application...
echo       ^(Keep this window open while the application is running.^)
echo.
%PY_CMD% "%~dp0efs_scraper_reverse_v9.py"

set ERR=%ERRORLEVEL%
echo.
echo ============================================================================
echo   Program finished ^(Exit Code: %ERR%^)
echo ============================================================================
pause
popd
exit /b %ERR%
