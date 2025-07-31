@echo off
echo ========================================
echo Freshservice Printer Asset Integration
echo ========================================
echo.

echo This script will scan for printers and add them to Freshservice.
echo Make sure you have configured freshservice_config.json first.
echo.

REM Check if config file exists
if not exist "freshservice_config.json" (
    echo ERROR: freshservice_config.json not found!
    echo Please configure your Freshservice settings first.
    echo.
    pause
    exit /b 1
)

echo Starting printer discovery and Freshservice integration...
echo.

REM Execute PowerShell script
powershell -ExecutionPolicy Bypass -File "freshservice_printer_asset_simple.ps1"

echo.
echo Integration completed!
pause 