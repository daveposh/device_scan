@echo off
echo ========================================
echo Freshservice API Integration Test
echo ========================================
echo.

echo This script will test the Freshservice API connection and asset creation.
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

echo Starting Freshservice API test...
echo.

REM Execute PowerShell script
powershell -ExecutionPolicy Bypass -File "test_freshservice_api.ps1"

echo.
echo Test completed!
pause 