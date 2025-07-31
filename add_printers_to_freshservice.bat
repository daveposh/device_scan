@echo off
echo ========================================
echo Freshservice Printer Asset Integration
echo ========================================
echo.

REM Check if parameters are provided
if "%1"=="" (
    echo Usage: add_printers_to_freshservice.bat [Domain] [ApiKey] [AssetType] [Location] [Department]
    echo.
    echo Parameters:
    echo   Domain     - Your Freshservice domain (e.g., yourcompany)
    echo   ApiKey     - Your Freshservice API key
    echo   AssetType  - Asset type name (optional, default: Printer)
    echo   Location   - Location name (optional)
    echo   Department - Department name (optional)
    echo.
    echo Examples:
    echo   add_printers_to_freshservice.bat yourcompany your-api-key
    echo   add_printers_to_freshservice.bat yourcompany your-api-key "Printer" "Main Office" "IT"
    echo.
    pause
    exit /b 1
)

REM Set parameters
set DOMAIN=%1
set APIKEY=%2
set ASSETTYPE=%3
set LOCATION=%4
set DEPARTMENT=%5

REM Build PowerShell command
set PS_CMD=powershell -ExecutionPolicy Bypass -File "freshservice_printer_asset.ps1" -FreshserviceDomain "%DOMAIN%" -ApiKey "%APIKEY%"

if not "%ASSETTYPE%"=="" set PS_CMD=%PS_CMD% -AssetType "%ASSETTYPE%"
if not "%LOCATION%"=="" set PS_CMD=%PS_CMD% -Location "%LOCATION%"
if not "%DEPARTMENT%"=="" set PS_CMD=%PS_CMD% -Department "%DEPARTMENT%"

echo Starting Freshservice integration...
echo Domain: %DOMAIN%
echo Asset Type: %ASSETTYPE%
echo Location: %LOCATION%
echo Department: %DEPARTMENT%
echo.

REM Execute PowerShell script
%PS_CMD%

echo.
echo Integration completed!
pause 