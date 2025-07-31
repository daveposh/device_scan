@echo off
echo ========================================
echo Comprehensive Freshservice Integration Test
echo ========================================
echo.

echo This script will test all components of the Freshservice integration.
echo.

echo Starting comprehensive test...
echo.

REM Execute PowerShell script
powershell -ExecutionPolicy Bypass -File "test_all_integration.ps1"

echo.
echo Test completed!
pause 