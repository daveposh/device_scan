@echo off
REM USB Printer Scanner Batch Wrapper
REM Designed for PDQ deployment
REM This script runs the PowerShell printer scanner with proper execution policy

echo ========================================
echo USB Printer Scanner for Windows
echo ========================================
echo.

REM Set execution policy for this session
powershell -ExecutionPolicy Bypass -Command "Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process -Force"

REM Run the enhanced printer scanner
echo Running printer scan...
powershell -ExecutionPolicy Bypass -File "printer_scanner_enhanced.ps1" -ExportCSV

echo.
echo ========================================
echo Scan completed. Check the output files:
echo - printer_scan_results.txt (full detailed log)
echo - printer_devices_only.txt (printer devices only)
echo - printer_scan_results.csv (if CSV export was enabled)
echo ========================================

pause 