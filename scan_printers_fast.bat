@echo off
REM Fast USB Printer Scanner Batch Wrapper
REM Designed for PDQ deployment - Quick scan for actual printers only
REM This script runs the PowerShell printer scanner with proper execution policy

echo ========================================
echo Fast USB Printer Scanner for Windows
echo ========================================
echo.

REM Set execution policy for this session
powershell -ExecutionPolicy Bypass -Command "Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process -Force"

REM Run the fast printer scanner
echo Running fast printer scan...
powershell -ExecutionPolicy Bypass -File "printer_scanner_fast.ps1"

echo.
echo ========================================
echo Fast scan completed. Check the output file:
echo - printer_scan_fast.txt
echo ========================================

pause 