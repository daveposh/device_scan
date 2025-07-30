# USB Printer Scanner for Windows

A native Windows tool to scan for USB printers and extract model and serial number information. Designed for PDQ deployment and enterprise environments.

## Features

- **Native Windows Tools**: Uses only built-in Windows PowerShell cmdlets and WMI
- **Multiple Detection Methods**: 
  - Windows Management Instrumentation (WMI)
  - Plug and Play (PnP) device enumeration
  - Registry scanning
  - Print Management API
- **Serial Number Extraction**: Attempts to extract USB device serial numbers from device IDs
- **PDQ Ready**: Designed for deployment via PDQ Deploy
- **Multiple Output Formats**: Text log and CSV export options
- **Error Handling**: Robust error handling for enterprise environments

## Files Included

- `printer_scanner.ps1` - Basic PowerShell script for USB printer scanning
- `printer_scanner_enhanced.ps1` - Enhanced version with multiple detection methods
- `scan_printers.bat` - Batch wrapper for easy PDQ deployment
- `requirements.txt` - Python dependencies (not needed for native Windows tools)

## PDQ Deployment Instructions

### Method 1: PowerShell Script Direct Deployment

1. **Upload Files**: Upload `printer_scanner_enhanced.ps1` to your PDQ repository
2. **Create Package**: Create a new package in PDQ Deploy
3. **Command Line**:
   ```
   powershell.exe -ExecutionPolicy Bypass -File "printer_scanner_enhanced.ps1" -ExportCSV
   ```
4. **Deploy**: Deploy to target computers

### Method 2: Batch File Deployment

1. **Upload Files**: Upload both `scan_printers.bat` and `printer_scanner_enhanced.ps1` to your PDQ repository
2. **Create Package**: Create a new package in PDQ Deploy
3. **Command Line**:
   ```
   scan_printers.bat
   ```
4. **Deploy**: Deploy to target computers

### Method 3: Custom Output Path

For centralized logging, specify a custom output path:

```
powershell.exe -ExecutionPolicy Bypass -File "printer_scanner_enhanced.ps1" -OutputPath "C:\Temp\printer_scan_%COMPUTERNAME%.txt" -ExportCSV
```

## Usage Examples

### Basic Scan
```powershell
.\printer_scanner_enhanced.ps1
```

### Scan with CSV Export
```powershell
.\printer_scanner_enhanced.ps1 -ExportCSV
```

### Scan with Custom Output Path
```powershell
.\printer_scanner_enhanced.ps1 -OutputPath "C:\Logs\printer_scan.txt" -ExportCSV
```

### Verbose Output
```powershell
.\printer_scanner_enhanced.ps1 -Verbose -ExportCSV
```

## Output Information

The scanner attempts to extract the following information for each USB printer:

- **Printer Name**: Display name of the printer
- **Model**: Printer model information
- **Serial Number**: USB device serial number (when available)
- **Manufacturer**: Manufacturer information (VID/PID when available)
- **Port Name**: USB port name or COM port
- **Driver Name**: Installed printer driver
- **Location**: Physical location (if configured)
- **Status**: Current printer status
- **Source**: Detection method used
- **Associated COM Port**: For Epson USB controllers, the virtual COM port created
- **Associated Printer**: For Epson USB controllers, the printer queue that uses the COM port

## Detection Methods

### 1. WMI (Win32_Printer)
- Scans Windows printer objects
- Identifies USB-connected printers by port name
- Extracts basic printer information

### 2. PnP Device Enumeration
- Scans USB devices using Get-PnpDevice
- Identifies printer devices by manufacturer names
- Extracts serial numbers from USB device IDs

### 3. Registry Scanning
- Scans Windows printer registry keys
- Identifies USB printer ports
- Provides additional printer configuration details

### 4. Print Management API
- Uses Get-Printer cmdlet
- Provides current printer status
- Extracts printer properties

### 5. Device Manager (Enhanced)
- Scans all device categories including "Other devices"
- Detects USB-to-serial bridge devices
- Identifies COM port printers

### 6. Epson USB Controller Detection
- Detects Epson USB controllers (TM/BA/EU series)
- Identifies associated COM port devices
- Links USB controllers to printer queues
- Extracts serial numbers from USB device IDs

## Serial Number Extraction

The scanner attempts to extract serial numbers using these methods:

1. **USB Device ID Parsing**: Extracts serial numbers from USB device instance IDs
2. **WMI Properties**: Searches WMI device properties for serial numbers
3. **Registry Values**: Checks registry for stored serial number information
4. **Hex Decoding**: Decodes hex-encoded serial numbers to human-readable format

### Hex Serial Number Decoding

The scanner includes enhanced hex decoding for serial numbers:

- **Standard ASCII Decoding**: Converts hex to ASCII characters
- **Epson-Specific Decoding**: Handles Epson's hex encoding format
  - Example: `5839584C1156290000` → `X9XL115629`
  - Filters out zero-padding bytes
  - Validates decoded results

Format: `USB\VID_XXXX&PID_XXXX\SERIALNUMBER`

## Requirements

- **Windows PowerShell 3.0 or later**
- **Administrator privileges** (recommended for full access)
- **Windows 7/8/10/11** (tested on Windows 10/11)

## Troubleshooting

### Common Issues

1. **No printers found**
   - Ensure USB printers are connected and powered on
   - Check that printer drivers are installed
   - Run as administrator for full access

2. **Execution policy errors**
   - Use the batch wrapper (`scan_printers.bat`)
   - Or run with `-ExecutionPolicy Bypass`

3. **Missing serial numbers**
   - Not all USB devices expose serial numbers
   - Some printers may not report serial numbers via USB
   - Check device manager for additional information

### Log Files

The scanner creates detailed log files:
- `printer_scan_results.txt` - Detailed scan log
- `printer_scan_results.csv` - CSV export (when enabled)

## PDQ Integration

### Return Codes
- **0**: Success (printers found or not found)
- **1**: Error during execution

### Output Variables
The script outputs standardized text for PDQ parsing:
- `USB_PRINTERS_FOUND: X` - Number of printers found
- `PRINTER: Name | Model: Model | Serial: SerialNumber` - Per printer details

### Example PDQ Collection Query
```sql
-- Find computers with USB printers
SELECT ComputerName, PrinterName, Model, SerialNumber 
FROM CustomInventory 
WHERE CollectionName = 'USB Printer Scan'
```

## Security Considerations

- Script uses only native Windows PowerShell cmdlets
- No external dependencies or downloads
- Execution policy bypass is scoped to the process only
- No network communication or data transmission

## Version History

- **v1.0**: Initial release with basic WMI scanning
- **v1.1**: Enhanced version with multiple detection methods
- **v1.2**: Added CSV export and PDQ integration features

## Support

For issues or questions:
1. Check the log files for detailed error information
2. Ensure all files are deployed together
3. Verify PowerShell execution policy settings
4. Test on a single machine before mass deployment 