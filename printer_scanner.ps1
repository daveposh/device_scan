# USB Printer Scanner for Windows
# This script scans for USB printers and extracts model and serial number information
# Designed for PDQ deployment

param(
    [string]$OutputPath = ".\printer_scan_results.txt",
    [switch]$Verbose
)

# Function to write log messages
function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Level] $Message"
    Write-Host $logMessage
    Add-Content -Path $OutputPath -Value $logMessage
}

# Function to get USB printer information using WMI
function Get-USBPrinterInfo {
    Write-Log "Starting USB printer scan..."
    
    try {
        # Get all USB devices
        $usbDevices = Get-WmiObject -Class Win32_USBHub -ErrorAction SilentlyContinue
        
        # Get printer devices
        $printers = Get-WmiObject -Class Win32_Printer -ErrorAction SilentlyContinue
        
        # Get USB controller devices
        $usbControllers = Get-WmiObject -Class Win32_USBController -ErrorAction SilentlyContinue
        
        # Get PnP devices that might be printers
        $pnpDevices = Get-WmiObject -Class Win32_PnPEntity -ErrorAction SilentlyContinue | 
                      Where-Object { $_.Name -like "*printer*" -or $_.Name -like "*print*" }
        
        $printerResults = @()
        
        # Method 1: Check printers that are connected via USB
        foreach ($printer in $printers) {
            if ($printer.PortName -like "*USB*" -or $printer.PortName -like "*USBPRN*") {
                $printerInfo = [PSCustomObject]@{
                    Name = $printer.Name
                    PortName = $printer.PortName
                    DriverName = $printer.DriverName
                    Location = $printer.Location
                    Comment = $printer.Comment
                    DeviceID = $printer.DeviceID
                    Source = "Win32_Printer"
                }
                $printerResults += $printerInfo
            }
        }
        
        # Method 2: Check PnP devices for printers
        foreach ($device in $pnpDevices) {
            if ($device.Name -like "*USB*" -or $device.DeviceID -like "*USB*") {
                $printerInfo = [PSCustomObject]@{
                    Name = $device.Name
                    DeviceID = $device.DeviceID
                    Manufacturer = $device.Manufacturer
                    Description = $device.Description
                    Status = $device.Status
                    Source = "Win32_PnPEntity"
                }
                $printerResults += $printerInfo
            }
        }
        
        # Method 3: Use registry to find USB printers
        $registryPrinters = Get-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\Print\Printers" -ErrorAction SilentlyContinue
        foreach ($regPrinter in $registryPrinters) {
            if ($regPrinter.PSPath -and (Test-Path "$($regPrinter.PSPath)\Ports")) {
                $portInfo = Get-ItemProperty -Path "$($regPrinter.PSPath)\Ports" -ErrorAction SilentlyContinue
                if ($portInfo -and $portInfo.PortName -like "*USB*") {
                    $printerInfo = [PSCustomObject]@{
                        Name = $regPrinter.PSChildName
                        PortName = $portInfo.PortName
                        Source = "Registry"
                    }
                    $printerResults += $printerInfo
                }
            }
        }
        
        return $printerResults
        
    } catch {
        Write-Log "Error scanning USB printers: $($_.Exception.Message)" "ERROR"
        return @()
    }
}

# Function to get detailed printer information including serial number
function Get-PrinterDetails {
    param([array]$Printers)
    
    Write-Log "Getting detailed printer information..."
    
    $detailedResults = @()
    
    foreach ($printer in $Printers) {
        try {
            $details = [PSCustomObject]@{
                Name = $printer.Name
                Model = $null
                SerialNumber = $null
                Manufacturer = $null
                PortName = $printer.PortName
                DriverName = $printer.DriverName
                Location = $printer.Location
                Status = $printer.Status
                Source = $printer.Source
                DeviceID = $printer.DeviceID
            }
            
            # Try to get model from device name
            if ($printer.Name) {
                $details.Model = $printer.Name
            }
            
            # Try to get serial number from device ID or other properties
            if ($printer.DeviceID) {
                # Extract serial number from device ID if present
                if ($printer.DeviceID -match "USB\\VID_([A-F0-9]{4})&PID_([A-F0-9]{4})\\([A-F0-9]+)") {
                    $details.SerialNumber = $matches[3]
                    $details.Manufacturer = "VID: $($matches[1]), PID: $($matches[2])"
                }
            }
            
            # Try to get more information using WMI
            try {
                $wmiPrinter = Get-WmiObject -Class Win32_Printer -Filter "Name='$($printer.Name)'" -ErrorAction SilentlyContinue
                if ($wmiPrinter) {
                    $details.Model = $wmiPrinter.Name
                    $details.Location = $wmiPrinter.Location
                    $details.Status = $wmiPrinter.Status
                }
            } catch {
                # Ignore WMI errors for individual printers
            }
            
            $detailedResults += $details
            
        } catch {
            Write-Log "Error getting details for printer $($printer.Name): $($_.Exception.Message)" "ERROR"
        }
    }
    
    return $detailedResults
}

# Function to get USB device serial numbers using devcon
function Get-USBDeviceSerialNumbers {
    Write-Log "Attempting to get USB device serial numbers..."
    
    $serialNumbers = @()
    
    try {
        # Try using PowerShell to get USB device information
        $usbDevices = Get-PnpDevice -Class USB -ErrorAction SilentlyContinue
        
        foreach ($device in $usbDevices) {
            $deviceInfo = [PSCustomObject]@{
                Name = $device.FriendlyName
                InstanceId = $device.InstanceId
                Status = $device.Status
                Class = $device.Class
            }
            
            # Extract serial number from instance ID if possible
            if ($device.InstanceId -match "USB\\VID_([A-F0-9]{4})&PID_([A-F0-9]{4})\\([A-F0-9]+)") {
                $deviceInfo.SerialNumber = $matches[3]
                $deviceInfo.VID = $matches[1]
                $deviceInfo.PID = $matches[2]
            }
            
            $serialNumbers += $deviceInfo
        }
        
    } catch {
        Write-Log "Error getting USB device serial numbers: $($_.Exception.Message)" "ERROR"
    }
    
    return $serialNumbers
}

# Main execution
Write-Log "=== USB Printer Scanner Started ==="
Write-Log "Computer Name: $env:COMPUTERNAME"
Write-Log "User: $env:USERNAME"
Write-Log "Date: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"

# Get USB printer information
$usbPrinters = Get-USBPrinterInfo

if ($usbPrinters.Count -eq 0) {
    Write-Log "No USB printers found on this system." "WARNING"
} else {
    Write-Log "Found $($usbPrinters.Count) potential USB printer(s)"
    
    # Get detailed information
    $detailedPrinters = Get-PrinterDetails -Printers $usbPrinters
    
    # Get USB device serial numbers
    $usbDevices = Get-USBDeviceSerialNumbers
    
    # Display results
    Write-Log "=== USB Printer Scan Results ==="
    
    foreach ($printer in $detailedPrinters) {
        Write-Log "Printer: $($printer.Name)"
        Write-Log "  Model: $($printer.Model)"
        Write-Log "  Serial Number: $($printer.SerialNumber)"
        Write-Log "  Manufacturer: $($printer.Manufacturer)"
        Write-Log "  Port: $($printer.PortName)"
        Write-Log "  Driver: $($printer.DriverName)"
        Write-Log "  Location: $($printer.Location)"
        Write-Log "  Status: $($printer.Status)"
        Write-Log "  Source: $($printer.Source)"
        Write-Log "  Device ID: $($printer.DeviceID)"
        Write-Log "---"
    }
    
    # Also show all USB devices for reference
    Write-Log "=== All USB Devices ==="
    foreach ($device in $usbDevices) {
        Write-Log "Device: $($device.Name)"
        Write-Log "  Serial: $($device.SerialNumber)"
        Write-Log "  VID: $($device.VID)"
        Write-Log "  PID: $($device.PID)"
        Write-Log "  Status: $($device.Status)"
        Write-Log "---"
    }
}

Write-Log "=== USB Printer Scanner Completed ==="
Write-Log "Results saved to: $OutputPath"

# Return results for PDQ
if ($usbPrinters.Count -gt 0) {
    Write-Output "USB_PRINTERS_FOUND: $($usbPrinters.Count)"
    foreach ($printer in $detailedPrinters) {
        Write-Output "PRINTER: $($printer.Name) | Model: $($printer.Model) | Serial: $($printer.SerialNumber)"
    }
} else {
    Write-Output "USB_PRINTERS_FOUND: 0"
} 