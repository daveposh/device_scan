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

# Function to decode hex serial numbers (including Epson format)
function Decode-HexSerial {
    param([string]$HexSerial)
    
    try {
        if ($HexSerial.Length -eq 20 -and $HexSerial -match "^([A-F0-9]{8})([0-9]{6})(0{4})$") {
            $hexPart = $matches[1]  # First 8 hex characters
            $numberPart = $matches[2]  # Next 6 numbers
            $padding = $matches[3]  # Last 4 zeros
            
            # Decode the hex part
            $hexDecoded = ""
            for ($i = 0; $i -lt $hexPart.Length; $i += 2) {
                $hexPair = $hexPart.Substring($i, 2)
                $byteValue = [Convert]::ToByte($hexPair, 16)
                if ($byteValue -ge 32 -and $byteValue -le 126) {
                    $hexDecoded += [char]$byteValue
                }
            }
            
            # Combine hex decoded part with number part
            $epsonDecoded = $hexDecoded + $numberPart
            
            if ($epsonDecoded.Length -gt 0) {
                return "$HexSerial (Epson Decoded: $epsonDecoded)"
            }
        }
        
        # If no Epson pattern found, return original
        return $HexSerial
    } catch {
        return $HexSerial
    }
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
                      Where-Object { 
                          # Exclude non-printer devices
                          $_.Name -notlike "*Fingerprint*" -and
                          $_.Name -notlike "*Scanner*" -and
                          $_.Name -notlike "*Camera*" -and
                          $_.Name -notlike "*Webcam*" -and
                          $_.Name -notlike "*Microphone*" -and
                          $_.Name -notlike "*Audio*" -and
                          $_.Name -notlike "*Speaker*" -and
                          $_.Name -notlike "*Headset*" -and
                          $_.Name -notlike "*Mouse*" -and
                          $_.Name -notlike "*Keyboard*" -and
                          $_.Name -notlike "*Touchpad*" -and
                          $_.Name -notlike "*Trackpad*" -and
                          $_.Name -notlike "*Monitor*" -and
                          $_.Name -notlike "*Display*" -and
                          $_.Name -notlike "*Graphics*" -and
                          $_.Name -notlike "*Video*" -and
                          $_.Name -notlike "*Network*" -and
                          $_.Name -notlike "*Ethernet*" -and
                          $_.Name -notlike "*WiFi*" -and
                          $_.Name -notlike "*Wireless*" -and
                          $_.Name -notlike "*Bluetooth*" -and
                          $_.Name -notlike "*Card Reader*" -and
                          $_.Name -notlike "*Smart Card*" -and
                          $_.Name -notlike "*USB Hub*" -and
                          $_.Name -notlike "*USB Root*" -and
                          $_.Name -notlike "*Composite*" -and
                          $_.Name -notlike "*Bus Enumerator*" -and
                          $_.Name -notlike "*Microsoft*" -and
                          $_.Name -notlike "*OneNote*" -and
                          $_.Name -notlike "*Fax*" -and
                          $_.Name -notlike "*XPS*" -and
                                                     $_.Name -notlike "*PDF*" -and
                           $_.Name -notlike "*Dock*" -and
                           $_.Name -notlike "*Dell Dock*" -and
                           $_.Name -notlike "*WD19S*" -and
                           $_.Name -notlike "*Docking*" -and
                           $_.Name -notlike "*Port Replicator*" -and
                          # Include printer devices
                          ($_.Name -like "*printer*" -or 
                           $_.Name -like "*print*" -or
                           $_.Name -like "*HP*" -or
                           $_.Name -like "*Canon*" -or
                           $_.Name -like "*Epson*" -or
                           $_.Name -like "*TM*" -or
                           $_.Name -like "*Thermal*" -or
                           $_.Name -like "*Receipt*" -or
                           $_.Name -like "*POS*" -or
                           $_.Name -like "*Point of Sale*" -or
                           $_.Name -like "*USB-to-Serial*" -or
                           $_.Name -like "*USB Serial*" -or
                           $_.Name -like "*USB Controller*" -or
                           $_.Name -like "*Serial Port*" -or
                           $_.Name -like "*COM*" -or
                           $_.Name -like "*Brother*" -or
                           $_.Name -like "*Lexmark*" -or
                           $_.Name -like "*Dell*" -and $_.Name -notlike "*Dock*" -and $_.Name -notlike "*WD19S*" -or
                           $_.Name -like "*Xerox*" -or
                           $_.Name -like "*Samsung*" -or
                           $_.Name -like "*Ricoh*" -or
                           $_.Name -like "*Kyocera*" -or
                           $_.Name -like "*Sharp*" -or
                           $_.Name -like "*Konica*" -or
                           $_.Name -like "*Minolta*" -or
                           $_.Name -like "*OKI*" -or
                           $_.Name -like "*Toshiba*" -or
                           $_.Name -like "*Panasonic*" -or
                           $_.Name -like "*Fuji*")
                      }
        
        $printerResults = @()
        
        # Method 1: Check printers that are connected via USB
        foreach ($printer in $printers) {
            # Exclude Microsoft services
            if ($printer.Name -notlike "*Microsoft*" -and
                $printer.Name -notlike "*OneNote*" -and
                $printer.Name -notlike "*Fax*" -and
                $printer.Name -notlike "*XPS*" -and
                $printer.Name -notlike "*PDF*" -and
                ($printer.PortName -like "*USB*" -or $printer.PortName -like "*USBPRN*")) {
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
                    $details.SerialNumber = Decode-HexSerial -HexSerial $matches[3]
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
        $usbDevices = Get-PnpDevice -Class USB -ErrorAction SilentlyContinue | 
                      Where-Object { 
                          # Exclude non-printer devices
                          $_.FriendlyName -notlike "*Fingerprint*" -and
                          $_.FriendlyName -notlike "*Scanner*" -and
                          $_.FriendlyName -notlike "*Camera*" -and
                          $_.FriendlyName -notlike "*Webcam*" -and
                          $_.FriendlyName -notlike "*Microphone*" -and
                          $_.FriendlyName -notlike "*Audio*" -and
                          $_.FriendlyName -notlike "*Speaker*" -and
                          $_.FriendlyName -notlike "*Headset*" -and
                          $_.FriendlyName -notlike "*Mouse*" -and
                          $_.FriendlyName -notlike "*Keyboard*" -and
                          $_.FriendlyName -notlike "*Touchpad*" -and
                          $_.FriendlyName -notlike "*Trackpad*" -and
                          $_.FriendlyName -notlike "*Monitor*" -and
                          $_.FriendlyName -notlike "*Display*" -and
                          $_.FriendlyName -notlike "*Graphics*" -and
                          $_.FriendlyName -notlike "*Video*" -and
                          $_.FriendlyName -notlike "*Network*" -and
                          $_.FriendlyName -notlike "*Ethernet*" -and
                          $_.FriendlyName -notlike "*WiFi*" -and
                          $_.FriendlyName -notlike "*Wireless*" -and
                          $_.FriendlyName -notlike "*Bluetooth*" -and
                          $_.FriendlyName -notlike "*Card Reader*" -and
                          $_.FriendlyName -notlike "*Smart Card*" -and
                          $_.FriendlyName -notlike "*USB Hub*" -and
                          $_.FriendlyName -notlike "*USB Root*" -and
                          $_.FriendlyName -notlike "*Composite*" -and
                          $_.FriendlyName -notlike "*Bus Enumerator*" -and
                          $_.FriendlyName -notlike "*Microsoft*" -and
                          $_.FriendlyName -notlike "*OneNote*" -and
                          $_.FriendlyName -notlike "*Fax*" -and
                          $_.FriendlyName -notlike "*XPS*" -and
                                                     $_.FriendlyName -notlike "*PDF*" -and
                           $_.FriendlyName -notlike "*Dock*" -and
                           $_.FriendlyName -notlike "*Dell Dock*" -and
                           $_.FriendlyName -notlike "*WD19S*" -and
                           $_.FriendlyName -notlike "*Docking*" -and
                           $_.FriendlyName -notlike "*Port Replicator*" -and
                          # Include printer devices
                          ($_.FriendlyName -like "*printer*" -or 
                           $_.FriendlyName -like "*print*" -or
                           $_.FriendlyName -like "*HP*" -or
                           $_.FriendlyName -like "*Canon*" -or
                           $_.FriendlyName -like "*Epson*" -or
                           $_.FriendlyName -like "*TM*" -or
                           $_.FriendlyName -like "*Thermal*" -or
                           $_.FriendlyName -like "*Receipt*" -or
                           $_.FriendlyName -like "*POS*" -or
                           $_.FriendlyName -like "*Point of Sale*" -or
                           $_.FriendlyName -like "*USB-to-Serial*" -or
                           $_.FriendlyName -like "*USB Serial*" -or
                           $_.FriendlyName -like "*USB Controller*" -or
                           $_.FriendlyName -like "*Serial Port*" -or
                           $_.FriendlyName -like "*COM*" -or
                           $_.FriendlyName -like "*Brother*" -or
                           $_.FriendlyName -like "*Lexmark*" -or
                           $_.FriendlyName -like "*Dell*" -or
                           $_.FriendlyName -like "*Xerox*" -or
                           $_.FriendlyName -like "*Samsung*" -or
                           $_.FriendlyName -like "*Ricoh*" -or
                           $_.FriendlyName -like "*Kyocera*" -or
                           $_.FriendlyName -like "*Sharp*" -or
                           $_.FriendlyName -like "*Konica*" -or
                           $_.FriendlyName -like "*Minolta*" -or
                           $_.FriendlyName -like "*OKI*" -or
                           $_.FriendlyName -like "*Toshiba*" -or
                           $_.FriendlyName -like "*Panasonic*" -or
                           $_.FriendlyName -like "*Fuji*")
                      }
        
        foreach ($device in $usbDevices) {
            $deviceInfo = [PSCustomObject]@{
                Name = $device.FriendlyName
                InstanceId = $device.InstanceId
                Status = $device.Status
                Class = $device.Class
            }
            
            # Extract serial number from instance ID if possible
            if ($device.InstanceId -match "USB\\VID_([A-F0-9]{4})&PID_([A-F0-9]{4})\\([A-F0-9]+)") {
                $deviceInfo.SerialNumber = Decode-HexSerial -HexSerial $matches[3]
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