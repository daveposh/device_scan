# Fast USB Printer Scanner for Windows
# Streamlined version that focuses only on finding actual printers quickly
# Designed for PDQ deployment

param(
    [string]$OutputPath = ".\printer_scan_fast.txt",
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
    
    $HexSerial = $HexSerial.Trim()  # Remove any leading/trailing whitespace or newlines
    Write-Log "Decoding serial: $HexSerial" "DEBUG"
    
    # Debug: Show ASCII values of each character
    for ($i = 0; $i -lt $HexSerial.Length; $i++) {
        $c = $HexSerial[$i]
        $ascii = [int]$c
        Write-Log "Char $i: '$c' (ASCII: $ascii)" "DEBUG"
    }
    
    try {
        if ($HexSerial.Length -eq 20 -and $HexSerial -match "^([A-Fa-f0-9]{8})([0-9]{6})(0{4})$") {
            Write-Log "Epson pattern matched for: $HexSerial" "DEBUG"
            $hexPart = $matches[1]  # First 8 hex characters
            $numberPart = $matches[2]  # Next 6 numbers
            $padding = $matches[3]  # Last 4 zeros
            
            # Decode the hex part
            $hexDecoded = ""
            for ($i = 0; $i -lt $hexPart.Length; $i += 2) {
                $hexPair = $hexPart.Substring($i, 2)
                $byteValue = [Convert]::ToByte($hexPair, 16)
                Write-Log "  Decoding $hexPair = $byteValue = '$([char]$byteValue)'" "DEBUG"
                # Allow all printable ASCII characters (32-126) and some control chars
                if ($byteValue -ge 32 -and $byteValue -le 126) {
                    $hexDecoded += [char]$byteValue
                    Write-Log "  Added character: '$([char]$byteValue)'" "DEBUG"
                } else {
                    Write-Log "  Skipped character (out of range): $byteValue" "DEBUG"
                }
            }
            
            # Combine hex decoded part with number part
            $epsonDecoded = $hexDecoded + $numberPart
            
            Write-Log "Epson decoded result: $epsonDecoded" "DEBUG"
            
            if ($epsonDecoded.Length -gt 0) {
                return "$HexSerial (Epson Decoded: $epsonDecoded)"
            }
        } else {
            Write-Log "Epson pattern NOT matched for: $HexSerial" "DEBUG"
        }
        
        # Standard ASCII decoding (try this after Epson-specific decoding)
        $bytes = @()
        for ($i = 0; $i -lt $HexSerial.Length; $i += 2) {
            $bytes += [Convert]::ToByte($HexSerial.Substring($i, 2), 16)
        }
        $decodedSerial = [System.Text.Encoding]::ASCII.GetString($bytes).TrimEnd([char]0)
        
        # Check if decoded result looks like a valid serial number
        if ($decodedSerial -match '^[A-Za-z0-9\-_]+$' -and $decodedSerial.Length -gt 0) {
            return "$HexSerial (Decoded: $decodedSerial)"
        }
        
        # If no valid decoding found, return original hex
        return $HexSerial
    } catch {
        # Return original hex if decoding fails
        return $HexSerial
    }
}

# Function to get printer information from Device Manager (FAST)
function Get-PrinterInfoFast {
    Write-Log "Getting printer information from Device Manager (FAST)..."
    
    try {
        $printerDevices = @()
        
        # Query only USB devices for speed
        $usbDevices = Get-WmiObject -Class Win32_PnPEntity -ErrorAction SilentlyContinue | 
                      Where-Object { $_.DeviceID -like "*USB*" }
        
        foreach ($device in $usbDevices) {
            # Exclude non-printer devices first
            if (
                ($device.Name -like "*Fingerprint*" -or
                 $device.Name -like "*Scanner*" -or
                 $device.Name -like "*Camera*" -or
                 $device.Name -like "*Webcam*" -or
                 $device.Name -like "*Microphone*" -or
                 $device.Name -like "*Audio*" -or
                 $device.Name -like "*Speaker*" -or
                 $device.Name -like "*Headset*" -or
                 $device.Name -like "*Mouse*" -or
                 $device.Name -like "*Keyboard*" -or
                 $device.Name -like "*Touchpad*" -or
                 $device.Name -like "*Trackpad*" -or
                 $device.Name -like "*Monitor*" -or
                 $device.Name -like "*Display*" -or
                 $device.Name -like "*Graphics*" -or
                 $device.Name -like "*Video*" -or
                 $device.Name -like "*Network*" -or
                 $device.Name -like "*Ethernet*" -or
                 $device.Name -like "*WiFi*" -or
                 $device.Name -like "*Wireless*" -or
                 $device.Name -like "*Bluetooth*" -or
                 $device.Name -like "*Card Reader*" -or
                 $device.Name -like "*Smart Card*" -or
                 $device.Name -like "*USB Hub*" -or
                 $device.Name -like "*USB Root*" -or
                 $device.Name -like "*Composite*" -or
                 $device.Name -like "*Bus Enumerator*" -or
                 $device.Name -like "*Microsoft*" -or
                 $device.Name -like "*OneNote*" -or
                 $device.Name -like "*Fax*" -or
                 $device.Name -like "*XPS*" -or
                 $device.Name -like "*PDF*" -or
                 $device.Name -like "*Dock*" -or
                 $device.Name -like "*Dell Dock*" -or
                 $device.Name -like "*WD19S*" -or
                 $device.Name -like "*Docking*" -or
                 $device.Name -like "*Port Replicator*" -or
                 $device.Name -like "*Dell Dock WD19S*")
            ) {
                continue
            }
            
            # Look for printer devices (exclude Dell Dock and other non-printer devices)
            if (
                ($device.Name -like "*printer*" -or 
                $device.Name -like "*print*" -or
                $device.Name -like "*HP*" -or
                $device.Name -like "*Canon*" -or
                $device.Name -like "*Epson*" -or
                $device.Name -like "*TM*" -or
                $device.Name -like "*Thermal*" -or
                $device.Name -like "*Receipt*" -or
                $device.Name -like "*POS*" -or
                $device.Name -like "*Point of Sale*" -or
                $device.Name -like "*USB-to-Serial*" -or
                $device.Name -like "*USB Serial*" -or
                $device.Name -like "*USB Controller*" -or
                $device.Name -like "*Serial Port*" -or
                $device.Name -like "*COM*" -or
                $device.Name -like "*Brother*" -or
                $device.Name -like "*Lexmark*" -or
                $device.Name -like "*Xerox*" -or
                $device.Name -like "*Samsung*" -or
                $device.Name -like "*Ricoh*" -or
                $device.Name -like "*Kyocera*" -or
                $device.Name -like "*Sharp*" -or
                $device.Name -like "*Konica*" -or
                $device.Name -like "*Minolta*" -or
                $device.Name -like "*OKI*" -or
                $device.Name -like "*Toshiba*" -or
                $device.Name -like "*Panasonic*" -or
                $device.Name -like "*Fuji*") -and
                $device.Name -notlike "*Dock*" -and
                $device.Name -notlike "*Dell Dock*" -and
                $device.Name -notlike "*WD19S*" -and
                $device.Name -notlike "*Dell Dock WD19S*"
            ) {
                
                $deviceInfo = [PSCustomObject]@{
                    Name = $device.Name
                    Model = $device.Name
                    DeviceID = $device.DeviceID
                    Manufacturer = $device.Manufacturer
                    Description = $device.Description
                    Status = $device.Status
                    Source = "DeviceManager_Fast"
                    SerialNumber = $null
                    VID = $null
                    PID = $null
                    HardwareID = $device.HardwareID
                    CompatibleID = $device.CompatibleID
                    Class = $device.PNPClass
                }
                
                # Extract serial number and manufacturer info from device ID
                if ($device.DeviceID -match "USB\\VID_([A-Fa-f0-9]{4})&PID_([A-Fa-f0-9]{4})\\([A-Fa-f0-9]+)") {
                    $deviceInfo.SerialNumber = Decode-HexSerial -HexSerial $matches[3]
                    $deviceInfo.VID = $matches[1]
                    $deviceInfo.PID = $matches[2]
                    $deviceInfo.Manufacturer = "VID: $($matches[1]), PID: $($matches[2])"
                }
                
                $printerDevices += $deviceInfo
            }
        }
        
        return $printerDevices
        
    } catch {
        Write-Log "Error getting Device Manager printer info: $($_.Exception.Message)" "ERROR"
        return @()
    }
}

# Function to get printer queue information (FAST)
function Get-PrinterQueueFast {
    Write-Log "Getting printer queue information (FAST)..."
    
    try {
        $printerQueues = @()
        
        # Get all printers from print management
        $printers = Get-Printer -ErrorAction SilentlyContinue
        
        foreach ($printer in $printers) {
            # Only include actual printer devices, not Microsoft services
            if ($printer.Name -notlike "*Microsoft*" -and
                $printer.Name -notlike "*OneNote*" -and
                $printer.Name -notlike "*Fax*" -and
                $printer.Name -notlike "*XPS*" -and
                $printer.Name -notlike "*PDF*") {
                
                $queueInfo = [PSCustomObject]@{
                    Name = $printer.Name
                    Model = $printer.Name
                    PortName = $printer.PortName
                    DriverName = $printer.DriverName
                    Location = $printer.Location
                    Status = $printer.PrinterStatus
                    Source = "PrintQueue_Fast"
                    SerialNumber = $null
                    Manufacturer = $null
                    QueueName = $printer.Name
                    DisplayName = $printer.Name
                    Shared = $printer.Shared
                    Published = $printer.Published
                }
                
                # Try to link with USB controller to get serial number
                try {
                    $usbDevices = Get-WmiObject -Class Win32_PnPEntity -ErrorAction SilentlyContinue | 
                                  Where-Object { $_.DeviceID -like "*USB*" -and $_.Name -like "*Epson*" }
                    
                    foreach ($usbDevice in $usbDevices) {
                        if ($usbDevice.DeviceID -match "USB\\VID_([A-Fa-f0-9]{4})&PID_([A-Fa-f0-9]{4})\\([A-Fa-f0-9]+)") {
                            $queueInfo.SerialNumber = Decode-HexSerial -HexSerial $matches[3]
                            $queueInfo.Manufacturer = "VID: $($matches[1]), PID: $($matches[2])"
                            break
                        }
                    }
                } catch {
                    # Continue if linking fails
                }
                
                $printerQueues += $queueInfo
            }
        }
        
        return $printerQueues
    } catch {
        Write-Log "Error getting printer queue info: $($_.Exception.Message)" "ERROR"
        return @()
    }
}

# Main execution
Write-Log "=== Fast USB Printer Scanner Started ==="
Write-Log "Computer Name: $env:COMPUTERNAME"
Write-Log "User: $env:USERNAME"
Write-Log "Date: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"

# Collect printer information from fast sources
$allPrinters = @()

# Method 1: Device Manager (FAST)
$deviceManagerPrinters = Get-PrinterInfoFast
$allPrinters += $deviceManagerPrinters
Write-Log "Found $($deviceManagerPrinters.Count) printer(s) via Device Manager (Fast)"

# Method 2: Printer Queues (FAST)
$printerQueues = Get-PrinterQueueFast
$allPrinters += $printerQueues
Write-Log "Found $($printerQueues.Count) printer queue(s) (Fast)"

# Display results
if ($allPrinters.Count -eq 0) {
    Write-Log "No USB printers found on this system." "WARNING"
    Write-Output "USB_PRINTERS_FOUND: 0"
} else {
    Write-Log "Found $($allPrinters.Count) USB printer(s)"
    Write-Log "=== Fast USB Printer Scan Results ==="
    
    foreach ($printer in $allPrinters) {
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
        Write-Log "  VID: $($printer.VID)"
        Write-Log "  PID: $($printer.PID)"
        if ($printer.QueueName -and $printer.QueueName -ne $printer.Name) {
            Write-Log "  Queue Name: $($printer.QueueName)"
        }
        if ($printer.DisplayName -and $printer.DisplayName -ne $printer.Name) {
            Write-Log "  Display Name: $($printer.DisplayName)"
        }
        if ($printer.Shared) {
            Write-Log "  Shared: $($printer.Shared)"
        }
        if ($printer.Published) {
            Write-Log "  Published: $($printer.Published)"
        }
        Write-Log "---"
        
        # Output for PDQ
        Write-Output "PRINTER: $($printer.Name) | Model: $($printer.Model) | Serial: $($printer.SerialNumber)"
    }
    
    Write-Output "USB_PRINTERS_FOUND: $($allPrinters.Count)"
}

Write-Log "=== Fast USB Printer Scanner Completed ==="
Write-Log "Results saved to: $OutputPath" 