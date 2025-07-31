# Enhanced USB Printer Scanner for Windows
# This script uses multiple native Windows methods to scan for USB printers
# and extract model and serial number information
# Designed for PDQ deployment

param(
    [string]$OutputPath = ".\printer_scan_results.txt",
    [switch]$Verbose,
    [switch]$ExportCSV
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
    
    try {
        if ($HexSerial.Length -eq 18 -and $HexSerial -match "^([A-Fa-f0-9]{8})([0-9]{6})(0{4})$") {
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

# Function to get printer information from Device Manager via WMI
function Get-PrinterInfoDeviceManager {
    Write-Log "Getting printer information from Device Manager via WMI..."
    
    try {
        $printerDevices = @()
        
        # Method 1: Query ALL Win32_PnPEntity devices (not just USB)
        $pnpDevices = Get-WmiObject -Class Win32_PnPEntity -ErrorAction SilentlyContinue
        
        # Also query USB devices specifically (like fast script)
        $usbDevices = Get-WmiObject -Class Win32_PnPEntity -ErrorAction SilentlyContinue | 
                      Where-Object { $_.DeviceID -like "*USB*" }
        
        foreach ($device in $pnpDevices) {
            # Exclude non-printer devices first
            if ($device.Name -like "*Fingerprint*" -or
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
                ($device.Name -like "*USB Controller*" -and $device.Name -notlike "*Epson*") -or
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
                $device.Name -like "*Dell Dock WD19S*" -or
                $device.Name -like "*Docking*" -or
                $device.Name -like "*Port Replicator*") {
                # Skip this device - it's not a printer
                continue
            }
            
            # Look for printer devices in ALL Device Manager categories
            if ($device.Name -like "*printer*" -or 
                $device.Name -like "*print*" -or
                $device.Name -like "*HP*" -or
                $device.Name -like "*Canon*" -or
                $device.Name -like "*Epson*" -or
                $device.Name -like "*TM*" -or
                $device.Name -like "*TM-T88*" -or
                $device.Name -like "*TM-T*" -or
                $device.Name -like "*TM-U*" -or
                $device.Name -like "*TM-H*" -or
                $device.Name -like "*TM-L*" -or
                $device.Name -like "*TM-S*" -or
                $device.Name -like "*TM-BA*" -or
                $device.Name -like "*TM-EU*" -or
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
                $device.Name -like "*Dell*" -or
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
                $device.Name -like "*Fuji*" -or
                $device.Name -like "*Fuji Xerox*" -or
                $device.Name -like "*HP LaserJet*" -or
                $device.Name -like "*HP OfficeJet*" -or
                $device.Name -like "*HP DeskJet*" -or
                $device.Name -like "*HP Envy*" -or
                $device.Name -like "*HP Photosmart*" -or
                $device.Name -like "*Canon PIXMA*" -or
                $device.Name -like "*Canon imageRUNNER*" -or
                $device.Name -like "*Epson WorkForce*" -or
                $device.Name -like "*Epson Expression*" -or
                $device.Name -like "*Epson TM-T88*" -or
                $device.Name -like "*Epson TM-T*" -or
                $device.Name -like "*Epson TM-U*" -or
                $device.Name -like "*Epson TM-H*" -or
                $device.Name -like "*Epson TM-L*" -or
                $device.Name -like "*Epson TM-S*" -or
                $device.Name -like "*Epson TM-BA*" -or
                $device.Name -like "*Epson TM-EU*" -or
                $device.Name -like "*Epson Receipt*" -or
                $device.Name -like "*Epson Thermal*" -or
                $device.Name -like "*Epson POS*" -or
                $device.Name -like "*Epson USB Controller*" -or
                $device.Name -like "*Brother HL*" -or
                $device.Name -like "*Brother MFC*" -or
                $device.Name -like "*Lexmark E*" -or
                $device.Name -like "*Lexmark X*" -or
                $device.Name -like "*Dell 1100*" -or
                $device.Name -like "*Dell 1130*" -or
                $device.Name -like "*Xerox Phaser*" -or
                $device.Name -like "*Xerox WorkCentre*" -or
                $device.Name -like "*Samsung ML*" -or
                $device.Name -like "*Samsung CLX*" -or
                $device.Name -like "*Ricoh Aficio*" -or
                $device.Name -like "*Ricoh SP*" -or
                $device.Name -like "*Kyocera FS*" -or
                $device.Name -like "*Kyocera ECOSYS*" -or
                $device.Name -like "*Sharp MX*" -or
                $device.Name -like "*Sharp AR*" -or
                $device.Name -like "*Konica Minolta*" -or
                $device.Name -like "*Unknown Device*" -and ($device.DeviceID -like "*USB*") -or
                $device.Name -like "*Other Device*" -and ($device.DeviceID -like "*USB*")) {
                
                $deviceInfo = [PSCustomObject]@{
                    Name = $device.Name
                    Model = $device.Name
                    DeviceID = $device.DeviceID
                    Manufacturer = $device.Manufacturer
                    Description = $device.Description
                    Status = $device.Status
                    Source = "DeviceManager_WMI"
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
        
        # Method 1.5: Query USB devices specifically (like fast script)
        foreach ($device in $usbDevices) {
            # Exclude non-printer devices first
            if ($device.Name -like "*Fingerprint*" -or
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
                ($device.Name -like "*USB Controller*" -and $device.Name -notlike "*Epson*") -or
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
                $device.Name -like "*Dell Dock WD19S*" -or
                $device.Name -like "*Docking*" -or
                $device.Name -like "*Port Replicator*") {
                # Skip this device - it's not a printer
                continue
            }
            
            # Look for printer devices (like fast script)
            if ($device.Name -like "*printer*" -or 
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
                $device.Name -like "*Fuji*") {
                
                $deviceInfo = [PSCustomObject]@{
                    Name = $device.Name
                    Model = $device.Name
                    DeviceID = $device.DeviceID
                    Manufacturer = $device.Manufacturer
                    Description = $device.Description
                    Status = $device.Status
                    Source = "DeviceManager_USB"
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
        
        # Method 2: Query USB devices specifically
        $usbDevices = Get-WmiObject -Class Win32_USBHub -ErrorAction SilentlyContinue
        
        foreach ($device in $usbDevices) {
            if ($device.Name -like "*printer*" -or 
                $device.Name -like "*print*" -or
                $device.Name -like "*HP*" -or
                $device.Name -like "*Canon*" -or
                $device.Name -like "*Epson*" -or
                $device.Name -like "*TM-T88*" -or
                $device.Name -like "*TM-T*" -or
                $device.Name -like "*TM-U*" -or
                $device.Name -like "*TM-H*" -or
                $device.Name -like "*TM-L*" -or
                $device.Name -like "*TM-S*" -or
                $device.Name -like "*Thermal*" -or
                $device.Name -like "*Receipt*" -or
                $device.Name -like "*POS*" -or
                $device.Name -like "*Point of Sale*" -or
                $device.Name -like "*Brother*" -or
                $device.Name -like "*Lexmark*" -or
                $device.Name -like "*Dell*" -or
                $device.Name -like "*Xerox*" -or
                $device.Name -like "*Samsung*" -or
                $device.Name -like "*Ricoh*" -or
                $device.Name -like "*Kyocera*" -or
                $device.Name -like "*Sharp*" -or
                $device.Name -like "*Konica*" -or
                $device.Name -like "*Minolta*") {
                
                $deviceInfo = [PSCustomObject]@{
                    Name = $device.Name
                    Model = $device.Name
                    DeviceID = $device.DeviceID
                    Manufacturer = $device.Manufacturer
                    Description = $device.Description
                    Status = $device.Status
                    Source = "DeviceManager_USBHub"
                    SerialNumber = $null
                    VID = $null
                    PID = $null
                    HardwareID = $null
                    CompatibleID = $null
                    Class = "USB"
                }
                
                # Extract serial number from device ID
                if ($device.DeviceID -match "USB\\VID_([A-Fa-f0-9]{4})&PID_([A-Fa-f0-9]{4})\\([A-Fa-f0-9]+)") {
                    $deviceInfo.SerialNumber = Decode-HexSerial -HexSerial $matches[3]
                    $deviceInfo.VID = $matches[1]
                    $deviceInfo.PID = $matches[2]
                    $deviceInfo.Manufacturer = "VID: $($matches[1]), PID: $($matches[2])"
                }
                # Also check for hex-encoded serial numbers in device ID
                elseif ($device.DeviceID -match "USB\\VID_([A-Fa-f0-9]{4})&PID_([A-Fa-f0-9]{4})\\([A-Fa-f0-9]{8,})") {
                    $hexSerial = $matches[3]
                    $deviceInfo.SerialNumber = $hexSerial
                    $deviceInfo.VID = $matches[1]
                    $deviceInfo.PID = $matches[2]
                    $deviceInfo.Manufacturer = "VID: $($matches[1]), PID: $($matches[2])"
                    
                    # Use enhanced hex decoding function
                    try {
                        if ($hexSerial.Length -ge 8 -and $hexSerial.Length % 2 -eq 0) {
                            $bytes = @()
                            for ($i = 0; $i -lt $hexSerial.Length; $i += 2) {
                                $bytes += [Convert]::ToByte($hexSerial.Substring($i, 2), 16)
                            }
                            $decodedSerial = [System.Text.Encoding]::ASCII.GetString($bytes).TrimEnd([char]0)
                            if ($decodedSerial -match '^[A-Za-z0-9\-_]+$') {
                                $deviceInfo.SerialNumber = "$hexSerial (Decoded: $decodedSerial)"
                            }
                        }
                    } catch {
                        # Keep original hex serial if decoding fails
                    }
                }
                
                $printerDevices += $deviceInfo
            }
        }
        
        # Method 3: Query ALL device classes for potential printers
        $allDeviceClasses = @("USB", "Unknown", "Other", "System", "Ports", "Printers", "Imaging")
        
        foreach ($deviceClass in $allDeviceClasses) {
            try {
                $classDevices = Get-WmiObject -Class Win32_PnPEntity -Filter "PNPClass='$deviceClass'" -ErrorAction SilentlyContinue
                
                foreach ($device in $classDevices) {
                    # Check if this device might be a printer
                    if ($device.Name -like "*printer*" -or 
                        $device.Name -like "*print*" -or
                        $device.Name -like "*HP*" -or
                        $device.Name -like "*Canon*" -or
                        $device.Name -like "*Epson*" -or
                        $device.Name -like "*TM-T88*" -or
                        $device.Name -like "*TM-T*" -or
                        $device.Name -like "*TM-U*" -or
                        $device.Name -like "*TM-H*" -or
                        $device.Name -like "*TM-L*" -or
                        $device.Name -like "*TM-S*" -or
                        $device.Name -like "*Thermal*" -or
                        $device.Name -like "*Receipt*" -or
                        $device.Name -like "*POS*" -or
                        $device.Name -like "*Point of Sale*" -or
                        $device.Name -like "*Brother*" -or
                        $device.Name -like "*Lexmark*" -or
                        $device.Name -like "*Dell*" -or
                        $device.Name -like "*Xerox*" -or
                        $device.Name -like "*Samsung*" -or
                        $device.Name -like "*Ricoh*" -or
                        $device.Name -like "*Kyocera*" -or
                        $device.Name -like "*Sharp*" -or
                        $device.Name -like "*Konica*" -or
                        $device.Name -like "*Minolta*" -or
                        ($device.DeviceID -like "*USB*" -and $device.Name -like "*Unknown*") -or
                        ($device.DeviceID -like "*USB*" -and $device.Name -like "*Other*")) {
                        
                        $deviceInfo = [PSCustomObject]@{
                            Name = $device.Name
                            Model = $device.Name
                            DeviceID = $device.DeviceID
                            Manufacturer = $device.Manufacturer
                            Description = $device.Description
                            Status = $device.Status
                            Source = "DeviceManager_$deviceClass"
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
            } catch {
                # Continue to next device class if there's an error
                continue
            }
        }
        
        return $printerDevices
        
    } catch {
        Write-Log "Error getting Device Manager printer info: $($_.Exception.Message)" "ERROR"
        return @()
    }
}

# Function to get printer information from Device Manager via Registry
function Get-PrinterInfoDeviceManagerRegistry {
    Write-Log "Getting printer information from Device Manager via Registry..."
    
    try {
        $printerDevices = @()
        
        # Query ALL device categories in registry, not just USB
        $deviceCategories = @(
            "USB",
            "Unknown",
            "Other",
            "System",
            "Ports", 
            "Printers",
            "Imaging",
            "Media",
            "Display",
            "HIDClass",
            "Net",
            "SCSIAdapter",
            "Volume",
            "Processor",
            "Computer",
            "SecurityDevices",
            "SmartCardReader",
            "Camera",
            "Image",
            "WPD"
        )
        
        foreach ($category in $deviceCategories) {
            try {
                $categoryPath = "HKLM:\SYSTEM\CurrentControlSet\Enum\$category"
                if (Test-Path $categoryPath) {
                    $devices = Get-ChildItem -Path $categoryPath -Recurse -ErrorAction SilentlyContinue
                    
                    foreach ($device in $devices) {
                        try {
                            $deviceInfo = Get-ItemProperty -Path $device.PSPath -ErrorAction SilentlyContinue
                            
                            if ($deviceInfo -and $deviceInfo.FriendlyName) {
                                $friendlyName = $deviceInfo.FriendlyName
                                
                                # Check if this looks like a printer device
                                if ($friendlyName -like "*printer*" -or 
                                    $friendlyName -like "*print*" -or
                                    $friendlyName -like "*HP*" -or
                                    $friendlyName -like "*Canon*" -or
                                    $friendlyName -like "*Epson*" -or
                                    $friendlyName -like "*TM-T88*" -or
                                    $friendlyName -like "*TM-T*" -or
                                    $friendlyName -like "*TM-U*" -or
                                    $friendlyName -like "*TM-H*" -or
                                    $friendlyName -like "*TM-L*" -or
                                    $friendlyName -like "*TM-S*" -or
                                    $friendlyName -like "*Thermal*" -or
                                    $friendlyName -like "*Receipt*" -or
                                    $friendlyName -like "*POS*" -or
                                    $friendlyName -like "*Point of Sale*" -or
                                    $friendlyName -like "*Brother*" -or
                                    $friendlyName -like "*Lexmark*" -or
                                    $friendlyName -like "*Dell*" -or
                                    $friendlyName -like "*Xerox*" -or
                                    $friendlyName -like "*Samsung*" -or
                                    $friendlyName -like "*Ricoh*" -or
                                    $friendlyName -like "*Kyocera*" -or
                                    $friendlyName -like "*Sharp*" -or
                                    $friendlyName -like "*Konica*" -or
                                    $friendlyName -like "*Minolta*" -or
                                    $friendlyName -like "*OKI*" -or
                                    $friendlyName -like "*Toshiba*" -or
                                    $friendlyName -like "*Panasonic*" -or
                                    $friendlyName -like "*Fuji*" -or
                                    $friendlyName -like "*Fuji Xerox*" -or
                                    $friendlyName -like "*HP LaserJet*" -or
                                    $friendlyName -like "*HP OfficeJet*" -or
                                    $friendlyName -like "*HP DeskJet*" -or
                                    $friendlyName -like "*HP Envy*" -or
                                    $friendlyName -like "*HP Photosmart*" -or
                                    $friendlyName -like "*Canon PIXMA*" -or
                                    $friendlyName -like "*Canon imageRUNNER*" -or
                                    $friendlyName -like "*Epson WorkForce*" -or
                                    $friendlyName -like "*Epson Expression*" -or
                                    $friendlyName -like "*Epson TM-T88*" -or
                                    $friendlyName -like "*Epson TM-T*" -or
                                    $friendlyName -like "*Epson TM-U*" -or
                                    $friendlyName -like "*Epson TM-H*" -or
                                    $friendlyName -like "*Epson TM-L*" -or
                                    $friendlyName -like "*Epson TM-S*" -or
                                    $friendlyName -like "*Epson Receipt*" -or
                                    $friendlyName -like "*Epson Thermal*" -or
                                    $friendlyName -like "*Epson POS*" -or
                                    $friendlyName -like "*Brother HL*" -or
                                    $friendlyName -like "*Brother MFC*" -or
                                    $friendlyName -like "*Lexmark E*" -or
                                    $friendlyName -like "*Lexmark X*" -or
                                    $friendlyName -like "*Dell 1100*" -or
                                    $friendlyName -like "*Dell 1130*" -or
                                    $friendlyName -like "*Xerox Phaser*" -or
                                    $friendlyName -like "*Xerox WorkCentre*" -or
                                    $friendlyName -like "*Samsung ML*" -or
                                    $friendlyName -like "*Samsung CLX*" -or
                                    $friendlyName -like "*Ricoh Aficio*" -or
                                    $friendlyName -like "*Ricoh SP*" -or
                                    $friendlyName -like "*Kyocera FS*" -or
                                    $friendlyName -like "*Kyocera ECOSYS*" -or
                                    $friendlyName -like "*Sharp MX*" -or
                                    $friendlyName -like "*Sharp AR*" -or
                                    $friendlyName -like "*Konica Minolta*" -or
                                    $friendlyName -like "*Unknown Device*" -or
                                    $friendlyName -like "*Other Device*" -or
                                    $friendlyName -like "*Generic USB*" -or
                                    $friendlyName -like "*USB Device*") {
                                    
                                    $deviceInfo = [PSCustomObject]@{
                                        Name = $friendlyName
                                        Model = $friendlyName
                                        DeviceID = $device.PSChildName
                                        Manufacturer = $deviceInfo.Mfg
                                        Description = $deviceInfo.DeviceDesc
                                        Status = "OK"
                                        Source = "DeviceManager_Registry_$category"
                                        SerialNumber = $null
                                        VID = $null
                                        PID = $null
                                        HardwareID = $null
                                        CompatibleID = $null
                                        Class = $category
                                    }
                                    
                                    # Extract serial number from device path
                                    $devicePath = $device.PSPath
                                    if ($devicePath -match "USB\\VID_([A-Fa-f0-9]{4})&PID_([A-Fa-f0-9]{4})\\([A-Fa-f0-9]+)") {
                                        $deviceInfo.SerialNumber = Decode-HexSerial -HexSerial $matches[3]
                                        $deviceInfo.VID = $matches[1]
                                        $deviceInfo.PID = $matches[2]
                                        $deviceInfo.Manufacturer = "VID: $($matches[1]), PID: $($matches[2])"
                                    }
                                    
                                    $printerDevices += $deviceInfo
                                }
                            }
                        } catch {
                            # Continue to next device if there's an error
                            continue
                        }
                    }
                }
            } catch {
                # Continue to next category if there's an error
                continue
            }
        }
        
        return $printerDevices
        
    } catch {
        Write-Log "Error getting Device Manager registry info: $($_.Exception.Message)" "ERROR"
        return @()
    }
}

# Function to get printer information using WMI
function Get-PrinterInfoWMI {
    Write-Log "Getting printer information via WMI..."
    
    try {
        $printers = Get-WmiObject -Class Win32_Printer -ErrorAction SilentlyContinue
        $usbPrinters = @()
        
        foreach ($printer in $printers) {
            # Check if printer is connected via USB
            if ($printer.PortName -like "*USB*" -or $printer.PortName -like "*USBPRN*" -or 
                $printer.PortName -like "*USB001*" -or $printer.PortName -like "*USB002*") {
                
                $printerInfo = [PSCustomObject]@{
                    Name = $printer.Name
                    Model = $printer.Name
                    PortName = $printer.PortName
                    DriverName = $printer.DriverName
                    Location = $printer.Location
                    Comment = $printer.Comment
                    Status = $printer.Status
                    Default = $printer.Default
                    Source = "Win32_Printer"
                    SerialNumber = $null
                    Manufacturer = $null
                }
                
                $usbPrinters += $printerInfo
            }
        }
        
        return $usbPrinters
    } catch {
        Write-Log "Error getting WMI printer info: $($_.Exception.Message)" "ERROR"
        return @()
    }
}

# Function to get USB device information using PnP
function Get-USBDeviceInfo {
    Write-Log "Getting USB device information via PnP..."
    
    try {
        $usbDevices = Get-PnpDevice -Class USB -ErrorAction SilentlyContinue
        $printerDevices = @()
        
        foreach ($device in $usbDevices) {
            # Look for devices that might be printers
            if ($device.FriendlyName -like "*printer*" -or 
                $device.FriendlyName -like "*print*" -or
                $device.FriendlyName -like "*HP*" -or
                $device.FriendlyName -like "*Canon*" -or
                $device.FriendlyName -like "*Epson*" -or
                $device.FriendlyName -like "*TM-T88*" -or
                $device.FriendlyName -like "*TM-T*" -or
                $device.FriendlyName -like "*TM-U*" -or
                $device.FriendlyName -like "*TM-H*" -or
                $device.FriendlyName -like "*TM-L*" -or
                $device.FriendlyName -like "*TM-S*" -or
                $device.FriendlyName -like "*Thermal*" -or
                $device.FriendlyName -like "*Receipt*" -or
                $device.FriendlyName -like "*POS*" -or
                $device.FriendlyName -like "*Point of Sale*" -or
                $device.FriendlyName -like "*Brother*" -or
                $device.FriendlyName -like "*Lexmark*" -or
                $device.FriendlyName -like "*Dell*" -or
                $device.FriendlyName -like "*Xerox*" -or
                $device.FriendlyName -like "*Samsung*" -or
                $device.FriendlyName -like "*Ricoh*" -or
                $device.FriendlyName -like "*Kyocera*" -or
                $device.FriendlyName -like "*Sharp*" -or
                $device.FriendlyName -like "*Konica*" -or
                $device.FriendlyName -like "*Minolta*") {
                
                $deviceInfo = [PSCustomObject]@{
                    Name = $device.FriendlyName
                    Model = $device.FriendlyName
                    InstanceId = $device.InstanceId
                    Status = $device.Status
                    Class = $device.Class
                    Source = "PnP_USB"
                    SerialNumber = $null
                    Manufacturer = $null
                    VID = $null
                    PID = $null
                }
                
                # Extract serial number and manufacturer info from instance ID
                if ($device.InstanceId -match "USB\\VID_([A-Fa-f0-9]{4})&PID_([A-Fa-f0-9]{4})\\([A-Fa-f0-9]+)") {
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
        Write-Log "Error getting USB device info: $($_.Exception.Message)" "ERROR"
        return @()
    }
}

# Function to get printer information from registry
function Get-PrinterInfoRegistry {
    Write-Log "Getting printer information from registry..."
    
    try {
        $registryPrinters = @()
        $printerKeys = Get-ChildItem -Path "HKLM:\SYSTEM\CurrentControlSet\Control\Print\Printers" -ErrorAction SilentlyContinue
        
        foreach ($printerKey in $printerKeys) {
            $printerName = $printerKey.PSChildName
            $portPath = "$($printerKey.PSPath)\Ports"
            
            if (Test-Path $portPath) {
                $portInfo = Get-ItemProperty -Path $portPath -ErrorAction SilentlyContinue
                
                if ($portInfo -and $portInfo.PortName -like "*USB*") {
                    $printerInfo = [PSCustomObject]@{
                        Name = $printerName
                        Model = $printerName
                        PortName = $portInfo.PortName
                        Source = "Registry"
                        SerialNumber = $null
                        Manufacturer = $null
                    }
                    
                    $registryPrinters += $printerInfo
                }
            }
        }
        
        return $registryPrinters
    } catch {
        Write-Log "Error getting registry printer info: $($_.Exception.Message)" "ERROR"
        return @()
    }
}

# Function to get detailed device information using WMI
function Get-DetailedDeviceInfo {
    param([array]$Devices)
    
    Write-Log "Getting detailed device information..."
    
    $detailedDevices = @()
    
    foreach ($device in $Devices) {
        try {
            # Try to get more information using WMI
            $wmiDevice = Get-WmiObject -Class Win32_PnPEntity -Filter "Name='$($device.Name)'" -ErrorAction SilentlyContinue
            
            if ($wmiDevice) {
                if ($device.PSObject.Properties.Name -contains "Manufacturer") {
                    $device.Manufacturer = $wmiDevice.Manufacturer
                }
                if ($device.PSObject.Properties.Name -contains "Description") {
                    $device.Description = $wmiDevice.Description
                }
                
                # Try to extract serial number from various properties
                if ($device.PSObject.Properties.Name -contains "SerialNumber" -and -not $device.SerialNumber) {
                    if ($wmiDevice.DeviceID -match "USB\\VID_([A-Fa-f0-9]{4})&PID_([A-Fa-f0-9]{4})\\([A-Fa-f0-9]+)") {
                        $device.SerialNumber = Decode-HexSerial -HexSerial $matches[3]
                    }
                    # Also check for hex-encoded serial numbers
                    elseif ($wmiDevice.DeviceID -match "USB\\VID_([A-Fa-f0-9]{4})&PID_([A-Fa-f0-9]{4})\\([A-Fa-f0-9]{8,})") {
                        $hexSerial = $matches[3]
                        $device.SerialNumber = Decode-HexSerial -HexSerial $hexSerial
                    }
                }
            }
            
            $detailedDevices += $device
            
        } catch {
            Write-Log "Error getting detailed info for $($device.Name): $($_.Exception.Message)" "ERROR"
            $detailedDevices += $device
        }
    }
    
    return $detailedDevices
}

# Function to get printer information using Print Management
function Get-PrinterInfoPrintManagement {
    Write-Log "Getting printer information via Print Management..."
    
    try {
        $printers = Get-Printer -ErrorAction SilentlyContinue
        $usbPrinters = @()
        
        foreach ($printer in $printers) {
            if ($printer.PortName -like "*USB*") {
                $printerInfo = [PSCustomObject]@{
                    Name = $printer.Name
                    Model = $printer.Name
                    PortName = $printer.PortName
                    DriverName = $printer.DriverName
                    Location = $printer.Location
                    Comment = $printer.Comment
                    Status = $printer.PrinterStatus
                    Default = $printer.Default
                    Source = "PrintManagement"
                    SerialNumber = $null
                    Manufacturer = $null
                }
                
                $usbPrinters += $printerInfo
            }
        }
        
        return $usbPrinters
    } catch {
        Write-Log "Error getting Print Management info: $($_.Exception.Message)" "ERROR"
        return @()
    }
}

# Function to get detailed printer queue information
function Get-PrinterQueueDetails {
    Write-Log "Getting detailed printer queue information..."
    
    try {
        $printerQueues = @()
        
        # Get all printers from print management
        $printers = Get-Printer -ErrorAction SilentlyContinue
        
        foreach ($printer in $printers) {
            $queueInfo = [PSCustomObject]@{
                QueueName = $printer.Name
                DisplayName = $printer.Name
                PortName = $printer.PortName
                DriverName = $printer.DriverName
                Location = $printer.Location
                Comment = $printer.Comment
                Status = $printer.PrinterStatus
                Default = $printer.Default
                Shared = $printer.Shared
                Published = $printer.Published
                Source = "PrintQueue"
                SerialNumber = $null
                Manufacturer = $null
                Model = $null
                DeviceID = $null
                IPAddress = $null
                NetworkProtocol = $null
            }
            
            # Try to get additional information from WMI
            try {
                $wmiPrinter = Get-WmiObject -Class Win32_Printer -Filter "Name='$($printer.Name)'" -ErrorAction SilentlyContinue
                if ($wmiPrinter) {
                    $queueInfo.Model = $wmiPrinter.Name
                    $queueInfo.Location = $wmiPrinter.Location
                    $queueInfo.Status = $wmiPrinter.Status
                    $queueInfo.DeviceID = $wmiPrinter.DeviceID
                }
            } catch {
                # Continue if WMI lookup fails
            }
            
            # Try to extract serial number from port name or other properties
            if ($printer.PortName -like "*USB*" -or $printer.PortName -like "*COM*") {
                # Look for serial number in port name
                if ($printer.PortName -match "USB([0-9]+)") {
                    $queueInfo.SerialNumber = "USB Port $($matches[1])"
                }
                elseif ($printer.PortName -match "COM([0-9]+)") {
                    $queueInfo.SerialNumber = "COM Port $($matches[1])"
                }
                
                # Try to match with device manager devices by port
                try {
                    $usbDevices = Get-WmiObject -Class Win32_PnPEntity -ErrorAction SilentlyContinue | 
                                  Where-Object { $_.DeviceID -like "*USB*" -or $_.DeviceID -like "*COM*" }
                    
                    foreach ($device in $usbDevices) {
                        # Try to match by port name or printer name
                        if ($device.Name -like "*$($printer.Name)*" -or 
                            $device.Name -like "*$($printer.DriverName)*" -or
                            $printer.Name -like "*$($device.Name)*") {
                            
                            if ($device.DeviceID -match "USB\\VID_([A-Fa-f0-9]{4})&PID_([A-Fa-f0-9]{4})\\([A-Fa-f0-9]+)") {
                                $queueInfo.SerialNumber = Decode-HexSerial -HexSerial $matches[3]
                                $queueInfo.Manufacturer = "VID: $($matches[1]), PID: $($matches[2])"
                            }
                            elseif ($device.DeviceID -match "USB\\VID_([A-Fa-f0-9]{4})&PID_([A-Fa-f0-9]{4})\\([A-Fa-f0-9]{8,})") {
                                $hexSerial = $matches[3]
                                $queueInfo.SerialNumber = Decode-HexSerial -HexSerial $hexSerial
                                $queueInfo.Manufacturer = "VID: $($matches[1]), PID: $($matches[2])"
                            }
                            break
                        }
                    }
                } catch {
                    # Continue if device matching fails
                }
            }
            
            # Extract IP address for network printers
            if ($printer.PortName -like "*IP*" -or $printer.PortName -like "*TCP*" -or $printer.PortName -like "*LPR*" -or $printer.PortName -like "*HTTP*") {
                # Extract IP address from port name
                if ($printer.PortName -match "(\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3})") {
                    $queueInfo.IPAddress = $matches[1]
                    $queueInfo.NetworkProtocol = "IP"
                }
                elseif ($printer.PortName -match "TCP:(\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3})") {
                    $queueInfo.IPAddress = $matches[1]
                    $queueInfo.NetworkProtocol = "TCP"
                }
                elseif ($printer.PortName -match "LPR:(\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3})") {
                    $queueInfo.IPAddress = $matches[1]
                    $queueInfo.NetworkProtocol = "LPR"
                }
                elseif ($printer.PortName -match "HTTP:(\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3})") {
                    $queueInfo.IPAddress = $matches[1]
                    $queueInfo.NetworkProtocol = "HTTP"
                }
                
                # Try to get IP from registry for this printer
                try {
                    $printerKey = "HKLM:\SYSTEM\CurrentControlSet\Control\Print\Printers\$($printer.Name)"
                    if (Test-Path $printerKey) {
                        $printerReg = Get-ItemProperty -Path $printerKey -ErrorAction SilentlyContinue
                        
                        # Check for IP address in various registry values
                        if ($printerReg.PortName -match "(\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3})") {
                            $queueInfo.IPAddress = $matches[1]
                        }
                        elseif ($printerReg.Comment -match "(\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3})") {
                            $queueInfo.IPAddress = $matches[1]
                        }
                        elseif ($printerReg.Location -match "(\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3})") {
                            $queueInfo.IPAddress = $matches[1]
                        }
                        
                        # Check for network protocol
                        if ($printerReg.PortName -like "*TCP*") {
                            $queueInfo.NetworkProtocol = "TCP"
                        }
                        elseif ($printerReg.PortName -like "*LPR*") {
                            $queueInfo.NetworkProtocol = "LPR"
                        }
                        elseif ($printerReg.PortName -like "*HTTP*") {
                            $queueInfo.NetworkProtocol = "HTTP"
                        }
                    }
                } catch {
                    # Continue if registry lookup fails
                }
            }
            
            $printerQueues += $queueInfo
        }
        
        return $printerQueues
        
    } catch {
        Write-Log "Error getting printer queue details: $($_.Exception.Message)" "ERROR"
        return @()
    }
}

# Function to detect network printers and extract IP addresses
function Get-NetworkPrinterInfo {
    Write-Log "Detecting network printers and extracting IP addresses..."
    
    try {
        $networkPrinters = @()
        
        # Get all printers from print management
        $printers = Get-Printer -ErrorAction SilentlyContinue
        
        foreach ($printer in $printers) {
            # Check if this is a network printer
            if ($printer.PortName -like "*IP*" -or 
                $printer.PortName -like "*TCP*" -or 
                $printer.PortName -like "*LPR*" -or 
                $printer.PortName -like "*HTTP*" -or
                $printer.PortName -like "*192.168.*" -or
                $printer.PortName -like "*10.*" -or
                $printer.PortName -like "*172.*" -or
                $printer.PortName -match "\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}") {
                
                $networkInfo = [PSCustomObject]@{
                    Name = $printer.Name
                    Model = $printer.Name
                    PortName = $printer.PortName
                    DriverName = $printer.DriverName
                    Location = $printer.Location
                    Comment = $printer.Comment
                    Status = $printer.PrinterStatus
                    Default = $printer.Default
                    Shared = $printer.Shared
                    Published = $printer.Published
                    Source = "NetworkPrinter"
                    SerialNumber = $null
                    Manufacturer = $null
                    IPAddress = $null
                    NetworkProtocol = $null
                    DeviceID = $null
                }
                
                # Extract IP address from port name
                if ($printer.PortName -match "(\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3})") {
                    $networkInfo.IPAddress = $matches[1]
                }
                
                # Determine network protocol
                if ($printer.PortName -like "*TCP*") {
                    $networkInfo.NetworkProtocol = "TCP"
                }
                elseif ($printer.PortName -like "*LPR*") {
                    $networkInfo.NetworkProtocol = "LPR"
                }
                elseif ($printer.PortName -like "*HTTP*") {
                    $networkInfo.NetworkProtocol = "HTTP"
                }
                elseif ($printer.PortName -like "*IP*") {
                    $networkInfo.NetworkProtocol = "IP"
                }
                else {
                    $networkInfo.NetworkProtocol = "Unknown"
                }
                
                # Try to get additional information from registry
                try {
                    $printerKey = "HKLM:\SYSTEM\CurrentControlSet\Control\Print\Printers\$($printer.Name)"
                    if (Test-Path $printerKey) {
                        $printerReg = Get-ItemProperty -Path $printerKey -ErrorAction SilentlyContinue
                        
                        # Check for IP address in various registry values
                        if (-not $networkInfo.IPAddress -and $printerReg.PortName -match "(\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3})") {
                            $networkInfo.IPAddress = $matches[1]
                        }
                        if (-not $networkInfo.IPAddress -and $printerReg.Comment -match "(\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3})") {
                            $networkInfo.IPAddress = $matches[1]
                        }
                        if (-not $networkInfo.IPAddress -and $printerReg.Location -match "(\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3})") {
                            $networkInfo.IPAddress = $matches[1]
                        }
                        
                        # Check for manufacturer info
                        if ($printer.Name -like "*Brother*") {
                            $networkInfo.Manufacturer = "Brother"
                        }
                        elseif ($printer.Name -like "*HP*") {
                            $networkInfo.Manufacturer = "HP"
                        }
                        elseif ($printer.Name -like "*Canon*") {
                            $networkInfo.Manufacturer = "Canon"
                        }
                        elseif ($printer.Name -like "*Epson*") {
                            $networkInfo.Manufacturer = "Epson"
                        }
                        elseif ($printer.Name -like "*Lexmark*") {
                            $networkInfo.Manufacturer = "Lexmark"
                        }
                        elseif ($printer.Name -like "*Xerox*") {
                            $networkInfo.Manufacturer = "Xerox"
                        }
                        elseif ($printer.Name -like "*Samsung*") {
                            $networkInfo.Manufacturer = "Samsung"
                        }
                        elseif ($printer.Name -like "*Ricoh*") {
                            $networkInfo.Manufacturer = "Ricoh"
                        }
                        elseif ($printer.Name -like "*Kyocera*") {
                            $networkInfo.Manufacturer = "Kyocera"
                        }
                        elseif ($printer.Name -like "*Sharp*") {
                            $networkInfo.Manufacturer = "Sharp"
                        }
                        elseif ($printer.Name -like "*Konica*" -or $printer.Name -like "*Minolta*") {
                            $networkInfo.Manufacturer = "Konica Minolta"
                        }
                    }
                } catch {
                    # Continue if registry lookup fails
                }
                
                $networkPrinters += $networkInfo
            }
        }
        
        return $networkPrinters
        
    } catch {
        Write-Log "Error detecting network printers: $($_.Exception.Message)" "ERROR"
        return @()
    }
}

# Function to detect Epson USB controllers and COM port printers
function Get-EpsonUSBControllers {
    Write-Log "Detecting Epson USB controllers and COM port printers..."
    
    try {
        $epsonControllers = @()
        
        # Get all USB devices that might be Epson controllers
        $usbDevices = Get-WmiObject -Class Win32_PnPEntity -ErrorAction SilentlyContinue | 
                      Where-Object { $_.DeviceID -like "*USB*" }
        
        foreach ($device in $usbDevices) {
            # Look for Epson USB controllers
            if ($device.Name -like "*Epson*" -and 
                ($device.Name -like "*USB*" -or 
                 $device.Name -like "*Controller*" -or
                 $device.Name -like "*Serial*" -or
                 $device.Name -like "*TM*" -or
                 $device.Name -like "*BA*" -or
                 $device.Name -like "*EU*")) {
                
                $controllerInfo = [PSCustomObject]@{
                    Name = $device.Name
                    Model = $device.Name
                    DeviceID = $device.DeviceID
                    Manufacturer = $device.Manufacturer
                    Description = $device.Description
                    Status = $device.Status
                    Source = "EpsonUSBController"
                    SerialNumber = $null
                    VID = $null
                    PID = $null
                    HardwareID = $device.HardwareID
                    CompatibleID = $device.CompatibleID
                    Class = $device.PNPClass
                    AssociatedCOM = $null
                    AssociatedPrinter = $null
                }
                
                # Extract serial number and manufacturer info from device ID
                if ($device.DeviceID -match "USB\\VID_([A-Fa-f0-9]{4})&PID_([A-Fa-f0-9]{4})\\([A-Fa-f0-9]+)") {
                    $controllerInfo.SerialNumber = Decode-HexSerial -HexSerial $matches[3]
                    $controllerInfo.VID = $matches[1]
                    $controllerInfo.PID = $matches[2]
                    $controllerInfo.Manufacturer = "VID: $($matches[1]), PID: $($matches[2])"
                }
                
                # Try to find associated COM port
                try {
                    $comDevices = Get-WmiObject -Class Win32_PnPEntity -ErrorAction SilentlyContinue | 
                                  Where-Object { $_.DeviceID -like "*COM*" -or $_.Name -like "*COM*" }
                    
                    foreach ($comDevice in $comDevices) {
                        # Check if this COM device might be associated with the Epson controller
                        if ($comDevice.Name -like "*Epson*" -or 
                            $comDevice.Name -like "*TM*" -or
                            $comDevice.Name -like "*Serial*" -or
                            $comDevice.DeviceID -like "*$($controllerInfo.SerialNumber)*") {
                            $controllerInfo.AssociatedCOM = $comDevice.Name
                            break
                        }
                    }
                } catch {
                    # Continue if COM port detection fails
                }
                
                # Try to find associated printer queue
                try {
                    $printers = Get-Printer -ErrorAction SilentlyContinue
                    foreach ($printer in $printers) {
                        if ($printer.PortName -like "*COM*" -and 
                            ($printer.Name -like "*Epson*" -or 
                             $printer.Name -like "*TM*" -or
                             $printer.Name -like "*Receipt*" -or
                             $printer.Name -like "*Thermal*")) {
                            $controllerInfo.AssociatedPrinter = $printer.Name
                            break
                        }
                    }
                    
                    # Also check for USB port printers that might be using the controller
                    foreach ($printer in $printers) {
                        if ($printer.PortName -like "*USB*" -and 
                            ($printer.Name -like "*Epson*" -or 
                             $printer.Name -like "*TM*" -or
                             $printer.Name -like "*Receipt*" -or
                             $printer.Name -like "*Thermal*")) {
                            $controllerInfo.AssociatedPrinter = $printer.Name
                            break
                        }
                    }
                } catch {
                    # Continue if printer queue detection fails
                }
                
                $epsonControllers += $controllerInfo
            }
        }
        
        return $epsonControllers
        
    } catch {
        Write-Log "Error detecting Epson USB controllers: $($_.Exception.Message)" "ERROR"
        return @()
    }
}

# Function to consolidate and deduplicate printer information
function Consolidate-PrinterInfo {
    param([array]$AllPrinters)
    
    Write-Log "Consolidating printer information..."
    
    $consolidated = @{}
    
    foreach ($printer in $AllPrinters) {
        # Skip if printer object is null or has no name
        if (-not $printer -or -not $printer.Name) {
            continue
        }
        
        # Use different keys for different source types
        $key = $printer.Name.ToLower()
        $queueKey = if ($printer.QueueName) { $printer.QueueName.ToLower() } else { $key }
        
        # Prefer print queue names over device names
        $preferredKey = if ($printer.Source -like "*PrintQueue*" -or $printer.Source -like "*PrintManagement*") { 
            $queueKey 
        } else { 
            $key 
        }
        
        if (-not $consolidated.ContainsKey($preferredKey)) {
            $consolidated[$preferredKey] = $printer
        } else {
            # Merge information from different sources
            $existing = $consolidated[$preferredKey]
            
            # Prefer print queue information for display name
            if ($printer.Source -like "*PrintQueue*" -or $printer.Source -like "*PrintManagement*") {
                if ($printer.QueueName) {
                    $existing.Name = $printer.QueueName
                }
                if ($printer.DisplayName) {
                    $existing.Name = $printer.DisplayName
                }
            }
            
            # Prefer non-null values
            if (-not $existing.SerialNumber -and $printer.SerialNumber) {
                $existing.SerialNumber = $printer.SerialNumber
            }
            if (-not $existing.Manufacturer -and $printer.Manufacturer) {
                $existing.Manufacturer = $printer.Manufacturer
            }
            if (-not $existing.Model -and $printer.Model) {
                $existing.Model = $printer.Model
            }
            if (-not $existing.PortName -and $printer.PortName) {
                $existing.PortName = $printer.PortName
            }
            if (-not $existing.DriverName -and $printer.DriverName) {
                $existing.DriverName = $printer.DriverName
            }
            if (-not $existing.Location -and $printer.Location) {
                $existing.Location = $printer.Location
            }
            if (-not $existing.Status -and $printer.Status) {
                $existing.Status = $printer.Status
            }
            
            # Add source information
            $existing.Source = "$($existing.Source), $($printer.Source)"
        }
    }
    
    return $consolidated.Values
}

# Function to generate a clean printer-only report
function Generate-PrinterOnlyReport {
    param(
        [array]$AllPrinters,
        [string]$OutputFile
    )
    
    try {
        $printerOnlyContent = @()
        $printerOnlyContent += "=== USB/COM Printer Devices Only ==="
        $printerOnlyContent += "Computer: $env:COMPUTERNAME"
        $printerOnlyContent += "Scan Date: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
        $printerOnlyContent += ""
        
        # Filter for actual printer devices
        $actualPrinters = @()
        
        foreach ($printer in $AllPrinters) {
            # Skip if printer object is null or has no name
            if (-not $printer -or -not $printer.Name) {
                continue
            }
            
            $isActualPrinter = $false
            
            # Check if this is an actual printer device
            # Exclude Microsoft services and non-printer devices
            if ($printer.Name -like "*Microsoft*" -or
                $printer.Name -like "*OneNote*" -or
                $printer.Name -like "*Fax*" -or
                $printer.Name -like "*XPS*" -or
                $printer.Name -like "*PDF*" -or
                $printer.Name -like "*Composite*" -or
                $printer.Name -like "*Bus Enumerator*" -or
                $printer.Name -like "*USB Composite Device*" -or
                $printer.Name -like "*Fingerprint*" -or
                $printer.Name -like "*Scanner*" -or
                $printer.Name -like "*Camera*" -or
                $printer.Name -like "*Webcam*" -or
                $printer.Name -like "*Microphone*" -or
                $printer.Name -like "*Audio*" -or
                $printer.Name -like "*Speaker*" -or
                $printer.Name -like "*Headset*" -or
                $printer.Name -like "*Mouse*" -or
                $printer.Name -like "*Keyboard*" -or
                $printer.Name -like "*Touchpad*" -or
                $printer.Name -like "*Trackpad*" -or
                $printer.Name -like "*Monitor*" -or
                $printer.Name -like "*Display*" -or
                $printer.Name -like "*Graphics*" -or
                $printer.Name -like "*Video*" -or
                $printer.Name -like "*Network*" -or
                $printer.Name -like "*Ethernet*" -or
                $printer.Name -like "*WiFi*" -or
                $printer.Name -like "*Wireless*" -or
                $printer.Name -like "*Bluetooth*" -or
                $printer.Name -like "*Card Reader*" -or
                $printer.Name -like "*Smart Card*" -or
                $printer.Name -like "*USB Hub*" -or
                $printer.Name -like "*USB Root*" -or
                $printer.Name -like "*USB Controller*" -and $printer.Name -notlike "*Epson*") {
                $isActualPrinter = $false
            }
            # Only include actual printer devices
            elseif ($printer.Source -like "*PrintQueue*" -and 
                   ($printer.Name -like "*Epson*" -or
                    $printer.Name -like "*HP*" -or
                    $printer.Name -like "*Canon*" -or
                    $printer.Name -like "*Brother*" -or
                    $printer.Name -like "*TM*" -or
                    $printer.Name -like "*Receipt*" -or
                    $printer.Name -like "*Thermal*" -or
                    $printer.Name -like "*Printer*")) {
                $isActualPrinter = $true
            }
            elseif ($printer.Source -like "*Win32_Printer*" -and 
                   ($printer.Name -like "*Epson*" -or
                    $printer.Name -like "*HP*" -or
                    $printer.Name -like "*Canon*" -or
                    $printer.Name -like "*Brother*" -or
                    $printer.Name -like "*TM*" -or
                    $printer.Name -like "*Receipt*" -or
                    $printer.Name -like "*Thermal*" -or
                    $printer.Name -like "*Printer*")) {
                $isActualPrinter = $true
            }
            elseif ($printer.Source -like "*PrintManagement*" -and 
                   ($printer.Name -like "*Epson*" -or
                    $printer.Name -like "*HP*" -or
                    $printer.Name -like "*Canon*" -or
                    $printer.Name -like "*Brother*" -or
                    $printer.Name -like "*TM*" -or
                    $printer.Name -like "*Receipt*" -or
                    $printer.Name -like "*Thermal*" -or
                    $printer.Name -like "*Printer*")) {
                $isActualPrinter = $true
            }
            elseif ($printer.Source -like "*EpsonUSBController*") {
                $isActualPrinter = $true
            }
            elseif ($printer.Source -like "*DeviceManager*" -and 
                   ($printer.Name -like "*Epson*" -or
                    $printer.Name -like "*HP*" -or
                    $printer.Name -like "*Canon*" -or
                    $printer.Name -like "*Brother*" -or
                    $printer.Name -like "*TM*" -or
                    $printer.Name -like "*Receipt*" -or
                    $printer.Name -like "*Thermal*" -or
                    $printer.Name -like "*USB Controller*" -or
                    $printer.Name -like "*COM Emulation*" -or
                    $printer.Name -like "*COM Port*")) {
                $isActualPrinter = $true
            }
            
            # Only include devices with actual printer-related ports (exclude Microsoft services)
            elseif ($printer.PortName -like "*USB*" -and 
                   $printer.Name -notlike "*Microsoft*" -and
                   $printer.Name -notlike "*OneNote*" -and
                   $printer.Name -notlike "*Fax*" -and
                   $printer.Name -notlike "*XPS*" -and
                   $printer.Name -notlike "*PDF*") {
                $isActualPrinter = $true
            }
            elseif ($printer.PortName -like "*COM*" -and 
                   ($printer.Name -like "*Epson*" -or
                    $printer.Name -like "*TM*" -or
                    $printer.Name -like "*Receipt*" -or
                    $printer.Name -like "*Thermal*")) {
                $isActualPrinter = $true
            }
            elseif ($printer.PortName -like "*TMUSB*" -or $printer.PortName -like "*LPT*") {
                $isActualPrinter = $true
            }
            
            # Only include devices with actual printer drivers (exclude Microsoft services)
            elseif ($printer.DriverName -like "*Epson*" -or
                   $printer.DriverName -like "*HP*" -or
                   $printer.DriverName -like "*Canon*" -or
                   $printer.DriverName -like "*Brother*" -or
                   $printer.DriverName -like "*TM*" -or
                   $printer.DriverName -like "*Receipt*" -or
                   $printer.DriverName -like "*Thermal*") {
                $isActualPrinter = $true
            }
            
            if ($isActualPrinter) {
                $actualPrinters += $printer
            }
        }
        
        if ($actualPrinters.Count -eq 0) {
            $printerOnlyContent += "No actual USB/COM printer devices found."
        } else {
            $printerOnlyContent += "Found $($actualPrinters.Count) actual USB/COM printer device(s):"
            $printerOnlyContent += ""
            
            foreach ($printer in $actualPrinters) {
                $printerOnlyContent += "=== PRINTER DEVICE ==="
                $printerOnlyContent += "Name: $($printer.Name)"
                $printerOnlyContent += "Model: $($printer.Model)"
                $printerOnlyContent += "Serial Number: $($printer.SerialNumber)"
                $printerOnlyContent += "Manufacturer: $($printer.Manufacturer)"
                $printerOnlyContent += "Port: $($printer.PortName)"
                $printerOnlyContent += "Driver: $($printer.DriverName)"
                $printerOnlyContent += "Location: $($printer.Location)"
                $printerOnlyContent += "Status: $($printer.Status)"
                $printerOnlyContent += "Source: $($printer.Source)"
                $printerOnlyContent += "Device ID: $($printer.DeviceID)"
                $printerOnlyContent += "VID: $($printer.VID)"
                $printerOnlyContent += "PID: $($printer.PID)"
                
                if ($printer.IPAddress) {
                    $printerOnlyContent += "IP Address: $($printer.IPAddress)"
                }
                if ($printer.NetworkProtocol) {
                    $printerOnlyContent += "Network Protocol: $($printer.NetworkProtocol)"
                }
                
                if ($printer.QueueName -and $printer.QueueName -ne $printer.Name) {
                    $printerOnlyContent += "Queue Name: $($printer.QueueName)"
                }
                if ($printer.DisplayName -and $printer.DisplayName -ne $printer.Name) {
                    $printerOnlyContent += "Display Name: $($printer.DisplayName)"
                }
                if ($printer.Shared) {
                    $printerOnlyContent += "Shared: $($printer.Shared)"
                }
                if ($printer.Published) {
                    $printerOnlyContent += "Published: $($printer.Published)"
                }
                if ($printer.AssociatedCOM) {
                    $printerOnlyContent += "Associated COM Port: $($printer.AssociatedCOM)"
                }
                if ($printer.AssociatedPrinter) {
                    $printerOnlyContent += "Associated Printer: $($printer.AssociatedPrinter)"
                }
                $printerOnlyContent += ""
            }
        }
        
        # Write to file
        $printerOnlyContent | Out-File -FilePath $OutputFile -Encoding UTF8
        Write-Log "Printer-only report generated: $OutputFile" "INFO"
        
    } catch {
        Write-Log "Error generating printer-only report: $($_.Exception.Message)" "ERROR"
    }
}

# Main execution
Write-Log "=== Enhanced USB Printer Scanner Started ==="
Write-Log "Computer Name: $env:COMPUTERNAME"
Write-Log "User: $env:USERNAME"
Write-Log "Date: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"

# Collect printer information from multiple sources
$allPrinters = @()

# Method 1: Device Manager via WMI (NEW - PRIMARY METHOD)
$deviceManagerPrinters = Get-PrinterInfoDeviceManager
$allPrinters += $deviceManagerPrinters
Write-Log "Found $($deviceManagerPrinters.Count) printer(s) via Device Manager WMI"

# Method 2: Device Manager via Registry (NEW - SECONDARY METHOD)
$deviceManagerRegistryPrinters = Get-PrinterInfoDeviceManagerRegistry
$allPrinters += $deviceManagerRegistryPrinters
Write-Log "Found $($deviceManagerRegistryPrinters.Count) printer(s) via Device Manager Registry"

# Method 3: WMI
$wmiPrinters = Get-PrinterInfoWMI
$allPrinters += $wmiPrinters
Write-Log "Found $($wmiPrinters.Count) printer(s) via WMI"

# Method 4: PnP USB Devices
$pnpPrinters = Get-USBDeviceInfo
$allPrinters += $pnpPrinters
Write-Log "Found $($pnpPrinters.Count) printer(s) via PnP"

# Method 5: Registry
$registryPrinters = Get-PrinterInfoRegistry
$allPrinters += $registryPrinters
Write-Log "Found $($registryPrinters.Count) printer(s) via Registry"

# Method 6: Print Management
$printMgmtPrinters = Get-PrinterInfoPrintManagement
$allPrinters += $printMgmtPrinters
Write-Log "Found $($printMgmtPrinters.Count) printer(s) via Print Management"

# Method 7: Printer Queue Details (NEW)
$printerQueues = Get-PrinterQueueDetails
$allPrinters += $printerQueues
Write-Log "Found $($printerQueues.Count) printer queue(s) with details"

# Method 8: Epson USB Controllers and COM Port Printers (NEW)
$epsonControllers = Get-EpsonUSBControllers
$allPrinters += $epsonControllers
Write-Log "Found $($epsonControllers.Count) Epson USB controller(s) with COM port printers"

# Method 9: Network Printers with IP Addresses (NEW)
$networkPrinters = Get-NetworkPrinterInfo
$allPrinters += $networkPrinters
Write-Log "Found $($networkPrinters.Count) network printer(s) with IP addresses"

# Get detailed information
$detailedPrinters = Get-DetailedDeviceInfo -Devices $allPrinters

# Consolidate results
$finalPrinters = Consolidate-PrinterInfo -AllPrinters $detailedPrinters

# Display results
if ($finalPrinters.Count -eq 0) {
    Write-Log "No USB printers found on this system." "WARNING"
    Write-Output "USB_PRINTERS_FOUND: 0"
} else {
    Write-Log "Found $($finalPrinters.Count) USB printer(s)"
    Write-Log "=== USB Printer Scan Results ==="
    
    $csvData = @()
    
            foreach ($printer in $finalPrinters) {
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
            if ($printer.IPAddress) {
                Write-Log "  IP Address: $($printer.IPAddress)"
            }
            if ($printer.NetworkProtocol) {
                Write-Log "  Network Protocol: $($printer.NetworkProtocol)"
            }
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
            if ($printer.AssociatedCOM) {
                Write-Log "  Associated COM Port: $($printer.AssociatedCOM)"
            }
            if ($printer.AssociatedPrinter) {
                Write-Log "  Associated Printer: $($printer.AssociatedPrinter)"
            }
            Write-Log "---"
        
        # Prepare CSV data
        $csvRow = [PSCustomObject]@{
            ComputerName = $env:COMPUTERNAME
            PrinterName = $printer.Name
            QueueName = $printer.QueueName
            DisplayName = $printer.DisplayName
            Model = $printer.Model
            SerialNumber = $printer.SerialNumber
            Manufacturer = $printer.Manufacturer
            PortName = $printer.PortName
            DriverName = $printer.DriverName
            Location = $printer.Location
            Status = $printer.Status
            Source = $printer.Source
            DeviceID = $printer.DeviceID
            VID = $printer.VID
            PID = $printer.PID
            IPAddress = $printer.IPAddress
            NetworkProtocol = $printer.NetworkProtocol
            Shared = $printer.Shared
            Published = $printer.Published
            AssociatedCOM = $printer.AssociatedCOM
            AssociatedPrinter = $printer.AssociatedPrinter
            ScanDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        }
        $csvData += $csvRow
        
        # Output for PDQ
        Write-Output "PRINTER: $($printer.Name) | Model: $($printer.Model) | Serial: $($printer.SerialNumber)"
    }
    
    Write-Output "USB_PRINTERS_FOUND: $($finalPrinters.Count)"
    
    # Export to CSV if requested
    if ($ExportCSV) {
        $csvPath = ".\printer_scan_results.csv"
        $csvData | Export-Csv -Path $csvPath -NoTypeInformation
        Write-Log "CSV results exported to: $csvPath"
    }
}

Write-Log "=== Enhanced USB Printer Scanner Completed ==="
Write-Log "Results saved to: $OutputPath"

# Generate clean printer-only report
$printerOnlyFile = ".\printer_devices_only.txt"
Generate-PrinterOnlyReport -AllPrinters $finalPrinters -OutputFile $printerOnlyFile
Write-Log "Printer-only report saved to: $printerOnlyFile" 