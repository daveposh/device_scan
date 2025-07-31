# Freshservice Printer Asset Integration Script (Simplified)
# This script scans for printers and adds them as assets in Freshservice using config file

param(
    [Parameter(Mandatory=$false)]
    [string]$ConfigFile = "freshservice_config.json",
    
    [Parameter(Mandatory=$false)]
    [switch]$DryRun,
    
    [Parameter(Mandatory=$false)]
    [switch]$TestConnection
)

function Write-Log {
    param([string]$Message)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Write-Host "[$timestamp] $Message"
}

# Import the printer scanner functions
try {
    . .\printer_scanner_enhanced.ps1
    Write-Log "✅ Successfully imported printer scanner functions"
} catch {
    Write-Log "⚠️ Could not import printer scanner, using built-in functions"
}

function Load-Configuration {
    param([string]$ConfigPath)
    
    try {
        if (-not (Test-Path $ConfigPath)) {
            Write-Log "❌ Configuration file not found: $ConfigPath"
            Write-Log "Please create $ConfigPath with your Freshservice settings"
            return $null
        }
        
        $config = Get-Content $ConfigPath -Raw | ConvertFrom-Json
        Write-Log "✅ Configuration loaded from $ConfigPath"
        return $config
    }
    catch {
        Write-Log "❌ Failed to load configuration: $($_.Exception.Message)"
        return $null
    }
}

function Test-FreshserviceConnection {
    param($Config)
    
    $BaseUrl = "https://$($Config.freshservice.domain).freshservice.com/api/v2"
    $Headers = @{
        "Content-Type" = "application/json"
        "Authorization" = "Basic " + [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes("$($Config.freshservice.api_key)`:"))
    }
    
    try {
        $response = Invoke-RestMethod -Uri "$BaseUrl/assets" -Headers $Headers -Method Get -ErrorAction Stop
        Write-Log "✅ Successfully connected to Freshservice API"
        return @{ BaseUrl = $BaseUrl; Headers = $Headers }
    }
    catch {
        Write-Log "❌ Failed to connect to Freshservice API: $($_.Exception.Message)"
        return $null
    }
}

function Get-FreshserviceAssetTypes {
    param($ApiConfig)
    
    try {
        $response = Invoke-RestMethod -Uri "$($ApiConfig.BaseUrl)/asset_types" -Headers $ApiConfig.Headers -Method Get -ErrorAction Stop
        Write-Log "✅ Retrieved $($response.asset_types.Count) asset types from Freshservice"
        return $response.asset_types
    }
    catch {
        Write-Log "❌ Failed to get asset types: $($_.Exception.Message)"
        return @()
    }
}

function Get-PrinterAssetType {
    param(
        $Config,
        [string]$PrinterName
    )
    
    $printerNameLower = $PrinterName.ToLower()
    
    foreach ($printerType in $Config.printer_types.PSObject.Properties) {
        $typeConfig = $printerType.Value
        foreach ($keyword in $typeConfig.keywords) {
            if ($printerNameLower -like "*$($keyword.ToLower())*") {
                return $typeConfig.asset_type
            }
        }
    }
    
    return $Config.freshservice.default_asset_type
}

function Find-ExistingAsset {
    param(
        $ApiConfig,
        [string]$SerialNumber,
        [string]$Name
    )
    
    try {
        # Search by serial number first
        if ($SerialNumber) {
            $response = Invoke-RestMethod -Uri "$($ApiConfig.BaseUrl)/assets?search_string=$SerialNumber" -Headers $ApiConfig.Headers -Method Get -ErrorAction Stop
            if ($response.assets.Count -gt 0) {
                return $response.assets[0]
            }
        }
        
        # Search by name if serial number not found
        if ($Name) {
            $response = Invoke-RestMethod -Uri "$($ApiConfig.BaseUrl)/assets?search_string=$Name" -Headers $ApiConfig.Headers -Method Get -ErrorAction Stop
            if ($response.assets.Count -gt 0) {
                return $response.assets[0]
            }
        }
        
        return $null
    }
    catch {
        Write-Log "⚠️ Error searching for existing asset: $($_.Exception.Message)"
        return $null
    }
}

function Find-ComputerAsset {
    param(
        $ApiConfig,
        [string]$Hostname
    )
    
    try {
        # Search for computer asset by hostname
        $response = Invoke-RestMethod -Uri "$($ApiConfig.BaseUrl)/assets?search_string=$Hostname" -Headers $ApiConfig.Headers -Method Get -ErrorAction Stop
        
        if ($response.assets.Count -gt 0) {
            # Look for a computer/laptop asset type
            foreach ($asset in $response.assets) {
                if ($asset.asset_type_name -like "*Computer*" -or 
                    $asset.asset_type_name -like "*Laptop*" -or 
                    $asset.asset_type_name -like "*Desktop*" -or
                    $asset.name -eq $Hostname) {
                    return $asset
                }
            }
        }
        
        return $null
    }
    catch {
        Write-Log "⚠️ Error searching for computer asset: $($_.Exception.Message)"
        return $null
    }
}

function Get-PrinterInfoSimple {
    # Fallback printer scanning function if the enhanced scanner is not available
    try {
        $printers = @()
        
        # Method 1: Get printers from WMI
        $wmiPrinters = Get-WmiObject -Class Win32_Printer -ErrorAction SilentlyContinue
        foreach ($printer in $wmiPrinters) {
            if ($printer.PortName -like "*USB*" -or $printer.PortName -like "*COM*") {
                $printerInfo = [PSCustomObject]@{
                    Name = $printer.Name
                    Model = $printer.Model
                    SerialNumber = $null
                    Manufacturer = $printer.Manufacturer
                    PortName = $printer.PortName
                    DriverName = $printer.DriverName
                    Status = $printer.Status
                    Source = "WMI_Simple"
                    DeviceID = $null
                    IPAddress = $null
                    NetworkProtocol = $null
                }
                $printers += $printerInfo
            }
        }
        
        # Method 2: Get USB devices
        $usbDevices = Get-WmiObject -Class Win32_PnPEntity -ErrorAction SilentlyContinue | 
                      Where-Object { $_.DeviceID -like "*USB*" }
        
        foreach ($device in $usbDevices) {
            if ($device.Name -like "*printer*" -or 
                $device.Name -like "*Epson*" -or 
                $device.Name -like "*HP*" -or 
                $device.Name -like "*Canon*" -or 
                $device.Name -like "*Brother*") {
                
                $serialNumber = $null
                if ($device.DeviceID -match "USB\\VID_([A-Fa-f0-9]{4})&PID_([A-Fa-f0-9]{4})\\([A-Fa-f0-9]+)") {
                    $serialNumber = $matches[3]
                }
                
                $printerInfo = [PSCustomObject]@{
                    Name = $device.Name
                    Model = $device.Name
                    SerialNumber = $serialNumber
                    Manufacturer = $device.Manufacturer
                    PortName = $null
                    DriverName = $device.DriverName
                    Status = $device.Status
                    Source = "USB_Simple"
                    DeviceID = $device.DeviceID
                    IPAddress = $null
                    NetworkProtocol = $null
                }
                $printers += $printerInfo
            }
        }
        
        return $printers
    }
    catch {
        Write-Log "❌ Error in simple printer scanning: $($_.Exception.Message)"
        return @()
    }
}

function New-FreshserviceAsset {
    param(
        $ApiConfig,
        $Config,
        [object]$Printer
    )
    
    try {
        # Check if asset already exists
        $existingAsset = Find-ExistingAsset -ApiConfig $ApiConfig -SerialNumber $Printer.SerialNumber -Name $Printer.Name
        if ($existingAsset) {
            Write-Log "⚠️ Asset already exists: $($existingAsset.display_id) - $($existingAsset.name)"
            return $existingAsset
        }
        
        # Find the computer asset to associate with
        $computerAsset = Find-ComputerAsset -ApiConfig $ApiConfig -Hostname $env:COMPUTERNAME
        if ($computerAsset) {
            Write-Log "🖥️ Found computer asset: $($computerAsset.display_id) - $($computerAsset.name)"
        } else {
            Write-Log "⚠️ Computer asset not found for hostname: $env:COMPUTERNAME"
        }
        
        # Get asset type ID based on printer
        $assetTypeName = Get-PrinterAssetType -Config $Config -PrinterName $Printer.Name
        $assetTypeId = ($Config.asset_types | Where-Object { $_.name -eq $assetTypeName }).id
        if (-not $assetTypeId) {
            # Fallback to default asset type
            $assetTypeId = ($Config.asset_types | Where-Object { $_.name -eq $Config.freshservice.default_asset_type }).id
            if (-not $assetTypeId) {
                # Use first available asset type
                $assetTypeId = ($Config.asset_types | Select-Object -First 1).id
            }
        }
        
        # Prepare asset data according to Freshservice API specification
        $assetTag = if ($Printer.SerialNumber) { $Printer.SerialNumber } else { "PRN-$(Get-Random -Minimum 1000 -Maximum 9999)" }
        
        $assetData = @{
            asset_type_id = $assetTypeId
            name = $Printer.Name
            description = "Auto-discovered printer from device scan on $env:COMPUTERNAME"
            asset_tag = $assetTag
            serial_number = $Printer.SerialNumber
            manufacturer = $Printer.Manufacturer
            model = $Printer.Model
            custom_fields = @{
                port_name = $Printer.PortName
                driver_name = $Printer.DriverName
                device_id = $Printer.DeviceID
                source = $Printer.Source
                status = $Printer.Status
                ip_address = $Printer.IPAddress
                network_protocol = $Printer.NetworkProtocol
                computer_name = $env:COMPUTERNAME
                discovery_date = (Get-Date -Format "yyyy-MM-dd")
            }
        }
        
        # Associate with computer asset if found
        if ($computerAsset) {
            $assetData.parent_asset_id = $computerAsset.id
        }
        
        if ($DryRun) {
            Write-Log "🔍 DRY RUN: Would create $assetTypeName asset for $($Printer.Name)"
            Write-Log "   Asset Type ID: $assetTypeId"
            if ($computerAsset) {
                Write-Log "   Associated with computer: $($computerAsset.name) (ID: $($computerAsset.id))"
            }
            Write-Log "   Asset Data: $($assetData | ConvertTo-Json -Depth 3)"
            return $null
        }
        
        $body = $assetData | ConvertTo-Json -Depth 3
        $response = Invoke-RestMethod -Uri "$($ApiConfig.BaseUrl)/assets" -Headers $ApiConfig.Headers -Method Post -Body $body -ErrorAction Stop
        
        Write-Log "✅ Successfully created asset: $($response.asset.display_id) - $($response.asset.name) (Type: $assetTypeName)"
        if ($computerAsset) {
            Write-Log "   Associated with computer: $($computerAsset.name)"
        }
        return $response.asset
        
    }
    catch {
        Write-Log "❌ Failed to create asset for $($Printer.Name): $($_.Exception.Message)"
        return $null
    }
}

# Main execution
Write-Log "🚀 Starting Freshservice Printer Asset Integration (Simplified)"
Write-Log "Computer: $env:COMPUTERNAME"

# Load configuration
$config = Load-Configuration -ConfigPath $ConfigFile
if (-not $config) {
    Write-Log "❌ Cannot proceed without configuration"
    exit 1
}

# Test connection
$apiConfig = Test-FreshserviceConnection -Config $config
if (-not $apiConfig) {
    Write-Log "❌ Cannot proceed without Freshservice connection"
    exit 1
}

if ($TestConnection) {
    Write-Log "✅ Connection test successful!"
    exit 0
}

# Get asset types from Freshservice
Write-Log "📋 Getting asset types from Freshservice..."
$assetTypes = Get-FreshserviceAssetTypes -ApiConfig $apiConfig
if ($assetTypes.Count -eq 0) {
    Write-Log "❌ Cannot proceed without asset types"
    exit 1
}

# Add asset types to config for easy lookup
$config | Add-Member -MemberType NoteProperty -Name "asset_types" -Value $assetTypes -Force

# Scan for printers
Write-Log "🔍 Scanning for printers..."
try {
    $printers = Get-PrinterInfoEnhanced
    Write-Log "✅ Using enhanced printer scanner"
} catch {
    Write-Log "⚠️ Enhanced scanner not available, using simple scanner"
    $printers = Get-PrinterInfoSimple
}

if ($printers.Count -eq 0) {
    Write-Log "⚠️ No printers found on this computer"
    exit 0
}

Write-Log "📊 Found $($printers.Count) printer(s)"

# Process each printer
$createdAssets = @()
$skippedAssets = @()

foreach ($printer in $printers) {
    Write-Log "🖨️ Processing printer: $($printer.Name)"
    
    $newAsset = New-FreshserviceAsset -ApiConfig $apiConfig -Config $config -Printer $printer
    if ($newAsset) {
        $createdAssets += $newAsset
    } else {
        $skippedAssets += $printer
    }
}

# Summary
Write-Log ""
Write-Log "📋 Integration Summary:"
Write-Log "   Created: $($createdAssets.Count) new assets"
Write-Log "   Skipped: $($skippedAssets.Count) printers"

if ($createdAssets.Count -gt 0) {
    Write-Log "✅ New assets created:"
    $createdAssets | ForEach-Object { Write-Log "   - $($_.display_id): $($_.name)" }
}

Write-Log "🎉 Freshservice integration completed!" 