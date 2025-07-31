# Freshservice Printer Asset Integration Script
# This script scans for printers and adds them as assets in Freshservice

param(
    [Parameter(Mandatory=$true)]
    [string]$FreshserviceDomain,
    
    [Parameter(Mandatory=$true)]
    [string]$ApiKey,
    
    [Parameter(Mandatory=$false)]
    [string]$AssetType = "Printer",
    
    [Parameter(Mandatory=$false)]
    [string]$Location = "",
    
    [Parameter(Mandatory=$false)]
    [string]$Department = "",
    
    [Parameter(Mandatory=$false)]
    [switch]$DryRun
)

# Import the printer scanner functions
. .\printer_scanner_enhanced.ps1

# Freshservice API Configuration
$BaseUrl = "https://$FreshserviceDomain.freshservice.com/api/v2"
$Headers = @{
    "Content-Type" = "application/json"
    "Authorization" = "Basic " + [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes("$ApiKey`:"))
}

function Write-Log {
    param([string]$Message)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Write-Host "[$timestamp] $Message"
}

function Test-FreshserviceConnection {
    try {
        $response = Invoke-RestMethod -Uri "$BaseUrl/assets" -Headers $Headers -Method Get -ErrorAction Stop
        Write-Log "✅ Successfully connected to Freshservice API"
        return $true
    }
    catch {
        Write-Log "❌ Failed to connect to Freshservice API: $($_.Exception.Message)"
        return $false
    }
}

function Get-FreshserviceAssetTypes {
    try {
        $response = Invoke-RestMethod -Uri "$BaseUrl/asset_types" -Headers $Headers -Method Get -ErrorAction Stop
        return $response.asset_types
    }
    catch {
        Write-Log "❌ Failed to get asset types: $($_.Exception.Message)"
        return @()
    }
}

function Get-FreshserviceLocations {
    try {
        $response = Invoke-RestMethod -Uri "$BaseUrl/locations" -Headers $Headers -Method Get -ErrorAction Stop
        return $response.locations
    }
    catch {
        Write-Log "❌ Failed to get locations: $($_.Exception.Message)"
        return @()
    }
}

function Get-FreshserviceDepartments {
    try {
        $response = Invoke-RestMethod -Uri "$BaseUrl/departments" -Headers $Headers -Method Get -ErrorAction Stop
        return $response.departments
    }
    catch {
        Write-Log "❌ Failed to get departments: $($_.Exception.Message)"
        return @()
    }
}

function Find-ExistingAsset {
    param(
        [string]$SerialNumber,
        [string]$Name
    )
    
    try {
        # Search by serial number first
        if ($SerialNumber) {
            $response = Invoke-RestMethod -Uri "$BaseUrl/assets?search_string=$SerialNumber" -Headers $Headers -Method Get -ErrorAction Stop
            if ($response.assets.Count -gt 0) {
                return $response.assets[0]
            }
        }
        
        # Search by name if serial number not found
        if ($Name) {
            $response = Invoke-RestMethod -Uri "$BaseUrl/assets?search_string=$Name" -Headers $Headers -Method Get -ErrorAction Stop
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
        [string]$Hostname
    )
    
    try {
        # Search for computer asset by hostname
        $response = Invoke-RestMethod -Uri "$BaseUrl/assets?search_string=$Hostname" -Headers $Headers -Method Get -ErrorAction Stop
        
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

function New-FreshserviceAsset {
    param(
        [object]$Printer
    )
    
    try {
        # Check if asset already exists
        $existingAsset = Find-ExistingAsset -SerialNumber $Printer.SerialNumber -Name $Printer.Name
        if ($existingAsset) {
            Write-Log "⚠️ Asset already exists: $($existingAsset.display_id) - $($existingAsset.name)"
            return $existingAsset
        }
        
        # Find the computer asset to associate with
        $computerAsset = Find-ComputerAsset -Hostname $env:COMPUTERNAME
        if ($computerAsset) {
            Write-Log "🖥️ Found computer asset: $($computerAsset.display_id) - $($computerAsset.name)"
        } else {
            Write-Log "⚠️ Computer asset not found for hostname: $env:COMPUTERNAME"
        }
        
        # Prepare asset data
        $assetData = @{
            asset_type_id = $AssetTypeId
            name = $Printer.Name
            description = "Auto-discovered printer from device scan on $env:COMPUTERNAME"
            asset_tag = if ($Printer.SerialNumber) { $Printer.SerialNumber } else { "PRN-$(Get-Random -Minimum 1000 -Maximum 9999)" }
            serial_number = $Printer.SerialNumber
            manufacturer = $Printer.Manufacturer
            model = $Printer.Model
            location_id = $LocationId
            department_id = $DepartmentId
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
            Write-Log "🔍 DRY RUN: Would create asset for $($Printer.Name)"
            if ($computerAsset) {
                Write-Log "   Associated with computer: $($computerAsset.name) (ID: $($computerAsset.id))"
            }
            Write-Log "   Asset Data: $($assetData | ConvertTo-Json -Depth 3)"
            return $null
        }
        
        $body = $assetData | ConvertTo-Json -Depth 3
        $response = Invoke-RestMethod -Uri "$BaseUrl/assets" -Headers $Headers -Method Post -Body $body -ErrorAction Stop
        
        Write-Log "✅ Successfully created asset: $($response.asset.display_id) - $($response.asset.name)"
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

function Update-FreshserviceAsset {
    param(
        [object]$Asset,
        [object]$Printer
    )
    
    try {
        # Find the computer asset to associate with
        $computerAsset = Find-ComputerAsset -Hostname $env:COMPUTERNAME
        
        $updateData = @{
            name = $Printer.Name
            description = "Auto-discovered printer from device scan on $env:COMPUTERNAME (Updated)"
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
                last_updated = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
            }
        }
        
        # Associate with computer asset if found and not already associated
        if ($computerAsset -and -not $Asset.parent_asset_id) {
            $updateData.parent_asset_id = $computerAsset.id
        }
        
        if ($DryRun) {
            Write-Log "🔍 DRY RUN: Would update asset $($Asset.display_id) - $($Asset.name)"
            if ($computerAsset -and -not $Asset.parent_asset_id) {
                Write-Log "   Would associate with computer: $($computerAsset.name)"
            }
            return $Asset
        }
        
        $body = $updateData | ConvertTo-Json -Depth 3
        $response = Invoke-RestMethod -Uri "$BaseUrl/assets/$($Asset.id)" -Headers $Headers -Method Put -Body $body -ErrorAction Stop
        
        Write-Log "✅ Successfully updated asset: $($response.asset.display_id) - $($response.asset.name)"
        if ($computerAsset -and -not $Asset.parent_asset_id) {
            Write-Log "   Associated with computer: $($computerAsset.name)"
        }
        return $response.asset
        
    }
    catch {
        Write-Log "❌ Failed to update asset $($Asset.display_id): $($_.Exception.Message)"
        return $Asset
    }
}

# Main execution
Write-Log "🚀 Starting Freshservice Printer Asset Integration"
Write-Log "Computer: $env:COMPUTERNAME"
Write-Log "Freshservice Domain: $FreshserviceDomain"

# Test connection
if (-not (Test-FreshserviceConnection)) {
    Write-Log "❌ Cannot proceed without Freshservice connection"
    exit 1
}

# Get asset type ID
Write-Log "📋 Getting asset types..."
$assetTypes = Get-FreshserviceAssetTypes
$AssetTypeId = ($assetTypes | Where-Object { $_.name -eq $AssetType }).id
if (-not $AssetTypeId) {
    Write-Log "⚠️ Asset type '$AssetType' not found. Available types:"
    $assetTypes | ForEach-Object { Write-Log "   - $($_.name) (ID: $($_.id))" }
    $AssetTypeId = ($assetTypes | Select-Object -First 1).id
    Write-Log "📋 Using first available asset type: ID $AssetTypeId"
}

# Get location ID if specified
$LocationId = $null
if ($Location) {
    Write-Log "📍 Getting locations..."
    $locations = Get-FreshserviceLocations
    $LocationId = ($locations | Where-Object { $_.name -eq $Location }).id
    if (-not $LocationId) {
        Write-Log "⚠️ Location '$Location' not found. Available locations:"
        $locations | ForEach-Object { Write-Log "   - $($_.name) (ID: $($_.id))" }
    }
}

# Get department ID if specified
$DepartmentId = $null
if ($Department) {
    Write-Log "🏢 Getting departments..."
    $departments = Get-FreshserviceDepartments
    $DepartmentId = ($departments | Where-Object { $_.name -eq $Department }).id
    if (-not $DepartmentId) {
        Write-Log "⚠️ Department '$Department' not found. Available departments:"
        $departments | ForEach-Object { Write-Log "   - $($_.name) (ID: $($_.id))" }
    }
}

# Scan for printers
Write-Log "🔍 Scanning for printers..."
$printers = Get-PrinterInfoEnhanced

if ($printers.Count -eq 0) {
    Write-Log "⚠️ No printers found on this computer"
    exit 0
}

Write-Log "📊 Found $($printers.Count) printer(s)"

# Process each printer
$createdAssets = @()
$updatedAssets = @()
$skippedAssets = @()

foreach ($printer in $printers) {
    Write-Log "🖨️ Processing printer: $($printer.Name)"
    
    # Check if asset already exists
    $existingAsset = Find-ExistingAsset -SerialNumber $printer.SerialNumber -Name $printer.Name
    
    if ($existingAsset) {
        Write-Log "📝 Updating existing asset: $($existingAsset.display_id)"
        $updatedAsset = Update-FreshserviceAsset -Asset $existingAsset -Printer $printer
        if ($updatedAsset) {
            $updatedAssets += $updatedAsset
        }
    } else {
        Write-Log "➕ Creating new asset for: $($printer.Name)"
        $newAsset = New-FreshserviceAsset -Printer $printer
        if ($newAsset) {
            $createdAssets += $newAsset
        } else {
            $skippedAssets += $printer
        }
    }
}

# Summary
Write-Log ""
Write-Log "📋 Integration Summary:"
Write-Log "   Created: $($createdAssets.Count) new assets"
Write-Log "   Updated: $($updatedAssets.Count) existing assets"
Write-Log "   Skipped: $($skippedAssets.Count) printers"

if ($createdAssets.Count -gt 0) {
    Write-Log "✅ New assets created:"
    $createdAssets | ForEach-Object { Write-Log "   - $($_.display_id): $($_.name)" }
}

if ($updatedAssets.Count -gt 0) {
    Write-Log "📝 Assets updated:"
    $updatedAssets | ForEach-Object { Write-Log "   - $($_.display_id): $($_.name)" }
}

Write-Log "🎉 Freshservice integration completed!" 