# Test Freshservice API Integration
# This script tests the Freshservice API connection and asset creation

param(
    [Parameter(Mandatory=$false)]
    [string]$ConfigFile = "freshservice_config.json",
    
    [Parameter(Mandatory=$false)]
    [switch]$TestAssetCreation
)

function Write-Log {
    param([string]$Message)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Write-Host "[$timestamp] $Message"
}

function Load-Configuration {
    param([string]$ConfigPath)
    
    try {
        if (-not (Test-Path $ConfigPath)) {
            Write-Log "❌ Configuration file not found: $ConfigPath"
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

function Test-AssetCreation {
    param($ApiConfig, $AssetTypes)
    
    try {
        # Create a test asset
        $testAssetData = @{
            asset_type_id = ($AssetTypes | Select-Object -First 1).id
            name = "Test Printer Asset - $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
            description = "Test asset created by API integration script"
            asset_tag = "TEST-$(Get-Random -Minimum 1000 -Maximum 9999)"
            serial_number = "TEST-SERIAL-$(Get-Random -Minimum 1000 -Maximum 9999)"
            manufacturer = "Test Manufacturer"
            model = "Test Model"
            custom_fields = @{
                test_field = "Test Value"
                computer_name = $env:COMPUTERNAME
                test_date = (Get-Date -Format "yyyy-MM-dd")
            }
        }
        
        $body = $testAssetData | ConvertTo-Json -Depth 3
        Write-Log "🔍 Testing asset creation with data:"
        Write-Log $body
        
        $response = Invoke-RestMethod -Uri "$($ApiConfig.BaseUrl)/assets" -Headers $ApiConfig.Headers -Method Post -Body $body -ErrorAction Stop
        
        Write-Log "✅ Successfully created test asset: $($response.asset.display_id) - $($response.asset.name)"
        
        # Clean up - delete the test asset
        Write-Log "🧹 Cleaning up test asset..."
        Invoke-RestMethod -Uri "$($ApiConfig.BaseUrl)/assets/$($response.asset.id)" -Headers $ApiConfig.Headers -Method Delete -ErrorAction Stop
        Write-Log "✅ Test asset deleted successfully"
        
        return $true
    }
    catch {
        Write-Log "❌ Failed to test asset creation: $($_.Exception.Message)"
        return $false
    }
}

# Main execution
Write-Log "🧪 Starting Freshservice API Integration Test"
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

# Get asset types
Write-Log "📋 Getting asset types from Freshservice..."
$assetTypes = Get-FreshserviceAssetTypes -ApiConfig $apiConfig
if ($assetTypes.Count -eq 0) {
    Write-Log "❌ Cannot proceed without asset types"
    exit 1
}

# Display available asset types
Write-Log "📋 Available asset types:"
$assetTypes | ForEach-Object { Write-Log "   - $($_.name) (ID: $($_.id))" }

# Test asset creation if requested
if ($TestAssetCreation) {
    Write-Log "🧪 Testing asset creation..."
    $success = Test-AssetCreation -ApiConfig $apiConfig -AssetTypes $assetTypes
    if ($success) {
        Write-Log "✅ Asset creation test passed!"
    } else {
        Write-Log "❌ Asset creation test failed!"
        exit 1
    }
}

Write-Log "🎉 Freshservice API integration test completed successfully!" 