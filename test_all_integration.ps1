# Comprehensive Freshservice Integration Test
# This script tests all components of the Freshservice integration

param(
    [Parameter(Mandatory=$false)]
    [string]$ConfigFile = "freshservice_config.json"
)

function Write-Log {
    param([string]$Message)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Write-Host "[$timestamp] $Message"
}

Write-Log "🧪 Starting Comprehensive Freshservice Integration Test"
Write-Log "Computer: $env:COMPUTERNAME"

# Test 1: Check if all required files exist
Write-Log ""
Write-Log "📁 Test 1: Checking required files..."
$requiredFiles = @(
    "freshservice_config.json",
    "freshservice_printer_asset_simple.ps1",
    "freshservice_printer_asset.ps1",
    "printer_scanner_enhanced.ps1",
    "add_printers_simple.bat",
    "add_printers_to_freshservice.bat"
)

$allFilesExist = $true
foreach ($file in $requiredFiles) {
    if (Test-Path $file) {
        Write-Log "   ✅ $file exists"
    } else {
        Write-Log "   ❌ $file missing"
        $allFilesExist = $false
    }
}

if (-not $allFilesExist) {
    Write-Log "❌ Some required files are missing"
    exit 1
}

# Test 2: Check configuration file
Write-Log ""
Write-Log "📋 Test 2: Checking configuration file..."
try {
    $config = Get-Content $ConfigFile -Raw | ConvertFrom-Json
    Write-Log "   ✅ Configuration file is valid JSON"
    
    if ($config.freshservice.domain -eq "yourcompany") {
        Write-Log "   ⚠️ Please update the domain in $ConfigFile"
    } else {
        Write-Log "   ✅ Domain configured: $($config.freshservice.domain)"
    }
    
    if ($config.freshservice.api_key -eq "your-api-key-here") {
        Write-Log "   ⚠️ Please update the API key in $ConfigFile"
    } else {
        Write-Log "   ✅ API key configured"
    }
    
} catch {
    Write-Log "   ❌ Configuration file is invalid: $($_.Exception.Message)"
    exit 1
}

# Test 3: Test printer scanner import
Write-Log ""
Write-Log "🔍 Test 3: Testing printer scanner import..."
try {
    . .\printer_scanner_enhanced.ps1
    Write-Log "   ✅ Enhanced printer scanner imported successfully"
    
    # Test if the main function exists
    if (Get-Command Get-PrinterInfoEnhanced -ErrorAction SilentlyContinue) {
        Write-Log "   ✅ Get-PrinterInfoEnhanced function available"
    } else {
        Write-Log "   ⚠️ Get-PrinterInfoEnhanced function not found"
    }
    
} catch {
    Write-Log "   ⚠️ Could not import enhanced printer scanner: $($_.Exception.Message)"
}

# Test 4: Test simple printer scanner
Write-Log ""
Write-Log "🖨️ Test 4: Testing simple printer scanner..."
try {
    . .\freshservice_printer_asset_simple.ps1
    Write-Log "   ✅ Simple integration script imported successfully"
    
    # Test if the fallback function exists
    if (Get-Command Get-PrinterInfoSimple -ErrorAction SilentlyContinue) {
        Write-Log "   ✅ Get-PrinterInfoSimple function available"
    } else {
        Write-Log "   ⚠️ Get-PrinterInfoSimple function not found"
    }
    
} catch {
    Write-Log "   ❌ Could not import simple integration script: $($_.Exception.Message)"
    exit 1
}

# Test 5: Test PowerShell syntax
Write-Log ""
Write-Log "🔧 Test 5: Testing PowerShell syntax..."
$scripts = @(
    "freshservice_printer_asset_simple.ps1",
    "freshservice_printer_asset.ps1"
)

foreach ($script in $scripts) {
    try {
        $null = [System.Management.Automation.PSParser]::Tokenize((Get-Content $script -Raw), [ref]$null)
        Write-Log "   ✅ $script syntax is valid"
    } catch {
        Write-Log "   ❌ $script has syntax errors: $($_.Exception.Message)"
    }
}

# Test 6: Test basic printer detection
Write-Log ""
Write-Log "🖨️ Test 6: Testing basic printer detection..."
try {
    $wmiPrinters = Get-WmiObject -Class Win32_Printer -ErrorAction SilentlyContinue
    Write-Log "   ✅ WMI printer query successful"
    Write-Log "   📊 Found $($wmiPrinters.Count) printers via WMI"
    
    $usbDevices = Get-WmiObject -Class Win32_PnPEntity -ErrorAction SilentlyContinue | 
                  Where-Object { $_.DeviceID -like "*USB*" }
    Write-Log "   ✅ USB device query successful"
    Write-Log "   📊 Found $($usbDevices.Count) USB devices"
    
} catch {
    Write-Log "   ❌ Basic printer detection failed: $($_.Exception.Message)"
}

# Test 7: Test Freshservice API connection (if configured)
Write-Log ""
Write-Log "🌐 Test 7: Testing Freshservice API connection..."
if ($config.freshservice.domain -ne "yourcompany" -and $config.freshservice.api_key -ne "your-api-key-here") {
    try {
        $BaseUrl = "https://$($config.freshservice.domain).freshservice.com/api/v2"
        $Headers = @{
            "Content-Type" = "application/json"
            "Authorization" = "Basic " + [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes("$($config.freshservice.api_key)`:"))
        }
        
        $response = Invoke-RestMethod -Uri "$BaseUrl/assets" -Headers $Headers -Method Get -ErrorAction Stop
        Write-Log "   ✅ Freshservice API connection successful"
        
        # Test asset types endpoint
        $assetTypesResponse = Invoke-RestMethod -Uri "$BaseUrl/asset_types" -Headers $Headers -Method Get -ErrorAction Stop
        Write-Log "   ✅ Asset types endpoint accessible"
        Write-Log "   📊 Found $($assetTypesResponse.asset_types.Count) asset types"
        
    } catch {
        Write-Log "   ❌ Freshservice API connection failed: $($_.Exception.Message)"
    }
} else {
    Write-Log "   ⚠️ Skipping API test - please configure domain and API key first"
}

Write-Log ""
Write-Log "🎉 Comprehensive test completed!"
Write-Log ""
Write-Log "📋 Next Steps:"
Write-Log "   1. Update freshservice_config.json with your domain and API key"
Write-Log "   2. Run: add_printers_simple.bat"
Write-Log "   3. Or test connection: .\freshservice_printer_asset_simple.ps1 -TestConnection"
Write-Log "   4. Or dry run: .\freshservice_printer_asset_simple.ps1 -DryRun" 