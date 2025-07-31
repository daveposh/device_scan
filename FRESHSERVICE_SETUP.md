# Freshservice Printer Asset Integration Setup Guide

This guide will help you set up and use the Freshservice printer asset integration system.

## 📋 Prerequisites

- Windows computer with PowerShell
- Freshservice account with API access
- Freshservice API key

## 🚀 Quick Start

### 1. Configure Freshservice Settings

Edit `freshservice_config.json` with your Freshservice domain and API key:

```json
{
    "freshservice": {
        "domain": "yourcompany",
        "api_key": "your-api-key-here",
        "default_asset_type": "Printer"
    }
}
```

### 2. Get Your Freshservice API Key

1. Log into your Freshservice instance
2. Go to Admin → API Settings
3. Generate a new API key
4. Copy the key to the config file

### 3. Test the Integration

Run the comprehensive test to verify everything is working:

```batch
test_all.bat
```

### 4. Run the Integration

```batch
add_printers_simple.bat
```

## 📁 File Overview

### Core Integration Files

- **`freshservice_printer_asset_simple.ps1`** - Main integration script (simplified)
- **`freshservice_printer_asset.ps1`** - Full-featured integration script
- **`freshservice_config.json`** - Configuration file
- **`printer_scanner_enhanced.ps1`** - Enhanced printer detection

### Batch Files

- **`add_printers_simple.bat`** - Simple integration runner
- **`add_printers_to_freshservice.bat`** - Advanced integration runner
- **`test_all.bat`** - Comprehensive test runner
- **`test_freshservice.bat`** - API test runner

### Test Files

- **`test_all_integration.ps1`** - Comprehensive integration test
- **`test_freshservice_api.ps1`** - API connectivity test

## 🎛️ Usage Options

### Simple Method (Recommended)

```batch
add_printers_simple.bat
```

### Advanced Method

```batch
add_printers_to_freshservice.bat yourcompany your-api-key "Label Printer" "Main Office" "IT"
```

### PowerShell Direct

```powershell
# Test connection
.\freshservice_printer_asset_simple.ps1 -TestConnection

# Dry run (preview)
.\freshservice_printer_asset_simple.ps1 -DryRun

# Full integration
.\freshservice_printer_asset_simple.ps1
```

## 🔍 Testing

### Comprehensive Test

```batch
test_all.bat
```

This tests:
- ✅ File existence
- ✅ Configuration validity
- ✅ Printer scanner import
- ✅ PowerShell syntax
- ✅ Basic printer detection
- ✅ Freshservice API connectivity

### API Test Only

```batch
test_freshservice.bat
```

## 🖨️ Printer Detection

The system detects printers using multiple methods:

1. **Enhanced Scanner** (preferred)
   - WMI queries
   - PnP device enumeration
   - Registry scanning
   - Print Management API

2. **Simple Scanner** (fallback)
   - Basic WMI printer queries
   - USB device detection

### Supported Printers

- **Epson Label Printers** (TM-T88, TM-T, TM-U, etc.)
- **HP Printers**
- **Canon Printers**
- **Brother Printers**
- **Network Printers** (with IP detection)

## 📊 Asset Information Captured

Each printer asset includes:

- **Basic Info**: Name, Description, Asset Tag, Serial Number
- **Hardware**: Manufacturer, Model
- **Connection**: Port Name, Driver Name, Device ID
- **Network**: IP Address, Network Protocol (for network printers)
- **Discovery**: Source, Status, Computer Name, Discovery Date
- **Computer Association**: Linked to computer asset as child asset

## 🔗 Computer Asset Association

Printer assets are automatically associated with the computer they're connected to:

- **Parent-Child Relationship**: Printers become child assets of the computer
- **Hostname Matching**: Searches for computer assets by hostname
- **Asset Type Detection**: Looks for Computer, Laptop, or Desktop asset types

## 🛡️ Error Handling

The system includes comprehensive error handling:

- **Graceful Import**: Fallback if enhanced scanner not available
- **API Error Handling**: Proper error messages for API failures
- **Configuration Validation**: Checks for required settings
- **Duplicate Detection**: Prevents creating duplicate assets

## 🔧 Troubleshooting

### Common Issues

1. **"Configuration file not found"**
   - Ensure `freshservice_config.json` exists in the same directory

2. **"Failed to connect to Freshservice API"**
   - Check your domain and API key in the config file
   - Verify your Freshservice instance is accessible

3. **"No printers found"**
   - Run `test_all.bat` to verify printer detection is working
   - Check if printers are properly connected and recognized by Windows

4. **"Enhanced scanner not available"**
   - The system will automatically fall back to simple scanner
   - This is normal and expected if `printer_scanner_enhanced.ps1` is not available

### Debug Mode

For detailed logging, run:

```powershell
.\freshservice_printer_asset_simple.ps1 -Verbose
```

## 📈 Asset Hierarchy in Freshservice

The integration creates a clean hierarchy:

```
📁 Computer Asset (DESKTOP-ABC123)
   └── 🏷️ Epson TM-T88 Receipt Printer
   └── 🖨️ HP LaserJet Pro
   └── 🖨️ Brother Network Printer
```

## 🎯 Epson Label Printer Support

Special handling for Epson label printers:

- **Serial Number Decoding**: Hex-encoded serials properly decoded
- **USB Controller Detection**: Links USB controllers to virtual COM ports
- **Asset Type**: Automatically categorized as "Label Printer"
- **Computer Association**: Linked to computer asset as child assets

## 📝 Customization

### Asset Types

Edit `freshservice_config.json` to customize asset types:

```json
"printer_types": {
    "epson": {
        "keywords": ["Epson", "TM-T88", "TM-T"],
        "asset_type": "Label Printer"
    },
    "custom": {
        "keywords": ["YourBrand"],
        "asset_type": "Custom Printer"
    }
}
```

### Custom Fields

The system captures printer-specific information in custom fields:

- Port Name
- Driver Name
- Device ID
- Discovery Source
- IP Address
- Network Protocol
- Computer Name
- Discovery Date

## 🚀 Deployment

### PDQ Deploy

For enterprise deployment via PDQ Deploy:

1. Upload all files to your PDQ repository
2. Create a package with command:
   ```
   add_printers_simple.bat
   ```
3. Deploy to target computers

### Manual Deployment

1. Copy all files to target computer
2. Configure `freshservice_config.json`
3. Run `add_printers_simple.bat`

## 📞 Support

If you encounter issues:

1. Run `test_all.bat` to identify the problem
2. Check the error messages in the console output
3. Verify your Freshservice configuration
4. Ensure printers are properly connected and recognized

## 🎉 Success Indicators

When the integration runs successfully, you should see:

- ✅ "Successfully connected to Freshservice API"
- ✅ "Using enhanced printer scanner" or "Using simple scanner"
- ✅ "Found X printer(s)"
- ✅ "Successfully created asset: ASSET-XXX - Printer Name"
- ✅ "Associated with computer: ComputerName"
- ✅ "Freshservice integration completed!" 