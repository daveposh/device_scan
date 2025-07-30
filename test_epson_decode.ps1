# Test script for Epson serial number decoding
function Decode-HexSerial {
    param([string]$HexSerial)
    
    Write-Host "Input: $HexSerial"
    Write-Host "Length: $($HexSerial.Length)"
    
    try {
        if ($HexSerial.Length -eq 20 -and $HexSerial -match "^([A-F0-9]{8})([0-9]{6})(0{4})$") {
            Write-Host "Epson pattern matched!"
            $hexPart = $matches[1]  # First 8 hex characters
            $numberPart = $matches[2]  # Next 6 numbers
            $padding = $matches[3]  # Last 4 zeros
            
            Write-Host "Hex part: $hexPart"
            Write-Host "Number part: $numberPart"
            Write-Host "Padding: $padding"
            
            # Decode the hex part
            $hexDecoded = ""
            for ($i = 0; $i -lt $hexPart.Length; $i += 2) {
                $hexPair = $hexPart.Substring($i, 2)
                $byteValue = [Convert]::ToByte($hexPair, 16)
                Write-Host "  $hexPair = $byteValue = '$([char]$byteValue)'"
                if ($byteValue -ge 32 -and $byteValue -le 126) {
                    $hexDecoded += [char]$byteValue
                }
            }
            
            Write-Host "Hex decoded: '$hexDecoded'"
            
            # Combine hex decoded part with number part
            $epsonDecoded = $hexDecoded + $numberPart
            
            Write-Host "Final Epson decoded: '$epsonDecoded'"
            
            if ($epsonDecoded.Length -gt 0) {
                return "$HexSerial (Epson Decoded: $epsonDecoded)"
            }
        } else {
            Write-Host "Epson pattern NOT matched"
        }
        
        # Standard ASCII decoding (try this after Epson-specific decoding)
        Write-Host "Trying standard ASCII decoding..."
        $bytes = @()
        for ($i = 0; $i -lt $HexSerial.Length; $i += 2) {
            $bytes += [Convert]::ToByte($HexSerial.Substring($i, 2), 16)
        }
        $decodedSerial = [System.Text.Encoding]::ASCII.GetString($bytes).TrimEnd([char]0)
        
        Write-Host "Standard decoded: '$decodedSerial'"
        
        # Check if decoded result looks like a valid serial number
        if ($decodedSerial -match '^[A-Za-z0-9\-_]+$' -and $decodedSerial.Length -gt 0) {
            return "$HexSerial (Decoded: $decodedSerial)"
        }
        
        # If no valid decoding found, return original hex
        return $HexSerial
    } catch {
        Write-Host "Error: $($_.Exception.Message)"
        # Return original hex if decoding fails
        return $HexSerial
    }
}

# Test the Epson serial number
$testSerial = "5839584C1156290000"
Write-Host "=== Testing Epson Serial Number Decoding ==="
$result = Decode-HexSerial -HexSerial $testSerial
Write-Host "Final result: $result"

# Test the regex pattern
if ($testSerial -match "^([A-F0-9]{8})([0-9]{6})(0{4})$") {
    Write-Host "Regex match successful!"
    Write-Host "Hex part: $($matches[1])"
    Write-Host "Number part: $($matches[2])"
    Write-Host "Padding: $($matches[3])"
} else {
    Write-Host "Regex match failed!"
} 