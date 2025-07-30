# Test script for Epson serial number decoding

# Function to decode hex serial numbers (including Epson format)
function Decode-HexSerial {
    param([string]$HexSerial)
    
    try {
        if ($HexSerial.Length -ge 8 -and $HexSerial.Length % 2 -eq 0) {
            # Special handling for Epson-style hex encoding: [8 hex chars][6 numbers][4 zero padding]
            if ($HexSerial.Length -eq 20 -and $HexSerial -match "^([A-F0-9]{8})([0-9]{6})(0000)$") {
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
        }
    } catch {
        # Return original hex if decoding fails
        return $HexSerial
    }
    
    return $HexSerial
}

# Test the Epson serial number from the scan results
$testSerial = "5839584C1156290000"
Write-Host "Testing hex decoding for: $testSerial"
$result = Decode-HexSerial -HexSerial $testSerial
Write-Host "Result: $result"

# Test a few more examples
$testCases = @(
    "5839584C1156290000",
    "41424344",
    "4142434400000000",
    "1234567890ABCDEF"
)

Write-Host "`nTesting multiple cases:"
foreach ($test in $testCases) {
    $result = Decode-HexSerial -HexSerial $test
    Write-Host "$test -> $result"
} 