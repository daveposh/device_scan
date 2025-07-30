# Test script for hex decoding function
# Test with the Epson serial number from the scan results

# Function to decode hex serial numbers (including Epson format)
function Decode-HexSerial {
    param([string]$HexSerial)
    
    try {
        if ($HexSerial.Length -ge 8 -and $HexSerial.Length % 2 -eq 0) {
            # Standard ASCII decoding
            $bytes = @()
            for ($i = 0; $i -lt $HexSerial.Length; $i += 2) {
                $bytes += [Convert]::ToByte($HexSerial.Substring($i, 2), 16)
            }
            $decodedSerial = [System.Text.Encoding]::ASCII.GetString($bytes).TrimEnd([char]0)
            
            # Check if decoded result looks like a valid serial number
            if ($decodedSerial -match '^[A-Za-z0-9\-_]+$' -and $decodedSerial.Length -gt 0) {
                return "$HexSerial (Decoded: $decodedSerial)"
            }
            
            # Special handling for Epson-style hex encoding (pairs of hex digits)
            if ($HexSerial.Length -ge 16 -and $HexSerial.Length % 2 -eq 0) {
                $epsonDecoded = ""
                for ($i = 0; $i -lt $HexSerial.Length; $i += 2) {
                    $hexPair = $HexSerial.Substring($i, 2)
                    $byteValue = [Convert]::ToByte($hexPair, 16)
                    # Only include non-zero bytes (Epson often pads with zeros)
                    if ($byteValue -ne 0) {
                        $epsonDecoded += [char]$byteValue
                    }
                }
                if ($epsonDecoded -match '^[A-Za-z0-9\-_]+$' -and $epsonDecoded.Length -gt 0) {
                    return "$HexSerial (Epson Decoded: $epsonDecoded)"
                }
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