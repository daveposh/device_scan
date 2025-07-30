# Test script for Epson regex pattern

$testSerial = "5839584C1156290000"

Write-Host "Testing Epson serial: $testSerial"

# Test the regex pattern
if ($testSerial -match "^([A-F0-9]{8})([0-9]{6})(0{4})$") {
    Write-Host "Regex match successful!"
    Write-Host "Hex part: $($matches[1])"
    Write-Host "Number part: $($matches[2])"
    Write-Host "Padding: $($matches[3])"
    
    $hexPart = $matches[1]
    $numberPart = $matches[2]
    
    # Decode the hex part
    $hexDecoded = ""
    for ($i = 0; $i -lt $hexPart.Length; $i += 2) {
        $hexPair = $hexPart.Substring($i, 2)
        $byteValue = [Convert]::ToByte($hexPair, 16)
        Write-Host "Hex pair: $hexPair = $byteValue = [char]$byteValue"
        if ($byteValue -ge 32 -and $byteValue -le 126) {
            $hexDecoded += [char]$byteValue
        }
    }
    
    Write-Host "Hex decoded: $hexDecoded"
    Write-Host "Number part: $numberPart"
    Write-Host "Final result: $hexDecoded$numberPart"
    
} else {
    Write-Host "Regex match failed!"
} 