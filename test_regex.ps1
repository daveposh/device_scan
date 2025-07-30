# Test the regex pattern for Epson serial number
$testSerial = "5839584C1156290000"

Write-Host "Testing serial: $testSerial"
Write-Host "Length: $($testSerial.Length)"

# Test the regex pattern
if ($testSerial -match "^([A-Fa-f0-9]{8})([0-9]{6})(0{4})$") {
    Write-Host "Regex MATCHED!"
    Write-Host "Hex part: $($matches[1])"
    Write-Host "Number part: $($matches[2])"
    Write-Host "Padding: $($matches[3])"
    
    # Decode the hex part
    $hexPart = $matches[1]
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
    $finalResult = $hexDecoded + $matches[2]
    Write-Host "Final result: '$finalResult'"
} else {
    Write-Host "Regex did NOT match!"
} 