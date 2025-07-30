# Test the regex pattern for Epson serial number
$testSerial = "5839584C1156290000"

Write-Host "Testing serial: '$testSerial'"
Write-Host "Length: $($testSerial.Length)"
Write-Host "Characters:"

# Show each character and its ASCII value
for ($i = 0; $i -lt $testSerial.Length; $i++) {
    $char = $testSerial[$i]
    $ascii = [int]$char
    Write-Host "  Position $i : '$char' (ASCII: $ascii)"
}

Write-Host "`nTesting regex patterns:"

# Test different regex patterns
$patterns = @(
    "^([A-Fa-f0-9]{8})([0-9]{6})(0{4})$",
    "^([A-Fa-f0-9]{8})([0-9]{6})(0{4})$",
    "^([A-Fa-f0-9]{8})([0-9]{6})(0{4})$"
)

foreach ($pattern in $patterns) {
    Write-Host "Testing pattern: $pattern"
    if ($testSerial -match $pattern) {
        Write-Host "  MATCHED!"
        Write-Host "  Hex part: '$($matches[1])'"
        Write-Host "  Number part: '$($matches[2])'"
        Write-Host "  Padding: '$($matches[3])'"
    } else {
        Write-Host "  NOT MATCHED"
    }
}

# Test individual parts
Write-Host "`nTesting individual parts:"
$first8 = $testSerial.Substring(0, 8)
$next6 = $testSerial.Substring(8, 6)
$last4 = $testSerial.Substring(14, 4)

Write-Host "First 8: '$first8' - Is hex: $($first8 -match '^[A-Fa-f0-9]{8}$')"
Write-Host "Next 6: '$next6' - Is numbers: $($next6 -match '^[0-9]{6}$')"
Write-Host "Last 4: '$last4' - Is zeros: $($last4 -match '^0{4}$')" 