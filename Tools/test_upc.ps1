function Extract-UPCs {
    param([string]$line)
    $step1 = [regex]::Replace($line, '\b\d{7,8}\b', '')
    $step2 = [regex]::Replace($step1, '\$?\d{1,3}(,\d{3})*\.\d+', '')
    $found = [regex]::Matches($step2, '\b\d{5}\b')
    return ($found | ForEach-Object { $_.Value } | Select-Object -Unique)
}

$tests = @(
    'The customer was already given the allowance, except for UPCs 11042 and 25225.',
    'UPC (1188, 11042) and UPC (25305) not received on invoice number 73817061.',
    'UPC 01058, 01060, 01186, 01189, 01201, and 25305 not received on invoice 73848920.',
    'UPC RATE 1058 188.00 1060 188.00 25225 663.00 11042 62.40',
    'Invoice# 73946508 shows UPC: 00511, 00521, 00525, 25660, 25743',
    'The above deduction is for an off-invoice allowance of $4.70, $2.60, $2.32 for UPC 01058 01060 01186 01188 not received on invoice number 73949602.'
)

foreach ($t in $tests) {
    $upcs = Extract-UPCs $t
    Write-Host "INPUT : $t"
    if ($upcs) {
        Write-Host "UPCS  : $($upcs -join ', ')"
    } else {
        Write-Host "UPCS  : (none found)"
    }
    Write-Host ""
}
