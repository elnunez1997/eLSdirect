$ErrorActionPreference = 'Stop'

$OutputPath = Join-Path $PSScriptRoot 'RepayLetters_Pilot_Intake.xlsx'

$excel = New-Object -ComObject Excel.Application
$excel.DisplayAlerts = $false
$excel.Visible = $false

$workbook = $null
$headerSheet = $null
$lineItemsSheet = $null
$instructionsSheet = $null
$referenceSheet = $null

try {
    if (Test-Path $OutputPath) {
        try {
            Remove-Item $OutputPath -Force
        }
        catch {
            $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
            $OutputPath = Join-Path $PSScriptRoot ("RepayLetters_Pilot_Intake_$timestamp.xlsx")
        }
    }

    $workbook = $excel.Workbooks.Add()

    while ($workbook.Worksheets.Count -lt 4) {
        $null = $workbook.Worksheets.Add()
    }

    while ($workbook.Worksheets.Count -gt 4) {
        $workbook.Worksheets.Item($workbook.Worksheets.Count).Delete()
    }

    $headerSheet = $workbook.Worksheets.Item(1)
    $lineItemsSheet = $workbook.Worksheets.Item(2)
    $instructionsSheet = $workbook.Worksheets.Item(3)
    $referenceSheet = $workbook.Worksheets.Item(4)

    $headerSheet.Name = 'Header'
    $lineItemsSheet.Name = 'LineItems'
    $instructionsSheet.Name = 'Instructions'
    $referenceSheet.Name = 'ReferenceData'

    $headerColumns = @(
        'LetterType','LetterDate','CustomerName','CustomerAddr1','CustomerCity','CustomerState','CustomerZip',
        'ValidatorName','ValidatorLocation','ValidatorEmail','ValidatorPhone','ValidatorFax','ClientName',
        'AcostaClaimNumber','ClientReferenceNumber','RepayAmount','CheckNumber','CheckDate',
        'CustomerReferenceNumber','ClientInvoiceNumber','CustomerPONumber','OriginalDeductionAmount',
        'VendorNumber','ManufacturerName','BracketCasePounds','BracketLevel','GeneratePDF',
        'RequestedByEmail','SubmissionStatus'
    )

    for ($i = 0; $i -lt $headerColumns.Count; $i++) {
        $headerSheet.Cells.Item(1, $i + 1).Value2 = $headerColumns[$i]
    }

    $headerSheet.Cells.Item(2,1).Value2 = 'Wrong Vendor'
    $headerSheet.Cells.Item(2,2).Value2 = '2026-07-30'
    $headerSheet.Cells.Item(2,3).Value2 = 'Sample Customer'
    $headerSheet.Cells.Item(2,4).Value2 = '123 Main St'
    $headerSheet.Cells.Item(2,5).Value2 = 'Chicago'
    $headerSheet.Cells.Item(2,6).Value2 = 'IL'
    $headerSheet.Cells.Item(2,7).Value2 = '60601'
    $headerSheet.Cells.Item(2,8).Value2 = 'Jane Analyst'
    $headerSheet.Cells.Item(2,9).Value2 = 'Acosta Chicago'
    $headerSheet.Cells.Item(2,10).Value2 = 'jane.analyst@example.com'
    $headerSheet.Cells.Item(2,11).Value2 = '555-0100'
    $headerSheet.Cells.Item(2,12).Value2 = '555-0101'
    $headerSheet.Cells.Item(2,13).Value2 = 'Sample Client'
    $headerSheet.Cells.Item(2,14).Value2 = 'ACO-10001'
    $headerSheet.Cells.Item(2,15).Value2 = 'REF-7788'
    $headerSheet.Cells.Item(2,16).Value2 = 1250.00
    $headerSheet.Cells.Item(2,17).Value2 = 'CHK-90210'
    $headerSheet.Cells.Item(2,18).Value2 = '2026-07-25'
    $headerSheet.Cells.Item(2,19).Value2 = 'CUST-REF-1'
    $headerSheet.Cells.Item(2,20).Value2 = 'INV-4477'
    $headerSheet.Cells.Item(2,21).Value2 = 'PO-9988'
    $headerSheet.Cells.Item(2,22).Value2 = 1250.00
    $headerSheet.Cells.Item(2,23).Value2 = 'V-100'
    $headerSheet.Cells.Item(2,24).Value2 = 'Sample Manufacturer'
    $headerSheet.Cells.Item(2,25).Value2 = ''
    $headerSheet.Cells.Item(2,26).Value2 = ''
    $headerSheet.Cells.Item(2,27).Value2 = 'Yes'
    $headerSheet.Cells.Item(2,28).Value2 = 'jane.analyst@example.com'
    $headerSheet.Cells.Item(2,29).Value2 = 'Draft'

    $headerRange = $headerSheet.Range($headerSheet.Cells.Item(1,1), $headerSheet.Cells.Item(2,$headerColumns.Count))
    $headerTable = $headerSheet.ListObjects.Add(1, $headerRange, $null, 1)
    $headerTable.Name = 'HeaderTable'
    $headerTable.TableStyle = 'TableStyleMedium2'

    $lineColumns = @('UPC','CustomerCode','AllowancePerUnit','CaseQty','LineTotal')

    for ($i = 0; $i -lt $lineColumns.Count; $i++) {
        $lineItemsSheet.Cells.Item(1, $i + 1).Value2 = $lineColumns[$i]
    }

    $lineItemsSheet.Cells.Item(2,1).Value2 = '000123456789'
    $lineItemsSheet.Cells.Item(2,2).Value2 = 'CC-100'
    $lineItemsSheet.Cells.Item(2,3).Value2 = 12.5
    $lineItemsSheet.Cells.Item(2,4).Value2 = 10
    $lineItemsSheet.Cells.Item(2,5).Formula = '=C2*D2'

    $lineItemsSheet.Cells.Item(3,1).Value2 = '000987654321'
    $lineItemsSheet.Cells.Item(3,2).Value2 = 'CC-100'
    $lineItemsSheet.Cells.Item(3,3).Value2 = 8.75
    $lineItemsSheet.Cells.Item(3,4).Value2 = 6
    $lineItemsSheet.Cells.Item(3,5).Formula = '=C3*D3'

    $lineRange = $lineItemsSheet.Range($lineItemsSheet.Cells.Item(1,1), $lineItemsSheet.Cells.Item(3,$lineColumns.Count))
    $lineTable = $lineItemsSheet.ListObjects.Add(1, $lineRange, $null, 1)
    $lineTable.Name = 'LineItemsTable'
    $lineTable.TableStyle = 'TableStyleMedium2'
    $lineTable.ShowTotals = $true
    $lineItemsSheet.Cells.Item(4,5).Formula = '=SUM(E2:E3)'

    $instructionsSheet.Cells.Item(1,1).Value2 = 'How to Use'
    $instructionsSheet.Cells.Item(2,1).Value2 = '1. Select a LetterType in HeaderTable.'
    $instructionsSheet.Cells.Item(3,1).Value2 = '2. Complete all required fields.'
    $instructionsSheet.Cells.Item(4,1).Value2 = '3. For Allow Shorted Product, complete LineItemsTable.'
    $instructionsSheet.Cells.Item(5,1).Value2 = '4. Set SubmissionStatus to Ready before triggering the flow.'
    $instructionsSheet.Cells.Item(7,1).Value2 = 'Pilot Qualifying Questions'
    $instructionsSheet.Cells.Item(8,1).Value2 = 'Allow Shorted Product: Did client short product? Compare PO vs invoice. Will client accept this deduction?'
    $instructionsSheet.Cells.Item(9,1).Value2 = 'Bracket Pricing: Did customer request wrong bracket on PO? Is a Price Discrepancy Notice attached? Is bracket pricing confirmed?'
    $instructionsSheet.Cells.Item(10,1).Value2 = 'Wrong Vendor: Are products sold by this client? Was there a recent acquisition?'

    $referenceSheet.Cells.Item(1,1).Value2 = 'LetterType'
    $referenceSheet.Cells.Item(2,1).Value2 = 'Allow Shorted Product'
    $referenceSheet.Cells.Item(3,1).Value2 = 'Bracket Pricing'
    $referenceSheet.Cells.Item(4,1).Value2 = 'Wrong Vendor'
    $referenceSheet.Cells.Item(1,2).Value2 = 'GeneratePDF'
    $referenceSheet.Cells.Item(2,2).Value2 = 'Yes'
    $referenceSheet.Cells.Item(3,2).Value2 = 'No'
    $referenceSheet.Cells.Item(1,3).Value2 = 'SubmissionStatus'
    $referenceSheet.Cells.Item(2,3).Value2 = 'Draft'
    $referenceSheet.Cells.Item(3,3).Value2 = 'Ready'
    $referenceSheet.Cells.Item(4,3).Value2 = 'Processed'
    $referenceSheet.Cells.Item(5,3).Value2 = 'Error'

    $headerSheet.Rows.Item(1).Font.Bold = $true
    $lineItemsSheet.Rows.Item(1).Font.Bold = $true
    $instructionsSheet.Rows.Item(1).Font.Bold = $true
    $instructionsSheet.Rows.Item(7).Font.Bold = $true
    $referenceSheet.Rows.Item(1).Font.Bold = $true

    $headerSheet.Range('P2').NumberFormat = '$#,##0.00'
    $headerSheet.Range('V2').NumberFormat = '$#,##0.00'
    $lineItemsSheet.Range('C2:C3').NumberFormat = '$#,##0.00'
    $lineItemsSheet.Range('E2:E4').NumberFormat = '$#,##0.00'

    $headerSheet.Columns.AutoFit() | Out-Null
    $lineItemsSheet.Columns.AutoFit() | Out-Null
    $instructionsSheet.Columns.AutoFit() | Out-Null
    $referenceSheet.Columns.AutoFit() | Out-Null

    $workbook.SaveAs($OutputPath, 51)
    $workbook.Close($true)

    Write-Output "Created $OutputPath"
}
finally {
    if ($excel) {
        $excel.Quit()
    }

    if ($referenceSheet) {
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($referenceSheet) | Out-Null
    }
    if ($instructionsSheet) {
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($instructionsSheet) | Out-Null
    }
    if ($lineItemsSheet) {
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($lineItemsSheet) | Out-Null
    }
    if ($headerSheet) {
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($headerSheet) | Out-Null
    }
    if ($workbook) {
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null
    }
    if ($excel) {
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    }

    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
