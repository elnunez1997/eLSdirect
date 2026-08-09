# Daily Task Tracker v2 - Input-driven with Dropdowns, Progress Bars & Charts

$OutputPath = "$PSScriptRoot\Daily_Task_Tracker_v2.xlsx"

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$wb = $excel.Workbooks.Add()

# ── Color helpers ────────────────────────────────────────────
function RGB([int]$r,[int]$g,[int]$b) { return $r + ($g * 256) + ($b * 65536) }

$cDark    = RGB 31 35 40
$cWhite   = RGB 255 255 255
$cBlue    = RGB 59 130 212
$cPurple  = RGB 124 92 216
$cRed     = RGB 207 34 46
$cAmber   = RGB 154 103 0
$cGreen   = RGB 26 127 55
$cGray    = RGB 247 248 250
$cMuted   = RGB 87 96 106
$cRedBG   = RGB 255 240 240
$cAmberBG = RGB 255 248 224
$cGreenBG = RGB 230 244 234
$cBlueBG  = RGB 238 242 255
$cPurBG   = RGB 243 232 255
$cYellow  = RGB 255 214 0
$cOrange  = RGB 255 140 0

$xlCenter  = -4108
$xlLeft    = -4131
$xlRight   = -4152
$xlVCenter = -4108

$dateStr = (Get-Date).ToString("dddd, MMMM dd yyyy")

# ── Set cell helper ──────────────────────────────────────────
function SetCell($ws,$r,$c,$val,$bold=$false,$bg=$null,$fg=$null,$ha=$null,$wrap=$false,$sz=$null,$va=$null,$border=$false,$italic=$false,$numfmt=$null) {
    $cell = $ws.Cells.Item($r,$c)
    if ($null -ne $val) { $cell.Value2 = $val }
    if ($bold)   { $cell.Font.Bold   = $true }
    if ($italic) { $cell.Font.Italic = $true }
    if ($sz)     { $cell.Font.Size   = $sz }
    if ($bg)     { $cell.Interior.Color = $bg }
    if ($fg)     { $cell.Font.Color   = $fg }
    if ($wrap)   { $cell.WrapText     = $true }
    if ($ha)     { $cell.HorizontalAlignment = $ha }
    if ($va)     { $cell.VerticalAlignment   = $va }
    if ($numfmt) { $cell.NumberFormat  = $numfmt }
    if ($border) {
        $b = $cell.Borders
        $b.Item(7).LineStyle = 1; $b.Item(7).Weight = 2   # left
        $b.Item(8).LineStyle = 1; $b.Item(8).Weight = 2   # top
        $b.Item(9).LineStyle = 1; $b.Item(9).Weight = 2   # bottom
        $b.Item(10).LineStyle = 1; $b.Item(10).Weight = 2 # right
    }
}

function Merge($ws,$r1,$c1,$r2,$c2) {
    $ws.Range($ws.Cells.Item($r1,$c1), $ws.Cells.Item($r2,$c2)).Merge() | Out-Null
}

function ColLetter([int]$n) {
    if ($n -le 26) { return [char](64+$n) }
    return [char](64+[int][math]::Floor(($n-1)/26)) + [char](64+(($n-1)%26)+1)
}

# ============================================================
# SHEET 1 — TASK INPUT (Editable)
# ============================================================
$ws1 = $wb.Worksheets.Item(1)
$ws1.Name = "TASK INPUT"
$ws1.Tab.Color = $cDark
$ws1.Rows.Item(1).RowHeight = 40
$ws1.Rows.Item(2).RowHeight = 20
$ws1.Rows.Item(3).RowHeight = 16
$ws1.Rows.Item(4).RowHeight = 22
$ws1.Rows.Item(5).RowHeight = 16

# Column widths:  A=#  B=Task  C=Role  D=Priority  E=Due  F=%Done  G=ProgBar  H=Status  I=Notes
$ws1.Columns.Item(1).ColumnWidth  = 4
$ws1.Columns.Item(2).ColumnWidth  = 46
$ws1.Columns.Item(3).ColumnWidth  = 14
$ws1.Columns.Item(4).ColumnWidth  = 18
$ws1.Columns.Item(5).ColumnWidth  = 12
$ws1.Columns.Item(6).ColumnWidth  = 10
$ws1.Columns.Item(7).ColumnWidth  = 20
$ws1.Columns.Item(8).ColumnWidth  = 16
$ws1.Columns.Item(9).ColumnWidth  = 30

# ── Title ────────────────────────────────────────────────────
Merge $ws1 1 1 1 9
SC $ws1 1 1 "DAILY TASK TRACKER  v2  |  Deduction Analyst + Reporting Analyst" $true $cDark $cWhite $xlLeft $false 15

Merge $ws1 2 1 2 9
SC $ws1 2 1 "Date: $dateStr  |  Fill in your tasks below. Use dropdowns for Role, Priority and Status. Update % Done (0-100) to drive progress bars." $false $cBlueBG $cBlue $xlLeft $false 9

# ── Legend row ───────────────────────────────────────────────
Merge $ws1 3 1 3 9
SC $ws1 3 1 "  LEGEND:   P1 = Urgent/Important  |  P2 = Important  |  P3 = Routine  |  Status drives colour coding  |  % Done drives the progress bar" $false $cGray $cMuted $xlLeft $false 9

# ── Column Headers ───────────────────────────────────────────
$hdrs = @("#","TASK / ACTIVITY DESCRIPTION","ROLE","PRIORITY","DUE TIME","% DONE","PROGRESS BAR","STATUS","NOTES / REMARKS")
for ($c=1; $c -le 9; $c++) {
    SC $ws1 4 $c $hdrs[$c-1] $true $cDark $cWhite $xlCenter $false 9 $null $true
}
$ws1.Rows.Item(4).RowHeight = 22

# ── 30 editable input rows (rows 6–35) ──────────────────────
$inputStart = 6
$inputEnd   = 35
$totalInputRows = $inputEnd - $inputStart + 1

for ($r = $inputStart; $r -le $inputEnd; $r++) {
    $ws1.Rows.Item($r).RowHeight = 22
    $rowBG = if (($r % 2) -eq 0) { $cGray } else { $cWhite }
    $num = $r - $inputStart + 1

    # Col A: row number (static)
    SC $ws1 $r 1 $num $false $rowBG $cMuted $xlCenter

    # Col B: Task description — editable, left-aligned
    SC $ws1 $r 2 "" $false $cWhite $cDark $xlLeft $true
    $ws1.Cells.Item($r,2).Interior.Color = $cWhite

    # Col C: Role dropdown — Deduction / Reporting / Both
    SC $ws1 $r 3 "" $false $cWhite $cBlue $xlCenter

    # Col D: Priority dropdown
    SC $ws1 $r 4 "" $false $cWhite $cRed $xlCenter

    # Col E: Due time — free text
    SC $ws1 $r 5 "" $false $cWhite $cMuted $xlCenter

    # Col F: % Done — number 0–100
    SC $ws1 $r 6 0 $false $cWhite $cDark $xlCenter $false $null $null $false $false "0"

    # Col G: Progress bar formula — REPT bar chars (visual bar)
    $pctCell = "F$r"
    $barFormula = "=IFERROR(REPT(CHAR(9608),ROUND($pctCell/10,0))&REPT(CHAR(9617),10-ROUND($pctCell/10,0)),REPT(CHAR(9617),10))"
    $ws1.Cells.Item($r,7).Formula = $barFormula
    $ws1.Cells.Item($r,7).Font.Name = "Courier New"
    $ws1.Cells.Item($r,7).Font.Size = 11
    $ws1.Cells.Item($r,7).Interior.Color = $cWhite
    $ws1.Cells.Item($r,7).HorizontalAlignment = $xlLeft
    $ws1.Cells.Item($r,7).Locked = $true

    # Col H: Status dropdown
    SC $ws1 $r 8 "" $false $cWhite $cGreen $xlCenter

    # Col I: Notes — free text
    SC $ws1 $r 9 "" $false $cWhite $cMuted $xlLeft $true
}

# ── Data Validation dropdowns ────────────────────────────────
$roleRange     = $ws1.Range("C${inputStart}:C${inputEnd}")
$priorityRange = $ws1.Range("D${inputStart}:D${inputEnd}")
$statusRange   = $ws1.Range("H${inputStart}:H${inputEnd}")
$pctRange      = $ws1.Range("F${inputStart}:F${inputEnd}")

# Role
$dv = $roleRange.Validation
$dv.Delete()
$dv.Add(3, 1, 1, "Deduction,Reporting,Both") | Out-Null  # xlValidateList=3, xlValidAlertStop=1
$dv.InCellDropdown = $true
$dv.ShowInput = $true
$dv.InputTitle = "Role"
$dv.InputMessage = "Select the analyst role for this task"

# Priority
$dv2 = $priorityRange.Validation
$dv2.Delete()
$dv2.Add(3, 1, 1, "P1 - Urgent,P2 - Important,P3 - Routine") | Out-Null
$dv2.InCellDropdown = $true
$dv2.ShowInput = $true
$dv2.InputTitle = "Priority"
$dv2.InputMessage = "P1=Urgent/Important  P2=Important  P3=Routine"

# Status
$dv3 = $statusRange.Validation
$dv3.Delete()
$dv3.Add(3, 1, 1, "Pending,In Progress,Partial,Complete") | Out-Null
$dv3.InCellDropdown = $true
$dv3.ShowInput = $true
$dv3.InputTitle = "Status"
$dv3.InputMessage = "Select the current status of this task"

# % Done: 0-100 integer only
$dv4 = $pctRange.Validation
$dv4.Delete()
$dv4.Add(1, 1, 1, "0", "100") | Out-Null  # xlValidateWholeNumber=1, between 0 and 100
$dv4.ShowInput = $true
$dv4.InputTitle = "% Done"
$dv4.InputMessage = "Enter a whole number from 0 to 100"
$dv4.ShowError = $true
$dv4.ErrorTitle = "Invalid Value"
$dv4.ErrorMessage = "Please enter a whole number between 0 and 100"

# ── Conditional Formatting: colour Status cells ──────────────
# Colour progress bar text by % Done
$barCFRange = $ws1.Range("G${inputStart}:G${inputEnd}")
# Green when F col = 100
$cf1 = $barCFRange.FormatConditions.Add(2, $null, "=F${inputStart}=100")  # xlExpression=2
$cf1.Font.Color = $cGreen
# Amber when 50-99
$cf2 = $barCFRange.FormatConditions.Add(2, $null, "=AND(F${inputStart}>=50,F${inputStart}<100)")
$cf2.Font.Color = $cAmber
# Red when < 50
$cf3 = $barCFRange.FormatConditions.Add(2, $null, "=F${inputStart}<50")
$cf3.Font.Color = $cRed

# ── % Done data bar CF ───────────────────────────────────────
$pctCFRange = $ws1.Range("F${inputStart}:F${inputEnd}")
try {
    $cfBar = $pctCFRange.FormatConditions.AddDatabar()
    $cfBar.ShowValue = $true
    $cfBar.BarColor.Color = $cBlue
} catch {}

# ── Colour-code Status column by value ──────────────────────
$stRange = $ws1.Range("H${inputStart}:H${inputEnd}")
# Complete = green
$cs1 = $stRange.FormatConditions.Add(1, 3, "Complete")  # xlCellValue=1, xlEqual=3
$cs1.Interior.Color = $cGreenBG; $cs1.Font.Color = $cGreen; $cs1.Font.Bold = $true
# Partial = amber
$cs2 = $stRange.FormatConditions.Add(1, 3, "Partial")
$cs2.Interior.Color = $cAmberBG; $cs2.Font.Color = $cAmber; $cs2.Font.Bold = $true
# In Progress = blue
$cs3 = $stRange.FormatConditions.Add(1, 3, "In Progress")
$cs3.Interior.Color = $cBlueBG; $cs3.Font.Color = $cBlue; $cs3.Font.Bold = $true
# Pending = red
$cs4 = $stRange.FormatConditions.Add(1, 3, "Pending")
$cs4.Interior.Color = $cRedBG; $cs4.Font.Color = $cRed; $cs4.Font.Bold = $true

# ── Colour-code Priority column ──────────────────────────────
$prRange = $ws1.Range("D${inputStart}:D${inputEnd}")
$cp1 = $prRange.FormatConditions.Add(1, 3, "P1 - Urgent")
$cp1.Interior.Color = $cRedBG; $cp1.Font.Color = $cRed; $cp1.Font.Bold = $true
$cp2 = $prRange.FormatConditions.Add(1, 3, "P2 - Important")
$cp2.Interior.Color = $cAmberBG; $cp2.Font.Color = $cAmber; $cp2.Font.Bold = $true
$cp3 = $prRange.FormatConditions.Add(1, 3, "P3 - Routine")
$cp3.Interior.Color = $cGreenBG; $cp3.Font.Color = $cGreen; $cp3.Font.Bold = $true

# ── Summary footer area (row 37-42) ─────────────────────────
$sumRow = $inputEnd + 2
Merge $ws1 $sumRow 1 $sumRow 9
SC $ws1 $sumRow 1 "SUMMARY (Auto-calculated from your inputs above)" $true $cDark $cWhite $xlLeft $false 10
$ws1.Rows.Item($sumRow).RowHeight = 22

$sr = $sumRow + 1
$lbls  = @("Total Tasks","% Done Avg","Complete","Partial","In Progress","Pending")
$fmlas = @(
    "=COUNTA(B${inputStart}:B${inputEnd})",
    "=IFERROR(AVERAGE(F${inputStart}:F${inputEnd}),0)",
    "=COUNTIF(H${inputStart}:H${inputEnd},""Complete"")",
    "=COUNTIF(H${inputStart}:H${inputEnd},""Partial"")",
    "=COUNTIF(H${inputStart}:H${inputEnd},""In Progress"")",
    "=COUNTIF(H${inputStart}:H${inputEnd},""Pending"")"
)
$sumFGs = @($cDark,$cBlue,$cGreen,$cAmber,$cBlue,$cRed)
$sumBGs = @($cGray,$cBlueBG,$cGreenBG,$cAmberBG,$cBlueBG,$cRedBG)

for ($i=0; $i -lt 6; $i++) {
    $col = ($i*1)+2   # cols B,C,D,E,F,G mapped as 2,3,4,5,6,7... use pairs
}
# Use two-column pairs: label + value
$pairs = @(
    @(2,3), @(4,5), @(6,7), @(2,3)   # reuse, layout below
)
# Layout: row sr = labels, row sr+1 = values — across columns B-G (2-7)
for ($i=0; $i -lt 6; $i++) {
    $col = $i + 2  # cols 2..7 (B..G)
    SC $ws1 $sr     $col $lbls[$i]  $true  $sumBGs[$i] $sumFGs[$i] $xlCenter $false 8
    $ws1.Rows.Item($sr).RowHeight = 18
    $cell = $ws1.Cells.Item($sr+1, $col)
    $cell.Formula = $fmlas[$i]
    $cell.Font.Bold = $true
    $cell.Font.Size = 16
    $cell.Interior.Color = $sumBGs[$i]
    $cell.Font.Color = $sumFGs[$i]
    $cell.HorizontalAlignment = $xlCenter
    if ($i -eq 1) { $cell.NumberFormat = "0.0" }
    $ws1.Rows.Item($sr+1).RowHeight = 28
}

# ── Named ranges for chart source data ──────────────────────
# Status counts in a hidden helper range: cols K-N, row 2-5
$ws1.Cells.Item(2,11).Value2 = "Status"
$ws1.Cells.Item(2,12).Value2 = "Count"
$statuses = @("Complete","Partial","In Progress","Pending")
for ($i=0;$i -lt 4;$i++) {
    $row = 3+$i
    $ws1.Cells.Item($row,11).Value2 = $statuses[$i]
    $ws1.Cells.Item($row,12).Formula = "=COUNTIF(H${inputStart}:H${inputEnd},K$row)"
}

# Priority counts: cols K-N, rows 8-10
$ws1.Cells.Item(8,11).Value2 = "Priority"
$ws1.Cells.Item(8,12).Value2 = "Count"
$priorities = @("P1 - Urgent","P2 - Important","P3 - Routine")
for ($i=0;$i -lt 3;$i++) {
    $row = 9+$i
    $ws1.Cells.Item($row,11).Value2 = $priorities[$i]
    $ws1.Cells.Item($row,12).Formula = "=COUNTIF(D${inputStart}:D${inputEnd},K$row)"
}

# Role counts: cols K-N, rows 13-15
$ws1.Cells.Item(13,11).Value2 = "Role"
$ws1.Cells.Item(13,12).Value2 = "Count"
$roles = @("Deduction","Reporting","Both")
for ($i=0;$i -lt 3;$i++) {
    $row = 14+$i
    $ws1.Cells.Item($row,11).Value2 = $roles[$i]
    $ws1.Cells.Item($row,12).Formula = "=COUNTIF(C${inputStart}:C${inputEnd},K$row)"
}

# Avg % Done per priority: rows 18-20
$ws1.Cells.Item(18,11).Value2 = "Priority"
$ws1.Cells.Item(18,12).Value2 = "Avg % Done"
for ($i=0;$i -lt 3;$i++) {
    $row = 19+$i
    $ws1.Cells.Item($row,11).Value2 = $priorities[$i]
    $ws1.Cells.Item($row,12).Formula = "=IFERROR(AVERAGEIF(D${inputStart}:D${inputEnd},K$row,F${inputStart}:F${inputEnd}),0)"
}

# Hide helper columns K-L
$ws1.Columns.Item(11).Hidden = $true
$ws1.Columns.Item(12).Hidden = $true

# ── Freeze panes + AutoFilter ────────────────────────────────
$ws1.Activate()
$ws1.Cells.Item($inputStart,1).Select()
$excel.ActiveWindow.FreezePanes = $true
$ws1.Rows.Item(4).AutoFilter() | Out-Null

# ── Sheet protection: lock formula col G, allow editing rest ─
$ws1.Cells.Item(1,1).Select()

# ============================================================
# SHEET 2 — CHARTS DASHBOARD
# ============================================================
$ws2 = $wb.Worksheets.Add([System.Reflection.Missing]::Value, $ws1)
$ws2.Name = "CHARTS"
$ws2.Tab.Color = $cBlue

$ws2.Columns.Item(1).ColumnWidth = 1

# Title
Merge $ws2 1 1 1 10
SC $ws2 1 1 "CHARTS DASHBOARD  |  Auto-updates from TASK INPUT sheet" $true $cDark $cWhite $xlLeft $false 14
$ws2.Rows.Item(1).RowHeight = 36

Merge $ws2 2 1 2 10
SC $ws2 2 1 "All charts refresh automatically as you add/update tasks in the TASK INPUT sheet." $false $cBlueBG $cBlue $xlLeft $false 10
$ws2.Rows.Item(2).RowHeight = 20

# ── Helper to add a chart ────────────────────────────────────
function Add-Chart {
    param($ws,$srcWs,$dataAddr,$chartTitle,$chartType,$left,$top,$width,$height,$colors)
    $co = $ws.Shapes.AddChart2(-1, $chartType, $left, $top, $width, $height)
    $chart = $co.Chart
    $chart.HasTitle = $true
    $chart.ChartTitle.Text = $chartTitle
    $chart.ChartTitle.Font.Size = 11
    $chart.ChartTitle.Font.Bold = $true

    $chart.SetSourceData($srcWs.Range($dataAddr))
    $chart.PlotBy = 2  # xlColumns

    # Style
    $chart.ChartArea.Format.Fill.ForeColor.RGB = $cWhite
    $chart.PlotArea.Format.Fill.ForeColor.RGB = $cGray
    $chart.HasLegend = $true
    $chart.Legend.Font.Size = 9

    try {
        $series = $chart.SeriesCollection(1)
        if ($colors) {
            for ($i=0;$i -lt $colors.Count;$i++) {
                try { $chart.SeriesCollection(1).Points($i+1).Format.Fill.ForeColor.RGB = $colors[$i] } catch {}
            }
        }
    } catch {}
    return $chart
}

# xlChartType constants
$xlPie         = 5
$xlBarClustered = 57
$xlColumnClustered = 51
$xlDoughnut    = -4120
$xlLine        = 4
$xlBarStacked  = 58

# Row px positions (approx: row height ~15px each)
$chartTop1  = 60
$chartTop2  = 320
$chartLeft1 = 20
$chartLeft2 = 380
$chartLeft3 = 740
$cW = 340
$cH = 240

# Chart 1: Status Breakdown — Donut
try {
    $co1 = $ws2.Shapes.AddChart2(-1, $xlDoughnut, $chartLeft1, $chartTop1, $cW, $cH)
    $ch1 = $co1.Chart
    $ch1.SetSourceData($ws1.Range("K3:L6"))
    $ch1.HasTitle = $true
    $ch1.ChartTitle.Text = "Task Status Breakdown"
    $ch1.ChartTitle.Font.Size = 11
    $ch1.ChartTitle.Font.Bold = $true
    $ch1.HasLegend = $true
    $ch1.Legend.Font.Size = 9
    $ch1.ChartArea.Format.Fill.ForeColor.RGB = $cWhite
    try {
        $ch1.SeriesCollection(1).Points(1).Format.Fill.ForeColor.RGB = $cGreen
        $ch1.SeriesCollection(1).Points(2).Format.Fill.ForeColor.RGB = $cAmber
        $ch1.SeriesCollection(1).Points(3).Format.Fill.ForeColor.RGB = $cBlue
        $ch1.SeriesCollection(1).Points(4).Format.Fill.ForeColor.RGB = $cRed
    } catch {}
} catch { Write-Host "Chart1 error: $_" }

# Chart 2: Priority Distribution — Column
try {
    $co2 = $ws2.Shapes.AddChart2(-1, $xlColumnClustered, $chartLeft2, $chartTop1, $cW, $cH)
    $ch2 = $co2.Chart
    $ch2.SetSourceData($ws1.Range("K9:L11"))
    $ch2.HasTitle = $true
    $ch2.ChartTitle.Text = "Tasks by Priority"
    $ch2.ChartTitle.Font.Size = 11
    $ch2.ChartTitle.Font.Bold = $true
    $ch2.HasLegend = $false
    $ch2.ChartArea.Format.Fill.ForeColor.RGB = $cWhite
    try {
        $ch2.SeriesCollection(1).Points(1).Format.Fill.ForeColor.RGB = $cRed
        $ch2.SeriesCollection(1).Points(2).Format.Fill.ForeColor.RGB = $cAmber
        $ch2.SeriesCollection(1).Points(3).Format.Fill.ForeColor.RGB = $cGreen
    } catch {}
} catch { Write-Host "Chart2 error: $_" }

# Chart 3: Role Distribution — Pie
try {
    $co3 = $ws2.Shapes.AddChart2(-1, $xlPie, $chartLeft3, $chartTop1, $cW, $cH)
    $ch3 = $co3.Chart
    $ch3.SetSourceData($ws1.Range("K14:L16"))
    $ch3.HasTitle = $true
    $ch3.ChartTitle.Text = "Tasks by Role"
    $ch3.ChartTitle.Font.Size = 11
    $ch3.ChartTitle.Font.Bold = $true
    $ch3.HasLegend = $true
    $ch3.Legend.Font.Size = 9
    $ch3.ChartArea.Format.Fill.ForeColor.RGB = $cWhite
    try {
        $ch3.SeriesCollection(1).Points(1).Format.Fill.ForeColor.RGB = $cBlue
        $ch3.SeriesCollection(1).Points(2).Format.Fill.ForeColor.RGB = $cPurple
        $ch3.SeriesCollection(1).Points(3).Format.Fill.ForeColor.RGB = $cGreen
    } catch {}
} catch { Write-Host "Chart3 error: $_" }

# Chart 4: Avg % Done by Priority — Bar
try {
    $co4 = $ws2.Shapes.AddChart2(-1, $xlBarClustered, $chartLeft1, $chartTop2, $cW, $cH)
    $ch4 = $co4.Chart
    $ch4.SetSourceData($ws1.Range("K19:L21"))
    $ch4.HasTitle = $true
    $ch4.ChartTitle.Text = "Avg % Completion by Priority"
    $ch4.ChartTitle.Font.Size = 11
    $ch4.ChartTitle.Font.Bold = $true
    $ch4.HasLegend = $false
    $ch4.ChartArea.Format.Fill.ForeColor.RGB = $cWhite
    try {
        $ch4.SeriesCollection(1).Points(1).Format.Fill.ForeColor.RGB = $cRed
        $ch4.SeriesCollection(1).Points(2).Format.Fill.ForeColor.RGB = $cAmber
        $ch4.SeriesCollection(1).Points(3).Format.Fill.ForeColor.RGB = $cGreen
    } catch {}
    # Set axis max to 100
    try { $ch4.Axes(2).MaximumScale = 100 } catch {}
} catch { Write-Host "Chart4 error: $_" }

# Chart 5: Complete vs Pending vs In Progress — Stacked Bar
try {
    $co5 = $ws2.Shapes.AddChart2(-1, $xlColumnClustered, $chartLeft2, $chartTop2, $cW*2+$chartLeft3-$chartLeft2-$cW, $cH)
    $ch5 = $co5.Chart
    $ch5.SetSourceData($ws1.Range("K3:L6"))
    $ch5.ChartType = $xlBarStacked
    $ch5.HasTitle = $true
    $ch5.ChartTitle.Text = "Status Count Overview"
    $ch5.ChartTitle.Font.Size = 11
    $ch5.ChartTitle.Font.Bold = $true
    $ch5.HasLegend = $true
    $ch5.ChartArea.Format.Fill.ForeColor.RGB = $cWhite
} catch { Write-Host "Chart5 error: $_" }

# ── Instructions panel on Charts sheet ──────────────────────
$instrRow = 36
Merge $ws2 $instrRow 1 $instrRow 10
SC $ws2 $instrRow 1 "HOW TO USE THIS TRACKER" $true $cDark $cWhite $xlLeft $false 11
$ws2.Rows.Item($instrRow).RowHeight = 24

$steps = @(
    "1. Go to the TASK INPUT sheet (tab at bottom).",
    "2. In column B, type your task description.",
    "3. Use the dropdown in column C to select Role: Deduction / Reporting / Both.",
    "4. Use the dropdown in column D to select Priority: P1-Urgent / P2-Important / P3-Routine.",
    "5. Enter Due Time in column E (e.g. 10:00 AM, EOD, 3:00 PM).",
    "6. Enter a number 0-100 in column F (% Done). The progress bar in column G updates automatically.",
    "7. Use the dropdown in column H to set Status: Pending / In Progress / Partial / Complete.",
    "8. Add any notes in column I.",
    "9. Come back to this CHARTS sheet - all 5 charts update live as you fill in tasks."
)
for ($i=0;$i -lt $steps.Count;$i++) {
    $rr = $instrRow+1+$i
    Merge $ws2 $rr 1 $rr 10
    $bg2 = if ($i % 2 -eq 0) { $cWhite } else { $cGray }
    SC $ws2 $rr 1 $steps[$i] $false $bg2 $cDark $xlLeft $false 10
    $ws2.Rows.Item($rr).RowHeight = 20
}

# ============================================================
# SHEET 3 — CATEGORIES (Lookup / Reference)
# ============================================================
$ws3 = $wb.Worksheets.Add([System.Reflection.Missing]::Value, $ws2)
$ws3.Name = "CATEGORIES"
$ws3.Tab.Color = $cPurple

$ws3.Columns.Item(1).ColumnWidth = 28
$ws3.Columns.Item(2).ColumnWidth = 50
$ws3.Columns.Item(3).ColumnWidth = 20

Merge $ws3 1 1 1 3
SC $ws3 1 1 "TASK CATEGORIES REFERENCE  |  Use as guide when entering tasks" $true $cPurple $cWhite $xlLeft $false 13
$ws3.Rows.Item(1).RowHeight = 32

Merge $ws3 2 1 2 3
SC $ws3 2 1 "Date: $dateStr" $false $cPurBG $cPurple $xlLeft $false 10
$ws3.Rows.Item(2).RowHeight = 18

$catHdrs = @("CATEGORY","EXAMPLE TASKS","ROLE")
for ($c=1;$c -le 3;$c++) {
    SC $ws3 4 $c $catHdrs[$c-1] $true $cDark $cWhite $xlCenter $false 9
}
$ws3.Rows.Item(4).RowHeight = 20

$cats = @(
    @("Deduction Resolution","Resolve escalated deductions / dispute filings / backup validation","Deduction"),
    @("ARCollect Updates","Activity lookup updates / Customer Name field / Activity Specific Fields","Deduction"),
    @("Aging Review","Review deductions over 60/90 days / prioritize for resolution","Deduction"),
    @("Short-Pay Reconciliation","Match customer short-pays against promotions / contracts","Deduction"),
    @("Dispute Filing","File formal disputes with backup / document in ARCollect","Deduction"),
    @("Client Follow-up","Follow up on pending deductions / await client response","Deduction"),
    @("Daily Aging Report","Compile and submit daily aging report to management","Reporting"),
    @("KPI Dashboard Update","Update daily KPI metrics - resolution rate / amounts cleared","Reporting"),
    @("Escalation Report","Prepare escalation report for management-flagged items","Reporting"),
    @("Trend Analysis","Weekly deduction pattern analysis by client or category","Reporting"),
    @("SLA Compliance","Track SLA deadlines / identify at-risk items","Reporting"),
    @("Root Cause Analysis","Identify and document top deduction reasons","Reporting"),
    @("Dashboard Maintenance","Update Power BI or Excel dashboards with latest data","Reporting"),
    @("Archive & Filing","Archive reports / file documents in SharePoint","Both"),
    @("Process Documentation","Document SOPs / process notes for team","Both"),
    @("Team Coordination","Status meetings / handoff notes / team updates","Both")
)

for ($i=0;$i -lt $cats.Count;$i++) {
    $r = 5+$i
    $ws3.Rows.Item($r).RowHeight = 20
    $rowBG3 = if ($i % 2 -eq 0) { $cWhite } else { $cGray }
    $roleC3 = switch ($cats[$i][2]) { "Deduction"{$cBlue} "Reporting"{$cPurple} default{$cGreen} }
    $roleBG3 = switch ($cats[$i][2]) { "Deduction"{$cBlueBG} "Reporting"{$cPurBG} default{$cGreenBG} }
    SC $ws3 $r 1 $cats[$i][0] $true  $rowBG3 $cDark $xlLeft $false 9
    SC $ws3 $r 2 $cats[$i][1] $false $rowBG3 $cMuted $xlLeft $true 9
    SC $ws3 $r 3 $cats[$i][2] $true  $roleBG3 $roleC3 $xlCenter $false 9
}

# ── Activate TASK INPUT and save ─────────────────────────────
$ws1.Activate()
$ws1.Cells.Item($inputStart, 2).Select()

$wb.SaveAs($OutputPath, 51)
$wb.Close($false)
$excel.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null

Write-Host "SUCCESS: File saved -> $OutputPath"
