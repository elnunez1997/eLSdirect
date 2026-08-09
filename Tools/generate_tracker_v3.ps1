# Daily Task Tracker v3 - Structured Headers, Priority Input, Progress Bars, Comments

$OutputPath = "$PSScriptRoot\Daily_Task_Tracker_v3.xlsx"

$excel = New-Object -ComObject Excel.Application
$excel.Visible  = $false
$excel.DisplayAlerts = $false
$wb = $excel.Workbooks.Add()

# Remove extra sheets until only 1 remains
while ($wb.Worksheets.Count -gt 1) {
    $wb.Worksheets.Item($wb.Worksheets.Count).Delete()
}

# ── Colors ───────────────────────────────────────────────────
function RGB([int]$r,[int]$g,[int]$b){ return $r+($g*256)+($b*65536) }

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
$cDarkBlu = RGB 30 64 120
$cDarkRed = RGB 120 20 20
$cSlate   = RGB 51 65 85
$cLBlue   = RGB 219 234 254
$cLPur    = RGB 237 233 254

$xlLeft   = -4131
$xlCenter = -4108
$xlRight  = -4152
$xlTop    = -4160
$xlVCenter= -4108
$xlBottom = -4107

$dateStr  = (Get-Date).ToString("dddd, MMMM dd, yyyy")
$timeStr  = (Get-Date).ToString("hh:mm tt")

# ── Cell helper ──────────────────────────────────────────────
function W($ws,$r,$c,$v,$bold=$false,$bg=$null,$fg=$null,$ha=$null,$sz=$null,$wrap=$false,$va=$null,$nb=$false,$italic=$false,$fmt=$null){
    $cell = $ws.Cells.Item($r,$c)
    if ($null -ne $v) { try { $cell.Value2 = $v } catch { $cell.Value2 = [string]$v } }
    if ($bold)   { $cell.Font.Bold   = $true }
    if ($italic) { $cell.Font.Italic = $true }
    if ($sz)     { $cell.Font.Size   = $sz }
    if ($bg)     { $cell.Interior.Color = $bg }
    if ($fg)     { $cell.Font.Color   = $fg }
    if ($ha)     { $cell.HorizontalAlignment = $ha }
    if ($va)     { $cell.VerticalAlignment   = $va }
    if ($wrap)   { $cell.WrapText     = $true }
    if ($fmt)    { $cell.NumberFormat  = $fmt }
    if ($nb) {
        1..4 | ForEach-Object { $cell.Borders.Item($_+6).LineStyle=1; $cell.Borders.Item($_+6).Weight=2 }
    }
}

function MR($ws,$r1,$c1,$r2,$c2){
    $ws.Range($ws.Cells.Item($r1,$c1),$ws.Cells.Item($r2,$c2)).Merge() | Out-Null
}

function RH($ws,$r,$h){ $ws.Rows.Item($r).RowHeight = $h }
function CW($ws,$c,$w){ $ws.Columns.Item($c).ColumnWidth = $w }

# ── Add validation dropdown ──────────────────────────────────
function AddDV($rng,$list,$title,$msg){
    $dv = $rng.Validation
    $dv.Delete()
    $dv.Add(3,1,1,$list) | Out-Null
    $dv.InCellDropdown = $true
    $dv.ShowInput = $true
    $dv.InputTitle   = $title
    $dv.InputMessage = $msg
}

# ── CF colour by cell value ──────────────────────────────────
function CF($rng,$val,$bgColor,$fgColor){
    $rule = $rng.FormatConditions.Add(1,3,$val)  # xlCellValue=1, xlEqual=3
    $rule.Interior.Color = $bgColor
    $rule.Font.Color     = $fgColor
    $rule.Font.Bold      = $true
}

# ╔══════════════════════════════════════════════════════════╗
# ║  SHEET 1 — MAIN TRACKER                                  ║
# ╚══════════════════════════════════════════════════════════╝
$ws = $wb.Worksheets.Item(1)
$ws.Name = "DAILY TRACKER"
$ws.Tab.Color = $cDark
try { $excel.ActiveWindow.DisplayGridlines = $false } catch {}

# Column layout:
# A(1)=# B(2)=Task  C(3)=Role  D(4)=Priority  E(5)=Due  F(6)=%Done  G(7)=ProgBar  H(8)=Status  I(9)=Comments
CW $ws 1  4    # #
CW $ws 2  46   # Task
CW $ws 3  14   # Role
CW $ws 4  18   # Priority
CW $ws 5  12   # Due
CW $ws 6  10   # % Done
CW $ws 7  22   # Progress Bar
CW $ws 8  16   # Status
CW $ws 9  38   # Comments
CW $ws 10  2   # spacer

# ══════════════════════════════════════════════════════════
# ROW 1  — MAIN TITLE BANNER
# ══════════════════════════════════════════════════════════
RH $ws 1 38
MR $ws 1 1 1 7
W $ws 1 1 "DAILY TASK TRACKER" $true $cDark $cWhite $xlLeft $false 18
$ws.Cells.Item(1,1).IndentLevel = 1

MR $ws 1 8 1 9
W $ws 1 8 "Date: $dateStr" $false $cDark $cMuted $xlRight $false 10

# ══════════════════════════════════════════════════════════
# ROW 2  — ROLE BADGES + SUBTITLE
# ══════════════════════════════════════════════════════════
RH $ws 2 22
MR $ws 2 1 2 4
W $ws 2 1 "  Deduction Analyst  +  Reporting Analyst  |  Combined View" $false $cSlate $cWhite $xlLeft $false 9

W $ws 2 5 "DEDUCTION" $true $cBlueBG $cBlue $xlCenter $false 9
W $ws 2 6 "" $false $cBlueBG $null $xlCenter
W $ws 2 7 "" $false $cBlueBG $null $xlCenter

W $ws 2 8 "REPORTING" $true $cPurBG $cPurple $xlCenter $false 9
W $ws 2 9 "" $false $cPurBG $null $xlCenter

MR $ws 2 5 2 7
MR $ws 2 8 2 9

# ══════════════════════════════════════════════════════════
# ROWS 3-5  — MASTER PROGRESS DASHBOARD
# ══════════════════════════════════════════════════════════
RH $ws 3 8   # thin separator
$ws.Rows.Item(3).Interior.Color = $cSlate

RH $ws 4 26
RH $ws 5 30

# Section label
MR $ws 4 1 4 9
W $ws 4 1 "  OVERALL PROGRESS DASHBOARD" $true $cSlate $cWhite $xlLeft $false 9

# Summary cards row 5: Total | Complete | Partial | In Progress | Pending | Avg %Done
# Spans: A-B=Total, C=Complete, D=Partial, E=InProgress, F=Pending, G-H=Avg%Done, I=OverallBar
$cardData = @(
    @(1,2,"TOTAL TASKS",    '=COUNTA(B12:B200)',             $cDark,   $cWhite),
    @(3,3,"COMPLETE",       '=COUNTIF(H12:H200,"Complete")', $cGreen,  $cWhite),
    @(4,4,"PARTIAL",        '=COUNTIF(H12:H200,"Partial")',  $cAmber,  $cWhite),
    @(5,5,"IN PROGRESS",    '=COUNTIF(H12:H200,"In Progress")',$cBlue, $cWhite),
    @(6,6,"PENDING",        '=COUNTIF(H12:H200,"Pending")',  $cRed,    $cWhite),
    @(7,8,"AVG % DONE",     '=IFERROR(AVERAGE(F12:F200),0)', $cSlate,  $cWhite)
)

foreach ($cd in $cardData) {
    $c1=$cd[0]; $c2=$cd[1]; $lbl=$cd[2]; $fml=$cd[3]; $bg=$cd[4]; $fg=$cd[5]
    if ($c1 -ne $c2) { MR $ws 5 $c1 5 $c2 }
    if ($c1 -ne $c2) { MR $ws 4 $c1 4 $c2 }
    $ws.Cells.Item(5,$c1).Formula = $fml
    $ws.Cells.Item(5,$c1).Font.Bold = $true
    $ws.Cells.Item(5,$c1).Font.Size = 16
    $ws.Cells.Item(5,$c1).Interior.Color = $bg
    $ws.Cells.Item(5,$c1).Font.Color = $fg
    $ws.Cells.Item(5,$c1).HorizontalAlignment = $xlCenter
    $ws.Cells.Item(5,$c1).VerticalAlignment = $xlVCenter
    if ($lbl -eq "AVG % DONE") { $ws.Cells.Item(5,$c1).NumberFormat = "0.0" }
    # Label in row 4
    $ws.Cells.Item(4,$c1).Value2 = $lbl
    $ws.Cells.Item(4,$c1).Font.Bold = $true
    $ws.Cells.Item(4,$c1).Font.Size = 7
    $ws.Cells.Item(4,$c1).Interior.Color = $bg
    $ws.Cells.Item(4,$c1).Font.Color = $fg
    $ws.Cells.Item(4,$c1).HorizontalAlignment = $xlCenter
}

# Master overall bar — col I row 4-5
MR $ws 4 9 5 9
$ws.Cells.Item(4,9).Formula = "=IFERROR(REPT(CHAR(9608),ROUND(AVERAGEIF(B12:B200,""<>"""",F12:F200)/10,0))&REPT(CHAR(9617),10-ROUND(AVERAGEIF(B12:B200,""<>"""",F12:F200)/10,0)),REPT(CHAR(9617),10))"
$ws.Cells.Item(4,9).Font.Name = "Courier New"
$ws.Cells.Item(4,9).Font.Size = 14
$ws.Cells.Item(4,9).Interior.Color = $cDark
$ws.Cells.Item(4,9).Font.Color = $cBlue
$ws.Cells.Item(4,9).HorizontalAlignment = $xlCenter
$ws.Cells.Item(4,9).VerticalAlignment = $xlVCenter
$ws.Cells.Item(4,9).WrapText = $false

# ══════════════════════════════════════════════════════════
# ROW 6-7  — PRIORITY LEVEL INPUT SECTION
# ══════════════════════════════════════════════════════════
RH $ws 6 8
$ws.Rows.Item(6).Interior.Color = $cGray

RH $ws 7 24

MR $ws 7 1 7 2
W $ws 7 1 "  SET PRIORITY FOCUS:" $true $cAmberBG $cAmber $xlLeft $false 9

# Priority selector cell D7 — user picks focus level
W $ws 7 3 "P1 - Urgent" $true $cRedBG $cRed $xlCenter $false 9
$ws.Cells.Item(7,3).Name = "PriorityFocus"

MR $ws 7 4 7 5
W $ws 7 4 "  Role Focus:" $true $cBlueBG $cBlue $xlLeft $false 9
W $ws 7 6 "Deduction" $true $cBlueBG $cBlue $xlCenter $false 9
$ws.Cells.Item(7,6).Name = "RoleFocus"

MR $ws 7 7 7 9
W $ws 7 7 "  Use dropdowns in each row to assign Priority and Role. Colour coding applies automatically." $false $cGray $cMuted $xlLeft $false 8

# Dropdowns for the input cells
AddDV ($ws.Range("C7")) "P1 - Urgent,P2 - Important,P3 - Routine" "Priority Focus" "Set your main priority focus for today"
AddDV ($ws.Range("F7")) "Deduction,Reporting,Both" "Role Focus" "Set your main role focus for today"

# ══════════════════════════════════════════════════════════
# ROW 8  — THIN SEPARATOR
# ══════════════════════════════════════════════════════════
RH $ws 8 6
$ws.Rows.Item(8).Interior.Color = $cSlate

# ══════════════════════════════════════════════════════════
# ROW 9  — SECTION LABEL: DEDUCTION
# ══════════════════════════════════════════════════════════
RH $ws 9 20
MR $ws 9 1 9 6
W $ws 9 1 "  DEDUCTION ANALYST TASKS" $true $cDarkBlu $cWhite $xlLeft $false 9

# Deduction mini-bar right side
MR $ws 9 7 9 9
$ws.Cells.Item(9,7).Formula = "=IFERROR(""Ded: ""&ROUND(AVERAGEIF(C12:C200,""Deduction"",F12:F200),0)&""%  ""&REPT(CHAR(9608),ROUND(AVERAGEIF(C12:C200,""Deduction"",F12:F200)/10,0))&REPT(CHAR(9617),10-ROUND(AVERAGEIF(C12:C200,""Deduction"",F12:F200)/10,0)),""No Deduction tasks yet"")"
$ws.Cells.Item(9,7).Font.Name = "Courier New"
$ws.Cells.Item(9,7).Font.Size = 10
$ws.Cells.Item(9,7).Interior.Color = $cDarkBlu
$ws.Cells.Item(9,7).Font.Color = $cLBlue
$ws.Cells.Item(9,7).HorizontalAlignment = $xlRight
$ws.Cells.Item(9,7).VerticalAlignment = $xlVCenter

# ══════════════════════════════════════════════════════════
# ROW 10  — SECTION LABEL: REPORTING
# ══════════════════════════════════════════════════════════
RH $ws 10 20
MR $ws 10 1 10 6
W $ws 10 1 "  REPORTING ANALYST TASKS" $true $cDarkRed $cWhite $xlLeft $false 9

MR $ws 10 7 10 9
$ws.Cells.Item(10,7).Formula = "=IFERROR(""Rep: ""&ROUND(AVERAGEIF(C12:C200,""Reporting"",F12:F200),0)&""%  ""&REPT(CHAR(9608),ROUND(AVERAGEIF(C12:C200,""Reporting"",F12:F200)/10,0))&REPT(CHAR(9617),10-ROUND(AVERAGEIF(C12:C200,""Reporting"",F12:F200)/10,0)),""No Reporting tasks yet"")"
$ws.Cells.Item(10,7).Font.Name = "Courier New"
$ws.Cells.Item(10,7).Font.Size = 10
$ws.Cells.Item(10,7).Interior.Color = $cDarkRed
$ws.Cells.Item(10,7).Font.Color = RGB 255 200 200
$ws.Cells.Item(10,7).HorizontalAlignment = $xlRight
$ws.Cells.Item(10,7).VerticalAlignment = $xlVCenter

# ══════════════════════════════════════════════════════════
# ROW 11  — COLUMN HEADERS
# ══════════════════════════════════════════════════════════
RH $ws 11 22
$colHdrs = @("#","TASK / ACTIVITY DESCRIPTION","ROLE","PRIORITY LEVEL","DUE TIME","% DONE","PROGRESS BAR","STATUS","COMMENTS / REMARKS")
for ($c=1; $c -le 9; $c++){
    W $ws 11 $c $colHdrs[$c-1] $true $cDark $cWhite $xlCenter $false 9
    $ws.Cells.Item(11,$c).VerticalAlignment = $xlVCenter
    # bottom border
    $ws.Cells.Item(11,$c).Borders.Item(9).LineStyle = 1
    $ws.Cells.Item(11,$c).Borders.Item(9).Weight    = 3
}

# ══════════════════════════════════════════════════════════
# ROWS 12–61  — 50 EDITABLE TASK ROWS
# ══════════════════════════════════════════════════════════
$dataStart = 12
$dataEnd   = 61

for ($r = $dataStart; $r -le $dataEnd; $r++){
    RH $ws $r 22
    $num   = $r - $dataStart + 1
    $even  = ($num % 2 -eq 0)
    $rowBG = if ($even) { $cGray } else { $cWhite }

    # A: row number
    W $ws $r 1 $num $false $rowBG $cMuted $xlCenter $false 9

    # B: Task description (editable, white bg)
    $ws.Cells.Item($r,2).Interior.Color = $cWhite
    $ws.Cells.Item($r,2).HorizontalAlignment = $xlLeft
    $ws.Cells.Item($r,2).WrapText = $true
    $ws.Cells.Item($r,2).IndentLevel = 1

    # C: Role (dropdown)
    $ws.Cells.Item($r,3).Interior.Color = $cWhite
    $ws.Cells.Item($r,3).HorizontalAlignment = $xlCenter
    $ws.Cells.Item($r,3).Font.Bold = $true
    $ws.Cells.Item($r,3).Font.Size = 9

    # D: Priority (dropdown)
    $ws.Cells.Item($r,4).Interior.Color = $cWhite
    $ws.Cells.Item($r,4).HorizontalAlignment = $xlCenter
    $ws.Cells.Item($r,4).Font.Bold = $true
    $ws.Cells.Item($r,4).Font.Size = 9

    # E: Due time
    $ws.Cells.Item($r,5).Interior.Color = $cWhite
    $ws.Cells.Item($r,5).HorizontalAlignment = $xlCenter
    $ws.Cells.Item($r,5).Font.Size = 9
    $ws.Cells.Item($r,5).Font.Color = $cMuted

    # F: % Done (0-100) — leave blank by default so empty rows don't pollute averages
    $ws.Cells.Item($r,6).Interior.Color = $cWhite
    $ws.Cells.Item($r,6).HorizontalAlignment = $xlCenter
    $ws.Cells.Item($r,6).Font.Bold = $true
    $ws.Cells.Item($r,6).Font.Size = 10
    $ws.Cells.Item($r,6).NumberFormat = "0"

    # G: Progress bar formula — safe: blank row = empty, numeric coercion via VALUE()
    $fCell = "F$r"
    $bCell = "B$r"
    $ws.Cells.Item($r,7).Formula = "=IF($bCell="""","""",IFERROR(REPT(CHAR(9608),ROUND(VALUE($fCell)/10,0))&REPT(CHAR(9617),10-ROUND(VALUE($fCell)/10,0)),REPT(CHAR(9617),10)))"
    $ws.Cells.Item($r,7).Font.Name = "Courier New"
    $ws.Cells.Item($r,7).Font.Size = 12
    $ws.Cells.Item($r,7).Interior.Color = $cGray
    $ws.Cells.Item($r,7).HorizontalAlignment = $xlLeft
    $ws.Cells.Item($r,7).VerticalAlignment = $xlVCenter

    # H: Status (dropdown)
    $ws.Cells.Item($r,8).Interior.Color = $cWhite
    $ws.Cells.Item($r,8).HorizontalAlignment = $xlCenter
    $ws.Cells.Item($r,8).Font.Bold = $true
    $ws.Cells.Item($r,8).Font.Size = 9

    # I: Comments (free text, wrap)
    $ws.Cells.Item($r,9).Interior.Color = $cWhite
    $ws.Cells.Item($r,9).HorizontalAlignment = $xlLeft
    $ws.Cells.Item($r,9).WrapText = $true
    $ws.Cells.Item($r,9).Font.Size = 9
    $ws.Cells.Item($r,9).Font.Color = $cMuted
    $ws.Cells.Item($r,9).IndentLevel = 1

    # Light bottom border on each row
    for ($c=1;$c -le 9;$c++){
        $ws.Cells.Item($r,$c).Borders.Item(9).LineStyle = 1
        $ws.Cells.Item($r,$c).Borders.Item(9).Weight    = 1
        $ws.Cells.Item($r,$c).Borders.Item(9).Color     = RGB 229 231 235
    }
}

# ── Bulk dropdowns ───────────────────────────────────────────
AddDV ($ws.Range("C${dataStart}:C${dataEnd}")) "Deduction,Reporting,Both" "Role" "Deduction / Reporting / Both"
AddDV ($ws.Range("D${dataStart}:D${dataEnd}")) "P1 - Urgent,P2 - Important,P3 - Routine" "Priority" "P1=Urgent  P2=Important  P3=Routine"
AddDV ($ws.Range("H${dataStart}:H${dataEnd}")) "Pending,In Progress,Partial,Complete" "Status" "Pending / In Progress / Partial / Complete"

# % Done validation 0-100
$pctR = $ws.Range("F${dataStart}:F${dataEnd}")
$dv4  = $pctR.Validation; $dv4.Delete()
$dv4.Add(1,1,1,"0","100") | Out-Null
$dv4.ShowError = $true
$dv4.ErrorTitle = "Invalid"
$dv4.ErrorMessage = "Enter 0 to 100"

# ══════════════════════════════════════════════════════════
# CONDITIONAL FORMATTING
# ══════════════════════════════════════════════════════════

# Role column colours
$roleRng = $ws.Range("C${dataStart}:C${dataEnd}")
CF $roleRng "Deduction" $cBlueBG $cBlue
CF $roleRng "Reporting"  $cPurBG  $cPurple
CF $roleRng "Both"       $cGreenBG $cGreen

# Priority column colours
$priRng = $ws.Range("D${dataStart}:D${dataEnd}")
CF $priRng "P1 - Urgent"     $cRedBG   $cRed
CF $priRng "P2 - Important"  $cAmberBG $cAmber
CF $priRng "P3 - Routine"    $cGreenBG $cGreen

# Status column colours
$stRng = $ws.Range("H${dataStart}:H${dataEnd}")
CF $stRng "Complete"    $cGreenBG $cGreen
CF $stRng "Partial"     $cAmberBG $cAmber
CF $stRng "In Progress" $cBlueBG  $cBlue
CF $stRng "Pending"     $cRedBG   $cRed

# Progress bar font colour — expression CF relative to column F same row
# Use =$F12 style (row of dataStart, will apply relatively across range)
$barRng = $ws.Range("G${dataStart}:G${dataEnd}")
try {
    $b1 = $barRng.FormatConditions.Add(2,$null,"=F$dataStart=100")
    $b1.Font.Color = $cGreen; $b1.Interior.Color = RGB 230 244 234
} catch {}
try {
    $b2 = $barRng.FormatConditions.Add(2,$null,"=AND(F$dataStart>=50,F$dataStart<100)")
    $b2.Font.Color = $cAmber; $b2.Interior.Color = RGB 255 248 224
} catch {}
try {
    $b3 = $barRng.FormatConditions.Add(2,$null,"=F$dataStart<50")
    $b3.Font.Color = $cRed; $b3.Interior.Color = RGB 255 240 240
} catch {}

# % Done column — Excel data bar (blue gradient)
try {
    $cfdb = $ws.Range("F${dataStart}:F${dataEnd}").FormatConditions.AddDatabar()
    $cfdb.ShowValue = $true
    $cfdb.BarColor.Color = $cBlue
} catch {}

# ══════════════════════════════════════════════════════════
# ROW AFTER DATA — summary/total bar
# ══════════════════════════════════════════════════════════
$sumR = $dataEnd + 2
RH $ws $sumR 20
MR $ws $sumR 1 $sumR 9
$ws.Cells.Item($sumR,1).Formula = '=IFERROR("  TODAY''S OVERALL: "&ROUND(AVERAGE(F12:F200),1)&"% done  |  Complete: "&COUNTIF(H12:H200,"Complete")&"  Partial: "&COUNTIF(H12:H200,"Partial")&"  In Progress: "&COUNTIF(H12:H200,"In Progress")&"  Pending: "&COUNTIF(H12:H200,"Pending")," ")'
$ws.Cells.Item($sumR,1).Interior.Color = $cDark
$ws.Cells.Item($sumR,1).Font.Color = $cWhite
$ws.Cells.Item($sumR,1).Font.Bold = $true
$ws.Cells.Item($sumR,1).Font.Size = 9
$ws.Cells.Item($sumR,1).HorizontalAlignment = $xlLeft

# ══════════════════════════════════════════════════════════
# FREEZE PANES — freeze rows 1-11
# ══════════════════════════════════════════════════════════
$ws.Activate()
$ws.Cells.Item($dataStart,1).Select()
$excel.ActiveWindow.FreezePanes = $true

# AutoFilter on header row
$ws.Rows.Item(11).AutoFilter() | Out-Null

# ╔══════════════════════════════════════════════════════════╗
# ║  SHEET 2 — CHARTS                                        ║
# ╚══════════════════════════════════════════════════════════╝
$wsCh = $wb.Worksheets.Add([System.Reflection.Missing]::Value, $ws)
$wsCh.Name = "CHARTS"
$wsCh.Tab.Color = $cBlue
$wsCh.DisplayGridlines = $false

RH $wsCh 1 36
MR $wsCh 1 1 1 10
$wsCh.Cells.Item(1,1).Value2 = "CHARTS DASHBOARD  |  Live from DAILY TRACKER sheet"
$wsCh.Cells.Item(1,1).Font.Bold = $true
$wsCh.Cells.Item(1,1).Font.Size = 14
$wsCh.Cells.Item(1,1).Interior.Color = $cDark
$wsCh.Cells.Item(1,1).Font.Color = $cWhite
$wsCh.Cells.Item(1,1).IndentLevel = 1

RH $wsCh 2 18
MR $wsCh 2 1 2 10
$wsCh.Cells.Item(2,1).Value2 = "All charts update automatically. Go to DAILY TRACKER, fill in tasks, come back here to see charts."
$wsCh.Cells.Item(2,1).Font.Size = 9
$wsCh.Cells.Item(2,1).Interior.Color = $cBlueBG
$wsCh.Cells.Item(2,1).Font.Color = $cBlue
$wsCh.Cells.Item(2,1).IndentLevel = 1

# Helper data for charts (hidden cols M-N)
$wsCh.Columns.Item(13).ColumnWidth = 18
$wsCh.Columns.Item(14).ColumnWidth = 10

$chHelp = @(
    @(4, @("Status","Count"), @(
        @("Complete",    '=COUNTIF(''DAILY TRACKER''!H12:H200,"Complete")'),
        @("Partial",     '=COUNTIF(''DAILY TRACKER''!H12:H200,"Partial")'),
        @("In Progress", '=COUNTIF(''DAILY TRACKER''!H12:H200,"In Progress")'),
        @("Pending",     '=COUNTIF(''DAILY TRACKER''!H12:H200,"Pending")')
    )),
    @(10, @("Priority","Count"), @(
        @("P1 - Urgent",    '=COUNTIF(''DAILY TRACKER''!D12:D200,"P1 - Urgent")'),
        @("P2 - Important", '=COUNTIF(''DAILY TRACKER''!D12:D200,"P2 - Important")'),
        @("P3 - Routine",   '=COUNTIF(''DAILY TRACKER''!D12:D200,"P3 - Routine")')
    )),
    @(15, @("Role","Count"), @(
        @("Deduction", '=COUNTIF(''DAILY TRACKER''!C12:C200,"Deduction")'),
        @("Reporting",  '=COUNTIF(''DAILY TRACKER''!C12:C200,"Reporting")'),
        @("Both",       '=COUNTIF(''DAILY TRACKER''!C12:C200,"Both")')
    )),
    @(20, @("Priority","Avg % Done"), @(
        @("P1 - Urgent",    '=IFERROR(AVERAGEIF(''DAILY TRACKER''!D12:D200,"P1 - Urgent",''DAILY TRACKER''!F12:F200),0)'),
        @("P2 - Important", '=IFERROR(AVERAGEIF(''DAILY TRACKER''!D12:D200,"P2 - Important",''DAILY TRACKER''!F12:F200),0)'),
        @("P3 - Routine",   '=IFERROR(AVERAGEIF(''DAILY TRACKER''!D12:D200,"P3 - Routine",''DAILY TRACKER''!F12:F200),0)')
    ))
)

foreach ($blk in $chHelp) {
    $startR = $blk[0]; $hdrs = $blk[1]; $rows = $blk[2]
    $wsCh.Cells.Item($startR,13).Value2 = $hdrs[0]
    $wsCh.Cells.Item($startR,14).Value2 = $hdrs[1]
    $wsCh.Cells.Item($startR,13).Font.Bold = $true
    $wsCh.Cells.Item($startR,14).Font.Bold = $true
    for ($i=0;$i -lt $rows.Count;$i++){
        $wsCh.Cells.Item($startR+1+$i,13).Value2 = $rows[$i][0]
        $wsCh.Cells.Item($startR+1+$i,14).Formula = $rows[$i][1]
    }
}

# ── Add Charts ───────────────────────────────────────────────
# NOTE: keep helper cols visible while adding charts, hide after
$xlDoughnut        = -4120
$xlColumnClustered =    51
$xlBarClustered    =    57
$xlPie             =     5
$xlLine            =     4

function AddChart($sheet,$srcRange,$type,$title,$left,$top,$w,$h,$ptColors){
    try {
        $co = $sheet.Shapes.AddChart2(-1,$type,$left,$top,$w,$h)
        $ch = $co.Chart
        $ch.SetSourceData($srcRange)
        $ch.HasTitle = $true
        $ch.ChartTitle.Text = $title
        $ch.ChartTitle.Font.Size = 11
        $ch.ChartTitle.Font.Bold = $true
        $ch.ChartArea.Format.Fill.ForeColor.RGB = (RGB 255 255 255)
        $ch.PlotArea.Format.Fill.ForeColor.RGB  = (RGB 247 248 250)
        $ch.HasLegend = $true
        $ch.Legend.Font.Size = 9
        if ($ptColors){
            for ($i=0;$i -lt $ptColors.Count;$i++){
                try { $ch.SeriesCollection(1).Points($i+1).Format.Fill.ForeColor.RGB = $ptColors[$i] } catch {}
            }
        }
        return $ch
    } catch { return $null }
}

$L1=20;  $L2=390; $L3=760
$T1=60;  $T2=340; $T3=620
$CW=350; $CH=260

# Chart 1: Status Donut  (M4:N8 = header + 4 data rows)
AddChart $wsCh ($wsCh.Range("M4:N8"))   $xlDoughnut        "Task Status Breakdown"      $L1 $T1 $CW $CH @($cGreen,$cAmber,$cBlue,$cRed) | Out-Null
# Chart 2: Priority Column (M10:N13 = header + 3 data rows)
AddChart $wsCh ($wsCh.Range("M10:N13")) $xlColumnClustered "Tasks by Priority"          $L2 $T1 $CW $CH @($cRed,$cAmber,$cGreen) | Out-Null
# Chart 3: Role Pie (M15:N18 = header + 3 data rows)
AddChart $wsCh ($wsCh.Range("M15:N18")) $xlPie             "Tasks by Role"              $L3 $T1 $CW $CH @($cBlue,$cPurple,$cGreen) | Out-Null
# Chart 4: Avg % Done by Priority Bar (M20:N23 = header + 3 data rows)
AddChart $wsCh ($wsCh.Range("M20:N23")) $xlBarClustered    "Avg % Done by Priority"     $L1 $T2 $CW $CH @($cRed,$cAmber,$cGreen) | Out-Null
# Chart 5: Status count column (reuse status range)
AddChart $wsCh ($wsCh.Range("M4:N8"))   $xlColumnClustered "Status Count Overview"      $L2 $T2 $CW $CH @($cGreen,$cAmber,$cBlue,$cRed) | Out-Null

# Hide helper columns AFTER charts are linked
$wsCh.Columns.Item(13).Hidden = $true
$wsCh.Columns.Item(14).Hidden = $true

# Instruction box
$instrR = 33
MR $wsCh $instrR 1 $instrR 10
$wsCh.Cells.Item($instrR,1).Value2 = "HOW TO USE - 9 Steps"
$wsCh.Cells.Item($instrR,1).Font.Bold = $true
$wsCh.Cells.Item($instrR,1).Font.Size = 11
$wsCh.Cells.Item($instrR,1).Interior.Color = $cDark
$wsCh.Cells.Item($instrR,1).Font.Color = $cWhite
RH $wsCh $instrR 22

$steps = @(
    "1  Go to the DAILY TRACKER sheet (tab at bottom).",
    "2  In row 7, set your PRIORITY FOCUS (P1/P2/P3) and ROLE FOCUS (Deduction/Reporting) using the dropdowns.",
    "3  In Column B (rows 12 onward), type each task description.",
    "4  Column C - select Role: Deduction / Reporting / Both (dropdown).",
    "5  Column D - select Priority: P1-Urgent / P2-Important / P3-Routine (dropdown, colours auto-apply).",
    "6  Column E - type Due Time (e.g. 9:00 AM / EOD / 3:00 PM).",
    "7  Column F - type a number 0 to 100. The progress bar in Column G updates instantly.",
    "8  Column H - select Status: Pending / In Progress / Partial / Complete (colour-coded automatically).",
    "9  Column I - type any comments, notes or blockers for that task."
)
for ($i=0;$i -lt $steps.Count;$i++){
    $rr = $instrR+1+$i
    MR $wsCh $rr 1 $rr 10
    $bg2 = if ($i%2 -eq 0){ RGB 255 255 255 } else { RGB 247 248 250 }
    $wsCh.Cells.Item($rr,1).Value2 = "  " + $steps[$i]
    $wsCh.Cells.Item($rr,1).Interior.Color = $bg2
    $wsCh.Cells.Item($rr,1).Font.Size = 10
    $wsCh.Cells.Item($rr,1).Font.Color = $cDark
    RH $wsCh $rr 20
}

# ══════════════════════════════════════════════════════════
# Activate DAILY TRACKER and save
# ══════════════════════════════════════════════════════════
$ws.Activate()
$ws.Cells.Item($dataStart,2).Select()

$wb.SaveAs($OutputPath, 51)
$wb.Close($false)
$excel.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null

Write-Host "SUCCESS: $OutputPath"
