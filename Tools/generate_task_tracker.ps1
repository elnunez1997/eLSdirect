# Daily Task Tracker Generator - Deduction & Reporting Analyst

$OutputPath = "$PSScriptRoot\Daily_Task_Tracker.xlsx"

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$wb = $excel.Workbooks.Add()

function Set-Cell($ws,$row,$col,$value,$bold=$false,$bg=$null,$fg=$null,$hAlign=$null,$wrap=$false,$size=$null) {
    $cell = $ws.Cells.Item($row,$col)
    $cell.Value2 = $value
    if ($bold)   { $cell.Font.Bold = $true }
    if ($size)   { $cell.Font.Size = $size }
    if ($bg)     { $cell.Interior.Color = $bg }
    if ($fg)     { $cell.Font.Color = $fg }
    if ($wrap)   { $cell.WrapText = $true }
    if ($hAlign) { $cell.HorizontalAlignment = $hAlign }
}

function Hex2Long([string]$hex) {
    $r = [Convert]::ToInt32($hex.Substring(0,2),16)
    $g = [Convert]::ToInt32($hex.Substring(2,2),16)
    $b = [Convert]::ToInt32($hex.Substring(4,2),16)
    return $r + ($g * 256) + ($b * 65536)
}

$cDark    = Hex2Long "1F2328"
$cWhite   = Hex2Long "FFFFFF"
$cBlue    = Hex2Long "3B82D4"
$cPurple  = Hex2Long "7C5CD8"
$cRed     = Hex2Long "CF222E"
$cAmber   = Hex2Long "9A6700"
$cGreen   = Hex2Long "1A7F37"
$cGray    = Hex2Long "F7F8FA"
$cText    = Hex2Long "1F2328"
$cMuted   = Hex2Long "57606A"
$cRedBG   = Hex2Long "FFF0F0"
$cAmberBG = Hex2Long "FFF8E0"
$cGreenBG = Hex2Long "E6F4EA"
$cBlueBG  = Hex2Long "EEF2FF"
$cPurBG   = Hex2Long "F3E8FF"

$xlCenter = -4108
$xlLeft   = -4131

$dateStr = (Get-Date).ToString("dddd, MMMM dd yyyy")

# Task list: id, role, priority, description, due, pct, status, notes
$tasks = @(
  @(1,"Deduction","P1 - Urgent","Review and resolve escalated deductions high-value over 10K","EOD",75,"Partial","3 items pending backup"),
  @(2,"Deduction","P1 - Urgent","Process dispute filings for deductions due today","12:00 PM",100,"Complete","7 filed"),
  @(3,"Deduction","P1 - Urgent","ARCollect Activity Lookup - update Customer Name fields","10:00 AM",50,"In Progress","Batch in progress"),
  @(4,"Deduction","P1 - Urgent","Follow up on Invalid/Pending deductions - client response needed","EOD",25,"Pending","Awaiting client"),
  @(5,"Deduction","P2 - Important","Validate backup documents for open deductions","3:00 PM",50,"Partial",""),
  @(6,"Deduction","P2 - Important","Update ARCollect Activity Specific Fields - UNFI Natural batch","2:00 PM",75,"Partial","UNFI Natural batch"),
  @(7,"Deduction","P2 - Important","Weekly aging review - deductions over 60 days","5:00 PM",0,"Pending","Scheduled 4PM"),
  @(8,"Deduction","P2 - Important","Reconcile customer short-pays against promotions","EOD",25,"In Progress",""),
  @(9,"Deduction","P3 - Routine","File resolved deduction backup in shared drive","EOD",100,"Complete",""),
  @(10,"Deduction","P3 - Routine","Update deduction tracker spreadsheet","EOD",0,"Pending",""),
  @(11,"Reporting","P1 - Urgent","Submit Daily Aging Report to management","9:00 AM",100,"Complete","Sent 8:52 AM"),
  @(12,"Reporting","P1 - Urgent","Escalation report - deductions flagged by management","11:00 AM",100,"Complete",""),
  @(13,"Reporting","P1 - Urgent","Daily KPI dashboard update - resolution rate and amounts cleared","EOD",75,"Partial",""),
  @(14,"Reporting","P2 - Important","Weekly trend analysis - deduction patterns by client","5:00 PM",50,"In Progress",""),
  @(15,"Reporting","P2 - Important","SLA compliance tracking - identify at-risk items","4:00 PM",25,"Pending",""),
  @(16,"Reporting","P2 - Important","Root cause analysis report - top 5 deduction reasons","EOD",0,"Pending",""),
  @(17,"Reporting","P3 - Routine","Update shared reporting dashboard Power BI and Excel","EOD",50,"Partial",""),
  @(18,"Reporting","P3 - Routine","Archive last weeks reports in SharePoint","EOD",100,"Complete",""),
  @(19,"Reporting","P3 - Routine","Document process notes for new team members","EOD",0,"Pending","")
)

function Get-PriColors($pri) {
    if ($pri -like "P1*") { return @($cRed,$cRedBG) }
    if ($pri -like "P2*") { return @($cAmber,$cAmberBG) }
    return @($cGreen,$cGreenBG)
}
function Get-StColors($st) {
    switch ($st) {
        "Complete"    { return @($cGreen,$cGreenBG) }
        "Partial"     { return @($cAmber,$cAmberBG) }
        "Pending"     { return @($cRed,$cRedBG) }
        "In Progress" { return @($cBlue,$cBlueBG) }
        default       { return @($cText,$cGray) }
    }
}

function Build-Sheet($ws,$title,$titleBG,$titleFG,$accentBG,$filterRole,$filterPri,$showRole) {
    # Column widths
    $ws.Columns.Item(1).ColumnWidth = 4
    $ws.Columns.Item(2).ColumnWidth = 52
    if ($showRole) {
        $ws.Columns.Item(3).ColumnWidth = 14
        $ws.Columns.Item(4).ColumnWidth = 22
        $ws.Columns.Item(5).ColumnWidth = 12
        $ws.Columns.Item(6).ColumnWidth = 13
        $ws.Columns.Item(7).ColumnWidth = 18
        $ws.Columns.Item(8).ColumnWidth = 16
        $ws.Columns.Item(9).ColumnWidth = 28
        $lastCol = 9
        $colLetters = "A:I"
    } else {
        $ws.Columns.Item(3).ColumnWidth = 22
        $ws.Columns.Item(4).ColumnWidth = 12
        $ws.Columns.Item(5).ColumnWidth = 13
        $ws.Columns.Item(6).ColumnWidth = 18
        $ws.Columns.Item(7).ColumnWidth = 16
        $ws.Columns.Item(8).ColumnWidth = 28
        $lastCol = 8
        $colLetters = "A:H"
    }

    # Title
    $ws.Range("A1:$([char](64+$lastCol))1").Merge()
    Set-Cell $ws 1 1 $title $true $titleBG $titleFG $xlLeft $false 13
    $ws.Rows.Item(1).RowHeight = 32

    # Sub-title
    $ws.Range("A2:$([char](64+$lastCol))2").Merge()
    Set-Cell $ws 2 1 "Date: $dateStr" $false $accentBG $titleBG $xlLeft $false 10
    $ws.Rows.Item(2).RowHeight = 18

    # Headers row 4
    if ($showRole) {
        $hdrs = @("#","TASK DESCRIPTION","ROLE","PRIORITY","DUE","% DONE","PROGRESS BAR","STATUS","NOTES")
    } else {
        $hdrs = @("#","TASK DESCRIPTION","PRIORITY","DUE","% DONE","PROGRESS BAR","STATUS","NOTES")
    }
    for ($c=1;$c -le $hdrs.Count;$c++) {
        Set-Cell $ws 4 $c $hdrs[$c-1] $true $cDark $cWhite $xlCenter $false 9
    }
    $ws.Rows.Item(4).RowHeight = 20

    # Filter tasks
    $filtered = $tasks
    if ($filterRole) { $filtered = $filtered | Where-Object { $_[1] -eq $filterRole } }
    if ($filterPri)  { $filtered = $filtered | Where-Object { $_[2] -like "$filterPri*" } }

    $r = 5
    foreach ($t in $filtered) {
        $ws.Rows.Item($r).RowHeight = 22
        $rowBG = if ($r % 2 -eq 0) { $accentBG } else { $cWhite }
        $pc = Get-PriColors $t[2]
        $sc = Get-StColors  $t[6]
        $roleC = if ($t[1] -eq "Deduction") { $cBlue } else { $cPurple }
        $roleBG = if ($t[1] -eq "Deduction") { $cBlueBG } else { $cPurBG }

        Set-Cell $ws $r 1 ($r-4) $false $rowBG $cMuted $xlCenter
        Set-Cell $ws $r 2 $t[3] $false $rowBG $cText $xlLeft $true
        $ws.Cells.Item($r,2).IndentLevel = 1

        if ($showRole) {
            Set-Cell $ws $r 3 $t[1]       $true  $roleBG $roleC $xlCenter
            Set-Cell $ws $r 4 $t[2]       $true  $pc[1] $pc[0] $xlCenter
            Set-Cell $ws $r 5 $t[4]       $false $rowBG $cMuted $xlCenter
            Set-Cell $ws $r 6 ($t[5]/100) $true  $rowBG $sc[0] $xlCenter
            $ws.Cells.Item($r,6).NumberFormat = "0%"
            Set-Cell $ws $r 7 ($t[5]/100) $false $rowBG $cText $xlCenter
            $ws.Cells.Item($r,7).NumberFormat = "0%"
            Set-Cell $ws $r 8 $t[6]       $true  $sc[1] $sc[0] $xlCenter
            Set-Cell $ws $r 9 $t[7]       $false $rowBG $cMuted $xlLeft $true
            $barCol = 7
        } else {
            Set-Cell $ws $r 3 $t[2]       $true  $pc[1] $pc[0] $xlCenter
            Set-Cell $ws $r 4 $t[4]       $false $rowBG $cMuted $xlCenter
            Set-Cell $ws $r 5 ($t[5]/100) $true  $rowBG $sc[0] $xlCenter
            $ws.Cells.Item($r,5).NumberFormat = "0%"
            Set-Cell $ws $r 6 ($t[5]/100) $false $rowBG $cText $xlCenter
            $ws.Cells.Item($r,6).NumberFormat = "0%"
            Set-Cell $ws $r 7 $t[6]       $true  $sc[1] $sc[0] $xlCenter
            Set-Cell $ws $r 8 $t[7]       $false $rowBG $cMuted $xlLeft $true
            $barCol = 6
        }
        $r++
    }

    # Data bar CF on progress bar column
    $barColLetter = [char](64+$barCol)
    $barRange = $ws.Range("${barColLetter}5:${barColLetter}$($r-1)")
    $cf = $barRange.FormatConditions.AddDatabar()
    $cf.MinPoint.Modify(1, 0)
    $cf.MaxPoint.Modify(1, 1)
    $cf.BarColor.Color = $titleBG
    $cf.ShowValue = $true

    # AutoFilter
    $ws.Rows.Item(4).AutoFilter() | Out-Null

    return $r
}

# ============================================================
# SHEET 1 - ALL TASKS
# ============================================================
$ws1 = $wb.Worksheets.Item(1)
$ws1.Name = "ALL TASKS"
$ws1.Tab.Color = $cDark

$lastRow1 = Build-Sheet $ws1 "DAILY TASK TRACKER  -  Deduction Analyst + Reporting Analyst" $cDark $cWhite $cGray $null $null $true

# Summary counts
$cntC=0; $cntP=0; $cntPend=0; $cntW=0; $totalPct=0
foreach ($t in $tasks) {
    switch ($t[6]) { "Complete"{$cntC++} "Partial"{$cntP++} "Pending"{$cntPend++} "In Progress"{$cntW++} }
    $totalPct += $t[5]
}
$overallPct = [math]::Round($totalPct / ($tasks.Count * 100) * 100)

# Summary bar row 3
$ws1.Range("A3:I3").Merge()
$barText = "Overall Progress: $overallPct% complete  |  Completed: $cntC  Partial: $cntP  Pending: $cntPend  In Progress: $cntW  Total: $($tasks.Count)"
Set-Cell $ws1 3 1 $barText $true $cBlueBG $cBlue $xlLeft $false 10
$ws1.Rows.Item(3).RowHeight = 18

$ws1.Activate()
$ws1.Cells.Item(5,1).Select()
$excel.ActiveWindow.FreezePanes = $true

# ============================================================
# SHEET 2 - URGENT P1
# ============================================================
$ws2 = $wb.Worksheets.Add([System.Reflection.Missing]::Value, $ws1)
$ws2.Name = "URGENT P1"
$ws2.Tab.Color = $cRed
Build-Sheet $ws2 "URGENT - P1 TASKS  -  Act on These First!" $cRed $cWhite $cRedBG $null "P1" $true | Out-Null

# ============================================================
# SHEET 3 - DEDUCTION
# ============================================================
$ws3 = $wb.Worksheets.Add([System.Reflection.Missing]::Value, $ws2)
$ws3.Name = "DEDUCTION"
$ws3.Tab.Color = $cBlue
Build-Sheet $ws3 "DEDUCTION ANALYST - Daily Tasks" $cBlue $cWhite $cBlueBG "Deduction" $null $false | Out-Null

# ============================================================
# SHEET 4 - REPORTING
# ============================================================
$ws4 = $wb.Worksheets.Add([System.Reflection.Missing]::Value, $ws3)
$ws4.Name = "REPORTING"
$ws4.Tab.Color = $cPurple
Build-Sheet $ws4 "REPORTING ANALYST - Daily Tasks" $cPurple $cWhite $cPurBG "Reporting" $null $false | Out-Null

# ============================================================
# SHEET 5 - KPI DASHBOARD
# ============================================================
$ws5 = $wb.Worksheets.Add([System.Reflection.Missing]::Value, $ws4)
$ws5.Name = "KPI DASHBOARD"
$ws5.Tab.Color = $cGreen

$ws5.Columns.Item(1).ColumnWidth = 32
$ws5.Columns.Item(2).ColumnWidth = 16
$ws5.Columns.Item(3).ColumnWidth = 16
$ws5.Columns.Item(4).ColumnWidth = 14
$ws5.Columns.Item(5).ColumnWidth = 16

$ws5.Range("A1:E1").Merge()
Set-Cell $ws5 1 1 "KPI DASHBOARD  -  Daily Performance Metrics" $true $cDark $cWhite $xlLeft $false 13
$ws5.Rows.Item(1).RowHeight = 32

$ws5.Range("A2:E2").Merge()
Set-Cell $ws5 2 1 "Date: $dateStr  |  Combined: Deduction + Reporting" $false $cGreenBG $cGreen $xlLeft $false 10
$ws5.Rows.Item(2).RowHeight = 18

# KPI Headers
$kpiHdrs = @("METRIC","ACTUAL","TARGET","% ACHIEVED","STATUS")
for ($c=1;$c -le 5;$c++) { Set-Cell $ws5 4 $c $kpiHdrs[$c-1] $true $cDark $cWhite $xlCenter $false 9 }
$ws5.Rows.Item(4).RowHeight = 20

$kpis = @(
    @("Deductions Resolved (count)","18","25","72%","On Track"),
    @("Amount Cleared","USD 142,500","USD 200,000","71%","On Track"),
    @("Reports Submitted","3","4","75%","On Track"),
    @("Aging Queue Reviewed","85%","100%","85%","On Track"),
    @("Disputes Filed","7","7","100%","Complete"),
    @("Overall Task Completion","$overallPct%","100%","$overallPct%","In Progress")
)

$rk = 5
foreach ($k in $kpis) {
    $ws5.Rows.Item($rk).RowHeight = 22
    $rowBGk = if ($rk % 2 -eq 0) { $cGray } else { $cWhite }
    $sc = Get-StColors $k[4]
    Set-Cell $ws5 $rk 1 $k[0] $false $rowBGk $cText $xlLeft
    Set-Cell $ws5 $rk 2 $k[1] $true  $rowBGk $sc[0] $xlCenter
    Set-Cell $ws5 $rk 3 $k[2] $false $rowBGk $cMuted $xlCenter
    Set-Cell $ws5 $rk 4 $k[3] $true  $rowBGk $sc[0] $xlCenter
    Set-Cell $ws5 $rk 5 $k[4] $true  $sc[1] $sc[0] $xlCenter
    $rk++
}

# Priority Matrix section
$ws5.Range("A$($rk+1):E$($rk+1)").Merge()
Set-Cell $ws5 ($rk+1) 1 "EISENHOWER PRIORITY MATRIX" $true $cDark $cWhite $xlCenter $false 10
$ws5.Rows.Item($rk+1).RowHeight = 24

$m = $rk+2
$ws5.Range("A${m}:B${m}").Merge()
Set-Cell $ws5 $m 1 "URGENT" $true $cRed $cWhite $xlCenter $false 10
$ws5.Range("C${m}:E${m}").Merge()
Set-Cell $ws5 $m 3 "NOT URGENT" $true $cAmber $cWhite $xlCenter $false 10
$ws5.Rows.Item($m).RowHeight = 20

$m2 = $m+1
$ws5.Range("A${m2}:B$($m2+3)").Merge()
Set-Cell $ws5 $m2 1 "P1 - DO FIRST: Deduction disputes due today / Escalated deductions / Missing backup / Daily aging report / Critical calls" $true $cRedBG $cRed $xlLeft $true 9
$ws5.Range("C${m2}:E$($m2+3)").Merge()
Set-Cell $ws5 $m2 3 "P2 - SCHEDULE: Weekly trend analysis / Root cause review / SLA tracking / Dashboard updates / Process improvement" $true $cAmberBG $cAmber $xlLeft $true 9
$ws5.Rows.Item($m2).RowHeight = 70

$m3 = $m2+4
$ws5.Range("A${m3}:B$($m3+3)").Merge()
Set-Cell $ws5 $m3 1 "P2B - DELEGATE: Routine email follow-ups / Document filing / Standard status updates / Data entry low-value" $false $cBlueBG $cBlue $xlLeft $true 9
$ws5.Range("C${m3}:E$($m3+3)").Merge()
Set-Cell $ws5 $m3 3 "P3 - ROUTINE: EOD reconciliation / File organization / Notes and documentation / Learning and training" $false $cGreenBG $cGreen $xlLeft $true 9
$ws5.Rows.Item($m3).RowHeight = 60

# Activate first sheet and save
$ws1.Activate()
$wb.SaveAs($OutputPath, 51)
$wb.Close($false)
$excel.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null

Write-Host "SUCCESS: File saved -> $OutputPath"
