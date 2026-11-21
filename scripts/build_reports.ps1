Add-Type -AssemblyName System.Drawing
$ErrorActionPreference = 'Stop'

$databasePath = '.\gp.accdb'
$logPath      = '.\build_reports.log'

Set-Content -Path $logPath -Value @('START ' + (Get-Date).ToString('s'))
function Log([string]$m) { Add-Content -Path $logPath -Value ((Get-Date).ToString('s') + ' ' + $m) }

# ── Access constants ──
$acDetail        = 0
$acPageHeader    = 3
$acPageFooter    = 4
$acReport        = 3
$acLabel         = 100
$acTextBox       = 109
$acLine          = 103
$dbFailOnError   = 128

function Get-OleColor([int]$r,[int]$g,[int]$b) {
    return [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::FromArgb($r,$g,$b))
}

# ── Colour palettes (varied per report) ──
$palettes = @(
    @{ Head = Get-OleColor 33 43 68;   HeadBg = Get-OleColor 220 230 241; Row = Get-OleColor 245 248 252; White = Get-OleColor 255 255 255; Text = Get-OleColor 40 40 40;  Accent = Get-OleColor 52 119 180 }
    @{ Head = Get-OleColor 60 20 80;   HeadBg = Get-OleColor 235 225 245; Row = Get-OleColor 250 245 255; White = Get-OleColor 255 255 255; Text = Get-OleColor 50 40 60;  Accent = Get-OleColor 128 60 160 }
    @{ Head = Get-OleColor 0 80 60;    HeadBg = Get-OleColor 215 240 230; Row = Get-OleColor 240 250 245; White = Get-OleColor 255 255 255; Text = Get-OleColor 30 50 40;  Accent = Get-OleColor 0 140 100  }
    @{ Head = Get-OleColor 120 50 10;  HeadBg = Get-OleColor 255 235 215; Row = Get-OleColor 255 248 240; White = Get-OleColor 255 255 255; Text = Get-OleColor 60 40 20;  Accent = Get-OleColor 200 100 30 }
    @{ Head = Get-OleColor 20 60 100;  HeadBg = Get-OleColor 210 228 245; Row = Get-OleColor 235 245 255; White = Get-OleColor 255 255 255; Text = Get-OleColor 30 40 60;  Accent = Get-OleColor 40 120 200 }
)

function Remove-ReportIfExists($access, [string]$name) {
    try { $access.DoCmd.DeleteObject($acReport, $name) } catch { }
}

function Add-OrReplaceQuery($db, [string]$qn, [string]$sql) {
    try { $db.QueryDefs.Delete($qn) } catch { }
    [void]$db.CreateQueryDef($qn, $sql)
    Log ('QUERY READY ' + $qn)
}

function Add-Lbl($access, [string]$rn, [int]$section, [string]$cap, [int]$l, [int]$t, [int]$w, [int]$h, [string]$font, [int]$fs, [int]$fw, $fc, $bc, [int]$align=1) {
    $c = $access.CreateReportControl($rn, $acLabel, $section, '', '', $l, $t, $w, $h)
    $c.Caption   = $cap
    $c.FontName  = $font
    $c.FontSize  = $fs
    $c.FontWeight = $fw
    $c.ForeColor = $fc
    $c.TextAlign = $align
    if ($null -ne $bc) { $c.BackStyle = 1; $c.BackColor = $bc } else { $c.BackStyle = 0 }
    $c.BorderStyle = 0
    return $c
}

function Add-Txt($access, [string]$rn, [int]$section, [string]$src, [int]$l, [int]$t, [int]$w, [int]$h, [string]$font, [int]$fs, [int]$fw, $fc, [int]$align=1, [string]$fmt='') {
    $c = $access.CreateReportControl($rn, $acTextBox, $section, '', '', $l, $t, $w, $h)
    $c.ControlSource = $src
    $c.FontName  = $font
    $c.FontSize  = $fs
    $c.FontWeight = $fw
    $c.ForeColor = $fc
    $c.TextAlign = $align
    $c.BackStyle = 0
    $c.BorderStyle = 0
    if ($fmt -ne '') { $c.Format = $fmt }
    return $c
}

function Add-HLine($access, [string]$rn, [int]$section, [int]$l, [int]$t, [int]$w, $color) {
    $c = $access.CreateReportControl($rn, $acLine, $section, '', '', $l, $t, $w, 0)
    $c.BorderColor = $color
    $c.BorderWidth = 1
    return $c
}

# ══════════════════════════════════════════════════════════
# OPEN ACCESS
# ══════════════════════════════════════════════════════════
Log 'OPENING ACCESS'
$access = New-Object -ComObject Access.Application
$access.Visible = $false
$access.OpenCurrentDatabase($databasePath, $true)  # Exclusive access required for saving reports
$db = $access.CurrentDb()
Log 'DATABASE OPEN'

# Delete any existing reports (including test leftovers)
$reportNames = @(
    'Rpt Vaccination by Month', 'Rpt Citizen Gender Distribution',
    'Rpt Vaccine Brand Popularity', 'Rpt Appointment Status Summary',
    'Rpt Side Effects by Severity', 'Rpt Center Workload by Region',
    'Rpt Citizen Age Group Distribution', 'Rpt Batch Stock and Expiry',
    'Rpt Symptom Frequency', 'Rpt Monthly Appointment Trend',
    'Test Report 1', 'Test Report 2', 'Test Report Alpha'
)
foreach ($n in $reportNames) {
    try { $access.DoCmd.DeleteObject($acReport, $n); Log "DELETED $n" } catch {}
}
# Also try to delete Chinese auto-named reports
for ($i = 1; $i -le 20; $i++) {
    try { $access.DoCmd.DeleteObject($acReport, "`u{5831}`u{8868}$i") } catch {}
}

# ══════════════════════════════════════════════════════════
# ENSURE REPORT QUERIES
# ══════════════════════════════════════════════════════════

Add-OrReplaceQuery $db 'qry_Rpt_VaxByMonth' @'
SELECT Format([VaccinationDate],"yyyy-mm") AS VaxMonth, Count(*) AS VaxCount
FROM VaccinationRecords
GROUP BY Format([VaccinationDate],"yyyy-mm")
ORDER BY Format([VaccinationDate],"yyyy-mm");
'@

Add-OrReplaceQuery $db 'qry_Rpt_GenderDist' @'
SELECT Gender, Count(*) AS CitizenCount
FROM Citizens
GROUP BY Gender
ORDER BY Gender;
'@

Add-OrReplaceQuery $db 'qry_Rpt_VaxBrandPop' @'
SELECT V.VaccineName, V.Manufacturer, Count(R.RecordID) AS TimesAdministered
FROM Vaccines AS V
LEFT JOIN VaccinationRecords AS R ON V.VaccineID = R.VaccineID
GROUP BY V.VaccineName, V.Manufacturer
ORDER BY Count(R.RecordID) DESC;
'@

Add-OrReplaceQuery $db 'qry_Rpt_ApptStatus' @'
SELECT AppointmentStatus, Count(*) AS StatusCount
FROM Appointments
GROUP BY AppointmentStatus
ORDER BY Count(*) DESC;
'@

Add-OrReplaceQuery $db 'qry_Rpt_SideEffectSev' @'
SELECT SeverityLevel, SymptomType, Count(*) AS ReportCount
FROM SideEffects
GROUP BY SeverityLevel, SymptomType
ORDER BY SeverityLevel, Count(*) DESC;
'@

Add-OrReplaceQuery $db 'qry_Rpt_CenterRegion' @'
SELECT VC.Region, VC.CenterName, Count(A.AppointmentID) AS TotalAppointments
FROM VaccinationCenters AS VC
LEFT JOIN Appointments AS A ON VC.CenterID = A.CenterID
GROUP BY VC.Region, VC.CenterName
ORDER BY VC.Region, Count(A.AppointmentID) DESC;
'@

Add-OrReplaceQuery $db 'qry_Rpt_AgeGroup' @'
SELECT
  IIf(DateDiff("yyyy",[DOB],Date()) - IIf(Format([DOB],"mmdd")>Format(Date(),"mmdd"),1,0) <= 11,"0-11",
  IIf(DateDiff("yyyy",[DOB],Date()) - IIf(Format([DOB],"mmdd")>Format(Date(),"mmdd"),1,0) <= 17,"12-17",
  IIf(DateDiff("yyyy",[DOB],Date()) - IIf(Format([DOB],"mmdd")>Format(Date(),"mmdd"),1,0) <= 29,"18-29",
  IIf(DateDiff("yyyy",[DOB],Date()) - IIf(Format([DOB],"mmdd")>Format(Date(),"mmdd"),1,0) <= 44,"30-44",
  IIf(DateDiff("yyyy",[DOB],Date()) - IIf(Format([DOB],"mmdd")>Format(Date(),"mmdd"),1,0) <= 59,"45-59","60+"))))) AS AgeGroup,
  Count(*) AS CitizenCount
FROM Citizens
GROUP BY
  IIf(DateDiff("yyyy",[DOB],Date()) - IIf(Format([DOB],"mmdd")>Format(Date(),"mmdd"),1,0) <= 11,"0-11",
  IIf(DateDiff("yyyy",[DOB],Date()) - IIf(Format([DOB],"mmdd")>Format(Date(),"mmdd"),1,0) <= 17,"12-17",
  IIf(DateDiff("yyyy",[DOB],Date()) - IIf(Format([DOB],"mmdd")>Format(Date(),"mmdd"),1,0) <= 29,"18-29",
  IIf(DateDiff("yyyy",[DOB],Date()) - IIf(Format([DOB],"mmdd")>Format(Date(),"mmdd"),1,0) <= 44,"30-44",
  IIf(DateDiff("yyyy",[DOB],Date()) - IIf(Format([DOB],"mmdd")>Format(Date(),"mmdd"),1,0) <= 59,"45-59","60+")))))
ORDER BY 1;
'@

Add-OrReplaceQuery $db 'qry_Rpt_BatchStock' @'
SELECT B.BatchNumber, V.VaccineName, V.Manufacturer, B.CurrentStock,
       B.ExpiryDate, B.Supplier,
       IIf([ExpiryDate]<Date(),"EXPIRED",IIf([ExpiryDate]<DateAdd("d",30,Date()),"EXPIRING SOON","OK")) AS StockAlert
FROM VaccineBatches AS B
INNER JOIN Vaccines AS V ON B.VaccineID = V.VaccineID
ORDER BY B.ExpiryDate;
'@

Add-OrReplaceQuery $db 'qry_Rpt_SymptomFreq' @'
SELECT SymptomType, Count(*) AS Occurrences,
       Round(Count(*)*100.0/(SELECT Count(*) FROM SideEffects),1) AS PctOfTotal
FROM SideEffects
GROUP BY SymptomType
ORDER BY Count(*) DESC;
'@

Add-OrReplaceQuery $db 'qry_Rpt_MonthlyApptTrend' @'
SELECT Format([AppointmentDate],"yyyy-mm") AS ApptMonth,
       AppointmentStatus,
       Count(*) AS ApptCount
FROM Appointments
GROUP BY Format([AppointmentDate],"yyyy-mm"), AppointmentStatus
ORDER BY Format([AppointmentDate],"yyyy-mm"), AppointmentStatus;
'@

Log 'ALL REPORT QUERIES READY'

# ══════════════════════════════════════════════════════════
# REPORT BUILDER — tabular (title + columns in PageHeader, data in Detail, page info in PageFooter)
# ══════════════════════════════════════════════════════════
function Build-TabularReport(
    [string]$reportName,
    [string]$queryName,
    [string]$title,
    [string]$subtitle,
    [hashtable]$pal,
    [array]$columns,
    [int]$reportWidth = 14400
) {
    Remove-ReportIfExists $access $reportName

    $rpt = $access.CreateReport()
    $rptName = $rpt.Name
    $rpt.RecordSource = $queryName
    $rpt.Caption = $reportName
    $rpt.Width = $reportWidth

    # Page Header — title + subtitle + column headers
    $rpt.Section($acPageHeader).Height = 1300
    $rpt.Section($acPageHeader).BackColor = $pal.HeadBg

    Add-Lbl $access $rptName $acPageHeader $title 200 80 ($reportWidth - 400) 400 'Georgia' 16 700 $pal.Head $null 1 | Out-Null
    Add-Lbl $access $rptName $acPageHeader $subtitle 200 500 ($reportWidth - 400) 280 'Calibri' 9 400 $pal.Text $null 1 | Out-Null

    $left = 200
    foreach ($col in $columns) {
        Add-Lbl $access $rptName $acPageHeader $col.Header $left 920 $col.Width 280 'Calibri' 9 700 $pal.Head $null $col.Align | Out-Null
        $left += $col.Width
    }
    Add-HLine $access $rptName $acPageHeader 100 1260 ($reportWidth - 200) $pal.Accent | Out-Null

    # Detail
    $rpt.Section($acDetail).Height = 340
    $rpt.Section($acDetail).BackColor = $pal.White

    $left = 200
    foreach ($col in $columns) {
        $fmt = ''
        if ($col.ContainsKey('Fmt')) { $fmt = $col.Fmt }
        Add-Txt $access $rptName $acDetail $col.Source $left 40 $col.Width 260 'Calibri' 9 400 $pal.Text $col.Align $fmt | Out-Null
        $left += $col.Width
    }

    # Page Footer
    $rpt.Section($acPageFooter).Height = 400
    Add-Lbl $access $rptName $acPageFooter 'Vaccine Registration System' 200 60 4000 260 'Calibri' 8 400 $pal.Text $null 1 | Out-Null
    Add-Txt $access $rptName $acPageFooter "='Page ' & [Page] & ' of ' & [Pages]" ($reportWidth - 2500) 60 2200 260 'Calibri' 8 400 $pal.Text 3 | Out-Null

    $access.RunCommand(20)  # acCmdSave — saves the active design object
    $access.DoCmd.Close($acReport, $rptName, 2)  # 2 = acSaveNo (already saved)
    if ($rptName -ne $reportName) {
        try { $access.DoCmd.Rename($reportName, $acReport, $rptName) } catch { Log ('RENAME SKIP ' + $reportName) }
    }
    Log ('REPORT READY ' + $reportName)
}

# ══════════════════════════════════════════════════════════
# REPORT BUILDER — summary cards in PageHeader + data columns
# ══════════════════════════════════════════════════════════
function Build-SummaryReport(
    [string]$reportName,
    [string]$queryName,
    [string]$title,
    [string]$subtitle,
    [hashtable]$pal,
    [array]$metrics,
    [array]$columns,
    [int]$reportWidth = 14400
) {
    Remove-ReportIfExists $access $reportName

    $rpt = $access.CreateReport()
    $rptName = $rpt.Name
    $rpt.RecordSource = $queryName
    $rpt.Caption = $reportName
    $rpt.Width = $reportWidth

    # Page Header — title + metric cards + column headers
    $rpt.Section($acPageHeader).Height = 2200
    $rpt.Section($acPageHeader).BackColor = $pal.HeadBg

    Add-Lbl $access $rptName $acPageHeader $title 200 80 ($reportWidth - 400) 380 'Georgia' 16 700 $pal.Head $null 1 | Out-Null
    Add-Lbl $access $rptName $acPageHeader $subtitle 200 480 ($reportWidth - 400) 240 'Calibri' 9 400 $pal.Text $null 1 | Out-Null

    # Metric cards
    $cardWidth = [math]::Floor(($reportWidth - 600) / [math]::Max($metrics.Count, 1)) - 200
    $cardLeft  = 300
    foreach ($m in $metrics) {
        # Card background
        Add-Lbl $access $rptName $acPageHeader '' $cardLeft 800 $cardWidth 600 'Calibri' 8 400 $pal.Text $pal.White 2 | Out-Null
        # Label
        Add-Lbl $access $rptName $acPageHeader $m.Label ($cardLeft + 80) 820 ($cardWidth - 160) 220 'Calibri' 8 600 $pal.Text $null 2 | Out-Null
        # Value
        Add-Txt $access $rptName $acPageHeader $m.Expr ($cardLeft + 80) 1060 ($cardWidth - 160) 300 'Georgia' 14 700 $pal.Head 2 | Out-Null
        $cardLeft += ($cardWidth + 200)
    }

    # Column headers
    $left = 200
    foreach ($col in $columns) {
        Add-Lbl $access $rptName $acPageHeader $col.Header $left 1600 $col.Width 280 'Calibri' 9 700 $pal.Head $null $col.Align | Out-Null
        $left += $col.Width
    }
    Add-HLine $access $rptName $acPageHeader 100 1920 ($reportWidth - 200) $pal.Accent | Out-Null

    # Detail
    $rpt.Section($acDetail).Height = 320
    $rpt.Section($acDetail).BackColor = $pal.White

    $left = 200
    foreach ($col in $columns) {
        $fmt = ''
        if ($col.ContainsKey('Fmt')) { $fmt = $col.Fmt }
        Add-Txt $access $rptName $acDetail $col.Source $left 30 $col.Width 260 'Calibri' 9 400 $pal.Text $col.Align $fmt | Out-Null
        $left += $col.Width
    }

    # Page Footer
    $rpt.Section($acPageFooter).Height = 400
    Add-Lbl $access $rptName $acPageFooter 'Vaccine Registration System' 200 60 4000 260 'Calibri' 8 400 $pal.Text $null 1 | Out-Null
    Add-Txt $access $rptName $acPageFooter "='Page ' & [Page] & ' of ' & [Pages]" ($reportWidth - 2500) 60 2200 260 'Calibri' 8 400 $pal.Text 3 | Out-Null

    $access.RunCommand(20)  # acCmdSave — saves the active design object
    $access.DoCmd.Close($acReport, $rptName, 2)  # 2 = acSaveNo (already saved)
    if ($rptName -ne $reportName) {
        try { $access.DoCmd.Rename($reportName, $acReport, $rptName) } catch { Log ('RENAME SKIP ' + $reportName) }
    }
    Log ('REPORT READY ' + $reportName)
}

# ══════════════════════════════════════════════════════════
# REPORT BUILDER — wide detail with alternating colour hint (no group, clean layout)
# ══════════════════════════════════════════════════════════
function Build-WideReport(
    [string]$reportName,
    [string]$queryName,
    [string]$title,
    [string]$subtitle,
    [hashtable]$pal,
    [array]$columns,
    [int]$reportWidth = 16000
) {
    Remove-ReportIfExists $access $reportName

    $rpt = $access.CreateReport()
    $rptName = $rpt.Name
    $rpt.RecordSource = $queryName
    $rpt.Caption = $reportName
    $rpt.Width = $reportWidth

    # Page Header
    $rpt.Section($acPageHeader).Height = 1400
    $rpt.Section($acPageHeader).BackColor = $pal.Row

    Add-Lbl $access $rptName $acPageHeader $title 200 60 ($reportWidth - 400) 380 'Georgia' 14 700 $pal.Head $null 1 | Out-Null
    Add-Lbl $access $rptName $acPageHeader $subtitle 200 450 ($reportWidth - 400) 240 'Calibri' 9 400 $pal.Text $null 1 | Out-Null
    Add-HLine $access $rptName $acPageHeader 100 780 ($reportWidth - 200) $pal.Accent | Out-Null

    # Column headers — bold, slightly larger
    $left = 200
    foreach ($col in $columns) {
        Add-Lbl $access $rptName $acPageHeader $col.Header $left 880 $col.Width 340 'Calibri' 10 700 $pal.Head $null $col.Align | Out-Null
        $left += $col.Width
    }
    Add-HLine $access $rptName $acPageHeader 100 1340 ($reportWidth - 200) $pal.Accent | Out-Null

    # Detail — slightly taller rows
    $rpt.Section($acDetail).Height = 380
    $rpt.Section($acDetail).BackColor = $pal.White

    $left = 200
    foreach ($col in $columns) {
        $fmt = ''
        if ($col.ContainsKey('Fmt')) { $fmt = $col.Fmt }
        Add-Txt $access $rptName $acDetail $col.Source $left 60 $col.Width 260 'Calibri' 9 400 $pal.Text $col.Align $fmt | Out-Null
        $left += $col.Width
    }
    Add-HLine $access $rptName $acDetail 200 360 ($reportWidth - 400) $pal.Row | Out-Null

    # Page Footer
    $rpt.Section($acPageFooter).Height = 460
    Add-HLine $access $rptName $acPageFooter 100 10 ($reportWidth - 200) $pal.Accent | Out-Null
    Add-Lbl $access $rptName $acPageFooter 'Vaccine Registration System' 200 60 4500 280 'Calibri' 8 400 $pal.Text $null 1 | Out-Null
    Add-Txt $access $rptName $acPageFooter "='Page ' & [Page] & ' of ' & [Pages]" ($reportWidth - 2800) 60 2500 280 'Calibri' 8 400 $pal.Text 3 | Out-Null
    Add-Txt $access $rptName $acPageFooter "='Printed: ' & Format(Now(),'dd-mmm-yyyy hh:nn')" 5500 60 4000 280 'Calibri' 8 400 $pal.Text 2 | Out-Null

    $access.RunCommand(20)  # acCmdSave — saves the active design object
    $access.DoCmd.Close($acReport, $rptName, 2)  # 2 = acSaveNo (already saved)
    if ($rptName -ne $reportName) {
        try { $access.DoCmd.Rename($reportName, $acReport, $rptName) } catch { Log ('RENAME SKIP ' + $reportName) }
    }
    Log ('REPORT READY ' + $reportName)
}

# ══════════════════════════════════════════════════════════
# BUILD ALL 10 REPORTS
# ══════════════════════════════════════════════════════════

# 1 — Monthly Vaccination Trend (tabular, compact)
Build-TabularReport `
    -reportName 'Rpt Vaccination by Month' `
    -queryName 'qry_Rpt_VaxByMonth' `
    -title 'Monthly Vaccination Trend' `
    -subtitle 'Number of vaccination records per calendar month.' `
    -pal $palettes[0] `
    -columns @(
        @{ Header='Month';  Source='VaxMonth';  Width=5000; Align=1 },
        @{ Header='Count';  Source='VaxCount';  Width=3000; Align=3 }
    ) `
    -reportWidth 10000

# 2 — Citizen Gender Distribution (summary cards)
Build-SummaryReport `
    -reportName 'Rpt Citizen Gender Distribution' `
    -queryName 'qry_Rpt_GenderDist' `
    -title 'Citizen Gender Distribution' `
    -subtitle 'Breakdown of registered citizens by gender.' `
    -pal $palettes[1] `
    -metrics @(
        @{ Label='Total Citizens'; Expr="=DCount('*','Citizens')" },
        @{ Label='Male';           Expr="=DCount('*','Citizens','[Gender]=""Male""')" },
        @{ Label='Female';         Expr="=DCount('*','Citizens','[Gender]=""Female""')" }
    ) `
    -columns @(
        @{ Header='Gender'; Source='Gender'; Width=5000; Align=1 },
        @{ Header='Count';  Source='CitizenCount'; Width=3000; Align=3 }
    ) `
    -reportWidth 10000

# 3 — Vaccine Brand Popularity (tabular, 3-column)
Build-TabularReport `
    -reportName 'Rpt Vaccine Brand Popularity' `
    -queryName 'qry_Rpt_VaxBrandPop' `
    -title 'Vaccine Brand Popularity' `
    -subtitle 'Number of times each vaccine has been administered.' `
    -pal $palettes[2] `
    -columns @(
        @{ Header='Vaccine Name';    Source='VaccineName';       Width=4500; Align=1 },
        @{ Header='Manufacturer';    Source='Manufacturer';      Width=3500; Align=1 },
        @{ Header='Times Given';     Source='TimesAdministered'; Width=2500; Align=3 }
    ) `
    -reportWidth 12500

# 4 — Appointment Status Summary (summary cards)
Build-SummaryReport `
    -reportName 'Rpt Appointment Status Summary' `
    -queryName 'qry_Rpt_ApptStatus' `
    -title 'Appointment Status Summary' `
    -subtitle 'Distribution of appointment statuses across all bookings.' `
    -pal $palettes[3] `
    -metrics @(
        @{ Label='Total Appointments'; Expr="=DCount('*','Appointments')" },
        @{ Label='Completed';          Expr="=DCount('*','Appointments','[AppointmentStatus]=""Completed""')" },
        @{ Label='Cancelled';          Expr="=DCount('*','Appointments','[AppointmentStatus]=""Cancelled""')" }
    ) `
    -columns @(
        @{ Header='Status'; Source='AppointmentStatus'; Width=5000; Align=1 },
        @{ Header='Count';  Source='StatusCount';       Width=3000; Align=3 }
    ) `
    -reportWidth 10000

# 5 — Side Effects by Severity (tabular)
Build-TabularReport `
    -reportName 'Rpt Side Effects by Severity' `
    -queryName 'qry_Rpt_SideEffectSev' `
    -title 'Side Effects by Severity Level' `
    -subtitle 'Reported side effects grouped by severity and symptom type.' `
    -pal $palettes[4] `
    -columns @(
        @{ Header='Severity';     Source='SeverityLevel'; Width=3000; Align=1 },
        @{ Header='Symptom';      Source='SymptomType';   Width=5000; Align=1 },
        @{ Header='Report Count'; Source='ReportCount';   Width=2500; Align=3 }
    ) `
    -reportWidth 12500

# 6 — Center Workload by Region (wide layout with row separator lines)
Build-WideReport `
    -reportName 'Rpt Center Workload by Region' `
    -queryName 'qry_Rpt_CenterRegion' `
    -title 'Center Workload by Region' `
    -subtitle 'Total appointments handled by each vaccination center, grouped by region.' `
    -pal $palettes[0] `
    -columns @(
        @{ Header='Region';       Source='Region';             Width=3500; Align=1 },
        @{ Header='Center Name';  Source='CenterName';         Width=7000; Align=1 },
        @{ Header='Appointments'; Source='TotalAppointments';  Width=2800; Align=3 }
    ) `
    -reportWidth 15300

# 7 — Citizen Age Group Distribution (summary cards)
Build-SummaryReport `
    -reportName 'Rpt Citizen Age Group Distribution' `
    -queryName 'qry_Rpt_AgeGroup' `
    -title 'Citizen Age Group Distribution' `
    -subtitle 'Citizens categorised into standard public-health age brackets.' `
    -pal $palettes[2] `
    -metrics @(
        @{ Label='Total Citizens'; Expr="=DCount('*','Citizens')" },
        @{ Label='Age Groups';     Expr="=DCount('AgeGroup','qry_Rpt_AgeGroup')" }
    ) `
    -columns @(
        @{ Header='Age Group'; Source='AgeGroup';     Width=4500; Align=1 },
        @{ Header='Count';     Source='CitizenCount';  Width=3000; Align=3 }
    ) `
    -reportWidth 9500

# 8 — Batch Stock & Expiry Overview (wide layout, 7 columns)
Build-WideReport `
    -reportName 'Rpt Batch Stock and Expiry' `
    -queryName 'qry_Rpt_BatchStock' `
    -title 'Vaccine Batch Stock and Expiry Overview' `
    -subtitle 'Current stock levels and expiry status for every batch on record.' `
    -pal $palettes[3] `
    -columns @(
        @{ Header='Batch #';      Source='BatchNumber';  Width=1800; Align=1 },
        @{ Header='Vaccine';       Source='VaccineName';  Width=2600; Align=1 },
        @{ Header='Manufacturer';  Source='Manufacturer'; Width=1900; Align=1 },
        @{ Header='Stock';         Source='CurrentStock'; Width=1200; Align=3 },
        @{ Header='Expiry Date';   Source='ExpiryDate';   Width=2000; Align=2; Fmt='dd-mmm-yyyy' },
        @{ Header='Supplier';      Source='Supplier';     Width=2200; Align=1 },
        @{ Header='Alert';         Source='StockAlert';   Width=2000; Align=2 }
    ) `
    -reportWidth 15700

# 9 — Symptom Frequency (summary cards)
Build-SummaryReport `
    -reportName 'Rpt Symptom Frequency' `
    -queryName 'qry_Rpt_SymptomFreq' `
    -title 'Side Effect Symptom Frequency' `
    -subtitle 'Frequency and percentage of each reported symptom type.' `
    -pal $palettes[4] `
    -metrics @(
        @{ Label='Total Reports';     Expr="=DCount('*','SideEffects')" },
        @{ Label='Distinct Symptoms'; Expr="=DCount('SymptomType','qry_Rpt_SymptomFreq')" },
        @{ Label='Most Common';       Expr="=DLookup('SymptomType','qry_Rpt_SymptomFreq')" }
    ) `
    -columns @(
        @{ Header='Symptom';     Source='SymptomType';  Width=4500; Align=1 },
        @{ Header='Occurrences'; Source='Occurrences';  Width=2000; Align=3 },
        @{ Header='% of Total';  Source='PctOfTotal';   Width=2000; Align=3; Fmt='0.0"%"' }
    ) `
    -reportWidth 10500

# 10 — Monthly Appointment Trend with Status (tabular)
Build-TabularReport `
    -reportName 'Rpt Monthly Appointment Trend' `
    -queryName 'qry_Rpt_MonthlyApptTrend' `
    -title 'Monthly Appointment Trend by Status' `
    -subtitle 'Appointment counts broken down by month and status for trend analysis.' `
    -pal $palettes[1] `
    -columns @(
        @{ Header='Month';   Source='ApptMonth';         Width=3500; Align=1 },
        @{ Header='Status';  Source='AppointmentStatus'; Width=4000; Align=1 },
        @{ Header='Count';   Source='ApptCount';         Width=2500; Align=3 }
    ) `
    -reportWidth 12000

# ══════════════════════════════════════════════════════════
# CLEANUP
# ══════════════════════════════════════════════════════════
Log 'ALL 10 REPORTS BUILT'

# List created reports
foreach ($r in $access.CurrentProject.AllReports) {
    Log ('  REPORT: ' + $r.Name)
}

$access.DoCmd.CloseDatabase()
$access.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($db)     | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($access) | Out-Null
$access = $null
$db = $null
[GC]::Collect()
Log 'DONE'
Write-Host 'All 10 reports created successfully.' -ForegroundColor Green
