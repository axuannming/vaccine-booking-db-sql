param(
    [Parameter(Mandatory = $true)]
    [string]$FormKey
)

Add-Type -AssemblyName System.Drawing
$ErrorActionPreference = 'Stop'

$databasePath = '.\gp.accdb'
$acDetail = 0
$acForm = 2
$acLabel = 100
$acCommandButton = 104
$acTextBox = 109
$script:FieldTopOffset = 2200

function Get-OleColor([int]$r, [int]$g, [int]$b) {
    return [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::FromArgb($r, $g, $b))
}

$colors = @{
    Header = Get-OleColor 246 236 239
    White = Get-OleColor 255 255 255
    Title = Get-OleColor 33 43 68
    Text = Get-OleColor 60 60 60
    Accent = Get-OleColor 79 147 205
    Button = Get-OleColor 182 223 255
    ButtonText = Get-OleColor 35 77 118
    Card = Get-OleColor 235 246 255
    Highlight = Get-OleColor 255 249 220
    Border = Get-OleColor 185 201 214
}

function Remove-FormIfExists($access, [string]$name) {
    try { $access.DoCmd.DeleteObject($acForm, $name) } catch { }
}

function Add-BannerBackground($access, [string]$formName, [int]$width, [int]$height) {
    $banner = $access.CreateControl($formName, $acLabel, $acDetail, $null, $null, 0, 0, $width, $height)
    $banner.Caption = ''
    $banner.BackStyle = 1
    $banner.BackColor = $colors.Header
    $banner.BorderStyle = 0
    return $banner
}

function Add-TitleLabel($access, [string]$formName, [string]$caption, [int]$left, [int]$top, [int]$width) {
    $lbl = $access.CreateControl($formName, $acLabel, $acDetail, $null, $null, $left, $top, $width, 420)
    $lbl.Caption = $caption
    $lbl.FontName = 'Georgia'
    $lbl.FontSize = 18
    $lbl.FontWeight = 700
    $lbl.ForeColor = $colors.Title
    $lbl.BackStyle = 0
    return $lbl
}

function Add-InfoLabel($access, [string]$formName, [string]$caption, [int]$left, [int]$top, [int]$width, [int]$height, [int]$fontSize = 10) {
    $lbl = $access.CreateControl($formName, $acLabel, $acDetail, $null, $null, $left, $top, $width, $height)
    $lbl.Caption = $caption
    $lbl.FontName = 'Calibri'
    $lbl.FontSize = $fontSize
    $lbl.ForeColor = $colors.Text
    $lbl.BackStyle = 0
    return $lbl
}

function Add-NavButton($access, [string]$formName, [string]$caption, [string]$targetName, [int]$left, [int]$top, [int]$width = 1700) {
    $btn = $access.CreateControl($formName, $acCommandButton, $acDetail, $null, $null, $left, $top, $width, 420)
    $btn.Caption = $caption
    $btn.FontName = 'Calibri'
    $btn.FontSize = 10
    $btn.FontWeight = 700
    $btn.BackColor = $colors.Button
    $btn.ForeColor = $colors.ButtonText
    $btn.BorderColor = $colors.Accent
    $btn.ControlTipText = 'Open from navigation pane: ' + $targetName
    return $btn
}

function Add-BoundBox($access, [string]$formName, [string]$fieldName, [int]$left, [int]$top, [int]$width, [bool]$highlight = $false) {
    $textbox = $access.CreateControl($formName, $acTextBox, $acDetail, $null, $null, $left, $top, $width, 330)
    $textbox.ControlSource = $fieldName
    $textbox.FontName = 'Calibri'
    $textbox.FontSize = 10
    $textbox.ForeColor = $colors.Text
    $textbox.BorderColor = $colors.Border
    $textbox.BackColor = $(if ($highlight) { $colors.Highlight } else { $colors.White })
    $textbox.TextAlign = 1
    return $textbox
}

function Add-FieldLabel($access, [string]$formName, [string]$caption, [int]$left, [int]$top, [int]$width = 1400) {
    $label = $access.CreateControl($formName, $acLabel, $acDetail, $null, $null, $left, $top + 40, $width, 250)
    $label.Caption = $caption
    $label.FontName = 'Calibri'
    $label.FontSize = 10
    $label.FontWeight = 600
    $label.ForeColor = $colors.Text
    $label.BackStyle = 0
    return $label
}

function Add-FieldRow($access, [string]$formName, [string]$caption, [string]$fieldName, [int]$labelLeft, [int]$boxLeft, [int]$rowTop, [int]$boxWidth, [bool]$highlight = $false) {
    $y = $rowTop + $script:FieldTopOffset
    Add-FieldLabel $access $formName $caption $labelLeft $y | Out-Null
    Add-BoundBox $access $formName $fieldName $boxLeft $y $boxWidth $highlight | Out-Null
}

function Add-Metric($access, [string]$formName, [string]$title, [string]$expr, [int]$left, [int]$top) {
    $y = $top + $script:FieldTopOffset
    Add-InfoLabel $access $formName $title $left $y 2400 260 10 | Out-Null
    $textbox = $access.CreateControl($formName, $acTextBox, $acDetail, $null, $null, $left, $y + 320, 2200, 520)
    $textbox.ControlSource = $expr
    $textbox.FontName = 'Georgia'
    $textbox.FontSize = 18
    $textbox.FontWeight = 700
    $textbox.ForeColor = $colors.Title
    $textbox.BackColor = $colors.Card
    $textbox.BorderColor = $colors.Accent
    $textbox.Locked = $true
    $textbox.TabStop = $false
    $textbox.TextAlign = 2
    return $textbox
}

function Save-AndRenameForm($access, $form, [string]$targetName) {
    $tempName = $form.Name
    $access.DoCmd.Save($acForm, $tempName)
    $access.DoCmd.Close($acForm, $tempName, 1)
    $access.DoCmd.Rename($targetName, $acForm, $tempName)
}

$before = @(Get-Process MSACCESS -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
try {
    $access = New-Object -ComObject Access.Application
    $access.Visible = $false
    $access.OpenCurrentDatabase($databasePath, $false)

    switch ($FormKey) {
        'HomePage' {
            Remove-FormIfExists $access '(Portal) HomePage'
            $form = $access.CreateForm()
            $form.Caption = '(Portal) HomePage'
            $form.Width = 15000
            $form.NavigationButtons = $false
            $form.RecordSelectors = $false
            $form.DividingLines = $false
            $form.AutoCenter = $true
            $form.Section($acDetail).Height = 6500
            $form.Section($acDetail).BackColor = $colors.White
            Add-BannerBackground $access $form.Name 15000 2000 | Out-Null
            Add-TitleLabel $access $form.Name 'Vaccine Appointment and Tracking System' 450 280 11000 | Out-Null
            Add-InfoLabel $access $form.Name 'Sample-inspired portal with cleaner spacing, simplified layout, and navigation-friendly sections.' 470 760 11000 260 10 | Out-Null
            Add-NavButton $access $form.Name 'Citizen Record' '(Citizen) Medical_Record' 450 1220 1700 | Out-Null
            Add-NavButton $access $form.Name 'Appointments' '(Public) Appointment Tracker' 2350 1220 1700 | Out-Null
            Add-NavButton $access $form.Name 'Stock Monitor' '(Public) Vaccine_Stock_Monitor' 4250 1220 1700 | Out-Null
            Add-NavButton $access $form.Name 'Second Dose' '(Public) Second Dose Reminder' 6150 1220 1700 | Out-Null
            Add-NavButton $access $form.Name 'Analytics' '(Public) Analytics Dashboard' 8050 1220 1700 | Out-Null
            Add-InfoLabel $access $form.Name 'Quick Access' 700 2450 3000 360 16 | Out-Null
            Add-InfoLabel $access $form.Name 'Open the saved queries below from the buttons or the navigation pane.' 700 2880 6200 260 10 | Out-Null
            Add-NavButton $access $form.Name 'Brand Usage Query' 'qry_VaccineUsageByBrand' 900 3380 2500 | Out-Null
            Add-NavButton $access $form.Name 'Side Effect Stats' 'qry_SideEffectByBrandSeverity' 3700 3380 2500 | Out-Null
            Add-NavButton $access $form.Name 'District Rate' 'qry_DistrictVaccinationRate' 6500 3380 2200 | Out-Null
            Add-NavButton $access $form.Name 'Reminder List' 'qry_SecondDoseReminder' 9000 3380 2200 | Out-Null
            Add-InfoLabel $access $form.Name 'Main Screens' 700 4320 3000 360 16 | Out-Null
            Add-InfoLabel $access $form.Name 'Citizen records, appointment tracking, stock monitoring, reminders, and dashboard analytics now follow one consistent UI style.' 700 4750 11800 420 10 | Out-Null
            Save-AndRenameForm $access $form '(Portal) HomePage'
            Write-Output 'CREATED: (Portal) HomePage'
        }
        'CitizenRecord' {
            Remove-FormIfExists $access '(Citizen) Medical_Record'
            $form = $access.CreateForm()
            $form.RecordSource = 'qry_UI_CitizenProfile'
            $form.Caption = '(Citizen) Medical_Record'
            $form.Width = 15000
            $form.NavigationButtons = $true
            $form.RecordSelectors = $false
            $form.DividingLines = $false
            $form.AutoCenter = $true
            $form.Section($acDetail).Height = 7600
            $form.Section($acDetail).BackColor = $colors.White
            Add-BannerBackground $access $form.Name 15000 2000 | Out-Null
            Add-TitleLabel $access $form.Name 'Citizen Medical Record - Patient Information' 450 280 11000 | Out-Null
            Add-NavButton $access $form.Name 'Home' '(Portal) HomePage' 450 1220 1500 | Out-Null
            Add-NavButton $access $form.Name 'Appointments' '(Public) Appointment Tracker' 2200 1220 1700 | Out-Null
            Add-NavButton $access $form.Name 'Stock' '(Public) Vaccine_Stock_Monitor' 4150 1220 1500 | Out-Null
            Add-NavButton $access $form.Name 'Reminder' '(Public) Second Dose Reminder' 5900 1220 1600 | Out-Null
            Add-NavButton $access $form.Name 'Analytics' '(Public) Analytics Dashboard' 7750 1220 1600 | Out-Null
            Add-FieldRow $access $form.Name 'Citizen ID' 'CitizenID' 500 2100 200 2400 $true
            Add-FieldRow $access $form.Name 'HKID' 'HKID' 7600 9300 200 2600 $true
            Add-FieldRow $access $form.Name 'First Name' 'FirstName' 500 2100 700 2400
            Add-FieldRow $access $form.Name 'Last Name' 'LastName' 7600 9300 700 2600
            Add-FieldRow $access $form.Name 'Age' 'CitizenAge' 500 2100 1200 1400
            Add-FieldRow $access $form.Name 'Gender' 'Gender' 7600 9300 1200 1600
            Add-FieldRow $access $form.Name 'Phone' 'Phone' 500 2100 1700 2400
            Add-FieldRow $access $form.Name 'District' 'District' 7600 9300 1700 2200
            Add-FieldRow $access $form.Name 'Email' 'Email' 500 2100 2200 5200
            Add-FieldRow $access $form.Name 'Address' 'Address' 500 2100 2700 5200
            Add-FieldRow $access $form.Name 'Occupation' 'Occupation' 500 2100 3200 2400
            Add-FieldRow $access $form.Name 'Place of Birth' 'PlaceOfBirth' 7600 9300 3200 2600
            Add-FieldRow $access $form.Name 'Allergy History' 'AllergyHistory' 500 2100 3700 5200
            Save-AndRenameForm $access $form '(Citizen) Medical_Record'
            Write-Output 'CREATED: (Citizen) Medical_Record'
        }
        'AppointmentTracker' {
            Remove-FormIfExists $access '(Public) Appointment Tracker'
            $form = $access.CreateForm()
            $form.RecordSource = 'qry_UI_Appointments'
            $form.Caption = '(Public) Appointment Tracker'
            $form.Width = 15000
            $form.NavigationButtons = $true
            $form.RecordSelectors = $false
            $form.DividingLines = $false
            $form.AutoCenter = $true
            $form.Section($acDetail).Height = 7200
            $form.Section($acDetail).BackColor = $colors.White
            Add-BannerBackground $access $form.Name 15000 2000 | Out-Null
            Add-TitleLabel $access $form.Name 'Vaccination Appointment Tracker' 450 280 10000 | Out-Null
            Add-NavButton $access $form.Name 'Home' '(Portal) HomePage' 450 1220 1500 | Out-Null
            Add-NavButton $access $form.Name 'Citizen' '(Citizen) Medical_Record' 2200 1220 1500 | Out-Null
            Add-NavButton $access $form.Name 'Stock' '(Public) Vaccine_Stock_Monitor' 3950 1220 1500 | Out-Null
            Add-NavButton $access $form.Name 'Reminder' '(Public) Second Dose Reminder' 5700 1220 1600 | Out-Null
            Add-NavButton $access $form.Name 'Analytics' '(Public) Analytics Dashboard' 7550 1220 1600 | Out-Null
            Add-FieldRow $access $form.Name 'Appointment ID' 'AppointmentID' 500 2200 200 2200 $true
            Add-FieldRow $access $form.Name 'Citizen Name' 'CitizenName' 7600 9300 200 3000
            Add-FieldRow $access $form.Name 'HKID' 'HKID' 500 2200 700 2200
            Add-FieldRow $access $form.Name 'Vaccine' 'VaccineName' 7600 9300 700 3000
            Add-FieldRow $access $form.Name 'Center' 'CenterName' 500 2200 1200 4200
            Add-FieldRow $access $form.Name 'Dose No.' 'DoseNo' 7600 9300 1200 1400
            Add-FieldRow $access $form.Name 'Appointment Date' 'AppointmentDate' 500 2200 1700 2200
            Add-FieldRow $access $form.Name 'Status' 'AppointmentStatus' 7600 9300 1700 2200
            Add-FieldRow $access $form.Name 'Actual Vaccination' 'ActualVaccinationDate' 500 2200 2200 2200
            Add-FieldRow $access $form.Name 'Cancellation Reason' 'CancellationReason' 7600 9300 2200 3000
            Add-FieldRow $access $form.Name 'Notes' 'Notes' 500 2200 2700 10100
            Save-AndRenameForm $access $form '(Public) Appointment Tracker'
            Write-Output 'CREATED: (Public) Appointment Tracker'
        }
        'StockMonitor' {
            Remove-FormIfExists $access '(Public) Vaccine_Stock_Monitor'
            $form = $access.CreateForm()
            $form.RecordSource = 'qry_ExpiringOrLowStockBatches'
            $form.Caption = '(Public) Vaccine_Stock_Monitor'
            $form.Width = 15000
            $form.NavigationButtons = $true
            $form.RecordSelectors = $false
            $form.DividingLines = $false
            $form.AutoCenter = $true
            $form.Section($acDetail).Height = 6200
            $form.Section($acDetail).BackColor = $colors.White
            Add-BannerBackground $access $form.Name 15000 2000 | Out-Null
            Add-TitleLabel $access $form.Name 'Vaccine Stock and Expiry Monitor' 450 280 10000 | Out-Null
            Add-NavButton $access $form.Name 'Home' '(Portal) HomePage' 450 1220 1500 | Out-Null
            Add-NavButton $access $form.Name 'Citizen' '(Citizen) Medical_Record' 2200 1220 1500 | Out-Null
            Add-NavButton $access $form.Name 'Appointments' '(Public) Appointment Tracker' 3950 1220 1800 | Out-Null
            Add-NavButton $access $form.Name 'Reminder' '(Public) Second Dose Reminder' 6000 1220 1600 | Out-Null
            Add-NavButton $access $form.Name 'Analytics' '(Public) Analytics Dashboard' 7850 1220 1600 | Out-Null
            Add-FieldRow $access $form.Name 'Batch ID' 'BatchID' 500 2200 300 1800 $true
            Add-FieldRow $access $form.Name 'Batch Number' 'BatchNumber' 7600 9300 300 2500
            Add-FieldRow $access $form.Name 'Vaccine' 'VaccineName' 500 2200 850 3200
            Add-FieldRow $access $form.Name 'Expiry Date' 'ExpiryDate' 7600 9300 850 2200
            Add-FieldRow $access $form.Name 'Current Stock' 'CurrentStock' 500 2200 1400 1800
            Add-FieldRow $access $form.Name 'Stock Flag' 'StockFlag' 7600 9300 1400 2200
            Add-FieldRow $access $form.Name 'Supplier' 'Supplier' 500 2200 1950 3200
            Add-FieldRow $access $form.Name 'Expiry Flag' 'ExpiryFlag' 7600 9300 1950 2200
            Add-InfoLabel $access $form.Name 'This screen shows the exact batches that need attention because of low stock or near expiry.' 500 5000 11000 260 10 | Out-Null
            Save-AndRenameForm $access $form '(Public) Vaccine_Stock_Monitor'
            Write-Output 'CREATED: (Public) Vaccine_Stock_Monitor'
        }
        'SecondDoseReminder' {
            Remove-FormIfExists $access '(Public) Second Dose Reminder'
            $form = $access.CreateForm()
            $form.RecordSource = 'qry_SecondDoseReminder'
            $form.Caption = '(Public) Second Dose Reminder'
            $form.Width = 15000
            $form.NavigationButtons = $true
            $form.RecordSelectors = $false
            $form.DividingLines = $false
            $form.AutoCenter = $true
            $form.Section($acDetail).Height = 6000
            $form.Section($acDetail).BackColor = $colors.White
            Add-BannerBackground $access $form.Name 15000 2000 | Out-Null
            Add-TitleLabel $access $form.Name 'Second Dose Reminder List' 450 280 10000 | Out-Null
            Add-NavButton $access $form.Name 'Home' '(Portal) HomePage' 450 1220 1500 | Out-Null
            Add-NavButton $access $form.Name 'Citizen' '(Citizen) Medical_Record' 2200 1220 1500 | Out-Null
            Add-NavButton $access $form.Name 'Appointments' '(Public) Appointment Tracker' 3950 1220 1800 | Out-Null
            Add-NavButton $access $form.Name 'Stock' '(Public) Vaccine_Stock_Monitor' 6000 1220 1500 | Out-Null
            Add-NavButton $access $form.Name 'Analytics' '(Public) Analytics Dashboard' 7750 1220 1600 | Out-Null
            Add-FieldRow $access $form.Name 'Citizen ID' 'CitizenID' 500 2200 300 1800 $true
            Add-FieldRow $access $form.Name 'Citizen Name' 'CitizenName' 7600 9300 300 3000
            Add-FieldRow $access $form.Name 'Vaccine' 'VaccineName' 500 2200 850 3200
            Add-FieldRow $access $form.Name 'Days Overdue' 'DaysOverdue' 7600 9300 850 1800 $true
            Add-FieldRow $access $form.Name 'First Dose Date' 'FirstDoseDate' 500 2200 1400 2200
            Add-FieldRow $access $form.Name 'Recommended Date' 'RecommendedSecondDoseDate' 7600 9300 1400 2200
            Add-InfoLabel $access $form.Name 'This list can be used for phone, SMS, or email follow-up.' 500 4500 9000 260 10 | Out-Null
            Save-AndRenameForm $access $form '(Public) Second Dose Reminder'
            Write-Output 'CREATED: (Public) Second Dose Reminder'
        }
        'AnalyticsDashboard' {
            Remove-FormIfExists $access '(Public) Analytics Dashboard'
            $form = $access.CreateForm()
            $form.Caption = '(Public) Analytics Dashboard'
            $form.Width = 15000
            $form.NavigationButtons = $false
            $form.RecordSelectors = $false
            $form.DividingLines = $false
            $form.AutoCenter = $true
            $form.Section($acDetail).Height = 7000
            $form.Section($acDetail).BackColor = $colors.White
            Add-BannerBackground $access $form.Name 15000 2000 | Out-Null
            Add-TitleLabel $access $form.Name 'Analytics Dashboard' 450 280 10000 | Out-Null
            Add-NavButton $access $form.Name 'Home' '(Portal) HomePage' 450 1220 1500 | Out-Null
            Add-NavButton $access $form.Name 'Citizen' '(Citizen) Medical_Record' 2200 1220 1500 | Out-Null
            Add-NavButton $access $form.Name 'Appointments' '(Public) Appointment Tracker' 3950 1220 1800 | Out-Null
            Add-NavButton $access $form.Name 'Stock' '(Public) Vaccine_Stock_Monitor' 6000 1220 1500 | Out-Null
            Add-NavButton $access $form.Name 'Reminder' '(Public) Second Dose Reminder' 7750 1220 1600 | Out-Null
            Add-Metric $access $form.Name 'Total Citizens' '=DCount("*","Citizens")' 700 300 | Out-Null
            Add-Metric $access $form.Name 'Total Appointments' '=DCount("*","Appointments")' 3400 300 | Out-Null
            Add-Metric $access $form.Name 'Completed Doses' '=DCount("*","Appointments","AppointmentStatus=''Completed''")' 6100 300 | Out-Null
            Add-Metric $access $form.Name 'Overdue Second Doses' '=DCount("*","qry_SecondDoseReminder")' 8800 300 | Out-Null
            Add-Metric $access $form.Name 'Low Stock / Expiring' '=DCount("*","qry_ExpiringOrLowStockBatches")' 700 1500 | Out-Null
            Add-Metric $access $form.Name 'Active Centers' '=DCount("*","VaccinationCenters","IsActive=True")' 3400 1500 | Out-Null
            Add-InfoLabel $access $form.Name 'Open statistical queries from the navigation pane or the shortcut buttons below.' 700 4700 8200 260 10 | Out-Null
            Add-NavButton $access $form.Name 'Brand Usage' 'qry_VaccineUsageByBrand' 700 5150 1800 | Out-Null
            Add-NavButton $access $form.Name 'Side Effects' 'qry_SideEffectByBrandSeverity' 2800 5150 1800 | Out-Null
            Add-NavButton $access $form.Name 'District Rate' 'qry_DistrictVaccinationRate' 4900 5150 1800 | Out-Null
            Add-NavButton $access $form.Name 'No-show Rate' 'qry_NoShowRate' 7000 5150 1800 | Out-Null
            Save-AndRenameForm $access $form '(Public) Analytics Dashboard'
            Write-Output 'CREATED: (Public) Analytics Dashboard'
        }
        default {
            throw "Unsupported FormKey: $FormKey"
        }
    }
} finally {
    try { if ($access) { $access.Quit(2) } } catch { }
    if ($access) { [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($access) }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    Start-Sleep -Seconds 2
    $after = @(Get-Process MSACCESS -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
    foreach ($procId in $after) {
        if ($before -notcontains $procId) {
            Stop-Process -Id $procId -Force -ErrorAction SilentlyContinue
        }
    }
}
