Add-Type -AssemblyName System.Drawing
$ErrorActionPreference = 'Stop'

$databasePath = '.\gp.accdb'
$logPath = '.\build_access_ui.log'

$acDetail = 0
$acForm = 2
$acLabel = 100
$acCommandButton = 104
$acCheckBox = 106
$acTextBox = 109
$acComboBox = 111
$dbFailOnError = 128

Set-Content -Path $logPath -Value @('START ' + (Get-Date).ToString('s'))

function Log([string]$message) {
    Add-Content -Path $logPath -Value ((Get-Date).ToString('s') + ' ' + $message)
}

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

function Close-OpenForms($access) {
    try {
        foreach ($item in $access.CurrentProject.AllForms) {
            try {
                if ($item.IsLoaded) {
                    $access.DoCmd.Close($acForm, $item.Name, 2)
                }
            } catch { }
        }
    } catch { }
}

function Field-Exists($db, [string]$tableName, [string]$fieldName) {
    foreach ($field in $db.TableDefs($tableName).Fields) {
        if ($field.Name -eq $fieldName) {
            return $true
        }
    }
    return $false
}

function Ensure-Field($db, [string]$tableName, [string]$fieldName, [string]$fieldDefinition) {
    if (-not (Field-Exists $db $tableName $fieldName)) {
        $db.Execute("ALTER TABLE [$tableName] ADD COLUMN [$fieldName] $fieldDefinition", $dbFailOnError)
        Log ("ADDED FIELD " + $tableName + "." + $fieldName)
    }
}

function Ensure-CitizenFields($db) {
    Ensure-Field $db 'Citizens' 'FirstName' 'TEXT(50)'
    Ensure-Field $db 'Citizens' 'LastName' 'TEXT(50)'
    Ensure-Field $db 'Citizens' 'Occupation' 'TEXT(100)'
    Ensure-Field $db 'Citizens' 'PlaceOfBirth' 'TEXT(100)'
}

function Add-OrReplaceQuery($db, [string]$queryName, [string]$sql) {
    try { $db.QueryDefs.Delete($queryName) } catch { }
    [void]$db.CreateQueryDef($queryName, $sql)
    Log ('QUERY READY ' + $queryName)
}

function Ensure-UiQueries($db) {
    $citizenSql = @'
SELECT
    Citizens.CitizenID,
    Citizens.HKID,
    Citizens.CitizenName,
    Citizens.FirstName,
    Citizens.LastName,
    Citizens.Gender,
    Citizens.DOB,
    DateDiff("yyyy",[DOB],Date()) - IIf(Format([DOB],"mmdd")>Format(Date(),"mmdd"),1,0) AS CitizenAge,
    IIf(DateDiff("yyyy",[DOB],Date()) - IIf(Format([DOB],"mmdd")>Format(Date(),"mmdd"),1,0)<=11,"0-11",IIf(DateDiff("yyyy",[DOB],Date()) - IIf(Format([DOB],"mmdd")>Format(Date(),"mmdd"),1,0)<=17,"12-17",IIf(DateDiff("yyyy",[DOB],Date()) - IIf(Format([DOB],"mmdd")>Format(Date(),"mmdd"),1,0)<=29,"18-29",IIf(DateDiff("yyyy",[DOB],Date()) - IIf(Format([DOB],"mmdd")>Format(Date(),"mmdd"),1,0)<=44,"30-44",IIf(DateDiff("yyyy",[DOB],Date()) - IIf(Format([DOB],"mmdd")>Format(Date(),"mmdd"),1,0)<=59,"45-59","60+"))))) AS AgeGroup,
    Citizens.Phone,
    Citizens.Email,
    Citizens.[Address],
    Citizens.District,
    Citizens.Occupation,
    Citizens.PlaceOfBirth,
    Citizens.AllergyHistory
FROM Citizens
ORDER BY Citizens.CitizenID;
'@

    $appointmentSql = @'
SELECT
    Appointments.AppointmentID,
    Appointments.CitizenID,
    Citizens.HKID,
    Citizens.CitizenName,
    Appointments.VaccineID,
    Vaccines.VaccineName,
    Appointments.CenterID,
    VaccinationCenters.CenterName,
    Appointments.SlotID,
    AppointmentSlots.SlotDate,
    Format([AppointmentSlots].[SlotTime],"hh:nn AM/PM") AS SlotTimeLabel,
    Appointments.BatchID,
    VaccineBatches.BatchNumber,
    Appointments.DoseNo,
    Appointments.AppointmentDate,
    Appointments.AppointmentStatus,
    Appointments.ActualVaccinationDate,
    Appointments.VaccinatedBy,
    Users.Username AS VaccinatedByName,
    Appointments.Notes,
    Appointments.CancellationReason
FROM ((((((Appointments
INNER JOIN Citizens ON Appointments.CitizenID = Citizens.CitizenID)
INNER JOIN Vaccines ON Appointments.VaccineID = Vaccines.VaccineID)
INNER JOIN VaccinationCenters ON Appointments.CenterID = VaccinationCenters.CenterID)
LEFT JOIN AppointmentSlots ON Appointments.SlotID = AppointmentSlots.SlotID)
LEFT JOIN VaccineBatches ON Appointments.BatchID = VaccineBatches.BatchID)
LEFT JOIN Users ON Appointments.VaccinatedBy = Users.UserID)
ORDER BY Appointments.AppointmentDate DESC, Appointments.AppointmentID DESC;
'@

    Add-OrReplaceQuery $db 'qry_UI_CitizenProfile' $citizenSql
    Add-OrReplaceQuery $db 'qry_UI_Appointments' $appointmentSql
}

function Add-BannerBackground($access, [string]$formName, [int]$width, [int]$height) {
    $banner = $access.CreateControl($formName, $acLabel, $acDetail, $null, $null, 0, 0, $width, $height)
    $banner.Caption = ''
    $banner.BackStyle = 1
    $banner.BackColor = $colors.Header
    $banner.BorderStyle = 0
    return $banner
}

function Add-TitleLabel($access, [string]$formName, [string]$caption, [int]$left, [int]$top, [int]$width, [int]$fontSize = 18) {
    $lbl = $access.CreateControl($formName, $acLabel, $acDetail, $null, $null, $left, $top, $width, 430)
    $lbl.Caption = $caption
    $lbl.FontName = 'Georgia'
    $lbl.FontSize = $fontSize
    $lbl.FontWeight = 700
    $lbl.ForeColor = $colors.Title
    $lbl.BackStyle = 0
    return $lbl
}

function Add-InfoLabel($access, [string]$formName, [string]$caption, [int]$left, [int]$top, [int]$width, [int]$height, [int]$fontSize = 10, [bool]$bold = $false) {
    $lbl = $access.CreateControl($formName, $acLabel, $acDetail, $null, $null, $left, $top, $width, $height)
    $lbl.Caption = $caption
    $lbl.FontName = 'Calibri'
    $lbl.FontSize = $fontSize
    $lbl.FontWeight = $(if ($bold) { 700 } else { 400 })
    $lbl.ForeColor = $colors.Text
    $lbl.BackStyle = 0
    return $lbl
}

function Add-FieldLabel($access, [string]$formName, [string]$caption, [int]$left, [int]$top, [int]$width = 1500) {
    $label = $access.CreateControl($formName, $acLabel, $acDetail, $null, $null, $left, $top + 40, $width, 260)
    $label.Caption = $caption
    $label.FontName = 'Calibri'
    $label.FontSize = 10
    $label.FontWeight = 600
    $label.ForeColor = $colors.Text
    $label.BackStyle = 0
    return $label
}

function Add-BoundTextControl($access, [string]$formName, [string]$fieldName, [int]$left, [int]$top, [int]$width, [int]$height = 330, [bool]$locked = $false, [bool]$highlight = $false) {
    $textbox = $access.CreateControl($formName, $acTextBox, $acDetail, $null, $null, $left, $top, $width, $height)
    $textbox.ControlSource = $fieldName
    $textbox.FontName = 'Calibri'
    $textbox.FontSize = 10
    $textbox.ForeColor = $colors.Text
    $textbox.BorderColor = $colors.Border
    $textbox.BackColor = $(if ($highlight) { $colors.Highlight } else { $colors.White })
    $textbox.TextAlign = 1
    if ($locked) {
        $textbox.Locked = $true
        $textbox.TabStop = $false
    }
    return $textbox
}

function Add-TextRow($access, [string]$formName, [string]$caption, [string]$fieldName, [int]$labelLeft, [int]$boxLeft, [int]$top, [int]$boxWidth, [bool]$locked = $false, [bool]$highlight = $false, [int]$boxHeight = 330) {
    Add-FieldLabel $access $formName $caption $labelLeft $top | Out-Null
    return Add-BoundTextControl $access $formName $fieldName $boxLeft $top $boxWidth $boxHeight $locked $highlight
}

function Add-BoundComboControl($access, [string]$formName, [string]$fieldName, [int]$left, [int]$top, [int]$width, [string]$rowSourceType, [string]$rowSource, [int]$columnCount = 1, [int]$boundColumn = 1, [string]$columnWidths = '', [int]$listWidth = 0) {
    $combo = $access.CreateControl($formName, $acComboBox, $acDetail, $null, $null, $left, $top, $width, 330)
    $combo.ControlSource = $fieldName
    $combo.RowSourceType = $rowSourceType
    $combo.RowSource = $rowSource
    $combo.ColumnCount = $columnCount
    $combo.BoundColumn = $boundColumn
    if ($columnWidths -ne '') {
        $combo.ColumnWidths = $columnWidths
    }
    if ($listWidth -gt 0) {
        $combo.ListWidth = $listWidth
    }
    $combo.LimitToList = $true
    $combo.AutoExpand = $true
    $combo.ListRows = 12
    $combo.FontName = 'Calibri'
    $combo.FontSize = 10
    $combo.ForeColor = $colors.Text
    $combo.BorderColor = $colors.Border
    $combo.BackColor = $colors.White
    return $combo
}

function Add-ComboRow($access, [string]$formName, [string]$caption, [string]$fieldName, [int]$labelLeft, [int]$boxLeft, [int]$top, [int]$boxWidth, [string]$rowSourceType, [string]$rowSource, [int]$columnCount = 1, [int]$boundColumn = 1, [string]$columnWidths = '', [int]$listWidth = 0) {
    Add-FieldLabel $access $formName $caption $labelLeft $top | Out-Null
    return Add-BoundComboControl $access $formName $fieldName $boxLeft $top $boxWidth $rowSourceType $rowSource $columnCount $boundColumn $columnWidths $listWidth
}

function Add-BoundCheckControl($access, [string]$formName, [string]$fieldName, [int]$left, [int]$top) {
    $check = $access.CreateControl($formName, $acCheckBox, $acDetail, $null, $null, $left, $top + 40, 260, 260)
    $check.ControlSource = $fieldName
    return $check
}

function Add-CheckRow($access, [string]$formName, [string]$caption, [string]$fieldName, [int]$labelLeft, [int]$boxLeft, [int]$top) {
    Add-FieldLabel $access $formName $caption $labelLeft $top | Out-Null
    return Add-BoundCheckControl $access $formName $fieldName $boxLeft $top
}

function Add-MetricBox($access, [string]$formName, [string]$title, [string]$expr, [int]$left, [int]$top) {
    Add-InfoLabel $access $formName $title $left $top 2300 260 10 $true | Out-Null
    $textbox = $access.CreateControl($formName, $acTextBox, $acDetail, $null, $null, $left, $top + 320, 2200, 560)
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

function Add-Procedure([System.Collections.ArrayList]$events, [string]$buttonName, [string]$body) {
    $procedure = "Private Sub ${buttonName}_Click()`r`n$body`r`nEnd Sub`r`n`r`n"
    [void]$events.Add($procedure)
}

function Add-EventButton($access, [string]$formName, [System.Collections.ArrayList]$events, [string]$buttonName, [string]$caption, [int]$left, [int]$top, [int]$width, [string]$body, [string]$toolTip = '') {
    $btn = $access.CreateControl($formName, $acCommandButton, $acDetail, $null, $null, $left, $top, $width, 420)
    $btn.Name = $buttonName
    $btn.Caption = $caption
    $btn.OnClick = '[Event Procedure]'
    $btn.FontName = 'Calibri'
    $btn.FontSize = 10
    $btn.FontWeight = 700
    $btn.BackColor = $colors.Button
    $btn.ForeColor = $colors.ButtonText
    $btn.BorderColor = $colors.Accent
    if ($toolTip -ne '') {
        $btn.ControlTipText = $toolTip
    }
    Add-Procedure $events $buttonName $body
    return $btn
}

function Add-FormOpenButton($access, [string]$formName, [System.Collections.ArrayList]$events, [string]$buttonName, [string]$caption, [string]$targetForm, [int]$left, [int]$top, [int]$width) {
    $body = '    DoCmd.OpenForm "' + $targetForm + '"'
    return Add-EventButton $access $formName $events $buttonName $caption $left $top $width $body ('Open form: ' + $targetForm)
}

function Add-QueryOpenButton($access, [string]$formName, [System.Collections.ArrayList]$events, [string]$buttonName, [string]$caption, [string]$targetQuery, [int]$left, [int]$top, [int]$width) {
    $body = '    DoCmd.OpenQuery "' + $targetQuery + '"'
    return Add-EventButton $access $formName $events $buttonName $caption $left $top $width $body ('Open query: ' + $targetQuery)
}

function Add-StandardNavBar($access, [string]$formName, [System.Collections.ArrayList]$events) {
    $rows = @(
        @(
            @{ Name = 'btnNavHome'; Caption = 'Home'; Target = '(Portal) HomePage'; Width = 1000 },
            @{ Name = 'btnNavCitizenView'; Caption = 'Citizen View'; Target = '(Citizen) Medical_Record'; Width = 1250 },
            @{ Name = 'btnNavCitizenEntry'; Caption = 'Citizen Entry'; Target = 'Citizen Entry Form'; Width = 1300 },
            @{ Name = 'btnNavApptView'; Caption = 'Appt View'; Target = '(Public) Appointment Tracker'; Width = 1200 },
            @{ Name = 'btnNavApptEntry'; Caption = 'Appt Entry'; Target = 'Appointment Entry Form'; Width = 1250 },
            @{ Name = 'btnNavVaxEntry'; Caption = 'Vax Record'; Target = 'Vaccination Record Entry Form'; Width = 1250 },
            @{ Name = 'btnNavScreening'; Caption = 'Screening'; Target = 'Screening Entry Form'; Width = 1200 }
        ),
        @(
            @{ Name = 'btnNavConsent'; Caption = 'Consent'; Target = 'Consent Entry Form'; Width = 1100 },
            @{ Name = 'btnNavObservation'; Caption = 'Observation'; Target = 'Observation Entry Form'; Width = 1250 },
            @{ Name = 'btnNavFollowup'; Caption = 'Follow-up'; Target = 'Adverse Follow-up Form'; Width = 1200 },
            @{ Name = 'btnNavInventory'; Caption = 'Inventory'; Target = 'Inventory Transaction Form'; Width = 1200 },
            @{ Name = 'btnNavStock'; Caption = 'Stock'; Target = '(Public) Vaccine_Stock_Monitor'; Width = 1000 },
            @{ Name = 'btnNavReminder'; Caption = 'Reminder'; Target = '(Public) Second Dose Reminder'; Width = 1100 },
            @{ Name = 'btnNavAnalytics'; Caption = 'Analytics'; Target = '(Public) Analytics Dashboard'; Width = 1200 }
        )
    )

    $tops = @(1180, 1630)
    for ($r = 0; $r -lt $rows.Count; $r++) {
        $left = 250
        foreach ($item in $rows[$r]) {
            Add-FormOpenButton $access $formName $events $item.Name $item.Caption $item.Target $left $tops[$r] $item.Width | Out-Null
            $left += ($item.Width + 110)
        }
    }
}

function Write-FormProcedures($form, [System.Collections.ArrayList]$events) {
    if ($events.Count -eq 0) {
        return
    }
    $module = $form.Module
    foreach ($procedure in $events) {
        $module.InsertLines($module.CountOfLines + 1, $procedure)
    }
}

function Save-AndRenameForm($access, $form, [string]$targetName) {
    $tempName = $form.Name
    $access.DoCmd.Save($acForm, $tempName)
    $access.DoCmd.Close($acForm, $tempName, 1)
    $access.DoCmd.Rename($targetName, $acForm, $tempName)
}

function Build-HomePage($access) {
    $events = [System.Collections.ArrayList]::new()
    $form = $access.CreateForm()
    $form.HasModule = $true
    $form.Caption = '(Portal) HomePage'
    $form.Width = 15000
    $form.NavigationButtons = $false
    $form.RecordSelectors = $false
    $form.DividingLines = $false
    $form.AllowEdits = $false
    $form.AllowAdditions = $false
    $form.AllowDeletions = $false
    $form.AutoCenter = $true
    $form.Section($acDetail).Height = 9300
    $form.Section($acDetail).BackColor = $colors.White

    Add-BannerBackground $access $form.Name 15000 2200 | Out-Null
    Add-TitleLabel $access $form.Name 'Vaccine Registration System v2' 420 260 11500 18 | Out-Null
    Add-InfoLabel $access $form.Name 'Expanded workflow: booking, screening, consent, vaccination, observation, side effects, follow-up, and inventory.' 430 760 13000 280 10 | Out-Null
    Add-StandardNavBar $access $form.Name $events

    Add-InfoLabel $access $form.Name 'Live Summary' 650 2550 2500 340 16 $true | Out-Null
    Add-MetricBox $access $form.Name 'Total Citizens' '=DCount("*","Citizens")' 700 2980 | Out-Null
    Add-MetricBox $access $form.Name 'Appointments' '=DCount("*","Appointments")' 3400 2980 | Out-Null
    Add-MetricBox $access $form.Name 'Vaccination Records' '=DCount("*","VaccinationRecords")' 6100 2980 | Out-Null
    Add-MetricBox $access $form.Name 'Due Reminders' '=DCount("*","qry_SecondDoseReminder")' 8800 2980 | Out-Null

    Add-InfoLabel $access $form.Name 'Quick Actions' 650 4380 2500 340 16 $true | Out-Null
    Add-InfoLabel $access $form.Name 'Use these to jump straight into the end-to-end vaccination workflow.' 700 4760 10000 260 10 | Out-Null
    Add-FormOpenButton $access $form.Name $events 'btnHomeCitizenEntry' 'Citizen Entry Form' 'Citizen Entry Form' 850 5240 2300 | Out-Null
    Add-FormOpenButton $access $form.Name $events 'btnHomeApptEntry' 'Appointment Entry Form' 'Appointment Entry Form' 3450 5240 2300 | Out-Null
    Add-FormOpenButton $access $form.Name $events 'btnHomeRecordEntry' 'Vaccination Record Form' 'Vaccination Record Entry Form' 6050 5240 2500 | Out-Null
    Add-FormOpenButton $access $form.Name $events 'btnHomeInventoryEntry' 'Inventory Transaction Form' 'Inventory Transaction Form' 8850 5240 2500 | Out-Null

    Add-InfoLabel $access $form.Name 'Safety Flow' 650 6200 2500 340 16 $true | Out-Null
    Add-InfoLabel $access $form.Name 'Suggested order: Appointment -> Screening -> Consent -> Vaccination Record -> Observation -> Side Effect Follow-up.' 700 6580 12400 260 10 | Out-Null
    Add-FormOpenButton $access $form.Name $events 'btnHomeScreening' 'Screening' 'Screening Entry Form' 850 7030 1800 | Out-Null
    Add-FormOpenButton $access $form.Name $events 'btnHomeConsent' 'Consent' 'Consent Entry Form' 2950 7030 1800 | Out-Null
    Add-FormOpenButton $access $form.Name $events 'btnHomeObservation' 'Observation' 'Observation Entry Form' 5050 7030 1800 | Out-Null
    Add-FormOpenButton $access $form.Name $events 'btnHomeFollowup' 'Follow-up' 'Adverse Follow-up Form' 7150 7030 1800 | Out-Null
    Add-QueryOpenButton $access $form.Name $events 'btnHomeAnalytics' 'Age Group Stats' 'qry_AgeGroupSideEffectRate' 9250 7030 2100 | Out-Null

    Add-InfoLabel $access $form.Name 'Forms and data sources:' 700 7900 3000 260 10 $true | Out-Null
    Add-InfoLabel $access $form.Name 'Citizen Medical Record -> qry_UI_CitizenProfile | Appointment Tracker -> qry_UI_Appointments | Entry forms -> direct workflow tables.' 700 8200 12800 420 10 | Out-Null

    Write-FormProcedures $form $events
    Save-AndRenameForm $access $form '(Portal) HomePage'
    Log 'FORM READY (Portal) HomePage'
}

function Build-CitizenRecordForm($access) {
    $events = [System.Collections.ArrayList]::new()
    $form = $access.CreateForm()
    $form.HasModule = $true
    $form.RecordSource = 'qry_UI_CitizenProfile'
    $form.Caption = '(Citizen) Medical_Record'
    $form.Width = 15000
    $form.NavigationButtons = $true
    $form.RecordSelectors = $false
    $form.DividingLines = $false
    $form.AllowEdits = $false
    $form.AllowAdditions = $false
    $form.AllowDeletions = $false
    $form.AutoCenter = $true
    $form.Section($acDetail).Height = 8100
    $form.Section($acDetail).BackColor = $colors.White

    Add-BannerBackground $access $form.Name 15000 2200 | Out-Null
    Add-TitleLabel $access $form.Name 'Citizen Medical Record' 420 260 10000 18 | Out-Null
    Add-InfoLabel $access $form.Name 'Bound to query: qry_UI_CitizenProfile' 430 760 6000 260 10 | Out-Null
    Add-StandardNavBar $access $form.Name $events

    Add-TextRow $access $form.Name 'Citizen ID' 'CitizenID' 500 2200 2600 2400 $true $true | Out-Null
    Add-TextRow $access $form.Name 'HKID' 'HKID' 7600 9200 2600 2600 $true $true | Out-Null
    Add-TextRow $access $form.Name 'First Name' 'FirstName' 500 2200 3150 2400 $true | Out-Null
    Add-TextRow $access $form.Name 'Last Name' 'LastName' 7600 9200 3150 2600 $true | Out-Null
    Add-TextRow $access $form.Name 'Age' 'CitizenAge' 500 2200 3700 1400 $true | Out-Null
    Add-TextRow $access $form.Name 'Age Group' 'AgeGroup' 7600 9200 3700 1600 $true $true | Out-Null
    Add-TextRow $access $form.Name 'Phone' 'Phone' 500 2200 4250 2400 $true | Out-Null
    Add-TextRow $access $form.Name 'Gender' 'Gender' 7600 9200 4250 1600 $true | Out-Null
    Add-TextRow $access $form.Name 'Email' 'Email' 500 2200 4800 5200 $true | Out-Null
    Add-TextRow $access $form.Name 'District' 'District' 7600 9200 4800 2200 $true | Out-Null
    Add-TextRow $access $form.Name 'Address' 'Address' 500 2200 5350 5200 $true | Out-Null
    Add-TextRow $access $form.Name 'Occupation' 'Occupation' 7600 9200 5350 2600 $true | Out-Null
    Add-TextRow $access $form.Name 'Place of Birth' 'PlaceOfBirth' 500 2200 5900 2400 $true | Out-Null
    Add-TextRow $access $form.Name 'Allergy History' 'AllergyHistory' 500 2200 6450 5200 $true | Out-Null

    Write-FormProcedures $form $events
    Save-AndRenameForm $access $form '(Citizen) Medical_Record'
    Log 'FORM READY (Citizen) Medical_Record'
}

function Build-AppointmentTrackerForm($access) {
    $events = [System.Collections.ArrayList]::new()
    $form = $access.CreateForm()
    $form.HasModule = $true
    $form.RecordSource = 'qry_UI_Appointments'
    $form.Caption = '(Public) Appointment Tracker'
    $form.Width = 15000
    $form.NavigationButtons = $true
    $form.RecordSelectors = $false
    $form.DividingLines = $false
    $form.AllowEdits = $false
    $form.AllowAdditions = $false
    $form.AllowDeletions = $false
    $form.AutoCenter = $true
    $form.Section($acDetail).Height = 7600
    $form.Section($acDetail).BackColor = $colors.White

    Add-BannerBackground $access $form.Name 15000 2200 | Out-Null
    Add-TitleLabel $access $form.Name 'Vaccination Appointment Tracker' 420 260 10000 18 | Out-Null
    Add-InfoLabel $access $form.Name 'Bound to query: qry_UI_Appointments' 430 760 6000 260 10 | Out-Null
    Add-StandardNavBar $access $form.Name $events

    Add-TextRow $access $form.Name 'Appointment ID' 'AppointmentID' 500 2200 2600 2200 $true $true | Out-Null
    Add-TextRow $access $form.Name 'Citizen Name' 'CitizenName' 7600 9200 2600 3000 $true | Out-Null
    Add-TextRow $access $form.Name 'HKID' 'HKID' 500 2200 3150 2200 $true | Out-Null
    Add-TextRow $access $form.Name 'Vaccine' 'VaccineName' 7600 9200 3150 3000 $true | Out-Null
    Add-TextRow $access $form.Name 'Center' 'CenterName' 500 2200 3700 4200 $true | Out-Null
    Add-TextRow $access $form.Name 'Dose No.' 'DoseNo' 7600 9200 3700 1400 $true | Out-Null
    Add-TextRow $access $form.Name 'Appointment Date' 'AppointmentDate' 500 2200 4250 2200 $true | Out-Null
    Add-TextRow $access $form.Name 'Status' 'AppointmentStatus' 7600 9200 4250 2200 $true | Out-Null
    Add-TextRow $access $form.Name 'Actual Vaccination' 'ActualVaccinationDate' 500 2200 4800 2200 $true | Out-Null
    Add-TextRow $access $form.Name 'Vaccinated By' 'VaccinatedByName' 7600 9200 4800 2200 $true | Out-Null
    Add-TextRow $access $form.Name 'Batch Number' 'BatchNumber' 500 2200 5350 2200 $true | Out-Null
    Add-TextRow $access $form.Name 'Cancellation Reason' 'CancellationReason' 7600 9200 5350 3000 $true | Out-Null
    Add-TextRow $access $form.Name 'Notes' 'Notes' 500 2200 5900 10000 $true | Out-Null

    Write-FormProcedures $form $events
    Save-AndRenameForm $access $form '(Public) Appointment Tracker'
    Log 'FORM READY (Public) Appointment Tracker'
}

function Build-StockMonitorForm($access) {
    $events = [System.Collections.ArrayList]::new()
    $form = $access.CreateForm()
    $form.HasModule = $true
    $form.RecordSource = 'qry_ExpiringOrLowStockBatches'
    $form.Caption = '(Public) Vaccine_Stock_Monitor'
    $form.Width = 15000
    $form.NavigationButtons = $true
    $form.RecordSelectors = $false
    $form.DividingLines = $false
    $form.AllowEdits = $false
    $form.AllowAdditions = $false
    $form.AllowDeletions = $false
    $form.AutoCenter = $true
    $form.Section($acDetail).Height = 6500
    $form.Section($acDetail).BackColor = $colors.White

    Add-BannerBackground $access $form.Name 15000 2200 | Out-Null
    Add-TitleLabel $access $form.Name 'Vaccine Stock and Expiry Monitor' 420 260 10000 18 | Out-Null
    Add-InfoLabel $access $form.Name 'Bound to query: qry_ExpiringOrLowStockBatches' 430 760 7000 260 10 | Out-Null
    Add-StandardNavBar $access $form.Name $events

    Add-TextRow $access $form.Name 'Batch ID' 'BatchID' 500 2200 2600 1800 $true $true | Out-Null
    Add-TextRow $access $form.Name 'Batch Number' 'BatchNumber' 7600 9200 2600 2600 $true | Out-Null
    Add-TextRow $access $form.Name 'Vaccine' 'VaccineName' 500 2200 3150 3200 $true | Out-Null
    Add-TextRow $access $form.Name 'Expiry Date' 'ExpiryDate' 7600 9200 3150 2200 $true | Out-Null
    Add-TextRow $access $form.Name 'Current Stock' 'CurrentStock' 500 2200 3700 1800 $true | Out-Null
    Add-TextRow $access $form.Name 'Stock Flag' 'StockFlag' 7600 9200 3700 2200 $true | Out-Null
    Add-TextRow $access $form.Name 'Supplier' 'Supplier' 500 2200 4250 3200 $true | Out-Null
    Add-TextRow $access $form.Name 'Expiry Flag' 'ExpiryFlag' 7600 9200 4250 2200 $true | Out-Null
    Add-InfoLabel $access $form.Name 'This form is read-only and follows the saved stock alert query exactly.' 500 5200 9800 260 10 | Out-Null

    Write-FormProcedures $form $events
    Save-AndRenameForm $access $form '(Public) Vaccine_Stock_Monitor'
    Log 'FORM READY (Public) Vaccine_Stock_Monitor'
}

function Build-SecondDoseReminderForm($access) {
    $events = [System.Collections.ArrayList]::new()
    $form = $access.CreateForm()
    $form.HasModule = $true
    $form.RecordSource = 'qry_SecondDoseReminder'
    $form.Caption = '(Public) Second Dose Reminder'
    $form.Width = 15000
    $form.NavigationButtons = $true
    $form.RecordSelectors = $false
    $form.DividingLines = $false
    $form.AllowEdits = $false
    $form.AllowAdditions = $false
    $form.AllowDeletions = $false
    $form.AutoCenter = $true
    $form.Section($acDetail).Height = 6200
    $form.Section($acDetail).BackColor = $colors.White

    Add-BannerBackground $access $form.Name 15000 2200 | Out-Null
    Add-TitleLabel $access $form.Name 'Second Dose Reminder List' 420 260 10000 18 | Out-Null
    Add-InfoLabel $access $form.Name 'Bound to query: qry_SecondDoseReminder' 430 760 6000 260 10 | Out-Null
    Add-StandardNavBar $access $form.Name $events

    Add-TextRow $access $form.Name 'Citizen ID' 'CitizenID' 500 2200 2600 1800 $true $true | Out-Null
    Add-TextRow $access $form.Name 'Citizen Name' 'CitizenName' 7600 9200 2600 3000 $true | Out-Null
    Add-TextRow $access $form.Name 'Vaccine' 'VaccineName' 500 2200 3150 3200 $true | Out-Null
    Add-TextRow $access $form.Name 'Days Overdue' 'DaysOverdue' 7600 9200 3150 1800 $true $true | Out-Null
    Add-TextRow $access $form.Name 'First Dose Date' 'FirstDoseDate' 500 2200 3700 2200 $true | Out-Null
    Add-TextRow $access $form.Name 'Recommended Date' 'RecommendedSecondDoseDate' 7600 9200 3700 2200 $true | Out-Null
    Add-InfoLabel $access $form.Name 'Use this list for follow-up calls, SMS messages, or email reminders.' 500 4650 9000 260 10 | Out-Null

    Write-FormProcedures $form $events
    Save-AndRenameForm $access $form '(Public) Second Dose Reminder'
    Log 'FORM READY (Public) Second Dose Reminder'
}

function Build-AnalyticsDashboard($access) {
    $events = [System.Collections.ArrayList]::new()
    $form = $access.CreateForm()
    $form.HasModule = $true
    $form.Caption = '(Public) Analytics Dashboard'
    $form.Width = 15000
    $form.NavigationButtons = $false
    $form.RecordSelectors = $false
    $form.DividingLines = $false
    $form.AllowEdits = $false
    $form.AllowAdditions = $false
    $form.AllowDeletions = $false
    $form.AutoCenter = $true
    $form.Section($acDetail).Height = 8400
    $form.Section($acDetail).BackColor = $colors.White

    Add-BannerBackground $access $form.Name 15000 2200 | Out-Null
    Add-TitleLabel $access $form.Name 'Analytics Dashboard' 420 260 10000 18 | Out-Null
    Add-InfoLabel $access $form.Name 'Metrics are live counts over tables and saved statistical queries.' 430 760 9000 260 10 | Out-Null
    Add-StandardNavBar $access $form.Name $events

    Add-MetricBox $access $form.Name 'Total Citizens' '=DCount("*","Citizens")' 700 2700 | Out-Null
    Add-MetricBox $access $form.Name 'Appointments' '=DCount("*","Appointments")' 3400 2700 | Out-Null
    Add-MetricBox $access $form.Name 'Vaccination Records' '=DCount("*","VaccinationRecords")' 6100 2700 | Out-Null
    Add-MetricBox $access $form.Name 'Follow-up Open Cases' '=DCount("*","qry_FollowUpOpenCases")' 8800 2700 | Out-Null
    Add-MetricBox $access $form.Name 'Overdue Second Doses' '=DCount("*","qry_SecondDoseReminder")' 700 4200 | Out-Null
    Add-MetricBox $access $form.Name 'Side Effect Reports' '=DCount("*","SideEffects")' 3400 4200 | Out-Null
    Add-MetricBox $access $form.Name 'Low Stock / Expiring' '=DCount("*","qry_ExpiringOrLowStockBatches")' 6100 4200 | Out-Null
    Add-MetricBox $access $form.Name 'Active Centers' '=DCount("*","VaccinationCenters","IsActive=True")' 8800 4200 | Out-Null

    Add-InfoLabel $access $form.Name 'Open statistical queries:' 700 5700 3000 260 10 $true | Out-Null
    Add-QueryOpenButton $access $form.Name $events 'btnAnaAge' 'Age Group Dist.' 'qry_AgeGroupDistribution' 700 6050 1800 | Out-Null
    Add-QueryOpenButton $access $form.Name $events 'btnAnaAgeSide' 'Age vs Side Effect' 'qry_AgeGroupSideEffectRate' 2800 6050 2100 | Out-Null
    Add-QueryOpenButton $access $form.Name $events 'btnAnaBrand' 'Brand Reaction' 'qry_VaccineBrandReactionComparison' 5250 6050 2100 | Out-Null
    Add-QueryOpenButton $access $form.Name $events 'btnAnaSide' 'Severity Mix' 'qry_SideEffectByBrandSeverity' 7700 6050 1800 | Out-Null
    Add-QueryOpenButton $access $form.Name $events 'btnAnaSlot' 'Slot Utilization' 'qry_SlotUtilizationByCenter' 9700 6050 1900 | Out-Null
    Add-QueryOpenButton $access $form.Name $events 'btnAnaFollow' 'Open Follow-ups' 'qry_FollowUpOpenCases' 11800 6050 1800 | Out-Null

    Add-InfoLabel $access $form.Name 'Operational queries:' 700 6950 2600 260 10 $true | Out-Null
    Add-QueryOpenButton $access $form.Name $events 'btnAnaDistrict' 'District Rate' 'qry_DistrictVaccinationRate' 700 7300 1800 | Out-Null
    Add-QueryOpenButton $access $form.Name $events 'btnAnaInventory' 'Inventory Usage' 'qry_InventoryUsageAndWastage' 2800 7300 2100 | Out-Null
    Add-QueryOpenButton $access $form.Name $events 'btnAnaObserve' 'Immediate Reactions' 'qry_ObservationImmediateReactions' 5250 7300 2300 | Out-Null
    Add-QueryOpenButton $access $form.Name $events 'btnAnaScreen' 'Risk Flags' 'qry_ScreeningRiskFlags' 7900 7300 1800 | Out-Null

    Write-FormProcedures $form $events
    Save-AndRenameForm $access $form '(Public) Analytics Dashboard'
    Log 'FORM READY (Public) Analytics Dashboard'
}

function Build-CitizenEntryForm($access) {
    $events = [System.Collections.ArrayList]::new()
    $form = $access.CreateForm()
    $form.HasModule = $true
    $form.RecordSource = 'Citizens'
    $form.Caption = 'Citizen Entry Form'
    $form.Width = 15000
    $form.NavigationButtons = $true
    $form.RecordSelectors = $false
    $form.DividingLines = $false
    $form.AllowEdits = $true
    $form.AllowAdditions = $true
    $form.AllowDeletions = $false
    $form.AutoCenter = $true
    $form.Section($acDetail).Height = 9600
    $form.Section($acDetail).BackColor = $colors.White

    Add-BannerBackground $access $form.Name 15000 2200 | Out-Null
    Add-TitleLabel $access $form.Name 'Citizen Entry Form' 420 260 10000 18 | Out-Null
    Add-InfoLabel $access $form.Name 'Writes directly to table: Citizens | Please use simple sample English names, not real personal names.' 430 760 12000 260 10 | Out-Null
    Add-StandardNavBar $access $form.Name $events

    Add-EventButton $access $form.Name $events 'btnCitizenNew' 'New Record' 500 2280 1500 '    DoCmd.GoToRecord , , acNewRec' 'Create a new citizen record' | Out-Null
    Add-EventButton $access $form.Name $events 'btnCitizenSave' 'Save Record' 2200 2280 1500 "    If Me.Dirty Then`r`n        Me.Dirty = False`r`n    End If`r`n    Me.Requery" 'Save the current citizen record' | Out-Null
    Add-EventButton $access $form.Name $events 'btnCitizenFullName' 'Fill Full Name' 3900 2280 1700 "    Me.CitizenName = Trim(Nz(Me.FirstName, """") & "" "" & Nz(Me.LastName, """"))" 'Combine first and last name into display name' | Out-Null

    Add-TextRow $access $form.Name 'Citizen ID' 'CitizenID' 500 2200 3200 2400 $true $true | Out-Null
    Add-TextRow $access $form.Name 'Display Name' 'CitizenName' 7600 9200 3200 3000 $false $true | Out-Null
    Add-TextRow $access $form.Name 'First Name' 'FirstName' 500 2200 3750 2400 | Out-Null
    Add-TextRow $access $form.Name 'Last Name' 'LastName' 7600 9200 3750 2600 | Out-Null
    Add-TextRow $access $form.Name 'HKID' 'HKID' 500 2200 4300 2400 | Out-Null
    $genderCombo = Add-ComboRow $access $form.Name 'Gender' 'Gender' 7600 9200 4300 2400 'Value List' 'Male;Female;Other' 1 1
    Add-TextRow $access $form.Name 'Phone' 'Phone' 500 2200 4850 2400 | Out-Null
    Add-TextRow $access $form.Name 'Date of Birth' 'DOB' 7600 9200 4850 2400 | Out-Null
    Add-TextRow $access $form.Name 'Age Group' '=IIf(DateDiff("yyyy",[DOB],Date()) - IIf(Format([DOB],"mmdd")>Format(Date(),"mmdd"),1,0)<=11,"0-11",IIf(DateDiff("yyyy",[DOB],Date()) - IIf(Format([DOB],"mmdd")>Format(Date(),"mmdd"),1,0)<=17,"12-17",IIf(DateDiff("yyyy",[DOB],Date()) - IIf(Format([DOB],"mmdd")>Format(Date(),"mmdd"),1,0)<=29,"18-29",IIf(DateDiff("yyyy",[DOB],Date()) - IIf(Format([DOB],"mmdd")>Format(Date(),"mmdd"),1,0)<=44,"30-44",IIf(DateDiff("yyyy",[DOB],Date()) - IIf(Format([DOB],"mmdd")>Format(Date(),"mmdd"),1,0)<=59,"45-59","60+")))))' 500 2200 5400 2400 $true $true | Out-Null
    $districtCombo = Add-ComboRow $access $form.Name 'District' 'District' 7600 9200 5400 2400 'Value List' 'Hong Kong Island;Kowloon;New Territories' 1 1
    Add-TextRow $access $form.Name 'Place of Birth' 'PlaceOfBirth' 500 2200 5950 2600 | Out-Null
    Add-TextRow $access $form.Name 'Email' 'Email' 500 2200 6500 5200 | Out-Null
    Add-TextRow $access $form.Name 'Address' 'Address' 500 2200 7050 5200 | Out-Null
    Add-TextRow $access $form.Name 'Occupation' 'Occupation' 500 2200 7600 2400 | Out-Null
    Add-TextRow $access $form.Name 'Allergy History' 'AllergyHistory' 500 2200 8150 5200 | Out-Null
    Add-InfoLabel $access $form.Name 'Tip: use Fill Full Name after typing first and last name, then save the record.' 500 8720 9000 260 10 | Out-Null

    Write-FormProcedures $form $events
    Save-AndRenameForm $access $form 'Citizen Entry Form'
    Log 'FORM READY Citizen Entry Form'
}

function Build-AppointmentEntryForm($access) {
    $events = [System.Collections.ArrayList]::new()
    $form = $access.CreateForm()
    $form.HasModule = $true
    $form.RecordSource = 'Appointments'
    $form.Caption = 'Appointment Entry Form'
    $form.Width = 15000
    $form.NavigationButtons = $true
    $form.RecordSelectors = $false
    $form.DividingLines = $false
    $form.AllowEdits = $true
    $form.AllowAdditions = $true
    $form.AllowDeletions = $false
    $form.AutoCenter = $true
    $form.Section($acDetail).Height = 8200
    $form.Section($acDetail).BackColor = $colors.White

    Add-BannerBackground $access $form.Name 15000 2200 | Out-Null
    Add-TitleLabel $access $form.Name 'Appointment Entry Form' 420 260 10000 18 | Out-Null
    Add-InfoLabel $access $form.Name 'Booking form only. Use this to create and manage appointment schedules before vaccination happens.' 430 760 12000 260 10 | Out-Null
    Add-StandardNavBar $access $form.Name $events

    Add-EventButton $access $form.Name $events 'btnApptNew' 'New Record' 500 2280 1500 '    DoCmd.GoToRecord , , acNewRec' 'Create a new appointment record' | Out-Null
    Add-EventButton $access $form.Name $events 'btnApptSave' 'Save Record' 2200 2280 1500 "    If Me.Dirty Then`r`n        Me.Dirty = False`r`n    End If`r`n    Me.Requery" 'Save the current appointment record' | Out-Null
    Add-QueryOpenButton $access $form.Name $events 'btnApptConflict' 'Conflict Check' 'qry_AppointmentConflictCheck' 3900 2280 1700 | Out-Null

    Add-TextRow $access $form.Name 'Appointment ID' 'AppointmentID' 500 2200 3200 2400 $true $true | Out-Null
    $statusCombo = Add-ComboRow $access $form.Name 'Status' 'AppointmentStatus' 7600 9200 3200 2400 'Value List' 'Scheduled;Cancelled;NoShow;Completed' 1 1
    $citizenCombo = Add-ComboRow $access $form.Name 'Citizen' 'CitizenID' 500 2200 3750 3400 'Table/Query' 'SELECT CitizenID, CitizenName, HKID FROM Citizens ORDER BY CitizenName;' 3 1 '0cm;4.5cm;3cm' 5000
    $vaccineCombo = Add-ComboRow $access $form.Name 'Vaccine' 'VaccineID' 7600 9200 3750 3400 'Table/Query' 'SELECT VaccineID, VaccineName FROM Vaccines ORDER BY VaccineName;' 2 1 '0cm;4.5cm' 4200
    $centerCombo = Add-ComboRow $access $form.Name 'Center' 'CenterID' 500 2200 4300 4300 'Table/Query' 'SELECT CenterID, CenterName FROM VaccinationCenters WHERE IsActive=True ORDER BY CenterName;' 2 1 '0cm;6cm' 6200
    $slotCombo = Add-ComboRow $access $form.Name 'Slot' 'SlotID' 7600 9200 4300 4800 'Table/Query' 'SELECT AppointmentSlots.SlotID, Format([AppointmentSlots].[SlotDate],''yyyy-mm-dd'') & '' '' & Format([AppointmentSlots].[SlotTime],''hh:nn AM/PM'') & '' | '' & [VaccinationCenters].[CenterName] & '' | '' & [Vaccines].[VaccineName] FROM (AppointmentSlots INNER JOIN VaccinationCenters ON AppointmentSlots.CenterID = VaccinationCenters.CenterID) INNER JOIN Vaccines ON AppointmentSlots.VaccineID = Vaccines.VaccineID ORDER BY AppointmentSlots.SlotDate, AppointmentSlots.SlotTime;' 2 1 '0cm;9cm' 9500
    $dateBox = Add-TextRow $access $form.Name 'Appointment Date' 'AppointmentDate' 500 2200 4850 2400
    $doseCombo = Add-ComboRow $access $form.Name 'Dose No.' 'DoseNo' 7600 9200 4850 1800 'Value List' '1;2;3' 1 1
    Add-TextRow $access $form.Name 'Notes' 'Notes' 500 2200 5400 10000 | Out-Null
    Add-TextRow $access $form.Name 'Cancellation Reason' 'CancellationReason' 500 2200 5950 10000 | Out-Null
    Add-InfoLabel $access $form.Name 'Tip: this form only handles booking status. Vaccination outcomes should be entered in Vaccination Record Entry Form.' 500 6600 12000 260 10 | Out-Null

    $statusCombo.DefaultValue = """Scheduled"""
    $doseCombo.DefaultValue = '1'
    $dateBox.DefaultValue = '=Date()'

    Write-FormProcedures $form $events
    Save-AndRenameForm $access $form 'Appointment Entry Form'
    Log 'FORM READY Appointment Entry Form'
}

function Build-VaccinationRecordEntryForm($access) {
    $events = [System.Collections.ArrayList]::new()
    $form = $access.CreateForm()
    $form.HasModule = $true
    $form.RecordSource = 'VaccinationRecords'
    $form.Caption = 'Vaccination Record Entry Form'
    $form.Width = 15000
    $form.NavigationButtons = $true
    $form.RecordSelectors = $false
    $form.DividingLines = $false
    $form.AllowEdits = $true
    $form.AllowAdditions = $true
    $form.AllowDeletions = $false
    $form.AutoCenter = $true
    $form.Section($acDetail).Height = 8400
    $form.Section($acDetail).BackColor = $colors.White

    Add-BannerBackground $access $form.Name 15000 2200 | Out-Null
    Add-TitleLabel $access $form.Name 'Vaccination Record Entry Form' 420 260 11000 18 | Out-Null
    Add-InfoLabel $access $form.Name 'Use this form after a booked citizen has actually received a dose.' 430 760 12000 260 10 | Out-Null
    Add-StandardNavBar $access $form.Name $events

    Add-EventButton $access $form.Name $events 'btnRecordNew' 'New Record' 500 2280 1500 '    DoCmd.GoToRecord , , acNewRec' '' | Out-Null
    Add-EventButton $access $form.Name $events 'btnRecordSave' 'Save Record' 2200 2280 1500 "    If Me.Dirty Then`r`n        Me.Dirty = False`r`n    End If`r`n    Me.Requery" '' | Out-Null

    Add-TextRow $access $form.Name 'Record ID' 'RecordID' 500 2200 3200 2400 $true $true | Out-Null
    Add-ComboRow $access $form.Name 'Appointment' 'AppointmentID' 7600 9200 3200 4000 'Table/Query' 'SELECT AppointmentID, CitizenID, AppointmentDate FROM Appointments ORDER BY AppointmentDate DESC;' 3 1 '0cm;2cm;4cm' 6500 | Out-Null
    Add-ComboRow $access $form.Name 'Citizen' 'CitizenID' 500 2200 3750 3400 'Table/Query' 'SELECT CitizenID, CitizenName FROM Citizens ORDER BY CitizenName;' 2 1 '0cm;4.5cm' 4200 | Out-Null
    Add-ComboRow $access $form.Name 'Vaccine' 'VaccineID' 7600 9200 3750 3400 'Table/Query' 'SELECT VaccineID, VaccineName FROM Vaccines ORDER BY VaccineName;' 2 1 | Out-Null
    Add-ComboRow $access $form.Name 'Batch' 'BatchID' 500 2200 4300 3400 'Table/Query' 'SELECT BatchID, BatchNumber FROM VaccineBatches ORDER BY BatchNumber;' 2 1 | Out-Null
    $doseCombo = Add-ComboRow $access $form.Name 'Dose No.' 'DoseNo' 7600 9200 4300 1800 'Value List' '1;2;3' 1 1
    $dateBox = Add-TextRow $access $form.Name 'Vaccination Date' 'VaccinationDate' 500 2200 4850 2400
    Add-ComboRow $access $form.Name 'Staff User' 'StaffUserID' 7600 9200 4850 3400 'Table/Query' 'SELECT UserID, Username FROM Users WHERE [Role] IN (''Admin'',''MedicalStaff'') ORDER BY Username;' 2 1 | Out-Null
    Add-ComboRow $access $form.Name 'Outcome Status' 'OutcomeStatus' 500 2200 5400 3400 'Value List' 'Completed;Completed with advice;Observed;Referred' 1 1 | Out-Null

    $doseCombo.DefaultValue = '1'
    $dateBox.DefaultValue = '=Now()'

    Write-FormProcedures $form $events
    Save-AndRenameForm $access $form 'Vaccination Record Entry Form'
    Log 'FORM READY Vaccination Record Entry Form'
}

function Build-ScreeningEntryForm($access) {
    $events = [System.Collections.ArrayList]::new()
    $form = $access.CreateForm()
    $form.HasModule = $true
    $form.RecordSource = 'PreVaccinationScreening'
    $form.Caption = 'Screening Entry Form'
    $form.Width = 15000
    $form.NavigationButtons = $true
    $form.RecordSelectors = $false
    $form.DividingLines = $false
    $form.AllowEdits = $true
    $form.AllowAdditions = $true
    $form.AllowDeletions = $false
    $form.AutoCenter = $true
    $form.Section($acDetail).Height = 8400
    $form.Section($acDetail).BackColor = $colors.White

    Add-BannerBackground $access $form.Name 15000 2200 | Out-Null
    Add-TitleLabel $access $form.Name 'Screening Entry Form' 420 260 11000 18 | Out-Null
    Add-InfoLabel $access $form.Name 'Capture allergy checks, fever checks, and nurse screening notes.' 430 760 12000 260 10 | Out-Null
    Add-StandardNavBar $access $form.Name $events

    Add-EventButton $access $form.Name $events 'btnScreenNew' 'New Record' 500 2280 1500 '    DoCmd.GoToRecord , , acNewRec' '' | Out-Null
    Add-EventButton $access $form.Name $events 'btnScreenSave' 'Save Record' 2200 2280 1500 "    If Me.Dirty Then`r`n        Me.Dirty = False`r`n    End If`r`n    Me.Requery" '' | Out-Null

    Add-TextRow $access $form.Name 'Screening ID' 'ScreeningID' 500 2200 3200 2400 $true $true | Out-Null
    Add-ComboRow $access $form.Name 'Appointment' 'AppointmentID' 7600 9200 3200 4000 'Table/Query' 'SELECT AppointmentID, CitizenID, AppointmentDate FROM Appointments ORDER BY AppointmentDate DESC;' 3 1 '0cm;2cm;4cm' 6500 | Out-Null
    Add-ComboRow $access $form.Name 'Citizen' 'CitizenID' 500 2200 3750 3400 'Table/Query' 'SELECT CitizenID, CitizenName FROM Citizens ORDER BY CitizenName;' 2 1 '0cm;4.5cm' 4200 | Out-Null
    $screenDate = Add-TextRow $access $form.Name 'Screening Date' 'ScreeningDate' 7600 9200 3750 2400
    Add-CheckRow $access $form.Name 'Risk Flag' 'RiskFlag' 500 2200 4300 | Out-Null
    Add-TextRow $access $form.Name 'Temperature' 'Temperature' 7600 9200 4300 1800 | Out-Null
    Add-CheckRow $access $form.Name 'Allergy Confirmed' 'AllergyConfirmed' 500 2200 4850 | Out-Null
    Add-ComboRow $access $form.Name 'Screened By' 'ScreenedBy' 7600 9200 4850 3400 'Table/Query' 'SELECT UserID, Username FROM Users WHERE [Role] IN (''Admin'',''MedicalStaff'') ORDER BY Username;' 2 1 | Out-Null
    Add-TextRow $access $form.Name 'Contraindication Note' 'ContraindicationNote' 500 2200 5400 10000 | Out-Null

    $screenDate.DefaultValue = '=Now()'

    Write-FormProcedures $form $events
    Save-AndRenameForm $access $form 'Screening Entry Form'
    Log 'FORM READY Screening Entry Form'
}

function Build-ConsentEntryForm($access) {
    $events = [System.Collections.ArrayList]::new()
    $form = $access.CreateForm()
    $form.HasModule = $true
    $form.RecordSource = 'ConsentRecords'
    $form.Caption = 'Consent Entry Form'
    $form.Width = 15000
    $form.NavigationButtons = $true
    $form.RecordSelectors = $false
    $form.DividingLines = $false
    $form.AllowEdits = $true
    $form.AllowAdditions = $true
    $form.AllowDeletions = $false
    $form.AutoCenter = $true
    $form.Section($acDetail).Height = 7800
    $form.Section($acDetail).BackColor = $colors.White

    Add-BannerBackground $access $form.Name 15000 2200 | Out-Null
    Add-TitleLabel $access $form.Name 'Consent Entry Form' 420 260 11000 18 | Out-Null
    Add-InfoLabel $access $form.Name 'Capture whether consent is granted, declined, or still pending.' 430 760 12000 260 10 | Out-Null
    Add-StandardNavBar $access $form.Name $events

    Add-EventButton $access $form.Name $events 'btnConsentNew' 'New Record' 500 2280 1500 '    DoCmd.GoToRecord , , acNewRec' '' | Out-Null
    Add-EventButton $access $form.Name $events 'btnConsentSave' 'Save Record' 2200 2280 1500 "    If Me.Dirty Then`r`n        Me.Dirty = False`r`n    End If`r`n    Me.Requery" '' | Out-Null

    Add-TextRow $access $form.Name 'Consent ID' 'ConsentID' 500 2200 3200 2400 $true $true | Out-Null
    Add-ComboRow $access $form.Name 'Appointment' 'AppointmentID' 7600 9200 3200 4000 'Table/Query' 'SELECT AppointmentID, CitizenID, AppointmentDate FROM Appointments ORDER BY AppointmentDate DESC;' 3 1 | Out-Null
    Add-ComboRow $access $form.Name 'Citizen' 'CitizenID' 500 2200 3750 3400 'Table/Query' 'SELECT CitizenID, CitizenName FROM Citizens ORDER BY CitizenName;' 2 1 | Out-Null
    Add-ComboRow $access $form.Name 'Consent Status' 'ConsentStatus' 7600 9200 3750 2400 'Value List' 'Granted;Declined;Pending' 1 1 | Out-Null
    $consentDate = Add-TextRow $access $form.Name 'Consent Date' 'ConsentDate' 500 2200 4300 2400
    Add-ComboRow $access $form.Name 'Witnessed By' 'WitnessedBy' 7600 9200 4300 3400 'Table/Query' 'SELECT UserID, Username FROM Users WHERE [Role] IN (''Admin'',''MedicalStaff'') ORDER BY Username;' 2 1 | Out-Null
    Add-TextRow $access $form.Name 'Consent Note' 'ConsentNote' 500 2200 4850 10000 | Out-Null

    $consentDate.DefaultValue = '=Now()'

    Write-FormProcedures $form $events
    Save-AndRenameForm $access $form 'Consent Entry Form'
    Log 'FORM READY Consent Entry Form'
}

function Build-ObservationEntryForm($access) {
    $events = [System.Collections.ArrayList]::new()
    $form = $access.CreateForm()
    $form.HasModule = $true
    $form.RecordSource = 'ObservationRecords'
    $form.Caption = 'Observation Entry Form'
    $form.Width = 15000
    $form.NavigationButtons = $true
    $form.RecordSelectors = $false
    $form.DividingLines = $false
    $form.AllowEdits = $true
    $form.AllowAdditions = $true
    $form.AllowDeletions = $false
    $form.AutoCenter = $true
    $form.Section($acDetail).Height = 8200
    $form.Section($acDetail).BackColor = $colors.White

    Add-BannerBackground $access $form.Name 15000 2200 | Out-Null
    Add-TitleLabel $access $form.Name 'Observation Entry Form' 420 260 11000 18 | Out-Null
    Add-InfoLabel $access $form.Name 'Record the 15-30 minute post-vaccination observation period.' 430 760 12000 260 10 | Out-Null
    Add-StandardNavBar $access $form.Name $events

    Add-EventButton $access $form.Name $events 'btnObsNew' 'New Record' 500 2280 1500 '    DoCmd.GoToRecord , , acNewRec' '' | Out-Null
    Add-EventButton $access $form.Name $events 'btnObsSave' 'Save Record' 2200 2280 1500 "    If Me.Dirty Then`r`n        Me.Dirty = False`r`n    End If`r`n    Me.Requery" '' | Out-Null

    Add-TextRow $access $form.Name 'Observation ID' 'ObservationID' 500 2200 3200 2400 $true $true | Out-Null
    Add-ComboRow $access $form.Name 'Record' 'RecordID' 7600 9200 3200 4000 'Table/Query' 'SELECT RecordID, CitizenID, VaccinationDate FROM VaccinationRecords ORDER BY VaccinationDate DESC;' 3 1 | Out-Null
    Add-ComboRow $access $form.Name 'Citizen' 'CitizenID' 500 2200 3750 3400 'Table/Query' 'SELECT CitizenID, CitizenName FROM Citizens ORDER BY CitizenName;' 2 1 | Out-Null
    $startBox = Add-TextRow $access $form.Name 'Observation Start' 'ObservationStart' 7600 9200 3750 2400
    Add-TextRow $access $form.Name 'Observation End' 'ObservationEnd' 500 2200 4300 2400 | Out-Null
    Add-CheckRow $access $form.Name 'Immediate Reaction' 'ImmediateReaction' 7600 9200 4300 | Out-Null
    Add-ComboRow $access $form.Name 'Observed By' 'ObservedBy' 500 2200 4850 3400 'Table/Query' 'SELECT UserID, Username FROM Users WHERE [Role] IN (''Admin'',''MedicalStaff'') ORDER BY Username;' 2 1 | Out-Null
    Add-ComboRow $access $form.Name 'Observation Outcome' 'ObservationOutcome' 7600 9200 4850 3400 'Value List' 'Stable discharge;Extended observation;Referred to clinic' 1 1 | Out-Null

    $startBox.DefaultValue = '=Now()'

    Write-FormProcedures $form $events
    Save-AndRenameForm $access $form 'Observation Entry Form'
    Log 'FORM READY Observation Entry Form'
}

function Build-AdverseFollowupForm($access) {
    $events = [System.Collections.ArrayList]::new()
    $form = $access.CreateForm()
    $form.HasModule = $true
    $form.RecordSource = 'AdverseEffectFollowUps'
    $form.Caption = 'Adverse Follow-up Form'
    $form.Width = 15000
    $form.NavigationButtons = $true
    $form.RecordSelectors = $false
    $form.DividingLines = $false
    $form.AllowEdits = $true
    $form.AllowAdditions = $true
    $form.AllowDeletions = $false
    $form.AutoCenter = $true
    $form.Section($acDetail).Height = 8200
    $form.Section($acDetail).BackColor = $colors.White

    Add-BannerBackground $access $form.Name 15000 2200 | Out-Null
    Add-TitleLabel $access $form.Name 'Adverse Follow-up Form' 420 260 11000 18 | Out-Null
    Add-InfoLabel $access $form.Name 'Track safety follow-up actions after a side effect report is received.' 430 760 12000 260 10 | Out-Null
    Add-StandardNavBar $access $form.Name $events

    Add-EventButton $access $form.Name $events 'btnFollowNew' 'New Record' 500 2280 1500 '    DoCmd.GoToRecord , , acNewRec' '' | Out-Null
    Add-EventButton $access $form.Name $events 'btnFollowSave' 'Save Record' 2200 2280 1500 "    If Me.Dirty Then`r`n        Me.Dirty = False`r`n    End If`r`n    Me.Requery" '' | Out-Null

    Add-TextRow $access $form.Name 'Follow-up ID' 'FollowUpID' 500 2200 3200 2400 $true $true | Out-Null
    Add-ComboRow $access $form.Name 'Report' 'ReportID' 7600 9200 3200 4200 'Table/Query' 'SELECT ReportID, CitizenID, SeverityLevel FROM SideEffects ORDER BY ReportDate DESC;' 3 1 | Out-Null
    $followDate = Add-TextRow $access $form.Name 'Follow-up Date' 'FollowUpDate' 500 2200 3750 2400
    Add-ComboRow $access $form.Name 'Staff User' 'StaffUserID' 7600 9200 3750 3400 'Table/Query' 'SELECT UserID, Username FROM Users WHERE [Role] IN (''Admin'',''MedicalStaff'') ORDER BY Username;' 2 1 | Out-Null
    Add-TextRow $access $form.Name 'Action Taken' 'ActionTaken' 500 2200 4300 10000 | Out-Null
    Add-ComboRow $access $form.Name 'Follow-up Outcome' 'FollowUpOutcome' 500 2200 4850 3400 'Value List' 'Recovered;Monitoring;Open;Needs Review' 1 1 | Out-Null
    Add-CheckRow $access $form.Name 'Escalation Flag' 'EscalationFlag' 7600 9200 4850 | Out-Null

    $followDate.DefaultValue = '=Date()'

    Write-FormProcedures $form $events
    Save-AndRenameForm $access $form 'Adverse Follow-up Form'
    Log 'FORM READY Adverse Follow-up Form'
}

function Build-InventoryTransactionForm($access) {
    $events = [System.Collections.ArrayList]::new()
    $form = $access.CreateForm()
    $form.HasModule = $true
    $form.RecordSource = 'InventoryTransactions'
    $form.Caption = 'Inventory Transaction Form'
    $form.Width = 15000
    $form.NavigationButtons = $true
    $form.RecordSelectors = $false
    $form.DividingLines = $false
    $form.AllowEdits = $true
    $form.AllowAdditions = $true
    $form.AllowDeletions = $false
    $form.AutoCenter = $true
    $form.Section($acDetail).Height = 8200
    $form.Section($acDetail).BackColor = $colors.White

    Add-BannerBackground $access $form.Name 15000 2200 | Out-Null
    Add-TitleLabel $access $form.Name 'Inventory Transaction Form' 420 260 11000 18 | Out-Null
    Add-InfoLabel $access $form.Name 'Track allocation, usage, wastage, and adjustment movements at each center.' 430 760 12000 260 10 | Out-Null
    Add-StandardNavBar $access $form.Name $events

    Add-EventButton $access $form.Name $events 'btnInvNew' 'New Record' 500 2280 1500 '    DoCmd.GoToRecord , , acNewRec' '' | Out-Null
    Add-EventButton $access $form.Name $events 'btnInvSave' 'Save Record' 2200 2280 1500 "    If Me.Dirty Then`r`n        Me.Dirty = False`r`n    End If`r`n    Me.Requery" '' | Out-Null

    Add-TextRow $access $form.Name 'Transaction ID' 'TransactionID' 500 2200 3200 2400 $true $true | Out-Null
    Add-ComboRow $access $form.Name 'Batch' 'BatchID' 7600 9200 3200 3400 'Table/Query' 'SELECT BatchID, BatchNumber FROM VaccineBatches ORDER BY BatchNumber;' 2 1 | Out-Null
    Add-ComboRow $access $form.Name 'Center' 'CenterID' 500 2200 3750 3400 'Table/Query' 'SELECT CenterID, CenterName FROM VaccinationCenters ORDER BY CenterName;' 2 1 | Out-Null
    Add-ComboRow $access $form.Name 'Type' 'TransactionType' 7600 9200 3750 2400 'Value List' 'Allocation;Used;Wastage;Adjustment' 1 1 | Out-Null
    Add-TextRow $access $form.Name 'Quantity' 'Quantity' 500 2200 4300 1800 | Out-Null
    $txnDate = Add-TextRow $access $form.Name 'Transaction Date' 'TransactionDate' 7600 9200 4300 2400
    Add-ComboRow $access $form.Name 'Related Record' 'RelatedRecordID' 500 2200 4850 3400 'Table/Query' 'SELECT RecordID, CitizenID FROM VaccinationRecords ORDER BY RecordID DESC;' 2 1 | Out-Null
    Add-ComboRow $access $form.Name 'Processed By' 'ProcessedBy' 7600 9200 4850 3400 'Table/Query' 'SELECT UserID, Username FROM Users WHERE [Role] IN (''Admin'',''MedicalStaff'') ORDER BY Username;' 2 1 | Out-Null

    $txnDate.DefaultValue = '=Now()'

    Write-FormProcedures $form $events
    Save-AndRenameForm $access $form 'Inventory Transaction Form'
    Log 'FORM READY Inventory Transaction Form'
}

$before = @(Get-Process MSACCESS -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
$createdAccess = $false

try {
    Log 'OPEN ACCESS'
    try {
        $access = New-Object -ComObject Access.Application
        $createdAccess = $true
    } catch {
        $access = [Runtime.InteropServices.Marshal]::GetActiveObject('Access.Application')
    }
    if ($createdAccess) {
        $access.Visible = $false
    }
    try { $access.CloseCurrentDatabase() } catch {}
    $access.OpenCurrentDatabase($databasePath, $true)
    Log 'DB OPENED'
    Close-OpenForms $access

    try {
        $db = $access.CurrentDb()
    } catch {
        $db = $null
    }
    if ($db) {
        Ensure-CitizenFields $db
        Ensure-UiQueries $db
        try { $db.Close() } catch {}
        $db = $null
    } else {
        Log 'DAO context unavailable in Access session; using existing schema and saved queries.'
    }

    foreach ($name in @(
        '(Portal) HomePage',
        '(Citizen) Medical_Record',
        '(Public) Appointment Tracker',
        '(Public) Vaccine_Stock_Monitor',
        '(Public) Second Dose Reminder',
        '(Public) Analytics Dashboard',
        'Citizen Entry Form',
        'Appointment Entry Form',
        'Vaccination Record Entry Form',
        'Screening Entry Form',
        'Consent Entry Form',
        'Observation Entry Form',
        'Adverse Follow-up Form',
        'Inventory Transaction Form'
    )) {
        Remove-FormIfExists $access $name
    }
    Log 'OLD FORMS CLEARED'

    Build-HomePage $access
    Build-CitizenRecordForm $access
    Build-AppointmentTrackerForm $access
    Build-StockMonitorForm $access
    Build-SecondDoseReminderForm $access
    Build-AnalyticsDashboard $access
    Build-CitizenEntryForm $access
    Build-AppointmentEntryForm $access
    Build-VaccinationRecordEntryForm $access
    Build-ScreeningEntryForm $access
    Build-ConsentEntryForm $access
    Build-ObservationEntryForm $access
    Build-AdverseFollowupForm $access
    Build-InventoryTransactionForm $access
    Log 'ALL FORMS BUILT'
} catch {
    Log ('ERROR: ' + $_.Exception.Message)
    throw
} finally {
    try {
        if ($access -and $createdAccess) {
            $access.CloseCurrentDatabase()
        }
    } catch {}
    $db = $null
    $access = $null
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    Start-Sleep -Seconds 2
    $after = @(Get-Process MSACCESS -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
    foreach ($procId in $after) {
        if ($before -notcontains $procId) {
            try {
                Stop-Process -Id $procId -Force -ErrorAction Stop
            } catch {
                try { taskkill /PID $procId /F | Out-Null } catch {}
            }
        }
    }
    Log 'END'
}
