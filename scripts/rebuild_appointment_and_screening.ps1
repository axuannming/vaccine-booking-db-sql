Add-Type -AssemblyName System.Drawing
$ErrorActionPreference = 'Stop'

$databasePath = '.\gp.accdb'
$logPath = '.\rebuild_appointment_and_screening.log'
$appointmentExportPath = '.\rebuild_appointment_entry_form.txt'
$screeningExportPath = '.\rebuild_screening_entry_form.txt'

$acDetail = 0
$acForm = 2
$acLabel = 100
$acCommandButton = 104
$acCheckBox = 106
$acTextBox = 109
$acComboBox = 111

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
    Highlight = Get-OleColor 255 249 220
    Border = Get-OleColor 185 201 214
}

function Remove-FormIfExists($access, [string]$name) {
    try { $access.DoCmd.DeleteObject($acForm, $name) } catch {}
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
    if ($events.Count -eq 0) { return }
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
    Add-ComboRow $access $form.Name 'Appointment' 'AppointmentID' 7600 9200 3200 4000 'Table/Query' 'SELECT Appointments.AppointmentID, Citizens.CitizenName, Appointments.AppointmentDate FROM Appointments INNER JOIN Citizens ON Appointments.CitizenID = Citizens.CitizenID ORDER BY Appointments.AppointmentDate DESC;' 3 1 '0cm;4cm;3.5cm' 7200 | Out-Null
    Add-ComboRow $access $form.Name 'Citizen' 'CitizenID' 500 2200 3750 3400 'Table/Query' 'SELECT CitizenID, CitizenName FROM Citizens ORDER BY CitizenName;' 2 1 '0cm;4.5cm' 4200 | Out-Null
    $screenDate = Add-TextRow $access $form.Name 'Screening Date' 'ScreeningDate' 7600 9200 3750 2400
    Add-CheckRow $access $form.Name 'Risk Flag' 'RiskFlag' 500 2200 4300 | Out-Null
    Add-TextRow $access $form.Name 'Temperature' 'Temperature' 7600 9200 4300 1800 | Out-Null
    Add-CheckRow $access $form.Name 'Allergy Confirmed' 'AllergyConfirmed' 500 2200 4850 | Out-Null
    Add-ComboRow $access $form.Name 'Screened By' 'ScreenedBy' 7600 9200 4850 3400 'Table/Query' 'SELECT UserID, Username FROM Users WHERE [Role] IN (''Admin'',''MedicalStaff'') ORDER BY Username;' 2 1 '0cm;4.5cm' 4200 | Out-Null
    Add-TextRow $access $form.Name 'Contraindication Note' 'ContraindicationNote' 500 2200 5400 10000 | Out-Null

    $screenDate.DefaultValue = '=Now()'

    Write-FormProcedures $form $events
    Save-AndRenameForm $access $form 'Screening Entry Form'
    Log 'FORM READY Screening Entry Form'
}

function Export-AndCheckForm($access, [string]$formName, [string]$targetPath) {
    $access.SaveAsText($acForm, $formName, $targetPath)
    $hasCodeBehind = Select-String -Path $targetPath -Pattern '^CodeBehindForm$' -Quiet
    Log ($formName + ' EXPORTED :: CodeBehindForm=' + $hasCodeBehind)
}

$before = @(Get-Process MSACCESS -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)

try {
    if (Test-Path '.\gp.laccdb') {
        Log 'WARNING: gp.laccdb exists before rebuild. Continuing only if Access can open the database.'
    }

    Log 'OPEN ACCESS'
    $access = New-Object -ComObject Access.Application
    $access.Visible = $true
    $access.UserControl = $true
    $access.OpenCurrentDatabase($databasePath, $true)
    Log 'DB OPENED'

    Remove-FormIfExists $access 'Appointment Entry Form'
    Remove-FormIfExists $access 'Screening Entry Form'
    Log 'TARGET FORMS CLEARED'

    Build-AppointmentEntryForm $access
    Build-ScreeningEntryForm $access

    Export-AndCheckForm $access 'Appointment Entry Form' $appointmentExportPath
    Export-AndCheckForm $access 'Screening Entry Form' $screeningExportPath
    Log 'DONE'
} catch {
    Log ('ERROR: ' + $_.Exception.Message)
    throw
} finally {
    try {
        if ($access) {
            $access.CloseCurrentDatabase()
            $access.Quit(2)
        }
    } catch {}
    if ($access) {
        try { [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($access) } catch {}
    }
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
