Add-Type -AssemblyName System.Drawing
$ErrorActionPreference = 'Stop'

$databasePath = '.\gp.accdb'
$logPath = '.\build_missing_v2_forms.log'

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

function Test-DatabaseExclusiveAccess([string]$path) {
    $probeEngine = $null
    $probeDb = $null
    try {
        $probeEngine = New-Object -ComObject DAO.DBEngine.120
        $probeDb = $probeEngine.OpenDatabase($path, $true, $false)
        return $true
    } catch {
        return $false
    } finally {
        if ($probeDb) {
            try { $probeDb.Close() } catch {}
            try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($probeDb) } catch {}
        }
        if ($probeEngine) {
            try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($probeEngine) } catch {}
        }
    }
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
    Add-ComboRow $access $form.Name 'Appointment' 'AppointmentID' 7600 9200 3200 4000 'Table/Query' 'SELECT AppointmentID, CitizenID, AppointmentDate FROM Appointments ORDER BY AppointmentDate DESC;' 3 1 '0cm;2cm;4cm' 6500 | Out-Null
    Add-ComboRow $access $form.Name 'Citizen' 'CitizenID' 500 2200 3750 3400 'Table/Query' 'SELECT CitizenID, CitizenName FROM Citizens ORDER BY CitizenName;' 2 1 '0cm;4.5cm' 4200 | Out-Null
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
    Add-ComboRow $access $form.Name 'Record' 'RecordID' 7600 9200 3200 4000 'Table/Query' 'SELECT RecordID, CitizenID, VaccinationDate FROM VaccinationRecords ORDER BY VaccinationDate DESC;' 3 1 '0cm;2cm;4cm' 6500 | Out-Null
    Add-ComboRow $access $form.Name 'Citizen' 'CitizenID' 500 2200 3750 3400 'Table/Query' 'SELECT CitizenID, CitizenName FROM Citizens ORDER BY CitizenName;' 2 1 '0cm;4.5cm' 4200 | Out-Null
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
    Add-ComboRow $access $form.Name 'Report' 'ReportID' 7600 9200 3200 4200 'Table/Query' 'SELECT ReportID, CitizenID, SeverityLevel FROM SideEffects ORDER BY ReportDate DESC;' 3 1 '0cm;2cm;3cm' 6500 | Out-Null
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
    if (-not (Test-DatabaseExclusiveAccess $databasePath)) {
        throw 'gp.accdb is genuinely locked by another Access session. Close all MSACCESS.EXE processes, then run this script again.'
    }

    $lockPath = '.\gp.laccdb'
    if (Test-Path $lockPath) {
        try {
            Remove-Item $lockPath -Force -ErrorAction Stop
            Log 'STALE LOCK FILE REMOVED'
        } catch {
            Log 'STALE LOCK FILE COULD NOT BE REMOVED; continuing because exclusive access probe succeeded.'
        }
    }

    Log 'OPEN ACCESS'
    $access = New-Object -ComObject Access.Application
    $createdAccess = $true
    $access.Visible = $true
    $access.UserControl = $true
    $access.OpenCurrentDatabase($databasePath, $true)
    Log 'DB OPENED'

    foreach ($name in @(
        'Consent Entry Form',
        'Observation Entry Form',
        'Adverse Follow-up Form',
        'Inventory Transaction Form'
    )) {
        Remove-FormIfExists $access $name
    }
    Log 'TARGET FORMS CLEARED'

    Build-ConsentEntryForm $access
    Build-ObservationEntryForm $access
    Build-AdverseFollowupForm $access
    Build-InventoryTransactionForm $access
    Log 'MISSING FORMS BUILT'
} catch {
    Log ('ERROR: ' + $_.Exception.Message)
    throw
} finally {
    try {
        if ($access -and $createdAccess) {
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
