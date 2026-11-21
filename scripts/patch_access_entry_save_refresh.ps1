$ErrorActionPreference = 'Stop'

$databasePath = '.\gp.accdb'
$workspaceRoot = '.\access_refresh_patch'
$exportDir = Join-Path $workspaceRoot 'exports'
$logPath = Join-Path $workspaceRoot 'patch.log'
$backupPath = Join-Path $workspaceRoot 'gp_before_entry_refresh_patch.accdb'

$acForm = 2

New-Item -ItemType Directory -Force -Path $workspaceRoot, $exportDir | Out-Null
Set-Content -Path $logPath -Value @('START ' + (Get-Date).ToString('s'))

function Log([string]$message) {
    Add-Content -Path $logPath -Value ((Get-Date).ToString('s') + ' ' + $message)
}

function Remove-ExistingHelperBlocks([string]$text) {
    $patterns = @(
        '(?ms)^Private Sub CodexRefreshAfterSave\(ByVal tableName As String\)(?:\r?\n).*?^End Sub(?:\r?\n)*',
        '(?ms)^Private Sub CodexRefreshOpenForm\(ByVal formName As String\)(?:\r?\n).*?^End Sub(?:\r?\n)*',
        '(?ms)^Private Sub CodexRefreshOpenTable\(ByVal tableName As String\)(?:\r?\n).*?^End Sub(?:\r?\n)*'
    )

    foreach ($pattern in $patterns) {
        $text = [regex]::Replace($text, $pattern, '')
    }
    return $text.TrimEnd() + "`r`n"
}

function Get-HelperCode() {
    return @"
Private Sub CodexRefreshAfterSave(ByVal tableName As String)
    On Error Resume Next

    CodexRefreshOpenForm "(Portal) HomePage"
    CodexRefreshOpenForm "(Citizen) Medical_Record"
    CodexRefreshOpenForm "(Public) Appointment Tracker"
    CodexRefreshOpenForm "(Public) Vaccine_Stock_Monitor"
    CodexRefreshOpenForm "(Public) Second Dose Reminder"
    CodexRefreshOpenForm "(Public) Analytics Dashboard"
    CodexRefreshOpenTable tableName

    DBEngine.Idle dbRefreshCache
    Application.RefreshDatabaseWindow

    On Error GoTo 0
End Sub

Private Sub CodexRefreshOpenForm(ByVal formName As String)
    On Error Resume Next
    If CurrentProject.AllForms(formName).IsLoaded Then
        Forms(formName).Requery
    End If
    On Error GoTo 0
End Sub

Private Sub CodexRefreshOpenTable(ByVal tableName As String)
    On Error Resume Next
    If CurrentData.AllTables(tableName).IsLoaded Then
        DoCmd.Close acTable, tableName, acSaveNo
        DoCmd.OpenTable tableName, acViewNormal, acEdit
    End If
    On Error GoTo 0
End Sub
"@
}

function Update-SaveProcedure([string]$text, [string]$procedureName, [string]$tableName) {
    $replacement = @"
Private Sub $procedureName()
    If Me.Dirty Then
        Me.Dirty = False
    End If
    Me.Requery
    CodexRefreshAfterSave "$tableName"
End Sub
"@

    $pattern = "(?ms)^Private Sub $([regex]::Escape($procedureName))\(\)(?:\r?\n).*?^End Sub"
    $updated = [regex]::Replace($text, $pattern, $replacement, 1)

    if ($updated -eq $text) {
        throw "Could not find procedure $procedureName"
    }

    return $updated
}

$forms = @(
    @{ FormName = 'Citizen Entry Form'; SaveProcedure = 'btnCitizenSave_Click'; TableName = 'Citizens' },
    @{ FormName = 'Appointment Entry Form'; SaveProcedure = 'btnApptSave_Click'; TableName = 'Appointments' },
    @{ FormName = 'Vaccination Record Entry Form'; SaveProcedure = 'btnRecordSave_Click'; TableName = 'VaccinationRecords' },
    @{ FormName = 'Screening Entry Form'; SaveProcedure = 'btnScreenSave_Click'; TableName = 'PreVaccinationScreening' },
    @{ FormName = 'Consent Entry Form'; SaveProcedure = 'btnConsentSave_Click'; TableName = 'ConsentRecords' },
    @{ FormName = 'Observation Entry Form'; SaveProcedure = 'btnObsSave_Click'; TableName = 'ObservationRecords' },
    @{ FormName = 'Adverse Follow-up Form'; SaveProcedure = 'btnFollowSave_Click'; TableName = 'AdverseEffectFollowUps' },
    @{ FormName = 'Inventory Transaction Form'; SaveProcedure = 'btnInvSave_Click'; TableName = 'InventoryTransactions' }
)

$createdAccess = $false
$attachedAccess = $false
$access = $null

try {
    try {
        Copy-Item -LiteralPath $databasePath -Destination $backupPath -Force
        Log "BACKUP $backupPath"
    } catch {
        Log ('BACKUP SKIPPED ' + $_.Exception.Message)
    }

    try {
        $access = [Runtime.InteropServices.Marshal]::GetActiveObject('Access.Application')
        $attachedAccess = $true
        Log 'ATTACHED TO RUNNING ACCESS'
    } catch {
        $access = New-Object -ComObject Access.Application
        $createdAccess = $true
        $access.Visible = $false
        Log 'CREATED NEW ACCESS INSTANCE'
    }

    if ($attachedAccess) {
        $currentDb = ''
        try { $currentDb = $access.CurrentProject.FullName } catch {}
        if ([string]::IsNullOrWhiteSpace($currentDb)) {
            throw 'Please open gp.accdb in Access first, then rerun this patch.'
        }
        if ([string]::Compare($currentDb, $databasePath, $true) -ne 0) {
            throw "The running Access window is attached to $currentDb, not $databasePath."
        }
        Log "USING OPEN DATABASE $currentDb"
    } else {
        $access.OpenCurrentDatabase($databasePath, $true)
        Log "DB OPENED $databasePath"
    }

    foreach ($form in $forms) {
        $safeName = ($form.FormName -replace '[\\/:*?""<>|()]', '_').Replace(' ', '_')
        $exportPath = Join-Path $exportDir ($safeName + '.txt')
        $wasOpen = $false

        try {
            $wasOpen = $access.CurrentProject.AllForms($form.FormName).IsLoaded
        } catch {}
        if ($wasOpen) {
            try {
                $access.DoCmd.Close($acForm, $form.FormName, 2)
                Log "CLOSED OPEN FORM $($form.FormName)"
            } catch {
                Log "CLOSE SKIPPED $($form.FormName)"
            }
        }

        Log "EXPORT $($form.FormName)"
        $access.SaveAsText($acForm, $form.FormName, $exportPath)

        $text = Get-Content -LiteralPath $exportPath -Raw -Encoding UTF8
        $text = Remove-ExistingHelperBlocks $text
        $text = Update-SaveProcedure $text $form.SaveProcedure $form.TableName
        $text = $text.TrimEnd() + "`r`n`r`n" + (Get-HelperCode)
        Set-Content -LiteralPath $exportPath -Value $text -Encoding UTF8

        Log "IMPORT $($form.FormName)"
        $access.DoCmd.DeleteObject($acForm, $form.FormName)
        $access.LoadFromText($acForm, $form.FormName, $exportPath)

        if ($wasOpen) {
            try {
                $access.DoCmd.OpenForm($form.FormName)
                Log "REOPENED FORM $($form.FormName)"
            } catch {
                Log "REOPEN FAILED $($form.FormName)"
            }
        }
    }

    try {
        $access.DoCmd.RunCommand(603)
        Log 'COMPILE OK'
    } catch {
        Log ('COMPILE SKIPPED ' + $_.Exception.Message)
    }

    foreach ($form in $forms) {
        $safeName = ($form.FormName -replace '[\\/:*?""<>|()]', '_').Replace(' ', '_')
        $verifyPath = Join-Path $exportDir ($safeName + '.verify.txt')
        $access.SaveAsText($acForm, $form.FormName, $verifyPath)
        Log "VERIFY EXPORT $($form.FormName)"
    }

    Log 'PATCH COMPLETE'
} catch {
    Log ('ERROR ' + $_.Exception.Message)
    throw
} finally {
    if ($access) {
        if ($createdAccess) {
            try { $access.CloseCurrentDatabase() } catch {}
            try { $access.Quit(2) } catch {}
        }
    }
}
