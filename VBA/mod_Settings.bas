Attribute VB_Name = "mod_Settings"
Option Explicit

'===============================================================================
' MODULE : mod_Settings
' DOMAINE / DOMAIN : Settings / Language
'
' FR
' Possede les reglages persistants, les langues par domaine et les controles du panneau Settings.
' Ne decide aucune politique de calcul planning.
'
' EN
' Owns persisted settings, per-domain languages and Settings panel controls.
' Does not decide planning calculation policy.
'
' CONTRATS / CONTRACTS : Settings_Initialize, Settings_HydrateRuntimeState, Settings_GetOwnerLanguage, Settings_ToggleInfoMessages, Settings_ToggleGlobalLanguage, Settings_ToggleGlobalActivated
' CALLBACKS EXTERNES / EXTERNAL CALLBACKS : Aucun / None
'===============================================================================


Private Const SETTINGS_SHEET As String = "SETTINGS"
Private Const SETTINGS_PREFIX As String = "SET_"

Private Const CELL_GLOBAL_LANGUAGE As String = "X2"
Private Const CELL_GLOBAL_ACTIVATED As String = "X3"
Private Const CELL_DASHBOARD_LANGUAGE As String = "X4"
Private Const CELL_GANTT_LANGUAGE As String = "X5"
Private Const CELL_SCURVE_LANGUAGE As String = "X6"
Private Const CELL_WBS_LANGUAGE As String = "X7"
Private Const CELL_EVENT_LANGUAGE As String = "X8"
Private Const CELL_INFO_ENABLED As String = "X9"
Private Const CELL_CONSTRAINTS_LANGUAGE As String = "X10"
Private Const CELL_DATE_DISPLAY_MODE As String = "X11"
Private Const CELL_FINANCE_LANGUAGE As String = "X20"
Private Const SETTINGS_STORAGE_VERSION As String = "SETTINGS_STORAGE_V2"
Private Const TUTORIAL_STORAGE_VERSION As String = "TUTORIAL_PROGRESS_V1"
Private mTutorialSession As Variant

Private Const DATE_MODE_DMY As String = "DMY"
Private Const DATE_MODE_MDY As String = "MDY"
Private Const DATE_MODE_ISO As String = "ISO"

Private Const DATE_FORMAT_DMY As String = "dd/mm/yyyy"
Private Const DATE_FORMAT_MDY As String = "mm/dd/yyyy"
Private Const DATE_FORMAT_ISO As String = "yyyy-mm-dd"
Private Const DATETIME_FORMAT_DMY As String = "dd/mm/yyyy hh:mm:ss"
Private Const DATETIME_FORMAT_MDY As String = "mm/dd/yyyy hh:mm:ss"
Private Const DATETIME_FORMAT_ISO As String = "yyyy-mm-dd hh:mm:ss"
Private Const DATE_COMPACT_FORMAT_DMY As String = "dd/mm"
Private Const DATE_COMPACT_FORMAT_MDY As String = "mm/dd"
Private Const DATE_COMPACT_FORMAT_ISO As String = "yyyy-mm-dd"

Private Const DATE_CONTROL_LABEL_NAME As String = "SET_DATE_FORMAT_VALUES"
Private Const DATE_CONTROL_BG_NAME As String = "SET_DATE_FORMAT_TRACK"
Private Const DATE_CONTROL_KNOB_NAME As String = "SET_DATE_FORMAT_KNOB"

Private Const GANTT_REFERENCE_TRACK_WIDTH As Double = 28
Private Const GANTT_THREE_POSITION_TRACK_WIDTH As Double = 46
Private Const GANTT_REFERENCE_TRACK_HEIGHT As Double = 12
Private Const GANTT_REFERENCE_KNOB_SIZE As Double = 9
Private Const GANTT_REFERENCE_LABEL_GAP As Double = 7
Private Const SETTINGS_TRACK_WIDTH As Double = 44
Private Const SETTINGS_TRACK_HEIGHT As Double = 16
Private Const SETTINGS_KNOB_SIZE As Double = 12
Private Const DATE_CONTROL_LABEL_WIDTH As Double = 104

Private Const MODULE_DASHBOARD As String = "DASHBOARD"
Private Const MODULE_GANTT As String = "GANTT"
Private Const MODULE_SCURVE As String = "SCURVE"
Private Const MODULE_WBS As String = "WBS"
Private Const MODULE_CONSTRAINTS As String = "CONSTRAINTS"
Private Const MODULE_EVENT As String = "EVENT"
Private Const MODULE_FINANCE As String = "FINANCE"

'------------------------------------------------------------------------------
' FR: Met a jour Settings Initialize dans le contexte settings and language.
' EN: Updates Settings Initialize in the settings and language context.
'------------------------------------------------------------------------------
Public Sub Settings_Initialize()

    Dim ws As Worksheet
    Dim oldEvents As Boolean
    Dim oldScreenUpdating As Boolean

    On Error GoTo SafeExit

    oldEvents = Application.EnableEvents
    oldScreenUpdating = Application.ScreenUpdating
    Application.EnableEvents = False
    Application.ScreenUpdating = False

    Set ws = Settings_EnsureSheet()
    Settings_BuildLayout ws

SafeExit:
    Application.ScreenUpdating = oldScreenUpdating
    Application.EnableEvents = oldEvents

End Sub

'------------------------------------------------------------------------------
' FR: Met a jour Settings Hydrate Runtime State dans le contexte settings and language.
' EN: Updates Settings Hydrate Runtime State in the settings and language context.
'------------------------------------------------------------------------------
Public Sub Settings_HydrateRuntimeState()

    EventHistory_SetShowInfo Settings_GetInfoEnabled()
    EventHistory_SetLanguage Settings_GetOwnerLanguage(MODULE_EVENT)
    WBS_SetLanguage Settings_GetOwnerLanguage(MODULE_WBS)
    Gantt_SetLanguage Settings_GetOwnerLanguage(MODULE_GANTT)
    SCurve_SetLanguage Settings_GetOwnerLanguage(MODULE_SCURVE)
    Dashboard_SetLanguage Settings_GetOwnerLanguage(MODULE_DASHBOARD)
    Constraints_SetLanguage Settings_GetOwnerLanguage(MODULE_CONSTRAINTS)

End Sub

' Retire only the legacy UI now owned by the Ribbon.
Public Sub Settings_RemoveLegacyMigrationCommand()
    Dim ws As Worksheet, i As Long
    Set ws = ThisWorkbook.Worksheets(SETTINGS_SHEET)
    For i = 1 To ws.Shapes.count
        If ws.Shapes(i).name = "SET_BTN_MIGRATION" Then
            ws.Shapes(i).Delete
            Exit Sub
        End If
    Next i
End Sub

' This owner alone knows the persisted Settings layout, including legacy V1.
Public Function Settings_MigrationSnapshot(ByVal source As Workbook) As Object
    Dim result As Object, ws As Worksheet, candidate As Worksheet
    Dim values(1 To 10, 1 To 1) As Variant, raw As Variant
    Dim i As Long, marker As String, defaults As String
    Set result = CreateObject("Scripting.Dictionary")
    For Each candidate In source.Worksheets
        If candidate.name = SETTINGS_SHEET Then Set ws = candidate
    Next candidate
    For i = 1 To 10
        values(i, 1) = "EN"
    Next i
    values(2, 1) = 1: values(8, 1) = 1: values(10, 1) = DATE_MODE_DMY
    If ws Is Nothing Then
        defaults = "GLOBAL=EN; languages=EN; GLOBAL activated=1; INFO=1; date=DMY"
        For Each candidate In source.Worksheets
            If candidate.name = "EVENT_HISTORY" Then
                Dim bg As shape, knob As shape
                On Error Resume Next
                Set bg = candidate.Shapes("btn_EventHistory_Language_BG")
                Set knob = candidate.Shapes("btn_EventHistory_Language_Knob")
                On Error GoTo 0
                If Not bg Is Nothing And Not knob Is Nothing Then
                    values(7, 1) = "FR"
                    If knob.left + knob.width / 2 >= bg.left + bg.width / 2 Then values(7, 1) = "EN"
                End If
                Set bg = Nothing: Set knob = Nothing
                On Error Resume Next
                Set bg = candidate.Shapes("btn_EventHistory_Info_BG")
                Set knob = candidate.Shapes("btn_EventHistory_Info_Knob")
                On Error GoTo 0
                If Not bg Is Nothing And Not knob Is Nothing Then
                    values(8, 1) = CLng(Abs(knob.left + knob.width / 2 >= bg.left + bg.width / 2))
                End If
            End If
        Next candidate
        marker = "NO_SETTINGS"
    Else
        marker = CStr(ws.Range("X1").Value2)
        If marker <> "SETTINGS_STORAGE_V1" And marker <> SETTINGS_STORAGE_VERSION Then
            Err.Raise 5, "Settings_MigrationSnapshot", "UNRECOGNIZED_SETTINGS_STORAGE"
        End If
        raw = ws.Range("X2:X11").Value2
        For i = 1 To 10
            If i = 9 And marker = "SETTINGS_STORAGE_V1" Then
                values(i, 1) = CStr(raw(1, 1))
                defaults = defaults & "CONSTRAINTS=GLOBAL; "
            ElseIf i = 10 And Len(CStr(raw(i, 1))) = 0 Then
                defaults = defaults & "date=DMY; "
            Else
                values(i, 1) = raw(i, 1)
            End If
        Next i
    End If
    For i = 1 To 10
        Select Case i
            Case 2, 8
                If VarType(values(i, 1)) = vbBoolean Then
                    values(i, 1) = CLng(Abs(CBool(values(i, 1))))
                Else
                    Select Case UCase$(Trim$(CStr(values(i, 1))))
                        Case "1", "TRUE", "VRAI", "YES", "Y", "ON": values(i, 1) = 1
                        Case "0", "FALSE", "FAUX", "NO", "N", "OFF": values(i, 1) = 0
                        Case Else: Err.Raise 5, "Settings_MigrationSnapshot", "INVALID_SETTINGS_BOOLEAN"
                    End Select
                End If
            Case 10
                If InStr(1, "|DMY|MDY|ISO|", "|" & CStr(values(i, 1)) & "|", vbBinaryCompare) = 0 Then Err.Raise 5, , "INVALID_DATE_DISPLAY_OPTION"
            Case Else
                If CStr(values(i, 1)) <> "FR" And CStr(values(i, 1)) <> "EN" Then Err.Raise 5, , "INVALID_SETTINGS_LANGUAGE"
        End Select
    Next i
    result.Add "Values", values
    result.Add "Storage", marker
    result.Add "Defaults", defaults
    Set Settings_MigrationSnapshot = result
End Function

Public Sub Settings_ApplyMigrationSnapshot(ByVal values As Variant)
    Dim owners As Variant, owner As Variant, headers As Object, ws As Worksheet
    Dim Index As Long, language As String, errorNumber As Long, errorSource As String, description As String
    On Error GoTo Failed
    Set ws = ThisWorkbook.Worksheets(SETTINGS_SHEET)
    owners = Settings_LanguageOwnerKeys(False)
    For Each owner In owners
        Index = ws.Range(Settings_ModuleStorageCell(CStr(owner))).row - 1
        Settings_PreflightLanguageOwner CStr(owner)
        VisibleHeaders_PreflightOwner CStr(owner), CStr(values(Index, 1))
    Next owner
    Set headers = VisibleHeaders_CaptureAllSnapshot()
    For Each owner In owners
        Index = ws.Range(Settings_ModuleStorageCell(CStr(owner))).row - 1
        VisibleHeaders_ApplyOwnerFromSnapshot headers, CStr(owner), CStr(values(Index, 1))
    Next owner
    ThisWorkbook.Worksheets(SETTINGS_SHEET).Range("X2:X11").Value2 = values
    ThisWorkbook.Worksheets(SETTINGS_SHEET).Range("X1").Value2 = SETTINGS_STORAGE_VERSION
    Settings_HydrateRuntimeState
    For Each owner In owners
        Settings_ApplySingleModule CStr(owner), Settings_GetOwnerLanguage(CStr(owner))
    Next owner
    Schema_ClearPhysicalHeaderLanguageOverrides
    Exit Sub
Failed:
    errorNumber = Err.number: errorSource = Err.source: description = Err.description
    Schema_ClearPhysicalHeaderLanguageOverrides
    Err.Raise errorNumber, errorSource, description
End Sub
'------------------------------------------------------------------------------
' FR: Retourne la langue persistante d'un owner sans modifier le workbook.
' EN: Returns an owner's persisted language without mutating the workbook.
'------------------------------------------------------------------------------
Public Function Settings_GetOwnerLanguage(ByVal ownerKey As String, Optional ByVal wb As Workbook) As String

    Dim ws As Worksheet
    Dim globalLanguage As String
    Dim storageCell As String

    On Error GoTo UseFallback
    If wb Is Nothing Then Set wb = ThisWorkbook
    Set ws = wb.Worksheets(SETTINGS_SHEET)
    globalLanguage = Settings_NormalizeLanguage(CStr(ws.Range(CELL_GLOBAL_LANGUAGE).Value2), "EN")
    storageCell = Settings_ModuleStorageCell(ownerKey)
    If storageCell = "" Then
        Settings_GetOwnerLanguage = globalLanguage
    Else
        Settings_GetOwnerLanguage = _
            Settings_NormalizeLanguage(CStr(ws.Range(storageCell).Value2), globalLanguage)
    End If
    Exit Function

UseFallback:
    Settings_GetOwnerLanguage = "EN"

End Function

'------------------------------------------------------------------------------
' Returns the persisted presentation language used by SETTINGS and, later, the
' workbook Ribbon. It does not replace any functional owner's language.
'------------------------------------------------------------------------------
Public Function Settings_GetGlobalDisplayLanguage() As String

    Dim ws As Worksheet

    On Error GoTo UseFallback
    Set ws = ThisWorkbook.Worksheets(SETTINGS_SHEET)
    Settings_GetGlobalDisplayLanguage = _
        Settings_NormalizeLanguage(CStr(ws.Range(CELL_GLOBAL_LANGUAGE).Value2), "EN")
    Exit Function

UseFallback:
    Settings_GetGlobalDisplayLanguage = "EN"

End Function

'------------------------------------------------------------------------------
' FR: Retourne le reglage Info persiste sans modifier le workbook.
' EN: Returns the persisted Info setting without mutating the workbook.
'------------------------------------------------------------------------------
Public Function Settings_GetInfoEnabled() As Boolean

    Dim ws As Worksheet

    On Error GoTo UseFallback
    Set ws = ThisWorkbook.Worksheets(SETTINGS_SHEET)
    Settings_GetInfoEnabled = Settings_InfoIsEnabled(ws)
    Exit Function

UseFallback:
    Settings_GetInfoEnabled = True

End Function

'------------------------------------------------------------------------------
' FR: Migre explicitement le stockage de langue vers le schema V2.
' EN: Explicitly migrates language storage to the V2 schema.
'------------------------------------------------------------------------------
Public Sub Settings_MigrateLanguageStorageV2()

    Dim ws As Worksheet
    Dim globalLanguage As String

    Set ws = Settings_EnsureSheet()
    globalLanguage = Settings_NormalizeLanguage(CStr(ws.Range(CELL_GLOBAL_LANGUAGE).Value2), "EN")

    If Not Settings_IsRecognizedLanguage(CStr(ws.Range(CELL_CONSTRAINTS_LANGUAGE).Value2)) Then
        ws.Range(CELL_CONSTRAINTS_LANGUAGE).Value2 = globalLanguage
    End If
    ws.Range("X1").Value2 = SETTINGS_STORAGE_VERSION
    ws.columns("X:Y").hidden = True

End Sub

'------------------------------------------------------------------------------
' FR: Met a jour Settings Toggle Info Messages dans le contexte settings and language.
' EN: Updates Settings Toggle Info Messages in the settings and language context.
'------------------------------------------------------------------------------
Public Sub Settings_ToggleInfoMessages()

    Dim ws As Worksheet
    Dim newValue As Boolean

    If Not WorkbookSchema_UserActionAllowed() Then Exit Sub
    Set ws = Settings_EnsureSheet()

    newValue = Not Settings_InfoIsEnabled(ws)
    ws.Range(CELL_INFO_ENABLED).value = newValue

    EventHistory_SetShowInfo newValue
    Refresh_EventHistory_View
    Settings_RefreshVisuals ws

End Sub

'------------------------------------------------------------------------------
' FR: Met a jour Settings Toggle Global Language dans le contexte settings and language.
' EN: Updates Settings Toggle Global Language in the settings and language context.
'------------------------------------------------------------------------------
Public Sub Settings_ToggleGlobalLanguage()

    Dim ws As Worksheet
    Dim languageCode As String

    If Not WorkbookSchema_UserActionAllowed() Then Exit Sub
    Set ws = Settings_EnsureSheet()

    languageCode = Settings_OppositeLanguage(CStr(ws.Range(CELL_GLOBAL_LANGUAGE).value))

    If Settings_GlobalIsActivated(ws) Then
        Settings_ApplyGlobalLanguageTransaction ws, languageCode, True
    Else
        ws.Range(CELL_GLOBAL_LANGUAGE).value = languageCode
        Settings_RefreshVisuals ws
    End If

    Call Finance_Present
    Ribbon_InvalidateLanguage

End Sub

Public Function Settings_SelectGlobalLanguage(ByVal languageCode As String) As Boolean
    Dim ws As Worksheet
    languageCode = UCase$(Trim$(languageCode))
    If languageCode <> "FR" And languageCode <> "EN" Then Err.Raise 5, "Settings_SelectGlobalLanguage", "Unsupported language"
    If ThisWorkbook.readOnly Then Exit Function
    If Not WorkbookSchema_UserActionAllowed() Then Exit Function
    Set ws = Settings_EnsureSheet()
    Settings_ApplyGlobalLanguageTransaction ws, languageCode, True, False
    Call Finance_Present
    Ribbon_InvalidateLanguage
    Settings_SelectGlobalLanguage = True
End Function

'------------------------------------------------------------------------------
' FR: Met a jour Settings Toggle Global Activated dans le contexte settings and language.
' EN: Updates Settings Toggle Global Activated in the settings and language context.
'------------------------------------------------------------------------------
Public Sub Settings_ToggleGlobalActivated()

    Dim ws As Worksheet
    Dim newValue As Boolean

    If Not WorkbookSchema_UserActionAllowed() Then Exit Sub
    Set ws = Settings_EnsureSheet()

    newValue = Not Settings_GlobalIsActivated(ws)
    If newValue Then
        Settings_ApplyGlobalLanguageTransaction ws, _
            Settings_NormalizeLanguage(CStr(ws.Range(CELL_GLOBAL_LANGUAGE).value), "EN"), False
    Else
        ws.Range(CELL_GLOBAL_ACTIVATED).value = False
        Settings_RefreshVisuals ws
    End If

End Sub

'------------------------------------------------------------------------------
' FR: Met a jour Settings Toggle Dashboard Language dans le contexte settings and language.
' EN: Updates Settings Toggle Dashboard Language in the settings and language context.
'------------------------------------------------------------------------------
Public Sub Settings_ToggleDashboardLanguage()
    Settings_ToggleModuleLanguage MODULE_DASHBOARD
End Sub

'------------------------------------------------------------------------------
' FR: Met a jour Settings Toggle Gantt Language dans le contexte settings and language.
' EN: Updates Settings Toggle Gantt Language in the settings and language context.
'------------------------------------------------------------------------------
Public Sub Settings_ToggleGanttLanguage()
    Settings_ToggleModuleLanguage MODULE_GANTT
End Sub

'------------------------------------------------------------------------------
' FR: Met a jour Settings Toggle SCurve Language dans le contexte settings and language.
' EN: Updates Settings Toggle SCurve Language in the settings and language context.
'------------------------------------------------------------------------------
Public Sub Settings_ToggleSCurveLanguage()
    Settings_ToggleModuleLanguage MODULE_SCURVE
End Sub

'------------------------------------------------------------------------------
' FR: Met a jour Settings Toggle WBSLanguage dans le contexte settings and language.
' EN: Updates Settings Toggle WBSLanguage in the settings and language context.
'------------------------------------------------------------------------------
Public Sub Settings_ToggleWBSLanguage()
    Settings_ToggleModuleLanguage MODULE_WBS
End Sub

'------------------------------------------------------------------------------
' FR: Bascule la langue de l'owner Constraints depuis SETTINGS.
' EN: Toggles the Constraints owner language from SETTINGS.
'------------------------------------------------------------------------------
Public Sub Settings_ToggleConstraintsLanguage()
    Settings_ToggleModuleLanguage MODULE_CONSTRAINTS
End Sub

'------------------------------------------------------------------------------
' FR: Met a jour Settings Toggle Event History Language dans le contexte settings and language.
' EN: Updates Settings Toggle Event History Language in the settings and language context.
'------------------------------------------------------------------------------
Public Sub Settings_ToggleEventHistoryLanguage()
    Settings_ToggleModuleLanguage MODULE_EVENT
End Sub

Public Sub Settings_ToggleFinanceLanguage()
    Settings_ToggleModuleLanguage MODULE_FINANCE
    Ribbon_InvalidateLanguage
End Sub

'------------------------------------------------------------------------------
' FR: Retourne le mode d'affichage des dates sans modifier le stockage.
' EN: Returns the date display mode without mutating persisted storage.
'------------------------------------------------------------------------------
Public Function Settings_GetDateDisplayMode() As String

    Dim ws As Worksheet

    On Error GoTo UseFallback
    Set ws = ThisWorkbook.Worksheets(SETTINGS_SHEET)
    Settings_GetDateDisplayMode = _
        Settings_NormalizeDateDisplayMode(CStr(ws.Range(CELL_DATE_DISPLAY_MODE).Value2))
    Exit Function

UseFallback:
    Settings_GetDateDisplayMode = DATE_MODE_DMY

End Function

'------------------------------------------------------------------------------
' FR: Retourne le format invariant de date visible selectionne.
' EN: Returns the selected invariant visible date format.
'------------------------------------------------------------------------------
Public Function Settings_GetDateNumberFormat() As String
    Settings_GetDateNumberFormat = _
        Settings_DateNumberFormatForMode(Settings_GetDateDisplayMode())
End Function

'------------------------------------------------------------------------------
' FR: Retourne le format invariant de date et heure visible selectionne.
' EN: Returns the selected invariant visible date-time format.
'------------------------------------------------------------------------------
Public Function Settings_GetDateTimeNumberFormat() As String
    Settings_GetDateTimeNumberFormat = _
        Settings_DateTimeNumberFormatForMode(Settings_GetDateDisplayMode())
End Function

'------------------------------------------------------------------------------
' FR: Retourne le format invariant compact utilise par les axes visibles.
' EN: Returns the invariant compact format used by visible date axes.
'------------------------------------------------------------------------------
Public Function Settings_GetCompactDateNumberFormat() As String
    Settings_GetCompactDateNumberFormat = _
        Settings_CompactDateNumberFormatForMode(Settings_GetDateDisplayMode())
End Function

'------------------------------------------------------------------------------
' FR: Fait avancer le selecteur partage DMY, MDY, ISO.
' EN: Advances the shared DMY, MDY, ISO selector.
'------------------------------------------------------------------------------
Public Sub Settings_ToggleDateDisplayMode()

    Dim requestedMode As String

    Select Case Settings_GetDateDisplayMode()
        Case DATE_MODE_DMY
            requestedMode = DATE_MODE_MDY
        Case DATE_MODE_MDY
            requestedMode = DATE_MODE_ISO
        Case Else
            requestedMode = DATE_MODE_DMY
    End Select

    Settings_ApplyDateDisplayMode requestedMode

End Sub

'------------------------------------------------------------------------------
' FR: Met a jour Settings Toggle Module Language dans le contexte settings and language.
' EN: Updates Settings Toggle Module Language in the settings and language context.
'------------------------------------------------------------------------------
Private Sub Settings_ToggleModuleLanguage(ByVal moduleKey As String)

    Dim ws As Worksheet
    Dim languageCode As String
    Dim storageCell As String
    Dim oldLanguage As String
    Dim oldStoredValue As Variant
    Dim oldGlobalActivated As Variant
    Dim oldEvents As Boolean
    Dim oldScreenUpdating As Boolean
    Dim headerSnapshot As Object
    Dim stateCaptured As Boolean
    Dim applyStarted As Boolean
    Dim headersApplied As Boolean
    Dim errorDescription As String

    If Not WorkbookSchema_UserActionAllowed() Then Exit Sub
    Set ws = Settings_EnsureSheet()

    storageCell = Settings_ModuleStorageCell(moduleKey)
    If storageCell = "" Then Exit Sub

    oldLanguage = Settings_ModuleLanguage(ws, moduleKey)
    oldStoredValue = ws.Range(storageCell).Value2
    oldGlobalActivated = ws.Range(CELL_GLOBAL_ACTIVATED).Value2
    languageCode = Settings_OppositeLanguage(oldLanguage)
    If languageCode = oldLanguage Then Exit Sub

    On Error GoTo ApplyFailed
    Settings_PreflightLanguageOwner moduleKey
    VisibleHeaders_PreflightOwner moduleKey, languageCode
    Set headerSnapshot = VisibleHeaders_CaptureOwnerSnapshot(moduleKey)

    oldEvents = Application.EnableEvents
    oldScreenUpdating = Application.ScreenUpdating
    stateCaptured = True
    Application.EnableEvents = False
    Application.ScreenUpdating = False

    VisibleHeaders_ApplyOwnerFromSnapshot headerSnapshot, moduleKey, languageCode
    headersApplied = True
    applyStarted = True
    Settings_ApplySingleModule moduleKey, languageCode
    ws.Range(storageCell).Value2 = languageCode
    ws.Range(CELL_GLOBAL_ACTIVATED).Value2 = False
    Settings_RefreshVisuals ws

CleanExit:
    Schema_ClearPhysicalHeaderLanguageOverrides
    If stateCaptured Then
        Application.ScreenUpdating = oldScreenUpdating
        Application.EnableEvents = oldEvents
    End If
    Exit Sub

ApplyFailed:
    errorDescription = Err.source & ": " & Err.description
    On Error Resume Next
    If headersApplied Then VisibleHeaders_RestoreSnapshot headerSnapshot
    If applyStarted Then Settings_ApplySingleModule moduleKey, oldLanguage
    ws.Range(storageCell).Value2 = oldStoredValue
    ws.Range(CELL_GLOBAL_ACTIVATED).Value2 = oldGlobalActivated
    Settings_RefreshVisuals ws
    Schema_ClearPhysicalHeaderLanguageOverrides
    On Error GoTo 0
    If stateCaptured Then
        Application.ScreenUpdating = oldScreenUpdating
        Application.EnableEvents = oldEvents
    End If
    CalcBridge_ShowSingleConsoleMessage "STOP", "COMMON.ERROR.OWNER_LANGUAGE", _
        TextCatalog_Arguments("Owner", moduleKey, "Details", errorDescription)

End Sub

'------------------------------------------------------------------------------
' FR: Met a jour Settings Apply Single Module dans le contexte settings and language.
' EN: Updates Settings Apply Single Module in the settings and language context.
'------------------------------------------------------------------------------
Private Sub Settings_ApplySingleModule( _
    ByVal moduleKey As String, _
    ByVal languageCode As String)

    Select Case UCase$(Trim$(moduleKey))
        Case MODULE_DASHBOARD
            Dashboard_ApplyLanguage languageCode
        Case MODULE_GANTT
            Gantt_ApplyLanguage languageCode
        Case MODULE_SCURVE
            SCurve_ApplyLanguage languageCode
        Case MODULE_WBS
            WBS_ApplyLanguage languageCode
        Case MODULE_CONSTRAINTS
            Constraints_ApplyLanguage languageCode
        Case MODULE_EVENT
            EventHistory_ApplyLanguage languageCode
        Case MODULE_FINANCE
            ThisWorkbook.Worksheets(SETTINGS_SHEET).Range(CELL_FINANCE_LANGUAGE).Value2 = languageCode
            Call Finance_Present
    End Select

End Sub

'------------------------------------------------------------------------------
' FR: Applique les six owners comme une transaction unique depuis GLOBAL.
' EN: Applies all six owners as one transaction from GLOBAL.
'------------------------------------------------------------------------------
Private Sub Settings_ApplyGlobalLanguageTransaction( _
    ByVal ws As Worksheet, _
    ByVal requestedLanguage As String, _
    ByVal persistGlobalLanguage As Boolean, _
    Optional ByVal reportFailureAsConsole As Boolean = True)

    Dim ownerKeys As Variant
    Dim oldLanguages(0 To 6) As String
    Dim oldStoredValues(0 To 6) As Variant
    Dim oldGlobalLanguage As Variant
    Dim oldGlobalActivated As Variant
    Dim oldEvents As Boolean
    Dim oldScreenUpdating As Boolean
    Dim headerSnapshot As Object
    Dim targetLanguage As String
    Dim stateCaptured As Boolean
    Dim applyStarted As Boolean
    Dim headersApplied As Boolean
    Dim failedOwner As String
    Dim errorDescription As String
    Dim i As Long

    targetLanguage = Settings_NormalizeLanguage(requestedLanguage, "EN")
    ownerKeys = Settings_LanguageOwnerKeys()

    On Error GoTo ApplyFailed

    oldGlobalLanguage = ws.Range(CELL_GLOBAL_LANGUAGE).Value2
    oldGlobalActivated = ws.Range(CELL_GLOBAL_ACTIVATED).Value2

    For i = LBound(ownerKeys) To UBound(ownerKeys)
        oldLanguages(i) = Settings_ModuleLanguage(ws, CStr(ownerKeys(i)))
        oldStoredValues(i) = ws.Range(Settings_ModuleStorageCell(CStr(ownerKeys(i)))).Value2
    Next i

    For i = LBound(ownerKeys) To UBound(ownerKeys)
        Settings_PreflightLanguageOwner CStr(ownerKeys(i))
    Next i
    VisibleHeaders_PreflightAll targetLanguage
    Set headerSnapshot = VisibleHeaders_CaptureAllSnapshot()

    oldEvents = Application.EnableEvents
    oldScreenUpdating = Application.ScreenUpdating
    stateCaptured = True
    Application.EnableEvents = False
    Application.ScreenUpdating = False

    VisibleHeaders_ApplyAllFromSnapshot headerSnapshot, targetLanguage
    headersApplied = True
    applyStarted = True
    For i = LBound(ownerKeys) To UBound(ownerKeys)
        failedOwner = CStr(ownerKeys(i))
        Settings_ApplySingleModule failedOwner, targetLanguage
    Next i

    For i = LBound(ownerKeys) To UBound(ownerKeys)
        ws.Range(Settings_ModuleStorageCell(CStr(ownerKeys(i)))).Value2 = targetLanguage
    Next i
    If persistGlobalLanguage Then ws.Range(CELL_GLOBAL_LANGUAGE).Value2 = targetLanguage
    ws.Range(CELL_GLOBAL_ACTIVATED).Value2 = True
    Settings_RefreshVisuals ws

CleanExit:
    Schema_ClearPhysicalHeaderLanguageOverrides
    If stateCaptured Then
        Application.ScreenUpdating = oldScreenUpdating
        Application.EnableEvents = oldEvents
    End If
    Exit Sub

ApplyFailed:
    errorDescription = Err.source & ": " & Err.description
    On Error Resume Next
    If headersApplied Then VisibleHeaders_RestoreSnapshot headerSnapshot
    If applyStarted Then
        For i = UBound(ownerKeys) To LBound(ownerKeys) Step -1
            Settings_ApplySingleModule CStr(ownerKeys(i)), oldLanguages(i)
        Next i
    End If
    For i = LBound(ownerKeys) To UBound(ownerKeys)
        ws.Range(Settings_ModuleStorageCell(CStr(ownerKeys(i)))).Value2 = oldStoredValues(i)
    Next i
    ws.Range(CELL_GLOBAL_LANGUAGE).Value2 = oldGlobalLanguage
    ws.Range(CELL_GLOBAL_ACTIVATED).Value2 = oldGlobalActivated
    Settings_RefreshVisuals ws
    Schema_ClearPhysicalHeaderLanguageOverrides
    On Error GoTo 0
    If stateCaptured Then
        Application.ScreenUpdating = oldScreenUpdating
        Application.EnableEvents = oldEvents
    End If
    If Not reportFailureAsConsole Then _
        Err.Raise vbObjectError + 5212, "Settings_ApplyGlobalLanguageTransaction", errorDescription
    CalcBridge_ShowSingleConsoleMessage "STOP", "COMMON.ERROR.GLOBAL_LANGUAGE", _
        TextCatalog_Arguments("Owner", failedOwner, "Details", errorDescription)

End Sub

'------------------------------------------------------------------------------
' FR: Retourne les owners dans l'ordre transactionnel canonique.
' EN: Returns owners in canonical transaction order.
'------------------------------------------------------------------------------
Private Function Settings_LanguageOwnerKeys(Optional ByVal includeFinance As Boolean = True) As Variant
    Dim keys As Variant
    keys = Array( _
        MODULE_DASHBOARD, MODULE_GANTT, MODULE_SCURVE, _
        MODULE_WBS, MODULE_CONSTRAINTS, MODULE_EVENT)
    If includeFinance Then
        ReDim Preserve keys(0 To 6)
        keys(6) = MODULE_FINANCE
    End If
    Settings_LanguageOwnerKeys = keys
End Function

'------------------------------------------------------------------------------
' FR: Valide les ressources minimales d'un owner avant toute mutation.
' EN: Validates an owner's minimum resources before any mutation.
'------------------------------------------------------------------------------
Private Sub Settings_PreflightLanguageOwner(ByVal moduleKey As String)

    Select Case UCase$(Trim$(moduleKey))
        Case MODULE_DASHBOARD
            Settings_RequireWorksheet "DASHBOARD"
        Case MODULE_GANTT
            Settings_RequireWorksheet "GANTT"
        Case MODULE_SCURVE
            Settings_RequireWorksheet "SCURVE"
        Case MODULE_WBS
            Settings_RequireTable "WBS", "tbl_WBS"
        Case MODULE_CONSTRAINTS
            Settings_RequireTable "CONSTRAINTS", "tbl_CONSTRAINTS"
        Case MODULE_EVENT
            Settings_RequireTable "EVENT_HISTORY", "tbl_EVENT_HISTORY"
            Settings_RequireTable "EVENT_ACK", "tbl_EVENT_ACK"
        Case MODULE_FINANCE
            Settings_RequireTable "FINANCE", FIN_TABLE
            Settings_RequireTable "FINANCE", FIN_EXPENSES
        Case Else
        Err.Raise vbObjectError + 5210, "Settings_PreflightLanguageOwner", _
            PlanningMessageText_Format("DIAG.TECH.UNKNOWN_LANGUAGE_OWNER", _
                TextCatalog_Arguments("Owner", moduleKey), TextCatalog_Arguments("Owner", moduleKey))
    End Select

End Sub

'------------------------------------------------------------------------------
' FR: Met a jour Settings Ensure Sheet dans le contexte settings and language.
' EN: Updates Settings Ensure Sheet in the settings and language context.
'------------------------------------------------------------------------------
Private Function Settings_EnsureSheet() As Worksheet

    Dim ws As Worksheet
    Dim wsAfter As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SETTINGS_SHEET)
    On Error GoTo 0

    If ws Is Nothing Then
        On Error Resume Next
        Set wsAfter = ThisWorkbook.Worksheets("DASHBOARD")
        On Error GoTo 0

        If wsAfter Is Nothing Then
            Set ws = ThisWorkbook.Worksheets.Add( _
                after:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
        Else
            Set ws = ThisWorkbook.Worksheets.Add(after:=wsAfter)
        End If

        ws.name = SETTINGS_SHEET
    End If

    ws.visible = xlSheetVisible
    Set Settings_EnsureSheet = ws

End Function

'------------------------------------------------------------------------------
' FR: Met a jour Settings Info Is Enabled dans le contexte settings and language.
' EN: Updates Settings Info Is Enabled in the settings and language context.
'------------------------------------------------------------------------------
Private Function Settings_InfoIsEnabled(ByVal ws As Worksheet) As Boolean

    Dim rawValue As Variant
    Dim textValue As String

    If ws Is Nothing Then Exit Function

    rawValue = ws.Range(CELL_INFO_ENABLED).value
    If VarType(rawValue) = vbBoolean Then
        Settings_InfoIsEnabled = CBool(rawValue)
        Exit Function
    End If

    textValue = UCase$(Trim$(CStr(rawValue)))
    Settings_InfoIsEnabled = _
        (textValue = "TRUE" Or textValue = "VRAI" Or textValue = "YES" Or _
         textValue = "Y" Or textValue = "1" Or textValue = "ON")

End Function

'------------------------------------------------------------------------------
' FR: Met a jour Settings Global Is Activated dans le contexte settings and language.
' EN: Updates Settings Global Is Activated in the settings and language context.
'------------------------------------------------------------------------------
Private Function Settings_GlobalIsActivated(ByVal ws As Worksheet) As Boolean

    Dim rawValue As Variant
    Dim textValue As String

    If ws Is Nothing Then Exit Function

    rawValue = ws.Range(CELL_GLOBAL_ACTIVATED).value

    If VarType(rawValue) = vbBoolean Then
        Settings_GlobalIsActivated = CBool(rawValue)
        Exit Function
    End If

    textValue = UCase$(Trim$(CStr(rawValue)))
    Settings_GlobalIsActivated = _
        (textValue = "TRUE" Or textValue = "VRAI" Or textValue = "YES" Or textValue = "Y" Or textValue = "1")

End Function

'------------------------------------------------------------------------------
' FR: Met a jour Settings Normalize Language dans le contexte settings and language.
' EN: Updates Settings Normalize Language in the settings and language context.
'------------------------------------------------------------------------------
Private Function Settings_NormalizeLanguage( _
    ByVal languageCode As String, _
    ByVal fallbackLanguage As String) As String

    Select Case UCase$(Trim$(languageCode))
        Case "FR"
            Settings_NormalizeLanguage = "FR"
        Case "EN"
            Settings_NormalizeLanguage = "EN"
        Case Else
            If UCase$(Trim$(fallbackLanguage)) = "FR" Then
                Settings_NormalizeLanguage = "FR"
            Else
                Settings_NormalizeLanguage = "EN"
            End If
    End Select

End Function

'------------------------------------------------------------------------------
' FR: Indique si un code de langue persiste est reconnu.
' EN: Returns whether a persisted language code is recognized.
'------------------------------------------------------------------------------
Private Function Settings_IsRecognizedLanguage(ByVal languageCode As String) As Boolean
    Select Case UCase$(Trim$(languageCode))
        Case "FR", "EN"
            Settings_IsRecognizedLanguage = True
    End Select
End Function

'------------------------------------------------------------------------------
' FR: Met a jour Settings Opposite Language dans le contexte settings and language.
' EN: Updates Settings Opposite Language in the settings and language context.
'------------------------------------------------------------------------------
Private Function Settings_OppositeLanguage(ByVal languageCode As String) As String

    If Settings_NormalizeLanguage(languageCode, "EN") = "FR" Then
        Settings_OppositeLanguage = "EN"
    Else
        Settings_OppositeLanguage = "FR"
    End If

End Function

'------------------------------------------------------------------------------
' FR: Met a jour Settings Module Storage Cell dans le contexte settings and language.
' EN: Updates Settings Module Storage Cell in the settings and language context.
'------------------------------------------------------------------------------
Private Function Settings_ModuleStorageCell(ByVal moduleKey As String) As String

    Select Case UCase$(Trim$(moduleKey))
        Case MODULE_DASHBOARD
            Settings_ModuleStorageCell = CELL_DASHBOARD_LANGUAGE
        Case MODULE_GANTT
            Settings_ModuleStorageCell = CELL_GANTT_LANGUAGE
        Case MODULE_SCURVE
            Settings_ModuleStorageCell = CELL_SCURVE_LANGUAGE
        Case MODULE_WBS
            Settings_ModuleStorageCell = CELL_WBS_LANGUAGE
        Case MODULE_CONSTRAINTS
            Settings_ModuleStorageCell = CELL_CONSTRAINTS_LANGUAGE
        Case MODULE_EVENT
            Settings_ModuleStorageCell = CELL_EVENT_LANGUAGE
        Case MODULE_FINANCE
            Settings_ModuleStorageCell = CELL_FINANCE_LANGUAGE
    End Select

End Function

'------------------------------------------------------------------------------
' FR: Met a jour Settings Module Language dans le contexte settings and language.
' EN: Updates Settings Module Language in the settings and language context.
'------------------------------------------------------------------------------
Private Function Settings_ModuleLanguage( _
    ByVal ws As Worksheet, _
    ByVal moduleKey As String) As String

    Dim cellAddress As String

    cellAddress = Settings_ModuleStorageCell(moduleKey)
    If cellAddress = "" Then
        Settings_ModuleLanguage = "EN"
    Else
        Settings_ModuleLanguage = _
            Settings_NormalizeLanguage( _
                CStr(ws.Range(cellAddress).value), _
                Settings_NormalizeLanguage(CStr(ws.Range(CELL_GLOBAL_LANGUAGE).Value2), "EN"))
    End If

End Function

'------------------------------------------------------------------------------
' FR: Met a jour Settings Build Layout dans le contexte settings and language.
' EN: Updates Settings Build Layout in the settings and language context.
'------------------------------------------------------------------------------
Private Sub Settings_BuildLayout(ByVal ws As Worksheet)

    Dim previousSheet As Object
    Dim dateTitle As shape
    Dim panelLeft As Double
    Dim panelTop As Double
    Dim settingsToggleCenterX As Double
    Dim horizontalScale As Double
    Dim dateTrackWidth As Double
    Dim dateTrackHeight As Double
    Dim dateKnobSize As Double
    Dim dateLabelGap As Double
    Dim dateTrackLeft As Double
    Dim dateTrackTop As Double
    Dim dateLabelLeft As Double

    If ws Is Nothing Then Exit Sub

    On Error Resume Next
    Set previousSheet = Application.ActiveSheet
    On Error GoTo 0

    Settings_DeleteGeneratedShapes ws
    Settings_PrepareCanvas ws

    panelLeft = ws.Range("B4").left
    panelTop = ws.Range("B4").top
    settingsToggleCenterX = panelLeft + 28 + 310 + (SETTINGS_TRACK_WIDTH / 2)
    horizontalScale = SETTINGS_TRACK_WIDTH / GANTT_REFERENCE_TRACK_WIDTH
    dateTrackWidth = GANTT_THREE_POSITION_TRACK_WIDTH * horizontalScale
    dateTrackHeight = GANTT_REFERENCE_TRACK_HEIGHT * _
        (SETTINGS_TRACK_HEIGHT / GANTT_REFERENCE_TRACK_HEIGHT)
    dateKnobSize = GANTT_REFERENCE_KNOB_SIZE * _
        (SETTINGS_KNOB_SIZE / GANTT_REFERENCE_KNOB_SIZE)
    dateLabelGap = GANTT_REFERENCE_LABEL_GAP * horizontalScale
    dateTrackLeft = settingsToggleCenterX - (dateTrackWidth / 2)
    dateTrackTop = panelTop + 432 + ((20 - dateTrackHeight) / 2)
    dateLabelLeft = dateTrackLeft - dateLabelGap - DATE_CONTROL_LABEL_WIDTH

    Settings_AddPanel ws, "SET_PANEL_LANGUAGE", panelLeft, panelTop, 610, 488, RGB(255, 255, 255), RGB(214, 220, 228)
    Settings_AddCenteredTitle ws, "SET_TITLE_LANGUAGE", TextCatalog_Get("SETTINGS.LANGUAGE.TITLE", Settings_ModuleDisplayLanguage(ws, "GLOBAL")), panelLeft + 20, panelTop + 16, 570, 28, 15

    Settings_AddLanguageSwitch ws, "GLOBAL", TextCatalog_Get("SETTINGS.OWNER.GLOBAL", Settings_ModuleDisplayLanguage(ws, "GLOBAL")), panelLeft + 28, panelTop + 64, _
        Settings_ModuleDisplayLanguage(ws, "GLOBAL"), "Settings_ToggleGlobalLanguage", True
    Settings_AddActivatedControl ws, panelLeft + 430, panelTop + 63, Settings_GlobalIsActivated(ws)

    Settings_AddLanguageSwitch ws, MODULE_DASHBOARD, TextCatalog_Get("SETTINGS.OWNER.DASHBOARD", Settings_ModuleDisplayLanguage(ws, "GLOBAL")), panelLeft + 28, panelTop + 118, _
        Settings_ModuleLanguage(ws, MODULE_DASHBOARD), "Settings_ToggleDashboardLanguage", False
    Settings_AddLanguageSwitch ws, MODULE_GANTT, TextCatalog_Get("SETTINGS.OWNER.GANTT", Settings_ModuleDisplayLanguage(ws, "GLOBAL")), panelLeft + 28, panelTop + 162, _
        Settings_ModuleLanguage(ws, MODULE_GANTT), "Settings_ToggleGanttLanguage", False
    Settings_AddLanguageSwitch ws, MODULE_SCURVE, TextCatalog_Get("SETTINGS.OWNER.SCURVE", Settings_ModuleDisplayLanguage(ws, "GLOBAL")), panelLeft + 28, panelTop + 206, _
        Settings_ModuleLanguage(ws, MODULE_SCURVE), "Settings_ToggleSCurveLanguage", False
    Settings_AddLanguageSwitch ws, MODULE_WBS, TextCatalog_Get("SETTINGS.OWNER.WBS", Settings_ModuleDisplayLanguage(ws, "GLOBAL")), panelLeft + 28, panelTop + 250, _
        Settings_ModuleLanguage(ws, MODULE_WBS), "Settings_ToggleWBSLanguage", False
    Settings_AddLanguageSwitch ws, MODULE_CONSTRAINTS, _
        TextCatalog_Get("SETTINGS.CONSTRAINTS.LABEL", Settings_ModuleDisplayLanguage(ws, "GLOBAL")), panelLeft + 28, panelTop + 294, _
        Settings_ModuleLanguage(ws, MODULE_CONSTRAINTS), "Settings_ToggleConstraintsLanguage", False
    Settings_AddLanguageSwitch ws, MODULE_EVENT, TextCatalog_Get("SETTINGS.OWNER.EVENT", Settings_ModuleDisplayLanguage(ws, "GLOBAL")), panelLeft + 28, panelTop + 338, _
        Settings_ModuleLanguage(ws, MODULE_EVENT), "Settings_ToggleEventHistoryLanguage", False
    Settings_AddInfoSwitch ws, panelLeft + 430, panelTop + 338, Settings_InfoIsEnabled(ws)
    Settings_AddLanguageSwitch ws, MODULE_FINANCE, TextCatalog_Get("FINANCE.GROUP", Settings_ModuleDisplayLanguage(ws, "GLOBAL")), panelLeft + 28, panelTop + 388, _
        Settings_ModuleLanguage(ws, MODULE_FINANCE), "Settings_ToggleFinanceLanguage", False

    Set dateTitle = Settings_AddTextShape(ws, "SET_DATE_FORMAT_TITLE", _
        TextCatalog_Get("SETTINGS.DATE_FORMAT.TITLE", Settings_ModuleDisplayLanguage(ws, "GLOBAL")), _
        panelLeft + 28, panelTop + 432, 150, 20, msoAlignLeft, 10.5, True)
    ThreePositionControl_Ensure ws, _
        DATE_CONTROL_LABEL_NAME, DATE_CONTROL_BG_NAME, DATE_CONTROL_KNOB_NAME, _
        dateLabelLeft, panelTop + 432, DATE_CONTROL_LABEL_WIDTH, TextCatalog_Get("SETTINGS.DATE_FORMAT.OPTIONS", Settings_ModuleDisplayLanguage(ws, "GLOBAL")), _
        dateTrackLeft, dateTrackTop, dateTrackWidth, dateTrackHeight, dateKnobSize, _
        "Settings_ToggleDateDisplayMode", xlFreeFloating
    ThreePositionControl_Refresh ws, DATE_CONTROL_BG_NAME, DATE_CONTROL_KNOB_NAME, _
        Settings_DateDisplayPosition(Settings_GetDateDisplayMode())

    Settings_RefreshVisuals ws

    Settings_RemoveLegacyMigrationCommand


    On Error Resume Next
    ws.Activate
    activeWindow.DisplayGridlines = False
    If Not previousSheet Is Nothing Then previousSheet.Activate
    On Error GoTo 0

End Sub

'------------------------------------------------------------------------------
' FR: Met a jour Settings Add Info Switch dans le contexte settings and language.
' EN: Updates Settings Add Info Switch in the settings and language context.
'------------------------------------------------------------------------------
Private Sub Settings_AddInfoSwitch( _
    ByVal ws As Worksheet, _
    ByVal leftPos As Double, _
    ByVal topPos As Double, _
    ByVal isEnabled As Boolean)

    Dim labelShape As shape
    Dim offShape As shape
    Dim trackShape As shape
    Dim knobShape As shape
    Dim onShape As shape
    Dim trackLeft As Double
    Dim knobLeft As Double

    trackLeft = leftPos + 58

    Set labelShape = Settings_AddTextShape(ws, "SET_INFO_LABEL", TextCatalog_Get("SETTINGS.INFO.LABEL", Settings_GetGlobalDisplayLanguage()), leftPos, topPos, 50, 24, msoAlignLeft, 9.5, True)
    Set offShape = Settings_AddTextShape(ws, "SET_INFO_OFF", TextCatalog_Get("SETTINGS.TOGGLE.OFF", Settings_GetGlobalDisplayLanguage()), trackLeft - 34, topPos, 28, 24, msoAlignCenter, 9.5, True)

    Set trackShape = ws.Shapes.AddShape(msoShapeRoundedRectangle, trackLeft, topPos + 4, 44, 16)
    trackShape.name = "SET_INFO_TRACK"
    trackShape.Adjustments.item(1) = 0.5
    trackShape.Placement = xlFreeFloating
    trackShape.OnAction = "Settings_ToggleInfoMessages"

    knobLeft = trackLeft + 2
    If isEnabled Then knobLeft = trackLeft + 28
    Set knobShape = ws.Shapes.AddShape(msoShapeOval, knobLeft, topPos + 6, 12, 12)
    knobShape.name = "SET_INFO_KNOB"
    knobShape.Placement = xlFreeFloating
    knobShape.OnAction = "Settings_ToggleInfoMessages"
    knobShape.fill.ForeColor.RGB = RGB(255, 255, 255)
    knobShape.Line.visible = msoFalse

    Set onShape = Settings_AddTextShape(ws, "SET_INFO_ON", TextCatalog_Get("SETTINGS.TOGGLE.ON", Settings_GetGlobalDisplayLanguage()), trackLeft + 50, topPos, 30, 24, msoAlignCenter, 9.5, True)

    labelShape.OnAction = "Settings_ToggleInfoMessages"
    offShape.OnAction = "Settings_ToggleInfoMessages"
    onShape.OnAction = "Settings_ToggleInfoMessages"
    Settings_FormatToggle trackShape, knobShape, isEnabled

End Sub

'------------------------------------------------------------------------------
' FR: Met a jour Settings Module Display Language dans le contexte settings and language.
' EN: Updates Settings Module Display Language in the settings and language context.
'------------------------------------------------------------------------------
Private Function Settings_ModuleDisplayLanguage( _
    ByVal ws As Worksheet, _
    ByVal moduleKey As String) As String

    If UCase$(moduleKey) = "GLOBAL" Then
        Settings_ModuleDisplayLanguage = _
            Settings_NormalizeLanguage(CStr(ws.Range(CELL_GLOBAL_LANGUAGE).value), "EN")
    Else
        Settings_ModuleDisplayLanguage = Settings_ModuleLanguage(ws, moduleKey)
    End If

End Function

'------------------------------------------------------------------------------
' FR: Met a jour Settings L dans le contexte settings and language.
' EN: Updates Settings L in the settings and language context.
'------------------------------------------------------------------------------
'------------------------------------------------------------------------------
' FR: Met a jour Settings Prepare Canvas dans le contexte settings and language.
' EN: Updates Settings Prepare Canvas in the settings and language context.
'------------------------------------------------------------------------------
Private Sub Settings_PrepareCanvas(ByVal ws As Worksheet)

    ws.cells.Interior.color = RGB(242, 244, 247)
    ws.cells.font.name = "Segoe UI"
    ws.cells.font.color = RGB(35, 45, 58)

    With ws.Range("A1:W30")
        .UnMerge
        .ClearContents
        .Interior.color = RGB(242, 244, 247)
        .font.name = "Segoe UI"
        .font.color = RGB(35, 45, 58)
    End With

    ws.columns("A").ColumnWidth = 2.5
    ws.columns("B:L").ColumnWidth = 12
    ws.rows("1:30").rowHeight = 22
    ws.rows("1:2").rowHeight = 28

    With ws.Range("B1:K2")
        .Merge
        .value = TextCatalog_Get("SETTINGS.PAGE.TITLE", Settings_ModuleDisplayLanguage(ws, "GLOBAL"))
        .font.name = "Segoe UI Semibold"
        .font.size = 22
        .font.bold = True
        .font.color = RGB(28, 42, 58)
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
    End With

    ws.Tab.color = RGB(96, 111, 128)

End Sub


'------------------------------------------------------------------------------
' FR: Met a jour Settings Delete Generated Shapes dans le contexte settings and language.
' EN: Updates Settings Delete Generated Shapes in the settings and language context.
'------------------------------------------------------------------------------
Private Sub Settings_DeleteGeneratedShapes(ByVal ws As Worksheet)

    Dim i As Long

    For i = ws.Shapes.count To 1 Step -1
        If left$(CStr(ws.Shapes(i).name), Len(SETTINGS_PREFIX)) = SETTINGS_PREFIX Then
            ws.Shapes(i).Delete
        End If
    Next i

End Sub

'------------------------------------------------------------------------------
' FR: Met a jour Settings Add Panel dans le contexte settings and language.
' EN: Updates Settings Add Panel in the settings and language context.
'------------------------------------------------------------------------------
Private Sub Settings_AddPanel( _
    ByVal ws As Worksheet, _
    ByVal shapeName As String, _
    ByVal leftPos As Double, _
    ByVal topPos As Double, _
    ByVal widthVal As Double, _
    ByVal heightVal As Double, _
    ByVal fillColor As Long, _
    ByVal lineColor As Long)

    Dim shp As shape

    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, leftPos, topPos, widthVal, heightVal)
    shp.name = shapeName
    shp.Placement = xlFreeFloating
    shp.Adjustments.item(1) = 0.08
    shp.fill.ForeColor.RGB = fillColor
    shp.Line.ForeColor.RGB = lineColor
    shp.Line.weight = 0.8
    shp.Shadow.visible = msoFalse
    shp.ZOrder msoSendToBack

End Sub

'------------------------------------------------------------------------------
' FR: Met a jour Settings Add Centered Title dans le contexte settings and language.
' EN: Updates Settings Add Centered Title in the settings and language context.
'------------------------------------------------------------------------------
Private Sub Settings_AddCenteredTitle( _
    ByVal ws As Worksheet, _
    ByVal shapeName As String, _
    ByVal titleText As String, _
    ByVal leftPos As Double, _
    ByVal topPos As Double, _
    ByVal widthVal As Double, _
    ByVal heightVal As Double, _
    ByVal fontSize As Double)

    Dim shp As shape

    Set shp = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, leftPos, topPos, widthVal, heightVal)
    shp.name = shapeName
    shp.Placement = xlFreeFloating
    shp.Line.visible = msoFalse
    shp.fill.visible = msoFalse
    shp.TextFrame2.TextRange.text = titleText
    shp.TextFrame2.VerticalAnchor = msoAnchorMiddle
    shp.TextFrame2.TextRange.ParagraphFormat.alignment = msoAlignCenter
    shp.TextFrame2.TextRange.font.name = "Segoe UI Semibold"
    shp.TextFrame2.TextRange.font.size = fontSize
    shp.TextFrame2.TextRange.font.bold = msoTrue
    shp.TextFrame2.TextRange.font.fill.ForeColor.RGB = RGB(35, 45, 58)

End Sub

'------------------------------------------------------------------------------
' FR: Met a jour Settings Add Language Switch dans le contexte settings and language.
' EN: Updates Settings Add Language Switch in the settings and language context.
'------------------------------------------------------------------------------
Private Sub Settings_AddLanguageSwitch( _
    ByVal ws As Worksheet, _
    ByVal keyName As String, _
    ByVal labelText As String, _
    ByVal leftPos As Double, _
    ByVal topPos As Double, _
    ByVal languageCode As String, _
    ByVal macroName As String, _
    ByVal emphasize As Boolean)

    Dim labelShape As shape
    Dim frShape As shape
    Dim trackShape As shape
    Dim knobShape As shape
    Dim enShape As shape
    Dim trackLeft As Double
    Dim knobLeft As Double
    Dim isEnglish As Boolean

    isEnglish = (Settings_NormalizeLanguage(languageCode, "EN") = "EN")
    trackLeft = leftPos + 310

    Set labelShape = Settings_AddTextShape(ws, "SET_LANG_" & keyName & "_LABEL", labelText, leftPos, topPos, 250, 24, msoAlignLeft, 10.5, emphasize)
    Set frShape = Settings_AddTextShape(ws, "SET_LANG_" & keyName & "_FR", TextCatalog_Get("SETTINGS.LANGUAGE.FR", Settings_GetGlobalDisplayLanguage()), trackLeft - 34, topPos, 28, 24, msoAlignCenter, 9.5, True)

    Set trackShape = ws.Shapes.AddShape(msoShapeRoundedRectangle, trackLeft, topPos + 4, 44, 16)
    trackShape.name = "SET_LANG_" & keyName & "_TRACK"
    trackShape.Adjustments.item(1) = 0.5
    trackShape.Placement = xlFreeFloating
    trackShape.OnAction = macroName

    knobLeft = trackLeft + 2
    If isEnglish Then knobLeft = trackLeft + 28
    Set knobShape = ws.Shapes.AddShape(msoShapeOval, knobLeft, topPos + 6, 12, 12)
    knobShape.name = "SET_LANG_" & keyName & "_KNOB"
    knobShape.Placement = xlFreeFloating
    knobShape.OnAction = macroName
    knobShape.fill.ForeColor.RGB = RGB(255, 255, 255)
    knobShape.Line.visible = msoFalse

    Set enShape = Settings_AddTextShape(ws, "SET_LANG_" & keyName & "_EN", TextCatalog_Get("SETTINGS.LANGUAGE.EN", Settings_GetGlobalDisplayLanguage()), trackLeft + 50, topPos, 30, 24, msoAlignCenter, 9.5, True)

    labelShape.OnAction = macroName
    frShape.OnAction = macroName
    enShape.OnAction = macroName
    Settings_FormatToggle trackShape, knobShape, isEnglish

End Sub

'------------------------------------------------------------------------------
' FR: Met a jour Settings Add Text Shape dans le contexte settings and language.
' EN: Updates Settings Add Text Shape in the settings and language context.
'------------------------------------------------------------------------------
Private Function Settings_AddTextShape( _
    ByVal ws As Worksheet, _
    ByVal shapeName As String, _
    ByVal textValue As String, _
    ByVal leftPos As Double, _
    ByVal topPos As Double, _
    ByVal widthVal As Double, _
    ByVal heightVal As Double, _
    ByVal alignment As MsoParagraphAlignment, _
    ByVal fontSize As Double, _
    ByVal isBold As Boolean) As shape

    Dim shp As shape

    Set shp = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, leftPos, topPos, widthVal, heightVal)
    shp.name = shapeName
    shp.Placement = xlFreeFloating
    shp.Line.visible = msoFalse
    shp.fill.visible = msoFalse
    shp.TextFrame2.MarginLeft = 0
    shp.TextFrame2.MarginRight = 0
    shp.TextFrame2.MarginTop = 0
    shp.TextFrame2.MarginBottom = 0
    shp.TextFrame2.VerticalAnchor = msoAnchorMiddle
    shp.TextFrame2.TextRange.text = textValue
    shp.TextFrame2.TextRange.ParagraphFormat.alignment = alignment
    shp.TextFrame2.TextRange.font.name = "Segoe UI"
    shp.TextFrame2.TextRange.font.size = fontSize
    shp.TextFrame2.TextRange.font.bold = isBold
    shp.TextFrame2.TextRange.font.fill.ForeColor.RGB = RGB(55, 67, 82)

    Set Settings_AddTextShape = shp

End Function

'------------------------------------------------------------------------------
' FR: Normalise la geometrie commune des toggles Settings, y compris les shapes persistantes.
' EN: Normalizes shared Settings toggle geometry, including persisted shapes.
'------------------------------------------------------------------------------
Private Sub Settings_NormalizeToggleGeometry( _
    ByVal trackShape As shape, _
    ByVal knobShape As shape)

    If trackShape Is Nothing Or knobShape Is Nothing Then Exit Sub

    trackShape.LockAspectRatio = msoFalse
    trackShape.width = 44
    trackShape.height = 16
    trackShape.Adjustments.item(1) = 0.5

    knobShape.LockAspectRatio = msoFalse
    knobShape.width = 12
    knobShape.height = 12

End Sub

'------------------------------------------------------------------------------
' FR: Aligne les labels gauche/droite sur la geometrie commune des toggles Settings.
' EN: Aligns left/right labels to the shared Settings toggle geometry.
'------------------------------------------------------------------------------
Private Sub Settings_NormalizeToggleLabels( _
    ByVal leftShape As shape, _
    ByVal trackShape As shape, _
    ByVal rightShape As shape)

    If trackShape Is Nothing Then Exit Sub

    If Not leftShape Is Nothing Then
        leftShape.left = trackShape.left - 34
        leftShape.top = trackShape.top - 4
        leftShape.width = 28
        leftShape.height = 24
        leftShape.TextFrame2.TextRange.font.size = 9.5
    End If

    If Not rightShape Is Nothing Then
        rightShape.left = trackShape.left + 50
        rightShape.top = trackShape.top - 4
        rightShape.width = 30
        rightShape.height = 24
        rightShape.TextFrame2.TextRange.font.size = 9.5
    End If

End Sub

'------------------------------------------------------------------------------
' FR: Normalise les labels de tous les toggles Settings existants.
' EN: Normalizes labels for all existing Settings toggles.
'------------------------------------------------------------------------------
Private Sub Settings_NormalizeAllToggleLabels(ByVal ws As Worksheet)

    Dim keys As Variant
    Dim keyName As Variant
    Dim leftShape As shape
    Dim trackShape As shape
    Dim rightShape As shape

    If ws Is Nothing Then Exit Sub

    keys = Array( _
        "GLOBAL", MODULE_DASHBOARD, MODULE_GANTT, MODULE_SCURVE, _
        MODULE_WBS, MODULE_CONSTRAINTS, MODULE_EVENT, MODULE_FINANCE)

    For Each keyName In keys
        Set leftShape = Nothing
        Set trackShape = Nothing
        Set rightShape = Nothing
        On Error Resume Next
        Set leftShape = ws.Shapes("SET_LANG_" & CStr(keyName) & "_FR")
        Set trackShape = ws.Shapes("SET_LANG_" & CStr(keyName) & "_TRACK")
        Set rightShape = ws.Shapes("SET_LANG_" & CStr(keyName) & "_EN")
        On Error GoTo 0
        Settings_NormalizeToggleLabels leftShape, trackShape, rightShape
    Next keyName

    Set leftShape = Nothing
    Set trackShape = Nothing
    Set rightShape = Nothing
    On Error Resume Next
    Set leftShape = ws.Shapes("SET_INFO_OFF")
    Set trackShape = ws.Shapes("SET_INFO_TRACK")
    Set rightShape = ws.Shapes("SET_INFO_ON")
    On Error GoTo 0
    Settings_NormalizeToggleLabels leftShape, trackShape, rightShape

End Sub

'------------------------------------------------------------------------------
' FR: Normalise ou formate Settings Format Toggle selon le contrat canonique du composant.
' EN: Normalizes or formats Settings Format Toggle according to the component contract.
'------------------------------------------------------------------------------

Private Sub Settings_FormatToggle( _
    ByVal trackShape As shape, _
    ByVal knobShape As shape, _
    ByVal isRightActive As Boolean)

    If trackShape Is Nothing Or knobShape Is Nothing Then Exit Sub
    Settings_NormalizeToggleGeometry trackShape, knobShape

    If isRightActive Then
        trackShape.fill.ForeColor.RGB = RGB(68, 114, 196)
        trackShape.Line.ForeColor.RGB = RGB(68, 114, 196)
        knobShape.left = trackShape.left + trackShape.width - knobShape.width - 2
    Else
        trackShape.fill.ForeColor.RGB = RGB(150, 160, 172)
        trackShape.Line.ForeColor.RGB = RGB(150, 160, 172)
        knobShape.left = trackShape.left + 2
    End If

    trackShape.Line.weight = 0.75
    knobShape.top = trackShape.top + ((trackShape.height - knobShape.height) / 2)

End Sub
'------------------------------------------------------------------------------
' FR: Met a jour Settings Add Activated Control dans le contexte settings and language.
' EN: Updates Settings Add Activated Control in the settings and language context.
'------------------------------------------------------------------------------
Private Sub Settings_AddActivatedControl( _
    ByVal ws As Worksheet, _
    ByVal leftPos As Double, _
    ByVal topPos As Double, _
    ByVal isActivated As Boolean)

    Dim boxShape As shape
    Dim labelShape As shape

    Set boxShape = ws.Shapes.AddShape(msoShapeRectangle, leftPos, topPos + 2, 18, 18)
    boxShape.name = "SET_GLOBAL_ACTIVATED_BOX"
    boxShape.Placement = xlFreeFloating
    boxShape.OnAction = "Settings_ToggleGlobalActivated"
    boxShape.TextFrame2.TextRange.text = IIf(isActivated, ChrW$(&H2713), "")
    boxShape.TextFrame2.TextRange.font.name = "Segoe UI Symbol"
    boxShape.TextFrame2.TextRange.font.size = 11
    boxShape.TextFrame2.TextRange.font.bold = msoTrue
    boxShape.TextFrame2.TextRange.ParagraphFormat.alignment = msoAlignCenter
    boxShape.TextFrame2.VerticalAnchor = msoAnchorMiddle
    boxShape.TextFrame2.MarginLeft = 0
    boxShape.TextFrame2.MarginRight = 0
    boxShape.TextFrame2.MarginTop = 0
    boxShape.TextFrame2.MarginBottom = 0

    Set labelShape = Settings_AddTextShape(ws, "SET_GLOBAL_ACTIVATED_LABEL", Settings_ActivatedLabel(), leftPos + 26, topPos, 130, 24, msoAlignLeft, 9.5, True)
    labelShape.OnAction = "Settings_ToggleGlobalActivated"

    Settings_FormatActivatedControl boxShape, isActivated

End Sub

'------------------------------------------------------------------------------
' FR: Met a jour Settings Format Activated Control dans le contexte settings and language.
' EN: Updates Settings Format Activated Control in the settings and language context.
'------------------------------------------------------------------------------
Private Sub Settings_FormatActivatedControl( _
    ByVal boxShape As shape, _
    ByVal isActivated As Boolean)

    If boxShape Is Nothing Then Exit Sub

    If isActivated Then
        boxShape.fill.ForeColor.RGB = RGB(68, 114, 196)
        boxShape.Line.ForeColor.RGB = RGB(68, 114, 196)
        boxShape.TextFrame2.TextRange.text = ChrW$(&H2713)
        boxShape.TextFrame2.TextRange.font.fill.ForeColor.RGB = RGB(255, 255, 255)
    Else
        boxShape.fill.ForeColor.RGB = RGB(255, 255, 255)
        boxShape.Line.ForeColor.RGB = RGB(150, 160, 172)
        boxShape.TextFrame2.TextRange.text = ""
    End If

End Sub

'------------------------------------------------------------------------------
' FR: Met a jour Settings Add Command Button dans le contexte settings and language.
' EN: Updates Settings Add Command Button in the settings and language context.
'------------------------------------------------------------------------------
Private Sub Settings_AddCommandButton( _
    ByVal ws As Worksheet, _
    ByVal shapeName As String, _
    ByVal captionText As String, _
    ByVal macroName As String, _
    ByVal leftPos As Double, _
    ByVal topPos As Double, _
    ByVal widthVal As Double, _
    ByVal heightVal As Double, _
    ByVal fillColor As Long, _
    ByVal textColor As Long)

    Dim shp As shape

    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, leftPos, topPos, widthVal, heightVal)
    shp.name = shapeName
    shp.Placement = xlFreeFloating
    shp.Adjustments.item(1) = 0.16
    shp.OnAction = macroName
    shp.fill.ForeColor.RGB = fillColor
    shp.Line.visible = msoFalse
    shp.Shadow.visible = msoFalse
    shp.TextFrame2.TextRange.text = captionText
    shp.TextFrame2.VerticalAnchor = msoAnchorMiddle
    shp.TextFrame2.TextRange.ParagraphFormat.alignment = msoAlignCenter
    shp.TextFrame2.TextRange.font.name = "Segoe UI Semibold"
    shp.TextFrame2.TextRange.font.size = 10
    shp.TextFrame2.TextRange.font.bold = msoTrue
    shp.TextFrame2.TextRange.font.fill.ForeColor.RGB = textColor

End Sub

'------------------------------------------------------------------------------
' FR: Met a jour Settings Refresh Visuals dans le contexte settings and language.
' EN: Updates Settings Refresh Visuals in the settings and language context.
'------------------------------------------------------------------------------
Private Sub Settings_RefreshVisuals(ByVal ws As Worksheet)

    If ws Is Nothing Then Exit Sub

    Settings_RefreshSwitch ws, "GLOBAL", Settings_ModuleDisplayLanguage(ws, "GLOBAL")
    Settings_RefreshSwitch ws, MODULE_DASHBOARD, Settings_ModuleLanguage(ws, MODULE_DASHBOARD)
    Settings_RefreshSwitch ws, MODULE_GANTT, Settings_ModuleLanguage(ws, MODULE_GANTT)
    Settings_RefreshSwitch ws, MODULE_SCURVE, Settings_ModuleLanguage(ws, MODULE_SCURVE)
    Settings_RefreshSwitch ws, MODULE_WBS, Settings_ModuleLanguage(ws, MODULE_WBS)
    Settings_RefreshSwitch ws, MODULE_CONSTRAINTS, Settings_ModuleLanguage(ws, MODULE_CONSTRAINTS)
    Settings_RefreshSwitch ws, MODULE_EVENT, Settings_ModuleLanguage(ws, MODULE_EVENT)
    Settings_RefreshSwitch ws, MODULE_FINANCE, Settings_ModuleLanguage(ws, MODULE_FINANCE)
    Settings_RefreshInfoSwitch ws, Settings_InfoIsEnabled(ws)
    ThreePositionControl_Refresh ws, DATE_CONTROL_BG_NAME, DATE_CONTROL_KNOB_NAME, _
        Settings_DateDisplayPosition(Settings_GetDateDisplayMode())
    Settings_NormalizeAllToggleLabels ws
    Settings_RefreshTitles ws

    On Error Resume Next
    Settings_FormatActivatedControl ws.Shapes("SET_GLOBAL_ACTIVATED_BOX"), Settings_GlobalIsActivated(ws)
    On Error GoTo 0

End Sub

'------------------------------------------------------------------------------
' FR: Met a jour Settings Refresh Titles dans le contexte settings and language.
' EN: Updates Settings Refresh Titles in the settings and language context.
'------------------------------------------------------------------------------
Private Sub Settings_RefreshTitles(ByVal ws As Worksheet)

    If ws Is Nothing Then Exit Sub

    ws.Range("B1").value = TextCatalog_Get("SETTINGS.PAGE.TITLE", Settings_ModuleDisplayLanguage(ws, "GLOBAL"))
    Settings_SetShapeText ws, "SET_TITLE_LANGUAGE", TextCatalog_Get("SETTINGS.LANGUAGE.TITLE", Settings_ModuleDisplayLanguage(ws, "GLOBAL"))
    Settings_SetShapeText ws, "SET_LANG_CONSTRAINTS_LABEL", _
        TextCatalog_Get("SETTINGS.CONSTRAINTS.LABEL", Settings_ModuleDisplayLanguage(ws, "GLOBAL"))
    Settings_SetShapeText ws, "SET_LANG_FINANCE_LABEL", TextCatalog_Get("FINANCE.GROUP", Settings_ModuleDisplayLanguage(ws, "GLOBAL"))
    Settings_SetShapeText ws, "SET_DATE_FORMAT_TITLE", TextCatalog_Get("SETTINGS.DATE_FORMAT.TITLE", Settings_ModuleDisplayLanguage(ws, "GLOBAL"))

End Sub

'------------------------------------------------------------------------------
' FR: Met a jour Settings Refresh Command Captions dans le contexte settings and language.
' EN: Updates Settings Refresh Command Captions in the settings and language context.
'------------------------------------------------------------------------------
'------------------------------------------------------------------------------
' FR: Met a jour Settings Set Shape Text dans le contexte settings and language.
' EN: Updates Settings Set Shape Text in the settings and language context.
'------------------------------------------------------------------------------
Private Sub Settings_SetShapeText( _
    ByVal ws As Worksheet, _
    ByVal shapeName As String, _
    ByVal captionText As String)

    On Error Resume Next
    ws.Shapes(shapeName).TextFrame2.TextRange.text = captionText
    On Error GoTo 0

End Sub

'------------------------------------------------------------------------------
' FR: Met a jour Settings Refresh Info Switch dans le contexte settings and language.
' EN: Updates Settings Refresh Info Switch in the settings and language context.
'------------------------------------------------------------------------------
Private Sub Settings_RefreshInfoSwitch( _
    ByVal ws As Worksheet, _
    ByVal isEnabled As Boolean)

    Dim trackShape As shape
    Dim knobShape As shape

    On Error Resume Next
    Set trackShape = ws.Shapes("SET_INFO_TRACK")
    Set knobShape = ws.Shapes("SET_INFO_KNOB")
    On Error GoTo 0

    Settings_FormatToggle trackShape, knobShape, isEnabled

End Sub

'------------------------------------------------------------------------------
' FR: Met a jour Settings Refresh Switch dans le contexte settings and language.
' EN: Updates Settings Refresh Switch in the settings and language context.
'------------------------------------------------------------------------------
Private Sub Settings_RefreshSwitch( _
    ByVal ws As Worksheet, _
    ByVal keyName As String, _
    ByVal languageCode As String)

    Dim trackShape As shape
    Dim knobShape As shape

    On Error Resume Next
    Set trackShape = ws.Shapes("SET_LANG_" & keyName & "_TRACK")
    Set knobShape = ws.Shapes("SET_LANG_" & keyName & "_KNOB")
    On Error GoTo 0

    Settings_FormatToggle trackShape, knobShape, _
        (Settings_NormalizeLanguage(languageCode, "EN") = "EN")

End Sub

'------------------------------------------------------------------------------
' FR: Met a jour Settings Reset Title dans le contexte settings and language.
' EN: Updates Settings Reset Title in the settings and language context.
'------------------------------------------------------------------------------
'------------------------------------------------------------------------------
' FR: Met a jour Settings Activated Label dans le contexte settings and language.
' EN: Updates Settings Activated Label in the settings and language context.
'------------------------------------------------------------------------------
Private Function Settings_ActivatedLabel() As String
    Settings_ActivatedLabel = TextCatalog_Get("SETTINGS.ACTIVATION.LABEL", Settings_GetGlobalDisplayLanguage())
End Function

'------------------------------------------------------------------------------
' FR: Applique transactionnellement un nouveau format visible aux surfaces produit.
' EN: Transactionally applies a new visible format to product-facing surfaces.
'------------------------------------------------------------------------------
Public Sub Settings_ApplyDateDisplayMode(ByVal requestedMode As String)

    Dim ws As Worksheet
    Dim oldMode As String
    Dim normalizedRequested As String
    Dim oldStoredValue As Variant
    Dim oldEvents As Boolean
    Dim oldScreenUpdating As Boolean
    Dim stateCaptured As Boolean
    Dim formattingStarted As Boolean
    Dim errorDescription As String

    If Not WorkbookSchema_UserActionAllowed() Then Exit Sub
    If Not Settings_IsValidDateDisplayMode(requestedMode) Then
        CalcBridge_ShowSingleConsoleMessage "STOP", "COMMON.ERROR.DATE_MODE", _
            TextCatalog_Arguments("Mode", requestedMode)
        Exit Sub
    End If

    normalizedRequested = Settings_NormalizeDateDisplayMode(requestedMode)
    oldMode = Settings_GetDateDisplayMode()
    If normalizedRequested = oldMode Then Exit Sub

    On Error GoTo ApplyFailed

    Set ws = ThisWorkbook.Worksheets(SETTINGS_SHEET)
    oldStoredValue = ws.Range(CELL_DATE_DISPLAY_MODE).Value2
    Settings_PreflightDateDisplaySurfaces

    oldEvents = Application.EnableEvents
    oldScreenUpdating = Application.ScreenUpdating
    stateCaptured = True
    Application.EnableEvents = False
    Application.ScreenUpdating = False

    formattingStarted = True
    Settings_ApplyDateDisplayFormatsForMode normalizedRequested
    ws.Range(CELL_DATE_DISPLAY_MODE).Value2 = normalizedRequested
    ThreePositionControl_Refresh ws, DATE_CONTROL_BG_NAME, DATE_CONTROL_KNOB_NAME, _
        Settings_DateDisplayPosition(normalizedRequested)

CleanExit:
    If stateCaptured Then
        Application.ScreenUpdating = oldScreenUpdating
        Application.EnableEvents = oldEvents
    End If
    Exit Sub

ApplyFailed:
    errorDescription = Err.description
    On Error Resume Next
    If formattingStarted Then Settings_ApplyDateDisplayFormatsForMode oldMode
    If Not ws Is Nothing Then
        ws.Range(CELL_DATE_DISPLAY_MODE).Value2 = oldStoredValue
        ThreePositionControl_Refresh ws, DATE_CONTROL_BG_NAME, DATE_CONTROL_KNOB_NAME, _
            Settings_DateDisplayPosition(oldMode)
    End If
    On Error GoTo 0
    If stateCaptured Then
        Application.ScreenUpdating = oldScreenUpdating
        Application.EnableEvents = oldEvents
    End If
    CalcBridge_ShowSingleConsoleMessage "STOP", "COMMON.ERROR.DATE_FORMAT", _
        TextCatalog_Arguments("Details", errorDescription)

End Sub

'------------------------------------------------------------------------------
' FR: Valide toutes les surfaces obligatoires avant la premiere ecriture de format.
' EN: Validates every required surface before the first format write.
'------------------------------------------------------------------------------
Private Sub Settings_PreflightDateDisplaySurfaces()

    Dim tbl As ListObject
    Dim scurveTbl As ListObject
    Dim ch As chart

    Set tbl = Settings_RequireTable("WBS", "tbl_WBS")
    Settings_RequireColumns tbl, VTS_TABLE_WBS, Array( _
        VTS_COL_BASELINE_START, VTS_COL_BASELINE_FINISH, VTS_COL_ACTUAL_START, VTS_COL_ACTUAL_FINISH, _
        VTS_COL_FORECAST_START, VTS_COL_FORECAST_FINISH, _
        VTS_COL_CALCULATED_START, VTS_COL_CALCULATED_FINISH)

    Settings_RequireWorksheet "GANTT"

    Set scurveTbl = Settings_RequireTable("SCURVE", "tbl_SCURVE")
    Settings_RequireColumns scurveTbl, VTS_TABLE_SCURVE, Array(VTS_COL_DATE)
    Set ch = Settings_RequireVisibleDateChart( _
        "SCURVE", "cht_SCurve", scurveTbl, VTS_TABLE_SCURVE, VTS_COL_DATE)

    Set tbl = Settings_RequireTable("CONSTRAINTS", "tbl_CONSTRAINTS")
    Settings_RequireColumns tbl, VTS_TABLE_CONSTRAINTS, Array( _
        VTS_COL_CALCULATED_START, VTS_COL_CALCULATED_FINISH, VTS_COL_START_CONSTRAINT_DATE, _
        VTS_COL_FINISH_CONSTRAINT_DATE, VTS_COL_DEADLINE)

    Set tbl = Settings_RequireTable("EVENT_HISTORY", "tbl_EVENT_HISTORY")
    Settings_RequireColumns tbl, VTS_TABLE_EVENT_HISTORY, Array(VTS_COL_DATE)
    Set tbl = Settings_RequireTable("EVENT_ACK", "tbl_EVENT_ACK")
    Settings_RequireColumns tbl, VTS_TABLE_EVENT_ACK, Array(VTS_COL_ACKNOWLEDGED_AT)

    Set ch = Settings_RequireVisibleDateChart( _
        "DASHBOARD", "cht_Dashboard_SCurve", scurveTbl, VTS_TABLE_SCURVE, VTS_COL_DATE)

End Sub

'------------------------------------------------------------------------------
' FR: Applique uniquement NumberFormat aux surfaces visibles contractuelles.
' EN: Applies only NumberFormat to the contracted visible surfaces.
'------------------------------------------------------------------------------
Private Sub Settings_ApplyDateDisplayFormatsForMode(ByVal modeValue As String)

    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim scurveTbl As ListObject
    Dim ch As chart
    Dim dateFormat As String
    Dim dateTimeFormat As String
    Dim compactFormat As String
    Dim lastRow As Long

    dateFormat = Settings_DateNumberFormatForMode(modeValue)
    dateTimeFormat = Settings_DateTimeNumberFormatForMode(modeValue)
    compactFormat = Settings_CompactDateNumberFormatForMode(modeValue)

    Set tbl = Settings_RequireTable("WBS", "tbl_WBS")
    Settings_ApplyTableColumnFormats tbl, VTS_TABLE_WBS, Array( _
        VTS_COL_BASELINE_START, VTS_COL_BASELINE_FINISH, VTS_COL_ACTUAL_START, VTS_COL_ACTUAL_FINISH, _
        VTS_COL_FORECAST_START, VTS_COL_FORECAST_FINISH, _
        VTS_COL_CALCULATED_START, VTS_COL_CALCULATED_FINISH), _
        dateFormat

    Set ws = Settings_RequireWorksheet("GANTT")
    lastRow = ws.cells(ws.rows.count, 1).End(xlUp).row
    If lastRow < 5 Then lastRow = 5
    ws.Range(ws.cells(5, 3), ws.cells(lastRow, 6)).numberFormat = dateFormat

    Set scurveTbl = Settings_RequireTable("SCURVE", "tbl_SCURVE")
    Settings_ApplyTableColumnFormats scurveTbl, VTS_TABLE_SCURVE, Array(VTS_COL_DATE), dateFormat
    Set ch = Settings_RequireVisibleDateChart( _
        "SCURVE", "cht_SCurve", scurveTbl, VTS_TABLE_SCURVE, VTS_COL_DATE)
    If Not ch Is Nothing Then ch.axes(xlCategory).TickLabels.numberFormat = dateFormat

    Set tbl = Settings_RequireTable("CONSTRAINTS", "tbl_CONSTRAINTS")
    Settings_ApplyTableColumnFormats tbl, VTS_TABLE_CONSTRAINTS, Array( _
        VTS_COL_CALCULATED_START, VTS_COL_CALCULATED_FINISH, VTS_COL_START_CONSTRAINT_DATE, _
        VTS_COL_FINISH_CONSTRAINT_DATE, VTS_COL_DEADLINE), dateFormat

    Set tbl = Settings_RequireTable("EVENT_HISTORY", "tbl_EVENT_HISTORY")
    Settings_ApplyTableColumnFormats tbl, VTS_TABLE_EVENT_HISTORY, Array(VTS_COL_DATE), dateFormat
    Set tbl = Settings_RequireTable("EVENT_ACK", "tbl_EVENT_ACK")
    Settings_ApplyTableColumnFormats tbl, VTS_TABLE_EVENT_ACK, Array(VTS_COL_ACKNOWLEDGED_AT), dateTimeFormat

    Set ch = Settings_RequireVisibleDateChart( _
        "DASHBOARD", "cht_Dashboard_SCurve", scurveTbl, VTS_TABLE_SCURVE, VTS_COL_DATE)
    If Not ch Is Nothing Then ch.axes(xlCategory).TickLabels.numberFormat = compactFormat

End Sub

Private Sub Settings_ApplyTableColumnFormats( _
    ByVal tbl As ListObject, _
    ByVal tableKey As String, _
    ByVal columnKeys As Variant, _
    ByVal numberFormat As String)

    Dim columnKey As Variant

    For Each columnKey In columnKeys
        SchemaListColumn(tbl, tableKey, CStr(columnKey)).Range.numberFormat = numberFormat
    Next columnKey

End Sub

Private Function Settings_RequireWorksheet(ByVal sheetName As String) As Worksheet

    On Error Resume Next
    Set Settings_RequireWorksheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If Settings_RequireWorksheet Is Nothing Then
    Err.Raise vbObjectError + 5281, "Settings_RequireWorksheet", _
        PlanningMessageText_Format("DIAG.TECH.REQUIRED_WORKSHEET_MISSING", _
            TextCatalog_Arguments("Sheet", sheetName), TextCatalog_Arguments("Sheet", sheetName))
    End If

End Function

Private Function Settings_RequireTable( _
    ByVal sheetName As String, _
    ByVal tableName As String) As ListObject

    Dim ws As Worksheet

    Set ws = Settings_RequireWorksheet(sheetName)
    On Error Resume Next
    Set Settings_RequireTable = ws.ListObjects(tableName)
    On Error GoTo 0
    If Settings_RequireTable Is Nothing Then
    Err.Raise vbObjectError + 5282, "Settings_RequireTable", _
        PlanningMessageText_Format("DIAG.TECH.REQUIRED_TABLE_MISSING", _
            TextCatalog_Arguments("Object", sheetName & "!" & tableName), _
            TextCatalog_Arguments("Object", sheetName & "!" & tableName))
    End If

End Function

Private Function Settings_RequireVisibleDateChart( _
    ByVal sheetName As String, _
    ByVal chartName As String, _
    ByVal dateTable As ListObject, _
    ByVal tableKey As String, _
    ByVal dateColumnKey As String) As chart

    Dim ws As Worksheet
    Dim dateRange As Range

    Set ws = Settings_RequireWorksheet(sheetName)
    On Error Resume Next
    Set Settings_RequireVisibleDateChart = ws.ChartObjects(chartName).chart
    On Error GoTo 0

    If Settings_RequireVisibleDateChart Is Nothing Then
        Set dateRange = SchemaListColumn(dateTable, tableKey, dateColumnKey).DataBodyRange
        If Not dateRange Is Nothing Then
            If Application.count(dateRange) > 0 Then
    Err.Raise vbObjectError + 5283, "Settings_RequireVisibleDateChart", _
        PlanningMessageText_Format("DIAG.TECH.REQUIRED_CHART_MISSING", _
            TextCatalog_Arguments("Object", sheetName & "!" & chartName), _
            TextCatalog_Arguments("Object", sheetName & "!" & chartName))
            End If
        End If
        Exit Function
    End If

    If Not Settings_RequireVisibleDateChart.HasAxis(xlCategory) Then
    Err.Raise vbObjectError + 5284, "Settings_RequireVisibleDateChart", _
        PlanningMessageText_Format("DIAG.TECH.REQUIRED_DATE_AXIS_MISSING", _
            TextCatalog_Arguments("Object", sheetName & "!" & chartName), _
            TextCatalog_Arguments("Object", sheetName & "!" & chartName))
    End If

End Function

Private Sub Settings_RequireColumns( _
    ByVal tbl As ListObject, _
    ByVal tableKey As String, _
    ByVal columnKeys As Variant)

    Dim columnKey As Variant
    Dim listColumn As listColumn

    For Each columnKey In columnKeys
        Set listColumn = Nothing
        On Error Resume Next
        Set listColumn = SchemaListColumn(tbl, tableKey, CStr(columnKey))
        On Error GoTo 0
        If listColumn Is Nothing Then
            Err.Raise vbObjectError + 5285, "Settings_RequireColumns", _
                PlanningMessageText_Format("DIAG.TECH.REQUIRED_COLUMN_MISSING", _
                    TextCatalog_Arguments("Column", tbl.name & "[" & CStr(columnKey) & "]"), _
                    TextCatalog_Arguments("Column", tbl.name & "[" & CStr(columnKey) & "]"))
        End If
    Next columnKey

End Sub

Private Function Settings_IsValidDateDisplayMode(ByVal modeValue As String) As Boolean

    Select Case UCase$(Trim$(modeValue))
        Case DATE_MODE_DMY, DATE_MODE_MDY, DATE_MODE_ISO
            Settings_IsValidDateDisplayMode = True
    End Select

End Function

Private Function Settings_NormalizeDateDisplayMode(ByVal modeValue As String) As String

    Select Case UCase$(Trim$(modeValue))
        Case DATE_MODE_MDY
            Settings_NormalizeDateDisplayMode = DATE_MODE_MDY
        Case DATE_MODE_ISO
            Settings_NormalizeDateDisplayMode = DATE_MODE_ISO
        Case Else
            Settings_NormalizeDateDisplayMode = DATE_MODE_DMY
    End Select

End Function

Private Function Settings_DateNumberFormatForMode(ByVal modeValue As String) As String

    Select Case Settings_NormalizeDateDisplayMode(modeValue)
        Case DATE_MODE_MDY
            Settings_DateNumberFormatForMode = DATE_FORMAT_MDY
        Case DATE_MODE_ISO
            Settings_DateNumberFormatForMode = DATE_FORMAT_ISO
        Case Else
            Settings_DateNumberFormatForMode = DATE_FORMAT_DMY
    End Select

End Function

Private Function Settings_DateTimeNumberFormatForMode(ByVal modeValue As String) As String

    Select Case Settings_NormalizeDateDisplayMode(modeValue)
        Case DATE_MODE_MDY
            Settings_DateTimeNumberFormatForMode = DATETIME_FORMAT_MDY
        Case DATE_MODE_ISO
            Settings_DateTimeNumberFormatForMode = DATETIME_FORMAT_ISO
        Case Else
            Settings_DateTimeNumberFormatForMode = DATETIME_FORMAT_DMY
    End Select

End Function

Private Function Settings_CompactDateNumberFormatForMode(ByVal modeValue As String) As String

    Select Case Settings_NormalizeDateDisplayMode(modeValue)
        Case DATE_MODE_MDY
            Settings_CompactDateNumberFormatForMode = DATE_COMPACT_FORMAT_MDY
        Case DATE_MODE_ISO
            Settings_CompactDateNumberFormatForMode = DATE_COMPACT_FORMAT_ISO
        Case Else
            Settings_CompactDateNumberFormatForMode = DATE_COMPACT_FORMAT_DMY
    End Select

End Function

Private Function Settings_DateDisplayPosition(ByVal modeValue As String) As Long

    Select Case Settings_NormalizeDateDisplayMode(modeValue)
        Case DATE_MODE_MDY
            Settings_DateDisplayPosition = 1
        Case DATE_MODE_ISO
            Settings_DateDisplayPosition = 2
        Case Else
            Settings_DateDisplayPosition = 0
    End Select

End Function

Public Function Settings_GetTutorialProgress(Optional ByVal mode As String = "DISCOVERY") As Variant
    Dim ws As Worksheet
    Dim values As Variant
    Dim slot As Long
    values = Settings_TutorialState()
    slot = IIf(mode = "PRACTICE", 4, 2)
    Settings_GetTutorialProgress = Array(CStr(values(0)), CStr(values(slot)), CStr(values(slot + 1)))
End Function

Private Function Settings_TutorialState() As Variant
    Dim values As Variant, state As Variant
    Dim i As Long
    If IsArray(mTutorialSession) Then
        Settings_TutorialState = mTutorialSession
        Exit Function
    End If
    state = Array("", "", "", "", "", "")
    On Error GoTo SessionOnly
    values = ThisWorkbook.Worksheets(SETTINGS_SHEET).Range("X12:X18").Value2
    If CStr(values(1, 1)) = "TUTORIAL_PROGRESS_V2" Then
        For i = 0 To 5
            state(i) = CStr(values(i + 2, 1))
        Next i
    ElseIf CStr(values(1, 1)) = TUTORIAL_STORAGE_VERSION Then
        ' V1 was a passive visit. Keep its key for explicit mapping by the course owner.
        state(0) = CStr(values(2, 1))
        state(1) = "DISCOVERY"
        state(2) = CStr(values(3, 1))
        state(3) = CStr(values(4, 1))
    End If
SessionOnly:
    mTutorialSession = state
    Settings_TutorialState = state
End Function

Public Function Settings_GetTutorialMode() As String
    Dim state As Variant
    state = Settings_TutorialState()
    Settings_GetTutorialMode = CStr(state(1))
End Function

Public Function Settings_TutorialConfigurationTarget() As Range
    Dim ws As Worksheet
    Dim panel As shape
    On Error GoTo missing
    Set ws = ThisWorkbook.Worksheets(SETTINGS_SHEET)
    Set panel = ws.Shapes("SET_PANEL_LANGUAGE")
    Set Settings_TutorialConfigurationTarget = ws.Range(panel.TopLeftCell, panel.BottomRightCell)
missing:
End Function

Public Function Settings_SetTutorialProgress(ByVal version As String, ByVal stepKey As String, _
                                             ByVal state As String, Optional ByVal mode As String = "DISCOVERY") As Boolean
    Dim ws As Worksheet
    Dim values(1 To 7, 1 To 1) As Variant
    Dim existing As Variant, session As Variant
    Dim i As Long, slot As Long
    If state <> "IN_PROGRESS" And state <> "COMPLETED" Then Err.Raise 5, "Settings_SetTutorialProgress", "Invalid tutorial state"
    If mode <> "DISCOVERY" And mode <> "PRACTICE" Then Err.Raise 5, "Settings_SetTutorialProgress", "Invalid tutorial mode"
    session = Settings_TutorialState()
    slot = IIf(mode = "PRACTICE", 4, 2)
    session(0) = version: session(1) = mode
    session(slot) = stepKey: session(slot + 1) = state
    mTutorialSession = session
    If ThisWorkbook.readOnly Or Not WorkbookSchema_WritesAllowed() Then Exit Function
    Set ws = ThisWorkbook.Worksheets(SETTINGS_SHEET)
    existing = ws.Range("X12:X18").Value2
    If CStr(existing(1, 1)) = TUTORIAL_STORAGE_VERSION Then
        For i = 5 To 7
            If Not IsEmpty(existing(i, 1)) Then Exit Function
        Next i
    ElseIf CStr(existing(1, 1)) <> "TUTORIAL_PROGRESS_V2" Then
        For i = 1 To 7
            If Not IsEmpty(existing(i, 1)) Then Exit Function
        Next i
    End If
    values(1, 1) = "TUTORIAL_PROGRESS_V2"
    For i = 0 To 5
        values(i + 2, 1) = session(i)
    Next i
    On Error GoTo SessionOnly
    ws.Range("X12:X18").Value2 = values
    Settings_SetTutorialProgress = True
SessionOnly:
End Function

Public Sub Settings_ResetTutorialProgress()
    Dim ws As Worksheet
    mTutorialSession = Empty
    Set ws = ThisWorkbook.Worksheets(SETTINGS_SHEET)
    If CStr(ws.Range("X12").Value2) = TUTORIAL_STORAGE_VERSION Then ws.Range("X12:X15").ClearContents
    If CStr(ws.Range("X12").Value2) = "TUTORIAL_PROGRESS_V2" Then ws.Range("X12:X18").ClearContents
End Sub

Public Function Settings_FinanceVocabulary(Optional ByVal wb As Workbook) As String
    Dim value As String
    If wb Is Nothing Then Set wb = ThisWorkbook
    value = CStr(wb.Worksheets("SETTINGS").Range("X19").Value2)
    If value = "" Then value = "BEGINNER"
    If value <> "BEGINNER" And value <> "EXPERT" Then Err.Raise 5, "Settings", "Invalid Finance vocabulary"
    Settings_FinanceVocabulary = value
End Function

Public Sub Settings_SetFinanceVocabulary(ByVal value As String)
    If value <> "BEGINNER" And value <> "EXPERT" Then Err.Raise 5, "Settings", "Invalid Finance vocabulary"
    ThisWorkbook.Worksheets("SETTINGS").Range("X19").Value2 = value
End Sub









