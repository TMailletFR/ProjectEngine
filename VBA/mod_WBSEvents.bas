Attribute VB_Name = "mod_WBSEvents"

'===============================================================================
' MODULE : mod_WBSEvents
' DOMAINE / DOMAIN : WBS
'
' FR
' Valide les edits WBS, restaure les formules gerees et route les changements autorises.
' Ne doit pas contourner les contrats publics des autres domaines.
'
' EN
' Validates WBS edits, restores managed formulas and routes authorized changes.
' Must not bypass public contracts owned by other domains.
'
' CONTRATS / CONTRACTS : Handle_WBS_Change
' CALLBACKS EXTERNES / EXTERNAL CALLBACKS : Handle_WBS_Change
'===============================================================================

'=================================================
' WBS EVENT HANDLER – INPUT / OUTPUT CONTROL
'
' Philosophy:
' - Blue columns = user input (editable)
' - Gray columns = calculated (read-only)
'
' Rules:
' - Manual edits in gray columns are blocked
' - Structural operations (row insert, table expansion) are allowed
' - No write operations here to preserve Excel Undo
'
' Validation:
' - WBS format
' - Predecessors WBS format
' - Lag numeric format
' - Duration numeric format
'
' Macro behavior:
' - events remain active as guard rails
' - engine writes in gray columns are allowed ONLY when
'   explicitly declared through BeginAuthorizedWBSWrite
' - all other engine writes trigger macro abort
'=================================================

'------------------------------------------------------------------------------
' FR: Traite un changement ou evenement pour WBS Change.
' EN: Handles a change or event for WBS Change.
'------------------------------------------------------------------------------
Public Sub Handle_WBS_Change(ByVal ws As Worksheet, ByVal Target As Range)

    Dim tbl As ListObject

    Dim rngWBS As Range
    Dim rngPredWBS As Range
    Dim rngTaskType As Range
    Dim rngSummaryDisplay As Range
    Dim rngCal As Range
    Dim rngBaselineDuration As Range

    Dim rngLockedCols As Range
    Dim rngEditableCols As Range
    Dim rngToCheck As Range
    Dim rngLockedTouched As Range

    Dim cell As Range
    Dim cellValue As String
    Dim normalizedPredWBS As String

    Dim reWBS As Object

    Dim col As ListColumn
    Dim authorizedHit As Boolean

    On Error GoTo SafeExit

    Set tbl = ws.ListObjects("tbl_WBS")
    If tbl.DataBodyRange Is Nothing Then Exit Sub

    Set rngWBS = SchemaListColumn(tbl, VTS_TABLE_WBS, VTS_COL_WBS).DataBodyRange
    Set rngPredWBS = SchemaListColumn(tbl, VTS_TABLE_WBS, VTS_COL_PREDECESSORS_WBS).DataBodyRange
    Set rngTaskType = SchemaListColumn(tbl, VTS_TABLE_WBS, VTS_COL_TASK_TYPE).DataBodyRange
    If WBSHasSchemaColumn(tbl, VTS_COL_S) Then Set rngSummaryDisplay = SchemaListColumn(tbl, VTS_TABLE_WBS, VTS_COL_S).DataBodyRange
    If WBSHasSchemaColumn(tbl, VTS_COL_CAL) Then Set rngCal = SchemaListColumn(tbl, VTS_TABLE_WBS, VTS_COL_CAL).DataBodyRange
    Set rngBaselineDuration = SchemaListColumn(tbl, VTS_TABLE_WBS, VTS_COL_BASELINE_DURATION).DataBodyRange

    Set rngEditableCols = GetWBSUserEditableRange(tbl)
    Set rngLockedCols = GetWBSLockedRange(tbl)

    '=================================================
    ' Gray columns protection
    '=================================================
    Set rngLockedTouched = Intersect(Target, rngLockedCols)

    If Not rngLockedTouched Is Nothing Then

        ' Structural mixed operations remain allowed if they also touch blue columns
        If Intersect(Target, rngEditableCols) Is Nothing Then

            '=================================================
            ' Authorized engine writes have priority over macro-run state.
            ' If the write guard is active and all touched locked columns
            ' are explicitly authorized, this is a valid internal write.
            '=================================================
            If IsAuthorizedWBSWriteActive() Then

                authorizedHit = True

                For Each col In tbl.ListColumns
                    If Not Intersect(rngLockedTouched, col.DataBodyRange) Is Nothing Then
                        If Not IsAuthorizedWBSColumn(col.Name) Then
                            authorizedHit = False
                            Exit For
                        End If
                    End If
                Next col

                If authorizedHit Then
                    GoTo ContinueValidation
                End If

            End If

            If IsWBSFormulaColumnAutofillEvent(tbl, Target, rngLockedTouched, rngEditableCols) Then
                GoTo SafeExit
            End If

            If IsMacroRunActive() Then

                RequestMacroAbortKey _
                    "Handle_WBS_Change", _
                    "WBS.EVENTS.MACRO.CALCULATED_EDIT", _
                    TextCatalog_Arguments("Source", GetAuthorizedWBSWriteSource())
                Exit Sub

            End If

            Application.EnableEvents = False
            Application.Undo

            WBS_ShowConsoleMessage vbExclamation, _
                "WBS.EVENTS.CALCULATED_EDIT"
            GoTo SafeExit

        End If
    End If

ContinueValidation:

    If rngCal Is Nothing Then
        Set rngToCheck = Intersect(Target, Union(rngWBS, rngPredWBS, rngTaskType, rngBaselineDuration))
    Else
        Set rngToCheck = Intersect(Target, Union(rngWBS, rngPredWBS, rngTaskType, rngCal, rngBaselineDuration))
    End If

    If Not rngSummaryDisplay Is Nothing Then
        If rngToCheck Is Nothing Then
            Set rngToCheck = Intersect(Target, rngSummaryDisplay)
        Else
            Set rngToCheck = Intersect(Target, Union(rngToCheck, rngSummaryDisplay))
        End If
    End If

    If rngToCheck Is Nothing Then GoTo SafeExit

    Set reWBS = CreateObject("VBScript.RegExp")
    reWBS.Pattern = "^\d+(\.\d+)*$"

    For Each cell In rngToCheck

        cellValue = Trim$(CStr(cell.value))

        If Not Intersect(cell, rngWBS) Is Nothing Then
            If cellValue <> "" Then
                If Not reWBS.Test(cellValue) Then

                    If IsMacroRunActive() Then
                        RequestMacroAbortKey _
                            "Handle_WBS_Change", _
                            "WBS.EVENTS.MACRO.WBS_FORMAT"
                        Exit Sub
                    End If

                    Application.EnableEvents = False
            Application.Undo

                    WBS_ShowConsoleMessage vbExclamation, _
                        "WBS.EVENTS.WBS_FORMAT"
                    GoTo SafeExit
                End If
            End If
        End If

        If Not Intersect(cell, rngPredWBS) Is Nothing Then
            If cellValue <> "" Then

                If Not IsMacroRunActive() Then
                    normalizedPredWBS = NormalizePredecessorsWBSLiveInput(cellValue)

                    If normalizedPredWBS <> cellValue Then
                        Application.EnableEvents = False
                        cell.NumberFormat = "@"
                        cell.value = normalizedPredWBS
                        Application.EnableEvents = True
                        cellValue = normalizedPredWBS
                    End If
                End If

                If cellValue = "" Or Not IsValidPredecessorsWBSInput(cellValue) Then

                    If IsMacroRunActive() Then
                        RequestMacroAbortKey _
                            "Handle_WBS_Change", _
                            "WBS.EVENTS.PREDECESSORS_FORMAT"
                        Exit Sub
                    End If

                    Application.EnableEvents = False
                    Application.Undo

                    WBS_ShowConsoleMessage vbExclamation, _
                        "WBS.EVENTS.PREDECESSORS_FORMAT"
                    GoTo SafeExit
                End If
            End If
        End If

        If Not Intersect(cell, rngTaskType) Is Nothing Then
            If cellValue <> "" Then
                If Not IsValidTaskTypeInput(cellValue) Then

                    If IsMacroRunActive() Then
                        RequestMacroAbortKey _
                            "Handle_WBS_Change", _
                            "WBS.EVENTS.MACRO.TASK_TYPE"
                        Exit Sub
                    End If

                    Application.EnableEvents = False
            Application.Undo

                    WBS_ShowConsoleMessage vbExclamation, _
                        "WBS.EVENTS.TASK_TYPE"
                    GoTo SafeExit
                End If
            End If
        End If

        If Not rngSummaryDisplay Is Nothing Then
            If Not Intersect(cell, rngSummaryDisplay) Is Nothing Then
                If cellValue <> "" Then
                    If Not IsValidSummaryDisplayInput(cellValue) Then

                        If IsMacroRunActive() Then
                            RequestMacroAbortKey _
                                "Handle_WBS_Change", _
                                "WBS.EVENTS.MACRO.SUMMARY_DISPLAY"
                            Exit Sub
                        End If

                        Application.EnableEvents = False
                        Application.Undo

                        WBS_ShowConsoleMessage vbExclamation, _
                            "WBS.EVENTS.SUMMARY_DISPLAY"
                        GoTo SafeExit
                    End If

                    If CStr(cell.value) <> UCase$(Trim$(CStr(cell.value))) Then
                        Application.EnableEvents = False
                        cell.NumberFormat = "@"
                        cell.value = UCase$(Trim$(CStr(cell.value)))
                        Application.EnableEvents = True
                        cellValue = CStr(cell.value)
                    End If
                End If
            End If
        End If


        If Not rngCal Is Nothing Then
            If Not Intersect(cell, rngCal) Is Nothing Then
                If cellValue <> "" Then
                    If Not IsValidCalendarType(cellValue) Then

                        If IsMacroRunActive() Then
                            RequestMacroAbortKey _
                                "Handle_WBS_Change", _
                                "WBS.EVENTS.MACRO.CALENDAR"
                            Exit Sub
                        End If

                        Application.EnableEvents = False
                        Application.Undo

                    WBS_ShowConsoleMessage vbExclamation, _
                        "WBS.EVENTS.CALENDAR"
                        GoTo SafeExit
                    End If
                End If
            End If
        End If
        If Not Intersect(cell, rngBaselineDuration) Is Nothing Then
            If cellValue <> "" Then
                If Not IsValidDurationInput(cellValue) Then

                    If IsMacroRunActive() Then
                        RequestMacroAbortKey _
                            "Handle_WBS_Change", _
                            "WBS.EVENTS.MACRO.DURATION"
                        Exit Sub
                    End If

                    Application.EnableEvents = False
                    Application.Undo

                WBS_ShowConsoleMessage vbExclamation, _
                    "WBS.EVENTS.DURATION"
                    GoTo SafeExit
                End If
            End If
        End If
    Next cell

SafeExit:
    If Not IsMacroAbortRequested() Then
        Application.EnableEvents = True
    End If

    If Err.Number <> 0 Then

        If IsMacroRunActive() Then
            RequestMacroAbortKey _
                "Handle_WBS_Change", _
                "WBS.EVENTS.MACRO.HANDLE_ERROR"
            Exit Sub
        End If

        WBS_ShowConsoleMessage vbCritical, _
            "WBS.EVENTS.HANDLE_ERROR"
    End If

End Sub

'------------------------------------------------------------------------------
' FR: Indique si WBSFormula Column Autofill Event est vrai pour le contexte courant.
' EN: Returns whether WBSFormula Column Autofill Event is true for the current context.
'------------------------------------------------------------------------------
Private Function IsWBSFormulaColumnAutofillEvent( _
    ByVal tbl As ListObject, _
    ByVal Target As Range, _
    ByVal rngLockedTouched As Range, _
    ByVal rngEditableCols As Range) As Boolean

    Dim col As ListColumn
    Dim hit As Range
    Dim touchedFormulaColumn As Boolean
    Dim columnKey As String

    On Error GoTo SafeExit

    If tbl Is Nothing Then Exit Function
    If Target Is Nothing Then Exit Function
    If rngLockedTouched Is Nothing Then Exit Function
    If Target.CountLarge <= 1 Then Exit Function
    If Not Intersect(Target, rngEditableCols) Is Nothing Then Exit Function

    For Each col In tbl.ListColumns
        Set hit = Nothing
        If Not col.DataBodyRange Is Nothing Then
            Set hit = Intersect(Target, col.DataBodyRange)
        End If

        If Not hit Is Nothing Then
            columnKey = WBSFormulaManagedColumnKey(tbl, col)
            If Len(columnKey) = 0 Then Exit Function
            If hit.Areas.Count <> 1 Then Exit Function
            If hit.Address(False, False) <> col.DataBodyRange.Address(False, False) Then Exit Function
            If Not WBSFormulaColumnHasExpectedFormula(col, columnKey) Then Exit Function
            touchedFormulaColumn = True
        End If
    Next col

    IsWBSFormulaColumnAutofillEvent = touchedFormulaColumn

SafeExit:
End Function

'------------------------------------------------------------------------------
' FR: Indique si WBSFormula Managed Column est vrai pour le contexte courant.
' EN: Returns whether WBSFormula Managed Column is true for the current context.
'------------------------------------------------------------------------------
Private Function WBSFormulaManagedColumnKey( _
    ByVal tbl As ListObject, _
    ByVal col As ListColumn) As String

    If col.Index = SchemaColumnIndex(tbl, VTS_TABLE_WBS, VTS_COL_BASELINE_FINISH) Then
        WBSFormulaManagedColumnKey = VTS_COL_BASELINE_FINISH
    ElseIf col.Index = SchemaColumnIndex(tbl, VTS_TABLE_WBS, VTS_COL_ACTUAL_DURATION) Then
        WBSFormulaManagedColumnKey = VTS_COL_ACTUAL_DURATION
    ElseIf col.Index = SchemaColumnIndex(tbl, VTS_TABLE_WBS, VTS_COL_CALCULATED_DURATION) Then
        WBSFormulaManagedColumnKey = VTS_COL_CALCULATED_DURATION
    End If

End Function

'------------------------------------------------------------------------------
' FR: Retourne la valeur WBS Formula Column Has Expected Formula sans modifier les donnees d'entree.
' EN: Returns the WBS Formula Column Has Expected Formula value without mutating input data.
'------------------------------------------------------------------------------

Private Function WBSFormulaColumnHasExpectedFormula( _
    ByVal col As ListColumn, _
    ByVal columnKey As String) As Boolean

    Dim expectedFormula As String
    Dim currentFormula As String

    On Error GoTo SafeExit

    If col Is Nothing Then Exit Function
    If col.DataBodyRange Is Nothing Then Exit Function
    If col.DataBodyRange.Cells.CountLarge = 0 Then Exit Function

    expectedFormula = ExpectedWBSFormulaInvariant(columnKey)

    currentFormula = CStr(col.DataBodyRange.Cells(1, 1).Formula)
    WBSFormulaColumnHasExpectedFormula = (StrComp(currentFormula, expectedFormula, vbTextCompare) = 0)

SafeExit:
End Function

'------------------------------------------------------------------------------
' FR: Retourne la formule WBS attendue dans la syntaxe anglaise invariante d'Excel.
' EN: Returns the expected WBS formula in Excel's invariant English syntax.
'------------------------------------------------------------------------------

Private Function ExpectedWBSFormulaInvariant(ByVal columnKey As String) As String

    ExpectedWBSFormulaInvariant = WBSFormulaWriter_ExpectedFormula(columnKey)

End Function
'------------------------------------------------------------------------------
' FR: Resout Project et accepte l'ancien nom Package jusqu'a la migration d'onboarding.
' EN: Resolves Project and accepts the legacy Package name until onboarding migration.
'------------------------------------------------------------------------------
Private Function WBSProjectListColumn(ByVal tbl As ListObject) As ListColumn

    If WBSHasSchemaColumn(tbl, VTS_COL_PROJECT) Then
        Set WBSProjectListColumn = SchemaListColumn(tbl, VTS_TABLE_WBS, VTS_COL_PROJECT)
    ElseIf WBSHasPhysicalColumn(tbl, "Package") Then
        Set WBSProjectListColumn = tbl.ListColumns("Package")
    Else
        Err.Raise vbObjectError + 2340, "WBSProjectListColumn", _
            PlanningMessageText_Format("WBS.ERROR.MISSING_PROJECT_COLUMN")
    End If

End Function

'------------------------------------------------------------------------------
' FR: Retourne l'union des colonnes WBS que l'utilisateur peut modifier.
' EN: Returns the union of WBS columns that the user may edit.
'------------------------------------------------------------------------------
Private Function GetWBSUserEditableRange(ByVal tbl As ListObject) As Range

    Dim rng As Range

    Set rng = Union( _
        SchemaListColumn(tbl, VTS_TABLE_WBS, VTS_COL_WBS).DataBodyRange, _
        SchemaListColumn(tbl, VTS_TABLE_WBS, VTS_COL_TASK_NAME).DataBodyRange, _
        SchemaListColumn(tbl, VTS_TABLE_WBS, VTS_COL_TASK_DESCRIPTION).DataBodyRange, _
        SchemaListColumn(tbl, VTS_TABLE_WBS, VTS_COL_DISCIPLINE).DataBodyRange, _
        SchemaListColumn(tbl, VTS_TABLE_WBS, VTS_COL_SUPPLIER).DataBodyRange, _
        WBSProjectListColumn(tbl).DataBodyRange, _
        SchemaListColumn(tbl, VTS_TABLE_WBS, VTS_COL_TASK_TYPE).DataBodyRange)

    If WBSHasSchemaColumn(tbl, VTS_COL_CAL) Then
        Set rng = Union(rng, SchemaListColumn(tbl, VTS_TABLE_WBS, VTS_COL_CAL).DataBodyRange)
    End If

    If WBSHasSchemaColumn(tbl, VTS_COL_S) Then
        Set rng = Union(rng, SchemaListColumn(tbl, VTS_TABLE_WBS, VTS_COL_S).DataBodyRange)
    End If
    Set GetWBSUserEditableRange = Union( _
        rng, _
        SchemaListColumn(tbl, VTS_TABLE_WBS, VTS_COL_PREDECESSORS_WBS).DataBodyRange, _
        SchemaListColumn(tbl, VTS_TABLE_WBS, VTS_COL_WEIGHT_PERCENT).DataBodyRange, _
        SchemaListColumn(tbl, VTS_TABLE_WBS, VTS_COL_PROGRESS_PERCENT).DataBodyRange, _
        SchemaListColumn(tbl, VTS_TABLE_WBS, VTS_COL_COMMENTS).DataBodyRange, _
        SchemaListColumn(tbl, VTS_TABLE_WBS, VTS_COL_BASELINE_START).DataBodyRange, _
        SchemaListColumn(tbl, VTS_TABLE_WBS, VTS_COL_BASELINE_DURATION).DataBodyRange, _
        SchemaListColumn(tbl, VTS_TABLE_WBS, VTS_COL_ACTUAL_START).DataBodyRange, _
        SchemaListColumn(tbl, VTS_TABLE_WBS, VTS_COL_ACTUAL_FINISH).DataBodyRange, _
        SchemaListColumn(tbl, VTS_TABLE_WBS, VTS_COL_FORECAST_START).DataBodyRange, _
        SchemaListColumn(tbl, VTS_TABLE_WBS, VTS_COL_FORECAST_FINISH).DataBodyRange _
    )

End Function
'------------------------------------------------------------------------------
' FR: Retourne la reference WBS Has Column sans modifier les donnees d'entree.
' EN: Returns the WBS Has Column reference without mutating input data.
'------------------------------------------------------------------------------

Private Function WBSHasSchemaColumn(ByVal tbl As ListObject, ByVal columnKey As String) As Boolean

    WBSHasSchemaColumn = WBSHasPhysicalColumn( _
        tbl, SchemaCurrentColumnTitle(VTS_TABLE_WBS, columnKey))

End Function

Private Function WBSHasPhysicalColumn(ByVal tbl As ListObject, ByVal columnName As String) As Boolean

    Dim col As ListColumn

    On Error Resume Next
    Set col = tbl.ListColumns(columnName)
    On Error GoTo 0

    WBSHasPhysicalColumn = Not col Is Nothing

End Function
'------------------------------------------------------------------------------
' FR: Retourne WBSLocked Range depuis le contexte WBS events.
' EN: Returns WBSLocked Range from the WBS events context.
'------------------------------------------------------------------------------
Private Function GetWBSLockedRange(ByVal tbl As ListObject) As Range

    Set GetWBSLockedRange = Union( _
        SchemaListColumn(tbl, VTS_TABLE_WBS, VTS_COL_BASELINE_FINISH).DataBodyRange, _
        SchemaListColumn(tbl, VTS_TABLE_WBS, VTS_COL_ACTUAL_DURATION).DataBodyRange, _
        SchemaListColumn(tbl, VTS_TABLE_WBS, VTS_COL_CALCULATED_START).DataBodyRange, _
        SchemaListColumn(tbl, VTS_TABLE_WBS, VTS_COL_CALCULATED_FINISH).DataBodyRange, _
        SchemaListColumn(tbl, VTS_TABLE_WBS, VTS_COL_CALCULATED_DURATION).DataBodyRange, _
        SchemaListColumn(tbl, VTS_TABLE_WBS, VTS_COL_START_VARIANCE).DataBodyRange, _
        SchemaListColumn(tbl, VTS_TABLE_WBS, VTS_COL_FINISH_VARIANCE).DataBodyRange, _
        SchemaListColumn(tbl, VTS_TABLE_WBS, VTS_COL_DURATION_VARIANCE).DataBodyRange, _
        SchemaListColumn(tbl, VTS_TABLE_WBS, VTS_COL_DRIVING_LOGIC).DataBodyRange, _
        SchemaListColumn(tbl, VTS_TABLE_WBS, VTS_COL_CRITICAL_PATH).DataBodyRange, _
        SchemaListColumn(tbl, VTS_TABLE_WBS, VTS_COL_LONGEST_PATH).DataBodyRange, _
        SchemaListColumn(tbl, VTS_TABLE_WBS, VTS_COL_CRITICAL_PATH_REX).DataBodyRange, _
        SchemaListColumn(tbl, VTS_TABLE_WBS, VTS_COL_TOTAL_FLOAT).DataBodyRange, _
        SchemaListColumn(tbl, VTS_TABLE_WBS, VTS_COL_FREE_FLOAT).DataBodyRange, _
        SchemaListColumn(tbl, VTS_TABLE_WBS, VTS_COL_TOTAL_FLOAT_REX).DataBodyRange, _
        SchemaListColumn(tbl, VTS_TABLE_WBS, VTS_COL_FREE_FLOAT_REX).DataBodyRange _
    )

End Function

'------------------------------------------------------------------------------
' FR: Normalise Predecessors WBSLive Input dans un format exploitable.
' EN: Normalizes Predecessors WBSLive Input into a usable format.
'------------------------------------------------------------------------------
Private Function NormalizePredecessorsWBSLiveInput(ByVal inputText As String) As String

    Dim cleaned As String
    Dim tokens As Variant
    Dim i As Long
    Dim tokenText As String
    Dim result As String

    cleaned = Trim$(CStr(inputText))
    cleaned = Replace$(cleaned, " ", "")
    cleaned = Replace$(cleaned, vbTab, "")
    cleaned = Replace$(cleaned, Chr$(160), "")

    If cleaned = "" Then
        NormalizePredecessorsWBSLiveInput = ""
        Exit Function
    End If

    tokens = Split(cleaned, ";")

    For i = LBound(tokens) To UBound(tokens)
        tokenText = Trim$(CStr(tokens(i)))

        If tokenText <> "" Then
            If result <> "" Then result = result & ";"
            result = result & tokenText
        End If
    Next i

    NormalizePredecessorsWBSLiveInput = result

End Function
'------------------------------------------------------------------------------
' FR: Construit la valeur WBS Build Predecessors Format Message FR a partir des donnees fournies par l'appelant.
' EN: Builds the WBS Build Predecessors Format Message FR value from data supplied by the caller.
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' FR: Indique si Valid Predecessors WBSInput est vrai pour le contexte courant.
' EN: Returns whether Valid Predecessors WBSInput is true for the current context.
'------------------------------------------------------------------------------
Private Function IsValidPredecessorsWBSInput(ByVal inputText As String) As Boolean

    Dim tokens As Variant
    Dim i As Long
    Dim tokenText As String

    Dim predWbs As String
    Dim linkType As String
    Dim lagVal As Long
    Dim rawToken As String
    Dim errText As String

    inputText = Trim$(CStr(inputText))

    If inputText = "" Then
        IsValidPredecessorsWBSInput = True
        Exit Function
    End If

    If InStr(1, inputText, " ", vbBinaryCompare) > 0 Then Exit Function

    tokens = Split(inputText, ";")

    For i = LBound(tokens) To UBound(tokens)

        tokenText = Trim$(CStr(tokens(i)))

        If tokenText = "" Then Exit Function

        If Not ParsePredecessorToken( _
            tokenText, _
            predWbs, _
            linkType, _
            lagVal, _
            rawToken, _
            errText) Then
            Exit Function
        End If

    Next i

    IsValidPredecessorsWBSInput = True

End Function

'------------------------------------------------------------------------------
' FR: Indique si Valid Predecessor Token est vrai pour le contexte courant.
' EN: Returns whether Valid Predecessor Token is true for the current context.
'------------------------------------------------------------------------------
Private Function IsValidPredecessorToken(ByVal tokenText As String) As Boolean

    Dim predWbs As String
    Dim linkType As String
    Dim lagVal As Long
    Dim rawToken As String
    Dim errText As String

    IsValidPredecessorToken = ParsePredecessorToken( _
        tokenText, _
        predWbs, _
        linkType, _
        lagVal, _
        rawToken, _
        errText)

End Function

'------------------------------------------------------------------------------
' FR: Indique si Valid Task Type Input est vrai pour le contexte courant.
' EN: Returns whether Valid Task Type Input is true for the current context.
'------------------------------------------------------------------------------
Private Function IsValidTaskTypeInput(ByVal inputText As String) As Boolean

    Select Case UCase$(Trim$(CStr(inputText)))

        Case "", "TASK", "MILESTONE", "LEVEL OF EFFORT"
            IsValidTaskTypeInput = True

        Case Else
            IsValidTaskTypeInput = False

    End Select

End Function

'------------------------------------------------------------------------------
' FR: Indique si Valid Summary Display Input est vrai pour le contexte courant.
' EN: Returns whether Valid Summary Display Input is true for the current context.
'------------------------------------------------------------------------------
Private Function IsValidSummaryDisplayInput(ByVal inputText As String) As Boolean

    Select Case UCase$(Trim$(CStr(inputText)))

        Case "", "Y", "N"
            IsValidSummaryDisplayInput = True

        Case Else
            IsValidSummaryDisplayInput = False

    End Select

End Function
'------------------------------------------------------------------------------
' FR: Indique si Valid Duration Input est vrai pour le contexte courant.
' EN: Returns whether Valid Duration Input is true for the current context.
'------------------------------------------------------------------------------
Private Function IsValidDurationInput(ByVal inputText As String) As Boolean

    Dim durationValue As Double

    inputText = Trim$(CStr(inputText))

    If inputText = "" Then
        IsValidDurationInput = True
        Exit Function
    End If

    If Not IsNumeric(inputText) Then Exit Function

    durationValue = CDbl(inputText)
    IsValidDurationInput = (durationValue > 0)

End Function

'------------------------------------------------------------------------------
' FR: Projette la collection WBS Show Console Message vers l'interface autorisee par la politique runtime.
' EN: Projects the WBS Show Console Message collection to the UI allowed by runtime policy.
'------------------------------------------------------------------------------

Private Sub WBS_ShowConsoleMessage( _
    ByVal boxStyle As VbMsgBoxStyle, _
    ByVal messageKey As String)

    Dim consoleMessages As Collection
    Dim msgType As String
    Dim msg As String

    msgType = WBS_MessageTypeFromMsgBoxStyle(boxStyle)

    msg = PlanningMessageText_Format(messageKey, Nothing, Nothing)

    Set consoleMessages = New Collection
    CalcBridge_AddConsoleMessage consoleMessages, msgType, msg
    CalcBridge_ShowPlanningConsole consoleMessages

End Sub

'------------------------------------------------------------------------------
' FR: Retourne la valeur WBS Message Type From Msg Box Style sans modifier les donnees d'entree.
' EN: Returns the WBS Message Type From Msg Box Style value without mutating input data.
'------------------------------------------------------------------------------

Private Function WBS_MessageTypeFromMsgBoxStyle(ByVal boxStyle As VbMsgBoxStyle) As String

    If (boxStyle And vbCritical) = vbCritical Then
        WBS_MessageTypeFromMsgBoxStyle = "STOP"
    ElseIf (boxStyle And vbExclamation) = vbExclamation Then
        WBS_MessageTypeFromMsgBoxStyle = "WARNING"
    Else
        WBS_MessageTypeFromMsgBoxStyle = "INFO"
    End If

End Function



