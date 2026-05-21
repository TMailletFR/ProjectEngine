Attribute VB_Name = "mod_WBSEvents"
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

Public Sub Handle_WBS_Change(ByVal ws As Worksheet, ByVal Target As Range)

    Dim tbl As ListObject

    Dim rngWBS As Range
    Dim rngPredWBS As Range
    Dim rngTaskType As Range
    Dim rngBaselineDuration As Range

    Dim rngLockedCols As Range
    Dim rngEditableCols As Range
    Dim rngToCheck As Range
    Dim rngLockedTouched As Range

    Dim cell As Range
    Dim cellValue As String

    Dim reWBS As Object

    Dim col As ListColumn
    Dim authorizedHit As Boolean

    On Error GoTo SafeExit

    Set tbl = ws.ListObjects("tbl_WBS")
    If tbl.DataBodyRange Is Nothing Then Exit Sub

    Set rngWBS = tbl.ListColumns("WBS").DataBodyRange
    Set rngPredWBS = tbl.ListColumns("Predecessors WBS").DataBodyRange
    Set rngTaskType = tbl.ListColumns("Task Type").DataBodyRange
    Set rngBaselineDuration = tbl.ListColumns("Baseline Duration").DataBodyRange

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

                RequestMacroAbort _
                    "Handle_WBS_Change", _
                    "Écriture moteur interdite dans une colonne calculée de WBS." & vbCrLf & _
                    "-> source : " & GetAuthorizedWBSWriteSource() & vbCrLf & _
                    "-> le macro-run est arrêté pour éviter une corruption de données.", _
                    "Engine write not allowed in a calculated WBS column." & vbCrLf & _
                    "-> source: " & GetAuthorizedWBSWriteSource() & vbCrLf & _
                    "-> the macro run was stopped to prevent data corruption."
                Exit Sub

            End If

            Application.EnableEvents = False
            Application.Undo

            WBS_ShowConsoleMessage vbExclamation, _
                "Modification interdite dans une colonne calculée (grise)." & vbCrLf & _
                "-> modifier uniquement les colonnes d'entrée (bleues).", _
                "Manual edit not allowed in calculated column (gray)." & vbCrLf & _
                "-> edit input columns only (blue)."
            GoTo SafeExit

        End If
    End If

ContinueValidation:

    Set rngToCheck = Intersect(Target, Union(rngWBS, rngPredWBS, rngTaskType, rngBaselineDuration))
    If rngToCheck Is Nothing Then GoTo SafeExit

    Set reWBS = CreateObject("VBScript.RegExp")
    reWBS.Pattern = "^\d+(\.\d+)*$"

    For Each cell In rngToCheck

        cellValue = Trim$(CStr(cell.value))

        If Not Intersect(cell, rngWBS) Is Nothing Then
            If cellValue <> "" Then
                If Not reWBS.Test(cellValue) Then

                    If IsMacroRunActive() Then
                        RequestMacroAbort _
                            "Handle_WBS_Change", _
                            "Format invalide détecté dans WBS pendant l'exécution d'un macro." & vbCrLf & _
                            "-> format attendu : 1 | 1.2 | 1.2.3", _
                            "Invalid WBS format detected during macro execution." & vbCrLf & _
                            "-> expected format: 1 | 1.2 | 1.2.3"
                        Exit Sub
                    End If

                    Application.EnableEvents = False
            Application.Undo

                    WBS_ShowConsoleMessage vbExclamation, _
                        "Format invalide dans WBS." & vbCrLf & _
                        "-> format attendu : 1 | 1.2 | 1.2.3", _
                        "Invalid format in WBS." & vbCrLf & _
                        "-> expected format: 1 | 1.2 | 1.2.3"
                    GoTo SafeExit
                End If
            End If
        End If

        If Not Intersect(cell, rngPredWBS) Is Nothing Then
            If cellValue <> "" Then
                If Not IsValidPredecessorsWBSInput(cellValue) Then

                    If IsMacroRunActive() Then
                        RequestMacroAbort _
                            "Handle_WBS_Change", _
                            "Format invalide détecté dans Predecessors WBS pendant l'exécution d'un macro." & vbCrLf & _
                            "-> formats autorisés : 1 | 1+3 | 1-2 | 1FS+3 | 1SS-2 | 1FF+4 | 1;2SS+3;4FF-2", _
                            "Invalid Predecessors WBS format detected during macro execution." & vbCrLf & _
                            "-> allowed formats: 1 | 1+3 | 1-2 | 1FS+3 | 1SS-2 | 1FF+4 | 1;2SS+3;4FF-2"
                        Exit Sub
                    End If

                    Application.EnableEvents = False
            Application.Undo

                    WBS_ShowConsoleMessage vbExclamation, _
                        "Format invalide dans Predecessors WBS." & vbCrLf & _
                        "-> formats autorisés : 1 | 1+3 | 1-2 | 1FS+3 | 1SS-2 | 1FF+4 | 1;2SS+3;4FF-2" & vbCrLf & _
                        "-> espaces interdits / token vide interdit", _
                        "Invalid format in Predecessors WBS." & vbCrLf & _
                        "-> allowed formats: 1 | 1+3 | 1-2 | 1FS+3 | 1SS-2 | 1FF+4 | 1;2SS+3;4FF-2" & vbCrLf & _
                        "-> spaces forbidden / empty token forbidden"
                    GoTo SafeExit
                End If
            End If
        End If

        If Not Intersect(cell, rngTaskType) Is Nothing Then
            If cellValue <> "" Then
                If Not IsValidTaskTypeInput(cellValue) Then

                    If IsMacroRunActive() Then
                        RequestMacroAbort _
                            "Handle_WBS_Change", _
                            "Task Type invalide détecté pendant l'exécution d'un macro." & vbCrLf & _
                            "-> valeurs autorisées : Task | Milestone | Level of Effort", _
                            "Invalid Task Type detected during macro execution." & vbCrLf & _
                            "-> allowed values: Task | Milestone | Level of Effort"
                        Exit Sub
                    End If

                    Application.EnableEvents = False
            Application.Undo

                    WBS_ShowConsoleMessage vbExclamation, _
                        "Task Type invalide." & vbCrLf & _
                        "-> valeurs autorisées : Task | Milestone | Level of Effort", _
                        "Invalid Task Type." & vbCrLf & _
                        "-> allowed values: Task | Milestone | Level of Effort"
                    GoTo SafeExit
                End If
            End If
        End If

        If Not Intersect(cell, rngBaselineDuration) Is Nothing Then
            If cellValue <> "" Then
                If Not IsValidDurationInput(cellValue) Then

                    If IsMacroRunActive() Then
                        RequestMacroAbort _
                            "Handle_WBS_Change", _
                            "Duree invalide detectee pendant l'execution d'un macro." & vbCrLf & _
                            "-> saisir une duree numerique strictement positive.", _
                            "Invalid duration detected during macro execution." & vbCrLf & _
                            "-> enter a strictly positive numeric duration."
                        Exit Sub
                    End If

                    Application.EnableEvents = False
                    Application.Undo

                    WBS_ShowConsoleMessage vbExclamation, _
                        "Duree invalide." & vbCrLf & _
                            "-> saisir une duree numerique strictement positive.", _
                        "Invalid duration." & vbCrLf & _
                        "-> enter a strictly positive numeric duration."
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
            RequestMacroAbort _
                "Handle_WBS_Change", _
                "Erreur VBA dans Handle_WBS_Change." & vbCrLf & _
                "-> le macro-run est arrêté pour éviter une cascade d'erreurs.", _
                "VBA error in Handle_WBS_Change." & vbCrLf & _
                "-> the macro run was stopped to avoid an error cascade."
            Exit Sub
        End If

        WBS_ShowConsoleMessage vbCritical, _
            "Erreur VBA dans Handle_WBS_Change" & vbCrLf & _
            "-> vérifier le dernier bloc modifié dans mod_WBSEvents", _
            "VBA error in Handle_WBS_Change" & vbCrLf & _
            "-> check the last edited block in mod_WBSEvents"
    End If

End Sub

Private Function IsWBSFormulaColumnAutofillEvent( _
    ByVal tbl As ListObject, _
    ByVal Target As Range, _
    ByVal rngLockedTouched As Range, _
    ByVal rngEditableCols As Range) As Boolean

    Dim col As ListColumn
    Dim hit As Range
    Dim touchedFormulaColumn As Boolean

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
            If Not IsWBSFormulaManagedColumn(col.Name) Then Exit Function
            If hit.Areas.Count <> 1 Then Exit Function
            If hit.Address(False, False) <> col.DataBodyRange.Address(False, False) Then Exit Function
            If Not WBSFormulaColumnHasExpectedFormula(col) Then Exit Function
            touchedFormulaColumn = True
        End If
    Next col

    IsWBSFormulaColumnAutofillEvent = touchedFormulaColumn

SafeExit:
End Function

Private Function IsWBSFormulaManagedColumn(ByVal columnName As String) As Boolean

    Select Case CStr(columnName)
        Case "Baseline Finish", "Actual Duration", "Calculated Duration"
            IsWBSFormulaManagedColumn = True
    End Select

End Function

Private Function WBSFormulaColumnHasExpectedFormula(ByVal col As ListColumn) As Boolean

    Dim expectedFormula As String
    Dim currentFormula As String

    On Error GoTo SafeExit

    If col Is Nothing Then Exit Function
    If col.DataBodyRange Is Nothing Then Exit Function
    If col.DataBodyRange.Cells.CountLarge = 0 Then Exit Function

    expectedFormula = ExpectedWBSFormulaLocal(col.Name)
    If Len(expectedFormula) = 0 Then Exit Function

    currentFormula = CStr(col.DataBodyRange.Cells(1, 1).FormulaLocal)
    WBSFormulaColumnHasExpectedFormula = (StrComp(currentFormula, expectedFormula, vbTextCompare) = 0)

SafeExit:
End Function

Private Function ExpectedWBSFormulaLocal(ByVal columnName As String) As String

    Select Case CStr(columnName)
        Case "Baseline Finish"
            ExpectedWBSFormulaLocal = "=SI(OU([@[Baseline Start]]="""";[@[Baseline Duration]]="""");"""";[@[Baseline Start]]+[@[Baseline Duration]]-1)"
        Case "Actual Duration"
            ExpectedWBSFormulaLocal = "=SI(OU([@[Actual Start]]="""";[@[Actual Finish]]="""");"""";[@[Actual Finish]]-[@[Actual Start]]+1)"
        Case "Calculated Duration"
            ExpectedWBSFormulaLocal = "=SI(OU([@[Calculated Start]]="""";[@[Calculated Finish]]="""");"""";[@[Calculated Finish]]-[@[Calculated Start]]+1)"
    End Select

End Function
Private Function GetWBSUserEditableRange(ByVal tbl As ListObject) As Range

    Set GetWBSUserEditableRange = Union( _
        tbl.ListColumns("ID").DataBodyRange, _
        tbl.ListColumns("WBS").DataBodyRange, _
        tbl.ListColumns("Task Name").DataBodyRange, _
        tbl.ListColumns("Task Description").DataBodyRange, _
        tbl.ListColumns("Discipline").DataBodyRange, _
        tbl.ListColumns("Supplier").DataBodyRange, _
        tbl.ListColumns("Package").DataBodyRange, _
        tbl.ListColumns("Task Type").DataBodyRange, _
        tbl.ListColumns("Predecessors WBS").DataBodyRange, _
        tbl.ListColumns("Weight (%)").DataBodyRange, _
        tbl.ListColumns("% Progress").DataBodyRange, _
        tbl.ListColumns("Comments").DataBodyRange, _
        tbl.ListColumns("Baseline Start").DataBodyRange, _
        tbl.ListColumns("Baseline Duration").DataBodyRange, _
        tbl.ListColumns("Actual Start").DataBodyRange, _
        tbl.ListColumns("Actual Finish").DataBodyRange, _
        tbl.ListColumns("Forecast Start").DataBodyRange, _
        tbl.ListColumns("Forecast Finish").DataBodyRange _
    )

End Function

Private Function GetWBSLockedRange(ByVal tbl As ListObject) As Range

    Set GetWBSLockedRange = Union( _
        tbl.ListColumns("Baseline Finish").DataBodyRange, _
        tbl.ListColumns("Actual Duration").DataBodyRange, _
        tbl.ListColumns("Calculated Start").DataBodyRange, _
        tbl.ListColumns("Calculated Finish").DataBodyRange, _
        tbl.ListColumns("Calculated Duration").DataBodyRange, _
        tbl.ListColumns("Start Variance").DataBodyRange, _
        tbl.ListColumns("Finish Variance").DataBodyRange, _
        tbl.ListColumns("Duration Variance").DataBodyRange, _
        tbl.ListColumns("Driving Logic").DataBodyRange, _
        tbl.ListColumns("Critical Path").DataBodyRange, _
        tbl.ListColumns("Longest Path").DataBodyRange, _
        tbl.ListColumns("Critical Path REX").DataBodyRange, _
        tbl.ListColumns("Total Float").DataBodyRange, _
        tbl.ListColumns("Free Float").DataBodyRange, _
        tbl.ListColumns("Total Float REX").DataBodyRange, _
        tbl.ListColumns("Free Float REX").DataBodyRange _
    )

End Function

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

Private Function IsValidTaskTypeInput(ByVal inputText As String) As Boolean

    Select Case UCase$(Trim$(CStr(inputText)))

        Case "", "TASK", "MILESTONE", "LEVEL OF EFFORT"
            IsValidTaskTypeInput = True

        Case Else
            IsValidTaskTypeInput = False

    End Select

End Function

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

Private Sub WBS_ShowConsoleMessage( _
    ByVal boxStyle As VbMsgBoxStyle, _
    ByVal frText As String, _
    ByVal enText As String)

    Dim consoleMessages As Collection
    Dim msgType As String
    Dim msg As String

    msgType = WBS_MessageTypeFromMsgBoxStyle(boxStyle)

    msg = _
        "FR:" & vbCrLf & _
        frText & vbCrLf & vbCrLf & _
        "EN:" & vbCrLf & _
        enText

    Set consoleMessages = New Collection
    CalcBridge_AddConsoleMessage consoleMessages, msgType, msg
    CalcBridge_ShowPlanningConsole consoleMessages

End Sub

Private Function WBS_MessageTypeFromMsgBoxStyle(ByVal boxStyle As VbMsgBoxStyle) As String

    If (boxStyle And vbCritical) = vbCritical Then
        WBS_MessageTypeFromMsgBoxStyle = "STOP"
    ElseIf (boxStyle And vbExclamation) = vbExclamation Then
        WBS_MessageTypeFromMsgBoxStyle = "WARNING"
    Else
        WBS_MessageTypeFromMsgBoxStyle = "INFO"
    End If

End Function



