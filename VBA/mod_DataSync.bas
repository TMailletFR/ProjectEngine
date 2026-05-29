Attribute VB_Name = "mod_DataSync"
Option Explicit

Private Const LOGIC_LINKS_TABLE_NAME As String = "tbl_LOGIC_LINKS"
Private Const LOGIC_LINKS_FIRST_CELL As String = "Z1"

Sub Sync_WBS_To_CALC(Optional ByVal preserveCalcOutputs As Boolean = False)

    Dim wsWBS As Worksheet
    Dim wsCalc As Worksheet
    Dim tblWBS As ListObject
    Dim tblCalc As ListObject

    Dim mapCalc As Object
    Dim mapWBS As Object
    Dim wbsIdRows As Object
    Dim wbsToId As Object
    Dim summaryWbsByWbs As Object

    Dim colsToCopy As Variant
    Dim colsToClear As Variant

    Dim i As Long
    Dim r As Long
    Dim maxId As Long
    Dim currentCalcRows As Long
    Dim targetRows As Long

    Dim idValue As Variant
    Dim idKey As String
    Dim wbsValue As String
    Dim wbsRowIndex As Long

    Dim predWbs As String
    Dim predRaw As String
    Dim taskTypeVal As String
    Dim calVal As String

    Dim consoleMessages As Collection

    On Error GoTo SafeExit

    Set consoleMessages = New Collection

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    Set wsWBS = ThisWorkbook.Worksheets("WBS")
    Set tblWBS = wsWBS.ListObjects("tbl_WBS")

    EnsureWBSTaskTypeInputSetup tblWBS
    EnsureWBSCalendarInputSetup tblWBS

    Ensure_Calc_Infrastructure

    Set wsCalc = ThisWorkbook.Worksheets("CALC")
    Set tblCalc = wsCalc.ListObjects("tbl_CALC")

    EnsureTaskTypeColumnExists tblWBS, tblCalc
    EnsureCalendarColumnExists tblWBS, tblCalc
    EnsureDeadlineOutputColumnsExist tblWBS, tblCalc
    EnsureLongestPathOutputColumnsExist tblWBS, tblCalc

    Set mapCalc = CreateObject("Scripting.Dictionary")
    Set mapWBS = CreateObject("Scripting.Dictionary")
    Set wbsIdRows = CreateObject("Scripting.Dictionary")
    Set wbsToId = CreateObject("Scripting.Dictionary")

    colsToCopy = Array( _
        "WBS", _
        "Task Name", _
        "Task Type", _
        "Cal", _
        "Predecessors WBS", _
        "Baseline Start", _
        "Baseline Duration", _
        "Baseline Finish", _
        "Actual Start", _
        "Actual Finish", _
        "Actual Duration", _
        "Forecast Start", _
        "Forecast Finish" _
    )

    colsToClear = Array( _
        "ParentID", _
        "IsSummary", _
        "Calculated Start", _
        "Calculated Finish", _
        "Calculated Duration", _
        "Driving Logic", _
        "Error flag", _
        "ErrorMsg", _
        "Critical Path", _
        "Longest Path", _
        "Total Float", _
        "Free Float", _
        "Critical Path REX", _
        "Total Float REX", _
        "Free Float REX", _
        "Deadline Float" _
    )

    For i = 1 To tblWBS.ListColumns.Count
        mapWBS(tblWBS.ListColumns(i).Name) = i
    Next i

    For i = 1 To tblCalc.ListColumns.Count
        mapCalc(tblCalc.ListColumns(i).Name) = i
    Next i

    If Not mapWBS.Exists("ID") Then
        DataSync_AddConsoleMessage consoleMessages, "STOP", _
            "La colonne ID est introuvable dans tbl_WBS.", _
            "Column ID was not found in tbl_WBS."
        GoTo SafeExit
    End If

    If Not mapWBS.Exists("WBS") Then
        DataSync_AddConsoleMessage consoleMessages, "STOP", _
            "La colonne WBS est introuvable dans tbl_WBS.", _
            "Column WBS was not found in tbl_WBS."
        GoTo SafeExit
    End If

    If Not mapWBS.Exists("Task Type") Then
        DataSync_AddConsoleMessage consoleMessages, "STOP", _
            "La colonne Task Type est introuvable dans tbl_WBS.", _
            "Column Task Type was not found in tbl_WBS."
        GoTo SafeExit
    End If

    If Not mapWBS.Exists("Predecessors WBS") Then
        DataSync_AddConsoleMessage consoleMessages, "STOP", _
            "La colonne Predecessors WBS est introuvable dans tbl_WBS.", _
            "Column Predecessors WBS was not found in WBS."
        GoTo SafeExit
    End If

    If Not mapCalc.Exists("ID") Then
        DataSync_AddConsoleMessage consoleMessages, "STOP", _
            "La colonne ID est introuvable dans tbl_CALC.", _
            "Column ID was not found in tbl_CALC."
        GoTo SafeExit
    End If

    If Not mapCalc.Exists("WBS") Then
        DataSync_AddConsoleMessage consoleMessages, "STOP", _
            "La colonne WBS est introuvable dans tbl_CALC.", _
            "Column WBS was not found in tbl_CALC."
        GoTo SafeExit
    End If

    If Not mapCalc.Exists("Task Type") Then
        DataSync_AddConsoleMessage consoleMessages, "STOP", _
            "La colonne Task Type est introuvable dans tbl_CALC.", _
            "Column Task Type was not found in tbl_CALC."
        GoTo SafeExit
    End If

    If Not mapCalc.Exists("Predecessors WBS") Then
        DataSync_AddConsoleMessage consoleMessages, "STOP", _
            "La colonne Predecessors WBS est introuvable dans tbl_CALC.", _
            "Column Predecessors WBS was not found in tbl_CALC."
        GoTo SafeExit
    End If

    maxId = 0

    If Not tblWBS.DataBodyRange Is Nothing Then
        For r = 1 To tblWBS.ListRows.Count

            idValue = tblWBS.DataBodyRange.Cells(r, mapWBS("ID")).value
            wbsValue = Trim$(CStr(tblWBS.DataBodyRange.Cells(r, mapWBS("WBS")).value))
            wbsValue = Replace(wbsValue, ",", ".")

            If Trim$(CStr(idValue)) <> "" Then

                If Not IsNumeric(idValue) Then
                    DataSync_AddConsoleMessage consoleMessages, "STOP", _
                        "ID non numérique détecté dans WBS : " & CStr(idValue), _
                        "Non-numeric ID detected in WBS: " & CStr(idValue)
                    GoTo SafeExit
                End If

                If CLng(idValue) < 1 Then
                    DataSync_AddConsoleMessage consoleMessages, "STOP", _
                        "ID invalide dans WBS (doit être >= 1) : " & CStr(idValue), _
                        "Invalid ID in WBS (must be >= 1): " & CStr(idValue)
                    GoTo SafeExit
                End If

                idKey = CStr(CLng(idValue))
                wbsIdRows(idKey) = r

                If CLng(idValue) > maxId Then
                    maxId = CLng(idValue)
                End If

                If wbsValue <> "" Then
                    If wbsToId.Exists(wbsValue) Then
                        DataSync_AddConsoleMessage consoleMessages, "STOP", _
                            "WBS dupliqué détecté dans WBS : " & wbsValue, _
                            "Duplicate WBS detected in WBS: " & wbsValue
                        GoTo SafeExit
                    Else
                        wbsToId(wbsValue) = CLng(idValue)
                    End If
                End If

            End If

        Next r
    End If

    Set summaryWbsByWbs = BuildWBSSummaryWbsLookup(tblWBS, mapWBS)
    NormalizeWBSCalendarValues tblWBS, mapWBS, summaryWbsByWbs

    If tblCalc.DataBodyRange Is Nothing Then
        currentCalcRows = 0
    Else
        currentCalcRows = tblCalc.ListRows.Count
    End If

    targetRows = maxId

    If targetRows = 0 Then
        Do While tblCalc.ListRows.Count > 0
            tblCalc.ListRows(tblCalc.ListRows.Count).Delete
        Loop
        GoTo SafeExit
    End If

    If currentCalcRows < targetRows Then
        For r = currentCalcRows + 1 To targetRows
            tblCalc.ListRows.Add
        Next r
    ElseIf currentCalcRows > targetRows Then
        For r = currentCalcRows To targetRows + 1 Step -1
            tblCalc.ListRows(r).Delete
        Next r
    End If

    For r = 1 To targetRows
        tblCalc.DataBodyRange.Cells(r, mapCalc("ID")).value = r
    Next r

    For r = 1 To targetRows

        idKey = CStr(r)

        If wbsIdRows.Exists(idKey) Then

            wbsRowIndex = wbsIdRows(idKey)

            For i = LBound(colsToCopy) To UBound(colsToCopy)
                If mapWBS.Exists(colsToCopy(i)) And mapCalc.Exists(colsToCopy(i)) Then

                    If colsToCopy(i) = "WBS" Or colsToCopy(i) = "Predecessors WBS" Then
                        With tblCalc.DataBodyRange.Cells(r, mapCalc(colsToCopy(i)))
                            .NumberFormat = "@"
                            .value = CStr(tblWBS.DataBodyRange.Cells(wbsRowIndex, mapWBS(colsToCopy(i))).value)
                        End With

                    ElseIf colsToCopy(i) = "Task Type" Then
                        taskTypeVal = NormalizeTaskTypeValue(tblWBS.DataBodyRange.Cells(wbsRowIndex, mapWBS("Task Type")).value)
                        With tblCalc.DataBodyRange.Cells(r, mapCalc("Task Type"))
                            .NumberFormat = "@"
                            .value = taskTypeVal
                        End With

                    ElseIf colsToCopy(i) = "Cal" Then
                        If IsCalendarIgnoredWBSRow(tblWBS, mapWBS, wbsRowIndex, summaryWbsByWbs) Then
                            calVal = vbNullString
                        Else
                            calVal = NormalizeCalendarType(tblWBS.DataBodyRange.Cells(wbsRowIndex, mapWBS("Cal")).value)
                        End If
                        With tblCalc.DataBodyRange.Cells(r, mapCalc("Cal"))
                            .NumberFormat = "@"
                            .value = calVal
                        End With

                    Else
                        tblCalc.DataBodyRange.Cells(r, mapCalc(colsToCopy(i))).value = _
                            tblWBS.DataBodyRange.Cells(wbsRowIndex, mapWBS(colsToCopy(i))).value
                    End If

                End If
            Next i
        Else

            For i = LBound(colsToCopy) To UBound(colsToCopy)
                If mapCalc.Exists(colsToCopy(i)) Then
                    tblCalc.DataBodyRange.Cells(r, mapCalc(colsToCopy(i))).ClearContents
                End If
            Next i

            For i = LBound(colsToClear) To UBound(colsToClear)
                If mapCalc.Exists(colsToClear(i)) Then
                    tblCalc.DataBodyRange.Cells(r, mapCalc(colsToClear(i))).ClearContents
                End If
            Next i

        End If

    Next r

    If Not preserveCalcOutputs Then
        For i = LBound(colsToClear) To UBound(colsToClear)
            If mapCalc.Exists(colsToClear(i)) Then
                If Not tblCalc.ListColumns(colsToClear(i)).DataBodyRange Is Nothing Then
                    tblCalc.ListColumns(colsToClear(i)).DataBodyRange.ClearContents
                End If
            End If
        Next i
    End If

SafeExit:
    Application.EnableEvents = True
    Application.ScreenUpdating = True

    If Err.Number <> 0 Then
        If consoleMessages Is Nothing Then Set consoleMessages = New Collection
        DataSync_AddConsoleMessage consoleMessages, "STOP", _
            "Erreur dans Sync_WBS_To_CALC : " & Err.Description, _
            "Error in Sync_WBS_To_CALC: " & Err.Description
    End If

    If Not consoleMessages Is Nothing Then
        CalcBridge_ShowPlanningConsole consoleMessages
    End If

End Sub

Private Sub EnsureTaskTypeColumnExists( _
    ByVal tblWBS As ListObject, _
    ByVal tblCalc As ListObject)

    Dim wbsTaskTypeIndex As Long
    Dim calcTaskTypeIndex As Long
    Dim newCol As ListColumn

    If tblWBS Is Nothing Then Exit Sub
    If tblCalc Is Nothing Then Exit Sub

    '--------------------------------------------------
    ' WBS must already contain the user input column.
    '--------------------------------------------------
    wbsTaskTypeIndex = 0

    On Error Resume Next
    wbsTaskTypeIndex = tblWBS.ListColumns("Task Type").Index
    On Error GoTo 0

    If wbsTaskTypeIndex <= 0 Then
        Err.Raise vbObjectError + 2301, "EnsureTaskTypeColumnExists", _
            "Missing required WBS input column: Task Type"
    End If

    '--------------------------------------------------
    ' CALC must contain Task Type, but we must NOT insert
    ' it inside the existing table structure.
    '
    ' Reason:
    ' - The previous patch inserted Task Type after IsSummary.
    ' - That changed the internal structure/order of tbl_CALC.
    ' - Some existing analytics plumbing is clearly still sensitive
    '   to the table structure / right-side analytics area.
    '
    ' Safer rule:
    ' - If Task Type exists: do nothing.
    ' - If missing: append it at the END of tbl_CALC only.
    ' - Do not move existing columns.
    ' - Do not touch analytics columns.
    '--------------------------------------------------
    calcTaskTypeIndex = 0

    On Error Resume Next
    calcTaskTypeIndex = tblCalc.ListColumns("Task Type").Index
    On Error GoTo 0

    If calcTaskTypeIndex <= 0 Then
        Set newCol = tblCalc.ListColumns.Add
        newCol.Name = "Task Type"
    End If

    If Not tblCalc.DataBodyRange Is Nothing Then
        tblCalc.ListColumns("Task Type").DataBodyRange.NumberFormat = "@"
    End If

End Sub

Private Sub EnsureCalendarColumnExists( _
    ByVal tblWBS As ListObject, _
    ByVal tblCalc As ListObject)

    Dim newCol As ListColumn

    If tblWBS Is Nothing Then Exit Sub
    If tblCalc Is Nothing Then Exit Sub

    If Not TableHasColumn(tblWBS, "Cal") Then
        Set newCol = tblWBS.ListColumns.Add
        newCol.Name = "Cal"
    End If

    If Not TableHasColumn(tblCalc, "Cal") Then
        Set newCol = tblCalc.ListColumns.Add
        newCol.Name = "Cal"
    End If

    If Not tblWBS.DataBodyRange Is Nothing Then
        tblWBS.ListColumns("Cal").DataBodyRange.NumberFormat = "@"
    End If

    If Not tblCalc.DataBodyRange Is Nothing Then
        tblCalc.ListColumns("Cal").DataBodyRange.NumberFormat = "@"
    End If

End Sub
Private Sub EnsureDeadlineOutputColumnsExist( _
    ByVal tblWBS As ListObject, _
    ByVal tblCalc As ListObject)

    Dim colIndex As Long
    Dim newCol As ListColumn

    If tblWBS Is Nothing Then Exit Sub
    If tblCalc Is Nothing Then Exit Sub

    
    colIndex = 0
    On Error Resume Next
    colIndex = tblWBS.ListColumns("Deadline").Index
    On Error GoTo 0

    If colIndex > 0 Then
        tblWBS.ListColumns("Deadline").Delete
    End If

    colIndex = 0
    On Error Resume Next
    colIndex = tblWBS.ListColumns("Deadline Float").Index
    On Error GoTo 0

    If colIndex <= 0 Then
        Set newCol = tblWBS.ListColumns.Add
        newCol.Name = "Deadline Float"
    End If

    colIndex = 0
    On Error Resume Next
    colIndex = tblCalc.ListColumns("Deadline").Index
    On Error GoTo 0

    If colIndex <= 0 Then
        Set newCol = tblCalc.ListColumns.Add
        newCol.Name = "Deadline"
    End If

    colIndex = 0
    On Error Resume Next
    colIndex = tblCalc.ListColumns("Deadline Float").Index
    On Error GoTo 0

    If colIndex <= 0 Then
        Set newCol = tblCalc.ListColumns.Add
        newCol.Name = "Deadline Float"
    End If

    If Not tblWBS.DataBodyRange Is Nothing Then
        tblWBS.ListColumns("Deadline Float").DataBodyRange.NumberFormat = "0"
    End If

    If Not tblCalc.DataBodyRange Is Nothing Then
        tblCalc.ListColumns("Deadline").DataBodyRange.NumberFormat = "dd/mm/yyyy"
        tblCalc.ListColumns("Deadline Float").DataBodyRange.NumberFormat = "0"
    End If

End Sub



Private Sub EnsureLongestPathOutputColumnsExist( _
    ByVal tblWBS As ListObject, _
    ByVal tblCalc As ListObject)

    Dim colIndex As Long
    Dim newCol As ListColumn

    If tblWBS Is Nothing Then Exit Sub
    If tblCalc Is Nothing Then Exit Sub

    colIndex = 0
    On Error Resume Next
    colIndex = tblWBS.ListColumns("Longest Path").Index
    On Error GoTo 0

    If colIndex <= 0 Then
        Set newCol = tblWBS.ListColumns.Add
        newCol.Name = "Longest Path"
    End If

    colIndex = 0
    On Error Resume Next
    colIndex = tblCalc.ListColumns("Longest Path").Index
    On Error GoTo 0

    If colIndex <= 0 Then
        Set newCol = tblCalc.ListColumns.Add
        newCol.Name = "Longest Path"
    End If

    If Not tblWBS.DataBodyRange Is Nothing Then
        tblWBS.ListColumns("Longest Path").DataBodyRange.NumberFormat = "@"
    End If

    If Not tblCalc.DataBodyRange Is Nothing Then
        tblCalc.ListColumns("Longest Path").DataBodyRange.NumberFormat = "@"
    End If

End Sub
Private Function NormalizeTaskTypeValue(ByVal rawValue As Variant) As String

    Dim s As String

    s = UCase$(Trim$(CStr(rawValue)))

    Select Case s

        Case "", "TASK", "STANDARD", "NORMAL"
            NormalizeTaskTypeValue = "Task"

        Case "MILESTONE", "MS", "JALON"
            NormalizeTaskTypeValue = "Milestone"

        Case "LEVEL OF EFFORT", "LOE", "LEVEL-OF-EFFORT", "LEVEL_OF_EFFORT"
            NormalizeTaskTypeValue = "Level of Effort"

        Case Else
            Err.Raise vbObjectError + 2302, "NormalizeTaskTypeValue", _
                "Invalid Task Type value: " & CStr(rawValue) & _
                " | Allowed values: Task, Milestone, Level of Effort"

    End Select

End Function

Function Sync_Forecast_Only() As Boolean

    Dim wsWBS As Worksheet
    Dim wsCalc As Worksheet
    Dim tblWBS As ListObject
    Dim tblCalc As ListObject

    Dim mapWBS As Object
    Dim mapCalc As Object
    Dim wbsIdRows As Object

    Dim r As Long
    Dim idValue As Variant
    Dim idKey As String
    Dim wbsRowIndex As Long

    Dim consoleMessages As Collection

    On Error GoTo SafeExit

    Sync_Forecast_Only = False
    Set consoleMessages = New Collection

    Set wsWBS = ThisWorkbook.Worksheets("WBS")
    Set wsCalc = ThisWorkbook.Worksheets("CALC")

    Set tblWBS = wsWBS.ListObjects("tbl_WBS")
    Set tblCalc = wsCalc.ListObjects("tbl_CALC")

    Set mapWBS = CreateObject("Scripting.Dictionary")
    Set mapCalc = CreateObject("Scripting.Dictionary")
    Set wbsIdRows = CreateObject("Scripting.Dictionary")

    For r = 1 To tblWBS.ListColumns.Count
        mapWBS(tblWBS.ListColumns(r).Name) = r
    Next r

    For r = 1 To tblCalc.ListColumns.Count
        mapCalc(tblCalc.ListColumns(r).Name) = r
    Next r

    If Not mapWBS.Exists("ID") Then
        DataSync_AddConsoleMessage consoleMessages, "STOP", _
            "La colonne ID est introuvable dans tbl_WBS.", _
            "Column ID was not found in tbl_WBS."
        GoTo SafeExit
    End If

    If Not mapCalc.Exists("ID") Then
        DataSync_AddConsoleMessage consoleMessages, "STOP", _
            "La colonne ID est introuvable dans tbl_CALC.", _
            "Column ID was not found in tbl_CALC."
        GoTo SafeExit
    End If

    If Not mapWBS.Exists("Forecast Start") Then
        DataSync_AddConsoleMessage consoleMessages, "STOP", _
            "La colonne Forecast Start est introuvable dans tbl_WBS.", _
            "Column Forecast Start was not found in tbl_WBS."
        GoTo SafeExit
    End If

    If Not mapWBS.Exists("Forecast Finish") Then
        DataSync_AddConsoleMessage consoleMessages, "STOP", _
            "La colonne Forecast Finish est introuvable dans tbl_WBS.", _
            "Column Forecast Finish was not found in tbl_WBS."
        GoTo SafeExit
    End If

    If Not mapCalc.Exists("Forecast Start") Then
        DataSync_AddConsoleMessage consoleMessages, "STOP", _
            "La colonne Forecast Start est introuvable dans tbl_CALC.", _
            "Column Forecast Start was not found in tbl_CALC."
        GoTo SafeExit
    End If

    If Not mapCalc.Exists("Forecast Finish") Then
        DataSync_AddConsoleMessage consoleMessages, "STOP", _
            "La colonne Forecast Finish est introuvable dans tbl_CALC.", _
            "Column Forecast Finish was not found in tbl_CALC."
        GoTo SafeExit
    End If

    If tblCalc.DataBodyRange Is Nothing Then GoTo SafeExit
    If tblWBS.DataBodyRange Is Nothing Then GoTo SafeExit

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    For r = 1 To tblWBS.ListRows.Count
        idValue = tblWBS.DataBodyRange.Cells(r, mapWBS("ID")).value
        If Trim(CStr(idValue)) <> "" Then
            idKey = CStr(idValue)
            wbsIdRows(idKey) = r
        End If
    Next r

    For r = 1 To tblCalc.ListRows.Count

        idValue = tblCalc.DataBodyRange.Cells(r, mapCalc("ID")).value

        If Trim(CStr(idValue)) <> "" Then
            idKey = CStr(idValue)

            If wbsIdRows.Exists(idKey) Then
                wbsRowIndex = wbsIdRows(idKey)

                tblCalc.DataBodyRange.Cells(r, mapCalc("Forecast Start")).value = _
                    tblWBS.DataBodyRange.Cells(wbsRowIndex, mapWBS("Forecast Start")).value

                tblCalc.DataBodyRange.Cells(r, mapCalc("Forecast Finish")).value = _
                    tblWBS.DataBodyRange.Cells(wbsRowIndex, mapWBS("Forecast Finish")).value
            End If
        End If

    Next r

    Sync_Forecast_Only = True

SafeExit:
    Application.EnableEvents = True
    Application.ScreenUpdating = True

    If Err.Number <> 0 Then
        If consoleMessages Is Nothing Then Set consoleMessages = New Collection
        DataSync_AddConsoleMessage consoleMessages, "STOP", _
            "Erreur dans Sync_Forecast_Only : " & Err.Description, _
            "Error in Sync_Forecast_Only: " & Err.Description
        Sync_Forecast_Only = False
    End If

    If Not consoleMessages Is Nothing Then
        CalcBridge_ShowPlanningConsole consoleMessages
    End If

End Function

Public Sub Push_Calculated_Back_To_WBS()

    Dim wsWBS As Worksheet
    Dim wsCalc As Worksheet
    Dim tblWBS As ListObject
    Dim tblCalc As ListObject

    Dim mapWBS As Object
    Dim mapCalc As Object
    Dim calcRowById As Object

    Dim allowedFields As Variant
    Dim authorizedFields As Variant
    Dim outCols As Object

    Dim arrWBS As Variant
    Dim arrCalc As Variant
    Dim outArr() As Variant

    Dim r As Long
    Dim c As Long
    Dim i As Long

    Dim wbsRows As Long
    Dim calcRows As Long

    Dim id As String
    Dim fieldName As String
    Dim calcRow As Long

    Dim consoleMessages As Collection

    On Error GoTo SafeExit

    Set consoleMessages = New Collection

    Set wsWBS = ThisWorkbook.Worksheets("WBS")
    Set wsCalc = ThisWorkbook.Worksheets("CALC")

    Set tblWBS = wsWBS.ListObjects("tbl_WBS")
    Set tblCalc = wsCalc.ListObjects("tbl_CALC")

    If tblWBS.DataBodyRange Is Nothing Then Exit Sub
    If tblCalc.DataBodyRange Is Nothing Then Exit Sub

    Set mapWBS = CreateObject("Scripting.Dictionary")
    Set mapCalc = CreateObject("Scripting.Dictionary")
    Set calcRowById = CreateObject("Scripting.Dictionary")
    Set outCols = CreateObject("Scripting.Dictionary")

    allowedFields = Array( _
        "Calculated Start", _
        "Calculated Finish", _
        "Driving Logic", _
        "Critical Path", _
        "Longest Path", _
        "Critical Path REX", _
        "Total Float", _
        "Free Float", _
        "Total Float REX", _
        "Free Float REX", _
        "Deadline Float")

    authorizedFields = Array( _
        "Calculated Start", _
        "Calculated Finish", _
        "Driving Logic", _
        "Critical Path", _
        "Longest Path", _
        "Critical Path REX", _
        "Total Float", _
        "Free Float", _
        "Total Float REX", _
        "Free Float REX", _
        "Deadline Float", _
        "Baseline Finish", _
        "Actual Duration", _
        "Calculated Duration")

    For c = 1 To tblWBS.ListColumns.Count
        mapWBS(tblWBS.ListColumns(c).Name) = c
    Next c

    For c = 1 To tblCalc.ListColumns.Count
        mapCalc(tblCalc.ListColumns(c).Name) = c
    Next c

    If Not mapWBS.Exists("ID") Then
        DataSync_AddConsoleMessage consoleMessages, "STOP", _
            "La colonne ID est introuvable dans tbl_WBS.", _
            "Column ID was not found in tbl_WBS."
        GoTo SafeExit
    End If

    If Not mapCalc.Exists("ID") Then
        DataSync_AddConsoleMessage consoleMessages, "STOP", _
            "La colonne ID est introuvable dans tbl_CALC.", _
            "Column ID was not found in tbl_CALC."
        GoTo SafeExit
    End If

    For i = LBound(allowedFields) To UBound(allowedFields)

        fieldName = CStr(allowedFields(i))

        If Not mapWBS.Exists(fieldName) Then
            DataSync_AddConsoleMessage consoleMessages, "STOP", _
                "Colonne de sortie introuvable dans tbl_WBS : " & fieldName, _
                "Output column not found in tbl_WBS: " & fieldName
            GoTo SafeExit
        End If

        If Not mapCalc.Exists(fieldName) Then
            DataSync_AddConsoleMessage consoleMessages, "STOP", _
                "Colonne de sortie introuvable dans tbl_CALC : " & fieldName, _
                "Output column not found in tbl_CALC: " & fieldName
            GoTo SafeExit
        End If

    Next i

    arrWBS = tblWBS.DataBodyRange.value
    arrCalc = tblCalc.DataBodyRange.value

    wbsRows = UBound(arrWBS, 1)
    calcRows = UBound(arrCalc, 1)

    For r = 1 To calcRows
        id = Trim$(CStr(arrCalc(r, mapCalc("ID"))))
        If id <> "" Then
            If Not calcRowById.Exists(id) Then
                calcRowById(id) = r
            End If
        End If
    Next r

    For i = LBound(allowedFields) To UBound(allowedFields)

        fieldName = CStr(allowedFields(i))
        ReDim outArr(1 To wbsRows, 1 To 1)

        For r = 1 To wbsRows

            id = Trim$(CStr(arrWBS(r, mapWBS("ID"))))

            If id <> "" Then
                If calcRowById.Exists(id) Then
                    calcRow = CLng(calcRowById(id))
                    outArr(r, 1) = arrCalc(calcRow, mapCalc(fieldName))
                Else
                    outArr(r, 1) = Empty
                End If
            Else
                outArr(r, 1) = Empty
            End If

        Next r

        outCols(fieldName) = outArr

    Next i

    BeginAuthorizedWBSWrite "Push_Calculated_Back_To_WBS", authorizedFields

    For i = LBound(allowedFields) To UBound(allowedFields)
        fieldName = CStr(allowedFields(i))
        tblWBS.ListColumns(fieldName).DataBodyRange.value = outCols(fieldName)
    Next i

    RestoreWBSFormulaColumns tblWBS

SafeExit:
    EndAuthorizedWBSWrite

    If Err.Number <> 0 Then
        If consoleMessages Is Nothing Then Set consoleMessages = New Collection
        DataSync_AddConsoleMessage consoleMessages, "STOP", _
            "Erreur dans Push_Calculated_Back_To_WBS : " & Err.Description, _
            "Error in Push_Calculated_Back_To_WBS: " & Err.Description
    End If

    If Not consoleMessages Is Nothing Then
        CalcBridge_ShowPlanningConsole consoleMessages
    End If

End Sub

Private Function IsAllowedCalculatedPushField(ByVal fieldName As String) As Boolean

    Select Case fieldName
        Case "Calculated Start", _
             "Calculated Finish", _
             "Driving Logic", _
             "Critical Path", _
             "Longest Path", _
             "Critical Path REX", _
             "Total Float", _
             "Free Float", _
             "Total Float REX", _
             "Free Float REX", _
             "Deadline Float", _
             "Error flag", _
             "ErrorMsg"

            IsAllowedCalculatedPushField = True

        Case Else
            IsAllowedCalculatedPushField = False

    End Select

End Function

Private Sub ApplyWBSDateFormats(ByVal tblWBS As ListObject)

    On Error Resume Next

    tblWBS.ListColumns("Baseline Start").DataBodyRange.NumberFormat = "dd/mm/yyyy"
    tblWBS.ListColumns("Baseline Finish").DataBodyRange.NumberFormat = "dd/mm/yyyy"
    tblWBS.ListColumns("Actual Start").DataBodyRange.NumberFormat = "dd/mm/yyyy"
    tblWBS.ListColumns("Actual Finish").DataBodyRange.NumberFormat = "dd/mm/yyyy"
    tblWBS.ListColumns("Forecast Start").DataBodyRange.NumberFormat = "dd/mm/yyyy"
    tblWBS.ListColumns("Forecast Finish").DataBodyRange.NumberFormat = "dd/mm/yyyy"
    tblWBS.ListColumns("Deadline Float").DataBodyRange.NumberFormat = "0"
    tblWBS.ListColumns("Calculated Start").DataBodyRange.NumberFormat = "dd/mm/yyyy"
    tblWBS.ListColumns("Calculated Finish").DataBodyRange.NumberFormat = "dd/mm/yyyy"

    On Error GoTo 0

End Sub

Public Sub RebuildLogicLinksTable()

    Dim wsWBS As Worksheet
    Dim wsCalc As Worksheet
    Dim tblWBS As ListObject
    Dim tblLinks As ListObject

    Dim mapWBS As Object
    Dim wbsToId As Object

    Dim arr As Variant
    Dim r As Long
    Dim rowCount As Long

    Dim succId As String
    Dim succWBS As String
    Dim predText As String

    Dim linksOut As Collection
    Dim errText As String
    Dim linkRow As Object

    Dim outArr() As Variant
    Dim outCount As Long
    Dim i As Long
    Dim writeRow As Long

    Dim consoleMessages As Collection

    On Error GoTo SafeExit

    Set consoleMessages = New Collection

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    Set wsWBS = ThisWorkbook.Worksheets("WBS")
    Set wsCalc = ThisWorkbook.Worksheets("CALC")

    Set tblWBS = wsWBS.ListObjects("tbl_WBS")
    Set tblLinks = EnsureLogicLinksTable(wsCalc)

    If tblWBS.DataBodyRange Is Nothing Then
        ClearLogicLinksTableRows tblLinks
        GoTo SafeExit
    End If

    Set mapWBS = CreateObject("Scripting.Dictionary")
    For i = 1 To tblWBS.ListColumns.Count
        mapWBS(tblWBS.ListColumns(i).Name) = i
    Next i

    Set wbsToId = BuildWbsToIdMapFromTable(tblWBS, mapWBS)

    arr = tblWBS.DataBodyRange.value
    rowCount = UBound(arr, 1)

    outCount = 0

    For r = 1 To rowCount

        succId = Trim$(CStr(arr(r, mapWBS("ID"))))
        succWBS = NormalizeWBS(CStr(arr(r, mapWBS("WBS"))))
        predText = Trim$(CStr(arr(r, mapWBS("Predecessors WBS"))))

        If succId <> "" And succWBS <> "" Then

            If Not ParsePredecessorsText(succId, succWBS, predText, wbsToId, linksOut, errText) Then
                DataSync_AddConsoleMessage consoleMessages, "STOP", _
                    "Erreur lors de la reconstruction de tbl_LOGIC_LINKS." & vbCrLf & _
                    "-> " & errText, _
                    "Error while rebuilding tbl_LOGIC_LINKS." & vbCrLf & _
                    "-> " & errText
                GoTo SafeExit
            End If

            outCount = outCount + linksOut.Count

        End If

    Next r

    If outCount <= 0 Then
        ClearLogicLinksTableRows tblLinks
        GoTo SafeExit
    End If

    ReDim outArr(1 To outCount, 1 To 7)

    writeRow = 0

    For r = 1 To rowCount

        succId = Trim$(CStr(arr(r, mapWBS("ID"))))
        succWBS = NormalizeWBS(CStr(arr(r, mapWBS("WBS"))))
        predText = Trim$(CStr(arr(r, mapWBS("Predecessors WBS"))))

        If succId <> "" And succWBS <> "" Then

            If Not ParsePredecessorsText(succId, succWBS, predText, wbsToId, linksOut, errText) Then
                DataSync_AddConsoleMessage consoleMessages, "STOP", _
                    "Erreur lors de la reconstruction de tbl_LOGIC_LINKS." & vbCrLf & _
                    "-> " & errText, _
                    "Error while rebuilding tbl_LOGIC_LINKS." & vbCrLf & _
                    "-> " & errText
                GoTo SafeExit
            End If

            If linksOut.Count > 0 Then
                For i = 1 To linksOut.Count
                    Set linkRow = linksOut(i)

                    writeRow = writeRow + 1

                    outArr(writeRow, 1) = CStr(linkRow("Succ ID"))
                    outArr(writeRow, 2) = CStr(linkRow("Succ WBS"))
                    outArr(writeRow, 3) = CStr(linkRow("Pred ID"))
                    outArr(writeRow, 4) = CStr(linkRow("Pred WBS"))
                    outArr(writeRow, 5) = CStr(linkRow("Link Type"))
                    outArr(writeRow, 6) = CLng(linkRow("Lag"))
                    outArr(writeRow, 7) = CStr(linkRow("Raw Token"))
                Next i
            End If

        End If

    Next r

    RewriteLogicLinksTable tblLinks, outArr, outCount

SafeExit:
    Application.EnableEvents = True
    Application.ScreenUpdating = True

    If Err.Number <> 0 Then
        If consoleMessages Is Nothing Then Set consoleMessages = New Collection
        DataSync_AddConsoleMessage consoleMessages, "STOP", _
            "Erreur VBA dans RebuildLogicLinksTable : " & Err.Description, _
            "VBA error in RebuildLogicLinksTable: " & Err.Description
    End If

    If Not consoleMessages Is Nothing Then
        CalcBridge_ShowPlanningConsole consoleMessages
    End If

End Sub

Private Function EnsureLogicLinksTable(ByVal wsCalc As Worksheet) As ListObject

    Dim tbl As ListObject
    Dim headerRange As Range
    Dim fullRange As Range

    On Error Resume Next
    Set tbl = wsCalc.ListObjects(LOGIC_LINKS_TABLE_NAME)
    On Error GoTo 0

    If tbl Is Nothing Then

        Set headerRange = wsCalc.Range(LOGIC_LINKS_FIRST_CELL).Resize(1, 7)
        WriteLogicLinksHeaders headerRange

        Set fullRange = headerRange.Resize(2, 7)
        Set tbl = wsCalc.ListObjects.Add(xlSrcRange, fullRange, , xlYes)
        tbl.Name = LOGIC_LINKS_TABLE_NAME

    Else
        WriteLogicLinksHeaders tbl.HeaderRowRange
    End If

    Set EnsureLogicLinksTable = tbl

End Function


Private Sub WriteLogicLinksHeaders(ByVal headerRange As Range)

    headerRange.Cells(1, 1).value = "Succ ID"
    headerRange.Cells(1, 2).value = "Succ WBS"
    headerRange.Cells(1, 3).value = "Pred ID"
    headerRange.Cells(1, 4).value = "Pred WBS"
    headerRange.Cells(1, 5).value = "Link Type"
    headerRange.Cells(1, 6).value = "Lag"
    headerRange.Cells(1, 7).value = "Raw Token"

End Sub


Private Sub ClearLogicLinksTableRows(ByVal tbl As ListObject)

    Do While tbl.ListRows.Count > 1
        tbl.ListRows(tbl.ListRows.Count).Delete
    Loop

    If tbl.ListRows.Count = 1 Then
        tbl.DataBodyRange.Rows(1).ClearContents
    End If

End Sub


Private Sub RewriteLogicLinksTable( _
    ByVal tbl As ListObject, _
    ByRef outArr() As Variant, _
    ByVal outCount As Long)

    Dim targetRows As Long
    Dim currentRows As Long
    Dim targetRange As Range

    ClearLogicLinksTableRows tbl

    If outCount <= 0 Then Exit Sub

    currentRows = tbl.ListRows.Count
    targetRows = outCount

    If currentRows < targetRows Then
        Do While tbl.ListRows.Count < targetRows
            tbl.ListRows.Add
        Loop
    ElseIf currentRows > targetRows Then
        Do While tbl.ListRows.Count > targetRows
            tbl.ListRows(tbl.ListRows.Count).Delete
        Loop
    End If

    Set targetRange = tbl.DataBodyRange.Resize(targetRows, 7)

    ' Force text format BEFORE writing values
    targetRange.Columns(1).NumberFormat = "@"
    targetRange.Columns(2).NumberFormat = "@"
    targetRange.Columns(3).NumberFormat = "@"
    targetRange.Columns(4).NumberFormat = "@"
    targetRange.Columns(5).NumberFormat = "@"
    targetRange.Columns(7).NumberFormat = "@"
    targetRange.Columns(6).NumberFormat = "0"

    targetRange.value = outArr

    ApplyLogicLinksTableFormats tbl

End Sub


Private Sub ApplyLogicLinksTableFormats(ByVal tbl As ListObject)

    On Error Resume Next

    tbl.ListColumns("Succ ID").DataBodyRange.NumberFormat = "@"
    tbl.ListColumns("Succ WBS").DataBodyRange.NumberFormat = "@"
    tbl.ListColumns("Pred ID").DataBodyRange.NumberFormat = "@"
    tbl.ListColumns("Pred WBS").DataBodyRange.NumberFormat = "@"
    tbl.ListColumns("Link Type").DataBodyRange.NumberFormat = "@"
    tbl.ListColumns("Raw Token").DataBodyRange.NumberFormat = "@"

    tbl.ListColumns("Lag").DataBodyRange.NumberFormat = "0"

    On Error GoTo 0

End Sub

Private Sub RestoreWBSFormulaColumns(ByVal tblWBS As ListObject)

    Dim consoleMessages As Collection

    On Error GoTo SafeExit

    Set consoleMessages = New Collection

    If tblWBS Is Nothing Then Exit Sub
    If tblWBS.DataBodyRange Is Nothing Then Exit Sub

    If TableHasColumn(tblWBS, "Baseline Finish") Then
        tblWBS.ListColumns("Baseline Finish").DataBodyRange.FormulaLocal = _
            "=SI(OU([@[Baseline Start]]="""";[@[Baseline Duration]]="""");"""";[@[Baseline Start]]+[@[Baseline Duration]]-1)"
        tblWBS.ListColumns("Baseline Finish").DataBodyRange.NumberFormat = "dd/mm/yyyy"
    End If

    If TableHasColumn(tblWBS, "Actual Duration") Then
        tblWBS.ListColumns("Actual Duration").DataBodyRange.FormulaLocal = _
            "=SI(OU([@[Actual Start]]="""";[@[Actual Finish]]="""");"""";[@[Actual Finish]]-[@[Actual Start]]+1)"
        tblWBS.ListColumns("Actual Duration").DataBodyRange.NumberFormat = "0"
    End If

    If TableHasColumn(tblWBS, "Calculated Duration") Then
        tblWBS.ListColumns("Calculated Duration").DataBodyRange.FormulaLocal = _
            "=SI(OU([@[Calculated Start]]="""";[@[Calculated Finish]]="""");"""";[@[Calculated Finish]]-[@[Calculated Start]]+1)"
        tblWBS.ListColumns("Calculated Duration").DataBodyRange.NumberFormat = "0"
    End If

SafeExit:
    If Err.Number <> 0 Then
        If consoleMessages Is Nothing Then Set consoleMessages = New Collection
        DataSync_AddConsoleMessage consoleMessages, "STOP", _
            "Erreur dans RestoreWBSFormulaColumns : " & Err.Description, _
            "Error in RestoreWBSFormulaColumns: " & Err.Description
        CalcBridge_ShowPlanningConsole consoleMessages
    End If

End Sub


Private Function TableHasColumn(ByVal tbl As ListObject, ByVal columnName As String) As Boolean

    Dim col As ListColumn

    On Error Resume Next
    Set col = tbl.ListColumns(columnName)
    On Error GoTo 0

    TableHasColumn = Not col Is Nothing

End Function

'=====================================================
' Task Type WBS input setup
'=====================================================
' Purpose:
' - Add dropdown validation on tbl_WBS[Task Type].
' - Fill blank Task Type cells with "Task".
' - Keep existing valid user values.
'
' Allowed values:
' - Task
' - Milestone
' - Level of Effort
'
' Where to call:
' - In Sync_WBS_To_CALC, after:
'       Set tblWBS = wsWBS.ListObjects("tbl_WBS")
'   and before reading data from tbl_WBS.
'
' Call line:
'       EnsureWBSTaskTypeInputSetup tblWBS
'=====================================================


Private Sub EnsureWBSTaskTypeInputSetup(ByVal tblWBS As ListObject)

    Dim rng As Range
    Dim cell As Range
    Dim normalizedValue As String

    If tblWBS Is Nothing Then Exit Sub

    If Not TableHasColumn(tblWBS, "Task Type") Then
        Err.Raise vbObjectError + 2310, "EnsureWBSTaskTypeInputSetup", _
            "Missing required WBS input column: Task Type"
    End If

    If tblWBS.DataBodyRange Is Nothing Then Exit Sub

    Set rng = tblWBS.ListColumns("Task Type").DataBodyRange

    'Default blank cells to Task.
    For Each cell In rng.Cells

        normalizedValue = NormalizeTaskTypeValue(cell.value)

        If Trim$(CStr(cell.value)) = "" Then
            cell.value = "Task"
        ElseIf CStr(cell.value) <> normalizedValue Then
            cell.value = normalizedValue
        End If

    Next cell

    'Apply dropdown validation.
    With rng.Validation
        .Delete
        .Add Type:=xlValidateList, _
             AlertStyle:=xlValidAlertStop, _
             Operator:=xlBetween, _
             Formula1:="Task,Milestone,Level of Effort"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = "Task Type"
        .InputMessage = "Choose: Task, Milestone, or Level of Effort."
        .ErrorTitle = "Invalid Task Type"
        .errorMessage = "Allowed values: Task, Milestone, Level of Effort."
        .ShowInput = True
        .ShowError = True
    End With

    rng.NumberFormat = "@"

End Sub

Private Sub EnsureWBSCalendarInputSetup(ByVal tblWBS As ListObject)

    Dim rng As Range
    Dim cell As Range
    Dim normalizedValue As String
    Dim newCol As ListColumn

    If tblWBS Is Nothing Then Exit Sub

    If Not TableHasColumn(tblWBS, "Cal") Then
        Set newCol = tblWBS.ListColumns.Add
        newCol.Name = "Cal"
    End If

    If tblWBS.DataBodyRange Is Nothing Then Exit Sub

    Set rng = tblWBS.ListColumns("Cal").DataBodyRange

    For Each cell In rng.Cells
        If Trim$(CStr(cell.value)) <> "" Then
            normalizedValue = NormalizeCalendarType(cell.value)
            If CStr(cell.value) <> normalizedValue Then cell.value = normalizedValue
        End If
    Next cell

    With rng.Validation
        .Delete
        .Add Type:=xlValidateList, _
             AlertStyle:=xlValidAlertStop, _
             Operator:=xlBetween, _
             Formula1:=CALENDAR_7D & "," & CALENDAR_6D & "," & CALENDAR_5D
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = "Cal"
        .InputMessage = "Choose: 7j/7, 6j/7, or 5j/7."
        .ErrorTitle = "Invalid Cal"
        .errorMessage = "Allowed values: blank, 7j/7, 6j/7, 5j/7."
        .ShowInput = True
        .ShowError = True
    End With

    rng.NumberFormat = "@"

End Sub

Private Function BuildWBSSummaryWbsLookup( _
    ByVal tblWBS As ListObject, _
    ByVal mapWBS As Object) As Object

    Dim summaryWbs As Object
    Dim r As Long
    Dim wbsVal As String
    Dim parentWbs As String

    Set summaryWbs = CreateObject("Scripting.Dictionary")

    If tblWBS Is Nothing Then
        Set BuildWBSSummaryWbsLookup = summaryWbs
        Exit Function
    End If
    If tblWBS.DataBodyRange Is Nothing Then
        Set BuildWBSSummaryWbsLookup = summaryWbs
        Exit Function
    End If
    If mapWBS Is Nothing Then
        Set BuildWBSSummaryWbsLookup = summaryWbs
        Exit Function
    End If
    If Not mapWBS.Exists("WBS") Then
        Set BuildWBSSummaryWbsLookup = summaryWbs
        Exit Function
    End If

    For r = 1 To tblWBS.ListRows.Count
        wbsVal = Replace$(Trim$(CStr(tblWBS.DataBodyRange.Cells(r, mapWBS("WBS")).value)), ",", ".")
        parentWbs = GetParentWBS(wbsVal)

        Do While parentWbs <> ""
            summaryWbs(parentWbs) = True
            parentWbs = GetParentWBS(parentWbs)
        Loop
    Next r

    Set BuildWBSSummaryWbsLookup = summaryWbs

End Function

Private Function IsCalendarIgnoredWBSRow( _
    ByVal tblWBS As ListObject, _
    ByVal mapWBS As Object, _
    ByVal rowIndex As Long, _
    ByVal summaryWbsByWbs As Object) As Boolean

    Dim wbsVal As String
    Dim taskTypeVal As String

    If tblWBS Is Nothing Then Exit Function
    If tblWBS.DataBodyRange Is Nothing Then Exit Function
    If mapWBS Is Nothing Then Exit Function
    If rowIndex < 1 Or rowIndex > tblWBS.ListRows.Count Then Exit Function

    If mapWBS.Exists("WBS") Then
        wbsVal = Replace$(Trim$(CStr(tblWBS.DataBodyRange.Cells(rowIndex, mapWBS("WBS")).value)), ",", ".")
        If Not summaryWbsByWbs Is Nothing Then
            If summaryWbsByWbs.Exists(wbsVal) Then
                IsCalendarIgnoredWBSRow = True
                Exit Function
            End If
        End If
    End If

    If mapWBS.Exists("Task Type") Then
        taskTypeVal = UCase$(NormalizeTaskTypeValue(tblWBS.DataBodyRange.Cells(rowIndex, mapWBS("Task Type")).value))
        IsCalendarIgnoredWBSRow = (taskTypeVal = "LEVEL OF EFFORT" Or taskTypeVal = "MILESTONE")
    End If

End Function

Private Sub NormalizeWBSCalendarValues( _
    ByVal tblWBS As ListObject, _
    ByVal mapWBS As Object, _
    ByVal summaryWbsByWbs As Object)

    Dim r As Long
    Dim cell As Range
    Dim normalizedValue As String

    If tblWBS Is Nothing Then Exit Sub
    If tblWBS.DataBodyRange Is Nothing Then Exit Sub
    If mapWBS Is Nothing Then Exit Sub
    If Not mapWBS.Exists("Cal") Then Exit Sub

    For r = 1 To tblWBS.ListRows.Count
        Set cell = tblWBS.DataBodyRange.Cells(r, mapWBS("Cal"))
        normalizedValue = NormalizeCalendarType(cell.value)

        If Trim$(CStr(cell.value)) = "" Then
            If Not IsCalendarIgnoredWBSRow(tblWBS, mapWBS, r, summaryWbsByWbs) Then
                cell.value = CALENDAR_7D
            End If
        ElseIf CStr(cell.value) <> normalizedValue Then
            cell.value = normalizedValue
        End If
    Next r

End Sub
Private Sub DataSync_AddConsoleMessage( _
    ByVal consoleMessages As Collection, _
    ByVal msgType As String, _
    ByVal frText As String, _
    ByVal enText As String)

    If consoleMessages Is Nothing Then Exit Sub

    CalcBridge_AddConsoleMessage consoleMessages, msgType, _
        BiMsg(frText, enText)

End Sub

