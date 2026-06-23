Attribute VB_Name = "mod_CalcState"
Option Explicit

Private Const CALC_STATE_SHEET_NAME As String = "CALC_STATE"
Private Const CALC_STATE_TABLE_NAME As String = "tbl_CALC_STATE"

Public Sub Ensure_CalcState_Table()

    Dim localConsole As Collection

    Set localConsole = New Collection
    Ensure_CalcState_Table_Console localConsole

    If localConsole.Count > 0 Then
        CalcBridge_ShowPlanningConsole localConsole
    End If

End Sub

Public Sub Ensure_CalcState_Table_Console(ByVal consoleMessages As Collection)

    Dim wsState As Worksheet
    Dim tblState As ListObject
    Dim wasSheetCreated As Boolean
    Dim wasTableCreated As Boolean

    On Error GoTo SafeExit

    Application.ScreenUpdating = False

    Set wsState = EnsureWorksheetExists_CalcState(CALC_STATE_SHEET_NAME, wasSheetCreated)
    Set tblState = EnsureTableWithHeaders_CalcState( _
        wsState, _
        CALC_STATE_TABLE_NAME, _
        wsState.Range("A1"), _
        BuildCalcStateHeaders(), _
        wasTableCreated)

    If wasSheetCreated Or wasTableCreated Then
        ApplyCalcStateFormats tblState
    End If

SafeExit:
    Application.ScreenUpdating = True

    If Err.Number <> 0 Then
        CalcState_AddStopToConsole consoleMessages, _
            "Erreur dans Ensure_CalcState_Table : " & Err.Description, _
            "Error in Ensure_CalcState_Table: " & Err.Description
    End If

End Sub

Public Sub Write_CalcState_Snapshot(ByVal runStatus As String)

    Dim localConsole As Collection

    Set localConsole = New Collection
    Write_CalcState_Snapshot_Console runStatus, localConsole

    If localConsole.Count > 0 Then
        CalcBridge_ShowPlanningConsole localConsole
    End If

End Sub


Public Sub Write_CalcState_Snapshot_Console( _
    ByVal runStatus As String, _
    ByVal consoleMessages As Collection)

    Dim perfScope As clsPerfScope

    Dim wsCalc As Worksheet
    Dim wsState As Worksheet
    Dim tblCalc As ListObject
    Dim tblState As ListObject

    Dim mapCalc As Object
    Dim mapState As Object

    Dim arrCalc As Variant
    Dim arrOut() As Variant

    Dim rowCount As Long
    Dim r As Long

    Set perfScope = Profiler_BeginScope("Write_CalcState_Snapshot_Console", "Excel Table Write")

    On Error GoTo SafeExit

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    Ensure_CalcState_Table_Console consoleMessages

    Set wsCalc = ThisWorkbook.Worksheets("CALC")
    Set wsState = ThisWorkbook.Worksheets(CALC_STATE_SHEET_NAME)

    Set tblCalc = wsCalc.ListObjects("tbl_CALC")
    Set tblState = wsState.ListObjects(CALC_STATE_TABLE_NAME)

    Set mapCalc = BuildColumnMap_CalcState(tblCalc)
    Set mapState = BuildColumnMap_CalcState(tblState)

    ValidateCalcStateSourceColumns mapCalc
    ValidateCalcStateTargetColumns mapState

    If tblCalc.DataBodyRange Is Nothing Then
        ResizeTableToRowCount_CalcState tblState, 0
        GoTo SafeExit
    End If

    arrCalc = tblCalc.DataBodyRange.value
    rowCount = UBound(arrCalc, 1)

    ResizeTableToRowCount_CalcState tblState, rowCount
    ReDim arrOut(1 To rowCount, 1 To tblState.ListColumns.Count)

    For r = 1 To rowCount

        arrOut(r, mapState("ID")) = GetCalcStateVal(arrCalc, r, mapCalc, "ID")
        arrOut(r, mapState("ParentID")) = GetCalcStateVal(arrCalc, r, mapCalc, "ParentID")
        arrOut(r, mapState("IsSummary")) = GetCalcStateVal(arrCalc, r, mapCalc, "IsSummary")
        arrOut(r, mapState("Predecessors WBS")) = GetCalcStateVal(arrCalc, r, mapCalc, "Predecessors WBS")
        arrOut(r, mapState("Cal")) = NormalizeCalendarType(GetCalcStateVal(arrCalc, r, mapCalc, "Cal"))

        arrOut(r, mapState("Baseline Start")) = GetCalcStateVal(arrCalc, r, mapCalc, "Baseline Start")
        arrOut(r, mapState("Baseline Duration")) = GetCalcStateVal(arrCalc, r, mapCalc, "Baseline Duration")
        arrOut(r, mapState("Actual Start")) = GetCalcStateVal(arrCalc, r, mapCalc, "Actual Start")
        arrOut(r, mapState("Actual Finish")) = GetCalcStateVal(arrCalc, r, mapCalc, "Actual Finish")
        arrOut(r, mapState("Forecast Start")) = GetCalcStateVal(arrCalc, r, mapCalc, "Forecast Start")
        arrOut(r, mapState("Forecast Finish")) = GetCalcStateVal(arrCalc, r, mapCalc, "Forecast Finish")
        arrOut(r, mapState("Deadline")) = GetCalcStateVal(arrCalc, r, mapCalc, "Deadline")
        arrOut(r, mapState("Constraint Active")) = GetCalcStateVal(arrCalc, r, mapCalc, "Constraint Active")
        arrOut(r, mapState("Start Constraint Type")) = GetCalcStateVal(arrCalc, r, mapCalc, "Start Constraint Type")
        arrOut(r, mapState("Start Constraint Date")) = GetCalcStateVal(arrCalc, r, mapCalc, "Start Constraint Date")
        arrOut(r, mapState("Finish Constraint Type")) = GetCalcStateVal(arrCalc, r, mapCalc, "Finish Constraint Type")
        arrOut(r, mapState("Finish Constraint Date")) = GetCalcStateVal(arrCalc, r, mapCalc, "Finish Constraint Date")

        arrOut(r, mapState("Row Signature")) = BuildCalcStateRowSignature(arrCalc, r, mapCalc)
        arrOut(r, mapState("Run Status")) = UCase$(Trim$(runStatus))

    Next r

    tblState.DataBodyRange.value = arrOut
    ApplyCalcStateFormats tblState

SafeExit:
    Application.EnableEvents = True
    Application.ScreenUpdating = True

    If Err.Number <> 0 Then
        CalcState_AddStopToConsole consoleMessages, _
            "Erreur dans Write_CalcState_Snapshot : " & Err.Description, _
            "Error in Write_CalcState_Snapshot: " & Err.Description
    End If

End Sub


Private Sub CalcState_AddStopToConsole( _
    ByVal consoleMessages As Collection, _
    ByVal frText As String, _
    ByVal enText As String)

    If consoleMessages Is Nothing Then Exit Sub

    CalcBridge_AddConsoleMessage consoleMessages, "STOP", _
        BiMsg(frText, enText)

End Sub

Private Function BuildCalcStateHeaders() As Variant

    Dim arr

    arr = Array( _
        "ID", _
        "ParentID", _
        "IsSummary", _
        "Predecessors WBS", _
        "Cal", _
        "Baseline Start", _
        "Baseline Duration", _
        "Actual Start", _
        "Actual Finish", _
        "Forecast Start", _
        "Forecast Finish", _
                "Deadline", _
        "Constraint Active", _
        "Start Constraint Type", _
        "Start Constraint Date", _
        "Finish Constraint Type", _
        "Finish Constraint Date", _
        "Row Signature", _
        "Run Status")

    BuildCalcStateHeaders = arr

End Function

Private Function EnsureWorksheetExists_CalcState( _
    ByVal sheetName As String, _
    Optional ByRef wasCreated As Boolean = False) As Worksheet

    Dim ws As Worksheet
    Dim wsStart As Worksheet

    wasCreated = False
    Set wsStart = ActiveSheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = sheetName
        wasCreated = True

        If Not wsStart Is Nothing Then
            wsStart.Activate
        End If
    End If

    Set EnsureWorksheetExists_CalcState = ws

End Function

Private Function EnsureTableWithHeaders_CalcState( _
    ByVal ws As Worksheet, _
    ByVal tableName As String, _
    ByVal topLeft As Range, _
    ByVal headers As Variant, _
    Optional ByRef wasCreated As Boolean = False) As ListObject

    Dim tbl As ListObject
    Dim i As Long
    Dim headerCount As Long
    Dim rng As Range
    Dim lc As ListColumn
    Dim headerName As String

    wasCreated = False

    On Error Resume Next
    Set tbl = ws.ListObjects(tableName)
    On Error GoTo 0

    headerCount = UBound(headers) - LBound(headers) + 1

    If tbl Is Nothing Then

        Set rng = ws.Range(topLeft, topLeft.Offset(1, headerCount - 1))
        rng.Clear

        For i = 0 To headerCount - 1
            rng.Cells(1, i + 1).value = headers(LBound(headers) + i)
        Next i

        Set tbl = ws.ListObjects.Add(xlSrcRange, rng, , xlYes)
        tbl.Name = tableName
        wasCreated = True

        If Not tbl.DataBodyRange Is Nothing Then
            tbl.DataBodyRange.ClearContents
        End If

    Else

        For i = LBound(headers) To UBound(headers)
            headerName = CStr(headers(i))

            Set lc = Nothing
            On Error Resume Next
            Set lc = tbl.ListColumns(headerName)
            On Error GoTo 0

            If lc Is Nothing Then
                Set lc = tbl.ListColumns.Add
                lc.Name = headerName
            End If
        Next i

    End If

    Set EnsureTableWithHeaders_CalcState = tbl

End Function

Private Function BuildColumnMap_CalcState(ByVal tbl As ListObject) As Object

    Dim d As Object
    Dim i As Long

    Set d = CreateObject("Scripting.Dictionary")

    For i = 1 To tbl.ListColumns.Count
        d(tbl.ListColumns(i).Name) = i
    Next i

    Set BuildColumnMap_CalcState = d

End Function

Private Sub ValidateCalcStateSourceColumns(ByVal mapCalc As Object)

    Dim requiredCols As Variant
    Dim c As Variant

    requiredCols = Array( _
        "ID", _
        "ParentID", _
        "IsSummary", _
        "Predecessors WBS", _
        "Cal", _
        "Baseline Start", _
        "Baseline Duration", _
        "Actual Start", _
        "Actual Finish", _
        "Forecast Start", _
        "Forecast Finish", _
                "Deadline", _
        "Constraint Active", _
        "Start Constraint Type", _
        "Start Constraint Date", _
        "Finish Constraint Type", _
        "Finish Constraint Date" _
    )

    For Each c In requiredCols
        If Not mapCalc.Exists(CStr(c)) Then
            Err.Raise vbObjectError + 9101, _
                "ValidateCalcStateSourceColumns", _
                "Missing source column in tbl_CALC: " & CStr(c)
        End If
    Next c

End Sub

Private Sub ValidateCalcStateTargetColumns(ByVal mapState As Object)

    Dim requiredCols As Variant
    Dim c As Variant

    requiredCols = Array( _
        "ID", _
        "ParentID", _
        "IsSummary", _
        "Predecessors WBS", _
        "Cal", _
        "Baseline Start", _
        "Baseline Duration", _
        "Actual Start", _
        "Actual Finish", _
        "Forecast Start", _
        "Forecast Finish", _
                "Deadline", _
        "Constraint Active", _
        "Start Constraint Type", _
        "Start Constraint Date", _
        "Finish Constraint Type", _
        "Finish Constraint Date", _
        "Row Signature", _
        "Run Status" _
    )

    For Each c In requiredCols
        If Not mapState.Exists(CStr(c)) Then
            Err.Raise vbObjectError + 9102, _
                "ValidateCalcStateTargetColumns", _
                "Missing target column in " & CALC_STATE_TABLE_NAME & ": " & CStr(c)
        End If
    Next c

End Sub


Private Sub ResizeTableToRowCount_CalcState(ByVal tbl As ListObject, ByVal targetRows As Long)

    Dim currentRows As Long
    Dim r As Long

    If tbl.DataBodyRange Is Nothing Then
        currentRows = 0
    Else
        currentRows = tbl.ListRows.Count
    End If

    If targetRows = 0 Then
        Do While tbl.ListRows.Count > 0
            tbl.ListRows(tbl.ListRows.Count).Delete
        Loop
        Exit Sub
    End If

    If currentRows < targetRows Then
        For r = currentRows + 1 To targetRows
            tbl.ListRows.Add
        Next r
    ElseIf currentRows > targetRows Then
        For r = currentRows To targetRows + 1 Step -1
            tbl.ListRows(r).Delete
        Next r
    End If

End Sub

Private Sub ApplyCalcStateFormats(ByVal tbl As ListObject)

    On Error Resume Next

    tbl.ListColumns("ID").Range.NumberFormat = "0"
    tbl.ListColumns("ParentID").Range.NumberFormat = "0"

    tbl.ListColumns("IsSummary").Range.NumberFormat = "@"
    tbl.ListColumns("Predecessors WBS").Range.NumberFormat = "@"
    tbl.ListColumns("Cal").Range.NumberFormat = "@"
    tbl.ListColumns("Constraint Active").Range.NumberFormat = "@"
    tbl.ListColumns("Start Constraint Type").Range.NumberFormat = "@"
    tbl.ListColumns("Finish Constraint Type").Range.NumberFormat = "@"
    tbl.ListColumns("Row Signature").Range.NumberFormat = "@"
    tbl.ListColumns("Run Status").Range.NumberFormat = "@"

    tbl.ListColumns("Baseline Start").Range.NumberFormat = "dd/mm/yyyy"
    tbl.ListColumns("Actual Start").Range.NumberFormat = "dd/mm/yyyy"
    tbl.ListColumns("Actual Finish").Range.NumberFormat = "dd/mm/yyyy"
    tbl.ListColumns("Forecast Start").Range.NumberFormat = "dd/mm/yyyy"
    tbl.ListColumns("Forecast Finish").Range.NumberFormat = "dd/mm/yyyy"
        tbl.ListColumns("Deadline").Range.NumberFormat = "dd/mm/yyyy"
    tbl.ListColumns("Start Constraint Date").Range.NumberFormat = "dd/mm/yyyy"
    tbl.ListColumns("Finish Constraint Date").Range.NumberFormat = "dd/mm/yyyy"

    tbl.ListColumns("Baseline Duration").Range.NumberFormat = "0"

    On Error GoTo 0

End Sub

Private Function GetCalcStateVal( _
    ByRef arr As Variant, _
    ByVal rowIdx As Long, _
    ByVal mapCol As Object, _
    ByVal colName As String) As Variant

    If Not mapCol.Exists(colName) Then
        GetCalcStateVal = Empty
        Exit Function
    End If

    GetCalcStateVal = arr(rowIdx, mapCol(colName))

End Function

Private Function BuildCalcStateRowSignature( _
    ByRef arrCalc As Variant, _
    ByVal rowIdx As Long, _
    ByVal mapCalc As Object) As String

    Dim s As String

    s = ""
    s = s & "|ID=" & NormalizeIncrementalSignatureValue(GetCalcStateVal(arrCalc, rowIdx, mapCalc, "ID"), "TEXT")
    s = s & "|ParentID=" & NormalizeIncrementalSignatureValue(GetCalcStateVal(arrCalc, rowIdx, mapCalc, "ParentID"), "TEXT")
    s = s & "|IsSummary=" & NormalizeIncrementalSignatureValue(GetCalcStateVal(arrCalc, rowIdx, mapCalc, "IsSummary"), "BOOLEAN")
    s = s & "|PredWBS=" & NormalizeIncrementalSignatureValue(GetCalcStateVal(arrCalc, rowIdx, mapCalc, "Predecessors WBS"), "PREDWBS")
    s = s & "|Cal=" & NormalizeIncrementalSignatureValue(NormalizeCalendarType(GetCalcStateVal(arrCalc, rowIdx, mapCalc, "Cal")), "TEXT")

    s = s & "|BS=" & NormalizeIncrementalSignatureValue(GetCalcStateVal(arrCalc, rowIdx, mapCalc, "Baseline Start"), "DATE")
    s = s & "|BD=" & NormalizeIncrementalSignatureValue(GetCalcStateVal(arrCalc, rowIdx, mapCalc, "Baseline Duration"), "NUMBER")
    s = s & "|AS=" & NormalizeIncrementalSignatureValue(GetCalcStateVal(arrCalc, rowIdx, mapCalc, "Actual Start"), "DATE")
    s = s & "|AF=" & NormalizeIncrementalSignatureValue(GetCalcStateVal(arrCalc, rowIdx, mapCalc, "Actual Finish"), "DATE")
    s = s & "|FS=" & NormalizeIncrementalSignatureValue(GetCalcStateVal(arrCalc, rowIdx, mapCalc, "Forecast Start"), "DATE")
    s = s & "|FF=" & NormalizeIncrementalSignatureValue(GetCalcStateVal(arrCalc, rowIdx, mapCalc, "Forecast Finish"), "DATE")
        s = s & "|DL=" & NormalizeIncrementalSignatureValue(GetCalcStateVal(arrCalc, rowIdx, mapCalc, "Deadline"), "DATE")
    s = s & "|CActive=" & NormalizeIncrementalSignatureValue(GetCalcStateVal(arrCalc, rowIdx, mapCalc, "Constraint Active"), "BOOLEAN")
    s = s & "|SCType=" & NormalizeIncrementalSignatureValue(GetCalcStateVal(arrCalc, rowIdx, mapCalc, "Start Constraint Type"), "TEXT")
    s = s & "|SCDate=" & NormalizeIncrementalSignatureValue(GetCalcStateVal(arrCalc, rowIdx, mapCalc, "Start Constraint Date"), "DATE")
    s = s & "|FCType=" & NormalizeIncrementalSignatureValue(GetCalcStateVal(arrCalc, rowIdx, mapCalc, "Finish Constraint Type"), "TEXT")
    s = s & "|FCDate=" & NormalizeIncrementalSignatureValue(GetCalcStateVal(arrCalc, rowIdx, mapCalc, "Finish Constraint Date"), "DATE")

    BuildCalcStateRowSignature = s

End Function


Public Function NormalizeIncrementalSignatureValue( _
    ByVal v As Variant, _
    Optional ByVal valueKind As String = "TEXT") As String

    Select Case UCase$(Trim$(valueKind))

        Case "BOOLEAN"
            NormalizeIncrementalSignatureValue = NormalizeIncrementalSignatureBoolean(v)

        Case "DATE"
            NormalizeIncrementalSignatureValue = NormalizeIncrementalSignatureDate(v)

        Case "NUMBER"
            NormalizeIncrementalSignatureValue = NormalizeIncrementalSignatureNumber(v)

        Case "PREDWBS"
            NormalizeIncrementalSignatureValue = NormalizeIncrementalSignatureText(v, True)

        Case Else
            NormalizeIncrementalSignatureValue = NormalizeIncrementalSignatureText(v, False)

    End Select

End Function

Private Function NormalizeIncrementalSignatureBoolean(ByVal v As Variant) As String

    Dim s As String

    If IsEmpty(v) Then Exit Function
    If IsNull(v) Then Exit Function

    If VarType(v) = vbBoolean Then
        If CBool(v) Then
            NormalizeIncrementalSignatureBoolean = "TRUE"
        Else
            NormalizeIncrementalSignatureBoolean = "FALSE"
        End If
        Exit Function
    End If

    s = UCase$(Trim$(CStr(v)))

    Select Case s

        Case "TRUE", "VRAI", "-1", "1", "YES", "OUI"
            NormalizeIncrementalSignatureBoolean = "TRUE"

        Case "FALSE", "FAUX", "0", "NO", "NON", ""
            NormalizeIncrementalSignatureBoolean = "FALSE"

        Case Else
            NormalizeIncrementalSignatureBoolean = s

    End Select

End Function

Private Function NormalizeIncrementalSignatureDate(ByVal v As Variant) As String

    If IsEmpty(v) Then Exit Function
    If IsNull(v) Then Exit Function
    If Trim$(CStr(v)) = "" Then Exit Function

    On Error GoTo Fallback

    If IsDate(v) Then
        NormalizeIncrementalSignatureDate = Format$(CDate(v), "yyyymmdd")
        Exit Function
    End If

    If IsNumeric(v) Then
        If CDbl(v) > 0 Then
            NormalizeIncrementalSignatureDate = Format$(CDate(CDbl(v)), "yyyymmdd")
            Exit Function
        End If
    End If

Fallback:
    NormalizeIncrementalSignatureDate = Trim$(CStr(v))

End Function

Private Function NormalizeIncrementalSignatureNumber(ByVal v As Variant) As String

    If IsEmpty(v) Then Exit Function
    If IsNull(v) Then Exit Function
    If Trim$(CStr(v)) = "" Then Exit Function

    If IsNumeric(v) Then
        NormalizeIncrementalSignatureNumber = Format$(CDbl(v), "0.############")
    Else
        NormalizeIncrementalSignatureNumber = Trim$(CStr(v))
    End If

End Function

Private Function NormalizeIncrementalSignatureText( _
    ByVal v As Variant, _
    Optional ByVal removeSpaces As Boolean = False) As String

    Dim s As String

    If IsEmpty(v) Then Exit Function
    If IsNull(v) Then Exit Function

    s = Trim$(CStr(v))

    If removeSpaces Then
        s = Replace$(s, " ", "")
    End If

    NormalizeIncrementalSignatureText = s

End Function

