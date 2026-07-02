Attribute VB_Name = "mod_GanttLive"
'=====================================================
' mod_GanttLive
'=====================================================
' Rôle :
' - simulation live du Gantt
' - construction et alimentation de CALC_GANTT_TEST
' - lecture des overlays de test
' - backend de recalcul test
' - lock changes avec backup / rollback
'
' Points d'attention :
' - priorité stricte : Test > Actual > Forecast > Baseline > Dependencies
' - leaf network uniquement pour le recalcul
' - aucun fallback silencieux
' - aucune écriture durable hors Lock Changes
'=====================================================

Option Explicit

Private Const GANTT_SHEET As String = "GANTT"
Private Const WBS_SHEET As String = "WBS"
Private Const WBS_TABLE As String = "tbl_WBS"
Private Const CALC_SHEET As String = "CALC"
Private Const CALC_TABLE As String = "tbl_CALC"

Private Const CALC_GANTT_TEST_SHEET As String = "CALC_GANTT_TEST"
Private Const CALC_GANTT_TEST_TABLE As String = "tbl_CALC_GANTT_TEST"

Private Const GANTT_FIRST_TASK_ROW As Long = 4
Private Const GANTT_COL_WBS As Long = 1

Private Const COL_TEST_START As Long = 5
Private Const COL_TEST_FINISH As Long = 6
Private Const COL_TEST_PROGRESS As Long = 9

'=====================================================
' One-shot render request.
' The test engine raises this flag just before calling
' Refresh_Gantt. Refresh_Gantt consumes it and clears it.
' This avoids stale simulation staying active on later
' normal refreshes.
'=====================================================
Private gPendingRenderMode As String
Private gActiveSimulationMode As String

Public Sub GanttLive_RequestTestRender()
    gPendingRenderMode = "TEST"
End Sub

Public Sub GanttLive_RequestScenarioRender()
    gPendingRenderMode = "SCENARIO"
End Sub

Public Sub GanttLive_ClearTestRenderRequest()
    gPendingRenderMode = ""
End Sub

Public Sub GanttLive_SafeEmptyState()

    GanttLive_ClearTestRenderRequest
    GanttLive_ClearActiveSimulationMode
    ClearCalcGanttTestResults

End Sub
Public Function GanttLive_IsTestRenderRequested() As Boolean
    GanttLive_IsTestRenderRequested = (Trim$(gPendingRenderMode) <> "")
End Function

Public Function GanttLive_GetPendingRenderMode() As String
    GanttLive_GetPendingRenderMode = UCase$(Trim$(gPendingRenderMode))
End Function

Public Sub Run_Gantt_Test_Engine( _
    Optional ByVal silentMode As Boolean = False, _
    Optional ByRef transactionSucceeded As Variant, _
    Optional ByRef transactionMessages As Variant, _
    Optional ByRef transactionGanttRebuilt As Variant, _
    Optional ByVal recordSilentMessages As Boolean = True)

    Dim wsGantt As Worksheet
    Dim wsWBS As Worksheet
    Dim wsCalc As Worksheet
    Dim wsTest As Worksheet

    Dim tblWBS As ListObject
    Dim tblCalc As ListObject
    Dim tblTest As ListObject

    Dim mapWBS As Object
    Dim mapCalc As Object
    Dim mapCalcDriving As Object
    Dim wbsToId As Object
    Dim calcById As Object

    Dim dataWBS As Variant
    Dim outArr() As Variant

    Dim rowCount As Long
    Dim r As Long
    Dim testRow As Long

    Dim idVal As String
    Dim wbsVal As String
    Dim taskTypeVal As String
    Dim ganttRow As Long

    Dim baseStart As Variant
    Dim baseFinish As Variant
    Dim baseDuration As Variant
    Dim baseProgress As Variant
    Dim drivingLogic As String

    Dim baselineStart As Variant
    Dim baselineDuration As Variant
    Dim actualStart As Variant
    Dim actualFinish As Variant
    Dim forecastStart As Variant
    Dim forecastFinish As Variant

    Dim testStart As Variant
    Dim testFinish As Variant
    Dim testProgressRaw As Variant
    Dim testProgressNorm As Variant

    Dim inputStart As Variant
    Dim inputFinish As Variant
    Dim inputDuration As Variant
    Dim inputProgress As Variant

    Dim isSummary As Boolean
    Dim hasActual As Boolean
    Dim anyTestValue As Boolean
    Dim hasRenderableDelta As Boolean
    Dim hasAnyTestInput As Boolean

    Dim consoleMessages As Collection
    Dim workflowStarted As Boolean
    Dim testSucceeded As Boolean
    Dim ganttRebuilt As Boolean

    On Error GoTo SafeExit

    If Not silentMode Then workflowStarted = EnsurePlanningWorkflowStarted("Run_Gantt_Test_Engine")
    ThisWorkbook.Init_AppEvents
    Set consoleMessages = New Collection

    Set wsGantt = ThisWorkbook.Worksheets(GANTT_SHEET)
    Set wsWBS = ThisWorkbook.Worksheets(WBS_SHEET)
    Set wsCalc = ThisWorkbook.Worksheets(CALC_SHEET)

    Set tblWBS = wsWBS.ListObjects(WBS_TABLE)
    Set tblCalc = wsCalc.ListObjects(CALC_TABLE)

    If tblWBS.DataBodyRange Is Nothing Then GoTo CleanExit
    If tblCalc.DataBodyRange Is Nothing Then GoTo CleanExit

    Set wsTest = Ensure_CalcGanttTest_Sheet()
    Set tblTest = Ensure_CalcGanttTest_Table(wsTest)

    Set mapWBS = BuildColumnMap_GanttLive(tblWBS)
    Set mapCalc = BuildColumnMap_GanttLive(tblCalc)
    Set mapCalcDriving = BuildCalcDrivingLogicMap_Live()
    Set wbsToId = BuildWbsToIdMapFromWBS_Live()
    Set calcById = BuildCalcByIdMap_Live(tblCalc, mapCalc)

    ValidateGanttTestSourceColumns mapWBS, mapCalc
    ValidateCalcGanttTestColumns tblTest

    dataWBS = tblWBS.DataBodyRange.value
    rowCount = UBound(dataWBS, 1)

    ResizeTableToRowCount_Generic tblTest, rowCount
    ReDim outArr(1 To rowCount, 1 To tblTest.ListColumns.Count)

    testRow = 0

    For r = 1 To rowCount

        idVal = Trim$(CStr(dataWBS(r, mapWBS("ID"))))
        wbsVal = NormalizeWBS(CStr(dataWBS(r, mapWBS("WBS"))))
        taskTypeVal = Trim$(CStr(dataWBS(r, mapWBS("Task Type"))))

        testRow = testRow + 1

        baseStart = GetCellValue(dataWBS(r, mapWBS("Calculated Start")))
        baseFinish = GetCellValue(dataWBS(r, mapWBS("Calculated Finish")))
        baseDuration = GetCellValue(dataWBS(r, mapWBS("Calculated Duration")))
        baseProgress = GetCellValue(dataWBS(r, mapWBS("% Progress")))

        baselineStart = GetCellValue(dataWBS(r, mapWBS("Baseline Start")))
        baselineDuration = GetCellValue(dataWBS(r, mapWBS("Baseline Duration")))
        actualStart = GetCellValue(dataWBS(r, mapWBS("Actual Start")))
        actualFinish = GetCellValue(dataWBS(r, mapWBS("Actual Finish")))
        forecastStart = GetCellValue(dataWBS(r, mapWBS("Forecast Start")))
        forecastFinish = GetCellValue(dataWBS(r, mapWBS("Forecast Finish")))

        If mapCalcDriving.Exists(idVal) Then
            drivingLogic = CStr(mapCalcDriving(idVal))
        Else
            drivingLogic = ""
        End If

        isSummary = (UCase$(drivingLogic) = "SUMMARY")
        hasActual = (HasValue(actualStart) Or HasValue(actualFinish))

        testStart = Empty
        testFinish = Empty
        testProgressRaw = Empty
        testProgressNorm = Empty
        inputStart = Empty
        inputFinish = Empty
        inputDuration = Empty
        inputProgress = Empty
        anyTestValue = False

        ganttRow = FindGanttRowByWBS(wsGantt, wbsVal)

        If ganttRow > 0 Then
            testStart = GetCellValue(wsGantt.Cells(ganttRow, COL_TEST_START).value)
            testFinish = GetCellValue(wsGantt.Cells(ganttRow, COL_TEST_FINISH).value)
            testProgressRaw = GetCellValue(wsGantt.Cells(ganttRow, COL_TEST_PROGRESS).value)
        End If

        If HasValue(testStart) Or HasValue(testFinish) Or HasValue(testProgressRaw) Then
            anyTestValue = True
        End If

        If HasValue(testProgressRaw) Then
            testProgressNorm = NormalizePercentInput(testProgressRaw)
        Else
            testProgressNorm = Empty
        End If

        BuildTestInputValues_FromWBSSource _
            baselineStart, baselineDuration, actualStart, actualFinish, forecastStart, forecastFinish, _
            baseProgress, testStart, testFinish, testProgressNorm, _
            inputStart, inputFinish, inputDuration, inputProgress

        outArr(testRow, GetColumnIndex_GanttLive(tblTest, "ID")) = idVal
        outArr(testRow, GetColumnIndex_GanttLive(tblTest, "WBS")) = wbsVal
        outArr(testRow, GetColumnIndex_GanttLive(tblTest, "Task Type")) = taskTypeVal
        outArr(testRow, GetColumnIndex_GanttLive(tblTest, "Cal")) = NormalizeCalendarType(dataWBS(r, mapWBS("Cal")))
        outArr(testRow, GetColumnIndex_GanttLive(tblTest, "Is Summary")) = IIf(isSummary, "YES", "NO")
        outArr(testRow, GetColumnIndex_GanttLive(tblTest, "Parent ID")) = GetParentIdFromWBS(wbsVal, wbsToId)

        outArr(testRow, GetColumnIndex_GanttLive(tblTest, "Base Start")) = baseStart
        outArr(testRow, GetColumnIndex_GanttLive(tblTest, "Base Finish")) = baseFinish
        outArr(testRow, GetColumnIndex_GanttLive(tblTest, "Base Duration")) = baseDuration
        outArr(testRow, GetColumnIndex_GanttLive(tblTest, "Base Progress")) = baseProgress
        outArr(testRow, GetColumnIndex_GanttLive(tblTest, "Driving Logic")) = drivingLogic
        outArr(testRow, GetColumnIndex_GanttLive(tblTest, "Has Actual")) = IIf(hasActual, "YES", "NO")

        outArr(testRow, GetColumnIndex_GanttLive(tblTest, "Test Start")) = testStart
        outArr(testRow, GetColumnIndex_GanttLive(tblTest, "Test Finish")) = testFinish
        outArr(testRow, GetColumnIndex_GanttLive(tblTest, "Test % Raw")) = testProgressRaw
        outArr(testRow, GetColumnIndex_GanttLive(tblTest, "Test % Normalized")) = testProgressNorm

        outArr(testRow, GetColumnIndex_GanttLive(tblTest, "Input Start")) = inputStart
        outArr(testRow, GetColumnIndex_GanttLive(tblTest, "Input Finish")) = inputFinish
        outArr(testRow, GetColumnIndex_GanttLive(tblTest, "Input Duration")) = inputDuration
        outArr(testRow, GetColumnIndex_GanttLive(tblTest, "Input Progress")) = inputProgress

        If calcById.Exists(idVal) Then
            outArr(testRow, GetColumnIndex_GanttLive(tblTest, "Predecessors")) = calcById(idVal)
        End If

        outArr(testRow, GetColumnIndex_GanttLive(tblTest, "Lag")) = Empty
        outArr(testRow, GetColumnIndex_GanttLive(tblTest, "Any Test Value")) = IIf(anyTestValue, "YES", "NO")
        outArr(testRow, GetColumnIndex_GanttLive(tblTest, "Warning Flag")) = ""
        outArr(testRow, GetColumnIndex_GanttLive(tblTest, "Warning Text")) = ""
        outArr(testRow, GetColumnIndex_GanttLive(tblTest, "Error Flag")) = ""

    Next r

    FormatCalcGanttTestColumns tblTest
    tblTest.DataBodyRange.value = outArr
    FormatCalcGanttTestColumns tblTest

    hasAnyTestInput = GanttLive_HasAnyTestInput(tblTest)

    If Not Run_Gantt_Test_Backend(tblTest, consoleMessages) Then
        GanttLive_AbortTestEngine wsGantt
        If silentMode Then
            If recordSilentMessages Then CalcBridge_RecordPlanningMessages consoleMessages, "Run_Gantt_Test_Engine"
        Else
            CalcBridge_ShowPlanningConsole consoleMessages
        End If
        GoTo CleanExit
    End If

    hasRenderableDelta = GanttLive_HasAnyRenderableTestDelta(tblTest)

    GanttLive_ApplyTestRender wsGantt
    If Not hasAnyTestInput Then GanttLive_ClearActiveSimulationMode
    testSucceeded = True
    ganttRebuilt = True

    If silentMode Then
        If recordSilentMessages Then CalcBridge_RecordPlanningMessages consoleMessages, "Run_Gantt_Test_Engine"
        GoTo CleanExit
    End If

    If Not hasAnyTestInput Then

        GanttLive_AddBiConsoleMessage consoleMessages, "INFO", _
            "Aucune saisie TEST détectée. L’affichage simulation a été réinitialisé.", _
            "No TEST input detected. Simulation display has been reset."

    ElseIf hasRenderableDelta Then

        GanttLive_AddBiConsoleMessage consoleMessages, "INFO", _
            "Simulation exécutée.", _
            "Simulation executed."

    Else

        GanttLive_AddBiConsoleMessage consoleMessages, "INFO", _
            "Simulation exécutée, mais aucun changement visible n'a été produit." & vbCrLf & _
            "-> cause probable : priorité locale (Actual/Forecast) et/ou contrainte réseau inchangée.", _
            "Simulation executed, but no visible change was produced." & vbCrLf & _
            "-> probable cause: local priority (Actual/Forecast) and/or unchanged network constraint."

    End If

    CalcBridge_ShowPlanningConsole consoleMessages

CleanExit:
    If Not IsMissing(transactionSucceeded) Then transactionSucceeded = testSucceeded
    If Not IsMissing(transactionGanttRebuilt) Then transactionGanttRebuilt = ganttRebuilt
    If Not IsMissing(transactionMessages) Then Set transactionMessages = consoleMessages

    If workflowStarted Then
        EndPlanningWorkflow
        On Error Resume Next
        GanttDrag_ReconcileWatchState
        On Error GoTo 0
    End If
    Exit Sub

SafeExit:
    testSucceeded = False
    GanttLive_SafeExit_TestEngine

    If Err.Number <> 0 Then
        If consoleMessages Is Nothing Then Set consoleMessages = New Collection

        GanttLive_AddVbaOrStructuredError consoleMessages, "Run_Gantt_Test_Engine", Err.Description

        If silentMode Then
            If recordSilentMessages Then CalcBridge_RecordPlanningMessages consoleMessages, "Run_Gantt_Test_Engine"
        Else
            CalcBridge_ShowPlanningConsole consoleMessages
        End If
    End If

    GoTo CleanExit

End Sub

Public Function GanttLive_RunTestTransaction( _
    ByRef consoleMessages As Collection, _
    ByRef ganttRebuilt As Boolean) As Boolean

    Dim transactionSucceeded As Variant
    Dim transactionMessages As Variant
    Dim transactionGanttRebuilt As Variant

    Set consoleMessages = New Collection
    ganttRebuilt = False

    Run_Gantt_Test_Engine _
        True, _
        transactionSucceeded, _
        transactionMessages, _
        transactionGanttRebuilt, _
        False

    If IsObject(transactionMessages) Then
        Set consoleMessages = transactionMessages
    End If

    If Not IsEmpty(transactionGanttRebuilt) Then
        ganttRebuilt = CBool(transactionGanttRebuilt)
    End If

    If Not IsEmpty(transactionSucceeded) Then
        GanttLive_RunTestTransaction = CBool(transactionSucceeded)
    End If

End Function

Private Function GanttLive_HasAnyTestInput(ByVal tblTest As ListObject) As Boolean

    Dim arr As Variant
    Dim mapTest As Object
    Dim r As Long

    If tblTest Is Nothing Then Exit Function
    If tblTest.DataBodyRange Is Nothing Then Exit Function

    Set mapTest = BuildColumnMap_GanttLive(tblTest)
    arr = tblTest.DataBodyRange.value

    For r = 1 To UBound(arr, 1)
        If Trim$(CStr(arr(r, mapTest("Any Test Value")))) = "YES" Then
            GanttLive_HasAnyTestInput = True
            Exit Function
        End If
    Next r

End Function

Private Function Ensure_CalcGanttTest_Sheet() As Worksheet

    Dim ws As Worksheet
    Dim wsGantt As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(CALC_GANTT_TEST_SHEET)
    On Error GoTo 0

    If ws Is Nothing Then
        Set wsGantt = ThisWorkbook.Worksheets(GANTT_SHEET)
        Set ws = ThisWorkbook.Worksheets.Add(After:=wsGantt)
        ws.Name = CALC_GANTT_TEST_SHEET
    End If

    Set Ensure_CalcGanttTest_Sheet = ws

End Function

Private Function Ensure_CalcGanttTest_Table(ByVal ws As Worksheet) As ListObject

    Dim tbl As ListObject
    Dim headers As Variant
    Dim i As Long
    Dim expectedColCount As Long
    Dim needsRebuild As Boolean

    On Error Resume Next
    Set tbl = ws.ListObjects(CALC_GANTT_TEST_TABLE)
    On Error GoTo 0

    headers = Array( _
        "ID", "WBS", "Task Type", "Cal", "Is Summary", "Parent ID", "Base Start", "Base Finish", "Base Duration", _
        "Base Progress", "Driving Logic", "Has Actual", "Test Start", "Test Finish", "Test % Raw", _
        "Test % Normalized", "Input Start", "Input Finish", "Input Duration", "Input Progress", _
        "Predecessors", "Lag", "Any Test Value", "Calc Test Start", "Calc Test Finish", _
        "Calc Test Duration", "Calc Test Progress", "Warning Flag", "Warning Text", "Error Flag")

    expectedColCount = UBound(headers) + 1

    If Not tbl Is Nothing Then
        If tbl.ListColumns.Count <> expectedColCount Then
            needsRebuild = True
        Else
            For i = LBound(headers) To UBound(headers)
                If tbl.ListColumns(i + 1).Name <> CStr(headers(i)) Then
                    needsRebuild = True
                    Exit For
                End If
            Next i
        End If
    End If

    If tbl Is Nothing Or needsRebuild Then

        ws.Cells.Clear

        For i = LBound(headers) To UBound(headers)
            ws.Cells(1, i + 1).value = headers(i)
        Next i

        Set tbl = ws.ListObjects.Add( _
            SourceType:=xlSrcRange, _
            Source:=ws.Range(ws.Cells(1, 1), ws.Cells(2, expectedColCount)), _
            XlListObjectHasHeaders:=xlYes)

        tbl.Name = CALC_GANTT_TEST_TABLE

        If Not tbl.DataBodyRange Is Nothing Then
            tbl.DataBodyRange.ClearContents
        End If
    End If

    Set Ensure_CalcGanttTest_Table = tbl

End Function

Private Sub ValidateGanttTestSourceColumns(ByVal mapWBS As Object, ByVal mapCalc As Object)

    Dim reqWBS As Variant
    Dim reqCalc As Variant
    Dim c As Variant

    reqWBS = Array( _
        "ID", "WBS", "Task Type", "Cal", _
        "Calculated Start", "Calculated Finish", "Calculated Duration", _
        "% Progress", "Driving Logic", _
        "Actual Start", "Actual Finish", _
        "Forecast Start", "Forecast Finish", _
        "Baseline Start", "Baseline Duration")

    reqCalc = Array("ID", "Predecessors WBS", "Driving Logic")

    For Each c In reqWBS
        Call RequireMapColumn_GanttLive(mapWBS, "tbl_WBS", CStr(c), "ValidateGanttTestSourceColumns")
    Next c

    For Each c In reqCalc
        Call RequireMapColumn_GanttLive(mapCalc, "tbl_CALC", CStr(c), "ValidateGanttTestSourceColumns")
    Next c

End Sub

Private Sub ValidateCalcGanttTestColumns(ByVal tbl As ListObject)

    Dim headers As Variant
    Dim i As Long

    headers = Array( _
        "ID", "WBS", "Task Type", "Cal", "Is Summary", "Parent ID", "Base Start", "Base Finish", "Base Duration", _
        "Base Progress", "Driving Logic", "Has Actual", "Test Start", "Test Finish", "Test % Raw", _
        "Test % Normalized", "Input Start", "Input Finish", "Input Duration", "Input Progress", _
        "Predecessors", "Lag", "Any Test Value", "Calc Test Start", "Calc Test Finish", _
        "Calc Test Duration", "Calc Test Progress", "Warning Flag", "Warning Text", "Error Flag")

    For i = LBound(headers) To UBound(headers)
        Call GetColumnIndex_GanttLive(tbl, CStr(headers(i)))
    Next i

End Sub

Private Sub ResizeTableToRowCount_Generic(ByVal tbl As ListObject, ByVal targetRows As Long)

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


Private Function FindGanttRowByWBS(ByVal ws As Worksheet, ByVal wbsVal As String) As Long

    Dim lastRow As Long
    Dim r As Long

    lastRow = GetLastGanttRow_Live(ws)

    For r = GANTT_FIRST_TASK_ROW To lastRow
        If NormalizeWBS(CStr(ws.Cells(r, GANTT_COL_WBS).value)) = wbsVal Then
            FindGanttRowByWBS = r
            Exit Function
        End If
    Next r

    FindGanttRowByWBS = 0

End Function


Private Function GetParentIdFromWBS(ByVal wbsVal As String, ByVal wbsToId As Object) As String

    Dim parentWbs As String

    parentWbs = GetParentWBS(wbsVal)

    If parentWbs <> "" Then
        If wbsToId.Exists(parentWbs) Then
            GetParentIdFromWBS = CStr(wbsToId(parentWbs))
            Exit Function
        End If
    End If

    GetParentIdFromWBS = ""

End Function

Private Sub FormatCalcGanttTestColumns(ByVal tbl As ListObject)

    If tbl.DataBodyRange Is Nothing Then Exit Sub

    tbl.ListColumns("ID").DataBodyRange.NumberFormat = "@"
    tbl.ListColumns("WBS").DataBodyRange.NumberFormat = "@"
    tbl.ListColumns("Task Type").DataBodyRange.NumberFormat = "@"
    tbl.ListColumns("Cal").DataBodyRange.NumberFormat = "@"
    tbl.ListColumns("Is Summary").DataBodyRange.NumberFormat = "@"
    tbl.ListColumns("Parent ID").DataBodyRange.NumberFormat = "@"
    tbl.ListColumns("Driving Logic").DataBodyRange.NumberFormat = "@"
    tbl.ListColumns("Has Actual").DataBodyRange.NumberFormat = "@"
    tbl.ListColumns("Predecessors").DataBodyRange.NumberFormat = "@"
    tbl.ListColumns("Any Test Value").DataBodyRange.NumberFormat = "@"
    tbl.ListColumns("Warning Flag").DataBodyRange.NumberFormat = "@"
    tbl.ListColumns("Warning Text").DataBodyRange.NumberFormat = "@"
    tbl.ListColumns("Error Flag").DataBodyRange.NumberFormat = "@"
    tbl.ListColumns("Base Start").DataBodyRange.NumberFormat = "dd/mm/yyyy"
    tbl.ListColumns("Base Finish").DataBodyRange.NumberFormat = "dd/mm/yyyy"
    tbl.ListColumns("Test Start").DataBodyRange.NumberFormat = "dd/mm/yyyy"
    tbl.ListColumns("Test Finish").DataBodyRange.NumberFormat = "dd/mm/yyyy"
    tbl.ListColumns("Input Start").DataBodyRange.NumberFormat = "dd/mm/yyyy"
    tbl.ListColumns("Input Finish").DataBodyRange.NumberFormat = "dd/mm/yyyy"
    tbl.ListColumns("Calc Test Start").DataBodyRange.NumberFormat = "dd/mm/yyyy"
    tbl.ListColumns("Calc Test Finish").DataBodyRange.NumberFormat = "dd/mm/yyyy"

    tbl.ListColumns("Base Progress").DataBodyRange.NumberFormat = "0%"
    tbl.ListColumns("Test % Raw").DataBodyRange.NumberFormat = "0%"
    tbl.ListColumns("Test % Normalized").DataBodyRange.NumberFormat = "0%"
    tbl.ListColumns("Input Progress").DataBodyRange.NumberFormat = "0%"
    tbl.ListColumns("Calc Test Progress").DataBodyRange.NumberFormat = "0%"

End Sub

Private Sub RaiseMissingGanttLiveColumn( _
    ByVal sourceName As String, _
    ByVal colName As String, _
    ByVal functionName As String)

    Err.Raise vbObjectError + 913, functionName, _
        "FR:" & vbCrLf & _
        "Colonne requise introuvable pour GANTT live/scenario" & vbCrLf & vbCrLf & _
        "Table/source : " & sourceName & vbCrLf & _
        "Colonne : " & colName & vbCrLf & _
        "Fonction : " & functionName & vbCrLf & vbCrLf & _
        "EN:" & vbCrLf & _
        "Required column missing for GANTT live/scenario" & vbCrLf & vbCrLf & _
        "Source/table: " & sourceName & vbCrLf & _
        "Column: " & colName & vbCrLf & _
        "Function: " & functionName

End Sub

Private Function RequireMapColumn_GanttLive( _
    ByVal mapObj As Object, _
    ByVal sourceName As String, _
    ByVal colName As String, _
    ByVal functionName As String) As Long

    If mapObj Is Nothing Then
        RaiseMissingGanttLiveColumn sourceName, colName, functionName
    End If

    If Not mapObj.Exists(colName) Then
        RaiseMissingGanttLiveColumn sourceName, colName, functionName
    End If

    RequireMapColumn_GanttLive = CLng(mapObj(colName))

End Function

Private Function GetColumnIndex_GanttLive(ByVal tbl As ListObject, ByVal colName As String) As Long

    Dim i As Long

    For i = 1 To tbl.ListColumns.Count
        If tbl.ListColumns(i).Name = colName Then
            GetColumnIndex_GanttLive = i
            Exit Function
        End If
    Next i

    RaiseMissingGanttLiveColumn tbl.Name, colName, "GetColumnIndex_GanttLive"

End Function

Private Function GetLastGanttRow_Live(ByVal ws As Worksheet) As Long

    GetLastGanttRow_Live = ws.Cells(ws.rows.Count, GANTT_COL_WBS).End(xlUp).Row

    If GetLastGanttRow_Live < GANTT_FIRST_TASK_ROW Then
        GetLastGanttRow_Live = GANTT_FIRST_TASK_ROW - 1
    End If

End Function

Private Function BuildColumnMap_GanttLive(ByVal tbl As ListObject) As Object

    Dim d As Object
    Dim i As Long

    Set d = CreateObject("Scripting.Dictionary")

    For i = 1 To tbl.ListColumns.Count
        d(tbl.ListColumns(i).Name) = i
    Next i

    Set BuildColumnMap_GanttLive = d

End Function


Private Function BuildCalcConstraintByIdMap_GanttLive() As Object

    Dim d As Object
    Dim wsCalc As Worksheet
    Dim tblCalc As ListObject
    Dim mapCalc As Object
    Dim arr As Variant
    Dim r As Long
    Dim idVal As String
    Dim req As Variant
    Dim c As Variant

    Set d = CreateObject("Scripting.Dictionary")
    Set wsCalc = ThisWorkbook.Worksheets(CALC_SHEET)
    Set tblCalc = wsCalc.ListObjects(CALC_TABLE)

    If tblCalc.DataBodyRange Is Nothing Then
        Set BuildCalcConstraintByIdMap_GanttLive = d
        Exit Function
    End If

    Set mapCalc = BuildColumnMap_GanttLive(tblCalc)
    req = Array( _
        "ID", _
        "Task Name", _
        "Constraint Active", _
        "Start Constraint Type", _
        "Start Constraint Date", _
        "Finish Constraint Type", _
        "Finish Constraint Date" _
    )

    For Each c In req
        Call RequireMapColumn_GanttLive(mapCalc, "tbl_CALC", CStr(c), "BuildCalcConstraintByIdMap_GanttLive")
    Next c

    arr = tblCalc.DataBodyRange.value

    For r = 1 To UBound(arr, 1)
        idVal = Trim$(CStr(arr(r, mapCalc("ID"))))

        If idVal <> "" Then
            d(idVal) = Array( _
                arr(r, mapCalc("Task Name")), _
                arr(r, mapCalc("Constraint Active")), _
                arr(r, mapCalc("Start Constraint Type")), _
                GetCellValue(arr(r, mapCalc("Start Constraint Date"))), _
                arr(r, mapCalc("Finish Constraint Type")), _
                GetCellValue(arr(r, mapCalc("Finish Constraint Date"))) _
            )
        End If
    Next r

    Set BuildCalcConstraintByIdMap_GanttLive = d

End Function

Private Function BuildTaskNameByIdFromWbs_Live() As Object

    Dim d As Object
    Dim wsWBS As Worksheet
    Dim tblWBS As ListObject
    Dim mapWBS As Object
    Dim arr As Variant
    Dim r As Long
    Dim idVal As String

    Set d = CreateObject("Scripting.Dictionary")
    Set wsWBS = ThisWorkbook.Worksheets(WBS_SHEET)
    Set tblWBS = wsWBS.ListObjects(WBS_TABLE)

    If tblWBS.DataBodyRange Is Nothing Then
        Set BuildTaskNameByIdFromWbs_Live = d
        Exit Function
    End If

    Set mapWBS = BuildColumnMap_GanttLive(tblWBS)

    Call RequireMapColumn_GanttLive(mapWBS, "tbl_WBS", "ID", "BuildTaskNameByIdFromWbs_Live")
    Call RequireMapColumn_GanttLive(mapWBS, "tbl_WBS", "Task Name", "BuildTaskNameByIdFromWbs_Live")

    arr = tblWBS.DataBodyRange.value

    For r = 1 To UBound(arr, 1)
        idVal = Trim$(CStr(arr(r, mapWBS("ID"))))
        If idVal <> "" Then d(idVal) = arr(r, mapWBS("Task Name"))
    Next r

    Set BuildTaskNameByIdFromWbs_Live = d

End Function

Private Function BuildCalcDrivingLogicMap_Live() As Object

    Dim wsCalc As Worksheet
    Dim tblCalc As ListObject
    Dim mapCalc As Object
    Dim arr As Variant
    Dim d As Object
    Dim r As Long
    Dim idVal As String

    Set d = CreateObject("Scripting.Dictionary")
    Set wsCalc = ThisWorkbook.Worksheets(CALC_SHEET)
    Set tblCalc = wsCalc.ListObjects(CALC_TABLE)

    If tblCalc.DataBodyRange Is Nothing Then
        Set BuildCalcDrivingLogicMap_Live = d
        Exit Function
    End If

    Set mapCalc = BuildColumnMap_GanttLive(tblCalc)
    arr = tblCalc.DataBodyRange.value

    For r = 1 To UBound(arr, 1)
        idVal = Trim$(CStr(arr(r, mapCalc("ID"))))
        If idVal <> "" Then
            d(idVal) = UCase$(Trim$(CStr(arr(r, mapCalc("Driving Logic")))))
        End If
    Next r

    Set BuildCalcDrivingLogicMap_Live = d

End Function

Private Function BuildWbsToIdMapFromWBS_Live() As Object

    Dim wsWBS As Worksheet
    Dim tblWBS As ListObject
    Dim mapWBS As Object
    Dim arr As Variant
    Dim d As Object
    Dim r As Long
    Dim wbsVal As String
    Dim idVal As String

    Set d = CreateObject("Scripting.Dictionary")
    Set wsWBS = ThisWorkbook.Worksheets(WBS_SHEET)
    Set tblWBS = wsWBS.ListObjects(WBS_TABLE)

    If tblWBS.DataBodyRange Is Nothing Then
        Set BuildWbsToIdMapFromWBS_Live = d
        Exit Function
    End If

    Set mapWBS = BuildColumnMap_GanttLive(tblWBS)
    arr = tblWBS.DataBodyRange.value

    For r = 1 To UBound(arr, 1)
        wbsVal = NormalizeWBS(CStr(arr(r, mapWBS("WBS"))))
        idVal = Trim$(CStr(arr(r, mapWBS("ID"))))

        If wbsVal <> "" And idVal <> "" Then
            d(wbsVal) = idVal
        End If
    Next r

    Set BuildWbsToIdMapFromWBS_Live = d

End Function

Private Function BuildCalcByIdMap_Live(ByVal tblCalc As ListObject, ByVal mapCalc As Object) As Object

    Dim d As Object
    Dim arr As Variant
    Dim r As Long
    Dim idVal As String
    Dim colId As Long
    Dim colPred As Long

    Set d = CreateObject("Scripting.Dictionary")

    colId = RequireMapColumn_GanttLive(mapCalc, "tbl_CALC", "ID", "BuildCalcByIdMap_Live")
    colPred = RequireMapColumn_GanttLive(mapCalc, "tbl_CALC", "Predecessors WBS", "BuildCalcByIdMap_Live")

    If tblCalc.DataBodyRange Is Nothing Then
        Set BuildCalcByIdMap_Live = d
        Exit Function
    End If

    arr = tblCalc.DataBodyRange.value

    For r = 1 To UBound(arr, 1)
        idVal = Trim$(CStr(arr(r, colId)))
        If idVal <> "" Then
            d(idVal) = Trim$(CStr(arr(r, colPred)))
        End If
    Next r

    Set BuildCalcByIdMap_Live = d

End Function

Private Function Run_Gantt_Test_Backend( _
    ByVal tblTest As ListObject, _
    Optional ByVal consoleMessages As Variant) As Boolean

    Dim dataCore As Variant
    Dim mapCore As Object
    Dim linksBySuccId As Object

    Dim outStart() As Variant
    Dim outFinish() As Variant
    Dim outDuration() As Variant
    Dim outProgress() As Variant
    Dim outWarnFlag() As Variant
    Dim outWarnText() As Variant
    Dim outErrFlag() As Variant

    Dim idToRowTest As Object
    Dim idToWbs As Object
    Dim errorIds As Object
    Dim rootErrorIds As Object
    Dim dependencyDiagnostics As Object
    Dim constraintDiagnostics As Object
    Dim cascadeDiagnostics As Object
    Dim warningActualIds As Object
    Dim constraintMessages As Collection

    Dim rowCount As Long
    Dim r As Long
    Dim idVal As String
    Dim errorMsg As String

    On Error GoTo SafeExit

    Run_Gantt_Test_Backend = False

    If tblTest Is Nothing Then Exit Function
    If tblTest.DataBodyRange Is Nothing Then
        Run_Gantt_Test_Backend = True
        Exit Function
    End If

    rowCount = tblTest.ListRows.Count

    ReDim outStart(1 To rowCount, 1 To 1)
    ReDim outFinish(1 To rowCount, 1 To 1)
    ReDim outDuration(1 To rowCount, 1 To 1)
    ReDim outProgress(1 To rowCount, 1 To 1)
    ReDim outWarnFlag(1 To rowCount, 1 To 1)
    ReDim outWarnText(1 To rowCount, 1 To 1)
    ReDim outErrFlag(1 To rowCount, 1 To 1)

    Set idToRowTest = CreateObject("Scripting.Dictionary")
    Set idToWbs = CreateObject("Scripting.Dictionary")
    Set errorIds = CreateObject("Scripting.Dictionary")
    Set rootErrorIds = CreateObject("Scripting.Dictionary")
    Set dependencyDiagnostics = CreateObject("Scripting.Dictionary")
    Set constraintDiagnostics = CreateObject("Scripting.Dictionary")
    Set cascadeDiagnostics = CreateObject("Scripting.Dictionary")
    Set warningActualIds = CreateObject("Scripting.Dictionary")

    If GanttLive_HasConsoleCollection(consoleMessages) Then
        Set constraintMessages = consoleMessages
    Else
        Set constraintMessages = New Collection
    End If

    If Not Sync_Constraints_To_CALC_ForWorkflow(constraintMessages) Then Exit Function

    BuildGanttTestCoreDataset _
        tblTest, dataCore, mapCore, idToRowTest, idToWbs, _
        outWarnFlag, outWarnText, warningActualIds

    Set linksBySuccId = BuildCoreLinksBySucc_FromLogicLinksTable_Expanded( _
        ThisWorkbook.Worksheets(CALC_SHEET).ListObjects(CALC_TABLE))

    Run_Calc_Core dataCore, mapCore, linksBySuccId, , dependencyDiagnostics, constraintDiagnostics, cascadeDiagnostics

    For r = 1 To rowCount

        idVal = Trim$(CStr(dataCore(r, mapCore("ID"))))
        If idVal = "" Then GoTo NextRow

        outStart(r, 1) = dataCore(r, mapCore("Calculated Start"))
        outFinish(r, 1) = dataCore(r, mapCore("Calculated Finish"))
        outDuration(r, 1) = dataCore(r, mapCore("Calculated Duration"))
        outProgress(r, 1) = dataCore(r, mapCore("Input Progress"))

        If UCase$(Trim$(CStr(dataCore(r, mapCore("Error flag"))))) = "ERROR" Then

            outErrFlag(r, 1) = "ERROR"
            errorIds(idVal) = True

            errorMsg = Trim$(CStr(dataCore(r, mapCore("ErrorMsg"))))

            If errorMsg <> "" Then
                outWarnText(r, 1) = errorMsg
            End If

            If Not GanttLive_IsInheritedCoreError(errorMsg) Then
                rootErrorIds(idVal) = True
            End If

        Else
            outErrFlag(r, 1) = ""
        End If

NextRow:
    Next r

    GanttLive_RemoveDerivedLOERootErrors dataCore, mapCore, errorIds, rootErrorIds

    GanttLive_ApplyActualImpactWarnings idToRowTest, warningActualIds, outWarnFlag, outWarnText

    tblTest.ListColumns("Calc Test Start").DataBodyRange.value = outStart
    tblTest.ListColumns("Calc Test Finish").DataBodyRange.value = outFinish
    tblTest.ListColumns("Calc Test Duration").DataBodyRange.value = outDuration
    tblTest.ListColumns("Calc Test Progress").DataBodyRange.value = outProgress
    tblTest.ListColumns("Warning Flag").DataBodyRange.value = outWarnFlag
    tblTest.ListColumns("Warning Text").DataBodyRange.value = outWarnText
    tblTest.ListColumns("Error Flag").DataBodyRange.value = outErrFlag

    FormatCalcGanttTestColumns tblTest

    If errorIds.Count > 0 Then
        If rootErrorIds.Count > 0 Then
            CalcBridge_AppendCoreErrorMessagesFromData consoleMessages, dataCore, mapCore, rootErrorIds, "TEST", dependencyDiagnostics, constraintDiagnostics, cascadeDiagnostics
        Else
            ShowGanttLiveGroupedMessage errorIds, idToWbs, _
                "Erreur de calcul dans le moteur live", "corriger les valeurs test ou la logique amont", _
                "Calculation error in live engine", "fix test values or upstream logic", vbCritical, consoleMessages
        End If

        Exit Function
    End If

    Run_Gantt_Test_Backend = True
    Exit Function

SafeExit:
    If Err.Number <> 0 Then
        GanttLive_AddVbaOrStructuredError consoleMessages, "Run_Gantt_Test_Backend", Err.Description
    End If

End Function

Private Sub ShowGanttLiveGroupedMessage( _
    ByVal idsDict As Object, _
    ByVal idToWbs As Object, _
    ByVal frProblem As String, _
    ByVal frAction As String, _
    ByVal enProblem As String, _
    ByVal enAction As String, _
    ByVal boxStyle As VbMsgBoxStyle, _
    Optional ByVal consoleMessages As Variant)

    Dim msgType As String
    Dim localMessages As Collection

    If idsDict Is Nothing Then Exit Sub
    If idsDict.Count = 0 Then Exit Sub

    msgType = GanttLive_MessageTypeFromMsgBoxStyle(boxStyle)

    If GanttLive_HasConsoleCollection(consoleMessages) Then
        CalcBridge_AddConsoleMessage consoleMessages, msgType, _
            BuildGanttLiveGroupedMessage(idsDict, idToWbs, frProblem, frAction, enProblem, enAction)
    Else
        Set localMessages = New Collection
        CalcBridge_AddConsoleMessage localMessages, msgType, _
            BuildGanttLiveGroupedMessage(idsDict, idToWbs, frProblem, frAction, enProblem, enAction)
        CalcBridge_ShowPlanningConsole localMessages
    End If

End Sub

Private Function BuildGanttLiveGroupedMessage( _
    ByVal idsDict As Object, _
    ByVal idToWbs As Object, _
    ByVal frProblem As String, _
    ByVal frAction As String, _
    ByVal enProblem As String, _
    ByVal enAction As String) As String

    Dim idsLine As String
    Dim wbsLine As String

    idsLine = BuildInlineList_GanttLive(idsDict, 20)
    wbsLine = BuildInlineWBSList_GanttLive(idsDict, idToWbs, 20)

    BuildGanttLiveGroupedMessage = _
        "FR:" & vbCrLf & _
        frProblem & vbCrLf & _
        "-> " & frAction & vbCrLf & vbCrLf & _
        "IDs : " & idsLine & vbCrLf & _
        "WBS : " & wbsLine & vbCrLf & vbCrLf & _
        "EN:" & vbCrLf & _
        enProblem & vbCrLf & _
        "-> " & enAction & vbCrLf & vbCrLf & _
        "IDs: " & idsLine & vbCrLf & _
        "WBS: " & wbsLine

End Function

Private Function BuildInlineList_GanttLive(ByVal idsDict As Object, ByVal maxItems As Long) As String

    Dim result As String
    Dim key As Variant
    Dim countShown As Long
    Dim totalCount As Long

    result = ""
    countShown = 0
    totalCount = idsDict.Count

    For Each key In idsDict.Keys
        countShown = countShown + 1
        If countShown <= maxItems Then
            If result <> "" Then result = result & " / "
            result = result & CStr(key)
        Else
            Exit For
        End If
    Next key

    If totalCount > maxItems Then
        result = result & " / +" & CStr(totalCount - maxItems)
    End If

    BuildInlineList_GanttLive = result

End Function

Private Function BuildInlineWBSList_GanttLive(ByVal idsDict As Object, ByVal idToWbs As Object, ByVal maxItems As Long) As String

    Dim result As String
    Dim key As Variant
    Dim countShown As Long
    Dim totalCount As Long
    Dim itemText As String

    result = ""
    countShown = 0
    totalCount = idsDict.Count

    For Each key In idsDict.Keys
        countShown = countShown + 1
        If countShown <= maxItems Then
            If idToWbs.Exists(CStr(key)) Then
                itemText = NormalizeWBS(CStr(idToWbs(CStr(key))))
            Else
                itemText = "-"
            End If

            If result <> "" Then result = result & " / "
            result = result & itemText
        Else
            Exit For
        End If
    Next key

    If totalCount > maxItems Then
        result = result & " / +" & CStr(totalCount - maxItems)
    End If

    BuildInlineWBSList_GanttLive = result

End Function

'=====================================================
' PUBLIC HELPERS FOR mod_Gantt RENDERER
'=====================================================

Public Function GanttLive_BuildTestByIdMap() As Object

    Dim perfScope As clsPerfScope

    Dim d As Object
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim mapTest As Object
    Dim arr As Variant
    Dim r As Long
    Dim idVal As String

    Set perfScope = Profiler_BeginScope("GanttLive_BuildTestByIdMap", "Excel Read")

    Set d = CreateObject("Scripting.Dictionary")

    On Error GoTo SafeExit

    Set ws = ThisWorkbook.Worksheets(CALC_GANTT_TEST_SHEET)
    Set tbl = ws.ListObjects(CALC_GANTT_TEST_TABLE)

    If tbl.DataBodyRange Is Nothing Then
        Set GanttLive_BuildTestByIdMap = d
        Exit Function
    End If

    Set mapTest = BuildColumnMap_GanttLive(tbl)
    arr = tbl.DataBodyRange.value

    For r = 1 To UBound(arr, 1)
        idVal = Trim$(CStr(arr(r, mapTest("ID"))))

        If idVal <> "" Then
            d(idVal) = Array( _
                GetCellValue(arr(r, mapTest("Calc Test Start"))), _
                GetCellValue(arr(r, mapTest("Calc Test Finish"))), _
                GetCellValue(arr(r, mapTest("Calc Test Duration"))), _
                GetCellValue(arr(r, mapTest("Calc Test Progress"))), _
                UCase$(Trim$(CStr(arr(r, mapTest("Is Summary"))))), _
                Trim$(CStr(arr(r, mapTest("Error Flag")))), _
                Trim$(CStr(arr(r, mapTest("Any Test Value")))) _
            )
        End If
    Next r

SafeExit:
    Set GanttLive_BuildTestByIdMap = d

End Function

Public Function GanttLive_BuildBaseByIdMap() As Object

    Dim perfScope As clsPerfScope

    Dim d As Object
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim mapWBS As Object
    Dim arr As Variant
    Dim r As Long
    Dim idVal As String

    Set perfScope = Profiler_BeginScope("GanttLive_BuildBaseByIdMap", "Excel Read")

    Set d = CreateObject("Scripting.Dictionary")

    On Error GoTo SafeExit

    Set ws = ThisWorkbook.Worksheets(WBS_SHEET)
    Set tbl = ws.ListObjects(WBS_TABLE)

    If tbl.DataBodyRange Is Nothing Then
        Set GanttLive_BuildBaseByIdMap = d
        Exit Function
    End If

    Set mapWBS = BuildColumnMap_GanttLive(tbl)
    arr = tbl.DataBodyRange.value

    For r = 1 To UBound(arr, 1)
        idVal = Trim$(CStr(arr(r, mapWBS("ID"))))

        If idVal <> "" Then
            d(idVal) = Array( _
                GetCellValue(arr(r, mapWBS("Calculated Start"))), _
                GetCellValue(arr(r, mapWBS("Calculated Finish"))), _
                GetCellValue(arr(r, mapWBS("Calculated Duration"))), _
                GetCellValue(arr(r, mapWBS("% Progress"))), _
                NormalizeCalendarType(arr(r, mapWBS("Cal"))) _
            )
        End If
    Next r

SafeExit:
    Set GanttLive_BuildBaseByIdMap = d

End Function

Public Function GanttLive_HasRenderableTestDelta(ByVal idVal As String, ByVal baseById As Object, ByVal testById As Object) As Boolean

    Dim baseData As Variant
    Dim testData As Variant

    If idVal = "" Then Exit Function
    If Not baseById.Exists(idVal) Then Exit Function
    If Not testById.Exists(idVal) Then Exit Function

    baseData = baseById(idVal)
    testData = testById(idVal)

    ' Pas de contour si la ligne test est en erreur
    If Trim$(CStr(testData(5))) <> "" Then Exit Function

    ' Contour si le rendu simulé affiché diffère du rendu de base
    If ValuesDiffer(baseData(0), testData(0)) Then
        GanttLive_HasRenderableTestDelta = True
        Exit Function
    End If

    If ValuesDiffer(baseData(1), testData(1)) Then
        GanttLive_HasRenderableTestDelta = True
        Exit Function
    End If

    If ValuesDiffer(baseData(3), testData(3)) Then
        GanttLive_HasRenderableTestDelta = True
        Exit Function
    End If

End Function

Public Function GanttLive_GetDisplayStart(ByVal idVal As String, ByVal baseById As Object, ByVal testById As Object, ByVal isTestMode As Boolean) As Variant

    If isTestMode Then
        If testById.Exists(idVal) Then
            If Trim$(CStr(testById(idVal)(5))) = "" Then
                If HasValue(testById(idVal)(0)) Then
                    GanttLive_GetDisplayStart = testById(idVal)(0)
                    Exit Function
                End If
            End If
        End If
    End If

    If baseById.Exists(idVal) Then
        GanttLive_GetDisplayStart = baseById(idVal)(0)
    Else
        GanttLive_GetDisplayStart = Empty
    End If

End Function

Public Function GanttLive_GetDisplayFinish(ByVal idVal As String, ByVal baseById As Object, ByVal testById As Object, ByVal isTestMode As Boolean) As Variant

    If isTestMode Then
        If testById.Exists(idVal) Then
            If Trim$(CStr(testById(idVal)(5))) = "" Then
                If HasValue(testById(idVal)(1)) Then
                    GanttLive_GetDisplayFinish = testById(idVal)(1)
                    Exit Function
                End If
            End If
        End If
    End If

    If baseById.Exists(idVal) Then
        GanttLive_GetDisplayFinish = baseById(idVal)(1)
    Else
        GanttLive_GetDisplayFinish = Empty
    End If

End Function

Public Function GanttLive_GetDisplayDuration(ByVal idVal As String, ByVal baseById As Object, ByVal testById As Object, ByVal isTestMode As Boolean) As Variant

    If isTestMode Then
        If testById.Exists(idVal) Then
            If Trim$(CStr(testById(idVal)(5))) = "" Then
                If HasValue(testById(idVal)(2)) Then
                    GanttLive_GetDisplayDuration = testById(idVal)(2)
                    Exit Function
                End If
            End If
        End If
    End If

    If baseById.Exists(idVal) Then
        GanttLive_GetDisplayDuration = baseById(idVal)(2)
    Else
        GanttLive_GetDisplayDuration = Empty
    End If

End Function

Public Function GanttLive_GetDisplayProgress(ByVal idVal As String, ByVal baseById As Object, ByVal testById As Object, ByVal isTestMode As Boolean) As Variant

    If isTestMode Then
        If testById.Exists(idVal) Then
            If Trim$(CStr(testById(idVal)(5))) = "" Then
                If HasValue(testById(idVal)(3)) Then
                    GanttLive_GetDisplayProgress = testById(idVal)(3)
                    Exit Function
                End If
            End If
        End If
    End If

    If baseById.Exists(idVal) Then
        GanttLive_GetDisplayProgress = baseById(idVal)(3)
    Else
        GanttLive_GetDisplayProgress = Empty
    End If

End Function

'=====================================================
' Lock button
'=====================================================
Public Sub Run_Gantt_Lock_Changes()

    Dim wsWBS As Worksheet
    Dim wsGantt As Worksheet
    Dim wsCalc As Worksheet
    Dim wsTest As Worksheet

    Dim tblWBS As ListObject
    Dim tblCalc As ListObject
    Dim tblTest As ListObject

    Dim mapWBS As Object
    Dim mapCalc As Object

    Dim wbsBackup As Object
    Dim ganttTestBackup As Object
    Dim appliedChanges As Object
    Dim simulatedById As Object

    Dim hasCalcErrors As Boolean
    Dim lockMatches As Boolean
    Dim consoleMessages As Collection
    Dim workflowStarted As Boolean

    On Error GoTo SafeExit

    workflowStarted = EnsurePlanningWorkflowStarted("Run_Gantt_Lock_Changes")
    ThisWorkbook.Init_AppEvents
    Set consoleMessages = New Collection

    Set wsWBS = ThisWorkbook.Worksheets(WBS_SHEET)
    Set wsGantt = ThisWorkbook.Worksheets(GANTT_SHEET)
    Set wsCalc = ThisWorkbook.Worksheets(CALC_SHEET)

    Set tblWBS = wsWBS.ListObjects(WBS_TABLE)
    Set tblCalc = wsCalc.ListObjects(CALC_TABLE)

    If tblWBS.DataBodyRange Is Nothing Then GoTo CleanExit
    If tblCalc.DataBodyRange Is Nothing Then GoTo CleanExit

    Set mapWBS = BuildColumnMap_GanttLive(tblWBS)
    Set mapCalc = BuildColumnMap_GanttLive(tblCalc)

    ValidateLockSourceColumns mapWBS, mapCalc

    If GanttLive_IsScenarioActive() Then
        CreateScenarioPlanningFromCurrentScenario
        GoTo CleanExit
    End If

    Set appliedChanges = BuildModifiedTestChangesMap(wsGantt, tblWBS, mapWBS)

    If appliedChanges.Count = 0 Then
        GanttLive_AddBiConsoleMessage consoleMessages, "WARNING", _
            "Aucune modification test à verrouiller.", _
            "No test changes to lock."

        CalcBridge_ShowPlanningConsole consoleMessages
        GoTo CleanExit
    End If

    Run_Gantt_Test_Engine True

    Set wsTest = ThisWorkbook.Worksheets(CALC_GANTT_TEST_SHEET)
    Set tblTest = wsTest.ListObjects(CALC_GANTT_TEST_TABLE)

    If GanttLive_CalcGanttTestHasErrors(tblTest) Then
        GanttLive_AddBiConsoleMessage consoleMessages, "WARNING", _
            "Lock annulé : la simulation TEST préalable contient des erreurs." & vbCrLf & _
            "-> corriger les valeurs test ou la logique amont avant de verrouiller.", _
            "Lock cancelled: the preliminary TEST simulation contains errors." & vbCrLf & _
            "-> fix test values or upstream logic before locking."

        CalcBridge_ShowPlanningConsole consoleMessages
        GoTo CleanExit
    End If

    Set appliedChanges = BuildModifiedTestChangesMap(wsGantt, tblWBS, mapWBS)

    If appliedChanges.Count = 0 Then
        GanttLive_AddBiConsoleMessage consoleMessages, "WARNING", _
            "Aucune modification test à verrouiller.", _
            "No test changes to lock."

        CalcBridge_ShowPlanningConsole consoleMessages
        GoTo CleanExit
    End If

    Set simulatedById = BuildSimulatedLockResultMap(appliedChanges)

    If simulatedById.Count = 0 Then
        GanttLive_AddBiConsoleMessage consoleMessages, "WARNING", _
            "Lock annulé : aucun résultat simulé exploitable n'a été trouvé après le refresh TEST.", _
            "Lock cancelled: no usable simulated result was found after TEST refresh."

        CalcBridge_ShowPlanningConsole consoleMessages
        GoTo CleanExit
    End If

    Set wbsBackup = BackupWBSLockColumns(tblWBS, mapWBS)
    Set ganttTestBackup = BackupGanttTestInputs(wsGantt)

    ApplyModifiedTestChangesToWBS tblWBS, mapWBS, appliedChanges

    GanttLive_ClearTestRenderRequest
    Run_Planning_Update
    DrainPlanningWorkflowDeferredDisplayMessages consoleMessages

    hasCalcErrors = CalcTableHasErrors(tblCalc, mapCalc)
    lockMatches = LockResultsMatchSimulatedResult(tblWBS, mapWBS, simulatedById)

    If hasCalcErrors Or (Not lockMatches) Then

        GanttLive_RollbackFailedLock _
            tblWBS, mapWBS, wsGantt, wbsBackup, ganttTestBackup

        If hasCalcErrors Then
            GanttLive_AddBiConsoleMessage consoleMessages, "WARNING", _
                "Lock annulé : le calcul a détecté des erreurs. Les valeurs WBS d'origine ont été restaurées et les colonnes test ont été conservées.", _
                "Lock cancelled: calculation found errors. Original WBS values were restored and test inputs were preserved."
        Else
            GanttLive_AddBiConsoleMessage consoleMessages, "WARNING", _
                "Lock annulé : le recalcul réel ne correspond pas au résultat simulé retenu. Les valeurs WBS d'origine ont été restaurées et les colonnes test ont été conservées.", _
                "Lock cancelled: the real recalculation does not match the retained simulated result. Original WBS values were restored and test inputs were preserved."
        End If

        CalcBridge_ShowPlanningConsole consoleMessages
        GoTo CleanExit
    End If

    GanttLive_FinalizeSuccessfulLock wsGantt, appliedChanges

    GanttLive_AddBiConsoleMessage consoleMessages, "INFO", _
        "Lock appliqué avec succès.", _
        "Lock successfully applied."

    CalcBridge_ShowPlanningConsole consoleMessages

CleanExit:
    If workflowStarted Then EndPlanningWorkflow
    Exit Sub

SafeExit:
    If consoleMessages Is Nothing Then Set consoleMessages = New Collection

    GanttLive_AddBiConsoleMessage consoleMessages, "STOP", _
        "Erreur VBA dans Run_Gantt_Lock_Changes" & vbCrLf & _
        "-> vérifier le dernier bloc modifié dans mod_GanttLive", _
        "VBA error in Run_Gantt_Lock_Changes" & vbCrLf & _
        "-> check the last edited block in mod_GanttLive"

    CalcBridge_ShowPlanningConsole consoleMessages
    Resume CleanExit

End Sub


Private Sub CreateScenarioPlanningFromCurrentScenario()

    Dim answer As VbMsgBoxResult
    Dim newPath As String
    Dim newWb As Workbook
    Dim oldScreenUpdating As Boolean
    Dim oldAlerts As Boolean
    Dim oldEvents As Boolean
    Dim macroName As String

    answer = MsgBox( _
        "Vous êtes actuellement en mode scénario." & vbCrLf & vbCrLf & _
        "Le lock direct n'est pas autorisé en mode scénario." & vbCrLf & vbCrLf & _
        "Voulez-vous créer un nouveau planning scénario basé sur l'état calculé actuel ?" & vbCrLf & vbCrLf & _
        "Le nouveau fichier :" & vbCrLf & _
        "* utilisera le scénario actuel comme nouvelle baseline ;" & vbCrLf & _
        "* conservera les % Progress du scénario ;" & vbCrLf & _
        "* videra Actual et Forecast ;" & vbCrLf & _
        "* désactivera les contraintes ;" & vbCrLf & _
        "* videra l'historique et les ACK ;" & vbCrLf & _
        "* sortira du mode scénario.", _
        vbQuestion + vbYesNo, _
        "Créer un planning scénario")

    If answer <> vbYes Then Exit Sub

    If Trim$(ThisWorkbook.Path) = "" Then
        MsgBox "Le fichier source doit être enregistré avant de créer un planning scénario.", vbExclamation, "Créer un planning scénario"
        Exit Sub
    End If

    oldScreenUpdating = Application.ScreenUpdating
    oldAlerts = Application.DisplayAlerts
    oldEvents = Application.EnableEvents

    On Error GoTo Fail

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False

    newPath = BuildScenarioPlanningCopyPath()
    ThisWorkbook.SaveCopyAs newPath

    Set newWb = Application.Workbooks.Open(newPath, UpdateLinks:=0, ReadOnly:=False)
    Application.EnableEvents = True

    macroName = "'" & Replace(newWb.Name, "'", "''") & "'!InitializeScenarioPlanningCopyFromCurrentWorkbook"
    Application.Run macroName

    newWb.Activate
    Application.ScreenUpdating = oldScreenUpdating
    Application.DisplayAlerts = oldAlerts
    Application.EnableEvents = oldEvents
    Exit Sub

Fail:
    Application.ScreenUpdating = oldScreenUpdating
    Application.DisplayAlerts = oldAlerts
    Application.EnableEvents = oldEvents
    MsgBox "Erreur pendant la création du planning scénario :" & vbCrLf & Err.Description, vbCritical, "Créer un planning scénario"

End Sub

Private Function BuildScenarioPlanningCopyPath() As String

    Dim folderPath As String
    Dim fileName As String
    Dim baseName As String
    Dim extName As String
    Dim dotPos As Long
    Dim candidate As String
    Dim suffix As Long
    Dim stamp As String

    folderPath = ThisWorkbook.Path
    fileName = ThisWorkbook.Name
    dotPos = InStrRev(fileName, ".")
    stamp = Format$(Now, "yyyymmdd_hhnn")

    If dotPos > 0 Then
        baseName = Left$(fileName, dotPos - 1)
        extName = Mid$(fileName, dotPos)
    Else
        baseName = fileName
        extName = ".xlsm"
    End If

    candidate = folderPath & Application.PathSeparator & baseName & "_SCENARIO_" & stamp & extName
    suffix = 1

    Do While Len(Dir$(candidate)) > 0
        suffix = suffix + 1
        candidate = folderPath & Application.PathSeparator & baseName & "_SCENARIO_" & stamp & "_" & Format$(suffix, "00") & extName
    Loop

    BuildScenarioPlanningCopyPath = candidate

End Function

Public Sub InitializeScenarioPlanningCopyFromCurrentWorkbook()

    Dim oldScreenUpdating As Boolean
    Dim oldEvents As Boolean
    Dim oldAlerts As Boolean
    Dim inputGuardStarted As Boolean
    Dim forcedUpdateGuardStarted As Boolean

    On Error GoTo Fail

    oldScreenUpdating = Application.ScreenUpdating
    oldEvents = Application.EnableEvents
    oldAlerts = Application.DisplayAlerts

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False

    BeginAuthorizedWBSWrite "ScenarioFork", Array("Baseline Start", "Baseline Duration", "Forecast Start", "Forecast Finish", "Actual Start", "Actual Finish", "% Progress")
    inputGuardStarted = True
    Application.EnableEvents = True
    ApplyScenarioBaselineToWBS_CurrentWorkbook
    Application.EnableEvents = False
    EndAuthorizedWBSWrite
    inputGuardStarted = False

    DeactivateScenarioConstraints_CurrentWorkbook
    ClearScenarioForkRuntimeTables_CurrentWorkbook
    Gantt_Clear_Test_State
    ClearCalcGanttTestResults
    GanttLive_ClearTestRenderRequest
    GanttLive_ClearActiveSimulationMode

    Application.EnableEvents = oldEvents
    BeginAuthorizedWBSWrite "ScenarioForkForcedUpdate", ScenarioForkCalculatedWBSColumns()
    forcedUpdateGuardStarted = True
    Run_Forced_Planning_Update

    ThisWorkbook.Save
    EndAuthorizedWBSWrite
    forcedUpdateGuardStarted = False

CleanExit:
    On Error Resume Next
    If forcedUpdateGuardStarted Then EndAuthorizedWBSWrite
    If inputGuardStarted Then EndAuthorizedWBSWrite
    On Error GoTo 0
    Application.DisplayAlerts = oldAlerts
    Application.ScreenUpdating = oldScreenUpdating
    Application.EnableEvents = oldEvents
    Exit Sub

Fail:
    MsgBox "Erreur pendant l'initialisation du nouveau planning scénario :" & vbCrLf & Err.Description, vbCritical, "Planning scénario"
    Resume CleanExit

End Sub

Private Function ScenarioForkCalculatedWBSColumns() As Variant

    ScenarioForkCalculatedWBSColumns = Array( _
        "Baseline Finish", _
        "Actual Duration", _
        "Calculated Start", _
        "Calculated Finish", _
        "Calculated Duration", _
        "Start Variance", _
        "Finish Variance", _
        "Duration Variance", _
        "Driving Logic", _
        "Critical Path", _
        "Longest Path", _
        "Critical Path REX", _
        "Total Float", _
        "Free Float", _
        "Total Float REX", _
        "Free Float REX", _
        "Deadline Float")

End Function

Private Sub ClearScenarioForkRuntimeTables_CurrentWorkbook()

    ClearScenarioForkTableRows "CALC_ALARM", "tbl_CALC_ALARM"
    ClearScenarioForkTableRows "EVENT_HISTORY", "tbl_EVENT_HISTORY"
    ClearScenarioForkTableRows "EVENT_ACK", "tbl_EVENT_ACK"

End Sub

Private Sub ClearScenarioForkTableRows(ByVal sheetName As String, ByVal tableName As String)

    Dim ws As Worksheet
    Dim tbl As ListObject

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    Set tbl = ws.ListObjects(tableName)
    On Error GoTo 0

    If tbl Is Nothing Then Exit Sub
    If tbl.DataBodyRange Is Nothing Then Exit Sub

    tbl.DataBodyRange.Delete

End Sub

Private Sub ApplyScenarioBaselineToWBS_CurrentWorkbook()

    Dim wsWBS As Worksheet
    Dim wsTest As Worksheet
    Dim tblWBS As ListObject
    Dim tblTest As ListObject
    Dim mapWBS As Object
    Dim mapTest As Object
    Dim testRowById As Object
    Dim arrWBS As Variant
    Dim arrTest As Variant
    Dim r As Long
    Dim testRow As Long
    Dim idVal As String
    Dim scenarioStart As Variant
    Dim scenarioDuration As Variant
    Dim scenarioProgress As Variant
    Dim isScenarioSummary As Boolean
    Dim isScenarioLOE As Boolean
    Dim writeScenarioInputs As Boolean
    Dim hasValidScenarioBaseline As Boolean

    Set wsWBS = ThisWorkbook.Worksheets(WBS_SHEET)
    Set wsTest = ThisWorkbook.Worksheets(CALC_GANTT_TEST_SHEET)
    Set tblWBS = wsWBS.ListObjects(WBS_TABLE)
    Set tblTest = wsTest.ListObjects(CALC_GANTT_TEST_TABLE)

    If tblWBS.DataBodyRange Is Nothing Then Exit Sub
    If tblTest.DataBodyRange Is Nothing Then Exit Sub

    Set mapWBS = BuildColumnMap_GanttLive(tblWBS)
    Set mapTest = BuildColumnMap_GanttLive(tblTest)
    Set testRowById = CreateObject("Scripting.Dictionary")

    RequireScenarioForkColumns mapWBS, mapTest

    arrWBS = tblWBS.DataBodyRange.value
    arrTest = tblTest.DataBodyRange.value

    For r = 1 To UBound(arrTest, 1)
        idVal = Trim$(CStr(arrTest(r, mapTest("ID"))))
        If idVal <> "" Then
            If Not testRowById.Exists(idVal) Then testRowById(idVal) = r
        End If
    Next r

    For r = 1 To UBound(arrWBS, 1)
        idVal = Trim$(CStr(arrWBS(r, mapWBS("ID"))))
        If idVal <> "" And testRowById.Exists(idVal) Then
            testRow = CLng(testRowById(idVal))

            scenarioStart = GetCellValue(arrTest(testRow, mapTest("Calc Test Start")))
            scenarioDuration = GetCellValue(arrTest(testRow, mapTest("Calc Test Duration")))
            scenarioProgress = GetCellValue(arrTest(testRow, mapTest("Input Progress")))
            isScenarioSummary = IsScenarioForkSummaryRow(arrTest, mapTest, testRow)
            isScenarioLOE = IsScenarioForkLevelOfEffortRow(arrTest, mapTest, testRow)
            writeScenarioInputs = (Not isScenarioSummary) And (Not isScenarioLOE)
            hasValidScenarioBaseline = HasValue(scenarioStart) And IsScenarioForkPositiveDuration(scenarioDuration)

            If writeScenarioInputs Then
                If hasValidScenarioBaseline Then
                    tblWBS.DataBodyRange.Cells(r, mapWBS("Baseline Start")).value = scenarioStart
                    tblWBS.DataBodyRange.Cells(r, mapWBS("Baseline Duration")).value = scenarioDuration
                End If

                If HasValue(scenarioProgress) Then
                    tblWBS.DataBodyRange.Cells(r, mapWBS("% Progress")).value = scenarioProgress
                End If
            End If

            tblWBS.DataBodyRange.Cells(r, mapWBS("Forecast Start")).ClearContents
            tblWBS.DataBodyRange.Cells(r, mapWBS("Forecast Finish")).ClearContents
            tblWBS.DataBodyRange.Cells(r, mapWBS("Actual Start")).ClearContents
            tblWBS.DataBodyRange.Cells(r, mapWBS("Actual Finish")).ClearContents
        End If
    Next r

    tblWBS.ListColumns("Baseline Start").DataBodyRange.NumberFormat = "dd/mm/yyyy"
    tblWBS.ListColumns("Baseline Duration").DataBodyRange.NumberFormat = "0"
    tblWBS.ListColumns("Forecast Start").DataBodyRange.NumberFormat = "dd/mm/yyyy"
    tblWBS.ListColumns("Forecast Finish").DataBodyRange.NumberFormat = "dd/mm/yyyy"
    tblWBS.ListColumns("Actual Start").DataBodyRange.NumberFormat = "dd/mm/yyyy"
    tblWBS.ListColumns("Actual Finish").DataBodyRange.NumberFormat = "dd/mm/yyyy"
    tblWBS.ListColumns("% Progress").DataBodyRange.NumberFormat = "0%"

End Sub

Private Sub RequireScenarioForkColumns(ByVal mapWBS As Object, ByVal mapTest As Object)

    Dim c As Variant

    For Each c In Array("ID", "Baseline Start", "Baseline Duration", "Forecast Start", "Forecast Finish", "Actual Start", "Actual Finish", "% Progress")
        If Not mapWBS.Exists(CStr(c)) Then Err.Raise vbObjectError + 1290, , "Missing WBS column: " & CStr(c)
    Next c

    For Each c In Array("ID", "Task Type", "Is Summary", "Calc Test Start", "Calc Test Duration", "Input Progress")
        If Not mapTest.Exists(CStr(c)) Then Err.Raise vbObjectError + 1291, , "Missing scenario column: " & CStr(c)
    Next c

End Sub

Private Function IsScenarioForkPositiveDuration(ByVal value As Variant) As Boolean

    If Not HasValue(value) Then Exit Function
    If Not IsNumeric(value) Then Exit Function

    IsScenarioForkPositiveDuration = (CDbl(value) > 0)

End Function

Private Function IsScenarioForkSummaryRow( _
    ByRef dataArr As Variant, _
    ByVal mapCol As Object, _
    ByVal rowIndex As Long) As Boolean

    Dim rawValue As String

    If mapCol Is Nothing Then Exit Function
    If Not mapCol.Exists("Is Summary") Then Exit Function

    rawValue = UCase$(Trim$(CStr(dataArr(rowIndex, mapCol("Is Summary")))))
    IsScenarioForkSummaryRow = (rawValue = "YES" Or rawValue = "TRUE" Or rawValue = "1" Or rawValue = "SUMMARY")

End Function

Private Function IsScenarioForkLevelOfEffortRow( _
    ByRef dataArr As Variant, _
    ByVal mapCol As Object, _
    ByVal rowIndex As Long) As Boolean

    Dim rawValue As String

    If mapCol Is Nothing Then Exit Function
    If Not mapCol.Exists("Task Type") Then Exit Function

    rawValue = UCase$(Trim$(CStr(dataArr(rowIndex, mapCol("Task Type")))))
    IsScenarioForkLevelOfEffortRow = _
        (rawValue = "LEVEL OF EFFORT") Or _
        (rawValue = "LOE") Or _
        (rawValue = "LEVEL-OF-EFFORT") Or _
        (rawValue = "LEVEL_OF_EFFORT")

End Function

Private Sub DeactivateScenarioConstraints_CurrentWorkbook()

    Dim wsConstraints As Worksheet
    Dim tblConstraints As ListObject
    Dim activeCol As ListColumn

    On Error Resume Next
    Set wsConstraints = ThisWorkbook.Worksheets("CONSTRAINTS")
    Set tblConstraints = wsConstraints.ListObjects("tbl_CONSTRAINTS")
    On Error GoTo 0

    If tblConstraints Is Nothing Then Exit Sub
    If tblConstraints.DataBodyRange Is Nothing Then Exit Sub

    Set activeCol = GetScenarioConstraintActiveColumn(tblConstraints)
    If activeCol Is Nothing Then Exit Sub

    activeCol.DataBodyRange.value = "No"

End Sub

Private Function GetScenarioConstraintActiveColumn(ByVal tbl As ListObject) As ListColumn

    Dim col As ListColumn
    Dim normalizedName As String

    For Each col In tbl.ListColumns
        normalizedName = UCase$(Trim$(CStr(col.Name)))
        Select Case normalizedName
            Case "CONSTRAINT ACTIVE", "ACTIVE", "IS ACTIVE"
                Set GetScenarioConstraintActiveColumn = col
                Exit Function
        End Select
    Next col

End Function

Private Function TableHasColumn_GanttLive(ByVal tbl As ListObject, ByVal colName As String) As Boolean

    Dim i As Long

    For i = 1 To tbl.ListColumns.Count
        If StrComp(CStr(tbl.ListColumns(i).Name), colName, vbTextCompare) = 0 Then
            TableHasColumn_GanttLive = True
            Exit Function
        End If
    Next i

End Function

Private Sub ValidateLockSourceColumns(ByVal mapWBS As Object, ByVal mapCalc As Object)

    Dim reqWBS As Variant
    Dim reqCalc As Variant
    Dim c As Variant

    reqWBS = Array( _
        "ID", _
        "Forecast Start", _
        "Forecast Finish", _
        "% Progress", _
        "Calculated Start", _
        "Calculated Finish" _
    )

    reqCalc = Array( _
        "ID", _
        "Error flag" _
    )

    For Each c In reqWBS
        If Not mapWBS.Exists(CStr(c)) Then
            Err.Raise vbObjectError + 980, , "Missing column in tbl_WBS: " & CStr(c)
        End If
    Next c

    For Each c In reqCalc
        If Not mapCalc.Exists(CStr(c)) Then
            Err.Raise vbObjectError + 981, , "Missing column in tbl_CALC: " & CStr(c)
        End If
    Next c

End Sub

Private Function BuildModifiedTestChangesMap( _
    ByVal wsGantt As Worksheet, _
    ByVal tblWBS As ListObject, _
    ByVal mapWBS As Object) As Object

    Dim d As Object
    Dim baseById As Object
    Dim wbsToId As Object
    Dim calcDrivingMap As Object

    Dim lastRow As Long
    Dim r As Long

    Dim wbsVal As String
    Dim idVal As String

    Dim testStart As Variant
    Dim testFinish As Variant
    Dim testProgressRaw As Variant
    Dim testProgressNorm As Variant

    Dim hasTestStart As Boolean
    Dim hasTestFinish As Boolean
    Dim hasTestProgress As Boolean

    Dim baseStart As Variant
    Dim baseFinish As Variant
    Dim baseProgress As Variant

    Set d = CreateObject("Scripting.Dictionary")
    Set baseById = GanttLive_BuildBaseByIdMap()
    Set wbsToId = BuildWbsToIdMapFromWBS_Live()
    Set calcDrivingMap = BuildCalcDrivingLogicMap_Live()

    lastRow = GetLastGanttRow_Live(wsGantt)
    If lastRow < GANTT_FIRST_TASK_ROW Then
        Set BuildModifiedTestChangesMap = d
        Exit Function
    End If

    For r = GANTT_FIRST_TASK_ROW To lastRow

        wbsVal = NormalizeWBS(CStr(wsGantt.Cells(r, GANTT_COL_WBS).value))
        If wbsVal = "" Then GoTo NextRow

        If Not wbsToId.Exists(wbsVal) Then GoTo NextRow
        idVal = CStr(wbsToId(wbsVal))

        If Not calcDrivingMap.Exists(idVal) Then GoTo NextRow
        If UCase$(Trim$(CStr(calcDrivingMap(idVal)))) = "SUMMARY" Then GoTo NextRow

        testStart = GetCellValue(wsGantt.Cells(r, COL_TEST_START).value)
        testFinish = GetCellValue(wsGantt.Cells(r, COL_TEST_FINISH).value)
        testProgressRaw = GetCellValue(wsGantt.Cells(r, COL_TEST_PROGRESS).value)

        hasTestStart = HasValue(testStart)
        hasTestFinish = HasValue(testFinish)
        hasTestProgress = HasValue(testProgressRaw)

        If hasTestProgress Then
            testProgressNorm = NormalizePercentInput(testProgressRaw)
        Else
            testProgressNorm = Empty
        End If

        If Not baseById.Exists(idVal) Then GoTo NextRow

        baseStart = baseById(idVal)(0)
        baseFinish = baseById(idVal)(1)
        baseProgress = baseById(idVal)(3)

        If hasTestStart Or hasTestFinish Or hasTestProgress Then
            If HasMeaningfulLockDelta(baseStart, baseFinish, baseProgress, testStart, testFinish, testProgressNorm) Then
                d(idVal) = Array(testStart, testFinish, testProgressNorm, hasTestStart, hasTestFinish, hasTestProgress)
            End If
        End If

NextRow:
    Next r

    Set BuildModifiedTestChangesMap = d

End Function

Private Function HasMeaningfulLockDelta( _
    ByVal baseStart As Variant, _
    ByVal baseFinish As Variant, _
    ByVal baseProgress As Variant, _
    ByVal testStart As Variant, _
    ByVal testFinish As Variant, _
    ByVal testProgress As Variant) As Boolean

    If HasValue(testStart) Then
        If ValuesDiffer(baseStart, testStart) Then
            HasMeaningfulLockDelta = True
            Exit Function
        End If
    End If

    If HasValue(testFinish) Then
        If ValuesDiffer(baseFinish, testFinish) Then
            HasMeaningfulLockDelta = True
            Exit Function
        End If
    End If

    If HasValue(testProgress) Then
        If ValuesDiffer(baseProgress, testProgress) Then
            HasMeaningfulLockDelta = True
            Exit Function
        End If
    End If

End Function

Private Function BackupWBSLockColumns(ByVal tblWBS As ListObject, ByVal mapWBS As Object) As Object

    Dim d As Object
    Dim arr As Variant
    Dim r As Long
    Dim idVal As String

    Set d = CreateObject("Scripting.Dictionary")

    If tblWBS.DataBodyRange Is Nothing Then
        Set BackupWBSLockColumns = d
        Exit Function
    End If

    arr = tblWBS.DataBodyRange.value

    For r = 1 To UBound(arr, 1)
        idVal = Trim$(CStr(arr(r, mapWBS("ID"))))
        If idVal <> "" Then
            d(idVal) = Array( _
                GetCellValue(arr(r, mapWBS("Forecast Start"))), _
                GetCellValue(arr(r, mapWBS("Forecast Finish"))), _
                GetCellValue(arr(r, mapWBS("% Progress"))), _
                NormalizeCalendarType(arr(r, mapWBS("Cal"))) _
            )
        End If
    Next r

    Set BackupWBSLockColumns = d

End Function

Private Sub RestoreWBSLockColumns(ByVal tblWBS As ListObject, ByVal mapWBS As Object, ByVal backup As Object)

    Dim r As Long
    Dim idVal As String
    Dim vals As Variant

    If tblWBS.DataBodyRange Is Nothing Then Exit Sub
    If backup Is Nothing Then Exit Sub

    For r = 1 To tblWBS.ListRows.Count
        idVal = Trim$(CStr(tblWBS.DataBodyRange.Cells(r, mapWBS("ID")).value))

        If idVal <> "" Then
            If backup.Exists(idVal) Then
                vals = backup(idVal)

                tblWBS.DataBodyRange.Cells(r, mapWBS("Forecast Start")).value = vals(0)
                tblWBS.DataBodyRange.Cells(r, mapWBS("Forecast Finish")).value = vals(1)
                tblWBS.DataBodyRange.Cells(r, mapWBS("% Progress")).value = vals(2)
            End If
        End If
    Next r

End Sub

Private Sub ApplyModifiedTestChangesToWBS( _
    ByVal tblWBS As ListObject, _
    ByVal mapWBS As Object, _
    ByVal changes As Object)

    Dim r As Long
    Dim idVal As String
    Dim vals As Variant

    Dim hasTestStart As Boolean
    Dim hasTestFinish As Boolean
    Dim hasTestProgress As Boolean

    If tblWBS.DataBodyRange Is Nothing Then Exit Sub
    If changes Is Nothing Then Exit Sub
    If changes.Count = 0 Then Exit Sub

    For r = 1 To tblWBS.ListRows.Count

        idVal = Trim$(CStr(tblWBS.DataBodyRange.Cells(r, mapWBS("ID")).value))

        If idVal <> "" Then
            If changes.Exists(idVal) Then

                vals = changes(idVal)

                hasTestStart = CBool(vals(3))
                hasTestFinish = CBool(vals(4))
                hasTestProgress = CBool(vals(5))

                'Important:
                'A blank TEST Start means "do not lock/change Forecast Start",
                'not "clear Forecast Start".
                If hasTestStart Then
                    tblWBS.DataBodyRange.Cells(r, mapWBS("Forecast Start")).value = vals(0)
                End If

                'Important:
                'A blank TEST Finish means "do not lock/change Forecast Finish",
                'not "clear Forecast Finish".
                If hasTestFinish Then
                    tblWBS.DataBodyRange.Cells(r, mapWBS("Forecast Finish")).value = vals(1)
                End If

                If hasTestProgress Then
                    tblWBS.DataBodyRange.Cells(r, mapWBS("% Progress")).value = vals(2)
                End If

            End If
        End If

    Next r

End Sub

Private Function CalcTableHasErrors(ByVal tblCalc As ListObject, ByVal mapCalc As Object) As Boolean

    Dim arr As Variant
    Dim r As Long

    If tblCalc.DataBodyRange Is Nothing Then Exit Function

    arr = tblCalc.DataBodyRange.value

    For r = 1 To UBound(arr, 1)
        If UCase$(Trim$(CStr(arr(r, mapCalc("Error flag"))))) = "ERROR" Then
            CalcTableHasErrors = True
            Exit Function
        End If
    Next r

End Function

Private Function BackupGanttTestInputs(ByVal wsGantt As Worksheet) As Object

    Dim d As Object
    Dim lastRow As Long
    Dim r As Long
    Dim wbsVal As String

    Dim testStart As Variant
    Dim testFinish As Variant
    Dim testProgress As Variant

    Set d = CreateObject("Scripting.Dictionary")

    lastRow = GetLastGanttRow_Live(wsGantt)
    If lastRow < GANTT_FIRST_TASK_ROW Then
        Set BackupGanttTestInputs = d
        Exit Function
    End If

    For r = GANTT_FIRST_TASK_ROW To lastRow

        wbsVal = NormalizeWBS(CStr(wsGantt.Cells(r, GANTT_COL_WBS).value))
        If wbsVal = "" Then GoTo NextRow

        testStart = GetCellValue(wsGantt.Cells(r, COL_TEST_START).value)
        testFinish = GetCellValue(wsGantt.Cells(r, COL_TEST_FINISH).value)
        testProgress = GetCellValue(wsGantt.Cells(r, COL_TEST_PROGRESS).value)

        If HasValue(testStart) Or HasValue(testFinish) Or HasValue(testProgress) Then
            d(wbsVal) = Array(testStart, testFinish, testProgress)
        End If

NextRow:
    Next r

    Set BackupGanttTestInputs = d

End Function

Private Sub RestoreGanttTestInputsFromBackup(ByVal wsGantt As Worksheet, ByVal backup As Object)

    Dim wbsToRow As Object
    Dim lastRow As Long
    Dim r As Long

    Dim wbsVal As String
    Dim vals As Variant
    Dim rowIndex As Long
    Dim key As Variant

    If backup Is Nothing Then Exit Sub
    If backup.Count = 0 Then Exit Sub

    Set wbsToRow = CreateObject("Scripting.Dictionary")

    lastRow = GetLastGanttRow_Live(wsGantt)
    If lastRow < GANTT_FIRST_TASK_ROW Then Exit Sub

    For r = GANTT_FIRST_TASK_ROW To lastRow
        wbsVal = NormalizeWBS(CStr(wsGantt.Cells(r, GANTT_COL_WBS).value))
        If wbsVal <> "" Then
            wbsToRow(wbsVal) = r
        End If
    Next r

    For Each key In backup.Keys

        wbsVal = CStr(key)

        If wbsToRow.Exists(wbsVal) Then

            rowIndex = wbsToRow(wbsVal)
            vals = backup(wbsVal)

            wsGantt.Cells(rowIndex, COL_TEST_START).value = vals(0)
            wsGantt.Cells(rowIndex, COL_TEST_FINISH).value = vals(1)
            wsGantt.Cells(rowIndex, COL_TEST_PROGRESS).value = vals(2)

            wsGantt.Cells(rowIndex, COL_TEST_START).NumberFormat = "dd/mm/yyyy"
            wsGantt.Cells(rowIndex, COL_TEST_FINISH).NumberFormat = "dd/mm/yyyy"
            wsGantt.Cells(rowIndex, COL_TEST_PROGRESS).NumberFormat = "0%"

        End If

    Next key

End Sub

Private Sub ClearGanttTestInputs_ForChangedTasks( _
    ByVal wsGantt As Worksheet, _
    ByVal appliedChanges As Object)

    Dim wbsToRow As Object
    Dim wbsFromId As Object
    Dim lastRow As Long
    Dim r As Long
    Dim rowIndex As Long
    Dim wbsVal As String
    Dim idVal As Variant

    Set wbsToRow = CreateObject("Scripting.Dictionary")

    lastRow = GetLastGanttRow_Live(wsGantt)
    If lastRow < GANTT_FIRST_TASK_ROW Then Exit Sub

    For r = GANTT_FIRST_TASK_ROW To lastRow
        wbsVal = NormalizeWBS(CStr(wsGantt.Cells(r, GANTT_COL_WBS).value))
        If wbsVal <> "" Then
            wbsToRow(wbsVal) = r
        End If
    Next r

    Set wbsFromId = BuildIdToWbsMapFromWBS_Live()

    For Each idVal In appliedChanges.Keys

        If wbsFromId.Exists(CStr(idVal)) Then

            wbsVal = CStr(wbsFromId(CStr(idVal)))

            If wbsToRow.Exists(wbsVal) Then
                rowIndex = CLng(wbsToRow(wbsVal))

                wsGantt.Cells(rowIndex, COL_TEST_START).ClearContents
                wsGantt.Cells(rowIndex, COL_TEST_FINISH).ClearContents
                wsGantt.Cells(rowIndex, COL_TEST_PROGRESS).ClearContents

                wsGantt.Cells(rowIndex, COL_TEST_START).Interior.Color = RGB(255, 255, 153)
                wsGantt.Cells(rowIndex, COL_TEST_FINISH).Interior.Color = RGB(255, 255, 153)
                wsGantt.Cells(rowIndex, COL_TEST_PROGRESS).Interior.Color = RGB(255, 255, 153)
            End If

        End If

    Next idVal

End Sub


Private Sub ClearCalcGanttTestResults()

    Dim ws As Worksheet
    Dim tbl As ListObject

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(CALC_GANTT_TEST_SHEET)
    Set tbl = ws.ListObjects(CALC_GANTT_TEST_TABLE)
    On Error GoTo 0

    If tbl Is Nothing Then Exit Sub
    If tbl.DataBodyRange Is Nothing Then Exit Sub

    tbl.DataBodyRange.ClearContents

End Sub

Private Function BuildIdToWbsMapFromWBS_Live() As Object

    Dim d As Object
    Dim wsWBS As Worksheet
    Dim tblWBS As ListObject
    Dim mapWBS As Object
    Dim arr As Variant
    Dim r As Long
    Dim idVal As String
    Dim wbsVal As String

    Set d = CreateObject("Scripting.Dictionary")
    Set wsWBS = ThisWorkbook.Worksheets(WBS_SHEET)
    Set tblWBS = wsWBS.ListObjects(WBS_TABLE)

    If tblWBS.DataBodyRange Is Nothing Then
        Set BuildIdToWbsMapFromWBS_Live = d
        Exit Function
    End If

    Set mapWBS = BuildColumnMap_GanttLive(tblWBS)
    arr = tblWBS.DataBodyRange.value

    For r = 1 To UBound(arr, 1)
        idVal = Trim$(CStr(arr(r, mapWBS("ID"))))
        wbsVal = NormalizeWBS(CStr(arr(r, mapWBS("WBS"))))

        If idVal <> "" And wbsVal <> "" Then
            d(idVal) = wbsVal
        End If
    Next r

    Set BuildIdToWbsMapFromWBS_Live = d

End Function

Private Function BuildLivePredsFromLogicLinks_FS_SS_FF_WithLagAndType( _
    ByVal validIds As Object, _
    ByVal predsById As Object, _
    ByVal childrenById As Object, _
    ByVal directChildrenById As Object, _
    ByVal predLagBySuccPred As Object, _
    ByVal predTypeBySuccPred As Object, _
    ByVal structuralErrors As Object, _
    ByVal errMissingPred As Object, _
    ByVal errUnsupportedLinkType As Object, _
    ByRef hasStructureError As Boolean) As Boolean

    Dim wsCalc As Worksheet
    Dim tblLinks As ListObject
    Dim mapLinks As Object
    Dim arrLinks As Variant

    Dim r As Long
    Dim i As Long

    Dim succId As String
    Dim predId As String
    Dim linkType As String
    Dim linkLag As Variant
    Dim linkKey As String

    Dim expandedLeafPreds As Collection
    Dim leafPredId As Variant

    BuildLivePredsFromLogicLinks_FS_SS_FF_WithLagAndType = False

    Set mapLinks = CreateObject("Scripting.Dictionary")

    On Error GoTo SafeExit

    Set wsCalc = ThisWorkbook.Worksheets(CALC_SHEET)
    Set tblLinks = wsCalc.ListObjects("tbl_LOGIC_LINKS")

    If tblLinks Is Nothing Then
        hasStructureError = True
        Exit Function
    End If

    For i = 1 To tblLinks.ListColumns.Count
        mapLinks(tblLinks.ListColumns(i).Name) = i
    Next i

    If Not mapLinks.Exists("Succ ID") Then
        hasStructureError = True
        Exit Function
    End If

    If Not mapLinks.Exists("Pred ID") Then
        hasStructureError = True
        Exit Function
    End If

    If Not mapLinks.Exists("Link Type") Then
        hasStructureError = True
        Exit Function
    End If

    If Not mapLinks.Exists("Lag") Then
        hasStructureError = True
        Exit Function
    End If

    If tblLinks.DataBodyRange Is Nothing Then
        BuildLivePredsFromLogicLinks_FS_SS_FF_WithLagAndType = True
        Exit Function
    End If

    arrLinks = tblLinks.DataBodyRange.value

    For r = 1 To UBound(arrLinks, 1)

        succId = Trim$(CStr(arrLinks(r, mapLinks("Succ ID"))))
        predId = Trim$(CStr(arrLinks(r, mapLinks("Pred ID"))))
        linkType = UCase$(Trim$(CStr(arrLinks(r, mapLinks("Link Type")))))
        linkLag = arrLinks(r, mapLinks("Lag"))

        If succId = "" Then GoTo NextLinkRow
        If Not validIds.Exists(succId) Then GoTo NextLinkRow

        If predId = "" Then
            structuralErrors(succId) = True
            errMissingPred(succId) = True
            hasStructureError = True
            GoTo NextLinkRow
        End If

        If linkType = "" Then linkType = "FS"

        If linkType <> "FS" And linkType <> "SS" And linkType <> "FF" Then
            structuralErrors(succId) = True
            errUnsupportedLinkType(succId) = True
            hasStructureError = True
            GoTo NextLinkRow
        End If

        If Not predsById.Exists(predId) Then
            structuralErrors(succId) = True
            errMissingPred(succId) = True
            hasStructureError = True
            GoTo NextLinkRow
        End If

        Set expandedLeafPreds = GetLiveLeafDescendantsFromHierarchy(predId, directChildrenById, validIds)

        If expandedLeafPreds Is Nothing Then
            structuralErrors(succId) = True
            errMissingPred(succId) = True
            hasStructureError = True
            GoTo NextLinkRow
        End If

        If expandedLeafPreds.Count = 0 Then
            structuralErrors(succId) = True
            errMissingPred(succId) = True
            hasStructureError = True
            GoTo NextLinkRow
        End If

        For Each leafPredId In expandedLeafPreds

            predsById(succId).Add CStr(leafPredId)
            childrenById(CStr(leafPredId)).Add succId

            linkKey = succId & "|" & CStr(leafPredId)

            If IsNumeric(linkLag) Then
                predLagBySuccPred(linkKey) = CDbl(linkLag)
            Else
                predLagBySuccPred(linkKey) = 0#
            End If

            predTypeBySuccPred(linkKey) = linkType

        Next leafPredId

NextLinkRow:
    Next r

    BuildLivePredsFromLogicLinks_FS_SS_FF_WithLagAndType = True
    Exit Function

SafeExit:
    BuildLivePredsFromLogicLinks_FS_SS_FF_WithLagAndType = False

End Function

Private Function GetLiveLeafDescendantsFromHierarchy( _
    ByVal startId As String, _
    ByVal directChildrenById As Object, _
    ByVal validIds As Object) As Collection

    Dim result As Collection
    Dim childId As Variant
    Dim childLeafs As Collection
    Dim leafId As Variant

    Set result = New Collection

    If validIds.Exists(startId) Then
        result.Add startId
        Set GetLiveLeafDescendantsFromHierarchy = result
        Exit Function
    End If

    If Not directChildrenById.Exists(startId) Then
        Set GetLiveLeafDescendantsFromHierarchy = result
        Exit Function
    End If

    If directChildrenById(startId).Count = 0 Then
        Set GetLiveLeafDescendantsFromHierarchy = result
        Exit Function
    End If

    For Each childId In directChildrenById(startId)

        Set childLeafs = GetLiveLeafDescendantsFromHierarchy(CStr(childId), directChildrenById, validIds)

        If Not childLeafs Is Nothing Then
            For Each leafId In childLeafs
                result.Add CStr(leafId)
            Next leafId
        End If

    Next childId

    Set GetLiveLeafDescendantsFromHierarchy = result

End Function

'=====================================================
' HELPERS – TEST RENDER FLOW
'=====================================================

Private Sub GanttLive_AbortTestEngine(ByVal wsGantt As Worksheet)

    SetGanttPreserveTestInputs False

    If Not wsGantt Is Nothing Then
        wsGantt.Activate
    End If

End Sub

Private Sub GanttLive_ApplyTestRender(ByVal wsGantt As Worksheet)

    Dim fastRenderApplied As Boolean

    SetGanttPreserveTestInputs True
    GanttLive_RequestTestRender
    fastRenderApplied = Gantt_TryApplyTestDayPredictiveRegistry()
    If Not fastRenderApplied Then Refresh_Gantt
    GanttLive_SetActiveSimulationMode "TEST"
    SetGanttPreserveTestInputs False

    If Not wsGantt Is Nothing Then
        wsGantt.Activate
    End If

End Sub

Private Sub GanttLive_SafeExit_TestEngine()

    SetGanttPreserveTestInputs False
    GanttLive_ClearTestRenderRequest

End Sub


'=====================================================
' HELPERS – LOCK FLOW
'=====================================================

Private Sub GanttLive_RollbackFailedLock( _
    ByVal tblWBS As ListObject, _
    ByVal mapWBS As Object, _
    ByVal wsGantt As Worksheet, _
    ByVal wbsBackup As Object, _
    ByVal ganttTestBackup As Object)

    RestoreWBSLockColumns tblWBS, mapWBS, wbsBackup

    GanttLive_ClearTestRenderRequest
    Run_Calc_Engine

    RestoreGanttTestInputsFromBackup wsGantt, ganttTestBackup

End Sub

Private Sub GanttLive_FinalizeSuccessfulLock( _
    ByVal wsGantt As Worksheet, _
    ByVal appliedChanges As Object)

    ClearGanttTestInputs_ForChangedTasks wsGantt, appliedChanges
    ClearCalcGanttTestResults
    GanttLive_ClearTestRenderRequest
    GanttLive_ClearActiveSimulationMode
    SetGanttPreserveTestInputs False
    Refresh_Gantt

End Sub

Private Function BuildSimulatedLockResultMap(ByVal appliedChanges As Object) As Object

    Dim d As Object
    Dim testById As Object
    Dim key As Variant
    Dim simVals As Variant
    Dim changeVals As Variant

    Dim compareStart As Boolean
    Dim compareFinish As Boolean
    Dim compareProgress As Boolean

    Set d = CreateObject("Scripting.Dictionary")
    Set testById = GanttLive_BuildTestByIdMap()

    If appliedChanges Is Nothing Then
        Set BuildSimulatedLockResultMap = d
        Exit Function
    End If

    For Each key In appliedChanges.Keys

        If testById.Exists(CStr(key)) Then

            simVals = testById(CStr(key))
            changeVals = appliedChanges(CStr(key))

            compareStart = CBool(changeVals(3))
            compareFinish = CBool(changeVals(4))
            compareProgress = CBool(changeVals(5))

            d(CStr(key)) = Array( _
                simVals(0), _
                simVals(1), _
                simVals(3), _
                compareStart, _
                compareFinish, _
                compareProgress)

        End If

    Next key

    Set BuildSimulatedLockResultMap = d

End Function

Private Function LockResultsMatchSimulatedResult( _
    ByVal tblWBS As ListObject, _
    ByVal mapWBS As Object, _
    ByVal simulatedById As Object) As Boolean

    Dim r As Long
    Dim idVal As String
    Dim simVals As Variant

    Dim calcStartVal As Variant
    Dim calcFinishVal As Variant
    Dim progressVal As Variant

    Dim compareStart As Boolean
    Dim compareFinish As Boolean
    Dim compareProgress As Boolean

    If tblWBS.DataBodyRange Is Nothing Then
        LockResultsMatchSimulatedResult = False
        Exit Function
    End If

    For r = 1 To tblWBS.ListRows.Count

        idVal = Trim$(CStr(tblWBS.DataBodyRange.Cells(r, mapWBS("ID")).value))

        If idVal <> "" Then
            If simulatedById.Exists(idVal) Then

                simVals = simulatedById(idVal)

                compareStart = CBool(simVals(3))
                compareFinish = CBool(simVals(4))
                compareProgress = CBool(simVals(5))

                If compareStart Then
                    calcStartVal = GetCellValue(tblWBS.DataBodyRange.Cells(r, mapWBS("Calculated Start")).value)
                    If ValuesDiffer(calcStartVal, simVals(0)) Then Exit Function
                End If

                If compareFinish Then
                    calcFinishVal = GetCellValue(tblWBS.DataBodyRange.Cells(r, mapWBS("Calculated Finish")).value)
                    If ValuesDiffer(calcFinishVal, simVals(1)) Then Exit Function
                End If

                If compareProgress Then
                    progressVal = GetCellValue(tblWBS.DataBodyRange.Cells(r, mapWBS("% Progress")).value)
                    If ValuesDiffer(progressVal, simVals(2)) Then Exit Function
                End If

            End If
        End If

    Next r

    LockResultsMatchSimulatedResult = True

End Function

Private Sub BuildTestInputValues_FromWBSSource( _
    ByVal baselineStart As Variant, _
    ByVal baselineDuration As Variant, _
    ByVal actualStart As Variant, _
    ByVal actualFinish As Variant, _
    ByVal forecastStart As Variant, _
    ByVal forecastFinish As Variant, _
    ByVal baseProgress As Variant, _
    ByVal testStart As Variant, _
    ByVal testFinish As Variant, _
    ByVal testProgress As Variant, _
    ByRef inputStart As Variant, _
    ByRef inputFinish As Variant, _
    ByRef inputDuration As Variant, _
    ByRef inputProgress As Variant)

    inputStart = Empty
    inputFinish = Empty
    inputDuration = baselineDuration
    inputProgress = baseProgress

    '----------------------------------------
    ' START candidate
    ' Priorité stricte : Test > Actual > Forecast > Baseline
    '----------------------------------------
    If HasValue(testStart) Then
        inputStart = testStart
    ElseIf HasValue(actualStart) Then
        inputStart = actualStart
    ElseIf HasValue(forecastStart) Then
        inputStart = forecastStart
    ElseIf HasValue(baselineStart) Then
        inputStart = baselineStart
    End If

    '----------------------------------------
    ' FINISH candidate
    ' Priorité stricte : Test > Actual > Forecast
    '----------------------------------------
    If HasValue(testFinish) Then
        inputFinish = testFinish
    ElseIf HasValue(actualFinish) Then
        inputFinish = actualFinish
    ElseIf HasValue(forecastFinish) Then
        inputFinish = forecastFinish
    End If

    '----------------------------------------
    ' Cas explicite Test Start + Test Finish :
    ' on dérive la durée du test saisi
    '----------------------------------------
    If HasValue(testStart) And HasValue(testFinish) Then
        inputDuration = CDbl(testFinish) - CDbl(testStart) + 1
    End If

    '----------------------------------------
    ' Progress
    '----------------------------------------
    If HasValue(testProgress) Then
        inputProgress = testProgress
    End If

End Sub

Private Sub BuildGanttTestCoreDataset( _
    ByVal tblTest As ListObject, _
    ByRef dataCore As Variant, _
    ByRef mapCore As Object, _
    ByVal idToRowTest As Object, _
    ByVal idToWbs As Object, _
    ByRef outWarnFlag() As Variant, _
    ByRef outWarnText() As Variant, _
    ByVal warningActualIds As Object)

    Dim arrTest As Variant
    Dim rowCount As Long
    Dim r As Long

    Dim headers As Variant
    Dim c As Long

    Dim idVal As String
    Dim wbsVal As String
    Dim taskTypeVal As String
    Dim isSummary As String
    Dim parentId As String

    Dim rawById As Object
    Dim calcConstraintById As Object
    Dim calcConstraintValues As Variant

    Dim rawActualStart As Variant
    Dim rawActualFinish As Variant
    Dim rawForecastStart As Variant
    Dim rawForecastFinish As Variant
    Dim rawBaselineStart As Variant
    Dim rawBaselineDuration As Variant
    Dim rawCal As String

    Dim taskName As String
    Dim constraintActive As String
    Dim startConstraintType As String
    Dim startConstraintDate As Variant
    Dim finishConstraintType As String
    Dim finishConstraintDate As Variant

    Dim testStart As Variant
    Dim testFinish As Variant
    Dim testProgress As Variant

    Dim actualStart As Variant
    Dim actualFinish As Variant
    Dim forecastStart As Variant
    Dim forecastFinish As Variant
    Dim baselineStart As Variant
    Dim baselineDuration As Variant
    Dim inputProgress As Variant

    Dim hasActual As Boolean
    Dim isSummaryBool As Boolean
    Dim rowHasMeaningfulTest As Boolean

    headers = Array( _
        "ID", "WBS", "Task Name", "Task Type", "Cal", "ParentID", "IsSummary", _
        "Actual Start", "Actual Finish", "Forecast Start", "Forecast Finish", _
        "Baseline Start", "Baseline Duration", "Constraint Active", _
        "Start Constraint Type", "Start Constraint Date", _
        "Finish Constraint Type", "Finish Constraint Date", _
        "Calculated Start", "Calculated Finish", "Calculated Duration", _
        "Error flag", "ErrorMsg", "Input Progress" _
    )

    Set mapCore = CreateObject("Scripting.Dictionary")
    For c = LBound(headers) To UBound(headers)
        mapCore(CStr(headers(c))) = c + 1
    Next c

    arrTest = tblTest.DataBodyRange.value
    rowCount = UBound(arrTest, 1)

    ReDim dataCore(1 To rowCount, 1 To UBound(headers) + 1)

    Set rawById = BuildRawWbsSourceById_Live()
    Set calcConstraintById = BuildCalcConstraintByIdMap_GanttLive()

    For r = 1 To rowCount

        idVal = Trim$(CStr(arrTest(r, GetColumnIndex_GanttLive(tblTest, "ID"))))
        wbsVal = NormalizeWBS(CStr(arrTest(r, GetColumnIndex_GanttLive(tblTest, "WBS"))))
        taskTypeVal = Trim$(CStr(arrTest(r, GetColumnIndex_GanttLive(tblTest, "Task Type"))))
        isSummary = UCase$(Trim$(CStr(arrTest(r, GetColumnIndex_GanttLive(tblTest, "Is Summary")))))
        parentId = Trim$(CStr(arrTest(r, GetColumnIndex_GanttLive(tblTest, "Parent ID"))))

        idToRowTest(idVal) = r
        idToWbs(idVal) = wbsVal

        rawActualStart = Empty
        rawActualFinish = Empty
        rawForecastStart = Empty
        rawForecastFinish = Empty
        rawBaselineStart = Empty
        rawBaselineDuration = Empty
        rawCal = CALENDAR_7D

        taskName = vbNullString
        constraintActive = "No"
        startConstraintType = vbNullString
        startConstraintDate = Empty
        finishConstraintType = vbNullString
        finishConstraintDate = Empty

        If rawById.Exists(idVal) Then
            rawActualStart = rawById(idVal)(0)
            rawActualFinish = rawById(idVal)(1)
            rawForecastStart = rawById(idVal)(2)
            rawForecastFinish = rawById(idVal)(3)
            rawBaselineStart = rawById(idVal)(4)
            rawBaselineDuration = rawById(idVal)(5)
            rawCal = CStr(rawById(idVal)(7))
        End If

        If calcConstraintById.Exists(idVal) Then
            calcConstraintValues = calcConstraintById(idVal)
            taskName = Trim$(CStr(calcConstraintValues(0)))
            constraintActive = Trim$(CStr(calcConstraintValues(1)))
            startConstraintType = Trim$(CStr(calcConstraintValues(2)))
            startConstraintDate = calcConstraintValues(3)
            finishConstraintType = Trim$(CStr(calcConstraintValues(4)))
            finishConstraintDate = calcConstraintValues(5)
        End If

        testStart = GetCellValue(arrTest(r, GetColumnIndex_GanttLive(tblTest, "Test Start")))
        testFinish = GetCellValue(arrTest(r, GetColumnIndex_GanttLive(tblTest, "Test Finish")))
        testProgress = GetCellValue(arrTest(r, GetColumnIndex_GanttLive(tblTest, "Test % Normalized")))

        hasActual = (HasValue(rawActualStart) Or HasValue(rawActualFinish))
        isSummaryBool = (isSummary = "YES")
        rowHasMeaningfulTest = (Trim$(CStr(arrTest(r, GetColumnIndex_GanttLive(tblTest, "Any Test Value")))) = "YES")

        outWarnFlag(r, 1) = ""
        outWarnText(r, 1) = ""

        actualStart = rawActualStart
        actualFinish = rawActualFinish
        forecastStart = rawForecastStart
        forecastFinish = rawForecastFinish
        baselineStart = rawBaselineStart
        baselineDuration = rawBaselineDuration
        inputProgress = GetCellValue(arrTest(r, GetColumnIndex_GanttLive(tblTest, "Base Progress")))

        If HasValue(testProgress) Then
            inputProgress = testProgress
        End If

        If rowHasMeaningfulTest And Not isSummaryBool Then

            If HasValue(testStart) Then forecastStart = testStart
            If HasValue(testFinish) Then forecastFinish = testFinish

            If HasValue(testStart) And HasValue(testFinish) Then
                baselineStart = testStart
                baselineDuration = CDbl(testFinish) - CDbl(testStart) + 1
            End If

            If hasActual Then
                warningActualIds(idVal) = True
                outWarnFlag(r, 1) = "WARNING"
                outWarnText(r, 1) = "TEST INPUT ON ACTUAL TASK - SIMULATION MAY SHOW NO EFFECT AND LOCK MAY NOT MATCH PROD"
            End If

        End If

        dataCore(r, mapCore("ID")) = idVal
        dataCore(r, mapCore("WBS")) = wbsVal
        dataCore(r, mapCore("Task Name")) = taskName
        dataCore(r, mapCore("Task Type")) = taskTypeVal
        dataCore(r, mapCore("Cal")) = NormalizeCalendarType(rawCal)
        dataCore(r, mapCore("ParentID")) = parentId
        dataCore(r, mapCore("IsSummary")) = isSummaryBool
        dataCore(r, mapCore("Actual Start")) = actualStart
        dataCore(r, mapCore("Actual Finish")) = actualFinish
        dataCore(r, mapCore("Forecast Start")) = forecastStart
        dataCore(r, mapCore("Forecast Finish")) = forecastFinish
        dataCore(r, mapCore("Baseline Start")) = baselineStart
        dataCore(r, mapCore("Baseline Duration")) = baselineDuration
        dataCore(r, mapCore("Constraint Active")) = constraintActive
        dataCore(r, mapCore("Start Constraint Type")) = startConstraintType
        dataCore(r, mapCore("Start Constraint Date")) = startConstraintDate
        dataCore(r, mapCore("Finish Constraint Type")) = finishConstraintType
        dataCore(r, mapCore("Finish Constraint Date")) = finishConstraintDate
        dataCore(r, mapCore("Calculated Start")) = Empty
        dataCore(r, mapCore("Calculated Finish")) = Empty
        dataCore(r, mapCore("Calculated Duration")) = Empty
        dataCore(r, mapCore("Error flag")) = ""
        dataCore(r, mapCore("ErrorMsg")) = ""
        dataCore(r, mapCore("Input Progress")) = inputProgress

    Next r

End Sub

Private Function GanttLive_BuildMeaningfulChangedTestIds(ByVal tblTest As ListObject) As Object

    Dim d As Object
    Dim arr As Variant
    Dim mapTest As Object
    Dim r As Long
    Dim idVal As String

    Dim baseStart As Variant
    Dim baseFinish As Variant
    Dim baseProgress As Variant

    Dim testStart As Variant
    Dim testFinish As Variant
    Dim testProgress As Variant

    Set d = CreateObject("Scripting.Dictionary")

    If tblTest Is Nothing Then
        Set GanttLive_BuildMeaningfulChangedTestIds = d
        Exit Function
    End If

    If tblTest.DataBodyRange Is Nothing Then
        Set GanttLive_BuildMeaningfulChangedTestIds = d
        Exit Function
    End If

    Set mapTest = BuildColumnMap_GanttLive(tblTest)
    arr = tblTest.DataBodyRange.value

    For r = 1 To UBound(arr, 1)

        idVal = Trim$(CStr(arr(r, mapTest("ID"))))
        If idVal = "" Then GoTo NextRow

        If UCase$(Trim$(CStr(arr(r, mapTest("Is Summary"))))) = "YES" Then GoTo NextRow
        If Trim$(CStr(arr(r, mapTest("Any Test Value")))) <> "YES" Then GoTo NextRow

        baseStart = GetCellValue(arr(r, mapTest("Base Start")))
        baseFinish = GetCellValue(arr(r, mapTest("Base Finish")))
        baseProgress = GetCellValue(arr(r, mapTest("Base Progress")))

        testStart = GetCellValue(arr(r, mapTest("Test Start")))
        testFinish = GetCellValue(arr(r, mapTest("Test Finish")))
        testProgress = GetCellValue(arr(r, mapTest("Test % Normalized")))

        If HasMeaningfulLockDelta(baseStart, baseFinish, baseProgress, testStart, testFinish, testProgress) Then
            d(idVal) = True
        End If

NextRow:
    Next r

    Set GanttLive_BuildMeaningfulChangedTestIds = d

End Function

Private Function GanttLive_FindActualTasksImpactedByTestChanges( _
    ByVal tblTest As ListObject, _
    ByVal changedTestIds As Object, _
    ByVal linksBySuccId As Object) As Object

    Dim impactedActualIds As Object
    Dim childrenByPred As Object
    Dim queue As Collection
    Dim visited As Object

    Dim mapTest As Object
    Dim arr As Variant
    Dim hasActualById As Object

    Dim r As Long
    Dim idVal As String
    Dim startId As Variant
    Dim currentId As String
    Dim childId As Variant

    Set impactedActualIds = CreateObject("Scripting.Dictionary")

    If tblTest Is Nothing Then
        Set GanttLive_FindActualTasksImpactedByTestChanges = impactedActualIds
        Exit Function
    End If

    If changedTestIds Is Nothing Then
        Set GanttLive_FindActualTasksImpactedByTestChanges = impactedActualIds
        Exit Function
    End If

    If changedTestIds.Count = 0 Then
        Set GanttLive_FindActualTasksImpactedByTestChanges = impactedActualIds
        Exit Function
    End If

    Set mapTest = BuildColumnMap_GanttLive(tblTest)
    arr = tblTest.DataBodyRange.value
    Set hasActualById = CreateObject("Scripting.Dictionary")

    For r = 1 To UBound(arr, 1)
        idVal = Trim$(CStr(arr(r, mapTest("ID"))))
        If idVal <> "" Then
            hasActualById(idVal) = (UCase$(Trim$(CStr(arr(r, mapTest("Has Actual"))))) = "YES")
        End If
    Next r

    Set childrenByPred = GanttLive_BuildChildrenByPred_FromLinks(arr, mapTest, linksBySuccId)

    Set queue = New Collection
    Set visited = CreateObject("Scripting.Dictionary")

    For Each startId In changedTestIds.Keys

        ' IMPORTANT :
        ' la tâche modifiée elle-même doit être incluse si elle a de l’Actual
        If hasActualById.Exists(CStr(startId)) Then
            If hasActualById(CStr(startId)) Then
                impactedActualIds(CStr(startId)) = True
            End If
        End If

        If Not visited.Exists(CStr(startId)) Then
            visited(CStr(startId)) = True
            queue.Add CStr(startId)
        End If

    Next startId

    Do While queue.Count > 0

        currentId = CStr(queue(1))
        queue.Remove 1

        If childrenByPred.Exists(currentId) Then
            For Each childId In childrenByPred(currentId)

                If Not visited.Exists(CStr(childId)) Then
                    visited(CStr(childId)) = True
                    queue.Add CStr(childId)
                End If

                If hasActualById.Exists(CStr(childId)) Then
                    If hasActualById(CStr(childId)) Then
                        impactedActualIds(CStr(childId)) = True
                    End If
                End If

            Next childId
        End If

    Loop

    Set GanttLive_FindActualTasksImpactedByTestChanges = impactedActualIds

End Function

Private Function GanttLive_BuildChildrenByPred_FromLinks( _
    ByRef arr As Variant, _
    ByVal mapTest As Object, _
    ByVal linksBySuccId As Object) As Object

    Dim childrenByPred As Object
    Dim r As Long
    Dim idVal As String
    Dim succId As Variant
    Dim oneLink As Variant
    Dim predId As String

    Set childrenByPred = CreateObject("Scripting.Dictionary")

    For r = 1 To UBound(arr, 1)
        idVal = Trim$(CStr(arr(r, mapTest("ID"))))
        If idVal <> "" Then
            Set childrenByPred(idVal) = New Collection
        End If
    Next r

    If linksBySuccId Is Nothing Then
        Set GanttLive_BuildChildrenByPred_FromLinks = childrenByPred
        Exit Function
    End If

    For Each succId In linksBySuccId.Keys
        For Each oneLink In linksBySuccId(CStr(succId))
            predId = Core_GetLinkPredId(oneLink)
            If predId <> "" Then
                If Not childrenByPred.Exists(predId) Then
                    Set childrenByPred(predId) = New Collection
                End If
                childrenByPred(predId).Add CStr(succId)
            End If
        Next oneLink
    Next succId

    Set GanttLive_BuildChildrenByPred_FromLinks = childrenByPred

End Function

Private Sub GanttLive_ApplyActualImpactWarnings( _
    ByVal idToRowTest As Object, _
    ByVal warningActualIds As Object, _
    ByRef outWarnFlag() As Variant, _
    ByRef outWarnText() As Variant)

    Dim idVal As Variant
    Dim rowIndex As Long
    Dim warnText As String

    If warningActualIds Is Nothing Then Exit Sub
    If warningActualIds.Count = 0 Then Exit Sub

    warnText = "TEST INPUT IMPACTS ACTUAL TASK - LOCK MAY NOT FULLY PERSIST"

    For Each idVal In warningActualIds.Keys

        If idToRowTest.Exists(CStr(idVal)) Then
            rowIndex = CLng(idToRowTest(CStr(idVal)))
            outWarnFlag(rowIndex, 1) = "WARNING"

            If Trim$(CStr(outWarnText(rowIndex, 1))) = "" Then
                outWarnText(rowIndex, 1) = warnText
            ElseIf InStr(1, CStr(outWarnText(rowIndex, 1)), warnText, vbTextCompare) = 0 Then
                outWarnText(rowIndex, 1) = CStr(outWarnText(rowIndex, 1)) & " | " & warnText
            End If
        End If

    Next idVal

End Sub

Private Function BuildRawWbsSourceById_Live() As Object

    Dim d As Object
    Dim wsWBS As Worksheet
    Dim tblWBS As ListObject
    Dim mapWBS As Object
    Dim arr As Variant
    Dim r As Long
    Dim idVal As String

    Set d = CreateObject("Scripting.Dictionary")
    Set wsWBS = ThisWorkbook.Worksheets(WBS_SHEET)
    Set tblWBS = wsWBS.ListObjects(WBS_TABLE)

    If tblWBS.DataBodyRange Is Nothing Then
        Set BuildRawWbsSourceById_Live = d
        Exit Function
    End If

    Set mapWBS = BuildColumnMap_GanttLive(tblWBS)
    arr = tblWBS.DataBodyRange.value

    For r = 1 To UBound(arr, 1)
        idVal = Trim$(CStr(arr(r, mapWBS("ID"))))

        If idVal <> "" Then
            d(idVal) = Array( _
                GetCellValue(arr(r, mapWBS("Actual Start"))), _
                GetCellValue(arr(r, mapWBS("Actual Finish"))), _
                GetCellValue(arr(r, mapWBS("Forecast Start"))), _
                GetCellValue(arr(r, mapWBS("Forecast Finish"))), _
                GetCellValue(arr(r, mapWBS("Baseline Start"))), _
                GetCellValue(arr(r, mapWBS("Baseline Duration"))), _
                GetCellValue(arr(r, mapWBS("% Progress"))), _
                NormalizeCalendarType(arr(r, mapWBS("Cal"))) _
            )
        End If
    Next r

    Set BuildRawWbsSourceById_Live = d

End Function

Public Sub Run_Gantt_Scenario_Engine( _
    Optional ByVal silentMode As Boolean = False, _
    Optional ByRef transactionSucceeded As Variant, _
    Optional ByRef transactionMessages As Variant, _
    Optional ByRef transactionGanttRebuilt As Variant, _
    Optional ByVal recordSilentMessages As Boolean = True)

    Dim wsGantt As Worksheet
    Dim wsWBS As Worksheet
    Dim wsCalc As Worksheet
    Dim wsTest As Worksheet

    Dim tblWBS As ListObject
    Dim tblCalc As ListObject
    Dim tblTest As ListObject

    Dim mapWBS As Object
    Dim mapCalc As Object
    Dim mapCalcDriving As Object
    Dim wbsToId As Object
    Dim calcById As Object
    Dim calcRowById As Object

    Dim dataWBS As Variant
    Dim dataCalc As Variant
    Dim outArr() As Variant

    Dim rowCount As Long
    Dim calcRowCount As Long
    Dim r As Long
    Dim calcRow As Long
    Dim testRow As Long

    Dim idVal As String
    Dim wbsVal As String
    Dim taskTypeVal As String
    Dim ganttRow As Long

    Dim baseStart As Variant
    Dim baseFinish As Variant
    Dim baseDuration As Variant
    Dim baseProgress As Variant
    Dim drivingLogic As String

    Dim testStart As Variant
    Dim testFinish As Variant
    Dim testProgressRaw As Variant
    Dim testProgressNorm As Variant

    Dim inputStart As Variant
    Dim inputFinish As Variant
    Dim inputDuration As Variant
    Dim inputProgress As Variant

    Dim isSummary As Boolean
    Dim hasActual As Boolean
    Dim anyTestValue As Boolean
    Dim consoleMessages As Collection
    Dim scenarioSucceeded As Boolean
    Dim ganttRebuilt As Boolean

    On Error GoTo SafeExit

    ThisWorkbook.Init_AppEvents
    Set consoleMessages = New Collection

    Set wsGantt = ThisWorkbook.Worksheets(GANTT_SHEET)
    Set wsWBS = ThisWorkbook.Worksheets(WBS_SHEET)
    Set wsCalc = ThisWorkbook.Worksheets(CALC_SHEET)

    Set tblWBS = wsWBS.ListObjects(WBS_TABLE)
    Set tblCalc = wsCalc.ListObjects(CALC_TABLE)

    If tblWBS.DataBodyRange Is Nothing Then Exit Sub
    If tblCalc.DataBodyRange Is Nothing Then Exit Sub

    Set wsTest = Ensure_CalcGanttTest_Sheet()
    Set tblTest = Ensure_CalcGanttTest_Table(wsTest)

    Set mapWBS = BuildColumnMap_GanttLive(tblWBS)
    Set mapCalc = BuildColumnMap_GanttLive(tblCalc)
    Set mapCalcDriving = BuildCalcDrivingLogicMap_Live()
    Set wbsToId = BuildWbsToIdMapFromWBS_Live()
    Set calcById = BuildCalcByIdMap_Live(tblCalc, mapCalc)
    Set calcRowById = CreateObject("Scripting.Dictionary")

    ValidateGanttTestSourceColumns mapWBS, mapCalc
    ValidateCalcGanttTestColumns tblTest

    If Not mapCalc.Exists("Calculated Start") Then
        Err.Raise vbObjectError + 9101, "Run_Gantt_Scenario_Engine", "Missing CALC column: Calculated Start"
    End If

    If Not mapCalc.Exists("Calculated Finish") Then
        Err.Raise vbObjectError + 9102, "Run_Gantt_Scenario_Engine", "Missing CALC column: Calculated Finish"
    End If

    dataWBS = tblWBS.DataBodyRange.value
    dataCalc = tblCalc.DataBodyRange.value

    rowCount = UBound(dataWBS, 1)
    calcRowCount = UBound(dataCalc, 1)

    For r = 1 To calcRowCount
        idVal = Trim$(CStr(dataCalc(r, mapCalc("ID"))))
        If idVal <> "" Then
            If Not calcRowById.Exists(idVal) Then
                calcRowById(idVal) = r
            End If
        End If
    Next r

    ResizeTableToRowCount_Generic tblTest, rowCount
    ReDim outArr(1 To rowCount, 1 To tblTest.ListColumns.Count)

    testRow = 0

    For r = 1 To rowCount

        idVal = Trim$(CStr(dataWBS(r, mapWBS("ID"))))
        wbsVal = NormalizeWBS(CStr(dataWBS(r, mapWBS("WBS"))))
        taskTypeVal = Trim$(CStr(dataWBS(r, mapWBS("Task Type"))))

        testRow = testRow + 1

        baseStart = Empty
        baseFinish = Empty
        baseDuration = Empty
        baseProgress = GetCellValue(dataWBS(r, mapWBS("% Progress")))
        drivingLogic = ""

        If calcRowById.Exists(idVal) Then

            calcRow = CLng(calcRowById(idVal))

            baseStart = GetCellValue(dataCalc(calcRow, mapCalc("Calculated Start")))
            baseFinish = GetCellValue(dataCalc(calcRow, mapCalc("Calculated Finish")))

            If HasValue(baseStart) And HasValue(baseFinish) Then
                baseDuration = CDbl(baseFinish) - CDbl(baseStart) + 1
            Else
                baseDuration = Empty
            End If

            If mapCalcDriving.Exists(idVal) Then
                drivingLogic = CStr(mapCalcDriving(idVal))
            ElseIf mapCalc.Exists("Driving Logic") Then
                drivingLogic = CStr(dataCalc(calcRow, mapCalc("Driving Logic")))
            End If

        End If

        isSummary = (UCase$(Trim$(drivingLogic)) = "SUMMARY")
        hasActual = False

        testStart = Empty
        testFinish = Empty
        testProgressRaw = Empty
        testProgressNorm = Empty
        inputStart = Empty
        inputFinish = Empty
        inputDuration = Empty
        inputProgress = Empty
        anyTestValue = False

        ganttRow = FindGanttRowByWBS(wsGantt, wbsVal)

        If ganttRow > 0 Then
            testStart = GetCellValue(wsGantt.Cells(ganttRow, COL_TEST_START).value)
            testFinish = GetCellValue(wsGantt.Cells(ganttRow, COL_TEST_FINISH).value)
            testProgressRaw = GetCellValue(wsGantt.Cells(ganttRow, COL_TEST_PROGRESS).value)
        End If

        If HasValue(testStart) Or HasValue(testFinish) Or HasValue(testProgressRaw) Then
            anyTestValue = True
        End If

        If HasValue(testProgressRaw) Then
            testProgressNorm = NormalizePercentInput(testProgressRaw)
        Else
            testProgressNorm = Empty
        End If

        inputStart = baseStart
        inputFinish = Empty
        inputDuration = baseDuration
        inputProgress = baseProgress

        If HasValue(testStart) Then inputStart = testStart
        If HasValue(testFinish) Then inputFinish = testFinish

        If HasValue(testStart) And HasValue(testFinish) Then
            inputDuration = CDbl(testFinish) - CDbl(testStart) + 1
        End If

        If HasValue(testProgressNorm) Then
            inputProgress = testProgressNorm
        End If

        outArr(testRow, GetColumnIndex_GanttLive(tblTest, "ID")) = idVal
        outArr(testRow, GetColumnIndex_GanttLive(tblTest, "WBS")) = wbsVal
        outArr(testRow, GetColumnIndex_GanttLive(tblTest, "Task Type")) = taskTypeVal
        outArr(testRow, GetColumnIndex_GanttLive(tblTest, "Cal")) = NormalizeCalendarType(dataWBS(r, mapWBS("Cal")))
        outArr(testRow, GetColumnIndex_GanttLive(tblTest, "Is Summary")) = IIf(isSummary, "YES", "NO")
        outArr(testRow, GetColumnIndex_GanttLive(tblTest, "Parent ID")) = GetParentIdFromWBS(wbsVal, wbsToId)

        outArr(testRow, GetColumnIndex_GanttLive(tblTest, "Base Start")) = baseStart
        outArr(testRow, GetColumnIndex_GanttLive(tblTest, "Base Finish")) = baseFinish
        outArr(testRow, GetColumnIndex_GanttLive(tblTest, "Base Duration")) = baseDuration
        outArr(testRow, GetColumnIndex_GanttLive(tblTest, "Base Progress")) = baseProgress
        outArr(testRow, GetColumnIndex_GanttLive(tblTest, "Driving Logic")) = drivingLogic
        outArr(testRow, GetColumnIndex_GanttLive(tblTest, "Has Actual")) = IIf(hasActual, "YES", "NO")

        outArr(testRow, GetColumnIndex_GanttLive(tblTest, "Test Start")) = testStart
        outArr(testRow, GetColumnIndex_GanttLive(tblTest, "Test Finish")) = testFinish
        outArr(testRow, GetColumnIndex_GanttLive(tblTest, "Test % Raw")) = testProgressRaw
        outArr(testRow, GetColumnIndex_GanttLive(tblTest, "Test % Normalized")) = testProgressNorm

        outArr(testRow, GetColumnIndex_GanttLive(tblTest, "Input Start")) = inputStart
        outArr(testRow, GetColumnIndex_GanttLive(tblTest, "Input Finish")) = inputFinish
        outArr(testRow, GetColumnIndex_GanttLive(tblTest, "Input Duration")) = inputDuration
        outArr(testRow, GetColumnIndex_GanttLive(tblTest, "Input Progress")) = inputProgress

        If calcById.Exists(idVal) Then
            outArr(testRow, GetColumnIndex_GanttLive(tblTest, "Predecessors")) = calcById(idVal)
        End If

        outArr(testRow, GetColumnIndex_GanttLive(tblTest, "Lag")) = Empty
        outArr(testRow, GetColumnIndex_GanttLive(tblTest, "Any Test Value")) = IIf(anyTestValue, "YES", "NO")
        outArr(testRow, GetColumnIndex_GanttLive(tblTest, "Warning Flag")) = ""
        outArr(testRow, GetColumnIndex_GanttLive(tblTest, "Warning Text")) = ""
        outArr(testRow, GetColumnIndex_GanttLive(tblTest, "Error Flag")) = ""

    Next r

    FormatCalcGanttTestColumns tblTest
    tblTest.DataBodyRange.value = outArr
    FormatCalcGanttTestColumns tblTest

    If Not Run_Gantt_Scenario_Backend(tblTest, consoleMessages) Then
        GanttLive_AbortTestEngine wsGantt
        If silentMode Then
            If recordSilentMessages Then CalcBridge_RecordPlanningMessages consoleMessages, "Run_Gantt_Scenario_Engine"
        Else
            CalcBridge_ShowPlanningConsole consoleMessages
        End If
        GoTo CleanExit
    End If

    GanttLive_ApplyScenarioRender wsGantt
    scenarioSucceeded = True
    ganttRebuilt = True

    If silentMode Then
        If recordSilentMessages Then CalcBridge_RecordPlanningMessages consoleMessages, "Run_Gantt_Scenario_Engine"
        GoTo CleanExit
    End If

    GanttLive_AddBiConsoleMessage consoleMessages, "INFO", _
        "Scénario mis à jour.", _
        "Scenario updated."

    CalcBridge_ShowPlanningConsole consoleMessages

CleanExit:
    If Not IsMissing(transactionSucceeded) Then transactionSucceeded = scenarioSucceeded
    If Not IsMissing(transactionGanttRebuilt) Then transactionGanttRebuilt = ganttRebuilt
    If Not IsMissing(transactionMessages) Then Set transactionMessages = consoleMessages
    Exit Sub

SafeExit:
    GanttLive_SafeExit_TestEngine

    If Err.Number <> 0 Then
        If consoleMessages Is Nothing Then Set consoleMessages = New Collection

        GanttLive_AddVbaOrStructuredError consoleMessages, "Run_Gantt_Scenario_Engine", Err.Description

        If silentMode Then
            If recordSilentMessages Then CalcBridge_RecordPlanningMessages consoleMessages, "Run_Gantt_Scenario_Engine"
        Else
            CalcBridge_ShowPlanningConsole consoleMessages
        End If
    End If

    Resume CleanExit

End Sub

Public Function GanttLive_RunScenarioTransaction( _
    ByRef consoleMessages As Collection, _
    ByRef ganttRebuilt As Boolean) As Boolean

    Dim transactionSucceeded As Variant
    Dim transactionMessages As Variant
    Dim transactionGanttRebuilt As Variant

    Set consoleMessages = New Collection
    ganttRebuilt = False

    Run_Gantt_Scenario_Engine _
        True, _
        transactionSucceeded, _
        transactionMessages, _
        transactionGanttRebuilt, _
        False

    If IsObject(transactionMessages) Then
        Set consoleMessages = transactionMessages
    End If

    If Not IsEmpty(transactionGanttRebuilt) Then
        ganttRebuilt = CBool(transactionGanttRebuilt)
    End If

    If Not IsEmpty(transactionSucceeded) Then
        GanttLive_RunScenarioTransaction = CBool(transactionSucceeded)
    End If

End Function

Private Function Run_Gantt_Scenario_Backend( _
    ByVal tblTest As ListObject, _
    Optional ByVal consoleMessages As Variant) As Boolean

    Dim dataCore As Variant
    Dim mapCore As Object
    Dim linksBySuccId As Object

    Dim outStart() As Variant
    Dim outFinish() As Variant
    Dim outDuration() As Variant
    Dim outProgress() As Variant
    Dim outWarnFlag() As Variant
    Dim outWarnText() As Variant
    Dim outErrFlag() As Variant

    Dim idToRowTest As Object
    Dim idToWbs As Object
    Dim errorIds As Object
    Dim rootErrorIds As Object
    Dim dependencyDiagnostics As Object
    Dim constraintDiagnostics As Object
    Dim cascadeDiagnostics As Object

    Dim rowCount As Long
    Dim r As Long
    Dim idVal As String
    Dim errorMsg As String

    On Error GoTo SafeExit

    Run_Gantt_Scenario_Backend = False

    If tblTest Is Nothing Then Exit Function
    If tblTest.DataBodyRange Is Nothing Then
        Run_Gantt_Scenario_Backend = True
        Exit Function
    End If

    rowCount = tblTest.ListRows.Count

    ReDim outStart(1 To rowCount, 1 To 1)
    ReDim outFinish(1 To rowCount, 1 To 1)
    ReDim outDuration(1 To rowCount, 1 To 1)
    ReDim outProgress(1 To rowCount, 1 To 1)
    ReDim outWarnFlag(1 To rowCount, 1 To 1)
    ReDim outWarnText(1 To rowCount, 1 To 1)
    ReDim outErrFlag(1 To rowCount, 1 To 1)

    Set idToRowTest = CreateObject("Scripting.Dictionary")
    Set idToWbs = CreateObject("Scripting.Dictionary")
    Set errorIds = CreateObject("Scripting.Dictionary")
    Set rootErrorIds = CreateObject("Scripting.Dictionary")
    Set dependencyDiagnostics = CreateObject("Scripting.Dictionary")
    Set constraintDiagnostics = CreateObject("Scripting.Dictionary")
    Set cascadeDiagnostics = CreateObject("Scripting.Dictionary")

    BuildGanttScenarioCoreDataset tblTest, dataCore, mapCore, idToRowTest, idToWbs

    Set linksBySuccId = BuildCoreLinksBySucc_FromLogicLinksTable_Expanded( _
        ThisWorkbook.Worksheets(CALC_SHEET).ListObjects(CALC_TABLE))

    Run_Calc_Core dataCore, mapCore, linksBySuccId, , dependencyDiagnostics, constraintDiagnostics, cascadeDiagnostics

    For r = 1 To rowCount

        idVal = Trim$(CStr(dataCore(r, mapCore("ID"))))
        If idVal = "" Then GoTo NextRow

        outStart(r, 1) = dataCore(r, mapCore("Calculated Start"))
        outFinish(r, 1) = dataCore(r, mapCore("Calculated Finish"))
        outDuration(r, 1) = dataCore(r, mapCore("Calculated Duration"))
        outProgress(r, 1) = dataCore(r, mapCore("Input Progress"))

        If UCase$(Trim$(CStr(dataCore(r, mapCore("Error flag"))))) = "ERROR" Then

            outErrFlag(r, 1) = "ERROR"
            errorIds(idVal) = True

            errorMsg = Trim$(CStr(dataCore(r, mapCore("ErrorMsg"))))

            If errorMsg <> "" Then
                outWarnText(r, 1) = errorMsg
            End If

            If Not GanttLive_IsInheritedCoreError(errorMsg) Then
                rootErrorIds(idVal) = True
            End If

        Else
            outErrFlag(r, 1) = ""
        End If

NextRow:
    Next r

    GanttLive_RemoveDerivedLOERootErrors dataCore, mapCore, errorIds, rootErrorIds

    tblTest.ListColumns("Calc Test Start").DataBodyRange.value = outStart
    tblTest.ListColumns("Calc Test Finish").DataBodyRange.value = outFinish
    tblTest.ListColumns("Calc Test Duration").DataBodyRange.value = outDuration
    tblTest.ListColumns("Calc Test Progress").DataBodyRange.value = outProgress
    tblTest.ListColumns("Warning Flag").DataBodyRange.value = outWarnFlag
    tblTest.ListColumns("Warning Text").DataBodyRange.value = outWarnText
    tblTest.ListColumns("Error Flag").DataBodyRange.value = outErrFlag

    FormatCalcGanttTestColumns tblTest

    If errorIds.Count > 0 Then
        If rootErrorIds.Count > 0 Then
            CalcBridge_AppendCoreErrorMessagesFromData consoleMessages, dataCore, mapCore, rootErrorIds, "SCENARIO", dependencyDiagnostics, constraintDiagnostics, cascadeDiagnostics
        Else
            ShowGanttLiveGroupedMessage errorIds, idToWbs, _
                "Erreur de calcul dans le scénario", "corriger les valeurs de test ou la logique amont", _
                "Calculation error in scenario", "fix test values or upstream logic", vbCritical, consoleMessages
        End If

        Exit Function
    End If

    Run_Gantt_Scenario_Backend = True
    Exit Function

SafeExit:
    If Err.Number <> 0 Then
        GanttLive_AddVbaOrStructuredError consoleMessages, "Run_Gantt_Scenario_Backend", Err.Description
    End If

End Function

Private Sub BuildGanttScenarioCoreDataset( _
    ByVal tblTest As ListObject, _
    ByRef dataCore As Variant, _
    ByRef mapCore As Object, _
    ByVal idToRowTest As Object, _
    ByVal idToWbs As Object)

    Dim arrTest As Variant
    Dim rowCount As Long
    Dim r As Long

    Dim headers As Variant
    Dim c As Long

    Dim idVal As String
    Dim wbsVal As String
    Dim taskTypeVal As String
    Dim isSummary As String
    Dim parentId As String
    Dim taskName As String
    Dim taskNameById As Object

    Dim baseStart As Variant
    Dim baseDuration As Variant
    Dim baseProgress As Variant

    Dim testStart As Variant
    Dim testFinish As Variant
    Dim testProgress As Variant

    Dim inputStart As Variant
    Dim inputDuration As Variant
    Dim inputProgress As Variant

    headers = Array( _
        "ID", "WBS", "Task Name", "Task Type", "Cal", "ParentID", "IsSummary", _
        "Actual Start", "Actual Finish", "Forecast Start", "Forecast Finish", _
        "Baseline Start", "Baseline Duration", "Constraint Active", _
        "Start Constraint Type", "Start Constraint Date", _
        "Finish Constraint Type", "Finish Constraint Date", _
        "Calculated Start", "Calculated Finish", "Calculated Duration", _
        "Error flag", "ErrorMsg", "Input Progress" _
    )

    Set mapCore = CreateObject("Scripting.Dictionary")
    For c = LBound(headers) To UBound(headers)
        mapCore(CStr(headers(c))) = c + 1
    Next c

    arrTest = tblTest.DataBodyRange.value
    rowCount = UBound(arrTest, 1)

    ReDim dataCore(1 To rowCount, 1 To UBound(headers) + 1)

    Set taskNameById = BuildTaskNameByIdFromWbs_Live()

    For r = 1 To rowCount

        idVal = Trim$(CStr(arrTest(r, GetColumnIndex_GanttLive(tblTest, "ID"))))
        wbsVal = NormalizeWBS(CStr(arrTest(r, GetColumnIndex_GanttLive(tblTest, "WBS"))))
        taskTypeVal = Trim$(CStr(arrTest(r, GetColumnIndex_GanttLive(tblTest, "Task Type"))))
        isSummary = UCase$(Trim$(CStr(arrTest(r, GetColumnIndex_GanttLive(tblTest, "Is Summary")))))
        parentId = Trim$(CStr(arrTest(r, GetColumnIndex_GanttLive(tblTest, "Parent ID"))))
        taskName = vbNullString
        If taskNameById.Exists(idVal) Then taskName = Trim$(CStr(taskNameById(idVal)))

        idToRowTest(idVal) = r
        idToWbs(idVal) = wbsVal

        baseStart = GetCellValue(arrTest(r, GetColumnIndex_GanttLive(tblTest, "Base Start")))
        baseDuration = GetCellValue(arrTest(r, GetColumnIndex_GanttLive(tblTest, "Base Duration")))
        baseProgress = GetCellValue(arrTest(r, GetColumnIndex_GanttLive(tblTest, "Base Progress")))

        testStart = GetCellValue(arrTest(r, GetColumnIndex_GanttLive(tblTest, "Test Start")))
        testFinish = GetCellValue(arrTest(r, GetColumnIndex_GanttLive(tblTest, "Test Finish")))
        testProgress = GetCellValue(arrTest(r, GetColumnIndex_GanttLive(tblTest, "Test % Normalized")))

        inputStart = baseStart
        inputDuration = baseDuration
        inputProgress = baseProgress

        If HasValue(testStart) Then
            inputStart = testStart
        End If

        If HasValue(testStart) And HasValue(testFinish) Then
            inputDuration = CDbl(testFinish) - CDbl(testStart) + 1
        End If

        If HasValue(testProgress) Then
            inputProgress = testProgress
        End If

        dataCore(r, mapCore("ID")) = idVal
        dataCore(r, mapCore("WBS")) = wbsVal
        dataCore(r, mapCore("Task Name")) = taskName
        dataCore(r, mapCore("Task Type")) = taskTypeVal
        dataCore(r, mapCore("Cal")) = NormalizeCalendarType(arrTest(r, GetColumnIndex_GanttLive(tblTest, "Cal")))
        dataCore(r, mapCore("ParentID")) = parentId
        dataCore(r, mapCore("IsSummary")) = IIf(isSummary = "YES", True, False)

        dataCore(r, mapCore("Actual Start")) = Empty
        dataCore(r, mapCore("Actual Finish")) = Empty
        dataCore(r, mapCore("Forecast Start")) = Empty
        dataCore(r, mapCore("Forecast Finish")) = Empty

        dataCore(r, mapCore("Baseline Start")) = inputStart
        dataCore(r, mapCore("Baseline Duration")) = inputDuration

        dataCore(r, mapCore("Constraint Active")) = "No"
        dataCore(r, mapCore("Start Constraint Type")) = vbNullString
        dataCore(r, mapCore("Start Constraint Date")) = Empty
        dataCore(r, mapCore("Finish Constraint Type")) = vbNullString
        dataCore(r, mapCore("Finish Constraint Date")) = Empty

        dataCore(r, mapCore("Calculated Start")) = Empty
        dataCore(r, mapCore("Calculated Finish")) = Empty
        dataCore(r, mapCore("Calculated Duration")) = Empty
        dataCore(r, mapCore("Error flag")) = ""
        dataCore(r, mapCore("ErrorMsg")) = ""
        dataCore(r, mapCore("Input Progress")) = inputProgress

    Next r

End Sub


Public Sub GanttLive_SetActiveSimulationMode(ByVal modeName As String)
    gActiveSimulationMode = UCase$(Trim$(modeName))
End Sub

Public Sub GanttLive_ClearActiveSimulationMode()
    gActiveSimulationMode = ""
End Sub

Public Function GanttLive_GetActiveSimulationMode() As String
    GanttLive_GetActiveSimulationMode = UCase$(Trim$(gActiveSimulationMode))
End Function

Public Function GanttLive_IsScenarioActive() As Boolean
    GanttLive_IsScenarioActive = (UCase$(Trim$(gActiveSimulationMode)) = "SCENARIO")
End Function

Public Function GanttLive_IsLiveTestActive() As Boolean
    GanttLive_IsLiveTestActive = (UCase$(Trim$(gActiveSimulationMode)) = "TEST")
End Function

Private Sub GanttLive_ApplyScenarioRender(ByVal wsGantt As Worksheet)

    SetGanttPreserveTestInputs True
    GanttLive_RequestScenarioRender
    Refresh_Gantt
    GanttLive_SetActiveSimulationMode "SCENARIO"
    SetGanttPreserveTestInputs False

    If Not wsGantt Is Nothing Then
        wsGantt.Activate
    End If

End Sub

Private Function GanttLive_HasAnyRenderableTestDelta(ByVal tblTest As ListObject) As Boolean

    Dim baseById As Object
    Dim testById As Object
    Dim arr As Variant
    Dim mapTest As Object
    Dim r As Long
    Dim idVal As String

    If tblTest Is Nothing Then Exit Function
    If tblTest.DataBodyRange Is Nothing Then Exit Function

    Set baseById = GanttLive_BuildBaseByIdMap()
    Set testById = GanttLive_BuildTestByIdMap()
    Set mapTest = BuildColumnMap_GanttLive(tblTest)

    arr = tblTest.DataBodyRange.value

    For r = 1 To UBound(arr, 1)

        idVal = Trim$(CStr(arr(r, mapTest("ID"))))
        If idVal = "" Then GoTo NextRow

        If UCase$(Trim$(CStr(arr(r, mapTest("Is Summary"))))) = "YES" Then GoTo NextRow

        If GanttLive_HasRenderableTestDelta(idVal, baseById, testById) Then
            GanttLive_HasAnyRenderableTestDelta = True
            Exit Function
        End If

NextRow:
    Next r

End Function

Private Function GanttLive_IsInheritedCoreError(ByVal errMsg As String) As Boolean

    Dim txt As String

    txt = Trim$(CStr(errMsg))

    GanttLive_IsInheritedCoreError = _
        (InStr(1, txt, "Blocked by predecessor error", vbTextCompare) > 0) Or _
        (InStr(1, txt, "Blocked by predecessor chain", vbTextCompare) > 0)

End Function

Private Sub GanttLive_RemoveDerivedLOERootErrors( _
    ByRef dataCore As Variant, _
    ByVal mapCore As Object, _
    ByVal errorIds As Object, _
    ByVal rootErrorIds As Object)

    Dim r As Long
    Dim idVal As String
    Dim errMsg As String
    Dim removeIds As Object
    Dim oneId As Variant

    If errorIds Is Nothing Then Exit Sub
    If rootErrorIds Is Nothing Then Exit Sub
    If rootErrorIds.Count = 0 Then Exit Sub

    Set removeIds = CreateObject("Scripting.Dictionary")

    For r = 1 To UBound(dataCore, 1)
        idVal = Trim$(CStr(dataCore(r, mapCore("ID"))))
        If idVal <> "" Then
            If rootErrorIds.Exists(idVal) Then
                errMsg = Trim$(CStr(dataCore(r, mapCore("ErrorMsg"))))
                If GanttLive_IsDerivedLOEPredecessorError(errMsg, errorIds) Then removeIds(idVal) = True
            End If
        End If
    Next r

    For Each oneId In removeIds.Keys
        If rootErrorIds.Exists(CStr(oneId)) Then rootErrorIds.Remove CStr(oneId)
    Next oneId

End Sub

Private Function GanttLive_IsDerivedLOEPredecessorError(ByVal errMsg As String, ByVal errorIds As Object) As Boolean

    Dim predId As String

    predId = GanttLive_ExtractLOEBlockedPredecessorId(errMsg)
    If predId = "" Then Exit Function

    GanttLive_IsDerivedLOEPredecessorError = errorIds.Exists(predId)

End Function

Private Function GanttLive_ExtractLOEBlockedPredecessorId(ByVal errMsg As String) As String

    Dim txt As String
    Dim marker As String
    Dim pos As Long
    Dim tail As String
    Dim i As Long
    Dim ch As String
    Dim result As String

    txt = Trim$(CStr(errMsg))
    If InStr(1, txt, "LOE blocked by SS predecessor error: ID", vbTextCompare) = 0 And _
       InStr(1, txt, "LOE blocked by FF predecessor error: ID", vbTextCompare) = 0 Then Exit Function

    marker = "error: ID"
    pos = InStr(1, txt, marker, vbTextCompare)
    If pos = 0 Then Exit Function

    tail = Trim$(Mid$(txt, pos + Len(marker)))
    For i = 1 To Len(tail)
        ch = Mid$(tail, i, 1)
        If ch >= "0" And ch <= "9" Then
            result = result & ch
        ElseIf result <> "" Then
            Exit For
        End If
    Next i

    GanttLive_ExtractLOEBlockedPredecessorId = result

End Function

Private Sub ShowGanttLiveUpstreamViolationMessage( _
    ByVal idsDict As Object, _
    ByVal idToWbs As Object, _
    ByVal frProblem As String, _
    ByVal frAction As String, _
    ByVal enProblem As String, _
    ByVal enAction As String, _
    ByVal boxStyle As VbMsgBoxStyle, _
    Optional ByVal consoleMessages As Variant)

    Dim itemsLine As String
    Dim msg As String
    Dim msgType As String
    Dim localMessages As Collection

    If idsDict Is Nothing Then Exit Sub
    If idsDict.Count = 0 Then Exit Sub

    itemsLine = BuildGanttLiveUpstreamViolationItems(idsDict, idToWbs, 20)
    msgType = GanttLive_MessageTypeFromMsgBoxStyle(boxStyle)

    msg = _
        "FR:" & vbCrLf & _
        frProblem & vbCrLf & _
        "-> " & frAction & vbCrLf & vbCrLf & _
        "Tâches : " & itemsLine & vbCrLf & vbCrLf & _
        "EN:" & vbCrLf & _
        enProblem & vbCrLf & _
        "-> " & enAction & vbCrLf & vbCrLf & _
        "Tasks: " & itemsLine

    If GanttLive_HasConsoleCollection(consoleMessages) Then
        CalcBridge_AddConsoleMessage consoleMessages, msgType, msg
    Else
        Set localMessages = New Collection
        CalcBridge_AddConsoleMessage localMessages, msgType, msg
        CalcBridge_ShowPlanningConsole localMessages
    End If

End Sub

Private Function BuildGanttLiveUpstreamViolationItems( _
    ByVal idsDict As Object, _
    ByVal idToWbs As Object, _
    ByVal maxItems As Long) As String

    Dim result As String
    Dim key As Variant
    Dim countShown As Long
    Dim totalCount As Long
    Dim wbsVal As String

    result = ""
    countShown = 0
    totalCount = idsDict.Count

    For Each key In idsDict.Keys

        countShown = countShown + 1

        If countShown <= maxItems Then

            If Not idToWbs Is Nothing Then
                If idToWbs.Exists(CStr(key)) Then
                    wbsVal = NormalizeWBS(CStr(idToWbs(CStr(key))))
                Else
                    wbsVal = "-"
                End If
            Else
                wbsVal = "-"
            End If

            If result <> "" Then result = result & " / "
            result = result & CStr(key) & " (" & wbsVal & ")"

        Else
            Exit For
        End If

    Next key

    If totalCount > maxItems Then
        result = result & " / +" & CStr(totalCount - maxItems)
    End If

    BuildGanttLiveUpstreamViolationItems = result

End Function

Private Function GanttLive_CalcGanttTestHasErrors(ByVal tblTest As ListObject) As Boolean

    Dim mapTest As Object
    Dim arr As Variant
    Dim r As Long

    On Error GoTo SafeExit

    GanttLive_CalcGanttTestHasErrors = False

    If tblTest Is Nothing Then Exit Function
    If tblTest.DataBodyRange Is Nothing Then Exit Function

    Set mapTest = BuildColumnMap_GanttLive(tblTest)

    If Not mapTest.Exists("Error Flag") Then Exit Function

    arr = tblTest.DataBodyRange.value

    For r = 1 To UBound(arr, 1)
        If UCase$(Trim$(CStr(arr(r, mapTest("Error Flag"))))) = "ERROR" Then
            GanttLive_CalcGanttTestHasErrors = True
            Exit Function
        End If
    Next r

    Exit Function

SafeExit:
    GanttLive_CalcGanttTestHasErrors = True

End Function

Private Function GanttLive_HasConsoleCollection(Optional ByVal consoleMessages As Variant) As Boolean

    On Error GoTo SafeExit

    If IsMissing(consoleMessages) Then Exit Function
    If IsObject(consoleMessages) Then
        If Not consoleMessages Is Nothing Then
            GanttLive_HasConsoleCollection = True
        End If
    End If

SafeExit:
End Function

Private Function GanttLive_MessageTypeFromMsgBoxStyle(ByVal boxStyle As VbMsgBoxStyle) As String

    If (boxStyle And vbCritical) = vbCritical Then
        GanttLive_MessageTypeFromMsgBoxStyle = "STOP"
    ElseIf (boxStyle And vbExclamation) = vbExclamation Then
        GanttLive_MessageTypeFromMsgBoxStyle = "WARNING"
    Else
        GanttLive_MessageTypeFromMsgBoxStyle = "INFO"
    End If

End Function


Private Function GanttLive_IsStructuredBiMessage(ByVal msgText As String) As Boolean

    Dim txt As String

    txt = LTrim$(CStr(msgText))

    GanttLive_IsStructuredBiMessage = _
        (Left$(txt, 3) = "FR:" And InStr(1, txt, "EN:", vbTextCompare) > 0)

End Function

Private Sub GanttLive_AddVbaOrStructuredError( _
    ByVal consoleMessages As Collection, _
    ByVal functionName As String, _
    ByVal errDescription As String)

    If consoleMessages Is Nothing Then Exit Sub

    If GanttLive_IsStructuredBiMessage(errDescription) Then
        CalcBridge_AddConsoleMessage consoleMessages, "STOP", Trim$(CStr(errDescription))
    Else
        GanttLive_AddBiConsoleMessage consoleMessages, "STOP", _
            "Erreur VBA dans " & functionName & vbCrLf & _
            "-> " & errDescription, _
            "VBA error in " & functionName & vbCrLf & _
            "-> " & errDescription
    End If

End Sub

Private Sub GanttLive_AddBiConsoleMessage( _
    ByVal consoleMessages As Collection, _
    ByVal msgType As String, _
    ByVal frText As String, _
    ByVal enText As String)

    Dim msg As String

    If consoleMessages Is Nothing Then Exit Sub

    msg = _
        "FR:" & vbCrLf & _
        frText & vbCrLf & vbCrLf & _
        "EN:" & vbCrLf & _
        enText

    CalcBridge_AddConsoleMessage consoleMessages, msgType, msg

End Sub




