Attribute VB_Name = "mod_GanttDragWatch"
Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Function SetTimer Lib "user32" ( _
        ByVal hwnd As LongPtr, _
        ByVal nIDEvent As LongPtr, _
        ByVal uElapse As Long, _
        ByVal lpTimerFunc As LongPtr) As LongPtr

    Private Declare PtrSafe Function KillTimer Lib "user32" ( _
        ByVal hwnd As LongPtr, _
        ByVal nIDEvent As LongPtr) As Long
#Else
    Private Declare Function SetTimer Lib "user32" ( _
        ByVal hwnd As Long, _
        ByVal nIDEvent As Long, _
        ByVal uElapse As Long, _
        ByVal lpTimerFunc As Long) As Long

    Private Declare Function KillTimer Lib "user32" ( _
        ByVal hwnd As Long, _
        ByVal nIDEvent As Long) As Long
#End If

Private Const GANTT_DRAG_SHEET As String = "GANTT"
Private Const GANTT_DRAG_SCALE_DAY As String = "DAY"
Private Const GANTT_DRAG_SCALE_WEEK As String = "WEEK"
Private Const GANTT_DRAG_SCALE_MONTH As String = "MONTH"
Private Const GANTT_DRAG_TIMER_MS As Long = 200
Private Const GANTT_DRAG_FIRST_TASK_ROW As Long = 5
Private Const GANTT_DRAG_HEADER_ROW As Long = 4
Private Const GANTT_DRAG_FIRST_TIMELINE_COL As Long = 11
Private Const GANTT_DRAG_COL_TEST_START As Long = 5
Private Const GANTT_DRAG_COL_TEST_FINISH As Long = 6
Private Const GANTT_DRAG_CHANGE_TOLERANCE As Double = 0.2
Private Const GANTT_DRAG_MAX_STRUCTURAL_ERRORS As Long = 3

#If VBA7 Then
    Private gTimerId As LongPtr
#Else
    Private gTimerId As Long
#End If

Private gWatchRequested As Boolean
Private gWatchEnabled As Boolean
Private gInTick As Boolean
Private gSuspendDepth As Long
Private gStructuralErrorCount As Long
Private gShapeState As Object
Private gLastDebugStatus As String
Private gTransactionActive As Boolean
Private gLastTransactionResult As String
Private gLastWrittenCells As String
Private gTransactionCount As Long

Public Sub GanttDrag_StartWatch()

    gWatchRequested = True
    GanttDrag_StartRuntime

End Sub

Public Sub GanttDrag_StopWatch(Optional ByVal showStatus As Boolean = False)

    GanttDrag_StopRuntime True, showStatus

End Sub

Public Sub GanttDrag_RebuildWatchMaps()

    Dim ws As Worksheet
    Dim shp As Shape
    Dim rowIndex As Long
    Dim taskType As String

    Set gShapeState = CreateObject("Scripting.Dictionary")

    If Not GanttDrag_IsGanttSheetActive(ws) Then Exit Sub
    If Not GanttDrag_IsSupportedTimelineScale() Then Exit Sub

    For Each shp In ws.Shapes
        rowIndex = 0
        taskType = vbNullString

        If GanttDrag_IsEligibleShape(ws, shp, rowIndex, taskType) Then
            GanttDrag_SaveShapeState shp, rowIndex, taskType
        End If
    Next shp

End Sub

Public Function GanttDrag_IsWatching() As Boolean

    GanttDrag_IsWatching = (gWatchEnabled And gTimerId <> 0)

End Function

Public Function GanttDrag_LastDebugStatus() As String

    GanttDrag_LastDebugStatus = gLastDebugStatus

End Function

Public Function GanttDrag_LastTransactionResult() As String

    GanttDrag_LastTransactionResult = gLastTransactionResult

End Function

Public Function GanttDrag_LastWrittenCells() As String

    GanttDrag_LastWrittenCells = gLastWrittenCells

End Function

Public Function GanttDrag_TransactionCount() As Long

    GanttDrag_TransactionCount = gTransactionCount

End Function

Public Function GanttDrag_IsShapeWatched(ByVal shapeName As String) As Boolean

    If gShapeState Is Nothing Then Exit Function
    GanttDrag_IsShapeWatched = gShapeState.Exists(shapeName)

End Function

Public Sub GanttDrag_ReconcileWatchState()

    If Not gWatchRequested Then
        GanttDrag_StopRuntime False, False
        Exit Sub
    End If

    If Not GanttDrag_CanStartWatch() Then
        GanttDrag_StopRuntime False, False
        Exit Sub
    End If

    If GanttDrag_IsWatching() Then
        GanttDrag_RebuildWatchMaps
    Else
        GanttDrag_StartRuntime
    End If

End Sub

Public Sub GanttDrag_PauseForLifecycle()

    GanttDrag_StopRuntime False, False

End Sub

Private Sub GanttDrag_StartRuntime()

    If Not GanttDrag_CanStartWatch() Then
        GanttDrag_StopRuntime False, False
        Exit Sub
    End If

    GanttDrag_StopRuntime False, False
    GanttDrag_RebuildWatchMaps

    If gShapeState Is Nothing Then Exit Sub
    If gShapeState.Count = 0 Then Exit Sub

    gStructuralErrorCount = 0
    gInTick = False
    gTransactionActive = False
    gLastTransactionResult = vbNullString
    gLastWrittenCells = vbNullString
    gTransactionCount = 0

#If VBA7 Then
    gTimerId = SetTimer(0, 0, GANTT_DRAG_TIMER_MS, AddressOf GanttDrag_TimerProc)
#Else
    gTimerId = SetTimer(0, 0, GANTT_DRAG_TIMER_MS, AddressOf GanttDrag_TimerProc)
#End If

    gWatchEnabled = (gTimerId <> 0)
    If Not gWatchEnabled Then gWatchRequested = False

End Sub

Private Sub GanttDrag_StopRuntime( _
    ByVal clearRequest As Boolean, _
    ByVal showStatus As Boolean)

    On Error Resume Next

    gWatchEnabled = False
    gInTick = False
    gSuspendDepth = 0
    gTransactionActive = False

    If gTimerId <> 0 Then
        KillTimer 0, gTimerId
        gTimerId = 0
    End If

    Set gShapeState = Nothing
    If clearRequest Then gWatchRequested = False

    If showStatus Then Debug.Print "Gantt Drag Watch stopped."

    On Error GoTo 0

End Sub

#If VBA7 Then
Private Sub GanttDrag_TimerProc( _
    ByVal hwnd As LongPtr, _
    ByVal uMsg As Long, _
    ByVal idEvent As LongPtr, _
    ByVal dwTime As Long)
#Else
Private Sub GanttDrag_TimerProc( _
    ByVal hwnd As Long, _
    ByVal uMsg As Long, _
    ByVal idEvent As Long, _
    ByVal dwTime As Long)
#End If

    On Error GoTo StructuralError

    If Not gWatchEnabled Then Exit Sub
    If gInTick Then Exit Sub
    If gSuspendDepth > 0 Then Exit Sub

    gInTick = True
    GanttDrag_WatchTick
    gStructuralErrorCount = 0

SafeExit:
    gInTick = False
    Exit Sub

StructuralError:
    gStructuralErrorCount = gStructuralErrorCount + 1
    If gStructuralErrorCount >= GANTT_DRAG_MAX_STRUCTURAL_ERRORS Then
        GanttDrag_StopRuntime True, False
    End If
    Resume SafeExit

End Sub

Private Sub GanttDrag_WatchTick()

    Dim ws As Worksheet
    Dim shapeName As Variant
    Dim shp As Shape
    Dim state As Variant
    Dim geometryChanged As Boolean

    If Not GanttDrag_IsGanttSheetActive(ws) Then
        GanttDrag_StopRuntime False, False
        Exit Sub
    End If

    If Not GanttDrag_IsSupportedTimelineScale() Then
        GanttDrag_StopRuntime False, False
        Exit Sub
    End If

    If IsPlanningWorkflowActive() Then Exit Sub
    If GetGanttInternalWrite() Then Exit Sub
    If Application.CalculationState <> xlDone Then Exit Sub

    If gShapeState Is Nothing Then GanttDrag_RebuildWatchMaps
    If gShapeState Is Nothing Then Exit Sub
    If gShapeState.Count = 0 Then
        GanttDrag_StopRuntime False, False
        Exit Sub
    End If

    For Each shapeName In gShapeState.Keys
        Set shp = Nothing
        On Error Resume Next
        Set shp = ws.Shapes(CStr(shapeName))
        On Error GoTo 0

        If shp Is Nothing Then
            GanttDrag_RebuildWatchMaps
            Exit Sub
        End If

        state = GanttDrag_GetShapeState(CStr(shapeName))
        geometryChanged = _
               Abs(CDbl(state(0)) - shp.Left) > GANTT_DRAG_CHANGE_TOLERANCE _
            Or Abs(CDbl(state(1)) - shp.Top) > GANTT_DRAG_CHANGE_TOLERANCE _
            Or Abs(CDbl(state(2)) - shp.Width) > GANTT_DRAG_CHANGE_TOLERANCE _
            Or Abs(CDbl(state(3)) - shp.Height) > GANTT_DRAG_CHANGE_TOLERANCE

        If geometryChanged Then
            gLastDebugStatus = _
                CStr(shapeName) & _
                " Left " & Format$(CDbl(state(0)), "0.00") & " -> " & Format$(shp.Left, "0.00") & _
                " | Top " & Format$(CDbl(state(1)), "0.00") & " -> " & Format$(shp.Top, "0.00") & _
                " | Width " & Format$(CDbl(state(2)), "0.00") & " -> " & Format$(shp.Width, "0.00") & _
                " | Height " & Format$(CDbl(state(3)), "0.00") & " -> " & Format$(shp.Height, "0.00")

            Debug.Print gLastDebugStatus
            GanttDrag_HandleShapeChange ws, shp, state
            Exit For
        End If
    Next shapeName

End Sub

Private Sub GanttDrag_HandleShapeChange( _
    ByVal ws As Worksheet, _
    ByVal shp As Shape, _
    ByVal oldState As Variant)

    Dim writtenCells As Collection
    Dim consoleMessages As Collection
    Dim revertMessages As Collection
    Dim testSucceeded As Boolean
    Dim revertSucceeded As Boolean
    Dim ganttRebuilt As Boolean
    Dim revertGanttRebuilt As Boolean
    Dim shouldRunTest As Boolean
    Dim displayFailure As Boolean
    Dim displayConsole As Boolean
    Dim dragInfo As Object
    Dim simulationMode As String

    On Error GoTo TransactionError

    If ws Is Nothing Then Exit Sub
    If shp Is Nothing Then Exit Sub
    If gTransactionActive Then Exit Sub

    Set writtenCells = New Collection
    shouldRunTest = GanttDrag_BuildTestInputs(ws, shp, oldState, writtenCells, dragInfo)

    If Not shouldRunTest Then
        GanttDrag_SaveShapeState shp, CLng(oldState(4)), CStr(oldState(5))
        Exit Sub
    End If

    gTransactionActive = True
    gTransactionCount = gTransactionCount + 1
    gLastTransactionResult = "RUNNING"
    gLastWrittenCells = GanttDrag_CellList(writtenCells)
    GanttDrag_Suspend

    simulationMode = GanttDrag_NormalizedSimulationMode(GanttLive_GetActiveSimulationMode())
    GanttDrag_SetDragInfoMode dragInfo, simulationMode

    If Not GanttDrag_IsSupportedSimulationMode(simulationMode) Then
        GanttDrag_ClearWrittenCells writtenCells
        If consoleMessages Is Nothing Then Set consoleMessages = New Collection
        GanttDrag_AddUnsupportedModeMessage consoleMessages, dragInfo, simulationMode
        gLastTransactionResult = "NO_ACTIVE_MODE"
        displayConsole = True
        GoTo CleanExit
    End If

    testSucceeded = GanttDrag_RunSimulationTransactionByMode(simulationMode, consoleMessages, ganttRebuilt)

    If testSucceeded Then
        gLastTransactionResult = "SUCCESS"
        GanttDrag_AddDragSuccessMessage consoleMessages, dragInfo
        displayConsole = True
    Else
        GanttDrag_ClearWrittenCells writtenCells
        revertSucceeded = GanttDrag_RunSimulationTransactionByMode(simulationMode, revertMessages, revertGanttRebuilt)

        If revertSucceeded Then
            gLastTransactionResult = "REVERTED"
        Else
            gLastTransactionResult = "REVERT_FAILED"
        End If

        GanttDrag_AddDragFailureMessage consoleMessages, dragInfo
        displayFailure = True
        displayConsole = True
    End If

CleanExit:
    On Error Resume Next
    GanttDrag_RebuildWatchMaps
    GanttDrag_Resume
    gTransactionActive = False

    If displayFailure Or displayConsole Then
        If Not consoleMessages Is Nothing Then CalcBridge_ShowPlanningConsole consoleMessages
    End If
    On Error GoTo 0
    Exit Sub

TransactionError:
    gLastTransactionResult = "ERROR"
    If Not writtenCells Is Nothing Then GanttDrag_ClearWrittenCells writtenCells
    If consoleMessages Is Nothing Then Set consoleMessages = New Collection
    GanttDrag_SetDragInfoMode dragInfo, simulationMode
    GanttDrag_AddDragFailureMessage consoleMessages, dragInfo
    displayFailure = True
    displayConsole = True
    Resume CleanExit

End Sub
Private Function GanttDrag_BuildTestInputs( _
    ByVal ws As Worksheet, _
    ByVal shp As Shape, _
    ByVal oldState As Variant, _
    ByVal writtenCells As Collection, _
    ByRef dragInfo As Object) As Boolean

    Dim oldLeft As Double
    Dim oldRight As Double
    Dim oldWidth As Double
    Dim newLeft As Double
    Dim newRight As Double
    Dim leftChanged As Boolean
    Dim rightChanged As Boolean
    Dim widthChanged As Boolean
    Dim testStart As Date
    Dim testFinish As Date
    Dim milestoneDate As Date
    Dim ganttRow As Long
    Dim taskType As String

    On Error GoTo SafeExit

    oldLeft = CDbl(oldState(0))
    oldWidth = CDbl(oldState(2))
    oldRight = oldLeft + oldWidth
    newLeft = CDbl(shp.Left)
    newRight = newLeft + CDbl(shp.Width)
    ganttRow = CLng(oldState(4))
    taskType = UCase$(Trim$(CStr(oldState(5))))

    leftChanged = Abs(newLeft - oldLeft) > GANTT_DRAG_CHANGE_TOLERANCE
    rightChanged = Abs(newRight - oldRight) > GANTT_DRAG_CHANGE_TOLERANCE
    widthChanged = Abs(CDbl(shp.Width) - oldWidth) > GANTT_DRAG_CHANGE_TOLERANCE

    If taskType = "MILESTONE" Then
        If Not leftChanged And Not rightChanged Then Exit Function
        If Not GanttDrag_DateFromX(ws, newLeft + (CDbl(shp.Width) / 2), 0, milestoneDate) Then Exit Function

        GanttDrag_WriteTestCell ws.Cells(ganttRow, GANTT_DRAG_COL_TEST_START), milestoneDate, writtenCells
        GanttDrag_WriteTestCell ws.Cells(ganttRow, GANTT_DRAG_COL_TEST_FINISH), milestoneDate, writtenCells
        Set dragInfo = GanttDrag_CreateDragInfo(ws, ganttRow, taskType, True, True, milestoneDate, milestoneDate)
        GanttDrag_BuildTestInputs = True
        Exit Function
    End If

    If taskType <> "TASK" Then Exit Function
    If Not leftChanged And Not rightChanged Then Exit Function

    If leftChanged And Not widthChanged Then
        If Not GanttDrag_DateFromX(ws, newLeft, -1, testStart) Then Exit Function
        If Not GanttDrag_DateFromX(ws, newRight, 1, testFinish) Then Exit Function

        GanttDrag_WriteTestCell ws.Cells(ganttRow, GANTT_DRAG_COL_TEST_START), testStart, writtenCells
        GanttDrag_WriteTestCell ws.Cells(ganttRow, GANTT_DRAG_COL_TEST_FINISH), testFinish, writtenCells
    ElseIf leftChanged And Not rightChanged Then
        If Not GanttDrag_DateFromX(ws, newLeft, -1, testStart) Then Exit Function
        GanttDrag_WriteTestCell ws.Cells(ganttRow, GANTT_DRAG_COL_TEST_START), testStart, writtenCells
    ElseIf Not leftChanged And rightChanged Then
        If Not GanttDrag_DateFromX(ws, newRight, 1, testFinish) Then Exit Function
        GanttDrag_WriteTestCell ws.Cells(ganttRow, GANTT_DRAG_COL_TEST_FINISH), testFinish, writtenCells
    Else
        If Not GanttDrag_DateFromX(ws, newLeft, -1, testStart) Then Exit Function
        If Not GanttDrag_DateFromX(ws, newRight, 1, testFinish) Then Exit Function

        GanttDrag_WriteTestCell ws.Cells(ganttRow, GANTT_DRAG_COL_TEST_START), testStart, writtenCells
        GanttDrag_WriteTestCell ws.Cells(ganttRow, GANTT_DRAG_COL_TEST_FINISH), testFinish, writtenCells
    End If

    If writtenCells.Count > 0 Then
        Set dragInfo = GanttDrag_CreateDragInfo(ws, ganttRow, taskType, _
            GanttDrag_WrittenCellsContainColumn(writtenCells, GANTT_DRAG_COL_TEST_START), _
            GanttDrag_WrittenCellsContainColumn(writtenCells, GANTT_DRAG_COL_TEST_FINISH), _
            testStart, testFinish)
    End If

    GanttDrag_BuildTestInputs = (writtenCells.Count > 0)

SafeExit:

End Function
Private Function GanttDrag_DateFromX( _
    ByVal ws As Worksheet, _
    ByVal xPos As Double, _
    ByVal anchorSide As Long, _
    ByRef resultDate As Date) As Boolean

    Select Case UCase$(Trim$(GetGanttTimelineScaleMode()))
        Case GANTT_DRAG_SCALE_WEEK
            GanttDrag_DateFromX = GanttDrag_WeekDateFromX(ws, xPos, anchorSide, resultDate)
        Case GANTT_DRAG_SCALE_MONTH
            GanttDrag_DateFromX = GanttDrag_MonthDateFromX(ws, xPos, anchorSide, resultDate)
        Case Else
            GanttDrag_DateFromX = GanttDrag_DayDateFromX(ws, xPos, anchorSide, resultDate)
    End Select

End Function

Private Function GanttDrag_DayDateFromX( _
    ByVal ws As Worksheet, _
    ByVal xPos As Double, _
    ByVal anchorSide As Long, _
    ByRef resultDate As Date) As Boolean

    Dim lastCol As Long
    Dim c As Long
    Dim headerValue As Variant
    Dim anchorX As Double
    Dim distance As Double
    Dim bestDistance As Double
    Dim foundDate As Boolean

    If ws Is Nothing Then Exit Function

    lastCol = ws.Cells(GANTT_DRAG_HEADER_ROW, ws.Columns.Count).End(xlToLeft).Column
    If lastCol < GANTT_DRAG_FIRST_TIMELINE_COL Then Exit Function

    bestDistance = 1E+30

    For c = GANTT_DRAG_FIRST_TIMELINE_COL To lastCol
        headerValue = ws.Cells(GANTT_DRAG_HEADER_ROW, c).value

        If IsDate(headerValue) Then
            Select Case anchorSide
                Case -1
                    anchorX = ws.Cells(GANTT_DRAG_HEADER_ROW, c).Left
                Case 1
                    anchorX = ws.Cells(GANTT_DRAG_HEADER_ROW, c).Left + _
                        ws.Cells(GANTT_DRAG_HEADER_ROW, c).Width
                Case Else
                    anchorX = ws.Cells(GANTT_DRAG_HEADER_ROW, c).Left + _
                        (ws.Cells(GANTT_DRAG_HEADER_ROW, c).Width / 2)
            End Select

            distance = Abs(xPos - anchorX)
            If distance < bestDistance Then
                bestDistance = distance
                resultDate = DateValue(CDate(headerValue))
                foundDate = True
            End If
        End If
    Next c

    GanttDrag_DayDateFromX = foundDate

End Function

Private Function GanttDrag_WeekDateFromX( _
    ByVal ws As Worksheet, _
    ByVal xPos As Double, _
    ByVal anchorSide As Long, _
    ByRef resultDate As Date) As Boolean

    Dim lastCol As Long
    Dim c As Long
    Dim targetCol As Long
    Dim cellLeft As Double
    Dim cellWidth As Double
    Dim cellRight As Double
    Dim distance As Double
    Dim bestDistance As Double
    Dim fraction As Double
    Dim dayOffset As Long
    Dim weekStart As Date

    If ws Is Nothing Then Exit Function

    lastCol = ws.Cells(GANTT_DRAG_HEADER_ROW, ws.Columns.Count).End(xlToLeft).Column
    If lastCol < GANTT_DRAG_FIRST_TIMELINE_COL Then Exit Function

    bestDistance = 1E+30

    For c = GANTT_DRAG_FIRST_TIMELINE_COL To lastCol
        cellLeft = ws.Cells(GANTT_DRAG_HEADER_ROW, c).Left
        cellWidth = ws.Cells(GANTT_DRAG_HEADER_ROW, c).Width
        cellRight = cellLeft + cellWidth

        If (anchorSide = 1 And xPos >= cellLeft And xPos <= cellRight) Or _
           (anchorSide <> 1 And xPos >= cellLeft And (xPos < cellRight Or c = lastCol)) Then
            targetCol = c
            Exit For
        End If

        distance = WorksheetFunction.Min(Abs(xPos - cellLeft), Abs(xPos - cellRight))
        If distance < bestDistance Then
            bestDistance = distance
            targetCol = c
        End If
    Next c

    If targetCol < GANTT_DRAG_FIRST_TIMELINE_COL Then Exit Function
    If Not GanttDrag_TimelineWeekStart(ws, targetCol, weekStart) Then Exit Function

    cellLeft = ws.Cells(GANTT_DRAG_HEADER_ROW, targetCol).Left
    cellWidth = ws.Cells(GANTT_DRAG_HEADER_ROW, targetCol).Width
    If cellWidth <= 0 Then Exit Function

    fraction = (xPos - cellLeft) / cellWidth
    If fraction < 0# Then fraction = 0#
    If fraction > 1# Then fraction = 1#

    Select Case anchorSide
        Case 1
            dayOffset = GanttDrag_CeilPositive(fraction * 7#) - 1
        Case Else
            dayOffset = CLng(Int(fraction * 7#))
    End Select

    If dayOffset < 0 Then dayOffset = 0
    If dayOffset > 6 Then dayOffset = 6

    resultDate = DateAdd("d", dayOffset, weekStart)
    GanttDrag_WeekDateFromX = True

End Function

Private Function GanttDrag_TimelineWeekStart( _
    ByVal ws As Worksheet, _
    ByVal timelineCol As Long, _
    ByRef weekStart As Date) As Boolean

    Dim weekNum As Long
    Dim isoYear As Long
    Dim yearValue As Variant
    Dim labelValue As String
    Dim janFourth As Date

    If ws Is Nothing Then Exit Function
    If timelineCol < GANTT_DRAG_FIRST_TIMELINE_COL Then Exit Function

    labelValue = CStr(ws.Cells(GANTT_DRAG_HEADER_ROW, timelineCol).value)
    weekNum = GanttDrag_ExtractFirstLong(labelValue)
    If weekNum < 1 Or weekNum > 53 Then Exit Function

    yearValue = ws.Cells(GANTT_DRAG_HEADER_ROW - 1, timelineCol).MergeArea.Cells(1, 1).value
    If Not IsNumeric(yearValue) Then Exit Function
    isoYear = CLng(yearValue)
    If isoYear < 1900 Then Exit Function

    janFourth = DateSerial(isoYear, 1, 4)
    weekStart = DateAdd("d", (weekNum - 1) * 7, janFourth - Weekday(janFourth, vbMonday) + 1)
    GanttDrag_TimelineWeekStart = True

End Function

Private Function GanttDrag_ExtractFirstLong(ByVal textValue As String) As Long

    Dim i As Long
    Dim ch As String
    Dim digits As String

    For i = 1 To Len(textValue)
        ch = Mid$(textValue, i, 1)
        If ch >= "0" And ch <= "9" Then
            digits = digits & ch
        ElseIf digits <> "" Then
            Exit For
        End If
    Next i

    If digits <> "" Then GanttDrag_ExtractFirstLong = CLng(digits)

End Function

Private Function GanttDrag_CeilPositive(ByVal value As Double) As Long

    If value <= 0# Then
        GanttDrag_CeilPositive = 0
    Else
        GanttDrag_CeilPositive = CLng(-Int(-value))
    End If

End Function

Private Function GanttDrag_MonthDateFromX( _
    ByVal ws As Worksheet, _
    ByVal xPos As Double, _
    ByVal anchorSide As Long, _
    ByRef resultDate As Date) As Boolean

    Dim targetCol As Long
    Dim cellLeft As Double
    Dim cellWidth As Double
    Dim fraction As Double
    Dim monthStart As Date
    Dim daysInMonth As Long
    Dim dayOffset As Long

    If ws Is Nothing Then Exit Function
    If Not GanttDrag_TimelineColumnFromX(ws, xPos, anchorSide, targetCol) Then Exit Function
    If Not GanttDrag_TimelineMonthStart(ws, targetCol, monthStart) Then Exit Function

    cellLeft = ws.Cells(GANTT_DRAG_HEADER_ROW, targetCol).Left
    cellWidth = ws.Cells(GANTT_DRAG_HEADER_ROW, targetCol).Width
    If cellWidth <= 0 Then Exit Function

    fraction = (xPos - cellLeft) / cellWidth
    If fraction < 0# Then fraction = 0#
    If fraction > 1# Then fraction = 1#

    daysInMonth = Day(DateSerial(Year(monthStart), Month(monthStart) + 1, 0))

    Select Case anchorSide
        Case 1
            dayOffset = GanttDrag_CeilPositive(fraction * CDbl(daysInMonth)) - 1
        Case Else
            dayOffset = CLng(Int(fraction * CDbl(daysInMonth)))
    End Select

    If dayOffset < 0 Then dayOffset = 0
    If dayOffset > daysInMonth - 1 Then dayOffset = daysInMonth - 1

    resultDate = DateAdd("d", dayOffset, monthStart)
    GanttDrag_MonthDateFromX = True

End Function

Private Function GanttDrag_TimelineColumnFromX( _
    ByVal ws As Worksheet, _
    ByVal xPos As Double, _
    ByVal anchorSide As Long, _
    ByRef targetCol As Long) As Boolean

    Dim lastCol As Long
    Dim c As Long
    Dim cellLeft As Double
    Dim cellWidth As Double
    Dim cellRight As Double
    Dim distance As Double
    Dim bestDistance As Double

    If ws Is Nothing Then Exit Function

    lastCol = ws.Cells(GANTT_DRAG_HEADER_ROW, ws.Columns.Count).End(xlToLeft).Column
    If lastCol < GANTT_DRAG_FIRST_TIMELINE_COL Then Exit Function

    bestDistance = 1E+30

    For c = GANTT_DRAG_FIRST_TIMELINE_COL To lastCol
        cellLeft = ws.Cells(GANTT_DRAG_HEADER_ROW, c).Left
        cellWidth = ws.Cells(GANTT_DRAG_HEADER_ROW, c).Width
        cellRight = cellLeft + cellWidth

        If (anchorSide = 1 And xPos >= cellLeft And xPos <= cellRight) Or _
           (anchorSide <> 1 And xPos >= cellLeft And (xPos < cellRight Or c = lastCol)) Then
            targetCol = c
            GanttDrag_TimelineColumnFromX = True
            Exit Function
        End If

        distance = Abs(xPos - cellLeft)
        If Abs(xPos - cellRight) < distance Then distance = Abs(xPos - cellRight)

        If distance < bestDistance Then
            bestDistance = distance
            targetCol = c
        End If
    Next c

    GanttDrag_TimelineColumnFromX = (targetCol >= GANTT_DRAG_FIRST_TIMELINE_COL)

End Function

Private Function GanttDrag_TimelineMonthStart( _
    ByVal ws As Worksheet, _
    ByVal timelineCol As Long, _
    ByRef monthStart As Date) As Boolean

    Dim yearValue As Variant
    Dim yearNum As Long
    Dim monthNum As Long

    If ws Is Nothing Then Exit Function
    If timelineCol < GANTT_DRAG_FIRST_TIMELINE_COL Then Exit Function

    yearValue = ws.Cells(GANTT_DRAG_HEADER_ROW - 1, timelineCol).MergeArea.Cells(1, 1).value
    If Not IsNumeric(yearValue) Then Exit Function
    yearNum = CLng(yearValue)
    If yearNum < 1900 Then Exit Function

    monthNum = GanttDrag_MonthNumberFromLabel(CStr(ws.Cells(GANTT_DRAG_HEADER_ROW, timelineCol).value))
    If monthNum < 1 Or monthNum > 12 Then Exit Function

    monthStart = DateSerial(yearNum, monthNum, 1)
    GanttDrag_TimelineMonthStart = True

End Function

Private Function GanttDrag_MonthNumberFromLabel(ByVal monthLabel As String) As Long

    Dim normalizedLabel As String

    normalizedLabel = LCase$(Trim$(monthLabel))

    Select Case normalizedLabel
        Case "jan", "janv", "jan.", "janv.": GanttDrag_MonthNumberFromLabel = 1
        Case "feb", "fév", "fev", "févr", "fevr", "fév.", "fev.", "févr.", "fevr.": GanttDrag_MonthNumberFromLabel = 2
        Case "mar", "mars": GanttDrag_MonthNumberFromLabel = 3
        Case "apr", "avr", "avr.": GanttDrag_MonthNumberFromLabel = 4
        Case "may", "mai": GanttDrag_MonthNumberFromLabel = 5
        Case "jun", "juin": GanttDrag_MonthNumberFromLabel = 6
        Case "jul", "juil", "juil.": GanttDrag_MonthNumberFromLabel = 7
        Case "aug", "août", "aout": GanttDrag_MonthNumberFromLabel = 8
        Case "sep", "sept", "sept.": GanttDrag_MonthNumberFromLabel = 9
        Case "oct", "oct.": GanttDrag_MonthNumberFromLabel = 10
        Case "nov", "nov.": GanttDrag_MonthNumberFromLabel = 11
        Case "dec", "déc", "dec.", "déc.": GanttDrag_MonthNumberFromLabel = 12
    End Select

End Function
Private Function GanttDrag_CreateDragInfo( _
    ByVal ws As Worksheet, _
    ByVal ganttRow As Long, _
    ByVal taskType As String, _
    ByVal changedStart As Boolean, _
    ByVal changedFinish As Boolean, _
    ByVal requestedStart As Variant, _
    ByVal requestedFinish As Variant) As Object

    Dim info As Object

    Set info = CreateObject("Scripting.Dictionary")

    info("TaskName") = Trim$(CStr(ws.Cells(ganttRow, 2).value))
    info("WBS") = Trim$(CStr(ws.Cells(ganttRow, 1).value))
    info("TaskType") = UCase$(Trim$(taskType))
    info("ChangedStart") = changedStart
    info("ChangedFinish") = changedFinish
    info("RequestedStart") = requestedStart
    info("RequestedFinish") = requestedFinish

    Set GanttDrag_CreateDragInfo = info

End Function

Private Function GanttDrag_WrittenCellsContainColumn( _
    ByVal writtenCells As Collection, _
    ByVal columnIndex As Long) As Boolean

    Dim item As Variant

    If writtenCells Is Nothing Then Exit Function

    For Each item In writtenCells
        If CLng(item.Column) = columnIndex Then
            GanttDrag_WrittenCellsContainColumn = True
            Exit Function
        End If
    Next item

End Function

Private Function GanttDrag_NormalizedSimulationMode(ByVal simulationMode As String) As String

    GanttDrag_NormalizedSimulationMode = UCase$(Trim$(simulationMode))
    If GanttDrag_NormalizedSimulationMode = "" Then GanttDrag_NormalizedSimulationMode = "TEST"

End Function

Private Function GanttDrag_IsSupportedSimulationMode(ByVal simulationMode As String) As Boolean

    Select Case UCase$(Trim$(simulationMode))
        Case "TEST", "SCENARIO"
            GanttDrag_IsSupportedSimulationMode = True
    End Select

End Function

Private Function GanttDrag_RunSimulationTransactionByMode( _
    ByVal simulationMode As String, _
    ByRef consoleMessages As Collection, _
    ByRef ganttRebuilt As Boolean) As Boolean

    Select Case UCase$(Trim$(simulationMode))
        Case "TEST"
            GanttDrag_RunSimulationTransactionByMode = GanttLive_RunTestTransaction(consoleMessages, ganttRebuilt)
        Case "SCENARIO"
            GanttDrag_RunSimulationTransactionByMode = GanttLive_RunScenarioTransaction(consoleMessages, ganttRebuilt)
    End Select

End Function

Private Sub GanttDrag_SetDragInfoMode( _
    ByVal dragInfo As Object, _
    ByVal simulationMode As String)

    If dragInfo Is Nothing Then Exit Sub
    dragInfo("EngineMode") = UCase$(Trim$(simulationMode))

End Sub

Private Sub GanttDrag_AddUnsupportedModeMessage( _
    ByVal consoleMessages As Collection, _
    ByVal dragInfo As Object, _
    ByVal simulationMode As String)

    Dim taskLabel As String
    Dim modeLabel As String

    If consoleMessages Is Nothing Then Exit Sub

    taskLabel = GanttDrag_InfoTaskLabel(dragInfo)
    modeLabel = Trim$(simulationMode)
    If modeLabel = "" Then modeLabel = "NONE"

    CalcBridge_AddConsoleMessage consoleMessages, "WARNING", BiMsg( _
        "Drag ignoré : aucun moteur de simulation actif compatible n'a été trouvé." & vbCrLf & _
        "Tâche : " & taskLabel & vbCrLf & _
        "Mode actif : " & modeLabel & vbCrLf & _
        "Activez TEST ou SCENARIO avant de déplacer une tâche.", _
        "Drag ignored: no compatible active simulation engine was found." & vbCrLf & _
        "Task: " & taskLabel & vbCrLf & _
        "Active mode: " & modeLabel & vbCrLf & _
        "Activate TEST or SCENARIO before moving a task.")

End Sub

Private Sub GanttDrag_AddDragSuccessMessage( _
    ByVal consoleMessages As Collection, _
    ByVal dragInfo As Object)

    If consoleMessages Is Nothing Then Exit Sub
    CalcBridge_AddConsoleMessage consoleMessages, "INFO", GanttDrag_BuildDragMessage(dragInfo, True)

End Sub

Private Sub GanttDrag_AddDragFailureMessage( _
    ByRef consoleMessages As Collection, _
    ByVal dragInfo As Object)

    If consoleMessages Is Nothing Then Set consoleMessages = New Collection
    CalcBridge_AddConsoleMessage consoleMessages, "STOP", GanttDrag_BuildDragMessage(dragInfo, False)

End Sub

Private Function GanttDrag_BuildDragMessage( _
    ByVal dragInfo As Object, _
    ByVal success As Boolean) As String

    Dim taskLabel As String
    Dim changesFr As String
    Dim changesEn As String
    Dim frText As String
    Dim enText As String

    taskLabel = GanttDrag_InfoTaskLabel(dragInfo)
    changesFr = GanttDrag_InfoChangesText(dragInfo, True)
    changesEn = GanttDrag_InfoChangesText(dragInfo, False)

    If success Then
        frText = "Modification " & GanttDrag_InfoEngineLabel(dragInfo, "Drag") & " appliquée." & vbCrLf & _
            "Tâche : " & taskLabel & vbCrLf & _
            "Modification demandée : " & changesFr & vbCrLf & _
            "Le planning a été recalculé. Les conséquences éventuelles sur les autres tâches proviennent du moteur planning." & vbCrLf & vbCrLf & _
            "Pour annuler cette modification, videz simplement les cellules TEST jaunes puis relancez TEST."

        enText = GanttDrag_InfoEngineLabel(dragInfo, "Drag") & " modification applied." & vbCrLf & _
            "Task: " & taskLabel & vbCrLf & _
            "Requested modification: " & changesEn & vbCrLf & _
            "The schedule has been recalculated. Any consequences on other tasks come from the planning engine." & vbCrLf & vbCrLf & _
            "To cancel this modification, simply clear the yellow TEST cells and run TEST again."
    Else
        frText = "La modification demandée par Drag n'a pas pu être appliquée." & vbCrLf & _
            "Tâche : " & taskLabel & vbCrLf & _
            "Modification demandée : " & changesFr

        enText = "The modification requested by Drag could not be applied." & vbCrLf & _
            "Task: " & taskLabel & vbCrLf & _
            "Requested modification: " & changesEn
    End If

    GanttDrag_BuildDragMessage = BiMsg(frText, enText)

End Function

Private Function GanttDrag_InfoEngineLabel( _
    ByVal dragInfo As Object, _
    ByVal defaultLabel As String) As String

    Dim modeLabel As String

    If dragInfo Is Nothing Then
        GanttDrag_InfoEngineLabel = defaultLabel
        Exit Function
    End If

    If dragInfo.Exists("EngineMode") Then modeLabel = UCase$(Trim$(CStr(dragInfo("EngineMode"))))

    Select Case modeLabel
        Case "TEST"
            GanttDrag_InfoEngineLabel = "Drag/Test"
        Case "SCENARIO"
            GanttDrag_InfoEngineLabel = "Drag/Scenario"
        Case Else
            GanttDrag_InfoEngineLabel = defaultLabel
    End Select

End Function

Private Function GanttDrag_InfoTaskLabel(ByVal dragInfo As Object) As String

    Dim taskName As String
    Dim wbsVal As String

    If dragInfo Is Nothing Then
        GanttDrag_InfoTaskLabel = "(unknown task)"
        Exit Function
    End If

    taskName = Trim$(CStr(dragInfo("TaskName")))
    wbsVal = Trim$(CStr(dragInfo("WBS")))

    If wbsVal <> "" And taskName <> "" Then
        GanttDrag_InfoTaskLabel = wbsVal & " - " & taskName
    ElseIf taskName <> "" Then
        GanttDrag_InfoTaskLabel = taskName
    ElseIf wbsVal <> "" Then
        GanttDrag_InfoTaskLabel = wbsVal
    Else
        GanttDrag_InfoTaskLabel = "(unknown task)"
    End If

End Function

Private Function GanttDrag_InfoChangesText( _
    ByVal dragInfo As Object, _
    ByVal french As Boolean) As String

    Dim changedStart As Boolean
    Dim changedFinish As Boolean
    Dim startText As String
    Dim finishText As String

    If dragInfo Is Nothing Then
        If french Then
            GanttDrag_InfoChangesText = "modification non identifiée"
        Else
            GanttDrag_InfoChangesText = "unidentified modification"
        End If
        Exit Function
    End If

    changedStart = CBool(dragInfo("ChangedStart"))
    changedFinish = CBool(dragInfo("ChangedFinish"))
    startText = GanttDrag_FormatDateValue(dragInfo("RequestedStart"))
    finishText = GanttDrag_FormatDateValue(dragInfo("RequestedFinish"))

    If changedStart And changedFinish Then
        If startText = finishText Then
            If french Then
                GanttDrag_InfoChangesText = "début et fin = " & startText
            Else
                GanttDrag_InfoChangesText = "start and finish = " & startText
            End If
        ElseIf french Then
            GanttDrag_InfoChangesText = "début = " & startText & ", fin = " & finishText
        Else
            GanttDrag_InfoChangesText = "start = " & startText & ", finish = " & finishText
        End If
    ElseIf changedStart Then
        If french Then
            GanttDrag_InfoChangesText = "début = " & startText
        Else
            GanttDrag_InfoChangesText = "start = " & startText
        End If
    ElseIf changedFinish Then
        If french Then
            GanttDrag_InfoChangesText = "fin = " & finishText
        Else
            GanttDrag_InfoChangesText = "finish = " & finishText
        End If
    ElseIf french Then
        GanttDrag_InfoChangesText = "aucune date modifiée"
    Else
        GanttDrag_InfoChangesText = "no date changed"
    End If

End Function

Private Function GanttDrag_FormatDateValue(ByVal value As Variant) As String

    If IsDate(value) Then
        GanttDrag_FormatDateValue = Format$(CDate(value), "dd/mm/yyyy")
    ElseIf Trim$(CStr(value)) <> "" Then
        GanttDrag_FormatDateValue = Trim$(CStr(value))
    Else
        GanttDrag_FormatDateValue = "-"
    End If

End Function

Private Sub GanttDrag_WriteTestCell( _
    ByVal targetCell As Range, _
    ByVal testDate As Date, _
    ByVal writtenCells As Collection)

    Dim oldEnableEvents As Boolean
    Dim oldInternalWrite As Boolean

    If targetCell Is Nothing Then Exit Sub
    If writtenCells Is Nothing Then Exit Sub

    oldEnableEvents = Application.EnableEvents
    oldInternalWrite = GetGanttInternalWrite()

    On Error GoTo CleanExit
    Application.EnableEvents = False
    SetGanttInternalWrite True

    targetCell.value = DateValue(testDate)
    writtenCells.Add targetCell

CleanExit:
    SetGanttInternalWrite oldInternalWrite
    Application.EnableEvents = oldEnableEvents

End Sub

Private Sub GanttDrag_ClearWrittenCells(ByVal writtenCells As Collection)

    Dim oldEnableEvents As Boolean
    Dim oldInternalWrite As Boolean
    Dim item As Variant

    If writtenCells Is Nothing Then Exit Sub

    oldEnableEvents = Application.EnableEvents
    oldInternalWrite = GetGanttInternalWrite()

    On Error GoTo CleanExit
    Application.EnableEvents = False
    SetGanttInternalWrite True

    For Each item In writtenCells
        item.ClearContents
    Next item

CleanExit:
    SetGanttInternalWrite oldInternalWrite
    Application.EnableEvents = oldEnableEvents

End Sub

Private Function GanttDrag_CellList(ByVal cells As Collection) As String

    Dim item As Variant
    Dim result As String

    If cells Is Nothing Then Exit Function

    For Each item In cells
        If result <> "" Then result = result & ","
        result = result & item.Address(False, False)
    Next item

    GanttDrag_CellList = result

End Function

Private Function GanttDrag_IsEligibleShape( _
    ByVal ws As Worksheet, _
    ByVal shp As Shape, _
    ByRef rowIndex As Long, _
    ByRef taskType As String) As Boolean

    Dim tblWBS As ListObject
    Dim shapeName As String
    Dim suffix As String
    Dim dataRow As Long
    Dim normalizedTaskType As String
    Dim isMilestoneShape As Boolean

    If ws Is Nothing Then Exit Function
    If shp Is Nothing Then Exit Function
    If shp.Visible = msoFalse Then Exit Function

    shapeName = CStr(shp.Name)

    If Left$(shapeName, 5) = "TASK_" Then
        suffix = Mid$(shapeName, 6)
        If suffix = "" Or suffix Like "*[!0-9]*" Then Exit Function
    ElseIf Left$(shapeName, 3) = "MS_" Then
        suffix = Mid$(shapeName, 4)
        If suffix = "" Or suffix Like "*[!0-9]*" Then Exit Function
        isMilestoneShape = True
    Else
        Exit Function
    End If

    dataRow = CLng(Val(suffix))
    If dataRow < 1 Then Exit Function

    Set tblWBS = ThisWorkbook.Worksheets("WBS").ListObjects("tbl_WBS")
    If tblWBS.DataBodyRange Is Nothing Then Exit Function
    If dataRow > tblWBS.ListRows.Count Then Exit Function

    normalizedTaskType = UCase$(Trim$(CStr( _
        tblWBS.DataBodyRange.Cells(dataRow, tblWBS.ListColumns("Task Type").Index).value)))

    If Not isMilestoneShape Then
        If normalizedTaskType <> "TASK" Then Exit Function
        taskType = "TASK"
    Else
        'MS_n is created only by DrawMilestone. The rendered shape type is the
        'authoritative watcher contract, independent of localized/input text.
        taskType = "MILESTONE"
    End If

    rowIndex = GANTT_DRAG_FIRST_TASK_ROW + dataRow - 1
    GanttDrag_IsEligibleShape = True

End Function

Private Sub GanttDrag_SaveShapeState( _
    ByVal shp As Shape, _
    ByVal rowIndex As Long, _
    ByVal taskType As String)

    Dim state(0 To 5) As Variant

    If shp Is Nothing Then Exit Sub
    If gShapeState Is Nothing Then Set gShapeState = CreateObject("Scripting.Dictionary")

    state(0) = CDbl(shp.Left)
    state(1) = CDbl(shp.Top)
    state(2) = CDbl(shp.Width)
    state(3) = CDbl(shp.Height)
    state(4) = CLng(rowIndex)
    state(5) = CStr(taskType)

    gShapeState(CStr(shp.Name)) = state

End Sub

Private Function GanttDrag_GetShapeState(ByVal shapeName As String) As Variant

    If gShapeState Is Nothing Then Exit Function
    If Not gShapeState.Exists(shapeName) Then Exit Function

    GanttDrag_GetShapeState = gShapeState(shapeName)

End Function

Private Sub GanttDrag_Suspend()

    gSuspendDepth = gSuspendDepth + 1

End Sub

Private Sub GanttDrag_Resume()

    If gSuspendDepth > 0 Then gSuspendDepth = gSuspendDepth - 1
    If gSuspendDepth = 0 And GanttDrag_IsWatching() Then GanttDrag_RebuildWatchMaps

End Sub

Private Function GanttDrag_IsSupportedTimelineScale() As Boolean

    Select Case UCase$(Trim$(GetGanttTimelineScaleMode()))
        Case GANTT_DRAG_SCALE_DAY, GANTT_DRAG_SCALE_WEEK, GANTT_DRAG_SCALE_MONTH
            GanttDrag_IsSupportedTimelineScale = True
    End Select

End Function
Private Function GanttDrag_CanStartWatch() As Boolean

    Dim ws As Worksheet

    If Not GanttDrag_IsGanttSheetActive(ws) Then Exit Function
    If Not GanttDrag_IsSupportedTimelineScale() Then Exit Function
    If IsPlanningWorkflowActive() Then Exit Function
    If GetGanttInternalWrite() Then Exit Function
    If Application.CalculationState <> xlDone Then Exit Function

    GanttDrag_CanStartWatch = True

End Function

Private Function GanttDrag_IsGanttSheetActive(ByRef ws As Worksheet) As Boolean

    On Error GoTo SafeExit

    If Application.ActiveWorkbook Is Nothing Then Exit Function
    If Not (Application.ActiveWorkbook Is ThisWorkbook) Then Exit Function
    If Application.ActiveSheet Is Nothing Then Exit Function
    If Not (Application.ActiveSheet.Parent Is ThisWorkbook) Then Exit Function
    If CStr(Application.ActiveSheet.Name) <> GANTT_DRAG_SHEET Then Exit Function

    Set ws = Application.ActiveSheet
    GanttDrag_IsGanttSheetActive = True

SafeExit:

End Function
