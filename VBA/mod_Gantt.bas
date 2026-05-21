Attribute VB_Name = "mod_Gantt"
'=====================================================
' mod_Gantt
'=====================================================
' Rôle :
' - rendu visuel uniquement
' - aucune logique métier
' - lit base/test et projette le planning dans la feuille GANTT
'
' Points d'attention :
' - toujours nettoyer les shapes avant redraw
' - DisplayOnly et Full Refresh doivent rester cohérents
' - ne jamais réintroduire de logique CALC ici
' - placement des shapes : xlMoveAndSize
'=====================================================

Option Explicit
Private gExpandedLinks As Object

Private Const GANTT_SHEET As String = "GANTT"
Private Const WBS_SHEET As String = "WBS"
Private Const WBS_TABLE As String = "tbl_WBS"
Private Const CALC_SHEET As String = "CALC"
Private Const CALC_TABLE As String = "tbl_CALC"

Private Const FIRST_TIMELINE_COL As Long = 11   ' K
Private Const TITLE_ROW As Long = 1
Private Const TOGGLE_ROW_TOP As Long = 2
Private Const TOGGLE_ROW_BOTTOM As Long = 3
Private Const HEADER_ROW_1 As Long = 3
Private Const HEADER_ROW_2 As Long = 4
Private Const FIRST_TASK_ROW As Long = 5

Private Const COL_WBS As Long = 1
Private Const COL_TASK As Long = 2
Private Const COL_START As Long = 3
Private Const COL_FINISH As Long = 4
Private Const COL_TEST_START As Long = 5
Private Const COL_TEST_FINISH As Long = 6
Private Const COL_DURATION As Long = 7
Private Const COL_PROGRESS As Long = 8
Private Const COL_TEST_PROGRESS As Long = 9
Private Const COL_LOGIC As Long = 10

Private Const BTN_VIEW_BG_NAME As String = "btn_Gantt_View_BG"
Private Const BTN_VIEW_KNOB_NAME As String = "btn_Gantt_View_Knob"
Private Const BTN_VIEW_LEFT_NAME As String = "btn_Gantt_View_Left"
Private Const BTN_VIEW_RIGHT_NAME As String = "btn_Gantt_View_Right"
Private Const BTN_VIEW_CELL_LEFT As String = "A2"
Private Const BTN_VIEW_CELL_RIGHT As String = "B2"

Private Const BTN_CP_BG_NAME As String = "btn_Gantt_CP_BG"
Private Const BTN_CP_KNOB_NAME As String = "btn_Gantt_CP_Knob"
Private Const BTN_CP_LEFT_NAME As String = "btn_Gantt_CP_Left"
Private Const BTN_CP_RIGHT_NAME As String = "btn_Gantt_CP_Right"
Private Const BTN_CP_CELL_LEFT As String = "A3"
Private Const BTN_CP_CELL_RIGHT As String = "B3"

Private Const GANTT_SCALE_DAY As String = "DAY"
Private Const GANTT_SCALE_WEEK As String = "WEEK"

Private Const BTN_CP_MULTI_BG_NAME As String = "shp_GANTT_CPMode_BG"
Private Const BTN_CP_MULTI_KNOB_NAME As String = "shp_GANTT_CPMode_Knob"
Private Const BTN_CP_MULTI_LEFT_NAME As String = "lbl_GANTT_CPMode_Toggle"

Private Const BTN_CONSTRAINT_BG_NAME As String = "btn_Gantt_Constraint_BG"
Private Const BTN_CONSTRAINT_KNOB_NAME As String = "btn_Gantt_Constraint_Knob"
Private Const BTN_CONSTRAINT_LEFT_NAME As String = "btn_Gantt_Constraint_Left"

Private Const GANTT_ROW_HEIGHT_HEADER_1 As Double = 21
Private Const GANTT_ROW_HEIGHT_HEADER_2 As Double = 18
Private Const GANTT_ROW_HEIGHT_TASK As Double = 18

Private Const BTN_SCALE_BG_NAME As String = "btn_Gantt_Scale_BG"
Private Const BTN_SCALE_KNOB_NAME As String = "btn_Gantt_Scale_Knob"
Private Const BTN_SCALE_LEFT_NAME As String = "btn_Gantt_Scale_Left"
Private Const BTN_SCALE_RIGHT_NAME As String = "btn_Gantt_Scale_Right"
Private Const BTN_SCALE_CELL_LEFT As String = "D2"
Private Const BTN_SCALE_CELL_RIGHT As String = "E2"

Private Const BTN_SCENARIO_NAME As String = "btn_Gantt_Scenario"
Private Const BTN_SCENARIO_CAPTION As String = "Scenario"

Private Const COLOR_TASK_BLUE As Long = 12874308
Private Const COLOR_TASK_CRITICAL As Long = 192

Private Const COLOR_PROGRESS_GREEN As Long = 4699504
Private Const COLOR_PROGRESS_ORANGE As Long = 3248093

Private Const LINK_STUB As Double = 6
Private Const LINK_MIN_CHANNEL_GAP As Double = 4
Private Const LINK_EDGE_PADDING As Double = 8

Private Const TEST_START_HEADER As String = "Test Start"
Private Const TEST_FINISH_HEADER As String = "Test Finish"
Private Const TEST_PROGRESS_HEADER As String = "Test %"

Private Const BTN_TEST_NAME As String = "btn_Gantt_Test"
Private Const BTN_LOCK_NAME As String = "btn_Gantt_Lock"

Private Const BTN_TEST_CAPTION As String = "Test"
Private Const BTN_LOCK_CAPTION As String = "Lock"

Private Const GANTT_VIEW_DETAIL As String = "DETAIL"
Private Const GANTT_VIEW_SUMMARY As String = "SUMMARY"

Private Const GANTT_ANALYTICS_PATH_NONE As String = "NONE"
Private Const GANTT_ANALYTICS_PATH_CP As String = "CP"
Private Const GANTT_ANALYTICS_PATH_LP As String = "LP"

Private Const LINK_ANCHOR_START As String = "START"
Private Const LINK_ANCHOR_FINISH As String = "FINISH"

Private gGanttViewMode As String
Private gPreserveTestInputs As Boolean
Private gGanttInternalWrite As Boolean
Private gAnalyticsPathMode As String
Private gShowConstraints As Boolean
Private gShowConstraintsInitialized As Boolean
Private gTimelineScaleMode As String
Private gGanttUiStateBootstrapped As Boolean


Public Sub SetGanttInternalWrite(ByVal internalWrite As Boolean)
    gGanttInternalWrite = internalWrite
End Sub

Public Function GetGanttInternalWrite() As Boolean
    GetGanttInternalWrite = gGanttInternalWrite
End Function

Public Sub SetGanttPreserveTestInputs(ByVal preserveValue As Boolean)
    gPreserveTestInputs = preserveValue
End Sub

Public Function GetGanttPreserveTestInputs() As Boolean
    GetGanttPreserveTestInputs = gPreserveTestInputs
End Function

Private Sub EnsureGanttViewInitialized()

    BootstrapGanttUiStateFromSheet

    If Trim$(gGanttViewMode) = "" Then
        gGanttViewMode = GANTT_VIEW_DETAIL
    End If

    If Trim$(gTimelineScaleMode) = "" Then
        gTimelineScaleMode = GANTT_SCALE_DAY
    End If

    If Trim$(gAnalyticsPathMode) = "" Then
        gAnalyticsPathMode = GANTT_ANALYTICS_PATH_NONE
    End If

    If Not gShowConstraintsInitialized Then
        gShowConstraints = True
        gShowConstraintsInitialized = True
    End If

End Sub

Private Sub BootstrapGanttUiStateFromSheet()

    Dim ws As Worksheet
    Dim isOn As Boolean
    Dim pathMode As String

    If gGanttUiStateBootstrapped Then Exit Sub
    gGanttUiStateBootstrapped = True

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(GANTT_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    If TryReadGanttTwoStateToggle(ws, BTN_VIEW_BG_NAME, BTN_VIEW_KNOB_NAME, isOn) Then
        If isOn Then
            gGanttViewMode = GANTT_VIEW_SUMMARY
        Else
            gGanttViewMode = GANTT_VIEW_DETAIL
        End If
    End If

    If TryReadGanttTwoStateToggle(ws, BTN_SCALE_BG_NAME, BTN_SCALE_KNOB_NAME, isOn) Then
        If isOn Then
            gTimelineScaleMode = GANTT_SCALE_WEEK
        Else
            gTimelineScaleMode = GANTT_SCALE_DAY
        End If
    End If

    If TryReadGanttTwoStateToggle(ws, BTN_CONSTRAINT_BG_NAME, BTN_CONSTRAINT_KNOB_NAME, isOn) Then
        gShowConstraints = isOn
        gShowConstraintsInitialized = True
    End If

    If TryReadGanttAnalyticsPathToggle(ws, pathMode) Then
        gAnalyticsPathMode = pathMode
    End If

End Sub

Private Function TryReadGanttTwoStateToggle( _
    ByVal ws As Worksheet, _
    ByVal bgName As String, _
    ByVal knobName As String, _
    ByRef isOn As Boolean) As Boolean

    Dim bg As Shape
    Dim knob As Shape

    On Error Resume Next
    Set bg = ws.Shapes(bgName)
    Set knob = ws.Shapes(knobName)
    On Error GoTo 0

    If bg Is Nothing Then Exit Function
    If knob Is Nothing Then Exit Function

    isOn = ((knob.Left + (knob.Width / 2)) >= (bg.Left + (bg.Width / 2)))
    TryReadGanttTwoStateToggle = True

End Function

Private Function TryReadGanttAnalyticsPathToggle( _
    ByVal ws As Worksheet, _
    ByRef pathMode As String) As Boolean

    Dim bg As Shape
    Dim knob As Shape
    Dim knobCenter As Double
    Dim ratio As Double

    On Error Resume Next
    Set bg = ws.Shapes(BTN_CP_BG_NAME)
    Set knob = ws.Shapes(BTN_CP_KNOB_NAME)
    On Error GoTo 0

    If bg Is Nothing Then Exit Function
    If knob Is Nothing Then Exit Function
    If bg.Width <= 0 Then Exit Function

    knobCenter = knob.Left + (knob.Width / 2)
    ratio = (knobCenter - bg.Left) / bg.Width

    If bg.Width < 40 Then
        If ratio >= 0.5 Then
            pathMode = GANTT_ANALYTICS_PATH_CP
        Else
            pathMode = GANTT_ANALYTICS_PATH_NONE
        End If
    ElseIf ratio >= 0.66 Then
        pathMode = GANTT_ANALYTICS_PATH_LP
    ElseIf ratio >= 0.34 Then
        pathMode = GANTT_ANALYTICS_PATH_CP
    Else
        pathMode = GANTT_ANALYTICS_PATH_NONE
    End If

    TryReadGanttAnalyticsPathToggle = True

End Function

Public Function GetGanttViewMode() As String

    EnsureGanttViewInitialized
    GetGanttViewMode = gGanttViewMode

End Function

Public Function GetGanttShowCriticalPath() As Boolean
    EnsureGanttViewInitialized
    GetGanttShowCriticalPath = (gAnalyticsPathMode = GANTT_ANALYTICS_PATH_CP)
End Function

Public Sub SetGanttShowCriticalPath(ByVal showValue As Boolean)
    If showValue Then
        gAnalyticsPathMode = GANTT_ANALYTICS_PATH_CP
    Else
        gAnalyticsPathMode = GANTT_ANALYTICS_PATH_NONE
    End If
End Sub
Public Function GetGanttTimelineScaleMode() As String

    EnsureGanttViewInitialized
    GetGanttTimelineScaleMode = gTimelineScaleMode

End Function

Public Sub SetGanttTimelineScaleMode(ByVal scaleMode As String)

    Select Case UCase$(Trim$(scaleMode))
        Case GANTT_SCALE_WEEK
            gTimelineScaleMode = GANTT_SCALE_WEEK
        Case Else
            gTimelineScaleMode = GANTT_SCALE_DAY
    End Select

End Sub

Private Function IsValidDisplayRange(ByVal startVal As Variant, ByVal finishVal As Variant) As Boolean

    If Not HasValue(startVal) Then Exit Function
    If Not HasValue(finishVal) Then Exit Function

    IsValidDisplayRange = (CDbl(finishVal) >= CDbl(startVal))

End Function

Private Function ShouldRenderTaskInCurrentView( _
    ByVal isParent As Boolean, _
    ByVal startVal As Variant, _
    ByVal finishVal As Variant) As Boolean

    ShouldRenderTaskInCurrentView = IsValidDisplayRange(startVal, finishVal)

End Function


Private Function BuildGanttTestInputMap(ByVal ws As Worksheet) As Object

    Dim d As Object
    Dim lastRow As Long
    Dim r As Long
    Dim wbsVal As String

    Set d = CreateObject("Scripting.Dictionary")

    lastRow = GetLastGanttRow(ws)
    If lastRow < FIRST_TASK_ROW Then
        Set BuildGanttTestInputMap = d
        Exit Function
    End If

    For r = FIRST_TASK_ROW To lastRow

        wbsVal = NormalizeWBS(CStr(ws.Cells(r, COL_WBS).value))

        If wbsVal <> "" Then
            d(wbsVal) = Array( _
                GetCellValue(ws.Cells(r, COL_TEST_START).value), _
                GetCellValue(ws.Cells(r, COL_TEST_FINISH).value), _
                GetCellValue(ws.Cells(r, COL_TEST_PROGRESS).value))
        End If

    Next r

    Set BuildGanttTestInputMap = d

End Function

Private Sub RestoreGanttTestInputs(ByVal ws As Worksheet, ByVal testInputMap As Object)

    Dim lastRow As Long
    Dim r As Long
    Dim wbsVal As String
    Dim savedVals As Variant

    If testInputMap Is Nothing Then Exit Sub

    lastRow = GetLastGanttRow(ws)
    If lastRow < FIRST_TASK_ROW Then Exit Sub

    For r = FIRST_TASK_ROW To lastRow

        wbsVal = NormalizeWBS(CStr(ws.Cells(r, COL_WBS).value))

        If wbsVal <> "" Then
            If testInputMap.Exists(wbsVal) Then
                savedVals = testInputMap(wbsVal)

                ws.Cells(r, COL_TEST_START).value = savedVals(0)
                ws.Cells(r, COL_TEST_FINISH).value = savedVals(1)
                ws.Cells(r, COL_TEST_PROGRESS).value = savedVals(2)

                ws.Cells(r, COL_TEST_START).NumberFormat = "dd/mm/yyyy"
                ws.Cells(r, COL_TEST_FINISH).NumberFormat = "dd/mm/yyyy"
                ws.Cells(r, COL_TEST_PROGRESS).NumberFormat = "0%"
            End If
        End If

    Next r

End Sub

Public Sub Refresh_Gantt(Optional ByVal isNewSheet As Boolean = False)

    RunGanttRefreshCore False, isNewSheet

End Sub

Public Sub Refresh_Gantt_DisplayOnly()

    RunGanttRefreshCore True, False

End Sub

Private Sub RunGanttRefreshCore( _
    ByVal displayOnly As Boolean, _
    ByVal isNewSheet As Boolean)

    Dim wsWBS As Worksheet
    Dim wsGantt As Worksheet
    Dim tblWBS As ListObject
    Dim dataArr As Variant
    Dim mapWBS As Object
    Dim hasChildren As Object
    Dim rowById As Object

    Dim projectStart As Variant
    Dim projectFinish As Variant
    Dim totalDays As Long

    Dim rowCount As Long
    Dim ganttRow As Long
    Dim r As Long

    Dim baseById As Object
    Dim testById As Object
    Dim isTestMode As Boolean
    Dim renderMode As String
    Dim renderConstraintMarkers As Boolean
    Dim renderDeadlineMarkers As Boolean
    Dim testInputMap As Object
    Dim calcDrivingMap As Object
    Dim constraintById As Object

    Dim wasGanttSheetCreated As Boolean
    Dim consoleMessages As Collection

    On Error GoTo SafeExit

    Set consoleMessages = New Collection
    Set gExpandedLinks = Nothing

    EnsureGanttViewInitialized
    SetGanttInternalWrite True
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    Set wsWBS = ThisWorkbook.Worksheets(WBS_SHEET)
    Set tblWBS = wsWBS.ListObjects(WBS_TABLE)

    If tblWBS.DataBodyRange Is Nothing Then GoTo SafeExit

    'Important:
    '- EnsureGanttSheet can detect whether the GANTT sheet has just been created.
    '- That information must be propagated to PrepareGanttFullLayout.
    '- Otherwise FreezeGanttAfterFinish is never called after deleting/recreating GANTT.
    Set wsGantt = EnsureGanttSheet(wasGanttSheetCreated)
    If wasGanttSheetCreated Then isNewSheet = True

    If GetGanttPreserveTestInputs() Then
        Set testInputMap = BuildGanttTestInputMap(wsGantt)
    Else
        Set testInputMap = Nothing
    End If

    Set mapWBS = BuildWBSColumnMap(tblWBS)
    ValidateGanttSourceColumns mapWBS

    dataArr = tblWBS.DataBodyRange.value
    rowCount = UBound(dataArr, 1)

    Set hasChildren = BuildHasChildrenMap(dataArr, mapWBS)
    Set rowById = BuildRowByIdMap(dataArr, mapWBS)

    Set baseById = GanttLive_BuildBaseByIdMap()
    Set testById = GanttLive_BuildTestByIdMap()
    renderMode = GanttLive_GetPendingRenderMode()
    isTestMode = GanttLive_IsTestRenderRequested()
    renderConstraintMarkers = (renderMode <> "SCENARIO") And gShowConstraints
    renderDeadlineMarkers = (renderMode = "") And gShowConstraints And IsAnalyticsEnabled()

    If renderConstraintMarkers Or renderDeadlineMarkers Then
        Set constraintById = BuildGanttConstraintMapFromCalc(renderDeadlineMarkers)
    Else
        Set constraintById = CreateObject("Scripting.Dictionary")
    End If

    ResolveDisplayedProjectRange dataArr, mapWBS, hasChildren, baseById, testById, isTestMode, projectStart, projectFinish
    If Not HasValue(projectStart) Or Not HasValue(projectFinish) Then GoTo SafeExit

    If IsWeekScaleMode() Then
        projectStart = GetWeekScaleProjectStart(projectStart)
        projectFinish = GetWeekScaleProjectFinish(projectFinish)
    End If

    totalDays = GetTimelineSlotCount(projectStart, projectFinish)
    If totalDays < 1 Then GoTo SafeExit

    If displayOnly Then
        PrepareGanttDisplayOnlyLayout wsGantt, rowCount, projectStart, totalDays, testInputMap
    Else
        Set calcDrivingMap = BuildCalcDrivingLogicMap()
        PrepareGanttFullLayout wsGantt, dataArr, mapWBS, hasChildren, calcDrivingMap, rowCount, projectStart, totalDays, testInputMap, isNewSheet
    End If

    wsGantt.Activate
    DoEvents

    DrawGanttShapes wsGantt, dataArr, mapWBS, hasChildren, projectStart, totalDays, baseById, testById, isTestMode
    DrawDependencyLinks wsGantt, mapWBS, dataArr, hasChildren, rowById, projectStart, totalDays, baseById, testById, isTestMode
    DrawTodayLine wsGantt, projectStart, totalDays, rowCount
    If renderConstraintMarkers Or renderDeadlineMarkers Then DrawConstraintMarkers_Gantt wsGantt, dataArr, mapWBS, hasChildren, projectStart, totalDays, constraintById, renderConstraintMarkers, renderDeadlineMarkers
    ApplyGanttUiState wsGantt

SafeExit:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    SetGanttInternalWrite False

    If Not GetGanttPreserveTestInputs() Then
        GanttLive_ClearTestRenderRequest
    End If

    If Err.Number <> 0 Then

        If displayOnly Then
            Gantt_AddConsoleMessage consoleMessages, "STOP", _
                "Erreur VBA dans Refresh_Gantt_DisplayOnly" & vbCrLf & _
                "-> vérifier le dernier bloc modifié dans mod_Gantt" & vbCrLf & _
                "-> " & Err.Description, _
                "VBA error in Refresh_Gantt_DisplayOnly" & vbCrLf & _
                "-> check the last edited block in mod_Gantt" & vbCrLf & _
                "-> " & Err.Description
        Else
            Gantt_AddConsoleMessage consoleMessages, "STOP", _
                "Erreur VBA dans Refresh_Gantt" & vbCrLf & _
                "-> vérifier le dernier bloc modifié dans mod_Gantt" & vbCrLf & _
                "-> " & Err.Description, _
                "VBA error in Refresh_Gantt" & vbCrLf & _
                "-> check the last edited block in mod_Gantt" & vbCrLf & _
                "-> " & Err.Description
        End If

        CalcBridge_ShowPlanningConsole consoleMessages

    End If

End Sub


Private Sub PrepareGanttDisplayOnlyLayout( _
    ByVal wsGantt As Worksheet, _
    ByVal rowCount As Long, _
    ByVal projectStart As Variant, _
    ByVal totalDays As Long, _
    ByVal testInputMap As Object)

    If rowCount > 0 Then
        wsGantt.Rows(FIRST_TASK_ROW & ":" & FIRST_TASK_ROW + rowCount - 1).Hidden = False
    End If

    ClearGanttRightPaneOnly wsGantt
    BuildTimeline wsGantt, projectStart, totalDays

    If GetGanttPreserveTestInputs() Then
        RestoreGanttTestInputs wsGantt, testInputMap
    End If

    FinalizeGanttSheet_DisplayOnly wsGantt, totalDays, rowCount

End Sub

Private Sub PrepareGanttFullLayout( _
    ByVal wsGantt As Worksheet, _
    ByVal dataArr As Variant, _
    ByVal mapWBS As Object, _
    ByVal hasChildren As Object, _
    ByVal calcDrivingMap As Object, _
    ByVal rowCount As Long, _
    ByVal projectStart As Variant, _
    ByVal totalDays As Long, _
    ByVal testInputMap As Object, _
    ByVal isNewSheet As Boolean)

    Dim ganttRow As Long
    Dim r As Long

    ClearGanttSheet wsGantt
    SetupStaticLayout wsGantt
    BuildTimeline wsGantt, projectStart, totalDays

    ganttRow = FIRST_TASK_ROW
    For r = 1 To rowCount
        WriteLeftPanelRow wsGantt, ganttRow, dataArr, r, mapWBS
        ApplyRowStyle wsGantt, ganttRow, dataArr, r, mapWBS, hasChildren, calcDrivingMap
        ganttRow = ganttRow + 1
    Next r

    If GetGanttPreserveTestInputs() Then
        RestoreGanttTestInputs wsGantt, testInputMap
    End If

    FinalizeGanttSheet wsGantt, totalDays, rowCount

    If isNewSheet Then
        FreezeGanttAfterFinish wsGantt, rowCount
    End If

End Sub

Private Sub ResolveDisplayedProjectRange( _
    ByVal dataArr As Variant, _
    ByVal mapWBS As Object, _
    ByVal hasChildren As Object, _
    ByVal baseById As Object, _
    ByVal testById As Object, _
    ByVal isTestMode As Boolean, _
    ByRef projectStart As Variant, _
    ByRef projectFinish As Variant)

    GetProjectDisplayRange dataArr, mapWBS, hasChildren, baseById, testById, isTestMode, projectStart, projectFinish

End Sub

Private Function IsManagedGanttShape(ByVal shapeName As String) As Boolean

    IsManagedGanttShape = _
           shapeName = "TODAY_LINE" _
        Or shapeName = "SCURVE_TODAY_LINE" _
        Or Left$(shapeName, 10) = "GANTT_BAR_" _
        Or Left$(shapeName, 13) = "GANTT_LINK_H_" _
        Or Left$(shapeName, 13) = "GANTT_LINK_V_" _
        Or Left$(shapeName, 12) = "GANTT_CP_BAR_" _
        Or Left$(shapeName, 15) = "GANTT_TEST_BOX_" _
        Or Left$(shapeName, 21) = "GANTT_PROGRESS_TEXT_" _
        Or Left$(shapeName, 4) = "SUM_" _
        Or Left$(shapeName, 3) = "MS_" _
        Or Left$(shapeName, 5) = "TASK_" _
        Or Left$(shapeName, 4) = "DEP_" _
        Or Left$(shapeName, 5) = "CSTR_"

End Function

Private Sub DeleteManagedGanttShapes(ByVal wsGantt As Worksheet)

    Dim shp As Shape
    Dim i As Long

    For i = wsGantt.Shapes.Count To 1 Step -1
        Set shp = wsGantt.Shapes(i)
        If IsManagedGanttShape(shp.Name) Then
            shp.Delete
        End If
    Next i

End Sub

Private Sub ClearGanttRightPaneOnly(ByVal wsGantt As Worksheet)

    Dim lastCol As Long

    lastCol = wsGantt.Cells(HEADER_ROW_2, wsGantt.Columns.Count).End(xlToLeft).Column
    If lastCol < FIRST_TIMELINE_COL Then lastCol = FIRST_TIMELINE_COL

    On Error Resume Next
    wsGantt.Range( _
        wsGantt.Cells(HEADER_ROW_1, FIRST_TIMELINE_COL), _
        wsGantt.Cells(HEADER_ROW_2, lastCol) _
    ).UnMerge
    On Error GoTo 0

    wsGantt.Range( _
        wsGantt.Cells(HEADER_ROW_1, FIRST_TIMELINE_COL), _
        wsGantt.Cells(wsGantt.Rows.Count, lastCol) _
    ).ClearContents

    wsGantt.Range( _
        wsGantt.Cells(HEADER_ROW_1, FIRST_TIMELINE_COL), _
        wsGantt.Cells(wsGantt.Rows.Count, lastCol) _
    ).Interior.Pattern = xlNone

    wsGantt.Range( _
        wsGantt.Cells(HEADER_ROW_1, FIRST_TIMELINE_COL), _
        wsGantt.Cells(wsGantt.Rows.Count, lastCol) _
    ).Borders.LineStyle = xlNone

    DeleteManagedGanttShapes wsGantt

End Sub

Private Sub FinalizeGanttSheet(ByVal ws As Worksheet, ByVal totalDays As Long, ByVal rowCount As Long)

    Dim lastCol As Long
    Dim lastRow As Long

    lastCol = FIRST_TIMELINE_COL + totalDays - 1
    lastRow = FIRST_TASK_ROW + rowCount - 1

    ws.Range(ws.Cells(FIRST_TASK_ROW, FIRST_TIMELINE_COL), ws.Cells(lastRow, lastCol)).Borders(xlInsideVertical).LineStyle = xlDot
    ws.Range(ws.Cells(FIRST_TASK_ROW, FIRST_TIMELINE_COL), ws.Cells(lastRow, lastCol)).Borders(xlInsideHorizontal).LineStyle = xlDot

    ws.Range(ws.Cells(FIRST_TASK_ROW, 1), ws.Cells(lastRow, COL_LOGIC)).Borders.LineStyle = xlContinuous

End Sub

Private Function EnsureGanttSheet(Optional ByRef isNewSheet As Boolean = False) As Worksheet

    On Error Resume Next
    Set EnsureGanttSheet = ThisWorkbook.Worksheets(GANTT_SHEET)
    On Error GoTo 0

    If EnsureGanttSheet Is Nothing Then
        Set EnsureGanttSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        EnsureGanttSheet.Name = GANTT_SHEET
        isNewSheet = True
    Else
        isNewSheet = False
    End If

End Function

Private Function BuildWBSColumnMap(ByVal tbl As ListObject) As Object

    Dim d As Object
    Dim i As Long

    Set d = CreateObject("Scripting.Dictionary")

    For i = 1 To tbl.ListColumns.Count
        d(tbl.ListColumns(i).Name) = i
    Next i

    Set BuildWBSColumnMap = d

End Function

Private Sub ValidateGanttSourceColumns(ByVal mapWBS As Object)

    Dim requiredCols As Variant
    Dim c As Variant

    requiredCols = Array( _
        "ID", _
        "WBS", _
        "Task Name", _
        "Calculated Start", _
        "Calculated Finish", _
        "Calculated Duration", _
        "% Progress", _
        "Driving Logic", _
        "Critical Path" _
    )

    For Each c In requiredCols
        If Not mapWBS.Exists(CStr(c)) Then
            Err.Raise vbObjectError + 700, , "Missing source column in tbl_WBS: " & CStr(c)
        End If
    Next c

End Sub

Private Sub ClearGanttSheet(ByVal ws As Worksheet)

    Dim shp As Shape

    For Each shp In ws.Shapes
        shp.Delete
    Next shp

    ws.Cells.Clear

End Sub

Private Sub SetupStaticLayout(ByVal ws As Worksheet)

    ws.Cells(TITLE_ROW, COL_WBS).value = "GANTT VIEW"

    ws.Range("A" & HEADER_ROW_2).value = "WBS"
    ws.Range("B" & HEADER_ROW_2).value = "Task Name"
    ws.Range("C" & HEADER_ROW_2).value = "Start"
    ws.Range("D" & HEADER_ROW_2).value = "Finish"
    ws.Range("E" & HEADER_ROW_2).value = TEST_START_HEADER
    ws.Range("F" & HEADER_ROW_2).value = TEST_FINISH_HEADER
    ws.Range("G" & HEADER_ROW_2).value = "Duration"
    ws.Range("H" & HEADER_ROW_2).value = "%"
    ws.Range("I" & HEADER_ROW_2).value = TEST_PROGRESS_HEADER
    ws.Range("J" & HEADER_ROW_2).value = "Logic"

    ws.Range(ws.Cells(TITLE_ROW, COL_WBS), ws.Cells(TITLE_ROW, COL_LOGIC)).Font.Bold = True
    ws.Range("A" & HEADER_ROW_2 & ":J" & HEADER_ROW_2).Font.Bold = True

    'Pixel-stable heights.
    'Avoid 20 pt task rows: it can create cumulative visual drift between grid and shapes.
    ws.Rows(TITLE_ROW).rowHeight = GANTT_ROW_HEIGHT_HEADER_1
    ws.Rows(TOGGLE_ROW_TOP).rowHeight = GANTT_ROW_HEIGHT_HEADER_2
    ws.Rows(TOGGLE_ROW_BOTTOM).rowHeight = GANTT_ROW_HEIGHT_HEADER_2
    ws.Rows(HEADER_ROW_2).rowHeight = GANTT_ROW_HEIGHT_HEADER_2

    ws.Columns("A").ColumnWidth = 12
    ws.Columns("B").ColumnWidth = 34
    ws.Columns("C:F").ColumnWidth = 11
    ws.Columns("G").ColumnWidth = 9
    ws.Columns("H:I").ColumnWidth = 8

    'Logic column calibrated to avoid text overflow into the timeline at 55% zoom.
    'Measured target: approx. 109 px / Excel width 14.86.
    ws.Columns("J").ColumnWidth = 14.86

    ws.Range("A1:FU3").Interior.Color = RGB(255, 255, 255)
    ws.Range("A1:FU3").Borders.LineStyle = xlNone

    ws.Range("A" & HEADER_ROW_2 & ":J" & HEADER_ROW_2).Interior.Color = RGB(217, 217, 217)
    ws.Range("A" & HEADER_ROW_2 & ":J" & HEADER_ROW_2).Borders.LineStyle = xlContinuous

End Sub




Private Function IsWeekScaleMode() As Boolean

    EnsureGanttViewInitialized
    IsWeekScaleMode = (gTimelineScaleMode = GANTT_SCALE_WEEK)

End Function

Private Function GetIsoWeekMonday(ByVal anyDate As Date) As Date

    GetIsoWeekMonday = CDate(anyDate - Weekday(anyDate, vbMonday) + 1)

End Function

Private Function GetIsoWeekLabel(ByVal anyDate As Date) As String

    GetIsoWeekLabel = "W" & Format$(WorksheetFunction.IsoWeekNum(anyDate), "00")

End Function

Private Function GetIsoWeekYear(ByVal anyDate As Date) As Long

    Dim thursdayDate As Date

    thursdayDate = GetIsoWeekMonday(anyDate) + 3
    GetIsoWeekYear = Year(thursdayDate)

End Function

Private Function GetWeekScaleProjectStart(ByVal rawProjectStart As Variant) As Variant

    If Not HasValue(rawProjectStart) Then Exit Function
    GetWeekScaleProjectStart = GetIsoWeekMonday(CDate(rawProjectStart))

End Function

Private Function GetWeekScaleProjectFinish(ByVal rawProjectFinish As Variant) As Variant

    If Not HasValue(rawProjectFinish) Then Exit Function
    GetWeekScaleProjectFinish = GetIsoWeekMonday(CDate(rawProjectFinish)) + 6

End Function

Private Function GetTimelineSlotCount(ByVal projectStart As Variant, ByVal projectFinish As Variant) As Long

    Dim startVal As Date
    Dim finishVal As Date

    If Not HasValue(projectStart) Then Exit Function
    If Not HasValue(projectFinish) Then Exit Function

    startVal = CDate(projectStart)
    finishVal = CDate(projectFinish)

    If finishVal < startVal Then Exit Function

    If IsWeekScaleMode() Then
        GetTimelineSlotCount = CLng(DateDiff("ww", GetIsoWeekMonday(startVal), GetIsoWeekMonday(finishVal), vbMonday, vbFirstFourDays)) + 1
    Else
        GetTimelineSlotCount = CLng(finishVal - startVal + 1)
    End If

End Function

Private Function GetTimelineSlotIndex(ByVal projectStart As Variant, ByVal anyDate As Variant) As Long

    Dim startVal As Date
    Dim dateVal As Date

    If Not HasValue(projectStart) Then
        GetTimelineSlotIndex = -1
        Exit Function
    End If

    If Not HasValue(anyDate) Then
        GetTimelineSlotIndex = -1
        Exit Function
    End If

    startVal = CDate(projectStart)
    dateVal = CDate(anyDate)

    If IsWeekScaleMode() Then
        GetTimelineSlotIndex = CLng(DateDiff("ww", GetIsoWeekMonday(startVal), GetIsoWeekMonday(dateVal), vbMonday, vbFirstFourDays))
    Else
        GetTimelineSlotIndex = CLng(CDbl(dateVal) - CDbl(startVal))
    End If

End Function


Private Function GetRenderStartForCurrentScale(ByVal startVal As Variant) As Variant

    If Not HasValue(startVal) Then Exit Function

    If IsWeekScaleMode() Then
        GetRenderStartForCurrentScale = GetIsoWeekMonday(CDate(startVal))
    Else
        GetRenderStartForCurrentScale = startVal
    End If

End Function

Private Function GetRenderFinishForCurrentScale(ByVal finishVal As Variant) As Variant

    If Not HasValue(finishVal) Then Exit Function

    If IsWeekScaleMode() Then
        GetRenderFinishForCurrentScale = GetIsoWeekMonday(CDate(finishVal)) + 6
    Else
        GetRenderFinishForCurrentScale = finishVal
    End If

End Function


Private Sub BuildTimeline(ByVal ws As Worksheet, ByVal projectStart As Variant, ByVal slotCount As Long)

    If IsWeekScaleMode() Then
        BuildTimeline_Week ws, projectStart, slotCount
    Else
        BuildTimeline_Day ws, projectStart, slotCount
    End If

End Sub

Private Sub BuildTimeline_Day(ByVal ws As Worksheet, ByVal projectStart As Variant, ByVal totalDays As Long)

    Dim i As Long
    Dim currentDate As Date
    Dim currentMonth As Long
    Dim monthStartCol As Long

    currentMonth = 0
    monthStartCol = FIRST_TIMELINE_COL

    For i = 0 To totalDays - 1
        currentDate = CDate(CDbl(projectStart) + i)

        ws.Cells(HEADER_ROW_2, FIRST_TIMELINE_COL + i).value = currentDate
        ws.Cells(HEADER_ROW_2, FIRST_TIMELINE_COL + i).NumberFormat = "dd"
        ws.Cells(HEADER_ROW_2, FIRST_TIMELINE_COL + i).HorizontalAlignment = xlCenter
        ws.Cells(HEADER_ROW_2, FIRST_TIMELINE_COL + i).ColumnWidth = 4

        If Month(currentDate) <> currentMonth Then
            If i > 0 Then
                With ws.Range(ws.Cells(HEADER_ROW_1, monthStartCol), ws.Cells(HEADER_ROW_1, FIRST_TIMELINE_COL + i - 1))
                    .Merge
                    .HorizontalAlignment = xlCenter
                    .VerticalAlignment = xlCenter
                    .value = Format(DateSerial(Year(CDate(CDbl(projectStart) + i - 1)), Month(CDate(CDbl(projectStart) + i - 1)), 1), "mmm yyyy")
                    .Interior.Color = RGB(191, 191, 191)
                End With
            End If

            currentMonth = Month(currentDate)
            monthStartCol = FIRST_TIMELINE_COL + i
        End If

        If Weekday(currentDate, vbMonday) >= 6 Then
            ws.Range(ws.Cells(FIRST_TASK_ROW, FIRST_TIMELINE_COL + i), ws.Cells(FIRST_TASK_ROW + 500, FIRST_TIMELINE_COL + i)).Interior.Color = RGB(242, 242, 242)
        End If
    Next i

    With ws.Range(ws.Cells(HEADER_ROW_1, monthStartCol), ws.Cells(HEADER_ROW_1, FIRST_TIMELINE_COL + totalDays - 1))
        .Merge
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .value = Format(DateSerial(Year(CDate(CDbl(projectStart) + totalDays - 1)), Month(CDate(CDbl(projectStart) + totalDays - 1)), 1), "mmm yyyy")
        .Interior.Color = RGB(191, 191, 191)
    End With

    ws.Range(ws.Cells(HEADER_ROW_1, FIRST_TIMELINE_COL), ws.Cells(HEADER_ROW_2, FIRST_TIMELINE_COL + totalDays - 1)).Borders.LineStyle = xlContinuous

End Sub

Private Sub BuildTimeline_Week(ByVal ws As Worksheet, ByVal projectStart As Variant, ByVal slotCount As Long)

    Dim i As Long
    Dim weekStart As Date
    Dim isoYear As Long
    Dim currentYear As Long
    Dim yearStartCol As Long

    currentYear = 0
    yearStartCol = FIRST_TIMELINE_COL

    For i = 0 To slotCount - 1

        weekStart = CDate(projectStart) + (i * 7)
        isoYear = GetIsoWeekYear(weekStart)

        ws.Cells(HEADER_ROW_2, FIRST_TIMELINE_COL + i).value = GetIsoWeekLabel(weekStart)
        ws.Cells(HEADER_ROW_2, FIRST_TIMELINE_COL + i).HorizontalAlignment = xlCenter
        ws.Cells(HEADER_ROW_2, FIRST_TIMELINE_COL + i).VerticalAlignment = xlCenter
        ws.Cells(HEADER_ROW_2, FIRST_TIMELINE_COL + i).ColumnWidth = 6

        If isoYear <> currentYear Then
            If i > 0 Then
                With ws.Range(ws.Cells(HEADER_ROW_1, yearStartCol), ws.Cells(HEADER_ROW_1, FIRST_TIMELINE_COL + i - 1))
                    .Merge
                    .HorizontalAlignment = xlCenter
                    .VerticalAlignment = xlCenter
                    .value = CStr(currentYear)
                    .Interior.Color = RGB(191, 191, 191)
                End With
            End If

            currentYear = isoYear
            yearStartCol = FIRST_TIMELINE_COL + i
        End If
    Next i

    With ws.Range(ws.Cells(HEADER_ROW_1, yearStartCol), ws.Cells(HEADER_ROW_1, FIRST_TIMELINE_COL + slotCount - 1))
        .Merge
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .value = CStr(currentYear)
        .Interior.Color = RGB(191, 191, 191)
    End With

    ws.Range(ws.Cells(HEADER_ROW_1, FIRST_TIMELINE_COL), ws.Cells(HEADER_ROW_2, FIRST_TIMELINE_COL + slotCount - 1)).Borders.LineStyle = xlContinuous

End Sub

Private Sub WriteLeftPanelRow(ByVal ws As Worksheet, ByVal ganttRow As Long, ByRef dataArr As Variant, ByVal dataRow As Long, ByVal mapWBS As Object)

    Dim logicVal As String
    Dim isLoE As Boolean

    ws.Rows(ganttRow).rowHeight = GANTT_ROW_HEIGHT_TASK

    logicVal = ""
    If mapWBS.Exists("Driving Logic") Then
        logicVal = CStr(dataArr(dataRow, mapWBS("Driving Logic")))
    End If

    isLoE = IsLevelOfEffortTaskType(dataArr, mapWBS, dataRow)

    ws.Cells(ganttRow, COL_WBS).NumberFormat = "@"
    ws.Cells(ganttRow, COL_WBS).value = NormalizeWBS(CStr(dataArr(dataRow, mapWBS("WBS"))))
    ws.Cells(ganttRow, COL_TASK).value = dataArr(dataRow, mapWBS("Task Name"))
    ws.Cells(ganttRow, COL_START).value = dataArr(dataRow, mapWBS("Calculated Start"))
    ws.Cells(ganttRow, COL_FINISH).value = dataArr(dataRow, mapWBS("Calculated Finish"))

    If Not GetGanttPreserveTestInputs() Then
        ws.Cells(ganttRow, COL_TEST_START).ClearContents
        ws.Cells(ganttRow, COL_TEST_FINISH).ClearContents
        ws.Cells(ganttRow, COL_TEST_PROGRESS).ClearContents
    End If

    ws.Cells(ganttRow, COL_DURATION).value = dataArr(dataRow, mapWBS("Calculated Duration"))

    If isLoE Then

        'Primavera-style LOE progress:
        '- Dynamic with TODAY()
        '- Based on displayed Calculated Start / Calculated Finish in columns C / D
        '- Clamped between 0% and 100%
        '- Formula is written in English/Formula syntax so Excel stores it robustly.
        ws.Cells(ganttRow, COL_PROGRESS).Formula = _
            "=IF(OR(" & _
                ws.Cells(ganttRow, COL_START).Address(False, False) & "=""""," & _
                ws.Cells(ganttRow, COL_FINISH).Address(False, False) & "=""""),0," & _
              "MAX(0,MIN(1,(TODAY()-" & ws.Cells(ganttRow, COL_START).Address(False, False) & "+1)/(" & _
                               ws.Cells(ganttRow, COL_FINISH).Address(False, False) & "-" & _
                               ws.Cells(ganttRow, COL_START).Address(False, False) & "+1))))"

    Else

        If HasValue(dataArr(dataRow, mapWBS("% Progress"))) Then
            ws.Cells(ganttRow, COL_PROGRESS).value = dataArr(dataRow, mapWBS("% Progress"))
        Else
            ws.Cells(ganttRow, COL_PROGRESS).value = 0
        End If

    End If

    ws.Cells(ganttRow, COL_LOGIC).value = logicVal

    ws.Cells(ganttRow, COL_START).NumberFormat = "dd/mm/yyyy"
    ws.Cells(ganttRow, COL_FINISH).NumberFormat = "dd/mm/yyyy"
    ws.Cells(ganttRow, COL_TEST_START).NumberFormat = "dd/mm/yyyy"
    ws.Cells(ganttRow, COL_TEST_FINISH).NumberFormat = "dd/mm/yyyy"
    ws.Cells(ganttRow, COL_PROGRESS).NumberFormat = "0%"
    ws.Cells(ganttRow, COL_TEST_PROGRESS).NumberFormat = "0%"

End Sub

Private Sub ApplyRowStyle(ByVal ws As Worksheet, ByVal ganttRow As Long, ByRef dataArr As Variant, ByVal dataRow As Long, ByVal mapWBS As Object, ByVal hasChildren As Object, ByVal calcDrivingMap As Object)

    Dim wbs As String
    Dim levelCount As Long
    Dim idVal As String
    Dim isLeaf As Boolean
    Dim hasActual As Boolean
    Dim isLoE As Boolean
    Dim logicVal As String

    wbs = NormalizeWBS(CStr(dataArr(dataRow, mapWBS("WBS"))))
    levelCount = WBSLevel(wbs)
    idVal = Trim$(CStr(dataArr(dataRow, mapWBS("ID"))))
    isLeaf = Not hasChildren.Exists(wbs)

    logicVal = ""
    If mapWBS.Exists("Driving Logic") Then
        logicVal = CStr(dataArr(dataRow, mapWBS("Driving Logic")))
    End If

    isLoE = IsLevelOfEffortTaskType(dataArr, mapWBS, dataRow)

    hasActual = False
    If mapWBS.Exists("Actual Start") Then
        If HasValue(dataArr(dataRow, mapWBS("Actual Start"))) Then hasActual = True
    End If
    If mapWBS.Exists("Actual Finish") Then
        If HasValue(dataArr(dataRow, mapWBS("Actual Finish"))) Then hasActual = True
    End If

    ws.Cells(ganttRow, COL_TASK).IndentLevel = WorksheetFunction.Min(levelCount - 1, 15)

    If hasChildren.Exists(wbs) Then
        ws.Range(ws.Cells(ganttRow, 1), ws.Cells(ganttRow, COL_LOGIC)).Font.Bold = True
        ws.Range(ws.Cells(ganttRow, 1), ws.Cells(ganttRow, COL_LOGIC)).Interior.Color = RGB(248, 248, 248)
    End If

    ApplyTestCellColoring ws, ganttRow, isLeaf, hasActual, isLoE

    ws.Range(ws.Cells(ganttRow, 1), ws.Cells(ganttRow, COL_LOGIC)).Borders(xlEdgeBottom).LineStyle = xlContinuous
    ws.Range(ws.Cells(ganttRow, FIRST_TIMELINE_COL), ws.Cells(ganttRow, FIRST_TIMELINE_COL + 500)).Borders(xlEdgeBottom).LineStyle = xlDot

End Sub

Private Sub DrawGanttShapes( _
    ByVal ws As Worksheet, _
    ByRef dataArr As Variant, _
    ByVal mapWBS As Object, _
    ByVal hasChildren As Object, _
    ByVal projectStart As Variant, _
    ByVal totalDays As Long, _
    ByVal baseById As Object, _
    ByVal testById As Object, _
    ByVal isTestMode As Boolean)

    Dim r As Long
    Dim rowCount As Long
    Dim ganttRow As Long

    Dim wbs As String
    Dim idVal As String

    Dim rawStartVal As Variant
    Dim rawFinishVal As Variant
    Dim renderStartVal As Variant
    Dim renderFinishVal As Variant

    Dim rawDurationDays As Long
    Dim progressVal As Double
    Dim isCritical As Boolean
    Dim hasDelta As Boolean
    Dim isParent As Boolean
    Dim isLoE As Boolean
    Dim isMilestone As Boolean
    Dim parentCompleteMap As Object

    rowCount = UBound(dataArr, 1)
    Set parentCompleteMap = BuildParentCompleteMap(dataArr, mapWBS)

    For r = 1 To rowCount
        ganttRow = FIRST_TASK_ROW + r - 1

        wbs = NormalizeWBS(CStr(dataArr(r, mapWBS("WBS"))))
        idVal = Trim$(CStr(dataArr(r, mapWBS("ID"))))

        rawStartVal = GanttLive_GetDisplayStart(idVal, baseById, testById, isTestMode)
        rawFinishVal = GanttLive_GetDisplayFinish(idVal, baseById, testById, isTestMode)

        renderStartVal = GetRenderStartForCurrentScale(rawStartVal)
        renderFinishVal = GetRenderFinishForCurrentScale(rawFinishVal)

        isParent = hasChildren.Exists(wbs)
        isMilestone = IsMilestoneTaskType(dataArr, mapWBS, r)
        isLoE = IsLevelOfEffortTaskType(dataArr, mapWBS, r)

        If Not ShouldRenderTaskInCurrentView(isParent, renderStartVal, renderFinishVal) Then GoTo NextShape
        If Not HasValue(rawStartVal) Or Not HasValue(rawFinishVal) Then GoTo NextShape
        If Not HasValue(renderStartVal) Or Not HasValue(renderFinishVal) Then GoTo NextShape

        rawDurationDays = CLng(CDbl(rawFinishVal) - CDbl(rawStartVal) + 1)

        progressVal = 0

        If isLoE Then

            'LOE Primavera-style progress:
            'The visual progress must be derived from TODAY and LOE calculated dates,
            'not from the manual % Progress stored in WBS/CALC.
            progressVal = GetLoEProgressFromToday(rawStartVal, rawFinishVal)

        Else

            If HasValue(GanttLive_GetDisplayProgress(idVal, baseById, testById, isTestMode)) Then
                progressVal = CDbl(GanttLive_GetDisplayProgress(idVal, baseById, testById, isTestMode))
            End If

        End If

        isCritical = ShouldHighlightGanttAnalyticsPath(dataArr, mapWBS, r)

        hasDelta = False
        If isTestMode Then
            hasDelta = GanttLive_HasRenderableTestDelta(idVal, baseById, testById)
        End If

        If isParent Then

            DrawSummaryBar ws, ganttRow, projectStart, renderStartVal, renderFinishVal, isCritical, totalDays, _
                "SUM_" & CStr(r), Trim$(CStr(dataArr(r, mapWBS("Task Name")))), hasDelta, _
                ParentIsCompleteFromMap(wbs, parentCompleteMap)

        ElseIf isMilestone Then

            DrawMilestone ws, ganttRow, projectStart, renderStartVal, progressVal, isCritical, totalDays, "MS_" & CStr(r), hasDelta

        ElseIf rawDurationDays <= 1 Then

            DrawSingleWeekTask ws, ganttRow, projectStart, renderStartVal, progressVal, isCritical, totalDays, _
                "TASK_" & CStr(r), hasDelta

        ElseIf IsWeekScaleMode() And IsTaskContainedInSingleIsoWeek(rawStartVal, rawFinishVal) Then

            DrawSingleWeekTask ws, ganttRow, projectStart, renderStartVal, progressVal, isCritical, totalDays, _
                "TASK_" & CStr(r), hasDelta

        Else

            DrawTaskBar ws, ganttRow, projectStart, renderStartVal, renderFinishVal, progressVal, isCritical, totalDays, _
                "TASK_" & CStr(r), hasDelta, rawStartVal, rawFinishVal, isLoE

        End If

NextShape:
    Next r

End Sub

Private Sub DrawTaskBar( _
    ByVal ws As Worksheet, _
    ByVal ganttRow As Long, _
    ByVal projectStart As Variant, _
    ByVal startVal As Variant, _
    ByVal finishVal As Variant, _
    ByVal progressVal As Double, _
    ByVal isCritical As Boolean, _
    ByVal totalDays As Long, _
    ByVal shapeKey As String, _
    ByVal hasDelta As Boolean, _
    ByVal rawStartVal As Variant, _
    ByVal rawFinishVal As Variant, _
    Optional ByVal isLoE As Boolean = False)

    Dim leftPos As Double
    Dim rightPos As Double
    Dim topPos As Double
    Dim barWidth As Double
    Dim barHeight As Double
    Dim fullBarTop As Double
    Dim fullBarHeight As Double
    Dim shp As Shape
    Dim shpProgress As Shape

    Dim progressColor As Variant
    Dim progressWidth As Double
    Dim progressLeft As Double
    Dim progressTop As Double
    Dim progressHeight As Double

    If Not HasValue(startVal) Then Exit Sub
    If Not HasValue(finishVal) Then Exit Sub

    'Strict date-column anchoring.
    'The bar geometry must be derived from the full Excel range/cell anchors,
    'not from accumulated pixel offsets.
    leftPos = TimelineLeft(ws, projectStart, startVal)
    rightPos = TimelineRightAfterFinish(ws, projectStart, finishVal)

    If rightPos <= leftPos Then Exit Sub

    fullBarTop = GetGanttBarTop(ws, ganttRow)
    fullBarHeight = GetGanttBarHeight(ws, ganttRow)

    If isLoE Then
        barHeight = fullBarHeight / 2
        If barHeight < 2 Then barHeight = 2
        topPos = fullBarTop + ((fullBarHeight - barHeight) / 2)
    Else
        barHeight = fullBarHeight
        topPos = fullBarTop
    End If

    barWidth = rightPos - leftPos

    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, leftPos, topPos, barWidth, barHeight)
    shp.Name = shapeKey
    ApplyGanttRenderShapePlacement shp

    With shp
        .Left = leftPos
        .Top = topPos
        .Width = barWidth
        .Height = barHeight
    End With

    If progressVal >= 1 Then
        shp.Fill.ForeColor.RGB = COLOR_PROGRESS_GREEN
    Else
        shp.Fill.ForeColor.RGB = GetTaskBaseColor(isCritical)
    End If

    If hasDelta Then
        shp.Line.Visible = msoTrue
        shp.Line.ForeColor.RGB = RGB(255, 192, 0)
        shp.Line.Weight = 2.75
    Else
        shp.Line.Visible = msoFalse
    End If

    If progressVal > 0 And progressVal < 1 Then

        progressColor = GetProgressFillColor(rawStartVal, rawFinishVal, progressVal)

        If Not IsEmpty(progressColor) Then

            progressWidth = barWidth * WorksheetFunction.Min(progressVal, 1)
            If progressWidth < 2 Then progressWidth = 2

            progressLeft = leftPos + 2
            progressTop = topPos + 2
            progressHeight = barHeight - 4

            If progressWidth > barWidth - 4 Then progressWidth = barWidth - 4
            If progressHeight < 2 Then progressHeight = 2

            Set shpProgress = ws.Shapes.AddShape( _
                msoShapeRoundedRectangle, _
                progressLeft, _
                progressTop, _
                progressWidth, _
                progressHeight)

            shpProgress.Name = shapeKey & "_P"
            ApplyGanttRenderShapePlacement shpProgress

            With shpProgress
                .Left = progressLeft
                .Top = progressTop
                .Width = progressWidth
                .Height = progressHeight
                .Line.Visible = msoFalse
                .Fill.ForeColor.RGB = CLng(progressColor)
                .Fill.Transparency = 0.25
            End With

        End If

    End If

End Sub



Private Function TimelineRightAfterFinish( _
    ByVal ws As Worksheet, _
    ByVal projectStart As Variant, _
    ByVal taskFinish As Variant) As Double

    Dim targetCol As Long

    targetCol = TimelineColumnFromHeaderDate_Exact(ws, projectStart, taskFinish)
    If targetCol < FIRST_TIMELINE_COL Then Exit Function

    TimelineRightAfterFinish = ws.Cells(HEADER_ROW_2, targetCol).Left + ws.Cells(HEADER_ROW_2, targetCol).Width

End Function

Private Function TimelineRight( _
    ByVal ws As Worksheet, _
    ByVal projectStart As Variant, _
    ByVal taskFinish As Variant) As Double

    TimelineRight = TimelineRightAfterFinish(ws, projectStart, taskFinish)

End Function

Private Sub DrawMilestone( _
    ByVal ws As Worksheet, _
    ByVal ganttRow As Long, _
    ByVal projectStart As Variant, _
    ByVal startVal As Variant, _
    ByVal progressVal As Double, _
    ByVal isCritical As Boolean, _
    ByVal totalDays As Long, _
    ByVal shapeKey As String, _
    ByVal hasDelta As Boolean)

    Dim leftPos As Double
    Dim topPos As Double
    Dim sizeVal As Double
    Dim cellMidX As Double
    Dim shp As Shape

    If Not HasValue(startVal) Then Exit Sub

    sizeVal = GetGanttMilestoneSize(ws, ganttRow)
    cellMidX = GetTaskMidX(ws, projectStart, startVal)
    If cellMidX <= 0 Then Exit Sub

    leftPos = cellMidX - (sizeVal / 2)
    topPos = GetGanttRowTop(ws, ganttRow) + ((GetGanttRowHeight(ws, ganttRow) - sizeVal) / 2)

    Set shp = ws.Shapes.AddShape(msoShapeDiamond, leftPos, topPos, sizeVal, sizeVal)
    shp.Name = shapeKey
    ApplyGanttRenderShapePlacement shp

    With shp
        .Left = leftPos
        .Top = topPos
        .Width = sizeVal
        .Height = sizeVal
    End With

    If progressVal >= 1 Then
        shp.Fill.ForeColor.RGB = COLOR_PROGRESS_GREEN
    Else
        shp.Fill.ForeColor.RGB = GetTaskBaseColor(isCritical)
    End If

    If hasDelta Then
        shp.Line.Visible = msoTrue
        shp.Line.ForeColor.RGB = RGB(255, 192, 0)
        shp.Line.Weight = 2.75
    Else
        shp.Line.Visible = msoFalse
    End If

End Sub


Private Function BuildGanttConstraintMapFromCalc(Optional ByVal includeDeadline As Boolean = False) As Object

    Dim d As Object
    Dim wsCalc As Worksheet
    Dim tblCalc As ListObject
    Dim mapCalc As Object
    Dim arr As Variant
    Dim r As Long
    Dim idVal As String
    Dim activeVal As String
    Dim isSummaryVal As String
    Dim drivingLogicVal As String
    Dim taskTypeVal As String
    Dim deadlineVal As Variant
    Dim hasDeadline As Boolean

    Set d = CreateObject("Scripting.Dictionary")

    Set wsCalc = ThisWorkbook.Worksheets(CALC_SHEET)
    Set tblCalc = wsCalc.ListObjects(CALC_TABLE)

    If tblCalc.DataBodyRange Is Nothing Then
        Set BuildGanttConstraintMapFromCalc = d
        Exit Function
    End If

    Set mapCalc = BuildWBSColumnMap(tblCalc)
    RequireGanttCalcColumn mapCalc, "ID", "BuildGanttConstraintMapFromCalc"
    RequireGanttCalcColumn mapCalc, "Constraint Active", "BuildGanttConstraintMapFromCalc"
    RequireGanttCalcColumn mapCalc, "Start Constraint Type", "BuildGanttConstraintMapFromCalc"
    RequireGanttCalcColumn mapCalc, "Start Constraint Date", "BuildGanttConstraintMapFromCalc"
    RequireGanttCalcColumn mapCalc, "Finish Constraint Type", "BuildGanttConstraintMapFromCalc"
    RequireGanttCalcColumn mapCalc, "Finish Constraint Date", "BuildGanttConstraintMapFromCalc"
    If includeDeadline Then RequireGanttCalcColumn mapCalc, "Deadline", "BuildGanttConstraintMapFromCalc"

    arr = tblCalc.DataBodyRange.value

    For r = 1 To UBound(arr, 1)

        idVal = Trim$(CStr(arr(r, mapCalc("ID"))))
        If idVal = "" Then GoTo NextRow

        isSummaryVal = ""
        drivingLogicVal = ""
        taskTypeVal = ""

        If mapCalc.Exists("IsSummary") Then isSummaryVal = UCase$(Trim$(CStr(arr(r, mapCalc("IsSummary")))))
        If mapCalc.Exists("Driving Logic") Then drivingLogicVal = UCase$(Trim$(CStr(arr(r, mapCalc("Driving Logic")))))
        If mapCalc.Exists("Task Type") Then taskTypeVal = UCase$(Trim$(CStr(arr(r, mapCalc("Task Type")))))

        If isSummaryVal = "TRUE" Or isSummaryVal = "YES" Then GoTo NextRow
        If drivingLogicVal = "SUMMARY" Or drivingLogicVal = "LOE" Then GoTo NextRow
        If InStr(1, taskTypeVal, "LEVEL OF EFFORT", vbTextCompare) > 0 Then GoTo NextRow

        activeVal = UCase$(Trim$(CStr(arr(r, mapCalc("Constraint Active")))))
        deadlineVal = Empty
        hasDeadline = False
        If includeDeadline Then
            deadlineVal = GetCellValue(arr(r, mapCalc("Deadline")))
            hasDeadline = HasValue(deadlineVal)
        End If

        If activeVal <> "YES" And Not hasDeadline Then GoTo NextRow

        d(idVal) = Array( _
            Trim$(CStr(arr(r, mapCalc("Start Constraint Type")))), _
            GetCellValue(arr(r, mapCalc("Start Constraint Date"))), _
            Trim$(CStr(arr(r, mapCalc("Finish Constraint Type")))), _
            GetCellValue(arr(r, mapCalc("Finish Constraint Date"))), _
            deadlineVal _
        )

NextRow:
    Next r

    Set BuildGanttConstraintMapFromCalc = d

End Function

Private Sub RequireGanttCalcColumn( _
    ByVal mapCalc As Object, _
    ByVal colName As String, _
    ByVal functionName As String)

    If Not mapCalc.Exists(colName) Then
        Err.Raise vbObjectError + 760, functionName, "Missing source column in tbl_CALC: " & colName
    End If

End Sub

Private Sub DrawConstraintMarkers_Gantt( _
    ByVal ws As Worksheet, _
    ByRef dataArr As Variant, _
    ByVal mapWBS As Object, _
    ByVal hasChildren As Object, _
    ByVal projectStart As Variant, _
    ByVal totalDays As Long, _
    ByVal constraintById As Object, _
    ByVal drawHardConstraints As Boolean, _
    ByVal drawDeadlineMarkers As Boolean)

    Dim r As Long
    Dim rowCount As Long
    Dim ganttRow As Long
    Dim idVal As String
    Dim wbsVal As String
    Dim vals As Variant
    Dim startType As String
    Dim finishType As String
    Dim startDate As Variant
    Dim finishDate As Variant
    Dim deadlineDate As Variant
    Dim isLoE As Boolean
    Dim isMilestone As Boolean

    If constraintById Is Nothing Then Exit Sub
    If constraintById.Count = 0 Then Exit Sub
    If Not HasValue(projectStart) Then Exit Sub
    If totalDays < 1 Then Exit Sub

    rowCount = UBound(dataArr, 1)

    For r = 1 To rowCount

        ganttRow = FIRST_TASK_ROW + r - 1
        If ws.Rows(ganttRow).Hidden Then GoTo NextRow

        idVal = Trim$(CStr(dataArr(r, mapWBS("ID"))))
        If idVal = "" Then GoTo NextRow
        If Not constraintById.Exists(idVal) Then GoTo NextRow

        wbsVal = NormalizeWBS(CStr(dataArr(r, mapWBS("WBS"))))
        If hasChildren.Exists(wbsVal) Then GoTo NextRow

        isLoE = IsLevelOfEffortTaskType(dataArr, mapWBS, r)
        If isLoE Then GoTo NextRow

        vals = constraintById(idVal)
        startType = UCase$(Trim$(CStr(vals(0))))
        startDate = vals(1)
        finishType = UCase$(Trim$(CStr(vals(2))))
        finishDate = vals(3)
        deadlineDate = Empty
        If UBound(vals) >= 4 Then deadlineDate = vals(4)

        If drawDeadlineMarkers And HasValue(deadlineDate) Then
            ApplyConstraintCellBorder_Gantt ws, ganttRow, projectStart, totalDays, deadlineDate, "RIGHT"
        End If

        If Not drawHardConstraints Then GoTo NextRow

        isMilestone = IsMilestoneTaskType(dataArr, mapWBS, r)

        If isMilestone _
            And startType = "MUST START ON" _
            And finishType = "MUST FINISH ON" _
            And HasValue(startDate) _
            And HasValue(finishDate) _
            And CLng(CDbl(startDate)) = CLng(CDbl(finishDate)) Then

            ApplyConstraintCellBorder_Gantt ws, ganttRow, projectStart, totalDays, startDate, "FULL"
            GoTo NextRow
        End If

        Select Case startType
            Case "START NO EARLIER THAN"
                ApplyConstraintCellBorder_Gantt ws, ganttRow, projectStart, totalDays, startDate, "LEFT"
                DrawConstraintBigCrossAtCellEdge_Gantt ws, ganttRow, projectStart, totalDays, startDate, "LEFT", CStr(r), "SNET"
            Case "START NO LATER THAN"
                ApplyConstraintCellBorder_Gantt ws, ganttRow, projectStart, totalDays, startDate, "RIGHT"
                DrawConstraintSmallCrossesAtCellEdge_Gantt ws, ganttRow, projectStart, totalDays, startDate, "RIGHT", "RIGHT", CStr(r), "SNLT"
            Case "MUST START ON"
                ApplyConstraintCellBorder_Gantt ws, ganttRow, projectStart, totalDays, startDate, "LEFT"
                DrawConstraintBigCrossAtCellEdge_Gantt ws, ganttRow, projectStart, totalDays, startDate, "LEFT", CStr(r), "SNET"
                DrawConstraintSmallCrossesAtCellEdge_Gantt ws, ganttRow, projectStart, totalDays, startDate, "LEFT", "RIGHT", CStr(r), "SNLT"
        End Select

        Select Case finishType
            Case "FINISH NO EARLIER THAN"
                ApplyConstraintCellBorder_Gantt ws, ganttRow, projectStart, totalDays, finishDate, "LEFT"
                DrawConstraintSmallCrossesAtCellEdge_Gantt ws, ganttRow, projectStart, totalDays, finishDate, "LEFT", "LEFT", CStr(r), "FNET"
            Case "FINISH NO LATER THAN"
                ApplyConstraintCellBorder_Gantt ws, ganttRow, projectStart, totalDays, finishDate, "RIGHT"
                DrawConstraintBigCrossAtCellEdge_Gantt ws, ganttRow, projectStart, totalDays, finishDate, "RIGHT", CStr(r), "FNLT"
            Case "MUST FINISH ON"
                ApplyConstraintCellBorder_Gantt ws, ganttRow, projectStart, totalDays, finishDate, "RIGHT"
                DrawConstraintSmallCrossesAtCellEdge_Gantt ws, ganttRow, projectStart, totalDays, finishDate, "RIGHT", "LEFT", CStr(r), "FNET"
                DrawConstraintBigCrossAtCellEdge_Gantt ws, ganttRow, projectStart, totalDays, finishDate, "RIGHT", CStr(r), "FNLT"
        End Select

NextRow:
    Next r

End Sub

Private Function GetConstraintTargetCell_Gantt( _
    ByVal ws As Worksheet, _
    ByVal ganttRow As Long, _
    ByVal projectStart As Variant, _
    ByVal totalDays As Long, _
    ByVal constraintDate As Variant) As Range

    Dim renderDate As Variant
    Dim projectFinish As Date
    Dim targetCol As Long

    If Not HasValue(constraintDate) Then Exit Function

    renderDate = GetRenderStartForCurrentScale(constraintDate)
    If Not HasValue(renderDate) Then Exit Function

    If IsWeekScaleMode() Then
        projectFinish = CDate(projectStart) + ((totalDays - 1) * 7) + 6
    Else
        projectFinish = CDate(CDbl(projectStart) + totalDays - 1)
    End If

    If CDate(renderDate) < CDate(projectStart) Or CDate(renderDate) > projectFinish Then Exit Function

    targetCol = TimelineColumnFromHeaderDate_Exact(ws, projectStart, renderDate)
    If targetCol < FIRST_TIMELINE_COL Then Exit Function

    Set GetConstraintTargetCell_Gantt = ws.Cells(ganttRow, targetCol)

End Function

Private Sub ApplyConstraintCellBorder_Gantt( _
    ByVal ws As Worksheet, _
    ByVal ganttRow As Long, _
    ByVal projectStart As Variant, _
    ByVal totalDays As Long, _
    ByVal constraintDate As Variant, _
    ByVal borderMode As String)

    Dim targetCell As Range

    Set targetCell = GetConstraintTargetCell_Gantt(ws, ganttRow, projectStart, totalDays, constraintDate)
    If targetCell Is Nothing Then Exit Sub

    Select Case UCase$(borderMode)
        Case "LEFT"
            ApplyConstraintOneBorder_Gantt targetCell.Borders(xlEdgeLeft)
        Case "RIGHT"
            ApplyConstraintOneBorder_Gantt targetCell.Borders(xlEdgeRight)
        Case "FULL"
            ApplyConstraintOneBorder_Gantt targetCell.Borders(xlEdgeLeft)
            ApplyConstraintOneBorder_Gantt targetCell.Borders(xlEdgeRight)
            ApplyConstraintOneBorder_Gantt targetCell.Borders(xlEdgeTop)
            ApplyConstraintOneBorder_Gantt targetCell.Borders(xlEdgeBottom)
    End Select

End Sub

Private Sub ApplyConstraintOneBorder_Gantt(ByVal borderObj As Border)

    With borderObj
        .LineStyle = xlContinuous
        .Weight = xlThick
        .Color = RGB(220, 0, 0)
    End With

End Sub

Private Sub DrawConstraintBigCrossAtCellEdge_Gantt( _
    ByVal ws As Worksheet, _
    ByVal ganttRow As Long, _
    ByVal projectStart As Variant, _
    ByVal totalDays As Long, _
    ByVal constraintDate As Variant, _
    ByVal edgeSide As String, _
    ByVal shapeSuffix As String, _
    ByVal token As String)

    Dim targetCell As Range
    Dim x As Double
    Dim y As Double
    Dim sizeVal As Double
    Dim offsetVal As Double

    Set targetCell = GetConstraintTargetCell_Gantt(ws, ganttRow, projectStart, totalDays, constraintDate)
    If targetCell Is Nothing Then Exit Sub

    sizeVal = 8
    offsetVal = 5

    If UCase$(edgeSide) = "RIGHT" Then
        x = targetCell.Left + targetCell.Width + offsetVal
    Else
        x = targetCell.Left - offsetVal
    End If

    y = GetGanttRowTop(ws, ganttRow) + (GetGanttRowHeight(ws, ganttRow) / 2)

    DrawConstraintCross_Gantt ws, x, y, sizeVal, "CSTR_" & shapeSuffix & "_" & token & "_", 1

End Sub

Private Sub DrawConstraintSmallCrossesAtCellEdge_Gantt( _
    ByVal ws As Worksheet, _
    ByVal ganttRow As Long, _
    ByVal projectStart As Variant, _
    ByVal totalDays As Long, _
    ByVal constraintDate As Variant, _
    ByVal wallEdge As String, _
    ByVal crossSide As String, _
    ByVal shapeSuffix As String, _
    ByVal token As String)

    Dim targetCell As Range
    Dim wallX As Double
    Dim x As Double
    Dim yTop As Double
    Dim yBottom As Double
    Dim sizeVal As Double
    Dim offsetVal As Double
    Dim rowTop As Double
    Dim rowHeight As Double
    Dim verticalShift As Double
    Dim minY As Double
    Dim maxY As Double

    Set targetCell = GetConstraintTargetCell_Gantt(ws, ganttRow, projectStart, totalDays, constraintDate)
    If targetCell Is Nothing Then Exit Sub

    If UCase$(wallEdge) = "RIGHT" Then
        wallX = targetCell.Left + targetCell.Width
    Else
        wallX = targetCell.Left
    End If

    sizeVal = 5
    offsetVal = 4

    If UCase$(crossSide) = "RIGHT" Then
        x = wallX + offsetVal
    Else
        x = wallX - offsetVal
    End If

    rowTop = GetGanttRowTop(ws, ganttRow)
    rowHeight = GetGanttRowHeight(ws, ganttRow)
    verticalShift = rowHeight * 0.1
    minY = rowTop + (sizeVal / 2)
    maxY = rowTop + rowHeight - (sizeVal / 2)

    yTop = rowTop + (sizeVal / 2) + 0.5 - verticalShift
    If yTop < minY Then yTop = minY

    yBottom = rowTop + rowHeight - (sizeVal / 2) - 0.5 + verticalShift
    If yBottom > maxY Then yBottom = maxY

    DrawConstraintCross_Gantt ws, x, yTop, sizeVal, "CSTR_" & shapeSuffix & "_" & token & "_", 1
    DrawConstraintCross_Gantt ws, x, yBottom, sizeVal, "CSTR_" & shapeSuffix & "_" & token & "_", 3

End Sub

Private Sub DrawConstraintCross_Gantt( _
    ByVal ws As Worksheet, _
    ByVal centerX As Double, _
    ByVal centerY As Double, _
    ByVal sizeVal As Double, _
    ByVal shapeBase As String, _
    ByVal startIndex As Long)

    Dim shp As Shape

    Set shp = ws.Shapes.AddShape(msoShapeMathMultiply, centerX - (sizeVal / 2), centerY - (sizeVal / 2), sizeVal, sizeVal)
    shp.Name = shapeBase & CStr(startIndex)
    ApplyGanttRenderShapePlacement shp

    With shp
        .Fill.Visible = msoTrue
        .Fill.ForeColor.RGB = RGB(220, 0, 0)
        .Fill.Transparency = 0
        .Line.Visible = msoFalse
    End With

    shp.ZOrder msoBringToFront

End Sub

Private Sub DrawTodayLine(ByVal ws As Worksheet, ByVal projectStart As Variant, ByVal totalDays As Long, ByVal rowCount As Long)

    Dim todayVal As Date
    Dim renderToday As Date
    Dim projectFinish As Date

    Dim targetCol As Long
    Dim x As Double
    Dim yTop As Double
    Dim yBottom As Double
    Dim shp As Shape

    If ws Is Nothing Then Exit Sub
    If Not HasValue(projectStart) Then Exit Sub
    If totalDays < 1 Then Exit Sub
    If rowCount < 1 Then Exit Sub

    todayVal = Date

    If IsWeekScaleMode() Then
        renderToday = GetIsoWeekMonday(todayVal)
        projectFinish = CDate(projectStart) + ((totalDays - 1) * 7) + 6
    Else
        renderToday = todayVal
        projectFinish = CDate(CDbl(projectStart) + totalDays - 1)
    End If

    If renderToday < CDate(projectStart) Or renderToday > projectFinish Then Exit Sub

    targetCol = TimelineColumnFromHeaderDate_Exact(ws, projectStart, renderToday)
    If targetCol < FIRST_TIMELINE_COL Then Exit Sub

    x = ws.Cells(HEADER_ROW_2, targetCol).Left + (ws.Cells(HEADER_ROW_2, targetCol).Width / 2)

    yTop = ws.Cells(HEADER_ROW_1, FIRST_TIMELINE_COL).Top
    yBottom = GetGanttRowTop(ws, FIRST_TASK_ROW + rowCount - 1) + _
              GetGanttRowHeight(ws, FIRST_TASK_ROW + rowCount - 1)

    If x <= 0 Then Exit Sub
    If yBottom <= yTop Then Exit Sub

    Set shp = ws.Shapes.AddLine(x, yTop, x, yBottom)
    shp.Name = "TODAY_LINE"
    ApplyGanttRenderLinePlacement shp

    With shp.Line
        .ForeColor.RGB = RGB(255, 192, 0)
        .Weight = 4.5 '3x thicker
        .DashStyle = msoLineDash
    End With

    shp.ZOrder msoBringToFront

End Sub



Private Sub FinalizeGanttSheet_DisplayOnly(ByVal wsGantt As Worksheet, ByVal totalDays As Long, ByVal rowCount As Long)

    Dim lastCol As Long
    Dim lastRow As Long

    lastCol = FIRST_TIMELINE_COL + totalDays - 1
    lastRow = FIRST_TASK_ROW + rowCount - 1

    If lastCol < FIRST_TIMELINE_COL Then lastCol = FIRST_TIMELINE_COL
    If lastRow < FIRST_TASK_ROW Then lastRow = FIRST_TASK_ROW

    wsGantt.Range(wsGantt.Cells(HEADER_ROW_2, FIRST_TIMELINE_COL), wsGantt.Cells(lastRow, lastCol)).Borders(xlInsideVertical).LineStyle = xlDot
    wsGantt.Range(wsGantt.Cells(FIRST_TASK_ROW, FIRST_TIMELINE_COL), wsGantt.Cells(lastRow, lastCol)).Borders(xlInsideHorizontal).LineStyle = xlDot

End Sub

Private Function GetTimelineTargetCol( _
    ByVal projectStart As Variant, _
    ByVal taskDate As Variant) As Long

    Dim slotOffset As Long

    If Not HasValue(projectStart) Then Exit Function
    If Not HasValue(taskDate) Then Exit Function

    slotOffset = GetTimelineSlotIndex(projectStart, taskDate)

    If slotOffset < 0 Then
        GetTimelineTargetCol = FIRST_TIMELINE_COL
    Else
        GetTimelineTargetCol = FIRST_TIMELINE_COL + slotOffset
    End If

End Function

Private Function TimelineLeft( _
    ByVal ws As Worksheet, _
    ByVal projectStart As Variant, _
    ByVal taskStart As Variant) As Double

    Dim targetCol As Long

    targetCol = TimelineColumnFromHeaderDate_Exact(ws, projectStart, taskStart)
    If targetCol < FIRST_TIMELINE_COL Then Exit Function

    TimelineLeft = ws.Cells(HEADER_ROW_2, targetCol).Left

End Function

Private Function TimelineWidth( _
    ByVal ws As Worksheet, _
    ByVal projectStart As Variant, _
    ByVal taskStart As Variant, _
    ByVal taskFinish As Variant) As Double

    Dim leftPos As Double
    Dim rightPos As Double

    If Not HasValue(projectStart) Then Exit Function
    If Not HasValue(taskStart) Then Exit Function
    If Not HasValue(taskFinish) Then Exit Function

    leftPos = TimelineLeft(ws, projectStart, taskStart)
    rightPos = TimelineRightAfterFinish(ws, projectStart, taskFinish)

    If rightPos <= leftPos Then
        TimelineWidth = 0
    Else
        TimelineWidth = rightPos - leftPos
    End If

End Function


Private Function WBSLevel(ByVal wbs As String) As Long

    If Trim$(wbs) = "" Then
        WBSLevel = 1
    Else
        WBSLevel = UBound(Split(wbs, ".")) + 1
    End If

End Function

Private Sub DrawDependencyLinks( _
    ByVal wsGantt As Worksheet, _
    ByVal mapWBS As Object, _
    ByRef dataArr As Variant, _
    ByVal hasChildren As Object, _
    ByVal rowById As Object, _
    ByVal projectStart As Variant, _
    ByVal totalDays As Long, _
    ByVal baseById As Object, _
    ByVal testById As Object, _
    ByVal isTestMode As Boolean)

    Dim succId As Variant
    Dim linkItem As Variant
    Dim predId As String
    Dim shapePrefix As String
    Dim linkIndex As Long

    If IsWeekScaleMode() Then Exit Sub

    On Error GoTo SafeExit

    EnsureExpandedLinksCacheFromCalc
    If Not HasExpandedLinksAvailable() Then Exit Sub

    For Each succId In gExpandedLinks.Keys

        linkIndex = 0

        For Each linkItem In gExpandedLinks(CStr(succId))

            predId = Trim$(CStr(linkItem("PredID")))

            If predId <> "" Then
                If rowById.Exists(predId) And rowById.Exists(CStr(succId)) Then

                    '==================================================
                    ' NOUVELLE RÈGLE :
                    ' on accepte aussi l'affichage des liens parent -> parent
                    ' si le lien existe dans CALC, on le dessine
                    '==================================================
                    linkIndex = linkIndex + 1
                    shapePrefix = "DEP_" & predId & "_" & CStr(succId) & "_" & CStr(linkIndex)

                    DrawSingleDependencyLink _
                        wsGantt, mapWBS, dataArr, hasChildren, rowById, _
                        projectStart, totalDays, _
                        predId, CStr(succId), _
                        baseById, testById, isTestMode, _
                        shapePrefix, _
                        GetLinkTypeFromItem(linkItem), _
                        GetLinkLagFromItem(linkItem)

                End If
            End If

NextLink:
        Next linkItem

    Next succId

SafeExit:
End Sub


Private Function GetLinkTypeFromItem(ByVal linkItem As Variant) As String

    On Error GoTo SafeExit

    If IsObject(linkItem) Then
        If linkItem.Exists("LinkType") Then
            GetLinkTypeFromItem = UCase$(Trim$(CStr(linkItem("LinkType"))))
            If GetLinkTypeFromItem = "" Then GetLinkTypeFromItem = "FS"
            Exit Function
        End If
    End If

SafeExit:
    GetLinkTypeFromItem = "FS"

End Function

Private Function GetLinkLagFromItem(ByVal linkItem As Variant) As Double

    On Error GoTo SafeExit

    If IsObject(linkItem) Then
        If linkItem.Exists("Lag") Then
            If IsNumeric(linkItem("Lag")) Then
                GetLinkLagFromItem = CDbl(linkItem("Lag"))
                Exit Function
            End If
        End If
    End If

SafeExit:
    GetLinkLagFromItem = 0#

End Function

Private Sub DrawSingleDependencyLink( _
    ByVal wsGantt As Worksheet, _
    ByVal mapWBS As Object, _
    ByRef dataArr As Variant, _
    ByVal hasChildren As Object, _
    ByVal rowById As Object, _
    ByVal projectStart As Variant, _
    ByVal totalDays As Long, _
    ByVal predId As String, _
    ByVal succId As String, _
    ByVal baseById As Object, _
    ByVal testById As Object, _
    ByVal isTestMode As Boolean, _
    ByVal shapePrefix As String, _
    ByVal linkType As String, _
    ByVal linkLag As Double)

    Dim predRow As Long
    Dim succRow As Long
    Dim predDataRow As Long
    Dim succDataRow As Long

    Dim predX As Double
    Dim predY As Double
    Dim succX As Double
    Dim succY As Double

    Dim succTopX As Double
    Dim succTopY As Double
    Dim succMidLeftX As Double
    Dim succMidLeftY As Double

    Dim predAnchorType As String
    Dim succAnchorType As String
    Dim predDate As Variant
    Dim succDate As Variant
    Dim gapDays As Long
    Dim useMidLeftEntry As Boolean

    If Not rowById.Exists(predId) Then Exit Sub
    If Not rowById.Exists(succId) Then Exit Sub

    predRow = CLng(rowById(predId))
    succRow = CLng(rowById(succId))

    predDataRow = predRow - FIRST_TASK_ROW + 1
    succDataRow = succRow - FIRST_TASK_ROW + 1

    If predDataRow < 1 Or predDataRow > UBound(dataArr, 1) Then Exit Sub
    If succDataRow < 1 Or succDataRow > UBound(dataArr, 1) Then Exit Sub

    linkType = UCase$(Trim$(linkType))
    If linkType = "" Then linkType = "FS"

    GetLinkAnchorTypes linkType, predAnchorType, succAnchorType

    predDate = GetLinkReferenceDate(predId, predAnchorType, baseById, testById, isTestMode)
    succDate = GetLinkReferenceDate(succId, succAnchorType, baseById, testById, isTestMode)

    If Not HasValue(predDate) Then Exit Sub
    If Not HasValue(succDate) Then Exit Sub

    GetTaskAnchorPointByType wsGantt, mapWBS, dataArr, hasChildren, projectStart, totalDays, predDataRow, _
        predAnchorType, predX, predY, baseById, testById, isTestMode

    Select Case linkType

        Case "SS"
            GetTaskAnchorPointByType wsGantt, mapWBS, dataArr, hasChildren, projectStart, totalDays, succDataRow, _
                succAnchorType, succX, succY, baseById, testById, isTestMode

            gapDays = CLng(CDbl(succDate) - CDbl(predDate) - linkLag)
            RouteDependencyLink_SS wsGantt, shapePrefix, predX, predY, succX, succY, gapDays

        Case "FF"
            GetTaskAnchorPointByType wsGantt, mapWBS, dataArr, hasChildren, projectStart, totalDays, succDataRow, _
                succAnchorType, succX, succY, baseById, testById, isTestMode

            gapDays = CLng(CDbl(succDate) - CDbl(predDate) - linkLag)
            RouteDependencyLink_FF wsGantt, shapePrefix, predX, predY, succX, succY, gapDays

        Case Else   ' FS
            gapDays = CLng(CDbl(succDate) - CDbl(predDate) - 1 - linkLag)

            ' On calcule les 2 points candidats côté successeur
            GetTaskTopEntryPoint wsGantt, mapWBS, dataArr, projectStart, totalDays, succDataRow, _
                succTopX, succTopY, baseById, testById, isTestMode

            GetTaskStartMidEntryPoint wsGantt, mapWBS, dataArr, hasChildren, projectStart, totalDays, succDataRow, _
                succMidLeftX, succMidLeftY, baseById, testById, isTestMode

            ' Règle corrigée :
            ' on ne décide PAS avec gapDays=0/1
            ' on décide avec la place horizontale réelle entre pred et l'entrée gauche du successeur
            useMidLeftEntry = HasRoomForFsMidLeftEntry(predX, succMidLeftX, wsGantt.Cells(HEADER_ROW_2, FIRST_TIMELINE_COL).Width)

            If gapDays < 0 Then
                succX = succMidLeftX
                succY = succMidLeftY
                RouteDependencyLink_FS_Negative wsGantt, shapePrefix, predX, predY, succX, succY, gapDays

            ElseIf useMidLeftEntry Then
                succX = succMidLeftX
                succY = succMidLeftY
                RouteDependencyLink_FS_Normal wsGantt, shapePrefix, predX, predY, succX, succY, gapDays

            Else
                succX = succTopX
                succY = succTopY
                RouteDependencyLink_FS_SameDay wsGantt, shapePrefix, predX, predY, succX, succY
            End If

    End Select

End Sub

Private Function HasRoomForFsMidLeftEntry( _
    ByVal predX As Double, _
    ByVal succMidLeftX As Double, _
    ByVal cellWidth As Double) As Boolean

    Dim minNeeded As Double

    ' Il faut une vraie place visuelle pour arriver par la gauche.
    ' gapDays = 0 peut quand même avoir assez de place à l’écran.
    minNeeded = WorksheetFunction.Max(10, cellWidth * 0.55)

    HasRoomForFsMidLeftEntry = ((succMidLeftX - predX) >= minNeeded)

End Function

Private Sub RouteDependencyLink_FS_SameDay( _
    ByVal wsGantt As Worksheet, _
    ByVal shapePrefix As String, _
    ByVal predX As Double, _
    ByVal predY As Double, _
    ByVal succX As Double, _
    ByVal succY As Double)

    DrawLinkSegment wsGantt, shapePrefix & "_1", predX, predY, succX, predY, False
    DrawLinkSegment wsGantt, shapePrefix & "_2", succX, predY, succX, succY, True

End Sub

Private Sub GetTaskStartMidEntryPoint( _
    ByVal ws As Worksheet, _
    ByVal mapWBS As Object, _
    ByRef dataArr As Variant, _
    ByVal hasChildren As Object, _
    ByVal projectStart As Variant, _
    ByVal totalDays As Long, _
    ByVal dataRow As Long, _
    ByRef xOut As Double, _
    ByRef yOut As Double, _
    ByVal baseById As Object, _
    ByVal testById As Object, _
    ByVal isTestMode As Boolean)

    ' Entrée milieu gauche = vrai point milieu côté gauche
    GetTaskAnchorPointBySide ws, mapWBS, dataArr, hasChildren, projectStart, totalDays, dataRow, _
        "LEFT", xOut, yOut, baseById, testById, isTestMode

End Sub

Private Sub RouteDependencyLink_FS_Normal( _
    ByVal wsGantt As Worksheet, _
    ByVal shapePrefix As String, _
    ByVal predX As Double, _
    ByVal predY As Double, _
    ByVal succX As Double, _
    ByVal succY As Double, _
    ByVal gapDays As Long)

    Dim endX As Double
    Dim directEnough As Boolean
    Dim routeAbove As Boolean
    Dim laneY As Double
    Dim bendX As Double
    Dim entryGap As Double
    Dim finalX As Double
    Dim cellWidth As Double

    cellWidth = wsGantt.Cells(HEADER_ROW_2, FIRST_TIMELINE_COL).Width

    ' IMPORTANT : plus aucun cas spécial ici basé sur gapDays = 0.
    ' Le choix top / milieu-gauche a déjà été fait en amont.

    If gapDays <= 1 Then
        entryGap = 8
    Else
        entryGap = 8 + (cellWidth / 2)
    End If

    endX = succX - entryGap
    If endX <= predX + 4 Then endX = succX - 4
    If endX <= predX + 2 Then endX = succX

    directEnough = (endX - predX >= LINK_STUB * 2)

    If directEnough Then
        If gapDays > 1 Then
            bendX = predX + LINK_STUB + (cellWidth / 3)
        Else
            bendX = predX + LINK_STUB
        End If

        If bendX > endX - LINK_STUB Then
            bendX = predX + ((endX - predX) / 2)
        End If

        finalX = succX - 2

        If finalX <= endX + 2 Then
            DrawLinkSegment wsGantt, shapePrefix & "_1", predX, predY, bendX, predY, False
            DrawLinkSegment wsGantt, shapePrefix & "_2", bendX, predY, bendX, succY, False
            DrawLinkSegment wsGantt, shapePrefix & "_3", bendX, succY, succX, succY, True
        Else
            DrawLinkSegment wsGantt, shapePrefix & "_1", predX, predY, bendX, predY, False
            DrawLinkSegment wsGantt, shapePrefix & "_2", bendX, predY, bendX, succY, False
            DrawLinkSegment wsGantt, shapePrefix & "_3", bendX, succY, endX, succY, False
            DrawLinkSegment wsGantt, shapePrefix & "_4", endX, succY, succX, succY, True
        End If

        Exit Sub
    End If

    routeAbove = (succY <= predY)

    If routeAbove Then
        laneY = WorksheetFunction.Min(predY, succY) - LINK_MIN_CHANNEL_GAP
    Else
        laneY = WorksheetFunction.Max(predY, succY) + LINK_MIN_CHANNEL_GAP
    End If

    If gapDays > 1 Then
        bendX = predX + LINK_STUB + (cellWidth / 3)
    Else
        bendX = predX + LINK_STUB
    End If

    If bendX >= succX - 6 Then
        bendX = predX + ((succX - predX) / 2)
    End If

    finalX = succX - 2

    If finalX <= endX + 2 Then
        DrawLinkSegment wsGantt, shapePrefix & "_1", predX, predY, bendX, predY, False
        DrawLinkSegment wsGantt, shapePrefix & "_2", bendX, predY, bendX, laneY, False
        DrawLinkSegment wsGantt, shapePrefix & "_3", bendX, laneY, succX - 2, laneY, False
        DrawLinkSegment wsGantt, shapePrefix & "_4", succX - 2, laneY, succX - 2, succY, False
        DrawLinkSegment wsGantt, shapePrefix & "_5", succX - 2, succY, succX, succY, True
    Else
        DrawLinkSegment wsGantt, shapePrefix & "_1", predX, predY, bendX, predY, False
        DrawLinkSegment wsGantt, shapePrefix & "_2", bendX, predY, bendX, laneY, False
        DrawLinkSegment wsGantt, shapePrefix & "_3", bendX, laneY, endX, laneY, False
        DrawLinkSegment wsGantt, shapePrefix & "_4", endX, laneY, endX, succY, False
        DrawLinkSegment wsGantt, shapePrefix & "_5", endX, succY, succX, succY, True
    End If

End Sub

Private Sub RouteDependencyLink_FS_Negative( _
    ByVal wsGantt As Worksheet, _
    ByVal shapePrefix As String, _
    ByVal predX As Double, _
    ByVal predY As Double, _
    ByVal succX As Double, _
    ByVal succY As Double, _
    ByVal gapDays As Long)

    Dim cellWidth As Double
    Dim laneY As Double
    Dim leftX As Double

    cellWidth = wsGantt.Cells(HEADER_ROW_2, FIRST_TIMELINE_COL).Width

    If succY <= predY Then
        laneY = WorksheetFunction.Min(predY, succY) - LINK_MIN_CHANNEL_GAP
    Else
        laneY = WorksheetFunction.Max(predY, succY) + LINK_MIN_CHANNEL_GAP
    End If

    leftX = WorksheetFunction.Min(predX, succX) - WorksheetFunction.Max(8, cellWidth / 2)
    leftX = leftX - WorksheetFunction.Max(6, Abs(gapDays) * (cellWidth / 2))

    DrawLinkSegment wsGantt, shapePrefix & "_1", predX, predY, predX, laneY, False
    DrawLinkSegment wsGantt, shapePrefix & "_2", predX, laneY, leftX, laneY, False
    DrawLinkSegment wsGantt, shapePrefix & "_3", leftX, laneY, leftX, succY, False
    DrawLinkSegment wsGantt, shapePrefix & "_4", leftX, succY, succX, succY, True

End Sub

Private Sub RouteDependencyLink_SS( _
    ByVal wsGantt As Worksheet, _
    ByVal shapePrefix As String, _
    ByVal predX As Double, _
    ByVal predY As Double, _
    ByVal succX As Double, _
    ByVal succY As Double, _
    ByVal gapDays As Long)

    Dim cellWidth As Double
    Dim busX As Double

    cellWidth = wsGantt.Cells(HEADER_ROW_2, FIRST_TIMELINE_COL).Width

    busX = WorksheetFunction.Min(predX, succX) - WorksheetFunction.Max(8, cellWidth / 2)

    If gapDays < 0 Then
        busX = busX - WorksheetFunction.Max(6, Abs(gapDays) * (cellWidth / 2))
    End If

    DrawLinkSegment wsGantt, shapePrefix & "_1", predX, predY, busX, predY, False
    DrawLinkSegment wsGantt, shapePrefix & "_2", busX, predY, busX, succY, False
    DrawLinkSegment wsGantt, shapePrefix & "_3", busX, succY, succX, succY, True

End Sub

Private Sub RouteDependencyLink_FF( _
    ByVal wsGantt As Worksheet, _
    ByVal shapePrefix As String, _
    ByVal predX As Double, _
    ByVal predY As Double, _
    ByVal succX As Double, _
    ByVal succY As Double, _
    ByVal gapDays As Long)

    Dim cellWidth As Double
    Dim busX As Double

    cellWidth = wsGantt.Cells(HEADER_ROW_2, FIRST_TIMELINE_COL).Width

    busX = WorksheetFunction.Min(predX, succX) - WorksheetFunction.Max(8, cellWidth / 2)

    If gapDays < 0 Then
        busX = busX - WorksheetFunction.Max(6, Abs(gapDays) * (cellWidth / 2))
    End If

    DrawLinkSegment wsGantt, shapePrefix & "_1", predX, predY, busX, predY, False
    DrawLinkSegment wsGantt, shapePrefix & "_2", busX, predY, busX, succY, False
    DrawLinkSegment wsGantt, shapePrefix & "_3", busX, succY, succX, succY, True

End Sub

Private Sub GetTaskTopEntryPoint( _
    ByVal ws As Worksheet, _
    ByVal mapWBS As Object, _
    ByRef dataArr As Variant, _
    ByVal projectStart As Variant, _
    ByVal totalDays As Long, _
    ByVal dataRow As Long, _
    ByRef xOut As Double, _
    ByRef yOut As Double, _
    ByVal baseById As Object, _
    ByVal testById As Object, _
    ByVal isTestMode As Boolean)

    Dim ganttRow As Long
    Dim idVal As String
    Dim startVal As Variant
    Dim finishVal As Variant
    Dim durationVal As Double
    Dim timelineLeftBound As Double
    Dim timelineRightBound As Double
    Dim topEntryOffset As Double

    ganttRow = FIRST_TASK_ROW + dataRow - 1
    idVal = Trim$(CStr(dataArr(dataRow, mapWBS("ID"))))

    startVal = GetRenderStartForCurrentScale(GanttLive_GetDisplayStart(idVal, baseById, testById, isTestMode))
    finishVal = GetRenderFinishForCurrentScale(GanttLive_GetDisplayFinish(idVal, baseById, testById, isTestMode))

    If Not HasValue(startVal) Or Not HasValue(finishVal) Then Exit Sub

    durationVal = CDbl(finishVal) - CDbl(startVal) + 1

    timelineLeftBound = ws.Cells(HEADER_ROW_2, FIRST_TIMELINE_COL).Left + LINK_EDGE_PADDING
    timelineRightBound = ws.Cells(HEADER_ROW_2, FIRST_TIMELINE_COL + totalDays - 1).Left + _
                         ws.Cells(HEADER_ROW_2, FIRST_TIMELINE_COL + totalDays - 1).Width - LINK_EDGE_PADDING

    If durationVal <= 1 Then
        xOut = GetTaskMidX(ws, projectStart, startVal)
        yOut = GetGanttRowTop(ws, ganttRow) + 3
    Else
        xOut = TimelineLeft(ws, projectStart, startVal) + 4
        yOut = GetGanttBarTop(ws, ganttRow)
    End If

    topEntryOffset = ws.Cells(HEADER_ROW_2, FIRST_TIMELINE_COL).Width * 0.15
    xOut = xOut + topEntryOffset

    If xOut < timelineLeftBound Then xOut = timelineLeftBound
    If xOut > timelineRightBound Then xOut = timelineRightBound

End Sub




Private Sub FormatDependencyLine(ByVal shp As Shape, ByVal withArrow As Boolean)

    With shp.Line
        .ForeColor.RGB = RGB(120, 120, 120)
        .Weight = 1
        .DashStyle = msoLineSolid
        If withArrow Then
            .EndArrowheadStyle = msoArrowheadTriangle
        End If
    End With

End Sub


Private Sub DrawLinkSegment( _
    ByVal ws As Worksheet, _
    ByVal shapeName As String, _
    ByVal x1 As Double, _
    ByVal y1 As Double, _
    ByVal x2 As Double, _
    ByVal y2 As Double, _
    ByVal withArrow As Boolean)

    Dim shp As Shape

    If Abs(x2 - x1) < 0.1 And Abs(y2 - y1) < 0.1 Then Exit Sub

    Set shp = ws.Shapes.AddLine(x1, y1, x2, y2)
    shp.Name = shapeName
    ApplyGanttRenderLinePlacement shp
    FormatDependencyLine shp, withArrow

End Sub



Private Function BuildCalcDrivingLogicMap() As Object

    Dim wsCalc As Worksheet
    Dim tblCalc As ListObject
    Dim d As Object
    Dim mapCalc As Object
    Dim arr As Variant
    Dim r As Long
    Dim idVal As String

    Set wsCalc = ThisWorkbook.Worksheets(CALC_SHEET)
    Set tblCalc = wsCalc.ListObjects(CALC_TABLE)
    Set d = CreateObject("Scripting.Dictionary")

    If tblCalc.DataBodyRange Is Nothing Then
        Set BuildCalcDrivingLogicMap = d
        Exit Function
    End If

    Set mapCalc = BuildWBSColumnMap(tblCalc)
    arr = tblCalc.DataBodyRange.value

    For r = 1 To UBound(arr, 1)
        idVal = Trim$(CStr(arr(r, mapCalc("ID"))))
        If idVal <> "" Then
            d(idVal) = UCase$(Trim$(CStr(arr(r, mapCalc("Driving Logic")))))
        End If
    Next r

    Set BuildCalcDrivingLogicMap = d

End Function

Private Function IsLeafById(ByVal idVal As String, ByVal calcDrivingMap As Object) As Boolean

    If idVal = "" Then
        IsLeafById = False
    ElseIf calcDrivingMap.Exists(idVal) Then
        IsLeafById = (calcDrivingMap(idVal) <> "SUMMARY")
    Else
        IsLeafById = False
    End If

End Function

Private Sub FreezeGanttAfterFinish(ByVal ws As Worksheet, ByVal rowCount As Long)

    ws.Activate
    ActiveWindow.FreezePanes = False
    ws.Cells(FIRST_TASK_ROW, COL_TEST_START).Select
    ActiveWindow.FreezePanes = True

End Sub

Private Function GetLastGanttRow(ByVal ws As Worksheet) As Long

    Dim lastCell As Range

    Set lastCell = ws.Columns(COL_WBS).Find( _
        What:="*", _
        After:=ws.Cells(1, COL_WBS), _
        LookIn:=xlValues, _
        LookAt:=xlPart, _
        SearchOrder:=xlByRows, _
        SearchDirection:=xlPrevious, _
        MatchCase:=False)

    If lastCell Is Nothing Then
        GetLastGanttRow = FIRST_TASK_ROW - 1
    ElseIf lastCell.Row < FIRST_TASK_ROW Then
        GetLastGanttRow = FIRST_TASK_ROW - 1
    Else
        GetLastGanttRow = lastCell.Row
    End If

End Function

Private Function IsAllowedTestColumn(ByVal colNum As Long) As Boolean

    IsAllowedTestColumn = ( _
        colNum = COL_TEST_START Or _
        colNum = COL_TEST_FINISH Or _
        colNum = COL_TEST_PROGRESS)

End Function

Private Function BuildWbsToIdMapFromWBS() As Object

    Dim wsWBS As Worksheet
    Dim tblWBS As ListObject
    Dim mapWBS As Object
    Dim arr As Variant
    Dim d As Object
    Dim r As Long
    Dim wbsVal As String
    Dim idVal As String

    Set wsWBS = ThisWorkbook.Worksheets(WBS_SHEET)
    Set tblWBS = wsWBS.ListObjects(WBS_TABLE)
    Set mapWBS = BuildWBSColumnMap(tblWBS)
    Set d = CreateObject("Scripting.Dictionary")

    If tblWBS.DataBodyRange Is Nothing Then
        Set BuildWbsToIdMapFromWBS = d
        Exit Function
    End If

    arr = tblWBS.DataBodyRange.value

    For r = 1 To UBound(arr, 1)
        wbsVal = NormalizeWBS(CStr(arr(r, mapWBS("WBS"))))
        idVal = Trim$(CStr(arr(r, mapWBS("ID"))))

        If wbsVal <> "" And idVal <> "" Then
            d(wbsVal) = idVal
        End If
    Next r

    Set BuildWbsToIdMapFromWBS = d

End Function

Public Function IsEditableTestCell(ByVal ws As Worksheet, ByVal targetCell As Range) As Boolean

    Dim rowNum As Long
    Dim colNum As Long
    Dim lastRow As Long
    Dim wbsVal As String
    Dim wbsToId As Object
    Dim calcDrivingMap As Object
    Dim idVal As String

    IsEditableTestCell = False

    rowNum = targetCell.Row
    colNum = targetCell.Column
    lastRow = GetLastGanttRow(ws)

    If rowNum < FIRST_TASK_ROW Or rowNum > lastRow Then Exit Function
    If Not IsAllowedTestColumn(colNum) Then Exit Function

    wbsVal = NormalizeWBS(CStr(ws.Cells(rowNum, COL_WBS).value))
    If wbsVal = "" Then Exit Function

    Set wbsToId = BuildWbsToIdMapFromWBS()
    If Not wbsToId.Exists(wbsVal) Then Exit Function

    idVal = CStr(wbsToId(wbsVal))
    Set calcDrivingMap = BuildCalcDrivingLogicMap()

    If IsLeafById(idVal, calcDrivingMap) Then
        IsEditableTestCell = True
    End If

End Function

Public Sub NormalizeGanttTestPercentCell(ByVal cell As Range)

    Dim rawVal As Variant
    Dim pctVal As Double

    rawVal = cell.value

    If Trim$(CStr(rawVal)) = "" Then Exit Sub
    If Not IsNumeric(rawVal) Then Exit Sub

    pctVal = CDbl(rawVal)

    If pctVal < 0 Then Exit Sub

    If pctVal <= 1 Then
        cell.value = pctVal
    Else
        cell.value = pctVal / 100#
    End If

    cell.NumberFormat = "0%"

End Sub

Public Sub Ensure_Gantt_Test_Buttons()

    Dim ws As Worksheet
    Dim scenarioCell As Range
    Dim testCell As Range
    Dim lockCell As Range

    Set ws = ThisWorkbook.Worksheets(GANTT_SHEET)

    With ws.Range("A1:H3")
        .Interior.Color = RGB(255, 255, 255)
        .Borders.LineStyle = xlNone
    End With

    EnsureGanttViewInitialized

    Set scenarioCell = ws.Range("B1")
    Set testCell = ws.Range("C1")
    Set lockCell = ws.Range("D1")

    CreateOrUpdateGanttButton ws, BTN_SCENARIO_NAME, BTN_SCENARIO_CAPTION, "Run_Gantt_Scenario_Engine", scenarioCell, 0.34, True
    CreateOrUpdateGanttButton ws, BTN_TEST_NAME, BTN_TEST_CAPTION, "Run_Gantt_Test_Engine", testCell, 1, False
    CreateOrUpdateGanttButton ws, BTN_LOCK_NAME, BTN_LOCK_CAPTION, "Run_Gantt_Lock_Changes", lockCell, 1, False

    DeleteShapeIfExists ws, BTN_VIEW_BG_NAME
    DeleteShapeIfExists ws, BTN_VIEW_KNOB_NAME
    DeleteShapeIfExists ws, BTN_VIEW_LEFT_NAME
    DeleteShapeIfExists ws, BTN_VIEW_RIGHT_NAME

    DeleteShapeIfExists ws, BTN_CP_BG_NAME
    DeleteShapeIfExists ws, BTN_CP_KNOB_NAME
    DeleteShapeIfExists ws, BTN_CP_LEFT_NAME
    DeleteShapeIfExists ws, BTN_CP_RIGHT_NAME

    DeleteShapeIfExists ws, BTN_SCALE_BG_NAME
    DeleteShapeIfExists ws, BTN_SCALE_KNOB_NAME
    DeleteShapeIfExists ws, BTN_SCALE_LEFT_NAME
    DeleteShapeIfExists ws, BTN_SCALE_RIGHT_NAME

    DeleteShapeIfExists ws, BTN_CP_MULTI_BG_NAME
    DeleteShapeIfExists ws, BTN_CP_MULTI_KNOB_NAME
    DeleteShapeIfExists ws, BTN_CP_MULTI_LEFT_NAME

    DeleteShapeIfExists ws, BTN_CONSTRAINT_BG_NAME
    DeleteShapeIfExists ws, BTN_CONSTRAINT_KNOB_NAME
    DeleteShapeIfExists ws, BTN_CONSTRAINT_LEFT_NAME

    BuildFixedHeaderToggles ws

End Sub

Public Sub BuildFixedHeaderToggles(ByVal ws As Worksheet)

    Dim x As Double
    Dim y As Double

    Dim groupGap As Double
    Dim trackGap As Double

    Dim labelW1 As Double
    Dim labelW2 As Double
    Dim labelW3 As Double
    Dim labelW4 As Double
    Dim labelW5 As Double
    Dim analyticsTrackW As Double

    Dim trackW As Double
    Dim trackH As Double
    Dim knobSize As Double

    x = ws.Cells(TOGGLE_ROW_TOP, COL_WBS).Left + 5
    y = ws.Cells(TOGGLE_ROW_TOP, COL_WBS).Top + 3

    trackGap = 7
    groupGap = 25
    
    'Detail / Summary
    labelW1 = 70
    
    'None / Critical Path / Longest Path
    labelW2 = 130
    
    'Day / Week
    labelW3 = 45
    
    'Single / Multiple Project
    labelW4 = 95
    
    'Constraint
    labelW5 = 45
    
    trackW = 28
    analyticsTrackW = 46
    trackH = 12
    knobSize = 9

    CreateFixedHeaderToggle ws, _
        BTN_VIEW_LEFT_NAME, BTN_VIEW_BG_NAME, BTN_VIEW_KNOB_NAME, _
        x, y, labelW1, "Detail / Summary", _
        x + labelW1 + trackGap, y + 2, trackW, trackH, knobSize, _
        (gGanttViewMode = GANTT_VIEW_SUMMARY), "Toggle_Gantt_View"

    x = x + labelW1 + trackGap + trackW + groupGap

    CreateFixedHeaderToggle ws, _
        BTN_SCALE_LEFT_NAME, BTN_SCALE_BG_NAME, BTN_SCALE_KNOB_NAME, _
        x, y, labelW3, "Day / Week", _
        x + labelW3 + trackGap, y + 2, trackW, trackH, knobSize, _
        (gTimelineScaleMode = GANTT_SCALE_WEEK), "Toggle_Gantt_Scale"

    x = x + labelW3 + trackGap + trackW + groupGap

    CreateFixedHeaderToggle ws, _
        BTN_CONSTRAINT_LEFT_NAME, BTN_CONSTRAINT_BG_NAME, BTN_CONSTRAINT_KNOB_NAME, _
        x, y, labelW5, "Constraint", _
        x + labelW5 + trackGap, y + 2, trackW, trackH, knobSize, _
        gShowConstraints, "Toggle_Gantt_Constraints"

    x = ws.Cells(TOGGLE_ROW_BOTTOM, COL_WBS).Left + 5
    y = ws.Cells(TOGGLE_ROW_BOTTOM, COL_WBS).Top + 3

    CreateFixedHeaderTriToggle ws, _
        BTN_CP_LEFT_NAME, BTN_CP_BG_NAME, BTN_CP_KNOB_NAME, _
        x, y, labelW2, "None / Critical Path / Longest Path", _
        x + labelW2 + trackGap, y + 2, analyticsTrackW, trackH, knobSize, _
        gAnalyticsPathMode, "Toggle_Gantt_CriticalPath"

    x = x + labelW2 + trackGap + analyticsTrackW + groupGap

    CreateFixedHeaderToggle ws, _
        BTN_CP_MULTI_LEFT_NAME, BTN_CP_MULTI_BG_NAME, BTN_CP_MULTI_KNOB_NAME, _
        x, y, labelW4, "Single / Multiple Project", _
        x + labelW4 + trackGap, y + 2, trackW, trackH, knobSize, _
        IsCriticalPathMultiNetworkEnabled(), "Toggle_CriticalPathMode_FromGantt"


End Sub

Private Sub CreateFixedHeaderToggle( _
    ByVal ws As Worksheet, _
    ByVal labelName As String, _
    ByVal bgName As String, _
    ByVal knobName As String, _
    ByVal labelLeft As Double, _
    ByVal labelTop As Double, _
    ByVal labelWidth As Double, _
    ByVal labelText As String, _
    ByVal trackLeft As Double, _
    ByVal trackTop As Double, _
    ByVal trackWidth As Double, _
    ByVal trackHeight As Double, _
    ByVal knobSize As Double, _
    ByVal isOn As Boolean, _
    ByVal macroName As String)

    Dim shpLabel As Shape
    Dim shpBg As Shape
    Dim shpKnob As Shape
    Dim knobLeft As Double
    Dim knobTop As Double
    Dim onColor As Long

    onColor = RGB(68, 114, 196)

    DeleteShapeIfExists ws, labelName
    DeleteShapeIfExists ws, bgName
    DeleteShapeIfExists ws, knobName

    Set shpLabel = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, labelLeft, labelTop, labelWidth, trackHeight + 4)
    shpLabel.Name = labelName

    With shpLabel
        .Line.Visible = msoFalse
        .Fill.Visible = msoFalse
        .OnAction = ""
        .Placement = xlMoveAndSize
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
        .TextFrame2.MarginLeft = 0
        .TextFrame2.MarginRight = 0
        .TextFrame2.MarginTop = 0
        .TextFrame2.MarginBottom = 0
        .TextFrame2.TextRange.Text = labelText
        .TextFrame2.TextRange.Font.Size = 9.5
        .TextFrame2.TextRange.Font.Bold = msoTrue
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
        .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignLeft
    End With

    Set shpBg = ws.Shapes.AddShape(msoShapeRoundedRectangle, trackLeft, trackTop, trackWidth, trackHeight)
    shpBg.Name = bgName

    With shpBg
        .Adjustments.item(1) = 0.5
        .OnAction = macroName
        .Placement = xlMoveAndSize
        .Line.Weight = 1
        .Shadow.Visible = msoFalse
    End With

    knobTop = trackTop + ((trackHeight - knobSize) / 2)

    If isOn Then
        knobLeft = trackLeft + trackWidth - knobSize - 2
    Else
        knobLeft = trackLeft + 2
    End If

    Set shpKnob = ws.Shapes.AddShape(msoShapeOval, knobLeft, knobTop, knobSize, knobSize)
    shpKnob.Name = knobName

    With shpKnob
        .OnAction = macroName
        .Placement = xlMoveAndSize
        .Fill.ForeColor.RGB = RGB(255, 255, 255)
        .Line.ForeColor.RGB = RGB(150, 150, 150)
        .Line.Weight = 0.75
        .Shadow.Visible = msoFalse
    End With

    If isOn Then
        shpBg.Fill.ForeColor.RGB = onColor
        shpBg.Line.ForeColor.RGB = onColor
    Else
        shpBg.Fill.ForeColor.RGB = RGB(230, 230, 230)
        shpBg.Line.ForeColor.RGB = RGB(170, 170, 170)
    End If

End Sub

Private Sub CreateFixedHeaderTriToggle( _
    ByVal ws As Worksheet, _
    ByVal labelName As String, _
    ByVal bgName As String, _
    ByVal knobName As String, _
    ByVal labelLeft As Double, _
    ByVal labelTop As Double, _
    ByVal labelWidth As Double, _
    ByVal labelText As String, _
    ByVal trackLeft As Double, _
    ByVal trackTop As Double, _
    ByVal trackWidth As Double, _
    ByVal trackHeight As Double, _
    ByVal knobSize As Double, _
    ByVal modeValue As String, _
    ByVal macroName As String)

    Dim shpLabel As Shape
    Dim shpBg As Shape
    Dim shpKnob As Shape
    Dim knobLeft As Double
    Dim knobTop As Double
    Dim normalizedMode As String
    Dim onColor As Long

    onColor = RGB(68, 114, 196)
    normalizedMode = UCase$(Trim$(modeValue))
    If normalizedMode <> GANTT_ANALYTICS_PATH_CP And normalizedMode <> GANTT_ANALYTICS_PATH_LP Then
        normalizedMode = GANTT_ANALYTICS_PATH_NONE
    End If

    DeleteShapeIfExists ws, labelName
    DeleteShapeIfExists ws, bgName
    DeleteShapeIfExists ws, knobName

    Set shpLabel = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, labelLeft, labelTop, labelWidth, trackHeight + 4)
    shpLabel.Name = labelName

    With shpLabel
        .Line.Visible = msoFalse
        .Fill.Visible = msoFalse
        .OnAction = ""
        .Placement = xlMoveAndSize
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
        .TextFrame2.MarginLeft = 0
        .TextFrame2.MarginRight = 0
        .TextFrame2.MarginTop = 0
        .TextFrame2.MarginBottom = 0
        .TextFrame2.TextRange.Text = labelText
        .TextFrame2.TextRange.Font.Size = 9
        .TextFrame2.TextRange.Font.Bold = msoTrue
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
        .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignLeft
    End With

    Set shpBg = ws.Shapes.AddShape(msoShapeRoundedRectangle, trackLeft, trackTop, trackWidth, trackHeight)
    shpBg.Name = bgName

    With shpBg
        .Adjustments.item(1) = 0.5
        .OnAction = macroName
        .Placement = xlMoveAndSize
        .Line.Weight = 1
        .Shadow.Visible = msoFalse
    End With

    knobTop = trackTop + ((trackHeight - knobSize) / 2)

    Select Case normalizedMode
        Case GANTT_ANALYTICS_PATH_LP
            knobLeft = trackLeft + trackWidth - knobSize - 2
        Case GANTT_ANALYTICS_PATH_CP
            knobLeft = trackLeft + ((trackWidth - knobSize) / 2)
        Case Else
            knobLeft = trackLeft + 2
    End Select

    Set shpKnob = ws.Shapes.AddShape(msoShapeOval, knobLeft, knobTop, knobSize, knobSize)
    shpKnob.Name = knobName

    With shpKnob
        .OnAction = macroName
        .Placement = xlMoveAndSize
        .Fill.ForeColor.RGB = RGB(255, 255, 255)
        .Line.ForeColor.RGB = RGB(150, 150, 150)
        .Line.Weight = 0.75
        .Shadow.Visible = msoFalse
    End With

    If normalizedMode = GANTT_ANALYTICS_PATH_NONE Then
        shpBg.Fill.ForeColor.RGB = RGB(230, 230, 230)
        shpBg.Line.ForeColor.RGB = RGB(170, 170, 170)
    Else
        shpBg.Fill.ForeColor.RGB = onColor
        shpBg.Line.ForeColor.RGB = onColor
    End If

End Sub

Private Sub RefreshFixedHeaderToggleVisuals(ByVal ws As Worksheet)

    EnsureGanttViewInitialized

    RefreshFixedHeaderToggleVisual ws, BTN_VIEW_BG_NAME, BTN_VIEW_KNOB_NAME, (gGanttViewMode = GANTT_VIEW_SUMMARY)
    RefreshFixedHeaderToggleVisual ws, BTN_SCALE_BG_NAME, BTN_SCALE_KNOB_NAME, (gTimelineScaleMode = GANTT_SCALE_WEEK)
    RefreshFixedHeaderToggleVisual ws, BTN_CONSTRAINT_BG_NAME, BTN_CONSTRAINT_KNOB_NAME, gShowConstraints
    RefreshFixedHeaderTriToggleVisual ws, BTN_CP_BG_NAME, BTN_CP_KNOB_NAME, gAnalyticsPathMode
    RefreshFixedHeaderToggleVisual ws, BTN_CP_MULTI_BG_NAME, BTN_CP_MULTI_KNOB_NAME, IsCriticalPathMultiNetworkEnabled()

End Sub

Private Sub RefreshFixedHeaderToggleVisual( _
    ByVal ws As Worksheet, _
    ByVal bgName As String, _
    ByVal knobName As String, _
    ByVal isOn As Boolean)

    Dim shpBg As Shape
    Dim shpKnob As Shape
    Dim knobLeft As Double
    Dim onColor As Long

    onColor = RGB(68, 114, 196)

    On Error Resume Next
    Set shpBg = ws.Shapes(bgName)
    Set shpKnob = ws.Shapes(knobName)
    On Error GoTo 0

    If shpBg Is Nothing Then Exit Sub
    If shpKnob Is Nothing Then Exit Sub

    If isOn Then
        shpBg.Fill.ForeColor.RGB = onColor
        shpBg.Line.ForeColor.RGB = onColor
        knobLeft = shpBg.Left + shpBg.Width - shpKnob.Width - 2
    Else
        shpBg.Fill.ForeColor.RGB = RGB(230, 230, 230)
        shpBg.Line.ForeColor.RGB = RGB(170, 170, 170)
        knobLeft = shpBg.Left + 2
    End If

    shpBg.Line.Weight = 1
    shpBg.Shadow.Visible = msoFalse
    shpKnob.Fill.ForeColor.RGB = RGB(255, 255, 255)
    shpKnob.Line.ForeColor.RGB = RGB(150, 150, 150)
    shpKnob.Line.Weight = 0.75
    shpKnob.Shadow.Visible = msoFalse
    shpKnob.Left = knobLeft
    shpKnob.Top = shpBg.Top + ((shpBg.Height - shpKnob.Height) / 2)

End Sub

Private Sub RefreshFixedHeaderTriToggleVisual( _
    ByVal ws As Worksheet, _
    ByVal bgName As String, _
    ByVal knobName As String, _
    ByVal modeValue As String)

    Dim shpBg As Shape
    Dim shpKnob As Shape
    Dim normalizedMode As String
    Dim knobLeft As Double
    Dim onColor As Long

    onColor = RGB(68, 114, 196)
    normalizedMode = UCase$(Trim$(modeValue))
    If normalizedMode <> GANTT_ANALYTICS_PATH_CP And normalizedMode <> GANTT_ANALYTICS_PATH_LP Then
        normalizedMode = GANTT_ANALYTICS_PATH_NONE
    End If

    On Error Resume Next
    Set shpBg = ws.Shapes(bgName)
    Set shpKnob = ws.Shapes(knobName)
    On Error GoTo 0

    If shpBg Is Nothing Then Exit Sub
    If shpKnob Is Nothing Then Exit Sub

    Select Case normalizedMode
        Case GANTT_ANALYTICS_PATH_LP
            knobLeft = shpBg.Left + shpBg.Width - shpKnob.Width - 2
        Case GANTT_ANALYTICS_PATH_CP
            knobLeft = shpBg.Left + ((shpBg.Width - shpKnob.Width) / 2)
        Case Else
            knobLeft = shpBg.Left + 2
    End Select

    If normalizedMode = GANTT_ANALYTICS_PATH_NONE Then
        shpBg.Fill.ForeColor.RGB = RGB(230, 230, 230)
        shpBg.Line.ForeColor.RGB = RGB(170, 170, 170)
    Else
        shpBg.Fill.ForeColor.RGB = onColor
        shpBg.Line.ForeColor.RGB = onColor
    End If

    shpBg.Line.Weight = 1
    shpBg.Shadow.Visible = msoFalse
    shpKnob.Fill.ForeColor.RGB = RGB(255, 255, 255)
    shpKnob.Line.ForeColor.RGB = RGB(150, 150, 150)
    shpKnob.Line.Weight = 0.75
    shpKnob.Shadow.Visible = msoFalse
    shpKnob.Left = knobLeft
    shpKnob.Top = shpBg.Top + ((shpBg.Height - shpKnob.Height) / 2)

End Sub

Private Sub DeleteShapeIfExists(ByVal ws As Worksheet, ByVal shapeName As String)

    Dim shp As Shape

    On Error Resume Next
    Set shp = ws.Shapes(shapeName)
    On Error GoTo 0

    If Not shp Is Nothing Then shp.Delete

End Sub

Private Sub CreateOrUpdateGanttButton( _
    ByVal ws As Worksheet, _
    ByVal shpName As String, _
    ByVal captionText As String, _
    ByVal macroName As String, _
    ByVal topLeftCell As Range, _
    Optional ByVal widthRatio As Double = 1, _
    Optional ByVal alignRightInCell As Boolean = False)

    Dim shp As Shape
    Dim leftPos As Double
    Dim topPos As Double
    Dim widthVal As Double
    Dim heightVal As Double
    Dim availableWidth As Double

    Set shp = GetShapeIfExists(ws, shpName)

    If widthRatio <= 0 Then widthRatio = 1
    If widthRatio > 1 Then widthRatio = 1

    availableWidth = topLeftCell.Width - 4
    widthVal = availableWidth * widthRatio
    heightVal = topLeftCell.Height - 4

    If widthVal < 36 Then widthVal = 36

    If alignRightInCell Then
        leftPos = topLeftCell.Left + 2 + (availableWidth - widthVal)
    Else
        leftPos = topLeftCell.Left + 2
    End If

    topPos = topLeftCell.Top + 2

    If shp Is Nothing Then
        Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, leftPos, topPos, widthVal, heightVal)
        shp.Name = shpName
    End If

    shp.Left = leftPos
    shp.Top = topPos
    shp.Width = widthVal
    shp.Height = heightVal
    shp.OnAction = macroName

    With shp
        .Adjustments.item(1) = 0.25
        .Placement = xlMoveAndSize
        .Line.Visible = msoFalse
        .Shadow.Visible = msoTrue
        .Shadow.Blur = 6
        .Shadow.OffsetX = 0
        .Shadow.OffsetY = 1.5
        .Shadow.Transparency = 0.45
    End With

    With shp.TextFrame2
        .TextRange.Text = captionText
        .TextRange.Font.Size = 10
        .TextRange.Font.Bold = msoTrue
        .VerticalAnchor = msoAnchorMiddle
        .MarginLeft = 0
        .MarginRight = 0
        .MarginTop = 0
        .MarginBottom = 0
        .TextRange.ParagraphFormat.Alignment = msoAlignCenter
    End With

    Select Case shpName
        Case BTN_SCENARIO_NAME, BTN_LOCK_NAME
            shp.Fill.ForeColor.RGB = RGB(192, 0, 0)
            shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)

        Case BTN_TEST_NAME
            shp.Fill.ForeColor.RGB = RGB(112, 173, 71)
            shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    End Select

End Sub

Public Sub ApplyTestCellColoring( _
    ByVal ws As Worksheet, _
    ByVal rowIndex As Long, _
    ByVal isLeaf As Boolean, _
    ByVal hasActual As Boolean, _
    Optional ByVal isLoE As Boolean = False)

    Dim colStart As Long
    Dim colFinish As Long
    Dim colProgress As Long
    Dim rngDates As Range
    Dim rngProgress As Range

    colStart = 5
    colFinish = 6
    colProgress = 9

    Set rngDates = Union( _
        ws.Cells(rowIndex, colStart), _
        ws.Cells(rowIndex, colFinish))

    Set rngProgress = ws.Cells(rowIndex, colProgress)

    If Not isLeaf Then
        rngDates.Interior.Pattern = xlNone
        rngProgress.Interior.Pattern = xlNone
        Exit Sub
    End If

    rngDates.Interior.Pattern = xlSolid

    If hasActual Then
        rngDates.Interior.Color = RGB(255, 255, 204)
    Else
        rngDates.Interior.Color = RGB(255, 255, 153)
    End If

    If isLoE Then
        'LOE progress is calculated automatically from dates.
        'Do not invite manual Test % input visually.
        rngProgress.Interior.Pattern = xlNone
    Else
        rngProgress.Interior.Pattern = xlSolid

        If hasActual Then
            rngProgress.Interior.Color = RGB(255, 255, 204)
        Else
            rngProgress.Interior.Color = RGB(255, 255, 153)
        End If
    End If

End Sub

Public Sub Toggle_Gantt_View()

    Dim ws As Worksheet
    Dim oldScreenUpdating As Boolean

    EnsureGanttViewInitialized
    Set ws = ThisWorkbook.Worksheets(GANTT_SHEET)

    If gGanttViewMode = GANTT_VIEW_SUMMARY Then
        gGanttViewMode = GANTT_VIEW_DETAIL
    Else
        gGanttViewMode = GANTT_VIEW_SUMMARY
    End If

    oldScreenUpdating = Application.ScreenUpdating
    Application.ScreenUpdating = False
    On Error GoTo SafeExit

    ApplyCurrentGanttView ws
    RefreshFixedHeaderToggleVisuals ws

SafeExit:
    Application.ScreenUpdating = oldScreenUpdating

End Sub

Public Sub Toggle_Gantt_CriticalPath()

    EnsureGanttViewInitialized

    Select Case gAnalyticsPathMode
        Case GANTT_ANALYTICS_PATH_NONE
            gAnalyticsPathMode = GANTT_ANALYTICS_PATH_CP
        Case GANTT_ANALYTICS_PATH_CP
            gAnalyticsPathMode = GANTT_ANALYTICS_PATH_LP
        Case Else
            gAnalyticsPathMode = GANTT_ANALYTICS_PATH_NONE
    End Select

    Refresh_Gantt_UI_Only

End Sub

Public Sub Toggle_Gantt_Constraints()

    EnsureGanttViewInitialized
    gShowConstraints = Not gShowConstraints
    Refresh_Gantt_UI_Only

End Sub

Public Sub Toggle_Gantt_Scale()

    EnsureGanttViewInitialized

    If gTimelineScaleMode = GANTT_SCALE_WEEK Then
        gTimelineScaleMode = GANTT_SCALE_DAY
    Else
        gTimelineScaleMode = GANTT_SCALE_WEEK
    End If

    Refresh_Gantt_UI_Only

End Sub

Private Sub SetShapeVisibilityIfExists(ByVal ws As Worksheet, ByVal shapeName As String, ByVal isVisible As Boolean)

    Dim shp As Shape

    On Error Resume Next
    Set shp = ws.Shapes(shapeName)
    On Error GoTo 0

    If shp Is Nothing Then Exit Sub

    shp.Visible = IIf(isVisible, msoTrue, msoFalse)

End Sub

Private Sub SetRowRenderedShapesVisibility(ByVal ws As Worksheet, ByVal rowNum As Long, ByVal isVisible As Boolean)

    Dim suffix As String

    If rowNum < FIRST_TASK_ROW Then Exit Sub

    suffix = CStr(rowNum - FIRST_TASK_ROW + 1)

    SetShapeVisibilityIfExists ws, "TASK_" & suffix, isVisible
    SetShapeVisibilityIfExists ws, "TASK_" & suffix & "_P", isVisible

    SetShapeVisibilityIfExists ws, "MS_" & suffix, isVisible

    SetShapeVisibilityIfExists ws, "SUM_" & suffix & "_H", isVisible
    SetShapeVisibilityIfExists ws, "SUM_" & suffix & "_L", isVisible
    SetShapeVisibilityIfExists ws, "SUM_" & suffix & "_R", isVisible
    SetShapeVisibilityIfExists ws, "SUM_" & suffix & "_TXT", isVisible

    SetConstraintMarkerVisibilityForRow ws, suffix, isVisible

End Sub


Private Sub SetConstraintMarkerVisibilityForRow( _
    ByVal ws As Worksheet, _
    ByVal suffix As String, _
    ByVal isVisible As Boolean)

    Dim tokens As Variant
    Dim token As Variant
    Dim i As Long

    tokens = Array("SNET", "SNLT", "FNET", "FNLT")

    For Each token In tokens
        For i = 1 To 5
            SetShapeVisibilityIfExists ws, "CSTR_" & suffix & "_" & CStr(token) & "_" & CStr(i), isVisible
        Next i
    Next token

End Sub

Private Sub ApplyCurrentGanttView(ByVal ws As Worksheet)

    Dim lastRow As Long
    Dim r As Long
    Dim showRow As Boolean
    Dim shp As Shape
    Dim oldScreenUpdating As Boolean

    EnsureGanttViewInitialized

    oldScreenUpdating = Application.ScreenUpdating
    Application.ScreenUpdating = False
    On Error GoTo SafeExit

    ws.Rows.Hidden = False

    lastRow = GetLastGanttRow(ws)
    If lastRow < FIRST_TASK_ROW Then GoTo SafeExit

    If gGanttViewMode = GANTT_VIEW_DETAIL Then

        For r = FIRST_TASK_ROW To lastRow
            SetRowRenderedShapesVisibility ws, r, True
        Next r

        For Each shp In ws.Shapes
            If Left$(shp.Name, 4) = "DEP_" Then
                If IsWeekScaleMode() Then
                    shp.Visible = msoFalse
                Else
                    shp.Visible = msoTrue
                End If
            End If
        Next shp

        GoTo SafeExit
    End If

    For r = FIRST_TASK_ROW To lastRow
        showRow = ShouldShowGanttRow(ws, r)
        ws.Rows(r).Hidden = Not showRow
        SetRowRenderedShapesVisibility ws, r, showRow
    Next r

    For Each shp In ws.Shapes
        If Left$(shp.Name, 4) = "DEP_" Then
            shp.Visible = msoFalse
        End If
    Next shp


SafeExit:
    Application.ScreenUpdating = oldScreenUpdating

End Sub

Private Function ShouldShowGanttRow(ByVal ws As Worksheet, ByVal rowNum As Long) As Boolean

    Dim logicVal As String
    Dim durationVal As Variant
    Dim isParent As Boolean
    Dim isMilestone As Boolean

    EnsureGanttViewInitialized

    If gGanttViewMode = GANTT_VIEW_DETAIL Then
        ShouldShowGanttRow = True
        Exit Function
    End If

    logicVal = UCase$(Trim$(CStr(ws.Cells(rowNum, COL_LOGIC).value)))
    durationVal = ws.Cells(rowNum, COL_DURATION).value

    isParent = (logicVal = "SUMMARY")
    isMilestone = False

    If IsNumeric(durationVal) Then
        isMilestone = (CDbl(durationVal) <= 1)
    End If

    ShouldShowGanttRow = (isParent Or isMilestone)

End Function

Private Sub EnsureOrUpdateCompactToggleSwitch( _
    ByVal ws As Worksheet, _
    ByVal bgName As String, _
    ByVal knobName As String, _
    ByVal leftName As String, _
    ByVal leftCell As Range, _
    ByVal rightCell As Range, _
    ByVal leftText As String, _
    ByVal isOn As Boolean, _
    ByVal macroName As String)

    Dim shpBg As Shape
    Dim shpKnob As Shape
    Dim shpLeft As Shape

    Dim groupLeft As Double
    Dim groupTop As Double
    Dim groupWidth As Double
    Dim groupHeight As Double

    Dim leftLabelWidth As Double
    Dim gapBeforeTrack As Double

    Dim trackWidth As Double
    Dim trackHeight As Double
    Dim knobSize As Double

    Dim trackLeft As Double
    Dim trackTop As Double
    Dim knobLeft As Double
    Dim knobTop As Double

    Set shpBg = GetShapeIfExists(ws, bgName)
    Set shpKnob = GetShapeIfExists(ws, knobName)
    Set shpLeft = GetShapeIfExists(ws, leftName)

    trackWidth = 28
    trackHeight = 14
    knobSize = 9

    groupLeft = leftCell.Left + 4
    groupTop = leftCell.Top + 2
    groupWidth = (rightCell.Left + rightCell.Width) - groupLeft - 2
    groupHeight = leftCell.Height - 4

    gapBeforeTrack = 4
    leftLabelWidth = groupWidth - trackWidth - gapBeforeTrack
    If leftLabelWidth < 26 Then leftLabelWidth = 26

    trackLeft = groupLeft + leftLabelWidth + gapBeforeTrack
    trackTop = groupTop + ((groupHeight - trackHeight) / 2)
    knobTop = trackTop + ((trackHeight - knobSize) / 2)

    If isOn Then
        knobLeft = trackLeft + trackWidth - knobSize - 2
    Else
        knobLeft = trackLeft + 2
    End If

    If shpBg Is Nothing Then
        Set shpBg = ws.Shapes.AddShape(msoShapeRoundedRectangle, trackLeft, trackTop, trackWidth, trackHeight)
    End If

    If shpKnob Is Nothing Then
        Set shpKnob = ws.Shapes.AddShape(msoShapeOval, knobLeft, knobTop, knobSize, knobSize)
    End If

    If shpLeft Is Nothing Then
        Set shpLeft = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, groupLeft, groupTop, leftLabelWidth, groupHeight)
    End If

    shpBg.Name = bgName
    shpKnob.Name = knobName
    shpLeft.Name = leftName

    With shpBg
        .Left = trackLeft
        .Top = trackTop
        .Width = trackWidth
        .Height = trackHeight
        .Adjustments.item(1) = 0.5
        .OnAction = macroName
        .Placement = xlMoveAndSize
        .Fill.ForeColor.RGB = RGB(230, 230, 230)
        .Line.ForeColor.RGB = RGB(170, 170, 170)
        .Line.Weight = 0.75
        .Shadow.Visible = msoFalse
    End With

    With shpKnob
        .Left = knobLeft
        .Top = knobTop
        .Width = knobSize
        .Height = knobSize
        .OnAction = macroName
        .Placement = xlMoveAndSize
        .Fill.ForeColor.RGB = RGB(255, 255, 255)
        .Line.ForeColor.RGB = RGB(150, 150, 150)
        .Line.Weight = 0.5
        .Shadow.Visible = msoFalse
    End With

    With shpLeft
        .Left = groupLeft
        .Top = groupTop
        .Width = leftLabelWidth
        .Height = groupHeight
        .Line.Visible = msoFalse
        .Fill.Visible = msoFalse
        .OnAction = macroName
        .Placement = xlMoveAndSize
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
        .TextFrame2.MarginLeft = 0
        .TextFrame2.MarginRight = 0
        .TextFrame2.MarginTop = 0
        .TextFrame2.MarginBottom = 0
        .TextFrame2.TextRange.Text = leftText
        .TextFrame2.TextRange.Font.Size = 8.5
        .TextFrame2.TextRange.Font.Bold = msoTrue
        .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignLeft
    End With

    If isOn Then
        shpLeft.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(90, 90, 90)
    Else
        shpLeft.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(68, 114, 196)
    End If

End Sub

Public Function GetShapeIfExists(ByVal ws As Worksheet, ByVal shapeName As String) As Shape

    On Error Resume Next
    Set GetShapeIfExists = ws.Shapes(shapeName)
    On Error GoTo 0

End Function

Private Function GetGanttAnalyticsPathColumnName() As String

    EnsureGanttViewInitialized

    Select Case gAnalyticsPathMode
        Case GANTT_ANALYTICS_PATH_CP
            GetGanttAnalyticsPathColumnName = "Critical Path"
        Case GANTT_ANALYTICS_PATH_LP
            GetGanttAnalyticsPathColumnName = "Longest Path"
        Case Else
            GetGanttAnalyticsPathColumnName = vbNullString
    End Select

End Function

Private Function ShouldHighlightGanttAnalyticsPath( _
    ByRef dataArr As Variant, _
    ByVal mapWBS As Object, _
    ByVal dataRow As Long) As Boolean

    Dim pathColumnName As String
    Dim pathValue As String

    pathColumnName = GetGanttAnalyticsPathColumnName()
    If Len(pathColumnName) = 0 Then Exit Function
    If mapWBS Is Nothing Then Exit Function
    If Not mapWBS.Exists(pathColumnName) Then Exit Function

    pathValue = UCase$(Trim$(CStr(dataArr(dataRow, mapWBS(pathColumnName)))))

    Select Case pathColumnName
        Case "Critical Path"
            ShouldHighlightGanttAnalyticsPath = (pathValue = "CRITICAL")
        Case "Longest Path"
            ShouldHighlightGanttAnalyticsPath = (pathValue = "LONGEST")
    End Select

End Function
Private Function IsGanttAnalyticsPathHighlightEnabled() As Boolean

    IsGanttAnalyticsPathHighlightEnabled = (Len(GetGanttAnalyticsPathColumnName()) > 0)

End Function

Private Function GetTaskBaseColor(ByVal isCritical As Boolean) As Long

    If IsGanttAnalyticsPathHighlightEnabled() And isCritical Then
        GetTaskBaseColor = COLOR_TASK_CRITICAL
    Else
        GetTaskBaseColor = COLOR_TASK_BLUE
    End If

End Function

Private Function GetSummaryLineColor( _
    ByVal isCritical As Boolean, _
    ByVal isComplete As Boolean) As Long

    If isComplete Then
        GetSummaryLineColor = COLOR_PROGRESS_GREEN
    ElseIf IsGanttAnalyticsPathHighlightEnabled() And isCritical Then
        GetSummaryLineColor = COLOR_TASK_CRITICAL
    Else
        GetSummaryLineColor = COLOR_TASK_BLUE
    End If

End Function

Private Function IsParentWBS(ByVal wbs As String, ByRef dataArr As Variant, ByVal mapWBS As Object) As Boolean

    Dim i As Long
    Dim otherWBS As String
    Dim prefix As String

    If Trim$(wbs) = "" Then Exit Function

    prefix = wbs & "."

    For i = 1 To UBound(dataArr, 1)
        otherWBS = NormalizeWBS(CStr(dataArr(i, mapWBS("WBS"))))
        If Left$(otherWBS, Len(prefix)) = prefix Then
            IsParentWBS = True
            Exit Function
        End If
    Next i

End Function


Private Function GetProgressFillColor( _
    ByVal startVal As Variant, _
    ByVal finishVal As Variant, _
    ByVal progressVal As Double) As Variant

    Dim durationVal As Double
    Dim coveredDays As Long
    Dim progressFinish As Double
    Dim todayVal As Double

    If progressVal <= 0 Then Exit Function

    If progressVal >= 1 Then
        GetProgressFillColor = COLOR_PROGRESS_GREEN
        Exit Function
    End If

    If Not HasValue(startVal) Then Exit Function
    If Not HasValue(finishVal) Then Exit Function

    durationVal = CDbl(finishVal) - CDbl(startVal) + 1
    If durationVal <= 0 Then Exit Function

    coveredDays = Int(durationVal * progressVal + 0.999999)
    If coveredDays < 1 Then coveredDays = 1

    progressFinish = CDbl(startVal) + coveredDays - 1
    todayVal = CDbl(Date)

    ' Tolérance : on considère "à jour" si le progress couvre aujourd'hui - 1
    If progressFinish >= todayVal - 1 Then
        GetProgressFillColor = COLOR_PROGRESS_GREEN
    Else
        GetProgressFillColor = COLOR_PROGRESS_ORANGE
    End If

End Function

Private Function BuildParentCompleteMap(ByRef dataArr As Variant, ByVal mapWBS As Object) As Object

    Dim parentMap As Object
    Dim i As Long
    Dim wbsVal As String
    Dim parentWbs As String
    Dim progressVal As Double

    Set parentMap = CreateObject("Scripting.Dictionary")

    For i = 1 To UBound(dataArr, 1)

        wbsVal = NormalizeWBS(CStr(dataArr(i, mapWBS("WBS"))))
        If wbsVal = "" Then GoTo NextI

        If IsParentWBS(wbsVal, dataArr, mapWBS) Then GoTo NextI

        If HasValue(dataArr(i, mapWBS("% Progress"))) Then
            progressVal = CDbl(dataArr(i, mapWBS("% Progress")))
        Else
            progressVal = 0
        End If

        parentWbs = GetParentWBS(wbsVal)

        Do While parentWbs <> ""
            If Not parentMap.Exists(parentWbs) Then
                parentMap(parentWbs) = True
            End If

            If progressVal < 1 Then
                parentMap(parentWbs) = False
            End If

            parentWbs = GetParentWBS(parentWbs)
        Loop

NextI:
    Next i

    Set BuildParentCompleteMap = parentMap

End Function

Private Function ParentIsCompleteFromMap(ByVal parentWbs As String, ByVal parentCompleteMap As Object) As Boolean

    If parentCompleteMap Is Nothing Then Exit Function
    If parentWbs = "" Then Exit Function

    If parentCompleteMap.Exists(parentWbs) Then
        ParentIsCompleteFromMap = CBool(parentCompleteMap(parentWbs))
    End If

End Function

Public Sub EnsureOrUpdateToggleSwitch( _
    ByVal ws As Worksheet, _
    ByVal bgName As String, _
    ByVal knobName As String, _
    ByVal leftName As String, _
    ByVal rightName As String, _
    ByVal leftCell As Range, _
    ByVal rightCell As Range, _
    ByVal leftText As String, _
    ByVal rightText As String, _
    ByVal isOn As Boolean, _
    ByVal macroName As String)

    Dim shpBg As Shape
    Dim shpKnob As Shape
    Dim shpLeft As Shape
    Dim shpRight As Shape

    Dim groupLeft As Double
    Dim groupTop As Double
    Dim groupHeight As Double

    Dim leftLabelWidth As Double
    Dim rightLabelWidth As Double
    Dim gapBeforeTrack As Double
    Dim gapAfterTrack As Double

    Dim trackWidth As Double
    Dim trackHeight As Double
    Dim knobSize As Double

    Dim trackLeft As Double
    Dim trackTop As Double
    Dim knobLeft As Double
    Dim knobTop As Double

    Set shpBg = GetShapeIfExists(ws, bgName)
    Set shpKnob = GetShapeIfExists(ws, knobName)
    Set shpLeft = GetShapeIfExists(ws, leftName)
    Set shpRight = GetShapeIfExists(ws, rightName)

    trackWidth = 42
    trackHeight = 20
    knobSize = 14

    If bgName = BTN_CP_BG_NAME Then
        groupLeft = leftCell.Left + 20
        leftLabelWidth = 88
        rightLabelWidth = 24
        gapBeforeTrack = 10
        gapAfterTrack = 8

    ElseIf bgName = BTN_SCALE_BG_NAME Then
        groupLeft = leftCell.Left + 80
        leftLabelWidth = 26
        rightLabelWidth = 40
        gapBeforeTrack = 10
        gapAfterTrack = 8

    Else
        groupLeft = leftCell.Left + 10
        leftLabelWidth = 34
        rightLabelWidth = 42
        gapBeforeTrack = 8
        gapAfterTrack = 8
    End If

    groupTop = leftCell.Top + 2
    groupHeight = leftCell.Height - 4

    trackLeft = groupLeft + leftLabelWidth + gapBeforeTrack
    trackTop = groupTop + ((groupHeight - trackHeight) / 2)
    knobTop = trackTop + ((trackHeight - knobSize) / 2)

    If isOn Then
        knobLeft = trackLeft + trackWidth - knobSize - 3
    Else
        knobLeft = trackLeft + 3
    End If

    If shpBg Is Nothing Then
        Set shpBg = ws.Shapes.AddShape(msoShapeRoundedRectangle, trackLeft, trackTop, trackWidth, trackHeight)
    End If

    If shpKnob Is Nothing Then
        Set shpKnob = ws.Shapes.AddShape(msoShapeOval, knobLeft, knobTop, knobSize, knobSize)
    End If

    If shpLeft Is Nothing Then
        Set shpLeft = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, groupLeft, groupTop, leftLabelWidth, groupHeight)
    End If

    If shpRight Is Nothing Then
        Set shpRight = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, trackLeft + trackWidth + gapAfterTrack, groupTop, rightLabelWidth, groupHeight)
    End If

    shpBg.Name = bgName
    shpKnob.Name = knobName
    shpLeft.Name = leftName
    shpRight.Name = rightName

    FormatToggleTrack shpBg, trackLeft, trackTop, macroName
    FormatToggleKnob shpKnob, knobLeft, knobTop, macroName
    FormatToggleLabel shpLeft, groupLeft, groupTop, leftLabelWidth, groupHeight, leftText, macroName
    FormatToggleLabel shpRight, trackLeft + trackWidth + gapAfterTrack, groupTop, rightLabelWidth, groupHeight, rightText, macroName

    If isOn Then
        shpRight.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(68, 114, 196)
        shpLeft.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(90, 90, 90)
    Else
        shpLeft.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(68, 114, 196)
        shpRight.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(90, 90, 90)
    End If

End Sub

Public Sub FormatToggleTrack( _
    ByVal shp As Shape, _
    ByVal leftPos As Double, _
    ByVal topPos As Double, _
    ByVal macroName As String)

    With shp
        .Left = leftPos
        .Top = topPos
        .Width = 42
        .Height = 20
        .Adjustments.item(1) = 0.5
        .OnAction = macroName
        .Placement = xlMoveAndSize
        .Fill.ForeColor.RGB = RGB(230, 230, 230)
        .Line.ForeColor.RGB = RGB(170, 170, 170)
        .Line.Weight = 1
        .Shadow.Visible = msoFalse
    End With

End Sub

Public Sub FormatToggleKnob( _
    ByVal shp As Shape, _
    ByVal leftPos As Double, _
    ByVal topPos As Double, _
    ByVal macroName As String)

    With shp
        .Left = leftPos
        .Top = topPos
        .Width = 14
        .Height = 14
        .OnAction = macroName
        .Placement = xlMoveAndSize
        .Fill.ForeColor.RGB = RGB(255, 255, 255)
        .Line.ForeColor.RGB = RGB(150, 150, 150)
        .Line.Weight = 0.75
        .Shadow.Visible = msoTrue
        .Shadow.Blur = 4
        .Shadow.OffsetX = 0
        .Shadow.OffsetY = 1
        .Shadow.Transparency = 0.45
    End With

End Sub

Public Sub FormatToggleLabel( _
    ByVal shp As Shape, _
    ByVal leftPos As Double, _
    ByVal topPos As Double, _
    ByVal widthVal As Double, _
    ByVal heightVal As Double, _
    ByVal labelText As String, _
    ByVal macroName As String)

    With shp
        .Left = leftPos
        .Top = topPos
        .Width = widthVal
        .Height = heightVal
        .Line.Visible = msoFalse
        .Fill.Visible = msoFalse
        .OnAction = macroName
        .Placement = xlMoveAndSize
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
        .TextFrame2.MarginLeft = 0
        .TextFrame2.MarginRight = 0
        .TextFrame2.MarginTop = 0
        .TextFrame2.MarginBottom = 0
        .TextFrame2.TextRange.Text = labelText
        .TextFrame2.TextRange.Font.Size = 9.5
        .TextFrame2.TextRange.Font.Bold = msoTrue
        .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    End With

End Sub

Private Sub ApplyGanttUiState(ByVal ws As Worksheet)

    Ensure_Gantt_Test_Buttons
    ApplyCurrentGanttView ws

End Sub

Private Sub Refresh_Gantt_UI_Only()

    Dim oldPreserve As Boolean

    On Error GoTo SafeExit

    oldPreserve = GetGanttPreserveTestInputs()

    SetGanttPreserveTestInputs True
    Refresh_Gantt_DisplayOnly

SafeExit:
    SetGanttPreserveTestInputs oldPreserve

End Sub

Private Sub DrawSingleWeekTask( _
    ByVal ws As Worksheet, _
    ByVal ganttRow As Long, _
    ByVal projectStart As Variant, _
    ByVal startVal As Variant, _
    ByVal progressVal As Double, _
    ByVal isCritical As Boolean, _
    ByVal totalDays As Long, _
    ByVal shapeKey As String, _
    ByVal hasDelta As Boolean)

    Dim leftPos As Double
    Dim topPos As Double
    Dim sizeVal As Double
    Dim shp As Shape
    Dim shpProgress As Shape
    Dim targetCol As Long
    Dim cellLeft As Double
    Dim cellWidth As Double
    Dim fillColor As Long
    Dim innerSize As Double

    If Not HasValue(startVal) Then Exit Sub

    targetCol = TimelineColumnFromHeaderDate_Exact(ws, projectStart, startVal)
    If targetCol < FIRST_TIMELINE_COL Then Exit Sub

    cellLeft = ws.Cells(HEADER_ROW_2, targetCol).Left
    cellWidth = ws.Cells(HEADER_ROW_2, targetCol).Width

    sizeVal = WorksheetFunction.Min(ws.Rows(ganttRow).Height - 6, cellWidth - 6)
    If sizeVal < 2 Then sizeVal = 2

    leftPos = cellLeft + ((cellWidth - sizeVal) / 2)
    topPos = ws.Cells(ganttRow, FIRST_TIMELINE_COL).Top + ((ws.Rows(ganttRow).Height - sizeVal) / 2)

    If progressVal >= 1 Then
        fillColor = COLOR_PROGRESS_GREEN
    Else
        fillColor = GetTaskBaseColor(isCritical)
    End If

    Set shp = ws.Shapes.AddShape(msoShapeOval, leftPos, topPos, sizeVal, sizeVal)
    shp.Name = shapeKey
    ApplyGanttRenderShapePlacement shp

    With shp
        .Left = leftPos
        .Top = topPos
        .Width = sizeVal
        .Height = sizeVal
        .Fill.ForeColor.RGB = fillColor
    End With

    If hasDelta Then
        shp.Line.Visible = msoTrue
        shp.Line.ForeColor.RGB = RGB(255, 192, 0)
        shp.Line.Weight = 2.75
    Else
        shp.Line.Visible = msoFalse
    End If

    If progressVal > 0 And progressVal < 1 Then

        innerSize = sizeVal * 0.55

        Set shpProgress = ws.Shapes.AddShape( _
            msoShapeOval, _
            leftPos + ((sizeVal - innerSize) / 2), _
            topPos + ((sizeVal - innerSize) / 2), _
            innerSize, _
            innerSize)

        shpProgress.Name = shapeKey & "_P"
        ApplyGanttRenderShapePlacement shpProgress

        With shpProgress
            .Left = leftPos + ((sizeVal - innerSize) / 2)
            .Top = topPos + ((sizeVal - innerSize) / 2)
            .Width = innerSize
            .Height = innerSize
            .Line.Visible = msoFalse
            .Fill.ForeColor.RGB = GetProgressFillColor(startVal, startVal, progressVal)
            .Fill.Transparency = 0.15
        End With

    End If

End Sub


Private Function IsTaskContainedInSingleIsoWeek(ByVal startVal As Variant, ByVal finishVal As Variant) As Boolean

    If Not HasValue(startVal) Then Exit Function
    If Not HasValue(finishVal) Then Exit Function

    IsTaskContainedInSingleIsoWeek = (GetIsoWeekMonday(CDate(startVal)) = GetIsoWeekMonday(CDate(finishVal)))

End Function

Private Sub GetRenderProgressBounds( _
    ByVal rawStartVal As Variant, _
    ByVal rawFinishVal As Variant, _
    ByVal progressVal As Double, _
    ByRef renderStartOut As Variant, _
    ByRef renderFinishOut As Variant)

    Dim rawDuration As Long
    Dim coveredDays As Long
    Dim rawProgressFinish As Date

    renderStartOut = Empty
    renderFinishOut = Empty

    If progressVal <= 0 Then Exit Sub
    If Not HasValue(rawStartVal) Then Exit Sub
    If Not HasValue(rawFinishVal) Then Exit Sub

    rawDuration = CLng(CDbl(rawFinishVal) - CDbl(rawStartVal) + 1)
    If rawDuration <= 0 Then Exit Sub

    coveredDays = Int(rawDuration * WorksheetFunction.Min(progressVal, 1) + 0.999999)
    If coveredDays < 1 Then coveredDays = 1

    rawProgressFinish = CDate(rawStartVal) + coveredDays - 1

    If IsWeekScaleMode() Then
        renderStartOut = GetIsoWeekMonday(CDate(rawStartVal))
        renderFinishOut = GetIsoWeekMonday(rawProgressFinish) + 6
    Else
        renderStartOut = rawStartVal
        renderFinishOut = rawProgressFinish
    End If

End Sub

Private Function GetGanttRowTop(ByVal ws As Worksheet, ByVal ganttRow As Long) As Double

    GetGanttRowTop = ws.Cells(ganttRow, FIRST_TIMELINE_COL).Top

End Function

Private Function GetGanttRowHeight(ByVal ws As Worksheet, ByVal ganttRow As Long) As Double

    GetGanttRowHeight = ws.Rows(ganttRow).Height

End Function

Private Function GetGanttRowMid(ByVal ws As Worksheet, ByVal ganttRow As Long) As Double

    GetGanttRowMid = GetGanttRowTop(ws, ganttRow) + (GetGanttRowHeight(ws, ganttRow) / 2)

End Function

Private Function GetGanttBarTop(ByVal ws As Worksheet, ByVal ganttRow As Long) As Double

    GetGanttBarTop = GetGanttRowTop(ws, ganttRow) + 4

End Function

Private Function GetGanttBarHeight(ByVal ws As Worksheet, ByVal ganttRow As Long) As Double

    GetGanttBarHeight = GetGanttRowHeight(ws, ganttRow) - 8
    If GetGanttBarHeight < 2 Then GetGanttBarHeight = 2

End Function

Private Function GetGanttMilestoneSize(ByVal ws As Worksheet, ByVal ganttRow As Long) As Double

    GetGanttMilestoneSize = GetGanttRowHeight(ws, ganttRow) - 6
    If GetGanttMilestoneSize < 2 Then GetGanttMilestoneSize = 2

End Function

Private Function GetGanttSummaryTop(ByVal ws As Worksheet, ByVal ganttRow As Long) As Double

    GetGanttSummaryTop = GetGanttRowTop(ws, ganttRow) + 4

End Function

Private Function GetGanttSummaryBottom(ByVal ws As Worksheet, ByVal ganttRow As Long) As Double

    GetGanttSummaryBottom = GetGanttRowTop(ws, ganttRow) + GetGanttRowHeight(ws, ganttRow) - 4

End Function

Private Sub ApplyGanttRenderShapePlacement(ByVal shp As Shape)

    On Error Resume Next

    'Important:
    'The Gantt renderer deletes/recreates shapes on each refresh.
    'Using xlMoveAndSize lets Excel re-anchor and micro-resize shapes when
    'columns/rows/merged headers are rebuilt, which can create cumulative drift.
    'Keep rendered shapes free-floating; their exact Left/Top/Width/Height are
    'fully controlled by the renderer.
    shp.Placement = xlMoveAndSize

    On Error GoTo 0

End Sub

Private Sub ApplyGanttRenderLinePlacement(ByVal shp As Shape)

    On Error Resume Next

    'Same rule as bars/milestones:
    'dependency lines and Today line are redrawn from scratch, so they must not
    'be cell-resized by Excel between layout rebuild and final render.
    shp.Placement = xlMoveAndSize

    On Error GoTo 0

End Sub


'=====================================================
' Helpers for factorized refresh core
'=====================================================

Private Function BuildHasChildrenMap( _
    ByRef dataArr As Variant, _
    ByVal mapWBS As Object) As Object

    Dim d As Object
    Dim r As Long
    Dim wbsVal As String
    Dim parentWbs As String

    Set d = CreateObject("Scripting.Dictionary")

    For r = 1 To UBound(dataArr, 1)
        wbsVal = NormalizeWBS(CStr(dataArr(r, mapWBS("WBS"))))
        parentWbs = GetParentWBS(wbsVal)

        If parentWbs <> "" Then
            d(parentWbs) = True
        End If
    Next r

    Set BuildHasChildrenMap = d

End Function

Private Function BuildRowByIdMap( _
    ByRef dataArr As Variant, _
    ByVal mapWBS As Object) As Object

    Dim d As Object
    Dim r As Long
    Dim idVal As String

    Set d = CreateObject("Scripting.Dictionary")

    For r = 1 To UBound(dataArr, 1)
        idVal = Trim$(CStr(dataArr(r, mapWBS("ID"))))
        If idVal <> "" Then
            d(idVal) = FIRST_TASK_ROW + r - 1
        End If
    Next r

    Set BuildRowByIdMap = d

End Function

Private Function BuildGanttBaseDisplayMap() As Object

    Set BuildGanttBaseDisplayMap = GanttLive_BuildBaseByIdMap()

End Function

Private Function BuildGanttTestDisplayMap() As Object

    Set BuildGanttTestDisplayMap = GanttLive_BuildTestByIdMap()

End Function

Private Sub GetProjectDisplayRange( _
    ByRef dataArr As Variant, _
    ByVal mapWBS As Object, _
    ByVal hasChildren As Object, _
    ByVal baseById As Object, _
    ByVal testById As Object, _
    ByVal isTestMode As Boolean, _
    ByRef projectStart As Variant, _
    ByRef projectFinish As Variant)

    Dim r As Long
    Dim idVal As String
    Dim displayStart As Variant
    Dim displayFinish As Variant

    projectStart = Empty
    projectFinish = Empty

    For r = 1 To UBound(dataArr, 1)

        idVal = Trim$(CStr(dataArr(r, mapWBS("ID"))))

        displayStart = GanttLive_GetDisplayStart(idVal, baseById, testById, isTestMode)
        displayFinish = GanttLive_GetDisplayFinish(idVal, baseById, testById, isTestMode)

        If HasValue(displayStart) And HasValue(displayFinish) Then

            If Not HasValue(projectStart) Then
                projectStart = displayStart
            ElseIf CDbl(displayStart) < CDbl(projectStart) Then
                projectStart = displayStart
            End If

            If Not HasValue(projectFinish) Then
                projectFinish = displayFinish
            ElseIf CDbl(displayFinish) > CDbl(projectFinish) Then
                projectFinish = displayFinish
            End If

        End If
    Next r

End Sub


Private Sub EnsureExpandedLinksCacheFromCalc()

    Set gExpandedLinks = Nothing
    Set gExpandedLinks = BuildExpandedLinksCacheFromLogicLinksTable()

End Sub
Private Function BuildExpandedLinksCacheFromLogicLinksTable() As Object

    Dim wsCalc As Worksheet
    Dim tblLinks As ListObject
    Dim mapLinks As Object
    Dim arrLinks As Variant

    Dim d As Object
    Dim linkCol As Collection
    Dim tokenInfo As Object

    Dim r As Long
    Dim i As Long

    Dim succId As String
    Dim predId As String
    Dim linkType As String
    Dim lagVal As Double
    Dim rawToken As String

    Set d = CreateObject("Scripting.Dictionary")
    Set mapLinks = CreateObject("Scripting.Dictionary")

    On Error GoTo SafeExit

    Set wsCalc = ThisWorkbook.Worksheets(CALC_SHEET)
    Set tblLinks = wsCalc.ListObjects("tbl_LOGIC_LINKS")

    If tblLinks Is Nothing Then
        Set BuildExpandedLinksCacheFromLogicLinksTable = d
        Exit Function
    End If

    If tblLinks.DataBodyRange Is Nothing Then
        Set BuildExpandedLinksCacheFromLogicLinksTable = d
        Exit Function
    End If

    For i = 1 To tblLinks.ListColumns.Count
        mapLinks(tblLinks.ListColumns(i).Name) = i
    Next i

    If Not mapLinks.Exists("Succ ID") Then GoTo SafeExit
    If Not mapLinks.Exists("Pred ID") Then GoTo SafeExit
    If Not mapLinks.Exists("Link Type") Then GoTo SafeExit
    If Not mapLinks.Exists("Lag") Then GoTo SafeExit

    arrLinks = tblLinks.DataBodyRange.value

    For r = 1 To UBound(arrLinks, 1)

        succId = Trim$(CStr(arrLinks(r, mapLinks("Succ ID"))))
        predId = Trim$(CStr(arrLinks(r, mapLinks("Pred ID"))))
        linkType = UCase$(Trim$(CStr(arrLinks(r, mapLinks("Link Type")))))

        If succId = "" Then GoTo NextRow
        If predId = "" Then GoTo NextRow

        If linkType = "" Then linkType = "FS"
        If linkType <> "FS" And linkType <> "SS" And linkType <> "FF" Then GoTo NextRow

        If IsNumeric(arrLinks(r, mapLinks("Lag"))) Then
            lagVal = CDbl(arrLinks(r, mapLinks("Lag")))
        Else
            lagVal = 0#
        End If

        If mapLinks.Exists("Raw Token") Then
            rawToken = Trim$(CStr(arrLinks(r, mapLinks("Raw Token"))))
        Else
            rawToken = vbNullString
        End If

        If Not d.Exists(succId) Then
            Set linkCol = New Collection
            d.Add succId, linkCol
        End If

        Set tokenInfo = CreateObject("Scripting.Dictionary")
        tokenInfo("PredID") = predId
        tokenInfo("LinkType") = linkType
        tokenInfo("Lag") = lagVal
        tokenInfo("RawToken") = rawToken

        d(succId).Add tokenInfo

NextRow:
    Next r

SafeExit:
    Set BuildExpandedLinksCacheFromLogicLinksTable = d

End Function

Private Function HasExpandedLinksAvailable() As Boolean

    On Error GoTo SafeExit

    If gExpandedLinks Is Nothing Then Exit Function
    If gExpandedLinks.Count <= 0 Then Exit Function

    HasExpandedLinksAvailable = True
    Exit Function

SafeExit:
    HasExpandedLinksAvailable = False

End Function

Private Sub GetLinkAnchorTypes( _
    ByVal linkType As String, _
    ByRef predAnchorType As String, _
    ByRef succAnchorType As String)

    Select Case UCase$(Trim$(linkType))
        Case "SS"
            predAnchorType = LINK_ANCHOR_START
            succAnchorType = LINK_ANCHOR_START

        Case "FF"
            predAnchorType = LINK_ANCHOR_FINISH
            succAnchorType = LINK_ANCHOR_FINISH

        Case Else
            predAnchorType = LINK_ANCHOR_FINISH
            succAnchorType = LINK_ANCHOR_START
    End Select

End Sub

Private Function GetLinkReferenceDate( _
    ByVal taskId As String, _
    ByVal anchorType As String, _
    ByVal baseById As Object, _
    ByVal testById As Object, _
    ByVal isTestMode As Boolean) As Variant

    Select Case UCase$(Trim$(anchorType))
        Case LINK_ANCHOR_START
            GetLinkReferenceDate = GanttLive_GetDisplayStart(taskId, baseById, testById, isTestMode)

        Case Else
            GetLinkReferenceDate = GanttLive_GetDisplayFinish(taskId, baseById, testById, isTestMode)
    End Select

End Function
Private Sub GetTaskAnchorPoint( _
    ByVal ws As Worksheet, _
    ByVal mapWBS As Object, _
    ByRef dataArr As Variant, _
    ByVal hasChildren As Object, _
    ByVal projectStart As Variant, _
    ByVal totalDays As Long, _
    ByVal dataRow As Long, _
    ByVal isFinishSide As Boolean, _
    ByRef xOut As Double, _
    ByRef yOut As Double, _
    ByVal baseById As Object, _
    ByVal testById As Object, _
    ByVal isTestMode As Boolean)

    Dim ganttRow As Long
    Dim wbs As String
    Dim idVal As String
    Dim startVal As Variant
    Dim finishVal As Variant
    Dim durationVal As Double
    Dim sizeVal As Double

    Dim timelineLeftBound As Double
    Dim timelineRightBound As Double

    ganttRow = FIRST_TASK_ROW + dataRow - 1
    wbs = NormalizeWBS(CStr(dataArr(dataRow, mapWBS("WBS"))))
    idVal = Trim$(CStr(dataArr(dataRow, mapWBS("ID"))))

    startVal = GetRenderStartForCurrentScale(GanttLive_GetDisplayStart(idVal, baseById, testById, isTestMode))
    finishVal = GetRenderFinishForCurrentScale(GanttLive_GetDisplayFinish(idVal, baseById, testById, isTestMode))

    If Not HasValue(startVal) Or Not HasValue(finishVal) Then Exit Sub

    durationVal = CDbl(finishVal) - CDbl(startVal) + 1

    timelineLeftBound = ws.Cells(HEADER_ROW_2, FIRST_TIMELINE_COL).Left + LINK_EDGE_PADDING
    timelineRightBound = ws.Cells(HEADER_ROW_2, FIRST_TIMELINE_COL + totalDays - 1).Left + _
                         ws.Cells(HEADER_ROW_2, FIRST_TIMELINE_COL + totalDays - 1).Width - LINK_EDGE_PADDING

    yOut = ws.Cells(ganttRow, FIRST_TIMELINE_COL).Top + (ws.Rows(ganttRow).Height / 2)

    If hasChildren.Exists(wbs) Then
        If isFinishSide Then
            xOut = TimelineRightAfterFinish(ws, projectStart, finishVal)
        Else
            xOut = TimelineLeft(ws, projectStart, startVal)
        End If

    ElseIf durationVal <= 1 Then
        sizeVal = ws.Rows(ganttRow).Height - 6

        If isFinishSide Then
            xOut = GetTaskMidX(ws, projectStart, startVal) + (sizeVal / 2)
        Else
            xOut = GetTaskMidX(ws, projectStart, startVal) - (sizeVal / 2)
        End If

    Else
        If isFinishSide Then
            xOut = TimelineRightAfterFinish(ws, projectStart, finishVal)
        Else
            xOut = TimelineLeft(ws, projectStart, startVal)
        End If
    End If

    If xOut < timelineLeftBound Then xOut = timelineLeftBound
    If xOut > timelineRightBound Then xOut = timelineRightBound

End Sub


Private Sub GetTaskAnchorPointByType( _
    ByVal ws As Worksheet, _
    ByVal mapWBS As Object, _
    ByRef dataArr As Variant, _
    ByVal hasChildren As Object, _
    ByVal projectStart As Variant, _
    ByVal totalDays As Long, _
    ByVal dataRow As Long, _
    ByVal anchorType As String, _
    ByRef xOut As Double, _
    ByRef yOut As Double, _
    ByVal baseById As Object, _
    ByVal testById As Object, _
    ByVal isTestMode As Boolean)

    GetTaskAnchorPointBySide ws, mapWBS, dataArr, hasChildren, projectStart, totalDays, dataRow, _
                             IIf(UCase$(Trim$(anchorType)) = LINK_ANCHOR_START, "LEFT", "RIGHT"), _
                             xOut, yOut, baseById, testById, isTestMode

End Sub

Public Sub GetTaskFinishEntryPoint( _
    ByVal ws As Worksheet, _
    ByVal mapWBS As Object, _
    ByRef dataArr As Variant, _
    ByVal projectStart As Variant, _
    ByVal totalDays As Long, _
    ByVal dataRow As Long, _
    ByRef xOut As Double, _
    ByRef yOut As Double, _
    ByVal baseById As Object, _
    ByVal testById As Object, _
    ByVal isTestMode As Boolean)

    Dim ganttRow As Long
    Dim idVal As String
    Dim startVal As Variant
    Dim finishVal As Variant
    Dim durationVal As Double
    Dim sizeVal As Double
    Dim timelineLeftBound As Double
    Dim timelineRightBound As Double

    ganttRow = FIRST_TASK_ROW + dataRow - 1
    idVal = Trim$(CStr(dataArr(dataRow, mapWBS("ID"))))

    startVal = GetRenderStartForCurrentScale(GanttLive_GetDisplayStart(idVal, baseById, testById, isTestMode))
    finishVal = GetRenderFinishForCurrentScale(GanttLive_GetDisplayFinish(idVal, baseById, testById, isTestMode))

    If Not HasValue(startVal) Or Not HasValue(finishVal) Then Exit Sub

    durationVal = CDbl(finishVal) - CDbl(startVal) + 1

    timelineLeftBound = ws.Cells(HEADER_ROW_2, FIRST_TIMELINE_COL).Left + LINK_EDGE_PADDING
    timelineRightBound = ws.Cells(HEADER_ROW_2, FIRST_TIMELINE_COL + totalDays - 1).Left + _
                         ws.Cells(HEADER_ROW_2, FIRST_TIMELINE_COL + totalDays - 1).Width - LINK_EDGE_PADDING

    yOut = ws.Cells(ganttRow, FIRST_TIMELINE_COL).Top + (ws.Rows(ganttRow).Height / 2)

    If durationVal <= 1 Then
        sizeVal = ws.Rows(ganttRow).Height - 6
        xOut = GetTaskMidX(ws, projectStart, startVal) + (sizeVal / 2)
    Else
        xOut = TimelineRightAfterFinish(ws, projectStart, finishVal)
    End If

    If xOut < timelineLeftBound Then xOut = timelineLeftBound
    If xOut > timelineRightBound Then xOut = timelineRightBound

End Sub

Private Function GetTaskLeftX( _
    ByVal ws As Worksheet, _
    ByVal projectStart As Variant, _
    ByVal taskDate As Variant) As Double

    GetTaskLeftX = TimelineLeft(ws, projectStart, taskDate)

End Function

Private Function GetTaskRightX( _
    ByVal ws As Worksheet, _
    ByVal projectStart As Variant, _
    ByVal taskDate As Variant) As Double

    GetTaskRightX = TimelineRightAfterFinish(ws, projectStart, taskDate)

End Function

Private Function GetTaskMidX( _
    ByVal ws As Worksheet, _
    ByVal projectStart As Variant, _
    ByVal taskDate As Variant) As Double

    Dim targetCol As Long

    targetCol = TimelineColumnFromHeaderDate_Exact(ws, projectStart, taskDate)
    If targetCol < FIRST_TIMELINE_COL Then Exit Function

    GetTaskMidX = ws.Cells(HEADER_ROW_2, targetCol).Left + (ws.Cells(HEADER_ROW_2, targetCol).Width / 2)

End Function

Private Sub GetTaskAnchorPointBySide( _
    ByVal ws As Worksheet, _
    ByVal mapWBS As Object, _
    ByRef dataArr As Variant, _
    ByVal hasChildren As Object, _
    ByVal projectStart As Variant, _
    ByVal totalDays As Long, _
    ByVal dataRow As Long, _
    ByVal anchorSide As String, _
    ByRef xOut As Double, _
    ByRef yOut As Double, _
    ByVal baseById As Object, _
    ByVal testById As Object, _
    ByVal isTestMode As Boolean)

    Dim ganttRow As Long
    Dim wbs As String
    Dim idVal As String
    Dim startVal As Variant
    Dim finishVal As Variant
    Dim durationVal As Double
    Dim sizeVal As Double
    Dim timelineLeftBound As Double
    Dim timelineRightBound As Double

    ganttRow = FIRST_TASK_ROW + dataRow - 1
    wbs = NormalizeWBS(CStr(dataArr(dataRow, mapWBS("WBS"))))
    idVal = Trim$(CStr(dataArr(dataRow, mapWBS("ID"))))

    startVal = GetRenderStartForCurrentScale(GanttLive_GetDisplayStart(idVal, baseById, testById, isTestMode))
    finishVal = GetRenderFinishForCurrentScale(GanttLive_GetDisplayFinish(idVal, baseById, testById, isTestMode))

    If Not HasValue(startVal) Or Not HasValue(finishVal) Then Exit Sub

    durationVal = CDbl(finishVal) - CDbl(startVal) + 1

    timelineLeftBound = ws.Cells(HEADER_ROW_2, FIRST_TIMELINE_COL).Left + LINK_EDGE_PADDING
    timelineRightBound = ws.Cells(HEADER_ROW_2, FIRST_TIMELINE_COL + totalDays - 1).Left + _
                         ws.Cells(HEADER_ROW_2, FIRST_TIMELINE_COL + totalDays - 1).Width - LINK_EDGE_PADDING

    yOut = ws.Cells(ganttRow, FIRST_TIMELINE_COL).Top + (ws.Rows(ganttRow).Height / 2)

    If hasChildren.Exists(wbs) Then
        Select Case UCase$(Trim$(anchorSide))
            Case "LEFT"
                xOut = TimelineLeft(ws, projectStart, startVal)
            Case "RIGHT"
                xOut = TimelineRightAfterFinish(ws, projectStart, finishVal)
            Case Else
                xOut = GetTaskMidX(ws, projectStart, startVal)
        End Select

    ElseIf durationVal <= 1 Then
        sizeVal = ws.Rows(ganttRow).Height - 6

        Select Case UCase$(Trim$(anchorSide))
            Case "LEFT"
                xOut = GetTaskMidX(ws, projectStart, startVal) - (sizeVal / 2)
            Case "RIGHT"
                xOut = GetTaskMidX(ws, projectStart, startVal) + (sizeVal / 2)
            Case Else
                xOut = GetTaskMidX(ws, projectStart, startVal)
        End Select

    Else
        Select Case UCase$(Trim$(anchorSide))
            Case "LEFT"
                xOut = TimelineLeft(ws, projectStart, startVal)
            Case "RIGHT"
                xOut = TimelineRightAfterFinish(ws, projectStart, finishVal)
            Case Else
                xOut = GetTaskMidX(ws, projectStart, startVal)
        End Select
    End If

    If xOut < timelineLeftBound Then xOut = timelineLeftBound
    If xOut > timelineRightBound Then xOut = timelineRightBound

End Sub

Public Sub Gantt_Clear_Test_State()

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim oldInternalWrite As Boolean
    Dim oldEvents As Boolean
    Dim oldScreenUpdating As Boolean

    On Error GoTo SafeExit

    oldInternalWrite = GetGanttInternalWrite()
    oldEvents = Application.EnableEvents
    oldScreenUpdating = Application.ScreenUpdating

    SetGanttInternalWrite True
    Application.EnableEvents = False
    Application.ScreenUpdating = False

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(GANTT_SHEET)
    On Error GoTo SafeExit

    If ws Is Nothing Then GoTo SafeExit

    lastRow = GetLastGanttRow(ws)
    If lastRow < FIRST_TASK_ROW Then GoTo SafeExit

    With ws.Range(ws.Cells(FIRST_TASK_ROW, COL_TEST_START), ws.Cells(lastRow, COL_TEST_START))
        .ClearContents
        .NumberFormat = "dd/mm/yyyy"
    End With

    With ws.Range(ws.Cells(FIRST_TASK_ROW, COL_TEST_FINISH), ws.Cells(lastRow, COL_TEST_FINISH))
        .ClearContents
        .NumberFormat = "dd/mm/yyyy"
    End With

    With ws.Range(ws.Cells(FIRST_TASK_ROW, COL_TEST_PROGRESS), ws.Cells(lastRow, COL_TEST_PROGRESS))
        .ClearContents
        .NumberFormat = "0%"
    End With

    GanttLive_ClearTestRenderRequest

SafeExit:
    Application.ScreenUpdating = oldScreenUpdating
    Application.EnableEvents = oldEvents
    SetGanttInternalWrite oldInternalWrite

End Sub

Private Sub DrawSummaryBar( _
    ByVal ws As Worksheet, _
    ByVal ganttRow As Long, _
    ByVal projectStart As Variant, _
    ByVal startVal As Variant, _
    ByVal finishVal As Variant, _
    ByVal isCritical As Boolean, _
    ByVal totalDays As Long, _
    ByVal shapeKey As String, _
    ByVal labelText As String, _
    ByVal hasDelta As Boolean, _
    ByVal isComplete As Boolean)

    Dim x1 As Double
    Dim x2 As Double
    Dim yTop As Double
    Dim yMid As Double
    Dim yBottom As Double
    Dim shp As Shape

    Dim displayText As String
    Dim maxLen As Long
    Dim txtShape As Shape
    Dim textLeft As Double
    Dim textTop As Double
    Dim lineColor As Long

    If Not HasValue(startVal) Then Exit Sub
    If Not HasValue(finishVal) Then Exit Sub

    'Strict date-column anchoring.
    x1 = TimelineLeft(ws, projectStart, startVal)
    x2 = TimelineRightAfterFinish(ws, projectStart, finishVal)

    If x2 <= x1 Then Exit Sub

    yTop = GetGanttSummaryTop(ws, ganttRow)
    yMid = GetGanttRowMid(ws, ganttRow)
    yBottom = GetGanttSummaryBottom(ws, ganttRow)

    lineColor = GetSummaryLineColor(isCritical, isComplete)

    Set shp = ws.Shapes.AddLine(x1, yMid, x2, yMid)
    shp.Name = shapeKey & "_H"
    ApplyGanttRenderLinePlacement shp
    shp.Line.Weight = 2.25
    If hasDelta Then
        shp.Line.ForeColor.RGB = RGB(255, 192, 0)
    Else
        shp.Line.ForeColor.RGB = lineColor
    End If

    Set shp = ws.Shapes.AddLine(x1, yTop, x1, yBottom)
    shp.Name = shapeKey & "_L"
    ApplyGanttRenderLinePlacement shp
    shp.Line.Weight = 2.25
    If hasDelta Then
        shp.Line.ForeColor.RGB = RGB(255, 192, 0)
    Else
        shp.Line.ForeColor.RGB = lineColor
    End If

    Set shp = ws.Shapes.AddLine(x2, yTop, x2, yBottom)
    shp.Name = shapeKey & "_R"
    ApplyGanttRenderLinePlacement shp
    shp.Line.Weight = 2.25
    If hasDelta Then
        shp.Line.ForeColor.RGB = RGB(255, 192, 0)
    Else
        shp.Line.ForeColor.RGB = lineColor
    End If

    maxLen = 28

    displayText = Trim$(labelText)
    If Len(displayText) > maxLen Then
        displayText = Left$(displayText, maxLen) & "..."
    End If

    textLeft = x2 + 8
    textTop = GetGanttRowTop(ws, ganttRow)

    Set txtShape = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, textLeft, textTop, 220, GetGanttRowHeight(ws, ganttRow))
    txtShape.Name = shapeKey & "_TXT"
    ApplyGanttRenderShapePlacement txtShape

    With txtShape
        .Line.Visible = msoFalse
        .Fill.Visible = msoFalse
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
        .TextFrame2.MarginLeft = 0
        .TextFrame2.MarginRight = 0
        .TextFrame2.MarginTop = 0
        .TextFrame2.MarginBottom = 0
    End With

    With txtShape.TextFrame2.TextRange
        .Text = displayText
        .Font.Size = 10.5
        .Font.Fill.ForeColor.RGB = RGB(80, 80, 80)
    End With

End Sub

'=====================================================
' DIRECT TIMELINE COLUMN MAPPING
'=====================================================

Private Function TimelineColumnFromDate( _
    ByVal ws As Worksheet, _
    ByVal anyDate As Variant) As Long

    Dim lastCol As Long
    Dim c As Long
    Dim targetDate As Date
    Dim headerVal As Variant

    If Not HasValue(anyDate) Then Exit Function

    targetDate = CDate(anyDate)

    lastCol = ws.Cells(HEADER_ROW_2, ws.Columns.Count).End(xlToLeft).Column
    If lastCol < FIRST_TIMELINE_COL Then Exit Function

    For c = FIRST_TIMELINE_COL To lastCol

        headerVal = ws.Cells(HEADER_ROW_2, c).value

        If IsWeekScaleMode() Then
            'In WEEK mode the header is text like W05, so direct date matching is not possible.
            'Fallback to deterministic slot index for week scale only.
            TimelineColumnFromDate = FIRST_TIMELINE_COL + GetTimelineSlotIndex(GetIsoWeekMonday(CDate(anyDate)), CDate(anyDate))
            Exit Function
        Else
            If IsDate(headerVal) Then
                If CLng(CDate(headerVal)) = CLng(targetDate) Then
                    TimelineColumnFromDate = c
                    Exit Function
                End If
            End If
        End If

    Next c

End Function

'=====================================================
' HELPER: TimelineColumnFromHeaderDate_Exact
'
' Rôle :
' - DAY mode : retrouve la colonne en lisant directement les dates construites en ligne HEADER_ROW_2
' - WEEK mode : conserve le mapping par slot ISO semaine
'
' Pourquoi :
' - le rendu doit être ancré sur les colonnes réelles du Gantt
' - pas sur une accumulation visuelle ou un calcul indirect de pixels
'=====================================================
Private Function TimelineColumnFromHeaderDate_Exact( _
    ByVal ws As Worksheet, _
    ByVal projectStart As Variant, _
    ByVal taskDate As Variant) As Long

    Dim targetDate As Date
    Dim targetCol As Long
    Dim lastCol As Long
    Dim c As Long
    Dim headerVal As Variant

    If ws Is Nothing Then Exit Function
    If Not HasValue(projectStart) Then Exit Function
    If Not HasValue(taskDate) Then Exit Function

    If IsWeekScaleMode() Then
        targetCol = GetTimelineTargetCol(projectStart, taskDate)
        If targetCol < FIRST_TIMELINE_COL Then Exit Function
        TimelineColumnFromHeaderDate_Exact = targetCol
        Exit Function
    End If

    targetDate = CDate(taskDate)

    lastCol = ws.Cells(HEADER_ROW_2, ws.Columns.Count).End(xlToLeft).Column
    If lastCol < FIRST_TIMELINE_COL Then Exit Function

    For c = FIRST_TIMELINE_COL To lastCol
        headerVal = ws.Cells(HEADER_ROW_2, c).value

        If IsDate(headerVal) Then
            If CLng(CDate(headerVal)) = CLng(targetDate) Then
                TimelineColumnFromHeaderDate_Exact = c
                Exit Function
            End If
        End If
    Next c

    'Fallback contrôlé : si la date n'est pas trouvée dans les headers,
    'on revient au mapping par offset, mais toujours vers une colonne entière.
    targetCol = GetTimelineTargetCol(projectStart, taskDate)
    If targetCol >= FIRST_TIMELINE_COL Then
        TimelineColumnFromHeaderDate_Exact = targetCol
    End If

End Function

Private Sub Gantt_AddConsoleMessage( _
    ByVal consoleMessages As Collection, _
    ByVal msgType As String, _
    ByVal frText As String, _
    ByVal enText As String)

    If consoleMessages Is Nothing Then Exit Sub

    CalcBridge_AddConsoleMessage consoleMessages, msgType, _
        BiMsg(frText, enText)

End Sub

Private Function IsLevelOfEffortTaskType( _
    ByRef dataArr As Variant, _
    ByVal mapWBS As Object, _
    ByVal dataRow As Long) As Boolean

    Dim taskTypeVal As String

    IsLevelOfEffortTaskType = False

    If mapWBS Is Nothing Then Exit Function
    If Not mapWBS.Exists("Task Type") Then Exit Function

    taskTypeVal = UCase$(Trim$(CStr(dataArr(dataRow, mapWBS("Task Type")))))
    taskTypeVal = Replace$(taskTypeVal, "-", " ")
    taskTypeVal = Replace$(taskTypeVal, "_", " ")

    Do While InStr(1, taskTypeVal, "  ", vbBinaryCompare) > 0
        taskTypeVal = Replace$(taskTypeVal, "  ", " ")
    Loop

    IsLevelOfEffortTaskType = _
        (taskTypeVal = "LOE") Or _
        (taskTypeVal = "LEVEL OF EFFORT") Or _
        (taskTypeVal = "LEVEL OF EFFORT TASK")

End Function

Public Function IsMilestoneTaskType( _
    ByRef dataArr As Variant, _
    ByVal mapWBS As Object, _
    ByVal dataRow As Long) As Boolean

    Dim taskTypeVal As String
    Dim deadlineVal As Variant
    Dim hasDeadline As Boolean

    If mapWBS Is Nothing Then Exit Function
    If Not mapWBS.Exists("Task Type") Then Exit Function

    taskTypeVal = UCase$(Trim$(CStr(dataArr(dataRow, mapWBS("Task Type")))))
    taskTypeVal = Replace$(taskTypeVal, "-", " ")
    taskTypeVal = Replace$(taskTypeVal, "_", " ")

    Do While InStr(1, taskTypeVal, "  ", vbBinaryCompare) > 0
        taskTypeVal = Replace$(taskTypeVal, "  ", " ")
    Loop

    IsMilestoneTaskType = _
        (taskTypeVal = "MILESTONE") Or _
        (taskTypeVal = "MS") Or _
        (taskTypeVal = "JALON")

End Function

Private Function GetLoEProgressFromToday( _
    ByVal startVal As Variant, _
    ByVal finishVal As Variant) As Double

    Dim durationVal As Double
    Dim rawProgress As Double

    GetLoEProgressFromToday = 0#

    If Not HasValue(startVal) Then Exit Function
    If Not HasValue(finishVal) Then Exit Function

    durationVal = CDbl(finishVal) - CDbl(startVal) + 1
    If durationVal <= 0 Then Exit Function

    rawProgress = (CDbl(Date) - CDbl(startVal) + 1) / durationVal

    If rawProgress < 0 Then
        GetLoEProgressFromToday = 0#
    ElseIf rawProgress > 1 Then
        GetLoEProgressFromToday = 1#
    Else
        GetLoEProgressFromToday = rawProgress
    End If

End Function

