Attribute VB_Name = "mod_GanttNavigation"
Option Explicit

' Navigation owns the viewport, never scheduling or a second render lifecycle.
Public Const GANTT_NAV_FULL As String = "FULL_TIMELINE"
Public Const GANTT_NAV_TASK As String = "SELECTED_TASK"
Private Const FIRST_TASK_ROW As Long = 5
Private Const FIRST_TIMELINE_COL As Long = 11
Private Const TASK_MIN_VIEWPORT_RATIO As Double = 0.15
Private Const TASK_TARGET_VIEWPORT_RATIO As Double = 0.3
Private Const TASK_MAX_VIEWPORT_RATIO As Double = 0.5
Private Const MAX_FROZEN_VIEWPORT_RATIO As Double = 0.3
Private Const MILESTONE_CONTEXT_SLOTS As Double = 4#
Private Const FULL_FIT_ZOOM_SAFETY_PERCENT As Long = 1
Private mLastResult As String
Private mLastDiagnostic As String
Private mLastTaskId As String
Private mLastSourceMode As String

Public Sub GanttNavigation_Run(ByVal contextWindow As Excel.Window, ByVal command As String)
    Dim result As String
    Dim messages As Collection
    Dim presentationStep As String
    Dim identityArguments As Object
    Dim identityHash As String
    On Error GoTo PresentationFailed
    presentationStep = "Execute"
    result = GanttNavigation_Execute(contextWindow, command)
    If Left$(result, 4) = "NAV." And result <> "NAV.ERROR.INVALID_CONTEXT" Then
        If result = "NAV.ERROR.FAILED" And command = GANTT_NAV_FULL Then result = "NAV.ERROR.FAILED_FULL"
        presentationStep = "Format"
        Set messages = New Collection
        Set identityArguments = CreateObject("Scripting.Dictionary")
        identityArguments("Command") = command
        identityArguments("Source") = mLastSourceMode
        identityHash = BuildPlanningEventIdentityV2( _
            "WARNING", result, IIf(Len(mLastTaskId) > 0, "TASK", "GLOBAL"), _
            mLastTaskId, identityArguments)
        CalcBridge_AddConsoleMessage messages, "WARNING", _
            PlanningMessageText_Format(result, Nothing, Nothing), _
            eventType:=result, eventHash:=identityHash, _
            subjectTaskId:=mLastTaskId, sourceSheet:="GANTT"
        presentationStep = "ShowConsole"
        CalcBridge_ShowPlanningConsole messages
    End If
    Exit Sub
PresentationFailed:
    mLastDiagnostic = mLastDiagnostic & "|Presentation." & presentationStep & "|" & _
        CStr(Err.Number) & "|" & Err.Source & "|" & Err.Description
    Debug.Print "GanttNavigation:" & mLastDiagnostic
End Sub

Public Function GanttNavigation_Execute(ByVal contextWindow As Excel.Window, ByVal command As String) As String
    Dim ws As Worksheet, source As Worksheet, selected As Range, tbl As ListObject
    Dim data As Variant, columns As Object, taskId As String, wbsKey As String
    Dim target As Range, r As Long, rowIndex As Long, matches As Long
    Dim result As String
    Dim oldScrollRow As Long, oldScrollColumn As Long, oldZoom As Long
    Dim scope As clsPerfScope, stepName As String, targetVisible As Boolean
    Dim renderedSlots As Long, renderedRows As Long, firstScrollableColumn As Long
    Dim taskShape As Shape
    Dim ownerName As String, windowId As String, sourceMode As String

    On Error GoTo Failed
    mLastDiagnostic = vbNullString
    mLastTaskId = vbNullString
    mLastSourceMode = vbNullString
    stepName = "Context"
    result = "NAV.ERROR.INVALID_CONTEXT"
    If Not Navigation_ContextValid(contextWindow) Then GoTo Finished
    ownerName = ThisWorkbook.Name
    windowId = CStr(contextWindow.Hwnd)
    If command <> GANTT_NAV_FULL And command <> GANTT_NAV_TASK Then GoTo Finished
    Set scope = Profiler_BeginScope("GanttNavigation_Execute", "Navigation")
    Set source = contextWindow.ActiveSheet
    Set ws = ThisWorkbook.Worksheets("GANTT")
    Set tbl = ThisWorkbook.Worksheets("WBS").ListObjects("tbl_WBS")
    sourceMode = GanttEngine_ActiveDataSource()
    result = "NAV.ERROR.EMPTY"
    If tbl.DataBodyRange Is Nothing Then GoTo Finished

    stepName = "TaskIdentity"
    If command = GANTT_NAV_TASK Then
        result = "NAV.ERROR.SELECT_ONE_TASK"
        If TypeName(contextWindow.Selection) <> "Range" Then GoTo Finished
        Set selected = contextWindow.RangeSelection
        If selected.Areas.Count <> 1 Or selected.Rows.Count <> 1 Then GoTo Finished
        Set columns = CanonicalIdentity_BuildColumnMap(tbl)
        data = tbl.DataBodyRange.Value2
        If source Is tbl.Parent Then
            If Application.Intersect(selected, tbl.DataBodyRange) Is Nothing Then GoTo Finished
            rowIndex = selected.Row - tbl.DataBodyRange.Row + 1
            If rowIndex < 1 Or rowIndex > UBound(data, 1) Then GoTo Finished
            taskId = Trim$(CStr(data(rowIndex, columns(VTS_COL_ID))))
        ElseIf source Is ws Then
            If selected.Row < FIRST_TASK_ROW Then GoTo Finished
            wbsKey = Trim$(CStr(ws.Cells(selected.Row, 1).Value2))
            If Len(wbsKey) = 0 Then GoTo Finished
            For r = 1 To UBound(data, 1)
                If Trim$(CStr(data(r, columns(VTS_COL_WBS)))) = wbsKey Then
                    matches = matches + 1
                    taskId = Trim$(CStr(data(r, columns(VTS_COL_ID))))
                End If
            Next r
            If matches <> 1 Then GoTo Finished
        Else
            GoTo Finished
        End If
        If Len(taskId) = 0 Then GoTo Finished
        matches = 0
        For r = 1 To UBound(data, 1)
            If Trim$(CStr(data(r, columns(VTS_COL_ID)))) = taskId Then
                rowIndex = r
                matches = matches + 1
            End If
        Next r
        If matches <> 1 Then GoTo Finished
        Profiler_RecordOperation "NavigationTaskIdBulkResolutions", 1, 0#
        If GetGanttViewMode() = "SUMMARY" Then
            result = "NAV.ERROR.TASK_HIDDEN"
            If UCase$(Trim$(CStr(data(rowIndex, columns(VTS_COL_S))))) <> "Y" Then GoTo Finished
        End If
        ' A hidden row can invalidate geometry. Reject the identified hidden target
        ' before readiness repair, which may restore the renderer's row visibility.
        If CStr(ws.Cells(FIRST_TASK_ROW + rowIndex - 1, 1).Value2) = CStr(data(rowIndex, columns(VTS_COL_WBS))) Then
            result = "NAV.ERROR.TASK_HIDDEN"
            If ws.Rows(FIRST_TASK_ROW + rowIndex - 1).Hidden Then GoTo Finished
        End If
    End If

    stepName = "Readiness"
    result = "NAV.ERROR.NOT_READY"
    ' Do not enable events on behalf of a caller that deliberately disabled them.
    If Not Application.EnableEvents Then GoTo Finished
    If Not contextWindow.ActiveSheet Is ws Then ws.Activate
    If GanttDeferredRender_IsPending() Then
        GanttDeferredRender_FinishIfPending "Navigation"
        If GanttDeferredRender_IsPending() Then GoTo Finished
    End If
    GanttColdState_EnsureReadyIfGanttActive "Navigation"
    If UCase$(Trim$(CalcState_GetRunStatus())) <> "OK" Then
        mLastDiagnostic = "Planning CALC state is not current. Update Planning before navigation."
        result = "NAV.ERROR.PLANNING_STALE"
        GoTo Finished
    End If
    If Gantt_IsReadyStateValid() Then
        Profiler_RecordOperation "NavigationReadyRenderSkipped", 1, 0#
    Else
        If Not EnsureGanttForCurrentPlanning(GANTT_ENSURE_RENDER_AND_SHOW, "NavigationExplicit") Then
            mLastDiagnostic = GanttRefresh_LastErrorDescription()
            If Len(mLastDiagnostic) > 0 Then result = "NAV.ERROR.PREPARATION_FAILED"
            GoTo Finished
        End If
        If Not Gantt_IsReadyStateValid() Then GoTo Finished
        Profiler_RecordOperation "NavigationReadyPrepared", 1, 0#
    End If
    sourceMode = GanttEngine_ActiveDataSource()
    stepName = "RenderedTarget"
    If command = GANTT_NAV_TASK Then
        ' Row-index shape names are renderer-owned. Validate the displayed WBS identity
        ' before using them, including after WBS reorder or insertion.
        If CStr(ws.Cells(FIRST_TASK_ROW + rowIndex - 1, 1).Value2) <> CStr(data(rowIndex, columns(VTS_COL_WBS))) Then GoTo Finished
        result = "NAV.ERROR.TASK_HIDDEN"
        If ws.Rows(FIRST_TASK_ROW + rowIndex - 1).Hidden Then GoTo Finished
        Set taskShape = Navigation_TaskShape(ws, rowIndex)
        result = "NAV.ERROR.TASK_NOT_RENDERED"
        If taskShape Is Nothing Then GoTo Finished
        Set target = ws.Range(taskShape.TopLeftCell, taskShape.BottomRightCell)
    Else
        If Not GanttRefresh_GetRenderedExtent(renderedSlots, renderedRows) Then GoTo Finished
        firstScrollableColumn = 1
        If contextWindow.FreezePanes And contextWindow.SplitColumn > 0 Then _
            firstScrollableColumn = contextWindow.SplitColumn + 1
        If firstScrollableColumn < FIRST_TIMELINE_COL Then firstScrollableColumn = FIRST_TIMELINE_COL
        If firstScrollableColumn > FIRST_TIMELINE_COL + renderedSlots - 1 Then GoTo Finished
        Set target = ws.Range( _
            ws.Cells(FIRST_TASK_ROW, firstScrollableColumn), _
            ws.Cells(FIRST_TASK_ROW + renderedRows - 1, FIRST_TIMELINE_COL + renderedSlots - 1))
        result = "NAV.ERROR.EMPTY"
        If target Is Nothing Then GoTo Finished
    End If

    ' Activation has completed through the normal workbook lifecycle.
    stepName = "Viewport"
    stepName = "Viewport.ReadScroll"
    oldScrollRow = contextWindow.ScrollRow
    oldScrollColumn = contextWindow.ScrollColumn
    oldZoom = contextWindow.Zoom
    stepName = "Viewport.Visibility"
    targetVisible = Navigation_IsVisible(contextWindow, target)
    If Len(mLastDiagnostic) > 0 Then
        result = "NAV.ERROR.FAILED"
        GoTo Finished
    End If
    If command = GANTT_NAV_TASK Then
        stepName = "Viewport.Select"
        If contextWindow.RangeSelection.Address <> target.Cells(1, 1).Address Then target.Cells(1, 1).Select
        ' Selecting a boundary milestone can scroll Excel; frame only afterwards.
        stepName = "Viewport.TaskFrame"
        Navigation_FrameTask contextWindow, taskShape, FIRST_TASK_ROW + rowIndex - 1, _
            data, columns(VTS_COL_S), rowIndex
    Else
        contextWindow.ScrollRow = FIRST_TASK_ROW
        result = "NAV.ERROR.ZOOM_LIMIT"
        If Not Navigation_FitTimeline(contextWindow, target, firstScrollableColumn) Then
            If Len(mLastDiagnostic) > 0 Then
                result = "NAV.ERROR.FAILED"
                GoTo Finished
            End If
            stepName = "Viewport.FullMinimumZoom"
            If contextWindow.Zoom <> 10 Then contextWindow.Zoom = 10
            Application.Goto ws.Cells(FIRST_TASK_ROW, firstScrollableColumn), True
            If Application.ActiveWindow.Hwnd <> contextWindow.Hwnd Then _
                Err.Raise 5, "GanttNavigation", "Window changed during minimum-zoom framing"
            mLastDiagnostic = "FULL_TIMELINE_PARTIAL_AT_MINIMUM_ZOOM"
            Profiler_RecordOperation "NavigationFullMinimumZoomFallback", 1, 0#
        End If
    End If
    result = "OK"
    If oldScrollRow = contextWindow.ScrollRow And oldScrollColumn = contextWindow.ScrollColumn And oldZoom = contextWindow.Zoom Then
        Profiler_RecordOperation "NavigationViewportNoop", 1, 0#
    Else
        Profiler_RecordOperation "NavigationViewportMoved", 1, 0#
    End If
Finished:
    mLastResult = result
    mLastTaskId = taskId
    mLastSourceMode = sourceMode
    GanttNavigation_Execute = result
    Exit Function
Failed:
    mLastDiagnostic = CStr(Err.Number) & "|" & Err.Source & "|" & Err.Description & _
        "|" & stepName & "|" & taskId & "|" & command & "|" & sourceMode & _
        "|" & ownerName & "|" & windowId
    Debug.Print "GanttNavigation:" & mLastDiagnostic
    result = "NAV.ERROR.FAILED"
    Resume Finished
End Function

Public Function GanttNavigation_LastDiagnostic() As String
    GanttNavigation_LastDiagnostic = mLastDiagnostic
End Function

Public Function GanttNavigation_LastResult() As String
    GanttNavigation_LastResult = mLastResult
End Function

Private Function Navigation_ContextValid(ByVal wnd As Excel.Window) As Boolean
    Dim active As Excel.Window, book As Workbook
    On Error GoTo Invalid
    If wnd Is Nothing Then Exit Function
    Set active = Application.ActiveWindow
    If active Is Nothing Then Exit Function
    If active.Hwnd <> wnd.Hwnd Then Exit Function
    If TypeName(wnd.ActiveSheet) <> "Worksheet" Then Exit Function
    Set book = wnd.ActiveSheet.Parent
    If Not book Is ThisWorkbook Then Exit Function
    If book.FullName <> ThisWorkbook.FullName Then Exit Function
    Navigation_ContextValid = True
Invalid:
End Function

Private Function Navigation_TaskShape(ByVal ws As Worksheet, ByVal rowIndex As Long) As Shape
    Dim prefix As Variant, shp As Shape
    ' At most three canonical lookups, never a Shapes collection scan.
    For Each prefix In Array("TASK_", "MS_", "SUM_")
        Set shp = Nothing
        On Error Resume Next
        If prefix = "SUM_" Then
            Set shp = ws.Shapes(CStr(prefix) & CStr(rowIndex) & "_H")
        Else
            Set shp = ws.Shapes(CStr(prefix) & CStr(rowIndex))
        End If
        On Error GoTo 0
        If Not shp Is Nothing Then
            If shp.Visible Then Set Navigation_TaskShape = shp
            Exit Function
        End If
    Next prefix
End Function

Private Sub Navigation_FrameTask( _
    ByVal wnd As Excel.Window, ByVal shp As Shape, ByVal taskRow As Long, _
    ByRef wbsData As Variant, ByVal summaryColumn As Long, ByVal taskIndex As Long)
    Dim ws As Worksheet, zoom As Long, slotWidth As Double, rowHeight As Double
    Dim frozenWidth As Double, frozenHeight As Double, readableWidth As Double
    Dim viewWidth As Double, viewHeight As Double, scrollLeft As Double
    Dim scrollColumn As Long, scrollRow As Long, firstScrollableColumn As Long
    Dim visibleRowsBefore As Long, r As Long
    Set ws = shp.Parent
    slotWidth = ws.Cells(FIRST_TASK_ROW, FIRST_TIMELINE_COL).Width
    rowHeight = ws.Rows(taskRow).Height
    If slotWidth <= 0 Or rowHeight <= 0 Then Exit Sub
    If wnd.FreezePanes Then
        If wnd.SplitColumn > 0 Then frozenWidth = ws.Cells(1, 1).Resize(1, wnd.SplitColumn).Width
        If wnd.SplitRow > 0 Then frozenHeight = ws.Cells(1, 1).Resize(wnd.SplitRow, 1).Height
    End If
    readableWidth = shp.Width
    If Left$(shp.Name, 3) = "MS_" Then readableWidth = MILESTONE_CONTEXT_SLOTS * slotWidth
    zoom = Navigation_TaskZoom(readableWidth, frozenWidth, wnd.UsableWidth)
    If wnd.Zoom <> zoom Then
        wnd.Zoom = zoom
        Profiler_RecordOperation "NavigationTaskZoomChanges", 1, 0#
    End If
    viewWidth = wnd.UsableWidth * 100# / zoom - frozenWidth
    viewHeight = wnd.UsableHeight * 100# / zoom - frozenHeight
    scrollLeft = shp.Left + shp.Width / 2# - viewWidth / 2#
    scrollColumn = FIRST_TIMELINE_COL + Fix((scrollLeft - ws.Cells(FIRST_TASK_ROW, FIRST_TIMELINE_COL).Left) / slotWidth)
    firstScrollableColumn = FIRST_TIMELINE_COL
    If wnd.FreezePanes And wnd.SplitColumn + 1 > firstScrollableColumn Then _
        firstScrollableColumn = wnd.SplitColumn + 1
    If scrollColumn < firstScrollableColumn Then scrollColumn = firstScrollableColumn
    visibleRowsBefore = Fix((viewHeight / rowHeight - 1#) / 2#)
    If GetGanttViewMode() = "SUMMARY" Then
        scrollRow = taskRow
        For r = taskIndex - 1 To 1 Step -1
            If UCase$(Trim$(CStr(wbsData(r, summaryColumn)))) = "Y" Then
                scrollRow = FIRST_TASK_ROW + r - 1
                visibleRowsBefore = visibleRowsBefore - 1
                If visibleRowsBefore <= 0 Then Exit For
            End If
        Next r
    Else
        scrollRow = taskRow - visibleRowsBefore
    End If
    If scrollRow < FIRST_TASK_ROW Then scrollRow = FIRST_TASK_ROW
    Application.Goto ws.Cells(scrollRow, scrollColumn), True
    If Application.ActiveWindow.Hwnd <> wnd.Hwnd Then Err.Raise 5, "GanttNavigation", "Window changed during task framing"
    shp.TopLeftCell.Select
End Sub

Private Function Navigation_TaskZoom(ByVal readableWidth As Double, ByVal frozenWidth As Double, _
    ByVal usableWidth As Double) As Long
    Dim requestedZoom As Double
    If readableWidth <= 0 Or usableWidth <= 0 Then Err.Raise 5, "GanttNavigation", "Invalid viewport geometry"
    ' Solve width * zoom / (usable width - frozen width * zoom) = target ratio.
    requestedZoom = 100# * TASK_TARGET_VIEWPORT_RATIO * usableWidth / _
        (readableWidth + TASK_TARGET_VIEWPORT_RATIO * frozenWidth)
    If frozenWidth > 0 Then
        If requestedZoom > 100# * MAX_FROZEN_VIEWPORT_RATIO * usableWidth / frozenWidth Then _
            requestedZoom = 100# * MAX_FROZEN_VIEWPORT_RATIO * usableWidth / frozenWidth
    End If
    If requestedZoom < 10# Then requestedZoom = 10#
    If requestedZoom > 400# Then requestedZoom = 400#
    Navigation_TaskZoom = CLng(requestedZoom)
End Function

Private Function Navigation_FitTimeline( _
    ByVal wnd As Excel.Window, ByVal target As Range, _
    ByVal firstScrollableColumn As Long) As Boolean

    Dim pane As Excel.Pane, visible As Range
    Dim ws As Worksheet, frozenWidth As Double, frozenHeight As Double
    Dim widthZoom As Double, heightZoom As Double, fittedZoom As Long

    Set ws = target.Worksheet
    If wnd.FreezePanes Then
        If wnd.SplitColumn > 0 Then frozenWidth = ws.Cells(1, 1).Resize(1, wnd.SplitColumn).Width
        If wnd.SplitRow > 0 Then frozenHeight = ws.Cells(1, 1).Resize(wnd.SplitRow, 1).Height
    End If
    widthZoom = 100# * wnd.UsableWidth / (frozenWidth + target.Width)
    heightZoom = 100# * wnd.UsableHeight / (frozenHeight + target.Height)
    fittedZoom = Fix(widthZoom)
    If heightZoom < widthZoom Then fittedZoom = Fix(heightZoom)
    fittedZoom = fittedZoom - FULL_FIT_ZOOM_SAFETY_PERCENT
    If frozenWidth > 0 Then
        If fittedZoom > Fix(100# * MAX_FROZEN_VIEWPORT_RATIO * wnd.UsableWidth / frozenWidth) Then _
            fittedZoom = Fix(100# * MAX_FROZEN_VIEWPORT_RATIO * wnd.UsableWidth / frozenWidth)
    End If
    If fittedZoom < 10 Then Exit Function
    If fittedZoom > 400 Then fittedZoom = 400

    If wnd.Zoom <> fittedZoom Then wnd.Zoom = fittedZoom
    ' Goto with Scroll=True asks Excel to reposition the physical scrollable
    ' pane. ScrollColumn alone can report E while E:J remain unpainted.
    Application.Goto ws.Cells(FIRST_TASK_ROW, firstScrollableColumn), True
    If Application.ActiveWindow.Hwnd <> wnd.Hwnd Then Err.Raise 5, "GanttNavigation", "Window changed during full framing"
    DoEvents
    Set pane = wnd.Panes.Item(wnd.Panes.Count)
    Set visible = pane.VisibleRange
    Navigation_FitTimeline = _
        (pane.ScrollRow = FIRST_TASK_ROW) And _
        (pane.ScrollColumn = firstScrollableColumn) And _
        (visible.Row <= FIRST_TASK_ROW) And _
        (visible.Column <= firstScrollableColumn) And _
        (visible.Row + visible.Rows.Count - 1 >= target.Row + target.Rows.Count - 1) And _
        (visible.Column + visible.Columns.Count - 1 >= target.Column + target.Columns.Count - 1)
End Function

Private Function Navigation_IsVisible(ByVal wnd As Excel.Window, ByVal target As Range) As Boolean
    Dim pane As Excel.Pane, visible As Range, hit As Range
    Dim operation As String, paneIndex As Long
    On Error GoTo Failed
    operation = "Panes"
    ' Excel Panes supports indexed access; its COM enumerator is not available
    ' in every supported host. There are at most four worksheet panes.
    For paneIndex = 1 To wnd.Panes.Count
        Set pane = wnd.Panes.Item(paneIndex)
        operation = "VisibleRange"
        Set visible = pane.VisibleRange
        operation = "Intersect"
        Set hit = Application.Intersect(visible, target)
        If Not hit Is Nothing Then
            operation = "CountLarge"
            If hit.CountLarge = target.CountLarge Then
                ' VisibleRange includes boundary cells. Do not require a spare cell:
                ' this detects missing cells, not sub-cell clipping at minimum zoom.
                Navigation_IsVisible = True
                Exit Function
            End If
        End If
    Next paneIndex
    Exit Function
Failed:
    mLastDiagnostic = CStr(Err.Number) & "|Navigation_IsVisible." & operation & "|" & Err.Description
End Function
