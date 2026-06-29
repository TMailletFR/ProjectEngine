Option Explicit

Dim gAppEvents As clsAppEvents

Private Sub Workbook_Open()
    Init_AppEvents
    Settings_HydrateRuntimeState
    Ensure_WBS_Main_Buttons
    StartGanttDragWatchIfActive
End Sub

Private Sub Workbook_Activate()
    Init_AppEvents
    StartGanttDragWatchIfActive
End Sub

Private Sub Workbook_SheetActivate(ByVal Sh As Object)
    Init_AppEvents
    If Not Sh Is Nothing Then
        If CStr(Sh.Name) = "GANTT" Then
            Gantt_RepairPendingGeometryIfNeeded
            GanttDrag_StartWatch
        End If
    End If
End Sub

Private Sub Workbook_SheetDeactivate(ByVal Sh As Object)
    If Sh Is Nothing Then Exit Sub
    If CStr(Sh.Name) = "GANTT" Then GanttDrag_PauseForLifecycle
End Sub

Private Sub Workbook_Deactivate()
    GanttDrag_PauseForLifecycle
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    GanttDrag_StopWatch False
End Sub

Private Sub StartGanttDragWatchIfActive()

    On Error GoTo SafeExit

    If Application.ActiveSheet Is Nothing Then Exit Sub
    If Not (Application.ActiveSheet.Parent Is Me) Then Exit Sub
    If CStr(Application.ActiveSheet.Name) <> "GANTT" Then Exit Sub

    Gantt_RepairPendingGeometryIfNeeded
    GanttDrag_StartWatch

SafeExit:

End Sub

Public Sub Init_AppEvents()

    If gAppEvents Is Nothing Then
        Set gAppEvents = New clsAppEvents
    End If

    If gAppEvents.App Is Nothing Then
        Set gAppEvents.App = Application
    End If

End Sub









