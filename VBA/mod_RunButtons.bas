Attribute VB_Name = "mod_RunButtons"
Option Explicit

'=====================================================
' USER ORCHESTRATION BUTTONS
'
' Règle :
' - les boutons ne font pas de push direct vers WBS ;
' - Run_Calc_Engine pilote le bridge + état ;
' - le bridge se charge du sync, du calcul, puis du push contrôlé.
'
' Console routing:
' - les erreurs VBA des boutons sont envoyées dans frmPlanningMessages
' - aucun MsgBox direct dans ce module
'=====================================================

Private Sub RunButtons_ShowDeferredWorkflowConsole( _
    Optional ByVal extraProcName As String = "")

    Dim deferredMessages As Collection
    Dim finalDisplayStarted As Boolean

    On Error GoTo SafeExit

    Set deferredMessages = New Collection
    DrainPlanningWorkflowDeferredDisplayMessages deferredMessages

    If Trim$(extraProcName) <> "" Then
        RunButtons_AddConsoleError deferredMessages, extraProcName
    End If

    If deferredMessages.Count > 0 Then
        BeginPlanningWorkflowFinalDisplay
        finalDisplayStarted = True
        CalcBridge_ShowPlanningConsole deferredMessages
    End If

SafeExit:
    If finalDisplayStarted Then EndPlanningWorkflowFinalDisplay
    If Err.Number <> 0 Then Err.Raise Err.Number, Err.Source, Err.Description

End Sub

Private Sub RunButtons_AddConsoleError( _
    ByVal consoleMessages As Collection, _
    ByVal procName As String)

    If consoleMessages Is Nothing Then Exit Sub

    CalcBridge_AddConsoleMessage consoleMessages, _
        "STOP", _
        "FR:" & vbCrLf & _
        "Erreur VBA dans " & procName & vbCrLf & _
        "-> vérifier le dernier bloc modifié dans mod_RunButtons" & vbCrLf & vbCrLf & _
        "EN:" & vbCrLf & _
        "VBA error in " & procName & vbCrLf & _
        "-> check the last edited block in mod_RunButtons"

End Sub

Private Sub RunButtons_ShowConsoleError(ByVal procName As String)

    Dim consoleMessages As Collection

    Set consoleMessages = New Collection
    RunButtons_AddConsoleError consoleMessages, procName
    CalcBridge_ShowPlanningConsole consoleMessages

End Sub

Public Sub Run_Planning_Update()

    Dim workflowStarted As Boolean

    On Error GoTo SafeExit

    workflowStarted = EnsurePlanningWorkflowStarted("Run_Planning_Update")
    BeginPlanningEventRun "Run_Planning_Update"
    Run_Calc_Engine

CleanExit:
    If workflowStarted Then EndPlanningWorkflow
    Exit Sub

SafeExit:
    RunButtons_ShowConsoleError "Run_Planning_Update"
    Resume CleanExit

End Sub

Public Sub Run_Forced_Planning_Update()

    Dim workflowStarted As Boolean

    On Error GoTo SafeExit

    workflowStarted = EnsurePlanningWorkflowStarted("Run_Forced_Planning_Update")
    BeginPlanningEventRun "Run_Forced_Planning_Update"
    Run_Calc_Engine True

CleanExit:
    If workflowStarted Then EndPlanningWorkflow
    Exit Sub

SafeExit:
    RunButtons_ShowConsoleError "Run_Forced_Planning_Update"
    Resume CleanExit

End Sub

Public Sub Run_Gantt_Update()

    Dim wsCaller As Worksheet
    Dim workflowStarted As Boolean

    On Error GoTo SafeExit

    workflowStarted = EnsurePlanningWorkflowStarted("Run_Gantt_Update")
    Set wsCaller = ActiveSheet

    'User-facing full update from the big Gantt Update button.
    '
    'Important:
    '- This button must reset any temporary Gantt test state.
    '- Otherwise, after recalculating the planning, the Gantt can still compare
    '  against an old test render state and highlight deltas in yellow incorrectly.
    '- The cleanup is done BEFORE Run_Calc_Engine / Refresh_Gantt.
    '- Refresh_Gantt itself must stay generic because it is also used by the test workflow.

    Gantt_Clear_Test_State
    SetGanttPreserveTestInputs False
    GanttLive_ClearTestRenderRequest

    Run_Calc_Engine

    'Run_Calc_Engine_CoreBridge already displayed the real business error message.
    'Do not launch Refresh_Gantt after a blocking calculation error.
    If CalcEngine_HasBlockingErrorsForState() Then
        If Planning_WBSIsEmpty() Then
            Planning_GanttSafeEmptyState
        Else
            Gantt_SafeEmptyState
        End If
        If Not wsCaller Is Nothing Then wsCaller.Activate
        GoTo CleanExit
    End If

    If IsMacroAbortRequested() Then
        If Not wsCaller Is Nothing Then wsCaller.Activate
        GoTo CleanExit
    End If

    Refresh_Gantt

CleanExit:
    If workflowStarted Then EndPlanningWorkflow
    Exit Sub

SafeExit:
    On Error Resume Next
    If Not wsCaller Is Nothing Then wsCaller.Activate
    On Error GoTo 0

    RunButtons_ShowConsoleError "Run_Gantt_Update"
    Resume CleanExit

End Sub

Public Sub Run_SCurve_Update()

    Dim consoleMessages As Collection

    On Error GoTo ErrHandler

    Set consoleMessages = New Collection

    ThisWorkbook.Init_AppEvents
    BeginMacroRun "Run_SCurve_Update"
    Run_SCurve_Engine

    If IsMacroAbortRequested() Then GoTo SafeExit

    ThisWorkbook.Worksheets("SCURVE").Activate

SafeExit:
    If IsMacroAbortRequested() Then
        ShowAbortMessageOnce
    End If

    EndMacroRun
    Exit Sub

ErrHandler:
    If consoleMessages Is Nothing Then Set consoleMessages = New Collection

    CalcBridge_AddConsoleMessage consoleMessages, "STOP", _
        BiMsg( _
            "Erreur dans Run_SCurve_Update" & vbCrLf & _
            "-> " & Err.Description, _
            "Error in Run_SCurve_Update" & vbCrLf & _
            "-> " & Err.Description)

    CalcBridge_ShowPlanningConsole consoleMessages
    Resume SafeExit

End Sub


Public Sub Run_Full_Update()

    Dim perfScope As clsPerfScope

    Dim workflowStarted As Boolean
    Dim wsCaller As Worksheet
    Dim deferredConsoleShown As Boolean

    Set perfScope = Profiler_BeginScope("Run_Full_Update", "Workflow")

    On Error GoTo SafeExit

    Set wsCaller = ActiveSheet
    workflowStarted = EnsurePlanningWorkflowStarted("Run_Full_Update")
    BeginPlanningEventRun "Run_Full_Update"
    Gantt_Clear_Test_State
    SetGanttPreserveTestInputs False
    GanttLive_ClearTestRenderRequest
    GanttLive_ClearActiveSimulationMode
    Run_Calc_Engine True

    If CalcEngine_HasBlockingErrorsForState() Then
        If Planning_WBSIsEmpty() Then
            Planning_FullSafeEmptyState
        Else
            Gantt_SafeEmptyState
        End If
        GoTo CleanExit
    End If
    If IsMacroAbortRequested() Then GoTo CleanExit

    Refresh_Gantt False, False

    If IsMacroAbortRequested() Then GoTo CleanExit
    Run_SCurve_Engine

CleanExit:
    On Error Resume Next
    If Not wsCaller Is Nothing Then wsCaller.Activate
    On Error GoTo 0
    If workflowStarted And Not deferredConsoleShown Then RunButtons_ShowDeferredWorkflowConsole
    If workflowStarted Then EndPlanningWorkflow
    Exit Sub

SafeExit:
    If workflowStarted Then
        RunButtons_ShowDeferredWorkflowConsole "Run_Full_Update"
        deferredConsoleShown = True
    Else
        RunButtons_ShowConsoleError "Run_Full_Update"
    End If
    Resume CleanExit

End Sub







