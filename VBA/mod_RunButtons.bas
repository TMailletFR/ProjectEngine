Attribute VB_Name = "mod_RunButtons"
Option Explicit

'===============================================================================
' MODULE : mod_RunButtons
' DOMAINE / DOMAIN : Shared Infrastructure
'
' FR
' Expose les cinq macros Run_* et orchestre leurs workflows utilisateur sous MacroGuard.
' Ne doit pas contourner les contrats publics des autres domaines.
'
' EN
' Exposes the five Run_* macros and orchestrates their user workflows under MacroGuard.
' Must not bypass public contracts owned by other domains.
'
' CONTRATS / CONTRACTS : Run_Planning_Update, Run_Forced_Planning_Update, Run_Gantt_Update, Run_SCurve_Update, Run_Full_Update
' CALLBACKS EXTERNES / EXTERNAL CALLBACKS : Aucun / None
'===============================================================================


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

'------------------------------------------------------------------------------
' FR: Orchestre Run Buttons Show Deferred Workflow Console en preservant l'ordre contractuel des etapes du domaine.
' EN: Orchestrates Run Buttons Show Deferred Workflow Console while preserving the domain's contractual step order.
'------------------------------------------------------------------------------

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
        RunButtonsTrace_Checkpoint "Console", "Deferred console display start"
        BeginPlanningWorkflowFinalDisplay
        finalDisplayStarted = True
        CalcBridge_ShowPlanningConsole deferredMessages
        RunButtonsTrace_Checkpoint "Console", "Deferred console display returned"
    End If

SafeExit:
    If finalDisplayStarted Then EndPlanningWorkflowFinalDisplay
    If Err.Number <> 0 Then Err.Raise Err.Number, Err.Source, Err.Description

End Sub

'------------------------------------------------------------------------------
' FR: Orchestre Run Buttons Add Console Error en preservant l'ordre contractuel des etapes du domaine.
' EN: Orchestrates Run Buttons Add Console Error while preserving the domain's contractual step order.
'------------------------------------------------------------------------------

Private Sub RunButtons_AddConsoleError( _
    ByVal consoleMessages As Collection, _
    ByVal procName As String)

    If consoleMessages Is Nothing Then Exit Sub

    If procName = "Run_Gantt_Update" Then
        CalcBridge_AddConsoleMessage consoleMessages, _
            "STOP", _
            PlanningMessageText_Format("DIAG.RUN_BUTTONS.GANTT_UPDATE_ERROR", Nothing, Nothing)
        Exit Sub
    End If

    CalcBridge_AddConsoleMessage consoleMessages, _
        "STOP", _
        PlanningMessageText_Format("DIAG.RUN_BUTTONS.VBA_ERROR", _
            TextCatalog_Arguments("Procedure", procName), _
            TextCatalog_Arguments("Procedure", procName))

End Sub

'------------------------------------------------------------------------------
' FR: Orchestre Run Buttons Show Console Error en preservant l'ordre contractuel des etapes du domaine.
' EN: Orchestrates Run Buttons Show Console Error while preserving the domain's contractual step order.
'------------------------------------------------------------------------------

Private Sub RunButtons_ShowConsoleError(ByVal procName As String)

    Dim consoleMessages As Collection

    Set consoleMessages = New Collection
    RunButtons_AddConsoleError consoleMessages, procName
    RunButtonsTrace_Checkpoint "Console", "Error console display start: " & procName
    CalcBridge_ShowPlanningConsole consoleMessages
    RunButtonsTrace_Checkpoint "Console", "Error console display returned: " & procName

End Sub

'------------------------------------------------------------------------------
' FR: Lance le workflow Planning Update.
' EN: Runs the Planning Update workflow.
'------------------------------------------------------------------------------
Public Sub Run_Planning_Update()

    Dim workflowStarted As Boolean

    On Error GoTo SafeExit

    RunButtonsTrace_Checkpoint "RunButtons", "Enter Run_Planning_Update"
    workflowStarted = EnsurePlanningWorkflowStarted("Run_Planning_Update")
    RunButtonsTrace_Checkpoint "Workflow stack", "Workflow started Run_Planning_Update=" & CStr(workflowStarted)
    BeginPlanningEventRun "Run_Planning_Update"
    RunButtonsTrace_Checkpoint "EventHistory", "BeginPlanningEventRun returned Run_Planning_Update"
    RunButtonsTrace_Checkpoint "CoreBridge", "Run_Calc_Engine start Run_Planning_Update"
    Run_Calc_Engine
    RunButtonsTrace_Checkpoint "CoreBridge", "Run_Calc_Engine returned Run_Planning_Update"

    If Not CalcEngine_HasBlockingErrorsForState() And Not IsMacroAbortRequested() Then
        If Not Planning_WBSIsEmpty() Then
            If Not GanttTimeline_HasPhysicalHeader(ThisWorkbook.Worksheets("GANTT")) Then
                If Not EnsureGanttForCurrentPlanning( _
                    GANTT_ENSURE_RENDER_OFFSCREEN, "Run_Planning_UpdateBootstrap") Then
                    Err.Raise 5, "Run_Planning_Update", _
                        PlanningMessageText_Format("GANTT.ERROR.UPDATE_NOT_READY")
                End If
            End If
        End If
    End If

CleanExit:
    RunButtonsTrace_Checkpoint "Workflow stack", "CleanExit Run_Planning_Update"
    If workflowStarted Then EndPlanningWorkflow
    RunButtonsTrace_Checkpoint "RunButtons", "Exit Run_Planning_Update"
    Exit Sub

SafeExit:
    RunButtonsTrace_Checkpoint "RunButtons", "SafeExit Run_Planning_Update Err=" & CStr(Err.Number)
    RunButtons_ShowConsoleError "Run_Planning_Update"
    Resume CleanExit

End Sub

'------------------------------------------------------------------------------
' FR: Lance le workflow Forced Planning Update.
' EN: Runs the Forced Planning Update workflow.
'------------------------------------------------------------------------------
Public Sub Run_Forced_Planning_Update()

    Dim workflowStarted As Boolean

    On Error GoTo SafeExit

    RunButtonsTrace_Checkpoint "RunButtons", "Enter Run_Forced_Planning_Update"
    workflowStarted = EnsurePlanningWorkflowStarted("Run_Forced_Planning_Update")
    RunButtonsTrace_Checkpoint "Workflow stack", "Workflow started Run_Forced_Planning_Update=" & CStr(workflowStarted)
    BeginPlanningEventRun "Run_Forced_Planning_Update"
    RunButtonsTrace_Checkpoint "EventHistory", "BeginPlanningEventRun returned Run_Forced_Planning_Update"
    RunButtonsTrace_Checkpoint "CoreBridge", "Run_Calc_Engine force start"
    Run_Calc_Engine True
    RunButtonsTrace_Checkpoint "CoreBridge", "Run_Calc_Engine force returned"

CleanExit:
    RunButtonsTrace_Checkpoint "Workflow stack", "CleanExit Run_Forced_Planning_Update"
    If workflowStarted Then EndPlanningWorkflow
    RunButtonsTrace_Checkpoint "RunButtons", "Exit Run_Forced_Planning_Update"
    Exit Sub

SafeExit:
    RunButtonsTrace_Checkpoint "RunButtons", "SafeExit Run_Forced_Planning_Update Err=" & CStr(Err.Number)
    RunButtons_ShowConsoleError "Run_Forced_Planning_Update"
    Resume CleanExit

End Sub

'------------------------------------------------------------------------------
' FR: Lance le workflow Gantt Update.
' EN: Runs the Gantt Update workflow.
'------------------------------------------------------------------------------
Public Sub Run_Gantt_Update()

    Dim wsCaller As Worksheet
    Dim workflowStarted As Boolean
    Dim finalConsoleShown As Boolean
    Dim hadLocalSnapshot As Boolean

    On Error GoTo SafeExit

    RunButtonsTrace_Checkpoint "RunButtons", "Enter Run_Gantt_Update"
    workflowStarted = EnsurePlanningWorkflowStarted("Run_Gantt_Update")
    RunButtonsTrace_Checkpoint "Workflow stack", "Workflow started Run_Gantt_Update=" & CStr(workflowStarted)
    Set wsCaller = ActiveSheet
    hadLocalSnapshot = GanttLocal_HasCommittedSnapshot()
    GanttLocal_PrimeNormalState

    'User-facing full update from the big Gantt Update button.
    '
    'Important:
    '- This button must reset any temporary Gantt test state.
    '- Otherwise, after recalculating the planning, the Gantt can still compare
    '  against an old test render state and highlight deltas in yellow incorrectly.
    '- The cleanup is done BEFORE Run_Calc_Engine / Refresh_Gantt.
    '- Refresh_Gantt itself must stay generic because it is also used by the test workflow.

    GanttSimulation_ResetToNormal False

    RunButtonsTrace_Checkpoint "CoreBridge", "Run_Calc_Engine start Run_Gantt_Update"
    Run_Calc_Engine
    RunButtonsTrace_Checkpoint "CoreBridge", "Run_Calc_Engine returned Run_Gantt_Update"

    'Run_Calc_Engine_CoreBridge already displayed the real business error message.
    'Do not launch Refresh_Gantt after a blocking calculation error.
    If CalcEngine_HasBlockingErrorsForState() Then
        RunButtonsTrace_Checkpoint "CoreBridge", "Blocking errors detected Run_Gantt_Update"
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

    If Not hadLocalSnapshot Then
        If Not GanttDependency_PrimeLocalIndex(ThisWorkbook.Worksheets("GANTT"), True) Then
            GanttLocal_Invalidate "ColdDependencyRefreshFailed"
        End If
    End If

    RunButtonsTrace_Checkpoint "Gantt", "Ensure RENDER_AND_SHOW start Run_Gantt_Update"
    If Not EnsureGanttForCurrentPlanning(GANTT_ENSURE_RENDER_AND_SHOW, "Run_Gantt_Update") Then
        Err.Raise 5, "Run_Gantt_Update", PlanningMessageText_Format("GANTT.ERROR.UPDATE_NOT_READY")
    End If
    RunButtonsTrace_Checkpoint "Gantt", "Ensure RENDER_AND_SHOW returned Run_Gantt_Update"

CleanExit:
    RunButtonsTrace_Checkpoint "Workflow stack", "CleanExit Run_Gantt_Update"
    If workflowStarted And Not finalConsoleShown Then RunButtons_ShowDeferredWorkflowConsole
    If workflowStarted Then EndPlanningWorkflow
    RunButtonsTrace_Checkpoint "RunButtons", "Exit Run_Gantt_Update"
    Exit Sub

SafeExit:
    RunButtonsTrace_Checkpoint "RunButtons", "SafeExit Run_Gantt_Update Err=" & CStr(Err.Number)
    If workflowStarted Then
        RunButtons_ShowDeferredWorkflowConsole "Run_Gantt_Update"
        finalConsoleShown = True
    Else
        RunButtons_ShowConsoleError "Run_Gantt_Update"
    End If
    Resume CleanExit

End Sub

'------------------------------------------------------------------------------
' FR: Lance le workflow SCurve Update.
' EN: Runs the SCurve Update workflow.
'------------------------------------------------------------------------------
Public Sub Run_SCurve_Update()

    Dim consoleMessages As Collection

    On Error GoTo ErrHandler

    RunButtonsTrace_Checkpoint "RunButtons", "Enter Run_SCurve_Update"
    Set consoleMessages = New Collection

    AppEvents_EnsureInitialized
    RunButtonsTrace_Checkpoint "Workflow stack", "Init_AppEvents returned Run_SCurve_Update"
    BeginMacroRun "Run_SCurve_Update"
    RunButtonsTrace_Checkpoint "Workflow stack", "BeginMacroRun returned Run_SCurve_Update"
    RunButtonsTrace_Checkpoint "SCurve", "Run_SCurve_Engine start"
    Run_SCurve_Engine
    RunButtonsTrace_Checkpoint "SCurve", "Run_SCurve_Engine returned"

    If IsMacroAbortRequested() Then GoTo SafeExit

    RunButtonsTrace_Checkpoint "SCurve", "Activate SCURVE start"
    ThisWorkbook.Worksheets("SCURVE").Activate
    RunButtonsTrace_Checkpoint "SCurve", "Activate SCURVE returned"

SafeExit:
    RunButtonsTrace_Checkpoint "Workflow stack", "SafeExit Run_SCurve_Update"
    If IsMacroAbortRequested() Then
        ShowAbortMessageOnce
    End If

    EndMacroRun
    RunButtonsTrace_Checkpoint "RunButtons", "Exit Run_SCurve_Update"
    Exit Sub

ErrHandler:
    RunButtonsTrace_Checkpoint "RunButtons", "ErrHandler Run_SCurve_Update Err=" & CStr(Err.Number)
    If consoleMessages Is Nothing Then Set consoleMessages = New Collection

    CalcBridge_AddConsoleMessage consoleMessages, "STOP", _
        PlanningMessageText_Format("DIAG.SCURVE.UPDATE_ERROR", _
            TextCatalog_Arguments("Details", Err.Description), _
            TextCatalog_Arguments("Details", Err.Description))

    RunButtonsTrace_Checkpoint "Console", "Error console display start Run_SCurve_Update"
    CalcBridge_ShowPlanningConsole consoleMessages
    RunButtonsTrace_Checkpoint "Console", "Error console display returned Run_SCurve_Update"
    Resume SafeExit

End Sub


'------------------------------------------------------------------------------
' FR: Lance le workflow Full Update.
' EN: Runs the Full Update workflow.
'------------------------------------------------------------------------------
Public Sub Run_Full_Update(Optional ByVal propagateErrors As Boolean = False)

    Dim perfScope As clsPerfScope

    Dim workflowStarted As Boolean
    Dim wsCaller As Worksheet
    Dim deferredConsoleShown As Boolean

    Dim failureNumber As Long, failureSource As String, failureDescription As String

    Set perfScope = Profiler_BeginScope("Run_Full_Update", "Workflow")

    On Error GoTo SafeExit

    RunButtonsTrace_Checkpoint "RunButtons", "Enter Run_Full_Update"
    Set wsCaller = ActiveSheet
    workflowStarted = EnsurePlanningWorkflowStarted("Run_Full_Update")
    RunButtonsTrace_Checkpoint "Workflow stack", "Workflow started Run_Full_Update=" & CStr(workflowStarted)
    BeginPlanningEventRun "Run_Full_Update"
    RunButtonsTrace_Checkpoint "EventHistory", "BeginPlanningEventRun returned Run_Full_Update"
    GanttSimulation_ResetToNormal False
    RunButtonsTrace_Checkpoint "CoreBridge", "Run_Calc_Engine force start Run_Full_Update"
    Run_Calc_Engine True
    RunButtonsTrace_Checkpoint "CoreBridge", "Run_Calc_Engine force returned Run_Full_Update"

    If CalcEngine_HasBlockingErrorsForState() Then
        RunButtonsTrace_Checkpoint "CoreBridge", "Blocking errors detected Run_Full_Update"
        If Planning_WBSIsEmpty(propagateErrors) Then
            Planning_FullSafeEmptyState
        Else
            If propagateErrors Then Err.Raise 5, "Run_Full_Update", "PLANNING_VALIDATION_FAILED"
            Gantt_SafeEmptyState
        End If
        GoTo CleanExit
    End If
    If IsMacroAbortRequested() Then
        If propagateErrors Then Err.Raise 5, "Run_Full_Update", "PLANNING_ABORTED"
        GoTo CleanExit
    End If

    If Not wsCaller Is Nothing Then
        If UCase$(CStr(wsCaller.Name)) = "GANTT" Then
            RunButtonsTrace_Checkpoint "Gantt", "Ensure RENDER_AND_SHOW start Run_Full_Update"
            If Not EnsureGanttForCurrentPlanning(GANTT_ENSURE_RENDER_AND_SHOW, "Run_Full_Update") Then
        Err.Raise 5, "Run_Full_Update", PlanningMessageText_Format("GANTT.ERROR.FULL_UPDATE_GANTT_NOT_READY")
            End If
            RunButtonsTrace_Checkpoint "Gantt", "Ensure RENDER_AND_SHOW returned Run_Full_Update"
        Else
            RunButtonsTrace_Checkpoint "Gantt", "Ensure NO_RENDER start Run_Full_Update"
            If Not EnsureGanttForCurrentPlanning(GANTT_ENSURE_NO_RENDER, "Run_Full_UpdateDeferredNonGantt") Then
        Err.Raise 5, "Run_Full_Update", PlanningMessageText_Format("GANTT.ERROR.NO_RENDER_INVALIDATION")
            End If
            RunButtonsTrace_Checkpoint "Gantt", "Ensure NO_RENDER returned Run_Full_Update"
        End If
    Else
        RunButtonsTrace_Checkpoint "Gantt", "Ensure RENDER_OFFSCREEN start Run_Full_Update"
        If Not EnsureGanttForCurrentPlanning(GANTT_ENSURE_RENDER_OFFSCREEN, "Run_Full_UpdateNoCaller") Then
        Err.Raise 5, "Run_Full_Update", PlanningMessageText_Format("GANTT.ERROR.FULL_UPDATE_NOT_READY")
        End If
        RunButtonsTrace_Checkpoint "Gantt", "Ensure RENDER_OFFSCREEN returned Run_Full_Update"
    End If

    If IsMacroAbortRequested() Then
        If propagateErrors Then Err.Raise 5, "Run_Full_Update", "PLANNING_ABORTED"
        GoTo CleanExit
    End If
    RunButtonsTrace_Checkpoint "SCurve", "Run_SCurve_Engine start Run_Full_Update"
    Run_SCurve_Engine propagateErrors
    RunButtonsTrace_Checkpoint "SCurve", "Run_SCurve_Engine returned Run_Full_Update"

CleanExit:
    RunButtonsTrace_Checkpoint "Workflow stack", "CleanExit Run_Full_Update"
    On Error Resume Next
    If Not wsCaller Is Nothing Then wsCaller.Activate
    On Error GoTo 0
    If workflowStarted And Not deferredConsoleShown Then RunButtons_ShowDeferredWorkflowConsole
    If workflowStarted Then EndPlanningWorkflow
    RunButtonsTrace_Checkpoint "RunButtons", "Exit Run_Full_Update"
    If propagateErrors And failureNumber <> 0 Then
        On Error GoTo 0
        Err.Raise failureNumber, failureSource, failureDescription
    End If
    Exit Sub

SafeExit:
    failureNumber = Err.Number: failureSource = Err.Source: failureDescription = Err.Description
    If propagateErrors Then Resume CleanExit
    RunButtonsTrace_Checkpoint "RunButtons", "SafeExit Run_Full_Update Err=" & CStr(Err.Number)
    If workflowStarted Then
        RunButtons_ShowDeferredWorkflowConsole "Run_Full_Update"
        deferredConsoleShown = True
    Else
        RunButtons_ShowConsoleError "Run_Full_Update"
    End If
    Resume CleanExit

End Sub







