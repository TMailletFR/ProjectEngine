Attribute VB_Name = "mod_ProjectTutorial"
Option Explicit

Private Const version As String = "5"
Private Const HIGHLIGHT_NAME As String = "PE_TUTORIAL_HIGHLIGHT"
Private Const HIGHLIGHT_MARKER As String = "ProjectEngine.TutorialHighlight.v1"
Private mActive As Boolean
Private mPresenting As Boolean
Private mStep As Long
Private mWindow As Excel.window
Private mOriginalSelection As Range
Private mOriginalSheet As Worksheet
Private mOriginalRow As Long
Private mOriginalColumn As Long
Private mOriginalZoom As Variant
Private mWasWatching As Boolean
Private mHighlightSheet As Worksheet
Private mLastDiagnostic As String
Private mClosing As Boolean
Private mCloseAt As Date
Private mCloseProcedure As String
Private mReturnWelcome As Boolean
Private mMode As String
Private mExpectedRun As Long
Private mBusinessAction As Boolean

Public Function ProjectTutorial_IsPassiveNavigation() As Boolean
    ProjectTutorial_IsPassiveNavigation = mPresenting And Not mBusinessAction
End Function

Public Function ProjectTutorial_Mode() As String
    ProjectTutorial_Mode = mMode
End Function

Public Function ProjectTutorial_IsEditableStep(ByVal stepNumber As Long) As Boolean
    If mMode <> "PRACTICE" Then Exit Function
    Select Case ProjectTutorial_StepKey(stepNumber)
        Case "WBS_STRUCTURE", "WBS_NAMES", "CONFIGURATION", "DATES", "DEPENDENCIES", _
             "SUMMARY_SELECTION", "WEIGHTS", "GANTT_CONTROLS", "GANTT_PATHS", _
             "SCENARIOS", "MONITORING", "CONSTRAINTS", "DEADLINE"
            ProjectTutorial_IsEditableStep = True
    End Select
End Function

Private Function ProjectTutorial_RibbonCommands(ByVal stepNumber As Long) As String
    Select Case ProjectTutorial_StepKey(stepNumber)
        Case "CALCULATION": ProjectTutorial_RibbonCommands = "pePlanning|peForced"
        Case "PROJECTIONS": ProjectTutorial_RibbonCommands = "peFull|peGantt|peSCurve"
        Case "GANTT_NAVIGATION": ProjectTutorial_RibbonCommands = "peFullTimeline|peSelectedTask"
        Case "SCENARIO_COMMIT": ProjectTutorial_RibbonCommands = "peScenario|peTest|peLock|peGanttReset"
        Case "SCURVE_BUILD": ProjectTutorial_RibbonCommands = "peSCurve"
        Case "GANTT_BUILD": ProjectTutorial_RibbonCommands = "peGantt"
        Case "GENERAL": ProjectTutorial_RibbonCommands = "peSettings|peImport|peTutorial|peFullTimeline|peSelectedTask|peClearHistory|peClearAck|peCleanDashboard|peResetPlanning|peFullReset"
        Case "FINANCE": ProjectTutorial_RibbonCommands = "peFinanceManage|peFinanceRefresh"
        Case "FINANCE_EXPENSES": ProjectTutorial_RibbonCommands = "peFinanceExpenses"
        Case "FINANCE_SUMMARY": ProjectTutorial_RibbonCommands = "peFinanceRefresh"
        Case "FINANCE_EXTRACT": ProjectTutorial_RibbonCommands = "peFinanceExtract"
    End Select
End Function

Public Sub ProjectTutorial_Home()
    ProjectTutorial_Cleanup
    mStep = 0
    mMode = vbNullString
    Ribbon_SetTutorialPresentation vbNullString, True
End Sub

Public Sub ProjectTutorial_SelectMode(ByVal mode As String)
    If mode <> "DISCOVERY" And mode <> "PRACTICE" Then Err.Raise 5
    ProjectTutorial_Home
    mMode = mode
End Sub

Public Function ProjectTutorial_ResumeStep() As Long
    Dim progress As Variant, keys As Variant, i As Long, preparation As Long
    progress = Settings_GetTutorialProgress(mMode)
    If progress(2) <> "IN_PROGRESS" And progress(2) <> "COMPLETED" Then Exit Function
    keys = ProjectTutorial_StepKeys()
    For i = 0 To UBound(keys)
        If CStr(progress(1)) = CStr(keys(i)) Then
            ProjectTutorial_ResumeStep = i + 1
            preparation = ProjectTutorial_ResumePreparationStep(i + 1)
            If preparation = 2 Then
                If MsgBox(ProjectTutorial_Text("RESUME.EMPTY"), vbQuestion + vbYesNo, ProjectTutorial_Text("TITLE")) = vbYes Then ProjectTutorial_ResumeStep = 2
            ElseIf preparation = 9 Then
                If MsgBox(ProjectTutorial_Text("RESUME.UNCONFIRMED"), vbQuestion + vbYesNo, ProjectTutorial_Text("TITLE")) = vbYes Then ProjectTutorial_ResumeStep = 9
            End If
            Exit Function
        End If
    Next i
    ' Unknown versions/removed steps explicitly return to the first preparation page.
    If Len(CStr(progress(1))) > 0 Then ProjectTutorial_ResumeStep = 1
End Function

Private Function ProjectTutorial_ResumePreparationStep(ByVal stepNumber As Long) As Long
    Dim result As Variant
    If mMode <> "PRACTICE" Or stepNumber < 9 Then Exit Function
    If Planning_WBSIsEmpty(True, False) Then
        ProjectTutorial_ResumePreparationStep = 2
        Exit Function
    End If
    result = WorkflowCommand_LastResult("Run_Planning_Update", "Run_Forced_Planning_Update")
    If CalcState_GetRunStatus() <> "OK" Or CLng(result(0)) = 0 Or Not CBool(result(3)) _
       Or (result(2) <> "SUCCESS" And result(2) <> "WARNING_NONBLOCKING") Then
        ProjectTutorial_ResumePreparationStep = 9
    End If
End Function

Public Function ProjectTutorial_Operation(ByVal stepNumber As Long) As String
    If mMode <> "PRACTICE" Then Exit Function
    Select Case ProjectTutorial_StepKey(stepNumber)
        Case "CALCULATION": ProjectTutorial_Operation = "pePlanning"
        Case "PROJECTIONS": ProjectTutorial_Operation = "peFull"
        Case "GANTT_BUILD": ProjectTutorial_Operation = "peGantt"
        Case "SCURVE_BUILD": ProjectTutorial_Operation = "peSCurve"
    End Select
End Function

Public Function ProjectTutorial_TryNext(ByVal stepNumber As Long) As Boolean
    Dim commandId As String, entry As String, result As Variant, equivalent As String
    If Not mActive Or mClosing Or Not ProjectTutorial_IsOwnerWindowActive() Then Exit Function
    commandId = ProjectTutorial_Operation(stepNumber)
    If Len(commandId) = 0 Then ProjectTutorial_TryNext = True: Exit Function
    entry = Ribbon_CommandProcedure(commandId)
    If commandId = "pePlanning" Then equivalent = Ribbon_CommandProcedure("peForced")
    result = WorkflowCommand_LastResult(entry, equivalent)
    If CLng(result(0)) > mExpectedRun And CBool(result(3)) Then
        If result(2) = "SUCCESS" Or result(2) = "WARNING_NONBLOCKING" Then
            ProjectTutorial_TryNext = True
            Exit Function
        End If
    End If
    If MsgBox(ProjectTutorial_Text("ACTION.CONFIRM") & vbCrLf & Ribbon_CommandLabel(commandId, Settings_GetGlobalDisplayLanguage()), _
        vbQuestion + vbYesNo + vbDefaultButton2, ProjectTutorial_Text("TITLE")) <> vbYes Then Exit Function
    If Not ProjectTutorial_IsOwnerWindowActive() Then Exit Function
    Ribbon_ExecuteCommand commandId, mWindow
    result = WorkflowCommand_LastResult(entry, equivalent)
    If CLng(result(0)) <= mExpectedRun Then Exit Function
    If Not CBool(result(3)) Then Exit Function
    ProjectTutorial_TryNext = (result(2) = "SUCCESS" Or result(2) = "WARNING_NONBLOCKING")
End Function

Public Sub ProjectTutorial_BeginBusinessAction()
    If Not mActive Then Exit Sub
    mBusinessAction = True
    ProjectTutorial_Cleanup
    GanttDrag_ReconcileWatchState
End Sub

Public Sub ProjectTutorial_EndBusinessAction()
    mBusinessAction = False
    If Not mActive Or mClosing Then Exit Sub
    If Not ProjectTutorial_IsOwnerWindowActive() Then
        ProjectTutorial_RequestClose
        Exit Sub
    End If
    If mStep > 0 Then Ribbon_SetTutorialPresentation ProjectTutorial_RibbonCommands(mStep), True
    frmProjectTutorial.RefreshCurrentPage
End Sub

Public Function ProjectTutorial_IsActive() As Boolean
    ProjectTutorial_IsActive = mActive
End Function

Public Function ProjectTutorial_IsOwnerWindowActive() As Boolean
    On Error GoTo NotOwner
    If mWindow Is Nothing Or Application.activeWindow Is Nothing Then Exit Function
    If Application.activeWindow.hwnd <> mWindow.hwnd Then Exit Function
    ProjectTutorial_IsOwnerWindowActive = (Application.activeWindow.ActiveSheet.parent Is ThisWorkbook)
NotOwner:
End Function

' Hide in the form event; unload once its event stack has returned to Excel.
Public Sub ProjectTutorial_RequestClose()
    On Error GoTo Failed
    If Not mActive Or mClosing Then Exit Sub
    PEConsole_WheelDetach frmProjectTutorial
    frmProjectTutorial.Hide
    ProjectTutorial_Cleanup
    Ribbon_SetTutorialPresentation vbNullString
    mClosing = True
    mCloseAt = Now + TimeSerial(0, 0, 1)
    mCloseProcedure = "'" & Replace(ThisWorkbook.name, "'", "''") & "'!ProjectTutorial_FinishClose"
    Application.OnTime mCloseAt, mCloseProcedure
    Exit Sub
Failed:
    mLastDiagnostic = "TUTORIAL_CLOSE_REQUEST:" & CStr(Err.number) & ":" & Err.description
    mClosing = False
    ' Keep the hidden form loaded rather than unload it inside its own handler.
    mActive = False
    ProjectTutorial_RestoreView
End Sub

Private Function ProjectTutorial_CancelClose() As Boolean
    ProjectTutorial_CancelClose = True
    If mClosing Then
        On Error Resume Next
        Application.OnTime mCloseAt, mCloseProcedure, Schedule:=False
        If Err.number <> 0 Then
            mLastDiagnostic = "TUTORIAL_CLOSE_CANCEL:" & CStr(Err.number) & ":" & Err.description
            Debug.Print mLastDiagnostic
            ProjectTutorial_CancelClose = False
        End If
        On Error GoTo 0
    End If
    If Not ProjectTutorial_CancelClose Then Exit Function
    mClosing = False
End Function

Public Sub ProjectTutorial_FinishClose()
    Dim returnWelcome As Boolean
    If Not mClosing Then Exit Sub
    mClosing = False
    returnWelcome = mReturnWelcome And ProjectTutorial_IsOwnerWindowActive()
    mReturnWelcome = False
    ProjectTutorial_Cleanup
    On Error Resume Next
    Unload frmProjectTutorial
    GanttDrag_ReconcileWatchState
    mActive = False
    Set mWindow = Nothing
    Set mOriginalSelection = Nothing
    Set mOriginalSheet = Nothing
    On Error GoTo 0
    Ribbon_WorksheetActivated
    If returnWelcome Then ProjectWelcome_ShowIfNeeded
End Sub

Public Sub ProjectTutorial_WorkbookDeactivated()
    If mBusinessAction Then
        ProjectTutorial_Cleanup
        Ribbon_SetTutorialPresentation vbNullString
        Exit Sub
    End If
    If mActive Then
        mReturnWelcome = False
        frmProjectTutorial.Hide
        ProjectTutorial_RequestClose
    End If
End Sub

Public Function ProjectTutorial_LastDiagnostic() As String
    ProjectTutorial_LastDiagnostic = mLastDiagnostic
End Function

Public Function ProjectTutorial_StepKeys() As Variant
    Dim definitions As Variant
    Dim keys() As Variant
    Dim i As Long
    definitions = ProjectTutorial_StepDefinitions()
    ReDim keys(0 To UBound(definitions))
    For i = 0 To UBound(definitions)
        keys(i) = definitions(i)(0)
    Next i
    ProjectTutorial_StepKeys = keys
End Function

Public Function ProjectTutorial_StepCount() As Long
    Dim definitions As Variant
    definitions = ProjectTutorial_StepDefinitions()
    ProjectTutorial_StepCount = UBound(definitions) + 1
End Function

Public Function ProjectTutorial_StepDefinitions() As Variant
    Dim definitions(0 To 32) As Variant
    ' StepKey, semantic target, sheet context: no image dependency.
    definitions(0) = Array("NAVIGATION", "WBS.STRUCTURE", "WBS")
    definitions(1) = Array("WBS_STRUCTURE", "WBS.IDENTITY", "WBS")
    definitions(2) = Array("WBS_NAMES", "WBS.TASK_NAME", "WBS")
    definitions(3) = Array("CONFIGURATION", "WBS.CALENDAR", "WBS")
    definitions(4) = Array("DATES", "WBS.DATES", "WBS")
    definitions(5) = Array("DEPENDENCIES", "WBS.PREDECESSORS_WBS", "WBS")
    definitions(6) = Array("SUMMARY_SELECTION", "WBS.SUMMARY_SELECTION", "WBS")
    definitions(7) = Array("WEIGHTS", "WBS.WEIGHT_PERCENT", "WBS")
    definitions(8) = Array("CALCULATION", "RIBBON.CALCULATION", "WBS")
    definitions(9) = Array("PROJECTIONS", "RIBBON.PROJECTIONS", "WBS")
    definitions(10) = Array("DATES_RESULT", "WBS.CALCULATED_DATES", "WBS")
    definitions(11) = Array("ANALYSIS", "WBS.ANALYTICS", "WBS")
    definitions(12) = Array("GANTT_BUILD", "RIBBON.GANTT", "WBS")
    definitions(13) = Array("GANTT", "GANTT.PLANNING", "GANTT")
    definitions(14) = Array("GANTT_CONTROLS", "GANTT.CONTROLS", "GANTT")
    definitions(15) = Array("GANTT_PATHS", "GANTT.CONTROLS", "GANTT")
    definitions(16) = Array("GANTT_NAVIGATION", "RIBBON.GANTT_NAVIGATION", "GANTT")
    definitions(17) = Array("SCENARIOS", "GANTT.TEST_INPUTS", "GANTT")
    definitions(18) = Array("SCENARIO_COMMIT", "RIBBON.SCENARIOS", "GANTT")
    definitions(19) = Array("SCURVE_BUILD", "RIBBON.SCURVE", "WBS")
    definitions(20) = Array("SCURVE", "SCURVE.CHART", "SCURVE")
    definitions(21) = Array("DASHBOARD_CAPTURE", "DASHBOARD.NEW_SNAPSHOT", "DASHBOARD")
    definitions(22) = Array("MONITORING", "DASHBOARD.SNAPSHOTS", "DASHBOARD")
    definitions(23) = Array("HISTORY", "EVENT_HISTORY.MESSAGES", "EVENT_HISTORY")
    definitions(24) = Array("GENERAL", "RIBBON.GENERAL", "WBS")
    definitions(25) = Array("DISPLAY", "SETTINGS.CONFIGURATION", "SETTINGS")
    definitions(26) = Array("CONSTRAINTS", "CONSTRAINTS.DATE_RULES", "CONSTRAINTS")
    definitions(27) = Array("DEADLINE", "CONSTRAINTS.DEADLINE", "CONSTRAINTS")
    definitions(28) = Array("FINANCE", "FINANCE.INPUTS", "FINANCE")
    definitions(29) = Array("FINANCE_EXPENSES", "FINANCE.EXPENSES", "FINANCE")
    definitions(30) = Array("FINANCE_SUMMARY", "FINANCE.SUMMARY", "FINANCE SUMMARY")
    definitions(31) = Array("FINANCE_EXTRACT", "RIBBON.FINANCE_EXTRACT", "FINANCE")
    definitions(32) = Array("FINISH", "NONE", "WBS")
    ProjectTutorial_StepDefinitions = definitions
End Function

Public Function ProjectTutorial_StepKey(ByVal stepNumber As Long) As String
    Dim keys As Variant
    keys = ProjectTutorial_StepKeys()
    If stepNumber < 1 Or stepNumber > UBound(keys) + 1 Then Err.Raise 5, "ProjectTutorial_StepKey", "Invalid tutorial step"
    ProjectTutorial_StepKey = CStr(keys(stepNumber - 1))
End Function

Public Function ProjectTutorial_Text(ByVal key As String, Optional ByVal language As String = "") As String
    Dim template As String
    Dim arguments As Object
    Dim bindings As Collection
    Dim binding As Variant
    Dim value As String
    If Len(language) = 0 Then language = Settings_GetGlobalDisplayLanguage()
    template = TextCatalog_Get("TUTORIAL." & key, language)
    Set arguments = TextCatalog_Arguments()
    Set bindings = ProjectTutorial_LabelBindings()
    For Each binding In bindings
        If InStr(1, template, "{" & CStr(binding(0)) & "}", vbBinaryCompare) > 0 Then
            Select Case CStr(binding(1))
                Case "RIBBON": value = Ribbon_CommandLabel(CStr(binding(2)), language)
                Case "WBS": value = SchemaColumnTitle("tbl_WBS", CStr(binding(2)), language)
                Case "CONSTRAINTS": value = SchemaColumnTitle("tbl_CONSTRAINTS", CStr(binding(2)), language)
                Case Else: value = TextCatalog_Get(CStr(binding(2)), language)
            End Select
            If ProjectTutorial_ButtonBinding(binding) Then value = UCase$(value)
            arguments(CStr(binding(0))) = Replace$(Replace$(value, vbCrLf, " "), vbLf, " ")
        End If
    Next binding
    ProjectTutorial_Text = Replace$(TextCatalog_Format("TUTORIAL." & key, language, arguments), "\n", vbCrLf)
End Function

Private Function ProjectTutorial_ButtonBinding(ByVal binding As Variant) As Boolean
    ProjectTutorial_ButtonBinding = (CStr(binding(1)) = "RIBBON" Or InStr("|NewSnapshot|Compare|Next|Previous|HomeButton|FinishButton|Confirm|Cancel|AddExpense|SaveExpense|DeleteExpense|AddFinancialProject|BulkExpenses|", "|" & CStr(binding(0)) & "|") > 0)
End Function

Public Function ProjectTutorial_ButtonLabels() As Object
    Dim labels As Object, bindings As Collection, binding As Variant, value As String, language As String
    Set labels = CreateObject("Scripting.Dictionary")
    language = Settings_GetGlobalDisplayLanguage()
    Set bindings = ProjectTutorial_LabelBindings()
    For Each binding In bindings
        If CStr(binding(1)) = "RIBBON" Or CStr(binding(0)) = "NewSnapshot" Or CStr(binding(0)) = "Compare" Then
            If CStr(binding(1)) = "RIBBON" Then
                value = Ribbon_CommandLabel(CStr(binding(2)), language)
            Else
                value = TextCatalog_Get(CStr(binding(2)), language)
            End If
            value = Replace$(Replace$(value, vbCrLf, " "), vbLf, " ")
            labels(UCase$(value)) = True
        End If
    Next binding
    Set ProjectTutorial_ButtonLabels = labels
End Function

Private Function ProjectTutorial_LabelBindings() As Collection
    Dim bindings As New Collection
    bindings.Add Array("UpdatePlanning", "RIBBON", "pePlanning")
    bindings.Add Array("FullUpdate", "RIBBON", "peFull")
    bindings.Add Array("UpdateGantt", "RIBBON", "peGantt")
    bindings.Add Array("UpdateSCurve", "RIBBON", "peSCurve")
    bindings.Add Array("ForcedUpdate", "RIBBON", "peForced")
    bindings.Add Array("Settings", "RIBBON", "peSettings")
    bindings.Add Array("Import", "RIBBON", "peImport")
    bindings.Add Array("Tutorial", "RIBBON", "peTutorial")
    bindings.Add Array("FullTimeline", "RIBBON", "peFullTimeline")
    bindings.Add Array("SelectedTask", "RIBBON", "peSelectedTask")
    bindings.Add Array("Lock", "RIBBON", "peLock")
    bindings.Add Array("GanttReset", "RIBBON", "peGanttReset")
    bindings.Add Array("Scenario", "RIBBON", "peScenario")
    bindings.Add Array("Test", "RIBBON", "peTest")
    bindings.Add Array("ResetPlanning", "RIBBON", "peResetPlanning")
    bindings.Add Array("FullReset", "RIBBON", "peFullReset")
    bindings.Add Array("ManageFinance", "RIBBON", "peFinanceManage")
    bindings.Add Array("RefreshFinance", "RIBBON", "peFinanceRefresh")
    bindings.Add Array("ManageExpenses", "RIBBON", "peFinanceExpenses")
    bindings.Add Array("ExtractFinance", "RIBBON", "peFinanceExtract")
    bindings.Add Array("CleanFinance", "RIBBON", "peFinanceClean")
    bindings.Add Array("AddFinancialProject", "CATALOG", "FINANCE.PROJECT_ADD")
    bindings.Add Array("Next", "CATALOG", "TUTORIAL.NEXT")
    bindings.Add Array("Previous", "CATALOG", "TUTORIAL.PREVIOUS")
    bindings.Add Array("HomeButton", "CATALOG", "TUTORIAL.HOME")
    bindings.Add Array("FinishButton", "CATALOG", "TUTORIAL.FINISH")
    bindings.Add Array("Confirm", "CATALOG", "FINANCE.CONFIRM")
    bindings.Add Array("Cancel", "CATALOG", "FINANCE.CANCEL")
    bindings.Add Array("AddExpense", "CATALOG", "FINANCE.ADD")
    bindings.Add Array("SaveExpense", "CATALOG", "FINANCE.MODIFY")
    bindings.Add Array("BulkExpenses", "CATALOG", "FINANCE.BULK")
    bindings.Add Array("DeleteExpense", "CATALOG", "FINANCE.DELETE")
    bindings.Add Array("TaskID", "WBS", "ID")
    bindings.Add Array("TaskName", "WBS", "TASK_NAME")
    bindings.Add Array("Cal", "WBS", "CAL")
    bindings.Add Array("TaskType", "WBS", "TASK_TYPE")
    bindings.Add Array("Predecessors", "WBS", "PREDECESSORS_WBS")
    bindings.Add Array("CalculatedStart", "WBS", "CALCULATED_START")
    bindings.Add Array("CalculatedFinish", "WBS", "CALCULATED_FINISH")
    bindings.Add Array("CalculatedDuration", "WBS", "CALCULATED_DURATION")
    bindings.Add Array("BaselineStart", "WBS", "BASELINE_START")
    bindings.Add Array("BaselineDuration", "WBS", "BASELINE_DURATION")
    bindings.Add Array("ActualStart", "WBS", "ACTUAL_START")
    bindings.Add Array("ActualFinish", "WBS", "ACTUAL_FINISH")
    bindings.Add Array("ForecastStart", "WBS", "FORECAST_START")
    bindings.Add Array("ForecastFinish", "WBS", "FORECAST_FINISH")
    bindings.Add Array("CriticalPath", "WBS", "CRITICAL_PATH")
    bindings.Add Array("TotalFloat", "WBS", "TOTAL_FLOAT")
    bindings.Add Array("FreeFloat", "WBS", "FREE_FLOAT")
    bindings.Add Array("Weight", "WBS", "WEIGHT_PERCENT")
    bindings.Add Array("Progress", "WBS", "PROGRESS_PERCENT")
    bindings.Add Array("Deadline", "CONSTRAINTS", "DEADLINE")
    bindings.Add Array("TestStart", "CATALOG", "GANTT.COLUMN.TEST_START")
    bindings.Add Array("TestFinish", "CATALOG", "GANTT.COLUMN.TEST_FINISH")
    bindings.Add Array("TestProgress", "CATALOG", "GANTT.COLUMN.TEST_PROGRESS")
    bindings.Add Array("Start", "CATALOG", "GANTT.COLUMN.START")
    bindings.Add Array("Finish", "CATALOG", "GANTT.COLUMN.FINISH")
    bindings.Add Array("View", "CATALOG", "GANTT.CONTROL.VIEW")
    bindings.Add Array("Scale", "CATALOG", "GANTT.CONTROL.SCALE")
    bindings.Add Array("ConstraintView", "CATALOG", "GANTT.CONTROL.CONSTRAINT")
    bindings.Add Array("Paths", "CATALOG", "GANTT.CONTROL.PATH")
    bindings.Add Array("MultiProject", "CATALOG", "GANTT.CONTROL.MULTI_PROJECT")
    bindings.Add Array("Baseline", "CATALOG", "SCURVE.SERIES.BASELINE")
    bindings.Add Array("Forecast", "CATALOG", "SCURVE.SERIES.FORECAST")
    bindings.Add Array("Actual", "CATALOG", "SCURVE.SERIES.ACTUAL")
    bindings.Add Array("NewSnapshot", "CATALOG", "DASHBOARD.COMMAND.SNAPSHOT")
    bindings.Add Array("Compare", "CATALOG", "DASHBOARD.COMMAND.COMPARE")
    bindings.Add Array("From", "CATALOG", "DASHBOARD.SNAPSHOT.FROM")
    bindings.Add Array("To", "CATALOG", "DASHBOARD.SNAPSHOT.TO")
    Set ProjectTutorial_LabelBindings = bindings
End Function

Public Sub ProjectTutorial_Show(Optional ByVal returnWelcome As Boolean = False)
    Dim progress As Variant
    Dim keys As Variant
    Dim i As Long
    Dim resumeStep As Long
    On Error GoTo Failed
    If mActive Then Exit Sub
    If Application.activeWindow Is Nothing Then Exit Sub
    If TypeName(Application.ActiveSheet) <> "Worksheet" Then Exit Sub
    If Not (Application.ActiveSheet.parent Is ThisWorkbook) Then Exit Sub
    If Not Application.EnableEvents Then
        mLastDiagnostic = "TUTORIAL_EVENTS_DISABLED"
        Exit Sub
    End If
    If Not ProjectTutorial_CancelClose() Then Exit Sub
    Unload frmProjectTutorial
    Set mWindow = Application.activeWindow
    Set mOriginalSheet = mWindow.ActiveSheet
    If TypeName(Application.Selection) = "Range" Then Set mOriginalSelection = Application.Selection
    mOriginalRow = mWindow.scrollRow
    mOriginalColumn = mWindow.scrollColumn
    mOriginalZoom = mWindow.zoom
    mWasWatching = GanttDrag_IsWatching()
    mActive = True
    mReturnWelcome = returnWelcome
    Load frmProjectTutorial
    frmProjectTutorial.Prepare
    If Not frmProjectTutorial.IsReadyToShow() Then GoTo CleanExit
    frmProjectTutorial.Show vbModeless
    Exit Sub
CleanExit:
    ProjectTutorial_CancelClose
    ProjectTutorial_Cleanup
    On Error Resume Next
    Unload frmProjectTutorial
    ProjectTutorial_RestoreView
    mActive = False
    mReturnWelcome = False
    Set mWindow = Nothing
    Set mOriginalSelection = Nothing
    Set mOriginalSheet = Nothing
    On Error GoTo 0
    Ribbon_WorksheetActivated
    Exit Sub
Failed:
    Debug.Print "ProjectTutorial_Show: " & CStr(Err.number) & ": " & Err.description
    Resume CleanExit
End Sub

Public Function ProjectTutorial_ResolveTarget(ByVal stepNumber As Long) As Range
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim key As String
    Dim tableKey As String
    Dim lastKey As String
    Dim firstColumn As Range
    Dim lastColumn As Range
    Dim rowCount As Long
    On Error GoTo missing
    Select Case ProjectTutorial_StepKey(stepNumber)
        Case "NAVIGATION": key = "ID": lastKey = "TASK_NAME"
        Case "WBS_STRUCTURE": key = "ID": lastKey = "WBS"
        Case "WBS_NAMES": key = "TASK_NAME"
        Case "CONFIGURATION": key = "CAL": lastKey = "TASK_TYPE"
        Case "SUMMARY_SELECTION": key = "TASK_TYPE": lastKey = "S"
        Case "DISPLAY"
            Set ProjectTutorial_ResolveTarget = Settings_TutorialConfigurationTarget()
            Exit Function
        Case "DATES": key = "BASELINE_START": lastKey = "FORECAST_FINISH"
        Case "DATES_RESULT": key = "CALCULATED_START": lastKey = "CALCULATED_DURATION"
        Case "DEPENDENCIES": key = "PREDECESSORS_WBS"
        Case "CONSTRAINTS": key = "START_CONSTRAINT_TYPE": lastKey = "FINISH_CONSTRAINT_DATE"
        Case "DEADLINE": key = "DEADLINE"
        Case "ANALYSIS": key = "CRITICAL_PATH": lastKey = "FREE_FLOAT_REX"
        Case "WEIGHTS": key = "WEIGHT_PERCENT": lastKey = "PROGRESS_PERCENT"
        Case "GANTT"
            Set ProjectTutorial_ResolveTarget = GanttLayout_TutorialTimelineTarget()
            Exit Function
        Case "SCENARIOS"
            Set ProjectTutorial_ResolveTarget = GanttLayout_TutorialTestTarget()
            Exit Function
        Case "GANTT_CONTROLS", "GANTT_PATHS"
            Set ProjectTutorial_ResolveTarget = GanttLayout_TutorialControlsTarget()
            Exit Function
        Case "DASHBOARD_CAPTURE"
            Set ProjectTutorial_ResolveTarget = Dashboard_TutorialNewSnapshotTarget()
            Exit Function
        Case "MONITORING"
            Set ProjectTutorial_ResolveTarget = Dashboard_TutorialSnapshotTarget()
            Exit Function
        Case "SCURVE"
            Set ProjectTutorial_ResolveTarget = SCurve_TutorialChartTarget()
            Exit Function
        Case "HISTORY"
            Set ProjectTutorial_ResolveTarget = EventHistory_TutorialMessagesTarget()
            Exit Function
        Case "FINANCE"
            Set ProjectTutorial_ResolveTarget = Finance_Table(FIN_TABLE).HeaderRowRange
            Exit Function
        Case "FINANCE_EXPENSES"
            Set tbl = Finance_Table(FIN_TABLE)
            Set ProjectTutorial_ResolveTarget = tbl.HeaderRowRange.cells(1, 1).Resize(1 + WorksheetFunction.Min(5, tbl.ListRows.count), 3)
            Exit Function
        Case "FINANCE_SUMMARY"
            Set ws = ThisWorkbook.Worksheets("FINANCE SUMMARY")
            If ws.ChartObjects.count > 0 Then
                Set ProjectTutorial_ResolveTarget = ws.Range(ws.ChartObjects(1).TopLeftCell, ws.ChartObjects(1).BottomRightCell)
            Else
                Set ProjectTutorial_ResolveTarget = ws.Range("A3:L12")
            End If
            Exit Function
        Case Else: Exit Function
    End Select
    tableKey = "tbl_WBS"
    Set ws = ThisWorkbook.Worksheets("WBS")
    If ProjectTutorial_StepKey(stepNumber) = "CONSTRAINTS" Or ProjectTutorial_StepKey(stepNumber) = "DEADLINE" Then
        tableKey = "tbl_CONSTRAINTS"
        Set ws = ThisWorkbook.Worksheets("CONSTRAINTS")
    End If
    Set tbl = ws.ListObjects(tableKey)
    If Len(lastKey) = 0 Then lastKey = key
    Set firstColumn = SchemaListColumn(tbl, tableKey, key).Range.cells(1, 1)
    Set lastColumn = SchemaListColumn(tbl, tableKey, lastKey).Range.cells(1, 1)
    rowCount = tbl.ListRows.count
    If mMode <> "PRACTICE" Then rowCount = 0
    If rowCount > 5 Then rowCount = 5
    If rowCount > 0 Then
        Set ProjectTutorial_ResolveTarget = ws.Range(firstColumn.offset(1, 0), lastColumn.offset(rowCount, 0))
    Else
        Set ProjectTutorial_ResolveTarget = ws.Range(firstColumn, lastColumn)
    End If
missing:
End Function

Public Function ProjectTutorial_ShowPage(ByVal stepNumber As Long, Optional ByVal persistProgress As Boolean = True, _
                                        Optional ByVal navigate As Boolean = True) As Boolean
    Dim target As Range
    Dim visibleTarget As Range
    Dim pane As Excel.pane
    Dim paneRange As Range
    Dim paneIndex As Long
    Dim eventsBefore As Boolean
    Dim updatingBefore As Boolean
    Dim shp As shape
    Dim occupied As shape
    Dim created As Boolean
    Dim statesCaptured As Boolean
    Dim definitions As Variant
    Dim contextSheet As Worksheet
    Dim operation As String, receipt As Variant
    If mPresenting Then Exit Function
    On Error GoTo Failed
    mLastDiagnostic = "SESSION_INACTIVE"
    If Not mActive Then Exit Function
    mLastDiagnostic = "WINDOW_MISSING"
    If Application.activeWindow Is Nothing Then Exit Function
    mLastDiagnostic = "WINDOW_NOT_OWNER"
    If Application.activeWindow.hwnd <> mWindow.hwnd Then Exit Function
    mLastDiagnostic = "WORKBOOK_NOT_OWNER"
    If Not (mWindow.ActiveSheet.parent Is ThisWorkbook) Then Exit Function
    mPresenting = True
    mLastDiagnostic = "TARGET_UNAVAILABLE"
    If mStep <> stepNumber Then
        operation = ProjectTutorial_Operation(stepNumber)
        If operation = "pePlanning" Then
            receipt = WorkflowCommand_LastResult(Ribbon_CommandProcedure(operation), Ribbon_CommandProcedure("peForced"))
        Else
            receipt = WorkflowCommand_LastResult(Ribbon_CommandProcedure(operation))
        End If
        mExpectedRun = CLng(receipt(0))
    End If
    mStep = stepNumber
    mLastDiagnostic = "PROGRESS"
    If persistProgress Then Settings_SetTutorialProgress version, ProjectTutorial_StepKey(mStep), "IN_PROGRESS", mMode
    mLastDiagnostic = "RESOLVE"
    definitions = ProjectTutorial_StepDefinitions()
    Set contextSheet = ThisWorkbook.Worksheets(CStr(definitions(mStep - 1)(2)))
    If navigate Then Ribbon_SetTutorialPresentation ProjectTutorial_RibbonCommands(mStep), True
    If contextSheet.visible <> xlSheetVisible Then GoTo CaptureFallback
    eventsBefore = Application.EnableEvents
    updatingBefore = Application.ScreenUpdating
    statesCaptured = True
    Application.EnableEvents = False
    Application.ScreenUpdating = True
    If navigate Then
        GanttDrag_PauseForLifecycle
        contextSheet.Activate
        mWindow.scrollColumn = 1
        mWindow.scrollRow = 1
    ElseIf Not (mWindow.ActiveSheet Is contextSheet) Then
        GoTo CaptureFallback
    End If
    Application.EnableEvents = eventsBefore
    Application.ScreenUpdating = updatingBefore
    Set target = ProjectTutorial_ResolveTarget(mStep)
    mLastDiagnostic = "TARGET_UNAVAILABLE"
    If target Is Nothing Then GoTo CaptureFallback
    If target.Worksheet.visible <> xlSheetVisible Then GoTo CaptureFallback
    If target.EntireRow.hidden Or target.EntireColumn.hidden Then GoTo CaptureFallback
    If target.MergeCells Then Set target = target.MergeArea
    If target.Worksheet.ProtectDrawingObjects Then GoTo CaptureFallback
    If Not mHighlightSheet Is Nothing Then
        If Not (mHighlightSheet Is target.Worksheet) Then
            If Not ProjectTutorial_Cleanup() Then GoTo CaptureFallback
        End If
    End If
    eventsBefore = Application.EnableEvents
    updatingBefore = Application.ScreenUpdating
    statesCaptured = True
    mLastDiagnostic = "NAVIGATE"
    Application.EnableEvents = False
    Application.ScreenUpdating = True
    If navigate Then
        target.Worksheet.Activate
        Application.Goto target.cells(1, 1), (target.Worksheet.name <> "FINANCE")
        If target.Worksheet.name = "FINANCE" Then mWindow.scrollRow = 1
        If target.Worksheet.name = "CONSTRAINTS" Then mWindow.scrollColumn = 1
    End If
    Application.EnableEvents = eventsBefore
    Application.ScreenUpdating = updatingBefore
    If navigate Then Ribbon_WorksheetActivated
    ' Let Office consume context invalidations only after the final navigation.
    DoEvents
    If Not mActive Then GoTo CaptureFallback
    mLastDiagnostic = "PANE_INTERSECTION"
    ' A single pane intersection avoids a rectangle spanning frozen panes.
    For paneIndex = 1 To mWindow.Panes.count
        Set pane = mWindow.Panes(paneIndex)
        mLastDiagnostic = "PANE_RANGE"
        Set paneRange = pane.VisibleRange
        mLastDiagnostic = "PANE_INTERSECT"
        Set visibleTarget = Application.Intersect(target, paneRange)
        If Not visibleTarget Is Nothing Then
            mLastDiagnostic = "PANE_AREAS"
            If visibleTarget.Areas.count = 1 Then
                If visibleTarget.cells.CountLarge = target.cells.CountLarge Then Exit For
            End If
        End If
        Set visibleTarget = Nothing
    Next paneIndex
    If visibleTarget Is Nothing Then GoTo CaptureFallback
    On Error Resume Next
    Set occupied = target.Worksheet.Shapes(HIGHLIGHT_NAME)
    On Error GoTo Failed
    If Not occupied Is Nothing Then
        If occupied.AlternativeText <> HIGHLIGHT_MARKER Then GoTo CaptureFallback
    End If
    Set mHighlightSheet = target.Worksheet
    mLastDiagnostic = "SHAPE_CREATE"
    If occupied Is Nothing Then
        Set shp = mHighlightSheet.Shapes.AddShape(msoShapeRectangle, visibleTarget.left, visibleTarget.top, _
            visibleTarget.width, visibleTarget.height)
        created = True
        shp.name = HIGHLIGHT_NAME
        shp.AlternativeText = HIGHLIGHT_MARKER
    Else
        Set shp = occupied
        shp.left = visibleTarget.left: shp.top = visibleTarget.top
        shp.width = visibleTarget.width: shp.height = visibleTarget.height
    End If
    shp.fill.visible = msoTrue
    shp.fill.ForeColor.RGB = RGB(255, 221, 0)
    mLastDiagnostic = "SHAPE_FORMAT"
    shp.fill.Transparency = 0.6
    shp.Line.ForeColor.RGB = RGB(205, 165, 0)
    shp.Line.weight = 1.25
    If ProjectTutorial_IsEditableStep(mStep) Then
        shp.fill.visible = msoFalse
        shp.Line.ForeColor.RGB = RGB(0, 155, 70)
        shp.Line.weight = 2
    End If
    shp.Placement = xlFreeFloating
    shp.ZOrder msoBringToFront
    ProjectTutorial_ShowPage = True
    mLastDiagnostic = vbNullString
    mPresenting = False
    Exit Function
CaptureFallback:
    ProjectTutorial_Cleanup
    If statesCaptured Then
        Application.EnableEvents = eventsBefore
        Application.ScreenUpdating = updatingBefore
    End If
    If navigate Then Ribbon_WorksheetActivated
    DoEvents
    mPresenting = False
    Exit Function
Failed:
    mLastDiagnostic = mLastDiagnostic & ":" & CStr(Err.number) & ":" & Err.description
    Debug.Print "ProjectTutorial_ShowPage: " & CStr(Err.number) & ": " & Err.description
    ' The object reference proves ownership before its reserved name is assigned.
    If created Then
        On Error Resume Next
        shp.Delete
        On Error GoTo 0
    End If
    If statesCaptured Then
        Application.EnableEvents = eventsBefore
        Application.ScreenUpdating = updatingBefore
    End If
    ProjectTutorial_Cleanup
    mPresenting = False
End Function

Public Sub ProjectTutorial_Complete()
    Settings_SetTutorialProgress version, ProjectTutorial_StepKey(ProjectTutorial_StepCount()), "COMPLETED", mMode
End Sub

Public Function ProjectTutorial_Cleanup() As Boolean
    Dim ws As Worksheet
    Dim remainingSheet As Worksheet
    Dim shp As shape
    ProjectTutorial_Cleanup = True
    For Each ws In ThisWorkbook.Worksheets
        Set shp = Nothing
        On Error Resume Next
        Set shp = ws.Shapes(HIGHLIGHT_NAME)
        On Error GoTo 0
        If Not shp Is Nothing Then
            If shp.AlternativeText = HIGHLIGHT_MARKER Then
                On Error Resume Next
                shp.Delete
                If Err.number <> 0 Then
                    mLastDiagnostic = "CLEANUP:" & CStr(Err.number) & ":" & Err.description
                    ProjectTutorial_Cleanup = False
                    Set remainingSheet = ws
                    Err.Clear
                End If
                On Error GoTo 0
            End If
        End If
    Next ws
    Set mHighlightSheet = remainingSheet
End Function

Public Function ProjectTutorial_BeforeSave() As Boolean
    If mClosing Then
        If Not ProjectTutorial_CancelClose() Then Exit Function
        mReturnWelcome = False
        mClosing = True
        ProjectTutorial_FinishClose
    End If
    ProjectTutorial_BeforeSave = ProjectTutorial_Cleanup()
End Function

Public Sub ProjectTutorial_AfterSave(ByVal Success As Boolean)
    If mActive And Not mClosing And mStep > 0 Then ProjectTutorial_ShowPage mStep, False, False
End Sub

Public Function ProjectTutorial_BeforeClose() As Boolean
    If Not ProjectTutorial_CancelClose() Then Exit Function
    If Not ProjectTutorial_Cleanup() Then Exit Function
    Ribbon_SetTutorialPresentation vbNullString
    If mActive Then
        PEConsole_WheelDetach frmProjectTutorial
        frmProjectTutorial.Hide
    End If
    mActive = False
    mReturnWelcome = False
    Set mWindow = Nothing
    Set mOriginalSelection = Nothing
    Set mOriginalSheet = Nothing
    ProjectTutorial_BeforeClose = True
End Function

Public Function ProjectTutorial_StopForReset() As Boolean
    If Not ProjectTutorial_BeforeClose() Then Exit Function
    On Error GoTo Failed
    Unload frmProjectTutorial
    ProjectTutorial_StopForReset = True
Failed:
End Function

Private Sub ProjectTutorial_RestoreView()
    Dim oldEvents As Boolean
    oldEvents = Application.EnableEvents
    On Error GoTo Finished
    If Not ProjectTutorial_IsOwnerWindowActive() Then Exit Sub
    Application.EnableEvents = False
    If mWindow Is Nothing Then GoTo Finished
    mWindow.Activate
    If Not mOriginalSheet Is Nothing Then mOriginalSheet.Activate
    If Not mOriginalSelection Is Nothing Then Application.Goto mOriginalSelection, False
    mWindow.zoom = mOriginalZoom
    mWindow.scrollRow = mOriginalRow
    mWindow.scrollColumn = mOriginalColumn
Finished:
    Application.EnableEvents = oldEvents
    If mWasWatching Then GanttDrag_ReconcileWatchState
    Ribbon_WorksheetActivated
End Sub






