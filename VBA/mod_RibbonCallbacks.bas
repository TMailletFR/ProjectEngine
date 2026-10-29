Attribute VB_Name = "mod_RibbonCallbacks"
Option Explicit

Private mRibbon As Office.IRibbonUI
Private mExecuting As Boolean
Private mLoads As Long
Private mActions As Long
Private mRejected As Long
Private mInvalidations As Long
Private mLastAction As String
Private mLastDiagnostic As String
Private mOpenInitialized As Boolean
Private mInitialTabPending As Boolean
Private mTutorialCommands As String
Private mTutorialPresentationActive As Boolean

Public Sub Ribbon_SetTutorialPresentation(ByVal commandIds As String, Optional ByVal restrictEmpty As Boolean = False)
    Dim enabled As Boolean
    enabled = restrictEmpty Or Len(commandIds) > 0
    If mTutorialCommands = commandIds And mTutorialPresentationActive = enabled Then Exit Sub
    mTutorialCommands = commandIds
    mTutorialPresentationActive = enabled
    Ribbon_InvalidateContext
End Sub

Private Function Ribbon_TutorialAllows(ByVal commandId As String) As Boolean
    Ribbon_TutorialAllows = True
    If Not mTutorialPresentationActive Then Exit Function
    If Not ProjectTutorial_IsActive() Then Exit Function
    If Not ProjectTutorial_IsOwnerWindowActive() Then Exit Function
    Ribbon_TutorialAllows = (InStr(1, "|" & mTutorialCommands & "|", "|" & commandId & "|", vbBinaryCompare) > 0)
End Function

Public Function Ribbon_CommandProcedure(ByVal commandId As String) As String
    Dim entry As Variant
    entry = Ribbon_ControlEntry(commandId)
    If Not IsEmpty(entry) Then Ribbon_CommandProcedure = CStr(entry(3))
End Function

Public Sub PERibbon_OnLoad(ByVal ribbon As Office.IRibbonUI)
    Set mRibbon = ribbon
    mLoads = mLoads + 1
    Ribbon_TryInitialTab
End Sub

Public Sub PERibbon_GetLabel(ByVal control As Office.IRibbonControl, ByRef returnedVal)
    returnedVal = Ribbon_Text(control, 0)
End Sub

Public Sub PERibbon_GetScreentip(ByVal control As Office.IRibbonControl, ByRef returnedVal)
    returnedVal = Ribbon_Text(control, 1)
End Sub

Public Sub PERibbon_GetSupertip(ByVal control As Office.IRibbonControl, ByRef returnedVal)
    returnedVal = Ribbon_Text(control, 2)
End Sub

Public Sub PERibbon_GetEnabled(ByVal control As Office.IRibbonControl, ByRef returnedVal)
    Dim owner As Workbook
    On Error GoTo SafeExit
    returnedVal = False
    Set owner = Ribbon_ControlWorkbook(control)
    If owner Is Nothing Then Exit Sub
    If Not Ribbon_IsActiveOwnerWindow(control.context, owner) Then Exit Sub
    If mExecuting Then Exit Sub
    If Not Ribbon_TutorialAllows(control.id) Then Exit Sub
    If control.id <> "peSettings" And control.id <> "peTutorial" Then
        If Not WorkbookSchema_WritesAllowed() Then Exit Sub
    End If
    If Ribbon_IsGanttCommand(control.id) Then
        returnedVal = (CStr(control.context.ActiveSheet.name) = "GANTT")
    ElseIf Ribbon_IsWBSCommand(control.id) Then
        returnedVal = (CStr(control.context.ActiveSheet.name) = "WBS")
    ElseIf left$(control.id, 9) = "peFinance" And control.id <> "peFinanceClean" Then
        returnedVal = Ribbon_ContextVisible("FINANCE")
        If control.id = "peFinanceExpenses" And returnedVal Then returnedVal = (Finance_ScopeId() <> 0)
    Else
        returnedVal = True
    End If
SafeExit:
End Sub

Public Sub PERibbon_GetGanttVisible(ByVal control As Office.IRibbonControl, ByRef returnedVal)
    returnedVal = Ribbon_ContextVisible("GANTT")
End Sub

Public Sub PERibbon_GetWBSVisible(ByVal control As Office.IRibbonControl, ByRef returnedVal)
    returnedVal = Ribbon_ContextVisible("WBS")
End Sub

Public Sub PERibbon_GetFinanceVisible(ByVal control As Office.IRibbonControl, ByRef returnedVal)
    returnedVal = Ribbon_ContextVisible("FINANCE")
End Sub

Private Function Ribbon_ContextVisible(ByVal sheetName As String) As Boolean
    Dim window As Excel.window
    On Error GoTo SafeExit
    ' Avoid the group's borrowed Office Context during a sheet transition.
    ' Visibility is passive: use the active Excel window only after owner validation.
    Set window = Application.activeWindow
    If Not Ribbon_IsActiveOwnerWindow(window, ThisWorkbook) Then Exit Function
    Ribbon_ContextVisible = (CStr(window.ActiveSheet.name) = sheetName)
    If sheetName = "FINANCE" And CStr(window.ActiveSheet.name) = "FINANCE SUMMARY" Then Ribbon_ContextVisible = True
    If sheetName = "FINANCE" And left$(CStr(window.ActiveSheet.name), 16) = "Extract Finance " Then Ribbon_ContextVisible = True
SafeExit:
End Function

Public Sub PERibbon_OnAction(ByVal control As Office.IRibbonControl)
    Ribbon_ExecuteCommand control.id, control.context
End Sub

Public Function Ribbon_ExecuteCommand(ByVal commandId As String, ByVal commandWindow As Object) As Boolean
    Dim owner As Workbook
    Dim ribbonWindow As Object
    Dim entry As Variant

    On Error GoTo Failed
    If commandWindow Is Nothing Then GoTo rejected
    Set owner = commandWindow.ActiveSheet.parent
    If Not (owner Is ThisWorkbook) Then GoTo rejected
    mLastDiagnostic = "RIBBON_OWNER_REJECTED"
    If owner Is Nothing Then GoTo rejected
    Set ribbonWindow = commandWindow
    ' Ribbon and Excel can expose distinct COM wrappers for the same native window.
    ' Existing owners use ActiveSheet: require the same window AND workbook first.
    mLastDiagnostic = "RIBBON_WINDOW_NOT_ACTIVE"
    If Not Ribbon_IsActiveOwnerWindow(ribbonWindow, owner) Then GoTo rejected
    mLastDiagnostic = "RIBBON_TUTORIAL_ACTIVE"
    If Not Ribbon_TutorialAllows(commandId) Then GoTo rejected
    If commandId <> "peSettings" And commandId <> "peTutorial" Then
        If Not WorkbookSchema_WritesAllowed() Then GoTo rejected
    End If
    mLastDiagnostic = "RIBBON_CONTEXT_NOT_WORKSHEET"
    If TypeName(ribbonWindow.ActiveSheet) <> "Worksheet" Then GoTo rejected
    entry = Ribbon_ControlEntry(commandId)
    mLastDiagnostic = "RIBBON_UNKNOWN_CONTROL"
    If IsEmpty(entry) Then GoTo rejected
    If left$(commandId, 9) = "peFinance" And commandId <> "peFinanceClean" Then
        If Not Ribbon_ContextVisible("FINANCE") Then GoTo rejected
    End If
    If Ribbon_IsGanttCommand(commandId) Then
        mLastDiagnostic = "RIBBON_GANTT_CONTEXT_REQUIRED"
        If CStr(ribbonWindow.ActiveSheet.name) <> "GANTT" Then GoTo rejected
    ElseIf Ribbon_IsWBSCommand(commandId) Then
        mLastDiagnostic = "RIBBON_WBS_CONTEXT_REQUIRED"
        If CStr(ribbonWindow.ActiveSheet.name) <> "WBS" Then GoTo rejected
    End If
    mLastDiagnostic = "RIBBON_BUSY"
    If mExecuting Then GoTo rejected

    mLastDiagnostic = vbNullString
    mExecuting = True
    If commandId <> "peTutorial" Then ProjectTutorial_BeginBusinessAction
    Select Case CStr(entry(3))
        Case "Finance_Manage": Finance_Manage
        Case "Finance_Refresh": Finance_Refresh
        Case "Finance_ExpensesCommand": Finance_ExpensesCommand
        Case "Finance_Extract": Finance_Extract
        Case "Finance_Clean": Finance_Clean
        Case "Run_Planning_Update": Run_Planning_Update
        Case "Run_Forced_Planning_Update": Run_Forced_Planning_Update
        Case "Run_Full_Update": Run_Full_Update
        Case "Run_Gantt_Update": Run_Gantt_Update
        Case "Run_SCurve_Update": Run_SCurve_Update
        Case "Reset_Planning": Reset_Planning
        Case "Settings": owner.Worksheets("SETTINGS").Activate
        Case "ProjectWelcome_ImportPrevious": ProjectWelcome_ImportPrevious
        Case "ProjectTutorial_Show": ProjectTutorial_Show
        Case "ClearPlanningEventHistory": ClearPlanningEventHistory
        Case "ClearPlanningWarningAcknowledgements": ClearPlanningWarningAcknowledgements
        Case "Reset_Dashboard": Reset_Dashboard
        Case "Armageddon": Armageddon
        Case "GanttSimulation_ResetToNormal": GanttSimulation_ResetToNormal
        Case "Run_Gantt_Scenario_Engine": Run_Gantt_Scenario_Engine
        Case "Run_Gantt_Test_Engine": Run_Gantt_Test_Engine
        Case "Run_Gantt_Lock_Changes": Run_Gantt_Lock_Changes
        Case "NavigationFull": GanttNavigation_Run ribbonWindow, GANTT_NAV_FULL
        Case "NavigationTask": GanttNavigation_Run ribbonWindow, GANTT_NAV_TASK
        Case Else
            mExecuting = False
            GoTo rejected
    End Select
    mLastAction = CStr(entry(3))
    mActions = mActions + 1
    mExecuting = False
    If commandId <> "peTutorial" Then ProjectTutorial_EndBusinessAction
    Ribbon_ExecuteCommand = True
    Ribbon_InvalidateContext
    Ribbon_ActivateOwnerTab
    Exit Function
rejected:
    mRejected = mRejected + 1
    Exit Function
Failed:
    mLastDiagnostic = "RIBBON_CALLBACK:" & CStr(Err.number) & ":" & Err.description
    Debug.Print mLastDiagnostic
    mExecuting = False
    ProjectTutorial_EndBusinessAction
    Ribbon_InvalidateContext
    If left$(commandId, 9) = "peFinance" Then
        CalcBridge_ShowSingleConsoleMessage "STOP", "COMMON.ERROR.PROCEDURE_INLINE", TextCatalog_Arguments("Procedure", "Finance", "Details", mLastDiagnostic)
    End If
End Function

Public Sub Ribbon_WorkbookOpened()
    mOpenInitialized = True
    mInitialTabPending = True
    Ribbon_TryInitialTab
End Sub

Public Sub Ribbon_WorkbookActivated()
    Ribbon_InvalidateContext
    Ribbon_TryInitialTab
End Sub

Public Sub Ribbon_WorksheetActivated()
    Ribbon_InvalidateContext
    Ribbon_ActivateOwnerTab
End Sub

Public Sub Ribbon_InvalidateContext()
    Dim controlId As Variant
    If mRibbon Is Nothing Then Exit Sub
    On Error GoTo Failed
    For Each controlId In Array("peGanttContext", "peUpdate", "pePlanning", "peForced", "peFull", "peGantt", "peSCurve", _
                               "peGanttReset", "peScenario", "peTest", "peLock", "peSettings", "peImport", "peTutorial", _
                               "peClearHistory", "peClearAck", "peCleanDashboard", "peResetPlanning", "peFullReset", _
                               "peFullTimeline", "peSelectedTask", "peFinanceContext", "peFinanceManage", "peFinanceRefresh", "peFinanceExpenses", "peFinanceExtract", "peFinanceClean")
        mRibbon.InvalidateControl CStr(controlId)
    Next controlId
    Exit Sub
Failed:
    mLastDiagnostic = "RIBBON_CONTEXT_INVALIDATION:" & CStr(Err.number) & ":" & Err.description
End Sub

Private Sub Ribbon_TryInitialTab()
    If Not mOpenInitialized Or Not mInitialTabPending Then Exit Sub
    If mRibbon Is Nothing Then Exit Sub
    On Error GoTo Failed
    If Not Ribbon_IsActiveOwnerWindow(Application.activeWindow, ThisWorkbook) Then Exit Sub
    ' Consume the initialization request independently of explicit sheet transitions.
    mInitialTabPending = False
    Ribbon_ActivateOwnerTab
    Exit Sub
Failed:
    mInitialTabPending = False
    mLastDiagnostic = "RIBBON_INITIAL_TAB:" & CStr(Err.number) & ":" & Err.description
End Sub

Private Sub Ribbon_ActivateOwnerTab()
    Dim window As Excel.window
    If mRibbon Is Nothing Then Exit Sub
    On Error GoTo Failed
    Set window = Application.activeWindow
    If Not Ribbon_IsActiveOwnerWindow(window, ThisWorkbook) Then Exit Sub
    mRibbon.ActivateTab "peTab"
    Exit Sub
Failed:
    mLastDiagnostic = "RIBBON_ACTIVATE_TAB:" & CStr(Err.number) & ":" & Err.description
End Sub

Private Function Ribbon_IsActiveOwnerWindow(ByVal window As Object, ByVal owner As Workbook) As Boolean
    Dim activeWindow As Excel.window
    Dim activeOwner As Workbook
    On Error GoTo SafeExit
    If window Is Nothing Or owner Is Nothing Then Exit Function
    Set activeWindow = Application.activeWindow
    If activeWindow Is Nothing Then Exit Function
    If window.hwnd <> activeWindow.hwnd Then Exit Function
    Set activeOwner = activeWindow.ActiveSheet.parent
    Ribbon_IsActiveOwnerWindow = (activeOwner Is owner)
SafeExit:
End Function

Private Function Ribbon_IsGanttCommand(ByVal controlId As String) As Boolean
    Select Case controlId
        Case "peGanttReset", "peScenario", "peTest", "peLock": Ribbon_IsGanttCommand = True
    End Select
End Function

Private Function Ribbon_IsWBSCommand(ByVal controlId As String) As Boolean
    Select Case controlId
        Case "pePlanning", "peForced", "peFull", "peGantt", "peSCurve": Ribbon_IsWBSCommand = True
    End Select
End Function

' Called only after the explicit global presentation-language click.
Public Sub Ribbon_InvalidateLanguage()
    Dim controlId As Variant
    If mRibbon Is Nothing Then Exit Sub
    On Error GoTo Failed
    For Each controlId In Array("peTab", "peUpdate", "pePlanning", "peForced", "peFull", "peGantt", "peSCurve", _
                               "pePlanningTools", "peResetPlanning", "peClearHistory", "peClearAck", "peCleanDashboard", "peFullReset", _
                               "peNavigation", "peSettings", "peImport", "peTutorial", "peFullTimeline", "peSelectedTask", "peGanttContext", _
                               "peGanttReset", "peScenario", "peTest", "peLock", _
                               "peFinanceContext", "peFinanceManage", "peFinanceRefresh", "peFinanceExpenses", "peFinanceExtract", "peFinanceClean")
        mRibbon.InvalidateControl CStr(controlId)
    Next controlId
    mInvalidations = mInvalidations + 1
    Exit Sub
Failed:
    mLastDiagnostic = "RIBBON_INVALIDATION:" & CStr(Err.number) & ":" & Err.description
    Debug.Print mLastDiagnostic
End Sub

' Passive technical observation; does not force loading or alter UI state.
Public Function Ribbon_State() As String
    Ribbon_State = "loads=" & CStr(mLoads) & ";actions=" & CStr(mActions) & _
        ";rejected=" & CStr(mRejected) & ";invalidations=" & CStr(mInvalidations) & _
        ";last=" & mLastAction & ";diagnostic=" & mLastDiagnostic
End Function

Private Function Ribbon_ControlWorkbook(ByVal control As Office.IRibbonControl) As Workbook
    Dim ribbonWindow As Object
    Dim contextSheet As Object
    Dim candidate As Workbook
    On Error GoTo InvalidContext
    If control Is Nothing Then Exit Function
    Set ribbonWindow = control.context
    If ribbonWindow Is Nothing Then Exit Function
    If TypeName(ribbonWindow) <> "Window" Then Exit Function
    Set contextSheet = ribbonWindow.ActiveSheet
    If contextSheet Is Nothing Then Exit Function
    Set candidate = contextSheet.parent
    If Not candidate Is ThisWorkbook Then Exit Function
    If StrComp(candidate.fullName, ThisWorkbook.fullName, vbTextCompare) <> 0 Then Exit Function
    If StrComp(candidate.name, ThisWorkbook.name, vbBinaryCompare) <> 0 Then Exit Function
    Set Ribbon_ControlWorkbook = candidate
InvalidContext:
End Function

Private Function Ribbon_Text(ByVal control As Office.IRibbonControl, ByVal field As Long) As String
    Dim owner As Workbook
    Dim entry As Variant
    On Error GoTo Failed
    Set owner = Ribbon_ControlWorkbook(control)
    If owner Is Nothing Then Exit Function
    entry = Ribbon_ControlEntry(control.id)
    If IsEmpty(entry) Then Exit Function
    If Len(CStr(entry(field))) = 0 Then Exit Function
    Ribbon_Text = TextCatalog_Get(CStr(entry(field)), Settings_GetGlobalDisplayLanguage())
    Exit Function
Failed:
    mLastDiagnostic = "RIBBON_TEXT:" & CStr(Err.number) & ":" & Err.description
End Function

' Single map: label, screentip, supertip, public command. No language selection here.
Private Function Ribbon_ControlEntry(ByVal controlId As String) As Variant
    Select Case controlId
        Case "peFinanceContext": Ribbon_ControlEntry = Array("FINANCE.GROUP", "", "", "")
        Case "peFinanceManage": Ribbon_ControlEntry = Array("FINANCE.MANAGE", "FINANCE.MANAGE", "", "Finance_Manage")
        Case "peFinanceRefresh": Ribbon_ControlEntry = Array("FINANCE.REFRESH", "FINANCE.REFRESH", "", "Finance_Refresh")
        Case "peFinanceExpenses": Ribbon_ControlEntry = Array("FINANCE.EXPENSES", "FINANCE.EXPENSES", "", "Finance_ExpensesCommand")
        Case "peFinanceExtract": Ribbon_ControlEntry = Array("FINANCE.EXTRACT", "FINANCE.EXTRACT", "", "Finance_Extract")
        Case "peFinanceClean": Ribbon_ControlEntry = Array("FINANCE.CLEAN", "FINANCE.CLEAN", "", "Finance_Clean")
        Case "peTab": Ribbon_ControlEntry = Array(TXT_RIBBON_TAB_LABEL, "", "", "")
        Case "peUpdate": Ribbon_ControlEntry = Array(TXT_RIBBON_UPDATE_LABEL, "", "", "")
        Case "pePlanning": Ribbon_ControlEntry = Array(TXT_RIBBON_PLANNING_LABEL, TXT_RIBBON_PLANNING_TIP, TXT_RIBBON_PLANNING_HELP, "Run_Planning_Update")
        Case "peForced": Ribbon_ControlEntry = Array("RIBBON.FORCED.LABEL", "RIBBON.FORCED.TIP", "", "Run_Forced_Planning_Update")
        Case "peFull": Ribbon_ControlEntry = Array(TXT_RIBBON_FULL_LABEL, TXT_RIBBON_FULL_TIP, TXT_RIBBON_FULL_HELP, "Run_Full_Update")
        Case "peGantt": Ribbon_ControlEntry = Array(TXT_RIBBON_GANTT_LABEL, TXT_RIBBON_GANTT_TIP, TXT_RIBBON_GANTT_HELP, "Run_Gantt_Update")
        Case "peSCurve": Ribbon_ControlEntry = Array(TXT_RIBBON_SCURVE_LABEL, TXT_RIBBON_SCURVE_TIP, TXT_RIBBON_SCURVE_HELP, "Run_SCurve_Update")
        Case "pePlanningTools": Ribbon_ControlEntry = Array("RIBBON.RESET_GROUP.LABEL", "", "", "")
        Case "peClearHistory": Ribbon_ControlEntry = Array("SETTINGS.CLEAR_HISTORY.LABEL", "RIBBON.CLEAR_HISTORY.TIP", "", "ClearPlanningEventHistory")
        Case "peClearAck": Ribbon_ControlEntry = Array("SETTINGS.CLEAR_ACK.LABEL", "RIBBON.CLEAR_ACK.TIP", "", "ClearPlanningWarningAcknowledgements")
        Case "peCleanDashboard": Ribbon_ControlEntry = Array("SETTINGS.CLEAN_DASHBOARD.LABEL", "RIBBON.CLEAN_DASHBOARD.TIP", "", "Reset_Dashboard")
        Case "peFullReset": Ribbon_ControlEntry = Array("SETTINGS.FULL_RESET.LABEL", "RIBBON.FULL_RESET.TIP", "", "Armageddon")
        Case "peResetPlanning": Ribbon_ControlEntry = Array(TXT_COMMON_RESET_PLANNING_LABEL, TXT_RIBBON_RESET_TIP, TXT_RIBBON_RESET_HELP, "Reset_Planning")
        Case "peNavigation": Ribbon_ControlEntry = Array("NAV.GROUP.LABEL", "", "", "")
        Case "peSettings": Ribbon_ControlEntry = Array("RIBBON.SETTINGS.LABEL", "RIBBON.SETTINGS.TIP", "", "Settings")
        Case "peImport": Ribbon_ControlEntry = Array("RIBBON.IMPORT.LABEL", "RIBBON.IMPORT.TIP", "", "ProjectWelcome_ImportPrevious")
        Case "peTutorial": Ribbon_ControlEntry = Array("TUTORIAL.LABEL", "TUTORIAL.TIP", "", "ProjectTutorial_Show")
        Case "peGanttContext": Ribbon_ControlEntry = Array("RIBBON.GANTT_CONTEXT.LABEL", "", "", "")
        Case "peGanttReset": Ribbon_ControlEntry = Array("GANTT.COMMAND.RESET", "RIBBON.GANTT_RESET.TIP", "", "GanttSimulation_ResetToNormal")
        Case "peScenario": Ribbon_ControlEntry = Array("GANTT.COMMAND.SCENARIO", "RIBBON.SCENARIO.TIP", "", "Run_Gantt_Scenario_Engine")
        Case "peTest": Ribbon_ControlEntry = Array("GANTT.COMMAND.TEST", "RIBBON.TEST.TIP", "", "Run_Gantt_Test_Engine")
        Case "peLock": Ribbon_ControlEntry = Array("GANTT.COMMAND.LOCK", "RIBBON.LOCK.TIP", "", "Run_Gantt_Lock_Changes")
        Case "peFullTimeline": Ribbon_ControlEntry = Array("NAV.FULL.LABEL", "NAV.FULL.TIP", "NAV.FULL.HELP", "NavigationFull")
        Case "peSelectedTask": Ribbon_ControlEntry = Array("NAV.TASK.LABEL", "NAV.TASK.TIP", "NAV.TASK.HELP", "NavigationTask")
    End Select
End Function

Public Function Ribbon_CommandLabel(ByVal controlId As String, ByVal language As String) As String
    Dim entry As Variant
    entry = Ribbon_ControlEntry(controlId)
    Ribbon_CommandLabel = TextCatalog_Get(CStr(entry(0)), language)
End Function




