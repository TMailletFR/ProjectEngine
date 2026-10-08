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
    If Not Ribbon_IsActiveOwnerWindow(control.Context, owner) Then Exit Sub
    If control.Id <> "peSettings" Then
        If Not WorkbookSchema_WritesAllowed() Then Exit Sub
    End If
    If Ribbon_IsGanttCommand(control.Id) Then
        returnedVal = (CStr(control.Context.ActiveSheet.Name) = "GANTT")
    ElseIf Ribbon_IsWBSCommand(control.Id) Then
        returnedVal = (CStr(control.Context.ActiveSheet.Name) = "WBS")
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

Private Function Ribbon_ContextVisible(ByVal sheetName As String) As Boolean
    Dim window As Excel.Window
    On Error GoTo SafeExit
    ' Avoid the group's borrowed Office Context during a sheet transition.
    ' Visibility is passive: use the active Excel window only after owner validation.
    Set window = Application.ActiveWindow
    If Not Ribbon_IsActiveOwnerWindow(window, ThisWorkbook) Then Exit Function
    Ribbon_ContextVisible = (CStr(window.ActiveSheet.Name) = sheetName)
SafeExit:
End Function

Public Sub PERibbon_OnAction(ByVal control As Office.IRibbonControl)
    Dim owner As Workbook
    Dim ribbonWindow As Object
    Dim entry As Variant

    On Error GoTo Failed
    Set owner = Ribbon_ControlWorkbook(control)
    mLastDiagnostic = "RIBBON_OWNER_REJECTED"
    If owner Is Nothing Then GoTo Rejected
    Set ribbonWindow = control.Context
    ' Ribbon and Excel can expose distinct COM wrappers for the same native window.
    ' Existing owners use ActiveSheet: require the same window AND workbook first.
    mLastDiagnostic = "RIBBON_WINDOW_NOT_ACTIVE"
    If Not Ribbon_IsActiveOwnerWindow(ribbonWindow, owner) Then GoTo Rejected
    If control.Id <> "peSettings" Then
        If Not WorkbookSchema_WritesAllowed() Then GoTo Rejected
    End If
    mLastDiagnostic = "RIBBON_CONTEXT_NOT_WORKSHEET"
    If TypeName(ribbonWindow.ActiveSheet) <> "Worksheet" Then GoTo Rejected
    entry = Ribbon_ControlEntry(control.Id)
    mLastDiagnostic = "RIBBON_UNKNOWN_CONTROL"
    If IsEmpty(entry) Then GoTo Rejected
    If Ribbon_IsGanttCommand(control.Id) Then
        mLastDiagnostic = "RIBBON_GANTT_CONTEXT_REQUIRED"
        If CStr(ribbonWindow.ActiveSheet.Name) <> "GANTT" Then GoTo Rejected
    ElseIf Ribbon_IsWBSCommand(control.Id) Then
        mLastDiagnostic = "RIBBON_WBS_CONTEXT_REQUIRED"
        If CStr(ribbonWindow.ActiveSheet.Name) <> "WBS" Then GoTo Rejected
    End If
    mLastDiagnostic = "RIBBON_BUSY"
    If mExecuting Then GoTo Rejected

    mLastDiagnostic = vbNullString
    mExecuting = True
    Select Case CStr(entry(3))
        Case "Run_Planning_Update": Run_Planning_Update
        Case "Run_Forced_Planning_Update": Run_Forced_Planning_Update
        Case "Run_Full_Update": Run_Full_Update
        Case "Run_Gantt_Update": Run_Gantt_Update
        Case "Run_SCurve_Update": Run_SCurve_Update
        Case "Reset_Planning": Reset_Planning
        Case "Settings": owner.Worksheets("SETTINGS").Activate
        Case "ProjectWelcome_ImportPrevious": ProjectWelcome_ImportPrevious
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
            GoTo Rejected
    End Select
    mLastAction = CStr(entry(3))
    mActions = mActions + 1
    mExecuting = False
    Exit Sub
Rejected:
    mRejected = mRejected + 1
    Exit Sub
Failed:
    mLastDiagnostic = "RIBBON_CALLBACK:" & CStr(Err.Number) & ":" & Err.Description
    Debug.Print mLastDiagnostic
    mExecuting = False
End Sub

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
                               "peGanttReset", "peScenario", "peTest", "peLock")
        mRibbon.InvalidateControl CStr(controlId)
    Next controlId
    Exit Sub
Failed:
    mLastDiagnostic = "RIBBON_CONTEXT_INVALIDATION:" & CStr(Err.Number) & ":" & Err.Description
End Sub

Private Sub Ribbon_TryInitialTab()
    If Not mOpenInitialized Or Not mInitialTabPending Then Exit Sub
    If mRibbon Is Nothing Then Exit Sub
    On Error GoTo Failed
    If Not Ribbon_IsActiveOwnerWindow(Application.ActiveWindow, ThisWorkbook) Then Exit Sub
    ' Consume the initialization request independently of explicit sheet transitions.
    mInitialTabPending = False
    Ribbon_ActivateOwnerTab
    Exit Sub
Failed:
    mInitialTabPending = False
    mLastDiagnostic = "RIBBON_INITIAL_TAB:" & CStr(Err.Number) & ":" & Err.Description
End Sub

Private Sub Ribbon_ActivateOwnerTab()
    Dim window As Excel.Window
    If mRibbon Is Nothing Then Exit Sub
    On Error GoTo Failed
    Set window = Application.ActiveWindow
    If Not Ribbon_IsActiveOwnerWindow(window, ThisWorkbook) Then Exit Sub
    mRibbon.ActivateTab "peTab"
    Exit Sub
Failed:
    mLastDiagnostic = "RIBBON_ACTIVATE_TAB:" & CStr(Err.Number) & ":" & Err.Description
End Sub

Private Function Ribbon_IsActiveOwnerWindow(ByVal window As Object, ByVal owner As Workbook) As Boolean
    Dim activeWindow As Excel.Window
    Dim activeOwner As Workbook
    On Error GoTo SafeExit
    If window Is Nothing Or owner Is Nothing Then Exit Function
    Set activeWindow = Application.ActiveWindow
    If activeWindow Is Nothing Then Exit Function
    If window.Hwnd <> activeWindow.Hwnd Then Exit Function
    Set activeOwner = activeWindow.ActiveSheet.Parent
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
                               "peNavigation", "peSettings", "peImport", "peFullTimeline", "peSelectedTask", "peGanttContext", _
                               "peGanttReset", "peScenario", "peTest", "peLock")
        mRibbon.InvalidateControl CStr(controlId)
    Next controlId
    mInvalidations = mInvalidations + 1
    Exit Sub
Failed:
    mLastDiagnostic = "RIBBON_INVALIDATION:" & CStr(Err.Number) & ":" & Err.Description
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
    Set ribbonWindow = control.Context
    If ribbonWindow Is Nothing Then Exit Function
    If TypeName(ribbonWindow) <> "Window" Then Exit Function
    Set contextSheet = ribbonWindow.ActiveSheet
    If contextSheet Is Nothing Then Exit Function
    Set candidate = contextSheet.Parent
    If Not candidate Is ThisWorkbook Then Exit Function
    If StrComp(candidate.FullName, ThisWorkbook.FullName, vbTextCompare) <> 0 Then Exit Function
    If StrComp(candidate.Name, ThisWorkbook.Name, vbBinaryCompare) <> 0 Then Exit Function
    Set Ribbon_ControlWorkbook = candidate
InvalidContext:
End Function

Private Function Ribbon_Text(ByVal control As Office.IRibbonControl, ByVal field As Long) As String
    Dim owner As Workbook
    Dim entry As Variant
    On Error GoTo Failed
    Set owner = Ribbon_ControlWorkbook(control)
    If owner Is Nothing Then Exit Function
    entry = Ribbon_ControlEntry(control.Id)
    If IsEmpty(entry) Then Exit Function
    If Len(CStr(entry(field))) = 0 Then Exit Function
    Ribbon_Text = TextCatalog_Get(CStr(entry(field)), Settings_GetGlobalDisplayLanguage())
    Exit Function
Failed:
    mLastDiagnostic = "RIBBON_TEXT:" & CStr(Err.Number) & ":" & Err.Description
End Function

' Single map: label, screentip, supertip, public command. No language selection here.
Private Function Ribbon_ControlEntry(ByVal controlId As String) As Variant
    Select Case controlId
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
        Case "peGanttContext": Ribbon_ControlEntry = Array("RIBBON.GANTT_CONTEXT.LABEL", "", "", "")
        Case "peGanttReset": Ribbon_ControlEntry = Array("GANTT.COMMAND.RESET", "RIBBON.GANTT_RESET.TIP", "", "GanttSimulation_ResetToNormal")
        Case "peScenario": Ribbon_ControlEntry = Array("GANTT.COMMAND.SCENARIO", "RIBBON.SCENARIO.TIP", "", "Run_Gantt_Scenario_Engine")
        Case "peTest": Ribbon_ControlEntry = Array("GANTT.COMMAND.TEST", "RIBBON.TEST.TIP", "", "Run_Gantt_Test_Engine")
        Case "peLock": Ribbon_ControlEntry = Array("GANTT.COMMAND.LOCK", "RIBBON.LOCK.TIP", "", "Run_Gantt_Lock_Changes")
        Case "peFullTimeline": Ribbon_ControlEntry = Array("NAV.FULL.LABEL", "NAV.FULL.TIP", "NAV.FULL.HELP", "NavigationFull")
        Case "peSelectedTask": Ribbon_ControlEntry = Array("NAV.TASK.LABEL", "NAV.TASK.TIP", "NAV.TASK.HELP", "NavigationTask")
    End Select
End Function
