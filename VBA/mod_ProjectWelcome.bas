Attribute VB_Name = "mod_ProjectWelcome"
Option Explicit

Private Const WELCOME_EVENT As String = "PROJECTENGINE_WELCOME"
Private mShowing As Boolean

Public Function ProjectWelcome_Message() As Object
    Dim messages As Collection, item As Object, message As String
    Set messages = New Collection
    message = "FR:" & vbCrLf & TextCatalog_Get("WELCOME.MESSAGE", "FR") & vbCrLf & _
              "EN:" & vbCrLf & TextCatalog_Get("WELCOME.MESSAGE", "EN")
    CalcBridge_AddConsoleMessage messages, "INFO", message, False, WELCOME_EVENT, _
        BuildPlanningEventIdentityV2("INFO", WELCOME_EVENT, "PROJECT", "WELCOME")
    Set item = messages(1)
    item("AcknowledgementRequired") = True
    Set ProjectWelcome_Message = item
End Function

Public Function ProjectWelcome_ShouldShow() As Boolean
    On Error GoTo NotEligible
    If Not WorkbookSchema_WritesAllowed() Or ThisWorkbook.ReadOnly Then Exit Function
    If Not Planning_IsNewOrFullyReset() Then Exit Function
    ProjectWelcome_ShouldShow = Not PlanningMessage_IsAcknowledged(ProjectWelcome_Message(), False)
NotEligible:
End Function

Public Sub ProjectWelcome_ShowIfNeeded()
    Dim messages As Collection, errorMessage As String, action As String, language As String
    If mShowing Or Not ProjectWelcome_ShouldShow() Then Exit Sub
    If Application.ActiveWorkbook Is Nothing Then Exit Sub
    If Not (Application.ActiveWorkbook Is ThisWorkbook) Then Exit Sub
    On Error GoTo CleanExit
    mShowing = True
    Set messages = New Collection
    messages.Add ProjectWelcome_Message()
    If Not PlanningEvents_LogConsoleMessagesSafe(messages, "ProjectWelcome", errorMessage) Then GoTo CleanExit
    If PlanningConsolePolicy_IsNonInteractive() Then
        PlanningConsolePolicy_CaptureDisplayMessages messages, "ProjectWelcome"
        GoTo CleanExit
    End If
    language = Settings_GetGlobalDisplayLanguage()
    Load frmProjectWelcome
    frmProjectWelcome.LoadWelcome TextCatalog_Get("WELCOME.TITLE", language), _
        TextCatalog_Get("WELCOME.MESSAGE", language), TextCatalog_Get("WELCOME.IMPORT", language), _
        TextCatalog_Get("WELCOME.NEW", language), TextCatalog_Get("CONSOLE.COMMAND.CLOSE", language)
    frmProjectWelcome.Show vbModal
    action = frmProjectWelcome.SelectedAction
    Unload frmProjectWelcome
    If action = "IMPORT" Then
        ProjectWelcome_ImportPrevious
    ElseIf action = "NEW" Then
        ProjectWelcome_StartNew
    End If
CleanExit:
    On Error Resume Next
    Unload frmProjectWelcome
    mShowing = False
End Sub

Public Sub ProjectWelcome_ImportPrevious()
    Migration_ImportPrevious
    ProjectWelcome_AcknowledgeImportSuccess
End Sub

Public Sub ProjectWelcome_AcknowledgeImportSuccess()
    If Migration_LastResult() <> "SUCCESS" Then Exit Sub
    ProjectWelcome_Acknowledge
End Sub

Public Sub ProjectWelcome_StartNew()
    If Not WorkbookSchema_UserActionAllowed() Then Exit Sub
    If Not Planning_IsNewOrFullyReset() Then Exit Sub
    ProjectWelcome_Acknowledge
    ThisWorkbook.Worksheets("WBS").Activate
    If Len(ThisWorkbook.Path) > 0 Then ThisWorkbook.Save
End Sub

Public Sub ProjectWelcome_Acknowledge()
    Dim item As Object
    Set item = ProjectWelcome_Message()
    If PlanningMessage_IsAcknowledged(item, False) Then Exit Sub
    SetPlanningWarningAckState item, True, True
    If Not PlanningMessage_IsAcknowledged(item) Then Err.Raise 5, "ProjectWelcome", "ACK_NOT_PERSISTED"
End Sub
