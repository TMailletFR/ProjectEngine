Attribute VB_Name = "mod_ProjectWelcome"
Option Explicit

Private Const WELCOME_EVENT As String = "PROJECTENGINE_WELCOME"
Private mShowing As Boolean
Private mTutorialVisitPending As Boolean
Private mDraftStarted As Boolean

Public Sub ProjectWelcome_ResetTutorialVisit()
    mTutorialVisitPending = False
    mDraftStarted = False
End Sub

Public Sub ProjectWelcome_ObserveTutorialEdit(ByVal Sh As Object, ByVal Target As Range)
    Dim body As Range
    If mDraftStarted Or Not mTutorialVisitPending Then Exit Sub
    If Not ProjectTutorial_IsActive() Or ProjectTutorial_Mode() <> "PRACTICE" Then Exit Sub
    If Sh Is Nothing Or Target Is Nothing Then Exit Sub
    If Not (Sh.parent Is ThisWorkbook) Or CStr(Sh.name) <> "WBS" Then Exit Sub
    Set body = Sh.ListObjects("tbl_WBS").DataBodyRange
    If body Is Nothing Then Exit Sub
    If Not Application.Intersect(body, Target) Is Nothing Then mDraftStarted = True
End Sub

Public Function ProjectWelcome_CanContinue() As Boolean
    ProjectWelcome_CanContinue = mTutorialVisitPending And mDraftStarted
End Function

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
    If Not mTutorialVisitPending Then
        If Not Planning_IsNewOrFullyReset() Then Exit Function
    End If
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
    Do
        language = Settings_GetGlobalDisplayLanguage()
        Load frmProjectWelcome
        If ProjectWelcome_CanContinue() Then
            errorMessage = TextCatalog_Get("WELCOME.MESSAGE.DRAFT", language)
        Else
            errorMessage = TextCatalog_Get("WELCOME.MESSAGE", language)
        End If
        frmProjectWelcome.LoadWelcome TextCatalog_Get("WELCOME.TITLE", language), _
            errorMessage, TextCatalog_Get("WELCOME.IMPORT", language), _
            TextCatalog_Get("WELCOME.NEW", language), TextCatalog_Get("CONSOLE.COMMAND.CLOSE", language), _
            ProjectWelcome_CanContinue()
        frmProjectWelcome.Show vbModal
        action = frmProjectWelcome.SelectedAction
        Unload frmProjectWelcome
        If action = "TUTORIAL" Then
            mTutorialVisitPending = True
            ProjectTutorial_Show True
            GoTo CleanExit
        End If
        If action = "IMPORT" Then
            ProjectWelcome_ImportPrevious
            If Migration_LastResult() <> "CANCELLED" Then Exit Do
            If Not ProjectWelcome_ShouldShow() Then Exit Do
            If Application.ActiveWorkbook Is Nothing Then Exit Do
            If Not (Application.ActiveWorkbook Is ThisWorkbook) Then Exit Do
        Else
            If action <> "NEW" Or Not ProjectWelcome_CanContinue() Then Exit Do
            If MsgBox(TextCatalog_Get("WELCOME.NEW.DISCARD.CONFIRM", language), _
                vbQuestion + vbYesNo + vbDefaultButton2, TextCatalog_Get("WELCOME.TITLE", language)) = vbYes Then Exit Do
        End If
    Loop
    If action = "NEW" Then
        ProjectWelcome_StartNew True
    ElseIf action = "CONTINUE" Then
        ProjectWelcome_ContinueProject
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
    ProjectWelcome_ResetTutorialVisit
End Sub

Public Sub ProjectWelcome_StartNew(Optional ByVal discardConfirmed As Boolean = False)
    If Not WorkbookSchema_UserActionAllowed() Then Exit Sub
    If ProjectWelcome_CanContinue() Then
        If Not discardConfirmed Then
            If MsgBox(TextCatalog_Get("WELCOME.NEW.DISCARD.CONFIRM", Settings_GetGlobalDisplayLanguage()), _
                vbQuestion + vbYesNo + vbDefaultButton2, TextCatalog_Get("WELCOME.TITLE", Settings_GetGlobalDisplayLanguage())) <> vbYes Then Exit Sub
        End If
        Armageddon True
    End If
    If Not Planning_IsNewOrFullyReset() Then Exit Sub
    ProjectWelcome_Acknowledge
    ProjectWelcome_ResetTutorialVisit
    ThisWorkbook.Worksheets("WBS").Activate
    If Len(ThisWorkbook.path) > 0 Then ThisWorkbook.Save
End Sub

Public Sub ProjectWelcome_ContinueProject()
    If Not WorkbookSchema_UserActionAllowed() Or Not ProjectWelcome_CanContinue() Then Exit Sub
    ProjectWelcome_Acknowledge
    ProjectWelcome_ResetTutorialVisit
    ThisWorkbook.Worksheets("WBS").Activate
End Sub

Public Sub ProjectWelcome_Acknowledge()
    Dim item As Object
    Set item = ProjectWelcome_Message()
    If PlanningMessage_IsAcknowledged(item, False) Then Exit Sub
    SetPlanningWarningAckState item, True, True
    If Not PlanningMessage_IsAcknowledged(item) Then Err.Raise 5, "ProjectWelcome", "ACK_NOT_PERSISTED"
End Sub




