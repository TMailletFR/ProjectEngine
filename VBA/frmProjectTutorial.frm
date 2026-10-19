VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmProjectTutorial 
   ClientHeight    =   7800
   ClientLeft      =   930
   ClientTop       =   3705
   ClientWidth     =   11400
   OleObjectBlob   =   "frmProjectTutorial.frx":0000
End
Attribute VB_Name = "frmProjectTutorial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private mPE As PEConsolePresentation
Private mHomeSizing As Boolean
Private mStep As Long, mResume As Long
Private mChoosing As Boolean, mBusy As Boolean, mReady As Boolean, mLanguageUpdating As Boolean
Private mSizing As Boolean, mGripX As Single, mGripY As Single
Private WithEvents mLanguage As MSForms.ComboBox
Attribute mLanguage.VB_VarHelpID = -1
Private WithEvents mHome As MSForms.CommandButton
Attribute mHome.VB_VarHelpID = -1
Private WithEvents mDiscovery As MSForms.CommandButton
Attribute mDiscovery.VB_VarHelpID = -1
Private WithEvents mPractice As MSForms.CommandButton
Attribute mPractice.VB_VarHelpID = -1
Private WithEvents mGrip As MSForms.Label
Attribute mGrip.VB_VarHelpID = -1
Private mModeLabel As MSForms.Label, mStatus As MSForms.Label
Private mContent As MSForms.frame
Private mBody As String

Public Function IsReadyToShow() As Boolean
    IsReadyToShow = mReady
End Function

Public Sub Prepare()
    Dim i As Long, control As Object
    Me.StartUpPosition = 0
    Me.width = 530: Me.height = 440
    Me.BackColor = RGB(247, 248, 250)
    Me.left = Application.left + Application.width - Me.width - 20
    Me.top = Application.top + 110
    If Me.left < Application.left Then Me.left = Application.left
    If Application.height > 460 Then
        If Me.height > Application.height - 70 Then Me.height = Application.height - 70
    End If
    lblCounter.Visible = False: lblMessageType.Visible = False
    chkWarningAck.Visible = False: cmdClearAck.Visible = False: cmdClearHistory.Visible = False
    txtMessage.Visible = False
    Set mHome = Me.Controls.Add("Forms.CommandButton.1", "tutorialHome", True)
    Set mLanguage = Me.Controls.Add("Forms.ComboBox.1", "tutorialLanguage", True)
    mLanguage.style = fmStyleDropDownList
    mLanguage.AddItem "Fran" & ChrW(231) & "ais": mLanguage.AddItem "English"
    Set mModeLabel = Me.Controls.Add("Forms.Label.1", "tutorialMode", True)
    Set mStatus = Me.Controls.Add("Forms.Label.1", "tutorialStatus", True)
    Set mDiscovery = Me.Controls.Add("Forms.CommandButton.1", "tutorialDiscovery", True)
    Set mPractice = Me.Controls.Add("Forms.CommandButton.1", "tutorialPractice", True)
    Set mContent = Me.Controls.Add("Forms.Frame.1", "tutorialContent", True)
    mContent.caption = "": mContent.BorderStyle = fmBorderStyleNone
    mContent.BackColor = Me.BackColor: mContent.ScrollBars = fmScrollBarsVertical
    For i = 0 To 3
        Set control = mContent.Controls.Add("Forms.Label.1", "heading" & CStr(i), True)
        control.font.name = "Segoe UI": control.font.size = 10: control.font.bold = True
        control.ForeColor = RGB(25, 85, 105): control.BackStyle = fmBackStyleTransparent
        Set control = mContent.Controls.Add("Forms.Label.1", "body" & CStr(i), True)
        control.font.name = "Segoe UI": control.font.size = 11
        control.BackStyle = fmBackStyleTransparent: control.WordWrap = True
    Next i
    Set mGrip = Me.Controls.Add("Forms.Label.1", "tutorialGrip", True)
    mGrip.caption = "/": mGrip.ForeColor = RGB(85, 85, 85)
    mGrip.ControlTipText = ProjectTutorial_Text("RESIZE")
    cmdClose.Cancel = True
    mReady = True
    ShowHome
End Sub

Private Sub Localize()
    PEBegin
    Me.caption = ProjectTutorial_Text("TITLE")
    mHome.caption = ProjectTutorial_Text("HOME")
    mDiscovery.caption = ProjectTutorial_Text("DISCOVERY.CHOICE")
    mPractice.caption = ProjectTutorial_Text("PRACTICE.CHOICE")
    cmdClose.caption = ProjectTutorial_Text("CLOSE")
    mLanguageUpdating = True
    mLanguage.ListIndex = IIf(Settings_GetGlobalDisplayLanguage() = "FR", 0, 1)
    mLanguage.Enabled = Not ThisWorkbook.ReadOnly
    mLanguageUpdating = False
    mModeLabel.caption = ProjectTutorial_Text(IIf(Len(ProjectTutorial_Mode()) = 0, "HOME", ProjectTutorial_Mode() & ".LABEL"))
    mModeLabel.Visible = False
End Sub

Private Sub ShowHome()
    PEBegin
    ProjectTutorial_Home
    mStep = 0: mChoosing = False
    Localize
    lblTitle.caption = ProjectTutorial_Text("HOME.TITLE")
    mBody = ProjectTutorial_Text("HOME.BODY")
    mContent.Visible = False
    mDiscovery.Visible = True: mPractice.Visible = True
    cmdPrevious.Visible = False: cmdNext.Visible = False
    mStatus.caption = ProjectTutorial_Text("HOME.STATUS")
    If ThisWorkbook.ReadOnly Then mStatus.caption = ProjectTutorial_Text("PROGRESS.READONLY")
    LayoutConsole
End Sub

Private Sub ChooseMode(ByVal mode As String)
    If mBusy Or Not ProjectTutorial_IsOwnerWindowActive() Then Exit Sub
    PEBegin
    ProjectTutorial_SelectMode mode
    mDiscovery.Visible = False: mPractice.Visible = False
    mContent.Visible = True
    mResume = ProjectTutorial_ResumeStep()
    mChoosing = (mResume > 0)
    If mChoosing Then
        Localize
        lblTitle.caption = ProjectTutorial_Text("RESUME")
        mBody = ProjectTutorial_Text("RESUME.MESSAGE")
        cmdPrevious.caption = ProjectTutorial_Text("RESTART"): cmdNext.caption = ProjectTutorial_Text("RESUME")
        cmdPrevious.Visible = True: cmdNext.Visible = True: cmdPrevious.Enabled = True
        mStatus.caption = ProjectTutorial_Text("RESUME.POLICY")
        LayoutConsole
    Else
        mStep = 1
        Me.height = 440
        RenderPage True
    End If
End Sub

Private Sub mDiscovery_Click()
    ChooseMode "DISCOVERY"
End Sub

Private Sub mPractice_Click()
    ChooseMode "PRACTICE"
End Sub

Private Sub mHome_Click()
    If Not mBusy Then ShowHome
End Sub

Private Sub RenderPage(ByVal navigate As Boolean)
    Dim key As String, live As Boolean
    If mBusy Then Exit Sub
    On Error GoTo Failed
    PEBegin
    mBusy = True
    Localize
    key = ProjectTutorial_StepKey(mStep)
    lblTitle.caption = CStr(mStep) & "/" & CStr(ProjectTutorial_StepCount()) & " - " & ProjectTutorial_Text(key & ".TITLE")
    mBody = ProjectTutorial_Text(key & ".BODY")
    live = ProjectTutorial_ShowPage(mStep, navigate, navigate)
    If ProjectTutorial_Mode() = "PRACTICE" Then
        mStatus.caption = ProjectTutorial_Text(IIf(ProjectTutorial_IsEditableStep(mStep), "PRACTICE.EDIT", "PRACTICE.OBSERVE"))
    Else
        mStatus.caption = ProjectTutorial_Text("DISCOVERY.STATUS")
    End If
    If Not live And Len(ProjectTutorial_Operation(mStep)) = 0 Then mStatus.caption = mStatus.caption & " " & ProjectTutorial_Text("EMPTY.STATUS")
    cmdPrevious.Visible = True: cmdPrevious.Enabled = (mStep > 1)
    cmdNext.Visible = True: cmdNext.Enabled = True
    cmdPrevious.caption = ProjectTutorial_Text("PREVIOUS")
    cmdNext.caption = ProjectTutorial_Text(IIf(mStep = ProjectTutorial_StepCount(), "FINISH", "NEXT"))
    LayoutConsole
    mBusy = False
    Exit Sub
Failed:
    Debug.Print "Tutorial form: " & CStr(Err.number) & ": " & Err.description
    mBusy = False
    CalcBridge_ShowSingleConsoleMessage "STOP", "TUTORIAL.ERROR", TextCatalog_Arguments("Details", Err.description)
    CloseConsole
End Sub

Public Sub RefreshCurrentPage()
    If mStep > 0 And Not mChoosing Then RenderPage False
End Sub

Private Sub LayoutConsole()
    Dim sections As Variant, separator As Long, i As Long
    Dim text As String, Y As Double, height As Double, heading As Object, body As Object
    If Not mReady Then Exit Sub
    If mStep = 0 And Not mChoosing Then
        LayoutHome
        Exit Sub
    End If
    Me.ScrollBars = fmScrollBarsNone: Me.ScrollTop = 0
    mStatus.Visible = True
    mHome.left = 16: mHome.top = 12: mHome.width = 76: mHome.height = 24
    mLanguage.left = 104: mLanguage.top = 12: mLanguage.width = 112: mLanguage.height = 24
    mModeLabel.left = 230: mModeLabel.top = 14: mModeLabel.width = Me.InsideWidth - 248: mModeLabel.height = 30
    mModeLabel.WordWrap = True
    lblTitle.left = 16: lblTitle.top = 50: lblTitle.width = Me.InsideWidth - 32: lblTitle.height = 42
    lblTitle.font.name = "Segoe UI": lblTitle.font.size = 12: lblTitle.font.bold = True: lblTitle.WordWrap = True
    height = (Me.InsideHeight - 182) / 2
    If height < 82 Then height = 82
    mDiscovery.left = 16: mDiscovery.top = 104: mDiscovery.width = Me.InsideWidth - 32: mDiscovery.height = height
    mPractice.left = 16: mPractice.top = mDiscovery.top + height + 12: mPractice.width = Me.InsideWidth - 32: mPractice.height = height
    mDiscovery.font.name = "Segoe UI": mDiscovery.font.size = 10
    mPractice.font.name = "Segoe UI": mPractice.font.size = 10
    mDiscovery.WordWrap = True: mPractice.WordWrap = True
    mContent.left = 16: mContent.top = IIf(mStep = 0 And Not mChoosing, 258, 102)
    mContent.width = Me.InsideWidth - 32
    height = Me.InsideHeight - mContent.top - 104
    If height < 18 Then height = 18
    mContent.height = height
    sections = Split(mBody, vbCrLf & vbCrLf)
    Y = 0
    For i = 0 To 3
        Set heading = mContent.Controls("heading" & CStr(i)): Set body = mContent.Controls("body" & CStr(i))
        heading.Visible = False: body.Visible = False
        If i <= UBound(sections) Then
            text = CStr(sections(i)): separator = InStr(text, vbCrLf)
            If separator > 0 Then
                heading.left = 4: heading.top = Y: heading.width = mContent.width - 28
                heading.caption = left$(text, separator - 1): heading.WordWrap = True
                heading.AutoSize = True: heading.Visible = True
                Y = heading.top + heading.height + 8: text = Mid$(text, separator + 2)
            End If
            body.AutoSize = False: body.left = 4: body.top = Y: body.width = mContent.width - 28
            body.caption = text: body.AutoSize = True: body.Visible = True
            Y = body.top + body.height + 20
        End If
    Next i
    mContent.ScrollHeight = Y
    mStatus.left = 18: mStatus.top = Me.InsideHeight - 94
    mStatus.width = Me.InsideWidth - 36: mStatus.height = 40: mStatus.WordWrap = True
    mStatus.font.size = 9: mStatus.ForeColor = RGB(70, 85, 95)
    cmdPrevious.width = (Me.InsideWidth - 56) / 3
    cmdNext.width = cmdPrevious.width: cmdClose.width = cmdPrevious.width
    cmdPrevious.left = 16: cmdNext.left = cmdPrevious.left + cmdPrevious.width + 12
    cmdClose.left = cmdNext.left + cmdNext.width + 12
    cmdPrevious.top = Me.InsideHeight - 46: cmdNext.top = cmdPrevious.top: cmdClose.top = cmdPrevious.top
    mGrip.left = Me.InsideWidth - 16: mGrip.top = Me.InsideHeight - 16: mGrip.width = 14: mGrip.height = 14
    PESync
    Me.Repaint
End Sub

Private Sub UserForm_Resize()
    PEBegin
    LayoutConsole
End Sub

Private Sub mGrip_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button <> 1 Then Exit Sub
    mSizing = True: mGripX = X: mGripY = Y
End Sub

Private Sub mGrip_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim width As Double, height As Double
    If Not mSizing Or Button <> 1 Then Exit Sub
    width = Me.width + X - mGripX: height = Me.height + Y - mGripY
    If width < 390 Then width = 390
    If height < 370 Then height = 370
    If width > 1000 Then width = 1000
    If height > 900 Then height = 900
    Me.width = width: Me.height = height
End Sub

Private Sub mGrip_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    mSizing = False
End Sub

Private Sub mLanguage_Change()
    If mBusy Or mLanguageUpdating Then Exit Sub
    If Not ProjectTutorial_IsOwnerWindowActive() Then Exit Sub
    If Not Settings_SelectGlobalLanguage(IIf(mLanguage.ListIndex = 0, "FR", "EN")) Then
        Localize
        mStatus.caption = ProjectTutorial_Text("LANGUAGE.FAILED")
        CalcBridge_ShowSingleConsoleMessage "STOP", "TUTORIAL.LANGUAGE.FAILED"
        Exit Sub
    End If
    If mStep = 0 Then
        If mChoosing Then
            Localize
            mBody = ProjectTutorial_Text("RESUME.MESSAGE")
            cmdPrevious.caption = ProjectTutorial_Text("RESTART"): cmdNext.caption = ProjectTutorial_Text("RESUME")
            LayoutConsole
        Else
            ShowHome
        End If
    Else
        RenderPage False
    End If
End Sub

Private Sub cmdPrevious_Click()
    If mBusy Then Exit Sub
    If mChoosing Then
        mChoosing = False: mStep = 1
    ElseIf mStep > 1 Then
        mStep = mStep - 1
    End If
    RenderPage True
End Sub

Private Sub cmdNext_Click()
    If mBusy Or Not ProjectTutorial_IsOwnerWindowActive() Then Exit Sub
    If mChoosing Then
        mChoosing = False: mStep = mResume
        RenderPage True
    ElseIf mStep = ProjectTutorial_StepCount() Then
        ProjectTutorial_Complete
        CloseConsole
    ElseIf mStep > 0 Then
        If Not ProjectTutorial_TryNext(mStep) Then
            mStatus.caption = ProjectTutorial_Text("ACTION.RETRY")
            Exit Sub
        End If
        mStep = mStep + 1: RenderPage True
    End If
End Sub

Private Sub CloseConsole()
    If Not ProjectTutorial_Cleanup() Then
        CalcBridge_ShowSingleConsoleMessage "STOP", "TUTORIAL.SAVE.BLOCKED"
        Exit Sub
    End If
    PEEnd
    Me.Hide
    ProjectTutorial_RequestClose
End Sub

Private Sub cmdClose_Click()
    If Not mBusy Then CloseConsole
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode <> vbFormControlMenu Then PEEnd
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        If Not mBusy Then CloseConsole
    End If
End Sub


Private Sub PEBegin()
    If Not mPE Is Nothing Then mPE.BeginUpdate
End Sub
Private Sub PESync()
    If mPE Is Nothing Then
        Set mPE = New PEConsolePresentation
        mPE.Attach Me, Array("cmdPrevious", "cmdNext", "cmdClose", "tutorialHome", "tutorialDiscovery", "tutorialPractice"), "cmdNext|tutorialDiscovery|tutorialPractice"
    End If
    PEConsole_Theme Me
    mPE.Sync
End Sub
Private Sub PEEnd()
    If Not mPE Is Nothing Then mPE.Detach
    Set mPE = Nothing
End Sub
Public Sub PEHover(ByVal name As String)
    If Not mPE Is Nothing Then mPE.Hover name
End Sub
Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    PEHover ""
End Sub
Private Sub peDefault_Click()
    If Not mPE Is Nothing Then mPE.InvokeDefault
End Sub
Private Sub peCancel_Click()
    If Not mPE Is Nothing Then mPE.InvokeCancel
End Sub
Public Sub PEInvoke(ByVal name As String)
    Select Case name
        Case "cmdPrevious": cmdPrevious_Click
        Case "cmdNext": cmdNext_Click
        Case "cmdClose": cmdClose_Click
        Case "tutorialHome": mHome_Click
        Case "tutorialDiscovery": mDiscovery_Click
        Case "tutorialPractice": mPractice_Click
    End Select
End Sub

Private Sub peFocus_cmdPrevious_Enter()
    If Not mPE Is Nothing Then mPE.FocusChanged "cmdPrevious", True
End Sub
Private Sub peFocus_cmdPrevious_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Not mPE Is Nothing Then mPE.FocusChanged "cmdPrevious", False
End Sub

Private Sub peFocus_cmdNext_Enter()
    If Not mPE Is Nothing Then mPE.FocusChanged "cmdNext", True
End Sub
Private Sub peFocus_cmdNext_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Not mPE Is Nothing Then mPE.FocusChanged "cmdNext", False
End Sub

Private Sub peFocus_cmdClose_Enter()
    If Not mPE Is Nothing Then mPE.FocusChanged "cmdClose", True
End Sub
Private Sub peFocus_cmdClose_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Not mPE Is Nothing Then mPE.FocusChanged "cmdClose", False
End Sub

Private Sub peFocus_tutorialHome_Enter()
    If Not mPE Is Nothing Then mPE.FocusChanged "tutorialHome", True
End Sub
Private Sub peFocus_tutorialHome_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Not mPE Is Nothing Then mPE.FocusChanged "tutorialHome", False
End Sub

Private Sub peFocus_tutorialDiscovery_Enter()
    If Not mPE Is Nothing Then mPE.FocusChanged "tutorialDiscovery", True
End Sub
Private Sub peFocus_tutorialDiscovery_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Not mPE Is Nothing Then mPE.FocusChanged "tutorialDiscovery", False
End Sub

Private Sub peFocus_tutorialPractice_Enter()
    If Not mPE Is Nothing Then mPE.FocusChanged "tutorialPractice", True
End Sub
Private Sub peFocus_tutorialPractice_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Not mPE Is Nothing Then mPE.FocusChanged "tutorialPractice", False
End Sub

Public Sub PERelease()
    PEEnd
End Sub
Public Sub PERefresh()
    PEBegin
    PESync
End Sub

Private Sub LayoutHome()
    Dim width As Single, firstHeight As Single, secondHeight As Single, clientHeight As Single
    Dim frameHeight As Single, maximum As Single, failure As Long, description As String
    If mHomeSizing Then Exit Sub
    On Error GoTo Failed
    mHomeSizing = True
    PEBegin
    PEConsole_Theme Me
    Me.ScrollTop = 0: Me.ScrollBars = fmScrollBarsNone
    width = Me.InsideWidth - 32
    mHome.Move 16, 12, 76, 24
    mLanguage.Move 104, 12, 112, 24
    mModeLabel.Visible = False: mStatus.Visible = False: mContent.Visible = False
    lblTitle.Move 16, 48, width, 24
    lblTitle.font.name = "Segoe UI": lblTitle.font.size = 12: lblTitle.font.bold = True
    mDiscovery.font.name = "Segoe UI": mDiscovery.font.size = 11
    mDiscovery.font.bold = False: mDiscovery.font.italic = False
    mPractice.font.name = "Segoe UI": mPractice.font.size = 11
    mPractice.font.bold = False: mPractice.font.italic = False
    firstHeight = PEChoiceHeight(mDiscovery.caption, width, "Segoe UI", 11)
    secondHeight = PEChoiceHeight(mPractice.caption, width, "Segoe UI", 11)
    mDiscovery.Move 16, 80, width, firstHeight
    mPractice.Move 16, 80 + firstHeight + 8, width, secondHeight
    mDiscovery.Visible = True: mPractice.Visible = True
    cmdPrevious.Visible = False: cmdNext.Visible = False
    cmdClose.Move Me.InsideWidth - 146, mPractice.top + secondHeight + 12, 130, 30
    clientHeight = cmdClose.top + cmdClose.height + 12
    frameHeight = Me.height - Me.InsideHeight
    maximum = Application.height - 70
    If maximum < 300 Then maximum = 300
    If clientHeight + frameHeight > maximum Then
        Me.height = maximum
        Me.ScrollBars = fmScrollBarsVertical
        Me.ScrollHeight = clientHeight
    Else
        Me.height = clientHeight + frameHeight
    End If
    mGrip.Move Me.InsideWidth - 16, Me.InsideHeight - 16, 14, 14
    PESync
    mHomeSizing = False
    Me.Repaint
    Exit Sub
Failed:
    failure = Err.number: description = Err.description
    mHomeSizing = False
    Err.Raise failure, , description
End Sub
