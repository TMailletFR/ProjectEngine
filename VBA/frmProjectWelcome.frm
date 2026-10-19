VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmProjectWelcome 
   ClientHeight    =   7800
   ClientLeft      =   930
   ClientTop       =   3705
   ClientWidth     =   11400
   OleObjectBlob   =   "frmProjectWelcome.frx":0000
End
Attribute VB_Name = "frmProjectWelcome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private mPE As PEConsolePresentation

Private mSelectedAction As String

Public Property Get SelectedAction() As String
    SelectedAction = mSelectedAction
End Property

Public Sub LoadWelcome(ByVal title As String, ByVal message As String, _
                       ByVal importLabel As String, ByVal newLabel As String, ByVal closeLabel As String, _
                       Optional ByVal canContinue As Boolean = False)
    PEBegin
    mSelectedAction = ""
    Me.caption = title
    Me.width = 640
    Me.height = 324
    Me.BackColor = RGB(247, 248, 250)
    Me.StartUpPosition = 1
    lblTitle.caption = title
    lblTitle.left = 24: lblTitle.top = 18
    lblTitle.width = 580: lblTitle.height = 24
    lblTitle.font.name = "Segoe UI": lblTitle.font.size = 12: lblTitle.font.bold = True
    lblTitle.BackStyle = fmBackStyleTransparent
    lblCounter.Visible = False
    lblMessageType.Visible = False
    chkWarningAck.Visible = False
    cmdClearAck.Visible = False
    cmdClearHistory.Visible = False
    txtMessage.text = message
    txtMessage.left = 24: txtMessage.top = 60
    txtMessage.width = 580: txtMessage.height = 105
    txtMessage.Multiline = True: txtMessage.WordWrap = True: txtMessage.Locked = True
    txtMessage.font.name = "Segoe UI": txtMessage.font.size = 11
    txtMessage.BackColor = RGB(247, 248, 250)
    txtMessage.SpecialEffect = fmSpecialEffectFlat
    txtMessage.ScrollBars = fmScrollBarsNone
    ConfigureButton cmdPrevious, importLabel, 24, 210
    ConfigureButton cmdNext, newLabel, 246, 220
    ConfigureButton cmdClose, closeLabel, 478, 126
    ConfigureButton cmdClearAck, TextCatalog_Get("WELCOME.TUTORIAL", Settings_GetGlobalDisplayLanguage()), 24, 210
    cmdClearAck.top = 234
    cmdClearAck.Visible = True
    ConfigureButton cmdClearHistory, TextCatalog_Get("WELCOME.CONTINUE", Settings_GetGlobalDisplayLanguage()), 246, 220
    cmdClearHistory.top = 234
    cmdClearHistory.Visible = canContinue
    cmdNext.Default = Not canContinue
    cmdClearHistory.Default = canContinue
    cmdClose.Cancel = True
    PEWelcome_Compact Me
    PESync
End Sub

Private Sub ConfigureButton(ByVal Button As Object, ByVal caption As String, ByVal left As Single, ByVal width As Single)
    Button.caption = caption
    Button.left = left: Button.top = 192
    Button.width = width: Button.height = 32
    Button.font.name = "Segoe UI": Button.font.size = 10
    Button.Enabled = True
End Sub

Private Sub cmdPrevious_Click()
    mSelectedAction = "IMPORT"
    PEEnd
    Me.Hide
End Sub

Private Sub cmdNext_Click()
    mSelectedAction = "NEW"
    PEEnd
    Me.Hide
End Sub

Private Sub cmdClose_Click()
    PEEnd
    Me.Hide
End Sub

Private Sub cmdClearAck_Click()
    mSelectedAction = "TUTORIAL"
    PEEnd
    Me.Hide
End Sub

Private Sub cmdClearHistory_Click()
    mSelectedAction = "CONTINUE"
    PEEnd
    Me.Hide
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode <> vbFormControlMenu Then PEEnd
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        PEEnd
    Me.Hide
    End If
End Sub


















































































































































Private Sub PEBegin()
    If Not mPE Is Nothing Then mPE.BeginUpdate
End Sub
Private Sub PESync()
    If mPE Is Nothing Then
        Set mPE = New PEConsolePresentation
        mPE.Attach Me, Array("cmdPrevious", "cmdNext", "cmdClose", "cmdClearAck", "cmdClearHistory"), "cmdNext|cmdClearHistory"
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
        Case "cmdClearAck": cmdClearAck_Click
        Case "cmdClearHistory": cmdClearHistory_Click
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

Private Sub peFocus_cmdClearAck_Enter()
    If Not mPE Is Nothing Then mPE.FocusChanged "cmdClearAck", True
End Sub
Private Sub peFocus_cmdClearAck_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Not mPE Is Nothing Then mPE.FocusChanged "cmdClearAck", False
End Sub

Private Sub peFocus_cmdClearHistory_Enter()
    If Not mPE Is Nothing Then mPE.FocusChanged "cmdClearHistory", True
End Sub
Private Sub peFocus_cmdClearHistory_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Not mPE Is Nothing Then mPE.FocusChanged "cmdClearHistory", False
End Sub

Public Sub PERelease()
    PEEnd
End Sub
Public Sub PERefresh()
    PEBegin
    PESync
End Sub


