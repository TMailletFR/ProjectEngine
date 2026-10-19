VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmPlanningMessages 
   ClientHeight    =   7800
   ClientLeft      =   930
   ClientTop       =   3705
   ClientWidth     =   11400
   OleObjectBlob   =   "frmPlanningMessages.frx":0000
End
Attribute VB_Name = "frmPlanningMessages"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private mPE As PEConsolePresentation

'===============================================================================
' MODULE : frmPlanningMessages
' DOMAINE / DOMAIN : Planning Console UI
'
' FR
' Affiche les messages planning, la navigation et les actions ACK dans la console modale.
' Ne prepare pas, ne filtre pas et ne journalise pas les diagnostics.
'
' EN
' Displays planning messages, navigation and ACK actions in the modal console.
' Does not prepare, filter or log diagnostics.
'
' CONTRATS / CONTRACTS : LoadMessages
' CALLBACKS EXTERNES / EXTERNAL CALLBACKS : UserForm_Initialize
'===============================================================================

Private mItems As Collection
Private mTitle As String
Private mIndex As Long
Private mLoadingMessage As Boolean

'------------------------------------------------------------------------------
' FR: Charge la collection Messages depuis sa source proprietaire sans appliquer de politique aval.
' EN: Loads the Messages collection from its owning source without applying downstream policy.
'------------------------------------------------------------------------------

Public Sub LoadMessages( _
    ByVal messages As Collection, _
    Optional ByVal windowTitle As String = vbNullString)

    Set mItems = MessageEngine_PrepareDisplayMessages(messages)
    If Len(windowTitle) = 0 Then
        mTitle = TextCatalog_Get("CONSOLE.WINDOW.TITLE", EventHistory_CurrentLanguage())
    Else
        mTitle = windowTitle
    End If
    mIndex = 1

    PEBegin
    ApplyFormLayout
    CenterFormOnExcel
    RenderCurrentMessage

End Sub

'------------------------------------------------------------------------------
' FR: Force la creation/layout des controles avant que la fenetre devienne visible.
' EN: Forces control creation/layout before the window becomes visible.
'------------------------------------------------------------------------------

Public Sub PrepareForImmediateShow()

    Me.Repaint

End Sub

'------------------------------------------------------------------------------
' FR: Traite le clic utilisateur cmd Previous Click dans planning message form.
' EN: Handles the cmd Previous Click user click in planning message form.
'------------------------------------------------------------------------------
Private Sub cmdPrevious_Click()

    If mItems Is Nothing Then Exit Sub
    If mItems.count = 0 Then Exit Sub

    If mIndex > 1 Then
        mIndex = mIndex - 1
        RenderCurrentMessage
    End If

End Sub

'------------------------------------------------------------------------------
' FR: Traite le clic utilisateur cmd Next Click dans planning message form.
' EN: Handles the cmd Next Click user click in planning message form.
'------------------------------------------------------------------------------
Private Sub cmdNext_Click()

    If mItems Is Nothing Then Exit Sub
    If mItems.count = 0 Then Exit Sub

    If mIndex < mItems.count Then
        mIndex = mIndex + 1
        RenderCurrentMessage
    End If

End Sub

'------------------------------------------------------------------------------
' FR: Traite le clic utilisateur cmd Close Click dans planning message form.
' EN: Handles the cmd Close Click user click in planning message form.
'------------------------------------------------------------------------------
Private Sub cmdClose_Click()

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

'------------------------------------------------------------------------------
' FR: Actualise Render Current Message sans modifier les regles metier qui produisent les donnees.
' EN: Refreshes Render Current Message without changing the business rules that produce the data.
'------------------------------------------------------------------------------

Private Sub RenderCurrentMessage()

    Dim item As Object
    Dim msgType As String
    PEBegin

    If mItems Is Nothing Then Exit Sub

    If mItems.count = 0 Then
        lblTitle.caption = TextCatalog_Get("CONSOLE.SECTION.MESSAGES", EventHistory_CurrentLanguage())
        lblCounter.caption = TextCatalog_Get("CONSOLE.EMPTY.COUNTER", EventHistory_CurrentLanguage())
        lblMessageType.caption = ""
        txtMessage.text = ""
        cmdPrevious.Enabled = False
        cmdNext.Enabled = False
        chkWarningAck.Visible = False
        PESync
        Exit Sub
    End If

    If mIndex < 1 Then mIndex = 1
    If mIndex > mItems.count Then mIndex = mItems.count

    Set item = mItems(mIndex)
    msgType = UCase$(Trim$(CStr(item("Type"))))
    If MessageIsWarning(msgType) Then item("Acknowledged") = PlanningMessage_IsAcknowledged(item)

    lblTitle.caption = TextCatalog_Get("CONSOLE.SECTION.MESSAGES", EventHistory_CurrentLanguage())
    lblCounter.caption = BuildCategoryProgressCaption(msgType)

    mLoadingMessage = True
    ApplyMessageTypeVisual msgType
    ApplyAckControlVisual item, msgType
    mLoadingMessage = False

    txtMessage.text = FormatPlanningConsoleMessageForCurrentLanguage(CStr(item("Message")))

    cmdPrevious.Enabled = (mIndex > 1)
    cmdNext.Enabled = (mIndex < mItems.count)
    PESync

End Sub

'------------------------------------------------------------------------------
' FR: Actualise Apply Message Type Visual sans modifier les regles metier qui produisent les donnees.
' EN: Refreshes Apply Message Type Visual without changing the business rules that produce the data.
'------------------------------------------------------------------------------

Private Sub ApplyMessageTypeVisual(ByVal msgType As String)

    Select Case UCase$(Trim$(msgType))

        Case "STOP", "ERROR"
            lblMessageType.caption = TextCatalog_Get("CONSOLE.SEVERITY.STOP", EventHistory_CurrentLanguage())
            lblMessageType.ForeColor = RGB(156, 0, 6)
            lblMessageType.BackColor = RGB(255, 235, 238)

        Case "WARNING"
            If CurrentMessageAcknowledged() Then
                lblMessageType.caption = TextCatalog_Get("CONSOLE.SEVERITY.WARNING", EventHistory_CurrentLanguage())
                lblMessageType.ForeColor = RGB(92, 74, 0)
                lblMessageType.BackColor = RGB(236, 232, 214)
            Else
                lblMessageType.caption = TextCatalog_Get("CONSOLE.SEVERITY.WARNING", EventHistory_CurrentLanguage())
                lblMessageType.ForeColor = RGB(156, 101, 0)
                lblMessageType.BackColor = RGB(255, 248, 225)
            End If

        Case "INFO"
            lblMessageType.caption = TextCatalog_Get("CONSOLE.SEVERITY.INFO", EventHistory_CurrentLanguage())
            lblMessageType.ForeColor = RGB(0, 97, 0)
            lblMessageType.BackColor = RGB(232, 245, 233)

        Case Else
            lblMessageType.caption = msgType
            ApplyMessageTypeBadgeLayout
            lblMessageType.ForeColor = RGB(60, 60, 60)
            lblMessageType.BackColor = RGB(240, 240, 240)

    End Select

End Sub

'------------------------------------------------------------------------------
' FR: Retourne la map Current Message Acknowledged sans exposer de mutateur sur l'etat source.
' EN: Returns the Current Message Acknowledged map without exposing a mutator for source state.
'------------------------------------------------------------------------------

Private Function CurrentMessageAcknowledged() As Boolean

    Dim item As Object

    On Error GoTo SafeExit

    If mItems Is Nothing Then Exit Function
    If mIndex < 1 Or mIndex > mItems.count Then Exit Function
    If TypeName(mItems(mIndex)) <> "Dictionary" Then Exit Function

    Set item = mItems(mIndex)
    CurrentMessageAcknowledged = CBool(item("Acknowledged"))

SafeExit:
End Function

'------------------------------------------------------------------------------
' FR: Retourne la valeur Message Is Warning sans modifier les donnees d'entree.
' EN: Returns the Message Is Warning value without mutating input data.
'------------------------------------------------------------------------------

Private Function MessageIsWarning(ByVal msgType As String) As Boolean

    MessageIsWarning = (UCase$(Trim$(msgType)) = "WARNING")

End Function

'------------------------------------------------------------------------------
' FR: Actualise Apply Ack Control Visual sans modifier les regles metier qui produisent les donnees.
' EN: Refreshes Apply Ack Control Visual without changing the business rules that produce the data.
'------------------------------------------------------------------------------

Private Sub ApplyAckControlVisual( _
    ByVal item As Object, _
    ByVal msgType As String)

    If Not MessageIsWarning(msgType) Then
        chkWarningAck.Visible = False
        Exit Sub
    End If

    chkWarningAck.Visible = PlanningMessage_CanAcknowledge(item)
    chkWarningAck.value = PlanningMessage_IsAcknowledged(item)
    chkWarningAck.caption = TextCatalog_Get("CONSOLE.COMMAND.HIDE_WARNING", EventHistory_CurrentLanguage())

End Sub

'------------------------------------------------------------------------------
' FR: Traite le clic utilisateur chk Warning Ack Click dans planning message form.
' EN: Handles the chk Warning Ack Click user click in planning message form.
'------------------------------------------------------------------------------
Private Sub chkWarningAck_Click()

    Dim item As Object

    If mLoadingMessage Then Exit Sub
    If mItems Is Nothing Then Exit Sub
    If mIndex < 1 Or mIndex > mItems.count Then Exit Sub
    If TypeName(mItems(mIndex)) <> "Dictionary" Then Exit Sub

    Set item = mItems(mIndex)
    If Not PlanningMessage_CanAcknowledge(item) Then Exit Sub

    SetPlanningWarningAckState item, CBool(chkWarningAck.value)
    item("Acknowledged") = PlanningMessage_IsAcknowledged(item)

    mLoadingMessage = True
    chkWarningAck.value = CBool(item("Acknowledged"))
    mLoadingMessage = False

    ApplyMessageTypeVisual CStr(item("Type"))
    PEConsole_Badge Me

End Sub

'------------------------------------------------------------------------------
' FR: Traite le clic utilisateur cmd Clear Ack Click dans planning message form.
' EN: Handles the cmd Clear Ack Click user click in planning message form.
'------------------------------------------------------------------------------
Private Sub cmdClearAck_Click()

    ClearPlanningWarningAcknowledgements
    RenderCurrentMessage

End Sub

'------------------------------------------------------------------------------
' FR: Traite le clic utilisateur cmd Clear History Click dans planning message form.
' EN: Handles the cmd Clear History Click user click in planning message form.
'------------------------------------------------------------------------------
Private Sub cmdClearHistory_Click()

    ClearPlanningEventHistory
    RenderCurrentMessage

End Sub
'------------------------------------------------------------------------------
' FR: Construit la valeur Category Progress Caption a partir des donnees fournies par l'appelant.
' EN: Builds the Category Progress Caption value from data supplied by the caller.
'------------------------------------------------------------------------------

Private Function BuildCategoryProgressCaption(ByVal currentType As String) As String

    BuildCategoryProgressCaption = MessageEngine_BuildCategoryProgressCaption( _
        mItems, _
        mIndex, _
        currentType, _
        EventHistory_CurrentLanguage())

End Function
'------------------------------------------------------------------------------
' FR: Actualise Apply Form Layout sans modifier les regles metier qui produisent les donnees.
' EN: Refreshes Apply Form Layout without changing the business rules that produce the data.
'------------------------------------------------------------------------------

Private Sub ApplyFormLayout()

    Dim formW As Single
    Dim formH As Single
    Dim margin As Single
    Dim headerTop As Single
    Dim headerH As Single
    Dim textTop As Single
    Dim textH As Single
    Dim footerH As Single
    Dim buttonTop As Single
    Dim buttonW As Single
    Dim buttonH As Single
    Dim gap As Single

    margin = 24
    headerTop = 18
    headerH = 42
    footerH = 50
    buttonW = 138
    buttonH = 30
    gap = 14

    PEMessageLayout_Calculate Me, mItems, formW, formH


    Me.caption = mTitle
    Me.StartUpPosition = 0
    Me.width = formW
    Me.height = formH
    Me.BackColor = RGB(247, 248, 250)

    lblTitle.left = margin
    lblTitle.top = headerTop
    lblTitle.width = 130
    lblTitle.height = 18
    lblTitle.caption = TextCatalog_Get("CONSOLE.SECTION.MESSAGES", EventHistory_CurrentLanguage())
    lblTitle.font.name = "Segoe UI"
    lblTitle.font.size = 10
    lblTitle.font.bold = True
    lblTitle.ForeColor = RGB(35, 35, 35)
    lblTitle.BackStyle = fmBackStyleTransparent

    lblMessageType.left = margin + 8
    lblMessageType.top = headerTop + 22 + 4
    lblMessageType.width = 110
    lblMessageType.height = 28
    lblMessageType.TextAlign = fmTextAlignCenter
    lblMessageType.font.name = "Segoe UI"
    lblMessageType.font.size = 10.66667
    lblMessageType.font.bold = True
    lblMessageType.BorderStyle = fmBorderStyleSingle
    lblMessageType.SpecialEffect = fmSpecialEffectFlat
    lblMessageType.WordWrap = False
    lblMessageType.AutoSize = False

    lblCounter.left = margin + 150
    lblCounter.top = headerTop
    lblCounter.width = Me.InsideWidth - (2 * margin) - 150
    lblCounter.height = 18
    lblCounter.TextAlign = fmTextAlignRight
    lblCounter.font.name = "Segoe UI"
    lblCounter.font.size = 10
    lblCounter.font.bold = True
    lblCounter.ForeColor = RGB(35, 35, 35)
    lblCounter.BackStyle = fmBackStyleTransparent

    textTop = headerTop + 22 + 28 + 8
    textH = Me.InsideHeight - textTop - footerH

    If textH < 24 Then textH = 24

    txtMessage.left = margin
    txtMessage.top = textTop
    txtMessage.width = Me.InsideWidth - (2 * margin)
    txtMessage.height = textH
    txtMessage.Multiline = True
    txtMessage.ScrollBars = fmScrollBarsVertical
    txtMessage.Locked = True
    txtMessage.WordWrap = True
    txtMessage.BorderStyle = fmBorderStyleSingle
    txtMessage.SpecialEffect = fmSpecialEffectFlat
    txtMessage.BackColor = RGB(247, 248, 250)
    txtMessage.ForeColor = RGB(25, 25, 25)
    txtMessage.font.name = "Segoe UI"
    txtMessage.font.size = 9

    buttonTop = txtMessage.top + txtMessage.height + 10

    cmdPrevious.left = margin
    cmdPrevious.top = buttonTop
    cmdPrevious.width = buttonW
    cmdPrevious.height = buttonH
    cmdPrevious.caption = TextCatalog_Get("CONSOLE.COMMAND.PREVIOUS", EventHistory_CurrentLanguage())
    cmdPrevious.font.name = "Segoe UI"
    cmdPrevious.font.size = 9

    cmdNext.left = cmdPrevious.left + buttonW + gap
    cmdNext.top = buttonTop
    cmdNext.width = buttonW
    cmdNext.height = buttonH
    cmdNext.caption = TextCatalog_Get("CONSOLE.COMMAND.NEXT", EventHistory_CurrentLanguage())
    cmdNext.font.name = "Segoe UI"
    cmdNext.font.size = 9

    cmdClose.left = Me.InsideWidth - margin - buttonW
    cmdClose.top = buttonTop
    cmdClose.width = buttonW
    cmdClose.height = buttonH
    cmdClose.caption = TextCatalog_Get("CONSOLE.COMMAND.CLOSE", EventHistory_CurrentLanguage())
    cmdClose.font.name = "Segoe UI"
    cmdClose.font.size = 9

    chkWarningAck.left = lblMessageType.left + lblMessageType.width + 10
    chkWarningAck.top = lblMessageType.top + 5
    chkWarningAck.width = 92
    chkWarningAck.height = 18
    chkWarningAck.caption = TextCatalog_Get("CONSOLE.COMMAND.HIDE_WARNING", EventHistory_CurrentLanguage())
    chkWarningAck.font.name = "Segoe UI"
    chkWarningAck.font.size = 9
    chkWarningAck.BackStyle = fmBackStyleTransparent
    chkWarningAck.Visible = False

    cmdClearAck.Visible = False
    cmdClearHistory.Visible = False
    cmdClose.Cancel = True

End Sub

'------------------------------------------------------------------------------
' FR: Retourne la map Calc Dynamic Form Width sans modifier les donnees d'entree.
' EN: Returns the Calc Dynamic Form Width map without mutating input data.
'------------------------------------------------------------------------------



'------------------------------------------------------------------------------
' FR: Retourne la map Calc Dynamic Form Height sans modifier les donnees d'entree.
' EN: Returns the Calc Dynamic Form Height map without mutating input data.
'------------------------------------------------------------------------------



'------------------------------------------------------------------------------
' FR: Positionne la valeur Form On Excel relativement a la fenetre Excel active.
' EN: Positions the Form On Excel value relative to the active Excel window.
'------------------------------------------------------------------------------

Private Sub CenterFormOnExcel()

    On Error GoTo SafeExit

    Me.StartUpPosition = 0

    Me.left = Application.left + ((Application.width - Me.width) / 2)
    Me.top = Application.top + ((Application.height - Me.height) / 2)

    If Me.left < 0 Then Me.left = 20
    If Me.top < 0 Then Me.top = 20

SafeExit:
End Sub

'------------------------------------------------------------------------------
' FR: Initialise l'etat visuel du formulaire de messages planning.
' EN: Initializes the planning message form visual state.
'------------------------------------------------------------------------------
Private Sub UserForm_Initialize()

    ApplyMessageTypeBadgeLayout

End Sub

'------------------------------------------------------------------------------
' FR: Actualise Apply Message Type Badge Layout sans modifier les regles metier qui produisent les donnees.
' EN: Refreshes Apply Message Type Badge Layout without changing the business rules that produce the data.
'------------------------------------------------------------------------------

Private Sub ApplyMessageTypeBadgeLayout()

    Const BADGE_HEIGHT As Single = 20

    Dim oldTop As Single
    Dim oldHeight As Single

    On Error Resume Next

    With lblMessageType

        oldTop = .top
        oldHeight = .height

        .AutoSize = False
        .WordWrap = False
        .TextAlign = fmTextAlignCenter

        .height = BADGE_HEIGHT

        If oldHeight > BADGE_HEIGHT Then
            .top = oldTop + ((oldHeight - BADGE_HEIGHT) / 2)
        End If

        .font.bold = True
        .font.size = 9

        .BorderStyle = fmBorderStyleSingle
        .SpecialEffect = fmSpecialEffectFlat

    End With

    On Error GoTo 0

End Sub








Private Sub PEBegin()
    If Not mPE Is Nothing Then mPE.BeginUpdate
End Sub
Private Sub PESync()
    If mPE Is Nothing Then
        Set mPE = New PEConsolePresentation
        mPE.Attach Me, Array("cmdPrevious", "cmdNext", "cmdClose", "cmdClearAck", "cmdClearHistory"), "cmdNext"
    End If
    PEConsole_Theme Me
    mPE.Sync
    PEConsole_Badge Me
End Sub
Private Sub PEEnd()
    If Not mPE Is Nothing Then mPE.Detach
    Set mPE = Nothing
    PEConsole_ClearBadge Me
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


