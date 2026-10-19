VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PEButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Event Click()
Private mSurface As Object
Private WithEvents mImage As MSForms.Image
Attribute mImage.VB_VarHelpID = -1
Private WithEvents mLabel As MSForms.Label
Attribute mLabel.VB_VarHelpID = -1
Private WithEvents mFocus As MSForms.TextBox
Attribute mFocus.VB_VarHelpID = -1
Private WithEvents mClassic As MSForms.CommandButton
Attribute mClassic.VB_VarHelpID = -1
Private mOwner As Object, mName As String, mCaption As String
Private mEnabled As Boolean, mHovered As Boolean, mDown As Boolean, mInside As Boolean
Private mFocused As Boolean, mKey As Long, mSuppressUp As Boolean
Private mRadius As Single, mFill As Long, mInk As Long, mBorder As Long
Private mHoverFill As Long, mPressedFill As Long
Private mFontSize As Single
Private mFontName As String, mFontBold As Boolean, mFontItalic As Boolean
Private mReference As MSForms.CommandButton
Private mUpdating As Boolean
Private mChoice As Boolean
Public Property Let ChoiceContent(ByVal value As Boolean)
    mChoice = value
End Property
Public Sub BeginUpdate()
    mUpdating = True
End Sub
Public Sub EndUpdate()
    mUpdating = False
    Repaint
End Sub
Public Sub AttachReference(ByVal owner As Object, ByVal focusRelay As MSForms.TextBox, ByVal reference As MSForms.CommandButton)
    Dim wasVisible As Boolean
    wasVisible = reference.Visible
    Set mReference = reference
    Attach owner, focusRelay, "pe" & reference.name, reference.left, reference.top, reference.width, reference.height, reference.tabIndex
    reference.Visible = wasVisible
End Sub
Public Sub Attach(ByVal owner As Object, ByVal focusRelay As MSForms.TextBox, ByVal name As String, ByVal left As Single, ByVal top As Single, _
    ByVal width As Single, ByVal height As Single, ByVal tabIndex As Long)
    Dim failure As String
    On Error GoTo Recoverable
    Set mOwner = owner: mName = name
    mEnabled = True: mRadius = 8: mFontSize = 12: mFontName = "Segoe UI"
    If Not mReference Is Nothing Then
        mFontName = mReference.font.name: mFontSize = mReference.font.size
        mFontBold = mReference.font.bold: mFontItalic = mReference.font.italic
        mCaption = mReference.caption: mEnabled = mReference.Enabled
    End If
    mFill = RGB(20, 105, 98): mInk = vbWhite: mBorder = RGB(20, 105, 98)
    mHoverFill = RGB(17, 90, 84): mPressedFill = RGB(12, 73, 68)
    Set mFocus = focusRelay
    With mFocus
        .left = left + width / 2: .top = top + height / 2: .width = 1: .height = 1
        .BorderStyle = 0: .BackStyle = 0: .Locked = True
        .TabStop = True: .tabIndex = tabIndex: .TabKeyBehavior = False
    End With
    If PEPresentation_ClassicRequested() Then
        CreateClassic left, top, width, height
        Exit Sub
    End If
    PEPresentation_CheckInitialization
    CreateSurface "IMAGE", left, top, width, height
    Repaint
    Exit Sub
Recoverable:
    failure = CStr(Err.number) & ": " & Err.description
    RemoveSurface
    PEPresentation_RecordFallback failure
    CreateClassic left, top, width, height
End Sub
Private Sub CreateClassic(ByVal left As Single, ByVal top As Single, ByVal width As Single, ByVal height As Single)
    If mReference Is Nothing Then
        Set mClassic = mOwner.Controls.Add("Forms.CommandButton.1", mName & "Classic", True)
    Else
        Set mClassic = mReference
    End If
    mClassic.Move left, top, width, height
    mClassic.tabIndex = mFocus.tabIndex
    mFocus.TabStop = False: mFocus.Visible = False
    mClassic.Visible = True
    Repaint
End Sub
Private Sub RemoveSurface()
    If mSurface Is Nothing Then Exit Sub
    Set mSurface.picture = Nothing
    Set mImage = Nothing: Set mLabel = Nothing: Set mSurface = Nothing
    mOwner.Controls.Remove mName & "Surface"
End Sub
Private Sub UseClassic(ByVal reason As String)
    Dim left As Single, top As Single, width As Single, height As Single
    If Not mClassic Is Nothing Then Exit Sub
    left = mSurface.left: top = mSurface.top: width = mSurface.width: height = mSurface.height
    RemoveSurface
    PEPresentation_RecordFallback reason
    CreateClassic left, top, width, height
End Sub
Public Property Get IsClassic() As Boolean
    IsClassic = Not mClassic Is Nothing
End Property
Private Sub CreateSurface(ByVal kind As String, ByVal left As Single, ByVal top As Single, ByVal width As Single, ByVal height As Single)
    If kind = "LABEL" Then
        Set mLabel = mOwner.Controls.Add("Forms.Label.1", mName & "Surface", True)
        Set mSurface = mLabel
        mLabel.caption = "": mLabel.PicturePosition = 12
    Else
        Set mImage = mOwner.Controls.Add("Forms.Image.1", mName & "Surface", True)
        Set mSurface = mImage
        mImage.PictureSizeMode = 1
    End If
    With mSurface
        .left = left: .top = top: .width = width: .height = height
        .BackStyle = 0: .BorderStyle = 0
        .Enabled = mEnabled
    End With
    mSurface.ZOrder 0
End Sub
Public Property Let SurfaceKind(ByVal kind As String)
    Dim left As Single, top As Single, width As Single, height As Single
    kind = UCase$(kind)
    If Not mClassic Is Nothing Then Exit Property
    If kind <> "IMAGE" And kind <> "LABEL" Then Err.Raise 5, , "SurfaceKind: IMAGE or LABEL"
    left = mSurface.left: top = mSurface.top: width = mSurface.width: height = mSurface.height
    Set mSurface.picture = Nothing
    Set mImage = Nothing: Set mLabel = Nothing: Set mSurface = Nothing
    mOwner.Controls.Remove mName & "Surface"
    mHovered = False: mDown = False
    CreateSurface kind, left, top, width, height
    Repaint
End Property
Public Sub Detach()
    RemoveSurface
    Set mImage = Nothing: Set mLabel = Nothing
    Set mClassic = Nothing
    Set mReference = Nothing
    Set mSurface = Nothing: Set mFocus = Nothing: Set mOwner = Nothing
End Sub
Public Property Get caption() As String
    caption = mCaption
End Property
Public Property Let caption(ByVal value As String)
    mCaption = value: Repaint
End Property
Public Property Get Enabled() As Boolean
    Enabled = mEnabled
End Property
Public Property Let Enabled(ByVal value As Boolean)
    mEnabled = value: mDown = False: mKey = 0: mHovered = False
    mFocus.Enabled = value
    If Not mSurface Is Nothing Then mSurface.Enabled = value
    If Not mClassic Is Nothing Then mClassic.Enabled = value
    Repaint
End Property
Public Property Get CornerRadius() As Single
    CornerRadius = mRadius
End Property
Public Property Get fontSize() As Single
    fontSize = mFontSize
End Property
Public Property Let fontSize(ByVal value As Single)
    If value < 8 Or value > 24 Then Err.Raise 5, , "FontSize must be between 8 and 24 pt"
    mFontSize = value: Repaint
End Property
Public Property Let CornerRadius(ByVal value As Single)
    If value < 0 Then Err.Raise 5, , "CornerRadius must be positive"
    mRadius = value: Repaint
End Property
Public Sub Appearance(ByVal fill As Long, ByVal ink As Long, ByVal border As Long, ByVal Hover As Long, ByVal pressed As Long)
    mFill = fill: mInk = ink: mBorder = border: mHoverFill = Hover: mPressedFill = pressed
    Repaint
End Sub
Public Sub Typography(ByVal name As String, ByVal size As Single, ByVal bold As Boolean, ByVal italic As Boolean)
    mFontName = name: mFontSize = size: mFontBold = bold: mFontItalic = italic
End Sub
Public Property Let Visible(ByVal value As Boolean)
    If Not mSurface Is Nothing Then mSurface.Visible = value
    If Not mClassic Is Nothing Then mClassic.Visible = value
    If Not mFocus Is Nothing Then mFocus.Visible = value And (mClassic Is Nothing)
End Property
Public Sub Move(ByVal left As Single, ByVal top As Single, ByVal width As Single, ByVal height As Single)
    If width < 16 Or height < 16 Then Err.Raise 5, , "PEButton minimum size is 16 pt"
    If Not mClassic Is Nothing Then
        mClassic.Move left, top, width, height
    Else
        mSurface.Move left, top, width, height
    End If
    mFocus.Move left + width / 2, top + height / 2, 1, 1
    Repaint
End Sub
Public Sub ClearHover()
    If mHovered Then mHovered = False: Repaint
End Sub
Public Sub SetFocus()
    If Not mEnabled Then Exit Sub
    If Not mClassic Is Nothing Then
        mClassic.SetFocus
    Else
        mFocus.SetFocus
    End If
End Sub
Private Sub Repaint()
    Dim fill As Long, ink As Long, edge As Long
    Dim failure As String
    If mUpdating Then Exit Sub
    On Error GoTo Recoverable
    If Not mClassic Is Nothing Then
        mClassic.caption = mCaption: mClassic.Enabled = mEnabled
        mClassic.font.name = mFontName: mClassic.font.size = mFontSize
        mClassic.font.bold = mFontBold: mClassic.font.italic = mFontItalic
        Exit Sub
    End If
    If mSurface Is Nothing Then Exit Sub
    If Not PEWMFCaptionSupported(mCaption) Then
        UseClassic "Caption cannot be represented faithfully in legacy WMF"
        Exit Sub
    End If
    fill = mFill: ink = mInk: edge = mBorder
    If mHovered Then fill = mHoverFill
    If (mDown And mInside) Or mKey <> 0 Then fill = mPressedFill
    If mFocused Then edge = RGB(47, 96, 156)
    If Not mEnabled Then fill = RGB(229, 233, 237): ink = RGB(128, 137, 147): edge = RGB(209, 216, 223)
    Set mSurface.picture = PEVector(mCaption, mSurface.width, mSurface.height, mRadius, fill, ink, edge, mFocused And mEnabled, mFontSize, mFontName, mFontBold, mFontItalic, mChoice)
    Exit Sub
Recoverable:
    failure = CStr(Err.number) & ": " & Err.description
    If mSurface Is Nothing Or Not mClassic Is Nothing Then Err.Raise 5, , failure
    UseClassic failure
End Sub
Private Sub mClassic_Click()
    ' A bound reference already owns its native Click handler.
    If mReference Is Nothing Then
        If mEnabled Then RaiseEvent Click
    End If
End Sub
Public Sub ClassicClickGate()
    If mClassic Is Nothing Then Err.Raise 5, , "Classic control absent"
    mClassic.value = True
End Sub
Private Function Hit(ByVal X As Single, ByVal Y As Single) As Boolean
    Dim r As Single, dx As Single, dy As Single, w As Single, h As Single
    w = mSurface.width: h = mSurface.height
    If X < 1 Or Y < 1 Or X > w - 1 Or Y > h - 1 Then Exit Function
    r = mRadius: If r > w / 2 Then r = w / 2
    If r > h / 2 Then r = h / 2
    If X < r Then dx = r - X
    If X > w - r Then dx = X - (w - r)
    If Y < r Then dy = r - Y
    If Y > h - r Then dy = Y - (h - r)
    Hit = (dx * dx + dy * dy <= r * r)
End Function
Private Sub Press(ByVal X As Single, ByVal Y As Single)
    If Not mEnabled Or Not Hit(X, Y) Then Exit Sub
    mFocus.SetFocus
    mDown = True: mInside = True: mSuppressUp = False
    Repaint
End Sub
Private Sub ReleasePointer(ByVal X As Single, ByVal Y As Single)
    Dim fire As Boolean
    fire = mEnabled And mDown And Hit(X, Y) And Not mSuppressUp
    mDown = False: mSuppressUp = False: Repaint
    If fire Then RaiseEvent Click
End Sub
Private Sub mImage_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button = 1 Then Press X, Y
End Sub
Private Sub mImage_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button = 1 Then ReleasePointer X, Y
End Sub
Private Sub PointerMove(ByVal X As Single, ByVal Y As Single)
    Dim inside As Boolean
    inside = Hit(X, Y)
    If inside Then mOwner.PEHover mName
    If inside <> mHovered Or inside <> mInside Then
        mInside = inside: mHovered = inside: Repaint
    End If
End Sub
Private Sub mImage_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    PointerMove X, Y
End Sub
Private Sub mImage_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Cancel = True: mSuppressUp = True
End Sub
Private Sub mLabel_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button = 1 Then Press X, Y
End Sub
Private Sub mLabel_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button = 1 Then ReleasePointer X, Y
End Sub
Private Sub mLabel_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    PointerMove X, Y
End Sub
Private Sub mLabel_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Cancel = True: mSuppressUp = True
End Sub
Public Sub FocusChanged(ByVal focused As Boolean)
    mFocused = focused
    If Not focused Then mDown = False: mKey = 0
    Repaint
End Sub
Private Sub mFocus_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If Not mEnabled Then Exit Sub
    Select Case CLng(KeyCode)
        Case 13, 32
            If mKey = 0 Then mKey = CLng(KeyCode): Repaint
            KeyCode = 0
    End Select
End Sub
Private Sub mFocus_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If mEnabled And mKey = CLng(KeyCode) Then
        mKey = 0: KeyCode = 0: Repaint: RaiseEvent Click
    End If
End Sub
Private Sub mFocus_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = 0
End Sub
Public Function ContractGate() As String
    Dim savedEnabled As Boolean
    If Hit(0, 0) Then Err.Raise 5, , "Corner hit-test failed"
    mDown = True: mInside = True
    ReleasePointer -5, -5
    If mDown Then Err.Raise 5, , "Drag-out cleanup failed"
    mDown = True
    ReleasePointer mSurface.width / 2, mSurface.height / 2
    ReleasePointer mSurface.width / 2, mSurface.height / 2
    savedEnabled = mEnabled: mEnabled = False: mDown = True
    ReleasePointer mSurface.width / 2, mSurface.height / 2
    mEnabled = savedEnabled: mDown = True: mSuppressUp = True
    ReleasePointer mSurface.width / 2, mSurface.height / 2
    ContractGate = "corner_rejected=True;drag_out_disarmed=True;one_click_expected=True;picture_type=" & mSurface.picture.Type
End Function
