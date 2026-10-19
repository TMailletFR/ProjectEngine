VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PEConsoleButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private WithEvents mButton As PEButton
Attribute mButton.VB_VarHelpID = -1
Private mReference As MSForms.CommandButton, mRelay As MSForms.TextBox
Private mOwner As Object
Private mVisible As Boolean, mTabStop As Boolean, mDefault As Boolean, mCancel As Boolean
Private mPrimary As Boolean
Public Sub Attach(ByVal owner As Object, ByVal reference As MSForms.CommandButton, ByVal relay As MSForms.TextBox, ByVal primary As Boolean)
    Set mOwner = owner: Set mReference = reference: Set mRelay = relay
    mTabStop = reference.TabStop: mPrimary = primary
    Set mButton = New PEButton
    mButton.BeginUpdate
    mButton.AttachReference owner, relay, reference
End Sub
Public Sub Restore()
    mReference.Visible = mVisible: mReference.TabStop = mTabStop
    mReference.Default = mDefault: mReference.Cancel = mCancel
End Sub
Public Sub Sync()
    mVisible = mReference.Visible: mDefault = mReference.Default: mCancel = mReference.Cancel
    mButton.BeginUpdate
    mButton.Typography mReference.font.name, mReference.font.size, mReference.font.bold, mReference.font.italic
    mButton.ChoiceContent = (mReference.name = "tutorialDiscovery" Or mReference.name = "tutorialPractice")
    If mPrimary Then
        mButton.Appearance RGB(40, 107, 145), vbWhite, RGB(40, 107, 145), RGB(34, 94, 129), RGB(27, 77, 106)
    Else
        mButton.Appearance RGB(231, 237, 241), mReference.ForeColor, RGB(199, 211, 220), RGB(219, 229, 236), RGB(204, 218, 228)
    End If
    mButton.Move mReference.left, mReference.top, mReference.width, mReference.height
    mButton.caption = mReference.caption
    mButton.Enabled = mReference.Enabled
    mButton.EndUpdate
    mButton.Visible = mVisible
    If Not mButton.IsClassic Then
        mReference.Visible = False: mReference.TabStop = False
        mReference.Default = False: mReference.Cancel = False
        mRelay.TabStop = mTabStop
    End If
End Sub
Public Property Get NeedsDefault() As Boolean
    NeedsDefault = mVisible And mDefault And mReference.Enabled And Not mButton.IsClassic
End Property
Public Property Get NeedsCancel() As Boolean
    NeedsCancel = mVisible And mCancel And mReference.Enabled And Not mButton.IsClassic
End Property
Public Property Get reference() As MSForms.CommandButton
    Set reference = mReference
End Property
Public Sub FocusChanged(ByVal focused As Boolean)
    If Not mButton Is Nothing Then mButton.FocusChanged focused
End Sub
Public Sub ClearHover()
    If Not mButton Is Nothing Then mButton.ClearHover
End Sub
Public Sub Detach()
    If Not mButton Is Nothing Then mButton.Detach
    Set mButton = Nothing: Set mOwner = Nothing: Set mReference = Nothing: Set mRelay = Nothing
End Sub
Private Sub mButton_Click()
    If mReference.Enabled And mVisible Then mOwner.PEInvoke mReference.name
End Sub
