VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PEConsolePresentation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private mButtons As Collection, mOwner As Object, mProjected As Boolean, mClosing As Boolean
Private mDefaultName As String, mCancelName As String
Public Sub Attach(ByVal owner As Object, ByVal names As Variant, ByVal primaryNames As String)
    Dim name As Variant, Button As PEConsoleButton
    Set mOwner = owner: Set mButtons = New Collection
    For Each name In names
        Set Button = New PEConsoleButton
        Button.Attach owner, owner.Controls(CStr(name)), owner.Controls("peFocus_" & CStr(name)), InStr(1, "|" & primaryNames & "|", "|" & CStr(name) & "|", vbBinaryCompare) > 0
        mButtons.Add Button, CStr(name)
    Next name
End Sub
Public Sub BeginUpdate()
    Dim Button As PEConsoleButton
    If mClosing Or Not mProjected Then Exit Sub
    mOwner.Controls("peDefault").Default = False
    mOwner.Controls("peCancel").Cancel = False
    For Each Button In mButtons
        Button.Restore
    Next Button
    mProjected = False
End Sub
Public Sub Sync()
    Dim Button As PEConsoleButton, reference As MSForms.CommandButton
    If mClosing Then Exit Sub
    If mProjected Then Err.Raise 5, "PEConsolePresentation", "BeginUpdate required before Sync"
    mDefaultName = "": mCancelName = ""
    For Each Button In mButtons
        Button.Sync
        Set reference = Button.reference
        If Button.NeedsDefault Then mDefaultName = reference.name: PlaceRelay "peDefault", reference
        If Button.NeedsCancel Then mCancelName = reference.name: PlaceRelay "peCancel", reference
    Next Button
    mOwner.Controls("peDefault").Default = Len(mDefaultName) > 0
    mOwner.Controls("peCancel").Cancel = Len(mCancelName) > 0
    mProjected = True
End Sub
Private Sub PlaceRelay(ByVal name As String, ByVal reference As MSForms.CommandButton)
    With mOwner.Controls(name)
        .Move reference.left + reference.width / 2, reference.top + reference.height / 2, 1, 1
        .caption = "": .TabStop = False: .Visible = True: .ZOrder 1
    End With
End Sub
Public Sub InvokeDefault()
    If Not mClosing And Len(mDefaultName) > 0 Then mOwner.PEInvoke mDefaultName
End Sub
Public Sub InvokeCancel()
    If Not mClosing And Len(mCancelName) > 0 Then mOwner.PEInvoke mCancelName
End Sub
Public Sub FocusChanged(ByVal name As String, ByVal focused As Boolean)
    If mClosing Or mButtons Is Nothing Then Exit Sub
    mButtons(name).FocusChanged focused
End Sub
Public Sub Hover(ByVal name As String)
    Dim Button As PEConsoleButton
    If mClosing Or mButtons Is Nothing Then Exit Sub
    For Each Button In mButtons
        If "pe" & Button.reference.name <> name Then Button.ClearHover
    Next Button
End Sub
Public Sub Detach()
    Dim Button As PEConsoleButton
    If mClosing Then Exit Sub
    BeginUpdate
    mClosing = True
    If Not mButtons Is Nothing Then
        For Each Button In mButtons
            Button.Detach
        Next Button
    End If
    Set mButtons = Nothing: Set mOwner = Nothing
End Sub
