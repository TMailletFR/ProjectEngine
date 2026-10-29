VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmFinanceResolution 
   Caption         =   "UserForm1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "frmFinanceResolution.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmFinanceResolution"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Accepted As Boolean
Public decisions As Object
Private mTargets As Collection, mRemoval As Boolean, mPE As PEConsolePresentation
Private mDuplicate As Boolean, mRebuilding As Boolean
Private mClickX As Single
Private mFuture As Boolean
Private WithEvents mChoice As MSForms.ComboBox
Attribute mChoice.VB_VarHelpID = -1
Private WithEvents mList As MSForms.ListBox
Attribute mList.VB_VarHelpID = -1
Private WithEvents mDefault As MSForms.CommandButton
Attribute mDefault.VB_VarHelpID = -1
Private WithEvents mCancel As MSForms.CommandButton
Attribute mCancel.VB_VarHelpID = -1

Public Sub Prepare(ByVal targets As Collection, Optional ByVal removal As Boolean = False)
    mDuplicate = False: mFuture = False
    Set mTargets = targets: mRemoval = removal: Accepted = False
    Set decisions = CreateObject("Scripting.Dictionary")
    FinanceUI_Init Me, Finance_Text("MANAGE"), 680, 410
    FinanceUI_Label Me, "message", Finance_Text(IIf(removal, "REMOVALS", "MISSING")), 18, 16, 630, 55
    Set mList = FinanceUI_Control(Me, "ListBox", "items", 18, 80, 630, 235)
    FinanceUI_Button Me, "keep", Finance_Text(IIf(removal, "RESTORE", "HISTORY")), 18, 330, 180
    FinanceUI_Button Me, "delete", Finance_Text("DELETE"), 206, 330, 120
    FinanceUI_Button Me, "confirm", Finance_Text("CONFIRM"), 334, 330, 140
    FinanceUI_Button Me, "cancel", Finance_Text("CANCEL"), 482, 330, 166, True
    Set mPE = FinanceUI_Buttons(Me, Array("keep", "delete", "confirm", "cancel"), mDefault, mCancel)
    Rebuild
End Sub

Public Sub PrepareDuplicates(ByVal targets As Collection)
    Dim choices(0 To 2, 0 To 1) As Variant
    Set mTargets = targets: mDuplicate = True: mFuture = False: Accepted = False
    Set decisions = CreateObject("Scripting.Dictionary")
    FinanceUI_Init Me, Finance_Text("DUP_TITLE"), 680, 520
    FinanceUI_Label Me, "message", Finance_Text("DUP_REQUIRED"), 18, 16, 630, 40
    FinanceUI_Label Me, "wbsHeader", "WBS", 18, 60, 95, 20
    FinanceUI_Label Me, "activityHeader", Finance_Text("ACTIVITY_LABEL"), 113, 60, 330, 20
    FinanceUI_Label Me, "relationHeader", Finance_Text("RELATION_LABEL"), 443, 60, 185, 20
    Set mList = FinanceUI_Control(Me, "ListBox", "items", 18, 84, 630, 236)
    mList.columnCount = 3: mList.ColumnWidths = "95 pt;330 pt;185 pt": mList.MultiSelect = fmMultiSelectSingle
    FinanceUI_Label Me, "choiceLabel", Finance_Text("DUP_CHOICE"), 18, 334, 320, 28
    Set mChoice = FinanceUI_Control(Me, "ComboBox", "choice", 350, 330, 298, 28)
    mChoice.style = fmStyleDropDownList: mChoice.columnCount = 2: mChoice.ColumnWidths = "0 pt;175 pt": mChoice.BoundColumn = 1
    mChoice.ListWidth = 203
    choices(0, 0) = "INDEPENDENT": choices(0, 1) = Finance_Text("INDEPENDENT")
    choices(1, 0) = "COPIED": choices(1, 1) = Finance_Text("COPIED")
    choices(2, 0) = "SHARED": choices(2, 1) = Finance_Text("SHARED")
    mChoice.list = choices
    mChoice.visible = False: Me.controls("choiceLabel").visible = False
    FinanceUI_Label Me, "explanation", Finance_Text("DUP_HELP"), 18, 334, 630, 94
    FinanceUI_Button Me, "confirm", Finance_Text("CONFIRM"), 334, 454, 140
    FinanceUI_Button Me, "cancel", Finance_Text("CANCEL"), 482, 454, 166, True
    Set mPE = FinanceUI_Buttons(Me, Array("confirm", "cancel"), mDefault, mCancel)
    RebuildDuplicates
End Sub

Public Sub PrepareFuture(ByVal targets As Collection)
    Dim data() As Variant, row As PEFinanceRow, i As Long
    Set mTargets = targets: mFuture = True: mDuplicate = False: Accepted = False
    Set decisions = CreateObject("Scripting.Dictionary")
    FinanceUI_Init Me, Finance_Text("FUTURE_TITLE"), 680, 410
    FinanceUI_Label Me, "message", Finance_Text("FUTURE_CONFIRM"), 18, 16, 630, 92
    FinanceUI_Label Me, "idHeader", "ID", 18, 120, 50
    FinanceUI_Label Me, "activityHeader", "WBS / " & Finance_Text("ACTIVITY_LABEL"), 78, 120, 430
    FinanceUI_Label Me, "dateHeader", Finance_Text("DATE_LABEL"), 518, 120, 130
    Set mList = FinanceUI_Control(Me, "ListBox", "items", 18, 144, 630, 170)
    mList.columnCount = 3: mList.ColumnWidths = "60 pt;440 pt;110 pt": mList.MultiSelect = fmMultiSelectSingle
    ReDim data(0 To targets.count - 1, 0 To 2)
    For Each row In targets
        data(i, 0) = row("TASK_ID"): data(i, 1) = FinanceUI_Caption(row)
        data(i, 2) = Format$(CDate(row("DATE")), "Short Date"): i = i + 1
    Next row
    mList.list = data
    FinanceUI_Button Me, "confirm", Finance_Text("CONFIRM"), 334, 330, 140
    FinanceUI_Button Me, "cancel", Finance_Text("CANCEL"), 482, 330, 166, True
    Set mPE = FinanceUI_Buttons(Me, Array("confirm", "cancel"), mDefault, mCancel)
End Sub

Private Sub RebuildDuplicates()
    Dim data() As Variant, d As Object, i As Long, selected As Long, top As Long, key As String
    If mTargets.count = 0 Then Exit Sub
    If Not mPE Is Nothing Then mPE.BeginUpdate
    selected = mList.ListIndex: top = mList.TopIndex: mRebuilding = True
    ReDim data(0 To mTargets.count - 1, 0 To 2)
    For Each d In mTargets
        key = CStr(d("WBS")): data(i, 0) = key: data(i, 1) = CStr(d("TASK_NAME"))
        If decisions.exists(key) Then data(i, 2) = Finance_Text(CStr(decisions(key))) Else data(i, 2) = Finance_Text("DUP_CHOOSE")
        i = i + 1
    Next d
    mList.list = data
    If selected < 0 Then selected = 0
    mList.ListIndex = selected
    If top >= 0 And top < mList.ListCount Then mList.TopIndex = top
    Me.controls("confirm").enabled = (decisions.count = mTargets.count)
    mRebuilding = False
    mList_Change
    If Not mPE Is Nothing Then mPE.Sync
End Sub

Private Sub mList_Change()
    Dim key As String, i As Long
    If Not mDuplicate Or mRebuilding Then Exit Sub
    mChoice.visible = False
    If mList.ListIndex < 0 Then Exit Sub
    key = CStr(mList.list(mList.ListIndex, 0)): mRebuilding = True
    Me.controls("choiceLabel").caption = Finance_Text("DUP_CHOICE") & " " & key
    mChoice.ListIndex = -1
    If decisions.exists(key) Then
        For i = 0 To mChoice.ListCount - 1
            If CStr(mChoice.list(i, 0)) = CStr(decisions(key)) Then mChoice.ListIndex = i: Exit For
        Next i
    End If
    mRebuilding = False
End Sub

Private Sub mList_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    mClickX = X
End Sub

Private Sub mList_Click()
    If mRebuilding Then Exit Sub
End Sub

Private Sub mList_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Not mDuplicate Or mRebuilding Or Button <> 1 Then Exit Sub
    If X >= 425 Then OpenRelationCell
End Sub

Private Sub mList_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If Not mDuplicate Then Exit Sub
    If KeyCode = 32 Or KeyCode = 113 Or KeyCode = 13 Then
        KeyCode = 0: OpenRelationCell
    End If
End Sub

Private Sub OpenRelationCell()
    Dim top As Single, rowHeight As Single
    If mList.ListIndex < 0 Then Exit Sub
    rowHeight = mList.font.size + 4
    top = mList.top + 2 + (mList.ListIndex - mList.TopIndex) * rowHeight
    If top + 22 > mList.top + mList.height Then top = mList.top + mList.height - 22
    mChoice.Move mList.left + 425, top, 203, 22
    mChoice.visible = True: mChoice.ZOrder 0: mChoice.SetFocus: mChoice.DropDown
End Sub

Private Sub mChoice_Change()
    If Not mDuplicate Or mRebuilding Then Exit Sub
    If mList.ListIndex < 0 Or mChoice.ListIndex < 0 Then Exit Sub
    decisions(CStr(mList.list(mList.ListIndex, 0))) = CStr(mChoice.list(mChoice.ListIndex, 0))
    RebuildDuplicates
End Sub

Private Sub Rebuild()
    Dim keys As New Collection, labels As New Collection, d As Object, key As String, text As String
    For Each d In mTargets
        key = Finance_Reference(d)
        If mRemoval And decisions.exists(key) Then
            If decisions(key) = "KEEP" Then GoTo NextTarget
        End If
        text = FinanceUI_Caption(d)
        If Finance_Has(d("BUDGET")) Then text = text & " | " & Format$(d("BUDGET"), "#,##0.00")
        If decisions.exists(key) Then text = text & " | " & Finance_Text(IIf(decisions(key) = "KEEP", IIf(mRemoval, "RESTORE", "HISTORY"), "DELETE"))
        keys.Add key: labels.Add text
NextTarget:
    Next d
    FinanceUI_List mList, keys, labels
End Sub

Public Sub PEInvoke(ByVal name As String)
    Dim i As Long, d As Object
    On Error GoTo Failed
    If mFuture And name = "confirm" Then
        Accepted = True: PEConsole_WheelDetach Me: Me.Hide: Exit Sub
    End If
    Select Case name
        Case "cancel": PEConsole_WheelDetach Me: Me.Hide
        Case "keep", "delete"
            For i = 0 To mList.ListCount - 1
                If mList.selected(i) Then decisions(CStr(mList.list(i, 0))) = IIf(name = "keep", "KEEP", "DELETE")
            Next i
            Rebuild
            If mRemoval And mList.ListCount = 0 Then Accepted = True: PEConsole_WheelDetach Me: Me.Hide
        Case "confirm"
            If mDuplicate Then
                If decisions.count <> mTargets.count Then Exit Sub
                Accepted = True: PEConsole_WheelDetach Me: Me.Hide: Exit Sub
            End If
            For Each d In mTargets
                If Not decisions.exists(Finance_Reference(d)) Then Exit Sub
            Next d
            Accepted = True: PEConsole_WheelDetach Me: Me.Hide
    End Select
    Exit Sub
Failed: FinanceUI_Error
End Sub
Private Sub mCancel_Click()
    mPE.InvokeCancel
End Sub
Private Sub mDefault_Click()
    mPE.InvokeDefault
End Sub
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    PEConsole_WheelDetach Me
    If CloseMode = 0 Then Cancel = True: Me.Hide
    If CloseMode <> 0 Then
        If Not mPE Is Nothing Then mPE.Detach
        Set mPE = Nothing
    End If
End Sub
Public Sub PEHover(ByVal name As String)
    If Not mPE Is Nothing Then mPE.Hover name
End Sub
Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    PEHover ""
End Sub
Private Sub peFocus_keep_Enter()
    If Not mPE Is Nothing Then mPE.FocusChanged "keep", True
End Sub
Private Sub peFocus_keep_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Not mPE Is Nothing Then mPE.FocusChanged "keep", False
End Sub
Private Sub peFocus_delete_Enter()
    If Not mPE Is Nothing Then mPE.FocusChanged "delete", True
End Sub
Private Sub peFocus_delete_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Not mPE Is Nothing Then mPE.FocusChanged "delete", False
End Sub
Private Sub peFocus_confirm_Enter()
    If Not mPE Is Nothing Then mPE.FocusChanged "confirm", True
End Sub
Private Sub peFocus_confirm_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Not mPE Is Nothing Then mPE.FocusChanged "confirm", False
End Sub
Private Sub peFocus_cancel_Enter()
    If Not mPE Is Nothing Then mPE.FocusChanged "cancel", True
End Sub
Private Sub peFocus_cancel_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Not mPE Is Nothing Then mPE.FocusChanged "cancel", False
End Sub
Private Sub UserForm_Terminate()
    PEConsole_WheelDetach Me
    Set mPE = Nothing
End Sub

Private Sub UserForm_Activate()
    PEConsole_WheelAttach Me, Me
End Sub
Public Sub PEListScrolled(ByVal name As String)
    If name = "items" And Not mChoice Is Nothing Then mChoice.visible = False
End Sub
Private Sub UserForm_Deactivate()
    PEConsole_WheelDetach Me
End Sub


