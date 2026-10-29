VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmFinanceActivities 
   Caption         =   "UserForm1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "frmFinanceActivities.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmFinanceActivities"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Ready As Boolean
Private mOriginal As Collection, mExpenses As Collection, mActivities As Collection, mChosen As Object, mAvailable As Object
Private mPE As PEConsolePresentation, mSource As String, mRemoved As Object
Private mGraph As Object, mEditScope As Long, mOriginScope As Long, mRevision As Long
Private mInputsValid As Boolean, mChangingProject As Boolean
Private WithEvents mProject As MSForms.ComboBox
Attribute mProject.VB_VarHelpID = -1
Private mProjectName As MSForms.TextBox
Private WithEvents mLeft As MSForms.ListBox
Attribute mLeft.VB_VarHelpID = -1
Private WithEvents mRight As MSForms.ListBox
Attribute mRight.VB_VarHelpID = -1
Private WithEvents mSearch As MSForms.TextBox
Attribute mSearch.VB_VarHelpID = -1
Private WithEvents mTrackedSearch As MSForms.TextBox
Attribute mTrackedSearch.VB_VarHelpID = -1
Private WithEvents mDefault As MSForms.CommandButton
Attribute mDefault.VB_VarHelpID = -1
Private WithEvents mCancel As MSForms.CommandButton
Attribute mCancel.VB_VarHelpID = -1

Public Sub Prepare()
    Dim d As Object, i As Long, j As Long, sorted() As Object, swap As Object
    Ready = False
    FinanceUI_Init Me, Finance_Text("MANAGE"), 830, 550
    mOriginScope = Finance_ScopeId(): mEditScope = mOriginScope: mRevision = Finance_Revision()
    Set mGraph = Finance_ActivityDraft(mInputsValid)
    Set mActivities = Finance_Activities()
    mSource = SourceState(mActivities)
    If Not Finance_ResolveMissing(mGraph, mActivities) Then Exit Sub
    Set mAvailable = CreateObject("Scripting.Dictionary")
    If mActivities.count > 0 Then
        ReDim sorted(1 To mActivities.count)
        For i = 1 To mActivities.count: Set sorted(i) = mActivities(i): Next i
        ' Numeric in-memory ordering; no worksheet sort or per-cell access.
        For i = 2 To UBound(sorted)
            Set swap = sorted(i): j = i - 1
            Do While j >= 1
                If CLng(sorted(j)("TASK_ID")) <= CLng(swap("TASK_ID")) Then Exit Do
                Set sorted(j + 1) = sorted(j): j = j - 1
            Loop
            Set sorted(j + 1) = swap
        Next i
        Set mActivities = New Collection
        For i = 1 To UBound(sorted): mActivities.Add sorted(i): Next i
    End If
    For Each d In mActivities: mAvailable.Add CStr(d("WBS")), d: Next d
    FinanceUI_Label Me, "projectLabel", Finance_Text("PROJECT_LABEL"), 18, 8, 230
    Set mProject = FinanceUI_Control(Me, "ComboBox", "project", 18, 32, 230, 27)
    mProject.style = fmStyleDropDownList: mProject.columnCount = 2: mProject.ColumnWidths = "0 pt;205 pt": mProject.BoundColumn = 1
    FinanceUI_Label Me, "projectNameLabel", Finance_Text("PROJECT_NAME"), 260, 8, 194
    Set mProjectName = FinanceUI_Control(Me, "TextBox", "projectName", 260, 32, 194, 27)
    FinanceUI_Button Me, "projectAdd", Finance_Text("PROJECT_ADD"), 470, 32, 110
    FinanceUI_Button Me, "projectRename", Finance_Text("PROJECT_RENAME"), 590, 32, 90
    FinanceUI_Button Me, "projectDelete", Finance_Text("PROJECT_DELETE"), 690, 32, 110
    FinanceUI_Label Me, "inputStatus", IIf(mInputsValid, "", Finance_Text("INVALID_DRAFT")), 18, 70, 782, 34
    FinanceUI_Label Me, "leftTitle", Finance_Text("AVAILABLE"), 18, 110, 310
    FinanceUI_Label Me, "rightTitle", Finance_Text("TRACKED"), 470, 110, 330
    Set mSearch = FinanceUI_Control(Me, "TextBox", "search", 18, 137, 310, 25)
    mSearch.ControlTipText = Finance_Text("SEARCH")
    Set mTrackedSearch = FinanceUI_Control(Me, "TextBox", "trackedSearch", 470, 137, 330, 25)
    mTrackedSearch.ControlTipText = Finance_Text("SEARCH")
    Set mLeft = FinanceUI_Control(Me, "ListBox", "available", 18, 173, 310, 260)
    Set mRight = FinanceUI_Control(Me, "ListBox", "tracked", 470, 173, 330, 260)
    FinanceUI_Button Me, "add", ">", 350, 218, 90
    FinanceUI_Button Me, "remove", "<", 350, 263, 90
    FinanceUI_Button Me, "all", ">>", 350, 308, 90
    FinanceUI_Button Me, "removeAll", "<<", 350, 353, 90
    FinanceUI_Button Me, "restore", Finance_Text("RESTORE"), 18, 460, 220
    FinanceUI_Button Me, "confirm", Finance_Text("CONFIRM"), 474, 460, 145
    FinanceUI_Button Me, "cancel", Finance_Text("CANCEL"), 635, 460, 165, True
    Set mPE = FinanceUI_Buttons(Me, Array("add", "remove", "all", "removeAll", "restore", "projectAdd", "projectRename", "projectDelete", "confirm", "cancel"), mDefault, mCancel)
    RefreshProjects
    If Not LoadProject Then Exit Sub
    Ready = True
End Sub

Private Sub RefreshProjects()
    Dim options As Variant, i As Long
    If Not mPE Is Nothing Then mPE.BeginUpdate
    mChangingProject = True
    options = Finance_ProjectOptions(mGraph, (mGraph("PROJECTS").count > 1))
    mProject.list = options
    For i = 0 To UBound(options, 1)
        If CLng(options(i, 0)) = mEditScope Then mProject.ListIndex = i: Exit For
    Next i
    mProjectName.text = ""
    If mEditScope <> 0 Then mProjectName.text = CStr(Finance_DomainIndex(mGraph("PROJECTS"), "PROJECT_ID")(CStr(mEditScope))("NAME"))
    Me.controls("projectRename").enabled = (mEditScope <> 0)
    Me.controls("projectDelete").enabled = (mEditScope <> 0)
    mChangingProject = False
    mProject.visible = (mGraph("PROJECTS").count > 1)
    FinanceUI_Label Me, "singleProject", IIf(mEditScope = 0, "", mProjectName.text), 18, 32, 230, 27
    Me.controls("singleProject").visible = Not mProject.visible
    If Not mPE Is Nothing Then mPE.Sync
End Sub

Private Function LoadProject() As Boolean
    Dim d As Object
    Set mOriginal = New Collection: Set mExpenses = New Collection
    If mEditScope <> 0 Then
        Finance_DomainViews mGraph, mEditScope, mOriginal, mExpenses
    End If
    Set mChosen = CreateObject("Scripting.Dictionary"): Set mRemoved = CreateObject("Scripting.Dictionary")
    For Each d In mOriginal: mChosen.Add Finance_Reference(d), d: Next d
    Dim controlName As Variant
    If Not mPE Is Nothing Then mPE.BeginUpdate
    For Each controlName In Array("available", "tracked", "search", "trackedSearch", "leftTitle", "rightTitle", "add", "remove", "all", "removeAll", "restore")
        Me.controls(CStr(controlName)).visible = (mEditScope <> 0)
    Next controlName
    FinanceUI_Label Me, "globalMessage", Finance_Text("GLOBAL_MANAGE_HINT"), 80, 190, 670, 100
    Me.controls("globalMessage").visible = (mEditScope = 0)
    Call Rebuild
    LoadProject = True
End Function

Private Function StageProject() As Boolean
    Dim rows As Collection, targets As Collection, decisions As Object, nextGraph As Object
    If mEditScope = 0 Then StageProject = True: Exit Function
    If Not ReviewRemovals Then Rebuild: Exit Function
    Set rows = Finance_SelectedRows(mOriginal, mChosen, mActivities)
    Set targets = Finance_DuplicationTargets(mGraph, mEditScope, rows)
    Set decisions = CreateObject("Scripting.Dictionary")
    If targets.count > 0 Then
        frmFinanceResolution.PrepareDuplicates targets
        frmFinanceResolution.Show vbModal
        If Not frmFinanceResolution.Accepted Then Unload frmFinanceResolution: Exit Function
        Set decisions = frmFinanceResolution.decisions
        Unload frmFinanceResolution
    End If
    Set nextGraph = FinanceStore_Clone(mGraph)
    Finance_StageActivities nextGraph, mEditScope, rows, decisions
    Set mGraph = nextGraph
    StageProject = True
End Function

Private Sub mProject_Change()
    Dim oldScope As Long, selected As Long
    If mChangingProject Or mProject.ListIndex < 0 Then Exit Sub
    On Error GoTo Failed
    oldScope = mEditScope: selected = CLng(mProject.list(mProject.ListIndex, 0))
    If selected = oldScope Then Exit Sub
    If Not StageProject Then RefreshProjects: Exit Sub
    mEditScope = selected
    If Not LoadProject Then mEditScope = oldScope: LoadProject
    RefreshProjects
    Exit Sub
Failed:
    mEditScope = oldScope: RefreshProjects: FinanceUI_Error
End Sub

Private Function SourceState(ByVal activities As Collection) As String
    Dim d As Object, value As String, key As Variant
    For Each d In activities
        For Each key In Array("WBS", "TASK_ID", "TASK_NAME", "PROJECT", "PARENT")
            value = CStr(d(key)): SourceState = SourceState & Len(value) & ":" & value
        Next key
    Next d
End Function

Private Sub Rebuild()
    Dim lk As New Collection, ll As New Collection, rk As New Collection, rl As New Collection, d As Object, key As Variant, search As String, trackedSearch As String
    If mChosen Is Nothing Or mLeft Is Nothing Or mRight Is Nothing Then Exit Sub
    search = LCase$(Trim$(mSearch.text))
    trackedSearch = LCase$(Trim$(mTrackedSearch.text))
    If mEditScope = 0 Then GoTo ListsReady
    For Each d In mActivities
        If Not CBool(d("PARENT")) And Not mChosen.exists("ACTIVE|" & d("WBS")) Then
            If InStr(1, LCase$(FinanceUI_Caption(d)), search, vbBinaryCompare) > 0 Or search = "" Then
                lk.Add CStr(d("WBS")): ll.Add FinanceUI_Caption(d)
            End If
        End If
        If mChosen.exists("ACTIVE|" & d("WBS")) Then
            If InStr(1, LCase$(FinanceUI_Caption(d)), trackedSearch, vbBinaryCompare) > 0 Or trackedSearch = "" Then
                rk.Add "ACTIVE|" & CStr(d("WBS")): rl.Add FinanceUI_Caption(d)
            End If
        End If
    Next d
    For Each key In mChosen.keys
        Set d = mChosen(key)
        If d("STATE") = "HISTORICAL" Then
            If InStr(1, LCase$(FinanceUI_Caption(d)), trackedSearch, vbBinaryCompare) > 0 Or trackedSearch = "" Then rk.Add CStr(key): rl.Add FinanceUI_Caption(d)
        End If
    Next key
ListsReady:
    FinanceUI_List mLeft, lk, ll: FinanceUI_List mRight, rk, rl
    If Not mPE Is Nothing Then mPE.BeginUpdate
    Me.controls("restore").enabled = (mRemoved.count > 0)
    Me.controls("add").enabled = (lk.count > 0): Me.controls("all").enabled = (lk.count > 0)
    Me.controls("remove").enabled = (rk.count > 0): Me.controls("removeAll").enabled = (rk.count > 0)
    If Not mPE Is Nothing Then mPE.Sync
End Sub

Private Sub AddOne(ByVal wbs As String)
    Dim d As Object, source As Object, key As Variant, old As Object
    If mChosen.exists("ACTIVE|" & wbs) Then Exit Sub
    If mRemoved.exists("ACTIVE|" & wbs) Then
        mChosen.Add "ACTIVE|" & wbs, mRemoved("ACTIVE|" & wbs)
        mRemoved.Remove "ACTIVE|" & wbs
        Exit Sub
    End If
    For Each old In mOriginal
        If Finance_Reference(old) = "ACTIVE|" & wbs Then mChosen.Add "ACTIVE|" & wbs, old: Exit Sub
    Next old
    Set d = Finance_NewRow(FIN_TABLE): Set source = mAvailable(wbs)
    For Each key In Array("WBS", "TASK_NAME", "STATE", "TASK_ID", "PROJECT", "PROGRESS"): d(key) = source(key): Next key
    mChosen.Add Finance_Reference(d), d
End Sub

Private Sub mTrackedSearch_Change()
    Rebuild
End Sub

Private Function ReviewRemovals() As Boolean
    Dim targets As Collection, d As Object, key As String
    Set targets = Finance_RemovalTargets(mOriginal, mChosen, mExpenses)
    If targets.count = 0 Then ReviewRemovals = True: Exit Function
    frmFinanceResolution.Prepare targets, True
    frmFinanceResolution.Show vbModal
    For Each d In targets
        key = Finance_Reference(d)
        If frmFinanceResolution.decisions.exists(key) Then
            If frmFinanceResolution.decisions(key) = "KEEP" Then
                If Not mChosen.exists(key) Then mChosen.Add key, d
                If mRemoved.exists(key) Then mRemoved.Remove key
            End If
        End If
    Next d
    If frmFinanceResolution.Accepted Then ReviewRemovals = (Finance_RemovalTargets(mOriginal, mChosen, mExpenses).count > 0)
    Unload frmFinanceResolution
End Function

Public Sub PEInvoke(ByVal name As String)
    Dim i As Long, rows As New Collection, expenses As New Collection, d As Object, key As Variant
    On Error GoTo Failed
    Select Case name
        Case "cancel": PEConsole_WheelDetach Me: Me.Hide
        Case "add", "all"
            For i = 0 To mLeft.ListCount - 1
                If name = "all" Or mLeft.selected(i) Then AddOne CStr(mLeft.list(i, 0))
            Next i
        Case "remove", "removeAll"
            For i = 0 To mRight.ListCount - 1
                If name = "removeAll" Or mRight.selected(i) Then
                    key = CStr(mRight.list(i, 0))
                    If Not mRemoved.exists(key) Then mRemoved.Add key, mChosen(key)
                    mChosen.Remove key
                End If
            Next i
        Case "restore": RestoreAll
        Case "projectAdd"
            If Not StageProject Then Rebuild: Exit Sub
            mEditScope = Finance_ProjectAdd(mGraph, mProjectName.text)
            RefreshProjects
            If Not LoadProject Then Exit Sub
        Case "projectRename"
            If mEditScope = 0 Then Exit Sub
            Finance_ProjectRename mGraph, mEditScope, mProjectName.text
            RefreshProjects
        Case "projectDelete"
            If mEditScope = 0 Then Exit Sub
            If MsgBox(mProjectName.text & vbCrLf & vbCrLf & Finance_Text("PROJECT_DELETE_CONFIRM"), vbYesNo + vbDefaultButton2 + vbExclamation, Finance_Text("PROJECT_DELETE")) <> vbYes Then Exit Sub
            Finance_ProjectDelete mGraph, mEditScope
            mEditScope = CLng(mGraph("META")(1)("SCOPE_ID"))
            RefreshProjects
            If Not LoadProject Then Exit Sub
        Case "confirm"
            If Not mInputsValid Then
                If Finance_DomainIndex(mGraph("PROJECTS"), "PROJECT_ID").exists(CStr(mOriginScope)) Then Err.Raise 5, , Finance_Text("INVALID_DRAFT")
            End If
            If Not StageProject Then Rebuild: Exit Sub
            If mSource <> SourceState(Finance_Activities()) Then Err.Raise 5, , Finance_Text("SOURCE_CHANGED")
            Finance_CheckRevision mOriginScope, mRevision
            Finance_CommitProjectDraft mGraph, mEditScope
            PEConsole_WheelDetach Me: Me.Hide: Exit Sub
    End Select
    Call Rebuild: Exit Sub
Failed: FinanceUI_Error
End Sub
Private Sub RestoreAll()
    Dim key As Variant
    For Each key In mRemoved.keys
        If Not mChosen.exists(key) Then mChosen.Add key, mRemoved(key)
    Next key
    mRemoved.RemoveAll
End Sub
Private Sub mSearch_Change()
    Rebuild
End Sub
Private Sub mLeft_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Cancel = True
    If mLeft.ListIndex >= 0 Then AddOne CStr(mLeft.list(mLeft.ListIndex, 0)): Rebuild
End Sub
Private Sub mRight_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Cancel = True
    If mRight.ListIndex >= 0 Then
        Dim key As String: key = CStr(mRight.list(mRight.ListIndex, 0))
        If Not mRemoved.exists(key) Then mRemoved.Add key, mChosen(key)
        mChosen.Remove key: Rebuild
    End If
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
Private Sub peFocus_add_Enter()
    If Not mPE Is Nothing Then mPE.FocusChanged "add", True
End Sub
Private Sub peFocus_add_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Not mPE Is Nothing Then mPE.FocusChanged "add", False
End Sub
Private Sub peFocus_remove_Enter()
    If Not mPE Is Nothing Then mPE.FocusChanged "remove", True
End Sub
Private Sub peFocus_remove_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Not mPE Is Nothing Then mPE.FocusChanged "remove", False
End Sub
Private Sub peFocus_all_Enter()
    If Not mPE Is Nothing Then mPE.FocusChanged "all", True
End Sub
Private Sub peFocus_all_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Not mPE Is Nothing Then mPE.FocusChanged "all", False
End Sub
Private Sub peFocus_restore_Enter()
    If Not mPE Is Nothing Then mPE.FocusChanged "restore", True
End Sub
Private Sub peFocus_removeAll_Enter()
    If Not mPE Is Nothing Then mPE.FocusChanged "removeAll", True
End Sub
Private Sub peFocus_removeAll_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Not mPE Is Nothing Then mPE.FocusChanged "removeAll", False
End Sub
Private Sub peFocus_restore_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Not mPE Is Nothing Then mPE.FocusChanged "restore", False
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
    PEConsole_WheelAttach Me
End Sub
Private Sub UserForm_Deactivate()
    PEConsole_WheelDetach Me
End Sub


