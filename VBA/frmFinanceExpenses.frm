VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmFinanceExpenses 
   Caption         =   "UserForm1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "frmFinanceExpenses.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmFinanceExpenses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mRows As Collection, mExpenses As Collection, mPE As PEConsolePresentation
Private mIndex As Long
Private mFiltering As Boolean, mSelectedReference As String
Private mScope As Long, mRevision As Long, mProjectName As String
Private WithEvents mActivity As MSForms.ComboBox
Attribute mActivity.VB_VarHelpID = -1
Private WithEvents mFilterWBS As MSForms.ComboBox
Attribute mFilterWBS.VB_VarHelpID = -1
Private WithEvents mFilterDate As MSForms.ComboBox
Attribute mFilterDate.VB_VarHelpID = -1
Private WithEvents mFilterAmount As MSForms.ComboBox
Attribute mFilterAmount.VB_VarHelpID = -1
Private WithEvents mFilterComment As MSForms.ComboBox
Attribute mFilterComment.VB_VarHelpID = -1
Private WithEvents mList As MSForms.ListBox
Attribute mList.VB_VarHelpID = -1
Private mDate As MSForms.TextBox, mAmount As MSForms.TextBox, mComment As MSForms.TextBox
Private mReferenceDate As MSForms.TextBox, mReferenceAmount As MSForms.TextBox, mReferenceComment As MSForms.TextBox
Private mReferenceWBS As MSForms.TextBox
Private WithEvents mDefault As MSForms.CommandButton
Attribute mDefault.VB_VarHelpID = -1
Private WithEvents mCancel As MSForms.CommandButton
Attribute mCancel.VB_VarHelpID = -1

Public Sub Prepare(Optional ByVal selectedReference As String = "")
    Dim d As Object, a() As Variant, i As Long
    Finance_RequireIndividual
    mScope = Finance_ScopeId(): mRevision = Finance_Revision()
    mProjectName = CStr(Finance_DomainIndex(Finance_DomainSnapshot()("PROJECTS"), "PROJECT_ID")(CStr(mScope))("NAME"))
    Set mRows = Finance_DraftRows(): Set mExpenses = Finance_DraftExpenses()
    Finance_Validate mRows, mExpenses
    mIndex = 0: mFiltering = True
    FinanceUI_Init Me, Finance_Text("EXPENSES") & " | " & mProjectName, 830, 565
    FinanceUI_Label Me, "wbsHeader", "WBS", 18, 10, 210
    FinanceUI_Label Me, "dateHeader", Finance_Text("DATE_LABEL"), 240, 10, 120
    FinanceUI_Label Me, "amountHeader", Finance_Text("AMOUNT_HINT") & " (" & Finance_CurrencySymbol() & ")", 375, 10, 120
    FinanceUI_Label Me, "commentHeader", Finance_Text("COMMENT_HINT"), 510, 10, 290
    Set mFilterWBS = FinanceUI_Control(Me, "ComboBox", "filterWBS", 18, 36, 210, 25)
    Set mFilterDate = FinanceUI_Control(Me, "ComboBox", "filterDate", 240, 36, 120, 25)
    Set mFilterAmount = FinanceUI_Control(Me, "ComboBox", "filterAmount", 375, 36, 120, 25)
    Set mFilterComment = FinanceUI_Control(Me, "ComboBox", "filterComment", 510, 36, 290, 25)
    Set mList = FinanceUI_Control(Me, "ListBox", "expenses", 18, 76, 782, 230)
    mList.columnCount = 5: mList.ColumnWidths = "0 pt;222 pt;135 pt;135 pt;270 pt": mList.MultiSelect = fmMultiSelectSingle
    FinanceUI_Label Me, "wbsLabel", "WBS", 18, 322, 210
    FinanceUI_Label Me, "dateLabel", Finance_Text("DATE_LABEL"), 240, 322, 120
    FinanceUI_Label Me, "amountLabel", Finance_Text("AMOUNT_HINT") & " (" & Finance_CurrencySymbol() & ")", 375, 322, 120
    FinanceUI_Label Me, "commentLabel", Finance_Text("COMMENT_HINT"), 510, 322, 290
    Set mReferenceWBS = FinanceUI_Control(Me, "TextBox", "referenceWBS", 18, 350, 210, 26)
    Set mReferenceDate = FinanceUI_Control(Me, "TextBox", "referenceDate", 240, 350, 120, 26)
    Set mReferenceAmount = FinanceUI_Control(Me, "TextBox", "referenceAmount", 375, 350, 120, 26)
    Set mReferenceComment = FinanceUI_Control(Me, "TextBox", "referenceComment", 510, 350, 290, 26)
    Set mActivity = FinanceUI_Control(Me, "ComboBox", "activity", 18, 385, 210, 27)
    mActivity.style = fmStyleDropDownList: mActivity.columnCount = 2: mActivity.ColumnWidths = "0 pt;425 pt": mActivity.ListWidth = 425: mActivity.BoundColumn = 1
    mActivity.MatchEntry = fmMatchEntryComplete
    Dim reference As Object, referenceName As Variant
    For Each referenceName In Array("referenceWBS", "referenceDate", "referenceAmount", "referenceComment")
        Set reference = Me.controls(CStr(referenceName))
        reference.Locked = True: reference.TabStop = False: reference.BackColor = RGB(231, 237, 241)
    Next referenceName
    Set mDate = FinanceUI_Control(Me, "TextBox", "date", 240, 385, 120, 27)
    Set mAmount = FinanceUI_Control(Me, "TextBox", "amount", 375, 385, 120, 27)
    Set mComment = FinanceUI_Control(Me, "TextBox", "comment", 510, 385, 290, 27)
    FinanceUI_Button Me, "new", Finance_Text("ADD"), 18, 430, 125
    FinanceUI_Button Me, "saveRow", Finance_Text("MODIFY"), 154, 430, 125
    FinanceUI_Button Me, "delete", Finance_Text("DELETE"), 290, 430, 125
    FinanceUI_Button Me, "bulk", Finance_Text("BULK"), 430, 430, 180
    FinanceUI_Button Me, "confirm", Finance_Text("CONFIRM"), 474, 487, 145
    FinanceUI_Button Me, "cancel", Finance_Text("CANCEL"), 635, 487, 165, True
    Set mPE = FinanceUI_Buttons(Me, Array("new", "saveRow", "delete", "bulk", "confirm", "cancel"), mDefault, mCancel, "confirm|new")
    mSelectedReference = selectedReference
    FilterActivities
    PopulateFilters
    Rebuild
    SelectionState
End Sub

Private Sub FilterActivities()
    Dim keys As New Collection, labels As New Collection, d As Object, query As String, a() As Variant, i As Long, selected As Long
    query = "": selected = -1
    For Each d In mRows
        If Len(query) = 0 Or InStr(1, LCase$(FinanceUI_Caption(d)), query, vbBinaryCompare) > 0 Then
            keys.Add Finance_Reference(d): labels.Add FinanceUI_Caption(d)
        End If
    Next d
    mFiltering = True: mActivity.Clear
    If keys.count > 0 Then
        ReDim a(0 To keys.count - 1, 0 To 1)
        For i = 1 To keys.count
            a(i - 1, 0) = keys(i): a(i - 1, 1) = labels(i)
            If CStr(keys(i)) = mSelectedReference Then selected = i - 1
        Next i
        mActivity.list = a
        If selected < 0 Then selected = 0
        mActivity.ListIndex = selected
    End If
    mFiltering = False
    mActivity_Change
    SelectionState
End Sub

Private Sub Rebuild()
    Dim n As Long, r As Long, a() As Variant, symbol As String, entry As PEFinanceRow
    If mList Is Nothing Then Exit Sub
    mList.Clear: mIndex = 0
    SelectionState
    symbol = Finance_CurrencySymbol()
    For Each entry In mExpenses
        If MatchesFilters(entry, symbol) Then n = n + 1
    Next entry
    If n = 0 Then Exit Sub
    ReDim a(0 To n - 1, 0 To 4): n = 0
    For Each entry In mExpenses
        r = r + 1
        If MatchesFilters(entry, symbol) Then
            a(n, 0) = r: a(n, 1) = CStr(entry("WBS"))
            If entry("STATE") = "HISTORICAL" Then a(n, 1) = a(n, 1) & " [" & Finance_Text("HISTORICAL") & "]"
            a(n, 2) = Format$(CDate(entry("DATE")), "yyyy-mm-dd")
            a(n, 3) = Finance_MoneyCaption(entry("AMOUNT"), symbol): a(n, 4) = entry("COMMENT"): n = n + 1
        End If
    Next entry
    mList.list = a
End Sub

Private Function ContainsFilter(ByVal value As String, ByVal query As String, Optional ByVal money As Boolean = False) As Boolean
    If money Then
        value = Replace$(Replace$(Replace$(value, ChrW$(160), ""), " ", ""), ",", ".")
        query = Replace$(Replace$(Replace$(query, ChrW$(160), ""), " ", ""), ",", ".")
    End If
    ContainsFilter = (Len(Trim$(query)) = 0 Or InStr(1, value, Trim$(query), vbTextCompare) > 0)
End Function

Private Function MatchesFilters(ByVal entry As PEFinanceRow, ByVal symbol As String) As Boolean
    If Not ContainsFilter(CStr(entry("WBS")), mFilterWBS.text) Then Exit Function
    If Not ContainsFilter(Format$(CDate(entry("DATE")), "yyyy-mm-dd") & " " & Format$(CDate(entry("DATE")), "dd/mm/yyyy"), mFilterDate.text) Then Exit Function
    If Not ContainsFilter(Finance_MoneyCaption(entry("AMOUNT"), symbol), mFilterAmount.text, True) Then Exit Function
    If Not ContainsFilter(CStr(entry("COMMENT")), mFilterComment.text) Then Exit Function
    MatchesFilters = True
End Function

Private Sub PopulateFilters()
    Dim options(0 To 3) As Object, entry As PEFinanceRow, values As Variant, i As Long, key As Variant, name As Variant, control As MSForms.ComboBox, text As String, symbol As String
    mFiltering = True
    symbol = Finance_CurrencySymbol()
    For i = 0 To 3: Set options(i) = CreateObject("Scripting.Dictionary"): options(i).CompareMode = vbTextCompare: Next i
    For Each entry In mExpenses
        values = Array(CStr(entry("WBS")), Format$(CDate(entry("DATE")), "yyyy-mm-dd"), Finance_MoneyCaption(entry("AMOUNT"), symbol), CStr(entry("COMMENT")))
        For i = 0 To 3
            If options(i).count < 100 And Len(CStr(values(i))) > 0 Then options(i)(CStr(values(i))) = True
        Next i
    Next entry
    i = 0
    For Each name In Array("filterWBS", "filterDate", "filterAmount", "filterComment")
        Set control = Me.controls(CStr(name)): text = control.text: control.Clear
        control.style = fmStyleDropDownCombo: control.MatchEntry = fmMatchEntryNone
        For Each key In options(i).keys: control.AddItem CStr(key): Next key
        control.text = text: i = i + 1
    Next name
    mFiltering = False
End Sub

Private Sub mFilterWBS_Change()
    If Not mFiltering Then Rebuild
End Sub
Private Sub mFilterDate_Change()
    If Not mFiltering Then Rebuild
End Sub
Private Sub mFilterAmount_Change()
    If Not mFiltering Then Rebuild
End Sub
Private Sub mFilterComment_Change()
    If Not mFiltering Then Rebuild
End Sub

Private Function SelectedFinancialId() As Long
    Dim row As Object, reference As String
    If mActivity.ListIndex < 0 Then Err.Raise 5, "Finance", Finance_Text("SELECT_ACTIVITY")
    reference = CStr(mActivity.list(mActivity.ListIndex, 0))
    For Each row In mRows
        If Finance_Reference(row) = reference Then SelectedFinancialId = CLng(row("FINANCIAL_ID")): Exit Function
    Next row
    Err.Raise 5, "Finance", Finance_Text("SOURCE_CHANGED")
End Function

Public Sub PEInvoke(ByVal name As String)
    Dim d As Object, v As Variant, parts As Variant, value As Date, i As Long, selected As Object, savedIndex As Long
    Dim pending As Collection
    On Error GoTo Failed
    Select Case name
        Case "cancel": PEConsole_WheelDetach Me: Me.Hide
        Case "new", "saveRow"
            If name = "saveRow" And mIndex = 0 Then Exit Sub
            If mActivity.ListIndex < 0 Then Err.Raise 5, , Finance_Text("SELECT_ACTIVITY")
            Set d = Finance_ExpenseDraftRow(mRows, CStr(mActivity.list(mActivity.ListIndex, 0)), mDate.text, mAmount.text, mComment.text)
            If name = "saveRow" Then
                If mExpenses(mIndex).exists("EXPENSE_ID") Then d.Add "EXPENSE_ID", mExpenses(mIndex)("EXPENSE_ID")
            End If
            Set pending = New Collection: pending.Add d
            If Not Finance_ConfirmExpenseDates(mRows, pending) Then Exit Sub
            If name = "new" Then
                mExpenses.Add d: savedIndex = mExpenses.count
            Else
                savedIndex = mIndex: mExpenses.Remove mIndex
                If savedIndex > mExpenses.count Then mExpenses.Add d Else mExpenses.Add d, before:=savedIndex
            End If
            Finance_Validate mRows, mExpenses: PopulateFilters: Rebuild
            For i = 0 To mList.ListCount - 1
                If CLng(mList.list(i, 0)) = savedIndex Then
                    mList.ListIndex = i: mList.selected(i) = True: mList_Click
                    Exit For
                End If
            Next i
        Case "bulk"
            frmFinanceBulk.Prepare mRows
            frmFinanceBulk.Show vbModal
            If frmFinanceBulk.Accepted Then
                For Each d In frmFinanceBulk.ResultRows: mExpenses.Add d: Next d
                Call PopulateFilters
                Call Rebuild
            End If
            Unload frmFinanceBulk
        Case "delete"
            Set selected = CreateObject("Scripting.Dictionary")
            For i = 0 To mList.ListCount - 1
                If mList.selected(i) Then selected(CStr(mList.list(i, 0))) = True
            Next i
            If selected.count = 0 Then Exit Sub
            If MsgBox(Finance_Text("DELETE_EXPENSE"), vbYesNo + vbDefaultButton2 + vbQuestion, Finance_Text("EXPENSES")) <> vbYes Then Exit Sub
            For i = mExpenses.count To 1 Step -1
                If selected.exists(CStr(i)) Then mExpenses.Remove i
            Next i
            Call PopulateFilters
            Call Rebuild
        Case "confirm"
            Finance_CheckRevision mScope, mRevision
            Finance_Commit mRows, mExpenses: PEConsole_WheelDetach Me: Me.Hide
    End Select
    Exit Sub
Failed: FinanceUI_Error
End Sub

Private Function FutureDateMessage(ByVal expense As Object) As String
    Dim row As Object, identity As String
    For Each row In mRows
        If Finance_Reference(row) = Finance_Reference(expense) Then
            identity = "ID: " & CStr(row("TASK_ID")) & " | WBS: " & CStr(row("WBS"))
            Exit For
        End If
    Next row
    FutureDateMessage = mProjectName & vbCrLf & Replace(Finance_Text("FUTURE_EXPENSE"), "{IDs}", identity)
End Function
Private Sub mActivity_Change()
    If mFiltering Then Exit Sub
    If mActivity.ListIndex >= 0 Then mSelectedReference = CStr(mActivity.list(mActivity.ListIndex, 0))
    SelectionState
End Sub
Private Sub mList_Click()
    If mList.ListIndex < 0 Then mIndex = 0: SelectionState: Exit Sub
    mIndex = CLng(mList.list(mList.ListIndex, 0))
    Dim i As Long, reference As String
    reference = Finance_Reference(mExpenses(mIndex)): mFiltering = True
    For i = 0 To mActivity.ListCount - 1
        If CStr(mActivity.list(i, 0)) = reference Then mActivity.ListIndex = i: Exit For
    Next i
    mFiltering = False
    mDate.text = Format$(CDate(mExpenses(mIndex)("DATE")), "yyyy-mm-dd")
    mAmount.text = CStr(mExpenses(mIndex)("AMOUNT")): mComment.text = CStr(mExpenses(mIndex)("COMMENT"))
    SelectionState
End Sub
Private Sub SelectionState()
    If mReferenceDate Is Nothing Then Exit Sub
    If Not mPE Is Nothing Then mPE.BeginUpdate
    mReferenceWBS.text = "": mReferenceDate.text = "": mReferenceAmount.text = "": mReferenceComment.text = ""
    If mIndex > 0 Then
        mReferenceWBS.text = CStr(mExpenses(mIndex)("WBS"))
        mReferenceDate.text = Format$(CDate(mExpenses(mIndex)("DATE")), "yyyy-mm-dd")
        mReferenceAmount.text = Finance_MoneyCaption(mExpenses(mIndex)("AMOUNT"))
        mReferenceComment.text = CStr(mExpenses(mIndex)("COMMENT"))
    End If
    Me.controls("saveRow").enabled = (mIndex > 0): Me.controls("delete").enabled = (mIndex > 0)
    Me.controls("new").enabled = (mActivity.ListIndex >= 0)
    If Not mPE Is Nothing Then mPE.Sync
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
Private Sub peFocus_new_Enter()
    If Not mPE Is Nothing Then mPE.FocusChanged "new", True
End Sub
Private Sub peFocus_new_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Not mPE Is Nothing Then mPE.FocusChanged "new", False
End Sub
Private Sub peFocus_saveRow_Enter()
    If Not mPE Is Nothing Then mPE.FocusChanged "saveRow", True
End Sub
Private Sub peFocus_saveRow_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Not mPE Is Nothing Then mPE.FocusChanged "saveRow", False
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
    PEConsole_WheelAttach Me
End Sub
Private Sub UserForm_Deactivate()
    PEConsole_WheelDetach Me
End Sub


