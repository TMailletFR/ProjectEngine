VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmFinanceBulk 
   Caption         =   "UserForm1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "frmFinanceBulk.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmFinanceBulk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Accepted As Boolean
Public ResultRows As Collection
Private mRows As Collection, mData() As Variant, mCount As Long, mRow As Long, mColumn As Long, mClickX As Single, mChanging As Boolean
Private mPE As PEConsolePresentation
Private WithEvents mGrid As MSForms.ListBox
Attribute mGrid.VB_VarHelpID = -1
Private WithEvents mEditor As MSForms.TextBox
Attribute mEditor.VB_VarHelpID = -1
Private WithEvents mWBS As MSForms.ComboBox
Attribute mWBS.VB_VarHelpID = -1
Private WithEvents mDefault As MSForms.CommandButton
Attribute mDefault.VB_VarHelpID = -1
Private WithEvents mCancel As MSForms.CommandButton
Attribute mCancel.VB_VarHelpID = -1

Public Sub Prepare(ByVal rows As Collection)
    Dim row As PEFinanceRow, options() As Variant, i As Long
    Accepted = False: Set ResultRows = New Collection: Set mRows = rows
    mRow = 0: mColumn = 0: mCount = 1: ReDim mData(0 To 3, 0 To 0)
    FinanceUI_Init Me, Finance_Text("BULK_TITLE"), 830, 535
    FinanceUI_Label Me, "wbsHeader", "WBS", 18, 14, 180
    FinanceUI_Label Me, "dateHeader", Finance_Text("DATE_LABEL"), 198, 14, 135
    FinanceUI_Label Me, "amountHeader", Finance_Text("AMOUNT_HINT"), 333, 14, 120
    FinanceUI_Label Me, "commentHeader", Finance_Text("COMMENT_HINT"), 453, 14, 347
    Set mGrid = FinanceUI_Control(Me, "ListBox", "grid", 18, 44, 782, 300)
    mGrid.columnCount = 4: mGrid.ColumnWidths = "180 pt;135 pt;120 pt;327 pt": mGrid.MultiSelect = fmMultiSelectSingle
    Set mEditor = FinanceUI_Control(Me, "TextBox", "cellEditor", 18, 44, 180, 18)
    mEditor.MaxLength = 32767: mEditor.visible = False
    Set mWBS = FinanceUI_Control(Me, "ComboBox", "wbsEditor", 18, 44, 180, 18)
    mWBS.style = fmStyleDropDownCombo: mWBS.MatchEntry = fmMatchEntryNone
    mWBS.columnCount = 2: mWBS.ColumnWidths = "0 pt;400 pt": mWBS.ListWidth = 425
    If rows.count > 0 Then
        ReDim options(0 To rows.count - 1, 0 To 1)
        For Each row In rows
            options(i, 0) = Finance_Reference(row): options(i, 1) = FinanceUI_Caption(row): i = i + 1
        Next row
        mChanging = True: mWBS.list = options: mChanging = False
    End If
    mWBS.visible = False
    FinanceUI_Label Me, "help", Finance_Text("BULK_HELP"), 18, 362, 782, 65
    FinanceUI_Button Me, "paste", Finance_Text("BULK_PASTE"), 18, 452, 145
    FinanceUI_Button Me, "confirm", Finance_Text("CONFIRM"), 474, 452, 145
    FinanceUI_Button Me, "cancel", Finance_Text("CANCEL"), 635, 452, 165, True
    Set mPE = FinanceUI_Buttons(Me, Array("paste", "confirm", "cancel"), mDefault, mCancel)
    RefreshGrid
End Sub

Public Property Get rowCount() As Long
    rowCount = mCount
End Property

Public Function GridData() As Variant
    Dim result() As Variant, r As Long, c As Long
    ReDim result(0 To mCount - 1, 0 To 3)
    For r = 0 To mCount - 1
        For c = 0 To 3: result(r, c) = mData(c, r): Next c
    Next r
    GridData = result
End Function

Private Sub RefreshGrid()
    Dim top As Long, wasChanging As Boolean
    wasChanging = mChanging: mChanging = True
    top = mGrid.TopIndex
    mGrid.list = GridData()
    If mRow >= 0 And mRow < mCount Then mGrid.ListIndex = mRow
    If top >= 0 And top < mCount Then mGrid.TopIndex = top
    mChanging = wasChanging
End Sub

Public Sub SetCell(ByVal row As Long, ByVal column As Long, ByVal value As String)
    Dim c As Long, nonempty As Boolean
    If row < 0 Or row >= mCount Or column < 0 Or column > 3 Then Err.Raise 5, "Finance", "Grid cell out of bounds"
    mData(column, row) = value: mGrid.list(row, column) = value
    If row = mCount - 1 Then
        For c = 0 To 3
            If Len(CStr(mData(c, row))) > 0 Then nonempty = True
        Next c
        If nonempty Then
            mCount = mCount + 1: ReDim Preserve mData(0 To 3, 0 To mCount - 1)
            mGrid.AddItem ""
        End If
    End If
End Sub

Public Sub PasteText(ByVal text As String)
    Dim data As Variant, first As Long, last As Long, required As Long, r As Long, c As Long, target As Long
    data = FinanceUI_ParseGridText(text)
    If IsEmpty(data) Then Exit Sub
    first = LBound(data, 1): last = UBound(data, 1)
    If Finance_BulkExpenseHeader(data, first) Then first = first + 1
    If first > last Then Exit Sub
    HideEditors
    required = mRow + last - first + 2
    If required > mCount Then mCount = required: ReDim Preserve mData(0 To 3, 0 To mCount - 1)
    target = mRow
    For r = first To last
        For c = 0 To 3: mData(c, target) = CStr(data(r, c)): Next c
        target = target + 1
    Next r
    RefreshGrid
End Sub

Private Sub PasteClipboard()
    Dim clipboard As New MSForms.DataObject
    clipboard.GetFromClipboard
    PasteText clipboard.GetText(1)
End Sub

Private Sub HideEditors()
    mEditor.visible = False: mWBS.visible = False
End Sub

Private Sub OpenCell()
    Dim left As Single, top As Single, width As Single, i As Long
    If mRow < 0 Or mRow >= mCount Then Exit Sub
    HideEditors
    mChanging = True
    left = mGrid.left
    For i = 0 To mColumn - 1: left = left + Choose(i + 1, 180, 135, 120, 327): Next i
    width = Choose(mColumn + 1, 180, 135, 120, 327)
    top = mGrid.top + 2 + (mRow - mGrid.TopIndex) * (mGrid.font.size + 4)
    If top + 18 > mGrid.top + mGrid.height Then top = mGrid.top + mGrid.height - 18
    If mColumn = 0 Then
        mWBS.Move left, top, width, 18: mWBS.ListIndex = -1: mWBS.text = CStr(mData(0, mRow))
        mWBS.visible = True: mWBS.ZOrder 0: mWBS.SetFocus
    Else
        mEditor.Move left, top, width, 18: mEditor.text = CStr(mData(mColumn, mRow))
        mEditor.visible = True: mEditor.ZOrder 0: mEditor.SetFocus: mEditor.SelStart = Len(mEditor.text)
    End If
    mChanging = False
End Sub

Private Sub mGrid_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    mClickX = X
End Sub
Private Sub mGrid_Click()
    If mChanging Then Exit Sub
    If mGrid.ListIndex < 0 Then Exit Sub
    mRow = mGrid.ListIndex
    Select Case mClickX
        Case Is < 180: mColumn = 0
        Case Is < 315: mColumn = 1
        Case Is < 435: mColumn = 2
        Case Else: mColumn = 3
    End Select
    OpenCell
End Sub
Private Sub mEditor_Change()
    If Not mChanging Then SetCell mRow, mColumn, mEditor.text
End Sub
Private Sub mWBS_Change()
    If mChanging Then Exit Sub
    If mWBS.ListIndex >= 0 Then
        SetCell mRow, 0, CStr(mWBS.list(mWBS.ListIndex, 1))
    Else
        SetCell mRow, 0, mWBS.text
    End If
End Sub

Private Sub EditKey(ByRef KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Dim clipboard As MSForms.DataObject, text As String
    If KeyCode = 86 And (Shift And 2) <> 0 Then
        Set clipboard = New MSForms.DataObject: clipboard.GetFromClipboard: text = clipboard.GetText(1)
        If InStr(text, vbTab) > 0 Then KeyCode = 0: PasteText text
    ElseIf KeyCode = 9 Then
        KeyCode = 0
        If (Shift And 1) <> 0 Then
            mColumn = mColumn - 1
            If mColumn < 0 Then mColumn = 3: mRow = mRow - 1
        Else
            mColumn = mColumn + 1
            If mColumn > 3 Then mColumn = 0: mRow = mRow + 1
        End If
        If mRow < 0 Then mRow = 0
        If mRow >= mCount Then mRow = mCount - 1
        mGrid.ListIndex = mRow: OpenCell
    ElseIf KeyCode = 13 Then
        KeyCode = 0: mRow = mRow + 1
        If mRow >= mCount Then mRow = mCount - 1
        mGrid.ListIndex = mRow: OpenCell
    ElseIf KeyCode = 27 Then
        KeyCode = 0: HideEditors: mGrid.SetFocus
    End If
End Sub
Private Sub mEditor_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    On Error GoTo Failed
    EditKey KeyCode, Shift
    Exit Sub
Failed: FinanceUI_Error
End Sub
Private Sub mWBS_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    On Error GoTo Failed
    EditKey KeyCode, Shift
    Exit Sub
Failed: FinanceUI_Error
End Sub
Private Sub mGrid_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    On Error GoTo Failed
    If KeyCode = 86 And (Shift And 2) <> 0 Then
        KeyCode = 0: PasteClipboard
    ElseIf KeyCode = 13 Or KeyCode = 113 Then
        KeyCode = 0: mRow = mGrid.ListIndex: OpenCell
    End If
    Exit Sub
Failed: FinanceUI_Error
End Sub

Public Sub PEInvoke(ByVal name As String)
    On Error GoTo Failed
    Select Case name
        Case "paste": PasteClipboard
        Case "cancel": Accepted = False: PEConsole_WheelDetach Me: Me.Hide
        Case "confirm"
            HideEditors
            Set ResultRows = Finance_BulkExpenseRows(mRows, GridData())
            If ResultRows.count = 0 Then Err.Raise 5, "Finance", Finance_Text("BULK_EMPTY")
            If Not Finance_ConfirmExpenseDates(mRows, ResultRows) Then Exit Sub
            Accepted = True: PEConsole_WheelDetach Me: Me.Hide
    End Select
    Exit Sub
Failed: FinanceUI_Error
End Sub
Private Sub mDefault_Click()
    mPE.InvokeDefault
End Sub
Private Sub mCancel_Click()
    mPE.InvokeCancel
End Sub
Public Sub PEHover(ByVal name As String)
    If Not mPE Is Nothing Then mPE.Hover name
End Sub
Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    PEHover ""
End Sub
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    PEConsole_WheelDetach Me
    If CloseMode = 0 Then Cancel = True: Me.Hide
    If CloseMode <> 0 Then
        If Not mPE Is Nothing Then mPE.Detach
        Set mPE = Nothing
    End If
End Sub
Private Sub UserForm_Activate()
    PEConsole_WheelAttach Me, Me
End Sub
Public Sub PEListScrolled(ByVal name As String)
    If name = "grid" Then HideEditors
End Sub
Private Sub UserForm_Deactivate()
    PEConsole_WheelDetach Me
End Sub
Private Sub UserForm_Terminate()
    PEConsole_WheelDetach Me
    Set mPE = Nothing
End Sub

