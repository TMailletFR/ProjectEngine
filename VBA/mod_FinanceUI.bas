Attribute VB_Name = "mod_FinanceUI"
Option Explicit


Public Sub FinanceUI_Init(ByVal view As Object, ByVal title As String, ByVal width As Single, ByVal height As Single)
    view.caption = title: view.width = width: view.height = height
    view.BackColor = RGB(247, 248, 250): view.font.name = "Segoe UI": view.font.size = 11
End Sub

Public Function FinanceUI_Control(ByVal view As Object, ByVal kind As String, ByVal name As String, ByVal X As Single, ByVal Y As Single, ByVal width As Single, ByVal height As Single) As Object
    Dim c As Object
    On Error Resume Next: Set c = view.controls(name): On Error GoTo 0
    If c Is Nothing Then Set c = view.controls.Add("Forms." & kind & ".1", name, True)
    c.Move X, Y, width, height: c.font.name = "Segoe UI": c.font.size = 11
    Set FinanceUI_Control = c
End Function

Public Sub FinanceUI_Label(ByVal view As Object, ByVal name As String, ByVal text As String, ByVal X As Single, ByVal Y As Single, ByVal width As Single, Optional ByVal height As Single = 20)
    Dim c As Object
    Set c = FinanceUI_Control(view, "Label", name, X, Y, width, height)
    c.caption = text: c.BackStyle = fmBackStyleTransparent: c.WordWrap = True
End Sub

Public Sub FinanceUI_Button(ByVal view As Object, ByVal name As String, ByVal text As String, ByVal X As Single, ByVal Y As Single, ByVal width As Single, Optional ByVal Cancel As Boolean = False)
    Dim c As Object, relay As Object
    Set c = FinanceUI_Control(view, "CommandButton", name, X, Y, width, 30)
    c.caption = text: c.Cancel = Cancel: c.Default = (name = "confirm" Or name = "save")
    Set relay = FinanceUI_Control(view, "TextBox", "peFocus_" & name, 0, 0, 1, 1)
    relay.TabStop = False: relay.Locked = True: relay.BorderStyle = 0
End Sub

Public Function FinanceUI_Buttons(ByVal view As Object, ByVal names As Variant, ByRef defaultRelay As MSForms.CommandButton, ByRef cancelRelay As MSForms.CommandButton, Optional ByVal primaryNames As String = "confirm|save") As PEConsolePresentation
    Dim pe As New PEConsolePresentation
    Set defaultRelay = FinanceUI_Control(view, "CommandButton", "peDefault", 0, 0, 1, 1)
    Set cancelRelay = FinanceUI_Control(view, "CommandButton", "peCancel", 0, 0, 1, 1)
    defaultRelay.visible = False: cancelRelay.visible = False
    pe.Attach view, names, primaryNames
    pe.Sync
    Set FinanceUI_Buttons = pe
End Function

Public Function FinanceUI_ParseGridText(ByVal text As String) As Variant
    Dim rows As New Collection, fields(0 To 3) As String, field As String, column As Long, i As Long
    Dim quoted As Boolean, closed As Boolean, endedRow As Boolean, character As String, row As Variant, output() As Variant, r As Long, c As Long
    If Len(text) = 0 Then Exit Function
    i = 1
    Do While i <= Len(text)
        character = Mid$(text, i, 1): endedRow = False
        If quoted Then
            If character = """" Then
                If Mid$(text, i + 1, 1) = """" Then
                    field = field & """": i = i + 1
                Else
                    quoted = False: closed = True
                End If
            Else
                field = field & character
            End If
        ElseIf character = vbTab Or character = vbCr Or character = vbLf Then
            If column > 3 Then GoTo invalid
            fields(column) = field: field = "": closed = False: column = column + 1
            If character <> vbTab Then
                If column <> 4 Then GoTo invalid
                rows.Add fields
                For c = 0 To 3: fields(c) = "": Next c
                column = 0: endedRow = True
                If character = vbCr And Mid$(text, i + 1, 1) = vbLf Then i = i + 1
            End If
        ElseIf closed Then
            GoTo invalid
        ElseIf character = """" And Len(field) = 0 Then
            quoted = True
        Else
            field = field & character
        End If
        i = i + 1
    Loop
    If quoted Then GoTo invalid
    If Not endedRow Then
        If column <> 3 Then GoTo invalid
        fields(column) = field: rows.Add fields
    End If
    If rows.count = 0 Then Exit Function
    ReDim output(0 To rows.count - 1, 0 To 3)
    For Each row In rows
        For c = 0 To 3: output(r, c) = row(c): Next c
        r = r + 1
    Next row
    FinanceUI_ParseGridText = output: Exit Function
invalid:
    Err.Raise 5, "Finance", Finance_Text("BULK_FOUR_COLUMNS")
End Function

Public Sub FinanceUI_List(ByVal list As MSForms.ListBox, ByVal keys As Collection, ByVal labels As Collection)
    Dim a() As Variant, i As Long
    list.Clear: list.columnCount = 2: list.ColumnWidths = "0 pt;900 pt"
    list.MultiSelect = fmMultiSelectExtended
    If keys.count = 0 Then Exit Sub
    ReDim a(0 To keys.count - 1, 0 To 1)
    For i = 1 To keys.count: a(i - 1, 0) = keys(i): a(i - 1, 1) = labels(i): Next i
    list.list = a
End Sub

Public Function FinanceUI_Caption(ByVal d As Object) As String
    FinanceUI_Caption = CStr(d("WBS")) & " " & CStr(d("TASK_NAME"))
    If d("STATE") = "HISTORICAL" Then FinanceUI_Caption = FinanceUI_Caption & " [" & Finance_Text("HISTORICAL") & "]"
End Function

Public Sub FinanceUI_Error()
    MsgBox Err.description, vbExclamation, "ProjectEngine Finance"
End Sub


