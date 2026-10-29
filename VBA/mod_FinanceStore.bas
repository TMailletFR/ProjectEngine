Attribute VB_Name = "mod_FinanceStore"
Option Explicit

' Technical storage adapter only. Finance owns validation, policies and transactions.
Public Const FIN_STORE_SHEET As String = "FINANCE_DATA"

Public Function FinanceStore_Kinds() As Variant
    FinanceStore_Kinds = Array("PROJECTS", "LINKS", "RECORDS", "JOURNAL", "META")
End Function

Public Function FinanceStore_Keys(ByVal kind As String) As Variant
    Select Case kind
        Case "PROJECTS": FinanceStore_Keys = Array("PROJECT_ID", "NAME", "DISPLAY_ORDER")
        Case "LINKS": FinanceStore_Keys = Array("LINK_ID", "PROJECT_ID", "WBS", "STATE", "FINANCIAL_ID", "ROW_ORDER", "TASK_ID", "TASK_NAME", "PROJECT", "START", "FINISH", "COPY_FROM")
        Case "RECORDS": FinanceStore_Keys = Array("FINANCIAL_ID", "BUDGET", "PROFILE", "ETC")
        Case "JOURNAL": FinanceStore_Keys = Array("EXPENSE_ID", "FINANCIAL_ID", "DATE", "AMOUNT", "COMMENT")
        Case "META": FinanceStore_Keys = Array("SCOPE_ID", "CURRENCY", "REFRESH_AT", "REVISION", "NEXT_PROJECT", "NEXT_LINK", "NEXT_FINANCIAL", "NEXT_EXPENSE")
        Case Else: Err.Raise 5, "FinanceStore", "Unknown storage table: " & kind
    End Select
End Function

Public Function FinanceStore_Row(ByVal kind As String) As Object
    Dim row As New PEFinanceRow
    row.Initialize FinanceStore_Keys(kind)
    Set FinanceStore_Row = row
End Function

Public Function FinanceStore_Exists(ByVal wb As Workbook) As Boolean
    Dim ws As Worksheet, n As Long, description As String
    On Error GoTo missing
    Set ws = wb.Worksheets(FIN_STORE_SHEET)
    FinanceStore_Exists = True
    Exit Function
missing:
    n = Err.number: description = Err.description
    If n <> 9 Then Err.Raise n, "FinanceStore", description
End Function

Public Function FinanceStore_Empty() As Object
    Dim graph As Object, kind As Variant, rows As Collection
    Set graph = CreateObject("Scripting.Dictionary")
    For Each kind In FinanceStore_Kinds()
        Set rows = New Collection
        graph.Add CStr(kind), rows
    Next kind
    Set FinanceStore_Empty = graph
End Function

Public Function FinanceStore_Read(ByVal wb As Workbook) As Object
    Dim graph As Object, ws As Worksheet, tbl As ListObject, kind As Variant
    Dim keys As Variant, cols As Object, data As Variant, headers As Variant
    Dim rows As Collection, row As PEFinanceRow, r As Long, c As Long
    Set graph = FinanceStore_Empty()
    Set ws = wb.Worksheets(FIN_STORE_SHEET)
    For Each kind In FinanceStore_Kinds()
        Set tbl = ws.ListObjects("tbl_FINANCE_" & CStr(kind))
        keys = FinanceStore_Keys(CStr(kind))
        If tbl.ListColumns.count <> UBound(keys) + 1 Then Err.Raise 5, "FinanceStore", "Invalid storage columns: " & kind
        Set cols = CreateObject("Scripting.Dictionary")
        headers = tbl.HeaderRowRange.Value2
        For c = 1 To UBound(headers, 2): cols.Add CStr(headers(1, c)), c: Next c
        For c = 0 To UBound(keys)
            If Not cols.exists(CStr(keys(c))) Then Err.Raise 5, "FinanceStore", "Missing storage column: " & keys(c)
        Next c
        Set rows = graph(CStr(kind))
        If Not tbl.DataBodyRange Is Nothing And tbl.ListRows.count > 0 Then
            data = tbl.DataBodyRange.Value2
            For r = 1 To UBound(data, 1)
                Set row = FinanceStore_Row(CStr(kind))
                For c = 0 To UBound(keys): row(CStr(keys(c))) = data(r, CLng(cols(CStr(keys(c))))): Next c
                rows.Add row
            Next r
        End If
    Next kind
    Set FinanceStore_Read = graph
End Function

Public Sub FinanceStore_Write(ByVal wb As Workbook, ByVal graph As Object, Optional ByVal onlyKind As String = "", Optional ByVal before As Object)
    Dim ws As Worksheet, tbl As ListObject, kind As Variant, keys As Variant
    Dim anchors As Variant, headers() As Variant, data() As Variant, previous As Range
    Dim rows As Collection, row As PEFinanceRow, i As Long, r As Long, c As Long, value As Variant, oldCount As Long
    Dim cols As Object, headerData As Variant
    Dim perf As clsPerfScope
    Set perf = Profiler_BeginScope("Finance.Storage.Write." & onlyKind, "Finance")
    If Not FinanceStore_Exists(wb) Then
        Set ws = wb.Worksheets.Add(after:=wb.Worksheets(wb.Worksheets.count))
        ws.name = FIN_STORE_SHEET
    Else
        Set ws = wb.Worksheets(FIN_STORE_SHEET)
    End If
    anchors = Array("A1", "F1", "U1", "Z1", "AF1")
    For Each kind In FinanceStore_Kinds()
        If onlyKind = "DATA" And CStr(kind) = "META" Then GoTo NextKind
        If onlyKind <> "" And onlyKind <> "DATA" And onlyKind <> CStr(kind) Then GoTo NextKind
        If Not before Is Nothing Then
            If FinanceStore_Token(graph, CStr(kind)) = FinanceStore_Token(before, CStr(kind)) Then GoTo NextKind
        End If
        keys = FinanceStore_Keys(CStr(kind)): Set tbl = Nothing
        For Each tbl In ws.ListObjects
            If tbl.name = "tbl_FINANCE_" & CStr(kind) Then Exit For
        Next tbl
        If Not tbl Is Nothing Then
            If tbl.name <> "tbl_FINANCE_" & CStr(kind) Then Set tbl = Nothing
        End If
        If tbl Is Nothing Then
            ReDim headers(1 To 1, 1 To UBound(keys) + 1)
            For c = 0 To UBound(keys): headers(1, c + 1) = keys(c): Next c
            ws.Range(anchors(i)).Resize(1, UBound(keys) + 1).Value2 = headers
            Set tbl = ws.ListObjects.Add(xlSrcRange, ws.Range(anchors(i)).Resize(1, UBound(keys) + 1), , xlYes)
            tbl.name = "tbl_FINANCE_" & CStr(kind)
        End If
        headerData = tbl.HeaderRowRange.Value2: Set cols = CreateObject("Scripting.Dictionary")
        For c = 1 To UBound(headerData, 2): cols.Add CStr(headerData(1, c)), c: Next c
        For c = 0 To UBound(keys)
            If Not cols.exists(CStr(keys(c))) Then Err.Raise 5, "FinanceStore", "Missing storage column: " & keys(c)
        Next c
        Set rows = graph(CStr(kind)): Set previous = tbl.DataBodyRange
        oldCount = tbl.ListRows.count
        If oldCount > rows.count And Not previous Is Nothing Then previous.rows(rows.count + 1).Resize(oldCount - rows.count).ClearContents
        If oldCount <> rows.count Then tbl.Resize tbl.HeaderRowRange.Resize(IIf(rows.count = 0, 2, rows.count + 1), UBound(keys) + 1)
        If rows.count > 0 Then
            ReDim data(1 To rows.count, 1 To UBound(keys) + 1): r = 0
            For Each row In rows
                r = r + 1
                For c = 0 To UBound(keys)
                    value = row(CStr(keys(c)))
                    If VarType(value) = vbString Then value = "'" & CStr(value)
                    data(r, CLng(cols(CStr(keys(c))))) = value
                Next c
            Next row
            tbl.HeaderRowRange.offset(1).Resize(rows.count, UBound(keys) + 1).formula = data
            If tbl.ListRows.count <> rows.count Then tbl.Resize tbl.HeaderRowRange.Resize(rows.count + 1, UBound(keys) + 1)
            If tbl.DataBodyRange Is Nothing Or tbl.ListRows.count <> rows.count Then Err.Raise 5, "FinanceStore", "Storage materialization failed"
        Else
            If tbl.ListRows.count > 0 Then tbl.ListRows(1).Delete
            If tbl.ListRows.count <> 0 Then Err.Raise 5, "FinanceStore", "Storage empty-table materialization failed"
        End If
NextKind:
        i = i + 1
    Next kind
    If ws.visible <> xlSheetVeryHidden Then ws.visible = xlSheetVeryHidden
End Sub

Public Function FinanceStore_Clone(ByVal graph As Object) As Object
    Dim result As Object, kind As Variant, keys As Variant, rows As Collection, row As PEFinanceRow, copy As PEFinanceRow, key As Variant
    Set result = FinanceStore_Empty()
    For Each kind In FinanceStore_Kinds()
        Set rows = result(CStr(kind)): keys = FinanceStore_Keys(CStr(kind))
        For Each row In graph(CStr(kind))
            Set copy = FinanceStore_Row(CStr(kind))
            For Each key In keys: copy(CStr(key)) = row(CStr(key)): Next key
            rows.Add copy
        Next row
    Next kind
    Set FinanceStore_Clone = result
End Function

Public Function FinanceStore_Token(ByVal graph As Object, Optional ByVal onlyKind As String = "") As String
    Dim kind As Variant, row As PEFinanceRow, key As Variant, text As String, keys As Variant
    Dim kinds(0 To 4) As String, values() As String, r As Long, i As Long
    For Each kind In FinanceStore_Kinds()
        If onlyKind <> "" And CStr(kind) <> onlyKind Then GoTo NextKind
        ReDim values(0 To graph(CStr(kind)).count): r = 0
        values(0) = CStr(kind) & ":" & graph(CStr(kind)).count & ":"
        keys = FinanceStore_Keys(CStr(kind))
        For Each row In graph(CStr(kind))
            r = r + 1
            For Each key In keys
                text = CStr(row(CStr(key)))
                values(r) = values(r) & Len(text) & ":" & text
            Next key
        Next row
        kinds(i) = Join(values, ""): i = i + 1
NextKind:
    Next kind
    FinanceStore_Token = Join(kinds, "")
End Function

