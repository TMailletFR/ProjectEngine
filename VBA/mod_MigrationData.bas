Attribute VB_Name = "mod_MigrationData"
Option Explicit

Public Function MigrationData_WbsInputs() As Variant
    MigrationData_WbsInputs = Array("ID", "WBS", "TASK_NAME", "TASK_DESCRIPTION", "DISCIPLINE", "SUPPLIER", _
        "PROJECT", "CAL", "TASK_TYPE", "S", "COMMENTS", "PREDECESSORS_WBS", "WEIGHT_PERCENT", _
        "PROGRESS_PERCENT", "BASELINE_START", "BASELINE_DURATION", "ACTUAL_START", "ACTUAL_FINISH", "FORECAST_START", "FORECAST_FINISH")
End Function

Public Function MigrationData_ConstraintInputs() As Variant
    MigrationData_ConstraintInputs = Array("ID", "START_CONSTRAINT_TYPE", "START_CONSTRAINT_DATE", _
        "FINISH_CONSTRAINT_TYPE", "FINISH_CONSTRAINT_DATE", "ACTIVE", "COMMENT", "DEADLINE")
End Function

Public Function MigrationData_Table(ByVal wb As Workbook, ByVal name As String, Optional ByVal required As Boolean = True) As ListObject
    Dim ws As Worksheet, tbl As ListObject
    For Each ws In wb.Worksheets
        For Each tbl In ws.ListObjects
            If tbl.Name = name Then
                If Not MigrationData_Table Is Nothing Then Err.Raise 5, , "DUPLICATE_TABLE:" & name
                Set MigrationData_Table = tbl
            End If
        Next tbl
    Next ws
    If required And MigrationData_Table Is Nothing Then Migration_Reject "MIGRATION.PROFILE", name
End Function

Public Function MigrationData_ColumnMap(ByVal tbl As ListObject) As Object
    Dim result As Object, titles As Object, key As Variant, column As ListColumn
    Dim keys As Variant, language As Variant, title As String
    Set result = CreateObject("Scripting.Dictionary")
    Set titles = CreateObject("Scripting.Dictionary")
    titles.CompareMode = vbTextCompare
    If SchemaIsVisibleTable(tbl.Name) Then
        keys = SchemaColumnKeys(tbl.Name)
        For Each key In keys
            For Each language In Array("EN", "FR")
                title = SchemaColumnTitle(tbl.Name, CStr(key), CStr(language))
                If Not titles.Exists(title) Then titles.Add title, CStr(key)
            Next language
        Next key
        If tbl.Name = VTS_TABLE_WBS Then titles.Add "Package", VTS_COL_PROJECT
        For Each column In tbl.ListColumns
            If titles.Exists(column.Name) Then
                key = titles(column.Name)
                If result.Exists(key) Then Migration_Reject "MIGRATION.INVALID", "Package/Project or duplicated header: " & column.Name
                result.Add key, column.Index
            Else
                Migration_Reject "MIGRATION.PROFILE", tbl.Name & ": " & column.Name
            End If
        Next column
    Else
        For Each column In tbl.ListColumns
            result.Add column.Name, column.Index
        Next column
    End If
    Set MigrationData_ColumnMap = result
End Function

Public Function MigrationData_Read(ByVal source As Workbook, ByVal target As Workbook) As Object
    Dim snapshot As Object, tables As Object, block As Object, settings As Object
    Dim tableName As Variant, tbl As ListObject, targetTbl As ListObject, keys As Variant
    Dim version As Long, profile As String, phase As String
    On Error GoTo Failed
    phase = "SCHEMA"
    version = WorkbookSchema_Version(source)
    If version > CURRENT_SCHEMA_VERSION Then Migration_Reject "MIGRATION.FUTURE"
    If version <> 0 And version <> CURRENT_SCHEMA_VERSION Then Migration_Reject "MIGRATION.PROFILE"
    If version = CURRENT_SCHEMA_VERSION Then
        Dim importState As String
        importState = CStr(WorkbookSchema_Property(source, "ProjectEngine.ImportState"))
        If importState = "FAILED" Or importState = "IMPORT_IN_PROGRESS" Then
            If Not (source Is target And Migration_IsRunning()) Then Migration_Reject "MIGRATION.PROFILE", "Incomplete source import"
        End If
    End If
    If source.Date1904 <> target.Date1904 Then Migration_Reject "MIGRATION.UNSAFE"
    Set snapshot = CreateObject("Scripting.Dictionary")
    Set tables = CreateObject("Scripting.Dictionary")
    phase = "SETTINGS"
    Set settings = Settings_MigrationSnapshot(source)
    phase = "tbl_WBS"
    Set tbl = MigrationData_Table(source, "tbl_WBS")
    Set block = MigrationData_ReadBlock(tbl, MigrationData_WbsInputs(), True)
    profile = CStr(settings("Storage"))
    Dim map As Object
    Set map = MigrationData_ColumnMap(tbl)
    If Not map.Exists("S") Then profile = profile & "/S_DEFAULT"
    If Not map.Exists("PROJECT") Then Migration_Reject "MIGRATION.PROFILE", "PROJECT"
    profile = profile & "/INPUTS_SCHEMA1"
    If version = 1 Then profile = "SCHEMA1/" & profile
    tables.Add "tbl_WBS", block
    phase = "tbl_CONSTRAINTS"
    Set tbl = MigrationData_Table(source, "tbl_CONSTRAINTS")
    Set block = MigrationData_ReadBlock(tbl, MigrationData_ConstraintInputs())
    tables.Add "tbl_CONSTRAINTS", block
    phase = "tbl_EVENT_ACK"
    Set tbl = MigrationData_Table(source, "tbl_EVENT_ACK")
    Set block = MigrationData_ReadBlock(tbl, SchemaColumnKeys("tbl_EVENT_ACK"))
    tables.Add "tbl_EVENT_ACK", block
    For Each tableName In Array("tbl_DASHBOARD_SNAPSHOTS", "tbl_CALC_ALARM")
        phase = CStr(tableName)
        Set targetTbl = MigrationData_Table(target, CStr(tableName))
        keys = MigrationData_PhysicalKeys(targetTbl)
        Set tbl = MigrationData_Table(source, CStr(tableName), False)
        If tbl Is Nothing Then
            If version = CURRENT_SCHEMA_VERSION And tableName = "tbl_DASHBOARD_SNAPSHOTS" Then Migration_Reject "MIGRATION.PROFILE", CStr(tableName)
            Set block = CreateObject("Scripting.Dictionary")
            block.Add "Keys", keys: block.Add "Values", Empty: block.Add "Count", 0&
            block.Add "Available", False
        ElseIf tableName = "tbl_CALC_ALARM" Then
            Set block = MigrationData_ReadOptionalHistory(tbl, keys)
        Else
            Set block = MigrationData_ReadBlock(tbl, keys)
            block.Add "Available", True
        End If
        tables.Add CStr(tableName), block
    Next tableName
    snapshot.Add "Tables", tables
    snapshot.Add "Settings", settings
    snapshot.Add "Profile", profile
    snapshot.Add "Source", source.FullName
    phase = "VALIDATION"
    MigrationData_Validate snapshot
    phase = "DIGEST"
    snapshot.Add "Digest", MigrationData_Digest(snapshot)
    Set MigrationData_Read = snapshot
    Exit Function
Failed:
    Dim errorNumber As Long, errorSource As String, errorDescription As String
    errorNumber = Err.Number: errorSource = Err.Source: errorDescription = Err.Description
    If errorSource = "MigrationValidation" Then Err.Raise errorNumber, errorSource, errorDescription
    If errorSource = "NormalizeTaskTypeValue" Or errorSource = "NormalizeSummaryDisplayValue" Then Migration_Reject "MIGRATION.INVALID", errorDescription
    Err.Raise errorNumber, "MigrationData_Read/" & phase, errorDescription
End Function

Private Function MigrationData_ReadOptionalHistory(ByVal tbl As ListObject, ByVal keys As Variant) As Object
    Dim block As Object, diagnostic As String, number As Long, origin As String
    On Error GoTo Incompatible
    Set block = MigrationData_ReadBlock(tbl, keys)
    MigrationData_ValidateLedger block, "tbl_CALC_ALARM"
    block.Add "Available", True
    Set MigrationData_ReadOptionalHistory = block
    Exit Function
Incompatible:
    number = Err.Number: origin = Err.Source: diagnostic = Err.Description
    If origin <> "MigrationValidation" Then Err.Raise number, origin, diagnostic
    Set block = CreateObject("Scripting.Dictionary")
    block.Add "Keys", keys: block.Add "Values", Empty: block.Add "Count", 0&
    block.Add "Available", False
    block.Add "UnavailableDiagnostic", diagnostic
    Set MigrationData_ReadOptionalHistory = block
End Function

Private Function MigrationData_PhysicalKeys(ByVal tbl As ListObject) As Variant
    Dim keys() As Variant, i As Long
    ReDim keys(0 To tbl.ListColumns.Count - 1)
    For i = 1 To tbl.ListColumns.Count
        keys(i - 1) = tbl.ListColumns(i).Name
    Next i
    MigrationData_PhysicalKeys = keys
End Function

Private Function MigrationData_ReadBlock(ByVal tbl As ListObject, ByVal keys As Variant, Optional ByVal wbs As Boolean = False) As Object
    Dim block As Object, map As Object, data As Variant, formulas As Variant, output() As Variant
    Dim r As Long, c As Long, n As Long, count As Long, key As String, index As Long, value As Variant
    Dim summaries As Object, emptyWbsRow As Boolean, hasFormula As Variant, inputFormulas As Object, preservedFormulas As Object, formula As String
    Set block = CreateObject("Scripting.Dictionary")
    If wbs Then Set preservedFormulas = CreateObject("Scripting.Dictionary")
    Set map = MigrationData_ColumnMap(tbl)
    If Not SchemaIsVisibleTable(tbl.Name) Then
        If tbl.ListColumns.Count <> UBound(keys) - LBound(keys) + 1 Then Migration_Reject "MIGRATION.PROFILE", tbl.Name & " column count"
    End If
    For c = LBound(keys) To UBound(keys)
        key = CStr(keys(c))
        If Not map.Exists(key) Then
            If Not (wbs And key = "S") Then Migration_Reject "MIGRATION.PROFILE", tbl.Name & ": " & key
        End If
    Next c
    count = tbl.ListRows.Count
    If Not tbl.DataBodyRange Is Nothing And count > 0 Then
        data = tbl.DataBodyRange.Value2
        For c = LBound(keys) To UBound(keys)
            key = CStr(keys(c))
            If map.Exists(key) Then
                hasFormula = tbl.ListColumns(map(key)).DataBodyRange.HasFormula
                If Not wbs Then
                    If IsNull(hasFormula) Then Migration_Reject "MIGRATION.FORMULA", tbl.ListColumns(map(key)).Name
                    If CBool(hasFormula) Then Migration_Reject "MIGRATION.FORMULA", tbl.ListColumns(map(key)).Name
                End If
            End If
        Next c
        If wbs Then
            formulas = tbl.DataBodyRange.Formula
            Set summaries = DataSync_BuildSummaryWbsLookupFromArray(data, map)
            Set inputFormulas = MigrationData_PrepareInputFormulas(tbl, map, keys, count, formulas)
        End If
        ReDim output(1 To count, 1 To UBound(keys) - LBound(keys) + 1)
        For r = 1 To count
            emptyWbsRow = False
            If wbs Then
                If IsError(data(r, map("ID"))) Then Migration_Reject "MIGRATION.INVALID", "ID"
                emptyWbsRow = (Len(CStr(data(r, map("ID")))) = 0)
            Else
                emptyWbsRow = True
                For c = LBound(keys) To UBound(keys)
                    If IsError(data(r, map(keys(c)))) Then
                        emptyWbsRow = False: Exit For
                    ElseIf Len(CStr(data(r, map(keys(c))))) > 0 Then
                        emptyWbsRow = False: Exit For
                    End If
                Next c
            End If
            If emptyWbsRow Then
                If wbs Then
                    For c = LBound(keys) To UBound(keys)
                        key = CStr(keys(c))
                        If map.Exists(key) Then
                            If Len(CStr(data(r, map(key)))) > 0 Then Migration_Reject "MIGRATION.INVALID", "WBS row without ID"
                        End If
                    Next c
                End If
            Else
                n = n + 1
                For c = LBound(keys) To UBound(keys)
                    key = CStr(keys(c))
                    formula = vbNullString
                    If wbs Then
                        If inputFormulas.Exists(CStr(r) & ":" & CStr(c - LBound(keys) + 1)) Then
                            formula = inputFormulas(CStr(r) & ":" & CStr(c - LBound(keys) + 1))
                            preservedFormulas.Add CStr(n) & ":" & CStr(c - LBound(keys) + 1), formula
                        End If
                    End If
                    If map.Exists(key) Then
                        index = CLng(map(key))
                        value = data(r, index)
                        If IsError(value) Then
                            If Len(formula) = 0 Then Migration_Reject "MIGRATION.INVALID", key
                            value = Empty
                        End If
                        output(n, c - LBound(keys) + 1) = value
                    End If
                    If wbs And key = "S" And Not map.Exists(key) Then
                            value = UCase$(NormalizeTaskTypeValue(data(r, map("TASK_TYPE"))))
                            output(n, c - LBound(keys) + 1) = IIf(summaries.Exists(NormalizeWBS(CStr(data(r, map("WBS"))))) Or value = "MILESTONE", "Y", "N")
                    End If
                Next c
            End If
        Next r
    End If
    Dim compact() As Variant
    If n > 0 Then
        ReDim compact(1 To n, 1 To UBound(keys) - LBound(keys) + 1)
        For r = 1 To n
            For c = 1 To UBound(compact, 2)
                compact(r, c) = output(r, c)
            Next c
        Next r
        block.Add "Values", compact
    Else
        block.Add "Values", Empty
    End If
    block.Add "Keys", keys: block.Add "Count", n
    If wbs Then block.Add "Formulas", preservedFormulas
    Set MigrationData_ReadBlock = block
End Function

Private Function MigrationData_PrepareInputFormulas(ByVal tbl As ListObject, ByVal map As Object, ByVal keys As Variant, ByVal count As Long, ByRef formulas As Variant) As Object
    Dim result As Object, flags() As Boolean, columnRange As Range, area As Range
    Dim c As Long, r As Long, firstRow As Long, lastRow As Long, coverage As Variant
    Set result = CreateObject("Scripting.Dictionary")
    For c = LBound(keys) To UBound(keys)
        If map.Exists(keys(c)) Then
            Set columnRange = tbl.ListColumns(map(keys(c))).DataBodyRange
            coverage = columnRange.HasFormula
            ReDim flags(1 To count)
            If IsNull(coverage) Then
                For Each area In columnRange.SpecialCells(xlCellTypeFormulas).Areas
                    firstRow = area.Row - tbl.DataBodyRange.Row + 1
                    lastRow = firstRow + area.Rows.Count - 1
                    For r = firstRow To lastRow
                        flags(r) = True
                    Next r
                Next area
            ElseIf CBool(coverage) Then
                For r = 1 To count
                    flags(r) = True
                Next r
            End If
            For r = 1 To count
                If flags(r) Then result.Add CStr(r) & ":" & CStr(c - LBound(keys) + 1), formulas(r, map(keys(c)))
            Next r
        End If
    Next c
    Set MigrationData_PrepareInputFormulas = result
End Function

Private Function MigrationData_Formula(ByVal block As Object, ByVal row As Long, ByVal column As Long) As String
    Dim identity As String, formulas As Object
    If Not block.Exists("Formulas") Then Exit Function
    Set formulas = block("Formulas")
    identity = CStr(row) & ":" & CStr(column)
    If formulas.Exists(identity) Then MigrationData_Formula = CStr(formulas(identity))
End Function

Private Function MigrationData_ContentToken(ByVal block As Object, ByRef values As Variant, ByVal row As Long, ByVal column As Long) As String
    Dim formula As String
    formula = MigrationData_Formula(block, row, column)
    If Len(formula) > 0 Then
        MigrationData_ContentToken = "F:" & formula
    Else
        MigrationData_ContentToken = MigrationData_ValueToken(values(row, column))
    End If
End Function


Public Function MigrationData_Index(ByVal block As Object, ByVal field As String) As Long
    Dim keys As Variant, i As Long
    keys = block("Keys")
    For i = LBound(keys) To UBound(keys)
        If CStr(keys(i)) = field Then MigrationData_Index = i - LBound(keys) + 1: Exit Function
    Next i
    Err.Raise 5, "MigrationData_Index", field
End Function

Private Sub MigrationData_Validate(ByVal snapshot As Object)
    Dim tables As Object, block As Object, ids As Object, wbsIds As Object, data As Variant
    Dim r As Long, c As Long, id As String, wbs As String, kind As String, field As Variant, seen As Object
    Set tables = snapshot("Tables")
    Set block = tables("tbl_WBS")
    Set ids = CreateObject("Scripting.Dictionary")
    Set wbsIds = CreateObject("Scripting.Dictionary")
    data = block("Values")
    For r = 1 To CLng(block("Count"))
        id = CStr(data(r, MigrationData_Index(block, "ID")))
        If Len(Trim$(id)) = 0 Or ids.Exists(id) Then Migration_Reject "MIGRATION.INVALID", "ID " & id
        ids.Add id, True
        wbs = CStr(data(r, MigrationData_Index(block, "WBS")))
        If Len(Trim$(wbs)) = 0 Or wbsIds.Exists(wbs) Then Migration_Reject "MIGRATION.INVALID", "WBS " & wbs
        wbsIds.Add wbs, True
        If MigrationData_Formula(block, r, MigrationData_Index(block, "TASK_TYPE")) = "" Then kind = NormalizeTaskTypeValue(data(r, MigrationData_Index(block, "TASK_TYPE")))
        If MigrationData_Formula(block, r, MigrationData_Index(block, "CAL")) = "" Then
            If Not IsValidCalendarType(data(r, MigrationData_Index(block, "CAL"))) Then Migration_Reject "MIGRATION.INVALID", "Calendar " & wbs
        End If
        If MigrationData_Formula(block, r, MigrationData_Index(block, "S")) = "" Then kind = NormalizeSummaryDisplayValue(data(r, MigrationData_Index(block, "S")))
        For Each field In Array("BASELINE_START", "ACTUAL_START", "ACTUAL_FINISH", "FORECAST_START", "FORECAST_FINISH")
            If MigrationData_Formula(block, r, MigrationData_Index(block, CStr(field))) = "" Then MigrationData_Date data(r, MigrationData_Index(block, CStr(field))), CStr(field) & " " & wbs
        Next field
    Next r
    Set block = tables("tbl_CONSTRAINTS")
    data = block("Values")
    Set seen = CreateObject("Scripting.Dictionary")
    For r = 1 To CLng(block("Count"))
        id = CStr(data(r, MigrationData_Index(block, "ID")))
        If Not ids.Exists(id) Then Migration_Reject "MIGRATION.ORPHAN", id
        If seen.Exists(id) Then Migration_Reject "MIGRATION.INVALID", "Constraint ID " & id
        seen.Add id, True
        For Each field In Array("START_CONSTRAINT_DATE", "FINISH_CONSTRAINT_DATE", "DEADLINE")
            MigrationData_Date data(r, MigrationData_Index(block, CStr(field))), CStr(field) & " " & id
        Next field
        kind = UCase$(Trim$(CStr(data(r, MigrationData_Index(block, "ACTIVE")))))
        If InStr(1, "||YES|NO|Y|N|TRUE|FALSE|1|0|OUI|NON|", "|" & kind & "|", vbBinaryCompare) = 0 Then Migration_Reject "MIGRATION.INVALID", "Active " & id
    Next r
    For Each field In Array("tbl_EVENT_ACK", "tbl_DASHBOARD_SNAPSHOTS", "tbl_CALC_ALARM")
        Set block = tables(field)
        MigrationData_ValidateLedger block, CStr(field)
    Next field
End Sub

Private Sub MigrationData_ValidateLedger(ByVal block As Object, ByVal tableName As String)
    Dim data As Variant, seen As Object, r As Long, id As String
    data = block("Values")
    Set seen = CreateObject("Scripting.Dictionary")
    For r = 1 To CLng(block("Count"))
            Select Case tableName
                Case "tbl_EVENT_ACK"
                    If Len(Trim$(CStr(data(r, MigrationData_Index(block, "HASH"))))) = 0 Then Migration_Reject "MIGRATION.INVALID", "ACK hash"
                    id = CStr(data(r, MigrationData_Index(block, "HASH"))) & "|" & CStr(data(r, MigrationData_Index(block, "EVENT_TYPE"))) & "|" & CStr(data(r, MigrationData_Index(block, "SEVERITY")))
                Case "tbl_DASHBOARD_SNAPSHOTS"
                    id = CStr(data(r, MigrationData_Index(block, "SnapshotId")))
                    If Not IsNumeric(id) Then Migration_Reject "MIGRATION.INVALID", "SnapshotId"
                    If CDbl(id) < 1 Or CDbl(id) <> Fix(CDbl(id)) Then Migration_Reject "MIGRATION.INVALID", "SnapshotId"
                    MigrationData_Date data(r, MigrationData_Index(block, "SnapshotDateTime")), "SnapshotDateTime"
                Case Else
                    id = CStr(data(r, MigrationData_Index(block, "Event ID")))
                    MigrationData_Date data(r, MigrationData_Index(block, "Timestamp")), "Timestamp"
            End Select
            If Len(id) = 0 Or seen.Exists(id) Then Migration_Reject "MIGRATION.INVALID", tableName & " identity " & id
            seen.Add id, True
    Next r
End Sub

Private Sub MigrationData_Date(ByVal value As Variant, ByVal field As String)
    If IsEmpty(value) Or Len(CStr(value)) = 0 Then Exit Sub
    If Not IsNumeric(value) Then Migration_Reject "MIGRATION.INVALID", field
    If CDbl(value) < 1 Or CDbl(value) >= 2958466# Then Migration_Reject "MIGRATION.INVALID", field
End Sub

Public Function MigrationData_Digest(ByVal snapshot As Object) As String
    Dim parts() As String, tables As Object, block As Object, tableName As Variant
    Dim data As Variant, r As Long, c As Long, value As String, total As Long, n As Long
    Set tables = snapshot("Tables")
    total = 15
    For Each tableName In Array("tbl_WBS", "tbl_CONSTRAINTS", "tbl_EVENT_ACK", "tbl_DASHBOARD_SNAPSHOTS", "tbl_CALC_ALARM")
        Set block = tables(tableName)
        data = block("Keys")
        total = total + CLng(block("Count")) * (UBound(data) - LBound(data) + 1)
    Next tableName
    ReDim parts(0 To total - 1)
    For Each tableName In Array("tbl_WBS", "tbl_CONSTRAINTS", "tbl_EVENT_ACK", "tbl_DASHBOARD_SNAPSHOTS", "tbl_CALC_ALARM")
        Set block = tables(tableName): data = block("Values")
        parts(n) = CStr(tableName) & ":" & CStr(block("Count")) & ":": n = n + 1
        For r = 1 To CLng(block("Count"))
            For c = 1 To UBound(data, 2)
                value = MigrationData_ContentToken(block, data, r, c)
                parts(n) = CStr(Len(value)) & ":" & value: n = n + 1
            Next c
        Next r
    Next tableName
    Set block = snapshot("Settings"): data = block("Values")
    For r = 1 To 10
        value = MigrationData_ValueToken(data(r, 1))
        parts(n) = CStr(Len(value)) & ":" & value: n = n + 1
    Next r
    MigrationData_Digest = DeterministicDigest_SHA256Hex(Join(parts, ""))
End Function

Private Function MigrationData_ValueToken(ByVal value As Variant) As String
    If IsEmpty(value) Then
        MigrationData_ValueToken = "S:"
    ElseIf VarType(value) = vbString Then
        MigrationData_ValueToken = "S:" & CStr(value)
    ElseIf VarType(value) = vbBoolean Then
        MigrationData_ValueToken = "B:" & CStr(Abs(CBool(value)))
    ElseIf IsNumeric(value) Then
        MigrationData_ValueToken = "N:" & Trim$(Str$(CDbl(value)))
    Else
        Err.Raise 5, "MigrationData_ValueToken", "UNSUPPORTED_VALUE_TYPE"
    End If
End Function

Public Sub MigrationData_Write(ByVal snapshot As Object, ByVal wb As Workbook)
    Dim tables As Object, block As Object, tableName As Variant, tbl As ListObject, map As Object
    Dim values As Variant, output() As Variant, keys As Variant, r As Long, c As Long, n As Long
    Dim formula As String, hasFormulas As Boolean, oldAutoFill As Boolean, captured As Boolean
    Dim errorNumber As Long, errorSource As String, errorDescription As String
    On Error GoTo Failed
    oldAutoFill = Application.AutoCorrect.AutoFillFormulasInLists: captured = True
    Application.AutoCorrect.AutoFillFormulasInLists = False
    Set tables = snapshot("Tables")
    For Each tableName In Array("tbl_WBS", "tbl_CONSTRAINTS", "tbl_EVENT_ACK", "tbl_DASHBOARD_SNAPSHOTS", "tbl_CALC_ALARM")
        Set block = tables(tableName)
        Set tbl = MigrationData_Table(wb, CStr(tableName))
        Set map = MigrationData_ColumnMap(tbl)
        n = CLng(block("Count"))
        If Not tbl.DataBodyRange Is Nothing Then tbl.DataBodyRange.Delete
        If n > 0 Then
            tbl.Resize tbl.HeaderRowRange.Resize(n + 1, tbl.ListColumns.Count)
            ReDim output(1 To n, 1 To tbl.ListColumns.Count)
            values = block("Values"): keys = block("Keys")
            hasFormulas = False
            If block.Exists("Formulas") Then hasFormulas = (block("Formulas").Count > 0)
            For r = 1 To n
                For c = LBound(keys) To UBound(keys)
                    output(r, map(keys(c))) = values(r, c - LBound(keys) + 1)
                    If VarType(output(r, map(keys(c)))) = vbString Then
                        If Len(output(r, map(keys(c)))) > 0 Then
                            output(r, map(keys(c))) = "'" & output(r, map(keys(c)))
                        End If
                    End If
                Next c
            Next r
            ' A one-row empty table exposes its insertion range but no DataBodyRange.
            ' The bulk write materializes that row; verify semantic size afterwards.
            ' Formula transports invariant formulas and explicitly quoted literal strings.
            ' Numeric variants remain numeric, including when no formula is present.
            tbl.HeaderRowRange.Offset(1, 0).Resize(n, tbl.ListColumns.Count).Formula = output
            ' Do not assign a mixed formula array as an Excel calculated column.
            ' Apply only captured formula cells after the bulk literal projection.
            If hasFormulas Then
                For r = 1 To n
                    For c = LBound(keys) To UBound(keys)
                        formula = MigrationData_Formula(block, r, c - LBound(keys) + 1)
                        If Len(formula) > 0 Then tbl.DataBodyRange.Cells(r, map(keys(c))).Formula = formula
                    Next c
                Next r
            End If
            If tbl.DataBodyRange Is Nothing Or tbl.ListRows.Count <> n Then
                Err.Raise 5, "MigrationData_Write", "TABLE_WRITE_SIZE:" & CStr(tableName) & ":expected=" & CStr(n) & ":actual=" & CStr(tbl.ListRows.Count)
            End If
        End If
    Next tableName
    Application.AutoCorrect.AutoFillFormulasInLists = oldAutoFill
    Exit Sub
Failed:
    errorNumber = Err.Number: errorSource = Err.Source: errorDescription = Err.Description
    If captured Then Application.AutoCorrect.AutoFillFormulasInLists = oldAutoFill
    Err.Raise errorNumber, errorSource, errorDescription
End Sub

Public Sub MigrationData_VerifyPreserved(ByVal expected As Object, ByVal actual As Object, Optional ByVal allowWelcomeAck As Boolean = False)
    Dim name As Variant, before As Object, after As Object, rows As Object
    Dim a As Variant, b As Variant, keys As Variant, r As Long, c As Long, targetRow As Long, identity As String
    Dim field As String, left As String, right As String
    Dim welcome As Object, welcomeIdentity As String, expectedAckCount As Long, welcomeInSource As Boolean
    If allowWelcomeAck Then
        Set welcome = ProjectWelcome_Message()
        welcomeIdentity = CStr(welcome("Hash")) & "|" & CStr(welcome("EventType")) & "|" & CStr(welcome("Type"))
    End If
    For Each name In Array("tbl_WBS", "tbl_CONSTRAINTS", "tbl_EVENT_ACK", "tbl_DASHBOARD_SNAPSHOTS", "tbl_CALC_ALARM")
        Set before = expected("Tables")(name): Set after = actual("Tables")(name)
        a = before("Values"): b = after("Values"): keys = before("Keys")
        Set rows = CreateObject("Scripting.Dictionary")
        For r = 1 To CLng(after("Count"))
            identity = MigrationData_RowIdentity(after, b, r, CStr(name))
            If rows.Exists(identity) Then Err.Raise 5, "MigrationData_VerifyPreserved", "DUPLICATE_RESULT_IDENTITY:" & CStr(name)
            rows.Add identity, r
        Next r
        If name = "tbl_WBS" Or (name = "tbl_EVENT_ACK" And Not allowWelcomeAck) Then
            If before("Count") <> after("Count") Then Err.Raise 5, "MigrationData_VerifyPreserved", "ROW_COUNT_CHANGED:" & CStr(name)
        End If
        If name = "tbl_EVENT_ACK" And allowWelcomeAck Then
            welcomeInSource = False
            For r = 1 To CLng(before("Count"))
                If MigrationData_RowIdentity(before, a, r, CStr(name)) = welcomeIdentity Then welcomeInSource = True
            Next r
            expectedAckCount = CLng(before("Count"))
            If Not welcomeInSource Then expectedAckCount = expectedAckCount + 1
            If CLng(after("Count")) <> expectedAckCount Or Not rows.Exists(welcomeIdentity) Then
                Err.Raise 5, "MigrationData_VerifyPreserved", "UNEXPECTED_ACK_ROWS"
            End If
            If Not CBool(b(CLng(rows(welcomeIdentity)), MigrationData_Index(after, "ACKNOWLEDGED"))) Then
                Err.Raise 5, "MigrationData_VerifyPreserved", "WELCOME_ACK_NOT_SET"
            End If
        End If
        For r = 1 To CLng(before("Count"))
            identity = MigrationData_RowIdentity(before, a, r, CStr(name))
            If Not rows.Exists(identity) Then Err.Raise 5, "MigrationData_VerifyPreserved", "IMPORTED_ROW_MISSING:" & CStr(name) & ":" & identity
            targetRow = CLng(rows(identity))
            If Not (name = "tbl_EVENT_ACK" And allowWelcomeAck And identity = welcomeIdentity) Then
            For c = LBound(keys) To UBound(keys)
                field = CStr(keys(c))
                left = MigrationData_ContentToken(before, a, r, c - LBound(keys) + 1)
                right = MigrationData_ContentToken(after, b, targetRow, MigrationData_Index(after, field))
                If left <> right Then Err.Raise 5, "MigrationData_VerifyPreserved", "IMPORTED_VALUE_CHANGED:" & CStr(name) & ":" & identity & ":" & field
            Next c
            End If
        Next r
    Next name
    a = expected("Settings")("Values"): b = actual("Settings")("Values")
    For r = 1 To 10
        If MigrationData_ValueToken(a(r, 1)) <> MigrationData_ValueToken(b(r, 1)) Then Err.Raise 5, "MigrationData_VerifyPreserved", "SETTINGS_CHANGED:" & CStr(r)
    Next r
End Sub

Private Function MigrationData_RowIdentity(ByVal block As Object, ByRef data As Variant, ByVal row As Long, ByVal tableName As String) As String
    Select Case tableName
        Case "tbl_WBS", "tbl_CONSTRAINTS"
            MigrationData_RowIdentity = CStr(data(row, MigrationData_Index(block, "ID")))
        Case "tbl_DASHBOARD_SNAPSHOTS"
            MigrationData_RowIdentity = CStr(data(row, MigrationData_Index(block, "SnapshotId")))
        Case "tbl_CALC_ALARM"
            MigrationData_RowIdentity = CStr(data(row, MigrationData_Index(block, "Event ID")))
        Case "tbl_EVENT_ACK"
            MigrationData_RowIdentity = CStr(data(row, MigrationData_Index(block, "HASH"))) & "|" & CStr(data(row, MigrationData_Index(block, "EVENT_TYPE"))) & "|" & CStr(data(row, MigrationData_Index(block, "SEVERITY")))
    End Select
End Function
