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
            If tbl.name = name Then
                If Not MigrationData_Table Is Nothing Then Err.Raise 5, , "DUPLICATE_TABLE:" & name
                Set MigrationData_Table = tbl
            End If
        Next tbl
    Next ws
    If required And MigrationData_Table Is Nothing Then Migration_Reject "MIGRATION.PROFILE", name
End Function

Public Function MigrationData_ColumnMap(ByVal tbl As ListObject) As Object
    If tbl.name = FIN_TABLE Or tbl.name = FIN_EXPENSES Then
        Set MigrationData_ColumnMap = SchemaBuildColumnKeyMap(tbl, tbl.name): Exit Function
    End If
    Dim result As Object, titles As Object, key As Variant, column As listColumn
    Dim keys As Variant, language As Variant, title As String
    Set result = CreateObject("Scripting.Dictionary")
    Set titles = CreateObject("Scripting.Dictionary")
    titles.CompareMode = vbTextCompare
    If SchemaIsVisibleTable(tbl.name) Then
        keys = SchemaColumnKeys(tbl.name)
        For Each key In keys
            For Each language In Array("EN", "FR")
                title = SchemaColumnTitle(tbl.name, CStr(key), CStr(language))
                If Not titles.exists(title) Then titles.Add title, CStr(key)
            Next language
        Next key
        If tbl.name = VTS_TABLE_WBS Then titles.Add "Package", VTS_COL_PROJECT
        For Each column In tbl.ListColumns
            If titles.exists(column.name) Then
                key = titles(column.name)
                If result.exists(key) Then Migration_Reject "MIGRATION.INVALID", "Package/Project or duplicated header: " & column.name
                result.Add key, column.Index
            Else
                Migration_Reject "MIGRATION.PROFILE", tbl.name & ": " & column.name
            End If
        Next column
    Else
        For Each column In tbl.ListColumns
            result.Add column.name, column.Index
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
    If version <> 0 And version <> 1 And version <> 2 And version <> CURRENT_SCHEMA_VERSION Then Migration_Reject "MIGRATION.PROFILE"
    If version >= 2 Then
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
    If Not map.exists("S") Then profile = profile & "/S_DEFAULT"
    If Not map.exists("PROJECT") Then Migration_Reject "MIGRATION.PROFILE", "PROJECT"
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
            If version >= 2 And tableName = "tbl_DASHBOARD_SNAPSHOTS" Then Migration_Reject "MIGRATION.PROFILE", CStr(tableName)
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
    snapshot.Add "Finance", MigrationData_ReadFinance(source, version)
    phase = "VALIDATION"
    MigrationData_Validate snapshot
    phase = "WBS_FORMATS"
    snapshot.Add "WbsFormats", MigrationData_ReadWbsFormats(MigrationData_Table(source, "tbl_WBS"))
    phase = "COLUMN_WIDTHS"
    snapshot.Add "ColumnWidths", MigrationData_ReadColumnWidths(source)
    phase = "DIGEST"
    snapshot.Add "Digest", MigrationData_Digest(snapshot)
    Set MigrationData_Read = snapshot
    Exit Function
Failed:
    Dim errorNumber As Long, errorSource As String, errorDescription As String
    errorNumber = Err.number: errorSource = Err.source: errorDescription = Err.description
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
    number = Err.number: origin = Err.source: diagnostic = Err.description
    If origin <> "MigrationValidation" Then Err.Raise number, origin, diagnostic
    Set block = CreateObject("Scripting.Dictionary")
    block.Add "Keys", keys: block.Add "Values", Empty: block.Add "Count", 0&
    block.Add "Available", False
    block.Add "UnavailableDiagnostic", diagnostic
    Set MigrationData_ReadOptionalHistory = block
End Function

Private Function MigrationData_PhysicalKeys(ByVal tbl As ListObject) As Variant
    Dim keys() As Variant, i As Long
    ReDim keys(0 To tbl.ListColumns.count - 1)
    For i = 1 To tbl.ListColumns.count
        keys(i - 1) = tbl.ListColumns(i).name
    Next i
    MigrationData_PhysicalKeys = keys
End Function

Private Function MigrationData_ReadBlock(ByVal tbl As ListObject, ByVal keys As Variant, Optional ByVal wbs As Boolean = False) As Object
    Dim block As Object, map As Object, data As Variant, formulas As Variant, output() As Variant
    Dim r As Long, c As Long, n As Long, count As Long, key As String, Index As Long, value As Variant
    Dim summaries As Object, emptyWbsRow As Boolean, hasFormula As Variant, inputFormulas As Object, preservedFormulas As Object, formula As String
    Set block = CreateObject("Scripting.Dictionary")
    If wbs Then Set preservedFormulas = CreateObject("Scripting.Dictionary")
    Set map = MigrationData_ColumnMap(tbl)
    If Not SchemaIsVisibleTable(tbl.name) Then
        If tbl.ListColumns.count <> UBound(keys) - LBound(keys) + 1 Then Migration_Reject "MIGRATION.PROFILE", tbl.name & " column count"
    End If
    For c = LBound(keys) To UBound(keys)
        key = CStr(keys(c))
        If Not map.exists(key) Then
            If Not (wbs And key = "S") Then Migration_Reject "MIGRATION.PROFILE", tbl.name & ": " & key
        End If
    Next c
    count = tbl.ListRows.count
    If Not tbl.DataBodyRange Is Nothing And count > 0 Then
        data = tbl.DataBodyRange.Value2
        For c = LBound(keys) To UBound(keys)
            key = CStr(keys(c))
            If map.exists(key) Then
                hasFormula = tbl.ListColumns(map(key)).DataBodyRange.hasFormula
                If Not wbs Then
                    If IsNull(hasFormula) Then Migration_Reject "MIGRATION.FORMULA", tbl.ListColumns(map(key)).name
                    If CBool(hasFormula) Then Migration_Reject "MIGRATION.FORMULA", tbl.ListColumns(map(key)).name
                End If
            End If
        Next c
        If wbs Then
            formulas = tbl.DataBodyRange.formula
            Set summaries = DataSync_BuildSummaryWbsLookupFromArray(data, map)
            Set inputFormulas = MigrationData_PrepareInputFormulas(tbl, map, keys, count, formulas)
        End If
        ReDim output(1 To count, 1 To UBound(keys) - LBound(keys) + 1)
        For r = 1 To count
            emptyWbsRow = False
            If wbs Then
                If isError(data(r, map("ID"))) Then Migration_Reject "MIGRATION.INVALID", "ID"
                emptyWbsRow = (Len(CStr(data(r, map("ID")))) = 0)
            Else
                emptyWbsRow = True
                For c = LBound(keys) To UBound(keys)
                    If isError(data(r, map(keys(c)))) Then
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
                        If map.exists(key) Then
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
                        If inputFormulas.exists(CStr(r) & ":" & CStr(c - LBound(keys) + 1)) Then
                            formula = inputFormulas(CStr(r) & ":" & CStr(c - LBound(keys) + 1))
                            preservedFormulas.Add CStr(n) & ":" & CStr(c - LBound(keys) + 1), formula
                        End If
                    End If
                    If map.exists(key) Then
                        Index = CLng(map(key))
                        value = data(r, Index)
                        If isError(value) Then
                            If Len(formula) = 0 Then Migration_Reject "MIGRATION.INVALID", key
                            value = Empty
                        End If
                        output(n, c - LBound(keys) + 1) = value
                    End If
                    If wbs And key = "S" And Not map.exists(key) Then
                            value = UCase$(NormalizeTaskTypeValue(data(r, map("TASK_TYPE"))))
                            output(n, c - LBound(keys) + 1) = IIf(summaries.exists(NormalizeWBS(CStr(data(r, map("WBS"))))) Or value = "MILESTONE", "Y", "N")
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
        If map.exists(keys(c)) Then
            Set columnRange = tbl.ListColumns(map(keys(c))).DataBodyRange
            coverage = columnRange.hasFormula
            ReDim flags(1 To count)
            If IsNull(coverage) Then
                For Each area In columnRange.SpecialCells(xlCellTypeFormulas).Areas
                    firstRow = area.row - tbl.DataBodyRange.row + 1
                    lastRow = firstRow + area.rows.count - 1
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
    If Not block.exists("Formulas") Then Exit Function
    Set formulas = block("Formulas")
    identity = CStr(row) & ":" & CStr(column)
    If formulas.exists(identity) Then MigrationData_Formula = CStr(formulas(identity))
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
        If Len(Trim$(id)) = 0 Or ids.exists(id) Then Migration_Reject "MIGRATION.INVALID", "ID " & id
        ids.Add id, True
        wbs = CStr(data(r, MigrationData_Index(block, "WBS")))
        If Len(Trim$(wbs)) = 0 Or wbsIds.exists(wbs) Then Migration_Reject "MIGRATION.INVALID", "WBS " & wbs
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
        If Not ids.exists(id) Then Migration_Reject "MIGRATION.ORPHAN", id
        If seen.exists(id) Then Migration_Reject "MIGRATION.INVALID", "Constraint ID " & id
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
            If Len(id) = 0 Or seen.exists(id) Then Migration_Reject "MIGRATION.INVALID", tableName & " identity " & id
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
    total = 16
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
    parts(n) = MigrationData_FinanceToken(snapshot("Finance"))
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
            tbl.Resize tbl.HeaderRowRange.Resize(n + 1, tbl.ListColumns.count)
            ReDim output(1 To n, 1 To tbl.ListColumns.count)
            values = block("Values"): keys = block("Keys")
            hasFormulas = False
            If block.exists("Formulas") Then hasFormulas = (block("Formulas").count > 0)
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
            tbl.HeaderRowRange.offset(1, 0).Resize(n, tbl.ListColumns.count).formula = output
            ' Do not assign a mixed formula array as an Excel calculated column.
            ' Apply only captured formula cells after the bulk literal projection.
            If hasFormulas Then
                For r = 1 To n
                    For c = LBound(keys) To UBound(keys)
                        formula = MigrationData_Formula(block, r, c - LBound(keys) + 1)
                        If Len(formula) > 0 Then tbl.DataBodyRange.cells(r, map(keys(c))).formula = formula
                    Next c
                Next r
            End If
            If tbl.DataBodyRange Is Nothing Or tbl.ListRows.count <> n Then
                Err.Raise 5, "MigrationData_Write", "TABLE_WRITE_SIZE:" & CStr(tableName) & ":expected=" & CStr(n) & ":actual=" & CStr(tbl.ListRows.count)
            End If
        End If
    Next tableName
    If snapshot.exists("WbsFormats") Then MigrationData_WriteWbsFormats snapshot("WbsFormats"), MigrationData_Table(wb, "tbl_WBS")
    MigrationData_WriteFinance snapshot("Finance"), wb
    Application.AutoCorrect.AutoFillFormulasInLists = oldAutoFill
    Exit Sub
Failed:
    errorNumber = Err.number: errorSource = Err.source: errorDescription = Err.description
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
            If rows.exists(identity) Then Err.Raise 5, "MigrationData_VerifyPreserved", "DUPLICATE_RESULT_IDENTITY:" & CStr(name)
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
            If CLng(after("Count")) <> expectedAckCount Or Not rows.exists(welcomeIdentity) Then
                Err.Raise 5, "MigrationData_VerifyPreserved", "UNEXPECTED_ACK_ROWS"
            End If
            If Not CBool(b(CLng(rows(welcomeIdentity)), MigrationData_Index(after, "ACKNOWLEDGED"))) Then
                Err.Raise 5, "MigrationData_VerifyPreserved", "WELCOME_ACK_NOT_SET"
            End If
        End If
        For r = 1 To CLng(before("Count"))
            identity = MigrationData_RowIdentity(before, a, r, CStr(name))
            If Not rows.exists(identity) Then Err.Raise 5, "MigrationData_VerifyPreserved", "IMPORTED_ROW_MISSING:" & CStr(name) & ":" & identity
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
    If MigrationData_FinanceToken(expected("Finance")) <> MigrationData_FinanceToken(actual("Finance")) Then Err.Raise 5, "Migration", "FINANCE_INPUTS_CHANGED"
    If expected.exists("WbsFormats") Then MigrationData_VerifyWbsFormats expected("WbsFormats"), actual("WbsFormats")
    If expected.exists("ColumnWidths") Then MigrationData_VerifyColumnWidths expected("ColumnWidths"), actual("ColumnWidths")
End Sub

Private Function MigrationData_WidthColumns(ByVal tbl As ListObject) As Object
    If tbl.name = FIN_TABLE Or tbl.name = FIN_EXPENSES Then
        Set MigrationData_WidthColumns = Finance_Map(tbl)
    Else
        Set MigrationData_WidthColumns = MigrationData_ColumnMap(tbl)
    End If
End Function

Private Function MigrationData_ReadColumnWidths(ByVal wb As Workbook) As Object
    Dim result As Object, widths As Object, map As Object, tbl As ListObject, ws As Worksheet
    Dim tableName As Variant, key As Variant, width As Double, i As Long, keys As Variant
    Dim scope As clsPerfScope
    Set scope = Profiler_BeginScope("Migration.ColumnWidths.Read", "Migration")
    Set result = CreateObject("Scripting.Dictionary")
    For Each tableName In Array(VTS_TABLE_WBS, VTS_TABLE_CONSTRAINTS, FIN_TABLE, FIN_EXPENSES)
        Set tbl = MigrationData_Table(wb, CStr(tableName), False)
        If Not tbl Is Nothing Then
            Set widths = CreateObject("Scripting.Dictionary")
            Set map = MigrationData_WidthColumns(tbl)
            For Each key In map.keys
                If Not (CStr(tableName) = FIN_TABLE And CStr(key) = "TASK_TYPE") Then
                    width = CDbl(tbl.ListColumns(CLng(map(key))).Range.ColumnWidth)
                    If width > 0 Then widths.Add CStr(key), width
                End If
            Next key
            result.Add CStr(tableName), widths
        End If
    Next tableName
    On Error Resume Next
    Set ws = wb.Worksheets("FINANCE SUMMARY")
    On Error GoTo 0
    If Not ws Is Nothing Then
        ' Monthly and total tables share worksheet columns, even before regeneration.
        keys = Finance_OutputKeys("tbl_FINANCE_TOTALS")
        Set widths = CreateObject("Scripting.Dictionary")
        For i = 1 To UBound(keys) - LBound(keys) + 1
            width = CDbl(ws.columns(i).ColumnWidth)
            If width > 0 Then widths.Add CStr(i), width
        Next i
        result.Add "FINANCE SUMMARY", widths
    End If
    Set MigrationData_ReadColumnWidths = result
End Function

Public Sub MigrationData_ApplyColumnWidths(ByVal snapshot As Object, ByVal wb As Workbook)
    Dim bundle As Object, widths As Object, map As Object, tbl As ListObject, ws As Worksheet
    Dim owner As Variant, key As Variant, target As Range, wasHidden As Boolean
    Dim scope As clsPerfScope
    If Not snapshot.exists("ColumnWidths") Then Exit Sub
    Set scope = Profiler_BeginScope("Migration.ColumnWidths.Write", "Migration")
    Set bundle = snapshot("ColumnWidths")
    For Each owner In bundle.keys
        Set widths = bundle(owner)
        If CStr(owner) = "FINANCE SUMMARY" Then
            Set ws = wb.Worksheets(CStr(owner))
        Else
            Set tbl = MigrationData_Table(wb, CStr(owner))
            Set map = MigrationData_WidthColumns(tbl)
        End If
        For Each key In widths.keys
            If CStr(owner) = "FINANCE SUMMARY" Then
                Set target = ws.columns(CLng(key))
            Else
                If Not map.exists(key) Then Err.Raise 5, "Migration", "WIDTH_FIELD_MISSING:" & owner & ":" & key
                Set target = tbl.ListColumns(CLng(map(key))).Range.EntireColumn
            End If
            If Abs(CDbl(target.ColumnWidth) - CDbl(widths(key))) > 0.001 Then
                wasHidden = CBool(target.hidden)
                target.ColumnWidth = CDbl(widths(key))
                target.hidden = wasHidden
            End If
        Next key
    Next owner
End Sub

Private Sub MigrationData_VerifyColumnWidths(ByVal expected As Object, ByVal actual As Object)
    Dim owner As Variant, key As Variant, before As Object, after As Object
    For Each owner In expected.keys
        If Not actual.exists(owner) Then Err.Raise 5, "Migration", "WIDTH_OWNER_MISSING:" & owner
        Set before = expected(owner): Set after = actual(owner)
        For Each key In before.keys
            If Not after.exists(key) Then Err.Raise 5, "Migration", "WIDTH_FIELD_MISSING:" & owner & ":" & key
            If Abs(CDbl(before(key)) - CDbl(after(key))) > 0.01 Then Err.Raise 5, "Migration", "COLUMN_WIDTH_CHANGED:" & owner & ":" & key
        Next key
    Next owner
End Sub

' Formats travel with the existing input schema and task ID, never whole rows.
Private Function MigrationData_ReadWbsFormats(ByVal tbl As ListObject) As Object
    Dim result As Object, map As Object, rows As Object, key As Variant, data As Variant
    Dim styles As Object, styleIds As Variant, styleId As String, p As Variant, r As Long, identity As String
    Dim scope As clsPerfScope
    Set scope = Profiler_BeginScope("Migration.WBS.Formats.Read", "Migration")
    Set result = CreateObject("Scripting.Dictionary")
    Set map = MigrationData_ColumnMap(tbl)
    If tbl.ListRows.count > 0 And Not tbl.DataBodyRange Is Nothing Then
        data = tbl.DataBodyRange.Value2
        styleIds = MigrationData_WbsStyleIds(tbl.DataBodyRange)
        Set styles = CreateObject("Scripting.Dictionary")
        For Each key In MigrationData_WbsInputs()
            If map.exists(key) Then
                Set rows = CreateObject("Scripting.Dictionary")
                For r = 1 To tbl.ListRows.count
                    identity = CStr(data(r, map("ID")))
                    If Len(identity) > 0 Then
                        styleId = MigrationData_StyleContext(styleIds, r, CLng(map(key)))
                        If Not styles.exists(styleId) Then
                            p = MigrationData_FormatProfile(tbl.DataBodyRange.cells(r, map(key)))
                            styles.Add styleId, p
                        End If
                        rows.Add identity, styles(styleId)
                    End If
                Next r
                result.Add CStr(key), rows
            End If
        Next key
    End If
    Set MigrationData_ReadWbsFormats = result
End Function

Private Function MigrationData_StyleContext(ByRef ids As Variant, ByVal row As Long, ByVal column As Long) As String
    Dim result As String
    result = CStr(ids(row, column))
    ' Excel exposes shared edge borders through adjacent cells, including range boundaries.
    If row > 1 Then result = result & "|" & ids(row - 1, column) Else result = result & "|TOP"
    If row < UBound(ids, 1) Then result = result & "|" & ids(row + 1, column) Else result = result & "|BOTTOM"
    If column > 1 Then result = result & "|" & ids(row, column - 1) Else result = result & "|LEFT"
    If column < UBound(ids, 2) Then result = result & "|" & ids(row, column + 1) Else result = result & "|RIGHT"
    MigrationData_StyleContext = result
End Function

Private Function MigrationData_WbsStyleIds(ByVal cells As Range) As Variant
    Dim doc As Object, table As Object, node As Object, cell As Object
    Dim ids() As String, columns() As String, r As Long, c As Long, i As Long, span As Long
    Dim n As Long, width As Long, rowStyle As String
    n = cells.rows.count: width = cells.columns.count
    ReDim ids(1 To n, 1 To width): ReDim columns(1 To width)
    Set doc = CreateObject("MSXML2.DOMDocument.6.0")
    doc.async = False: doc.resolveExternals = False
    doc.SetProperty "ProhibitDTD", True
    doc.SetProperty "SelectionNamespaces", "xmlns:ss='urn:schemas-microsoft-com:office:spreadsheet'"
    ' Native bulk projection supplies style identities only; business values/formulas keep their existing path.
    If Not doc.LoadXML(CStr(cells.value(xlRangeValueXMLSpreadsheet))) Then Err.Raise 5, "MigrationData_WbsStyleIds", "INVALID_RANGE_STYLE_XML"
    Set table = doc.SelectSingleNode("/ss:Workbook/ss:Worksheet/ss:Table")
    If table Is Nothing Then Err.Raise 5, "MigrationData_WbsStyleIds", "MISSING_RANGE_STYLE_TABLE"
    For i = 1 To width: columns(i) = "Default": Next i
    For Each node In table.SelectNodes("ss:Column")
        c = CLng(MigrationData_StyleAttribute(node, "Index", CStr(c + 1)))
        span = CLng(MigrationData_StyleAttribute(node, "Span", "0"))
        If c < 1 Or c + span > width Then Err.Raise 5, "MigrationData_WbsStyleIds", "STYLE_COLUMN_OUT_OF_RANGE"
        For i = c To c + span
            columns(i) = MigrationData_StyleAttribute(node, "StyleID", "Default")
        Next i
        c = c + span
    Next node
    For r = 1 To n
        For c = 1 To width: ids(r, c) = columns(c): Next c
    Next r
    r = 0
    For Each node In table.SelectNodes("ss:Row")
        r = CLng(MigrationData_StyleAttribute(node, "Index", CStr(r + 1)))
        If r < 1 Or r > n Then Err.Raise 5, "MigrationData_WbsStyleIds", "STYLE_ROW_OUT_OF_RANGE"
        rowStyle = MigrationData_StyleAttribute(node, "StyleID", "")
        If Len(rowStyle) > 0 Then
            For c = 1 To width: ids(r, c) = rowStyle: Next c
        End If
        c = 0
        For Each cell In node.SelectNodes("ss:Cell")
            c = CLng(MigrationData_StyleAttribute(cell, "Index", CStr(c + 1)))
            If c < 1 Or c > width Then Err.Raise 5, "MigrationData_WbsStyleIds", "STYLE_CELL_OUT_OF_RANGE"
            ids(r, c) = MigrationData_StyleAttribute(cell, "StyleID", ids(r, c))
            If Not cell.SelectSingleNode("ss:Data/*") Is Nothing Then ids(r, c) = ids(r, c) & "|INLINE:" & CStr(r) & ":" & CStr(c)
        Next cell
    Next node
    MigrationData_WbsStyleIds = ids
End Function

Private Function MigrationData_StyleAttribute(ByVal node As Object, ByVal key As String, ByVal defaultValue As String) As String
    Dim value As Variant
    value = node.getAttribute("ss:" & key)
    If IsNull(value) Then
        MigrationData_StyleAttribute = defaultValue
    ElseIf Len(CStr(value)) = 0 Then
        MigrationData_StyleAttribute = defaultValue
    Else
        MigrationData_StyleAttribute = CStr(value)
    End If
End Function

Private Function MigrationData_FormatProfile(ByVal cells As Range) As Variant
    ' Font 0..2, fill 3..7, six borders 8..31, then bold/italic/font name 32..34.
    Dim p(0 To 34) As Variant, f As font, fill As Interior, edge As border
    Dim axes As Variant, i As Long, j As Long, at As Long
    Set f = cells.font: Set fill = cells.Interior
    p(0) = f.ColorIndex: p(1) = f.color: p(2) = f.size
    p(32) = f.bold: p(33) = f.italic: p(34) = f.name
    p(3) = fill.Pattern: p(4) = fill.ColorIndex: p(5) = fill.color
    p(6) = fill.PatternColorIndex: p(7) = fill.PatternColor
    axes = Array(xlDiagonalDown, xlDiagonalUp, xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight)
    For i = 0 To 5
        Set edge = cells.Borders(axes(i)): at = 8 + i * 4
        p(at) = edge.LineStyle: p(at + 1) = edge.ColorIndex
        p(at + 2) = edge.color: p(at + 3) = edge.weight
    Next i
    For i = 0 To 34
        If IsNull(p(i)) Then Migration_Reject "MIGRATION.PROFILE", "WBS mixed character formatting"
    Next i
    ' Preserve resolved colors, not workbook-specific palette/automatic metadata.
    For Each axes In Array(0, 4, 6, 9, 13, 17, 21, 25, 29)
        p(axes) = 0&
    Next axes
    If p(3) = xlNone Then
        p(4) = 0&: p(5) = vbWhite: p(6) = 0&: p(7) = vbBlack
    End If
    For i = 0 To 5
        at = 8 + i * 4
        If p(at) = xlLineStyleNone Then
            p(at + 1) = 0&: p(at + 2) = vbBlack: p(at + 3) = xlThin
        End If
    Next i
    MigrationData_FormatProfile = p
End Function

Private Function MigrationData_FormatCanGroup(ByVal p As Variant) As Boolean
    Dim j As Long
    If p(8) <> xlLineStyleNone Or p(12) <> xlLineStyleNone Then Exit Function
    For j = 0 To 3
        If p(20 + j) <> p(24 + j) Then Exit Function
    Next j
    MigrationData_FormatCanGroup = True
End Function

Private Function MigrationData_FormatToken(ByVal p As Variant) As String
    Dim i As Long, result As String
    For i = LBound(p) To UBound(p)
        result = result & "|" & CStr(p(i))
    Next i
    MigrationData_FormatToken = result
End Function

Private Sub MigrationData_WriteWbsFormats(ByVal formats As Object, ByVal tbl As ListObject)
    Dim map As Object, rows As Object, data As Variant, key As Variant
    Dim r As Long, last As Long, identity As String, p As Variant, nextProfile As Variant, token As String
    Dim scope As clsPerfScope
    Set scope = Profiler_BeginScope("Migration.WBS.Formats.Write", "Migration")
    If tbl.DataBodyRange Is Nothing Or tbl.ListRows.count = 0 Then Exit Sub
    Set map = MigrationData_ColumnMap(tbl)
    data = tbl.DataBodyRange.Value2
    For Each key In formats.keys
        Set rows = formats(key): r = 1
        Do While r <= tbl.ListRows.count
            identity = CStr(data(r, map("ID")))
            If Not rows.exists(identity) Then Err.Raise 5, "MigrationData_WriteWbsFormats", "FORMAT_ID_MISSING:" & identity
            p = rows(identity): token = MigrationData_FormatToken(p): last = r
            If MigrationData_FormatCanGroup(p) Then
                Do While last < tbl.ListRows.count
                    identity = CStr(data(last + 1, map("ID")))
                    If Not rows.exists(identity) Then Exit Do
                    nextProfile = rows(identity)
                    If MigrationData_FormatToken(nextProfile) <> token Then Exit Do
                    last = last + 1
                Loop
            End If
            MigrationData_ApplyFormat tbl.DataBodyRange.cells(r, map(key)).Resize(last - r + 1), p
            r = last + 1
        Loop
    Next key
End Sub

Private Sub MigrationData_ApplyFormat(ByVal cells As Range, ByVal p As Variant)
    Dim axes As Variant, edge As border, i As Long, at As Long
    With cells.font
        If p(0) = xlColorIndexAutomatic Then .ColorIndex = p(0) Else .color = p(1)
        .size = p(2)
        .bold = p(32): .italic = p(33): .name = p(34)
    End With
    With cells.Interior
        If p(3) <> xlNone Then
            .color = p(5)
            If p(6) = xlColorIndexAutomatic Then .PatternColorIndex = p(6) Else .PatternColor = p(7)
        End If
        .Pattern = p(3)
    End With
    axes = Array(xlDiagonalDown, xlDiagonalUp, xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight, xlInsideHorizontal)
    For i = 0 To IIf(cells.rows.count > 1, 6, 5)
        at = 8 + i * 4
        If i = 6 Then at = 20
        Set edge = cells.Borders(axes(i))
        If p(at) <> xlLineStyleNone Then
            If p(at + 1) = xlColorIndexAutomatic Then edge.ColorIndex = p(at + 1) Else edge.color = p(at + 2)
            edge.weight = p(at + 3)
        End If
        edge.LineStyle = p(at)
    Next i
End Sub

Private Sub MigrationData_VerifyWbsFormats(ByVal expected As Object, ByVal actual As Object)
    Dim key As Variant, identity As Variant, rows As Object, result As Object, a As Variant, b As Variant, i As Long
    For Each key In expected.keys
        If Not actual.exists(key) Then Err.Raise 5, "MigrationData_VerifyPreserved", "FORMAT_FIELD_MISSING:" & key
        Set rows = expected(key): Set result = actual(key)
        If rows.count <> result.count Then Err.Raise 5, "MigrationData_VerifyPreserved", "FORMAT_ROW_COUNT_CHANGED:" & key
        For Each identity In rows.keys
            If Not result.exists(identity) Then Err.Raise 5, "MigrationData_VerifyPreserved", "FORMAT_ID_MISSING:" & identity
            a = rows(identity): b = result(identity)
            For i = LBound(a) To UBound(a)
                If a(i) <> b(i) Then Err.Raise 5, "MigrationData_VerifyPreserved", _
                    "IMPORTED_FORMAT_CHANGED:" & key & ":" & identity & ":property=" & CStr(i) & ":expected=" & CStr(a(i)) & ":actual=" & CStr(b(i))
            Next i
        Next identity
    Next key
End Sub

Private Function MigrationData_ReadFinance(ByVal wb As Workbook, ByVal version As Long) As Object
    Dim result As Object, rows As Collection, expenses As Collection, d As Object, graph As Object
    Set result = CreateObject("Scripting.Dictionary")
    Set rows = New Collection: Set expenses = New Collection
    result.Add "Currency", "EUR": result.Add "Vocabulary", "BEGINNER"
    result.Add "Language", Settings_GetOwnerLanguage("FINANCE", wb)
    If version = 3 Then
        Set graph = Finance_DomainSnapshot(wb)
        result("Currency") = CStr(graph("META")(1)("CURRENCY"))
        result("Vocabulary") = Settings_FinanceVocabulary(wb)
    ElseIf version = 2 Then
        Set rows = Finance_Read(Finance_Table(FIN_TABLE, wb))
        Set expenses = Finance_Read(Finance_Table(FIN_EXPENSES, wb))
        Finance_Validate rows, expenses
        result("Currency") = CStr(wb.Worksheets("FINANCE").Range("S1").Value2)
        If Len(result("Currency")) <> 3 Or result("Currency") Like "*[!A-Z]*" Then Migration_Reject "MIGRATION.INVALID", "Finance currency"
        result("Vocabulary") = Settings_FinanceVocabulary(wb)
        For Each d In rows
            d("ACTUAL") = Empty: d("EAC") = Empty: d("VARIANCE") = Empty: d("PROGRESS") = Empty: d("DIAGNOSTIC") = Empty
            If d("STATE") = "ACTIVE" Then d("START") = Empty: d("FINISH") = Empty
        Next d
    End If
    If graph Is Nothing Then Set graph = Finance_DomainFromLegacy(rows, expenses, CStr(result("Currency")))
    result.Add "Domain", graph
    result.Add "Rows", rows: result.Add "Expenses", expenses
    Set MigrationData_ReadFinance = result
End Function

Private Sub MigrationData_WriteFinance(ByVal bundle As Object, ByVal wb As Workbook)
    Dim ws As Worksheet, graph As Object, rows As Collection, expenses As Collection, emptyRows As New Collection
    Set ws = wb.Worksheets("FINANCE"): ws.Unprotect
    Set graph = bundle("Domain"): Finance_DomainValidate graph
    FinanceStore_Write wb, graph
    Finance_DomainViews graph, CLng(graph("META")(1)("SCOPE_ID")), rows, expenses, True
    Finance_PrepareImportedViews graph, rows, expenses
    Finance_Write Finance_Table(FIN_TABLE, wb), rows
    Finance_Write Finance_Table(FIN_EXPENSES, wb), expenses
    ws.Range("S1").Value2 = graph("META")(1)("CURRENCY")
    ws.Range("S2").Value2 = graph("META")(1)("REFRESH_AT")
    wb.Worksheets("SETTINGS").Range("X19").Value2 = bundle("Vocabulary")
    wb.Worksheets("SETTINGS").Range("X20").Value2 = bundle("Language")
    ' Migration runs only in its target workbook, never in the source.
    If wb Is ThisWorkbook Then
        Finance_BindSavedView
        Finance_Present
    End If
End Sub

Private Function MigrationData_FinanceToken(ByVal bundle As Object) As String
    MigrationData_FinanceToken = CStr(bundle("Vocabulary")) & "|" & CStr(bundle("Language")) & "|" & FinanceStore_Token(bundle("Domain"))
End Function

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




