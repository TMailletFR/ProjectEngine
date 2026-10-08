Attribute VB_Name = "mod_Migration"
Option Explicit

Private mRunning As Boolean
Private mLastDiagnostic As String
Private mLastResult As String

Public Function Migration_IsRunning() As Boolean
    Migration_IsRunning = mRunning
End Function

Public Function Migration_LastResult() As String
    Migration_LastResult = mLastResult
End Function

Public Function Migration_LastDiagnostic() As String
    Migration_LastDiagnostic = mLastDiagnostic
End Function

Public Function Migration_Text(ByVal key As String, Optional ByVal arguments As Object = Nothing) As String
    Migration_Text = Replace$(TextCatalog_Format(key, Settings_GetGlobalDisplayLanguage(), arguments), "\n", vbCrLf)
End Function

Public Sub Migration_Reject(ByVal key As String, Optional ByVal field As String = "")
    Err.Raise vbObjectError + 5721, "MigrationValidation", Migration_Text(key, TextCatalog_Arguments("Field", field))
End Sub

Public Sub Migration_ImportPrevious()
    Dim selected As String, targetName As String, targetPath As String
    Dim dialog As Object
    On Error GoTo Failed
    mLastResult = "CANCELLED": mLastDiagnostic = ""
    WorkbookSchema_RequireWritable
    If Application.ActiveWorkbook Is Nothing Then Exit Sub
    If Not (Application.ActiveWorkbook Is ThisWorkbook) Then Exit Sub
    targetName = ThisWorkbook.Name: targetPath = ThisWorkbook.FullName
    Set dialog = Application.FileDialog(msoFileDialogFilePicker)
    dialog.Title = Migration_Text("MIGRATION.SELECT")
    dialog.AllowMultiSelect = False
    dialog.Filters.Clear
    dialog.Filters.Add "Excel", "*.xlsm;*.xlsx"
    If dialog.Show <> -1 Then Exit Sub
    selected = dialog.SelectedItems(1)
    If ThisWorkbook.Name <> targetName Or ThisWorkbook.FullName <> targetPath Then Err.Raise 5, , "TARGET_IDENTITY_CHANGED"
    Migration_ImportFile selected
    Exit Sub
Failed:
    mLastResult = "FAILED"
    mLastDiagnostic = CStr(Err.Number) & ":" & Err.Source & ":" & Err.Description
    If Err.Source = "MigrationValidation" Then
        MsgBox Err.Description, vbExclamation, Migration_Text("MIGRATION.TITLE")
    Else
        MsgBox Migration_Text("MIGRATION.PROFILE"), vbExclamation, Migration_Text("MIGRATION.TITLE")
    End If
End Sub

' Public non-dialog entry used by the UI and isolated test tooling alike.
Public Function Migration_ImportFile(ByVal sourcePath As String, Optional ByVal confirmation As Boolean = False) As String
    Dim perfScope As clsPerfScope
    Dim source As Workbook, target As Workbook, item As Workbook
    Dim snapshot As Object, current As Object, receipt As Object
    Dim sourceName As String, fullSource As String, targetName As String, targetPath As String
    Dim operationId As String, phase As String, applied As String
    Dim oldEvents As Boolean, oldScreen As Boolean, oldAutoFill As Boolean, oldSecurity As Long
    Dim captured As Boolean, sourceOwned As Boolean, watchPaused As Boolean, mutationStarted As Boolean
    Dim errorNumber As Long, errorSource As String, errorDescription As String, finalSaved As Boolean
    On Error GoTo Failed
    mLastResult = "": mLastDiagnostic = ""
    If mRunning Then Err.Raise 5, , "IMPORT_REENTRANT"
    Set target = ThisWorkbook
    WorkbookSchema_RequireWritable
    If target.ReadOnly Or Len(target.Path) = 0 Or target.ProtectStructure Then Migration_Reject "MIGRATION.TARGET"
    targetName = target.Name: targetPath = target.FullName
    fullSource = CreateObject("Scripting.FileSystemObject").GetAbsolutePathName(sourcePath)
    sourceName = CreateObject("Scripting.FileSystemObject").GetFileName(fullSource)
    If StrComp(fullSource, targetPath, vbTextCompare) = 0 Then Migration_Reject "MIGRATION.PROFILE"
    For Each item In Application.Workbooks
        If StrComp(item.FullName, fullSource, vbTextCompare) = 0 Or StrComp(item.Name, sourceName, vbTextCompare) = 0 Then Migration_Reject "MIGRATION.OPEN_SOURCE"
    Next item
    If Not confirmation Then
        If MsgBox(Migration_Text("MIGRATION.DESTRUCTIVE"), vbOKCancel + vbExclamation + vbDefaultButton2, Migration_Text("MIGRATION.WARNING_TITLE")) <> vbOK Then
            mLastResult = "CANCELLED": GoTo CleanExit
        End If
    End If
    mRunning = True
    phase = "SOURCE_SAFETY"
    Migration_PreflightPackage fullSource
    oldEvents = Application.EnableEvents: oldScreen = Application.ScreenUpdating
    oldSecurity = Application.AutomationSecurity
    oldAutoFill = Application.AutoCorrect.AutoFillFormulasInLists
    captured = True
    Application.EnableEvents = False
    Application.AutomationSecurity = msoAutomationSecurityForceDisable
    phase = "READ_SOURCE"
    Set perfScope = Profiler_BeginScope("Migration.Source.Open", "Migration")
    Set source = Application.Workbooks.Open(Filename:=fullSource, UpdateLinks:=0, ReadOnly:=True, Password:="PE_IMPORT_NO_PASSWORD", AddToMru:=False)
    Set perfScope = Nothing
    Profiler_RecordCounter "Migration.Source.Opens", 1
    sourceOwned = True
    Application.AutomationSecurity = oldSecurity
    Migration_AssertIdentity source, fullSource, sourceName
    If Not source.ReadOnly Or source.Connections.Count <> 0 Then Migration_Reject "MIGRATION.UNSAFE"
    Dim links As Variant
    links = source.LinkSources(xlExcelLinks)
    If Not IsEmpty(links) Then Migration_Reject "MIGRATION.UNSAFE"
    Set perfScope = Profiler_BeginScope("Migration.Source.Read", "Migration")
    Set snapshot = MigrationData_Read(source, target)
    Set perfScope = Nothing
    Migration_AssertIdentity source, fullSource, sourceName
    Set perfScope = Profiler_BeginScope("Migration.Source.Close", "Migration")
    source.Close False: sourceOwned = False: Set source = Nothing
    Set perfScope = Nothing
    Profiler_RecordCounter "Migration.Source.Closes", 1

    Migration_AssertIdentity target, targetPath, targetName
    If target.ReadOnly Then Err.Raise 5, , "TARGET_READONLY"
    GanttDrag_PauseForLifecycle: watchPaused = True
    Application.ScreenUpdating = False
    Application.AutoCorrect.AutoFillFormulasInLists = False
    operationId = Format$(Now, "yyyymmdd-hhnnss") & "-" & Replace$(Trim$(Str$(Timer * 100)), ".", "-")
    snapshot.Add "OperationId", operationId
    snapshot.Add "RecoveryBackup", ""
    phase = "APPLY"
    mutationStarted = True
    target.Activate
    Set perfScope = Profiler_BeginScope("Migration.Target.Apply", "Migration")
    applied = Migration_ApplyInputs(snapshot)
    Set perfScope = Nothing
    If applied <> "PASS" Then Err.Raise 5, "MigrationApply", applied
    phase = "VERIFY"
    Set perfScope = Profiler_BeginScope("Migration.Target.Verify", "Migration")
    Set current = MigrationData_Read(target, target)
    MigrationData_VerifyPreserved snapshot, current
    Migration_VerifyEmptyOutputs
    Set perfScope = Nothing
    Migration_Log "INFO", "MIGRATION.SUCCESS", operationId, "", CStr(snapshot("Profile")), snapshot, receipt
    WorkbookSchema_SetProperty target, "ProjectEngine.ImportOperation", operationId
    WorkbookSchema_SetProperty target, "ProjectEngine.ImportSource", fullSource
    WorkbookSchema_SetProperty target, "ProjectEngine.ImportDigest", CStr(snapshot("Digest"))
    WorkbookSchema_SetProperty target, "ProjectEngine.ImportProfile", CStr(snapshot("Profile"))
    WorkbookSchema_SetProperty target, "ProjectEngine.ImportBackup", ""
    WorkbookSchema_SetProperty target, "ProjectEngine.ImportTimestamp", Format$(Now, "yyyy-mm-dd hh:nn:ss")
    WorkbookSchema_SetProperty target, "ProjectEngine.ImportState", "SUCCESS"
    phase = "ACK"
    Set perfScope = Profiler_BeginScope("Migration.Target.Ack", "Migration")
    ProjectWelcome_Acknowledge
    Set perfScope = Nothing
    phase = "FINAL_SAVE"
    Set perfScope = Profiler_BeginScope("Migration.Target.Save", "Migration")
    Profiler_RecordCounter "Migration.Target.Saves", 1
    target.Save
    Set perfScope = Nothing
    finalSaved = True
    phase = "POST_SAVE"
    Set perfScope = Profiler_BeginScope("Migration.Target.PostSave", "Migration")
    Set current = MigrationData_Read(target, target)
    MigrationData_VerifyPreserved snapshot, current, True
    If Not PlanningMessage_IsAcknowledged(ProjectWelcome_Message(), False) Then Err.Raise 5, "Migration", "WELCOME_ACK_MISSING_AFTER_SAVE"
    Set perfScope = Nothing
    mLastResult = "SUCCESS"
    GoTo CleanExit
Failed:
    errorNumber = Err.Number: errorSource = Err.Source: errorDescription = Err.Description
    mLastDiagnostic = phase & ":" & CStr(errorNumber) & ":" & errorSource & ":" & errorDescription
    mLastResult = "FAILED"
    Set perfScope = Nothing
    ' Failure never intentionally saves a partially imported target.
    On Error Resume Next
    If mutationStarted Then
        If Not finalSaved Then SetPlanningWarningAckState ProjectWelcome_Message(), False
        WorkbookSchema_SetProperty target, "ProjectEngine.ImportState", "FAILED"
        WorkbookSchema_SetProperty target, "ProjectEngine.ImportDiagnostic", Left$(mLastDiagnostic, 240)
        Err.Clear
        Migration_Log "STOP", IIf(finalSaved, "MIGRATION.POST_SAVE_FAILED", "MIGRATION.FAILED"), operationId, "", CStr(snapshot("Profile"))
        If Err.Number <> 0 Then mLastDiagnostic = mLastDiagnostic & "|FAILURE_RECEIPT:" & CStr(Err.Number) & ":" & Err.Description
    End If
    On Error GoTo 0
CleanExit:
    On Error GoTo CleanupFailed
    If sourceOwned Then
        Migration_AssertIdentity source, fullSource, sourceName
        source.Close False: sourceOwned = False
    End If
    If captured Then
        Application.AutomationSecurity = oldSecurity
        Application.AutoCorrect.AutoFillFormulasInLists = oldAutoFill
        Application.EnableEvents = oldEvents
        Application.ScreenUpdating = oldScreen
    End If
    If mLastResult = "SUCCESS" Then target.Worksheets("WBS").Activate
    If watchPaused Then
        If WorkbookSchema_WritesAllowed() Then GanttDrag_EnsureArmed
    End If
    mRunning = False
    Migration_ImportFile = mLastResult
    If Not confirmation Then
        If mLastResult = "SUCCESS" Then
            MsgBox Migration_Text("MIGRATION.SUCCESS"), vbInformation, Migration_Text("MIGRATION.TITLE")
        ElseIf mLastResult = "FAILED" Then
            If Not mutationStarted And errorSource = "MigrationValidation" Then
                MsgBox errorDescription, vbExclamation, Migration_Text("MIGRATION.TITLE")
            Else
                MsgBox Migration_Text(IIf(finalSaved, "MIGRATION.POST_SAVE_FAILED", "MIGRATION.FAILED")), vbExclamation, Migration_Text("MIGRATION.TITLE")
            End If
        End If
    End If
    Exit Function
CleanupFailed:
    mLastDiagnostic = mLastDiagnostic & "|CLEANUP:" & CStr(Err.Number) & ":" & Err.Description
    On Error Resume Next
    If captured Then
        Application.AutomationSecurity = oldSecurity
        Application.AutoCorrect.AutoFillFormulasInLists = oldAutoFill
        Application.EnableEvents = oldEvents
        Application.ScreenUpdating = oldScreen
    End If
    mRunning = False: mLastResult = "FAILED": Migration_ImportFile = "FAILED"
    If Not confirmation Then MsgBox Migration_Text("MIGRATION.FAILED"), vbExclamation, Migration_Text("MIGRATION.TITLE")
End Function

Private Sub Migration_AssertIdentity(ByVal wb As Workbook, ByVal fullName As String, ByVal name As String)
    If wb Is Nothing Then Err.Raise 5, , "WORKBOOK_OBJECT_MISSING"
    If StrComp(wb.FullName, fullName, vbTextCompare) <> 0 Or wb.Name <> name Then Err.Raise 5, , "WORKBOOK_IDENTITY_CHANGED"
End Sub

Public Function Migration_ApplyInputs(ByVal snapshot As Object) As String
    Dim previousRunning As Boolean, tbl As ListObject
    Dim key As Variant, n As Long, errorNumber As Long, description As String
    Dim oldEvents As Boolean, eventsCaptured As Boolean
    On Error GoTo Failed
    previousRunning = mRunning: mRunning = True
    If ThisWorkbook.ReadOnly Then Err.Raise 5, , "TARGET_READONLY"
    oldEvents = Application.EnableEvents: eventsCaptured = True
    Application.EnableEvents = False
    GanttDrag_PauseForLifecycle
    Set tbl = MigrationData_Table(ThisWorkbook, "tbl_WBS")
    If Not tbl.DataBodyRange Is Nothing Then tbl.DataBodyRange.Delete
    Planning_FullSafeEmptyState
    Reset_Dashboard
    Migration_VerifyEmptyOutputs
    Dim options As Object
    Set options = snapshot("Settings")
    Settings_ApplyMigrationSnapshot options("Values")
    MigrationData_Write snapshot, ThisWorkbook
    WBS_ApplyLanguage Settings_GetOwnerLanguage("WBS")
    Dashboard_RebindSnapshotSelectors
    Set tbl = MigrationData_Table(ThisWorkbook, "tbl_WBS")
    RestoreWBSFormulaColumns tbl
    If Not tbl.DataBodyRange Is Nothing Then
        Dim restored As Variant, expected As String, actual As String
        For Each key In Array(VTS_COL_BASELINE_FINISH, VTS_COL_ACTUAL_DURATION, VTS_COL_CALCULATED_DURATION)
            expected = WBSFormulaWriter_ExpectedFormula(CStr(key))
            restored = SchemaListColumn(tbl, VTS_TABLE_WBS, CStr(key)).DataBodyRange.Formula
            For n = 1 To tbl.ListRows.Count
                If tbl.ListRows.Count = 1 Then actual = CStr(restored) Else actual = CStr(restored(n, 1))
                actual = Replace$(actual, tbl.Name & "[[#This Row],[", "[@[")
                If actual <> expected Then Err.Raise 5, "Migration", "STANDARD_FORMULA_NOT_RESTORED:" & CStr(key)
            Next n
        Next key
    End If
    Import_WBS_To_Constraints
    If MigrationData_Table(ThisWorkbook, "tbl_CONSTRAINTS").ListRows.Count <> tbl.ListRows.Count Then
        Err.Raise 5, "Migration", "CONSTRAINTS_INPUT_PROJECTION_INCOMPLETE"
    End If
    Migration_Log "INFO", "MIGRATION.STARTED", CStr(snapshot("OperationId")), CStr(snapshot("RecoveryBackup")), CStr(snapshot("Profile")), snapshot
    Migration_VerifyEmptyOutputs
    Refresh_EventHistory_View True, True
    Settings_Initialize
    GanttDrag_PauseForLifecycle
    If eventsCaptured Then Application.EnableEvents = oldEvents
    Migration_ApplyInputs = "PASS"
    mRunning = previousRunning
    Exit Function
Failed:
    Migration_ApplyInputs = "FAIL:" & CStr(Err.Number) & ":" & Err.Source & ":" & Err.Description
    GanttDrag_PauseForLifecycle
    If eventsCaptured Then Application.EnableEvents = oldEvents
    mRunning = previousRunning
End Function

Private Sub Migration_VerifyEmptyOutputs()
    Dim name As Variant, tbl As ListObject, shape As Shape
    For Each name In Array("tbl_CALC", "tbl_LOGIC_LINKS", "tbl_SCURVE", "tbl_CALC_STATE", "tbl_CALC_GANTT_TEST", "tbl_CALC_SCURVE")
        Set tbl = MigrationData_Table(ThisWorkbook, CStr(name))
        If tbl.ListRows.Count <> 0 Or Not tbl.DataBodyRange Is Nothing Then Err.Raise 5, "Migration", "EMPTY_OUTPUT_NOT_CLEARED:" & CStr(name)
    Next name
    For Each shape In ThisWorkbook.Worksheets("GANTT").Shapes
        If GanttShapeRegistry_IsBusinessShapeName(shape.Name) Or GanttDependencySvg_IsLayerName(shape.Name) Then
            Err.Raise 5, "Migration", "EMPTY_GANTT_SHAPE_NOT_CLEARED:" & shape.Name
        End If
    Next shape
End Sub


Private Sub Migration_Log(ByVal severity As String, ByVal key As String, ByVal operation As String, ByVal backup As String, ByVal profile As String, Optional ByVal snapshot As Object = Nothing, Optional ByRef emittedReceipt As Object = Nothing)
    Dim arguments As Object, receipt As Object, hash As String
    Set arguments = TextCatalog_Arguments("Backup", backup, "Operation", operation, "Profile", profile, "MessageKey", key)
    If Not snapshot Is Nothing Then
        arguments.Add "Source", CStr(snapshot("Source"))
        arguments.Add "SourceDigest", CStr(snapshot("Digest"))
        arguments.Add "HistoryImported", CBool(snapshot("Tables")("tbl_CALC_ALARM")("Available"))
        If snapshot("Tables")("tbl_CALC_ALARM").Exists("UnavailableDiagnostic") Then arguments.Add "HistoryDiagnostic", CStr(snapshot("Tables")("tbl_CALC_ALARM")("UnavailableDiagnostic"))
        Dim name As Variant
        For Each name In Array("tbl_WBS", "tbl_CONSTRAINTS", "tbl_EVENT_ACK", "tbl_DASHBOARD_SNAPSHOTS", "tbl_CALC_ALARM")
            arguments.Add CStr(name) & ".Count", CLng(snapshot("Tables")(name)("Count"))
        Next name
    End If
    hash = BuildPlanningEventIdentityV2(severity, "MIGRATION", "WORKBOOK", operation, arguments)
    Set receipt = LogPlanningEvent(severity, "MIGRATION", hash, TextCatalog_Format(key, "FR", arguments), _
        TextCatalog_Format(key, "EN", arguments), mLastDiagnostic, mLastDiagnostic, "Migration_ImportFile", "SETTINGS", "", "", "", "", True)
    If receipt Is Nothing Then Err.Raise 5, , "MIGRATION_LOG_RECEIPT_MISSING"
    If Len(CStr(receipt("EventId"))) = 0 Or CStr(receipt("Hash")) <> hash Then Err.Raise 5, , "MIGRATION_LOG_RECEIPT_INVALID"
    Set emittedReceipt = receipt
    WorkbookSchema_SetProperty ThisWorkbook, "ProjectEngine.ImportEventId", CStr(receipt("EventId"))
    WorkbookSchema_SetProperty ThisWorkbook, "ProjectEngine.ImportEventHash", CStr(receipt("Hash"))
End Sub

' Inspect ZIP directory names before Excel opens the source (including XLM sheets).
Private Sub Migration_PreflightPackage(ByVal path As String)
    Dim fileNo As Integer, bytes() As Byte, length As Long, p As Long, eocd As Long
    Dim offset As Long, count As Long, i As Long, nameLength As Long, extra As Long, comment As Long
    Dim name As String, j As Long, foundWorkbook As Boolean
    On Error GoTo Failed
    If LCase$(Right$(path, 5)) <> ".xlsm" And LCase$(Right$(path, 5)) <> ".xlsx" Then Migration_Reject "MIGRATION.PROFILE"
    fileNo = FreeFile
    Open path For Binary Access Read Lock Write As #fileNo
    length = LOF(fileNo)
    If length < 22 Or length > 268435456 Then Err.Raise 5, , "UNSUPPORTED_PACKAGE_SIZE"
    ReDim bytes(0 To length - 1)
    Get #fileNo, , bytes
    Close #fileNo: fileNo = 0
    For p = length - 22 To Application.Max(0, length - 65557) Step -1
        If bytes(p) = &H50 And bytes(p + 1) = &H4B And bytes(p + 2) = &H5 And bytes(p + 3) = &H6 Then eocd = p: Exit For
    Next p
    If eocd = 0 Then Err.Raise 5, , "ZIP_DIRECTORY_MISSING"
    count = Migration_U16(bytes, eocd + 10)
    If count = 65535 Then Err.Raise 5, , "ZIP64_NOT_SUPPORTED"
    offset = Migration_U32(bytes, eocd + 16)
    For i = 1 To count
        If offset + 46 > length Then Err.Raise 5, , "CORRUPT_ZIP_DIRECTORY"
        If bytes(offset) <> &H50 Or bytes(offset + 1) <> &H4B Or bytes(offset + 2) <> 1 Or bytes(offset + 3) <> 2 Then Err.Raise 5, , "CORRUPT_ZIP_DIRECTORY"
        nameLength = Migration_U16(bytes, offset + 28)
        extra = Migration_U16(bytes, offset + 30): comment = Migration_U16(bytes, offset + 32)
        If offset + 46 + nameLength + extra + comment > length Then Err.Raise 5, , "CORRUPT_ZIP_DIRECTORY"
        name = ""
        For j = 0 To nameLength - 1
            name = name & Chr$(bytes(offset + 46 + j))
        Next j
        name = LCase$(Replace$(name, "\", "/"))
        If InStr(name, "macrosheets/") > 0 Or InStr(name, "externalLinks/") > 0 Or name = "xl/connections.xml" Then Migration_Reject "MIGRATION.UNSAFE"
        If InStr(name, "externallinks/") > 0 Then Migration_Reject "MIGRATION.UNSAFE"
        If name = "xl/workbook.xml" Then foundWorkbook = True
        offset = offset + 46 + nameLength + extra + comment
    Next i
    If Not foundWorkbook Then Migration_Reject "MIGRATION.PROFILE"
    Exit Sub
Failed:
    Dim errorNumber As Long, description As String, source As String
    errorNumber = Err.Number: description = Err.Description: source = Err.Source
    If fileNo <> 0 Then Close #fileNo
    Err.Raise errorNumber, source, description
End Sub

Private Function Migration_U16(ByRef bytes() As Byte, ByVal offset As Long) As Long
    Migration_U16 = CLng(bytes(offset)) + CLng(bytes(offset + 1)) * 256&
End Function

Private Function Migration_U32(ByRef bytes() As Byte, ByVal offset As Long) As Long
    Dim value As Double
    value = CDbl(bytes(offset)) + CDbl(bytes(offset + 1)) * 256# + CDbl(bytes(offset + 2)) * 65536# + CDbl(bytes(offset + 3)) * 16777216#
    If value > 2147483647# Then Err.Raise 5, , "ZIP_OFFSET_NOT_SUPPORTED"
    Migration_U32 = CLng(value)
End Function
