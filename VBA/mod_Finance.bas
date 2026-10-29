Attribute VB_Name = "mod_Finance"
Option Explicit

' Finance owns its inputs, explicit calculation and outputs. No planning caller.
Public Const FIN_TABLE As String = "tbl_FINANCE"
Public Const FIN_EXPENSES As String = "tbl_FINANCE_EXPENSES"
Private Const FIN_SUMMARY As String = "FINANCE SUMMARY"
Private Const FIN_TODAY_LINE As String = "FINANCE_TODAY_LINE"
Private Const FIN_NATIVE_EXPENSE_TRIAL As Boolean = True
Public Const FIN_EXTRACT_COUNTER As String = "AZ1"
Private mCalculations As Long, mReads As Long, mBusy As Boolean
Private mViewBound As Boolean, mViewScope As Long, mViewRevision As Long
Private mTransaction As Boolean, mStorageFailed As Boolean, mDraftDepth As Long
Private mProjectionDomain As Object
Private mProjectionJournal As Collection
Private mProjectionOutputsChanged As Boolean
Private mExpenseSheet As Worksheet, mExpenseTable As ListObject
Private mExpenseRows As Collection, mExpenseOriginal As Collection
Private mExpenseOriginalIndex As Object
Private mExpenseHeaders As Variant
Private mExpenseScope As Long, mExpenseRevision As Long, mExpenseErrorRow As Long
Private mExpenseActive As Boolean, mExpenseEditing As Boolean
Private mExpenseCommitted As Boolean

Public Function Finance_InputKeys() As Variant
    Finance_InputKeys = Array("BUDGET", "PROFILE", "ETC")
End Function

Public Function Finance_DomainIndex(ByVal rows As Collection, ByVal key As String) As Object
    Dim result As Object, row As PEFinanceRow, value As String
    Set result = CreateObject("Scripting.Dictionary")
    For Each row In rows
        value = CStr(row(key))
        If result.exists(value) Then Err.Raise 5, "Finance", Finance_Text("INVALID_DOMAIN") & " " & key & "=" & value
        result.Add value, row
    Next row
    Set Finance_DomainIndex = result
End Function

Private Function Finance_Id(ByVal value As Variant, Optional ByVal allowZero As Boolean = False) As Long
    If isError(value) Or IsEmpty(value) Then GoTo invalid
    If Not IsNumeric(value) Then GoTo invalid
    If CDec(value) <> Fix(CDec(value)) Or CDec(value) > 2147483646# Then GoTo invalid
    If CDec(value) < IIf(allowZero, 0, 1) Then GoTo invalid
    Finance_Id = CLng(value): Exit Function
invalid:
    Err.Raise 5, "Finance", Finance_Text("INVALID_DOMAIN")
End Function

Private Function Finance_ValidWBS(ByVal value As String) As Boolean
    If Len(value) = 0 Then Exit Function
    If value Like "*[!0-9.]*" Then Exit Function
    Finance_ValidWBS = (NormalizeWBS(value) = value)
End Function

Public Sub Finance_DomainValidate(ByVal graph As Object)
    Dim projects As Object, records As Object, links As Object, journal As Object, names As Object
    Dim references As Object, groups As Object, group As Variant, uniqueLinks As Object, pairs As Object, pair As String
    Dim row As PEFinanceRow, meta As PEFinanceRow, key As Variant, id As Long, f As String, wbs As String, name As String
    Dim maxProject As Long, maxLink As Long, maxFinancial As Long, maxExpense As Long
    Dim perf As clsPerfScope
    Set perf = Profiler_BeginScope("Finance.Storage.Validate", "Finance")
    Set projects = Finance_DomainIndex(graph("PROJECTS"), "PROJECT_ID")
    Set records = Finance_DomainIndex(graph("RECORDS"), "FINANCIAL_ID")
    Set links = Finance_DomainIndex(graph("LINKS"), "LINK_ID")
    Set journal = Finance_DomainIndex(graph("JOURNAL"), "EXPENSE_ID")
    If projects.count = 0 Or graph("META").count <> 1 Then GoTo invalid
    Set names = CreateObject("Scripting.Dictionary"): names.CompareMode = vbTextCompare
    For Each row In graph("PROJECTS")
        id = Finance_Id(row("PROJECT_ID")): If id > maxProject Then maxProject = id
        name = Trim$(CStr(row("NAME")))
        If Len(name) = 0 Or Len(name) > 120 Or names.exists(name) Then GoTo invalid
        names.Add name, True
        id = Finance_Id(row("DISPLAY_ORDER"))
    Next row
    For Each row In graph("RECORDS")
        id = Finance_Id(row("FINANCIAL_ID")): If id > maxFinancial Then maxFinancial = id
        row("BUDGET") = Finance_Money(row("BUDGET")): row("ETC") = Finance_Money(row("ETC"))
        row("PROFILE") = Finance_Profile(CStr(row("PROFILE")))
        Select Case CStr(row("PROFILE")): Case "START", "END", "LINEAR": Case Else: GoTo invalid: End Select
    Next row
    Set references = CreateObject("Scripting.Dictionary"): Set groups = CreateObject("Scripting.Dictionary"): Set pairs = CreateObject("Scripting.Dictionary")
    Set uniqueLinks = CreateObject("Scripting.Dictionary")
    For Each row In graph("LINKS")
        id = Finance_Id(row("LINK_ID")): If id > maxLink Then maxLink = id
        If Not projects.exists(CStr(Finance_Id(row("PROJECT_ID")))) Then GoTo invalid
        f = CStr(Finance_Id(row("FINANCIAL_ID"))): If Not records.exists(f) Then GoTo invalid
        wbs = CStr(row("WBS")): If Not Finance_ValidWBS(wbs) Then GoTo invalid
        If references.exists(f) Then
            If CStr(references(f)) <> wbs Then GoTo invalid
        Else
            references.Add f, wbs
        End If
        If row("STATE") <> "ACTIVE" And row("STATE") <> "HISTORICAL" Then GoTo invalid
        key = CStr(row("PROJECT_ID")) & "|" & Finance_Reference(row)
        If uniqueLinks.exists(key) Then GoTo invalid
        uniqueLinks.Add key, True
        id = Finance_Id(row("ROW_ORDER"))
        key = "ORDER|" & CStr(row("PROJECT_ID")) & "|" & CStr(id)
        If uniqueLinks.exists(key) Then GoTo invalid
        uniqueLinks.Add key, True
        If Finance_Has(row("TASK_ID")) Then id = Finance_Id(row("TASK_ID"))
        If Finance_Has(row("START")) Then
            If Not Finance_UsableDate(row("START")) Then GoTo invalid
        End If
        If Finance_Has(row("FINISH")) Then
            If Not Finance_UsableDate(row("FINISH")) Then GoTo invalid
        End If
        If row("STATE") = "ACTIVE" Then
            If Not groups.exists(wbs) Then
                groups.Add wbs, Array(0&, 0&)
            End If
            group = groups(wbs): group(0) = CLng(group(0)) + 1
            pair = wbs & "|" & f
            If Not pairs.exists(pair) Then pairs.Add pair, True: group(1) = CLng(group(1)) + 1
            groups(wbs) = group
        End If
    Next row
    For Each key In groups.keys
        group = groups(key)
        If CLng(group(1)) <> 1 And CLng(group(1)) <> CLng(group(0)) Then GoTo invalid
    Next key
    For Each key In records.keys: If Not references.exists(CStr(key)) Then GoTo invalid
    Next key
    For Each row In graph("JOURNAL")
        id = Finance_Id(row("EXPENSE_ID")): If id > maxExpense Then maxExpense = id
        If Not records.exists(CStr(Finance_Id(row("FINANCIAL_ID")))) Then GoTo invalid
        If Not Finance_UsableDate(row("DATE")) Then GoTo invalid
        If CDbl(row("DATE")) <> Fix(CDbl(row("DATE"))) Then GoTo invalid
        row("AMOUNT") = Finance_Money(row("AMOUNT"), True)
        If Not Finance_Has(row("AMOUNT")) Then GoTo invalid
        If isError(row("COMMENT")) Then GoTo invalid
    Next row
    Set meta = graph("META")(1)
    id = Finance_Id(meta("SCOPE_ID"), True)
    If id = 0 Then
        If projects.count < 2 Then GoTo invalid
    ElseIf Not projects.exists(CStr(id)) Then
        GoTo invalid
    End If
    id = Finance_Id(meta("REVISION"), True)
    If Finance_Id(meta("NEXT_PROJECT")) <= maxProject Or Finance_Id(meta("NEXT_LINK")) <= maxLink Then GoTo invalid
    If Finance_Id(meta("NEXT_FINANCIAL")) <= maxFinancial Or Finance_Id(meta("NEXT_EXPENSE")) <= maxExpense Then GoTo invalid
    If Len(CStr(meta("CURRENCY"))) <> 3 Or CStr(meta("CURRENCY")) Like "*[!A-Z]*" Then GoTo invalid
    If Finance_Has(meta("REFRESH_AT")) Then
        If Not IsNumeric(meta("REFRESH_AT")) Then GoTo invalid
        If Not Finance_UsableDate(Fix(CDbl(meta("REFRESH_AT")))) Then GoTo invalid
    End If
    Exit Sub
invalid:
    Err.Raise 5, "Finance", Finance_Text("INVALID_DOMAIN")
End Sub

Private Function Finance_AllocateId(ByVal graph As Object, ByVal field As String) As Long
    Dim meta As Object
    Set meta = graph("META")(1)
    Finance_AllocateId = Finance_Id(meta(field))
    If Finance_AllocateId >= 2147483646 Then Err.Raise 5, "Finance", Finance_Text("INVALID_DOMAIN")
    meta(field) = Finance_AllocateId + 1
End Function

Public Function Finance_DomainBlank(ByVal currencyCode As String) As Object
    Dim graph As Object, project As Object, meta As Object
    Set graph = FinanceStore_Empty()
    Set project = FinanceStore_Row("PROJECTS")
    project("PROJECT_ID") = 1&: project("NAME") = Finance_Text("DEFAULT_PROJECT"): project("DISPLAY_ORDER") = 1&
    graph("PROJECTS").Add project
    Set meta = FinanceStore_Row("META")
    meta("SCOPE_ID") = 1&: meta("CURRENCY") = currencyCode: meta("REVISION") = 0&
    meta("NEXT_PROJECT") = 2&: meta("NEXT_LINK") = 1&: meta("NEXT_FINANCIAL") = 1&: meta("NEXT_EXPENSE") = 1&
    graph("META").Add meta
    Set Finance_DomainBlank = graph
End Function

Public Function Finance_DomainFromLegacy(ByVal rows As Collection, ByVal expenses As Collection, ByVal currencyCode As String) As Object
    Dim graph As Object, indexes As Object, row As Object, link As Object, record As Object, entry As Object, key As Variant, n As Long
    Finance_Validate rows, expenses
    Set graph = Finance_DomainBlank(currencyCode): Set indexes = CreateObject("Scripting.Dictionary")
    For Each row In rows
        n = n + 1: Set link = FinanceStore_Row("LINKS"): Set record = FinanceStore_Row("RECORDS")
        record("FINANCIAL_ID") = Finance_AllocateId(graph, "NEXT_FINANCIAL")
        For Each key In Finance_InputKeys(): record(CStr(key)) = row(CStr(key)): Next key
        link("LINK_ID") = Finance_AllocateId(graph, "NEXT_LINK"): link("PROJECT_ID") = 1&
        link("FINANCIAL_ID") = record("FINANCIAL_ID"): link("ROW_ORDER") = n
        For Each key In Array("WBS", "STATE", "TASK_ID", "TASK_NAME", "PROJECT", "START", "FINISH"): link(CStr(key)) = row(CStr(key)): Next key
        graph("LINKS").Add link: graph("RECORDS").Add record
        indexes.Add Finance_Reference(row), record("FINANCIAL_ID")
    Next row
    For Each row In expenses
        Set entry = FinanceStore_Row("JOURNAL")
        entry("EXPENSE_ID") = Finance_AllocateId(graph, "NEXT_EXPENSE")
        entry("FINANCIAL_ID") = indexes(Finance_Reference(row))
        For Each key In Array("DATE", "AMOUNT", "COMMENT"): entry(CStr(key)) = row(CStr(key)): Next key
        graph("JOURNAL").Add entry
    Next row
    Finance_DomainValidate graph
    Set Finance_DomainFromLegacy = graph
End Function

Public Function Finance_DomainSnapshot(Optional ByVal wb As Workbook) As Object
    Dim graph As Object, perf As clsPerfScope
    Set perf = Profiler_BeginScope("Finance.Storage.Read", "Finance")
    If wb Is Nothing Then Set wb = ThisWorkbook
    Set graph = FinanceStore_Read(wb): Finance_DomainValidate graph
    Set Finance_DomainSnapshot = graph
End Function

Public Sub Finance_UpgradeStorage()
    Dim graph As Object
    If FinanceStore_Exists(ThisWorkbook) Then Err.Raise 5, "Finance", "Storage already initialized"
    Set graph = Finance_DomainFromLegacy(Finance_Read(Finance_Table(FIN_TABLE)), Finance_Read(Finance_Table(FIN_EXPENSES)), CStr(ThisWorkbook.Worksheets("FINANCE").Range("S1").Value2))
    graph("META")(1)("REFRESH_AT") = ThisWorkbook.Worksheets("FINANCE").Range("S2").Value2
    Finance_DomainValidate graph
    FinanceStore_Write ThisWorkbook, graph
    Finance_BindSavedView
End Sub

Public Sub Finance_BindSavedView()
    Dim meta As Object, graph As Object
    If Not FinanceStore_Exists(ThisWorkbook) Then Exit Sub
    Set graph = FinanceStore_Read(ThisWorkbook): Finance_DomainValidate graph
    Set meta = graph("META")(1)
    mViewScope = CLng(meta("SCOPE_ID")): mViewRevision = CLng(meta("REVISION"))
    mViewBound = True
End Sub

Public Function Finance_ScopeId() As Long
    If mTransaction And Not mProjectionDomain Is Nothing Then Finance_ScopeId = CLng(mProjectionDomain("META")(1)("SCOPE_ID")): Exit Function
    If Not mViewBound Then Finance_BindSavedView
    If Not mViewBound Or mStorageFailed Then Err.Raise 5, "Finance", Finance_Text("INVALID_DOMAIN")
    Finance_ScopeId = mViewScope
End Function

Public Sub Finance_RequireIndividual()
    If Finance_ScopeId() = 0 Then Err.Raise 5, "Finance", Finance_Text("GLOBAL_READONLY")
End Sub

Public Sub Finance_Draft(ByVal entering As Boolean)
    If entering Then mDraftDepth = mDraftDepth + 1 Else mDraftDepth = mDraftDepth - 1
    If mDraftDepth < 0 Then mDraftDepth = 0
End Sub

Public Function Finance_ActivityDraft(ByRef inputsValid As Boolean) As Object
    Dim graph As Object
    Set graph = Finance_DomainSnapshot(): Finance_AssertBinding graph
    On Error GoTo InvalidInputs
    Finance_ReadVisibleInputs graph
    inputsValid = True: Set Finance_ActivityDraft = graph
    Exit Function
InvalidInputs:
    inputsValid = False: Set Finance_ActivityDraft = Finance_DomainSnapshot()
End Function

Public Sub Finance_StageActivities(ByVal graph As Object, ByVal projectId As Long, ByVal rows As Collection, ByVal decisions As Object)
    If projectId = 0 Then Err.Raise 5, "Finance", Finance_Text("GLOBAL_READONLY")
    Finance_ApplyMemberships graph, projectId, rows, decisions
End Sub

Public Sub Finance_CommitProjectDraft(ByVal graph As Object, ByVal scope As Long)
    Finance_CheckRevision Finance_ScopeId(), CLng(graph("META")(1)("REVISION"))
    Finance_CheckDraftInputs graph
    If graph("PROJECTS").count = 1 Then scope = CLng(graph("PROJECTS")(1)("PROJECT_ID"))
    graph("META")(1)("SCOPE_ID") = scope
    Finance_DomainValidate graph
    Finance_CommitDomain graph, IIf(scope = mViewScope And scope <> 0, 1, 2)
End Sub

Public Function Finance_PresentationDomain() As Object
    If mTransaction And Not mProjectionDomain Is Nothing Then
        Set Finance_PresentationDomain = mProjectionDomain
    Else
        Set Finance_PresentationDomain = Finance_DomainSnapshot()
    End If
End Function

Public Function Finance_ProjectOptions(ByVal graph As Object, Optional ByVal includeGlobal As Boolean = True) As Variant
    Dim values() As Variant, row As Object, i As Long, count As Long
    count = graph("PROJECTS").count
    If includeGlobal Then count = count + 1
    ReDim values(0 To count - 1, 0 To 1)
    If includeGlobal Then values(0, 0) = 0&: values(0, 1) = Finance_Text("GLOBAL"): i = 1
    For Each row In graph("PROJECTS")
        values(i, 0) = row("PROJECT_ID"): values(i, 1) = row("NAME")
        If StrComp(CStr(row("NAME")), Finance_Text("GLOBAL"), vbTextCompare) = 0 Then values(i, 1) = Finance_Text("PROJECT_LABEL") & ": " & CStr(row("NAME"))
        i = i + 1
    Next row
    Finance_ProjectOptions = values
End Function

Public Function Finance_ProjectAdd(ByVal graph As Object, ByVal name As String) As Long
    Dim row As Object
    Finance_CheckProjectName graph, name
    Set row = FinanceStore_Row("PROJECTS")
    row("PROJECT_ID") = Finance_AllocateId(graph, "NEXT_PROJECT")
    row("NAME") = Trim$(name): row("DISPLAY_ORDER") = row("PROJECT_ID")
    graph("PROJECTS").Add row
    On Error GoTo Failed
    Finance_DomainValidate graph
    Finance_ProjectAdd = CLng(row("PROJECT_ID"))
    Exit Function
Failed:
    Dim n As Long, description As String
    n = Err.number: description = Err.description
    graph("PROJECTS").Remove graph("PROJECTS").count
    graph("META")(1)("NEXT_PROJECT") = row("PROJECT_ID")
    Err.Raise n, "Finance", description
End Function

Public Sub Finance_ProjectRename(ByVal graph As Object, ByVal projectId As Long, ByVal name As String)
    Dim row As Object, original As String, n As Long, description As String
    Finance_CheckProjectName graph, name, projectId
    Set row = Finance_DomainIndex(graph("PROJECTS"), "PROJECT_ID")(CStr(projectId))
    original = CStr(row("NAME")): row("NAME") = Trim$(name)
    On Error GoTo Failed
    Finance_DomainValidate graph
    Exit Sub
Failed:
    n = Err.number: description = Err.description: row("NAME") = original
    Err.Raise n, "Finance", description
End Sub

Private Sub Finance_CheckProjectName(ByVal graph As Object, ByVal name As String, Optional ByVal excludingId As Long = 0)
    Dim row As Object, normalized As String
    normalized = Trim$(name)
    If Len(normalized) = 0 Or Len(normalized) > 120 Then Err.Raise 5, "Finance", Finance_Text("PROJECT_NAME_INVALID")
    For Each row In graph("PROJECTS")
        If CLng(row("PROJECT_ID")) <> excludingId Then
            If StrComp(Trim$(CStr(row("NAME"))), normalized, vbTextCompare) = 0 Then Err.Raise 5, "Finance", Finance_Text("PROJECT_NAME_INVALID")
        End If
    Next row
End Sub

Public Sub Finance_ProjectDelete(ByVal graph As Object, ByVal projectId As Long)
    Dim projects As New Collection, links As New Collection, row As Object, found As Boolean, scope As Long
    For Each row In graph("PROJECTS")
        If CLng(row("PROJECT_ID")) <> projectId Then projects.Add row Else found = True
    Next row
    If Not found Then Err.Raise 5, "Finance", Finance_Text("SOURCE_CHANGED")
    For Each row In graph("LINKS")
        If CLng(row("PROJECT_ID")) <> projectId Then links.Add row
    Next row
    Set graph("PROJECTS") = projects: Set graph("LINKS") = links
    Finance_GarbageCollect graph
    If projects.count = 0 Then
        graph("META")(1)("REFRESH_AT") = Empty
        Set row = FinanceStore_Row("PROJECTS")
        row("PROJECT_ID") = Finance_AllocateId(graph, "NEXT_PROJECT")
        row("NAME") = Finance_Text("DEFAULT_PROJECT"): row("DISPLAY_ORDER") = 1&
        projects.Add row
    End If
    scope = CLng(graph("META")(1)("SCOPE_ID"))
    If projects.count = 1 Or scope = projectId Then graph("META")(1)("SCOPE_ID") = projects(1)("PROJECT_ID")
    Finance_DomainValidate graph
End Sub

Public Sub Finance_SelectProject(ByVal projectId As Long)
    Dim graph As Object
    If Not WorkbookSchema_UserActionAllowed() Then Exit Sub
    If mDraftDepth > 0 Then Err.Raise 5, "Finance", Finance_Text("DRAFT_PENDING")
    Set graph = Finance_DomainSnapshot(): Finance_ReadVisibleInputs graph
    If CLng(graph("META")(1)("SCOPE_ID")) = projectId Then Finance_CommitDomain graph: Exit Sub
    graph("META")(1)("SCOPE_ID") = projectId
    Finance_DomainValidate graph
    Finance_CommitDomain graph, 2
End Sub

Public Sub Finance_SelectScope()
    Dim choice As DropDown, options As Variant, graph As Object, projectId As Long, description As String
    If Not (Application.ActiveWorkbook Is ThisWorkbook) Then Exit Sub
    If Application.ActiveSheet.name <> "FINANCE" Then Exit Sub
    On Error GoTo Failed
    Set graph = Finance_DomainSnapshot(): options = Finance_ProjectOptions(graph)
    Set choice = ThisWorkbook.Worksheets("FINANCE").DropDowns("peFinanceScope")
    If choice.value < 1 Or choice.value > UBound(options, 1) + 1 Then Err.Raise 5, "Finance", Finance_Text("SOURCE_CHANGED")
    projectId = CLng(options(choice.value - 1, 0))
    Finance_SelectProject projectId
    Exit Sub
Failed:
    description = Err.description
    On Error Resume Next: Finance_PresentationControls ThisWorkbook.Worksheets("FINANCE"): On Error GoTo 0
    CalcBridge_ShowSingleConsoleMessage "STOP", "COMMON.ERROR.PROCEDURE_INLINE", TextCatalog_Arguments("Procedure", "Finance", "Details", description)
End Sub

Private Sub Finance_ProjectControl(ByVal ws As Worksheet)
    Dim graph As Object, options As Variant, data() As Variant, choice As DropDown, i As Long, selected As Long
    Dim storage As Worksheet, stamp As shape, lastRow As Long
    If Not FinanceStore_Exists(ThisWorkbook) Then Exit Sub
    Set graph = Finance_PresentationDomain()
    On Error Resume Next: Set choice = ws.DropDowns("peFinanceScope"): On Error GoTo 0
    Set storage = ThisWorkbook.Worksheets(FIN_STORE_SHEET)
    If Not choice Is Nothing Then choice.ListFillRange = ""
    lastRow = storage.cells(storage.rows.count, "AP").End(xlUp).row
    storage.Range("AP1:AP" & CStr(lastRow)).ClearContents
    If graph("PROJECTS").count = 1 Then
        If Not choice Is Nothing Then
            choice.visible = False
        End If
        ws.Shapes("peFinanceLastRefresh").left = ws.Range("L1").left
        Exit Sub
    End If
    options = Finance_ProjectOptions(graph)
    ReDim data(1 To UBound(options, 1) + 1, 1 To 1)
    For i = 0 To UBound(options, 1)
        data(i + 1, 1) = "'" & CStr(options(i, 1))
        If CLng(options(i, 0)) = CLng(graph("META")(1)("SCOPE_ID")) Then selected = i + 1
    Next i
    storage.Range("AP1").Resize(UBound(data, 1), 1).formula = data
    If choice Is Nothing Then Set choice = ws.DropDowns.Add(336, 5, 180, 23): choice.name = "peFinanceScope"
    choice.ListFillRange = "'" & FIN_STORE_SHEET & "'!$AP$1:$AP$" & CStr(UBound(data, 1))
    choice.value = selected: choice.visible = (graph("PROJECTS").count > 1)
    choice.OnAction = "'" & Replace(ThisWorkbook.name, "'", "''") & "'!Finance_SelectScope"
    choice.Placement = xlFreeFloating
    Set stamp = ws.Shapes("peFinanceLastRefresh")
    stamp.left = ws.Range("L1").left
    If choice.visible Then
        If stamp.left < choice.left + choice.width + 10 Then stamp.left = choice.left + choice.width + 10
    End If
End Sub

Private Sub Finance_EnsureProjectionColumns(ByVal tbl As ListObject)
    Dim map As Object, key As Variant, width As Long, fields As Variant
    Set map = Finance_Map(tbl)
    If tbl.name = FIN_EXPENSES Then fields = Array("PROJECT_NAME") Else fields = Array("OTHERS", "RELATION")
    For Each key In fields
        If Not map.exists(CStr(key)) Then
            width = tbl.ListColumns.count + 1
            tbl.Resize tbl.Range.Resize(, width)
            tbl.ListColumns(width).name = Finance_ColumnTitle(CStr(key), Settings_GetOwnerLanguage("FINANCE"))
        End If
    Next key
End Sub

Private Sub Finance_ContextColumns(ByVal tbl As ListObject)
    Dim graph As Object, projects As Object, groups As Object, group As Collection, map As Object, data As Variant
    Dim context() As Variant, names As String, relation As String, link As Object, current As Object, i As Long
    Dim scope As Long, visible As Boolean, relationWidth As Double, oldScreen As Boolean, widthError As Long, widthDetails As String
    Set graph = Finance_PresentationDomain(): scope = CLng(graph("META")(1)("SCOPE_ID"))
    visible = (graph("PROJECTS").count > 1 And scope <> 0)
    Set map = Finance_Map(tbl)
    If CBool(tbl.ListColumns(CLng(map("OTHERS"))).Range.EntireColumn.hidden) = visible Then tbl.ListColumns(CLng(map("OTHERS"))).Range.EntireColumn.hidden = Not visible
    If CBool(tbl.ListColumns(CLng(map("RELATION"))).Range.EntireColumn.hidden) = visible Then tbl.ListColumns(CLng(map("RELATION"))).Range.EntireColumn.hidden = Not visible
    If Not visible Then Exit Sub
    If tbl.DataBodyRange Is Nothing Then Exit Sub
    Set projects = Finance_DomainIndex(graph("PROJECTS"), "PROJECT_ID"): Set groups = Finance_ActiveGroups(graph)
    data = tbl.DataBodyRange.Value2: ReDim context(1 To UBound(data, 1), 1 To 2)
    For i = 1 To UBound(data, 1)
        names = "": relation = "": Set current = Nothing
        If visible And groups.exists(CStr(data(i, CLng(map("WBS"))))) Then
            Set group = groups(CStr(data(i, CLng(map("WBS")))))
            If group.count > 1 Then
                For Each link In group
                    If CLng(link("PROJECT_ID")) <> scope Then
                        If Len(names) > 0 Then names = names & "; "
                        names = names & CStr(projects(CStr(link("PROJECT_ID")))("NAME"))
                    Else
                        Set current = link
                    End If
                Next link
                If Not current Is Nothing Then
                    If CLng(group(1)("FINANCIAL_ID")) = CLng(group(2)("FINANCIAL_ID")) Then
                        relation = Finance_Text("SHARED")
                    ElseIf Finance_Has(current("COPY_FROM")) Then
                        relation = Finance_Text("COPIED")
                    Else
                        relation = Finance_Text("INDEPENDENT")
                    End If
                End If
            End If
        End If
        context(i, 1) = "'" & names: context(i, 2) = "'" & relation
    Next i
    tbl.ListColumns(CLng(map("OTHERS"))).DataBodyRange.Resize(, 2).formula = context
    oldScreen = Application.ScreenUpdating
    On Error GoTo WidthFailed
    Application.ScreenUpdating = False
    With tbl.ListColumns(CLng(map("RELATION"))).Range
        relationWidth = .ColumnWidth
        .columns.AutoFit
        If .ColumnWidth < relationWidth Then .ColumnWidth = relationWidth
    End With
    Application.ScreenUpdating = oldScreen
    Exit Sub
WidthFailed:
    widthError = Err.number: widthDetails = Err.description
    Application.ScreenUpdating = oldScreen
    On Error GoTo 0
    Err.Raise widthError, "Finance_ContextColumns", widthDetails
End Sub

Private Function Finance_ChartExpenses() As Collection
    Dim graph As Object, rows As Collection, expenses As Collection, groups As Object
    If Not mProjectionJournal Is Nothing Then
        Set Finance_ChartExpenses = mProjectionJournal: Exit Function
    End If
    Set graph = Finance_PresentationDomain()
    Finance_DomainViews graph, CLng(graph("META")(1)("SCOPE_ID")), rows, expenses
    Set groups = Finance_ContributionGroups(rows)
    Set Finance_ChartExpenses = Finance_JournalProjection(groups, expenses)
End Function

Public Sub Finance_CheckRevision(ByVal scope As Long, ByVal revision As Long)
    If Finance_ScopeId() <> scope Or mViewRevision <> revision Then Err.Raise 5, "Finance", Finance_Text("SOURCE_CHANGED")
End Sub

Public Function Finance_Revision() As Long
    Dim scope As Long: scope = Finance_ScopeId()
    Finance_Revision = mViewRevision
End Function

Private Sub Finance_LinkSort(ByRef rows() As Object, ByVal low As Long, ByVal high As Long)
    Dim i As Long, j As Long, pivot As Long, swap As Object
    i = low: j = high: pivot = CLng(rows((low + high) \ 2)("ROW_ORDER"))
    Do While i <= j
        Do While CLng(rows(i)("ROW_ORDER")) < pivot: i = i + 1: Loop
        Do While CLng(rows(j)("ROW_ORDER")) > pivot: j = j - 1: Loop
        If i <= j Then
            Set swap = rows(i): Set rows(i) = rows(j): Set rows(j) = swap
            i = i + 1: j = j - 1
        End If
    Loop
    If low < j Then Finance_LinkSort rows, low, j
    If i < high Then Finance_LinkSort rows, i, high
End Sub

Public Sub Finance_DomainViews(ByVal graph As Object, ByVal scope As Long, ByRef rows As Collection, ByRef expenses As Collection, Optional ByVal contributions As Boolean = False, Optional ByVal expandJournal As Boolean = False)
    Dim records As Object, projects As Object, journal As Object, entries As Collection
    Dim link As PEFinanceRow, record As PEFinanceRow, row As PEFinanceRow, e As PEFinanceRow, entry As PEFinanceRow, key As Variant
    Dim selected() As Object, n As Long, i As Long, f As String, seen As Object
    Dim globalSeen As Object, projectNames As Object, names As Object, projectKey As Variant, label As String, position As Long
    Set records = Finance_DomainIndex(graph("RECORDS"), "FINANCIAL_ID")
    Set projects = Finance_DomainIndex(graph("PROJECTS"), "PROJECT_ID")
    Set journal = CreateObject("Scripting.Dictionary"): Set seen = CreateObject("Scripting.Dictionary")
    Set globalSeen = CreateObject("Scripting.Dictionary"): Set projectNames = CreateObject("Scripting.Dictionary")
    For Each entry In graph("JOURNAL")
        f = CStr(entry("FINANCIAL_ID"))
        If Not journal.exists(f) Then Set entries = New Collection: journal.Add f, entries
        journal(f).Add entry
    Next entry
    Set rows = New Collection: Set expenses = New Collection
    If graph("LINKS").count = 0 Then Exit Sub
    ReDim selected(1 To graph("LINKS").count)
    For Each link In graph("LINKS")
        If scope = 0 Or CLng(link("PROJECT_ID")) = scope Then
            f = CStr(link("FINANCIAL_ID"))
            If scope = 0 Then
                If Not projectNames.exists(f) Then Set names = CreateObject("Scripting.Dictionary"): projectNames.Add f, names
                projectNames(f)(CStr(link("PROJECT_ID"))) = CStr(projects(CStr(link("PROJECT_ID")))("NAME"))
                If globalSeen.exists(f) Then
                    position = CLng(globalSeen(f))
                    If selected(position)("STATE") = "HISTORICAL" And link("STATE") = "ACTIVE" Then Set selected(position) = link
                    GoTo NextViewLink
                End If
            End If
            n = n + 1: Set selected(n) = link
            If scope = 0 Then globalSeen.Add f, n
        End If
NextViewLink:
    Next link
    If n > 1 And scope <> 0 Then Finance_LinkSort selected, 1, n
    For i = 1 To n
        Set link = selected(i): f = CStr(link("FINANCIAL_ID")): Set record = records(f)
        Set row = Finance_NewRow(FIN_TABLE)
        For Each key In Array("WBS", "STATE", "TASK_ID", "TASK_NAME", "PROJECT", "START", "FINISH"): row(CStr(key)) = link(CStr(key)): Next key
        For Each key In Finance_InputKeys(): row(CStr(key)) = record(CStr(key)): Next key
        For Each key In Array("LINK_ID", "PROJECT_ID", "FINANCIAL_ID", "COPY_FROM"): row.Add CStr(key), link(CStr(key)): Next key
        row.Add "PROJECT_NAME", projects(CStr(link("PROJECT_ID")))("NAME")
        If scope = 0 Then
            label = ""
            For Each projectKey In projectNames(f).keys
                If Len(label) > 0 Then label = label & "; "
                label = label & CStr(projectNames(f)(projectKey))
            Next projectKey
            row("PROJECT_NAME") = label
        End If
        If contributions Then row.Add "CONTRIBUTION", link("LINK_ID")
        rows.Add row
        If journal.exists(f) And (expandJournal Or Not seen.exists(f)) Then
            For Each entry In journal(f)
                Set e = Finance_NewRow(FIN_EXPENSES)
                e("WBS") = link("WBS"): e("STATE") = link("STATE")
                For Each key In Array("DATE", "AMOUNT", "COMMENT"): e(CStr(key)) = entry(CStr(key)): Next key
                For Each key In Array("LINK_ID", "PROJECT_ID", "FINANCIAL_ID"): e.Add CStr(key), link(CStr(key)): Next key
                e.Add "EXPENSE_ID", entry("EXPENSE_ID")
                e("PROJECT_NAME") = row("PROJECT_NAME")
                If contributions Then e.Add "CONTRIBUTION", link("LINK_ID")
                expenses.Add e
            Next entry
        End If
        seen(f) = True
    Next i
End Sub

Private Function Finance_ContributionGroups(ByVal rows As Collection) As Object
    Dim groups As Object, members As Collection, row As Object, f As String
    Set groups = CreateObject("Scripting.Dictionary")
    For Each row In rows
        f = CStr(row("FINANCIAL_ID"))
        If Not groups.exists(f) Then Set members = New Collection: groups.Add f, members
        groups(f).Add row
    Next row
    Set Finance_ContributionGroups = groups
End Function

Private Function Finance_JournalProjection(ByVal groups As Object, ByVal expenses As Collection) As Collection
    Dim projected As New Collection, entry As PEFinanceRow, display As PEFinanceRow, members As Collection, member As PEFinanceRow
    Dim key As Variant, names As Object, labels As Object, f As String, text As String, state As String, nameKey As String
    Set labels = CreateObject("Scripting.Dictionary"): Set names = CreateObject("Scripting.Dictionary")
    For Each entry In expenses
        f = CStr(entry("FINANCIAL_ID")): Set members = groups(f)
        If Not labels.exists(f) Then
            text = "": state = CStr(members(1)("STATE"))
            For Each member In members
                nameKey = f & "|" & CStr(member("PROJECT_ID"))
                If Not names.exists(nameKey) Then
                    If Len(text) > 0 Then text = text & "; "
                    text = text & CStr(member("PROJECT_NAME")): names.Add nameKey, True
                End If
                If CStr(member("STATE")) <> state Then state = "MIXED"
            Next member
            If Len(text) > 30000 Then text = left$(text, 29970) & "..."
            labels.Add f, Array(text, state)
        End If
        Set display = Finance_NewRow(FIN_EXPENSES)
        For Each key In Finance_Keys(FIN_EXPENSES): display(CStr(key)) = entry(CStr(key)): Next key
        display("AMOUNT") = CCur(entry("AMOUNT")) * CLng(members.count)
        display("PROJECT_NAME") = CStr(labels(f)(0)): display("STATE") = CStr(labels(f)(1))
        projected.Add display
    Next entry
    Set Finance_JournalProjection = projected
End Function

Private Function Finance_FinancialReference(ByVal row As PEFinanceRow) As String
    If row.exists("FINANCIAL_ID") Then
        Finance_FinancialReference = "FIN|" & CStr(row("FINANCIAL_ID"))
    Else
        Finance_FinancialReference = Finance_Reference(row)
    End If
End Function

Private Sub Finance_AssertBinding(ByVal graph As Object)
    If Not mViewBound Then Finance_BindSavedView
    If mStorageFailed Or Not mViewBound Then Err.Raise 5, "Finance", Finance_Text("STORAGE_BLOCKED")
    If CLng(graph("META")(1)("SCOPE_ID")) <> mViewScope Or CLng(graph("META")(1)("REVISION")) <> mViewRevision Then Err.Raise 5, "Finance", Finance_Text("SOURCE_CHANGED")
End Sub

Private Sub Finance_ReadVisibleInputs(ByVal graph As Object)
    Dim rows As Collection, links As Object, records As Object, link As Object, row As Object, key As Variant
    Dim ref As String, count As Long, seen As Object, previousOrder As Long, reordered As Boolean
    Dim perf As clsPerfScope
    Set perf = Profiler_BeginScope("Finance.View.FlushRead", "Finance")
    Finance_AssertBinding graph
    If mViewScope = 0 Then Exit Sub
    Set links = CreateObject("Scripting.Dictionary"): Set seen = CreateObject("Scripting.Dictionary")
    Set records = Finance_DomainIndex(graph("RECORDS"), "FINANCIAL_ID")
    For Each link In graph("LINKS")
        If CLng(link("PROJECT_ID")) = mViewScope Then links.Add Finance_Reference(link), link
    Next link
    Set rows = Finance_Read(Finance_Table(FIN_TABLE))
    If rows.count <> links.count Then Err.Raise 5, "Finance", Finance_Text("SOURCE_CHANGED")
    For Each row In rows
        ref = Finance_Reference(row)
        If Not links.exists(ref) Or seen.exists(ref) Then Err.Raise 5, "Finance", Finance_Text("SOURCE_CHANGED")
        seen.Add ref, True: Set link = links(ref): count = count + 1
        records(CStr(link("FINANCIAL_ID")))("BUDGET") = Finance_Money(row("BUDGET"))
        records(CStr(link("FINANCIAL_ID")))("ETC") = Finance_Money(row("ETC"))
        records(CStr(link("FINANCIAL_ID")))("PROFILE") = Finance_Profile(CStr(row("PROFILE")))
        Select Case CStr(records(CStr(link("FINANCIAL_ID")))("PROFILE"))
            Case "START", "END", "LINEAR"
            Case Else: Err.Raise 5, "Finance", Finance_Text("INVALID_PROFILE")
        End Select
        If CLng(link("ROW_ORDER")) <= previousOrder Then reordered = True
        previousOrder = CLng(link("ROW_ORDER"))
    Next row
    If reordered Then
        count = 0
        For Each row In rows
            count = count + 1
            links(Finance_Reference(row))("ROW_ORDER") = count
        Next row
    End If
End Sub

Private Sub Finance_GarbageCollect(ByVal graph As Object)
    Dim referenced As Object, records As New Collection, journal As New Collection, row As Object
    Set referenced = CreateObject("Scripting.Dictionary")
    For Each row In graph("LINKS"): referenced(CStr(row("FINANCIAL_ID"))) = True: Next row
    For Each row In graph("RECORDS")
        If referenced.exists(CStr(row("FINANCIAL_ID"))) Then records.Add row
    Next row
    For Each row In graph("JOURNAL")
        If referenced.exists(CStr(row("FINANCIAL_ID"))) Then journal.Add row
    Next row
    Set graph("RECORDS") = records: Set graph("JOURNAL") = journal
End Sub

Private Function Finance_ActiveGroups(ByVal graph As Object) As Object
    Dim groups As Object, links As Collection, row As Object, wbs As String
    Set groups = CreateObject("Scripting.Dictionary")
    For Each row In graph("LINKS")
        If row("STATE") = "ACTIVE" Then
            wbs = CStr(row("WBS"))
            If Not groups.exists(wbs) Then Set links = New Collection: groups.Add wbs, links
            groups(wbs).Add row
        End If
    Next row
    Set Finance_ActiveGroups = groups
End Function

Public Function Finance_DuplicationTargets(ByVal graph As Object, ByVal scope As Long, ByVal rows As Collection) As Collection
    Dim result As New Collection, existing As Object, groups As Object, row As Object, link As Object
    Set existing = CreateObject("Scripting.Dictionary"): Set groups = Finance_ActiveGroups(graph)
    For Each link In graph("LINKS")
        If CLng(link("PROJECT_ID")) = scope Then existing(Finance_Reference(link)) = True
    Next link
    For Each row In rows
        If row("STATE") = "ACTIVE" And Not existing.exists(Finance_Reference(row)) Then
            If groups.exists(CStr(row("WBS"))) Then
                If groups(CStr(row("WBS"))).count = 1 Then result.Add row
            End If
        End If
    Next row
    Set Finance_DuplicationTargets = result
End Function

Private Function Finance_NewFinancial(ByVal graph As Object, Optional ByVal sourceId As Long = 0) As Object
    Dim record As Object, source As Object, e As Object, copy As Object, key As Variant, copies As New Collection
    Set record = FinanceStore_Row("RECORDS")
    record("FINANCIAL_ID") = Finance_AllocateId(graph, "NEXT_FINANCIAL"): record("PROFILE") = "LINEAR"
    If sourceId > 0 Then
        Set source = Finance_DomainIndex(graph("RECORDS"), "FINANCIAL_ID")(CStr(sourceId))
        For Each key In Finance_InputKeys(): record(CStr(key)) = source(CStr(key)): Next key
        For Each e In graph("JOURNAL")
            If CLng(e("FINANCIAL_ID")) = sourceId Then
                Set copy = FinanceStore_Row("JOURNAL")
                copy("EXPENSE_ID") = Finance_AllocateId(graph, "NEXT_EXPENSE"): copy("FINANCIAL_ID") = record("FINANCIAL_ID")
                For Each key In Array("DATE", "AMOUNT", "COMMENT"): copy(CStr(key)) = e(CStr(key)): Next key
                copies.Add copy
            End If
        Next e
        For Each e In copies: graph("JOURNAL").Add e: Next e
    End If
    graph("RECORDS").Add record
    Set Finance_NewFinancial = record
End Function

Private Sub Finance_ApplyMemberships(ByVal graph As Object, ByVal scope As Long, ByVal rows As Collection, ByVal decisions As Object)
    Dim previous As Object, kept As New Collection, groups As Object, group As Collection, records As Object
    Dim row As Object, link As Object, source As Object, record As Object, key As Variant, ref As String, choice As String, i As Long
    Dim phase As String, n As Long, description As String
    On Error GoTo Failed
    phase = "Index active links"
    Set previous = CreateObject("Scripting.Dictionary"): Set groups = Finance_ActiveGroups(graph)
    For Each link In graph("LINKS")
        If CLng(link("PROJECT_ID")) = scope Then
            previous.Add Finance_Reference(link), link
        Else
            kept.Add link
        End If
    Next link
    Set records = Finance_DomainIndex(graph("RECORDS"), "FINANCIAL_ID")
    For Each row In rows
        phase = "Resolve existing link"
        i = i + 1: ref = Finance_Reference(row): Set link = Nothing
        If previous.exists(ref) Then
            Set link = previous(ref)
        ElseIf row.exists("LINK_ID") Then
            If previous.count > 0 Then
                For Each key In previous.keys
                    Set source = previous(CStr(key))
                    If CLng(source("LINK_ID")) = CLng(row("LINK_ID")) Then Set link = source: Exit For
                Next key
            End If
        End If
        If link Is Nothing Then
            phase = "Allocate link"
            Set link = FinanceStore_Row("LINKS")
            link("LINK_ID") = Finance_AllocateId(graph, "NEXT_LINK"): link("PROJECT_ID") = scope
            choice = "INDEPENDENT": Set group = Nothing
            If groups.exists(CStr(row("WBS"))) Then
                phase = "Read sharing group"
                Set group = groups(CStr(row("WBS")))
                If group.count = 1 Then
                    If decisions Is Nothing Then Err.Raise 5, "Finance", Finance_Text("DUP_REQUIRED")
                    If Not decisions.exists(CStr(row("WBS"))) Then Err.Raise 5, "Finance", Finance_Text("DUP_REQUIRED")
                    choice = CStr(decisions(CStr(row("WBS"))))
                ElseIf CLng(group(1)("FINANCIAL_ID")) = CLng(group(2)("FINANCIAL_ID")) Then
                    choice = "SHARED"
                End If
            End If
            Select Case choice
                Case "SHARED"
                    phase = "Reference shared financial record"
                    If group Is Nothing Then Err.Raise 5, "Finance", Finance_Text("INVALID_DOMAIN")
                    link("FINANCIAL_ID") = group(1)("FINANCIAL_ID")
                Case "COPIED"
                    If group Is Nothing Then Err.Raise 5, "Finance", Finance_Text("INVALID_DOMAIN")
                    Set record = Finance_NewFinancial(graph, CLng(group(1)("FINANCIAL_ID")))
                    link("FINANCIAL_ID") = record("FINANCIAL_ID"): link("COPY_FROM") = group(1)("FINANCIAL_ID")
                Case "INDEPENDENT"
                    Set record = Finance_NewFinancial(graph): link("FINANCIAL_ID") = record("FINANCIAL_ID")
                Case Else: Err.Raise 5, "Finance", Finance_Text("DUP_REQUIRED")
            End Select
        End If
        phase = "Copy link metadata"
        link("ROW_ORDER") = i
        For Each key In Array("WBS", "STATE", "TASK_ID", "TASK_NAME", "PROJECT", "START", "FINISH"): link(CStr(key)) = row(CStr(key)): Next key
        kept.Add link
    Next row
    Set graph("LINKS") = kept
    phase = "Collect unreferenced records"
    Finance_GarbageCollect graph
    phase = "Validate link graph"
    Finance_DomainValidate graph
    Exit Sub
Failed:
    n = Err.number: description = Err.description
    Err.Raise n, "Finance", phase & ": " & description
End Sub

Private Sub Finance_ApplyJournal(ByVal graph As Object, ByVal scope As Long, ByVal expenses As Collection)
    Dim links As Object, allowed As Object, seen As Object, original As Object, kept As New Collection
    Dim link As Object, row As Object, entry As Object, key As Variant, id As Long, f As String
    Set links = CreateObject("Scripting.Dictionary"): Set allowed = CreateObject("Scripting.Dictionary")
    Set seen = CreateObject("Scripting.Dictionary"): Set original = Finance_DomainIndex(graph("JOURNAL"), "EXPENSE_ID")
    For Each link In graph("LINKS")
        If CLng(link("PROJECT_ID")) = scope Then
            links.Add Finance_Reference(link), link: allowed(CStr(link("FINANCIAL_ID"))) = True
        End If
    Next link
    For Each row In graph("JOURNAL")
        If Not allowed.exists(CStr(row("FINANCIAL_ID"))) Then kept.Add row
    Next row
    For Each row In expenses
        If Not links.exists(Finance_Reference(row)) Then Err.Raise 5, "Finance", Finance_Text("INVALID_DOMAIN")
        f = CStr(links(Finance_Reference(row))("FINANCIAL_ID")): id = 0
        If row.exists("EXPENSE_ID") Then
            If Finance_Has(row("EXPENSE_ID")) Then id = Finance_Id(row("EXPENSE_ID"))
        End If
        If id > 0 Then
            If Not original.exists(CStr(id)) Then Err.Raise 5, "Finance", Finance_Text("SOURCE_CHANGED")
            If Not allowed.exists(CStr(original(CStr(id))("FINANCIAL_ID"))) Then Err.Raise 5, "Finance", Finance_Text("SOURCE_CHANGED")
        Else
            id = Finance_AllocateId(graph, "NEXT_EXPENSE")
        End If
        If seen.exists(CStr(id)) Then Err.Raise 5, "Finance", Finance_Text("INVALID_DOMAIN")
        seen.Add CStr(id), True
        Set entry = FinanceStore_Row("JOURNAL"): entry("EXPENSE_ID") = id: entry("FINANCIAL_ID") = CLng(f)
        For Each key In Array("DATE", "AMOUNT", "COMMENT"): entry(CStr(key)) = row(CStr(key)): Next key
        kept.Add entry
    Next row
    Set graph("JOURNAL") = kept
    Finance_DomainValidate graph
End Sub

Private Function Finance_GlobalRows(ByVal rows As Collection) As Collection
    Dim grouped As Object, result As New Collection, row As Object, total As Object, key As Variant, ref As String
    Set grouped = CreateObject("Scripting.Dictionary")
    For Each row In rows
        ref = CStr(row("STATE")) & "|" & CStr(row("WBS"))
        If Not grouped.exists(ref) Then
            Set total = Finance_NewRow(FIN_TABLE)
            For Each key In Finance_Keys(FIN_TABLE): total(CStr(key)) = row(CStr(key)): Next key
            grouped.Add ref, total: result.Add total
        Else
            Set total = grouped(ref)
            For Each key In Array("BUDGET", "ACTUAL", "ETC", "EAC", "VARIANCE")
                If Finance_Has(row(CStr(key))) Then
                    If Finance_Has(total(CStr(key))) Then total(CStr(key)) = CCur(total(CStr(key))) + CCur(row(CStr(key))) Else total(CStr(key)) = row(CStr(key))
                End If
            Next key
            If CStr(total("PROFILE")) <> CStr(row("PROFILE")) Then total("PROFILE") = "MIXED"
            For Each key In Array("TASK_ID", "TASK_NAME", "PROJECT", "START", "FINISH")
                If CStr(total(CStr(key))) <> CStr(row(CStr(key))) Then total(CStr(key)) = Empty
            Next key
            If Len(CStr(row("DIAGNOSTIC"))) > 0 Then
                If InStr(1, CStr(total("DIAGNOSTIC")), CStr(row("DIAGNOSTIC")), vbBinaryCompare) = 0 Then
                    If Len(CStr(total("DIAGNOSTIC"))) > 0 Then total("DIAGNOSTIC") = CStr(total("DIAGNOSTIC")) & vbLf
                    total("DIAGNOSTIC") = CStr(total("DIAGNOSTIC")) & CStr(row("DIAGNOSTIC"))
                End If
            End If
        End If
    Next row
    Set Finance_GlobalRows = result
End Function

Public Sub Finance_PrepareImportedViews(ByVal graph As Object, ByRef rows As Collection, ByRef expenses As Collection)
    Dim emptyRows As New Collection, groups As Object
    Finance_EchoRows rows, expenses, emptyRows, False
    Set groups = Finance_ContributionGroups(rows)
    Set expenses = Finance_JournalProjection(groups, expenses)
    If CLng(graph("META")(1)("SCOPE_ID")) = 0 Then Set rows = Finance_GlobalRows(rows)
End Sub

Private Sub Finance_EchoRows(ByVal rows As Collection, ByVal expenses As Collection, ByVal previous As Collection, ByVal sameScope As Boolean)
    Dim old As Object, actuals As Object, row As PEFinanceRow, ref As String, key As Variant
    Set old = CreateObject("Scripting.Dictionary")
    If sameScope Then
        For Each row In previous: old.Add Finance_Reference(row), row: Next row
    End If
    Set actuals = Finance_RecordedCosts(expenses)
    For Each row In rows
        ref = Finance_Reference(row)
        If old.exists(ref) Then
            For Each key In Array("PROGRESS", "START", "FINISH", "EAC", "VARIANCE", "DIAGNOSTIC"): row(CStr(key)) = old(ref)(CStr(key)): Next key
        End If
        row("ACTUAL") = Empty
        If actuals.exists(Finance_FinancialReference(row)) Then row("ACTUAL") = actuals(Finance_FinancialReference(row))
        If old.exists(ref) Then
            If CStr(old(ref)("ACTUAL")) <> CStr(row("ACTUAL")) Then row("EAC") = Empty: row("VARIANCE") = Empty
        End If
    Next row
End Sub

Private Sub Finance_CommitDomain(ByVal candidate As Object, Optional ByVal projection As Long = 0, Optional ByVal activities As Collection)
    Dim before As Object, actual As Object, computed As Object, meta As Object, links As Object, groups As Object
    Dim oldRows As Collection, oldExpenses As Collection, rows As Collection, expenses As Collection
    Dim oldMonths As Variant, oldTotals As Variant, oldStamp As Variant, row As Object, key As Variant
    Dim report As Worksheet, ws As Worksheet, targetScope As Long, n As Long, description As String, restoreError As String
    Dim oldEvents As Boolean, oldScreen As Boolean, oldScope As Long, oldRevision As Long, stored As Boolean
    Dim oldChartCount As Long, oldChartName As String, oldChartVisible As Boolean
    Dim oldSeries(1 To 4) As String, seriesIndex As Long
    Dim oldSeriesValues(1 To 4) As Variant
    Dim phase As String
    Dim perf As clsPerfScope, phasePerf As clsPerfScope
    Set perf = Profiler_BeginScope("Finance.Transaction", "Finance")
    Set phasePerf = Profiler_BeginScope("Finance.Transaction.Capture", "Finance")
    If mTransaction Or mStorageFailed Then Err.Raise 5, "Finance", Finance_Text("STORAGE_BLOCKED")
    oldEvents = Application.EnableEvents: oldScreen = Application.ScreenUpdating
    oldScope = mViewScope: oldRevision = mViewRevision
    phase = "Capture"
    On Error GoTo Failed
    ' The caller validated its read boundary. Recheck binding and validate the candidate once.
    Set before = FinanceStore_Read(ThisWorkbook): Finance_AssertBinding before
    Finance_DomainValidate candidate
    If FinanceStore_Token(candidate) = FinanceStore_Token(before) And projection = 0 Then Exit Sub
    oldEvents = Application.EnableEvents: oldScreen = Application.ScreenUpdating
    oldScope = mViewScope: oldRevision = mViewRevision
    Set ws = ThisWorkbook.Worksheets("FINANCE"): Set report = Finance_SummarySheet()
    targetScope = CLng(candidate("META")(1)("SCOPE_ID"))
    Set oldRows = Finance_Read(Finance_Table(FIN_TABLE)): Set oldExpenses = Finance_Read(Finance_Table(FIN_EXPENSES))
    oldStamp = ws.Range("S2").Value2
    If projection > 0 Then
        oldMonths = Finance_OutputSnapshot(report, "tbl_FINANCE_MONTHS")
        oldTotals = Finance_OutputSnapshot(report, "tbl_FINANCE_TOTALS")
        oldChartCount = report.ChartObjects.count
        If oldChartCount > 0 Then
            oldChartName = report.ChartObjects(1).name: oldChartVisible = report.ChartObjects(1).visible
            For seriesIndex = 1 To report.ChartObjects(1).chart.SeriesCollection.count
                If seriesIndex > 4 Then Exit For
                oldSeries(seriesIndex) = report.ChartObjects(1).chart.SeriesCollection(seriesIndex).formula
                If seriesIndex = 2 Or seriesIndex = 4 Then oldSeriesValues(seriesIndex) = Finance_ChartSeriesValues(report.ChartObjects(1).chart.SeriesCollection(seriesIndex))
            Next seriesIndex
        End If
    End If
    mTransaction = True: Set mProjectionDomain = candidate: mProjectionOutputsChanged = (projection = 2)
    phase = "Calculate"
    Set phasePerf = Nothing: Set phasePerf = Profiler_BeginScope("Finance.Transaction.Calculate", "Finance")
    If projection = 2 Then
        Finance_DomainViews candidate, targetScope, rows, expenses, True
        Set groups = Finance_ContributionGroups(rows)
        Set computed = Finance_Compute(rows, expenses, activities, groups)
        Set links = Finance_DomainIndex(candidate("LINKS"), "LINK_ID")
        For Each row In rows
            For Each key In Array("TASK_ID", "TASK_NAME", "PROJECT", "START", "FINISH"): links(CStr(row("LINK_ID")))(CStr(key)) = row(CStr(key)): Next key
        Next row
        Set expenses = Finance_JournalProjection(groups, expenses)
        If targetScope = 0 Then Set rows = Finance_GlobalRows(rows)
    ElseIf projection = 1 Then
        Finance_DomainViews candidate, targetScope, rows, expenses
        Finance_EchoRows rows, expenses, oldRows, (targetScope = oldScope)
        Set groups = Finance_ContributionGroups(rows)
        Set expenses = Finance_JournalProjection(groups, expenses)
        If targetScope = 0 Then Set rows = Finance_GlobalRows(rows)
    End If
    If oldRevision >= 2147483646 Then Err.Raise 5, "Finance", Finance_Text("INVALID_DOMAIN")
    If Not expenses Is Nothing Then Set mProjectionJournal = expenses
    candidate("META")(1)("REVISION") = oldRevision + 1
    Application.EnableEvents = False: Application.ScreenUpdating = False
    stored = True
    phase = "Store data"
    Set phasePerf = Nothing: Set phasePerf = Profiler_BeginScope("Finance.Transaction.StoreData", "Finance")
    If projection > 0 Then
        ' Publish selection only after both projections have succeeded.
        FinanceStore_Write ThisWorkbook, candidate, "DATA", before
        Set phasePerf = Nothing: Set phasePerf = Profiler_BeginScope("Finance.Transaction.ProjectRows", "Finance")
        phase = "Project rows"
        If projection <> 3 Then
            Finance_Write Finance_Table(FIN_TABLE), rows
            Finance_Write Finance_Table(FIN_EXPENSES), expenses
        End If
        ws.Range("S1").Value2 = candidate("META")(1)("CURRENCY")
        ws.Range("S2").Value2 = candidate("META")(1)("REFRESH_AT")
        Set phasePerf = Nothing: Set phasePerf = Profiler_BeginScope("Finance.Transaction.OutputTables", "Finance")
        phase = "Output tables"
        If projection = 2 Then Finance_WriteComputedOutputs computed("MONTHS"), computed("TOTALS")
        Set phasePerf = Nothing: Set phasePerf = Profiler_BeginScope("Finance.Transaction.Present", "Finance")
        phase = "Present"
        Finance_Present (projection = 2)
        If report.ChartObjects.count > 0 And projection = 2 Then report.ChartObjects(1).visible = (CLng(computed("COUNT")) > 0)
        Set phasePerf = Nothing: Set phasePerf = Profiler_BeginScope("Finance.Transaction.PublishMeta", "Finance")
        phase = "Publish metadata"
        FinanceStore_Write ThisWorkbook, candidate, "META"
    Else
        FinanceStore_Write ThisWorkbook, candidate, , before
    End If
    Set phasePerf = Nothing: Set phasePerf = Profiler_BeginScope("Finance.Transaction.Verify", "Finance")
    phase = "Verify"
    Set actual = FinanceStore_Read(ThisWorkbook)
    If FinanceStore_Token(actual) <> FinanceStore_Token(candidate) Then Err.Raise 5, "Finance", Finance_Text("INVALID_DOMAIN")
    mViewScope = targetScope: mViewRevision = CLng(candidate("META")(1)("REVISION")): mViewBound = True
    Ribbon_InvalidateContext
    GoTo CleanExit
Failed:
    n = Err.number: description = phase & ": " & Err.description
    Resume Rollback
Rollback:
    On Error GoTo RestoreFailed
    If stored Then
        Set mProjectionDomain = before
        Set mProjectionJournal = oldExpenses
        Set actual = FinanceStore_Read(ThisWorkbook)
        If FinanceStore_Token(actual) <> FinanceStore_Token(before) Then FinanceStore_Write ThisWorkbook, before, , actual
        If projection > 0 Then
            If Not Finance_RowsEqual(oldRows, Finance_Read(Finance_Table(FIN_TABLE)), FIN_TABLE) Then Finance_Write Finance_Table(FIN_TABLE), oldRows
            If Not Finance_RowsEqual(oldExpenses, Finance_Read(Finance_Table(FIN_EXPENSES)), FIN_EXPENSES) Then Finance_Write Finance_Table(FIN_EXPENSES), oldExpenses
            ws.Range("S1").Value2 = before("META")(1)("CURRENCY"): ws.Range("S2").Value2 = oldStamp
            Finance_WriteComputedOutputs oldMonths, oldTotals
            Finance_Present True
            If oldChartCount = 0 Then
                Do While report.ChartObjects.count > 0: report.ChartObjects(1).Delete: Loop
            Else
                If report.ChartObjects.count <> oldChartCount Or report.ChartObjects(1).name <> oldChartName Then Err.Raise 5, , "ChartObject rollback mismatch"
                report.ChartObjects(1).visible = oldChartVisible
                For seriesIndex = 1 To 4
                    If Len(oldSeries(seriesIndex)) > 0 Then
                        If seriesIndex = 2 Or seriesIndex = 4 Then report.ChartObjects(1).chart.SeriesCollection(seriesIndex).values = oldSeriesValues(seriesIndex)
                        If report.ChartObjects(1).chart.SeriesCollection(seriesIndex).formula <> oldSeries(seriesIndex) Then Err.Raise 5, , "Chart series rollback mismatch"
                    End If
                Next seriesIndex
            End If
            If Not Finance_RowsEqual(oldRows, Finance_Read(Finance_Table(FIN_TABLE)), FIN_TABLE) Then Err.Raise 5, , "Input rollback mismatch"
            If Not Finance_RowsEqual(oldExpenses, Finance_Read(Finance_Table(FIN_EXPENSES)), FIN_EXPENSES) Then Err.Raise 5, , "Journal rollback mismatch"
            If Not Finance_OutputEqual(oldMonths, Finance_OutputSnapshot(report, "tbl_FINANCE_MONTHS")) Then Err.Raise 5, , "Monthly rollback mismatch"
            If Not Finance_OutputEqual(oldTotals, Finance_OutputSnapshot(report, "tbl_FINANCE_TOTALS")) Then Err.Raise 5, , "Totals rollback mismatch"
        End If
        Set actual = FinanceStore_Read(ThisWorkbook)
        If FinanceStore_Token(actual) <> FinanceStore_Token(before) Then Err.Raise 5, , "Storage rollback mismatch"
    End If
    GoTo CleanExit
RestoreFailed:
    restoreError = Err.description: mStorageFailed = True
    Resume CleanExit
CleanExit:
    On Error GoTo 0
    Set phasePerf = Nothing
    If n <> 0 Then mViewScope = oldScope: mViewRevision = oldRevision
    Set mProjectionDomain = Nothing: Set mProjectionJournal = Nothing: mProjectionOutputsChanged = False: mTransaction = False
    Application.EnableEvents = oldEvents: Application.ScreenUpdating = oldScreen
    If Len(restoreError) > 0 Then Err.Raise 5, "Finance", Finance_Text("STORAGE_BLOCKED") & vbCrLf & description & vbCrLf & restoreError
    If n <> 0 Then Err.Raise n, "Finance", description
End Sub

Private Function Finance_ChartSeriesValues(ByVal series As series) As Variant
    Dim data As Variant, singleValue(1 To 1) As Variant, i As Long
    data = series.values
    If Not IsArray(data) Then
        singleValue(1) = data: data = singleValue
    End If
    ' Excel returns chart #N/A points as Empty; restore actual gaps, not zeroes.
    For i = LBound(data) To UBound(data)
        If IsEmpty(data(i)) Then data(i) = CVErr(xlErrNA)
    Next i
    Finance_ChartSeriesValues = data
End Function

Private Sub Finance_WriteComputedOutputs(ByVal months As Variant, ByVal totals As Variant)
    Dim report As Worksheet, tbl As ListObject, count As Long
    Set report = Finance_SummarySheet()
    If Not IsEmpty(months) Then count = UBound(months, 1)
    On Error Resume Next: Set tbl = Finance_Table("tbl_FINANCE_TOTALS"): On Error GoTo 0
    If Not tbl Is Nothing Then
        tbl.Range.EntireRow.hidden = False: tbl.Range.EntireRow.OutlineLevel = 1
        If report.Range(Finance_TotalsAddress(count)).row > tbl.Range.row Then Finance_PositionOutput tbl, Finance_TotalsAddress(count), report
    End If
    Finance_Output report, "tbl_FINANCE_MONTHS", "A25", Finance_OutputHeaders("tbl_FINANCE_MONTHS"), months
    Finance_Output report, "tbl_FINANCE_TOTALS", Finance_TotalsAddress(), Finance_OutputHeaders("tbl_FINANCE_TOTALS"), totals
End Sub

Public Sub Finance_Flush()
    Dim graph As Object
    If Not FinanceStore_Exists(ThisWorkbook) Then Exit Sub
    If mTransaction Then Err.Raise 5, "Finance", Finance_Text("STORAGE_BLOCKED")
    If mDraftDepth > 0 Then Err.Raise 5, "Finance", Finance_Text("DRAFT_PENDING")
    Set graph = Finance_DomainSnapshot(): Finance_ReadVisibleInputs graph
    Finance_CommitDomain graph
End Sub

Public Function Finance_DraftRows() As Collection
    Dim graph As Object, rows As Collection, expenses As Collection
    Set graph = Finance_DomainSnapshot(): Finance_AssertBinding graph
    Finance_RequireIndividual
    Finance_DomainViews graph, mViewScope, rows, expenses
    Finance_EchoRows rows, expenses, Finance_Read(Finance_Table(FIN_TABLE)), True
    Set Finance_DraftRows = rows
End Function

Public Function Finance_DraftExpenses() As Collection
    Dim graph As Object, rows As Collection, expenses As Collection
    Set graph = Finance_DomainSnapshot(): Finance_AssertBinding graph
    Finance_RequireIndividual
    Finance_DomainViews graph, mViewScope, rows, expenses
    Set Finance_DraftExpenses = expenses
End Function

Public Sub Finance_Refresh()
    Dim graph As Object, activities As Collection
    Dim n As Long, description As String, perf As clsPerfScope
    If Not WorkbookSchema_UserActionAllowed() Then Exit Sub
    If mBusy Then Err.Raise 5, "Finance", Finance_Text("STORAGE_BLOCKED")
    On Error GoTo Failed
    Set perf = Profiler_BeginScope("Finance.Command.Refresh", "Finance")
    mBusy = True
    Set graph = Finance_DomainSnapshot(): Finance_ReadVisibleInputs graph
    Set activities = Finance_Activities()
    If Not Finance_ResolveMissing(graph, activities) Then GoTo CleanExit
    graph("META")(1)("REFRESH_AT") = CDbl(Now)
    Finance_CommitDomain graph, 2, activities
    mCalculations = mCalculations + 1
CleanExit:
    mBusy = False
    If n <> 0 Then Err.Raise n, "Finance_Refresh", description
    Exit Sub
Failed:
    n = Err.number: description = Err.description: Resume CleanExit
End Sub

Public Function Finance_Keys(ByVal tableName As String) As Variant
    Select Case tableName
        Case FIN_TABLE
            Finance_Keys = Array("WBS", "TASK_NAME", "STATE", "PROGRESS", "BUDGET", "PROFILE", "ACTUAL", "ETC", "EAC", "VARIANCE", "START", "FINISH", "PROJECT", "TASK_ID", "DIAGNOSTIC", "OTHERS", "RELATION")
        Case FIN_EXPENSES
            Finance_Keys = Array("WBS", "STATE", "DATE", "AMOUNT", "COMMENT", "PROJECT_NAME")
        Case Else: Err.Raise 5, "Finance", "Unknown Finance table"
    End Select
End Function

Public Function Finance_Text(ByVal key As String) As String
    Finance_Text = TextCatalog_Get("FINANCE." & key, Settings_GetOwnerLanguage("FINANCE"))
End Function

Public Function Finance_CurrencySymbol() As String
    Dim code As String
    code = CStr(ThisWorkbook.Worksheets("FINANCE").Range("S1").Value2)
    Select Case code
        Case "EUR": Finance_CurrencySymbol = ChrW(&H20AC)
        Case "USD": Finance_CurrencySymbol = "$"
        Case "GBP": Finance_CurrencySymbol = ChrW(&HA3)
        Case "CAD": Finance_CurrencySymbol = "C$"
        Case "AUD": Finance_CurrencySymbol = "A$"
        Case Else: Finance_CurrencySymbol = code
    End Select
End Function

Private Function Finance_MoneyFormat() As String
    Finance_MoneyFormat = "#,##0.00 """ & Replace$(Finance_CurrencySymbol(), """", """""") & """"
End Function

Public Function Finance_MoneyCaption(ByVal value As Variant, Optional ByVal symbol As String = "") As String
    If symbol = "" Then symbol = Finance_CurrencySymbol()
    If Finance_Has(value) Then Finance_MoneyCaption = Format$(value, "#,##0.00") & " " & symbol
End Function

Private Function Finance_WBSSortKey(ByVal wbs As String) As String
    Dim part As Variant
    For Each part In Split(wbs, ".")
        Finance_WBSSortKey = Finance_WBSSortKey & Format$(Len(CStr(part)), "000000") & CStr(part) & "/"
    Next part
End Function

Private Sub Finance_OutlineTotals(ByVal tbl As ListObject, Optional ByVal regroup As Boolean = True)
    Dim data As Variant, keys() As String, ranks() As Long, nodes As Object, i As Long, j As Long, parent As String, token As String, row As Range
    Dim perf As clsPerfScope
    Set perf = Profiler_BeginScope("Finance.OutlineTotals", "Finance")
    If tbl.DataBodyRange Is Nothing Then Exit Sub
    data = tbl.DataBodyRange.Value2: ReDim keys(1 To UBound(data, 1)): ReDim ranks(1 To UBound(data, 1))
    Set nodes = CreateObject("Scripting.Dictionary")
    For i = 2 To UBound(data, 1)
        token = Split(Trim$(CStr(data(i, 1))), " ")(0)
        If left$(token, 1) Like "#" Then keys(i) = token: nodes(token) = True
    Next i
    tbl.DataBodyRange.font.bold = False: tbl.DataBodyRange.font.size = 11
    tbl.DataBodyRange.columns(1).IndentLevel = 0
    For i = 1 To UBound(data, 1)
        Set row = tbl.DataBodyRange.rows(i)
        If i = 1 Then
            row.font.bold = True: row.font.size = 16: If row.rowHeight < 24 Then row.rowHeight = 24
        ElseIf Len(keys(i)) > 0 Then
            ranks(i) = 1: parent = GetParentWBS(keys(i))
            Do While Len(parent) > 0
                If nodes.exists(parent) Then ranks(i) = ranks(i) + 1
                parent = GetParentWBS(parent)
            Loop
            If ranks(i) = 1 Then row.font.bold = True: row.font.size = 13
            If ranks(i) = 1 And row.rowHeight < 21 Then row.rowHeight = 21
            j = ranks(i) - 1: If j > 15 Then j = 15
            row.cells(1, 1).IndentLevel = j
        End If
    Next i
    If Not regroup Then Exit Sub
    If tbl.parent.Outline.SummaryRow <> xlAbove Then tbl.parent.Outline.SummaryRow = xlAbove
    For i = UBound(keys) To 2 Step -1
        If Len(keys(i)) > 0 And ranks(i) < 7 Then
            j = i + 1
            Do While j <= UBound(keys)
                If left$(keys(j), Len(keys(i)) + 1) <> keys(i) & "." Then Exit Do
                j = j + 1
            Loop
            If j > i + 1 Then tbl.DataBodyRange.rows(i + 1).Resize(j - i - 1).EntireRow.group
        End If
    Next i
    If UBound(keys) > 1 Then tbl.DataBodyRange.rows(2).Resize(UBound(keys) - 1).EntireRow.group
    tbl.parent.Outline.ShowLevels RowLevels:=8
End Sub

Private Function Finance_IsExtractSheet(ByVal ws As Worksheet) As Boolean
    Dim tbl As ListObject
    If ws.name Like "Extract Finance #*" Then Finance_IsExtractSheet = True: Exit Function
    For Each tbl In ws.ListObjects
        If tbl.name Like "tbl_FINANCE_EXTRACT_#*" Then Finance_IsExtractSheet = True: Exit Function
    Next tbl
End Function

Public Sub Finance_Extract()
    Dim rows As Collection, expenses As Collection, known As Object, d As PEFinanceRow, e As PEFinanceRow, ws As Worksheet, existing As Worksheet, tbl As ListObject
    Dim serial As Long, i As Long, col As Long, stamp As Date, output() As Variant, fields As Variant, headers(1 To 1, 1 To 8) As Variant
    Dim graph As Object, scopeName As String, scope As Long
    Dim sheetName As String, tableName As String, available As Boolean, oldEvents As Boolean, oldAlerts As Boolean, errorNumber As Long, description As String, currencyCode As String
    If Not WorkbookSchema_UserActionAllowed() Then Exit Sub
    oldEvents = Application.EnableEvents: oldAlerts = Application.DisplayAlerts
    On Error GoTo Failed
    Finance_Flush
    Set graph = Finance_DomainSnapshot(): scope = CLng(graph("META")(1)("SCOPE_ID"))
    Finance_DomainViews graph, scope, rows, expenses, True, True
    Finance_Validate rows, expenses
    Set known = CreateObject("Scripting.Dictionary")
    currencyCode = CStr(graph("META")(1)("CURRENCY"))
    scopeName = Finance_Text("GLOBAL")
    If scope <> 0 Then scopeName = CStr(Finance_DomainIndex(graph("PROJECTS"), "PROJECT_ID")(CStr(scope))("NAME"))
    For Each d In rows: known.Add Finance_Reference(d), d: Next d
    Finance_MigrateExtractCounter
    serial = CLng(Val(CStr(ThisWorkbook.Worksheets("FINANCE").Range(FIN_EXTRACT_COUNTER).Value2)))
    Do
        serial = serial + 1: available = True
        sheetName = "Extract Finance " & CStr(serial): tableName = "tbl_FINANCE_EXTRACT_" & CStr(serial)
        For Each existing In ThisWorkbook.Worksheets
            If StrComp(existing.name, sheetName, vbTextCompare) = 0 Then available = False
            For Each tbl In existing.ListObjects
                If StrComp(tbl.name, tableName, vbTextCompare) = 0 Then available = False
            Next tbl
        Next existing
    Loop Until available
    Application.EnableEvents = False
    Set ws = ThisWorkbook.Worksheets.Add(after:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
    ws.name = sheetName: stamp = Now
    ws.Range("A1").Value2 = Finance_Text("EXTRACT_AT") & " " & Format$(stamp, Settings_GetDateTimeNumberFormat()) & " | " & scopeName & " | " & currencyCode
    fields = Array("PROJECT_NAME", "WBS", "TASK_NAME", "STATE", "DATE", "AMOUNT", "CURRENCY_CODE", "COMMENT")
    For col = 0 To UBound(fields)
        If fields(col) = "CURRENCY_CODE" Then
            headers(1, col + 1) = Finance_Text("CURRENCY_CODE")
        ElseIf fields(col) = "PROJECT_NAME" Then
            headers(1, col + 1) = Finance_Text("PROJECT_NAME")
        ElseIf right$(CStr(fields(col)), 3) = "_ID" Then
            headers(1, col + 1) = Finance_Text("ID_" & CStr(fields(col)))
        Else
            headers(1, col + 1) = Finance_ColumnTitle(CStr(fields(col)), Settings_GetOwnerLanguage("FINANCE"))
        End If
    Next col
    ws.Range("A2:H2").Value2 = headers
    Set tbl = ws.ListObjects.Add(xlSrcRange, ws.Range("A2:H2").Resize(WorksheetFunction.Max(2, expenses.count + 1)), , xlYes)
    tbl.name = tableName: tbl.TableStyle = ThisWorkbook.Worksheets("CONSTRAINTS").ListObjects("tbl_CONSTRAINTS").TableStyle.name
    If expenses.count > 0 Then
        ReDim output(1 To expenses.count, 1 To 8)
        For Each e In expenses
            i = i + 1: Set d = known(Finance_Reference(e))
            output(i, 1) = "'" & CStr(e("PROJECT_NAME")): output(i, 2) = "'" & CStr(e("WBS"))
            output(i, 3) = "'" & CStr(d("TASK_NAME")): output(i, 4) = "'" & Finance_Text(CStr(e("STATE")))
            output(i, 5) = CDbl(e("DATE")): output(i, 6) = CCur(e("AMOUNT"))
            output(i, 7) = "'" & currencyCode: output(i, 8) = "'" & CStr(e("COMMENT"))
        Next e
        Finance_AssignBody tbl, output, True
        tbl.ListColumns(5).DataBodyRange.numberFormat = Settings_GetDateNumberFormat()
        tbl.ListColumns(6).DataBodyRange.numberFormat = Finance_MoneyFormat()
    Else
        If tbl.ListRows.count > 0 Then tbl.ListRows(1).Delete
    End If
    ws.columns("A:H").ColumnWidth = 20: ws.columns("A").ColumnWidth = 32: ws.columns("C").ColumnWidth = 38: ws.columns("H").ColumnWidth = 50
    tbl.HeaderRowRange.WrapText = True: tbl.HeaderRowRange.rowHeight = 30
    ws.Unprotect
    ThisWorkbook.Worksheets("FINANCE").Range(FIN_EXTRACT_COUNTER).Value2 = serial
    Application.EnableEvents = oldEvents
    ws.Activate
    Exit Sub
Failed:
    errorNumber = Err.number: description = Err.description
    On Error Resume Next
    If Not ws Is Nothing Then Application.DisplayAlerts = False: ws.Delete
    Application.DisplayAlerts = oldAlerts: Application.EnableEvents = oldEvents
    On Error GoTo 0
    Err.Raise errorNumber, "Finance_Extract", description
End Sub

Private Sub Finance_MigrateExtractCounter()
    Dim ws As Worksheet, counter As Range
    Set ws = ThisWorkbook.Worksheets("FINANCE"): Set counter = ws.Range(FIN_EXTRACT_COUNTER)
    If Len(CStr(counter.Value2)) = 0 Then
        counter.Value2 = 0
        If Finance_Table(FIN_EXPENSES).HeaderRowRange.column <> 18 Then
            If Len(CStr(ws.Range("S3").Value2)) > 0 And IsNumeric(ws.Range("S3").Value2) Then
                counter.Value2 = CLng(ws.Range("S3").Value2): ws.Range("S3").ClearContents
            End If
        End If
    End If
    counter.numberFormat = ";;;": counter.EntireColumn.hidden = True
End Sub

Public Function Finance_ColumnTitle(ByVal key As String, ByVal language As String, Optional ByVal mode As String = "") As String
    If mode = "" Then mode = Settings_FinanceVocabulary()
    Finance_ColumnTitle = TextCatalog_Get("FINANCE.COL." & key & "." & mode, language)
End Function

Private Function Finance_VisibleTitle(ByVal key As String, Optional ByVal tableName As String = FIN_TABLE) As String
    If FinanceStore_Exists(ThisWorkbook) Then
        If Finance_ScopeId() = 0 Then
            If tableName = FIN_EXPENSES And key = "AMOUNT" Then Finance_VisibleTitle = Finance_Text("GLOBAL.COL.AMOUNT"): Exit Function
            Select Case key
                Case "BUDGET", "ACTUAL", "ETC", "EAC", "VARIANCE"
                    Finance_VisibleTitle = Finance_Text("GLOBAL.COL." & key): Exit Function
            End Select
        End If
    End If
    Finance_VisibleTitle = Finance_ColumnTitle(key, Settings_GetOwnerLanguage("FINANCE"))
End Function

Private Function Finance_ScopeCaption() As String
    Dim graph As Object, scope As Long
    Set graph = Finance_PresentationDomain(): scope = CLng(graph("META")(1)("SCOPE_ID"))
    If graph("PROJECTS").count = 1 Then
        Finance_ScopeCaption = "GLOBAL"
    ElseIf scope = 0 Then
        Finance_ScopeCaption = Finance_Text("GLOBAL")
    Else
        Finance_ScopeCaption = Finance_Text("SCOPE_TOTAL") & " - " & CStr(Finance_DomainIndex(graph("PROJECTS"), "PROJECT_ID")(CStr(scope))("NAME"))
    End If
End Function

Public Function Finance_Map(ByVal tbl As ListObject) As Object
    Dim result As Object, aliases As Object, key As Variant, lang As Variant, mode As Variant, col As listColumn, title As String
    Dim perf As clsPerfScope
    Set perf = Profiler_BeginScope("Finance.View.Map." & tbl.name, "Finance")
    Set result = CreateObject("Scripting.Dictionary")
    Set aliases = CreateObject("Scripting.Dictionary"): aliases.CompareMode = vbTextCompare
    For Each key In Finance_Keys(tbl.name)
        For Each lang In Array("FR", "EN")
            For Each mode In Array("BEGINNER", "EXPERT")
                title = Finance_ColumnTitle(CStr(key), CStr(lang), CStr(mode))
                If aliases.exists(title) Then
                    If aliases(title) <> CStr(key) Then Err.Raise 5, "Finance", "Ambiguous Finance header"
                Else
                    aliases.Add title, CStr(key)
                End If
            Next mode
        Next lang
    Next key
    If tbl.name = FIN_TABLE Then
        For Each key In Array("BUDGET", "ACTUAL", "ETC", "EAC", "VARIANCE")
            For Each lang In Array("FR", "EN")
                title = TextCatalog_Get("FINANCE.GLOBAL.COL." & CStr(key), CStr(lang))
                If aliases.exists(title) Then
                    If CStr(aliases(title)) <> CStr(key) Then Err.Raise 5, "Finance", "Ambiguous Finance header"
                Else
                    aliases.Add title, CStr(key)
                End If
            Next lang
        Next key
        For Each lang In Array("FR", "EN")
            For Each mode In Array("BEGINNER", "EXPERT")
                aliases(Finance_ColumnTitle("TASK_TYPE", CStr(lang), CStr(mode))) = "TASK_TYPE"
            Next mode
        Next lang
        aliases("Task ID") = "TASK_ID"
        aliases("Coût constaté (AC)") = "ACTUAL"
        aliases("ETC - Estimation pour achever") = "ETC"
        aliases("EAC - Estimation à terminaison") = "EAC"
        aliases("Expected overspend") = "VARIANCE"
        aliases("Dépassement prévu") = "VARIANCE"
    ElseIf tbl.name = FIN_EXPENSES Then
        For Each lang In Array("FR", "EN")
            title = TextCatalog_Get("FINANCE.GLOBAL.COL.AMOUNT", CStr(lang))
            aliases(title) = "AMOUNT"
        Next lang
    End If
    For Each col In tbl.ListColumns
        If Not aliases.exists(col.name) Then Err.Raise 5, "Finance", "Unknown Finance header"
        key = aliases(col.name)
        If result.exists(key) Then Err.Raise 5, "Finance", "Duplicate Finance column"
        result.Add CStr(key), col.Index
    Next col
    For Each key In Finance_Keys(tbl.name)
        If Not result.exists(key) Then
            If key <> "OTHERS" And key <> "RELATION" And key <> "PROJECT_NAME" Then Err.Raise 5, "Finance", "Missing Finance columns"
        End If
    Next key
    Set Finance_Map = result
End Function

Private Sub Finance_OrderColumns(ByVal tbl As ListObject)
    Dim order As Variant, map As Object, rows As Collection, widths() As Double, i As Long, changed As Boolean, retired As Range
    order = Array("TASK_ID", "WBS", "TASK_NAME", "PROJECT", "STATE", "PROGRESS", "START", "FINISH", "BUDGET", "PROFILE", "ACTUAL", "ETC", "EAC", "VARIANCE", "DIAGNOSTIC", "OTHERS", "RELATION")
    Set map = Finance_Map(tbl)
    changed = map.exists("TASK_TYPE")
    For i = 0 To UBound(order): If CLng(map(order(i))) <> i + 1 Then changed = True
    Next i
    If Not changed Then Exit Sub
    Set rows = Finance_Read(tbl): ReDim widths(0 To UBound(order))
    For i = 0 To UBound(order): widths(i) = tbl.ListColumns(CLng(map(order(i)))).Range.ColumnWidth: Next i
    If map.exists("TASK_TYPE") Then
        WBS_Onboarding_ClearCellComments tbl.HeaderRowRange.cells(1, CLng(map("TASK_TYPE"))).offset(-1)
        tbl.HeaderRowRange.cells(1, CLng(map("TASK_TYPE"))).offset(-1).ClearContents
        tbl.ListColumns(CLng(map("TASK_TYPE"))).Range.EntireColumn.hidden = False
        Set retired = tbl.ListColumns(tbl.ListColumns.count).Range
        WBS_Onboarding_ClearCellComments retired.cells(1, 1).offset(-1)
        retired.cells(1, 1).offset(-1).ClearContents
        tbl.Resize tbl.Range.Resize(, UBound(order) + 1)
        retired.Clear
    End If
    For i = 0 To UBound(order): tbl.ListColumns(i + 1).name = "__FIN_ORDER_" & CStr(i): Next i
    For i = 0 To UBound(order)
        tbl.ListColumns(i + 1).name = Finance_ColumnTitle(CStr(order(i)), Settings_GetOwnerLanguage("FINANCE"))
        tbl.ListColumns(i + 1).Range.EntireColumn.ColumnWidth = widths(i)
    Next i
    Finance_Write tbl, rows
End Sub

Public Function Finance_OutputKeys(ByVal name As String) As Variant
    If name = "tbl_FINANCE_MONTHS" Then
        Finance_OutputKeys = Array("MONTH", "SERIES1", "SERIES3", "SERIES2", "SERIES4")
    Else
        Finance_OutputKeys = Array("SCOPE", "TOTAL_BUDGET", "TOTAL_ACTUAL", "TOTAL_ETC", "TOTAL_EAC", "TOTAL_VARIANCE", "COUNT", "KNOWN_BUDGET", "KNOWN_ACTUAL", "KNOWN_ETC", "KNOWN_EAC", "UNALLOCATED_AMOUNT", "UNALLOCATED_COUNT", "ALLOCATED_AMOUNT")
    End If
End Function

Private Sub Finance_ColumnHelp(ByVal tbl As ListObject)
    Dim keys As Variant, map As Object, key As Variant, cell As Range, text As String, status As String, language As String, i As Long, col As Long
    Dim modern As Boolean, exists As Boolean, actual As String, title As String
    language = Settings_GetOwnerLanguage("FINANCE")
    modern = ExcelCompatibility_SupportsThreadedComments()
    If tbl.name = FIN_TABLE Then
        keys = Finance_Keys(FIN_TABLE): Set map = Finance_Map(tbl)
    Else
        keys = Finance_OutputKeys(tbl.name)
    End If
    For i = 0 To UBound(keys)
        key = keys(i)
        If map Is Nothing Then col = i + 1 Else col = CLng(map(key))
        Set cell = tbl.HeaderRowRange.cells(1, col).offset(-1)
        status = "WBS.ONBOARDING.CALCULATED"
        If tbl.name = FIN_TABLE And Finance_ScopeId() <> 0 Then
            If key = "BUDGET" Or key = "ETC" Then status = "WBS.ONBOARDING.OPTIONAL"
            If key = "PROFILE" Then status = "WBS.ONBOARDING.REQUIRED"
        End If
        cell.Value2 = TextCatalog_Get(status, language)
        cell.font.name = tbl.HeaderRowRange.font.name: cell.font.size = 9: cell.font.bold = False
        cell.font.color = RGB(90, 90, 90): cell.HorizontalAlignment = xlCenter: cell.WrapText = False
        title = CStr(tbl.HeaderRowRange.cells(1, col).Value2)
        text = title & vbLf & vbLf & Replace$(Finance_Text("HELP." & CStr(key)), ". ", "." & vbLf & "- ")
        WBS_Onboarding_ApplyHelpCell cell, text, Len(title)
    Next i
    If tbl.HeaderRowRange.offset(-1).rowHeight < 18 Then tbl.HeaderRowRange.offset(-1).rowHeight = 18
End Sub

Private Sub Finance_ClearColumnHelp(ByVal tbl As ListObject)
    Dim cell As Range, exists As Boolean, text As String
    For Each cell In tbl.HeaderRowRange.offset(-1).cells
        text = WBS_Onboarding_ThreadedCommentText(cell, exists)
        If exists Then WBS_Onboarding_ClearCellComments cell: cell.ClearContents
        If Not cell.comment Is Nothing Then
            If cell.comment.Author = "ProjectEngine" Then WBS_Onboarding_ClearCellComments cell: cell.ClearContents
        End If
    Next cell
End Sub

Private Sub Finance_LocalizeValues(ByVal tbl As ListObject, ByVal map As Object)
    Dim key As Variant, codes As Variant, code As Variant, lang As Variant, translations As Object, field As listColumn
    Dim values As Variant, r As Long, oldValue As String, newValue As String, changed As Boolean, replacement As Variant
    If tbl.DataBodyRange Is Nothing Then Exit Sub
    For Each key In Array("STATE", "PROFILE", "DIAGNOSTIC")
        If Not map.exists(CStr(key)) Then GoTo NextField
        Set translations = CreateObject("Scripting.Dictionary"): translations.CompareMode = vbTextCompare
        Select Case CStr(key)
            Case "STATE": codes = Array("ACTIVE", "HISTORICAL", "MIXED")
            Case "PROFILE": codes = Array("PROFILE_START", "PROFILE_END", "PROFILE_LINEAR", "MIXED")
            Case Else: codes = Array("UNALLOCATED", "OVER_BUDGET", "GLOBAL_COVERAGE", "MISSING")
        End Select
        For Each code In codes
            For Each lang In Array("FR", "EN")
                translations(TextCatalog_Get("FINANCE." & CStr(code), CStr(lang))) = Finance_Text(CStr(code))
            Next lang
            If key = "STATE" Then translations(CStr(code)) = Finance_Text(CStr(code))
            If key = "PROFILE" Then
                If left$(CStr(code), 8) = "PROFILE_" Then translations(Mid$(CStr(code), 9)) = Finance_Text(CStr(code)) Else translations(CStr(code)) = Finance_Text(CStr(code))
            End If
        Next code
        Set field = tbl.ListColumns(CLng(map(key)))
        If field.DataBodyRange.rows.count = 1 Then
            ReDim values(1 To 1, 1 To 1): values(1, 1) = field.DataBodyRange.Value2
        Else
            values = field.DataBodyRange.Value2
        End If
        changed = False
        For r = 1 To UBound(values, 1)
            oldValue = CStr(values(r, 1)): newValue = oldValue
            If key = "DIAGNOSTIC" Then
                For Each replacement In translations.keys
                    newValue = Replace$(newValue, CStr(replacement), CStr(translations(replacement)))
                Next replacement
            ElseIf translations.exists(oldValue) Then
                newValue = translations(oldValue)
            End If
            If newValue <> oldValue Then values(r, 1) = newValue: changed = True
        Next r
        If changed Then field.DataBodyRange.Value2 = values
NextField:
    Next key
End Sub

Public Function Finance_Table(ByVal name As String, Optional ByVal wb As Workbook) As ListObject
    If wb Is Nothing Then Set wb = ThisWorkbook
    If name = "tbl_FINANCE_MONTHS" Or name = "tbl_FINANCE_TOTALS" Then
        On Error Resume Next
        Set Finance_Table = wb.Worksheets(FIN_SUMMARY).ListObjects(name)
        On Error GoTo 0
        If Not Finance_Table Is Nothing Then Exit Function
    End If
    Set Finance_Table = wb.Worksheets("FINANCE").ListObjects(name)
End Function

Private Function Finance_SummarySheet() As Worksheet
    Dim ws As Worksheet
    On Error Resume Next: Set ws = ThisWorkbook.Worksheets(FIN_SUMMARY): On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(after:=ThisWorkbook.Worksheets("FINANCE"))
        ws.name = FIN_SUMMARY: ws.Tab.color = ThisWorkbook.Worksheets("FINANCE").Tab.color
        ws.rows(1).rowHeight = 32: ws.rows(2).rowHeight = 8
        ws.columns("A:N").ColumnWidth = 18: ws.columns("A").ColumnWidth = 32
    End If
    Set Finance_SummarySheet = ws
End Function

Private Function Finance_TotalsAddress(Optional ByVal monthCount As Long = -1) As String
    Dim tbl As ListObject, row As Long
    row = 29
    If monthCount >= 0 Then Finance_TotalsAddress = "A" & CStr(29 + monthCount): Exit Function
    On Error Resume Next: Set tbl = Finance_Table("tbl_FINANCE_MONTHS"): On Error GoTo 0
    If Not tbl Is Nothing Then row = tbl.Range.row + tbl.Range.rows.count + 3
    If row < 29 Then row = 29
    Finance_TotalsAddress = "A" & CStr(row)
End Function

Public Function Finance_Read(ByVal tbl As ListObject) As Collection
    Dim rows As New Collection, keys As Variant, data As Variant, map As Object, r As Long, c As Long, item As PEFinanceRow
    mReads = mReads + 1
    Set map = Finance_Map(tbl): keys = Finance_Keys(tbl.name)
    If tbl.ListRows.count > 0 And Not tbl.DataBodyRange Is Nothing Then
        data = tbl.DataBodyRange.Value2
        For r = 1 To UBound(data, 1)
            Set item = New PEFinanceRow
            item.Initialize keys
            For c = LBound(keys) To UBound(keys)
                If map.exists(CStr(keys(c))) Then item(CStr(keys(c))) = data(r, map(CStr(keys(c))))
            Next c
            If item.exists("STATE") Then item("STATE") = Finance_State(CStr(item("STATE")))
            If item.exists("PROFILE") Then item("PROFILE") = Finance_Profile(CStr(item("PROFILE")))
            rows.Add item
        Next r
    End If
    Set Finance_Read = rows
End Function

Public Function Finance_NewRow(ByVal tableName As String) As Object
    Dim d As New PEFinanceRow
    d.Initialize Finance_Keys(tableName)
    If tableName = FIN_TABLE Then d("PROFILE") = "LINEAR"
    Set Finance_NewRow = d
End Function

Private Function Finance_RowEquals(ByVal left As Object, ByVal right As Object, ByVal keys As Variant) As Boolean
    Dim key As Variant
    For Each key In keys
        If CStr(left(key)) <> CStr(right(key)) Then Exit Function
    Next key
    Finance_RowEquals = True
End Function

Private Function Finance_RowsEqual(ByVal left As Collection, ByVal right As Collection, ByVal tableName As String) As Boolean
    Dim i As Long, keys As Variant
    If left.count <> right.count Then Exit Function
    keys = Finance_Keys(tableName)
    For i = 1 To left.count
        If Not Finance_RowEquals(left(i), right(i), keys) Then Exit Function
    Next i
    Finance_RowsEqual = True
End Function

Private Function Finance_StableSubset(ByVal rows As Collection, ByVal original As Collection, ByVal tableName As String) As Boolean
    Dim row As Object, wanted As Long, keys As Variant
    If rows.count > original.count Then Exit Function
    wanted = 1: keys = Finance_Keys(tableName)
    For Each row In original
        If wanted > rows.count Then Exit For
        If Finance_Reference(row) = Finance_Reference(rows(wanted)) Then
            If Not Finance_RowEquals(row, rows(wanted), keys) Then Exit Function
            wanted = wanted + 1
        End If
    Next row
    Finance_StableSubset = (wanted = rows.count + 1)
End Function

Public Function Finance_SelectedRows(ByVal original As Collection, ByVal chosen As Object, ByVal activities As Collection) As Collection
    Dim rows As New Collection, included As Object, row As Object, reference As String
    Set included = CreateObject("Scripting.Dictionary")
    For Each row In original
        reference = Finance_Reference(row)
        If chosen.exists(reference) Then rows.Add chosen(reference): included.Add reference, True
    Next row
    For Each row In activities
        reference = Finance_Reference(row)
        If chosen.exists(reference) And Not included.exists(reference) Then
            rows.Add chosen(reference): included.Add reference, True
        End If
    Next row
    If rows.count <> chosen.count Then Err.Raise 5, "Finance_SelectedRows", Finance_Text("SOURCE_CHANGED")
    Set Finance_SelectedRows = rows
End Function

Public Sub Finance_Write(ByVal tbl As ListObject, ByVal rows As Collection)
    Dim keys As Variant, map As Object, a() As Variant, r As Long, c As Long, value As Variant, row As PEFinanceRow
    Dim activeCaption As String, historyCaption As String
    Dim startCaption As String, endCaption As String, linearCaption As String
    Dim oldEvents As Boolean, errorNumber As Long, description As String
    Dim original As Collection, previous As Range, tail As Range, i As Long, wanted As Long, last As Long
    Dim previousRow As Long, previousColumn As Long, previousCount As Long, previousWidth As Long
    Dim perf As clsPerfScope
    Set perf = Profiler_BeginScope("Finance.View.Write." & tbl.name, "Finance")
    oldEvents = Application.EnableEvents
    On Error GoTo Failed
    Application.EnableEvents = False
    Set map = Finance_Map(tbl): keys = Finance_Keys(tbl.name)
    If Not tbl.DataBodyRange Is Nothing Then
        Set previous = tbl.DataBodyRange
        previousRow = previous.row: previousColumn = previous.column
        previousCount = previous.rows.count: previousWidth = previous.columns.count
    End If
    If rows.count < tbl.ListRows.count Then
        Set original = Finance_Read(tbl)
        If Finance_StableSubset(rows, original, tbl.name) Then
            i = original.count: wanted = rows.count
            Do While i > 0
                If wanted > 0 Then
                    If Finance_Reference(original(i)) = Finance_Reference(rows(wanted)) Then
                        i = i - 1: wanted = wanted - 1
                        GoTo NextRemoval
                    End If
                End If
                last = i
                Do While i > 0
                    If wanted > 0 Then
                        If Finance_Reference(original(i)) = Finance_Reference(rows(wanted)) Then Exit Do
                    End If
                    i = i - 1
                Loop
                tbl.ListRows(i + 1).Range.Resize(last - i, tbl.ListColumns.count).Delete Shift:=xlShiftUp
NextRemoval:
            Loop
            If tbl.ListRows.count <> rows.count Then Err.Raise 5, "Finance_Write", "Finance deletion count mismatch"
            GoTo CleanExit
        End If
    End If
    If Not tbl.DataBodyRange Is Nothing Then
        If rows.count <> previousCount Then tbl.DataBodyRange.ClearContents
    End If
    If rows.count = 0 Then
        tbl.Resize tbl.HeaderRowRange.Resize(2, tbl.ListColumns.count)
        If tbl.ListRows.count > 0 Then tbl.ListRows(1).Delete
        GoTo CleanExit
    End If
    activeCaption = Finance_Text("ACTIVE"): historyCaption = Finance_Text("HISTORICAL")
    startCaption = Finance_Text("PROFILE_START"): endCaption = Finance_Text("PROFILE_END"): linearCaption = Finance_Text("PROFILE_LINEAR")
    ReDim a(1 To rows.count, 1 To tbl.ListColumns.count)
    For Each row In rows
        r = r + 1
        For c = LBound(keys) To UBound(keys)
            If Not map.exists(CStr(keys(c))) Then GoTo NextWriteColumn
            value = row(CStr(keys(c)))
            If CStr(keys(c)) = "STATE" Then
                If CStr(value) = "ACTIVE" Then value = activeCaption
                If CStr(value) = "HISTORICAL" Then value = historyCaption
                If CStr(value) = "MIXED" Then value = Finance_Text("MIXED")
            End If
            If CStr(keys(c)) = "PROFILE" Then
                Select Case CStr(value)
                    Case "START": value = startCaption
                    Case "END": value = endCaption
                    Case "LINEAR": value = linearCaption
                    Case "MIXED": value = Finance_Text("MIXED")
                End Select
            End If
            If VarType(value) = vbString Then
                If Len(value) > 0 Then value = "'" & value
            End If
            a(r, map(CStr(keys(c)))) = value
NextWriteColumn:
        Next c
    Next row
    Finance_AssignBody tbl, a, True
CleanExit:
    If errorNumber = 0 And Not previous Is Nothing Then
        If rows.count < previousCount Then
            Set tail = tbl.parent.cells(previousRow + rows.count, previousColumn).Resize(previousCount - rows.count, previousWidth)
            If Application.WorksheetFunction.CountA(tail) = 0 Then tail.Clear
        End If
    End If
    Application.EnableEvents = oldEvents
    On Error GoTo 0
    If errorNumber <> 0 Then Err.Raise errorNumber, "Finance_Write", description
    Exit Sub
Failed:
    errorNumber = Err.number: description = Err.description
    Resume CleanExit
End Sub

Private Sub Finance_AssignBody(ByVal tbl As ListObject, ByVal data As Variant, Optional ByVal formulaSafe As Boolean = False)
    Dim target As Range, count As Long, previous As Range
    If Not tbl.DataBodyRange Is Nothing Then Set previous = tbl.DataBodyRange
    count = UBound(data, 1)
    Set target = tbl.HeaderRowRange.offset(1).Resize(count, tbl.ListColumns.count)
    tbl.Resize tbl.HeaderRowRange.Resize(count + 1, tbl.ListColumns.count)
    ' Write the declared body first: a blank insertion row is not a ListRow yet.
    ' ListRows.Add can shift neighbouring tables, even after Resize.
    If formulaSafe Then target.formula = data Else target.Value2 = data
    If tbl.ListRows.count <> count Then tbl.Resize tbl.HeaderRowRange.Resize(count + 1, tbl.ListColumns.count)
    If tbl.DataBodyRange Is Nothing Or tbl.ListRows.count <> count Then Err.Raise 5, "Finance_AssignBody", "Finance row materialization failed"
    If Not previous Is Nothing Then
        If previous.rows.count > count Then previous.rows(count + 1).Resize(previous.rows.count - count).Clear
    End If
End Sub

Public Function Finance_Money(ByVal value As Variant, Optional ByVal allowCredit As Boolean = False) As Variant
    Dim amount As Currency
    If isError(value) Then GoTo InvalidAmount
    If IsEmpty(value) Or Len(Trim$(CStr(value))) = 0 Then Finance_Money = Empty: Exit Function
    If Not IsNumeric(value) Then GoTo InvalidAmount
    If CDec(value) < 0 And Not allowCredit Then GoTo InvalidAmount
    If Abs(CDec(value)) > CDec(1000000000000#) Then GoTo InvalidAmount
    amount = CCur(value)
    If CDec(amount) * 100 <> Fix(CDec(amount) * 100) Or CDec(amount) <> CDec(value) Then GoTo InvalidAmount
    Finance_Money = amount
    Exit Function
InvalidAmount:
    Err.Raise 5, "Finance", Finance_Text(IIf(allowCredit, "INVALID_EXPENSE", "INVALID_MONEY"))
End Function

Public Function Finance_ParseAmount(ByVal text As String) As Variant
    Dim value As String, separator As String
    value = Trim$(Replace$(Replace$(text, ChrW$(160), " "), ChrW$(8239), " "))
    If Len(value) = 0 Then Finance_ParseAmount = Empty: Exit Function
    If InStr(value, ".") > 0 And InStr(value, ",") > 0 Then Err.Raise 5, "Finance", Finance_Text("INVALID_EXPENSE")
    separator = CStr(Application.International(xlDecimalSeparator))
    value = Replace$(value, " ", "")
    value = Replace$(Replace$(value, ".", separator), ",", separator)
    Finance_ParseAmount = Finance_Money(value, True)
End Function

Public Function Finance_ParseExpenseDate(ByVal value As Variant) As Double
    Dim text As String, parts As Variant, result As Date
    On Error GoTo invalid
    If isError(value) Or IsEmpty(value) Then GoTo invalid
    If VarType(value) <> vbString And IsNumeric(value) Then
        If Finance_UsableDate(value) Then Finance_ParseExpenseDate = CDbl(value): Exit Function
        GoTo invalid
    End If
    text = Trim$(CStr(value))
    If text Like "####-##-##" Then
        parts = Split(text, "-"): result = DateSerial(CInt(parts(0)), CInt(parts(1)), CInt(parts(2)))
        If Format$(result, "yyyy-mm-dd") <> text Then GoTo invalid
    Else
        If Not IsDate(text) Then GoTo invalid
        result = dateValue(CDate(text))
    End If
    If Not Finance_UsableDate(CDbl(result)) Then GoTo invalid
    Finance_ParseExpenseDate = CDbl(result): Exit Function
invalid:
    Err.Raise 5, "Finance", Finance_Text("INVALID_DATE")
End Function

Private Function Finance_ExpenseInputIndex(ByVal rows As Collection) As Object
    Dim Index As Object, row As PEFinanceRow, key As String
    Set Index = CreateObject("Scripting.Dictionary"): Index.CompareMode = vbTextCompare
    For Each row In rows
        Set Index(Finance_Reference(row)) = row
        Set Index(FinanceUI_Caption(row)) = row
        key = CStr(row("WBS"))
        If Not Index.exists(key) Then
            Index.Add key, row
        ElseIf CStr(row("STATE")) = "ACTIVE" Then
            Set Index(key) = row
        End If
    Next row
    Set Finance_ExpenseInputIndex = Index
End Function

Private Function Finance_ExpenseInputRow(ByVal Index As Object, ByVal reference As String, ByVal recordedDate As Variant, ByVal amount As Variant, ByVal comment As String) As Object
    Dim source As PEFinanceRow, row As PEFinanceRow, money As Variant
    reference = Trim$(reference)
    If Not Index.exists(reference) Then Err.Raise 5, "Finance", Replace$(Finance_Text("UNKNOWN_WBS"), "{WBS}", reference)
    Set source = Index(reference)
    money = Finance_ParseAmount(CStr(amount))
    If Not Finance_Has(money) Then Err.Raise 5, "Finance", Finance_Text("INVALID_EXPENSE")
    If Len(comment) > 32767 Then Err.Raise 5, "Finance", Finance_Text("COMMENT_LIMIT")
    Set row = Finance_NewRow(FIN_EXPENSES)
    row("WBS") = source("WBS"): row("STATE") = source("STATE")
    row("DATE") = Finance_ParseExpenseDate(recordedDate): row("AMOUNT") = money: row("COMMENT") = comment
    row("PROJECT_NAME") = source("PROJECT_NAME")
    row.Add "FINANCIAL_ID", source("FINANCIAL_ID"): row.Add "TASK_ID", source("TASK_ID")
    Set Finance_ExpenseInputRow = row
End Function

Public Function Finance_ExpenseDraftRow(ByVal rows As Collection, ByVal reference As String, ByVal recordedDate As Variant, ByVal amount As Variant, ByVal comment As String) As Object
    Set Finance_ExpenseDraftRow = Finance_ExpenseInputRow(Finance_ExpenseInputIndex(rows), reference, recordedDate, amount, comment)
End Function

Public Function Finance_BulkExpenseHeader(ByVal data As Variant, ByVal row As Long) As Boolean
    Dim c As Long
    c = LBound(data, 2)
    If UBound(data, 2) - c <> 3 Then Err.Raise 5, "Finance", Finance_Text("BULK_FOUR_COLUMNS")
    Finance_BulkExpenseHeader = Finance_ExpenseHeaderTitles(CStr(data(row, c)), CStr(data(row, c + 1)), CStr(data(row, c + 2)), CStr(data(row, c + 3)))
End Function

Private Function Finance_ExpenseHeaderTitles(ByVal wbs As String, ByVal recordedDate As String, ByVal amount As String, ByVal comment As String) As Boolean
    Dim title As String
    If UCase$(Trim$(wbs)) <> "WBS" Then Exit Function
    If LCase$(Trim$(recordedDate)) <> "date" Then Err.Raise 5, "Finance", Finance_Text("BULK_FOUR_COLUMNS")
    title = LCase$(Trim$(amount))
    If title <> "amount" And title <> "montant" Then Err.Raise 5, "Finance", Finance_Text("BULK_FOUR_COLUMNS")
    title = LCase$(Trim$(comment))
    If title <> "comment" And title <> "commentaire" Then Err.Raise 5, "Finance", Finance_Text("BULK_FOUR_COLUMNS")
    Finance_ExpenseHeaderTitles = True
End Function

Public Function Finance_BulkExpenseRows(ByVal rows As Collection, ByVal data As Variant) As Collection
    Dim result As New Collection, Index As Object, row As PEFinanceRow, r As Long, c As Long, firstColumn As Long, blank As Boolean, details As String, rowNumber As Long
    On Error GoTo Failed
    If Not IsArray(data) Then Err.Raise 5, , Finance_Text("BULK_FOUR_COLUMNS")
    firstColumn = LBound(data, 2)
    If UBound(data, 2) - firstColumn <> 3 Then Err.Raise 5, , Finance_Text("BULK_FOUR_COLUMNS")
    Set Index = Finance_ExpenseInputIndex(rows)
    For r = LBound(data, 1) To UBound(data, 1)
        rowNumber = r - LBound(data, 1) + 1
        blank = True
        For c = firstColumn To firstColumn + 3
            If Len(Trim$(CStr(data(r, c)))) > 0 Then blank = False
        Next c
        If Not blank Then
            If r = LBound(data, 1) Then
                If Finance_BulkExpenseHeader(data, r) Then GoTo NextBulkRow
            End If
            Set row = Finance_ExpenseInputRow(Index, CStr(data(r, firstColumn)), data(r, firstColumn + 1), data(r, firstColumn + 2), CStr(data(r, firstColumn + 3)))
            result.Add row
        End If
NextBulkRow:
    Next r
    Set Finance_BulkExpenseRows = result: Exit Function
Failed:
    details = Err.description
    If rowNumber = 0 Then Err.Raise 5, "Finance", details
    Err.Raise 5, "Finance", TextCatalog_Format("FINANCE.BULK_ROW_ERROR", Settings_GetOwnerLanguage("FINANCE"), TextCatalog_Arguments("Row", CStr(rowNumber), "Details", details))
End Function

Public Function Finance_ConfirmExpenseDates(ByVal rows As Collection, ByVal entries As Collection) As Boolean
    Dim targets As New Collection, Index As Object, entry As PEFinanceRow, source As PEFinanceRow, target As PEFinanceRow
    Set Index = Finance_ExpenseInputIndex(rows)
    For Each entry In entries
        If CDbl(entry("DATE")) > CDbl(Date) Then
            Set source = Index(Finance_Reference(entry))
            Set target = New PEFinanceRow
            target.Initialize Array("TASK_ID", "WBS", "TASK_NAME", "STATE", "DATE")
            target("TASK_ID") = source("TASK_ID"): target("WBS") = source("WBS")
            target("TASK_NAME") = source("TASK_NAME"): target("STATE") = source("STATE"): target("DATE") = entry("DATE")
            targets.Add target
        End If
    Next entry
    If targets.count = 0 Then Finance_ConfirmExpenseDates = True: Exit Function
    frmFinanceResolution.PrepareFuture targets
    frmFinanceResolution.Show vbModal
    Finance_ConfirmExpenseDates = frmFinanceResolution.Accepted
    Unload frmFinanceResolution
End Function

Private Function Finance_State(ByVal value As String) As String
    Dim key As Variant, lang As Variant
    For Each key In Array("ACTIVE", "HISTORICAL", "MIXED")
        If value = key Then Finance_State = CStr(key): Exit Function
        For Each lang In Array("FR", "EN")
            If value = TextCatalog_Get("FINANCE." & CStr(key), CStr(lang)) Then Finance_State = CStr(key): Exit Function
        Next lang
    Next key
    Finance_State = value
End Function

Private Function Finance_Profile(ByVal value As String) As String
    Dim key As Variant, lang As Variant
    value = Trim$(value)
    Select Case UCase$(value)
        Case "START", "END", "LINEAR", "MIXED": Finance_Profile = UCase$(value): Exit Function
    End Select
    For Each lang In Array("FR", "EN")
        If StrComp(value, TextCatalog_Get("FINANCE.MIXED", CStr(lang)), vbTextCompare) = 0 Then Finance_Profile = "MIXED": Exit Function
    Next lang
    For Each key In Array("START", "END", "LINEAR")
        If UCase$(value) = CStr(key) Then Finance_Profile = CStr(key): Exit Function
        For Each lang In Array("FR", "EN")
            If StrComp(value, TextCatalog_Get("FINANCE.PROFILE_" & CStr(key), CStr(lang)), vbTextCompare) = 0 Then Finance_Profile = CStr(key): Exit Function
        Next lang
    Next key
    Finance_Profile = value
End Function

Public Function Finance_Has(ByVal value As Variant) As Boolean
    If isError(value) Then Exit Function
    Finance_Has = Not IsEmpty(value) And Len(CStr(value)) > 0
End Function

Public Function Finance_Reference(ByVal row As Object) As String
    If row.exists("CONTRIBUTION") Then
        If Finance_Has(row("CONTRIBUTION")) Then Finance_Reference = "LINK|" & CStr(row("CONTRIBUTION")): Exit Function
    End If
    Finance_Reference = CStr(row("STATE")) & "|" & CStr(row("WBS"))
End Function

Public Function Finance_RemovalTargets(ByVal original As Collection, ByVal chosen As Object, ByVal expenses As Collection) As Collection
    Dim targets As New Collection, recorded As Object, row As Object, reference As String
    Set recorded = CreateObject("Scripting.Dictionary")
    For Each row In expenses
        recorded(Finance_FinancialReference(row)) = True
    Next row
    For Each row In original
        reference = Finance_Reference(row)
        If Not chosen.exists(reference) Then
            ' A declared zero and even a zero-value ledger entry are real user data.
            If Finance_Has(row("BUDGET")) Or Finance_Has(row("ETC")) Or recorded.exists(Finance_FinancialReference(row)) Then targets.Add row
        End If
    Next row
    Set Finance_RemovalTargets = targets
End Function

Public Function Finance_Activities() As Collection
    Dim tbl As ListObject, map As Object, a As Variant, rows As New Collection, seen As Object, ids As Object, parents As Object
    Dim r As Long, d As Object, wbs As String, raw As String, id As Variant, re As Object, parent As String, key As Variant
    Set tbl = ThisWorkbook.Worksheets("WBS").ListObjects("tbl_WBS")
    Set map = CanonicalIdentity_BuildColumnMap(tbl)
    Set seen = CreateObject("Scripting.Dictionary"): Set ids = CreateObject("Scripting.Dictionary")
    Set parents = CreateObject("Scripting.Dictionary")
    Set re = CreateObject("VBScript.RegExp"): re.Pattern = "^\d+(\.\d+)*$"
    If tbl.DataBodyRange Is Nothing Then Set Finance_Activities = rows: Exit Function
    a = tbl.DataBodyRange.Value2
    For r = 1 To UBound(a, 1)
        id = a(r, map(VTS_COL_ID)): raw = Trim$(CStr(a(r, map(VTS_COL_WBS))))
        If Len(CStr(id)) > 0 Then
            If Not IsNumeric(id) Then Err.Raise 5, "Finance", "Invalid Task ID"
            If CDbl(id) < 1 Or CDbl(id) <> Fix(CDbl(id)) Or CDbl(id) > 2147483647# Then Err.Raise 5, "Finance", "Invalid Task ID"
            If ids.exists(CStr(CLng(id))) Then Err.Raise 5, "Finance", "Duplicate Task ID"
            ids.Add CStr(CLng(id)), True
            If Not re.Test(raw) Then Err.Raise 5, "Finance", "Invalid WBS: " & raw
            wbs = NormalizeWBS(raw)
            If seen.exists(wbs) Then Err.Raise 5, "Finance", "Duplicate WBS: " & wbs
            Set d = Finance_NewRow(FIN_TABLE)
            d("WBS") = wbs: d("TASK_ID") = CLng(id): d("STATE") = "ACTIVE"
            d("TASK_NAME") = a(r, map(VTS_COL_TASK_NAME)): d("PROJECT") = a(r, map(VTS_COL_PROJECT))
            d("PROGRESS") = a(r, map(VTS_COL_PROGRESS_PERCENT))
            d.Add "PARENT", False
            seen.Add wbs, d: rows.Add d
            parent = GetParentWBS(wbs)
            Do While parent <> ""
                parents(parent) = True
                parent = GetParentWBS(parent)
            Loop
        End If
    Next r
    For Each d In rows
        d("PARENT") = parents.exists(CStr(d("WBS")))
    Next d
    Set Finance_Activities = rows
End Function

Public Sub Finance_Validate(ByVal rows As Collection, ByVal expenses As Collection)
    Dim refs As Object, d As PEFinanceRow, e As PEFinanceRow, key As String, w As String, re As Object, v As Variant
    Set refs = CreateObject("Scripting.Dictionary")
    For Each d In rows
        w = CStr(d("WBS"))
        If Not Finance_ValidWBS(w) Then Err.Raise 5, "Finance", "Invalid Finance WBS"
        If d("STATE") <> "ACTIVE" And d("STATE") <> "HISTORICAL" Then Err.Raise 5, "Finance", "Invalid Finance state"
        key = Finance_Reference(d)
        If refs.exists(key) Then Err.Raise 5, "Finance", Finance_Text("HISTORY_COLLISION")
        refs.Add key, True
        v = Finance_Money(d("BUDGET")): v = Finance_Money(d("ETC"))
        If d("PROFILE") <> "START" And d("PROFILE") <> "END" And d("PROFILE") <> "LINEAR" Then Err.Raise 5, "Finance", "Invalid budget profile"
    Next d
    For Each e In expenses
        If Not refs.exists(Finance_Reference(e)) Then Err.Raise 5, "Finance", "Orphan expense"
        If Not IsNumeric(e("DATE")) Then Err.Raise 5, "Finance", Finance_Text("INVALID_DATE")
        If CDbl(e("DATE")) < 1 Or CDbl(e("DATE")) > 2958465# Or CDbl(e("DATE")) <> Fix(CDbl(e("DATE"))) Then Err.Raise 5, "Finance", Finance_Text("INVALID_DATE")
        v = Finance_Money(e("AMOUNT"), True)
        If Not Finance_Has(v) Then Err.Raise 5, "Finance", Finance_Text("INVALID_MONEY")
    Next e
End Sub

' Build-time only. No caller in Workbook_Open or planning workflows.
Public Sub Finance_Install()
    Dim ws As Worksheet, tbl As ListObject, keys As Variant, i As Long, shp As shape
    On Error Resume Next: Set ws = ThisWorkbook.Worksheets("FINANCE"): On Error GoTo 0
    If Not ws Is Nothing Then Err.Raise 5, "Finance", "Finance already installed"
    Set ws = ThisWorkbook.Worksheets.Add(after:=ThisWorkbook.Worksheets("CONSTRAINTS"))
    ws.name = "FINANCE": ws.Tab.ThemeColor = ThisWorkbook.Worksheets("CONSTRAINTS").Tab.ThemeColor
    ws.Tab.TintAndShade = ThisWorkbook.Worksheets("CONSTRAINTS").Tab.TintAndShade
    ws.cells.font.name = "Segoe UI": ws.cells.font.size = 11
    ws.rows(1).rowHeight = 32: ws.rows(2).rowHeight = 8
    ws.Range("R1").Value2 = "Currency": ws.Range("S1").Value2 = "EUR"
    ws.Range("R2").Value2 = "LastCalculatedAt"
    ws.Range("R1:S2").numberFormat = ";;;"
    For Each keys In Array(Finance_Keys(FIN_TABLE), Finance_Keys(FIN_EXPENSES))
        i = 6
        If UBound(keys) = 14 Then
            Set tbl = ws.ListObjects.Add(xlSrcRange, ws.Range("A3:O4"), , xlYes): tbl.name = FIN_TABLE
        Else
            Set tbl = ws.ListObjects.Add(xlSrcRange, ws.Range("R3:V4"), , xlYes): tbl.name = FIN_EXPENSES
        End If
        For i = 0 To UBound(keys): tbl.ListColumns(i + 1).name = Finance_ColumnTitle(CStr(keys(i)), Settings_GetOwnerLanguage("FINANCE")): Next i
        If tbl.ListRows.count > 0 Then tbl.ListRows(1).Delete
        tbl.TableStyle = "TableStyleMedium2"
        tbl.HeaderRowRange.WrapText = True: tbl.HeaderRowRange.rowHeight = 42
    Next keys
    ws.columns("A:P").ColumnWidth = 16: ws.columns("B").ColumnWidth = 32
    ws.columns("P").ColumnWidth = 42
    Finance_Present
    Settings_Initialize
End Sub

Private Function Finance_StartButton(ByVal ws As Worksheet) As shape
    Dim shp As shape
    On Error Resume Next: Set shp = ws.Shapes("peFinanceStart"): On Error GoTo 0
    If shp Is Nothing Then
        Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, 160, 140, 310, 48)
        shp.name = "peFinanceStart"
        shp.fill.ForeColor.RGB = RGB(40, 107, 145): shp.Line.visible = msoFalse
        shp.TextFrame2.TextRange.font.size = 14
        shp.TextFrame2.TextRange.font.fill.ForeColor.RGB = vbWhite
        shp.TextFrame2.VerticalAnchor = msoAnchorMiddle
        shp.TextFrame2.TextRange.ParagraphFormat.alignment = msoAlignCenter
    End If
    shp.Placement = xlFreeFloating
    shp.OnAction = "'" & Replace(ThisWorkbook.name, "'", "''") & "'!Finance_Start"
    Set Finance_StartButton = shp
End Function

Private Sub Finance_StatusCaption(ByVal ws As Worksheet, ByVal text As String)
    Dim shp As shape, anchor As Range
    Set anchor = ws.Range("L1")
    On Error Resume Next: Set shp = ws.Shapes("peFinanceLastRefresh"): On Error GoTo 0
    If shp Is Nothing Then
        Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, anchor.left, anchor.top + 3, 270, 26)
        shp.name = "peFinanceLastRefresh"
    End If
    With shp
        .left = anchor.left: .top = anchor.top + 3: .width = 270: .height = 26
        .Placement = xlMove: .Line.visible = msoFalse
        .fill.ForeColor.RGB = RGB(40, 107, 145): .Adjustments.item(1) = 0.18
        With .TextFrame2
            .MarginLeft = 12: .MarginRight = 12: .MarginTop = 3: .MarginBottom = 3
            .VerticalAnchor = msoAnchorMiddle: .WordWrap = msoFalse
            .TextRange.text = text: .TextRange.font.name = "Segoe UI": .TextRange.font.size = 10
            .TextRange.font.fill.ForeColor.RGB = vbWhite
        End With
    End With
End Sub

Public Sub Finance_Protect(Optional ByVal regroupTotals As Boolean = False)
    Dim ws As Worksheet, tbl As ListObject, map As Object, key As Variant, template As ListObject, name As Variant, i As Long
    Dim autoHeader As Long, autoBody As Long, moneyFormat As String, phase As String, errorNumber As Long, description As String
    Dim tail As Range, tailCell As Range, retired As Boolean, hasComment As Boolean, commentText As String, readOnly As Boolean
    On Error GoTo Failed
    phase = "Input style"
    Set ws = ThisWorkbook.Worksheets("FINANCE"): ws.Unprotect
    ' Historical public entry point: formatting only; input guards are event based.
    Set template = ThisWorkbook.Worksheets("CONSTRAINTS").ListObjects("tbl_CONSTRAINTS")
    autoHeader = template.HeaderRowRange.cells(1, 1).Interior.color
    autoBody = RGB(232, 232, 232)
    If Not template.DataBodyRange Is Nothing Then autoBody = template.DataBodyRange.cells(1, 1).Interior.color
    moneyFormat = Finance_MoneyFormat()
    Set tbl = Finance_Table(FIN_TABLE): Set map = Finance_Map(tbl)
    tbl.TableStyle = template.TableStyle.name: tbl.ShowTableStyleRowStripes = template.ShowTableStyleRowStripes
    tbl.Range.font.name = template.HeaderRowRange.cells(1, 1).font.name
    tbl.Range.font.size = template.HeaderRowRange.cells(1, 1).font.size
    tbl.HeaderRowRange.Interior.color = autoHeader: tbl.HeaderRowRange.font.color = vbBlack
    If FinanceStore_Exists(ThisWorkbook) Then readOnly = (Finance_ScopeId() = 0)
    For Each key In Finance_InputKeys()
        If Not readOnly Then
            tbl.ListColumns(CLng(map(key))).Range.cells(1, 1).Interior.Pattern = xlPatternNone
            tbl.ListColumns(CLng(map(key))).Range.cells(1, 1).font.color = vbWhite
        End If
    Next key
    If Not tbl.DataBodyRange Is Nothing Then
        phase = "Input conditions"
        tbl.DataBodyRange.FormatConditions.Delete
        tbl.DataBodyRange.Interior.color = autoBody: tbl.DataBodyRange.font.color = vbBlack
        With tbl.DataBodyRange.FormatConditions.Add(xlExpression, , "=" & tbl.DataBodyRange.cells(1, CLng(map("STATE"))).address(False, True) & "=""" & Replace$(Finance_Text("HISTORICAL"), """", """""") & """")
            .Interior.color = RGB(231, 237, 241)
            .font.color = RGB(80, 87, 94)
        End With
        For Each key In Finance_InputKeys()
            If Not readOnly Then tbl.ListColumns(CLng(map(key))).DataBodyRange.Interior.Pattern = xlPatternNone
        Next key
        tbl.ListColumns(CLng(map("PROGRESS"))).DataBodyRange.numberFormat = "0.0%"
        For Each key In Array("START", "FINISH")
            tbl.ListColumns(CLng(map(key))).DataBodyRange.numberFormat = Settings_GetDateNumberFormat()
        Next key
        For Each key In Array("BUDGET", "ACTUAL", "ETC", "EAC", "VARIANCE")
            phase = "Input money " & CStr(key)
            tbl.ListColumns(CLng(map(key))).DataBodyRange.numberFormat = moneyFormat
        Next key
        phase = "Profile list name"
        ThisWorkbook.names.Add name:="peFinanceProfiles", RefersTo:="='FINANCE'!$AX$1:$AX$3"
        With tbl.ListColumns(CLng(map("PROFILE"))).DataBodyRange.validation
            phase = "Profile validation delete"
            .Delete
            phase = "Profile validation add"
            If Not readOnly Then .Add xlValidateList, xlValidAlertStop, xlBetween, "=peFinanceProfiles"
        End With
    End If
    phase = "Retired input row format"
    Set tail = tbl.HeaderRowRange.offset(tbl.ListRows.count + 1).Resize(1, tbl.ListColumns.count)
    If Application.WorksheetFunction.CountA(tail) = 0 Then
        If tail.cells(1, CLng(map("TASK_ID"))).Interior.color = autoBody And _
           tail.cells(1, CLng(map("WBS"))).Interior.color = autoBody And _
           tail.cells(1, CLng(map("ACTUAL"))).Interior.color = autoBody And _
           tail.cells(1, CLng(map("PROGRESS"))).numberFormat = "0.0%" Then
            retired = True
            For Each tailCell In tail.cells
                If Not tailCell.comment Is Nothing Then retired = False: Exit For
                commentText = WBS_Onboarding_ThreadedCommentText(tailCell, hasComment)
                If hasComment Then retired = False: Exit For
            Next tailCell
            If retired Then tail.Clear
        End If
    End If
    Set tbl = Finance_Table(FIN_EXPENSES): Set map = Finance_Map(tbl)
    tbl.TableStyle = template.TableStyle.name
    tbl.HeaderRowRange.Interior.color = autoHeader: tbl.HeaderRowRange.font.color = vbBlack
    If Not tbl.DataBodyRange Is Nothing Then
        tbl.DataBodyRange.Interior.color = autoBody: tbl.DataBodyRange.font.color = vbBlack
        tbl.ListColumns(CLng(map("DATE"))).DataBodyRange.numberFormat = Settings_GetDateNumberFormat()
        tbl.ListColumns(CLng(map("AMOUNT"))).DataBodyRange.numberFormat = moneyFormat
    End If
    For Each name In Array("tbl_FINANCE_MONTHS", "tbl_FINANCE_TOTALS")
        Set tbl = Nothing
        On Error Resume Next: Set tbl = Finance_Table(CStr(name)): On Error GoTo Failed
        If Not tbl Is Nothing Then
            phase = "Output style " & CStr(name)
            tbl.TableStyle = template.TableStyle.name
            tbl.Range.font.name = template.HeaderRowRange.cells(1, 1).font.name
            tbl.Range.font.size = template.HeaderRowRange.cells(1, 1).font.size
            tbl.HeaderRowRange.Interior.color = autoHeader: tbl.HeaderRowRange.font.color = vbBlack
            If Not tbl.DataBodyRange Is Nothing Then
                tbl.DataBodyRange.Interior.color = autoBody: tbl.DataBodyRange.font.color = vbBlack
                If CStr(name) = "tbl_FINANCE_MONTHS" Then
                    phase = "Monthly money"
                    For i = 2 To 5: tbl.ListColumns(i).DataBodyRange.numberFormat = moneyFormat: Next i
                Else
                    phase = "Totals money"
                    For Each key In Array(2, 3, 4, 5, 6, 12, 14): tbl.ListColumns(CLng(key)).DataBodyRange.numberFormat = moneyFormat: Next key
                    phase = "Totals hierarchy style"
                    Finance_OutlineTotals tbl, regroupTotals
                End If
            End If
        End If
    Next name
    Exit Sub
Failed:
    errorNumber = Err.number: description = phase & ": " & Err.description
    On Error GoTo 0
    Err.Raise errorNumber, "Finance_Protect", description
End Sub

Public Sub Handle_Finance_Change(ByVal ws As Worksheet, ByVal target As Range)
    Dim tbl As ListObject, hit As Range, editable As Range, map As Object, key As Variant, name As Variant
    Dim oldEvents As Boolean, details As String
    oldEvents = Application.EnableEvents
    On Error GoTo InvalidEdit
    If ws.name = "FINANCE" Then
    Set tbl = ws.ListObjects(FIN_TABLE)
    If Not Intersect(target, tbl.Range) Is Nothing Then
        Set hit = Intersect(target, tbl.Range)
        Set map = Finance_Map(tbl)
        If Not tbl.DataBodyRange Is Nothing Then
            For Each key In Finance_InputKeys()
                If editable Is Nothing Then
                    Set editable = tbl.ListColumns(CLng(map(CStr(key)))).DataBodyRange
                Else
                    Set editable = Application.Union(editable, tbl.ListColumns(CLng(map(CStr(key)))).DataBodyRange)
                End If
            Next key
        End If
        If Not editable Is Nothing Then
            If Finance_ScopeId() <> 0 Then
                If Not Intersect(target, editable) Is Nothing Then
                    If Intersect(target, editable).CountLarge = target.CountLarge Then Exit Sub
                End If
            End If
        End If
    End If
    End If
    For Each name In Array(FIN_EXPENSES, "tbl_FINANCE_MONTHS", "tbl_FINANCE_TOTALS")
        Set tbl = Nothing
        On Error Resume Next: Set tbl = ws.ListObjects(CStr(name)): On Error GoTo InvalidEdit
        If Not tbl Is Nothing Then
            If Not Intersect(target, tbl.Range) Is Nothing Then Set hit = Intersect(target, tbl.Range)
        End If
    Next name
    If hit Is Nothing Then Exit Sub
RejectEdit:
    Application.EnableEvents = False
    On Error GoTo UndoFailed
    Application.Undo
    If Finance_ScopeId() = 0 Then
        CalcBridge_ShowSingleConsoleMessage "WARNING", "FINANCE.GLOBAL_READONLY"
    Else
        CalcBridge_ShowSingleConsoleMessage "WARNING", "FINANCE.READ_ONLY"
    End If
SafeExit:
    Application.EnableEvents = oldEvents
    Exit Sub
InvalidEdit:
    details = Err.description
    If Not hit Is Nothing Then Resume RejectEdit
    Resume UndoFailed
UndoFailed:
    If Len(details) = 0 Then details = Err.description
    Application.EnableEvents = oldEvents
    CalcBridge_ShowSingleConsoleMessage "STOP", "COMMON.ERROR.PROCEDURE_INLINE", TextCatalog_Arguments("Procedure", "Finance input guard", "Details", details)
End Sub

Public Sub Finance_Present(Optional ByVal regroupTotals As Boolean = False)
    Dim ws As Worksheet, name As Variant, tbl As ListObject, keys As Variant, i As Long, shp As shape, options(1 To 3, 1 To 1) As Variant
    Dim oldEvents As Boolean, errorNumber As Long, description As String, chartTop As Double, chart As ChartObject, report As Worksheet, activeBefore As Object, phase As String, statusText As String
    Dim currentHeaders As Variant, wantedTitles() As String, renameHeaders As Boolean
    Dim perf As clsPerfScope
    Set perf = Profiler_BeginScope("Finance.Presentation", "Finance")
    oldEvents = Application.EnableEvents
    Set activeBefore = Application.ActiveSheet
    On Error GoTo Failed
    Application.EnableEvents = False
    phase = "Input headers"
    Set ws = ThisWorkbook.Worksheets("FINANCE"): ws.Unprotect
    Finance_EnsureProjectionColumns Finance_Table(FIN_TABLE)
    Finance_EnsureProjectionColumns Finance_Table(FIN_EXPENSES)
    Finance_OrderColumns Finance_Table(FIN_TABLE)
    For Each name In Array(FIN_TABLE, FIN_EXPENSES)
        Set tbl = Finance_Table(CStr(name)): keys = Finance_Keys(CStr(name))
        ' Resolve before renaming so presentation never changes column identity.
        Dim map As Object: Set map = Finance_Map(tbl)
        currentHeaders = tbl.HeaderRowRange.Value2: ReDim wantedTitles(0 To UBound(keys)): renameHeaders = False
        For i = 0 To UBound(keys)
            wantedTitles(i) = Finance_VisibleTitle(CStr(keys(i)), CStr(name))
            If CStr(currentHeaders(1, CLng(map(keys(i))))) <> wantedTitles(i) Then renameHeaders = True
        Next i
        If renameHeaders Then
            For i = 0 To UBound(keys): tbl.ListColumns(CLng(map(keys(i)))).name = "__FIN_" & CStr(i): Next i
            For i = 0 To UBound(keys): tbl.ListColumns(CLng(map(keys(i)))).name = wantedTitles(i): Next i
        End If
        Finance_LocalizeValues tbl, map
    Next name
    options(1, 1) = Finance_Text("PROFILE_START"): options(2, 1) = Finance_Text("PROFILE_END"): options(3, 1) = Finance_Text("PROFILE_LINEAR")
    ws.Range("AX1:AX3").Value2 = options: ws.columns("AX:AY").hidden = True
    statusText = Finance_Text("NEVER")
    If IsNumeric(ws.Range("S2").Value2) And Len(CStr(ws.Range("S2").Value2)) > 0 Then statusText = Finance_Text("LAST") & " " & Format$(CDate(ws.Range("S2").Value2), Replace$(Settings_GetDateTimeNumberFormat(), ":ss", ""))
    ws.Range("E1").ClearContents
    Finance_StatusCaption ws, statusText
    Set shp = Finance_StartButton(ws)
    shp.TextFrame2.TextRange.text = Finance_Text("START")
    shp.visible = (Finance_Table(FIN_TABLE).ListRows.count = 0)
    phase = "Controls": Finance_PresentationControls ws
    phase = "Summary sheet"
    Set report = Finance_SummarySheet()
    report.Range("A1").Value2 = statusText
    If FinanceStore_Exists(ThisWorkbook) Then
        Dim domain As Object: Set domain = Finance_PresentationDomain()
        If domain("PROJECTS").count > 1 Then report.Range("A1").Value2 = statusText & " | " & Finance_ScopeCaption()
    End If
    Dim outputName As Variant, headers As Variant
    For Each outputName In Array("tbl_FINANCE_MONTHS", "tbl_FINANCE_TOTALS")
        Set tbl = Nothing
        On Error Resume Next: Set tbl = Finance_Table(CStr(outputName)): On Error GoTo Failed
        If Not tbl Is Nothing Then
            phase = "Position " & CStr(outputName)
            If CStr(outputName) = "tbl_FINANCE_MONTHS" Then
                Finance_PositionOutput tbl, "A25", report
            Else
                Finance_PositionOutput tbl, Finance_TotalsAddress(), report
            End If
            headers = Finance_OutputHeaders(CStr(outputName))
            currentHeaders = tbl.HeaderRowRange.Value2
            For i = 0 To UBound(headers)
                If CStr(currentHeaders(1, i + 1)) <> CStr(headers(i)) Then tbl.ListColumns(i + 1).name = CStr(headers(i))
            Next i
            If CStr(outputName) = "tbl_FINANCE_TOTALS" And tbl.ListRows.count > 0 Then tbl.DataBodyRange.cells(1, 1).Value2 = Finance_ScopeCaption()
        End If
    Next outputName
    phase = "Ledger position"
    Finance_MigrateExtractCounter
    Set tbl = Finance_Table(FIN_EXPENSES)
    If tbl.HeaderRowRange.column <> 19 Then
        Dim destination As Range, overlap As Range, occupied As Double, oldGuide As Range, guideExists As Boolean, guideText As String, ledgerWidths() As Double
        Set destination = ws.Range("S3").Resize(tbl.Range.rows.count, tbl.ListColumns.count)
        occupied = WorksheetFunction.CountA(destination)
        Set overlap = Application.Intersect(destination, tbl.Range)
        If Not overlap Is Nothing Then occupied = occupied - WorksheetFunction.CountA(overlap)
        If occupied > 0 Then Err.Raise 5, "Finance", "Ledger destination is not empty"
        ReDim ledgerWidths(1 To tbl.ListColumns.count)
        For i = 1 To tbl.ListColumns.count: ledgerWidths(i) = tbl.ListColumns(i).Range.EntireColumn.ColumnWidth: Next i
        tbl.Range.EntireColumn.hidden = False
        Set map = Finance_Map(tbl)
        For Each outputName In Array("PROJECT_NAME", "AMOUNT")
            Set oldGuide = tbl.HeaderRowRange.cells(1, CLng(map(CStr(outputName)))).offset(-1)
            guideText = WBS_Onboarding_ThreadedCommentText(oldGuide, guideExists)
            If guideExists Or Not oldGuide.comment Is Nothing Then oldGuide.Cut destination:=oldGuide.offset(0, 19 - tbl.HeaderRowRange.column)
        Next outputName
        Finance_PositionOutput tbl, "S3", ws
        For i = 1 To tbl.ListColumns.count: tbl.ListColumns(i).Range.EntireColumn.ColumnWidth = ledgerWidths(i): Next i
    End If
    ws.columns("R").hidden = False
    tbl.Range.Resize(, 5).EntireColumn.hidden = False: tbl.HeaderRowRange.numberFormat = "General"
    phase = "Formats"
    Finance_Protect regroupTotals
    If FinanceStore_Exists(ThisWorkbook) Then
        Set tbl = Finance_Table(FIN_EXPENSES): Set map = Finance_Map(tbl)
        Dim hideProject As Boolean: hideProject = (Finance_ScopeId() <> 0)
        With tbl.ListColumns(CLng(map("PROJECT_NAME"))).Range.EntireColumn
            If CBool(.hidden) <> hideProject Then .hidden = hideProject
        End With
        Dim guide As Range, guideKey As Variant, guideTitle As String
        For Each guideKey In Array("AMOUNT", "PROJECT_NAME")
            Set guide = tbl.HeaderRowRange.cells(1, CLng(map(CStr(guideKey)))).offset(-1)
            guide.numberFormat = "General": guide.Value2 = TextCatalog_Get("WBS.ONBOARDING.CALCULATED", Settings_GetOwnerLanguage("FINANCE"))
            guide.font.size = 9: guide.font.color = RGB(90, 90, 90): guide.HorizontalAlignment = xlCenter
            guideTitle = CStr(tbl.ListColumns(CLng(map(CStr(guideKey)))).name)
            WBS_Onboarding_ApplyHelpCell guide, guideTitle & vbLf & vbLf & Finance_Text("HELP.JOURNAL." & CStr(guideKey)), Len(guideTitle)
        Next guideKey
    End If
    Set tbl = Finance_Table(FIN_TABLE): Set map = Finance_Map(tbl)
    Finance_ColumnHelp Finance_Table(FIN_TABLE)
    If FinanceStore_Exists(ThisWorkbook) Then Finance_ContextColumns Finance_Table(FIN_TABLE)
    For Each outputName In Array("tbl_FINANCE_MONTHS", "tbl_FINANCE_TOTALS")
        Set tbl = Nothing
        On Error Resume Next: Set tbl = Finance_Table(CStr(outputName)): On Error GoTo Failed
        If Not tbl Is Nothing Then Finance_ColumnHelp tbl
    Next outputName
    phase = "Move chart"
    If ws.ChartObjects.count > 0 Then ws.ChartObjects(1).chart.Location xlLocationAsObject, report.name
    Set tbl = Nothing
    On Error Resume Next: Set tbl = Finance_Table("tbl_FINANCE_MONTHS"): On Error GoTo Failed
    If Not tbl Is Nothing Then
        If tbl.ListRows.count > 0 Then phase = "Chart projection": Finance_DrawChart report, mProjectionOutputsChanged
    End If
    If report.ChartObjects.count > 0 Then
        phase = "Chart layout"
        Set chart = report.ChartObjects(1)
        chart.top = report.Range("A3").top: chart.left = report.Range("A3").left: chart.height = report.Range("A25").top - chart.top - 18
        chart.Placement = xlFreeFloating
    End If
    Finance_DrawTodayLine report
CleanExit:
    On Error Resume Next
    If Not activeBefore Is Nothing Then
        If Not (Application.ActiveSheet Is activeBefore) Then activeBefore.Activate
    End If
    Application.EnableEvents = oldEvents
    On Error GoTo 0
    If errorNumber <> 0 Then Err.Raise errorNumber, "Finance_Present", description
    Exit Sub
Failed:
    errorNumber = Err.number: description = phase & ": " & Err.description
    Resume CleanExit
End Sub

Private Function Finance_OutputHeaders(ByVal name As String) As Variant
    Dim keys As Variant, i As Long, globalScope As Boolean
    keys = Finance_OutputKeys(name)
    If FinanceStore_Exists(ThisWorkbook) Then globalScope = (Finance_ScopeId() = 0)
    For i = 0 To UBound(keys)
        Select Case CStr(keys(i))
            Case "COUNT", "KNOWN_BUDGET", "KNOWN_ACTUAL", "KNOWN_ETC", "KNOWN_EAC"
                If globalScope Then keys(i) = Finance_Text("GLOBAL." & CStr(keys(i))) Else keys(i) = Finance_Text(CStr(keys(i)))
            Case Else: keys(i) = Finance_Text(CStr(keys(i)))
        End Select
    Next i
    Finance_OutputHeaders = keys
End Function

Public Sub Finance_Start()
    If Not (Application.ActiveWorkbook Is ThisWorkbook) Then Exit Sub
    If Application.ActiveSheet.name <> "FINANCE" Or TypeName(Application.Caller) <> "String" Then Exit Sub
    If CStr(Application.Caller) <> "peFinanceStart" Then Exit Sub
    Finance_Manage
End Sub

Public Sub Finance_Manage()
    Dim n As Long, description As String, entered As Boolean
    If Not WorkbookSchema_UserActionAllowed() Then Exit Sub
    If mExpenseActive Then Err.Raise 5, "Finance", Finance_Text("DRAFT_PENDING")
    On Error GoTo Failed
    frmFinanceActivities.Prepare
    If frmFinanceActivities.Ready Then
        Finance_Draft True: entered = True
        frmFinanceActivities.Show vbModal
    End If
CleanExit:
    If entered Then Finance_Draft False
    Unload frmFinanceActivities
    If n <> 0 Then Err.Raise n, "Finance", description
    Exit Sub
Failed:
    n = Err.number: description = Err.description: Resume CleanExit
End Sub

Public Sub Finance_ExpensesCommand()
    Dim n As Long, description As String, entered As Boolean
    If Not WorkbookSchema_UserActionAllowed() Then Exit Sub
    ' Native-sheet trial. The previous MSForms path remains available below.
    If FIN_NATIVE_EXPENSE_TRIAL Then
        Finance_ExpenseSheetOpen
        Exit Sub
    End If
    On Error GoTo Failed
    Finance_RequireIndividual
    Finance_Flush
    frmFinanceExpenses.Prepare Finance_SelectedReference()
    Finance_Draft True: entered = True
    frmFinanceExpenses.Show vbModal
CleanExit:
    If entered Then Finance_Draft False
    Unload frmFinanceExpenses
    If n <> 0 Then Err.Raise n, "Finance", description
    Exit Sub
Failed:
    n = Err.number: description = Err.description: Resume CleanExit
End Sub

Public Function Finance_ExpenseSheetActive() As Boolean
    Finance_ExpenseSheetActive = mExpenseActive
End Function

Public Sub Finance_ExpenseSheetOpen(Optional ByVal showPanel As Boolean = True)
    Dim ws As Worksheet, rows() As Variant, headers(1 To 1, 1 To 5) As Variant, entry As PEFinanceRow
    Dim name As String, suffix As Long, found As Boolean, i As Long, oldEvents As Boolean, n As Long, description As String, projectName As String
    If mExpenseActive Then
        mExpenseSheet.Activate
        If showPanel Then frmFinanceExpensePanel.Show vbModeless
        Exit Sub
    End If
    Call Finance_RequireIndividual
    Call Finance_Flush
    mExpenseCommitted = False
    Set mExpenseRows = Finance_DraftRows(): Set mExpenseOriginal = Finance_DraftExpenses()
    Set mExpenseOriginalIndex = Finance_DomainIndex(mExpenseOriginal, "EXPENSE_ID")
    mExpenseScope = Finance_ScopeId(): mExpenseRevision = Finance_Revision()
    projectName = CStr(Finance_DomainIndex(Finance_DomainSnapshot()("PROJECTS"), "PROJECT_ID")(CStr(mExpenseScope))("NAME"))
    oldEvents = Application.EnableEvents
    On Error GoTo Failed
    Application.EnableEvents = False: mExpenseEditing = True
    name = "FINANCE INPUT"
    Do
        found = False
        For Each ws In ThisWorkbook.Worksheets
            If StrComp(ws.name, name, vbTextCompare) = 0 Then found = True: Exit For
        Next ws
        If Not found Then Exit Do
        suffix = suffix + 1: name = "FINANCE INPUT " & CStr(suffix)
    Loop
    Set mExpenseSheet = ThisWorkbook.Worksheets.Add(after:=ThisWorkbook.Worksheets("FINANCE"))
    mExpenseSheet.name = name
    headers(1, 1) = "WBS": headers(1, 2) = Finance_Text("DATE_LABEL"): headers(1, 3) = Finance_Text("AMOUNT_HINT")
    headers(1, 4) = Finance_Text("COMMENT_HINT"): headers(1, 5) = "__ExpenseId"
    mExpenseHeaders = headers
    mExpenseSheet.Range("A1").Value2 = projectName
    mExpenseSheet.Range("A1").font.bold = True
    mExpenseSheet.Range("A3:E3").Value2 = headers
    mExpenseSheet.columns(1).numberFormat = "@"
    Set mExpenseTable = mExpenseSheet.ListObjects.Add(xlSrcRange, mExpenseSheet.Range("A3:E4").Resize(mExpenseOriginal.count + 2), , xlYes)
    mExpenseTable.TableStyle = Finance_Table(FIN_EXPENSES).TableStyle.name
    If mExpenseOriginal.count > 0 Then
        ReDim rows(1 To mExpenseOriginal.count, 1 To 5)
        For Each entry In mExpenseOriginal
            i = i + 1: rows(i, 1) = CStr(entry("WBS")): rows(i, 2) = CDbl(entry("DATE"))
            rows(i, 3) = CCur(entry("AMOUNT")): rows(i, 4) = CStr(entry("COMMENT")): rows(i, 5) = CLng(entry("EXPENSE_ID"))
        Next entry
        mExpenseTable.DataBodyRange.cells(1, 1).Resize(mExpenseOriginal.count, 5).Value2 = rows
    End If
    mExpenseSheet.columns(1).ColumnWidth = 20: mExpenseSheet.columns(2).ColumnWidth = 18
    mExpenseSheet.columns(3).ColumnWidth = 20: mExpenseSheet.columns(4).ColumnWidth = 60
    mExpenseSheet.columns(5).hidden = True
    mExpenseTable.ListColumns(2).DataBodyRange.numberFormat = Settings_GetDateNumberFormat()
    mExpenseTable.ListColumns(3).DataBodyRange.numberFormat = Finance_MoneyFormat()
    Unload frmFinanceExpensePanel
    frmFinanceExpensePanel.Prepare projectName, mExpenseSheet.name
    Finance_Draft True: mExpenseActive = True
    mExpenseSheet.Activate: mExpenseSheet.Range("A4").Select
    mExpenseEditing = False: Application.EnableEvents = oldEvents
    Ribbon_WorksheetActivated
    If showPanel Then
        frmFinanceExpensePanel.Show vbModeless
        frmFinanceExpensePanel.left = Application.left + 20
        frmFinanceExpensePanel.top = Application.top + Application.height - frmFinanceExpensePanel.height - 30
    End If
    Exit Sub
Failed:
    n = Err.number: description = Err.description
    mExpenseEditing = False: Application.EnableEvents = oldEvents
    Finance_ExpenseSheetCancel
    Err.Raise n, "Finance", description
End Sub

Private Function Finance_ExpenseSheetColumnsValid() As Boolean
    Dim i As Long
    On Error GoTo invalid
    If mExpenseTable Is Nothing Or IsEmpty(mExpenseHeaders) Then Exit Function
    If mExpenseTable.ListColumns.count <> 5 Then Exit Function
    For i = 1 To 5
        If mExpenseTable.ListColumns(i).name <> CStr(mExpenseHeaders(1, i)) Then Exit Function
    Next i
    Finance_ExpenseSheetColumnsValid = True
invalid:
End Function

Public Function Finance_ExpenseSheetChanged(ByVal sheet As Object, ByVal target As Range) As Boolean
    Dim last As Range, row As Long, oldEvents As Boolean, n As Long, description As String
    If mExpenseSheet Is Nothing Then Exit Function
    If Not (sheet Is mExpenseSheet) Then Exit Function
    Finance_ExpenseSheetChanged = True
    If mExpenseEditing Or Not mExpenseActive Then Exit Function
    If Not Finance_ExpenseSheetColumnsValid() Then Exit Function
    If Intersect(target, sheet.Range("A4:D" & CStr(sheet.rows.count))) Is Nothing Then Exit Function
    oldEvents = Application.EnableEvents
    On Error GoTo Failed
    Application.EnableEvents = False: mExpenseEditing = True
    Set last = sheet.Range("A4:E" & CStr(sheet.rows.count)).Find(What:="*", LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False, SearchFormat:=False)
    row = 4
    If Not last Is Nothing Then row = last.row + 1
    If row >= sheet.rows.count Then Err.Raise 5, "Finance", Finance_Text("NATIVE_FULL")
    If row > mExpenseTable.Range.row + mExpenseTable.Range.rows.count - 1 Then mExpenseTable.Resize sheet.Range("A3:E" & CStr(row))
    mExpenseTable.ListColumns(1).DataBodyRange.numberFormat = "@"
    mExpenseTable.ListColumns(2).DataBodyRange.numberFormat = Settings_GetDateNumberFormat()
    mExpenseTable.ListColumns(3).DataBodyRange.numberFormat = Finance_MoneyFormat()
CleanExit:
    mExpenseEditing = False: Application.EnableEvents = oldEvents
    On Error GoTo 0
    If n <> 0 Then Err.Raise n, "Finance", description
    Exit Function
Failed:
    n = Err.number: description = Err.description: Resume CleanExit
End Function

Public Function Finance_ExpenseSheetRows() As Collection
    Dim result As New Collection, Index As Object, ids As Object, seen As Object, data As Variant, r As Long, c As Long, blank As Boolean
    Dim entry As PEFinanceRow, id As String, details As String, reference As String
    Set Index = Finance_ExpenseInputIndex(mExpenseRows): Set ids = mExpenseOriginalIndex
    Set seen = CreateObject("Scripting.Dictionary")
    On Error GoTo Failed
    If Not Finance_ExpenseSheetColumnsValid() Then Err.Raise 5, , Finance_Text("NATIVE_COLUMNS")
    If mExpenseTable.DataBodyRange Is Nothing Then Set Finance_ExpenseSheetRows = result: Exit Function
    If IsNull(mExpenseTable.DataBodyRange.hasFormula) Then Err.Raise 5, , Finance_Text("NATIVE_VALUES_ONLY")
    If mExpenseTable.DataBodyRange.hasFormula Then Err.Raise 5, , Finance_Text("NATIVE_VALUES_ONLY")
    data = mExpenseTable.DataBodyRange.Value2
    For r = 1 To UBound(data, 1)
        mExpenseErrorRow = mExpenseTable.HeaderRowRange.row + r: blank = True
        For c = 1 To 4
            If Len(Trim$(CStr(data(r, c)))) > 0 Then blank = False
        Next c
        If Not blank Then
            id = CStr(data(r, 5)): reference = CStr(data(r, 1))
            If Len(id) = 0 Then
                If Finance_ExpenseHeaderTitles(reference, CStr(data(r, 2)), CStr(data(r, 3)), CStr(data(r, 4))) Then GoTo NextNativeDraftRow
            End If
            If Len(id) > 0 Then
                If Not ids.exists(id) Or seen.exists(id) Then Err.Raise 5, , Finance_Text("SOURCE_CHANGED")
                If CStr(ids(id)("WBS")) = reference Then reference = Finance_Reference(ids(id))
                seen.Add id, True
            End If
            Set entry = Finance_ExpenseInputRow(Index, reference, data(r, 2), data(r, 3), CStr(data(r, 4)))
            If Len(id) > 0 Then entry.Add "EXPENSE_ID", CLng(id)
            result.Add entry
        End If
NextNativeDraftRow:
    Next r
    mExpenseErrorRow = 0: Set Finance_ExpenseSheetRows = result: Exit Function
Failed:
    details = Err.description
    Err.Raise 5, "Finance", TextCatalog_Format("FINANCE.BULK_ROW_ERROR", Settings_GetOwnerLanguage("FINANCE"), TextCatalog_Arguments("Row", CStr(mExpenseErrorRow), "Details", details))
End Function

Public Sub Finance_ExpenseSheetAction(ByVal name As String)
    Dim entry As PEFinanceRow, entries As Collection, pending As New Collection, before As Object, current As PEFinanceRow, activities As Object
    Dim id As String, n As Long, description As String
    If Not mExpenseActive Then Exit Sub
    If name = "cancel" Then Finance_ExpenseSheetCancel: Exit Sub
    If mExpenseCommitted Then
        If name = "confirm" Then Finance_ExpenseSheetCancel: Exit Sub
        Err.Raise 5, "Finance", Finance_Text("NATIVE_COMMITTED")
    End If
    On Error GoTo Failed
    Finance_CheckRevision mExpenseScope, mExpenseRevision
    If Not Finance_ExpenseSheetColumnsValid() Then Err.Raise 5, "Finance", Finance_Text("NATIVE_COLUMNS")
    Select Case name
        Case "confirm"
            Set activities = Finance_DomainIndex(Finance_Activities(), "WBS")
            For Each entry In mExpenseRows
                If entry("STATE") = "ACTIVE" Then
                    If Not activities.exists(CStr(entry("WBS"))) Then Err.Raise 5, "Finance", Finance_Text("SOURCE_CHANGED")
                    If CLng(activities(CStr(entry("WBS")))("TASK_ID")) <> CLng(entry("TASK_ID")) Then Err.Raise 5, "Finance", Finance_Text("SOURCE_CHANGED")
                End If
            Next entry
            Set entries = Finance_ExpenseSheetRows(): Set before = mExpenseOriginalIndex
            For Each entry In entries
                If Not entry.exists("EXPENSE_ID") Then
                    pending.Add entry
                Else
                    id = CStr(entry("EXPENSE_ID")): Set current = before(id)
                    If Finance_Reference(entry) <> Finance_Reference(current) Or CDbl(entry("DATE")) <> CDbl(current("DATE")) Or CCur(entry("AMOUNT")) <> CCur(current("AMOUNT")) Or CStr(entry("COMMENT")) <> CStr(current("COMMENT")) Then pending.Add entry
                End If
            Next entry
            If Not Finance_ConfirmExpenseDates(mExpenseRows, pending) Then Exit Sub
            If ThisWorkbook.ProtectStructure Then Err.Raise 5, "Finance", Finance_Text("NATIVE_STRUCTURE")
            Finance_Commit mExpenseRows, entries
            mExpenseCommitted = True
            Finance_ExpenseSheetCancel
    End Select
    Exit Sub
Failed:
    n = Err.number: description = Err.description
    On Error Resume Next
    If mExpenseErrorRow > 0 And Not mExpenseSheet Is Nothing Then
        mExpenseSheet.Activate: mExpenseSheet.Range("A" & CStr(mExpenseErrorRow) & ":D" & CStr(mExpenseErrorRow)).Select
    End If
    On Error GoTo 0
    Err.Raise n, "Finance", description
End Sub

Public Sub Finance_ExpenseSheetCancel(Optional ByVal unloadPanel As Boolean = True)
    Dim oldAlerts As Boolean, oldEvents As Boolean, wasActive As Boolean, wasVisible As Boolean, n As Long, description As String
    Dim sheet As Worksheet, present As Boolean
    oldAlerts = Application.DisplayAlerts: oldEvents = Application.EnableEvents: wasActive = mExpenseActive
    wasVisible = frmFinanceExpensePanel.visible
    On Error GoTo Failed
    mExpenseActive = False: mExpenseEditing = True
    frmFinanceExpensePanel.Hide
    Application.DisplayAlerts = False: Application.EnableEvents = False
    If Not mExpenseSheet Is Nothing Then
        For Each sheet In ThisWorkbook.Worksheets
            If sheet Is mExpenseSheet Then present = True: Exit For
        Next sheet
        If present Then mExpenseSheet.Delete
    End If
CleanExit:
    If n = 0 Then
        Set mExpenseSheet = Nothing: Set mExpenseTable = Nothing: Set mExpenseRows = Nothing: Set mExpenseOriginal = Nothing
        Set mExpenseOriginalIndex = Nothing
        mExpenseHeaders = Empty
        mExpenseCommitted = False
        If wasActive Then Finance_Draft False
    Else
        mExpenseActive = wasActive
    End If
    mExpenseEditing = False: Application.DisplayAlerts = oldAlerts: Application.EnableEvents = oldEvents
    On Error GoTo 0
    If n <> 0 Then
        If wasVisible Then frmFinanceExpensePanel.Show vbModeless
        If mExpenseCommitted Then description = Finance_Text("NATIVE_COMMITTED") & vbCrLf & description
        Err.Raise n, "Finance", description
    End If
    If unloadPanel Then Unload frmFinanceExpensePanel
    ThisWorkbook.Worksheets("FINANCE").Activate
    Ribbon_WorksheetActivated
    Exit Sub
Failed:
    n = Err.number: description = Err.description: Resume CleanExit
End Sub

Public Sub Finance_Clean()
    If Not WorkbookSchema_UserActionAllowed() Then Exit Sub
    If mDraftDepth > 0 Or mBusy Or mTransaction Or mStorageFailed Then Err.Raise 5, "Finance", Finance_Text("STORAGE_BLOCKED")
    If MsgBox(Finance_Text("CLEAN_CONFIRM"), vbYesNo + vbDefaultButton2 + vbExclamation, Finance_Text("CLEAN")) <> vbYes Then Exit Sub
    Finance_Reset True
End Sub

Private Function Finance_SelectedReference() As String
    Dim tbl As ListObject, cell As Range, map As Object, row As Long
    If Not (Application.ActiveWorkbook Is ThisWorkbook) Then Exit Function
    Set tbl = Finance_Table(FIN_TABLE): Set cell = Application.ActiveCell
    If Not (cell.parent Is tbl.parent) Then Exit Function
    If tbl.DataBodyRange Is Nothing Then Exit Function
    If Intersect(cell, tbl.DataBodyRange) Is Nothing Then Exit Function
    row = cell.row - tbl.HeaderRowRange.row: Set map = Finance_Map(tbl)
    Finance_SelectedReference = Finance_State(CStr(tbl.DataBodyRange.cells(row, CLng(map("STATE"))).Value2)) & "|" & NormalizeWBS(tbl.DataBodyRange.cells(row, CLng(map("WBS"))).Value2)
End Function

Private Sub Finance_PresentationControls(ByVal ws As Worksheet)
    Dim choice As DropDown, codes As Variant, values() As Variant, i As Long, selected As Long, current As String, macro As String
    macro = "'" & Replace(ThisWorkbook.name, "'", "''") & "'!"
    CreateFixedHeaderToggle ws, "peFinanceVocabularyLabel", "peFinanceVocabularyTrack", "peFinanceVocabularyKnob", _
        10, 6, 90, Finance_Text("BEGINNER") & " / " & Finance_Text("EXPERT"), 105, 8, 36, 16, 12, macro & "Finance_ToggleVocabulary"
    ws.Shapes("peFinanceVocabularyLabel").TextFrame2.TextRange.text = Finance_Text("BEGINNER") & " / " & Finance_Text("EXPERT")
    RefreshFixedHeaderToggleVisual ws, "peFinanceVocabularyTrack", "peFinanceVocabularyKnob", (Settings_FinanceVocabulary() = "EXPERT")
    On Error Resume Next: Set choice = ws.DropDowns("peFinanceCurrency"): On Error GoTo 0
    If choice Is Nothing Then Set choice = ws.DropDowns.Add(238, 5, 90, 23): choice.name = "peFinanceCurrency"
    current = CStr(ws.Range("S1").Value2): codes = Array("EUR", "USD", "GBP", "CHF", "CAD", "AUD")
    For i = 0 To UBound(codes)
        If CStr(codes(i)) = current Then selected = i + 1
    Next i
    If selected = 0 Then
        ReDim Preserve codes(0 To UBound(codes) + 1)
        codes(UBound(codes)) = current: selected = UBound(codes) + 1
    End If
    ReDim values(1 To UBound(codes) + 1, 1 To 1)
    For i = 0 To UBound(codes): values(i + 1, 1) = codes(i): Next i
    ws.Range("AY1:AY10").ClearContents
    ws.Range("AY1").Resize(UBound(values, 1), 1).Value2 = values
    choice.ListFillRange = "'FINANCE'!$AY$1:$AY$" & CStr(UBound(values, 1))
    choice.value = selected: choice.OnAction = macro & "Finance_SelectCurrency"
    choice.Placement = xlFreeFloating
    ws.Shapes("peFinanceVocabularyLabel").Placement = xlFreeFloating
    ws.Shapes("peFinanceVocabularyTrack").Placement = xlFreeFloating
    ws.Shapes("peFinanceVocabularyKnob").Placement = xlFreeFloating
    Finance_ProjectControl ws
End Sub

Public Sub Finance_ToggleVocabulary()
    Dim oldMode As String
    If Not (Application.ActiveWorkbook Is ThisWorkbook) Then Exit Sub
    If Application.ActiveSheet.name <> "FINANCE" Then Exit Sub
    If Not WorkbookSchema_UserActionAllowed() Then Exit Sub
    oldMode = Settings_FinanceVocabulary()
    On Error GoTo Failed
    Settings_SetFinanceVocabulary IIf(oldMode = "EXPERT", "BEGINNER", "EXPERT")
    Call Finance_Present
    Exit Sub
Failed:
    Dim description As String: description = Err.description
    On Error Resume Next
    Settings_SetFinanceVocabulary oldMode
    Call Finance_Present
    On Error GoTo 0
    CalcBridge_ShowSingleConsoleMessage "STOP", "COMMON.ERROR.PROCEDURE_INLINE", TextCatalog_Arguments("Procedure", "Finance", "Details", description)
End Sub

Public Sub Finance_SelectCurrency()
    Dim ws As Worksheet, value As String, oldValue As String, choice As DropDown, graph As Object
    If Not (Application.ActiveWorkbook Is ThisWorkbook) Then Exit Sub
    If Application.ActiveSheet.name <> "FINANCE" Then Exit Sub
    If Not WorkbookSchema_UserActionAllowed() Then Exit Sub
    Set ws = ThisWorkbook.Worksheets("FINANCE"): oldValue = CStr(ws.Range("S1").Value2)
    On Error GoTo Failed
    Finance_Flush
    Set graph = Finance_DomainSnapshot()
    Set choice = ws.DropDowns("peFinanceCurrency")
    If choice.value < 1 Or choice.value > choice.ListCount Then Err.Raise 5, "Finance", "Invalid currency selection"
    value = CStr(choice.list(choice.value))
    If value <> oldValue And graph("RECORDS").count > 0 Then
        If MsgBox(Finance_Text("CURRENCY_CHANGE"), vbYesNo + vbDefaultButton2 + vbExclamation) <> vbYes Then GoTo Repaint
    End If
    graph("META")(1)("CURRENCY") = value
    Finance_CommitDomain graph, 3
    Exit Sub
Repaint:
    Call Finance_Present
    Exit Sub
Failed:
    Dim description As String: description = Err.description
    On Error Resume Next
    ws.Unprotect: ws.Range("S1").Value2 = oldValue
    Call Finance_Present
    On Error GoTo 0
    CalcBridge_ShowSingleConsoleMessage "STOP", "COMMON.ERROR.PROCEDURE_INLINE", TextCatalog_Arguments("Procedure", "Finance", "Details", description)
End Sub

Public Sub Finance_Reset(Optional ByVal clearExtracts As Boolean = False)
    If mExpenseActive Then Err.Raise 5, "Finance", Finance_Text("DRAFT_PENDING")
    Dim emptyRows As New Collection, ws As Worksheet, report As Worksheet, i As Long, oldAlerts As Boolean, oldEvents As Boolean, errorNumber As Long, description As String
    Dim graph As Object, currencyCode As String
    oldAlerts = Application.DisplayAlerts: oldEvents = Application.EnableEvents
    On Error GoTo Failed
    Application.EnableEvents = False
    Set ws = ThisWorkbook.Worksheets("FINANCE"): ws.Unprotect
    If mTransaction Then Err.Raise 5, "Finance", Finance_Text("STORAGE_BLOCKED")
    currencyCode = CStr(ws.Range("S1").Value2)
    If FinanceStore_Exists(ThisWorkbook) Then currencyCode = CStr(Finance_DomainSnapshot()("META")(1)("CURRENCY"))
    mViewBound = False
    Set graph = Finance_DomainBlank(currencyCode)
    FinanceStore_Write ThisWorkbook, graph
    Finance_Write Finance_Table(FIN_TABLE), emptyRows: Finance_Write Finance_Table(FIN_EXPENSES), emptyRows
    ws.Range("S2").ClearContents
    Finance_ClearOutputs ws
    If clearExtracts Then
        Application.DisplayAlerts = False
        For i = ThisWorkbook.Worksheets.count To 1 Step -1
            If Finance_IsExtractSheet(ThisWorkbook.Worksheets(i)) Then ThisWorkbook.Worksheets(i).Delete
        Next i
        ws.Range(FIN_EXTRACT_COUNTER).Value2 = 0
        Application.DisplayAlerts = oldAlerts
    End If
    Finance_BindSavedView
    Call Finance_Present
CleanExit:
    Application.DisplayAlerts = oldAlerts: Application.EnableEvents = oldEvents
    On Error GoTo 0
    If errorNumber <> 0 Then Err.Raise errorNumber, "Finance_Reset", description
    Exit Sub
Failed:
    errorNumber = Err.number: description = Err.description
    Resume CleanExit
End Sub

Public Function Finance_CalculationCount() As Long
    Finance_CalculationCount = mCalculations
End Function

Public Function Finance_ReadCount() As Long
    Finance_ReadCount = mReads
End Function

Public Sub Finance_ClearOutputs(ByVal ws As Worksheet)
    Dim name As Variant, tbl As ListObject, area As Range, report As Worksheet
    Dim marker As shape
    If Not (ws.parent Is ThisWorkbook) Then Err.Raise 5, "Finance_ClearOutputs", "Foreign workbook"
    For Each name In Array("tbl_FINANCE_MONTHS", "tbl_FINANCE_TOTALS")
        Set tbl = Nothing
        On Error Resume Next: Set tbl = Finance_Table(CStr(name)): On Error GoTo 0
        If Not tbl Is Nothing Then
            Set area = tbl.Range
            If CStr(name) = "tbl_FINANCE_TOTALS" Then
                area.EntireRow.hidden = False
                area.EntireRow.OutlineLevel = 1
            End If
            If Not tbl.DataBodyRange Is Nothing Then
                tbl.DataBodyRange.ClearContents
                tbl.Resize tbl.HeaderRowRange.Resize(2, tbl.ListColumns.count)
                If tbl.ListRows.count > 0 Then tbl.ListRows(1).Delete
            End If
        End If
    Next name
    Do While ws.ChartObjects.count > 0: ws.ChartObjects(1).Delete: Loop
    Set report = Finance_SummarySheet()
    Do While report.ChartObjects.count > 0: report.ChartObjects(1).Delete: Loop
    On Error Resume Next
    Set marker = report.Shapes(FIN_TODAY_LINE)
    On Error GoTo 0
    If Not marker Is Nothing Then marker.Delete
End Sub

Public Sub Finance_Commit(ByVal rows As Collection, ByVal expenses As Collection)
    Dim graph As Object
    Finance_RequireIndividual
    Set graph = Finance_DomainSnapshot(): Finance_ReadVisibleInputs graph
    Finance_Validate rows, expenses
    Finance_ApplyJournal graph, mViewScope, expenses
    Finance_CommitDomain graph, 1
End Sub

Public Sub Finance_CommitActivities(ByVal graph As Object, ByVal scope As Long, ByVal rows As Collection, ByVal decisions As Object)
    Dim phase As String, n As Long, description As String
    On Error GoTo Failed
    phase = "Validate activity draft"
    Finance_CheckRevision Finance_ScopeId(), CLng(graph("META")(1)("REVISION"))
    Finance_CheckDraftInputs graph
    If scope = 0 Then Err.Raise 5, "Finance", Finance_Text("GLOBAL_READONLY")
    phase = "Apply memberships"
    Finance_ApplyMemberships graph, scope, rows, decisions
    graph("META")(1)("SCOPE_ID") = scope
    phase = "Commit activity domain"
    Finance_CommitDomain graph, IIf(scope = mViewScope, 1, 2)
    Exit Sub
Failed:
    n = Err.number: description = Err.description
    Err.Raise n, "Finance_CommitActivities", phase & ": " & description
End Sub

Private Sub Finance_CheckDraftInputs(ByVal graph As Object)
    Dim before As Object, records As Object, allowed As Object, row As Object, key As Variant, id As String
    Dim candidateRecords As Object, originalJournal As Object, candidateJournal As Object, f As String
    Set before = Finance_DomainSnapshot(): Set records = Finance_DomainIndex(before("RECORDS"), "FINANCIAL_ID")
    Set allowed = CreateObject("Scripting.Dictionary")
    If mViewScope <> 0 Then
        For Each row In before("LINKS")
            If CLng(row("PROJECT_ID")) = mViewScope Then allowed(CStr(row("FINANCIAL_ID"))) = True
        Next row
    End If
    For Each row In graph("RECORDS")
        id = CStr(row("FINANCIAL_ID"))
        If records.exists(id) And Not allowed.exists(id) Then
            For Each key In Finance_InputKeys()
                If CStr(row(CStr(key))) <> CStr(records(id)(CStr(key))) Then Err.Raise 5, "Finance", Finance_Text("GLOBAL_READONLY")
            Next key
        End If
    Next row
    Set candidateRecords = Finance_DomainIndex(graph("RECORDS"), "FINANCIAL_ID")
    Set originalJournal = Finance_DomainIndex(before("JOURNAL"), "EXPENSE_ID")
    Set candidateJournal = Finance_DomainIndex(graph("JOURNAL"), "EXPENSE_ID")
    For Each row In before("JOURNAL")
        id = CStr(row("EXPENSE_ID")): f = CStr(row("FINANCIAL_ID"))
        If candidateRecords.exists(f) And Not allowed.exists(f) Then
            If Not candidateJournal.exists(id) Then Err.Raise 5, "Finance", Finance_Text("GLOBAL_READONLY")
            For Each key In FinanceStore_Keys("JOURNAL")
                If CStr(row(CStr(key))) <> CStr(candidateJournal(id)(CStr(key))) Then Err.Raise 5, "Finance", Finance_Text("GLOBAL_READONLY")
            Next key
        End If
    Next row
    For Each row In graph("JOURNAL")
        f = CStr(row("FINANCIAL_ID")): id = CStr(row("EXPENSE_ID"))
        If records.exists(f) And Not allowed.exists(f) Then
            If Not originalJournal.exists(id) Then Err.Raise 5, "Finance", Finance_Text("GLOBAL_READONLY")
            If CStr(originalJournal(id)("FINANCIAL_ID")) <> f Then Err.Raise 5, "Finance", Finance_Text("GLOBAL_READONLY")
        End If
    Next row
End Sub

Private Function Finance_RecordedCosts(ByVal expenses As Collection) As Object
    Dim actuals As Object, e As Object, ref As String
    Set actuals = CreateObject("Scripting.Dictionary")
    For Each e In expenses
        ref = Finance_FinancialReference(e)
        If Not actuals.exists(ref) Then actuals.Add ref, CCur(0)
        actuals(ref) = CCur(actuals(ref)) + CCur(Finance_Money(e("AMOUNT"), True))
    Next e
    Set Finance_RecordedCosts = actuals
End Function

Public Function Finance_MissingTargets(ByVal graph As Object, ByVal activities As Collection) As Collection
    Dim available As Object, unique As Object, projects As Object, d As Object, link As Object, target As Object, missing As New Collection, wbs As String
    Set available = CreateObject("Scripting.Dictionary")
    For Each d In activities: available.Add CStr(d("WBS")), d: Next d
    Set unique = CreateObject("Scripting.Dictionary")
    For Each link In graph("LINKS")
        If link("STATE") = "ACTIVE" Then
            wbs = CStr(link("WBS"))
            Dim absent As Boolean: absent = Not available.exists(wbs)
            If Not absent Then absent = CBool(available(wbs)("PARENT"))
            If absent Then
                If projects Is Nothing Then Set projects = Finance_DomainIndex(graph("PROJECTS"), "PROJECT_ID")
                If Not unique.exists(wbs) Then
                    Set target = Finance_NewRow(FIN_TABLE)
                    target("WBS") = wbs: target("STATE") = "ACTIVE": target("TASK_NAME") = link("TASK_NAME")
                    unique.Add wbs, target: missing.Add target
                End If
                Set target = unique(wbs)
                target("TASK_NAME") = CStr(target("TASK_NAME")) & " | " & CStr(projects(CStr(link("PROJECT_ID")))("NAME"))
            End If
        End If
    Next link
    Set Finance_MissingTargets = missing
End Function

Public Sub Finance_ApplyMissingDecisions(ByVal graph As Object, ByVal decisions As Object)
    Dim historical As Object, matched As Object, link As Object, kept As New Collection, key As String, historyKey As String, choice As String, d As Variant
    Set historical = CreateObject("Scripting.Dictionary"): Set matched = CreateObject("Scripting.Dictionary")
    For Each link In graph("LINKS")
        If link("STATE") = "HISTORICAL" Then historical(CStr(link("PROJECT_ID")) & "|" & CStr(link("WBS"))) = True
    Next link
    For Each link In graph("LINKS")
        key = Finance_Reference(link)
        If link("STATE") = "ACTIVE" And decisions.exists(key) Then
            choice = CStr(decisions(key))
            If choice <> "KEEP" And choice <> "DELETE" Then Err.Raise 5, "Finance", Finance_Text("INVALID_DOMAIN")
            historyKey = CStr(link("PROJECT_ID")) & "|" & CStr(link("WBS"))
            If choice = "KEEP" And historical.exists(historyKey) Then Err.Raise 5, "Finance", Finance_Text("HISTORY_COLLISION")
            matched(key) = True
        End If
    Next link
    For Each d In decisions.keys
        If Not matched.exists(CStr(d)) Then Err.Raise 5, "Finance", Finance_Text("SOURCE_CHANGED")
    Next d
    For Each link In graph("LINKS")
        key = Finance_Reference(link)
        If link("STATE") = "ACTIVE" And decisions.exists(key) Then
            If decisions(key) = "KEEP" Then link("STATE") = "HISTORICAL": kept.Add link
        Else
            kept.Add link
        End If
    Next link
    Set graph("LINKS") = kept
    Finance_GarbageCollect graph
    Finance_DomainValidate graph
End Sub

Public Function Finance_ResolveMissing(ByVal graph As Object, ByVal activities As Collection) As Boolean
    Dim missing As Collection
    Set missing = Finance_MissingTargets(graph, activities)
    If missing.count = 0 Then Finance_ResolveMissing = True: Exit Function
    frmFinanceResolution.Prepare missing
    frmFinanceResolution.Show vbModal
    Finance_ResolveMissing = frmFinanceResolution.Accepted
    If Finance_ResolveMissing Then
        Finance_ApplyMissingDecisions graph, frmFinanceResolution.decisions
    End If
    Unload frmFinanceResolution
End Function

Private Function Finance_UsableDate(ByVal v As Variant) As Boolean
    If isError(v) Or IsEmpty(v) Then Exit Function
    If Not IsNumeric(v) Then Exit Function
    If Len(CStr(v)) = 0 Then Exit Function
    Finance_UsableDate = (CDbl(v) >= 1 And CDbl(v) <= 2958465# And CDbl(v) = Fix(CDbl(v)))
End Function

Private Function Finance_Month(ByVal v As Variant) As Long
    Finance_Month = CLng(DateSerial(Year(CDate(v)), Month(CDate(v)), 1))
End Function

Private Sub Finance_Output(ByVal ws As Worksheet, ByVal name As String, ByVal address As String, ByVal headers As Variant, ByVal data As Variant)
    Dim tbl As ListObject, i As Long, n As Long, oldCount As Long, physicalRows As Long, oldHeaderRow As Long
    If Not IsEmpty(data) Then n = UBound(data, 1)
    physicalRows = n + 1: If physicalRows < 2 Then physicalRows = 2
    On Error Resume Next
    Set tbl = Finance_Table(name)
    On Error GoTo 0
    If tbl Is Nothing Then
        Set tbl = ws.ListObjects.Add(xlSrcRange, ws.Range(address).Resize(physicalRows, UBound(headers) + 1), , xlYes)
        tbl.name = name
    Else
        Finance_PositionOutput tbl, address, ws
        oldCount = tbl.Range.rows.count
        tbl.Resize ws.Range(address).Resize(physicalRows, UBound(headers) + 1)
        If oldCount > physicalRows Then ws.Range(address).offset(physicalRows).Resize(oldCount - physicalRows, UBound(headers) + 1).Clear
    End If
    For i = 0 To UBound(headers): tbl.ListColumns(i + 1).name = CStr(headers(i)): Next i
    tbl.TableStyle = "TableStyleMedium2"
    tbl.HeaderRowRange.WrapText = True
    If tbl.HeaderRowRange.rowHeight < 60 Then tbl.HeaderRowRange.rowHeight = 60
    If n = 0 Then
        If tbl.ListRows.count > 0 Then tbl.ListRows(1).Delete
        Exit Sub
    End If
    Finance_AssignBody tbl, data
    If name = "tbl_FINANCE_MONTHS" Then
        tbl.ListColumns(1).DataBodyRange.numberFormat = "mmm yyyy"
        tbl.DataBodyRange.offset(0, 1).Resize(n, 4).numberFormat = "#,##0.00"
    Else
        tbl.DataBodyRange.numberFormat = "#,##0.00"
        tbl.DataBodyRange.offset(0, 6).Resize(n, 5).numberFormat = "0"
        tbl.ListColumns(13).DataBodyRange.numberFormat = "0"
    End If
End Sub

Private Sub Finance_PositionOutput(ByRef tbl As ListObject, ByVal address As String, Optional ByVal destination As Worksheet)
    Dim oldHeaderRow As Long, ws As Worksheet, original As Worksheet, tableName As String
    Dim help As Range, modern As Boolean, text As String, moveHelp As Boolean
    Set original = tbl.parent
    tableName = tbl.name
    If destination Is Nothing Then Set destination = original
    Set ws = destination: oldHeaderRow = tbl.HeaderRowRange.row
    If Not (original Is ws) Or tbl.Range.cells(1, 1).address <> ws.Range(address).address Then
        If tableName = "tbl_FINANCE_MONTHS" Or tableName = "tbl_FINANCE_TOTALS" Then
            Set help = tbl.HeaderRowRange.offset(-1)
            text = WBS_Onboarding_ThreadedCommentText(help.cells(1, 1), modern)
            moveHelp = modern Or Not help.cells(1, 1).comment Is Nothing
        End If
        If moveHelp Then
            help.Resize(tbl.Range.rows.count + 1, tbl.ListColumns.count).Cut destination:=ws.Range(address).offset(-1)
        Else
            tbl.Range.Cut destination:=ws.Range(address)
        End If
        Set tbl = ws.ListObjects(tableName)
        If original.name = "FINANCE" And oldHeaderRow <> 3 And original.rows(oldHeaderRow).rowHeight = 42 Then original.rows(oldHeaderRow).rowHeight = original.StandardHeight
    End If
    tbl.HeaderRowRange.WrapText = True
    If tbl.HeaderRowRange.rowHeight < 60 Then tbl.HeaderRowRange.rowHeight = 60
End Sub

Private Function Finance_Compute(ByVal rows As Collection, ByVal expenses As Collection, Optional ByVal activities As Collection, Optional ByVal groups As Object) As Object
    Dim ws As Worksheet, result As Object
    Dim available As Object, calcMap As Object, calcIndex As Object, tbl As ListObject, calc As Variant
    Dim totals As Object, actuals As Object, datedBudgets As Object, d As PEFinanceRow, e As PEFinanceRow, source As PEFinanceRow, ref As String
    Dim low As Long, high As Long, monthKey As Long, lastMonth As Long, count As Long, i As Long, r As Long
    Dim months() As Variant, monthData As Variant, spent() As Currency, budget() As Currency, a As Variant, key As Variant
    Dim firstDate As Long, lastDate As Long, startDate As Long, finishDate As Long, days As Long, eligible As Long
    Dim amount As Currency, allocated As Currency, piece As Currency, cb As Currency, ca As Currency, v As Variant
    Dim oldRows As Collection, oldExpenses As Collection, oldStamp As Variant, oldMonths As Variant, oldTotals As Variant
    Dim outTotals As Variant, scope As String, parent As String, chart As ChartObject, series As series
    Dim errorNumber As Long, description As String, scopeList As Collection, data As Variant, j As Long, report As Worksheet
    Dim captions As Object, ordered() As String, sortKeys() As String, savedScope As String, savedSort As String, sourceMissing As Boolean
    Dim perf As clsPerfScope, phaseScope As clsPerfScope, commitStarted As Boolean, oldEvents As Boolean, oldScreenUpdating As Boolean, phase As String, unallocatedMessage As String, overBudgetMessage As String
    phase = "Read inputs"
    On Error GoTo Failed
    phase = "Profiler"
    Set perf = Profiler_BeginScope("Finance.Refresh", "Finance")
    phase = "Read inputs"
    Set phaseScope = Profiler_BeginScope("Finance.ReadAndValidate", "Finance")
    Set ws = ThisWorkbook.Worksheets("FINANCE")
    If activities Is Nothing Then Set activities = Finance_Activities()
    phase = "Validate inputs"
    ' CommitDomain already validated the authority; these rows are its pure projection.
    If rows.count = 0 Then GoTo BuildResult
    Set available = CreateObject("Scripting.Dictionary")
    For Each d In activities: available.Add CStr(d("WBS")), d: Next d
    Set actuals = Finance_RecordedCosts(expenses)
    If groups Is Nothing Then Set groups = Finance_ContributionGroups(rows)
    Set datedBudgets = CreateObject("Scripting.Dictionary")
    For Each e In expenses
        monthKey = Finance_Month(e("DATE"))
        If low = 0 Or monthKey < low Then low = monthKey
        If monthKey > high Then high = monthKey
    Next e
    phase = "Read planning"
    Set phaseScope = Nothing
    Set phaseScope = Profiler_BeginScope("Finance.PlanningProjection", "Finance")
    Set tbl = ThisWorkbook.Worksheets("CALC").ListObjects("tbl_CALC")
    Set calcMap = CanonicalIdentity_BuildColumnMap(tbl): Set calcIndex = CreateObject("Scripting.Dictionary")
    If Not tbl.DataBodyRange Is Nothing Then
        calc = tbl.DataBodyRange.Value2
        Set calcIndex = CanonicalIdentity_BuildCalcRowById(calc, calcMap)
    End If
    phase = "Project activities"
    unallocatedMessage = Finance_Text("UNALLOCATED")
    overBudgetMessage = Finance_Text("OVER_BUDGET")
    For Each d In rows
        d("ACTUAL") = Empty: d("EAC") = Empty: d("VARIANCE") = Empty: d("DIAGNOSTIC") = ""
        sourceMissing = False
        If d("STATE") = "ACTIVE" Then
            sourceMissing = Not available.exists(CStr(d("WBS")))
            If Not sourceMissing Then sourceMissing = CBool(available(CStr(d("WBS")))("PARENT"))
            If sourceMissing Then
                ' Do not delete or silently historize an unresolved link on a scope change.
                d("PROGRESS") = Empty: d("DIAGNOSTIC") = Finance_Text("MISSING")
            Else
            Set source = available(CStr(d("WBS")))
            For Each key In Array("TASK_NAME", "TASK_ID", "PROJECT", "PROGRESS"): d(key) = source(key): Next key
            d("START") = Empty: d("FINISH") = Empty
            ref = CStr(d("TASK_ID"))
            If calcIndex.exists(ref) Then
                r = CLng(calcIndex(ref))
                If NormalizeWBS(calc(r, calcMap("WBS"))) = d("WBS") And Not Finance_Has(calc(r, calcMap("Error flag"))) Then
                    d("START") = calc(r, calcMap("Calculated Start")): d("FINISH") = calc(r, calcMap("Calculated Finish"))
                End If
            End If
            End If
        Else
            d("PROGRESS") = Empty
        End If
        ref = Finance_Reference(d)
        If actuals.exists(Finance_FinancialReference(d)) Then d("ACTUAL") = actuals(Finance_FinancialReference(d))
        If Finance_Has(d("ACTUAL")) And Finance_Has(d("ETC")) Then d("EAC") = CCur(d("ACTUAL")) + CCur(Finance_Money(d("ETC")))
        If Finance_Has(d("EAC")) And Finance_Has(d("BUDGET")) Then d("VARIANCE") = CCur(d("EAC")) - CCur(Finance_Money(d("BUDGET")))
        If Finance_Has(d("BUDGET")) Then
            If Not sourceMissing And Finance_UsableDate(d("START")) And Finance_UsableDate(d("FINISH")) Then
                If CLng(d("FINISH")) < CLng(d("START")) Then d("DIAGNOSTIC") = unallocatedMessage
            Else
                If Not sourceMissing Then d("DIAGNOSTIC") = unallocatedMessage
            End If
            If d("DIAGNOSTIC") = "" Then
                datedBudgets.Add ref, True
                monthKey = Finance_Month(d("START")): lastMonth = Finance_Month(d("FINISH"))
                If low = 0 Or monthKey < low Then low = monthKey
                If lastMonth > high Then high = lastMonth
            End If
            If Finance_Has(d("ACTUAL")) Then
                If CCur(d("ACTUAL")) > CCur(d("BUDGET")) Then
                    If Len(CStr(d("DIAGNOSTIC"))) > 0 Then d("DIAGNOSTIC") = CStr(d("DIAGNOSTIC")) & vbLf
                    d("DIAGNOSTIC") = CStr(d("DIAGNOSTIC")) & overBudgetMessage
                End If
            End If
        End If
    Next d
    phase = "Allocate months"
    Set phaseScope = Nothing
    Set phaseScope = Profiler_BeginScope("Finance.MonthlyAllocation", "Finance")
    If low <> 0 Then count = DateDiff("m", CDate(low), CDate(high)) + 1
    If count > 1200 Then Err.Raise 5, "Finance", Finance_Text("HORIZON")
    If count > 0 Then
        ReDim spent(1 To count): ReDim budget(1 To count): ReDim months(1 To count, 1 To 5)
    End If
    For Each e In expenses
        i = DateDiff("m", CDate(low), CDate(e("DATE"))) + 1
        spent(i) = spent(i) + CCur(e("AMOUNT")) * CLng(groups(CStr(e("FINANCIAL_ID"))).count)
    Next e
    For Each d In rows
        If datedBudgets.exists(Finance_Reference(d)) Then
            startDate = CLng(d("START")): finishDate = CLng(d("FINISH")): amount = CCur(d("BUDGET"))
            If d("PROFILE") <> "LINEAR" Then
                monthKey = IIf(d("PROFILE") = "START", startDate, finishDate)
                i = DateDiff("m", CDate(low), CDate(monthKey)) + 1: budget(i) = budget(i) + amount
            Else
                allocated = 0: days = finishDate - startDate + 1
                monthKey = Finance_Month(startDate)
                Do While monthKey <= finishDate
                    i = DateDiff("m", CDate(low), CDate(monthKey)) + 1
                    firstDate = monthKey: If firstDate < startDate Then firstDate = startDate
                    lastDate = CLng(DateAdd("m", 1, CDate(monthKey))) - 1: If lastDate > finishDate Then lastDate = finishDate
                    eligible = lastDate - firstDate + 1
                    If lastDate = finishDate Then
                        piece = amount - allocated
                    Else
                        piece = CCur(Fix(CDec(amount) * 100 * eligible / days)) / 100
                    End If
                    budget(i) = budget(i) + piece: allocated = allocated + piece
                    monthKey = CLng(DateAdd("m", 1, CDate(monthKey)))
                Loop
            End If
        End If
    Next d
    For i = 1 To count
        cb = cb + budget(i): ca = ca + spent(i)
        months(i, 1) = CDbl(DateAdd("m", i - 1, CDate(low)))
        months(i, 2) = budget(i): months(i, 3) = cb: months(i, 4) = spent(i): months(i, 5) = ca
    Next i
    If count > 0 Then monthData = months
    phase = "Consolidate"
    Set phaseScope = Nothing
    Set phaseScope = Profiler_BeginScope("Finance.Consolidation", "Finance")
    Set totals = CreateObject("Scripting.Dictionary"): Set scopeList = New Collection: Set captions = CreateObject("Scripting.Dictionary")
    For Each d In rows
        Finance_AddTotal totals, scopeList, "GLOBAL", d, datedBudgets.exists(Finance_Reference(d))
        If d("STATE") = "ACTIVE" Then
            parent = CStr(d("WBS"))
            Do While parent <> ""
                If available.exists(parent) Then
                    Finance_AddTotal totals, scopeList, "WBS: " & parent, d, datedBudgets.exists(Finance_Reference(d))
                    captions("WBS: " & parent) = parent & " " & CStr(available(parent)("TASK_NAME"))
                End If
                parent = GetParentWBS(parent)
            Loop
        Else
            scope = "HISTORY: " & CStr(d("WBS"))
            Finance_AddTotal totals, scopeList, scope, d, datedBudgets.exists(Finance_Reference(d))
            captions(scope) = "[" & Finance_Text("HISTORICAL") & "] " & CStr(d("WBS")) & " " & CStr(d("TASK_NAME"))
        End If
    Next d
    ReDim ordered(1 To scopeList.count): ReDim sortKeys(1 To scopeList.count)
    For i = 1 To scopeList.count
        ordered(i) = CStr(scopeList(i))
        If left$(ordered(i), 5) = "WBS: " Then sortKeys(i) = Finance_WBSSortKey(Mid$(ordered(i), 6))
        If left$(ordered(i), 9) = "HISTORY: " Then sortKeys(i) = "~" & Finance_WBSSortKey(Mid$(ordered(i), 10))
    Next i
    For i = 2 To UBound(ordered)
        savedScope = ordered(i): savedSort = sortKeys(i): j = i - 1
        Do While j > 0
            If StrComp(sortKeys(j), savedSort, vbBinaryCompare) <= 0 Then Exit Do
            ordered(j + 1) = ordered(j): sortKeys(j + 1) = sortKeys(j): j = j - 1
        Loop
        ordered(j + 1) = savedScope: sortKeys(j + 1) = savedSort
    Next i
    ReDim outTotals(1 To scopeList.count, 1 To 14)
    For i = 1 To scopeList.count
        scope = ordered(i): a = totals(scope): outTotals(i, 1) = scope
        If captions.exists(scope) Then outTotals(i, 1) = captions(scope)
        For j = 0 To 12: outTotals(i, j + 2) = a(j): Next j
        For j = 0 To 3
            If CLng(a(6 + j)) = 0 Then outTotals(i, j + 2) = Empty
        Next j
        If CLng(a(13)) = 0 Then outTotals(i, 6) = Empty
        If CLng(a(6)) = 0 Then outTotals(i, 12) = Empty: outTotals(i, 14) = Empty
    Next i
BuildResult:
    Set result = CreateObject("Scripting.Dictionary")
    result.Add "ROWS", rows: result.Add "EXPENSES", expenses
    result.Add "MONTHS", monthData: result.Add "TOTALS", outTotals: result.Add "COUNT", count
    Set phaseScope = Nothing: Set perf = Nothing
    Set Finance_Compute = result
    Exit Function
Failed:
    errorNumber = Err.number: description = phase & ": " & Err.description
    Set phaseScope = Nothing: Set perf = Nothing
    Err.Raise errorNumber, "Finance_Compute", description
End Function

Private Sub Finance_AddTotal(ByVal totals As Object, ByVal order As Collection, ByVal scope As String, ByVal d As Object, ByVal allocated As Boolean)
    Dim a As Variant, fields As Variant, i As Long
    If Not totals.exists(scope) Then
        totals.Add scope, Array(CCur(0), CCur(0), CCur(0), CCur(0), CCur(0), 0&, 0&, 0&, 0&, 0&, CCur(0), 0&, CCur(0), 0&): order.Add scope
    End If
    a = totals(scope): fields = Array("BUDGET", "ACTUAL", "ETC", "EAC", "VARIANCE")
    a(5) = a(5) + 1
    For i = 0 To 4
        If Finance_Has(d(fields(i))) Then
            a(i) = CCur(a(i)) + CCur(d(fields(i)))
            If i < 4 Then a(6 + i) = a(6 + i) + 1 Else a(13) = a(13) + 1
        End If
    Next i
    If Finance_Has(d("BUDGET")) Then
        If Not allocated Then
            a(10) = CCur(a(10)) + CCur(d("BUDGET")): a(11) = a(11) + 1
        Else
            a(12) = CCur(a(12)) + CCur(d("BUDGET"))
        End If
    End If
    totals(scope) = a
End Sub

Private Function Finance_OutputSnapshot(ByVal ws As Worksheet, ByVal name As String) As Variant
    Dim tbl As ListObject
    On Error Resume Next: Set tbl = Finance_Table(name): On Error GoTo 0
    If Not tbl Is Nothing Then
        If tbl.ListRows.count > 0 Then Finance_OutputSnapshot = tbl.DataBodyRange.Value2
    End If
End Function

Private Function Finance_OutputEqual(ByVal expected As Variant, ByVal actual As Variant) As Boolean
    Dim r As Long, c As Long
    If IsEmpty(expected) Or IsEmpty(actual) Then
        Finance_OutputEqual = IsEmpty(expected) And IsEmpty(actual): Exit Function
    End If
    If UBound(expected, 1) <> UBound(actual, 1) Or UBound(expected, 2) <> UBound(actual, 2) Then Exit Function
    For r = 1 To UBound(expected, 1)
        For c = 1 To UBound(expected, 2)
            If CStr(expected(r, c)) <> CStr(actual(r, c)) Then Exit Function
        Next c
    Next r
    Finance_OutputEqual = True
End Function

Private Sub Finance_RestoreOutput(ByVal ws As Worksheet, ByVal name As String, ByVal address As String, ByVal data As Variant)
    Dim tbl As ListObject
    If IsEmpty(data) Then Exit Sub
    ws.Range(address).Resize(UBound(data, 1), UBound(data, 2)).Value2 = data
    Set tbl = ws.ListObjects.Add(xlSrcRange, ws.Range(address).Resize(UBound(data, 1), UBound(data, 2)), , xlYes): tbl.name = name
    If UBound(data, 1) = 1 And tbl.ListRows.count > 0 Then tbl.ListRows(1).Delete
End Sub

Private Sub Finance_DrawChart(ByVal ws As Worksheet, Optional ByVal rebuildActual As Boolean = False)
    Dim chart As ChartObject, tbl As ListObject, series As series, i As Long, columns As Variant, moneyFormat As String
    Dim data As Variant, actualMonths() As Variant, actualCumulative() As Variant, r As Long, asOfMonth As Long
    Dim asOfDate As Double, futureInCurrentMonth As Currency, expense As Object
    Set tbl = Finance_Table("tbl_FINANCE_MONTHS")
    data = tbl.DataBodyRange.Value2
    If ws.ChartObjects.count = 0 Then
        rebuildActual = True
    ElseIf ws.ChartObjects(1).chart.SeriesCollection.count <> 4 Then
        rebuildActual = True
    End If
    If rebuildActual Then
    ReDim actualMonths(1 To UBound(data, 1)): ReDim actualCumulative(1 To UBound(data, 1))
    asOfDate = Finance_ChartDate()
    If asOfDate > 0 Then
        asOfMonth = Finance_Month(asOfDate)
        ' The journal and monthly business totals retain confirmed future entries.
        For Each expense In Finance_ChartExpenses()
            If CDbl(expense("DATE")) > asOfDate Then
                If Finance_Month(CDbl(expense("DATE"))) = asOfMonth Then futureInCurrentMonth = futureInCurrentMonth + CCur(expense("AMOUNT"))
            End If
        Next expense
    End If
    For r = 1 To UBound(data, 1)
        If asOfMonth > 0 And CDbl(data(r, 1)) <= asOfMonth Then
            actualMonths(r) = data(r, 4): actualCumulative(r) = data(r, 5)
            If CDbl(data(r, 1)) = asOfMonth Then
                actualMonths(r) = CDbl(CCur(actualMonths(r)) - futureInCurrentMonth)
                actualCumulative(r) = CDbl(CCur(actualCumulative(r)) - futureInCurrentMonth)
            End If
        Else
            actualMonths(r) = CVErr(xlErrNA): actualCumulative(r) = CVErr(xlErrNA)
        End If
    Next r
    End If
    If ws.ChartObjects.count = 0 Then
        Set chart = ws.ChartObjects.Add(ws.Range("R12").left, ws.Range("R12").top, 880, 310): chart.name = "cht_FINANCE"
    Else
        Set chart = ws.ChartObjects(1)
    End If
    moneyFormat = CStr(tbl.ListColumns(2).DataBodyRange.cells(1, 1).NumberFormatLocal)
    With chart.chart
        Do While .SeriesCollection.count > 4: .SeriesCollection(.SeriesCollection.count).Delete: Loop
        Do While .SeriesCollection.count < 4: .SeriesCollection.NewSeries: Loop
        .HasTitle = True: .ChartTitle.text = Finance_Text("CHART") & " (" & CStr(ThisWorkbook.Worksheets("FINANCE").Range("S1").Value2) & ")": .HasLegend = True
        .PlotVisibleOnly = False
        columns = Array(2, 4, 3, 5)
        For i = 0 To 3
            Set series = .SeriesCollection(i + 1)
            series.name = "=" & tbl.HeaderRowRange.cells(1, CLng(columns(i))).address(External:=True)
            series.XValues = tbl.ListColumns(1).DataBodyRange
            Select Case i
                Case 1: If rebuildActual Then series.values = actualMonths
                Case 3: If rebuildActual Then series.values = actualCumulative
                Case Else: series.values = tbl.ListColumns(CLng(columns(i))).DataBodyRange
            End Select
            If i < 2 Then
                series.ChartType = xlColumnClustered: series.AxisGroup = xlPrimary
            Else
                series.ChartType = xlLine: series.AxisGroup = xlSecondary
                series.MarkerStyle = xlMarkerStyleCircle: series.MarkerSize = 5
            End If
        Next i
        .axes(xlCategory).TickLabels.numberFormat = "mmm yyyy"
        .axes(xlCategory).CategoryType = xlCategoryScale
        .axes(xlCategory).axisBetweenCategories = True
        .axes(xlCategory).TickLabelPosition = xlTickLabelPositionLow
        .axes(xlValue, xlPrimary).MinimumScaleIsAuto = True: .axes(xlValue, xlPrimary).MaximumScaleIsAuto = True
        .axes(xlValue, xlSecondary).MinimumScaleIsAuto = True: .axes(xlValue, xlSecondary).MaximumScaleIsAuto = True
        .axes(xlValue, xlPrimary).TickLabels.NumberFormatLinked = False
        .axes(xlValue, xlSecondary).TickLabels.NumberFormatLinked = False
        .axes(xlValue, xlPrimary).TickLabels.NumberFormatLocal = moneyFormat
        .axes(xlValue, xlSecondary).TickLabels.NumberFormatLocal = moneyFormat
        .Refresh
    End With
End Sub

Private Function Finance_ChartDate() As Double
    Dim stamp As Variant
    stamp = ThisWorkbook.Worksheets("FINANCE").Range("S2").Value2
    If IsNumeric(stamp) And Len(CStr(stamp)) > 0 Then Finance_ChartDate = Fix(CDbl(stamp))
End Function

Private Sub Finance_DrawTodayLine(ByVal ws As Worksheet)
    Dim tbl As ListObject, marker As shape, endExclusive As Double, asOfDate As Double
    asOfDate = Finance_ChartDate()
    On Error Resume Next
    Set marker = ws.Shapes(FIN_TODAY_LINE)
    Set tbl = Finance_Table("tbl_FINANCE_MONTHS")
    On Error GoTo 0
    If ws.ChartObjects.count = 0 Or tbl Is Nothing Or asOfDate = 0 Then GoTo HideLine
    If tbl.DataBodyRange Is Nothing Or tbl.ListRows.count = 0 Then GoTo HideLine
    endExclusive = CDbl(DateAdd("m", 1, CDate(tbl.DataBodyRange.cells(tbl.ListRows.count, 1).Value2)))
    ws.ChartObjects(1).chart.Refresh
    DrawChartTodayVerticalLine ws.ChartObjects(1).chart, ws, tbl.ListColumns(1).DataBodyRange, FIN_TODAY_LINE, endExclusive, asOfDate
    Exit Sub
HideLine:
    If Not marker Is Nothing Then marker.visible = msoFalse
End Sub








