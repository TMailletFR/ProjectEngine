Attribute VB_Name = "mod_Dashboard"
Option Explicit

'===============================================================================
' MODULE : mod_Dashboard
' DOMAINE / DOMAIN : Dashboard Rendering
'
' FR
' Construit et actualise les KPI, textes, graphiques et controles du Dashboard.
' Ne possede pas les calculs Core ou S-Curve sources.
'
' EN
' Builds and refreshes Dashboard KPIs, text, charts and controls.
' Does not own source Core or S-Curve calculations.
'
' CONTRATS / CONTRACTS : Refresh_Dashboard, Refresh_Dashboard_FullBuild, Refresh_Dashboard_ContentOnly, Dashboard_SetLanguage, Dashboard_ApplyLanguage, Toggle_Dashboard_Language, Reset_Dashboard, Run_Dashboard_Update
' CALLBACKS EXTERNES / EXTERNAL CALLBACKS : Aucun / None
'===============================================================================


Private Const DASHBOARD_SHEET As String = "DASHBOARD"
Private Const DASHBOARD_SNAPSHOT_SHEET As String = "DASHBOARD_SNAPSHOTS"
Private Const DASHBOARD_SNAPSHOT_TABLE As String = "tbl_DASHBOARD_SNAPSHOTS"
Private Const DASH_CHART_SCURVE As String = "cht_Dashboard_SCurve"
Private Const DASH_SCURVE_TODAY_LINE As String = "DASHBOARD_SCURVE_TODAY_LINE"
Private Const DASH_PREFIX As String = "DASH_"
Private Const DASH_FROM_CELL As String = "C4"
Private Const DASH_TO_CELL As String = "G4"
Private gDashboardLanguage As String
Private gHotSpotHelper As String
Private gHotSpotStep As String

Public Function Dashboard_TutorialSnapshotTarget() As Range
    Dim ws As Worksheet
    Dim snapshots As ListObject
    On Error GoTo missing
    Set snapshots = ThisWorkbook.Worksheets(DASHBOARD_SNAPSHOT_SHEET).ListObjects(DASHBOARD_SNAPSHOT_TABLE)
    If snapshots.DataBodyRange Is Nothing Then Exit Function
    Set ws = ThisWorkbook.Worksheets(DASHBOARD_SHEET)
    If Len(CStr(ws.Range(DASH_FROM_CELL).Value2)) = 0 Then Exit Function
    Set Dashboard_TutorialSnapshotTarget = ws.Range(ws.cells(ws.Range(DASH_FROM_CELL).row, 2), ws.Shapes("btn_Dashboard_Refresh_Comparison").BottomRightCell)
missing:
End Function

Public Function Dashboard_TutorialNewSnapshotTarget() As Range
    Dim ws As Worksheet
    Dim Button As shape
    On Error GoTo missing
    Set ws = ThisWorkbook.Worksheets(DASHBOARD_SHEET)
    Set Button = ws.Shapes("btn_Update_Dashboard")
    Set Dashboard_TutorialNewSnapshotTarget = ws.Range(Button.TopLeftCell, Button.BottomRightCell)
missing:
End Function

'------------------------------------------------------------------------------
' FR: Rafraichit Dashboard a partir de l'etat courant.
' EN: Refreshes Dashboard from the current state.
'------------------------------------------------------------------------------
Public Sub Refresh_Dashboard(Optional ByVal preserveSnapshotSelection As Boolean = False)

    Dashboard_SetLanguage Settings_GetOwnerLanguage("DASHBOARD")
    If Dashboard_HasUsableLayout() Then
        Refresh_Dashboard_ContentOnly preserveSnapshotSelection
    Else
        Refresh_Dashboard_FullBuild preserveSnapshotSelection
    End If

End Sub
'------------------------------------------------------------------------------
' FR: Rafraichit Dashboard Full Build a partir de l'etat courant.
' EN: Refreshes Dashboard Full Build from the current state.
'------------------------------------------------------------------------------
Public Sub Refresh_Dashboard_FullBuild(Optional ByVal preserveSnapshotSelection As Boolean = False, Optional ByVal propagateErrors As Boolean = False)

    Dim ws As Worksheet
    Dim readContext As clsDashboardReadContext
    Dim fromLabel As String
    Dim toLabel As String
    Dim langCode As String
    Dim failureNumber As Long, failureDescription As String

    On Error GoTo ErrHandler
    Dashboard_SetLanguage Settings_GetOwnerLanguage("DASHBOARD")
    Set ws = Dashboard_EnsureSheet()
    If preserveSnapshotSelection Then
        fromLabel = CStr(ws.Range(DASH_FROM_CELL).value)
        toLabel = CStr(ws.Range(DASH_TO_CELL).value)
    End If
    langCode = Dashboard_CurrentLanguage()
    Dashboard_Clear ws
    Dashboard_SetupCanvas ws

    Set readContext = DashboardReadContext_Load()

    Dashboard_RenderHeader ws, fromLabel, toLabel
    Dashboard_RenderExecutiveSummary ws, readContext.wbsTable, readContext.WbsColumns, readContext.CalcTable, readContext.calcColumns, readContext.scurveProjection
    Dashboard_RenderSCurveChart ws, readContext.scurveProjection
    Dashboard_RenderPlanningOverview ws, readContext.wbsTable, readContext.WbsColumns, readContext.CalcTable, readContext.calcColumns
    Dashboard_RenderHotSpots ws, readContext.wbsTable, readContext.WbsColumns, readContext.CalcTable, readContext.calcColumns

SafeExit:
    Exit Sub
ErrHandler:
    If propagateErrors Then
        failureNumber = Err.number: failureDescription = Err.description
        On Error GoTo 0
        Err.Raise failureNumber, "Refresh_Dashboard_FullBuild", failureDescription
    End If
    Debug.Print "Refresh_Dashboard_FullBuild error " & Err.number & ": " & Err.description
    Resume SafeExit
End Sub

'------------------------------------------------------------------------------
' FR: Rafraichit Dashboard Content Only a partir de l'etat courant.
' EN: Refreshes Dashboard Content Only from the current state.
'------------------------------------------------------------------------------
Public Sub Refresh_Dashboard_ContentOnly(Optional ByVal preserveSnapshotSelection As Boolean = False, Optional ByVal propagateErrors As Boolean = False)

    Dim ws As Worksheet
    Dim readContext As clsDashboardReadContext
    Dim fromLabel As String
    Dim toLabel As String
    Dim oldScreenUpdating As Boolean
    Dim failureNumber As Long, failureDescription As String
    On Error GoTo ErrHandler
    Dashboard_SetLanguage Settings_GetOwnerLanguage("DASHBOARD")

    If Not Dashboard_HasUsableLayout() Then
        Refresh_Dashboard_FullBuild preserveSnapshotSelection, propagateErrors
        Exit Sub
    End If

    Set ws = ThisWorkbook.Worksheets(DASHBOARD_SHEET)
    If preserveSnapshotSelection Then
        fromLabel = CStr(ws.Range(DASH_FROM_CELL).value)
        toLabel = CStr(ws.Range(DASH_TO_CELL).value)
    End If
    oldScreenUpdating = Application.ScreenUpdating
    Application.ScreenUpdating = False

    Set readContext = DashboardReadContext_Load()
    Dashboard_UpdateHeaderTexts ws
    Dashboard_UpdateHeaderTimestamp ws
    Dashboard_SetupSnapshotControls ws, fromLabel, toLabel, False
    Dashboard_EnsureResetButton ws
    Dashboard_RemoveLegacyLanguageToggle ws
    Dashboard_UpdateKnownShapeTexts ws
    Dashboard_UpdateKpiCardsInPlace ws, readContext.wbsTable, readContext.WbsColumns, readContext.CalcTable, readContext.calcColumns, readContext.scurveProjection
    Dashboard_UpdateSCurveChartInPlace ws, readContext.scurveProjection
    Dashboard_UpdatePlanningOverviewContent ws, readContext.wbsTable, readContext.WbsColumns, readContext.CalcTable, readContext.calcColumns
    Dashboard_UpdateHotSpotsContent ws, readContext.wbsTable, readContext.WbsColumns, readContext.CalcTable, readContext.calcColumns

SafeExit:
    Application.ScreenUpdating = oldScreenUpdating
    If propagateErrors And failureNumber <> 0 Then
        On Error GoTo 0
        Err.Raise failureNumber, "Refresh_Dashboard_ContentOnly", failureDescription
    End If
    Exit Sub
ErrHandler:
    failureNumber = Err.number: failureDescription = Err.description
    Debug.Print "Refresh_Dashboard_ContentOnly error " & Err.number & ": " & Err.description
    Resume SafeExit
End Sub
'------------------------------------------------------------------------------
' FR: Active ou initialise Set Language dans l'etat runtime du composant.
' EN: Activates or initializes Set Language in the component runtime state.
'------------------------------------------------------------------------------

Public Sub Dashboard_SetLanguage(ByVal languageCode As String)

    Select Case UCase$(Trim$(languageCode))
        Case "FR"
            gDashboardLanguage = "FR"
        Case "EN"
            gDashboardLanguage = "EN"
        Case Else
            gDashboardLanguage = "EN"
    End Select

End Sub

'------------------------------------------------------------------------------
' FR: Actualise Apply Language sans modifier les regles metier qui produisent les donnees.
' EN: Refreshes Apply Language without changing the business rules that produce the data.
'------------------------------------------------------------------------------

Public Sub Dashboard_ApplyLanguage(Optional ByVal languageCode As String = "")

    If Trim$(languageCode) <> "" Then Dashboard_SetLanguage languageCode
    Refresh_Dashboard_TextsAndComparisonOnly

End Sub

'------------------------------------------------------------------------------
' FR: Reconnecte les selecteurs de snapshots a leur source possedee sans rendu.
' EN: Rebinds snapshot selectors to their owned source without rendering.
'------------------------------------------------------------------------------
Public Sub Dashboard_RebindSnapshotSelectors()

    Dim ws As Worksheet
    Dim selectedFromLabel As String
    Dim selectedToLabel As String

    Dashboard_SetLanguage Settings_GetOwnerLanguage("DASHBOARD")
    Set ws = Dashboard_EnsureSheet()
    selectedFromLabel = CStr(ws.Range("C4").Value2)
    selectedToLabel = CStr(ws.Range("G4").Value2)
    Dashboard_SetupSnapshotControls ws, selectedFromLabel, selectedToLabel, False

End Sub

'------------------------------------------------------------------------------
' FR: Bascule l'etat Dashboard Language et met a jour les sorties associees.
' EN: Toggles Dashboard Language state and updates related outputs.
'------------------------------------------------------------------------------
Public Sub Toggle_Dashboard_Language()
    Settings_ToggleDashboardLanguage
End Sub

'------------------------------------------------------------------------------
' FR: Reinitialise Reset Dashboard dans le perimetre possede par le composant.
' EN: Resets Reset Dashboard within the state owned by the component.
'------------------------------------------------------------------------------

Public Sub Reset_Dashboard()
    If Not WorkbookSchema_UserActionAllowed() Then Exit Sub

    Dim ws As Worksheet
    Dim oldScreenUpdating As Boolean
    Dim oldEvents As Boolean

    On Error GoTo SafeExit
    Dashboard_SetLanguage Settings_GetOwnerLanguage("DASHBOARD")

    oldScreenUpdating = Application.ScreenUpdating
    oldEvents = Application.EnableEvents
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    Set ws = Dashboard_EnsureSheet()

    Dashboard_ClearSnapshotsForReset
    Dashboard_ResetVisualCaches
    Dashboard_Clear ws
    Dashboard_SetupCanvas ws
    Dashboard_RenderEmptyShell ws

SafeExit:
    Application.EnableEvents = oldEvents
    Application.ScreenUpdating = oldScreenUpdating

End Sub
'------------------------------------------------------------------------------
' FR: Lance le workflow Dashboard Update.
' EN: Runs the Dashboard Update workflow.
'------------------------------------------------------------------------------
Public Sub Run_Dashboard_Update()

    Dim workflowStarted As Boolean
    Dim stopOnlyStarted As Boolean
    Dim consoleMessages As Collection
    Dim newSnapshot As ListRow
    Dim failureDescription As String
    Dim previousFrom As String, previousTo As String
On Error GoTo ErrHandler

    workflowStarted = EnsurePlanningWorkflowStarted("Run_Dashboard_Update")
    BeginPlanningEventRun "Run_Dashboard_Update"
    BeginPlanningWorkflowStopOnlyDisplay
    stopOnlyStarted = True
    Run_Full_Update True

    If CalcEngine_HasBlockingErrorsForState() Then GoTo CleanExit
    If IsMacroAbortRequested() Then GoTo CleanExit
    previousFrom = CStr(ThisWorkbook.Worksheets(DASHBOARD_SHEET).Range(DASH_FROM_CELL).Value2)
    previousTo = CStr(ThisWorkbook.Worksheets(DASHBOARD_SHEET).Range(DASH_TO_CELL).Value2)
    Set newSnapshot = Dashboard_CreateSnapshotFromCurrentData()
    Refresh_Dashboard_ContentOnly False, True
    ThisWorkbook.Worksheets(DASHBOARD_SHEET).Activate
    Set newSnapshot = Nothing

CleanExit:
    If stopOnlyStarted Then
        EndPlanningWorkflowStopOnlyDisplay
        stopOnlyStarted = False
    End If
    If IsMacroAbortRequested() Then ShowAbortMessageOnce
    If workflowStarted Then EndPlanningWorkflow
    Exit Sub

ErrHandler:
    failureDescription = Err.description
    If Not newSnapshot Is Nothing Then
        newSnapshot.Delete
        Set newSnapshot = Nothing
        Dashboard_SetupSnapshotControls ThisWorkbook.Worksheets(DASHBOARD_SHEET), previousFrom, previousTo, False
    End If
    If stopOnlyStarted Then
        EndPlanningWorkflowStopOnlyDisplay
        stopOnlyStarted = False
    End If

    Set consoleMessages = New Collection
    CalcBridge_AddConsoleMessage consoleMessages, "STOP", _
        PlanningMessageText_Format("DASHBOARD.ERROR.UPDATE", _
            TextCatalog_Arguments("Details", failureDescription), _
            TextCatalog_Arguments("Details", failureDescription))
    CalcBridge_ShowPlanningConsole consoleMessages
    Resume CleanExit

End Sub

'------------------------------------------------------------------------------
' FR: Cree ou remet en conformite la reference Sheet de maniere idempotente.
' EN: Creates or restores the Sheet reference to a compliant state idempotently.
'------------------------------------------------------------------------------

Private Function Dashboard_EnsureSheet() As Worksheet

    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(DASHBOARD_SHEET)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(after:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
        ws.name = DASHBOARD_SHEET
    End If

    Set Dashboard_EnsureSheet = ws

End Function

'------------------------------------------------------------------------------
' FR: Indique si la reference Usable Layout satisfait la condition attendue, sans modifier les donnees source.
' EN: Returns whether the Usable Layout reference satisfies the expected condition without mutating source data.
'------------------------------------------------------------------------------

Private Function Dashboard_HasUsableLayout() As Boolean

    Dim ws As Worksheet
    Dim chartObj As ChartObject

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(DASHBOARD_SHEET)
    If ws Is Nothing Then Exit Function

    Set chartObj = ws.ChartObjects(DASH_CHART_SCURVE)
    If chartObj Is Nothing Then Exit Function

    If Dashboard_GetShape(ws, "DASH_KPI_1_BG") Is Nothing Then Exit Function
    If Dashboard_GetShape(ws, "DASH_KPI_1_ACCENT") Is Nothing Then Exit Function
    If Dashboard_GetShape(ws, "DASH_KPI_1_TITLE") Is Nothing Then Exit Function
    If Dashboard_GetShape(ws, "DASH_KPI_1_HERO") Is Nothing Then Exit Function
    If Dashboard_GetShape(ws, "DASH_KPI_1_SUB") Is Nothing Then Exit Function

    Dashboard_HasUsableLayout = True
    On Error GoTo 0

End Function

'------------------------------------------------------------------------------
' FR: Reinitialise Clear dans le perimetre possede par le composant.
' EN: Resets Clear within the state owned by the component.
' FR - Effet de bord : efface uniquement les donnees ou objets cibles du contrat.
' EN - Side effect: clears only data or objects targeted by the contract.
'------------------------------------------------------------------------------

Private Sub Dashboard_Clear(ByVal ws As Worksheet)

    Dim i As Long

    ws.cells.Clear

    For i = ws.ChartObjects.count To 1 Step -1
        ws.ChartObjects(i).Delete
    Next i

    For i = ws.Shapes.count To 1 Step -1
        ws.Shapes(i).Delete
    Next i

End Sub

'------------------------------------------------------------------------------
' FR: Active ou initialise Setup Canvas dans l'etat runtime du composant.
' EN: Activates or initializes Setup Canvas in the component runtime state.
'------------------------------------------------------------------------------

Private Sub Dashboard_SetupCanvas(ByVal ws As Worksheet)

    Dim i As Long

    ws.cells.Interior.color = RGB(246, 248, 251)
    ws.cells.font.name = "Segoe UI"
    ws.cells.font.size = 10

    For i = 1 To 24
        ws.columns(i).ColumnWidth = 10
    Next i

    ws.columns("A").ColumnWidth = 2
    ws.columns("B:C").ColumnWidth = 11
    ws.columns("D:F").ColumnWidth = 12
    ws.columns("G:H").ColumnWidth = 11
    ws.columns("I:N").ColumnWidth = 12
    ws.columns("O:W").ColumnWidth = 11

    For i = 1 To 62
        ws.rows(i).rowHeight = 18
    Next i

    ws.Range("A1:W62").Interior.color = RGB(246, 248, 251)
    activeWindow.DisplayGridlines = False

    With ws.PageSetup
        .PrintArea = ""
    End With

End Sub

'------------------------------------------------------------------------------
' FR: Rafraichit Dashboard Comparison a partir de l'etat courant.
' EN: Refreshes Dashboard Comparison from the current state.
'------------------------------------------------------------------------------
Public Sub Refresh_Dashboard_Comparison()

    On Error GoTo ErrHandler

    Refresh_Dashboard_TextsAndComparisonOnly
    ThisWorkbook.Worksheets(DASHBOARD_SHEET).Activate
    Exit Sub

ErrHandler:
    CalcBridge_ShowSingleConsoleMessage "WARNING", "DASHBOARD.ERROR.COMPARISON", _
        TextCatalog_Arguments("Details", Err.description)

End Sub

Public Sub Dashboard_HandleSnapshotSelectionChange(ByVal ws As Worksheet, ByVal Target As Range)
    Dim selected As Range, cell As Range
    Dim snapshots As ListObject, label As String
    Dim oldEvents As Boolean
    oldEvents = Application.EnableEvents
    On Error GoTo Failed
    If Not (ws.parent Is ThisWorkbook) Then Exit Sub
    Set selected = Application.Intersect(Target, ws.Range(DASH_FROM_CELL & "," & DASH_TO_CELL))
    If selected Is Nothing Then Exit Sub
    Set snapshots = ThisWorkbook.Worksheets(DASHBOARD_SNAPSHOT_SHEET).ListObjects(DASHBOARD_SNAPSHOT_TABLE)
    For Each cell In selected.cells
        ' UI reconstruction may temporarily remove validation; the owned history remains authoritative.
        label = CStr(cell.Value2)
        If Len(label) > 0 Then
            If Not Dashboard_SnapshotLabelExists(snapshots, label) Then
                Application.EnableEvents = False
                Application.Undo
                Application.EnableEvents = oldEvents
                CalcBridge_ShowSingleConsoleMessage "STOP", "DASHBOARD.VALIDATION.SNAPSHOT.ERROR"
                Exit Sub
            End If
        End If
    Next cell
    Exit Sub
Failed:
    Application.EnableEvents = oldEvents
    CalcBridge_ShowSingleConsoleMessage "STOP", "DASHBOARD.ERROR.COMPARISON", _
        TextCatalog_Arguments("Details", Err.description)
End Sub

'------------------------------------------------------------------------------
' FR: Rafraichit Dashboard Texts And Comparison Only a partir de l'etat courant.
' EN: Refreshes Dashboard Texts And Comparison Only from the current state.
'------------------------------------------------------------------------------
Private Sub Refresh_Dashboard_TextsAndComparisonOnly()

    Dim ws As Worksheet
    Dim readContext As clsDashboardReadContext
    Dim fromLabel As String
    Dim toLabel As String

    On Error GoTo ErrHandler
    Set ws = ThisWorkbook.Worksheets(DASHBOARD_SHEET)
    fromLabel = CStr(ws.Range(DASH_FROM_CELL).value)
    toLabel = CStr(ws.Range(DASH_TO_CELL).value)
    Set readContext = DashboardReadContext_Load()

    Dashboard_UpdateHeaderTexts ws
    Dashboard_UpdateHeaderTimestamp ws
    Dashboard_SetupSnapshotControls ws, fromLabel, toLabel, False
    Dashboard_EnsureResetButton ws
    Dashboard_RemoveLegacyLanguageToggle ws
    Dashboard_UpdateKnownShapeTexts ws
    Dashboard_UpdateKpiCardsInPlace ws, readContext.wbsTable, readContext.WbsColumns, readContext.CalcTable, readContext.calcColumns, readContext.scurveProjection
    Dashboard_UpdateSCurveChartInPlace ws, readContext.scurveProjection
    Dashboard_UpdatePlanningOverviewContent ws, readContext.wbsTable, readContext.WbsColumns, readContext.CalcTable, readContext.calcColumns
    Dashboard_UpdateHotSpotsContent ws, readContext.wbsTable, readContext.WbsColumns, readContext.CalcTable, readContext.calcColumns

SafeExit:
    Exit Sub
ErrHandler:
    Debug.Print "Refresh_Dashboard_TextsAndComparisonOnly error " & Err.number & ": " & Err.description
    Err.Raise Err.number, "Refresh_Dashboard_TextsAndComparisonOnly", Err.description
End Sub

'------------------------------------------------------------------------------
' FR: Reinitialise Clear Snapshots For Reset dans le perimetre possede par le composant.
' EN: Resets Clear Snapshots For Reset within the state owned by the component.
' FR - Effet de bord : efface uniquement les donnees ou objets cibles du contrat.
' EN - Side effect: clears only data or objects targeted by the contract.
'------------------------------------------------------------------------------

Private Sub Dashboard_ClearSnapshotsForReset()

    Dim tbl As ListObject

    Set tbl = Dashboard_EnsureSnapshotsTable()

    If Not tbl Is Nothing Then
        Do While tbl.ListRows.count > 0
            tbl.ListRows(tbl.ListRows.count).Delete
        Loop
    End If

    On Error Resume Next
    ThisWorkbook.names("DashboardSnapshotLabels").Delete
    On Error GoTo 0

End Sub

'------------------------------------------------------------------------------
' FR: Reinitialise Reset Visual Caches dans le perimetre possede par le composant.
' EN: Resets Reset Visual Caches within the state owned by the component.
'------------------------------------------------------------------------------

Private Sub Dashboard_ResetVisualCaches()

    gHotSpotHelper = vbNullString
    gHotSpotStep = vbNullString

End Sub

'------------------------------------------------------------------------------
' FR: Actualise Render Empty Shell sans modifier les regles metier qui produisent les donnees.
' EN: Refreshes Render Empty Shell without changing the business rules that produce the data.
'------------------------------------------------------------------------------

Private Sub Dashboard_RenderEmptyShell(ByVal ws As Worksheet)

    Dashboard_RenderHeader ws
    Dashboard_RenderExecutiveSummaryEmpty ws
    Dashboard_RenderSCurveEmpty ws
    Dashboard_RenderPlanningOverviewEmpty ws
    Dashboard_RenderHotSpotsEmpty ws

End Sub

'------------------------------------------------------------------------------
' FR: Actualise Render Executive Summary Empty sans modifier les regles metier qui produisent les donnees.
' EN: Refreshes Render Executive Summary Empty without changing the business rules that produce the data.
'------------------------------------------------------------------------------

Private Sub Dashboard_RenderExecutiveSummaryEmpty(ByVal ws As Worksheet)

    Dim leftPos As Double
    Dim topPos As Double
    Dim widthVal As Double
    Dim cardWidth As Double
    Dim gapVal As Double

    leftPos = ws.Range("B5").left
    topPos = ws.Range("B7").top
    widthVal = ws.Range("B5:Q29").width
    gapVal = 12
    cardWidth = (widthVal - (gapVal * 3)) / 4

    ws.Range("B6:M6").Merge
    ws.Range("B6").value = TextCatalog_Get("DASHBOARD.TITLE.EXECUTIVE_SUMMARY", Dashboard_CurrentLanguage())
    ws.Range("B6").font.bold = True
    ws.Range("B6").font.size = 15
    ws.Range("B6").font.color = RGB(20, 34, 51)

    With ws.Range("N6:Q6")
        .Merge
        .value = Dashboard_NoProjectLoadedText()
        .font.size = 9
        .font.color = RGB(96, 111, 128)
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlCenter
    End With

    Dashboard_UpdateKpiCard ws, 1, TextCatalog_Get("DASHBOARD.KPI.PROGRESS", Dashboard_CurrentLanguage()), Dashboard_NoDataText(), _
        Dashboard_NoProjectLoadedText(), _
        leftPos + 18, topPos, cardWidth - 10, 78, RGB(160, 170, 181)

    Dashboard_UpdateKpiCard ws, 2, TextCatalog_Get("DASHBOARD.KPI.FORECAST_FINISH", Dashboard_CurrentLanguage()), Dashboard_NoDataText(), _
        Dashboard_NoProjectLoadedText(), _
        leftPos + 18 + cardWidth + gapVal, topPos, cardWidth - 10, 78, RGB(160, 170, 181)

    Dashboard_UpdateKpiCard ws, 3, TextCatalog_Get("DASHBOARD.KPI.CRITICAL_ACTIVITIES", Dashboard_CurrentLanguage()), Dashboard_NoDataText(), _
        Dashboard_NoProjectLoadedText(), _
        leftPos + 18 + ((cardWidth + gapVal) * 2), topPos, cardWidth - 10, 78, RGB(160, 170, 181)

    Dashboard_UpdateKpiCard ws, 4, TextCatalog_Get("DASHBOARD.KPI.MOMENTUM", Dashboard_CurrentLanguage()), Dashboard_NoDataText(), _
        Dashboard_NoProjectLoadedText(), _
        leftPos + 18 + ((cardWidth + gapVal) * 3), topPos, cardWidth - 10, 78, RGB(160, 170, 181)

End Sub

'------------------------------------------------------------------------------
' FR: Actualise Render S Curve Empty sans modifier les regles metier qui produisent les donnees.
' EN: Refreshes Render S Curve Empty without changing the business rules that produce the data.
'------------------------------------------------------------------------------

Private Sub Dashboard_RenderSCurveEmpty(ByVal ws As Worksheet)

    Dim leftPos As Double
    Dim topPos As Double
    Dim widthVal As Double
    Dim heightVal As Double

    leftPos = ws.Range("B13").left + 18
    topPos = ws.Range("B13").top
    widthVal = (ws.Range("I13").left + ws.Range("I13").width) - leftPos
    heightVal = ws.Range("B13:I29").height

    Dashboard_AddPanel ws, leftPos, topPos, widthVal, heightVal, RGB(255, 255, 255)
    Dashboard_AddCenteredSectionTitle ws, TextCatalog_Get("DASHBOARD.SECTION.SCURVE", Dashboard_CurrentLanguage()), leftPos + 22, topPos + 10, widthVal - 44
    Dashboard_WriteDashboardEmptyState ws, leftPos + 22, topPos + 70, widthVal - 44

End Sub

'------------------------------------------------------------------------------
' FR: Actualise Render Planning Overview Empty sans modifier les regles metier qui produisent les donnees.
' EN: Refreshes Render Planning Overview Empty without changing the business rules that produce the data.
'------------------------------------------------------------------------------

Private Sub Dashboard_RenderPlanningOverviewEmpty(ByVal ws As Worksheet)

    Dim leftPos As Double
    Dim topPos As Double
    Dim widthVal As Double
    Dim heightVal As Double

    leftPos = ws.Range("J13").left + 8
    topPos = ws.Range("J13").top
    widthVal = ws.Range("J13:Q29").width - 20
    heightVal = ws.Range("J13:Q29").height

    Dashboard_AddPanel ws, leftPos, topPos, widthVal, heightVal, RGB(255, 255, 255)
    Dashboard_AddCenteredSectionTitle ws, TextCatalog_Get("DASHBOARD.SECTION.OVERVIEW", Dashboard_CurrentLanguage()), leftPos + 22, topPos + 10, widthVal - 44
    Dashboard_WriteDashboardEmptyState ws, leftPos + 22, topPos + 70, widthVal - 44

End Sub

'------------------------------------------------------------------------------
' FR: Actualise Render Hot Spots Empty sans modifier les regles metier qui produisent les donnees.
' EN: Refreshes Render Hot Spots Empty without changing the business rules that produce the data.
'------------------------------------------------------------------------------

Private Sub Dashboard_RenderHotSpotsEmpty(ByVal ws As Worksheet)

    Dim leftPos As Double
    Dim topPos As Double
    Dim widthVal As Double
    Dim heightVal As Double
    Dim cardTop As Double
    Dim cardHeight As Double
    Dim gapVal As Double
    Dim rightW As Double
    Dim largeW As Double
    Dim smallH As Double

    leftPos = ws.Range("B32").left
    topPos = ws.Range("B32").top
    widthVal = ws.Range("B32:Q56").width
    heightVal = ws.Range("B32:Q56").height

    Dashboard_AddSectionTitle ws, TextCatalog_Get("DASHBOARD.SECTION.HOT_SPOTS", Dashboard_CurrentLanguage()), leftPos + 22, topPos + 10, widthVal - 36

    cardTop = topPos + 48
    cardHeight = heightVal - 64
    gapVal = 14
    rightW = 240
    largeW = (widthVal - 44 - rightW - (gapVal * 2)) / 2
    smallH = (cardHeight - gapVal) / 2

    Dashboard_AddEmptyHotSpotCard ws, leftPos + 22, cardTop, largeW, cardHeight, TextCatalog_Get("DASHBOARD.CARD.TOP_DELAYS", Dashboard_CurrentLanguage())
    Dashboard_AddEmptyHotSpotCard ws, leftPos + 22 + largeW + gapVal, cardTop, largeW, cardHeight, TextCatalog_Get("DASHBOARD.CARD.DEADLINES", Dashboard_CurrentLanguage())
    Dashboard_AddEmptyHotSpotCard ws, leftPos + 22 + (largeW * 2) + (gapVal * 2), cardTop, rightW, smallH, TextCatalog_Get("DASHBOARD.CARD.NEXT_MILESTONE", Dashboard_CurrentLanguage())
    Dashboard_AddEmptyHotSpotCard ws, leftPos + 22 + (largeW * 2) + (gapVal * 2), cardTop + smallH + gapVal, rightW, smallH, TextCatalog_Get("DASHBOARD.CARD.NEXT_CRITICAL", Dashboard_CurrentLanguage())

End Sub
'------------------------------------------------------------------------------
' FR: Actualise Update Header Texts sans modifier les regles metier qui produisent les donnees.
' EN: Refreshes Update Header Texts without changing the business rules that produce the data.
' FR - Effet de bord : cree ou met a jour des shapes Excel.
' EN - Side effect: creates or updates Excel shapes.
'------------------------------------------------------------------------------

Private Sub Dashboard_UpdateHeaderTexts(ByVal ws As Worksheet)

    With ws.Range("B1:P2")
        .value = TextCatalog_Get("DASHBOARD.TITLE.PAGE", Dashboard_CurrentLanguage())
    End With

    With ws.Range("B3:P3")
        .value = TextCatalog_Get("DASHBOARD.TITLE.SUBTITLE", Dashboard_CurrentLanguage())
    End With

    On Error Resume Next
    ws.Shapes("btn_Update_Dashboard").TextFrame2.TextRange.text = TextCatalog_Get("DASHBOARD.COMMAND.SNAPSHOT", Dashboard_CurrentLanguage())
    ws.Shapes("btn_Dashboard_Refresh_Comparison").TextFrame2.TextRange.text = TextCatalog_Get("DASHBOARD.COMMAND.COMPARE", Dashboard_CurrentLanguage())
    ws.Shapes("btn_Dashboard_Reset").TextFrame2.TextRange.text = TextCatalog_Get("DASHBOARD.COMMAND.CLEAN", Dashboard_CurrentLanguage())
    On Error GoTo 0

End Sub

'------------------------------------------------------------------------------
' FR: Actualise Update Header Timestamp sans modifier les regles metier qui produisent les donnees.
' EN: Refreshes Update Header Timestamp without changing the business rules that produce the data.
'------------------------------------------------------------------------------

Private Sub Dashboard_UpdateHeaderTimestamp(ByVal ws As Worksheet)

    Dim shp As shape
    Dim tsText As String
    Dim found As Boolean

    tsText = Dashboard_FormatTimestamp(Now)

    On Error Resume Next
    For Each shp In ws.Shapes
        If shp.TextFrame2.HasText Then
            If shp.top < ws.Range("A4").top Then
                If Dashboard_IsTimestampText(Trim$(shp.TextFrame2.TextRange.text)) Then
                    shp.TextFrame2.TextRange.text = tsText
                    found = True
                    Exit For
                End If
            End If
        End If
    Next shp
    On Error GoTo 0

    If Not found Then
        Dashboard_AddBadge ws, tsText, ws.Range("T1").left, ws.Range("T1").top + 5, 150, 22, RGB(226, 240, 255), RGB(29, 91, 158)
    End If

End Sub
'------------------------------------------------------------------------------
' FR: Reinitialise Remove Legacy Language Toggle dans le perimetre possede par le composant.
' EN: Resets Remove Legacy Language Toggle within the state owned by the component.
' FR - Effet de bord : efface uniquement les donnees ou objets cibles du contrat.
' EN - Side effect: clears only data or objects targeted by the contract.
'------------------------------------------------------------------------------

Private Sub Dashboard_RemoveLegacyLanguageToggle(ByVal ws As Worksheet)

    If ws Is Nothing Then Exit Sub

    On Error Resume Next
    ws.Shapes("DASH_Language_Label").Delete
    ws.Shapes("DASH_Language_BG").Delete
    ws.Shapes("DASH_Language_Knob").Delete
    On Error GoTo 0

End Sub

'------------------------------------------------------------------------------
' FR: Actualise Update Language Toggle sans modifier les regles metier qui produisent les donnees.
' EN: Refreshes Update Language Toggle without changing the business rules that produce the data.
' FR - Effet de bord : cree ou met a jour des shapes Excel.
' EN - Side effect: creates or updates Excel shapes.
'------------------------------------------------------------------------------

Private Sub Dashboard_UpdateLanguageToggle(ByVal ws As Worksheet)

    Dim trackShape As shape
    Dim knobShape As shape
    Dim isOn As Boolean
    Dim knobSize As Double
    Dim knobLeft As Double

    On Error Resume Next
    Set trackShape = ws.Shapes("DASH_Language_BG")
    Set knobShape = ws.Shapes("DASH_Language_Knob")
    On Error GoTo 0
    If trackShape Is Nothing Or knobShape Is Nothing Then Exit Sub

    isOn = Not Dashboard_IsFrench()
    Dashboard_FormatSwitchTrack trackShape, isOn

    knobSize = knobShape.width
    If isOn Then
        knobLeft = trackShape.left + trackShape.width - knobSize - 2
    Else
        knobLeft = trackShape.left + 2
    End If
    knobShape.left = knobLeft
    knobShape.top = trackShape.top + ((trackShape.height - knobSize) / 2)

End Sub

'------------------------------------------------------------------------------
' FR: Reinitialise Clear Kpi Shapes dans le perimetre possede par le composant.
' EN: Resets Clear Kpi Shapes within the state owned by the component.
' FR - Effet de bord : efface uniquement les donnees ou objets cibles du contrat.
' EN - Side effect: clears only data or objects targeted by the contract.
'------------------------------------------------------------------------------

Private Sub Dashboard_ClearKpiShapes(ByVal ws As Worksheet)

    Dim i As Long
    Dim shp As shape
    Dim kpiLeft As Double
    Dim kpiRight As Double
    Dim kpiTop As Double
    Dim kpiBottom As Double

    kpiLeft = ws.Range("B7").left - 4
    kpiRight = ws.Range("Q7").left + ws.Range("Q7").width + 4
    kpiTop = ws.Range("B7").top - 6
    kpiBottom = ws.Range("B7").top + 90

    For i = ws.Shapes.count To 1 Step -1
        Set shp = ws.Shapes(i)
        If left$(shp.name, Len(DASH_PREFIX & "KPI")) = DASH_PREFIX & "KPI" Or _
            (shp.left >= kpiLeft And shp.left <= kpiRight And shp.top >= kpiTop And shp.top <= kpiBottom) Then
            shp.Delete
        End If
    Next i

End Sub

'------------------------------------------------------------------------------
' FR: Actualise Update Known Shape Texts sans modifier les regles metier qui produisent les donnees.
' EN: Refreshes Update Known Shape Texts without changing the business rules that produce the data.
'------------------------------------------------------------------------------

Private Sub Dashboard_UpdateKnownShapeTexts(ByVal ws As Worksheet)

    Dim shp As shape
    Dim txt As String

    On Error Resume Next
    For Each shp In ws.Shapes
        If shp.TextFrame2.HasText Then
            txt = Trim$(shp.TextFrame2.TextRange.text)
            Select Case txt
                Case TextCatalog_Get("DASHBOARD.SECTION.SCURVE", "FR"), TextCatalog_Get("DASHBOARD.SECTION.SCURVE", "EN")
                    shp.TextFrame2.TextRange.text = TextCatalog_Get("DASHBOARD.SECTION.SCURVE", Dashboard_CurrentLanguage())
                Case TextCatalog_Get("DASHBOARD.SECTION.OVERVIEW", "FR"), TextCatalog_Get("DASHBOARD.SECTION.OVERVIEW", "EN")
                    shp.TextFrame2.TextRange.text = TextCatalog_Get("DASHBOARD.SECTION.OVERVIEW", Dashboard_CurrentLanguage())
                Case TextCatalog_Get("DASHBOARD.SECTION.HOT_SPOTS", "FR"), TextCatalog_Get("DASHBOARD.SECTION.HOT_SPOTS", "EN")
                    shp.TextFrame2.TextRange.text = TextCatalog_Get("DASHBOARD.SECTION.HOT_SPOTS", Dashboard_CurrentLanguage())
                Case TextCatalog_Get("DASHBOARD.CARD.TOP_DELAYS", "FR"), TextCatalog_Get("DASHBOARD.CARD.TOP_DELAYS", "EN"), TextCatalog_Get("DASHBOARD.LEGACY_ALIAS.TOP_DELAYS", "FR")
                    shp.TextFrame2.TextRange.text = TextCatalog_Get("DASHBOARD.CARD.TOP_DELAYS", Dashboard_CurrentLanguage())
                Case TextCatalog_Get("DASHBOARD.CARD.DEADLINE_RISKS", "FR"), TextCatalog_Get("DASHBOARD.CARD.DEADLINE_RISKS", "EN")
                    shp.TextFrame2.TextRange.text = TextCatalog_Get("DASHBOARD.CARD.DEADLINE_RISKS", Dashboard_CurrentLanguage())
                Case TextCatalog_Get("DASHBOARD.CARD.FORECAST_ISSUES", "FR"), TextCatalog_Get("DASHBOARD.CARD.FORECAST_ISSUES", "EN")
                    shp.TextFrame2.TextRange.text = TextCatalog_Get("DASHBOARD.CARD.FORECAST_ISSUES", Dashboard_CurrentLanguage())
                Case TextCatalog_Get("DASHBOARD.EMPTY.NO_DELAYS", "FR"), TextCatalog_Get("DASHBOARD.EMPTY.NO_DELAYS", "EN"), TextCatalog_Get("DASHBOARD.LEGACY_ALIAS.NO_DELAYS", "FR")
                    shp.TextFrame2.TextRange.text = TextCatalog_Get("DASHBOARD.EMPTY.NO_DELAYS", Dashboard_CurrentLanguage())
                Case TextCatalog_Get("DASHBOARD.EMPTY.NO_DEADLINE_RISKS", "FR"), TextCatalog_Get("DASHBOARD.EMPTY.NO_DEADLINE_RISKS", "EN"), TextCatalog_Get("DASHBOARD.LEGACY_ALIAS.NO_DEADLINE_RISKS", "FR")
                    shp.TextFrame2.TextRange.text = TextCatalog_Get("DASHBOARD.EMPTY.NO_DEADLINE_RISKS", Dashboard_CurrentLanguage())
                Case TextCatalog_Get("DASHBOARD.EMPTY.NO_FORECAST_ISSUES", "FR"), TextCatalog_Get("DASHBOARD.EMPTY.NO_FORECAST_ISSUES", "EN"), TextCatalog_Get("DASHBOARD.LEGACY_ALIAS.NO_FORECAST_ISSUES", "FR")
                    shp.TextFrame2.TextRange.text = TextCatalog_Get("DASHBOARD.EMPTY.NO_FORECAST_ISSUES", Dashboard_CurrentLanguage())
                Case TextCatalog_Get("DASHBOARD.EMPTY.NO_PROJECT", "FR"), TextCatalog_Get("DASHBOARD.EMPTY.NO_PROJECT", "EN"), TextCatalog_Get("DASHBOARD.LEGACY_ALIAS.NO_PROJECT", "FR")
                    shp.TextFrame2.TextRange.text = Dashboard_NoProjectLoadedText()
                Case TextCatalog_Get("DASHBOARD.EMPTY.NO_DATA", "FR"), TextCatalog_Get("DASHBOARD.EMPTY.NO_DATA", "EN"), TextCatalog_Get("DASHBOARD.LEGACY_ALIAS.NO_DATA", "FR")
                    shp.TextFrame2.TextRange.text = Dashboard_NoDataText()
            End Select
        End If
    Next shp
    On Error GoTo 0

End Sub

'------------------------------------------------------------------------------
' FR: Actualise Render Header sans modifier les regles metier qui produisent les donnees.
' EN: Refreshes Render Header without changing the business rules that produce the data.
'------------------------------------------------------------------------------

Private Sub Dashboard_RenderHeader(ByVal ws As Worksheet, Optional ByVal selectedFromLabel As String = "", Optional ByVal selectedToLabel As String = "")

    With ws.Range("B1:P2")
        .Merge
        .value = TextCatalog_Get("DASHBOARD.TITLE.PAGE", Dashboard_CurrentLanguage())
        .font.size = 24
        .font.bold = True
        .font.color = RGB(20, 34, 51)
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
    End With

    With ws.Range("B3:P3")
        .Merge
        .value = TextCatalog_Get("DASHBOARD.TITLE.SUBTITLE", Dashboard_CurrentLanguage())
        .font.size = 10
        .font.color = RGB(96, 111, 128)
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
    End With

    Dashboard_AddUpdateButton ws, ws.Range("Q1").left, ws.Range("Q1").top + 4, 126, 24
    Dashboard_AddBadge ws, Dashboard_FormatTimestamp(Now), ws.Range("T1").left, ws.Range("T1").top + 5, 150, 22, RGB(226, 240, 255), RGB(29, 91, 158)
    Dashboard_SetupSnapshotControls ws, selectedFromLabel, selectedToLabel, True
    Dashboard_AddResetButton ws, ws.Range("Q3").left, ws.Range("Q3").top + 4, 126, 24
    Dashboard_RemoveLegacyLanguageToggle ws

End Sub

'------------------------------------------------------------------------------
' FR: Ajoute la reference Update Button a la structure cible fournie par l'appelant.
' EN: Adds the Update Button reference to the target structure supplied by the caller.
' FR - Effet de bord : cree ou met a jour des shapes Excel.
' EN - Side effect: creates or updates Excel shapes.
'------------------------------------------------------------------------------

Private Sub Dashboard_AddUpdateButton(ByVal ws As Worksheet, ByVal X As Double, ByVal Y As Double, ByVal w As Double, ByVal h As Double)

    Dim btn As shape

    Set btn = ws.Shapes.AddShape(msoShapeRoundedRectangle, X, Y, w, h)
    btn.name = "btn_Update_Dashboard"
    btn.fill.ForeColor.RGB = RGB(20, 34, 51)
    btn.Line.Visible = msoFalse
    btn.OnAction = "Run_Dashboard_Update"

    With btn.TextFrame2
        .MarginLeft = 8
        .MarginRight = 8
        .MarginTop = 2
        .MarginBottom = 2
        .TextRange.text = TextCatalog_Get("DASHBOARD.COMMAND.SNAPSHOT", Dashboard_CurrentLanguage())
        .TextRange.font.name = "Segoe UI"
        .TextRange.font.size = 9
        .TextRange.font.bold = msoTrue
        .TextRange.font.fill.ForeColor.RGB = RGB(255, 255, 255)
        .TextRange.ParagraphFormat.alignment = msoAlignCenter
        .VerticalAnchor = msoAnchorMiddle
    End With

End Sub

'------------------------------------------------------------------------------
' FR: Ajoute la reference Language Toggle a la structure cible fournie par l'appelant.
' EN: Adds the Language Toggle reference to the target structure supplied by the caller.
' FR - Effet de bord : cree ou met a jour des shapes Excel.
' EN - Side effect: creates or updates Excel shapes.
'------------------------------------------------------------------------------

Private Sub Dashboard_AddLanguageToggle(ByVal ws As Worksheet, ByVal X As Double, ByVal Y As Double, ByVal w As Double, ByVal h As Double)

    Dim labelShape As shape
    Dim trackShape As shape
    Dim knobShape As shape
    Dim labelW As Double
    Dim trackW As Double
    Dim trackH As Double
    Dim knobSize As Double
    Dim trackLeft As Double
    Dim knobLeft As Double
    Dim knobTop As Double
    Dim isOn As Boolean

    labelW = 56
    trackW = 36
    trackH = 14
    knobSize = 10
    trackLeft = X + labelW
    isOn = Not Dashboard_IsFrench()

    Set labelShape = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, X, Y, labelW, trackH)
    labelShape.name = "DASH_Language_Label"
    labelShape.OnAction = "Toggle_Dashboard_Language"
    labelShape.Placement = xlMove
    Dashboard_FormatSwitchLabel labelShape, TextCatalog_Get("DASHBOARD.LANGUAGE.SWITCH", Dashboard_CurrentLanguage())

    Set trackShape = ws.Shapes.AddShape(msoShapeRoundedRectangle, trackLeft, Y, trackW, trackH)
    trackShape.name = "DASH_Language_BG"
    trackShape.OnAction = "Toggle_Dashboard_Language"
    trackShape.Placement = xlMove
    trackShape.Adjustments.item(1) = 0.5
    Dashboard_FormatSwitchTrack trackShape, isOn

    knobTop = Y + ((trackH - knobSize) / 2)
    If isOn Then
        knobLeft = trackLeft + trackW - knobSize - 2
    Else
        knobLeft = trackLeft + 2
    End If

    Set knobShape = ws.Shapes.AddShape(msoShapeOval, knobLeft, knobTop, knobSize, knobSize)
    knobShape.name = "DASH_Language_Knob"
    knobShape.OnAction = "Toggle_Dashboard_Language"
    knobShape.Placement = xlMove
    Dashboard_FormatSwitchKnob knobShape

End Sub

'------------------------------------------------------------------------------
' FR: Normalise ou formate Format Switch Label selon le contrat canonique du composant.
' EN: Normalizes or formats Format Switch Label according to the component contract.
'------------------------------------------------------------------------------

Private Sub Dashboard_FormatSwitchLabel(ByVal shp As shape, ByVal caption As String)

    With shp
        .Line.Visible = msoFalse
        .fill.Visible = msoFalse
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
        .TextFrame2.MarginLeft = 0
        .TextFrame2.MarginRight = 0
        .TextFrame2.MarginTop = 0
        .TextFrame2.MarginBottom = 0
        .TextFrame2.TextRange.text = caption
        .TextFrame2.TextRange.font.name = "Segoe UI"
        .TextFrame2.TextRange.font.size = 9
        .TextFrame2.TextRange.font.bold = msoTrue
        .TextFrame2.TextRange.font.fill.ForeColor.RGB = RGB(68, 114, 196)
        .TextFrame2.TextRange.ParagraphFormat.alignment = msoAlignCenter
    End With

End Sub

'------------------------------------------------------------------------------
' FR: Normalise ou formate Format Switch Track selon le contrat canonique du composant.
' EN: Normalizes or formats Format Switch Track according to the component contract.
'------------------------------------------------------------------------------

Private Sub Dashboard_FormatSwitchTrack(ByVal shp As shape, ByVal isOn As Boolean)

    With shp
        .Line.weight = 1
        .Shadow.Visible = msoFalse
        If isOn Then
            .fill.ForeColor.RGB = RGB(68, 114, 196)
            .Line.ForeColor.RGB = RGB(68, 114, 196)
        Else
            .fill.ForeColor.RGB = RGB(230, 230, 230)
            .Line.ForeColor.RGB = RGB(170, 170, 170)
        End If
    End With

End Sub

'------------------------------------------------------------------------------
' FR: Normalise ou formate Format Switch Knob selon le contrat canonique du composant.
' EN: Normalizes or formats Format Switch Knob according to the component contract.
'------------------------------------------------------------------------------

Private Sub Dashboard_FormatSwitchKnob(ByVal shp As shape)

    With shp
        .fill.ForeColor.RGB = RGB(255, 255, 255)
        .Line.ForeColor.RGB = RGB(150, 150, 150)
        .Line.weight = 0.75
        .Shadow.Visible = msoFalse
    End With

End Sub

'------------------------------------------------------------------------------
' FR: Ajoute la reference Refresh Comparison Button a la structure cible fournie par l'appelant.
' EN: Adds the Refresh Comparison Button reference to the target structure supplied by the caller.
' FR - Effet de bord : cree ou met a jour des shapes Excel.
' EN - Side effect: creates or updates Excel shapes.
'------------------------------------------------------------------------------

Private Sub Dashboard_AddRefreshComparisonButton(ByVal ws As Worksheet, ByVal X As Double, ByVal Y As Double, ByVal w As Double, ByVal h As Double)

    Dim btn As shape

    Set btn = ws.Shapes.AddShape(msoShapeRoundedRectangle, X, Y, w, h)
    btn.name = "btn_Dashboard_Refresh_Comparison"
    btn.fill.ForeColor.RGB = RGB(96, 111, 128)
    btn.Line.Visible = msoFalse
    btn.OnAction = "Refresh_Dashboard_Comparison"

    With btn.TextFrame2
        .MarginLeft = 8
        .MarginRight = 8
        .MarginTop = 2
        .MarginBottom = 2
        .TextRange.text = TextCatalog_Get("DASHBOARD.COMMAND.COMPARE", Dashboard_CurrentLanguage())
        .TextRange.font.name = "Segoe UI"
        .TextRange.font.size = 9
        .TextRange.font.bold = msoTrue
        .TextRange.font.fill.ForeColor.RGB = RGB(255, 255, 255)
        .TextRange.ParagraphFormat.alignment = msoAlignCenter
        .VerticalAnchor = msoAnchorMiddle
    End With

End Sub

'------------------------------------------------------------------------------
' FR: Ajoute la reference Reset Button a la structure cible fournie par l'appelant.
' EN: Adds the Reset Button reference to the target structure supplied by the caller.
' FR - Effet de bord : cree ou met a jour des shapes Excel.
' EN - Side effect: creates or updates Excel shapes.
'------------------------------------------------------------------------------

Private Sub Dashboard_AddResetButton(ByVal ws As Worksheet, ByVal X As Double, ByVal Y As Double, ByVal w As Double, ByVal h As Double)

    Dim btn As shape

    Set btn = ws.Shapes.AddShape(msoShapeRoundedRectangle, X, Y, w, h)
    btn.name = "btn_Dashboard_Reset"
    btn.fill.ForeColor.RGB = RGB(192, 80, 77)
    btn.Line.Visible = msoFalse
    btn.OnAction = "Reset_Dashboard"

    With btn.TextFrame2
        .MarginLeft = 8
        .MarginRight = 8
        .MarginTop = 2
        .MarginBottom = 2
        .TextRange.text = TextCatalog_Get("DASHBOARD.COMMAND.CLEAN", Dashboard_CurrentLanguage())
        .TextRange.font.name = "Segoe UI"
        .TextRange.font.size = 9
        .TextRange.font.bold = msoTrue
        .TextRange.font.fill.ForeColor.RGB = RGB(255, 255, 255)
        .TextRange.ParagraphFormat.alignment = msoAlignCenter
        .VerticalAnchor = msoAnchorMiddle
    End With

End Sub

'------------------------------------------------------------------------------
' FR: Cree ou remet en conformite la reference Reset Button de maniere idempotente.
' EN: Creates or restores the Reset Button reference to a compliant state idempotently.
'------------------------------------------------------------------------------

Private Sub Dashboard_EnsureResetButton(ByVal ws As Worksheet)

    Dim btn As shape

    Set btn = Dashboard_GetShape(ws, "btn_Dashboard_Reset")
    If btn Is Nothing Then
        Dashboard_AddResetButton ws, ws.Range("Q3").left, ws.Range("Q3").top + 4, 126, 24
    Else
        Dashboard_PositionKpiShape btn, ws.Range("Q3").left, ws.Range("Q3").top + 4, 126, 24
        btn.fill.ForeColor.RGB = RGB(192, 80, 77)
        btn.Line.Visible = msoFalse
        btn.Visible = msoTrue
        btn.OnAction = "Reset_Dashboard"
        btn.TextFrame2.TextRange.text = TextCatalog_Get("DASHBOARD.COMMAND.CLEAN", Dashboard_CurrentLanguage())
    End If

End Sub
'------------------------------------------------------------------------------
' FR: Active ou initialise Setup Snapshot Controls dans l'etat runtime du composant.
' EN: Activates or initializes Setup Snapshot Controls in the component runtime state.
' FR - Effet de bord : efface uniquement les donnees ou objets cibles du contrat.
' EN - Side effect: clears only data or objects targeted by the contract.
'------------------------------------------------------------------------------

Private Sub Dashboard_SetupSnapshotControls(ByVal ws As Worksheet, ByVal selectedFromLabel As String, ByVal selectedToLabel As String, Optional ByVal recreateButton As Boolean = True)

    Dim tbl As ListObject
    Dim fromDefault As String
    Dim toDefault As String

    Set tbl = Dashboard_EnsureSnapshotsTable()
    Dashboard_GetDefaultSnapshotLabels tbl, fromDefault, toDefault

    If selectedFromLabel <> "" And Dashboard_SnapshotLabelExists(tbl, selectedFromLabel) Then fromDefault = selectedFromLabel
    If selectedToLabel <> "" And Dashboard_SnapshotLabelExists(tbl, selectedToLabel) Then toDefault = selectedToLabel

    ws.Range("B4").value = TextCatalog_Get("DASHBOARD.SNAPSHOT.FROM", Dashboard_CurrentLanguage())
    ws.Range("F4").value = TextCatalog_Get("DASHBOARD.SNAPSHOT.TO", Dashboard_CurrentLanguage())
    ws.Range("B4,F4").font.bold = True
    ws.Range("B4,F4").font.color = RGB(96, 111, 128)

    ws.Range("C4:D4").Merge
    ws.Range("G4:H4").Merge

    With ws.Range("C4:D4")
        .value = fromDefault
        .Interior.color = RGB(255, 255, 255)
        .font.color = RGB(20, 34, 51)
        .HorizontalAlignment = xlLeft
        .validation.Delete
        If Not tbl.DataBodyRange Is Nothing Then
            .validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="=DashboardSnapshotLabels"
            .validation.ShowError = False
        End If
    End With

    With ws.Range("G4:H4")
        .value = toDefault
        .Interior.color = RGB(255, 255, 255)
        .font.color = RGB(20, 34, 51)
        .HorizontalAlignment = xlLeft
        .validation.Delete
        If Not tbl.DataBodyRange Is Nothing Then
            .validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="=DashboardSnapshotLabels"
            .validation.ShowError = False
        End If
    End With

    If recreateButton Then Dashboard_AddRefreshComparisonButton ws, ws.Range("J4").left, ws.Range("J4").top - 1, 132, 22

End Sub

'------------------------------------------------------------------------------
' FR: Actualise Render Executive Summary sans modifier les regles metier qui produisent les donnees.
' EN: Refreshes Render Executive Summary without changing the business rules that produce the data.
'------------------------------------------------------------------------------

Private Sub Dashboard_RenderExecutiveSummary( _
    ByVal ws As Worksheet, _
    ByVal tblWBS As ListObject, _
    ByVal mapWBS As Object, _
    ByVal tblCalc As ListObject, _
    ByVal mapCalc As Object, _
    ByVal scurveProjection As Object)

    Dashboard_UpdateKpiCardsInPlace ws, tblWBS, mapWBS, tblCalc, mapCalc, scurveProjection

End Sub

'------------------------------------------------------------------------------
' FR: Actualise Update Kpi Cards In Place sans modifier les regles metier qui produisent les donnees.
' EN: Refreshes Update Kpi Cards In Place without changing the business rules that produce the data.
'------------------------------------------------------------------------------

Private Sub Dashboard_UpdateKpiCardsInPlace( _
    ByVal ws As Worksheet, _
    ByVal tblWBS As ListObject, _
    ByVal mapWBS As Object, _
    ByVal tblCalc As ListObject, _
    ByVal mapCalc As Object, _
    ByVal scurveProjection As Object)

    Dim actualProgress As Double
    Dim plannedProgress As Double
    Dim progressVar As Double
    Dim baselineFinish As Variant
    Dim calcFinish As Variant
    Dim driftDays As Variant
    Dim cpCount As Long
    Dim lpCount As Long
    Dim lpRemaining As Long
    Dim contractText As String
    Dim contractColor As Long
    Dim momentumText As String
    Dim momentumColor As Long
    Dim compareText As String
    Dim momentumSubText As String
    Dim progressDeltaText As String
    Dim forecastDeltaText As String
    Dim progressColor As Long
    Dim forecastColor As Long
    Dim criticalHeroText As String
    Dim criticalSubText As String
    Dim criticalColor As Long
    Dim fromRow As Object
    Dim toRow As Object
    Dim fromLabel As String
    Dim toLabel As String
    Dim leftPos As Double
    Dim topPos As Double
    Dim widthVal As Double
    Dim cardWidth As Double
    Dim gapVal As Double

    If tblCalc Is Nothing Or tblCalc.DataBodyRange Is Nothing Then
        Dashboard_RenderExecutiveSummaryEmpty ws
        Exit Sub
    End If

    actualProgress = CDbl(scurveProjection("ActualProgress"))
    plannedProgress = CDbl(scurveProjection("PlannedProgress"))
    progressVar = actualProgress - plannedProgress

    baselineFinish = Dashboard_MaxDate(tblCalc, mapCalc, "Baseline Finish")
    calcFinish = Dashboard_MaxDate(tblCalc, mapCalc, "Calculated Finish")
    If Dashboard_HasDateValue(baselineFinish) And Dashboard_HasDateValue(calcFinish) Then driftDays = CLng(Dashboard_DateNumber(calcFinish) - Dashboard_DateNumber(baselineFinish))

    cpCount = Dashboard_CountMarker(tblCalc, mapCalc, "Critical Path")
    lpCount = Dashboard_CountMarker(tblCalc, mapCalc, "Longest Path")
    lpRemaining = Dashboard_CountMarkerRemaining(tblCalc, mapCalc, "Longest Path")

    fromLabel = CStr(ws.Range(DASH_FROM_CELL).value)
    toLabel = CStr(ws.Range(DASH_TO_CELL).value)
    Dashboard_GetComparisonRows fromLabel, toLabel, fromRow, toRow

    If Not toRow Is Nothing Then
        actualProgress = CDbl(toRow("ActualProgress"))
        plannedProgress = CDbl(toRow("PlannedProgress"))
        progressVar = CDbl(toRow("ProgressVariance"))
        calcFinish = toRow("ForecastFinish")
        baselineFinish = toRow("BaselineFinish")
        driftDays = toRow("DriftDays")
        contractText = CStr(toRow("ContractStatus"))
    Else
        contractText = Dashboard_ContractStatus(driftDays, contractColor)
    End If

    contractText = Dashboard_ContractStatus(driftDays, contractColor)
    compareText = Dashboard_ComparisonText(fromRow, toRow)
    progressDeltaText = Dashboard_BehindPlanText(plannedProgress, actualProgress, progressColor)
    forecastDeltaText = Dashboard_ContractDriftText(driftDays, forecastColor)
    criticalHeroText = Dashboard_CriticalActivitiesHero(fromRow, toRow, criticalColor, criticalSubText)
    momentumText = Dashboard_ScheduleMomentumStatus(fromRow, toRow, momentumColor, momentumSubText)

    leftPos = ws.Range("B5").left
    topPos = ws.Range("B5").top
    widthVal = ws.Range("B5:Q29").width
    gapVal = 12
    cardWidth = (widthVal - (gapVal * 3)) / 4

    ws.Range("B6:M6").Merge
    ws.Range("B6").value = TextCatalog_Get("DASHBOARD.TITLE.EXECUTIVE_SUMMARY", Dashboard_CurrentLanguage())
    ws.Range("B6").font.bold = True
    ws.Range("B6").font.size = 15
    ws.Range("B6").font.color = RGB(20, 34, 51)

    Dashboard_UpdateKpiCard ws, 1, TextCatalog_Get("DASHBOARD.KPI.PROGRESS", Dashboard_CurrentLanguage()), Format$(actualProgress, "0%"), _
        progressDeltaText, _
        leftPos + 18, ws.Range("B7").top, cardWidth - 10, 78, progressColor

    Dashboard_UpdateKpiCard ws, 2, TextCatalog_Get("DASHBOARD.KPI.FORECAST_FINISH", Dashboard_CurrentLanguage()), Dashboard_FormatDateShort(calcFinish), _
        forecastDeltaText, _
        leftPos + 18 + cardWidth + gapVal, ws.Range("B7").top, cardWidth - 10, 78, forecastColor

    Dashboard_UpdateKpiCard ws, 3, TextCatalog_Get("DASHBOARD.KPI.CRITICAL_ACTIVITIES", Dashboard_CurrentLanguage()), criticalHeroText, _
        criticalSubText, _
        leftPos + 18 + ((cardWidth + gapVal) * 2), ws.Range("B7").top, cardWidth - 10, 78, criticalColor

    Dashboard_UpdateKpiCard ws, 4, TextCatalog_Get("DASHBOARD.KPI.MOMENTUM", Dashboard_CurrentLanguage()), Dashboard_LocalMomentumStatus(momentumText), _
        momentumSubText, _
        leftPos + 18 + ((cardWidth + gapVal) * 3), ws.Range("B7").top, cardWidth - 10, 78, momentumColor

    With ws.Range("N6:Q6")
        .Merge
        .value = compareText
        .font.size = 9
        .font.color = RGB(96, 111, 128)
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlCenter
    End With

End Sub

'------------------------------------------------------------------------------
' FR: Actualise Update S Curve Chart In Place sans modifier les regles metier qui produisent les donnees.
' EN: Refreshes Update S Curve Chart In Place without changing the business rules that produce the data.
' FR - Effet de bord : efface uniquement les donnees ou objets cibles du contrat.
' EN - Side effect: clears only data or objects targeted by the contract.
'------------------------------------------------------------------------------

Private Sub Dashboard_UpdateSCurveChartInPlace(ByVal ws As Worksheet, ByVal scurveProjection As Object)

    Dim chartObj As ChartObject
    Dim ch As chart
    Dim leftPos As Double
    Dim topPos As Double
    Dim widthVal As Double
    Dim heightVal As Double
    Dim xRange As Range
    Dim baselineRange As Range
    Dim actualRange As Range
    Dim forecastRange As Range

    If scurveProjection Is Nothing Then Exit Sub
    If Not scurveProjection.exists("DateRange") Then Exit Sub
    Set xRange = scurveProjection("DateRange")
    If scurveProjection.exists("BaselineRange") Then Set baselineRange = scurveProjection("BaselineRange")
    If scurveProjection.exists("ActualRange") Then Set actualRange = scurveProjection("ActualRange")
    If scurveProjection.exists("ForecastRange") Then Set forecastRange = scurveProjection("ForecastRange")

    On Error Resume Next
    Set chartObj = ws.ChartObjects(DASH_CHART_SCURVE)
    On Error GoTo 0
    If chartObj Is Nothing Then
        Dashboard_RenderSCurveChart ws, scurveProjection
        Exit Sub
    End If

    leftPos = ws.Range("B13").left + 18
    topPos = ws.Range("B13").top
    widthVal = (ws.Range("I13").left + ws.Range("I13").width) - leftPos
    heightVal = ws.Range("B13:I29").height
    chartObj.left = leftPos + 16
    chartObj.top = topPos + 38
    chartObj.width = widthVal - 24
    chartObj.height = heightVal - 56

    Set ch = chartObj.chart
    Do While ch.SeriesCollection.count > 0
        ch.SeriesCollection(1).Delete
    Loop
    ch.ChartType = xlLine
    ch.HasTitle = False
    ch.HasLegend = True
    ch.Legend.position = xlLegendPositionBottom
    Dashboard_AddLineSeries ch, TextCatalog_Get("DASHBOARD.SERIES.BASELINE", Dashboard_CurrentLanguage()), xRange, baselineRange, RGB(150, 150, 150), 1.5, False
    Dashboard_AddLineSeries ch, TextCatalog_Get("DASHBOARD.SERIES.ACTUAL", Dashboard_CurrentLanguage()), xRange, actualRange, RGB(0, 145, 112), 2.5, False
    Dashboard_AddLineSeries ch, TextCatalog_Get("DASHBOARD.SERIES.FORECAST", Dashboard_CurrentLanguage()), xRange, forecastRange, RGB(43, 106, 176), 2.5, True

    On Error Resume Next
    ch.Axes(xlValue).TickLabels.numberFormat = "0%"
    ch.Axes(xlCategory).TickLabels.numberFormat = Dashboard_ChartDateNumberFormat(False)
    ch.Axes(xlCategory).TickLabels.orientation = 45
    ch.Axes(xlValue).MajorGridlines.Format.Line.ForeColor.RGB = RGB(225, 231, 238)
    ch.ChartArea.Format.Line.Visible = msoFalse
    ch.PlotArea.Format.Line.Visible = msoFalse
    ch.ChartArea.Format.fill.Visible = msoFalse
    ch.PlotArea.Format.fill.Visible = msoFalse
    On Error GoTo 0
    SCurve_DrawDashboardTodayLine ch, ws, DASH_SCURVE_TODAY_LINE

End Sub
'------------------------------------------------------------------------------
' FR: Actualise Render S Curve Chart sans modifier les regles metier qui produisent les donnees.
' EN: Refreshes Render S Curve Chart without changing the business rules that produce the data.
'------------------------------------------------------------------------------

Private Sub Dashboard_RenderSCurveChart(ByVal ws As Worksheet, ByVal scurveProjection As Object)

    Dim chartObj As ChartObject
    Dim ch As chart
    Dim xRange As Range
    Dim baselineRange As Range
    Dim actualRange As Range
    Dim forecastRange As Range
    Dim leftPos As Double
    Dim topPos As Double
    Dim widthVal As Double
    Dim heightVal As Double

    If scurveProjection Is Nothing Then Exit Sub
    If Not scurveProjection.exists("DateRange") Then Exit Sub
    Set xRange = scurveProjection("DateRange")
    If scurveProjection.exists("BaselineRange") Then Set baselineRange = scurveProjection("BaselineRange")
    If scurveProjection.exists("ActualRange") Then Set actualRange = scurveProjection("ActualRange")
    If scurveProjection.exists("ForecastRange") Then Set forecastRange = scurveProjection("ForecastRange")
    leftPos = ws.Range("B13").left + 18
    topPos = ws.Range("B13").top
    widthVal = (ws.Range("I13").left + ws.Range("I13").width) - leftPos
    heightVal = ws.Range("B13:I29").height
    Dashboard_AddPanel ws, leftPos, topPos, widthVal, heightVal, RGB(255, 255, 255)
    Dashboard_AddCenteredSectionTitle ws, TextCatalog_Get("DASHBOARD.SECTION.SCURVE", Dashboard_CurrentLanguage()), leftPos + 22, topPos + 10, widthVal - 44
    Set chartObj = ws.ChartObjects.Add(leftPos + 16, topPos + 38, widthVal - 24, heightVal - 56)
    chartObj.name = DASH_CHART_SCURVE
    Set ch = chartObj.chart
    ch.ChartType = xlLine
    ch.HasTitle = False
    ch.HasLegend = True
    ch.Legend.position = xlLegendPositionBottom
    Dashboard_AddLineSeries ch, TextCatalog_Get("DASHBOARD.SERIES.BASELINE", Dashboard_CurrentLanguage()), xRange, baselineRange, RGB(150, 150, 150), 1.5, False
    Dashboard_AddLineSeries ch, TextCatalog_Get("DASHBOARD.SERIES.ACTUAL", Dashboard_CurrentLanguage()), xRange, actualRange, RGB(0, 145, 112), 2.5, False
    Dashboard_AddLineSeries ch, TextCatalog_Get("DASHBOARD.SERIES.FORECAST", Dashboard_CurrentLanguage()), xRange, forecastRange, RGB(43, 106, 176), 2.5, True
    On Error Resume Next
    ch.Axes(xlValue).TickLabels.numberFormat = "0%"
    ch.Axes(xlCategory).TickLabels.numberFormat = Dashboard_ChartDateNumberFormat(False)
    ch.Axes(xlCategory).TickLabels.orientation = 45
    ch.Axes(xlValue).MajorGridlines.Format.Line.ForeColor.RGB = RGB(225, 231, 238)
    ch.ChartArea.Format.Line.Visible = msoFalse
    ch.PlotArea.Format.Line.Visible = msoFalse
    ch.ChartArea.Format.fill.Visible = msoFalse
    ch.PlotArea.Format.fill.Visible = msoFalse
    On Error GoTo 0
    SCurve_DrawDashboardTodayLine ch, ws, DASH_SCURVE_TODAY_LINE

End Sub

'------------------------------------------------------------------------------
' FR: Ajoute la reference Line Series a la structure cible fournie par l'appelant.
' EN: Adds the Line Series reference to the target structure supplied by the caller.
'------------------------------------------------------------------------------

Private Sub Dashboard_AddLineSeries( _
    ByVal ch As chart, _
    ByVal seriesName As String, _
    ByVal xRange As Range, _
    ByVal yRange As Range, _
    ByVal rgbColor As Long, _
    ByVal weightVal As Double, _
    ByVal dashed As Boolean)

    Dim s As series
    If xRange Is Nothing Or yRange Is Nothing Then Exit Sub
    Set s = ch.SeriesCollection.NewSeries
    With s
        .name = seriesName
        .XValues = xRange
        .values = yRange
        .ChartType = xlLine
        .Format.Line.ForeColor.RGB = rgbColor
        .Format.Line.weight = weightVal
        If dashed Then .Format.Line.DashStyle = msoLineDash
        .MarkerStyle = xlMarkerStyleNone
    End With

End Sub

'------------------------------------------------------------------------------
' FR: Actualise Update Planning Overview Content sans modifier les regles metier qui produisent les donnees.
' EN: Refreshes Update Planning Overview Content without changing the business rules that produce the data.
'------------------------------------------------------------------------------

Private Sub Dashboard_UpdatePlanningOverviewContent(ByVal ws As Worksheet, ByVal tblWBS As ListObject, ByVal mapWBS As Object, ByVal tblCalc As ListObject, ByVal mapCalc As Object)

    Dim leftPos As Double
    Dim topPos As Double
    Dim widthVal As Double
    Dim heightVal As Double

    leftPos = ws.Range("J13").left + 8
    topPos = ws.Range("J13").top
    widthVal = ws.Range("J13:Q29").width - 20
    heightVal = ws.Range("J13:Q29").height

    Dashboard_DeleteShapesInRect ws, leftPos + 1, topPos + 36, widthVal - 2, heightVal - 38
    Dashboard_RenderPlanningOverview ws, tblWBS, mapWBS, tblCalc, mapCalc, False

End Sub
'------------------------------------------------------------------------------
' FR: Actualise Render Planning Overview sans modifier les regles metier qui produisent les donnees.
' EN: Refreshes Render Planning Overview without changing the business rules that produce the data.
'------------------------------------------------------------------------------

Private Sub Dashboard_RenderPlanningOverview(ByVal ws As Worksheet, ByVal tblWBS As ListObject, ByVal mapWBS As Object, ByVal tblCalc As ListObject, ByVal mapCalc As Object, Optional ByVal renderShell As Boolean = True)

    Dim leftPos As Double
    Dim topPos As Double
    Dim widthVal As Double
    Dim heightVal As Double
    Dim arr As Variant
    Dim r As Long
    Dim rowsToShow As Collection
    Dim rowRef As Variant
    Dim rowIndex As Long
    Dim rowCount As Long
    Dim shown As Long
    Dim rowH As Double
    Dim rowCenter As Double
    Dim shapeCenterY As Double
    Dim separatorTopY As Double
    Dim separatorBottomY As Double
    Dim firstRowCenter As Double
    Dim labelTop As Double
    Dim labelText As String
    Dim rowStart As Variant
    Dim rowFinish As Variant
    Dim baseStart As Variant
    Dim baseFinish As Variant
    Dim x1 As Double
    Dim x2 As Double
    Dim todayX As Double
    Dim axisLeft As Double
    Dim axisWidth As Double
    Dim axisRight As Double
    Dim progressSummary As Double
    Dim progressById As Object
    Dim summaryProgressByWbs As Object
    Dim visibleParents As Object
    Dim summaryMetrics As Variant
    Dim progressColor As Variant
    Dim axisMin As Double
    Dim axisMax As Double
    Dim axisMargin As Double
    Dim usefulLeft As Double
    Dim usefulTop As Double
    Dim usefulWidth As Double
    Dim usefulHeight As Double
    Dim labelWidth As Double
    Dim plotTop As Double
    Dim plotBottom As Double
    Dim axisY As Double
    Dim fontSize As Double
    Dim barHeight As Double
    Dim baselineHeight As Double
    Dim wbsVal As String
    Dim taskName As String
    Dim isMilestone As Boolean

    If tblCalc Is Nothing Then Exit Sub
    If tblCalc.DataBodyRange Is Nothing Then Exit Sub
    If Not mapCalc.exists("IsSummary") Then Exit Sub
    If Not mapCalc.exists("Calculated Start") Then Exit Sub
    If Not mapCalc.exists("Calculated Finish") Then Exit Sub

    leftPos = ws.Range("J13").left + 8
    topPos = ws.Range("J13").top
    widthVal = ws.Range("J13:Q29").width - 20
    heightVal = ws.Range("J13:Q29").height

    If renderShell Then
        Dashboard_AddPanel ws, leftPos, topPos, widthVal, heightVal, RGB(255, 255, 255)
        Dashboard_AddCenteredSectionTitle ws, TextCatalog_Get("DASHBOARD.SECTION.OVERVIEW", Dashboard_CurrentLanguage()), leftPos + 22, topPos + 10, widthVal - 44
    End If

    arr = tblCalc.DataBodyRange.value
    Set progressById = Dashboard_BuildProgressById(tblWBS, mapWBS)

    Set rowsToShow = Dashboard_PlanningOverviewRows(arr, mapCalc, axisMin, axisMax)
    If rowsToShow.count = 0 Or axisMin <= 0 Or axisMax < axisMin Then
        Dashboard_WriteEmptyState ws, leftPos + 22, topPos + 70, widthVal - 44, TextCatalog_Get("DASHBOARD.EMPTY.NO_SUMMARY", Dashboard_CurrentLanguage())
        Exit Sub
    End If

    Set visibleParents = CreateObject("Scripting.Dictionary")
    For Each rowRef In rowsToShow
        rowIndex = CLng(rowRef)
        If Dashboard_IsSummaryArrayRow(arr, mapCalc, rowIndex) Then
            visibleParents(NormalizeWBS(CStr(Dashboard_ArrayVal(arr, mapCalc, rowIndex, "WBS")))) = True
        End If
    Next rowRef
    Set summaryProgressByWbs = Dashboard_BuildSummaryProgress(arr, mapCalc, progressById, visibleParents)

    axisMargin = Dashboard_PlanningAxisMargin(axisMin, axisMax)
    axisMin = axisMin - axisMargin
    axisMax = axisMax + axisMargin
    If axisMax <= axisMin Then axisMax = axisMin + 1

    usefulLeft = leftPos + 22
    usefulTop = topPos + 42
    usefulWidth = widthVal - 44
    usefulHeight = heightVal - 62
    labelWidth = WorksheetFunction.Min(usefulWidth * 0.25, 132)
    axisLeft = usefulLeft + labelWidth + 14
    axisWidth = usefulWidth - labelWidth - 14
    axisRight = axisLeft + axisWidth
    axisY = usefulTop + 18
    plotTop = usefulTop + 46
    plotBottom = topPos + heightVal - 36

    Dashboard_RenderTimeAxis ws, axisMin, axisMax, axisLeft, axisY, axisWidth

    rowCount = rowsToShow.count
    If rowCount <= 0 Then Exit Sub
    rowH = (plotBottom - plotTop) / rowCount
    If rowH > 22 Then rowH = 22
    If rowH < 6 Then rowH = 6

    fontSize = rowH - 2
    If fontSize > 9 Then fontSize = 9
    If fontSize < 8 Then fontSize = 8

    barHeight = rowH * 0.42
    If barHeight > 8 Then barHeight = 8
    If barHeight < 4 Then barHeight = 4
    baselineHeight = 3

    firstRowCenter = plotTop + ((plotBottom - plotTop) - (rowH * rowCount)) / 2 + (rowH / 2)
    If firstRowCenter < plotTop + (rowH / 2) Then firstRowCenter = plotTop + (rowH / 2)

    shown = 0
    For Each rowRef In rowsToShow
        rowIndex = CLng(rowRef)
        rowStart = GetCellValue(Dashboard_ArrayVal(arr, mapCalc, rowIndex, "Calculated Start"))
        rowFinish = GetCellValue(Dashboard_ArrayVal(arr, mapCalc, rowIndex, "Calculated Finish"))
        If Dashboard_HasDateValue(rowStart) And Dashboard_HasDateValue(rowFinish) Then
            shown = shown + 1
            rowCenter = firstRowCenter + ((shown - 1) * rowH)
            shapeCenterY = rowCenter + 1.5
            separatorTopY = shapeCenterY - (rowH / 2)
            separatorBottomY = shapeCenterY + (rowH / 2)

            If shown = 1 Then Dashboard_AddLine ws, usefulLeft, separatorTopY, axisRight, separatorTopY, RGB(224, 230, 237), 0.5
            Dashboard_AddLine ws, usefulLeft, separatorBottomY, axisRight, separatorBottomY, RGB(224, 230, 237), 0.5

            wbsVal = CStr(Dashboard_ArrayVal(arr, mapCalc, rowIndex, "WBS"))
            taskName = CStr(Dashboard_ArrayVal(arr, mapCalc, rowIndex, "Task Name"))
            labelText = Dashboard_TruncateText(wbsVal & "  " & taskName, 26)
            labelTop = shapeCenterY - ((fontSize + 4) / 2)
            Dashboard_AddMiniGanttLabel ws, labelText, usefulLeft, labelTop, labelWidth, fontSize

            progressSummary = Dashboard_RowProgress(arr, mapCalc, progressById, rowIndex, _
                CStr(GetCellValue(Dashboard_ArrayVal(arr, mapCalc, rowIndex, "ID"))))
            progressColor = Empty
            If Dashboard_IsSummaryArrayRow(arr, mapCalc, rowIndex) Then
                If summaryProgressByWbs.exists(NormalizeWBS(wbsVal)) Then
                    summaryMetrics = summaryProgressByWbs(NormalizeWBS(wbsVal))
                    progressSummary = summaryMetrics(1) / summaryMetrics(0)
                    progressColor = summaryMetrics(2)
                End If
            End If
            isMilestone = (Dashboard_DateNumber(rowFinish) <= Dashboard_DateNumber(rowStart))

            baseStart = GetCellValue(Dashboard_ArrayVal(arr, mapCalc, rowIndex, "Baseline Start"))
            baseFinish = GetCellValue(Dashboard_ArrayVal(arr, mapCalc, rowIndex, "Baseline Finish"))
            If Dashboard_HasDateValue(baseStart) And Dashboard_HasDateValue(baseFinish) Then
                x1 = Dashboard_DateToX(Dashboard_DateNumber(baseStart), axisMin, axisMax, axisLeft, axisWidth)
                x2 = Dashboard_DateToX(Dashboard_DateNumber(baseFinish), axisMin, axisMax, axisLeft, axisWidth)
                If (x2 - x1) >= 6 Then
                    Dashboard_AddBar ws, x1, shapeCenterY - (baselineHeight / 2), x2 - x1, baselineHeight, RGB(189, 198, 208)
                End If
            End If

            x1 = Dashboard_DateToX(Dashboard_DateNumber(rowStart), axisMin, axisMax, axisLeft, axisWidth)
            x2 = Dashboard_DateToX(Dashboard_DateNumber(rowFinish), axisMin, axisMax, axisLeft, axisWidth)
            If isMilestone Then
                Dashboard_AddMilestoneDiamond ws, x1, shapeCenterY, WorksheetFunction.Max(7, barHeight + 2), Dashboard_MiniGanttCurrentColor(progressSummary)
            ElseIf (x2 - x1) >= 6 Then
                Dashboard_AddProgressBar ws, x1, shapeCenterY - (barHeight / 2), x2 - x1, barHeight, progressSummary, rowStart, rowFinish, progressColor
            End If
        End If
    Next rowRef

    If shown = 0 Then
        Dashboard_WriteEmptyState ws, leftPos + 22, topPos + 70, widthVal - 44, TextCatalog_Get("DASHBOARD.EMPTY.NO_SUMMARY", Dashboard_CurrentLanguage())
    End If

    If CDbl(Date) >= axisMin And CDbl(Date) <= axisMax Then
        todayX = Dashboard_DateToX(CDbl(Date), axisMin, axisMax, axisLeft, axisWidth)
        Dashboard_AddLine ws, todayX, plotTop, todayX, plotBottom, RGB(0, 145, 112), 1.25
    End If

End Sub

'------------------------------------------------------------------------------
' FR: Actualise Update Hot Spots Content sans modifier les regles metier qui produisent les donnees.
' EN: Refreshes Update Hot Spots Content without changing the business rules that produce the data.
'------------------------------------------------------------------------------

Private Sub Dashboard_UpdateHotSpotsContent(ByVal ws As Worksheet, ByVal tblWBS As ListObject, ByVal mapWBS As Object, ByVal tblCalc As ListObject, ByVal mapCalc As Object)

    Dim leftPos As Double
    Dim topPos As Double
    Dim widthVal As Double
    Dim heightVal As Double

    leftPos = ws.Range("B32").left
    topPos = ws.Range("B32").top
    widthVal = ws.Range("B32:Q56").width
    heightVal = ws.Range("B32:Q56").height

    Dashboard_DeleteShapesInRect ws, leftPos + 1, topPos + 40, widthVal - 2, heightVal - 42
    Dashboard_RenderHotSpots ws, tblWBS, mapWBS, tblCalc, mapCalc, False

End Sub
'------------------------------------------------------------------------------
' FR: Actualise Render Hot Spots sans modifier les regles metier qui produisent les donnees.
' EN: Refreshes Render Hot Spots without changing the business rules that produce the data.
'------------------------------------------------------------------------------

Private Sub Dashboard_RenderHotSpots(ByVal ws As Worksheet, ByVal tblWBS As ListObject, ByVal mapWBS As Object, ByVal tblCalc As ListObject, ByVal mapCalc As Object, Optional ByVal renderShell As Boolean = True)

    Dim leftPos As Double
    Dim topPos As Double
    Dim widthVal As Double
    Dim heightVal As Double
    Dim cardTop As Double
    Dim cardHeight As Double
    Dim gapVal As Double
    Dim rightW As Double
    Dim largeW As Double
    Dim smallH As Double

    leftPos = ws.Range("B32").left
    topPos = ws.Range("B32").top
    widthVal = ws.Range("B32:Q56").width
    heightVal = ws.Range("B32:Q56").height

    If renderShell Then Dashboard_AddSectionTitle ws, TextCatalog_Get("DASHBOARD.SECTION.HOT_SPOTS", Dashboard_CurrentLanguage()), leftPos + 22, topPos + 10, widthVal - 36

    cardTop = topPos + 48
    cardHeight = heightVal - 64
    gapVal = 14
    rightW = 240
    largeW = (widthVal - 44 - rightW - (gapVal * 2)) / 2
    smallH = (cardHeight - gapVal) / 2

    Dashboard_RenderTopDelays ws, tblWBS, mapWBS, leftPos + 22, cardTop, largeW, cardHeight
    Dashboard_RenderDeadlineHealth ws, tblCalc, mapCalc, leftPos + 22 + largeW + gapVal, cardTop, largeW, cardHeight
    Dashboard_RenderNextMilestone ws, tblCalc, mapCalc, leftPos + 22 + (largeW * 2) + (gapVal * 2), cardTop, rightW, smallH
    Dashboard_RenderNextCriticalActivity ws, tblCalc, mapCalc, leftPos + 22 + (largeW * 2) + (gapVal * 2), cardTop + smallH + gapVal, rightW, smallH

End Sub

'------------------------------------------------------------------------------
' FR: Actualise Render Top Delays sans modifier les regles metier qui produisent les donnees.
' EN: Refreshes Render Top Delays without changing the business rules that produce the data.
'------------------------------------------------------------------------------

Private Sub Dashboard_RenderTopDelays(ByVal ws As Worksheet, ByVal tblWBS As ListObject, ByVal mapWBS As Object, ByVal X As Double, ByVal Y As Double, ByVal w As Double, ByVal h As Double)

    Const TOP_DELAY_COUNT As Long = 4

    Dim topIds(1 To 5) As String
    Dim topWbs(1 To 5) As String
    Dim topNames(1 To 5) As String
    Dim topVals(1 To 5) As Double
    Dim arr As Variant
    Dim r As Long
    Dim v As Variant
    Dim i As Long
    Dim lineY As Double
    Dim accentColor As Long

    On Error GoTo RenderFailed

    If tblWBS Is Nothing Then GoTo NoData
    If mapWBS Is Nothing Then GoTo NoData
    If tblWBS.DataBodyRange Is Nothing Then GoTo NoData
    If Not mapWBS.exists(VTS_COL_FINISH_VARIANCE) Then GoTo NoData

    arr = tblWBS.DataBodyRange.value
    For r = 1 To UBound(arr, 1)
        v = Dashboard_ArrayVal(arr, mapWBS, r, "Finish Variance")
        If IsNumeric(v) Then
            If CDbl(v) > 0 Then
                Dashboard_InsertTopDelay topIds, topWbs, topNames, topVals, _
                    CStr(Dashboard_ArrayVal(arr, mapWBS, r, "ID")), _
                    CStr(Dashboard_ArrayVal(arr, mapWBS, r, "WBS")), _
                    CStr(Dashboard_ArrayVal(arr, mapWBS, r, "Task Name")), _
                    CDbl(v)
            End If
        End If
    Next r

    If topIds(1) <> "" Then
        accentColor = RGB(192, 80, 77)
    Else
        accentColor = RGB(0, 145, 112)
    End If

    Dashboard_AddHotSpotCard ws, X, Y, w, h, TextCatalog_Get("DASHBOARD.CARD.TOP_DELAYS", Dashboard_CurrentLanguage()), TextCatalog_Get("DASHBOARD.HELP.DELAYS", Dashboard_CurrentLanguage()), accentColor

    For i = 1 To TOP_DELAY_COUNT
        lineY = Y + 66 + ((i - 1) * 68)
        If topIds(i) <> "" Then
            Dashboard_AddHotSpotRankingItem ws, X + 18, lineY, w - 36, 54, _
                CStr(i), _
                topWbs(i), _
                Dashboard_TruncateText(topNames(i), 30), _
                "+" & CStr(CLng(topVals(i))) & Dashboard_DurationSuffix(), _
                RGB(192, 80, 77)
        End If
    Next i

    If topIds(1) = "" Then
        Dashboard_AddHotSpotText ws, TextCatalog_Get("DASHBOARD.EMPTY.NO_DELAYS", Dashboard_CurrentLanguage()), X + 18, Y + 90, w - 36, 18, RGB(96, 111, 128), 8, False, xlHAlignLeft
    End If

    Exit Sub

NoData:
    Dashboard_AddHotSpotCard ws, X, Y, w, h, TextCatalog_Get("DASHBOARD.CARD.TOP_DELAYS", Dashboard_CurrentLanguage()), TextCatalog_Get("DASHBOARD.HELP.DELAYS", Dashboard_CurrentLanguage()), RGB(160, 170, 181)
    Dashboard_AddHotSpotText ws, TextCatalog_Get("DASHBOARD.EMPTY.NO_DELAY_DATA", Dashboard_CurrentLanguage()), X + 18, Y + 90, w - 36, 18, RGB(96, 111, 128), 8, False, xlHAlignLeft
    Exit Sub

RenderFailed:
    Dashboard_WriteHotSpotRuntimeError ws, X + 18, Y + 90, w - 36, TextCatalog_Get("DASHBOARD.CARD.TOP_DELAYS", Dashboard_CurrentLanguage())

End Sub

'------------------------------------------------------------------------------
' FR: Actualise Render Deadline Health sans modifier les regles metier qui produisent les donnees.
' EN: Refreshes Render Deadline Health without changing the business rules that produce the data.
'------------------------------------------------------------------------------

Private Sub Dashboard_RenderDeadlineHealth(ByVal ws As Worksheet, ByVal tblCalc As ListObject, ByVal mapCalc As Object, ByVal X As Double, ByVal Y As Double, ByVal w As Double, ByVal h As Double)

    Dim arr As Variant
    Dim r As Long
    Dim deadlineFloat As Variant
    Dim floatVal As Double
    Dim overdueCount As Long
    Dim onTrackCount As Long
    Dim activeDeadlineCount As Long
    Dim worstFound As Boolean
    Dim riskFound As Boolean
    Dim worstFloat As Double
    Dim riskFloat As Double
    Dim worstWbs As String
    Dim worstName As String
    Dim riskWbs As String
    Dim riskName As String
    Dim accentColor As Long
    Dim riskColor As Long
    Dim riskFill As Long
    Dim closestY As Double

    On Error GoTo RenderFailed

    If tblCalc Is Nothing Then GoTo NoData
    If mapCalc Is Nothing Then GoTo NoData
    If tblCalc.DataBodyRange Is Nothing Then GoTo NoData
    If Not mapCalc.exists("Deadline") Then GoTo NoData
    If Not mapCalc.exists("Deadline Float") Then GoTo NoData

    arr = tblCalc.DataBodyRange.value
    For r = 1 To UBound(arr, 1)
        If Dashboard_IsActiveDashboardTask(arr, mapCalc, r) Then
            If HasValue(Dashboard_ArrayVal(arr, mapCalc, r, "Deadline")) Then
                deadlineFloat = Dashboard_ArrayVal(arr, mapCalc, r, "Deadline Float")
                If IsNumeric(deadlineFloat) Then
                    activeDeadlineCount = activeDeadlineCount + 1
                    floatVal = CDbl(deadlineFloat)
                    If floatVal < 0# Then
                        overdueCount = overdueCount + 1
                        If Not worstFound Or floatVal < worstFloat Then
                            worstFound = True
                            worstFloat = floatVal
                            worstWbs = CStr(Dashboard_ArrayVal(arr, mapCalc, r, "WBS"))
                            worstName = CStr(Dashboard_ArrayVal(arr, mapCalc, r, "Task Name"))
                        End If
                    Else
                        onTrackCount = onTrackCount + 1
                        If Not riskFound Or floatVal < riskFloat Then
                            riskFound = True
                            riskFloat = floatVal
                            riskWbs = CStr(Dashboard_ArrayVal(arr, mapCalc, r, "WBS"))
                            riskName = CStr(Dashboard_ArrayVal(arr, mapCalc, r, "Task Name"))
                        End If
                    End If
                End If
            End If
        End If
    Next r

    If activeDeadlineCount = 0 Then
        accentColor = RGB(160, 170, 181)
    ElseIf overdueCount > 0 Then
        accentColor = RGB(192, 80, 77)
    Else
        accentColor = RGB(0, 145, 112)
    End If

    Dashboard_AddHotSpotCard ws, X, Y, w, h, TextCatalog_Get("DASHBOARD.CARD.DEADLINES", Dashboard_CurrentLanguage()), TextCatalog_Get("DASHBOARD.HELP.COMMITMENTS", Dashboard_CurrentLanguage()), accentColor
    Dashboard_AddHotSpotKpiBlock ws, X + 18, Y + 62, (w - 48) / 2, 64, TextCatalog_Get("DASHBOARD.STATUS.OVERDUE", Dashboard_CurrentLanguage()), CStr(overdueCount), IIf(overdueCount > 0, RGB(192, 80, 77), RGB(0, 145, 112)), IIf(overdueCount > 0, RGB(252, 235, 232), RGB(229, 246, 239))
    Dashboard_AddHotSpotKpiBlock ws, X + 30 + ((w - 48) / 2), Y + 62, (w - 48) / 2, 64, TextCatalog_Get("DASHBOARD.STATUS.ON_TRACK", Dashboard_CurrentLanguage()), CStr(onTrackCount), RGB(0, 145, 112), RGB(229, 246, 239)

    closestY = Y + 152
    If worstFound Then
        Dashboard_AddHotSpotInsightBlock ws, X + 18, closestY, w - 36, 72, TextCatalog_Get("DASHBOARD.LABEL.WORST", Dashboard_CurrentLanguage()), worstWbs, Dashboard_TruncateText(worstName, 30), CStr(CLng(worstFloat)) & Dashboard_DurationSuffix(), RGB(192, 80, 77), RGB(252, 235, 232)
        closestY = Y + 242
    End If

    If riskFound Then
        If riskFloat <= 5# Then
            riskColor = RGB(238, 156, 68)
            riskFill = RGB(255, 244, 226)
        Else
            riskColor = RGB(0, 145, 112)
            riskFill = RGB(229, 246, 239)
        End If
        Dashboard_AddHotSpotInsightBlock ws, X + 18, closestY, w - 36, 72, TextCatalog_Get("DASHBOARD.LABEL.NEAREST_RISK", Dashboard_CurrentLanguage()), riskWbs, Dashboard_TruncateText(riskName, 30), "+" & CStr(CLng(riskFloat)) & Dashboard_DurationSuffix(), riskColor, riskFill
    ElseIf activeDeadlineCount = 0 Then
        Dashboard_AddHotSpotText ws, TextCatalog_Get("DASHBOARD.EMPTY.NO_ACTIVE_DEADLINE", Dashboard_CurrentLanguage()), X + 18, closestY + 10, w - 36, 18, RGB(96, 111, 128), 8, False, xlHAlignLeft
    End If

    Exit Sub

NoData:
    Dashboard_AddHotSpotCard ws, X, Y, w, h, TextCatalog_Get("DASHBOARD.CARD.DEADLINES", Dashboard_CurrentLanguage()), TextCatalog_Get("DASHBOARD.HELP.COMMITMENTS", Dashboard_CurrentLanguage()), RGB(160, 170, 181)
    Dashboard_AddHotSpotText ws, TextCatalog_Get("DASHBOARD.EMPTY.NO_DEADLINE_DATA", Dashboard_CurrentLanguage()), X + 18, Y + 90, w - 36, 18, RGB(96, 111, 128), 8, False, xlHAlignLeft
    Exit Sub

RenderFailed:
    Dashboard_WriteHotSpotRuntimeError ws, X + 18, Y + 90, w - 36, TextCatalog_Get("DASHBOARD.CARD.DEADLINES", Dashboard_CurrentLanguage())

End Sub

'------------------------------------------------------------------------------
' FR: Actualise Render Next Milestone sans modifier les regles metier qui produisent les donnees.
' EN: Refreshes Render Next Milestone without changing the business rules that produce the data.
'------------------------------------------------------------------------------

Private Sub Dashboard_RenderNextMilestone(ByVal ws As Worksheet, ByVal tblCalc As ListObject, ByVal mapCalc As Object, ByVal X As Double, ByVal Y As Double, ByVal w As Double, ByVal h As Double)

    Dim arr As Variant
    Dim r As Long
    Dim finishVal As Variant
    Dim finishSerial As Double
    Dim todaySerial As Double
    Dim bestFutureFound As Boolean
    Dim bestPastFound As Boolean
    Dim bestFutureDate As Double
    Dim bestPastDate As Double
    Dim bestWbs As String
    Dim bestName As String
    Dim bestDate As Double
    Dim pastWbs As String
    Dim pastName As String
    Dim accentColor As Long
    Dim dateColor As Long

    On Error GoTo RenderFailed

    If tblCalc Is Nothing Then GoTo NoData
    If mapCalc Is Nothing Then GoTo NoData
    If tblCalc.DataBodyRange Is Nothing Then GoTo NoData
    If Not mapCalc.exists("Task Type") Then GoTo NoData
    If Not mapCalc.exists("Calculated Finish") Then GoTo NoData

    arr = tblCalc.DataBodyRange.value
    todaySerial = CDbl(Date)

    For r = 1 To UBound(arr, 1)
        If Dashboard_IsMilestoneRow(arr, mapCalc, r) And Dashboard_IsActiveDashboardTask(arr, mapCalc, r) Then
            finishVal = GetCellValue(Dashboard_ArrayVal(arr, mapCalc, r, "Calculated Finish"))
            If Dashboard_HasDateValue(finishVal) Then
                finishSerial = Dashboard_DateNumber(finishVal)
                If finishSerial >= todaySerial Then
                    If Not bestFutureFound Or finishSerial < bestFutureDate Then
                        bestFutureFound = True
                        bestFutureDate = finishSerial
                        bestWbs = CStr(Dashboard_ArrayVal(arr, mapCalc, r, "WBS"))
                        bestName = CStr(Dashboard_ArrayVal(arr, mapCalc, r, "Task Name"))
                    End If
                ElseIf Not bestPastFound Or finishSerial > bestPastDate Then
                    bestPastFound = True
                    bestPastDate = finishSerial
                    pastWbs = CStr(Dashboard_ArrayVal(arr, mapCalc, r, "WBS"))
                    pastName = CStr(Dashboard_ArrayVal(arr, mapCalc, r, "Task Name"))
                End If
            End If
        End If
    Next r

    If bestFutureFound Then
        bestDate = bestFutureDate
        accentColor = RGB(68, 114, 196)
        dateColor = RGB(68, 114, 196)
    ElseIf bestPastFound Then
        bestDate = bestPastDate
        bestWbs = pastWbs
        bestName = pastName
        accentColor = RGB(192, 80, 77)
        dateColor = RGB(192, 80, 77)
    Else
        accentColor = RGB(160, 170, 181)
        dateColor = RGB(96, 111, 128)
    End If

    Dashboard_AddHotSpotCard ws, X, Y, w, h, TextCatalog_Get("DASHBOARD.CARD.NEXT_MILESTONE", Dashboard_CurrentLanguage()), "", accentColor

    If bestFutureFound Or bestPastFound Then
        Dashboard_AddHotSpotHero ws, Dashboard_TruncateText(bestName, 24), X + 16, Y + 58, w - 32, 28, RGB(20, 34, 51)
        Dashboard_AddHotSpotText ws, bestWbs, X + 16, Y + 86, w - 32, 14, RGB(96, 111, 128), 8, False, xlHAlignLeft
        Dashboard_AddHotSpotText ws, Dashboard_FormatDate(bestDate, True), X + 16, Y + 112, w - 32, 16, RGB(42, 52, 65), 9, False, xlHAlignLeft
        Dashboard_AddHotSpotText ws, Dashboard_DaysRemainingText(CLng(bestDate - todaySerial)), X + 16, Y + 136, w - 32, 16, dateColor, 9, True, xlHAlignLeft
    Else
        Dashboard_AddHotSpotText ws, TextCatalog_Get("DASHBOARD.EMPTY.NO_ACTIVE_MILESTONE", Dashboard_CurrentLanguage()), X + 16, Y + 76, w - 32, 18, RGB(96, 111, 128), 8, False, xlHAlignLeft
    End If

    Exit Sub

NoData:
    Dashboard_AddHotSpotCard ws, X, Y, w, h, TextCatalog_Get("DASHBOARD.CARD.NEXT_MILESTONE", Dashboard_CurrentLanguage()), "", RGB(160, 170, 181)
    Dashboard_AddHotSpotText ws, TextCatalog_Get("DASHBOARD.EMPTY.NO_MILESTONE_DATA", Dashboard_CurrentLanguage()), X + 16, Y + 76, w - 32, 18, RGB(96, 111, 128), 8, False, xlHAlignLeft
    Exit Sub

RenderFailed:
    Dashboard_WriteHotSpotRuntimeError ws, X + 16, Y + 76, w - 32, TextCatalog_Get("DASHBOARD.CARD.NEXT_MILESTONE", Dashboard_CurrentLanguage())

End Sub

'------------------------------------------------------------------------------
' FR: Actualise Render Next Critical Activity sans modifier les regles metier qui produisent les donnees.
' EN: Refreshes Render Next Critical Activity without changing the business rules that produce the data.
'------------------------------------------------------------------------------

Private Sub Dashboard_RenderNextCriticalActivity(ByVal ws As Worksheet, ByVal tblCalc As ListObject, ByVal mapCalc As Object, ByVal X As Double, ByVal Y As Double, ByVal w As Double, ByVal h As Double)

    Dim arr As Variant
    Dim r As Long
    Dim startVal As Variant
    Dim startSerial As Double
    Dim todaySerial As Double
    Dim bestFound As Boolean
    Dim bestDate As Double
    Dim bestWbs As String
    Dim bestName As String
    Dim bestFloat As Variant
    Dim fallbackFound As Boolean
    Dim fallbackDate As Double
    Dim fallbackWbs As String
    Dim fallbackName As String
    Dim fallbackFloat As Variant
    Dim accentColor As Long
    Dim dateColor As Long
    Dim daysToStart As Double

    On Error GoTo RenderFailed

    If tblCalc Is Nothing Then GoTo NoData
    If mapCalc Is Nothing Then GoTo NoData
    If tblCalc.DataBodyRange Is Nothing Then GoTo NoData
    If Not mapCalc.exists("Critical Path") Then GoTo NoData
    If Not mapCalc.exists("Calculated Start") Then GoTo NoData

    arr = tblCalc.DataBodyRange.value
    todaySerial = CDbl(Date)

    For r = 1 To UBound(arr, 1)
        If Dashboard_IsActiveDashboardTask(arr, mapCalc, r) Then
            If UCase$(Trim$(CStr(Dashboard_ArrayVal(arr, mapCalc, r, "Critical Path")))) = "CRITICAL" Then
                If Not Dashboard_IsMilestoneRow(arr, mapCalc, r) Then
                    startVal = GetCellValue(Dashboard_ArrayVal(arr, mapCalc, r, "Calculated Start"))
                    If Dashboard_HasDateValue(startVal) Then
                        startSerial = Dashboard_DateNumber(startVal)
                        If startSerial >= todaySerial Then
                            If Not bestFound Or startSerial < bestDate Then
                                bestFound = True
                                bestDate = startSerial
                                bestWbs = CStr(Dashboard_ArrayVal(arr, mapCalc, r, "WBS"))
                                bestName = CStr(Dashboard_ArrayVal(arr, mapCalc, r, "Task Name"))
                                If mapCalc.exists("Total Float") Then bestFloat = Dashboard_ArrayVal(arr, mapCalc, r, "Total Float") Else bestFloat = Empty
                            End If
                        ElseIf Not fallbackFound Or startSerial > fallbackDate Then
                            fallbackFound = True
                            fallbackDate = startSerial
                            fallbackWbs = CStr(Dashboard_ArrayVal(arr, mapCalc, r, "WBS"))
                            fallbackName = CStr(Dashboard_ArrayVal(arr, mapCalc, r, "Task Name"))
                            If mapCalc.exists("Total Float") Then fallbackFloat = Dashboard_ArrayVal(arr, mapCalc, r, "Total Float") Else fallbackFloat = Empty
                        End If
                    End If
                End If
            End If
        End If
    Next r

    If Not bestFound And fallbackFound Then
        bestFound = True
        bestDate = fallbackDate
        bestWbs = fallbackWbs
        bestName = fallbackName
        bestFloat = fallbackFloat
    End If

    accentColor = RGB(160, 170, 181)
    dateColor = RGB(96, 111, 128)
    If bestFound Then
        daysToStart = bestDate - todaySerial
        If (IsNumeric(bestFloat) And CDbl(bestFloat) <= 0#) Or daysToStart <= 0# Then
            accentColor = RGB(192, 80, 77)
            dateColor = RGB(192, 80, 77)
        ElseIf daysToStart <= 14# Then
            accentColor = RGB(238, 156, 68)
            dateColor = RGB(238, 156, 68)
        Else
            accentColor = RGB(68, 114, 196)
            dateColor = RGB(68, 114, 196)
        End If
    End If

    Dashboard_AddHotSpotCard ws, X, Y, w, h, TextCatalog_Get("DASHBOARD.CARD.NEXT_CRITICAL", Dashboard_CurrentLanguage()), "", accentColor

    If bestFound Then
        Dashboard_AddHotSpotHero ws, Dashboard_TruncateText(bestName, 24), X + 16, Y + 58, w - 32, 28, RGB(20, 34, 51)
        Dashboard_AddHotSpotText ws, bestWbs, X + 16, Y + 86, w - 32, 14, RGB(96, 111, 128), 8, False, xlHAlignLeft
        Dashboard_AddHotSpotText ws, TextCatalog_Format("DASHBOARD.TASK.STARTS", Dashboard_CurrentLanguage(), TextCatalog_Arguments("Date", Dashboard_FormatDate(bestDate, True))), X + 16, Y + 112, w - 32, 16, RGB(42, 52, 65), 9, False, xlHAlignLeft
        Dashboard_AddHotSpotText ws, Dashboard_DaysRemainingText(CLng(bestDate - todaySerial)), X + 16, Y + 136, w - 32, 16, dateColor, 9, True, xlHAlignLeft
        If IsNumeric(bestFloat) Then Dashboard_AddHotSpotText ws, TextCatalog_Format("DASHBOARD.TASK.FLOAT", Dashboard_CurrentLanguage(), TextCatalog_Arguments("Days", CStr(CLng(bestFloat)) & Dashboard_DurationSuffix())), X + 16, Y + h - 24, w - 32, 16, RGB(96, 111, 128), 8, False, xlHAlignLeft
    Else
        Dashboard_AddHotSpotText ws, TextCatalog_Get("DASHBOARD.EMPTY.NO_ACTIVE_CRITICAL", Dashboard_CurrentLanguage()), X + 16, Y + 76, w - 32, 18, RGB(96, 111, 128), 8, False, xlHAlignLeft
    End If

    Exit Sub

NoData:
    Dashboard_AddHotSpotCard ws, X, Y, w, h, TextCatalog_Get("DASHBOARD.CARD.NEXT_CRITICAL", Dashboard_CurrentLanguage()), "", RGB(160, 170, 181)
    Dashboard_AddHotSpotText ws, TextCatalog_Get("DASHBOARD.EMPTY.NO_CRITICAL_DATA", Dashboard_CurrentLanguage()), X + 16, Y + 76, w - 32, 18, RGB(96, 111, 128), 8, False, xlHAlignLeft
    Exit Sub

RenderFailed:
    Dashboard_WriteHotSpotRuntimeError ws, X + 16, Y + 76, w - 32, TextCatalog_Get("DASHBOARD.CARD.NEXT_CRITICAL", Dashboard_CurrentLanguage())

End Sub

'------------------------------------------------------------------------------
' FR: Emet un checkpoint de diagnostic pour la construction des hot spots Dashboard lorsque la trace est active.
' EN: Emits a diagnostic checkpoint for Dashboard hot-spot construction when tracing is enabled.
'------------------------------------------------------------------------------

Private Sub Dashboard_HotSpotTrace(ByVal helperName As String, ByVal stepName As String)

    gHotSpotHelper = helperName
    gHotSpotStep = stepName
    Debug.Print "Dashboard Hot Spots | " & helperName & " | " & stepName

End Sub

'------------------------------------------------------------------------------
' FR: Ecrit ou synchronise Write Hot Spot Runtime Error dans le stockage possede par le domaine.
' EN: Writes or synchronizes Write Hot Spot Runtime Error in the store owned by the domain.
'------------------------------------------------------------------------------

Private Sub Dashboard_WriteHotSpotRuntimeError( _
    ByVal ws As Worksheet, _
    ByVal X As Double, _
    ByVal Y As Double, _
    ByVal w As Double, _
    ByVal cardName As String)

    Dim errNum As Long
    Dim errDesc As String
    Dim msg As String

    errNum = Err.number
    errDesc = Err.description
    msg = TextCatalog_Format("DASHBOARD.ERROR.RENDER", Dashboard_CurrentLanguage(), TextCatalog_Arguments("Number", CStr(errNum), "Details", errDesc, "Card", cardName, "Helper", gHotSpotHelper, "Step", gHotSpotStep))
    Debug.Print "Dashboard Hot Spots | " & cardName & " | " & gHotSpotHelper & " | " & gHotSpotStep & " | RenderFailed " & CStr(errNum) & ": " & errDesc

    On Error Resume Next
    Dashboard_AddHotSpotText ws, msg, X, Y, w, 28, RGB(192, 80, 77), 8, False, xlHAlignLeft

End Sub

'------------------------------------------------------------------------------
' FR: Ajoute la reference Hot Spot Card a la structure cible fournie par l'appelant.
' EN: Adds the Hot Spot Card reference to the target structure supplied by the caller.
' FR - Effet de bord : cree ou met a jour des shapes Excel.
' EN - Side effect: creates or updates Excel shapes.
'------------------------------------------------------------------------------

Private Sub Dashboard_AddHotSpotCard( _
    ByVal ws As Worksheet, _
    ByVal X As Double, _
    ByVal Y As Double, _
    ByVal w As Double, _
    ByVal h As Double, _
    ByVal titleText As String, _
    ByVal subtitleText As String, _
    ByVal accentColor As Long)

    Dim bg As shape
    Dim accent As shape

    Dashboard_HotSpotTrace "Dashboard_AddHotSpotCard", "Add background shape"
    Set bg = ws.Shapes.AddShape(msoShapeRoundedRectangle, X, Y, w, h)

    Dashboard_HotSpotTrace "Dashboard_AddHotSpotCard", "Format background"
    bg.name = DASH_PREFIX & "HOTSPOT_CARD"
    bg.fill.ForeColor.RGB = RGB(255, 255, 255)
    bg.Line.ForeColor.RGB = RGB(224, 230, 237)

    Dashboard_HotSpotTrace "Dashboard_AddHotSpotCard", "Add accent shape"
    Set accent = ws.Shapes.AddShape(msoShapeRectangle, X, Y, 5, h)

    Dashboard_HotSpotTrace "Dashboard_AddHotSpotCard", "Format accent"
    accent.name = DASH_PREFIX & "HOTSPOT_ACCENT"
    accent.fill.ForeColor.RGB = accentColor
    accent.Line.Visible = msoFalse

    Dashboard_HotSpotTrace "Dashboard_AddHotSpotCard", "Add title text"
    Dashboard_AddHotSpotText ws, titleText, X + 16, Y + 11, w - 30, 18, RGB(20, 34, 51), 11, True, xlHAlignLeft

    If Len(Trim$(subtitleText)) > 0 Then
        Dashboard_HotSpotTrace "Dashboard_AddHotSpotCard", "Add subtitle text"
        Dashboard_AddHotSpotText ws, subtitleText, X + 16, Y + 32, w - 30, 14, RGB(126, 139, 153), 7, False, xlHAlignLeft
    End If

End Sub

'------------------------------------------------------------------------------
' FR: Ajoute la reference Empty Hot Spot Card a la structure cible fournie par l'appelant.
' EN: Adds the Empty Hot Spot Card reference to the target structure supplied by the caller.
'------------------------------------------------------------------------------

Private Sub Dashboard_AddEmptyHotSpotCard(ByVal ws As Worksheet, ByVal X As Double, ByVal Y As Double, ByVal w As Double, ByVal h As Double, ByVal titleText As String)

    Dashboard_AddHotSpotCard ws, X, Y, w, h, titleText, "", RGB(160, 170, 181)
    Dashboard_AddHotSpotText ws, Dashboard_NoProjectLoadedText(), X + 16, Y + 58, w - 32, 18, RGB(96, 111, 128), 9, True, xlHAlignLeft

End Sub

'------------------------------------------------------------------------------
' FR: Ajoute la reference Hot Spot Kpi Block a la structure cible fournie par l'appelant.
' EN: Adds the Hot Spot Kpi Block reference to the target structure supplied by the caller.
'------------------------------------------------------------------------------

Private Sub Dashboard_AddHotSpotKpiBlock( _
    ByVal ws As Worksheet, _
    ByVal X As Double, _
    ByVal Y As Double, _
    ByVal w As Double, _
    ByVal h As Double, _
    ByVal labelText As String, _
    ByVal valueText As String, _
    ByVal colorVal As Long, _
    ByVal fillColor As Long)

    Dashboard_HotSpotTrace "Dashboard_AddHotSpotKpiBlock", "Add KPI background"
    Dashboard_AddHotSpotBlock ws, X, Y, w, h, fillColor, colorVal
    Dashboard_AddHotSpotText ws, valueText, X + 12, Y + 10, w - 24, 28, colorVal, 18, True, xlHAlignLeft
    Dashboard_AddHotSpotText ws, labelText, X + 12, Y + 40, w - 24, 14, RGB(96, 111, 128), 8, False, xlHAlignLeft

End Sub

'------------------------------------------------------------------------------
' FR: Ajoute la reference Hot Spot Insight Block a la structure cible fournie par l'appelant.
' EN: Adds the Hot Spot Insight Block reference to the target structure supplied by the caller.
'------------------------------------------------------------------------------

Private Sub Dashboard_AddHotSpotInsightBlock( _
    ByVal ws As Worksheet, _
    ByVal X As Double, _
    ByVal Y As Double, _
    ByVal w As Double, _
    ByVal h As Double, _
    ByVal labelText As String, _
    ByVal wbsText As String, _
    ByVal nameText As String, _
    ByVal metricText As String, _
    ByVal colorVal As Long, _
    ByVal fillColor As Long)

    Dashboard_HotSpotTrace "Dashboard_AddHotSpotInsightBlock", "Add insight background"
    Dashboard_AddHotSpotBlock ws, X, Y, w, h, fillColor, colorVal
    Dashboard_AddHotSpotText ws, labelText, X + 12, Y + 8, w - 24, 13, RGB(96, 111, 128), 7, True, xlHAlignLeft
    If Len(wbsText) > 0 Then Dashboard_AddHotSpotText ws, wbsText, X + 12, Y + 27, 72, 14, colorVal, 8, True, xlHAlignLeft
    Dashboard_AddHotSpotText ws, nameText, X + 12, Y + 43, w - 86, 16, RGB(20, 34, 51), 8, False, xlHAlignLeft
    If Len(metricText) > 0 Then Dashboard_AddHotSpotText ws, metricText, X + w - 70, Y + 30, 58, 18, colorVal, 10, True, xlHAlignRight

End Sub

'------------------------------------------------------------------------------
' FR: Ajoute la reference Hot Spot Ranking Item a la structure cible fournie par l'appelant.
' EN: Adds the Hot Spot Ranking Item reference to the target structure supplied by the caller.
'------------------------------------------------------------------------------

Private Sub Dashboard_AddHotSpotRankingItem( _
    ByVal ws As Worksheet, _
    ByVal X As Double, _
    ByVal Y As Double, _
    ByVal w As Double, _
    ByVal h As Double, _
    ByVal rankText As String, _
    ByVal wbsText As String, _
    ByVal nameText As String, _
    ByVal metricText As String, _
    ByVal colorVal As Long)

    Dashboard_HotSpotTrace "Dashboard_AddHotSpotRankingItem", "Add ranking background"
    Dashboard_AddHotSpotBlock ws, X, Y, w, h, RGB(249, 251, 253), colorVal
    Dashboard_AddHotSpotPill ws, rankText, X + 12, Y + 13, 28, 28, Dashboard_LightenColor(colorVal, 0.86), colorVal, 9, True
    Dashboard_AddHotSpotText ws, wbsText, X + 50, Y + 9, w - 124, 13, RGB(96, 111, 128), 7, False, xlHAlignLeft
    Dashboard_AddHotSpotText ws, nameText, X + 50, Y + 26, w - 124, 16, RGB(20, 34, 51), 9, False, xlHAlignLeft
    Dashboard_AddHotSpotPill ws, metricText, X + w - 72, Y + 15, 58, 24, Dashboard_LightenColor(colorVal, 0.88), colorVal, 8, True

End Sub

'------------------------------------------------------------------------------
' FR: Ajoute la reference Hot Spot Metric a la structure cible fournie par l'appelant.
' EN: Adds the Hot Spot Metric reference to the target structure supplied by the caller.
'------------------------------------------------------------------------------

Private Sub Dashboard_AddHotSpotMetric( _
    ByVal ws As Worksheet, _
    ByVal X As Double, _
    ByVal Y As Double, _
    ByVal w As Double, _
    ByVal labelText As String, _
    ByVal valueText As String, _
    ByVal colorVal As Long)

    Dashboard_AddHotSpotKpiBlock ws, X, Y, w, 62, labelText, valueText, colorVal, Dashboard_LightenColor(colorVal, 0.9)

End Sub

'------------------------------------------------------------------------------
' FR: Ajoute la reference Hot Spot Label a la structure cible fournie par l'appelant.
' EN: Adds the Hot Spot Label reference to the target structure supplied by the caller.
'------------------------------------------------------------------------------

Private Sub Dashboard_AddHotSpotLabel(ByVal ws As Worksheet, ByVal txt As String, ByVal X As Double, ByVal Y As Double, ByVal w As Double)

    Dashboard_HotSpotTrace "Dashboard_AddHotSpotLabel", "Add label text"
    Dashboard_AddHotSpotText ws, txt, X, Y, w, 14, RGB(96, 111, 128), 8, True, xlHAlignLeft

End Sub

'------------------------------------------------------------------------------
' FR: Ajoute la reference Hot Spot Hero a la structure cible fournie par l'appelant.
' EN: Adds the Hot Spot Hero reference to the target structure supplied by the caller.
'------------------------------------------------------------------------------

Private Sub Dashboard_AddHotSpotHero( _
    ByVal ws As Worksheet, _
    ByVal txt As String, _
    ByVal X As Double, _
    ByVal Y As Double, _
    ByVal w As Double, _
    ByVal h As Double, _
    ByVal colorVal As Long)

    Dashboard_HotSpotTrace "Dashboard_AddHotSpotHero", "Add hero text"
    Dashboard_AddHotSpotText ws, txt, X, Y, w, h, colorVal, 13, True, xlHAlignLeft

End Sub

'------------------------------------------------------------------------------
' FR: Ajoute la collection Hot Spot List Item a la structure cible fournie par l'appelant.
' EN: Adds the Hot Spot List Item collection to the target structure supplied by the caller.
'------------------------------------------------------------------------------

Private Sub Dashboard_AddHotSpotListItem( _
    ByVal ws As Worksheet, _
    ByVal X As Double, _
    ByVal Y As Double, _
    ByVal w As Double, _
    ByVal wbsText As String, _
    ByVal nameText As String, _
    ByVal metricText As String, _
    ByVal metricColor As Long)

    Dashboard_AddHotSpotRankingItem ws, X, Y, w, 48, "", wbsText, nameText, metricText, metricColor

End Sub

'------------------------------------------------------------------------------
' FR: Ajoute la reference Hot Spot Block a la structure cible fournie par l'appelant.
' EN: Adds the Hot Spot Block reference to the target structure supplied by the caller.
' FR - Effet de bord : cree ou met a jour des shapes Excel.
' EN - Side effect: creates or updates Excel shapes.
'------------------------------------------------------------------------------

Private Sub Dashboard_AddHotSpotBlock( _
    ByVal ws As Worksheet, _
    ByVal X As Double, _
    ByVal Y As Double, _
    ByVal w As Double, _
    ByVal h As Double, _
    ByVal fillColor As Long, _
    ByVal lineColor As Long)

    Dim shp As shape

    Dashboard_HotSpotTrace "Dashboard_AddHotSpotBlock", "Add block shape"
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, X, Y, w, h)
    shp.name = DASH_PREFIX & "HOTSPOT_BLOCK"
    shp.fill.ForeColor.RGB = fillColor
    shp.Line.ForeColor.RGB = Dashboard_LightenColor(lineColor, 0.7)

End Sub

'------------------------------------------------------------------------------
' FR: Ajoute la reference Hot Spot Pill a la structure cible fournie par l'appelant.
' EN: Adds the Hot Spot Pill reference to the target structure supplied by the caller.
' FR - Effet de bord : cree ou met a jour des shapes Excel.
' EN - Side effect: creates or updates Excel shapes.
'------------------------------------------------------------------------------

Private Sub Dashboard_AddHotSpotPill( _
    ByVal ws As Worksheet, _
    ByVal txt As String, _
    ByVal X As Double, _
    ByVal Y As Double, _
    ByVal w As Double, _
    ByVal h As Double, _
    ByVal fillColor As Long, _
    ByVal colorVal As Long, _
    ByVal fontSize As Double, _
    ByVal isBold As Boolean)

    Dim shp As shape

    Dashboard_HotSpotTrace "Dashboard_AddHotSpotPill", "Add pill shape"
    Set shp = ws.Shapes.AddShape(msoShapeOval, X, Y, w, h)
    shp.name = DASH_PREFIX & "HOTSPOT_PILL"
    shp.fill.ForeColor.RGB = fillColor
    shp.Line.Visible = msoFalse
    Dashboard_AddHotSpotText ws, txt, X, Y + 1, w, h - 2, colorVal, fontSize, isBold, xlHAlignCenter

End Sub

'------------------------------------------------------------------------------
' FR: Ajoute la reference Hot Spot Text a la structure cible fournie par l'appelant.
' EN: Adds the Hot Spot Text reference to the target structure supplied by the caller.
' FR - Effet de bord : cree ou met a jour des shapes Excel.
' EN - Side effect: creates or updates Excel shapes.
'------------------------------------------------------------------------------

Private Sub Dashboard_AddHotSpotText( _
    ByVal ws As Worksheet, _
    ByVal txt As String, _
    ByVal X As Double, _
    ByVal Y As Double, _
    ByVal w As Double, _
    ByVal h As Double, _
    ByVal colorVal As Long, _
    ByVal fontSize As Double, _
    ByVal isBold As Boolean, _
    ByVal horizontalAlign As Long)

    Dim shp As shape

    Dashboard_HotSpotTrace "Dashboard_AddHotSpotText", "Add textbox"
    Set shp = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, X, Y, w, h)

    Dashboard_HotSpotTrace "Dashboard_AddHotSpotText", "Format shape fill line"
    shp.name = DASH_PREFIX & "HOTSPOT_TXT"
    shp.fill.Visible = msoFalse
    shp.Line.Visible = msoFalse

    Dashboard_HotSpotTrace "Dashboard_AddHotSpotText", "Format legacy TextFrame"
    With shp.TextFrame
        .MarginLeft = 0
        .MarginRight = 0
        .MarginTop = 0
        .MarginBottom = 0
        .HorizontalAlignment = horizontalAlign
        .VerticalAlignment = xlVAlignCenter
        .characters.text = txt
        .characters.font.name = "Segoe UI"
        .characters.font.size = fontSize
        .characters.font.color = colorVal
        .characters.font.bold = isBold
    End With

End Sub

'------------------------------------------------------------------------------
' FR: Retourne la valeur Lighten Color sans modifier les donnees d'entree.
' EN: Returns the Lighten Color value without mutating input data.
'------------------------------------------------------------------------------

Private Function Dashboard_LightenColor(ByVal colorVal As Long, ByVal ratio As Double) As Long

    Dim r As Long
    Dim g As Long
    Dim b As Long

    r = colorVal Mod 256
    g = (colorVal \ 256) Mod 256
    b = (colorVal \ 65536) Mod 256

    r = CLng(r + ((255 - r) * ratio))
    g = CLng(g + ((255 - g) * ratio))
    b = CLng(b + ((255 - b) * ratio))

    Dashboard_LightenColor = RGB(r, g, b)

End Function
'------------------------------------------------------------------------------
' FR: Indique si la map Active Dashboard Task satisfait la condition attendue, sans modifier les donnees source.
' EN: Returns whether the Active Dashboard Task map satisfies the expected condition without mutating source data.
'------------------------------------------------------------------------------

Private Function Dashboard_IsActiveDashboardTask(ByRef arr As Variant, ByVal mapTbl As Object, ByVal r As Long) As Boolean

    Dim progressVal As Variant
    Dim actualFinish As Variant

    If mapTbl.exists("Actual Finish") Then
        actualFinish = GetCellValue(Dashboard_ArrayVal(arr, mapTbl, r, "Actual Finish"))
        If HasValue(actualFinish) Then Exit Function
    End If

    If mapTbl.exists("% Progress") Then
        progressVal = GetCellValue(Dashboard_ArrayVal(arr, mapTbl, r, "% Progress"))
        If IsNumeric(progressVal) Then
            If Dashboard_NormalizeProgress(progressVal) >= 1# Then Exit Function
        End If
    End If

    Dashboard_IsActiveDashboardTask = True

End Function

'------------------------------------------------------------------------------
' FR: Indique si la map Milestone Row satisfait la condition attendue, sans modifier les donnees source.
' EN: Returns whether the Milestone Row map satisfies the expected condition without mutating source data.
'------------------------------------------------------------------------------

Private Function Dashboard_IsMilestoneRow(ByRef arr As Variant, ByVal mapTbl As Object, ByVal r As Long) As Boolean

    Dim taskTypeVal As String
    Dim durationVal As Variant

    If mapTbl.exists("Task Type") Then
        taskTypeVal = UCase$(Trim$(CStr(Dashboard_ArrayVal(arr, mapTbl, r, "Task Type"))))
        Dashboard_IsMilestoneRow = (taskTypeVal = "MILESTONE")
        Exit Function
    End If

    If mapTbl.exists("Calculated Duration") Then
        durationVal = GetCellValue(Dashboard_ArrayVal(arr, mapTbl, r, "Calculated Duration"))
        If IsNumeric(durationVal) Then Dashboard_IsMilestoneRow = (CDbl(durationVal) <= 1#)
    End If

End Function

'------------------------------------------------------------------------------
' FR: Retourne la valeur Days Remaining Text sans modifier les donnees d'entree.
' EN: Returns the Days Remaining Text value without mutating input data.
'------------------------------------------------------------------------------

Private Function Dashboard_DaysRemainingText(ByVal daysVal As Long) As String

    If daysVal < 0 Then
        Dashboard_DaysRemainingText = TextCatalog_Format("DASHBOARD.DEADLINE.OVERDUE", Dashboard_CurrentLanguage(), TextCatalog_Arguments("Count", CStr(Abs(daysVal))))
    ElseIf daysVal = 0 Then
        Dashboard_DaysRemainingText = TextCatalog_Get("DASHBOARD.LABEL.TODAY", Dashboard_CurrentLanguage())
    Else
        Dashboard_DaysRemainingText = TextCatalog_Format("DASHBOARD.DEADLINE.REMAINING", Dashboard_CurrentLanguage(), TextCatalog_Arguments("Count", CStr(daysVal)))
    End If

End Function
'------------------------------------------------------------------------------
' FR: Actualise Render Forecast Issues sans modifier les regles metier qui produisent les donnees.
' EN: Refreshes Render Forecast Issues without changing the business rules that produce the data.
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' FR: Retourne la map Max Date sans modifier les donnees d'entree.
' EN: Returns the Max Date map without mutating input data.
'------------------------------------------------------------------------------

Private Function Dashboard_MaxDate(ByVal tbl As ListObject, ByVal mapTbl As Object, ByVal colName As String) As Variant

    Dim arr As Variant
    Dim r As Long
    Dim v As Variant
    Dim hasDate As Boolean
    Dim maxVal As Double

    If tbl Is Nothing Then Exit Function
    If tbl.DataBodyRange Is Nothing Then Exit Function
    If Not mapTbl.exists(colName) Then Exit Function

    arr = tbl.DataBodyRange.value
    For r = 1 To UBound(arr, 1)
        v = GetCellValue(arr(r, mapTbl(colName)))
        If Dashboard_HasDateValue(v) Then
            If Not hasDate Then
                maxVal = Dashboard_DateNumber(v)
                hasDate = True
            ElseIf Dashboard_DateNumber(v) > maxVal Then
                maxVal = Dashboard_DateNumber(v)
            End If
        End If
    Next r

    If hasDate Then Dashboard_MaxDate = maxVal

End Function

'------------------------------------------------------------------------------
' FR: Retourne la map Min Date Across sans modifier les donnees d'entree.
' EN: Returns the Min Date Across map without mutating input data.
'------------------------------------------------------------------------------

Private Function Dashboard_MinDateAcross(ByVal tbl As ListObject, ByVal mapTbl As Object, ByVal colNames As Variant) As Variant

    Dim arr As Variant
    Dim r As Long
    Dim c As Variant
    Dim v As Variant
    Dim hasDate As Boolean
    Dim minVal As Double

    If tbl Is Nothing Then Exit Function
    If tbl.DataBodyRange Is Nothing Then Exit Function
    arr = tbl.DataBodyRange.value

    For r = 1 To UBound(arr, 1)
        For Each c In colNames
            If mapTbl.exists(CStr(c)) Then
                v = GetCellValue(arr(r, mapTbl(CStr(c))))
                If Dashboard_HasDateValue(v) Then
                    If Not hasDate Then
                        minVal = Dashboard_DateNumber(v)
                        hasDate = True
                    ElseIf Dashboard_DateNumber(v) < minVal Then
                        minVal = Dashboard_DateNumber(v)
                    End If
                End If
            End If
        Next c
    Next r

    If hasDate Then Dashboard_MinDateAcross = minVal

End Function

'------------------------------------------------------------------------------
' FR: Retourne la map Max Date Across sans modifier les donnees d'entree.
' EN: Returns the Max Date Across map without mutating input data.
'------------------------------------------------------------------------------

Private Function Dashboard_MaxDateAcross(ByVal tbl As ListObject, ByVal mapTbl As Object, ByVal colNames As Variant) As Variant

    Dim arr As Variant
    Dim r As Long
    Dim c As Variant
    Dim v As Variant
    Dim hasDate As Boolean
    Dim maxVal As Double

    If tbl Is Nothing Then Exit Function
    If tbl.DataBodyRange Is Nothing Then Exit Function
    arr = tbl.DataBodyRange.value

    For r = 1 To UBound(arr, 1)
        For Each c In colNames
            If mapTbl.exists(CStr(c)) Then
                v = GetCellValue(arr(r, mapTbl(CStr(c))))
                If Dashboard_HasDateValue(v) Then
                    If Not hasDate Then
                        maxVal = Dashboard_DateNumber(v)
                        hasDate = True
                    ElseIf Dashboard_DateNumber(v) > maxVal Then
                        maxVal = Dashboard_DateNumber(v)
                    End If
                End If
            End If
        Next c
    Next r

    If hasDate Then Dashboard_MaxDateAcross = maxVal

End Function

'------------------------------------------------------------------------------
' FR: Retourne la map Numeric Below sans exposer de mutateur sur l'etat source.
' EN: Returns the Numeric Below map without exposing a mutator for source state.
'------------------------------------------------------------------------------

Private Function Dashboard_CountNumericBelow(ByVal tbl As ListObject, ByVal mapTbl As Object, ByVal colName As String, ByVal threshold As Double) As Long

    Dim arr As Variant
    Dim r As Long
    Dim v As Variant

    If tbl Is Nothing Then Exit Function
    If tbl.DataBodyRange Is Nothing Then Exit Function
    If Not mapTbl.exists(colName) Then Exit Function

    arr = tbl.DataBodyRange.value
    For r = 1 To UBound(arr, 1)
        v = GetCellValue(arr(r, mapTbl(colName)))
        If IsNumeric(v) Then
            If CDbl(v) < threshold Then Dashboard_CountNumericBelow = Dashboard_CountNumericBelow + 1
        End If
    Next r

End Function

'------------------------------------------------------------------------------
' FR: Compte les lignes portant un diagnostic Forecast structure.
' EN: Counts rows carrying a structured Forecast diagnostic.
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' FR: Reconnait les diagnostics Forecast par code stable, jamais par texte rendu.
' EN: Recognizes Forecast diagnostics by stable code, never rendered text.
'------------------------------------------------------------------------------
'------------------------------------------------------------------------------
' FR: Retourne la map Marker sans exposer de mutateur sur l'etat source.
' EN: Returns the Marker map without exposing a mutator for source state.
'------------------------------------------------------------------------------

Private Function Dashboard_CountMarker(ByVal tbl As ListObject, ByVal mapTbl As Object, ByVal colName As String) As Long

    Dim arr As Variant
    Dim r As Long
    Dim txt As String

    If tbl Is Nothing Then Exit Function
    If tbl.DataBodyRange Is Nothing Then Exit Function
    If Not mapTbl.exists(colName) Then Exit Function

    arr = tbl.DataBodyRange.value
    For r = 1 To UBound(arr, 1)
        txt = UCase$(Trim$(CStr(arr(r, mapTbl(colName)))))
        If txt <> "" And txt <> "NO" And txt <> "FALSE" And txt <> "0" Then Dashboard_CountMarker = Dashboard_CountMarker + 1
    Next r

End Function

'------------------------------------------------------------------------------
' FR: Retourne la map Marker Remaining sans exposer de mutateur sur l'etat source.
' EN: Returns the Marker Remaining map without exposing a mutator for source state.
'------------------------------------------------------------------------------

Private Function Dashboard_CountMarkerRemaining(ByVal tbl As ListObject, ByVal mapTbl As Object, ByVal colName As String) As Long

    Dim arr As Variant
    Dim r As Long
    Dim txt As String
    Dim progressVal As Double
    Dim hasProgress As Boolean
    Dim actualFinish As Variant

    If tbl Is Nothing Then Exit Function
    If tbl.DataBodyRange Is Nothing Then Exit Function
    If Not mapTbl.exists(colName) Then Exit Function

    hasProgress = mapTbl.exists("% Progress")

    arr = tbl.DataBodyRange.value
    For r = 1 To UBound(arr, 1)
        txt = UCase$(Trim$(CStr(arr(r, mapTbl(colName)))))
        If txt <> "" And txt <> "NO" And txt <> "FALSE" And txt <> "0" Then
            If hasProgress Then
                progressVal = Dashboard_NormalizeProgress(GetCellValue(arr(r, mapTbl("% Progress"))))
            ElseIf mapTbl.exists("Actual Finish") Then
                actualFinish = GetCellValue(arr(r, mapTbl("Actual Finish")))
                If Dashboard_HasDateValue(actualFinish) Then progressVal = 1# Else progressVal = 0#
            Else
                progressVal = 0#
            End If
            If progressVal < 1# Then Dashboard_CountMarkerRemaining = Dashboard_CountMarkerRemaining + 1
        End If
    Next r

End Function

'------------------------------------------------------------------------------
' FR: Retourne la valeur Contract Status sans modifier les donnees d'entree.
' EN: Returns the Contract Status value without mutating input data.
'------------------------------------------------------------------------------

Private Function Dashboard_ContractStatus(ByVal driftDays As Variant, ByRef statusColor As Long) As String

    statusColor = RGB(96, 111, 128)

    If Not HasValue(driftDays) Then
        Dashboard_ContractStatus = "UNKNOWN"
    ElseIf CLng(driftDays) <= 0 Then
        Dashboard_ContractStatus = "ON CONTRACT"
        statusColor = RGB(0, 145, 112)
    ElseIf CLng(driftDays) <= 10 Then
        Dashboard_ContractStatus = "MINOR DELAY"
        statusColor = RGB(238, 156, 68)
    Else
        Dashboard_ContractStatus = "CONTRACT DELAY"
        statusColor = RGB(192, 80, 77)
    End If

End Function

'------------------------------------------------------------------------------
' FR: Construit la map Metrics a partir des donnees fournies par l'appelant.
' EN: Builds the Metrics map from data supplied by the caller.
'------------------------------------------------------------------------------

Private Function Dashboard_BuildMetrics( _
    ByVal tblCalc As ListObject, _
    ByVal mapCalc As Object, _
    ByVal scurveProjection As Object) As Object

    Dim metrics As Object
    Dim baselineFinish As Variant
    Dim calcFinish As Variant
    Dim driftDays As Variant
    Dim contractColor As Long

    Set metrics = CreateObject("Scripting.Dictionary")
    metrics("ActualProgress") = CDbl(scurveProjection("ActualProgress"))
    metrics("PlannedProgress") = CDbl(scurveProjection("PlannedProgress"))
    metrics("ProgressVariance") = CDbl(metrics("ActualProgress")) - CDbl(metrics("PlannedProgress"))
    baselineFinish = Dashboard_MaxDate(tblCalc, mapCalc, "Baseline Finish")
    calcFinish = Dashboard_MaxDate(tblCalc, mapCalc, "Calculated Finish")
    metrics("BaselineFinish") = baselineFinish
    metrics("ForecastFinish") = calcFinish
    If Dashboard_HasDateValue(baselineFinish) And Dashboard_HasDateValue(calcFinish) Then
        driftDays = CLng(Dashboard_DateNumber(calcFinish) - Dashboard_DateNumber(baselineFinish))
        metrics("DriftDays") = driftDays
    Else
        metrics("DriftDays") = Empty
    End If
    metrics("DeadlineRiskCount") = Dashboard_CountNumericBelow(tblCalc, mapCalc, "Deadline Float", 0#)
    metrics("LongestPathCount") = Dashboard_CountMarker(tblCalc, mapCalc, "Longest Path")
    metrics("LongestPathRemaining") = Dashboard_CountMarkerRemaining(tblCalc, mapCalc, "Longest Path")
    metrics("CriticalPathCount") = Dashboard_CountMarker(tblCalc, mapCalc, "Critical Path")
    metrics("ContractStatus") = Dashboard_ContractStatus(metrics("DriftDays"), contractColor)
    Set Dashboard_BuildMetrics = metrics

End Function

'------------------------------------------------------------------------------
' FR: Construit la map Snapshot From Current Data a partir des donnees fournies par l'appelant.
' EN: Builds the Snapshot From Current Data map from data supplied by the caller.
' FR - Effet de bord : ecrit dans une table Excel detenue par le workflow.
' EN - Side effect: writes to an Excel table owned by the workflow.
'------------------------------------------------------------------------------

Private Function Dashboard_CreateSnapshotFromCurrentData() As ListRow

    Dim tblCalc As ListObject
    Dim mapCalc As Object
    Dim scurveProjection As Object
    Dim metrics As Object
    Dim tblSnap As ListObject
    Dim newRow As ListRow
    Dim snapshotId As Long
    Dim ts As Date
    Dim payload() As Variant
    Dim key As Variant
    Dim failureNumber As Long, failureDescription As String

    On Error GoTo Failed

    Set tblCalc = ThisWorkbook.Worksheets("CALC").ListObjects("tbl_CALC")
    Set mapCalc = CanonicalIdentity_BuildColumnMap(tblCalc)
    Set scurveProjection = SCurve_BuildDashboardProjection()
    If Not CBool(scurveProjection("IsComplete")) Then
        Err.Raise 5, "Dashboard_CreateSnapshotFromCurrentData", TextCatalog_Get("DASHBOARD.SNAPSHOT.INCOMPLETE", Dashboard_CurrentLanguage())
    End If
    Set metrics = Dashboard_BuildMetrics(tblCalc, mapCalc, scurveProjection)
    If Not Dashboard_SnapshotMetricsComplete(metrics) Then
        Err.Raise 5, "Dashboard_CreateSnapshotFromCurrentData", TextCatalog_Get("DASHBOARD.SNAPSHOT.INCOMPLETE", Dashboard_CurrentLanguage())
    End If
    Set tblSnap = Dashboard_EnsureSnapshotsTable()
    snapshotId = Dashboard_NextSnapshotId(tblSnap)
    ts = Now
    ReDim payload(1 To 1, 1 To tblSnap.ListColumns.count)
    payload(1, tblSnap.ListColumns("SnapshotId").Index) = snapshotId
    payload(1, tblSnap.ListColumns("SnapshotDateTime").Index) = CDbl(ts)
    payload(1, tblSnap.ListColumns("SnapshotLabel").Index) = Dashboard_SnapshotLabel(snapshotId, ts)
    For Each key In metrics.keys
        payload(1, tblSnap.ListColumns(CStr(key)).Index) = metrics(key)
    Next key
    Set newRow = tblSnap.ListRows.Add
    newRow.Range.Value2 = payload
    Set Dashboard_CreateSnapshotFromCurrentData = newRow
    Exit Function
Failed:
    failureNumber = Err.number: failureDescription = Err.description
    If Not newRow Is Nothing Then newRow.Delete
    On Error GoTo 0
    Err.Raise failureNumber, "Dashboard_CreateSnapshotFromCurrentData", failureDescription
End Function

Private Function Dashboard_SnapshotMetricsComplete(ByVal metrics As Object) As Boolean
    Dim key As Variant
    For Each key In Array("ActualProgress", "PlannedProgress", "ProgressVariance", "DriftDays", _
        "DeadlineRiskCount", "LongestPathCount", "LongestPathRemaining", "CriticalPathCount")
        If Not metrics.exists(CStr(key)) Then Exit Function
        If isError(metrics(key)) Then Exit Function
        If Not HasValue(metrics(key)) Or Not IsNumeric(metrics(key)) Then Exit Function
    Next key
    For Each key In Array("ForecastFinish", "BaselineFinish")
        If Not metrics.exists(CStr(key)) Then Exit Function
        If Not Dashboard_HasDateValue(metrics(key)) Then Exit Function
    Next key
    If Not metrics.exists("ContractStatus") Then Exit Function
    If Len(CStr(metrics("ContractStatus"))) = 0 Then Exit Function
    Dashboard_SnapshotMetricsComplete = True
End Function

'------------------------------------------------------------------------------
' FR: Cree ou remet en conformite la reference Snapshots Table de maniere idempotente.
' EN: Creates or restores the Snapshots Table reference to a compliant state idempotently.
' FR - Effet de bord : ecrit dans une table Excel detenue par le workflow.
' FR - Effet de bord : efface uniquement les donnees ou objets cibles du contrat.
' EN - Side effect: writes to an Excel table owned by the workflow.
' EN - Side effect: clears only data or objects targeted by the contract.
'------------------------------------------------------------------------------

Private Function Dashboard_EnsureSnapshotsTable() As ListObject

    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim headers As Variant
    Dim i As Long
    Dim nm As name

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(DASHBOARD_SNAPSHOT_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(after:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
        ws.name = DASHBOARD_SNAPSHOT_SHEET
    End If

    On Error Resume Next
    Set tbl = ws.ListObjects(DASHBOARD_SNAPSHOT_TABLE)
    On Error GoTo 0

    If tbl Is Nothing Then
        headers = Array("SnapshotId", "SnapshotDateTime", "SnapshotLabel", "ActualProgress", "PlannedProgress", "ProgressVariance", "ForecastFinish", "BaselineFinish", "DriftDays", "ContractStatus", "DeadlineRiskCount", "ForecastIssueCount", "LongestPathCount", "LongestPathRemaining", "CriticalPathCount")
        For i = LBound(headers) To UBound(headers)
            ws.cells(1, i + 1).value = CStr(headers(i))
        Next i
        Set tbl = ws.ListObjects.Add(xlSrcRange, ws.Range(ws.cells(1, 1), ws.cells(1, UBound(headers) + 1)), , xlYes)
        tbl.name = DASHBOARD_SNAPSHOT_TABLE
    End If

    On Error Resume Next
    ThisWorkbook.names("DashboardSnapshotLabels").Delete
    On Error GoTo 0
    If Not tbl.DataBodyRange Is Nothing Then
        ' Refresh the owned physical extent whenever snapshot controls are rebuilt.
        ThisWorkbook.names.Add name:="DashboardSnapshotLabels", _
            RefersTo:="='" & Replace$(ws.name, "'", "''") & "'!" & _
                tbl.ListColumns("SnapshotLabel").DataBodyRange.address(True, True, xlA1)
    End If

    Set Dashboard_EnsureSnapshotsTable = tbl

End Function

'------------------------------------------------------------------------------
' FR: Retourne la reference Next Snapshot ID sans modifier les donnees d'entree.
' EN: Returns the Next Snapshot ID reference without mutating input data.
'------------------------------------------------------------------------------

Private Function Dashboard_NextSnapshotId(ByVal tbl As ListObject) As Long

    Dim arr As Variant
    Dim r As Long
    Dim v As Variant
    Dim maxId As Long

    If tbl Is Nothing Then Exit Function
    If tbl.DataBodyRange Is Nothing Then
        Dashboard_NextSnapshotId = 1
        Exit Function
    End If

    arr = tbl.DataBodyRange.value
    For r = 1 To UBound(arr, 1)
        v = arr(r, tbl.ListColumns("SnapshotId").Index)
        If IsNumeric(v) Then If CLng(v) > maxId Then maxId = CLng(v)
    Next r

    Dashboard_NextSnapshotId = maxId + 1

End Function

'------------------------------------------------------------------------------
' FR: Retourne la valeur Snapshot Label sans modifier les donnees d'entree.
' EN: Returns the Snapshot Label value without mutating input data.
'------------------------------------------------------------------------------

Private Function Dashboard_SnapshotLabel(ByVal snapshotId As Long, ByVal snapshotDate As Date) As String

    Dashboard_SnapshotLabel = TextCatalog_Format( _
        "DASHBOARD.SNAPSHOT.LABEL", Dashboard_CurrentLanguage(), _
        TextCatalog_Arguments( _
            "Id", CStr(snapshotId), _
            "Date", Format$(snapshotDate, "dd/mm/yyyy hh:nn")))

End Function

'------------------------------------------------------------------------------
' FR: Retourne la reference Default Snapshot Labels sans exposer de mutateur sur l'etat source.
' EN: Returns the Default Snapshot Labels reference without exposing a mutator for source state.
'------------------------------------------------------------------------------

Private Sub Dashboard_GetDefaultSnapshotLabels(ByVal tbl As ListObject, ByRef fromLabel As String, ByRef toLabel As String)

    Dim n As Long

    fromLabel = ""
    toLabel = ""
    If tbl Is Nothing Then Exit Sub
    If tbl.DataBodyRange Is Nothing Then Exit Sub

    n = tbl.ListRows.count
    toLabel = CStr(tbl.DataBodyRange.cells(n, tbl.ListColumns("SnapshotLabel").Index).value)
    If n > 1 Then
        fromLabel = CStr(tbl.DataBodyRange.cells(n - 1, tbl.ListColumns("SnapshotLabel").Index).value)
    Else
        fromLabel = toLabel
    End If

End Sub

'------------------------------------------------------------------------------
' FR: Retourne la reference Snapshot Label Exists sans modifier les donnees d'entree.
' EN: Returns the Snapshot Label Exists reference without mutating input data.
'------------------------------------------------------------------------------

Private Function Dashboard_SnapshotLabelExists(ByVal tbl As ListObject, ByVal snapshotLabel As String) As Boolean

    Dim arr As Variant
    Dim r As Long

    If tbl Is Nothing Then Exit Function
    If tbl.DataBodyRange Is Nothing Then Exit Function

    arr = tbl.DataBodyRange.value
    For r = 1 To UBound(arr, 1)
        If CStr(arr(r, tbl.ListColumns("SnapshotLabel").Index)) = snapshotLabel Then
            Dashboard_SnapshotLabelExists = True
            Exit Function
        End If
    Next r

End Function

'------------------------------------------------------------------------------
' FR: Retourne la map Momentum Status sans modifier les donnees d'entree.
' EN: Returns the Momentum Status map without mutating input data.
'------------------------------------------------------------------------------

Private Function Dashboard_MomentumStatus(ByVal fromLabel As String, ByVal toLabel As String, ByRef statusColor As Long, ByRef compareText As String, ByRef subText As String) As String

    Dim tbl As ListObject
    Dim fromRow As Object
    Dim toRow As Object
    Dim progressDelta As Double
    Dim forecastDelta As Double
    Dim riskDelta As Double
    Dim fromDefault As String
    Dim toDefault As String

    statusColor = RGB(96, 111, 128)
    compareText = TextCatalog_Get("DASHBOARD.EMPTY.NO_SNAPSHOTS", Dashboard_CurrentLanguage())
    subText = TextCatalog_Get("DASHBOARD.EMPTY.HISTORY_TITLE", Dashboard_CurrentLanguage())
    Dashboard_MomentumStatus = "INSUFFICIENT HISTORY"

    Set tbl = Dashboard_EnsureSnapshotsTable()
    Dashboard_GetDefaultSnapshotLabels tbl, fromDefault, toDefault
    If fromLabel = "" Then fromLabel = fromDefault
    If toLabel = "" Then toLabel = toDefault

    If fromLabel = "" Or toLabel = "" Then Exit Function

    Set fromRow = Dashboard_GetSnapshotByLabel(tbl, fromLabel)
    Set toRow = Dashboard_GetSnapshotByLabel(tbl, toLabel)
    If fromRow Is Nothing Or toRow Is Nothing Then Exit Function

    compareText = Dashboard_FormatDate(fromRow("SnapshotDateTime"), True) & " " & ChrW(8594) & " " & Dashboard_FormatDate(toRow("SnapshotDateTime"), True)
    If CStr(fromRow("SnapshotId")) = CStr(toRow("SnapshotId")) Then Exit Function

    progressDelta = CDbl(toRow("ActualProgress")) - CDbl(fromRow("ActualProgress"))
    forecastDelta = Dashboard_DateNumber(toRow("ForecastFinish")) - Dashboard_DateNumber(fromRow("ForecastFinish"))
    riskDelta = CDbl(toRow("DeadlineRiskCount")) - CDbl(fromRow("DeadlineRiskCount"))
    subText = TextCatalog_Format("DASHBOARD.MOMENTUM.SUMMARY", Dashboard_CurrentLanguage(), TextCatalog_Arguments("Progress", Dashboard_FormatPercentSigned(progressDelta), "Forecast", Dashboard_FormatSignedCompactDays(forecastDelta), "Risks", Dashboard_FormatSignedNumber(riskDelta)))

    If forecastDelta > 0 Or riskDelta > 0 Then
        Dashboard_MomentumStatus = "DETERIORATING"
        statusColor = RGB(192, 80, 77)
    ElseIf progressDelta > 0 And (forecastDelta < 0 Or riskDelta < 0) Then
        Dashboard_MomentumStatus = "IMPROVING"
        statusColor = RGB(0, 145, 112)
    Else
        Dashboard_MomentumStatus = "STABLE"
        statusColor = RGB(238, 156, 68)
    End If

End Function

'------------------------------------------------------------------------------
' FR: Retourne la map Snapshot By Label sans exposer de mutateur sur l'etat source.
' EN: Returns the Snapshot By Label map without exposing a mutator for source state.
'------------------------------------------------------------------------------

Private Function Dashboard_GetSnapshotByLabel(ByVal tbl As ListObject, ByVal snapshotLabel As String) As Object

    Dim snap As Object
    Dim arr As Variant
    Dim r As Long
    Dim c As Long

    If tbl Is Nothing Then Exit Function
    If tbl.DataBodyRange Is Nothing Then Exit Function

    arr = tbl.DataBodyRange.value
    For r = 1 To UBound(arr, 1)
        If CStr(arr(r, tbl.ListColumns("SnapshotLabel").Index)) = snapshotLabel Then
            Set snap = CreateObject("Scripting.Dictionary")
            For c = 1 To tbl.ListColumns.count
                snap(tbl.ListColumns(c).name) = arr(r, c)
            Next c
            If Not Dashboard_SnapshotMetricsComplete(snap) Then Exit Function
            Set Dashboard_GetSnapshotByLabel = snap
            Exit Function
        End If
    Next r

End Function

'------------------------------------------------------------------------------
' FR: Retourne la map Comparison Rows sans exposer de mutateur sur l'etat source.
' EN: Returns the Comparison Rows map without exposing a mutator for source state.
'------------------------------------------------------------------------------

Private Sub Dashboard_GetComparisonRows(ByVal fromLabel As String, ByVal toLabel As String, ByRef fromRow As Object, ByRef toRow As Object)

    Dim tbl As ListObject
    Dim fromDefault As String
    Dim toDefault As String

    Set tbl = Dashboard_EnsureSnapshotsTable()
    Dashboard_GetDefaultSnapshotLabels tbl, fromDefault, toDefault
    If fromLabel = "" Then fromLabel = fromDefault
    If toLabel = "" Then toLabel = toDefault

    Set fromRow = Dashboard_GetSnapshotByLabel(tbl, fromLabel)
    Set toRow = Dashboard_GetSnapshotByLabel(tbl, toLabel)

End Sub

'------------------------------------------------------------------------------
' FR: Retourne la map Comparison Text sans modifier les donnees d'entree.
' EN: Returns the Comparison Text map without mutating input data.
'------------------------------------------------------------------------------

Private Function Dashboard_ComparisonText(ByVal fromRow As Object, ByVal toRow As Object) As String

    If fromRow Is Nothing Or toRow Is Nothing Then
        Dashboard_ComparisonText = TextCatalog_Get("DASHBOARD.EMPTY.COMPARISON", Dashboard_CurrentLanguage())
        Exit Function
    End If

    If Not Dashboard_HasDateValue(fromRow("SnapshotDateTime")) Or Not Dashboard_HasDateValue(toRow("SnapshotDateTime")) Then
        Dashboard_ComparisonText = TextCatalog_Get("DASHBOARD.EMPTY.COMPARISON", Dashboard_CurrentLanguage())
        Exit Function
    End If

    Dashboard_ComparisonText = Dashboard_FormatDate(fromRow("SnapshotDateTime"), True) & _
        " " & ChrW(8594) & " " & _
        Dashboard_FormatDate(toRow("SnapshotDateTime"), True)

End Function

'------------------------------------------------------------------------------
' FR: Retourne la map Progress Delta Text sans modifier les donnees d'entree.
' EN: Returns the Progress Delta Text map without mutating input data.
'------------------------------------------------------------------------------

Private Function Dashboard_ProgressDeltaText(ByVal fromRow As Object, ByVal toRow As Object) As String

    Dim deltaVal As Double

    If fromRow Is Nothing Or toRow Is Nothing Then
        Dashboard_ProgressDeltaText = TextCatalog_Get("DASHBOARD.EMPTY.DELTA", Dashboard_CurrentLanguage())
        Exit Function
    End If

    deltaVal = CDbl(toRow("ActualProgress")) - CDbl(fromRow("ActualProgress"))
    Dashboard_ProgressDeltaText = TextCatalog_Format("DASHBOARD.PROGRESS.DELTA", Dashboard_CurrentLanguage(), TextCatalog_Arguments("Delta", Dashboard_FormatPercentSigned(deltaVal)))

End Function

'------------------------------------------------------------------------------
' FR: Retourne la valeur Behind Plan Text sans modifier les donnees d'entree.
' EN: Returns the Behind Plan Text value without mutating input data.
'------------------------------------------------------------------------------

Private Function Dashboard_BehindPlanText(ByVal plannedProgress As Double, ByVal actualProgress As Double, ByRef statusColor As Long) As String

    Dim behindPlan As Double

    behindPlan = plannedProgress - actualProgress

    If behindPlan <= 0# Then
        statusColor = RGB(0, 145, 112)
        Dashboard_BehindPlanText = TextCatalog_Get("DASHBOARD.STATUS.AHEAD", Dashboard_CurrentLanguage())
    ElseIf behindPlan <= 0.1 Then
        statusColor = RGB(238, 156, 68)
        Dashboard_BehindPlanText = TextCatalog_Format("DASHBOARD.PROGRESS.BEHIND", Dashboard_CurrentLanguage(), TextCatalog_Arguments("Progress", Format$(behindPlan, "0%")))
    Else
        statusColor = RGB(192, 80, 77)
        Dashboard_BehindPlanText = TextCatalog_Format("DASHBOARD.PROGRESS.BEHIND", Dashboard_CurrentLanguage(), TextCatalog_Arguments("Progress", Format$(behindPlan, "0%")))
    End If

End Function

'------------------------------------------------------------------------------
' FR: Retourne la map Forecast Delta Text sans modifier les donnees d'entree.
' EN: Returns the Forecast Delta Text map without mutating input data.
'------------------------------------------------------------------------------

Private Function Dashboard_ForecastDeltaText(ByVal fromRow As Object, ByVal toRow As Object) As String

    Dim deltaDays As Long

    If fromRow Is Nothing Or toRow Is Nothing Then
        Dashboard_ForecastDeltaText = TextCatalog_Get("DASHBOARD.EMPTY.COMPARISON", Dashboard_CurrentLanguage())
        Exit Function
    End If
    If Not Dashboard_HasDateValue(fromRow("ForecastFinish")) Or Not Dashboard_HasDateValue(toRow("ForecastFinish")) Then
        Dashboard_ForecastDeltaText = TextCatalog_Get("DASHBOARD.EMPTY.FORECAST_COMPARISON", Dashboard_CurrentLanguage())
        Exit Function
    End If

    deltaDays = CLng(Dashboard_DateNumber(toRow("ForecastFinish")) - Dashboard_DateNumber(fromRow("ForecastFinish")))
    If deltaDays < 0 Then
        Dashboard_ForecastDeltaText = TextCatalog_Format("DASHBOARD.FORECAST.IMPROVED", Dashboard_CurrentLanguage(), TextCatalog_Arguments("Count", CStr(Abs(deltaDays))))
    ElseIf deltaDays > 0 Then
        Dashboard_ForecastDeltaText = TextCatalog_Format("DASHBOARD.FORECAST.SLIPPED", Dashboard_CurrentLanguage(), TextCatalog_Arguments("Count", CStr(deltaDays)))
    Else
        Dashboard_ForecastDeltaText = TextCatalog_Get("DASHBOARD.STATUS.FORECAST_STABLE", Dashboard_CurrentLanguage())
    End If

End Function

'------------------------------------------------------------------------------
' FR: Retourne la valeur Contract Drift Text sans modifier les donnees d'entree.
' EN: Returns the Contract Drift Text value without mutating input data.
'------------------------------------------------------------------------------

Private Function Dashboard_ContractDriftText(ByVal driftDays As Variant, ByRef statusColor As Long) As String

    If Not HasValue(driftDays) Then
        statusColor = RGB(96, 111, 128)
        Dashboard_ContractDriftText = TextCatalog_Get("DASHBOARD.EMPTY.CONTRACT_DRIFT", Dashboard_CurrentLanguage())
    ElseIf CLng(driftDays) <= 0 Then
        statusColor = RGB(0, 145, 112)
        Dashboard_ContractDriftText = TextCatalog_Format("DASHBOARD.CONTRACT.DRIFT", Dashboard_CurrentLanguage(), TextCatalog_Arguments("Days", CStr(CLng(driftDays))))
    ElseIf CLng(driftDays) <= 30 Then
        statusColor = RGB(238, 156, 68)
        Dashboard_ContractDriftText = TextCatalog_Format("DASHBOARD.CONTRACT.DRIFT", Dashboard_CurrentLanguage(), TextCatalog_Arguments("Days", "+" & CStr(CLng(driftDays))))
    Else
        statusColor = RGB(192, 80, 77)
        Dashboard_ContractDriftText = TextCatalog_Format("DASHBOARD.CONTRACT.DRIFT", Dashboard_CurrentLanguage(), TextCatalog_Arguments("Days", "+" & CStr(CLng(driftDays))))
    End If

End Function

'------------------------------------------------------------------------------
' FR: Retourne la map Critical Activities Hero sans modifier les donnees d'entree.
' EN: Returns the Critical Activities Hero map without mutating input data.
'------------------------------------------------------------------------------

Private Function Dashboard_CriticalActivitiesHero(ByVal fromRow As Object, ByVal toRow As Object, ByRef statusColor As Long, ByRef subText As String) As String

    Dim fromCount As Long
    Dim toCount As Long
    Dim deltaVal As Long

    statusColor = RGB(96, 111, 128)
    subText = TextCatalog_Get("DASHBOARD.EMPTY.HISTORY_SENTENCE", Dashboard_CurrentLanguage())
    Dashboard_CriticalActivitiesHero = "-"

    If toRow Is Nothing Then Exit Function

    toCount = CLng(toRow("CriticalPathCount"))
    If fromRow Is Nothing Then
        Dashboard_CriticalActivitiesHero = CStr(toCount) & " " & ChrW(8594)
        Exit Function
    End If

    fromCount = CLng(fromRow("CriticalPathCount"))
    deltaVal = toCount - fromCount

    If deltaVal > 0 Then
        statusColor = RGB(192, 80, 77)
        Dashboard_CriticalActivitiesHero = CStr(toCount) & " " & ChrW(8599)
        subText = TextCatalog_Format("DASHBOARD.CRITICAL.ADDED", Dashboard_CurrentLanguage(), TextCatalog_Arguments("Count", CStr(deltaVal)))
    ElseIf deltaVal < 0 Then
        statusColor = RGB(0, 145, 112)
        Dashboard_CriticalActivitiesHero = CStr(toCount) & " " & ChrW(8600)
        subText = TextCatalog_Format("DASHBOARD.CRITICAL.REMOVED", Dashboard_CurrentLanguage(), TextCatalog_Arguments("Count", CStr(deltaVal)))
    Else
        statusColor = RGB(238, 156, 68)
        Dashboard_CriticalActivitiesHero = CStr(toCount) & " " & ChrW(8594)
        subText = TextCatalog_Get("DASHBOARD.STATUS.CRITICAL_UNCHANGED", Dashboard_CurrentLanguage())
    End If

End Function

'------------------------------------------------------------------------------
' FR: Retourne la map Schedule Momentum Status sans modifier les donnees d'entree.
' EN: Returns the Schedule Momentum Status map without mutating input data.
'------------------------------------------------------------------------------

Private Function Dashboard_ScheduleMomentumStatus(ByVal fromRow As Object, ByVal toRow As Object, ByRef statusColor As Long, ByRef subText As String) As String

    Dim delayFrom As Double
    Dim delayTo As Double
    Dim delayDelta As Double

    statusColor = RGB(96, 111, 128)
    subText = TextCatalog_Get("DASHBOARD.EMPTY.SNAPSHOTS_REQUIRED", Dashboard_CurrentLanguage())
    Dashboard_ScheduleMomentumStatus = "INSUFFICIENT HISTORY"

    If fromRow Is Nothing Or toRow Is Nothing Then Exit Function
    If CStr(fromRow("SnapshotId")) = CStr(toRow("SnapshotId")) Then Exit Function

    delayFrom = CDbl(fromRow("PlannedProgress")) - CDbl(fromRow("ActualProgress"))
    delayTo = CDbl(toRow("PlannedProgress")) - CDbl(toRow("ActualProgress"))
    delayDelta = delayTo - delayFrom

    If delayDelta < 0 Then
        statusColor = RGB(0, 145, 112)
        Dashboard_ScheduleMomentumStatus = "IMPROVING"
        subText = TextCatalog_Format("DASHBOARD.MOMENTUM.DELAY", Dashboard_CurrentLanguage(), TextCatalog_Arguments("Delta", Dashboard_FormatMomentumDelayPercent(delayDelta)))
    ElseIf delayDelta > 0 Then
        statusColor = RGB(192, 80, 77)
        Dashboard_ScheduleMomentumStatus = "DETERIORATING"
        subText = TextCatalog_Format("DASHBOARD.MOMENTUM.DELAY", Dashboard_CurrentLanguage(), TextCatalog_Arguments("Delta", Dashboard_FormatMomentumDelayPercent(delayDelta)))
    Else
        statusColor = RGB(238, 156, 68)
        Dashboard_ScheduleMomentumStatus = "STABLE"
        subText = TextCatalog_Get("DASHBOARD.STATUS.DELAY_UNCHANGED", Dashboard_CurrentLanguage())
    End If

End Function

'------------------------------------------------------------------------------
' FR: Retourne la valeur Local Contract Status sans modifier les donnees d'entree.
' EN: Returns the Local Contract Status value without mutating input data.
'------------------------------------------------------------------------------

Private Function Dashboard_LocalContractStatus(ByVal statusText As String) As String

    Select Case UCase$(Trim$(statusText))
        Case "ON CONTRACT": Dashboard_LocalContractStatus = TextCatalog_Get("DASHBOARD.STATUS.ON_CONTRACT", Dashboard_CurrentLanguage())
        Case "MINOR DELAY": Dashboard_LocalContractStatus = TextCatalog_Get("DASHBOARD.STATUS.MINOR_DELAY", Dashboard_CurrentLanguage())
        Case "CONTRACT DELAY": Dashboard_LocalContractStatus = TextCatalog_Get("DASHBOARD.STATUS.CONTRACT_DELAY", Dashboard_CurrentLanguage())
        Case Else: Dashboard_LocalContractStatus = statusText
    End Select

End Function

'------------------------------------------------------------------------------
' FR: Retourne la valeur Local Momentum Status sans modifier les donnees d'entree.
' EN: Returns the Local Momentum Status value without mutating input data.
'------------------------------------------------------------------------------

Private Function Dashboard_LocalMomentumStatus(ByVal statusText As String) As String

    Select Case UCase$(Trim$(statusText))
        Case "IMPROVING": Dashboard_LocalMomentumStatus = TextCatalog_Get("DASHBOARD.STATUS.IMPROVING", Dashboard_CurrentLanguage())
        Case "DETERIORATING": Dashboard_LocalMomentumStatus = TextCatalog_Get("DASHBOARD.STATUS.DETERIORATING", Dashboard_CurrentLanguage())
        Case "STABLE": Dashboard_LocalMomentumStatus = TextCatalog_Get("DASHBOARD.STATUS.STABLE", Dashboard_CurrentLanguage())
        Case "INSUFFICIENT HISTORY": Dashboard_LocalMomentumStatus = TextCatalog_Get("DASHBOARD.STATUS.NO_HISTORY", Dashboard_CurrentLanguage())
        Case Else: Dashboard_LocalMomentumStatus = statusText
    End Select

End Function

'------------------------------------------------------------------------------
' FR: Retourne la map Array Val sans modifier les donnees d'entree.
' EN: Returns the Array Val map without mutating input data.
'------------------------------------------------------------------------------

Private Function Dashboard_ArrayVal(ByRef arr As Variant, ByVal mapTbl As Object, ByVal r As Long, ByVal colName As String) As Variant

    If mapTbl.exists(colName) Then Dashboard_ArrayVal = arr(r, mapTbl(colName))

End Function

'------------------------------------------------------------------------------
' FR: Indique si la map Summary Array Row satisfait la condition attendue, sans modifier les donnees source.
' EN: Returns whether the Summary Array Row map satisfies the expected condition without mutating source data.
'------------------------------------------------------------------------------

Private Function Dashboard_IsSummaryArrayRow(ByRef arr As Variant, ByVal mapTbl As Object, ByVal r As Long) As Boolean

    Dim v As Variant
    Dim txt As String

    If Not mapTbl.exists("IsSummary") Then Exit Function
    v = arr(r, mapTbl("IsSummary"))

    If VarType(v) = vbBoolean Then
        Dashboard_IsSummaryArrayRow = CBool(v)
    ElseIf IsNumeric(v) Then
        Dashboard_IsSummaryArrayRow = (CDbl(v) <> 0)
    Else
        txt = UCase$(Trim$(CStr(v)))
        Dashboard_IsSummaryArrayRow = (txt = "TRUE" Or txt = "YES" Or txt = "SUMMARY" Or txt = "1")
    End If

End Function

'------------------------------------------------------------------------------
' FR: Construit l'index Progress By ID a partir des donnees fournies par l'appelant.
' EN: Builds the Progress By ID index from data supplied by the caller.
'------------------------------------------------------------------------------

Private Function Dashboard_BuildProgressById(ByVal tblWBS As ListObject, ByVal mapWBS As Object) As Object

    Dim progressById As Object
    Dim arr As Variant
    Dim r As Long
    Dim idVal As String

    Set progressById = CreateObject("Scripting.Dictionary")

    If tblWBS Is Nothing Then GoTo CleanExit
    If tblWBS.DataBodyRange Is Nothing Then GoTo CleanExit
    If Not mapWBS.exists(VTS_COL_ID) Then GoTo CleanExit
    If Not mapWBS.exists(VTS_COL_PROGRESS_PERCENT) Then GoTo CleanExit

    arr = tblWBS.DataBodyRange.value
    For r = 1 To UBound(arr, 1)
        idVal = CStr(GetCellValue(arr(r, mapWBS(VTS_COL_ID))))
        If idVal <> "" Then progressById(idVal) = Dashboard_NormalizeProgress(GetCellValue(arr(r, mapWBS(VTS_COL_PROGRESS_PERCENT))))
    Next r

CleanExit:
    Set Dashboard_BuildProgressById = progressById

End Function

'------------------------------------------------------------------------------
' FR: Retourne la map Summary Progress sans modifier les donnees d'entree.
' EN: Returns the Summary Progress map without mutating input data.
'------------------------------------------------------------------------------

Private Function Dashboard_BuildSummaryProgress(ByRef arr As Variant, ByVal mapTbl As Object, ByVal progressById As Object, Optional ByVal visibleParents As Object = Nothing) As Object
    Dim result As Object
    Dim eventsByParent As Object
    Dim events As Object
    Dim r As Long
    Dim parentWbs As String
    Dim taskId As String
    Dim startVal As Variant
    Dim finishVal As Variant
    Dim progressVal As Double
    Dim leafColor As Variant
    Dim completeColor As Long
    Dim metrics As Variant
    Dim coveredThrough As Variant
    Dim key As Variant
    Dim day As Long, i As Long
    Dim eventDays() As Long
    Dim keys As Variant
    Dim intervalStart As Long, intervalFinish As Long
    Dim active As Long, covered As Long
    Dim delta As Variant

    Set result = CreateObject("Scripting.Dictionary")
    Set eventsByParent = CreateObject("Scripting.Dictionary")
    Set Dashboard_BuildSummaryProgress = result
    If Not mapTbl.exists("WBS") Then Exit Function

    ' Parent geometry comes from CALC; no scheduling or inferred dates here.
    For r = 1 To UBound(arr, 1)
        If Not Dashboard_IsSummaryArrayRow(arr, mapTbl, r) Then GoTo NextParent
        parentWbs = NormalizeWBS(CStr(Dashboard_ArrayVal(arr, mapTbl, r, "WBS")))
        If Not visibleParents Is Nothing Then
            If Not visibleParents.exists(parentWbs) Then GoTo NextParent
        End If
        startVal = GetCellValue(Dashboard_ArrayVal(arr, mapTbl, r, "Calculated Start"))
        finishVal = GetCellValue(Dashboard_ArrayVal(arr, mapTbl, r, "Calculated Finish"))
        If Not Dashboard_HasDateValue(startVal) Or Not Dashboard_HasDateValue(finishVal) Then GoTo NextParent
        startVal = CLng(Dashboard_DateNumber(startVal))
        finishVal = CLng(Dashboard_DateNumber(finishVal))
        If finishVal < startVal Then GoTo NextParent
        completeColor = CLng(GetProgressFillColor(startVal, finishVal, 1#))
        result(parentWbs) = Array(CDbl(finishVal - startVal + 1), 0#, completeColor, startVal, finishVal, 0&)
        Set events = CreateObject("Scripting.Dictionary")
        Set eventsByParent(parentWbs) = events
NextParent:
    Next r

    ' Sparse daily deltas: each terminal task counts once per ancestor, with equal weight on each day.
    For r = 1 To UBound(arr, 1)
        If Dashboard_IsSummaryArrayRow(arr, mapTbl, r) Then GoTo NextTask
        taskId = CStr(GetCellValue(Dashboard_ArrayVal(arr, mapTbl, r, "ID")))
        If taskId = "" Then GoTo NextTask
        startVal = GetCellValue(Dashboard_ArrayVal(arr, mapTbl, r, "Calculated Start"))
        finishVal = GetCellValue(Dashboard_ArrayVal(arr, mapTbl, r, "Calculated Finish"))
        If Not Dashboard_HasDateValue(startVal) Or Not Dashboard_HasDateValue(finishVal) Then GoTo NextTask
        startVal = CLng(Dashboard_DateNumber(startVal))
        finishVal = CLng(Dashboard_DateNumber(finishVal))
        If CDbl(finishVal) < CDbl(startVal) Then GoTo NextTask
        progressVal = Dashboard_RowProgress(arr, mapTbl, progressById, r, taskId)
        leafColor = GetProgressFillColor(startVal, finishVal, progressVal, coveredThrough)
        If TaskTypeRules_IsMilestoneRow(arr, mapTbl, r) Then leafColor = Empty
        completeColor = CLng(GetProgressFillColor(startVal, finishVal, 1#))
        parentWbs = GetParentWBS(NormalizeWBS(CStr(Dashboard_ArrayVal(arr, mapTbl, r, "WBS"))))
        Do While parentWbs <> ""
            If result.exists(parentWbs) Then
                metrics = result(parentWbs)
                metrics(5) = CLng(metrics(5)) + 1
                Set events = eventsByParent(parentWbs)
                Dashboard_AddDailyProgressDelta events, CLng(startVal), 1, 0
                Dashboard_AddDailyProgressDelta events, CLng(finishVal) + 1, -1, 0
                If Not IsEmpty(coveredThrough) Then
                    Dashboard_AddDailyProgressDelta events, CLng(startVal), 0, 1
                    Dashboard_AddDailyProgressDelta events, CLng(coveredThrough) + 1, 0, -1
                End If
                If Not IsEmpty(leafColor) Then
                    If CLng(leafColor) <> completeColor Then metrics(2) = CLng(leafColor)
                End If
                result(parentWbs) = metrics
            End If
            parentWbs = GetParentWBS(parentWbs)
        Loop
NextTask:
    Next r

    ' Counts are constant between event dates; integrate daily fractions over those intervals.
    For Each key In result.keys
        metrics = result(key)
        If CLng(metrics(5)) = 0 Then
            result.Remove key
            GoTo NextSummary
        End If
        Set events = eventsByParent(key)
        Dashboard_AddDailyProgressDelta events, CLng(metrics(3)), 0, 0
        Dashboard_AddDailyProgressDelta events, CLng(metrics(4)) + 1, 0, 0
        If CLng(Date) > CLng(metrics(3)) And CLng(Date) <= CLng(metrics(4)) Then
            Dashboard_AddDailyProgressDelta events, CLng(Date), 0, 0
        End If
        keys = events.keys
        ReDim eventDays(0 To events.count - 1)
        For i = 0 To UBound(keys)
            eventDays(i) = CLng(keys(i))
        Next i
        Dashboard_SortProgressDays eventDays, LBound(eventDays), UBound(eventDays)
        active = 0: covered = 0
        For i = LBound(eventDays) To UBound(eventDays) - 1
            day = eventDays(i)
            delta = events(CStr(day))
            active = active + CLng(delta(0))
            covered = covered + CLng(delta(1))
            intervalStart = day: intervalFinish = eventDays(i + 1)
            If intervalStart < CLng(metrics(3)) Then intervalStart = CLng(metrics(3))
            If intervalFinish > CLng(metrics(4)) + 1 Then intervalFinish = CLng(metrics(4)) + 1
            If intervalFinish > intervalStart Then
                If active > 0 Then
                    metrics(1) = metrics(1) + (intervalFinish - intervalStart) * CDbl(covered) / active
                ElseIf intervalStart < CLng(Date) Then
                    metrics(1) = metrics(1) + intervalFinish - intervalStart
                End If
            End If
        Next i
        result(key) = metrics
NextSummary:
    Next key
End Function

Private Sub Dashboard_AddDailyProgressDelta(ByVal events As Object, ByVal day As Long, ByVal active As Long, ByVal covered As Long)
    Dim delta As Variant
    If events.exists(CStr(day)) Then
        delta = events(CStr(day))
    Else
        delta = Array(0&, 0&)
    End If
    delta(0) = CLng(delta(0)) + active
    delta(1) = CLng(delta(1)) + covered
    events(CStr(day)) = delta
End Sub

Private Sub Dashboard_SortProgressDays(ByRef days() As Long, ByVal first As Long, ByVal last As Long)
    Dim left As Long, right As Long, pivot As Long, swap As Long
    left = first: right = last: pivot = days((first + last) \ 2)
    Do While left <= right
        Do While days(left) < pivot: left = left + 1: Loop
        Do While days(right) > pivot: right = right - 1: Loop
        If left <= right Then
            swap = days(left): days(left) = days(right): days(right) = swap
            left = left + 1: right = right - 1
        End If
    Loop
    If first < right Then Dashboard_SortProgressDays days, first, right
    If left < last Then Dashboard_SortProgressDays days, left, last
End Sub

'------------------------------------------------------------------------------
' FR: Retourne la map Row Progress sans modifier les donnees d'entree.
' EN: Returns the Row Progress map without mutating input data.
'------------------------------------------------------------------------------

Private Function Dashboard_RowProgress(ByRef arr As Variant, ByVal mapTbl As Object, ByVal progressById As Object, ByVal r As Long, ByVal idVal As String) As Double

    Dim actualFinish As Variant

    If TaskTypeRules_IsLevelOfEffortRow(arr, mapTbl, r) Then
        Dashboard_RowProgress = GetLoEProgressFromToday( _
            GetCellValue(Dashboard_ArrayVal(arr, mapTbl, r, "Calculated Start")), _
            GetCellValue(Dashboard_ArrayVal(arr, mapTbl, r, "Calculated Finish")))
        Exit Function
    End If

    If Not progressById Is Nothing Then
        If idVal <> "" And progressById.exists(idVal) Then
            Dashboard_RowProgress = CDbl(progressById(idVal))
            Exit Function
        End If
    End If

    If mapTbl.exists("% Progress") Then
        Dashboard_RowProgress = Dashboard_NormalizeProgress(GetCellValue(Dashboard_ArrayVal(arr, mapTbl, r, "% Progress")))
    ElseIf mapTbl.exists("Actual Finish") Then
        actualFinish = GetCellValue(Dashboard_ArrayVal(arr, mapTbl, r, "Actual Finish"))
        If Dashboard_HasDateValue(actualFinish) Then Dashboard_RowProgress = 1#
    End If

End Function

'------------------------------------------------------------------------------
' FR: Normalise ou formate Normalize Progress selon le contrat canonique du composant.
' EN: Normalizes or formats Normalize Progress according to the component contract.
'------------------------------------------------------------------------------

Private Function Dashboard_NormalizeProgress(ByVal progressVal As Variant) As Double

    If Not HasValue(progressVal) Or Not IsNumeric(progressVal) Then Exit Function

    Dashboard_NormalizeProgress = CDbl(progressVal)
    If Dashboard_NormalizeProgress > 1# Then Dashboard_NormalizeProgress = Dashboard_NormalizeProgress / 100#
    If Dashboard_NormalizeProgress < 0# Then Dashboard_NormalizeProgress = 0#
    If Dashboard_NormalizeProgress > 1# Then Dashboard_NormalizeProgress = 1#

End Function

'------------------------------------------------------------------------------
' FR: Retourne la valeur Wbs Level sans modifier les donnees d'entree.
' EN: Returns the Wbs Level value without mutating input data.
'------------------------------------------------------------------------------

Private Function Dashboard_WbsLevel(ByVal wbsValue As String) As Long

    Dim txt As String
    Dim i As Long

    txt = Replace$(Trim$(wbsValue), ",", ".")
    If txt = "" Then Exit Function

    Dashboard_WbsLevel = 1
    For i = 1 To Len(txt)
        If Mid$(txt, i, 1) = "." Then Dashboard_WbsLevel = Dashboard_WbsLevel + 1
    Next i

End Function

'------------------------------------------------------------------------------
' FR: Retourne la valeur Date To X sans modifier les donnees d'entree.
' EN: Returns the Date To X value without mutating input data.
'------------------------------------------------------------------------------

Private Function Dashboard_DateToX(ByVal dateVal As Double, ByVal minDate As Double, ByVal maxDate As Double, ByVal leftPos As Double, ByVal widthVal As Double) As Double

    If maxDate <= minDate Then
        Dashboard_DateToX = leftPos
    Else
        Dashboard_DateToX = leftPos + (((dateVal - minDate) / (maxDate - minDate)) * widthVal)
    End If

End Function

'------------------------------------------------------------------------------
' FR: Retourne la collection Planning Overview Rows sans modifier les donnees d'entree.
' EN: Returns the Planning Overview Rows collection without mutating input data.
'------------------------------------------------------------------------------

Private Function Dashboard_PlanningOverviewRows(ByRef arr As Variant, ByVal mapTbl As Object, ByRef minOut As Double, ByRef maxOut As Double) As Collection

    Const MAX_OVERVIEW_ROWS As Long = 10

    Dim activeRows As Collection
    Dim beforeRows As Collection
    Dim afterRows As Collection
    Dim selected As Collection
    Dim selectedRows As Object
    Dim r As Long
    Dim rowStart As Variant
    Dim rowFinish As Variant
    Dim startSerial As Double
    Dim finishSerial As Double
    Dim todaySerial As Double
    Dim remainingSlots As Long
    Dim beforeQuota As Long
    Dim afterQuota As Long
    Dim hasDate As Boolean
    Dim orderedRows As Collection

    Set activeRows = New Collection
    Set beforeRows = New Collection
    Set afterRows = New Collection
    Set selected = New Collection
    Set selectedRows = CreateObject("Scripting.Dictionary")

    todaySerial = CDbl(Date)

    For r = 1 To UBound(arr, 1)
        If Dashboard_IsPlanningOverviewVisibleRow(arr, mapTbl, r) Then
            If Dashboard_PlanningRowHasRenderableDates(arr, mapTbl, r) Then
                rowStart = GetCellValue(Dashboard_ArrayVal(arr, mapTbl, r, "Calculated Start"))
                rowFinish = GetCellValue(Dashboard_ArrayVal(arr, mapTbl, r, "Calculated Finish"))
                startSerial = Dashboard_DateNumber(rowStart)
                finishSerial = Dashboard_DateNumber(rowFinish)

                If startSerial <= todaySerial And finishSerial >= todaySerial Then
                    activeRows.Add r
                ElseIf finishSerial < todaySerial Then
                    Dashboard_InsertPlanningCandidateByDistance beforeRows, r, todaySerial - finishSerial
                ElseIf startSerial > todaySerial Then
                    Dashboard_InsertPlanningCandidateByDistance afterRows, r, startSerial - todaySerial
                End If
            End If
        End If
    Next r

    Dashboard_AddPlanningCandidates selected, selectedRows, activeRows, MAX_OVERVIEW_ROWS

    remainingSlots = MAX_OVERVIEW_ROWS - selected.count
    If remainingSlots > 0 Then
        beforeQuota = remainingSlots \ 2
        afterQuota = remainingSlots - beforeQuota

        Dashboard_AddPlanningCandidates selected, selectedRows, beforeRows, beforeQuota
        Dashboard_AddPlanningCandidates selected, selectedRows, afterRows, afterQuota

        If selected.count < MAX_OVERVIEW_ROWS Then
            Dashboard_AddPlanningCandidates selected, selectedRows, beforeRows, MAX_OVERVIEW_ROWS - selected.count
        End If
        If selected.count < MAX_OVERVIEW_ROWS Then
            Dashboard_AddPlanningCandidates selected, selectedRows, afterRows, MAX_OVERVIEW_ROWS - selected.count
        End If
    End If

    Set orderedRows = Dashboard_OrderPlanningRowsBySourceOrder(selected)

    For r = 1 To orderedRows.count
        Dashboard_IncludeAxisDate arr, mapTbl, CLng(orderedRows(r)), "Baseline Start", hasDate, minOut, maxOut
        Dashboard_IncludeAxisDate arr, mapTbl, CLng(orderedRows(r)), "Baseline Finish", hasDate, minOut, maxOut
        Dashboard_IncludeAxisDate arr, mapTbl, CLng(orderedRows(r)), "Calculated Start", hasDate, minOut, maxOut
        Dashboard_IncludeAxisDate arr, mapTbl, CLng(orderedRows(r)), "Calculated Finish", hasDate, minOut, maxOut
    Next r

    Set Dashboard_PlanningOverviewRows = orderedRows

End Function

'------------------------------------------------------------------------------
' FR: Traite la collection Insert Planning Candidate By Distance sans modifier les donnees d'entree.
' EN: Handles the Insert Planning Candidate By Distance collection without mutating input data.
'------------------------------------------------------------------------------

Private Sub Dashboard_InsertPlanningCandidateByDistance(ByVal rows As Collection, ByVal rowIndex As Long, ByVal distanceVal As Double)

    Dim item As Object
    Dim existing As Object
    Dim tmp As Collection
    Dim inserted As Boolean
    Dim i As Long

    Set item = CreateObject("Scripting.Dictionary")
    item("Row") = rowIndex
    item("Distance") = distanceVal

    Set tmp = New Collection
    For i = 1 To rows.count
        Set existing = rows(i)
        If Not inserted Then
            If distanceVal < CDbl(existing("Distance")) Then
                tmp.Add item
                inserted = True
            End If
        End If
        tmp.Add existing
    Next i

    If Not inserted Then tmp.Add item

    Do While rows.count > 0
        rows.Remove 1
    Loop
    For i = 1 To tmp.count
        rows.Add tmp(i)
    Next i

End Sub

'------------------------------------------------------------------------------
' FR: Ajoute la collection Planning Candidates a la structure cible fournie par l'appelant.
' EN: Adds the Planning Candidates collection to the target structure supplied by the caller.
'------------------------------------------------------------------------------

Private Sub Dashboard_AddPlanningCandidates( _
    ByVal selected As Collection, _
    ByVal selectedRows As Object, _
    ByVal candidates As Collection, _
    ByVal maxToAdd As Long)

    Dim item As Variant
    Dim rowIndex As Long
    Dim added As Long

    If maxToAdd <= 0 Then Exit Sub

    For Each item In candidates
        rowIndex = Dashboard_PlanningCandidateRow(item)
        If rowIndex > 0 Then
            If Not selectedRows.exists(CStr(rowIndex)) Then
                selectedRows(CStr(rowIndex)) = True
                selected.Add rowIndex
                added = added + 1
                If added >= maxToAdd Then Exit Sub
            End If
        End If
    Next item

End Sub

'------------------------------------------------------------------------------
' FR: Retourne la valeur Planning Candidate Row sans modifier les donnees d'entree.
' EN: Returns the Planning Candidate Row value without mutating input data.
'------------------------------------------------------------------------------

Private Function Dashboard_PlanningCandidateRow(ByVal item As Variant) As Long

    If IsObject(item) Then
        Dashboard_PlanningCandidateRow = CLng(item("Row"))
    Else
        Dashboard_PlanningCandidateRow = CLng(item)
    End If

End Function

'------------------------------------------------------------------------------
' FR: Retourne la collection Order Planning Rows By Source Order sans modifier les donnees d'entree.
' EN: Returns the Order Planning Rows By Source Order collection without mutating input data.
'------------------------------------------------------------------------------

Private Function Dashboard_OrderPlanningRowsBySourceOrder(ByVal rows As Collection) As Collection

    Dim ordered As Collection
    Dim i As Long
    Dim j As Long
    Dim currentRow As Long
    Dim inserted As Boolean

    Set ordered = New Collection

    For i = 1 To rows.count
        currentRow = CLng(rows(i))
        inserted = False

        For j = 1 To ordered.count
            If currentRow < CLng(ordered(j)) Then
                ordered.Add currentRow, , j
                inserted = True
                Exit For
            End If
        Next j

        If Not inserted Then ordered.Add currentRow
    Next i

    Set Dashboard_OrderPlanningRowsBySourceOrder = ordered

End Function
'------------------------------------------------------------------------------
' FR: Indique si la map Planning Overview Visible Row satisfait la condition attendue, sans modifier les donnees source.
' EN: Returns whether the Planning Overview Visible Row map satisfies the expected condition without mutating source data.
'------------------------------------------------------------------------------

Private Function Dashboard_IsPlanningOverviewVisibleRow(ByRef arr As Variant, ByVal mapTbl As Object, ByVal r As Long) As Boolean

    Dim durationVal As Variant
    Dim summaryDisplayVal As String

    If mapTbl.exists("S") Then
        summaryDisplayVal = UCase$(Trim$(CStr(Dashboard_ArrayVal(arr, mapTbl, r, "S"))))
        Dashboard_IsPlanningOverviewVisibleRow = (summaryDisplayVal = "Y")
        Exit Function
    End If

    If Dashboard_IsSummaryArrayRow(arr, mapTbl, r) Then
        Dashboard_IsPlanningOverviewVisibleRow = True
        Exit Function
    End If

    If mapTbl.exists("Calculated Duration") Then
        durationVal = GetCellValue(Dashboard_ArrayVal(arr, mapTbl, r, "Calculated Duration"))
        If IsNumeric(durationVal) Then
            Dashboard_IsPlanningOverviewVisibleRow = (CDbl(durationVal) <= 1#)
        End If
    End If

End Function

'------------------------------------------------------------------------------
' FR: Retourne la map Planning Row Has Renderable Dates sans modifier les donnees d'entree.
' EN: Returns the Planning Row Has Renderable Dates map without mutating input data.
'------------------------------------------------------------------------------

Private Function Dashboard_PlanningRowHasRenderableDates(ByRef arr As Variant, ByVal mapTbl As Object, ByVal r As Long) As Boolean

    Dim startVal As Variant
    Dim finishVal As Variant

    startVal = GetCellValue(Dashboard_ArrayVal(arr, mapTbl, r, "Calculated Start"))
    finishVal = GetCellValue(Dashboard_ArrayVal(arr, mapTbl, r, "Calculated Finish"))
    Dashboard_PlanningRowHasRenderableDates = (Dashboard_HasDateValue(startVal) And Dashboard_HasDateValue(finishVal))

End Function

'------------------------------------------------------------------------------
' FR: Retourne la valeur Planning Axis Margin sans modifier les donnees d'entree.
' EN: Returns the Planning Axis Margin value without mutating input data.
'------------------------------------------------------------------------------

Private Function Dashboard_PlanningAxisMargin(ByVal minDate As Double, ByVal maxDate As Double) As Double

    Dim spanDays As Double

    spanDays = maxDate - minDate
    If spanDays <= 14 Then
        Dashboard_PlanningAxisMargin = 1
    ElseIf spanDays <= 62 Then
        Dashboard_PlanningAxisMargin = 3
    ElseIf spanDays <= 730 Then
        Dashboard_PlanningAxisMargin = 14
    Else
        Dashboard_PlanningAxisMargin = 45
    End If

End Function

'------------------------------------------------------------------------------
' FR: Retourne la map Planning Axis Bounds sans exposer de mutateur sur l'etat source.
' EN: Returns the Planning Axis Bounds map without exposing a mutator for source state.
'------------------------------------------------------------------------------

Private Sub Dashboard_GetPlanningAxisBounds(ByRef arr As Variant, ByVal mapTbl As Object, ByRef minOut As Double, ByRef maxOut As Double)

    Dim r As Long
    Dim shown As Long
    Dim v As Variant
    Dim hasDate As Boolean

    For r = 1 To UBound(arr, 1)
        If Dashboard_IsSummaryArrayRow(arr, mapTbl, r) Then
            If Dashboard_WbsLevel(CStr(Dashboard_ArrayVal(arr, mapTbl, r, "WBS"))) <= 3 Then
                shown = shown + 1
                If shown > 7 Then Exit For
                Dashboard_IncludeAxisDate arr, mapTbl, r, "Baseline Start", hasDate, minOut, maxOut
                Dashboard_IncludeAxisDate arr, mapTbl, r, "Baseline Finish", hasDate, minOut, maxOut
                Dashboard_IncludeAxisDate arr, mapTbl, r, "Calculated Start", hasDate, minOut, maxOut
                Dashboard_IncludeAxisDate arr, mapTbl, r, "Calculated Finish", hasDate, minOut, maxOut
            End If
        End If
    Next r

    If hasDate Then
        minOut = minOut - 7
        maxOut = maxOut + 10
    End If

End Sub

'------------------------------------------------------------------------------
' FR: Traite la map Include Axis Date sans modifier les donnees d'entree.
' EN: Handles the Include Axis Date map without mutating input data.
'------------------------------------------------------------------------------

Private Sub Dashboard_IncludeAxisDate(ByRef arr As Variant, ByVal mapTbl As Object, ByVal r As Long, ByVal colName As String, ByRef hasDate As Boolean, ByRef minOut As Double, ByRef maxOut As Double)

    Dim v As Variant
    Dim d As Double

    v = GetCellValue(Dashboard_ArrayVal(arr, mapTbl, r, colName))
    If Not Dashboard_HasDateValue(v) Then Exit Sub

    d = Dashboard_DateNumber(v)
    If Not hasDate Then
        minOut = d
        maxOut = d
        hasDate = True
    ElseIf d < minOut Then
        minOut = d
    ElseIf d > maxOut Then
        maxOut = d
    End If

End Sub

'------------------------------------------------------------------------------
' FR: Insere un candidat dans le classement borne des plus forts retards du Dashboard.
' EN: Inserts a candidate into the Dashboard's bounded highest-delay ranking.
'------------------------------------------------------------------------------

Private Sub Dashboard_InsertTopDelay( _
    ByRef topIds() As String, _
    ByRef topWbs() As String, _
    ByRef topNames() As String, _
    ByRef topVals() As Double, _
    ByVal idVal As String, _
    ByVal wbsVal As String, _
    ByVal nameVal As String, _
    ByVal metricVal As Double)

    Dim i As Long
    Dim j As Long

    For i = 1 To 5
        If metricVal > topVals(i) Then
            For j = 5 To i + 1 Step -1
                topIds(j) = topIds(j - 1)
                topWbs(j) = topWbs(j - 1)
                topNames(j) = topNames(j - 1)
                topVals(j) = topVals(j - 1)
            Next j
            topIds(i) = idVal
            topWbs(i) = wbsVal
            topNames(i) = nameVal
            topVals(i) = metricVal
            Exit Sub
        End If
    Next i

End Sub

'------------------------------------------------------------------------------
' FR: Insere un candidat dans le classement borne des risques les plus faibles du Dashboard.
' EN: Inserts a candidate into the Dashboard's bounded lowest-risk ranking.
'------------------------------------------------------------------------------

Private Sub Dashboard_InsertLowestRisk( _
    ByRef ids() As String, _
    ByRef wbsVals() As String, _
    ByRef names() As String, _
    ByRef vals() As Double, _
    ByVal idVal As String, _
    ByVal wbsVal As String, _
    ByVal nameVal As String, _
    ByVal metricVal As Double)

    Dim i As Long
    Dim j As Long

    For i = 1 To 5
        If metricVal < vals(i) Then
            For j = 5 To i + 1 Step -1
                ids(j) = ids(j - 1)
                wbsVals(j) = wbsVals(j - 1)
                names(j) = names(j - 1)
                vals(j) = vals(j - 1)
            Next j
            ids(i) = idVal
            wbsVals(i) = wbsVal
            names(i) = nameVal
            vals(i) = metricVal
            Exit Sub
        End If
    Next i

End Sub

'------------------------------------------------------------------------------
' FR: Actualise Update Kpi Card sans modifier les regles metier qui produisent les donnees.
' EN: Refreshes Update Kpi Card without changing the business rules that produce the data.
' FR - Effet de bord : cree ou met a jour des shapes Excel.
' EN - Side effect: creates or updates Excel shapes.
'------------------------------------------------------------------------------

Private Sub Dashboard_UpdateKpiCard(ByVal ws As Worksheet, ByVal kpiIndex As Long, ByVal titleText As String, ByVal valueText As String, ByVal subText As String, ByVal X As Double, ByVal Y As Double, ByVal w As Double, ByVal h As Double, ByVal accentColor As Long)

    Dim bg As shape
    Dim accent As shape
    Dim titleBox As shape
    Dim valueBox As shape
    Dim subBox As shape
    Dim prefixName As String

    prefixName = DASH_PREFIX & "KPI_" & CStr(kpiIndex) & "_"

    Set bg = Dashboard_GetShape(ws, prefixName & "BG")
    Set accent = Dashboard_GetShape(ws, prefixName & "ACCENT")
    Set titleBox = Dashboard_GetShape(ws, prefixName & "TITLE")
    Set valueBox = Dashboard_GetShape(ws, prefixName & "HERO")
    Set subBox = Dashboard_GetShape(ws, prefixName & "SUB")

    If bg Is Nothing Or accent Is Nothing Or titleBox Is Nothing Or valueBox Is Nothing Or subBox Is Nothing Then
        Dashboard_DeleteShapesInRect ws, X - 2, Y - 2, w + 4, h + 4

        Set bg = ws.Shapes.AddShape(msoShapeRoundedRectangle, X, Y, w, h)
        bg.name = prefixName & "BG"
        bg.fill.ForeColor.RGB = RGB(255, 255, 255)
        bg.Line.ForeColor.RGB = RGB(224, 230, 237)

        Set accent = ws.Shapes.AddShape(msoShapeRectangle, X, Y, 6, h)
        accent.name = prefixName & "ACCENT"
        accent.Line.Visible = msoFalse

        Set titleBox = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, X + 16, Y + 8, w - 24, 16)
        titleBox.name = prefixName & "TITLE"
        Dashboard_FormatStableTextBox titleBox, 9, RGB(96, 111, 128), False

        Set valueBox = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, X + 16, Y + 24, w - 24, 34)
        valueBox.name = prefixName & "HERO"
        Dashboard_FormatStableTextBox valueBox, 21, RGB(20, 34, 51), True

        Set subBox = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, X + 16, Y + h - 20, w - 24, 18)
        subBox.name = prefixName & "SUB"
        Dashboard_FormatStableTextBox subBox, 8, RGB(96, 111, 128), False
    End If

    Dashboard_PositionKpiShape bg, X, Y, w, h
    Dashboard_PositionKpiShape accent, X, Y, 6, h
    Dashboard_PositionKpiShape titleBox, X + 16, Y + 8, w - 24, 16
    Dashboard_PositionKpiShape valueBox, X + 16, Y + 24, w - 24, 34
    Dashboard_PositionKpiShape subBox, X + 16, Y + h - 20, w - 24, 18

    bg.fill.ForeColor.RGB = RGB(255, 255, 255)
    bg.Line.ForeColor.RGB = RGB(224, 230, 237)
    accent.fill.ForeColor.RGB = accentColor
    accent.Line.Visible = msoFalse

    Dashboard_SetStableText titleBox, titleText, 9, RGB(96, 111, 128), False
    Dashboard_SetStableText valueBox, valueText, 21, RGB(20, 34, 51), True
    Dashboard_SetStableText subBox, subText, 8, RGB(96, 111, 128), False

End Sub

'------------------------------------------------------------------------------
' FR: Retourne la reference Shape sans exposer de mutateur sur l'etat source.
' EN: Returns the Shape reference without exposing a mutator for source state.
' FR - Effet de bord : cree ou met a jour des shapes Excel.
' EN - Side effect: creates or updates Excel shapes.
'------------------------------------------------------------------------------

Private Function Dashboard_GetShape(ByVal ws As Worksheet, ByVal shapeName As String) As shape

    On Error Resume Next
    Set Dashboard_GetShape = ws.Shapes(shapeName)
    On Error GoTo 0

End Function

'------------------------------------------------------------------------------
' FR: Traite la reference Position Kpi Shape sans modifier les donnees d'entree.
' EN: Handles the Position Kpi Shape reference without mutating input data.
'------------------------------------------------------------------------------

Private Sub Dashboard_PositionKpiShape(ByVal shp As shape, ByVal X As Double, ByVal Y As Double, ByVal w As Double, ByVal h As Double)

    shp.left = X
    shp.top = Y
    shp.width = w
    shp.height = h

End Sub

'------------------------------------------------------------------------------
' FR: Reinitialise Delete Shapes In Rect dans le perimetre possede par le composant.
' EN: Resets Delete Shapes In Rect within the state owned by the component.
' FR - Effet de bord : efface uniquement les donnees ou objets cibles du contrat.
' EN - Side effect: clears only data or objects targeted by the contract.
'------------------------------------------------------------------------------

Private Sub Dashboard_DeleteShapesInRect(ByVal ws As Worksheet, ByVal X As Double, ByVal Y As Double, ByVal w As Double, ByVal h As Double)

    Dim i As Long
    Dim shp As shape
    Dim rightEdge As Double
    Dim bottomEdge As Double

    rightEdge = X + w
    bottomEdge = Y + h

    For i = ws.Shapes.count To 1 Step -1
        Set shp = ws.Shapes(i)
        If shp.left >= X And shp.left <= rightEdge And shp.top >= Y And shp.top <= bottomEdge Then
            shp.Delete
        End If
    Next i

End Sub

'------------------------------------------------------------------------------
' FR: Normalise ou formate Format Stable Text Box selon le contrat canonique du composant.
' EN: Normalizes or formats Format Stable Text Box according to the component contract.
'------------------------------------------------------------------------------

Private Sub Dashboard_FormatStableTextBox(ByVal shp As shape, ByVal fontSize As Double, ByVal colorVal As Long, ByVal isBold As Boolean)

    shp.fill.Visible = msoFalse
    shp.Line.Visible = msoFalse
    shp.TextFrame2.MarginLeft = 0
    shp.TextFrame2.MarginRight = 0
    shp.TextFrame2.MarginTop = 0
    shp.TextFrame2.MarginBottom = 0
    shp.TextFrame2.VerticalAnchor = msoAnchorMiddle
    shp.TextFrame2.WordWrap = msoFalse
    shp.TextFrame2.TextRange.font.name = "Segoe UI"
    shp.TextFrame2.TextRange.font.size = fontSize
    shp.TextFrame2.TextRange.font.fill.ForeColor.RGB = colorVal
    shp.TextFrame2.TextRange.font.bold = isBold

End Sub

'------------------------------------------------------------------------------
' FR: Active ou initialise Set Stable Text dans l'etat runtime du composant.
' EN: Activates or initializes Set Stable Text in the component runtime state.
'------------------------------------------------------------------------------

Private Sub Dashboard_SetStableText(ByVal shp As shape, ByVal txt As String, ByVal fontSize As Double, ByVal colorVal As Long, ByVal isBold As Boolean)

    shp.TextFrame2.TextRange.text = txt
    Dashboard_FormatStableTextBox shp, fontSize, colorVal, isBold

End Sub
'------------------------------------------------------------------------------
' FR: Ajoute la reference Panel a la structure cible fournie par l'appelant.
' EN: Adds the Panel reference to the target structure supplied by the caller.
' FR - Effet de bord : cree ou met a jour des shapes Excel.
' EN - Side effect: creates or updates Excel shapes.
'------------------------------------------------------------------------------

Private Sub Dashboard_AddPanel(ByVal ws As Worksheet, ByVal X As Double, ByVal Y As Double, ByVal w As Double, ByVal h As Double, ByVal fillColor As Long)

    Dim shp As shape

    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, X, Y, w, h)
    shp.name = DASH_PREFIX & "PANEL"
    shp.fill.ForeColor.RGB = fillColor
    shp.Line.ForeColor.RGB = RGB(224, 230, 237)

End Sub

'------------------------------------------------------------------------------
' FR: Ajoute la reference Slide Surface a la structure cible fournie par l'appelant.
' EN: Adds the Slide Surface reference to the target structure supplied by the caller.
' FR - Effet de bord : cree ou met a jour des shapes Excel.
' EN - Side effect: creates or updates Excel shapes.
'------------------------------------------------------------------------------

Private Sub Dashboard_AddSlideSurface(ByVal ws As Worksheet, ByVal X As Double, ByVal Y As Double, ByVal w As Double, ByVal h As Double)

    Dim shp As shape

    Set shp = ws.Shapes.AddShape(msoShapeRectangle, X, Y, w, h)
    shp.name = DASH_PREFIX & "SLIDE"
    shp.fill.ForeColor.RGB = RGB(255, 255, 255)
    shp.Line.ForeColor.RGB = RGB(224, 230, 237)
    shp.ZOrder msoSendToBack

End Sub

'------------------------------------------------------------------------------
' FR: Ajoute la reference Section Title a la structure cible fournie par l'appelant.
' EN: Adds the Section Title reference to the target structure supplied by the caller.
' FR - Effet de bord : cree ou met a jour des shapes Excel.
' EN - Side effect: creates or updates Excel shapes.
'------------------------------------------------------------------------------

Private Sub Dashboard_AddSectionTitle(ByVal ws As Worksheet, ByVal titleText As String, ByVal X As Double, ByVal Y As Double, ByVal w As Double)

    Dim shp As shape

    Set shp = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, X, Y, w, 20)
    Dashboard_FormatTextBox shp, titleText, 12, RGB(20, 34, 51), True

End Sub

'------------------------------------------------------------------------------
' FR: Ajoute la reference Centered Section Title a la structure cible fournie par l'appelant.
' EN: Adds the Centered Section Title reference to the target structure supplied by the caller.
' FR - Effet de bord : cree ou met a jour des shapes Excel.
' EN - Side effect: creates or updates Excel shapes.
'------------------------------------------------------------------------------

Private Sub Dashboard_AddCenteredSectionTitle(ByVal ws As Worksheet, ByVal titleText As String, ByVal X As Double, ByVal Y As Double, ByVal w As Double)

    Dim shp As shape

    Set shp = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, X, Y, w, 20)
    Dashboard_FormatTextBox shp, titleText, 12, RGB(20, 34, 51), True
    shp.TextFrame2.TextRange.ParagraphFormat.alignment = msoAlignCenter

End Sub

'------------------------------------------------------------------------------
' FR: Ajoute la reference Table Title a la structure cible fournie par l'appelant.
' EN: Adds the Table Title reference to the target structure supplied by the caller.
' FR - Effet de bord : cree ou met a jour des shapes Excel.
' EN - Side effect: creates or updates Excel shapes.
'------------------------------------------------------------------------------

Private Sub Dashboard_AddTableTitle(ByVal ws As Worksheet, ByVal titleText As String, ByVal X As Double, ByVal Y As Double, ByVal w As Double)

    Dim shp As shape

    Set shp = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, X, Y, w, 18)
    Dashboard_FormatTextBox shp, titleText, 11, RGB(20, 34, 51), True

End Sub

'------------------------------------------------------------------------------
' FR: Ecrit ou synchronise Write Hotspot Header dans le stockage possede par le domaine.
' EN: Writes or synchronizes Write Hotspot Header in the store owned by the domain.
'------------------------------------------------------------------------------

Private Sub Dashboard_WriteHotspotHeader(ByVal ws As Worksheet, ByVal X As Double, ByVal Y As Double, ByVal headers As Variant)

    Dashboard_AddTinyText ws, CStr(headers(0)), X, Y, 54, 14, RGB(96, 111, 128)
    Dashboard_AddTinyText ws, CStr(headers(1)), X + 58, Y, 138, 14, RGB(96, 111, 128)
    Dashboard_AddTinyText ws, CStr(headers(2)), X + 202, Y, 52, 14, RGB(96, 111, 128)

End Sub

'------------------------------------------------------------------------------
' FR: Ecrit ou synchronise Write Hotspot Row dans le stockage possede par le domaine.
' EN: Writes or synchronizes Write Hotspot Row in the store owned by the domain.
'------------------------------------------------------------------------------

Private Sub Dashboard_WriteHotspotRow(ByVal ws As Worksheet, ByVal X As Double, ByVal Y As Double, ByVal wbsVal As String, ByVal nameVal As String, ByVal metricText As String, ByVal metricColor As Long)

    Dashboard_AddTinyText ws, wbsVal, X, Y, 54, 16, RGB(20, 34, 51)
    Dashboard_AddTinyText ws, nameVal, X + 58, Y, 138, 16, RGB(42, 52, 65)
    Dashboard_AddTinyText ws, metricText, X + 202, Y, 52, 16, metricColor

End Sub

'------------------------------------------------------------------------------
' FR: Ajoute la reference Mini Gantt Label a la structure cible fournie par l'appelant.
' EN: Adds the Mini Gantt Label reference to the target structure supplied by the caller.
' FR - Effet de bord : cree ou met a jour des shapes Excel.
' EN - Side effect: creates or updates Excel shapes.
'------------------------------------------------------------------------------

Private Sub Dashboard_AddMiniGanttLabel(ByVal ws As Worksheet, ByVal txt As String, ByVal X As Double, ByVal Y As Double, ByVal w As Double, ByVal fontSize As Double)

    Dim shp As shape

    Set shp = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, X, Y, w, fontSize + 4)
    Dashboard_FormatTextBox shp, txt, fontSize, RGB(42, 52, 65), False
    shp.TextFrame2.VerticalAnchor = msoAnchorMiddle
    shp.TextFrame2.TextRange.ParagraphFormat.alignment = msoAlignLeft

End Sub

'------------------------------------------------------------------------------
' FR: Retourne la valeur Truncate Text sans modifier les donnees d'entree.
' EN: Returns the Truncate Text value without mutating input data.
'------------------------------------------------------------------------------

Private Function Dashboard_TruncateText(ByVal txt As String, ByVal maxLen As Long) As String

    If Len(txt) <= maxLen Then
        Dashboard_TruncateText = txt
    ElseIf maxLen <= 3 Then
        Dashboard_TruncateText = left$(txt, maxLen)
    Else
        Dashboard_TruncateText = left$(txt, maxLen - 3) & "..."
    End If

End Function

'------------------------------------------------------------------------------
' FR: Ajoute la reference Tiny Label a la structure cible fournie par l'appelant.
' EN: Adds the Tiny Label reference to the target structure supplied by the caller.
'------------------------------------------------------------------------------

Private Sub Dashboard_AddTinyLabel(ByVal ws As Worksheet, ByVal txt As String, ByVal X As Double, ByVal Y As Double, ByVal w As Double)

    Dashboard_AddTinyText ws, txt, X, Y, w, 14, RGB(96, 111, 128)

End Sub

'------------------------------------------------------------------------------
' FR: Ajoute la reference Tiny Text a la structure cible fournie par l'appelant.
' EN: Adds the Tiny Text reference to the target structure supplied by the caller.
' FR - Effet de bord : cree ou met a jour des shapes Excel.
' EN - Side effect: creates or updates Excel shapes.
'------------------------------------------------------------------------------

Private Sub Dashboard_AddTinyText(ByVal ws As Worksheet, ByVal txt As String, ByVal X As Double, ByVal Y As Double, ByVal w As Double, ByVal h As Double, ByVal colorVal As Long)

    Dim shp As shape

    Set shp = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, X, Y, w, h)
    Dashboard_FormatTextBox shp, txt, 8, colorVal, False

End Sub

'------------------------------------------------------------------------------
' FR: Ajoute la reference Axis Tick Label a la structure cible fournie par l'appelant.
' EN: Adds the Axis Tick Label reference to the target structure supplied by the caller.
' FR - Effet de bord : cree ou met a jour des shapes Excel.
' EN - Side effect: creates or updates Excel shapes.
'------------------------------------------------------------------------------

Private Sub Dashboard_AddAxisTickLabel(ByVal ws As Worksheet, ByVal txt As String, ByVal X As Double, ByVal Y As Double, ByVal w As Double, ByVal h As Double, ByVal colorVal As Long)

    Dim shp As shape

    Set shp = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, X, Y, w, h)
    Dashboard_FormatTextBox shp, txt, 8, colorVal, False
    shp.TextFrame2.TextRange.ParagraphFormat.alignment = msoAlignCenter

End Sub

'------------------------------------------------------------------------------
' FR: Ecrit ou synchronise Write Empty State dans le stockage possede par le domaine.
' EN: Writes or synchronizes Write Empty State in the store owned by the domain.
'------------------------------------------------------------------------------

Private Sub Dashboard_WriteEmptyState(ByVal ws As Worksheet, ByVal X As Double, ByVal Y As Double, ByVal w As Double, ByVal txt As String)

    Dashboard_AddTinyText ws, TextCatalog_Format("DASHBOARD.STATUS.OK_DETAIL", Dashboard_CurrentLanguage(), TextCatalog_Arguments("Details", txt)), X, Y, w, 16, RGB(0, 145, 112)

End Sub

'------------------------------------------------------------------------------
' FR: Ecrit ou synchronise Write Dashboard Empty State dans le stockage possede par le domaine.
' EN: Writes or synchronizes Write Dashboard Empty State in the store owned by the domain.
' FR - Effet de bord : cree ou met a jour des shapes Excel.
' EN - Side effect: creates or updates Excel shapes.
'------------------------------------------------------------------------------

Private Sub Dashboard_WriteDashboardEmptyState(ByVal ws As Worksheet, ByVal X As Double, ByVal Y As Double, ByVal w As Double)

    Dim shp As shape

    Set shp = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, X, Y, w, 18)
    Dashboard_FormatTextBox shp, Dashboard_NoProjectLoadedText(), 9, RGB(96, 111, 128), True
    shp.TextFrame2.TextRange.ParagraphFormat.alignment = msoAlignCenter

End Sub

'------------------------------------------------------------------------------
' FR: Ajoute la reference Badge a la structure cible fournie par l'appelant.
' EN: Adds the Badge reference to the target structure supplied by the caller.
' FR - Effet de bord : cree ou met a jour des shapes Excel.
' EN - Side effect: creates or updates Excel shapes.
'------------------------------------------------------------------------------

Private Sub Dashboard_AddBadge(ByVal ws As Worksheet, ByVal txt As String, ByVal X As Double, ByVal Y As Double, ByVal w As Double, ByVal h As Double, ByVal fillColor As Long, ByVal textColor As Long)

    Dim bg As shape
    Dim t As shape

    Set bg = ws.Shapes.AddShape(msoShapeRoundedRectangle, X, Y, w, h)
    bg.name = DASH_PREFIX & "BADGE"
    bg.fill.ForeColor.RGB = fillColor
    bg.Line.Visible = msoFalse

    Set t = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, X + 8, Y + 3, w - 16, h - 6)
    Dashboard_FormatTextBox t, txt, 8, textColor, True
    t.TextFrame2.TextRange.ParagraphFormat.alignment = msoAlignCenter

End Sub

'------------------------------------------------------------------------------
' FR: Ajoute la reference Bar a la structure cible fournie par l'appelant.
' EN: Adds the Bar reference to the target structure supplied by the caller.
' FR - Effet de bord : cree ou met a jour des shapes Excel.
' EN - Side effect: creates or updates Excel shapes.
'------------------------------------------------------------------------------

Private Sub Dashboard_AddBar(ByVal ws As Worksheet, ByVal X As Double, ByVal Y As Double, ByVal w As Double, ByVal h As Double, ByVal colorVal As Long)

    Dim shp As shape

    If w < 2 Then w = 2
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, X, Y, w, h)
    shp.name = DASH_PREFIX & "BAR"
    shp.fill.ForeColor.RGB = colorVal
    shp.Line.Visible = msoFalse

End Sub

'------------------------------------------------------------------------------
' FR: Ajoute la reference Milestone Diamond a la structure cible fournie par l'appelant.
' EN: Adds the Milestone Diamond reference to the target structure supplied by the caller.
' FR - Effet de bord : cree ou met a jour des shapes Excel.
' EN - Side effect: creates or updates Excel shapes.
'------------------------------------------------------------------------------

Private Sub Dashboard_AddMilestoneDiamond(ByVal ws As Worksheet, ByVal centerX As Double, ByVal centerY As Double, ByVal sizeVal As Double, ByVal colorVal As Long)

    Dim shp As shape

    If sizeVal < 7 Then sizeVal = 7
    Set shp = ws.Shapes.AddShape(msoShapeDiamond, centerX - (sizeVal / 2), centerY - (sizeVal / 2), sizeVal, sizeVal)
    shp.name = DASH_PREFIX & "MILESTONE"
    shp.fill.ForeColor.RGB = colorVal
    shp.Line.Visible = msoFalse

End Sub

'------------------------------------------------------------------------------
' FR: Retourne la valeur Mini Gantt Current Color sans modifier les donnees d'entree.
' EN: Returns the Mini Gantt Current Color value without mutating input data.
'------------------------------------------------------------------------------

Private Function Dashboard_MiniGanttCurrentColor(ByVal progressVal As Double) As Long

    If progressVal >= 1# Then
        Dashboard_MiniGanttCurrentColor = RGB(112, 180, 71)
    Else
        Dashboard_MiniGanttCurrentColor = RGB(68, 114, 196)
    End If

End Function

'------------------------------------------------------------------------------
' FR: Ajoute la reference Progress Bar a la structure cible fournie par l'appelant.
' EN: Adds the Progress Bar reference to the target structure supplied by the caller.
' FR - Effet de bord : cree ou met a jour des shapes Excel.
' EN - Side effect: creates or updates Excel shapes.
'------------------------------------------------------------------------------

Private Sub Dashboard_AddProgressBar(ByVal ws As Worksheet, ByVal X As Double, ByVal Y As Double, ByVal w As Double, ByVal h As Double, ByVal progressVal As Double, ByVal startVal As Variant, ByVal finishVal As Variant, Optional ByVal consolidatedColor As Variant)

    Dim baseColor As Long
    Dim progressColor As Variant
    Dim progressWidth As Double
    Dim shp As shape

    baseColor = Dashboard_MiniGanttCurrentColor(progressVal)

    Dashboard_AddBar ws, X, Y, w, h, baseColor

    If progressVal > 0# And progressVal < 1# Then
            If IsMissing(consolidatedColor) Or IsEmpty(consolidatedColor) Then
                progressColor = GetProgressFillColor(Dashboard_DateNumber(startVal), Dashboard_DateNumber(finishVal), progressVal)
            Else
                progressColor = consolidatedColor
            End If
        If Not IsEmpty(progressColor) Then
            progressWidth = w * progressVal
            If progressWidth >= 6 Then
                If progressWidth > w - 2 Then progressWidth = w - 2
                Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, X + 1, Y + 1, progressWidth, h - 2)
                shp.name = DASH_PREFIX & "BAR_PROGRESS"
                shp.fill.ForeColor.RGB = CLng(progressColor)
                shp.fill.Transparency = 0.25
                shp.Line.Visible = msoFalse
            End If
        End If
    End If

End Sub

'------------------------------------------------------------------------------
' FR: Actualise Render Time Axis sans modifier les regles metier qui produisent les donnees.
' EN: Refreshes Render Time Axis without changing the business rules that produce the data.
'------------------------------------------------------------------------------

Private Sub Dashboard_RenderTimeAxis(ByVal ws As Worksheet, ByVal minDate As Double, ByVal maxDate As Double, ByVal X As Double, ByVal Y As Double, ByVal w As Double)

    Dim spanDays As Double
    Dim tickDate As Date
    Dim tickEnd As Date
    Dim tickX As Double
    Dim stepKind As String
    Dim stepCount As Long
    Dim labelText As String

    If maxDate <= minDate Then Exit Sub

    Dashboard_AddLine ws, X, Y + 12, X + w, Y + 12, RGB(224, 230, 237), 0.75

    spanDays = maxDate - minDate
    If spanDays <= 14 Then
        stepKind = "d"
        stepCount = 1
        tickDate = CDate(Int(minDate))
    ElseIf spanDays <= 62 Then
        stepKind = "ww"
        stepCount = 1
        tickDate = DateAdd("d", 1 - Weekday(CDate(Int(minDate)), vbMonday), CDate(Int(minDate)))
    ElseIf spanDays <= 730 Then
        stepKind = "m"
        stepCount = 1
        tickDate = DateSerial(Year(CDate(minDate)), Month(CDate(minDate)), 1)
    Else
        stepKind = "yyyy"
        stepCount = 1
        tickDate = DateSerial(Year(CDate(minDate)), 1, 1)
    End If

    tickEnd = CDate(maxDate)
    Do While CDbl(tickDate) <= CDbl(tickEnd)
        tickX = Dashboard_DateToX(CDbl(tickDate), minDate, maxDate, X, w)
        If tickX >= X - 1 And tickX <= X + w + 1 Then
            Dashboard_AddLine ws, tickX, Y + 8, tickX, Y + 15, RGB(189, 198, 208), 0.75
            Select Case stepKind
                Case "d"
                    labelText = Dashboard_FormatDate(tickDate, False)
                Case "ww"
                    labelText = TextCatalog_Get("DASHBOARD.UNIT.WEEK_PREFIX", Dashboard_CurrentLanguage()) & Format$(tickDate, "ww", vbMonday, vbFirstFourDays)
                Case "m"
                    labelText = Dashboard_FormatMonthShort(tickDate)
                Case Else
                    labelText = Format$(tickDate, "yyyy")
            End Select
            Dashboard_AddAxisTickLabel ws, labelText, tickX - 21, Y - 5, 42, 12, RGB(96, 111, 128)
        End If

        tickDate = DateAdd(stepKind, stepCount, tickDate)
    Loop

End Sub

'------------------------------------------------------------------------------
' FR: Ajoute la reference Line a la structure cible fournie par l'appelant.
' EN: Adds the Line reference to the target structure supplied by the caller.
' FR - Effet de bord : cree ou met a jour des shapes Excel.
' EN - Side effect: creates or updates Excel shapes.
'------------------------------------------------------------------------------

Private Sub Dashboard_AddLine(ByVal ws As Worksheet, ByVal x1 As Double, ByVal y1 As Double, ByVal x2 As Double, ByVal y2 As Double, ByVal colorVal As Long, ByVal weightVal As Double)

    Dim shp As shape

    Set shp = ws.Shapes.AddLine(x1, y1, x2, y2)
    shp.name = DASH_PREFIX & "LINE"
    shp.Line.ForeColor.RGB = colorVal
    shp.Line.weight = weightVal
    shp.Line.DashStyle = msoLineDash

End Sub

'------------------------------------------------------------------------------
' FR: Normalise ou formate Format Text Box selon le contrat canonique du composant.
' EN: Normalizes or formats Format Text Box according to the component contract.
'------------------------------------------------------------------------------

Private Sub Dashboard_FormatTextBox(ByVal shp As shape, ByVal txt As String, ByVal fontSize As Double, ByVal colorVal As Long, ByVal isBold As Boolean)

    shp.name = DASH_PREFIX & "TXT"
    shp.fill.Visible = msoFalse
    shp.Line.Visible = msoFalse
    shp.TextFrame2.MarginLeft = 0
    shp.TextFrame2.MarginRight = 0
    shp.TextFrame2.MarginTop = 0
    shp.TextFrame2.MarginBottom = 0
    shp.TextFrame2.TextRange.text = txt
    shp.TextFrame2.TextRange.font.name = "Segoe UI"
    shp.TextFrame2.TextRange.font.size = fontSize
    shp.TextFrame2.TextRange.font.fill.ForeColor.RGB = colorVal
    shp.TextFrame2.TextRange.font.bold = IIf(isBold, msoTrue, msoFalse)

End Sub

'------------------------------------------------------------------------------
' FR: Normalise ou formate Format Date Short selon le contrat canonique du composant.
' EN: Normalizes or formats Format Date Short according to the component contract.
'------------------------------------------------------------------------------

Private Function Dashboard_FormatDateShort(ByVal dateVal As Variant) As String

    If Dashboard_HasDateValue(dateVal) Then
        Dashboard_FormatDateShort = Dashboard_FormatDate(dateVal, True)
    Else
        Dashboard_FormatDateShort = "-"
    End If

End Function

'------------------------------------------------------------------------------
' FR: Indique si la valeur Date Value satisfait la condition attendue, sans modifier les donnees source.
' EN: Returns whether the Date Value value satisfies the expected condition without mutating source data.
'------------------------------------------------------------------------------

Private Function Dashboard_HasDateValue(ByVal dateVal As Variant) As Boolean

    Dim v As Variant

    v = GetCellValue(dateVal)
    If Not HasValue(v) Then Exit Function

    Dashboard_HasDateValue = (IsNumeric(v) Or IsDate(v))

End Function

'------------------------------------------------------------------------------
' FR: Retourne la valeur Date Number sans modifier les donnees d'entree.
' EN: Returns the Date Number value without mutating input data.
'------------------------------------------------------------------------------

Private Function Dashboard_DateNumber(ByVal dateVal As Variant) As Double

    If IsDate(dateVal) Then
        Dashboard_DateNumber = CDbl(CDate(dateVal))
    Else
        Dashboard_DateNumber = CDbl(dateVal)
    End If

End Function

'------------------------------------------------------------------------------
' FR: Normalise ou formate Format Signed Days selon le contrat canonique du composant.
' EN: Normalizes or formats Format Signed Days according to the component contract.
'------------------------------------------------------------------------------

Private Function Dashboard_FormatSignedDays(ByVal dayVal As Variant) As String

    If Not HasValue(dayVal) Then
        Dashboard_FormatSignedDays = "-"
    ElseIf CLng(dayVal) > 0 Then
        Dashboard_FormatSignedDays = "+" & CStr(CLng(dayVal)) & TextCatalog_Get("DASHBOARD.UNIT.DAY_LONG", Dashboard_CurrentLanguage())
    Else
        Dashboard_FormatSignedDays = CStr(CLng(dayVal)) & TextCatalog_Get("DASHBOARD.UNIT.DAY_LONG", Dashboard_CurrentLanguage())
    End If

End Function

'------------------------------------------------------------------------------
' FR: Normalise ou formate Format Signed Compact Days selon le contrat canonique du composant.
' EN: Normalizes or formats Format Signed Compact Days according to the component contract.
'------------------------------------------------------------------------------

Private Function Dashboard_FormatSignedCompactDays(ByVal dayVal As Variant) As String

    If Not HasValue(dayVal) Then
        Dashboard_FormatSignedCompactDays = "-"
    ElseIf CLng(dayVal) > 0 Then
        Dashboard_FormatSignedCompactDays = "+" & CStr(CLng(dayVal)) & TextCatalog_Get("DASHBOARD.UNIT.DAY_COMPACT", Dashboard_CurrentLanguage())
    Else
        Dashboard_FormatSignedCompactDays = CStr(CLng(dayVal)) & TextCatalog_Get("DASHBOARD.UNIT.DAY_COMPACT", Dashboard_CurrentLanguage())
    End If

End Function

'------------------------------------------------------------------------------
' FR: Normalise ou formate Format Signed Number selon le contrat canonique du composant.
' EN: Normalizes or formats Format Signed Number according to the component contract.
'------------------------------------------------------------------------------

Private Function Dashboard_FormatSignedNumber(ByVal valueVal As Variant) As String

    If Not HasValue(valueVal) Then
        Dashboard_FormatSignedNumber = "-"
    ElseIf CLng(valueVal) > 0 Then
        Dashboard_FormatSignedNumber = "+" & CStr(CLng(valueVal))
    Else
        Dashboard_FormatSignedNumber = CStr(CLng(valueVal))
    End If

End Function

'------------------------------------------------------------------------------
' FR: Normalise ou formate Format Momentum Delay Percent selon le contrat canonique du composant.
' EN: Normalizes or formats Format Momentum Delay Percent according to the component contract.
'------------------------------------------------------------------------------

Private Function Dashboard_FormatMomentumDelayPercent(ByVal valueVal As Double) As String

    If Abs(valueVal) < 0.01 Then
        If valueVal > 0 Then
            Dashboard_FormatMomentumDelayPercent = "+" & Format$(valueVal, "0.0%")
        Else
            Dashboard_FormatMomentumDelayPercent = Format$(valueVal, "0.0%")
        End If
    Else
        Dashboard_FormatMomentumDelayPercent = Dashboard_FormatPercentSigned(valueVal)
    End If

End Function

'------------------------------------------------------------------------------
' FR: Normalise ou formate Format Percent Signed selon le contrat canonique du composant.
' EN: Normalizes or formats Format Percent Signed according to the component contract.
'------------------------------------------------------------------------------

Private Function Dashboard_FormatPercentSigned(ByVal valueVal As Double) As String

    If valueVal > 0 Then
        Dashboard_FormatPercentSigned = "+" & Format$(valueVal, "0%")
    Else
        Dashboard_FormatPercentSigned = Format$(valueVal, "0%")
    End If

End Function

'------------------------------------------------------------------------------
' FR: Retourne la valeur Language sans exposer de mutateur sur l'etat source.
' EN: Returns the Language value without exposing a mutator for source state.
'------------------------------------------------------------------------------

Public Function Dashboard_CurrentLanguage() As String

    If UCase$(Trim$(gDashboardLanguage)) <> "FR" And UCase$(Trim$(gDashboardLanguage)) <> "EN" Then
        gDashboardLanguage = "EN"
    End If
    Dashboard_CurrentLanguage = UCase$(Trim$(gDashboardLanguage))

End Function

'------------------------------------------------------------------------------
' FR: Indique si la valeur French satisfait la condition attendue, sans modifier les donnees source.
' EN: Returns whether the French value satisfies the expected condition without mutating source data.
'------------------------------------------------------------------------------

Private Function Dashboard_IsFrench() As Boolean

    Dashboard_IsFrench = (Dashboard_CurrentLanguage() = "FR")

End Function

'------------------------------------------------------------------------------
' FR: Retourne la valeur L sans modifier les donnees d'entree.
' EN: Returns the L value without mutating input data.
'------------------------------------------------------------------------------


'------------------------------------------------------------------------------
' FR: Retourne la valeur No Project Loaded Text sans modifier les donnees d'entree.
' EN: Returns the No Project Loaded Text value without mutating input data.
'------------------------------------------------------------------------------

Private Function Dashboard_NoProjectLoadedText() As String

    Dashboard_NoProjectLoadedText = TextCatalog_Get("DASHBOARD.EMPTY.NO_PROJECT", Dashboard_CurrentLanguage())

End Function

'------------------------------------------------------------------------------
' FR: Retourne la valeur No Data Text sans modifier les donnees d'entree.
' EN: Returns the No Data Text value without mutating input data.
'------------------------------------------------------------------------------

Private Function Dashboard_NoDataText() As String

    Dashboard_NoDataText = TextCatalog_Get("DASHBOARD.EMPTY.NO_DATA", Dashboard_CurrentLanguage())

End Function

'------------------------------------------------------------------------------
' FR: Retourne la valeur Duration Suffix sans modifier les donnees d'entree.
' EN: Returns the Duration Suffix value without mutating input data.
'------------------------------------------------------------------------------

Private Function Dashboard_DurationSuffix() As String

    Dashboard_DurationSuffix = TextCatalog_Get("DASHBOARD.UNIT.DAY_COMPACT", Dashboard_CurrentLanguage())

End Function

'------------------------------------------------------------------------------
' FR: Normalise ou formate Format Month Short selon le contrat canonique du composant.
' EN: Normalizes or formats Format Month Short according to the component contract.
'------------------------------------------------------------------------------

Private Function Dashboard_FormatMonthShort(ByVal dateVal As Variant) As String

    Dim monthNames As Variant
    Dim d As Date

    If Not Dashboard_HasDateValue(dateVal) Then Exit Function
    d = CDate(Dashboard_DateNumber(dateVal))

    monthNames = Array("DASHBOARD.MONTH.JAN", "DASHBOARD.MONTH.FEB", "DASHBOARD.MONTH.MAR", "DASHBOARD.MONTH.APR", "DASHBOARD.MONTH.MAY", "DASHBOARD.MONTH.JUN", "DASHBOARD.MONTH.JUL", "DASHBOARD.MONTH.AUG", "DASHBOARD.MONTH.SEP", "DASHBOARD.MONTH.OCT", "DASHBOARD.MONTH.NOV", "DASHBOARD.MONTH.DEC")
    Dashboard_FormatMonthShort = TextCatalog_Get(CStr(monthNames(Month(d) - 1)), Dashboard_CurrentLanguage())

End Function

'------------------------------------------------------------------------------
' FR: Normalise ou formate Format Date selon le contrat canonique du composant.
' EN: Normalizes or formats Format Date according to the component contract.
'------------------------------------------------------------------------------

Private Function Dashboard_FormatDate(ByVal dateVal As Variant, Optional ByVal includeYear As Boolean = True) As String

    Dim d As Date

    If Not Dashboard_HasDateValue(dateVal) Then
        Dashboard_FormatDate = "-"
        Exit Function
    End If

    d = CDate(Dashboard_DateNumber(dateVal))
    Dashboard_FormatDate = Format$(d, "dd") & " " & Dashboard_FormatMonthShort(d)
    If includeYear Then Dashboard_FormatDate = Dashboard_FormatDate & " " & Format$(d, "yyyy")

End Function

'------------------------------------------------------------------------------
' FR: Normalise ou formate Format Timestamp selon le contrat canonique du composant.
' EN: Normalizes or formats Format Timestamp according to the component contract.
'------------------------------------------------------------------------------

Private Function Dashboard_FormatTimestamp(ByVal dateVal As Date) As String

    Dashboard_FormatTimestamp = TextCatalog_Format("DASHBOARD.TIMESTAMP.DISPLAY", Dashboard_CurrentLanguage(), TextCatalog_Arguments("Date", Dashboard_FormatDate(dateVal, True), "Time", Format$(dateVal, "hh:nn")))

End Function

'------------------------------------------------------------------------------
' FR: Indique si la valeur Timestamp Text satisfait la condition attendue, sans modifier les donnees source.
' EN: Returns whether the Timestamp Text value satisfies the expected condition without mutating source data.
'------------------------------------------------------------------------------

Private Function Dashboard_IsTimestampText(ByVal txt As String) As Boolean

    Dashboard_IsTimestampText = (left$(txt, Len(TextCatalog_Get("DASHBOARD.TIMESTAMP.PREFIX", "EN"))) = TextCatalog_Get("DASHBOARD.TIMESTAMP.PREFIX", "EN") Or _
        left$(txt, Len(TextCatalog_Get("DASHBOARD.TIMESTAMP.PREFIX", "FR"))) = TextCatalog_Get("DASHBOARD.TIMESTAMP.PREFIX", "FR"))

End Function

'------------------------------------------------------------------------------
' FR: Retourne la valeur Chart Date Number Format sans modifier les donnees d'entree.
' EN: Returns the Chart Date Number Format value without mutating input data.
'------------------------------------------------------------------------------

Private Function Dashboard_ChartDateNumberFormat(Optional ByVal includeYear As Boolean = False) As String

    If includeYear Then
        Dashboard_ChartDateNumberFormat = Settings_GetDateNumberFormat()
    Else
        Dashboard_ChartDateNumberFormat = Settings_GetCompactDateNumberFormat()
    End If

End Function

'------------------------------------------------------------------------------
' FR: Retourne la valeur Progress Color sans modifier les donnees d'entree.
' EN: Returns the Progress Color value without mutating input data.
'------------------------------------------------------------------------------

Private Function Dashboard_ProgressColor(ByVal progressVar As Double) As Long

    If progressVar < -0.03 Then
        Dashboard_ProgressColor = RGB(192, 80, 77)
    ElseIf progressVar < 0 Then
        Dashboard_ProgressColor = RGB(238, 156, 68)
    Else
        Dashboard_ProgressColor = RGB(0, 145, 112)
    End If

End Function

'------------------------------------------------------------------------------
' FR: Retourne la valeur Drift Color sans modifier les donnees d'entree.
' EN: Returns the Drift Color value without mutating input data.
'------------------------------------------------------------------------------

Private Function Dashboard_DriftColor(ByVal driftDays As Variant) As Long

    If Not HasValue(driftDays) Then
        Dashboard_DriftColor = RGB(96, 111, 128)
    ElseIf CLng(driftDays) > 0 Then
        Dashboard_DriftColor = RGB(192, 80, 77)
    Else
        Dashboard_DriftColor = RGB(0, 145, 112)
    End If

End Function




