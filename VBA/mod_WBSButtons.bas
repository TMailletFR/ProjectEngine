Attribute VB_Name = "mod_WBSButtons"
Option Explicit

'===============================================================================
' MODULE : mod_WBSButtons
' DOMAINE / DOMAIN : WBS
'
' FR
' Possede Reset/Armageddon, les boutons WBS, leur langue, le guide de saisie et la mise en forme des inputs.
' Ne doit pas contourner les contrats publics des autres domaines.
'
' EN
' Owns Reset/Armageddon, WBS buttons, their language, the onboarding guide and input formatting.
' Must not bypass public contracts owned by other domains.
'
' CONTRATS / CONTRACTS : Armageddon, Reset_Planning, WBS_EnsureOnboardingGuide, Ensure_WBS_Main_Buttons, WBS_ApplyLanguage, WBS_SetLanguage, WBS_CurrentLanguage
' CALLBACKS EXTERNES / EXTERNAL CALLBACKS : Aucun / None
'===============================================================================


Private Const RESET_PLANNING_EMPTY_WBS_ROWS As Long = 5

Private Const WBS_ONBOARDING_FONT_NAME As String = "Aptos Narrow"
Private Const WBS_ONBOARDING_FONT_SIZE As Double = 8
Private Const WBS_ONBOARDING_ROW_HEIGHT As Double = 14
Private gWBSLanguage As String
Private gWBSOnboardingLocalizedCheckCount As Long
Private gWBSOnboardingLocalizedRebuildCount As Long
Private gWBSOnboardingStructuralMutationCount As Long
Private gWBSOnboardingThreadedReadCount As Long
Private gWBSOnboardingThreadedAddCount As Long
Private gWBSOnboardingThreadedClearCount As Long
'------------------------------------------------------------------------------
' FR: Traite la reference Armageddon sans modifier les donnees d'entree.
' EN: Handles the Armageddon reference without mutating input data.
'------------------------------------------------------------------------------

Public Sub Armageddon(Optional ByVal skipConfirmation As Boolean = False)
    If Not WorkbookSchema_UserActionAllowed() Then Exit Sub

    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim answer As VbMsgBoxResult
    Dim oldEvents As Boolean
    Dim oldScreenUpdating As Boolean

    oldEvents = Application.EnableEvents
    oldScreenUpdating = Application.ScreenUpdating

    On Error GoTo ErrHandler

    If Not skipConfirmation Then
        answer = MsgBox( _
            TextCatalog_Get("WBS.FULL_RESET.CONFIRMATION", EventHistory_CurrentLanguage()), _
            vbQuestion + vbYesNo + vbDefaultButton2, _
            TextCatalog_Get("WBS.FULL_RESET.TITLE", EventHistory_CurrentLanguage()))

        If answer <> vbYes Then Exit Sub
    End If

    Set ws = ThisWorkbook.Worksheets("WBS")
    Set tbl = ws.ListObjects("tbl_WBS")

    Application.EnableEvents = False
    Application.ScreenUpdating = False
    ResetPlanning_PrepareEmptyWBS ws, tbl
    Application.ScreenUpdating = oldScreenUpdating
    Application.EnableEvents = oldEvents

    Run_Full_Update
    Reset_Dashboard
    ClearPlanningWarningAcknowledgements
    ClearPlanningEventHistory
    ProjectWelcome_ShowIfNeeded
    Exit Sub

ErrHandler:
    On Error Resume Next
    Application.ScreenUpdating = oldScreenUpdating
    Application.EnableEvents = oldEvents
    On Error GoTo 0

    WBSButtons_ShowConsoleError _
        "WBS.ERROR.FULL_RESET", _
        TextCatalog_Arguments("Details", Err.Description)

End Sub

'------------------------------------------------------------------------------
' FR: Reinitialise Reset Planning dans le perimetre possede par le composant.
' EN: Resets Reset Planning within the state owned by the component.
'------------------------------------------------------------------------------

Public Sub Reset_Planning()
    If Not WorkbookSchema_UserActionAllowed() Then Exit Sub

    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim answer As VbMsgBoxResult
    Dim oldEvents As Boolean
    Dim oldScreenUpdating As Boolean

    On Error GoTo ErrHandler

    oldEvents = Application.EnableEvents
    oldScreenUpdating = Application.ScreenUpdating
    answer = MsgBox( _
        TextCatalog_Get(TXT_COMMON_RESET_PLANNING_CONFIRMATION, EventHistory_CurrentLanguage()), _
        vbQuestion + vbYesNo + vbDefaultButton2, _
        TextCatalog_Get(TXT_COMMON_RESET_PLANNING_LABEL, EventHistory_CurrentLanguage()))

    If answer <> vbYes Then Exit Sub

    Set ws = ThisWorkbook.Worksheets("WBS")
    Set tbl = ws.ListObjects("tbl_WBS")

    oldEvents = Application.EnableEvents
    oldScreenUpdating = Application.ScreenUpdating
    Application.EnableEvents = False
    Application.ScreenUpdating = False

    ResetPlanning_PrepareEmptyWBS ws, tbl

CleanReset:
    Application.ScreenUpdating = oldScreenUpdating
    Application.EnableEvents = oldEvents

    Run_Planning_Update
    Exit Sub

ErrHandler:
    On Error Resume Next
    Application.ScreenUpdating = oldScreenUpdating
    Application.EnableEvents = oldEvents
    On Error GoTo 0

    WBSButtons_ShowConsoleError _
        "WBS.ERROR.RESET_PLANNING", _
        TextCatalog_Arguments("Details", Err.Description)

End Sub

'------------------------------------------------------------------------------
' FR: Reinitialise Reset Planning Prepare Empty WBS dans le perimetre possede par le composant.
' EN: Resets Reset Planning Prepare Empty WBS within the state owned by the component.
' FR - Effet de bord : efface uniquement les donnees ou objets cibles du contrat.
' EN - Side effect: clears only data or objects targeted by the contract.
'------------------------------------------------------------------------------

Private Sub ResetPlanning_PrepareEmptyWBS( _
    ByVal ws As Worksheet, _
    ByVal tbl As ListObject)

    If ws Is Nothing Then Exit Sub
    If tbl Is Nothing Then Exit Sub

    Do While tbl.ListRows.Count > RESET_PLANNING_EMPTY_WBS_ROWS
        tbl.ListRows(tbl.ListRows.Count).Delete
    Loop

    Do While tbl.ListRows.Count < RESET_PLANNING_EMPTY_WBS_ROWS
        tbl.ListRows.Add
    Loop

    If Not tbl.DataBodyRange Is Nothing Then
        tbl.DataBodyRange.ClearContents
        RestoreWBSFormulaColumns tbl
        ResetPlanning_ApplyWBSInputSetup tbl
    End If

End Sub

'------------------------------------------------------------------------------
' FR: Reinitialise Reset Planning Apply WBS Input Setup dans le perimetre possede par le composant.
' EN: Resets Reset Planning Apply WBS Input Setup within the state owned by the component.
'------------------------------------------------------------------------------

Private Sub ResetPlanning_ApplyWBSInputSetup(ByVal tbl As ListObject)

    If tbl Is Nothing Then Exit Sub
    If tbl.DataBodyRange Is Nothing Then Exit Sub

    WBS_ApplyLocalizedInputValidations tbl
    ResetPlanning_ApplyWBSFormats tbl

End Sub

'------------------------------------------------------------------------------
' FR: Rematerialise les validations WBS dans la langue de l'owner sans toucher aux donnees.
' EN: Rematerializes WBS validations in the owner language without changing data.
'------------------------------------------------------------------------------
Private Sub WBS_ApplyLocalizedInputValidations(ByVal tbl As ListObject)

    If tbl Is Nothing Then Exit Sub
    If tbl.DataBodyRange Is Nothing Then Exit Sub

    ' Calendar is user input: remove the inherited calculated-column fill.
    If WBS_TableHasSchemaColumn(tbl, VTS_COL_CAL) Then
        SchemaListColumn(tbl, VTS_TABLE_WBS, VTS_COL_CAL).DataBodyRange.Interior.Pattern = xlNone
    End If

    ResetPlanning_ApplyListValidation tbl, VTS_COL_TASK_TYPE, _
        TextCatalog_Get("WBS.VALIDATION.TASK_TYPE.LIST", WBS_CurrentLanguage()), _
        TextCatalog_Get("WBS.VALIDATION.TASK_TYPE.INPUT_TITLE", WBS_CurrentLanguage()), _
        TextCatalog_Get("WBS.VALIDATION.TASK_TYPE.INPUT_MESSAGE", WBS_CurrentLanguage()), _
        TextCatalog_Get("WBS.VALIDATION.TASK_TYPE.ERROR_TITLE", WBS_CurrentLanguage()), _
        TextCatalog_Get("WBS.VALIDATION.TASK_TYPE.ERROR_MESSAGE", WBS_CurrentLanguage())

    ResetPlanning_ApplyListValidation tbl, VTS_COL_S, _
        TextCatalog_Get("WBS.VALIDATION.SUMMARY.LIST", WBS_CurrentLanguage()), _
        TextCatalog_Get("WBS.VALIDATION.SUMMARY.INPUT_TITLE", WBS_CurrentLanguage()), _
        TextCatalog_Get("WBS.VALIDATION.SUMMARY.INPUT_MESSAGE", WBS_CurrentLanguage()), _
        TextCatalog_Get("WBS.VALIDATION.SUMMARY.ERROR_TITLE", WBS_CurrentLanguage()), _
        TextCatalog_Get("WBS.VALIDATION.SUMMARY.ERROR_MESSAGE", WBS_CurrentLanguage())

    ResetPlanning_ApplyListValidation tbl, VTS_COL_CAL, _
        TextCatalog_Get("WBS.VALIDATION.CALENDAR.LIST", WBS_CurrentLanguage()), _
        TextCatalog_Get("WBS.VALIDATION.CALENDAR.INPUT_TITLE", WBS_CurrentLanguage()), _
        TextCatalog_Get("WBS.VALIDATION.CALENDAR.INPUT_MESSAGE", WBS_CurrentLanguage()), _
        TextCatalog_Get("WBS.VALIDATION.CALENDAR.ERROR_TITLE", WBS_CurrentLanguage()), _
        TextCatalog_Get("WBS.VALIDATION.CALENDAR.ERROR_MESSAGE", WBS_CurrentLanguage())

End Sub

'------------------------------------------------------------------------------
' FR: Reinitialise Reset Planning Apply List Validation dans le perimetre possede par le composant.
' EN: Resets Reset Planning Apply List Validation within the state owned by the component.
' FR - Effet de bord : efface uniquement les donnees ou objets cibles du contrat.
' EN - Side effect: clears only data or objects targeted by the contract.
'------------------------------------------------------------------------------

Private Sub ResetPlanning_ApplyListValidation( _
    ByVal tbl As ListObject, _
    ByVal columnKey As String, _
    ByVal listFormula As String, _
    ByVal inputTitle As String, _
    ByVal inputMessage As String, _
    ByVal errorTitle As String, _
    ByVal errorMessage As String)

    Dim rng As Range

    If Not WBS_TableHasSchemaColumn(tbl, columnKey) Then Exit Sub
    Set rng = SchemaListColumn(tbl, VTS_TABLE_WBS, columnKey).DataBodyRange
    If rng Is Nothing Then Exit Sub

    rng.NumberFormat = "@"
    With rng.Validation
        .Delete
        .Add Type:=xlValidateList, _
             AlertStyle:=xlValidAlertStop, _
             Operator:=xlBetween, _
             Formula1:=listFormula
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = inputTitle
        .InputMessage = inputMessage
        .ErrorTitle = errorTitle
        .errorMessage = errorMessage
        .ShowInput = True
        .ShowError = True
    End With

End Sub

'------------------------------------------------------------------------------
' FR: Reinitialise Reset Planning Apply WBS Formats dans le perimetre possede par le composant.
' EN: Resets Reset Planning Apply WBS Formats within the state owned by the component.
'------------------------------------------------------------------------------

Private Sub ResetPlanning_ApplyWBSFormats(ByVal tbl As ListObject)

    Dim dateFormat As String

    dateFormat = Settings_GetDateNumberFormat()

    If WBS_TableHasSchemaColumn(tbl, VTS_COL_WBS) Then SchemaListColumn(tbl, VTS_TABLE_WBS, VTS_COL_WBS).DataBodyRange.NumberFormat = "@"
    If WBS_TableHasSchemaColumn(tbl, VTS_COL_ID) Then SchemaListColumn(tbl, VTS_TABLE_WBS, VTS_COL_ID).DataBodyRange.NumberFormat = "0"
    If WBS_TableHasSchemaColumn(tbl, VTS_COL_BASELINE_START) Then SchemaListColumn(tbl, VTS_TABLE_WBS, VTS_COL_BASELINE_START).DataBodyRange.NumberFormat = dateFormat
    If WBS_TableHasSchemaColumn(tbl, VTS_COL_BASELINE_FINISH) Then SchemaListColumn(tbl, VTS_TABLE_WBS, VTS_COL_BASELINE_FINISH).DataBodyRange.NumberFormat = dateFormat
    If WBS_TableHasSchemaColumn(tbl, VTS_COL_ACTUAL_START) Then SchemaListColumn(tbl, VTS_TABLE_WBS, VTS_COL_ACTUAL_START).DataBodyRange.NumberFormat = dateFormat
    If WBS_TableHasSchemaColumn(tbl, VTS_COL_ACTUAL_FINISH) Then SchemaListColumn(tbl, VTS_TABLE_WBS, VTS_COL_ACTUAL_FINISH).DataBodyRange.NumberFormat = dateFormat
    If WBS_TableHasSchemaColumn(tbl, VTS_COL_FORECAST_START) Then SchemaListColumn(tbl, VTS_TABLE_WBS, VTS_COL_FORECAST_START).DataBodyRange.NumberFormat = dateFormat
    If WBS_TableHasSchemaColumn(tbl, VTS_COL_FORECAST_FINISH) Then SchemaListColumn(tbl, VTS_TABLE_WBS, VTS_COL_FORECAST_FINISH).DataBodyRange.NumberFormat = dateFormat
    If WBS_TableHasSchemaColumn(tbl, VTS_COL_BASELINE_DURATION) Then SchemaListColumn(tbl, VTS_TABLE_WBS, VTS_COL_BASELINE_DURATION).DataBodyRange.NumberFormat = "0"
    If WBS_TableHasSchemaColumn(tbl, VTS_COL_ACTUAL_DURATION) Then SchemaListColumn(tbl, VTS_TABLE_WBS, VTS_COL_ACTUAL_DURATION).DataBodyRange.NumberFormat = "0"
    If WBS_TableHasSchemaColumn(tbl, VTS_COL_CALCULATED_DURATION) Then SchemaListColumn(tbl, VTS_TABLE_WBS, VTS_COL_CALCULATED_DURATION).DataBodyRange.NumberFormat = "0"

End Sub

'=====================================================
' mod_WBSButtons
'
' WBS main buttons + Task Type setup.
'
' Console routing:
' - confirmation utilisateur autorisee pour Reset Planning
' - les erreurs sont envoyées vers frmPlanningMessages
'=====================================================

'------------------------------------------------------------------------------
' FR: Garantit le layout d'aide WBS v1.0.1 et migre un ancien layout compatible.
' EN: Ensures the v1.0.1 WBS help layout and migrates a compatible legacy layout.
' FR - Effet de bord : insere la ligne 3 une seule fois si tbl_WBS commence en ligne 4.
' EN - Side effect: inserts row 3 exactly once when tbl_WBS starts on row 4.
'------------------------------------------------------------------------------
Public Sub WBS_EnsureOnboardingGuide()

    Dim ws As Worksheet

    Set ws = ThisWorkbook.Worksheets("WBS")
    Ensure_WBS_Onboarding_Guide ws

End Sub

'------------------------------------------------------------------------------
' FR: Migre et maintient le guide de saisie WBS sans dependre de la position des colonnes.
' EN: Migrates and maintains the WBS onboarding guide without relying on column positions.
' FR - Effet de bord : insere la ligne 3 une seule fois dans les anciens classeurs compatibles.
' EN - Side effect: inserts row 3 exactly once in compatible legacy workbooks.
'------------------------------------------------------------------------------
Private Sub Ensure_WBS_Onboarding_Guide( _
    ByVal ws As Worksheet, _
    Optional ByVal applyLocalizedContent As Boolean = True)

    Dim tbl As ListObject
    Dim statuses As Object
    Dim key As Variant
    Dim missingColumns As String
    Dim headerRow As Long
    Dim firstDataRow As Long
    Dim oldEvents As Boolean
    Dim oldScreenUpdating As Boolean
    Dim renamedLegacyProjectColumn As Boolean
    Dim structureNeedsMutation As Boolean
    Dim errorNumber As Long
    Dim errorSource As String
    Dim errorDescription As String

    If ws Is Nothing Then Exit Sub

    oldEvents = Application.EnableEvents
    oldScreenUpdating = Application.ScreenUpdating
    On Error GoTo ErrHandler

    Set tbl = ws.ListObjects("tbl_WBS")
    Set statuses = WBS_Onboarding_BuildStatusMap()

    If WBS_TableHasSchemaColumn(tbl, VTS_COL_PROJECT) And WBS_TableHasPhysicalColumn(tbl, "Package") Then
        Err.Raise vbObjectError + 2329, "Ensure_WBS_Onboarding_Guide", _
            PlanningMessageText_Format("WBS.ERROR.AMBIGUOUS_SCHEMA")
    End If

    For Each key In statuses.Keys
        If Not WBS_TableHasSchemaColumn(tbl, CStr(key)) Then
            If CStr(key) <> VTS_COL_PROJECT Or Not WBS_TableHasPhysicalColumn(tbl, "Package") Then
                If missingColumns <> "" Then missingColumns = missingColumns & ", "
                missingColumns = missingColumns & CStr(key)
            End If
        End If
    Next key

    If missingColumns <> "" Then
        Err.Raise vbObjectError + 2330, "Ensure_WBS_Onboarding_Guide", _
            PlanningMessageText_Format("WBS.ERROR.MISSING_CANONICAL_COLUMNS", _
                TextCatalog_Arguments("Columns", missingColumns), TextCatalog_Arguments("Columns", missingColumns))
    End If

    headerRow = tbl.HeaderRowRange.Row
    If headerRow <= 1 Then
        Err.Raise vbObjectError + 2331, "Ensure_WBS_Onboarding_Guide", _
            PlanningMessageText_Format("WBS.ERROR.NO_ONBOARDING_ROW")
    End If

    structureNeedsMutation = Not WBS_TableHasSchemaColumn(tbl, VTS_COL_PROJECT)

    If Not structureNeedsMutation Then
        structureNeedsMutation = Not WBS_Onboarding_StructureIsCurrent(ws, tbl, statuses)
    End If

    If structureNeedsMutation Then
        Application.EnableEvents = False
        Application.ScreenUpdating = False

        If Not WBS_TableHasSchemaColumn(tbl, VTS_COL_PROJECT) Then
            tbl.ListColumns("Package").Name = SchemaCurrentColumnTitle(VTS_TABLE_WBS, VTS_COL_PROJECT)
            renamedLegacyProjectColumn = True
        End If

        Set tbl = ws.ListObjects("tbl_WBS")
        WBS_Onboarding_ApplyStructure ws, tbl, statuses
        gWBSOnboardingStructuralMutationCount = _
            gWBSOnboardingStructuralMutationCount + 1
    End If

    If Not tbl.DataBodyRange Is Nothing Then
        firstDataRow = tbl.DataBodyRange.Row
        If firstDataRow <> tbl.HeaderRowRange.Row + 1 Then
            Err.Raise vbObjectError + 2333, "Ensure_WBS_Onboarding_Guide", _
                PlanningMessageText_Format("WBS.ERROR.UNEXPECTED_FIRST_DATA_ROW", _
                    TextCatalog_Arguments("Row", CStr(firstDataRow)), TextCatalog_Arguments("Row", CStr(firstDataRow)))
        End If
    End If

    Application.ScreenUpdating = oldScreenUpdating
    Application.EnableEvents = oldEvents

    If applyLocalizedContent Then
        WBS_SetLanguage Settings_GetOwnerLanguage("WBS")
        WBS_Onboarding_ApplyLocalizedContent ws, tbl
    End If
    Exit Sub

ErrHandler:
    errorNumber = Err.Number
    errorSource = Err.Source
    errorDescription = Err.Description

    On Error Resume Next
    Application.EnableEvents = False
    If renamedLegacyProjectColumn Then
        Set tbl = ws.ListObjects("tbl_WBS")
        If WBS_TableHasSchemaColumn(tbl, VTS_COL_PROJECT) Then SchemaListColumn(tbl, VTS_TABLE_WBS, VTS_COL_PROJECT).Name = "Package"
    End If
    Application.ScreenUpdating = oldScreenUpdating
    Application.EnableEvents = oldEvents
    On Error GoTo 0

    Err.Raise errorNumber, errorSource, errorDescription

End Sub

'------------------------------------------------------------------------------
' FR: Resout la ligne d'onboarding relativement au vrai header de tbl_WBS.
' EN: Resolves the onboarding row relative to the physical tbl_WBS header.
'------------------------------------------------------------------------------
Private Function WBS_OnboardingHelpRow(ByVal tbl As ListObject) As Long

    If tbl Is Nothing Or tbl.HeaderRowRange.Row <= 1 Then
        Err.Raise vbObjectError + 2334, "WBS_OnboardingHelpRow", _
            PlanningMessageText_Format("WBS.ERROR.INVALID_ONBOARDING_ROW")
    End If
    WBS_OnboardingHelpRow = tbl.HeaderRowRange.Row - 1

End Function

Public Function WBS_OnboardingHarness_CheckStructure() As String
    Dim ws As Worksheet
    Dim tbl As ListObject

    Set ws = ThisWorkbook.Worksheets("WBS")
    Set tbl = ws.ListObjects("tbl_WBS")
    WBS_OnboardingHarness_CheckStructure = CStr( _
        WBS_Onboarding_StructureIsCurrent(ws, tbl, WBS_Onboarding_BuildStatusMap()))
End Function

Public Function WBS_OnboardingHarness_CheckLocalized( _
    Optional ByVal languageCode As String = "") As String
    Dim ws As Worksheet
    Dim tbl As ListObject

    Set ws = ThisWorkbook.Worksheets("WBS")
    Set tbl = ws.ListObjects("tbl_WBS")
    If Trim$(languageCode) <> "" Then
        WBS_SetLanguage languageCode
    Else
        WBS_SetLanguage Settings_GetOwnerLanguage("WBS")
    End If
    WBS_OnboardingHarness_CheckLocalized = CStr( _
        WBS_Onboarding_LocalizedContentIsCurrent(ws, tbl))
End Function

Public Function WBS_OnboardingHarness_ApplyLocalized() As String
    Dim ws As Worksheet
    Dim tbl As ListObject

    Set ws = ThisWorkbook.Worksheets("WBS")
    Set tbl = ws.ListObjects("tbl_WBS")
    WBS_SetLanguage Settings_GetOwnerLanguage("WBS")
    WBS_Onboarding_ApplyLocalizedContent ws, tbl
    WBS_OnboardingHarness_ApplyLocalized = "PASS"
End Function

'------------------------------------------------------------------------------
' FR: Verifie en lecture seule que la ligne d'onboarding utilise la structure canonique.
' EN: Read-only check that the onboarding row uses the canonical structure.
'------------------------------------------------------------------------------
Private Function WBS_Onboarding_StructureIsCurrent( _
    ByVal ws As Worksheet, _
    ByVal tbl As ListObject, _
    ByVal statuses As Object) As Boolean

    Dim key As Variant
    Dim cell As Range


    On Error GoTo NotCurrent

    If ws Is Nothing Or tbl Is Nothing Or statuses Is Nothing Then Exit Function
    If tbl.HeaderRowRange.Row <= 1 Then Exit Function
    If Not tbl.DataBodyRange Is Nothing Then
        If tbl.DataBodyRange.Row <> tbl.HeaderRowRange.Row + 1 Then Exit Function
    End If
    If Abs(CDbl(ws.Rows(WBS_OnboardingHelpRow(tbl)).RowHeight) - _
        WBS_ONBOARDING_ROW_HEIGHT) > 0.05 Then Exit Function

    For Each key In statuses.Keys
        Set cell = ws.Cells(WBS_OnboardingHelpRow(tbl), _
            SchemaListColumn(tbl, VTS_TABLE_WBS, CStr(key)).Range.Column)


        If CStr(cell.Font.Name) <> WBS_ONBOARDING_FONT_NAME Then Exit Function
        If Abs(CDbl(cell.Font.Size) - WBS_ONBOARDING_FONT_SIZE) > 0.01 Then Exit Function
        If Not CBool(cell.Font.Bold) Then Exit Function
        If CBool(cell.Font.Italic) Then Exit Function
        If cell.Interior.Pattern <> xlSolid Then Exit Function
        If cell.Interior.Color <> RGB(255, 255, 255) Then Exit Function
        If cell.HorizontalAlignment <> xlCenter Then Exit Function
        If cell.VerticalAlignment <> xlCenter Then Exit Function
        If Not CBool(cell.WrapText) Then Exit Function
        If CBool(cell.ShrinkToFit) Then Exit Function
        If cell.Borders(xlEdgeBottom).LineStyle <> xlContinuous Then Exit Function
        If cell.Borders(xlEdgeBottom).Color <> RGB(0, 0, 0) Then Exit Function
        If cell.Borders(xlEdgeBottom).Weight <> xlThin Then Exit Function


    Next key

    WBS_Onboarding_StructureIsCurrent = True
    Exit Function

NotCurrent:
    WBS_Onboarding_StructureIsCurrent = False

End Function

'------------------------------------------------------------------------------
' FR: Repare la ligne d'onboarding lorsqu'un ecart structurel est demontre.
' EN: Repairs the onboarding row when a structural mismatch is proven.
'------------------------------------------------------------------------------
Private Sub WBS_Onboarding_ApplyStructure( _
    ByVal ws As Worksheet, _
    ByVal tbl As ListObject, _
    ByVal statuses As Object)

    Dim helpRange As Range
    Dim key As Variant
    Dim cell As Range

    Set helpRange = ws.Range( _
        ws.Cells(WBS_OnboardingHelpRow(tbl), tbl.Range.Column), _
        ws.Cells(WBS_OnboardingHelpRow(tbl), _
            tbl.Range.Column + tbl.ListColumns.Count - 1))

    helpRange.ClearContents
    With helpRange
        .Font.Name = WBS_ONBOARDING_FONT_NAME
        .Font.Size = WBS_ONBOARDING_FONT_SIZE
        .Font.Bold = True
        .Font.Italic = False
        .Font.Color = RGB(0, 0, 0)
        .Interior.Pattern = xlSolid
        .Interior.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .ShrinkToFit = False
        With .Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Color = RGB(0, 0, 0)
            .Weight = xlThin
        End With
    End With

    ws.Rows(WBS_OnboardingHelpRow(tbl)).RowHeight = WBS_ONBOARDING_ROW_HEIGHT

    For Each key In statuses.Keys
        Set cell = ws.Cells(WBS_OnboardingHelpRow(tbl), _
            SchemaListColumn(tbl, VTS_TABLE_WBS, CStr(key)).Range.Column)
        cell.Value = CStr(statuses(key))
        WBS_Onboarding_FormatStatusCell cell
    Next key

End Sub

'------------------------------------------------------------------------------
' FR: Reconstruit les statuts, le Quick Start et les commentaires dans la langue active.
' EN: Rebuilds statuses, the Quick Start and comments in the active language.
'------------------------------------------------------------------------------
Private Sub WBS_Onboarding_ApplyLocalizedContent( _
    ByVal ws As Worksheet, _
    Optional ByVal tbl As ListObject = Nothing)

    Dim comments As Object
    Dim statuses As Object
    Dim key As Variant
    Dim statusCell As Range
    Dim listColumn As ListColumn
    Dim helpCell As Range
    Dim headerCell As Range
    Dim localizedPair As Variant
    Dim expectedText As String
    Dim actualText As String
    Dim hasThreadedComment As Boolean
    Dim threadedCommentsSupported As Boolean
    Dim languageIndex As Long
    Dim columnKeys As Variant
    Dim oldEvents As Boolean
    Dim oldScreenUpdating As Boolean
    Dim errorNumber As Long
    Dim errorSource As String
    Dim errorDescription As String

    If ws Is Nothing Then Exit Sub

    oldEvents = Application.EnableEvents
    oldScreenUpdating = Application.ScreenUpdating
    On Error GoTo ErrHandler

    If tbl Is Nothing Then Set tbl = ws.ListObjects("tbl_WBS")
    Set comments = WBS_Onboarding_BuildHelpCommentMap()
    Set statuses = WBS_Onboarding_BuildStatusMap()
    threadedCommentsSupported = ExcelCompatibility_SupportsThreadedComments()
    languageIndex = IIf(WBS_CurrentLanguage() = "FR", 0, 1)
    columnKeys = SchemaColumnKeys(VTS_TABLE_WBS)

    Application.EnableEvents = False
    Application.ScreenUpdating = False
    gWBSOnboardingLocalizedRebuildCount = _
        gWBSOnboardingLocalizedRebuildCount + 1

    For Each key In statuses.Keys
        Set statusCell = ws.Cells(WBS_OnboardingHelpRow(tbl), _
            SchemaListColumn(tbl, VTS_TABLE_WBS, CStr(key)).Range.Column)
        If CStr(statusCell.Value2) <> CStr(statuses(key)) Then
            statusCell.Value = CStr(statuses(key))
        End If
        WBS_Onboarding_FormatStatusCell statusCell
    Next key

    For Each key In columnKeys
        Set listColumn = SchemaListColumn(tbl, VTS_TABLE_WBS, CStr(key))
        Set helpCell = ws.Cells(WBS_OnboardingHelpRow(tbl), listColumn.Range.Column)
        Set headerCell = listColumn.Range.Cells(1, 1)

        If WBS_Onboarding_CellHasAnyComment(headerCell) Then
            WBS_Onboarding_ClearCellComments headerCell
        End If

        If comments.Exists(CStr(key)) Then
            localizedPair = comments(CStr(key))
            expectedText = CStr(localizedPair(languageIndex))
            If threadedCommentsSupported Then
                actualText = WBS_Onboarding_ThreadedCommentText( _
                    helpCell, hasThreadedComment)
                If WBS_Onboarding_CellHasLegacyComment(helpCell) Or _
                   Not hasThreadedComment Or _
                   WBS_Onboarding_NormalizeComparisonText(actualText) <> _
                        WBS_Onboarding_NormalizeComparisonText(expectedText) Then
                    WBS_Onboarding_ClearCellComments helpCell
                    WBS_Onboarding_AddThreadedComment helpCell, expectedText
                End If
            Else
                actualText = WBS_Onboarding_LegacyCommentText( _
                    helpCell, hasThreadedComment)
                If Not hasThreadedComment Or _
                   WBS_Onboarding_NormalizeComparisonText(actualText) <> _
                        WBS_Onboarding_NormalizeComparisonText(expectedText) Then
                    WBS_Onboarding_ClearCellComments helpCell
                    WBS_Onboarding_AddLegacyComment helpCell, expectedText
                End If
            End If
        ElseIf WBS_Onboarding_CellHasAnyComment(helpCell) Then
            WBS_Onboarding_ClearCellComments helpCell
        End If
    Next key

    WBS_Onboarding_WriteQuickStart ws

    Application.ScreenUpdating = oldScreenUpdating
    Application.EnableEvents = oldEvents
    Exit Sub

ErrHandler:
    errorNumber = Err.Number
    errorSource = Err.Source
    errorDescription = Err.Description

    On Error Resume Next
    Application.ScreenUpdating = oldScreenUpdating
    Application.EnableEvents = oldEvents
    On Error GoTo 0

    Err.Raise errorNumber, errorSource, errorDescription

End Sub

'------------------------------------------------------------------------------
' FR: Verifie en lecture seule que les contenus localises correspondent au catalogue actif.
' EN: Read-only check that localized content matches the active catalog.
'------------------------------------------------------------------------------
Private Function WBS_Onboarding_LocalizedContentIsCurrent( _
    ByVal ws As Worksheet, _
    Optional ByVal tbl As ListObject = Nothing) As Boolean

    Dim comments As Object
    Dim statuses As Object
    Dim key As Variant
    Dim statusCell As Range
    Dim requiredFont As Font
    Dim suffixFont As Font
    Dim requiredLabel As String
    Dim listColumn As ListColumn
    Dim helpCell As Range
    Dim headerCell As Range
    Dim noteArea As Range
    Dim localizedPair As Variant
    Dim expectedText As String
    Dim actualText As String
    Dim hasThreadedComment As Boolean
    Dim threadedCommentsSupported As Boolean
    Dim languageIndex As Long
    Dim columnKeys As Variant

    gWBSOnboardingLocalizedCheckCount = _
        gWBSOnboardingLocalizedCheckCount + 1

    On Error GoTo NotCurrent
    If ws Is Nothing Then Exit Function
    If tbl Is Nothing Then Set tbl = ws.ListObjects("tbl_WBS")

    Set comments = WBS_Onboarding_BuildHelpCommentMap()
    Set statuses = WBS_Onboarding_BuildStatusMap()
    threadedCommentsSupported = ExcelCompatibility_SupportsThreadedComments()
    requiredLabel = WBS_Onboarding_RequiredLabel()
    If comments.Count <> 37 Then Exit Function
    If Not comments.Exists(VTS_COL_PROJECT) Then Exit Function
    If Not comments.Exists(VTS_COL_LONGEST_PATH) Then Exit Function
    If Not comments.Exists(VTS_COL_LONGEST_PATH_REX) Then Exit Function
    If Not comments.Exists(VTS_COL_DEADLINE_FLOAT) Then Exit Function

    languageIndex = IIf(WBS_CurrentLanguage() = "FR", 0, 1)
    columnKeys = SchemaColumnKeys(VTS_TABLE_WBS)
    Set noteArea = ws.Range("O1:R2")

    If Not CBool(noteArea.MergeCells) Then Exit Function
    If ws.Range("O1").MergeArea.Address(False, False) <> "O1:R2" Then Exit Function
    If WBS_Onboarding_NormalizeComparisonText(CStr(ws.Range("O1").Value2)) <> _
        WBS_Onboarding_NormalizeComparisonText(WBS_Onboarding_QuickStartText()) Then Exit Function
    If noteArea.HorizontalAlignment <> xlLeft Then Exit Function
    If noteArea.VerticalAlignment <> xlTop Then Exit Function
    If Not CBool(noteArea.WrapText) Then Exit Function

    For Each key In statuses.Keys
        Set statusCell = ws.Cells(WBS_OnboardingHelpRow(tbl), _
            SchemaListColumn(tbl, VTS_TABLE_WBS, CStr(key)).Range.Column)
        actualText = CStr(statusCell.Value2)
        If actualText <> CStr(statuses(key)) Then Exit Function

        If Left$(actualText, Len(requiredLabel)) = requiredLabel Then
            Set requiredFont = statusCell.Characters(1, Len(requiredLabel)).Font
            If requiredFont.Color <> RGB(192, 0, 0) Then Exit Function
            If Not CBool(requiredFont.Bold) Then Exit Function

            If Len(actualText) > Len(requiredLabel) Then
                Set suffixFont = statusCell.Characters( _
                    Len(requiredLabel) + 1, _
                    Len(actualText) - Len(requiredLabel)).Font
                If suffixFont.Color <> RGB(0, 0, 0) Then Exit Function
            End If
        ElseIf statusCell.Font.Color <> RGB(0, 0, 0) Then
            Exit Function
        End If
    Next key

    For Each key In columnKeys
        Set listColumn = SchemaListColumn(tbl, VTS_TABLE_WBS, CStr(key))
        Set helpCell = ws.Cells(WBS_OnboardingHelpRow(tbl), listColumn.Range.Column)
        Set headerCell = listColumn.Range.Cells(1, 1)

        If WBS_Onboarding_CellHasAnyComment(headerCell) Then Exit Function

        If threadedCommentsSupported Then
            If WBS_Onboarding_CellHasLegacyComment(helpCell) Then Exit Function
            actualText = WBS_Onboarding_ThreadedCommentText( _
                helpCell, hasThreadedComment)
        Else
            actualText = WBS_Onboarding_LegacyCommentText( _
                helpCell, hasThreadedComment)
        End If

        If comments.Exists(CStr(key)) Then
            If Not hasThreadedComment Then Exit Function
            localizedPair = comments(CStr(key))
            expectedText = CStr(localizedPair(languageIndex))
            If WBS_Onboarding_NormalizeComparisonText(actualText) <> _
                WBS_Onboarding_NormalizeComparisonText(expectedText) Then Exit Function
        ElseIf hasThreadedComment Then
            Exit Function
        End If
    Next key

    WBS_Onboarding_LocalizedContentIsCurrent = True
    Exit Function

NotCurrent:
    WBS_Onboarding_LocalizedContentIsCurrent = False

End Function

'------------------------------------------------------------------------------
' FR: Retourne le Quick Start canonique dans la langue runtime WBS.
' EN: Returns the canonical Quick Start in the WBS runtime language.
'------------------------------------------------------------------------------
Private Function WBS_Onboarding_QuickStartText() As String

    WBS_Onboarding_QuickStartText = TextCatalog_Get("WBS.ONBOARDING.QUICK_START", WBS_CurrentLanguage())

End Function

'------------------------------------------------------------------------------
' FR: Normalise uniquement les fins de ligne avant une comparaison de contenu.
' EN: Normalizes line endings only before comparing content.
'------------------------------------------------------------------------------
Private Function WBS_Onboarding_NormalizeComparisonText( _
    ByVal value As String) As String

    value = Replace(value, vbCrLf, vbLf)
    value = Replace(value, vbCr, vbLf)
    WBS_Onboarding_NormalizeComparisonText = value

End Function

'------------------------------------------------------------------------------
' FR: Indique si une cellule possede un commentaire classique.
' EN: Reports whether a cell owns a legacy comment.
'------------------------------------------------------------------------------
Private Function WBS_Onboarding_CellHasLegacyComment( _
    ByVal cell As Range) As Boolean

    Dim legacyComment As Object

    On Error Resume Next
    Set legacyComment = cell.Comment
    On Error GoTo 0
    WBS_Onboarding_CellHasLegacyComment = Not legacyComment Is Nothing

End Function

'------------------------------------------------------------------------------
' FR: Lit une Note Excel classique sans utiliser l'API threaded.
' EN: Reads a classic Excel Note without using the threaded API.
'------------------------------------------------------------------------------
Private Function WBS_Onboarding_LegacyCommentText( _
    ByVal cell As Range, _
    ByRef commentExists As Boolean) As String

    Dim legacyComment As Object

    commentExists = False
    On Error Resume Next
    Set legacyComment = cell.Comment
    On Error GoTo 0
    If legacyComment Is Nothing Then Exit Function

    commentExists = True
    WBS_Onboarding_LegacyCommentText = CStr( _
        CallByName(legacyComment, "Text", VbMethod))

End Function

'------------------------------------------------------------------------------
' FR: Lit un commentaire threaded sans modifier la cellule.
' EN: Reads a threaded comment without mutating the cell.
'------------------------------------------------------------------------------
Private Function WBS_Onboarding_ThreadedCommentText( _
    ByVal cell As Range, _
    ByRef commentExists As Boolean) As String

    Dim threadedComment As Object

    commentExists = False
    If Not ExcelCompatibility_SupportsThreadedComments() Then Exit Function

    gWBSOnboardingThreadedReadCount = gWBSOnboardingThreadedReadCount + 1
    On Error Resume Next
    Set threadedComment = cell.CommentThreaded
    On Error GoTo 0

    If threadedComment Is Nothing Then Exit Function
    commentExists = True
    WBS_Onboarding_ThreadedCommentText = CStr( _
        CallByName(threadedComment, "Text", VbMethod))

End Function

'------------------------------------------------------------------------------
' FR: Indique si une cellule possede un commentaire classique ou threaded.
' EN: Reports whether a cell owns a legacy or threaded comment.
'------------------------------------------------------------------------------
Private Function WBS_Onboarding_CellHasAnyComment( _
    ByVal cell As Range) As Boolean

    Dim hasThreadedComment As Boolean
    Dim ignoredText As String

    ignoredText = WBS_Onboarding_ThreadedCommentText(cell, hasThreadedComment)
    WBS_Onboarding_CellHasAnyComment = _
        hasThreadedComment Or WBS_Onboarding_CellHasLegacyComment(cell)

End Function

'------------------------------------------------------------------------------
' FR: Supprime les commentaires d'une cellule uniquement lorsqu'une reparation est requise.
' EN: Clears cell comments only when a repair is required.
'------------------------------------------------------------------------------
Private Sub WBS_Onboarding_ClearCellComments(ByVal cell As Range)

    Dim legacyComment As Object
    Dim threadedComment As Object

    On Error Resume Next
    Set legacyComment = cell.Comment
    If Not legacyComment Is Nothing Then legacyComment.Delete
    If ExcelCompatibility_SupportsThreadedComments() Then
        Set threadedComment = cell.CommentThreaded
        If Not threadedComment Is Nothing Then
            gWBSOnboardingThreadedClearCount = gWBSOnboardingThreadedClearCount + 1
            CallByName threadedComment, "Delete", VbMethod
        End If
    End If
    On Error GoTo 0

End Sub

'------------------------------------------------------------------------------
' FR: Ajoute un commentaire threaded uniquement lorsque l'API Excel est disponible.
' EN: Adds a threaded comment only when the Excel API is available.
'------------------------------------------------------------------------------
Private Sub WBS_Onboarding_AddThreadedComment( _
    ByVal cell As Range, _
    ByVal commentText As String)

    If Not ExcelCompatibility_SupportsThreadedComments() Then Exit Sub

    gWBSOnboardingThreadedAddCount = gWBSOnboardingThreadedAddCount + 1
    CallByName cell, "AddCommentThreaded", VbMethod, commentText

End Sub

'------------------------------------------------------------------------------
' FR: Ajoute une Note Excel classique sur les versions sans commentaires threaded.
' EN: Adds a classic Excel Note on versions without threaded comments.
'------------------------------------------------------------------------------
Private Sub WBS_Onboarding_AddLegacyComment( _
    ByVal cell As Range, _
    ByVal commentText As String)

    cell.AddComment commentText

End Sub

'------------------------------------------------------------------------------
' FR: Construit le catalogue canonique des aides de colonnes WBS en FR et EN.
' EN: Builds the canonical FR and EN catalog of WBS column help text.
'------------------------------------------------------------------------------
Private Function WBS_Onboarding_BuildHelpCommentMap() As Object

    Dim comments As Object
    Dim columnKeys As Variant
    Dim columnKey As Variant

    Set comments = CreateObject("Scripting.Dictionary")
    comments.CompareMode = vbTextCompare
    columnKeys = SchemaColumnKeys(VTS_TABLE_WBS)

    For Each columnKey In columnKeys
        If CStr(columnKey) <> VTS_COL_CAL Then
            comments(CStr(columnKey)) = Array( _
                TextCatalogWBS_TutorialText(CStr(columnKey), TEXT_LANGUAGE_FR), _
                TextCatalogWBS_TutorialText(CStr(columnKey), TEXT_LANGUAGE_EN))
        End If
    Next columnKey

    Set WBS_Onboarding_BuildHelpCommentMap = comments

End Function

' FR: Definit le statut fonctionnel de chaque colonne canonique WBS.
' EN: Defines the functional status of every canonical WBS column.
'------------------------------------------------------------------------------
Private Function WBS_Onboarding_BuildStatusMap() As Object

    Dim statuses As Object
    Dim requiredLabel As String
    Dim optionalLabel As String
    Dim calculatedLabel As String

    requiredLabel = WBS_Onboarding_RequiredLabel()
    optionalLabel = TextCatalog_Get("WBS.ONBOARDING.OPTIONAL", WBS_CurrentLanguage())
    calculatedLabel = TextCatalog_Get("WBS.ONBOARDING.CALCULATED", WBS_CurrentLanguage())

    Set statuses = CreateObject("Scripting.Dictionary")
    statuses.CompareMode = vbTextCompare

    statuses.Add VTS_COL_ID, requiredLabel
    statuses.Add VTS_COL_WBS, requiredLabel
    statuses.Add VTS_COL_TASK_NAME, requiredLabel
    statuses.Add VTS_COL_TASK_DESCRIPTION, optionalLabel
    statuses.Add VTS_COL_DISCIPLINE, optionalLabel
    statuses.Add VTS_COL_SUPPLIER, optionalLabel
    statuses.Add VTS_COL_PROJECT, optionalLabel
    statuses.Add VTS_COL_CAL, optionalLabel
    statuses.Add VTS_COL_TASK_TYPE, optionalLabel
    statuses.Add VTS_COL_S, "O"
    statuses.Add VTS_COL_COMMENTS, optionalLabel
    statuses.Add VTS_COL_PREDECESSORS_WBS, requiredLabel & " G*"
    statuses.Add VTS_COL_WEIGHT_PERCENT, requiredLabel & " S"
    statuses.Add VTS_COL_PROGRESS_PERCENT, requiredLabel & " S"
    statuses.Add VTS_COL_BASELINE_START, requiredLabel & " G*"
    statuses.Add VTS_COL_BASELINE_DURATION, requiredLabel & " G*"
    statuses.Add VTS_COL_BASELINE_FINISH, calculatedLabel
    statuses.Add VTS_COL_ACTUAL_START, optionalLabel
    statuses.Add VTS_COL_ACTUAL_FINISH, optionalLabel
    statuses.Add VTS_COL_ACTUAL_DURATION, calculatedLabel
    statuses.Add VTS_COL_FORECAST_START, optionalLabel
    statuses.Add VTS_COL_FORECAST_FINISH, optionalLabel
    statuses.Add VTS_COL_CALCULATED_START, calculatedLabel
    statuses.Add VTS_COL_CALCULATED_FINISH, calculatedLabel
    statuses.Add VTS_COL_CALCULATED_DURATION, calculatedLabel
    statuses.Add VTS_COL_START_VARIANCE, calculatedLabel
    statuses.Add VTS_COL_FINISH_VARIANCE, calculatedLabel
    statuses.Add VTS_COL_DURATION_VARIANCE, calculatedLabel
    statuses.Add VTS_COL_DRIVING_LOGIC, calculatedLabel
    statuses.Add VTS_COL_CRITICAL_PATH, calculatedLabel
    statuses.Add VTS_COL_CRITICAL_PATH_REX, calculatedLabel
    statuses.Add VTS_COL_LONGEST_PATH, calculatedLabel
    statuses.Add VTS_COL_LONGEST_PATH_REX, calculatedLabel
    statuses.Add VTS_COL_TOTAL_FLOAT, calculatedLabel
    statuses.Add VTS_COL_FREE_FLOAT, calculatedLabel
    statuses.Add VTS_COL_TOTAL_FLOAT_REX, calculatedLabel
    statuses.Add VTS_COL_FREE_FLOAT_REX, calculatedLabel
    statuses.Add VTS_COL_DEADLINE_FLOAT, calculatedLabel

    Set WBS_Onboarding_BuildStatusMap = statuses

End Function

'------------------------------------------------------------------------------
' FR: Retourne le libelle Requis dans la langue runtime WBS.
' EN: Returns the Required label in the WBS runtime language.
'------------------------------------------------------------------------------
Private Function WBS_Onboarding_RequiredLabel() As String

    WBS_Onboarding_RequiredLabel = TextCatalog_Get("WBS.ONBOARDING.REQUIRED", WBS_CurrentLanguage())

End Function

'------------------------------------------------------------------------------
' FR: Applique le style compact et met uniquement le statut Requis en evidence.
' EN: Applies the compact style and highlights only the Required status.
'------------------------------------------------------------------------------
Private Sub WBS_Onboarding_FormatStatusCell(ByVal cell As Range)

    Dim requiredStart As Long
    Dim requiredLabel As String

    If cell Is Nothing Then Exit Sub
    requiredLabel = WBS_Onboarding_RequiredLabel()

    cell.Font.Color = RGB(0, 0, 0)
    cell.Font.Bold = True

    requiredStart = InStr(1, CStr(cell.Value2), requiredLabel, vbTextCompare)
    If requiredStart > 0 Then
        With cell.Characters(requiredStart, Len(requiredLabel)).Font
            .Color = RGB(192, 0, 0)
            .Bold = True
        End With
    End If

End Sub

'------------------------------------------------------------------------------
' FR: Affiche le guide général WBS dans O1:R2 en respectant la langue active.
' EN: Displays the general WBS guide in O1:R2 using the active language.
'------------------------------------------------------------------------------
Private Sub WBS_Onboarding_WriteQuickStart(ByVal ws As Worksheet)

    Dim noteArea As Range
    Dim noteText As String
    Dim requiredLabel As String
    Dim requiredStart As Long
    Dim textChanged As Boolean

    Set noteArea = ws.Range("O1:R2")
    noteText = WBS_Onboarding_QuickStartText()

    requiredLabel = WBS_Onboarding_RequiredLabel()

    If Not CBool(noteArea.MergeCells) Or _
        ws.Range("O1").MergeArea.Address(False, False) <> "O1:R2" Then
        noteArea.UnMerge
        noteArea.ClearContents
        noteArea.Merge
        textChanged = True
    End If

    If WBS_Onboarding_NormalizeComparisonText(CStr(ws.Range("O1").Value2)) <> _
        WBS_Onboarding_NormalizeComparisonText(noteText) Then
        noteArea.Value = noteText
        textChanged = True
    End If

    If noteArea.HorizontalAlignment <> xlLeft Then _
        noteArea.HorizontalAlignment = xlLeft
    If noteArea.VerticalAlignment <> xlTop Then _
        noteArea.VerticalAlignment = xlTop
    If Not CBool(noteArea.WrapText) Then noteArea.WrapText = True
    If CBool(noteArea.ShrinkToFit) Then noteArea.ShrinkToFit = False

    If textChanged Then
        requiredStart = InStr(1, noteText, requiredLabel, vbTextCompare)
        If requiredStart > 0 Then
            With noteArea.Characters(requiredStart, Len(requiredLabel)).Font
                .Color = RGB(192, 0, 0)
                .Bold = True
            End With
        End If
    End If

End Sub

'------------------------------------------------------------------------------
' FR: Verifie ou cree WBS Main Buttons si necessaire.
' EN: Ensures or creates WBS Main Buttons when needed.
'------------------------------------------------------------------------------
Public Sub Ensure_WBS_Main_Buttons()

    Dim ws As Worksheet

    Set ws = ThisWorkbook.Worksheets("WBS")
    WBS_SetLanguage Settings_GetOwnerLanguage("WBS")

    Ensure_WBS_Onboarding_Guide ws, False

    On Error Resume Next
    ws.Shapes("btn_WBS_FullReset").Delete
    On Error GoTo 0

    Ensure_WBS_TaskType_Input_Setup ws
    Ensure_CriticalPathMode_Toggle

    'Sheet command buttons are presented exclusively by the Ribbon.
    Dim commandName As Variant
    Dim commandShape As Shape
    For Each commandName In Array("btn_WBS_Planning", "btn_WBS_Gantt", "btn_WBS_SCurve", "btn_WBS_ForcedPlanning", "btn_WBS_Full", "btn_WBS_ResetPlanning")
        Set commandShape = Nothing
        On Error Resume Next
        Set commandShape = ws.Shapes(CStr(commandName))
        On Error GoTo 0
        If Not commandShape Is Nothing Then commandShape.Delete
    Next commandName

    WBS_SetCellValueIfDifferent ws.Range("L1"), _
        TextCatalog_Get("WBS.SUMMARY.BASELINE_FINISH", WBS_CurrentLanguage())
    WBS_SetCellValueIfDifferent ws.Range("M1"), _
        TextCatalog_Get("WBS.SUMMARY.CALCULATED_FINISH", WBS_CurrentLanguage())
    WBS_SetCellValueIfDifferent ws.Range("N1"), _
        TextCatalog_Get("WBS.SUMMARY.DELAY", WBS_CurrentLanguage())

End Sub

'------------------------------------------------------------------------------
' FR: Actualise WBS Apply Language sans modifier les regles metier qui produisent les donnees.
' EN: Refreshes WBS Apply Language without changing the business rules that produce the data.
'------------------------------------------------------------------------------

Public Sub WBS_ApplyLanguage(Optional ByVal languageCode As String = "")

    Dim ws As Worksheet
    Dim tbl As ListObject

    On Error GoTo ErrHandler

    Set ws = ThisWorkbook.Worksheets("WBS")

    If Trim$(languageCode) <> "" Then
        WBS_SetLanguage languageCode
    Else
        EnsureWBSLanguageInitialized
    End If

    WBS_SetShapeText ws, "btn_WBS_Planning", TextCatalog_Get(TXT_WBS_UPDATE_PLANNING_LABEL, WBS_CurrentLanguage())
    WBS_SetShapeText ws, "btn_WBS_Gantt", TextCatalog_Get(TXT_WBS_UPDATE_GANTT_LABEL, WBS_CurrentLanguage())
    WBS_SetShapeText ws, "btn_WBS_SCurve", TextCatalog_Get(TXT_WBS_UPDATE_SCURVE_LABEL, WBS_CurrentLanguage())
    WBS_SetShapeText ws, "btn_WBS_ForcedPlanning", TextCatalog_Get(TXT_WBS_FORCED_UPDATE_LABEL, WBS_CurrentLanguage())
    WBS_SetShapeText ws, "btn_WBS_Full", TextCatalog_Get(TXT_WBS_FULL_UPDATE_LABEL, WBS_CurrentLanguage())
    WBS_SetShapeText ws, "btn_WBS_ResetPlanning", TextCatalog_Get(TXT_WBS_RESET_PLANNING_LABEL, WBS_CurrentLanguage())

    WBS_SetCellValueIfDifferent ws.Range("L1"), _
        TextCatalog_Get("WBS.SUMMARY.BASELINE_FINISH", WBS_CurrentLanguage())
    WBS_SetCellValueIfDifferent ws.Range("M1"), _
        TextCatalog_Get("WBS.SUMMARY.CALCULATED_FINISH", WBS_CurrentLanguage())
    WBS_SetCellValueIfDifferent ws.Range("N1"), _
        TextCatalog_Get("WBS.SUMMARY.DELAY", WBS_CurrentLanguage())

    Set tbl = ws.ListObjects("tbl_WBS")
    WBS_ApplyLocalizedInputValidations tbl
    If Not WBS_Onboarding_LocalizedContentIsCurrent(ws, tbl) Then
        WBS_Onboarding_ApplyLocalizedContent ws, tbl
    End If

    Exit Sub

ErrHandler:
    Err.Raise Err.Number, "WBS_ApplyLanguage", Err.Description

End Sub

'------------------------------------------------------------------------------
' FR: Reinitialise les compteurs runtime utilises uniquement par le harnais WBS.
' EN: Resets runtime counters used only by the WBS proof harness.
'------------------------------------------------------------------------------
Public Sub WBS_OnboardingInstrumentation_Reset()

    gWBSOnboardingLocalizedCheckCount = 0
    gWBSOnboardingLocalizedRebuildCount = 0
    gWBSOnboardingStructuralMutationCount = 0
    gWBSOnboardingThreadedReadCount = 0
    gWBSOnboardingThreadedAddCount = 0
    gWBSOnboardingThreadedClearCount = 0

End Sub

'------------------------------------------------------------------------------
' FR: Expose un snapshot des compteurs runtime au harnais non interactif.
' EN: Exposes a runtime counter snapshot to the noninteractive proof harness.
'------------------------------------------------------------------------------
Public Function WBS_OnboardingInstrumentation_Snapshot() As String

    WBS_OnboardingInstrumentation_Snapshot = _
        "checks=" & CStr(gWBSOnboardingLocalizedCheckCount) & _
        ";rebuilds=" & CStr(gWBSOnboardingLocalizedRebuildCount) & _
        ";structural=" & CStr(gWBSOnboardingStructuralMutationCount) & _
        ";threadedReads=" & CStr(gWBSOnboardingThreadedReadCount) & _
        ";threadedAdds=" & CStr(gWBSOnboardingThreadedAddCount) & _
        ";threadedClears=" & CStr(gWBSOnboardingThreadedClearCount) & _
        ";threadedSupport=" & ExcelCompatibility_ThreadedCommentsProbeStatus()

End Function

'------------------------------------------------------------------------------
' FR: Ecrit une valeur de cellule uniquement lorsqu'elle differe.
' EN: Writes a cell value only when it differs.
'------------------------------------------------------------------------------
Private Sub WBS_SetCellValueIfDifferent( _
    ByVal target As Range, _
    ByVal expectedValue As String)

    If target Is Nothing Then Exit Sub
    If CStr(target.Value2) <> expectedValue Then target.Value = expectedValue

End Sub

'------------------------------------------------------------------------------
' FR: Active ou initialise WBS Set Language dans l'etat runtime du composant.
' EN: Activates or initializes WBS Set Language in the component runtime state.
'------------------------------------------------------------------------------

Public Sub WBS_SetLanguage(ByVal languageCode As String)

    Select Case UCase$(Trim$(languageCode))
        Case "FR"
            gWBSLanguage = "FR"
        Case "EN"
            gWBSLanguage = "EN"
        Case Else
            gWBSLanguage = "EN"
    End Select

End Sub

'------------------------------------------------------------------------------
' FR: Retourne la valeur WBS Current Language sans modifier les donnees d'entree.
' EN: Returns the WBS Current Language value without mutating input data.
'------------------------------------------------------------------------------

Public Function WBS_CurrentLanguage() As String

    EnsureWBSLanguageInitialized
    WBS_CurrentLanguage = gWBSLanguage

End Function

'------------------------------------------------------------------------------
' FR: Verifie ou cree WBSLanguage Initialized si necessaire.
' EN: Ensures or creates WBSLanguage Initialized when needed.
'------------------------------------------------------------------------------
Private Sub EnsureWBSLanguageInitialized()

    If UCase$(Trim$(gWBSLanguage)) <> "FR" And UCase$(Trim$(gWBSLanguage)) <> "EN" Then
        gWBSLanguage = "EN"
    End If

End Sub

'------------------------------------------------------------------------------
' FR: Retourne la valeur WBS L sans modifier les donnees d'entree.
' EN: Returns the WBS L value without mutating input data.
'------------------------------------------------------------------------------


'------------------------------------------------------------------------------
' FR: Active ou initialise WBS Set Shape Text dans l'etat runtime du composant.
' EN: Activates or initializes WBS Set Shape Text in the component runtime state.
' FR - Effet de bord : cree ou met a jour des shapes Excel.
' EN - Side effect: creates or updates Excel shapes.
'------------------------------------------------------------------------------

Private Sub WBS_SetShapeText( _
    ByVal ws As Worksheet, _
    ByVal shapeName As String, _
    ByVal captionText As String)

    Dim shp As Shape

    If ws Is Nothing Then Exit Sub

    On Error Resume Next
    Set shp = ws.Shapes(shapeName)
    On Error GoTo 0

    If shp Is Nothing Then Exit Sub

    If CStr(shp.TextFrame2.TextRange.Text) <> captionText Then
        shp.TextFrame2.TextRange.Text = captionText
    End If

End Sub
'------------------------------------------------------------------------------
' FR: Cree Or Update WBSFloating Button pour le domaine WBS buttons.
' EN: Creates Or Update WBSFloating Button for the WBS buttons domain.
'------------------------------------------------------------------------------
Private Sub CreateOrUpdateWBSFloatingButton( _
    ByVal ws As Worksheet, _
    ByVal shpName As String, _
    ByVal captionText As String, _
    ByVal macroName As String, _
    ByVal leftPos As Double, _
    ByVal topPos As Double, _
    ByVal btnWidth As Double, _
    ByVal btnHeight As Double, _
    ByVal fillColor As Long)

    Dim shp As Shape

    On Error Resume Next
    Set shp = ws.Shapes(shpName)
    On Error GoTo 0

    If shp Is Nothing Then
        Set shp = ws.Shapes.AddShape( _
            msoShapeRoundedRectangle, _
            leftPos, _
            topPos, _
            btnWidth, _
            btnHeight)
        shp.Name = shpName
    End If

    shp.Left = leftPos
    shp.Top = topPos
    shp.Width = btnWidth
    shp.Height = btnHeight
    shp.OnAction = macroName
    shp.Placement = xlFreeFloating

    With shp.TextFrame2
        .TextRange.Text = captionText
        .TextRange.Font.Size = 10
        .TextRange.Font.Bold = msoTrue
        .TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .VerticalAnchor = msoAnchorMiddle
        .TextRange.ParagraphFormat.Alignment = msoAlignCenter
        .MarginLeft = 6
        .MarginRight = 6
        .MarginTop = 2
        .MarginBottom = 2
    End With

    shp.Fill.ForeColor.RGB = fillColor
    shp.Line.ForeColor.RGB = RGB(150, 150, 150)
    shp.Line.Weight = 1

End Sub

'------------------------------------------------------------------------------
' FR: Verifie ou cree WBS Task Type Input Setup si necessaire.
' EN: Ensures or creates WBS Task Type Input Setup when needed.
'------------------------------------------------------------------------------
Private Sub Ensure_WBS_TaskType_Input_Setup(ByVal ws As Worksheet)

    Dim tbl As ListObject
    Dim rng As Range
    Dim cell As Range
    Dim normalizedValue As String
    Dim rowIndex As Long
    Dim oldEvents As Boolean

    If ws Is Nothing Then Exit Sub

    On Error GoTo SafeExit

    On Error Resume Next
    Set tbl = ws.ListObjects("tbl_WBS")
    On Error GoTo SafeExit

    If tbl Is Nothing Then Exit Sub
    If tbl.DataBodyRange Is Nothing Then Exit Sub

    If Not WBS_TableHasSchemaColumn(tbl, VTS_COL_TASK_TYPE) Then
        Err.Raise vbObjectError + 2310, "Ensure_WBS_TaskType_Input_Setup", _
            PlanningMessageText_Format("WBS.ERROR.MISSING_TASK_TYPE_COLUMN")
    End If

    Set rng = SchemaListColumn(tbl, VTS_TABLE_WBS, VTS_COL_TASK_TYPE).DataBodyRange

    oldEvents = Application.EnableEvents
    Application.EnableEvents = False

    For Each cell In rng.Cells

        rowIndex = cell.Row - rng.Row + 1

        If Not WBSButtons_RowHasTaskIdentity(tbl, rowIndex) Then
            If Trim$(CStr(cell.value)) <> "" Then cell.ClearContents
        Else
            normalizedValue = Normalize_WBS_TaskType_Value(cell.value)

            If Trim$(CStr(cell.value)) = "" Then
                cell.value = "Task"
            ElseIf CStr(cell.value) <> normalizedValue Then
                cell.value = normalizedValue
            End If
        End If

    Next cell

    WBS_ApplyLocalizedInputValidations tbl

    rng.NumberFormat = "@"

SafeExit:
    Application.EnableEvents = oldEvents

    If Err.Number <> 0 Then
        WBSButtons_ShowConsoleError _
            "WBS.ERROR.TASK_TYPE_SETUP", _
            TextCatalog_Arguments("Details", Err.Description)
    End If

End Sub

'------------------------------------------------------------------------------
' FR: Normalise WBS Task Type Value dans un format exploitable.
' EN: Normalizes WBS Task Type Value into a usable format.
'------------------------------------------------------------------------------
Private Function Normalize_WBS_TaskType_Value(ByVal rawValue As Variant) As String

    Dim s As String

    s = UCase$(Trim$(CStr(rawValue)))

    Select Case s

        Case "", "TASK", "STANDARD", "NORMAL"
            Normalize_WBS_TaskType_Value = "Task"

        Case "MILESTONE", "MS", "JALON"
            Normalize_WBS_TaskType_Value = "Milestone"

        Case "LEVEL OF EFFORT", "LOE", "LEVEL-OF-EFFORT", "LEVEL_OF_EFFORT"
            Normalize_WBS_TaskType_Value = "Level of Effort"

        Case Else
            Normalize_WBS_TaskType_Value = CStr(rawValue)

    End Select

End Function

'------------------------------------------------------------------------------
' FR: Retourne la reference WBS Buttons Row Has Task IDentity sans modifier les donnees d'entree.
' EN: Returns the WBS Buttons Row Has Task IDentity reference without mutating input data.
'------------------------------------------------------------------------------

Private Function WBSButtons_RowHasTaskIdentity( _
    ByVal tbl As ListObject, _
    ByVal rowIndex As Long) As Boolean

    Dim idVal As String
    Dim wbsVal As String

    On Error GoTo SafeExit

    If tbl Is Nothing Then Exit Function
    If tbl.DataBodyRange Is Nothing Then Exit Function
    If rowIndex < 1 Or rowIndex > tbl.ListRows.Count Then Exit Function

    If WBS_TableHasSchemaColumn(tbl, VTS_COL_ID) Then
        idVal = Trim$(CStr(SchemaListColumn(tbl, VTS_TABLE_WBS, VTS_COL_ID).DataBodyRange.Cells(rowIndex, 1).value))
    End If

    If WBS_TableHasSchemaColumn(tbl, VTS_COL_WBS) Then
        wbsVal = Trim$(CStr(SchemaListColumn(tbl, VTS_TABLE_WBS, VTS_COL_WBS).DataBodyRange.Cells(rowIndex, 1).value))
    End If

    wbsVal = Replace$(wbsVal, ",", ".")
    WBSButtons_RowHasTaskIdentity = (idVal <> "" Or wbsVal <> "")

SafeExit:
End Function
'------------------------------------------------------------------------------
' FR: Retourne la reference WBS Table Has Column sans modifier les donnees d'entree.
' EN: Returns the WBS Table Has Column reference without mutating input data.
'------------------------------------------------------------------------------

Private Function WBS_TableHasSchemaColumn(ByVal tbl As ListObject, ByVal columnKey As String) As Boolean

    WBS_TableHasSchemaColumn = WBS_TableHasPhysicalColumn( _
        tbl, SchemaCurrentColumnTitle(VTS_TABLE_WBS, columnKey))

End Function

Private Function WBS_TableHasPhysicalColumn(ByVal tbl As ListObject, ByVal columnName As String) As Boolean

    Dim col As ListColumn

    On Error Resume Next
    Set col = tbl.ListColumns(columnName)
    On Error GoTo 0

    WBS_TableHasPhysicalColumn = Not col Is Nothing

End Function

'------------------------------------------------------------------------------
' FR: Projette la collection WBS Buttons Show Console Error vers l'interface autorisee par la politique runtime.
' EN: Projects the WBS Buttons Show Console Error collection to the UI allowed by runtime policy.
'------------------------------------------------------------------------------

Private Sub WBSButtons_ShowConsoleError( _
    ByVal messageKey As String, _
    Optional ByVal namedArguments As Object = Nothing)

    Dim consoleMessages As Collection

    Set consoleMessages = New Collection

    CalcBridge_AddConsoleMessage consoleMessages, _
        "STOP", _
        PlanningMessageText_Format(messageKey, namedArguments, namedArguments)

    CalcBridge_ShowPlanningConsole consoleMessages

End Sub













