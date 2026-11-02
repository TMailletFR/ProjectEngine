Attribute VB_Name = "mod_TextCatalogHarness"
Option Explicit

' Reusable release gate for the localization foundation.

Public Function TextCatalogHarness_RunPureGate() As String

    Dim i As Long
    Dim arguments As Object
    Dim formatted As String
    Dim diagnosticCountBefore As Long

    On Error GoTo Failed
    TextCatalog_ResetForTests

    TextCatalogHarness_Assert TextCatalog_ValidateAll() = "PASS", "Catalog validation failed: " & TextCatalog_ValidateAll()
    TextCatalogHarness_Assert TextCatalog_BuildCount() = 1, "Catalog was not built exactly once."
    TextCatalogHarness_Assert TextCatalog_Get(TXT_WBS_UPDATE_PLANNING_LABEL, "EN") = "Update" & vbCrLf & "Planning", "EN lookup failed."
    TextCatalogHarness_Assert TextCatalog_Get(TXT_WBS_UPDATE_PLANNING_LABEL, "FR") = "Mettre à jour" & vbCrLf & "Planning", "FR lookup failed."

    Set arguments = TextCatalog_Arguments("Error", "controlled failure")
    formatted = TextCatalog_Format(TXT_GANTT_SCENARIO_CREATE_ERROR, "EN", arguments)
    TextCatalogHarness_Assert InStr(1, formatted, "controlled failure", vbBinaryCompare) > 0, "Named placeholder replacement failed."

    formatted = TextCatalog_Get(TXT_WBS_UPDATE_PLANNING_LABEL, "DE")
    TextCatalogHarness_Assert formatted = "Update" & vbCrLf & "Planning", "Unknown language did not fall back to EN."
    formatted = TextCatalog_Get("UNKNOWN.PILOT.KEY", "FR")
    TextCatalogHarness_Assert formatted = "[missing text:UNKNOWN.PILOT.KEY/FR]", "Unknown key marker is incorrect."

    diagnosticCountBefore = TextCatalog_DiagnosticCount()
    Set arguments = TextCatalog_Arguments("Unexpected", "value")
    formatted = TextCatalog_Format(TXT_GANTT_SCENARIO_CREATE_ERROR, "FR", arguments)
    TextCatalogHarness_Assert InStr(1, formatted, "{Error}", vbBinaryCompare) > 0, "Missing placeholder token was hidden."
    TextCatalogHarness_Assert TextCatalog_DiagnosticCount() = diagnosticCountBefore + 2, "Placeholder diagnostics are incomplete."

    For i = 1 To 10000
        formatted = TextCatalog_Get(TXT_WBS_FULL_UPDATE_LABEL, "EN")
    Next i
    TextCatalogHarness_Assert TextCatalog_BuildCount() = 1, "Repeated lookups rebuilt the catalog."
    TextCatalogHarness_Assert TextCatalog_LookupCount() >= 10000, "Repeated lookups were not observed."

    TextCatalogHarness_RunPureGate = "PASS|builds=" & CStr(TextCatalog_BuildCount()) & _
        "|lookups=" & CStr(TextCatalog_LookupCount()) & _
        "|diagnostics=" & CStr(TextCatalog_DiagnosticCount())
    Exit Function

Failed:
    TextCatalogHarness_RunPureGate = "FAIL|" & CStr(Err.Number) & "|" & Err.Source & "|" & Err.Description

End Function

Public Function TextCatalogHarness_RunOwnerIsolationGate() As String

    Dim originalWBSLanguage As String
    Dim originalGanttLanguage As String
    Dim wbsText As String
    Dim ganttText As String
    Dim arguments As Object
    Dim errorNumber As Long
    Dim errorSource As String
    Dim errorDescription As String

    On Error GoTo Failed
    originalWBSLanguage = WBS_CurrentLanguage()
    originalGanttLanguage = Gantt_CurrentLanguage()

    WBS_SetLanguage "FR"
    Gantt_SetLanguage "EN"
    wbsText = TextCatalog_Get(TXT_WBS_RESET_PLANNING_LABEL, WBS_CurrentLanguage())
    Set arguments = TextCatalog_Arguments("Error", "X")
    ganttText = TextCatalog_Format(TXT_GANTT_SCENARIO_CREATE_ERROR, Gantt_CurrentLanguage(), arguments)

    TextCatalogHarness_Assert Left$(wbsText, Len("Réinitialiser")) = "Réinitialiser", "WBS owner language was not used."
    TextCatalogHarness_Assert Left$(ganttText, Len("Error while")) = "Error while", "Gantt owner language was not used."

    WBS_SetLanguage originalWBSLanguage
    Gantt_SetLanguage originalGanttLanguage
    TextCatalogHarness_RunOwnerIsolationGate = "PASS"
    Exit Function

Failed:
    errorNumber = Err.Number
    errorSource = Err.Source
    errorDescription = Err.Description
    On Error Resume Next
    WBS_SetLanguage originalWBSLanguage
    Gantt_SetLanguage originalGanttLanguage
    On Error GoTo 0
    TextCatalogHarness_RunOwnerIsolationGate = "FAIL|" & CStr(errorNumber) & "|" & errorSource & "|" & errorDescription

End Function

Public Function TextCatalogHarness_RunMission54_6DeltaGate() As String

    Dim predecessorWbs As String
    Dim linkType As String
    Dim lagValue As Long
    Dim rawToken As String
    Dim englishError As String
    Dim frenchError As Variant
    Dim links As Collection
    Dim wbsToId As Object
    Dim dummyArguments As Object
    Dim bootstrapDescription As String

    On Error GoTo Failed

    TextCatalogHarness_Assert _
        TextCatalog_Get("CONSOLE.WINDOW.TITLE", TEXT_LANGUAGE_FR) = "Console de planification", _
        "French console title is not localized."
    TextCatalogHarness_Assert _
        TextCatalog_Get("CONSOLE.SEVERITY.WARNING", TEXT_LANGUAGE_FR) = "AVERTISSEMENT", _
        "French console severity is not localized."
    TextCatalogHarness_Assert _
        TextCatalog_Get("WBS.FULL_RESET.TITLE", TEXT_LANGUAGE_FR) = "Réinitialisation complète", _
        "French Full Reset title is not localized."
    TextCatalogHarness_Assert _
        TextCatalog_Get("GANTT.CONTROL.MULTI_CRITICAL_PATH", TEXT_LANGUAGE_FR) = "Chemins critiques multiples", _
        "French critical-path label is not localized."

    frenchError = Empty
    TextCatalogHarness_Assert Not ParsePredecessorToken( _
        "1. 2", predecessorWbs, linkType, lagValue, rawToken, englishError, frenchError), _
        "Invalid predecessor token was accepted."
    TextCatalogHarness_Assert englishError = _
        TextCatalog_Get("DIAG.PARSER.PREDECESSOR.SPACES", TEXT_LANGUAGE_EN), _
        "English predecessor error is not catalog-backed."
    TextCatalogHarness_Assert CStr(frenchError) = _
        TextCatalog_Get("DIAG.PARSER.PREDECESSOR.SPACES", TEXT_LANGUAGE_FR), _
        "French predecessor error is not catalog-backed."

    Set wbsToId = CreateObject("Scripting.Dictionary")
    frenchError = Empty
    TextCatalogHarness_Assert Not ParsePredecessorsText( _
        "2", "1.2", "1.1;bad", wbsToId, links, englishError, frenchError), _
        "Invalid predecessor token in a list was accepted."
    TextCatalogHarness_Assert InStr(1, englishError, "Successor WBS 1.2", vbBinaryCompare) = 1, _
        "English predecessor context is incorrect."
    TextCatalogHarness_Assert InStr(1, CStr(frenchError), "WBS successeur 1.2", vbBinaryCompare) = 1, _
        "French predecessor context is incorrect."

    frenchError = Empty
    TextCatalogHarness_Assert Not ParsePredecessorsText( _
        "2", "1.2", "1.1;;1.3", wbsToId, links, englishError, frenchError), _
        "Empty predecessor token in a list was accepted."
    TextCatalogHarness_Assert englishError = TextCatalog_Format( _
        "DIAG.PARSER.PREDECESSOR.EMPTY_IN_TEXT", TEXT_LANGUAGE_EN, _
        TextCatalog_Arguments("Text", "1.1;;1.3")), _
        "English empty-token list error is incorrect."
    TextCatalogHarness_Assert CStr(frenchError) = TextCatalog_Format( _
        "DIAG.PARSER.PREDECESSOR.EMPTY_IN_TEXT", TEXT_LANGUAGE_FR, _
        TextCatalog_Arguments("Text", "1.1;;1.3")), _
        "French empty-token list error is incorrect."

    On Error Resume Next
    Err.Clear
    Set dummyArguments = TextCatalog_Arguments("OnlyName")
    bootstrapDescription = Err.Description
    Err.Clear
    On Error GoTo Failed
    TextCatalogHarness_Assert bootstrapDescription = _
        TextCatalogBootstrap_ErrorWire("TEXTCATALOG.ERROR.ARGUMENT_PAIRS"), _
        "Bootstrap argument error did not use its registered definition."
    TextCatalogHarness_Assert _
        TextCatalogBootstrap_ErrorWire("TEXTCATALOG.ERROR.UNKNOWN") = _
            "[missing text:TEXTCATALOG.ERROR.UNKNOWN]", _
        "Unknown bootstrap key did not use the deterministic marker."

    TextCatalogHarness_Assert Len(TextCatalog_Get("SCURVE.ROW.WARNING.INVALID_WEIGHT", TEXT_LANGUAGE_EN)) > 0, _
        "S-Curve EN warning is missing."
    TextCatalogHarness_Assert Len(TextCatalog_Get("SCURVE.ROW.WARNING.INVALID_WEIGHT", TEXT_LANGUAGE_FR)) > 0, _
        "S-Curve FR warning is missing."
    TextCatalogHarness_Assert Len(TextCatalog_Get("GANTT.TEST.ROW.ACTUAL_WARNING", TEXT_LANGUAGE_EN)) > 0, _
        "Gantt Test EN warning is missing."
    TextCatalogHarness_Assert Len(TextCatalog_Get("GANTT.TEST.ROW.ACTUAL_WARNING", TEXT_LANGUAGE_FR)) > 0, _
        "Gantt Test FR warning is missing."

    TextCatalogHarness_AssertHistoricalDiagnosticValues

    TextCatalogHarness_RunMission54_6DeltaGate = "PASS"
    Exit Function

Failed:
    TextCatalogHarness_RunMission54_6DeltaGate = _
        "FAIL|" & CStr(Err.Number) & "|" & Err.Source & "|" & Err.Description

End Function

Private Sub TextCatalogHarness_AssertHistoricalDiagnosticValues()

    Dim messageKeys As Variant
    Dim expectedValues As Variant
    Dim index As Long

    messageKeys = Array( _
        "DIAG.LABEL.START", "DIAG.LABEL.FINISH", _
        "DIAG.LABEL.ACTUAL_START", "DIAG.LABEL.ACTUAL_FINISH", _
        "DIAG.LABEL.FORECAST_START", "DIAG.LABEL.FORECAST_FINISH", _
        "DIAG.LABEL.CALCULATED_START", "DIAG.LABEL.CALCULATED_FINISH", _
        "DIAG.LABEL.EFFECTIVE_DURATION", _
        "DIAG.CONSTRAINT.TYPE.MUST_START_ON", "DIAG.CONSTRAINT.TYPE.MUST_FINISH_ON", _
        "DIAG.CONSTRAINT.TYPE.UPSTREAM_FINISH", _
        "DIAG.CONSTRAINT.TYPE.START_GENERIC", "DIAG.CONSTRAINT.TYPE.FINISH_GENERIC", _
        "DIAG.CONSTRAINT.TYPE.START_NO_EARLIER", "DIAG.CONSTRAINT.TYPE.START_NO_LATER", _
        "DIAG.CONSTRAINT.TYPE.FINISH_NO_EARLIER", "DIAG.CONSTRAINT.TYPE.FINISH_NO_LATER", _
        "DIAG.CONSTRAINT.TYPE.MUST_START_ON_CANONICAL", _
        "DIAG.CONSTRAINT.TYPE.MUST_FINISH_ON_CANONICAL")
    expectedValues = Array( _
        "START", "FINISH", "Actual Start", "Actual Finish", _
        "Forecast Start", "Forecast Finish", "Calculated Start", "Calculated Finish", _
        "Effective Duration", "Must Start On", "Must Finish On", _
        "Upstream finish constraint", "START constraint", "FINISH constraint", _
        "Start No Earlier Than", "Start No Later Than", _
        "Finish No Earlier Than", "Finish No Later Than", _
        "Must Start On", "Must Finish On")

    For index = LBound(messageKeys) To UBound(messageKeys)
        TextCatalogHarness_Assert _
            TextCatalog_Get(CStr(messageKeys(index)), TEXT_LANGUAGE_EN) = CStr(expectedValues(index)), _
            CStr(messageKeys(index)) & " changed its historical EN wire value."
        TextCatalogHarness_Assert _
            TextCatalog_Get(CStr(messageKeys(index)), TEXT_LANGUAGE_FR) = CStr(expectedValues(index)), _
            CStr(messageKeys(index)) & " changed its historical FR wire value."
    Next index

End Sub

Public Function TextCatalogHarness_RunLocalizedValidationGate() As String

    Dim wbsTable As ListObject
    Dim constraintsTable As ListObject
    Dim originalWbsLanguage As String
    Dim originalConstraintsLanguage As String
    Dim errorNumber As Long
    Dim errorSource As String
    Dim errorDescription As String

    On Error GoTo Failed

    Set wbsTable = ThisWorkbook.Worksheets("WBS").ListObjects("tbl_WBS")
    Set constraintsTable = ThisWorkbook.Worksheets("CONSTRAINTS").ListObjects("tbl_CONSTRAINTS")
    TextCatalogHarness_Assert Not wbsTable.DataBodyRange Is Nothing, "WBS validation fixture has no data rows."
    TextCatalogHarness_Assert Not constraintsTable.DataBodyRange Is Nothing, "Constraints validation fixture has no data rows."

    originalWbsLanguage = WBS_CurrentLanguage()
    originalConstraintsLanguage = Constraints_CurrentLanguage()

    WBS_ApplyLanguage TEXT_LANGUAGE_EN
    TextCatalogHarness_AssertValidation _
        SchemaListColumn(wbsTable, VTS_TABLE_WBS, VTS_COL_TASK_TYPE), _
        "WBS.VALIDATION.TASK_TYPE", TEXT_LANGUAGE_EN
    TextCatalogHarness_AssertValidation _
        SchemaListColumn(wbsTable, VTS_TABLE_WBS, VTS_COL_S), _
        "WBS.VALIDATION.SUMMARY", TEXT_LANGUAGE_EN
    TextCatalogHarness_AssertValidation _
        SchemaListColumn(wbsTable, VTS_TABLE_WBS, VTS_COL_CAL), _
        "WBS.VALIDATION.CALENDAR", TEXT_LANGUAGE_EN

    WBS_ApplyLanguage TEXT_LANGUAGE_FR
    TextCatalogHarness_AssertValidation _
        SchemaListColumn(wbsTable, VTS_TABLE_WBS, VTS_COL_TASK_TYPE), _
        "WBS.VALIDATION.TASK_TYPE", TEXT_LANGUAGE_FR
    TextCatalogHarness_AssertValidation _
        SchemaListColumn(wbsTable, VTS_TABLE_WBS, VTS_COL_S), _
        "WBS.VALIDATION.SUMMARY", TEXT_LANGUAGE_FR
    TextCatalogHarness_AssertValidation _
        SchemaListColumn(wbsTable, VTS_TABLE_WBS, VTS_COL_CAL), _
        "WBS.VALIDATION.CALENDAR", TEXT_LANGUAGE_FR

    Constraints_ApplyLanguage TEXT_LANGUAGE_EN
    TextCatalogHarness_AssertValidation _
        SchemaListColumn(constraintsTable, VTS_TABLE_CONSTRAINTS, VTS_COL_START_CONSTRAINT_TYPE), _
        "CONSTRAINTS.VALIDATION.START_TYPE", TEXT_LANGUAGE_EN
    TextCatalogHarness_AssertValidation _
        SchemaListColumn(constraintsTable, VTS_TABLE_CONSTRAINTS, VTS_COL_FINISH_CONSTRAINT_TYPE), _
        "CONSTRAINTS.VALIDATION.FINISH_TYPE", TEXT_LANGUAGE_EN
    TextCatalogHarness_AssertValidation _
        SchemaListColumn(constraintsTable, VTS_TABLE_CONSTRAINTS, VTS_COL_ACTIVE), _
        "CONSTRAINTS.VALIDATION.ACTIVE", TEXT_LANGUAGE_EN

    Constraints_ApplyLanguage TEXT_LANGUAGE_FR
    TextCatalogHarness_AssertValidation _
        SchemaListColumn(constraintsTable, VTS_TABLE_CONSTRAINTS, VTS_COL_START_CONSTRAINT_TYPE), _
        "CONSTRAINTS.VALIDATION.START_TYPE", TEXT_LANGUAGE_FR
    TextCatalogHarness_AssertValidation _
        SchemaListColumn(constraintsTable, VTS_TABLE_CONSTRAINTS, VTS_COL_FINISH_CONSTRAINT_TYPE), _
        "CONSTRAINTS.VALIDATION.FINISH_TYPE", TEXT_LANGUAGE_FR
    TextCatalogHarness_AssertValidation _
        SchemaListColumn(constraintsTable, VTS_TABLE_CONSTRAINTS, VTS_COL_ACTIVE), _
        "CONSTRAINTS.VALIDATION.ACTIVE", TEXT_LANGUAGE_FR

    WBS_ApplyLanguage originalWbsLanguage
    Constraints_ApplyLanguage originalConstraintsLanguage
    TextCatalogHarness_RunLocalizedValidationGate = "PASS"
    Exit Function

Failed:
    errorNumber = Err.Number
    errorSource = Err.Source
    errorDescription = Err.Description
    On Error Resume Next
    WBS_ApplyLanguage originalWbsLanguage
    Constraints_ApplyLanguage originalConstraintsLanguage
    On Error GoTo 0
    TextCatalogHarness_RunLocalizedValidationGate = _
        "FAIL|" & CStr(errorNumber) & "|" & errorSource & "|" & errorDescription

End Function

Private Sub TextCatalogHarness_AssertValidation( _
    ByVal listColumn As ListColumn, _
    ByVal keyPrefix As String, _
    ByVal languageKey As String)

    Dim validation As Validation

    TextCatalogHarness_Assert Not listColumn Is Nothing, keyPrefix & " column is missing."
    TextCatalogHarness_Assert Not listColumn.DataBodyRange Is Nothing, keyPrefix & " range is empty."
    Set validation = listColumn.DataBodyRange.Validation

    TextCatalogHarness_Assert _
        Replace$(validation.Formula1, ";", ",") = TextCatalog_Get(keyPrefix & ".LIST", languageKey), _
        keyPrefix & " list formula is not localized."
    TextCatalogHarness_Assert validation.InputTitle = TextCatalog_Get(keyPrefix & ".INPUT_TITLE", languageKey), _
        keyPrefix & " input title is not localized."
    TextCatalogHarness_Assert validation.InputMessage = TextCatalog_Get( _
        keyPrefix & IIf(keyPrefix = "CONSTRAINTS.VALIDATION.START_TYPE" Or _
                         keyPrefix = "CONSTRAINTS.VALIDATION.FINISH_TYPE" Or _
                         keyPrefix = "CONSTRAINTS.VALIDATION.ACTIVE", ".INPUT", ".INPUT_MESSAGE"), _
        languageKey), _
        keyPrefix & " input message is not localized."
    TextCatalogHarness_Assert validation.ErrorTitle = TextCatalog_Get(keyPrefix & ".ERROR_TITLE", languageKey), _
        keyPrefix & " error title is not localized."
    TextCatalogHarness_Assert validation.ErrorMessage = TextCatalog_Get( _
        keyPrefix & IIf(keyPrefix = "CONSTRAINTS.VALIDATION.START_TYPE" Or _
                         keyPrefix = "CONSTRAINTS.VALIDATION.FINISH_TYPE" Or _
                         keyPrefix = "CONSTRAINTS.VALIDATION.ACTIVE", ".ERROR", ".ERROR_MESSAGE"), _
        languageKey), _
        keyPrefix & " error message is not localized."

End Sub

Private Sub TextCatalogHarness_Assert(ByVal condition As Boolean, ByVal message As String)
    If Not condition Then Err.Raise vbObjectError + 5490, "mod_TextCatalogHarness", message
End Sub
