Attribute VB_Name = "mod_CoreBridgeDiagnostics"
Option Explicit

'===============================================================================
' MODULE : mod_CoreBridgeDiagnostics
' DOMAINE / DOMAIN : Core Bridge
'
' FR
' Projette les erreurs Core et analytics en messages structures, puis les route vers console et historique.
' Ne doit pas contourner les contrats publics des autres domaines.
'
' EN
' Projects Core and analytics errors into structured messages and routes them to console and history.
' Must not bypass public contracts owned by other domains.
'
' CONTRATS / CONTRACTS : CalcBridge_AppendCoreErrorMessages, CalcBridge_AppendCoreErrorMessagesFromData, CalcBridge_ShowGroupedErrorMessage, CalcBridge_AddInfoMessage, CalcBridge_AddGroupedWarningToCollection, CalcBridge_AddConsoleMessage, CalcBridge_ShowPlanningConsole, CalcBridge_RecordPlanningMessages
' CALLBACKS EXTERNES / EXTERNAL CALLBACKS : Aucun / None
'===============================================================================



'------------------------------------------------------------------------------
' FR: Ajoute la collection Constraint Root Messages a la structure cible fournie par l'appelant.
' EN: Adds the Constraint Root Messages collection to the target structure supplied by the caller.
'------------------------------------------------------------------------------

Private Sub CalcBridge_AddConstraintRootMessages( _
    ByVal consoleMessages As Collection, _
    ByVal constraintMessagesById As Object)

    Dim idVal As Variant

    If consoleMessages Is Nothing Then Exit Sub
    If constraintMessagesById Is Nothing Then Exit Sub

    For Each idVal In constraintMessagesById.Keys
        CalcBridge_AddConsoleMessage consoleMessages, "STOP", _
            CalcBridge_ToPMConstraintMessage(CStr(constraintMessagesById(CStr(idVal))))
    Next idVal

End Sub


'------------------------------------------------------------------------------
' FR: Lit les erreurs bloquantes de CALC et ajoute leurs diagnostics structures a la collection console. En cas de source illisible, ajoute un STOP fallback fail-closed.
' EN: Reads blocking CALC errors and appends their structured diagnostics to the console collection. If the source is unreadable, appends a fail-closed fallback STOP.
'------------------------------------------------------------------------------

Public Sub CalcBridge_AppendCoreErrorMessages( _
    ByVal consoleMessages As Collection, _
    ByVal tblCalc As ListObject)

    Dim mapCalc As Object
    Dim arr As Variant

    On Error GoTo FailSafe

    If consoleMessages Is Nothing Then Exit Sub
    If tblCalc Is Nothing Then GoTo FailSafe
    If tblCalc.DataBodyRange Is Nothing Then GoTo FailSafe

    Set mapCalc = CanonicalIdentity_BuildColumnMap(tblCalc)
    arr = tblCalc.DataBodyRange.value

    CalcBridge_AppendCoreErrorMessagesFromData consoleMessages, arr, mapCalc, Nothing, "PROD"
    Exit Sub

FailSafe:
    CalcBridge_AddConsoleMessage consoleMessages, "STOP", _
        PlanningMessageText_Format("DIAG.PLANNING.SOURCE_DIAGNOSTIC_FAILED", Nothing, Nothing)

End Sub


'------------------------------------------------------------------------------
' FR: Projette les erreurs d'un dataset Core en messages racine, dependance, contrainte et cascade sans recalculer le planning. Les donnees invalides produisent un STOP fallback.
' EN: Projects Core dataset errors into root, dependency, constraint and cascade messages without recalculating planning. Invalid input produces a fallback STOP.
'------------------------------------------------------------------------------

Public Sub CalcBridge_AppendCoreErrorMessagesFromData( _
    ByVal consoleMessages As Collection, _
    ByRef dataArr As Variant, _
    ByVal mapCore As Object, _
    Optional ByVal rootErrorIds As Object = Nothing, _
    Optional ByVal contextMode As String = "PROD", _
    Optional ByVal dependencyDiagnostics As Object = Nothing, _
    Optional ByVal constraintDiagnostics As Object = Nothing, _
    Optional ByVal cascadeDiagnostics As Object = Nothing, _
    Optional ByVal coreDiagnostics As Object = Nothing)

    Dim mapCalc As Object
    Dim arr As Variant
    Dim r As Long

    Dim idToWbs As Object
    Dim idToTaskName As Object

    Dim errMissingPred As Object
    Dim errCycle As Object
    Dim errUnsupportedLinkType As Object
    Dim errActualStartConflict As Object
    Dim errActualFinishConflict As Object
    Dim errForecastConflict As Object
    Dim errForecastFinishConflict As Object
    Dim errMissingDuration As Object
    Dim errStartNotComputable As Object
    Dim errFinishBeforeStart As Object
    Dim errLOEAsPredecessor As Object
    Dim errLOEMissingSS As Object
    Dim errLOEMissingFF As Object
    Dim errLOEInvalidLink As Object
    Dim errOtherRoot As Object
    Dim errConstraintRootMessages As Object
    Dim cycleDetailMessage As String

    Dim idVal As String
    Dim wbsVal As String
    Dim taskNameVal As String
    Dim errMsg As String
    Dim hasSpecificRootError As Boolean
    Dim contextKey As String

    On Error GoTo FailSafe

    If consoleMessages Is Nothing Then Exit Sub
    If Not IsArray(dataArr) Then GoTo FailSafe
    If mapCore Is Nothing Then GoTo FailSafe

    Set mapCalc = mapCore
    arr = dataArr
    contextKey = UCase$(Trim$(contextMode))
    Set errMissingPred = CreateObject("Scripting.Dictionary")
    Set errCycle = CreateObject("Scripting.Dictionary")
    Set errUnsupportedLinkType = CreateObject("Scripting.Dictionary")
    Set errActualStartConflict = CreateObject("Scripting.Dictionary")
    Set errActualFinishConflict = CreateObject("Scripting.Dictionary")
    Set errForecastConflict = CreateObject("Scripting.Dictionary")
    Set errForecastFinishConflict = CreateObject("Scripting.Dictionary")
    Set errMissingDuration = CreateObject("Scripting.Dictionary")
    Set errStartNotComputable = CreateObject("Scripting.Dictionary")
    Set errFinishBeforeStart = CreateObject("Scripting.Dictionary")
    Set errLOEAsPredecessor = CreateObject("Scripting.Dictionary")
    Set errLOEMissingSS = CreateObject("Scripting.Dictionary")
    Set errLOEMissingFF = CreateObject("Scripting.Dictionary")
    Set errLOEInvalidLink = CreateObject("Scripting.Dictionary")
    Set errOtherRoot = CreateObject("Scripting.Dictionary")
    Set errConstraintRootMessages = CreateObject("Scripting.Dictionary")
    Set idToWbs = CreateObject("Scripting.Dictionary")
    Set idToTaskName = CreateObject("Scripting.Dictionary")

    If Not mapCalc.Exists("ID") Then GoTo FailSafe
    If Not mapCalc.Exists("Error flag") Then GoTo FailSafe
    If Not mapCalc.Exists("ErrorMsg") Then GoTo FailSafe

    For r = 1 To UBound(arr, 1)

        idVal = Trim$(CStr(arr(r, mapCalc("ID"))))

        If idVal <> "" Then

            If mapCalc.Exists("WBS") Then
                wbsVal = Trim$(CStr(arr(r, mapCalc("WBS"))))
            Else
                wbsVal = ""
            End If

            idToWbs(idVal) = wbsVal
            If mapCalc.Exists("Task Name") Then
                taskNameVal = Trim$(CStr(arr(r, mapCalc("Task Name"))))
            Else
                taskNameVal = ""
            End If
            idToTaskName(idVal) = taskNameVal

            If UCase$(Trim$(CStr(arr(r, mapCalc("Error flag"))))) = "ERROR" Then

                errMsg = Trim$(CStr(arr(r, mapCalc("ErrorMsg"))))

                If coreDiagnostics Is Nothing Then GoTo FailSafe

                'Inherited errors remain visible in the source table,
                'but are not highlighted in the main popup.
                If Not CoreDiagnostics_TaskHasClassification(coreDiagnostics, idVal, "ROOT") Then GoTo NextRow

                If Not rootErrorIds Is Nothing Then
                    If Not rootErrorIds.Exists(idVal) Then GoTo NextRow
                End If

                CalcBridge_ClassifyStructuredCoreDiagnostics coreDiagnostics, idVal, errMsg, _
                    errMissingPred, errCycle, errUnsupportedLinkType, _
                    errActualStartConflict, errActualFinishConflict, errForecastConflict, _
                    errForecastFinishConflict, errMissingDuration, errStartNotComputable, _
                    errFinishBeforeStart, errLOEAsPredecessor, errLOEMissingSS, _
                    errLOEMissingFF, errLOEInvalidLink, errOtherRoot, _
                    errConstraintRootMessages, cycleDetailMessage

            End If
        End If

NextRow:
    Next r

    hasSpecificRootError = _
        (errMissingPred.Count > 0) Or _
        (errUnsupportedLinkType.Count > 0) Or _
        (errCycle.Count > 0) Or _
        (errActualStartConflict.Count > 0) Or _
        (errActualFinishConflict.Count > 0) Or _
        (errForecastConflict.Count > 0) Or _
        (errForecastFinishConflict.Count > 0) Or _
        (errMissingDuration.Count > 0) Or _
        (errStartNotComputable.Count > 0) Or _
        (errFinishBeforeStart.Count > 0) Or _
        (errConstraintRootMessages.Count > 0) Or _
        (errLOEAsPredecessor.Count > 0) Or _
        (errLOEMissingSS.Count > 0) Or _
        (errLOEMissingFF.Count > 0) Or _
        (errLOEInvalidLink.Count > 0)

    If errLOEAsPredecessor.Count > 0 Then
        CalcBridge_AddGroupedStopToCollection consoleMessages, errLOEAsPredecessor, idToWbs, _
            "DIAG.GROUP.LOE.PREDECESSOR"
    End If

    If errLOEMissingSS.Count > 0 Then
        CalcBridge_AddGroupedStopToCollection consoleMessages, errLOEMissingSS, idToWbs, _
            "DIAG.GROUP.LOE.MISSING_SS"
    End If

    If errLOEMissingFF.Count > 0 Then
        CalcBridge_AddGroupedStopToCollection consoleMessages, errLOEMissingFF, idToWbs, _
            "DIAG.GROUP.LOE.MISSING_FF"
    End If

    If errLOEInvalidLink.Count > 0 Then
        CalcBridge_AddGroupedStopToCollection consoleMessages, errLOEInvalidLink, idToWbs, _
            "DIAG.GROUP.LOE.INVALID_LINK"
    End If

    If errMissingPred.Count > 0 Then
        CalcBridge_AddGroupedStopToCollection consoleMessages, errMissingPred, idToWbs, _
            "DIAG.GROUP.DEPENDENCY.MISSING_PREDECESSOR"
    End If

    If errUnsupportedLinkType.Count > 0 Then
        CalcBridge_AddGroupedStopToCollection consoleMessages, errUnsupportedLinkType, idToWbs, _
            "DIAG.GROUP.DEPENDENCY.UNSUPPORTED_LINK"
    End If

    If errCycle.Count > 0 Then
        If CalcBridge_CycleDetailMessageFromCoreError(cycleDetailMessage) <> "" Then
            CalcBridge_AddConsoleMessage consoleMessages, "STOP", _
                CalcBridge_CycleDetailMessageFromCoreError(cycleDetailMessage)
        Else
            CalcBridge_AddGroupedStopToCollection consoleMessages, errCycle, idToWbs, _
                "DIAG.GROUP.DEPENDENCY.CYCLE"
        End If
    End If

    If errActualStartConflict.Count > 0 Then
        CalcBridge_AddUpstreamStopToCollection consoleMessages, errActualStartConflict, idToWbs, _
            "DIAG.UPSTREAM.ACTUAL_START_CONFLICT"
    End If

    If errActualFinishConflict.Count > 0 Then
        If Not CalcBridge_TryAddConstraintDiagnosticStops(consoleMessages, errActualFinishConflict, idToWbs, idToTaskName, constraintDiagnostics, cascadeDiagnostics, contextKey) Then
            CalcBridge_AddUpstreamStopToCollection consoleMessages, errActualFinishConflict, idToWbs, _
                "DIAG.UPSTREAM.ACTUAL_FINISH_CONFLICT"
        End If
    End If

    If errForecastConflict.Count > 0 Then
        If contextKey = "TEST" Or contextKey = "SCENARIO" Then
            If Not CalcBridge_TryAddForecastStartDependencyDiagnosticStops( _
                consoleMessages, errForecastConflict, idToWbs, idToTaskName, dependencyDiagnostics, contextKey) Then

                CalcBridge_AddGroupedStopToCollection consoleMessages, errForecastConflict, idToWbs, _
                    "DIAG.GROUP.TEST.START_CONFLICT"
            End If
        Else
            CalcBridge_AddGroupedStopToCollection consoleMessages, errForecastConflict, idToWbs, _
                "DIAG.GROUP.FORECAST.START_CONFLICT"
        End If
    End If
    If errForecastFinishConflict.Count > 0 Then
        If CalcBridge_TryAddConstraintDiagnosticStops(consoleMessages, errForecastFinishConflict, idToWbs, idToTaskName, constraintDiagnostics, cascadeDiagnostics, contextKey) Then
            'Structured constraint diagnostic already rendered.
        ElseIf contextKey = "TEST" Or contextKey = "SCENARIO" Then
            CalcBridge_AddGroupedStopToCollection consoleMessages, errForecastFinishConflict, idToWbs, _
                "DIAG.GROUP.TEST.FINISH_CONFLICT"
        Else
            CalcBridge_AddGroupedStopToCollection consoleMessages, errForecastFinishConflict, idToWbs, _
                "DIAG.GROUP.FORECAST.FINISH_CONFLICT"
        End If
    End If
    If errConstraintRootMessages.Count > 0 Then
        If Not CalcBridge_TryAddConstraintDiagnosticStops(consoleMessages, errConstraintRootMessages, idToWbs, idToTaskName, constraintDiagnostics, cascadeDiagnostics, contextKey) Then
            CalcBridge_AddConstraintRootMessages consoleMessages, errConstraintRootMessages
        End If
    End If

    If errMissingDuration.Count > 0 Then
        CalcBridge_AddGroupedStopToCollection consoleMessages, errMissingDuration, idToWbs, _
            "DIAG.GROUP.BASELINE.MISSING_DURATION"
    End If

    If errStartNotComputable.Count > 0 Then
        CalcBridge_AddGroupedStopToCollection consoleMessages, errStartNotComputable, idToWbs, _
            "DIAG.GROUP.DATES.START_UNRESOLVED"
    End If

    If errFinishBeforeStart.Count > 0 Then
        CalcBridge_AddGroupedStopToCollection consoleMessages, errFinishBeforeStart, idToWbs, _
            "DIAG.GROUP.DATES.FINISH_BEFORE_START"
    End If

    If errOtherRoot.Count > 0 And Not hasSpecificRootError Then
        If contextKey = "TEST" Then
            CalcBridge_AddGroupedStopToCollection consoleMessages, errOtherRoot, idToWbs, _
                "DIAG.GROUP.ENGINE.TEST_ERROR"
        ElseIf contextKey = "SCENARIO" Then
            CalcBridge_AddGroupedStopToCollection consoleMessages, errOtherRoot, idToWbs, _
                "DIAG.GROUP.ENGINE.SCENARIO_ERROR"
        Else
            CalcBridge_AddGroupedStopToCollection consoleMessages, errOtherRoot, idToWbs, _
                "DIAG.GROUP.ENGINE.ERROR"
        End If
    End If
    Exit Sub

FailSafe:
    CalcBridge_AddConsoleMessage consoleMessages, "STOP", _
        PlanningMessageText_Format("DIAG.PLANNING.DATA_DIAGNOSTIC_FAILED", Nothing, Nothing)

End Sub


'------------------------------------------------------------------------------
' FR: Projette la collection Grouped Error Message vers l'interface autorisee par la politique runtime.
' EN: Projects the Grouped Error Message collection to the UI allowed by runtime policy.
'------------------------------------------------------------------------------

Public Sub CalcBridge_ShowGroupedErrorMessage( _
    ByVal idsDict As Object, _
    ByVal idToWbs As Object, _
    ByVal messageKey As String)

    Dim consoleMessages As Collection

    If idsDict Is Nothing Then Exit Sub
    If idsDict.Count = 0 Then Exit Sub

    Set consoleMessages = New Collection

    CalcBridge_AddGroupedStopToCollection consoleMessages, idsDict, idToWbs, _
        messageKey

    CalcBridge_ShowPlanningConsole consoleMessages

End Sub


'------------------------------------------------------------------------------
' FR: Ajoute la collection Info Message a la structure cible fournie par l'appelant.
' EN: Adds the Info Message collection to the target structure supplied by the caller.
'------------------------------------------------------------------------------

Public Sub CalcBridge_AddInfoMessage( _
    ByVal messages As Collection, _
    ByVal messageKey As String)

    Dim msg As String

    If messages Is Nothing Then Exit Sub

    msg = PlanningMessageText_Format(messageKey, Nothing, Nothing)

    CalcBridge_AddConsoleMessage messages, "INFO", msg

End Sub


'------------------------------------------------------------------------------
' FR: Ajoute la collection Grouped Warning To Collection a la structure cible fournie par l'appelant.
' EN: Adds the Grouped Warning To Collection collection to the target structure supplied by the caller.
'------------------------------------------------------------------------------

Public Sub CalcBridge_AddGroupedWarningToCollection( _
    ByVal messages As Collection, _
    ByVal idsDict As Object, _
    ByVal idToWbs As Object, _
    ByVal messageKey As String, _
    Optional ByVal historyHandled As Boolean = False, _
    Optional ByVal ackTokens As String = "", _
    Optional ByVal memberReceipts As Object = Nothing)

    CalcBridge_AddGroupedConsoleMessage messages, "WARNING", idsDict, idToWbs, _
        messageKey, ackTokens, memberReceipts

End Sub

Public Sub CalcBridge_AddGroupedConsoleMessage( _
    ByVal messages As Collection, _
    ByVal severity As String, _
    ByVal idsDict As Object, _
    ByVal idToWbs As Object, _
    ByVal messageKey As String, _
    Optional ByVal ackTokens As String = "", _
    Optional ByVal memberReceipts As Object = Nothing)

    Dim msg As String
    Dim item As Object
    Dim members As Collection
    Dim id As Variant
    Dim receipt As Object

    If messages Is Nothing Then Exit Sub
    If idsDict Is Nothing Then Exit Sub
    If idsDict.Count = 0 Then Exit Sub

    msg = CalcBridge_BuildGroupedMessage(idsDict, idToWbs, messageKey)
    Set item = CalcBridge_CreateConsoleItem(severity, msg, False, messageKey, , ackTokens)
    Set members = New Collection
    For Each id In idsDict.Keys
        If Not memberReceipts Is Nothing Then
            If Not memberReceipts.Exists(CStr(id)) Then _
                Err.Raise 5, "CalcBridge_AddGroupedConsoleMessage", "Missing member receipt."
            Set receipt = memberReceipts(CStr(id))
            members.Add CStr(receipt("Hash"))
        Else
            members.Add "TASK:" & CStr(id)
        End If
    Next id
    Set item("GroupMembers") = members
    messages.Add item

End Sub


'------------------------------------------------------------------------------
' FR: Construit la map Console Item a partir des donnees fournies par l'appelant.
' EN: Builds the Console Item map from data supplied by the caller.
'------------------------------------------------------------------------------

Private Function CalcBridge_CreateConsoleItem( _
    ByVal msgType As String, _
    ByVal msgText As String, _
    Optional ByVal historyHandled As Boolean = False, _
    Optional ByVal eventType As String = "", _
    Optional ByVal eventHash As String = "", _
    Optional ByVal ackTokens As String = "", _
    Optional ByVal historyReceipt As Object = Nothing, _
    Optional ByVal subjectTaskId As String = "", _
    Optional ByVal sourceSheet As String = "") As Object

    Dim item As Object

    Set item = CreateObject("Scripting.Dictionary")
    item("Type") = UCase$(Trim$(msgType))
    item("Message") = CStr(msgText)
    item("HistoryState") = "NOT_LOGGED"
    If Not historyReceipt Is Nothing Then
        Set item("HistoryReceipt") = historyReceipt
        item("HistoryState") = "PERSISTED"
    End If
    If Trim$(eventType) <> "" Then item("EventType") = Trim$(eventType)
    If Trim$(eventHash) <> "" Then item("Hash") = Trim$(eventHash)
    If Trim$(ackTokens) <> "" Then item("AckTokens") = Trim$(ackTokens)
    If Trim$(subjectTaskId) <> "" Then item("TaskId") = Trim$(subjectTaskId)
    If Trim$(sourceSheet) <> "" Then item("SourceSheet") = Trim$(sourceSheet)

    Set CalcBridge_CreateConsoleItem = item

End Function


'------------------------------------------------------------------------------
' FR: Ajoute la collection Console Message a la structure cible fournie par l'appelant.
' EN: Adds the Console Message collection to the target structure supplied by the caller.
'------------------------------------------------------------------------------

Public Sub CalcBridge_AddConsoleMessage( _
    ByVal targetMessages As Collection, _
    ByVal msgType As String, _
    ByVal msgText As String, _
    Optional ByVal historyHandled As Boolean = False, _
    Optional ByVal eventType As String = "", _
    Optional ByVal eventHash As String = "", _
    Optional ByVal ackTokens As String = "", _
    Optional ByVal historyReceipt As Object = Nothing, _
    Optional ByVal subjectTaskId As String = "", _
    Optional ByVal sourceSheet As String = "")

    If targetMessages Is Nothing Then Exit Sub
    If Trim$(msgText) = "" Then Exit Sub

    targetMessages.Add CalcBridge_CreateConsoleItem(msgType, msgText, historyHandled, eventType, eventHash, ackTokens, historyReceipt, subjectTaskId, sourceSheet)

End Sub


'------------------------------------------------------------------------------
' FR: Projette la collection Planning Console vers l'interface autorisee par la politique runtime.
' EN: Projects the Planning Console collection to the UI allowed by runtime policy.
'------------------------------------------------------------------------------

Public Sub CalcBridge_ShowPlanningConsole(ByVal messages As Collection)

    Dim historyMessages As Collection
    Dim displayMessages As Collection
    Dim historyErrorMessage As String
    Dim prepScope As clsPerfScope
    Dim showScope As clsPerfScope
    Dim historyOk As Boolean
    Dim failureItem As Object

    If messages Is Nothing Then Exit Sub
    If messages.Count = 0 Then Exit Sub

    Profiler_RecordOperation "PlanningConsole_ENGINE_READY", 1, 0#
    Profiler_RecordOperation "PlanningConsole_PREP_START", 1, 0#
    Set prepScope = Profiler_BeginScope("PlanningConsole_PrepareBeforeShow", "Planning Console")

    Set historyMessages = MessageEngine_PrepareConsoleMessages(messages)

    historyOk = PlanningEvents_LogConsoleMessagesSafe( _
        historyMessages, _
        "CalcBridge_ShowPlanningConsole", _
        historyErrorMessage)
    If Not historyOk Then
        Set failureItem = CalcBridge_CreateConsoleItem("STOP", historyErrorMessage)
        failureItem("HistoryState") = "STORE_UNAVAILABLE"
        historyMessages.Add failureItem
    End If

    If Profiler_ShouldSuppressUserInterface() Then Exit Sub

    If ShouldDeferCurrentWorkflowDisplayToRoot("CalcBridge_ShowPlanningConsole") Then
        DeferPlanningWorkflowDisplayMessages historyMessages
        Exit Sub
    End If

    If IsPlanningWorkflowStopOnlyDisplay() Then
        Set displayMessages = CalcBridge_FilterConsoleMessagesBySeverity(historyMessages, "STOP")
    Else
        Set displayMessages = MessageEngine_PrepareDisplayMessages(historyMessages)
    End If
    If Not MessageEngine_ShouldShowConsole(displayMessages) Then Exit Sub
    If Not CanCurrentWorkflowDisplay("CalcBridge_ShowPlanningConsole") Then Exit Sub
    If historyOk Then
        If Not PlanningEvents_LogConsoleMessagesSafe( _
            displayMessages, "CalcBridge_ShowPlanningConsole", historyErrorMessage) Then
            Set failureItem = CalcBridge_CreateConsoleItem("STOP", historyErrorMessage)
            failureItem("HistoryState") = "STORE_UNAVAILABLE"
            displayMessages.Add failureItem
        End If
    End If

    If PlanningConsolePolicy_IsNonInteractive() Then
        Set prepScope = Nothing
        RunButtonsTrace_Checkpoint "ConsolePolicy", "Planning console noninteractive capture start"
        PlanningConsolePolicy_CaptureDisplayMessages displayMessages, "Planning console"
        RunButtonsTrace_Checkpoint "ConsolePolicy", "Planning console noninteractive capture returned"
        Exit Sub
    End If

    Load frmPlanningMessages
    RunButtonsTrace_Checkpoint "Console", "Planning console modal load start"
    frmPlanningMessages.LoadMessages displayMessages, "Planning console"
    frmPlanningMessages.PrepareForImmediateShow
    RunButtonsTrace_Checkpoint "Console", "Planning console ready before show"
    Profiler_RecordOperation "PlanningConsole_PREP_END", 1, 0#
    Set prepScope = Nothing

    Profiler_RecordOperation "PlanningConsole_SHOW_START", 1, 0#
    RunButtonsTrace_Checkpoint "Console", "Planning console modal show start"
    Set showScope = Profiler_BeginScope("PlanningConsole_ShowModalLifetime", "Planning Console")
    frmPlanningMessages.Show vbModal
    Set showScope = Nothing
    RunButtonsTrace_Checkpoint "Console", "Planning console modal show returned"
    Unload frmPlanningMessages
    Profiler_RecordOperation "PlanningConsole_SHOW_RETURN", 1, 0#

End Sub


'------------------------------------------------------------------------------
' FR: Retourne la collection Filter Console Messages By Severity sans modifier les donnees d'entree.
' EN: Returns the Filter Console Messages By Severity collection without mutating input data.
'------------------------------------------------------------------------------

Private Function CalcBridge_FilterConsoleMessagesBySeverity( _
    ByVal messages As Collection, _
    ByVal severity As String) As Collection

    Dim result As Collection
    Dim item As Variant

    Set result = New Collection

    If messages Is Nothing Then
        Set CalcBridge_FilterConsoleMessagesBySeverity = result
        Exit Function
    End If

    For Each item In messages
        If UCase$(Trim$(CStr(item("Type")))) = UCase$(Trim$(severity)) Then
            result.Add item
        End If
    Next item

    Set CalcBridge_FilterConsoleMessagesBySeverity = result

End Function


'------------------------------------------------------------------------------
' FR: Ecrit ou synchronise Record Planning Messages dans le stockage possede par le domaine.
' EN: Writes or synchronizes Record Planning Messages in the store owned by the domain.
'------------------------------------------------------------------------------

Public Function CalcBridge_RecordPlanningMessages( _
    ByVal messages As Collection, _
    Optional ByVal sourceProcedure As String = "CalcBridge_RecordPlanningMessages") As Boolean

    Dim historyMessages As Collection
    Dim historyErrorMessage As String

    If messages Is Nothing Then
        CalcBridge_RecordPlanningMessages = True
        Exit Function
    End If

    If messages.Count = 0 Then
        CalcBridge_RecordPlanningMessages = True
        Exit Function
    End If

    Set historyMessages = MessageEngine_PrepareConsoleMessages(messages)
    If historyMessages.Count = 0 Then
        CalcBridge_RecordPlanningMessages = True
        Exit Function
    End If

    CalcBridge_RecordPlanningMessages = PlanningEvents_LogConsoleMessagesSafe( _
        historyMessages, _
        sourceProcedure, _
        historyErrorMessage)

    If Not CalcBridge_RecordPlanningMessages Then
        Debug.Print "Runtime history logging failed in " & sourceProcedure & ": " & historyErrorMessage
    End If

End Function


'------------------------------------------------------------------------------
' FR: Retourne la collection Try Add Constraint Diagnostic Stops sans modifier les donnees d'entree.
' EN: Returns the Try Add Constraint Diagnostic Stops collection without mutating input data.
'------------------------------------------------------------------------------

Private Function CalcBridge_TryAddConstraintDiagnosticStops( _
    ByVal messages As Collection, _
    ByVal constraintMessagesById As Object, _
    ByVal idToWbs As Object, _
    ByVal idToTaskName As Object, _
    ByVal constraintDiagnostics As Object, _
    ByVal cascadeDiagnostics As Object, _
    ByVal contextKey As String) As Boolean

    Dim key As Variant
    Dim diag As Object
    Dim detailMsg As String
    Dim addedAny As Boolean

    If messages Is Nothing Then Exit Function
    If constraintMessagesById Is Nothing Then Exit Function
    If constraintMessagesById.Count = 0 Then Exit Function

    For Each key In constraintMessagesById.Keys
        detailMsg = ""

        If Not constraintDiagnostics Is Nothing Then
            If constraintDiagnostics.Exists(CStr(key)) Then
                If IsObject(constraintDiagnostics(CStr(key))) Then
                    Set diag = constraintDiagnostics(CStr(key))
                    detailMsg = CalcBridge_BuildConstraintDiagnosticMessage( _
                        diag, idToWbs, idToTaskName, cascadeDiagnostics, contextKey)
                End If
            End If
        End If

        If Trim$(detailMsg) = "" Then
            detailMsg = CalcBridge_ToPMConstraintMessage(CStr(constraintMessagesById(CStr(key))))
        End If

        If Trim$(detailMsg) <> "" Then
            CalcBridge_AddConsoleMessage messages, "STOP", detailMsg
            addedAny = True
        End If
    Next key

    CalcBridge_TryAddConstraintDiagnosticStops = addedAny

End Function


'------------------------------------------------------------------------------
' FR: Retourne la collection Try Add Forecast Start Dependency Diagnostic Stops sans modifier les donnees d'entree.
' EN: Returns the Try Add Forecast Start Dependency Diagnostic Stops collection without mutating input data.
'------------------------------------------------------------------------------

Private Function CalcBridge_TryAddForecastStartDependencyDiagnosticStops( _
    ByVal messages As Collection, _
    ByVal idsDict As Object, _
    ByVal idToWbs As Object, _
    ByVal idToTaskName As Object, _
    ByVal dependencyDiagnostics As Object, _
    ByVal contextKey As String) As Boolean

    Dim key As Variant
    Dim diag As Object
    Dim detailMsg As String
    Dim addedAny As Boolean

    If messages Is Nothing Then Exit Function
    If idsDict Is Nothing Then Exit Function
    If idsDict.Count = 0 Then Exit Function
    If dependencyDiagnostics Is Nothing Then Exit Function

    For Each key In idsDict.Keys
        If dependencyDiagnostics.Exists(CStr(key)) Then
            If IsObject(dependencyDiagnostics(CStr(key))) Then
                Set diag = dependencyDiagnostics(CStr(key))
                detailMsg = CalcBridge_BuildForecastStartDependencyDiagnosticMessage( _
                    diag, idToWbs, idToTaskName, contextKey)

                If Trim$(detailMsg) <> "" Then
                    CalcBridge_AddConsoleMessage messages, "STOP", detailMsg
                    addedAny = True
                End If
            End If
        End If
    Next key

    CalcBridge_TryAddForecastStartDependencyDiagnosticStops = addedAny

End Function

'------------------------------------------------------------------------------
' FR: Ajoute la collection Grouped Stop To Collection a la structure cible fournie par l'appelant.
' EN: Adds the Grouped Stop To Collection collection to the target structure supplied by the caller.
'------------------------------------------------------------------------------

Public Sub CalcBridge_AddGroupedStopToCollection( _
    ByVal messages As Collection, _
    ByVal idsDict As Object, _
    ByVal idToWbs As Object, _
    ByVal messageKey As String)

    If messages Is Nothing Then Exit Sub
    If idsDict Is Nothing Then Exit Sub
    If idsDict.Count = 0 Then Exit Sub

    CalcBridge_AddConsoleMessage messages, "STOP", _
        CalcBridge_BuildGroupedMessage(idsDict, idToWbs, messageKey)

End Sub


'------------------------------------------------------------------------------
' FR: Ajoute la collection Upstream Stop To Collection a la structure cible fournie par l'appelant.
' EN: Adds the Upstream Stop To Collection collection to the target structure supplied by the caller.
'------------------------------------------------------------------------------

Private Sub CalcBridge_AddUpstreamStopToCollection( _
    ByVal messages As Collection, _
    ByVal idsDict As Object, _
    ByVal idToWbs As Object, _
    ByVal messageKey As String)

    Dim itemsLine As String
    Dim msg As String

    If messages Is Nothing Then Exit Sub
    If idsDict Is Nothing Then Exit Sub
    If idsDict.Count = 0 Then Exit Sub

    itemsLine = CalcBridge_BuildUpstreamViolationItems(idsDict, idToWbs, 20)

    msg = PlanningMessageText_Format(messageKey, _
        TextCatalog_Arguments("Items", itemsLine), _
        TextCatalog_Arguments("Items", itemsLine))

    CalcBridge_AddConsoleMessage messages, "STOP", msg

End Sub


'------------------------------------------------------------------------------
' FR: Projette la collection Single Console Message vers l'interface autorisee par la politique runtime.
' EN: Projects the Single Console Message collection to the UI allowed by runtime policy.
'------------------------------------------------------------------------------

Public Sub CalcBridge_ShowSingleConsoleMessage( _
    ByVal msgType As String, _
    ByVal messageKey As String, _
    Optional ByVal namedArguments As Object = Nothing)

    Dim consoleMessages As Collection

    Set consoleMessages = New Collection

    CalcBridge_AddConsoleMessage consoleMessages, msgType, _
        PlanningMessageText_Format(messageKey, namedArguments, namedArguments)

    CalcBridge_ShowPlanningConsole consoleMessages

End Sub


'------------------------------------------------------------------------------
' FR: Ajoute la collection Or Show Raw Console Message a la structure cible fournie par l'appelant.
' EN: Adds the Or Show Raw Console Message collection to the target structure supplied by the caller.
'------------------------------------------------------------------------------

Public Sub CalcBridge_AddOrShowRawConsoleMessage( _
    ByVal consoleMessages As Collection, _
    ByVal msgType As String, _
    ByVal msgText As String)

    Dim localMessages As Collection

    If Trim$(CStr(msgText)) = "" Then Exit Sub

    If consoleMessages Is Nothing Then
        Set localMessages = New Collection
        CalcBridge_AddConsoleMessage localMessages, msgType, msgText
        CalcBridge_ShowPlanningConsole localMessages
    Else
        CalcBridge_AddConsoleMessage consoleMessages, msgType, msgText
    End If

End Sub


'------------------------------------------------------------------------------
' FR: Ajoute la collection Or Show Console Message a la structure cible fournie par l'appelant.
' EN: Adds the Or Show Console Message collection to the target structure supplied by the caller.
'------------------------------------------------------------------------------

Public Sub CalcBridge_AddOrShowConsoleMessage( _
    ByVal consoleMessages As Collection, _
    ByVal msgType As String, _
    ByVal messageKey As String, _
    Optional ByVal namedArguments As Object = Nothing)

    CalcBridge_AddOrShowRawConsoleMessage consoleMessages, msgType, _
        PlanningMessageText_Format(messageKey, namedArguments, namedArguments)

End Sub



'------------------------------------------------------------------------------
' FR: Classe les diagnostics Core par codes stables, jamais par texte rendu.
' EN: Classifies Core diagnostics by stable codes, never rendered text.
'------------------------------------------------------------------------------
Private Sub CalcBridge_ClassifyStructuredCoreDiagnostics( _
    ByVal coreDiagnostics As Object, _
    ByVal taskId As String, _
    ByVal renderedErrorText As String, _
    ByVal errMissingPred As Object, _
    ByVal errCycle As Object, _
    ByVal errUnsupportedLinkType As Object, _
    ByVal errActualStartConflict As Object, _
    ByVal errActualFinishConflict As Object, _
    ByVal errForecastConflict As Object, _
    ByVal errForecastFinishConflict As Object, _
    ByVal errMissingDuration As Object, _
    ByVal errStartNotComputable As Object, _
    ByVal errFinishBeforeStart As Object, _
    ByVal errLOEAsPredecessor As Object, _
    ByVal errLOEMissingSS As Object, _
    ByVal errLOEMissingFF As Object, _
    ByVal errLOEInvalidLink As Object, _
    ByVal errOtherRoot As Object, _
    ByVal errConstraintRootMessages As Object, _
    ByRef cycleDetailMessage As String)

    Dim records As Collection
    Dim record As Variant
    Dim code As String
    Dim foundSpecific As Boolean

    Set records = CoreDiagnostics_ForTask(coreDiagnostics, taskId)
    If records Is Nothing Then
        errOtherRoot(taskId) = True
        Exit Sub
    End If

    For Each record In records
        If IsObject(record) Then
            If StrComp(CStr(record("Classification")), "ROOT", vbTextCompare) = 0 Then
                code = UCase$(Trim$(CStr(record("Code"))))

                Select Case code
                    Case "CORE.ERROR.LOE_AS_PREDECESSOR", "CORE.ERROR.INVALID_LOE_PREDECESSOR_ID"
                        errLOEAsPredecessor(taskId) = True: foundSpecific = True
                    Case "CORE.ERROR.LOE_SS_REQUIRED", "CORE.ERROR.LOE_SS_PREDECESSOR_MISSING", _
                         "CORE.ERROR.LOE_SS_PREDECESSOR_NOT_FOUND", "CORE.ERROR.LOE_SS_START_UNAVAILABLE"
                        errLOEMissingSS(taskId) = True: foundSpecific = True
                    Case "CORE.ERROR.LOE_FF_REQUIRED", "CORE.ERROR.LOE_FF_PREDECESSOR_MISSING", _
                         "CORE.ERROR.LOE_FF_PREDECESSOR_NOT_FOUND", "CORE.ERROR.LOE_FF_FINISH_UNAVAILABLE"
                        errLOEMissingFF(taskId) = True: foundSpecific = True
                    Case "CORE.ERROR.LOE_LINK_TYPE"
                        errLOEInvalidLink(taskId) = True: foundSpecific = True
                    Case "CORE.ERROR.MISSING_PREDECESSOR", "CORE.ERROR.MISSING_PREDECESSOR_ID"
                        errMissingPred(taskId) = True: foundSpecific = True
                    Case "DIAG.CALC_ENGINE.CYCLE_MARKER"
                        errCycle(taskId) = True: foundSpecific = True
                        If cycleDetailMessage = "" Then cycleDetailMessage = renderedErrorText
                    Case "CORE.ERROR.UNSUPPORTED_LINK_TYPE"
                        errUnsupportedLinkType(taskId) = True: foundSpecific = True
                    Case "CORE.ERROR.ACTUAL_START_DEPENDENCIES"
                        errActualStartConflict(taskId) = True: foundSpecific = True
                    Case "CORE.ERROR.ACTUAL_FINISH_CONSTRAINTS"
                        errActualFinishConflict(taskId) = True: foundSpecific = True
                    Case "CORE.ERROR.FORECAST_START_DEPENDENCIES"
                        errForecastConflict(taskId) = True: foundSpecific = True
                    Case "CORE.ERROR.FORECAST_FINISH_CONSTRAINTS"
                        errForecastFinishConflict(taskId) = True: foundSpecific = True
                    Case "CORE.ERROR.BASELINE_DURATION_MISSING"
                        errMissingDuration(taskId) = True: foundSpecific = True
                    Case "CORE.ERROR.START_NOT_COMPUTABLE"
                        errStartNotComputable(taskId) = True: foundSpecific = True
                    Case "CORE.ERROR.FINISH_BEFORE_START", "CORE.ERROR.LOE_FINISH_BEFORE_START"
                        errFinishBeforeStart(taskId) = True: foundSpecific = True
                    Case Else
                        If Left$(code, Len("DIAG.CONSTRAINT.")) = "DIAG.CONSTRAINT." Then
                            errConstraintRootMessages(taskId) = renderedErrorText
                            foundSpecific = True
                        End If
                End Select
            End If
        End If
    Next record

    If Not foundSpecific Then errOtherRoot(taskId) = True

End Sub

