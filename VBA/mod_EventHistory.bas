Attribute VB_Name = "mod_EventHistory"
Option Explicit

Private Const CALC_ALARM_SHEET As String = "CALC_ALARM"
Private Const CALC_ALARM_TABLE As String = "tbl_CALC_ALARM"
Private Const EVENT_HISTORY_SHEET As String = "EVENT_HISTORY"
Private Const EVENT_HISTORY_TABLE As String = "tbl_EVENT_HISTORY"
Private Const EVENT_ACK_SHEET As String = "EVENT_ACK"
Private Const EVENT_ACK_TABLE As String = "tbl_EVENT_ACK"
Private Const EVENT_HISTORY_HEADER_ROW As Long = 3

Private Const EVENT_HISTORY_INFO_LABEL As String = "btn_EventHistory_Info_Label"
Private Const EVENT_HISTORY_INFO_BG As String = "btn_EventHistory_Info_BG"
Private Const EVENT_HISTORY_INFO_KNOB As String = "btn_EventHistory_Info_Knob"
Private Const EVENT_HISTORY_LANG_LABEL As String = "btn_EventHistory_Language_Label"
Private Const EVENT_HISTORY_LANG_BG As String = "btn_EventHistory_Language_BG"
Private Const EVENT_HISTORY_LANG_KNOB As String = "btn_EventHistory_Language_Knob"
Private Const EVENT_ACK_CLEAR_BUTTON As String = "btn_EventAck_ClearList"
Private Const EVENT_HISTORY_CLEAR_BUTTON As String = "btn_EventHistory_ClearHistory"

Private gPlanningEventRunId As String
Private gPlanningEventRunActive As Boolean
Private gEventHistoryShowInfo As Boolean
Private gEventHistoryShowInfoInitialized As Boolean
Private gPlanningLanguage As String
Private gEventHistoryUiStateBootstrapped As Boolean
Private gEventHistoryInternalWriteDepth As Long

Public Sub BeginPlanningEventRun(Optional ByVal sourceProcedure As String = "")

    gPlanningEventRunId = BuildPlanningEventRunId(sourceProcedure)
    gPlanningEventRunActive = True

End Sub

Public Sub EndPlanningEventRun()

    gPlanningEventRunActive = False
    gPlanningEventRunId = vbNullString

End Sub

Public Function GetPlanningConsoleLanguage() As String

    EnsureEventHistoryState
    GetPlanningConsoleLanguage = gPlanningLanguage

End Function

Public Function ShouldShowInfoOnlyPlanningConsole() As Boolean

    EnsureEventHistoryState
    ShouldShowInfoOnlyPlanningConsole = gEventHistoryShowInfo

End Function

Public Sub Toggle_EventHistory_Info()

    Dim wsHistory As Worksheet
    Dim oldScreenUpdating As Boolean

    EnsureEventHistoryState
    gEventHistoryShowInfo = Not gEventHistoryShowInfo

    oldScreenUpdating = Application.ScreenUpdating
    Application.ScreenUpdating = False
    On Error GoTo SafeExit

    Set wsHistory = EnsurePlanningEventSheet(EVENT_HISTORY_SHEET)
    EnsureEventHistoryToggleShapes wsHistory
    RefreshEventHistoryToggleVisuals wsHistory

SafeExit:
    Application.ScreenUpdating = oldScreenUpdating

End Sub

Public Sub Toggle_EventHistory_Language()

    Dim oldScreenUpdating As Boolean

    EnsureEventHistoryState

    If UCase$(Trim$(gPlanningLanguage)) = "EN" Then
        gPlanningLanguage = "FR"
    Else
        gPlanningLanguage = "EN"
    End If

    oldScreenUpdating = Application.ScreenUpdating
    Application.ScreenUpdating = False
    On Error GoTo SafeExit

    Refresh_EventHistory_View

SafeExit:
    Application.ScreenUpdating = oldScreenUpdating

End Sub

Public Sub Handle_EventHistory_Change(ByVal ws As Worksheet, ByVal Target As Range)

    On Error GoTo SafeExit

    If IsPlanningEventInternalWriteActive() Then Exit Sub

    If Not ws Is Nothing Then
        If CStr(ws.Name) = EVENT_ACK_SHEET Then
            If IsAllowedManualEventAckChange(ws, Target) Then
                Application.EnableEvents = False
                NormalizeManualEventAckRows ws, Target
                Refresh_EventHistory_View
                GoTo SafeExit
            End If
        End If
    End If

    Application.EnableEvents = False
    Application.Undo

SafeExit:
    If Not IsMacroAbortRequested() Then
        Application.EnableEvents = True
    End If

End Sub

Private Function IsAllowedManualEventAckChange(ByVal ws As Worksheet, ByVal Target As Range) As Boolean

    Dim tblAck As ListObject
    Dim editableRange As Range

    If ws Is Nothing Then Exit Function
    If Target Is Nothing Then Exit Function

    On Error Resume Next
    Set tblAck = ws.ListObjects(EVENT_ACK_TABLE)
    On Error GoTo 0

    If tblAck Is Nothing Then Exit Function
    If Intersect(Target, tblAck.HeaderRowRange) Is Nothing Then
        If Not tblAck.DataBodyRange Is Nothing Then
            Set editableRange = tblAck.DataBodyRange
        Else
            Set editableRange = tblAck.Range.Offset(1, 0).Resize(1, tblAck.Range.Columns.Count)
        End If

        IsAllowedManualEventAckChange = Not Intersect(Target, editableRange) Is Nothing
    End If

End Function

Private Sub NormalizeManualEventAckRows(ByVal ws As Worksheet, ByVal Target As Range)

    Dim tblAck As ListObject
    Dim touched As Range
    Dim rowRange As Range
    Dim acknowledgedCell As Range
    Dim severityCell As Range

    If ws Is Nothing Then Exit Sub
    If Target Is Nothing Then Exit Sub

    On Error Resume Next
    Set tblAck = ws.ListObjects(EVENT_ACK_TABLE)
    On Error GoTo 0

    If tblAck Is Nothing Then Exit Sub
    If tblAck.DataBodyRange Is Nothing Then Exit Sub

    Set touched = Intersect(Target, tblAck.DataBodyRange)
    If touched Is Nothing Then Exit Sub

    For Each rowRange In touched.rows
        Set acknowledgedCell = Intersect(rowRange.EntireRow, tblAck.ListColumns("Acknowledged").DataBodyRange)
        If Not acknowledgedCell Is Nothing Then
            If Trim$(CStr(acknowledgedCell.Cells(1, 1).value)) = "" Then acknowledgedCell.Cells(1, 1).value = True
        End If

        Set severityCell = Intersect(rowRange.EntireRow, tblAck.ListColumns("Severity").DataBodyRange)
        If Not severityCell Is Nothing Then
            If Trim$(CStr(severityCell.Cells(1, 1).value)) <> "" Then severityCell.Cells(1, 1).value = UCase$(Trim$(CStr(severityCell.Cells(1, 1).value)))
        End If
    Next rowRange

End Sub

Public Sub Refresh_EventHistory_View()

    Dim wsAlarm As Worksheet
    Dim wsHistory As Worksheet
    Dim wsAck As Worksheet
    Dim tblAlarm As ListObject
    Dim tblHistory As ListObject
    Dim tblAck As ListObject
    Dim ackLookup As Object

    Dim r As Long
    Dim newRow As ListRow
    Dim msgText As String
    Dim severity As String
    Dim eventType As String
    Dim eventHash As String
    Dim internalWriteStarted As Boolean
    Dim oldScreenUpdating As Boolean

    oldScreenUpdating = Application.ScreenUpdating
    Application.ScreenUpdating = False
    On Error GoTo CleanFail
    BeginPlanningEventInternalWrite wsAlarm, wsHistory, wsAck
    internalWriteStarted = True
    Set tblAlarm = wsAlarm.ListObjects(CALC_ALARM_TABLE)
    Set tblHistory = wsHistory.ListObjects(EVENT_HISTORY_TABLE)
    Set tblAck = wsAck.ListObjects(EVENT_ACK_TABLE)
    Set ackLookup = BuildEventAckLookup(tblAck)

    ClearPlanningEventTableRows tblHistory

    If Not tblAlarm.DataBodyRange Is Nothing Then
        For r = tblAlarm.ListRows.Count To 1 Step -1
            severity = UCase$(Trim$(CStr(tblAlarm.DataBodyRange.Cells(r, tblAlarm.ListColumns("Severity").Index).value)))
            eventType = Trim$(CStr(tblAlarm.DataBodyRange.Cells(r, tblAlarm.ListColumns("Event Type").Index).value))
            eventHash = Trim$(CStr(tblAlarm.DataBodyRange.Cells(r, tblAlarm.ListColumns("Hash").Index).value))
            msgText = BuildEventHistoryDisplayMessage( _
                CStr(tblAlarm.DataBodyRange.Cells(r, tblAlarm.ListColumns("FR Message").Index).value), _
                CStr(tblAlarm.DataBodyRange.Cells(r, tblAlarm.ListColumns("EN Message").Index).value), _
                CStr(tblAlarm.DataBodyRange.Cells(r, tblAlarm.ListColumns("FR Details").Index).value), _
                CStr(tblAlarm.DataBodyRange.Cells(r, tblAlarm.ListColumns("EN Details").Index).value))

            Set newRow = tblHistory.ListRows.Add
            With newRow.Range
                .Cells(1, tblHistory.ListColumns("Date").Index).value = tblAlarm.DataBodyRange.Cells(r, tblAlarm.ListColumns("Date").Index).value
                .Cells(1, tblHistory.ListColumns("Heure").Index).value = tblAlarm.DataBodyRange.Cells(r, tblAlarm.ListColumns("Time").Index).value
                .Cells(1, tblHistory.ListColumns("Severity").Index).value = severity
                .Cells(1, tblHistory.ListColumns("Message").Index).value = msgText
                .Cells(1, tblHistory.ListColumns("Acknowledged").Index).value = IsEventAcknowledged(ackLookup, severity, eventType, eventHash)
            End With
        Next r
    End If

    ApplyPlanningEventFormats tblAlarm, tblHistory, tblAck
    EnsureEventHistoryToggleShapes wsHistory
    RefreshEventHistoryToggleVisuals wsHistory

CleanExit:
    If internalWriteStarted Then
        EndPlanningEventInternalWrite wsAlarm, wsHistory, wsAck
    End If
    Application.ScreenUpdating = oldScreenUpdating
    Exit Sub

CleanFail:
    Resume CleanExit

End Sub

Public Function FormatPlanningConsoleMessageForCurrentLanguage(ByVal rawMessage As String) As String

    Dim txt As String
    Dim upperTxt As String
    Dim enPos As Long
    Dim enMarkerLen As Long
    Dim frPart As String
    Dim enPart As String

    EnsureEventHistoryState

    txt = TrimPlanningMessageLineEdges(CStr(rawMessage))
    upperTxt = UCase$(txt)

    If Left$(upperTxt, 3) <> "FR:" Then
        FormatPlanningConsoleMessageForCurrentLanguage = TrimPlanningMessageLineEdges(txt)
        Exit Function
    End If

    enPos = InStr(1, upperTxt, vbCrLf & "EN:", vbTextCompare)
    enMarkerLen = Len(vbCrLf & "EN:")
    If enPos = 0 Then
        enPos = InStr(1, upperTxt, vbLf & "EN:", vbTextCompare)
        enMarkerLen = Len(vbLf & "EN:")
    End If
    If enPos = 0 Then
        FormatPlanningConsoleMessageForCurrentLanguage = TrimPlanningMessageLineEdges(txt)
        Exit Function
    End If

    frPart = Mid$(txt, 4, enPos - 4)
    enPart = Mid$(txt, enPos + enMarkerLen)

    If UCase$(Trim$(gPlanningLanguage)) = "EN" Then
        FormatPlanningConsoleMessageForCurrentLanguage = TrimPlanningMessageLineEdges(enPart)
    Else
        FormatPlanningConsoleMessageForCurrentLanguage = TrimPlanningMessageLineEdges(frPart)
    End If

End Function

Public Function BuildPlanningWarningAckToken( _
    ByVal eventType As String, _
    ByVal eventHash As String) As String

    If Trim$(eventType) = "" Then Exit Function
    If Trim$(eventHash) = "" Then Exit Function

    BuildPlanningWarningAckToken = UCase$(Trim$(eventType)) & "|" & Trim$(eventHash)

End Function

Public Function PlanningMessage_IsAcknowledged(ByVal item As Variant) As Boolean

    Dim msgType As String
    Dim ackTokens As String
    Dim tblAck As ListObject
    Dim ackLookup As Object

    On Error GoTo SafeExit

    msgType = NormalizeConsoleEventSeverity(CStr(item("Type")))
    If msgType <> "WARNING" Then Exit Function

    On Error Resume Next
    Set tblAck = ThisWorkbook.Worksheets(EVENT_ACK_SHEET).ListObjects(EVENT_ACK_TABLE)
    On Error GoTo SafeExit
    If tblAck Is Nothing Then
        EnsurePlanningEventHistoryInfrastructure
        Set tblAck = ThisWorkbook.Worksheets(EVENT_ACK_SHEET).ListObjects(EVENT_ACK_TABLE)
    End If

    Set ackLookup = BuildEventAckLookup(tblAck)

    ackTokens = PlanningMessage_BuildAckTokens(item)
    If Trim$(ackTokens) <> "" Then
        PlanningMessage_IsAcknowledged = ArePlanningWarningAckTokensAcknowledged(ackLookup, ackTokens)
        Exit Function
    End If

SafeExit:
End Function

Public Function PlanningMessage_CanAcknowledge(ByVal item As Variant) As Boolean

    On Error GoTo SafeExit

    If NormalizeConsoleEventSeverity(CStr(item("Type"))) <> "WARNING" Then Exit Function
    PlanningMessage_CanAcknowledge = (Trim$(PlanningMessage_BuildAckTokens(item)) <> "")

SafeExit:
End Function

Public Sub SetPlanningWarningAckState( _
    ByVal item As Variant, _
    ByVal acknowledged As Boolean)

    Dim ackTokens As String
    Dim tokens() As String
    Dim oneToken As Variant
    Dim wsAlarm As Worksheet
    Dim wsHistory As Worksheet
    Dim wsAck As Worksheet
    Dim internalWriteStarted As Boolean

    On Error GoTo SafeExit

    If Not PlanningMessage_CanAcknowledge(item) Then Exit Sub

    BeginPlanningEventInternalWrite wsAlarm, wsHistory, wsAck
    internalWriteStarted = True

    ackTokens = PlanningMessage_BuildAckTokens(item)
    tokens = Split(ackTokens, ";")

    For Each oneToken In tokens
        If Trim$(CStr(oneToken)) <> "" Then
            If acknowledged Then
                UpsertPlanningWarningAckToken CStr(oneToken), item
            Else
                RemovePlanningWarningAckToken CStr(oneToken)
            End If
        End If
    Next oneToken

    item("Acknowledged") = PlanningMessage_IsAcknowledged(item)
    Refresh_EventHistory_View

SafeExit:
    If internalWriteStarted Then
        EndPlanningEventInternalWrite wsAlarm, wsHistory, wsAck
    End If
End Sub

Public Sub ClearPlanningWarningAcknowledgements()

    Dim wsAlarm As Worksheet
    Dim wsHistory As Worksheet
    Dim wsAck As Worksheet
    Dim tblAck As ListObject
    Dim internalWriteStarted As Boolean

    On Error GoTo CleanFail
    BeginPlanningEventInternalWrite wsAlarm, wsHistory, wsAck
    internalWriteStarted = True
    Set tblAck = wsAck.ListObjects(EVENT_ACK_TABLE)

    ClearPlanningEventTableRows tblAck

    Refresh_EventHistory_View

CleanExit:
    If internalWriteStarted Then
        EndPlanningEventInternalWrite wsAlarm, wsHistory, wsAck
    End If
    Exit Sub

CleanFail:
    Resume CleanExit

End Sub

Public Sub ClearPlanningEventHistory()

    Dim wsAlarm As Worksheet
    Dim wsHistory As Worksheet
    Dim wsAck As Worksheet
    Dim tblAlarm As ListObject
    Dim tblHistory As ListObject
    Dim internalWriteStarted As Boolean

    On Error GoTo CleanFail
    BeginPlanningEventInternalWrite wsAlarm, wsHistory, wsAck
    internalWriteStarted = True
    Set tblAlarm = wsAlarm.ListObjects(CALC_ALARM_TABLE)
    Set tblHistory = wsHistory.ListObjects(EVENT_HISTORY_TABLE)

    ClearPlanningEventTableRows tblAlarm
    ClearPlanningEventTableRows tblHistory

    Refresh_EventHistory_View

CleanExit:
    If internalWriteStarted Then
        EndPlanningEventInternalWrite wsAlarm, wsHistory, wsAck
    End If
    Exit Sub

CleanFail:
    Resume CleanExit

End Sub

Public Function PlanningMessage_BuildAckTokens(ByVal item As Variant) As String

    Dim rawMessage As String
    Dim eventType As String
    Dim eventHash As String
    Dim frMessage As String
    Dim enMessage As String
    Dim frDetails As String
    Dim enDetails As String

    On Error GoTo SafeExit

    If NormalizeConsoleEventSeverity(CStr(item("Type"))) <> "WARNING" Then Exit Function

    PlanningMessage_BuildAckTokens = PlanningMessage_GetOptionalText(item, "AckTokens")
    If Trim$(PlanningMessage_BuildAckTokens) <> "" Then Exit Function

    eventType = PlanningMessage_GetOptionalText(item, "EventType")
    eventHash = PlanningMessage_GetOptionalText(item, "Hash")

    If eventType = "" Or eventHash = "" Then
        rawMessage = CStr(item("Message"))
        SplitConsoleMessageForHistory rawMessage, frMessage, enMessage, frDetails, enDetails
        eventType = "CONSOLE_WARNING"
        eventHash = BuildPlanningEventHash( _
            "WARNING", eventType, frMessage, enMessage, frDetails, enDetails, _
            "MessageEngine")
    End If

    PlanningMessage_BuildAckTokens = BuildPlanningWarningAckToken(eventType, eventHash)

SafeExit:
End Function

Public Sub EnsurePlanningEventHistoryInfrastructure()

    Dim wsAlarm As Worksheet
    Dim wsHistory As Worksheet
    Dim wsAck As Worksheet
    Dim tblAlarm As ListObject
    Dim tblHistory As ListObject
    Dim tblAck As ListObject
    Dim internalWriteStarted As Boolean

    On Error GoTo CleanFail
    gEventHistoryInternalWriteDepth = gEventHistoryInternalWriteDepth + 1
    internalWriteStarted = True

    Set wsAlarm = EnsurePlanningEventSheet(CALC_ALARM_SHEET)
    Set wsHistory = EnsurePlanningEventSheet(EVENT_HISTORY_SHEET)
    Set wsAck = EnsurePlanningEventSheet(EVENT_ACK_SHEET)

    EnsurePlanningEventTable wsAlarm, CALC_ALARM_TABLE, CalcAlarmHeaders()
    EnsureEventHistoryTopRows wsHistory
    EnsureEventAckTopRow wsAck
    EnsurePlanningEventTable wsHistory, EVENT_HISTORY_TABLE, EventHistoryHeaders(), EVENT_HISTORY_HEADER_ROW
    EnsurePlanningEventTable wsAck, EVENT_ACK_TABLE, EventAckHeaders(), 2

    Set tblAlarm = wsAlarm.ListObjects(CALC_ALARM_TABLE)
    Set tblHistory = wsHistory.ListObjects(EVENT_HISTORY_TABLE)
    Set tblAck = wsAck.ListObjects(EVENT_ACK_TABLE)
    ApplyPlanningEventFormats tblAlarm, tblHistory, tblAck
    EnsureEventHistoryToggleShapes wsHistory
    EnsureEventHistoryCommandButtons wsHistory, wsAck
    RefreshEventHistoryToggleVisuals wsHistory

CleanExit:
    If internalWriteStarted Then
        gEventHistoryInternalWriteDepth = gEventHistoryInternalWriteDepth - 1
        If gEventHistoryInternalWriteDepth < 0 Then gEventHistoryInternalWriteDepth = 0
    End If
    Exit Sub

CleanFail:
    Resume CleanExit

End Sub

Public Sub LogPlanningEvent( _
    ByVal severity As String, _
    ByVal eventType As String, _
    ByVal eventHash As String, _
    ByVal frMessage As String, _
    ByVal enMessage As String, _
    ByVal frDetails As String, _
    ByVal enDetails As String, _
    ByVal sourceProcedure As String, _
    Optional ByVal sourceSheet As String = "", _
    Optional ByVal sourceTable As String = "", _
    Optional ByVal taskId As String = "", _
    Optional ByVal wbsValue As String = "", _
    Optional ByVal taskName As String = "", _
    Optional ByVal refreshView As Boolean = True)

    Dim wsAlarm As Worksheet
    Dim wsHistory As Worksheet
    Dim wsAck As Worksheet
    Dim tblAlarm As ListObject
    Dim tblHistory As ListObject
    Dim alarmRow As ListRow
    Dim eventTs As Date
    Dim finalHash As String
    Dim normalizedWbs As String
    Dim internalWriteStarted As Boolean

    On Error GoTo CleanFail
    If Trim$(severity) = "" Then Exit Sub
    If Trim$(eventType) = "" Then Exit Sub

    EnsurePlanningEventRunId sourceProcedure
    frMessage = TrimPlanningMessageLineEdges(frMessage)
    enMessage = TrimPlanningMessageLineEdges(enMessage)
    frDetails = TrimPlanningMessageLineEdges(frDetails)
    enDetails = TrimPlanningMessageLineEdges(enDetails)
    normalizedWbs = NormalizeWBS(wbsValue)

    BeginPlanningEventInternalWrite wsAlarm, wsHistory, wsAck
    internalWriteStarted = True
    Set tblAlarm = wsAlarm.ListObjects(CALC_ALARM_TABLE)
    Set tblHistory = wsHistory.ListObjects(EVENT_HISTORY_TABLE)

    eventTs = Now
    finalHash = Trim$(eventHash)
    If finalHash = "" Then
        finalHash = BuildPlanningEventHash( _
            severity, eventType, frMessage, enMessage, frDetails, enDetails, _
            sourceProcedure, sourceSheet, sourceTable, taskId, normalizedWbs, taskName)
    End If

    Set alarmRow = tblAlarm.ListRows.Add
    With alarmRow.Range
        .Cells(1, 1).value = BuildPlanningEventId(tblAlarm.ListRows.Count, eventTs)
        .Cells(1, 2).value = gPlanningEventRunId
        .Cells(1, 3).value = eventTs
        .Cells(1, 4).value = dateValue(eventTs)
        .Cells(1, 5).value = TimeValue(eventTs)
        .Cells(1, 6).value = UCase$(Trim$(severity))
        .Cells(1, 7).value = Trim$(eventType)
        .Cells(1, 8).value = finalHash
        .Cells(1, 9).value = False
        .Cells(1, 10).value = Trim$(sourceSheet)
        .Cells(1, 11).value = Trim$(sourceTable)
        .Cells(1, 12).value = Trim$(taskId)
        .Cells(1, 13).NumberFormat = "@"
        .Cells(1, 13).value = normalizedWbs
        .Cells(1, 14).value = Trim$(taskName)
        .Cells(1, 15).value = frMessage
        .Cells(1, 16).value = enMessage
        .Cells(1, 17).value = frDetails
        .Cells(1, 18).value = enDetails
        .Cells(1, 19).value = Trim$(sourceProcedure)
    End With

    If refreshView Then
        Refresh_EventHistory_View
    End If

CleanExit:
    If internalWriteStarted Then
        EndPlanningEventInternalWrite wsAlarm, wsHistory, wsAck
    End If
    Exit Sub

CleanFail:
    Resume CleanExit

End Sub

Public Function PlanningEvents_LogConsoleMessagesSafe( _
    ByVal consoleMessages As Collection, _
    ByVal sourceProcedure As String, _
    ByRef errorMessage As String) As Boolean

    Dim temporaryWorkflowStarted As Boolean

    On Error GoTo Fail

    errorMessage = vbNullString

    If consoleMessages Is Nothing Then
        PlanningEvents_LogConsoleMessagesSafe = True
        Exit Function
    End If

    temporaryWorkflowStarted = EnsurePlanningWorkflowStarted(sourceProcedure)

    If IsMacroRunActive() Then
        EnsurePlanningEventRunId sourceProcedure
    Else
        BeginPlanningEventRun sourceProcedure
    End If

    LogPlanningConsoleMessages consoleMessages, sourceProcedure
    PlanningEvents_LogConsoleMessagesSafe = True

CleanExit:
    If temporaryWorkflowStarted Then EndPlanningWorkflow
    Exit Function

Fail:
    errorMessage = BiMsg( _
        "Erreur pendant l'historisation de la console runtime" & vbCrLf & _
        "-> " & Err.Description, _
        "Error while logging the runtime console" & vbCrLf & _
        "-> " & Err.Description)
    PlanningEvents_LogConsoleMessagesSafe = False
    Resume CleanExit

End Function
Public Sub LogPlanningConsoleMessages( _
    ByVal consoleMessages As Collection, _
    Optional ByVal sourceProcedure As String = "CalcBridge_ShowPlanningConsole")

    Dim item As Variant
    Dim severity As String
    Dim eventType As String
    Dim rawMessage As String
    Dim frMessage As String
    Dim enMessage As String
    Dim frDetails As String
    Dim enDetails As String
    Dim seenInConsole As Object
    Dim localKey As String
    Dim wsAlarm As Worksheet
    Dim wsHistory As Worksheet
    Dim wsAck As Worksheet
    Dim internalWriteStarted As Boolean

    If consoleMessages Is Nothing Then Exit Sub
    If consoleMessages.Count = 0 Then Exit Sub

    On Error GoTo CleanFail
    BeginPlanningEventInternalWrite wsAlarm, wsHistory, wsAck
    internalWriteStarted = True
    EnsurePlanningEventRunId sourceProcedure
    Set seenInConsole = CreateObject("Scripting.Dictionary")

    For Each item In consoleMessages
        If Not ConsoleMessageHistoryHandled(item) Then
            severity = NormalizeConsoleEventSeverity(CStr(item("Type")))
            If severity <> "" Then
                rawMessage = CStr(item("Message"))
                SplitConsoleMessageForHistory rawMessage, frMessage, enMessage, frDetails, enDetails
                eventType = PlanningMessage_GetOptionalText(item, "EventType")
                If eventType = "" Then eventType = "CONSOLE_" & severity
                localKey = BuildConsoleLocalDedupKey(severity, frMessage, enMessage, frDetails, enDetails)

                If Not seenInConsole.Exists(localKey) Then
                    seenInConsole(localKey) = True

                    LogPlanningEvent _
                        severity, _
                        eventType, _
                        PlanningMessage_GetOptionalText(item, "Hash"), _
                        frMessage, _
                        enMessage, _
                        frDetails, _
                        enDetails, _
                        sourceProcedure, _
                        vbNullString, _
                        vbNullString, _
                        vbNullString, _
                        vbNullString, _
                        vbNullString, _
                        False
                End If
            End If
        End If
    Next item

    Refresh_EventHistory_View

CleanExit:
    If internalWriteStarted Then
        EndPlanningEventInternalWrite wsAlarm, wsHistory, wsAck
    End If
    Exit Sub

CleanFail:
    Resume CleanExit

End Sub

Public Function BuildPlanningEventHash( _
    ByVal severity As String, _
    ByVal eventType As String, _
    ByVal frMessage As String, _
    ByVal enMessage As String, _
    ByVal frDetails As String, _
    ByVal enDetails As String, _
    ByVal sourceProcedure As String, _
    Optional ByVal sourceSheet As String = "", _
    Optional ByVal sourceTable As String = "", _
    Optional ByVal taskId As String = "", _
    Optional ByVal wbsValue As String = "", _
    Optional ByVal taskName As String = "") As String

    Dim rawText As String

    rawText = BuildPlanningEventSignature( _
        severity, eventType, frMessage, enMessage, frDetails, enDetails, _
        sourceSheet, sourceTable, taskId, wbsValue, taskName)

    BuildPlanningEventHash = HashPlanningEventSignature(rawText)

End Function

Public Function BuildPlanningEventSignature( _
    ByVal severity As String, _
    ByVal eventType As String, _
    ByVal frMessage As String, _
    ByVal enMessage As String, _
    ByVal frDetails As String, _
    ByVal enDetails As String, _
    Optional ByVal sourceSheet As String = "", _
    Optional ByVal sourceTable As String = "", _
    Optional ByVal taskId As String = "", _
    Optional ByVal wbsValue As String = "", _
    Optional ByVal taskName As String = "") As String

    BuildPlanningEventSignature = _
        UCase$(Trim$(severity)) & "|" & _
        UCase$(Trim$(eventType)) & "|" & _
        UCase$(Trim$(sourceSheet)) & "|" & _
        UCase$(Trim$(sourceTable)) & "|" & _
        UCase$(Trim$(taskId)) & "|" & _
        UCase$(Trim$(wbsValue)) & "|" & _
        UCase$(Trim$(taskName)) & "|" & _
        NormalizePlanningEventSignatureText(FirstPlanningEventMessageLine(frMessage)) & "|" & _
        NormalizePlanningEventSignatureText(FirstPlanningEventMessageLine(enMessage)) & "|" & _
        NormalizePlanningEventSignatureText(ExtractPlanningEventContextLine(frMessage, "IDS")) & "|" & _
        NormalizePlanningEventSignatureText(ExtractPlanningEventContextLine(enMessage, "IDS")) & "|" & _
        NormalizePlanningEventSignatureText(ExtractPlanningEventContextLine(frMessage, "ID")) & "|" & _
        NormalizePlanningEventSignatureText(ExtractPlanningEventContextLine(enMessage, "ID")) & "|" & _
        NormalizePlanningEventSignatureText(ExtractPlanningEventContextLine(frMessage, "WBS")) & "|" & _
        NormalizePlanningEventSignatureText(ExtractPlanningEventContextLine(enMessage, "WBS")) & "|" & _
        NormalizePlanningEventSignatureText(ExtractPlanningEventContextLine(frMessage, "TASK")) & "|" & _
        NormalizePlanningEventSignatureText(ExtractPlanningEventContextLine(enMessage, "TASK"))

End Function

Private Function HashPlanningEventSignature(ByVal rawText As String) As String

    Dim i As Long
    Dim acc As Double
    Dim nextVal As Double

    acc = 5381#
    For i = 1 To Len(rawText)
        nextVal = (acc * 33#) + AscW(Mid$(rawText, i, 1))
        acc = nextVal - (Fix(nextVal / 2147483647#) * 2147483647#)
    Next i

    HashPlanningEventSignature = "EVH-" & Format$(CLng(acc), "0000000000")

End Function

Private Function FirstPlanningEventMessageLine(ByVal value As String) As String

    Dim lines() As String
    Dim i As Long
    Dim txt As String

    txt = Replace(TrimPlanningMessageLineEdges(value), vbCrLf, vbLf)
    txt = Replace(txt, vbCr, vbLf)
    lines = Split(txt, vbLf)

    For i = LBound(lines) To UBound(lines)
        If Trim$(lines(i)) <> "" And Left$(Trim$(lines(i)), 2) <> "->" Then
            FirstPlanningEventMessageLine = Trim$(lines(i))
            Exit Function
        End If
    Next i

End Function

Private Function ExtractPlanningEventContextLine( _
    ByVal value As String, _
    ByVal labelName As String) As String

    Dim lines() As String
    Dim i As Long
    Dim txt As String
    Dim lineText As String
    Dim labelUpper As String

    txt = Replace(TrimPlanningMessageLineEdges(value), vbCrLf, vbLf)
    txt = Replace(txt, vbCr, vbLf)
    lines = Split(txt, vbLf)
    labelUpper = UCase$(Trim$(labelName))

    For i = LBound(lines) To UBound(lines)
        lineText = Trim$(lines(i))
        If UCase$(Left$(lineText, Len(labelUpper) + 1)) = labelUpper & ":" Then
            ExtractPlanningEventContextLine = Trim$(Mid$(lineText, Len(labelUpper) + 2))
            Exit Function
        End If
        If UCase$(Left$(lineText, Len(labelUpper) + 2)) = labelUpper & " :" Then
            ExtractPlanningEventContextLine = Trim$(Mid$(lineText, Len(labelUpper) + 3))
            Exit Function
        End If
    Next i

End Function

Private Function NormalizePlanningEventSignatureText(ByVal value As String) As String

    Dim txt As String

    txt = UCase$(TrimPlanningMessageLineEdges(value))
    txt = Replace(txt, vbTab, " ")

    Do While InStr(1, txt, "  ", vbBinaryCompare) > 0
        txt = Replace(txt, "  ", " ")
    Loop

    NormalizePlanningEventSignatureText = Trim$(txt)

End Function

Private Function ConsoleMessageHistoryHandled(ByVal item As Variant) As Boolean

    On Error Resume Next
    ConsoleMessageHistoryHandled = CBool(item("HistoryHandled"))
    On Error GoTo 0

End Function

Private Function NormalizeConsoleEventSeverity(ByVal msgType As String) As String

    Select Case UCase$(Trim$(msgType))
        Case "ERROR", "STOP"
            NormalizeConsoleEventSeverity = "STOP"
        Case "WARNING"
            NormalizeConsoleEventSeverity = "WARNING"
        Case "INFO"
            NormalizeConsoleEventSeverity = "INFO"
        Case Else
            NormalizeConsoleEventSeverity = vbNullString
    End Select

End Function

Private Sub SplitConsoleMessageForHistory( _
    ByVal rawMessage As String, _
    ByRef frMessage As String, _
    ByRef enMessage As String, _
    ByRef frDetails As String, _
    ByRef enDetails As String)

    Dim txt As String
    Dim upperTxt As String
    Dim enPos As Long
    Dim enMarkerLen As Long

    txt = TrimPlanningMessageLineEdges(CStr(rawMessage))
    upperTxt = UCase$(txt)
    frDetails = vbNullString
    enDetails = vbNullString

    If Left$(upperTxt, 3) <> "FR:" Then
        frMessage = txt
        enMessage = txt
        Exit Sub
    End If

    enPos = InStr(1, upperTxt, vbCrLf & "EN:", vbTextCompare)
    enMarkerLen = Len(vbCrLf & "EN:")
    If enPos = 0 Then
        enPos = InStr(1, upperTxt, vbLf & "EN:", vbTextCompare)
        enMarkerLen = Len(vbLf & "EN:")
    End If

    If enPos = 0 Then
        frMessage = txt
        enMessage = txt
        Exit Sub
    End If

    frMessage = TrimPlanningMessageLineEdges(Mid$(txt, 4, enPos - 4))
    enMessage = TrimPlanningMessageLineEdges(Mid$(txt, enPos + enMarkerLen))

End Sub

Private Function TrimPlanningMessageLineEdges(ByVal value As String) As String

    Dim txt As String

    txt = CStr(value)

    Do While Len(txt) > 0 And (Left$(txt, 1) = vbCr Or Left$(txt, 1) = vbLf)
        txt = Mid$(txt, 2)
    Loop

    Do While Len(txt) > 0 And (Right$(txt, 1) = vbCr Or Right$(txt, 1) = vbLf)
        txt = Left$(txt, Len(txt) - 1)
    Loop

    TrimPlanningMessageLineEdges = txt

End Function

Private Function BuildConsoleLocalDedupKey( _
    ByVal severity As String, _
    ByVal frMessage As String, _
    ByVal enMessage As String, _
    ByVal frDetails As String, _
    ByVal enDetails As String) As String

    BuildConsoleLocalDedupKey = _
        UCase$(Trim$(severity)) & "|" & _
        Trim$(frMessage) & "|" & _
        Trim$(enMessage) & "|" & _
        Trim$(frDetails) & "|" & _
        Trim$(enDetails)

End Function
Private Sub EnsurePlanningEventRunId(ByVal sourceProcedure As String)

    If Trim$(gPlanningEventRunId) = "" Or Not gPlanningEventRunActive Then
        gPlanningEventRunId = BuildPlanningEventRunId(sourceProcedure)
        gPlanningEventRunActive = True
    End If

End Sub

Private Function BuildPlanningEventRunId(ByVal sourceProcedure As String) As String

    BuildPlanningEventRunId = _
        "RUN-" & Format$(Now, "yyyymmdd-hhnnss") & _
        IIf(Trim$(sourceProcedure) = "", "", "-" & Trim$(sourceProcedure))

End Function

Private Function BuildPlanningEventId(ByVal rowCount As Long, ByVal eventTs As Date) As String

    BuildPlanningEventId = _
        "EVT-" & Format$(eventTs, "yyyymmdd-hhnnss") & "-" & Format$(rowCount, "000000")

End Function

Private Function EnsurePlanningEventSheet(ByVal sheetName As String) As Worksheet

    On Error Resume Next
    Set EnsurePlanningEventSheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If EnsurePlanningEventSheet Is Nothing Then
        Set EnsurePlanningEventSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        EnsurePlanningEventSheet.Name = sheetName
    End If

End Function

Private Sub EnsurePlanningEventTable( _
    ByVal ws As Worksheet, _
    ByVal tableName As String, _
    ByVal headers As Variant, _
    Optional ByVal headerRow As Long = 1)

    Dim tbl As ListObject
    Dim rng As Range
    Dim i As Long

    On Error Resume Next
    Set tbl = ws.ListObjects(tableName)
    On Error GoTo 0

    If tbl Is Nothing Then
        Set rng = ws.Range(ws.Cells(headerRow, 1), ws.Cells(headerRow, UBound(headers) + 1))
        For i = LBound(headers) To UBound(headers)
            rng.Cells(1, i + 1).value = CStr(headers(i))
        Next i
        Set tbl = ws.ListObjects.Add(xlSrcRange, rng, , xlYes)
        tbl.Name = tableName
    Else
        EnsurePlanningEventTableHeaders tbl, headers
    End If

End Sub

Private Sub EnsureEventHistoryState()

    BootstrapEventHistoryUiStateFromSheet

    If Not gEventHistoryShowInfoInitialized Then
        gEventHistoryShowInfo = True
        gEventHistoryShowInfoInitialized = True
    End If

    If Trim$(gPlanningLanguage) = "" Then gPlanningLanguage = "FR"

End Sub

Private Sub BootstrapEventHistoryUiStateFromSheet()

    Dim ws As Worksheet
    Dim isOn As Boolean

    If gEventHistoryUiStateBootstrapped Then Exit Sub
    gEventHistoryUiStateBootstrapped = True

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(EVENT_HISTORY_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    If TryReadEventHistoryTwoStateToggle(ws, EVENT_HISTORY_INFO_BG, EVENT_HISTORY_INFO_KNOB, isOn) Then
        gEventHistoryShowInfo = isOn
        gEventHistoryShowInfoInitialized = True
    End If

    If TryReadEventHistoryTwoStateToggle(ws, EVENT_HISTORY_LANG_BG, EVENT_HISTORY_LANG_KNOB, isOn) Then
        If isOn Then
            gPlanningLanguage = "EN"
        Else
            gPlanningLanguage = "FR"
        End If
    End If

End Sub

Private Function TryReadEventHistoryTwoStateToggle( _
    ByVal ws As Worksheet, _
    ByVal bgName As String, _
    ByVal knobName As String, _
    ByRef isOn As Boolean) As Boolean

    Dim bg As Shape
    Dim knob As Shape

    On Error Resume Next
    Set bg = ws.Shapes(bgName)
    Set knob = ws.Shapes(knobName)
    On Error GoTo 0

    If bg Is Nothing Then Exit Function
    If knob Is Nothing Then Exit Function

    isOn = ((knob.Left + (knob.Width / 2)) >= (bg.Left + (bg.Width / 2)))
    TryReadEventHistoryTwoStateToggle = True

End Function

Private Sub EnsureEventHistoryTopRows(ByVal ws As Worksheet)

    Dim tbl As ListObject

    On Error Resume Next
    Set tbl = ws.ListObjects(EVENT_HISTORY_TABLE)
    On Error GoTo 0

    If Not tbl Is Nothing Then
        If tbl.HeaderRowRange.Row < EVENT_HISTORY_HEADER_ROW Then
            ws.rows("1:2").Insert Shift:=xlDown
        End If
    End If

    ws.rows("1:2").rowHeight = 22

End Sub

Private Sub EnsureEventAckTopRow(ByVal ws As Worksheet)

    Dim tbl As ListObject

    On Error Resume Next
    Set tbl = ws.ListObjects(EVENT_ACK_TABLE)
    On Error GoTo 0

    If Not tbl Is Nothing Then
        If tbl.HeaderRowRange.Row < 2 Then ws.rows("1:1").Insert Shift:=xlDown
    End If

    ws.rows("1:1").rowHeight = 22

End Sub

Private Function BuildEventHistoryDisplayMessage( _
    ByVal frMessage As String, _
    ByVal enMessage As String, _
    ByVal frDetails As String, _
    ByVal enDetails As String) As String

    Dim msg As String
    Dim details As String

    EnsureEventHistoryState

    If UCase$(Trim$(gPlanningLanguage)) = "EN" Then
        msg = TrimPlanningMessageLineEdges(enMessage)
        details = TrimPlanningMessageLineEdges(enDetails)
    Else
        msg = TrimPlanningMessageLineEdges(frMessage)
        details = TrimPlanningMessageLineEdges(frDetails)
    End If

    If details <> "" Then
        If msg <> "" Then
            msg = msg & vbLf & "-> " & details
        Else
            msg = details
        End If
    End If

    BuildEventHistoryDisplayMessage = msg

End Function

Private Function BuildEventAckLookup(ByVal tblAck As ListObject) As Object

    Dim lookup As Object
    Dim r As Long
    Dim severity As String
    Dim eventType As String
    Dim eventHash As String

    Set lookup = CreateObject("Scripting.Dictionary")

    If tblAck Is Nothing Then
        Set BuildEventAckLookup = lookup
        Exit Function
    End If

    If tblAck.DataBodyRange Is Nothing Then
        Set BuildEventAckLookup = lookup
        Exit Function
    End If

    For r = 1 To tblAck.ListRows.Count
        severity = UCase$(Trim$(CStr(tblAck.DataBodyRange.Cells(r, tblAck.ListColumns("Severity").Index).value)))
        eventType = Trim$(CStr(tblAck.DataBodyRange.Cells(r, tblAck.ListColumns("Event Type").Index).value))
        eventHash = Trim$(CStr(tblAck.DataBodyRange.Cells(r, tblAck.ListColumns("Hash").Index).value))

        If severity = "WARNING" And IsTruthy(tblAck.DataBodyRange.Cells(r, tblAck.ListColumns("Acknowledged").Index).value) Then
            If eventType <> "" And eventHash <> "" Then
                lookup(BuildEventAckKey(severity, eventType, eventHash)) = True
            End If
        End If
    Next r

    Set BuildEventAckLookup = lookup

End Function

Private Function IsEventAcknowledged( _
    ByVal ackLookup As Object, _
    ByVal severity As String, _
    ByVal eventType As String, _
    ByVal eventHash As String) As Boolean

    If ackLookup Is Nothing Then Exit Function
    If UCase$(Trim$(severity)) <> "WARNING" Then Exit Function
    If Trim$(eventType) = "" Then Exit Function
    If Trim$(eventHash) = "" Then Exit Function

    IsEventAcknowledged = ackLookup.Exists(BuildEventAckKey(severity, eventType, eventHash))

End Function

Private Function BuildEventAckKey( _
    ByVal severity As String, _
    ByVal eventType As String, _
    ByVal eventHash As String) As String

    BuildEventAckKey = _
        UCase$(Trim$(severity)) & "|" & _
        UCase$(Trim$(eventType)) & "|" & _
        Trim$(eventHash)

End Function

Private Function ArePlanningWarningAckTokensAcknowledged( _
    ByVal ackLookup As Object, _
    ByVal ackTokens As String) As Boolean

    Dim tokens() As String
    Dim oneToken As Variant
    Dim parts() As String
    Dim eventType As String
    Dim eventHash As String
    Dim foundToken As Boolean

    If ackLookup Is Nothing Then Exit Function

    tokens = Split(ackTokens, ";")

    For Each oneToken In tokens
        If Trim$(CStr(oneToken)) <> "" Then
            parts = Split(CStr(oneToken), "|")
            If UBound(parts) <> 1 Then Exit Function

            eventType = Trim$(parts(0))
            eventHash = Trim$(parts(1))
            If eventType = "" Or eventHash = "" Then Exit Function

            foundToken = True
            If Not ackLookup.Exists(BuildEventAckKey("WARNING", eventType, eventHash)) Then
                ArePlanningWarningAckTokensAcknowledged = False
                Exit Function
            End If
        End If
    Next oneToken

    ArePlanningWarningAckTokensAcknowledged = foundToken

End Function

Private Sub UpsertPlanningWarningAckToken( _
    ByVal ackToken As String, _
    ByVal item As Variant)

    Dim tblAck As ListObject
    Dim parts() As String
    Dim eventType As String
    Dim eventHash As String
    Dim r As Long
    Dim targetRow As ListRow
    Dim rawMessage As String
    Dim frMessage As String
    Dim enMessage As String
    Dim frDetails As String
    Dim enDetails As String
    Dim wbsValue As String
    Dim taskName As String

    parts = Split(CStr(ackToken), "|")
    If UBound(parts) <> 1 Then Exit Sub

    eventType = Trim$(parts(0))
    eventHash = Trim$(parts(1))
    If eventType = "" Or eventHash = "" Then Exit Sub

    rawMessage = PlanningMessage_GetOptionalText(item, "Message")
    SplitConsoleMessageForHistory rawMessage, frMessage, enMessage, frDetails, enDetails
    wbsValue = NormalizeWBS(ExtractPlanningEventContextLine(frMessage, "WBS"))
    taskName = ExtractPlanningEventContextLine(frMessage, "Task")
    If taskName = "" Then taskName = ExtractPlanningEventContextLine(enMessage, "Task")

    Set tblAck = ThisWorkbook.Worksheets(EVENT_ACK_SHEET).ListObjects(EVENT_ACK_TABLE)
    If taskName = "" Then taskName = FindEventTaskNameFromAlarm(eventType, eventHash)
    If wbsValue = "" Then wbsValue = NormalizeWBS(FindEventWbsFromAlarm(eventType, eventHash))

    If Not tblAck.DataBodyRange Is Nothing Then
        For r = 1 To tblAck.ListRows.Count
            If UCase$(Trim$(CStr(tblAck.DataBodyRange.Cells(r, tblAck.ListColumns("Severity").Index).value))) = "WARNING" And _
               UCase$(Trim$(CStr(tblAck.DataBodyRange.Cells(r, tblAck.ListColumns("Event Type").Index).value))) = UCase$(eventType) And _
               Trim$(CStr(tblAck.DataBodyRange.Cells(r, tblAck.ListColumns("Hash").Index).value)) = eventHash Then

                Set targetRow = tblAck.ListRows(r)
                Exit For
            End If
        Next r
    End If

    If targetRow Is Nothing Then Set targetRow = tblAck.ListRows.Add

    With targetRow.Range
        .Cells(1, tblAck.ListColumns("Hash").Index).value = eventHash
        .Cells(1, tblAck.ListColumns("Severity").Index).value = "WARNING"
        .Cells(1, tblAck.ListColumns("Event Type").Index).value = eventType
        .Cells(1, tblAck.ListColumns("Acknowledged").Index).value = True
        .Cells(1, tblAck.ListColumns("Acknowledged At").Index).value = Now
        .Cells(1, tblAck.ListColumns("Acknowledged By").Index).value = Environ$("USERNAME")
        .Cells(1, tblAck.ListColumns("WBS").Index).value = wbsValue
        .Cells(1, tblAck.ListColumns("Task Name").Index).value = taskName
        .Cells(1, tblAck.ListColumns("FR Message").Index).value = frMessage
        .Cells(1, tblAck.ListColumns("EN Message").Index).value = enMessage
    End With

End Sub

Private Function FindEventTaskNameFromAlarm( _
    ByVal eventType As String, _
    ByVal eventHash As String) As String

    Dim tblAlarm As ListObject
    Dim r As Long

    On Error GoTo SafeExit

    Set tblAlarm = ThisWorkbook.Worksheets(CALC_ALARM_SHEET).ListObjects(CALC_ALARM_TABLE)
    If tblAlarm.DataBodyRange Is Nothing Then Exit Function

    For r = tblAlarm.ListRows.Count To 1 Step -1
        If UCase$(Trim$(CStr(tblAlarm.DataBodyRange.Cells(r, tblAlarm.ListColumns("Severity").Index).value))) = "WARNING" And _
           UCase$(Trim$(CStr(tblAlarm.DataBodyRange.Cells(r, tblAlarm.ListColumns("Event Type").Index).value))) = UCase$(Trim$(eventType)) And _
           Trim$(CStr(tblAlarm.DataBodyRange.Cells(r, tblAlarm.ListColumns("Hash").Index).value)) = Trim$(eventHash) Then

            FindEventTaskNameFromAlarm = Trim$(CStr(tblAlarm.DataBodyRange.Cells(r, tblAlarm.ListColumns("Task Name").Index).value))
            Exit Function
        End If
    Next r

SafeExit:
End Function

Private Function FindEventWbsFromAlarm( _
    ByVal eventType As String, _
    ByVal eventHash As String) As String

    Dim tblAlarm As ListObject
    Dim r As Long

    On Error GoTo SafeExit

    Set tblAlarm = ThisWorkbook.Worksheets(CALC_ALARM_SHEET).ListObjects(CALC_ALARM_TABLE)
    If tblAlarm.DataBodyRange Is Nothing Then Exit Function

    For r = tblAlarm.ListRows.Count To 1 Step -1
        If UCase$(Trim$(CStr(tblAlarm.DataBodyRange.Cells(r, tblAlarm.ListColumns("Severity").Index).value))) = "WARNING" And _
           UCase$(Trim$(CStr(tblAlarm.DataBodyRange.Cells(r, tblAlarm.ListColumns("Event Type").Index).value))) = UCase$(Trim$(eventType)) And _
           Trim$(CStr(tblAlarm.DataBodyRange.Cells(r, tblAlarm.ListColumns("Hash").Index).value)) = Trim$(eventHash) Then

            FindEventWbsFromAlarm = Trim$(CStr(tblAlarm.DataBodyRange.Cells(r, tblAlarm.ListColumns("WBS").Index).value))
            Exit Function
        End If
    Next r

SafeExit:
End Function

Private Sub RemovePlanningWarningAckToken(ByVal ackToken As String)

    Dim tblAck As ListObject
    Dim parts() As String
    Dim eventType As String
    Dim eventHash As String
    Dim r As Long

    parts = Split(CStr(ackToken), "|")
    If UBound(parts) <> 1 Then Exit Sub

    eventType = Trim$(parts(0))
    eventHash = Trim$(parts(1))
    If eventType = "" Or eventHash = "" Then Exit Sub

    Set tblAck = ThisWorkbook.Worksheets(EVENT_ACK_SHEET).ListObjects(EVENT_ACK_TABLE)
    If tblAck.DataBodyRange Is Nothing Then Exit Sub

    For r = tblAck.ListRows.Count To 1 Step -1
        If UCase$(Trim$(CStr(tblAck.DataBodyRange.Cells(r, tblAck.ListColumns("Severity").Index).value))) = "WARNING" And _
           UCase$(Trim$(CStr(tblAck.DataBodyRange.Cells(r, tblAck.ListColumns("Event Type").Index).value))) = UCase$(eventType) And _
           Trim$(CStr(tblAck.DataBodyRange.Cells(r, tblAck.ListColumns("Hash").Index).value)) = eventHash Then

            tblAck.ListRows(r).Delete
        End If
    Next r

End Sub

Private Function PlanningMessage_GetOptionalText( _
    ByVal item As Variant, _
    ByVal keyName As String) As String

    On Error Resume Next
    PlanningMessage_GetOptionalText = Trim$(CStr(item(keyName)))
    On Error GoTo 0

End Function

Private Function IsTruthy(ByVal value As Variant) As Boolean

    Dim txt As String

    If VarType(value) = vbBoolean Then
        IsTruthy = CBool(value)
        Exit Function
    End If

    txt = UCase$(Trim$(CStr(value)))
    IsTruthy = (txt = "TRUE" Or txt = "YES" Or txt = "OUI" Or txt = "1")

End Function

Private Sub EnsurePlanningEventTableHeaders(ByVal tbl As ListObject, ByVal headers As Variant)

    Dim i As Long
    Dim headerName As String

    For i = LBound(headers) To UBound(headers)
        headerName = CStr(headers(i))
        If Not TableHasColumn(tbl, headerName) Then tbl.ListColumns.Add.Name = headerName
    Next i

End Sub

Private Function TableHasColumn(ByVal tbl As ListObject, ByVal columnName As String) As Boolean

    On Error Resume Next
    TableHasColumn = Not tbl.ListColumns(columnName) Is Nothing
    On Error GoTo 0

End Function

Private Sub ApplyPlanningEventFormats( _
    ByVal tblAlarm As ListObject, _
    ByVal tblHistory As ListObject, _
    ByVal tblAck As ListObject)

    Dim r As Long
    Dim sev As String

    On Error Resume Next

    tblAlarm.ListColumns("Event ID").Range.NumberFormat = "@"
    tblAlarm.ListColumns("Run ID").Range.NumberFormat = "@"
    tblAlarm.ListColumns("Severity").Range.NumberFormat = "@"
    tblAlarm.ListColumns("Event Type").Range.NumberFormat = "@"
    tblAlarm.ListColumns("Hash").Range.NumberFormat = "@"
    tblAlarm.ListColumns("Sheet").Range.NumberFormat = "@"
    tblAlarm.ListColumns("Table").Range.NumberFormat = "@"
    tblAlarm.ListColumns("ID").Range.NumberFormat = "@"
    tblAlarm.ListColumns("WBS").Range.NumberFormat = "@"
    tblAlarm.ListColumns("Task Name").Range.NumberFormat = "@"
    tblAlarm.ListColumns("FR Message").Range.NumberFormat = "@"
    tblAlarm.ListColumns("EN Message").Range.NumberFormat = "@"
    tblAlarm.ListColumns("FR Details").Range.NumberFormat = "@"
    tblAlarm.ListColumns("EN Details").Range.NumberFormat = "@"
    tblAlarm.ListColumns("Source Procedure").Range.NumberFormat = "@"

    If Not tblAlarm.DataBodyRange Is Nothing Then
        tblAlarm.ListColumns("Timestamp").DataBodyRange.NumberFormat = "dd/mm/yyyy hh:mm:ss"
        tblAlarm.ListColumns("Date").DataBodyRange.NumberFormat = "dd/mm/yyyy"
        tblAlarm.ListColumns("Time").DataBodyRange.NumberFormat = "hh:mm:ss"
    End If

    tblHistory.ListColumns("Severity").Range.NumberFormat = "@"
    tblHistory.ListColumns("Message").Range.NumberFormat = "@"
    tblAck.ListColumns("Hash").Range.NumberFormat = "@"
    tblAck.ListColumns("Severity").Range.NumberFormat = "@"
    tblAck.ListColumns("Event Type").Range.NumberFormat = "@"
    tblAck.ListColumns("Acknowledged By").Range.NumberFormat = "@"
    tblAck.ListColumns("Comment").Range.NumberFormat = "@"
    tblAck.ListColumns("WBS").Range.NumberFormat = "@"
    tblAck.ListColumns("Task Name").Range.NumberFormat = "@"
    tblAck.ListColumns("FR Message").Range.NumberFormat = "@"
    tblAck.ListColumns("EN Message").Range.NumberFormat = "@"

    If Not tblHistory.DataBodyRange Is Nothing Then
        tblHistory.ListColumns("Date").DataBodyRange.NumberFormat = "dd/mm/yyyy"
        tblHistory.ListColumns("Heure").DataBodyRange.NumberFormat = "hh:mm:ss"
        tblHistory.ListColumns("Message").DataBodyRange.WrapText = True

        For r = 1 To tblHistory.ListRows.Count
            sev = UCase$(Trim$(CStr(tblHistory.DataBodyRange.Cells(r, tblHistory.ListColumns("Severity").Index).value)))
            Select Case sev
                Case "ERROR", "STOP"
                    tblHistory.DataBodyRange.rows(r).Interior.Color = RGB(255, 235, 238)
                Case "WARNING"
                    tblHistory.DataBodyRange.rows(r).Interior.Color = RGB(255, 248, 225)
                Case "INFO"
                    tblHistory.DataBodyRange.rows(r).Interior.Color = RGB(235, 242, 250)
                Case Else
                    tblHistory.DataBodyRange.rows(r).Interior.Pattern = xlNone
            End Select
        Next r
    End If

    If Not tblAck.DataBodyRange Is Nothing Then
        tblAck.ListColumns("Acknowledged At").DataBodyRange.NumberFormat = "dd/mm/yyyy hh:mm:ss"
    End If

    tblAlarm.Range.Columns.AutoFit
    tblHistory.Range.Columns.AutoFit
    tblAck.Range.Columns.AutoFit
    tblHistory.ListColumns("Message").Range.ColumnWidth = 72
    tblHistory.ListColumns("Message").Range.WrapText = True
    tblHistory.Range.VerticalAlignment = xlTop
    tblAck.Range.VerticalAlignment = xlTop
    tblHistory.Range.EntireRow.Hidden = False

    On Error GoTo 0

End Sub

Private Sub EnsureEventHistoryToggleShapes(ByVal ws As Worksheet)

    Dim leftPos As Double
    Dim topPos As Double
    Dim labelW As Double
    Dim trackW As Double
    Dim trackH As Double
    Dim knobSize As Double
    Dim gap As Double

    EnsureEventHistoryState

    leftPos = ws.Range("A1").Left + 4
    topPos = ws.Range("A1").Top + 4
    labelW = 42
    trackW = 36
    trackH = 14
    knobSize = 10
    gap = 18

    AddEventHistoryToggleLabel ws, EVENT_HISTORY_INFO_LABEL, "Info", leftPos, topPos, labelW, trackH, "Toggle_EventHistory_Info"
    AddEventHistoryToggleTrack ws, EVENT_HISTORY_INFO_BG, leftPos + labelW, topPos, trackW, trackH, "Toggle_EventHistory_Info"
    AddEventHistoryToggleKnob ws, EVENT_HISTORY_INFO_KNOB, leftPos + labelW, topPos + 2, knobSize, "Toggle_EventHistory_Info"

    leftPos = leftPos + labelW + trackW + gap
    AddEventHistoryToggleLabel ws, EVENT_HISTORY_LANG_LABEL, "FR / EN", leftPos, topPos, 56, trackH, "Toggle_EventHistory_Language"
    AddEventHistoryToggleTrack ws, EVENT_HISTORY_LANG_BG, leftPos + 56, topPos, trackW, trackH, "Toggle_EventHistory_Language"
    AddEventHistoryToggleKnob ws, EVENT_HISTORY_LANG_KNOB, leftPos + 56, topPos + 2, knobSize, "Toggle_EventHistory_Language"

End Sub

Private Sub EnsureEventHistoryCommandButtons( _
    ByVal wsHistory As Worksheet, _
    ByVal wsAck As Worksheet)

    AddEventHistoryCommandButton _
        wsHistory, EVENT_HISTORY_CLEAR_BUTTON, _
        "Nettoyer historique / Clear History", _
        wsHistory.Range("A2").Left, wsHistory.Range("A2").Top + 2, _
        wsHistory.Range("A2:C2").Width, 18, "ClearPlanningEventHistory"

    AddEventHistoryCommandButton _
        wsAck, EVENT_ACK_CLEAR_BUTTON, _
        "Nettoyer cache / Clear list", _
        wsAck.Range("A1").Left + 4, wsAck.Range("A1").Top + 2, _
        154, 18, "ClearPlanningWarningAcknowledgements"

End Sub

Private Sub AddEventHistoryCommandButton( _
    ByVal ws As Worksheet, _
    ByVal shapeName As String, _
    ByVal caption As String, _
    ByVal leftPos As Double, _
    ByVal topPos As Double, _
    ByVal widthVal As Double, _
    ByVal heightVal As Double, _
    ByVal macroName As String)

    Dim shp As Shape

    On Error Resume Next
    Set shp = ws.Shapes(shapeName)
    On Error GoTo 0

    If shp Is Nothing Then
        Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, leftPos, topPos, widthVal, heightVal)
        shp.Name = shapeName
    End If

    shp.Left = leftPos
    shp.Top = topPos
    shp.Width = widthVal
    shp.Height = heightVal
    shp.OnAction = macroName
    shp.Placement = xlMove
    shp.Adjustments.item(1) = 0.2
    shp.Fill.ForeColor.RGB = RGB(245, 245, 245)
    shp.Line.ForeColor.RGB = RGB(160, 160, 160)
    shp.Line.Weight = 0.75
    shp.Shadow.Visible = msoFalse
    shp.TextFrame2.TextRange.Text = caption
    shp.TextFrame2.VerticalAnchor = msoAnchorMiddle
    shp.TextFrame2.MarginLeft = 6
    shp.TextFrame2.MarginRight = 6
    shp.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    shp.TextFrame2.TextRange.Font.Name = "Segoe UI"
    shp.TextFrame2.TextRange.Font.Size = 8
    shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(55, 55, 55)

End Sub

Private Sub RefreshEventHistoryToggleVisuals(ByVal ws As Worksheet)

    Dim shpInfoBg As Shape
    Dim shpInfoKnob As Shape
    Dim shpInfoLabel As Shape
    Dim shpLangBg As Shape
    Dim shpLangKnob As Shape
    Dim shpLangLabel As Shape
    Dim langIsEN As Boolean

    EnsureEventHistoryState

    On Error Resume Next
    Set shpInfoBg = ws.Shapes(EVENT_HISTORY_INFO_BG)
    Set shpInfoKnob = ws.Shapes(EVENT_HISTORY_INFO_KNOB)
    Set shpInfoLabel = ws.Shapes(EVENT_HISTORY_INFO_LABEL)
    Set shpLangBg = ws.Shapes(EVENT_HISTORY_LANG_BG)
    Set shpLangKnob = ws.Shapes(EVENT_HISTORY_LANG_KNOB)
    Set shpLangLabel = ws.Shapes(EVENT_HISTORY_LANG_LABEL)
    On Error GoTo 0

    If Not shpInfoBg Is Nothing Then FormatEventHistoryToggleTrack shpInfoBg, gEventHistoryShowInfo
    If Not shpInfoKnob Is Nothing Then PositionEventHistoryToggleKnob shpInfoKnob, shpInfoBg, gEventHistoryShowInfo
    If Not shpInfoLabel Is Nothing Then FormatEventHistoryToggleLabel shpInfoLabel, gEventHistoryShowInfo

    langIsEN = (UCase$(Trim$(gPlanningLanguage)) = "EN")
    If Not shpLangBg Is Nothing Then FormatEventHistoryToggleTrack shpLangBg, langIsEN
    If Not shpLangKnob Is Nothing Then PositionEventHistoryToggleKnob shpLangKnob, shpLangBg, langIsEN
    If Not shpLangLabel Is Nothing Then FormatEventHistoryToggleLabel shpLangLabel, True

End Sub

Private Sub AddEventHistoryToggleLabel( _
    ByVal ws As Worksheet, _
    ByVal shapeName As String, _
    ByVal caption As String, _
    ByVal leftPos As Double, _
    ByVal topPos As Double, _
    ByVal widthVal As Double, _
    ByVal heightVal As Double, _
    ByVal macroName As String)

    Dim shp As Shape

    On Error Resume Next
    Set shp = ws.Shapes(shapeName)
    On Error GoTo 0

    If shp Is Nothing Then
        Set shp = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, leftPos, topPos, widthVal, heightVal)
        shp.Name = shapeName
    End If

    shp.Left = leftPos
    shp.Top = topPos
    shp.Width = widthVal
    shp.Height = heightVal
    shp.TextFrame2.TextRange.Text = caption
    shp.OnAction = macroName
    shp.Placement = xlMove
    FormatEventHistoryToggleLabel shp, True

End Sub

Private Sub AddEventHistoryToggleTrack( _
    ByVal ws As Worksheet, _
    ByVal shapeName As String, _
    ByVal leftPos As Double, _
    ByVal topPos As Double, _
    ByVal widthVal As Double, _
    ByVal heightVal As Double, _
    ByVal macroName As String)

    Dim shp As Shape

    On Error Resume Next
    Set shp = ws.Shapes(shapeName)
    On Error GoTo 0

    If shp Is Nothing Then
        Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, leftPos, topPos, widthVal, heightVal)
        shp.Name = shapeName
    End If

    shp.Left = leftPos
    shp.Top = topPos
    shp.Width = widthVal
    shp.Height = heightVal
    shp.OnAction = macroName
    shp.Placement = xlMove
    shp.Adjustments.item(1) = 0.5

End Sub

Private Sub AddEventHistoryToggleKnob( _
    ByVal ws As Worksheet, _
    ByVal shapeName As String, _
    ByVal leftPos As Double, _
    ByVal topPos As Double, _
    ByVal knobSize As Double, _
    ByVal macroName As String)

    Dim shp As Shape

    On Error Resume Next
    Set shp = ws.Shapes(shapeName)
    On Error GoTo 0

    If shp Is Nothing Then
        Set shp = ws.Shapes.AddShape(msoShapeOval, leftPos, topPos, knobSize, knobSize)
        shp.Name = shapeName
        shp.Left = leftPos
        shp.Top = topPos
    End If

    shp.Width = knobSize
    shp.Height = knobSize
    shp.OnAction = macroName
    shp.Placement = xlMove
    shp.Fill.ForeColor.RGB = RGB(255, 255, 255)
    shp.Line.ForeColor.RGB = RGB(150, 150, 150)
    shp.Line.Weight = 0.75
    shp.Shadow.Visible = msoFalse

End Sub

Private Sub FormatEventHistoryToggleTrack(ByVal shp As Shape, ByVal isOn As Boolean)

    shp.Line.Weight = 1
    shp.Shadow.Visible = msoFalse

    If isOn Then
        shp.Fill.ForeColor.RGB = RGB(68, 114, 196)
        shp.Line.ForeColor.RGB = RGB(68, 114, 196)
    Else
        shp.Fill.ForeColor.RGB = RGB(230, 230, 230)
        shp.Line.ForeColor.RGB = RGB(170, 170, 170)
    End If

End Sub

Private Sub FormatEventHistoryToggleLabel(ByVal shp As Shape, ByVal isOn As Boolean)

    shp.Line.Visible = msoFalse
    shp.Fill.Visible = msoFalse
    shp.TextFrame2.VerticalAnchor = msoAnchorMiddle
    shp.TextFrame2.MarginLeft = 0
    shp.TextFrame2.MarginRight = 0
    shp.TextFrame2.TextRange.Font.Name = "Segoe UI"
    shp.TextFrame2.TextRange.Font.Size = 9
    shp.TextFrame2.TextRange.Font.Bold = msoTrue

    If isOn Then
        shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(68, 114, 196)
    Else
        shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(90, 90, 90)
    End If

End Sub

Private Sub PositionEventHistoryToggleKnob(ByVal knob As Shape, ByVal track As Shape, ByVal isOn As Boolean)

    If track Is Nothing Then Exit Sub

    knob.Top = track.Top + ((track.Height - knob.Height) / 2)
    If isOn Then
        knob.Left = track.Left + track.Width - knob.Width - 2
    Else
        knob.Left = track.Left + 2
    End If

End Sub

Private Sub DeleteEventHistoryShapeIfExists(ByVal ws As Worksheet, ByVal shapeName As String)

    On Error Resume Next
    ws.Shapes(shapeName).Delete
    On Error GoTo 0

End Sub

Private Sub ClearPlanningEventTableRows(ByVal tbl As ListObject)

    If tbl Is Nothing Then Exit Sub
    If tbl.DataBodyRange Is Nothing Then Exit Sub

    tbl.DataBodyRange.Delete

End Sub

Public Function IsPlanningEventInternalWriteActive() As Boolean

    IsPlanningEventInternalWriteActive = (gEventHistoryInternalWriteDepth > 0)

End Function

Private Sub BeginPlanningEventInternalWrite( _
    ByRef wsAlarm As Worksheet, _
    ByRef wsHistory As Worksheet, _
    ByRef wsAck As Worksheet)

    gEventHistoryInternalWriteDepth = gEventHistoryInternalWriteDepth + 1

    If Not IsPlanningEventInfrastructureReady() Then
        EnsurePlanningEventHistoryInfrastructure
    End If

    Set wsAlarm = ThisWorkbook.Worksheets(CALC_ALARM_SHEET)
    Set wsHistory = ThisWorkbook.Worksheets(EVENT_HISTORY_SHEET)
    Set wsAck = ThisWorkbook.Worksheets(EVENT_ACK_SHEET)

End Sub

Private Sub EndPlanningEventInternalWrite( _
    ByVal wsAlarm As Worksheet, _
    ByVal wsHistory As Worksheet, _
    ByVal wsAck As Worksheet)

    If gEventHistoryInternalWriteDepth <= 0 Then
        gEventHistoryInternalWriteDepth = 0
        Exit Sub
    End If

    gEventHistoryInternalWriteDepth = gEventHistoryInternalWriteDepth - 1

End Sub

Private Function IsPlanningEventInfrastructureReady() As Boolean

    Dim wsAlarm As Worksheet
    Dim wsHistory As Worksheet
    Dim wsAck As Worksheet
    Dim tblAlarm As ListObject
    Dim tblHistory As ListObject
    Dim tblAck As ListObject

    On Error GoTo SafeExit

    Set wsAlarm = ThisWorkbook.Worksheets(CALC_ALARM_SHEET)
    Set wsHistory = ThisWorkbook.Worksheets(EVENT_HISTORY_SHEET)
    Set wsAck = ThisWorkbook.Worksheets(EVENT_ACK_SHEET)
    Set tblAlarm = wsAlarm.ListObjects(CALC_ALARM_TABLE)
    Set tblHistory = wsHistory.ListObjects(EVENT_HISTORY_TABLE)
    Set tblAck = wsAck.ListObjects(EVENT_ACK_TABLE)

    IsPlanningEventInfrastructureReady = _
        TableHasColumn(tblAlarm, "Event ID") And _
        TableHasColumn(tblAlarm, "Source Procedure") And _
        TableHasColumn(tblHistory, "Acknowledged") And _
        TableHasColumn(tblAck, "Hash") And _
        TableHasColumn(tblAck, "EN Message")

SafeExit:
End Function

Private Function CalcAlarmHeaders() As Variant

    CalcAlarmHeaders = Array( _
        "Event ID", _
        "Run ID", _
        "Timestamp", _
        "Date", _
        "Time", _
        "Severity", _
        "Event Type", _
        "Hash", _
        "Acknowledged", _
        "Sheet", _
        "Table", _
        "ID", _
        "WBS", _
        "Task Name", _
        "FR Message", _
        "EN Message", _
        "FR Details", _
        "EN Details", _
        "Source Procedure")

End Function

Private Function EventHistoryHeaders() As Variant

    EventHistoryHeaders = Array( _
        "Date", _
        "Heure", _
        "Severity", _
        "Message", _
        "Acknowledged")

End Function

Private Function EventAckHeaders() As Variant

    EventAckHeaders = Array( _
        "Hash", _
        "Severity", _
        "Event Type", _
        "Acknowledged", _
        "Acknowledged At", _
        "Acknowledged By", _
        "Comment", _
        "WBS", _
        "Task Name", _
        "FR Message", _
        "EN Message")

End Function









