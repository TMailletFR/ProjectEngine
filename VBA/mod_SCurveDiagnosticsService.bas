Attribute VB_Name = "mod_SCurveDiagnosticsService"
Option Explicit

'===============================================================================
' MODULE : mod_SCurveDiagnosticsService
' DOMAINE / DOMAIN : S-Curve
'
' FR
' Possede le workflow specialise indique par son nom et expose ses contrats stables.
' Ne possede pas les domaines appeles en dependance.
'
' EN
' Owns the named specialized workflow and exposes its stable contracts.
' Does not own the domains it calls as dependencies.
'
' CONTRATS / CONTRACTS : SCurve_AddConsoleMessage, SCurve_AddGroupedMessage, SCurve_LogGroupedWarningEvents
' CALLBACKS EXTERNES / EXTERNAL CALLBACKS : Aucun / None
'===============================================================================


'------------------------------------------------------------------------------
' FR: Ajoute un message console S-Curve simple sans modifier la logique de calcul.
' EN: Adds a simple S-Curve console message without changing calculation logic.
'------------------------------------------------------------------------------
Public Sub SCurve_AddConsoleMessage( _
    ByVal consoleMessages As Collection, _
    ByVal msgType As String, _
    ByVal messageKey As String, _
    Optional ByVal namedArguments As Object = Nothing, _
    Optional ByVal eventType As String = "", _
    Optional ByVal eventHash As String = "")

    If consoleMessages Is Nothing Then Exit Sub

    CalcBridge_AddConsoleMessage consoleMessages, msgType, _
        PlanningMessageText_Format(messageKey, namedArguments, namedArguments), _
        False, _
        eventType, _
        eventHash

End Sub

'------------------------------------------------------------------------------
' FR: Ajoute un diagnostic S-Curve groupe avec contexte ID/WBS et ACK eventuel.
' EN: Adds a grouped S-Curve diagnostic with ID/WBS context and optional ACK.
'------------------------------------------------------------------------------
Public Sub SCurve_AddGroupedMessage( _
    ByVal consoleMessages As Collection, _
    ByVal msgType As String, _
    ByVal idsDict As Object, _
    ByVal idToWbs As Object, _
    ByVal messageKey As String, _
    Optional ByVal historyHandled As Boolean = False, _
    Optional ByVal ackTokens As String = "", _
    Optional ByVal memberReceipts As Object = Nothing)

    If consoleMessages Is Nothing Then Exit Sub
    If idsDict Is Nothing Then Exit Sub
    If idsDict.Count = 0 Then Exit Sub

    CalcBridge_AddGroupedConsoleMessage consoleMessages, msgType, idsDict, idToWbs, _
        messageKey, ackTokens, memberReceipts

End Sub

'------------------------------------------------------------------------------
' FR: Journalise les warnings S-Curve groupes et retourne les tokens ACK.
' EN: Logs grouped S-Curve warnings and returns ACK tokens.
'------------------------------------------------------------------------------
Public Function SCurve_LogGroupedWarningEvents( _
    ByVal idsDict As Object, _
    ByVal idToWbs As Object, _
    ByVal eventType As String, _
    ByVal messageKey As String, _
    ByVal detailsKey As String, _
    Optional ByVal receiptsById As Object = Nothing) As String

    Dim key As Variant
    Dim idVal As String
    Dim wbsVal As String
    Dim eventHash As String
    Dim tokens As String
    Dim frMessage As String
    Dim enMessage As String
    Dim frDetails As String
    Dim enDetails As String
    Dim receipt As Object

    If idsDict Is Nothing Then Exit Function
    frMessage = TextCatalog_Get(messageKey, TEXT_LANGUAGE_FR)
    enMessage = TextCatalog_Get(messageKey, TEXT_LANGUAGE_EN)
    frDetails = TextCatalog_Get(detailsKey, TEXT_LANGUAGE_FR)
    enDetails = TextCatalog_Get(detailsKey, TEXT_LANGUAGE_EN)

    For Each key In idsDict.Keys
        idVal = Trim$(CStr(key))
        If idVal <> "" Then
            wbsVal = vbNullString
            If Not idToWbs Is Nothing Then
                If idToWbs.Exists(idVal) Then wbsVal = CStr(idToWbs(idVal))
            End If

            eventHash = BuildPlanningEventIdentityV2( _
                "WARNING", eventType, "TASK", idVal)

            Set receipt = LogPlanningEvent( _
                "WARNING", _
                eventType, _
                eventHash, _
                frMessage, _
                enMessage, _
                frDetails, _
                enDetails, _
                "Run_SCurve_Engine", _
                "SCURVE", _
                "tbl_SCURVE", _
                idVal, _
                wbsVal, _
                vbNullString, _
                False)
            If Not receiptsById Is Nothing Then Set receiptsById(idVal) = receipt

            If tokens <> "" Then tokens = tokens & ";"
            tokens = tokens & BuildPlanningWarningAckToken(eventType, eventHash)
        End If
    Next key

    SCurve_LogGroupedWarningEvents = tokens

End Function

'------------------------------------------------------------------------------
' FR: Construit la collection S Curve Diagnostics Build Inline List a partir des donnees fournies par l'appelant.
' EN: Builds the S Curve Diagnostics Build Inline List collection from data supplied by the caller.
'------------------------------------------------------------------------------

Private Function SCurveDiagnostics_BuildInlineList(ByVal idsDict As Object, ByVal maxItems As Long) As String

    Dim result As String
    Dim key As Variant
    Dim countShown As Long
    Dim totalCount As Long

    result = ""
    countShown = 0
    totalCount = idsDict.Count

    For Each key In idsDict.Keys
        countShown = countShown + 1
        If countShown <= maxItems Then
            If result <> "" Then result = result & " / "
            result = result & CStr(key)
        Else
            Exit For
        End If
    Next key

    If totalCount > maxItems Then
        result = result & " / +" & CStr(totalCount - maxItems)
    End If

    SCurveDiagnostics_BuildInlineList = result

End Function

'------------------------------------------------------------------------------
' FR: Construit la collection S Curve Diagnostics Build Inline WBS List a partir des donnees fournies par l'appelant.
' EN: Builds the S Curve Diagnostics Build Inline WBS List collection from data supplied by the caller.
'------------------------------------------------------------------------------

Private Function SCurveDiagnostics_BuildInlineWBSList(ByVal idsDict As Object, ByVal idToWbs As Object, ByVal maxItems As Long) As String

    Dim result As String
    Dim key As Variant
    Dim countShown As Long
    Dim totalCount As Long
    Dim itemText As String

    result = ""
    countShown = 0
    totalCount = idsDict.Count

    For Each key In idsDict.Keys
        countShown = countShown + 1
        If countShown <= maxItems Then
            If Not idToWbs Is Nothing Then
                If idToWbs.Exists(CStr(key)) Then
                    itemText = CStr(idToWbs(CStr(key)))
                Else
                    itemText = "-"
                End If
            Else
                itemText = "-"
            End If

            If result <> "" Then result = result & " / "
            result = result & itemText
        Else
            Exit For
        End If
    Next key

    If totalCount > maxItems Then
        result = result & " / +" & CStr(totalCount - maxItems)
    End If

    SCurveDiagnostics_BuildInlineWBSList = result

End Function
