Attribute VB_Name = "mod_GanttSimulationMessages"
Option Explicit

'===============================================================================
' MODULE : mod_GanttSimulationMessages
' DOMAINE / DOMAIN : Gantt
'
' FR
' Transforme les resultats TEST/SCENARIO/LOCK en messages bilingues sans recalcul metier.
' Ne doit pas contourner les contrats publics des autres domaines.
'
' EN
' Transforms TEST/SCENARIO/LOCK results into bilingual messages without business recalculation.
' Must not bypass public contracts owned by other domains.
'
' CONTRATS / CONTRACTS : ShowGanttLiveGroupedMessage, GanttLive_IsInheritedCoreError, GanttLive_RemoveDerivedLOERootErrors, GanttLive_CalcGanttTestHasErrors, GanttLive_HasConsoleCollection, GanttLive_AddVbaOrStructuredError, GanttLive_AddBiConsoleMessage
' CALLBACKS EXTERNES / EXTERNAL CALLBACKS : Aucun / None
'===============================================================================



'------------------------------------------------------------------------------
' FR: Publie un message groupe live dans la console existante ou dans une console locale.
' EN: Publishes a grouped live message to the existing console or to a local console.
'------------------------------------------------------------------------------
Public Sub ShowGanttLiveGroupedMessage( _
    ByVal idsDict As Object, _
    ByVal idToWbs As Object, _
    ByVal messageKey As String, _
    ByVal boxStyle As VbMsgBoxStyle, _
    Optional ByVal consoleMessages As Variant)

    Dim msgType As String
    Dim localMessages As Collection

    If idsDict Is Nothing Then Exit Sub
    If idsDict.Count = 0 Then Exit Sub

    msgType = GanttLive_MessageTypeFromMsgBoxStyle(boxStyle)

    If GanttLive_HasConsoleCollection(consoleMessages) Then
        CalcBridge_AddConsoleMessage consoleMessages, msgType, _
            CalcBridge_BuildGroupedMessage(idsDict, idToWbs, messageKey)
    Else
        Set localMessages = New Collection
        CalcBridge_AddConsoleMessage localMessages, msgType, _
            CalcBridge_BuildGroupedMessage(idsDict, idToWbs, messageKey)
        CalcBridge_ShowPlanningConsole localMessages
    End If

End Sub

'------------------------------------------------------------------------------
' FR: Formate une liste compacte d'IDs avec limite d'affichage.
' EN: Formats a compact capped ID list.
'------------------------------------------------------------------------------
Private Function BuildInlineList_GanttLive(ByVal idsDict As Object, ByVal maxItems As Long) As String

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

    BuildInlineList_GanttLive = result

End Function

'------------------------------------------------------------------------------
' FR: Formate une liste compacte de WBS correspondant a une liste d'IDs.
' EN: Formats a compact WBS list corresponding to an ID list.
'------------------------------------------------------------------------------
Private Function BuildInlineWBSList_GanttLive(ByVal idsDict As Object, ByVal idToWbs As Object, ByVal maxItems As Long) As String

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
            If idToWbs.Exists(CStr(key)) Then
                itemText = NormalizeWBS(CStr(idToWbs(CStr(key))))
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

    BuildInlineWBSList_GanttLive = result

End Function


'------------------------------------------------------------------------------
' FR: Retire des causes racines les erreurs LOE qui ne font que refleter un predecesseur deja en erreur.
' EN: Removes LOE errors from root causes when they only reflect an already failing predecessor.
'------------------------------------------------------------------------------
Public Sub GanttLive_RemoveDerivedLOERootErrors( _
    ByVal coreDiagnostics As Object, _
    ByVal errorIds As Object, _
    ByVal rootErrorIds As Object)

    Dim removeIds As Object
    Dim oneId As Variant

    If errorIds Is Nothing Then Exit Sub
    If rootErrorIds Is Nothing Then Exit Sub
    If rootErrorIds.Count = 0 Then Exit Sub

    Set removeIds = CreateObject("Scripting.Dictionary")

    For Each oneId In rootErrorIds.Keys
        If Not CoreDiagnostics_TaskHasClassification(coreDiagnostics, CStr(oneId), "ROOT") Then
            removeIds(CStr(oneId)) = True
        End If
    Next oneId

    For Each oneId In removeIds.Keys
        If rootErrorIds.Exists(CStr(oneId)) Then rootErrorIds.Remove CStr(oneId)
    Next oneId

End Sub

'------------------------------------------------------------------------------
' FR: Detecte les erreurs presentes dans tbl_CALC_GANTT_TEST avant un lock.
' EN: Detects errors present in tbl_CALC_GANTT_TEST before lock.
'------------------------------------------------------------------------------
Public Function GanttLive_CalcGanttTestHasErrors(ByVal tblTest As ListObject) As Boolean

    Dim mapTest As Object
    Dim arr As Variant
    Dim r As Long

    On Error GoTo SafeExit

    GanttLive_CalcGanttTestHasErrors = False

    If tblTest Is Nothing Then Exit Function
    If tblTest.DataBodyRange Is Nothing Then Exit Function

    Set mapTest = CanonicalIdentity_BuildColumnMap(tblTest)

    If Not mapTest.Exists("Error Flag") Then Exit Function

    arr = tblTest.DataBodyRange.value

    For r = 1 To UBound(arr, 1)
        If UCase$(Trim$(CStr(arr(r, mapTest("Error Flag"))))) = "ERROR" Then
            GanttLive_CalcGanttTestHasErrors = True
            Exit Function
        End If
    Next r

    Exit Function

SafeExit:
    GanttLive_CalcGanttTestHasErrors = True

End Function

'------------------------------------------------------------------------------
' FR: Verifie si un parametre optionnel contient une collection console utilisable.
' EN: Checks whether an optional argument contains a usable console collection.
'------------------------------------------------------------------------------
Public Function GanttLive_HasConsoleCollection(Optional ByVal consoleMessages As Variant) As Boolean

    On Error GoTo SafeExit

    If IsMissing(consoleMessages) Then Exit Function
    If IsObject(consoleMessages) Then
        If Not consoleMessages Is Nothing Then
            GanttLive_HasConsoleCollection = True
        End If
    End If

SafeExit:
End Function

'------------------------------------------------------------------------------
' FR: Convertit un style MsgBox en type de message console STOP/WARNING/INFO.
' EN: Converts a MsgBox style into a STOP/WARNING/INFO console message type.
'------------------------------------------------------------------------------
Private Function GanttLive_MessageTypeFromMsgBoxStyle(ByVal boxStyle As VbMsgBoxStyle) As String

    If (boxStyle And vbCritical) = vbCritical Then
        GanttLive_MessageTypeFromMsgBoxStyle = "STOP"
    ElseIf (boxStyle And vbExclamation) = vbExclamation Then
        GanttLive_MessageTypeFromMsgBoxStyle = "WARNING"
    Else
        GanttLive_MessageTypeFromMsgBoxStyle = "INFO"
    End If

End Function

'------------------------------------------------------------------------------
' FR: Detecte un message deja structure en blocs FR/EN.
' EN: Detects a message already structured as FR/EN blocks.
'------------------------------------------------------------------------------
Private Function GanttLive_IsStructuredBiMessage(ByVal msgText As String) As Boolean

    Dim txt As String

    txt = LTrim$(CStr(msgText))

    GanttLive_IsStructuredBiMessage = _
        (Left$(txt, 3) = "FR:" And InStr(1, txt, "EN:", vbTextCompare) > 0)

End Function

'------------------------------------------------------------------------------
' FR: Ajoute a la console une erreur VBA, en preservant les messages FR/EN deja structures.
' EN: Adds a VBA error to the console while preserving already structured FR/EN messages.
'------------------------------------------------------------------------------
Public Sub GanttLive_AddVbaOrStructuredError( _
    ByVal consoleMessages As Collection, _
    ByVal functionName As String, _
    ByVal errDescription As String)

    If consoleMessages Is Nothing Then Exit Sub

    If GanttLive_IsStructuredBiMessage(errDescription) Then
        CalcBridge_AddConsoleMessage consoleMessages, "STOP", Trim$(CStr(errDescription))
    Else
        GanttLive_AddBiConsoleMessage consoleMessages, "STOP", _
            "GANTT.SIMULATION.VBA_ERROR", _
            TextCatalog_Arguments("Function", functionName, "Details", errDescription)
    End If

End Sub

'------------------------------------------------------------------------------
' FR: Ajoute un message console bilingue FR/EN avec le type fourni.
' EN: Adds a bilingual FR/EN console message with the supplied type.
'------------------------------------------------------------------------------
Public Sub GanttLive_AddBiConsoleMessage( _
    ByVal consoleMessages As Collection, _
    ByVal msgType As String, _
    ByVal messageKey As String, _
    Optional ByVal namedArguments As Object = Nothing)

    Dim msg As String

    If consoleMessages Is Nothing Then Exit Sub

    msg = PlanningMessageText_Format(messageKey, namedArguments, namedArguments)

    CalcBridge_AddConsoleMessage consoleMessages, msgType, msg

End Sub
