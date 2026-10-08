Attribute VB_Name = "mod_CoreBridgeDiagFormatting"
Option Explicit

'===============================================================================
' MODULE : mod_CoreBridgeDiagFormatting
' DOMAINE / DOMAIN : Core Bridge
'
' FR
' Construit les textes bilingues des diagnostics CoreBridge sans lire Excel ni recalculer le planning.
' Ne doit pas contourner les contrats publics des autres domaines.
'
' EN
' Builds bilingual CoreBridge diagnostic text without reading Excel or recalculating planning.
' Must not bypass public contracts owned by other domains.
'
' CONTRATS / CONTRACTS : CalcBridge_BuildGroupedMessage, CalcBridge_BuildInlineList, CalcBridge_BuildInlineWBSList, CalcBridge_BuildUpstreamViolationItems, CalcBridge_CycleDetailMessageFromCoreError, CalcBridge_ToPMConstraintMessage, CalcBridge_BuildMissingBaselineRexMessage, CalcBridge_BuildLOEExplicitPredecessorMessage
' CALLBACKS EXTERNES / EXTERNAL CALLBACKS : Aucun / None
'===============================================================================


'------------------------------------------------------------------------------
' FR: Construit un message bilingue groupe pour des IDs/WBS.
' EN: Builds a bilingual grouped message for IDs/WBS values.
'------------------------------------------------------------------------------
Public Function CalcBridge_BuildGroupedMessage( _
    ByVal idsDict As Object, _
    ByVal idToWbs As Object, _
    ByVal messageKey As String) As String

    Dim idsLine As String
    Dim wbsLine As String
    Dim arguments As Object

    idsLine = CalcBridge_BuildInlineList(idsDict, 20)
    wbsLine = CalcBridge_BuildInlineWBSList(idsDict, idToWbs, 20)

    Set arguments = TextCatalog_Arguments("Ids", idsLine, "Wbs", wbsLine)
    CalcBridge_BuildGroupedMessage = PlanningMessageText_Format(messageKey, arguments, arguments)

End Function

'------------------------------------------------------------------------------
' FR: Formate une liste compacte d'IDs pour un diagnostic groupe.
' EN: Formats a compact ID list for a grouped diagnostic.
'------------------------------------------------------------------------------
Public Function CalcBridge_BuildInlineList( _
    ByVal idsDict As Object, _
    ByVal maxItems As Long) As String

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

    CalcBridge_BuildInlineList = result

End Function

'------------------------------------------------------------------------------
' FR: Formate une liste compacte de WBS correspondant a des IDs.
' EN: Formats a compact WBS list matching IDs.
'------------------------------------------------------------------------------
Public Function CalcBridge_BuildInlineWBSList( _
    ByVal idsDict As Object, _
    ByVal idToWbs As Object, _
    ByVal maxItems As Long) As String

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

    CalcBridge_BuildInlineWBSList = result

End Function

'------------------------------------------------------------------------------
' FR: Formate les couples ID/WBS des violations amont.
' EN: Formats ID/WBS pairs for upstream violations.
'------------------------------------------------------------------------------
Public Function CalcBridge_BuildUpstreamViolationItems( _
    ByVal idsDict As Object, _
    ByVal idToWbs As Object, _
    ByVal maxItems As Long) As String

    Dim result As String
    Dim key As Variant
    Dim countShown As Long
    Dim totalCount As Long
    Dim wbsVal As String

    result = ""
    countShown = 0
    totalCount = idsDict.Count

    For Each key In idsDict.Keys

        countShown = countShown + 1

        If countShown <= maxItems Then

            If Not idToWbs Is Nothing Then
                If idToWbs.Exists(CStr(key)) Then
                    wbsVal = CStr(idToWbs(CStr(key)))
                Else
                    wbsVal = "-"
                End If
            Else
                wbsVal = "-"
            End If

            If result <> "" Then result = result & " / "
            result = result & CStr(key) & " (" & wbsVal & ")"

        Else
            Exit For
        End If

    Next key

    If totalCount > maxItems Then
        result = result & " / +" & CStr(totalCount - maxItems)
    End If

    CalcBridge_BuildUpstreamViolationItems = result

End Function

'------------------------------------------------------------------------------
' FR: Lit une valeur texte dans un dictionnaire diagnostic.
' EN: Reads a text value from a diagnostic dictionary.
'------------------------------------------------------------------------------
Private Function CalcBridge_DiagString(ByVal diag As Object, ByVal key As String) As String

    If diag Is Nothing Then Exit Function
    If diag.Exists(key) Then CalcBridge_DiagString = Trim$(CStr(diag(key)))

End Function

'------------------------------------------------------------------------------
' FR: Lit une valeur brute dans un dictionnaire diagnostic.
' EN: Reads a raw value from a diagnostic dictionary.
'------------------------------------------------------------------------------
Private Function CalcBridge_DiagValue(ByVal diag As Object, ByVal key As String) As Variant

    If diag Is Nothing Then Exit Function
    If diag.Exists(key) Then CalcBridge_DiagValue = diag(key)

End Function

'------------------------------------------------------------------------------
' FR: Formate une valeur diagnostic quelconque en texte affichable.
' EN: Formats any diagnostic value as display text.
'------------------------------------------------------------------------------
Private Function CalcBridge_DiagnosticAnyValueText(ByVal value As Variant) As String

    If Not HasValue(value) Then
        CalcBridge_DiagnosticAnyValueText = "-"
    ElseIf IsDate(value) Then
        CalcBridge_DiagnosticAnyValueText = Format$(CDate(value), "dd/mm/yyyy")
    Else
        CalcBridge_DiagnosticAnyValueText = CStr(value)
    End If

End Function

'------------------------------------------------------------------------------
' FR: Formate le libelle lisible d'une tache dans un diagnostic.
' EN: Formats the readable task label used in diagnostics.
'------------------------------------------------------------------------------
Private Function CalcBridge_DiagnosticTaskLabel( _
    ByVal taskId As String, _
    ByVal idToWbs As Object, _
    ByVal idToTaskName As Object) As String

    Dim wbsVal As String
    Dim nameVal As String

    If Not idToWbs Is Nothing Then
        If idToWbs.Exists(taskId) Then wbsVal = Trim$(CStr(idToWbs(taskId)))
    End If

    If Not idToTaskName Is Nothing Then
        If idToTaskName.Exists(taskId) Then nameVal = Trim$(CStr(idToTaskName(taskId)))
    End If

    If wbsVal <> "" And nameVal <> "" Then
        CalcBridge_DiagnosticTaskLabel = wbsVal & " " & nameVal
    ElseIf wbsVal <> "" Then
        CalcBridge_DiagnosticTaskLabel = wbsVal
    ElseIf nameVal <> "" Then
        CalcBridge_DiagnosticTaskLabel = nameVal
    Else
        CalcBridge_DiagnosticTaskLabel = TextCatalog_Format( _
            "DIAG.COMMON.ID_TASK_LABEL", TEXT_LANGUAGE_EN, _
            TextCatalog_Arguments("Id", taskId))
    End If

End Function

'------------------------------------------------------------------------------
' FR: Formate une date de diagnostic ou un tiret si absente.
' EN: Formats a diagnostic date or a dash when absent.
'------------------------------------------------------------------------------
Private Function CalcBridge_DiagnosticDateText(ByVal dateValue As Variant) As String

    If HasValue(dateValue) Then
        CalcBridge_DiagnosticDateText = Format$(CDate(dateValue), "dd/mm/yyyy")
    Else
        CalcBridge_DiagnosticDateText = "-"
    End If

End Function

'------------------------------------------------------------------------------
' FR: Formate un lag diagnostic avec signe explicite.
' EN: Formats a diagnostic lag with an explicit sign.
'------------------------------------------------------------------------------
Private Function CalcBridge_FormatDiagnosticLag(ByVal lagValue As Double) As String

    Dim lagText As String

    lagText = Replace$(Format$(Abs(lagValue), "0.##"), ",", ".")

    If lagValue >= 0# Then
        CalcBridge_FormatDiagnosticLag = "+" & lagText
    Else
        CalcBridge_FormatDiagnosticLag = "-" & lagText
    End If

End Function


'------------------------------------------------------------------------------
' FR: Retourne la valeur Cycle Detail Message From Core Error sans modifier les donnees d'entree.
' EN: Returns the Cycle Detail Message From Core Error value without mutating input data.
'------------------------------------------------------------------------------

Public Function CalcBridge_CycleDetailMessageFromCoreError(ByVal errMsg As String) As String

    Dim markerPos As Long
    Dim msg As String

    msg = CStr(errMsg)
    markerPos = InStr(1, msg, "FR:", vbTextCompare)

    If markerPos > 0 Then
        CalcBridge_CycleDetailMessageFromCoreError = Trim$(Mid$(msg, markerPos))
    Else
        CalcBridge_CycleDetailMessageFromCoreError = ""
    End If

End Function


'------------------------------------------------------------------------------
' FR: Retourne la valeur To PM Constraint Message sans modifier les donnees d'entree.
' EN: Returns the To PM Constraint Message value without mutating input data.
'------------------------------------------------------------------------------

Public Function CalcBridge_ToPMConstraintMessage(ByVal rawMessage As String) As String

    Dim msg As String

    msg = CStr(rawMessage)

    CalcBridge_ReplacePMMessageKey msg, "DIAG.CONSTRAINT.UNKNOWN_START_TYPE", "DIAG.CONSTRAINT.PM.UNKNOWN_START_TYPE"
    CalcBridge_ReplacePMMessageKey msg, "DIAG.CONSTRAINT.UNKNOWN_FINISH_TYPE", "DIAG.CONSTRAINT.PM.UNKNOWN_FINISH_TYPE"
    CalcBridge_ReplacePMMessageKey msg, "DIAG.CONSTRAINT.ACTUAL_START_BEFORE_START", "DIAG.CONSTRAINT.PM.ACTUAL_START_BEFORE_START"
    CalcBridge_ReplacePMMessageKey msg, "DIAG.CONSTRAINT.FORECAST_START_BEFORE_START", "DIAG.CONSTRAINT.PM.FORECAST_START_BEFORE_START"
    CalcBridge_ReplacePMMessageKey msg, "DIAG.CONSTRAINT.ACTUAL_START_AFTER_LATEST", "DIAG.CONSTRAINT.PM.ACTUAL_START_AFTER_LATEST"
    CalcBridge_ReplacePMMessageKey msg, "DIAG.CONSTRAINT.FORECAST_START_AFTER_LATEST", "DIAG.CONSTRAINT.PM.FORECAST_START_AFTER_LATEST"
    CalcBridge_ReplacePMMessageKey msg, "DIAG.CONSTRAINT.CALCULATED_START_AFTER_LATEST", "DIAG.CONSTRAINT.PM.CALCULATED_START_AFTER_LATEST"
    CalcBridge_ReplacePMMessageKey msg, "DIAG.CONSTRAINT.ACTUAL_START_DIFFERS_MSO", "DIAG.CONSTRAINT.PM.ACTUAL_START_DIFFERS_MSO"
    CalcBridge_ReplacePMMessageKey msg, "DIAG.CONSTRAINT.FORECAST_START_DIFFERS_MSO", "DIAG.CONSTRAINT.PM.FORECAST_START_DIFFERS_MSO"
    CalcBridge_ReplacePMMessageKey msg, "DIAG.CONSTRAINT.CALCULATED_START_DIFFERS_MSO", "DIAG.CONSTRAINT.PM.CALCULATED_START_DIFFERS_MSO"
    CalcBridge_ReplacePMMessageKey msg, "DIAG.CONSTRAINT.ACTUAL_FINISH_BEFORE_FINISH", "DIAG.CONSTRAINT.PM.ACTUAL_FINISH_BEFORE_FINISH"
    CalcBridge_ReplacePMMessageKey msg, "DIAG.CONSTRAINT.FORECAST_FINISH_BEFORE_UPSTREAM", "DIAG.CONSTRAINT.PM.FORECAST_FINISH_BEFORE_UPSTREAM"
    CalcBridge_ReplacePMMessageKey msg, "DIAG.CONSTRAINT.FORECAST_FINISH_BEFORE_FINISH", "DIAG.CONSTRAINT.PM.FORECAST_FINISH_BEFORE_FINISH"
    CalcBridge_ReplacePMMessageKey msg, "DIAG.CONSTRAINT.ACTUAL_FINISH_AFTER_LATEST", "DIAG.CONSTRAINT.PM.ACTUAL_FINISH_AFTER_LATEST"
    CalcBridge_ReplacePMMessageKey msg, "DIAG.CONSTRAINT.FORECAST_FINISH_AFTER_LATEST", "DIAG.CONSTRAINT.PM.FORECAST_FINISH_AFTER_LATEST"
    CalcBridge_ReplacePMMessageKey msg, "DIAG.CONSTRAINT.CALCULATED_FINISH_AFTER_LATEST", "DIAG.CONSTRAINT.PM.CALCULATED_FINISH_AFTER_LATEST"
    CalcBridge_ReplacePMMessageKey msg, "DIAG.CONSTRAINT.ACTUAL_FINISH_DIFFERS_MFO", "DIAG.CONSTRAINT.PM.ACTUAL_FINISH_DIFFERS_MFO"
    CalcBridge_ReplacePMMessageKey msg, "DIAG.CONSTRAINT.FORECAST_FINISH_DIFFERS_MFO", "DIAG.CONSTRAINT.PM.FORECAST_FINISH_DIFFERS_MFO"
    CalcBridge_ReplacePMMessageKey msg, "DIAG.CONSTRAINT.CALCULATED_FINISH_DIFFERS_MFO", "DIAG.CONSTRAINT.PM.CALCULATED_FINISH_DIFFERS_MFO"
    CalcBridge_ReplacePMMessageKey msg, "DIAG.CONSTRAINT.CALCULATED_START_BEFORE_MFO_NETWORK", "DIAG.CONSTRAINT.PM.CALCULATED_START_BEFORE_MFO_NETWORK"
    CalcBridge_ReplacePMMessageKey msg, "DIAG.CONSTRAINT.ACTUAL_START_DIFFERS_MFO_IMPLIED", "DIAG.CONSTRAINT.PM.ACTUAL_START_DIFFERS_MFO_IMPLIED"
    CalcBridge_ReplacePMMessageKey msg, "DIAG.CONSTRAINT.FORECAST_START_DIFFERS_MFO_IMPLIED", "DIAG.CONSTRAINT.PM.FORECAST_START_DIFFERS_MFO_IMPLIED"
    CalcBridge_ReplacePMMessageKey msg, "DIAG.CONSTRAINT.DURATION_INCOMPATIBLE_MSO_MFO", "DIAG.CONSTRAINT.PM.DURATION_INCOMPATIBLE_MSO_MFO"

    CalcBridge_ToPMConstraintMessage = msg

End Function


'------------------------------------------------------------------------------
' FR: Transforme la valeur Replace PM Message sans modifier la semantique du message source.
' EN: Transforms the Replace PM Message value without changing source-message semantics.
'------------------------------------------------------------------------------

Private Sub CalcBridge_ReplacePMMessageKey( _
    ByRef msg As String, _
    ByVal sourceKey As String, _
    ByVal replacementKey As String)

    msg = Replace(msg, TextCatalog_Get(sourceKey, TEXT_LANGUAGE_FR), _
        TextCatalog_Get(replacementKey, TEXT_LANGUAGE_FR), 1, -1, vbTextCompare)
    msg = Replace(msg, TextCatalog_Get(sourceKey, TEXT_LANGUAGE_EN), _
        TextCatalog_Get(replacementKey, TEXT_LANGUAGE_EN), 1, -1, vbTextCompare)

End Sub


'------------------------------------------------------------------------------
' FR: Construit la map Missing Baseline Rex Message a partir des donnees fournies par l'appelant.
' EN: Builds the Missing Baseline Rex Message map from data supplied by the caller.
'------------------------------------------------------------------------------

Public Function CalcBridge_BuildMissingBaselineRexMessage( _
    ByVal missingIds As Object, _
    ByVal idToWbs As Object) As String

    Dim idsLine As String
    Dim wbsLine As String

    idsLine = CalcBridge_BuildInlineList(missingIds, 20)
    wbsLine = CalcBridge_BuildInlineWBSList(missingIds, idToWbs, 20)

    CalcBridge_BuildMissingBaselineRexMessage = PlanningMessageText_Format( _
        "DIAG.ANALYTICS.MISSING_BASELINE_REX", _
        TextCatalog_Arguments("Ids", idsLine, "Wbs", wbsLine), _
        TextCatalog_Arguments("Ids", idsLine, "Wbs", wbsLine))

End Function


'------------------------------------------------------------------------------
' FR: Construit la collection LOE Explicit Predecessor Message a partir des donnees fournies par l'appelant.
' EN: Builds the LOE Explicit Predecessor Message collection from data supplied by the caller.
'------------------------------------------------------------------------------

Public Function CalcBridge_BuildLOEExplicitPredecessorMessage(ByVal details As Collection) As String

    Dim detail As Object
    Dim blocksFR As String
    Dim blocksEN As String

    For Each detail In details
        CalcBridge_AppendTextBlock blocksFR, _
            TextCatalog_Format("DIAG.LOE.EXPLICIT.DETAIL_BLOCK", TEXT_LANGUAGE_FR, _
                TextCatalog_Arguments("SuccWbs", CStr(detail("Succ WBS")), "SuccId", CStr(detail("Succ ID")), _
                    "LoeWbs", CStr(detail("LOE WBS")), "LoeId", CStr(detail("LOE ID")), "Link", CalcBridge_FormatLOELinkLabel(detail)))

        CalcBridge_AppendTextBlock blocksEN, _
            TextCatalog_Format("DIAG.LOE.EXPLICIT.DETAIL_BLOCK", TEXT_LANGUAGE_EN, _
                TextCatalog_Arguments("SuccWbs", CStr(detail("Succ WBS")), "SuccId", CStr(detail("Succ ID")), _
                    "LoeWbs", CStr(detail("LOE WBS")), "LoeId", CStr(detail("LOE ID")), "Link", CalcBridge_FormatLOELinkLabel(detail)))
    Next detail

    CalcBridge_BuildLOEExplicitPredecessorMessage = PlanningMessageText_Format( _
        "DIAG.LOE.EXPLICIT.MESSAGE", _
        TextCatalog_Arguments("Details", blocksFR), TextCatalog_Arguments("Details", blocksEN))

End Function


'------------------------------------------------------------------------------
' FR: Construit la collection LOE Parent Predecessor Message a partir des donnees fournies par l'appelant.
' EN: Builds the LOE Parent Predecessor Message collection from data supplied by the caller.
'------------------------------------------------------------------------------

Public Function CalcBridge_BuildLOEParentPredecessorMessage(ByVal details As Collection) As String

    Dim detail As Object
    Dim blocksFR As String
    Dim blocksEN As String

    For Each detail In details
        CalcBridge_AppendTextBlock blocksFR, _
            TextCatalog_Format("DIAG.LOE.PARENT.DETAIL_BLOCK", TEXT_LANGUAGE_FR, _
                TextCatalog_Arguments("SuccWbs", CStr(detail("Succ WBS")), "SuccId", CStr(detail("Succ ID")), _
                    "ParentWbs", CStr(detail("Parent WBS")), "ParentId", CStr(detail("Parent ID")), _
                    "LoeWbs", CStr(detail("LOE WBS")), "LoeId", CStr(detail("LOE ID")), "Link", CalcBridge_FormatLOELinkLabel(detail)))

        CalcBridge_AppendTextBlock blocksEN, _
            TextCatalog_Format("DIAG.LOE.PARENT.DETAIL_BLOCK", TEXT_LANGUAGE_EN, _
                TextCatalog_Arguments("SuccWbs", CStr(detail("Succ WBS")), "SuccId", CStr(detail("Succ ID")), _
                    "ParentWbs", CStr(detail("Parent WBS")), "ParentId", CStr(detail("Parent ID")), _
                    "LoeWbs", CStr(detail("LOE WBS")), "LoeId", CStr(detail("LOE ID")), "Link", CalcBridge_FormatLOELinkLabel(detail)))
    Next detail

    CalcBridge_BuildLOEParentPredecessorMessage = PlanningMessageText_Format( _
        "DIAG.LOE.PARENT.MESSAGE", _
        TextCatalog_Arguments("Details", blocksFR), TextCatalog_Arguments("Details", blocksEN))

End Function


'------------------------------------------------------------------------------
' FR: Ajoute la valeur Text Block a la structure cible fournie par l'appelant.
' EN: Adds the Text Block value to the target structure supplied by the caller.
'------------------------------------------------------------------------------

Private Sub CalcBridge_AppendTextBlock( _
    ByRef target As String, _
    ByVal blockText As String)

    If target <> "" Then target = target & vbCrLf & vbCrLf
    target = target & blockText

End Sub


'------------------------------------------------------------------------------
' FR: Retourne la map Wbs For ID sans exposer de mutateur sur l'etat source.
' EN: Returns the Wbs For ID map without exposing a mutator for source state.
'------------------------------------------------------------------------------

Public Function CalcBridge_GetWbsForId( _
    ByVal idToWbs As Object, _
    ByVal idVal As String) As String

    idVal = Trim$(CStr(idVal))

    If idVal <> "" Then
        If Not idToWbs Is Nothing Then
            If idToWbs.Exists(idVal) Then
                CalcBridge_GetWbsForId = CStr(idToWbs(idVal))
                Exit Function
            End If
        End If
    End If

    CalcBridge_GetWbsForId = idVal

End Function


'------------------------------------------------------------------------------
' FR: Normalise ou formate Format LOE Link Label selon le contrat canonique du composant.
' EN: Normalizes or formats Format LOE Link Label according to the component contract.
'------------------------------------------------------------------------------

Private Function CalcBridge_FormatLOELinkLabel(ByVal detail As Object) As String

    Dim lagVal As Double
    Dim signText As String

    lagVal = CDbl(detail("Lag"))

    If lagVal >= 0 Then
        signText = "+"
    Else
        signText = ""
    End If

    CalcBridge_FormatLOELinkLabel = CStr(detail("Link Type")) & signText & CStr(CLng(lagVal))

End Function


'------------------------------------------------------------------------------
' FR: Construit la collection Warning Ack Token List a partir des donnees fournies par l'appelant.
' EN: Builds the Warning Ack Token List collection from data supplied by the caller.
'------------------------------------------------------------------------------

Public Function CalcBridge_BuildWarningAckTokenList( _
    ByVal idsDict As Object, _
    ByVal ackTokensById As Object) As String

    Dim key As Variant
    Dim tokenText As String

    If idsDict Is Nothing Then Exit Function
    If ackTokensById Is Nothing Then Exit Function

    For Each key In idsDict.Keys
        If ackTokensById.Exists(CStr(key)) Then
            If Trim$(CStr(ackTokensById(CStr(key)))) <> "" Then
                If tokenText <> "" Then tokenText = tokenText & ";"
                tokenText = tokenText & Trim$(CStr(ackTokensById(CStr(key))))
            End If
        End If
    Next key

    CalcBridge_BuildWarningAckTokenList = tokenText

End Function



'------------------------------------------------------------------------------
' FR: Construit la map Constraint Diagnostic Message a partir des donnees fournies par l'appelant.
' EN: Builds the Constraint Diagnostic Message map from data supplied by the caller.
'------------------------------------------------------------------------------

Public Function CalcBridge_BuildConstraintDiagnosticMessage( _
    ByVal diag As Object, _
    ByVal idToWbs As Object, _
    ByVal idToTaskName As Object, _
    ByVal cascadeDiagnostics As Object, _
    ByVal contextKey As String) As String

    Dim taskId As String
    Dim taskLabel As String
    Dim constraintType As String
    Dim checkedField As String
    Dim relationText As String
    Dim cascadeText As String
    Dim frenchArguments As Object
    Dim englishArguments As Object

    If diag Is Nothing Then Exit Function
    If Not diag.Exists("TaskID") Then Exit Function

    taskId = CStr(diag("TaskID"))
    taskLabel = CalcBridge_DiagnosticTaskLabel(taskId, idToWbs, idToTaskName)
    constraintType = CalcBridge_DiagString(diag, "ConstraintType")
    checkedField = CalcBridge_DiagString(diag, "CheckedField")
    relationText = CalcBridge_DiagString(diag, "ExpectedOperator")
    cascadeText = CalcBridge_BuildCascadeRootCauseText(taskId, cascadeDiagnostics, idToWbs, idToTaskName)

    If constraintType = "" Then
        If CalcBridge_DiagString(diag, "ConstraintSide") = "FINISH" Then
            constraintType = "DIAG.CONSTRAINT.TYPE.FINISH_GENERIC"
        Else
            constraintType = "DIAG.CONSTRAINT.TYPE.START_GENERIC"
        End If
    End If
    If checkedField = "" Then checkedField = "DIAG.LABEL.CALCULATED_DATE"
    If relationText = "" Then relationText = "="

    Set frenchArguments = TextCatalog_Arguments( _
        "Task", taskLabel, "ConstraintType", CalcBridge_LocalizedDiagnosticToken(constraintType, TEXT_LANGUAGE_FR), _
        "ConstraintDate", CalcBridge_DiagnosticAnyValueText(CalcBridge_DiagValue(diag, "ConstraintDate")), _
        "CheckedField", CalcBridge_LocalizedDiagnosticToken(checkedField, TEXT_LANGUAGE_FR), "CheckedValue", CalcBridge_DiagnosticAnyValueText(CalcBridge_DiagValue(diag, "CheckedValue")), _
        "Relation", relationText, "AllowedValue", CalcBridge_DiagnosticAnyValueText(CalcBridge_DiagValue(diag, "AllowedValue")), _
        "CalculatedStart", CalcBridge_DiagnosticAnyValueText(CalcBridge_DiagValue(diag, "CalculatedStart")), _
        "CalculatedFinish", CalcBridge_DiagnosticAnyValueText(CalcBridge_DiagValue(diag, "CalculatedFinish")), _
        "Cascade", cascadeText)
    Set englishArguments = TextCatalog_Arguments( _
        "Task", taskLabel, "ConstraintType", CalcBridge_LocalizedDiagnosticToken(constraintType, TEXT_LANGUAGE_EN), _
        "ConstraintDate", CalcBridge_DiagnosticAnyValueText(CalcBridge_DiagValue(diag, "ConstraintDate")), _
        "CheckedField", CalcBridge_LocalizedDiagnosticToken(checkedField, TEXT_LANGUAGE_EN), "CheckedValue", CalcBridge_DiagnosticAnyValueText(CalcBridge_DiagValue(diag, "CheckedValue")), _
        "Relation", relationText, "AllowedValue", CalcBridge_DiagnosticAnyValueText(CalcBridge_DiagValue(diag, "AllowedValue")), _
        "CalculatedStart", CalcBridge_DiagnosticAnyValueText(CalcBridge_DiagValue(diag, "CalculatedStart")), _
        "CalculatedFinish", CalcBridge_DiagnosticAnyValueText(CalcBridge_DiagValue(diag, "CalculatedFinish")), _
        "Cascade", vbNullString)
    CalcBridge_BuildConstraintDiagnosticMessage = PlanningMessageText_Format( _
        "DIAG.CONSTRAINT.STRUCTURED_DETAIL", frenchArguments, englishArguments)

End Function

Private Function CalcBridge_LocalizedDiagnosticToken( _
    ByVal value As String, _
    ByVal languageKey As String) As String

    Select Case value
        Case "Start No Earlier Than": value = "DIAG.CONSTRAINT.TYPE.START_NO_EARLIER"
        Case "Start No Later Than": value = "DIAG.CONSTRAINT.TYPE.START_NO_LATER"
        Case "Finish No Earlier Than": value = "DIAG.CONSTRAINT.TYPE.FINISH_NO_EARLIER"
        Case "Finish No Later Than": value = "DIAG.CONSTRAINT.TYPE.FINISH_NO_LATER"
        Case "Must Start On": value = "DIAG.CONSTRAINT.TYPE.MUST_START_ON_CANONICAL"
        Case "Must Finish On": value = "DIAG.CONSTRAINT.TYPE.MUST_FINISH_ON_CANONICAL"
    End Select

    If Left$(value, 5) = "DIAG." Then
        CalcBridge_LocalizedDiagnosticToken = TextCatalog_Get(value, languageKey)
    Else
        CalcBridge_LocalizedDiagnosticToken = value
    End If

End Function



'------------------------------------------------------------------------------
' FR: Construit la map Cascade Root Cause Text a partir des donnees fournies par l'appelant.
' EN: Builds the Cascade Root Cause Text map from data supplied by the caller.
'------------------------------------------------------------------------------

Private Function CalcBridge_BuildCascadeRootCauseText( _
    ByVal taskId As String, _
    ByVal cascadeDiagnostics As Object, _
    ByVal idToWbs As Object, _
    ByVal idToTaskName As Object) As String

    Dim cascadeDiag As Object
    Dim parentId As String
    Dim rootId As String

    If cascadeDiagnostics Is Nothing Then Exit Function
    If Not cascadeDiagnostics.Exists(CStr(taskId)) Then Exit Function
    If Not IsObject(cascadeDiagnostics(CStr(taskId))) Then Exit Function

    Set cascadeDiag = cascadeDiagnostics(CStr(taskId))
    parentId = CalcBridge_DiagString(cascadeDiag, "ParentPropagatedFrom")
    rootId = CalcBridge_DiagString(cascadeDiag, "RootErrorID")

    If parentId = "" And rootId = "" Then Exit Function

    CalcBridge_BuildCascadeRootCauseText = TextCatalog_Format( _
        "DIAG.CONSTRAINT.CASCADE", TEXT_LANGUAGE_FR, _
        TextCatalog_Arguments( _
            "Parent", CalcBridge_DiagnosticTaskLabel(parentId, idToWbs, idToTaskName), _
            "Root", CalcBridge_DiagnosticTaskLabel(rootId, idToWbs, idToTaskName)))

End Function


'------------------------------------------------------------------------------
' FR: Construit la map Forecast Start Dependency Diagnostic Message a partir des donnees fournies par l'appelant.
' EN: Builds the Forecast Start Dependency Diagnostic Message map from data supplied by the caller.
'------------------------------------------------------------------------------

Public Function CalcBridge_BuildForecastStartDependencyDiagnosticMessage( _
    ByVal diag As Object, _
    ByVal idToWbs As Object, _
    ByVal idToTaskName As Object, _
    ByVal contextKey As String) As String

    Dim taskId As String
    Dim predId As String
    Dim taskLabel As String
    Dim predLabel As String
    Dim linkText As String
    Dim predDateKindFr As String
    Dim predDateKindEn As String
    Dim requestedLabelFr As String
    Dim requestedLabelEn As String
    Dim frenchArguments As Object
    Dim englishArguments As Object

    If diag Is Nothing Then Exit Function
    If Not diag.Exists("TaskID") Then Exit Function
    If Not diag.Exists("BlockingPredecessorID") Then Exit Function

    taskId = CStr(diag("TaskID"))
    predId = CStr(diag("BlockingPredecessorID"))
    taskLabel = CalcBridge_DiagnosticTaskLabel(taskId, idToWbs, idToTaskName)
    predLabel = CalcBridge_DiagnosticTaskLabel(predId, idToWbs, idToTaskName)
    linkText = CStr(diag("BlockingLinkType")) & " " & CalcBridge_FormatDiagnosticLag(CDbl(diag("BlockingLag")))

    If UCase$(Trim$(CStr(diag("BlockingPredecessorDateKind")))) = "START" Then
        predDateKindFr = TextCatalog_Get("DIAG.LABEL.PREDECESSOR_START", TEXT_LANGUAGE_FR)
        predDateKindEn = TextCatalog_Get("DIAG.LABEL.PREDECESSOR_START", TEXT_LANGUAGE_EN)
    Else
        predDateKindFr = TextCatalog_Get("DIAG.LABEL.PREDECESSOR_FINISH", TEXT_LANGUAGE_FR)
        predDateKindEn = TextCatalog_Get("DIAG.LABEL.PREDECESSOR_FINISH", TEXT_LANGUAGE_EN)
    End If

    If UCase$(Trim$(contextKey)) = "SCENARIO" Then
        requestedLabelFr = TextCatalog_Get("DIAG.LABEL.REQUESTED_SCENARIO_START", TEXT_LANGUAGE_FR)
        requestedLabelEn = TextCatalog_Get("DIAG.LABEL.REQUESTED_SCENARIO_START", TEXT_LANGUAGE_EN)
    Else
        requestedLabelFr = TextCatalog_Get("DIAG.LABEL.REQUESTED_TEST_START", TEXT_LANGUAGE_FR)
        requestedLabelEn = TextCatalog_Get("DIAG.LABEL.REQUESTED_TEST_START", TEXT_LANGUAGE_EN)
    End If

    Set frenchArguments = TextCatalog_Arguments( _
        "Task", taskLabel, "Predecessor", predLabel, "Link", linkText, _
        "PredecessorDateLabel", predDateKindFr, _
        "PredecessorDate", CalcBridge_DiagnosticDateText(diag("BlockingPredecessorDate")), _
        "MinimumStart", CalcBridge_DiagnosticDateText(diag("MinimumAllowedStart")), _
        "RequestedLabel", requestedLabelFr, "RequestedStart", CalcBridge_DiagnosticDateText(diag("RequestedStart")))
    Set englishArguments = TextCatalog_Arguments( _
        "Task", taskLabel, "Predecessor", predLabel, "Link", linkText, _
        "PredecessorDateLabel", predDateKindEn, _
        "PredecessorDate", CalcBridge_DiagnosticDateText(diag("BlockingPredecessorDate")), _
        "MinimumStart", CalcBridge_DiagnosticDateText(diag("MinimumAllowedStart")), _
        "RequestedLabel", requestedLabelEn, "RequestedStart", CalcBridge_DiagnosticDateText(diag("RequestedStart")))
    CalcBridge_BuildForecastStartDependencyDiagnosticMessage = PlanningMessageText_Format( _
        "DIAG.FORECAST_START.DEPENDENCY", frenchArguments, englishArguments)

End Function
