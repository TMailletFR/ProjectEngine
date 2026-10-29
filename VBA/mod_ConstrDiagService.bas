Attribute VB_Name = "mod_ConstrDiagService"
Option Explicit

'===============================================================================
' MODULE : mod_ConstrDiagService
' DOMAINE / DOMAIN : Shared Infrastructure
'
' FR
' Possede le workflow specialise indique par son nom et expose ses contrats stables.
' Ne possede pas les domaines appeles en dependance.
'
' EN
' Owns the named specialized workflow and exposes its stable contracts.
' Does not own the domains it calls as dependencies.
'
' CONTRATS / CONTRACTS : BuildConstraintValidationMessage, AddConstraintWarning
' CALLBACKS EXTERNES / EXTERNAL CALLBACKS : Aucun / None
'===============================================================================


'------------------------------------------------------------------------------
' FR: Construit le message bilingue d'un diagnostic Constraints pour une ligne source.
' EN: Builds the bilingual Constraints diagnostic message for a source row.
'------------------------------------------------------------------------------
Public Function BuildConstraintValidationMessage( _
    ByRef arrConstraints As Variant, _
    ByVal rowIdx As Long, _
    ByVal mapConstraints As Object, _
    ByVal messageKey As String, _
    Optional ByVal explanationKey As String = "") As String

    Dim idVal As String
    Dim wbsVal As String
    Dim taskName As String
    Dim frText As String
    Dim enText As String
    Dim frPrefix As String
    Dim enPrefix As String
    Dim frExplanation As String
    Dim enExplanation As String

    idVal = Trim$(CStr(arrConstraints(rowIdx, mapConstraints(VTS_COL_ID))))
    wbsVal = Trim$(CStr(arrConstraints(rowIdx, mapConstraints(VTS_COL_WBS))))
    taskName = Trim$(CStr(arrConstraints(rowIdx, mapConstraints(VTS_COL_TASK_NAME))))
    frPrefix = TextCatalog_Get(messageKey, TEXT_LANGUAGE_FR)
    enPrefix = TextCatalog_Get(messageKey, TEXT_LANGUAGE_EN)
    If Trim$(explanationKey) <> "" Then
        frExplanation = TextCatalog_Get(explanationKey, TEXT_LANGUAGE_FR)
        enExplanation = TextCatalog_Get(explanationKey, TEXT_LANGUAGE_EN)
    End If

    frText = frPrefix
    If Trim$(frExplanation) <> "" Then
        frText = frText & vbCrLf & vbCrLf & "-> " & frExplanation
    End If
    frText = frText & vbCrLf & vbCrLf & TextCatalog_Format( _
        "DIAG.COMMON.CONTEXT", _
        TEXT_LANGUAGE_FR, _
        TextCatalog_Arguments("Id", idVal, "Wbs", wbsVal, "Task", taskName))

    enText = enPrefix
    If Trim$(enExplanation) <> "" Then
        enText = enText & vbCrLf & vbCrLf & "-> " & enExplanation
    End If
    enText = enText & vbCrLf & vbCrLf & TextCatalog_Format( _
        "DIAG.COMMON.CONTEXT", _
        TEXT_LANGUAGE_EN, _
        TextCatalog_Arguments("Id", idVal, "Wbs", wbsVal, "Task", taskName))

    BuildConstraintValidationMessage = BiMsg(frText, enText)

End Function

'------------------------------------------------------------------------------
' FR: Ajoute un warning Constraints au flux console sans relire les regles metier.
' EN: Adds a Constraints warning to the console stream without re-reading business rules.
'------------------------------------------------------------------------------
Public Sub AddConstraintWarning( _
    ByVal consoleMessages As Collection, _
    ByVal messageText As String, _
    Optional ByVal historyHandled As Boolean = False, _
    Optional ByVal eventType As String = "", _
    Optional ByVal eventHash As String = "", _
    Optional ByVal historyReceipt As Object = Nothing)

    If consoleMessages Is Nothing Then Exit Sub

    CalcBridge_AddConsoleMessage consoleMessages, "WARNING", messageText, historyHandled, eventType, eventHash, _
        historyReceipt:=historyReceipt

End Sub

