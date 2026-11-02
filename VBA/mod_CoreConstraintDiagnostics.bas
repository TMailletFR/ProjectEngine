Attribute VB_Name = "mod_CoreConstraintDiagnostics"
Option Explicit

'------------------------------------------------------------------------------
Public Function Core_BuildConstraintDiagnosticMessage( _
    ByRef dataArr As Variant, _
    ByVal rowIdx As Long, _
    ByVal mapCol As Object, _
    ByVal constraintDiagnostics As Object, _
    ByVal taskId As String, _
    ByVal startConstraintType As String, _
    ByVal startConstraintDate As Variant, _
    ByVal finishConstraintType As String, _
    ByVal finishConstraintDate As Variant, _
    ByVal actualStart As Variant, _
    ByVal actualFinish As Variant, _
    ByVal forecastStart As Variant, _
    ByVal forecastFinish As Variant, _
    ByVal calcStart As Variant, _
    ByVal calcFinish As Variant, _
    ByVal predAllowedStart As Variant, _
    ByVal predAllowedFinish As Variant, _
    ByVal allowedStart As Variant, _
    ByVal allowedFinish As Variant, _
    ByVal effectiveDuration As Variant, _
    ByVal mustFinishStart As Variant, _
    ByVal messageKey As String) As String

    Core_RecordConstraintDiagnostic constraintDiagnostics, dataArr, rowIdx, mapCol, taskId, _
        startConstraintType, startConstraintDate, finishConstraintType, finishConstraintDate, _
        actualStart, actualFinish, forecastStart, forecastFinish, calcStart, calcFinish, _
        predAllowedStart, predAllowedFinish, allowedStart, allowedFinish, effectiveDuration, _
        mustFinishStart, messageKey

    Core_BuildConstraintDiagnosticMessage = Core_BuildConstraintCoreMessage(dataArr, rowIdx, mapCol, messageKey)

End Function


'------------------------------------------------------------------------------
' FR: Construit l'objet diagnostic detaille d'une violation de contrainte a partir du contexte de calcul de la tache.
' EN: Builds the detailed diagnostic object for a constraint violation from the task calculation context.
'------------------------------------------------------------------------------
Private Sub Core_RecordConstraintDiagnostic( _
    ByVal constraintDiagnostics As Object, _
    ByRef dataArr As Variant, _
    ByVal rowIdx As Long, _
    ByVal mapCol As Object, _
    ByVal taskId As String, _
    ByVal startConstraintType As String, _
    ByVal startConstraintDate As Variant, _
    ByVal finishConstraintType As String, _
    ByVal finishConstraintDate As Variant, _
    ByVal actualStart As Variant, _
    ByVal actualFinish As Variant, _
    ByVal forecastStart As Variant, _
    ByVal forecastFinish As Variant, _
    ByVal calcStart As Variant, _
    ByVal calcFinish As Variant, _
    ByVal predAllowedStart As Variant, _
    ByVal predAllowedFinish As Variant, _
    ByVal allowedStart As Variant, _
    ByVal allowedFinish As Variant, _
    ByVal effectiveDuration As Variant, _
    ByVal mustFinishStart As Variant, _
    ByVal messageKey As String)

    Dim diag As Object
    Dim idVal As String
    Dim wbsVal As String
    Dim taskName As String
    Dim frText As String
    Dim enText As String
    Dim frContext As String
    Dim enContext As String
    Dim sideVal As String
    Dim typeVal As String
    Dim dateVal As Variant
    Dim checkedField As String
    Dim checkedValue As Variant
    Dim relationVal As String
    Dim allowedValue As Variant

    If constraintDiagnostics Is Nothing Then Exit Sub

    idVal = Trim$(CStr(taskId))
    wbsVal = Trim$(CStr(Core_GetVal(dataArr, rowIdx, mapCol, "WBS")))
    taskName = Trim$(CStr(Core_GetVal(dataArr, rowIdx, mapCol, "Task Name")))
    frText = TextCatalog_Get(messageKey, TEXT_LANGUAGE_FR)
    enText = TextCatalog_Get(messageKey, TEXT_LANGUAGE_EN)
    frContext = TextCatalog_Format( _
        "DIAG.COMMON.CONTEXT", _
        TEXT_LANGUAGE_FR, _
        TextCatalog_Arguments("Id", idVal, "Wbs", wbsVal, "Task", taskName))
    enContext = TextCatalog_Format( _
        "DIAG.COMMON.CONTEXT", _
        TEXT_LANGUAGE_EN, _
        TextCatalog_Arguments("Id", idVal, "Wbs", wbsVal, "Task", taskName))

    Select Case messageKey
        Case "DIAG.CONSTRAINT.UNKNOWN_START_TYPE"
            sideVal = "START": typeVal = Trim$(startConstraintType): dateVal = startConstraintDate
            checkedField = "DIAG.LABEL.START": checkedValue = calcStart: relationVal = "=": allowedValue = startConstraintDate
        Case "DIAG.CONSTRAINT.UNKNOWN_FINISH_TYPE"
            sideVal = "FINISH": typeVal = Trim$(finishConstraintType): dateVal = finishConstraintDate
            checkedField = "DIAG.LABEL.FINISH": checkedValue = calcFinish: relationVal = "=": allowedValue = finishConstraintDate
        Case "DIAG.CONSTRAINT.ACTUAL_START_BEFORE_START"
            sideVal = "START": typeVal = Trim$(startConstraintType): dateVal = startConstraintDate
            checkedField = "DIAG.LABEL.ACTUAL_START": checkedValue = actualStart: relationVal = ">=": allowedValue = Core_FirstValue(allowedStart, startConstraintDate)
        Case "DIAG.CONSTRAINT.ACTUAL_START_AFTER_LATEST"
            sideVal = "START": typeVal = Trim$(startConstraintType): dateVal = startConstraintDate
            checkedField = "DIAG.LABEL.ACTUAL_START": checkedValue = actualStart: relationVal = "<=": allowedValue = startConstraintDate
        Case "DIAG.CONSTRAINT.ACTUAL_START_DIFFERS_MSO"
            sideVal = "START": typeVal = "DIAG.CONSTRAINT.TYPE.MUST_START_ON": dateVal = startConstraintDate
            checkedField = "DIAG.LABEL.ACTUAL_START": checkedValue = actualStart: relationVal = "=": allowedValue = startConstraintDate
        Case "DIAG.CONSTRAINT.ACTUAL_FINISH_BEFORE_FINISH"
            sideVal = "FINISH": typeVal = Trim$(finishConstraintType): dateVal = finishConstraintDate
            checkedField = "DIAG.LABEL.ACTUAL_FINISH": checkedValue = actualFinish: relationVal = ">=": allowedValue = Core_FirstValue(allowedFinish, finishConstraintDate)
        Case "DIAG.CONSTRAINT.ACTUAL_FINISH_AFTER_LATEST"
            sideVal = "FINISH": typeVal = Trim$(finishConstraintType): dateVal = finishConstraintDate
            checkedField = "DIAG.LABEL.ACTUAL_FINISH": checkedValue = actualFinish: relationVal = "<=": allowedValue = finishConstraintDate
        Case "DIAG.CONSTRAINT.ACTUAL_FINISH_DIFFERS_MFO"
            sideVal = "FINISH": typeVal = "DIAG.CONSTRAINT.TYPE.MUST_FINISH_ON": dateVal = finishConstraintDate
            checkedField = "DIAG.LABEL.ACTUAL_FINISH": checkedValue = actualFinish: relationVal = "=": allowedValue = finishConstraintDate
        Case "DIAG.CONSTRAINT.FORECAST_START_BEFORE_START"
            sideVal = "START": typeVal = Trim$(startConstraintType): dateVal = startConstraintDate
            checkedField = "DIAG.LABEL.FORECAST_START": checkedValue = forecastStart: relationVal = ">=": allowedValue = Core_FirstValue(allowedStart, startConstraintDate)
        Case "DIAG.CONSTRAINT.FORECAST_START_AFTER_LATEST"
            sideVal = "START": typeVal = Trim$(startConstraintType): dateVal = startConstraintDate
            checkedField = "DIAG.LABEL.FORECAST_START": checkedValue = forecastStart: relationVal = "<=": allowedValue = startConstraintDate
        Case "DIAG.CONSTRAINT.FORECAST_START_DIFFERS_MSO"
            sideVal = "START": typeVal = "DIAG.CONSTRAINT.TYPE.MUST_START_ON": dateVal = startConstraintDate
            checkedField = "DIAG.LABEL.FORECAST_START": checkedValue = forecastStart: relationVal = "=": allowedValue = startConstraintDate
        Case "DIAG.CONSTRAINT.FORECAST_FINISH_BEFORE_UPSTREAM"
            sideVal = "FINISH": typeVal = "DIAG.CONSTRAINT.TYPE.UPSTREAM_FINISH": dateVal = predAllowedFinish
            checkedField = "DIAG.LABEL.FORECAST_FINISH": checkedValue = forecastFinish: relationVal = ">=": allowedValue = predAllowedFinish
        Case "DIAG.CONSTRAINT.FORECAST_FINISH_BEFORE_FINISH"
            sideVal = "FINISH": typeVal = Trim$(finishConstraintType): dateVal = finishConstraintDate
            checkedField = "DIAG.LABEL.FORECAST_FINISH": checkedValue = forecastFinish: relationVal = ">=": allowedValue = Core_FirstValue(allowedFinish, finishConstraintDate)
        Case "DIAG.CONSTRAINT.FORECAST_FINISH_AFTER_LATEST"
            sideVal = "FINISH": typeVal = Trim$(finishConstraintType): dateVal = finishConstraintDate
            checkedField = "DIAG.LABEL.FORECAST_FINISH": checkedValue = forecastFinish: relationVal = "<=": allowedValue = finishConstraintDate
        Case "DIAG.CONSTRAINT.FORECAST_FINISH_DIFFERS_MFO"
            sideVal = "FINISH": typeVal = "DIAG.CONSTRAINT.TYPE.MUST_FINISH_ON": dateVal = finishConstraintDate
            checkedField = "DIAG.LABEL.FORECAST_FINISH": checkedValue = forecastFinish: relationVal = "=": allowedValue = finishConstraintDate
        Case "DIAG.CONSTRAINT.CALCULATED_FINISH_DIFFERS_MFO"
            sideVal = "FINISH": typeVal = "DIAG.CONSTRAINT.TYPE.MUST_FINISH_ON": dateVal = finishConstraintDate
            checkedField = "DIAG.LABEL.CALCULATED_FINISH": checkedValue = calcFinish: relationVal = "=": allowedValue = finishConstraintDate
        Case "DIAG.CONSTRAINT.CALCULATED_START_BEFORE_MFO_NETWORK"
            sideVal = "FINISH": typeVal = "DIAG.CONSTRAINT.TYPE.MUST_FINISH_ON": dateVal = finishConstraintDate
            checkedField = "DIAG.LABEL.CALCULATED_START": checkedValue = calcStart: relationVal = ">=": allowedValue = Core_FirstValue(allowedFinish, finishConstraintDate)
        Case "DIAG.CONSTRAINT.DURATION_INCOMPATIBLE_MSO_MFO"
            sideVal = "FINISH": typeVal = "DIAG.CONSTRAINT.TYPE.MUST_FINISH_ON": dateVal = finishConstraintDate
            checkedField = "DIAG.LABEL.EFFECTIVE_DURATION": checkedValue = effectiveDuration: relationVal = "=": allowedValue = finishConstraintDate
        Case "DIAG.CONSTRAINT.ACTUAL_START_DIFFERS_MFO_IMPLIED"
            sideVal = "FINISH": typeVal = "DIAG.CONSTRAINT.TYPE.MUST_FINISH_ON": dateVal = finishConstraintDate
            checkedField = "DIAG.LABEL.ACTUAL_START": checkedValue = actualStart: relationVal = "=": allowedValue = finishConstraintDate
        Case "DIAG.CONSTRAINT.FORECAST_START_DIFFERS_MFO_IMPLIED"
            sideVal = "FINISH": typeVal = "DIAG.CONSTRAINT.TYPE.MUST_FINISH_ON": dateVal = finishConstraintDate
            checkedField = "DIAG.LABEL.FORECAST_START": checkedValue = forecastStart: relationVal = "=": allowedValue = finishConstraintDate
        Case "DIAG.CONSTRAINT.CALCULATED_START_DIFFERS_MSO"
            sideVal = "START": typeVal = "DIAG.CONSTRAINT.TYPE.MUST_START_ON": dateVal = startConstraintDate
            checkedField = "DIAG.LABEL.CALCULATED_START": checkedValue = calcStart: relationVal = "=": allowedValue = startConstraintDate
        Case "DIAG.CONSTRAINT.CALCULATED_START_AFTER_LATEST"
            sideVal = "START": typeVal = Trim$(startConstraintType): dateVal = startConstraintDate
            checkedField = "DIAG.LABEL.CALCULATED_START": checkedValue = calcStart: relationVal = "<=": allowedValue = startConstraintDate
        Case "DIAG.CONSTRAINT.CALCULATED_FINISH_AFTER_LATEST"
            sideVal = "FINISH": typeVal = Trim$(finishConstraintType): dateVal = finishConstraintDate
            checkedField = "DIAG.LABEL.CALCULATED_FINISH": checkedValue = calcFinish: relationVal = "<=": allowedValue = finishConstraintDate
        Case Else
            Err.Raise vbObjectError + 2891, "Core_RecordConstraintDiagnostic", _
                PlanningMessageText_Format("DIAG.TECH.UNKNOWN_CONSTRAINT_KEY", _
                    TextCatalog_Arguments("Key", messageKey), TextCatalog_Arguments("Key", messageKey))
    End Select

    Set diag = CreateObject("Scripting.Dictionary")
    diag("Code") = UCase$(Trim$(messageKey))
    diag("TaskID") = CStr(taskId)
    diag("WBS") = Trim$(CStr(Core_GetVal(dataArr, rowIdx, mapCol, "WBS")))
    diag("TaskName") = Trim$(CStr(Core_GetVal(dataArr, rowIdx, mapCol, "Task Name")))
    diag("ConstraintSide") = sideVal
    diag("ConstraintType") = typeVal
    diag("ConstraintDate") = dateVal
    diag("CheckedField") = checkedField
    diag("CheckedValue") = checkedValue
    diag("ExpectedOperator") = relationVal
    diag("AllowedValue") = allowedValue
    diag("ActualStart") = actualStart
    diag("ActualFinish") = actualFinish
    diag("ForecastStart") = forecastStart
    diag("ForecastFinish") = forecastFinish
    diag("CalculatedStart") = calcStart
    diag("CalculatedFinish") = calcFinish
    diag("PredAllowedStart") = predAllowedStart
    diag("PredAllowedFinish") = predAllowedFinish
    diag("AllowedStart") = allowedStart
    diag("AllowedFinish") = allowedFinish
    diag("EffectiveDuration") = effectiveDuration
    diag("MustFinishImpliedStart") = mustFinishStart
    diag("RawFR") = frText
    diag("RawEN") = enText

    Set constraintDiagnostics.Item(CStr(taskId)) = diag

End Sub


'------------------------------------------------------------------------------
' FR: Retourne la premiere valeur renseignee entre une valeur prioritaire et une valeur de repli.
' EN: Returns the first populated value between a primary value and a fallback value.
'------------------------------------------------------------------------------
Private Function Core_FirstValue(ByVal primaryValue As Variant, ByVal fallbackValue As Variant) As Variant

    If HasValue(primaryValue) Then
        Core_FirstValue = primaryValue
    Else
        Core_FirstValue = fallbackValue
    End If

End Function

'------------------------------------------------------------------------------
' FR: Construit le message Core FR/EN minimal associe a une violation de contrainte, avec ID, WBS et nom de tache.
' EN: Builds the minimal FR/EN Core message for a constraint violation, including task ID, WBS, and name.
'------------------------------------------------------------------------------
Private Function Core_BuildConstraintCoreMessage( _
    ByRef dataArr As Variant, _
    ByVal rowIdx As Long, _
    ByVal mapCol As Object, _
    ByVal messageKey As String) As String

    Dim idVal As String
    Dim wbsVal As String
    Dim taskName As String
    Dim frText As String
    Dim enText As String
    Dim frContext As String
    Dim enContext As String

    idVal = Trim$(CStr(Core_GetVal(dataArr, rowIdx, mapCol, "ID")))
    wbsVal = Trim$(CStr(Core_GetVal(dataArr, rowIdx, mapCol, "WBS")))
    taskName = Trim$(CStr(Core_GetVal(dataArr, rowIdx, mapCol, "Task Name")))
    frText = TextCatalog_Get(messageKey, TEXT_LANGUAGE_FR)
    enText = TextCatalog_Get(messageKey, TEXT_LANGUAGE_EN)
    frContext = TextCatalog_Format( _
        "DIAG.COMMON.CONTEXT", _
        TEXT_LANGUAGE_FR, _
        TextCatalog_Arguments("Id", idVal, "Wbs", wbsVal, "Task", taskName))
    enContext = TextCatalog_Format( _
        "DIAG.COMMON.CONTEXT", _
        TEXT_LANGUAGE_EN, _
        TextCatalog_Arguments("Id", idVal, "Wbs", wbsVal, "Task", taskName))

    Core_BuildConstraintCoreMessage = BiMsg( _
        frText & vbCrLf & vbCrLf & frContext, _
        enText & vbCrLf & vbCrLf & enContext)

End Function

'------------------------------------------------------------------------------
' FR:
' Enregistre une erreur bloquante sur une tache et la reporte immediatement
' dans les colonnes d'erreur de la ligne Core.
'
' EN:
' Registers a blocking error on a task and immediately mirrors it to the Core
' row error columns.
'
' Entrees / Inputs:
' - currentId, rowIdx et message metier deja formule.
' - dataArr/mapCol pour mettre a jour la ligne.
' - blockingErrors pour memoriser la racine de blocage.
'
' Sorties / Outputs:
' - blockingErrors(currentId) = message.
' - Error flag/ErrorMsg mis a jour dans dataArr.
'
' Appele par / Called by:
' - Core_ComputeOneLeafTask, Core_MarkTopoFailure, Core_ValidateLOEAsNonPredecessor, Core_ApplyLOEPostProcess.
'
' Notes:
' - Point commun de creation des erreurs dures qui seront ensuite propagees.

