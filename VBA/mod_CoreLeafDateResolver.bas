Attribute VB_Name = "mod_CoreLeafDateResolver"
Option Explicit

Public Function CoreLeaf_ResolveDatesAndConstraints( _
    ByRef dataArr As Variant, _
    ByVal rowIdx As Long, ByVal mapCol As Object, _
    ByVal blockingErrors As Object, ByVal taskId As String, _
    ByVal constraintDiagnostics As Object, ByVal dependencyDiagnostics As Object, _
    ByVal coreDiagnostics As Object, ByVal constraintActive As String, _
    ByVal startConstraintType As String, ByVal startConstraintDate As Variant, _
    ByVal finishConstraintType As String, ByVal finishConstraintDate As Variant, _
    ByVal actualStart As Variant, ByVal actualFinish As Variant, _
    ByVal forecastStart As Variant, ByVal forecastFinish As Variant, _
    ByVal baselineStart As Variant, ByVal effectiveDuration As Variant, _
    ByVal calType As String, ByVal predAllowedStart As Variant, _
    ByVal predAllowedFinish As Variant, ByVal predDiagPredId As String, _
    ByVal predDiagLinkType As String, ByVal predDiagLag As Double, _
    ByVal predDiagCandidateDate As Variant, ByVal predDiagPredecessorDate As Variant, _
    ByVal predDiagPredecessorDateKind As String, ByVal predDiagSummarySourceId As String, _
    ByVal leafProfileEnabled As Boolean, ByRef leafPhaseStart As Double, _
    ByRef calcStart As Variant, _
    ByRef calcFinish As Variant) As Boolean

    Dim constraintAllowedStart As Variant
    Dim constraintAllowedFinish As Variant
    Dim constraintLatestStart As Variant
    Dim constraintLatestFinish As Variant
    Dim constraintMustStart As Variant
    Dim constraintMustFinish As Variant
    Dim allowedStart As Variant
    Dim allowedFinish As Variant
    Dim mustFinishStart As Variant
    Dim sourceStart As Variant
    Dim sourceFinish As Variant
    Dim hasExplicitStart As Boolean

    If constraintActive = "YES" Then
        If startConstraintType = "Start No Earlier Than" Then
            constraintAllowedStart = startConstraintDate
        ElseIf startConstraintType = "Start No Later Than" Then
            constraintLatestStart = startConstraintDate
        ElseIf startConstraintType = "Must Start On" Then
            constraintMustStart = startConstraintDate
        ElseIf startConstraintType <> "" Then
            Core_AddBlockingError dataArr, rowIdx, mapCol, blockingErrors, taskId, _
                Core_BuildConstraintDiagnosticMessage(dataArr, rowIdx, mapCol, constraintDiagnostics, taskId, startConstraintType, startConstraintDate, finishConstraintType, finishConstraintDate, actualStart, actualFinish, forecastStart, forecastFinish, calcStart, calcFinish, predAllowedStart, predAllowedFinish, allowedStart, allowedFinish, effectiveDuration, mustFinishStart, _
                    "DIAG.CONSTRAINT.UNKNOWN_START_TYPE")
            Exit Function
        End If

        If finishConstraintType = "Finish No Earlier Than" Then
            constraintAllowedFinish = finishConstraintDate
        ElseIf finishConstraintType = "Finish No Later Than" Then
            constraintLatestFinish = finishConstraintDate
        ElseIf finishConstraintType = "Must Finish On" Then
            constraintMustFinish = finishConstraintDate
        ElseIf finishConstraintType <> "" Then
            Core_AddBlockingError dataArr, rowIdx, mapCol, blockingErrors, taskId, _
                Core_BuildConstraintDiagnosticMessage(dataArr, rowIdx, mapCol, constraintDiagnostics, taskId, startConstraintType, startConstraintDate, finishConstraintType, finishConstraintDate, actualStart, actualFinish, forecastStart, forecastFinish, calcStart, calcFinish, predAllowedStart, predAllowedFinish, allowedStart, allowedFinish, effectiveDuration, mustFinishStart, _
                    "DIAG.CONSTRAINT.UNKNOWN_FINISH_TYPE")
            Exit Function
        End If
    End If

    allowedStart = Core_MaxDateIfBoth(predAllowedStart, constraintAllowedStart)
    allowedFinish = Core_MaxDateIfBoth(predAllowedFinish, constraintAllowedFinish)
    hasExplicitStart = HasValue(actualStart) Or HasValue(forecastStart) Or HasValue(constraintAllowedStart) Or HasValue(constraintMustStart)

    If HasValue(actualStart) And HasValue(predAllowedStart) Then
        If CDbl(actualStart) < CDbl(predAllowedStart) Then
            Core_AddBlockingError dataArr, rowIdx, mapCol, blockingErrors, taskId, _
                TextCatalog_Get("CORE.ERROR.ACTUAL_START_DEPENDENCIES", TEXT_LANGUAGE_EN), _
                "CORE.ERROR.ACTUAL_START_DEPENDENCIES", Nothing, predDiagPredId, coreDiagnostics, "ROOT", "DEPENDENCY"
            Exit Function
        End If
    End If

    If HasValue(actualStart) And HasValue(constraintAllowedStart) Then
        If CDbl(actualStart) < CDbl(constraintAllowedStart) Then
            Core_AddBlockingError dataArr, rowIdx, mapCol, blockingErrors, taskId, _
                Core_BuildConstraintDiagnosticMessage(dataArr, rowIdx, mapCol, constraintDiagnostics, taskId, startConstraintType, startConstraintDate, finishConstraintType, finishConstraintDate, actualStart, actualFinish, forecastStart, forecastFinish, calcStart, calcFinish, predAllowedStart, predAllowedFinish, allowedStart, allowedFinish, effectiveDuration, mustFinishStart, _
                    "DIAG.CONSTRAINT.ACTUAL_START_BEFORE_START")
            Exit Function
        End If
    End If

    If HasValue(actualStart) And HasValue(constraintLatestStart) Then
        If CDbl(actualStart) > CDbl(constraintLatestStart) Then
            Core_AddBlockingError dataArr, rowIdx, mapCol, blockingErrors, taskId, _
                Core_BuildConstraintDiagnosticMessage(dataArr, rowIdx, mapCol, constraintDiagnostics, taskId, startConstraintType, startConstraintDate, finishConstraintType, finishConstraintDate, actualStart, actualFinish, forecastStart, forecastFinish, calcStart, calcFinish, predAllowedStart, predAllowedFinish, allowedStart, allowedFinish, effectiveDuration, mustFinishStart, _
                    "DIAG.CONSTRAINT.ACTUAL_START_AFTER_LATEST")
            Exit Function
        End If
    End If

    If HasValue(actualStart) And HasValue(constraintMustStart) Then
        If CDbl(actualStart) <> CDbl(constraintMustStart) Then
            Core_AddBlockingError dataArr, rowIdx, mapCol, blockingErrors, taskId, _
                Core_BuildConstraintDiagnosticMessage(dataArr, rowIdx, mapCol, constraintDiagnostics, taskId, startConstraintType, startConstraintDate, finishConstraintType, finishConstraintDate, actualStart, actualFinish, forecastStart, forecastFinish, calcStart, calcFinish, predAllowedStart, predAllowedFinish, allowedStart, allowedFinish, effectiveDuration, mustFinishStart, _
                    "DIAG.CONSTRAINT.ACTUAL_START_DIFFERS_MSO")
            Exit Function
        End If
    End If

    If HasValue(actualFinish) And HasValue(predAllowedFinish) Then
        If CDbl(actualFinish) < CDbl(predAllowedFinish) Then
            Core_AddBlockingError dataArr, rowIdx, mapCol, blockingErrors, taskId, _
                TextCatalog_Get("CORE.ERROR.ACTUAL_FINISH_CONSTRAINTS", TEXT_LANGUAGE_EN), _
                "CORE.ERROR.ACTUAL_FINISH_CONSTRAINTS", Nothing, vbNullString, coreDiagnostics, "ROOT", "CONSTRAINT"
            Exit Function
        End If
    End If

    If HasValue(actualFinish) And HasValue(constraintAllowedFinish) Then
        If CDbl(actualFinish) < CDbl(constraintAllowedFinish) Then
            Core_AddBlockingError dataArr, rowIdx, mapCol, blockingErrors, taskId, _
                Core_BuildConstraintDiagnosticMessage(dataArr, rowIdx, mapCol, constraintDiagnostics, taskId, startConstraintType, startConstraintDate, finishConstraintType, finishConstraintDate, actualStart, actualFinish, forecastStart, forecastFinish, calcStart, calcFinish, predAllowedStart, predAllowedFinish, allowedStart, allowedFinish, effectiveDuration, mustFinishStart, _
                    "DIAG.CONSTRAINT.ACTUAL_FINISH_BEFORE_FINISH")
            Exit Function
        End If
    End If

    If HasValue(actualFinish) And HasValue(constraintLatestFinish) Then
        If CDbl(actualFinish) > CDbl(constraintLatestFinish) Then
            Core_AddBlockingError dataArr, rowIdx, mapCol, blockingErrors, taskId, _
                Core_BuildConstraintDiagnosticMessage(dataArr, rowIdx, mapCol, constraintDiagnostics, taskId, startConstraintType, startConstraintDate, finishConstraintType, finishConstraintDate, actualStart, actualFinish, forecastStart, forecastFinish, calcStart, calcFinish, predAllowedStart, predAllowedFinish, allowedStart, allowedFinish, effectiveDuration, mustFinishStart, _
                    "DIAG.CONSTRAINT.ACTUAL_FINISH_AFTER_LATEST")
            Exit Function
        End If
    End If

    If HasValue(actualFinish) And HasValue(constraintMustFinish) Then
        If CDbl(actualFinish) <> CDbl(constraintMustFinish) Then
            Core_AddBlockingError dataArr, rowIdx, mapCol, blockingErrors, taskId, _
                Core_BuildConstraintDiagnosticMessage(dataArr, rowIdx, mapCol, constraintDiagnostics, taskId, startConstraintType, startConstraintDate, finishConstraintType, finishConstraintDate, actualStart, actualFinish, forecastStart, forecastFinish, calcStart, calcFinish, predAllowedStart, predAllowedFinish, allowedStart, allowedFinish, effectiveDuration, mustFinishStart, _
                    "DIAG.CONSTRAINT.ACTUAL_FINISH_DIFFERS_MFO")
            Exit Function
        End If
    End If

    If HasValue(forecastStart) And HasValue(predAllowedStart) Then
        If CDbl(forecastStart) < CDbl(predAllowedStart) Then
            Core_RecordForecastStartDependencyDiagnostic dependencyDiagnostics, taskId, forecastStart, predAllowedStart, _
                predDiagPredId, predDiagLinkType, predDiagLag, predDiagCandidateDate, _
                predDiagPredecessorDate, predDiagPredecessorDateKind, predDiagSummarySourceId
            Core_AddBlockingError dataArr, rowIdx, mapCol, blockingErrors, taskId, _
                TextCatalog_Get("CORE.ERROR.FORECAST_START_DEPENDENCIES", TEXT_LANGUAGE_EN), _
                "CORE.ERROR.FORECAST_START_DEPENDENCIES", Nothing, predDiagPredId, coreDiagnostics, "ROOT", "FORECAST"
            Exit Function
        End If
    End If

    If HasValue(forecastStart) And HasValue(constraintAllowedStart) Then
        If CDbl(forecastStart) < CDbl(constraintAllowedStart) Then
            Core_AddBlockingError dataArr, rowIdx, mapCol, blockingErrors, taskId, _
                Core_BuildConstraintDiagnosticMessage(dataArr, rowIdx, mapCol, constraintDiagnostics, taskId, startConstraintType, startConstraintDate, finishConstraintType, finishConstraintDate, actualStart, actualFinish, forecastStart, forecastFinish, calcStart, calcFinish, predAllowedStart, predAllowedFinish, allowedStart, allowedFinish, effectiveDuration, mustFinishStart, _
                    "DIAG.CONSTRAINT.FORECAST_START_BEFORE_START")
            Exit Function
        End If
    End If

    If HasValue(forecastStart) And HasValue(constraintLatestStart) Then
        If CDbl(forecastStart) > CDbl(constraintLatestStart) Then
            Core_AddBlockingError dataArr, rowIdx, mapCol, blockingErrors, taskId, _
                Core_BuildConstraintDiagnosticMessage(dataArr, rowIdx, mapCol, constraintDiagnostics, taskId, startConstraintType, startConstraintDate, finishConstraintType, finishConstraintDate, actualStart, actualFinish, forecastStart, forecastFinish, calcStart, calcFinish, predAllowedStart, predAllowedFinish, allowedStart, allowedFinish, effectiveDuration, mustFinishStart, _
                    "DIAG.CONSTRAINT.FORECAST_START_AFTER_LATEST")
            Exit Function
        End If
    End If

    If HasValue(forecastStart) And HasValue(constraintMustStart) Then
        If CDbl(forecastStart) <> CDbl(constraintMustStart) Then
            Core_AddBlockingError dataArr, rowIdx, mapCol, blockingErrors, taskId, _
                Core_BuildConstraintDiagnosticMessage(dataArr, rowIdx, mapCol, constraintDiagnostics, taskId, startConstraintType, startConstraintDate, finishConstraintType, finishConstraintDate, actualStart, actualFinish, forecastStart, forecastFinish, calcStart, calcFinish, predAllowedStart, predAllowedFinish, allowedStart, allowedFinish, effectiveDuration, mustFinishStart, _
                    "DIAG.CONSTRAINT.FORECAST_START_DIFFERS_MSO")
            Exit Function
        End If
    End If

    If HasValue(forecastFinish) And HasValue(predAllowedFinish) Then
        If CDbl(forecastFinish) < CDbl(predAllowedFinish) Then
            Core_AddBlockingError dataArr, rowIdx, mapCol, blockingErrors, taskId, _
                Core_BuildConstraintDiagnosticMessage(dataArr, rowIdx, mapCol, constraintDiagnostics, taskId, startConstraintType, startConstraintDate, finishConstraintType, finishConstraintDate, actualStart, actualFinish, forecastStart, forecastFinish, calcStart, calcFinish, predAllowedStart, predAllowedFinish, allowedStart, allowedFinish, effectiveDuration, mustFinishStart, _
                    "DIAG.CONSTRAINT.FORECAST_FINISH_BEFORE_UPSTREAM")
            Exit Function
        End If
    End If

    If HasValue(forecastFinish) And HasValue(constraintAllowedFinish) Then
        If CDbl(forecastFinish) < CDbl(constraintAllowedFinish) Then
            Core_AddBlockingError dataArr, rowIdx, mapCol, blockingErrors, taskId, _
                Core_BuildConstraintDiagnosticMessage(dataArr, rowIdx, mapCol, constraintDiagnostics, taskId, startConstraintType, startConstraintDate, finishConstraintType, finishConstraintDate, actualStart, actualFinish, forecastStart, forecastFinish, calcStart, calcFinish, predAllowedStart, predAllowedFinish, allowedStart, allowedFinish, effectiveDuration, mustFinishStart, _
                    "DIAG.CONSTRAINT.FORECAST_FINISH_BEFORE_FINISH")
            Exit Function
        End If
    End If

    If HasValue(forecastFinish) And HasValue(constraintLatestFinish) Then
        If CDbl(forecastFinish) > CDbl(constraintLatestFinish) Then
            Core_AddBlockingError dataArr, rowIdx, mapCol, blockingErrors, taskId, _
                Core_BuildConstraintDiagnosticMessage(dataArr, rowIdx, mapCol, constraintDiagnostics, taskId, startConstraintType, startConstraintDate, finishConstraintType, finishConstraintDate, actualStart, actualFinish, forecastStart, forecastFinish, calcStart, calcFinish, predAllowedStart, predAllowedFinish, allowedStart, allowedFinish, effectiveDuration, mustFinishStart, _
                    "DIAG.CONSTRAINT.FORECAST_FINISH_AFTER_LATEST")
            Exit Function
        End If
    End If

    If HasValue(forecastFinish) And HasValue(constraintMustFinish) Then
        If CDbl(forecastFinish) <> CDbl(constraintMustFinish) Then
            Core_AddBlockingError dataArr, rowIdx, mapCol, blockingErrors, taskId, _
                Core_BuildConstraintDiagnosticMessage(dataArr, rowIdx, mapCol, constraintDiagnostics, taskId, startConstraintType, startConstraintDate, finishConstraintType, finishConstraintDate, actualStart, actualFinish, forecastStart, forecastFinish, calcStart, calcFinish, predAllowedStart, predAllowedFinish, allowedStart, allowedFinish, effectiveDuration, mustFinishStart, _
                    "DIAG.CONSTRAINT.FORECAST_FINISH_DIFFERS_MFO")
            Exit Function
        End If
    End If

    If HasValue(constraintMustFinish) Then
        If Not HasValue(effectiveDuration) Then
            Core_AddBlockingError dataArr, rowIdx, mapCol, blockingErrors, taskId, _
                TextCatalog_Get("CORE.ERROR.BASELINE_DURATION_MISSING", TEXT_LANGUAGE_EN), _
                "CORE.ERROR.BASELINE_DURATION_MISSING", Nothing, vbNullString, coreDiagnostics, "ROOT", "BASELINE"
            Exit Function
        End If

        mustFinishStart = SubtractWorkingDays(constraintMustFinish, effectiveDuration, calType)

        If HasValue(allowedFinish) Then
            If CDbl(allowedFinish) > CDbl(constraintMustFinish) Then
                Core_AddBlockingError dataArr, rowIdx, mapCol, blockingErrors, taskId, _
                    Core_BuildConstraintDiagnosticMessage(dataArr, rowIdx, mapCol, constraintDiagnostics, taskId, startConstraintType, startConstraintDate, finishConstraintType, finishConstraintDate, actualStart, actualFinish, forecastStart, forecastFinish, calcStart, calcFinish, predAllowedStart, predAllowedFinish, allowedStart, allowedFinish, effectiveDuration, mustFinishStart, _
                        "DIAG.CONSTRAINT.CALCULATED_FINISH_DIFFERS_MFO")
                Exit Function
            End If
        End If

        If HasValue(allowedStart) Then
            If CDbl(allowedStart) > CDbl(mustFinishStart) Then
                Core_AddBlockingError dataArr, rowIdx, mapCol, blockingErrors, taskId, _
                    Core_BuildConstraintDiagnosticMessage(dataArr, rowIdx, mapCol, constraintDiagnostics, taskId, startConstraintType, startConstraintDate, finishConstraintType, finishConstraintDate, actualStart, actualFinish, forecastStart, forecastFinish, calcStart, calcFinish, predAllowedStart, predAllowedFinish, allowedStart, allowedFinish, effectiveDuration, mustFinishStart, _
                        "DIAG.CONSTRAINT.CALCULATED_START_BEFORE_MFO_NETWORK")
                Exit Function
            End If
        End If

        If HasValue(constraintMustStart) Then
            If CDbl(constraintMustStart) <> CDbl(mustFinishStart) Then
                Core_AddBlockingError dataArr, rowIdx, mapCol, blockingErrors, taskId, _
                    Core_BuildConstraintDiagnosticMessage(dataArr, rowIdx, mapCol, constraintDiagnostics, taskId, startConstraintType, startConstraintDate, finishConstraintType, finishConstraintDate, actualStart, actualFinish, forecastStart, forecastFinish, calcStart, calcFinish, predAllowedStart, predAllowedFinish, allowedStart, allowedFinish, effectiveDuration, mustFinishStart, _
                        "DIAG.CONSTRAINT.DURATION_INCOMPATIBLE_MSO_MFO")
                Exit Function
            End If
        End If

        If HasValue(actualStart) Then
            If CDbl(actualStart) <> CDbl(mustFinishStart) Then
                Core_AddBlockingError dataArr, rowIdx, mapCol, blockingErrors, taskId, _
                    Core_BuildConstraintDiagnosticMessage(dataArr, rowIdx, mapCol, constraintDiagnostics, taskId, startConstraintType, startConstraintDate, finishConstraintType, finishConstraintDate, actualStart, actualFinish, forecastStart, forecastFinish, calcStart, calcFinish, predAllowedStart, predAllowedFinish, allowedStart, allowedFinish, effectiveDuration, mustFinishStart, _
                        "DIAG.CONSTRAINT.ACTUAL_START_DIFFERS_MFO_IMPLIED")
                Exit Function
            End If
        End If

        If HasValue(forecastStart) Then
            If CDbl(forecastStart) <> CDbl(mustFinishStart) Then
                Core_AddBlockingError dataArr, rowIdx, mapCol, blockingErrors, taskId, _
                    Core_BuildConstraintDiagnosticMessage(dataArr, rowIdx, mapCol, constraintDiagnostics, taskId, startConstraintType, startConstraintDate, finishConstraintType, finishConstraintDate, actualStart, actualFinish, forecastStart, forecastFinish, calcStart, calcFinish, predAllowedStart, predAllowedFinish, allowedStart, allowedFinish, effectiveDuration, mustFinishStart, _
                        "DIAG.CONSTRAINT.FORECAST_START_DIFFERS_MFO_IMPLIED")
                Exit Function
            End If
        End If
    End If

    If HasValue(constraintMustStart) And HasValue(allowedStart) Then
        If CDbl(allowedStart) > CDbl(constraintMustStart) Then
            Core_AddBlockingError dataArr, rowIdx, mapCol, blockingErrors, taskId, _
                Core_BuildConstraintDiagnosticMessage(dataArr, rowIdx, mapCol, constraintDiagnostics, taskId, startConstraintType, startConstraintDate, finishConstraintType, finishConstraintDate, actualStart, actualFinish, forecastStart, forecastFinish, calcStart, calcFinish, predAllowedStart, predAllowedFinish, allowedStart, allowedFinish, effectiveDuration, mustFinishStart, _
                    "DIAG.CONSTRAINT.CALCULATED_START_DIFFERS_MSO")
            Exit Function
        End If
    End If

    If leafProfileEnabled Then
        CoreLeafProfile_AddPhase "ConstraintEvaluation", leafPhaseStart
        leafPhaseStart = CoreLeafProfile_Timestamp()
    End If

    If HasValue(actualStart) Then
        sourceStart = actualStart
    ElseIf HasValue(forecastStart) Then
        sourceStart = forecastStart
    ElseIf HasValue(baselineStart) Then
        sourceStart = baselineStart
    Else
        sourceStart = Empty
    End If
    If leafProfileEnabled Then
        If HasValue(actualStart) Or HasValue(actualFinish) Then
            CoreLeafProfile_Count "PolicyActual"
        ElseIf HasValue(forecastStart) Or HasValue(forecastFinish) Then
            CoreLeafProfile_Count "PolicyForecast"
        ElseIf HasValue(baselineStart) Then
            CoreLeafProfile_Count "PolicyBaseline"
        Else
            CoreLeafProfile_Count "PolicyDependenciesOnly"
        End If
    End If

    If HasValue(constraintMustFinish) Then
        calcStart = mustFinishStart
    ElseIf HasValue(constraintMustStart) Then
        calcStart = constraintMustStart
    ElseIf HasValue(sourceStart) Then
        calcStart = Core_MaxDateIfBoth(sourceStart, allowedStart)
    Else
        calcStart = Empty

        If HasValue(allowedFinish) And HasValue(effectiveDuration) Then
            calcStart = SubtractWorkingDays(allowedFinish, effectiveDuration, calType)
        End If

        If HasValue(allowedStart) Then
            calcStart = Core_MaxDateIfBoth(calcStart, allowedStart)
        End If
    End If

    If Not HasValue(calcStart) Then
        Core_AddBlockingError dataArr, rowIdx, mapCol, blockingErrors, taskId, _
            TextCatalog_Get("CORE.ERROR.START_NOT_COMPUTABLE", TEXT_LANGUAGE_EN), _
            "CORE.ERROR.START_NOT_COMPUTABLE", Nothing, vbNullString, coreDiagnostics, "ROOT", "DATES"
        Exit Function
    End If

    If HasValue(constraintLatestStart) Then
        If CDbl(calcStart) > CDbl(constraintLatestStart) Then
            Core_AddBlockingError dataArr, rowIdx, mapCol, blockingErrors, taskId, _
                Core_BuildConstraintDiagnosticMessage(dataArr, rowIdx, mapCol, constraintDiagnostics, taskId, startConstraintType, startConstraintDate, finishConstraintType, finishConstraintDate, actualStart, actualFinish, forecastStart, forecastFinish, calcStart, calcFinish, predAllowedStart, predAllowedFinish, allowedStart, allowedFinish, effectiveDuration, mustFinishStart, _
                    "DIAG.CONSTRAINT.CALCULATED_START_AFTER_LATEST")
            Exit Function
        End If
    End If

    If HasValue(constraintMustStart) Then
        If CDbl(calcStart) <> CDbl(constraintMustStart) Then
            Core_AddBlockingError dataArr, rowIdx, mapCol, blockingErrors, taskId, _
                Core_BuildConstraintDiagnosticMessage(dataArr, rowIdx, mapCol, constraintDiagnostics, taskId, startConstraintType, startConstraintDate, finishConstraintType, finishConstraintDate, actualStart, actualFinish, forecastStart, forecastFinish, calcStart, calcFinish, predAllowedStart, predAllowedFinish, allowedStart, allowedFinish, effectiveDuration, mustFinishStart, _
                    "DIAG.CONSTRAINT.CALCULATED_START_DIFFERS_MSO")
            Exit Function
        End If
    End If

    If HasValue(actualFinish) Then
        sourceFinish = actualFinish
    ElseIf HasValue(forecastFinish) Then
        sourceFinish = forecastFinish
    Else
        sourceFinish = Empty
    End If

    If HasValue(constraintMustFinish) Then
        calcFinish = constraintMustFinish
    ElseIf HasValue(sourceFinish) Then
        calcFinish = sourceFinish
    Else
        If Not HasValue(effectiveDuration) Then
            Core_AddBlockingError dataArr, rowIdx, mapCol, blockingErrors, taskId, _
                TextCatalog_Get("CORE.ERROR.BASELINE_DURATION_MISSING", TEXT_LANGUAGE_EN), _
                "CORE.ERROR.BASELINE_DURATION_MISSING", Nothing, vbNullString, coreDiagnostics, "ROOT", "BASELINE"
            Exit Function
        End If

        calcFinish = AddWorkingDays(calcStart, effectiveDuration, calType)
    End If

    If HasValue(allowedFinish) Then
        If CDbl(calcFinish) < CDbl(allowedFinish) Then
            calcFinish = allowedFinish

            If Not hasExplicitStart Then
                If HasValue(effectiveDuration) Then
                    calcStart = SubtractWorkingDays(calcFinish, effectiveDuration, calType)

                    If HasValue(allowedStart) Then
                        If CDbl(calcStart) < CDbl(allowedStart) Then calcStart = allowedStart
                    End If
                End If
            End If
        End If
    End If

    If HasValue(constraintLatestFinish) Then
        If CDbl(calcFinish) > CDbl(constraintLatestFinish) Then
            Core_AddBlockingError dataArr, rowIdx, mapCol, blockingErrors, taskId, _
                Core_BuildConstraintDiagnosticMessage(dataArr, rowIdx, mapCol, constraintDiagnostics, taskId, startConstraintType, startConstraintDate, finishConstraintType, finishConstraintDate, actualStart, actualFinish, forecastStart, forecastFinish, calcStart, calcFinish, predAllowedStart, predAllowedFinish, allowedStart, allowedFinish, effectiveDuration, mustFinishStart, _
                    "DIAG.CONSTRAINT.CALCULATED_FINISH_AFTER_LATEST")
            Exit Function
        End If
    End If

    If HasValue(constraintMustFinish) Then
        If CDbl(calcFinish) <> CDbl(constraintMustFinish) Then
            Core_AddBlockingError dataArr, rowIdx, mapCol, blockingErrors, taskId, _
                Core_BuildConstraintDiagnosticMessage(dataArr, rowIdx, mapCol, constraintDiagnostics, taskId, startConstraintType, startConstraintDate, finishConstraintType, finishConstraintDate, actualStart, actualFinish, forecastStart, forecastFinish, calcStart, calcFinish, predAllowedStart, predAllowedFinish, allowedStart, allowedFinish, effectiveDuration, mustFinishStart, _
                    "DIAG.CONSTRAINT.CALCULATED_FINISH_DIFFERS_MFO")
            Exit Function
        End If
    End If

    If CDbl(calcFinish) < CDbl(calcStart) Then
        Core_AddBlockingError dataArr, rowIdx, mapCol, blockingErrors, taskId, _
            TextCatalog_Get("CORE.ERROR.FINISH_BEFORE_START", TEXT_LANGUAGE_EN), _
            "CORE.ERROR.FINISH_BEFORE_START", Nothing, vbNullString, coreDiagnostics, "ROOT", "DATES"
        Exit Function
    End If

    If leafProfileEnabled Then
        CoreLeafProfile_AddPhase "DatePolicyAndFinalization", leafPhaseStart
        leafPhaseStart = CoreLeafProfile_Timestamp()
    End If

    CoreLeaf_ResolveDatesAndConstraints = True

End Function
