Attribute VB_Name = "mod_CalcCoreEngine"
Option Explicit

'=====================================================
' mod_CalcCoreEngine
'=====================================================
' Rôle :
' - futur cœur unique de calcul
' - travaille uniquement en mémoire
' - ne lit / n'écrit aucune feuille
' - ne connaît ni WBS ni Gantt ni TEST
'
' Entrées :
' - dataArr(row, col)
' - mapCol(name) = col index
' - linksBySuccId(succId) = Collection de link()
'
' Sorties :
' - modifie dataArr en place :
'   * Calculated Start
'   * Calculated Finish
'   * Calculated Duration
'   * Error flag
'   * ErrorMsg
'
' Convention :
' - les dépendances sont déjà préparées
' - les summaries sont déjà identifiées
' - les warnings restent en dehors du cœur pour l'instant
'=====================================================

Public Sub Run_Calc_Core( _
    ByRef dataArr As Variant, _
    ByVal mapCol As Object, _
    ByVal linksBySuccId As Object, _
    Optional ByVal recalcScope As Object)

    Dim requiredCols As Variant

    Dim rowById As Object
    Dim parentIds As Object
    Dim directChildrenById As Object
    Dim childrenByPred As Object
    Dim validIds As Object
    Dim loeIds As Object
    Dim indegree As Object
    Dim topoOrder As Collection

    Dim calcStartById As Object
    Dim calcFinishById As Object
    Dim blockingErrors As Object

    Dim currentId As Variant
    Dim rowIdx As Long
    Dim shouldCompute As Boolean
    Dim isPartialMode As Boolean

    requiredCols = Array( _
        "ID", _
        "WBS", _
        "Task Name", _
        "ParentID", _
        "IsSummary", _
        "Task Type", _
        "Actual Start", _
        "Actual Finish", _
        "Forecast Start", _
        "Forecast Finish", _
        "Baseline Start", _
        "Baseline Duration", _
        "Constraint Active", _
        "Start Constraint Type", _
        "Start Constraint Date", _
        "Finish Constraint Type", _
        "Finish Constraint Date", _
        "Calculated Start", _
        "Calculated Finish", _
        "Calculated Duration", _
        "Error flag", _
        "ErrorMsg" _
    )

    Core_RequireColumns mapCol, requiredCols, "Run_Calc_Core"

    isPartialMode = Not (recalcScope Is Nothing)

    Set rowById = Core_BuildRowById(dataArr, mapCol)
    Set parentIds = Core_BuildParentIds(dataArr, mapCol, rowById)
    Set directChildrenById = Core_BuildDirectChildrenById(dataArr, mapCol, rowById)
    Set childrenByPred = Core_BuildChildrenByPred(rowById, linksBySuccId)

    Set loeIds = Core_BuildLOEIds(dataArr, mapCol, rowById)

    Set validIds = Core_BuildValidLeafIds(rowById, parentIds)
    Core_RemoveIdsFromDictionary validIds, loeIds

    Set indegree = Core_BuildIndegree(validIds, linksBySuccId)
    Set topoOrder = Core_TopoSortLeafNetwork(validIds, childrenByPred, indegree)

    Set calcStartById = CreateObject("Scripting.Dictionary")
    Set calcFinishById = CreateObject("Scripting.Dictionary")
    Set blockingErrors = CreateObject("Scripting.Dictionary")

    If isPartialMode Then
        Core_LoadExistingCalcOutputs dataArr, mapCol, rowById, calcStartById, calcFinishById
        Core_ClearCalcOutputs_ForScope dataArr, mapCol, rowById, recalcScope, calcStartById, calcFinishById
    Else
        Core_ClearAllCalcOutputs dataArr, mapCol
    End If

    Core_ValidateLOEAsNonPredecessor dataArr, mapCol, rowById, linksBySuccId, loeIds, blockingErrors

    If Core_HasTopoFailure(topoOrder, validIds) Then
        Core_MarkTopoFailure dataArr, mapCol, validIds, rowById, linksBySuccId, blockingErrors
        GoTo SafeExit
    End If

    For Each currentId In topoOrder

        shouldCompute = False

        If isPartialMode Then
            If recalcScope.Exists(CStr(currentId)) Then
                shouldCompute = True
            End If
        Else
            shouldCompute = True
        End If

        If shouldCompute Then
            Core_ComputeOneLeafTask _
                dataArr, mapCol, CStr(currentId), rowById, linksBySuccId, _
                calcStartById, calcFinishById, blockingErrors
        End If

    Next currentId

    Core_ApplyLOEPostProcess dataArr, mapCol, rowById, linksBySuccId, loeIds, _
                             calcStartById, calcFinishById, blockingErrors

    If blockingErrors.Count > 0 Then
        Core_PropagateBlockingErrors dataArr, mapCol, blockingErrors, childrenByPred, rowById
    End If

    For Each currentId In calcStartById.Keys
        If rowById.Exists(CStr(currentId)) Then
            rowIdx = CLng(rowById(CStr(currentId)))
            Core_SetCalcTriplet dataArr, rowIdx, mapCol, _
                calcStartById(CStr(currentId)), _
                calcFinishById(CStr(currentId))
        End If
    Next currentId

    Core_RollupSummaryDates dataArr, mapCol, rowById, directChildrenById, parentIds

SafeExit:
End Sub


Private Sub Core_LoadExistingCalcOutputs( _
    ByRef dataArr As Variant, _
    ByVal mapCol As Object, _
    ByVal rowById As Object, _
    ByVal calcStartById As Object, _
    ByVal calcFinishById As Object)

    Dim idVal As Variant
    Dim rowIdx As Long
    Dim calcStart As Variant
    Dim calcFinish As Variant

    For Each idVal In rowById.Keys

        rowIdx = CLng(rowById(CStr(idVal)))

        calcStart = Core_GetVal(dataArr, rowIdx, mapCol, "Calculated Start")
        calcFinish = Core_GetVal(dataArr, rowIdx, mapCol, "Calculated Finish")

        If HasValue(calcStart) Then
            calcStartById(CStr(idVal)) = calcStart
        End If

        If HasValue(calcFinish) Then
            calcFinishById(CStr(idVal)) = calcFinish
        End If

    Next idVal

End Sub

Private Sub Core_ComputeOneLeafTask( _
    ByRef dataArr As Variant, _
    ByVal mapCol As Object, _
    ByVal taskId As String, _
    ByVal rowById As Object, _
    ByVal linksBySuccId As Object, _
    ByVal calcStartById As Object, _
    ByVal calcFinishById As Object, _
    ByVal blockingErrors As Object)

    Dim rowIdx As Long

    Dim baselineStart As Variant
    Dim baselineDuration As Variant
    Dim actualStart As Variant
    Dim actualFinish As Variant
    Dim forecastStart As Variant
    Dim forecastFinish As Variant

    Dim predAllowedStart As Variant
    Dim predAllowedFinish As Variant
    Dim constraintAllowedStart As Variant
    Dim constraintLatestStart As Variant
    Dim constraintMustStart As Variant
    Dim constraintAllowedFinish As Variant
    Dim constraintLatestFinish As Variant
    Dim constraintMustFinish As Variant
    Dim allowedStart As Variant
    Dim allowedFinish As Variant
    Dim constraintActive As String
    Dim startConstraintType As String
    Dim finishConstraintType As String
    Dim startConstraintDate As Variant
    Dim finishConstraintDate As Variant
    Dim hasExplicitStart As Boolean

    Dim normalAllowedStart As Variant
    Dim summaryAllowedStart As Variant
    Dim summaryStartBySource As Object

    Dim sourceStart As Variant
    Dim sourceFinish As Variant
    Dim calcStart As Variant
    Dim calcFinish As Variant
    Dim mustFinishStart As Variant

    Dim oneLink As Variant
    Dim predId As String
    Dim linkType As String
    Dim lagVal As Double
    Dim summarySourceId As String

    Dim candidateStart As Variant
    Dim candidateFinish As Variant
    Dim parentKey As Variant

    Dim effectiveDuration As Variant
    Dim taskTypeVal As String

    If Not rowById.Exists(taskId) Then Exit Sub
    rowIdx = CLng(rowById(taskId))

    baselineStart = Core_GetVal(dataArr, rowIdx, mapCol, "Baseline Start")
    baselineDuration = Core_GetVal(dataArr, rowIdx, mapCol, "Baseline Duration")
    actualStart = Core_GetVal(dataArr, rowIdx, mapCol, "Actual Start")
    actualFinish = Core_GetVal(dataArr, rowIdx, mapCol, "Actual Finish")
    forecastStart = Core_GetVal(dataArr, rowIdx, mapCol, "Forecast Start")
    forecastFinish = Core_GetVal(dataArr, rowIdx, mapCol, "Forecast Finish")
    constraintActive = UCase$(Trim$(CStr(Core_GetVal(dataArr, rowIdx, mapCol, "Constraint Active"))))
    startConstraintType = Trim$(CStr(Core_GetVal(dataArr, rowIdx, mapCol, "Start Constraint Type")))
    startConstraintDate = Core_GetVal(dataArr, rowIdx, mapCol, "Start Constraint Date")
    finishConstraintType = Trim$(CStr(Core_GetVal(dataArr, rowIdx, mapCol, "Finish Constraint Type")))
    finishConstraintDate = Core_GetVal(dataArr, rowIdx, mapCol, "Finish Constraint Date")

    effectiveDuration = baselineDuration

    'Milestone rule:
    'A milestone without explicit duration is treated as 1 day by the core.
    'This is a calculation default only; it does not write 1 back to WBS/CALC input fields.
    If Not HasValue(effectiveDuration) Then
        taskTypeVal = Core_NormalizeTaskType(Core_GetVal(dataArr, rowIdx, mapCol, "Task Type"))

        If taskTypeVal = "MILESTONE" Then
            effectiveDuration = 1
        End If
    End If

    predAllowedStart = Empty
    predAllowedFinish = Empty
    constraintAllowedStart = Empty
    constraintLatestStart = Empty
    constraintMustStart = Empty
    constraintAllowedFinish = Empty
    constraintLatestFinish = Empty
    constraintMustFinish = Empty
    mustFinishStart = Empty
    allowedStart = Empty
    allowedFinish = Empty

    normalAllowedStart = Empty
    summaryAllowedStart = Empty
    Set summaryStartBySource = CreateObject("Scripting.Dictionary")

    If Not linksBySuccId Is Nothing Then
        If linksBySuccId.Exists(taskId) Then
            For Each oneLink In linksBySuccId(taskId)

                predId = Core_GetLinkPredId(oneLink)
                linkType = Core_GetLinkType(oneLink)
                lagVal = Core_GetLinkLag(oneLink)
                summarySourceId = Core_GetLinkSummarySourceId(oneLink)

                If predId = "" Then
                    Core_AddBlockingError dataArr, rowIdx, mapCol, blockingErrors, taskId, _
                        "Missing predecessor"
                    Exit Sub
                End If

                If Not rowById.Exists(predId) Then
                    Core_AddBlockingError dataArr, rowIdx, mapCol, blockingErrors, taskId, _
                        "Missing predecessor: ID " & predId
                    Exit Sub
                End If

                If blockingErrors.Exists(predId) Then
                    Core_AddBlockingError dataArr, rowIdx, mapCol, blockingErrors, taskId, _
                        "Blocked by predecessor error: ID " & predId
                    Exit Sub
                End If

                Select Case linkType

                    Case "SS"
                        If calcStartById.Exists(predId) Then
                            candidateStart = CDbl(calcStartById(predId)) + lagVal

                            If summarySourceId <> "" Then
                                If summaryStartBySource.Exists(summarySourceId) Then
                                    summaryStartBySource(summarySourceId) = _
                                        Core_MinDateIfBoth(summaryStartBySource(summarySourceId), candidateStart)
                                Else
                                    summaryStartBySource(summarySourceId) = candidateStart
                                End If
                            Else
                                normalAllowedStart = Core_MaxDateIfBoth(normalAllowedStart, candidateStart)
                            End If
                        End If

                    Case "FF"
                        If calcFinishById.Exists(predId) Then
                            candidateFinish = CDbl(calcFinishById(predId)) + lagVal
                            predAllowedFinish = Core_MaxDateIfBoth(predAllowedFinish, candidateFinish)
                        End If

                    Case "FS", ""
                        If calcFinishById.Exists(predId) Then
                            candidateStart = CDbl(calcFinishById(predId)) + 1 + lagVal
                            normalAllowedStart = Core_MaxDateIfBoth(normalAllowedStart, candidateStart)
                        End If

                    Case Else
                        Core_AddBlockingError dataArr, rowIdx, mapCol, blockingErrors, taskId, _
                            "Unsupported link type: " & linkType
                        Exit Sub

                End Select

            Next oneLink
        End If
    End If

    For Each parentKey In summaryStartBySource.Keys
        summaryAllowedStart = Core_MaxDateIfBoth(summaryAllowedStart, summaryStartBySource(CStr(parentKey)))
    Next parentKey

    predAllowedStart = Core_MaxDateIfBoth(normalAllowedStart, summaryAllowedStart)

    If constraintActive = "YES" Then
        If startConstraintType = "Start No Earlier Than" Then
            constraintAllowedStart = startConstraintDate
        ElseIf startConstraintType = "Start No Later Than" Then
            constraintLatestStart = startConstraintDate
        ElseIf startConstraintType = "Must Start On" Then
            constraintMustStart = startConstraintDate
        ElseIf startConstraintType <> "" Then
            Core_AddBlockingError dataArr, rowIdx, mapCol, blockingErrors, taskId, _
                Core_BuildConstraintCoreMessage(dataArr, rowIdx, mapCol, _
                    "Type de contrainte debut non reconnu", _
                    "Unknown start constraint type")
            Exit Sub
        End If

        If finishConstraintType = "Finish No Earlier Than" Then
            constraintAllowedFinish = finishConstraintDate
        ElseIf finishConstraintType = "Finish No Later Than" Then
            constraintLatestFinish = finishConstraintDate
        ElseIf finishConstraintType = "Must Finish On" Then
            constraintMustFinish = finishConstraintDate
        ElseIf finishConstraintType <> "" Then
            Core_AddBlockingError dataArr, rowIdx, mapCol, blockingErrors, taskId, _
                Core_BuildConstraintCoreMessage(dataArr, rowIdx, mapCol, _
                    "Type de contrainte fin non reconnu", _
                    "Unknown finish constraint type")
            Exit Sub
        End If
    End If

    allowedStart = Core_MaxDateIfBoth(predAllowedStart, constraintAllowedStart)
    allowedFinish = Core_MaxDateIfBoth(predAllowedFinish, constraintAllowedFinish)
    hasExplicitStart = HasValue(actualStart) Or HasValue(forecastStart) Or HasValue(constraintAllowedStart) Or HasValue(constraintMustStart)

    If HasValue(actualStart) And HasValue(predAllowedStart) Then
        If CDbl(actualStart) < CDbl(predAllowedStart) Then
            Core_AddBlockingError dataArr, rowIdx, mapCol, blockingErrors, taskId, _
                "Actual Start violates dependencies"
            Exit Sub
        End If
    End If

    If HasValue(actualStart) And HasValue(constraintAllowedStart) Then
        If CDbl(actualStart) < CDbl(constraintAllowedStart) Then
            Core_AddBlockingError dataArr, rowIdx, mapCol, blockingErrors, taskId, _
                Core_BuildConstraintCoreMessage(dataArr, rowIdx, mapCol, _
                    "Actual Start avant contrainte debut", _
                    "Actual Start is before start constraint")
            Exit Sub
        End If
    End If

    If HasValue(actualStart) And HasValue(constraintLatestStart) Then
        If CDbl(actualStart) > CDbl(constraintLatestStart) Then
            Core_AddBlockingError dataArr, rowIdx, mapCol, blockingErrors, taskId, _
                Core_BuildConstraintCoreMessage(dataArr, rowIdx, mapCol, _
                    "Actual Start apres contrainte debut max", _
                    "Actual Start is after latest start constraint")
            Exit Sub
        End If
    End If

    If HasValue(actualStart) And HasValue(constraintMustStart) Then
        If CDbl(actualStart) <> CDbl(constraintMustStart) Then
            Core_AddBlockingError dataArr, rowIdx, mapCol, blockingErrors, taskId, _
                Core_BuildConstraintCoreMessage(dataArr, rowIdx, mapCol, _
                    "Actual Start different de contrainte Must Start On", _
                    "Actual Start differs from Must Start On constraint")
            Exit Sub
        End If
    End If

    If HasValue(actualFinish) And HasValue(predAllowedFinish) Then
        If CDbl(actualFinish) < CDbl(predAllowedFinish) Then
            Core_AddBlockingError dataArr, rowIdx, mapCol, blockingErrors, taskId, _
                "Actual Finish violates finish constraints"
            Exit Sub
        End If
    End If

    If HasValue(actualFinish) And HasValue(constraintAllowedFinish) Then
        If CDbl(actualFinish) < CDbl(constraintAllowedFinish) Then
            Core_AddBlockingError dataArr, rowIdx, mapCol, blockingErrors, taskId, _
                Core_BuildConstraintCoreMessage(dataArr, rowIdx, mapCol, _
                    "Actual Finish avant contrainte fin", _
                    "Actual Finish is before finish constraint")
            Exit Sub
        End If
    End If

    If HasValue(actualFinish) And HasValue(constraintLatestFinish) Then
        If CDbl(actualFinish) > CDbl(constraintLatestFinish) Then
            Core_AddBlockingError dataArr, rowIdx, mapCol, blockingErrors, taskId, _
                Core_BuildConstraintCoreMessage(dataArr, rowIdx, mapCol, _
                    "Actual Finish apres contrainte fin max", _
                    "Actual Finish is after latest finish constraint")
            Exit Sub
        End If
    End If

    If HasValue(actualFinish) And HasValue(constraintMustFinish) Then
        If CDbl(actualFinish) <> CDbl(constraintMustFinish) Then
            Core_AddBlockingError dataArr, rowIdx, mapCol, blockingErrors, taskId, _
                Core_BuildConstraintCoreMessage(dataArr, rowIdx, mapCol, _
                    "Actual Finish different de contrainte Must Finish On", _
                    "Actual Finish differs from Must Finish On constraint")
            Exit Sub
        End If
    End If

    If HasValue(forecastStart) And HasValue(predAllowedStart) Then
        If CDbl(forecastStart) < CDbl(predAllowedStart) Then
            Core_AddBlockingError dataArr, rowIdx, mapCol, blockingErrors, taskId, _
                "Forecast Start violates dependencies"
            Exit Sub
        End If
    End If

    If HasValue(forecastStart) And HasValue(constraintAllowedStart) Then
        If CDbl(forecastStart) < CDbl(constraintAllowedStart) Then
            Core_AddBlockingError dataArr, rowIdx, mapCol, blockingErrors, taskId, _
                Core_BuildConstraintCoreMessage(dataArr, rowIdx, mapCol, _
                    "Forecast Start avant contrainte debut", _
                    "Forecast Start is before start constraint")
            Exit Sub
        End If
    End If

    If HasValue(forecastStart) And HasValue(constraintLatestStart) Then
        If CDbl(forecastStart) > CDbl(constraintLatestStart) Then
            Core_AddBlockingError dataArr, rowIdx, mapCol, blockingErrors, taskId, _
                Core_BuildConstraintCoreMessage(dataArr, rowIdx, mapCol, _
                    "Forecast Start apres contrainte debut max", _
                    "Forecast Start is after latest start constraint")
            Exit Sub
        End If
    End If

    If HasValue(forecastStart) And HasValue(constraintMustStart) Then
        If CDbl(forecastStart) <> CDbl(constraintMustStart) Then
            Core_AddBlockingError dataArr, rowIdx, mapCol, blockingErrors, taskId, _
                Core_BuildConstraintCoreMessage(dataArr, rowIdx, mapCol, _
                    "Forecast Start different de contrainte Must Start On", _
                    "Forecast Start differs from Must Start On constraint")
            Exit Sub
        End If
    End If

    If HasValue(forecastFinish) And HasValue(predAllowedFinish) Then
        If CDbl(forecastFinish) < CDbl(predAllowedFinish) Then
            Core_AddBlockingError dataArr, rowIdx, mapCol, blockingErrors, taskId, _
                Core_BuildConstraintCoreMessage(dataArr, rowIdx, mapCol, _
                    "Forecast Finish avant contrainte fin amont", _
                    "Forecast Finish is before upstream finish constraint")
            Exit Sub
        End If
    End If

    If HasValue(forecastFinish) And HasValue(constraintAllowedFinish) Then
        If CDbl(forecastFinish) < CDbl(constraintAllowedFinish) Then
            Core_AddBlockingError dataArr, rowIdx, mapCol, blockingErrors, taskId, _
                Core_BuildConstraintCoreMessage(dataArr, rowIdx, mapCol, _
                    "Forecast Finish avant contrainte fin", _
                    "Forecast Finish is before finish constraint")
            Exit Sub
        End If
    End If

    If HasValue(forecastFinish) And HasValue(constraintLatestFinish) Then
        If CDbl(forecastFinish) > CDbl(constraintLatestFinish) Then
            Core_AddBlockingError dataArr, rowIdx, mapCol, blockingErrors, taskId, _
                Core_BuildConstraintCoreMessage(dataArr, rowIdx, mapCol, _
                    "Forecast Finish apres contrainte fin max", _
                    "Forecast Finish is after latest finish constraint")
            Exit Sub
        End If
    End If

    If HasValue(forecastFinish) And HasValue(constraintMustFinish) Then
        If CDbl(forecastFinish) <> CDbl(constraintMustFinish) Then
            Core_AddBlockingError dataArr, rowIdx, mapCol, blockingErrors, taskId, _
                Core_BuildConstraintCoreMessage(dataArr, rowIdx, mapCol, _
                    "Forecast Finish different de contrainte Must Finish On", _
                    "Forecast Finish differs from Must Finish On constraint")
            Exit Sub
        End If
    End If

    If HasValue(constraintMustFinish) Then
        If Not HasValue(effectiveDuration) Then
            Core_AddBlockingError dataArr, rowIdx, mapCol, blockingErrors, taskId, _
                "Baseline Duration missing"
            Exit Sub
        End If

        mustFinishStart = CDbl(constraintMustFinish) - CDbl(effectiveDuration) + 1

        If HasValue(allowedFinish) Then
            If CDbl(allowedFinish) > CDbl(constraintMustFinish) Then
                Core_AddBlockingError dataArr, rowIdx, mapCol, blockingErrors, taskId, _
                    Core_BuildConstraintCoreMessage(dataArr, rowIdx, mapCol, _
                        "Calculated Finish different de contrainte Must Finish On", _
                        "Calculated Finish differs from Must Finish On constraint")
                Exit Sub
            End If
        End If

        If HasValue(allowedStart) Then
            If CDbl(allowedStart) > CDbl(mustFinishStart) Then
                Core_AddBlockingError dataArr, rowIdx, mapCol, blockingErrors, taskId, _
                    Core_BuildConstraintCoreMessage(dataArr, rowIdx, mapCol, _
                        "Calculated Start avant reseau impose par contrainte Must Finish On", _
                        "Calculated Start is before network allowed start due to Must Finish On constraint")
                Exit Sub
            End If
        End If

        If HasValue(constraintMustStart) Then
            If CDbl(constraintMustStart) <> CDbl(mustFinishStart) Then
                Core_AddBlockingError dataArr, rowIdx, mapCol, blockingErrors, taskId, _
                    Core_BuildConstraintCoreMessage(dataArr, rowIdx, mapCol, _
                        "Duree incompatible avec contraintes Must Start On / Must Finish On", _
                        "Duration is incompatible with Must Start On / Must Finish On constraints")
                Exit Sub
            End If
        End If

        If HasValue(actualStart) Then
            If CDbl(actualStart) <> CDbl(mustFinishStart) Then
                Core_AddBlockingError dataArr, rowIdx, mapCol, blockingErrors, taskId, _
                    Core_BuildConstraintCoreMessage(dataArr, rowIdx, mapCol, _
                        "Actual Start different du debut impose par Must Finish On", _
                        "Actual Start differs from start implied by Must Finish On constraint")
                Exit Sub
            End If
        End If

        If HasValue(forecastStart) Then
            If CDbl(forecastStart) <> CDbl(mustFinishStart) Then
                Core_AddBlockingError dataArr, rowIdx, mapCol, blockingErrors, taskId, _
                    Core_BuildConstraintCoreMessage(dataArr, rowIdx, mapCol, _
                        "Forecast Start different du debut impose par Must Finish On", _
                        "Forecast Start differs from start implied by Must Finish On constraint")
                Exit Sub
            End If
        End If
    End If

    If HasValue(constraintMustStart) And HasValue(allowedStart) Then
        If CDbl(allowedStart) > CDbl(constraintMustStart) Then
            Core_AddBlockingError dataArr, rowIdx, mapCol, blockingErrors, taskId, _
                Core_BuildConstraintCoreMessage(dataArr, rowIdx, mapCol, _
                    "Calculated Start different de contrainte Must Start On", _
                    "Calculated Start differs from Must Start On constraint")
            Exit Sub
        End If
    End If

    sourceStart = Core_GetSourceStart(dataArr, rowIdx, mapCol)

    If HasValue(constraintMustFinish) Then
        calcStart = mustFinishStart
    ElseIf HasValue(constraintMustStart) Then
        calcStart = constraintMustStart
    ElseIf HasValue(sourceStart) Then
        calcStart = Core_MaxDateIfBoth(sourceStart, allowedStart)
    Else
        calcStart = Empty

        If HasValue(allowedFinish) And HasValue(effectiveDuration) Then
            calcStart = CDbl(allowedFinish) - CDbl(effectiveDuration) + 1
        End If

        If HasValue(allowedStart) Then
            calcStart = Core_MaxDateIfBoth(calcStart, allowedStart)
        End If
    End If

    If Not HasValue(calcStart) Then
        Core_AddBlockingError dataArr, rowIdx, mapCol, blockingErrors, taskId, _
            "Start date not computable"
        Exit Sub
    End If

    If HasValue(constraintLatestStart) Then
        If CDbl(calcStart) > CDbl(constraintLatestStart) Then
            Core_AddBlockingError dataArr, rowIdx, mapCol, blockingErrors, taskId, _
                Core_BuildConstraintCoreMessage(dataArr, rowIdx, mapCol, _
                    "Calculated Start apres contrainte debut max", _
                    "Calculated Start is after latest start constraint")
            Exit Sub
        End If
    End If

    If HasValue(constraintMustStart) Then
        If CDbl(calcStart) <> CDbl(constraintMustStart) Then
            Core_AddBlockingError dataArr, rowIdx, mapCol, blockingErrors, taskId, _
                Core_BuildConstraintCoreMessage(dataArr, rowIdx, mapCol, _
                    "Calculated Start different de contrainte Must Start On", _
                    "Calculated Start differs from Must Start On constraint")
            Exit Sub
        End If
    End If

    sourceFinish = Core_GetSourceFinish(dataArr, rowIdx, mapCol)

    If HasValue(constraintMustFinish) Then
        calcFinish = constraintMustFinish
    ElseIf HasValue(sourceFinish) Then
        calcFinish = sourceFinish
    Else
        If Not HasValue(effectiveDuration) Then
            Core_AddBlockingError dataArr, rowIdx, mapCol, blockingErrors, taskId, _
                "Baseline Duration missing"
            Exit Sub
        End If

        calcFinish = CDbl(calcStart) + CDbl(effectiveDuration) - 1
    End If

    If HasValue(allowedFinish) Then
        If CDbl(calcFinish) < CDbl(allowedFinish) Then
            calcFinish = allowedFinish

            If Not hasExplicitStart Then
                If HasValue(effectiveDuration) Then
                    calcStart = CDbl(calcFinish) - CDbl(effectiveDuration) + 1

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
                Core_BuildConstraintCoreMessage(dataArr, rowIdx, mapCol, _
                    "Calculated Finish apres contrainte fin max", _
                    "Calculated Finish is after latest finish constraint")
            Exit Sub
        End If
    End If

    If HasValue(constraintMustFinish) Then
        If CDbl(calcFinish) <> CDbl(constraintMustFinish) Then
            Core_AddBlockingError dataArr, rowIdx, mapCol, blockingErrors, taskId, _
                Core_BuildConstraintCoreMessage(dataArr, rowIdx, mapCol, _
                    "Calculated Finish different de contrainte Must Finish On", _
                    "Calculated Finish differs from Must Finish On constraint")
            Exit Sub
        End If
    End If

    If CDbl(calcFinish) < CDbl(calcStart) Then
        Core_AddBlockingError dataArr, rowIdx, mapCol, blockingErrors, taskId, _
            "Finish before start"
        Exit Sub
    End If

    calcStartById(taskId) = calcStart
    calcFinishById(taskId) = calcFinish

End Sub

Private Function Core_BuildConstraintCoreMessage( _
    ByRef dataArr As Variant, _
    ByVal rowIdx As Long, _
    ByVal mapCol As Object, _
    ByVal frText As String, _
    ByVal enText As String) As String

    Dim idVal As String
    Dim wbsVal As String
    Dim taskName As String

    idVal = Trim$(CStr(Core_GetVal(dataArr, rowIdx, mapCol, "ID")))
    wbsVal = Trim$(CStr(Core_GetVal(dataArr, rowIdx, mapCol, "WBS")))
    taskName = Trim$(CStr(Core_GetVal(dataArr, rowIdx, mapCol, "Task Name")))

    Core_BuildConstraintCoreMessage = _
        "FR:" & vbCrLf & _
        frText & vbCrLf & vbCrLf & _
        "ID : " & idVal & vbCrLf & _
        "WBS : " & wbsVal & vbCrLf & _
        "Task : " & taskName & vbCrLf & vbCrLf & _
        "EN:" & vbCrLf & _
        enText & vbCrLf & vbCrLf & _
        "ID: " & idVal & vbCrLf & _
        "WBS: " & wbsVal & vbCrLf & _
        "Task: " & taskName

End Function


Private Sub Core_AddBlockingError( _
    ByRef dataArr As Variant, _
    ByVal rowIdx As Long, _
    ByVal mapCol As Object, _
    ByVal blockingErrors As Object, _
    ByVal taskId As String, _
    ByVal errText As String)

    blockingErrors(CStr(taskId)) = True
    Core_AddErrorMessage_Row dataArr, rowIdx, mapCol, errText, True

End Sub

Private Sub Core_MarkTopoFailure( _
    ByRef dataArr As Variant, _
    ByVal mapCol As Object, _
    ByVal validIds As Object, _
    ByVal rowById As Object, _
    ByVal linksBySuccId As Object, _
    ByVal blockingErrors As Object)

    Dim taskId As Variant
    Dim rowIdx As Long
    Dim cycleMessage As String

    cycleMessage = Core_BuildTopoCycleDiagnosticMessage(dataArr, mapCol, validIds, rowById, linksBySuccId)
    If Trim$(cycleMessage) = "" Then cycleMessage = "Cycle detected"

    For Each taskId In validIds.Keys
        If rowById.Exists(CStr(taskId)) Then
            rowIdx = CLng(rowById(CStr(taskId)))
            blockingErrors(CStr(taskId)) = True
            Core_AddErrorMessage_Row dataArr, rowIdx, mapCol, cycleMessage, True
        End If
    Next taskId

End Sub

Private Function Core_BuildTopoCycleDiagnosticMessage( _
    ByRef dataArr As Variant, _
    ByVal mapCol As Object, _
    ByVal validIds As Object, _
    ByVal rowById As Object, _
    ByVal linksBySuccId As Object) As String

    Dim cycleEdges As Collection
    Dim cycleText As String

    On Error GoTo SafeFallback

    Set cycleEdges = Core_FindFirstTopoCycleEdges(validIds, linksBySuccId)

    If cycleEdges Is Nothing Then GoTo SafeFallback
    If cycleEdges.Count = 0 Then GoTo SafeFallback

    cycleText = Core_FormatTopoCycleEdges(dataArr, mapCol, rowById, cycleEdges)
    If Trim$(cycleText) = "" Then GoTo SafeFallback

    Core_BuildTopoCycleDiagnosticMessage = _
        "Cycle detected" & vbCrLf & _
        BiMsg( _
            "Cycle logique détecté :" & vbCrLf & vbCrLf & cycleText & vbCrLf & vbCrLf & _
            "-> corriger un des liens de cette boucle dans Predecessors WBS.", _
            "Logic cycle detected:" & vbCrLf & vbCrLf & cycleText & vbCrLf & vbCrLf & _
            "-> fix one of the links in this loop in Predecessors WBS.")
    Exit Function

SafeFallback:
    Core_BuildTopoCycleDiagnosticMessage = "Cycle detected"

End Function

Private Function Core_FindFirstTopoCycleEdges( _
    ByVal validIds As Object, _
    ByVal linksBySuccId As Object) As Collection

    Dim childrenByPred As Object
    Dim state As Object
    Dim stackIndex As Object
    Dim stackIds As Collection
    Dim stackEdges As Collection
    Dim cycleEdges As Collection
    Dim idKey As Variant
    Dim succId As Variant
    Dim oneLink As Variant
    Dim predId As String
    Dim edgeInfo As Object

    Set childrenByPred = CreateObject("Scripting.Dictionary")
    Set state = CreateObject("Scripting.Dictionary")
    Set stackIndex = CreateObject("Scripting.Dictionary")
    Set stackIds = New Collection
    Set stackEdges = New Collection
    Set cycleEdges = New Collection

    If validIds Is Nothing Then
        Set Core_FindFirstTopoCycleEdges = cycleEdges
        Exit Function
    End If

    For Each idKey In validIds.Keys
        Set childrenByPred(CStr(idKey)) = New Collection
        state(CStr(idKey)) = 0
    Next idKey

    If Not linksBySuccId Is Nothing Then
        For Each succId In linksBySuccId.Keys
            If validIds.Exists(CStr(succId)) Then
                For Each oneLink In linksBySuccId(CStr(succId))
                    predId = Core_GetLinkPredId(oneLink)
                    If predId <> "" Then
                        If validIds.Exists(predId) Then
                            Set edgeInfo = CreateObject("Scripting.Dictionary")
                            edgeInfo("FromId") = predId
                            edgeInfo("ToId") = CStr(succId)
                            edgeInfo("LinkType") = Core_GetLinkType(oneLink)
                            edgeInfo("Lag") = Core_GetLinkLag(oneLink)
                            childrenByPred(predId).Add edgeInfo
                        End If
                    End If
                Next oneLink
            End If
        Next succId
    End If

    For Each idKey In validIds.Keys
        If CLng(state(CStr(idKey))) = 0 Then
            Core_DFS_FindTopoCycleEdges _
                CStr(idKey), childrenByPred, state, stackIndex, stackIds, stackEdges, cycleEdges
            If cycleEdges.Count > 0 Then Exit For
        End If
    Next idKey

    Set Core_FindFirstTopoCycleEdges = cycleEdges

End Function

Private Sub Core_DFS_FindTopoCycleEdges( _
    ByVal currentId As String, _
    ByVal childrenByPred As Object, _
    ByVal state As Object, _
    ByVal stackIndex As Object, _
    ByVal stackIds As Collection, _
    ByVal stackEdges As Collection, _
    ByVal cycleEdges As Collection)

    Dim edgeInfo As Variant
    Dim childId As String
    Dim startIdx As Long
    Dim i As Long

    If cycleEdges.Count > 0 Then Exit Sub

    state(currentId) = 1
    stackIds.Add currentId
    stackIndex(currentId) = stackIds.Count

    If childrenByPred.Exists(currentId) Then
        For Each edgeInfo In childrenByPred(currentId)
            If cycleEdges.Count > 0 Then Exit For

            childId = CStr(edgeInfo("ToId"))
            If Not state.Exists(childId) Then state(childId) = 0

            If CLng(state(childId)) = 0 Then
                stackEdges.Add edgeInfo
                Core_DFS_FindTopoCycleEdges _
                    childId, childrenByPred, state, stackIndex, stackIds, stackEdges, cycleEdges
                If cycleEdges.Count > 0 Then Exit For
                If stackEdges.Count > 0 Then stackEdges.Remove stackEdges.Count

            ElseIf CLng(state(childId)) = 1 Then
                If stackIndex.Exists(childId) Then
                    startIdx = CLng(stackIndex(childId))
                    For i = startIdx To stackEdges.Count
                        cycleEdges.Add stackEdges(i)
                    Next i
                    cycleEdges.Add edgeInfo
                    Exit For
                End If
            End If
        Next edgeInfo
    End If

    state(currentId) = 2

    If stackIds.Count > 0 Then
        If CStr(stackIds(stackIds.Count)) = currentId Then
            stackIds.Remove stackIds.Count
        End If
    End If

    If stackIndex.Exists(currentId) Then stackIndex.Remove currentId

End Sub

Private Function Core_FormatTopoCycleEdges( _
    ByRef dataArr As Variant, _
    ByVal mapCol As Object, _
    ByVal rowById As Object, _
    ByVal cycleEdges As Collection) As String

    Dim result As String
    Dim edgeInfo As Variant
    Dim fromId As String
    Dim toId As String
    Dim linkType As String
    Dim lagVal As Double

    result = ""

    For Each edgeInfo In cycleEdges
        fromId = CStr(edgeInfo("FromId"))
        toId = CStr(edgeInfo("ToId"))
        linkType = UCase$(Trim$(CStr(edgeInfo("LinkType"))))
        lagVal = CDbl(edgeInfo("Lag"))

        If result <> "" Then result = result & vbCrLf & vbCrLf

        result = result & _
            Core_CycleTaskLabel(dataArr, mapCol, rowById, fromId) & vbCrLf & _
            "--" & linkType & Core_FormatCycleLag(lagVal) & "-->" & vbCrLf & _
            Core_CycleTaskLabel(dataArr, mapCol, rowById, toId)
    Next edgeInfo

    Core_FormatTopoCycleEdges = result

End Function

Private Function Core_CycleTaskLabel( _
    ByRef dataArr As Variant, _
    ByVal mapCol As Object, _
    ByVal rowById As Object, _
    ByVal taskId As String) As String

    Dim rowIdx As Long
    Dim wbsVal As String
    Dim taskNameVal As String

    If rowById Is Nothing Then GoTo Fallback
    If Not rowById.Exists(taskId) Then GoTo Fallback

    rowIdx = CLng(rowById(taskId))
    wbsVal = Trim$(CStr(Core_GetVal(dataArr, rowIdx, mapCol, "WBS")))
    taskNameVal = Trim$(CStr(Core_GetVal(dataArr, rowIdx, mapCol, "Task Name")))

    If wbsVal <> "" And taskNameVal <> "" Then
        Core_CycleTaskLabel = wbsVal & " " & taskNameVal
    ElseIf wbsVal <> "" Then
        Core_CycleTaskLabel = wbsVal
    ElseIf taskNameVal <> "" Then
        Core_CycleTaskLabel = taskNameVal
    Else
        Core_CycleTaskLabel = "ID " & taskId
    End If
    Exit Function

Fallback:
    Core_CycleTaskLabel = "ID " & taskId

End Function

Private Function Core_FormatCycleLag(ByVal lagVal As Double) As String

    Dim lagText As String

    If Abs(lagVal) < 0.0000001 Then
        Core_FormatCycleLag = "+0"
        Exit Function
    End If

    lagText = Replace$(Format$(Abs(lagVal), "0.##"), ",", ".")

    If lagVal > 0 Then
        Core_FormatCycleLag = "+" & lagText
    Else
        Core_FormatCycleLag = "-" & lagText
    End If

End Function

Private Sub Core_PropagateBlockingErrors( _
    ByRef dataArr As Variant, _
    ByVal mapCol As Object, _
    ByVal blockingErrors As Object, _
    ByVal childrenByPred As Object, _
    ByVal rowById As Object)

    Dim allErrorIds As Object
    Dim taskId As Variant
    Dim rowIdx As Long

    Set allErrorIds = CreateObject("Scripting.Dictionary")

    For Each taskId In blockingErrors.Keys
        allErrorIds(CStr(taskId)) = True
    Next taskId

    For Each taskId In blockingErrors.Keys
        Core_PropagateErrorToChildren CStr(taskId), childrenByPred, allErrorIds
    Next taskId

    For Each taskId In allErrorIds.Keys
        If rowById.Exists(CStr(taskId)) Then
            rowIdx = CLng(rowById(CStr(taskId)))
            Core_SetErrorFlag_Row dataArr, rowIdx, mapCol, True

            If Not blockingErrors.Exists(CStr(taskId)) Then
                Core_AddErrorMessage_Row dataArr, rowIdx, mapCol, _
                    "Blocked by predecessor chain", True
            End If
        End If
    Next taskId

End Sub

Private Sub Core_ClearCalcOutputs_ForScope( _
    ByRef dataArr As Variant, _
    ByVal mapCol As Object, _
    ByVal rowById As Object, _
    ByVal recalcScope As Object, _
    Optional ByVal calcStartById As Object, _
    Optional ByVal calcFinishById As Object)

    Dim idKey As Variant
    Dim taskId As String
    Dim rowIdx As Long

    If recalcScope Is Nothing Then Exit Sub
    If rowById Is Nothing Then Exit Sub

    For Each idKey In recalcScope.Keys

        taskId = CStr(idKey)

        If rowById.Exists(taskId) Then
            rowIdx = CLng(rowById(taskId))
            Core_ClearCalcOutputs_Row dataArr, rowIdx, mapCol
        End If

        If Not calcStartById Is Nothing Then
            If calcStartById.Exists(taskId) Then calcStartById.Remove taskId
        End If

        If Not calcFinishById Is Nothing Then
            If calcFinishById.Exists(taskId) Then calcFinishById.Remove taskId
        End If

    Next idKey

End Sub

Private Function Core_BuildLOEIds( _
    ByRef dataArr As Variant, _
    ByVal mapCol As Object, _
    ByVal rowById As Object) As Object

    Dim d As Object
    Dim idVal As Variant
    Dim rowIdx As Long
    Dim taskType As String

    Set d = CreateObject("Scripting.Dictionary")

    If rowById Is Nothing Then
        Set Core_BuildLOEIds = d
        Exit Function
    End If

    For Each idVal In rowById.Keys

        rowIdx = CLng(rowById(CStr(idVal)))
        taskType = Core_NormalizeTaskType(Core_GetVal(dataArr, rowIdx, mapCol, "Task Type"))

        If taskType = "LEVEL OF EFFORT" Then
            d(CStr(idVal)) = True
        End If

    Next idVal

    Set Core_BuildLOEIds = d

End Function

Private Function Core_NormalizeTaskType(ByVal rawValue As Variant) As String

    Dim s As String

    s = UCase$(Trim$(CStr(rawValue)))

    Select Case s

        Case "", "TASK", "STANDARD", "NORMAL"
            Core_NormalizeTaskType = "TASK"

        Case "MILESTONE", "MS", "JALON"
            Core_NormalizeTaskType = "MILESTONE"

        Case "LEVEL OF EFFORT", "LOE", "LEVEL-OF-EFFORT", "LEVEL_OF_EFFORT"
            Core_NormalizeTaskType = "LEVEL OF EFFORT"

        Case Else
            Core_NormalizeTaskType = s

    End Select

End Function

Private Sub Core_RemoveIdsFromDictionary( _
    ByVal targetDict As Object, _
    ByVal idsToRemove As Object)

    Dim idVal As Variant

    If targetDict Is Nothing Then Exit Sub
    If idsToRemove Is Nothing Then Exit Sub

    For Each idVal In idsToRemove.Keys
        If targetDict.Exists(CStr(idVal)) Then
            targetDict.Remove CStr(idVal)
        End If
    Next idVal

End Sub

Private Sub Core_ValidateLOEAsNonPredecessor( _
    ByRef dataArr As Variant, _
    ByVal mapCol As Object, _
    ByVal rowById As Object, _
    ByVal linksBySuccId As Object, _
    ByVal loeIds As Object, _
    ByVal blockingErrors As Object)

    Dim succId As Variant
    Dim oneLink As Variant
    Dim predId As String
    Dim loeRow As Long
    Dim succRow As Long

    If linksBySuccId Is Nothing Then Exit Sub
    If loeIds Is Nothing Then Exit Sub
    If blockingErrors Is Nothing Then Exit Sub

    For Each succId In linksBySuccId.Keys

        If linksBySuccId.Exists(CStr(succId)) Then

            For Each oneLink In linksBySuccId(CStr(succId))

                predId = Core_GetLinkPredId(oneLink)

                If predId <> "" Then
                    If loeIds.Exists(predId) Then

                        If rowById.Exists(predId) Then
                            loeRow = CLng(rowById(predId))
                            Core_AddBlockingError dataArr, loeRow, mapCol, blockingErrors, predId, _
                                "LOE cannot be used as predecessor"
                        End If

                        If rowById.Exists(CStr(succId)) Then
                            succRow = CLng(rowById(CStr(succId)))
                            Core_AddBlockingError dataArr, succRow, mapCol, blockingErrors, CStr(succId), _
                                "Blocked by invalid LOE predecessor: ID " & predId
                        End If

                    End If
                End If

            Next oneLink

        End If

    Next succId

End Sub

Private Sub Core_ApplyLOEPostProcess( _
    ByRef dataArr As Variant, _
    ByVal mapCol As Object, _
    ByVal rowById As Object, _
    ByVal linksBySuccId As Object, _
    ByVal loeIds As Object, _
    ByVal calcStartById As Object, _
    ByVal calcFinishById As Object, _
    ByVal blockingErrors As Object)

    Dim loeId As Variant
    Dim rowIdx As Long

    Dim ssCount As Long
    Dim ffCount As Long
    Dim invalidCount As Long

    Dim oneLink As Variant
    Dim linkType As String
    Dim predId As String
    Dim lagVal As Double
    Dim summarySourceId As String

    Dim loeStart As Variant
    Dim loeFinish As Variant
    Dim normalSSStart As Variant
    Dim summarySSStart As Variant
    Dim summarySSStartBySource As Object
    Dim parentKey As Variant

    Dim candidateStart As Variant
    Dim candidateFinish As Variant

    If loeIds Is Nothing Then Exit Sub
    If loeIds.Count = 0 Then Exit Sub

    For Each loeId In loeIds.Keys

        If Not rowById.Exists(CStr(loeId)) Then GoTo NextLOE
        rowIdx = CLng(rowById(CStr(loeId)))

        ssCount = 0
        ffCount = 0
        invalidCount = 0

        loeStart = Empty
        loeFinish = Empty
        normalSSStart = Empty
        summarySSStart = Empty
        Set summarySSStartBySource = CreateObject("Scripting.Dictionary")

        If Not linksBySuccId Is Nothing Then
            If linksBySuccId.Exists(CStr(loeId)) Then

                For Each oneLink In linksBySuccId(CStr(loeId))

                    predId = Core_GetLinkPredId(oneLink)
                    linkType = Core_GetLinkType(oneLink)
                    lagVal = Core_GetLinkLag(oneLink)
                    summarySourceId = Core_GetLinkSummarySourceId(oneLink)

                    Select Case linkType

                        Case "SS"
                            ssCount = ssCount + 1

                            If predId = "" Then
                                Core_AddBlockingError dataArr, rowIdx, mapCol, blockingErrors, CStr(loeId), _
                                    "LOE SS predecessor missing"
                                GoTo NextLOE
                            End If

                            If Not rowById.Exists(predId) Then
                                Core_AddBlockingError dataArr, rowIdx, mapCol, blockingErrors, CStr(loeId), _
                                    "LOE SS predecessor not found: ID " & predId
                                GoTo NextLOE
                            End If

                            If blockingErrors.Exists(predId) Then
                                Core_AddBlockingError dataArr, rowIdx, mapCol, blockingErrors, CStr(loeId), _
                                    "LOE blocked by SS predecessor error: ID " & predId
                                GoTo NextLOE
                            End If

                            If Not calcStartById.Exists(predId) Then
                                Core_AddBlockingError dataArr, rowIdx, mapCol, blockingErrors, CStr(loeId), _
                                    "LOE SS predecessor start not available: ID " & predId
                                GoTo NextLOE
                            End If

                            candidateStart = CDbl(calcStartById(predId)) + lagVal

                            If summarySourceId <> "" Then
                                If summarySSStartBySource.Exists(summarySourceId) Then
                                    summarySSStartBySource(summarySourceId) = _
                                        Core_MinDateIfBoth(summarySSStartBySource(summarySourceId), candidateStart)
                                Else
                                    summarySSStartBySource(summarySourceId) = candidateStart
                                End If
                            Else
                                normalSSStart = Core_MaxDateIfBoth(normalSSStart, candidateStart)
                            End If

                        Case "FF"
                            ffCount = ffCount + 1

                            If predId = "" Then
                                Core_AddBlockingError dataArr, rowIdx, mapCol, blockingErrors, CStr(loeId), _
                                    "LOE FF predecessor missing"
                                GoTo NextLOE
                            End If

                            If Not rowById.Exists(predId) Then
                                Core_AddBlockingError dataArr, rowIdx, mapCol, blockingErrors, CStr(loeId), _
                                    "LOE FF predecessor not found: ID " & predId
                                GoTo NextLOE
                            End If

                            If blockingErrors.Exists(predId) Then
                                Core_AddBlockingError dataArr, rowIdx, mapCol, blockingErrors, CStr(loeId), _
                                    "LOE blocked by FF predecessor error: ID " & predId
                                GoTo NextLOE
                            End If

                            If Not calcFinishById.Exists(predId) Then
                                Core_AddBlockingError dataArr, rowIdx, mapCol, blockingErrors, CStr(loeId), _
                                    "LOE FF predecessor finish not available: ID " & predId
                                GoTo NextLOE
                            End If

                            candidateFinish = CDbl(calcFinishById(predId)) + lagVal
                            loeFinish = Core_MaxDateIfBoth(loeFinish, candidateFinish)

                        Case Else
                            invalidCount = invalidCount + 1

                    End Select

                Next oneLink

            End If
        End If

        If ssCount = 0 Then
            Core_AddBlockingError dataArr, rowIdx, mapCol, blockingErrors, CStr(loeId), _
                "LOE must have at least one SS predecessor"
            GoTo NextLOE
        End If

        If ffCount = 0 Then
            Core_AddBlockingError dataArr, rowIdx, mapCol, blockingErrors, CStr(loeId), _
                "LOE must have at least one FF predecessor"
            GoTo NextLOE
        End If

        If invalidCount > 0 Then
            Core_AddBlockingError dataArr, rowIdx, mapCol, blockingErrors, CStr(loeId), _
                "LOE only supports SS and FF predecessors"
            GoTo NextLOE
        End If

        For Each parentKey In summarySSStartBySource.Keys
            summarySSStart = Core_MaxDateIfBoth(summarySSStart, summarySSStartBySource(CStr(parentKey)))
        Next parentKey

        loeStart = Core_MaxDateIfBoth(normalSSStart, summarySSStart)

        If Not HasValue(loeStart) Then
            Core_AddBlockingError dataArr, rowIdx, mapCol, blockingErrors, CStr(loeId), _
                "LOE start not computable"
            GoTo NextLOE
        End If

        If Not HasValue(loeFinish) Then
            Core_AddBlockingError dataArr, rowIdx, mapCol, blockingErrors, CStr(loeId), _
                "LOE finish not computable"
            GoTo NextLOE
        End If

        If CDbl(loeFinish) < CDbl(loeStart) Then
            Core_AddBlockingError dataArr, rowIdx, mapCol, blockingErrors, CStr(loeId), _
                "LOE finish before start"
            GoTo NextLOE
        End If

        calcStartById(CStr(loeId)) = loeStart
        calcFinishById(CStr(loeId)) = loeFinish

        If mapCol.Exists("Driving Logic") Then
            dataArr(rowIdx, mapCol("Driving Logic")) = "LOE"
        End If

        If mapCol.Exists("Critical Path") Then
            dataArr(rowIdx, mapCol("Critical Path")) = vbNullString
        End If

        If mapCol.Exists("Total Float") Then
            dataArr(rowIdx, mapCol("Total Float")) = vbNullString
        End If

        If mapCol.Exists("Free Float") Then
            dataArr(rowIdx, mapCol("Free Float")) = vbNullString
        End If

        If mapCol.Exists("Critical Path REX") Then
            dataArr(rowIdx, mapCol("Critical Path REX")) = vbNullString
        End If

        If mapCol.Exists("Total Float REX") Then
            dataArr(rowIdx, mapCol("Total Float REX")) = vbNullString
        End If

        If mapCol.Exists("Free Float REX") Then
            dataArr(rowIdx, mapCol("Free Float REX")) = vbNullString
        End If

NextLOE:
    Next loeId

End Sub



