Attribute VB_Name = "mod_CoreLeafCalculator"
Option Explicit

Private Const CORE_INPUT_CAL As Long = 1
Private Const CORE_INPUT_BASELINE_START As Long = 2
Private Const CORE_INPUT_BASELINE_DURATION As Long = 3
Private Const CORE_INPUT_ACTUAL_START As Long = 4
Private Const CORE_INPUT_ACTUAL_FINISH As Long = 5
Private Const CORE_INPUT_FORECAST_START As Long = 6
Private Const CORE_INPUT_FORECAST_FINISH As Long = 7
Private Const CORE_INPUT_CONSTRAINT_ACTIVE As Long = 8
Private Const CORE_INPUT_START_CONSTRAINT_TYPE As Long = 9
Private Const CORE_INPUT_START_CONSTRAINT_DATE As Long = 10
Private Const CORE_INPUT_FINISH_CONSTRAINT_TYPE As Long = 11
Private Const CORE_INPUT_FINISH_CONSTRAINT_DATE As Long = 12
Private Const CORE_INPUT_TASK_TYPE As Long = 13

'------------------------------------------------------------------------------
Public Sub Core_ComputeOneLeafTask( _
    ByRef dataArr As Variant, _
    ByVal mapCol As Object, _
    ByVal taskId As String, _
    ByVal rowById As Object, _
    ByVal linksBySuccId As Object, _
    ByVal calcStartById As Object, _
    ByVal calcFinishById As Object, _
    ByVal blockingErrors As Object, _
    Optional ByVal dependencyDiagnostics As Object, Optional ByVal constraintDiagnostics As Object, _
    Optional ByVal cascadeDiagnostics As Object, Optional ByVal coreDiagnostics As Object, _
    Optional ByVal useIndexedCore As Boolean = False, Optional ByVal nodeIndex As Long = 0, _
    Optional ByRef rowsByNode As Variant, _
    Optional ByRef corePredOffsets As Variant, _
    Optional ByRef corePredNodes As Variant, _
    Optional ByRef corePredIds As Variant, _
    Optional ByRef corePredTypes As Variant, _
    Optional ByRef corePredLags As Variant, Optional ByRef corePredSummarySourceIds As Variant, _
    Optional ByRef indexedInputCols As Variant, _
    Optional ByRef calcStartByNode As Variant, _
    Optional ByRef calcFinishByNode As Variant, _
    Optional ByRef hasCalcByNode As Variant, _
    Optional ByRef blockedByNode As Variant)

    Dim perfScope As clsPerfScope

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
    Dim summaryStartDiagBySource As Object
    Dim predDiagPredId As String
    Dim predDiagLinkType As String
    Dim predDiagLag As Double
    Dim predDiagCandidateDate As Variant
    Dim predDiagPredecessorDate As Variant
    Dim candidatePredecessorDate As Variant
    Dim predDiagPredecessorDateKind As String
    Dim predDiagSummarySourceId As String
    Dim summaryDiagData As Variant

    Dim sourceStart As Variant
    Dim sourceFinish As Variant
    Dim calcStart As Variant
    Dim calcFinish As Variant
    Dim mustFinishStart As Variant

    Dim oneLink As Variant
    Dim taskLinks As Object
    Dim linkPosition As Long
    Dim linkStart As Long
    Dim linkEnd As Long
    Dim predNodeIndex As Long
    Dim predExists As Boolean
    Dim predBlocked As Boolean
    Dim predHasStart As Boolean
    Dim predHasFinish As Boolean
    Dim predId As String
    Dim linkType As String
    Dim lagVal As Double
    Dim summarySourceId As String

    Dim candidateStart As Variant
    Dim candidateFinish As Variant
    Dim parentKey As Variant

    Dim effectiveDuration As Variant
    Dim taskTypeVal As String
    Dim calType As String
    Dim leafProfileEnabled As Boolean
    Dim leafPhaseStart As Double

    Set perfScope = Profiler_BeginScope("Core_ComputeOneLeafTask", "Core Leaf")
    leafProfileEnabled = CoreLeafProfile_IsEnabled()
    If leafProfileEnabled Then leafPhaseStart = CoreLeafProfile_Timestamp()

    If useIndexedCore Then
        If nodeIndex <= 0 Then Exit Sub
        rowIdx = CLng(rowsByNode(nodeIndex))
    Else
        If Not rowById.Exists(taskId) Then Exit Sub
        rowIdx = CLng(rowById(taskId))
    End If

    If useIndexedCore Then
        calType = NormalizeCalendarType(dataArr(rowIdx, indexedInputCols(CORE_INPUT_CAL)))
        baselineStart = dataArr(rowIdx, indexedInputCols(CORE_INPUT_BASELINE_START))
        baselineDuration = dataArr(rowIdx, indexedInputCols(CORE_INPUT_BASELINE_DURATION))
        actualStart = dataArr(rowIdx, indexedInputCols(CORE_INPUT_ACTUAL_START))
        actualFinish = dataArr(rowIdx, indexedInputCols(CORE_INPUT_ACTUAL_FINISH))
        forecastStart = dataArr(rowIdx, indexedInputCols(CORE_INPUT_FORECAST_START))
        forecastFinish = dataArr(rowIdx, indexedInputCols(CORE_INPUT_FORECAST_FINISH))
        constraintActive = UCase$(Trim$(CStr(dataArr(rowIdx, indexedInputCols(CORE_INPUT_CONSTRAINT_ACTIVE)))))
        startConstraintType = Trim$(CStr(dataArr(rowIdx, indexedInputCols(CORE_INPUT_START_CONSTRAINT_TYPE))))
        startConstraintDate = dataArr(rowIdx, indexedInputCols(CORE_INPUT_START_CONSTRAINT_DATE))
        finishConstraintType = Trim$(CStr(dataArr(rowIdx, indexedInputCols(CORE_INPUT_FINISH_CONSTRAINT_TYPE))))
        finishConstraintDate = dataArr(rowIdx, indexedInputCols(CORE_INPUT_FINISH_CONSTRAINT_DATE))
    Else
        calType = NormalizeCalendarType(Core_GetVal(dataArr, rowIdx, mapCol, "Cal"))
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
    End If

    effectiveDuration = baselineDuration

    'Milestone rule:
    'A milestone without explicit duration is treated as 1 day by the core.
    'This is a calculation default only; it does not write 1 back to WBS/CALC input fields.
    If Not HasValue(effectiveDuration) Then
        If useIndexedCore Then
            taskTypeVal = Core_NormalizeTaskType(dataArr(rowIdx, indexedInputCols(CORE_INPUT_TASK_TYPE)))
        Else
            taskTypeVal = Core_NormalizeTaskType(Core_GetVal(dataArr, rowIdx, mapCol, "Task Type"))
        End If

        If taskTypeVal = "MILESTONE" Then
            effectiveDuration = 1
        End If
    End If

    If leafProfileEnabled Then
        CoreLeafProfile_Count "LeavesProcessed"
        CoreLeafProfile_Count "DataArrayReads", 12
        If useIndexedCore Then
            CoreLeafProfile_Count "IndexedColumnReads", 12
        Else
            CoreLeafProfile_Count "ColumnMapExists", 12
            CoreLeafProfile_Count "ColumnMapItems", 12
        End If
        CoreLeafProfile_AddPhase "TaskPreparation", leafPhaseStart
        leafPhaseStart = CoreLeafProfile_Timestamp()
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

    linkStart = 1
    linkEnd = 0
    If useIndexedCore Then
        linkStart = CLng(corePredOffsets(nodeIndex))
        linkEnd = CLng(corePredOffsets(nodeIndex + 1)) - 1
    ElseIf Not linksBySuccId Is Nothing Then
        If linksBySuccId.Exists(taskId) Then
            Set taskLinks = linksBySuccId(taskId)
            linkEnd = taskLinks.Count
        End If
    End If

    For linkPosition = linkStart To linkEnd
        If leafProfileEnabled Then CoreLeafProfile_Count "PredecessorsEvaluated"
        If useIndexedCore Then
            predNodeIndex = CLng(corePredNodes(linkPosition))
            predId = CStr(corePredIds(linkPosition))
            linkType = CStr(corePredTypes(linkPosition))
            lagVal = CDbl(corePredLags(linkPosition))
            summarySourceId = CStr(corePredSummarySourceIds(linkPosition))
        Else
            If IsObject(taskLinks.Item(linkPosition)) Then
                Set oneLink = taskLinks.Item(linkPosition)
            Else
                oneLink = taskLinks.Item(linkPosition)
            End If
            predNodeIndex = 0
            predId = Core_GetLinkPredId(oneLink)
            linkType = Core_GetLinkType(oneLink)
            lagVal = Core_GetLinkLag(oneLink)
            summarySourceId = Core_GetLinkSummarySourceId(oneLink)
        End If

        If predId = "" Then
            Core_AddBlockingError dataArr, rowIdx, mapCol, blockingErrors, taskId, _
                TextCatalog_Get("CORE.ERROR.MISSING_PREDECESSOR", TEXT_LANGUAGE_EN), _
                "CORE.ERROR.MISSING_PREDECESSOR", Nothing, vbNullString, coreDiagnostics, "ROOT", "DEPENDENCY"
            Exit Sub
        End If

        predExists = False
        predBlocked = False
        predHasStart = False
        predHasFinish = False
        If useIndexedCore Then
            predExists = (predNodeIndex > 0)
            If predExists Then
                predBlocked = CBool(blockedByNode(predNodeIndex))
                predHasStart = CBool(hasCalcByNode(predNodeIndex))
                predHasFinish = predHasStart
            End If
        Else
            predExists = rowById.Exists(predId)
            If predExists Then
                predBlocked = blockingErrors.Exists(predId)
                predHasStart = calcStartById.Exists(predId)
                predHasFinish = calcFinishById.Exists(predId)
            End If
        End If

        If Not predExists Then
            Core_AddBlockingError dataArr, rowIdx, mapCol, blockingErrors, taskId, _
                TextCatalog_Format("CORE.ERROR.MISSING_PREDECESSOR_ID", TEXT_LANGUAGE_EN, _
                    TextCatalog_Arguments("Id", predId)), _
                "CORE.ERROR.MISSING_PREDECESSOR_ID", TextCatalog_Arguments("Id", predId), _
                predId, coreDiagnostics, "ROOT", "DEPENDENCY"
            Exit Sub
        End If

        If predBlocked Then
            Core_AddBlockingError dataArr, rowIdx, mapCol, blockingErrors, taskId, _
                TextCatalog_Format("CORE.ERROR.BLOCKED_PREDECESSOR_ID", TEXT_LANGUAGE_EN, _
                    TextCatalog_Arguments("Id", predId)), _
                "CORE.ERROR.BLOCKED_PREDECESSOR_ID", TextCatalog_Arguments("Id", predId), _
                predId, coreDiagnostics, "INHERITED", "CASCADE"
            Exit Sub
        End If

        Select Case linkType

            Case "SS"
                If leafProfileEnabled Then CoreLeafProfile_Count "LinksSS"
                If predHasStart Then
                    If useIndexedCore Then
                        candidateStart = ApplyLag(calcStartByNode(predNodeIndex), lagVal, calType, "SS")
                        candidatePredecessorDate = calcStartByNode(predNodeIndex)
                    Else
                        candidateStart = ApplyLag(calcStartById(predId), lagVal, calType, "SS")
                        candidatePredecessorDate = calcStartById(predId)
                    End If

                    If summarySourceId <> "" Then
                        If summaryStartBySource Is Nothing Then
                            Set summaryStartBySource = CreateObject("Scripting.Dictionary")
                            Set summaryStartDiagBySource = CreateObject("Scripting.Dictionary")
                        End If
                        If summaryStartBySource.Exists(summarySourceId) Then
                            If CDbl(candidateStart) < CDbl(summaryStartBySource(summarySourceId)) Then
                                summaryStartBySource(summarySourceId) = candidateStart
                                summaryStartDiagBySource(summarySourceId) = Array( _
                                    predId, linkType, lagVal, candidateStart, _
                                    candidatePredecessorDate, _
                                    "START", summarySourceId)
                            End If
                        Else
                            summaryStartBySource(summarySourceId) = candidateStart
                            summaryStartDiagBySource(summarySourceId) = Array( _
                                predId, linkType, lagVal, candidateStart, _
                                candidatePredecessorDate, _
                                "START", summarySourceId)
                        End If
                    Else
                        If Not HasValue(normalAllowedStart) Or CDbl(candidateStart) > CDbl(normalAllowedStart) Then
                            predDiagPredId = predId
                            predDiagLinkType = linkType
                            predDiagLag = lagVal
                            predDiagCandidateDate = candidateStart
                            predDiagPredecessorDate = candidatePredecessorDate
                            predDiagPredecessorDateKind = "START"
                            predDiagSummarySourceId = summarySourceId
                        End If
                        normalAllowedStart = Core_MaxDateIfBoth(normalAllowedStart, candidateStart)
                    End If
                End If

            Case "FF"
                If leafProfileEnabled Then CoreLeafProfile_Count "LinksFF"
                If predHasFinish Then
                    If useIndexedCore Then
                        candidateFinish = ApplyLag(calcFinishByNode(predNodeIndex), lagVal, calType, "FF")
                    Else
                        candidateFinish = ApplyLag(calcFinishById(predId), lagVal, calType, "FF")
                    End If
                    predAllowedFinish = Core_MaxDateIfBoth(predAllowedFinish, candidateFinish)
                End If

            Case "FS", ""
                If leafProfileEnabled Then CoreLeafProfile_Count "LinksFS"
                If predHasFinish Then
                    If useIndexedCore Then
                        candidateStart = ApplyLag(calcFinishByNode(predNodeIndex), lagVal, calType, "FS")
                    Else
                        candidateStart = ApplyLag(calcFinishById(predId), lagVal, calType, "FS")
                    End If
                    If Not HasValue(normalAllowedStart) Or CDbl(candidateStart) > CDbl(normalAllowedStart) Then
                        predDiagPredId = predId
                        predDiagLinkType = IIf(linkType = "", "FS", linkType)
                        predDiagLag = lagVal
                        predDiagCandidateDate = candidateStart
                        If useIndexedCore Then
                            predDiagPredecessorDate = calcFinishByNode(predNodeIndex)
                        Else
                            predDiagPredecessorDate = calcFinishById(predId)
                        End If
                        predDiagPredecessorDateKind = "FINISH"
                        predDiagSummarySourceId = summarySourceId
                    End If
                    normalAllowedStart = Core_MaxDateIfBoth(normalAllowedStart, candidateStart)
                End If

            Case Else
                Core_AddBlockingError dataArr, rowIdx, mapCol, blockingErrors, taskId, _
                    TextCatalog_Format("CORE.ERROR.UNSUPPORTED_LINK_TYPE", TEXT_LANGUAGE_EN, _
                        TextCatalog_Arguments("LinkType", linkType)), _
                    "CORE.ERROR.UNSUPPORTED_LINK_TYPE", TextCatalog_Arguments("LinkType", linkType), _
                    vbNullString, coreDiagnostics, "ROOT", "DEPENDENCY"
                Exit Sub

        End Select

        If leafProfileEnabled Then
            If lagVal < 0# Then
                CoreLeafProfile_Count "LagsNegative"
            ElseIf lagVal > 0# Then
                CoreLeafProfile_Count "LagsPositive"
            Else
                CoreLeafProfile_Count "LagsZero"
            End If
        End If
    Next linkPosition

    If Not summaryStartBySource Is Nothing Then
        For Each parentKey In summaryStartBySource.Keys
            If Not HasValue(summaryAllowedStart) Or CDbl(summaryStartBySource(CStr(parentKey))) > CDbl(summaryAllowedStart) Then
                If summaryStartDiagBySource.Exists(CStr(parentKey)) Then
                    summaryDiagData = summaryStartDiagBySource(CStr(parentKey))
                End If
            End If
            summaryAllowedStart = Core_MaxDateIfBoth(summaryAllowedStart, summaryStartBySource(CStr(parentKey)))
        Next parentKey
    End If

    predAllowedStart = Core_MaxDateIfBoth(normalAllowedStart, summaryAllowedStart)
    If HasValue(predAllowedStart) Then
        If HasValue(summaryAllowedStart) And CDbl(predAllowedStart) = CDbl(summaryAllowedStart) Then
            If IsArray(summaryDiagData) Then
                predDiagPredId = CStr(summaryDiagData(0))
                predDiagLinkType = CStr(summaryDiagData(1))
                predDiagLag = CDbl(summaryDiagData(2))
                predDiagCandidateDate = summaryDiagData(3)
                predDiagPredecessorDate = summaryDiagData(4)
                predDiagPredecessorDateKind = CStr(summaryDiagData(5))
                predDiagSummarySourceId = CStr(summaryDiagData(6))
            End If
        End If
    End If

    If leafProfileEnabled Then
        CoreLeafProfile_AddPhase "PredecessorEvaluation", leafPhaseStart
        leafPhaseStart = CoreLeafProfile_Timestamp()
    End If

    If Not CoreLeaf_ResolveDatesAndConstraints( _
        dataArr, rowIdx, mapCol, blockingErrors, taskId, constraintDiagnostics, _
        dependencyDiagnostics, coreDiagnostics, constraintActive, startConstraintType, _
        startConstraintDate, finishConstraintType, finishConstraintDate, actualStart, _
        actualFinish, forecastStart, forecastFinish, baselineStart, effectiveDuration, _
        calType, predAllowedStart, predAllowedFinish, predDiagPredId, predDiagLinkType, _
        predDiagLag, predDiagCandidateDate, predDiagPredecessorDate, _
        predDiagPredecessorDateKind, predDiagSummarySourceId, leafProfileEnabled, _
        leafPhaseStart, calcStart, calcFinish) Then Exit Sub

    If useIndexedCore Then
        calcStartByNode(nodeIndex) = calcStart
        calcFinishByNode(nodeIndex) = calcFinish
        hasCalcByNode(nodeIndex) = True
    Else
        calcStartById(taskId) = calcStart
        calcFinishById(taskId) = calcFinish
    End If

    If leafProfileEnabled Then CoreLeafProfile_AddPhase "ResultBuffers", leafPhaseStart

End Sub

'------------------------------------------------------------------------------
' FR: Enregistre le diagnostic structure d'une contrainte puis renvoie le message bilingue stocke dans ErrorMsg.
' EN: Records the structured constraint diagnostic and returns the bilingual message stored in ErrorMsg.
'------------------------------------------------------------------------------
Public Sub Core_AddBlockingError( _
    ByRef dataArr As Variant, _
    ByVal rowIdx As Long, _
    ByVal mapCol As Object, _
    ByVal blockingErrors As Object, _
    ByVal taskId As String, _
    ByVal errText As String, _
    Optional ByVal errorCode As String = "", _
    Optional ByVal errorArguments As Object = Nothing, _
    Optional ByVal relatedObjectId As String = "", _
    Optional ByVal coreDiagnostics As Object = Nothing, _
    Optional ByVal classification As String = "ROOT", _
    Optional ByVal family As String = "CORE")

    blockingErrors(CStr(taskId)) = True
    Core_AddErrorMessage_Row dataArr, rowIdx, mapCol, errText, True
    CoreDiagnostics_Record coreDiagnostics, CStr(taskId), errorCode, errorArguments, _
        classification, relatedObjectId, family

End Sub

'------------------------------------------------------------------------------
' FR:
' Transforme un echec de tri topologique en erreur bloquante sur toutes les
' taches feuilles valides, avec un diagnostic de cycle quand il est retrouvable.
'
' EN:
' Converts a topological sort failure into a blocking error on every valid leaf
' task, including a cycle diagnostic when one can be recovered.
'
' Entrees / Inputs:
' - validIds, linksBySuccId, rowById, dataArr et mapCol du run Core.
' - blockingErrors a completer.
'
' Sorties / Outputs:
' - Marque les taches impactees en erreur bloquante.
'
' Appele par / Called by:
' - Run_Calc_Core lorsque l'ordre topologique ne couvre pas tous les IDs valides.
'
' Notes:
' - Cette erreur stoppe le calcul normal car l'ordre de dependance n'est plus fiable.
