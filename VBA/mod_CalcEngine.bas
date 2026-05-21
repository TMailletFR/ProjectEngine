Attribute VB_Name = "mod_CalcEngine"
Option Explicit

Public Sub Run_Calc_Engine(Optional ByVal forceFullRecalcOverride As Boolean = False)

    Dim consoleMessages As Collection

    On Error GoTo ErrHandler

    Set consoleMessages = New Collection

    Run_Calc_Engine_CoreBridge forceFullRecalcOverride

    If IsMacroAbortRequested() Then Exit Sub

    Exit Sub

ErrHandler:
    On Error Resume Next
    Write_CalcState_Snapshot "ERROR"
    On Error GoTo 0

    If consoleMessages Is Nothing Then Set consoleMessages = New Collection

    CalcBridge_AddConsoleMessage consoleMessages, "STOP", _
        BiMsg( _
            "Erreur dans Run_Calc_Engine" & vbCrLf & _
            "-> " & Err.Description, _
            "Error in Run_Calc_Engine" & vbCrLf & _
            "-> " & Err.Description)

    CalcBridge_ShowPlanningConsole consoleMessages

End Sub
Private Sub MarkErrorsInCalc( _
    ByVal tblCalc As ListObject, _
    ByVal mapCalc As Object, _
    ByRef outError() As Variant, _
    ByVal errorIds As Object)

    Dim r As Long
    Dim id As String

    If tblCalc.DataBodyRange Is Nothing Then Exit Sub

    For r = 1 To tblCalc.ListRows.Count
        id = Trim(CStr(tblCalc.DataBodyRange.Cells(r, mapCalc("ID")).value))
        If id <> "" Then
            If errorIds.Exists(id) Then
                outError(r, 1) = "ERROR"
            Else
                outError(r, 1) = ""
            End If
        Else
            outError(r, 1) = ""
        End If
    Next r

    tblCalc.ListColumns("Error flag").DataBodyRange.value = outError

End Sub

Private Sub ShowCalcErrorMessages( _
    ByVal idsDict As Object, _
    ByVal idToWbs As Object, _
    ByVal frProblem As String, _
    ByVal frAction As String, _
    ByVal enProblem As String, _
    ByVal enAction As String)

    If idsDict Is Nothing Then Exit Sub
    If idsDict.Count = 0 Then Exit Sub

    CalcEngine_ShowSingleConsoleMessage "STOP", _
        BuildGroupedMessage(idsDict, idToWbs, frProblem, frAction, enProblem, enAction)

End Sub




Private Function BuildGroupedMessage( _
    ByVal idsDict As Object, _
    ByVal idToWbs As Object, _
    ByVal frProblem As String, _
    ByVal frAction As String, _
    ByVal enProblem As String, _
    ByVal enAction As String) As String

    Dim idsLine As String
    Dim wbsLine As String

    idsLine = BuildInlineList(idsDict, 20)
    wbsLine = BuildInlineWBSList(idsDict, idToWbs, 20)

    BuildGroupedMessage = _
        frProblem & vbCrLf & _
        "-> " & frAction & vbCrLf & vbCrLf & _
        "IDs : " & idsLine & vbCrLf & _
        "WBS : " & wbsLine & vbCrLf & vbCrLf & _
        enProblem & vbCrLf & _
        "-> " & enAction & vbCrLf & vbCrLf & _
        "IDs: " & idsLine & vbCrLf & _
        "WBS: " & wbsLine

End Function

Private Function BuildInlineList(ByVal idsDict As Object, ByVal maxItems As Long) As String

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

    BuildInlineList = result

End Function

Private Function BuildInlineWBSList(ByVal idsDict As Object, ByVal idToWbs As Object, ByVal maxItems As Long) As String

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
                itemText = CStr(idToWbs(CStr(key)))
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

    BuildInlineWBSList = result

End Function

Public Sub ComputeCurrentFloatAndCritical( _
    ByVal tblCalc As ListObject, _
    ByVal mapCalc As Object, _
    ByVal idToRow As Object, _
    ByVal childrenById As Object, _
    ByVal directChildrenById As Object, _
    ByVal parentIds As Object, _
    ByVal validIds As Object, _
    ByVal topoOrder As Collection, _
    Optional ByVal consoleMessages As Collection = Nothing)

    Dim dataArr As Variant
    Dim currentLateStartById As Object
    Dim currentLateFinishById As Object
    Dim currentTotalFloatById As Object
    Dim currentFreeFloatById As Object
    Dim predLagBySuccPred As Object
    Dim predTypeBySuccPred As Object
    Dim reverseTopo As Collection
    Dim networkFinishById As Object

    Dim projectCurrentFinish As Variant
    Dim rowIndex As Long
    Dim r As Long
    Dim taskId As String
    Dim idKey As Variant
    Dim succId As Variant

    Dim calcStartVal As Variant
    Dim calcFinishVal As Variant
    Dim succStartVal As Variant
    Dim succFinishVal As Variant
    Dim currentDuration As Double
    Dim candidateLateFinish As Variant
    Dim candidateFreeFloat As Double
    Dim minFreeFloat As Variant
    Dim linkKey As String
    Dim effectiveLag As Double
    Dim linkType As String
    Dim candidateLateStart As Double
    Dim constraintActive As String
    Dim startConstraintType As String
    Dim finishConstraintType As String
    Dim startConstraintDate As Variant
    Dim finishConstraintDate As Variant
    Dim constraintLateFinish As Variant

    Dim outTF() As Variant
    Dim outFF() As Variant
    Dim outCritical() As Variant

    Dim hasNegativeFloat As Boolean
    Dim useMultiNetwork As Boolean

    If tblCalc Is Nothing Then Exit Sub
    If tblCalc.DataBodyRange Is Nothing Then Exit Sub

    Set currentLateStartById = CreateObject("Scripting.Dictionary")
    Set currentLateFinishById = CreateObject("Scripting.Dictionary")
    Set currentTotalFloatById = CreateObject("Scripting.Dictionary")
    Set currentFreeFloatById = CreateObject("Scripting.Dictionary")
    Set predLagBySuccPred = BuildPredLagMapFromLogicLinks()
    Set predTypeBySuccPred = BuildPredTypeMapFromLogicLinks()
    Set reverseTopo = New Collection
    Set networkFinishById = CreateObject("Scripting.Dictionary")

    dataArr = tblCalc.DataBodyRange.value
    projectCurrentFinish = Empty
    useMultiNetwork = IsCriticalPathMultiNetworkEnabled()

    For Each idKey In validIds.Keys
        rowIndex = idToRow(CStr(idKey))
        calcFinishVal = GetCellValue(dataArr(rowIndex, mapCalc("Calculated Finish")))

        If HasValue(calcFinishVal) Then
            If Not HasValue(projectCurrentFinish) Then
                projectCurrentFinish = calcFinishVal
            ElseIf CDbl(calcFinishVal) > CDbl(projectCurrentFinish) Then
                projectCurrentFinish = calcFinishVal
            End If
        End If
    Next idKey

    If useMultiNetwork Then
        Set networkFinishById = BuildCurrentNetworkFinishById(dataArr, mapCalc, idToRow, childrenById, validIds)
    End If

    ReDim outTF(1 To tblCalc.ListRows.Count, 1 To 1)
    ReDim outFF(1 To tblCalc.ListRows.Count, 1 To 1)
    ReDim outCritical(1 To tblCalc.ListRows.Count, 1 To 1)

    If Not HasValue(projectCurrentFinish) Then
        tblCalc.ListColumns("Total Float").DataBodyRange.value = outTF
        tblCalc.ListColumns("Free Float").DataBodyRange.value = outFF
        tblCalc.ListColumns("Critical Path").DataBodyRange.value = outCritical
        Exit Sub
    End If

    For r = topoOrder.Count To 1 Step -1
        reverseTopo.Add CStr(topoOrder(r))
    Next r

    For Each idKey In reverseTopo

        taskId = CStr(idKey)
        rowIndex = idToRow(taskId)

        calcStartVal = GetCellValue(dataArr(rowIndex, mapCalc("Calculated Start")))
        calcFinishVal = GetCellValue(dataArr(rowIndex, mapCalc("Calculated Finish")))

        If Not HasValue(calcStartVal) Or Not HasValue(calcFinishVal) Then
            GoTo NextCurrentBackwardTask
        End If

        currentDuration = CDbl(calcFinishVal) - CDbl(calcStartVal) + 1
        candidateLateFinish = Empty

        For Each succId In childrenById(taskId)
            If validIds.Exists(CStr(succId)) Then

                linkKey = CStr(succId) & "|" & taskId

                If predLagBySuccPred.Exists(linkKey) Then
                    effectiveLag = CDbl(predLagBySuccPred(linkKey))
                Else
                    effectiveLag = 0#
                End If

                If predTypeBySuccPred.Exists(linkKey) Then
                    linkType = CStr(predTypeBySuccPred(linkKey))
                Else
                    linkType = "FS"
                End If

                Select Case linkType

                    Case "SS"
                        If currentLateStartById.Exists(CStr(succId)) Then
                            candidateLateStart = CDbl(currentLateStartById(CStr(succId))) - effectiveLag

                            If Not HasValue(candidateLateFinish) Then
                                candidateLateFinish = candidateLateStart + currentDuration - 1
                            ElseIf candidateLateStart + currentDuration - 1 < CDbl(candidateLateFinish) Then
                                candidateLateFinish = candidateLateStart + currentDuration - 1
                            End If
                        End If

                    Case "FF"
                        If currentLateFinishById.Exists(CStr(succId)) Then
                            If Not HasValue(candidateLateFinish) Then
                                candidateLateFinish = CDbl(currentLateFinishById(CStr(succId))) - effectiveLag
                            ElseIf CDbl(currentLateFinishById(CStr(succId))) - effectiveLag < CDbl(candidateLateFinish) Then
                                candidateLateFinish = CDbl(currentLateFinishById(CStr(succId))) - effectiveLag
                            End If
                        End If

                    Case Else
                        If currentLateStartById.Exists(CStr(succId)) Then
                            If Not HasValue(candidateLateFinish) Then
                                candidateLateFinish = CDbl(currentLateStartById(CStr(succId))) - effectiveLag - 1
                            ElseIf CDbl(currentLateStartById(CStr(succId))) - effectiveLag - 1 < CDbl(candidateLateFinish) Then
                                candidateLateFinish = CDbl(currentLateStartById(CStr(succId))) - effectiveLag - 1
                            End If
                        End If

                End Select
            End If
        Next succId

        If Not HasValue(candidateLateFinish) Then
            If useMultiNetwork And networkFinishById.Exists(taskId) Then
                currentLateFinishById(taskId) = networkFinishById(taskId)
            Else
                currentLateFinishById(taskId) = projectCurrentFinish
            End If
        Else
            currentLateFinishById(taskId) = candidateLateFinish
        End If

        constraintActive = vbNullString
        startConstraintType = vbNullString
        finishConstraintType = vbNullString
        startConstraintDate = Empty
        finishConstraintDate = Empty
        constraintLateFinish = Empty

        If mapCalc.Exists("Constraint Active") Then
            constraintActive = UCase$(Trim$(CStr(dataArr(rowIndex, mapCalc("Constraint Active")))))
        End If

        If constraintActive = "YES" Then
            If mapCalc.Exists("Start Constraint Type") Then
                startConstraintType = Trim$(CStr(dataArr(rowIndex, mapCalc("Start Constraint Type"))))
            End If
            If mapCalc.Exists("Start Constraint Date") Then
                startConstraintDate = GetCellValue(dataArr(rowIndex, mapCalc("Start Constraint Date")))
            End If
            If mapCalc.Exists("Finish Constraint Type") Then
                finishConstraintType = Trim$(CStr(dataArr(rowIndex, mapCalc("Finish Constraint Type"))))
            End If
            If mapCalc.Exists("Finish Constraint Date") Then
                finishConstraintDate = GetCellValue(dataArr(rowIndex, mapCalc("Finish Constraint Date")))
            End If

            If (startConstraintType = "Start No Later Than" Or startConstraintType = "Must Start On") _
                And HasValue(startConstraintDate) Then
                constraintLateFinish = CDbl(startConstraintDate) + currentDuration - 1
                If CDbl(constraintLateFinish) < CDbl(currentLateFinishById(taskId)) Then
                    currentLateFinishById(taskId) = constraintLateFinish
                End If
            End If

            If (finishConstraintType = "Finish No Later Than" Or finishConstraintType = "Must Finish On") _
                And HasValue(finishConstraintDate) Then
                If CDbl(finishConstraintDate) < CDbl(currentLateFinishById(taskId)) Then
                    currentLateFinishById(taskId) = finishConstraintDate
                End If
            End If
        End If

        currentLateStartById(taskId) = CDbl(currentLateFinishById(taskId)) - currentDuration + 1

NextCurrentBackwardTask:
    Next idKey

    For Each idKey In validIds.Keys

        taskId = CStr(idKey)
        rowIndex = idToRow(taskId)

        calcStartVal = GetCellValue(dataArr(rowIndex, mapCalc("Calculated Start")))

        If HasValue(calcStartVal) And currentLateStartById.Exists(taskId) Then
            currentTotalFloatById(taskId) = CDbl(currentLateStartById(taskId)) - CDbl(calcStartVal)
            outTF(rowIndex, 1) = currentTotalFloatById(taskId)

            If CDbl(currentTotalFloatById(taskId)) < 0 Then hasNegativeFloat = True

            If Not IsActualFinishedInCalcArray(dataArr, rowIndex, mapCalc) Then
                If CDbl(currentTotalFloatById(taskId)) <= 0 Then
                    outCritical(rowIndex, 1) = "CRITICAL"
                Else
                    outCritical(rowIndex, 1) = ""
                End If
            Else
                outCritical(rowIndex, 1) = ""
            End If
        End If

    Next idKey

    For Each idKey In validIds.Keys

        taskId = CStr(idKey)
        rowIndex = idToRow(taskId)

        calcStartVal = GetCellValue(dataArr(rowIndex, mapCalc("Calculated Start")))
        calcFinishVal = GetCellValue(dataArr(rowIndex, mapCalc("Calculated Finish")))

        If Not HasValue(calcStartVal) Or Not HasValue(calcFinishVal) Then GoTo NextCurrentFreeFloatTask

        If childrenById(taskId).Count = 0 Then
            If currentTotalFloatById.Exists(taskId) Then
                currentFreeFloatById(taskId) = currentTotalFloatById(taskId)
                outFF(rowIndex, 1) = currentFreeFloatById(taskId)

                If CDbl(currentFreeFloatById(taskId)) < 0 Then hasNegativeFloat = True
            End If
        Else
            minFreeFloat = Empty

            For Each succId In childrenById(taskId)
                If validIds.Exists(CStr(succId)) Then

                    linkKey = CStr(succId) & "|" & taskId

                    If predLagBySuccPred.Exists(linkKey) Then
                        effectiveLag = CDbl(predLagBySuccPred(linkKey))
                    Else
                        effectiveLag = 0#
                    End If

                    If predTypeBySuccPred.Exists(linkKey) Then
                        linkType = CStr(predTypeBySuccPred(linkKey))
                    Else
                        linkType = "FS"
                    End If

                    Select Case linkType

                        Case "SS"
                            succStartVal = GetCellValue(dataArr(idToRow(CStr(succId)), mapCalc("Calculated Start")))
                            If HasValue(succStartVal) Then
                                candidateFreeFloat = CDbl(succStartVal) - effectiveLag - CDbl(calcStartVal)

                                If Not HasValue(minFreeFloat) Then
                                    minFreeFloat = candidateFreeFloat
                                ElseIf candidateFreeFloat < CDbl(minFreeFloat) Then
                                    minFreeFloat = candidateFreeFloat
                                End If
                            End If

                        Case "FF"
                            succFinishVal = GetCellValue(dataArr(idToRow(CStr(succId)), mapCalc("Calculated Finish")))
                            If HasValue(succFinishVal) Then
                                candidateFreeFloat = CDbl(succFinishVal) - effectiveLag - CDbl(calcFinishVal)

                                If Not HasValue(minFreeFloat) Then
                                    minFreeFloat = candidateFreeFloat
                                ElseIf candidateFreeFloat < CDbl(minFreeFloat) Then
                                    minFreeFloat = candidateFreeFloat
                                End If
                            End If

                        Case Else
                            succStartVal = GetCellValue(dataArr(idToRow(CStr(succId)), mapCalc("Calculated Start")))
                            If HasValue(succStartVal) Then
                                candidateFreeFloat = CDbl(succStartVal) - effectiveLag - 1 - CDbl(calcFinishVal)

                                If Not HasValue(minFreeFloat) Then
                                    minFreeFloat = candidateFreeFloat
                                ElseIf candidateFreeFloat < CDbl(minFreeFloat) Then
                                    minFreeFloat = candidateFreeFloat
                                End If
                            End If

                    End Select
                End If
            Next succId

            If HasValue(minFreeFloat) Then
                currentFreeFloatById(taskId) = minFreeFloat
                outFF(rowIndex, 1) = currentFreeFloatById(taskId)

                If CDbl(currentFreeFloatById(taskId)) < 0 Then hasNegativeFloat = True
            End If
        End If

        If currentFreeFloatById.Exists(taskId) Then
            minFreeFloat = currentFreeFloatById(taskId)

            constraintActive = vbNullString
            startConstraintType = vbNullString
            finishConstraintType = vbNullString
            startConstraintDate = Empty
            finishConstraintDate = Empty

            If mapCalc.Exists("Constraint Active") Then
                constraintActive = UCase$(Trim$(CStr(dataArr(rowIndex, mapCalc("Constraint Active")))))
            End If

            If constraintActive = "YES" Then
                If mapCalc.Exists("Start Constraint Type") Then
                    startConstraintType = Trim$(CStr(dataArr(rowIndex, mapCalc("Start Constraint Type"))))
                End If
                If mapCalc.Exists("Start Constraint Date") Then
                    startConstraintDate = GetCellValue(dataArr(rowIndex, mapCalc("Start Constraint Date")))
                End If
                If mapCalc.Exists("Finish Constraint Type") Then
                    finishConstraintType = Trim$(CStr(dataArr(rowIndex, mapCalc("Finish Constraint Type"))))
                End If
                If mapCalc.Exists("Finish Constraint Date") Then
                    finishConstraintDate = GetCellValue(dataArr(rowIndex, mapCalc("Finish Constraint Date")))
                End If

                If (startConstraintType = "Start No Later Than" Or startConstraintType = "Must Start On") _
                    And HasValue(startConstraintDate) Then
                    candidateFreeFloat = CDbl(startConstraintDate) - CDbl(calcStartVal)
                    If candidateFreeFloat < CDbl(minFreeFloat) Then minFreeFloat = candidateFreeFloat
                End If

                If (finishConstraintType = "Finish No Later Than" Or finishConstraintType = "Must Finish On") _
                    And HasValue(finishConstraintDate) Then
                    candidateFreeFloat = CDbl(finishConstraintDate) - CDbl(calcFinishVal)
                    If candidateFreeFloat < CDbl(minFreeFloat) Then minFreeFloat = candidateFreeFloat
                End If
            End If

            currentFreeFloatById(taskId) = minFreeFloat
            outFF(rowIndex, 1) = currentFreeFloatById(taskId)

            If CDbl(currentFreeFloatById(taskId)) < 0 Then hasNegativeFloat = True
        End If

NextCurrentFreeFloatTask:
    Next idKey

    RollupFloatToParents idToRow, directChildrenById, parentIds, outTF, outFF

    For r = 1 To tblCalc.ListRows.Count
        If HasValue(outTF(r, 1)) Then
            If Not IsActualFinishedInCalcArray(dataArr, r, mapCalc) Then
                If CDbl(outTF(r, 1)) <= 0 Then
                    outCritical(r, 1) = "CRITICAL"
                Else
                    outCritical(r, 1) = ""
                End If
            Else
                outCritical(r, 1) = ""
            End If
        Else
            outCritical(r, 1) = ""
        End If
    Next r

    tblCalc.ListColumns("Total Float").DataBodyRange.value = outTF
    tblCalc.ListColumns("Free Float").DataBodyRange.value = outFF
    tblCalc.ListColumns("Critical Path").DataBodyRange.value = outCritical

    If hasNegativeFloat Then
        CalcEngine_AddOrShowConsoleMessage consoleMessages, "WARNING", _
            BiMsg( _
                "Float négatif détecté dans le planning actuel" & vbCrLf & _
                "-> vérifier la logique, les dates, les lags ou les prévisions", _
                "Negative float detected in the current schedule" & vbCrLf & _
                "-> check logic, dates, lags or forecasts")
    End If

End Sub

Public Sub ComputeLongestPath( _
    ByVal tblCalc As ListObject, _
    ByVal mapCalc As Object, _
    ByVal idToRow As Object, _
    ByVal predsById As Object, _
    ByVal childrenById As Object, _
    ByVal validIds As Object)

    Dim dataArr As Variant
    Dim outLP() As Variant
    Dim lpById As Object
    Dim queuedById As Object
    Dim predLagBySuccPred As Object
    Dim predTypeBySuccPred As Object
    Dim networkFinishById As Object
    Dim queue As Collection

    Dim projectFinish As Variant
    Dim idKey As Variant
    Dim taskId As String
    Dim predId As Variant
    Dim rowIndex As Long
    Dim finishVal As Variant
    Dim targetFinish As Variant
    Dim linkKey As String
    Dim linkType As String
    Dim effectiveLag As Double

    If tblCalc Is Nothing Then Exit Sub
    If tblCalc.DataBodyRange Is Nothing Then Exit Sub
    If mapCalc Is Nothing Then Exit Sub
    If idToRow Is Nothing Then Exit Sub
    If predsById Is Nothing Then Exit Sub
    If childrenById Is Nothing Then Exit Sub
    If validIds Is Nothing Then Exit Sub
    If Not mapCalc.Exists("Calculated Start") Then Exit Sub
    If Not mapCalc.Exists("Calculated Finish") Then Exit Sub
    If Not mapCalc.Exists("Longest Path") Then Exit Sub

    dataArr = tblCalc.DataBodyRange.value
    ReDim outLP(1 To tblCalc.ListRows.Count, 1 To 1)

    Set lpById = CreateObject("Scripting.Dictionary")
    Set queuedById = CreateObject("Scripting.Dictionary")
    Set predLagBySuccPred = BuildPredLagMapFromLogicLinks()
    Set predTypeBySuccPred = BuildPredTypeMapFromLogicLinks()
    Set queue = New Collection

    projectFinish = Empty

    For Each idKey In validIds.Keys
        rowIndex = CLng(idToRow(CStr(idKey)))
        finishVal = GetCellValue(dataArr(rowIndex, mapCalc("Calculated Finish")))

        If HasValue(finishVal) Then
            If Not IsActualFinishedInCalcArray(dataArr, rowIndex, mapCalc) Then
                If Not HasValue(projectFinish) Then
                    projectFinish = finishVal
                ElseIf CDbl(finishVal) > CDbl(projectFinish) Then
                    projectFinish = finishVal
                End If
            End If
        End If
    Next idKey

    If Not HasValue(projectFinish) Then
        tblCalc.ListColumns("Longest Path").DataBodyRange.value = outLP
        Exit Sub
    End If

    If IsCriticalPathMultiNetworkEnabled() Then
        Set networkFinishById = BuildCurrentNetworkFinishById(dataArr, mapCalc, idToRow, childrenById, validIds)
    Else
        Set networkFinishById = Nothing
    End If

    For Each idKey In validIds.Keys
        taskId = CStr(idKey)
        rowIndex = CLng(idToRow(taskId))
        finishVal = GetCellValue(dataArr(rowIndex, mapCalc("Calculated Finish")))

        If HasValue(finishVal) Then
            If Not IsActualFinishedInCalcArray(dataArr, rowIndex, mapCalc) Then
                If Not networkFinishById Is Nothing Then
                    If networkFinishById.Exists(taskId) Then
                        targetFinish = networkFinishById(taskId)
                    Else
                        targetFinish = Empty
                    End If
                Else
                    targetFinish = projectFinish
                End If

                If HasValue(targetFinish) Then
                    If CalcEngine_DatesEqual(finishVal, targetFinish) Then
                        CalcEngine_AddLongestPathTask taskId, lpById, queuedById, queue
                    End If
                End If
            End If
        End If
    Next idKey

    Do While queue.Count > 0
        taskId = CStr(queue(1))
        queue.Remove 1

        If predsById.Exists(taskId) Then
            For Each predId In predsById(taskId)
                If validIds.Exists(CStr(predId)) Then
                    linkKey = taskId & "|" & CStr(predId)

                    If predLagBySuccPred.Exists(linkKey) Then
                        effectiveLag = CDbl(predLagBySuccPred(linkKey))
                    Else
                        effectiveLag = 0#
                    End If

                    If predTypeBySuccPred.Exists(linkKey) Then
                        linkType = CStr(predTypeBySuccPred(linkKey))
                    Else
                        linkType = "FS"
                    End If

                    If CalcEngine_IsDrivingLongestPathLink(dataArr, mapCalc, idToRow, taskId, CStr(predId), linkType, effectiveLag) Then
                        CalcEngine_AddLongestPathTask CStr(predId), lpById, queuedById, queue
                    End If
                End If
            Next predId
        End If
    Loop

    For Each idKey In lpById.Keys
        rowIndex = CLng(idToRow(CStr(idKey)))
        If Not IsActualFinishedInCalcArray(dataArr, rowIndex, mapCalc) Then
            outLP(rowIndex, 1) = "LONGEST"
        End If
    Next idKey

    tblCalc.ListColumns("Longest Path").DataBodyRange.value = outLP

End Sub

Private Sub CalcEngine_AddLongestPathTask( _
    ByVal taskId As String, _
    ByVal lpById As Object, _
    ByVal queuedById As Object, _
    ByVal queue As Collection)

    If Len(taskId) = 0 Then Exit Sub

    If Not lpById.Exists(taskId) Then
        lpById(taskId) = True
    End If

    If Not queuedById.Exists(taskId) Then
        queuedById(taskId) = True
        queue.Add taskId
    End If

End Sub

Private Function CalcEngine_IsDrivingLongestPathLink( _
    ByRef dataArr As Variant, _
    ByVal mapCalc As Object, _
    ByVal idToRow As Object, _
    ByVal succId As String, _
    ByVal predId As String, _
    ByVal linkType As String, _
    ByVal effectiveLag As Double) As Boolean

    Dim succRow As Long
    Dim predRow As Long
    Dim succStart As Variant
    Dim succFinish As Variant
    Dim predStart As Variant
    Dim predFinish As Variant

    If Not idToRow.Exists(succId) Then Exit Function
    If Not idToRow.Exists(predId) Then Exit Function

    succRow = CLng(idToRow(succId))
    predRow = CLng(idToRow(predId))

    succStart = GetCellValue(dataArr(succRow, mapCalc("Calculated Start")))
    succFinish = GetCellValue(dataArr(succRow, mapCalc("Calculated Finish")))
    predStart = GetCellValue(dataArr(predRow, mapCalc("Calculated Start")))
    predFinish = GetCellValue(dataArr(predRow, mapCalc("Calculated Finish")))

    Select Case UCase$(Trim$(linkType))
        Case "SS"
            If HasValue(succStart) And HasValue(predStart) Then
                CalcEngine_IsDrivingLongestPathLink = CalcEngine_DatesEqual(CDbl(succStart), CDbl(predStart) + effectiveLag)
            End If

        Case "FF"
            If HasValue(succFinish) And HasValue(predFinish) Then
                CalcEngine_IsDrivingLongestPathLink = CalcEngine_DatesEqual(CDbl(succFinish), CDbl(predFinish) + effectiveLag)
            End If

        Case Else
            If HasValue(succStart) And HasValue(predFinish) Then
                CalcEngine_IsDrivingLongestPathLink = CalcEngine_DatesEqual(CDbl(succStart), CDbl(predFinish) + 1 + effectiveLag)
            End If
    End Select

End Function

Private Function CalcEngine_DatesEqual(ByVal leftVal As Variant, ByVal rightVal As Variant) As Boolean

    If Not HasValue(leftVal) Then Exit Function
    If Not HasValue(rightVal) Then Exit Function

    CalcEngine_DatesEqual = (Abs(CDbl(leftVal) - CDbl(rightVal)) < 0.000001)

End Function
Public Sub ComputeCriticalPathREX( _
    ByVal tblCalc As ListObject, _
    ByVal mapCalc As Object, _
    ByVal idToRow As Object, _
    ByVal predsById As Object, _
    ByVal childrenById As Object, _
    ByVal directChildrenById As Object, _
    ByVal parentIds As Object, _
    ByVal validIds As Object, _
    ByVal topoOrder As Collection, _
    ByVal idToWbs As Object, _
    ByRef outCriticalREX() As Variant, _
    ByVal errMissingBaselineForREX As Object, _
    Optional ByVal consoleMessages As Collection = Nothing)

    Dim dataArr As Variant
    Dim rexStartById As Object
    Dim rexFinishById As Object
    Dim rexLateStartById As Object
    Dim rexLateFinishById As Object
    Dim rexTotalFloatById As Object
    Dim rexFreeFloatById As Object
    Dim predLagBySuccPred As Object
    Dim predTypeBySuccPred As Object
    Dim reverseTopo As Collection
    Dim networkFinishById As Object

    Dim projectBaselineFinish As Variant

    Dim r As Long
    Dim rowIndex As Long
    Dim taskId As String
    Dim idKey As Variant
    Dim predId As Variant
    Dim succId As Variant
    Dim preds As Collection

    Dim baselineDuration As Variant
    Dim baselineStart As Variant
    Dim startVal As Variant
    Dim finishVal As Variant
    Dim effectiveLag As Double
    Dim linkKey As String
    Dim linkType As String
    Dim candidateStart As Double
    Dim bestStart As Variant
    Dim candidateLateFinish As Variant
    Dim candidateLateStart As Double
    Dim hasValidPred As Boolean

    Dim hasNegativeFloat As Boolean
    Dim tf As Double
    Dim ff As Variant
    Dim minFF As Variant
    Dim candidateFF As Double

    Dim outTF() As Variant
    Dim outFF() As Variant
    Dim useMultiNetwork As Boolean

    If tblCalc Is Nothing Then Exit Sub
    If tblCalc.DataBodyRange Is Nothing Then Exit Sub

    Set rexStartById = CreateObject("Scripting.Dictionary")
    Set rexFinishById = CreateObject("Scripting.Dictionary")
    Set rexLateStartById = CreateObject("Scripting.Dictionary")
    Set rexLateFinishById = CreateObject("Scripting.Dictionary")
    Set rexTotalFloatById = CreateObject("Scripting.Dictionary")
    Set rexFreeFloatById = CreateObject("Scripting.Dictionary")
    Set predLagBySuccPred = BuildPredLagMapFromLogicLinks()
    Set predTypeBySuccPred = BuildPredTypeMapFromLogicLinks()
    Set reverseTopo = New Collection
    Set networkFinishById = CreateObject("Scripting.Dictionary")

    dataArr = tblCalc.DataBodyRange.value
    projectBaselineFinish = Empty
    useMultiNetwork = IsCriticalPathMultiNetworkEnabled()

    ReDim outTF(1 To tblCalc.ListRows.Count, 1 To 1)
    ReDim outFF(1 To tblCalc.ListRows.Count, 1 To 1)

    For Each idKey In topoOrder

        taskId = CStr(idKey)
        rowIndex = idToRow(taskId)

        baselineDuration = GetCellValue(dataArr(rowIndex, mapCalc("Baseline Duration")))
        baselineStart = GetCellValue(dataArr(rowIndex, mapCalc("Baseline Start")))

        If Not HasValue(baselineDuration) Then
            If IsMilestoneTaskType(dataArr, mapCalc, rowIndex) Then
                baselineDuration = 1
            End If
        End If

        bestStart = Empty
        hasValidPred = False

        If Not HasValue(baselineDuration) Then
            errMissingBaselineForREX(taskId) = True
            GoTo NextRexForward
        End If

        Set preds = predsById(taskId)

        For Each predId In preds

            If validIds.Exists(CStr(predId)) Then

                hasValidPred = True
                linkKey = taskId & "|" & CStr(predId)

                If predLagBySuccPred.Exists(linkKey) Then
                    effectiveLag = CDbl(predLagBySuccPred(linkKey))
                Else
                    effectiveLag = 0#
                End If

                If predTypeBySuccPred.Exists(linkKey) Then
                    linkType = CStr(predTypeBySuccPred(linkKey))
                Else
                    linkType = "FS"
                End If

                Select Case linkType

                    Case "SS"
                        If rexStartById.Exists(CStr(predId)) Then
                            candidateStart = CDbl(rexStartById(CStr(predId))) + effectiveLag

                            If Not HasValue(bestStart) Then
                                bestStart = candidateStart
                            ElseIf candidateStart > CDbl(bestStart) Then
                                bestStart = candidateStart
                            End If
                        End If

                    Case "FF"
                        If rexFinishById.Exists(CStr(predId)) Then
                            candidateStart = CDbl(rexFinishById(CStr(predId))) + effectiveLag - CDbl(baselineDuration) + 1

                            If Not HasValue(bestStart) Then
                                bestStart = candidateStart
                            ElseIf candidateStart > CDbl(bestStart) Then
                                bestStart = candidateStart
                            End If
                        End If

                    Case Else
                        If rexFinishById.Exists(CStr(predId)) Then
                            candidateStart = CDbl(rexFinishById(CStr(predId))) + 1 + effectiveLag

                            If Not HasValue(bestStart) Then
                                bestStart = candidateStart
                            ElseIf candidateStart > CDbl(bestStart) Then
                                bestStart = candidateStart
                            End If
                        End If

                End Select

            End If

        Next predId

        If hasValidPred Then
            If Not HasValue(bestStart) Then
                errMissingBaselineForREX(taskId) = True
                GoTo NextRexForward
            End If

            startVal = bestStart
        Else
            If Not HasValue(baselineStart) Then
                errMissingBaselineForREX(taskId) = True
                GoTo NextRexForward
            End If

            startVal = baselineStart
        End If

        finishVal = CDbl(startVal) + CDbl(baselineDuration) - 1

        rexStartById(taskId) = startVal
        rexFinishById(taskId) = finishVal

        If Not HasValue(projectBaselineFinish) Then
            projectBaselineFinish = finishVal
        ElseIf CDbl(finishVal) > CDbl(projectBaselineFinish) Then
            projectBaselineFinish = finishVal
        End If

NextRexForward:
    Next idKey

    If errMissingBaselineForREX.Count > 0 Then Exit Sub
    If Not HasValue(projectBaselineFinish) Then Exit Sub

    If useMultiNetwork Then
        Set networkFinishById = BuildRexNetworkFinishById(rexFinishById, childrenById, validIds)
    End If

    For r = topoOrder.Count To 1 Step -1
        reverseTopo.Add CStr(topoOrder(r))
    Next r

    For Each idKey In reverseTopo

        taskId = CStr(idKey)
        rowIndex = idToRow(taskId)

        baselineDuration = GetCellValue(dataArr(rowIndex, mapCalc("Baseline Duration")))

        If Not HasValue(baselineDuration) Then
            If IsMilestoneTaskType(dataArr, mapCalc, rowIndex) Then
                baselineDuration = 1
            End If
        End If

        candidateLateFinish = Empty

        If Not HasValue(baselineDuration) Then GoTo NextRexBackward

        For Each succId In childrenById(taskId)

            If validIds.Exists(CStr(succId)) Then

                linkKey = CStr(succId) & "|" & taskId

                If predLagBySuccPred.Exists(linkKey) Then
                    effectiveLag = CDbl(predLagBySuccPred(linkKey))
                Else
                    effectiveLag = 0#
                End If

                If predTypeBySuccPred.Exists(linkKey) Then
                    linkType = CStr(predTypeBySuccPred(linkKey))
                Else
                    linkType = "FS"
                End If

                Select Case linkType

                    Case "SS"
                        If rexLateStartById.Exists(CStr(succId)) Then
                            candidateLateStart = CDbl(rexLateStartById(CStr(succId))) - effectiveLag

                            If Not HasValue(candidateLateFinish) Then
                                candidateLateFinish = candidateLateStart + CDbl(baselineDuration) - 1
                            ElseIf candidateLateStart + CDbl(baselineDuration) - 1 < CDbl(candidateLateFinish) Then
                                candidateLateFinish = candidateLateStart + CDbl(baselineDuration) - 1
                            End If
                        End If

                    Case "FF"
                        If rexLateFinishById.Exists(CStr(succId)) Then
                            If Not HasValue(candidateLateFinish) Then
                                candidateLateFinish = CDbl(rexLateFinishById(CStr(succId))) - effectiveLag
                            ElseIf CDbl(rexLateFinishById(CStr(succId))) - effectiveLag < CDbl(candidateLateFinish) Then
                                candidateLateFinish = CDbl(rexLateFinishById(CStr(succId))) - effectiveLag
                            End If
                        End If

                    Case Else
                        If rexLateStartById.Exists(CStr(succId)) Then
                            If Not HasValue(candidateLateFinish) Then
                                candidateLateFinish = CDbl(rexLateStartById(CStr(succId))) - effectiveLag - 1
                            ElseIf CDbl(rexLateStartById(CStr(succId))) - effectiveLag - 1 < CDbl(candidateLateFinish) Then
                                candidateLateFinish = CDbl(rexLateStartById(CStr(succId))) - effectiveLag - 1
                            End If
                        End If

                End Select
            End If
        Next succId

        If Not HasValue(candidateLateFinish) Then
            If useMultiNetwork And networkFinishById.Exists(taskId) Then
                rexLateFinishById(taskId) = networkFinishById(taskId)
            Else
                rexLateFinishById(taskId) = projectBaselineFinish
            End If
        Else
            rexLateFinishById(taskId) = candidateLateFinish
        End If

        rexLateStartById(taskId) = CDbl(rexLateFinishById(taskId)) - CDbl(baselineDuration) + 1

NextRexBackward:
    Next idKey

    For Each idKey In validIds.Keys

        taskId = CStr(idKey)
        rowIndex = idToRow(taskId)

        If rexLateStartById.Exists(taskId) And rexStartById.Exists(taskId) Then

            tf = CDbl(rexLateStartById(taskId)) - CDbl(rexStartById(taskId))
            rexTotalFloatById(taskId) = tf

            If tf < 0 Then hasNegativeFloat = True

            If tf <= 0 Then
                outCriticalREX(rowIndex, 1) = "CRITICAL"
            Else
                outCriticalREX(rowIndex, 1) = ""
            End If

        End If

    Next idKey

    For Each idKey In validIds.Keys

        taskId = CStr(idKey)
        rowIndex = idToRow(taskId)
        ff = Empty

        If Not rexStartById.Exists(taskId) Then GoTo NextRexFreeFloatTask
        If Not rexFinishById.Exists(taskId) Then GoTo NextRexFreeFloatTask

        If childrenById(taskId).Count = 0 Then
            If rexTotalFloatById.Exists(taskId) Then ff = rexTotalFloatById(taskId)
        Else
            minFF = Empty

            For Each succId In childrenById(taskId)

                If validIds.Exists(CStr(succId)) Then

                    If rexStartById.Exists(CStr(succId)) And rexFinishById.Exists(CStr(succId)) Then

                        linkKey = CStr(succId) & "|" & taskId

                        If predLagBySuccPred.Exists(linkKey) Then
                            effectiveLag = CDbl(predLagBySuccPred(linkKey))
                        Else
                            effectiveLag = 0#
                        End If

                        If predTypeBySuccPred.Exists(linkKey) Then
                            linkType = CStr(predTypeBySuccPred(linkKey))
                        Else
                            linkType = "FS"
                        End If

                        Select Case linkType
                            Case "SS"
                                candidateFF = CDbl(rexStartById(CStr(succId))) - effectiveLag - CDbl(rexStartById(taskId))

                            Case "FF"
                                candidateFF = CDbl(rexFinishById(CStr(succId))) - effectiveLag - CDbl(rexFinishById(taskId))

                            Case Else
                                candidateFF = CDbl(rexStartById(CStr(succId))) - effectiveLag - 1 - CDbl(rexFinishById(taskId))
                        End Select

                        If Not HasValue(minFF) Then
                            minFF = candidateFF
                        ElseIf candidateFF < CDbl(minFF) Then
                            minFF = candidateFF
                        End If

                    End If

                End If

            Next succId

            ff = minFF
        End If

        If HasValue(ff) Then
            rexFreeFloatById(taskId) = ff
            If CDbl(ff) < 0 Then hasNegativeFloat = True
        End If

NextRexFreeFloatTask:
    Next idKey

    For Each idKey In validIds.Keys
        rowIndex = idToRow(CStr(idKey))

        If rexTotalFloatById.Exists(CStr(idKey)) Then
            outTF(rowIndex, 1) = rexTotalFloatById(CStr(idKey))
        End If

        If rexFreeFloatById.Exists(CStr(idKey)) Then
            outFF(rowIndex, 1) = rexFreeFloatById(CStr(idKey))
        End If
    Next idKey

    RollupFloatToParents idToRow, directChildrenById, parentIds, outTF, outFF

    For r = 1 To tblCalc.ListRows.Count
        If HasValue(outTF(r, 1)) Then
            If CDbl(outTF(r, 1)) <= 0 Then
                outCriticalREX(r, 1) = "CRITICAL"
            Else
                outCriticalREX(r, 1) = ""
            End If
        Else
            outCriticalREX(r, 1) = ""
        End If
    Next r

    tblCalc.ListColumns("Total Float REX").DataBodyRange.value = outTF
    tblCalc.ListColumns("Free Float REX").DataBodyRange.value = outFF
    tblCalc.ListColumns("Critical Path REX").DataBodyRange.value = outCriticalREX

    If hasNegativeFloat Then
        CalcEngine_AddOrShowConsoleMessage consoleMessages, "WARNING", _
            BiMsg( _
                "Float négatif détecté -> planning incohérent, corriger les données.", _
                "Negative float detected -> inconsistent schedule, fix inputs.")
    End If

End Sub


Public Sub Push_CriticalPathREX_Back_To_WBS()

    Dim wsWBS As Worksheet
    Dim wsCalc As Worksheet
    Dim tblWBS As ListObject
    Dim tblCalc As ListObject
    Dim mapWBS As Object
    Dim mapCalc As Object
    Dim calcById As Object
    Dim r As Long
    Dim id As String

    Set wsWBS = ThisWorkbook.Worksheets("WBS")
    Set wsCalc = ThisWorkbook.Worksheets("CALC")

    Set tblWBS = wsWBS.ListObjects("tbl_WBS")
    Set tblCalc = wsCalc.ListObjects("tbl_CALC")

    If tblWBS.DataBodyRange Is Nothing Then Exit Sub
    If tblCalc.DataBodyRange Is Nothing Then Exit Sub

    Set mapWBS = CreateObject("Scripting.Dictionary")
    Set mapCalc = CreateObject("Scripting.Dictionary")
    Set calcById = CreateObject("Scripting.Dictionary")

    For r = 1 To tblWBS.ListColumns.Count
        mapWBS(tblWBS.ListColumns(r).Name) = r
    Next r

    For r = 1 To tblCalc.ListColumns.Count
        mapCalc(tblCalc.ListColumns(r).Name) = r
    Next r

    If Not mapWBS.Exists("ID") Then Exit Sub
    If Not mapWBS.Exists("Critical Path REX") Then Exit Sub
    If Not mapCalc.Exists("ID") Then Exit Sub
    If Not mapCalc.Exists("Critical Path REX") Then Exit Sub

    For r = 1 To tblCalc.ListRows.Count
        id = Trim(CStr(tblCalc.DataBodyRange.Cells(r, mapCalc("ID")).value))
        If id <> "" Then
            calcById(id) = tblCalc.DataBodyRange.Cells(r, mapCalc("Critical Path REX")).value
        End If
    Next r

    BeginAuthorizedWBSWrite "Push_CriticalPathREX_Back_To_WBS", Array("Critical Path REX")

    For r = 1 To tblWBS.ListRows.Count
        id = Trim(CStr(tblWBS.DataBodyRange.Cells(r, mapWBS("ID")).value))
        If id <> "" Then
            If calcById.Exists(id) Then
                tblWBS.DataBodyRange.Cells(r, mapWBS("Critical Path REX")).value = calcById(id)
            Else
                tblWBS.DataBodyRange.Cells(r, mapWBS("Critical Path REX")).ClearContents
            End If
        End If
    Next r

    EndAuthorizedWBSWrite

End Sub


Private Sub DetectCyclesDFS(ByVal currentId As String, ByVal predsById As Object, ByVal idToRow As Object, ByVal state As Object, ByVal cycleNodes As Object)

    Dim pred As Variant

    state(currentId) = 1

    For Each pred In predsById(currentId)
        If idToRow.Exists(CStr(pred)) Then
            If state(CStr(pred)) = 0 Then
                DetectCyclesDFS CStr(pred), predsById, idToRow, state, cycleNodes
            ElseIf state(CStr(pred)) = 1 Then
                cycleNodes(currentId) = True
                cycleNodes(CStr(pred)) = True
            End If
        End If
    Next pred

    state(currentId) = 2

End Sub

Private Sub PropagateErrorToChildren(ByVal startId As String, ByVal childrenById As Object, ByVal errorDict As Object)

    Dim q As Collection
    Dim currentId As Variant
    Dim childId As Variant

    Set q = New Collection
    q.Add startId
    errorDict(startId) = True

    Do While q.Count > 0
        currentId = q(1)
        q.Remove 1

        If childrenById.Exists(CStr(currentId)) Then
            For Each childId In childrenById(CStr(currentId))
                If Not errorDict.Exists(CStr(childId)) Then
                    errorDict(CStr(childId)) = True
                    q.Add CStr(childId)
                End If
            Next childId
        End If
    Loop

End Sub

Private Sub RollupFloatToParents( _
    ByVal idToRow As Object, _
    ByVal directChildrenById As Object, _
    ByVal parentIds As Object, _
    ByRef outTF() As Variant, _
    ByRef outFF() As Variant)

    Dim changed As Boolean
    Dim key As Variant
    Dim id As String
    Dim rowIndex As Long
    Dim childId As Variant
    Dim maxPass As Long
    Dim passCount As Long

    maxPass = idToRow.Count
    passCount = 0

    Do
        changed = False
        passCount = passCount + 1

        For Each key In parentIds.Keys

            id = CStr(key)
            rowIndex = idToRow(id)

            Dim minTF As Variant
            Dim minFF As Variant

            minTF = Empty
            minFF = Empty

            For Each childId In directChildrenById(id)

                If Not IsEmpty(outTF(idToRow(CStr(childId)), 1)) Then
                    If IsEmpty(minTF) Then
                        minTF = outTF(idToRow(CStr(childId)), 1)
                    ElseIf CDbl(outTF(idToRow(CStr(childId)), 1)) < CDbl(minTF) Then
                        minTF = outTF(idToRow(CStr(childId)), 1)
                    End If
                End If

                If Not IsEmpty(outFF(idToRow(CStr(childId)), 1)) Then
                    If IsEmpty(minFF) Then
                        minFF = outFF(idToRow(CStr(childId)), 1)
                    ElseIf CDbl(outFF(idToRow(CStr(childId)), 1)) < CDbl(minFF) Then
                        minFF = outFF(idToRow(CStr(childId)), 1)
                    End If
                End If

            Next childId

            If Not IsEmpty(minTF) Then
                If IsEmpty(outTF(rowIndex, 1)) Then
                    outTF(rowIndex, 1) = minTF
                    changed = True
                ElseIf CDbl(outTF(rowIndex, 1)) <> CDbl(minTF) Then
                    outTF(rowIndex, 1) = minTF
                    changed = True
                End If
            End If

            If Not IsEmpty(minFF) Then
                If IsEmpty(outFF(rowIndex, 1)) Then
                    outFF(rowIndex, 1) = minFF
                    changed = True
                ElseIf CDbl(outFF(rowIndex, 1)) <> CDbl(minFF) Then
                    outFF(rowIndex, 1) = minFF
                    changed = True
                End If
            End If

        Next key

    Loop While changed And passCount <= maxPass

End Sub

Public Sub Push_Analytics_Back_To_WBS()

    Dim wsWBS As Worksheet
    Dim wsCalc As Worksheet
    Dim tblWBS As ListObject
    Dim tblCalc As ListObject

    Dim mapWBS As Object
    Dim mapCalc As Object
    Dim arrCalc As Variant
    Dim calcRowById As Object

    Dim r As Long
    Dim calcRow As Long
    Dim idVal As String

    On Error GoTo ErrHandler

    Set wsWBS = ThisWorkbook.Worksheets("WBS")
    Set wsCalc = ThisWorkbook.Worksheets("CALC")

    Set tblWBS = wsWBS.ListObjects("tbl_WBS")
    Set tblCalc = wsCalc.ListObjects("tbl_CALC")

    If tblWBS.DataBodyRange Is Nothing Then Exit Sub
    If tblCalc.DataBodyRange Is Nothing Then Exit Sub

    Set mapWBS = Core_BuildColumnMap_FromListObject(tblWBS)
    Set mapCalc = Core_BuildColumnMap_FromListObject(tblCalc)

    If Not mapWBS.Exists("ID") Then Err.Raise vbObjectError + 1301, "Push_Analytics_Back_To_WBS", "Missing column in tbl_WBS: ID"
    If Not mapCalc.Exists("ID") Then Err.Raise vbObjectError + 1302, "Push_Analytics_Back_To_WBS", "Missing column in tbl_CALC: ID"

    arrCalc = tblCalc.DataBodyRange.value

    Set calcRowById = CreateObject("Scripting.Dictionary")

    For r = 1 To UBound(arrCalc, 1)
        idVal = Trim$(CStr(arrCalc(r, mapCalc("ID"))))
        If idVal <> "" Then
            calcRowById(idVal) = r
        End If
    Next r

    Application.EnableEvents = False

    BeginAuthorizedWBSWrite "Push_Analytics_Back_To_WBS", Array( _
        "Critical Path", _
        "Longest Path", _
        "Critical Path REX", _
        "Total Float", _
        "Free Float", _
        "Total Float REX", _
        "Free Float REX", _
        "Deadline Float")

    For r = 1 To tblWBS.ListRows.Count

        idVal = Trim$(CStr(tblWBS.DataBodyRange.Cells(r, mapWBS("ID")).value))

        If idVal <> "" Then
            If calcRowById.Exists(idVal) Then

                calcRow = CLng(calcRowById(idVal))

                PushOneAnalyticsCellIfExists tblWBS, arrCalc, r, calcRow, mapWBS, mapCalc, "Critical Path"
                PushOneAnalyticsCellIfExists tblWBS, arrCalc, r, calcRow, mapWBS, mapCalc, "Longest Path"
                PushOneAnalyticsCellIfExists tblWBS, arrCalc, r, calcRow, mapWBS, mapCalc, "Critical Path REX"
                PushOneAnalyticsCellIfExists tblWBS, arrCalc, r, calcRow, mapWBS, mapCalc, "Total Float"
                PushOneAnalyticsCellIfExists tblWBS, arrCalc, r, calcRow, mapWBS, mapCalc, "Free Float"
                PushOneAnalyticsCellIfExists tblWBS, arrCalc, r, calcRow, mapWBS, mapCalc, "Total Float REX"
                PushOneAnalyticsCellIfExists tblWBS, arrCalc, r, calcRow, mapWBS, mapCalc, "Free Float REX"
                PushOneAnalyticsCellIfExists tblWBS, arrCalc, r, calcRow, mapWBS, mapCalc, "Deadline Float"

            Else

                ClearOneAnalyticsCellIfExists tblWBS, r, mapWBS, "Critical Path"
                ClearOneAnalyticsCellIfExists tblWBS, r, mapWBS, "Longest Path"
                ClearOneAnalyticsCellIfExists tblWBS, r, mapWBS, "Critical Path REX"
                ClearOneAnalyticsCellIfExists tblWBS, r, mapWBS, "Total Float"
                ClearOneAnalyticsCellIfExists tblWBS, r, mapWBS, "Free Float"
                ClearOneAnalyticsCellIfExists tblWBS, r, mapWBS, "Total Float REX"
                ClearOneAnalyticsCellIfExists tblWBS, r, mapWBS, "Free Float REX"
                ClearOneAnalyticsCellIfExists tblWBS, r, mapWBS, "Deadline Float"

            End If
        End If

    Next r

    EndAuthorizedWBSWrite

SafeExit:
    Application.EnableEvents = True
    Exit Sub

ErrHandler:
    On Error Resume Next
    EndAuthorizedWBSWrite
    Application.EnableEvents = True
    On Error GoTo 0
    Err.Raise Err.Number, "Push_Analytics_Back_To_WBS", Err.Description

End Sub

Private Sub PushOneAnalyticsCellIfExists( _
    ByVal tblWBS As ListObject, _
    ByRef arrCalc As Variant, _
    ByVal wbsRow As Long, _
    ByVal calcRow As Long, _
    ByVal mapWBS As Object, _
    ByVal mapCalc As Object, _
    ByVal fieldName As String)

    If Not IsAllowedAnalyticsPushField(fieldName) Then
        Err.Raise vbObjectError + 1303, "PushOneAnalyticsCellIfExists", _
            "Forbidden analytics WBS write attempted: " & fieldName
    End If

    If mapWBS.Exists(fieldName) Then
        If mapCalc.Exists(fieldName) Then
            tblWBS.DataBodyRange.Cells(wbsRow, mapWBS(fieldName)).value = _
                arrCalc(calcRow, mapCalc(fieldName))
        End If
    End If

End Sub

Private Sub ClearOneAnalyticsCellIfExists( _
    ByVal tblWBS As ListObject, _
    ByVal wbsRow As Long, _
    ByVal mapWBS As Object, _
    ByVal fieldName As String)

    If Not IsAllowedAnalyticsPushField(fieldName) Then
        Err.Raise vbObjectError + 1304, "ClearOneAnalyticsCellIfExists", _
            "Forbidden analytics WBS clear attempted: " & fieldName
    End If

    If mapWBS.Exists(fieldName) Then
        tblWBS.DataBodyRange.Cells(wbsRow, mapWBS(fieldName)).ClearContents
    End If

End Sub

Private Function IsAllowedAnalyticsPushField(ByVal fieldName As String) As Boolean

    Select Case fieldName
        Case "Critical Path", _
             "Longest Path", _
             "Critical Path REX", _
             "Total Float", _
             "Free Float", _
             "Total Float REX", _
             "Free Float REX", _
             "Deadline Float"
            IsAllowedAnalyticsPushField = True

        Case Else
            IsAllowedAnalyticsPushField = False
    End Select

End Function

Public Sub Validate_LogicLinksNetwork()

    Dim wsWBS As Worksheet
    Dim wsCalc As Worksheet
    Dim tblWBS As ListObject
    Dim tblLinks As ListObject

    Dim mapWBS As Object
    Dim mapLinks As Object
    Dim taskInfoById As Object
    Dim childrenByPred As Object
    Dim incomingCount As Object
    Dim allTaskIds As Object
    Dim state As Object
    Dim positionable As Object
    Dim indegree As Object
    Dim directChildrenById As Object
    Dim q As Collection
    Dim topo As Collection

    Dim errMissingPred As Object
    Dim errCycle As Object
    Dim errNotPositionable As Object
    Dim errLOEPred As Object

    Dim r As Long
    Dim i As Long
    Dim idVal As Variant
    Dim wbsVal As String
    Dim predId As String
    Dim succId As String
    Dim taskTypeVal As String

    Dim arrWBS As Variant
    Dim arrLinks As Variant

    Dim hasTaskType As Boolean
    Dim hasErrors As Boolean
    Dim isParent As Boolean

    Dim info As Object
    Dim colChildren As Collection
    Dim colSuccChildren As Collection
    Dim currentId As String
    Dim childId As Variant
    Dim parentWbs As String
    Dim parentId As String
    Dim wbsToId As Object

    On Error GoTo SafeExit

    Set wsWBS = ThisWorkbook.Worksheets("WBS")
    Set wsCalc = ThisWorkbook.Worksheets("CALC")

    Set tblWBS = wsWBS.ListObjects("tbl_WBS")
    Set tblLinks = wsCalc.ListObjects("tbl_LOGIC_LINKS")

    If tblWBS.DataBodyRange Is Nothing Then
        CalcEngine_ShowSingleConsoleMessage "WARNING", _
            BiMsg( _
                "tbl_WBS est vide.", _
                "tbl_WBS is empty.")
        Exit Sub
    End If

    Set mapWBS = CreateObject("Scripting.Dictionary")
    Set mapLinks = CreateObject("Scripting.Dictionary")
    Set taskInfoById = CreateObject("Scripting.Dictionary")
    Set childrenByPred = CreateObject("Scripting.Dictionary")
    Set incomingCount = CreateObject("Scripting.Dictionary")
    Set allTaskIds = CreateObject("Scripting.Dictionary")
    Set state = CreateObject("Scripting.Dictionary")
    Set positionable = CreateObject("Scripting.Dictionary")
    Set indegree = CreateObject("Scripting.Dictionary")
    Set directChildrenById = CreateObject("Scripting.Dictionary")
    Set wbsToId = CreateObject("Scripting.Dictionary")

    Set errMissingPred = CreateObject("Scripting.Dictionary")
    Set errCycle = CreateObject("Scripting.Dictionary")
    Set errNotPositionable = CreateObject("Scripting.Dictionary")
    Set errLOEPred = CreateObject("Scripting.Dictionary")

    For i = 1 To tblWBS.ListColumns.Count
        mapWBS(tblWBS.ListColumns(i).Name) = i
    Next i

    If tblLinks.ListColumns.Count > 0 Then
        For i = 1 To tblLinks.ListColumns.Count
            mapLinks(tblLinks.ListColumns(i).Name) = i
        Next i
    End If

    If Not mapWBS.Exists("ID") Then Err.Raise vbObjectError + 801, , "Missing column in tbl_WBS: ID"
    If Not mapWBS.Exists("WBS") Then Err.Raise vbObjectError + 802, , "Missing column in tbl_WBS: WBS"
    If Not mapWBS.Exists("Baseline Start") Then Err.Raise vbObjectError + 803, , "Missing column in tbl_WBS: Baseline Start"
    If Not mapWBS.Exists("Forecast Start") Then Err.Raise vbObjectError + 804, , "Missing column in tbl_WBS: Forecast Start"
    If Not mapWBS.Exists("Actual Start") Then Err.Raise vbObjectError + 805, , "Missing column in tbl_WBS: Actual Start"

    hasTaskType = mapWBS.Exists("Task Type")

    If Not mapLinks.Exists("Succ ID") Then Err.Raise vbObjectError + 806, , "Missing column in tbl_LOGIC_LINKS: Succ ID"
    If Not mapLinks.Exists("Pred ID") Then Err.Raise vbObjectError + 807, , "Missing column in tbl_LOGIC_LINKS: Pred ID"

    arrWBS = tblWBS.DataBodyRange.value

    For r = 1 To UBound(arrWBS, 1)

        idVal = Trim$(CStr(arrWBS(r, mapWBS("ID"))))
        wbsVal = NormalizeWBS(arrWBS(r, mapWBS("WBS")))

        If idVal <> "" Then

            If hasTaskType Then
                taskTypeVal = UCase$(Trim$(CStr(arrWBS(r, mapWBS("Task Type")))))
            Else
                taskTypeVal = ""
            End If

            Set info = CreateObject("Scripting.Dictionary")
            info("WBS") = wbsVal
            info("HasBaselineStart") = HasValue(arrWBS(r, mapWBS("Baseline Start")))
            info("HasForecastStart") = HasValue(arrWBS(r, mapWBS("Forecast Start")))
            info("HasActualStart") = HasValue(arrWBS(r, mapWBS("Actual Start")))
            info("Task Type") = taskTypeVal

            If taskInfoById.Exists(CStr(idVal)) Then
                Set taskInfoById(CStr(idVal)) = info
            Else
                taskInfoById.Add CStr(idVal), info
            End If

            allTaskIds(CStr(idVal)) = True
            wbsToId(wbsVal) = CStr(idVal)

            If Not childrenByPred.Exists(CStr(idVal)) Then
                Set colChildren = New Collection
                childrenByPred.Add CStr(idVal), colChildren
            End If

            If Not incomingCount.Exists(CStr(idVal)) Then incomingCount(CStr(idVal)) = 0
            If Not directChildrenById.Exists(CStr(idVal)) Then
                Set colChildren = New Collection
                directChildrenById.Add CStr(idVal), colChildren
            End If

        End If

    Next r

    For Each idVal In taskInfoById.Keys

        wbsVal = CStr(taskInfoById(CStr(idVal))("WBS"))
        parentWbs = GetParentWBS(wbsVal)

        If parentWbs <> "" Then
            If wbsToId.Exists(parentWbs) Then
                parentId = CStr(wbsToId(parentWbs))
                directChildrenById(parentId).Add CStr(idVal)
            End If
        End If

    Next idVal

    If Not tblLinks.DataBodyRange Is Nothing Then

        arrLinks = tblLinks.DataBodyRange.value

        For r = 1 To UBound(arrLinks, 1)

            succId = Trim$(CStr(arrLinks(r, mapLinks("Succ ID"))))
            predId = Trim$(CStr(arrLinks(r, mapLinks("Pred ID"))))

            If succId <> "" Then
                If Not allTaskIds.Exists(succId) Then allTaskIds(succId) = True

                If Not childrenByPred.Exists(succId) Then
                    Set colSuccChildren = New Collection
                    childrenByPred.Add succId, colSuccChildren
                End If

                If Not incomingCount.Exists(succId) Then incomingCount(succId) = 0
            End If

            If predId = "" Then
                If succId <> "" Then errMissingPred(succId) = True
            Else
                If Not taskInfoById.Exists(predId) Then
                    If succId <> "" Then errMissingPred(succId) = True
                Else
                    If succId <> "" Then
                        childrenByPred(predId).Add succId
                        incomingCount(succId) = incomingCount(succId) + 1
                    End If

                    If hasTaskType Then
                        If UCase$(Trim$(CStr(taskInfoById(predId)("Task Type")))) = "LEVEL OF EFFORT" Then
                            If succId <> "" Then errLOEPred(succId) = True
                        End If
                    End If
                End If
            End If

        Next r

    End If

    For Each idVal In allTaskIds.Keys
        state(CStr(idVal)) = 0
    Next idVal

    For Each idVal In allTaskIds.Keys
        If state(CStr(idVal)) = 0 Then
            DetectLogicCyclesDFS CStr(idVal), childrenByPred, state, errCycle
        End If
    Next idVal

    Set q = New Collection
    Set topo = New Collection

    For Each idVal In allTaskIds.Keys
        indegree(CStr(idVal)) = incomingCount(CStr(idVal))
        If indegree(CStr(idVal)) = 0 Then q.Add CStr(idVal)

        If taskInfoById.Exists(CStr(idVal)) Then
            positionable(CStr(idVal)) = _
                CBool(taskInfoById(CStr(idVal))("HasActualStart")) Or _
                CBool(taskInfoById(CStr(idVal))("HasForecastStart")) Or _
                CBool(taskInfoById(CStr(idVal))("HasBaselineStart"))
        Else
            positionable(CStr(idVal)) = False
        End If
    Next idVal

    Do While q.Count > 0

        currentId = CStr(q(1))
        q.Remove 1
        topo.Add currentId

        If childrenByPred.Exists(currentId) Then
            For Each childId In childrenByPred(currentId)

                If positionable.Exists(currentId) Then
                    If CBool(positionable(currentId)) Then
                        positionable(CStr(childId)) = True
                    End If
                End If

                indegree(CStr(childId)) = indegree(CStr(childId)) - 1
                If indegree(CStr(childId)) = 0 Then q.Add CStr(childId)

            Next childId
        End If

    Loop

    For Each idVal In taskInfoById.Keys

        If directChildrenById.Exists(CStr(idVal)) Then
            If directChildrenById(CStr(idVal)).Count > 0 Then

                For Each childId In directChildrenById(CStr(idVal))
                    If positionable.Exists(CStr(childId)) Then
                        If CBool(positionable(CStr(childId))) Then
                            positionable(CStr(idVal)) = True
                            Exit For
                        End If
                    End If
                Next childId

            End If
        End If

    Next idVal

    For Each idVal In allTaskIds.Keys

        If taskInfoById.Exists(CStr(idVal)) Then

            isParent = False
            If directChildrenById.Exists(CStr(idVal)) Then
                isParent = (directChildrenById(CStr(idVal)).Count > 0)
            End If

            If Not isParent Then
                If Not CBool(positionable(CStr(idVal))) Then
                    errNotPositionable(CStr(idVal)) = True
                End If
            End If

        End If

    Next idVal

    hasErrors = _
        (errMissingPred.Count > 0) Or _
        (errCycle.Count > 0) Or _
        (errNotPositionable.Count > 0) Or _
        (errLOEPred.Count > 0)

    If hasErrors Then

        If errMissingPred.Count > 0 Then
            ShowLogicLinksErrorMessages errMissingPred, taskInfoById, _
                "Prédécesseur introuvable dans tbl_LOGIC_LINKS", "vérifier la colonne Predecessors WBS", _
                "Missing predecessor in tbl_LOGIC_LINKS", "check the Predecessors WBS column"
        End If

        If errLOEPred.Count > 0 Then
            ShowLogicLinksErrorMessages errLOEPred, taskInfoById, _
                "Un Level of Effort ne peut pas être prédécesseur", "corriger la logique de liaison", _
                "A Level of Effort cannot be used as predecessor", "fix the logical relationship"
        End If

        If errCycle.Count > 0 Then
            ShowLogicLinksErrorMessages errCycle, taskInfoById, _
                "Boucle logique détectée", "corriger les relations de dépendance", _
                "Logical cycle detected", "fix the dependency relationships"
        End If

        If errNotPositionable.Count > 0 Then
            ShowLogicLinksErrorMessages errNotPositionable, taskInfoById, _
                "Tâche ou chaîne non positionnable", "ajouter une date de début ou une logique amont ancrée", _
                "Task or chain cannot be positioned", "add a start date or an anchored upstream logic"
        End If

        Exit Sub
    End If

    CalcEngine_ShowSingleConsoleMessage "INFO", _
        BiMsg( _
            "Validation réseau OK." & vbCrLf & _
            "-> aucun prédécesseur manquant, aucun cycle, aucune tâche non positionnable détectée.", _
            "Network validation OK." & vbCrLf & _
            "-> no missing predecessor, no cycle, no non-positionable task detected.")

SafeExit:
    If Err.Number <> 0 Then
        CalcEngine_ShowSingleConsoleMessage "STOP", _
            BiMsg( _
                "Erreur VBA dans Validate_LogicLinksNetwork : " & Err.Description, _
                "VBA error in Validate_LogicLinksNetwork: " & Err.Description)
    End If

End Sub


Private Sub DetectLogicCyclesDFS( _
    ByVal currentId As String, _
    ByVal childrenByPred As Object, _
    ByVal state As Object, _
    ByVal cycleIds As Object)

    Dim childId As Variant

    state(currentId) = 1

    If childrenByPred.Exists(currentId) Then
        For Each childId In childrenByPred(currentId)

            If state(CStr(childId)) = 0 Then
                DetectLogicCyclesDFS CStr(childId), childrenByPred, state, cycleIds
            ElseIf state(CStr(childId)) = 1 Then
                cycleIds(currentId) = True
                cycleIds(CStr(childId)) = True
            End If

        Next childId
    End If

    state(currentId) = 2

End Sub

Private Sub ShowLogicLinksErrorMessages( _
    ByVal idsDict As Object, _
    ByVal taskInfoById As Object, _
    ByVal frProblem As String, _
    ByVal frAction As String, _
    ByVal enProblem As String, _
    ByVal enAction As String)

    Dim idsLine As String
    Dim wbsLine As String
    Dim msgText As String

    If idsDict Is Nothing Then Exit Sub
    If idsDict.Count = 0 Then Exit Sub

    idsLine = BuildInlineList_LogicLinks(idsDict, 20)
    wbsLine = BuildInlineWBSList_LogicLinks(idsDict, taskInfoById, 20)

    msgText = _
        frProblem & vbCrLf & _
        "-> " & frAction & vbCrLf & vbCrLf & _
        "IDs : " & idsLine & vbCrLf & _
        "WBS : " & wbsLine & vbCrLf & vbCrLf & _
        enProblem & vbCrLf & _
        "-> " & enAction & vbCrLf & vbCrLf & _
        "IDs: " & idsLine & vbCrLf & _
        "WBS: " & wbsLine

    CalcEngine_ShowSingleConsoleMessage "STOP", msgText

End Sub



Private Function BuildInlineList_LogicLinks(ByVal idsDict As Object, ByVal maxItems As Long) As String

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

    BuildInlineList_LogicLinks = result

End Function


Private Function BuildInlineWBSList_LogicLinks( _
    ByVal idsDict As Object, _
    ByVal taskInfoById As Object, _
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

            If taskInfoById.Exists(CStr(key)) Then
                itemText = CStr(taskInfoById(CStr(key))("WBS"))
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

    BuildInlineWBSList_LogicLinks = result

End Function
'=================================================
' HELPER 2
' Build an inline message on IDs + WBS from current engine dictionaries.
'=================================================
Private Sub ShowCalcUnsupportedLinkTypeMessages( _
    ByVal idsDict As Object, _
    ByVal idToWbs As Object)

    If idsDict Is Nothing Then Exit Sub
    If idsDict.Count = 0 Then Exit Sub

    CalcEngine_ShowSingleConsoleMessage "STOP", _
        BuildGroupedMessage( _
            idsDict, _
            idToWbs, _
            "Type de lien non encore supporté par le moteur", _
            "à ce stade, seuls les liens FS sont calculés", _
            "Link type not yet supported by the engine", _
            "at this stage, only FS links are calculated")

End Sub


'=================================================
' HELPER 3
' Validate presence of tbl_LOGIC_LINKS.
'=================================================
Private Function GetLogicLinksTable() As ListObject

    Dim wsCalc As Worksheet

    On Error GoTo SafeExit

    Set wsCalc = ThisWorkbook.Worksheets("CALC")
    Set GetLogicLinksTable = wsCalc.ListObjects("tbl_LOGIC_LINKS")
    Exit Function

SafeExit:
    Set GetLogicLinksTable = Nothing

End Function


'=================================================
' HELPER 4
' Rebuild network messages for missing / unsupported link types
' before date calculation starts.
'=================================================
Private Sub ShowLogicLinksStructuralMessages_Stage5A( _
    ByVal errMissingPred As Object, _
    ByVal errUnsupportedLinkType As Object, _
    ByVal idToWbs As Object)

    If errMissingPred.Count > 0 Then
        ShowCalcErrorMessages errMissingPred, idToWbs, _
            "Prédécesseur introuvable dans tbl_LOGIC_LINKS", "vérifier la table des liens logiques", _
            "Missing predecessor in tbl_LOGIC_LINKS", "check the logical links table"
    End If

    If errUnsupportedLinkType.Count > 0 Then
        ShowCalcUnsupportedLinkTypeMessages errUnsupportedLinkType, idToWbs
    End If

End Sub

Public Sub Test_WBS_UnauthorizedWrite()

    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim oldVal As Variant
    Dim consoleMessages As Collection

    On Error GoTo SafeExit

    Set ws = ThisWorkbook.Worksheets("WBS")
    Set tbl = ws.ListObjects("tbl_WBS")

    If tbl.DataBodyRange Is Nothing Then Exit Sub

    BeginMacroRun "Test_WBS_UnauthorizedWrite"

    oldVal = tbl.ListColumns("Calculated Start").DataBodyRange.Cells(1, 1).value

    If IsDate(oldVal) Then
        tbl.ListColumns("Calculated Start").DataBodyRange.Cells(1, 1).value = CDate(oldVal) + 1
    Else
        tbl.ListColumns("Calculated Start").DataBodyRange.Cells(1, 1).value = Date + 7
    End If

    If IsMacroAbortRequested() Then
        ShowAbortMessageOnce
    Else
        Set consoleMessages = New Collection

        CalcBridge_AddConsoleMessage consoleMessages, "WARNING", _
            "FR:" & vbCrLf & _
            "KO test : aucune demande d'abort n'a été détectée." & vbCrLf & vbCrLf & _
            "EN:" & vbCrLf & _
            "KO test: no abort request was detected."

        CalcBridge_ShowPlanningConsole consoleMessages
    End If

SafeExit:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    EndMacroRun

End Sub

'=================================================
' HELPER
' Build lag map from tbl_LOGIC_LINKS
'
' Key format:
'   SuccID|PredID
'
' Temporary rule:
' - FS supported
' - blank lag = 0
' - SS / FF ignored here because they are still blocked upstream
'=================================================
Private Function BuildPredLagMapFromLogicLinks() As Object

    Dim wsCalc As Worksheet
    Dim tblLinks As ListObject
    Dim mapLinks As Object
    Dim arrLinks As Variant
    Dim d As Object

    Dim r As Long
    Dim i As Long

    Dim succId As String
    Dim predId As String
    Dim linkType As String
    Dim linkKey As String
    Dim lagVal As Variant

    Set d = CreateObject("Scripting.Dictionary")
    Set mapLinks = CreateObject("Scripting.Dictionary")

    On Error GoTo SafeExit

    Set wsCalc = ThisWorkbook.Worksheets("CALC")
    Set tblLinks = wsCalc.ListObjects("tbl_LOGIC_LINKS")

    If tblLinks Is Nothing Then
        Set BuildPredLagMapFromLogicLinks = d
        Exit Function
    End If

    For i = 1 To tblLinks.ListColumns.Count
        mapLinks(tblLinks.ListColumns(i).Name) = i
    Next i

    If Not mapLinks.Exists("Succ ID") Then GoTo SafeExit
    If Not mapLinks.Exists("Pred ID") Then GoTo SafeExit
    If Not mapLinks.Exists("Link Type") Then GoTo SafeExit
    If Not mapLinks.Exists("Lag") Then GoTo SafeExit

    If tblLinks.DataBodyRange Is Nothing Then
        Set BuildPredLagMapFromLogicLinks = d
        Exit Function
    End If

    arrLinks = tblLinks.DataBodyRange.value

    For r = 1 To UBound(arrLinks, 1)

        succId = Trim$(CStr(arrLinks(r, mapLinks("Succ ID"))))
        predId = Trim$(CStr(arrLinks(r, mapLinks("Pred ID"))))
        linkType = UCase$(Trim$(CStr(arrLinks(r, mapLinks("Link Type")))))
        lagVal = arrLinks(r, mapLinks("Lag"))

        If succId = "" Or predId = "" Then GoTo NextRow

        If linkType = "" Then linkType = "FS"

        Select Case linkType
            Case "FS", "SS", "FF"
                linkKey = succId & "|" & predId

                If IsNumeric(lagVal) Then
                    d(linkKey) = CDbl(lagVal)
                Else
                    d(linkKey) = 0#
                End If
        End Select

NextRow:
    Next r

SafeExit:
    Set BuildPredLagMapFromLogicLinks = d

End Function

'=================================================
' HELPER
' Build link-type map from tbl_LOGIC_LINKS
'
' Key format:
'   SuccID|PredID
'
' Default type:
' - blank = FS
'=================================================
Private Function BuildPredTypeMapFromLogicLinks() As Object

    Dim wsCalc As Worksheet
    Dim tblLinks As ListObject
    Dim mapLinks As Object
    Dim arrLinks As Variant
    Dim d As Object

    Dim r As Long
    Dim i As Long

    Dim succId As String
    Dim predId As String
    Dim linkType As String
    Dim linkKey As String

    Set d = CreateObject("Scripting.Dictionary")
    Set mapLinks = CreateObject("Scripting.Dictionary")

    On Error GoTo SafeExit

    Set wsCalc = ThisWorkbook.Worksheets("CALC")
    Set tblLinks = wsCalc.ListObjects("tbl_LOGIC_LINKS")

    If tblLinks Is Nothing Then
        Set BuildPredTypeMapFromLogicLinks = d
        Exit Function
    End If

    For i = 1 To tblLinks.ListColumns.Count
        mapLinks(tblLinks.ListColumns(i).Name) = i
    Next i

    If Not mapLinks.Exists("Succ ID") Then GoTo SafeExit
    If Not mapLinks.Exists("Pred ID") Then GoTo SafeExit
    If Not mapLinks.Exists("Link Type") Then GoTo SafeExit

    If tblLinks.DataBodyRange Is Nothing Then
        Set BuildPredTypeMapFromLogicLinks = d
        Exit Function
    End If

    arrLinks = tblLinks.DataBodyRange.value

    For r = 1 To UBound(arrLinks, 1)

        succId = Trim$(CStr(arrLinks(r, mapLinks("Succ ID"))))
        predId = Trim$(CStr(arrLinks(r, mapLinks("Pred ID"))))
        linkType = UCase$(Trim$(CStr(arrLinks(r, mapLinks("Link Type")))))

        If succId = "" Or predId = "" Then GoTo NextRow
        If linkType = "" Then linkType = "FS"

        Select Case linkType
            Case "FS", "SS", "FF"
                linkKey = succId & "|" & predId
                d(linkKey) = linkType
        End Select

NextRow:
    Next r

SafeExit:
    Set BuildPredTypeMapFromLogicLinks = d

End Function

'=================================================
' HELPER
' Build predecessor collections from tbl_LOGIC_LINKS
'
' Stage FF:
' - FS supported
' - SS supported
' - FF supported for dates
'
' Outputs:
' - predsById
' - childrenById
' - predLagBySuccPred  : key = SuccID|PredID
' - predTypeBySuccPred : key = SuccID|PredID
'=================================================
Private Function BuildPredsFromLogicLinks_FS_SS_FF_WithLagAndType_ExpandedToLeafs( _
    ByVal tblLinks As ListObject, _
    ByVal validIds As Object, _
    ByVal predsById As Object, _
    ByVal childrenById As Object, _
    ByVal directChildrenById As Object, _
    ByVal predLagBySuccPred As Object, _
    ByVal predTypeBySuccPred As Object, _
    ByVal structuralErrors As Object, _
    ByVal errMissingPred As Object, _
    ByVal errUnsupportedLinkType As Object, _
    ByRef hasStructuralError As Boolean) As Boolean

    Dim mapLinks As Object
    Dim arrLinks As Variant
    Dim r As Long
    Dim i As Long

    Dim rawSuccId As String
    Dim rawPredId As String
    Dim linkType As String
    Dim linkLag As Variant
    Dim linkKey As String

    Dim expandedLeafPreds As Collection
    Dim expandedLeafSuccs As Collection
    Dim leafPredId As Variant
    Dim leafSuccId As Variant

    BuildPredsFromLogicLinks_FS_SS_FF_WithLagAndType_ExpandedToLeafs = False

    Set mapLinks = CreateObject("Scripting.Dictionary")

    If tblLinks Is Nothing Then
        hasStructuralError = True
        Exit Function
    End If

    For i = 1 To tblLinks.ListColumns.Count
        mapLinks(tblLinks.ListColumns(i).Name) = i
    Next i

    If Not mapLinks.Exists("Succ ID") Then
        hasStructuralError = True
        Exit Function
    End If

    If Not mapLinks.Exists("Pred ID") Then
        hasStructuralError = True
        Exit Function
    End If

    If Not mapLinks.Exists("Link Type") Then
        hasStructuralError = True
        Exit Function
    End If

    If Not mapLinks.Exists("Lag") Then
        hasStructuralError = True
        Exit Function
    End If

    If tblLinks.DataBodyRange Is Nothing Then
        BuildPredsFromLogicLinks_FS_SS_FF_WithLagAndType_ExpandedToLeafs = True
        Exit Function
    End If

    arrLinks = tblLinks.DataBodyRange.value

    For r = 1 To UBound(arrLinks, 1)

        rawSuccId = Trim$(CStr(arrLinks(r, mapLinks("Succ ID"))))
        rawPredId = Trim$(CStr(arrLinks(r, mapLinks("Pred ID"))))
        linkType = UCase$(Trim$(CStr(arrLinks(r, mapLinks("Link Type")))))
        linkLag = arrLinks(r, mapLinks("Lag"))

        If rawSuccId = "" Then GoTo NextLinkRow
        If rawPredId = "" Then
            structuralErrors(rawSuccId) = True
            errMissingPred(rawSuccId) = True
            hasStructuralError = True
            GoTo NextLinkRow
        End If

        If linkType = "" Then linkType = "FS"

        If linkType <> "FS" And linkType <> "SS" And linkType <> "FF" Then
            structuralErrors(rawSuccId) = True
            errUnsupportedLinkType(rawSuccId) = True
            hasStructuralError = True
            GoTo NextLinkRow
        End If

        Set expandedLeafPreds = GetLeafDescendantsForCalcNetwork(rawPredId, directChildrenById, validIds)
        Set expandedLeafSuccs = GetLeafDescendantsForCalcNetwork(rawSuccId, directChildrenById, validIds)

        If expandedLeafPreds Is Nothing Then
            structuralErrors(rawSuccId) = True
            errMissingPred(rawSuccId) = True
            hasStructuralError = True
            GoTo NextLinkRow
        End If

        If expandedLeafSuccs Is Nothing Then
            structuralErrors(rawSuccId) = True
            errMissingPred(rawSuccId) = True
            hasStructuralError = True
            GoTo NextLinkRow
        End If

        If expandedLeafPreds.Count = 0 Then
            structuralErrors(rawSuccId) = True
            errMissingPred(rawSuccId) = True
            hasStructuralError = True
            GoTo NextLinkRow
        End If

        If expandedLeafSuccs.Count = 0 Then
            structuralErrors(rawSuccId) = True
            errMissingPred(rawSuccId) = True
            hasStructuralError = True
            GoTo NextLinkRow
        End If

        For Each leafSuccId In expandedLeafSuccs
            For Each leafPredId In expandedLeafPreds

                predsById(CStr(leafSuccId)).Add CStr(leafPredId)
                childrenById(CStr(leafPredId)).Add CStr(leafSuccId)

                linkKey = CStr(leafSuccId) & "|" & CStr(leafPredId)

                If IsNumeric(linkLag) Then
                    predLagBySuccPred(linkKey) = CDbl(linkLag)
                Else
                    predLagBySuccPred(linkKey) = 0#
                End If

                predTypeBySuccPred(linkKey) = linkType

            Next leafPredId
        Next leafSuccId

NextLinkRow:
    Next r

    BuildPredsFromLogicLinks_FS_SS_FF_WithLagAndType_ExpandedToLeafs = True

End Function


Private Function GetLeafDescendantsForCalcNetwork( _
    ByVal startId As String, _
    ByVal directChildrenById As Object, _
    ByVal validIds As Object) As Collection

    Dim result As Collection
    Dim childId As Variant
    Dim childLeafs As Collection
    Dim leafId As Variant

    Set result = New Collection

    If validIds.Exists(startId) Then
        result.Add startId
        Set GetLeafDescendantsForCalcNetwork = result
        Exit Function
    End If

    If Not directChildrenById.Exists(startId) Then
        Set GetLeafDescendantsForCalcNetwork = result
        Exit Function
    End If

    If directChildrenById(startId).Count = 0 Then
        Set GetLeafDescendantsForCalcNetwork = result
        Exit Function
    End If

    For Each childId In directChildrenById(startId)

        Set childLeafs = GetLeafDescendantsForCalcNetwork(CStr(childId), directChildrenById, validIds)

        If Not childLeafs Is Nothing Then
            For Each leafId In childLeafs
                result.Add CStr(leafId)
            Next leafId
        End If

    Next childId

    Set GetLeafDescendantsForCalcNetwork = result

End Function

Private Function BuildUpstreamViolationMessages( _
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
            If idToWbs.Exists(CStr(key)) Then
                wbsVal = CStr(idToWbs(CStr(key)))
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

    BuildUpstreamViolationMessages = result

End Function

Private Sub ShowUpstreamViolationMessages( _
    ByVal idsDict As Object, _
    ByVal idToWbs As Object, _
    ByVal frProblem As String, _
    ByVal frAction As String, _
    ByVal enProblem As String, _
    ByVal enAction As String)

    Dim itemsLine As String
    Dim msgText As String
    Dim consoleMessages As Collection

    If idsDict Is Nothing Then Exit Sub
    If idsDict.Count = 0 Then Exit Sub

    itemsLine = BuildUpstreamViolationMessages(idsDict, idToWbs, 20)

    msgText = _
        "FR:" & vbCrLf & _
        frProblem & vbCrLf & _
        "-> " & frAction & vbCrLf & vbCrLf & _
        "Tâches : " & itemsLine & vbCrLf & vbCrLf & _
        "EN:" & vbCrLf & _
        enProblem & vbCrLf & _
        "-> " & enAction & vbCrLf & vbCrLf & _
        "Tasks: " & itemsLine

    Set consoleMessages = New Collection
    CalcBridge_AddConsoleMessage consoleMessages, "STOP", msgText
    CalcBridge_ShowPlanningConsole consoleMessages

End Sub


Public Function CalcEngine_HasBlockingErrorsForState() As Boolean

    Dim wsCalc As Worksheet
    Dim tblCalc As ListObject
    Dim mapCalc As Object
    Dim arr As Variant
    Dim r As Long

    On Error GoTo FailSafe

    CalcEngine_HasBlockingErrorsForState = True

    Set wsCalc = ThisWorkbook.Worksheets("CALC")
    Set tblCalc = wsCalc.ListObjects("tbl_CALC")

    If tblCalc.DataBodyRange Is Nothing Then Exit Function

    Set mapCalc = CreateObject("Scripting.Dictionary")

    For r = 1 To tblCalc.ListColumns.Count
        mapCalc(tblCalc.ListColumns(r).Name) = r
    Next r

    If Not mapCalc.Exists("Error flag") Then Exit Function

    arr = tblCalc.DataBodyRange.value

    For r = 1 To UBound(arr, 1)
        If UCase$(Trim$(CStr(arr(r, mapCalc("Error flag"))))) = "ERROR" Then
            CalcEngine_HasBlockingErrorsForState = True
            Exit Function
        End If
    Next r

    CalcEngine_HasBlockingErrorsForState = False
    Exit Function

FailSafe:
    CalcEngine_HasBlockingErrorsForState = True

End Function

Private Function IsActualFinishedInCalcArray( _
    ByRef dataArr As Variant, _
    ByVal rowIndex As Long, _
    ByVal mapCalc As Object) As Boolean

    If mapCalc Is Nothing Then Exit Function
    If Not mapCalc.Exists("Actual Finish") Then Exit Function

    IsActualFinishedInCalcArray = HasValue(GetCellValue(dataArr(rowIndex, mapCalc("Actual Finish"))))

End Function

Private Function HasUsableCalcDates( _
    ByRef dataArr As Variant, _
    ByVal rowIndex As Long, _
    ByVal mapCalc As Object) As Boolean

    If mapCalc Is Nothing Then Exit Function
    If Not mapCalc.Exists("Calculated Start") Then Exit Function
    If Not mapCalc.Exists("Calculated Finish") Then Exit Function

    HasUsableCalcDates = _
        HasValue(GetCellValue(dataArr(rowIndex, mapCalc("Calculated Start")))) And _
        HasValue(GetCellValue(dataArr(rowIndex, mapCalc("Calculated Finish"))))

End Function

Private Sub CalcEngine_ShowSingleConsoleMessage( _
    ByVal msgType As String, _
    ByVal msgText As String)

    CalcBridge_AddOrShowRawConsoleMessage Nothing, msgType, msgText

End Sub

Private Sub CalcEngine_AddOrShowConsoleMessage( _
    ByVal consoleMessages As Collection, _
    ByVal msgType As String, _
    ByVal msgText As String)

    CalcBridge_AddOrShowRawConsoleMessage consoleMessages, msgType, msgText

End Sub

Private Function BuildCurrentNetworkFinishById( _
    ByRef dataArr As Variant, _
    ByVal mapCalc As Object, _
    ByVal idToRow As Object, _
    ByVal childrenById As Object, _
    ByVal validIds As Object) As Object

    Dim finishById As Object
    Dim componentIds As Collection
    Dim visited As Object
    Dim idKey As Variant
    Dim compId As Variant
    Dim rowIndex As Long
    Dim finishVal As Variant
    Dim componentFinish As Variant

    Set finishById = CreateObject("Scripting.Dictionary")
    Set visited = CreateObject("Scripting.Dictionary")

    For Each idKey In validIds.Keys

        If Not visited.Exists(CStr(idKey)) Then

            Set componentIds = GetUndirectedNetworkComponent(CStr(idKey), childrenById, validIds, visited)
            componentFinish = Empty

            For Each compId In componentIds
                rowIndex = idToRow(CStr(compId))
                finishVal = GetCellValue(dataArr(rowIndex, mapCalc("Calculated Finish")))

                If HasValue(finishVal) Then
                    If Not HasValue(componentFinish) Then
                        componentFinish = finishVal
                    ElseIf CDbl(finishVal) > CDbl(componentFinish) Then
                        componentFinish = finishVal
                    End If
                End If
            Next compId

            If HasValue(componentFinish) Then
                For Each compId In componentIds
                    finishById(CStr(compId)) = componentFinish
                Next compId
            End If

        End If

    Next idKey

    Set BuildCurrentNetworkFinishById = finishById

End Function

Private Function BuildRexNetworkFinishById( _
    ByVal rexFinishById As Object, _
    ByVal childrenById As Object, _
    ByVal validIds As Object) As Object

    Dim finishById As Object
    Dim componentIds As Collection
    Dim visited As Object
    Dim idKey As Variant
    Dim compId As Variant
    Dim componentFinish As Variant

    Set finishById = CreateObject("Scripting.Dictionary")
    Set visited = CreateObject("Scripting.Dictionary")

    For Each idKey In validIds.Keys

        If Not visited.Exists(CStr(idKey)) Then

            Set componentIds = GetUndirectedNetworkComponent(CStr(idKey), childrenById, validIds, visited)
            componentFinish = Empty

            For Each compId In componentIds
                If rexFinishById.Exists(CStr(compId)) Then
                    If Not HasValue(componentFinish) Then
                        componentFinish = rexFinishById(CStr(compId))
                    ElseIf CDbl(rexFinishById(CStr(compId))) > CDbl(componentFinish) Then
                        componentFinish = rexFinishById(CStr(compId))
                    End If
                End If
            Next compId

            If HasValue(componentFinish) Then
                For Each compId In componentIds
                    finishById(CStr(compId)) = componentFinish
                Next compId
            End If

        End If

    Next idKey

    Set BuildRexNetworkFinishById = finishById

End Function

Private Function GetUndirectedNetworkComponent( _
    ByVal startId As String, _
    ByVal childrenById As Object, _
    ByVal validIds As Object, _
    ByVal visited As Object) As Collection

    Dim componentIds As Collection
    Dim queue As Collection
    Dim reverseChildrenById As Object
    Dim currentId As String
    Dim nextId As Variant

    Set componentIds = New Collection
    Set queue = New Collection
    Set reverseChildrenById = BuildReverseChildrenMap(childrenById, validIds)

    visited(startId) = True
    queue.Add startId

    Do While queue.Count > 0

        currentId = CStr(queue(1))
        queue.Remove 1
        componentIds.Add currentId

        If childrenById.Exists(currentId) Then
            For Each nextId In childrenById(currentId)
                If validIds.Exists(CStr(nextId)) Then
                    If Not visited.Exists(CStr(nextId)) Then
                        visited(CStr(nextId)) = True
                        queue.Add CStr(nextId)
                    End If
                End If
            Next nextId
        End If

        If reverseChildrenById.Exists(currentId) Then
            For Each nextId In reverseChildrenById(currentId)
                If validIds.Exists(CStr(nextId)) Then
                    If Not visited.Exists(CStr(nextId)) Then
                        visited(CStr(nextId)) = True
                        queue.Add CStr(nextId)
                    End If
                End If
            Next nextId
        End If

    Loop

    Set GetUndirectedNetworkComponent = componentIds

End Function

Private Function BuildReverseChildrenMap( _
    ByVal childrenById As Object, _
    ByVal validIds As Object) As Object

    Dim reverseMap As Object
    Dim idKey As Variant
    Dim childId As Variant
    Dim col As Collection

    Set reverseMap = CreateObject("Scripting.Dictionary")

    For Each idKey In validIds.Keys
        Set reverseMap(CStr(idKey)) = New Collection
    Next idKey

    If childrenById Is Nothing Then
        Set BuildReverseChildrenMap = reverseMap
        Exit Function
    End If

    For Each idKey In childrenById.Keys
        If validIds.Exists(CStr(idKey)) Then
            For Each childId In childrenById(CStr(idKey))
                If validIds.Exists(CStr(childId)) Then
                    If Not reverseMap.Exists(CStr(childId)) Then
                        Set col = New Collection
                        reverseMap.Add CStr(childId), col
                    End If
                    reverseMap(CStr(childId)).Add CStr(idKey)
                End If
            Next childId
        End If
    Next idKey

    Set BuildReverseChildrenMap = reverseMap

End Function

