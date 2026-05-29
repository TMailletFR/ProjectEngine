Attribute VB_Name = "mod_CalcEngineCoreBridge"
Option Explicit
Public Sub Run_Calc_Engine_CoreBridge(Optional ByVal forceFullRecalcOverride As Boolean = False)

    Const PARTIAL_RECALC_MAX_IMPACT_RATIO As Double = 0.3

    Dim wsCalc As Worksheet
    Dim wsWBS As Worksheet
    Dim tblCalc As ListObject
    Dim tblWBS As ListObject

    Dim mapCalc As Object
    Dim mapWBS As Object
    Dim dataArr As Variant
    Dim linksBySuccId As Object

    Dim changedIds As Object
    Dim forceFullRecalc As Boolean

    Dim succByPred As Object
    Dim parentById As Object
    Dim impactedIds As Object
    Dim usePartialWrite As Boolean

    Dim totalRows As Long
    Dim impactedCount As Long
    Dim impactRatio As Double

    Dim runAnalytics As Boolean
    Dim consoleMessages As Collection

    On Error GoTo ErrHandler

    ThisWorkbook.Init_AppEvents
    BeginMacroRun "Run_Calc_Engine_CoreBridge"

    Set consoleMessages = New Collection

    Set wsWBS = ThisWorkbook.Worksheets("WBS")
    Set tblWBS = wsWBS.ListObjects("tbl_WBS")

    If tblWBS.DataBodyRange Is Nothing Then
        CalcBridge_AddConsoleMessage consoleMessages, "STOP", _
            BiMsg( _
                "La table tbl_WBS est vide.", _
                "Table tbl_WBS is empty.")
        Write_CalcState_Snapshot_Console "ERROR", consoleMessages
        CalcBridge_ShowPlanningConsole consoleMessages
        GoTo SafeExit
    End If

    Ensure_Calc_Infrastructure consoleMessages
    Ensure_Analytics_Toggle
    AbortIfRequested "Run_Calc_Engine_CoreBridge.AfterEnsureCalc"

    Set wsCalc = ThisWorkbook.Worksheets("CALC")
    Set tblCalc = wsCalc.ListObjects("tbl_CALC")

    'Preserve existing calculated outputs during sync.
    'Full mode still clears/rebuilds inside Run_Calc_Core.
    'Partial mode needs previous calculated dates available as predecessor state.
    Sync_WBS_To_CALC True
    AbortIfRequested "Run_Calc_Engine_CoreBridge.AfterSync"

    If Not Sync_Constraints_To_CALC_ForWorkflow(consoleMessages) Then
        Write_CalcState_Snapshot_Console "ERROR", consoleMessages
        CalcBridge_ShowPlanningConsole consoleMessages
        GoTo SafeExit
    End If
    AbortIfRequested "Run_Calc_Engine_CoreBridge.AfterConstraintSync"

    Set wsCalc = ThisWorkbook.Worksheets("CALC")
    Set tblCalc = wsCalc.ListObjects("tbl_CALC")

    If Not ValidateCalcAfterSync(tblCalc) Then
        CalcBridge_AddConsoleMessage consoleMessages, "STOP", _
            BiMsg( _
                "Le sync WBS -> CALC a échoué ou a laissé tbl_CALC dans un état invalide." & vbCrLf & _
                "Le calcul est arrêté pour éviter un faux succès.", _
                "WBS -> CALC sync failed or left tbl_CALC in an invalid state." & vbCrLf & _
                "Calculation stopped to avoid a false success.")
        Write_CalcState_Snapshot_Console "ERROR", consoleMessages
        CalcBridge_ShowPlanningConsole consoleMessages
        GoTo SafeExit
    End If

    ApplyCalcDateFormats tblCalc

    RebuildLogicLinksTable
    AbortIfRequested "Run_Calc_Engine_CoreBridge.AfterRebuildLogicLinks"

    Set wsCalc = ThisWorkbook.Worksheets("CALC")
    Set tblCalc = wsCalc.ListObjects("tbl_CALC")
    Set mapCalc = Core_BuildColumnMap_FromListObject(tblCalc)

    FillCalcParentAndSummaryFromWBS tblCalc, tblWBS
    CalcBridge_ShowParentDateWarnings tblCalc, mapCalc, consoleMessages

    Set changedIds = Get_Changed_TaskIds(forceFullRecalc)

    If forceFullRecalcOverride Then
        forceFullRecalc = True
        Set changedIds = CreateObject("Scripting.Dictionary")
        Debug.Print "FULL RECALC FORCED BY USER"
    ElseIf forceFullRecalc Then
        Debug.Print "FULL RECALC FORCED"
    Else
        Debug.Print "Changed tasks count: " & changedIds.Count
    End If

    Set mapCalc = Core_BuildColumnMap_FromListObject(tblCalc)
    Set mapWBS = Core_BuildColumnMap_FromListObject(tblWBS)
    dataArr = tblCalc.DataBodyRange.value

    If Not forceFullRecalc Then
        If Not changedIds Is Nothing Then
            If changedIds.Count = 0 Then

                Debug.Print "NO CHANGE DETECTED - CORE SKIPPED"

                CalcBridge_AppendTaskTypeWarnings consoleMessages, tblWBS, mapWBS, tblCalc, mapCalc
                Write_CalcState_Snapshot_Console "OK", consoleMessages

                CalcBridge_AddInfoMessage consoleMessages, _
                    "Aucune modification détectée : calcul moteur non relancé.", _
                    "No change detected: core calculation was not rerun."

                CalcBridge_ShowPlanningConsole consoleMessages
                GoTo SafeExit

            End If
        End If
    End If

    If CalcBridge_PreCore_CheckLOEAsPredecessor(tblCalc, mapCalc, consoleMessages) Then
        Write_CalcState_Snapshot_Console "ERROR", consoleMessages
        CalcBridge_ShowPlanningConsole consoleMessages
        GoTo SafeExit
    End If

    Set linksBySuccId = BuildCoreLinksBySucc_FromLogicLinksTable_Expanded(tblCalc)

    Set succByPred = Build_Successor_Map(linksBySuccId)
    Set parentById = BuildParentByIdMap_FromCalc(dataArr, mapCalc, Core_BuildRowById(dataArr, mapCalc))
    Set impactedIds = Build_Impacted_TaskIds(changedIds, succByPred, parentById)

    Debug.Print "Impacted tasks count: " & impactedIds.Count

    usePartialWrite = False

    If Not forceFullRecalc Then
        If Not changedIds Is Nothing Then
            If changedIds.Count > 0 Then
                If Not impactedIds Is Nothing Then
                    If impactedIds.Count > 0 Then

                        totalRows = 0
                        impactedCount = impactedIds.Count
                        impactRatio = 1#

                        If Not tblCalc.DataBodyRange Is Nothing Then
                            totalRows = tblCalc.ListRows.Count
                        End If

                        If totalRows > 0 Then
                            impactRatio = CDbl(impactedCount) / CDbl(totalRows)
                        End If

                        Debug.Print "Partial impact ratio: " & Format$(impactRatio, "0.00%")

                        If impactRatio <= PARTIAL_RECALC_MAX_IMPACT_RATIO Then
                            usePartialWrite = True
                        Else
                            usePartialWrite = False
                            Debug.Print "PARTIAL DISABLED - IMPACT RATIO ABOVE " & Format$(PARTIAL_RECALC_MAX_IMPACT_RATIO, "0%")
                        End If

                    End If
                End If
            End If
        End If
    End If

    If usePartialWrite Then

        Debug.Print "PARTIAL CORE MODE ENABLED"
        Debug.Print "Partial impacted tasks count: " & impactedIds.Count

        Run_Calc_Core dataArr, mapCalc, linksBySuccId, impactedIds

        WriteCoreOutputsToCalc_Partial tblCalc, mapCalc, dataArr, impactedIds
        WriteCoreDrivingLogicToCalc_Partial tblCalc, mapCalc, dataArr, impactedIds

    Else

        Debug.Print "FULL CORE MODE"

        Run_Calc_Core dataArr, mapCalc, linksBySuccId

        WriteCoreOutputsToCalc tblCalc, mapCalc, dataArr
        WriteCoreDrivingLogicToCalc tblCalc, mapCalc, dataArr

    End If

    ApplyCalcDateFormats tblCalc
    AbortIfRequested "Run_Calc_Engine_CoreBridge.AfterCoreWrite"

    If CalcBridge_HasCoreErrors(tblCalc) Then
        CalcBridge_AppendCoreErrorMessages consoleMessages, tblCalc
        Write_CalcState_Snapshot_Console "ERROR", consoleMessages
        CalcBridge_ShowPlanningConsole consoleMessages
        GoTo SafeExit
    End If

    'Non-blocking warnings only after a clean core calculation.
    'Warnings are collected into the planning console and shown once at the end.
    runAnalytics = IsAnalyticsEnabled()


    CalcBridge_AppendTaskTypeWarnings consoleMessages, tblWBS, mapWBS, tblCalc, mapCalc

    'Analytics are controlled by the WBS user toggle and stored in VBA memory.
    'When enabled: analytics are recomputed globally using existing analytics engine.
    'When disabled: analytics outputs are cleared to avoid stale CP / float values.
    'No partial analytics calculation is implemented at this stage.

    If runAnalytics Then
        Run_Analytics_Only consoleMessages
        AbortIfRequested "Run_Calc_Engine_CoreBridge.AfterAnalytics"
    Else
        Clear_Analytics_Outputs
    End If

    If usePartialWrite Then
        'Partial WBS push for core outputs only (dates / driving logic).
        'Analytics are handled separately by toggle logic above.
        Push_Calculated_Back_To_WBS_Partial impactedIds
    Else
        Push_Calculated_Back_To_WBS
    End If

    'Variances are recomputed globally.
    'This is deliberate for consistency because they depend on baseline references
    'and summary roll-ups.
    Compute_And_Push_Variances consoleMessages

    Write_CalcState_Snapshot_Console "OK", consoleMessages

    CalcBridge_AddInfoMessage consoleMessages, _
        "Calcul terminé avec succès.", _
        "Calculation completed successfully."

    CalcBridge_ShowPlanningConsole consoleMessages

SafeExit:
    If IsMacroAbortRequested() Then
        ShowAbortMessageOnce
        EndMacroRun
        Exit Sub
    End If

    EndMacroRun
    Exit Sub

ErrHandler:
    If consoleMessages Is Nothing Then Set consoleMessages = New Collection

    CalcBridge_AddConsoleMessage consoleMessages, "STOP", _
        BiMsg( _
            "Erreur dans Run_Calc_Engine_CoreBridge" & vbCrLf & _
            "-> " & Err.Description, _
            "Error in Run_Calc_Engine_CoreBridge" & vbCrLf & _
            "-> " & Err.Description)

    On Error Resume Next
    Write_CalcState_Snapshot_Console "ERROR", consoleMessages
    CalcBridge_ShowPlanningConsole consoleMessages
    On Error GoTo 0

    Resume SafeExit

End Sub


Private Function ValidateCalcAfterSync(ByVal tblCalc As ListObject) As Boolean

    Dim mapCalc As Object
    Dim arr As Variant
    Dim hasAnyId As Boolean
    Dim r As Long

    On Error GoTo Fail

    If tblCalc Is Nothing Then GoTo Fail
    If tblCalc.DataBodyRange Is Nothing Then GoTo Fail

    Set mapCalc = Core_BuildColumnMap_FromListObject(tblCalc)

    If Not mapCalc.Exists("ID") Then GoTo Fail
    If Not mapCalc.Exists("Calculated Start") Then GoTo Fail
    If Not mapCalc.Exists("Calculated Finish") Then GoTo Fail

    arr = tblCalc.DataBodyRange.value

    For r = 1 To UBound(arr, 1)
        If Trim$(CStr(arr(r, mapCalc("ID")))) <> "" Then
            hasAnyId = True
            Exit For
        End If
    Next r

    If Not hasAnyId Then GoTo Fail

    ValidateCalcAfterSync = True
    Exit Function

Fail:
    ValidateCalcAfterSync = False

End Function

Private Sub ApplyCalcDateFormats(ByVal tblCalc As ListObject)

    On Error Resume Next

    tblCalc.ListColumns("Baseline Start").DataBodyRange.NumberFormat = "dd/mm/yyyy"
    tblCalc.ListColumns("Baseline Finish").DataBodyRange.NumberFormat = "dd/mm/yyyy"
    tblCalc.ListColumns("Actual Start").DataBodyRange.NumberFormat = "dd/mm/yyyy"
    tblCalc.ListColumns("Actual Finish").DataBodyRange.NumberFormat = "dd/mm/yyyy"
    tblCalc.ListColumns("Forecast Start").DataBodyRange.NumberFormat = "dd/mm/yyyy"
    tblCalc.ListColumns("Forecast Finish").DataBodyRange.NumberFormat = "dd/mm/yyyy"
    tblCalc.ListColumns("Deadline").DataBodyRange.NumberFormat = "dd/mm/yyyy"
    tblCalc.ListColumns("Calculated Start").DataBodyRange.NumberFormat = "dd/mm/yyyy"
    tblCalc.ListColumns("Calculated Finish").DataBodyRange.NumberFormat = "dd/mm/yyyy"
    tblCalc.ListColumns("Deadline Float").DataBodyRange.NumberFormat = "0"

    On Error GoTo 0

End Sub



Public Function BuildPredsBySucc_FromExpandedLinks( _
    ByVal rowById As Object, _
    ByVal linksBySuccId As Object) As Object

    Dim predsBySucc As Object
    Dim anyId As Variant
    Dim succId As Variant
    Dim oneLink As Variant
    Dim predId As String

    Set predsBySucc = CreateObject("Scripting.Dictionary")

    For Each anyId In rowById.Keys
        Set predsBySucc(CStr(anyId)) = New Collection
    Next anyId

    If linksBySuccId Is Nothing Then
        Set BuildPredsBySucc_FromExpandedLinks = predsBySucc
        Exit Function
    End If

    For Each succId In linksBySuccId.Keys
        If Not predsBySucc.Exists(CStr(succId)) Then
            Set predsBySucc(CStr(succId)) = New Collection
        End If

        For Each oneLink In linksBySuccId(CStr(succId))
            predId = Core_GetLinkPredId(oneLink)
            If predId <> "" Then
                predsBySucc(CStr(succId)).Add CStr(predId)
            End If
        Next oneLink
    Next succId

    Set BuildPredsBySucc_FromExpandedLinks = predsBySucc

End Function

Private Sub WriteCoreDrivingLogicToCalc( _
    ByVal tblCalc As ListObject, _
    ByVal mapCalc As Object, _
    ByRef dataArr As Variant)

    Dim rowById As Object
    Dim parentIds As Object
    Dim rowCount As Long
    Dim outDriving() As Variant
    Dim r As Long
    Dim idVal As String
    Dim baselineStart As Variant
    Dim actualStart As Variant
    Dim actualFinish As Variant
    Dim forecastStart As Variant
    Dim forecastFinish As Variant
    Dim calcStart As Variant

    Set rowById = Core_BuildRowById(dataArr, mapCalc)
    Set parentIds = Core_BuildParentIds(dataArr, mapCalc, rowById)

    rowCount = UBound(dataArr, 1)
    ReDim outDriving(1 To rowCount, 1 To 1)

    For r = 1 To rowCount

        idVal = Trim$(CStr(dataArr(r, mapCalc("ID"))))
        outDriving(r, 1) = ""

        If idVal <> "" Then

            If parentIds.Exists(idVal) Then

                outDriving(r, 1) = "SUMMARY"

            ElseIf CalcBridge_IsLevelOfEffortRow(dataArr, mapCalc, r) Then

                outDriving(r, 1) = "LOE"

            Else

                baselineStart = dataArr(r, mapCalc("Baseline Start"))
                actualStart = dataArr(r, mapCalc("Actual Start"))
                actualFinish = dataArr(r, mapCalc("Actual Finish"))
                forecastStart = dataArr(r, mapCalc("Forecast Start"))
                forecastFinish = dataArr(r, mapCalc("Forecast Finish"))
                calcStart = dataArr(r, mapCalc("Calculated Start"))

                If CalcBridge_IsDrivenByActiveConstraint(dataArr, mapCalc, r) Then
                    outDriving(r, 1) = "CONSTRAINT"
                ElseIf HasValue(actualStart) Or HasValue(actualFinish) Then
                    outDriving(r, 1) = "ACTUAL"
                ElseIf HasValue(forecastStart) Or HasValue(forecastFinish) Then
                    outDriving(r, 1) = "FORECAST"
                ElseIf HasValue(calcStart) And HasValue(baselineStart) Then
                    If CDbl(calcStart) > CDbl(baselineStart) Then
                        outDriving(r, 1) = "DEPENDENCY"
                    Else
                        outDriving(r, 1) = "BASELINE"
                    End If
                Else
                    outDriving(r, 1) = "BASELINE"
                End If

            End If

        End If

    Next r

    tblCalc.ListColumns("Driving Logic").DataBodyRange.value = outDriving

End Sub

Private Function CalcBridge_IsDrivenByActiveConstraint( _
    ByRef dataArr As Variant, _
    ByVal mapCalc As Object, _
    ByVal rowIdx As Long) As Boolean

    Dim activeVal As String
    Dim startType As String
    Dim finishType As String
    Dim startDate As Variant
    Dim finishDate As Variant
    Dim calcStart As Variant
    Dim calcFinish As Variant

    CalcBridge_IsDrivenByActiveConstraint = False

    If mapCalc Is Nothing Then Exit Function
    If Not mapCalc.Exists("Constraint Active") Then Exit Function
    If Not mapCalc.Exists("Start Constraint Type") Then Exit Function
    If Not mapCalc.Exists("Start Constraint Date") Then Exit Function
    If Not mapCalc.Exists("Finish Constraint Type") Then Exit Function
    If Not mapCalc.Exists("Finish Constraint Date") Then Exit Function
    If Not mapCalc.Exists("Calculated Start") Then Exit Function
    If Not mapCalc.Exists("Calculated Finish") Then Exit Function

    activeVal = UCase$(Trim$(CStr(dataArr(rowIdx, mapCalc("Constraint Active")))))
    If activeVal <> "YES" Then Exit Function

    startType = UCase$(Trim$(CStr(dataArr(rowIdx, mapCalc("Start Constraint Type")))))
    finishType = UCase$(Trim$(CStr(dataArr(rowIdx, mapCalc("Finish Constraint Type")))))
    startDate = dataArr(rowIdx, mapCalc("Start Constraint Date"))
    finishDate = dataArr(rowIdx, mapCalc("Finish Constraint Date"))
    calcStart = dataArr(rowIdx, mapCalc("Calculated Start"))
    calcFinish = dataArr(rowIdx, mapCalc("Calculated Finish"))

    If startType = "MUST START ON" Then
        CalcBridge_IsDrivenByActiveConstraint = True
        Exit Function
    End If

    If finishType = "MUST FINISH ON" Then
        CalcBridge_IsDrivenByActiveConstraint = True
        Exit Function
    End If

    Select Case startType
        Case "START NO EARLIER THAN", "START NO LATER THAN"
            If CalcBridge_DatesMatch(calcStart, startDate) Then
                CalcBridge_IsDrivenByActiveConstraint = True
                Exit Function
            End If
    End Select

    Select Case finishType
        Case "FINISH NO EARLIER THAN", "FINISH NO LATER THAN"
            If CalcBridge_DatesMatch(calcFinish, finishDate) Then
                CalcBridge_IsDrivenByActiveConstraint = True
                Exit Function
            End If
    End Select

End Function

Private Function CalcBridge_DatesMatch(ByVal leftVal As Variant, ByVal rightVal As Variant) As Boolean

    CalcBridge_DatesMatch = False

    If Not HasValue(leftVal) Then Exit Function
    If Not HasValue(rightVal) Then Exit Function
    If Not IsDate(leftVal) Then Exit Function
    If Not IsDate(rightVal) Then Exit Function

    CalcBridge_DatesMatch = (CLng(CDbl(CDate(leftVal))) = CLng(CDbl(CDate(rightVal))))

End Function

Private Function BuildIdToWbsFromWBS(ByVal tblWBS As ListObject) As Object

    Dim mapWBS As Object
    Dim arrWBS As Variant
    Dim idToWbs As Object
    Dim r As Long
    Dim idVal As String
    Dim wbsVal As String
    Dim taskNameVal As String

    Set mapWBS = Core_BuildColumnMap_FromListObject(tblWBS)
    Set idToWbs = CreateObject("Scripting.Dictionary")

    If tblWBS.DataBodyRange Is Nothing Then
        Set BuildIdToWbsFromWBS = idToWbs
        Exit Function
    End If

    arrWBS = tblWBS.DataBodyRange.value

    For r = 1 To UBound(arrWBS, 1)
        idVal = Trim$(CStr(arrWBS(r, mapWBS("ID"))))
        wbsVal = NormalizeWBS(arrWBS(r, mapWBS("WBS")))

        If idVal <> "" Then
            idToWbs(idVal) = wbsVal
        End If
    Next r

    Set BuildIdToWbsFromWBS = idToWbs

End Function

Private Function CalcBridge_GetCoreErrorSummary( _
    ByVal tblCalc As ListObject, _
    Optional ByVal maxRows As Long = 12) As String

    Dim mapCalc As Object
    Dim arr As Variant
    Dim r As Long
    Dim countShown As Long
    Dim msg As String

    Dim idVal As String
    Dim wbsVal As String
    Dim taskNameVal As String
    Dim errMsg As String

    On Error GoTo FailSafe

    CalcBridge_GetCoreErrorSummary = ""

    If tblCalc Is Nothing Then Exit Function
    If tblCalc.DataBodyRange Is Nothing Then Exit Function

    Set mapCalc = Core_BuildColumnMap_FromListObject(tblCalc)

    If Not mapCalc.Exists("Error flag") Then Exit Function
    If Not mapCalc.Exists("ErrorMsg") Then Exit Function
    If Not mapCalc.Exists("ID") Then Exit Function

    arr = tblCalc.DataBodyRange.value

    For r = 1 To UBound(arr, 1)

        If UCase$(Trim$(CStr(arr(r, mapCalc("Error flag"))))) = "ERROR" Then

            idVal = Trim$(CStr(arr(r, mapCalc("ID"))))
            errMsg = Trim$(CStr(arr(r, mapCalc("ErrorMsg"))))

            If mapCalc.Exists("WBS") Then
                wbsVal = Trim$(CStr(arr(r, mapCalc("WBS"))))
            Else
                wbsVal = ""
            End If

            countShown = countShown + 1

            If msg <> "" Then msg = msg & vbCrLf

            msg = msg & "- ID " & idVal
            If wbsVal <> "" Then msg = msg & " / WBS " & wbsVal
            If errMsg <> "" Then msg = msg & " : " & errMsg

            If countShown >= maxRows Then Exit For

        End If

    Next r

    If msg = "" Then
        msg = "Error flag detected, but no detailed message was available."
    End If

    CalcBridge_GetCoreErrorSummary = msg
    Exit Function

FailSafe:
    CalcBridge_GetCoreErrorSummary = "Unable to build core error summary."

End Function

Private Function CalcBridge_HasCoreErrors(ByVal tblCalc As ListObject) As Boolean

    Dim mapCalc As Object
    Dim arr As Variant
    Dim r As Long

    On Error GoTo FailSafe

    CalcBridge_HasCoreErrors = False

    If tblCalc Is Nothing Then Exit Function
    If tblCalc.DataBodyRange Is Nothing Then Exit Function

    Set mapCalc = Core_BuildColumnMap_FromListObject(tblCalc)

    If Not mapCalc.Exists("Error flag") Then Exit Function

    arr = tblCalc.DataBodyRange.value

    For r = 1 To UBound(arr, 1)
        If UCase$(Trim$(CStr(arr(r, mapCalc("Error flag"))))) = "ERROR" Then
            CalcBridge_HasCoreErrors = True
            Exit Function
        End If
    Next r

    Exit Function

FailSafe:
    CalcBridge_HasCoreErrors = True

End Function

Private Function CalcBridge_IsInheritedCoreError(ByVal errMsg As String) As Boolean

    Dim txt As String

    txt = Trim$(CStr(errMsg))

    CalcBridge_IsInheritedCoreError = _
        (InStr(1, txt, "Blocked by predecessor error", vbTextCompare) > 0) Or _
        (InStr(1, txt, "Blocked by predecessor chain", vbTextCompare) > 0)

End Function
Private Function CalcBridge_IsConstraintCoreError(ByVal errMsg As String) As Boolean

    Dim txt As String

    txt = Trim$(CStr(errMsg))

    CalcBridge_IsConstraintCoreError = _
        (InStr(1, txt, "avant contrainte debut", vbTextCompare) > 0) Or _
        (InStr(1, txt, "before start constraint", vbTextCompare) > 0) Or _
        (InStr(1, txt, "apres contrainte debut max", vbTextCompare) > 0) Or _
        (InStr(1, txt, "after latest start constraint", vbTextCompare) > 0) Or _
        (InStr(1, txt, "different de contrainte Must Start On", vbTextCompare) > 0) Or _
        (InStr(1, txt, "differs from Must Start On constraint", vbTextCompare) > 0) Or _
        (InStr(1, txt, "avant contrainte fin", vbTextCompare) > 0) Or _
        (InStr(1, txt, "before finish constraint", vbTextCompare) > 0) Or _
        (InStr(1, txt, "apres contrainte fin max", vbTextCompare) > 0) Or _
        (InStr(1, txt, "after latest finish constraint", vbTextCompare) > 0) Or _
        (InStr(1, txt, "contrainte Must Finish On", vbTextCompare) > 0) Or _
        (InStr(1, txt, "Must Finish On constraint", vbTextCompare) > 0) Or _
        (InStr(1, txt, "contraintes Must Start On / Must Finish On", vbTextCompare) > 0) Or _
        (InStr(1, txt, "Must Start On / Must Finish On constraints", vbTextCompare) > 0) Or _
        (InStr(1, txt, "Type de contrainte debut non reconnu", vbTextCompare) > 0) Or _
        (InStr(1, txt, "Unknown start constraint type", vbTextCompare) > 0) Or _
        (InStr(1, txt, "Type de contrainte fin non reconnu", vbTextCompare) > 0) Or _
        (InStr(1, txt, "Unknown finish constraint type", vbTextCompare) > 0)

End Function

Private Sub CalcBridge_AddConstraintRootMessages( _
    ByVal consoleMessages As Collection, _
    ByVal constraintMessagesById As Object)

    Dim idVal As Variant

    If consoleMessages Is Nothing Then Exit Sub
    If constraintMessagesById Is Nothing Then Exit Sub

    For Each idVal In constraintMessagesById.Keys
        CalcBridge_AddConsoleMessage consoleMessages, "STOP", _
            CStr(constraintMessagesById(CStr(idVal)))
    Next idVal

End Sub

Private Sub CalcBridge_AppendCoreErrorMessages( _
    ByVal consoleMessages As Collection, _
    ByVal tblCalc As ListObject)

    Dim mapCalc As Object
    Dim arr As Variant
    Dim r As Long

    Dim idToWbs As Object

    Dim errMissingPred As Object
    Dim errCycle As Object
    Dim errUnsupportedLinkType As Object
    Dim errActualStartConflict As Object
    Dim errActualFinishConflict As Object
    Dim errForecastConflict As Object
    Dim errForecastFinishConflict As Object
    Dim errMissingDuration As Object
    Dim errStartNotComputable As Object
    Dim errFinishBeforeStart As Object
    Dim errLOEAsPredecessor As Object
    Dim errLOEMissingSS As Object
    Dim errLOEMissingFF As Object
    Dim errLOEInvalidLink As Object
    Dim errOtherRoot As Object
    Dim errConstraintRootMessages As Object
    Dim cycleDetailMessage As String

    Dim idVal As String
    Dim wbsVal As String
    Dim taskNameVal As String
    Dim errMsg As String
    Dim hasSpecificRootError As Boolean

    On Error GoTo FailSafe

    If consoleMessages Is Nothing Then Exit Sub

    Set errMissingPred = CreateObject("Scripting.Dictionary")
    Set errCycle = CreateObject("Scripting.Dictionary")
    Set errUnsupportedLinkType = CreateObject("Scripting.Dictionary")
    Set errActualStartConflict = CreateObject("Scripting.Dictionary")
    Set errActualFinishConflict = CreateObject("Scripting.Dictionary")
    Set errForecastConflict = CreateObject("Scripting.Dictionary")
    Set errForecastFinishConflict = CreateObject("Scripting.Dictionary")
    Set errMissingDuration = CreateObject("Scripting.Dictionary")
    Set errStartNotComputable = CreateObject("Scripting.Dictionary")
    Set errFinishBeforeStart = CreateObject("Scripting.Dictionary")
    Set errLOEAsPredecessor = CreateObject("Scripting.Dictionary")
    Set errLOEMissingSS = CreateObject("Scripting.Dictionary")
    Set errLOEMissingFF = CreateObject("Scripting.Dictionary")
    Set errLOEInvalidLink = CreateObject("Scripting.Dictionary")
    Set errOtherRoot = CreateObject("Scripting.Dictionary")
    Set errConstraintRootMessages = CreateObject("Scripting.Dictionary")
    Set idToWbs = CreateObject("Scripting.Dictionary")

    If tblCalc Is Nothing Then GoTo FailSafe
    If tblCalc.DataBodyRange Is Nothing Then GoTo FailSafe

    Set mapCalc = Core_BuildColumnMap_FromListObject(tblCalc)

    If Not mapCalc.Exists("ID") Then GoTo FailSafe
    If Not mapCalc.Exists("Error flag") Then GoTo FailSafe
    If Not mapCalc.Exists("ErrorMsg") Then GoTo FailSafe

    arr = tblCalc.DataBodyRange.value

    For r = 1 To UBound(arr, 1)

        idVal = Trim$(CStr(arr(r, mapCalc("ID"))))

        If idVal <> "" Then

            If mapCalc.Exists("WBS") Then
                wbsVal = Trim$(CStr(arr(r, mapCalc("WBS"))))
            Else
                wbsVal = ""
            End If

            idToWbs(idVal) = wbsVal

            If UCase$(Trim$(CStr(arr(r, mapCalc("Error flag"))))) = "ERROR" Then

                errMsg = Trim$(CStr(arr(r, mapCalc("ErrorMsg"))))

                'Inherited errors remain visible in tbl_CALC,
                'but are not highlighted in the main popup.
                If CalcBridge_IsInheritedCoreError(errMsg) Then
                    GoTo NextRow
                End If

                Select Case True

                    Case InStr(1, errMsg, "LOE cannot be used as predecessor", vbTextCompare) > 0 Or _
                         InStr(1, errMsg, "Blocked by invalid LOE predecessor", vbTextCompare) > 0
                        errLOEAsPredecessor(idVal) = True

                    Case InStr(1, errMsg, "LOE must have at least one SS predecessor", vbTextCompare) > 0 Or _
                         InStr(1, errMsg, "LOE SS predecessor missing", vbTextCompare) > 0 Or _
                         InStr(1, errMsg, "LOE SS predecessor not found", vbTextCompare) > 0 Or _
                         InStr(1, errMsg, "LOE SS predecessor start not available", vbTextCompare) > 0
                        errLOEMissingSS(idVal) = True

                    Case InStr(1, errMsg, "LOE must have at least one FF predecessor", vbTextCompare) > 0 Or _
                         InStr(1, errMsg, "LOE FF predecessor missing", vbTextCompare) > 0 Or _
                         InStr(1, errMsg, "LOE FF predecessor not found", vbTextCompare) > 0 Or _
                         InStr(1, errMsg, "LOE FF predecessor finish not available", vbTextCompare) > 0
                        errLOEMissingFF(idVal) = True

                    Case InStr(1, errMsg, "LOE only supports SS and FF predecessors", vbTextCompare) > 0
                        errLOEInvalidLink(idVal) = True

                    Case InStr(1, errMsg, "Missing predecessor", vbTextCompare) > 0
                        errMissingPred(idVal) = True

                    Case InStr(1, errMsg, "Cycle detected", vbTextCompare) > 0
                        errCycle(idVal) = True
                        If cycleDetailMessage = "" Then cycleDetailMessage = errMsg

                    Case InStr(1, errMsg, "Unsupported link type", vbTextCompare) > 0
                        errUnsupportedLinkType(idVal) = True

                    Case InStr(1, errMsg, "Actual Start violates dependencies", vbTextCompare) > 0
                        errActualStartConflict(idVal) = True

                    Case InStr(1, errMsg, "Actual Finish violates finish constraints", vbTextCompare) > 0
                        errActualFinishConflict(idVal) = True

                    Case InStr(1, errMsg, "Forecast Start violates dependencies", vbTextCompare) > 0
                        errForecastConflict(idVal) = True

                    Case InStr(1, errMsg, "Forecast Finish violates finish constraints", vbTextCompare) > 0
                        errForecastFinishConflict(idVal) = True

                    Case InStr(1, errMsg, "Baseline Duration missing", vbTextCompare) > 0
                        errMissingDuration(idVal) = True

                    Case InStr(1, errMsg, "Start date not computable", vbTextCompare) > 0
                        errStartNotComputable(idVal) = True

                    Case InStr(1, errMsg, "Finish before start", vbTextCompare) > 0
                        errFinishBeforeStart(idVal) = True

                    Case CalcBridge_IsConstraintCoreError(errMsg)
                        errConstraintRootMessages(idVal) = errMsg

                    Case Else
                        errOtherRoot(idVal) = True

                End Select

            End If
        End If

NextRow:
    Next r

    hasSpecificRootError = _
        (errMissingPred.Count > 0) Or _
        (errUnsupportedLinkType.Count > 0) Or _
        (errCycle.Count > 0) Or _
        (errActualStartConflict.Count > 0) Or _
        (errActualFinishConflict.Count > 0) Or _
        (errForecastConflict.Count > 0) Or _
        (errForecastFinishConflict.Count > 0) Or _
        (errMissingDuration.Count > 0) Or _
        (errStartNotComputable.Count > 0) Or _
        (errFinishBeforeStart.Count > 0) Or _
        (errConstraintRootMessages.Count > 0) Or _
        (errLOEAsPredecessor.Count > 0) Or _
        (errLOEMissingSS.Count > 0) Or _
        (errLOEMissingFF.Count > 0) Or _
        (errLOEInvalidLink.Count > 0)

    If errLOEAsPredecessor.Count > 0 Then
        CalcBridge_AddGroupedStopToCollection consoleMessages, errLOEAsPredecessor, idToWbs, _
            "LOE utilisée comme prédécesseur", _
            "supprimer la LOE de la logique amont ; une LOE est pilotée par le réseau mais ne doit pas piloter d'autres tâches", _
            "LOE used as predecessor", _
            "remove the LOE from upstream logic; a LOE is driven by the network but must not drive other tasks"
    End If

    If errLOEMissingSS.Count > 0 Then
        CalcBridge_AddGroupedStopToCollection consoleMessages, errLOEMissingSS, idToWbs, _
            "LOE sans lien SS exploitable", _
            "ajouter au moins un prédécesseur SS valide pour définir le début de la LOE", _
            "LOE without usable SS link", _
            "add at least one valid SS predecessor to define the LOE start"
    End If

    If errLOEMissingFF.Count > 0 Then
        CalcBridge_AddGroupedStopToCollection consoleMessages, errLOEMissingFF, idToWbs, _
            "LOE sans lien FF exploitable", _
            "ajouter au moins un prédécesseur FF valide pour définir la fin de la LOE", _
            "LOE without usable FF link", _
            "add at least one valid FF predecessor to define the LOE finish"
    End If

    If errLOEInvalidLink.Count > 0 Then
        CalcBridge_AddGroupedStopToCollection consoleMessages, errLOEInvalidLink, idToWbs, _
            "Lien invalide sur LOE", _
            "une LOE accepte uniquement des liens SS et FF", _
            "Invalid link on LOE", _
            "a LOE only supports SS and FF links"
    End If

    If errMissingPred.Count > 0 Then
        CalcBridge_AddGroupedStopToCollection consoleMessages, errMissingPred, idToWbs, _
            "Prédécesseur introuvable", _
            "vérifier la colonne Predecessors WBS", _
            "Missing predecessor", _
            "check the Predecessors WBS column"
    End If

    If errUnsupportedLinkType.Count > 0 Then
        CalcBridge_AddGroupedStopToCollection consoleMessages, errUnsupportedLinkType, idToWbs, _
            "Type de lien non supporté par le moteur", _
            "corriger le type de lien dans Predecessors WBS ou tbl_LOGIC_LINKS", _
            "Link type not supported by the engine", _
            "fix the link type in Predecessors WBS or tbl_LOGIC_LINKS"
    End If

    If errCycle.Count > 0 Then
        If CalcBridge_CycleDetailMessageFromCoreError(cycleDetailMessage) <> "" Then
            CalcBridge_AddConsoleMessage consoleMessages, "STOP", _
                CalcBridge_CycleDetailMessageFromCoreError(cycleDetailMessage)
        Else
            CalcBridge_AddGroupedStopToCollection consoleMessages, errCycle, idToWbs, _
                "Boucle de dépendance détectée", _
                "corriger la colonne Predecessors WBS", _
                "Dependency cycle detected", _
                "fix the Predecessors WBS column"
        End If
    End If

    If errActualStartConflict.Count > 0 Then
        CalcBridge_AddUpstreamStopToCollection consoleMessages, errActualStartConflict, idToWbs, _
            "Actual Start incohérent avec les dépendances amont", _
            "corriger la logique, le lag, ou la date actual", _
            "Actual Start violates upstream dependencies", _
            "fix logic, lag, or actual date"
    End If

    If errActualFinishConflict.Count > 0 Then
        CalcBridge_AddUpstreamStopToCollection consoleMessages, errActualFinishConflict, idToWbs, _
            "Actual Finish incohérent avec les contraintes de fin amont", _
            "corriger la logique, le lag, ou la date actual", _
            "Actual Finish violates upstream finish constraints", _
            "fix logic, lag, or actual date"
    End If

    If errForecastConflict.Count > 0 Then
        CalcBridge_AddGroupedStopToCollection consoleMessages, errForecastConflict, idToWbs, _
            "Forecast incohérent avec les dépendances", _
            "ajuster Forecast Start", _
            "Forecast violates dependencies", _
            "adjust Forecast Start"
    End If

    If errForecastFinishConflict.Count > 0 Then
        CalcBridge_AddGroupedStopToCollection consoleMessages, errForecastFinishConflict, idToWbs, _
            "Forecast Finish incohérent avec les contraintes de fin amont", _
            "ajuster Forecast Finish", _
            "Forecast Finish violates upstream finish constraints", _
            "adjust Forecast Finish"
    End If

    If errConstraintRootMessages.Count > 0 Then
        CalcBridge_AddConstraintRootMessages consoleMessages, errConstraintRootMessages
    End If

    If errMissingDuration.Count > 0 Then
        CalcBridge_AddGroupedStopToCollection consoleMessages, errMissingDuration, idToWbs, _
            "Baseline Duration manquante", _
            "compléter la durée baseline", _
            "Missing Baseline Duration", _
            "please fill in Baseline Duration"
    End If

    If errStartNotComputable.Count > 0 Then
        CalcBridge_AddGroupedStopToCollection consoleMessages, errStartNotComputable, idToWbs, _
            "Date de début non déterminable", _
            "vérifier les dépendances ou la baseline", _
            "Start date not computable", _
            "check dependencies or baseline"
    End If

    If errFinishBeforeStart.Count > 0 Then
        CalcBridge_AddGroupedStopToCollection consoleMessages, errFinishBeforeStart, idToWbs, _
            "Fin avant début détectée", _
            "vérifier les dates ou la durée", _
            "Finish before start detected", _
            "check dates or duration"
    End If

    If errOtherRoot.Count > 0 And Not hasSpecificRootError Then
        CalcBridge_AddGroupedStopToCollection consoleMessages, errOtherRoot, idToWbs, _
            "Erreur bloquante détectée par le moteur", _
            "vérifier ErrorMsg dans tbl_CALC pour le détail technique", _
            "Blocking error detected by the engine", _
            "check ErrorMsg in tbl_CALC for technical details"
    End If

    Exit Sub

FailSafe:
    CalcBridge_AddConsoleMessage consoleMessages, "STOP", _
        BiMsg( _
            "Calcul arrêté : le moteur a détecté des erreurs bloquantes." & vbCrLf & _
            "Impossible de reconstruire le message détaillé." & vbCrLf & _
            "-> vérifier les colonnes Error flag et ErrorMsg dans tbl_CALC." & vbCrLf & _
            "-> aucune donnée calculée n'a été repoussée vers WBS.", _
            "Calculation stopped: the engine detected blocking errors." & vbCrLf & _
            "Unable to rebuild the detailed message." & vbCrLf & _
            "-> check Error flag and ErrorMsg columns in tbl_CALC." & vbCrLf & _
            "-> no calculated data was pushed back to WBS.")

End Sub

Private Sub CalcBridge_ShowGroupedErrorMessage( _
    ByVal idsDict As Object, _
    ByVal idToWbs As Object, _
    ByVal frProblem As String, _
    ByVal frAction As String, _
    ByVal enProblem As String, _
    ByVal enAction As String)

    Dim consoleMessages As Collection

    If idsDict Is Nothing Then Exit Sub
    If idsDict.Count = 0 Then Exit Sub

    Set consoleMessages = New Collection

    CalcBridge_AddGroupedStopToCollection consoleMessages, idsDict, idToWbs, _
        frProblem, frAction, enProblem, enAction

    CalcBridge_ShowPlanningConsole consoleMessages

End Sub

Private Function CalcBridge_CycleDetailMessageFromCoreError(ByVal errMsg As String) As String

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
Private Function CalcBridge_BuildGroupedMessage( _
    ByVal idsDict As Object, _
    ByVal idToWbs As Object, _
    ByVal frProblem As String, _
    ByVal frAction As String, _
    ByVal enProblem As String, _
    ByVal enAction As String) As String

    Dim idsLine As String
    Dim wbsLine As String

    idsLine = CalcBridge_BuildInlineList(idsDict, 20)
    wbsLine = CalcBridge_BuildInlineWBSList(idsDict, idToWbs, 20)

    CalcBridge_BuildGroupedMessage = _
        "FR:" & vbCrLf & _
        frProblem & vbCrLf & _
        "-> " & frAction & vbCrLf & vbCrLf & _
        "IDs : " & idsLine & vbCrLf & _
        "WBS : " & wbsLine & vbCrLf & vbCrLf & _
        "EN:" & vbCrLf & _
        enProblem & vbCrLf & _
        "-> " & enAction & vbCrLf & vbCrLf & _
        "IDs: " & idsLine & vbCrLf & _
        "WBS: " & wbsLine

End Function

Private Function CalcBridge_BuildInlineList( _
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

Private Function CalcBridge_BuildInlineWBSList( _
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

Private Function CalcBridge_BuildUpstreamViolationItems( _
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

Private Function CalcBridge_PreCore_CheckMissingPredecessors( _
    ByVal tblCalc As ListObject, _
    ByVal mapCalc As Object, _
    ByVal rowById As Object, _
    Optional ByVal consoleMessages As Collection) As Boolean

    Dim wsCalc As Worksheet
    Dim tblLinks As ListObject
    Dim mapLinks As Object
    Dim arrLinks As Variant

    Dim idToWbs As Object
    Dim errMissingPred As Object

    Dim i As Long
    Dim r As Long
    Dim succId As String
    Dim predId As String

    On Error GoTo FailSafe

    CalcBridge_PreCore_CheckMissingPredecessors = False

    Set errMissingPred = CreateObject("Scripting.Dictionary")
    Set idToWbs = CalcBridge_PreCore_BuildIdToWbsFromCalc(tblCalc, mapCalc)

    Set wsCalc = ThisWorkbook.Worksheets("CALC")
    Set tblLinks = wsCalc.ListObjects("tbl_LOGIC_LINKS")

    If tblLinks Is Nothing Then Exit Function
    If tblLinks.DataBodyRange Is Nothing Then Exit Function

    Set mapLinks = CreateObject("Scripting.Dictionary")
    For i = 1 To tblLinks.ListColumns.Count
        mapLinks(tblLinks.ListColumns(i).Name) = i
    Next i

    If Not mapLinks.Exists("Succ ID") Then Exit Function
    If Not mapLinks.Exists("Pred ID") Then Exit Function

    arrLinks = tblLinks.DataBodyRange.value

    For r = 1 To UBound(arrLinks, 1)

        succId = Trim$(CStr(arrLinks(r, mapLinks("Succ ID"))))
        predId = Trim$(CStr(arrLinks(r, mapLinks("Pred ID"))))

        If succId <> "" Then
            If predId = "" Then
                errMissingPred(succId) = True
            ElseIf Not rowById.Exists(predId) Then
                errMissingPred(succId) = True
            End If
        End If

    Next r

    If errMissingPred.Count > 0 Then

        CalcBridge_PreCore_MarkErrorIdsInCalc tblCalc, mapCalc, errMissingPred, "Missing predecessor"

        If consoleMessages Is Nothing Then
            CalcBridge_ShowGroupedErrorMessage errMissingPred, idToWbs, _
                "Prédécesseur introuvable dans tbl_LOGIC_LINKS", _
                "vérifier la table des liens logiques", _
                "Missing predecessor in tbl_LOGIC_LINKS", _
                "check the logical links table"
        Else
            CalcBridge_AddGroupedStopToCollection consoleMessages, errMissingPred, idToWbs, _
                "Prédécesseur introuvable dans tbl_LOGIC_LINKS", _
                "vérifier la table des liens logiques", _
                "Missing predecessor in tbl_LOGIC_LINKS", _
                "check the logical links table"
        End If

        CalcBridge_PreCore_CheckMissingPredecessors = True
    End If

    Exit Function

FailSafe:
    CalcBridge_AddOrShowConsoleMessage consoleMessages, "STOP", _
        "Erreur pendant le contrôle des prédécesseurs." & vbCrLf & _
        "-> calcul arrêté avant écriture WBS.", _
        "Error while checking predecessors." & vbCrLf & _
        "-> calculation stopped before WBS write."

    CalcBridge_PreCore_CheckMissingPredecessors = True

End Function

Private Function CalcBridge_PreCore_CheckCycles( _
    ByVal tblCalc As ListObject, _
    ByVal mapCalc As Object, _
    ByVal validIds As Object, _
    ByVal linksBySuccId As Object, _
    Optional ByVal consoleMessages As Collection) As Boolean

    Dim cycleIds As Object
    Dim idToWbs As Object

    On Error GoTo FailSafe

    CalcBridge_PreCore_CheckCycles = False

    Set cycleIds = CalcBridge_PreCore_DetectCycleIds(validIds, linksBySuccId)

    If cycleIds.Count > 0 Then

        Set idToWbs = CalcBridge_PreCore_BuildIdToWbsFromCalc(tblCalc, mapCalc)

        CalcBridge_PreCore_MarkErrorIdsInCalc tblCalc, mapCalc, cycleIds, "Cycle detected"

        If consoleMessages Is Nothing Then
            CalcBridge_ShowGroupedErrorMessage cycleIds, idToWbs, _
                "Boucle de dépendance détectée", _
                "corriger la colonne Predecessors WBS", _
                "Dependency cycle detected", _
                "fix the Predecessors WBS column"
        Else
            CalcBridge_AddGroupedStopToCollection consoleMessages, cycleIds, idToWbs, _
                "Boucle de dépendance détectée", _
                "corriger la colonne Predecessors WBS", _
                "Dependency cycle detected", _
                "fix the Predecessors WBS column"
        End If

        CalcBridge_PreCore_CheckCycles = True
    End If

    Exit Function

FailSafe:
    CalcBridge_AddOrShowConsoleMessage consoleMessages, "STOP", _
        "Erreur pendant le contrôle des cycles." & vbCrLf & _
        "-> calcul arrêté avant écriture WBS.", _
        "Error while checking dependency cycles." & vbCrLf & _
        "-> calculation stopped before WBS write."

    CalcBridge_PreCore_CheckCycles = True

End Function


Private Function CalcBridge_PreCore_DetectCycleIds( _
    ByVal validIds As Object, _
    ByVal linksBySuccId As Object) As Object

    Dim childrenByPred As Object
    Dim state As Object
    Dim stackIndex As Object
    Dim stack As Collection
    Dim cycleIds As Object

    Dim idKey As Variant
    Dim succId As Variant
    Dim oneLink As Variant
    Dim predId As String
    Dim colChildren As Collection

    Set childrenByPred = CreateObject("Scripting.Dictionary")
    Set state = CreateObject("Scripting.Dictionary")
    Set stackIndex = CreateObject("Scripting.Dictionary")
    Set stack = New Collection
    Set cycleIds = CreateObject("Scripting.Dictionary")

    For Each idKey In validIds.Keys

        Set colChildren = New Collection
        childrenByPred.Add CStr(idKey), colChildren

        state(CStr(idKey)) = 0

    Next idKey

    If Not linksBySuccId Is Nothing Then

        For Each succId In linksBySuccId.Keys

            If validIds.Exists(CStr(succId)) Then

                For Each oneLink In linksBySuccId(CStr(succId))

                    predId = Core_GetLinkPredId(oneLink)

                    If predId <> "" Then
                        If validIds.Exists(predId) Then

                            If Not childrenByPred.Exists(predId) Then
                                Set colChildren = New Collection
                                childrenByPred.Add predId, colChildren
                            End If

                            childrenByPred(predId).Add CStr(succId)

                        End If
                    End If

                Next oneLink

            End If

        Next succId

    End If

    For Each idKey In validIds.Keys

        If CLng(state(CStr(idKey))) = 0 Then

            CalcBridge_PreCore_DFS_Cycle _
                CStr(idKey), childrenByPred, state, stackIndex, stack, cycleIds

            If cycleIds.Count > 0 Then Exit For

        End If

    Next idKey

    Set CalcBridge_PreCore_DetectCycleIds = cycleIds

End Function

Private Sub CalcBridge_PreCore_DFS_Cycle( _
    ByVal currentId As String, _
    ByVal childrenByPred As Object, _
    ByVal state As Object, _
    ByVal stackIndex As Object, _
    ByVal stack As Collection, _
    ByVal cycleIds As Object)

    Dim childId As Variant
    Dim i As Long
    Dim startIdx As Long

    If cycleIds.Count > 0 Then Exit Sub

    state(currentId) = 1

    stack.Add currentId
    stackIndex(currentId) = stack.Count

    If childrenByPred.Exists(currentId) Then

        For Each childId In childrenByPred(currentId)

            If cycleIds.Count > 0 Then Exit For

            If Not state.Exists(CStr(childId)) Then
                state(CStr(childId)) = 0
            End If

            If CLng(state(CStr(childId))) = 0 Then

                CalcBridge_PreCore_DFS_Cycle _
                    CStr(childId), childrenByPred, state, stackIndex, stack, cycleIds

            ElseIf CLng(state(CStr(childId))) = 1 Then

                If stackIndex.Exists(CStr(childId)) Then

                    startIdx = CLng(stackIndex(CStr(childId)))

                    For i = startIdx To stack.Count
                        cycleIds(CStr(stack(i))) = True
                    Next i

                    Exit For

                End If

            End If

        Next childId

    End If

    state(currentId) = 2

    If stack.Count > 0 Then
        If CStr(stack(stack.Count)) = currentId Then
            stack.Remove stack.Count
        End If
    End If

    If stackIndex.Exists(currentId) Then
        stackIndex.Remove currentId
    End If

End Sub

Private Function CalcBridge_PreCore_BuildIdToWbsFromCalc( _
    ByVal tblCalc As ListObject, _
    ByVal mapCalc As Object) As Object

    Dim idToWbs As Object
    Dim arr As Variant
    Dim r As Long
    Dim idVal As String
    Dim wbsVal As String
    Dim taskNameVal As String

    Set idToWbs = CreateObject("Scripting.Dictionary")

    If tblCalc Is Nothing Then
        Set CalcBridge_PreCore_BuildIdToWbsFromCalc = idToWbs
        Exit Function
    End If

    If tblCalc.DataBodyRange Is Nothing Then
        Set CalcBridge_PreCore_BuildIdToWbsFromCalc = idToWbs
        Exit Function
    End If

    If Not mapCalc.Exists("ID") Then
        Set CalcBridge_PreCore_BuildIdToWbsFromCalc = idToWbs
        Exit Function
    End If

    arr = tblCalc.DataBodyRange.value

    For r = 1 To UBound(arr, 1)

        idVal = Trim$(CStr(arr(r, mapCalc("ID"))))

        If idVal <> "" Then

            If mapCalc.Exists("WBS") Then
                wbsVal = Trim$(CStr(arr(r, mapCalc("WBS"))))
            Else
                wbsVal = "-"
            End If

            idToWbs(idVal) = wbsVal

        End If

    Next r

    Set CalcBridge_PreCore_BuildIdToWbsFromCalc = idToWbs

End Function

Private Sub CalcBridge_PreCore_MarkErrorIdsInCalc( _
    ByVal tblCalc As ListObject, _
    ByVal mapCalc As Object, _
    ByVal errorIds As Object, _
    ByVal errMsg As String)

    Dim arr As Variant
    Dim r As Long
    Dim idVal As String

    If tblCalc Is Nothing Then Exit Sub
    If tblCalc.DataBodyRange Is Nothing Then Exit Sub

    If Not mapCalc.Exists("ID") Then Exit Sub
    If Not mapCalc.Exists("Error flag") Then Exit Sub
    If Not mapCalc.Exists("ErrorMsg") Then Exit Sub

    arr = tblCalc.DataBodyRange.value

    For r = 1 To UBound(arr, 1)

        idVal = Trim$(CStr(arr(r, mapCalc("ID"))))

        If idVal <> "" Then
            If errorIds.Exists(idVal) Then
                arr(r, mapCalc("Error flag")) = "ERROR"
                arr(r, mapCalc("ErrorMsg")) = errMsg
            End If
        End If

    Next r

    tblCalc.DataBodyRange.value = arr

End Sub

Public Sub CalcBridge_ComputeDeadlineAnalytics( _
    ByVal tblCalc As ListObject, _
    ByVal mapCalc As Object, _
    Optional ByVal consoleMessages As Collection)

    Dim dataArr As Variant
    Dim outDeadlineFloat() As Variant
    Dim exceededIds As Object
    Dim idToWbs As Object
    Dim ackTokensById As Object
    Dim ackTokens As String
    Dim r As Long
    Dim rowCount As Long
    Dim idVal As String
    Dim wbsVal As String
    Dim taskNameVal As String
    Dim deadlineVal As Variant
    Dim calcFinishVal As Variant
    Dim deadlineFloatVal As Double
    Dim eventHashVal As String

    On Error GoTo SafeExit

    If tblCalc Is Nothing Then Exit Sub
    If tblCalc.DataBodyRange Is Nothing Then Exit Sub
    If mapCalc Is Nothing Then Exit Sub

    If Not mapCalc.Exists("ID") Then Exit Sub
    If Not mapCalc.Exists("WBS") Then Exit Sub
    If Not mapCalc.Exists("Task Name") Then Exit Sub
    If Not mapCalc.Exists("Deadline") Then Exit Sub
    If Not mapCalc.Exists("Deadline Float") Then Exit Sub
    If Not mapCalc.Exists("Calculated Finish") Then Exit Sub

    dataArr = tblCalc.DataBodyRange.value
    rowCount = UBound(dataArr, 1)
    ReDim outDeadlineFloat(1 To rowCount, 1 To 1)

    Set exceededIds = CreateObject("Scripting.Dictionary")
    Set idToWbs = CreateObject("Scripting.Dictionary")
    Set ackTokensById = CreateObject("Scripting.Dictionary")

    For r = 1 To rowCount

        idVal = Trim$(CStr(dataArr(r, mapCalc("ID"))))
        wbsVal = NormalizeWBS(dataArr(r, mapCalc("WBS")))
        taskNameVal = Trim$(CStr(dataArr(r, mapCalc("Task Name"))))
        idToWbs(idVal) = wbsVal

        deadlineVal = dataArr(r, mapCalc("Deadline"))
        calcFinishVal = dataArr(r, mapCalc("Calculated Finish"))

        If HasValue(deadlineVal) And HasValue(calcFinishVal) Then
            If IsDate(deadlineVal) And IsDate(calcFinishVal) Then

                deadlineFloatVal = CDbl(CDate(deadlineVal)) - CDbl(CDate(calcFinishVal))
                outDeadlineFloat(r, 1) = deadlineFloatVal

                If deadlineFloatVal < 0# Then
                    If idVal <> "" Then
                        exceededIds(idVal) = True

                        eventHashVal = BuildPlanningEventHash( _
                            "WARNING", _
                            "DEADLINE_EXCEEDED", _
                            "Deadline depassee", _
                            "Deadline exceeded", _
                            "la date calculee finit apres la deadline", _
                            "calculated finish is after the deadline", _
                            "CalcBridge_ComputeDeadlineAnalytics", _
                            "CALC", _
                            "tbl_CALC", _
                            idVal, _
                            wbsVal, _
                            taskNameVal)

                        ackTokensById(idVal) = BuildPlanningWarningAckToken("DEADLINE_EXCEEDED", eventHashVal)

                        LogPlanningEvent _
                            "WARNING", _
                            "DEADLINE_EXCEEDED", _
                            eventHashVal, _
                            "Deadline depassee", _
                            "Deadline exceeded", _
                            "la date calculee finit apres la deadline", _
                            "calculated finish is after the deadline", _
                            "CalcBridge_ComputeDeadlineAnalytics", _
                            "CALC", _
                            "tbl_CALC", _
                            idVal, _
                            wbsVal, _
                            taskNameVal
                    End If
                End If
            End If
        End If

    Next r

    tblCalc.ListColumns("Deadline Float").DataBodyRange.value = outDeadlineFloat

    If exceededIds.Count > 0 Then
        ackTokens = CalcBridge_BuildWarningAckTokenList(exceededIds, ackTokensById)

        If consoleMessages Is Nothing Then
            Set consoleMessages = New Collection

            CalcBridge_AddGroupedWarningToCollection consoleMessages, exceededIds, idToWbs, _
                "Deadline depassee", _
                "la date calculee finit apres la deadline", _
                "Deadline exceeded", _
                "calculated finish is after the deadline", _
                True, _
                ackTokens

            CalcBridge_ShowPlanningConsole consoleMessages
        Else
            CalcBridge_AddGroupedWarningToCollection consoleMessages, exceededIds, idToWbs, _
                "Deadline depassee", _
                "la date calculee finit apres la deadline", _
                "Deadline exceeded", _
                "calculated finish is after the deadline", _
                True, _
                ackTokens
        End If
    End If

SafeExit:
End Sub
Private Sub CalcBridge_ShowParentDateWarnings( _
    ByVal tblCalc As ListObject, _
    ByVal mapCalc As Object, _
    Optional ByVal consoleMessages As Collection)

    Dim dataArr As Variant
    Dim rowById As Object
    Dim parentIds As Object
    Dim warnParentDates As Object
    Dim idToWbs As Object
    Dim ackTokensById As Object
    Dim ackTokens As String
    Dim eventHashVal As String

    Dim key As Variant
    Dim rowIdx As Long
    Dim idVal As String
    Dim wbsVal As String
    Dim taskNameVal As String

    On Error GoTo SafeExit

    If tblCalc Is Nothing Then Exit Sub
    If tblCalc.DataBodyRange Is Nothing Then Exit Sub

    If Not mapCalc.Exists("ID") Then Exit Sub
    If Not mapCalc.Exists("Actual Start") Then Exit Sub
    If Not mapCalc.Exists("Actual Finish") Then Exit Sub
    If Not mapCalc.Exists("Forecast Start") Then Exit Sub
    If Not mapCalc.Exists("Forecast Finish") Then Exit Sub
    If Not mapCalc.Exists("Baseline Start") Then Exit Sub
    If Not mapCalc.Exists("Baseline Duration") Then Exit Sub
    If Not mapCalc.Exists("Baseline Finish") Then Exit Sub

    dataArr = tblCalc.DataBodyRange.value

    Set rowById = Core_BuildRowById(dataArr, mapCalc)
    Set parentIds = Core_BuildParentIds(dataArr, mapCalc, rowById)

    Set warnParentDates = CreateObject("Scripting.Dictionary")
    Set idToWbs = CreateObject("Scripting.Dictionary")
    Set ackTokensById = CreateObject("Scripting.Dictionary")

    For Each key In rowById.Keys

        idVal = CStr(key)
        rowIdx = CLng(rowById(idVal))

        If mapCalc.Exists("WBS") Then
            wbsVal = NormalizeWBS(dataArr(rowIdx, mapCalc("WBS")))
        Else
            wbsVal = "-"
        End If

        idToWbs(idVal) = wbsVal

        If parentIds.Exists(idVal) Then

            If HasValue(dataArr(rowIdx, mapCalc("Actual Start"))) Or _
               HasValue(dataArr(rowIdx, mapCalc("Actual Finish"))) Or _
               HasValue(dataArr(rowIdx, mapCalc("Forecast Start"))) Or _
               HasValue(dataArr(rowIdx, mapCalc("Forecast Finish"))) Or _
               HasValue(dataArr(rowIdx, mapCalc("Baseline Start"))) Or _
               HasValue(dataArr(rowIdx, mapCalc("Baseline Duration"))) Or _
               HasValue(dataArr(rowIdx, mapCalc("Baseline Finish"))) Then

                warnParentDates(idVal) = True

                If mapCalc.Exists("Task Name") Then
                    taskNameVal = Trim$(CStr(dataArr(rowIdx, mapCalc("Task Name"))))
                Else
                    taskNameVal = vbNullString
                End If

                eventHashVal = BuildPlanningEventHash( _
                    "WARNING", _
                    "PARENT_DATES_IGNORED", _
                    "Dates saisies sur tache parent", _
                    "Dates entered on summary task", _
                    "les valeurs sont ignorees, calcul par les taches enfants", _
                    "values are ignored and calculated from child tasks", _
                    "CalcBridge_ShowParentDateWarnings", _
                    "CALC", _
                    "tbl_CALC", _
                    idVal, _
                    wbsVal, _
                    taskNameVal)
                ackTokensById(idVal) = BuildPlanningWarningAckToken("PARENT_DATES_IGNORED", eventHashVal)

                On Error Resume Next

                LogPlanningEvent _
                    "WARNING", _
                    "PARENT_DATES_IGNORED", _
                    eventHashVal, _
                    "Dates saisies sur tache parent", _
                    "Dates entered on summary task", _
                    "les valeurs sont ignorees, calcul par les taches enfants", _
                    "values are ignored and calculated from child tasks", _
                    "CalcBridge_ShowParentDateWarnings", _
                    "CALC", _
                    "tbl_CALC", _
                    idVal, _
                    wbsVal, _
                    taskNameVal
                On Error GoTo SafeExit

            End If

        End If

    Next key

    If warnParentDates.Count > 0 Then

        ackTokens = CalcBridge_BuildWarningAckTokenList(warnParentDates, ackTokensById)

        If consoleMessages Is Nothing Then
            Set consoleMessages = New Collection

            CalcBridge_AddGroupedWarningToCollection consoleMessages, warnParentDates, idToWbs, _
                "Dates saisies sur tâche parent", _
                "les valeurs sont ignorées, calcul par les tâches enfants", _
                "Dates entered on summary task", _
                "values are ignored and calculated from child tasks", _
                True, _
                ackTokens

            CalcBridge_ShowPlanningConsole consoleMessages
        Else
            CalcBridge_AddGroupedWarningToCollection consoleMessages, warnParentDates, idToWbs, _
                "Dates saisies sur tâche parent", _
                "les valeurs sont ignorées, calcul par les tâches enfants", _
                "Dates entered on summary task", _
                "values are ignored and calculated from child tasks", _
                True, _
                ackTokens
        End If

    End If

SafeExit:
End Sub

Private Function BuildParentByIdMap_FromCalc( _
    ByRef dataArr As Variant, _
    ByVal mapCalc As Object, _
    ByVal rowById As Object) As Object

    Dim parentById As Object
    Dim idKey As Variant
    Dim rowIdx As Long
    Dim parentId As String

    Set parentById = CreateObject("Scripting.Dictionary")

    If Not mapCalc.Exists("ParentID") Then
        Set BuildParentByIdMap_FromCalc = parentById
        Exit Function
    End If

    For Each idKey In rowById.Keys
        rowIdx = CLng(rowById(CStr(idKey)))
        parentId = Trim$(CStr(dataArr(rowIdx, mapCalc("ParentID"))))

        If parentId <> "" Then
            parentById(CStr(idKey)) = parentId
        End If
    Next idKey

    Set BuildParentByIdMap_FromCalc = parentById

End Function

Public Function Build_Successor_Map(ByVal linksBySuccId As Object) As Object

    Dim succByPred As Object
    Dim succId As Variant
    Dim oneLink As Variant
    Dim predId As String

    Set succByPred = CreateObject("Scripting.Dictionary")

    If linksBySuccId Is Nothing Then
        Set Build_Successor_Map = succByPred
        Exit Function
    End If

    For Each succId In linksBySuccId.Keys

        For Each oneLink In linksBySuccId(CStr(succId))

            predId = Core_GetLinkPredId(oneLink)

            If predId <> "" Then
                If Not succByPred.Exists(predId) Then
                    Set succByPred(predId) = New Collection
                End If

                succByPred(predId).Add CStr(succId)
            End If

        Next oneLink

    Next succId

    Set Build_Successor_Map = succByPred

End Function

Public Function Get_Impacted_Descendants( _
    ByVal changedIds As Object, _
    ByVal succByPred As Object) As Object

    Dim impacted As Object
    Dim queue As Collection
    Dim idVal As Variant
    Dim currentId As String
    Dim succId As Variant

    Set impacted = CreateObject("Scripting.Dictionary")
    Set queue = New Collection

    If changedIds Is Nothing Then
        Set Get_Impacted_Descendants = impacted
        Exit Function
    End If

    For Each idVal In changedIds.Keys
        impacted(CStr(idVal)) = True
        queue.Add CStr(idVal)
    Next idVal

    Do While queue.Count > 0

        currentId = CStr(queue(1))
        queue.Remove 1

        If Not succByPred Is Nothing Then
            If succByPred.Exists(currentId) Then

                For Each succId In succByPred(currentId)

                    If Not impacted.Exists(CStr(succId)) Then
                        impacted(CStr(succId)) = True
                        queue.Add CStr(succId)
                    End If

                Next succId

            End If
        End If

    Loop

    Set Get_Impacted_Descendants = impacted

End Function

Private Sub WriteCoreOutputsToCalc_Partial( _
    ByVal tblCalc As ListObject, _
    ByVal mapCalc As Object, _
    ByRef dataArr As Variant, _
    ByVal impactedIds As Object)

    Dim rowById As Object
    Dim idVal As Variant
    Dim rowIdx As Long

    If tblCalc Is Nothing Then Exit Sub
    If tblCalc.DataBodyRange Is Nothing Then Exit Sub
    If impactedIds Is Nothing Then Exit Sub
    If impactedIds.Count = 0 Then Exit Sub

    If Not mapCalc.Exists("ID") Then Exit Sub
    If Not mapCalc.Exists("Calculated Start") Then Exit Sub
    If Not mapCalc.Exists("Calculated Finish") Then Exit Sub
    If Not mapCalc.Exists("Calculated Duration") Then Exit Sub
    If Not mapCalc.Exists("Error flag") Then Exit Sub
    If Not mapCalc.Exists("ErrorMsg") Then Exit Sub

    Set rowById = Core_BuildRowById(dataArr, mapCalc)

    For Each idVal In impactedIds.Keys

        If rowById.Exists(CStr(idVal)) Then

            rowIdx = CLng(rowById(CStr(idVal)))

            tblCalc.DataBodyRange.Cells(rowIdx, mapCalc("Calculated Start")).value = _
                dataArr(rowIdx, mapCalc("Calculated Start"))

            tblCalc.DataBodyRange.Cells(rowIdx, mapCalc("Calculated Finish")).value = _
                dataArr(rowIdx, mapCalc("Calculated Finish"))

            tblCalc.DataBodyRange.Cells(rowIdx, mapCalc("Calculated Duration")).value = _
                dataArr(rowIdx, mapCalc("Calculated Duration"))

            tblCalc.DataBodyRange.Cells(rowIdx, mapCalc("Error flag")).value = _
                dataArr(rowIdx, mapCalc("Error flag"))

            tblCalc.DataBodyRange.Cells(rowIdx, mapCalc("ErrorMsg")).value = _
                dataArr(rowIdx, mapCalc("ErrorMsg"))

        End If

    Next idVal

End Sub

Private Sub WriteCoreDrivingLogicToCalc_Partial( _
    ByVal tblCalc As ListObject, _
    ByVal mapCalc As Object, _
    ByRef dataArr As Variant, _
    ByVal impactedIds As Object)

    Dim rowById As Object
    Dim parentIds As Object
    Dim idVal As Variant
    Dim rowIdx As Long

    Dim baselineStart As Variant
    Dim actualStart As Variant
    Dim actualFinish As Variant
    Dim forecastStart As Variant
    Dim forecastFinish As Variant
    Dim calcStart As Variant
    Dim drivingLogic As String

    If tblCalc Is Nothing Then Exit Sub
    If tblCalc.DataBodyRange Is Nothing Then Exit Sub
    If impactedIds Is Nothing Then Exit Sub
    If impactedIds.Count = 0 Then Exit Sub

    If Not mapCalc.Exists("ID") Then Exit Sub
    If Not mapCalc.Exists("Driving Logic") Then Exit Sub
    If Not mapCalc.Exists("Baseline Start") Then Exit Sub
    If Not mapCalc.Exists("Actual Start") Then Exit Sub
    If Not mapCalc.Exists("Actual Finish") Then Exit Sub
    If Not mapCalc.Exists("Forecast Start") Then Exit Sub
    If Not mapCalc.Exists("Forecast Finish") Then Exit Sub
    If Not mapCalc.Exists("Calculated Start") Then Exit Sub

    Set rowById = Core_BuildRowById(dataArr, mapCalc)
    Set parentIds = Core_BuildParentIds(dataArr, mapCalc, rowById)

    For Each idVal In impactedIds.Keys

        If rowById.Exists(CStr(idVal)) Then

            rowIdx = CLng(rowById(CStr(idVal)))
            drivingLogic = ""

            If parentIds.Exists(CStr(idVal)) Then

                drivingLogic = "SUMMARY"

            ElseIf CalcBridge_IsLevelOfEffortRow(dataArr, mapCalc, rowIdx) Then

                drivingLogic = "LOE"

            Else

                baselineStart = dataArr(rowIdx, mapCalc("Baseline Start"))
                actualStart = dataArr(rowIdx, mapCalc("Actual Start"))
                actualFinish = dataArr(rowIdx, mapCalc("Actual Finish"))
                forecastStart = dataArr(rowIdx, mapCalc("Forecast Start"))
                forecastFinish = dataArr(rowIdx, mapCalc("Forecast Finish"))
                calcStart = dataArr(rowIdx, mapCalc("Calculated Start"))

                If CalcBridge_IsDrivenByActiveConstraint(dataArr, mapCalc, rowIdx) Then
                    drivingLogic = "CONSTRAINT"
                ElseIf HasValue(actualStart) Or HasValue(actualFinish) Then
                    drivingLogic = "ACTUAL"
                ElseIf HasValue(forecastStart) Or HasValue(forecastFinish) Then
                    drivingLogic = "FORECAST"
                ElseIf HasValue(calcStart) And HasValue(baselineStart) Then
                    If CDbl(calcStart) > CDbl(baselineStart) Then
                        drivingLogic = "DEPENDENCY"
                    Else
                        drivingLogic = "BASELINE"
                    End If
                Else
                    drivingLogic = "BASELINE"
                End If

            End If

            tblCalc.DataBodyRange.Cells(rowIdx, mapCalc("Driving Logic")).value = drivingLogic

        End If

    Next idVal

End Sub

Public Sub CalcBridge_RunAnalyticsAndPush( _
    ByVal tblCalc As ListObject, _
    ByVal mapCalc As Object, _
    ByVal linksBySuccId As Object, _
    Optional ByVal consoleMessages As Collection)

    Dim dataArr As Variant
    Dim rowById As Object
    Dim parentIds As Object
    Dim validIds As Object
    Dim directChildrenById As Object
    Dim predsById As Object
    Dim childrenById As Object
    Dim topoOrder As Collection
    Dim idToWbs As Object
    Dim outCriticalREX() As Variant
    Dim errMissingBaselineForREX As Object

    On Error GoTo ErrHandler

    If tblCalc Is Nothing Then Exit Sub
    If tblCalc.DataBodyRange Is Nothing Then Exit Sub
    If mapCalc Is Nothing Then Exit Sub

    dataArr = tblCalc.DataBodyRange.value

    Set rowById = Core_BuildRowById(dataArr, mapCalc)
    Set parentIds = Core_BuildParentIds(dataArr, mapCalc, rowById)

    Set validIds = CalcBridge_BuildLeafIds(rowById, parentIds, dataArr, mapCalc)

    Set directChildrenById = CalcBridge_BuildDirectChildrenById(dataArr, mapCalc, rowById)
    Set predsById = CalcBridge_BuildEmptyCollections(validIds)
    Set childrenById = CalcBridge_BuildEmptyCollections(validIds)
    Set idToWbs = CalcBridge_BuildIdToWbsFromData(dataArr, mapCalc, rowById)
    Set errMissingBaselineForREX = CreateObject("Scripting.Dictionary")

    CalcBridge_FillPredsAndChildrenFromLinks linksBySuccId, validIds, predsById, childrenById
    Set topoOrder = CalcBridge_TopologicalOrder(validIds, predsById, childrenById)

    If topoOrder.Count <> validIds.Count Then

        CalcBridge_AddOrShowConsoleMessage consoleMessages, "WARNING", _
            "Analytics non calculées : ordre topologique incomplet." & vbCrLf & _
            "-> vérifier les cycles ou la reconstruction tbl_LOGIC_LINKS.", _
            "Analytics not calculated: incomplete topological order." & vbCrLf & _
            "-> check cycles or tbl_LOGIC_LINKS rebuild."

        Exit Sub
    End If

    ComputeCurrentFloatAndCritical _
        tblCalc, mapCalc, rowById, childrenById, directChildrenById, parentIds, validIds, topoOrder, _
        consoleMessages

    ReDim outCriticalREX(1 To tblCalc.ListRows.Count, 1 To 1)

    ComputeLongestPath _
        tblCalc, mapCalc, rowById, predsById, childrenById, validIds

    ComputeCriticalPathREX _
        tblCalc, mapCalc, rowById, predsById, childrenById, directChildrenById, _
        parentIds, validIds, topoOrder, idToWbs, outCriticalREX, errMissingBaselineForREX, _
        consoleMessages

    If errMissingBaselineForREX.Count > 0 Then
        If consoleMessages Is Nothing Then
            Set consoleMessages = New Collection
            CalcBridge_AddConsoleMessage consoleMessages, "WARNING", _
                CalcBridge_BuildMissingBaselineRexMessage(errMissingBaselineForREX, idToWbs)
            CalcBridge_ShowPlanningConsole consoleMessages
        Else
            CalcBridge_AddConsoleMessage consoleMessages, "WARNING", _
                CalcBridge_BuildMissingBaselineRexMessage(errMissingBaselineForREX, idToWbs)
        End If
    End If

    Exit Sub

ErrHandler:
    Err.Raise Err.Number, "CalcBridge_RunAnalyticsAndPush", Err.Description

End Sub

Private Function CalcBridge_BuildLeafIds( _
    ByVal rowById As Object, _
    ByVal parentIds As Object, _
    Optional ByRef dataArr As Variant, _
    Optional ByVal mapCalc As Object) As Object

    Dim d As Object
    Dim key As Variant
    Dim rowIdx As Long

    Set d = CreateObject("Scripting.Dictionary")

    If rowById Is Nothing Then
        Set CalcBridge_BuildLeafIds = d
        Exit Function
    End If

    For Each key In rowById.Keys

        If Not parentIds Is Nothing Then
            If parentIds.Exists(CStr(key)) Then GoTo NextKey
        End If

        rowIdx = CLng(rowById(CStr(key)))

        'LOE is not part of analytics / critical path / floats / REX.
        'This exclusion happens before REX baseline-duration validation.
        If CalcBridge_IsLevelOfEffortRow(dataArr, mapCalc, rowIdx) Then
            GoTo NextKey
        End If

        d(CStr(key)) = True

NextKey:
    Next key

    Set CalcBridge_BuildLeafIds = d

End Function

Private Function CalcBridge_BuildEmptyCollections(ByVal ids As Object) As Object

    Dim d As Object
    Dim key As Variant

    Set d = CreateObject("Scripting.Dictionary")

    For Each key In ids.Keys
        Set d(CStr(key)) = New Collection
    Next key

    Set CalcBridge_BuildEmptyCollections = d

End Function

Private Function CalcBridge_BuildDirectChildrenById( _
    ByRef dataArr As Variant, _
    ByVal mapCalc As Object, _
    ByVal rowById As Object) As Object

    Dim d As Object
    Dim key As Variant
    Dim r As Long
    Dim idVal As String
    Dim parentId As String

    Set d = CreateObject("Scripting.Dictionary")

    For Each key In rowById.Keys
        Set d(CStr(key)) = New Collection
    Next key

    If Not mapCalc.Exists("ID") Then
        Set CalcBridge_BuildDirectChildrenById = d
        Exit Function
    End If

    If Not mapCalc.Exists("ParentID") Then
        Set CalcBridge_BuildDirectChildrenById = d
        Exit Function
    End If

    For r = 1 To UBound(dataArr, 1)
        idVal = Trim$(CStr(dataArr(r, mapCalc("ID"))))
        parentId = Trim$(CStr(dataArr(r, mapCalc("ParentID"))))

        If idVal <> "" And parentId <> "" Then
            If d.Exists(parentId) Then
                d(parentId).Add idVal
            End If
        End If
    Next r

    Set CalcBridge_BuildDirectChildrenById = d

End Function

Private Sub CalcBridge_FillPredsAndChildrenFromLinks( _
    ByVal linksBySuccId As Object, _
    ByVal validIds As Object, _
    ByVal predsById As Object, _
    ByVal childrenById As Object)

    Dim succId As Variant
    Dim oneLink As Variant
    Dim predId As String

    If linksBySuccId Is Nothing Then Exit Sub

    For Each succId In linksBySuccId.Keys
        If validIds.Exists(CStr(succId)) Then
            For Each oneLink In linksBySuccId(CStr(succId))
                predId = Core_GetLinkPredId(oneLink)

                If predId <> "" Then
                    If validIds.Exists(predId) Then
                        predsById(CStr(succId)).Add predId
                        childrenById(predId).Add CStr(succId)
                    End If
                End If
            Next oneLink
        End If
    Next succId

End Sub

Private Function CalcBridge_TopologicalOrder( _
    ByVal validIds As Object, _
    ByVal predsById As Object, _
    ByVal childrenById As Object) As Collection

    Dim q As Collection
    Dim topo As Collection
    Dim indegree As Object
    Dim key As Variant
    Dim childId As Variant
    Dim currentId As String

    Set q = New Collection
    Set topo = New Collection
    Set indegree = CreateObject("Scripting.Dictionary")

    For Each key In validIds.Keys
        indegree(CStr(key)) = predsById(CStr(key)).Count
        If CLng(indegree(CStr(key))) = 0 Then q.Add CStr(key)
    Next key

    Do While q.Count > 0
        currentId = CStr(q(1))
        q.Remove 1
        topo.Add currentId

        If childrenById.Exists(currentId) Then
            For Each childId In childrenById(currentId)
                indegree(CStr(childId)) = CLng(indegree(CStr(childId))) - 1
                If CLng(indegree(CStr(childId))) = 0 Then q.Add CStr(childId)
            Next childId
        End If
    Loop

    Set CalcBridge_TopologicalOrder = topo

End Function

Private Function CalcBridge_BuildIdToWbsFromData( _
    ByRef dataArr As Variant, _
    ByVal mapCalc As Object, _
    ByVal rowById As Object) As Object

    Dim d As Object
    Dim key As Variant
    Dim rowIdx As Long

    Set d = CreateObject("Scripting.Dictionary")

    For Each key In rowById.Keys
        rowIdx = CLng(rowById(CStr(key)))

        If mapCalc.Exists("WBS") Then
            d(CStr(key)) = Trim$(CStr(dataArr(rowIdx, mapCalc("WBS"))))
        Else
            d(CStr(key)) = "-"
        End If
    Next key

    Set CalcBridge_BuildIdToWbsFromData = d

End Function

Public Sub Push_Calculated_Back_To_WBS_Partial(ByVal impactedIds As Object)

    Dim wsWBS As Worksheet
    Dim wsCalc As Worksheet
    Dim tblWBS As ListObject
    Dim tblCalc As ListObject

    Dim mapWBS As Object
    Dim mapCalc As Object
    Dim calcRowById As Object

    Dim allowedFields As Variant
    Dim r As Long
    Dim i As Long
    Dim id As String
    Dim calcRow As Long

    On Error GoTo SafeExit

    If impactedIds Is Nothing Then Exit Sub
    If impactedIds.Count = 0 Then Exit Sub

    Set wsWBS = ThisWorkbook.Worksheets("WBS")
    Set wsCalc = ThisWorkbook.Worksheets("CALC")

    Set tblWBS = wsWBS.ListObjects("tbl_WBS")
    Set tblCalc = wsCalc.ListObjects("tbl_CALC")

    If tblWBS.DataBodyRange Is Nothing Then Exit Sub
    If tblCalc.DataBodyRange Is Nothing Then Exit Sub

    Set mapWBS = Core_BuildColumnMap_FromListObject(tblWBS)
    Set mapCalc = Core_BuildColumnMap_FromListObject(tblCalc)
    Set calcRowById = CreateObject("Scripting.Dictionary")

    If Not mapWBS.Exists("ID") Then
        Err.Raise vbObjectError + 2101, "Push_Calculated_Back_To_WBS_Partial", _
            "Missing column in tbl_WBS: ID"
    End If

    If Not mapCalc.Exists("ID") Then
        Err.Raise vbObjectError + 2102, "Push_Calculated_Back_To_WBS_Partial", _
            "Missing column in tbl_CALC: ID"
    End If

    allowedFields = Array( _
        "Calculated Start", _
        "Calculated Finish", _
        "Driving Logic", _
        "Deadline Float" _
    )

    For i = LBound(allowedFields) To UBound(allowedFields)

        If Not mapWBS.Exists(CStr(allowedFields(i))) Then
            Err.Raise vbObjectError + 2110 + i, "Push_Calculated_Back_To_WBS_Partial", _
                "Missing output column in tbl_WBS: " & CStr(allowedFields(i))
        End If

        If Not mapCalc.Exists(CStr(allowedFields(i))) Then
            Err.Raise vbObjectError + 2120 + i, "Push_Calculated_Back_To_WBS_Partial", _
                "Missing output column in tbl_CALC: " & CStr(allowedFields(i))
        End If

    Next i

    For r = 1 To tblCalc.ListRows.Count
        id = Trim$(CStr(tblCalc.DataBodyRange.Cells(r, mapCalc("ID")).value))
        If id <> "" Then
            calcRowById(id) = r
        End If
    Next r

    BeginAuthorizedWBSWrite "Push_Calculated_Back_To_WBS_Partial", allowedFields

    For r = 1 To tblWBS.ListRows.Count

        id = Trim$(CStr(tblWBS.DataBodyRange.Cells(r, mapWBS("ID")).value))

        If id <> "" Then
            If impactedIds.Exists(id) Then

                If calcRowById.Exists(id) Then

                    calcRow = CLng(calcRowById(id))

                    For i = LBound(allowedFields) To UBound(allowedFields)
                        tblWBS.DataBodyRange.Cells(r, mapWBS(CStr(allowedFields(i)))).value = _
                            tblCalc.DataBodyRange.Cells(calcRow, mapCalc(CStr(allowedFields(i)))).value
                    Next i

                Else

                    For i = LBound(allowedFields) To UBound(allowedFields)
                        tblWBS.DataBodyRange.Cells(r, mapWBS(CStr(allowedFields(i)))).ClearContents
                    Next i

                End If

            End If
        End If

    Next r

SafeExit:
    EndAuthorizedWBSWrite

    If Err.Number <> 0 Then
        CalcBridge_ShowSingleConsoleMessage _
            "STOP", _
            "Erreur dans Push_Calculated_Back_To_WBS_Partial : " & Err.Description, _
            "Error in Push_Calculated_Back_To_WBS_Partial: " & Err.Description
    End If

End Sub

Private Function CalcBridge_IsLevelOfEffortRow( _
    ByRef dataArr As Variant, _
    ByVal mapCalc As Object, _
    ByVal rowIdx As Long) As Boolean

    Dim taskTypeVal As String

    On Error GoTo SafeExit

    CalcBridge_IsLevelOfEffortRow = False

    If mapCalc Is Nothing Then Exit Function
    If Not mapCalc.Exists("Task Type") Then Exit Function

    If rowIdx < LBound(dataArr, 1) Then Exit Function
    If rowIdx > UBound(dataArr, 1) Then Exit Function

    taskTypeVal = UCase$(Trim$(CStr(dataArr(rowIdx, mapCalc("Task Type")))))
    taskTypeVal = Replace$(taskTypeVal, "-", " ")
    taskTypeVal = Replace$(taskTypeVal, "_", " ")

    Do While InStr(1, taskTypeVal, "  ", vbBinaryCompare) > 0
        taskTypeVal = Replace$(taskTypeVal, "  ", " ")
    Loop

    CalcBridge_IsLevelOfEffortRow = _
        (taskTypeVal = "LOE") Or _
        (taskTypeVal = "LEVEL OF EFFORT") Or _
        (taskTypeVal = "LEVEL OF EFFORT TASK")

SafeExit:
End Function

Private Function CalcBridge_BuildMissingBaselineRexMessage( _
    ByVal missingIds As Object, _
    ByVal idToWbs As Object) As String

    Dim idsLine As String
    Dim wbsLine As String

    idsLine = CalcBridge_BuildInlineList(missingIds, 20)
    wbsLine = CalcBridge_BuildInlineWBSList(missingIds, idToWbs, 20)

    CalcBridge_BuildMissingBaselineRexMessage = _
        "FR:" & vbCrLf & _
        "Analytics REX partiellement non calculées : Baseline Duration manquante sur au moins une tâche feuille." & vbCrLf & _
        "-> compléter Baseline Duration sur les lignes listées ci-dessous." & vbCrLf & vbCrLf & _
        "IDs : " & idsLine & vbCrLf & _
        "WBS : " & wbsLine & vbCrLf & vbCrLf & _
        "EN:" & vbCrLf & _
        "REX analytics partially not calculated: Baseline Duration is missing on at least one leaf task." & vbCrLf & _
        "-> fill Baseline Duration on the lines listed below." & vbCrLf & vbCrLf & _
        "IDs: " & idsLine & vbCrLf & _
        "WBS: " & wbsLine

End Function

Private Function CalcBridge_PreCore_CheckLOEAsPredecessor( _
    ByVal tblCalc As ListObject, _
    ByVal mapCalc As Object, _
    Optional ByVal consoleMessages As Collection) As Boolean

    Dim wsCalc As Worksheet
    Dim tblLinks As ListObject
    Dim mapLinks As Object
    Dim arrCalc As Variant
    Dim arrLinks As Variant
    Dim rowById As Object
    Dim idToWbs As Object
    Dim errIds As Object

    Dim i As Long
    Dim r As Long
    Dim succId As String
    Dim predId As String
    Dim predRow As Long

    On Error GoTo FailSafe

    CalcBridge_PreCore_CheckLOEAsPredecessor = False

    If tblCalc Is Nothing Then Exit Function
    If tblCalc.DataBodyRange Is Nothing Then Exit Function
    If mapCalc Is Nothing Then Exit Function

    If Not mapCalc.Exists("ID") Then Exit Function
    If Not mapCalc.Exists("Task Type") Then Exit Function

    Set wsCalc = ThisWorkbook.Worksheets("CALC")
    Set tblLinks = wsCalc.ListObjects("tbl_LOGIC_LINKS")

    If tblLinks Is Nothing Then Exit Function
    If tblLinks.DataBodyRange Is Nothing Then Exit Function

    Set mapLinks = CreateObject("Scripting.Dictionary")
    For i = 1 To tblLinks.ListColumns.Count
        mapLinks(tblLinks.ListColumns(i).Name) = i
    Next i

    If Not mapLinks.Exists("Succ ID") Then Exit Function
    If Not mapLinks.Exists("Pred ID") Then Exit Function

    arrCalc = tblCalc.DataBodyRange.value
    arrLinks = tblLinks.DataBodyRange.value

    Set rowById = Core_BuildRowById(arrCalc, mapCalc)
    Set idToWbs = CalcBridge_PreCore_BuildIdToWbsFromCalc(tblCalc, mapCalc)
    Set errIds = CreateObject("Scripting.Dictionary")

    For r = 1 To UBound(arrLinks, 1)

        succId = Trim$(CStr(arrLinks(r, mapLinks("Succ ID"))))
        predId = Trim$(CStr(arrLinks(r, mapLinks("Pred ID"))))

        If succId <> "" And predId <> "" Then
            If rowById.Exists(predId) Then
                predRow = CLng(rowById(predId))

                If CalcBridge_IsLevelOfEffortRow(arrCalc, mapCalc, predRow) Then
                    errIds(predId) = True
                    errIds(succId) = True
                End If
            End If
        End If

    Next r

    If errIds.Count > 0 Then

        CalcBridge_PreCore_MarkErrorIdsInCalc tblCalc, mapCalc, errIds, "LOE cannot be used as predecessor"

        If consoleMessages Is Nothing Then
            CalcBridge_ShowGroupedErrorMessage errIds, idToWbs, _
                "LOE utilisée comme prédécesseur", _
                "supprimer la LOE de la logique amont ; une LOE est pilotée par le réseau mais ne doit pas piloter d'autres tâches", _
                "LOE used as predecessor", _
                "remove the LOE from upstream logic; a LOE is driven by the network but must not drive other tasks"
        Else
            CalcBridge_AddGroupedStopToCollection consoleMessages, errIds, idToWbs, _
                "LOE utilisée comme prédécesseur", _
                "supprimer la LOE de la logique amont ; une LOE est pilotée par le réseau mais ne doit pas piloter d'autres tâches", _
                "LOE used as predecessor", _
                "remove the LOE from upstream logic; a LOE is driven by the network but must not drive other tasks"
        End If

        CalcBridge_PreCore_CheckLOEAsPredecessor = True

    End If

    Exit Function

FailSafe:
    CalcBridge_AddOrShowConsoleMessage consoleMessages, "STOP", _
        "Erreur pendant le contrôle des LOE utilisées comme prédécesseur." & vbCrLf & _
        "-> calcul arrêté avant écriture WBS.", _
        "Error while checking LOE used as predecessor." & vbCrLf & _
        "-> calculation stopped before WBS write."

    CalcBridge_PreCore_CheckLOEAsPredecessor = True

End Function

Private Function CalcBridge_IsProgressStrictlyBetweenZeroAndOne(ByVal v As Variant) As Boolean

    Dim p As Variant

    If Not HasValue(v) Then Exit Function

    p = NormalizePercentInput(v)
    If Not HasValue(p) Then Exit Function

    CalcBridge_IsProgressStrictlyBetweenZeroAndOne = _
        (CDbl(p) > 0.0000001 And CDbl(p) < 0.9999999)

End Function

Private Sub CalcBridge_AppendTaskTypeWarnings( _
    ByVal warningMessages As Collection, _
    ByVal tblWBS As ListObject, _
    ByVal mapWBS As Object, _
    ByVal tblCalc As ListObject, _
    ByVal mapCalc As Object)

    Dim arrWBS As Variant
    Dim arrCalc As Variant
    Dim idToWbs As Object
    Dim calcRowById As Object

    Dim warnLOEProgress As Object
    Dim warnLOEBaseline As Object
    Dim warnMilestoneProgress As Object
    Dim warnMilestoneDuration As Object
    Dim warnIgnoredCal As Object

    Dim r As Long
    Dim calcRow As Long
    Dim idVal As String
    Dim calVal As String

    On Error GoTo SafeExit

    If warningMessages Is Nothing Then Exit Sub
    If tblWBS Is Nothing Then Exit Sub
    If tblCalc Is Nothing Then Exit Sub
    If tblWBS.DataBodyRange Is Nothing Then Exit Sub
    If tblCalc.DataBodyRange Is Nothing Then Exit Sub
    If mapWBS Is Nothing Then Exit Sub
    If mapCalc Is Nothing Then Exit Sub

    If Not mapWBS.Exists("ID") Then Exit Sub
    If Not mapWBS.Exists("WBS") Then Exit Sub
    If Not mapWBS.Exists("Task Type") Then Exit Sub

    If Not mapCalc.Exists("ID") Then Exit Sub
    If Not mapCalc.Exists("Task Type") Then Exit Sub

    arrWBS = tblWBS.DataBodyRange.value
    arrCalc = tblCalc.DataBodyRange.value

    Set idToWbs = BuildIdToWbsFromWBS(tblWBS)
    Set calcRowById = Core_BuildRowById(arrCalc, mapCalc)

    Set warnLOEProgress = CreateObject("Scripting.Dictionary")
    Set warnLOEBaseline = CreateObject("Scripting.Dictionary")
    Set warnMilestoneProgress = CreateObject("Scripting.Dictionary")
    Set warnMilestoneDuration = CreateObject("Scripting.Dictionary")
    Set warnIgnoredCal = CreateObject("Scripting.Dictionary")

    For r = 1 To UBound(arrWBS, 1)

        idVal = Trim$(CStr(arrWBS(r, mapWBS("ID"))))
        If idVal = "" Then GoTo NextRow

        If mapWBS.Exists("Cal") Then
            calVal = NormalizeCalendarType(arrWBS(r, mapWBS("Cal")))
            If calVal = CALENDAR_5D Or calVal = CALENDAR_6D Then
                If calcRowById.Exists(idVal) Then
                    calcRow = CLng(calcRowById(idVal))
                    If Core_IsSummaryRow(arrCalc, calcRow, mapCalc) Or _
                       CalcBridge_IsLevelOfEffortRow(arrWBS, mapWBS, r) Or _
                       IsMilestoneTaskType(arrWBS, mapWBS, r) Then
                        warnIgnoredCal(idVal) = True
                    End If
                ElseIf CalcBridge_IsLevelOfEffortRow(arrWBS, mapWBS, r) Or _
                       IsMilestoneTaskType(arrWBS, mapWBS, r) Then
                    warnIgnoredCal(idVal) = True
                End If
            End If
        End If
        If CalcBridge_IsLevelOfEffortRow(arrWBS, mapWBS, r) Then

            If mapWBS.Exists("% Progress") Then
                If HasValue(arrWBS(r, mapWBS("% Progress"))) Then warnLOEProgress(idVal) = True
            End If

            If (mapWBS.Exists("Baseline Start") And HasValue(arrWBS(r, mapWBS("Baseline Start")))) Or _
               (mapWBS.Exists("Baseline Duration") And HasValue(arrWBS(r, mapWBS("Baseline Duration")))) Or _
               (mapWBS.Exists("Baseline Finish") And HasValue(arrWBS(r, mapWBS("Baseline Finish")))) Then
                warnLOEBaseline(idVal) = True
            End If

        ElseIf IsMilestoneTaskType(arrWBS, mapWBS, r) Then

            If mapWBS.Exists("% Progress") Then
                If CalcBridge_IsProgressStrictlyBetweenZeroAndOne(arrWBS(r, mapWBS("% Progress"))) Then
                    warnMilestoneProgress(idVal) = True
                End If
            End If

            If calcRowById.Exists(idVal) Then
                calcRow = CLng(calcRowById(idVal))

                If CalcBridge_RowHasDurationGreaterThanOne(arrCalc, mapCalc, calcRow) Then
                    warnMilestoneDuration(idVal) = True
                End If
            End If

        End If

NextRow:
    Next r

    If warnIgnoredCal.Count > 0 Then
        CalcBridge_AddGroupedWarningToCollection warningMessages, warnIgnoredCal, idToWbs, _
            "Calendrier ignoré sur tâche Summary / LOE / Milestone", _
            "le calendrier n'est utilisé que pour les tâches normales", _
            "Calendar ignored on Summary / LOE / Milestone task", _
            "calendars apply only to normal tasks"
    End If
    If warnLOEProgress.Count > 0 Then
        CalcBridge_AddGroupedWarningToCollection warningMessages, warnLOEProgress, idToWbs, _
            "% Progress renseigné sur LOE", _
            "le progress LOE est calculé automatiquement par date du jour ; la saisie manuelle est ignorée dans le Gantt", _
            "% Progress entered on LOE", _
            "LOE progress is automatically calculated from today's date; manual input is ignored in the Gantt"
    End If

    If warnLOEBaseline.Count > 0 Then
        CalcBridge_AddGroupedWarningToCollection warningMessages, warnLOEBaseline, idToWbs, _
            "Baseline renseignée sur LOE", _
            "une LOE doit être pilotée par ses liens SS/FF ; vérifier que la baseline saisie ne crée pas de confusion", _
            "Baseline entered on LOE", _
            "a LOE must be driven by its SS/FF links; check that entered baseline values are not misleading"
    End If

    If warnMilestoneProgress.Count > 0 Then
        CalcBridge_AddGroupedWarningToCollection warningMessages, warnMilestoneProgress, idToWbs, _
            "% Progress partiel renseigné sur Milestone", _
            "une milestone doit être à 0% ou 100% ; toute valeur intermédiaire doit être corrigée", _
            "Partial % Progress entered on Milestone", _
            "a milestone must be either 0% or 100%; any intermediate value should be corrected"
    End If

    If warnMilestoneDuration.Count > 0 Then
        CalcBridge_AddGroupedWarningToCollection warningMessages, warnMilestoneDuration, idToWbs, _
            "Durée supérieure à 1 jour sur Milestone", _
            "la tâche sera rendue comme milestone mais la durée saisie peut induire en erreur", _
            "Duration greater than 1 day on Milestone", _
            "the task will be rendered as a milestone but the entered duration may be misleading"
    End If

SafeExit:
End Sub

Private Sub CalcBridge_AddInfoMessage( _
    ByVal messages As Collection, _
    ByVal frText As String, _
    ByVal enText As String)

    Dim msg As String

    If messages Is Nothing Then Exit Sub

    msg = _
        "FR:" & vbCrLf & _
        frText & vbCrLf & vbCrLf & _
        "EN:" & vbCrLf & _
        enText

    CalcBridge_AddConsoleMessage messages, "INFO", msg

End Sub

Private Function CalcBridge_RowHasDurationGreaterThanOne( _
    ByRef dataArr As Variant, _
    ByVal mapCalc As Object, _
    ByVal rowIdx As Long) As Boolean

    If mapCalc.Exists("Calculated Duration") Then
        If IsNumeric(dataArr(rowIdx, mapCalc("Calculated Duration"))) Then
            If CDbl(dataArr(rowIdx, mapCalc("Calculated Duration"))) > 1# Then
                CalcBridge_RowHasDurationGreaterThanOne = True
                Exit Function
            End If
        End If
    End If

    If mapCalc.Exists("Baseline Duration") Then
        If IsNumeric(dataArr(rowIdx, mapCalc("Baseline Duration"))) Then
            If CDbl(dataArr(rowIdx, mapCalc("Baseline Duration"))) > 1# Then
                CalcBridge_RowHasDurationGreaterThanOne = True
                Exit Function
            End If
        End If
    End If

    If mapCalc.Exists("Actual Duration") Then
        If IsNumeric(dataArr(rowIdx, mapCalc("Actual Duration"))) Then
            If CDbl(dataArr(rowIdx, mapCalc("Actual Duration"))) > 1# Then
                CalcBridge_RowHasDurationGreaterThanOne = True
                Exit Function
            End If
        End If
    End If

    If CalcBridge_PairDurationGreaterThanOne(dataArr, mapCalc, rowIdx, "Baseline Start", "Baseline Finish") Then
        CalcBridge_RowHasDurationGreaterThanOne = True
        Exit Function
    End If

    If CalcBridge_PairDurationGreaterThanOne(dataArr, mapCalc, rowIdx, "Actual Start", "Actual Finish") Then
        CalcBridge_RowHasDurationGreaterThanOne = True
        Exit Function
    End If

    If CalcBridge_PairDurationGreaterThanOne(dataArr, mapCalc, rowIdx, "Forecast Start", "Forecast Finish") Then
        CalcBridge_RowHasDurationGreaterThanOne = True
        Exit Function
    End If

    If CalcBridge_PairDurationGreaterThanOne(dataArr, mapCalc, rowIdx, "Calculated Start", "Calculated Finish") Then
        CalcBridge_RowHasDurationGreaterThanOne = True
        Exit Function
    End If

End Function

Private Function CalcBridge_PairDurationGreaterThanOne( _
    ByRef dataArr As Variant, _
    ByVal mapCalc As Object, _
    ByVal rowIdx As Long, _
    ByVal startColName As String, _
    ByVal finishColName As String) As Boolean

    Dim startVal As Variant
    Dim finishVal As Variant

    If Not mapCalc.Exists(startColName) Then Exit Function
    If Not mapCalc.Exists(finishColName) Then Exit Function

    startVal = dataArr(rowIdx, mapCalc(startColName))
    finishVal = dataArr(rowIdx, mapCalc(finishColName))

    If Not HasValue(startVal) Then Exit Function
    If Not HasValue(finishVal) Then Exit Function

    CalcBridge_PairDurationGreaterThanOne = _
        ((CDbl(finishVal) - CDbl(startVal) + 1) > 1#)

End Function

Private Sub CalcBridge_AddGroupedWarningToCollection( _
    ByVal messages As Collection, _
    ByVal idsDict As Object, _
    ByVal idToWbs As Object, _
    ByVal frProblem As String, _
    ByVal frAction As String, _
    ByVal enProblem As String, _
    ByVal enAction As String, _
    Optional ByVal historyHandled As Boolean = False, _
    Optional ByVal ackTokens As String = "")

    Dim idsLine As String
    Dim wbsLine As String
    Dim msg As String

    If messages Is Nothing Then Exit Sub
    If idsDict Is Nothing Then Exit Sub
    If idsDict.Count = 0 Then Exit Sub

    idsLine = CalcBridge_BuildInlineList(idsDict, 20)
    wbsLine = CalcBridge_BuildInlineWBSList(idsDict, idToWbs, 20)

    msg = _
        "FR:" & vbCrLf & _
        frProblem & vbCrLf & _
        "-> " & frAction & vbCrLf & vbCrLf & _
        "IDs : " & idsLine & vbCrLf & _
        "WBS : " & wbsLine & vbCrLf & vbCrLf & _
        "EN:" & vbCrLf & _
        enProblem & vbCrLf & _
        "-> " & enAction & vbCrLf & vbCrLf & _
        "IDs: " & idsLine & vbCrLf & _
        "WBS: " & wbsLine

    CalcBridge_AddConsoleMessage messages, "WARNING", msg, historyHandled, ackTokens:=ackTokens

End Sub

Private Function CalcBridge_BuildWarningAckTokenList( _
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

Private Function CalcBridge_CreateConsoleItem( _
    ByVal msgType As String, _
    ByVal msgText As String, _
    Optional ByVal historyHandled As Boolean = False, _
    Optional ByVal eventType As String = "", _
    Optional ByVal eventHash As String = "", _
    Optional ByVal ackTokens As String = "") As Object

    Dim item As Object

    Set item = CreateObject("Scripting.Dictionary")
    item("Type") = UCase$(Trim$(msgType))
    item("Message") = CStr(msgText)
    item("HistoryHandled") = historyHandled
    If Trim$(eventType) <> "" Then item("EventType") = Trim$(eventType)
    If Trim$(eventHash) <> "" Then item("Hash") = Trim$(eventHash)
    If Trim$(ackTokens) <> "" Then item("AckTokens") = Trim$(ackTokens)

    Set CalcBridge_CreateConsoleItem = item

End Function

Public Sub CalcBridge_AddConsoleMessage( _
    ByVal targetMessages As Collection, _
    ByVal msgType As String, _
    ByVal msgText As String, _
    Optional ByVal historyHandled As Boolean = False, _
    Optional ByVal eventType As String = "", _
    Optional ByVal eventHash As String = "", _
    Optional ByVal ackTokens As String = "")

    If targetMessages Is Nothing Then Exit Sub
    If Trim$(msgText) = "" Then Exit Sub

    targetMessages.Add CalcBridge_CreateConsoleItem(msgType, msgText, historyHandled, eventType, eventHash, ackTokens)

End Sub

Public Sub CalcBridge_ShowPlanningConsole(ByVal messages As Collection)

    Dim historyMessages As Collection
    Dim displayMessages As Collection
    Dim historyErrorMessage As String

    If messages Is Nothing Then Exit Sub
    If messages.Count = 0 Then Exit Sub

    Set historyMessages = MessageEngine_PrepareConsoleMessages(messages)

    If Not PlanningEvents_LogConsoleMessagesSafe( _
        historyMessages, _
        "CalcBridge_ShowPlanningConsole", _
        historyErrorMessage) Then

        CalcBridge_AddConsoleMessage historyMessages, "STOP", historyErrorMessage, True
    End If

    If ShouldDeferCurrentWorkflowDisplayToRoot("CalcBridge_ShowPlanningConsole") Then
        DeferPlanningWorkflowDisplayMessages historyMessages
        Exit Sub
    End If

    Set displayMessages = MessageEngine_PrepareDisplayMessages(historyMessages)
    If Not MessageEngine_ShouldShowConsole(displayMessages) Then Exit Sub
    If Not CanCurrentWorkflowDisplay("CalcBridge_ShowPlanningConsole") Then Exit Sub

    frmPlanningMessages.LoadMessages displayMessages, "Planning console"
    frmPlanningMessages.Show vbModal

End Sub

Public Function CalcBridge_RecordPlanningMessages( _
    ByVal messages As Collection, _
    Optional ByVal sourceProcedure As String = "CalcBridge_RecordPlanningMessages") As Boolean

    Dim historyMessages As Collection
    Dim historyErrorMessage As String

    If messages Is Nothing Then
        CalcBridge_RecordPlanningMessages = True
        Exit Function
    End If

    If messages.Count = 0 Then
        CalcBridge_RecordPlanningMessages = True
        Exit Function
    End If

    Set historyMessages = MessageEngine_PrepareConsoleMessages(messages)
    If historyMessages.Count = 0 Then
        CalcBridge_RecordPlanningMessages = True
        Exit Function
    End If

    CalcBridge_RecordPlanningMessages = PlanningEvents_LogConsoleMessagesSafe( _
        historyMessages, _
        sourceProcedure, _
        historyErrorMessage)

    If Not CalcBridge_RecordPlanningMessages Then
        Debug.Print "Runtime history logging failed in " & sourceProcedure & ": " & historyErrorMessage
    End If

End Function

Private Sub CalcBridge_AddGroupedStopToCollection( _
    ByVal messages As Collection, _
    ByVal idsDict As Object, _
    ByVal idToWbs As Object, _
    ByVal frProblem As String, _
    ByVal frAction As String, _
    ByVal enProblem As String, _
    ByVal enAction As String)

    If messages Is Nothing Then Exit Sub
    If idsDict Is Nothing Then Exit Sub
    If idsDict.Count = 0 Then Exit Sub

    CalcBridge_AddConsoleMessage messages, "STOP", _
        CalcBridge_BuildGroupedMessage(idsDict, idToWbs, frProblem, frAction, enProblem, enAction)

End Sub

Private Sub CalcBridge_AddUpstreamStopToCollection( _
    ByVal messages As Collection, _
    ByVal idsDict As Object, _
    ByVal idToWbs As Object, _
    ByVal frProblem As String, _
    ByVal frAction As String, _
    ByVal enProblem As String, _
    ByVal enAction As String)

    Dim itemsLine As String
    Dim msg As String

    If messages Is Nothing Then Exit Sub
    If idsDict Is Nothing Then Exit Sub
    If idsDict.Count = 0 Then Exit Sub

    itemsLine = CalcBridge_BuildUpstreamViolationItems(idsDict, idToWbs, 20)

    msg = _
        "FR:" & vbCrLf & _
        frProblem & vbCrLf & _
        "-> " & frAction & vbCrLf & vbCrLf & _
        "Tâches : " & itemsLine & vbCrLf & vbCrLf & _
        "EN:" & vbCrLf & _
        enProblem & vbCrLf & _
        "-> " & enAction & vbCrLf & vbCrLf & _
        "Tasks: " & itemsLine

    CalcBridge_AddConsoleMessage messages, "STOP", msg

End Sub

Public Sub CalcBridge_ShowSingleConsoleMessage( _
    ByVal msgType As String, _
    ByVal frText As String, _
    ByVal enText As String)

    Dim consoleMessages As Collection

    Set consoleMessages = New Collection

    CalcBridge_AddConsoleMessage consoleMessages, msgType, _
        BiMsg(frText, enText)

    CalcBridge_ShowPlanningConsole consoleMessages

End Sub

Public Sub CalcBridge_AddOrShowRawConsoleMessage( _
    ByVal consoleMessages As Collection, _
    ByVal msgType As String, _
    ByVal msgText As String)

    Dim localMessages As Collection

    If Trim$(CStr(msgText)) = "" Then Exit Sub

    If consoleMessages Is Nothing Then
        Set localMessages = New Collection
        CalcBridge_AddConsoleMessage localMessages, msgType, msgText
        CalcBridge_ShowPlanningConsole localMessages
    Else
        CalcBridge_AddConsoleMessage consoleMessages, msgType, msgText
    End If

End Sub

Public Sub CalcBridge_AddOrShowConsoleMessage( _
    ByVal consoleMessages As Collection, _
    ByVal msgType As String, _
    ByVal frText As String, _
    ByVal enText As String)

    CalcBridge_AddOrShowRawConsoleMessage consoleMessages, msgType, _
        BiMsg(frText, enText)

End Sub









