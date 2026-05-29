Attribute VB_Name = "mod_SCurve"
Option Explicit

Private Const SCURVE_SHEET As String = "SCURVE"
Private Const SCURVE_TABLE As String = "tbl_SCURVE"

Private Const CALC_SCURVE_SHEET As String = "CALC_SCURVE"
Private Const CALC_SCURVE_TABLE As String = "tbl_CALC_SCURVE"

Private Const WBS_SHEET As String = "WBS"
Private Const WBS_TABLE As String = "tbl_WBS"

Public Sub Run_SCurve_Engine()

    Dim wsWBS As Worksheet
    Dim wsCalc As Worksheet
    Dim wsSCurve As Worksheet
    Dim wsCalcSCurve As Worksheet

    Dim tblWBS As ListObject
    Dim tblCalc As ListObject
    Dim tblSCurve As ListObject
    Dim tblCalcSCurve As ListObject

    Dim mapWBS As Object
    Dim mapCalc As Object
    Dim mapSCurve As Object
    Dim mapCalcSCurve As Object

    Dim dataWBS As Variant
    Dim dataCalc As Variant
    Dim outCalc() As Variant
    Dim rowCount As Long
    Dim r As Long

    Dim hasChildren As Object
    Dim idToWbs As Object
    Dim calcRowById As Object

    Dim includedIds As Object
    Dim missingWeightIds As Object
    Dim blockingErrorIds As Object
    Dim excludedTaskTypeWithWeightIds As Object

    Dim baselineDailyByDate As Object
    Dim actualizedDailyByDate As Object
    Dim remainingDailyByDate As Object
    Dim allDates As Object

    Dim totalRawWeight As Double
    Dim consoleMessages As Collection
    Dim ackTokens As String

    On Error GoTo SafeExit

    Set consoleMessages = New Collection

    Set wsWBS = ThisWorkbook.Worksheets(WBS_SHEET)
    Set wsCalc = ThisWorkbook.Worksheets("CALC")
    Set wsSCurve = ThisWorkbook.Worksheets(SCURVE_SHEET)
    Set wsCalcSCurve = ThisWorkbook.Worksheets(CALC_SCURVE_SHEET)

    Set tblWBS = wsWBS.ListObjects(WBS_TABLE)
    Set tblCalc = wsCalc.ListObjects("tbl_CALC")
    Set tblSCurve = wsSCurve.ListObjects(SCURVE_TABLE)
    Set tblCalcSCurve = wsCalcSCurve.ListObjects(CALC_SCURVE_TABLE)

    If tblWBS.DataBodyRange Is Nothing Then Exit Sub
    If tblCalc.DataBodyRange Is Nothing Then Exit Sub

    Set mapWBS = BuildColumnMapGeneric(tblWBS)
    Set mapCalc = BuildColumnMapGeneric(tblCalc)
    Set mapSCurve = BuildColumnMapGeneric(tblSCurve)
    Set mapCalcSCurve = BuildColumnMapGeneric(tblCalcSCurve)

    ValidateSCurveSourceColumns mapWBS, mapCalc
    ValidateSCurveOutputColumns mapSCurve
    ValidateCalcSCurveColumns mapCalcSCurve

    dataWBS = tblWBS.DataBodyRange.value
    dataCalc = tblCalc.DataBodyRange.value
    rowCount = UBound(dataWBS, 1)

    Set hasChildren = CreateObject("Scripting.Dictionary")
    Set idToWbs = CreateObject("Scripting.Dictionary")
    Set calcRowById = CreateObject("Scripting.Dictionary")

    Set includedIds = CreateObject("Scripting.Dictionary")
    Set missingWeightIds = CreateObject("Scripting.Dictionary")
    Set blockingErrorIds = CreateObject("Scripting.Dictionary")
    Set excludedTaskTypeWithWeightIds = CreateObject("Scripting.Dictionary")

    Set baselineDailyByDate = CreateObject("Scripting.Dictionary")
    Set actualizedDailyByDate = CreateObject("Scripting.Dictionary")
    Set remainingDailyByDate = CreateObject("Scripting.Dictionary")
    Set allDates = CreateObject("Scripting.Dictionary")

    BuildWBSParentMaps dataCalc, mapCalc, hasChildren, idToWbs
    BuildCalcRowByIdMap dataCalc, mapCalc, calcRowById

    ResizeTableToRowCount tblCalcSCurve, rowCount
    ReDim outCalc(1 To rowCount, 1 To tblCalcSCurve.ListColumns.Count)

    totalRawWeight = 0

    '--------------------------------------------------
    ' PASS 1 : scan tasks / fill debug base / validate
    '--------------------------------------------------
    For r = 1 To rowCount

        Dim taskId As String
        Dim taskWBS As String
        Dim calcRow As Long
        Dim isLeaf As Boolean

        Dim rawWeight As Variant
        Dim baselineStart As Variant
        Dim baselineDuration As Variant
        Dim baselineFinish As Variant

        Dim calcStart As Variant
        Dim calcFinish As Variant
        Dim progressVal As Double
        Dim drivingLogic As String
        Dim taskTypeVal As String

        Dim baselineDaily As Variant
        Dim calculatedDaily As Variant
        Dim actualizedWeight As Variant

        Dim errorText As String

        taskId = Trim$(CStr(dataWBS(r, mapWBS("ID"))))
        baselineStart = Empty
        baselineDuration = Empty
        baselineFinish = Empty
        calcStart = Empty
        calcFinish = Empty
        drivingLogic = ""
        taskTypeVal = ""

        outCalc(r, mapCalcSCurve("ID")) = taskId

        If taskId = "" Then
            outCalc(r, mapCalcSCurve("SCurve Included")) = "NO"
            outCalc(r, mapCalcSCurve("SCurve Warning")) = ""
            outCalc(r, mapCalcSCurve("Error flag")) = ""
            GoTo NextPass1Row
        End If

        If Not calcRowById.Exists(taskId) Then
            blockingErrorIds(taskId) = True
            errorText = "TASK MISSING IN CALC"
            outCalc(r, mapCalcSCurve("SCurve Included")) = "NO"
            outCalc(r, mapCalcSCurve("SCurve Warning")) = errorText
            outCalc(r, mapCalcSCurve("Error flag")) = "ERROR"
            GoTo WritePass1Row
        End If

        calcRow = CLng(calcRowById(taskId))
        taskWBS = Trim$(CStr(dataCalc(calcRow, mapCalc("WBS"))))
        isLeaf = Not hasChildren.Exists(NormalizeWBSLocal(taskWBS))

        rawWeight = dataWBS(r, mapWBS("Weight (%)"))
        progressVal = ProgressToDoubleLocal(dataWBS(r, mapWBS("% Progress")))

        baselineStart = GetCellValue(dataCalc(calcRow, mapCalc("Baseline Start")))
        baselineDuration = GetCellValue(dataCalc(calcRow, mapCalc("Baseline Duration")))
        calcStart = GetCellValue(dataCalc(calcRow, mapCalc("Calculated Start")))
        calcFinish = GetCellValue(dataCalc(calcRow, mapCalc("Calculated Finish")))
        drivingLogic = UCase$(Trim$(CStr(dataCalc(calcRow, mapCalc("Driving Logic")))))
        taskTypeVal = Trim$(CStr(dataCalc(calcRow, mapCalc("Task Type"))))

        baselineDaily = Empty
        calculatedDaily = Empty
        actualizedWeight = Empty
        errorText = ""

        If Not isLeaf Then
            outCalc(r, mapCalcSCurve("SCurve Included")) = "NO"
            outCalc(r, mapCalcSCurve("SCurve Warning")) = "SUMMARY TASK EXCLUDED"
            outCalc(r, mapCalcSCurve("Error flag")) = ""
            GoTo WritePass1Row
        End If

        If IsSCurveExcludedTaskTypeLocal(taskTypeVal) Then
            If HasValue(rawWeight) Then
                excludedTaskTypeWithWeightIds(taskId) = True
            End If

            outCalc(r, mapCalcSCurve("SCurve Included")) = "NO"
            outCalc(r, mapCalcSCurve("SCurve Warning")) = "TASK TYPE EXCLUDED FROM S-CURVE"
            outCalc(r, mapCalcSCurve("Error flag")) = ""
            GoTo WritePass1Row
        End If

        If Not HasValue(rawWeight) Then
            missingWeightIds(taskId) = True
            outCalc(r, mapCalcSCurve("SCurve Included")) = "NO"
            outCalc(r, mapCalcSCurve("SCurve Warning")) = "MISSING WEIGHT - EXCLUDED"
            outCalc(r, mapCalcSCurve("Error flag")) = ""
            GoTo WritePass1Row
        End If

        If Not IsNumeric(rawWeight) Then
            blockingErrorIds(taskId) = True
            errorText = "INVALID WEIGHT"
            outCalc(r, mapCalcSCurve("SCurve Included")) = "NO"
            outCalc(r, mapCalcSCurve("SCurve Warning")) = errorText
            outCalc(r, mapCalcSCurve("Error flag")) = "ERROR"
            GoTo WritePass1Row
        End If

        If CDbl(rawWeight) <= 0 Then
            blockingErrorIds(taskId) = True
            errorText = "NON-POSITIVE WEIGHT"
            outCalc(r, mapCalcSCurve("SCurve Included")) = "NO"
            outCalc(r, mapCalcSCurve("SCurve Warning")) = errorText
            outCalc(r, mapCalcSCurve("Error flag")) = "ERROR"
            GoTo WritePass1Row
        End If

        If (Not HasValue(baselineStart) Or Not HasValue(baselineDuration)) And _
           HasValue(calcStart) And HasValue(calcFinish) And drivingLogic = "BASELINE" Then

            baselineStart = calcStart
            baselineDuration = CDbl(calcFinish) - CDbl(calcStart) + 1
        End If

        If Not HasValue(baselineStart) Or Not HasValue(baselineDuration) Then
            blockingErrorIds(taskId) = True
            errorText = "MISSING BASELINE DATA"
            outCalc(r, mapCalcSCurve("SCurve Included")) = "NO"
            outCalc(r, mapCalcSCurve("SCurve Warning")) = errorText
            outCalc(r, mapCalcSCurve("Error flag")) = "ERROR"
            GoTo WritePass1Row
        End If

        If CDbl(baselineDuration) <= 0 Then
            blockingErrorIds(taskId) = True
            errorText = "INVALID BASELINE DURATION"
            outCalc(r, mapCalcSCurve("SCurve Included")) = "NO"
            outCalc(r, mapCalcSCurve("SCurve Warning")) = errorText
            outCalc(r, mapCalcSCurve("Error flag")) = "ERROR"
            GoTo WritePass1Row
        End If

        If Not HasValue(calcStart) Or Not HasValue(calcFinish) Then
            blockingErrorIds(taskId) = True
            errorText = "MISSING CALCULATED DATES"
            outCalc(r, mapCalcSCurve("SCurve Included")) = "NO"
            outCalc(r, mapCalcSCurve("SCurve Warning")) = errorText
            outCalc(r, mapCalcSCurve("Error flag")) = "ERROR"
            GoTo WritePass1Row
        End If

        If CDbl(calcFinish) < CDbl(calcStart) Then
            blockingErrorIds(taskId) = True
            errorText = "CALCULATED FINISH BEFORE START"
            outCalc(r, mapCalcSCurve("SCurve Included")) = "NO"
            outCalc(r, mapCalcSCurve("SCurve Warning")) = errorText
            outCalc(r, mapCalcSCurve("Error flag")) = "ERROR"
            GoTo WritePass1Row
        End If

        includedIds(taskId) = True
        totalRawWeight = totalRawWeight + CDbl(rawWeight)

        baselineFinish = CDbl(baselineStart) + CDbl(baselineDuration) - 1
        baselineDaily = 1# / CDbl(baselineDuration)
        calculatedDaily = 1# / (CDbl(calcFinish) - CDbl(calcStart) + 1)

        actualizedWeight = CDbl(rawWeight) * progressVal

        outCalc(r, mapCalcSCurve("SCurve Included")) = "YES"
        outCalc(r, mapCalcSCurve("SCurve Warning")) = ""
        outCalc(r, mapCalcSCurve("Error flag")) = ""

WritePass1Row:
        outCalc(r, mapCalcSCurve("SCurve Weight Raw")) = rawWeight
        outCalc(r, mapCalcSCurve("SCurve Weight Normalized")) = ""
        outCalc(r, mapCalcSCurve("SCurve Baseline Start")) = baselineStart
        outCalc(r, mapCalcSCurve("SCurve Baseline Finish")) = baselineFinish
        outCalc(r, mapCalcSCurve("SCurve Calculated Start")) = calcStart
        outCalc(r, mapCalcSCurve("SCurve Calculated Finish")) = calcFinish
        outCalc(r, mapCalcSCurve("SCurve Progress")) = progressVal
        outCalc(r, mapCalcSCurve("SCurve Baseline Daily")) = baselineDaily
        outCalc(r, mapCalcSCurve("SCurve Calculated Daily")) = calculatedDaily
        outCalc(r, mapCalcSCurve("SCurve Actualized Weight")) = actualizedWeight

NextPass1Row:
    Next r

    tblCalcSCurve.DataBodyRange.value = outCalc
    FormatCalcSCurveColumns tblCalcSCurve, mapCalcSCurve

    If blockingErrorIds.Count > 0 Then
        SCurve_AddGroupedMessage consoleMessages, "STOP", blockingErrorIds, idToWbs, _
            "Données bloquantes pour S-curve", "corriger les champs nécessaires avant recalcul", _
            "Blocking data for S-curve", "fix required fields before recalculation"
        GoTo SafeExit
    End If

    If includedIds.Count = 0 Then
        SCurve_AddConsoleMessage consoleMessages, "WARNING", _
            "Aucune tâche feuille exploitable pour la S-curve.", _
            "No valid leaf task available for S-curve.", _
            "SCURVE_NO_VALID_LEAF_TASK"
        GoTo SafeExit
    End If

    If totalRawWeight <= 0 Then
        SCurve_AddConsoleMessage consoleMessages, "STOP", _
            "La somme des poids exploités pour la S-curve est nulle ou invalide.", _
            "The total usable S-curve weight is zero or invalid."
        GoTo SafeExit
    End If

    If missingWeightIds.Count > 0 Then
        ackTokens = SCurve_LogGroupedWarningEvents( _
            missingWeightIds, idToWbs, _
            "SCURVE_MISSING_WEIGHT", _
            "Poids manquant sur certaines taches feuilles", _
            "Missing weight on some leaf tasks", _
            "les taches sont exclues de la S-curve ; completer Weight (%) si necessaire", _
            "tasks are excluded from the S-curve; fill Weight (%) if needed")

        SCurve_AddGroupedMessage consoleMessages, "WARNING", missingWeightIds, idToWbs, _
            "Poids manquant sur certaines tâches feuilles - non prises en compte dans la S-curve", "compléter Weight (%) si nécessaire", _
            "Missing weight on some leaf tasks - excluded from S-curve", "fill Weight (%) if needed", _
            True, _
            ackTokens
    End If

    If excludedTaskTypeWithWeightIds.Count > 0 Then
        ackTokens = SCurve_LogGroupedWarningEvents( _
            excludedTaskTypeWithWeightIds, idToWbs, _
            "SCURVE_EXCLUDED_TASK_TYPE_WITH_WEIGHT", _
            "Poids renseigne sur des Milestones ou Level of Effort", _
            "Weight entered on Milestones or Level of Effort", _
            "les taches sont exclues de la S-curve ; supprimer Weight (%) si vous voulez eviter ce warning", _
            "tasks are excluded from the S-curve; remove Weight (%) if you want to avoid this warning")

        SCurve_AddGroupedMessage consoleMessages, "WARNING", excludedTaskTypeWithWeightIds, idToWbs, _
            "Poids renseigné sur des Milestones ou Level of Effort - non pris en compte dans la S-curve", "supprimer le Weight (%) si vous voulez éviter ce warning", _
            "Weight entered on Milestones or Level of Effort - excluded from S-curve", "remove Weight (%) if you want to avoid this warning", _
            True, _
            ackTokens
    End If

    '--------------------------------------------------
    ' PASS 2 : normalize + build daily curves
    '--------------------------------------------------
    outCalc = tblCalcSCurve.DataBodyRange.value

    For r = 1 To rowCount

        Dim currentId As String
        Dim normWeight As Double
        Dim rowBaselineStart As Variant
        Dim rowBaselineFinish As Variant
        Dim rowCalcStart As Variant
        Dim rowCalcFinish As Variant
        Dim rowProgress As Double
        Dim rowRawWeight As Double
        Dim baselineDays As Long
        Dim calcDays As Long
        Dim dailyCalcUnit As Double
        Dim dailyBaselineUnit As Double
        Dim actualDays As Long
        Dim remainingDays As Long
        Dim daySerial As Long
        Dim i As Long

        currentId = Trim$(CStr(outCalc(r, mapCalcSCurve("ID"))))

        If currentId = "" Then GoTo NextPass2Row
        If Trim$(CStr(outCalc(r, mapCalcSCurve("SCurve Included")))) <> "YES" Then GoTo NextPass2Row

        rowRawWeight = CDbl(outCalc(r, mapCalcSCurve("SCurve Weight Raw")))
        normWeight = rowRawWeight / totalRawWeight
        outCalc(r, mapCalcSCurve("SCurve Weight Normalized")) = normWeight

        rowBaselineStart = outCalc(r, mapCalcSCurve("SCurve Baseline Start"))
        rowBaselineFinish = outCalc(r, mapCalcSCurve("SCurve Baseline Finish"))
        rowCalcStart = outCalc(r, mapCalcSCurve("SCurve Calculated Start"))
        rowCalcFinish = outCalc(r, mapCalcSCurve("SCurve Calculated Finish"))
        rowProgress = ProgressToDoubleLocal(outCalc(r, mapCalcSCurve("SCurve Progress")))

        baselineDays = CLng(CDbl(rowBaselineFinish) - CDbl(rowBaselineStart) + 1)
        calcDays = CLng(CDbl(rowCalcFinish) - CDbl(rowCalcStart) + 1)

        dailyBaselineUnit = normWeight / baselineDays
        dailyCalcUnit = normWeight / calcDays

        actualDays = CLng(Fix(calcDays * rowProgress + 0.0000001))
        If actualDays < 0 Then actualDays = 0
        If actualDays > calcDays Then actualDays = calcDays

        remainingDays = calcDays - actualDays

        outCalc(r, mapCalcSCurve("SCurve Baseline Daily")) = dailyBaselineUnit
        outCalc(r, mapCalcSCurve("SCurve Calculated Daily")) = dailyCalcUnit
        outCalc(r, mapCalcSCurve("SCurve Actualized Weight")) = dailyCalcUnit * actualDays

        For daySerial = CLng(CDbl(rowBaselineStart)) To CLng(CDbl(rowBaselineFinish))
            AddValueByDate baselineDailyByDate, daySerial, dailyBaselineUnit
            allDates(CStr(daySerial)) = True
        Next daySerial

        If actualDays > 0 Then
            For i = 0 To actualDays - 1
                daySerial = CLng(CDbl(rowCalcStart)) + i
                AddValueByDate actualizedDailyByDate, daySerial, dailyCalcUnit
                allDates(CStr(daySerial)) = True
            Next i
        End If

        If remainingDays > 0 Then
            For i = actualDays To calcDays - 1
                daySerial = CLng(CDbl(rowCalcStart)) + i
                AddValueByDate remainingDailyByDate, daySerial, dailyCalcUnit
                allDates(CStr(daySerial)) = True
            Next i
        End If

NextPass2Row:
    Next r

    tblCalcSCurve.DataBodyRange.value = outCalc
    FormatCalcSCurveColumns tblCalcSCurve, mapCalcSCurve

    If allDates.Count = 0 Then
        SCurve_AddConsoleMessage consoleMessages, "STOP", _
            "Aucune date exploitable pour générer la S-curve.", _
            "No usable date found to generate the S-curve."
        GoTo SafeExit
    End If

    WriteSCurveDailyTable tblSCurve, mapSCurve, allDates, baselineDailyByDate, actualizedDailyByDate, remainingDailyByDate
    Ensure_SCurve_Chart

SafeExit:
    If Err.Number <> 0 Then
        If consoleMessages Is Nothing Then Set consoleMessages = New Collection

        SCurve_AddConsoleMessage consoleMessages, "STOP", _
            "Erreur VBA dans Run_SCurve_Engine" & vbCrLf & _
            "-> vérifier le dernier bloc modifié dans mod_SCurve" & vbCrLf & _
            "-> " & Err.Description, _
            "VBA error in Run_SCurve_Engine" & vbCrLf & _
            "-> check the last edited block in mod_SCurve" & vbCrLf & _
            "-> " & Err.Description
    End If

    If Not consoleMessages Is Nothing Then
        If consoleMessages.Count > 0 Then
            CalcBridge_ShowPlanningConsole consoleMessages
        End If
    End If

End Sub

Private Sub ValidateSCurveSourceColumns(ByVal mapWBS As Object, ByVal mapCalc As Object)

    Dim requiredWbsCols As Variant
    Dim requiredCalcCols As Variant
    Dim c As Variant

    requiredWbsCols = Array( _
        "ID", _
        "Weight (%)", _
        "% Progress" _
    )

    requiredCalcCols = Array( _
        "ID", _
        "WBS", _
        "Task Type", _
        "Baseline Start", _
        "Baseline Duration", _
        "Calculated Start", _
        "Calculated Finish", _
        "Driving Logic" _
    )

    For Each c In requiredWbsCols
        If Not mapWBS.Exists(CStr(c)) Then
            Err.Raise vbObjectError + 810, , "Missing source column in tbl_WBS: " & CStr(c)
        End If
    Next c

    For Each c In requiredCalcCols
        If Not mapCalc.Exists(CStr(c)) Then
            Err.Raise vbObjectError + 813, , "Missing source column in tbl_CALC: " & CStr(c)
        End If
    Next c

End Sub

Private Sub ValidateSCurveOutputColumns(ByVal mapSCurve As Object)

    Dim requiredCols As Variant
    Dim c As Variant

    requiredCols = Array( _
        "Date", _
        "Daily Baseline", _
        "Cumulative Baseline", _
        "Daily Actualized", _
        "Cumulative Actualized", _
        "Daily Remaining Forecast", _
        "Cumulative Remaining Forecast", _
        "Calculated Curve Solid", _
        "Calculated Curve Dashed", _
        "Cumulative Actual" _
    )

    For Each c In requiredCols
        If Not mapSCurve.Exists(CStr(c)) Then
            Err.Raise vbObjectError + 811, , "Missing column in tbl_SCURVE: " & CStr(c)
        End If
    Next c

End Sub

Private Sub ValidateCalcSCurveColumns(ByVal mapCalcSCurve As Object)

    Dim requiredCols As Variant
    Dim c As Variant

    requiredCols = Array( _
        "ID", _
        "SCurve Included", _
        "SCurve Warning", _
        "SCurve Weight Raw", _
        "SCurve Weight Normalized", _
        "SCurve Baseline Start", _
        "SCurve Baseline Finish", _
        "SCurve Calculated Start", _
        "SCurve Calculated Finish", _
        "SCurve Progress", _
        "SCurve Baseline Daily", _
        "SCurve Calculated Daily", _
        "SCurve Actualized Weight", _
        "Error flag" _
    )

    For Each c In requiredCols
        If Not mapCalcSCurve.Exists(CStr(c)) Then
            Err.Raise vbObjectError + 812, , "Missing column in tbl_CALC_SCURVE: " & CStr(c)
        End If
    Next c

End Sub

Private Function BuildColumnMapGeneric(ByVal tbl As ListObject) As Object

    Dim d As Object
    Dim i As Long

    Set d = CreateObject("Scripting.Dictionary")

    For i = 1 To tbl.ListColumns.Count
        d(tbl.ListColumns(i).Name) = i
    Next i

    Set BuildColumnMapGeneric = d

End Function

Private Sub BuildWBSParentMaps( _
    ByRef dataWBS As Variant, _
    ByVal mapWBS As Object, _
    ByVal hasChildren As Object, _
    ByVal idToWbs As Object)

    Dim r As Long
    Dim rowCount As Long
    Dim currentWBS As String
    Dim parentWbs As String
    Dim taskId As String

    rowCount = UBound(dataWBS, 1)

    For r = 1 To rowCount
        currentWBS = NormalizeWBSLocal(CStr(dataWBS(r, mapWBS("WBS"))))
        taskId = Trim$(CStr(dataWBS(r, mapWBS("ID"))))

        If currentWBS <> "" Then
            parentWbs = GetParentWBS(currentWBS)
            If parentWbs <> "" Then
                hasChildren(parentWbs) = True
            End If
        End If

        If taskId <> "" Then
            idToWbs(taskId) = currentWBS
        End If
    Next r

End Sub

Private Sub ResizeTableToRowCount(ByVal tbl As ListObject, ByVal targetRows As Long)

    Dim currentRows As Long
    Dim r As Long

    If tbl.DataBodyRange Is Nothing Then
        currentRows = 0
    Else
        currentRows = tbl.ListRows.Count
    End If

    If targetRows = 0 Then
        Do While tbl.ListRows.Count > 0
            tbl.ListRows(tbl.ListRows.Count).Delete
        Loop
        Exit Sub
    End If

    If currentRows < targetRows Then
        For r = currentRows + 1 To targetRows
            tbl.ListRows.Add
        Next r
    ElseIf currentRows > targetRows Then
        For r = currentRows To targetRows + 1 Step -1
            tbl.ListRows(r).Delete
        Next r
    End If

End Sub

Private Sub AddValueByDate(ByVal dict As Object, ByVal dayKey As Long, ByVal valueToAdd As Double)

    If dict.Exists(CStr(dayKey)) Then
        dict(CStr(dayKey)) = CDbl(dict(CStr(dayKey))) + valueToAdd
    Else
        dict(CStr(dayKey)) = valueToAdd
    End If

End Sub

Private Sub WriteSCurveDailyTable( _
    ByVal tblSCurve As ListObject, _
    ByVal mapSCurve As Object, _
    ByVal allDates As Object, _
    ByVal baselineDailyByDate As Object, _
    ByVal actualizedDailyByDate As Object, _
    ByVal remainingDailyByDate As Object)

    Dim sortedDates() As Long
    Dim idx As Long
    Dim key As Variant
    Dim outArr() As Variant

    Dim cumulativeBaseline As Double
    Dim cumulativeActualized As Double
    Dim cumulativeRemainingForecast As Double

    Dim currentDateKey As String
    Dim dailyBaseline As Double
    Dim dailyActualized As Double
    Dim dailyRemainingForecast As Double

    Dim todaySerial As Long
    Dim currentDate As Long

    ReDim sortedDates(1 To allDates.Count)

    idx = 0
    For Each key In allDates.Keys
        idx = idx + 1
        sortedDates(idx) = CLng(key)
    Next key

    SortLongArray sortedDates

    ResizeTableToRowCount tblSCurve, UBound(sortedDates)
    ReDim outArr(1 To UBound(sortedDates), 1 To tblSCurve.ListColumns.Count)

    cumulativeBaseline = 0
    cumulativeActualized = 0
    cumulativeRemainingForecast = 0
    todaySerial = CLng(Date)

    For idx = 1 To UBound(sortedDates)

        currentDateKey = CStr(sortedDates(idx))
        currentDate = sortedDates(idx)

        If baselineDailyByDate.Exists(currentDateKey) Then
            dailyBaseline = CDbl(baselineDailyByDate(currentDateKey))
        Else
            dailyBaseline = 0
        End If

        If actualizedDailyByDate.Exists(currentDateKey) Then
            dailyActualized = CDbl(actualizedDailyByDate(currentDateKey))
        Else
            dailyActualized = 0
        End If

        If remainingDailyByDate.Exists(currentDateKey) Then
            dailyRemainingForecast = CDbl(remainingDailyByDate(currentDateKey))
        Else
            dailyRemainingForecast = 0
        End If

        cumulativeBaseline = cumulativeBaseline + dailyBaseline
        cumulativeActualized = cumulativeActualized + dailyActualized
        cumulativeRemainingForecast = cumulativeRemainingForecast + dailyRemainingForecast

        outArr(idx, mapSCurve("Date")) = CLng(sortedDates(idx))
        outArr(idx, mapSCurve("Daily Baseline")) = dailyBaseline
        outArr(idx, mapSCurve("Cumulative Baseline")) = cumulativeBaseline
        outArr(idx, mapSCurve("Daily Actualized")) = dailyActualized
        outArr(idx, mapSCurve("Cumulative Actualized")) = cumulativeActualized
        outArr(idx, mapSCurve("Daily Remaining Forecast")) = dailyRemainingForecast
        outArr(idx, mapSCurve("Cumulative Remaining Forecast")) = cumulativeRemainingForecast
        outArr(idx, mapSCurve("Calculated Curve Solid")) = CVErr(xlErrNA)
        outArr(idx, mapSCurve("Calculated Curve Dashed")) = CVErr(xlErrNA)
        outArr(idx, mapSCurve("Cumulative Actual")) = CVErr(xlErrNA)

        If currentDate <= todaySerial Then
            outArr(idx, mapSCurve("Cumulative Actual")) = cumulativeActualized
        End If

    Next idx

    For idx = 1 To UBound(sortedDates)

        currentDate = sortedDates(idx)

        If currentDate < todaySerial Then
            outArr(idx, mapSCurve("Calculated Curve Solid")) = _
                CDbl(outArr(idx, mapSCurve("Cumulative Actualized"))) + CDbl(outArr(idx, mapSCurve("Cumulative Remaining Forecast")))
        Else
            outArr(idx, mapSCurve("Calculated Curve Solid")) = CVErr(xlErrNA)
        End If

        If currentDate >= todaySerial Then
            outArr(idx, mapSCurve("Calculated Curve Dashed")) = _
                CDbl(outArr(idx, mapSCurve("Cumulative Actualized"))) + CDbl(outArr(idx, mapSCurve("Cumulative Remaining Forecast")))
        Else
            outArr(idx, mapSCurve("Calculated Curve Dashed")) = CVErr(xlErrNA)
        End If

    Next idx

    tblSCurve.DataBodyRange.value = outArr

    tblSCurve.ListColumns("Date").DataBodyRange.NumberFormat = "dd/mm/yyyy"
    tblSCurve.ListColumns("Daily Baseline").DataBodyRange.NumberFormat = "0.00%"
    tblSCurve.ListColumns("Cumulative Baseline").DataBodyRange.NumberFormat = "0.00%"
    tblSCurve.ListColumns("Daily Actualized").DataBodyRange.NumberFormat = "0.00%"
    tblSCurve.ListColumns("Cumulative Actualized").DataBodyRange.NumberFormat = "0.00%"
    tblSCurve.ListColumns("Daily Remaining Forecast").DataBodyRange.NumberFormat = "0.00%"
    tblSCurve.ListColumns("Cumulative Remaining Forecast").DataBodyRange.NumberFormat = "0.00%"
    tblSCurve.ListColumns("Calculated Curve Solid").DataBodyRange.NumberFormat = "0.00%"
    tblSCurve.ListColumns("Calculated Curve Dashed").DataBodyRange.NumberFormat = "0.00%"
    tblSCurve.ListColumns("Cumulative Actual").DataBodyRange.NumberFormat = "0.00%"

End Sub

Private Sub SortLongArray(ByRef arr() As Long)

    Dim i As Long
    Dim j As Long
    Dim tempVal As Long

    For i = LBound(arr) To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If arr(j) < arr(i) Then
                tempVal = arr(i)
                arr(i) = arr(j)
                arr(j) = tempVal
            End If
        Next j
    Next i

End Sub

Private Sub FormatCalcSCurveColumns(ByVal tblCalcSCurve As ListObject, ByVal mapCalcSCurve As Object)

    If tblCalcSCurve.DataBodyRange Is Nothing Then Exit Sub

    tblCalcSCurve.ListColumns("SCurve Baseline Start").DataBodyRange.NumberFormat = "dd/mm/yyyy"
    tblCalcSCurve.ListColumns("SCurve Baseline Finish").DataBodyRange.NumberFormat = "dd/mm/yyyy"
    tblCalcSCurve.ListColumns("SCurve Calculated Start").DataBodyRange.NumberFormat = "dd/mm/yyyy"
    tblCalcSCurve.ListColumns("SCurve Calculated Finish").DataBodyRange.NumberFormat = "dd/mm/yyyy"

    tblCalcSCurve.ListColumns("SCurve Weight Raw").DataBodyRange.NumberFormat = "0.00%"
    tblCalcSCurve.ListColumns("SCurve Weight Normalized").DataBodyRange.NumberFormat = "0.00%"
    tblCalcSCurve.ListColumns("SCurve Progress").DataBodyRange.NumberFormat = "0.00%"
    tblCalcSCurve.ListColumns("SCurve Baseline Daily").DataBodyRange.NumberFormat = "0.0000%"
    tblCalcSCurve.ListColumns("SCurve Calculated Daily").DataBodyRange.NumberFormat = "0.0000%"
    tblCalcSCurve.ListColumns("SCurve Actualized Weight").DataBodyRange.NumberFormat = "0.00%"

End Sub

Private Function NormalizeWBSLocal(ByVal wbs As String) As String

    NormalizeWBSLocal = Replace(Trim$(wbs), ",", ".")

End Function

Private Function ProgressToDoubleLocal(ByVal v As Variant) As Double

    If IsNumeric(v) Then
        ProgressToDoubleLocal = CDbl(v)
    Else
        ProgressToDoubleLocal = 0
    End If

    If ProgressToDoubleLocal < 0 Then ProgressToDoubleLocal = 0
    If ProgressToDoubleLocal > 1 Then ProgressToDoubleLocal = 1

End Function

Private Sub ShowSCurveGroupedMessage( _
    ByVal idsDict As Object, _
    ByVal idToWbs As Object, _
    ByVal frProblem As String, _
    ByVal frAction As String, _
    ByVal enProblem As String, _
    ByVal enAction As String, _
    ByVal boxStyle As VbMsgBoxStyle)

    Dim consoleMessages As Collection
    Dim ackTokens As String
    Dim msgType As String

    If idsDict Is Nothing Then Exit Sub
    If idsDict.Count = 0 Then Exit Sub

    If (boxStyle And vbCritical) = vbCritical Then
        msgType = "STOP"
    ElseIf (boxStyle And vbExclamation) = vbExclamation Then
        msgType = "WARNING"
    Else
        msgType = "INFO"
    End If

    Set consoleMessages = New Collection

    SCurve_AddGroupedMessage consoleMessages, msgType, idsDict, idToWbs, _
        frProblem, frAction, enProblem, enAction

    CalcBridge_ShowPlanningConsole consoleMessages

End Sub

Private Function BuildSCurveGroupedMessage( _
    ByVal idsDict As Object, _
    ByVal idToWbs As Object, _
    ByVal frProblem As String, _
    ByVal frAction As String, _
    ByVal enProblem As String, _
    ByVal enAction As String) As String

    Dim idsLine As String
    Dim wbsLine As String

    idsLine = BuildInlineListLocal(idsDict, 20)
    wbsLine = BuildInlineWBSListLocal(idsDict, idToWbs, 20)

    BuildSCurveGroupedMessage = _
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

Private Function BuildInlineListLocal(ByVal idsDict As Object, ByVal maxItems As Long) As String

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

    BuildInlineListLocal = result

End Function

Private Function BuildInlineWBSListLocal(ByVal idsDict As Object, ByVal idToWbs As Object, ByVal maxItems As Long) As String

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

    BuildInlineWBSListLocal = result

End Function

Private Sub Ensure_SCurve_Chart()

    Dim ws As Worksheet
    Dim chartObj As ChartObject

    Set ws = ThisWorkbook.Worksheets(SCURVE_SHEET)

    On Error Resume Next
    Set chartObj = ws.ChartObjects("cht_SCurve")
    On Error GoTo 0

    If chartObj Is Nothing Then
        Create_SCurve_Chart
    Else
        Update_SCurve_Chart chartObj.Chart
    End If

End Sub

Private Sub Create_SCurve_Chart()

    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim chartObj As ChartObject
    Dim ch As Chart
    Dim s As Series

    Set ws = ThisWorkbook.Worksheets(SCURVE_SHEET)
    Set tbl = ws.ListObjects(SCURVE_TABLE)

    If tbl.DataBodyRange Is Nothing Then Exit Sub

    Set chartObj = ws.ChartObjects.Add( _
        Left:=ws.Range("L2").Left, _
        Top:=ws.Range("L2").Top, _
        Width:=1100, _
        Height:=500)

    chartObj.Name = "cht_SCurve"
    Set ch = chartObj.Chart

    Do While ch.SeriesCollection.Count > 0
        ch.SeriesCollection(1).Delete
    Loop

    ch.HasTitle = True
    ch.ChartTitle.Text = "S-Curve"
    ch.HasLegend = True

    Set s = ch.SeriesCollection.NewSeries
    With s
        .Name = "Daily Actualized"
        .XValues = tbl.ListColumns("Date").DataBodyRange
        .Values = tbl.ListColumns("Daily Actualized").DataBodyRange
        .AxisGroup = xlPrimary
        .ChartType = xlColumnStacked
        .Format.Fill.ForeColor.RGB = RGB(31, 78, 121)
        .Format.Line.Visible = msoFalse
    End With

    Set s = ch.SeriesCollection.NewSeries
    With s
        .Name = "Daily Remaining Forecast"
        .XValues = tbl.ListColumns("Date").DataBodyRange
        .Values = tbl.ListColumns("Daily Remaining Forecast").DataBodyRange
        .AxisGroup = xlPrimary
        .ChartType = xlColumnStacked
        .Format.Fill.ForeColor.RGB = RGB(191, 191, 191)
        .Format.Line.Visible = msoFalse
    End With

    Set s = ch.SeriesCollection.NewSeries
    With s
        .Name = "Baseline"
        .XValues = tbl.ListColumns("Date").DataBodyRange
        .Values = tbl.ListColumns("Cumulative Baseline").DataBodyRange
        .AxisGroup = xlSecondary
        .ChartType = xlLine
        .Format.Line.ForeColor.RGB = RGB(166, 166, 166)
        .Format.Line.Weight = 1.5
        .MarkerStyle = xlMarkerStyleNone
    End With

    Set s = ch.SeriesCollection.NewSeries
    With s
        .Name = "Calculated"
        .XValues = tbl.ListColumns("Date").DataBodyRange
        .Values = tbl.ListColumns("Calculated Curve Solid").DataBodyRange
        .AxisGroup = xlSecondary
        .ChartType = xlLine
        .Format.Line.ForeColor.RGB = RGB(68, 114, 196)
        .Format.Line.Weight = 2.75
        .MarkerStyle = xlMarkerStyleNone
    End With

    Set s = ch.SeriesCollection.NewSeries
    With s
        .Name = "Forecast"
        .XValues = tbl.ListColumns("Date").DataBodyRange
        .Values = tbl.ListColumns("Calculated Curve Dashed").DataBodyRange
        .AxisGroup = xlSecondary
        .ChartType = xlLine
        .Format.Line.ForeColor.RGB = RGB(157, 195, 230)
        .Format.Line.Weight = 2.75
        .Format.Line.DashStyle = msoLineDash
        .MarkerStyle = xlMarkerStyleNone
    End With

    Set s = ch.SeriesCollection.NewSeries
    With s
        .Name = "Actual"
        .XValues = tbl.ListColumns("Date").DataBodyRange
        .Values = tbl.ListColumns("Cumulative Actual").DataBodyRange
        .AxisGroup = xlSecondary
        .ChartType = xlLine
        .Format.Line.ForeColor.RGB = RGB(31, 78, 121)
        .Format.Line.Weight = 2.75
        .MarkerStyle = xlMarkerStyleNone
    End With

    Format_SCurve_Chart ch
    DrawTodayVerticalLine ch

End Sub

Private Sub Update_SCurve_Chart(ByVal ch As Chart)

    Dim tbl As ListObject
    Dim rebuildNeeded As Boolean

    Set tbl = ThisWorkbook.Worksheets(SCURVE_SHEET).ListObjects(SCURVE_TABLE)

    If tbl.DataBodyRange Is Nothing Then Exit Sub

    rebuildNeeded = False

    If ch.SeriesCollection.Count <> 6 Then
        rebuildNeeded = True
    Else
        On Error Resume Next
        rebuildNeeded = ( _
            ch.SeriesCollection(1).ChartType <> xlColumnStacked Or _
            ch.SeriesCollection(2).ChartType <> xlColumnStacked Or _
            ch.SeriesCollection(1).AxisGroup <> xlPrimary Or _
            ch.SeriesCollection(2).AxisGroup <> xlPrimary Or _
            ch.SeriesCollection(3).ChartType <> xlLine Or _
            ch.SeriesCollection(3).AxisGroup <> xlSecondary)
        On Error GoTo 0
    End If

    If rebuildNeeded Then
        ch.Parent.Delete
        Create_SCurve_Chart
        Exit Sub
    End If

    ch.SeriesCollection(1).XValues = tbl.ListColumns("Date").DataBodyRange
    ch.SeriesCollection(1).Values = tbl.ListColumns("Daily Actualized").DataBodyRange

    ch.SeriesCollection(2).XValues = tbl.ListColumns("Date").DataBodyRange
    ch.SeriesCollection(2).Values = tbl.ListColumns("Daily Remaining Forecast").DataBodyRange

    ch.SeriesCollection(3).XValues = tbl.ListColumns("Date").DataBodyRange
    ch.SeriesCollection(3).Values = tbl.ListColumns("Cumulative Baseline").DataBodyRange

    ch.SeriesCollection(4).XValues = tbl.ListColumns("Date").DataBodyRange
    ch.SeriesCollection(4).Values = tbl.ListColumns("Calculated Curve Solid").DataBodyRange

    ch.SeriesCollection(5).XValues = tbl.ListColumns("Date").DataBodyRange
    ch.SeriesCollection(5).Values = tbl.ListColumns("Calculated Curve Dashed").DataBodyRange

    ch.SeriesCollection(6).XValues = tbl.ListColumns("Date").DataBodyRange
    ch.SeriesCollection(6).Values = tbl.ListColumns("Cumulative Actual").DataBodyRange

    ch.SeriesCollection(1).AxisGroup = xlPrimary
    ch.SeriesCollection(1).ChartType = xlColumnStacked

    ch.SeriesCollection(2).AxisGroup = xlPrimary
    ch.SeriesCollection(2).ChartType = xlColumnStacked

    ch.SeriesCollection(3).AxisGroup = xlSecondary
    ch.SeriesCollection(3).ChartType = xlLine

    ch.SeriesCollection(4).AxisGroup = xlSecondary
    ch.SeriesCollection(4).ChartType = xlLine

    ch.SeriesCollection(5).AxisGroup = xlSecondary
    ch.SeriesCollection(5).ChartType = xlLine

    ch.SeriesCollection(6).AxisGroup = xlSecondary
    ch.SeriesCollection(6).ChartType = xlLine

    Format_SCurve_Chart ch
    DrawTodayVerticalLine ch

End Sub


Private Sub Format_SCurve_Chart(ByVal ch As Chart)

    With ch.Axes(xlValue, xlPrimary)
        .MinimumScale = 0
        .MaximumScale = 0.12
        .MajorUnit = 0.02
        .TickLabels.NumberFormat = "0%"
        .HasMajorGridlines = False
    End With

    With ch.Axes(xlValue, xlSecondary)
        .MinimumScale = 0
        .MaximumScale = 1
        .MajorUnit = 0.1
        .TickLabels.NumberFormat = "0%"
        .HasMajorGridlines = True
    End With

    With ch.Axes(xlCategory)
        .TickLabels.NumberFormat = "dd/mm/yyyy"
        .TickLabels.Orientation = 45
        .CategoryType = xlCategoryScale
    End With

    ch.Legend.Position = xlLegendPositionRight

    On Error Resume Next
    ch.ChartGroups(1).GapWidth = 35
    ch.ChartGroups(1).Overlap = 100
    On Error GoTo 0

End Sub

Private Sub DrawTodayVerticalLine(ByVal ch As Chart)

    Dim ws As Worksheet
    Dim shp As Shape
    Dim todaySerial As Double
    Dim minDate As Double
    Dim maxDate As Double
    Dim plotLeft As Double
    Dim plotTop As Double
    Dim plotWidth As Double
    Dim plotHeight As Double
    Dim xPos As Double
    Dim ratio As Double
    Dim chartObj As ChartObject
    Dim dateRange As Range
    Dim dateCount As Long
    Dim categoryIndex As Double
    Dim r As Long
    Dim currentDate As Double
    Dim nextDate As Double
    Dim axisBetweenCategories As Boolean

    Set ws = ThisWorkbook.Worksheets(SCURVE_SHEET)
    Set chartObj = ch.Parent
    Set dateRange = ws.ListObjects(SCURVE_TABLE).ListColumns("Date").DataBodyRange

    todaySerial = CDbl(Date)
    dateCount = dateRange.Rows.Count

    minDate = CDbl(dateRange.Cells(1, 1).value)
    maxDate = CDbl(dateRange.Cells(dateCount, 1).value)

    On Error Resume Next
    ws.Shapes("SCURVE_TODAY_LINE").Delete
    On Error GoTo 0

    If todaySerial < minDate Or todaySerial > maxDate Then Exit Sub
    If maxDate <= minDate Then Exit Sub
    If dateCount < 2 Then Exit Sub

    categoryIndex = 0#
    For r = 1 To dateCount
        currentDate = CDbl(dateRange.Cells(r, 1).value)

        If currentDate = todaySerial Then
            categoryIndex = CDbl(r)
            Exit For
        End If

        If r < dateCount Then
            nextDate = CDbl(dateRange.Cells(r + 1, 1).value)
            If todaySerial > currentDate And todaySerial < nextDate Then
                categoryIndex = CDbl(r) + ((todaySerial - currentDate) / (nextDate - currentDate))
                Exit For
            End If
        End If
    Next r

    If categoryIndex = 0# Then Exit Sub

    plotLeft = chartObj.Left + ch.PlotArea.InsideLeft
    plotTop = chartObj.Top + ch.PlotArea.InsideTop
    plotWidth = ch.PlotArea.InsideWidth
    plotHeight = ch.PlotArea.InsideHeight

    axisBetweenCategories = False
    On Error Resume Next
    axisBetweenCategories = ch.Axes(xlCategory).axisBetweenCategories
    On Error GoTo 0

    If axisBetweenCategories Then
        ratio = (categoryIndex - 0.5) / CDbl(dateCount)
    Else
        ratio = (categoryIndex - 1#) / CDbl(dateCount - 1)
    End If

    xPos = plotLeft + (plotWidth * ratio)

    Set shp = ws.Shapes.AddLine(xPos, plotTop, xPos, plotTop + plotHeight)
    shp.Name = "SCURVE_TODAY_LINE"

    With shp.Line
        .ForeColor.RGB = RGB(0, 176, 80)
        .Weight = 1.5
        .DashStyle = msoLineDash
    End With

End Sub

Private Sub BuildCalcRowByIdMap(ByRef dataCalc As Variant, ByVal mapCalc As Object, ByVal calcRowById As Object)

    Dim r As Long
    Dim rowCount As Long
    Dim taskId As String

    rowCount = UBound(dataCalc, 1)

    For r = 1 To rowCount
        taskId = Trim$(CStr(dataCalc(r, mapCalc("ID"))))
        If taskId <> "" Then
            calcRowById(taskId) = r
        End If
    Next r

End Sub

Private Sub SCurve_AddConsoleMessage( _
    ByVal consoleMessages As Collection, _
    ByVal msgType As String, _
    ByVal frText As String, _
    ByVal enText As String, _
    Optional ByVal eventType As String = "", _
    Optional ByVal eventHash As String = "")

    If consoleMessages Is Nothing Then Exit Sub

    CalcBridge_AddConsoleMessage consoleMessages, msgType, _
        "FR:" & vbCrLf & _
        frText & vbCrLf & vbCrLf & _
        "EN:" & vbCrLf & _
        enText, _
        False, _
        eventType, _
        eventHash

End Sub

Private Sub SCurve_AddGroupedMessage( _
    ByVal consoleMessages As Collection, _
    ByVal msgType As String, _
    ByVal idsDict As Object, _
    ByVal idToWbs As Object, _
    ByVal frProblem As String, _
    ByVal frAction As String, _
    ByVal enProblem As String, _
    ByVal enAction As String, _
    Optional ByVal historyHandled As Boolean = False, _
    Optional ByVal ackTokens As String = "")

    If consoleMessages Is Nothing Then Exit Sub
    If idsDict Is Nothing Then Exit Sub
    If idsDict.Count = 0 Then Exit Sub

    CalcBridge_AddConsoleMessage consoleMessages, msgType, _
        BuildSCurveGroupedMessage(idsDict, idToWbs, frProblem, frAction, enProblem, enAction), _
        historyHandled, _
        ackTokens:=ackTokens

End Sub

Private Function SCurve_LogGroupedWarningEvents( _
    ByVal idsDict As Object, _
    ByVal idToWbs As Object, _
    ByVal eventType As String, _
    ByVal frMessage As String, _
    ByVal enMessage As String, _
    ByVal frDetails As String, _
    ByVal enDetails As String) As String

    Dim key As Variant
    Dim idVal As String
    Dim wbsVal As String
    Dim eventHash As String
    Dim tokens As String

    If idsDict Is Nothing Then Exit Function

    For Each key In idsDict.Keys
        idVal = Trim$(CStr(key))
        If idVal <> "" Then
            wbsVal = vbNullString
            If Not idToWbs Is Nothing Then
                If idToWbs.Exists(idVal) Then wbsVal = CStr(idToWbs(idVal))
            End If

            eventHash = BuildPlanningEventHash( _
                "WARNING", eventType, frMessage, enMessage, frDetails, enDetails, _
                "Run_SCurve_Engine", _
                "SCURVE", _
                "tbl_SCURVE", _
                idVal, _
                wbsVal, _
                vbNullString)

            LogPlanningEvent _
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
                False

            If tokens <> "" Then tokens = tokens & ";"
            tokens = tokens & BuildPlanningWarningAckToken(eventType, eventHash)
        End If
    Next key

    SCurve_LogGroupedWarningEvents = tokens

End Function
Private Function IsSCurveExcludedTaskTypeLocal(ByVal taskTypeValue As Variant) As Boolean

    Dim s As String

    s = UCase$(Trim$(CStr(taskTypeValue)))
    s = Replace$(s, "-", " ")
    s = Replace$(s, "_", " ")

    Do While InStr(1, s, "  ", vbBinaryCompare) > 0
        s = Replace$(s, "  ", " ")
    Loop

    IsSCurveExcludedTaskTypeLocal = _
        (s = "MILESTONE") Or _
        (s = "MS") Or _
        (s = "JALON") Or _
        (s = "LEVEL OF EFFORT") Or _
        (s = "LOE")

End Function


