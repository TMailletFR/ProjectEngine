Attribute VB_Name = "mod_CalcIncremental"
Option Explicit

Private Const CALC_SHEET_NAME_INC As String = "CALC"
Private Const CALC_TABLE_NAME_INC As String = "tbl_CALC"
Private Const CALC_STATE_SHEET_NAME_INC As String = "CALC_STATE"
Private Const CALC_STATE_TABLE_NAME_INC As String = "tbl_CALC_STATE"

Public Function Get_Changed_TaskIds(ByRef forceFullRecalc As Boolean) As Object

    Dim wsCalc As Worksheet
    Dim wsState As Worksheet
    Dim tblCalc As ListObject
    Dim tblState As ListObject

    Dim mapCalc As Object
    Dim mapState As Object

    Dim arrCalc As Variant
    Dim arrState As Variant

    Dim stateSigById As Object
    Dim stateRunStatus As String
    Dim changedIds As Object

    Dim r As Long
    Dim idVal As String
    Dim currentSig As String

    Set changedIds = CreateObject("Scripting.Dictionary")
    forceFullRecalc = False

    On Error GoTo ForceFull

    Set wsCalc = ThisWorkbook.Worksheets(CALC_SHEET_NAME_INC)
    Set tblCalc = wsCalc.ListObjects(CALC_TABLE_NAME_INC)

    On Error Resume Next
    Set wsState = ThisWorkbook.Worksheets(CALC_STATE_SHEET_NAME_INC)
    If Not wsState Is Nothing Then
        Set tblState = wsState.ListObjects(CALC_STATE_TABLE_NAME_INC)
    End If
    On Error GoTo ForceFull

    If tblCalc Is Nothing Then GoTo ForceFull
    If tblCalc.DataBodyRange Is Nothing Then
        Debug.Print "INCREMENTAL CHECK - tbl_CALC empty: no changed task"
        Set Get_Changed_TaskIds = changedIds
        Exit Function
    End If

    If tblState Is Nothing Then GoTo ForceFull
    If tblState.DataBodyRange Is Nothing Then GoTo ForceFull

    Set mapCalc = BuildColumnMap_Incremental(tblCalc)
    Set mapState = BuildColumnMap_Incremental(tblState)

    ValidateIncrementalCalcColumns mapCalc
    ValidateIncrementalStateColumns mapState

    arrCalc = tblCalc.DataBodyRange.value
    arrState = tblState.DataBodyRange.value

    Set stateSigById = BuildStateSignatureMap(arrState, mapState)
    stateRunStatus = GetStateRunStatus(arrState, mapState)

    If UCase$(Trim$(stateRunStatus)) <> "OK" Then GoTo ForceFull

    For r = 1 To UBound(arrCalc, 1)

        idVal = Trim$(CStr(arrCalc(r, mapCalc("ID"))))
        If idVal = "" Then GoTo NextRow

        currentSig = BuildCalcCurrentRowSignature(arrCalc, r, mapCalc)

        If Not stateSigById.Exists(idVal) Then
            changedIds(idVal) = True
            GoTo NextRow
        End If

        If CStr(stateSigById(idVal)) <> currentSig Then
            changedIds(idVal) = True
        End If

NextRow:
    Next r

    Debug.Print "INCREMENTAL CHECK - changed tasks count: " & changedIds.Count

    Set Get_Changed_TaskIds = changedIds
    Exit Function

ForceFull:
    Debug.Print "INCREMENTAL CHECK - FULL FORCED: " & Err.Description

    forceFullRecalc = True
    Set changedIds = CreateObject("Scripting.Dictionary")
    Set Get_Changed_TaskIds = changedIds

End Function



Private Function BuildColumnMap_Incremental(ByVal tbl As ListObject) As Object

    Dim d As Object
    Dim i As Long

    Set d = CreateObject("Scripting.Dictionary")

    For i = 1 To tbl.ListColumns.Count
        d(tbl.ListColumns(i).Name) = i
    Next i

    Set BuildColumnMap_Incremental = d

End Function

Private Sub ValidateIncrementalCalcColumns(ByVal mapCalc As Object)

    Dim requiredCols As Variant
    Dim c As Variant

    requiredCols = Array( _
        "ID", _
        "ParentID", _
        "IsSummary", _
        "Predecessors WBS", _
        "Baseline Start", _
        "Baseline Duration", _
        "Actual Start", _
        "Actual Finish", _
        "Forecast Start", _
        "Forecast Finish", _
        "Deadline", _
        "Constraint Active", _
        "Start Constraint Type", _
        "Start Constraint Date", _
        "Finish Constraint Type", _
        "Finish Constraint Date" _
    )

    For Each c In requiredCols
        If Not mapCalc.Exists(CStr(c)) Then
            Err.Raise vbObjectError + 9201, _
                "ValidateIncrementalCalcColumns", _
                "Missing required column in tbl_CALC: " & CStr(c)
        End If
    Next c

End Sub

Private Sub ValidateIncrementalStateColumns(ByVal mapState As Object)

    Dim requiredCols As Variant
    Dim c As Variant

    requiredCols = Array( _
        "ID", _
        "ParentID", _
        "IsSummary", _
        "Predecessors WBS", _
        "Baseline Start", _
        "Baseline Duration", _
        "Actual Start", _
        "Actual Finish", _
        "Forecast Start", _
        "Forecast Finish", _
        "Deadline", _
        "Constraint Active", _
        "Start Constraint Type", _
        "Start Constraint Date", _
        "Finish Constraint Type", _
        "Finish Constraint Date", _
        "Row Signature", _
        "Run Status" _
    )

    For Each c In requiredCols
        If Not mapState.Exists(CStr(c)) Then
            Err.Raise vbObjectError + 9202, _
                "ValidateIncrementalStateColumns", _
                "Missing required column in tbl_CALC_STATE: " & CStr(c)
        End If
    Next c

End Sub


Private Function BuildStateSignatureMap( _
    ByRef arrState As Variant, _
    ByVal mapState As Object) As Object

    Dim d As Object
    Dim r As Long
    Dim idVal As String

    Set d = CreateObject("Scripting.Dictionary")

    For r = 1 To UBound(arrState, 1)
        idVal = Trim$(CStr(arrState(r, mapState("ID"))))
        If idVal <> "" Then
            d(idVal) = Trim$(CStr(arrState(r, mapState("Row Signature"))))
        End If
    Next r

    Set BuildStateSignatureMap = d

End Function

Private Function GetStateRunStatus( _
    ByRef arrState As Variant, _
    ByVal mapState As Object) As String

    If UBound(arrState, 1) < 1 Then
        GetStateRunStatus = ""
        Exit Function
    End If

    GetStateRunStatus = Trim$(CStr(arrState(1, mapState("Run Status"))))

End Function

Private Function BuildCalcCurrentRowSignature( _
    ByRef arrCalc As Variant, _
    ByVal rowIdx As Long, _
    ByVal mapCalc As Object) As String

    Dim s As String

    s = ""
    s = s & "|ID=" & NormalizeIncrementalSignatureValue(arrCalc(rowIdx, mapCalc("ID")), "TEXT")
    s = s & "|ParentID=" & NormalizeIncrementalSignatureValue(arrCalc(rowIdx, mapCalc("ParentID")), "TEXT")
    s = s & "|IsSummary=" & NormalizeIncrementalSignatureValue(arrCalc(rowIdx, mapCalc("IsSummary")), "BOOLEAN")
    s = s & "|PredWBS=" & NormalizeIncrementalSignatureValue(arrCalc(rowIdx, mapCalc("Predecessors WBS")), "PREDWBS")

    s = s & "|BS=" & NormalizeIncrementalSignatureValue(arrCalc(rowIdx, mapCalc("Baseline Start")), "DATE")
    s = s & "|BD=" & NormalizeIncrementalSignatureValue(arrCalc(rowIdx, mapCalc("Baseline Duration")), "NUMBER")
    s = s & "|AS=" & NormalizeIncrementalSignatureValue(arrCalc(rowIdx, mapCalc("Actual Start")), "DATE")
    s = s & "|AF=" & NormalizeIncrementalSignatureValue(arrCalc(rowIdx, mapCalc("Actual Finish")), "DATE")
    s = s & "|FS=" & NormalizeIncrementalSignatureValue(arrCalc(rowIdx, mapCalc("Forecast Start")), "DATE")
    s = s & "|FF=" & NormalizeIncrementalSignatureValue(arrCalc(rowIdx, mapCalc("Forecast Finish")), "DATE")
    s = s & "|DL=" & NormalizeIncrementalSignatureValue(arrCalc(rowIdx, mapCalc("Deadline")), "DATE")
    s = s & "|CActive=" & NormalizeIncrementalSignatureValue(arrCalc(rowIdx, mapCalc("Constraint Active")), "BOOLEAN")
    s = s & "|SCType=" & NormalizeIncrementalSignatureValue(arrCalc(rowIdx, mapCalc("Start Constraint Type")), "TEXT")
    s = s & "|SCDate=" & NormalizeIncrementalSignatureValue(arrCalc(rowIdx, mapCalc("Start Constraint Date")), "DATE")
    s = s & "|FCType=" & NormalizeIncrementalSignatureValue(arrCalc(rowIdx, mapCalc("Finish Constraint Type")), "TEXT")
    s = s & "|FCDate=" & NormalizeIncrementalSignatureValue(arrCalc(rowIdx, mapCalc("Finish Constraint Date")), "DATE")

    BuildCalcCurrentRowSignature = s

End Function

Private Function GetArrVal_Incremental( _
    ByRef arr As Variant, _
    ByVal rowIdx As Long, _
    ByVal mapCol As Object, _
    ByVal colName As String) As Variant

    If Not mapCol.Exists(colName) Then
        GetArrVal_Incremental = Empty
        Exit Function
    End If

    GetArrVal_Incremental = arr(rowIdx, mapCol(colName))

End Function

Private Function NormalizeSignatureToken_Incremental(ByVal v As Variant) As String

    If isError(v) Then
        NormalizeSignatureToken_Incremental = "#ERR"
        Exit Function
    End If

    If IsEmpty(v) Then
        NormalizeSignatureToken_Incremental = ""
        Exit Function
    End If

    If VarType(v) = vbNull Then
        NormalizeSignatureToken_Incremental = ""
        Exit Function
    End If

    If Trim$(CStr(v)) = "" Then
        NormalizeSignatureToken_Incremental = ""
        Exit Function
    End If

    If IsDate(v) Then
        NormalizeSignatureToken_Incremental = Format$(CDate(v), "yyyy-mm-dd")
    Else
        NormalizeSignatureToken_Incremental = Trim$(CStr(v))
    End If

End Function


'=====================================================
' Build_Impacted_TaskIds
'=====================================================
' Entrée :
' - changedIds (Dictionary of String -> True)
' - childrenByPred (predId -> children list)
' - parentById (id -> parentId)
'
' Sortie :
' - impactedIds (Dictionary)
'
' Règles :
' - inclut changed
' - inclut tous les descendants
' - inclut tous les parents (rollup)
'=====================================================

Public Function Build_Impacted_TaskIds( _
    ByVal changedIds As Object, _
    ByVal childrenByPred As Object, _
    ByVal parentById As Object) As Object

    Dim impactedIds As Object
    Dim idVal As Variant
    Dim snapshotKeys As Variant
    Dim i As Long

    Set impactedIds = CreateObject("Scripting.Dictionary")

    If changedIds Is Nothing Then
        Set Build_Impacted_TaskIds = impactedIds
        Exit Function
    End If

    ' 1. Seed = changed tasks
    For Each idVal In changedIds.Keys
        impactedIds(CStr(idVal)) = True
    Next idVal

    ' 2. Add all downstream successors / descendants in the dependency network
    For Each idVal In changedIds.Keys
        AddDescendants_Incremental CStr(idVal), childrenByPred, impactedIds
    Next idVal

    ' 3. Add parents for roll-up refresh
    If impactedIds.Count > 0 Then
        snapshotKeys = impactedIds.Keys

        For i = LBound(snapshotKeys) To UBound(snapshotKeys)
            AddParents_Incremental CStr(snapshotKeys(i)), parentById, impactedIds
        Next i
    End If

    Set Build_Impacted_TaskIds = impactedIds

End Function


Private Sub AddDescendants( _
    ByVal rootId As String, _
    ByVal childrenByPred As Object, _
    ByVal impactedIds As Object)

    Dim stack As Object
    Dim currentId As String
    Dim child As Variant

    Set stack = CreateObject("System.Collections.ArrayList")
    stack.Add rootId

    Do While stack.Count > 0

        currentId = CStr(stack.item(stack.Count - 1))
        stack.RemoveAt stack.Count - 1

        If childrenByPred.Exists(currentId) Then
            For Each child In childrenByPred(currentId)

                If Not impactedIds.Exists(CStr(child)) Then
                    impactedIds(CStr(child)) = True
                    stack.Add CStr(child)
                End If

            Next child
        End If

    Loop

End Sub


Private Sub AddParents( _
    ByVal startId As String, _
    ByVal parentById As Object, _
    ByVal impactedIds As Object)

    Dim currentId As String
    Dim parentId As String

    currentId = startId

    Do While parentById.Exists(currentId)

        parentId = CStr(parentById(currentId))
        If parentId = "" Then Exit Do

        If Not impactedIds.Exists(parentId) Then
            impactedIds(parentId) = True
        End If

        currentId = parentId

    Loop

End Sub

Private Sub AddDescendants_Incremental( _
    ByVal rootId As String, _
    ByVal childrenByPred As Object, _
    ByVal impactedIds As Object)

    Dim queue As Collection
    Dim currentId As String
    Dim child As Variant

    If childrenByPred Is Nothing Then Exit Sub

    Set queue = New Collection
    queue.Add rootId

    Do While queue.Count > 0

        currentId = CStr(queue(1))
        queue.Remove 1

        If childrenByPred.Exists(currentId) Then
            For Each child In childrenByPred(currentId)

                If Not impactedIds.Exists(CStr(child)) Then
                    impactedIds(CStr(child)) = True
                    queue.Add CStr(child)
                End If

            Next child
        End If

    Loop

End Sub

Private Sub AddParents_Incremental( _
    ByVal startId As String, _
    ByVal parentById As Object, _
    ByVal impactedIds As Object)

    Dim currentId As String
    Dim parentId As String

    If parentById Is Nothing Then Exit Sub

    currentId = startId

    Do While parentById.Exists(currentId)

        parentId = Trim$(CStr(parentById(currentId)))
        If parentId = "" Then Exit Do

        If Not impactedIds.Exists(parentId) Then
            impactedIds(parentId) = True
        End If

        currentId = parentId

    Loop

End Sub


