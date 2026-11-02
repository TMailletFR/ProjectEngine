Attribute VB_Name = "mod_CoreLOEPostProcess"
Option Explicit

'------------------------------------------------------------------------------
Public Sub Core_ApplyLOEPostProcess( _
    ByRef dataArr As Variant, _
    ByVal mapCol As Object, _
    ByVal rowById As Object, _
    ByVal linksBySuccId As Object, _
    ByVal loeIds As Object, _
    ByVal calcStartById As Object, _
    ByVal calcFinishById As Object, _
    ByVal blockingErrors As Object, _
    ByVal coreDiagnostics As Object)

    Dim perfScope As clsPerfScope

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
    Dim calType As String

    Set perfScope = Profiler_BeginScope("Core_ApplyLOEPostProcess", "Core LOE")

    If loeIds Is Nothing Then Exit Sub
    If loeIds.Count = 0 Then Exit Sub

    For Each loeId In loeIds.Keys

        If Not rowById.Exists(CStr(loeId)) Then GoTo NextLOE
        rowIdx = CLng(rowById(CStr(loeId)))
        calType = NormalizeCalendarType(Core_GetVal(dataArr, rowIdx, mapCol, "Cal"))

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
                                    TextCatalog_Get("CORE.ERROR.LOE_SS_PREDECESSOR_MISSING", TEXT_LANGUAGE_EN), _
                                    "CORE.ERROR.LOE_SS_PREDECESSOR_MISSING", Nothing, vbNullString, coreDiagnostics, "ROOT", "LOE"
                                GoTo NextLOE
                            End If

                            If Not rowById.Exists(predId) Then
                                Core_AddBlockingError dataArr, rowIdx, mapCol, blockingErrors, CStr(loeId), _
                                    TextCatalog_Format("CORE.ERROR.LOE_SS_PREDECESSOR_NOT_FOUND", TEXT_LANGUAGE_EN, _
                                        TextCatalog_Arguments("Id", predId)), _
                                    "CORE.ERROR.LOE_SS_PREDECESSOR_NOT_FOUND", TextCatalog_Arguments("Id", predId), _
                                    predId, coreDiagnostics, "ROOT", "LOE"
                                GoTo NextLOE
                            End If

                            If blockingErrors.Exists(predId) Then
                                Core_AddBlockingError dataArr, rowIdx, mapCol, blockingErrors, CStr(loeId), _
                                    TextCatalog_Format("CORE.ERROR.LOE_SS_PREDECESSOR_BLOCKED", TEXT_LANGUAGE_EN, _
                                        TextCatalog_Arguments("Id", predId)), _
                                    "CORE.ERROR.LOE_SS_PREDECESSOR_BLOCKED", TextCatalog_Arguments("Id", predId), _
                                    predId, coreDiagnostics, "INHERITED", "LOE"
                                GoTo NextLOE
                            End If

                            If Not calcStartById.Exists(predId) Then
                                Core_AddBlockingError dataArr, rowIdx, mapCol, blockingErrors, CStr(loeId), _
                                    TextCatalog_Format("CORE.ERROR.LOE_SS_START_UNAVAILABLE", TEXT_LANGUAGE_EN, _
                                        TextCatalog_Arguments("Id", predId)), _
                                    "CORE.ERROR.LOE_SS_START_UNAVAILABLE", TextCatalog_Arguments("Id", predId), _
                                    predId, coreDiagnostics, "ROOT", "LOE"
                                GoTo NextLOE
                            End If

                            candidateStart = ApplyLag(calcStartById(predId), lagVal, calType, "SS")

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
                                    TextCatalog_Get("CORE.ERROR.LOE_FF_PREDECESSOR_MISSING", TEXT_LANGUAGE_EN), _
                                    "CORE.ERROR.LOE_FF_PREDECESSOR_MISSING", Nothing, vbNullString, coreDiagnostics, "ROOT", "LOE"
                                GoTo NextLOE
                            End If

                            If Not rowById.Exists(predId) Then
                                Core_AddBlockingError dataArr, rowIdx, mapCol, blockingErrors, CStr(loeId), _
                                    TextCatalog_Format("CORE.ERROR.LOE_FF_PREDECESSOR_NOT_FOUND", TEXT_LANGUAGE_EN, _
                                        TextCatalog_Arguments("Id", predId)), _
                                    "CORE.ERROR.LOE_FF_PREDECESSOR_NOT_FOUND", TextCatalog_Arguments("Id", predId), _
                                    predId, coreDiagnostics, "ROOT", "LOE"
                                GoTo NextLOE
                            End If

                            If blockingErrors.Exists(predId) Then
                                Core_AddBlockingError dataArr, rowIdx, mapCol, blockingErrors, CStr(loeId), _
                                    TextCatalog_Format("CORE.ERROR.LOE_FF_PREDECESSOR_BLOCKED", TEXT_LANGUAGE_EN, _
                                        TextCatalog_Arguments("Id", predId)), _
                                    "CORE.ERROR.LOE_FF_PREDECESSOR_BLOCKED", TextCatalog_Arguments("Id", predId), _
                                    predId, coreDiagnostics, "INHERITED", "LOE"
                                GoTo NextLOE
                            End If

                            If Not calcFinishById.Exists(predId) Then
                                Core_AddBlockingError dataArr, rowIdx, mapCol, blockingErrors, CStr(loeId), _
                                    TextCatalog_Format("CORE.ERROR.LOE_FF_FINISH_UNAVAILABLE", TEXT_LANGUAGE_EN, _
                                        TextCatalog_Arguments("Id", predId)), _
                                    "CORE.ERROR.LOE_FF_FINISH_UNAVAILABLE", TextCatalog_Arguments("Id", predId), _
                                    predId, coreDiagnostics, "ROOT", "LOE"
                                GoTo NextLOE
                            End If

                            candidateFinish = ApplyLag(calcFinishById(predId), lagVal, calType, "FF")
                            loeFinish = Core_MaxDateIfBoth(loeFinish, candidateFinish)

                        Case Else
                            invalidCount = invalidCount + 1

                    End Select

                Next oneLink

            End If
        End If

        If ssCount = 0 Then
            Core_AddBlockingError dataArr, rowIdx, mapCol, blockingErrors, CStr(loeId), _
                TextCatalog_Get("CORE.ERROR.LOE_SS_REQUIRED", TEXT_LANGUAGE_EN), _
                "CORE.ERROR.LOE_SS_REQUIRED", Nothing, vbNullString, coreDiagnostics, "ROOT", "LOE"
            GoTo NextLOE
        End If

        If ffCount = 0 Then
            Core_AddBlockingError dataArr, rowIdx, mapCol, blockingErrors, CStr(loeId), _
                TextCatalog_Get("CORE.ERROR.LOE_FF_REQUIRED", TEXT_LANGUAGE_EN), _
                "CORE.ERROR.LOE_FF_REQUIRED", Nothing, vbNullString, coreDiagnostics, "ROOT", "LOE"
            GoTo NextLOE
        End If

        If invalidCount > 0 Then
            Core_AddBlockingError dataArr, rowIdx, mapCol, blockingErrors, CStr(loeId), _
                TextCatalog_Get("CORE.ERROR.LOE_LINK_TYPE", TEXT_LANGUAGE_EN), _
                "CORE.ERROR.LOE_LINK_TYPE", Nothing, vbNullString, coreDiagnostics, "ROOT", "LOE"
            GoTo NextLOE
        End If

        For Each parentKey In summarySSStartBySource.Keys
            summarySSStart = Core_MaxDateIfBoth(summarySSStart, summarySSStartBySource(CStr(parentKey)))
        Next parentKey

        loeStart = Core_MaxDateIfBoth(normalSSStart, summarySSStart)

        If Not HasValue(loeStart) Then
            Core_AddBlockingError dataArr, rowIdx, mapCol, blockingErrors, CStr(loeId), _
                TextCatalog_Get("CORE.ERROR.LOE_START_NOT_COMPUTABLE", TEXT_LANGUAGE_EN), _
                "CORE.ERROR.LOE_START_NOT_COMPUTABLE", Nothing, vbNullString, coreDiagnostics, "ROOT", "LOE"
            GoTo NextLOE
        End If

        If Not HasValue(loeFinish) Then
            Core_AddBlockingError dataArr, rowIdx, mapCol, blockingErrors, CStr(loeId), _
                TextCatalog_Get("CORE.ERROR.LOE_FINISH_NOT_COMPUTABLE", TEXT_LANGUAGE_EN), _
                "CORE.ERROR.LOE_FINISH_NOT_COMPUTABLE", Nothing, vbNullString, coreDiagnostics, "ROOT", "LOE"
            GoTo NextLOE
        End If

        If CDbl(loeFinish) < CDbl(loeStart) Then
            Core_AddBlockingError dataArr, rowIdx, mapCol, blockingErrors, CStr(loeId), _
                TextCatalog_Get("CORE.ERROR.LOE_FINISH_BEFORE_START", TEXT_LANGUAGE_EN), _
                "CORE.ERROR.LOE_FINISH_BEFORE_START", Nothing, vbNullString, coreDiagnostics, "ROOT", "LOE"
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

