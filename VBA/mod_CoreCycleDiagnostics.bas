Attribute VB_Name = "mod_CoreCycleDiagnostics"
Option Explicit

'------------------------------------------------------------------------------
Public Sub Core_MarkTopoFailure( _
    ByRef dataArr As Variant, _
    ByVal mapCol As Object, _
    ByVal validIds As Object, _
    ByVal rowById As Object, _
    ByVal linksBySuccId As Object, _
    ByVal blockingErrors As Object, _
    ByVal coreDiagnostics As Object)

    Dim taskId As Variant
    Dim rowIdx As Long
    Dim cycleMessage As String
    Dim cycleText As String
    Dim args As Object

    cycleMessage = Core_BuildTopoCycleDiagnosticMessage(dataArr, mapCol, validIds, rowById, linksBySuccId, cycleText)
    If Trim$(cycleMessage) = "" Then cycleMessage = TextCatalog_Get("DIAG.CALC_ENGINE.CYCLE_MARKER", TEXT_LANGUAGE_EN)

    For Each taskId In validIds.Keys
        If rowById.Exists(CStr(taskId)) Then
            rowIdx = CLng(rowById(CStr(taskId)))
            blockingErrors(CStr(taskId)) = True
            Core_AddErrorMessage_Row dataArr, rowIdx, mapCol, cycleMessage, True
            Set args = TextCatalog_Arguments("Cycle", cycleText)
            CoreDiagnostics_Record coreDiagnostics, CStr(taskId), _
                "DIAG.CALC_ENGINE.CYCLE_MARKER", args, "ROOT", vbNullString, "CYCLE"
        End If
    Next taskId

End Sub

'------------------------------------------------------------------------------
' FR: Construit le message utilisateur pour un cycle logique detecte par le tri topologique.
' EN: Builds the user-facing message for a logical cycle detected by the topological sort.
'------------------------------------------------------------------------------
Private Function Core_BuildTopoCycleDiagnosticMessage( _
    ByRef dataArr As Variant, _
    ByVal mapCol As Object, _
    ByVal validIds As Object, _
    ByVal rowById As Object, _
    ByVal linksBySuccId As Object, _
    Optional ByRef cycleTextOut As String) As String

    Dim cycleEdges As Collection
    Dim cycleText As String

    On Error GoTo SafeFallback

    Set cycleEdges = Core_FindFirstTopoCycleEdges(validIds, linksBySuccId)

    If cycleEdges Is Nothing Then GoTo SafeFallback
    If cycleEdges.Count = 0 Then GoTo SafeFallback

    cycleText = Core_FormatTopoCycleEdges(dataArr, mapCol, rowById, cycleEdges)
    cycleTextOut = cycleText
    If Trim$(cycleText) = "" Then GoTo SafeFallback

    Core_BuildTopoCycleDiagnosticMessage = _
        TextCatalog_Get("DIAG.CALC_ENGINE.CYCLE_MARKER", TEXT_LANGUAGE_EN) & vbCrLf & _
        PlanningMessageText_Format("DIAG.CALC_ENGINE.TOPO_CYCLE_DETAIL", _
            TextCatalog_Arguments("Cycle", cycleText), _
            TextCatalog_Arguments("Cycle", cycleText))
    Exit Function

SafeFallback:
    Core_BuildTopoCycleDiagnosticMessage = TextCatalog_Get("DIAG.CALC_ENGINE.CYCLE_MARKER", TEXT_LANGUAGE_EN)

End Function

'------------------------------------------------------------------------------
' FR: Recherche dans le graphe des liens le premier cycle entre taches feuilles valides.
' EN: Searches the dependency graph for the first cycle between valid leaf tasks.
'------------------------------------------------------------------------------
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

'------------------------------------------------------------------------------
' FR: Parcourt le graphe en profondeur et extrait les aretes du premier cycle rencontre.
' EN: Traverses the graph depth-first and extracts the edges of the first encountered cycle.
'------------------------------------------------------------------------------
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

'------------------------------------------------------------------------------
' FR: Transforme les aretes d'un cycle en texte lisible avec taches, type de lien et lag.
' EN: Formats cycle edges as readable text with tasks, link type, and lag.
'------------------------------------------------------------------------------
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

'------------------------------------------------------------------------------
' FR:
' Cree le diagnostic source qui explique quelle dependance impose une date de
' debut minimale plus tardive que la date demandee.
'
' EN:
' Creates the source diagnostic explaining which dependency imposes a minimum
' start date later than the requested date.
'
' Entrees / Inputs:
' - Task cible, predecesseur bloquant, type de lien, lag et dates candidates.
' - Source eventuelle quand le predecesseur bloquant provient d'un summary.
'
' Sorties / Outputs:
' - Dictionnaire diagnostic pret a etre stocke ou enrichi.
'
' Appele par / Called by:
' - Core_ComputeOneLeafTask pendant l'analyse des liens de debut.
'
' Notes:
' - Ce diagnostic est la base des messages dependency et des cascades aval.
'------------------------------------------------------------------------------



'------------------------------------------------------------------------------
' FR:
' Stocke le diagnostic structure lorsqu'une Forecast Start est anterieure a la
' date minimale imposee par les dependances.
'
' EN:
' Stores the structured diagnostic when a Forecast Start is earlier than the
' minimum start date imposed by dependencies.
'
' Entrees / Inputs:
' - Task cible, Forecast Start demandee, date minimale autorisee.
' - Diagnostic source de la dependance bloquante.
'
' Sorties / Outputs:
' - dependencyDiagnostics(taskId) avec les champs de cause et de date.
'
' Appele par / Called by:
' - Core_ComputeOneLeafTask au moment de valider Forecast Start.
'
' Notes:
' - Ne formule pas le message utilisateur; elle preserve les donnees auditables.
'------------------------------------------------------------------------------
Public Sub Core_RecordForecastStartDependencyDiagnostic( _
    ByVal dependencyDiagnostics As Object, _
    ByVal taskId As String, _
    ByVal requestedStart As Variant, _
    ByVal minimumAllowedStart As Variant, _
    ByVal blockingPredecessorId As String, _
    ByVal blockingLinkType As String, _
    ByVal blockingLag As Double, _
    ByVal blockingCandidateDate As Variant, _
    ByVal blockingPredecessorDate As Variant, _
    ByVal blockingPredecessorDateKind As String, _
    ByVal expandedFrom As String)

    Dim diag As Object

    If dependencyDiagnostics Is Nothing Then Exit Sub
    If blockingPredecessorId = "" Then Exit Sub

    Set diag = CreateObject("Scripting.Dictionary")
    If CoreLeafProfile_IsEnabled() Then CoreLeafProfile_Count "DependencyDiagnosticDictionaries"

    diag("TaskID") = CStr(taskId)
    diag("RequestedStart") = requestedStart
    diag("MinimumAllowedStart") = minimumAllowedStart
    diag("BlockingPredecessorID") = blockingPredecessorId
    diag("BlockingLinkType") = UCase$(Trim$(blockingLinkType))
    diag("BlockingLag") = blockingLag
    diag("BlockingCandidateDate") = blockingCandidateDate
    diag("BlockingPredecessorDate") = blockingPredecessorDate
    diag("BlockingPredecessorDateKind") = UCase$(Trim$(blockingPredecessorDateKind))
    diag("ExpandedFrom") = expandedFrom

    Set dependencyDiagnostics.Item(CStr(taskId)) = diag

End Sub

'------------------------------------------------------------------------------
' FR: Construit le libelle humain d'une tache de cycle a partir du WBS, du nom ou de l'ID.
' EN: Builds the human-readable label for a cycle task from WBS, task name, or ID.
'------------------------------------------------------------------------------
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
        Core_CycleTaskLabel = TextCatalog_Format( _
            "DIAG.COMMON.ID_TASK_LABEL", TEXT_LANGUAGE_EN, _
            TextCatalog_Arguments("Id", taskId))
    End If
    Exit Function

Fallback:
    Core_CycleTaskLabel = TextCatalog_Format( _
        "DIAG.COMMON.ID_TASK_LABEL", TEXT_LANGUAGE_EN, _
        TextCatalog_Arguments("Id", taskId))

End Function

'------------------------------------------------------------------------------
' FR: Formate le lag d'un lien de cycle avec signe explicite et decimal standardise.
' EN: Formats a cycle link lag with an explicit sign and standardized decimal separator.
'------------------------------------------------------------------------------
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

'------------------------------------------------------------------------------
' FR:
' Propage les erreurs bloquantes racines vers les successeurs afin que toute la
' chaine aval soit marquee non calculable.
'
' EN:
' Propagates root blocking errors to successors so the full downstream chain is
' marked as non-computable.
'
' Entrees / Inputs:
' - blockingErrors racines, childrenByPred, rowById, dataArr et mapCol.
' - cascadeDiagnostics optionnel pour tracer la propagation.
'
' Sorties / Outputs:
' - Error flag/ErrorMsg sur les descendants impactes.
' - cascadeDiagnostics rempli avec la racine et le parent de propagation.
'
' Appele par / Called by:
' - Run_Calc_Core apres le calcul standard et le post-process LOE.
'
' Notes:
' - Les descendants recoivent un message de chaine bloquee s'ils n'ont pas deja
'   leur propre erreur racine.

