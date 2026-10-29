Attribute VB_Name = "mod_CalcCoreEngine"
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
Public Const CORE_DIAGNOSTIC_SCHEMA_VERSION As Long = 1

'===============================================================================
' MODULE : mod_CalcCoreEngine
' DOMAINE / DOMAIN : Core Calculation
'
' FR
' Execute le calcul planning unique sur les taches feuilles, les contraintes, les dependances et le post-process LOE.
' Ne doit pas contourner les contrats publics des autres domaines.
'
' EN
' Runs the single planning calculation over leaf tasks, constraints, dependencies and LOE post-processing.
' Must not bypass public contracts owned by other domains.
'
' CONTRATS / CONTRACTS : Run_Calc_Core
' CALLBACKS EXTERNES / EXTERNAL CALLBACKS : Aucun / None
'===============================================================================

'------------------------------------------------------------------------------
' FR:
' Orchestre le calcul Core en memoire: prepare les index, nettoie le scope de
' recalcul, calcule les taches feuilles en ordre topologique, traite les LOE,
' propage les erreurs bloquantes puis consolide les summaries.
'
' EN:
' Orchestrates the in-memory Core calculation: builds indexes, clears the
' recalculation scope, computes leaf tasks in topological order, processes LOE
' tasks, propagates blocking errors, and rolls up summary dates.
'
' Entrees / Inputs:
' - dataArr avec les colonnes Core deja mappees dans mapCol.
' - linksBySuccId prepare par le Bridge, indexe par successeur.
' - recalcScope optionnel pour un recalcul partiel.
' - dictionnaires optionnels de diagnostics dependency/constraint/cascade.
'
' Sorties / Outputs:
' - Met a jour dans dataArr les dates calculees, durees, flags et messages.
' - Renseigne les diagnostics structures passes par le caller.
' - Preserve les sorties hors scope lors d'un recalcul partiel.
'
' Appele par / Called by:
' - Les adaptateurs Bridge/production qui preparent les donnees Core.
' - Les workflows TEST/SCENARIO qui executent le moteur en memoire.
'
' Notes:
' - Point d'entree stable du moteur Core; aucune lecture/ecriture feuille.
' - Les taches LOE sont exclues du topo standard puis calculees en post-process.
'------------------------------------------------------------------------------
Public Sub Run_Calc_Core( _
    ByRef dataArr As Variant, _
    ByVal mapCol As Object, _
    ByVal linksBySuccId As Object, _
    Optional ByVal recalcScope As Object, _
    Optional ByVal dependencyDiagnostics As Object, _
    Optional ByVal constraintDiagnostics As Object, _
    Optional ByVal cascadeDiagnostics As Object, _
    Optional ByVal executionNetwork As clsCompiledExecutionNetwork, _
    Optional ByVal coreDiagnostics As Object)

    Dim perfScope As clsPerfScope

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
    Dim useIndexedCore As Boolean
    Dim nodeIndex As Long
    Dim topoPosition As Long
    Dim nodeIds As Variant
    Dim rowsByNode As Variant
    Dim topoNodes As Variant
    Dim corePredOffsets As Variant
    Dim corePredNodes As Variant
    Dim corePredIds As Variant
    Dim corePredTypes As Variant
    Dim corePredLags As Variant
    Dim corePredSummarySourceIds As Variant
    Dim indexedInputCols As Variant
    Dim calcStartByNode As Variant
    Dim calcFinishByNode As Variant
    Dim hasCalcByNode As Variant
    Dim blockedByNode As Variant

    Set perfScope = Profiler_BeginScope("Run_Calc_Core", "Core Calculation")

    requiredCols = Array( _
        "ID", _
        "WBS", _
        "Task Name", _
        "ParentID", _
        "IsSummary", _
        "Task Type", _
        "Cal", _
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
    useIndexedCore = Not (executionNetwork Is Nothing) And Not isPartialMode

    If executionNetwork Is Nothing Then
        Set rowById = Core_BuildRowById(dataArr, mapCol)
        Set parentIds = Core_BuildParentIds(dataArr, mapCol, rowById)
        Set directChildrenById = Core_BuildDirectChildrenById(dataArr, mapCol, rowById)
        Set childrenByPred = Core_BuildChildrenByPred(rowById, linksBySuccId)
        Set loeIds = Core_BuildLOEIds(dataArr, mapCol, rowById)
        Set validIds = Core_BuildValidLeafIds(rowById, parentIds)
        Core_RemoveIdsFromDictionary validIds, loeIds
        Set indegree = Core_BuildIndegree(validIds, linksBySuccId)
        Set topoOrder = Core_TopoSortLeafNetwork(validIds, childrenByPred, indegree)
    Else
        Set rowById = executionNetwork.RowById
        Set parentIds = executionNetwork.ParentIds
        Set directChildrenById = executionNetwork.DirectChildrenById
        Set childrenByPred = executionNetwork.ChildrenByPred
        Set loeIds = executionNetwork.LoeIds
        Set validIds = executionNetwork.ValidLeafIds
        Set topoOrder = executionNetwork.TopoOrder
    End If

    Set calcStartById = CreateObject("Scripting.Dictionary")
    Set calcFinishById = CreateObject("Scripting.Dictionary")
    Set blockingErrors = CreateObject("Scripting.Dictionary")

    If Not coreDiagnostics Is Nothing Then
        If dependencyDiagnostics Is Nothing Then Set dependencyDiagnostics = CreateObject("Scripting.Dictionary")
        If constraintDiagnostics Is Nothing Then Set constraintDiagnostics = CreateObject("Scripting.Dictionary")
        If cascadeDiagnostics Is Nothing Then Set cascadeDiagnostics = CreateObject("Scripting.Dictionary")
        CoreDiagnostics_ClearScope coreDiagnostics, recalcScope
    End If

    If useIndexedCore Then
        nodeIds = executionNetwork.NodeIds
        rowsByNode = executionNetwork.RowsByNode
        topoNodes = executionNetwork.TopoNodes
        corePredOffsets = executionNetwork.CorePredOffsets
        corePredNodes = executionNetwork.CorePredNodes
        corePredIds = executionNetwork.CorePredIds
        corePredTypes = executionNetwork.CorePredTypes
        corePredLags = executionNetwork.CorePredLags
        corePredSummarySourceIds = executionNetwork.CorePredSummarySourceIds
        ReDim indexedInputCols(1 To CORE_INPUT_TASK_TYPE)
        indexedInputCols(CORE_INPUT_CAL) = CLng(mapCol("Cal"))
        indexedInputCols(CORE_INPUT_BASELINE_START) = CLng(mapCol("Baseline Start"))
        indexedInputCols(CORE_INPUT_BASELINE_DURATION) = CLng(mapCol("Baseline Duration"))
        indexedInputCols(CORE_INPUT_ACTUAL_START) = CLng(mapCol("Actual Start"))
        indexedInputCols(CORE_INPUT_ACTUAL_FINISH) = CLng(mapCol("Actual Finish"))
        indexedInputCols(CORE_INPUT_FORECAST_START) = CLng(mapCol("Forecast Start"))
        indexedInputCols(CORE_INPUT_FORECAST_FINISH) = CLng(mapCol("Forecast Finish"))
        indexedInputCols(CORE_INPUT_CONSTRAINT_ACTIVE) = CLng(mapCol("Constraint Active"))
        indexedInputCols(CORE_INPUT_START_CONSTRAINT_TYPE) = CLng(mapCol("Start Constraint Type"))
        indexedInputCols(CORE_INPUT_START_CONSTRAINT_DATE) = CLng(mapCol("Start Constraint Date"))
        indexedInputCols(CORE_INPUT_FINISH_CONSTRAINT_TYPE) = CLng(mapCol("Finish Constraint Type"))
        indexedInputCols(CORE_INPUT_FINISH_CONSTRAINT_DATE) = CLng(mapCol("Finish Constraint Date"))
        indexedInputCols(CORE_INPUT_TASK_TYPE) = CLng(mapCol("Task Type"))

        ReDim calcStartByNode(1 To executionNetwork.NodeCount)
        ReDim calcFinishByNode(1 To executionNetwork.NodeCount)
        ReDim hasCalcByNode(1 To executionNetwork.NodeCount)
        ReDim blockedByNode(1 To executionNetwork.NodeCount)
    End If

    If isPartialMode Then
        Core_LoadExistingCalcOutputs dataArr, mapCol, rowById, calcStartById, calcFinishById
        Core_ClearCalcOutputs_ForScope dataArr, mapCol, rowById, recalcScope, calcStartById, calcFinishById
    Else
        Core_ClearAllCalcOutputs dataArr, mapCol
    End If

    Core_ValidateLOEAsNonPredecessor dataArr, mapCol, rowById, linksBySuccId, loeIds, blockingErrors, coreDiagnostics

    If Core_HasTopoFailure(topoOrder, validIds) Then
        Core_MarkTopoFailure dataArr, mapCol, validIds, rowById, linksBySuccId, blockingErrors, coreDiagnostics
        GoTo SafeExit
    End If

    If useIndexedCore Then
        For topoPosition = LBound(topoNodes) To UBound(topoNodes)
            nodeIndex = CLng(topoNodes(topoPosition))
            currentId = CStr(nodeIds(nodeIndex))

            Core_ComputeOneLeafTask _
                dataArr, mapCol, CStr(currentId), rowById, linksBySuccId, _
                calcStartById, calcFinishById, blockingErrors, dependencyDiagnostics, constraintDiagnostics, cascadeDiagnostics, coreDiagnostics, _
                True, nodeIndex, rowsByNode, corePredOffsets, corePredNodes, corePredIds, corePredTypes, corePredLags, corePredSummarySourceIds, _
                indexedInputCols, calcStartByNode, calcFinishByNode, hasCalcByNode, blockedByNode

            blockedByNode(nodeIndex) = blockingErrors.Exists(CStr(currentId))
            Profiler_RecordCounter "CoreIndexedLeafCalls", 1
        Next topoPosition

        Core_MaterializeIndexedResults _
            nodeIds, calcStartByNode, calcFinishByNode, hasCalcByNode, _
            calcStartById, calcFinishById
    Else
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
                    calcStartById, calcFinishById, blockingErrors, dependencyDiagnostics, constraintDiagnostics, cascadeDiagnostics, coreDiagnostics
            End If

        Next currentId
    End If

    Core_ApplyLOEPostProcess dataArr, mapCol, rowById, linksBySuccId, loeIds, _
                             calcStartById, calcFinishById, blockingErrors, coreDiagnostics

    If blockingErrors.Count > 0 Then
        Core_PropagateBlockingErrors dataArr, mapCol, blockingErrors, childrenByPred, rowById, constraintDiagnostics, dependencyDiagnostics, cascadeDiagnostics, coreDiagnostics
    End If

    For Each currentId In calcStartById.Keys
        If rowById.Exists(CStr(currentId)) Then
            rowIdx = CLng(rowById(CStr(currentId)))
            Core_SetCalcTriplet dataArr, rowIdx, mapCol, _
                calcStartById(CStr(currentId)), _
                calcFinishById(CStr(currentId)), _
                NormalizeCalendarType(Core_GetVal(dataArr, rowIdx, mapCol, "Cal"))
        End If
    Next currentId

    Core_RollupSummaryDates dataArr, mapCol, rowById, directChildrenById, parentIds

SafeExit:
    CoreDiagnostics_MergeSpecialized coreDiagnostics, constraintDiagnostics, cascadeDiagnostics
End Sub


'------------------------------------------------------------------------------
' FR: Materialise une seule fois les buffers indexes pour les post-process historiques.
' EN: Materializes indexed buffers once for the historical post-process boundaries.
'------------------------------------------------------------------------------
Private Sub Core_MaterializeIndexedResults( _
    ByRef nodeIds As Variant, _
    ByRef calcStartByNode As Variant, _
    ByRef calcFinishByNode As Variant, _
    ByRef hasCalcByNode As Variant, _
    ByVal calcStartById As Object, _
    ByVal calcFinishById As Object)

    Dim nodeIndex As Long
    Dim taskId As String

    For nodeIndex = LBound(nodeIds) To UBound(nodeIds)
        If hasCalcByNode(nodeIndex) Then
            taskId = CStr(nodeIds(nodeIndex))
            calcStartById(taskId) = calcStartByNode(nodeIndex)
            calcFinishById(taskId) = calcFinishByNode(nodeIndex)
        End If
    Next nodeIndex

End Sub


'------------------------------------------------------------------------------
' FR:
' Recharge les dates calculees deja presentes dans dataArr afin qu'un recalcul
' partiel puisse reutiliser les predecesseurs et successeurs hors scope.
'
' EN:
' Reloads calculated dates already present in dataArr so a partial recalculation
' can reuse predecessors and successors that are outside the scope.
'
' Entrees / Inputs:
' - dataArr, mapCol et rowById du run Core courant.
' - Dictionnaires calcStartById et calcFinishById a alimenter.
'
' Sorties / Outputs:
' - Ajoute les dates Calculated Start/Finish existantes dans les caches Core.
'
' Appele par / Called by:
' - Run_Calc_Core en mode recalcul partiel.
'
' Notes:
' - Ne calcule rien; cette procedure ne fait que rehydrater l'etat existant.
'------------------------------------------------------------------------------
Private Sub Core_LoadExistingCalcOutputs( _
    ByRef dataArr As Variant, _
    ByVal mapCol As Object, _
    ByVal rowById As Object, _
    ByVal calcStartById As Object, _
    ByVal calcFinishById As Object)

    Dim perfScope As clsPerfScope

    Dim idVal As Variant
    Dim rowIdx As Long
    Dim calcStart As Variant
    Dim calcFinish As Variant

    Set perfScope = Profiler_BeginScope("Core_LoadExistingCalcOutputs", "Array")

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


'------------------------------------------------------------------------------
' FR:
' Calcule une tache feuille standard a partir de ses dates terrain, de ses
' contraintes et de ses liens predecesseurs; enregistre les blocages metier
' des que la tache devient non calculable.
'
' EN:
' Computes one standard leaf task from actual/forecast dates, constraints, and
' predecessor links; records business blocking errors as soon as the task cannot
' be scheduled consistently.
'
' Entrees / Inputs:
' - currentId, dataArr, mapCol et rowById pour lire la ligne de tache.
' - linksBySuccId, directChildrenById et calcStart/FinishById pour les dependances.
' - Dictionnaires de diagnostics dependency/constraint et erreurs bloquantes.
'
' Sorties / Outputs:
' - Ajoute start/finish calcules dans calcStartById et calcFinishById.
' - Ajoute Error flag/ErrorMsg et diagnostics en cas de violation.
'
' Appele par / Called by:
' - Run_Calc_Core pendant le parcours topologique des taches feuilles.
'
' Notes:
' - Gere FS/SS/FF, lag, sources summary, contraintes start/finish et dates
'   actual/forecast.
' - Les LOE ne passent pas par ce calcul; elles sont traitees ensuite.
'------------------------------------------------------------------------------
Private Sub Core_PropagateBlockingErrors( _
    ByRef dataArr As Variant, _
    ByVal mapCol As Object, _
    ByVal blockingErrors As Object, _
    ByVal childrenByPred As Object, _
    ByVal rowById As Object, _
    Optional ByVal constraintDiagnostics As Object, _
    Optional ByVal dependencyDiagnostics As Object, _
    Optional ByVal cascadeDiagnostics As Object, _
    Optional ByVal coreDiagnostics As Object)

    Dim allErrorIds As Object
    Dim taskId As Variant
    Dim rowIdx As Long

    Set allErrorIds = CreateObject("Scripting.Dictionary")

    For Each taskId In blockingErrors.Keys
        allErrorIds(CStr(taskId)) = True
    Next taskId

    For Each taskId In blockingErrors.Keys
        Core_PropagateErrorToChildren CStr(taskId), childrenByPred, allErrorIds
        Core_RecordCascadeDiagnosticsForRoot CStr(taskId), childrenByPred, constraintDiagnostics, dependencyDiagnostics, cascadeDiagnostics
    Next taskId

    For Each taskId In allErrorIds.Keys
        If rowById.Exists(CStr(taskId)) Then
            rowIdx = CLng(rowById(CStr(taskId)))
            Core_SetErrorFlag_Row dataArr, rowIdx, mapCol, True

            If Not blockingErrors.Exists(CStr(taskId)) Then
                Core_AddErrorMessage_Row dataArr, rowIdx, mapCol, _
                    TextCatalog_Get("CORE.ERROR.BLOCKED_PREDECESSOR_CHAIN", TEXT_LANGUAGE_EN), True
                CoreDiagnostics_Record coreDiagnostics, CStr(taskId), _
                    "CORE.ERROR.BLOCKED_PREDECESSOR_CHAIN", Nothing, "INHERITED", vbNullString, "CASCADE"
            End If
        End If
    Next taskId

End Sub


'------------------------------------------------------------------------------
' FR:
' Cree les diagnostics de cascade pour tous les descendants d'une erreur racine,
' en conservant la cause initiale et le parent qui propage le blocage.
'
' EN:
' Creates cascade diagnostics for all descendants of one root error, preserving
' the initial cause and the parent that propagated the block.
'
' Entrees / Inputs:
' - rootId, childrenByPred, allErrorIds et diagnostics racines disponibles.
'
' Sorties / Outputs:
' - cascadeDiagnostics enrichi pour chaque descendant impacte.
'
' Appele par / Called by:
' - Core_PropagateBlockingErrors, une fois par erreur racine.
'
' Notes:
' - Parcours en largeur pour produire une chaine de responsabilite lisible.
'------------------------------------------------------------------------------
Private Sub Core_RecordCascadeDiagnosticsForRoot( _
    ByVal rootId As String, _
    ByVal childrenByPred As Object, _
    ByVal constraintDiagnostics As Object, _
    ByVal dependencyDiagnostics As Object, _
    ByVal cascadeDiagnostics As Object)

    Dim q As Collection
    Dim currentId As Variant
    Dim childId As Variant
    Dim parentById As Object
    Dim diag As Object
    Dim rootType As String

    If cascadeDiagnostics Is Nothing Then Exit Sub
    If childrenByPred Is Nothing Then Exit Sub

    If Not constraintDiagnostics Is Nothing Then
        If constraintDiagnostics.Exists(CStr(rootId)) Then rootType = "CONSTRAINT"
    End If

    If rootType = "" Then
        If Not dependencyDiagnostics Is Nothing Then
            If dependencyDiagnostics.Exists(CStr(rootId)) Then rootType = "DEPENDENCY"
        End If
    End If

    If rootType = "" Then rootType = "CORE"

    Set q = New Collection
    Set parentById = CreateObject("Scripting.Dictionary")
    q.Add CStr(rootId)
    parentById(CStr(rootId)) = ""

    Do While q.Count > 0
        currentId = q(1)
        q.Remove 1

        If childrenByPred.Exists(CStr(currentId)) Then
            For Each childId In childrenByPred(CStr(currentId))
                If Not parentById.Exists(CStr(childId)) Then
                    parentById(CStr(childId)) = CStr(currentId)
                    q.Add CStr(childId)

                    Set diag = CreateObject("Scripting.Dictionary")
                    diag("TaskID") = CStr(childId)
                    diag("RootErrorID") = CStr(rootId)
                    diag("RootErrorType") = rootType
                    diag("ParentPropagatedFrom") = CStr(currentId)

                    If rootType = "CONSTRAINT" Then
                        If Not constraintDiagnostics Is Nothing Then
                            If constraintDiagnostics.Exists(CStr(rootId)) Then Set diag("RootConstraintDiagnostic") = constraintDiagnostics(CStr(rootId))
                        End If
                    ElseIf rootType = "DEPENDENCY" Then
                        If Not dependencyDiagnostics Is Nothing Then
                            If dependencyDiagnostics.Exists(CStr(rootId)) Then Set diag("RootDependencyDiagnostic") = dependencyDiagnostics(CStr(rootId))
                        End If
                    End If

                    Set cascadeDiagnostics.Item(CStr(childId)) = diag
                End If
            Next childId
        End If
    Loop

End Sub

'------------------------------------------------------------------------------
' FR: Efface les sorties calculees uniquement pour les IDs du recalcul partiel et retire leurs caches en memoire.
' EN: Clears calculated outputs only for partial recalculation IDs and removes their in-memory cache entries.
'------------------------------------------------------------------------------
Private Sub Core_ClearCalcOutputs_ForScope( _
    ByRef dataArr As Variant, _
    ByVal mapCol As Object, _
    ByVal rowById As Object, _
    ByVal recalcScope As Object, _
    Optional ByVal calcStartById As Object, _
    Optional ByVal calcFinishById As Object)

    Dim perfScope As clsPerfScope

    Dim idKey As Variant
    Dim taskId As String
    Dim rowIdx As Long

    Set perfScope = Profiler_BeginScope("Core_ClearCalcOutputs_ForScope", "Array")

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


'------------------------------------------------------------------------------
' FR: Identifie les taches Level of Effort afin de les exclure du calcul topo standard et de les traiter a part.
' EN: Identifies Level of Effort tasks so they can be excluded from standard topological calculation and handled separately.
'------------------------------------------------------------------------------
Private Function Core_BuildLOEIds( _
    ByRef dataArr As Variant, _
    ByVal mapCol As Object, _
    ByVal rowById As Object) As Object

    Dim perfScope As clsPerfScope

    Dim d As Object
    Dim idVal As Variant
    Dim rowIdx As Long
    Dim taskType As String

    Set perfScope = Profiler_BeginScope("Core_BuildLOEIds", "Dictionary")

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

'------------------------------------------------------------------------------
' FR: Normalise les libelles de type de tache consommes par le Core.
' EN: Normalizes task type labels consumed by the Core.
'------------------------------------------------------------------------------
Public Function Core_NormalizeTaskType(ByVal rawValue As Variant) As String

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

'------------------------------------------------------------------------------
' FR: Retire d'un dictionnaire de travail tous les IDs presents dans un second dictionnaire.
' EN: Removes from a working dictionary every ID present in a second dictionary.
'------------------------------------------------------------------------------
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

'------------------------------------------------------------------------------
' FR: Valide Core Validate LOE As Non Predecessor et applique la politique d'erreur definie par le composant.
' EN: Validates Core Validate LOE As Non Predecessor and applies the component's defined failure policy.
'------------------------------------------------------------------------------

Private Sub Core_ValidateLOEAsNonPredecessor( _
    ByRef dataArr As Variant, _
    ByVal mapCol As Object, _
    ByVal rowById As Object, _
    ByVal linksBySuccId As Object, _
    ByVal loeIds As Object, _
    ByVal blockingErrors As Object, _
    ByVal coreDiagnostics As Object)

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
                                TextCatalog_Get("CORE.ERROR.LOE_AS_PREDECESSOR", TEXT_LANGUAGE_EN), _
                                "CORE.ERROR.LOE_AS_PREDECESSOR", Nothing, CStr(succId), coreDiagnostics, "ROOT", "LOE"
                        End If

                        If rowById.Exists(CStr(succId)) Then
                            succRow = CLng(rowById(CStr(succId)))
                            Core_AddBlockingError dataArr, succRow, mapCol, blockingErrors, CStr(succId), _
                                TextCatalog_Format("CORE.ERROR.INVALID_LOE_PREDECESSOR_ID", TEXT_LANGUAGE_EN, _
                                    TextCatalog_Arguments("Id", predId)), _
                                "CORE.ERROR.INVALID_LOE_PREDECESSOR_ID", TextCatalog_Arguments("Id", predId), _
                                predId, coreDiagnostics, "ROOT", "LOE"
                        End If

                    End If
                End If

            Next oneLink

        End If

    Next succId

End Sub


'------------------------------------------------------------------------------
' FR:
' Calcule les taches Level of Effort apres les taches standard, a partir de
' liens SS pour le debut et FF pour la fin, puis neutralise les champs analytiques
' qui ne s'appliquent pas a une LOE.
'
' EN:
' Computes Level of Effort tasks after standard tasks, using SS links for start
' and FF links for finish, then clears analytics fields that do not apply to LOE.
'
' Entrees / Inputs:
' - loeIds, dataArr, mapCol, rowById, linksBySuccId et directChildrenById.
' - calcStartById/calcFinishById issus du calcul standard.
' - blockingErrors pour signaler les LOE invalides ou non calculables.
'
' Sorties / Outputs:
' - Dates calculees LOE ajoutees aux caches Core.
' - Erreurs bloquantes sur LOE si les liens SS/FF requis sont absents ou invalides.
' - Champs Driving Logic, Critical Path, Total/Free Float et REX nettoyes si presents.
'
' Appele par / Called by:
' - Run_Calc_Core apres le parcours topologique des taches non-LOE.
'
' Notes:
' - Les LOE doivent avoir au moins un SS et un FF; FS et autres types sont refuses.
' - Les sources summary SS sont regroupees pour choisir le debut le plus tot du groupe.

