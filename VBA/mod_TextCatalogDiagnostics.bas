Attribute VB_Name = "mod_TextCatalogDiagnostics"
Option Explicit

' Preserve historical spelling/whitespace: these texts participate in event hashes.
Public Function TextCatalogDiagnostics_Definitions() As Variant
    Dim definitions(0 To 255) As Variant
    definitions(0) = Array("DIAG.LOE.EXPLICIT.MESSAGE", _
            "LOE used as predecessor" & vbCrLf & vbCrLf & "{Details}" & vbCrLf & vbCrLf & _
            "A LOE is driven by the network but must not drive other tasks." & vbCrLf & vbCrLf & _
            "-> replace the LOE with a real leaf task or milestone.", _
            "LOE utilisee comme predecesseur" & vbCrLf & vbCrLf & "{Details}" & vbCrLf & vbCrLf & _
            "Une LOE est pilotee par le reseau mais ne doit pas piloter d'autres taches." & vbCrLf & vbCrLf & _
            "-> remplacer la LOE par une vraie tache feuille ou une milestone.")
    definitions(1) = Array("DIAG.LOE.PARENT.MESSAGE", _
            "LOE used as predecessor through a parent link" & vbCrLf & vbCrLf & "{Details}" & vbCrLf & vbCrLf & _
            "-> replace the parent with a leaf task or finish milestone.", _
            "LOE utilisee comme predecesseur via un lien parent" & vbCrLf & vbCrLf & "{Details}" & vbCrLf & vbCrLf & _
            "-> remplacer le parent par une tache feuille ou une milestone de fin.")
    definitions(2) = Array("DIAG.GROUP.LOE.PREDECESSOR", _
        "LOE used as predecessor" & vbCrLf & "-> " & "remove the LOE from upstream logic; a LOE is driven by the network but must not drive other tasks" & vbCrLf & vbCrLf & _
        "IDs: {Ids}" & vbCrLf & "WBS: {Wbs}", _
        "LOE utilisée comme prédécesseur" & vbCrLf & "-> " & "supprimer la LOE de la logique amont ; une LOE est pilotée par le réseau mais ne doit pas piloter d'autres tâches" & vbCrLf & vbCrLf & _
        "IDs : {Ids}" & vbCrLf & "WBS : {Wbs}")
    definitions(3) = Array("DIAG.GROUP.LOE.MISSING_SS", _
        "LOE without usable SS link" & vbCrLf & "-> " & "add at least one valid SS predecessor to define the LOE start" & vbCrLf & vbCrLf & _
        "IDs: {Ids}" & vbCrLf & "WBS: {Wbs}", _
        "LOE sans lien SS exploitable" & vbCrLf & "-> " & "ajouter au moins un prédécesseur SS valide pour définir le début de la LOE" & vbCrLf & vbCrLf & _
        "IDs : {Ids}" & vbCrLf & "WBS : {Wbs}")
    definitions(4) = Array("DIAG.GROUP.LOE.MISSING_FF", _
        "LOE without usable FF link" & vbCrLf & "-> " & "add at least one valid FF predecessor to define the LOE finish" & vbCrLf & vbCrLf & _
        "IDs: {Ids}" & vbCrLf & "WBS: {Wbs}", _
        "LOE sans lien FF exploitable" & vbCrLf & "-> " & "ajouter au moins un prédécesseur FF valide pour définir la fin de la LOE" & vbCrLf & vbCrLf & _
        "IDs : {Ids}" & vbCrLf & "WBS : {Wbs}")
    definitions(5) = Array("DIAG.GROUP.LOE.INVALID_LINK", _
        "Invalid link on LOE" & vbCrLf & "-> " & "a LOE only supports SS and FF links" & vbCrLf & vbCrLf & _
        "IDs: {Ids}" & vbCrLf & "WBS: {Wbs}", _
        "Lien invalide sur LOE" & vbCrLf & "-> " & "une LOE accepte uniquement des liens SS et FF" & vbCrLf & vbCrLf & _
        "IDs : {Ids}" & vbCrLf & "WBS : {Wbs}")
    definitions(6) = Array("DIAG.GROUP.DEPENDENCY.MISSING_PREDECESSOR", _
        "Missing predecessor" & vbCrLf & "-> " & "check the Predecessors WBS column" & vbCrLf & vbCrLf & _
        "IDs: {Ids}" & vbCrLf & "WBS: {Wbs}", _
        "Prédécesseur introuvable" & vbCrLf & "-> " & "vérifier la colonne Predecessors WBS" & vbCrLf & vbCrLf & _
        "IDs : {Ids}" & vbCrLf & "WBS : {Wbs}")
    definitions(7) = Array("DIAG.GROUP.DEPENDENCY.UNSUPPORTED_LINK", _
        "Link type not supported by the engine" & vbCrLf & "-> " & "fix the link type in Predecessors WBS or tbl_LOGIC_LINKS" & vbCrLf & vbCrLf & _
        "IDs: {Ids}" & vbCrLf & "WBS: {Wbs}", _
        "Type de lien non supporté par le moteur" & vbCrLf & "-> " & "corriger le type de lien dans Predecessors WBS ou tbl_LOGIC_LINKS" & vbCrLf & vbCrLf & _
        "IDs : {Ids}" & vbCrLf & "WBS : {Wbs}")
    definitions(8) = Array("DIAG.GROUP.DEPENDENCY.CYCLE", _
        "Dependency cycle detected" & vbCrLf & "-> " & "fix the Predecessors WBS column" & vbCrLf & vbCrLf & _
        "IDs: {Ids}" & vbCrLf & "WBS: {Wbs}", _
        "Boucle de dépendance détectée" & vbCrLf & "-> " & "corriger la colonne Predecessors WBS" & vbCrLf & vbCrLf & _
        "IDs : {Ids}" & vbCrLf & "WBS : {Wbs}")
    definitions(9) = Array("DIAG.GROUP.TEST.START_CONFLICT", _
        "Test Start is incompatible with upstream dependencies" & vbCrLf & "-> " & "fix Test Start or upstream logic" & vbCrLf & vbCrLf & _
        "IDs: {Ids}" & vbCrLf & "WBS: {Wbs}", _
        "Test Start incompatible avec les d" & ChrW$(233) & "pendances amont" & vbCrLf & "-> " & "corriger Test Start ou la logique amont" & vbCrLf & vbCrLf & _
        "IDs : {Ids}" & vbCrLf & "WBS : {Wbs}")
    definitions(10) = Array("DIAG.GROUP.FORECAST.START_CONFLICT", _
        "Forecast Start is incompatible with upstream dependencies" & vbCrLf & "-> " & "fix Forecast Start or upstream logic" & vbCrLf & vbCrLf & _
        "IDs: {Ids}" & vbCrLf & "WBS: {Wbs}", _
        "Forecast Start incompatible avec les d" & ChrW$(233) & "pendances amont" & vbCrLf & "-> " & "corriger Forecast Start ou la logique amont" & vbCrLf & vbCrLf & _
        "IDs : {Ids}" & vbCrLf & "WBS : {Wbs}")
    definitions(11) = Array("DIAG.GROUP.TEST.FINISH_CONFLICT", _
        "Test Finish is incompatible with upstream finish constraints" & vbCrLf & "-> " & "fix Test Finish or upstream logic" & vbCrLf & vbCrLf & _
        "IDs: {Ids}" & vbCrLf & "WBS: {Wbs}", _
        "Test Finish incompatible avec les contraintes de fin amont" & vbCrLf & "-> " & "corriger Test Finish ou la logique amont" & vbCrLf & vbCrLf & _
        "IDs : {Ids}" & vbCrLf & "WBS : {Wbs}")
    definitions(12) = Array("DIAG.GROUP.FORECAST.FINISH_CONFLICT", _
        "Forecast Finish is incompatible with upstream finish constraints" & vbCrLf & "-> " & "fix Forecast Finish or upstream logic" & vbCrLf & vbCrLf & _
        "IDs: {Ids}" & vbCrLf & "WBS: {Wbs}", _
        "Forecast Finish incompatible avec les contraintes de fin amont" & vbCrLf & "-> " & "corriger Forecast Finish ou la logique amont" & vbCrLf & vbCrLf & _
        "IDs : {Ids}" & vbCrLf & "WBS : {Wbs}")
    definitions(13) = Array("DIAG.GROUP.BASELINE.MISSING_DURATION", _
        "Missing Baseline Duration" & vbCrLf & "-> " & "please fill in Baseline Duration" & vbCrLf & vbCrLf & _
        "IDs: {Ids}" & vbCrLf & "WBS: {Wbs}", _
        "Baseline Duration manquante" & vbCrLf & "-> " & "compléter la durée baseline" & vbCrLf & vbCrLf & _
        "IDs : {Ids}" & vbCrLf & "WBS : {Wbs}")
    definitions(14) = Array("DIAG.GROUP.DATES.START_UNRESOLVED", _
        "Start date not computable" & vbCrLf & "-> " & "check dependencies or baseline" & vbCrLf & vbCrLf & _
        "IDs: {Ids}" & vbCrLf & "WBS: {Wbs}", _
        "Date de début non déterminable" & vbCrLf & "-> " & "vérifier les dépendances ou la baseline" & vbCrLf & vbCrLf & _
        "IDs : {Ids}" & vbCrLf & "WBS : {Wbs}")
    definitions(15) = Array("DIAG.GROUP.DATES.FINISH_BEFORE_START", _
        "Finish is incompatible with start" & vbCrLf & "-> " & "fix dates or duration" & vbCrLf & vbCrLf & _
        "IDs: {Ids}" & vbCrLf & "WBS: {Wbs}", _
        "Fin incompatible avec le d" & ChrW$(233) & "but" & vbCrLf & "-> " & "corriger les dates ou la dur" & ChrW$(233) & "e" & vbCrLf & vbCrLf & _
        "IDs : {Ids}" & vbCrLf & "WBS : {Wbs}")
    definitions(16) = Array("DIAG.GROUP.ENGINE.TEST_ERROR", _
        "Calculation error in live engine" & vbCrLf & "-> " & "fix test values or upstream logic" & vbCrLf & vbCrLf & _
        "IDs: {Ids}" & vbCrLf & "WBS: {Wbs}", _
        "Erreur de calcul dans le moteur live" & vbCrLf & "-> " & "corriger les valeurs test ou la logique amont" & vbCrLf & vbCrLf & _
        "IDs : {Ids}" & vbCrLf & "WBS : {Wbs}")
    definitions(17) = Array("DIAG.GROUP.ENGINE.SCENARIO_ERROR", _
        "Calculation error in scenario" & vbCrLf & "-> " & "fix test values or upstream logic" & vbCrLf & vbCrLf & _
        "IDs: {Ids}" & vbCrLf & "WBS: {Wbs}", _
        "Erreur de calcul dans le sc" & ChrW$(233) & "nario" & vbCrLf & "-> " & "corriger les valeurs de test ou la logique amont" & vbCrLf & vbCrLf & _
        "IDs : {Ids}" & vbCrLf & "WBS : {Wbs}")
    definitions(18) = Array("DIAG.GROUP.ENGINE.ERROR", _
        "Blocking error detected by the engine" & vbCrLf & "-> " & "check ErrorMsg in tbl_CALC for technical details" & vbCrLf & vbCrLf & _
        "IDs: {Ids}" & vbCrLf & "WBS: {Wbs}", _
        "Erreur bloquante d" & ChrW$(233) & "tect" & ChrW$(233) & "e par le moteur" & vbCrLf & "-> " & "v" & ChrW$(233) & "rifier ErrorMsg dans tbl_CALC pour le d" & ChrW$(233) & "tail technique" & vbCrLf & vbCrLf & _
        "IDs : {Ids}" & vbCrLf & "WBS : {Wbs}")
    definitions(19) = Array("DIAG.GROUP.DEPENDENCY.MISSING_LINK_PREDECESSOR", _
        "Missing predecessor in tbl_LOGIC_LINKS" & vbCrLf & "-> " & "check the logical links table" & vbCrLf & vbCrLf & _
        "IDs: {Ids}" & vbCrLf & "WBS: {Wbs}", _
        "Prédécesseur introuvable dans tbl_LOGIC_LINKS" & vbCrLf & "-> " & "vérifier la table des liens logiques" & vbCrLf & vbCrLf & _
        "IDs : {Ids}" & vbCrLf & "WBS : {Wbs}")
    definitions(20) = Array("DIAG.GROUP.DEADLINE.EXCEEDED", _
        "Deadline exceeded" & vbCrLf & "-> " & "calculated finish is after the deadline" & vbCrLf & vbCrLf & _
        "IDs: {Ids}" & vbCrLf & "WBS: {Wbs}", _
        "Deadline depassee" & vbCrLf & "-> " & "la date calculee finit apres la deadline" & vbCrLf & vbCrLf & _
        "IDs : {Ids}" & vbCrLf & "WBS : {Wbs}")
    definitions(21) = Array("DIAG.GROUP.SUMMARY.IGNORED_DATES", _
        "Dates entered on summary task" & vbCrLf & "-> " & "values are ignored and calculated from child tasks" & vbCrLf & vbCrLf & _
        "IDs: {Ids}" & vbCrLf & "WBS: {Wbs}", _
        "Dates saisies sur tâche parent" & vbCrLf & "-> " & "les valeurs sont ignorées, calcul par les tâches enfants" & vbCrLf & vbCrLf & _
        "IDs : {Ids}" & vbCrLf & "WBS : {Wbs}")
    definitions(22) = Array("DIAG.GROUP.TASK_TYPE.IGNORED_CALENDAR", _
        "Calendar ignored on Summary / LOE / Milestone task" & vbCrLf & "-> " & "calendars apply only to normal tasks" & vbCrLf & vbCrLf & _
        "IDs: {Ids}" & vbCrLf & "WBS: {Wbs}", _
        "Calendrier ignoré sur tâche Summary / LOE / Milestone" & vbCrLf & "-> " & "le calendrier n'est utilisé que pour les tâches normales" & vbCrLf & vbCrLf & _
        "IDs : {Ids}" & vbCrLf & "WBS : {Wbs}")
    definitions(23) = Array("DIAG.GROUP.LOE.IGNORED_PROGRESS", _
        "% Progress entered on LOE" & vbCrLf & "-> " & "LOE progress is automatically calculated from today's date; manual input is ignored in the Gantt" & vbCrLf & vbCrLf & _
        "IDs: {Ids}" & vbCrLf & "WBS: {Wbs}", _
        "% Progress renseigné sur LOE" & vbCrLf & "-> " & "le progress LOE est calculé automatiquement par date du jour ; la saisie manuelle est ignorée dans le Gantt" & vbCrLf & vbCrLf & _
        "IDs : {Ids}" & vbCrLf & "WBS : {Wbs}")
    definitions(24) = Array("DIAG.GROUP.LOE.BASELINE", _
        "Baseline entered on LOE" & vbCrLf & "-> " & "a LOE must be driven by its SS/FF links; check that entered baseline values are not misleading" & vbCrLf & vbCrLf & _
        "IDs: {Ids}" & vbCrLf & "WBS: {Wbs}", _
        "Baseline renseignée sur LOE" & vbCrLf & "-> " & "une LOE doit être pilotée par ses liens SS/FF ; vérifier que la baseline saisie ne crée pas de confusion" & vbCrLf & vbCrLf & _
        "IDs : {Ids}" & vbCrLf & "WBS : {Wbs}")
    definitions(25) = Array("DIAG.GROUP.MILESTONE.PARTIAL_PROGRESS", _
        "Partial % Progress entered on Milestone" & vbCrLf & "-> " & "a milestone must be either 0% or 100%; any intermediate value should be corrected" & vbCrLf & vbCrLf & _
        "IDs: {Ids}" & vbCrLf & "WBS: {Wbs}", _
        "% Progress partiel renseigné sur Milestone" & vbCrLf & "-> " & "une milestone doit être à 0% ou 100% ; toute valeur intermédiaire doit être corrigée" & vbCrLf & vbCrLf & _
        "IDs : {Ids}" & vbCrLf & "WBS : {Wbs}")
    definitions(26) = Array("DIAG.GROUP.MILESTONE.DURATION", _
        "Duration greater than 1 day on Milestone" & vbCrLf & "-> " & "the task will be rendered as a milestone but the entered duration may be misleading" & vbCrLf & vbCrLf & _
        "IDs: {Ids}" & vbCrLf & "WBS: {Wbs}", _
        "Durée supérieure à 1 jour sur Milestone" & vbCrLf & "-> " & "la tâche sera rendue comme milestone mais la durée saisie peut induire en erreur" & vbCrLf & vbCrLf & _
        "IDs : {Ids}" & vbCrLf & "WBS : {Wbs}")
    definitions(27) = Array("DIAG.PLANNING.NO_CHANGE", _
        "No change detected: core calculation was not rerun.", _
        "Aucune modification détectée : calcul moteur non relancé.")
    definitions(28) = Array("DIAG.PLANNING.COMPLETE", _
        "Calculation completed successfully.", "Calcul terminé avec succès.")
    definitions(29) = Array("DIAG.PLANNING.SYNC_INVALID", _
        "WBS -> CALC sync failed or left tbl_CALC in an invalid state." & vbCrLf & _
        "Calculation stopped to avoid a false success.", _
        "Le sync WBS -> CALC a échoué ou a laissé tbl_CALC dans un état invalide." & vbCrLf & _
        "Le calcul est arrêté pour éviter un faux succès.")
    definitions(30) = Array("DIAG.PLANNING.RUN_ERROR", _
        "Error in Run_Calc_Engine_CoreBridge" & vbCrLf & "-> {Details}", _
        "Erreur dans Run_Calc_Engine_CoreBridge" & vbCrLf & "-> {Details}")
    definitions(31) = Array("DIAG.PLANNING.SOURCE_DIAGNOSTIC_FAILED", _
        "Calculation stopped: the engine detected blocking errors." & vbCrLf & _
        "Unable to rebuild the detailed message." & vbCrLf & _
        "-> check Error flag and ErrorMsg columns in tbl_CALC." & vbCrLf & _
        "-> no calculated data was pushed back to WBS.", _
        "Calcul arrete : le moteur a detecte des erreurs bloquantes." & vbCrLf & _
        "Impossible de reconstruire le message detaille." & vbCrLf & _
        "-> verifier les colonnes Error flag et ErrorMsg dans tbl_CALC." & vbCrLf & _
        "-> aucune donnee calculee n'a ete repoussee vers WBS.")
    definitions(32) = Array("DIAG.PLANNING.DATA_DIAGNOSTIC_FAILED", _
        "Calculation stopped: the engine detected blocking errors." & vbCrLf & _
        "Unable to rebuild the detailed message." & vbCrLf & _
        "-> check Error flag and ErrorMsg columns in tbl_CALC." & vbCrLf & _
        "-> no calculated data was pushed back to WBS.", _
        "Calcul arrêté : le moteur a détecté des erreurs bloquantes." & vbCrLf & _
        "Impossible de reconstruire le message détaillé." & vbCrLf & _
        "-> vérifier les colonnes Error flag et ErrorMsg dans tbl_CALC." & vbCrLf & _
        "-> aucune donnée calculée n'a été repoussée vers WBS.")
    definitions(33) = Array("DIAG.PRECHECK.LOE_FAILED", _
        "Error while checking LOE used as predecessor." & vbCrLf & "-> calculation stopped before WBS write.", _
        "Erreur pendant le controle des LOE utilisees comme predecesseur." & vbCrLf & "-> calcul arrete avant ecriture WBS.")
    definitions(34) = Array("DIAG.PRECHECK.CYCLES_FAILED", _
        "Error while checking dependency cycles." & vbCrLf & "-> calculation stopped before WBS write.", _
        "Erreur pendant le contrôle des cycles." & vbCrLf & "-> calcul arrêté avant écriture WBS.")
    definitions(35) = Array("DIAG.PRECHECK.PREDECESSORS_FAILED", _
        "Error while checking predecessors." & vbCrLf & "-> calculation stopped before WBS write.", _
        "Erreur pendant le contrôle des prédécesseurs." & vbCrLf & "-> calcul arrêté avant écriture WBS.")
    definitions(36) = Array("DIAG.ANALYTICS.TOPOLOGY_INCOMPLETE", _
        "Analytics not calculated: incomplete topological order." & vbCrLf & "-> check cycles or tbl_LOGIC_LINKS rebuild.", _
        "Analytics non calculées : ordre topologique incomplet." & vbCrLf & "-> vérifier les cycles ou la reconstruction tbl_LOGIC_LINKS.")
    definitions(37) = Array("DIAG.INFRASTRUCTURE.ERROR", _
        "Error Ensure_Calc_Infrastructure" & vbCrLf & "-> " & "{Details}", _
        "Erreur Ensure_Calc_Infrastructure" & vbCrLf & "-> " & "{Details}")
    definitions(38) = Array("DIAG.VARIANCES.ERROR", _
        "Error in Compute_And_Push_Variances" & vbCrLf & "-> " & "{Details}", _
        "Erreur dans Compute_And_Push_Variances" & vbCrLf & "-> " & "{Details}")
    definitions(39) = Array("DIAG.UPSTREAM.ACTUAL_START_CONFLICT", _
        "Actual Start is incompatible with upstream dependencies" & vbCrLf & _
        "-> fix Actual Start, upstream logic, or lag" & vbCrLf & vbCrLf & _
        "Tasks: {Items}", _
        "Actual Start incompatible avec les dépendances amont" & vbCrLf & _
        "-> corriger Actual Start, la logique amont ou le lag" & vbCrLf & vbCrLf & _
        "Tâches : {Items}")
    definitions(40) = Array("DIAG.UPSTREAM.ACTUAL_FINISH_CONFLICT", _
        "Actual Finish is incompatible with upstream finish constraints" & vbCrLf & _
        "-> fix Actual Finish, upstream logic, or lag" & vbCrLf & vbCrLf & _
        "Tasks: {Items}", _
        "Actual Finish incompatible avec les contraintes de fin amont" & vbCrLf & _
        "-> corriger Actual Finish, la logique amont ou le lag" & vbCrLf & vbCrLf & _
        "Tâches : {Items}")
    definitions(41) = Array("DIAG.ANALYTICS.EMPTY", _
        "Analytics not recalculated: tbl_CALC is empty.", _
        "Analytics non recalculées : tbl_CALC est vide.")
    definitions(42) = Array("DIAG.ANALYTICS.RUN_ERROR", _
        "Error in Run_Analytics_Only" & vbCrLf & "-> {Details}", _
        "Erreur dans Run_Analytics_Only" & vbCrLf & "-> {Details}")
    definitions(43) = Array("DIAG.ANALYTICS.CLEAR_ERROR", _
        "Error in Clear_Analytics_Outputs" & vbCrLf & "-> {Details}", _
        "Erreur dans Clear_Analytics_Outputs" & vbCrLf & "-> {Details}")
    definitions(44) = Array("DIAG.ANALYTICS.NEGATIVE_FLOAT", _
        "Negative float detected in the current schedule" & vbCrLf & _
        "-> check logic, dates, lags or forecasts", _
        "Float négatif détecté dans le planning actuel" & vbCrLf & _
        "-> vérifier la logique, les dates, les lags ou les prévisions")
    definitions(45) = Array("DIAG.CALC_ENGINE.RUN_ERROR", _
        "Error in Run_Calc_Engine" & vbCrLf & "-> {Details}", _
        "Erreur dans Run_Calc_Engine" & vbCrLf & "-> {Details}")
    definitions(46) = Array("DIAG.CALC_ENGINE.WBS_EMPTY", _
        "tbl_WBS is empty.", "tbl_WBS est vide.")
    definitions(47) = Array("DIAG.CALC_ENGINE.NETWORK_OK", _
        "Network validation OK." & vbCrLf & _
        "-> no missing predecessor, no cycle, no non-positionable task detected.", _
        "Validation réseau OK." & vbCrLf & _
        "-> aucun prédécesseur manquant, aucun cycle, aucune tâche non positionnable détectée.")
    definitions(48) = Array("DIAG.CALC_ENGINE.NETWORK_ERROR", _
        "VBA error in Validate_LogicLinksNetwork: {Details}", _
        "Erreur VBA dans Validate_LogicLinksNetwork : {Details}")
    definitions(49) = Array("DIAG.EVENT_HISTORY.LOG_ERROR", _
        "The message history could not be saved. Close and reopen the workbook before continuing.", _
        "L'historique des messages n'a pas pu être enregistré. Fermez puis rouvrez le classeur avant de continuer.")
    definitions(50) = Array("DIAG.SCURVE.UPDATE_ERROR", _
        "Error in Run_SCurve_Update" & vbCrLf & "-> {Details}", _
        "Erreur dans Run_SCurve_Update" & vbCrLf & "-> {Details}")
    definitions(51) = Array("DIAG.CALC_STATE.WRITE_ERROR", _
        "Error in Write_CalcState_Snapshot: {Details}", _
        "Erreur dans Write_CalcState_Snapshot : {Details}")
    definitions(52) = Array("DIAG.OUTPUT.WBS_ID_MISSING", _
        "Column ID was not found in tbl_WBS.", _
        "La colonne ID est introuvable dans tbl_WBS.")
    definitions(53) = Array("DIAG.OUTPUT.CALC_ID_MISSING", _
        "Column ID was not found in tbl_CALC.", _
        "La colonne ID est introuvable dans tbl_CALC.")
    definitions(54) = Array("DIAG.OUTPUT.WBS_COLUMN_MISSING", _
        "Output column not found in tbl_WBS: {Column}", _
        "Colonne de sortie introuvable dans tbl_WBS : {Column}")
    definitions(55) = Array("DIAG.OUTPUT.CALC_COLUMN_MISSING", _
        "Output column not found in tbl_CALC: {Column}", _
        "Colonne de sortie introuvable dans tbl_CALC : {Column}")
    definitions(56) = Array("DIAG.OUTPUT.PUSH_ERROR", _
        "Error in Push_Calculated_Back_To_WBS: {Details}", _
        "Erreur dans Push_Calculated_Back_To_WBS : {Details}")
    definitions(57) = Array("DIAG.WBS_FORMULAS.RESTORE_ERROR", _
        "Error in RestoreWBSFormulaColumns: {Details}", _
        "Erreur dans RestoreWBSFormulaColumns : {Details}")
    definitions(58) = Array("DIAG.DATASYNC.WBS_ID_MISSING", "Column ID was not found in tbl_WBS.", "La colonne ID est introuvable dans tbl_WBS.")
    definitions(59) = Array("DIAG.DATASYNC.WBS_WBS_MISSING", "Column WBS was not found in tbl_WBS.", "La colonne WBS est introuvable dans tbl_WBS.")
    definitions(60) = Array("DIAG.DATASYNC.WBS_TASK_TYPE_MISSING", "Column Task Type was not found in tbl_WBS.", "La colonne Task Type est introuvable dans tbl_WBS.")
    definitions(61) = Array("DIAG.DATASYNC.WBS_S_MISSING", "Column S was not found in tbl_WBS.", "La colonne S est introuvable dans tbl_WBS.")
    definitions(62) = Array("DIAG.DATASYNC.WBS_PREDECESSORS_MISSING", "Column Predecessors WBS was not found in WBS.", "La colonne Predecessors WBS est introuvable dans tbl_WBS.")
    definitions(63) = Array("DIAG.DATASYNC.CALC_ID_MISSING", "Column ID was not found in tbl_CALC.", "La colonne ID est introuvable dans tbl_CALC.")
    definitions(64) = Array("DIAG.DATASYNC.CALC_WBS_MISSING", "Column WBS was not found in tbl_CALC.", "La colonne WBS est introuvable dans tbl_CALC.")
    definitions(65) = Array("DIAG.DATASYNC.CALC_TASK_TYPE_MISSING", "Column Task Type was not found in tbl_CALC.", "La colonne Task Type est introuvable dans tbl_CALC.")
    definitions(66) = Array("DIAG.DATASYNC.CALC_S_MISSING", "Column S was not found in tbl_CALC.", "La colonne S est introuvable dans tbl_CALC.")
    definitions(67) = Array("DIAG.DATASYNC.CALC_PREDECESSORS_MISSING", "Column Predecessors WBS was not found in tbl_CALC.", "La colonne Predecessors WBS est introuvable dans tbl_CALC.")
    definitions(68) = Array("DIAG.DATASYNC.SYNC_ERROR", "Error in Sync_WBS_To_CALC: {Details}", "Erreur dans Sync_WBS_To_CALC : {Details}")
    definitions(69) = Array("DIAG.DATASYNC.NON_NUMERIC_ID", "Non-numeric ID detected in WBS: {Value}", "ID non numerique detecte dans WBS : {Value}")
    definitions(70) = Array("DIAG.DATASYNC.INVALID_ID", "Invalid ID in WBS (must be >= 1): {Value}", "ID invalide dans WBS (doit etre >= 1) : {Value}")
    definitions(71) = Array("DIAG.DATASYNC.DUPLICATE_WBS", "Duplicate WBS detected in WBS: {Value}", "WBS duplique detecte dans WBS : {Value}")
    definitions(72) = Array("DIAG.DATASYNC.WBS_FORECAST_START_MISSING", "Column Forecast Start was not found in tbl_WBS.", "La colonne Forecast Start est introuvable dans tbl_WBS.")
    definitions(73) = Array("DIAG.DATASYNC.WBS_FORECAST_FINISH_MISSING", "Column Forecast Finish was not found in tbl_WBS.", "La colonne Forecast Finish est introuvable dans tbl_WBS.")
    definitions(74) = Array("DIAG.DATASYNC.CALC_FORECAST_START_MISSING", "Column Forecast Start was not found in tbl_CALC.", "La colonne Forecast Start est introuvable dans tbl_CALC.")
    definitions(75) = Array("DIAG.DATASYNC.CALC_FORECAST_FINISH_MISSING", "Column Forecast Finish was not found in tbl_CALC.", "La colonne Forecast Finish est introuvable dans tbl_CALC.")
    definitions(76) = Array("DIAG.DATASYNC.FORECAST_SYNC_ERROR", "Error in Sync_Forecast_Only: {Details}", "Erreur dans Sync_Forecast_Only : {Details}")
    definitions(77) = Array("DIAG.DATASYNC.LOGIC_LINKS_REBUILD_ERROR", "Error while rebuilding tbl_LOGIC_LINKS." & vbCrLf & "-> {Details}", "Erreur lors de la reconstruction de tbl_LOGIC_LINKS." & vbCrLf & "-> {Details}")
    definitions(78) = Array("DIAG.DATASYNC.MISSING_PREDECESSOR", _
        "Missing predecessor" & vbCrLf & vbCrLf & "The following predecessors do not exist in the planning." & vbCrLf & vbCrLf & _
        "IDs:" & vbCrLf & "{Ids}" & vbCrLf & vbCrLf & "WBS:" & vbCrLf & "{Wbs}" & vbCrLf & vbCrLf & _
        "Entered predecessors:" & vbCrLf & "{Entered}" & vbCrLf & vbCrLf & "Referenced WBS:" & vbCrLf & "{Referenced}" & vbCrLf & vbCrLf & _
        "-> create the missing tasks or correct the links.", _
        "Prédécesseur introuvable" & vbCrLf & vbCrLf & "Les prédécesseurs suivants n'existent pas dans le planning." & vbCrLf & vbCrLf & _
        "IDs :" & vbCrLf & "{Ids}" & vbCrLf & vbCrLf & "WBS :" & vbCrLf & "{Wbs}" & vbCrLf & vbCrLf & _
        "Prédécesseurs saisis :" & vbCrLf & "{Entered}" & vbCrLf & vbCrLf & "WBS recherchées :" & vbCrLf & "{Referenced}" & vbCrLf & vbCrLf & _
        "-> créer les tâches manquantes ou corriger les liens.")
    definitions(79) = Array("DIAG.DATASYNC.LOGIC_LINKS_VBA_ERROR", "VBA error in RebuildLogicLinksTable: {Details}", "Erreur VBA dans RebuildLogicLinksTable : {Details}")
    definitions(80) = Array("DIAG.CALC_ENGINE.TOPO_CYCLE_DETAIL", _
        "Logic cycle detected:" & vbCrLf & vbCrLf & "{Cycle}" & vbCrLf & vbCrLf & _
        "-> fix one of the links in this loop in Predecessors WBS.", _
        "Cycle logique détecté :" & vbCrLf & vbCrLf & "{Cycle}" & vbCrLf & vbCrLf & _
        "-> corriger un des liens de cette boucle dans Predecessors WBS.")
    definitions(81) = Array("DIAG.CALC_STATE.ENSURE_ERROR", _
        "Error in Ensure_CalcState_Table: {Details}", _
        "Erreur dans Ensure_CalcState_Table : {Details}")
    definitions(82) = Array("DIAG.CORE_PILOT.CALC_EMPTY", "Table tbl_CALC is empty.", "La table tbl_CALC est vide.")
    definitions(83) = Array("DIAG.CORE_PILOT.WBS_EMPTY", "Table tbl_WBS is empty.", "La table tbl_WBS est vide.")
    definitions(84) = Array("DIAG.CORE_PILOT.COMPLETE", "Pilot core completed." & vbCrLf & "-> results written to tbl_CALC", "Pilot core terminé." & vbCrLf & "-> résultats écrits dans tbl_CALC")
    definitions(85) = Array("DIAG.CORE_PILOT.ERROR", "Error in Run_Calc_Core_PROD_Pilot" & vbCrLf & "-> {Details}", "Erreur dans Run_Calc_Core_PROD_Pilot" & vbCrLf & "-> {Details}")
    definitions(86) = Array("DIAG.CONSTRAINT.UNKNOWN_START_TYPE", "Unknown start constraint type", "Type de contrainte debut non reconnu")
    definitions(87) = Array("DIAG.CONSTRAINT.UNKNOWN_FINISH_TYPE", "Unknown finish constraint type", "Type de contrainte fin non reconnu")
    definitions(88) = Array("DIAG.CONSTRAINT.ACTUAL_START_BEFORE_START", "Actual Start is before start constraint", "Actual Start avant contrainte debut")
    definitions(89) = Array("DIAG.CONSTRAINT.ACTUAL_START_AFTER_LATEST", "Actual Start is after latest start constraint", "Actual Start apres contrainte debut max")
    definitions(90) = Array("DIAG.CONSTRAINT.ACTUAL_START_DIFFERS_MSO", "Actual Start differs from Must Start On constraint", "Actual Start different de contrainte Must Start On")
    definitions(91) = Array("DIAG.CONSTRAINT.ACTUAL_FINISH_BEFORE_FINISH", "Actual Finish is before finish constraint", "Actual Finish avant contrainte fin")
    definitions(92) = Array("DIAG.CONSTRAINT.ACTUAL_FINISH_AFTER_LATEST", "Actual Finish is after latest finish constraint", "Actual Finish apres contrainte fin max")
    definitions(93) = Array("DIAG.CONSTRAINT.ACTUAL_FINISH_DIFFERS_MFO", "Actual Finish differs from Must Finish On constraint", "Actual Finish different de contrainte Must Finish On")
    definitions(94) = Array("DIAG.CONSTRAINT.FORECAST_START_BEFORE_START", "Forecast Start is before start constraint", "Forecast Start avant contrainte debut")
    definitions(95) = Array("DIAG.CONSTRAINT.FORECAST_START_AFTER_LATEST", "Forecast Start is after latest start constraint", "Forecast Start apres contrainte debut max")
    definitions(96) = Array("DIAG.CONSTRAINT.FORECAST_START_DIFFERS_MSO", "Forecast Start differs from Must Start On constraint", "Forecast Start different de contrainte Must Start On")
    definitions(97) = Array("DIAG.CONSTRAINT.FORECAST_FINISH_BEFORE_UPSTREAM", "Forecast Finish is before upstream finish constraint", "Forecast Finish avant contrainte fin amont")
    definitions(98) = Array("DIAG.CONSTRAINT.FORECAST_FINISH_BEFORE_FINISH", "Forecast Finish is before finish constraint", "Forecast Finish avant contrainte fin")
    definitions(99) = Array("DIAG.CONSTRAINT.FORECAST_FINISH_AFTER_LATEST", "Forecast Finish is after latest finish constraint", "Forecast Finish apres contrainte fin max")
    definitions(100) = Array("DIAG.CONSTRAINT.FORECAST_FINISH_DIFFERS_MFO", "Forecast Finish differs from Must Finish On constraint", "Forecast Finish different de contrainte Must Finish On")
    definitions(101) = Array("DIAG.CONSTRAINT.CALCULATED_FINISH_DIFFERS_MFO", "Calculated Finish differs from Must Finish On constraint", "Calculated Finish different de contrainte Must Finish On")
    definitions(102) = Array("DIAG.CONSTRAINT.CALCULATED_START_BEFORE_MFO_NETWORK", "Calculated Start is before network allowed start due to Must Finish On constraint", "Calculated Start avant reseau impose par contrainte Must Finish On")
    definitions(103) = Array("DIAG.CONSTRAINT.DURATION_INCOMPATIBLE_MSO_MFO", "Duration is incompatible with Must Start On / Must Finish On constraints", "Duree incompatible avec contraintes Must Start On / Must Finish On")
    definitions(104) = Array("DIAG.CONSTRAINT.ACTUAL_START_DIFFERS_MFO_IMPLIED", "Actual Start differs from start implied by Must Finish On constraint", "Actual Start different du debut impose par Must Finish On")
    definitions(105) = Array("DIAG.CONSTRAINT.FORECAST_START_DIFFERS_MFO_IMPLIED", "Forecast Start differs from start implied by Must Finish On constraint", "Forecast Start different du debut impose par Must Finish On")
    definitions(106) = Array("DIAG.CONSTRAINT.CALCULATED_START_DIFFERS_MSO", "Calculated Start differs from Must Start On constraint", "Calculated Start different de contrainte Must Start On")
    definitions(107) = Array("DIAG.CONSTRAINT.CALCULATED_START_AFTER_LATEST", "Calculated Start is after latest start constraint", "Calculated Start apres contrainte debut max")
    definitions(108) = Array("DIAG.CONSTRAINT.CALCULATED_FINISH_AFTER_LATEST", "Calculated Finish is after latest finish constraint", "Calculated Finish apres contrainte fin max")
    definitions(109) = Array("DIAG.CALC_ENGINE.LINKS_MISSING_PREDECESSOR_WBS", _
        "Missing predecessor in tbl_LOGIC_LINKS" & vbCrLf & "-> check the Predecessors WBS column" & vbCrLf & vbCrLf & "IDs: {Ids}" & vbCrLf & "WBS: {Wbs}", _
        "Prédécesseur introuvable dans tbl_LOGIC_LINKS" & vbCrLf & "-> vérifier la colonne Predecessors WBS" & vbCrLf & vbCrLf & "IDs : {Ids}" & vbCrLf & "WBS : {Wbs}")
    definitions(110) = Array("DIAG.CALC_ENGINE.LOE_PREDECESSOR", _
        "A Level of Effort cannot be used as predecessor" & vbCrLf & "-> fix the logical relationship" & vbCrLf & vbCrLf & "IDs: {Ids}" & vbCrLf & "WBS: {Wbs}", _
        "Un Level of Effort ne peut pas être prédécesseur" & vbCrLf & "-> corriger la logique de liaison" & vbCrLf & vbCrLf & "IDs : {Ids}" & vbCrLf & "WBS : {Wbs}")
    definitions(111) = Array("DIAG.CALC_ENGINE.LOGICAL_CYCLE", _
        "Logical cycle detected" & vbCrLf & "-> fix the dependency relationships" & vbCrLf & vbCrLf & "IDs: {Ids}" & vbCrLf & "WBS: {Wbs}", _
        "Boucle logique détectée" & vbCrLf & "-> corriger les relations de dépendance" & vbCrLf & vbCrLf & "IDs : {Ids}" & vbCrLf & "WBS : {Wbs}")
    definitions(112) = Array("DIAG.CALC_ENGINE.NOT_POSITIONABLE", _
        "Task or chain cannot be positioned" & vbCrLf & "-> add a start date or an anchored upstream logic" & vbCrLf & vbCrLf & "IDs: {Ids}" & vbCrLf & "WBS: {Wbs}", _
        "Tâche ou chaîne non positionnable" & vbCrLf & "-> ajouter une date de début ou une logique amont ancrée" & vbCrLf & vbCrLf & "IDs : {Ids}" & vbCrLf & "WBS : {Wbs}")
    definitions(113) = Array("DIAG.CALC_ENGINE.LINKS_MISSING_PREDECESSOR_TABLE", _
        "Missing predecessor in tbl_LOGIC_LINKS" & vbCrLf & "-> check the logical links table" & vbCrLf & vbCrLf & "IDs: {Ids}" & vbCrLf & "WBS: {Wbs}", _
        "Prédécesseur introuvable dans tbl_LOGIC_LINKS" & vbCrLf & "-> vérifier la table des liens logiques" & vbCrLf & vbCrLf & "IDs : {Ids}" & vbCrLf & "WBS : {Wbs}")
    definitions(114) = Array("DIAG.CALC_ENGINE.UNSUPPORTED_LINK_TYPE", _
        "Link type not yet supported by the engine" & vbCrLf & "-> at this stage, only FS links are calculated" & vbCrLf & vbCrLf & "IDs: {Ids}" & vbCrLf & "WBS: {Wbs}", _
        "Type de lien non encore supporté par le moteur" & vbCrLf & "-> à ce stade, seuls les liens FS sont calculés" & vbCrLf & vbCrLf & "IDs : {Ids}" & vbCrLf & "WBS : {Wbs}")
    definitions(115) = Array("DIAG.RUN_BUTTONS.GANTT_UPDATE_ERROR", _
        "The Gantt could not be updated and remains unavailable." & vbCrLf & "Retry the action. If the problem persists, send the workbook to support.", _
        "Le Gantt n'a pas pu être mis à jour et reste indisponible." & vbCrLf & "Réessayez l'action. Si le problème persiste, transmettez le classeur au support.")
    definitions(116) = Array("DIAG.RUN_BUTTONS.VBA_ERROR", _
        "VBA error in {Procedure}" & vbCrLf & "-> check the last edited block in mod_RunButtons", _
        "Erreur VBA dans {Procedure}" & vbCrLf & "-> vérifier le dernier bloc modifié dans mod_RunButtons")
    definitions(117) = Array("DIAG.CALC_ENGINE.UNAUTHORIZED_WRITE_TEST", "KO test: no abort request was detected.", "KO test : aucune demande d'abort n'a été détectée.")
    definitions(118) = Array("DIAG.ANALYTICS.MISSING_BASELINE_REX", _
        "REX analytics partially not calculated: the Baseline temporal state is incomplete on at least one leaf task." & vbCrLf & _
        "-> complete an explicit Baseline or a usable Baseline dependency chain on the lines listed below." & vbCrLf & vbCrLf & _
        "IDs: {Ids}" & vbCrLf & "WBS: {Wbs}", _
        "Analytics REX partiellement non calculées : état Baseline temporel incomplet sur au moins une tâche feuille." & vbCrLf & _
        "-> compléter une Baseline explicite ou une chaîne de dépendances Baseline exploitable sur les lignes listées ci-dessous." & vbCrLf & vbCrLf & _
        "IDs : {Ids}" & vbCrLf & "WBS : {Wbs}")
    definitions(119) = Array("DIAG.CONSTRAINT.STRUCTURED_DETAIL", _
        "Constraint cannot be met." & vbCrLf & vbCrLf & "Task: {Task}" & vbCrLf & "Constraint: {ConstraintType}" & vbCrLf & _
        "Constraint date: {ConstraintDate}" & vbCrLf & "{CheckedField}: {CheckedValue}" & vbCrLf & _
        "Expected value: {Relation} {AllowedValue}" & vbCrLf & "Calculated Start: {CalculatedStart}" & vbCrLf & _
        "Calculated Finish: {CalculatedFinish}{Cascade}", _
        "Contrainte impossible a respecter." & vbCrLf & vbCrLf & "Tache : {Task}" & vbCrLf & "Contrainte : {ConstraintType}" & vbCrLf & _
        "Date contrainte : {ConstraintDate}" & vbCrLf & "{CheckedField} : {CheckedValue}" & vbCrLf & _
        "Valeur attendue : {Relation} {AllowedValue}" & vbCrLf & "Calculated Start : {CalculatedStart}" & vbCrLf & _
        "Calculated Finish : {CalculatedFinish}{Cascade}")
    definitions(120) = Array("DIAG.CONSTRAINT.CASCADE", _
        vbCrLf & vbCrLf & "Propagation:" & vbCrLf & "Blocked by: {Parent}" & vbCrLf & "Root cause: {Root}", _
        vbCrLf & vbCrLf & "Propagation :" & vbCrLf & "Bloque par : {Parent}" & vbCrLf & "Cause racine : {Root}")
    definitions(121) = Array("DIAG.FORECAST_START.DEPENDENCY", _
        "Test Start impossible." & vbCrLf & vbCrLf & "Task: {Task}" & vbCrLf & _
        "Blocking dependency: {Predecessor} ({Link})" & vbCrLf & "{PredecessorDateLabel}: {PredecessorDate}" & vbCrLf & _
        "Earliest allowed start: {MinimumStart}" & vbCrLf & "{RequestedLabel}: {RequestedStart}" & vbCrLf & _
        "-> move the predecessor earlier, update the link/lag, or choose a Start >= {MinimumStart}.", _
        "Test Start impossible." & vbCrLf & vbCrLf & "Tache : {Task}" & vbCrLf & _
        "Dependance bloquante : {Predecessor} ({Link})" & vbCrLf & "{PredecessorDateLabel} : {PredecessorDate}" & vbCrLf & _
        "Debut minimum autorise : {MinimumStart}" & vbCrLf & "{RequestedLabel} : {RequestedStart}" & vbCrLf & _
        "-> avancer le predecesseur, modifier le lien/lag, ou choisir un Start >= {MinimumStart}.")
    definitions(122) = Array("DIAG.LOE.EXPLICIT.DETAIL_BLOCK", _
        "Task {SuccWbs} directly references LOE {LoeWbs}." & vbCrLf & vbCrLf & "Details:" & vbCrLf & vbCrLf & _
        "Successor:" & vbCrLf & "{SuccWbs} (ID {SuccId})" & vbCrLf & vbCrLf & _
        "Detected LOE:" & vbCrLf & "{LoeWbs} (ID {LoeId})" & vbCrLf & vbCrLf & "Link:" & vbCrLf & "{Link}", _
        "La tache {SuccWbs} reference directement la LOE {LoeWbs}." & vbCrLf & vbCrLf & "Details :" & vbCrLf & vbCrLf & _
        "Successeur :" & vbCrLf & "{SuccWbs} (ID {SuccId})" & vbCrLf & vbCrLf & _
        "LOE detectee :" & vbCrLf & "{LoeWbs} (ID {LoeId})" & vbCrLf & vbCrLf & "Lien :" & vbCrLf & "{Link}")
    definitions(123) = Array("DIAG.LOE.PARENT.DETAIL_BLOCK", _
        "Task {SuccWbs} references parent {ParentWbs}." & vbCrLf & "This parent contains LOE {LoeWbs}." & vbCrLf & vbCrLf & _
        "When the parent link is expanded, the LOE becomes an indirect predecessor." & vbCrLf & vbCrLf & "Details:" & vbCrLf & vbCrLf & _
        "Successor:" & vbCrLf & "{SuccWbs} (ID {SuccId})" & vbCrLf & vbCrLf & _
        "Entered predecessor:" & vbCrLf & "{ParentWbs} (ID {ParentId})" & vbCrLf & vbCrLf & _
        "Detected LOE:" & vbCrLf & "{LoeWbs} (ID {LoeId})" & vbCrLf & vbCrLf & "Link:" & vbCrLf & "{Link}", _
        "La tache {SuccWbs} reference le parent {ParentWbs}." & vbCrLf & "Ce parent contient la LOE {LoeWbs}." & vbCrLf & vbCrLf & _
        "Lors de l'expansion du lien parent, la LOE devient predecesseur indirect." & vbCrLf & vbCrLf & "Details :" & vbCrLf & vbCrLf & _
        "Successeur :" & vbCrLf & "{SuccWbs} (ID {SuccId})" & vbCrLf & vbCrLf & _
        "Predecesseur saisi :" & vbCrLf & "{ParentWbs} (ID {ParentId})" & vbCrLf & vbCrLf & _
        "LOE detectee :" & vbCrLf & "{LoeWbs} (ID {LoeId})" & vbCrLf & vbCrLf & "Lien :" & vbCrLf & "{Link}")
    definitions(124) = Array("DIAG.LABEL.PREDECESSOR_START", "Predecessor start", "Debut predecesseur")
    definitions(125) = Array("DIAG.LABEL.PREDECESSOR_FINISH", "Predecessor finish", "Fin predecesseur")
    definitions(126) = Array("DIAG.LABEL.REQUESTED_SCENARIO_START", "Requested Scenario Start", "Scenario Start demande")
    definitions(127) = Array("DIAG.LABEL.REQUESTED_TEST_START", "Requested Test Start", "Test Start demande")
    definitions(128) = Array("CONSOLE.WINDOW.TITLE", "Planning console", "Console de planification")
    definitions(129) = Array("CONSOLE.SECTION.MESSAGES", "Messages", "Messages")
    definitions(130) = Array("CONSOLE.COMMAND.PREVIOUS", "Previous", "Précédent")
    definitions(131) = Array("CONSOLE.COMMAND.NEXT", "Next", "Suivant")
    definitions(132) = Array("CONSOLE.COMMAND.CLOSE", "Close", "Fermer")
    definitions(133) = Array("CONSOLE.COMMAND.HIDE_WARNING", "Hide", "Cacher")
    definitions(134) = Array("EVENT_HISTORY.COMMAND.CLEAR_HISTORY", "Clear History", "Nettoyer historique")
    definitions(135) = Array("EVENT_HISTORY.COMMAND.CLEAR_ACK", "Clear list", "Nettoyer cache")
    definitions(136) = Array("CONSOLE.EMPTY.COUNTER", "STOP 0/0 | WARNING 0/0 | INFO 0/0", "ARRÊT 0/0 | AVERTISSEMENT 0/0 | INFO 0/0")
    definitions(137) = Array("CONSOLE.SEVERITY.STOP", "STOP", "ARRÊT")
    definitions(138) = Array("CONSOLE.SEVERITY.WARNING", "WARNING", "AVERTISSEMENT")
    definitions(139) = Array("CONSOLE.SEVERITY.INFO", "INFO", "INFO")
    definitions(140) = Array("CONSOLE.SEVERITY.ERROR", "ERROR", "ERREUR")
    definitions(141) = Array("CONSOLE.COUNTER.INFO_ACTIVE", "INFO {Index}/{Total}", "INFO {Index}/{Total}")
    definitions(142) = Array("CONSOLE.COUNTER.INFO_MUTED", "INFO Muted", "INFO masqué")
    definitions(143) = Array("CONSOLE.COUNTER.ALL", _
        "STOP {StopIndex}/{StopTotal} | WARNING {WarningIndex}/{WarningTotal} | {InfoCaption}", _
        "ARRÊT {StopIndex}/{StopTotal} | AVERTISSEMENT {WarningIndex}/{WarningTotal} | {InfoCaption}")
    definitions(144) = Array("DIAG.COMMON.CONTEXT", _
        "ID: {Id}" & vbCrLf & "WBS: {Wbs}" & vbCrLf & "Task: {Task}", _
        "ID : {Id}" & vbCrLf & "WBS : {Wbs}" & vbCrLf & "Task : {Task}")
    definitions(145) = Array("DIAG.CALC_ENGINE.CYCLE_MARKER", "Cycle detected", "Cycle detected")
    definitions(146) = Array("DIAG.CORE_ERROR_SUMMARY.ROW", "- ID {Id}{WbsPart}{DetailsPart}", "- ID {Id}{WbsPart}{DetailsPart}")
    definitions(147) = Array("DIAG.CORE_ERROR_SUMMARY.WBS_PART", " / WBS {Wbs}", " / WBS {Wbs}")
    definitions(148) = Array("DIAG.CORE_ERROR_SUMMARY.NO_DETAILS", "Error flag detected, but no detailed message was available.", "Error flag detected, but no detailed message was available.")
    definitions(149) = Array("DIAG.CORE_ERROR_SUMMARY.BUILD_FAILED", "Unable to build core error summary.", "Unable to build core error summary.")
    definitions(150) = Array("DIAG.CONSTRAINT.PM.UNKNOWN_START_TYPE", "Invalid start constraint type" & vbCrLf & "-> choose a supported start constraint", "Type de contrainte début invalide" & vbCrLf & "-> choisir une contrainte début supportée")
    definitions(151) = Array("DIAG.CONSTRAINT.PM.UNKNOWN_FINISH_TYPE", "Invalid finish constraint type" & vbCrLf & "-> choose a supported finish constraint", "Type de contrainte fin invalide" & vbCrLf & "-> choisir une contrainte fin supportée")
    definitions(152) = Array("DIAG.CONSTRAINT.PM.ACTUAL_START_BEFORE_START", "Actual Start is incompatible with Start No Earlier Than" & vbCrLf & "-> move Actual Start later or update the constraint", "Actual Start incompatible avec Start No Earlier Than" & vbCrLf & "-> repousser Actual Start ou modifier la contrainte")
    definitions(153) = Array("DIAG.CONSTRAINT.PM.FORECAST_START_BEFORE_START", "Forecast Start is incompatible with Start No Earlier Than" & vbCrLf & "-> move Forecast Start later or update the constraint", "Forecast Start incompatible avec Start No Earlier Than" & vbCrLf & "-> repousser Forecast Start ou modifier la contrainte")
    definitions(154) = Array("DIAG.CONSTRAINT.PM.ACTUAL_START_AFTER_LATEST", "Actual Start is incompatible with Start No Later Than" & vbCrLf & "-> align Actual Start or update the constraint", "Actual Start incompatible avec Start No Later Than" & vbCrLf & "-> aligner Actual Start ou modifier la contrainte")
    definitions(155) = Array("DIAG.CONSTRAINT.PM.FORECAST_START_AFTER_LATEST", "Forecast Start is incompatible with Start No Later Than" & vbCrLf & "-> align Forecast Start or update the constraint", "Forecast Start incompatible avec Start No Later Than" & vbCrLf & "-> aligner Forecast Start ou modifier la contrainte")
    definitions(156) = Array("DIAG.CONSTRAINT.PM.CALCULATED_START_AFTER_LATEST", "Start No Later Than cannot be met" & vbCrLf & "-> fix logic, duration, or the constraint", "Start No Later Than impossible à respecter" & vbCrLf & "-> corriger la logique, la durée ou la contrainte")
    definitions(157) = Array("DIAG.CONSTRAINT.PM.ACTUAL_START_DIFFERS_MSO", "Actual Start is incompatible with Must Start On" & vbCrLf & "-> align Actual Start or update the constraint", "Actual Start incompatible avec Must Start On" & vbCrLf & "-> aligner Actual Start ou modifier la contrainte")
    definitions(158) = Array("DIAG.CONSTRAINT.PM.FORECAST_START_DIFFERS_MSO", "Forecast Start is incompatible with Must Start On" & vbCrLf & "-> align Forecast Start or update the constraint", "Forecast Start incompatible avec Must Start On" & vbCrLf & "-> aligner Forecast Start ou modifier la contrainte")
    definitions(159) = Array("DIAG.CONSTRAINT.PM.CALCULATED_START_DIFFERS_MSO", "Must Start On cannot be met" & vbCrLf & "-> fix upstream logic, duration, or the constraint", "Must Start On impossible à respecter" & vbCrLf & "-> corriger la logique amont, la durée ou la contrainte")
    definitions(160) = Array("DIAG.CONSTRAINT.PM.ACTUAL_FINISH_BEFORE_FINISH", "Actual Finish is incompatible with Finish No Earlier Than" & vbCrLf & "-> move Actual Finish later or update the constraint", "Actual Finish incompatible avec Finish No Earlier Than" & vbCrLf & "-> repousser Actual Finish ou modifier la contrainte")
    definitions(161) = Array("DIAG.CONSTRAINT.PM.FORECAST_FINISH_BEFORE_UPSTREAM", "Forecast Finish is incompatible with upstream finish constraints" & vbCrLf & "-> move Forecast Finish later or fix upstream logic", "Forecast Finish incompatible avec les contraintes de fin amont" & vbCrLf & "-> repousser Forecast Finish ou corriger la logique amont")
    definitions(162) = Array("DIAG.CONSTRAINT.PM.FORECAST_FINISH_BEFORE_FINISH", "Forecast Finish is incompatible with Finish No Earlier Than" & vbCrLf & "-> move Forecast Finish later or update the constraint", "Forecast Finish incompatible avec Finish No Earlier Than" & vbCrLf & "-> repousser Forecast Finish ou modifier la contrainte")
    definitions(163) = Array("DIAG.CONSTRAINT.PM.ACTUAL_FINISH_AFTER_LATEST", "Actual Finish is incompatible with Finish No Later Than" & vbCrLf & "-> align Actual Finish or update the constraint", "Actual Finish incompatible avec Finish No Later Than" & vbCrLf & "-> aligner Actual Finish ou modifier la contrainte")
    definitions(164) = Array("DIAG.CONSTRAINT.PM.FORECAST_FINISH_AFTER_LATEST", "Forecast Finish is incompatible with Finish No Later Than" & vbCrLf & "-> move Forecast Finish earlier, reduce duration, or update the constraint", "Forecast Finish incompatible avec Finish No Later Than" & vbCrLf & "-> avancer Forecast Finish, réduire la durée ou modifier la contrainte")
    definitions(165) = Array("DIAG.CONSTRAINT.PM.CALCULATED_FINISH_AFTER_LATEST", "Finish No Later Than cannot be met" & vbCrLf & "-> fix logic, duration, or the constraint", "Finish No Later Than impossible à respecter" & vbCrLf & "-> corriger la logique, la durée ou la contrainte")
    definitions(166) = Array("DIAG.CONSTRAINT.PM.ACTUAL_FINISH_DIFFERS_MFO", "Actual Finish is incompatible with Must Finish On" & vbCrLf & "-> align Actual Finish or update the constraint", "Actual Finish incompatible avec Must Finish On" & vbCrLf & "-> aligner Actual Finish ou modifier la contrainte")
    definitions(167) = Array("DIAG.CONSTRAINT.PM.FORECAST_FINISH_DIFFERS_MFO", "Forecast Finish is incompatible with Must Finish On" & vbCrLf & "-> align Forecast Finish or update the constraint", "Forecast Finish incompatible avec Must Finish On" & vbCrLf & "-> aligner Forecast Finish ou modifier la contrainte")
    definitions(168) = Array("DIAG.CONSTRAINT.PM.CALCULATED_FINISH_DIFFERS_MFO", "Must Finish On cannot be met" & vbCrLf & "-> fix upstream logic, duration, or the constraint", "Must Finish On impossible à respecter" & vbCrLf & "-> corriger la logique amont, la durée ou la contrainte")
    definitions(169) = Array("DIAG.CONSTRAINT.PM.CALCULATED_START_BEFORE_MFO_NETWORK", "Must Finish On is incompatible with upstream logic" & vbCrLf & "-> the network forces a start too late to meet the imposed finish", "Must Finish On incompatible avec la logique amont" & vbCrLf & "-> le réseau impose un démarrage trop tardif pour respecter la fin imposée")
    definitions(170) = Array("DIAG.CONSTRAINT.PM.ACTUAL_START_DIFFERS_MFO_IMPLIED", "Actual Start is incompatible with Must Finish On" & vbCrLf & "-> the entered date cannot meet the imposed finish with the current duration", "Actual Start incompatible avec Must Finish On" & vbCrLf & "-> la date saisie ne permet pas de respecter la fin imposée avec la durée actuelle")
    definitions(171) = Array("DIAG.CONSTRAINT.PM.FORECAST_START_DIFFERS_MFO_IMPLIED", "Forecast Start is incompatible with Must Finish On" & vbCrLf & "-> the entered date cannot meet the imposed finish with the current duration", "Forecast Start incompatible avec Must Finish On" & vbCrLf & "-> la date saisie ne permet pas de respecter la fin imposée avec la durée actuelle")
    definitions(172) = Array("DIAG.CONSTRAINT.PM.DURATION_INCOMPATIBLE_MSO_MFO", "Duration is incompatible with Must Start On and Must Finish On" & vbCrLf & "-> fix duration or one of the two constraints", "Durée incompatible avec Must Start On et Must Finish On" & vbCrLf & "-> corriger la durée ou l'une des deux contraintes")
    definitions(173) = Array("EVENT.DEADLINE_EXCEEDED.MESSAGE", "Deadline exceeded", "Deadline depassee")
    definitions(174) = Array("EVENT.DEADLINE_EXCEEDED.DETAIL", "calculated finish is after the deadline", "la date calculee finit apres la deadline")
    definitions(175) = Array("EVENT.PARENT_DATES_IGNORED.MESSAGE", "Dates entered on summary task", "Dates saisies sur tache parent")
    definitions(176) = Array("EVENT.PARENT_DATES_IGNORED.DETAIL", "values are ignored and calculated from child tasks", "les valeurs sont ignorees, calcul par les taches enfants")
    definitions(177) = Array("CORE.ERROR.MISSING_PREDECESSOR", "Missing predecessor", "Missing predecessor")
    definitions(178) = Array("CORE.ERROR.MISSING_PREDECESSOR_ID", "Missing predecessor: ID {Id}", "Missing predecessor: ID {Id}")
    definitions(179) = Array("CORE.ERROR.BLOCKED_PREDECESSOR_ID", "Blocked by predecessor error: ID {Id}", "Blocked by predecessor error: ID {Id}")
    definitions(180) = Array("CORE.ERROR.UNSUPPORTED_LINK_TYPE", "Unsupported link type: {LinkType}", "Unsupported link type: {LinkType}")
    definitions(181) = Array("CORE.ERROR.ACTUAL_START_DEPENDENCIES", "Actual Start violates dependencies", "Actual Start violates dependencies")
    definitions(182) = Array("CORE.ERROR.ACTUAL_FINISH_CONSTRAINTS", "Actual Finish violates finish constraints", "Actual Finish violates finish constraints")
    definitions(183) = Array("CORE.ERROR.FORECAST_START_DEPENDENCIES", "Forecast Start violates dependencies", "Forecast Start violates dependencies")
    definitions(184) = Array("CORE.ERROR.FORECAST_FINISH_CONSTRAINTS", "Forecast Finish violates finish constraints", "Forecast Finish violates finish constraints")
    definitions(185) = Array("CORE.ERROR.BASELINE_DURATION_MISSING", "Baseline Duration missing", "Baseline Duration missing")
    definitions(186) = Array("CORE.ERROR.START_NOT_COMPUTABLE", "Start date not computable", "Start date not computable")
    definitions(187) = Array("CORE.ERROR.FINISH_BEFORE_START", "Finish before start", "Finish before start")
    definitions(188) = Array("CORE.ERROR.BLOCKED_PREDECESSOR_CHAIN", "Blocked by predecessor chain", "Blocked by predecessor chain")
    definitions(189) = Array("CORE.ERROR.LOE_AS_PREDECESSOR", "LOE cannot be used as predecessor", "LOE cannot be used as predecessor")
    definitions(190) = Array("CORE.ERROR.INVALID_LOE_PREDECESSOR_ID", "Blocked by invalid LOE predecessor: ID {Id}", "Blocked by invalid LOE predecessor: ID {Id}")
    definitions(191) = Array("CORE.ERROR.LOE_SS_PREDECESSOR_MISSING", "LOE SS predecessor missing", "LOE SS predecessor missing")
    definitions(192) = Array("CORE.ERROR.LOE_SS_PREDECESSOR_NOT_FOUND", "LOE SS predecessor not found: ID {Id}", "LOE SS predecessor not found: ID {Id}")
    definitions(193) = Array("CORE.ERROR.LOE_SS_PREDECESSOR_BLOCKED", "LOE blocked by SS predecessor error: ID {Id}", "LOE blocked by SS predecessor error: ID {Id}")
    definitions(194) = Array("CORE.ERROR.LOE_SS_START_UNAVAILABLE", "LOE SS predecessor start not available: ID {Id}", "LOE SS predecessor start not available: ID {Id}")
    definitions(195) = Array("CORE.ERROR.LOE_FF_PREDECESSOR_MISSING", "LOE FF predecessor missing", "LOE FF predecessor missing")
    definitions(196) = Array("CORE.ERROR.LOE_FF_PREDECESSOR_NOT_FOUND", "LOE FF predecessor not found: ID {Id}", "LOE FF predecessor not found: ID {Id}")
    definitions(197) = Array("CORE.ERROR.LOE_FF_PREDECESSOR_BLOCKED", "LOE blocked by FF predecessor error: ID {Id}", "LOE blocked by FF predecessor error: ID {Id}")
    definitions(198) = Array("CORE.ERROR.LOE_FF_FINISH_UNAVAILABLE", "LOE FF predecessor finish not available: ID {Id}", "LOE FF predecessor finish not available: ID {Id}")
    definitions(199) = Array("CORE.ERROR.LOE_SS_REQUIRED", "LOE must have at least one SS predecessor", "LOE must have at least one SS predecessor")
    definitions(200) = Array("CORE.ERROR.LOE_FF_REQUIRED", "LOE must have at least one FF predecessor", "LOE must have at least one FF predecessor")
    definitions(201) = Array("CORE.ERROR.LOE_LINK_TYPE", "LOE only supports SS and FF predecessors", "LOE only supports SS and FF predecessors")
    definitions(202) = Array("CORE.ERROR.LOE_START_NOT_COMPUTABLE", "LOE start not computable", "LOE start not computable")
    definitions(203) = Array("CORE.ERROR.LOE_FINISH_NOT_COMPUTABLE", "LOE finish not computable", "LOE finish not computable")
    definitions(204) = Array("CORE.ERROR.LOE_FINISH_BEFORE_START", "LOE finish before start", "LOE finish before start")
    definitions(205) = Array("CORE.ERROR.UNSUPPORTED_CALENDAR_LAG_TYPE", "Unsupported link type for calendar lag: {LinkType}", "Unsupported link type for calendar lag: {LinkType}")
    definitions(206) = Array("DIAG.LABEL.CALCULATED_DATE", "Calculated date", "Date calculée")
    definitions(207) = Array("DIAG.TECH.UNKNOWN_CONSTRAINT_KEY", "Unknown constraint diagnostic key: {Key}", "Clé de diagnostic de contrainte inconnue : {Key}")
    definitions(208) = Array("DIAG.TECH.MISSING_COLUMN", "Missing column in {Table}: {Column}", "Colonne absente de {Table} : {Column}")
    definitions(209) = Array("DIAG.TECH.MISSING_REQUIRED_COLUMN", "Missing required column in {Table}: {Column}", "Colonne obligatoire absente de {Table} : {Column}")
    definitions(210) = Array("DIAG.TECH.MISSING_SOURCE_COLUMN", "Missing source column in {Table}: {Column}", "Colonne source absente de {Table} : {Column}")
    definitions(211) = Array("DIAG.TECH.MISSING_TARGET_COLUMN", "Missing target column in {Table}: {Column}", "Colonne cible absente de {Table} : {Column}")
    definitions(212) = Array("DIAG.TECH.MISSING_OUTPUT_COLUMN", "Missing output column in {Table}: {Column}", "Colonne de sortie absente de {Table} : {Column}")
    definitions(213) = Array("DIAG.TECH.MISSING_REQUIRED_COLUMNS", "Missing required columns in {Context}:" & vbCrLf & "{Columns}", "Colonnes obligatoires absentes de {Context} :" & vbCrLf & "{Columns}")
    definitions(214) = Array("DIAG.TECH.FORBIDDEN_ANALYTICS_WRITE", "Forbidden analytics WBS write attempted: {Column}", "Tentative interdite d'écriture Analytics dans WBS : {Column}")
    definitions(215) = Array("DIAG.TECH.FORBIDDEN_ANALYTICS_CLEAR", "Forbidden analytics WBS clear attempted: {Column}", "Tentative interdite d'effacement Analytics dans WBS : {Column}")
    definitions(216) = Array("DIAG.TECH.INVALID_CALENDAR_TYPE", "Invalid calendar type: {Value}", "Type de calendrier invalide : {Value}")
    definitions(217) = Array("DIAG.TECH.UNSUPPORTED_CALENDAR_TYPE", "Unsupported calendar type: {Value}", "Type de calendrier non pris en charge : {Value}")
    definitions(218) = Array("DIAG.TECH.INVALID_DIGEST_CONTRACT", "Windows CNG returned an invalid SHA-256 contract.", "Windows CNG a retourné un contrat SHA-256 invalide.")
    definitions(219) = Array("DIAG.TECH.NATIVE_OPERATION_FAILED", "{Operation} failed with NTSTATUS {Status}.", "Échec de {Operation} avec NTSTATUS {Status}.")
    definitions(220) = Array("DIAG.TECH.UNKNOWN_LANGUAGE_OWNER", "Unknown language owner: {Owner}", "Owner de langue inconnu : {Owner}")
    definitions(221) = Array("DIAG.TECH.REQUIRED_WORKSHEET_MISSING", "Required worksheet is missing: {Sheet}", "Feuille obligatoire absente : {Sheet}")
    definitions(222) = Array("DIAG.TECH.REQUIRED_TABLE_MISSING", "Required table is missing: {Object}", "Table obligatoire absente : {Object}")
    definitions(223) = Array("DIAG.TECH.REQUIRED_CHART_MISSING", "Required chart is missing: {Object}", "Graphique obligatoire absent : {Object}")
    definitions(224) = Array("DIAG.TECH.REQUIRED_DATE_AXIS_MISSING", "Required date axis is missing: {Object}", "Axe de dates obligatoire absent : {Object}")
    definitions(225) = Array("DIAG.TECH.REQUIRED_COLUMN_MISSING", "Required column is missing: {Column}", "Colonne obligatoire absente : {Column}")
    definitions(226) = Array("DIAG.HARNESS.RUNTIME_ABORT", "Test abort", "Arrêt test")
    definitions(227) = Array("DIAG.PARSER.PREDECESSOR.SPACES", "Spaces are not allowed in predecessor tokens.", "Les espaces ne sont pas autorisés dans les prédécesseurs.")
    definitions(228) = Array("DIAG.PARSER.PREDECESSOR.EMPTY_TOKEN", "Empty predecessor token.", "Prédécesseur vide.")
    definitions(229) = Array("DIAG.PARSER.PREDECESSOR.EMPTY_WBS", "Empty predecessor WBS.", "WBS prédécesseur vide.")
    definitions(230) = Array("DIAG.PARSER.PREDECESSOR.INVALID_WBS", "Invalid predecessor WBS: {Value}", "WBS prédécesseur invalide : {Value}")
    definitions(231) = Array("DIAG.PARSER.PREDECESSOR.INVALID_SUFFIX", "Invalid predecessor suffix: {Value}", "Suffixe de prédécesseur invalide : {Value}")
    definitions(232) = Array("DIAG.PARSER.PREDECESSOR.INVALID_LAG", "Invalid predecessor lag: {Value}", "Décalage de prédécesseur invalide : {Value}")
    definitions(233) = Array("DIAG.PARSER.PREDECESSOR.EMPTY_IN_TEXT", "Empty predecessor token in: {Text}", "Prédécesseur vide dans : {Text}")
    definitions(234) = Array("DIAG.PARSER.PREDECESSOR.SUCCESSOR_CONTEXT", "Successor WBS {Successor} -> {Details}", "WBS successeur {Successor} -> {Details}")
    ' These labels preserve the historical diagnostic wire text. Identical EN/FR
    ' values are explicit catalog entries, not producer-owned literals.
    definitions(235) = Array("DIAG.LABEL.START", "START", "START")
    definitions(236) = Array("DIAG.LABEL.FINISH", "FINISH", "FINISH")
    definitions(237) = Array("DIAG.LABEL.ACTUAL_START", "Actual Start", "Actual Start")
    definitions(238) = Array("DIAG.LABEL.ACTUAL_FINISH", "Actual Finish", "Actual Finish")
    definitions(239) = Array("DIAG.LABEL.FORECAST_START", "Forecast Start", "Forecast Start")
    definitions(240) = Array("DIAG.LABEL.FORECAST_FINISH", "Forecast Finish", "Forecast Finish")
    definitions(241) = Array("DIAG.LABEL.CALCULATED_START", "Calculated Start", "Calculated Start")
    definitions(242) = Array("DIAG.LABEL.CALCULATED_FINISH", "Calculated Finish", "Calculated Finish")
    definitions(243) = Array("DIAG.LABEL.EFFECTIVE_DURATION", "Effective Duration", "Effective Duration")
    definitions(244) = Array("DIAG.CONSTRAINT.TYPE.MUST_START_ON", "Must Start On", "Must Start On")
    definitions(245) = Array("DIAG.CONSTRAINT.TYPE.MUST_FINISH_ON", "Must Finish On", "Must Finish On")
    definitions(246) = Array("DIAG.CONSTRAINT.TYPE.UPSTREAM_FINISH", "Upstream finish constraint", "Upstream finish constraint")
    definitions(247) = Array("DIAG.CONSTRAINT.TYPE.START_GENERIC", "START constraint", "START constraint")
    definitions(248) = Array("DIAG.CONSTRAINT.TYPE.FINISH_GENERIC", "FINISH constraint", "FINISH constraint")
    definitions(249) = Array("DIAG.CONSTRAINT.TYPE.START_NO_EARLIER", "Start No Earlier Than", "Start No Earlier Than")
    definitions(250) = Array("DIAG.CONSTRAINT.TYPE.START_NO_LATER", "Start No Later Than", "Start No Later Than")
    definitions(251) = Array("DIAG.CONSTRAINT.TYPE.FINISH_NO_EARLIER", "Finish No Earlier Than", "Finish No Earlier Than")
    definitions(252) = Array("DIAG.CONSTRAINT.TYPE.FINISH_NO_LATER", "Finish No Later Than", "Finish No Later Than")
    definitions(253) = Array("DIAG.CONSTRAINT.TYPE.MUST_START_ON_CANONICAL", "Must Start On", "Must Start On")
    definitions(254) = Array("DIAG.CONSTRAINT.TYPE.MUST_FINISH_ON_CANONICAL", "Must Finish On", "Must Finish On")
    definitions(255) = Array("DIAG.COMMON.ID_TASK_LABEL", "ID {Id}", "ID {Id}")
    TextCatalogDiagnostics_Definitions = definitions
End Function
