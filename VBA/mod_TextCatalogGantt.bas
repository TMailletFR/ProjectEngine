Attribute VB_Name = "mod_TextCatalogGantt"
Option Explicit

' Gantt text definitions. Language selection remains owned by mod_GanttLanguage.

Public Const TXT_GANTT_SCENARIO_CREATE_ERROR As String = "GANTT.SCENARIO.CREATE.ERROR"
Public Const TXT_GANTT_SCENARIO_CREATE_TITLE As String = "GANTT.SCENARIO.CREATE.TITLE"

Public Function TextCatalogGantt_Definitions() As Variant
    Dim definitions(0 To 97) As Variant
    definitions(0) = Array(TXT_GANTT_SCENARIO_CREATE_ERROR, _
        "Error while creating the scenario planning:" & vbCrLf & "{Error}", _
        "Erreur pendant la création du planning scénario :" & vbCrLf & "{Error}")
    definitions(1) = Array(TXT_GANTT_SCENARIO_CREATE_TITLE, _
        "Create scenario planning", "Créer un planning scénario")
    definitions(2) = Array("GANTT.VIEW.TITLE", "GANTT VIEW", "VUE GANTT")
    definitions(3) = Array("GANTT.COLUMN.TASK_NAME", "Task Name", "Nom tâche")
    definitions(4) = Array("GANTT.COLUMN.START", "Start", "Début")
    definitions(5) = Array("GANTT.COLUMN.FINISH", "Finish", "Fin")
    definitions(6) = Array("GANTT.COLUMN.TEST_START", "Test Start", "Début test")
    definitions(7) = Array("GANTT.COLUMN.TEST_FINISH", "Test Finish", "Fin test")
    definitions(8) = Array("GANTT.COLUMN.DURATION", "Duration", "Durée")
    definitions(9) = Array("GANTT.COLUMN.TEST_PROGRESS", "Test %", "Test %")
    definitions(10) = Array("GANTT.COLUMN.LOGIC", "Logic", "Logique")
    definitions(11) = Array("GANTT.COMMAND.RESET", "Reset", "Réinitialiser")
    definitions(12) = Array("GANTT.COMMAND.SCENARIO", "Scenario", "Scénario")
    definitions(13) = Array("GANTT.COMMAND.TEST", "Test", "Test")
    definitions(14) = Array("GANTT.COMMAND.LOCK", "Lock", "Verrouiller")
    definitions(15) = Array("GANTT.CONTROL.VIEW", "Detail / Summary", "Détail / Synthèse")
    definitions(16) = Array("GANTT.CONTROL.SCALE", "Day / Week / Month", "Jour / Sem. / Mois")
    definitions(17) = Array("GANTT.CONTROL.CONSTRAINT", "Constraint", "Contrainte")
    definitions(18) = Array("GANTT.CONTROL.PATH", "None / Critical Path / Longest Path", "N/A / Chem. Crit. / Le plus long")
    definitions(19) = Array("GANTT.CONTROL.MULTI_PROJECT", "Single / Multiple Project", "Unique / Multi-projet")
    definitions(20) = Array("GANTT.SCENARIO.CREATE.UNSAVED", _
        "The source workbook must be saved before creating a scenario planning workbook.", _
        "Le fichier source doit être enregistré avant de créer un planning scénario.")
    definitions(21) = Array("GANTT.SCENARIO.INITIALIZE.ERROR", _
        "Error while initializing the new scenario planning:" & vbCrLf & "{Error}", _
        "Erreur pendant l'initialisation du nouveau planning scénario :" & vbCrLf & "{Error}")
    definitions(22) = Array("GANTT.SCENARIO.CREATE.CONFIRMATION", _
        "You are currently in scenario mode." & vbCrLf & vbCrLf & _
        "Direct locking is not allowed in scenario mode." & vbCrLf & vbCrLf & _
        "Do you want to create a new scenario planning workbook based on the current calculated state?" & vbCrLf & vbCrLf & _
        "The new file will:" & vbCrLf & _
        "* use the current scenario as its new baseline;" & vbCrLf & _
        "* keep the scenario % Progress;" & vbCrLf & _
        "* clear Actual and Forecast;" & vbCrLf & _
        "* deactivate constraints;" & vbCrLf & _
        "* clear history and ACK;" & vbCrLf & _
        "* leave scenario mode.", _
        "Vous êtes actuellement en mode scénario." & vbCrLf & vbCrLf & _
        "Le lock direct n'est pas autorisé en mode scénario." & vbCrLf & vbCrLf & _
        "Voulez-vous créer un nouveau planning scénario basé sur l'état calculé actuel ?" & vbCrLf & vbCrLf & _
        "Le nouveau fichier :" & vbCrLf & _
        "* utilisera le scénario actuel comme nouvelle baseline ;" & vbCrLf & _
        "* conservera les % Progress du scénario ;" & vbCrLf & _
        "* videra Actual et Forecast ;" & vbCrLf & _
        "* désactivera les contraintes ;" & vbCrLf & _
        "* videra l'historique et les ACK ;" & vbCrLf & _
        "* sortira du mode scénario.")
    definitions(23) = Array("GANTT.DRAG.NO_ENGINE", _
        "Drag ignored: no compatible active simulation engine was found." & vbCrLf & _
        "Task: {Task}" & vbCrLf & "Active mode: {Mode}" & vbCrLf & _
        "Activate TEST or SCENARIO before moving a task.", _
        "Drag ignoré : aucun moteur de simulation actif compatible n'a été trouvé." & vbCrLf & _
        "Tâche : {Task}" & vbCrLf & "Mode actif : {Mode}" & vbCrLf & _
        "Activez TEST ou SCENARIO avant de déplacer une tâche.")
    definitions(24) = Array("GANTT.DRAG.SUCCESS", _
        "{Engine} modification applied." & vbCrLf & "Task: {Task}" & vbCrLf & _
        "Requested modification: {Changes}" & vbCrLf & _
        "The schedule has been recalculated. Any consequences on other tasks come from the planning engine." & vbCrLf & vbCrLf & _
        "To abandon this simulation and return to the last calculated schedule, click Reset.", _
        "Modification {Engine} appliquée." & vbCrLf & "Tâche : {Task}" & vbCrLf & _
        "Modification demandée : {Changes}" & vbCrLf & _
        "Le planning a été recalculé. Les conséquences éventuelles sur les autres tâches proviennent du moteur planning." & vbCrLf & vbCrLf & _
        "Pour abandonner cette simulation et revenir au dernier planning calculé, cliquez sur Réinitialiser.")
    definitions(25) = Array("GANTT.DRAG.SUCCESS.TEST", _
        "{Engine} modification applied." & vbCrLf & "Task: {Task}" & vbCrLf & _
        "Requested modification: {Changes}" & vbCrLf & _
        "The schedule has been recalculated. Any consequences on other tasks come from the planning engine." & vbCrLf & vbCrLf & _
        "To abandon this simulation and return to the last calculated schedule, click Reset." & vbCrLf & _
        "To remove only one assumption, clear the corresponding yellow TEST cell and run TEST again.", _
        "Modification {Engine} appliquée." & vbCrLf & "Tâche : {Task}" & vbCrLf & _
        "Modification demandée : {Changes}" & vbCrLf & _
        "Le planning a été recalculé. Les conséquences éventuelles sur les autres tâches proviennent du moteur planning." & vbCrLf & vbCrLf & _
        "Pour abandonner cette simulation et revenir au dernier planning calculé, cliquez sur Réinitialiser." & vbCrLf & _
        "Pour retirer uniquement une hypothèse, videz la cellule TEST jaune correspondante puis relancez TEST.")
    definitions(26) = Array("GANTT.DRAG.SUCCESS.SCENARIO", _
        "{Engine} modification applied." & vbCrLf & "Task: {Task}" & vbCrLf & _
        "Requested modification: {Changes}" & vbCrLf & _
        "The schedule has been recalculated. Any consequences on other tasks come from the planning engine." & vbCrLf & vbCrLf & _
        "To abandon this simulation and return to the last calculated schedule, click Reset." & vbCrLf & _
        "To remove only one assumption, clear the corresponding yellow cell and click Scenario.", _
        "Modification {Engine} appliquée." & vbCrLf & "Tâche : {Task}" & vbCrLf & _
        "Modification demandée : {Changes}" & vbCrLf & _
        "Le planning a été recalculé. Les conséquences éventuelles sur les autres tâches proviennent du moteur planning." & vbCrLf & vbCrLf & _
        "Pour abandonner cette simulation et revenir au dernier planning calculé, cliquez sur Réinitialiser." & vbCrLf & _
        "Pour retirer uniquement une hypothèse, videz la cellule jaune correspondante puis cliquez sur Scénario.")
    definitions(27) = Array("GANTT.DRAG.FAILURE", _
        "The modification requested by Drag could not be applied." & vbCrLf & _
        "Task: {Task}" & vbCrLf & "Requested modification: {Changes}", _
        "La modification demandée par Drag n'a pas pu être appliquée." & vbCrLf & _
        "Tâche : {Task}" & vbCrLf & "Modification demandée : {Changes}")
    definitions(28) = Array("GANTT.DRAG.CHANGES.UNKNOWN", "unidentified modification", "modification non identifiée")
    definitions(29) = Array("GANTT.DRAG.CHANGES.SAME", "start and finish = {Start}", "début et fin = {Start}")
    definitions(30) = Array("GANTT.DRAG.CHANGES.BOTH", "start = {Start}, finish = {Finish}", "début = {Start}, fin = {Finish}")
    definitions(31) = Array("GANTT.DRAG.CHANGES.START", "start = {Start}", "début = {Start}")
    definitions(32) = Array("GANTT.DRAG.CHANGES.FINISH", "finish = {Finish}", "fin = {Finish}")
    definitions(33) = Array("GANTT.DRAG.CHANGES.NONE", "no date changed", "aucune date modifiée")
    definitions(34) = Array("GANTT.EVENTS.INPUT_AREA_INVALID", "Input is only allowed in yellow test columns for leaf tasks.", "Saisie autorisée uniquement dans les colonnes test jaunes des tâches leaf.")
    definitions(35) = Array("GANTT.EVENTS.DATE_INVALID", "Invalid date in a test column.", "Date invalide dans une colonne test.")
    definitions(36) = Array("GANTT.EVENTS.PERCENT_INVALID", "Invalid value in Test %.", "Valeur invalide dans Test %.")
    definitions(37) = Array("GANTT.EVENTS.HANDLE_ERROR", _
        "VBA error in Handle_Gantt_Change." & vbCrLf & "-> check the last edited block in mod_GanttEvents.", _
        "Erreur VBA dans Handle_Gantt_Change." & vbCrLf & "-> vérifier le dernier bloc modifié dans mod_GanttEvents.")
    definitions(38) = Array("GANTT.SIMULATION.VBA_ERROR", "VBA error in {Function}" & vbCrLf & "-> {Details}", "Erreur VBA dans {Function}" & vbCrLf & "-> {Details}")
    definitions(39) = Array("GANTT.LOCK.NO_CHANGES", "No test changes to lock.", "Aucune modification test à verrouiller.")
    definitions(40) = Array("GANTT.LOCK.PRELIMINARY_ERRORS", _
        "Lock cancelled: the preliminary TEST simulation contains errors." & vbCrLf & "-> fix test values or upstream logic before locking.", _
        "Lock annulé : la simulation TEST préalable contient des erreurs." & vbCrLf & "-> corriger les valeurs test ou la logique amont avant de verrouiller.")
    definitions(41) = Array("GANTT.LOCK.NO_SIMULATED_RESULT", "Lock cancelled: no usable simulated result was found after TEST refresh.", "Lock annulé : aucun résultat simulé exploitable n'a été trouvé après le refresh TEST.")
    definitions(42) = Array("GANTT.LOCK.CALCULATION_ERRORS", "Lock cancelled: calculation found errors. Original WBS values were restored and test inputs were preserved.", "Lock annulé : le calcul a détecté des erreurs. Les valeurs WBS d'origine ont été restaurées et les colonnes test ont été conservées.")
    definitions(43) = Array("GANTT.LOCK.RESULT_MISMATCH", "Lock cancelled: the real recalculation does not match the retained simulated result. Original WBS values were restored and test inputs were preserved.", "Lock annulé : le recalcul réel ne correspond pas au résultat simulé retenu. Les valeurs WBS d'origine ont été restaurées et les colonnes test ont été conservées.")
    definitions(44) = Array("GANTT.LOCK.SUCCESS", "Lock successfully applied.", "Lock appliqué avec succès.")
    definitions(45) = Array("GANTT.LOCK.RUN_ERROR", "VBA error in Run_Gantt_Lock_Changes" & vbCrLf & "-> check the last edited block in mod_GanttLive", "Erreur VBA dans Run_Gantt_Lock_Changes" & vbCrLf & "-> vérifier le dernier bloc modifié dans mod_GanttLive")
    definitions(46) = Array("GANTT.SCENARIO.UPDATED", "Scenario updated.", "Scénario mis à jour.")
    definitions(47) = Array("GANTT.TEST.NO_INPUT_NORMAL", "No TEST input detected. The normal display is already restored.", "Aucune saisie TEST détectée. L'affichage normal est déjà restauré.")
    definitions(48) = Array("GANTT.TEST.NO_INPUT_RESET", "No TEST input detected. Simulation display has been reset.", "Aucune saisie TEST détectée. L’affichage simulation a été réinitialisé.")
    definitions(49) = Array("GANTT.TEST.EXECUTED", "Simulation executed.", "Simulation exécutée.")
    definitions(50) = Array("GANTT.TEST.NO_VISIBLE_CHANGE", _
        "Simulation executed, but no visible change was produced." & vbCrLf & "-> probable cause: local priority (Actual/Forecast) and/or unchanged network constraint.", _
        "Simulation exécutée, mais aucun changement visible n'a été produit." & vbCrLf & "-> cause probable : priorité locale (Actual/Forecast) et/ou contrainte réseau inchangée.")
    definitions(51) = Array("GANTT.SCENARIO.CALCULATION_ERROR_GROUP", _
        "Calculation error in scenario" & vbCrLf & "-> fix test values or upstream logic" & vbCrLf & vbCrLf & "IDs: {Ids}" & vbCrLf & "WBS: {Wbs}", _
        "Erreur de calcul dans le scénario" & vbCrLf & "-> corriger les valeurs de test ou la logique amont" & vbCrLf & vbCrLf & "IDs : {Ids}" & vbCrLf & "WBS : {Wbs}")
    definitions(52) = Array("GANTT.TEST.CALCULATION_ERROR_GROUP", _
        "Calculation error in live engine" & vbCrLf & "-> fix test values or upstream logic" & vbCrLf & vbCrLf & "IDs: {Ids}" & vbCrLf & "WBS: {Wbs}", _
        "Erreur de calcul dans le moteur live" & vbCrLf & "-> corriger les valeurs test ou la logique amont" & vbCrLf & vbCrLf & "IDs : {Ids}" & vbCrLf & "WBS : {Wbs}")
    definitions(53) = Array("GANTT.SIMULATION.MISSING_COLUMN", _
        "Required column missing for GANTT live/scenario" & vbCrLf & vbCrLf & _
        "Source/table: {Source}" & vbCrLf & "Column: {Column}" & vbCrLf & "Function: {Function}", _
        "Colonne requise introuvable pour GANTT live/scenario" & vbCrLf & vbCrLf & _
        "Table/source : {Source}" & vbCrLf & "Colonne : {Column}" & vbCrLf & "Fonction : {Function}")
    definitions(54) = Array("GANTT.TIMELINE.WEEK_PREFIX", "W", "S")
    definitions(55) = Array("GANTT.TIMELINE.MONTH.M01", "Jan", "jan")
    definitions(56) = Array("GANTT.TIMELINE.MONTH.M02", "Feb", "fév")
    definitions(57) = Array("GANTT.TIMELINE.MONTH.M03", "Mar", "mar")
    definitions(58) = Array("GANTT.TIMELINE.MONTH.M04", "Apr", "avr")
    definitions(59) = Array("GANTT.TIMELINE.MONTH.M05", "May", "mai")
    definitions(60) = Array("GANTT.TIMELINE.MONTH.M06", "Jun", "juin")
    definitions(61) = Array("GANTT.TIMELINE.MONTH.M07", "Jul", "juil")
    definitions(62) = Array("GANTT.TIMELINE.MONTH.M08", "Aug", "août")
    definitions(63) = Array("GANTT.TIMELINE.MONTH.M09", "Sep", "sep")
    definitions(64) = Array("GANTT.TIMELINE.MONTH.M10", "Oct", "oct")
    definitions(65) = Array("GANTT.TIMELINE.MONTH.M11", "Nov", "nov")
    definitions(66) = Array("GANTT.TIMELINE.MONTH.M12", "Dec", "déc")
    definitions(67) = Array("GANTT.CONTROL.MULTI_CRITICAL_PATH", "Multi Critical Path", "Chemins critiques multiples")
    definitions(68) = Array("GANTT.DRAG.TASK.UNKNOWN", "(unknown task)", "(tâche inconnue)")
    definitions(69) = Array("GANTT.DRAG.ENGINE.DRAG", "Drag", "Glisser")
    definitions(70) = Array("GANTT.DRAG.ENGINE.TEST", "Drag/Test", "Glisser/Test")
    definitions(71) = Array("GANTT.DRAG.ENGINE.SCENARIO", "Drag/Scenario", "Glisser/Scénario")
    definitions(72) = Array("GANTT.TEST.ACTUAL_IMPACT_WARNING", _
        "TEST INPUT IMPACTS ACTUAL TASK - LOCK MAY NOT FULLY PERSIST", _
        "LA SAISIE TEST AFFECTE UNE TACHE ACTUAL - LE VERROUILLAGE PEUT NE PAS PERSISTER ENTIÈREMENT")
    definitions(73) = Array("GANTT.ERROR.UNKNOWN_ENSURE_MODE", "Unknown Gantt ensure mode: {Mode}", "Mode de préparation Gantt inconnu : {Mode}")
    definitions(74) = Array("GANTT.ERROR.UNKNOWN_UPDATE_SCOPE", "Unknown Gantt update scope: {Scope}", "Périmètre de mise à jour Gantt inconnu : {Scope}")
    definitions(75) = Array("GANTT.ERROR.UNSUPPORTED_DEPENDENCY_RENDER_MODE", "Unsupported dependency render mode: {Mode}", "Mode de rendu des dépendances non pris en charge : {Mode}")
    definitions(76) = Array("GANTT.ERROR.RENDER_NOT_READY", "{Context} did not reach READY state.", "{Context} n'a pas atteint l'état READY.")
    definitions(77) = Array("GANTT.ERROR.DEPENDENCY_ROUTING", "Dependency routing failed: {Details}", "Échec du routage des dépendances : {Details}")
    definitions(78) = Array("GANTT.ERROR.GEOMETRY_RECONCILIATION", "Gantt geometry reconciliation failed.", "Échec de la réconciliation de la géométrie Gantt.")
    definitions(79) = Array("GANTT.ERROR.CONSTRAINT_OVERLAY", "Constraint overlay refresh failed.", "Échec du rafraîchissement de la surcouche de contraintes.")
    definitions(80) = Array("GANTT.ERROR.MISSING_SOURCE_COLUMN", "Missing source column in {Table}: {Column}", "Colonne source absente de {Table} : {Column}")
    definitions(81) = Array("GANTT.ERROR.MISSING_WBS_COLUMN", "Missing WBS column: {Column}", "Colonne WBS absente : {Column}")
    definitions(82) = Array("GANTT.ERROR.MISSING_CALC_COLUMN", "Missing CALC column: {Column}", "Colonne CALC absente : {Column}")
    definitions(83) = Array("GANTT.ERROR.MISSING_SCENARIO_COLUMN", "Missing scenario column: {Column}", "Colonne scénario absente : {Column}")
    definitions(84) = Array("GANTT.ERROR.STYLE_SHAPE_MISSING", "Style-only render found a missing shape: {Shape}", "Le rendu de style a détecté une shape absente : {Shape}")
    definitions(85) = Array("GANTT.ERROR.STYLE_SHAPE_INCOMPATIBLE", "Style-only render found an incompatible shape: {Shape}", "Le rendu de style a détecté une shape incompatible : {Shape}")
    definitions(86) = Array("GANTT.ERROR.UPDATE_NOT_READY", "Gantt renderer did not reach READY state after Update Gantt.", "Le renderer Gantt n'a pas atteint l'état READY après Update Gantt.")
    definitions(87) = Array("GANTT.ERROR.FULL_UPDATE_GANTT_NOT_READY", "Gantt renderer did not reach READY state after Full Update on GANTT.", "Le renderer Gantt n'a pas atteint l'état READY après Full Update sur GANTT.")
    definitions(88) = Array("GANTT.ERROR.NO_RENDER_INVALIDATION", "Gantt NO_RENDER invalidation failed.", "Échec de l'invalidation Gantt NO_RENDER.")
    definitions(89) = Array("GANTT.ERROR.FULL_UPDATE_NOT_READY", "Gantt renderer did not reach READY state after Full Update.", "Le renderer Gantt n'a pas atteint l'état READY après Full Update.")
    definitions(90) = Array("GANTT.ERROR.CRITICAL_PATH_NOT_READY", "Gantt render did not reach READY state after Critical Path mode change.", "Le rendu Gantt n'a pas atteint l'état READY après le changement de mode Critical Path.")
    definitions(91) = Array("GANTT.ERROR.LOCK_NOT_READY", "Gantt LOCK render did not reach READY state.", "Le rendu Gantt LOCK n'a pas atteint l'état READY.")
    definitions(92) = Array("GANTT.ERROR.SCENARIO_NOT_READY", "Gantt SCENARIO render did not reach READY state.", "Le rendu Gantt SCENARIO n'a pas atteint l'état READY.")
    definitions(93) = Array("GANTT.ERROR.RESET_NOT_READY", "Gantt Reset render did not reach READY state.", "Le rendu Gantt Reset n'a pas atteint l'état READY.")
    definitions(94) = Array("GANTT.ERROR.TEST_EMPTY_NOT_READY", "Gantt TEST empty convergence did not reach READY state.", "La convergence vide Gantt TEST n'a pas atteint l'état READY.")
    definitions(95) = Array("GANTT.ERROR.TEST_NO_INPUT_NOT_READY", "Gantt TEST no-input convergence did not reach READY state.", "La convergence sans saisie Gantt TEST n'a pas atteint l'état READY.")
    definitions(96) = Array("GANTT.ERROR.TEST_NOT_READY", "Gantt TEST render did not reach READY state.", "Le rendu Gantt TEST n'a pas atteint l'état READY.")
    definitions(97) = Array("GANTT.TEST.ROW.ACTUAL_WARNING", _
        "TEST INPUT ON ACTUAL TASK - SIMULATION MAY SHOW NO EFFECT AND LOCK MAY NOT MATCH PROD", _
        "SAISIE TEST SUR UNE TÂCHE ACTUAL - LA SIMULATION PEUT NE PRODUIRE AUCUN EFFET ET LE VERROUILLAGE PEUT DIFFÉRER DE PROD")
    TextCatalogGantt_Definitions = definitions
End Function
