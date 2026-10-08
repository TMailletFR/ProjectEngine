Attribute VB_Name = "mod_TextCatalogWBS"
Option Explicit

' WBS UI definitions; consumers supply their existing owner's language.

Public Const TXT_WBS_UPDATE_PLANNING_LABEL As String = "WBS.ACTION.UPDATE_PLANNING.LABEL"
Public Const TXT_WBS_UPDATE_GANTT_LABEL As String = "WBS.ACTION.UPDATE_GANTT.LABEL"
Public Const TXT_WBS_UPDATE_SCURVE_LABEL As String = "WBS.ACTION.UPDATE_SCURVE.LABEL"
Public Const TXT_WBS_FORCED_UPDATE_LABEL As String = "WBS.ACTION.FORCED_UPDATE.LABEL"
Public Const TXT_WBS_FULL_UPDATE_LABEL As String = "WBS.ACTION.FULL_UPDATE.LABEL"
Public Const TXT_WBS_RESET_PLANNING_LABEL As String = "WBS.ACTION.RESET_PLANNING.LABEL"
Private Const WBS_TUTORIAL_KEY_PREFIX As String = "WBS.TUTORIAL."

Public Function TextCatalogWBS_Definitions() As Variant
    Dim definitions() As Variant
    Dim tutorialSources As Object
    Dim columnKey As Variant
    Dim localizedPair As Variant
    Dim definitionIndex As Long

    Set tutorialSources = TextCatalogWBS_TutorialSourceMap()
    ReDim definitions(0 To 60 + tutorialSources.Count)
    definitions(0) = Array(TXT_WBS_UPDATE_PLANNING_LABEL, "Update" & vbCrLf & "Planning", "Mettre à jour" & vbCrLf & "Planning")
    definitions(1) = Array(TXT_WBS_UPDATE_GANTT_LABEL, "Update" & vbCrLf & "Gantt", "Mettre à jour" & vbCrLf & "Gantt")
    definitions(2) = Array(TXT_WBS_UPDATE_SCURVE_LABEL, "Update" & vbCrLf & "S-Curve", "Mettre à jour" & vbCrLf & "S-Curve")
    definitions(3) = Array(TXT_WBS_FORCED_UPDATE_LABEL, "Forced" & vbCrLf & "Planning Update", "MàJ forcée" & vbCrLf & "Planning")
    definitions(4) = Array(TXT_WBS_FULL_UPDATE_LABEL, "Full" & vbCrLf & "Update", "Mise à jour" & vbCrLf & "complète")
    definitions(5) = Array(TXT_WBS_RESET_PLANNING_LABEL, "Reset" & vbCrLf & "Planning", "Réinitialiser" & vbCrLf & "Planning")
    definitions(6) = Array("WBS.ONBOARDING.QUICK_START", "Quick Start: Fill in the columns marked Required. Optional columns may be left blank, while Calculated columns are populated automatically." & vbLf & "* Predecessors WBS, Baseline Start and Baseline Duration must together provide enough information to position the task on the timeline.", "Démarrage rapide : renseignez les colonnes portant le statut Requis. Les colonnes Optionnel peuvent rester vides et les colonnes Calculé sont renseignées automatiquement." & vbLf & "* Predecessors WBS, Baseline Start et Baseline Duration doivent ensemble fournir suffisamment d'informations pour positionner la tâche sur la chronologie.")
    definitions(7) = Array("WBS.ONBOARDING.OPTIONAL", "Optional", "Optionnel")
    definitions(8) = Array("WBS.ONBOARDING.CALCULATED", "Calculated", "Calculé")
    definitions(9) = Array("WBS.ONBOARDING.REQUIRED", "Required", "Requis")
    definitions(10) = Array("WBS.SUMMARY.BASELINE_FINISH", "Project Baseline" & vbCrLf & "Finish", "Fin baseline" & vbCrLf & "projet")
    definitions(11) = Array("WBS.SUMMARY.CALCULATED_FINISH", "Project" & vbCrLf & "Calculated Finish", "Fin calculée" & vbCrLf & "projet")
    definitions(12) = Array("WBS.SUMMARY.DELAY", "Project Delay" & vbCrLf & "(days)", "Retard projet" & vbCrLf & "(jours)")
    definitions(13) = Array("WBS.FULL_RESET.CONFIRMATION", _
        "This will clear the planning, Dashboard, event history and acknowledgements." & vbCrLf & "Continue?", _
        "Cette action va vider le planning, le Dashboard, l'historique et les acquittements." & vbCrLf & "Continuer ?")
    definitions(14) = Array("WBS.FULL_RESET.TITLE", "Full Reset", "Réinitialisation complète")
    definitions(15) = Array("WBS.ERROR.FULL_RESET", "Error in Full Reset: {Details}", "Erreur dans Full Reset : {Details}")
    definitions(16) = Array("WBS.ERROR.RESET_PLANNING", "Error in Reset Planning: {Details}", "Erreur dans Reset Planning : {Details}")
    definitions(17) = Array("WBS.ERROR.TASK_TYPE_SETUP", "Error in Ensure_WBS_TaskType_Input_Setup: {Details}", "Erreur dans Ensure_WBS_TaskType_Input_Setup : {Details}")
    definitions(18) = Array("WBS.EVENTS.CALCULATED_EDIT", "Manual edit not allowed in calculated column (gray)." & vbCrLf & "-> edit input columns only (blue).", "Modification interdite dans une colonne calculée (grise)." & vbCrLf & "-> modifier uniquement les colonnes d'entrée (bleues).")
    definitions(19) = Array("WBS.EVENTS.MACRO.CALCULATED_EDIT", "Engine write not allowed in a calculated WBS column." & vbCrLf & "-> source: {Source}" & vbCrLf & "-> the macro run was stopped to prevent data corruption.", "Écriture moteur interdite dans une colonne calculée de WBS." & vbCrLf & "-> source : {Source}" & vbCrLf & "-> le macro-run est arrêté pour éviter une corruption de données.")
    definitions(20) = Array("WBS.EVENTS.WBS_FORMAT", "Invalid format in WBS." & vbCrLf & "-> expected format: 1 | 1.2 | 1.2.3", "Format invalide dans WBS." & vbCrLf & "-> format attendu : 1 | 1.2 | 1.2.3")
    definitions(21) = Array("WBS.EVENTS.MACRO.WBS_FORMAT", "Invalid WBS format detected during macro execution." & vbCrLf & "-> expected format: 1 | 1.2 | 1.2.3", "Format invalide détecté dans WBS pendant l'exécution d'un macro." & vbCrLf & "-> format attendu : 1 | 1.2 | 1.2.3")
    definitions(22) = Array("WBS.EVENTS.PREDECESSORS_FORMAT", _
        "Invalid format in Predecessors WBS" & vbCrLf & vbCrLf & "Accepted formats:" & vbCrLf & vbCrLf & "* 1" & vbCrLf & "* 1+3" & vbCrLf & "* 1-2" & vbCrLf & "* 1SS" & vbCrLf & "* 1FF" & vbCrLf & "* 1SS-2" & vbCrLf & "* 1FF+4" & vbCrLf & "* 1;2SS+3;4FF-2" & vbCrLf & vbCrLf & "Rules:" & vbCrLf & vbCrLf & "* FS is implicit when no type is provided" & vbCrLf & "* zero lag is implicit" & vbCrLf & "* multiple links are separated by ;" & vbCrLf & "* spaces are not allowed" & vbCrLf & "* empty tokens are not allowed", _
        "Format invalide dans Predecessors WBS" & vbCrLf & vbCrLf & "Formats acceptes :" & vbCrLf & vbCrLf & "* 1" & vbCrLf & "* 1+3" & vbCrLf & "* 1-2" & vbCrLf & "* 1SS" & vbCrLf & "* 1FF" & vbCrLf & "* 1SS-2" & vbCrLf & "* 1FF+4" & vbCrLf & "* 1;2SS+3;4FF-2" & vbCrLf & vbCrLf & "Regles :" & vbCrLf & vbCrLf & "* FS est implicite si aucun type n'est indique" & vbCrLf & "* le lag 0 est implicite" & vbCrLf & "* plusieurs liens sont separes par ;" & vbCrLf & "* les espaces ne sont pas autorises" & vbCrLf & "* les elements vides ne sont pas autorises")
    definitions(23) = Array("WBS.EVENTS.TASK_TYPE", "Invalid Task Type." & vbCrLf & "-> allowed values: Task | Milestone | Level of Effort", "Task Type invalide." & vbCrLf & "-> valeurs autorisées : Task | Milestone | Level of Effort")
    definitions(24) = Array("WBS.EVENTS.MACRO.TASK_TYPE", "Invalid Task Type detected during macro execution." & vbCrLf & "-> allowed values: Task | Milestone | Level of Effort", "Task Type invalide détecté pendant l'exécution d'un macro." & vbCrLf & "-> valeurs autorisées : Task | Milestone | Level of Effort")
    definitions(25) = Array("WBS.EVENTS.SUMMARY_DISPLAY", "Invalid S." & vbCrLf & "-> allowed values: blank | Y | N", "Valeur invalide dans S." & vbCrLf & "-> valeurs autorisees : vide | Y | N")
    definitions(26) = Array("WBS.EVENTS.MACRO.SUMMARY_DISPLAY", "Invalid S detected during macro execution." & vbCrLf & "-> allowed values: blank | Y | N", "Valeur S invalide detectee pendant l'execution d'un macro." & vbCrLf & "-> valeurs autorisees : vide | Y | N")
    definitions(27) = Array("WBS.EVENTS.CALENDAR", "Invalid Cal." & vbCrLf & "-> allowed values: blank | 7j/7 | 6j/7 | 5j/7", "Calendrier invalide." & vbCrLf & "-> valeurs autorisees : vide | 7j/7 | 6j/7 | 5j/7")
    definitions(28) = Array("WBS.EVENTS.MACRO.CALENDAR", "Invalid Cal detected during macro execution." & vbCrLf & "-> allowed values: blank | 7j/7 | 6j/7 | 5j/7", "Calendrier invalide detecte pendant l'execution d'un macro." & vbCrLf & "-> valeurs autorisees : vide | 7j/7 | 6j/7 | 5j/7")
    definitions(29) = Array("WBS.EVENTS.DURATION", "Invalid duration." & vbCrLf & "-> enter a strictly positive numeric duration.", "Duree invalide." & vbCrLf & "-> saisir une duree numerique strictement positive.")
    definitions(30) = Array("WBS.EVENTS.MACRO.DURATION", "Invalid duration detected during macro execution." & vbCrLf & "-> enter a strictly positive numeric duration.", "Duree invalide detectee pendant l'execution d'un macro." & vbCrLf & "-> saisir une duree numerique strictement positive.")
    definitions(31) = Array("WBS.EVENTS.HANDLE_ERROR", "VBA error in Handle_WBS_Change" & vbCrLf & "-> check the last edited block in mod_WBSEvents", "Erreur VBA dans Handle_WBS_Change" & vbCrLf & "-> vérifier le dernier bloc modifié dans mod_WBSEvents")
    definitions(32) = Array("WBS.EVENTS.MACRO.HANDLE_ERROR", "VBA error in Handle_WBS_Change." & vbCrLf & "-> the macro run was stopped to avoid an error cascade.", "Erreur VBA dans Handle_WBS_Change." & vbCrLf & "-> le macro-run est arrêté pour éviter une cascade d'erreurs.")
    definitions(33) = Array("WBS.VALIDATION.TASK_TYPE.INPUT_TITLE", "Task Type", "Type de tâche")
    definitions(34) = Array("WBS.VALIDATION.TASK_TYPE.INPUT_MESSAGE", "Choose: Task, Milestone, or Level of Effort.", "Choisir : Task, Milestone ou Level of Effort.")
    definitions(35) = Array("WBS.VALIDATION.TASK_TYPE.ERROR_TITLE", "Invalid Task Type", "Type de tâche invalide")
    definitions(36) = Array("WBS.VALIDATION.TASK_TYPE.ERROR_MESSAGE", "Allowed values: Task, Milestone, Level of Effort.", "Valeurs autorisées : Task, Milestone, Level of Effort.")
    definitions(37) = Array("WBS.VALIDATION.SUMMARY.INPUT_TITLE", "S", "S")
    definitions(38) = Array("WBS.VALIDATION.SUMMARY.INPUT_MESSAGE", "Choose Y to show in Summary, N to hide.", "Choisir Y pour afficher dans Summary, N pour masquer.")
    definitions(39) = Array("WBS.VALIDATION.SUMMARY.ERROR_TITLE", "Invalid S", "Valeur S invalide")
    definitions(40) = Array("WBS.VALIDATION.SUMMARY.ERROR_MESSAGE", "Allowed values: blank, Y, N.", "Valeurs autorisées : vide, Y, N.")
    definitions(41) = Array("WBS.VALIDATION.CALENDAR.INPUT_TITLE", "Cal", "Cal")
    definitions(42) = Array("WBS.VALIDATION.CALENDAR.INPUT_MESSAGE", "Choose: 7j/7, 6j/7, or 5j/7.", "Choisir : 7j/7, 6j/7 ou 5j/7.")
    definitions(43) = Array("WBS.VALIDATION.CALENDAR.ERROR_TITLE", "Invalid Cal", "Calendrier invalide")
    definitions(44) = Array("WBS.VALIDATION.CALENDAR.ERROR_MESSAGE", "Allowed values: blank, 7j/7, 6j/7, 5j/7.", "Valeurs autorisées : vide, 7j/7, 6j/7, 5j/7.")
    definitions(45) = Array("WBS.ERROR.AMBIGUOUS_SCHEMA", "Ambiguous WBS schema: both Project and legacy Package columns are present.", "Schéma WBS ambigu : les colonnes Project et l'ancienne colonne Package sont toutes deux présentes.")
    definitions(46) = Array("WBS.ERROR.MISSING_CANONICAL_COLUMNS", "Missing or renamed canonical WBS column(s): {Columns}", "Colonne(s) WBS canonique(s) absente(s) ou renommée(s) : {Columns}")
    definitions(47) = Array("WBS.ERROR.NO_ONBOARDING_ROW", "tbl_WBS has no row available immediately above its physical header for onboarding.", "tbl_WBS ne possède aucune ligne disponible juste au-dessus de son en-tête physique pour l'onboarding.")
    definitions(48) = Array("WBS.ERROR.UNEXPECTED_FIRST_DATA_ROW", "Unexpected first tbl_WBS data row. Expected the row immediately below the physical header, found {Row}.", "Première ligne de données tbl_WBS inattendue. La ligne située juste sous l'en-tête physique était attendue ; ligne trouvée : {Row}.")
    definitions(49) = Array("WBS.ERROR.INVALID_ONBOARDING_ROW", "tbl_WBS has no valid onboarding row immediately above its physical header.", "tbl_WBS ne possède aucune ligne d'onboarding valide juste au-dessus de son en-tête physique.")
    definitions(50) = Array("WBS.ERROR.MISSING_TASK_TYPE_COLUMN", "Missing required WBS input column: Task Type", "Colonne de saisie WBS obligatoire absente : Task Type")
    definitions(51) = Array("WBS.ERROR.MISSING_PROJECT_COLUMN", "Missing canonical WBS column: Project (legacy alias: Package).", "Colonne WBS canonique absente : Project (ancien alias : Package).")
    definitions(52) = Array("WBS.ERROR.UNKNOWN_FORMULA_KEY", "Unknown managed WBS formula column key '{Key}'.", "Clé de colonne de formule WBS gérée inconnue '{Key}'.")
    definitions(53) = Array("WBS.ERROR.NO_WRITE_SCOPE", "No active WBS write scope matches token {Token}.", "Aucun scope d'écriture WBS actif ne correspond au jeton {Token}.")
    definitions(54) = Array("WBS.ERROR.WRITE_SCOPE_ORDER", "WBS write scopes must close in LIFO order. Active token={ActiveToken}, requested token={RequestedToken}.", "Les scopes d'écriture WBS doivent se fermer dans l'ordre LIFO. Jeton actif={ActiveToken}, jeton demandé={RequestedToken}.")
    definitions(55) = Array("WBS.ERROR.INVALID_TASK_TYPE", "Invalid Task Type value: {Value} | Allowed values: Task, Milestone, Level of Effort", "Valeur Task Type invalide : {Value} | Valeurs autorisées : Task, Milestone, Level of Effort")
    definitions(56) = Array("WBS.ERROR.LOGIC_LINKS_ROWS", "tbl_LOGIC_LINKS did not materialize its data rows.", "tbl_LOGIC_LINKS n'a pas matérialisé ses lignes de données.")
    definitions(57) = Array("WBS.ERROR.INVALID_SUMMARY_VALUE", "Invalid S value: {Value} (allowed: blank, Y, N)", "Valeur S invalide : {Value} (valeurs autorisées : vide, Y, N)")
    definitions(58) = Array("WBS.VALIDATION.TASK_TYPE.LIST", _
        "Task,Milestone,Level of Effort", _
        "Task,Milestone,Level of Effort")
    definitions(59) = Array("WBS.VALIDATION.SUMMARY.LIST", "Y,N", "Y,N")
    definitions(60) = Array("WBS.VALIDATION.CALENDAR.LIST", "7j/7,6j/7,5j/7", "7j/7,6j/7,5j/7")

    definitionIndex = 61
    For Each columnKey In tutorialSources.Keys
        localizedPair = tutorialSources(CStr(columnKey))
        definitions(definitionIndex) = Array( _
            "WBS.TUTORIAL." & CStr(columnKey), _
            CStr(localizedPair(1)), _
            CStr(localizedPair(0)))
        definitionIndex = definitionIndex + 1
    Next columnKey

    TextCatalogWBS_Definitions = definitions
End Function

Public Function TextCatalogWBS_TutorialText( _
    ByVal columnKey As String, _
    ByVal languageKey As String) As String

    TextCatalogWBS_TutorialText = TextCatalog_Get( _
        WBS_TUTORIAL_KEY_PREFIX & UCase$(Trim$(columnKey)), _
        languageKey)

End Function

