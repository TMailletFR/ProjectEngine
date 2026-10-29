Attribute VB_Name = "mod_TextCatalogConstraints"
Option Explicit

Public Function TextCatalogConstraints_Definitions() As Variant
    Dim definitions(0 To 44) As Variant
    definitions(0) = Array("CONSTRAINTS.VALIDATION.START_TYPE.INPUT_TITLE", _
        "Start Constraint Type", _
        "Type de contrainte de debut")
    definitions(1) = Array("CONSTRAINTS.VALIDATION.START_TYPE.INPUT", _
        "Choose blank, Start No Earlier Than, Start No Later Than, or Must Start On.", _
        "Choisir vide, Start No Earlier Than, Start No Later Than ou Must Start On.")
    definitions(2) = Array("CONSTRAINTS.VALIDATION.START_TYPE.ERROR_TITLE", _
        "Unknown Start Constraint Type", _
        "Contrainte de debut inconnue")
    definitions(3) = Array("CONSTRAINTS.VALIDATION.START_TYPE.ERROR", _
        "Use Start No Earlier Than, Start No Later Than, or Must Start On.", _
        "Utiliser Start No Earlier Than, Start No Later Than ou Must Start On.")
    definitions(4) = Array("CONSTRAINTS.VALIDATION.FINISH_TYPE.INPUT_TITLE", _
        "Finish Constraint Type", _
        "Type de contrainte de fin")
    definitions(5) = Array("CONSTRAINTS.VALIDATION.FINISH_TYPE.INPUT", _
        "Choose blank, Finish No Earlier Than, Finish No Later Than, or Must Finish On.", _
        "Choisir vide, Finish No Earlier Than, Finish No Later Than ou Must Finish On.")
    definitions(6) = Array("CONSTRAINTS.VALIDATION.FINISH_TYPE.ERROR_TITLE", _
        "Unknown Finish Constraint Type", _
        "Contrainte de fin inconnue")
    definitions(7) = Array("CONSTRAINTS.VALIDATION.FINISH_TYPE.ERROR", _
        "Use Finish No Earlier Than, Finish No Later Than, or Must Finish On.", _
        "Utiliser Finish No Earlier Than, Finish No Later Than ou Must Finish On.")
    definitions(8) = Array("CONSTRAINTS.VALIDATION.ACTIVE.INPUT_TITLE", _
        "Active", _
        "Actif")
    definitions(9) = Array("CONSTRAINTS.VALIDATION.ACTIVE.INPUT", _
        "Choose Yes or No.", _
        "Choisir Yes ou No.")
    definitions(10) = Array("CONSTRAINTS.VALIDATION.ACTIVE.ERROR_TITLE", _
        "Unknown Active value", _
        "Valeur Actif inconnue")
    definitions(11) = Array("CONSTRAINTS.VALIDATION.ACTIVE.ERROR", _
        "Recommended values: Yes or No.", _
        "Valeurs recommandees : Yes ou No.")
    definitions(12) = Array("CONSTRAINTS.IMPORT.ERROR", _
        "Error in Import_WBS_To_Constraints" & vbCrLf & "-> {Details}", _
        "Erreur dans Import_WBS_To_Constraints" & vbCrLf & "-> {Details}")
    definitions(13) = Array("CONSTRAINTS.SYNC.ERROR", _
        "Error in Sync_Constraints_To_CALC" & vbCrLf & "-> {Details}", _
        "Erreur dans Sync_Constraints_To_CALC" & vbCrLf & "-> {Details}")
    definitions(14) = Array("CONSTRAINTS.DIAG.PARENT_IGNORED", "Active constraint ignored on a summary task", "Contrainte active ignoree sur une tache parent")
    definitions(15) = Array("CONSTRAINTS.DIAG.PARENT_IGNORED.DETAIL", "constraints on summary tasks are not exported to CALC", "les contraintes sur taches parent ne sont pas exportees vers CALC")
    definitions(16) = Array("CONSTRAINTS.DIAG.LOE_IGNORED", "Active constraint ignored on a Level of Effort task", "Contrainte active ignoree sur une tache Level of Effort")
    definitions(17) = Array("CONSTRAINTS.DIAG.LOE_IGNORED.DETAIL", "constraints on LOE tasks are not exported to CALC", "les contraintes sur LOE ne sont pas exportees vers CALC")
    definitions(18) = Array("CONSTRAINTS.DIAG.INVALID_DEADLINE", "Invalid deadline", "Deadline renseignee invalide")
    definitions(19) = Array("CONSTRAINTS.DIAG.ID_NOT_IN_CALC", "Active constraint references an ID not found in CALC", "Contrainte active sur un ID absent de CALC")
    definitions(20) = Array("CONSTRAINTS.DIAG.UNKNOWN_START_TYPE", "Unknown start constraint type", "Type de contrainte debut non reconnu")
    definitions(21) = Array("CONSTRAINTS.DIAG.UNKNOWN_FINISH_TYPE", "Unknown finish constraint type", "Type de contrainte fin non reconnu")
    definitions(22) = Array("CONSTRAINTS.DIAG.START_TYPE_WITHOUT_DATE", "Start constraint type defined without constraint date", "Type de contrainte debut renseigne sans date")
    definitions(23) = Array("CONSTRAINTS.DIAG.START_DATE_WITHOUT_TYPE", "Start constraint date defined without constraint type", "Date de contrainte debut renseignee sans type")
    definitions(24) = Array("CONSTRAINTS.DIAG.FINISH_TYPE_WITHOUT_DATE", "Finish constraint type defined without constraint date", "Type de contrainte fin renseigne sans date")
    definitions(25) = Array("CONSTRAINTS.DIAG.FINISH_DATE_WITHOUT_TYPE", "Finish constraint date defined without constraint type", "Date de contrainte fin renseignee sans type")
    definitions(26) = Array("CONSTRAINTS.DIAG.EMPTY_IGNORED", "Active empty constraint ignored", "Contrainte active vide ignoree")
    definitions(27) = Array("CONSTRAINTS.DIAG.EMPTY_IGNORED.DETAIL", "no start/finish constraint is defined ; the row is not exported to CALC", "aucune contrainte debut/fin n'est definie ; la ligne n'est pas exportee vers CALC")
    definitions(28) = Array("CONSTRAINTS.ERROR.MISSING_TABLE", "Missing table {Table}. Run Import WBS to Constraints first.", "Table {Table} introuvable. Exécuter d'abord Import WBS vers Constraints.")
    definitions(29) = Array("CONSTRAINTS.ERROR.CALC_EMPTY", "Active constraints exist but tbl_CALC is empty.", "Des contraintes actives existent mais tbl_CALC est vide.")
    definitions(30) = Array("CONSTRAINTS.ERROR.INVALID_DEADLINE_ID", "Invalid deadline in tbl_CONSTRAINTS for ID: {Id}", "Deadline invalide dans tbl_CONSTRAINTS pour l'ID : {Id}")
    definitions(31) = Array("CONSTRAINTS.ERROR.ACTIVE_EMPTY_ID", "Active constraint row has an empty ID.", "Une ligne de contrainte active possède un ID vide.")
    definitions(32) = Array("CONSTRAINTS.ERROR.ID_NOT_IN_CALC", "Active constraint references ID not found in tbl_CALC: {Id}", "Une contrainte active référence un ID absent de tbl_CALC : {Id}")
    definitions(33) = Array("CONSTRAINTS.ERROR.MISSING_REQUIRED_COLUMN", "Missing required column in {Table}: {Column}", "Colonne obligatoire absente de {Table} : {Column}")
    definitions(34) = Array("CONSTRAINTS.VALIDATION.START_TYPE.LIST", _
        "Start No Earlier Than,Start No Later Than,Must Start On", _
        "Start No Earlier Than,Start No Later Than,Must Start On")
    definitions(35) = Array("CONSTRAINTS.VALIDATION.FINISH_TYPE.LIST", _
        "Finish No Earlier Than,Finish No Later Than,Must Finish On", _
        "Finish No Earlier Than,Finish No Later Than,Must Finish On")
    definitions(36) = Array("CONSTRAINTS.VALIDATION.ACTIVE.LIST", "Yes,No", "Yes,No")
    definitions(37) = Array("CONSTRAINTS.HELP.AUTO", _
        "Automatic projection" & vbLf & "- Identity and hierarchy come from WBS." & vbLf & "- Calculated dates, duration and driving logic come from the planning engine." & vbLf & vbLf & "Update the source or run the planning command; do not enter a value here.", _
        "Projection automatique" & vbLf & "- Identite et hierarchie issues de WBS." & vbLf & "- Dates calculees, duree et logique pilote issues du moteur de planning." & vbLf & vbLf & "Modifier la source ou actualiser le planning ; ne pas saisir de valeur ici.")
    definitions(38) = Array("CONSTRAINTS.HELP.START_CONSTRAINT_TYPE", _
        "Optional input" & vbLf & "Choose a start constraint from the list and enter its date in the adjacent column." & vbLf & "Leave both fields empty when no start constraint is needed.", _
        "Saisie facultative" & vbLf & "Choisir une contrainte de debut dans la liste et saisir sa date dans la colonne voisine." & vbLf & "Laisser les deux champs vides sans contrainte de debut.")
    definitions(39) = Array("CONSTRAINTS.HELP.START_CONSTRAINT_DATE", _
        "Optional input" & vbLf & "Required when a start constraint type is selected." & vbLf & "Enter an Excel date, for example 15/06/2026 in a day/month/year locale.", _
        "Saisie facultative" & vbLf & "Obligatoire si un type de contrainte de debut est choisi." & vbLf & "Saisir une date Excel, par exemple le 15/06/2026 avec un format jour/mois/annee.")
    definitions(40) = Array("CONSTRAINTS.HELP.FINISH_CONSTRAINT_TYPE", _
        "Optional input" & vbLf & "Choose a finish constraint from the list and enter its date in the adjacent column." & vbLf & "Leave both fields empty when no finish constraint is needed.", _
        "Saisie facultative" & vbLf & "Choisir une contrainte de fin dans la liste et saisir sa date dans la colonne voisine." & vbLf & "Laisser les deux champs vides sans contrainte de fin.")
    definitions(41) = Array("CONSTRAINTS.HELP.FINISH_CONSTRAINT_DATE", _
        "Optional input" & vbLf & "Required when a finish constraint type is selected." & vbLf & "Use a real Excel date, not free text.", _
        "Saisie facultative" & vbLf & "Obligatoire si un type de contrainte de fin est choisi." & vbLf & "Utiliser une vraie date Excel, pas un texte libre.")
    definitions(42) = Array("CONSTRAINTS.HELP.DEADLINE", _
        "Optional input" & vbLf & "Target date used to assess deadline float; it does not force the activity finish date." & vbLf & "Leave empty if no deadline is specified.", _
        "Saisie facultative" & vbLf & "Date cible utilisee pour evaluer la marge deadline ; elle ne force pas la fin de l'activite." & vbLf & "Laisser vide sans echeance definie.")
    definitions(43) = Array("CONSTRAINTS.HELP.ACTIVE", _
        "Optional input" & vbLf & "Yes: apply the row constraints. No: retain the inputs without applying them." & vbLf & "Summary and Level of Effort restrictions remain applicable.", _
        "Saisie facultative" & vbLf & "Yes : appliquer les contraintes de la ligne. No : conserver les saisies sans les appliquer." & vbLf & "Les restrictions des parents et activites Level of Effort restent applicables.")
    definitions(44) = Array("CONSTRAINTS.HELP.COMMENT", _
        "Optional input" & vbLf & "Explain the constraint or its assumption." & vbLf & "Example: equipment available after delivery.", _
        "Saisie facultative" & vbLf & "Expliquer la contrainte ou son hypothese." & vbLf & "Exemple : equipement disponible apres livraison.")
    TextCatalogConstraints_Definitions = definitions
End Function


