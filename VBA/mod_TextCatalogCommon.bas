Attribute VB_Name = "mod_TextCatalogCommon"
Option Explicit

Public Const TXT_COMMON_RESET_PLANNING_LABEL As String = "COMMON.RESET_PLANNING.LABEL"
Public Const TXT_COMMON_RESET_PLANNING_CONFIRMATION As String = "COMMON.RESET_PLANNING.CONFIRMATION"

Public Function TextCatalogCommon_Definitions() As Variant
    Dim definitions(0 To 43) As Variant

    definitions(0) = Array(TXT_COMMON_RESET_PLANNING_LABEL, "Reset Planning", "Réinitialiser planning")
    definitions(1) = Array(TXT_COMMON_RESET_PLANNING_CONFIRMATION, _
        "This will clear the WBS and clean planning calculation outputs." & vbCrLf & _
        "Gantt, S-Curve, Dashboard, history and acknowledgements will not be modified." & vbCrLf & vbCrLf & _
        "Continue?", _
        "Cette action va vider le WBS et nettoyer les sorties calcul planning." & vbCrLf & _
        "Gantt, S-Curve, Dashboard, historique et acknowledgements ne seront pas modifies." & vbCrLf & vbCrLf & _
        "Continuer ?")
    definitions(2) = Array("COMMON.ERROR.PROCEDURE_INLINE", "Error in {Procedure}: {Details}", "Erreur dans {Procedure} : {Details}")
    definitions(3) = Array("COMMON.ERROR.GANTT_RESET", "Error while resetting the Gantt: {Details}", "Erreur pendant la réinitialisation du Gantt : {Details}")
    definitions(4) = Array("COMMON.ERROR.INVALID_DEADLINE", _
        "Invalid deadline." & vbCrLf & "-> enter a valid date or leave blank.", _
        "Deadline invalide." & vbCrLf & "-> saisir une date valide ou laisser vide.")
    definitions(5) = Array("COMMON.ERROR.OWNER_LANGUAGE", "The {Owner} language was not changed. {Details}", "La langue {Owner} n'a pas ete modifiee. {Details}")
    definitions(6) = Array("COMMON.ERROR.GLOBAL_LANGUAGE", "The GLOBAL language change failed on {Owner}. {Details}", "Le changement GLOBAL a echoue sur {Owner}. {Details}")
    definitions(7) = Array("COMMON.ERROR.DATE_MODE", "Unknown date display mode: {Mode}", "Mode de format de date inconnu : {Mode}")
    definitions(8) = Array("COMMON.ERROR.DATE_FORMAT", "The date format was not changed. {Details}", "Le format des dates n'a pas ete modifie. {Details}")
    definitions(9) = Array("COMMON.LABEL.ANALYTICS", "Analytics", "Analytics")
    definitions(10) = Array("COMMON.STATUS.NO_PROJECT_DATA", "No project data - calculation outputs cleared.", "Aucune donnée projet - sorties de calcul nettoyées.")
    definitions(11) = Array("VISIBLE_HEADERS.ERROR.LANGUAGE_MISMATCH", "Table '{Table}' is physically '{Actual}' but expected '{Expected}'.", "La table '{Table}' est physiquement en '{Actual}' mais '{Expected}' était attendu.")
    definitions(12) = Array("VISIBLE_HEADERS.ERROR.PROTECTED_SHEET", "Worksheet '{Sheet}' is protected; table '{Table}' headers cannot be renamed.", "La feuille '{Sheet}' est protégée ; les en-têtes de la table '{Table}' ne peuvent pas être renommés.")
    definitions(13) = Array("VISIBLE_HEADERS.ERROR.PERSISTED_LANGUAGE_MISMATCH", "Physical header language mismatch for table '{Table}'. Persisted owner language is '{Persisted}' but headers are '{Actual}'.", "La langue physique des en-têtes de la table '{Table}' ne correspond pas. La langue persistée de l'owner est '{Persisted}' mais les en-têtes sont en '{Actual}'.")
    definitions(14) = Array("VISIBLE_HEADERS.ERROR.COLUMN_COUNT", "Table '{Table}' has {Actual} columns; expected {Expected}.", "La table '{Table}' possède {Actual} colonnes ; {Expected} étaient attendues.")
    definitions(15) = Array("VISIBLE_HEADERS.ERROR.DUPLICATE_PHYSICAL_HEADER", "Duplicate physical header '{Header}' in table '{Table}'.", "En-tête physique '{Header}' dupliqué dans la table '{Table}'.")
    definitions(16) = Array("VISIBLE_HEADERS.ERROR.UNKNOWN_PHYSICAL_HEADER", "Unknown physical header in table '{Table}' at column {Column}. Observed '{Observed}'. Expected EN '{English}' or FR '{French}'.", "En-tête physique inconnu dans la table '{Table}' à la colonne {Column}. Valeur observée '{Observed}'. Valeur attendue EN '{English}' ou FR '{French}'.")
    definitions(17) = Array("VISIBLE_HEADERS.ERROR.DUPLICATE_SCHEMA_KEY", "Duplicate schema column key '{Key}' in table '{Table}'.", "Clé de colonne de schéma '{Key}' dupliquée dans la table '{Table}'.")
    definitions(18) = Array("VISIBLE_HEADERS.ERROR.MISSING_SCHEMA_KEY", "Missing schema column key '{Key}' in table '{Table}'.", "Clé de colonne de schéma '{Key}' absente de la table '{Table}'.")
    definitions(19) = Array("VISIBLE_HEADERS.ERROR.MIXED_LANGUAGE", "Mixed EN/FR physical headers in table '{Table}'.", "En-têtes physiques EN/FR mélangés dans la table '{Table}'.")
    definitions(20) = Array("VISIBLE_HEADERS.ERROR.UNMAPPED_HEADER", "Unable to map physical header '{Header}' in table '{Table}' to language '{Language}'.", "Impossible d'associer l'en-tête physique '{Header}' de la table '{Table}' à la langue '{Language}'.")
    definitions(21) = Array("VISIBLE_HEADERS.ERROR.TABLE_NOT_FOUND", "Visible table '{Table}' was not found.", "Table visible '{Table}' introuvable.")
    definitions(22) = Array("VISIBLE_HEADERS.ERROR.UNKNOWN_OWNER", "Unknown language owner '{Owner}'.", "Owner de langue inconnu '{Owner}'.")
    definitions(23) = Array("VISIBLE_HEADERS.ERROR.UNSUPPORTED_LANGUAGE", "Unsupported target language '{Language}'. Expected EN or FR.", "Langue cible '{Language}' non prise en charge. EN ou FR attendu.")
    definitions(24) = Array("VISIBLE_HEADERS.ERROR.TEMPORARY_RESIDUAL", "Temporary header residual '{Header}' remains in table '{Table}'.", "L'en-tête temporaire résiduel '{Header}' subsiste dans la table '{Table}'.")
    definitions(25) = Array("VISIBLE_HEADERS.ERROR.UNEXPECTED_HEADER_ROW", "Unexpected physical header row for table '{Table}'. Found {Actual}, expected {Legacy} before migration or {Migrated} after migration.", "Ligne d'en-tête physique inattendue pour la table '{Table}'. Ligne {Actual} trouvée ; {Legacy} attendue avant migration ou {Migrated} après migration.")
    definitions(26) = Array("VISIBLE_HEADERS.ERROR.LEGACY_ROW_POSITION", "The candidate legacy row is not immediately above table '{Table}'.", "La ligne legacy candidate n'est pas située juste au-dessus de la table '{Table}'.")
    definitions(27) = Array("VISIBLE_HEADERS.ERROR.LEGACY_COLUMN_COUNT", "Unexpected column count in table '{Table}'.", "Nombre de colonnes inattendu dans la table '{Table}'.")
    definitions(28) = Array("VISIBLE_HEADERS.ERROR.LEGACY_MERGED_CELLS", "Merged cells exist in the candidate legacy row for table '{Table}'.", "La ligne legacy candidate de la table '{Table}' contient des cellules fusionnées.")
    definitions(29) = Array("VISIBLE_HEADERS.ERROR.LEGACY_LABEL_COUNT", "The candidate legacy row for table '{Table}' does not contain exactly one label per visible column.", "La ligne legacy candidate de la table '{Table}' ne contient pas exactement un libellé par colonne visible.")
    definitions(30) = Array("VISIBLE_HEADERS.ERROR.LEGACY_OUTSIDE_DATA", "Unexpected data exists {Position} the legacy labels on sheet '{Sheet}'.", "Des données inattendues existent {Position} les libellés legacy sur la feuille '{Sheet}'.")
    definitions(31) = Array("VISIBLE_HEADERS.ERROR.LEGACY_UNKNOWN_CONTENT", "Unknown content '{Content}' in the candidate legacy row for table '{Table}' at column {Column}.", "Contenu inconnu '{Content}' dans la ligne legacy candidate de la table '{Table}' à la colonne {Column}.")
    definitions(32) = Array("VISIBLE_HEADERS.ERROR.LEGACY_SHAPE_INTERSECTION", "Shape '{Shape}' intersects the candidate legacy row on sheet '{Sheet}'.", "La shape '{Shape}' recoupe la ligne legacy candidate sur la feuille '{Sheet}'.")
    definitions(33) = Array("VISIBLE_HEADERS.ERROR.POST_DELETE_VALIDATION", "Post-delete structure validation failed for table '{Table}'.", "Échec de la validation structurelle après suppression pour la table '{Table}'.")
    definitions(34) = Array("VISIBLE_SCHEMA.ERROR.UNSUPPORTED_OVERRIDE_LANGUAGE", "Unsupported transient schema language '{Language}' for table '{Table}'. Expected EN or FR.", "Langue de schéma transitoire '{Language}' non prise en charge pour la table '{Table}'. EN ou FR attendu.")
    definitions(35) = Array("VISIBLE_SCHEMA.ERROR.UNKNOWN_TABLE_COLUMN", "Unknown visible-table key '{Table}' for column key '{Column}'.", "Clé de table visible '{Table}' inconnue pour la clé de colonne '{Column}'.")
    definitions(36) = Array("VISIBLE_SCHEMA.ERROR.UNKNOWN_COLUMN", "Unknown visible-table column key '{Column}' for table '{Table}'.", "Clé de colonne visible '{Column}' inconnue pour la table '{Table}'.")
    definitions(37) = Array("VISIBLE_SCHEMA.ERROR.LISTOBJECT_NOTHING", "The ListObject is Nothing for table key '{Table}'.", "Le ListObject est Nothing pour la clé de table '{Table}'.")
    definitions(38) = Array("VISIBLE_SCHEMA.ERROR.LISTOBJECT_MISMATCH", "ListObject mismatch. Expected '{Expected}' but received '{Actual}'.", "ListObject incorrect. '{Expected}' attendu, '{Actual}' reçu.")
    definitions(39) = Array("VISIBLE_SCHEMA.ERROR.COLUMN_NOT_RESOLVED", "The column key '{Column}' could not be resolved in table '{Table}'. Expected physical header: '{Header}'.", "La clé de colonne '{Column}' n'a pas pu être résolue dans la table '{Table}'. En-tête physique attendu : '{Header}'.")
    definitions(40) = Array("VISIBLE_SCHEMA.ERROR.UNSUPPORTED_LANGUAGE", "Unsupported schema language '{Language}' for table '{Table}' and column key '{Column}'. Expected EN or FR.", "Langue de schéma '{Language}' non prise en charge pour la table '{Table}' et la clé de colonne '{Column}'. EN ou FR attendu.")
    definitions(41) = Array("VISIBLE_SCHEMA.ERROR.UNKNOWN_TABLE", "Unknown visible-table key '{Table}'.", "Clé de table visible '{Table}' inconnue.")
    definitions(42) = Array("COMMON.POSITION.BEFORE", "before", "avant")
    definitions(43) = Array("COMMON.POSITION.AFTER", "after", "après")

    TextCatalogCommon_Definitions = definitions
End Function
