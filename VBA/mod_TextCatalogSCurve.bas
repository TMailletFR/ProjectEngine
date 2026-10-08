Attribute VB_Name = "mod_TextCatalogSCurve"
Option Explicit

Public Function TextCatalogSCurve_Definitions() As Variant
    Dim definitions(0 To 32) As Variant
    definitions(0) = Array("SCURVE.CHART.TITLE", "S-Curve", "S-Curve")
    definitions(1) = Array("SCURVE.SERIES.DAILY_ACTUAL", "Daily Actualized", "Réel journalier")
    definitions(2) = Array("SCURVE.SERIES.DAILY_FORECAST", "Daily Remaining Forecast", "Prévu restant journalier")
    definitions(3) = Array("SCURVE.SERIES.BASELINE", "Baseline", "Référence")
    definitions(4) = Array("SCURVE.SERIES.CALCULATED", "Calculated", "Calculé")
    definitions(5) = Array("SCURVE.SERIES.FORECAST", "Forecast", "Prévu")
    definitions(6) = Array("SCURVE.SERIES.ACTUAL", "Actual", "Réel")
    definitions(7) = Array("SCURVE.DIAG.NO_VALID_LEAF", "No valid leaf task available for S-curve.", "Aucune tâche feuille exploitable pour la S-curve.")
    definitions(8) = Array("SCURVE.DIAG.INVALID_TOTAL_WEIGHT", "The total usable S-curve weight is zero or invalid.", "La somme des poids exploités pour la S-curve est nulle ou invalide.")
    definitions(9) = Array("SCURVE.DIAG.NO_USABLE_DATE", "No usable date found to generate the S-curve.", "Aucune date exploitable pour générer la S-curve.")
    definitions(10) = Array("SCURVE.DIAG.RUN_ERROR", _
        "VBA error in Run_SCurve_Engine" & vbCrLf & "-> check the last edited block in mod_SCurve" & vbCrLf & "-> {Details}", _
        "Erreur VBA dans Run_SCurve_Engine" & vbCrLf & "-> vérifier le dernier bloc modifié dans mod_SCurve" & vbCrLf & "-> {Details}")
    definitions(11) = Array("SCURVE.DIAG.BLOCKING_GROUP", _
        "Blocking data for S-curve" & vbCrLf & "-> fix required fields before recalculation" & vbCrLf & vbCrLf & "IDs: {Ids}" & vbCrLf & "WBS: {Wbs}", _
        "Données bloquantes pour S-curve" & vbCrLf & "-> corriger les champs nécessaires avant recalcul" & vbCrLf & vbCrLf & "IDs : {Ids}" & vbCrLf & "WBS : {Wbs}")
    definitions(12) = Array("SCURVE.DIAG.MISSING_WEIGHT_GROUP", _
        "Missing weight on some leaf tasks - excluded from S-curve" & vbCrLf & "-> fill Weight (%) if needed" & vbCrLf & vbCrLf & "IDs: {Ids}" & vbCrLf & "WBS: {Wbs}", _
        "Poids manquant sur certaines tâches feuilles - non prises en compte dans la S-curve" & vbCrLf & "-> compléter Weight (%) si nécessaire" & vbCrLf & vbCrLf & "IDs : {Ids}" & vbCrLf & "WBS : {Wbs}")
    definitions(13) = Array("SCURVE.DIAG.EXCLUDED_WEIGHT_GROUP", _
        "Weight entered on Milestones or Level of Effort - excluded from S-curve" & vbCrLf & "-> remove Weight (%) if you want to avoid this warning" & vbCrLf & vbCrLf & "IDs: {Ids}" & vbCrLf & "WBS: {Wbs}", _
        "Poids renseigné sur des Milestones ou Level of Effort - non pris en compte dans la S-curve" & vbCrLf & "-> supprimer le Weight (%) si vous voulez éviter ce warning" & vbCrLf & vbCrLf & "IDs : {Ids}" & vbCrLf & "WBS : {Wbs}")
    definitions(14) = Array("SCURVE.EVENT.MISSING_WEIGHT.MESSAGE", "Missing weight on some leaf tasks", "Poids manquant sur certaines taches feuilles")
    definitions(15) = Array("SCURVE.EVENT.MISSING_WEIGHT.DETAIL", "tasks are excluded from the S-curve; fill Weight (%) if needed", "les taches sont exclues de la S-curve ; completer Weight (%) si necessaire")
    definitions(16) = Array("SCURVE.EVENT.EXCLUDED_WEIGHT.MESSAGE", "Weight entered on Milestones or Level of Effort", "Poids renseigne sur des Milestones ou Level of Effort")
    definitions(17) = Array("SCURVE.EVENT.EXCLUDED_WEIGHT.DETAIL", "tasks are excluded from the S-curve; remove Weight (%) if you want to avoid this warning", "les taches sont exclues de la S-curve ; supprimer Weight (%) si vous voulez eviter ce warning")
    definitions(18) = Array("SCURVE.ERROR.MISSING_SOURCE_COLUMN", "Missing source column in {Table}: {Column}", "Colonne source absente de {Table} : {Column}")
    definitions(19) = Array("SCURVE.ERROR.MISSING_OUTPUT_COLUMN", "Missing column in {Table}: {Column}", "Colonne absente de {Table} : {Column}")
    definitions(20) = Array("SCURVE.ERROR.INVALID_SIGNATURE_KEY", "Invalid fixed-length S-Curve chart signature key.", "Clé de signature fixe du graphique S-Curve invalide.")
    definitions(21) = Array("SCURVE.ERROR.MISSING_TABLE", "Required S-Curve table is missing: {Table}", "Table S-Curve obligatoire absente : {Table}")
    definitions(22) = Array("SCURVE.ERROR.MISSING_CHART", "Required S-Curve chart is missing: {Chart}", "Graphique S-Curve obligatoire absent : {Chart}")
    definitions(23) = Array("SCURVE.ROW.WARNING.TASK_MISSING", "TASK MISSING IN CALC", "TÂCHE ABSENTE DE CALC")
    definitions(24) = Array("SCURVE.ROW.WARNING.SUMMARY_EXCLUDED", "SUMMARY TASK EXCLUDED", "TÂCHE PARENTE EXCLUE")
    definitions(25) = Array("SCURVE.ROW.WARNING.TYPE_EXCLUDED", "TASK TYPE EXCLUDED FROM S-CURVE", "TYPE DE TÂCHE EXCLU DE LA S-CURVE")
    definitions(26) = Array("SCURVE.ROW.WARNING.MISSING_WEIGHT", "MISSING WEIGHT - EXCLUDED", "POIDS MANQUANT - EXCLU")
    definitions(27) = Array("SCURVE.ROW.WARNING.INVALID_WEIGHT", "INVALID WEIGHT", "POIDS INVALIDE")
    definitions(28) = Array("SCURVE.ROW.WARNING.NON_POSITIVE_WEIGHT", "NON-POSITIVE WEIGHT", "POIDS NON POSITIF")
    definitions(29) = Array("SCURVE.ROW.WARNING.MISSING_BASELINE", "MISSING BASELINE DATA", "DONNÉES BASELINE MANQUANTES")
    definitions(30) = Array("SCURVE.ROW.WARNING.INVALID_BASELINE_DURATION", "INVALID BASELINE DURATION", "DURÉE BASELINE INVALIDE")
    definitions(31) = Array("SCURVE.ROW.WARNING.MISSING_CALCULATED_DATES", "MISSING CALCULATED DATES", "DATES CALCULÉES MANQUANTES")
    definitions(32) = Array("SCURVE.ROW.WARNING.FINISH_BEFORE_START", "CALCULATED FINISH BEFORE START", "FIN CALCULÉE ANTÉRIEURE AU DÉBUT")
    TextCatalogSCurve_Definitions = definitions
End Function
