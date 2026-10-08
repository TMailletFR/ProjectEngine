Attribute VB_Name = "mod_TextCatalogRibbon"
Option Explicit

' Definitions only: the shared TextCatalog owns lookup and fallback.
Public Const TXT_RIBBON_TAB_LABEL As String = "RIBBON.TAB.LABEL"
Public Const TXT_RIBBON_UPDATE_LABEL As String = "RIBBON.UPDATE.LABEL"
Public Const TXT_RIBBON_PLANNING_LABEL As String = "RIBBON.PLANNING.LABEL"
Public Const TXT_RIBBON_PLANNING_TIP As String = "RIBBON.PLANNING.SCREENTIP"
Public Const TXT_RIBBON_PLANNING_HELP As String = "RIBBON.PLANNING.SUPERTIP"
Public Const TXT_RIBBON_FULL_LABEL As String = "RIBBON.FULL.LABEL"
Public Const TXT_RIBBON_FULL_TIP As String = "RIBBON.FULL.SCREENTIP"
Public Const TXT_RIBBON_FULL_HELP As String = "RIBBON.FULL.SUPERTIP"
Public Const TXT_RIBBON_GANTT_LABEL As String = "RIBBON.GANTT.LABEL"
Public Const TXT_RIBBON_GANTT_TIP As String = "RIBBON.GANTT.SCREENTIP"
Public Const TXT_RIBBON_GANTT_HELP As String = "RIBBON.GANTT.SUPERTIP"
Public Const TXT_RIBBON_SCURVE_LABEL As String = "RIBBON.SCURVE.LABEL"
Public Const TXT_RIBBON_SCURVE_TIP As String = "RIBBON.SCURVE.SCREENTIP"
Public Const TXT_RIBBON_SCURVE_HELP As String = "RIBBON.SCURVE.SUPERTIP"
Public Const TXT_RIBBON_PLANNING_TOOLS_LABEL As String = "RIBBON.PLANNING_TOOLS.LABEL"
Public Const TXT_RIBBON_RESET_TIP As String = "RIBBON.RESET_PLANNING.SCREENTIP"
Public Const TXT_RIBBON_RESET_HELP As String = "RIBBON.RESET_PLANNING.SUPERTIP"

Public Function TextCatalogRibbon_Definitions() As Variant
    Dim definitions As Variant, additions As Variant, combined() As Variant
    Dim i As Long, offset As Long
    definitions = Array( _
        Array(TXT_RIBBON_TAB_LABEL, "ProjectEngine", "ProjectEngine"), _
        Array(TXT_RIBBON_UPDATE_LABEL, "WBS", "WBS"), _
        Array(TXT_RIBBON_PLANNING_LABEL, "Update Planning", "Actualiser le planning"), _
        Array(TXT_RIBBON_PLANNING_TIP, "Recalculate the planning", "Recalculer le planning"), _
        Array(TXT_RIBBON_PLANNING_HELP, "Run the same planning update as the WBS button in this workbook.", "Exécuter la même mise à jour du planning que le bouton WBS de ce classeur."), _
        Array(TXT_RIBBON_FULL_LABEL, "Full Update", "Mise à jour complète"), _
        Array(TXT_RIBBON_FULL_TIP, "Run a full update", "Exécuter une mise à jour complète"), _
        Array(TXT_RIBBON_FULL_HELP, "Run the existing full update workflow, including planning and S-Curve. Gantt rendering follows the current sheet context.", "Exécuter le workflow complet existant, incluant planning et S-Curve. Le rendu Gantt suit le contexte de la feuille courante."), _
        Array(TXT_RIBBON_GANTT_LABEL, "Update Gantt", "Actualiser le Gantt"), _
        Array(TXT_RIBBON_GANTT_TIP, "Recalculate and display the Gantt", "Recalculer et afficher le Gantt"), _
        Array(TXT_RIBBON_GANTT_HELP, "Use the existing Gantt update workflow. Temporary simulation state is reset as with the WBS button.", "Utiliser le workflow de mise à jour Gantt existant. L'état temporaire de simulation est réinitialisé comme avec le bouton WBS."), _
        Array(TXT_RIBBON_SCURVE_LABEL, "Update S-Curve", "Actualiser la S-Curve"), _
        Array(TXT_RIBBON_SCURVE_TIP, "Update and display the S-Curve", "Actualiser et afficher la S-Curve"), _
        Array(TXT_RIBBON_SCURVE_HELP, "Run the existing S-Curve workflow using this workbook's planning data.", "Exécuter le workflow S-Curve existant avec les données de planning de ce classeur."), _
        Array(TXT_RIBBON_PLANNING_TOOLS_LABEL, "Planning", "Planning"), _
        Array(TXT_RIBBON_RESET_TIP, "Clear planning inputs after confirmation", "Vider les données du planning après confirmation"), _
        Array(TXT_RIBBON_RESET_HELP, "Clear WBS inputs and planning calculation outputs using the existing Reset Planning command. Confirmation is required; No is selected by default.", "Vider les données WBS et les sorties du calcul planning avec la commande de réinitialisation existante. Une confirmation est obligatoire ; Non est sélectionné par défaut."))
    additions = Array( _
        Array("RIBBON.FORCED.LABEL", "Forced Planning Update", "Actualisation forcée du planning"), _
        Array("RIBBON.FORCED.TIP", "Force the existing planning recalculation workflow", "Forcer le recalcul du planning existant"), _
        Array("RIBBON.RESET_GROUP.LABEL", "Reset", "Réinitialisation"), _
        Array("RIBBON.SETTINGS.LABEL", "Settings", "Paramètres"), _
        Array("RIBBON.IMPORT.LABEL", "Import project", "Importer un projet"), _
        Array("RIBBON.IMPORT.TIP", "Import data from a previous ProjectEngine workbook", "Importer les données d'un ancien classeur ProjectEngine"), _
        Array("RIBBON.SETTINGS.TIP", "Open this workbook's settings", "Ouvrir les paramètres de ce classeur"), _
        Array("RIBBON.GANTT_CONTEXT.LABEL", "Gantt", "Gantt"), _
        Array("RIBBON.CLEAR_HISTORY.TIP", "Clear the planning event history", "Effacer l'historique des événements du planning"), _
        Array("RIBBON.CLEAR_ACK.TIP", "Clear warning acknowledgements", "Effacer les acquittements des avertissements"), _
        Array("RIBBON.CLEAN_DASHBOARD.TIP", "Clear Dashboard snapshots and reset its display", "Effacer les instantanés du Dashboard et réinitialiser son affichage"), _
        Array("RIBBON.FULL_RESET.TIP", "Reset this workbook completely after confirmation", "Réinitialiser entièrement ce classeur après confirmation"), _
        Array("RIBBON.GANTT_RESET.TIP", "Discard simulation assumptions and return to the calculated planning", "Abandonner les hypothèses de simulation et revenir au planning calculé"), _
        Array("RIBBON.SCENARIO.TIP", "Run the existing Gantt Scenario workflow", "Exécuter le scénario Gantt existant"), _
        Array("RIBBON.TEST.TIP", "Run the existing Gantt Test workflow", "Exécuter le test Gantt existant"), _
        Array("RIBBON.LOCK.TIP", "Apply the current Gantt assumptions to the planning", "Appliquer les hypothèses Gantt actuelles au planning"))
    offset = UBound(definitions) + 1
    ReDim combined(0 To offset + UBound(additions))
    For i = 0 To UBound(definitions)
        combined(i) = definitions(i)
    Next i
    For i = 0 To UBound(additions)
        combined(offset + i) = additions(i)
    Next i
    TextCatalogRibbon_Definitions = combined
End Function
