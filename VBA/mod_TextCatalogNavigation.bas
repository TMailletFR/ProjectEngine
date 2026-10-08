Attribute VB_Name = "mod_TextCatalogNavigation"
Option Explicit

Public Function TextCatalogNavigation_Definitions() As Variant
    TextCatalogNavigation_Definitions = Array( _
        Array("NAV.GROUP.LABEL", "Navigation", "Navigation"), _
        Array("NAV.FULL.LABEL", "Show Full Timeline", "Afficher toute la chronologie"), _
        Array("NAV.FULL.TIP", "Fit the rendered timeline", "Cadrer la chronologie affichée"), _
        Array("NAV.FULL.HELP", "Fit the current timeline using Excel zoom, without recalculating planning or changing Day/Week/Month. At Excel's zoom limit, enlarge the window or reduce the visible extent. Week/Month reduces width, not the number of task rows.", "Cadrer la chronologie avec le zoom Excel, sans recalculer le planning ni changer Jour/Semaine/Mois. A la limite de zoom Excel, agrandir la fenêtre ou réduire l'étendue visible. Semaine/Mois réduit la largeur, pas le nombre de lignes de tâches."), _
        Array("NAV.TASK.LABEL", "Go to selected task", "Aller à la tâche sélectionnée"), _
        Array("NAV.TASK.TIP", "Find the selected task in Gantt", "Retrouver la tâche sélectionnée dans le Gantt"), _
        Array("NAV.TASK.HELP", "Select cells on one task row in WBS or the Gantt left panel. Center its rendered bar and adjust zoom only when needed, without recalculating planning or changing Day/Week/Month. Hidden tasks remain hidden.", "Sélectionner des cellules sur une ligne de tâche dans WBS ou le panneau gauche du Gantt. Centrer sa barre et ajuster le zoom seulement si nécessaire, sans recalculer le planning ni changer Jour/Semaine/Mois. Les tâches masquées restent masquées."), _
        Array("NAV.ERROR.INVALID_CONTEXT", "Navigation requires this workbook's active window.", "La navigation nécessite la fenêtre active de ce classeur."), _
        Array("NAV.ERROR.EMPTY", "No calculated timeline is available. Update Planning first.", "Aucune chronologie calculée disponible. Actualiser d'abord le planning."), _
        Array("NAV.ERROR.SELECT_ONE_TASK", "Select one task row in WBS or Gantt, then try again.", "Sélectionnez une seule ligne de tâche dans WBS ou Gantt, puis réessayez."), _
        Array("NAV.ERROR.NOT_READY", "The Gantt could not be prepared. Reopen the workbook, then try again.", "L'affichage du Gantt n'a pas pu être préparé. Rouvrez le classeur, puis réessayez."), _
        Array("NAV.ERROR.PLANNING_STALE", "Update Planning before navigating in the Gantt.", "Actualisez le planning avant de naviguer dans le Gantt."), _
        Array("NAV.ERROR.PREPARATION_FAILED", "The Gantt view could not be prepared. Update Gantt, then try again.", "L'affichage du Gantt n'a pas pu être préparé. Actualisez le Gantt, puis réessayez."), _
        Array("NAV.ERROR.TASK_HIDDEN", "This task is hidden in the current Gantt view. Adjust the view or filter first.", "Cette tâche est masquée dans la vue Gantt courante. Ajuster d'abord la vue ou le filtre."), _
        Array("NAV.ERROR.TASK_NOT_RENDERED", "No rendered bar is available for this task. Check its calculated dates and the Gantt view.", "Aucune barre disponible pour cette tâche. Vérifier ses dates calculées et la vue Gantt."), _
        Array("NAV.ERROR.ZOOM_LIMIT", "The rendered timeline cannot fit entirely within this window at Excel's zoom limits. Enlarge the window or reduce the visible extent. Week/Month reduces width, not the number of task rows.", "La chronologie affichée ne tient pas entièrement dans cette fenêtre aux limites de zoom Excel. Agrandir la fenêtre ou réduire l'étendue visible. Semaine/Mois réduit la largeur, pas le nombre de lignes de tâches."), _
        Array("NAV.ERROR.FAILED", "The selected task could not be displayed in the Gantt. Reopen the workbook; if it persists, report an issue on the project's GitHub repository.", "Impossible d'afficher la tâche sélectionnée dans le Gantt. Rouvrez le classeur ; si le problème persiste, signalez-le dans une issue sur le dépôt GitHub du projet."), _
        Array("NAV.ERROR.FAILED_FULL", "The full Gantt view could not be displayed. Reopen the workbook; if it persists, report an issue on the project's GitHub repository.", "Impossible d'afficher toute la vue Gantt. Rouvrez le classeur ; si le problème persiste, signalez-le dans une issue sur le dépôt GitHub du projet."))
End Function
