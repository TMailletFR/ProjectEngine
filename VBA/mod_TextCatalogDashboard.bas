Attribute VB_Name = "mod_TextCatalogDashboard"
Option Explicit

Public Function TextCatalogDashboard_Definitions() As Variant
    Dim definitions(0 To 107) As Variant
    definitions(0) = Array("DASHBOARD.ERROR.COMPARISON", _
        "Error in Refresh_Dashboard_Comparison" & vbCrLf & "{Details}", _
        "Erreur dans Refresh_Dashboard_Comparison" & vbCrLf & "{Details}")
    definitions(1) = Array("DASHBOARD.TITLE.EXECUTIVE_SUMMARY", _
        "Executive Summary", _
        "Synthèse exécutive")
    definitions(2) = Array("DASHBOARD.KPI.PROGRESS", _
        "Project Progress", _
        "Avancement projet")
    definitions(3) = Array("DASHBOARD.KPI.FORECAST_FINISH", _
        "Forecast Finish", _
        "Fin prévisionnelle")
    definitions(4) = Array("DASHBOARD.KPI.CRITICAL_ACTIVITIES", _
        "Critical Activities", _
        "Activités critiques")
    definitions(5) = Array("DASHBOARD.KPI.MOMENTUM", _
        "Schedule Momentum", _
        "Momentum planning")
    definitions(6) = Array("DASHBOARD.SECTION.SCURVE", _
        "S-Curve Snapshot", _
        "Snapshot S-Curve")
    definitions(7) = Array("DASHBOARD.SECTION.OVERVIEW", _
        "Planning Overview", _
        "Vue planning")
    definitions(8) = Array("DASHBOARD.SECTION.HOT_SPOTS", _
        "Hot Spots", _
        "Points chauds")
    definitions(9) = Array("DASHBOARD.CARD.TOP_DELAYS", _
        "Top Delays", _
        "Dérives majeures")
    definitions(10) = Array("DASHBOARD.CARD.DEADLINES", _
        "Deadline Health", _
        "Santé deadlines")
    definitions(11) = Array("DASHBOARD.CARD.NEXT_MILESTONE", _
        "Next Milestone", _
        "Prochain jalon")
    definitions(12) = Array("DASHBOARD.CARD.NEXT_CRITICAL", _
        "Next Critical Activity", _
        "Activité critique")
    definitions(13) = Array("DASHBOARD.TITLE.PAGE", _
        "Dashboard", _
        "Tableau de bord")
    definitions(14) = Array("DASHBOARD.TITLE.SUBTITLE", _
        "PM / Engineering dashboard - comparison snapshots", _
        "Pilotage projet PM / Engineering - snapshot de comparaison")
    definitions(15) = Array("DASHBOARD.COMMAND.SNAPSHOT", _
        "New snapshot", _
        "Nouveau snapshot")
    definitions(16) = Array("DASHBOARD.COMMAND.COMPARE", _
        "Refresh Comparison", _
        "Rafraîchir comparaison")
    definitions(17) = Array("DASHBOARD.COMMAND.CLEAN", _
        "Clean Dashboard", _
        "Nettoyer Dashboard")
    definitions(18) = Array("DASHBOARD.CARD.DEADLINE_RISKS", _
        "Deadline Risks", _
        "Risques jalons")
    definitions(19) = Array("DASHBOARD.CARD.FORECAST_ISSUES", _
        "Forecast Issues", _
        "Alertes forecast")
    definitions(20) = Array("DASHBOARD.EMPTY.NO_DELAYS", _
        "No schedule delays detected", _
        "Aucune dérive planning détectée")
    definitions(21) = Array("DASHBOARD.EMPTY.NO_DEADLINE_RISKS", _
        "No deadline risks detected", _
        "Aucun risque deadline détecté")
    definitions(22) = Array("DASHBOARD.EMPTY.NO_FORECAST_ISSUES", _
        "No forecast issues detected", _
        "Aucune alerte forecast détectée")
    definitions(23) = Array("DASHBOARD.SNAPSHOT.FROM", _
        "From", _
        "De")
    definitions(24) = Array("DASHBOARD.SNAPSHOT.TO", _
        "To", _
        "A")
    definitions(25) = Array("DASHBOARD.SERIES.BASELINE", _
        "Baseline", _
        "Référence")
    definitions(26) = Array("DASHBOARD.SERIES.ACTUAL", _
        "Actual", _
        "Réel")
    definitions(27) = Array("DASHBOARD.SERIES.FORECAST", _
        "Forecast", _
        "Prévision")
    definitions(28) = Array("DASHBOARD.EMPTY.NO_SUMMARY", _
        "No summary schedule data available", _
        "Aucune donnée planning summary disponible")
    definitions(29) = Array("DASHBOARD.HELP.DELAYS", _
        "Where is the delay?", _
        "Où est le retard ?")
    definitions(30) = Array("DASHBOARD.EMPTY.NO_DELAY_DATA", _
        "No delay data available", _
        "Aucune donnée de dérive disponible")
    definitions(31) = Array("DASHBOARD.HELP.COMMITMENTS", _
        "Are commitments safe?", _
        "Mes engagements sont-ils tenus ?")
    definitions(32) = Array("DASHBOARD.STATUS.OVERDUE", _
        "Overdue", _
        "En retard")
    definitions(33) = Array("DASHBOARD.STATUS.ON_TRACK", _
        "On Track", _
        "OK")
    definitions(34) = Array("DASHBOARD.LABEL.WORST", _
        "Worst Offender", _
        "Plus critique")
    definitions(35) = Array("DASHBOARD.LABEL.NEAREST_RISK", _
        "Closest Risk", _
        "Risque proche")
    definitions(36) = Array("DASHBOARD.EMPTY.NO_ACTIVE_DEADLINE", _
        "No active deadline", _
        "Aucune deadline active")
    definitions(37) = Array("DASHBOARD.EMPTY.NO_DEADLINE_DATA", _
        "No deadline data available", _
        "Aucune deadline disponible")
    definitions(38) = Array("DASHBOARD.EMPTY.NO_ACTIVE_MILESTONE", _
        "No active milestone", _
        "Aucun jalon actif")
    definitions(39) = Array("DASHBOARD.EMPTY.NO_MILESTONE_DATA", _
        "No milestone data available", _
        "Aucun jalon disponible")
    definitions(40) = Array("DASHBOARD.EMPTY.NO_ACTIVE_CRITICAL", _
        "No active critical activity", _
        "Aucune activité critique active")
    definitions(41) = Array("DASHBOARD.EMPTY.NO_CRITICAL_DATA", _
        "No critical activity data available", _
        "Aucune activité critique disponible")
    definitions(42) = Array("DASHBOARD.LABEL.TODAY", _
        "today", _
        "aujourd'hui")
    definitions(43) = Array("DASHBOARD.COLUMN.TASK", _
        "Task", _
        "Tâche")
    definitions(44) = Array("DASHBOARD.COLUMN.ISSUE", _
        "Issue", _
        "Alerte")
    definitions(45) = Array("DASHBOARD.EMPTY.NO_SNAPSHOTS", _
        "No snapshots", _
        "Aucun snapshot")
    definitions(46) = Array("DASHBOARD.EMPTY.HISTORY_TITLE", _
        "Insufficient history", _
        "Historique insuffisant")
    definitions(47) = Array("DASHBOARD.EMPTY.COMPARISON", _
        "Comparison unavailable", _
        "Comparaison indisponible")
    definitions(48) = Array("DASHBOARD.EMPTY.DELTA", _
        "Delta unavailable", _
        "Delta indisponible")
    definitions(49) = Array("DASHBOARD.STATUS.AHEAD", _
        "ahead or on plan", _
        "a l'heure ou en avance")
    definitions(50) = Array("DASHBOARD.EMPTY.FORECAST_COMPARISON", _
        "Forecast not comparable", _
        "Fin non comparable")
    definitions(51) = Array("DASHBOARD.UNIT.DAY_COMPACT", _
        "d", _
        "j")
    definitions(52) = Array("DASHBOARD.STATUS.FORECAST_STABLE", _
        "Forecast stable vs From", _
        "Fin stable vs Début")
    definitions(53) = Array("DASHBOARD.EMPTY.CONTRACT_DRIFT", _
        "contract drift unavailable", _
        "dérive contrat indisponible")
    definitions(54) = Array("DASHBOARD.EMPTY.HISTORY_SENTENCE", _
        "insufficient history", _
        "historique insuffisant")
    definitions(55) = Array("DASHBOARD.STATUS.CRITICAL_UNCHANGED", _
        "0 critical activities change", _
        "0 changement activités critiques")
    definitions(56) = Array("DASHBOARD.EMPTY.SNAPSHOTS_REQUIRED", _
        "Need snapshots", _
        "Snapshots requis")
    definitions(57) = Array("DASHBOARD.STATUS.DELAY_UNCHANGED", _
        "0% delay change", _
        "0% évolution du retard")
    definitions(58) = Array("DASHBOARD.STATUS.ON_CONTRACT", _
        "ON CONTRACT", _
        "CONTRAT OK")
    definitions(59) = Array("DASHBOARD.STATUS.MINOR_DELAY", _
        "MINOR DELAY", _
        "RETARD MINEUR")
    definitions(60) = Array("DASHBOARD.STATUS.CONTRACT_DELAY", _
        "CONTRACT DELAY", _
        "RETARD CONTRAT")
    definitions(61) = Array("DASHBOARD.STATUS.IMPROVING", _
        "IMPROVING", _
        "AMÉLIORATION")
    definitions(62) = Array("DASHBOARD.STATUS.DETERIORATING", _
        "DETERIORATING", _
        "DÉGRADATION")
    definitions(63) = Array("DASHBOARD.STATUS.STABLE", _
        "STABLE", _
        "STABLE")
    definitions(64) = Array("DASHBOARD.STATUS.NO_HISTORY", _
        "NO HISTORY", _
        "PAS D'HIST.")
    definitions(65) = Array("DASHBOARD.UNIT.WEEK_PREFIX", _
        "W", _
        "S")
    definitions(66) = Array("DASHBOARD.UNIT.DAY_LONG", _
        " days", _
        " jours")
    definitions(67) = Array("DASHBOARD.EMPTY.NO_PROJECT", _
        "No project loaded", _
        "Aucun projet chargé")
    definitions(68) = Array("DASHBOARD.EMPTY.NO_DATA", _
        "NO DATA", _
        "AUCUNE DONNÉE")
    definitions(69) = Array("DASHBOARD.MONTH.JAN", _
        "Jan", _
        "janv")
    definitions(70) = Array("DASHBOARD.MONTH.FEB", _
        "Feb", _
        "fév")
    definitions(71) = Array("DASHBOARD.MONTH.MAR", _
        "Mar", _
        "mars")
    definitions(72) = Array("DASHBOARD.MONTH.APR", _
        "Apr", _
        "avr")
    definitions(73) = Array("DASHBOARD.MONTH.MAY", _
        "May", _
        "mai")
    definitions(74) = Array("DASHBOARD.MONTH.JUN", _
        "Jun", _
        "juin")
    definitions(75) = Array("DASHBOARD.MONTH.JUL", _
        "Jul", _
        "juil")
    definitions(76) = Array("DASHBOARD.MONTH.AUG", _
        "Aug", _
        "août")
    definitions(77) = Array("DASHBOARD.MONTH.SEP", _
        "Sep", _
        "sept")
    definitions(78) = Array("DASHBOARD.MONTH.OCT", _
        "Oct", _
        "oct")
    definitions(79) = Array("DASHBOARD.MONTH.NOV", _
        "Nov", _
        "nov")
    definitions(80) = Array("DASHBOARD.MONTH.DEC", _
        "Dec", _
        "déc")
    definitions(81) = Array("DASHBOARD.ERROR.UPDATE", _
        "Error in Run_Dashboard_Update" & vbCrLf & "-> {Details}", _
        "Erreur dans Run_Dashboard_Update" & vbCrLf & "-> {Details}")
    definitions(82) = Array("DASHBOARD.TASK.STARTS", _
        "Starts {Date}", _
        "Début {Date}")
    definitions(83) = Array("DASHBOARD.TASK.FLOAT", _
        "Float: {Days}", _
        "Marge: {Days}")
    definitions(84) = Array("DASHBOARD.ERROR.RENDER", _
        "Render error {Number} - {Details} | {Card} | {Helper} | {Step}", _
        "Erreur rendu {Number} - {Details} | {Card} | {Helper} | {Step}")
    definitions(85) = Array("DASHBOARD.DEADLINE.OVERDUE", _
        "{Count} days overdue", _
        "{Count}j de retard")
    definitions(86) = Array("DASHBOARD.DEADLINE.REMAINING", _
        "{Count} days remaining", _
        "{Count}j restants")
    definitions(87) = Array("DASHBOARD.MOMENTUM.SUMMARY", _
        "Progress {Progress} | Forecast {Forecast} | Risks {Risks}", _
        "Avancement {Progress} | Fin {Forecast} | Risques {Risks}")
    definitions(88) = Array("DASHBOARD.PROGRESS.DELTA", _
        "{Delta} vs From", _
        "{Delta} vs Début")
    definitions(89) = Array("DASHBOARD.PROGRESS.BEHIND", _
        "{Progress} behind plan", _
        "{Progress} de retard vs plan")
    definitions(90) = Array("DASHBOARD.FORECAST.IMPROVED", _
        "Forecast improved by {Count}d", _
        "Fin améliorée de {Count}j")
    definitions(91) = Array("DASHBOARD.FORECAST.SLIPPED", _
        "Forecast slipped by {Count}d", _
        "Fin décalée de {Count}j")
    definitions(92) = Array("DASHBOARD.CONTRACT.DRIFT", _
        "{Days}d vs contract", _
        "{Days}j vs contrat")
    definitions(93) = Array("DASHBOARD.CRITICAL.ADDED", _
        "+{Count} new critical activities", _
        "+{Count} nouvelles activités critiques")
    definitions(94) = Array("DASHBOARD.CRITICAL.REMOVED", _
        "{Count} critical activities removed", _
        "{Count} activités critiques retirées")
    definitions(95) = Array("DASHBOARD.MOMENTUM.DELAY", _
        "{Delta} delay vs previous snapshot", _
        "{Delta} retard vs snapshot precedent")
    definitions(96) = Array("DASHBOARD.STATUS.OK_DETAIL", _
        "OK - {Details}", _
        "OK - {Details}")
    definitions(97) = Array("DASHBOARD.TIMESTAMP.DISPLAY", _
        "Updated {Date} {Time}", _
        "Mis à jour {Date} {Time}")
    definitions(98) = Array("DASHBOARD.TIMESTAMP.PREFIX", _
        "Updated ", _
        "Mis à jour ")
    definitions(99) = Array("DASHBOARD.LANGUAGE.SWITCH", "FR / EN", "FR / EN")
    definitions(100) = Array("DASHBOARD.LEGACY_ALIAS.TOP_DELAYS", "Derives majeures", "Derives majeures")
    definitions(101) = Array("DASHBOARD.LEGACY_ALIAS.NO_DELAYS", "Aucune derive planning detectee", "Aucune derive planning detectee")
    definitions(102) = Array("DASHBOARD.LEGACY_ALIAS.NO_DEADLINE_RISKS", "Aucun risque deadline detecte", "Aucun risque deadline detecte")
    definitions(103) = Array("DASHBOARD.LEGACY_ALIAS.NO_FORECAST_ISSUES", "Aucune alerte forecast detectee", "Aucune alerte forecast detectee")
    definitions(104) = Array("DASHBOARD.LEGACY_ALIAS.NO_PROJECT", "Aucun projet charge", "Aucun projet charge")
    definitions(105) = Array("DASHBOARD.LEGACY_ALIAS.NO_DATA", "AUCUNE DONNEE", "AUCUNE DONNEE")
    definitions(106) = Array("DASHBOARD.SNAPSHOT.LABEL", "#{Id} - {Date}", "#{Id} - {Date}")
    definitions(107) = Array("DASHBOARD.LABEL.FORECAST_ISSUE", _
        "Forecast", _
        "Prévision")
    TextCatalogDashboard_Definitions = definitions
End Function
