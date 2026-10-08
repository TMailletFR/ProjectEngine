Attribute VB_Name = "mod_TextCatalogSettings"
Option Explicit

Public Function TextCatalogSettings_Definitions() As Variant
    TextCatalogSettings_Definitions = Array( _
        Array("SETTINGS.LANGUAGE.TITLE", "Language", "Langue"), _
        Array("SETTINGS.CONSTRAINTS.LABEL", "Constraints", "Contraintes"), _
        Array("SETTINGS.DATE_FORMAT.TITLE", "Date format", "Format des dates"), _
        Array("SETTINGS.RESET.TITLE", "Reset", "Réinitialisation"), _
        Array("SETTINGS.CLEAR_HISTORY.LABEL", "Clear History", "Nettoyer historique"), _
        Array("SETTINGS.CLEAR_ACK.LABEL", "Clear Acknowledged", "Nettoyer les messages acquités"), _
        Array("SETTINGS.CLEAN_DASHBOARD.LABEL", "Clean Dashboard", "Nettoyer Dashboard"), _
        Array("SETTINGS.DANGER.TITLE", "Danger Zone", "Zone de danger"), _
        Array("SETTINGS.FULL_RESET.LABEL", "Full Reset", "Réinitialisation complète"), _
        Array("SETTINGS.PAGE.TITLE", "Settings", "Options"), _
        Array("SETTINGS.ACTIVATION.LABEL", "Activated", "Activé"), _
        Array("SETTINGS.OWNER.GLOBAL", "GLOBAL", "GLOBAL"), _
        Array("SETTINGS.OWNER.DASHBOARD", "Dashboard", "Dashboard"), _
        Array("SETTINGS.OWNER.GANTT", "Gantt", "Gantt"), _
        Array("SETTINGS.OWNER.SCURVE", "S-Curve", "S-Curve"), _
        Array("SETTINGS.OWNER.WBS", "WBS", "WBS"), _
        Array("SETTINGS.OWNER.EVENT", "Messages & Event History", "Messages et historique"), _
        Array("SETTINGS.DATE_FORMAT.OPTIONS", "DMY / MDY / ISO", "DMY / MDY / ISO"), _
        Array("SETTINGS.INFO.LABEL", "Info", "Info"), _
        Array("SETTINGS.TOGGLE.OFF", "OFF", "OFF"), _
        Array("SETTINGS.TOGGLE.ON", "ON", "ON"), _
        Array("SETTINGS.LANGUAGE.FR", "FR", "FR"), _
        Array("SETTINGS.LANGUAGE.EN", "EN", "EN"))
End Function
