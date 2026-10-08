Attribute VB_Name = "mod_TextCatalogWelcome"
Option Explicit

Public Function TextCatalogWelcome_Definitions() As Variant
    TextCatalogWelcome_Definitions = Array( _
        Array("WELCOME.TITLE", "Welcome to ProjectEngine", "Bienvenue dans ProjectEngine"), _
        Array("WELCOME.MESSAGE", "You can import data from an existing ProjectEngine workbook or start a new project.", "Vous pouvez importer les données d'un ProjectEngine existant ou commencer un nouveau projet."), _
        Array("WELCOME.IMPORT", "Import existing project", "Importer un projet existant"), _
        Array("WELCOME.NEW", "Start a new project", "Commencer un nouveau projet"))
End Function
