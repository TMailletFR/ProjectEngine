Attribute VB_Name = "mod_TextCatalogWelcome"
Option Explicit

Public Function TextCatalogWelcome_Definitions() As Variant
    TextCatalogWelcome_Definitions = Array( _
        Array("WELCOME.TITLE", "Welcome to ProjectEngine", "Bienvenue dans ProjectEngine"), _
        Array("WELCOME.MESSAGE", "You can import data from an existing ProjectEngine workbook or start a new project.", "Vous pouvez importer les données d'un ProjectEngine existant ou commencer un nouveau projet."), _
        Array("WELCOME.IMPORT", "Import existing project", "Importer un projet existant"), _
        Array("WELCOME.NEW", "Start a new project", "Commencer un nouveau projet"), _
        Array("WELCOME.TUTORIAL", "Discover ProjectEngine", "Découvrir ProjectEngine"), _
        Array("WELCOME.CONTINUE", "Continue the project", "Continuer le projet"), _
        Array("WELCOME.MESSAGE.DRAFT", "You have started entering your project. Continue keeps those inputs. Starting a new project clears them only after confirmation.", "Vous avez commencé à renseigner votre projet. Continuer conserve ces saisies. Commencer un nouveau projet les efface uniquement après confirmation."), _
        Array("WELCOME.NEW.DISCARD.CONFIRM", "Discard the data entered during the tutorial and start a new empty project?", "Effacer les données saisies pendant le tutoriel et commencer un nouveau projet vide ?"))
End Function
