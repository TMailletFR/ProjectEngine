Attribute VB_Name = "mod_TextCatalogMigration"
Option Explicit

Public Function TextCatalogMigration_Definitions() As Variant
    TextCatalogMigration_Definitions = Array( _
        Array("MIGRATION.TITLE", "Import previous ProjectEngine", "Importer un ancien ProjectEngine"), _
        Array("MIGRATION.WARNING_TITLE", "Warning - Import an existing project", "Attention - Import d'un projet existant"), _
        Array("MIGRATION.DESTRUCTIVE", "The data in this workbook will be replaced by the imported project.\nUnsaved changes may also be lost.\n\nTo keep the current state, make your own copy before continuing.\nThere is no automatic backup or rollback.\n\nOK: continue import. Cancel: keep this workbook unchanged.", _
            "Les données de ce fichier seront remplacées par le projet importé.\nLes modifications non enregistrées pourront aussi être perdues.\n\nPour conserver l'état actuel, créez une copie avant de continuer.\nAucune sauvegarde automatique ni retour en arrière.\n\nOK : continuer l'import. Annuler : conserver ce fichier intact."), _
        Array("MIGRATION.TARGET", "The destination must be saved, writable and structurally unprotected.", "La destination doit être enregistrée, modifiable et sans protection de structure."), _
        Array("MIGRATION.POST_SAVE_FAILED", "Import verification failed after saving. The source is unchanged. Start from a clean ProjectEngine copy; no automatic rollback is available.", "La vérification de l'import a échoué après enregistrement. La source est intacte. Repartez d'un ProjectEngine propre ; aucun retour automatique n'est disponible."), _
        Array("MIGRATION.SELECT", "Select your previous ProjectEngine workbook", "Choisir votre ancien classeur ProjectEngine"), _
        Array("MIGRATION.BLOCKED", "This workbook is incompatible or an import failed. Start from a clean compatible ProjectEngine workbook.", "Ce classeur est incompatible ou un import a échoué. Repartez d'un ProjectEngine propre et compatible."), _
        Array("MIGRATION.PROFILE", "The selected workbook is not a supported ProjectEngine profile. Use manual migration; nothing has been replaced.", "Le classeur choisi ne correspond pas à un profil ProjectEngine pris en charge. Utilisez la migration manuelle ; aucune donnée n'a été remplacée."), _
        Array("MIGRATION.FUTURE", "This file requires a newer ProjectEngine version. Download a compatible version before importing.", "Ce fichier nécessite une version plus récente de ProjectEngine. Téléchargez une version compatible avant l'import."), _
        Array("MIGRATION.OPEN_SOURCE", "Close the source workbook yourself after saving any edits, then select it again.", "Fermez vous-même le classeur source après avoir enregistré vos modifications, puis sélectionnez-le à nouveau."), _
        Array("MIGRATION.UNSAFE", "This source has external connections, links, macro sheets or a different date system. Automatic import is not supported; use manual migration.", "Cette source contient des connexions, des liens externes, des feuilles macros ou un système de dates différent. L'import automatique n'est pas pris en charge ; utilisez la migration manuelle."), _
        Array("MIGRATION.FORMULA", "A custom formula or override cannot be preserved automatically: {Field}. Use manual migration.", "Une formule personnalisée ou une valeur remplacée ne peut pas être conservée automatiquement : {Field}. Utilisez la migration manuelle."), _
        Array("MIGRATION.INVALID", "Invalid source data: {Field}. Correct it in the source file, save and retry. Nothing has been replaced.", "Donnée source invalide : {Field}. Corrigez-la dans le fichier source, enregistrez et recommencez. Aucune donnée n'a été remplacée."), _
        Array("MIGRATION.ORPHAN", "A constraint refers to a missing task: ID {Field}. Correct the source before importing.", "Une contrainte fait référence à une tâche introuvable : ID {Field}. Corrigez la source avant l'import."), _
        Array("MIGRATION.FAILED", "Import failed. The source is unchanged. No recovery or rollback is provided; start from a clean ProjectEngine copy before retrying. No final save was intentionally performed.", "L'import a échoué. La source est intacte. Aucun retour automatique ni sauvegarde de récupération : repartez d'un ProjectEngine propre. Aucun enregistrement final volontaire n'a été effectué."), _
        Array("MIGRATION.SUCCESS", "Import completed and saved. Check WBS and Constraints, then run Update Planning when ready.\nSource unchanged.", "Import terminé et enregistré. Vérifiez WBS et Constraints, puis lancez Update Planning quand vous le souhaitez.\nSource inchangée."), _
        Array("MIGRATION.STARTED", "ProjectEngine import started.", "Import ProjectEngine commencé."), _
        Array("MIGRATION.HISTORY_SKIPPED", "Old history unavailable; it will not be preserved.", "Ancien historique indisponible ; il ne sera pas conservé."), _
        Array("MIGRATION.DEFAULTS", "Missing task types become ordinary tasks. Missing Summary visibility is derived from the task hierarchy and milestones.", "Un type de tâche absent devient une tâche ordinaire. La visibilité Summary manquante est déduite de la hiérarchie et des jalons."), _
        Array("MIGRATION.OPTIONS", "Display language: {Language}; dates: {Date}. Review preferences in Settings after import.", "Langue d'affichage : {Language} ; dates : {Date}. Vérifiez vos préférences dans Settings après l'import."))
End Function


