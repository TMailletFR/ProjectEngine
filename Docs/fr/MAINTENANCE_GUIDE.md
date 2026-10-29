# Guide de maintenance

## Règle générale

Identifier d'abord le propriétaire de la donnée ou de la politique. Modifier ensuite la plus petite frontière correcte, puis choisir une validation proportionnée aux consommateurs réels.

## Ajouter une colonne WBS ou CALC

1. Déterminer si la colonne est une entrée WBS, une donnée de travail CALC ou une sortie calculée.
2. Mettre à jour le propriétaire du schéma dans DataSync/Infrastructure, pas chaque consommateur séparément.
3. Ajouter une validation Pre-Core si une absence ou une valeur invalide doit bloquer le calcul.
4. Ajouter le champ au dataset Core uniquement si le moteur le consomme réellement.
5. Mettre à jour les Full et Partial Output Writers si la colonne est écrite en sortie.
6. Ajouter la colonne aux scopes WBS Write Guard si le moteur l'écrit dans WBS.
7. Modifier `mod_IncrementalSignature` uniquement si la variation du champ doit invalider le calcul. Préserver l'ordre et versionner le contrat.
8. Rechercher tous les accès directs au nom de colonne et vérifier qu'ils appartiennent au bon propriétaire.

L'appartenance d'une colonne à la signature incrémentale est une décision métier, pas une conséquence automatique de son ajout.

## Ajouter un diagnostic

1. Produire le diagnostic dans le domaine qui détecte la situation.
2. Réutiliser les services de formatting existants pour le texte, sans recalcul métier.
3. Choisir explicitement `STOP`, `WARNING` ou `INFO`.
4. Définir un code et un hash stables si EventHistory ou ACK doivent reconnaître l'événement.
5. Router le message par MessageEngine ; ne jamais appeler directement la UserForm.
6. Journaliser par le contrat propriétaire d'EventHistory.
7. Ajouter un cas au harnais du producteur et vérifier les projections interactive et non interactive.

Le producteur possède le sens et la sévérité. MessageEngine possède la préparation. EventHistory possède la persistance et les ACK. La UserForm possède uniquement l'affichage.

## Ajouter un callback ou un bouton

1. Placer le callback dans un module standard ou un objet Excel compatible avec son mécanisme.
2. Utiliser un wrapper Public stable pour `OnAction`, `Application.Run`, `AddressOf`, timer ou event Excel.
3. Déléguer immédiatement au service propriétaire ; ne pas mettre de logique métier dans le callback.
4. Ajouter le nom dans l'en-tête `CALLBACKS EXTERNES / EXTERNAL CALLBACKS`.
5. Scanner `OnAction`, `Application.Run`, `SetTimer` et `AddressOf` avant tout renommage.
6. Tester le callback sur une copie temporaire.

Ne jamais renommer un callback externe uniquement pour des raisons de style. Conserver un wrapper lorsque la compatibilité l'exige.

## Maintenance du Ribbon, de l'accueil et de l'import (v1.5.0)

**Ribbon :** modifier `customUI.xml` et ses icônes embarquées, dont celles de Finance par le packaging existant. `mod_RibbonCallbacks` reste un adaptateur qui valide l'ownership workbook/fenêtre avant délégation métier. Conserver la visibilité contextuelle et l'invalidation des contrôles enfants ; **aucune invalidation Ribbon dans `Workbook_Deactivate`**. Compiler/sauvegarder/rouvrir, vérifier XML, relations, nombre d'images, labels et callbacks, puis tester dans une vraie fenêtre Excel possédée.

**Accueil :** `mod_ProjectWelcome` orchestre `frmProjectWelcome` et le propriétaire `EVENT_ACK`. Pas de store parallèle `WelcomeShown`. Start New ou un import réellement réussi peuvent acquitter ; Cancel/FAIL/X ne le doivent pas. Full Reset rend l'accueil éligible par le nettoyage existant. Vérifier qu'un projet actif ne déclenche pas l'accueil.

**Migration :** suivre `Migration_ImportFile` -> `MigrationData_Read` -> `Migration_ApplyInputs` -> propriétaires Settings/WBS/Constraints/Dashboard/EventHistory. Source en lecture seule, macros désactivées, inchangée. Import **data-only** : ni Core ni graphique automatique. En v1.5.0, l'import direct **ne crée pas** de backup automatique de la destination ni de classeur de staging : avertissement destructif et copie préalable par l'utilisateur indispensables. Vérifier les données, les distinctions formule/texte/valeur et l'ACK de succès, limiter à un Save final, propager les erreurs sans faux SUCCESS. Refuser explicitement les schémas/formules non supportés. Ne pas garantir une restauration transactionnelle ou une protection contre AutoSave.

## Maintenance du tutoriel, des consoles et des métadonnées (v1.5.0)

**Tutoriel :** `mod_ProjectTutorial` possède une session non modale, 29 étapes sémantiques et une seule Shape de surlignage ; `mod_Settings` possède deux bookmarks Découverte/Pratique. Une page passive n'appelle jamais Core, ne reconstruit pas Gantt et n’injecte pas d’exemple. Une page active propose une vraie commande avec consentement ; un retour normal de `Sub` ne prouve pas la réussite métier. Exiger un reçu terminal frais des owners. Préserver la saisie WBS, la validation classeur/fenêtre, les contextes Ribbon et le nettoyage au Save/SaveAs/Reset/Close/erreur. Valider STOP/correction, warning, planning vide, annulation d’import et coexistence avec la vraie console modale.

**Présentation :** conserver les trois UserForms et l’unique infrastructure PEButton/présentation, avec fallback CommandButton MSForms natif. Pas de fenêtre hôte parallèle, ni seconde console métier. Vérifier libellés, clavier, Tab/Maj+Tab, Entrée/Échap, édition Excel derrière le tutoriel, tailles/scroll et ACK. Les glyphes Unicode non rendables fidèlement en WMF passent par le contrôle classique BSTR, jamais par une translittération silencieuse. Excel 32 bits, LTSC et multi-DPI restent non testés nativement : fallback prudent.

**Confidentialité :** ne pas persister `source.FullName`, chemins absolus ou exceptions brutes dans les reçus d’import, l’historique ou les propriétés. Préserver `ImportState`, `SchemaVersion`, `ReleaseId`, digest et ACK. Garder les alertes Excel actives. Après la dernière sauvegarde du XLSM public, assainir et relancer le gate de confidentialité : Excel peut recréer ses propres métadonnées de localisation. Examiner séparément les attributions personnelles EVENT_ACK et les contenus métier. Le nettoyage des releases historiques ne modifie pas leur ancien code VBA d’import.

**Limite connue :** le rejet `MIGRATION.PROFILE` rapporté puis disparu après Full Update n’a pas de cause démontrée, faute de fichier BEFORE intact. Ne pas annoncer la correction ni élargir la reconnaissance sans preuve ; préserver le refus sans mutation de la cible.

## Maintenir Finance et le multiprojet (v1.5.0)

- **Autorité unique :** `mod_Finance` possède validation, contrats financiers, commit et calcul ; les cinq tables de **FINANCE_DATA** (VeryHidden) constituent le stockage central via l'adaptateur FinanceStore. `FINANCE`, `FINANCE SUMMARY` et `FINANCE INPUT` sont respectivement projections et brouillon temporaire : ne pas y créer une seconde source de vérité.
- **Contexte :** une activité liée à plusieurs projets ne se reconnaît pas par `STATE|WBS` seul. Préserver les IDs et associations persistants, et distinguer relation partagée (dédupliquée dans Global), copie indépendante et fiche indépendante (additionnées).
- **Transactions :** tableau natif filtrable/triable, modifications en brouillon, confirmation/annulation. Vérifier gardes de modification directe, collages mixtes, sauvegarde/fermeture et absence de mutation sur Cancel. Deux dépenses identiques peuvent être deux écritures distinctes ; ne pas les fusionner arbitrairement.
- **Montants et temps :** valider signes, centimes, crédits, ETC/EAC et règle de distribution START/END/LINEAR. Le journal conserve les opérations futures ; **le graphique du réalisé** s'arrête au jour de référence du dernier `Finance_Refresh`. Le changement de langue et l'ouverture ne déclenchent pas Finance_Refresh. Une devise affichée ne réalise pas une conversion de change.
- **Rendu :** conserver le ChartObject Finance et le helper physique commun du repère Aujourd'hui ; ne pas dupliquer le renderer S-Curve. Les tables mensuelles/totaux peuvent conserver les opérations futures, même lorsque le graphique ne les montre pas encore.
- **Import v3 :** `mod_Migration` / `mod_MigrationData` lisent et écrivent le **domaine entier**, même si seule la vue d'un projet est affichée. Ne pas ressusciter l'ancien cycle Finance_Reset complet avant écriture ; vider seulement les sorties dérivées avec `Finance_ClearOutputs` puis remplacer l'autorité une fois. Vérifier Save final réel, reouverture, version future refusée, ancien classeur sans Finance, source inchangée et absence de refresh implicite.
- **Full Reset :** nettoyer via les owners, y compris FINANCE_DATA, journal, FINANCE SUMMARY, FINANCE_TODAY_LINE, extractions, liste de projets en AP, cache SVG et métadonnées d'ancien import. Conserver structures et préférences prévues ; vérifier un nouveau projet après Reset. Un Full Reset n'est pas une transaction disque atomique.
- **Mise en forme :** préserver le format Texte de WBS/Predecessors WBS, les formats de saisie WBS transportables et les largeurs des colonnes WBS/Constraints/Finance/Finance Summary par clés du schéma. Un collage Excel complet peut remplacer le format Texte par celui de la source ; ne pas promettre l'immunité aux collages.

**Gates proportionnés :** compilation, véritables transactions Finance, Global et partage/déduplication, débits/crédits/futur, import ancien/v3, Save/reopen, données des trois maîtres conservées, Full Reset + nouveau projet, scan confidentialité du candidat public. Les tests 32 bits, LTSC, autres DPI et reprise après crash ne sont pas certifiés par les seules preuves v1.5. Le dernier gate d'audit du maitre peut légitimement signaler des ACK personnels : un **artefact de release** doit être vérifié après Full Reset et dernière sauvegarde, sans exception de confidentialité destinée au développement.


## Ajouter une règle Task Type

1. Ajouter la normalisation ou la classification dans `mod_TaskTypeRules`.
2. Ne pas recopier la règle dans Core, Constraints, S-Curve ou Gantt.
3. Laisser chaque consommateur appliquer sa propre politique après classification : exclusion du rendu, validation et calcul sont des décisions différentes.
4. Ajouter des cas LOE, Milestone et valeur inconnue aux harnais directement consommateurs.

## Modifier la signature incrémentale

1. Modifier uniquement `mod_IncrementalSignature` pour les champs, l'ordre et la sérialisation.
2. Considérer ce format comme un contrat persistant.
3. Capturer les signatures avant modification avec `mod_IncrementalSignatureHarness`.
4. Toute modification de champ exige une stratégie explicite de version ou d'invalidation de CALC_STATE.
5. Vérifier `mod_CalcState` et `mod_CalcIncremental` sans y redéfinir le contrat.
6. Exiger une golden capture. Toute différence bit à bit doit être voulue et expliquée.

## Intervenir sur le Gantt

| Changement | Propriétaire | Validation minimale |
|---|---|---|
| date -> position ou géométrie | Geometry / TimelineGeometry | compilation et harnais visuel ciblé |
| barre, milestone ou summary | GanttRenderer | Visual Regression et TEST/fallback selon les consommateurs |
| dépendances | DependencyRenderer | Registry et Visual Regression |
| contraintes ou deadlines | ConstraintRenderer | Visual Regression et scénario concerné |
| registry, diff ou Lazy Repair | ShapeRegistry | Predictive Registry et Visual Regression |
| boutons, toggles ou langue | UiControls / ViewState / Language | signature UI ciblée |
| drag ou timer | GanttDragWatch | smoke Drag/TEST et cycle de vie timer |
| TEST, SCENARIO ou LOCK | service correspondant | smoke transactionnel sur copie |

Ne jamais modifier les noms de Shapes, `OnAction`, z-order, tolérances ou fallback au cours d'un simple nettoyage. Ne jamais créer un second renderer ou un second moteur de simulation.

## Contrats performance Full Update, Gantt et EventHistory

- **Full Update :** conserver l'exécution complète Core/Analytics. Dans `mod_CoreBridgeOutputWriter`, sauter une *colonne* WBS uniquement après comparaison intégrale et en absence de **toute formule** dans cette colonne. Un résultat modifié ou une formule parasite doivent être corrigés par Full Update. Ne pas présenter cela comme un gain de vitesse du moteur de calcul.
- **Réutilisation Gantt :** les caches conservés par DEFER restent des candidats non validés. Vérifier le contexte visuel complet et l'état physique, et toutes les routes incidentes aux Shapes localement modifiées. Garder le fallback FULL canonique pour les changements de structure, contraintes, layout ou géométrie inconnue. Comparer formes physiques et routes à un rendu FULL canonique ; mesurer séparément unchanged, small, large et contre-exemples.
- **EventHistory :** conserver l'unique store et owner ACK, les reçus d'écriture et la propagation d'erreur. Le bulk ou le regroupement ne doivent ni perdre d'événement, ni modifier une identité, ni dupliquer un ACK, ni masquer une panne de stockage. Vérifier store **et** projection visible.
- **Performance :** mesurer le travail réellement évité, temps muraux, état de cache et variance Excel. Un chrono warm isolé ne prouve pas une accélération universelle.

## Choisir le niveau de validation

### Niveau S : structure

Renommage local, visibilité, suppression de code mort ou commentaires : audit statique, import ciblé, compilation et harnais du consommateur uniquement si nécessaire.

### Niveau M : module

Extraction locale, nouveau contrat interne ou changement de projection : compilation, golden ciblée et workflows consommateurs directs.

### Niveau A : architecture ou comportement

Nouvelle frontière, moteur, writer, callback externe ou workflow principal : validation élargie incluant les guards et transactions affectés.

Ne pas lancer une batterie globale par habitude. Ne pas omettre un harnais réellement consommateur pour gagner du temps.

## Encodages CP1252 et UTF-8

- Lire les octets et détecter UTF-8 BOM, puis UTF-8 valide, puis CP1252 en fallback.
- Réécrire avec le même encodage et les mêmes fins de ligne.
- Ne jamais convertir implicitement un fichier VBA historique CP1252 en UTF-8.
- Préférer l'ASCII dans le code et les commentaires VBA lorsque la source l'utilise.
- Conserver les documents Markdown et TSV en UTF-8.
- Après une transformation documentaire, comparer un hash du code hors commentaires.

## Importer et compiler le VBA

Outils de référence :

```powershell
powershell -ExecutionPolicy Bypass -File codex_tools\Import_VBA_To_Workbook.ps1
powershell -ExecutionPolicy Bypass -File codex_tools\tmp_compile_vba_project.ps1
```

Lire les paramètres du script et vérifier le classeur cible avant exécution. Les smokes doivent créer leur propre copie sous `%TEMP%`, conserver le hash du classeur source et tracer le processus Excel qu'ils créent.

Après import, compiler le VBAProject complet. Un scan statique ne remplace pas la compilation VBA : visibilité, arguments optionnels et types d'arrays peuvent échouer uniquement dans le VBE.

## Ne jamais fermer l'Excel utilisateur

1. Capturer les processus Excel existants avant le worker.
2. Créer une instance COM dédiée et mémoriser son PID ou handle.
3. Ouvrir uniquement une copie temporaire identifiée.
4. Fermer **uniquement le classeur de test possédé**, après vérification exacte de `Name` et `FullName` ; libérer les références COM. Appeler `Application.Quit` seulement si la possession exclusive de l’instance est certaine et qu’aucun classeur utilisateur n’y est attaché.
5. En cas de timeout ou perte RPC, documenter et distinguer produit, harnais, fixture et session COM. Ne pas tuer Excel sur la seule base d’un PID créé ou associé au test ; ne jamais forcer la fermeture d’une instance d’ownership incertain.
6. Interdire `taskkill`, `Stop-Process EXCEL` génériques et fermetures par nom partiel.
7. Tout Excel utilisateur reste hors périmètre, même s'il ouvre un autre classeur.

## Choisir entre Public, Friend et Private

| Visibilité | Utilisation |
|---|---|
| `Private` | helper consommé dans un seul module ou état interne d'une classe |
| `Friend` | contrat nécessaire dans le même VBAProject mais non destiné à Excel ou à une macro externe |
| `Public` | macro utilisateur, callback externe, harnais appelé par worker ou contrat inter-modules démontré |

En cas de doute sur un appel dynamique, conserver Public et documenter le contrat. Ne pas rendre un helper Public uniquement parce qu'il a été déplacé.

## Quand créer un module

Créer un module lorsqu'une responsabilité complète possède un contrat, un invariant et plusieurs procédures cohérentes, ou lorsqu'il devient propriétaire unique d'un contrat canonique réellement dupliqué.

Ne pas créer de module pour :

- quelques lignes d'orchestration ;
- un helper isolé sans politique propre ;
- éviter de transmettre deux paramètres ;
- masquer une dépendance qui devrait rester visible ;
- regrouper des procédures parce que leurs noms se ressemblent.

## Quand ne pas créer Wrapper, Context ou Read Model

- Pas de wrapper pour un renommage local atomique sans contrat dynamique.
- Pas de Context si ses membres n'ont pas le même cycle de vie.
- Pas de DTO universel mélangeant WBS, CALC, UI, diagnostics et options.
- Pas de Read Model si les consommateurs partagent seulement une table source mais pas la même projection.
- Pas de booléens de politique pour forcer deux workflows différents dans une abstraction commune.

## Erreurs d'architecture à éviter

- façade passe-plat sans valeur de stabilité ;
- DTO universel ;
- politique métier cachée dans `Utils`, `Helper` ou `Manager` ;
- lecture cellule par cellule lorsqu'un array de table suffit ;
- duplication d'Identity Index, Parsed Network ou Incremental Signature ;
- accès direct au store d'un autre propriétaire ;
- reset global connaissant tous les schémas ;
- fermeture d'un scope WBS non possédé ;
- dépendance TEST -> SCENARIO ou LOCK -> TEST Service ;
- validation disproportionnée au risque ;
- commentaire qui répète seulement le nom de la procédure ;
- timeout transformé en PASS ou fallback masqué par un harnais.

## Checklist avant livraison

1. Propriétaire et frontière identifiés.
2. Contrats externes et callbacks scannés.
3. Corps métier inchangé ou changement explicitement autorisé.
4. Documentation du module et des procédures mise à jour.
5. Encodage préservé et hash hors commentaires contrôlé pour un changement documentaire.
6. Import et compilation adaptés au risque.
7. Harnais des consommateurs directs verts.
8. Classeur source protégé.
9. Aucune instance Excel utilisateur fermée.
10. Rapport proportionné et limites de couverture explicites.
