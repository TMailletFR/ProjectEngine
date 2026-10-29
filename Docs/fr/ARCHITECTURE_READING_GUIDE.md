# Carte de lecture de l'architecture

## Découverte en 15 minutes

Lire les éléments suivants dans cet ordre pour comprendre les frontières du projet sans entrer immédiatement dans les algorithmes :

1. `mod_RunButtons` : macros utilisateur stables `Run_*`.
2. `mod_RuntimeWorkflow`, `mod_MacroGuard`, `mod_PlanningConsolePolicy` : cycle de vie d'une commande.
3. `mod_CalcEngineCoreBridge` : orchestration générale du planning.
4. `mod_DataSync`, puis `mod_CalcCoreProdWrapper` : passage de WBS/CALC vers le Core.
5. `mod_CalcCoreEngine`, `mod_CalcCoreNetwork`, `mod_CalendarEngine` : moteur de calcul unique.
6. `mod_CoreBridgeAnalytics` et `mod_CoreBridgeOutputWriter` : analytics et sorties.
7. `mod_GanttRefreshPipeline`, `mod_GanttRenderer`, `mod_GanttShapeRegistry` : rendu Gantt.
8. `mod_GanttLive` et les services TEST/SCENARIO/LOCK : simulation.
9. `mod_MessageEngine`, `mod_EventHistory`, `frmPlanningMessages` : chemin des diagnostics jusqu'à l'utilisateur.
10. `PROJECT_GLOSSARY.md` et `MAINTENANCE_GUIDE.md` : vocabulaire et règles de modification.

## Entrées Ribbon et accueil (v1.5.0)

`customUI.xml` est embarqué dans chaque `.xlsm` avec des commandes illustrées, dont quatre actions Finance. `mod_RibbonCallbacks` est un **adaptateur UI mince**, pas un propriétaire de calcul. Il valide le classeur d'origine et la fenêtre active avant d'appeler les entrées publiques existantes. Les Updates sont contextuelles sur WBS ; Test/Scenario/Lock/Reset sur GANTT ; Settings, Navigation, Reset et Import suivent leur périmètre global. `ThisWorkbook` orchestre les événements du classeur et l'activation rafraîchit le Ribbon. **Ne pas invalider les contrôles Ribbon dans `Workbook_Deactivate`** (coexistence de plusieurs fichiers ProjectEngine).

Sur un classeur vide et modifiable, `mod_ProjectWelcome` affiche `frmProjectWelcome` via `ProjectWelcome_ShowIfNeeded`. Start New acquitte l'accueil via le propriétaire EventHistory/ACK existant ; Import réutilise la même entrée que le Ribbon. Fermeture, annulation ou échec n'enregistrent pas un accueil réussi. Full Reset rend l'accueil à nouveau éligible. Le formulaire est une vue dédiée, sans nouveau store de messages, moteur de planning ou catalogue de traduction.

## Tutoriel interactif et reçus de commandes (v1.5.0)

La commande Ribbon de tutoriel (`peTutorial`) et **Discover ProjectEngine** depuis l’accueil délèguent à `mod_ProjectTutorial` ; `frmProjectTutorial` est une **vue non modale**, pas un propriétaire du planning. Le catalogue TextCatalog commun fournit le FR/EN et Settings possède la progression persistante. Découverte et Pratique partagent **29 StepKeys sémantiques**, avec des repères de reprise indépendants. La navigation passive résout et surligne les cibles réelles sans calcul, réparation Gantt, injection de données ni watcher permanent.

La pratique conserve la vraie saisie Excel et propose explicitement les commandes normales par le **dispatcher Ribbon et les owners métier existants**. `mod_RuntimeWorkflow` fournit l’identité d'exécution et le résultat terminal ; le tutoriel exige un reçu frais et des entrées toujours valides. STOP, échec, annulation et refus de guard n’autorisent pas une progression comme succès. L’état vide n'est une réussite que si l’owner confirme une opération safe-empty accomplie : pas de second moteur ni validateur pédagogique.

Le surlignage est une unique Shape temporaire identifiée (`PE_TUTORIAL_HIGHLIGHT`), retirée pour Save/SaveAs/Close et sans mutation de valeurs, formats ou Shapes métier. Le formulaire non modal vérifie son classeur et sa fenêtre, puis ferme proprement à la perte d’ownership. Les bookmarks résident dans Settings, pas dans ACK ; Full Reset remet à zéro les deux modes ; un import conserve la progression de destination.

## Fenêtres modernisées et frontière de présentation (v1.5.0)

`frmPlanningMessages` est **modale**, `frmProjectWelcome` est **modale**, `frmProjectTutorial` est **non modale**. Elles utilisent l’infrastructure visuelle commune (`PEButton`, `PEConsoleButton`, `PEConsolePresentation`, `mod_PEPaint`, `mod_PEPresentation`, `mod_PEConsoleStyle`), sans nouveaux owners métier ni installation externe. MessageEngine/EventHistory/ACK et les handlers Welcome/Tutorial font autorité. Les vrais CommandButtons MSForms restent le fallback, notamment pour les légendes incompatibles avec le renderer vectoriel. La console mesure son texte natif avant Show et garde une géométrie stable entre pages. Préserver clavier, focus, Échap, Default/Cancel et édition Excel derrière le tutoriel.

## Confidentialité de l'import et des artefacts publiés (v1.5.0)

`mod_MigrationData` et `mod_Migration` ne persistent plus le chemin absolu de la source dans le snapshot, le journal ou les identités d’événements. Les détails complets d’exception restent de session ; le diagnostic persistant est borné à la phase/code d’erreur. `mod_WorkbookSchema` conserve l’autorité sur les propriétés. La source est toujours en lecture seule, macros désactivées. Excel peut recréer des métadonnées de localisation après sauvegarde : **assainir l’artefact final après sa dernière sauvegarde et relancer le gate de confidentialité** ; nettoyer séparément les attributions nominatives `EVENT_ACK` destinées à la publication. Ne pas activer RemovePersonalInformation/filterPrivacy pour forcer le gate au prix d’un avertissement de sauvegarde.

## Domaine Finance v1.5.0 : autorité, vues et calcul explicite

La **Finance** possède un owner métier unique (`mod_Finance`) et un stockage persistant central dans la feuille technique **FINANCE_DATA** (VeryHidden), avec cinq tableaux de domaine : projets, rattachements, fiches financières, journal et métadonnées. Les feuilles **FINANCE** et **FINANCE SUMMARY** sont des *projections* et ne sont plus les magasins canoniques de toutes les données financières.

```text
Ribbon (vérification du workbook) / édition native Excel
  -> mod_Finance : contrôles / transaction / commit explicite
  -> données financières canoniques dans FINANCE_DATA (FinanceStore = I/O)
  -> Finance_Refresh (WBS/CALC lus, jamais de second scheduler)
  -> FINANCE + FINANCE SUMMARY : projections / graphique existant
```

Les projets financiers peuvent référencer des fiches indépendantes, copiées ou partagées. La vue **Global** déduplique les contributions réellement partagées et additionne les fiches indépendantes. Les identités financières ne sont pas les seuls IDs WBS : un même WBS peut exister dans plusieurs projets financiers. Les deux vues visibles restent un seul contexte sélectionné ; une vue non affichée ne doit pas perdre son stockage.

**FINANCE INPUT** est une table temporaire d'édition Excel, non une seconde autorité. Les formulaires et le Ribbon restent des adaptateurs ; seul l'owner Finance valide et confirme. Les dépenses sont datées, signées, et historisées ; le réalisé du graphique est limité au jour du dernier Refresh Finance alors que le journal peut contenir des dépenses futures. START/END/LINEAR sont des profils de budget, distincts de la S-Curve physique. Finance ne s'exécute pas automatiquement au Planning Update, à l'ouverture ou à l'import.

`SchemaVersion=3` décrit le nouveau domaine persistant. `mod_Migration`/`mod_MigrationData` transportent le domaine Finance entier, pas les seules tables visibles. Le Full Reset appelle le reset de l'owner Finance et vide les caches/projections appartenant aux autres owners. L'état final ne doit conserver aucune ancienne fiche, dépense, extraction ou ancienne identité d'ACK ; préférences autorisées et structures persistent.


## Flux principal

```text
Callback Excel / macro Run_* / OnAction
    -> RuntimeWorkflow + MacroGuard
    -> DataSync (WBS -> CALC + LOGIC_LINKS)
    -> validations Pre-Core
    -> CalcCoreProdWrapper
    -> CalcCoreEngine + CalcCoreNetwork + CalendarEngine
    -> CoreBridgeOutputWriter (CALC puis WBS, sous WBS Write Guard)
    -> CoreBridgeAnalytics / Variances
    -> rafraîchissements Gantt / S-Curve / Dashboard
    -> MessageEngine
    -> PlanningConsolePolicy
    -> EventHistory / ACK + frmPlanningMessages
```

Les flèches indiquent la direction d'orchestration. Un domaine accède à un autre domaine par ses contrats publics, jamais par son état privé. Les diagnostics se déplacent vers MessageEngine ; ils ne remontent pas dans le calcul.

## Import depuis un ancien classeur ProjectEngine

```text
Ribbon Import / accueil Import
    -> ProjectWelcome_ImportPrevious
    -> Migration_ImportPrevious / Migration_ImportFile
    -> vérification source/destination/schéma et confirmation destructive
    -> source ouverte en lecture seule, macros désactivées, liens non mis à jour
    -> MigrationData_Read (entrées et stores pris en charge)
    -> fermeture de la source sans sauvegarde
    -> Migration_ApplyInputs (remplacement des données cibles prises en charge)
    -> propriétaires Settings / formules WBS / Constraints / Dashboard / EventHistory / Finance
    -> vérification des données et des sorties dérivées vides
    -> reçu SUCCESS + acquittement accueil, puis une sauvegarde de la cible
    -> vérification après sauvegarde
```

L'import est **data-only** : aucun appel automatique au Core, aucun Gantt/S-Curve reconstruit, aucune macro du fichier source exécutée. L'utilisateur vérifie WBS et Constraints puis lance explicitement Planning Update ou Full Update. Les données prises en charge incluent les inputs WBS (formules utilisateur compatibles), Constraints, Settings, EVENT_ACK et snapshots Dashboard ; un historique compatible peut être repris, sinon le bonus incompatible est écarté explicitement. La cible conserve son propre VBA, Ribbon et ses contrats de sorties.

**Frontière transactionnelle v1.4.0 :** import direct et destructif, sans création automatique d'une copie de récupération de la destination ni workbook de staging. La source reste intacte. Demander une **sauvegarde préalable de la destination**. Une erreur avant le Save final ne doit pas produire SUCCESS/ACK de succès ; ne pas présenter cette procédure comme un rollback atomique garanti face à Excel, AutoSave ou une interruption après Save.

`ProjectEngine.SchemaVersion` versionne le schéma persistant ; `ProjectEngine.ReleaseId` identifie séparément la release et n'est pas inventé par l'import. Un legacy sans marqueur est reconnu par son profil structurel validé. Schémas futurs/incompatibles et entrées dangereuses sont refusés ; le VBA source n'est jamais exécuté pour la migration.

## Workflow Update Planning

1. `Run_Planning_Update` ouvre MacroGuard et un workflow runtime.
2. `Run_Calc_Engine_CoreBridge` traite le Safe Empty State, prépare l'infrastructure et synchronise les tables.
3. Les validations Pre-Core produisent des STOP sans créer un second moteur.
4. `Run_Calc_Core_PROD_Pilot` prépare le dataset de travail et le réseau parsé.
5. `Run_Calc_Core` calcule les tâches feuilles, propage les erreurs et applique le post-traitement LOE.
6. Les writers persistent CALC, puis WBS, dans l'ordre contractuel.
7. Les analytics calculent paths, floats, deadlines, variances et warnings.
8. MessageEngine prépare la console ; Runtime peut différer son affichage jusqu'au workflow racine.

## Chemin WBS -> CALC -> Core -> sorties

| Étape | Propriétaire | Données | Invariant |
|---|---|---|---|
| Entrées utilisateur | WBS et `mod_WBSEvents` | `tbl_WBS` | Les colonnes calculées restent protégées. |
| Synchronisation | `mod_DataSync` | `tbl_CALC`, `tbl_LOGIC_LINKS` | WBS fournit les entrées ; CALC est le dataset moteur. |
| Identité | `mod_CanonicalIdentityIndex` | maps ID/WBS/lignes | Les maps exposées sont en lecture seule. |
| Réseau | `mod_ParsedPlanningNetwork` | Succ/Pred/Type/Lag | Le parsing est commun ; les projections métier restent séparées. |
| Calcul | `mod_CalcCoreEngine` | array Core mutable | Il n'existe qu'un seul moteur planning. |
| Persistance | `mod_CoreBridgeOutputWriter` | CALC puis WBS | Full et Partial conservent les mêmes champs et le même ordre. |
| Protection WBS | `mod_WBSWriteGuard` | scopes tokenisés | Un appelant ne ferme que son propre token, en ordre LIFO. |

## Full Update et publication des sorties

`Run_Full_Update` impose toujours un calcul Core et Analytics complet. En v1.5.0, le writer **compare les colonnes de résultats déjà calculés** avec les valeurs WBS existantes et saute uniquement une colonne entièrement identique sans formule. Toute différence, ou toute formule dans une colonne de sortie calculée, déclenche l'écriture bulk normale. Ce n'est pas un calcul Core incrémental, et un premier rendu complet n'est pas promis plus rapide. L'owner des sorties reste unique.

## Domaine Gantt

`Refresh_Gantt` est le wrapper public stable. `mod_GanttRefreshPipeline` acquiert les données et choisit le chemin Full ou Display Only. Les renderers reçoivent des arrays et maps déjà préparés.

- `mod_GanttRenderer` dessine tâches, summaries, milestones et ligne du jour.
- `mod_GanttDependencyRenderer` route et dessine les dépendances.
- `mod_GanttConstraintRenderer` dessine contraintes et deadlines.
- `mod_GanttShapeRegistry` possède les records de Shapes, le cache et le diff prédictif.
- `mod_GanttGeometry` et `mod_GanttTimelineGeometry` fournissent des calculs purs.
- `mod_GanttUiControls`, `mod_GanttViewState` et `mod_GanttLanguage` possèdent l'UI, pas le calcul.

Zones sensibles :

- noms, `OnAction`, z-order et géométrie des Shapes ;
- fast path prédictif Day, fallback Week/Month et Lazy Repair ;
- cycle de vie du watcher Drag et de son timer ;
- cohérence entre registry attendu et état réel de la feuille.

## Navigation Gantt, READY et réutilisation sûre

`mod_GanttNavigation` résout les Task IDs et la visibilité ; `mod_GanttViewState` possède la projection Detail/Summary. La navigation peut demander `EnsureGanttForCurrentPlanning` si la vue doit être préparée, mais ne lance pas silencieusement un calcul de planning obsolète. Show Full Timeline cadre l'étendue rendue dans les limites physiques de zoom Excel et utilise un fallback explicite si le fit complet est impossible.

Le rendu conserve les owners `mod_GanttRefreshPipeline`, `mod_GanttRenderer`, `mod_GanttDependencyRenderer` et `mod_GanttDependencySvg`. DEFER peut conserver la géométrie et les routes comme **candidats à la réutilisation**, jamais comme preuve READY automatique. La réutilisation demande des vérifications d'inputs, contexte et état physique ; la géométrie locale modifiée doit couvrir toutes les routes affectées. Changement structurel, contraintes, layout de la frise ou mutation physique non suivie : fallback FULL canonique. Day rend les dépendances ; Week/Month conservent leur politique de visibilité agrégée. Aucun deuxième renderer, store de routes ou autorité READY.

## TEST, SCENARIO et LOCK

| Mode | Propriétaire | Entrée | Sortie | Interdiction principale |
|---|---|---|---|---|
| TEST | `mod_GanttTestService` | cellules TEST jaunes | `tbl_CALC_GANTT_TEST`, overlay prédictif | aucune écriture durable dans WBS |
| SCENARIO | `mod_GanttScenarioService` | planning ou copie scénario | dataset et rendu scénario | ne dépend pas de TEST comme moteur parent |
| LOCK | `mod_GanttLockService` | simulation validée | Forecast WBS durable | ne contourne jamais WBS Write Guard |

`mod_GanttLive` conserve les wrappers historiques et les transactions publiques. `mod_GanttSimulationState` possède le mode et la demande de rendu. `mod_GanttSimulationTableStore` possède le schéma et le reset de `tbl_CALC_GANTT_TEST`. Scenario Fork conserve son contrat `Application.Run`.

## S-Curve et Dashboard

`mod_SCurve` est l'unique moteur des séries temporelles et le propriétaire de leurs sorties. `SCurve_BuildDashboardProjection` expose une projection dédiée au Dashboard.

`mod_DashboardReadContext` acquiert une fois WBS, CALC et cette projection pour les trois modes Dashboard. `mod_Dashboard` conserve séparément les politiques Full Build, Content Only et Texts/Comparison.

## Correctifs Dashboard et S-Curve v1.5.0

Le Dashboard possède sa projection journalière relative des summaries, sans recalcul du planning ni pondération de poids S-Curve. Les instantanés incomplets sont rejetés avant engagement et les nouvelles écritures partielles font l’objet d’un rollback ciblé. La sélection des instantanés se valide contre l’historique métier, jamais contre une validation Excel transitoire. La S-Curve peut résoudre la baseline manquante via le Core existant : pas de baseline inventée depuis Forecast.


## Localisation commune

Le TextCatalog logique partagé résout **TextKey + LanguageKey fourni par le consumer**. Le catalogue ne possède pas la langue. Les six owners indépendants sont conservés ; GLOBAL coordonne les changements sans devenir l'autorité unique. Ribbon, accueil, import et navigation consomment cette infrastructure. Les fallbacks doivent être explicites et testables ; l'API logique ne doit pas être enfermée dans le stockage CP1252.

## Diagnostics, console, EventHistory et ACK

```text
Producteurs CoreBridge / Constraints / S-Curve
    -> collections de messages structurés
    -> filtrage et regroupement MessageEngine
    -> PlanningConsolePolicy
       -> mode interactif : frmPlanningMessages.Show vbModal
       -> mode harnais : capture sans affichage
    -> journalisation EventHistory
    -> ACK éventuel du warning, sans suppression de l'historique
```

Le producteur décide le sens et la sévérité du diagnostic. MessageEngine prépare et regroupe sans recalculer. EventHistory possède le stockage et les ACK. La UserForm affiche uniquement une projection déjà préparée.

## Stores, snapshots et contrats canoniques

| Composant | Possède | Ne possède pas |
|---|---|---|
| Canonical Identity Index | ID, WBS normalisé, index de lignes, Driving Logic | hiérarchie métier et calcul |
| Parsed Planning Network | parsing immuable des liens | tri topologique, validation et routing Gantt |
| Incremental Signature | 17 champs, ordre, normalisation, sérialisation | CALC_STATE et décision de recalcul |
| CalcState | persistance du snapshot incrémental | définition de la signature |
| Dashboard Read Context | acquisition commune d'un refresh | rendu des trois modes |
| Finance / FinanceStore | FINANCE_DATA canonique, validation et projections financières | dates Core ni logique métier Ribbon autonome |
| Simulation Table Store | `tbl_CALC_GANTT_TEST` et son reset | politique TEST/SCENARIO/LOCK |

## Guards et invariants de sécurité

- `MacroGuard` empêche les exécutions concurrentes et transporte les demandes d'abandon.
- `RuntimeWorkflow` maintient profondeur, workflow racine et messages différés.
- `PlanningConsolePolicy` est interactive par défaut ; seul un harnais active le mode non interactif.
- `WBSWriteGuard` utilise des scopes tokenisés LIFO possédés par l'appelant.
- Le Core reste la source unique du calcul planning.
- TEST, SCENARIO et LOCK restent des services frères.
- Un reset est toujours demandé au propriétaire du store.
- Un fallback de sécurité ne doit jamais devenir un PASS générique.

## Harnais et niveau de preuve

| Harnais | Ce qu'il prouve | À lancer lorsque |
|---|---|---|
| WBS Write Guard | scopes, imbrication, erreurs, état final | guard, writer, Runtime ou LOCK |
| RuntimeWorkflow / RunButtons | workflows complets non interactifs | wrappers `Run_*`, MacroGuard ou console policy |
| MessageEngine / EventHistory | filtrage, regroupement, ACK, historique | diagnostics ou console |
| Diagnostic Producers | STOP/WARNING/INFO de bout en bout | producteurs CoreBridge/Constraints/S-Curve |
| Gantt Visual Regression | signature Shapes et feuille | renderer, layout ou UI Gantt |
| Predictive Registry | fast path, fallback, reuse, Lazy Repair | registry et renderers spécialisés |
| TEST / fallback / SCENARIO | transactions de simulation | GanttLive et services simulation |
| LOCK instrumenté | succès durable sur copie, source intacte | LOCK, writer WBS ou guard |
| Incremental Signature | compatibilité bit à bit | signature, CalcState ou Incremental |

## Localiser rapidement une modification

| Besoin | Commencer par |
|---|---|
| règle de date ou lag | `mod_CalendarEngine`, puis Core |
| dépendance FS/SS/FF | Parsed Network, `mod_CalcCoreNetwork`, Core |
| nouveau warning | producteur propriétaire, puis contrats MessageEngine/EventHistory |
| nouvelle colonne | DataSync, Pre-Core, contrat Core, writers, Incremental Signature |
| rendu d'une barre | `mod_GanttRenderer`, Geometry, ShapeRegistry |
| lien Gantt | `mod_GanttDependencyRenderer` |
| drag/resize | `mod_GanttDragWatch`, transaction TEST/SCENARIO |
| KPI Dashboard | Dashboard Read Context, puis `mod_Dashboard` |
| série S-Curve | `mod_SCurve` |
| bouton ou callback | module UI propriétaire et registre des callbacks |

## Compréhension en une heure

Après le parcours de 15 minutes, lire les en-têtes des modules du domaine ciblé, puis seulement leurs APIs Public. Utiliser `MODULE_AND_PROCEDURE_DOCUMENTATION_COVERAGE.tsv` pour retrouver un composant et `NAMING_AUDIT_AND_RETAINED_LEGACY_CONTRACTS.tsv` pour identifier les contrats historiques. N'ouvrir les helpers Private que lorsque le contrat Public ne suffit pas à expliquer l'invariant recherché.
