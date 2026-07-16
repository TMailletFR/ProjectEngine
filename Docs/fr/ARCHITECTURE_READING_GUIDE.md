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
