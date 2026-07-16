# Glossaire du projet

Ce vocabulaire décrit l'usage réel des termes dans ce classeur. Il constitue la convention de nommage pour les futures évolutions.

| Terme | Définition | Peut faire | Ne doit pas faire | Exemple réel |
|---|---|---|---|---|
| **Bridge** | Frontière traduisant les tables Excel en contrats moteur et orchestrant leur échange. | Adapter les structures et déléguer. | Recalculer une règle Core ou dessiner l'UI. | `mod_CalcEngineCoreBridge` |
| **Context** | État d'exécution cohérent partagé pendant un workflow précis. | Regrouper un état ayant le même lifecycle. | Devenir un sac universel de paramètres. | contexte de `mod_RuntimeWorkflow` |
| **Data Provider** | Lecteur propriétaire d'une projection destinée à un domaine. | Lire, valider le schéma et construire des maps. | Décider le rendu ou le calcul métier. | `mod_GanttDataProvider`, `mod_GanttLiveDataProvider` |
| **Diagnostics** | Transformation d'un résultat ou d'une violation en message structuré. | Construire code, sévérité, texte et détails. | Recalculer le résultat métier. | `mod_CoreBridgeDiagnostics` |
| **Engine** | Composant qui exécute une logique métier de calcul. | Calculer et produire un dataset de sortie. | Dépendre des Shapes, UserForms ou callbacks Excel. | `mod_CalcCoreEngine`, `mod_CalendarEngine` |
| **Facade** | Contrat stable protégeant les appelants d'une organisation interne appelée à évoluer. | Déléguer vers le propriétaire réel avec une valeur de découplage. | Être un passe-plat sans consommateur ou stabilité requise. | wrappers publics de `mod_GanttLive` |
| **Harness / Harnais** | Code de preuve reproductible exécuté sur une copie ou un état contrôlé. | Capturer, comparer, tracer et échouer explicitement. | Participer à un workflow utilisateur ou masquer un timeout. | `mod_GanttVisualRegression`, `mod_WBSWriteGuardHarness` |
| **Pipeline** | Séquence ordonnée d'étapes déléguées à leurs propriétaires. | Imposer l'ordre, gérer la sortie et le rollback orchestral. | Réimplémenter les étapes appelées. | `mod_GanttRefreshPipeline` |
| **Projection** | Vue de données adaptée à un consommateur sans être la donnée canonique. | Filtrer et façonner des données de lecture. | Être réutilisée si les politiques des consommateurs divergent. | projection Dashboard de `mod_SCurve` |
| **Read Model** | Représentation de lecture cohérente, sans mutation métier, partagée parce que sa sémantique est identique. | Centraliser parsing, identité ou contrat persistant dupliqué. | Masquer quelques lignes d'orchestration ou fusionner des politiques. | Canonical Identity Index, Parsed Planning Network |
| **Registry** | Index durable pendant un rendu permettant de retrouver et comparer les Shapes attendues. | Créer, réutiliser, invalider et réparer ses records. | Lire WBS/CALC comme source métier générale. | `mod_GanttShapeRegistry` |
| **Renderer** | Composant produisant des cellules, Shapes ou graphiques à partir de données préparées. | Calculer la géométrie visuelle et appliquer le style. | Décider TEST/SCENARIO ou recalculer le planning. | `mod_GanttRenderer`, `mod_GanttDependencyRenderer` |
| **Rules** | Propriétaire d'une politique métier pure et réutilisée. | Normaliser et classifier selon une règle unique. | Lire ou écrire une table au nom des consommateurs. | `mod_TaskTypeRules` |
| **Service** | Propriétaire d'un workflow métier cohérent avec une API explicite. | Orchestrer ses étapes et utiliser des propriétaires externes. | Posséder les données ou politiques des services appelés. | `mod_GanttTestService`, `mod_ConstrDiagService` |
| **Snapshot** | Capture cohérente d'un état à un instant ou une version déterminée. | Être comparé, persisté ou consommé en lecture. | Être modifié silencieusement par ses consommateurs. | `CALC_STATE`, Dashboard snapshots |
| **Store** | Propriétaire d'un stockage, de son schéma et de son reset. | Lire, écrire, redimensionner et vider son stockage. | Décider pourquoi un orchestrateur demande le reset. | `mod_GanttSimulationTableStore` |
| **Workflow** | Transaction utilisateur ou système avec début, fin, profondeur et politique d'affichage. | Orchestrer des domaines et conserver son état runtime. | Posséder les implémentations internes des domaines. | `mod_RuntimeWorkflow` |
| **Wrapper** | Nom stable conservé pour un callback ou un contrat historique. | Déléguer sans altérer signature ni comportement. | Être créé uniquement pour éviter une migration locale sûre. | `Run_Gantt_Test_Engine` |
| **Owner / Propriétaire** | Seul composant autorisé à connaître le schéma ou l'état interne d'une donnée. | Exposer des contrats métier ciblés. | Exposer `GetCell` ou un accès générique à son stockage. | EventHistory pour EVENT_HISTORY et ACK |
| **Canonical / Canonique** | Source unique d'un contrat strictement équivalent dans tous ses usages. | Définir ordre, parsing ou identité communs. | Absorber des différences métier volontaires. | `mod_CanonicalIdentityIndex`, `mod_IncrementalSignature` |
| **Full / Partial** | Deux portées d'un même writer : dataset complet ou IDs impactés. | Partager un contrat de sortie tout en gardant leurs stratégies. | Modifier l'ordre ou les champs écrits pour uniformiser le code. | Full/Partial Output Writer |
| **TEST** | Simulation issue des cellules jaunes, sans mutation durable de WBS. | Appeler le Core existant et alimenter le rendu prédictif. | Écrire durablement WBS ou devenir un second Core. | `mod_GanttTestService` |
| **SCENARIO** | Simulation complète projetée depuis le planning courant ou une copie scénario. | Construire son dataset et rendre ses résultats. | Dépendre du service TEST comme moteur parent. | `mod_GanttScenarioService` |
| **LOCK** | Transaction durable appliquant à WBS un résultat de simulation validé. | Sauvegarder, valider, écrire, finaliser ou rollback. | Contourner WBS Write Guard ou dupliquer TEST/SCENARIO. | `mod_GanttLockService` |
| **Safe Empty State** | État visuel et de données sûr lorsque le projet n'a aucune tâche exploitable. | Vider uniquement les sorties possédées et laisser le classeur utilisable. | Être utilisé comme reset destructif générique. | `Planning_FullSafeEmptyState`, `SCurve_SafeEmptyState` |
| **Incremental** | Calcul limité aux tâches modifiées et à leurs dépendants selon une signature persistée. | Comparer le contrat canonique et construire le scope impacté. | Modifier la définition des 17 champs hors version contrôlée. | `mod_CalcIncremental`, `mod_IncrementalSignature` |
| **Driving Logic** | Référence de la logique ou du prédécesseur qui pilote la date calculée d'une tâche. | Être produite par le calcul et consommée par analytics/rendu. | Être reconstruite différemment par chaque consommateur. | Canonical Identity Driving Logic map |
| **Blocking Error** | Erreur empêchant de considérer les sorties Core comme valides. | Produire STOP, propagation et diagnostics structurés. | Être transformée en simple warning par un renderer. | `Core_AddBlockingError` |
| **ACK** | Token d'acquittement permettant de masquer un warning déjà reconnu sans supprimer son historique. | Être stable, persisté et associé au hash du diagnostic. | Modifier la sévérité ou effacer l'événement source. | EVENT_ACK / EventHistory |

## Règle de choix

Utiliser le terme le plus étroit qui décrit une responsabilité possédée. Si un nouveau composant nécessite plusieurs termes contradictoires, sa frontière est probablement trop large.
