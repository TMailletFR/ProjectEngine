# Project Glossary

This vocabulary describes how terms are actually used in this workbook. It is the naming convention for future changes.

| Term | Definition | May do | Must not do | Actual example |
|---|---|---|---|---|
| **Bridge** | Boundary translating Excel tables into engine contracts and orchestrating their exchange. | Adapt structures and delegate. | Recalculate a Core rule or draw UI. | `mod_CalcEngineCoreBridge` |
| **Context** | Coherent execution state shared during one specific workflow. | Group state sharing the same lifecycle. | Become a universal parameter bag. | contexte de `mod_RuntimeWorkflow` |
| **Data Provider** | Owner-operated reader of a projection required by one domain. | Read, validate schema and build maps. | Decide rendering or business calculation. | `mod_GanttDataProvider`, `mod_GanttLiveDataProvider` |
| **Diagnostics** | Transformation of a result or violation into a structured message. | Build code, severity, text and details. | Recalculate the business result. | `mod_CoreBridgeDiagnostics` |
| **Engine** | Component executing business calculation logic. | Calculate and produce an output dataset. | Depend on Shapes, UserForms or Excel callbacks. | `mod_CalcCoreEngine`, `mod_CalendarEngine` |
| **Facade** | Stable contract shielding callers from evolving internal organization. | Delegate to the actual owner while reducing coupling. | Be a pass-through without a stability requirement. | wrappers publics de `mod_GanttLive` |
| **Harness / Harnais** | Reproducible proof code executed on a copy or controlled state. | Capture, compare, trace and fail explicitly. | Participate in user workflows or hide timeouts. | `mod_GanttVisualRegression`, `mod_WBSWriteGuardHarness` |
| **Pipeline** | Ordered sequence of stages delegated to their owners. | Enforce order and manage orchestration exits. | Reimplement delegated stages. | `mod_GanttRefreshPipeline` |
| **Projection** | Consumer-oriented view that is not the canonical dataset itself. | Filter and shape read data for a consumer. | Be shared when consumer policies differ. | projection Dashboard de `mod_SCurve` |
| **Read Model** | Coherent read-only representation shared because its semantics are identical. | Centralize duplicated parsing, identity or persistent contracts. | Hide a few orchestration lines or merge policies. | Canonical Identity Index, Parsed Planning Network |
| **Registry** | Render-lifetime index used to locate and compare expected Shapes. | Create, reuse, invalidate and repair its records. | Read WBS/CALC as a general business source. | `mod_GanttShapeRegistry` |
| **Renderer** | Component producing cells, Shapes or charts from prepared data. | Calculate visual geometry and apply style. | Decide TEST/SCENARIO or recalculate planning. | `mod_GanttRenderer`, `mod_GanttDependencyRenderer` |
| **Rules** | Owner of a reusable pure business policy. | Normalize and classify through one shared rule. | Read or write tables on behalf of consumers. | `mod_TaskTypeRules` |
| **Service** | Owner of one coherent business workflow with an explicit API. | Orchestrate its stages and use external owners. | Own data or policies belonging to called services. | `mod_GanttTestService`, `mod_ConstrDiagService` |
| **Snapshot** | Coherent capture of state at a defined time or version. | Be compared, persisted or consumed read-only. | Be silently mutated by consumers. | `CALC_STATE`, Dashboard snapshots |
| **Store** | Owner of a store, its schema and reset contract. | Read, write, resize and reset owned storage. | Decide why an orchestrator requests a reset. | `mod_GanttSimulationTableStore` |
| **Workflow** | User or system transaction with start, end, nesting and display policy. | Orchestrate domains and retain runtime state. | Own internal implementations of called domains. | `mod_RuntimeWorkflow` |
| **Wrapper** | Stable name retained for a callback or historical contract. | Delegate without changing name, signature or behavior. | Exist only to avoid a safe local migration. | `Run_Gantt_Test_Engine` |
| **Owner / Propriétaire** | Sole component allowed to know a datum's internal schema or state. | Expose targeted business contracts. | Expose generic GetCell-style access to storage. | EventHistory pour EVENT_HISTORY et ACK |
| **Canonical / Canonique** | Single source of a contract that is strictly equivalent across uses. | Define shared ordering, parsing or identity once. | Absorb intentional business differences. | `mod_CanonicalIdentityIndex`, `mod_IncrementalSignature` |
| **Full / Partial** | Two scopes of one writer: complete dataset or impacted IDs. | Share output contracts while retaining write strategies. | Change fields or ordering merely to unify code. | Full/Partial Output Writer |
| **TEST** | Simulation driven by yellow input cells, without durable WBS mutation. | Call the existing Core and feed predictive rendering. | Write WBS durably or become another Core. | `mod_GanttTestService` |
| **SCENARIO** | Full simulation projected from current planning or a scenario copy. | Build and render its complete simulation dataset. | Depend on TEST as a parent engine. | `mod_GanttScenarioService` |
| **LOCK** | Durable transaction applying a validated simulation result to WBS. | Backup, validate, write, finalize or roll back. | Bypass WBS Write Guard or duplicate simulation engines. | `mod_GanttLockService` |
| **Safe Empty State** | Safe visual and data state when the project has no usable task. | Clear owned outputs and keep the workbook usable. | Act as a generic destructive reset. | `Planning_FullSafeEmptyState`, `SCurve_SafeEmptyState` |
| **Incremental** | Calculation limited to changed tasks and dependants according to a persisted signature. | Compare canonical signatures and build impacted scope. | Change the 17-field contract without controlled versioning. | `mod_CalcIncremental`, `mod_IncrementalSignature` |
| **Driving Logic** | Reference to the logic or predecessor driving a task's calculated date. | Be produced by calculation and consumed by analytics/rendering. | Be reconstructed differently by each consumer. | Canonical Identity Driving Logic map |
| **Blocking Error** | Error preventing Core outputs from being considered valid. | Produce STOP, propagation and structured diagnostics. | Be downgraded to a warning by rendering code. | `Core_AddBlockingError` |
| **ACK** | Acknowledgement token used to hide an accepted warning without deleting its history. | Remain stable, persisted and linked to a diagnostic hash. | Change severity or delete the source event. | EVENT_ACK / EventHistory |

## Selection rule

Use the narrowest term that describes an owned responsibility. If a new component needs several conflicting terms, its boundary is probably too broad.
