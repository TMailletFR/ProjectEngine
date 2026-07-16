# Architecture Reading Guide

## 15-minute tour

Read the following in order to understand project boundaries before entering detailed algorithms:

1. `mod_RunButtons`: stable user-facing `Run_*` macros.
2. `mod_RuntimeWorkflow`, `mod_MacroGuard`, `mod_PlanningConsolePolicy`: command lifecycle.
3. `mod_CalcEngineCoreBridge`: planning orchestration.
4. `mod_DataSync`, then `mod_CalcCoreProdWrapper`: WBS/CALC to Core transition.
5. `mod_CalcCoreEngine`, `mod_CalcCoreNetwork`, `mod_CalendarEngine`: the single calculation engine.
6. `mod_CoreBridgeAnalytics` and `mod_CoreBridgeOutputWriter`: analytics and outputs.
7. `mod_GanttRefreshPipeline`, `mod_GanttRenderer`, `mod_GanttShapeRegistry`: Gantt rendering.
8. `mod_GanttLive` and TEST/SCENARIO/LOCK services: simulation.
9. `mod_MessageEngine`, `mod_EventHistory`, `frmPlanningMessages`: diagnostics reaching the user.
10. `PROJECT_GLOSSARY.md` and `MAINTENANCE_GUIDE.md`: terminology and change rules.

## Main flow

```text
Excel callback / Run_* macro / OnAction
    -> RuntimeWorkflow + MacroGuard
    -> DataSync (WBS -> CALC + LOGIC_LINKS)
    -> Pre-Core validation
    -> CalcCoreProdWrapper
    -> CalcCoreEngine + CalcCoreNetwork + CalendarEngine
    -> CoreBridgeOutputWriter (CALC then WBS, under WBS Write Guard)
    -> CoreBridgeAnalytics / Variances
    -> Gantt / S-Curve / Dashboard refreshes
    -> MessageEngine
    -> PlanningConsolePolicy
    -> EventHistory / ACK + frmPlanningMessages
```

Arrows show orchestration direction. Domains access one another through public contracts, never through private state. Diagnostics flow toward MessageEngine; they do not flow back into calculation.

## Update Planning workflow

1. `Run_Planning_Update` opens MacroGuard and a runtime workflow.
2. `Run_Calc_Engine_CoreBridge` handles Safe Empty State, prepares infrastructure and synchronizes tables.
3. Pre-Core validation emits STOP diagnostics without creating another engine.
4. `Run_Calc_Core_PROD_Pilot` prepares the working dataset and parsed network.
5. `Run_Calc_Core` calculates leaf tasks, propagates errors and applies LOE post-processing.
6. Writers persist CALC and then WBS in contractual order.
7. Analytics calculates paths, floats, deadlines, variances and warnings.
8. MessageEngine prepares console output; Runtime may defer display to the root workflow.

## WBS -> CALC -> Core -> outputs

| Stage | Owner | Data | Invariant |
|---|---|---|---|
| User input | WBS and `mod_WBSEvents` | `tbl_WBS` | Calculated columns remain protected. |
| Synchronization | `mod_DataSync` | `tbl_CALC`, `tbl_LOGIC_LINKS` | WBS supplies inputs; CALC is the engine dataset. |
| Identity | `mod_CanonicalIdentityIndex` | ID/WBS/row maps | Exposed maps are read-only. |
| Network | `mod_ParsedPlanningNetwork` | Succ/Pred/Type/Lag | Parsing is shared; business projections remain separate. |
| Calculation | `mod_CalcCoreEngine` | mutable Core array | There is only one planning engine. |
| Persistence | `mod_CoreBridgeOutputWriter` | CALC then WBS | Full and Partial preserve fields and write order. |
| WBS protection | `mod_WBSWriteGuard` | tokenized scopes | A caller closes only its own token in LIFO order. |

## Gantt domain

`Refresh_Gantt` is the stable public wrapper. `mod_GanttRefreshPipeline` acquires data and selects Full or Display Only processing. Renderers receive prepared arrays and maps.

- `mod_GanttRenderer` draws tasks, summaries, milestones and the today line.
- `mod_GanttDependencyRenderer` routes and draws dependencies.
- `mod_GanttConstraintRenderer` draws constraints and deadlines.
- `mod_GanttShapeRegistry` owns Shape records, cache and predictive diff.
- `mod_GanttGeometry` and `mod_GanttTimelineGeometry` provide pure calculations.
- `mod_GanttUiControls`, `mod_GanttViewState` and `mod_GanttLanguage` own UI concerns, not calculation.

Sensitive areas:

- Shape names, `OnAction`, z-order and geometry;
- Day predictive fast path, Week/Month fallback and Lazy Repair;
- Drag watcher and timer lifecycle;
- consistency between expected registry and actual sheet state.

## TEST, SCENARIO and LOCK

| Mode | Owner | Input | Output | Main prohibition |
|---|---|---|---|---|
| TEST | `mod_GanttTestService` | yellow TEST cells | `tbl_CALC_GANTT_TEST`, predictive overlay | no durable WBS write |
| SCENARIO | `mod_GanttScenarioService` | planning or scenario copy | scenario dataset and rendering | must not use TEST as a parent engine |
| LOCK | `mod_GanttLockService` | validated simulation | durable WBS Forecast values | must never bypass WBS Write Guard |

`mod_GanttLive` retains historical wrappers and public transactions. `mod_GanttSimulationState` owns mode and render requests. `mod_GanttSimulationTableStore` owns the schema and reset of `tbl_CALC_GANTT_TEST`. Scenario Fork retains its `Application.Run` contract.

## S-Curve and Dashboard

`mod_SCurve` is the single time-series engine and owns its outputs. `SCurve_BuildDashboardProjection` exposes a Dashboard-specific projection.

`mod_DashboardReadContext` acquires WBS, CALC and that projection once for all three Dashboard modes. `mod_Dashboard` keeps Full Build, Content Only and Texts/Comparison rendering policies separate.

## Diagnostics, console, EventHistory and ACK

```text
CoreBridge / Constraints / S-Curve producers
    -> structured message collections
    -> MessageEngine filtering and grouping
    -> PlanningConsolePolicy
       -> interactive mode: frmPlanningMessages.Show vbModal
       -> harness mode: capture without display
    -> EventHistory logging
    -> optional warning ACK without deleting history
```

The producer decides diagnostic meaning and severity. MessageEngine prepares and groups without recalculation. EventHistory owns storage and ACK state. The UserForm only displays an already prepared projection.

## Stores, snapshots and canonical contracts

| Component | Owns | Does not own |
|---|---|---|
| Canonical Identity Index | ID, normalized WBS, row indexes, Driving Logic | business hierarchy or calculation |
| Parsed Planning Network | immutable link parsing | topological sort, validation or Gantt routing |
| Incremental Signature | 17 fields, order, normalization, serialization | CALC_STATE or recalculation decisions |
| CalcState | incremental snapshot persistence | signature definition |
| Dashboard Read Context | shared acquisition for one refresh | three-mode rendering |
| Simulation Table Store | `tbl_CALC_GANTT_TEST` and reset | TEST/SCENARIO/LOCK policy |

## Guards and safety invariants

- `MacroGuard` prevents concurrent execution and carries abort requests.
- `RuntimeWorkflow` maintains depth, root workflow and deferred messages.
- `PlanningConsolePolicy` is interactive by default; only harnesses enable noninteractive mode.
- `WBSWriteGuard` uses caller-owned, tokenized LIFO scopes.
- Core remains the single source of planning calculation.
- TEST, SCENARIO and LOCK remain sibling services.
- Resets are always requested from the store owner.
- A safety fallback must never be converted into a generic PASS.

## Harnesses and proof level

| Harness | What it proves | Run when changing |
|---|---|---|
| WBS Write Guard | scopes, nesting, errors, final state | guard, writer, Runtime or LOCK |
| RuntimeWorkflow / RunButtons | complete noninteractive workflows | `Run_*` wrappers, MacroGuard or console policy |
| MessageEngine / EventHistory | filtering, grouping, ACK, history | diagnostics or console |
| Diagnostic Producers | end-to-end STOP/WARNING/INFO | CoreBridge/Constraints/S-Curve producers |
| Gantt Visual Regression | Shape and sheet signature | Gantt renderer, layout or UI |
| Predictive Registry | fast path, fallback, reuse, Lazy Repair | registry and specialized renderers |
| TEST / fallback / SCENARIO | simulation transactions | GanttLive and simulation services |
| Instrumented LOCK | durable success on a copy, unchanged source | LOCK, WBS writer or guard |
| Incremental Signature | bit-for-bit compatibility | signature, CalcState or Incremental |

## Quickly locating a change

| Need | Start with |
|---|---|
| date or lag rule | `mod_CalendarEngine`, then Core |
| FS/SS/FF dependency | Parsed Network, `mod_CalcCoreNetwork`, Core |
| new warning | owning producer, then MessageEngine/EventHistory contracts |
| new column | DataSync, Pre-Core, Core contract, writers, Incremental Signature |
| task-bar rendering | `mod_GanttRenderer`, Geometry, ShapeRegistry |
| Gantt dependency | `mod_GanttDependencyRenderer` |
| drag/resize | `mod_GanttDragWatch`, TEST/SCENARIO transaction |
| Dashboard KPI | Dashboard Read Context, then `mod_Dashboard` |
| S-Curve series | `mod_SCurve` |
| button or callback | owning UI module and callback registry |

## One-hour understanding

After the 15-minute tour, read module headers in the target domain and then only their Public APIs. Use `MODULE_AND_PROCEDURE_DOCUMENTATION_COVERAGE.tsv` to locate components and `NAMING_AUDIT_AND_RETAINED_LEGACY_CONTRACTS.tsv` to identify historical contracts. Open Private helpers only when the Public contract does not sufficiently explain the required invariant.
