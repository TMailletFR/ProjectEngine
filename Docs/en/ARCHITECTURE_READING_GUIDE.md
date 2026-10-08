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

## Ribbon and welcome entry points (v1.3.0)

`customUI.xml` is embedded in each `.xlsm` with eighteen image-backed commands. `mod_RibbonCallbacks` is a **thin UI adapter**, not a scheduling owner. It validates the originating workbook and active window before dispatching to existing public operations. WBS update commands are WBS-contextual; Gantt Test/Scenario/Lock/Reset are Gantt-contextual; Settings, Navigation, Reset and Import have their defined global scopes. `ThisWorkbook` coordinates lifecycle events, while activation refreshes Ribbon state. **Do not invalidate Ribbon controls during `Workbook_Deactivate`** (coexistence with another open ProjectEngine workbook).

For an eligible empty, writable workbook, `mod_ProjectWelcome` displays `frmProjectWelcome` via `ProjectWelcome_ShowIfNeeded`. Start New acknowledges welcome through the existing EventHistory/ACK owner; import shares the same migration entry as Ribbon. Closing/cancelling/failing does not acknowledge successful onboarding. Full Reset makes welcome eligible again. This form is a dedicated view, not a new message store, planning engine or translation catalogue.

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

## Importing a previous ProjectEngine workbook

```text
Ribbon Import / Welcome Import
    -> ProjectWelcome_ImportPrevious
    -> Migration_ImportPrevious / Migration_ImportFile
    -> validate source/destination/schema and show destructive confirmation
    -> open source read-only, macros disabled, links not updated
    -> MigrationData_Read (supported inputs and selected persistent stores)
    -> close source without saving
    -> Migration_ApplyInputs (replace destination's supported data)
    -> existing Settings / WBS formula / Constraints / Dashboard / EventHistory owners
    -> verify resulting data and empty derived outputs
    -> success receipt + welcome acknowledgement, then one destination Save
    -> post-save verification
```

The import is **data-only**: it does not run Core, rebuild the Gantt, calculate the S-Curve or execute source macros. The user reviews WBS and Constraints, then explicitly runs Planning Update or Full Update. Supported data include WBS inputs (including supported input formulas), Constraints, Settings, EVENT_ACK and Dashboard snapshots; compatible history can be preserved, incompatible optional history skipped explicitly. The target retains its own VBA, Ribbon and derived data contracts.

**Current v1.3.0 transaction boundary:** this is a direct destructive import with no automatically created destination recovery backup or staging workbook. The source remains unchanged. Ask the user to back up the destination **before** starting. An error before the final Save must not be reported as SUCCESS or create a success ACK; do not call the workflow atomically rollback-safe against Excel/AutoSave or an interruption after Save.

`ProjectEngine.SchemaVersion` describes the persistent schema; `ProjectEngine.ReleaseId` identifies a release separately and is not auto-invented by import. Older unmarked workbooks use validated structure/profile detection. Unsupported/future schemas and unsafe inputs fail closed; source workbooks must never be executed as migration code.

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

## Full Update and output publication

`Run_Full_Update` still forces complete Core scheduling and analytics. The v1.3.0 output writer **compares already-calculated output columns** with existing WBS values and skips only columns that are identical and contain no formulas. Differing columns, or any formula in a calculated output column, follow the normal bulk rewrite path. This is not an incremental-Core shortcut and does **not** promise a faster first full reconstruction. The same single output authority remains in place.

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

## Gantt navigation, readiness and safe reuse

`mod_GanttNavigation` resolves Task IDs and visible targets; `mod_GanttViewState` owns Detail/Summary projection. Navigation may call the existing `EnsureGanttForCurrentPlanning` preparation when needed, but does not silently recalculate a stale planning state. Show Full Timeline fits the rendered extent subject to Excel zoom limits and uses an explicit fallback when fitting is physically impossible.

Rendering uses the existing `mod_GanttRefreshPipeline`, `mod_GanttRenderer`, `mod_GanttDependencyRenderer` and `mod_GanttDependencySvg` owners. Deferred rendering may retain existing route/shape state as a **candidate**, never as unconditional READY. Reuse is permitted only after input/context and physical-state checks; changed task geometry must have complete affected-route coverage. Structural, constraint, timeline-layout or untracked physical changes cause a conservative canonical FULL fallback. Day renders dependency routes; aggregated Week/Month scales use their existing visibility policy. No second renderer, route store or READY authority was introduced.

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

## Shared localization and user-facing text

The logical shared TextCatalog resolves a **TextKey + supplied LanguageKey** to text. The catalogue does not own the language. Six independent language owners are retained; GLOBAL coordinates changes without overriding each domain's authority. Ribbon, welcome, import and navigation consume this common localization infrastructure. Use explicit, testable fallbacks and avoid coupling the logical API to CP1252 storage.

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
