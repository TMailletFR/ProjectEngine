# ProjectEngine

Excel-native, dependency-driven planning engine focused on transparent calculations, controlled simulation, actionable diagnostics, and project controls workflows.

ProjectEngine transforms a standard Excel workbook into a lightweight scheduling and project controls platform powered entirely by VBA. It combines a single-source planning engine, incremental recalculation, network analytics, interactive Gantt rendering, S-Curve reporting, constraints management, scenario simulation, and traceable runtime diagnostics in one portable workbook.

The workbook remains familiar to Excel users while calculation, rendering, diagnostics, and mutation workflows are separated into explicit components with guarded boundaries.

---

## Screenshots

### WBS Planning Inputs

![WBS Planning Inputs](Screenshots/wbs-inputs.png)

Excel-native planning inputs for hierarchy, task classification, dependencies, baseline, actuals, forecast, progress, and calculated outputs.

### Interactive Gantt

![Interactive Gantt](Screenshots/hero-gantt-sc.png)

Dependency-driven scheduling with progress tracking, milestones, summary tasks, critical path analysis, and real-time project status visualization.

### Executive Dashboard

![Executive Dashboard](Screenshots/dashboard.png)

High-level project overview including progress indicators, forecast finish analysis, critical activities, schedule momentum, S-Curve snapshot, and planning overview.

### Hot Spots Analysis

![Hot Spots Analysis](Screenshots/dashboard2.png)

Dedicated planning analytics highlighting top delays, deadline health, upcoming milestones, and next critical activities requiring attention.

### Planning Analytics

![Planning Analytics](Screenshots/analytics.png)

Detailed schedule analysis covering variances, Driving Logic, Critical and Longest Path, Total and Free Float, REX comparisons, and Deadline Float.

### Month View

![Month View](Screenshots/gantt-month.png)

Long-range schedule visualization with monthly scaling, WBS hierarchy, task classifications, constraints, dependencies, milestones, and progress status.

### S-Curve Analytics

![S-Curve Analytics](Screenshots/s-curve.png)

Baseline, Actual, Forecast, and Calculated progress curves combined with daily workload distribution for project controls reporting and performance analysis.

### Diagnostic Console

![Diagnostic Console](Screenshots/event-console.png)

Structured INFO, WARNING, and STOP diagnostics with bilingual messages, navigation, Event History, and acknowledgement controls.

### Animated Demo

![ProjectEngine Demo](Screenshots/hero-gantt.gif)

---

## Core Features

### Scheduling Engine

* Dependency-driven scheduling
* FS / SS / FF relationships
* Positive and negative lag support
* Multi-predecessor logic
* Automatic schedule propagation
* Parent / child rollups
* Level of Effort (LOE) support
* Milestone support
* 5-day and 6-day working calendars
* Cycle and missing-predecessor validation

### Planning Logic

* Baseline planning
* Forecast planning
* Actual progress integration
* Calculated schedule generation
* Signature-based incremental recalculation
* Forced full recalculation
* Full and partial output synchronization
* Protected WBS write lifecycle

### Analytics

* Critical Path analysis
* Longest Path analysis
* Total Float calculation
* Free Float calculation
* Multi-project criticality mode
* Delay analysis
* Deadline monitoring
* Schedule momentum tracking
* Parent-date and task-type warnings
* Baseline and forecast variance analysis

### Constraints & Controls

* Hard constraints engine
* Deadline management
* Constraint diagnostics
* Constraint impact analysis
* Dedicated constraints table
* Warning acknowledgement (ACK)

### Visualization

* Interactive Gantt chart
* Day / Week / Month timeline scaling
* Dependency link rendering
* Critical path overlays
* Progress overlays
* Constraint overlays
* Deadline markers
* Predictive Shape Registry and lazy repair
* Executive Dashboard
* S-Curve reporting

### Simulation

* TEST mode for focused, non-destructive changes
* SCENARIO mode for complete schedule alternatives
* LOCK workflow for applying a validated simulation to WBS
* Schedule comparison workflows
* Drag and resize simulation from the Gantt

### Diagnostics

* Runtime warning console
* INFO / WARNING / STOP severities
* Event History
* Alarm History
* Warning acknowledgement workflow
* Stable event signature tracking
* Input validation framework
* Grouped and bilingual planning messages

### Maintenance & Safety

* Safe Empty State for Gantt
* Safe Empty State for S-Curve
* Planning Reset workflow
* Full Reset workflow
* Protected calculated columns
* Tokenized, caller-owned write scopes
* Deterministic visual and workflow regression harnesses
* Noninteractive validation on isolated workbook copies

---

## Architecture

ProjectEngine is organized around explicit domains rather than one monolithic macro.

| Domain | Responsibility |
| --- | --- |
| Runtime Workflow | Macro lifecycle, nesting, abort handling, and deferred display |
| WBS / DataSync | User inputs and synchronization into canonical calculation tables |
| Core Calculation | Dependency network, calendars, constraints, task dates, LOE, and rollups |
| Analytics | Critical/Longest Path, float, deadlines, variances, and schedule health |
| Output Writers | Deterministic full or partial writes to CALC and protected WBS columns |
| Gantt | Refresh pipeline, renderers, geometry, Shape Registry, interaction, and Drag |
| Simulation | Sibling TEST, SCENARIO, and LOCK services using the existing Core |
| Dashboard / S-Curve | Executive reporting and time-phased progress analytics |
| Diagnostics | Message preparation, console policy, Event History, alarms, and ACK |
| Canonical Contracts | Shared identity indexes, parsed planning network, and incremental signatures |

The scheduling engine is the single source of truth. Simulation services call the same Core; they do not implement competing calculation engines. Visualization layers consume calculated projections but never drive schedule calculations.

### Design Principles

* Excel remains the user interface and portable runtime.
* Business calculation is separated from rendering and UserForms.
* Data stores expose owner-controlled contracts instead of generic table access.
* Public callbacks and workbook entry points remain stable wrappers.
* Safety-critical writes use explicit guards and caller-owned scopes.
* Refactors are validated with deterministic, copy-based harnesses.

---

## Documentation

Detailed architecture, maintenance, and project terminology are available in dedicated developer documentation:

- [Documentation française](Docs/fr/README.md)
- [English documentation](Docs/en/README.md)

---

## Typical Workflow

1. Enter tasks, hierarchy, calendars, dates, progress, and dependencies in WBS.
2. Define hard constraints or deadlines when required.
3. Run Planning Update to synchronize WBS, calculate the network, and publish outputs.
4. Review STOP, WARNING, and INFO diagnostics in the planning console and Event History.
5. Analyze schedule health through Critical/Longest Path, float, deadlines, Dashboard, and S-Curve views.
6. Test focused changes through yellow TEST inputs or build a complete SCENARIO alternative.
7. Review predictive Gantt differences and lock validated changes back into WBS when appropriate.
8. Refresh Gantt, S-Curve, or Dashboard independently for reporting.

---

## Project Status

### Current Stage

**Beta**

The engine is actively used and continuously improved through real-world planning use cases.

Scheduling, incremental calculation, analytics, simulation, Gantt, Dashboard, S-Curve, diagnostics, ACK, reset, and guarded write workflows are operational.

The current architecture is documented and protected by deterministic validation harnesses. Future releases will continue improving usability, visualization, and advanced planning capabilities without introducing competing calculation paths.

---

## Support the Project

If ProjectEngine helps your work and you would like to support ongoing development:

![QRCode](Screenshots/qrcode.png)

---

## License

GPL-3.0
