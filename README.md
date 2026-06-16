# ProjectEngine

Excel-native dependency-driven planning engine focused on transparency, simulation, diagnostics, and project controls workflows.

ProjectEngine transforms a standard Excel workbook into a lightweight scheduling and project controls platform powered entirely by VBA. It combines dependency-driven schedule calculation, planning analytics, Gantt visualization, S-Curve reporting, constraints management, simulation tools, and runtime diagnostics in a single portable workbook.

---

## Screenshots

### Interactive Gantt

![Interactive Gantt](Screenshots/hero-gantt-sc.png)

Dependency-driven scheduling with progress tracking, milestones, summary tasks, critical path analysis, and real-time project status visualization.

### Executive Dashboard

![Executive Dashboard](Screenshots/dashboard.png)

High-level project overview including progress indicators, forecast finish analysis, critical activities, schedule momentum, S-Curve snapshot, and planning overview.

### Hot Spots Analysis

![Hot Spots Analysis](Screenshots/dashboard2.png)

Dedicated planning analytics highlighting top delays, deadline health, upcoming milestones, and next critical activities requiring attention.

### Month View

![Month View](Screenshots/gantt-month.png)

Long-range schedule visualization with monthly scaling, WBS hierarchy, task classifications, constraints, dependencies, milestones, and progress status.

### S-Curve Analytics

![S-Curve Analytics](Screenshots/s-curve.png)

Baseline, Actual, Forecast, and Calculated progress curves combined with daily workload distribution for project controls reporting and performance analysis.

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

### Planning Logic

* Baseline planning
* Forecast planning
* Actual progress integration
* Calculated schedule generation
* Incremental recalculation
* Forced full recalculation

### Analytics

* Critical Path analysis
* Longest Path analysis
* Total Float calculation
* Free Float calculation
* Multi-project criticality mode
* Delay analysis
* Deadline monitoring
* Schedule momentum tracking

### Constraints & Controls

* Hard constraints engine
* Deadline management
* Constraint diagnostics
* Constraint impact analysis
* Dedicated constraints table

### Visualization

* Interactive Gantt chart
* Day / Week / Month timeline scaling
* Dependency link rendering
* Critical path overlays
* Progress overlays
* Constraint overlays
* Executive Dashboard
* S-Curve reporting

### Simulation

* Test mode planning
* Scenario mode planning
* Schedule comparison workflows
* Non-destructive schedule simulation

### Diagnostics

* Runtime warning console
* INFO / WARNING / STOP severities
* Event History
* Alarm History
* Warning acknowledgement workflow
* Stable event signature tracking
* Input validation framework

### Maintenance & Safety

* Safe Empty State for Gantt
* Safe Empty State for S-Curve
* Planning Reset workflow
* Full Reset workflow
* Protected calculated columns
* Controlled system writes

---

## Architecture

ProjectEngine is organized around a calculation core and multiple visualization layers.

| Layer         | Purpose                            |
| ------------- | ---------------------------------- |
| WBS           | User planning inputs and outputs   |
| CALC          | Scheduling engine runtime layer    |
| CONSTRAINTS   | Constraints and deadlines          |
| GANTT         | Schedule visualization             |
| SCURVE        | Progress and workload analytics    |
| DASHBOARD     | Executive reporting                |
| EVENT_HISTORY | Runtime traceability               |
| CALC_ALARM    | Active warnings and stops          |
| EVENT_ACK     | Warning acknowledgement management |

The scheduling engine is the single source of truth.

Visualization layers consume calculated data but never drive schedule calculations.

---

## Typical Workflow

1. Enter planning data in WBS.
2. Run Planning Update.
3. Review warnings and diagnostics.
4. Analyze schedule health through Dashboard and S-Curve.
5. Review Critical Path and Float analytics.
6. Simulate alternatives using Test and Scenario modes.
7. Generate reporting through Gantt and Dashboard views.

---

## Project Status

### Current Stage

**Beta**

The engine is actively used and continuously improved through real-world planning use cases.

Major scheduling, analytics, dashboard, diagnostics, and reset frameworks are operational.

Future releases will continue improving usability, visualization, and advanced planning capabilities.

---

## Support the Project

If ProjectEngine helps your work and you would like to support ongoing development:

![QRCode](Screenshots/qrcode.png)

---

## License

GPL-3.0
