# Maintenance Guide

## General rule

First identify the owner of the data or policy. Change the smallest correct boundary, then select validation proportional to actual consumers.

## Adding a WBS or CALC column

1. Decide whether the column is a WBS input, CALC working value or calculated output.
2. Update the schema owner in DataSync/Infrastructure, not every consumer independently.
3. Add Pre-Core validation when a missing or invalid value must block calculation.
4. Add the field to the Core dataset only when the engine actually consumes it.
5. Update Full and Partial Output Writers when the column is written as output.
6. Add the column to WBS Write Guard scopes when the engine writes it to WBS.
7. Change `mod_IncrementalSignature` only when field variation must invalidate calculation. Preserve order and version the contract.
8. Search all direct references to the column name and verify that each belongs to the correct owner.

Membership in the incremental signature is a business decision, not an automatic consequence of adding a column.

## Adding a diagnostic

1. Produce the diagnostic in the domain that detects the condition.
2. Reuse existing formatting services for text, without business recalculation.
3. Explicitly select `STOP`, `WARNING` or `INFO`.
4. Define a stable code and hash when EventHistory or ACK must recognize the event.
5. Route through MessageEngine; never call the UserForm directly.
6. Log through EventHistory's owning contract.
7. Add a case to the producer harness and verify both interactive and noninteractive projections.

The producer owns meaning and severity. MessageEngine owns preparation. EventHistory owns persistence and ACK. The UserForm owns display only.

## Adding a callback or button

1. Place the callback in a standard module or Excel object compatible with its mechanism.
2. Use a stable Public wrapper for `OnAction`, `Application.Run`, `AddressOf`, timers or Excel events.
3. Delegate immediately to the owning service; keep business logic out of the callback.
4. Add the name to the module header under `CALLBACKS EXTERNES / EXTERNAL CALLBACKS`.
5. Scan `OnAction`, `Application.Run`, `SetTimer` and `AddressOf` before renaming.
6. Test the callback on a temporary copy.

Never rename an external callback for style alone. Keep a wrapper when compatibility requires it.

## Adding a Task Type rule

1. Add normalization or classification to `mod_TaskTypeRules`.
2. Do not duplicate the rule in Core, Constraints, S-Curve or Gantt.
3. Let each consumer apply its own policy after classification: rendering exclusion, validation and calculation are different decisions.
4. Add LOE, Milestone and unknown-value cases to directly consuming harnesses.

## Changing the incremental signature

1. Change only `mod_IncrementalSignature` for fields, order and serialization.
2. Treat the format as a persistent contract.
3. Capture signatures before modification with `mod_IncrementalSignatureHarness`.
4. Any field change requires an explicit CALC_STATE versioning or invalidation strategy.
5. Verify `mod_CalcState` and `mod_CalcIncremental` without redefining the contract there.
6. Require a golden capture. Every bit-level difference must be intentional and explained.

## Changing Gantt code

| Change | Owner | Minimum validation |
|---|---|---|
| date-to-position or geometry | Geometry / TimelineGeometry | compile and targeted visual harness |
| task bar, milestone or summary | GanttRenderer | Visual Regression and TEST/fallback as consumed |
| dependencies | DependencyRenderer | Registry and Visual Regression |
| constraints or deadlines | ConstraintRenderer | Visual Regression and relevant scenario |
| registry, diff or Lazy Repair | ShapeRegistry | Predictive Registry and Visual Regression |
| buttons, toggles or language | UiControls / ViewState / Language | targeted UI signature |
| drag or timer | GanttDragWatch | Drag/TEST smoke and timer lifecycle |
| TEST, SCENARIO or LOCK | corresponding service | transactional smoke on a copy |

Never change Shape names, `OnAction`, z-order, tolerances or fallback during cosmetic cleanup. Never create a second renderer or simulation engine.

## Selecting validation level

### Level S: structure

Local rename, visibility, dead-code removal or comments: static audit, targeted import, compile and consumer harness only when required.

### Level M: module

Local extraction, new internal contract or projection change: compile, targeted golden capture and direct consuming workflows.

### Level A: architecture or behavior

New boundary, engine, writer, external callback or main workflow: broader validation including affected guards and transactions.

Do not run a global suite by habit. Do not omit an actual consuming harness merely to save time.

## CP1252 and UTF-8 encodings

- Read bytes and detect UTF-8 BOM, then valid UTF-8, then CP1252 as fallback.
- Write using the same encoding and line endings.
- Never implicitly convert a historical CP1252 VBA file to UTF-8.
- Prefer ASCII in VBA code and comments when the source already does.
- Keep Markdown and TSV documents in UTF-8.
- After a documentation-only transformation, compare a non-comment code hash.

## Importing and compiling VBA

Reference tools:

```powershell
powershell -ExecutionPolicy Bypass -File codex_tools\Import_VBA_To_Workbook.ps1
powershell -ExecutionPolicy Bypass -File codex_tools\tmp_compile_vba_project.ps1
```

Read script parameters and verify the target workbook before execution. Smokes must create their own copy under `%TEMP%`, preserve the source workbook hash and trace the Excel process they create.

Compile the complete VBAProject after import. Static scanning cannot replace VBA compilation: visibility, optional arguments and array types may fail only in VBE.

## Never closing user Excel

1. Capture existing Excel processes before starting a worker.
2. Create a dedicated COM instance and retain its PID or handle.
3. Open only an identified temporary copy.
4. Close the copy, call `Quit` on the owned instance and release its COM objects.
5. On timeout, stop only the PID created by the worker.
6. Never perform global process-name or partial-workbook-name termination.
7. Every user Excel instance remains out of scope, even when it opens another workbook.

## Choosing Public, Friend or Private

| Visibility | Use when |
|---|---|
| `Private` | helper consumed by one module or internal class state |
| `Friend` | contract required inside the VBAProject but not intended for Excel or external macros |
| `Public` | user macro, external callback, worker-invoked harness or demonstrated inter-module contract |

When a dynamic call is uncertain, keep Public and document the contract. Do not make a helper Public merely because it was moved.

## When to create a module

Create a module when a complete responsibility has a contract, an invariant and several coherent procedures, or when it becomes the sole owner of a genuinely duplicated canonical contract.

Do not create a module for:

- a few orchestration lines;
- one isolated helper with no policy;
- avoiding two parameters;
- hiding a dependency that should remain visible;
- grouping procedures only because their names look similar.

## When not to create a Wrapper, Context or Read Model

- No wrapper for an atomic local rename without a dynamic contract.
- No Context when its members do not share the same lifecycle.
- No universal DTO mixing WBS, CALC, UI, diagnostics and options.
- No Read Model when consumers share only a source table, not the same projection.
- No policy booleans used to force different workflows into one abstraction.

## Architecture mistakes to avoid

- pass-through facade with no stability value;
- universal DTO;
- business policy hidden in `Utils`, `Helper` or `Manager`;
- cell-by-cell reads when a table array is sufficient;
- duplication of Identity Index, Parsed Network or Incremental Signature;
- direct access to another owner's store;
- global reset aware of every schema;
- closing an unowned WBS scope;
- TEST -> SCENARIO or LOCK -> TEST Service dependency;
- validation disproportionate to risk;
- comments that merely repeat procedure names;
- turning a timeout into PASS or hiding fallback in a harness.

## Delivery checklist

1. Owner and boundary identified.
2. External contracts and callbacks scanned.
3. Business body unchanged or change explicitly authorized.
4. Module and procedure documentation updated.
5. Encoding preserved and non-comment hash checked for documentation-only changes.
6. Import and compilation proportional to risk.
7. Direct consumer harnesses green.
8. Source workbook protected.
9. No user Excel instance closed.
10. Report proportional and coverage limitations explicit.
