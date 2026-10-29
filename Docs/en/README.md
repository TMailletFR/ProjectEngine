# English documentation

[Documentation française](../fr/README.md)

## Start here

- [Architecture Reading Guide](ARCHITECTURE_READING_GUIDE.md): understand domains, workflows, owners and safeguards.
- [Maintenance Guide](MAINTENANCE_GUIDE.md): change the project without breaking its contracts.
- [Project Glossary](PROJECT_GLOSSARY.md): official vocabulary used in code and documentation.

## Using ProjectEngine v1.5.0

- **First use:** the welcome screen offers **Start New Project**, **Import Previous Project**, or **Discover ProjectEngine** when the workbook is eligible (empty and writable). The welcome choice is stored through the existing acknowledgement system; Full Reset can bring the welcome screen back.
- **Excel Ribbon:** the embedded **ProjectEngine** tab groups planning updates, Gantt simulation controls, Settings, navigation, reset and import. Controls change with the active worksheet; no separate add-in is required.
- **Importing older projects:** select **Import Previous Project** in the welcome screen or Ribbon. The destination's supported project inputs are replaced, the source workbook remains unchanged, and the user must review WBS/Constraints and explicitly run Planning Update or Full Update afterward. **Back up the destination workbook yourself before importing.**
- **Interactive tutorial:** open **Tutorial** from the Ribbon or **Discover ProjectEngine** from Welcome. Choose **Discovery** for a passive tour or **Practice** to edit actual WBS inputs and execute normal commands with explicit consent. Both modes have independent bookmarks across 29 steps. The tutorial does not inject sample tasks or recalculate on passive Next.
- **Modern dialogs:** Welcome, Planning Messages and Tutorial use updated layouts and controls. Message windows are sized once before showing and retain selectable, scrollable text. Conventional native controls remain available as a fallback when required.
- **Gantt navigation:** **Go to Selected Task** and **Show Full Timeline** use the existing Gantt owner. Complete vertical/horizontal fit is limited by Excel's minimum zoom on very large schedules.

The [Architecture Reading Guide](ARCHITECTURE_READING_GUIDE.md) explains ownership and processing flows; the [Maintenance Guide](MAINTENANCE_GUIDE.md) covers changes and safety constraints.

## Finance and multi-project controls (new in v1.5.0)

- **FINANCE:** attach WBS activities, enter Budget and ETC, review Actual, EAC and Variance, and select **START**, **END** or **LINEAR** planned-cost profiles.
- **Expense journal:** record dated expenses and negative credits; prepare multiple edits using a native Excel draft table with filtering, sorting and pasting, then explicitly confirm or cancel.
- **Multi-project:** maintain independent financial projects and independent, copied or shared financial records. The **Global** view consolidates shared financial data once.
- **FINANCE SUMMARY:** inspect monthly and cumulative planned/actual spending. Future expenses remain recorded but are excluded from actual chart values until the Finance refresh reference date reaches them.
- **Explicit calculation:** select **Refresh Finance**; opening the workbook, changing language or running Planning Update does not automatically refresh Finance.
- **Migration:** **SchemaVersion 3** preserves supported financial-domain data, editable WBS cell formats and selected column widths across compatible workbook versions. The source remains unchanged; back up the destination first.
- **Full Reset:** clears the previous project's Finance data, extracted sheets, event history and old caches while retaining the workbook infrastructure and supported preferences.

Finance provides **project cost reporting**, not general-ledger accounting or invoicing. START/END/LINEAR apply to Finance cost timing, not progress S-Curve shapes. Global Finance does not combine separate planning workbooks.


## Recommended path

1. Read the Architecture Reading Guide for a 15-minute overview.
2. Check the glossary before creating or renaming a component.
3. Use the Maintenance Guide to select the correct boundary and validation protocol.
4. Read FR/EN module and procedure headers for detailed contracts in the current code.
