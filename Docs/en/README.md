# English documentation

[Documentation française](../fr/README.md)

## Start here

- [Architecture Reading Guide](ARCHITECTURE_READING_GUIDE.md): understand domains, workflows, owners and safeguards.
- [Maintenance Guide](MAINTENANCE_GUIDE.md): change the project without breaking its contracts.
- [Project Glossary](PROJECT_GLOSSARY.md): official vocabulary used in code and documentation.

## Using ProjectEngine v1.3.0

- **First use:** the welcome screen offers **Start New Project** or **Import Previous Project** when the workbook is eligible (empty and writable). The welcome choice is stored through the existing acknowledgement system; Full Reset can bring the welcome screen back.
- **Excel Ribbon:** the embedded **ProjectEngine** tab groups planning updates, Gantt simulation controls, Settings, navigation, reset and import. Controls change with the active worksheet; no separate add-in is required.
- **Importing older projects:** select **Import Previous Project** in the welcome screen or Ribbon. The destination's supported project inputs are replaced, the source workbook remains unchanged, and the user must review WBS/Constraints and explicitly run Planning Update or Full Update afterward. **Back up the destination workbook yourself before importing.**
- **Gantt navigation:** **Go to Selected Task** and **Show Full Timeline** use the existing Gantt owner. Complete vertical/horizontal fit is limited by Excel's minimum zoom on very large schedules.

The [Architecture Reading Guide](ARCHITECTURE_READING_GUIDE.md) explains ownership and processing flows; the [Maintenance Guide](MAINTENANCE_GUIDE.md) covers changes and safety constraints.

## Recommended path

1. Read the Architecture Reading Guide for a 15-minute overview.
2. Check the glossary before creating or renaming a component.
3. Use the Maintenance Guide to select the correct boundary and validation protocol.
4. Read FR/EN module and procedure headers for detailed contracts in the current code.
