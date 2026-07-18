# Using MassTrack

1. **Save your design first.** The workbook is named after the document.
2. Select the **bodies** you want counted, then **MassTrack → Mark**. (The picker defaults to bodies; switch to components if you prefer.)
3. For parts whose CAD mass is a placeholder, use **Set Mass** to enter the real mass in grams.
4. **MassTrack → Open in Excel** writes the workbook and opens it.

Marks are stored inside the Fusion file, so they travel when the file is shared (F3D/F3Z or Hub, not STEP).

## Commands

- **Mark / Unmark** adds or removes a part from the budget.
- **Set Mass / Clear Mass** overrides a part's mass with a known value in grams.
- **Snapshot** records a labelled point in the mass history.
- **Show / Highlight** lists or highlights everything currently marked.
- **Open in Excel** regenerates and opens the workbook.
- **Diagram** captures the current 3D view into a Diagram sheet in the workbook.

## Sheets

- **Overview** shows the total mass, breakdowns by subassembly and by material, the heaviest parts, and a mass history chart.
- **Parts** lists every marked part, linked to its source file.
- **Simple** has just the key, name, and mass.
- **History / Trend** track totals and per-part mass over snapshots.
- **Work** is your own scratch sheet. The tool never overwrites it.
