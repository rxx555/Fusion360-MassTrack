# Using MassTrack

1. **Save your design first** — the workbook is named after the document.
2. Select the **bodies** you want counted, then **MassTrack → Mark**. (The picker defaults to bodies; switch to components if you prefer.)
3. For parts whose CAD mass is a placeholder, use **Set Mass** to enter the real mass in grams.
4. **MassTrack → Open in Excel** — writes the workbook and opens it.

Marks are stored inside the Fusion file, so they travel when the file is shared (F3D/F3Z or Hub — not STEP).

## Commands

- **Mark / Unmark** — add or remove a part from the budget.
- **Set Mass / Clear Mass** — override a part's mass with a known value in grams.
- **Snapshot** — record a labelled point in the mass history.
- **Show / Highlight** — list or highlight everything currently marked.
- **Open in Excel** — regenerate and open the workbook.
- **Diagram** — capture the current 3D view into a Diagram sheet in the workbook.

## Sheets

- **Overview** — total mass, breakdown by subassembly and material, heaviest parts, mass history chart.
- **Parts** — every marked part, linked to its source file.
- **Simple** — key, name, mass.
- **History / Trend** — totals and per-part mass over snapshots.
- **Work** — your own scratch sheet; the tool never overwrites it.
