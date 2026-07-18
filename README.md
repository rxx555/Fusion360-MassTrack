# MassTrack

A Fusion 360 add-in that tracks the mass of the parts you mark and exports them to an Excel workbook — per part, with history and charts.

Mark the bodies (or components) you care about; MassTrack reads their physical properties and writes a styled workbook: total mass, breakdown by subassembly and by material, heaviest parts, and mass history. Parts are opt-in — nothing is exported until you mark it, so tool bodies and cutters stay out of the budget.

For parts whose CAD mass is a placeholder (electronics, bought-in parts), set a known mass in grams and it overrides the CAD value.

## Install

1. Download `dist/MassTrack_addin.zip` and unzip it — you get a `MassTrack` folder.
2. Move that folder into Fusion's add-ins directory:
   - **macOS:** `~/Library/Application Support/Autodesk/Autodesk Fusion 360/API/AddIns/`
   - **Windows:** `%appdata%\Autodesk\Autodesk Fusion 360\API\AddIns\`
3. In Fusion: **Utilities → ADD-INS → MassTrack → Run** (tick **Run on Startup**).

A **MassTrack** panel appears in the Design ribbon.

## Use

See [USAGE.md](USAGE.md).

## Licence

GPL-3.0-or-later. See [LICENSE](LICENSE).
