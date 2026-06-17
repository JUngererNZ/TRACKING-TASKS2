
This script is a tracking update automation tool that synchronizes shipping/cargo status data from two external sources into a BARTRAC tracking spreadsheet. Here's what it does:

Overall Purpose
Updates the BARTRAC tracking spreadsheet with the latest truck/cargo statuses from FML and ORIENTO tracking reports, automatically matching vehicles and highlighting changes in pink.

Key Components
Setup & Utilities:

create_backup() - Creates a timestamped backup of BARTRAC before making changes
extract_component_keyword() - Cleans component names by removing "CAT 6060" prefix to help match records
FML File Processing (process_fml_file()):

Reads the FML Excel file looking for "Load Desc." row (component names) and "Truck" row (vehicle registrations)
Finds the last date row (latest status data)
Maps each component to its truck registration and latest status
Searches BARTRAC for matching rows by truck registration and cargo details
Updates the "ACTUAL STATUS" column with FML statuses and highlights them pink
ORIENTO File Processing (process_oriento_file()):

Reads the ORIENTO tracking report (sheet "BA3189")
Extracts truck numbers and their current statuses from the report
Matches trucks in BARTRAC by vehicle registration
Updates statuses and highlights them pink
Logs any unmatched vehicles
Main Workflow:

Backs up the original BARTRAC file
Opens BARTRAC workbook and locates the header row
Maps required columns (CARGO DETAILS, HORSE REG NO, ACTUAL STATUS)
Processes FML data first
Processes ORIENTO data second
Saves all changes to the original BARTRAC file
Default Files
The script uses these files if not provided as command-line arguments:

FML: vendor-report/FML CAT 6060.xlsx
BARTRAC: BARTRAC - KCC TRACKING AS OF 15-06-2026.xlsx
ORIENTO: ORIENTO TRACKING REPORT TO FML (DURBAN PORT TO KCC_.xlsx
The script essentially keeps your BART

This is an **advanced batch update script** that automates status synchronization for multiple BARTRAC tracking files from multiple vendor sources. It's essentially a more comprehensive version of the previous script.

## High-Level Purpose
Processes all BARTRAC files in a folder, updating each one with the latest tracking data from FML, ORIENTO, NATRANS, and Vanito tracking reports, plus manual overrides from a JSON file.

## Key Sections

**Configuration:**
- Defines paths to BARTRAC folder, vendor report folder, and specific tracking files
- Lists 5 ORIENTO files, 2 NATRANS files, 1 Vanito file, and a JSON override file
- All paths are hardcoded at the top for easy modification

**Helper Functions:**
- `create_backup()` - Same backup functionality as before
- `get_sheet_name_for_bartrac()` - Intelligently detects the correct sheet based on filename keywords (Congo, ERG, Kamoa, KCC, MUMI)
- `find_oriento_sheet()` - Searches for the correct sheet in ORIENTO files by looking for "TRUCK NUMBER"
- `extract_component_keyword()` - Cleans component names

**Data Processing Functions** (handles 4 different tracking sources):
- `process_fml_file()` - Same FML processing
- `process_oriento_file()` - Enhanced: auto-detects sheets, combines location + status, has special hardcoded handling for BA2951/EG20 files
- `process_natrans_file()` - New: reads NATRANS tracking files, extracts milestones (LOAD, DISP, BB SA, etc.), determines last status reached
- `process_vanito_file()` - New: reads Vanito tracking reports, auto-selects the latest dated sheet, extracts current locations

**Manual Overrides:**
- `load_horse_overrides()` - Reads horse-registration.json containing vehicle codes with manual status/color overrides
- `apply_horse_overrides()` - Applies these overrides to matching rows in BARTRAC, allows setting custom colors (like pink)

**Main Processing:**
- `update_bartrac_file()` - Processes a single BARTRAC file:
  1. Determines correct sheet name
  2. Creates backup
  3. Loads workbook and finds header
  4. Processes all vendor files in order (FML → ORIENTO → NATRANS → Vanito)
  5. Applies manual overrides from JSON
  6. Saves changes

- `main()` - Entry point:
  1. Validates all required files exist
  2. Finds all `.xlsx` files in BARTRAC folder (excluding temp files `~$*`)
  3. Processes each BARTRAC file sequentially
  4. Handles errors gracefully for each file

## Key Differences from the Simple Version
- **Batch mode**: Processes all BARTRAC files automatically instead of one
- **Multiple sources**: Handles 4 vendor sources vs. 2
- **Smart sheet detection**: Infers the correct sheet dynamically
- **Manual overrides**: Supports JSON-based custom status updates
- **Better error handling**: Skips missing files and continues with others
- **Milestone tracking**: NATRANS processor tracks journey milestones (locations reached)

This script is designed for daily/regular updates across multiple shipment tracking files from different logistics providers.

Looking at the log alongside the script, here are important operational insights:

## Data Matching Challenges

**The matching logic is truck-registration-based:**
- The script successfully updates rows only when the `HORSE REG NO` in BARTRAC exactly matches the truck registration from vendor files
- The log shows **many unmatched trucks** (`❌ No row found for truck 'XYZ'`) — suggesting either:
  - Different BARTRAC files have different trucks (each tracks its own shipment)
  - Vendor files contain trucks not in that specific BARTRAC file
  - Registration format mismatches (spaces, case sensitivity)

**Example success:** KCC BARTRAC updated 10 FML rows and 10 NATRANS rows, while KAMOA and CONGO files got 0 updates from the same sources.

## File/Path Issues

**Missing file warning:**
```
⚠️ ORIENTO file not found: C:\Users\Jason\...\BA2951...
```
One of the 5 configured ORIENTO files doesn't exist, but the script gracefully skips it and continues.

**Temporary file bug:**
The script tries to process `~$BARTRAC - CONGO TRACKING...` (Excel temp file) and crashes with permission denied. The filter logic `if not f.name.startswith("~$")` should prevent this but may not be catching all temp files.

## Data Staleness

**Vanito sheet is outdated:**
- Current date: 2026-06-16 (when script ran)
- Latest sheet found: `28 MAY 2026` (19 days old)
- The script auto-selects the latest dated sheet, but if no new sheet has been added since May 28, it keeps using old data

## JSON Override Failure

```
❌ Invalid JSON in horse-registration.json: line 13 column 3
```
The JSON file has a syntax error (missing quotes around a property name). Manual overrides never apply because the file fails to parse.

## Processing Order Matters

The script processes vendor files in this order: **FML → ORIENTO → NATRANS → Vanito**. If a truck gets a status from FML, NATRANS can overwrite it with different data from a later source if they both match the same truck.

## Summary

The script is **robust for multi-file batch processing** but has practical limitations:
- Depends heavily on truck registration matching (can fail silently if registrations don't align)
- File path configuration is hardcoded (any file rename breaks it)
- Gracefully handles missing files but not malformed data
- Catches temp files inconsistently
- Relies on vendor files being regularly updated (Vanito is stale in this run)

Plan — what I'll describe and provide:
1. Read vendor and bartrac listings from the two JSON files instead of hardcoding.
2. Auto-discover roles (FML, ORIENTO, NATRANS, Vanito) by filename patterns.
3. Fallback to existing behavior if JSONs are missing.
4. Provide a ready-to-paste code patch for update_all_bartrac.py.

What to change
- Add two new config paths pointing to the JSON manifests:
  - bartrac-directory_structure.json
  - vendor-directory_structure.json
- Replace the hardcoded vendor lists / folder glob in `main()` with logic that:
  - Loads vendor file paths from the vendor JSON and categorises them using regex/patterns.
  - Loads the BARTRAC files list from the bartrac JSON.
  - Falls back to the original folder-glob behavior if JSON files are missing or invalid.

Patch (insert or replace the relevant sections in update_all_bartrac.py)
- Replace the top CONFIG block and the start of `main()` with the following code (paste into the file, adapt minor variable names if needed):

```python
import json
import re
from pathlib import Path
# ... keep your other imports

# New config: point to the JSON manifests
BARTRAC_JSON = r"C:\Users\Jason\Projects\TRACKING-TASKS2\bartrac-directory_structure.json"
VENDOR_JSON = r"C:\Users\Jason\Projects\TRACKING-TASKS2\vendor-directory_structure.json"

def load_paths_from_manifest(json_path):
    p = Path(json_path)
    if not p.exists():
        return []
    try:
        with p.open('r', encoding='utf-8') as f:
            data = json.load(f)
    except Exception:
        return []
    children = data.get("children", [])
    paths = []
    for c in children:
        path = c.get("path") or c.get("filePath") or c.get("name")
        if path:
            paths.append(Path(path))
    return paths

def choose_fml(vendor_paths):
    # prefer CAT 6060 FML if present, else any file with 'FML' in name
    candidates = [p for p in vendor_paths if re.search(r'FML.*6060', p.name, re.I)]
    if not candidates:
        candidates = [p for p in vendor_paths if 'FML' in p.name.upper()]
    return candidates[0] if candidates else None

def choose_orientos(vendor_paths):
    return [p for p in vendor_paths if 'ORIENTO' in p.name.upper()]

def choose_natrans(vendor_paths):
    # pick files likely to be NATRANS/FML milestone sheets
    return [p for p in vendor_paths if re.search(r'FML|NATRANS', p.name, re.I) and ('BA' in p.name.upper() or 'KOLWEZI' in p.name.upper() or 'SAKANIA' in p.name.upper())]

def choose_vanito(vendor_paths):
    for p in vendor_paths:
        if 'VANITO' in p.name.upper():
            return p
    return None

def main():
    # load manifests if available
    vendor_paths = load_paths_from_manifest(VENDOR_JSON)
    bartrac_manifest_paths = load_paths_from_manifest(BARTRAC_JSON)

    # fallback to folder scan if manifests not present
    if not vendor_paths:
        vendor_folder = Path(r"C:\Users\Jason\Projects\TRACKING-TASKS2\vendor-report")
        vendor_paths = list(vendor_folder.glob("*.xlsx"))

    if bartrac_manifest_paths:
        bartrac_files = [p for p in bartrac_manifest_paths if p.suffix.lower() == ".xlsx" and not p.name.startswith("~$")]
    else:
        bartrac_folder = Path(r"C:\Users\Jason\Projects\TRACKING-TASKS2\BARTRAC-TRACKING")
        bartrac_files = [f for f in bartrac_folder.glob("*.xlsx") if not f.name.startswith("~$")]

    if not bartrac_files:
        print(f"No .xlsx files found to process.")
        return

    # auto-select vendor files
    fml_path = choose_fml(vendor_paths)
    oriento_paths = choose_orientos(vendor_paths)
    natrans_paths = choose_natrans(vendor_paths)
    vanito_path = choose_vanito(vendor_paths)

    print(f"Discovered vendor files: FML={fml_path}, ORIENTO={len(oriento_paths)}, NATRANS={len(natrans_paths)}, VANITO={vanito_path}")
    for bartrac_file in bartrac_files:
        try:
            update_bartrac_file(bartrac_file, fml_path, oriento_paths, natrans_paths, vanito_path, HORSE_OVERRIDE_JSON)
        except Exception as e:
            print(f"❌ Failed to process {bartrac_file.name}: {e}")

    print("\n✅ All BARTRAC files processed.")
```

Notes and recommendations
- The manifest JSONs you already have (examples: bartrac-directory_structure.json and vendor-directory_structure.json) contain full paths — this makes discovery simple and reliable.
- The categorization above uses filename patterns (case-insensitive). If your vendor filenames change format, either:
  - Add a small JSON config mapping (file role → regex), or
  - Put a consistent keyword in those filenames (e.g., include "ORIENTO" or "VANITO").
- Keep the fallback glob behavior — safe if manifests fail.
- Add logging (already present) for discovered paths so you can verify daily runs.
- Fix the existing broken JSON (horse-registration.json) so overrides apply; the loader already reports JSON errors.

If you want, I can:
- Produce a full patch/diff for update_all_bartrac.py ready to apply.
- Add a small `config.json` (role→regex) to make discovery explicit and robust. Which would you prefer?

**Plan**
- Review your dynamic script and list any issues or improvements.
- Suggest concise fixes you can accept.
- Apply the patch to update_all_bartrac.py if you want me to.

**Quick Review & Suggestions**
- **Year regex**: Replace `^2026-\d{2}-\d{2}` with a generic `^\d{4}-\d{2}-\d{2}` or make it configurable to avoid hardcoding 2026.  
- **FML discovery**: Fall back to any `FML` file if `"FML CAT 6060"` not found; prefer most-recent modified file when multiples exist.  
- **ORIENTO matching**: Matching by `"CATERPILLAR"` may capture unrelated files — prefer `ORIENTO` or sheet patterns (or use regex like `BA\d{3,4}`).  
- **Vanito sheet parsing**: Trim sheet names and use locale-safe month parsing; skip sheets with trailing spaces.  
- **Column lookup robustness**: Normalize header names (strip, upper) when building `col_index` to tolerate small header variations.  
- **Double workbook load**: `update_bartrac_file` calls `load_workbook` twice; remove the extra load to save time.  
- **JSON overrides**: Script already tolerates invalid JSON — recommend surfacing the JSON error so overrides can be fixed (see horse-registration.json).  
- **Logging & dry-run**: Add a `--dry-run` flag and verbose logging to preview changes before saving.  
- **Fallback behavior**: Keep your current fallbacks but add explicit console warnings when a critical file (like FML) is missing.

Shall I apply these improvements and patch update_all_bartrac.py now?