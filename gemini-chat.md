The short answer is **yes, absolutely**—and doing this will make your workflow vastly more resilient.

Looking closely at your script, the logic is sound, but it breaks down daily because it relies on explicit, hardcoded filenames. Meanwhile, your vendor files change dynamically as new jobs, dates, and reference numbers arrive.

Instead of parsing the raw `.json` directory snapshots directly (which would require you to generate those JSONs via an external process every day), we can make your Python script completely **self-sufficient**. By utilizing Python's native `pathlib` module, the script can look directly inside `vendor-report` and `BARTRAC-TRACKING`, matching files dynamically using keywords, pattern matching, and file metadata.

Here is the updated, dynamic version of your script.

### Updated Python Script (`update_all_bartrac.py`)

```python
import pandas as pd
import re
import sys
import shutil
import json
from pathlib import Path
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

PINK_FILL = PatternFill(start_color="FFC0CB", end_color="FFC0CB", fill_type="solid")

# ------------------- CONFIGURATION -------------------
BARTRAC_FOLDER = Path(r"C:\Users\Jason\Projects\TRACKING-TASKS2\BARTRAC-TRACKING")
VENDOR_FOLDER = Path(r"C:\Users\Jason\Projects\TRACKING-TASKS2\vendor-report")
HORSE_OVERRIDE_JSON = Path(r"C:\Users\Jason\Projects\TRACKING-TASKS2\horse-registration.json")

# ------------------- DYNAMIC FILE DISCOVERY -------------------
def discover_vendor_files(vendor_dir: Path):
    """
    Dynamically scans the vendor folder and groups files by provider based on naming patterns.
    """
    fml_file = None
    oriento_files = []
    natrans_files = []
    vanito_file = None

    if not vendor_dir.exists():
        print(f"❌ ERROR: Vendor directory does not exist: {vendor_dir}")
        sys.exit(1)

    # Scan all excel files in the folder (ignoring temp files)
    all_files = [f for f in vendor_dir.glob("*.xlsx") if not f.name.startswith("~$")]

    for file_path in all_files:
        name_upper = file_path.name.upper()
        
        # 1. Match FML Master file
        if "FML CAT 6060" in name_upper:
            fml_file = file_path
            
        # 2. Match Vanito
        elif "VANITO" in name_upper:
            vanito_file = file_path
            
        # 3. Match NATRANS
        elif "NATRANS" in name_upper:
            natrans_files.append(file_path)
            
        # 4. Match ORIENTO (Explicit keyword or standard pattern from your tracking files)
        elif "ORIENTO" in name_upper or "CATERPILLAR" in name_upper:
            oriento_files.append(file_path)

    # Fallback: If FML master isn't explicitly named 'FML CAT 6060', look for any NATRANS files 
    # and map others accordingly, or fall back to the first available if strict matching isn't required.
    print("\n🔍 Dynamic Discovery Summary:")
    print(f"  🔹 FML Master File: {fml_file.name if fml_file else '❌ NOT FOUND (Make sure it is named correctly or uncommented)'}")
    print(f"  🔹 Vanito Report:   {vanito_file.name if vanito_file else '⚠️ NOT FOUND'}")
    print(f"  🔹 NATRANS Files:   {[f.name for f in natrans_files]}")
    print(f"  🔹 ORIENTO Files:   {[f.name for f in oriento_files]}")
    print("-" * 60)

    return fml_file, oriento_files, natrans_files, vanito_file

# ------------------- CORE UTILITIES -------------------
def create_backup(file_path):
    file_path = Path(file_path)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_name = f"{file_path.stem}_backup_{timestamp}{file_path.suffix}"
    backup_path = file_path.parent / backup_name
    shutil.copy2(file_path, backup_path)
    print(f"📁 Backup created: {backup_path}")
    return backup_path

def extract_component_keyword(comp_name):
    cleaned = comp_name.replace("CAT 6060", "").strip()
    return cleaned

def get_sheet_name_for_bartrac(bartrac_path):
    filename = bartrac_path.name.lower()
    if "congo" in filename:
        return "CURRENT SHIPMENTS"
    elif "erg" in filename:
        return "CURRENT SHIPMENTS"
    elif "kamoa" in filename:
        return "ENROUTE SITE"
    elif "kcc" in filename:
        return "ENROUTE SITE"
    elif "mumi" in filename:
        return "ENROUTE SITE"
    else:
        return None

def find_oriento_sheet(file_path):
    try:
        xl = pd.ExcelFile(file_path)
        for sheet in xl.sheet_names:
            df_sample = pd.read_excel(file_path, sheet_name=sheet, header=None, nrows=20)
            for _, row in df_sample.iterrows():
                for cell in row.iloc[:20]:
                    if pd.notna(cell) and "TRUCK NUMBER" in str(cell).upper():
                        return sheet
        return None
    except Exception as e:
        print(f"⚠️ Could not read sheets from {file_path}: {e}")
        return None

# ------------------- PROCESSING ENGINES -------------------
def process_oriento_file(oriento_path, wb, ws, col_index):
    oriento_path = Path(oriento_path)
    if not oriento_path.exists():
        print(f"⚠️ ORIENTO file not found: {oriento_path}")
        return

    if "BA2951" in oriento_path.name or "2603DSI2788" in oriento_path.name:
        sheet_name = "2603DSI2788"
        print(f"Using hardcoded sheet '{sheet_name}' for BA2951/EG20 file: {oriento_path.name}")
    else:
        sheet_name = find_oriento_sheet(oriento_path)
        if sheet_name is None:
            print(f"⚠️ Could not find sheet with 'TRUCK NUMBER' in {oriento_path.name}. Skipping.")
            return

    print(f"Reading ORIENTO file: {oriento_path} (sheet: {sheet_name})")
    try:
        df_raw = pd.read_excel(oriento_path, sheet_name=sheet_name, header=None, dtype=str)
    except Exception as e:
        print(f"⚠️ Could not read sheet '{sheet_name}': {e}")
        return

    header_idx = None
    for idx, row in df_raw.iterrows():
        for cell in row.iloc[:20]:
            if pd.notna(cell) and "TRUCK NUMBER" in str(cell).upper():
                header_idx = idx
                break
        if header_idx is not None:
            break
    if header_idx is None:
        print(f"❌ Could not find header row in ORIENTO sheet {sheet_name}")
        return

    header_row = df_raw.iloc[header_idx]
    truck_col, location_col, status_col = None, None, None
    for i, val in enumerate(header_row):
        if pd.notna(val):
            val_str = str(val).strip().upper()
            if "TRUCK NUMBER" in val_str:
                truck_col = i
            elif "CURRENT LOCATION" in val_str:
                location_col = i
            elif "CURRENT STATUS" in val_str:
                status_col = i

    if truck_col is None:
        print(f"❌ Could not find 'TRUCK NUMBER' column in ORIENTO sheet {sheet_name}")
        return

    data_rows = df_raw.iloc[header_idx + 1:].copy()
    data_rows.columns = range(data_rows.shape[1])

    truck_to_status = {}
    for _, row in data_rows.iterrows():
        truck = str(row.iloc[truck_col]).strip() if pd.notna(row.iloc[truck_col]) else ""
        if not truck or truck == 'nan':
            continue

        location = str(row.iloc[location_col]).strip() if location_col is not None and pd.notna(row.iloc[location_col]) else ""
        status = str(row.iloc[status_col]).strip() if status_col is not None and pd.notna(row.iloc[status_col]) else ""

        if location and status:
            combined = f"{location} - {status}"
        elif location:
            combined = location
        elif status:
            combined = status
        else:
            continue

        truck_to_status[truck] = combined

    header_row_idx = None
    for row in ws.iter_rows(min_row=1, max_row=100, values_only=False):
        for cell in row:
            if cell.value and "CLIENT PO" in str(cell.value).upper():
                header_row_idx = cell.row
                break
        if header_row_idx is not None:
            break

    updated_count = 0
    for truck_reg, status in truck_to_status.items():
        matching_rows = []
        for row in ws.iter_rows(min_row=header_row_idx+1, values_only=False):
            horse_cell = row[col_index["HORSE REG NO"] - 1]
            if horse_cell.value and str(horse_cell.value).strip() == truck_reg.strip():
                matching_rows.append(row)

        for row in matching_rows:
            status_cell = row[col_index["ACTUAL STATUS"] - 1]
            old_value = status_cell.value
            status_cell.value = status.upper()
            status_cell.fill = PINK_FILL
            updated_count += 1
            print(f"✅ ORIENTO ({sheet_name}): Updated row {row[0].row} (truck {truck_reg}) → '{status.upper()}' (was: '{old_value}')")

    print(f"📊 ORIENTO ({sheet_name}) total rows updated: {updated_count}")

def process_fml_file(fml_path, wb, ws, col_index):
    if not fml_path:
        print("⚠️ Skipping FML Process: No dynamic file discovered.")
        return
    print(f"Reading FML file: {fml_path}")
    df_fml_raw = pd.read_excel(fml_path, sheet_name="Sheet1", header=None, dtype=str)

    load_desc_idx = None
    for idx, row in df_fml_raw.iterrows():
        first_cell = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ""
        if first_cell == "Load Desc.":
            load_desc_idx = idx
            break
    if load_desc_idx is None:
        raise ValueError("Could not find 'Load Desc.' row in FML file")
    components = [str(c).strip() for c in df_fml_raw.iloc[load_desc_idx, 1:] if pd.notna(c)]

    truck_idx = None
    for idx, row in df_fml_raw.iterrows():
        first_cell = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ""
        if first_cell == "Truck":
            truck_idx = idx
            break
    if truck_idx is None:
        raise ValueError("Could not find 'Truck' row in FML file")
    trucks = [str(t).strip() if pd.notna(t) else "" for t in df_fml_raw.iloc[truck_idx, 1:]]

    comp_to_truck = {comp: truck for comp, truck in zip(components, trucks) if truck}

    date_pattern = re.compile(r'^2026-\d{2}-\d{2}')
    data_rows = []
    for idx, row in df_fml_raw.iterrows():
        first_cell = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ""
        if date_pattern.match(first_cell):
            data_rows.append(idx)
    if not data_rows:
        raise ValueError("No date rows found in FML file")
    last_data_idx = data_rows[-1]
    latest_statuses = [str(s).strip() if pd.notna(s) else "" for s in df_fml_raw.iloc[last_data_idx, 1:]]
    comp_to_status = {comp: status for comp, status in zip(components, latest_statuses) if comp}

    header_row_idx = None
    for row in ws.iter_rows(min_row=1, max_row=100, values_only=False):
        for cell in row:
            if cell.value and "CLIENT PO" in str(cell.value).upper():
                header_row_idx = cell.row
                break
        if header_row_idx is not None:
            break

    updated_count = 0
    for comp, truck_reg in comp_to_truck.items():
        keyword = extract_component_keyword(comp)
        status = comp_to_status.get(comp, "")
        if not status:
            continue

        matching_rows = []
        for row in ws.iter_rows(min_row=header_row_idx+1, values_only=False):
            horse_cell = row[col_index["HORSE REG NO"] - 1]
            if horse_cell.value and str(horse_cell.value).strip() == truck_reg.strip():
                matching_rows.append(row)

        for row in matching_rows:
            status_cell = row[col_index["ACTUAL STATUS"] - 1]
            old_value = status_cell.value
            status_cell.value = status.upper()
            status_cell.fill = PINK_FILL
            updated_count += 1
            print(f"✅ FML: Updated row {row[0].row} for {comp} (truck {truck_reg}) → '{status.upper()}' (was: '{old_value}')")

    print(f"📊 FML total rows updated: {updated_count}")

def process_natrans_file(natrans_path, wb, ws, col_index):
    print(f"Reading NATRANS file: {natrans_path.name}")
    try:
        df_raw = pd.read_excel(natrans_path, sheet_name="NATRANS", header=None, dtype=str)
    except ValueError:
        print(f"⚠️ Sheet 'NATRANS' not found in {natrans_path.name}. Skipping.")
        return

    header_idx = None
    for idx, row in df_raw.iterrows():
        for cell in row.iloc[:10]:
            if pd.notna(cell) and str(cell).strip().upper() in ["LOAD", "DISP", "BB SA", "BB ZIM"]:
                header_idx = idx
                break
        if header_idx is not None:
            break
    if header_idx is None:
        print(f"⚠️ Could not find header row in NATRANS sheet of {natrans_path.name}")
        return

    header_row = df_raw.iloc[header_idx]
    col_names = {str(val).strip(): i for i, val in enumerate(header_row) if pd.notna(val)}

    horse_col = None
    for possible in ["HORSE", "HORSE REG NO", "HORSE REG"]:
        if possible in col_names:
            horse_col = col_names[possible]
            break
    if horse_col is None:
        print("⚠️ Could not find HORSE column in NATRANS file")
        return

    milestone_columns = []
    for col in ["LOAD", "DISP", "BB SA", "BB ZIM", "CHI ZIM", "CHI ZAM", 
                "KASUM ZAM", "KASUM DRC", "WHISKEY", "SITE", "OFFLOAD", 
                "LUFUA ARR", "LUFUA DISP"]:
        if col in col_names:
            milestone_columns.append(col)

    data_rows = df_raw.iloc[header_idx + 1:].copy()
    data_rows.columns = range(data_rows.shape[1])

    truck_to_status = {}
    for _, row in data_rows.iterrows():
        truck = str(row.iloc[horse_col]).strip() if pd.notna(row.iloc[horse_col]) else ""
        if not truck or truck == 'nan':
            continue

        last_status = None
        for milestone in milestone_columns:
            col_idx = col_names[milestone]
            cell_value = row.iloc[col_idx]
            if pd.notna(cell_value) and str(cell_value).strip() not in ['', 'nan']:
                last_status = milestone

        if "COMMENTS" in col_names:
            comments_idx = col_names["COMMENTS"]
            comments_val = row.iloc[comments_idx]
            if pd.notna(comments_val) and str(comments_val).strip() not in ['', 'nan']:
                truck_to_status[truck] = str(comments_val).strip()
            elif last_status:
                truck_to_status[truck] = last_status
        else:
            if last_status:
                truck_to_status[truck] = last_status

    header_row_idx = None
    for row in ws.iter_rows(min_row=1, max_row=100, values_only=False):
        for cell in row:
            if cell.value and "CLIENT PO" in str(cell.value).upper():
                header_row_idx = cell.row
                break
        if header_row_idx is not None:
            break

    updated_count = 0
    for truck_reg, status in truck_to_status.items():
        matching_rows = []
        for row in ws.iter_rows(min_row=header_row_idx+1, values_only=False):
            horse_cell = row[col_index["HORSE REG NO"] - 1]
            if horse_cell.value and str(horse_cell.value).strip() == truck_reg.strip():
                matching_rows.append(row)

        for row in matching_rows:
            status_cell = row[col_index["ACTUAL STATUS"] - 1]
            old_value = status_cell.value
            status_cell.value = status.upper()
            status_cell.fill = PINK_FILL
            updated_count += 1
            print(f"✅ NATRANS ({natrans_path.name}): Updated row {row[0].row} (truck {truck_reg}) → '{status.upper()}' (was: '{old_value}')")

    print(f"📊 NATRANS ({natrans_path.name}) total rows updated: {updated_count}")

def process_vanito_file(vanito_path, wb, ws, col_index):
    if not vanito_path:
        print("⚠️ Skipping Vanito Process: No file discovered.")
        return
    print(f"Reading Vanito file: {vanito_path}")

    try:
        xl = pd.ExcelFile(vanito_path)
        sheets = xl.sheet_names
    except Exception as e:
        print(f"⚠️ Could not read sheets from Vanito file: {e}")
        return

    date_pattern = re.compile(r'^(\d{1,2})\s+([A-Z]{3,})\s+(\d{4})$')
    sheet_date_map = {}
    for sheet in sheets:
        m = date_pattern.match(sheet.strip())
        if m:
            day, month_str, year = m.groups()
            try:
                month_num = datetime.strptime(month_str, "%b").month
                dt = datetime(int(year), month_num, int(day))
                sheet_date_map[dt] = sheet
            except ValueError:
                continue

    if not sheet_date_map:
        print("⚠️ No sheets with date format found in Vanito file.")
        return

    latest_date = max(sheet_date_map.keys())
    latest_sheet = sheet_date_map[latest_date]
    print(f"Using latest sheet: '{latest_sheet}'")

    try:
        df_raw = pd.read_excel(vanito_path, sheet_name=latest_sheet, header=None, dtype=str)
    except Exception as e:
        print(f"⚠️ Could not read sheet '{latest_sheet}': {e}")
        return

    header_idx = None
    for idx, row in df_raw.iterrows():
        for cell in row.iloc[:20]:
            if pd.notna(cell) and "HORSE REG" in str(cell).upper():
                header_idx = idx
                break
        if header_idx is not None:
            break
    if header_idx is None:
        print(f"⚠️ Could not find header row with 'HORSE REG' in sheet '{latest_sheet}'")
        return

    header_row = df_raw.iloc[header_idx]
    horse_col, location_col = None, None
    for i, val in enumerate(header_row):
        val_str = str(val).strip().upper()
        if "HORSE REG" in val_str:
            horse_col = i
        elif "CURRENT LOCATION" in val_str or "CURRENT STATUS" in val_str or "LOCATION" in val_str:
            location_col = i

    if horse_col is None or location_col is None:
        return

    data_rows = df_raw.iloc[header_idx + 1:].copy()
    data_rows.columns = range(data_rows.shape[1])

    truck_to_status = {}
    for _, row in data_rows.iterrows():
        truck = str(row.iloc[horse_col]).strip() if pd.notna(row.iloc[horse_col]) else ""
        location = str(row.iloc[location_col]).strip() if pd.notna(row.iloc[location_col]) else ""
        if truck and location and truck != 'nan' and location != 'nan':
            truck_to_status[truck] = location

    bartrac_header_idx = None
    for row in ws.iter_rows(min_row=1, max_row=100, values_only=False):
        for cell in row:
            if cell.value and "CLIENT PO" in str(cell.value).upper():
                bartrac_header_idx = cell.row
                break
        if bartrac_header_idx is not None:
            break

    updated_count = 0
    for truck_reg, status in truck_to_status.items():
        matching_rows = []
        for row in ws.iter_rows(min_row=bartrac_header_idx+1, values_only=False):
            horse_cell = row[col_index["HORSE REG NO"] - 1]
            if horse_cell.value and str(horse_cell.value).strip() == truck_reg.strip():
                matching_rows.append(row)

        for row in matching_rows:
            status_cell = row[col_index["ACTUAL STATUS"] - 1]
            old_value = status_cell.value
            status_cell.value = status.upper()
            status_cell.fill = PINK_FILL
            updated_count += 1
            print(f"✅ Vanito: Updated row {row[0].row} (truck {truck_reg}) → '{status.upper()}' (was: '{old_value}')")

    print(f"📊 Vanito total rows updated: {updated_count}")

def load_horse_overrides(json_path):
    overrides = {}
    if not json_path.exists():
        return overrides
    try:
        with open(json_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
    except Exception:
        return overrides

    for _, entries in data.items():
        if isinstance(entries, list):
            for entry in entries:
                if isinstance(entry, dict) and "code" in entry:
                    code = entry.get("code")
                    status = entry.get("status")
                    colour = entry.get("colour")
                    if code and status:
                        overrides[code] = {"status": status, "colour": colour}
    return overrides

def apply_horse_overrides(wb, ws, col_index, overrides):
    header_row_idx = None
    for row in ws.iter_rows(min_row=1, max_row=100, values_only=False):
        for cell in row:
            if cell.value and "CLIENT PO" in str(cell.value).upper():
                header_row_idx = cell.row
                break
        if header_row_idx is not None:
            break
    if header_row_idx is None:
        return

    updated_count = 0
    for row in ws.iter_rows(min_row=header_row_idx+1, values_only=False):
        horse_cell = row[col_index["HORSE REG NO"] - 1]
        if not horse_cell.value:
            continue
        reg = str(horse_cell.value).strip()
        if reg in overrides:
            override = overrides[reg]
            status_cell = row[col_index["ACTUAL STATUS"] - 1]
            old_value = status_cell.value
            status_cell.value = override["status"].upper()
            if override.get("colour") == "pink":
                status_cell.fill = PINK_FILL
            updated_count += 1
    print(f"📌 Total manual overrides applied: {updated_count}")

# ------------------- MAIN EXECUTION BRIDGE -------------------
def update_bartrac_file(bartrac_path, fml_path, oriento_paths, natrans_paths, vanito_path, horse_json):
    print(f"\n{'='*60}")
    print(f"Processing: {bartrac_path.name}")
    print(f"{'='*60}")

    sheet_name = get_sheet_name_for_bartrac(bartrac_path)
    try:
        wb = load_workbook(bartrac_path)
        if sheet_name not in wb.sheetnames:
            if "ENROUTE SITE" in wb.sheetnames: sheet_name = "ENROUTE SITE"
            elif "CURRENT SHIPMENTS" in wb.sheetnames: sheet_name = "CURRENT SHIPMENTS"
            else: raise ValueError(f"No match for sheet tracking structure.")
    except Exception as e:
        print(f"❌ Cannot open or find tracking sheet context: {e}")
        return

    print(f"Using target tracking sheet: '{sheet_name}'")
    create_backup(bartrac_path)
    
    wb = load_workbook(bartrac_path)
    ws = wb[sheet_name]

    header_row_idx = None
    header_row = None
    for row in ws.iter_rows(min_row=1, max_row=100, values_only=False):
        for cell in row:
            if cell.value and "CLIENT PO" in str(cell.value).upper():
                header_row_idx = cell.row
                header_row = row
                break
        if header_row_idx is not None:
            break

    if header_row_idx is None:
        print("❌ Could not locate header column array matching standard schema.")
        return

    col_index = {str(cell.value).strip(): col_idx for col_idx, cell in enumerate(header_row, start=1) if cell.value}

    # Run engines sequentially
    process_fml_file(fml_path, wb, ws, col_index)
    for oriento_path in oriento_paths:
        process_oriento_file(oriento_path, wb, ws, col_index)
    for natrans_path in natrans_paths:
        process_natrans_file(natrans_path, wb, ws, col_index)
    process_vanito_file(vanito_path, wb, ws, col_index)

    overrides = load_horse_overrides(horse_json)
    if overrides:
        apply_horse_overrides(wb, ws, col_index, overrides)

    print(f"💾 Saving changes: {bartrac_path}")
    wb.save(bartrac_path)

def main():
    # 1. Dynamically find all relevant files inside your vendor directory
    fml_path, oriento_paths, natrans_paths, vanito_path = discover_vendor_files(VENDOR_FOLDER)

    # 2. Dynamically pull all tracking target sheets from BARTRAC folder
    bartrac_files = [f for f in BARTRAC_FOLDER.glob("*.xlsx") if not f.name.startswith("~$")]
    
    if not bartrac_files:
        print(f"No Target .xlsx files found in {BARTRAC_FOLDER}")
        sys.exit(0)

    print(f"Found {len(bartrac_files)} master BARTRAC file(s) to process dynamically.")
    
    # 3. Run updates
    for bartrac_file in bartrac_files:
        try:
            update_bartrac_file(bartrac_file, fml_path, oriento_paths, natrans_paths, vanito_path, HORSE_OVERRIDE_JSON)
        except Exception as e:
            print(f"❌ Failed to process {bartrac_file.name}: {e}")

    print("\n✅ All operations completed successfully.")

if __name__ == "__main__":
    main()

```

---

### What changed?

1. **Removed Global String Arrays:** The hardcoded `ORIENTO_FILES` array, `NATRANS_FILE` variables, and commented out `FML_FILE` values are gone.
2. **Added `discover_vendor_files()` engine:** This helper scans `C:\Users\Jason\Projects\TRACKING-TASKS2\vendor-report` using broad rules (like checking if the file contains the keyword `ORIENTO`, `NATRANS`, or `VANITO`). This handles daily naming format drifts gracefully without throwing code errors.
3. **Resiliency Over Failures:** If a specific vendor doesn't drop a file on a given day (e.g., Vanito skips a day), the engine flags a soft warning message in the console pipeline and seamlessly continues tracking updates for everything else instead of crashing with a `FileNotFoundError`.


Yes, this makes absolute sense and is exactly how you move away from brittle scripts to a truly zero-maintenance pipeline!

Instead of trying to keep up with changing dates, references, and job numbers by editing the script every day, reading the structure directly from your `.json` manifests completely decouples the script logic from the filenames.

Implementing this strategy provides several key benefits:

* **Zero Manual Path Edits:** When new vendor updates drop or a new BARTRAC reporting file is created for the day, your generation process updates `vendor-directory_structure.json` and `bartrac-directory_structure.json`. The Python script automatically adapts.
* **Intelligent Keyword Classification:** Rather than expecting an exact string match, the script uses case-insensitive pattern matching (e.g., looking for `ORIENTO`, `NATRANS`, `VANITO`, or `FML`) across whatever file list the JSON provides.
* **Fail-Safe Architecture:** If the JSONs are missing or malformed on a given run, the script seamlessly falls back to scanning the physical folders via `pathlib` so your tracking pipeline doesn't break.

---

### The Complete, JSON-Driven Script

Here is your fully production-ready, refactored `update_all_bartrac.py` script. It includes robust manifest parsing, pattern-based vendor assignment, and fixes the double-workbook load issue to speed up processing time.

```python
import pandas as pd
import re
import sys
import shutil
import json
from pathlib import Path
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

PINK_FILL = PatternFill(start_color="FFC0CB", end_color="FFC0CB", fill_type="solid")

# ------------------- CONFIGURATION -------------------
# Path manifests that change daily
BARTRAC_JSON = Path(r"C:\Users\Jason\Projects\TRACKING-TASKS2\bartrac-directory_structure.json")
VENDOR_JSON = Path(r"C:\Users\Jason\Projects\TRACKING-TASKS2\vendor-directory_structure.json")

# Core static configurations
HORSE_OVERRIDE_JSON = Path(r"C:\Users\Jason\Projects\TRACKING-TASKS2\horse-registration.json")
DEFAULT_BARTRAC_FOLDER = Path(r"C:\Users\Jason\Projects\TRACKING-TASKS2\BARTRAC-TRACKING")
DEFAULT_VENDOR_FOLDER = Path(r"C:\Users\Jason\Projects\TRACKING-TASKS2\vendor-report")

# ------------------- MANIFEST & DISCOVERY ENGINE -------------------
def load_paths_from_manifest(json_path: Path) -> list:
    """Reads a directory structure JSON manifest and extracts file paths."""
    if not json_path.exists():
        print(f"⚠️ Manifest not found: {json_path}. Falling back to folder scan.")
        return []
    try:
        with open(json_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
        
        children = data.get("children", [])
        paths = []
        for child in children:
            if child.get("type") == "file":
                path_str = child.get("path")
                if path_str:
                    paths.append(Path(path_str))
        return paths
    except Exception as e:
        print(f"⚠️ Error reading manifest {json_path.name}: {e}. Falling back to folder scan.")
        return []

def resolve_vendor_roles(vendor_paths: list):
    """Dynamically categorizes vendor files based on name patterns from the manifest."""
    fml_file = None
    oriento_files = []
    natrans_files = []
    vanito_file = None

    # Sort to ensure predictable processing orders if multiple exist
    for file_path in sorted(vendor_paths):
        name_upper = file_path.name.upper()
        if file_path.name.startswith("~$") or file_path.suffix.lower() != ".xlsx":
            continue

        # 1. Match FML Master file
        if "FML" in name_upper and "6060" in name_upper:
            fml_file = file_path
        # 2. Match Vanito Reports
        elif "VANITO" in name_upper:
            vanito_file = file_path
        # 3. Match NATRANS manifests
        elif "NATRANS" in name_upper or ("FML CAT60" in name_upper and any(k in name_upper for k in ["KOLWEZI", "SAKANIA"])):
            natrans_files.append(file_path)
        # 4. Match ORIENTO Reports
        elif "ORIENTO" in name_upper:
            oriento_files.append(file_path)

    # Robust fallback: If FML master didn't have 6060 but an FML file exists, don't leave it empty
    if not fml_file:
        fml_candidates = [p for p in vendor_paths if "FML" in p.name.upper() and not p.name.startswith("~$")]
        if fml_candidates:
            fml_file = fml_candidates[0]

    print("\n🔍 Dynamic File Discovery Summary:")
    print(f"  🔹 FML Master:      {fml_file.name if fml_file else '❌ NOT FOUND'}")
    print(f"  🔹 Vanito Report:   {vanito_file.name if vanito_file else '⚠️ NOT FOUND'}")
    print(f"  🔹 NATRANS Files:   {[f.name for f in natrans_files]}")
    print(f"  🔹 ORIENTO Files:   {[f.name for f in oriento_files]}")
    print("-" * 60)

    return fml_file, oriento_files, natrans_files, vanito_file

# ------------------- CORE UTILITIES -------------------
def create_backup(file_path):
    file_path = Path(file_path)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_name = f"{file_path.stem}_backup_{timestamp}{file_path.suffix}"
    backup_path = file_path.parent / backup_name
    shutil.copy2(file_path, backup_path)
    print(f"📁 Backup created: {backup_path.name}")
    return backup_path

def extract_component_keyword(comp_name):
    cleaned = comp_name.replace("CAT 6060", "").strip()
    return cleaned

def get_sheet_name_for_bartrac(bartrac_path):
    filename = bartrac_path.name.lower()
    if "congo" in filename or "erg" in filename:
        return "CURRENT SHIPMENTS"
    elif any(k in filename for k in ["kamoa", "kcc", "mumi"]):
        return "ENROUTE SITE"
    return None

def find_oriento_sheet(file_path):
    try:
        xl = pd.ExcelFile(file_path)
        for sheet in xl.sheet_names:
            df_sample = pd.read_excel(file_path, sheet_name=sheet, header=None, nrows=20)
            for _, row in df_sample.iterrows():
                for cell in row.iloc[:20]:
                    if pd.notna(cell) and "TRUCK NUMBER" in str(cell).upper():
                        return sheet
        return None
    except Exception as e:
        print(f"⚠️ Could not read sheets from {file_path.name}: {e}")
        return None

# ------------------- PROCESSING ENGINES -------------------
def process_oriento_file(oriento_path, wb, ws, col_index):
    if not oriento_path.exists():
        return

    if "BA2951" in oriento_path.name or "2603DSI2788" in oriento_path.name:
        sheet_name = "2603DSI2788"
    else:
        sheet_name = find_oriento_sheet(oriento_path)
        if sheet_name is None:
            print(f"⚠️ Could not auto-detect sheet in ORIENTO file {oriento_path.name}. Skipping.")
            return

    try:
        df_raw = pd.read_excel(oriento_path, sheet_name=sheet_name, header=None, dtype=str)
    except Exception as e:
        print(f"⚠️ Could not read sheet '{sheet_name}' in {oriento_path.name}: {e}")
        return

    header_idx = None
    for idx, row in df_raw.iterrows():
        for cell in row.iloc[:20]:
            if pd.notna(cell) and "TRUCK NUMBER" in str(cell).upper():
                header_idx = idx
                break
        if header_idx is not None:
            break
    if header_idx is None:
        return

    header_row = df_raw.iloc[header_idx]
    truck_col, location_col, status_col = None, None, None
    for i, val in enumerate(header_row):
        if pd.notna(val):
            val_str = str(val).strip().upper()
            if "TRUCK NUMBER" in val_str: truck_col = i
            elif "CURRENT LOCATION" in val_str: location_col = i
            elif "CURRENT STATUS" in val_str: status_col = i

    if truck_col is None: return

    data_rows = df_raw.iloc[header_idx + 1:].copy()
    data_rows.columns = range(data_rows.shape[1])

    truck_to_status = {}
    for _, row in data_rows.iterrows():
        truck = str(row.iloc[truck_col]).strip() if pd.notna(row.iloc[truck_col]) else ""
        if not truck or truck == 'nan': continue

        location = str(row.iloc[location_col]).strip() if location_col is not None and pd.notna(row.iloc[location_col]) else ""
        status = str(row.iloc[status_col]).strip() if status_col is not None and pd.notna(row.iloc[status_col]) else ""

        if location and status: combined = f"{location} - {status}"
        elif location: combined = location
        elif status: combined = status
        else: continue

        truck_to_status[truck] = combined

    header_row_idx = None
    for row in ws.iter_rows(min_row=1, max_row=100, values_only=False):
        for cell in row:
            if cell.value and "CLIENT PO" in str(cell.value).upper():
                header_row_idx = cell.row
                break
        if header_row_idx is not None: break

    updated_count = 0
    for truck_reg, status in truck_to_status.items():
        for row in ws.iter_rows(min_row=header_row_idx+1, values_only=False):
            horse_cell = row[col_index["HORSE REG NO"] - 1]
            if horse_cell.value and str(horse_cell.value).strip() == truck_reg.strip():
                status_cell = row[col_index["ACTUAL STATUS"] - 1]
                status_cell.value = status.upper()
                status_cell.fill = PINK_FILL
                updated_count += 1
    print(f"✅ ORIENTO ({sheet_name}): Updated {updated_count} rows.")

def process_fml_file(fml_path, wb, ws, col_index):
    if not fml_path or not fml_path.exists():
        return
    
    df_fml_raw = pd.read_excel(fml_path, sheet_name="Sheet1", header=None, dtype=str)

    load_desc_idx = None
    for idx, row in df_fml_raw.iterrows():
        if str(row.iloc[0]).strip() == "Load Desc.":
            load_desc_idx = idx
            break
    if load_desc_idx is None: return
    
    components = [str(c).strip() for c in df_fml_raw.iloc[load_desc_idx, 1:] if pd.notna(c)]

    truck_idx = None
    for idx, row in df_fml_raw.iterrows():
        if str(row.iloc[0]).strip() == "Truck":
            truck_idx = idx
            break
    if truck_idx is None: return
    trucks = [str(t).strip() if pd.notna(t) else "" for t in df_fml_raw.iloc[truck_idx, 1:]]

    comp_to_truck = {comp: truck for comp, truck in zip(components, trucks) if truck}

    # Generic 4-digit year regex to stay future-proof
    date_pattern = re.compile(r'^\d{4}-\d{2}-\d{2}')
    data_rows = [idx for idx, row in df_fml_raw.iterrows() if date_pattern.match(str(row.iloc[0]).strip())]
    if not data_rows: return
    
    last_data_idx = data_rows[-1]
    latest_statuses = [str(s).strip() if pd.notna(s) else "" for s in df_fml_raw.iloc[last_data_idx, 1:]]
    comp_to_status = {comp: status for comp, status in zip(components, latest_statuses) if comp}

    header_row_idx = None
    for row in ws.iter_rows(min_row=1, max_row=100, values_only=False):
        for cell in row:
            if cell.value and "CLIENT PO" in str(cell.value).upper():
                header_row_idx = cell.row
                break
        if header_row_idx is not None: break

    updated_count = 0
    for comp, truck_reg in comp_to_truck.items():
        status = comp_to_status.get(comp, "")
        if not status: continue

        for row in ws.iter_rows(min_row=header_row_idx+1, values_only=False):
            horse_cell = row[col_index["HORSE REG NO"] - 1]
            if horse_cell.value and str(horse_cell.value).strip() == truck_reg.strip():
                status_cell = row[col_index["ACTUAL STATUS"] - 1]
                status_cell.value = status.upper()
                status_cell.fill = PINK_FILL
                updated_count += 1
    print(f"✅ FML: Updated {updated_count} rows.")

def process_natrans_file(natrans_path, wb, ws, col_index):
    if not natrans_path.exists(): return

    try:
        df_raw = pd.read_excel(natrans_path, sheet_name="NATRANS", header=None, dtype=str)
    except ValueError:
        return

    header_idx = None
    for idx, row in df_raw.iterrows():
        for cell in row.iloc[:10]:
            if pd.notna(cell) and str(cell).strip().upper() in ["LOAD", "DISP", "BB SA", "BB ZIM"]:
                header_idx = idx
                break
        if header_idx is not None: break
    if header_idx is None: return

    header_row = df_raw.iloc[header_idx]
    col_names = {str(val).strip(): i for i, val in enumerate(header_row) if pd.notna(val)}

    horse_col = next((col_names[p] for p in ["HORSE", "HORSE REG NO", "HORSE REG"] if p in col_names), None)
    if horse_col is None: return

    milestone_columns = [col for col in ["LOAD", "DISP", "BB SA", "BB ZIM", "CHI ZIM", "CHI ZAM", 
                                        "KASUM ZAM", "KASUM DRC", "WHISKEY", "SITE", "OFFLOAD", 
                                        "LUFUA ARR", "LUFUA DISP"] if col in col_names]

    data_rows = df_raw.iloc[header_idx + 1:].copy()
    data_rows.columns = range(data_rows.shape[1])

    truck_to_status = {}
    for _, row in data_rows.iterrows():
        truck = str(row.iloc[horse_col]).strip() if pd.notna(row.iloc[horse_col]) else ""
        if not truck or truck == 'nan': continue

        last_status = None
        for milestone in milestone_columns:
            cell_value = row.iloc[col_names[milestone]]
            if pd.notna(cell_value) and str(cell_value).strip() not in ['', 'nan']:
                last_status = milestone

        if "COMMENTS" in col_names and pd.notna(row.iloc[col_names["COMMENTS"]]) and str(row.iloc[col_names["COMMENTS"]]).strip() not in ['', 'nan']:
            truck_to_status[truck] = str(row.iloc[col_names["COMMENTS"]]).strip()
        elif last_status:
            truck_to_status[truck] = last_status

    header_row_idx = None
    for row in ws.iter_rows(min_row=1, max_row=100, values_only=False):
        for cell in row:
            if cell.value and "CLIENT PO" in str(cell.value).upper():
                header_row_idx = cell.row
                break
        if header_row_idx is not None: break

    updated_count = 0
    for truck_reg, status in truck_to_status.items():
        for row in ws.iter_rows(min_row=header_row_idx+1, values_only=False):
            horse_cell = row[col_index["HORSE REG NO"] - 1]
            if horse_cell.value and str(horse_cell.value).strip() == truck_reg.strip():
                status_cell = row[col_index["ACTUAL STATUS"] - 1]
                status_cell.value = status.upper()
                status_cell.fill = PINK_FILL
                updated_count += 1
    print(f"✅ NATRANS ({natrans_path.name}): Updated {updated_count} rows.")

def process_vanito_file(vanito_path, wb, ws, col_index):
    if not vanito_path or not vanito_path.exists(): return

    try:
        xl = pd.ExcelFile(vanito_path)
        sheets = xl.sheet_names
    except Exception: return

    date_pattern = re.compile(r'^(\d{1,2})\s+([A-Z]{3,})\s+(\d{4})$')
    sheet_date_map = {}
    for sheet in sheets:
        m = date_pattern.match(sheet.strip().upper())
        if m:
            day, month_str, year = m.groups()
            try:
                month_num = datetime.strptime(month_str, "%b").month
                dt = datetime(int(year), month_num, int(day))
                sheet_date_map[dt] = sheet
            except ValueError: continue

    if not sheet_date_map: return
    latest_sheet = sheet_date_map[max(sheet_date_map.keys())]

    try:
        df_raw = pd.read_excel(vanito_path, sheet_name=latest_sheet, header=None, dtype=str)
    except Exception: return

    header_idx = None
    for idx, row in df_raw.iterrows():
        for cell in row.iloc[:20]:
            if pd.notna(cell) and "HORSE REG" in str(cell).upper():
                header_idx = idx
                break
        if header_idx is not None: break
    if header_idx is None: return

    header_row = df_raw.iloc[header_idx]
    horse_col, location_col = None, None
    for i, val in enumerate(header_row):
        val_str = str(val).strip().upper()
        if "HORSE REG" in val_str: horse_col = i
        elif any(k in val_str for k in ["CURRENT LOCATION", "CURRENT STATUS", "LOCATION"]): location_col = i

    if horse_col is None or location_col is None: return

    data_rows = df_raw.iloc[header_idx + 1:].copy()
    data_rows.columns = range(data_rows.shape[1])

    truck_to_status = {}
    for _, row in data_rows.iterrows():
        truck = str(row.iloc[horse_col]).strip() if pd.notna(row.iloc[horse_col]) else ""
        location = str(row.iloc[location_col]).strip() if pd.notna(row.iloc[location_col]) else ""
        if truck and location and truck != 'nan' and location != 'nan':
            truck_to_status[truck] = location

    bartrac_header_idx = None
    for row in ws.iter_rows(min_row=1, max_row=100, values_only=False):
        for cell in row:
            if cell.value and "CLIENT PO" in str(cell.value).upper():
                bartrac_header_idx = cell.row
                break
        if bartrac_header_idx is not None: break

    updated_count = 0
    for truck_reg, status in truck_to_status.items():
        for row in ws.iter_rows(min_row=bartrac_header_idx+1, values_only=False):
            horse_cell = row[col_index["HORSE REG NO"] - 1]
            if horse_cell.value and str(horse_cell.value).strip() == truck_reg.strip():
                status_cell = row[col_index["ACTUAL STATUS"] - 1]
                status_cell.value = status.upper()
                status_cell.fill = PINK_FILL
                updated_count += 1
    print(f"✅ VANITO ({latest_sheet}): Updated {updated_count} rows.")

def load_horse_overrides(json_path):
    overrides = {}
    if not json_path.exists(): return overrides
    try:
        with open(json_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
    except Exception: return overrides

    for _, entries in data.items():
        if isinstance(entries, list):
            for entry in entries:
                if isinstance(entry, dict) and "code" in entry:
                    code = entry.get("code")
                    status = entry.get("status")
                    colour = entry.get("colour")
                    if code and status:
                        overrides[code] = {"status": status, "colour": colour}
    return overrides

def apply_horse_overrides(ws, col_index, overrides):
    header_row_idx = None
    for row in ws.iter_rows(min_row=1, max_row=100, values_only=False):
        for cell in row:
            if cell.value and "CLIENT PO" in str(cell.value).upper():
                header_row_idx = cell.row
                break
        if header_row_idx is not None: break

    updated_count = 0
    for row in ws.iter_rows(min_row=header_row_idx+1, values_only=False):
        horse_cell = row[col_index["HORSE REG NO"] - 1]
        if not horse_cell.value: continue
        reg = str(horse_cell.value).strip()
        if reg in overrides:
            override = overrides[reg]
            status_cell = row[col_index["ACTUAL STATUS"] - 1]
            status_cell.value = override["status"].upper()
            if override.get("colour") == "pink":
                status_cell.fill = PINK_FILL
            updated_count += 1
    print(f"📌 Total manual overrides applied: {updated_count}")

# ------------------- WORKFLOW BRIDGE -------------------
def update_bartrac_file(bartrac_path, fml_path, oriento_paths, natrans_paths, vanito_path, horse_json):
    print(f"\n{'='*60}\nProcessing: {bartrac_path.name}\n{'='*60}")

    sheet_name = get_sheet_name_for_bartrac(bartrac_path)
    try:
        wb = load_workbook(bartrac_path)
        if sheet_name not in wb.sheetnames:
            if "ENROUTE SITE" in wb.sheetnames: sheet_name = "ENROUTE SITE"
            elif "CURRENT SHIPMENTS" in wb.sheetnames: sheet_name = "CURRENT SHIPMENTS"
            else: raise ValueError("No standard tracker matching sheets.")
    except Exception as e:
        print(f"❌ Aborted workbook parsing: {e}")
        return

    print(f"Targeting active tracking sheet: '{sheet_name}'")
    create_backup(bartrac_path)
    
    ws = wb[sheet_name]

    header_row = None
    for row in ws.iter_rows(min_row=1, max_row=100, values_only=False):
        for cell in row:
            if cell.value and "CLIENT PO" in str(cell.value).upper():
                header_row = row
                break
        if header_row is not None: break

    if header_row is None:
        print("❌ Could not locate structural layout schema.")
        return

    col_index = {str(cell.value).strip().upper(): col_idx for col_idx, cell in enumerate(header_row, start=1) if cell.value}

    # Execute synchronization engines
    process_fml_file(fml_path, wb, ws, col_index)
    for oriento_path in oriento_paths:
        process_oriento_file(oriento_path, wb, ws, col_index)
    for natrans_path in natrans_paths:
        process_natrans_file(natrans_path, wb, ws, col_index)
    process_vanito_file(vanito_path, wb, ws, col_index)

    overrides = load_horse_overrides(horse_json)
    if overrides:
        apply_horse_overrides(ws, col_index, overrides)

    print(f"💾 Saving updates to tracking: {bartrac_path.name}")
    wb.save(bartrac_path)

def main():
    # 1. Pull dynamic file lists from the manifests
    vendor_paths = load_paths_from_manifest(VENDOR_JSON)
    bartrac_manifest_paths = load_paths_from_manifest(BARTRAC_JSON)

    # 2. Folder scan fallbacks if json maps are blank/missing
    if not vendor_paths and DEFAULT_VENDOR_FOLDER.exists():
        vendor_paths = list(DEFAULT_VENDOR_FOLDER.glob("*.xlsx"))

    if bartrac_manifest_paths:
        bartrac_files = [p for p in bartrac_manifest_paths if p.suffix.lower() == ".xlsx" and not p.name.startswith("~$")]
    elif DEFAULT_BARTRAC_FOLDER.exists():
        bartrac_files = [f for f in DEFAULT_BARTRAC_FOLDER.glob("*.xlsx") if not f.name.startswith("~$")]
    else:
        bartrac_files = []

    if not bartrac_files:
        print("❌ Critical Execution Failure: No target files resolved to update.")
        sys.exit(1)

    # 3. Classify paths to correct logic blocks using naming footprints
    fml_path, oriento_paths, natrans_paths, vanito_path = resolve_vendor_roles(vendor_paths)

    # 4. Trigger Batch Processing Loop
    print(f"Beginning sync for {len(bartrac_files)} tracking targets.")
    for bartrac_file in bartrac_files:
        try:
            update_bartrac_file(bartrac_file, fml_path, oriento_paths, natrans_paths, vanito_path, HORSE_OVERRIDE_JSON)
        except Exception as e:
            print(f"❌ Error processing tracking data file {bartrac_file.name}: {e}")

    print("\n✅ Execution loop finalized successfully.")

if __name__ == "__main__":
    main()

```