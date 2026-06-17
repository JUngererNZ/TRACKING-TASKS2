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
BARTRAC_FOLDER = r"C:\Users\Jason\Projects\TRACKING-TASKS2\BARTRAC-TRACKING"
VENDOR_FOLDER = r"C:\Users\Jason\Projects\TRACKING-TASKS2\vendor-report"

# FML file (uncommented)
# FML_FILE = r"C:\Users\Jason\Projects\TRACKING-TASKS2\vendor-report\FML CAT 6060.xlsx"

ORIENTO_FILES = [
    r"C:\Users\Jason\Projects\TRACKING-TASKS2\vendor-report\1 X CATERPILLAR 6030 IN CKD FORM - LOAD10-13 - VARIOUS - 2 X LINK+ 1 X TRI AXLE - 2604DSI2802- BA3198 -DURBAN PORT TO MUTANDA MINING, DRC.xlsx",
    r"C:\Users\Jason\Projects\TRACKING-TASKS2\vendor-report\1 X CATERPILLAR 6060 IN CKD FORM - LOAD 10-16 - VARIOUS - 6 X LINKS - 2604DSI2804- BA3188 -DURBAN PORT TO KAMOTO COPPER COMPANY.xlsx",
    r"C:\Users\Jason\Projects\TRACKING-TASKS2\vendor-report\ORIENTO TRACKING REPORT TO FML (DURBAN PORT TO FRONTIER MINE)...xlsx",
    r"C:\Users\Jason\Projects\TRACKING-TASKS2\vendor-report\ORIENTO TRACKING REPORT TO FML (DURBAN PORT TO KCC 16 June 2026.xlsx",
    r"C:\Users\Jason\Projects\TRACKING-TASKS2\vendor-report\ORIENTO TRACKING REPORT. - 1 X TRI-AXLE TO LOAD 1 x EG20 MOTOR GRADER -2603DSI2788 - BA2951-FREIGHTSTATIONS SA DURBAN TO KAMOA COPPER SA KOLWEZI DRC.xlsx"
]

NATRANS_FILE1 = r"C:\Users\Jason\Projects\TRACKING-TASKS2\vendor-report\FML CAT6060 KOLWEZI BA3188.xlsx"
NATRANS_FILE2 = r"C:\Users\Jason\Projects\TRACKING-TASKS2\vendor-report\FML CAT6030 SAKANIA BA3159.xlsx"

VANITO_FILE = r"C:\Users\Jason\Projects\TRACKING-TASKS2\vendor-report\Vanito Tracking 2026 - FML DBN TO DRC.xlsx"

HORSE_OVERRIDE_JSON = r"C:\Users\Jason\Projects\TRACKING-TASKS2\horse-registration.json"

# ------------------------------------------------------------------
# (the rest of the script remains exactly as you had it)
# ------------------------------------------------------------------

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

def process_oriento_file(oriento_path, wb, ws, col_index):
    """Generic ORIENTO processor, with special case for BA2951 (EG20)."""
    oriento_path = Path(oriento_path)
    if not oriento_path.exists():
        print(f"⚠️ ORIENTO file not found: {oriento_path}")
        return

    # Special case for BA2951 (EG20) file – force sheet "2603DSI2788"
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

    # Find header row
    header_idx = None
    for idx, row in df_raw.iterrows():
        for cell in row.iloc[:20]:
            if pd.notna(cell) and "TRUCK NUMBER" in str(cell).upper():
                header_idx = idx
                break
        if header_idx is not None:
            break
    if header_idx is None:
        raise ValueError(f"Could not find header row in ORIENTO sheet {sheet_name}")

    header_row = df_raw.iloc[header_idx]
    truck_col = None
    location_col = None
    status_col = None
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
        raise ValueError(f"Could not find 'TRUCK NUMBER' column in ORIENTO sheet {sheet_name}")
    if location_col is None and status_col is None:
        raise ValueError(f"Could not find 'CURRENT LOCATION' or 'CURRENT STATUS' column in ORIENTO sheet {sheet_name}")

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

    print(f"Found {len(truck_to_status)} truck entries in ORIENTO ({sheet_name})")

    # Find BARTRAC header row
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

        if not matching_rows:
            print(f"❌ ORIENTO ({sheet_name}): No row found for truck '{truck_reg}'")
            continue

        for row in matching_rows:
            status_cell = row[col_index["ACTUAL STATUS"] - 1]
            old_value = status_cell.value
            status_cell.value = status.upper()
            status_cell.fill = PINK_FILL
            updated_count += 1
            print(f"✅ ORIENTO ({sheet_name}): Updated row {row[0].row} (truck {truck_reg}) → '{status.upper()}' (was: '{old_value}')")

    print(f"📊 ORIENTO ({sheet_name}) total rows updated: {updated_count}")

def process_fml_file(fml_path, wb, ws, col_index):
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
            print(f"⚠️ FML: No status found for component: {comp}")
            continue

        matching_rows = []
        for row in ws.iter_rows(min_row=header_row_idx+1, values_only=False):
            horse_cell = row[col_index["HORSE REG NO"] - 1]
            if horse_cell.value and str(horse_cell.value).strip() == truck_reg.strip():
                cargo_cell = row[col_index["CARGO DETAILS"] - 1]
                if cargo_cell.value and keyword.upper() in str(cargo_cell.value).upper():
                    matching_rows.append(row)
                else:
                    matching_rows.append(row)

        if not matching_rows:
            print(f"❌ FML: No row found for truck '{truck_reg}' (component: {comp})")
            continue

        for row in matching_rows:
            status_cell = row[col_index["ACTUAL STATUS"] - 1]
            old_value = status_cell.value
            status_cell.value = status.upper()
            status_cell.fill = PINK_FILL
            updated_count += 1
            print(f"✅ FML: Updated row {row[0].row} for {comp} (truck {truck_reg}) → '{status.upper()}' (was: '{old_value}')")

    print(f"📊 FML total rows updated: {updated_count}")

def process_natrans_file(natrans_path, wb, ws, col_index):
    natrans_path = Path(natrans_path)
    if not natrans_path.exists():
        print(f"⚠️ NATRANS file not found: {natrans_path}")
        return

    print(f"Reading NATRANS file: {natrans_path}")
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
        raise ValueError(f"Could not find header row in NATRANS sheet of {natrans_path.name}")

    header_row = df_raw.iloc[header_idx]
    col_names = {}
    for i, val in enumerate(header_row):
        if pd.notna(val):
            col_names[str(val).strip()] = i

    horse_col = None
    for possible in ["HORSE", "HORSE REG NO", "HORSE REG"]:
        if possible in col_names:
            horse_col = col_names[possible]
            break
    if horse_col is None:
        raise ValueError("Could not find HORSE column in NATRANS file")

    milestone_columns = []
    for col in ["LOAD", "DISP", "BB SA", "BB ZIM", "CHI ZIM", "CHI ZAM", 
                "KASUM ZAM", "KASUM DRC", "WHISKEY", "SITE", "OFFLOAD", 
                "LUFUA ARR", "LUFUA DISP"]:
        if col in col_names:
            milestone_columns.append(col)
    for col, idx in col_names.items():
        if any(keyword in col.upper() for keyword in ["ARR", "DISP", "ZAM", "DRC", "WHISKEY", "SITE", "OFFLOAD"]):
            if col not in milestone_columns:
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

    print(f"Found {len(truck_to_status)} truck entries in {natrans_path.name}")

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

        if not matching_rows:
            print(f"❌ NATRANS ({natrans_path.name}): No row found for truck '{truck_reg}'")
            continue

        for row in matching_rows:
            status_cell = row[col_index["ACTUAL STATUS"] - 1]
            old_value = status_cell.value
            status_cell.value = status.upper()
            status_cell.fill = PINK_FILL
            updated_count += 1
            print(f"✅ NATRANS ({natrans_path.name}): Updated row {row[0].row} (truck {truck_reg}) → '{status.upper()}' (was: '{old_value}')")

    print(f"📊 NATRANS ({natrans_path.name}) total rows updated: {updated_count}")

def process_vanito_file(vanito_path, wb, ws, col_index):
    vanito_path = Path(vanito_path)
    if not vanito_path.exists():
        print(f"⚠️ Vanito file not found: {vanito_path}")
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
    print(f"Using latest sheet: '{latest_sheet}' (date: {latest_date.strftime('%Y-%m-%d')})")

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
    horse_col = None
    location_col = None
    for i, val in enumerate(header_row):
        val_str = str(val).strip().upper()
        if "HORSE REG" in val_str:
            horse_col = i
        elif "CURRENT LOCATION" in val_str:
            location_col = i
        if location_col is None and "CURRENT STATUS" in val_str:
            location_col = i
        if location_col is None and "LOCATION" in val_str:
            location_col = i

    if horse_col is None:
        print(f"⚠️ Could not find 'HORSE REG' column in sheet '{latest_sheet}'")
        return
    if location_col is None:
        print(f"⚠️ Could not find 'CURRENT LOCATION' (or equivalent) column in sheet '{latest_sheet}'")
        return

    data_rows = df_raw.iloc[header_idx + 1:].copy()
    data_rows.columns = range(data_rows.shape[1])

    truck_to_status = {}
    for _, row in data_rows.iterrows():
        truck = str(row.iloc[horse_col]).strip() if pd.notna(row.iloc[horse_col]) else ""
        location = str(row.iloc[location_col]).strip() if pd.notna(row.iloc[location_col]) else ""
        if truck and location and truck != 'nan' and location != 'nan':
            truck_to_status[truck] = location

    print(f"Found {len(truck_to_status)} truck entries in Vanito (sheet: {latest_sheet})")

    bartrac_header_idx = None
    for row in ws.iter_rows(min_row=1, max_row=100, values_only=False):
        for cell in row:
            if cell.value and "CLIENT PO" in str(cell.value).upper():
                bartrac_header_idx = cell.row
                break
        if bartrac_header_idx is not None:
            break
    if bartrac_header_idx is None:
        raise ValueError("Could not find header row containing 'CLIENT PO' in BARTRAC sheet")

    updated_count = 0
    for truck_reg, status in truck_to_status.items():
        matching_rows = []
        for row in ws.iter_rows(min_row=bartrac_header_idx+1, values_only=False):
            horse_cell = row[col_index["HORSE REG NO"] - 1]
            if horse_cell.value and str(horse_cell.value).strip() == truck_reg.strip():
                matching_rows.append(row)

        if not matching_rows:
            print(f"❌ Vanito: No row found for truck '{truck_reg}'")
            continue

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
    if not Path(json_path).exists():
        print(f"⚠️ Horse registration file not found: {json_path}")
        return overrides
    try:
        with open(json_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
    except json.JSONDecodeError as e:
        print(f"❌ Invalid JSON in {json_path}: {e}")
        return overrides

    for section, entries in data.items():
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
        print("❌ Could not find header row for manual overrides")
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
            print(f"🎨 MANUAL OVERRIDE: Row {row[0].row} | {reg} → '{override['status'].upper()}' (was '{old_value}') colour={override.get('colour')}")
    print(f"📌 Total manual overrides applied: {updated_count}")

def update_bartrac_file(bartrac_path, fml_path, oriento_paths, natrans_paths, vanito_path, horse_json):
    print(f"\n{'='*60}")
    print(f"Processing: {bartrac_path.name}")
    print(f"{'='*60}")

    sheet_name = get_sheet_name_for_bartrac(bartrac_path)
    if sheet_name is None:
        try:
            wb = load_workbook(bartrac_path)
            if "ENROUTE SITE" in wb.sheetnames:
                sheet_name = "ENROUTE SITE"
            elif "CURRENT SHIPMENTS" in wb.sheetnames:
                sheet_name = "CURRENT SHIPMENTS"
            else:
                raise ValueError(f"No known sheet found in {bartrac_path.name}")
        except Exception as e:
            print(f"❌ Cannot determine sheet for {bartrac_path.name}: {e}")
            return
    else:
        try:
            wb = load_workbook(bartrac_path)
            if sheet_name not in wb.sheetnames:
                for s in wb.sheetnames:
                    if s.strip() == sheet_name.strip():
                        sheet_name = s
                        break
                if sheet_name not in wb.sheetnames:
                    raise ValueError(f"Sheet '{sheet_name}' not found in {bartrac_path.name}")
        except Exception as e:
            print(f"❌ Cannot open or find sheet: {e}")
            return

    print(f"Using sheet: '{sheet_name}'")

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
        raise ValueError("Could not find header row containing 'CLIENT PO' in sheet")

    col_index = {}
    for col_idx, cell in enumerate(header_row, start=1):
        if cell.value:
            col_name = str(cell.value).strip()
            col_index[col_name] = col_idx

    required = ["CARGO DETAILS", "HORSE REG NO", "ACTUAL STATUS"]
    for col in required:
        if col not in col_index:
            raise ValueError(f"Column '{col}' not found in header row")

    process_fml_file(fml_path, wb, ws, col_index)
    for oriento_path in oriento_paths:
        process_oriento_file(oriento_path, wb, ws, col_index)
    for natrans_path in natrans_paths:
        process_natrans_file(natrans_path, wb, ws, col_index)
    process_vanito_file(vanito_path, wb, ws, col_index)

    overrides = load_horse_overrides(horse_json)
    if overrides:
        apply_horse_overrides(wb, ws, col_index, overrides)
    else:
        print("ℹ️ No manual overrides loaded.")

    print(f"💾 Saving changes to original file: {bartrac_path}")
    wb.save(bartrac_path)

def main():
    fml_path = Path(FML_FILE)
    oriento_paths = [Path(p) for p in ORIENTO_FILES]
    natrans_paths = [Path(NATRANS_FILE1), Path(NATRANS_FILE2)]
    vanito_path = Path(VANITO_FILE)
    bartrac_folder = Path(BARTRAC_FOLDER)
    horse_json = Path(HORSE_OVERRIDE_JSON)

    if not bartrac_folder.exists():
        print(f"ERROR: BARTRAC folder does not exist: {bartrac_folder}")
        sys.exit(1)

    if not fml_path.exists():
        print(f"ERROR: FML file not found: {fml_path}")
        sys.exit(1)

    bartrac_files = list(bartrac_folder.glob("*.xlsx"))
    # Filter out temporary files (starting with ~$)
    bartrac_files = [f for f in bartrac_files if not f.name.startswith("~$")]
    if not bartrac_files:
        print(f"No .xlsx files found in {bartrac_folder}")
        sys.exit(0)

    print(f"Found {len(bartrac_files)} BARTRAC file(s) to process.")
    for bartrac_file in bartrac_files:
        try:
            update_bartrac_file(bartrac_file, fml_path, oriento_paths, natrans_paths, vanito_path, horse_json)
        except Exception as e:
            print(f"❌ Failed to process {bartrac_file.name}: {e}")

    print("\n✅ All BARTRAC files processed.")

if __name__ == "__main__":
    main()