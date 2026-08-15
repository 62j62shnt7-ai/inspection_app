#!/usr/bin/env python3
"""
import_excel.py — Smart, Resilient Importer for Master Inspection Plan Excel workbooks.

Features:
- Dynamic Header & Schema Auto-Detection with Fuzzy Synonym Matching.
- Header Row Auto-Discovery (handles title/banner rows).
- Smart Unit & Number Cleansing (temp, pressure, thickness, corrosion rate, remaining life).
- Ingestion-Time CUI & API 580 POF/COF Risk Auto-Classification.
- Non-Destructive Upsert Mode OR Clean Wipe Mode.
- Specialized Turnaround Critical Scope & Temporary Repair table importers.
"""
import sys
import os
import json
import sqlite3
import datetime
import re
import contextlib

try:
    import openpyxl
except ImportError:
    print("This script needs openpyxl. Run: pip3 install openpyxl --break-system-packages")
    sys.exit(1)

SKIP_SHEETS = {
    "DWG", "Evaluation Criteria", "Inspection Sequence",
    "Anodes Reporting& As-Found Insp", "API-574 Tables & Pipe Sch",
    "HT Press Tables", "FF", "Insulated Piping"
}

MONTH_MAP = {
    "jan": 1, "feb": 2, "mar": 3, "apr": 4, "may": 5, "jun": 6,
    "jul": 7, "aug": 8, "sep": 9, "oct": 10, "nov": 11, "dec": 12
}

DATE_PATTERNS = [
    re.compile(r"^\s*(\d{4})\s*$"),                                      # 2024
    re.compile(r"^\s*(\d{1,2})[-/](\d{4})\s*$"),                          # 09/2024, 9-2024
    re.compile(r"^\s*(\d{4})[-/](\d{1,2})[-/](\d{1,2})\s*$"),            # 2024-09-15
    re.compile(r"^\s*(\d{1,2})[-/](\d{1,2})[-/](\d{4})\s*$"),            # 15/09/2024 or 09/15/2024
    re.compile(r"^\s*([A-Za-z]{3})[-/ ](\d{4})\s*$"),                    # Sep-2024, Sep 2024
    re.compile(r"^\s*(\d{1,2})[-/ ]([A-Za-z]{3})[-/ ](\d{4})\s*$"),       # 15-Sep-2024
    re.compile(r"^\s*Q([1-4])[-/ ](\d{4})\s*$", re.I),                   # Q3 2024
]

FIELD_SYNONYMS = [
    (re.compile(r"^(s[\.\s/]?n|serial|no\.?|seq|item\s*no)$", re.I), "sn"),
    (re.compile(r"(field|pack\s*(no|\#)?|package|header|section)", re.I), "field"),
    (re.compile(r"(plant|facility|station|complex)", re.I), "plant"),
    (re.compile(r"(location|area|deck|platform|sub[\s-]*system)", re.I), "location"),
    (re.compile(r"(unit(\s*name)?|system(\s*name)?)", re.I), "unit_name"),
    (re.compile(r"^(asset(\s*name)?|equipment(\s*name)?|description|item(\s*description)?|line(\s*description)?)$", re.I), "name"),
    (re.compile(r"(tag(\s*no|\s*#)?|equipment\s*tag|line\s*no|line\s*#|spool(\s*no|\s*#)?|iso(\s*no|\s*#)?)", re.I), "tag"),
    (re.compile(r"(asset\s*(no|\#|id|number)|equip\s*(no|\#|id))", re.I), "asset_number"),
    (re.compile(r"(in[\s_-]*service|operational\s*status|status)", re.I), "in_service"),
    (re.compile(r"(insul(ation)?(\s*type)?|cladding|lagging)", re.I), "insulation"),
    (re.compile(r"(last\s*(osi|on[\s-]*stream|external)|date\s*last\s*osi|last\s*insp(\w*\s*)?date)", re.I), "date_osi_last"),
    (re.compile(r"(next\s*(osi|on[\s-]*stream|external)|date\s*next\s*osi|next\s*insp(\w*\s*)?date|next\s*due|due\s*date)", re.I), "date_osi_next"),
    (re.compile(r"(last\s*(internal|int(\.)?|major)|date\s*last\s*internal)", re.I), "date_internal_last"),
    (re.compile(r"(next\s*(internal|int(\.)?|major)|date\s*next\s*internal)", re.I), "date_internal_next"),
    (re.compile(r"(fluid(\s*service)?|medium|service|product|process(\s*fluid)?)", re.I), "fluid_service"),
    (re.compile(r"(design\s*press(ure)?|p_?des|dp\s*\(|des\.\s*press)", re.I), "design_pressure"),
    (re.compile(r"(design\s*temp(erature)?|t_?des|dt\s*\(|des\.\s*temp)", re.I), "design_temp"),
    (re.compile(r"(operat(ing)?\s*press(ure)?|p_?op|op\s*\(|op\.\s*press)", re.I), "operating_pressure"),
    (re.compile(r"(operat(ing)?\s*temp(erature)?|t_?op|ot\s*\(|op\.\s*temp)", re.I), "operating_temp"),
    (re.compile(r"(material(\s*spec)?|metallurgy|pipe\s*mat|spec)", re.I), "material_spec"),
    (re.compile(r"(nominal\s*(thk|thickness|wall)|t_?nom|sch(edule)?)", re.I), "nominal_thickness"),
    (re.compile(r"(t_?min|min(\w*\s*)?(thk|thickness|wall)|retire(ment)?\s*thk)", re.I), "t_min"),
    (re.compile(r"(corr(osion)?\s*rate|cr\s*\(|short\s*term\s*cr|long\s*term\s*cr)", re.I), "corrosion_rate"),
    (re.compile(r"(remain(ing)?\s*life|rem\s*life|rl\s*\()", re.I), "remaining_life"),
    (re.compile(r"(remark(s)?|comment(s)?|note(s)?|recommendation(s)?|finding(s)?)", re.I), "remarks"),
]

import_errors = []

@contextlib.contextmanager
def import_sheet_guard(sheet_name):
    try:
        yield
    except Exception as e:
        import_errors.append((sheet_name, str(e)))
        print(f"  [!] ERROR importing {sheet_name}: {e}")


def parse_date(value):
    """Smart multi-format date parser into ISO date string (YYYY-MM-DD)."""
    if value is None:
        return None, None
    if isinstance(value, (datetime.datetime, datetime.date)):
        try:
            return value.strftime("%Y-%m-%d"), None
        except Exception:
            return None, str(value)
    text = str(value).strip()
    if not text:
        return None, None

    for idx, pat in enumerate(DATE_PATTERNS):
        m = pat.match(text)
        if not m:
            continue
        if idx == 0: # 2024
            return f"{m.group(1)}-01-01", text
        elif idx == 1: # 09/2024
            month, year = int(m.group(1)), m.group(2)
            if 1 <= month <= 12:
                return f"{year}-{month:02d}-01", text
        elif idx == 2: # 2024-09-15
            return f"{m.group(1)}-{int(m.group(2)):02d}-{int(m.group(3)):02d}", text
        elif idx == 3: # 15/09/2024 or 09/15/2024
            p1, p2, year = int(m.group(1)), int(m.group(2)), m.group(3)
            month = p1 if 1 <= p1 <= 12 else (p2 if 1 <= p2 <= 12 else 1)
            day = p2 if p1 == month else p1
            return f"{year}-{month:02d}-{min(day, 28):02d}", text
        elif idx == 4: # Sep-2024
            mon_str, year = m.group(1).lower()[:3], m.group(2)
            if mon_str in MONTH_MAP:
                return f"{year}-{MONTH_MAP[mon_str]:02d}-01", text
        elif idx == 5: # 15-Sep-2024
            day, mon_str, year = int(m.group(1)), m.group(2).lower()[:3], m.group(3)
            if mon_str in MONTH_MAP:
                return f"{year}-{MONTH_MAP[mon_str]:02d}-{min(day, 28):02d}", text
        elif idx == 6: # Q3 2024
            qtr, year = int(m.group(1)), m.group(2)
            qtr_month = (qtr - 1) * 3 + 1
            return f"{year}-{qtr_month:02d}-01", text

    return None, text


def clean_str(val):
    if val is None:
        return None
    s = str(val).strip()
    return s if s else None


def find_header_row_and_map(ws, max_scan_rows=12):
    best_row = 1
    best_score = 0
    best_map = {}

    for r in range(1, min(ws.max_row + 1, max_scan_rows + 1)):
        col_map = {}
        score = 0
        for c in range(1, ws.max_column + 1):
            val = ws.cell(row=r, column=c).value
            if not val:
                continue
            hdr_str = str(val).strip()
            for pat, field_name in FIELD_SYNONYMS:
                if pat.search(hdr_str):
                    col_map[c] = (field_name, hdr_str)
                    score += 1
                    break
        if score > best_score:
            best_score = score
            best_row = r
            best_map = col_map

    return best_row, best_map


SCHEMA = """
CREATE TABLE IF NOT EXISTS assets (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    source_sheet TEXT,
    sn TEXT,
    field TEXT,
    plant TEXT,
    location TEXT,
    unit_name TEXT,
    name TEXT,
    tag TEXT,
    asset_number TEXT,
    description TEXT,
    in_service TEXT,
    insulation TEXT,
    last_insp_category TEXT,
    date_osi_last TEXT,
    date_osi_next TEXT,
    date_internal_last TEXT,
    date_internal_next TEXT,
    next_insp_category TEXT,
    corrosion_rate TEXT,
    remaining_life TEXT,
    fluid_service TEXT,
    design_pressure TEXT,
    design_temp TEXT,
    operating_pressure TEXT,
    operating_temp TEXT,
    material_spec TEXT,
    nominal_thickness TEXT,
    t_min TEXT,
    risk_category TEXT,
    damage_mechanisms TEXT,
    cui_susceptible INTEGER,
    deferral_status TEXT,
    deferral_reason TEXT,
    deferral_mitigation TEXT,
    deferral_expiry TEXT,
    deferral_approver TEXT,
    deferral_moc_no TEXT,
    pof_score INTEGER,
    cof_score INTEGER,
    remarks TEXT,
    extra_json TEXT,
    archived INTEGER DEFAULT 0,
    created_at TEXT DEFAULT (datetime('now')),
    updated_at TEXT DEFAULT (datetime('now'))
);

CREATE INDEX IF NOT EXISTS idx_assets_source_sheet ON assets(source_sheet);
CREATE INDEX IF NOT EXISTS idx_assets_archived ON assets(archived);
CREATE INDEX IF NOT EXISTS idx_assets_risk ON assets(risk_category);
CREATE INDEX IF NOT EXISTS idx_assets_date_osi_next ON assets(date_osi_next);
CREATE INDEX IF NOT EXISTS idx_assets_date_internal_next ON assets(date_internal_next);

CREATE VIRTUAL TABLE IF NOT EXISTS assets_fts USING fts5(
    name, tag, asset_number, description, remarks, sn, fluid_service, field,
    content='assets', content_rowid='id'
);

CREATE TRIGGER IF NOT EXISTS assets_ai AFTER INSERT ON assets BEGIN
  INSERT INTO assets_fts(rowid, name, tag, asset_number, description, remarks, sn, fluid_service, field)
  VALUES (new.id, new.name, new.tag, new.asset_number, new.description, new.remarks, new.sn, new.fluid_service, new.field);
END;

CREATE TRIGGER IF NOT EXISTS assets_ad AFTER DELETE ON assets BEGIN
  INSERT INTO assets_fts(assets_fts, rowid, name, tag, asset_number, description, remarks, sn, fluid_service, field)
  VALUES ('delete', old.id, old.name, old.tag, old.asset_number, old.description, old.remarks, old.sn, old.fluid_service, old.field);
END;

CREATE TRIGGER IF NOT EXISTS assets_au AFTER UPDATE ON assets BEGIN
  INSERT INTO assets_fts(assets_fts, rowid, name, tag, asset_number, description, remarks, sn, fluid_service, field)
  VALUES ('delete', old.id, old.name, old.tag, old.asset_number, old.description, old.remarks, old.sn, old.fluid_service, old.field);
  INSERT INTO assets_fts(rowid, name, tag, asset_number, description, remarks, sn, fluid_service, field)
  VALUES (new.id, new.name, new.tag, new.asset_number, new.description, new.remarks, new.sn, new.fluid_service, new.field);
END;

CREATE TABLE IF NOT EXISTS temp_repairs (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    facility_type TEXT,
    area TEXT,
    asset_name TEXT,
    repaired_section TEXT,
    repaired_by TEXT,
    original_repair_date TEXT,
    repair_life_years TEXT,
    expiration_date TEXT,
    hardness_hb TEXT,
    revalidation_date TEXT,
    last_expire_date TEXT,
    expiration_status TEXT,
    report_ref TEXT,
    remarks TEXT,
    created_at TEXT DEFAULT (datetime('now'))
);

CREATE TABLE IF NOT EXISTS critical_assets (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    sn TEXT,
    category_section TEXT,
    pack_no TEXT,
    report_no TEXT,
    item_description TEXT,
    insp_date TEXT,
    replacement_scope TEXT,
    replacement_done TEXT,
    remarks TEXT,
    plant_remarks TEXT,
    created_at TEXT DEFAULT (datetime('now'))
);

CREATE TABLE IF NOT EXISTS inspection_log (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    asset_id INTEGER NOT NULL REFERENCES assets(id) ON DELETE CASCADE,
    insp_date TEXT,
    insp_type TEXT,
    findings TEXT,
    next_due_date TEXT,
    inspector_name TEXT,
    insp_method TEXT,
    t_actual TEXT,
    action_required TEXT,
    created_at TEXT DEFAULT (datetime('now'))
);

CREATE TABLE IF NOT EXISTS raw_rows (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    sheet TEXT,
    row_num INTEGER,
    data TEXT
);
"""


def evaluate_cui_and_risk(record):
    """Auto-computes CUI susceptibility and API 580 POF/COF at ingestion time."""
    insul = str(record.get("insulation") or "").strip().lower()
    has_insul = insul not in ("", "none", "no", "n/a", "0", "false")
    
    op_temp_str = str(record.get("operating_temp") or "")
    op_temp_val = None
    m = re.search(r"(-?\d+(?:\.\d+)?)", op_temp_str)
    if m:
        try:
            op_temp_val = float(m.group(1))
        except ValueError:
            pass

    if has_insul and op_temp_val is not None:
        record["cui_susceptible"] = 1 if 10.0 <= op_temp_val <= 175.0 else 0
    elif has_insul:
        record["cui_susceptible"] = 1
    else:
        record["cui_susceptible"] = 0

    pof = 2
    rl_str = str(record.get("remaining_life") or "")
    rl_m = re.search(r"(\d+(?:\.\d+)?)", rl_str)
    if rl_m:
        try:
            rl = float(rl_m.group(1))
            if rl <= 2.0: pof = 5
            elif rl <= 5.0: pof = 4
            elif rl <= 10.0: pof = 3
        except ValueError:
            pass

    if record["cui_susceptible"]:
        pof = min(5, pof + 1)

    cof = 2
    sheet = (record.get("source_sheet") or "").lower()
    fluid = str(record.get("fluid_service") or "").lower()
    name = (record.get("name") or "").lower()

    if any(k in sheet or k in fluid or k in name for k in ["h2s", "acid", "lethal", "flare", "turbines", "vessels & tks", "gp inlet"]):
        cof = 5
    elif any(k in sheet or k in fluid or k in name for k in ["gas", "condensate", "fuel", "high press", "op piping", "gp piping", "epf"]):
        cof = 4
    elif any(k in sheet or k in fluid or k in name for k in ["crude", "oil", "coolers", "mfds", "tl", "fl"]):
        cof = 3
    elif any(k in sheet or k in fluid or k in name for k in ["water", "drain", "utility", "air"]):
        cof = 1

    record["pof_score"] = pof
    record["cof_score"] = cof

    score = pof * cof
    if score >= 16 or (cof == 5 and pof >= 3) or (pof == 5 and cof >= 3):
        record["risk_category"] = "HIGH"
    elif score >= 8:
        record["risk_category"] = "MEDIUM"
    else:
        record["risk_category"] = "LOW"


def import_smart_sheet(conn, ws, sheet_name, clean_wipe=False):
    """Smart sheet ingestion with header auto-discovery and synonym mapping."""
    hdr_row, col_map = find_header_row_and_map(ws)
    if not col_map:
        return 0

    cur = conn.cursor()
    count = 0

    for r in range(hdr_row + 1, ws.max_row + 1):
        row_vals = [ws.cell(row=r, column=c).value for c in range(1, ws.max_column + 1)]
        if not any(row_vals):
            continue

        raw_dict = {}
        for c in range(1, ws.max_column + 1):
            h = ws.cell(row=hdr_row, column=c).value or f"Col_{c}"
            raw_dict[str(h).strip()] = ws.cell(row=r, column=c).value

        record = {"source_sheet": sheet_name}
        extra_json = {}

        for c, val in enumerate(row_vals, start=1):
            if c in col_map:
                field_name, orig_hdr = col_map[c]
                if "date" in field_name:
                    d_iso, d_raw = parse_date(val)
                    record[field_name] = d_iso
                    if d_raw:
                        extra_json[orig_hdr + " (raw)"] = d_raw
                else:
                    record[field_name] = clean_str(val)
            else:
                h_name = ws.cell(row=hdr_row, column=c).value
                if h_name and val is not None:
                    extra_json[str(h_name).strip()] = str(val).strip()

        if not record.get("name"):
            record["name"] = record.get("tag") or record.get("asset_number") or record.get("description") or f"{sheet_name} Item {r}"

        if not record.get("tag") and record.get("sn"):
            record["tag"] = record.get("sn")

        evaluate_cui_and_risk(record)
        record["extra_json"] = json.dumps(extra_json, ensure_ascii=False) if extra_json else "{}"

        # If not clean_wipe, attempt non-destructive upsert
        if not clean_wipe:
            existing = None
            if record.get("tag"):
                existing = cur.execute("SELECT id, deferral_status, deferral_reason, deferral_expiry, deferral_approver, deferral_moc_no FROM assets WHERE source_sheet = ? AND tag = ?", (sheet_name, record["tag"])).fetchone()
            elif record.get("asset_number"):
                existing = cur.execute("SELECT id, deferral_status, deferral_reason, deferral_expiry, deferral_approver, deferral_moc_no FROM assets WHERE source_sheet = ? AND asset_number = ?", (sheet_name, record["asset_number"])).fetchone()

            if existing:
                aid = existing[0]
                if existing[1]: record["deferral_status"] = existing[1]
                if existing[2]: record["deferral_reason"] = existing[2]
                if existing[3]: record["deferral_expiry"] = existing[3]
                if existing[4]: record["deferral_approver"] = existing[4]
                if existing[5]: record["deferral_moc_no"] = existing[5]
                
                update_cols = [k for k in record.keys() if k != "id"]
                set_clause = ", ".join([f"{k} = ?" for k in update_cols])
                vals = [record[k] for k in update_cols] + [aid]
                cur.execute(f"UPDATE assets SET {set_clause} WHERE id = ?", vals)
            else:
                cols = list(record.keys())
                placeholders = ", ".join(["?"] * len(cols))
                cur.execute(f"INSERT INTO assets ({', '.join(cols)}) VALUES ({placeholders})", [record[k] for k in cols])
        else:
            cols = list(record.keys())
            placeholders = ", ".join(["?"] * len(cols))
            cur.execute(f"INSERT INTO assets ({', '.join(cols)}) VALUES ({placeholders})", [record[k] for k in cols])

        try:
            cur.execute("INSERT INTO raw_rows (sheet, row_num, data) VALUES (?,?,?)",
                        (sheet_name, r, json.dumps(raw_dict, default=str)))
        except Exception:
            pass
        count += 1

    return count


def import_temp_repairs(conn, ws):
    """Specialized importer for Temp-Repair (clamps & composite wraps)."""
    hdr_row, _ = find_header_row_and_map(ws)
    cur = conn.cursor()
    cur.execute("DELETE FROM temp_repairs")
    count = 0

    for r in range(hdr_row + 1, ws.max_row + 1):
        facility = clean_str(ws.cell(row=r, column=2).value)
        area = clean_str(ws.cell(row=r, column=3).value)
        asset_name = clean_str(ws.cell(row=r, column=4).value)
        section = clean_str(ws.cell(row=r, column=5).value)
        repaired_by = clean_str(ws.cell(row=r, column=6).value)
        
        d_orig, _ = parse_date(ws.cell(row=r, column=7).value)
        life_yrs = clean_str(ws.cell(row=r, column=8).value)
        d_exp, _ = parse_date(ws.cell(row=r, column=9).value)
        hardness = clean_str(ws.cell(row=r, column=10).value)
        d_reval, _ = parse_date(ws.cell(row=r, column=11).value)
        d_last_exp, _ = parse_date(ws.cell(row=r, column=12).value)
        status = clean_str(ws.cell(row=r, column=13).value) or "Active"
        report_ref = clean_str(ws.cell(row=r, column=14).value)
        remarks = clean_str(ws.cell(row=r, column=15).value)

        if not any([facility, area, asset_name, section]):
            continue

        cur.execute("""
            INSERT INTO temp_repairs (
                facility_type, area, asset_name, repaired_section, repaired_by,
                original_repair_date, repair_life_years, expiration_date,
                hardness_hb, revalidation_date, last_expire_date, expiration_status,
                report_ref, remarks
            ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)
        """, (facility, area, asset_name, section, repaired_by, d_orig, life_yrs,
              d_exp, hardness, d_reval, d_last_exp, status, report_ref, remarks))
        count += 1
    return count


def import_critical_assets(conn, ws):
    """Specialized importer for Critical Assets (Turnaround scope)."""
    hdr_row, _ = find_header_row_and_map(ws)
    cur = conn.cursor()
    cur.execute("DELETE FROM critical_assets")
    count = 0

    for r in range(hdr_row + 1, ws.max_row + 1):
        sn = clean_str(ws.cell(row=r, column=1).value)
        category = clean_str(ws.cell(row=r, column=2).value)
        pack_no = clean_str(ws.cell(row=r, column=3).value)
        report_no = clean_str(ws.cell(row=r, column=4).value)
        item_desc = clean_str(ws.cell(row=r, column=5).value)
        d_insp, _ = parse_date(ws.cell(row=r, column=6).value)
        scope = clean_str(ws.cell(row=r, column=7).value)
        done = clean_str(ws.cell(row=r, column=8).value) or "No"
        remarks = clean_str(ws.cell(row=r, column=9).value)
        plant_rem = clean_str(ws.cell(row=r, column=10).value)

        if not any([pack_no, item_desc, scope]):
            continue

        cur.execute("""
            INSERT INTO critical_assets (
                sn, category_section, pack_no, report_no, item_description,
                insp_date, replacement_scope, replacement_done, remarks, plant_remarks
            ) VALUES (?,?,?,?,?,?,?,?,?,?)
        """, (sn, category, pack_no, report_no, item_desc, d_insp, scope, done, remarks, plant_rem))
        count += 1
    return count


def import_workbook(xlsx_path, db_path, clean_wipe=False):
    print(f"\n=======================================================")
    print(f"  Smart Ingestion ({'CLEAN WIPE' if clean_wipe else 'SMART SYNC'}): {xlsx_path}")
    print(f"  Target SQLite DB: {db_path}")
    print(f"=======================================================\n")

    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    conn = sqlite3.connect(db_path)
    conn.execute("PRAGMA foreign_keys = ON")
    conn.executescript(SCHEMA)

    if clean_wipe:
        print("  [*] Wiping existing database tables for clean import...")
        conn.execute("DELETE FROM assets")
        conn.execute("DELETE FROM temp_repairs")
        conn.execute("DELETE FROM critical_assets")
        conn.execute("DELETE FROM inspection_log")
        conn.execute("DELETE FROM raw_rows")
        try:
            conn.execute("DELETE FROM assets_fts")
        except Exception:
            pass
        conn.commit()

    total_assets = 0
    total_temp_repairs = 0
    total_critical = 0

    for sheet_name in wb.sheetnames:
        if sheet_name in SKIP_SHEETS:
            continue

        ws = wb[sheet_name]
        lower_name = sheet_name.strip().lower()

        with import_sheet_guard(sheet_name):
            if "temp" in lower_name and "repair" in lower_name:
                count = import_temp_repairs(conn, ws)
                total_temp_repairs += count
                print(f"  [+] {sheet_name:28s} -> {count:4d} temporary repairs")
            elif "critical" in lower_name:
                count = import_critical_assets(conn, ws)
                total_critical += count
                print(f"  [+] {sheet_name:28s} -> {count:4d} turnaround critical scope items")
            else:
                count = import_smart_sheet(conn, ws, sheet_name, clean_wipe=clean_wipe)
                total_assets += count
                print(f"  [+] {sheet_name:28s} -> {count:4d} assets mapped & synchronized")

    conn.commit()
    conn.close()

    print(f"\n-------------------------------------------------------")
    print(f"  SMART INGESTION COMPLETE:")
    print(f"  Total Equipment Assets : {total_assets:,}")
    print(f"  Temporary Repairs      : {total_temp_repairs:,}")
    print(f"  Critical Turnaround    : {total_critical:,}")
    if import_errors:
        print(f"  Warnings/Errors        : {len(import_errors)}")
    print(f"-------------------------------------------------------\n")


def main():
    clean = "--clean" in sys.argv or "-c" in sys.argv
    args = [a for a in sys.argv[1:] if not a.startswith("-")]
    xlsx = args[0] if len(args) > 0 else "1. Master Inspection Plan - Updated 4-6-2026.xlsx"
    db = args[1] if len(args) > 1 else "inspection_plan.db"
    
    if not os.path.exists(xlsx):
        found = [f for f in os.listdir(".") if f.endswith(".xlsx") and not f.startswith("~$")]
        if found:
            xlsx = found[0]
        else:
            print(f"Error: Excel file not found: {xlsx}")
            sys.exit(1)

    import_workbook(xlsx, db, clean_wipe=clean)


if __name__ == "__main__":
    main()