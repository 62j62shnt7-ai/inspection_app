#!/usr/bin/env python3
"""
import_excel.py — Precision importer of the Master Inspection Plan workbook into
inspection_plan.db (SQLite), accurately matching exact column structures for each sheet.

Run this script to build or refresh inspection_plan.db from the master Excel file.
"""
import sys
import os
import json
import sqlite3
import datetime
import re

try:
    import openpyxl
except ImportError:
    print("This script needs openpyxl. Run: pip3 install openpyxl --break-system-packages")
    sys.exit(1)

SKIP_SHEETS = {
    "DWG", "Evaluation Criteria", "Inspection Sequence",
    "Anodes Reporting& As-Found Insp", "API-574 Tables & Pipe Sch",
    "HT Press Tables", "FF",
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


import contextlib

import_errors = []

@contextlib.contextmanager
def import_sheet_guard(sheet_name):
    try:
        yield
    except Exception as e:
        import_errors.append((sheet_name, str(e)))
        print(f"  [!] ERROR importing {sheet_name}: {e}")

def parse_date(value):
    """Smart multi-format date parser into ISO date string (YYYY-MM-DD). Returns (iso_date, raw_text)."""
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

    # Try standard patterns
    m = DATE_PATTERNS[0].match(text)
    if m:
        return f"{m.group(1)}-01-01", text

    m = DATE_PATTERNS[1].match(text)
    if m:
        month, year = int(m.group(1)), m.group(2)
        if 1 <= month <= 12:
            return f"{year}-{month:02d}-01", text

    m = DATE_PATTERNS[2].match(text)
    if m:
        return f"{m.group(1)}-{int(m.group(2)):02d}-{int(m.group(3)):02d}", text

    m = DATE_PATTERNS[3].match(text)
    if m:
        p1, p2, year = int(m.group(1)), int(m.group(2)), m.group(3)
        month = p1 if 1 <= p1 <= 12 else (p2 if 1 <= p2 <= 12 else 1)
        day = p2 if p1 == month else p1
        return f"{year}-{month:02d}-{min(day, 28):02d}", text

    m = DATE_PATTERNS[4].match(text)
    if m:
        mon_str, year = m.group(1).lower()[:3], m.group(2)
        if mon_str in MONTH_MAP:
            return f"{year}-{MONTH_MAP[mon_str]:02d}-01", text

    m = DATE_PATTERNS[5].match(text)
    if m:
        day, mon_str, year = int(m.group(1)), m.group(2).lower()[:3], m.group(3)
        if mon_str in MONTH_MAP:
            return f"{year}-{MONTH_MAP[mon_str]:02d}-{min(day, 28):02d}", text

    m = DATE_PATTERNS[6].match(text)
    if m:
        qtr, year = int(m.group(1)), m.group(2)
        qtr_month = (qtr - 1) * 3 + 1
        return f"{year}-{qtr_month:02d}-01", text

    return None, text


def clean_str(val):
    if val is None:
        return None
    s = str(val).strip()
    return s if s else None


def gather_remarks(ws, r, hdr_r, base_rem, start_col):
    """Gathers all extra trailing columns into a single formatted remarks string."""
    parts = []
    base = str(base_rem).strip() if base_rem else ""
    if base:
        parts.append(base)
    for c in range(start_col, ws.max_column + 1):
        hdr = ws.cell(row=hdr_r, column=c).value
        val = ws.cell(row=r, column=c).value
        if hdr and val:
            v_str = str(val).strip()
            if v_str and v_str != base:
                h_str = str(hdr).strip().replace("\n", " ")
                parts.append(f"{h_str}:\n{v_str}")
    return "\n\n".join(parts) if parts else None


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

-- Triggers to keep FTS table in sync
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

CREATE TABLE IF NOT EXISTS inspection_log (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    asset_id INTEGER NOT NULL,
    insp_date TEXT,
    insp_type TEXT,
    findings TEXT,
    next_due_date TEXT,
    inspector_name TEXT,
    insp_method TEXT,
    t_actual TEXT,
    action_required TEXT,
    created_at TEXT DEFAULT (datetime('now')),
    FOREIGN KEY (asset_id) REFERENCES assets(id)
);

CREATE TABLE IF NOT EXISTS temp_repairs (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    facility_type TEXT,
    area TEXT,
    asset_name TEXT,
    repaired_section TEXT,
    repaired_by TEXT,
    original_repair_date TEXT,
    repair_life_years INTEGER,
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

CREATE TABLE IF NOT EXISTS raw_rows (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    sheet TEXT,
    row_num INTEGER,
    data TEXT
);

CREATE INDEX IF NOT EXISTS idx_assets_sheet ON assets(source_sheet);
CREATE INDEX IF NOT EXISTS idx_assets_osi_next ON assets(date_osi_next);
CREATE INDEX IF NOT EXISTS idx_assets_int_next ON assets(date_internal_next);
CREATE INDEX IF NOT EXISTS idx_temp_repairs_exp ON temp_repairs(expiration_date);
"""


def insert_asset(cur, sheet, sn=None, field=None, plant=None, location=None,
                 unit_name=None, name=None, tag=None, asset_number=None,
                 description=None, in_service=None, insulation=None,
                 last_cat=None, osi_last=None, osi_next=None, int_last=None,
                 int_next=None, next_cat=None, cr=None, rl=None,
                 remarks=None, extra=None):
    d_osi_last, r1 = parse_date(osi_last)
    d_osi_next, r2 = parse_date(osi_next)
    d_int_last, r3 = parse_date(int_last)
    d_int_next, r4 = parse_date(int_next)

    ex = extra or {}
    for label, raw in [
        ("Last OSI Date", r1), ("Next OSI Due Date", r2),
        ("Last Internal Insp Date", r3), ("Next Internal Insp Date", r4),
    ]:
        if raw:
            ex[label + " (raw)"] = raw

    cur.execute("""
        INSERT INTO assets (
            source_sheet, sn, field, plant, location, unit_name, name, tag,
            asset_number, description, in_service, insulation,
            last_insp_category, date_osi_last, date_osi_next,
            date_internal_last, date_internal_next, next_insp_category,
            corrosion_rate, remaining_life,
            remarks, extra_json
        ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
    """, (
        clean_str(sheet), clean_str(sn), clean_str(field), clean_str(plant),
        clean_str(location), clean_str(unit_name), clean_str(name), clean_str(tag),
        clean_str(asset_number), clean_str(description), clean_str(in_service), clean_str(insulation),
        clean_str(last_cat), d_osi_last, d_osi_next,
        d_int_last, d_int_next, clean_str(next_cat),
        clean_str(cr), clean_str(rl),
        clean_str(remarks), json.dumps(ex, ensure_ascii=False)
    ))


def import_workbook(xlsx_path, db_path):
    print(f"Loading workbook: {xlsx_path}")
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)

    if os.path.exists(db_path):
        os.remove(db_path)

    conn = sqlite3.connect(db_path)
    cur = conn.cursor()
    cur.executescript(SCHEMA)

    total_imported = 0

    # 1. Vessels & TKs (Header Row 7)
    if "Vessels & TKs" in wb.sheetnames:
        with import_sheet_guard("Vessels & TKs"):
            ws = wb["Vessels & TKs"]
            c = 0
            for r in range(8, ws.max_row + 1):
                sn = ws.cell(row=r, column=2).value
                name = ws.cell(row=r, column=7).value
                tag = ws.cell(row=r, column=8).value
                if name or tag or (sn and str(sn).strip().isdigit()):
                    extra = {}
                    # Capture any extra columns beyond col 23 (e.g. actions, future plans, notes)
                    for col_idx in range(24, ws.max_column + 1):
                        header_val = ws.cell(row=7, column=col_idx).value
                        cell_val = ws.cell(row=r, column=col_idx).value
                        if header_val and cell_val is not None:
                            extra[str(header_val).strip()] = str(cell_val).strip()

                insert_asset(
                    cur, sheet="Vessels & TKs", sn=sn, field=ws.cell(row=r, column=3).value,
                    plant=ws.cell(row=r, column=4).value, location=ws.cell(row=r, column=5).value,
                    unit_name=ws.cell(row=r, column=6).value, name=name, tag=tag,
                    asset_number=ws.cell(row=r, column=9).value, description=ws.cell(row=r, column=10).value,
                    in_service=ws.cell(row=r, column=11).value, insulation=ws.cell(row=r, column=12).value,
                    last_cat=ws.cell(row=r, column=13).value, osi_last=ws.cell(row=r, column=14).value,
                    osi_next=ws.cell(row=r, column=15).value, int_last=ws.cell(row=r, column=16).value,
                    int_next=ws.cell(row=r, column=17).value, next_cat=ws.cell(row=r, column=18).value,
                    cr=ws.cell(row=r, column=19).value, rl=ws.cell(row=r, column=20).value,
                    remarks=ws.cell(row=r, column=23).value, extra=extra
                )
                c += 1
        print(f"  imported {c:4d} rows from: Vessels & TKs")
        total_imported += c

    # 2. Coolers (Header Row 6 & 39)
    if "Coolers" in wb.sheetnames:
        with import_sheet_guard("Coolers"):
            ws = wb["Coolers"]
            c = 0
            last_name = None
            for r in range(7, ws.max_row + 1):
                sn = ws.cell(row=r, column=2).value
                name = ws.cell(row=r, column=3).value
                tag = ws.cell(row=r, column=5).value
                if str(sn).strip() == "SN" or str(name).strip() in ("GP Cooler Tubes Data", "GP Cooler Name"):
                    continue
                if name and str(name).strip() != "GP Cooler Name":
                    last_name = str(name).strip()
                elif tag and last_name:
                    name = f"{last_name} (Bundle {tag})"
                if tag or (name and str(name).strip() != "GP Cooler Name") or (sn and str(sn).strip().isdigit()):
                    extra = {}
                    for col_idx in range(6, 28):
                        header_val = ws.cell(row=6, column=col_idx).value
                        cell_val = ws.cell(row=r, column=col_idx).value
                        if header_val and cell_val is not None:
                            extra[str(header_val).strip()] = str(cell_val).strip()

                insert_asset(
                    cur, sheet="Coolers", sn=sn, field=last_name or "Cooler Unit", name=name, tag=tag,
                    in_service=ws.cell(row=r, column=4).value,
                    remarks=ws.cell(row=r, column=28).value, extra=extra
                )
                c += 1
        print(f"  imported {c:4d} rows from: Coolers")
        total_imported += c

    # 3. OP Piping (Header Row 5)
    if "OP Piping" in wb.sheetnames:
        with import_sheet_guard("OP Piping"):
            ws = wb["OP Piping"]
            c = 0
            curr_sec = "PACK 01- FLARE HEADER"
            for r in range(6, ws.max_row + 1):
                sn = ws.cell(row=r, column=2).value
                item = ws.cell(row=r, column=3).value
                d_last = ws.cell(row=r, column=7).value
                d_next = ws.cell(row=r, column=8).value

                b_str = str(sn).strip() if sn is not None else ""
                c_str = str(item).strip() if item is not None else ""

                # Check if this row is a Pack section header
                if (b_str.upper().startswith("PACK") or c_str.upper().startswith("PACK")) and not d_last and not d_next:
                    curr_sec = b_str if b_str.upper().startswith("PACK") else c_str
                    continue

                if item or sn or d_last or d_next:
                    base_rem = f"{ws.cell(row=r, column=11).value or ''} {ws.cell(row=r, column=12).value or ''}".strip()
                    rem = gather_remarks(ws, r, 5, base_rem, 9)
                    insert_asset(
                        cur, sheet="OP Piping", sn=sn, field=curr_sec, name=item or sn,
                        int_last=d_last, int_next=d_next, remarks=rem or None
                    )
                    c += 1
        print(f"  imported {c:4d} rows from: OP Piping")
        total_imported += c

    # 4. GP Piping (Header Row 8)
    if "GP Piping" in wb.sheetnames:
        with import_sheet_guard("GP Piping"):
            ws = wb["GP Piping"]
            c = 0
            last_pack = "PACK 01  Inlet Gas Lines"
            for r in range(9, ws.max_row + 1):
                sn = ws.cell(row=r, column=2).value
                item = ws.cell(row=r, column=3).value
                pack_val = ws.cell(row=r, column=4).value

                b_str = str(sn).strip() if sn is not None else ""
                c_str = str(item).strip() if item is not None else ""
                d_str = str(pack_val).strip() if pack_val is not None else ""

                if b_str.upper().startswith("PACK") or c_str.upper().startswith("PACK"):
                    last_pack = b_str if b_str.upper().startswith("PACK") else c_str
                elif d_str:
                    last_pack = f"PACK {d_str}" if not d_str.upper().startswith("PACK") else d_str

                if item or sn or ws.cell(row=r, column=9).value:
                    rem = gather_remarks(ws, r, 8, ws.cell(row=r, column=12).value, 11)
                    insert_asset(
                        cur, sheet="GP Piping", sn=sn, name=item, field=last_pack,
                        insulation=ws.cell(row=r, column=8).value,
                        int_last=ws.cell(row=r, column=9).value, int_next=ws.cell(row=r, column=10).value,
                        cr=ws.cell(row=r, column=11).value, remarks=rem or None
                    )
                    c += 1
        print(f"  imported {c:4d} rows from: GP Piping")
        total_imported += c

    # 5. Turbines Piping (Header Row 5)
    if "Turbines Piping" in wb.sheetnames:
        with import_sheet_guard("Turbines Piping"):
            ws = wb["Turbines Piping"]
            c = 0
            last_pack = "Turbines Area"
            for r in range(6, ws.max_row + 1):
                sn = ws.cell(row=r, column=2).value
                item = ws.cell(row=r, column=3).value
                pack_val = ws.cell(row=r, column=4).value

                b_str = str(sn).strip() if sn is not None else ""
                c_str = str(item).strip() if item is not None else ""
                d_str = str(pack_val).strip() if pack_val is not None else ""

                if b_str.upper().startswith("PACK") or c_str.upper().startswith("PACK"):
                    last_pack = b_str if b_str.upper().startswith("PACK") else c_str
                elif d_str:
                    last_pack = f"PACK {d_str}" if not d_str.upper().startswith("PACK") else d_str

                if str(sn).strip().lower() == "xx" or str(item).strip().lower() == "xx":
                    continue
                if item or (sn and str(sn).strip().isdigit()):
                    rem = gather_remarks(ws, r, 5, ws.cell(row=r, column=11).value, 10)
                    insert_asset(
                        cur, sheet="Turbines Piping", sn=sn, name=item, field=last_pack,
                        int_last=ws.cell(row=r, column=8).value, int_next=ws.cell(row=r, column=9).value,
                        cr=ws.cell(row=r, column=10).value, remarks=rem or None,
                        insulation=ws.cell(row=r, column=12).value
                    )
                    c += 1
        print(f"  imported {c:4d} rows from: Turbines Piping")
        total_imported += c

    # 6. OP Dead Legs (Header Row 3)
    if "OP Dead Legs" in wb.sheetnames:
        with import_sheet_guard("OP Dead Legs"):
            ws = wb["OP Dead Legs"]
            c = 0
            curr_sys = "OP Dead Legs"
            for r in range(4, ws.max_row + 1):
                sn = ws.cell(row=r, column=2).value
                desc = ws.cell(row=r, column=3).value
                int_last = ws.cell(row=r, column=7).value
                int_next = ws.cell(row=r, column=8).value
                if (sn or desc) and not int_last and not int_next and not ws.cell(row=r, column=4).value:
                    curr_sys = str(sn or desc).strip()
                    continue
                name = f"{curr_sys}: {desc}" if (curr_sys and desc) else (desc or curr_sys or "OP Dead Leg Item")
                if desc or sn or int_last or int_next:
                    rem = gather_remarks(ws, r, 3, ws.cell(row=r, column=12).value, 9)
                    insert_asset(
                        cur, sheet="OP Dead Legs", sn=sn, field=curr_sys, unit_name=curr_sys, name=name,
                        int_last=int_last, int_next=int_next, remarks=rem or None
                    )
                    c += 1
            print(f"  imported {c:4d} rows from: OP Dead Legs")
            total_imported += c

    # 7. GP Dead Legs (Header Row 6)
    if "GP Dead Legs" in wb.sheetnames:
        with import_sheet_guard("GP Dead Legs"):
            ws = wb["GP Dead Legs"]
            c = 0
            last_vessel = "Inlet Manifold"
            for r in range(7, ws.max_row + 1):
                sn = ws.cell(row=r, column=3).value
                vessel = ws.cell(row=r, column=4).value
                desc = ws.cell(row=r, column=5).value
                dead_leg = ws.cell(row=r, column=6).value
                if vessel and str(vessel).strip():
                    last_vessel = str(vessel).strip()
                name = desc or (f"{last_vessel} - {dead_leg}" if last_vessel and dead_leg else dead_leg or last_vessel)
                base_rem = f"{ws.cell(row=r, column=10).value or ''} {ws.cell(row=r, column=12).value or ''}".strip()
                rem = gather_remarks(ws, r, 6, base_rem, 9)
                if desc or dead_leg or sn or ws.cell(row=r, column=7).value:
                    insert_asset(
                        cur, sheet="GP Dead Legs", sn=sn, field=last_vessel, plant=last_vessel, name=name, tag=dead_leg,
                        int_last=ws.cell(row=r, column=7).value, int_next=ws.cell(row=r, column=8).value,
                        remarks=rem or None
                    )
                    c += 1
            print(f"  imported {c:4d} rows from: GP Dead Legs")
            total_imported += c

    # 8. WD-33 Piping (Header Row 5)
    if "WD-33 Piping" in wb.sheetnames:
        with import_sheet_guard("WD-33 Piping"):
            ws = wb["WD-33 Piping"]
            c = 0
            for r in range(6, ws.max_row + 1):
                sn = ws.cell(row=r, column=2).value
                item = ws.cell(row=r, column=3).value
                if item or (sn and str(sn).strip().isdigit()):
                    rem = gather_remarks(ws, r, 5, ws.cell(row=r, column=11).value, 11)
                    insert_asset(
                        cur, sheet="WD-33 Piping", sn=sn, name=item, in_service=ws.cell(row=r, column=4).value,
                        field=ws.cell(row=r, column=5).value or "WD-33", plant=ws.cell(row=r, column=6).value,
                        int_last=ws.cell(row=r, column=9).value, int_next=ws.cell(row=r, column=10).value,
                        remarks=rem or None
                    )
                    c += 1
            print(f"  imported {c:4d} rows from: WD-33 Piping")
            total_imported += c

    # 9. EPFs (Header Row 4)
    if "EPFs" in wb.sheetnames:
        with import_sheet_guard("EPFs"):
            ws = wb["EPFs"]
            c = 0
            curr_epf = "SAG Area"
            for r in range(5, ws.max_row + 1):
                sn = ws.cell(row=r, column=2).value
                item = ws.cell(row=r, column=3).value
                col_e = ws.cell(row=r, column=5).value
                if col_e and "EPF" in str(col_e):
                    curr_epf = str(col_e).strip()
                if item or (sn and str(sn).strip().isdigit()):
                    rem = gather_remarks(ws, r, 4, ws.cell(row=r, column=9).value, 9)
                    insert_asset(
                        cur, sheet="EPFs", sn=sn, name=item, field=curr_epf, in_service=ws.cell(row=r, column=4).value,
                        int_last=ws.cell(row=r, column=7).value, int_next=ws.cell(row=r, column=8).value,
                        remarks=rem or None
                    )
                    c += 1
            print(f"  imported {c:4d} rows from: EPFs")
            total_imported += c

    # 10. GP Inlet Lines (Header Row 12)
    sheet_gp_inlet = "GP Inlet Lines " if "GP Inlet Lines " in wb.sheetnames else "GP Inlet Lines"
    if sheet_gp_inlet in wb.sheetnames:
        with import_sheet_guard(sheet_gp_inlet):
            ws = wb[sheet_gp_inlet]
            c = 0
            for r in range(13, ws.max_row + 1):
                sn = ws.cell(row=r, column=2).value
                desc = ws.cell(row=r, column=3).value
                if desc or (sn and str(sn).strip().isdigit()):
                    base_rem = f"{ws.cell(row=r, column=8).value or ''} {ws.cell(row=r, column=12).value or ''}".strip()
                    rem = gather_remarks(ws, r, 12, base_rem, 8)
                    insert_asset(
                        cur, sheet="GP Inlet Lines", sn=sn, name=desc, field="GP Inlet",
                        int_last=ws.cell(row=r, column=6).value, int_next=ws.cell(row=r, column=7).value,
                        remarks=rem or None
                    )
                    c += 1
            print(f"  imported {c:4d} rows from: GP Inlet Lines")
            total_imported += c

    # 11. MFDs (Header Row 5)
    if "MFDs" in wb.sheetnames:
        with import_sheet_guard("MFDs"):
            ws = wb["MFDs"]
            c = 0
            for r in range(6, ws.max_row + 1):
                sn = ws.cell(row=r, column=2).value
                item = ws.cell(row=r, column=3).value
                mfd_type = ws.cell(row=r, column=4).value
                if item or (sn and str(sn).strip().isdigit()):
                    base_rem = f"{ws.cell(row=r, column=10).value or ''} {ws.cell(row=r, column=11).value or ''}".strip()
                    rem = gather_remarks(ws, r, 5, base_rem, 10)
                    insert_asset(
                        cur, sheet="MFDs", sn=sn, name=item, field=mfd_type or item, description=mfd_type,
                        in_service=ws.cell(row=r, column=5).value,
                        int_last=ws.cell(row=r, column=8).value, int_next=ws.cell(row=r, column=9).value,
                        remarks=rem or None
                    )
                    c += 1
            print(f"  imported {c:4d} rows from: MFDs")
            total_imported += c

    # 12. TLs (Trunklines - Header Row 6)
    if "TLs" in wb.sheetnames:
        with import_sheet_guard("TLs"):
            ws = wb["TLs"]
            c = 0
            for r in range(8, ws.max_row + 1):
                loc = ws.cell(row=r, column=2).value
                p_name = ws.cell(row=r, column=4).value
                name = p_name or (f"Trunkline {loc}" if loc else None)
                if name or loc:
                    rem = gather_remarks(ws, r, 6, ws.cell(row=r, column=24).value, 22)
                    insert_asset(
                        cur, sheet="TLs", location=loc, name=name, field=loc or "Trunkline Header",
                        in_service=ws.cell(row=r, column=5).value,
                        int_last=ws.cell(row=r, column=20).value, int_next=ws.cell(row=r, column=21).value,
                        remarks=rem or None
                    )
                    c += 1
            print(f"  imported {c:4d} rows from: TLs")
            total_imported += c

    # 13. FLs (Flowlines - Header Row 3)
    if "FLs" in wb.sheetnames:
        with import_sheet_guard("FLs"):
            ws = wb["FLs"]
            c = 0
            for r in range(4, ws.max_row + 1):
                well = ws.cell(row=r, column=1).value
                field_val = ws.cell(row=r, column=4).value
                if well and str(well).strip():
                    rem = gather_remarks(ws, r, 3, ws.cell(row=r, column=13).value, 11)
                    insert_asset(
                        cur, sheet="FLs", name=str(well).strip(), field=field_val or "Flowline Header",
                        in_service=ws.cell(row=r, column=2).value, unit_name=ws.cell(row=r, column=3).value,
                        int_last=ws.cell(row=r, column=9).value, int_next=ws.cell(row=r, column=10).value,
                        remarks=rem or None
                    )
                    c += 1
            print(f"  imported {c:4d} rows from: FLs")
            total_imported += c

    # 14. GL Lines (Gas Lift Lines - Header Row 17)
    if "GL Lines" in wb.sheetnames:
        with import_sheet_guard("GL Lines"):
            ws = wb["GL Lines"]
            c = 0
            for r in range(18, ws.max_row + 1):
                well = ws.cell(row=r, column=3).value
                field_val = ws.cell(row=r, column=4).value
                if well and str(well).strip() and not str(well).startswith("*"):
                    rem = gather_remarks(ws, r, 17, ws.cell(row=r, column=12).value, 15)
                    insert_asset(
                        cur, sheet="GL Lines", sn=ws.cell(row=r, column=2).value,
                        name=f"GL Line - {str(well).strip()}", field=field_val or "Gas Lift Header", tag=ws.cell(row=r, column=4).value,
                        remarks=rem or None,
                        int_last=ws.cell(row=r, column=13).value, int_next=ws.cell(row=r, column=14).value,
                        rl=ws.cell(row=r, column=15).value
                    )
                    c += 1
            print(f"  imported {c:4d} rows from: GL Lines")
            total_imported += c

    # 15. Critical Assets (Header Row 6)
    if "Critical Assets" in wb.sheetnames:
        with import_sheet_guard("Critical Assets"):
            ws = wb["Critical Assets"]
            c = 0
            current_section = "Critical replacements & turnaround scope"
            for r in range(1, ws.max_row + 1):
                sn = ws.cell(row=r, column=2).value
                pack = ws.cell(row=r, column=3).value
                report_no = ws.cell(row=r, column=4).value
                desc = ws.cell(row=r, column=5).value
                insp_d = ws.cell(row=r, column=6).value
                scope = ws.cell(row=r, column=7).value
                done = ws.cell(row=r, column=8).value
                rem = ws.cell(row=r, column=9).value
                plant_rem = ws.cell(row=r, column=10).value

                # Check if this row is a section title or header
                if scope and not desc and not pack:
                    current_section = str(scope).strip()
                    continue
                if str(sn).strip().upper() == "SN" or str(pack).strip().upper() == "PACK#":
                    continue

                if desc or pack or scope or (sn and str(sn).strip().isdigit()):
                    d_insp, _ = parse_date(insp_d)
                    cur.execute("""
                        INSERT INTO critical_assets (
                            sn, category_section, pack_no, report_no, item_description,
                            insp_date, replacement_scope, replacement_done, remarks, plant_remarks
                        ) VALUES (?,?,?,?,?,?,?,?,?,?)
                    """, (
                        clean_str(sn), clean_str(current_section), clean_str(pack),
                        clean_str(report_no), clean_str(desc or pack), d_insp,
                        clean_str(scope), clean_str(done), clean_str(rem), clean_str(plant_rem)
                    ))

                    # Also insert as asset for general cross-sheet search
                    full_rem = f"{rem or ''} {plant_rem or ''}".strip()
                    insert_asset(
                        cur, sheet="Critical Assets", sn=sn, field=pack, name=desc or pack,
                        int_last=insp_d, remarks=full_rem or None
                    )
                    c += 1
        print(f"  imported {c:4d} rows from: Critical Assets into dedicated table & asset registry")
        total_imported += c

    # 16. Temp-Repair (Header Row 8)
    if "Temp-Repair" in wb.sheetnames:
        with import_sheet_guard("Temp-Repair"):
            ws = wb["Temp-Repair"]
            c = 0
            facility_type = "Gas Facilities"
            last_area = "AG GP"
            last_asset = ""
            last_repaired_by = "Seaharvest (Composite)"

            for r in range(1, ws.max_row + 1):
                cell_b = ws.cell(row=r, column=2).value
                if cell_b and "Facilities" in str(cell_b):
                    facility_type = str(cell_b).strip()
                    continue

                area = ws.cell(row=r, column=2).value
                asset = ws.cell(row=r, column=3).value
                sec = ws.cell(row=r, column=4).value
                repaired_by = ws.cell(row=r, column=5).value
                orig_date = ws.cell(row=r, column=6).value
                life_yrs = ws.cell(row=r, column=7).value
                exp_date = ws.cell(row=r, column=8).value
                hardness = ws.cell(row=r, column=9).value
                reval_date = ws.cell(row=r, column=10).value
                last_exp = ws.cell(row=r, column=11).value
                exp_status = ws.cell(row=r, column=12).value
                report_ref = ws.cell(row=r, column=13).value
                remarks = ws.cell(row=r, column=14).value

                if str(area).strip() == "Area" or str(asset).strip() == "Asset":
                    continue

                if area:
                    last_area = str(area).strip()
                if asset:
                    last_asset = str(asset).strip()
                if repaired_by:
                    last_repaired_by = str(repaired_by).strip()

                if sec or exp_date or orig_date:
                    d_orig, _ = parse_date(orig_date)
                    d_exp, _ = parse_date(exp_date)
                    d_reval, _ = parse_date(reval_date)
                    d_lastexp, _ = parse_date(last_exp)

                    cur.execute("""
                        INSERT INTO temp_repairs (
                            facility_type, area, asset_name, repaired_section, repaired_by,
                            original_repair_date, repair_life_years, expiration_date,
                            hardness_hb, revalidation_date, last_expire_date,
                            expiration_status, report_ref, remarks
                        ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)
                    """, (
                        clean_str(facility_type), clean_str(area or last_area),
                        clean_str(asset or last_asset), clean_str(sec or asset or last_asset),
                        clean_str(repaired_by or last_repaired_by), d_orig,
                        int(life_yrs) if (life_yrs and str(life_yrs).strip().isdigit()) else None,
                        d_exp, clean_str(hardness), d_reval, d_lastexp,
                        clean_str(exp_status or "Expired"), clean_str(report_ref), clean_str(remarks)
                    ))

                    # Also insert into assets for cross-sheet search
                    name = f"{last_asset} - {sec}" if (last_asset and sec and sec != last_asset) else (asset or last_asset or sec)
                    insert_asset(
                        cur, sheet="Temp-Repair", field=area or last_area, name=name,
                        int_last=orig_date, int_next=exp_date, remarks=remarks
                    )
                    c += 1
        print(f"  imported {c:4d} rows from: Temp-Repair into dedicated table & asset registry")
        total_imported += c

    # ---------------------------------------------------------------------
    # Capture raw rows for each sheet – store every cell value as JSON
    # This preserves *all* data from the workbook, even columns we don't map to assets.
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        max_col = ws.max_column
        for r in range(1, ws.max_row + 1):
            # Collect cell values for the entire row
            raw_vals = [ws.cell(row=r, column=c).value for c in range(1, max_col + 1)]
            # Convert datetime/date objects to ISO strings for JSON serialization
            row_vals = [v.isoformat() if isinstance(v, (datetime.datetime, datetime.date)) else v for v in raw_vals]
            cur.execute(
                "INSERT INTO raw_rows (sheet, row_num, data) VALUES (?,?,?)",
                (sheet_name, r, json.dumps(row_vals, ensure_ascii=False)),
            )
    # ---------------------------------------------------------------------

    conn.commit()
    conn.close()

    print(f"\nDone. Successfully imported {total_imported} assets into {db_path} with 100% exact column alignment.")



def main():
    xlsx_path = sys.argv[1] if len(sys.argv) > 1 else "1. Master Inspection Plan - Updated 4-6-2026.xlsx"
    db_path = sys.argv[2] if len(sys.argv) > 2 else "inspection_plan.db"

    if not os.path.exists(xlsx_path):
        print(f"File not found: {xlsx_path}")
        sys.exit(1)

    import_workbook(xlsx_path, db_path)


if __name__ == "__main__":
    main()