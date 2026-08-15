#!/usr/bin/env python3
"""
import_excel.py — Precision Engineered Importer for Master Inspection Plan Excel workbooks.

Accurately maps every sheet's unique layout, columns, dates, envelope specs, and remarks
without column misalignments.
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
    "june": 6, "july": 7, "jul": 7, "aug": 8, "sep": 9, "sept": 9, "oct": 10, "nov": 11, "dec": 12
}

DATE_PATTERNS = [
    re.compile(r"^\s*(\d{4})\s*$"),                                      # 2024
    re.compile(r"^\s*(\d{1,2})[-/](\d{4})\s*$"),                          # 09/2024, 9-2024
    re.compile(r"^\s*(\d{4})[-/](\d{1,2})[-/](\d{1,2})\s*$"),            # 2024-09-15
    re.compile(r"^\s*(\d{1,2})[-/](\d{1,2})[-/](\d{4})\s*$"),            # 15/09/2024 or 09/15/2024
    re.compile(r"^\s*([A-Za-z]+)\s*[-/ ]\s*(\d{4})\s*$"),                # Sep-2024, June 2024
    re.compile(r"^\s*([A-Za-z]+)\s*[-/ ]\s*(\d{2})\s*$"),                # June -19, Jun-22
    re.compile(r"^\s*(\d{1,2})\s*[-/ ]\s*([A-Za-z]+)\s*[-/ ]\s*(\d{4})\s*$"),  # 15-Sep-2024
    re.compile(r"^\s*(\d{1,2})\s*[-/ ]\s*([A-Za-z]+)\s*[-/ ]\s*(\d{2})\s*$"),    # 15-Sep-24
    re.compile(r"^\s*Q([1-4])[-/ ](\d{4})\s*$", re.I),                   # Q3 2024
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
        elif idx == 4: # Sep-2024, June 2024
            mon_str = m.group(1).lower()
            mon_num = MONTH_MAP.get(mon_str) or MONTH_MAP.get(mon_str[:3])
            if mon_num:
                return f"{m.group(2)}-{mon_num:02d}-01", text
        elif idx == 5: # June -19, Jun-22
            mon_str = m.group(1).lower()
            mon_num = MONTH_MAP.get(mon_str) or MONTH_MAP.get(mon_str[:3])
            yr_2digit = int(m.group(2))
            full_year = 2000 + yr_2digit if yr_2digit < 70 else 1900 + yr_2digit
            if mon_num:
                return f"{full_year}-{mon_num:02d}-01", text
        elif idx == 6: # 15-Sep-2024
            day, mon_str, year = int(m.group(1)), m.group(2).lower(), m.group(3)
            mon_num = MONTH_MAP.get(mon_str) or MONTH_MAP.get(mon_str[:3])
            if mon_num:
                return f"{year}-{mon_num:02d}-{min(day, 28):02d}", text
        elif idx == 7: # 15-Sep-24
            day, mon_str, yr_2digit = int(m.group(1)), m.group(2).lower(), int(m.group(3))
            mon_num = MONTH_MAP.get(mon_str) or MONTH_MAP.get(mon_str[:3])
            full_year = 2000 + yr_2digit if yr_2digit < 70 else 1900 + yr_2digit
            if mon_num:
                return f"{full_year}-{mon_num:02d}-{min(day, 28):02d}", text
        elif idx == 8: # Q3 2024
            qtr, year = int(m.group(1)), m.group(2)
            qtr_month = (qtr - 1) * 3 + 1
            return f"{year}-{qtr_month:02d}-01", text

    return None, text


def clean_str(val):
    if val is None:
        return None
    s = str(val).strip()
    return s if s else None


def gather_extra(ws, r, hdr_r, start_col):
    extra = {}
    for c in range(start_col, ws.max_column + 1):
        hdr = ws.cell(row=hdr_r, column=c).value
        val = ws.cell(row=r, column=c).value
        if hdr and val is not None:
            extra[str(hdr).strip().replace("\n", " ")] = str(val).strip()
    return extra


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


def save_record(cur, record, extra_dict, clean_wipe):
    evaluate_cui_and_risk(record)
    record["extra_json"] = json.dumps(extra_dict, ensure_ascii=False) if extra_dict else "{}"

    if not clean_wipe:
        existing = None
        if record.get("tag"):
            existing = cur.execute("SELECT id, deferral_status, deferral_reason, deferral_expiry, deferral_approver, deferral_moc_no FROM assets WHERE source_sheet = ? AND tag = ?", (record["source_sheet"], record["tag"])).fetchone()
        elif record.get("asset_number"):
            existing = cur.execute("SELECT id, deferral_status, deferral_reason, deferral_expiry, deferral_approver, deferral_moc_no FROM assets WHERE source_sheet = ? AND asset_number = ?", (record["source_sheet"], record["asset_number"])).fetchone()

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
            return
    
    cols = list(record.keys())
    placeholders = ", ".join(["?"] * len(cols))
    cur.execute(f"INSERT INTO assets ({', '.join(cols)}) VALUES ({placeholders})", [record[k] for k in cols])


# ---------------------------------------------------------------- Precision Sheet Parsers

def parse_vessels_and_tks(conn, ws, clean_wipe):
    cur = conn.cursor()
    count = 0
    hdr_r = 7
    for r in range(8, ws.max_row + 1):
        sn = clean_str(ws.cell(r, 2).value)
        name = clean_str(ws.cell(r, 7).value)
        tag = clean_str(ws.cell(r, 8).value)
        if not sn and not name and not tag:
            continue
        
        d_osi_last, _ = parse_date(ws.cell(r, 14).value)
        d_osi_next, _ = parse_date(ws.cell(r, 15).value)
        d_int_last, _ = parse_date(ws.cell(r, 16).value)
        d_int_next, _ = parse_date(ws.cell(r, 17).value)

        extra = gather_extra(ws, r, hdr_r, 24)
        if ws.cell(r, 21).value:
            extra["RBI Due Date (Internal)"] = str(ws.cell(r, 21).value)
        if ws.cell(r, 22).value:
            extra["RBI Due Date (OSI)"] = str(ws.cell(r, 22).value)

        record = {
            "source_sheet": "Vessels & TKs",
            "sn": sn, "field": clean_str(ws.cell(r, 3).value),
            "plant": clean_str(ws.cell(r, 4).value), "location": clean_str(ws.cell(r, 5).value),
            "unit_name": clean_str(ws.cell(r, 6).value), "name": name or tag,
            "tag": tag, "asset_number": clean_str(ws.cell(r, 9).value),
            "description": clean_str(ws.cell(r, 10).value),
            "in_service": clean_str(ws.cell(r, 11).value),
            "insulation": clean_str(ws.cell(r, 12).value),
            "last_insp_category": clean_str(ws.cell(r, 13).value),
            "date_osi_last": d_osi_last, "date_osi_next": d_osi_next,
            "date_internal_last": d_int_last, "date_internal_next": d_int_next,
            "next_insp_category": clean_str(ws.cell(r, 18).value),
            "corrosion_rate": clean_str(ws.cell(r, 19).value),
            "remaining_life": clean_str(ws.cell(r, 20).value),
            "remarks": clean_str(ws.cell(r, 23).value),
        }
        save_record(cur, record, extra, clean_wipe)
        count += 1
    return count


def parse_coolers(conn, ws, clean_wipe):
    cur = conn.cursor()
    count = 0
    hdr_r = 6
    for r in range(7, ws.max_row + 1):
        sn = clean_str(ws.cell(r, 2).value)
        name = clean_str(ws.cell(r, 3).value)
        tag = clean_str(ws.cell(r, 5).value)
        if not sn and not name and not tag:
            continue

        extra = gather_extra(ws, r, hdr_r, 20)
        record = {
            "source_sheet": "Coolers",
            "sn": sn, "name": name or tag, "tag": tag,
            "plant": "Turbines / Gas Plant", "location": "Cooler Bay",
            "material_spec": clean_str(ws.cell(r, 6).value),
            "nominal_thickness": clean_str(ws.cell(r, 8).value),
            "operating_temp": f"{clean_str(ws.cell(r, 18).value) or ''} / {clean_str(ws.cell(r, 19).value) or ''}".strip(" /"),
            "operating_pressure": clean_str(ws.cell(r, 16).value),
            "fluid_service": clean_str(ws.cell(r, 15).value),
            "remarks": clean_str(ws.cell(r, 24).value),
        }
        save_record(cur, record, extra, clean_wipe)
        count += 1
    return count


def parse_op_piping(conn, ws, clean_wipe):
    cur = conn.cursor()
    count = 0
    hdr_r = 5
    for r in range(6, ws.max_row + 1):
        sn = clean_str(ws.cell(r, 2).value)
        name = clean_str(ws.cell(r, 3).value)
        if not sn and not name:
            continue

        d_last, _ = parse_date(ws.cell(r, 7).value)
        d_next, _ = parse_date(ws.cell(r, 8).value)

        extra = gather_extra(ws, r, hdr_r, 9)
        rem_parts = [clean_str(ws.cell(r, 10).value), clean_str(ws.cell(r, 11).value), clean_str(ws.cell(r, 12).value)]
        rem_str = "\n".join([p for p in rem_parts if p])

        record = {
            "source_sheet": "OP Piping",
            "sn": sn, "name": name, "plant": "Oil Processing (OP)",
            "operating_pressure": clean_str(ws.cell(r, 4).value),
            "design_pressure": clean_str(ws.cell(r, 6).value),
            "date_osi_last": d_last, "date_osi_next": d_next,
            "remarks": rem_str,
        }
        save_record(cur, record, extra, clean_wipe)
        count += 1
    return count


def parse_gp_piping(conn, ws, clean_wipe):
    cur = conn.cursor()
    count = 0
    hdr_r = 8
    current_pack = "PACK 01"

    for r in range(9, ws.max_row + 1):
        c2 = clean_str(ws.cell(r, 2).value)
        c3 = clean_str(ws.cell(r, 3).value)
        if c2 and "PACK" in c2.upper():
            current_pack = c2
            continue
        if not c2 and not c3:
            continue

        d_last, _ = parse_date(ws.cell(r, 9).value)
        d_next, _ = parse_date(ws.cell(r, 10).value)

        pack_val = clean_str(ws.cell(r, 4).value) or current_pack
        extra = gather_extra(ws, r, hdr_r, 13)

        record = {
            "source_sheet": "GP Piping",
            "sn": c2, "name": c3 or f"{pack_val} Line", "field": pack_val,
            "plant": "Gas Plant (GP)", "operating_pressure": clean_str(ws.cell(r, 5).value),
            "design_pressure": clean_str(ws.cell(r, 7).value),
            "insulation": clean_str(ws.cell(r, 8).value),
            "date_osi_last": d_last, "date_osi_next": d_next,
            "corrosion_rate": clean_str(ws.cell(r, 11).value),
            "remarks": clean_str(ws.cell(r, 12).value),
        }
        save_record(cur, record, extra, clean_wipe)
        count += 1
    return count


def parse_turbines_piping(conn, ws, clean_wipe):
    cur = conn.cursor()
    count = 0
    hdr_r = 5
    for r in range(6, ws.max_row + 1):
        sn = clean_str(ws.cell(r, 2).value)
        name = clean_str(ws.cell(r, 3).value)
        if not sn and not name:
            continue

        d_last, _ = parse_date(ws.cell(r, 8).value)
        d_next, _ = parse_date(ws.cell(r, 9).value)

        extra = gather_extra(ws, r, hdr_r, 13)
        record = {
            "source_sheet": "Turbines Piping",
            "sn": sn, "name": name, "field": clean_str(ws.cell(r, 4).value),
            "plant": "Turbines", "operating_pressure": clean_str(ws.cell(r, 5).value),
            "design_pressure": clean_str(ws.cell(r, 7).value),
            "date_osi_last": d_last, "date_osi_next": d_next,
            "corrosion_rate": clean_str(ws.cell(r, 10).value),
            "remarks": clean_str(ws.cell(r, 11).value),
            "insulation": clean_str(ws.cell(r, 12).value),
        }
        save_record(cur, record, extra, clean_wipe)
        count += 1
    return count


def parse_op_dead_legs(conn, ws, clean_wipe):
    cur = conn.cursor()
    count = 0
    hdr_r = 3
    for r in range(4, ws.max_row + 1):
        sn = clean_str(ws.cell(r, 2).value)
        name = clean_str(ws.cell(r, 3).value)
        if not sn and not name:
            continue

        d_last, _ = parse_date(ws.cell(r, 7).value)
        d_next, _ = parse_date(ws.cell(r, 8).value)

        extra = gather_extra(ws, r, hdr_r, 9)
        record = {
            "source_sheet": "OP Dead Legs",
            "sn": sn, "name": name, "plant": "Oil Processing (OP)",
            "location": "Dead Leg", "date_osi_last": d_last, "date_osi_next": d_next,
            "remarks": clean_str(ws.cell(r, 12).value),
        }
        save_record(cur, record, extra, clean_wipe)
        count += 1
    return count


def parse_gp_dead_legs(conn, ws, clean_wipe):
    cur = conn.cursor()
    count = 0
    hdr_r = 6
    for r in range(7, ws.max_row + 1):
        sn = clean_str(ws.cell(r, 3).value) or clean_str(ws.cell(r, 2).value)
        vessel = clean_str(ws.cell(r, 4).value)
        desc = clean_str(ws.cell(r, 5).value)
        if not sn and not desc:
            continue

        d_last, _ = parse_date(ws.cell(r, 7).value)
        d_next, _ = parse_date(ws.cell(r, 8).value)

        extra = gather_extra(ws, r, hdr_r, 9)
        name_str = f"{vessel} - {desc}" if vessel and desc else (desc or vessel or f"GP Dead Leg {sn}")
        record = {
            "source_sheet": "GP Dead Legs",
            "sn": sn, "name": name_str, "plant": "Gas Plant (GP)",
            "location": clean_str(ws.cell(r, 6).value) or "Dead Leg",
            "date_osi_last": d_last, "date_osi_next": d_next,
            "design_pressure": clean_str(ws.cell(r, 13).value),
            "remarks": clean_str(ws.cell(r, 10).value),
        }
        save_record(cur, record, extra, clean_wipe)
        count += 1
    return count


def parse_wd33_piping(conn, ws, clean_wipe):
    cur = conn.cursor()
    count = 0
    hdr_r = 5
    for r in range(6, ws.max_row + 1):
        sn = clean_str(ws.cell(r, 2).value)
        name = clean_str(ws.cell(r, 3).value)
        if not sn and not name:
            continue

        d_last, _ = parse_date(ws.cell(r, 9).value)
        d_next, _ = parse_date(ws.cell(r, 10).value)

        extra = gather_extra(ws, r, hdr_r, 11)
        record = {
            "source_sheet": "WD-33 Piping",
            "sn": sn, "name": name, "in_service": clean_str(ws.cell(r, 4).value),
            "location": clean_str(ws.cell(r, 5).value) or "WD-33",
            "field": clean_str(ws.cell(r, 6).value),
            "operating_pressure": clean_str(ws.cell(r, 7).value),
            "design_pressure": clean_str(ws.cell(r, 8).value),
            "date_osi_last": d_last, "date_osi_next": d_next,
            "remarks": clean_str(ws.cell(r, 11).value),
        }
        save_record(cur, record, extra, clean_wipe)
        count += 1
    return count


def parse_epfs(conn, ws, clean_wipe):
    cur = conn.cursor()
    count = 0
    hdr_r = 4
    for r in range(5, ws.max_row + 1):
        sn = clean_str(ws.cell(r, 2).value)
        name = clean_str(ws.cell(r, 3).value)
        if not sn and not name:
            continue

        d_last, _ = parse_date(ws.cell(r, 7).value)
        d_next, _ = parse_date(ws.cell(r, 8).value)

        extra = gather_extra(ws, r, hdr_r, 10)
        record = {
            "source_sheet": "EPFs",
            "sn": sn, "name": name, "in_service": clean_str(ws.cell(r, 4).value),
            "operating_pressure": clean_str(ws.cell(r, 5).value),
            "design_pressure": clean_str(ws.cell(r, 6).value),
            "date_osi_last": d_last, "date_osi_next": d_next,
            "remarks": clean_str(ws.cell(r, 9).value),
        }
        save_record(cur, record, extra, clean_wipe)
        count += 1
    return count


def parse_gp_inlet_lines(conn, ws, clean_wipe):
    cur = conn.cursor()
    count = 0
    for r in range(1, ws.max_row + 1):
        sn = clean_str(ws.cell(r, 2).value)
        desc = clean_str(ws.cell(r, 3).value)
        if not sn or not sn.isdigit():
            continue

        d_last, _ = parse_date(ws.cell(r, 6).value)
        d_next, _ = parse_date(ws.cell(r, 7).value)

        rem_parts = [clean_str(ws.cell(r, 8).value), clean_str(ws.cell(r, 10).value), clean_str(ws.cell(r, 12).value)]
        rem_str = "\n".join([p for p in rem_parts if p])

        record = {
            "source_sheet": "GP Inlet Lines",
            "sn": sn, "name": desc or f"GP Inlet Line {sn}", "plant": "Gas Plant (GP)",
            "nominal_thickness": clean_str(ws.cell(r, 4).value),
            "operating_pressure": clean_str(ws.cell(r, 5).value),
            "date_osi_last": d_last, "date_osi_next": d_next,
            "remarks": rem_str,
        }
        save_record(cur, record, {}, clean_wipe)
        count += 1
    return count


def parse_mfds(conn, ws, clean_wipe):
    cur = conn.cursor()
    count = 0
    hdr_r = 5
    for r in range(6, ws.max_row + 1):
        sn = clean_str(ws.cell(r, 2).value)
        name = clean_str(ws.cell(r, 3).value)
        if not sn and not name:
            continue

        d_last, _ = parse_date(ws.cell(r, 8).value)
        d_next, _ = parse_date(ws.cell(r, 9).value)

        extra = gather_extra(ws, r, hdr_r, 10)
        record = {
            "source_sheet": "MFDs",
            "sn": sn, "name": name, "field": clean_str(ws.cell(r, 4).value),
            "fluid_service": clean_str(ws.cell(r, 5).value),
            "operating_pressure": clean_str(ws.cell(r, 7).value) or clean_str(ws.cell(r, 6).value),
            "date_osi_last": d_last, "date_osi_next": d_next,
            "remarks": clean_str(ws.cell(r, 11).value),
        }
        save_record(cur, record, extra, clean_wipe)
        count += 1
    return count


def parse_tls(conn, ws, clean_wipe):
    cur = conn.cursor()
    count = 0
    hdr_r = 6
    for r in range(7, ws.max_row + 1):
        loc = clean_str(ws.cell(r, 2).value)
        name = clean_str(ws.cell(r, 4).value)
        if not name and not loc:
            continue

        d_last, _ = parse_date(ws.cell(r, 20).value)
        d_next, _ = parse_date(ws.cell(r, 21).value)

        extra = gather_extra(ws, r, hdr_r, 22)
        record = {
            "source_sheet": "TLs",
            "name": name or f"Trunkline {loc}", "location": loc,
            "in_service": clean_str(ws.cell(r, 5).value),
            "operating_pressure": clean_str(ws.cell(r, 6).value),
            "design_pressure": clean_str(ws.cell(r, 7).value),
            "fluid_service": clean_str(ws.cell(r, 8).value),
            "nominal_thickness": f'{clean_str(ws.cell(r, 9).value) or ""} (Sch {clean_str(ws.cell(r, 10).value) or ""})'.strip(),
            "date_osi_last": d_last, "date_osi_next": d_next,
            "remarks": clean_str(ws.cell(r, 24).value),
        }
        save_record(cur, record, extra, clean_wipe)
        count += 1
    return count


def parse_fls(conn, ws, clean_wipe):
    cur = conn.cursor()
    count = 0
    hdr_r = 3
    for r in range(4, ws.max_row + 1):
        well = clean_str(ws.cell(r, 1).value)
        if not well or "updated" in well.lower():
            continue

        d_last, _ = parse_date(ws.cell(r, 9).value)
        d_next, _ = parse_date(ws.cell(r, 10).value)

        extra = gather_extra(ws, r, hdr_r, 11)
        record = {
            "source_sheet": "FLs",
            "name": f"Flowline {well}", "tag": well, "in_service": clean_str(ws.cell(r, 2).value),
            "fluid_service": clean_str(ws.cell(r, 3).value),
            "operating_pressure": clean_str(ws.cell(r, 5).value),
            "design_pressure": clean_str(ws.cell(r, 6).value),
            "date_osi_last": d_last, "date_osi_next": d_next,
            "remarks": clean_str(ws.cell(r, 13).value),
        }
        save_record(cur, record, extra, clean_wipe)
        count += 1
    return count


def parse_gl_lines(conn, ws, clean_wipe):
    cur = conn.cursor()
    count = 0
    hdr_r = 17
    for r in range(18, ws.max_row + 1):
        sn = clean_str(ws.cell(r, 2).value)
        well = clean_str(ws.cell(r, 3).value)
        if not well or not sn:
            continue

        d_last, _ = parse_date(ws.cell(r, 13).value)
        d_next, _ = parse_date(ws.cell(r, 14).value)

        extra = gather_extra(ws, r, hdr_r, 15)
        record = {
            "source_sheet": "GL Lines",
            "sn": sn, "name": f"Gas Lift Line {well}", "tag": well,
            "fluid_service": "Gas Lift", "nominal_thickness": clean_str(ws.cell(r, 4).value),
            "operating_pressure": clean_str(ws.cell(r, 5).value),
            "operating_temp": clean_str(ws.cell(r, 7).value),
            "t_min": clean_str(ws.cell(r, 11).value),
            "date_osi_last": d_last, "date_osi_next": d_next,
            "remaining_life": clean_str(ws.cell(r, 15).value),
            "remarks": clean_str(ws.cell(r, 12).value),
        }
        save_record(cur, record, extra, clean_wipe)
        count += 1
    return count


def import_temp_repairs(conn, ws):
    cur = conn.cursor()
    cur.execute("DELETE FROM temp_repairs")
    count = 0
    for r in range(9, ws.max_row + 1):
        facility = clean_str(ws.cell(r, 2).value)
        asset_name = clean_str(ws.cell(r, 3).value)
        section = clean_str(ws.cell(r, 4).value)
        repaired_by = clean_str(ws.cell(r, 5).value)
        
        d_orig, _ = parse_date(ws.cell(r, 6).value)
        life_yrs = clean_str(ws.cell(r, 7).value)
        d_exp, _ = parse_date(ws.cell(r, 8).value)
        hardness = clean_str(ws.cell(r, 9).value)
        d_reval, _ = parse_date(ws.cell(r, 10).value)
        d_last_exp, _ = parse_date(ws.cell(r, 11).value)
        status = clean_str(ws.cell(r, 12).value) or "Active"
        report_ref = clean_str(ws.cell(r, 13).value)
        remarks = clean_str(ws.cell(r, 14).value)

        if not any([facility, asset_name, section]):
            continue

        cur.execute("""
            INSERT INTO temp_repairs (
                facility_type, area, asset_name, repaired_section, repaired_by,
                original_repair_date, repair_life_years, expiration_date,
                hardness_hb, revalidation_date, last_expire_date, expiration_status,
                report_ref, remarks
            ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)
        """, (facility, facility, asset_name, section, repaired_by, d_orig, life_yrs,
              d_exp, hardness, d_reval, d_last_exp, status, report_ref, remarks))
        count += 1
    return count


def import_critical_assets(conn, ws):
    cur = conn.cursor()
    cur.execute("DELETE FROM critical_assets")
    count = 0
    current_category = "Critical Replacement Scope"

    for r in range(7, ws.max_row + 1):
        sn = clean_str(ws.cell(r, 2).value)
        pack_no = clean_str(ws.cell(r, 3).value)
        report_no = clean_str(ws.cell(r, 4).value)
        item_desc = clean_str(ws.cell(r, 5).value)
        
        # Section header detection
        if sn and not item_desc and not report_no:
            current_category = sn
            continue

        if not item_desc:
            continue

        d_insp, _ = parse_date(ws.cell(r, 6).value)
        scope = clean_str(ws.cell(r, 7).value)
        done = clean_str(ws.cell(r, 8).value) or "No"
        remarks = clean_str(ws.cell(r, 9).value)
        plant_rem = clean_str(ws.cell(r, 10).value)

        cur.execute("""
            INSERT INTO critical_assets (
                sn, category_section, pack_no, report_no, item_description,
                insp_date, replacement_scope, replacement_done, remarks, plant_remarks
            ) VALUES (?,?,?,?,?,?,?,?,?,?)
        """, (sn, current_category, pack_no, report_no, item_desc, d_insp, scope, done, remarks, plant_rem))
        count += 1
    return count


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


PARSERS = {
    "Vessels & TKs": parse_vessels_and_tks,
    "Coolers": parse_coolers,
    "OP Piping": parse_op_piping,
    "GP Piping": parse_gp_piping,
    "Turbines Piping": parse_turbines_piping,
    "OP Dead Legs": parse_op_dead_legs,
    "GP Dead Legs": parse_gp_dead_legs,
    "WD-33 Piping": parse_wd33_piping,
    "EPFs": parse_epfs,
    "GP Inlet Lines ": parse_gp_inlet_lines,
    "GP Inlet Lines": parse_gp_inlet_lines,
    "MFDs": parse_mfds,
    "TLs": parse_tls,
    "FLs": parse_fls,
    "GL Lines": parse_gl_lines,
}


def import_workbook(xlsx_path, db_path, clean_wipe=False):
    print(f"\n=======================================================")
    print(f"  Precision Ingestion ({'CLEAN WIPE' if clean_wipe else 'SMART SYNC'}): {xlsx_path}")
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
            elif sheet_name in PARSERS:
                count = PARSERS[sheet_name](conn, ws, clean_wipe)
                total_assets += count
                print(f"  [+] {sheet_name:28s} -> {count:4d} assets mapped & synchronized")

    conn.commit()
    conn.close()

    print(f"\n-------------------------------------------------------")
    print(f"  PRECISION INGESTION COMPLETE:")
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