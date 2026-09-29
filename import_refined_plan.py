#!/usr/bin/env python3
"""
import_refined_plan.py — Ingests and cross-references 'refined plan.xlsx'
with the Master Inspection Plan database ('inspection_plan.db').

Links each tactical campaign item to its corresponding statutory master asset
using multi-tier heuristics (Asset#, Tag, Package/SN, Normalized Piping Name).
Computes date variances and preserves field execution progress.
"""
import sys
import os
import re
import sqlite3
import datetime

# Prefer the openpyxl bundled with the app (python/site-packages).
_BUNDLED_SITE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "python", "site-packages")
if os.path.isdir(_BUNDLED_SITE) and _BUNDLED_SITE not in sys.path:
    sys.path.append(_BUNDLED_SITE)

import openpyxl

SCHEMA_REFINED = """
CREATE TABLE IF NOT EXISTS refined_plan_items (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    asset_id INTEGER REFERENCES assets(id) ON DELETE SET NULL,
    category TEXT,
    pkg_or_asset_no TEXT,
    item_description TEXT,
    extracted_tag TEXT,
    last_insp_date TEXT,
    planned_insp_date TEXT,
    scope_category TEXT,
    priority TEXT,
    insp_scope_remarks TEXT,
    ut_progress TEXT,
    report_issued TEXT,
    api_eval_done TEXT,
    kpc_updated TEXT,
    kpc_sent TEXT,
    remarks TEXT,
    match_status TEXT, -- 'MATCHED', 'UNMATCHED', 'MANUAL'
    match_method TEXT, -- 'ASSET_NO', 'TAG', 'PACK_SN', 'NAME_MATCH', 'ALIAS'
    date_variance_days INTEGER, -- planned_insp_date - master_due_date (positive = scheduled after due date)
    synced_to_master INTEGER DEFAULT 0,
    synced_at TEXT,
    created_at TEXT DEFAULT (datetime('now'))
);

CREATE INDEX IF NOT EXISTS idx_refined_asset_id ON refined_plan_items(asset_id);
CREATE INDEX IF NOT EXISTS idx_refined_match_status ON refined_plan_items(match_status);
CREATE INDEX IF NOT EXISTS idx_refined_planned_date ON refined_plan_items(planned_insp_date);
"""

KNOWN_ALIASES = {
    "brakish water tank": "brackish water tank",
    "pacific-3": "pacific-03",
    "olympic-n1x": "olympic-n01x",
    "neag #4": "neag-04",
    "ag #93": "ag-93",
    "op skimmer": "skimmer-7",
    "lp separator": "lp sep",
    "com 1842 to amrya suction": "from comp. 1842 to amerya suction",
    "gpc line": "gpc line",
    "gas boot 201": "gas boot-t 201",
    "gas boot 202": "gas boot-t 202",
    "separator (203) 33-1": "sep. at wd-33-1",
}


def parse_date_str(val):
    if val is None:
        return None
    if isinstance(val, (datetime.datetime, datetime.date)):
        yr = val.year
        if 1930 <= yr <= 1969:
            val = val.replace(year=yr + 100)
        return val.strftime("%Y-%m-%d")
    text = str(val).strip()
    if not text or text.lower() in ["none", "never", "n/a", "na"]:
        return None
    m = re.search(r"(\d{4})[/-](\d{1,2})[/-](\d{1,2})", text)
    if m:
        return f"{m.group(1)}-{int(m.group(2)):02d}-{int(m.group(3)):02d}"
    m_my = re.search(r"(\d{1,2})[/-](\d{4})", text)
    if m_my:
        return f"{m_my.group(2)}-{int(m_my.group(1)):02d}-01"
    m_y = re.search(r"\b(20\d\d)\b", text)
    if m_y:
        return f"{m_y.group(1)}-01-01"
    return None


GENERIC_BANNED = {'line', 'lines', 'drain', 'header', 'separator', 'tank', 'vessel', 'gasboot', 'boot', 'pipe', 'piping', 'skimmer', 'skimmerbox'}
INVALID_TAGS = {"??", "?", "na", "n/a", "none", "-", "", "nan", "null"}
WORK_NOTES = {'site visit', 'cutting location', 'lifting lugs', 'came from beni suef', 'union leak'}


def clean_tokens(text):
    words = re.findall(r'[a-zA-Z0-9]+', str(text or '').lower())
    stop = {'from', 'to', 'at', 'the', 'of', 'and', 'in', 'sec', 'section', 'with', 'for', 'fm', 'line', 'lines', 'header', 'drain'}
    return set(w for w in words if w not in stop and re.search(r'[a-z]', w) and len(w) > 1)


def clean_fl_code(t):
    t = re.sub(r'(?:flowline|\(operated.*?\)|wsw|wiw)\s*', '', str(t or ''), flags=re.I).strip()
    t = re.sub(r'[\s\-#]+', '', t).upper()
    return re.sub(r'([A-Z])0+(\d)', r'\1\2', t)


def match_asset(cat, pkg_or_asset_no, item_desc, all_assets):
    """
    Foolproof Multi-Stage Matcher linking a refined item to an asset in all_assets.
    Prioritizes exact tags, packages, structural routing (trunklines),
    canonical flowline codes, and domain aliases over naive token overlap.
    Returns: (matched_asset, match_method, extracted_tag)
    """
    pkg = str(pkg_or_asset_no or "").strip()
    item = str(item_desc or "").strip()
    cat_str = str(cat or "").strip()
    pkg_lower = pkg.lower()
    item_lower = item.lower()
    item_norm = re.sub(r'[\s\"\'=\-]+', '', item_lower)
    item_norm_clean = re.sub(r't(\d+)', r'\1', item_norm)

    # Stage 0: Detect Tactical Work Notes / Non-Equipment Activities
    if any(wn in item_lower for wn in WORK_NOTES):
        return None, "WORK_NOTE", None

    # Stage 1: Exact Asset# or Tag match on pkg_or_asset_no
    if pkg and pkg_lower not in INVALID_TAGS:
        for a in all_assets:
            a_id, a_sheet, a_field, a_plant, a_loc, a_sn, a_tag, a_asset_no, a_name = a[:9]
            if a_asset_no and a_asset_no.strip().lower() == pkg_lower:
                return a, "ASSET_NO", a_tag
            if a_tag and a_tag.strip().lower() not in INVALID_TAGS and a_tag.strip().lower() == pkg_lower:
                return a, "TAG", a_tag

    # Stage 2: Extract Embedded TAG from item_desc (e.g. TAG: V-1842B, TAG: T-6601, TAG: 10-1201)
    extracted_tag = None
    if item:
        m_tag = re.search(r"TAG:\s*([A-Za-z0-9\-_/]+)", item, re.I)
        if m_tag:
            extracted_tag = m_tag.group(1).strip()
            t_lower = extracted_tag.lower()
            if t_lower not in INVALID_TAGS:
                t_clean = t_lower.replace("-", "").replace(" ", "")
                for a in all_assets:
                    a_tag = a[6]
                    if a_tag and a_tag.strip().lower() not in INVALID_TAGS:
                        a_tag_lower = a_tag.strip().lower()
                        if a_tag_lower == t_lower or a_tag_lower.replace("-", "").replace(" ", "") == t_clean:
                            return a, "TAG_EXTRACTED", extracted_tag

    # Stage 3: Exact Clean Name Match (including multi-line note prefix match)
    if item and len(item_norm) >= 6 and item_norm not in GENERIC_BANNED:
        for a in all_assets:
            if not a[8]:
                continue
            # Compare full name or primary name before newline note
            a_clean_name = a[8].split("\n")[0].strip()
            a_norm = re.sub(r'[\s\"\'=\-]+', '', a[8]).lower()
            a_norm_prefix = re.sub(r'[\s\"\'=\-]+', '', a_clean_name).lower()
            a_norm_clean = re.sub(r't(\d+)', r'\1', a_norm)
            if a_norm == item_norm or a_norm_prefix == item_norm or a_norm_clean == item_norm_clean:
                if 'gp' in cat_str.lower() and a[1] == 'GP Piping':
                    return a, "EXACT_NAME", a[6]
                if 'op' in cat_str.lower() and a[1] == 'OP Piping':
                    return a, "EXACT_NAME", a[6]
                return a, "EXACT_NAME", a[6]

    # Stage 4: Domain-Specific Aliases & Disambiguation
    full_str = f"{pkg} {item}".lower()

    # 4a: Flare Lines in OP / GP
    if "new flare line" in item_lower:
        for a in all_assets:
            if a[1] == "OP Piping" and a[8] and "new flare line" in a[8].lower():
                return a, "ALIAS_FLARE_LINE", a[6]

    # 4b: Skimmers in Vessels & TKs (cross-referencing location & remarks)
    if "rzk skimmer" in item_lower:
        for a in all_assets:
            rem = str(a[12] or "").lower() if len(a) > 12 else ""
            if a[1] == "Vessels & TKs" and ("rzk" in rem or "rzk" in str(a[4] or "").lower()):
                return a, "ALIAS_RZK_SKIMMER", a[6]

    if "skimmer 33-15" in item_lower or ("skimmer box" in item_lower and "33-15" in item_lower):
        for a in all_assets:
            loc = str(a[4] or "").lower()
            name_low = str(a[8] or "").lower()
            if a[1] == "Vessels & TKs" and "33-15" in loc and "skimmer" in name_low:
                return a, "ALIAS_SKIMMER_33_15", a[6]

    if "op skimmer" in item_lower:
        for a in all_assets:
            if (a[6] or "").upper() == "V-0201":
                return a, "ALIAS_MATCH", a[6]

    # 4c: Water Tanks & Vessels
    if "brakish water tank" in item_lower or "brackish water tank" in item_lower:
        for a in all_assets:
            if (a[6] or "").upper() == "T-8302":
                return a, "ALIAS_MATCH", a[6]

    if "lp separator" in item_lower and "sep" in full_str:
        for a in all_assets:
            if (a[6] or "").upper() == "V-1300":
                return a, "ALIAS_MATCH", a[6]

    if "separator (203) 33-1" in item_lower:
        for a in all_assets:
            if a[8] and "sigma sep" in a[8].lower() and "33-15" in a[8].lower():
                return a, "ALIAS_MATCH", a[6]

    # 4d: Compressor piping tie-ins
    if "com 1842 to amrya suction" in item_lower:
        for a in all_assets:
            if a[8] and "1842 to amerya suction" in a[8].lower():
                return a, "ALIAS_MATCH", a[6]

    if "gpc line" in item_lower:
        for a in all_assets:
            if a[8] and "gpc line" in a[8].lower():
                return a, "ALIAS_MATCH", a[6]

    # 4e: Gas boots
    if "gas boot 201" in item_lower:
        for a in all_assets:
            if a[8] and "gas boot-t 201" in a[8].lower():
                return a, "GAS_BOOT_MATCH", a[6]

    if "gas boot 202" in item_lower:
        for a in all_assets:
            if a[8] and "gas boot-t 202" in a[8].lower():
                return a, "GAS_BOOT_MATCH", a[6]

    # 4f: Specific piping drain lines
    if "drain line from tank 201" in item_lower and "burn pit" in item_lower:
        for a in all_assets:
            if a[1] == "OP Piping" and a[8] and "tanks 201,202 to skimmer and burn pit" in a[8].lower():
                return a, "EXACT_PIPING_DRAIN", a[6]

    if "bapetco hp" in item_lower and "gas plant" in item_lower:
        for a in all_assets:
            if a[1] == "GP Piping" and a[8] and "bapetco hp gas line" in a[8].lower():
                return a, "EXACT_BAPETCO_LINE", a[6]

    # Stage 5: Flowlines (FLs) Canonical Code Matching
    if pkg.upper() in ['FLS', 'FL', 'FLOW LINE'] or 'flow line' in cat_str.lower() or re.search(r'^(AG|SWAG|EBED|SHADOW|HAWK|WD|NEAG|OLYMPIC|PACIFIC|MONTEZOUMA|BOLT|NAG)', item, re.I):
        if not any(k in item_lower for k in ['area', 'separator', 'boot', 'skimmer', 'header']):
            # Try exact tag match first (e.g. NAG-1 to E-MFD)
            for a in all_assets:
                if a[1] == 'FLs' and a[6] and a[6].strip().lower() == item.strip().lower():
                    return a, "FLOWLINE_MATCH", a[6]

            fl_code = clean_fl_code(item)
            if len(fl_code) >= 3:
                for a in all_assets:
                    if a[1] == 'FLs':
                        a_code = clean_fl_code(a[8] or '')
                        a_tag_code = clean_fl_code(a[6] or '')
                        if fl_code == a_code or fl_code == a_tag_code:
                            return a, "FLOWLINE_MATCH", a[6]

    # Stage 6: Trunkline (TLs) Structural Routing Matcher
    if pkg.upper() in ['TLS', 'TL'] or 'trunk' in item_lower or 'tl-' in item_lower or 'tl#' in item_lower or 'sag line-b' in item_lower:
        # Special case: SAG Line-B from GP Fence to OP
        if "sag line-b" in item_lower and "to op" in item_lower:
            for a in all_assets:
                if a[1] == "TLs" and a[8] and "header b at ard mfd to oil process" in a[8].lower():
                    return a, "TL_STRUCTURAL_MATCH", a[6]

        m_let = re.search(r'\b(?:trunk\s*line|tl|header)\s*[-#]?\s*([a-g])\b', item_lower)
        req_letter = m_let.group(1).upper() if m_let else None

        is_to_gp = bool(re.search(r'(?:to|&)\s*(?:gp|gas\s*plant)\b', item_lower))
        is_to_op = bool(re.search(r'\bto\s*(?:op|oil\s*process)\b', item_lower))
        is_to_horst = bool(re.search(r'\bto\s*horst\b', item_lower)) and not is_to_gp
        is_to_ard = bool(re.search(r'\bto\s*ard\b', item_lower))
        is_from_horst = bool(re.search(r'\b(?:from|at)\s*horst\b', item_lower))
        is_from_sag = bool(re.search(r'\b(?:from|at)\s*sag\b', item_lower))
        is_from_ard = bool(re.search(r'\b(?:from|at)\s*ard\b', item_lower))

        best_tl = None
        best_score = -1

        for a in all_assets:
            if a[1] != 'TLs' or not a[8]:
                continue
            a_name_low = a[8].lower()
            sc = 0

            # Letter match
            m_a_let = re.search(r'\b(?:trunk\s*line|tl|header)\s*[-#]?\s*([a-g])\b', a_name_low)
            a_let = m_a_let.group(1).upper() if m_a_let else None
            if req_letter:
                if a_let == req_letter:
                    sc += 40
                elif a_let and a_let != req_letter:
                    continue

            # Strict origin & destination verification
            if is_to_gp:
                if "to gas plant" in a_name_low or "to gp" in a_name_low:
                    sc += 30
                elif "to horst" in a_name_low or "to ard" in a_name_low or "to op" in a_name_low:
                    continue  # Direction clash!

            if is_to_op:
                if "to oil process" in a_name_low or "to op" in a_name_low:
                    sc += 30
                elif "to gas plant" in a_name_low or "to horst" in a_name_low:
                    continue

            if is_to_horst:
                if "to horst" in a_name_low:
                    sc += 30
                elif "to gas plant" in a_name_low or "to op" in a_name_low:
                    continue

            if is_to_ard:
                if "to ard" in a_name_low:
                    sc += 30
                elif "to gas plant" in a_name_low or "to op" in a_name_low:
                    continue

            if is_from_horst and ("from horst" in a_name_low or "at horst" in a_name_low): sc += 20
            if is_from_sag and ("from sag" in a_name_low or "at sag" in a_name_low): sc += 20
            if is_from_ard and ("from ard" in a_name_low or "at ard" in a_name_low): sc += 20

            # Token overlap
            c_tokens = clean_tokens(a[8])
            i_tokens = clean_tokens(item)
            sc += len(c_tokens.intersection(i_tokens)) * 5

            if sc > best_score:
                best_score = sc
                best_tl = a

        if best_tl and best_score >= 45:
            return best_tl, "TL_STRUCTURAL_MATCH", best_tl[6]

    # Stage 7: Piping Package Code & Line Number Match
    if pkg:
        m_pkg = re.search(r"PACK\s*([0-9]+[A-Za-z]?)\s*[-–]\s*([0-9]+[A-Za-z]?)", pkg, re.I)
        if m_pkg:
            p_code = m_pkg.group(1).upper()
            p_int = str(int(re.sub(r'[^0-9]', '', p_code))) if re.search(r'\d', p_code) else p_code
            sn_raw = m_pkg.group(2).upper()
            sn_code = str(int(re.sub(r'[^0-9]', '', sn_raw))) if re.search(r'\d', sn_raw) else sn_raw

            for a in all_assets:
                if a[1] in ['OP Piping', 'GP Piping', 'Turbines Piping']:
                    a_f = (a[2] or "").upper()
                    a_sn = (a[5] or "").strip().upper()
                    if (p_code in a_f or p_int == a_f or f"PACK {p_code}" in a_f or f"PACK {p_int}" in a_f) and (a_sn == sn_code or a_sn == sn_raw):
                        return a, "PACK_SN", a[6]

    # Stage 8: Scored Token Overlap with Strict Sanity Guards
    best_match = None
    best_score = 0
    i_tokens = clean_tokens(item)

    for a in all_assets:
        if not a[8]:
            continue

        # Sanity Guard 1: Flare line must not match non-flare
        if "flare" in item_lower and "flare" not in a[8].lower() and "flare" not in str(a[2] or "").lower():
            continue
        # Sanity Guard 2: Scrubber must not match non-scrubber
        if "scrubber" in item_lower and "scrubb" not in a[8].lower() and "pv" not in a[8].lower():
            continue
        # Sanity Guard 3: Tank must not match non-tank
        if ("tank" in item_lower or " tk" in item_lower) and ("tank" not in a[8].lower() and " tk" not in a[8].lower()):
            continue
        # Sanity Guard 4: Separator must not match non-separator
        if ("separator" in item_lower or " sep" in item_lower) and ("sep" not in a[8].lower() and "separator" not in a[8].lower()):
            continue

        a_tokens = clean_tokens(a[8])
        common = i_tokens.intersection(a_tokens)
        if len(common) >= 3:
            score = len(common) * 15
            if score > best_score:
                best_score = score
                best_match = a

    if best_match and best_score >= 40:
        return best_match, "TOKEN_MATCH", best_match[6]

    return None, "UNMATCHED", extracted_tag


def calculate_variance(planned_date_iso, matched_asset, scope_cat):
    """
    Calculates date variance in days: planned_date - master_due_date.
    Positive value: planned LATER than master statutory due date (potential overdue/delay).
    Negative value: planned EARLIER than master statutory due date (proactive).
    """
    if not planned_date_iso or not matched_asset:
        return None
    try:
        p_date = datetime.date.fromisoformat(planned_date_iso)
    except Exception:
        return None

    # Determine master reference date based on scope
    a_id, a_sheet, a_field, a_plant, a_loc, a_sn, a_tag, a_asset_no, a_name, a_osi_n, a_int_n = matched_asset[:11]
    ref_date_str = None
    scope_lower = (scope_cat or "").lower()

    if "internal" in scope_lower:
        ref_date_str = a_int_n or a_osi_n
    elif "osi" in scope_lower:
        ref_date_str = a_osi_n or a_int_n
    else:
        # Earliest due date
        dates = [d for d in [a_osi_n, a_int_n] if d]
        if dates:
            ref_date_str = min(dates)

    if not ref_date_str:
        return None

    try:
        m_date = datetime.date.fromisoformat(ref_date_str)
        return (p_date - m_date).days
    except Exception:
        return None


def import_refined_plan(xlsx_path="refined plan.xlsx", db_path="inspection_plan.db"):
    if not os.path.exists(xlsx_path):
        print(f"Refined plan not found: {xlsx_path}")
        return 0, 0
    if not os.path.exists(db_path):
        print(f"Target database not found: {db_path}")
        return 0, 0

    print(f"\n=======================================================")
    print(f"  Ingesting & Cross-Referencing: {xlsx_path}")
    print(f"  Target SQLite DB: {db_path}")
    print(f"=======================================================\n")

    conn = sqlite3.connect(db_path)
    conn.execute("PRAGMA foreign_keys = ON")
    conn.executescript(SCHEMA_REFINED)
    try:
        conn.execute("ALTER TABLE refined_plan_items ADD COLUMN synced_to_master INTEGER DEFAULT 0")
    except Exception:
        pass
    try:
        conn.execute("ALTER TABLE refined_plan_items ADD COLUMN synced_at TEXT")
    except Exception:
        pass

    # Wipe existing refined items for fresh link
    conn.execute("DELETE FROM refined_plan_items")
    conn.commit()

    # Load master assets for matching
    all_assets = conn.execute(
        "SELECT id, source_sheet, field, plant, location, sn, tag, asset_number, name, date_osi_next, date_internal_next, unit_name, remarks FROM assets WHERE archived = 0"
    ).fetchall()

    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    sheet_name = "Full Inspection Plan"
    if sheet_name not in wb.sheetnames:
        sheet_name = wb.sheetnames[0]
    ws = wb[sheet_name]

    count = 0
    matched_count = 0

    for r in range(3, ws.max_row + 1):
        vals = [ws.cell(row=r, column=col).value for col in range(1, 15)]
        if not any(v is not None for v in vals):
            continue

        cat = str(vals[0] or "").strip() if vals[0] else None
        pkg_or_asset_no = str(vals[1] or "").strip() if vals[1] else None
        item_desc = str(vals[2] or "").strip() if vals[2] else None
        last_date = parse_date_str(vals[3])
        planned_date = parse_date_str(vals[4])
        scope_cat = str(vals[5] or "").strip() if vals[5] else None
        priority = str(vals[6] or "").strip() if vals[6] else None
        insp_scope_remarks = str(vals[7] or "").strip() if vals[7] else None
        ut_progress = str(vals[8] or "").strip() if vals[8] else None
        report_issued = str(vals[9] or "").strip() if vals[9] else None
        api_eval_done = str(vals[10] or "").strip() if vals[10] else None
        kpc_updated = str(vals[11] or "").strip() if vals[11] else None
        kpc_sent = str(vals[12] or "").strip() if vals[12] else None
        remarks = str(vals[13] or "").strip() if vals[13] else None

        # Match against Master Plan
        matched_a, match_method, extracted_tag = match_asset(cat, pkg_or_asset_no, item_desc, all_assets)
        asset_id = matched_a[0] if matched_a else None
        match_status = "MATCHED" if matched_a else "UNMATCHED"
        variance_days = calculate_variance(planned_date, matched_a, scope_cat)

        if match_status == "MATCHED":
            matched_count += 1

        # Check if already logged in master inspection_log
        is_synced = 0
        if asset_id:
            chk = conn.execute(
                "SELECT id FROM inspection_log WHERE asset_id = ? AND (insp_date = ? OR inspector_name LIKE '%Tactical%')",
                (asset_id, planned_date)
            ).fetchone()
            if chk:
                is_synced = 1

        conn.execute("""
            INSERT INTO refined_plan_items (
                asset_id, category, pkg_or_asset_no, item_description, extracted_tag,
                last_insp_date, planned_insp_date, scope_category, priority,
                insp_scope_remarks, ut_progress, report_issued, api_eval_done,
                kpc_updated, kpc_sent, remarks, match_status, match_method,
                date_variance_days, synced_to_master
            ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
        """, (
            asset_id, cat, pkg_or_asset_no, item_desc, extracted_tag,
            last_date, planned_date, scope_cat, priority,
            insp_scope_remarks, ut_progress, report_issued, api_eval_done,
            kpc_updated, kpc_sent, remarks, match_status, match_method,
            variance_days, is_synced
        ))
        count += 1

    conn.commit()
    try:
        conn.execute("PRAGMA wal_checkpoint(TRUNCATE)")
    except Exception:
        pass
    conn.close()

    rate = (matched_count / count * 100) if count else 0
    print(f"  [+] Ingested {count} campaign items.")
    print(f"  [+] Matched {matched_count} / {count} with Master Plan ({rate:.1f}% Match Rate).")
    print(f"  [+] Unmatched items flagged for review: {count - matched_count}.\n")
    return count, matched_count


if __name__ == "__main__":
    xlsx = sys.argv[1] if len(sys.argv) > 1 else "refined plan.xlsx"
    db = sys.argv[2] if len(sys.argv) > 2 else "inspection_plan.db"
    import_refined_plan(xlsx, db)
