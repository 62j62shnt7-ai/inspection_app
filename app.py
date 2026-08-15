#!/usr/bin/env python3
"""
app.py — Master Inspection Plan web app.

Standard library only (http.server + sqlite3) — no pip installs needed to
RUN this, which matters on a locked-down corporate Windows machine. Only the
one-time import_excel.py step (run on your Mac) needs openpyxl.

Usage:
    python3 app.py [path/to/inspection_plan.db] [port]

Then open http://localhost:8642 in your browser.
"""
import sys
import os
import json
import csv
import io
import re
import sqlite3
import datetime
import mimetypes
import shutil
import webbrowser
import threading
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import urlparse, parse_qs

DB_PATH = sys.argv[1] if len(sys.argv) > 1 else "inspection_plan.db"
PORT = int(sys.argv[2]) if len(sys.argv) > 2 else 8642
STATIC_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "static")

if not os.path.exists(DB_PATH):
    print(f"Database not found: {DB_PATH}")
    print("Run import_excel.py first to create it.")
    sys.exit(1)


def get_conn():
    conn = sqlite3.connect(DB_PATH, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    # Auto-migrate schema for AIMS columns if missing
    cur = conn.cursor()
    asset_cols = [r[1] for r in cur.execute("PRAGMA table_info(assets)").fetchall()]
    new_cols = [
        ("fluid_service", "TEXT"), ("design_pressure", "TEXT"), ("design_temp", "TEXT"),
        ("operating_pressure", "TEXT"), ("operating_temp", "TEXT"), ("material_spec", "TEXT"),
        ("nominal_thickness", "TEXT"), ("t_min", "TEXT"), ("risk_category", "TEXT"),
        ("damage_mechanisms", "TEXT"), ("cui_susceptible", "INTEGER")
    ]
    for col_name, col_type in new_cols:
        if col_name not in asset_cols:
            try:
                cur.execute(f"ALTER TABLE assets ADD COLUMN {col_name} {col_type}")
            except Exception:
                pass
    conn.commit()
    return conn


CONN_LOCK = threading.Lock()
CONN = get_conn()

ASSET_COLUMNS = frozenset([
    "source_sheet", "sn", "field", "plant", "location", "unit_name", "name",
    "tag", "asset_number", "description", "in_service", "insulation",
    "last_insp_category", "date_osi_last", "date_osi_next",
    "date_internal_last", "date_internal_next", "next_insp_category",
    "corrosion_rate", "remaining_life",
    "fluid_service", "design_pressure", "design_temp", "operating_pressure",
    "operating_temp", "material_spec", "nominal_thickness", "t_min",
    "risk_category", "damage_mechanisms", "cui_susceptible",
    "remarks", "archived",
])

ISO_DATE_RE = re.compile(r"^\d{4}-\d{2}-\d{2}$")


def today():
    """Always-fresh today string so overdue logic survives midnight."""
    return datetime.date.today().isoformat()


def row_to_dict(row):
    d = dict(row)
    try:
        d["extra"] = json.loads(d.get("extra_json") or "{}")
    except Exception:
        d["extra"] = {}
    d.pop("extra_json", None)

    # Calculate CUI (Corrosion Under Insulation) susceptibility:
    # Operating temperature between 10C and 175C (or 50F to 350F) with Insulation present
    insul = str(d.get("insulation") or "").strip().lower()
    has_insulation = insul not in ("", "none", "no", "n/a", "0")
    
    op_temp_str = str(d.get("operating_temp") or "")
    op_temp_val = None
    m = re.search(r"(-?\d+(?:\.\d+)?)", op_temp_str)
    if m:
        try:
            op_temp_val = float(m.group(1))
        except ValueError:
            pass

    if d.get("cui_susceptible") is not None:
        d["is_cui"] = str(d.get("cui_susceptible")).strip().lower() in ("1", "yes", "true")
    else:
        # Auto detect based on insulation & temperature envelope
        if has_insulation and op_temp_val is not None:
            d["is_cui"] = 10.0 <= op_temp_val <= 175.0
        elif has_insulation:
            d["is_cui"] = True
        else:
            d["is_cui"] = False

    # Risk Category default assignment if missing
    risk = str(d.get("risk_category") or "").strip().upper()
    if risk not in ("HIGH", "MEDIUM", "LOW"):
        # Infer default risk if remaining life < 5 yrs or overdue
        rl_str = str(d.get("remaining_life") or "")
        try:
            rl_val = float(re.search(r"(\d+(?:\.\d+)?)", rl_str).group(1)) if re.search(r"(\d+(?:\.\d+)?)", rl_str) else None
        except Exception:
            rl_val = None

        if rl_val is not None and rl_val <= 3.0:
            risk = "HIGH"
        elif rl_val is not None and rl_val <= 10.0:
            risk = "MEDIUM"
        else:
            risk = "LOW"
    d["risk_category"] = risk

    # Earliest real upcoming/overdue date among the tracked due-date fields
    due_candidates = [d.get("date_osi_next"), d.get("date_internal_next")]
    due_candidates = [x for x in due_candidates if x]
    d["next_due"] = min(due_candidates) if due_candidates else None
    if d["next_due"]:
        d["overdue"] = d["next_due"] < today()
    else:
        d["overdue"] = False
    return d


def api_list_assets(params):
    q = params.get("q", [""])[0].strip()
    sheet = params.get("sheet", [""])[0].strip()
    overdue_only = params.get("overdue", ["0"])[0] == "1"
    include_archived = params.get("archived", ["0"])[0] == "1"
    risk_filter = params.get("risk", [""])[0].strip().upper()

    sql = "SELECT * FROM assets WHERE 1=1"
    args = []
    
    if q:
        # FTS5 matches tokens. If they search "pump", we want "pump*".
        # We replace double quotes to prevent syntax errors.
        safe_q = q.replace('"', '""')
        sql = "SELECT * FROM assets WHERE id IN (SELECT rowid FROM assets_fts WHERE assets_fts MATCH ?) "
        args.append(f'"{safe_q}"*')

    if not include_archived:
        sql += " AND archived = 0"
    if sheet:
        sql += " AND source_sheet = ?"
        args.append(sheet)

    with CONN_LOCK:
        rows = [row_to_dict(r) for r in CONN.execute(sql, args).fetchall()]

    if overdue_only:
        rows = [r for r in rows if r["overdue"]]
    if risk_filter in ("HIGH", "MEDIUM", "LOW"):
        rows = [r for r in rows if r["risk_category"] == risk_filter]

    rows.sort(key=lambda r: (r["next_due"] is None, r["next_due"] or ""))
    return rows


def api_dashboard():
    with CONN_LOCK:
        rows = [row_to_dict(r) for r in CONN.execute(
            "SELECT * FROM assets WHERE archived = 0").fetchall()]
        temp_repairs = [dict(r) for r in CONN.execute(
            "SELECT * FROM temp_repairs ORDER BY expiration_date ASC").fetchall()]
        critical_assets = [dict(r) for r in CONN.execute(
            "SELECT * FROM critical_assets ORDER BY id ASC").fetchall()]

    overdue = [r for r in rows if r["overdue"]]
    high_risk = [r for r in rows if r["risk_category"] == "HIGH"]
    cui_flagged = [r for r in rows if r.get("is_cui")]
    expired_repairs = [tr for tr in temp_repairs if (tr.get("expiration_status") or "").lower() == "expired" or (tr.get("expiration_date") and tr.get("expiration_date") < today())]
    pending_critical = [ca for ca in critical_assets if (ca.get("replacement_done") or "").lower() not in ("yes", "1", "true", "completed")]

    in_30 = []
    in_90 = []
    d30 = (datetime.date.today() + datetime.timedelta(days=30)).isoformat()
    d90 = (datetime.date.today() + datetime.timedelta(days=90)).isoformat()
    for r in rows:
        nd = r["next_due"]
        if not nd or r["overdue"]:
            continue
        if nd <= d30:
            in_30.append(r)
        elif nd <= d90:
            in_90.append(r)

    by_sheet = {}
    for r in rows:
        by_sheet.setdefault(r["source_sheet"], {"total": 0, "overdue": 0})
        by_sheet[r["source_sheet"]]["total"] += 1
        if r["overdue"]:
            by_sheet[r["source_sheet"]]["overdue"] += 1

    overdue.sort(key=lambda r: r["next_due"])
    return {
        "total_assets": len(rows),
        "overdue_count": len(overdue),
        "high_risk_count": len(high_risk),
        "cui_count": len(cui_flagged),
        "temp_repairs_count": len(temp_repairs),
        "expired_temp_repairs_count": len(expired_repairs),
        "critical_assets_count": len(critical_assets),
        "pending_critical_count": len(pending_critical),
        "due_30_count": len(in_30),
        "due_90_count": len(in_90),
        "overdue": overdue[:200],
        "due_30": sorted(in_30, key=lambda r: r["next_due"])[:200],
        "high_risk": high_risk[:200],
        "expired_repairs": expired_repairs[:50],
        "pending_critical": pending_critical[:50],
        "by_sheet": by_sheet,
        "sheets": sorted(by_sheet.keys()),
    }


def api_list_temp_repairs():
    with CONN_LOCK:
        rows = CONN.execute("SELECT * FROM temp_repairs ORDER BY expiration_date ASC").fetchall()
    return [dict(r) for r in rows]


def api_create_temp_repair(payload):
    with CONN_LOCK:
        cur = CONN.execute("""
            INSERT INTO temp_repairs (
                facility_type, area, asset_name, repaired_section, repaired_by,
                original_repair_date, repair_life_years, expiration_date,
                hardness_hb, revalidation_date, expiration_status, report_ref, remarks
            ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)
        """, (
            payload.get("facility_type"), payload.get("area"), payload.get("asset_name"),
            payload.get("repaired_section"), payload.get("repaired_by"), payload.get("original_repair_date"),
            payload.get("repair_life_years"), payload.get("expiration_date"), payload.get("hardness_hb"),
            payload.get("revalidation_date"), payload.get("expiration_status", "Active"),
            payload.get("report_ref"), payload.get("remarks")
        ))
        CONN.commit()
        new_id = cur.lastrowid
        row = CONN.execute("SELECT * FROM temp_repairs WHERE id = ?", (new_id,)).fetchone()
    return dict(row)


def api_update_temp_repair(repair_id, payload):
    fields, args = [], []
    allowed = ["facility_type", "area", "asset_name", "repaired_section", "repaired_by",
               "original_repair_date", "repair_life_years", "expiration_date",
               "hardness_hb", "revalidation_date", "last_expire_date", "expiration_status",
               "report_ref", "remarks"]
    for col in allowed:
        if col in payload:
            fields.append(f"{col} = ?")
            args.append(payload[col])
    if fields:
        args.append(repair_id)
        with CONN_LOCK:
            CONN.execute(f"UPDATE temp_repairs SET {', '.join(fields)} WHERE id = ?", args)
            CONN.commit()
    with CONN_LOCK:
        row = CONN.execute("SELECT * FROM temp_repairs WHERE id = ?", (repair_id,)).fetchone()
    return dict(row) if row else None


def api_list_critical_assets():
    with CONN_LOCK:
        rows = CONN.execute("SELECT * FROM critical_assets ORDER BY id ASC").fetchall()
    return [dict(r) for r in rows]


def api_create_critical_asset(payload):
    with CONN_LOCK:
        cur = CONN.execute("""
            INSERT INTO critical_assets (
                sn, category_section, pack_no, report_no, item_description,
                insp_date, replacement_scope, replacement_done, remarks, plant_remarks
            ) VALUES (?,?,?,?,?,?,?,?,?,?)
        """, (
            payload.get("sn"), payload.get("category_section"), payload.get("pack_no"),
            payload.get("report_no"), payload.get("item_description"), payload.get("insp_date"),
            payload.get("replacement_scope"), payload.get("replacement_done", "No"),
            payload.get("remarks"), payload.get("plant_remarks")
        ))
        CONN.commit()
        new_id = cur.lastrowid
        row = CONN.execute("SELECT * FROM critical_assets WHERE id = ?", (new_id,)).fetchone()
    return dict(row)


def api_update_critical_asset(critical_id, payload):
    fields, args = [], []
    allowed = ["sn", "category_section", "pack_no", "report_no", "item_description",
               "insp_date", "replacement_scope", "replacement_done", "remarks", "plant_remarks"]
    for col in allowed:
        if col in payload:
            fields.append(f"{col} = ?")
            args.append(payload[col])
    if fields:
        args.append(critical_id)
        with CONN_LOCK:
            CONN.execute(f"UPDATE critical_assets SET {', '.join(fields)} WHERE id = ?", args)
            CONN.commit()
    with CONN_LOCK:
        row = CONN.execute("SELECT * FROM critical_assets WHERE id = ?", (critical_id,)).fetchone()
    return dict(row) if row else None


def api_yearly_plan(params):
    year = params.get("year", [str(datetime.date.today().year)])[0]
    start = f"{year}-01-01"
    end = f"{year}-12-31"
    with CONN_LOCK:
        rows = [row_to_dict(r) for r in CONN.execute(
            "SELECT * FROM assets WHERE archived = 0").fetchall()]
    plan = [r for r in rows if r["next_due"] and start <= r["next_due"] <= end]
    plan.sort(key=lambda r: r["next_due"])
    return plan


def api_get_asset(asset_id):
    with CONN_LOCK:
        row = CONN.execute("SELECT * FROM assets WHERE id = ?", (asset_id,)).fetchone()
        log = CONN.execute(
            "SELECT * FROM inspection_log WHERE asset_id = ? ORDER BY insp_date DESC",
            (asset_id,)).fetchall()
    if not row:
        return None
    d = row_to_dict(row)
    d["log"] = [dict(x) for x in log]
    return d


def api_update_asset(asset_id, payload):
    fields, args = [], []
    for col in ASSET_COLUMNS:
        if col in payload:
            assert col in ASSET_COLUMNS, f"Unexpected column: {col}"
            fields.append(f"{col} = ?")
            args.append(payload[col])
    if "extra" in payload:
        fields.append("extra_json = ?")
        args.append(json.dumps(payload["extra"], ensure_ascii=False))
    if not fields:
        return api_get_asset(asset_id)
    fields.append("updated_at = ?")
    args.append(datetime.datetime.now().isoformat(timespec="seconds"))
    args.append(asset_id)
    with CONN_LOCK:
        CONN.execute(f"UPDATE assets SET {', '.join(fields)} WHERE id = ?", args)
        CONN.commit()
    return api_get_asset(asset_id)


def api_create_asset(payload):
    # Ensure source_sheet has a default, then build column list (no duplicates)
    payload.setdefault("source_sheet", "Manual Entry")
    cols = [c for c in ASSET_COLUMNS if c in payload]
    vals = [payload[c] for c in cols]
    placeholders = ", ".join(["?"] * len(cols))
    with CONN_LOCK:
        cur = CONN.execute(
            f"INSERT INTO assets ({', '.join(cols)}, extra_json) "
            f"VALUES ({placeholders}, ?)",
            vals + ["{}"])
        CONN.commit()
        new_id = cur.lastrowid
    return api_get_asset(new_id)


def _validate_iso_date(value, field_name):
    """Return the value if it's a valid YYYY-MM-DD string, else raise."""
    if value is None:
        return None
    if not ISO_DATE_RE.match(str(value)):
        raise ValueError(f"Invalid date for {field_name}: expected YYYY-MM-DD, got '{value}'")
    return str(value)


def api_add_log(asset_id, payload):
    insp_date = _validate_iso_date(payload.get("insp_date"), "insp_date")
    next_due = _validate_iso_date(payload.get("next_due_date"), "next_due_date")
    insp_type = (payload.get("insp_type") or "").upper()
    inspector_name = payload.get("inspector_name")
    insp_method = payload.get("insp_method")
    t_actual = payload.get("t_actual")
    action_required = payload.get("action_required")
    
    with CONN_LOCK:
        # Ensure table columns exist
        try:
            CONN.execute("ALTER TABLE inspection_log ADD COLUMN inspector_name TEXT")
            CONN.execute("ALTER TABLE inspection_log ADD COLUMN insp_method TEXT")
            CONN.execute("ALTER TABLE inspection_log ADD COLUMN t_actual TEXT")
            CONN.execute("ALTER TABLE inspection_log ADD COLUMN action_required TEXT")
        except Exception:
            pass

        CONN.execute(
            "INSERT INTO inspection_log (asset_id, insp_date, insp_type, findings, next_due_date, inspector_name, insp_method, t_actual, action_required) "
            "VALUES (?,?,?,?,?,?,?,?,?)",
            (asset_id, insp_date, payload.get("insp_type"),
             payload.get("findings"), next_due, inspector_name, insp_method, t_actual, action_required))
        
        # Determine whether this is OSI or Internal inspection
        is_osi = "OSI" in insp_type
        
        # Build update for asset
        updates = []
        params = []
        
        if insp_date:
            last_field = "date_osi_last" if is_osi else "date_internal_last"
            updates.append(f"{last_field} = ?")
            params.append(insp_date)
            
        if next_due:
            next_field = "date_osi_next" if is_osi else "date_internal_next"
            updates.append(f"{next_field} = ?")
            params.append(next_due)
        else:
            # If no new next due date provided, clear next due date for this type so it isn't marked overdue
            next_field = "date_osi_next" if is_osi else "date_internal_next"
            updates.append(f"{next_field} = NULL")

        if updates:
            params.append(asset_id)
            CONN.execute(f"UPDATE assets SET {', '.join(updates)} WHERE id = ?", params)
        CONN.commit()
    return api_get_asset(asset_id)


def api_export_csv(params):
    sheet = params.get("sheet", [""])[0].strip()
    with CONN_LOCK:
        if sheet:
            rows = CONN.execute(
                "SELECT * FROM assets WHERE archived = 0 AND source_sheet = ?", (sheet,)).fetchall()
        else:
            rows = CONN.execute("SELECT * FROM assets WHERE archived = 0").fetchall()
    
    out = io.StringIO()
    if not rows:
        writer = csv.DictWriter(out, fieldnames=list(ASSET_COLUMNS))
        writer.writeheader()
        return out.getvalue()

    # Collect standard base columns excluding extra_json
    base_cols = [c for c in rows[0].keys() if c != "extra_json"]
    
    # Collect all unique extra_json keys across all rows
    extra_keys = []
    parsed_extras = []
    for r in rows:
        try:
            ex = json.loads(r["extra_json"] or "{}")
        except Exception:
            ex = {}
        parsed_extras.append(ex)
        for k in ex.keys():
            if k not in extra_keys:
                extra_keys.append(k)

    all_cols = base_cols + extra_keys
    writer = csv.DictWriter(out, fieldnames=all_cols, extrasaction="ignore")
    writer.writeheader()

    for r, ex in zip(rows, parsed_extras):
        row_dict = {k: r[k] for k in base_cols}
        row_dict.update(ex)
        writer.writerow(row_dict)

    return out.getvalue()


def _backup_db():
    """Create a timestamped backup of the database before destructive reimport."""
    if os.path.exists(DB_PATH):
        ts = datetime.datetime.now().strftime("%Y%m%d-%H%M%S")
        backup = f"{DB_PATH}.bak.{ts}"
        shutil.copy2(DB_PATH, backup)
        print(f"Database backed up to: {backup}")
        return backup
    return None


def _do_reimport(xlsx_path, display_name):
    """Shared reimport logic: backup DB, run import, reconnect."""
    import import_excel
    _backup_db()
    with CONN_LOCK:
        import_excel.import_workbook(xlsx_path, DB_PATH)
        global CONN
        try:
            CONN.close()
        except Exception:
            pass
        CONN = get_conn()
        row_count = CONN.execute("SELECT COUNT(*) FROM assets").fetchone()[0]
    return {
        "success": True,
        "message": f"Successfully re-imported {row_count} assets from '{display_name}'.",
        "total_assets": row_count,
        "excel_file": display_name,
    }


def api_reimport(payload=None):
    xlsx_file = (payload or {}).get("xlsx_path") or "1. Master Inspection Plan - Updated 4-6-2026.xlsx"
    if not os.path.exists(xlsx_file):
        cwd_files = [f for f in os.listdir(".") if f.endswith(".xlsx") and not f.startswith("~$")]
        if cwd_files:
            xlsx_file = cwd_files[0]
        else:
            raise FileNotFoundError(f"Excel file not found: {xlsx_file}")
    return _do_reimport(xlsx_file, os.path.basename(xlsx_file))


def api_reimport_file(payload):
    import base64
    filename = payload.get("filename") or "uploaded.xlsx"
    filedata = payload.get("filedata")
    if not filedata:
        raise ValueError("No file content uploaded.")

    save_dir = os.path.dirname(os.path.abspath(DB_PATH))
    temp_path = os.path.join(save_dir, "temp_imported_plan.xlsx")

    binary_data = base64.b64decode(filedata)
    with open(temp_path, "wb") as f:
        f.write(binary_data)

    return _do_reimport(temp_path, filename)


def api_delete_log(log_id):
    with CONN_LOCK:
        row = CONN.execute("SELECT asset_id, insp_type FROM inspection_log WHERE id = ?", (log_id,)).fetchone()
        if not row:
            raise ValueError("Log record not found")
        asset_id = row["asset_id"]
        insp_type = (row["insp_type"] or "").upper()
        is_osi = "OSI" in insp_type

        # Delete log entry
        CONN.execute("DELETE FROM inspection_log WHERE id = ?", (log_id,))

        # Recalculate latest inspection & next due date for this asset from remaining logs
        logs = CONN.execute(
            "SELECT insp_date, next_due_date, insp_type FROM inspection_log WHERE asset_id = ? ORDER BY insp_date DESC",
            (asset_id,)).fetchall()
        
        type_logs = [l for l in logs if ("OSI" in (l["insp_type"] or "").upper()) == is_osi]
        
        last_field = "date_osi_last" if is_osi else "date_internal_last"
        next_field = "date_osi_next" if is_osi else "date_internal_next"

        latest_insp = type_logs[0]["insp_date"] if type_logs else None
        latest_next = type_logs[0]["next_due_date"] if type_logs else None

        CONN.execute(
            f"UPDATE assets SET {last_field} = ?, {next_field} = ? WHERE id = ?",
            (latest_insp, latest_next, asset_id)
        )
        CONN.commit()
    return api_get_asset(asset_id)


# MIME types for static file serving
mimetypes.init()
MIME_OVERRIDES = {
    ".js": "application/javascript",
    ".css": "text/css",
    ".html": "text/html",
    ".json": "application/json",
    ".svg": "image/svg+xml",
    ".ico": "image/x-icon",
    ".png": "image/png",
    ".woff2": "font/woff2",
}


class Handler(BaseHTTPRequestHandler):
    def log_message(self, fmt, *args):
        # Quiet normal traffic, but still print errors (status >= 400)
        try:
            status = int(args[1]) if len(args) > 1 else 0
        except (ValueError, IndexError):
            status = 0
        if status >= 400:
            sys.stderr.write(f"[{self.log_date_time_string()}] {fmt % args}\n")

    def _send_json(self, data, status=200):
        body = json.dumps(data, default=str).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def _send_file(self, path, content_type=None):
        if not os.path.exists(path):
            self.send_error(404)
            return
        if content_type is None:
            ext = os.path.splitext(path)[1].lower()
            content_type = MIME_OVERRIDES.get(ext) or mimetypes.guess_type(path)[0] or "application/octet-stream"
        with open(path, "rb") as f:
            body = f.read()
        self.send_response(200)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def _read_json_body(self):
        length = int(self.headers.get("Content-Length", 0))
        if not length:
            return {}
        return json.loads(self.rfile.read(length).decode("utf-8"))

    def _handle_exception(self, e):
        err_msg = str(e)
        if "not found" in err_msg.lower():
            self._send_json({"error": err_msg}, 404)
        elif isinstance(e, ValueError):
            self._send_json({"error": err_msg}, 400)
        else:
            self._send_json({"error": err_msg}, 500)

    def do_GET(self):
        parsed = urlparse(self.path)
        path = parsed.path
        params = parse_qs(parsed.query)

        try:
            if path == "/" or path == "/index.html":
                self._send_file(os.path.join(STATIC_DIR, "index.html"), "text/html")
            elif path == "/api/assets":
                self._send_json(api_list_assets(params))
            elif path == "/api/dashboard":
                self._send_json(api_dashboard())
            elif path == "/api/temp_repairs":
                self._send_json(api_list_temp_repairs())
            elif path == "/api/critical_assets":
                self._send_json(api_list_critical_assets())
            elif path == "/api/yearly_plan":
                self._send_json(api_yearly_plan(params))
            elif path.startswith("/api/assets/"):
                asset_id = int(path.rsplit("/", 1)[-1])
                d = api_get_asset(asset_id)
                if d is None:
                    self._send_json({"error": "not found"}, 404)
                else:
                    self._send_json(d)
            elif path == "/api/export.csv":
                csv_text = api_export_csv(params)
                body = csv_text.encode("utf-8")
                self.send_response(200)
                self.send_header("Content-Type", "text/csv")
                self.send_header("Content-Disposition", 'attachment; filename="inspection_plan_export.csv"')
                self.send_header("Content-Length", str(len(body)))
                self.end_headers()
                self.wfile.write(body)
            else:
                # Serve static files from STATIC_DIR (js, css, images, etc.)
                safe = os.path.normpath(path.lstrip("/"))
                if ".." not in safe:
                    full = os.path.join(STATIC_DIR, safe)
                    if os.path.isfile(full):
                        self._send_file(full)
                        return
                self.send_error(404)
        except Exception as e:
            self._handle_exception(e)

    def do_POST(self):
        parsed = urlparse(self.path)
        path = parsed.path
        try:
            payload = self._read_json_body()
            if path == "/api/assets":
                self._send_json(api_create_asset(payload))
            elif path == "/api/temp_repairs":
                self._send_json(api_create_temp_repair(payload))
            elif path == "/api/critical_assets":
                self._send_json(api_create_critical_asset(payload))
            elif path == "/api/reimport":
                self._send_json(api_reimport(payload))
            elif path == "/api/reimport_file":
                self._send_json(api_reimport_file(payload))
            elif path.endswith("/log") and path.startswith("/api/assets/"):
                asset_id = int(path.split("/")[3])
                self._send_json(api_add_log(asset_id, payload))
            else:
                self.send_error(404)
        except Exception as e:
            self._handle_exception(e)

    def do_PUT(self):
        parsed = urlparse(self.path)
        path = parsed.path
        try:
            payload = self._read_json_body()
            if path.startswith("/api/temp_repairs/"):
                repair_id = int(path.rsplit("/", 1)[-1])
                self._send_json(api_update_temp_repair(repair_id, payload))
            elif path.startswith("/api/critical_assets/"):
                critical_id = int(path.rsplit("/", 1)[-1])
                self._send_json(api_update_critical_asset(critical_id, payload))
            elif path.startswith("/api/assets/"):
                asset_id = int(path.rsplit("/", 1)[-1])
                self._send_json(api_update_asset(asset_id, payload))
            else:
                self.send_error(404)
        except Exception as e:
            self._handle_exception(e)

    def do_DELETE(self):
        parsed = urlparse(self.path)
        path = parsed.path
        try:
            if path.startswith("/api/logs/"):
                log_id = int(path.rsplit("/", 1)[-1])
                self._send_json(api_delete_log(log_id))
            else:
                self.send_error(404)
        except Exception as e:
            self._handle_exception(e)


def main():
    port = PORT
    server = None
    for attempt in range(10):
        try:
            server = ThreadingHTTPServer(("127.0.0.1", port), Handler)
            break
        except OSError as err:
            if err.errno == 48 or getattr(err, "winerror", None) == 10048 or "Address already in use" in str(err):
                # Check if an instance is already running
                if attempt == 0:
                    try:
                        import urllib.request
                        res = urllib.request.urlopen(f"http://127.0.0.1:{port}/api/dashboard", timeout=1)
                        if res.status == 200:
                            url = f"http://127.0.0.1:{port}"
                            print(f"Master Inspection Plan is ALREADY running at {url}")
                            webbrowser.open(url)
                            return
                    except Exception:
                        pass
                port += 1
            else:
                raise err

    if not server:
        print("ERROR: Could not find an available port to launch server.")
        sys.exit(1)

    url = f"http://127.0.0.1:{port}"
    print(f"Master Inspection Plan running at {url}")
    print(f"Database: {os.path.abspath(DB_PATH)}")
    print("Press Ctrl+C to stop.")
    threading.Timer(0.6, lambda: webbrowser.open(url)).start()
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nStopped.")


if __name__ == "__main__":
    main()
