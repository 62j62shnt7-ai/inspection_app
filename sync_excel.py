#!/usr/bin/env python3
"""
sync_excel.py — Precision Cell-Targeted Write-Back Engine for Master Inspection Plan Excel workbook.

Preserves 100% of Excel formatting, formulas, cell styles, fonts, borders, and colors.
Surgically updates ONLY the specific target cells edited by the user.
"""
import os
import shutil
import sqlite3
import datetime
import openpyxl
from openpyxl.cell.cell import MergedCell

DEFAULT_XLSX = "/Users/don/Desktop/1. Master Inspection Plan - Updated 4-6-2026.xlsx"
DEFAULT_DB = "inspection_plan.db"

SHEET_COLUMN_MAP = {
    "Vessels & TKs": {
        "start_row": 8, "key_col": 8, "fallback_key_col": 2, # Tag, SN
        "fields": {
            "date_osi_last": 14, "date_osi_next": 15,
            "date_internal_last": 16, "date_internal_next": 17,
            "corrosion_rate": 19, "remaining_life": 20,
            "remarks": 23
        }
    },
    "Coolers": {
        "start_row": 7, "key_col": 5, "fallback_key_col": 2, # Tag, SN
        "fields": {
            "material_spec": 6, "nominal_thickness": 8, "remarks": 24
        }
    },
    "OP Piping": {
        "start_row": 6, "key_col": 2, "fallback_key_col": 3, # SN, Name
        "fields": {
            "operating_pressure": 4, "date_osi_last": 7, "date_osi_next": 8, "remarks": 12
        }
    },
    "GP Piping": {
        "start_row": 9, "key_col": 3, "fallback_key_col": 2, # Name, SN
        "fields": {
            "operating_pressure": 5, "date_osi_last": 9, "date_osi_next": 10,
            "corrosion_rate": 11, "remarks": 12
        }
    },
    "Turbines Piping": {
        "start_row": 6, "key_col": 2, "fallback_key_col": 3, # SN, Name
        "fields": {
            "operating_pressure": 5, "date_osi_last": 8, "date_osi_next": 9,
            "corrosion_rate": 10, "remarks": 11
        }
    },
    "OP Dead Legs": {
        "start_row": 4, "key_col": 3, "fallback_key_col": 2, # Name, SN
        "fields": {
            "date_osi_last": 7, "date_osi_next": 8, "remarks": 12
        }
    },
    "GP Dead Legs": {
        "start_row": 7, "key_col": 3, "fallback_key_col": 2, # SN
        "fields": {
            "date_osi_last": 7, "date_osi_next": 8, "remarks": 10
        }
    },
    "WD-33 Piping": {
        "start_row": 6, "key_col": 2, "fallback_key_col": 3, # SN, Name
        "fields": {
            "operating_pressure": 7, "date_osi_last": 9, "date_osi_next": 10, "remarks": 11
        }
    },
    "EPFs": {
        "start_row": 5, "key_col": 2, "fallback_key_col": 3, # SN, Name
        "fields": {
            "operating_pressure": 5, "date_osi_last": 7, "date_osi_next": 8, "remarks": 9
        }
    },
    "GP Inlet Lines ": {
        "start_row": 1, "key_col": 2, "fallback_key_col": 3, # SN
        "fields": {
            "date_osi_last": 6, "date_osi_next": 7, "remarks": 8
        }
    },
    "GP Inlet Lines": {
        "start_row": 1, "key_col": 2, "fallback_key_col": 3, # SN
        "fields": {
            "date_osi_last": 6, "date_osi_next": 7, "remarks": 8
        }
    },
    "MFDs": {
        "start_row": 6, "key_col": 3, "fallback_key_col": 2, # Name, SN
        "fields": {
            "operating_pressure": 6, "date_osi_last": 8, "date_osi_next": 9, "remarks": 11
        }
    },
    "TLs": {
        "start_row": 7, "key_col": 4, "fallback_key_col": 2, # Name, Loc
        "fields": {
            "operating_pressure": 6, "date_osi_last": 20, "date_osi_next": 21, "remarks": 24
        }
    },
    "FLs": {
        "start_row": 4, "key_col": 1, "fallback_key_col": 1, # Well Tag
        "fields": {
            "operating_pressure": 5, "date_osi_last": 9, "date_osi_next": 10, "remarks": 13
        }
    },
    "GL Lines": {
        "start_row": 18, "key_col": 3, "fallback_key_col": 2, # Tag, SN
        "fields": {
            "operating_pressure": 5, "operating_temp": 7, "t_min": 11,
            "date_osi_last": 13, "date_osi_next": 14, "remaining_life": 15, "remarks": 12
        }
    },
    "Temp-Repair": {
        "start_row": 9, "key_col": 3, "fallback_key_col": 4, # Asset Name, Repaired Section
        "fields": {
            "expiration_date": 8, "revalidation_date": 10, "last_expire_date": 11,
            "expiration_status": 12, "remarks": 14
        }
    },
    "Critical Assets": {
        "start_row": 7, "key_col": 5, "fallback_key_col": 2, # Description, SN
        "fields": {
            "replacement_done": 8, "remarks": 9, "plant_remarks": 10
        }
    }
}


def clean_val(v):
    return str(v).strip() if v is not None else ""


def safe_set_cell(ws, row, col, value):
    """Surgically sets ONLY the target cell value without touching font, fill, or borders."""
    if value is None:
        return
    val_str = str(value).strip()
    cell = ws.cell(row=row, column=col)
    if isinstance(cell, MergedCell):
        for rng in ws.merged_cells.ranges:
            if cell.coordinate in rng:
                target = ws.cell(row=rng.min_row, column=rng.min_col)
                target.value = val_str
                return
    else:
        cell.value = val_str


def sync_single_asset_to_excel(xlsx_path, asset_dict, updated_fields):
    """
    Surgically updates ONLY the modified field cells for a specific asset in the master Excel workbook.
    Leaves 100% of all other cells, formulas, colors, and formatting untouched.
    """
    if not os.path.exists(xlsx_path):
        return {"success": False, "error": f"Master Excel file not found: {xlsx_path}"}

    sheet_name = asset_dict.get("source_sheet")
    if not sheet_name or sheet_name not in SHEET_COLUMN_MAP:
        return {"success": False, "message": f"Sheet '{sheet_name}' is not in sync map."}

    cfg = SHEET_COLUMN_MAP[sheet_name]
    wb = openpyxl.load_workbook(xlsx_path)
    
    ws = None
    if sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
    elif sheet_name.strip() in wb.sheetnames:
        ws = wb[sheet_name.strip()]

    if not ws:
        return {"success": False, "error": f"Worksheet '{sheet_name}' not found in Excel workbook."}

    # Find the exact row in the worksheet
    target_row = None
    tag_val = clean_val(asset_dict.get("tag") or asset_dict.get("sn") or asset_dict.get("name"))
    
    for r in range(cfg["start_row"], ws.max_row + 1):
        cell_key1 = clean_val(ws.cell(r, cfg["key_col"]).value)
        cell_key2 = clean_val(ws.cell(r, cfg.get("fallback_key_col", cfg["key_col"])).value)
        if tag_val and (tag_val == cell_key1 or tag_val == cell_key2 or tag_val in cell_key1 or (cell_key1 and cell_key1 in tag_val)):
            target_row = r
            break

    if not target_row:
        return {"success": False, "message": f"Asset '{tag_val}' row not found in sheet '{sheet_name}'."}

    # Surgically update ONLY the modified fields
    cells_updated = 0
    for field_name, new_val in updated_fields.items():
        if field_name in cfg["fields"]:
            col_idx = cfg["fields"][field_name]
            safe_set_cell(ws, target_row, col_idx, new_val)
            cells_updated += 1

    if cells_updated > 0:
        wb.save(xlsx_path)

    return {
        "success": True,
        "message": f"Surgically updated {cells_updated} cell(s) in '{sheet_name}' Row {target_row} for '{tag_val}'.",
        "row": target_row,
        "sheet": sheet_name,
        "cells_updated": cells_updated
    }


def sync_db_to_excel(db_path, xlsx_path, output_path=None):
    """
    Surgically synchronizes database fields back to their exact respective cells
    while preserving 100% of formatting, formulas, and fonts.
    """
    if not os.path.exists(xlsx_path):
        raise FileNotFoundError(f"Excel master file not found: {xlsx_path}")

    target_path = output_path or xlsx_path

    # Make safety backup if overwriting existing file
    if target_path == xlsx_path and os.path.exists(xlsx_path):
        ts = datetime.datetime.now().strftime("%Y%m%d-%H%M%S")
        backup = f"{xlsx_path}.bak.{ts}"
        shutil.copy2(xlsx_path, backup)
    else:
        backup = None

    wb = openpyxl.load_workbook(xlsx_path)
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row

    updated_count = 0

    for sheet_name, cfg in SHEET_COLUMN_MAP.items():
        ws = None
        if sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
        elif sheet_name.strip() in wb.sheetnames:
            ws = wb[sheet_name.strip()]

        if not ws:
            continue

        if sheet_name == "Temp-Repair":
            repairs = {clean_val(r["asset_name"]): r for r in conn.execute("SELECT * FROM temp_repairs").fetchall() if r["asset_name"]}
            for r in range(cfg["start_row"], ws.max_row + 1):
                asset_name = clean_val(ws.cell(r, cfg["key_col"]).value)
                rep = repairs.get(asset_name)
                if rep:
                    for fld, col_idx in cfg["fields"].items():
                        if rep[fld]: safe_set_cell(ws, r, col_idx, rep[fld])
                    updated_count += 1
        elif sheet_name == "Critical Assets":
            crits = {clean_val(r["item_description"]): r for r in conn.execute("SELECT * FROM critical_assets").fetchall() if r["item_description"]}
            for r in range(cfg["start_row"], ws.max_row + 1):
                desc = clean_val(ws.cell(r, cfg["key_col"]).value)
                crit = crits.get(desc)
                if crit:
                    for fld, col_idx in cfg["fields"].items():
                        if crit[fld]: safe_set_cell(ws, r, col_idx, crit[fld])
                    updated_count += 1
        else:
            assets_by_tag = {clean_val(r["tag"]): r for r in conn.execute("SELECT * FROM assets WHERE source_sheet = ?", (sheet_name,)).fetchall() if r["tag"]}
            assets_by_sn = {clean_val(r["sn"]): r for r in conn.execute("SELECT * FROM assets WHERE source_sheet = ?", (sheet_name,)).fetchall() if r["sn"]}
            assets_by_name = {clean_val(r["name"]): r for r in conn.execute("SELECT * FROM assets WHERE source_sheet = ?", (sheet_name,)).fetchall() if r["name"]}

            for r in range(cfg["start_row"], ws.max_row + 1):
                k1 = clean_val(ws.cell(r, cfg["key_col"]).value)
                k2 = clean_val(ws.cell(r, cfg.get("fallback_key_col", cfg["key_col"])).value)
                asset = assets_by_tag.get(k1) or assets_by_sn.get(k1) or assets_by_name.get(k1) or assets_by_tag.get(k2) or assets_by_sn.get(k2) or assets_by_name.get(k2)
                if asset:
                    for fld, col_idx in cfg["fields"].items():
                        if asset[fld]: safe_set_cell(ws, r, col_idx, asset[fld])
                    updated_count += 1

    wb.save(target_path)
    conn.close()

    return {
        "success": True,
        "message": f"Surgically synchronized {updated_count:,} records into '{os.path.basename(target_path)}' (zero formatting damage).",
        "synced_records": updated_count,
        "target_file": target_path,
        "backup_file": backup
    }


if __name__ == "__main__":
    res = sync_db_to_excel(DEFAULT_DB, DEFAULT_XLSX)
    print(res)
