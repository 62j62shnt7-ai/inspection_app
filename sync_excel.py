#!/usr/bin/env python3
"""
sync_excel.py — Write-back engine to synchronize database edits back into Master Inspection Plan Excel workbook.

Preserves all Excel formulas, formatting, colors, and fonts while updating inspection dates,
next due dates, remaining life, turnaround replacement scope, and temporary repairs.
"""
import os
import shutil
import sqlite3
import datetime
import openpyxl
from openpyxl.cell.cell import MergedCell

DEFAULT_XLSX = "/Users/don/Desktop/1. Master Inspection Plan - Updated 4-6-2026.xlsx"
DEFAULT_DB = "inspection_plan.db"


def clean_val(v):
    return str(v).strip() if v is not None else ""


def safe_set_cell(ws, row, col, value):
    if value is None or str(value).strip() == "":
        return
    cell = ws.cell(row=row, column=col)
    if isinstance(cell, MergedCell):
        for rng in ws.merged_cells.ranges:
            if cell.coordinate in rng:
                ws.cell(row=rng.min_row, column=rng.min_col).value = value
                return
    else:
        cell.value = value


def sync_db_to_excel(db_path, xlsx_path, output_path=None):
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

    # 1. Vessels & TKs (Header Row 7)
    if "Vessels & TKs" in wb.sheetnames:
        ws = wb["Vessels & TKs"]
        rows = {clean_val(r["tag"]): r for r in conn.execute("SELECT * FROM assets WHERE source_sheet = 'Vessels & TKs'").fetchall() if r["tag"]}
        sn_rows = {clean_val(r["sn"]): r for r in conn.execute("SELECT * FROM assets WHERE source_sheet = 'Vessels & TKs'").fetchall() if r["sn"]}
        
        for r in range(8, ws.max_row + 1):
            sn = clean_val(ws.cell(r, 2).value)
            tag = clean_val(ws.cell(r, 8).value)
            asset = rows.get(tag) or sn_rows.get(sn)
            if asset:
                if asset["date_osi_last"]: safe_set_cell(ws, r, 14, asset["date_osi_last"])
                if asset["date_osi_next"]: safe_set_cell(ws, r, 15, asset["date_osi_next"])
                if asset["date_internal_last"]: safe_set_cell(ws, r, 16, asset["date_internal_last"])
                if asset["date_internal_next"]: safe_set_cell(ws, r, 17, asset["date_internal_next"])
                if asset["corrosion_rate"]: safe_set_cell(ws, r, 19, asset["corrosion_rate"])
                if asset["remaining_life"]: safe_set_cell(ws, r, 20, asset["remaining_life"])
                if asset["remarks"]: safe_set_cell(ws, r, 23, asset["remarks"])
                updated_count += 1

    # 2. Coolers (Header Row 6)
    if "Coolers" in wb.sheetnames:
        ws = wb["Coolers"]
        tag_rows = {clean_val(r["tag"]): r for r in conn.execute("SELECT * FROM assets WHERE source_sheet = 'Coolers'").fetchall() if r["tag"]}
        sn_rows = {clean_val(r["sn"]): r for r in conn.execute("SELECT * FROM assets WHERE source_sheet = 'Coolers'").fetchall() if r["sn"]}
        for r in range(7, ws.max_row + 1):
            sn = clean_val(ws.cell(r, 2).value)
            tag = clean_val(ws.cell(r, 5).value)
            asset = tag_rows.get(tag) or sn_rows.get(sn)
            if asset:
                if asset["material_spec"]: safe_set_cell(ws, r, 6, asset["material_spec"])
                if asset["nominal_thickness"]: safe_set_cell(ws, r, 8, asset["nominal_thickness"])
                if asset["remarks"]: safe_set_cell(ws, r, 24, asset["remarks"])
                updated_count += 1

    # 3. OP Piping (Header Row 5)
    if "OP Piping" in wb.sheetnames:
        ws = wb["OP Piping"]
        rows = {clean_val(r["sn"]): r for r in conn.execute("SELECT * FROM assets WHERE source_sheet = 'OP Piping'").fetchall() if r["sn"]}
        for r in range(6, ws.max_row + 1):
            sn = clean_val(ws.cell(r, 2).value)
            asset = rows.get(sn)
            if asset:
                if asset["date_osi_last"]: safe_set_cell(ws, r, 7, asset["date_osi_last"])
                if asset["date_osi_next"]: safe_set_cell(ws, r, 8, asset["date_osi_next"])
                if asset["operating_pressure"]: safe_set_cell(ws, r, 4, asset["operating_pressure"])
                if asset["remarks"]: safe_set_cell(ws, r, 12, asset["remarks"])
                updated_count += 1

    # 4. GP Piping (Header Row 8)
    if "GP Piping" in wb.sheetnames:
        ws = wb["GP Piping"]
        rows = {clean_val(r["name"]): r for r in conn.execute("SELECT * FROM assets WHERE source_sheet = 'GP Piping'").fetchall() if r["name"]}
        for r in range(9, ws.max_row + 1):
            name = clean_val(ws.cell(r, 3).value)
            asset = rows.get(name)
            if asset:
                if asset["date_osi_last"]: safe_set_cell(ws, r, 9, asset["date_osi_last"])
                if asset["date_osi_next"]: safe_set_cell(ws, r, 10, asset["date_osi_next"])
                if asset["corrosion_rate"]: safe_set_cell(ws, r, 11, asset["corrosion_rate"])
                if asset["remarks"]: safe_set_cell(ws, r, 12, asset["remarks"])
                updated_count += 1

    # 5. Turbines Piping (Header Row 5)
    if "Turbines Piping" in wb.sheetnames:
        ws = wb["Turbines Piping"]
        rows = {clean_val(r["sn"]): r for r in conn.execute("SELECT * FROM assets WHERE source_sheet = 'Turbines Piping'").fetchall() if r["sn"]}
        for r in range(6, ws.max_row + 1):
            sn = clean_val(ws.cell(r, 2).value)
            asset = rows.get(sn)
            if asset:
                if asset["date_osi_last"]: safe_set_cell(ws, r, 8, asset["date_osi_last"])
                if asset["date_osi_next"]: safe_set_cell(ws, r, 9, asset["date_osi_next"])
                if asset["corrosion_rate"]: safe_set_cell(ws, r, 10, asset["corrosion_rate"])
                if asset["remarks"]: safe_set_cell(ws, r, 11, asset["remarks"])
                updated_count += 1

    # 6. OP Dead Legs (Header Row 3)
    if "OP Dead Legs" in wb.sheetnames:
        ws = wb["OP Dead Legs"]
        rows = {clean_val(r["name"]): r for r in conn.execute("SELECT * FROM assets WHERE source_sheet = 'OP Dead Legs'").fetchall() if r["name"]}
        for r in range(4, ws.max_row + 1):
            name = clean_val(ws.cell(r, 3).value)
            asset = rows.get(name)
            if asset:
                if asset["date_osi_last"]: safe_set_cell(ws, r, 7, asset["date_osi_last"])
                if asset["date_osi_next"]: safe_set_cell(ws, r, 8, asset["date_osi_next"])
                if asset["remarks"]: safe_set_cell(ws, r, 12, asset["remarks"])
                updated_count += 1

    # 7. GP Dead Legs (Header Row 6)
    if "GP Dead Legs" in wb.sheetnames:
        ws = wb["GP Dead Legs"]
        rows = {clean_val(r["sn"]): r for r in conn.execute("SELECT * FROM assets WHERE source_sheet = 'GP Dead Legs'").fetchall() if r["sn"]}
        for r in range(7, ws.max_row + 1):
            sn = clean_val(ws.cell(r, 3).value) or clean_val(ws.cell(r, 2).value)
            asset = rows.get(sn)
            if asset:
                if asset["date_osi_last"]: safe_set_cell(ws, r, 7, asset["date_osi_last"])
                if asset["date_osi_next"]: safe_set_cell(ws, r, 8, asset["date_osi_next"])
                if asset["remarks"]: safe_set_cell(ws, r, 10, asset["remarks"])
                updated_count += 1

    # 8. WD-33 Piping (Header Row 5)
    if "WD-33 Piping" in wb.sheetnames:
        ws = wb["WD-33 Piping"]
        rows = {clean_val(r["sn"]): r for r in conn.execute("SELECT * FROM assets WHERE source_sheet = 'WD-33 Piping'").fetchall() if r["sn"]}
        for r in range(6, ws.max_row + 1):
            sn = clean_val(ws.cell(r, 2).value)
            asset = rows.get(sn)
            if asset:
                if asset["date_osi_last"]: safe_set_cell(ws, r, 9, asset["date_osi_last"])
                if asset["date_osi_next"]: safe_set_cell(ws, r, 10, asset["date_osi_next"])
                if asset["remarks"]: safe_set_cell(ws, r, 11, asset["remarks"])
                updated_count += 1

    # 9. EPFs (Header Row 4)
    if "EPFs" in wb.sheetnames:
        ws = wb["EPFs"]
        rows = {clean_val(r["sn"]): r for r in conn.execute("SELECT * FROM assets WHERE source_sheet = 'EPFs'").fetchall() if r["sn"]}
        for r in range(5, ws.max_row + 1):
            sn = clean_val(ws.cell(r, 2).value)
            asset = rows.get(sn)
            if asset:
                if asset["date_osi_last"]: safe_set_cell(ws, r, 7, asset["date_osi_last"])
                if asset["date_osi_next"]: safe_set_cell(ws, r, 8, asset["date_osi_next"])
                if asset["remarks"]: safe_set_cell(ws, r, 9, asset["remarks"])
                updated_count += 1

    # 10. GP Inlet Lines (Header Row 12)
    if "GP Inlet Lines " in wb.sheetnames or "GP Inlet Lines" in wb.sheetnames:
        sheet_key = "GP Inlet Lines " if "GP Inlet Lines " in wb.sheetnames else "GP Inlet Lines"
        ws = wb[sheet_key]
        rows = {clean_val(r["sn"]): r for r in conn.execute("SELECT * FROM assets WHERE source_sheet LIKE 'GP Inlet%'").fetchall() if r["sn"]}
        for r in range(1, ws.max_row + 1):
            sn = clean_val(ws.cell(r, 2).value)
            asset = rows.get(sn)
            if asset:
                if asset["date_osi_last"]: safe_set_cell(ws, r, 6, asset["date_osi_last"])
                if asset["date_osi_next"]: safe_set_cell(ws, r, 7, asset["date_osi_next"])
                updated_count += 1

    # 11. MFDs (Header Row 5)
    if "MFDs" in wb.sheetnames:
        ws = wb["MFDs"]
        rows = {clean_val(r["name"]): r for r in conn.execute("SELECT * FROM assets WHERE source_sheet = 'MFDs'").fetchall() if r["name"]}
        for r in range(6, ws.max_row + 1):
            name = clean_val(ws.cell(r, 3).value)
            asset = rows.get(name)
            if asset:
                if asset["date_osi_last"]: safe_set_cell(ws, r, 8, asset["date_osi_last"])
                if asset["date_osi_next"]: safe_set_cell(ws, r, 9, asset["date_osi_next"])
                if asset["remarks"]: safe_set_cell(ws, r, 11, asset["remarks"])
                updated_count += 1

    # 12. TLs (Header Row 6)
    if "TLs" in wb.sheetnames:
        ws = wb["TLs"]
        rows = {clean_val(r["name"]): r for r in conn.execute("SELECT * FROM assets WHERE source_sheet = 'TLs'").fetchall() if r["name"]}
        for r in range(7, ws.max_row + 1):
            name = clean_val(ws.cell(r, 4).value)
            asset = rows.get(name)
            if asset:
                if asset["date_osi_last"]: safe_set_cell(ws, r, 20, asset["date_osi_last"])
                if asset["date_osi_next"]: safe_set_cell(ws, r, 21, asset["date_osi_next"])
                if asset["remarks"]: safe_set_cell(ws, r, 24, asset["remarks"])
                updated_count += 1

    # 13. FLs (Header Row 3)
    if "FLs" in wb.sheetnames:
        ws = wb["FLs"]
        rows = {clean_val(r["tag"]): r for r in conn.execute("SELECT * FROM assets WHERE source_sheet = 'FLs'").fetchall() if r["tag"]}
        for r in range(4, ws.max_row + 1):
            well = clean_val(ws.cell(r, 1).value)
            asset = rows.get(well)
            if asset:
                if asset["date_osi_last"]: safe_set_cell(ws, r, 9, asset["date_osi_last"])
                if asset["date_osi_next"]: safe_set_cell(ws, r, 10, asset["date_osi_next"])
                if asset["remarks"]: safe_set_cell(ws, r, 13, asset["remarks"])
                updated_count += 1

    # 14. GL Lines (Header Row 17)
    if "GL Lines" in wb.sheetnames:
        ws = wb["GL Lines"]
        rows = {clean_val(r["tag"]): r for r in conn.execute("SELECT * FROM assets WHERE source_sheet = 'GL Lines'").fetchall() if r["tag"]}
        for r in range(18, ws.max_row + 1):
            well = clean_val(ws.cell(r, 3).value)
            asset = rows.get(well)
            if asset:
                if asset["date_osi_last"]: safe_set_cell(ws, r, 13, asset["date_osi_last"])
                if asset["date_osi_next"]: safe_set_cell(ws, r, 14, asset["date_osi_next"])
                if asset["remaining_life"]: safe_set_cell(ws, r, 15, asset["remaining_life"])
                if asset["remarks"]: safe_set_cell(ws, r, 12, asset["remarks"])
                updated_count += 1

    # 15. Temp-Repair (Header Row 8)
    if "Temp-Repair" in wb.sheetnames:
        ws = wb["Temp-Repair"]
        repairs = {clean_val(r["asset_name"]): r for r in conn.execute("SELECT * FROM temp_repairs").fetchall() if r["asset_name"]}
        for r in range(9, ws.max_row + 1):
            asset_name = clean_val(ws.cell(r, 3).value)
            rep = repairs.get(asset_name)
            if rep:
                if rep["expiration_date"]: safe_set_cell(ws, r, 8, rep["expiration_date"])
                if rep["revalidation_date"]: safe_set_cell(ws, r, 10, rep["revalidation_date"])
                if rep["last_expire_date"]: safe_set_cell(ws, r, 11, rep["last_expire_date"])
                if rep["expiration_status"]: safe_set_cell(ws, r, 12, rep["expiration_status"])
                if rep["remarks"]: safe_set_cell(ws, r, 14, rep["remarks"])
                updated_count += 1

    # 16. Critical Assets (Header Row 6)
    if "Critical Assets" in wb.sheetnames:
        ws = wb["Critical Assets"]
        crits = {clean_val(r["item_description"]): r for r in conn.execute("SELECT * FROM critical_assets").fetchall() if r["item_description"]}
        for r in range(7, ws.max_row + 1):
            desc = clean_val(ws.cell(r, 5).value)
            crit = crits.get(desc)
            if crit:
                if crit["replacement_done"]: safe_set_cell(ws, r, 8, crit["replacement_done"])
                if crit["remarks"]: safe_set_cell(ws, r, 9, crit["remarks"])
                if crit["plant_remarks"]: safe_set_cell(ws, r, 10, crit["plant_remarks"])
                updated_count += 1

    wb.save(target_path)
    conn.close()

    return {
        "success": True,
        "message": f"Successfully synchronized {updated_count:,} records back into '{os.path.basename(target_path)}'.",
        "synced_records": updated_count,
        "target_file": target_path,
        "backup_file": backup
    }


if __name__ == "__main__":
    res = sync_db_to_excel(DEFAULT_DB, DEFAULT_XLSX)
    print(res)
