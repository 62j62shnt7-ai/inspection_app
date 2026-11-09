#!/usr/bin/env python3
"""Create a lossless, reviewable Master Register worksheet from a workbook."""

from __future__ import annotations

import argparse
import copy
import hashlib
import json
import os
import posixpath
import re
import shutil
import tempfile
import zipfile
import xml.etree.ElementTree as ET
from collections import Counter
from datetime import date, datetime
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation


REGISTER_SHEET = "Master Register"
MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
DOC_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
REGISTER_HEADERS = [
    "Asset ID", "Asset Class", "Asset Name", "Tag", "Asset Number",
    "Field", "Plant", "Location", "Unit / Pack", "In Service",
    "OSI Applicable", "Last OSI Date", "Next OSI Date", "OSI Status / Hold Reason",
    "Internal Applicable", "Last Internal Date", "Next Internal Date",
    "Internal Status / Hold Reason", "Remarks", "Source Sheet", "Source Row",
    "Row Type", "Review Flags", "Source Data JSON",
]

DATE_KEYS = {
    "osi_last": ("last osi date", "last osi", "last inspection date", "last inspection", "last insp date", "last insp"),
    "osi_next": ("next osi due date", "next osi date", "next osi", "next inspection date", "next inspection", "next insp date", "next insp", "next due"),
    "internal_last": ("last internal",),
    "internal_next": ("next internal",),
}


def _text(value):
    if value is None:
        return ""
    if isinstance(value, (datetime, date)):
        return value.isoformat()
    return re.sub(r"\s+", " ", str(value)).strip()


def _normal(value):
    return re.sub(r"[^a-z0-9]+", " ", _text(value).lower()).strip()


def _date_value(value):
    if isinstance(value, datetime):
        return value.date().isoformat()
    if isinstance(value, date):
        return value.isoformat()
    s = _text(value)
    if not s:
        return None
    if _is_ambiguous_numeric_date(s):
        return None
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y", "%d-%m-%Y", "%m-%d-%Y", "%d-%b-%Y", "%d-%B-%Y"):
        try:
            return datetime.strptime(s, fmt).date().isoformat()
        except ValueError:
            continue
    month_year = re.fullmatch(r"([A-Za-z]+)\s*-\s*(\d{2,4})", s)
    if month_year:
        month_name, year_text = month_year.groups()
        for fmt in ("%B", "%b"):
            try:
                month = datetime.strptime(month_name, fmt).month
                year = int(year_text)
                if year < 100:
                    year += 2000 if year < 70 else 1900
                return date(year, month, 1).isoformat()
            except ValueError:
                continue
    return None


def _is_month_year(value):
    return bool(re.fullmatch(r"[A-Za-z]+\s*-\s*\d{2,4}", _text(value)))


def _is_ambiguous_numeric_date(value):
    match = re.fullmatch(r"(\d{1,2})[/-](\d{1,2})[/-](\d{4})", _text(value))
    if not match:
        return False
    first, second = int(match.group(1)), int(match.group(2))
    return first != second and 1 <= first <= 12 and 1 <= second <= 12


def _find_header(ws):
    keys = ("last insp", "next insp", "last inspection", "next inspection", "asset name", "pipeline name", "tag", "well", "item")
    best_row, best_score = None, 0
    for row_num in range(1, min(ws.max_row, 30) + 1):
        vals = [_normal(ws.cell(row_num, col).value) for col in range(1, ws.max_column + 1)]
        score = sum(1 for v in vals if v and any(k in v for k in keys))
        if score > best_score:
            best_row, best_score = row_num, score
    return best_row if best_score >= 2 else None


def _sheet_class(sheet_name, asset_name=""):
    sn = _normal(sheet_name)
    nm = _normal(asset_name)
    if "vessel" in sn or "tank" in sn:
        if "tank" in nm:
            return "Tank"
        vessel_terms = ("vessel", "separator", "scrubber", "drum", "exchanger", "hex", "filter", "tower", "column", "accumulator", "kod")
        return "Vessel" if any(term in nm for term in vessel_terms) else "Other Equipment"
    if "trunkline" in sn or sn == "tls":
        return "Trunkline"
    if "flowline" in sn or sn == "fls":
        return "Flowline"
    if "gas lift" in sn or sn == "gl lines":
        return "Gas Lift Line"
    if "dead leg" in sn:
        return "Dead Leg"
    if "piping" in sn or "line" in sn or sn in ("epfs", "mfds"):
        return "Piping / Line"
    return "Other Asset"


def _header_index(headers, fragments):
    normalized_headers = [_normal(h) for h in headers]
    normalized_fragments = sorted((_normal(f) for f in fragments), key=len, reverse=True)
    # Prefer a complete header match before substring matching. This avoids
    # mapping "Last Insp" to "Number of Leaks after last Insp date".
    for fragment in normalized_fragments:
        for i, header in enumerate(normalized_headers):
            if header and header == fragment:
                return i
    for fragment in normalized_fragments:
        for i, header in enumerate(normalized_headers):
            if header and fragment in header:
                return i
    return None


def _pick(headers, values, *fragments):
    idx = _header_index(headers, fragments)
    return values[idx] if idx is not None and idx < len(values) else None


def _raw_map(headers, values):
    result = {}
    for i, value in enumerate(values):
        if value is None:
            continue
        key = _text(headers[i]) if i < len(headers) and _text(headers[i]) else f"Column {get_column_letter(i + 1)}"
        if key in result:
            key = f"{key} [{get_column_letter(i + 1)}]"
        result[key] = value.isoformat() if isinstance(value, (datetime, date)) else value
    return result


def classify_row(sheet_name, row_num, headers, values, header_row=None):
    """Map a row without discarding source cells; non-asset rows are references."""
    nonempty = any(v is not None and _text(v) for v in values)
    row_type = "Asset"
    flags = []
    if not nonempty:
        row_type = "Empty"
    elif header_row is None or row_num <= header_row:
        row_type = "Review / Reference"
        flags.append("REFERENCE_ROW")

    name_headers = ("asset name", "pipeline name", "line desc", "item description", "item", "well", "gas lift line")
    name_idx = _header_index(headers, name_headers) if row_type == "Asset" else None
    name_value = values[name_idx] if name_idx is not None and name_idx < len(values) else None
    if row_type == "Asset" and name_idx is not None and not _text(name_value):
        row_type = "Review / Reference"
        flags.append("REFERENCE_ROW")
    name = _text(name_value)
    if not name and row_type == "Asset" and name_idx is None:
        name = next((_text(v) for v in values if _text(v)), "")

    klass = _sheet_class(sheet_name, name) if row_type == "Asset" else "Review / Reference"
    name_normal = _normal(name)
    if row_type == "Asset" and (name.lower().startswith(("pack ", "*", "notes", "note:", "common recommendations"))
                                 or name_normal in ("from", "to")
                                 or "for ut at" in name.lower() or "all gas lift lines" in name.lower()):
        row_type = "Review / Reference"
        klass = "Review / Reference"
        osi_applicable = "No"
        internal_applicable = "No"
        flags.append("REFERENCE_ROW")
    internal_applicable = "Yes" if klass in ("Tank", "Vessel") else "No"
    if klass == "Other Equipment" and ("vessel" in _normal(sheet_name) or "tank" in _normal(sheet_name)):
        internal_applicable = "Review"
        flags.append("INTERNAL_APPLICABILITY_REVIEW")
    osi_applicable = "Yes" if row_type == "Asset" else "No"

    tag_value = _pick(headers, values, "tag")
    asset_num_value = _pick(headers, values, "asset number", "asset no", "asset #", "oracle")
    strong_id = _text(tag_value) or _text(asset_num_value) or _text(_pick(headers, values, "well", "pipeline name", "line desc", "gas lift line"))
    source_id = strong_id or _text(_pick(headers, values, "report", "sn", "no"))
    if not source_id and row_type == "Asset":
        source_id = name
    if row_type == "Asset" and not strong_id:
        flags.append("ID_FROM_ASSET_NAME_REVIEW")
    if row_type == "Asset" and not source_id:
        flags.append("GENERATED_ID_SOURCE_ID_MISSING")

    row_text = " ".join(_text(v).lower() for v in values if isinstance(v, str))
    update_metadata = any(marker in row_text for marker in ("updated on", "updated date", "last update", "update on"))
    if row_type == "Review / Reference" and update_metadata:
        if any(isinstance(v, (date, datetime)) and (v.date() if isinstance(v, datetime) else v) > date.today() for v in values):
            flags.append("FUTURE_UPDATE_DATE_REVIEW")

    date_outputs = {}
    status_outputs = {}
    for field, keys in DATE_KEYS.items():
        if row_type != "Asset":
            date_outputs[field] = None
            status_outputs[field] = ""
            continue
        idx = _header_index(headers, keys)
        raw = values[idx] if idx is not None and idx < len(values) else None
        parsed = _date_value(raw)
        date_outputs[field] = parsed
        status = "" if parsed or raw is None else _text(raw)
        status_outputs[field] = status
        if parsed and _is_month_year(raw):
            flags.append("PARTIAL_DATE_ASSUMED_DAY_01")
        if raw is not None and _is_ambiguous_numeric_date(raw):
            flags.append("AMBIGUOUS_DATE_ORDER_REVIEW")
        if raw is not None and _text(raw) and not parsed:
            flags.append(f"UNPARSED_{field.upper()}_REVIEW")

    if row_type == "Asset" and not date_outputs.get("osi_next") and not status_outputs.get("osi_next"):
        flags.append("OSI_DUE_INFORMATION_MISSING")

    remarks_idx = _header_index(headers, ("remarks", "recommendation", "concern", "corrosion remarks", "findings"))
    remarks = _text(values[remarks_idx]) if remarks_idx is not None and remarks_idx < len(values) else ""
    if row_type == "Asset" and not name:
        flags.append("ASSET_NAME_MISSING")
    if row_type == "Asset" and not source_id:
        flags.append("IDENTITY_REVIEW")

    return {
        "asset_class": klass, "asset_name": name, "tag": _text(tag_value),
        "asset_number": _text(asset_num_value),
        "field": _text(_pick(headers, values, "field")),
        "plant": _text(_pick(headers, values, "plant")),
        "location": _text(_pick(headers, values, "location", "area")),
        "unit_pack": _text(_pick(headers, values, "unit", "pack", "report")),
        "in_service": _text(_pick(headers, values, "in service")),
        "osi_applicable": osi_applicable,
        "osi_last_date": date_outputs["osi_last"], "osi_next_date": date_outputs["osi_next"],
        "osi_status": status_outputs["osi_next"],
        "internal_applicable": internal_applicable,
        "internal_last_date": date_outputs["internal_last"] if internal_applicable in ("Yes", "Review") else None,
        "internal_next_date": date_outputs["internal_next"] if internal_applicable in ("Yes", "Review") else None,
        "internal_status": status_outputs["internal_next"] if internal_applicable in ("Yes", "Review") else "",
        "remarks": remarks, "row_type": row_type,
        "review_flags": "; ".join(dict.fromkeys(flags)),
        "source_raw": json.dumps(_raw_map(headers, values), ensure_ascii=False, default=str),
        "source_id": source_id,
        "strong_source_id": bool(strong_id),
    }


def _slug(value):
    s = re.sub(r"[^A-Za-z0-9]+", "-", value).strip("-").upper()
    return s or "SOURCE"


def _rows_for_sheet(ws, header_row):
    headers = []
    if header_row:
        headers = [_text(ws.cell(header_row, col).value) for col in range(1, ws.max_column + 1)]
    for row_num in range(1, ws.max_row + 1):
        values = [ws.cell(row_num, col).value for col in range(1, ws.max_column + 1)]
        if any(v is not None and _text(v) for v in values):
            yield row_num, headers, values


def _worksheet_parts(archive):
    workbook_root = ET.fromstring(archive.read("xl/workbook.xml"))
    rels_root = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
    rel_targets = {rel.attrib["Id"]: rel.attrib["Target"] for rel in rels_root.findall(f"{{{PKG_REL_NS}}}Relationship")}
    parts = {}
    for sheet in workbook_root.findall(f".//{{{MAIN_NS}}}sheet"):
        rel_id = sheet.attrib[f"{{{DOC_REL_NS}}}id"]
        target = rel_targets[rel_id]
        parts[sheet.attrib["name"]] = target.lstrip("/") if target.startswith("/") else posixpath.normpath(posixpath.join("xl", target))
    return parts


def _append_style_children(styles_xml, generated_root, collection_name, children, count_offset=None):
    namespace = f"{{{MAIN_NS}}}"
    generated = generated_root.find(namespace + collection_name)
    if generated is None or not children:
        return styles_xml
    pattern = re.compile(rb"(<" + collection_name.encode() + rb"\b[^>]*>)(.*?)(</" + collection_name.encode() + rb">)", re.DOTALL)
    match = pattern.search(styles_xml)
    if not match:
        if collection_name != "numFmts":
            return styles_xml
        opening_name = b"<numFmts count=\"" + str(len(children)).encode() + b"\">"
        added = b"".join(ET.tostring(child, encoding="utf-8") for child in children)
        insert_at = styles_xml.find(b"<fonts")
        if insert_at < 0:
            return styles_xml.replace(b"</styleSheet>", opening_name + added + b"</numFmts></styleSheet>", 1)
        return styles_xml[:insert_at] + opening_name + added + b"</numFmts>" + styles_xml[insert_at:]
    opening, old_children, closing = match.groups()
    if count_offset is not None:
        opening = re.sub(rb'\bcount="\d+"', b'count="' + str(count_offset + len(children)).encode() + b'"', opening, count=1)
    added = b"".join(ET.tostring(child, encoding="utf-8") for child in children)
    return styles_xml[:match.start()] + opening + old_children + added + closing + styles_xml[match.end():]


def _merge_register_into_source(source_path, generated_path, output_path):
    """Add only the new sheet/package metadata; keep source parts byte-for-byte."""
    main = f"{{{MAIN_NS}}}"
    with zipfile.ZipFile(source_path, "r") as source_zip, zipfile.ZipFile(generated_path, "r") as generated_zip:
        source_parts = _worksheet_parts(source_zip)
        generated_parts = _worksheet_parts(generated_zip)
        if REGISTER_SHEET in source_parts:
            raise ValueError("Input already contains a Master Register; this generator is for the one-time source migration.")
        generated_sheet = generated_zip.read(generated_parts[REGISTER_SHEET])

        source_styles_xml = source_zip.read("xl/styles.xml")
        generated_styles_root = ET.fromstring(generated_zip.read("xl/styles.xml"))
        source_styles_root = ET.fromstring(source_styles_xml)
        style_child_offsets = {}
        for name in ("fonts", "fills", "borders", "cellXfs", "numFmts"):
            current = source_styles_root.find(main + name)
            style_child_offsets[name] = len(current) if current is not None else 0

        generated_numfmts = generated_styles_root.find(main + "numFmts")
        original_numfmts = source_styles_root.find(main + "numFmts")
        numfmt_map = {}
        if generated_numfmts is not None:
            max_id = max((int(x.attrib.get("numFmtId", 163)) for x in (list(original_numfmts) if original_numfmts is not None else [])), default=163)
            generated_num_ids = {int(x.attrib["numFmtId"]) for x in generated_numfmts}
            numfmt_map = {old: max_id + i + 1 for i, old in enumerate(sorted(generated_num_ids))}
            numfmt_children = []
            for node in generated_numfmts:
                clone = copy.deepcopy(node)
                clone.attrib["numFmtId"] = str(numfmt_map[int(clone.attrib["numFmtId"])])
                numfmt_children.append(clone)
        else:
            numfmt_children = []

        generated_fonts = generated_styles_root.find(main + "fonts")
        generated_fills = generated_styles_root.find(main + "fills")
        generated_borders = generated_styles_root.find(main + "borders")
        generated_xfs = generated_styles_root.find(main + "cellXfs")
        fonts = list(generated_fonts) if generated_fonts is not None else []
        fills = list(generated_fills) if generated_fills is not None else []
        borders = list(generated_borders) if generated_borders is not None else []
        xfs = []
        for node in (list(generated_xfs) if generated_xfs is not None else []):
            clone = copy.deepcopy(node)
            for attr, collection in (("fontId", "fonts"), ("fillId", "fills"), ("borderId", "borders")):
                if attr in clone.attrib:
                    clone.attrib[attr] = str(int(clone.attrib[attr]) + style_child_offsets[collection])
            if "numFmtId" in clone.attrib and int(clone.attrib["numFmtId"]) in numfmt_map:
                clone.attrib["numFmtId"] = str(numfmt_map[int(clone.attrib["numFmtId"])])
            xfs.append(clone)

        styles_xml = source_styles_xml
        styles_xml = _append_style_children(styles_xml, generated_styles_root, "numFmts", numfmt_children, style_child_offsets["numFmts"])
        styles_xml = _append_style_children(styles_xml, generated_styles_root, "fonts", fonts, style_child_offsets["fonts"])
        styles_xml = _append_style_children(styles_xml, generated_styles_root, "fills", fills, style_child_offsets["fills"])
        styles_xml = _append_style_children(styles_xml, generated_styles_root, "borders", borders, style_child_offsets["borders"])
        styles_xml = _append_style_children(styles_xml, generated_styles_root, "cellXfs", xfs, style_child_offsets["cellXfs"])

        style_offset = style_child_offsets["cellXfs"]
        generated_sheet = re.sub(
            rb'<c\b[^>]*>',
            lambda m: re.sub(rb'\bs="(\d+)"', lambda s: b's="' + str(style_offset + int(s.group(1))).encode() + b'"', m.group(0)),
            generated_sheet,
        )

        workbook_xml = source_zip.read("xl/workbook.xml")
        workbook_rels = source_zip.read("xl/_rels/workbook.xml.rels")
        content_types = source_zip.read("[Content_Types].xml")
        used_ids = [int(x) for x in re.findall(rb'\brId(\d+)\b', workbook_rels)]
        rel_id = f"rId{max(used_ids, default=0) + 1}"
        sheet_ids = [int(x) for x in re.findall(rb'\bsheetId="(\d+)"', workbook_xml)]
        sheet_id = max(sheet_ids, default=0) + 1
        sheet_part = f"xl/worksheets/sheet{sheet_id}.xml"
        sheet_tag = (f'<sheet name="{REGISTER_SHEET}" sheetId="{sheet_id}" r:id="{rel_id}"/>').encode()
        workbook_xml = workbook_xml.replace(b"</sheets>", sheet_tag + b"</sheets>", 1)
        rel_tag = (f'<Relationship Id="{rel_id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{sheet_id}.xml"/>').encode()
        workbook_rels = workbook_rels.replace(b"</Relationships>", rel_tag + b"</Relationships>", 1)
        override = (f'<Override PartName="/{sheet_part}" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>').encode()
        content_types = content_types.replace(b"</Types>", override + b"</Types>", 1)

        replacements = {
            "xl/workbook.xml": workbook_xml,
            "xl/_rels/workbook.xml.rels": workbook_rels,
            "[Content_Types].xml": content_types,
            "xl/styles.xml": styles_xml,
            sheet_part: generated_sheet,
        }
        with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as result_zip:
            for info in source_zip.infolist():
                result_zip.writestr(info, replacements.pop(info.filename, source_zip.read(info.filename)))
            for name, payload in replacements.items():
                result_zip.writestr(name, payload, compress_type=zipfile.ZIP_DEFLATED)


def build_register(input_path, output_path):
    """Copy workbook, retaining original sheets and adding the register sheet."""
    input_path, output_path = Path(input_path), Path(output_path)
    if input_path.resolve() == output_path.resolve():
        raise ValueError("Output must be a new workbook path; the source workbook is never overwritten.")
    if output_path.exists():
        raise FileExistsError(f"Output already exists and may contain your review edits: {output_path}")
    fd, generated_path = tempfile.mkstemp(prefix="register-sheet-", suffix=".xlsx", dir=str(output_path.parent))
    os.close(fd)
    try:
        shutil.copy2(input_path, generated_path)
        wb = openpyxl.load_workbook(generated_path, data_only=False)
        if REGISTER_SHEET in wb.sheetnames:
            del wb[REGISTER_SHEET]
        ws_out = wb.create_sheet(REGISTER_SHEET)
        ws_out.append(REGISTER_HEADERS)

        totals = Counter()
        id_counts = Counter()
        pending = []
        for ws in wb.worksheets:
            if ws.title == REGISTER_SHEET:
                continue
            header_row = _find_header(ws)
            for row_num, headers, values in _rows_for_sheet(ws, header_row):
                mapped = classify_row(ws.title, row_num, headers, values, header_row)
                source_id = mapped.pop("source_id")
                strong_source_id = mapped.pop("strong_source_id")
                if source_id and strong_source_id:
                    id_counts[(ws.title, source_id.casefold())] += 1
                identity = source_id if strong_source_id else mapped["asset_name"]
                if mapped["row_type"] != "Asset" or not identity:
                    identity = f"row:{row_num}"
                pending.append((ws.title, row_num, mapped, source_id, strong_source_id, identity))

        identity_counts = Counter((sheet, identity.casefold()) for sheet, _, mapped, _, _, identity in pending if mapped["row_type"] == "Asset")
        used_ids = set()
        for sheet, row_num, mapped, source_id, strong_source_id, identity in pending:
            flags = [v for v in mapped["review_flags"].split("; ") if v]
            if source_id and strong_source_id and id_counts[(sheet, source_id.casefold())] > 1:
                flags.append("DUPLICATE_SOURCE_ID")
            mapped["review_flags"] = "; ".join(dict.fromkeys(flags))
            seed = f"{sheet.casefold()}|{identity.casefold()}"
            if mapped["row_type"] == "Asset" and identity_counts[(sheet, identity.casefold())] > 1:
                secondary = "|".join(_text(mapped[k]).casefold() for k in ("asset_name", "tag", "asset_number", "field", "plant", "location", "unit_pack"))
                seed += f"|{secondary}"
                if not secondary.strip("|"):
                    seed += f"|row:{row_num}"
            if mapped["row_type"] != "Asset":
                seed = f"{sheet.casefold()}|row:{row_num}"
            rid = f"{_slug(sheet)}-{hashlib.sha256(seed.encode('utf-8')).hexdigest()[:14].upper()}"
            if rid in used_ids:
                rid = f"{rid}-R{row_num:04d}"
                flags.append("ID_COLLISION_ROW_SUFFIX")
                mapped["review_flags"] = "; ".join(dict.fromkeys(flags))
            used_ids.add(rid)
            out_values = [
                rid, mapped["asset_class"], mapped["asset_name"], mapped["tag"], mapped["asset_number"],
                mapped["field"], mapped["plant"], mapped["location"], mapped["unit_pack"], mapped["in_service"],
                mapped["osi_applicable"], date.fromisoformat(mapped["osi_last_date"]) if mapped["osi_last_date"] else None,
                date.fromisoformat(mapped["osi_next_date"]) if mapped["osi_next_date"] else None, mapped["osi_status"],
                mapped["internal_applicable"], date.fromisoformat(mapped["internal_last_date"]) if mapped["internal_last_date"] else None,
                date.fromisoformat(mapped["internal_next_date"]) if mapped["internal_next_date"] else None, mapped["internal_status"],
                mapped["remarks"], sheet, row_num, mapped["row_type"], mapped["review_flags"], mapped["source_raw"],
            ]
            ws_out.append(out_values)
            totals["total_rows"] += 1
            totals[f"sheet:{sheet}"] += 1
            totals["asset_rows" if mapped["row_type"] == "Asset" else "reference_rows"] += 1
            if mapped["review_flags"]:
                totals["flagged_rows"] += 1

        ws_out.freeze_panes = "A2"
        ws_out.auto_filter.ref = ws_out.dimensions
        ws_out.row_dimensions[1].height = 32
        for cell in ws_out[1]:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill("solid", fgColor="17365D")
            cell.alignment = Alignment(wrap_text=True, vertical="center")
        widths = {1: 20, 2: 20, 3: 40, 4: 20, 5: 22, 6: 16, 7: 20, 8: 24, 9: 20,
                  10: 16, 11: 16, 12: 16, 13: 16, 14: 32, 15: 20, 16: 18, 17: 18,
                  18: 32, 19: 64, 20: 24, 21: 14, 22: 22, 23: 42, 24: 70}
        for col, width in widths.items():
            ws_out.column_dimensions[get_column_letter(col)].width = width
        for row in ws_out.iter_rows(min_row=2):
            for col in (12, 13, 16, 17):
                row[col - 1].number_format = "yyyy-mm-dd"
            row[22].fill = PatternFill("solid", fgColor="FFF2CC") if row[22].value else PatternFill(fill_type=None)
            row[23].alignment = Alignment(wrap_text=True, vertical="top")
            row[22].alignment = Alignment(wrap_text=True, vertical="top")
        if ws_out.max_row > 1:
            dv = DataValidation(type="list", formula1='"Yes,No,Review"', allow_blank=True)
            ws_out.add_data_validation(dv)
            dv.add(f"K2:K{ws_out.max_row}")
            dv.add(f"O2:O{ws_out.max_row}")
            date_dv = DataValidation(type="date", operator="between", formula1="DATE(1900,1,1)", formula2="DATE(2200,12,31)", allow_blank=True)
            ws_out.add_data_validation(date_dv)
            for col in ("L", "M", "P", "Q"):
                date_dv.add(f"{col}2:{col}{ws_out.max_row}")
        wb.save(generated_path)
        _merge_register_into_source(input_path, generated_path, output_path)
    finally:
        if os.path.exists(generated_path):
            os.remove(generated_path)
    return dict(totals)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("input", help="Source .xlsx workbook (never modified)")
    parser.add_argument("output", help="New output .xlsx workbook")
    args = parser.parse_args()
    print(json.dumps(build_register(args.input, args.output), indent=2, ensure_ascii=False))


if __name__ == "__main__":
    main()
