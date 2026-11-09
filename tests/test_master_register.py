import json
import os
import tempfile
import unittest
import re
import zipfile

import openpyxl

from build_master_register import build_register, classify_row


class MasterRegisterTests(unittest.TestCase):
    def make_workbook(self, path):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Vessels & TKs"
        ws.append(["Update", "2026-01-01"])
        ws.append(["SN", "Field", "Asset Name", "Tag #", "Asset #", "Last Insp. Category", "Last OSI Date", "Next OSI Due Date", "Last Internal Insp. Date", "Next Internal Insp. Date", "Remarks"])
        ws.append([1, "AG", "Test Vessel", "V-1", "OR-1", "OSI", "2020-01-01", "OOS", "2019-01-01", "2029-01-01", "note"])
        ws.append([2, "AG", "Tank", None, None, "2021-02-01", "2026-02-01", "Never", "Next SD", None])
        line = wb.create_sheet("TLs")
        line.append(["Location", "Pipeline Name", "Last Insp", "Next Insp"])
        line.append(["From", "To", None, None])
        line.append(["A", "Trunk line A", "2022-01-01", "W/Eval"])
        notes = wb.create_sheet("GL Lines")
        notes.append(["*", "Global note", None])
        wb.save(path)

    def test_classification_separates_dates_and_preserves_raw_status(self):
        result = classify_row("Vessels & TKs", 3,
                              ["SN", "Field", "Asset Name", "Tag #", "Asset #", "Last Insp. Category", "Last OSI Date", "Next OSI Due Date", "Last Internal Insp. Date", "Next Internal Insp. Date", "Remarks"],
                              [1, "AG", "Test Vessel", "V-1", "OR-1", "OSI", "2020-01-01", "OOS", "2019-01-01", "2029-01-01", "note"],
                              header_row=2)
        self.assertEqual(result["asset_class"], "Vessel")
        self.assertEqual(result["osi_last_date"], "2020-01-01")
        self.assertEqual(result["osi_next_date"], None)
        self.assertEqual(result["osi_status"], "OOS")
        self.assertEqual(result["internal_applicable"], "Yes")
        raw = json.loads(result["source_raw"])
        self.assertEqual(raw["SN"], 1)
        self.assertEqual(raw["Last Insp. Category"], "OSI")
        self.assertEqual(raw["Last OSI Date"], "2020-01-01")

    def test_register_retains_every_nonempty_source_row_and_flags_reference_rows(self):
        with tempfile.TemporaryDirectory() as tmp:
            src, out = os.path.join(tmp, "source.xlsx"), os.path.join(tmp, "register.xlsx")
            self.make_workbook(src)
            counts = build_register(src, out)
            wb = openpyxl.load_workbook(out, data_only=True)
            reg = wb["Master Register"]
            rows = list(reg.iter_rows(values_only=True))
            headers = rows[0]
            records = [dict(zip(headers, row)) for row in rows[1:]]
            expected = sum(1 for s in openpyxl.load_workbook(src, data_only=True).worksheets
                           for row in s.iter_rows(values_only=True) if any(v is not None for v in row))
            self.assertEqual(len(records), expected)
            self.assertEqual(len({r["Asset ID"] for r in records}), expected)
            self.assertTrue(any("REFERENCE_ROW" in (r["Review Flags"] or "") for r in records))
            self.assertTrue(any(r["Asset Class"] == "Trunkline" and r["Internal Applicable"] == "No" for r in records))
            self.assertEqual(counts["total_rows"], expected)
            self.assertTrue(reg.freeze_panes)
            self.assertTrue(reg.auto_filter.ref)

    def test_ids_repeat_across_runs(self):
        with tempfile.TemporaryDirectory() as tmp:
            src, out1, out2 = [os.path.join(tmp, n) for n in ("source.xlsx", "one.xlsx", "two.xlsx")]
            self.make_workbook(src)
            build_register(src, out1)
            build_register(src, out2)
            def ids(path):
                ws = openpyxl.load_workbook(path, data_only=True)["Master Register"]
                return [r[0] for r in ws.iter_rows(min_row=2, values_only=True)]
            self.assertEqual(ids(out1), ids(out2))

    def test_pack_heading_is_preserved_as_reference_not_asset(self):
        result = classify_row("GP Piping", 10, ["SN#", "Item", "Pack#", "Last Insp", "Next Insp"],
                              ["PACK 01A Inlet gas", None, None, None, None], header_row=8)
        self.assertEqual(result["row_type"], "Review / Reference")
        self.assertIn("REFERENCE_ROW", result["review_flags"])

    def test_gas_lift_line_uses_line_name_as_source_identifier(self):
        result = classify_row("GL Lines", 18, ["SN", "Well", "Gas Lift Line", "Last Insp. Date", "Next Insp. Date"],
                              [1, "AG-2", "AG-2", "June -19", "2019-12-02"], header_row=17)
        self.assertEqual(result["asset_name"], "AG-2")
        self.assertNotIn("GENERATED_ID_SOURCE_ID_MISSING", result["review_flags"])

    def test_asset_number_prefers_asset_number_column_over_oracle_description(self):
        result = classify_row("Vessels & TKs", 8,
                              ["Asset Name", "Tag #", "Asset #", "Full asset Description (Oracle Name)", "Last OSI Date"],
                              ["Vessel", "V-1", "HEST-001", "Full equipment description", "2020-01-01"], 7)
        self.assertEqual(result["asset_number"], "HEST-001")

    def test_last_insp_date_does_not_match_leak_count_header(self):
        result = classify_row("FLs", 5,
                              ["Well", "Last Insp", "Next Insp", "Number of Leaks after last Insp date"],
                              ["AG-102", "2017-05-16", "2021-05-25", 1], 3)
        self.assertEqual(result["osi_last_date"], "2017-05-16")
        self.assertEqual(result["osi_next_date"], "2021-05-25")

    def test_generated_asset_ids_survive_inserting_source_rows(self):
        with tempfile.TemporaryDirectory() as tmp:
            src1, src2, out1, out2 = [os.path.join(tmp, n) for n in ("one.xlsx", "two.xlsx", "one-out.xlsx", "two-out.xlsx")]
            self.make_workbook(src1)
            wb = openpyxl.load_workbook(src1)
            ws = wb["Vessels & TKs"]
            ws.insert_rows(3)
            ws.cell(3, 1, "Added note")
            wb.save(src2)
            build_register(src1, out1)
            build_register(src2, out2)
            def asset_id(path, tag):
                ws = openpyxl.load_workbook(path, data_only=True)["Master Register"]
                h = [c.value for c in ws[1]]
                ix = {v:i for i,v in enumerate(h)}
                return next(r[ix["Asset ID"]] for r in ws.iter_rows(min_row=2, values_only=True) if r[ix["Tag"]] == tag)
            self.assertEqual(asset_id(out1, "V-1"), asset_id(out2, "V-1"))

    def test_month_year_date_is_explicitly_flagged_as_assumed_day(self):
        result = classify_row("GL Lines", 18,
                              ["Well", "Gas Lift Line", "Last Insp. Date", "Next Insp. Date"],
                              ["AG-2", "AG-2", "June -19", "2019-12-02"], 17)
        self.assertEqual(result["osi_last_date"], "2019-06-01")
        self.assertIn("PARTIAL_DATE_ASSUMED_DAY_01", result["review_flags"])

    def test_ambiguous_numeric_date_is_not_silently_day_month_swapped(self):
        result = classify_row("FLs", 5, ["Well", "Last Insp", "Next Insp"],
                              ["AG-2", "01/02/2025", "2026-01-01"], 3)
        self.assertIsNone(result["osi_last_date"])
        self.assertIn("AMBIGUOUS_DATE_ORDER_REVIEW", result["review_flags"])

    def test_name_derived_asset_id_is_flagged_for_manual_retention(self):
        result = classify_row("OP Piping", 7, ["Report#", "Item", "Next insp"],
                              [1, "Oil line to separator", "2027-01-01"], 5)
        self.assertIn("ID_FROM_ASSET_NAME_REVIEW", result["review_flags"])

    def test_future_update_metadata_is_flagged(self):
        from datetime import date
        result = classify_row("FLs", 1, ["Well", "In-service?"],
                              ["Updated on", date(2100, 1, 1)], 3)
        self.assertIn("FUTURE_UPDATE_DATE_REVIEW", result["review_flags"])

    def test_uncertain_vessel_sheet_item_is_not_assumed_internal(self):
        result = classify_row("Vessels & TKs", 21,
                              ["SN", "Asset Name", "Tag #", "Asset #", "Last Internal Insp. Date", "Next Internal Insp. Date"],
                              [14, "Process Heater 1401", "H-1401", "A-1401", "2020-01-01", "2030-01-01"], 7)
        self.assertEqual(result["asset_class"], "Other Equipment")
        self.assertEqual(result["internal_applicable"], "Review")
        self.assertIn("INTERNAL_APPLICABILITY_REVIEW", result["review_flags"])
        self.assertEqual(result["internal_last_date"], "2020-01-01")
        self.assertEqual(result["internal_next_date"], "2030-01-01")

    def test_shell_and_tube_hex_is_vessel_class_with_internal_inspection(self):
        result = classify_row("Vessels & TKs", 8,
                              ["Asset Name", "Tag #", "Asset #", "Last Internal Insp. Date"],
                              ["Gas/Gas HEX-1301-A/B [Shell & Tube HEX]", "E-1301 A/B", "HEST-001", "2024-01-01"], 7)
        self.assertEqual(result["asset_class"], "Vessel")
        self.assertEqual(result["internal_applicable"], "Yes")

    def test_note_and_from_to_rows_are_reference_rows(self):
        epf = classify_row("EPFs", 68, ["NO", "Item", "Last Insp", "Next Insp"],
                           [None, "Note: As per attached mail", None, None], 4)
        tl = classify_row("TLs", 7, ["Location", "", "", "Pipeline Name"],
                          ["From", "To", None, None], 6)
        self.assertEqual(epf["row_type"], "Review / Reference")
        self.assertEqual(tl["row_type"], "Review / Reference")
        turbines = classify_row("Turbines Piping", 6, ["Rep.#", "Item", "Last Inspection", "Next Inspection"],
                                 ["GP Comp. #1845", None, None, None], 5)
        self.assertEqual(turbines["row_type"], "Review / Reference")

    def test_asset_without_due_date_or_status_is_flagged(self):
        result = classify_row("FLs", 20, ["Well", "Last Insp", "Next Insp"], ["AG-1", "2020-01-01", None], 3)
        self.assertIn("OSI_DUE_INFORMATION_MISSING", result["review_flags"])

    def test_generated_date_cells_are_excel_dates_with_validation(self):
        with tempfile.TemporaryDirectory() as tmp:
            src, out = os.path.join(tmp, "source.xlsx"), os.path.join(tmp, "out.xlsx")
            self.make_workbook(src)
            build_register(src, out)
            wb = openpyxl.load_workbook(out, data_only=True)
            ws = wb["Master Register"]
            self.assertEqual(ws["L4"].value.strftime("%Y-%m-%d"), "2020-01-01")
            self.assertTrue(any(dv.type == "date" for dv in ws.data_validations.dataValidation))

    def test_generator_refuses_to_overwrite_existing_review_workbook(self):
        with tempfile.TemporaryDirectory() as tmp:
            src, out = os.path.join(tmp, "source.xlsx"), os.path.join(tmp, "out.xlsx")
            self.make_workbook(src)
            build_register(src, out)
            before = os.path.getsize(out)
            with self.assertRaises(FileExistsError):
                build_register(src, out)
            self.assertEqual(os.path.getsize(out), before)

    def test_duplicate_source_rows_still_receive_unique_ids(self):
        with tempfile.TemporaryDirectory() as tmp:
            src, out = os.path.join(tmp, "source.xlsx"), os.path.join(tmp, "out.xlsx")
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "FLs"
            ws.append(["Well", "Last Insp", "Next Insp"])
            ws.append(["AG-1", "2020-01-01", "2025-01-01"])
            ws.append(["AG-1", "2020-01-01", "2025-01-01"])
            wb.save(src)
            build_register(src, out)
            reg = openpyxl.load_workbook(out, data_only=True)["Master Register"]
            ids = [row[0] for row in reg.iter_rows(min_row=2, values_only=True)]
            self.assertEqual(len(ids), len(set(ids)))

    def test_formula_cached_value_is_retained_in_source_sheet(self):
        with tempfile.TemporaryDirectory() as tmp:
            src, out = os.path.join(tmp, "formula.xlsx"), os.path.join(tmp, "formula-out.xlsx")
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "TLs"
            ws["A1"] = "Summary"
            ws["A2"] = "Total"
            ws["B2"] = "=SUM(1,2)"
            wb.save(src)
            temp_src = os.path.join(tmp, "formula-cached.xlsx")
            with zipfile.ZipFile(src) as zin, zipfile.ZipFile(temp_src, "w", zipfile.ZIP_DEFLATED) as zout:
                for info in zin.infolist():
                    payload = zin.read(info.filename)
                    if info.filename == "xl/worksheets/sheet1.xml":
                        payload = payload.replace(b"<f>SUM(1,2)</f><v /></c>", b"<f>SUM(1,2)</f><v>3</v></c>")
                    zout.writestr(info, payload)
            build_register(temp_src, out)
            result = openpyxl.load_workbook(out, data_only=True)
            self.assertEqual(result["TLs"]["B2"].value, 3)


if __name__ == "__main__":
    unittest.main()
