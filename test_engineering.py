"""
test_engineering.py — stdlib-only test suite for the Master Inspection Plan
engineering logic. Run:  python3 test_engineering.py
Covers: POF/COF engine, remaining-life calculator, overdue/next-due derivation,
estimated-date flags, and the fresh-schema reconciliation fix.
"""
import unittest
import datetime
import sys
import os

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

# Import app module without starting the server (guarded by __main__)
import app


class TestParseFloatOrNone(unittest.TestCase):
    def test_pure_numbers(self):
        self.assertEqual(app._parse_float_or_none("5.71"), 5.71)
        self.assertEqual(app._parse_float_or_none("0.19"), 0.19)
        self.assertEqual(app._parse_float_or_none(3), 3.0)

    def test_junk_text_rejected(self):
        # '1st insp.' must NOT become remaining life of 1 year (POF bug fix)
        self.assertIsNone(app._parse_float_or_none("1st insp."))
        self.assertIsNone(app._parse_float_or_none("tubing material"))
        self.assertIsNone(app._parse_float_or_none(""))
        self.assertIsNone(app._parse_float_or_none(None))
        self.assertIsNone(app._parse_float_or_none("NA"))


class TestPofEngine(unittest.TestCase):
    def base_row(self, **kw):
        row = {"remaining_life": None, "corrosion_rate": None,
               "next_due": None, "is_cui": False}
        row.update(kw)
        return row

    def test_default_pof_with_no_data_is_uncertainty_not_floor(self):
        pof, cof, risk = app.calculate_pof_cof(self.base_row())
        self.assertEqual(pof, 3, "no due date + no RL + no CR => POF 3 (uncertainty)")

    def test_scheduled_healthy_asset_stays_pof2(self):
        future = (datetime.date.today() + datetime.timedelta(days=400)).isoformat()
        pof, _, _ = app.calculate_pof_cof(self.base_row(next_due=future))
        self.assertEqual(pof, 2)

    def test_text_remaining_life_does_not_drive_pof(self):
        # THE BUG: '1st insp.' extracted as RL=1.0 -> POF 5. Now must be ignored.
        future = (datetime.date.today() + datetime.timedelta(days=400)).isoformat()
        pof, _, _ = app.calculate_pof_cof(
            self.base_row(next_due=future, remaining_life="1st insp."))
        self.assertEqual(pof, 2)

    def test_numeric_remaining_life_bands(self):
        future = (datetime.date.today() + datetime.timedelta(days=400)).isoformat()
        self.assertEqual(app.calculate_pof_cof(
            self.base_row(next_due=future, remaining_life="1.5"))[0], 5)
        self.assertEqual(app.calculate_pof_cof(
            self.base_row(next_due=future, remaining_life="4"))[0], 4)
        self.assertEqual(app.calculate_pof_cof(
            self.base_row(next_due=future, remaining_life="8.1"))[0], 3)

    def test_corrosion_rate_bands(self):
        future = (datetime.date.today() + datetime.timedelta(days=400)).isoformat()
        self.assertEqual(app.calculate_pof_cof(
            self.base_row(next_due=future, corrosion_rate="0.6"))[0], 5)
        self.assertEqual(app.calculate_pof_cof(
            self.base_row(next_due=future, corrosion_rate="0.3"))[0], 4)
        self.assertEqual(app.calculate_pof_cof(
            self.base_row(next_due=future, corrosion_rate="0.15"))[0], 3)
        self.assertEqual(app.calculate_pof_cof(
            self.base_row(next_due=future, corrosion_rate="0.05"))[0], 2)

    def test_overdue_bands(self):
        for days, expected in [(200, 5), (70, 4), (10, 3)]:
            nd = (datetime.date.today() - datetime.timedelta(days=days)).isoformat()
            self.assertEqual(app.calculate_pof_cof(
                self.base_row(next_due=nd))[0], expected, f"{days}d overdue")

    def test_cui_adds_one(self):
        future = (datetime.date.today() + datetime.timedelta(days=400)).isoformat()
        self.assertEqual(app.calculate_pof_cof(
            self.base_row(next_due=future, is_cui=True))[0], 3)

    def test_cof_keywords(self):
        self.assertEqual(app.calculate_pof_cof(
            self.base_row(source_sheet="Vessels & TKs"))[1], 5)
        self.assertEqual(app.calculate_pof_cof(
            self.base_row(source_sheet="Coolers"))[1], 3)
        self.assertEqual(app.calculate_pof_cof(
            self.base_row(source_sheet="WD-33 Piping"))[1], 2, "unknown sheet => default")

    def test_risk_matrix_boundaries(self):
        # Coolers => COF 3. POF 5 (RL 1.5y): 5x3=15 AND (pof=5, cof>=3) => HIGH
        self.assertEqual(app.calculate_pof_cof(
            self.base_row(source_sheet="Coolers", remaining_life="1.5"))[2], "HIGH")
        # POF 3 (CR 0.15) x COF 3 = 9 => MEDIUM
        self.assertEqual(app.calculate_pof_cof(
            self.base_row(source_sheet="Coolers", corrosion_rate="0.15"))[2], "MEDIUM")
        # Default COF 2 x POF 2 = 4 => LOW (scheduled asset)
        future = (datetime.date.today() + datetime.timedelta(days=400)).isoformat()
        self.assertEqual(app.calculate_pof_cof(self.base_row(next_due=future))[2], "LOW")


class TestRemainingLifeCalculator(unittest.TestCase):
    def calc(self, **kw):
        return app.api_calculate_remaining_life(kw)

    def test_basic_half_life(self):
        r = self.calc(t_act=8.2, t_prev=9.0, t_min=6.0, years_between=2)
        self.assertTrue(r["success"])
        self.assertAlmostEqual(r["corrosion_rate_mm_yr"], 0.4)
        self.assertAlmostEqual(r["remaining_life_years"], 5.5)
        self.assertAlmostEqual(r["half_life_interval_years"], 2.8)

    def test_zero_corrosion_no_longer_grants_max_interval(self):
        r = self.calc(t_act=8.2, t_prev=8.2, t_min=6.0, years_between=3)
        self.assertTrue(r["success"])
        self.assertEqual(r["half_life_interval_years"], 5.0, "CR=0 => capped at 5y, not 10y")
        self.assertIn("warning", r)
        self.assertIsNotNone(r["warning"])

    def test_short_term_rate_warning(self):
        r = self.calc(t_act=7.9, t_prev=8.0, t_min=6.0, years_between=0.1)
        self.assertTrue(r["success"])
        self.assertIsNotNone(r["warning"], "0.1mm in 0.1y = 1mm/yr needs a caution")

    def test_retirement_condition(self):
        r = self.calc(t_act=5.5, t_prev=6.0, t_min=6.0, years_between=2)
        self.assertTrue(r["success"])
        self.assertEqual(r["remaining_life_years"], 0.0)
        self.assertTrue(r["retire_now"])

    def test_dates_derive_years(self):
        r = self.calc(t_act=8.2, t_prev=9.0, t_min=6.0,
                      d_prev="2021-01-01", d_act="2026-01-01")
        self.assertTrue(r["success"])
        self.assertAlmostEqual(r["corrosion_rate_mm_yr"], 0.16, places=2)

    def test_invalid_input_fails_gracefully(self):
        r = self.calc(t_act="abc", t_min=3)
        self.assertFalse(r["success"])


class TestPlaceholderFlags(unittest.TestCase):
    PLACEHOLDER = "2020-01-01"

    def _dict(self, **over):
        base = {
            "extra_json": "{}", "insulation": None, "operating_temp": None,
            "cui_susceptible": None, "date_osi_next": "2030-01-01",
            "date_internal_next": None, "status_osi_next": "Scheduled",
            "status_internal_next": None, "pof_score": None, "cof_score": None,
            "risk_category": None, "deferral_status": None, "deferral_expiry": None,
            "date_osi_last": None, "date_internal_last": None,
            "corrosion_rate": None, "remaining_life": None, "source_sheet": "Coolers",
            "name": "X", "archived": 0,
        }
        base.update(over)
        return base

    def test_placeholder_flagged_as_estimated(self):
        d = app.row_to_dict(self._dict(date_osi_next=self.PLACEHOLDER))
        self.assertTrue(d["due_date_estimated"])
        self.assertTrue(d["overdue"])

    def test_real_date_not_flagged(self):
        d = app.row_to_dict(self._dict())
        self.assertFalse(d["due_date_estimated"])

    def test_no_dates_asset_gets_uncertainty_pof(self):
        d = app.row_to_dict(self._dict(date_osi_next=None, status_osi_next=None))
        self.assertEqual(d["pof_score"], 3)


class TestFreshSchemaReconciliation(unittest.TestCase):
    """A DB created by app.py alone (no refined import yet) must not 500."""

    def test_refined_table_has_sync_columns(self):
        import sqlite3, tempfile
        with tempfile.TemporaryDirectory() as td:
            db = os.path.join(td, "t.db")
            open(db, "w").close()
            old = app.DB_PATH
            app.DB_PATH = db
            try:
                conn = app.get_conn()
                cols = {r[1] for r in conn.execute("PRAGMA table_info(refined_plan_items)")}
                conn.close()
            finally:
                app.DB_PATH = old
            self.assertIn("synced_to_master", cols)
            self.assertIn("synced_at", cols)

    def test_audit_and_meta_tables_exist(self):
        import sqlite3, tempfile
        with tempfile.TemporaryDirectory() as td:
            db = os.path.join(td, "t.db")
            open(db, "w").close()
            old = app.DB_PATH
            app.DB_PATH = db
            try:
                conn = app.get_conn()
                tables = {r[0] for r in conn.execute(
                    "SELECT name FROM sqlite_master WHERE type='table'")}
                conn.close()
            finally:
                app.DB_PATH = old
            self.assertIn("audit_log", tables)
            self.assertIn("app_meta", tables)


if __name__ == "__main__":
    unittest.main(verbosity=2)
