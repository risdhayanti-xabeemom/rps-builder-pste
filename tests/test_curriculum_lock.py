import sys
import types
import unittest
from pathlib import Path

import pandas as pd


if "streamlit" not in sys.modules:
    sys.modules["streamlit"] = types.SimpleNamespace(
        cache_data=lambda *args, **kwargs: lambda function: function
    )

import app


APP_SOURCE = (Path(__file__).resolve().parents[1] / "app.py").read_text(encoding="utf-8")


class CurriculumLockTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        workbook = app.load_program_workbook("D4 Teknik Elektronika", "2026")
        key = workbook["Master_MK"].loc[workbook["Master_MK"]["nama_mk"] == "Fisika", "id_penawaran"].iloc[0]
        cls.payload = app.build_course_payload(workbook, key)

    def test_no_master_upload_for_lecturers(self):
        self.assertNotIn("file_uploader", APP_SOURCE)

    def test_locked_weekly_columns_are_restored_from_master(self):
        master = self.payload["weekly"]
        edited = app.records_to_editor(master, app.RPS_WEEKLY_COLUMNS)
        edited.loc[0, "kode_cpmk"] = "CPMK99.99"
        edited.loc[1, "bobot"] = 50
        edited.loc[2, "materi"] = "Materi versi dosen"
        result = app.enforce_locked_weekly(edited, master, "RTE261002", "RTE261002-02")
        expected = app.records_to_editor(master, app.RPS_WEEKLY_COLUMNS)
        self.assertEqual(result["kode_cpmk"].tolist(), expected["kode_cpmk"].tolist())
        self.assertEqual(result["bobot"].tolist(), expected["bobot"].tolist())
        self.assertEqual(result.loc[2, "materi"], "Materi versi dosen")

    def test_cpmk_frame_matches_master(self):
        frame = app.locked_cpmk_frame(self.payload, "RTE261002", "RTE261002-02")
        self.assertEqual(frame["kode_cpmk"].tolist(), ["CPMK01.01", "CPMK04.02", "CPMK08.02"])

    def test_cpmk_weight_summary_totals_100(self):
        summary = app.cpmk_weight_summary(self.payload["weekly"])
        self.assertEqual(dict(zip(summary["Kode CPMK"], summary["Bobot (%)"])),
                         {"CPMK01.01": 26.0, "CPMK04.02": 49.0, "CPMK08.02": 25.0})


if __name__ == "__main__":
    unittest.main()
