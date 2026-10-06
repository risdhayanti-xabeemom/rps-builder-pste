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
                         {"CPMK01.01": 43.76, "CPMK04.02": 28.12, "CPMK08.02": 28.12})


class ProgramHeaderTests(unittest.TestCase):
    def _header_lines(self, program, cohort):
        import io
        from docx import Document
        workbook = app.load_program_workbook(program, cohort)
        payload = app.build_course_payload(workbook, app.dataframe_course_keys(workbook["Master_MK"]).iloc[0])
        cpmk = app.locked_cpmk_frame(payload, payload["mk"]["kode_mk"], payload["mk"].get("id_penawaran", ""))
        weekly = app.records_to_editor(payload["weekly"], app.RPS_WEEKLY_COLUMNS)
        refs = app.records_to_editor(payload["references"], ["kode_mk", "referensi"])
        context = app.make_context(payload, "", "", cpmk, weekly, refs)
        document = Document(io.BytesIO(app.render_docx(context, payload, cpmk, weekly)))
        return sorted({p.text for p in app.iter_document_paragraphs(document) if "PROGRAM STUDI" in p.text})

    def test_d4_word_header_says_d4(self):
        self.assertEqual(self._header_lines("D4 Teknik Elektronika", "2026"), ["PROGRAM STUDI : D4 TEKNIK ELEKTRONIKA"])

    def test_d3_word_header_unchanged(self):
        self.assertEqual(self._header_lines("D3 Teknik Elektro", "2026"), ["PROGRAM STUDI : D3 TEKNIK ELEKTRONIKA"])

    def test_program_header_text(self):
        self.assertEqual(app.program_header_text("D4 Teknik Elektronika"), "PROGRAM STUDI : D4 TEKNIK ELEKTRONIKA")
        self.assertEqual(app.program_header_text("D-IV Teknik Elektronika"), "PROGRAM STUDI : D4 TEKNIK ELEKTRONIKA")
        self.assertEqual(app.program_header_text("D3 Teknik Elektro"), "PROGRAM STUDI : D3 TEKNIK ELEKTRONIKA")


class D3CpmkWeightTests(unittest.TestCase):
    def test_d3_weights_follow_bobot_table(self):
        for cohort in ("2024", "2025", "2026"):
            workbook = app.load_program_workbook("D3 Teknik Elektro", cohort)
            cpmk = workbook["Master_CPMK"]
            weekly = workbook["RPS_Pertemuan"].copy()
            weekly["bobot"] = weekly["bobot"].map(app.as_float)
            sums = weekly.groupby(["kode_mk", "kode_cpmk"])["bobot"].sum()
            for row in cpmk.to_dict("records"):
                self.assertAlmostEqual(sums[(row["kode_mk"], row["kode_cpmk"])], row["bobot_cpmk_mk_persen"], places=2)
            for total in weekly.groupby("kode_mk")["bobot"].sum():
                self.assertAlmostEqual(total, 100, places=2)

    def test_d3_2026_matches_2025(self):
        def weights(cohort):
            frame = app.load_program_workbook("D3 Teknik Elektro", cohort)["Master_CPMK"]
            return [(row["kode_mk"][5:], row["kode_cpmk"], row["bobot_cpmk_mk_persen"]) for row in frame.to_dict("records")]
        self.assertEqual(weights("2025"), weights("2026"))
        fisika = [w for w in weights("2025") if w[0] == "1001"]
        self.assertEqual(fisika, [("1001", "CPMK1.1", 52.17), ("1001", "CPMK1.2", 47.83)])


if __name__ == "__main__":
    unittest.main()
