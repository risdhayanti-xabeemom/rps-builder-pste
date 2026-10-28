import io
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
import rangkuman


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
            # Rangkaian Listrik 1 (x1008) hanya berubah di angkatan 2026.
            return [(row["kode_mk"][5:], row["kode_cpmk"], row["bobot_cpmk_mk_persen"]) for row in frame.to_dict("records") if row["kode_mk"][5:] != "1008"]
        self.assertEqual(weights("2025"), weights("2026"))
        fisika = [w for w in weights("2025") if w[0] == "1001"]
        self.assertEqual(fisika, [("1001", "CPMK1.1", 52.17), ("1001", "CPMK1.2", 47.83)])

    def test_komponen_pengukuran_without_cpl08_from_2025(self):
        komponen = {
            "CPMK1.1": 14.81, "CPMK1.2": 13.58, "CPMK3.1": 17.28, "CPMK3.2": 16.05,
            "CPMK4.1": 4.94, "CPMK4.2": 14.81, "CPMK4.3": 18.53,
        }
        magang = {
            "CPMK10.1": 7.41, "CPMK10.3": 3.7, "CPMK6.1": 7.41, "CPMK6.3": 16.67, "CPMK7.1": 11.11,
            "CPMK7.2": 1.85, "CPMK8.1": 25.92, "CPMK8.3": 14.81, "CPMK9.1": 5.56, "CPMK9.2": 5.56,
        }
        for cohort in ("2025", "2026"):
            workbook = app.load_program_workbook("D3 Teknik Elektro", cohort)
            frame = workbook["Master_CPMK"]
            kp, mg = f"REC{cohort[-2:]}1002", f"REC{cohort[-2:]}5001"
            got = lambda code: dict(zip(frame[frame["kode_mk"] == code]["kode_cpmk"], frame[frame["kode_mk"] == code]["bobot_cpmk_mk_persen"]))
            self.assertEqual(got(kp), komponen)
            self.assertEqual(got(mg), magang)
            mapping = workbook["Mapping_MK_CPL"]
            self.assertEqual(sorted(mapping[mapping["kode_mk"] == kp]["kode_cpl"]), ["CPL1", "CPL3", "CPL4"])
            weeks = workbook["RPS_Pertemuan"]
            self.assertNotIn("CPMK8.1", set(weeks[weeks["kode_mk"] == kp]["kode_cpmk"]))

    def test_d3_2024_follows_obaobe_rencana_evaluasi(self):
        frame = app.load_program_workbook("D3 Teknik Elektro", "2024")["Master_CPMK"]
        got = lambda code: dict(zip(frame[frame["kode_mk"] == code]["kode_cpmk"], frame[frame["kode_mk"] == code]["bobot_cpmk_mk_persen"]))
        self.assertEqual(got("REC241001"), {"CPMK1.1": 52.17, "CPMK1.2": 47.83})
        # Angkatan 2024 tidak ikut pemindahan CPL08: Komponen dan Pengukuran tetap memuat CPMK8.1.
        self.assertEqual(got("REC241002"), {
            "CPMK1.1": 13.64, "CPMK1.2": 12.5, "CPMK3.1": 15.91, "CPMK3.2": 14.77,
            "CPMK4.1": 4.55, "CPMK4.2": 13.64, "CPMK4.3": 17.04, "CPMK8.1": 7.95,
        })
        self.assertEqual(got("REC246003"), {"CPMK10.2": 8.7, "CPMK8.1": 30.43, "CPMK8.2": 26.09, "CPMK8.3": 34.78})

    def test_elektronika_analog_2_without_cpl5_from_2025(self):
        expected = {"CPMK1.1": 17.0, "CPMK2.2": 20.0, "CPMK2.3": 5.0, "CPMK3.1": 19.0, "CPMK4.2": 17.0, "CPMK4.3": 22.0}
        for cohort in ("2025", "2026"):
            workbook = app.load_program_workbook("D3 Teknik Elektro", cohort)
            code = f"REC{cohort[-2:]}3001"
            frame = workbook["Master_CPMK"]
            frame = frame[frame["kode_mk"] == code]
            self.assertEqual(dict(zip(frame["kode_cpmk"], frame["bobot_cpmk_mk_persen"])), expected)
            mapping = workbook["Mapping_MK_CPL"]
            self.assertNotIn("CPL5", set(mapping[mapping["kode_mk"] == code]["kode_cpl"]))
            weekly = workbook["RPS_Pertemuan"]
            weekly = weekly[weekly["kode_mk"] == code]
            self.assertAlmostEqual(weekly["bobot"].map(app.as_float).sum(), 100, places=2)
            self.assertFalse({"CPMK5.1", "CPMK5.3"} & set(weekly["kode_cpmk"]))
        # Angkatan 2024 tidak ikut perubahan.
        old = app.load_program_workbook("D3 Teknik Elektro", "2024")["Master_CPMK"]
        self.assertIn("CPMK5.1", set(old[old["kode_mk"] == "REC243001"]["kode_cpmk"]))

    def test_rangkaian_listrik_1_cpl1_only_in_2026(self):
        workbook = app.load_program_workbook("D3 Teknik Elektro", "2026")
        frame = workbook["Master_CPMK"]
        frame = frame[frame["kode_mk"] == "REC261008"]
        self.assertEqual(dict(zip(frame["kode_cpmk"], frame["bobot_cpmk_mk_persen"])), {"CPMK1.1": 52.17, "CPMK1.2": 47.83})
        mapping = workbook["Mapping_MK_CPL"]
        self.assertEqual(list(mapping[mapping["kode_mk"] == "REC261008"]["kode_cpl"]), ["CPL1"])
        old = app.load_program_workbook("D3 Teknik Elektro", "2025")["Master_CPMK"]
        self.assertEqual(len(old[old["kode_mk"] == "REC251008"]), 7)


class RangkumanTests(unittest.TestCase):
    def test_summary_covers_every_program_and_cohort(self):
        for program in ("D3 Teknik Elektro", "D4 Teknik Elektronika"):
            for cohort in ("2023", "2024", "2025", "2026"):
                workbook = app.load_program_workbook(program, cohort)
                sheets = rangkuman.build_summary(workbook)
                courses = sheets["Ringkasan MK"]
                self.assertEqual(len(courses), len(workbook["Master_MK"]))
                self.assertTrue((courses["total_bobot_persen"].round(2) == 100).all(), (program, cohort))
                self.assertEqual(len(sheets["CPMK dan Bobot"]), len(workbook["Master_CPMK"]))
                self.assertEqual(len(sheets["Ringkasan CPL"]), 10)
                if (program, cohort) != ("D3 Teknik Elektro", "2023"):  # master REC23 tidak punya CPMK CPL7
                    self.assertTrue((sheets["Ringkasan CPL"]["jumlah_cpmk"] > 0).all(), (program, cohort))

    def test_excel_export_has_all_sheets(self):
        workbook = app.load_program_workbook("D4 Teknik Elektronika", "2026")
        sheets = rangkuman.build_summary(workbook)
        loaded = pd.read_excel(io.BytesIO(rangkuman.summary_to_excel_bytes(sheets)), sheet_name=None)
        self.assertEqual(list(loaded), list(sheets))

    def test_menu_is_wired_in_app(self):
        self.assertIn("Rangkuman CPL dan CPMK", APP_SOURCE)


if __name__ == "__main__":
    unittest.main()
