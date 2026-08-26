import sys
import types
import unittest
from pathlib import Path


if "streamlit" not in sys.modules:
    sys.modules["streamlit"] = types.SimpleNamespace(
        cache_data=lambda *args, **kwargs: lambda function: function
    )

import app


SAMPLE_DIR = Path(__file__).resolve().parents[1] / "sample_data"


class CourseSelectionTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.d3 = app.load_master_excel((SAMPLE_DIR / "master_rps_d3_pste.xlsx").read_bytes())
        cls.d4 = app.load_master_excel((SAMPLE_DIR / "master_rps_d4_pste.xlsx").read_bytes())

    def test_bundled_masters_remain_separate(self):
        self.assertEqual(len(self.d3["Master_MK"]), 44)
        self.assertEqual(len(self.d4["Master_MK"]), 51)

    def test_process_control_uses_only_pdf_mapping(self):
        payload = app.build_course_payload(self.d4, "RTE267006-45")
        self.assertEqual(payload["mk"]["nama_mk"], "Sistem Kendali Proses")
        self.assertEqual(
            [row["kode_cpmk"] for row in payload["cpmk"]],
            ["CPMK02.02", "CPMK04.02"],
        )
        self.assertEqual(
            [row["kode_cpl"] for row in payload["cpl"]], ["CPL2", "CPL4"]
        )
        self.assertEqual(len(payload["weekly"]), 16)

    def test_d4_course_structure_is_complete(self):
        courses = self.d4["Master_MK"]
        semester_three = courses.loc[courses["semester"] == 3, "nama_mk"].tolist()
        self.assertIn("Matematika Diskrit", semester_three)
        self.assertIn("Proyek Akhir", courses["nama_mk"].tolist())
        self.assertNotIn("Skripsi", courses["nama_mk"].tolist())
        semester_seven = courses[courses["semester"] == 7]
        self.assertTrue(semester_seven["jenis_mk"].str.contains("Pilihan").any())
        self.assertTrue(semester_seven["jenis_mk"].str.contains("Paket").any())

    def test_duplicate_code_does_not_mix_industrial_project(self):
        process = app.build_course_payload(self.d4, "RTE267006-45")
        project = app.build_course_payload(self.d4, "RTE267006-49")
        self.assertEqual(project["mk"]["nama_mk"], "Proyek Industri")
        self.assertNotEqual(
            {row["kode_cpmk"] for row in process["cpmk"]},
            {row["kode_cpmk"] for row in project["cpmk"]},
        )
        self.assertEqual(len(project["weekly"]), 16)


if __name__ == "__main__":
    unittest.main()
