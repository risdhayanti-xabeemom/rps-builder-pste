import sys
import types
import unittest
from pathlib import Path


if "streamlit" not in sys.modules:
    sys.modules["streamlit"] = types.SimpleNamespace(
        cache_data=lambda *args, **kwargs: lambda function: function
    )

import app
from curriculum import load_catalog


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


class MultiCurriculumTests(unittest.TestCase):
    EXPECTED_COUNTS = {
        ("D3 Teknik Elektro", "2024"): 36,
        ("D3 Teknik Elektro", "2025"): 35,
        ("D3 Teknik Elektro", "2026"): 35,
        ("D4 Teknik Elektronika", "2023"): 57,
        ("D4 Teknik Elektronika", "2024"): 51,
        ("D4 Teknik Elektronika", "2025"): 51,
        ("D4 Teknik Elektronika", "2026"): 51,
    }

    def test_all_active_curricula_use_exact_source_codes(self):
        catalog = load_catalog()
        for (program, cohort), expected_count in self.EXPECTED_COUNTS.items():
            with self.subTest(program=program, cohort=cohort):
                workbook = app.load_program_workbook(program, cohort)
                level = program.split()[0]
                expected_codes = [
                    row["kode_mk"] for row in catalog["curricula"][f"{level}-{cohort}"]
                ]
                self.assertEqual(len(workbook["Master_MK"]), expected_count)
                self.assertEqual(workbook["Master_MK"]["kode_mk"].tolist(), expected_codes)
                self.assertEqual(app.validate_workbook_schema(workbook), [])

    def test_known_historical_code_exceptions_are_preserved(self):
        d4_2023 = app.load_program_workbook("D4 Teknik Elektronika", "2023")["Master_MK"]
        d4_2024 = app.load_program_workbook("D4 Teknik Elektronika", "2024")["Master_MK"]
        self.assertEqual(
            d4_2023.loc[d4_2023["nama_mk"] == "Proyek Industri", "kode_mk"].iloc[0],
            "RTE227011",
        )
        self.assertEqual(
            d4_2024.loc[d4_2024["nama_mk"] == "Proyek Akhir", "kode_mk"].iloc[0],
            "RTE218001",
        )
        self.assertEqual(
            d4_2024.loc[d4_2024["nama_mk"].str.lower() == "seminar hasil", "kode_mk"].iloc[0],
            "RTE218002",
        )

    def test_d3_agama_is_two_credit_theory(self):
        for cohort in ("2024", "2025", "2026"):
            courses = app.load_program_workbook("D3 Teknik Elektro", cohort)["Master_MK"]
            agama = courses[courses["nama_mk"] == "Agama"].iloc[0]
            self.assertEqual(float(agama["sks_teori"]), 2)
            self.assertEqual(float(agama["sks_praktek"]), 0)
            self.assertEqual(float(agama["jam_teori_minggu"]), 2)

    def test_d4_2026_process_control_keeps_official_mapping(self):
        workbook = app.load_program_workbook("D4 Teknik Elektronika", "2026")
        courses = workbook["Master_MK"]
        process_key = courses.loc[
            courses["nama_mk"] == "Sistem Kendali Proses", "id_penawaran"
        ].iloc[0]
        payload = app.build_course_payload(workbook, process_key)
        self.assertEqual([row["kode_cpmk"] for row in payload["cpmk"]], ["CPMK02.02", "CPMK04.02"])
        self.assertEqual([row["kode_cpl"] for row in payload["cpl"]], ["CPL2", "CPL4"])

    def test_duplicate_2026_code_uses_distinct_offering_ids(self):
        workbook = app.load_program_workbook("D4 Teknik Elektronika", "2026")
        duplicate = workbook["Master_MK"][workbook["Master_MK"]["kode_mk"] == "RTE267006"]
        self.assertEqual(duplicate["nama_mk"].tolist(), ["Sistem Kendali Proses", "Proyek Industri"])
        self.assertEqual(duplicate["id_penawaran"].nunique(), 2)

    def test_d4_weekly_material_comes_from_corrected_syllabus(self):
        workbook = app.load_program_workbook("D4 Teknik Elektronika", "2026")
        course = workbook["Master_MK"][workbook["Master_MK"]["nama_mk"] == "Sistem Kendali Proses"].iloc[0]
        payload = app.build_course_payload(workbook, course["id_penawaran"])
        materials = " ".join(str(row["materi"]) for row in payload["weekly"])
        self.assertIn("Pemodelan CSTR", materials)
        self.assertNotIn("ruang lingkup mata kuliah", materials.lower())
        self.assertEqual(len(payload["weekly"]), 16)


class IntegrityValidationTests(unittest.TestCase):
    def test_legacy_d3_cpmk4_is_remapped_and_marked(self):
        workbook = app.load_master_excel(
            (SAMPLE_DIR / "master_rps_d3_pste.xlsx").read_bytes()
        )
        marked = workbook["RPS_Pertemuan"]["catatan_integritas"] == "AUTO_REMAP_REVIEW_DOSEN"
        self.assertEqual(int(marked.sum()), 156)
        valid_pairs = {
            (app.course_key_from_row(row), str(row["kode_cpmk"]))
            for row in workbook["Master_CPMK"].to_dict("records")
        }
        for row in workbook["RPS_Pertemuan"].loc[marked].to_dict("records"):
            self.assertIn((app.course_key_from_row(row), row["kode_cpmk"]), valid_pairs)

    def test_all_generated_curricula_pass_blocking_integrity_checks(self):
        for program, cohorts in app.CURRICULUM_OPTIONS.items():
            for cohort in cohorts:
                with self.subTest(program=program, cohort=cohort):
                    workbook = app.load_program_workbook(program, cohort)
                    self.assertEqual(app.validate_workbook_integrity(workbook), [])

    def test_mapping_conflict_is_blocking(self):
        workbook = app.load_program_workbook("D3 Teknik Elektro", "2026")
        workbook["Mapping_MK_CPL"] = workbook["Mapping_MK_CPL"].copy()
        workbook["Mapping_MK_CPL"].loc[0, "kode_cpl"] = "CPL10"
        errors = app.validate_workbook_integrity(workbook)
        self.assertTrue(any("Konflik CPL" in error for error in errors))

    def test_direct_cpl_must_match_parent_ik(self):
        workbook = app.load_program_workbook("D4 Teknik Elektronika", "2026")
        workbook["Master_CPMK"] = workbook["Master_CPMK"].copy()
        workbook["Master_CPMK"].loc[0, "kode_cpl"] = "CPL10"
        errors = app.validate_workbook_integrity(workbook)
        self.assertTrue(any("IK induknya" in error for error in errors))

    def test_duplicate_course_requires_offering_id_in_every_related_sheet(self):
        workbook = app.load_program_workbook("D4 Teknik Elektronika", "2026")
        duplicate_row = workbook["RPS_Pertemuan"]["kode_mk"] == "RTE267006"
        workbook["RPS_Pertemuan"] = workbook["RPS_Pertemuan"].copy()
        workbook["RPS_Pertemuan"].loc[duplicate_row, "id_penawaran"] = ""
        errors = app.validate_workbook_integrity(workbook)
        self.assertTrue(any("id_penawaran` kosong" in error for error in errors))

    def test_ik_normalization_removes_zero_padding(self):
        self.assertEqual(app.normalize_ik_code("IK01.01"), "IK1.1")
        self.assertEqual(app.normalize_ik_code("IK1.1"), "IK1.1")


if __name__ == "__main__":
    unittest.main()
