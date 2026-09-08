from __future__ import annotations

import io
import json
import re
from collections import Counter
from difflib import SequenceMatcher
from pathlib import Path
from typing import Any

import pandas as pd


CATALOG_PATH = Path(__file__).parent / "data" / "curriculum_catalog.json"
CURRICULUM_OPTIONS = {
    "D3 Teknik Elektro": ["2024", "2025", "2026"],
    "D4 Teknik Elektronika": ["2023", "2024", "2025", "2026"],
}


def load_catalog() -> dict[str, Any]:
    return json.loads(CATALOG_PATH.read_text(encoding="utf-8"))


def curriculum_key(program: str, cohort: str) -> str:
    level = "D4" if program.startswith("D4") else "D3"
    key = f"{level}-{cohort}"
    if cohort not in CURRICULUM_OPTIONS[program]:
        raise ValueError(f"Angkatan {cohort} tidak tersedia untuk {program}.")
    return key


def _norm(value: Any) -> str:
    text = str(value or "").casefold()
    text = text.replace("praktek", "praktik").replace("praktikum", "praktik")
    text = text.replace("mesin-mesin", "mesin").replace("mesin mesin", "mesin")
    return re.sub(r"[^a-z0-9]+", " ", text).strip()


ALIASES = {
    "fisika": "fisika dasar",
    "fisika elektronika": "fisika dasar",
    "fisika elektronika 1": "fisika dasar",
    "fisika terapan": "fisika dasar",
    "gambar teknik": "gambar listrik",
    "bengkel mekanik": "praktik bengkel mekanik",
    "praktek bengkel mekanik": "praktik bengkel mekanik",
    "bengkel elektronika": "praktik bengkel elektronika",
    "praktek bengkel elektronika": "praktik bengkel elektronika",
    "algoritma pemrograman": "praktik algoritma dan pemrograman",
    "algoritma dan pemrograman": "praktik algoritma dan pemrograman",
    "metode numerik dan pemrograman": "praktik algoritma dan pemrograman",
    "praktik metode numerik dan pemrograman": "praktik algoritma dan pemrograman",
    "sensor aktuator": "sensor dan aktuator",
    "praktik sensor aktuator": "praktik sensor dan aktuator",
    "rangkaian logika digital": "elektronika digital 1",
    "teknik digital 1": "elektronika digital 1",
    "teknik digital 2": "elektronika digital 2",
    "elektronika digital": "elektronika digital 2",
    "sistem mikrokontroler": "sistem mikrokontroler 1",
    "teknologi mikrokontroler": "sistem mikrokontroler 2",
    "sistem kendali kontinyu": "sistem kendali analog",
    "sistem kendali 1": "sistem kendali analog",
    "sistem kendali diskrit": "sistem kendali digital",
    "sistem kendali 2": "sistem kendali digital",
    "sistem kendali cerdas": "kendali cerdas",
    "sistem dcs dan scada": "dcs dan scada",
    "mesin listrik": "mesin listrik",
    "mesin listrik listrik": "mesin listrik",
    "perawatan perbaikan": "perawatan dan perbaikan",
    "magang industri": "praktik kerja industri prakerin",
    "proyek akhir": "skripsi",
    "seminar hasil": "skripsi",
    "quality management system": "manajemen industri",
    "karya tulis ilmiah": "bahasa indonesia",
    "machine learning": "kendali cerdas",
    "pengolahan sinyal multimedia": "pengolahan sinyal digital",
    "proyek industri": "desain proyek",
    "elektrokimia": "fisika dasar",
    "elektro kimia": "fisika dasar",
    "persamaan diferensial": "matematika 2",
    "matematika diskrit": "matematika 2",
    "matematika 3": "matematika 2",
    "matematika terapan 1": "matematika 2",
    "keselamatan dan kesehatan kerja lingkungan": "keselamatan dan kesehatan kerja k3",
    "bahasa inggris teknik": "bahasa inggris 1",
    "bahasa inggris untuk komunikasi": "bahasa inggris 2",
    "sistem embedded": "sistem mikrokontroler 2",
}

BASE_ALIASES = {
    "fisika dasar": "fisika",
    "fisika elektronika": "fisika dasar",
    "fisika elektronika 1": "fisika",
    "praktik rangkaian listrik 1": "rangkaian listrik 1",
    "praktik rangkaian listrik 2": "rangkaian listrik 2",
    "praktik elektronika analog 1": "elektronika analog 1",
    "praktik elektronika analog 2": "elektronika analog 2",
    "praktik elektronika digital 1": "elektronika digital",
    "praktik sensor aktuator": "sensor aktuator",
    "praktik sistem mikrokontroler 1": "sistem mikrokontroler",
    "teknik digital 1": "elektronika digital",
    "teknik digital 2": "elektronika digital",
    "elektronika digital 1": "elektronika digital",
    "elektronika digital 2": "elektronika digital",
    "sistem mikrokontroler 1": "sistem mikrokontroler",
    "sistem mikrokontroler 2": "teknologi mikrokontroler",
    "bahasa inggris 1": "bahasa inggris teknik",
    "bahasa inggris 2": "bahasa inggris untuk komunikasi",
    "matematika 3": "matematika 2",
    "matematika terapan 1": "matematika 2",
}


def _best_name(name: str, choices: list[str], syllabus: bool = False) -> str:
    if not choices:
        return ""
    target = _norm(name)
    exact = {_norm(choice): choice for choice in choices}
    if target in exact:
        return exact[target]
    aliases = ALIASES if syllabus else BASE_ALIASES
    target = aliases.get(target, target)
    if target in exact:
        return exact[target]
    return max(choices, key=lambda choice: SequenceMatcher(None, target, _norm(choice)).ratio())


def _course_rows(df: pd.DataFrame, source: dict[str, Any]) -> pd.DataFrame:
    if df.empty:
        return df.copy()
    offer = str(source.get("id_penawaran", ""))
    if offer and "id_penawaran" in df.columns:
        hit = df[df["id_penawaran"].astype(str) == offer]
        if not hit.empty:
            return hit.copy()
    return df[df["kode_mk"].astype(str) == str(source.get("kode_mk", ""))].copy()


def _replace_identity(df: pd.DataFrame, target: dict[str, Any]) -> pd.DataFrame:
    result = df.copy()
    if result.empty:
        return result
    old_codes = result["kode_mk"].astype(str).unique().tolist() if "kode_mk" in result else []
    result["kode_mk"] = target["kode_mk"]
    result["id_penawaran"] = target["id_penawaran"]
    for column in ("kode_cpmk", "key_mk_cpmk"):
        if column in result.columns:
            for old_code in old_codes:
                if old_code:
                    result[column] = result[column].astype(str).str.replace(
                        old_code, target["kode_mk"], regex=False
                    )
    for column in ("nama_mk", "mata_kuliah"):
        if column in result.columns:
            result[column] = target["nama_mk"]
    if "semester" in result.columns:
        result["semester"] = target["semester"]
    return result


def _target_courses(courses: list[dict[str, Any]], program: str) -> list[dict[str, Any]]:
    duplicate_counts = Counter(course["kode_mk"] for course in courses)
    sequence: Counter[str] = Counter()
    rows = []
    for course in courses:
        row = dict(course)
        if program.startswith("D3") and _norm(row["nama_mk"]) == "agama":
            row.update({
                "sks_teori": 2, "sks_praktek": 0, "total_sks": 2,
                "jam_teori_minggu": 2, "jam_praktik_minggu": 0,
                "total_jam_minggu": 2,
            })
        sequence[row["kode_mk"]] += 1
        suffix = sequence[row["kode_mk"]]
        row["id_penawaran"] = (
            f"{row['kode_mk']}-{suffix:02d}"
            if duplicate_counts[row["kode_mk"]] > 1
            else row["kode_mk"]
        )
        row["nama_prodi"] = program
        rows.append(row)
    return rows


def _official_d3_cpmk(catalog: dict[str, Any], cohort: str) -> pd.DataFrame:
    rows = []
    for source in catalog["d3_2026_cpmk"]:
        row = {
            "kode_mk": source["kode_mk"].replace("REC26", f"REC{cohort[-2:]}"),
            "id_penawaran": source["kode_mk"].replace("REC26", f"REC{cohort[-2:]}"),
            "kode_cpmk": source["kode_cpmk"],
            "deskripsi_cpmk": source["rumusan_cpmk"],
            "kode_ik": source["kode_ik"],
            "kode_cpl": source["kode_cpl"],
        }
        rows.append(row)
    return pd.DataFrame(rows)


def _weekly_rows(target: dict[str, Any], materials: list[str], cpmk: pd.DataFrame, weeks: int) -> pd.DataFrame:
    topics = [item for item in materials if str(item).strip()] or [target["nama_mk"]]
    cpmk_codes = cpmk.get("kode_cpmk", pd.Series(dtype=str)).astype(str).tolist()
    rows = []
    final_week = weeks
    for week in range(1, weeks + 1):
        topic_index = week - 1 if week < 9 else week - 2
        material = topics[topic_index % len(topics)]
        sub = f"Mampu menjelaskan dan menerapkan {material}."
        technique, weight = "", 0
        if week == 9:
            material, sub, technique, weight = "UTS", "Evaluasi tengah semester", "UTS", 20
        elif week == final_week:
            material = "Integrasi dan evaluasi: " + topics[-1]
            sub, technique, weight = "Evaluasi akhir atau proyek akhir semester", "UAS", 40
        elif week in (4, 12):
            technique, weight = "Tugas", 10
        elif week in (2, 6, 10, 14):
            technique, weight = "Kuis", 5
        rows.append({
            "kode_mk": target["kode_mk"], "id_penawaran": target["id_penawaran"],
            "minggu": week, "sub_cpmk": sub, "materi": material, "modalitas": "Luring",
            "bentuk_pembelajaran": "Praktikum" if target.get("sks_praktek", 0) else "Kuliah",
            "metode": "Praktik Terbimbing" if target.get("sks_praktek", 0) else "Ceramah Interaktif",
            "pengalaman_belajar": f"Mahasiswa mempelajari dan menerapkan {material}.",
            "teknik_asesmen": technique, "indikator_penilaian": f"Ketepatan penguasaan {material}.",
            "bobot": weight, "referensi": "", "kode_cpmk": cpmk_codes[(week - 1) % len(cpmk_codes)] if cpmk_codes else "",
        })
    return pd.DataFrame(rows)


def build_cohort_workbook(
    base: dict[str, pd.DataFrame], program: str, cohort: str
) -> dict[str, pd.DataFrame]:
    catalog = load_catalog()
    key = curriculum_key(program, cohort)
    targets = _target_courses(catalog["curricula"][key], program)
    base_mk = base["Master_MK"].fillna("")
    base_names = base_mk["nama_mk"].astype(str).tolist()
    syllabus_records = catalog["d4_short_syllabus"]
    syllabus_names = [record["nama_mk"] for record in syllabus_records]
    syllabus_by_name = {record["nama_mk"]: record for record in syllabus_records}

    output: dict[str, list[pd.DataFrame]] = {
        name: [] for name in ["Master_CPMK", "Mapping_MK_CPL", "Short_Silabus", "Referensi", "RPS_Pertemuan"]
    }
    master_rows = []
    official_d3 = _official_d3_cpmk(catalog, cohort) if key in {"D3-2025", "D3-2026"} else pd.DataFrame()

    for target in targets:
        master_rows.append(target)
        source_name = _best_name(target["nama_mk"], base_names)
        source_mk = base_mk[base_mk["nama_mk"].astype(str) == source_name].iloc[0].to_dict()

        if not official_d3.empty:
            cpmk = official_d3[official_d3["kode_mk"] == target["kode_mk"]].copy()
        else:
            cpmk = _replace_identity(_course_rows(base["Master_CPMK"], source_mk), target)
        if cpmk.empty:
            cpmk = pd.DataFrame([{
                "kode_mk": target["kode_mk"], "id_penawaran": target["id_penawaran"],
                "kode_cpmk": "CPMK1", "deskripsi_cpmk": f"Mampu menerapkan kompetensi {target['nama_mk']}.",
                "kode_ik": "", "kode_cpl": "",
            }])
        output["Master_CPMK"].append(cpmk)
        mapping = cpmk[["kode_mk", "id_penawaran", "kode_cpl"]].drop_duplicates()
        mapping = mapping[mapping["kode_cpl"].astype(str) != ""]
        output["Mapping_MK_CPL"].append(mapping)

        if program.startswith("D4"):
            syllabus_name = _best_name(target["nama_mk"], syllabus_names, syllabus=True)
            syllabus = syllabus_by_name[syllabus_name]
            materials = syllabus["bahan_kajian"]
            description = syllabus["deskripsi_mk"]
            references = syllabus["referensi"]
            output["RPS_Pertemuan"].append(_weekly_rows(target, materials, cpmk, 16))
        else:
            source_syllabus = _course_rows(base["Short_Silabus"], source_mk)
            description = str(source_syllabus.iloc[0].get("deskripsi_mk", "")) if not source_syllabus.empty else ""
            materials_text = str(source_syllabus.iloc[0].get("bahan_kajian", "")) if not source_syllabus.empty else ""
            materials = [item.strip() for item in re.split(r"[;\n]+", materials_text) if item.strip()]
            source_refs = _course_rows(base["Referensi"], source_mk)
            references = source_refs.get("referensi", pd.Series(dtype=str)).astype(str).tolist()
            weekly = _replace_identity(_course_rows(base["RPS_Pertemuan"], source_mk), target)
            if not official_d3.empty:
                weekly = _weekly_rows(target, materials, cpmk, 17)
            elif not weekly.empty:
                codes = cpmk["kode_cpmk"].astype(str).tolist()
                if codes:
                    weekly["kode_cpmk"] = [
                        codes[index % len(codes)] for index in range(len(weekly))
                    ]
            output["RPS_Pertemuan"].append(
                weekly if not weekly.empty else _weekly_rows(target, materials, cpmk, 17)
            )

        output["Short_Silabus"].append(pd.DataFrame([{
            "kode_mk": target["kode_mk"], "id_penawaran": target["id_penawaran"],
            "deskripsi_mk": description, "bahan_kajian": "; ".join(materials),
        }]))
        output["Referensi"].append(pd.DataFrame([{
            "kode_mk": target["kode_mk"], "id_penawaran": target["id_penawaran"],
            "referensi": "\n".join(references),
        }]))

    result = {
        "Master_MK": pd.DataFrame(master_rows),
        "Master_CPL": base["Master_CPL"].copy(),
        "Master_IK": base["Master_IK"].copy(),
    }
    for name, frames in output.items():
        result[name] = pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()
    return result


def workbook_to_excel_bytes(workbook: dict[str, pd.DataFrame]) -> bytes:
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        for sheet_name, frame in workbook.items():
            frame.to_excel(writer, sheet_name=sheet_name, index=False)
    return buffer.getvalue()
