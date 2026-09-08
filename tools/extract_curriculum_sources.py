from __future__ import annotations

import json
import math
import re
import sys
from pathlib import Path

import pandas as pd
import pdfplumber


ROOT = Path(__file__).resolve().parents[1]
UPLOAD = ROOT.parent / "upload"
LIBRARY_SOURCE = ROOT.parent / "source_library" / "Dokumen Kurikulum D3"


def clean(value: object) -> str:
    if value is None or (isinstance(value, float) and math.isnan(value)):
        return ""
    return re.sub(r"\s+", " ", str(value)).strip()


def number(value: object) -> int | float:
    try:
        result = float(value)
    except (TypeError, ValueError):
        return 0
    return int(result) if result.is_integer() else result


def split_course_name(value: object) -> tuple[str, str]:
    text = clean(value)
    if "/" not in text:
        return text, ""
    left, right = text.split("/", 1)
    return clean(left), clean(right)


def parse_mapping(path: Path, sheet_names: dict[str, tuple[str, int]]) -> dict[str, list[dict[str, object]]]:
    result: dict[str, list[dict[str, object]]] = {}
    for key, (sheet_name, expected_year) in sheet_names.items():
        df = pd.read_excel(path, sheet_name=sheet_name, header=None, dtype=object)
        courses: list[dict[str, object]] = []
        for _, row in df.iterrows():
            values = row.tolist()
            match_index = None
            code = ""
            for idx, value in enumerate(values):
                candidate = clean(value).upper()
                if re.fullmatch(r"(?:REC|RTE)\d{6}", candidate):
                    match_index = idx
                    code = candidate
                    break
            if match_index is None:
                continue
            name_id, name_en = split_course_name(
                values[match_index + 1] if match_index + 1 < len(values) else ""
            )
            theory = number(values[match_index + 2] if match_index + 2 < len(values) else 0)
            practice = number(values[match_index + 3] if match_index + 3 < len(values) else 0)
            total = number(values[match_index + 4] if match_index + 4 < len(values) else 0)
            theory_hours = number(values[match_index + 5] if match_index + 5 < len(values) else 0)
            practice_hours = number(values[match_index + 6] if match_index + 6 < len(values) else 0)
            total_hours = number(values[match_index + 7] if match_index + 7 < len(values) else 0)
            status = clean(values[match_index + 8] if match_index + 8 < len(values) else "")
            note = clean(values[match_index + 9] if match_index + 9 < len(values) else "")
            semester = int(code[5])
            if int(code[3:5]) not in {expected_year, 21, 22}:
                raise ValueError(f"Unexpected cohort code {code} in {sheet_name}")
            if not theory_hours and theory:
                theory_hours = theory
            if not practice_hours and practice:
                practice_hours = practice * 2
            if not total_hours:
                total_hours = theory_hours + practice_hours
            if not total:
                total = theory + practice
            courses.append(
                {
                    "semester": semester,
                    "kode_mk": code,
                    "nama_mk": name_id,
                    "course_name": name_en,
                    "sks_teori": theory,
                    "sks_praktek": practice,
                    "total_sks": total,
                    "jam_teori_minggu": theory_hours,
                    "jam_praktik_minggu": practice_hours,
                    "total_jam_minggu": total_hours,
                    "jenis_mk": status or "Wajib",
                    "catatan": note,
                }
            )
        result[key] = courses
    return result


def extract_pdf_text(path: Path) -> str:
    with pdfplumber.open(path) as pdf:
        return "\n".join((page.extract_text(x_tolerance=2, y_tolerance=3) or "") for page in pdf.pages)


def parse_short_syllabus(path: Path) -> list[dict[str, object]]:
    text = extract_pdf_text(path)
    text = text.replace("\r", "")
    course_pattern = re.compile(
        r"(?m)^Mata Kuliah\s*:?[ \t]+(?P<name>[^\n]+)"
    )
    semester_pattern = re.compile(r"(?m)^SEMESTER\s+(?P<semester>\d+)\s*:?[ \t]*$")
    course_matches = list(course_pattern.finditer(text))
    semester_matches = list(semester_pattern.finditer(text))
    records: list[dict[str, object]] = []
    for index, match in enumerate(course_matches):
        start = match.start()
        end = course_matches[index + 1].start() if index + 1 < len(course_matches) else len(text)
        block = text[start:end]
        semester = 0
        for semester_match in semester_matches:
            if semester_match.start() <= start:
                semester = int(semester_match.group("semester"))
            else:
                break
        name = clean(match.group("name")).lstrip(":").strip()
        description_match = re.search(
            r"(?is)Deskripsi\s*:?[ \t]*(.*?)(?=\n\s*Capaian Pembelajaran\s*:?)",
            block,
        )
        outcome_match = re.search(
            r"(?is)Capaian Pembelajaran\s*:?[ \t]*(.*?)(?=\n\s*Materi(?:\s+Kuliah)?\s*:?)",
            block,
        )
        material_match = re.search(
            r"(?is)\n\s*Materi(?:\s+Kuliah)?\s*:?[ \t]*(.*?)(?=\n\s*Referensi\s*:?)",
            block,
        )
        reference_match = re.search(
            r"(?is)\n\s*Referensi\s*:?[ \t]*(.*)$",
            block,
        )
        description = clean(description_match.group(1)) if description_match else ""
        outcomes_raw = outcome_match.group(1).strip() if outcome_match else ""
        materials_raw = material_match.group(1).strip() if material_match else ""
        references_raw = reference_match.group(1).strip() if reference_match else ""

        def lines(raw: str) -> list[str]:
            normalized = raw.replace("●", "\n").replace("•", "\n")
            parts = []
            for line in normalized.splitlines():
                item = clean(line)
                item = re.sub(r"^(?:\d+[.)]|[a-zA-Z][.)])\s*", "", item)
                if item:
                    parts.append(item)
            return parts

        records.append(
            {
                "semester": semester,
                "nama_mk": name,
                "deskripsi_mk": description,
                "capaian_pembelajaran": lines(outcomes_raw),
                "bahan_kajian": lines(materials_raw),
                "referensi": lines(references_raw),
            }
        )
    return records


def parse_d3_cpmk(path: Path) -> list[dict[str, str]]:
    df = pd.read_excel(path, sheet_name="Master CPMK-MK", header=3, dtype=object)
    df.columns = [
        "no",
        "semester",
        "kode_mk",
        "nama_mk",
        "course_name",
        "sks_teori",
        "sks_praktek",
        "total_sks",
        "jam_teori_minggu",
        "jam_praktik_minggu",
        "total_jam_minggu",
        "jenis_mk",
        "kode_cpl",
        "kode_ik",
        "kode_cpmk",
        "rumusan_cpmk",
        "catatan",
    ]
    rows = []
    for row in df.to_dict("records"):
        code = clean(row["kode_mk"]).upper()
        if not re.fullmatch(r"REC26\d{4}", code):
            continue
        rows.append({key: clean(value) for key, value in row.items() if key not in {"no"}})
    return rows


def main() -> None:
    d4_mapping = parse_mapping(
        UPLOAD / "2026-akademik - mapping.xlsx",
        {
            "D4-2023": ("D4-2023 (edit)-fix", 23),
            "D4-2024": ("D4-2024 (mila-pipit-lili)-fix", 24),
            "D4-2025": ("D4-2025-fix", 25),
            "D4-2026": ("D4-2026", 26),
        },
    )
    d3_mapping = parse_mapping(
        LIBRARY_SOURCE / "Mapping_Kurikulum_TA_2026_2027_REC24_REC25_REC26.xlsx",
        {
            "D3-2024": ("D3-2024 fix  (name-edit) obe", 24),
            "D3-2025": ("D3-2025", 25),
            "D3-2026": ("D3-2026", 26),
        },
    )
    payload = {
        "source_notes": {
            "D3": "Mapping_Kurikulum_TA_2026_2027_REC24_REC25_REC26.xlsx",
            "D4": "2026-akademik - mapping.xlsx",
            "D4_silabus": "Short Silabus Kurikulum D4 PSTE.pdf",
            "D4_CPL": "MK-CPL D4 TE (1).pdf",
            "D4_CPMK": "!CPL-CPMK-Matkul 2022.xlsx",
        },
        "curricula": {**d3_mapping, **d4_mapping},
        "d4_short_syllabus": parse_short_syllabus(
            UPLOAD / "Short Silabus Kurikulum D4 PSTE.pdf"
        ),
        "d3_2026_cpmk": parse_d3_cpmk(
            LIBRARY_SOURCE / "Master_CPMK_MK_Angkatan_2026_REC26.xlsx"
        ),
    }
    json.dump(payload, sys.stdout, ensure_ascii=False, indent=2)


if __name__ == "__main__":
    main()
