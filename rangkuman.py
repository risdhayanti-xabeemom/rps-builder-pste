"""Rangkuman CPL, IK, dan CPMK dari master kurikulum yang sedang dipilih."""

from __future__ import annotations

from io import BytesIO

import pandas as pd


def _cpl_sort_key(code: str) -> tuple[int, str]:
    digits = "".join(ch for ch in str(code) if ch.isdigit())
    return (int(digits) if digits else 0, str(code))


def _joined(values) -> str:
    return ", ".join(dict.fromkeys(str(v) for v in values if str(v).strip()))


def _mk_labels(workbook: dict[str, pd.DataFrame]) -> pd.DataFrame:
    mk = workbook["Master_MK"].fillna("").copy()
    mk["kode_mk"] = mk["kode_mk"].astype(str)
    if "id_penawaran" not in mk.columns:
        mk["id_penawaran"] = mk["kode_mk"]
    return mk[["semester", "kode_mk", "id_penawaran", "nama_mk"]]


def summarize_cpl(workbook: dict[str, pd.DataFrame]) -> pd.DataFrame:
    """Satu baris per CPL: jumlah IK, jumlah MK pengampu, dan jumlah CPMK."""
    cpl = workbook["Master_CPL"].fillna("")
    ik = workbook["Master_IK"].fillna("")
    cpmk = summarize_cpmk(workbook)
    rows = []
    for item in cpl.to_dict("records"):
        code = item["kode_cpl"]
        cpmk_cpl = cpmk[cpmk["kode_cpl"].astype(str) == str(code)]
        rows.append({
            "kode_cpl": code,
            "deskripsi_cpl": item.get("deskripsi_cpl", ""),
            "jumlah_ik": int((ik["kode_cpl"].astype(str) == str(code)).sum()),
            "jumlah_mk": int(cpmk_cpl["kode_mk"].nunique()),
            "jumlah_cpmk": int(len(cpmk_cpl)),
        })
    return pd.DataFrame(rows).sort_values("kode_cpl", key=lambda s: s.map(_cpl_sort_key), ignore_index=True)


def summarize_ik(workbook: dict[str, pd.DataFrame]) -> pd.DataFrame:
    """Satu baris per IK dengan MK pengampu dan jumlah CPMK-nya."""
    ik = workbook["Master_IK"].fillna("")
    cpmk = workbook["Master_CPMK"].fillna("").copy()
    cpmk["ik_list"] = cpmk["kode_ik"].astype(str).str.findall(r"IK\s*0*\d+\.0*\d+")
    rows = []
    for item in ik.to_dict("records"):
        code = str(item["kode_ik"])
        number = "".join(ch for ch in code if ch.isdigit() or ch == ".")
        major, _, minor = number.partition(".")
        pattern = f"{int(major or 0)}.{int(minor or 0)}"
        hit = cpmk[cpmk["ik_list"].map(
            lambda found: any(
                f"{int(a)}.{int(b)}" == pattern
                for a, b in (x.replace("IK", "").strip().split(".") for x in found)
            )
        )]
        rows.append({
            "kode_ik": code, "kode_cpl": item.get("kode_cpl", ""),
            "deskripsi_ik": item.get("deskripsi_ik", ""),
            "jumlah_mk": int(hit["id_penawaran"].nunique()), "jumlah_cpmk": int(len(hit)),
            "mata_kuliah": _joined(hit["kode_mk"]),
        })
    return pd.DataFrame(rows)


def summarize_cpmk(workbook: dict[str, pd.DataFrame]) -> pd.DataFrame:
    """Satu baris per CPMK dengan nama MK, semester, IK, CPL, dan bobotnya dalam MK."""
    mk = _mk_labels(workbook)
    cpmk = workbook["Master_CPMK"].fillna("").copy()
    cpmk["id_penawaran"] = cpmk["id_penawaran"].astype(str)
    cpmk = cpmk.drop(columns=[c for c in ("semester", "nama_mk") if c in cpmk.columns])
    merged = cpmk.merge(mk[["id_penawaran", "semester", "nama_mk"]], on="id_penawaran", how="left")
    merged["bobot_persen"] = pd.to_numeric(merged.get("bobot_cpmk_mk_persen"), errors="coerce")
    if merged["bobot_persen"].isna().all():  # master lama tanpa kolom bobot: jumlahkan bobot mingguan
        weekly = workbook.get("RPS_Pertemuan", pd.DataFrame())
        if not weekly.empty:
            weekly = weekly.assign(bobot=pd.to_numeric(weekly["bobot"], errors="coerce").fillna(0))
            sums = weekly.groupby(["kode_mk", "kode_cpmk"])["bobot"].sum()
            merged["bobot_persen"] = [
                sums.get((row.kode_mk, row.kode_cpmk)) for row in merged.itertuples()
            ]
    blank = merged["kode_cpl"].astype(str).str.strip() == ""
    if blank.any():  # master lama: CPL diambil dari IK
        by_ik = dict(zip(workbook["Master_IK"]["kode_ik"].astype(str), workbook["Master_IK"]["kode_cpl"].astype(str)))
        merged.loc[blank, "kode_cpl"] = merged.loc[blank, "kode_ik"].astype(str).map(by_ik).fillna("")
    merged["semester"] = pd.to_numeric(merged["semester"], errors="coerce")
    merged = merged.sort_values(["semester", "kode_mk", "kode_cpmk"], ignore_index=True)
    return merged[[
        "semester", "kode_mk", "nama_mk", "kode_cpmk", "deskripsi_cpmk", "kode_ik", "kode_cpl", "bobot_persen",
    ]]


def cpl_by_course(workbook: dict[str, pd.DataFrame]) -> pd.DataFrame:
    """Matriks MK x CPL: jumlah CPMK MK itu yang membebani CPL tersebut."""
    cpmk = summarize_cpmk(workbook)
    cpmk = cpmk[cpmk["kode_cpl"].astype(str) != ""]
    table = pd.crosstab([cpmk["semester"], cpmk["kode_mk"], cpmk["nama_mk"]], cpmk["kode_cpl"])
    table = table[sorted(table.columns, key=_cpl_sort_key)].reset_index()
    return table


def course_totals(workbook: dict[str, pd.DataFrame]) -> pd.DataFrame:
    """Satu baris per MK: jumlah CPMK, total bobot (harus 100), dan CPL yang dibebankan."""
    cpmk = summarize_cpmk(workbook)
    rows = []
    for (semester, code, name), group in cpmk.groupby(["semester", "kode_mk", "nama_mk"], dropna=False, sort=True):
        weights = group["bobot_persen"]
        rows.append({
            "semester": semester, "kode_mk": code, "nama_mk": name, "jumlah_cpmk": int(len(group)),
            "total_bobot_persen": round(float(weights.sum()), 2) if weights.notna().any() else None,
            "cpl": _joined(sorted(group["kode_cpl"].astype(str), key=_cpl_sort_key)),
        })
    return pd.DataFrame(rows)


def build_summary(workbook: dict[str, pd.DataFrame]) -> dict[str, pd.DataFrame]:
    return {
        "Ringkasan CPL": summarize_cpl(workbook),
        "IK per CPL": summarize_ik(workbook),
        "CPL per MK": cpl_by_course(workbook),
        "Ringkasan MK": course_totals(workbook),
        "CPMK dan Bobot": summarize_cpmk(workbook),
    }


def summary_to_excel_bytes(sheets: dict[str, pd.DataFrame]) -> bytes:
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        for name, frame in sheets.items():
            frame.to_excel(writer, sheet_name=name[:31], index=False)
            sheet = writer.sheets[name[:31]]
            sheet.freeze_panes = "A2"
            for column in sheet.columns:
                longest = max((len(str(cell.value)) for cell in column if cell.value is not None), default=8)
                sheet.column_dimensions[column[0].column_letter].width = min(max(longest + 2, 8), 60)
    return buffer.getvalue()
