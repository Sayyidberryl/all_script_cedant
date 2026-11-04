from __future__ import annotations

import argparse
from pathlib import Path
import re
import pandas as pd

HERE = Path(__file__).resolve().parent
DEFAULT_INPUT = HERE / "output" / "final_output_v1.xlsx"
DEFAULT_OUTPUT = HERE / "output" / "audit_fac_code_flag_kbru.xlsx"


def text_value(value) -> str:
    if pd.isna(value):
        return ""
    val = str(value).strip()
    if val.upper() in {"NAN", "NONE", "NULL", "NAT", "-"}:
        return ""
    return re.sub(r"(?<=\d)\.0$", "", val)


def flag_category(flag: str) -> str | None:
    normalized = re.sub(r"\s+", " ", flag.strip()).casefold()
    if normalized.startswith("new entry"):
        return "New Entry"
    if normalized.startswith("adjustment"):
        return "Adjustment"
    return None


def audit(source: Path, target: Path) -> dict:
    if not source.is_file():
        raise FileNotFoundError(f"File hasil matching KBRU tidak ditemukan: {source}")

    df = pd.read_excel(source, dtype={"CCOS_REF_CODE": str}, engine="openpyxl")
    mandatory = {"CCOS_REF_CODE", "FLAG_PROD"}
    missing = mandatory - set(df.columns)
    if missing:
        raise ValueError(f"Kolom wajib tidak ada: {', '.join(sorted(missing))}")

    df = df.copy()
    df.insert(0, "BARIS_EXCEL", range(2, len(df) + 2))
    df["FAC_CODE_AUDIT"] = df["CCOS_REF_CODE"].map(text_value)
    df["FLAG_AUDIT"] = df["FLAG_PROD"].map(text_value)
    df["KATEGORI_FLAG_AUDIT"] = df["FLAG_AUDIT"].map(flag_category)
    selected = df.loc[df["KATEGORI_FLAG_AUDIT"].notna()].copy()

    # Hanya FAC CODE tunggal yang bisa dikelompokkan dengan aman.
    # FAC CODE kosong / gabungan masuk lembar pemeriksaan tersendiri.
    multi_mask = selected["FAC_CODE_AUDIT"].str.contains(",|;|\n", regex=True, na=False)
    missing_mask = selected["FAC_CODE_AUDIT"].eq("")
    exceptions = selected.loc[multi_mask | missing_mask].copy()
    exceptions["CATATAN_AUDIT"] = exceptions["FAC_CODE_AUDIT"].map(
        lambda x: "FAC CODE kosong" if not x else "Kolom berisi beberapa FAC CODE; perlu pemeriksaan manual"
    )
    valid = selected.loc[~(multi_mask | missing_mask)].copy()

    summary_records = []
    for code, part in valid.groupby("FAC_CODE_AUDIT", sort=True, dropna=False):
        unique_flags = sorted(part["FLAG_AUDIT"].dropna().unique().tolist())
        categories = sorted(part["KATEGORI_FLAG_AUDIT"].dropna().unique().tolist())
        summary_records.append({
            "CCOS_REF_CODE": code,
            "JUMLAH_BARIS": len(part),
            "JUMLAH_JENIS_FLAG_PROD": len(unique_flags),
            "DAFTAR_FLAG_PROD": " | ".join(unique_flags),
            "KATEGORI_TERLIBAT": " + ".join(categories),
            "JUMLAH_NEW_ENTRY": int((part["KATEGORI_FLAG_AUDIT"] == "New Entry").sum()),
            "JUMLAH_ADJUSTMENT": int((part["KATEGORI_FLAG_AUDIT"] == "Adjustment").sum()),
            "STATUS_AUDIT": "KONFLIK: >1 jenis FLAG_PROD" if len(unique_flags) > 1 else "OK: 1 jenis FLAG_PROD",
            "BARIS_EXCEL": ", ".join(map(str, part["BARIS_EXCEL"])),
        })

    cols = ["CCOS_REF_CODE", "JUMLAH_BARIS", "JUMLAH_JENIS_FLAG_PROD", "DAFTAR_FLAG_PROD", "KATEGORI_TERLIBAT", "JUMLAH_NEW_ENTRY", "JUMLAH_ADJUSTMENT", "STATUS_AUDIT", "BARIS_EXCEL"]
    summary = pd.DataFrame(summary_records, columns=cols)
    conflicts = summary.loc[summary["JUMLAH_JENIS_FLAG_PROD"] > 1].copy()
    conflicting_codes = set(conflicts["CCOS_REF_CODE"])
    detail = valid.loc[valid["FAC_CODE_AUDIT"].isin(conflicting_codes)].copy()

    headline = pd.DataFrame([
        {"METRIK": "Baris New Entry + Adjustment", "NILAI": len(selected)},
        {"METRIK": "FAC CODE unik dengan kode tunggal", "NILAI": len(summary)},
        {"METRIK": "FAC CODE dengan satu jenis FLAG_PROD", "NILAI": len(summary) - len(conflicts)},
        {"METRIK": "FAC CODE dengan >1 jenis FLAG_PROD", "NILAI": len(conflicts)},
        {"METRIK": "FAC CODE dengan New Entry DAN Adjustment", "NILAI": int(summary["KATEGORI_TERLIBAT"].eq("Adjustment + New Entry").sum())},
        {"METRIK": "Baris perlu pemeriksaan FAC CODE", "NILAI": len(exceptions)},
    ])

    target.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(target, engine="openpyxl") as writer:
        headline.to_excel(writer, sheet_name="Statistik", index=False)
        summary.to_excel(writer, sheet_name="Semua FAC CODE", index=False)
        conflicts.to_excel(writer, sheet_name="FAC CODE Multi Flag", index=False)
        detail.to_excel(writer, sheet_name="Detail Multi Flag", index=False)
        exceptions.to_excel(writer, sheet_name="Periksa FAC CODE", index=False)
        from openpyxl.styles import Font, PatternFill, Alignment
        for ws in writer.book.worksheets:
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions
            for cell in ws[1]:
                cell.font = Font(bold=True, color="FFFFFF")
                cell.fill = PatternFill("solid", fgColor="17365D")
                cell.alignment = Alignment(vertical="center", wrap_text=True)
            for col_cells in ws.columns:
                header = str(col_cells[0].value or "")
                width = min(max(len(header) + 3, 16), 48)
                if header in {"DAFTAR_FLAG_PROD", "BARIS_EXCEL", "CATATAN_AUDIT"}:
                    width = 56
                ws.column_dimensions[col_cells[0].column_letter].width = width

    return {
        "total_flagged_rows": len(selected),
        "unique_fac_codes": len(summary),
        "multiple_flag_fac_codes": len(conflicts),
        "exceptions": len(exceptions),
        "output": str(target),
    }


def main():
    parser = argparse.ArgumentParser(description="Audit FAC CODE KBRU dengan lebih dari satu FLAG_PROD New Entry/Adjustment")
    parser.add_argument("--input", type=Path, default=DEFAULT_INPUT, help="Excel hasil matching KBRU")
    parser.add_argument("--output", type=Path, default=DEFAULT_OUTPUT, help="Excel hasil audit KBRU")
    args = parser.parse_args()
    results = audit(args.input, args.output)
    print("=" * 60)
    print("AUDIT FAC CODE KBRU: NEW ENTRY / ADJUSTMENT")
    print(f"Total baris New Entry/Adjustment : {results['total_flagged_rows']:,}")
    print(f"Jumlah FAC CODE unik             : {results['unique_fac_codes']:,}")
    print(f"FAC CODE dengan >1 jenis FLAG    : {results['multiple_flag_fac_codes']:,}")
    print(f"Baris perlu pemeriksaan         : {results['exceptions']:,}")
    print(f"Hasil audit                     : {results['output']}")


if __name__ == "__main__":
    main()
