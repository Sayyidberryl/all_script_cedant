from __future__ import annotations
import argparse
from pathlib import Path
import re
import pandas as pd

HERE = Path(__file__).resolve().parent
DEFAULT_INPUT = HERE / "output matching" / "matching_jasarp.xlsx"
DEFAULT_OUTPUT = HERE / "output matching" / "audit_fac_code_flag_jasarp.xlsx"


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
        raise FileNotFoundError(f"File hasil matching JASARP tidak ditemukan: {source}")

    # Baca sebagai teks agar FAC CODE tidak berubah menjadi angka/notasi ilmiah.
    df = pd.read_excel(source, dtype=str, engine="openpyxl")
    mandatory = {"CCOS_REF_CODE", "FLAG_PROD"}
    missing = mandatory - set(df.columns)
    if missing:
        raise ValueError("Kolom wajib tidak ada: " + ", ".join(sorted(missing)))

    df = df.copy()
    df.insert(0, "BARIS_EXCEL", range(2, len(df) + 2))
    df["FAC_CODE_AUDIT"] = df["CCOS_REF_CODE"].map(text_value)
    df["FLAG_AUDIT"] = df["FLAG_PROD"].map(text_value)
    df["KATEGORI_FLAG_AUDIT"] = df["FLAG_AUDIT"].map(flag_category)
    selected = df.loc[df["KATEGORI_FLAG_AUDIT"].notna()].copy()

    # FAC CODE kosong/gabungan dikecualikan dari grouping agar tidak memicu
    # konflik palsu; tetap dilaporkan untuk pemeriksaan manual.
    multi_mask = selected["FAC_CODE_AUDIT"].str.contains(r"[,;\n\r|]", regex=True, na=False)
    missing_mask = selected["FAC_CODE_AUDIT"].eq("")
    exceptions = selected.loc[multi_mask | missing_mask].copy()
    exceptions["CATATAN_AUDIT"] = exceptions["FAC_CODE_AUDIT"].map(
        lambda x: "FAC CODE kosong" if not x
        else "Berisi beberapa FAC CODE; perlu pemeriksaan manual"
    )
    valid = selected.loc[~(multi_mask | missing_mask)].copy()

    records = []
    for code, part in valid.groupby("FAC_CODE_AUDIT", sort=True, dropna=False):
        flags = sorted(part["FLAG_AUDIT"].dropna().unique().tolist())
        categories = sorted(part["KATEGORI_FLAG_AUDIT"].dropna().unique().tolist())
        cross_conflict = "New Entry" in categories and "Adjustment" in categories
        if cross_conflict:
            status = "KONFLIK: New Entry + Adjustment"
        elif len(flags) > 1:
            status = "PERIKSA: >1 subjenis FLAG_PROD"
        else:
            status = "OK: satu jenis FLAG_PROD"
        records.append({
            "CCOS_REF_CODE": code,
            "JUMLAH_BARIS": len(part),
            "JUMLAH_JENIS_FLAG_PROD": len(flags),
            "DAFTAR_FLAG_PROD": " | ".join(flags),
            "KATEGORI_TERLIBAT": " + ".join(categories),
            "JUMLAH_NEW_ENTRY": int((part["KATEGORI_FLAG_AUDIT"] == "New Entry").sum()),
            "JUMLAH_ADJUSTMENT": int((part["KATEGORI_FLAG_AUDIT"] == "Adjustment").sum()),
            "DOUBLE_FLAG_NEW_ENTRY_ADJUSTMENT": "YA" if cross_conflict else "TIDAK",
            "STATUS_AUDIT": status,
            "BARIS_EXCEL": ", ".join(map(str, part["BARIS_EXCEL"])),
        })

    cols = [
        "CCOS_REF_CODE", "JUMLAH_BARIS", "JUMLAH_JENIS_FLAG_PROD",
        "DAFTAR_FLAG_PROD", "KATEGORI_TERLIBAT", "JUMLAH_NEW_ENTRY",
        "JUMLAH_ADJUSTMENT", "DOUBLE_FLAG_NEW_ENTRY_ADJUSTMENT",
        "STATUS_AUDIT", "BARIS_EXCEL",
    ]
    summary = pd.DataFrame(records, columns=cols)
    conflicts = summary.loc[summary["DOUBLE_FLAG_NEW_ENTRY_ADJUSTMENT"] == "YA"].copy()
    multi_flags = summary.loc[summary["JUMLAH_JENIS_FLAG_PROD"] > 1].copy()
    conflict_codes = set(conflicts["CCOS_REF_CODE"])
    detail = valid.loc[valid["FAC_CODE_AUDIT"].isin(conflict_codes)].copy()

    stats = pd.DataFrame([
        {"METRIK": "Total baris New Entry + Adjustment", "NILAI": len(selected)},
        {"METRIK": "FAC CODE unik dengan kode tunggal", "NILAI": len(summary)},
        {"METRIK": "FAC CODE New Entry DAN Adjustment", "NILAI": len(conflicts)},
        {"METRIK": "FAC CODE dengan satu kategori utama", "NILAI": len(summary) - len(conflicts)},
        {"METRIK": "FAC CODE dengan >1 subjenis FLAG_PROD", "NILAI": len(multi_flags)},
        {"METRIK": "Baris FAC CODE kosong/gabungan perlu diperiksa", "NILAI": len(exceptions)},
    ])

    target.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(target, engine="openpyxl") as writer:
        stats.to_excel(writer, sheet_name="Statistik", index=False)
        summary.to_excel(writer, sheet_name="Semua FAC CODE", index=False)
        conflicts.to_excel(writer, sheet_name="Double Flag NEntry Adj", index=False)
        multi_flags.to_excel(writer, sheet_name="Multi Subjenis Flag", index=False)
        detail.to_excel(writer, sheet_name="Detail Double Flag", index=False)
        exceptions.to_excel(writer, sheet_name="Periksa FAC CODE", index=False)

        from openpyxl.styles import Alignment, Font, PatternFill
        for ws in writer.book.worksheets:
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions
            for cell in ws[1]:
                cell.font = Font(bold=True, color="FFFFFF")
                cell.fill = PatternFill("solid", fgColor="17365D")
                cell.alignment = Alignment(vertical="center", wrap_text=True)
            for cells in ws.columns:
                header = str(cells[0].value or "")
                width = min(max(len(header) + 3, 16), 48)
                if header in {"DAFTAR_FLAG_PROD", "BARIS_EXCEL", "CATATAN_AUDIT", "KATEGORI_TERLIBAT"}:
                    width = 56
                ws.column_dimensions[cells[0].column_letter].width = width

    return {
        "total_flagged_rows": len(selected),
        "unique_fac_codes": len(summary),
        "double_flag_fac_codes": len(conflicts),
        "multiple_subtype_fac_codes": len(multi_flags),
        "exceptions": len(exceptions),
        "output": str(target),
    }


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Audit FAC CODE JASARP untuk double flag New Entry/Adjustment."
    )
    parser.add_argument("--input", type=Path, default=DEFAULT_INPUT,
                        help="Excel hasil matching JASARP")
    parser.add_argument("--output", type=Path, default=DEFAULT_OUTPUT,
                        help="Excel hasil audit")
    args = parser.parse_args()
    result = audit(args.input, args.output)
    print("=" * 62)
    print("AUDIT FAC CODE JASARP: NEW ENTRY / ADJUSTMENT")
    print(f"Total baris New Entry/Adjustment : {result['total_flagged_rows']:,}")
    print(f"Jumlah FAC CODE unik             : {result['unique_fac_codes']:,}")
    print(f"FAC CODE New Entry + Adjustment : {result['double_flag_fac_codes']:,}")
    print(f"FAC CODE >1 subjenis FLAG_PROD   : {result['multiple_subtype_fac_codes']:,}")
    print(f"Baris perlu pemeriksaan          : {result['exceptions']:,}")
    print(f"Hasil audit                      : {result['output']}")
    print("=" * 62)


if __name__ == "__main__":
    main()
