import os
import re
import numpy as np
import pandas as pd

print("SCRIPT BERJALAN")

# ============================================================
# KONFIGURASI MULTI ARTHA GUNA
# ============================================================

BASE_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))

def _find_input_file(candidates: list) -> str:
    for c in candidates:
        if os.path.exists(c):
            return c
    return candidates[0]

INPUT_FILE = _find_input_file([
    os.path.join(BASE_DIR, "data_raw", "3b. Database Suspense 150826.xlsx"),
    os.path.join(BASE_DIR, "input", "3b. Database Suspense 150826.xlsx"),
    os.path.join("data_raw", "3b. Database Suspense 150826.xlsx"),
    os.path.join("input", "3b. Database Suspense 150826.xlsx"),
])
OUTPUT_FILE = os.path.join(BASE_DIR, "data_cleaned", "mag_output_suspend.xlsx")
SHEET_NAME = 0

CEDANT_COL = "CEDANT NAME"
CEDANT_VALUE = "PT.ASURANSI MULTI ARTHA GUNA"

MAX_BREAKDOWN = 5


# ============================================================
# INSURED RULE
# ============================================================

_INSURED_TAIL_RE = re.compile(
    r"""
      \bAS\b\s+(?:THE\s+)?(?:PRINCIPAL|OFF-TAKER|MAINTENANCE|CONTRACTOR).*
    | \bBEING\s+(?:THE\s+)?(?:PRINCIPAL|OFF-TAKER).*
    | \bAND\s+ALL\s+SUBSIDIARI.*
    | \bINCLUDING\s+ALL\s+SUBSIDIARI.*
    | \bINCLUDING\s+ANY\s+SUBSIDIAR.*
    | \bINCLUDING\s+ANY\s+SUB\.?.*
    | \bCOMPRISING\s+OF.*
    | \bINSTALLMENT\b.*
    | \bRELATED\s+COMPANY\b.*
    | \bPURCHASED\s+OR\s+OTHERWISE\b.*
    | \bWPC\s+DATE\s*:.* 
    """,
    re.IGNORECASE | re.VERBOSE,
)

_INSURED_JUNK_WORDS = frozenset({
    "SHANGHAI", "PR OF CHINA", "CHINA", "INDONESIA", "JAKARTA",
    "OFFICERS", "EMPLOYEES", "ALL OTHER CONTRACTORS",
    "SUB-CONTRACTORS", "SUB CONTRACTORS",
    "COMPANIES", "AFFILIATED", "AFFILIATES",
    "CORPORATIONS AND INCLUDING PARTNERSHIP",
    "JOINT VENTURES AND AGREEMENT OR BY LAW",
    "AS THEIR RESPECTIVE INTEREST MAY APPEAR",
    "AS THEIR RESPECTIVE INTERESTS MAY APPEAR",
    "SUBSIDIARY", "SUBSIDIARIES", "ANY SUBSIDIARY COMPANY",
    "RELATED COMPANY",
    "FOR THEIRS RESPECTIVE RIGHTS AND INTEREST",
    "MIGRASI AS400", "THE PRINCIPAL", "PRINCIPAL", "OWNER",
})

_INSURED_TITLE_RE = re.compile(
    r"\b(?:BAPAK|IBU|BPK|NY\.?|SDR\.?|SDRI\.?|TN\.?|MR\.?|MRS\.?|MS\.?)\b\s*",
    re.IGNORECASE,
)


# ============================================================
# HELPER
# ============================================================

def _cap_breakdown(parts: list, sep: str = "; ") -> list:
    parts = [p for p in parts if p]
    if len(parts) > MAX_BREAKDOWN:
        return [sep.join(parts)]
    return parts


# ============================================================
# CLEAN POLIS
# ============================================================

def clean_polis(val) -> list:
    """
    Etiqa:
    - 3010010924004217 -> 3010010924004217
    - 3010010924004217-000239 -> 3010010924004217
    - 3010010924004217-1/0 -> 3010010924004217
    - separator . / - dibersihkan
    """

    if pd.isna(val):
        return []

    val = str(val).strip()

    if not val or val == "-":
        return []

    val = val.replace(",", "")
    raw_parts = [p.strip() for p in val.split("+") if p.strip()]
    cleaned_parts = []

    for p in raw_parts:
        cur = p.strip()

        # Ambil polis utama saja.
        # Contoh:
        # 40010924046082-000667-1/0 -> 40010924046082
        # 40010924046082-000689-1/1 -> 40010924046082
        # 45013022023211-000047 -> 45013022023211
        # Hapus suffix certificate + fraction
        # Contoh: 40010924046082-000667-1/0 -> 40010924046082
        cur = re.sub(r"-\d{1,6}(?:-\d+/\d+)?$", "", cur)

        # Kalau hanya ada fraction:
        # Contoh: 40013024000741-11/10 -> 40013024000741
        cur = re.sub(r"-\d+/\d+$", "", cur)

        # Hapus separator polis
        cur = cur.replace(".", "")
        cur = cur.replace("/", "")
        cur = cur.replace("-", "")

        # Jangan masukkan token kosong
        if cur:
            cleaned_parts.append(cur)

    return _cap_breakdown(cleaned_parts)


# ============================================================
# CLEAN CERTIFICATE
# ============================================================

# def clean_certificate(polis_ori):
#     """
#     Certificate Etiqa diambil dari suffix numeric setelah dash.

#     Contoh:
#       3010010924004217-000239 -> 000239
#       3010010524002579-000920 -> 000920
#       3010010924004217-1/1   -> blank
#       3010010924004217       -> blank

#     Certificate hanya valid jika suffix numeric 1-6 digit.
#     """

#     if pd.isna(polis_ori):
#         return ""

#     polis = str(polis_ori).strip().upper()

#     if not polis or polis == "-":
#         return ""

#     # Data yang memang bukan certificate
#     if re.search(r"\bSUSPENSE\b|\bENDORSEMENT\b", polis):
#         return ""

#     # Jika berakhiran fraksi -1/1, -2/2, -11/6, dst.,
#     # fraksi bukan certificate.
#     if re.search(r"-\d+/\d+$", polis):
#         return ""

#     m = re.search(r"^.*?-(\d{1,6})(?:-\d+/\d+)?$", polis)

#     if not m:
#         return ""

#     return m.group(1).zfill(6)

def clean_certificate(polis_ori):
    if pd.isna(polis_ori):
        return ""

    polis = str(polis_ori).strip().upper()

    if not polis or polis == "-":
        return ""

    if re.search(r"\bSUSPENSE\b|\bENDORSEMENT\b", polis):
        return ""

    # Ambil certificate setelah polis utama.
    # 40010924046082-000667-1/0 -> 000667
    # 40010924046082-000689-1/1 -> 000689
    # 45013022023211-000047-3/1 -> 000047
    # 45013022021309-000005-12/9 -> 000005
    # 45013022023211-000047 -> 000047
    m = re.search(r"^.*?-(\d{1,6})(?:-\d+/\d+)?$", polis)

    if not m:
        return ""

    return m.group(1).zfill(6)

# ============================================================
# CLEAN SLIP
# ============================================================

def clean_slip(val) -> list:
    """
    Rule Etiqa:
    - 3010010924004217 -> 3010010924004217
    - 3010010924004217-28/9 -> 3010010924004217
    - 1071031124000024-000251-1/0 -> 1071031124000024-000251
    - CN/010/25/01/007263 -> CN0102501007263
    - beberapa slip dengan + tetap dipisahkan
    """

    if pd.isna(val):
        return []

    val = str(val).strip()

    if not val or val == "-":
        return []

    raw_parts = [p.strip() for p in val.split("+") if p.strip()]
    cleaned_parts = []

    for p in raw_parts:
        p = p.strip()

        # Hapus suffix fraksi paling belakang.
        # Contoh:
        # 3010010924004217-28/9 -> 3010010924004217
        # 1071031124000024-000251-1/0 -> 1071031124000024-000251
        p = re.sub(r"-.*$", "", p)
        
        # Hilangkan separator . dan /
        p = p.replace(".", "")
        p = p.replace("/", "")

        # Untuk nomor slip Etiqa, dash yang tersisa adalah bagian
        # nomor slip/pola nomor dan dipertahankan.
        # Contoh 1071031124000024-000251 tetap seperti itu.
        if p:
            cleaned_parts.append(p)

    return _cap_breakdown(cleaned_parts)


# ============================================================
# INSURED
# ============================================================

def _remove_polis_slip_from_text(text: str, polis_ori, slip_ori) -> str:
    if pd.notna(polis_ori):
        for token in [str(polis_ori).strip()] + clean_polis(polis_ori):
            if token and token != "-":
                text = text.replace(token, "")

    if pd.notna(slip_ori):
        for token in [str(slip_ori).strip()] + clean_slip(slip_ori):
            if token and token != "-":
                text = text.replace(token, "")

    return text


def _normalize_insured_part(p: str) -> str:
    p = _INSURED_TAIL_RE.sub("", p)
    p = _INSURED_TITLE_RE.sub("", p)

    p = re.sub(r"\bKB\b", "", p, flags=re.IGNORECASE)
    p = re.sub(r"\bA\.?W\.?\b", "", p, flags=re.IGNORECASE)

    p = re.sub(r"\(\s*\)", "", p)
    p = re.sub(r"\*\)", "", p)
    p = p.replace("*", "")

    p = re.sub(r"^(?:AND|OR)\b\s*", "", p, flags=re.IGNORECASE)
    p = re.sub(r"\s*\b(?:AND|OR)$", "", p, flags=re.IGNORECASE)

    p = re.sub(r"^[^a-zA-Z0-9(]+", "", p)
    p = re.sub(r"[^a-zA-Z0-9)]+$", "", p)

    return re.sub(r"\s+", " ", p).strip()


def _is_valid_insured_part(p: str) -> bool:
    if len(p) <= 2:
        return False

    up = p.upper()

    if re.match(r"^[\d\/\-\.]+$", up):
        return False

    if re.search(
        r"\b(?:NO\.\s*\d+|BUILDING|FLOOR|ROOM|ROAD|STREET|TOWER|KAV\.?|BLOK)\b",
        up,
    ):
        return False

    if re.search(
        r"\b(?:PLTGU|PLTMH|PLTU|POWER PLANT|COMBINED CYCLE|MW|HYDRO ELECTRIC)\b",
        up,
    ):
        return False

    return up not in _INSURED_JUNK_WORDS


def clean_insured(val, polis_ori, slip_ori) -> list:

    if pd.isna(val):
        return []

    val = str(val).strip()

    if not val:
        return []

    up = re.sub(r"\s+", " ", val).strip().upper()

    # ========================================================
    # KHUSUS INSURED - NAMA YANG HARUS DIANGGAP SATU KESATUAN
    # ========================================================
    SPECIAL_SINGLE_INSURED = [
        (
            r"PT\s+BUMI\s+MENTARI\s+KARYA\s*\(MUKO-MUKO\)",
            "BUMI MENTARI KARYA (MUKO-MUKO)"
        ),
        (
            r"PT\s+BSG\s+GASES\s*\(\s*PT\s+BEKASI\s+SEJATI\s+GAS\s*\)",
            "BSG GASES (BEKASI SEJATI GAS)"
        ),
        (
            r"KOPERASI\s+JASA\s+KARYAWAN\s+PT\s+BIO\s+FARMA",
            "KOPERASI JASA KARYAWAN BIO FARMA"
        ),
        (
            r"PRIMA\s+KARYA\s+HUSADA\s*\(RS\.\s*ROYAL\s+SURABAYA\)",
            "PRIMA KARYA HUSADA (RS. ROYAL SURABAYA)"
        ),
        (
            r"BIONET-ASIA\s+CO\.,?\s+LTD\.?",
            "BIONET-ASIA CO., LTD."
        ),
    ]

    for pattern, cleaned_name in SPECIAL_SINGLE_INSURED:
        if re.fullmatch(pattern, up, flags=re.IGNORECASE):
            return [cleaned_name]

    # Hapus nomor polis/slip yang ikut masuk ke insured
    val = _remove_polis_slip_from_text(up, polis_ori, slip_ori)

    # Hapus bagian deskriptif
    val = _INSURED_TAIL_RE.sub("", val)

    # Hapus keterangan subsidiary / parent subsidiary
    val = re.sub(
        r"\s+AND\s+ANY\s+SUBSIDIARY\b",
        "",
        val,
        flags=re.IGNORECASE,
    )

    val = re.sub(
        r"\s+AND\s+SUBSIDIARY\b",
        "",
        val,
        flags=re.IGNORECASE,
    )

    val = re.sub(
        r"\s+INCLUDING\s+ANY\s+PARENT[`']?S\s+SUBSIDIARY\b",
        "",
        val,
        flags=re.IGNORECASE,
    )

    # ========================================================
    # KHUSUS: AND/OR ... = KETERANGAN, BUKAN INSURED
    # Contoh:
    # PT BINTANG MAS GLASSOLUTIONS s and/or subsidiary
    # and/or associated and/or related companies
    # for their respective rights and interest
    # PT BANK CIMB NIAGA TBK
    #
    # Hasil:
    # BINTANG MAS GLASSOLUTIONS
    # BANK CIMB NIAGA TBK
    # ========================================================
    val = re.sub(
        r"\s+(?:S\s+)?AND\s*/\s*OR\s+"
        r"(?:SUBSIDIARY|SUBSIDIARIES|ASSOCIATED|"
        r"RELATED\s+COMPAN(?:Y|IES)|AFFILIATED|AFFILIATES)"
        r"(?:\s+AND\s*/\s*OR\s+"
        r"(?:SUBSIDIARY|SUBSIDIARIES|ASSOCIATED|"
        r"RELATED\s+COMPAN(?:Y|IES)|AFFILIATED|AFFILIATES))*"
        r"(?:\s+FOR\s+THEIR\s+RESPECTIVE\s+RIGHTS?\s+AND\s+INTERESTS?)?"
        r"(?=\s+PT\.?\s+)",
        " ",
        val,
        flags=re.IGNORECASE,
    )

    # Hapus isi dalam kurung
    val = re.sub(r"\([^)]*\)", "", val)

    # Normalisasi pemisah AND/OR
    val = re.sub(
        r"\s+(?:S\s+)?AND\s*/\s*OR\b.*?(?=\s+PT\.?\s+)",
        " ",
        val,
        flags=re.IGNORECASE,
    )

    val = re.sub(r"\bAND\s+OR\b", ",", val, flags=re.IGNORECASE)
    val = re.sub(r"\bCO\.,?\s*LTD\.?\b", ",", val, flags=re.IGNORECASE)

    # Hapus bentuk badan usaha tertentu
    for pattern in [
        r"\bTBK\.?\b",
        r"\(PERSERO\)",
        r"\bPERSERO\b",
        r"\bLTD\.?\b",
        r"\(FCI\.\s*I\.\)",
    ]:
        val = re.sub(pattern, "", val, flags=re.IGNORECASE)

    # QQ adalah pemisah insured pada data Etiqa.
    # PT. A QQ PT. B -> A ; B
    split_pattern = (
        r"\bQQ\b"
        r"|/"
        r"|,"
        r"|&"
        r"|\+"
        r"|-(?!\s*(?:19|20)\d{2}\b)"
        r"|\d+\."
        r"|\bPT\.?\b"
        r"|\bCV\.?\b"
        r"|:"
        r"|;"
    )

    parts = re.split(split_pattern, val, flags=re.IGNORECASE)

    cleaned = []

    for p in parts:
        p = _normalize_insured_part(p)

        if _is_valid_insured_part(p):
            cleaned.append(p.upper())

    # Hilangkan duplikat berurutan / duplikat persis
    final = []
    seen = set()

    for item in cleaned:
        if item not in seen:
            final.append(item)
            seen.add(item)

    return _cap_breakdown(final)


# ============================================================
# COLUMN EXPANSION
# ============================================================

def _expand_clean_columns(
    df: pd.DataFrame,
    all_lists: list,
    prefix: str,
    max_cols: int,
) -> list:

    added = []

    for i in range(1, max_cols + 1):
        col_name = f"clean {prefix} {i}"

        df[col_name] = [
            lst[i - 1] if i - 1 < len(lst) else None
            for lst in all_lists
        ]

        added.append(col_name)

    return added


# ============================================================
# PROCESS DATA
# ============================================================

def process_data(input_file: str, output_file: str) -> None:

    print(f"[1/6] Membaca data dari: {input_file} ...")

    cache_file = input_file + ".cache.pkl"
    if os.path.exists(cache_file) and os.path.getmtime(cache_file) >= os.path.getmtime(input_file):
        print(f"      Loading dari cache: {cache_file} ...")
        try:
            df = pd.read_pickle(cache_file)
        except Exception:
            print("      Cache rusak, membaca ulang dari Excel ...")
            df = pd.read_excel(
                input_file,
                sheet_name=SHEET_NAME,
                header=2,
            )
    else:
        df = pd.read_excel(
            input_file,
            sheet_name=SHEET_NAME,
            header=2,
        )
        try:
            df.to_pickle(cache_file)
            print(f"      Cache raw disimpan: {cache_file}")
        except Exception as e:
            print(f"      Gagal menyimpan cache raw: {e}")

    # ========================================================
    # FILTER STATUS - HANYA SUSPENSE
    # ADJUSTED DIBUANG PERMANEN DARI OUTPUT
    # ========================================================

    if "STATUS" not in df.columns:
        raise ValueError("Kolom STATUS tidak ditemukan di data.")

    sebelum_status = len(df)

    df = df[
        df["STATUS"]
        .fillna("")
        .astype(str)
        .str.strip()
        .str.upper()
        .eq("SUSPENSE")
    ].copy()

    sesudah_status = len(df)

    print(
        f"[STATUS] Hanya SUSPENSE yang dipertahankan : "
        f"{sesudah_status:,} baris"
    )
    print(
        f"[STATUS] ADJUSTED yang dibuang             : "
        f"{sebelum_status - sesudah_status:,} baris"
    )

    print(f"      Total baris keseluruhan : {len(df):,}")

    required_cols = [
        CEDANT_COL,
        "INSURED",
        "POLIS",
        "SLIP NO",
    ]

    missing = [c for c in required_cols if c not in df.columns]

    if missing:
        raise ValueError(
            f"Kolom wajib tidak ditemukan: {missing}"
        )

    # ========================================================
    # FILTER CEDANT MULTI ARTHA GUNA
    # ========================================================

    cedant_mask = (
        df[CEDANT_COL]
        .fillna("")
        .astype(str)
        .str.strip()
        .eq(CEDANT_VALUE)
    )
    df = df[cedant_mask].copy()

    print(
        f"[2/6] Data cedant "
        f"'{CEDANT_VALUE}' : {len(df):,} baris"
    )

    # ========================================================
    # RENAME ORIGINAL COLUMNS
    # ========================================================

    rename_map = {
        "INSURED": "insured_ori",
        "POLIS": "polis_ori",
        "SLIP NO": "slip_ori",
    }

    df.rename(
        columns={
            k: v
            for k, v in rename_map.items()
            if k in df.columns
        },
        inplace=True,
    )

    print("[3/6] Menjalankan cleansing Multi Artha Guna ...")

    all_clean_polis = []
    all_clean_slip = []
    all_clean_ins = []
    all_certificates = []

    max_polis = 1
    max_slip = 1
    max_ins = 1

    polis_series = df["polis_ori"].tolist() if "polis_ori" in df.columns else [""] * len(df)
    slip_series = df["slip_ori"].tolist() if "slip_ori" in df.columns else [""] * len(df)
    insured_series = df["insured_ori"].tolist() if "insured_ori" in df.columns else [""] * len(df)

    for p_ori, s_ori, i_ori in zip(polis_series, slip_series, insured_series):
        c_polis = clean_polis(p_ori)
        c_slip = clean_slip(s_ori)
        c_ins = clean_insured(
            i_ori,
            p_ori,
            s_ori,
        )
        c_certificate = clean_certificate(p_ori)

        max_polis = max(max_polis, len(c_polis))
        max_slip = max(max_slip, len(c_slip))
        max_ins = max(max_ins, len(c_ins))

        all_clean_polis.append(c_polis)
        all_clean_slip.append(c_slip)
        all_clean_ins.append(c_ins)
        all_certificates.append(c_certificate)

    max_polis = min(max_polis, MAX_BREAKDOWN)
    max_slip = min(max_slip, MAX_BREAKDOWN)
    max_ins = min(max_ins, MAX_BREAKDOWN)

    print(
        f"      Maks breakdown polis   : {max_polis}"
    )
    print(
        f"      Maks breakdown slip    : {max_slip}"
    )
    print(
        f"      Maks breakdown insured : {max_ins}"
    )

    # ========================================================
    # SUSUN OUTPUT
    # ========================================================

    print("[4/6] Menyusun kolom output ...")

    df["CERTIFICATE"] = all_certificates

    new_columns = []

    for col in list(df.columns):

        if col == "CERTIFICATE":
            continue

        if col == "polis_ori":

            new_columns.append(col)

            clean_polis_columns = _expand_clean_columns(
                df,
                all_clean_polis,
                "polis",
                max_polis,
            )

            if clean_polis_columns:

                # Urutan:
                # POLIS ORI
                # clean polis 1
                # CERTIFICATE
                # clean polis 2
                # clean polis 3 ...
                new_columns.append(
                    clean_polis_columns[0]
                )

                new_columns.append(
                    "CERTIFICATE"
                )

                new_columns += (
                    clean_polis_columns[1:]
                )

            else:
                new_columns.append(
                    "CERTIFICATE"
                )

        elif col == "slip_ori":

            new_columns.append(col)

            new_columns += _expand_clean_columns(
                df,
                all_clean_slip,
                "slip",
                max_slip,
            )

        elif col == "insured_ori":

            new_columns.append(col)

            new_columns += _expand_clean_columns(
                df,
                all_clean_ins,
                "insured",
                max_ins,
            )

        else:
            new_columns.append(col)

    df = df[new_columns]

    # ========================================================
    # SAVE
    # ========================================================

    print("[5/6] Menyimpan hasil ...")

    os.makedirs(
        os.path.dirname(output_file),
        exist_ok=True,
    )

    df.to_excel(
        output_file,
        index=False,
    )

    # Simpan juga cache .cache.pkl agar matching_mag.py dapat membaca instan
    output_cache = output_file + ".cache.pkl"
    try:
        data_records = df.to_dict("records")
        header_cols = list(df.columns)
        pd.to_pickle((data_records, header_cols), output_cache)
        print(f"      Cache output disimpan: {output_cache}")
    except Exception as e:
        print(f"      Gagal menyimpan cache output: {e}")

    print("[6/6] SELESAI")
    print("=" * 65)
    print(f"  Cedant       : {CEDANT_VALUE}")
    print(f"  Total baris  : {len(df):,}")
    print(f"  Total kolom  : {len(df.columns)}")
    print(f"  Output       : {output_file}")
    print("=" * 65)


# ============================================================
# ENTRY POINT
# ============================================================


if __name__ == "__main__":

    print("MASUK MAIN")

    process_data(
        INPUT_FILE,
        OUTPUT_FILE,
    )

# if __name__ == "__main__":
#     print("=" * 80)
#     print("TEST INSURED MAG - RULE BARU")
#     print("=" * 80)

#     test_cases = [
#         "PT BUMI MENTARI KARYA (MUKO-MUKO)",
#         "PT BSG GASES (PT BEKASI SEJATI GAS)",
#         "PT SRIBOGA FLOUR MILL and/or subsidiary and/or associated and/or related companies for their respective rights an interests and PT BANK MANDIRI (PERSERO) TBK",
#         "KOPERASI JASA KARYAWAN PT BIO FARMA",
#         "PRIMA KARYA HUSADA (RS. ROYAL SURABAYA)",
#         "BioNet-Asia Co., Ltd.",
#         "LIFERE AGRO KAPUAS AND ANY SUBSIDIARY",
#         "FERRON PAR PHARMACEUTICALS AND SUBSIDIARY",
#         "BUDI MUARATEX INCLUDING ANY PARENT`S SUBSIDIARY",
#     ]

#     for i, insured in enumerate(test_cases, 1):
#         hasil = clean_insured(
#             insured,
#             polis_ori="",
#             slip_ori=""
#         )

#         print(f"\nTEST {i}")
#         print("ORI   :", insured)
#         print("CLEAN :", hasil)
