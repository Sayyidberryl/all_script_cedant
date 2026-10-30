import os
import re

import numpy as np
import pandas as pd

print("SCRIPT BERJALAN")

# ─────────────────────────────────────────────────────────────────────────────
# KONFIGURASI
# ─────────────────────────────────────────────────────────────────────────────

INPUT_FILE  = os.path.join("data_raw", "3b. Database Suspense 150826.xlsx")
OUTPUT_FILE = os.path.join("data_cleaned", "tokio_output_suspend.xlsx")
SHEET_NAME  = 0

# Filter cedant: hanya proses dan simpan cedant Tokio
CEDANT_KEYWORD = "TOKIO"

# Header output suspend (urutan & nama kolom final)
OUTPUT_HEADERS = [
    "RECEIPT NO", "CREDIT NOTES", "DETAIL RINCIAN NO", "RECEIPT DATE",
    "CEDANT NAME", "CEDANT SHRT NAME",
    "INSURED", "INSURED_CLEAN",
    "CURR ORI", "AMOUNT ORI", "CURR PAY", "AMOUNT PAY",
    "POLIS", "POLIS_CLEAN", "CERTIFICATE_1",
    "SLIP NO", "SLIP_NO_CLEAN",
    "DESC 1", "DESC 2", "DESC 3", "DESC 4",
    "STATUS", "REC_TYPE", "CEK FASE",
]

# Kalau hasil clean terdiri dari >1 bagian, digabung dalam satu sel dengan pemisah ini
CLEAN_JOIN_SEP = "; "

# Business rule: kalau hasil breakdown (insured/polis/slip) > 5 bagian,
# tidak usah dipecah per kolom -> digabung lagi jadi satu kolom.
MAX_BREAKDOWN = 5

# Regex untuk sub-removal dalam nama insured
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

# Gelar / sapaan yang ikut dihapus dari nama insured
_INSURED_TITLE_RE = re.compile(
    r"\b(?:BAPAK|IBU|BPK|NY\.?|SDR\.?|SDRI\.?|TN\.?|MR\.?|MRS\.?|MS\.?)\b\s*",
    re.IGNORECASE,
)


# ─────────────────────────────────────────────────────────────────────────────
# CLEAN FUNCTIONS
# ─────────────────────────────────────────────────────────────────────────────

def _cap_breakdown(parts: list, sep: str = "; ") -> list:
    """
    Business rule: jika jumlah bagian hasil breakdown > MAX_BREAKDOWN,
    jangan dipecah per kolom -> gabungkan lagi jadi satu nilai.
    """
    parts = [p for p in parts if p]
    if len(parts) > MAX_BREAKDOWN:
        return [sep.join(parts)]
    return parts


def clean_polis(val) -> list:
    """
    Clean nomor polis.

    Rules:
    - Hapus prefix seperti TMI/FEAQ/23-
    - Hapus suffix -000-01, -000-02, dst.
    - Hapus suffix numerik lain seperti -01 atau -02/03
    - Hapus titik (.)
    - Split jika dipisahkan dengan +
    """
    if pd.isna(val):
        return []

    val = str(val).strip()
    if not val or val == "-":
        return []

    # Hapus koma
    val = val.replace(",", "")

    # Split jika ada lebih dari satu polis
    raw_parts = [p.strip() for p in val.split("+") if p.strip()]
    cleaned_parts = []

    for p in raw_parts:
        cur = p.strip()

        # Hapus suffix
        cur = re.sub(r"-\d{3}-\d+$", "", cur)
        cur = re.sub(r"-000-\d+$", "", cur)

        while True:
            stripped = re.sub(r"-\d+(?:/\d+)?$", "", cur)
            if stripped == cur:
                break
            cur = stripped

        # Hapus tanda pemisah
        cur = cur.replace(".", "")
        cur = cur.replace("/", "")
        cur = cur.replace("-", "")

        if cur:
            cleaned_parts.append(cur)

    return _cap_breakdown(cleaned_parts)

def clean_Certificate(polis_ori):
    """
    Mengambil Certificate dari POLIS ORI (Khusus Tokio Marine).
    - Jika bernilai -000- atau -0 -> Dianggap BUKAN Sertifikat (Kembalikan BLANK "")
    - Jika bernilai -006- -> Kembalikan "000006"
    """
    if pd.isna(polis_ori):
        return ""

    polis = str(polis_ori).strip().upper()

    if not polis:
        return ""

    # Hapus data kotor / suspense
    if re.search(r"\bSUSPENSE\b|\bENDORSEMENT\b", polis):
        return ""

    # 1. Handling Pola Tokio Marine: -XXX-YY
    m_tokio = re.search(r"-(\d{3})-\d+$", polis)
    if m_tokio:
        val = m_tokio.group(1)
        # Jika bernilai 000 / 0 -> Anggap BUKAN sertifikat (BLANK)
        if int(val) == 0:
            return ""
        return val.zfill(6)

    # 2. Handling Pola Single Dash Biasa (-006 / -000006)
    m_single = re.search(r"-(\d{1,6})$", polis)
    if m_single:
        val = m_single.group(1)
        if int(val) == 0:
            return ""
        return val.zfill(6)

    return ""
def clean_slip(val) -> list:
    """Kembalikan nilai slip (tanpa titik), pecah kalau ada lebih dari satu (+)."""
    if pd.isna(val):
        return []
    val = str(val).strip()
    if not val or val == "-":
        return []

    raw_parts = [
        p.strip()
        .replace(".", "")
        .replace("/", "")
        .replace("-", "")
        for p in val.split("+")
        if p.strip()
    ]
    return _cap_breakdown(raw_parts)


def _remove_polis_slip_from_text(text: str, polis_ori, slip_ori) -> str:
    """Hapus nomor polis & slip dari teks insured secara agresif."""
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
    """Terapkan sub-removal dan normalisasi pada satu bagian nama insured."""
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

    return p.strip()


def _is_valid_insured_part(p: str) -> bool:
    """Return True jika bagian insured layak dipertahankan."""
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
    """Bersihkan nama insured dengan menghapus nomor polis/slip, gelar, dan junk words."""
    if pd.isna(val):
        return []

    val = str(val).strip()
    if not val:
        return []

    if re.search(r"KSO\s+MCC19-KJK\s*\(MMP\s*PROJECT\)", val, re.IGNORECASE):
        return [re.sub(r"\s{2,}", " ", val).strip().upper()]

    val = _remove_polis_slip_from_text(val, polis_ori, slip_ori)
    val = _INSURED_TAIL_RE.sub("", val)
    val = re.sub(r"\([^)]*\)", "", val)

    val = re.sub(
        r"\b(?:AND|AN|OR)\s*/\s*(?:AND|OR)\b",
        ",",
        val,
        flags=re.IGNORECASE,
    )
    val = re.sub(r"\bAND\s+OR\b", ",", val, flags=re.IGNORECASE)
    val = re.sub(r"\bCO\.,?\s*LTD\.?\b", ",", val, flags=re.IGNORECASE)

    for pattern in [
        r"\bTBK\.?\b",
        r"\(PERSERO\)",
        r"\bPERSERO\b",
        r"\bLTD\.?\b",
        r"\(FCI\.\s*I\)",
    ]:
        val = re.sub(pattern, "", val, flags=re.IGNORECASE)

    split_pattern = (
        r"\bQQ\b|/|,|&|\+"
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

    return _cap_breakdown(cleaned)


# ─────────────────────────────────────────────────────────────────────────────
# COLUMN EXPANSION HELPER
# ─────────────────────────────────────────────────────────────────────────────

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


# ─────────────────────────────────────────────────────────────────────────────
# PROCESS DATA
# ─────────────────────────────────────────────────────────────────────────────

def process_data(input_file: str, output_file: str) -> None:
    print(f"[1/5] Membaca data dari: {input_file} (sheet: {SHEET_NAME}) ...")
    df = pd.read_excel(input_file, sheet_name=SHEET_NAME, header=2)
    print(f"      Total baris keseluruhan: {len(df):,}")

    # ========================================================
    # FILTER STATUS - HANYA SUSPENSE
    # ADJUSTED DIBUANG PERMANEN
    # ========================================================

    if "STATUS" not in df.columns:
        raise ValueError("Kolom STATUS tidak ditemukan di data.")

    total_sebelum = len(df)

    df = df[
        df["STATUS"]
        .fillna("")
        .astype(str)
        .str.strip()
        .str.upper()
        .eq("SUSPENSE")
    ].copy()

    total_sesudah = len(df)
    total_adjusted = total_sebelum - total_sesudah

    print(f"      Data SUSPENSE : {total_sesudah:,} baris")
    print(f"      Data ADJUSTED : {total_adjusted:,} baris dibuang")

    # ========================================================
    # FILTER CEDANT - HANYA TOKIO
    # CEDANT LAIN DIBUANG PERMANEN
    # ========================================================
    ced_cols = [c for c in df.columns if "CEDANT" in str(c).upper()]
    if not ced_cols:
        raise ValueError("Kolom CEDANT tidak ditemukan di data.")

    mask_tokio = pd.Series(False, index=df.index)
    for c in ced_cols:
        mask_tokio |= df[c].fillna("").astype(str).str.upper().str.contains(CEDANT_KEYWORD)

    total_sebelum_cedant = len(df)
    df = df[mask_tokio].reset_index(drop=True)
    total_non_tokio = total_sebelum_cedant - len(df)

    print(f"      Data TOKIO    : {len(df):,} baris")
    print(f"      Non-TOKIO     : {total_non_tokio:,} baris dibuang")
    print(f"[2/5] Total baris yang akan diproses: {len(df):,}")

    # Rename kolom asli -> _ori
    rename_map = {
        "INSURED": "insured_ori",
        "POLIS": "polis_ori",
        "SLIP NO": "slip_ori",
    }

    df.rename(
        columns={k: v for k, v in rename_map.items() if k in df.columns},
        inplace=True,
    )

    print("[3/5] Menjalankan proses cleaning untuk baris TOKIO ...")

    all_clean_polis = []
    all_clean_slip = []
    all_clean_ins = []
    all_Certificates = []  # <--- 1. TAMBAH PENAMPUNG Certificate

    max_polis = 1
    max_slip = 1
    max_ins = 1

    for idx, (_, row) in enumerate(df.iterrows()):
        p_ori = row.get("polis_ori", "")
        s_ori = row.get("slip_ori", "")
        i_ori = row.get("insured_ori", "")

        c_polis = clean_polis(p_ori)
        c_slip = clean_slip(s_ori)
        c_ins = clean_insured(i_ori, p_ori, s_ori)
        c_cert = clean_Certificate(p_ori)  # <--- 2. PANGGIL CLEAN_Certificate

        max_polis = max(max_polis, len(c_polis))
        max_slip = max(max_slip, len(c_slip))
        max_ins = max(max_ins, len(c_ins))

        all_clean_polis.append(c_polis)
        all_clean_slip.append(c_slip)
        all_clean_ins.append(c_ins)
        all_Certificates.append(c_cert)  # <--- 3. SIMPAN KE LIST

    # Safety net
    max_polis = min(max_polis, MAX_BREAKDOWN)
    max_slip = min(max_slip, MAX_BREAKDOWN)
    max_ins = min(max_ins, MAX_BREAKDOWN)

    print("[4/5] Menyusun kolom output ...")

    def _join(lst: list):
        return CLEAN_JOIN_SEP.join(lst) if lst else None

    df["INSURED_CLEAN"] = [_join(x) for x in all_clean_ins]
    df["POLIS_CLEAN"] = [_join(x) for x in all_clean_polis]
    df["CERTIFICATE_1"] = all_Certificates

    _expand_clean_columns(df, all_clean_ins, "insured", max_ins)
    _expand_clean_columns(df, all_clean_polis, "polis", max_polis)
    _expand_clean_columns(df, all_clean_slip, "slip", max_slip)

    # Kembalikan nama kolom asli (sebelumnya di-rename jadi *_ori)
    df = df.rename(columns={
        "insured_ori": "INSURED",
        "polis_ori": "POLIS",
        "slip_ori": "SLIP NO",
    })

    missing = [c for c in OUTPUT_HEADERS if c not in df.columns]
    if missing:
        raise ValueError(f"Kolom berikut tidak ada di data: {missing}")

    extra = [c for c in df.columns if c not in OUTPUT_HEADERS]
    if extra:
        print(f"      [INFO] Kolom di luar header yang ditentukan (ditaruh di paling kanan): {extra}")

    df = df[OUTPUT_HEADERS + extra]

    print(f"[5/5] Menyimpan hasil ke: {output_file} ...")

    os.makedirs(os.path.dirname(output_file), exist_ok=True)
    try:
        df.to_excel(output_file, index=False)
    except PermissionError:
        base, ext = os.path.splitext(output_file)
        alt_path = f"{base}_new{ext}"
        print(f"\n[WARNING] File '{output_file}' sedang dibuka di aplikasi lain.")
        print(f"          Menyimpan hasil ke file alternatif: '{alt_path}'")
        df.to_excel(alt_path, index=False)
        output_file = alt_path

    print(f"\n{'=' * 55}")
    print("  [OK] Selesai!")
    print(f"  Total baris output  : {len(df):,}")
    print(f"  Total kolom output  : {len(df.columns)}")
    print(f"  File disimpan di    : {output_file}")
    print(f"{'=' * 55}")


# ─────────────────────────────────────────────────────────────────────────────
# ENTRY POINT
# ─────────────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    print("MASUK MAIN")
    process_data(INPUT_FILE, OUTPUT_FILE)