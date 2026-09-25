import os
import re
import sys

sys.modules['numexpr'] = None
sys.modules['bottleneck'] = None

import numpy as np
import pandas as pd


INPUT_FILE  = os.path.join("data", "suspend.xlsx")
OUTPUT_FILE = os.path.join("data", "suspend_clean_aca.xlsx")

CEDANT_FILTER_COL   = "CEDANT SHRT NAME"
CEDANT_FILTER_VALUE = "CENTRAL"

# Regex untuk mengekstrak nomor sertifikat dari polis.
# Format 1: '{base_polis} - {1-6digit}', contoh: '100030825120000155 - 000211' → '000211'
# Format 2: '{base_polis}-{1-6digit}', contoh: '100030825120000155-001007' → '001007'
_CERT_FROM_POLIS_RE = re.compile(r'[-\s]+\s*(\d{1,6})(?:[^0-9]|$)')

# Regex untuk range S/D pada sertifikat: '{1-6digit}S/D{1-6digit}'
# Contoh: '000100S/D000110' atau '100 S/D 110'
_CERT_SD_RE = re.compile(r'(\d{1,6})\s*S/D\s*(\d{1,6})', re.IGNORECASE)

# Sertifikat valid: 1-6 digit pure numeric
_SERTIF_DIGIT_RE = re.compile(r'^\d{1,6}$')

# Max kolom sertifikat di output (maksimal 3 kolom sertif)
_MAX_SERTIF_COLS = 3

_INSURED_TAIL_RE = re.compile(
    r"""
    \bAS\b\s+(?:THE\s+)?(?:PRINCIPAL|OFF-TAKER|MAINTENANCE|CONTRACTOR).*
  | \bBEING\s+(?:THE\s+)?(?:PRINCIPAL|OFF-TAKER).*
  | \bAND\s+ALL\s+SUBSIDIARI.*
  | \bINCLUDING\s+ALL\s+SUBSIDIARI.*
  | \bINCLUDING\s+ANY\s+SUBSIDIAR.*
  | \bCOMPRISING\s+OF.*
  | \bINSTALLMENT\b.*
  | \bRELATED\s+COMPANY\b.*
  | \bPURCHASED\s+OR\s+OTHERWISE\b.*
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


# =============================================================================
# CLEAN FUNCTIONS
# =============================================================================

def clean_polis(val) -> list:
    """Strip numeric suffixes from policy number (e.g. -01, -02/03)."""
    if pd.isna(val):
        return []
    val = str(val).strip()
    if not val:
        return []

    p = val
    while True:
        stripped = re.sub(r"-(?:\d+(?:/\d+)?|EXT\(\d+\))$", "", p, flags=re.IGNORECASE)
        if stripped == p:
            break
        p = stripped

    return [p] if p else []


def clean_slip(val) -> list:
    """Return slip value as-is (no transformation needed)."""
    if pd.isna(val):
        return []
    val = str(val).strip()
    return [val] if val else []


def extract_cert_from_polis(val) -> list:
    """Ekstrak nomor sertifikat dari polis_ori di Suspend.

    Di Data 3 (Suspend), sertifikat menempel di kolom polis:
    Format 1 (suffix): '{16digit_polis} - {1-6digit_sertif}'
    Format 2 (S/D range): '{base_polis} - {1-6digit}S/D{1-6digit}'

    Aturan:
    - Cari suffix digit (1-6) setelah tanda '-' atau '- '
    - Jika format S/D: breakdown range, maksimal 5 nilai
    - Zero-pad setiap nilai ke 6 digit
    - Kembalikan list of strings (max _MAX_SERTIF_COLS = 5)

    Contoh:
        '100030825120000155-001007'       -> ['001007']
        '7283738299277344 - 000211'       -> ['000211']
        '1000308250000100S/D000110'       -> ['000100','000101','000102','000103','000104']
        '100030825120000155'              -> []  (tidak ada sertif)
    """
    if pd.isna(val):
        return []
    s = str(val).strip()
    if not s:
        return []

    # Cari pola S/D terlebih dahulu (lebih spesifik)
    m_sd = _CERT_SD_RE.search(s)
    if m_sd:
        start = int(m_sd.group(1))
        end   = int(m_sd.group(2))
        if start > end:
            start, end = end, start
        certs = [str(i).zfill(6) for i in range(start, end + 1)]
        return certs[:_MAX_SERTIF_COLS]   # max 5 kolom

    # Cari suffix digit (1-6) setelah tanda '-' atau '- '
    # Cek dari kanan untuk ambil suffix terakhir (nomor sertifikat)
    m = _CERT_FROM_POLIS_RE.search(s)
    if m:
        raw = m.group(1)
        if _SERTIF_DIGIT_RE.match(raw):
            return [raw.zfill(6)]

    return []


def _remove_polis_slip_from_text(text: str, polis_ori, slip_ori) -> str:
    """Remove policy and slip numbers from insured name text."""
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
    """Apply tail removal and normalization to one insured name segment."""
    p = _INSURED_TAIL_RE.sub("", p)
    p = re.sub(r"\bKB\b",    "", p, flags=re.IGNORECASE)
    p = re.sub(r"\bA\.?W\.?\b", "", p, flags=re.IGNORECASE)
    p = re.sub(r"\(\s*\)",   "", p)
    p = re.sub(r"^(?:AND|OR)\b\s*", "", p, flags=re.IGNORECASE)
    p = re.sub(r"\s*\b(?:AND|OR)$", "", p, flags=re.IGNORECASE)
    p = re.sub(r"^[^a-zA-Z0-9(]+", "", p)
    p = re.sub(r"[^a-zA-Z0-9)]+$", "", p)
    return p.strip()


def _is_valid_insured_part(p: str) -> bool:
    """Return True if an insured name segment is worth keeping."""
    if len(p) <= 2:
        return False
    up = p.upper()
    if re.match(r"^[\d\/\-\.]+$", up):
        return False
    if re.search(r"\b(?:NO\.\s*\d+|BUILDING|FLOOR|ROOM|ROAD|STREET|TOWER|KAV\.?|BLOK)\b", up):
        return False
    if re.search(r"\b(?:PLTGU|PLTMH|PLTU|POWER PLANT|COMBINED CYCLE|MW|HYDRO ELECTRIC)\b", up):
        return False
    return up not in _INSURED_JUNK_WORDS


def clean_insured(val, polis_ori, slip_ori) -> list:
    """Clean insured name by removing policy/slip numbers and junk words."""
    if pd.isna(val):
        return []
    val = str(val).strip()
    if not val:
        return []

    val = _remove_polis_slip_from_text(val, polis_ori, slip_ori)

    val = re.sub(r"\(\s*[\d\.\/\-]+\s*\)", "", val)
    val = re.sub(r"\b(?:AND|AN|OR)\s*/\s*(?:AND|OR)\b", ",", val, flags=re.IGNORECASE)
    val = re.sub(r"\bAND\s+OR\b", ",", val, flags=re.IGNORECASE)
    val = re.sub(r"\bCO\.,?\s*LTD\.?\b", ",", val, flags=re.IGNORECASE)
    for pattern in [r"\bTBK\.?\b", r"\(PERSERO\)", r"\bPERSERO\b", r"\bLTD\.?\b", r"\(FCI\.\s*I\)"]:
        val = re.sub(pattern, "", val, flags=re.IGNORECASE)

    split_pattern = r"\bQQ\b|/|,|-(?!\s*(?:19|20)\d{2}\b)|\d+\.|\bPT\.?\b|\bCV\.?\b|:|;"
    parts = re.split(split_pattern, val, flags=re.IGNORECASE)

    cleaned = []
    for p in parts:
        p = _normalize_insured_part(p)
        if _is_valid_insured_part(p):
            cleaned.append(p)

    return cleaned


# =============================================================================
# COLUMN EXPANSION
# =============================================================================

def _expand_clean_columns(df: pd.DataFrame, all_lists: list, ori_col: str, prefix: str, max_cols: int) -> list:
    """Add clean_{prefix}_1..N columns after ori_col. Returns list of added column names."""
    added = []
    for i in range(1, max_cols + 1):
        col_name = f"clean {prefix} {i}"
        df[col_name] = [lst[i - 1] if i - 1 < len(lst) else "" for lst in all_lists]
        added.append(col_name)
    return added


# =============================================================================
# PROCESS DATA
# =============================================================================

def process_data(input_file: str, output_file: str) -> None:
    print(f"[1/5] Reading: {input_file} ...")
    df = pd.read_excel(input_file, header=2)
    print(f"      Total rows: {len(df):,}")

    # Remove illegal XML control characters (e.g. \x1f) that corrupt Excel workbooks
    for c in df.select_dtypes(include=['object']).columns:
        df[c] = (df[c]
                 .fillna('')
                 .astype(str)
                 .str.replace(r'[\x00-\x08\x0B\x0C\x0E-\x1F]', '', regex=True))

    if CEDANT_FILTER_COL not in df.columns:
        print(f"\n[ERROR] Column '{CEDANT_FILTER_COL}' not found. Available: {list(df.columns)}")
        return

    df = df[df[CEDANT_FILTER_COL] == CEDANT_FILTER_VALUE].copy()
    print(f"[2/5] Filter '{CEDANT_FILTER_VALUE}': {len(df):,} rows.")

    if "STATUS" in df.columns:
        df = df[df["STATUS"].astype(str).str.strip().str.upper() == "SUSPENSE"].copy()
        print(f"      Filter STATUS 'SUSPENSE': {len(df):,} rows.")

    if df.empty:
        print("\n[WARN] No data after filter. Stopping.")
        return

    rename_map = {"INSURED": "insured_ori", "POLIS": "polis_ori", "SLIP NO": "slip_ori"}
    df.rename(columns={k: v for k, v in rename_map.items() if k in df.columns}, inplace=True)

    print("[3/5] Cleaning ...")

    # Vectorized apply — jauh lebih cepat dari iterrows() untuk dataset besar.
    # clean_insured butuh 2 kolom (polis_ori + slip_ori) sehingga pakai apply(axis=1).
    print(f"      Cleaning polis & slip ({len(df):,} baris) ...", flush=True)
    all_polis = df["polis_ori"].map(clean_polis).tolist()
    all_slip  = df["slip_ori"].map(clean_slip).tolist()

    print(f"      Cleaning insured ...", flush=True)
    all_ins = df.apply(
        lambda row: clean_insured(row.get("insured_ori", ""),
                                  row.get("polis_ori",   ""),
                                  row.get("slip_ori",    "")),
        axis=1,
    ).tolist()

    max_polis = max(1, max((len(x) for x in all_polis), default=1))
    max_slip  = max(1, max((len(x) for x in all_slip),  default=1))
    max_ins   = max(1, max((len(x) for x in all_ins),   default=1))

    # Ekstrak sertifikat 6-digit dari polis_ori (sertifikat menempel di polis untuk Suspend)
    print(f"      Extracting sertifikat dari polis ...", flush=True)
    all_sertif = df["polis_ori"].map(extract_cert_from_polis).tolist()
    max_sertif = max(1, max((len(x) for x in all_sertif), default=1))

    print("[4/5] Building output columns ...")

    new_columns = []
    for col in df.columns:
        new_columns.append(col)
        if col == "polis_ori":
            new_columns += _expand_clean_columns(df, all_polis,  col, "polis",   max_polis)
            # Tambahkan kolom sertifikat tepat setelah kolom polis bersih
            new_columns += _expand_clean_columns(df, all_sertif, col, "sertif",  max_sertif)
        elif col == "slip_ori":
            new_columns += _expand_clean_columns(df, all_slip,   col, "slip",    max_slip)
        elif col == "insured_ori":
            new_columns += _expand_clean_columns(df, all_ins,    col, "insured", max_ins)

    df = df[new_columns]

    print(f"[5/5] Saving to: {output_file} ...")
    df.to_excel(output_file, index=False)

    try:
        from excel_styler import apply_purple_column_style
        apply_purple_column_style(output_file, "suspend_clean")
    except Exception as e:
        print(f"  [WARN] Styling failed: {e}")

    print(f"\n{'=' * 50}")
    print(f"  Done. {len(df):,} rows, {len(df.columns)} cols -> {output_file}")
    print(f"{'=' * 50}")


if __name__ == "__main__":
    process_data(INPUT_FILE, OUTPUT_FILE)
