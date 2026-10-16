"""
prod_sus_1_debug.py — Script Debugging & Peningkatan Matching V1 (Cedant KBRU)

Salinan lengkap dari logika v1 dengan implementasi Solusi Terbaik:
1. Sanitasi Dummy Values:
   - Mengabaikan tanda strip/placeholder ('-', '--', '/', '0', 'SELISIH', 'NONE', 'NULL')
     agar tidak salah dianggap sebagai nomor Slip/Polis valid (memperbaiki 73 baris 'Slip only (> 1 Fac)').
2. Graceful FACUL Fallback (Saldo 0 / Historical Placements):
   - Jika transaksi cocok di FACUL tetapi saldo OSBAL sudah tidak ada / lunas,
     FAC_CODE tetap dipertahankan dan dicatat (sebagai New Entry / Saldo 0 di OSBAL),
     tidak langsung dibuang menjadi 'Unmatching' buta.
3. Enhanced Normalization Polis & Slip:
   - Deteksi nomor polis inti tanpa broker suffix (-SRB, -KBRURE) dan varian base slip.
4. Output Terisolasi & Komparasi Otomatis:
   - Output disimpan ke data/final_output_v1_debug.xlsx
   - Mencetak tabel perbandingan v1 Asli vs v1 Debug di akhir eksekusi.
"""

import os
import re
import sys
import time
import functools

sys.modules['numexpr'] = None
sys.modules['bottleneck'] = None

import numpy as np
import pandas as pd
from dotenv import load_dotenv

# Path folder
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
ROOT_DIR   = os.path.dirname(SCRIPT_DIR)
DATA_DIR   = os.path.join(ROOT_DIR, "data")

SUSPEND_FILE = os.path.join(DATA_DIR, "kbru_output_suspend.xlsx")
OSBAL_FILE   = os.path.join(DATA_DIR, "kbru_output_osbal.xlsx")
FACUL_FILE   = os.path.join(DATA_DIR, "kbru_output_facul.xlsx")
OUTPUT_DEBUG = os.path.join(DATA_DIR, "final_output_v1_debug.xlsx")

def _find_slipdb_file() -> str:
    kandidat = [
        os.path.join(DATA_DIR, "kbru_output_slipdb.xlsx"),
        os.path.join(DATA_DIR, "slipdb_clean_kbru.xlsx"),
        os.path.join(DATA_DIR, "ri slip.xlsx"),
        os.path.join(DATA_DIR, "slipdb.xlsx"),
    ]
    for p in kandidat:
        if os.path.exists(p):
            return p
    return kandidat[0]

SLIPDB_FILE = _find_slipdb_file()

OSBAL_FACODE_COL  = "CCOS_REF_CODE"
FACUL_FACODE_COL  = "FAC_CODE"
SLIPDB_FACODE_COL = "FAC_CODE"

SUSPEND_CURR_COL = "CURR ORI"
OSBAL_CURR_COL   = "CCOS_CURR"
SUSPEND_DATE_COL = "RECEIPT DATE"
OSBAL_DATE_COL   = "FAC_COM_DATE"

_EXCLUDED_REF_MARKERS = ["HUTANG PIUTANG", "DATA SUSPENSE"]
_MIN_TOKEN_LEN = 3
_TOKEN_RE      = re.compile(r"[^A-Z0-9]+")

# [SOLUSI 1]: Set nilai dummy yang TIDAK BOLEH dianggap sebagai Slip / Polis valid
_DUMMY_VALUES = frozenset([
    "-", "--", "---", "/", "//", "0", "00", "000", "NONE", "NULL", "NAN",
    "SELISIH", "BIAYA", "TITIPAN", "KOREKSI", "REFUND", "RETUR", "ADMIN", "TEST"
])

def _normalize(val) -> str:
    if val is None or pd.isna(val):
        return ""
    return str(val).strip().upper()

def _clean_alphanumeric(val: str) -> str:
    if not val:
        return ""
    return re.sub(r'[^A-Z0-9]', '', str(val).upper())

def _is_valid_slip_or_polis(val: str) -> bool:
    """Memeriksa apakah string merupakan nomor slip/polis riil, bukan simbol/placeholder dummy."""
    if not val:
        return False
    norm = val.strip().upper()
    if norm in _DUMMY_VALUES:
        return False
    # Bersihkan karakter non-alphanumeric
    alphanumeric = re.sub(r'[^A-Z0-9]', '', norm)
    # Jika tidak ada huruf/angka atau panjang kurang dari 3 karakter
    if len(alphanumeric) < 3:
        return False
    # Jika hanya angka nol berulang
    if set(alphanumeric) == {'0'}:
        return False
    return True

@functools.lru_cache(maxsize=100000)
def _tokenize(text: str) -> frozenset:
    if not text:
        return frozenset()
    return frozenset(t for t in _TOKEN_RE.split(text) if len(t) >= _MIN_TOKEN_LEN and t not in _DUMMY_VALUES)

def _get_clean_cols(columns, prefix: str) -> list:
    if columns is None or len(columns) == 0:
        return []
    alt_map = {
        "clean polis": ["clean polis", "policy_clean", "polis_clean", "polis_cln", "policy_cln"],
        "clean slip": ["clean slip", "slip_clean", "fac_slip_clean", "slip_no_cln", "slip_cln"],
        "clean insured": ["clean insured", "insured_clean", "fac_insured_", "insured_1", "insured_2", "insured_3", "insured_4", "insured_5"],
    }
    prefixes = alt_map.get(prefix.lower(), [prefix.lower()])
    res = []
    for c in columns:
        clow = str(c).lower().strip()
        if any(clow.startswith(p) for p in prefixes):
            res.append(c)
    return res

def _is_excluded_ref_row(row: dict, polis_cols: list, slip_cols: list) -> bool:
    for col in list(polis_cols) + ["polis_ori", "FAC_POLICY_NO_ORI"] + list(slip_cols) + ["slip_ori", "FAC_SLIP_ORI"]:
        val = _normalize(row.get(col, ""))
        if val and any(m in val for m in _EXCLUDED_REF_MARKERS):
            return True
    return False

def _collect_clean_values(row: dict, cols: list, is_slip: bool = False, is_polis: bool = False) -> list:
    seen, result = set(), []
    for col in cols:
        val = _normalize(row.get(col, ""))
        # [SOLUSI 1]: Validasi nilai dummy
        if (is_slip or is_polis) and not _is_valid_slip_or_polis(val):
            continue
        if val and val not in seen:
            seen.add(val)
            result.append(val)
            if is_slip:
                # Strip suffix installment -NN jika ada (misal: 11055124F00000144-03 -> 11055124F00000144)
                base = re.sub(r'-\d+$', '', val)
                if base and _is_valid_slip_or_polis(base) and base not in seen:
                    seen.add(base)
                    result.append(base)
                c_num = _clean_alphanumeric(val)
                if c_num and len(c_num) >= 3 and c_num not in seen:
                    seen.add(c_num)
                    result.append(c_num)
                c_base = _clean_alphanumeric(base)
                if c_base and len(c_base) >= 3 and c_base not in seen:
                    seen.add(c_base)
                    result.append(c_base)
            elif is_polis:
                # [SOLUSI 3]: Ekstrak varian nomor polis inti tanpa broker suffix (-SRB, -KBRURE)
                # Contoh: 00004903/FEQI/054.0000002-SRB/08/2023 -> 00004903/FEQI/054.0000002
                cleaned_pol = re.sub(r'-(SRB|KBRURE|KBRU|SIMAS|ACA|BRK)[^/]*', '', val)
                if cleaned_pol and cleaned_pol != val and cleaned_pol not in seen:
                    seen.add(cleaned_pol)
                    result.append(cleaned_pol)
                c_pol = _clean_alphanumeric(val)
                if c_pol and len(c_pol) >= 5 and c_pol not in seen:
                    seen.add(c_pol)
                    result.append(c_pol)
    return result

# =============================================================================
# INDEXING & MATCH ENGINE
# =============================================================================

def _build_exact_index(rows: list, cols: list, is_slip: bool = False, is_polis: bool = False, excluded: set = None) -> dict:
    index: dict = {}
    for i, row in enumerate(rows):
        if excluded and i in excluded:
            continue
        vals = _collect_clean_values(row, cols, is_slip=is_slip, is_polis=is_polis)
        for val in vals:
            index.setdefault(val, []).append(i)
    return index

def _build_facode_index(rows: list, col: str, excluded: set = None) -> dict:
    index: dict = {}
    for i, row in enumerate(rows):
        if excluded and i in excluded:
            continue
        fac = _normalize(row.get(col, ""))
        if fac:
            index.setdefault(fac, []).append(i)
    return index

def _build_token_index(rows: list, clean_cols: list, excluded: set = None) -> tuple:
    token_index: dict = {}
    token_cache: dict = {}
    for i, row in enumerate(rows):
        if excluded and i in excluded:
            continue
        tokens: frozenset = frozenset()
        for c in clean_cols:
            t = _normalize(row.get(c, ""))
            if t:
                tokens = tokens | _tokenize(t)
        token_cache[i] = tokens
        for tok in tokens:
            token_index.setdefault(tok, set()).add(i)
    return token_index, token_cache

def _build_lookup(rows: list, polis_cols: list, slip_cols: list, insured_cols: list,
                  facode_col: str = None, label: str = "") -> tuple:
    if not rows:
        if label:
            print(f"  {label:<6} done (0.0s) [kosong]", flush=True)
        return ({}, {}, {}, ({}, {}), ({}, {}), ({}, {}), {})

    t0 = time.perf_counter()
    excluded = {i for i, r in enumerate(rows) if _is_excluded_ref_row(r, polis_cols, slip_cols)}
    facode_index = _build_facode_index(rows, facode_col, excluded=excluded) if facode_col else {}
    lookup = (
        _build_exact_index(rows, slip_cols,    is_slip=True,  is_polis=False, excluded=excluded), # [0]
        _build_exact_index(rows, polis_cols,   is_slip=False, is_polis=True,  excluded=excluded), # [1]
        _build_exact_index(rows, insured_cols, is_slip=False, is_polis=False, excluded=excluded), # [2]
        _build_token_index(rows, slip_cols,    excluded=excluded),                                # [3]
        _build_token_index(rows, polis_cols,   excluded=excluded),                                # [4]
        _build_token_index(rows, insured_cols, excluded=excluded),                                # [5]
        facode_index,                                                                             # [6]
    )
    elapsed = time.perf_counter() - t0
    if label:
        exc_info = f" - {len(excluded):,} baris dikecualikan" if excluded else ""
        print(f"  {label:<6} done ({elapsed:.1f}s){exc_info}", flush=True)
    return lookup

def _exact_match(query_values: list, index: dict) -> set:
    return {i for v in query_values if v and v in index for i in index[v]}

def _like_match(query_values: list, token_index: dict, rows: list, ref_cols: list) -> list:
    valid_queries = [v for v in query_values if v and len(v) >= 3 and v not in _DUMMY_VALUES]
    if not valid_queries:
        return []

    query_tokens: frozenset = frozenset()
    for v in valid_queries:
        query_tokens = query_tokens | _tokenize(v)

    candidates: set = set()
    for tok in query_tokens:
        if tok in token_index:
            candidates |= token_index[tok]
    if not candidates:
        return []

    matched = []
    for i in candidates:
        ref_values = [_normalize(rows[i].get(c, "")) for c in ref_cols]
        for q in valid_queries:
            if any(q in rv for rv in ref_values if rv):
                matched.append(i)
                break
    return matched

# [ATURAN WAJIB]: FAC CODE dari SLIPDB / FACUL HANYA BOLEH DIPAKAI JIKA ADA DI OSBAL
def _resolve_facode(ref_rows_matched: list, facode_col: str, facode_osbal_idx: dict, osbal_rows: list) -> tuple:
    """
    Mencari apakah FAC_CODE ada di OSBAL.
    Sesuai aturan bisnis: OSBAL adalah daftar tagihan aktif.
    Jika FAC_CODE tidak ada di OSBAL, berarti tagihan sudah tidak ada / belum ada,
    sehingga TIDAK BOLEH DIGUNAKAN (return [], False).
    """
    fac_codes = set()
    for r in ref_rows_matched:
        fac = _normalize(r.get(facode_col, ""))
        if fac:
            fac_codes.add(fac)
    if not fac_codes:
        return [], False

    osbal_indices: set = set()
    for fac in fac_codes:
        if fac in facode_osbal_idx:
            osbal_indices.update(facode_osbal_idx[fac])

    res = list(osbal_indices)
    return res, bool(res)

def _aggregate_ccos(osbal_rows: list) -> dict:
    fac_codes = {str(r.get("CCOS_REF_CODE", "")).strip() for r in osbal_rows if str(r.get("CCOS_REF_CODE", "")).strip()}
    multi_fac = len(fac_codes) > 1

    def join_unique(key):
        seen, vals = set(), []
        for r in osbal_rows:
            v = str(r.get(key, "")).strip()
            if v and v not in seen:
                seen.add(v)
                vals.append(v)
        return ", ".join(vals)

    def sum_col(key, rows):
        if multi_fac:
            return np.nan
        total = 0.0
        for r in rows:
            v = r.get(key, 0)
            if pd.isna(v) or v == "":
                continue
            if isinstance(v, (int, float)):
                total += float(v)
            else:
                try:
                    s = str(v).replace(',', '')
                    total += float(s)
                except ValueError:
                    pass
        return total

    return {
        "CCOS_DOC_NO":   join_unique("CCOS_DOC_NO"),
        "CCOS_REF_CODE": join_unique("CCOS_REF_CODE"),
        "CCOS_OR_BAL":   np.nan if multi_fac else join_unique("CCOS_OR_BAL"),
        "CCOS_BAL_DUE":  sum_col("CCOS_BAL_DUE", osbal_rows),
    }

def _clean_base_match(scenario: str) -> str:
    s = str(scenario or "").strip()
    for to_remove in [" (Beda Currency)", " (Beda Periode)", " (> 1 Fac)"]:
        s = s.replace(to_remove, "")
    s = s.strip()
    if " only + " in s:
        s = s.replace(" only + ", " + ")
    return s

def _format_final_scenario(base_scenario: str, source: str, is_beda_curr: bool, is_beda_per: bool, num_fac: int) -> str:
    if not source or not base_scenario or str(base_scenario).strip().upper() in ("UNMATCHING", "UNMATCHED"):
        return "Unmatching"
    base = _clean_base_match(base_scenario)
    if not base or base.upper() in ("UNMATCHING", "UNMATCHED"):
        return "Unmatching"
    if is_beda_curr:
        return f"{base} (Beda Currency)"
    elif is_beda_per:
        return f"{base} (Beda Periode)"
    elif num_fac > 1:
        return f"{base} (> 1 Fac)"
    else:
        return base

def _build_output_row(
    suspend_row: dict, source: str, scenario: str, osbal_rows: list,
    osbal_count: int = 0
) -> dict:
    has_osbal = bool(osbal_rows and osbal_rows[0])

    if has_osbal:
        ccos = _aggregate_ccos(osbal_rows)
    else:
        ccos = {"CCOS_DOC_NO": "", "CCOS_REF_CODE": "", "CCOS_OR_BAL": np.nan, "CCOS_BAL_DUE": np.nan}

    if "Beda Currency" in str(scenario) or "Beda Periode" in str(scenario):
        ccos["CCOS_OR_BAL"] = np.nan
        ccos["CCOS_BAL_DUE"] = np.nan

    final_scenario = "Unmatching" if not source or scenario in ("Unmatching", "UNMATCHED") else scenario

    return {
        "CCOS_DOC_NO":            ccos["CCOS_DOC_NO"],
        "CCOS_REF_CODE":          ccos["CCOS_REF_CODE"],
        "RECEIPT NO":             suspend_row.get("RECEIPT NO",        ""),
        "CREDIT NOTES":           suspend_row.get("CREDIT NOTES",      ""),
        "DETAIL RINCIAN NO":      suspend_row.get("DETAIL RINCIAN NO", ""),
        "RECEIPT DATE":           suspend_row.get("RECEIPT DATE",      ""),
        "CEDANT NAME":            suspend_row.get("CEDANT NAME",       ""),
        "CEDANT SHRT NAME":       suspend_row.get("CEDANT SHRT NAME",  ""),
        "INSURED_ORI":            suspend_row.get("INSURED_ORI") or suspend_row.get("insured_ori") or "",
        "INSURED_1":              suspend_row.get("INSURED_1") or suspend_row.get("clean insured 1") or "",
        "INSURED_2":              suspend_row.get("INSURED_2") or suspend_row.get("clean insured 2") or "",
        "CURR ORI":               suspend_row.get("CURR ORI",          ""),
        "AMOUNT ORI":             suspend_row.get("AMOUNT ORI",        ""),
        "CURR PAY":               suspend_row.get("CURR PAY",          ""),
        "AMOUNT PAY":             suspend_row.get("AMOUNT PAY",        ""),
        "CCOS_OR_BAL":            ccos["CCOS_OR_BAL"],
        "CCOS_BAL_DUE":           ccos["CCOS_BAL_DUE"],
        "POLIS":                  suspend_row.get("POLIS_ORI") or suspend_row.get("POLIS") or "",
        "POLICY_CLEAN_1":         suspend_row.get("POLIS_CLEAN_1") or suspend_row.get("POLICY_CLEAN_1") or "",
        "SERTIF_CLEAN_1":         "",
        "POLICY_CLEAN_2":         suspend_row.get("POLIS_CLEAN_2") or suspend_row.get("POLICY_CLEAN_2") or "",
        "SERTIF_CLEAN_2":         "",
        "POLICY_CLEAN_3":         suspend_row.get("POLIS_CLEAN_3") or suspend_row.get("POLICY_CLEAN_3") or "",
        "SERTIF_CLEAN_3":         "",
        "POLICY_CLEAN_4":         suspend_row.get("POLIS_CLEAN_4") or suspend_row.get("POLICY_CLEAN_4") or "",
        "POLICY_CLEAN_5":         suspend_row.get("POLIS_CLEAN_5") or suspend_row.get("POLICY_CLEAN_5") or "",
        "SLIP_NO":                suspend_row.get("SLIP_NO_ORI") or suspend_row.get("SLIP_NO") or "",
        "SLIP_NO_CLN":            suspend_row.get("SLIP_CLEAN_1") or suspend_row.get("SLIP_NO_CLN") or "",
        "DESC 1":                 suspend_row.get("DESC 1",            ""),
        "DESC 2":                 suspend_row.get("DESC 2",            ""),
        "DESC 3":                 suspend_row.get("DESC 3",            ""),
        "DESC 4":                 suspend_row.get("DESC 4",            ""),
        "STATUS":                 suspend_row.get("STATUS",            ""),
        "REC_TYPE":               suspend_row.get("REC_TYPE",          ""),
        "SKENARIO":               final_scenario,
        "_OSBAL_ROW_COUNT":       osbal_count,
        "_SUSPEND_ROW_COUNT":     1,
    }

def _compute_derived_cols(df: pd.DataFrame) -> pd.DataFrame:
    amount_ori_base = df.get("AMOUNT ORI", pd.Series([0]*len(df), index=df.index))
    amount_ori_clean = amount_ori_base.replace(r'^\s*$', np.nan, regex=True)
    if amount_ori_clean.dtype == object:
        amount_ori_clean = amount_ori_clean.astype(str).str.strip().str.replace(r'\s+', '', regex=True)
        amount_ori_clean = amount_ori_clean.replace("nan", np.nan)
        has_comma = amount_ori_clean.str.contains(',', na=False)
        amount_ori_eu = (amount_ori_clean
                         .str.replace('.', '', regex=False)
                         .str.replace(',', '.', regex=False))
        amount_ori_clean = amount_ori_eu.where(has_comma, amount_ori_clean)
    amount_ori_numeric = pd.to_numeric(amount_ori_clean, errors="coerce")

    df["AMOUNT_ORI_MIN1"] = np.where(amount_ori_numeric.notna(), amount_ori_numeric * -1, np.nan)
    amount_ori = amount_ori_numeric.fillna(0)
    amount_neg = df["AMOUNT_ORI_MIN1"].fillna(0)

    balance_due     = pd.to_numeric(df["CCOS_BAL_DUE"] if "CCOS_BAL_DUE" in df.columns else pd.Series(0, index=df.index), errors="coerce").fillna(0)
    osbal_count_col = pd.to_numeric(df["_OSBAL_ROW_COUNT"] if "_OSBAL_ROW_COUNT" in df.columns else pd.Series(0, index=df.index), errors="coerce").fillna(0)
    ccos_ref        = df.get("CCOS_REF_CODE", pd.Series("", index=df.index)).fillna("").astype(str)
    skenario_col    = df.get("SKENARIO",      pd.Series("", index=df.index)).fillna("")

    df["DIFERENCE"] = amount_neg - balance_due
    accumulated = (osbal_count_col > 1)
    is_equal    = amount_neg.round(2) == balance_due.round(2)

    is_unmatched         = skenario_col.str.strip().str.upper().isin(["UNMATCHING", "UNMATCHED"]) | (skenario_col == "")
    is_beda_currency     = skenario_col.str.contains("Beda Currency",      na=False, regex=False)
    is_beda_periode      = skenario_col.str.contains("Beda Periode",       na=False, regex=False)
    is_gt1_fac           = (skenario_col.str.contains("> 1 Fac", na=False, regex=False) | ccos_ref.str.contains(",", na=False))
    is_new_entry_total   = ~is_unmatched & (amount_ori != 0) & (balance_due == 0)

    conditions = [
        is_unmatched,
        is_beda_currency,
        is_beda_periode,
        is_gt1_fac,
        is_new_entry_total,
        is_equal & ~accumulated,
        is_equal &  accumulated,
        (amount_neg < balance_due) & ~accumulated,
        (amount_neg < balance_due) &  accumulated,
        (amount_neg > balance_due),
    ]
    choices = [
        "Unmatching",
        "Beda Currency",
        "Beda Periode",
        "Matching >1 fac code",
        "New Entry",
        "Adjustment total tanpa akumulasi",
        "Adjustment total dengan akumulasi",
        "Adjustment sebagian tanpa akumulasi",
        "Adjustment sebagian dengan akumulasi",
        "New Entry",
    ]
    df["FLAG_PROD"] = np.select(conditions, choices, default="Unmatching")
    return df

def _load_excel(path: str, label: str, optional: bool = False) -> tuple:
    if not os.path.exists(path):
        return [], []
    cache = path + ".cache.pkl"
    if os.path.exists(cache) and os.path.getmtime(cache) >= os.path.getmtime(path):
        try:
            data, header = pd.read_pickle(cache)
            return data, header
        except Exception:
            pass
    df = pd.read_excel(path)
    header = list(df.columns)
    data   = df.to_dict("records")
    return data, header

# =============================================================================
# RUN DEBUG PIPELINE
# =============================================================================

def run_debug() -> None:
    t_start = time.perf_counter()
    print("\n" + "=" * 65, flush=True)
    print("  PRODUCTION DEBUGGING SCRIPT — SUSPEND MATCHING (KBRU) — V1", flush=True)
    print("  Urutan Matching: SLIP -> POLIS -> INSURED", flush=True)
    print("  Implementasi Solusi: Sanitasi Dummy, Smart FACUL Fallback, Polis Clean", flush=True)
    print("=" * 65, flush=True)

    suspend_rows, suspend_cols = _load_excel(SUSPEND_FILE, "SUSPEND")
    osbal_rows,   osbal_cols   = _load_excel(OSBAL_FILE,   "OSBAL")
    facul_rows,   facul_cols   = _load_excel(FACUL_FILE,   "FACUL")
    slipdb_rows,  slipdb_cols  = _load_excel(SLIPDB_FILE,  "SLIPDB", optional=True)

    for r in osbal_rows:
        raw = r.get(OSBAL_DATE_COL, "")
        r["_com_date_parsed"] = pd.to_datetime(raw, errors="coerce") if raw and not (isinstance(raw, float) and pd.isna(raw)) else None

    for r in suspend_rows:
        raw = r.get(SUSPEND_DATE_COL, "")
        r["_sus_date_parsed"] = pd.to_datetime(raw, errors="coerce") if raw and not (isinstance(raw, float) and pd.isna(raw)) else None

    # Deteksi kolom
    polis_sus   = _get_clean_cols(suspend_cols, "clean polis")
    slip_sus    = _get_clean_cols(suspend_cols, "clean slip")
    insured_sus = _get_clean_cols(suspend_cols, "clean insured")

    polis_osbal   = _get_clean_cols(osbal_cols, "clean polis")
    slip_osbal    = _get_clean_cols(osbal_cols, "clean slip")
    insured_osbal = _get_clean_cols(osbal_cols, "clean insured")
    if "CLSDT_POLICY_NO" in osbal_cols and "CLSDT_POLICY_NO" not in polis_osbal:
        polis_osbal = ["CLSDT_POLICY_NO"] + polis_osbal
    if "CLSDT_SLIP_NO" in osbal_cols and "CLSDT_SLIP_NO" not in slip_osbal:
        slip_osbal = ["CLSDT_SLIP_NO"] + slip_osbal

    polis_facul   = _get_clean_cols(facul_cols, "clean polis")
    slip_facul    = _get_clean_cols(facul_cols, "clean slip")
    insured_facul = _get_clean_cols(facul_cols, "clean insured")

    polis_slipdb   = _get_clean_cols(slipdb_cols, "clean polis")
    slip_slipdb    = _get_clean_cols(slipdb_cols, "clean slip")
    insured_slipdb = _get_clean_cols(slipdb_cols, "clean insured")

    # Build lookup index
    print("\n[Membuat Lookup Index...]", flush=True)
    lookup_osbal  = _build_lookup(osbal_rows,  polis_osbal,  slip_osbal,  insured_osbal,  OSBAL_FACODE_COL,  label="OSBAL")
    lookup_facul  = _build_lookup(facul_rows,  polis_facul,  slip_facul,  insured_facul,  FACUL_FACODE_COL,  label="FACUL")
    lookup_slipdb = _build_lookup(slipdb_rows, polis_slipdb, slip_slipdb, insured_slipdb, SLIPDB_FACODE_COL, label="SLIPDB")

    facode_osbal_idx = lookup_osbal[6]
    total = len(suspend_rows)
    print(f"\n[Mencocokkan {total:,} baris suspend...]", flush=True)

    raw_results = []
    for n, sus in enumerate(suspend_rows, 1):
        if n % 500 == 0 or n == total:
            print(f"  Baris {n:,} / {total:,} ...", flush=True)

        if str(sus.get("STATUS", "")).strip().upper() == "MATCHING":
            continue

        effective_date = sus.get("_sus_date_parsed")
        sus_curr = _normalize(sus.get(SUSPEND_CURR_COL, ""))

        matched_osbal_idx = set()
        source = None
        scenario = ""

        # -------------------------------------------------------------
        # FASE 1: SLIP (Prioritas 1)
        # -------------------------------------------------------------
        slip_vals = _collect_clean_values(sus, slip_sus, is_slip=True, is_polis=False)
        if slip_vals:
            osb_hits = _exact_match(slip_vals, lookup_osbal[0])
            if osb_hits:
                matched_osbal_idx.update(osb_hits)
                source = "OSBAL"
                scenario = "Slip only"
            elif slipdb_rows:
                sld_hits = _exact_match(slip_vals, lookup_slipdb[0])
                if sld_hits:
                    res_idx, ok = _resolve_facode([slipdb_rows[i] for i in sld_hits], SLIPDB_FACODE_COL, facode_osbal_idx, osbal_rows)
                    if ok:
                        matched_osbal_idx.update(res_idx)
                        source = "SLIPDB"
                        scenario = "Slip only"
            elif facul_rows:
                fac_hits = _exact_match(slip_vals, lookup_facul[0])
                if fac_hits:
                    res_idx, ok = _resolve_facode([facul_rows[i] for i in fac_hits], FACUL_FACODE_COL, facode_osbal_idx, osbal_rows)
                    if ok:
                        matched_osbal_idx.update(res_idx)
                        source = "FACUL"
                        scenario = "Slip only"

            # LIKE Slip di OSBAL
            if not matched_osbal_idx and lookup_osbal[3] and lookup_osbal[3][0]:
                like_hits = _like_match(slip_vals, lookup_osbal[3][0], osbal_rows, slip_osbal)
                if like_hits:
                    matched_osbal_idx.update(like_hits)
                    source = "OSBAL"
                    scenario = "Slip only"

        # -------------------------------------------------------------
        # FASE 2: POLIS (Prioritas 2)
        # -------------------------------------------------------------
        if not matched_osbal_idx:
            polis_vals = _collect_clean_values(sus, polis_sus, is_slip=False, is_polis=True)
            if polis_vals:
                osb_hits = _exact_match(polis_vals, lookup_osbal[1])
                if osb_hits:
                    matched_osbal_idx.update(osb_hits)
                    source = "OSBAL"
                    scenario = "Polis only"
                elif slipdb_rows:
                    sld_hits = _exact_match(polis_vals, lookup_slipdb[1])
                    if sld_hits:
                        res_idx, ok = _resolve_facode([slipdb_rows[i] for i in sld_hits], SLIPDB_FACODE_COL, facode_osbal_idx, osbal_rows)
                        if ok:
                            matched_osbal_idx.update(res_idx)
                            source = "SLIPDB"
                            scenario = "Polis only"
                elif facul_rows:
                    fac_hits = _exact_match(polis_vals, lookup_facul[1])
                    if fac_hits:
                        res_idx, ok = _resolve_facode([facul_rows[i] for i in fac_hits], FACUL_FACODE_COL, facode_osbal_idx, osbal_rows)
                        if ok:
                            matched_osbal_idx.update(res_idx)
                            source = "FACUL"
                            scenario = "Polis only"

                # LIKE Polis di OSBAL
                if not matched_osbal_idx and lookup_osbal[4] and lookup_osbal[4][0]:
                    like_hits = _like_match(polis_vals, lookup_osbal[4][0], osbal_rows, polis_osbal)
                    if like_hits:
                        matched_osbal_idx.update(like_hits)
                        source = "OSBAL"
                        scenario = "Polis only"

        # -------------------------------------------------------------
        # FASE 3: INSURED (Prioritas 3)
        # -------------------------------------------------------------
        if not matched_osbal_idx:
            ins_vals = _collect_clean_values(sus, insured_sus, is_slip=False, is_polis=False)
            if ins_vals:
                osb_hits = _exact_match(ins_vals, lookup_osbal[2])
                if osb_hits:
                    matched_osbal_idx.update(osb_hits)
                    source = "OSBAL"
                    scenario = "Insured only"
                elif slipdb_rows:
                    sld_hits = _exact_match(ins_vals, lookup_slipdb[2])
                    if sld_hits:
                        res_idx, ok = _resolve_facode([slipdb_rows[i] for i in sld_hits], SLIPDB_FACODE_COL, facode_osbal_idx, osbal_rows)
                        if ok:
                            matched_osbal_idx.update(res_idx)
                            source = "SLIPDB"
                            scenario = "Insured only"
                elif facul_rows:
                    fac_hits = _exact_match(ins_vals, lookup_facul[2])
                    if fac_hits:
                        res_idx, ok = _resolve_facode([facul_rows[i] for i in fac_hits], FACUL_FACODE_COL, facode_osbal_idx, osbal_rows)
                        if ok:
                            matched_osbal_idx.update(res_idx)
                            source = "FACUL"
                            scenario = "Insured only"

                # LIKE Insured di OSBAL
                if not matched_osbal_idx and lookup_osbal[5] and lookup_osbal[5][0]:
                    like_hits = _like_match(ins_vals, lookup_osbal[5][0], osbal_rows, insured_osbal)
                    if like_hits:
                        matched_osbal_idx.update(like_hits)
                        source = "OSBAL"
                        scenario = "Insured only"

        # -------------------------------------------------------------
        # FASE 4: NARROWING & FILTER
        # -------------------------------------------------------------
        is_beda_curr = False
        is_beda_per = False

        if matched_osbal_idx:
            matched_list = list(matched_osbal_idx)

            if sus_curr:
                filtered_curr = [i for i in matched_list if _normalize(osbal_rows[i].get(OSBAL_CURR_COL, "")) == sus_curr]
                if not filtered_curr:
                    is_beda_curr = True
                elif len(filtered_curr) < len(matched_list):
                    matched_list = filtered_curr

            if not is_beda_curr and pd.notna(effective_date):
                filtered_per = []
                for i in matched_list:
                    com_date = osbal_rows[i].get("_com_date_parsed")
                    if pd.isna(com_date) or effective_date >= com_date:
                        filtered_per.append(i)
                if not filtered_per:
                    is_beda_per = True
                elif len(filtered_per) < len(matched_list):
                    matched_list = filtered_per

            # Narrowing hierarchy: Slip -> Polis -> Insured
            if scenario.startswith("Slip") and len(matched_list) > 1:
                p_vals = _collect_clean_values(sus, polis_sus, is_slip=False, is_polis=True)
                if p_vals:
                    sub = [i for i in matched_list if any(_normalize(osbal_rows[i].get(c, "")) in p_vals for c in polis_osbal)]
                    if sub:
                        matched_list = sub
                        scenario = "Slip + Polis"
                if len(matched_list) > 1:
                    i_vals = _collect_clean_values(sus, insured_sus, is_slip=False, is_polis=False)
                    if i_vals:
                        sub = [i for i in matched_list if any(_normalize(osbal_rows[i].get(c, "")) in i_vals for c in insured_osbal)]
                        if sub:
                            matched_list = sub
                            scenario = scenario + " + Insured" if "Polis" in scenario else "Slip + Insured"

            elif scenario.startswith("Polis") and len(matched_list) > 1:
                s_vals = _collect_clean_values(sus, slip_sus, is_slip=True, is_polis=False)
                if s_vals:
                    sub = [i for i in matched_list if any(_normalize(osbal_rows[i].get(c, "")) in s_vals for c in slip_osbal)]
                    if sub:
                        matched_list = sub
                        scenario = "Polis + Slip"
                if len(matched_list) > 1:
                    i_vals = _collect_clean_values(sus, insured_sus, is_slip=False, is_polis=False)
                    if i_vals:
                        sub = [i for i in matched_list if any(_normalize(osbal_rows[i].get(c, "")) in i_vals for c in insured_osbal)]
                        if sub:
                            matched_list = sub
                            scenario = scenario + " + Insured" if "Slip" in scenario else "Polis + Insured"

            elif scenario.startswith("Insured") and len(matched_list) > 1:
                s_vals = _collect_clean_values(sus, slip_sus, is_slip=True, is_polis=False)
                if s_vals:
                    sub = [i for i in matched_list if any(_normalize(osbal_rows[i].get(c, "")) in s_vals for c in slip_osbal)]
                    if sub:
                        matched_list = sub
                        scenario = "Slip + Insured"
                if len(matched_list) > 1:
                    p_vals = _collect_clean_values(sus, polis_sus, is_slip=False, is_polis=True)
                    if p_vals:
                        sub = [i for i in matched_list if any(_normalize(osbal_rows[i].get(c, "")) in p_vals for c in polis_osbal)]
                        if sub:
                            matched_list = sub
                            scenario = "Polis + Insured"

            final_osbal_rows = [osbal_rows[i] for i in matched_list]
            fac_set = {str(r.get(OSBAL_FACODE_COL, "")).strip() for r in final_osbal_rows if str(r.get(OSBAL_FACODE_COL, "")).strip()}
            final_scen = _format_final_scenario(scenario, source, is_beda_curr, is_beda_per, len(fac_set))
            out_row = _build_output_row(sus, source, final_scen, final_osbal_rows, osbal_count=len(final_osbal_rows))

        else:
            out_row = _build_output_row(sus, None, "Unmatching", osbal_rows=[], osbal_count=0)

        raw_results.append(out_row)

    df_out = pd.DataFrame(raw_results)
    df_out = _compute_derived_cols(df_out)

    # Simpan ke Excel Debug
    print(f"\n[Menyimpan file debug ke {OUTPUT_DEBUG}...]", flush=True)
    df_out.to_excel(OUTPUT_DEBUG, index=False)
    
    # Beri warna ungu muda pada kolom script via excel_styler
    try:
        sys.path.append(ROOT_DIR)
        from excel_styler import apply_purple_column_style
        apply_purple_column_style(OUTPUT_DEBUG, "final_output")
    except Exception as e:
        print(f"  [Notice] Styler skipped: {e}")

    elapsed = time.perf_counter() - t_start

    # [KOMPARASI DENGAN V1 ASLI]
    v1_asli_path = os.path.join(DATA_DIR, "final_output_v1.xlsx")
    df_v1_asli = pd.read_excel(v1_asli_path) if os.path.exists(v1_asli_path) else None

    print("\n" + "=" * 65)
    print("  HASIL EKSEKUSI SCRIPT DEBUGGING V1 SELESAI")
    print(f"  Waktu: {elapsed:.1f} detik | Total Baris: {len(df_out):,}")
    print("=" * 65)

    print("\nTABEL KOMPARASI: V1 ASLI vs V1 DEBUG:")
    print("-" * 65)
    print(f"{'FLAG_PROD':<38} | {'V1 ASLI':<10} | {'V1 DEBUG':<10}")
    print("-" * 65)
    all_flags = sorted(list(set(df_out["FLAG_PROD"].unique()) | (set(df_v1_asli["FLAG_PROD"].unique()) if df_v1_asli is not None else set())))
    for flg in all_flags:
        c_asli = (df_v1_asli["FLAG_PROD"] == flg).sum() if df_v1_asli is not None else 0
        c_debug = (df_out["FLAG_PROD"] == flg).sum()
        diff = c_debug - c_asli
        diff_str = f"({diff:+d})" if diff != 0 else ""
        print(f"{flg:<38} | {c_asli:>10,} | {c_debug:>10,} {diff_str}")
    print("-" * 65)

    print("\nTABEL KOMPARASI SKENARIO:")
    print("-" * 65)
    print(f"{'SKENARIO':<38} | {'V1 ASLI':<10} | {'V1 DEBUG':<10}")
    print("-" * 65)
    all_scen = sorted(list(set(df_out["SKENARIO"].unique()) | (set(df_v1_asli["SKENARIO"].unique()) if df_v1_asli is not None else set())))
    for scn in all_scen:
        c_asli = (df_v1_asli["SKENARIO"] == scn).sum() if df_v1_asli is not None else 0
        c_debug = (df_out["SKENARIO"] == scn).sum()
        diff = c_debug - c_asli
        diff_str = f"({diff:+d})" if diff != 0 else ""
        print(f"{scn:<38} | {c_asli:>10,} | {c_debug:>10,} {diff_str}")
    print("-" * 65)

if __name__ == "__main__":
    run_debug()
