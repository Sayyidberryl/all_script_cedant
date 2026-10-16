"""
prod_sus_2.py — Suspend Matching Script untuk Cedant ASKRINDO — V2 (Akumulasi / Merged Rows)

Aturan Matching:
  - Urutan matching: POLIS -> SLIP -> INSURED (Exact -> LIKE)
  - Narrowing wajib: Currency (CCOS_CURR == CURR ORI) & Periode (RECEIPT DATE >= FAC_COM_DATE)
  - Narrowing basic:
      * Jika POLIS   -> + SLIP + INSURED
      * Jika SLIP    -> + POLIS + INSURED
      * Jika INSURED -> + POLIS + SLIP
  - Mode V2: Suspend rows yang share 1 FAC CODE di-merge / diakumulasi menjadi 1 baris
  - TIDAK menerapkan Periode 2 (penagihan + 1 bulan)
  - TIDAK menerapkan Narrowing Bordero (Open Cover)
  - Non-destruktif: tetap pertahankan kandidat dengan flag kondisi khusus jika semua berbeda
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
from sqlalchemy import create_engine, text


# =============================================================================
# KONFIGURASI
# =============================================================================

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR   = os.path.join(SCRIPT_DIR, "data")

def _find_file(candidates: list) -> str:
    """Kembalikan path pertama yang ada dari daftar kandidat."""
    for p in candidates:
        if os.path.exists(p):
            return p
    return candidates[0]


SUSPEND_FILE = _find_file([
    os.path.join(DATA_DIR, "suspend_clean_askrindo.xlsx"),
    os.path.join(DATA_DIR, "askrindo_output_suspend.xlsx"),
    os.path.join(DATA_DIR, "suspend_clean_aca.xlsx"),
])
OSBAL_FILE = _find_file([
    os.path.join(DATA_DIR, "osbal_clean_askrindo.xlsx"),
    os.path.join(DATA_DIR, "askrindo_output_osbal.xlsx"),
    os.path.join(DATA_DIR, "osbal_clean_aca.xlsx"),
])
FACUL_FILE = _find_file([
    os.path.join(DATA_DIR, "facul_clean_askrindo.xlsx"),
    os.path.join(DATA_DIR, "askrindo_output_facul.xlsx"),
    os.path.join(DATA_DIR, "facul_clean_aca.xlsx"),
])
OUTPUT_FILE  = os.path.join(DATA_DIR, "final_output_v2.xlsx")

# Kolom FAC CODE
OSBAL_FACODE_COL = "CCOS_REF_CODE"
FACUL_FACODE_COL = "FAC_CODE"

# Kolom currency & periode
SUSPEND_CURR_COL = "CURR ORI"
OSBAL_CURR_COL   = "CCOS_CURR"
SUSPEND_DATE_COL = "RECEIPT DATE"
OSBAL_DATE_COL   = "FAC_COM_DATE"

# Marker baris administratif yang dikecualikan
_EXCLUDED_REF_MARKERS = ["HUTANG PIUTANG", "DATA SUSPENSE"]

# Marker LINESLIP
_LINESLIP_MARKERS = ("LINESLIP", "LINE SLIP")

# Token matching config
_MIN_TOKEN_LEN = 3
_TOKEN_RE      = re.compile(r"[^A-Z0-9]+")
_DELIMITER_RE  = re.compile(r"[,;|+\s]+")

# Kolom suspend yang di-join saat akumulasi
_SUSPEND_JOIN_COLS = [
    "RECEIPT NO", "CREDIT NOTES", "DETAIL RINCIAN NO", "RECEIPT DATE",
    "CEDANT NAME", "CEDANT SHRT NAME",
    "INSURED", "INSURED_CLEAN_1", "INSURED_CLEAN_2",
    "CURR ORI", "CURR PAY", "AMOUNT PAY",
    "POLIS", "POLIS_CLEAN",
    "SLIP NO", "SLIP_NO_CLEAN",
    "DESC 1", "DESC 2", "DESC 3", "DESC 4",
    "STATUS", "REC_TYPE",
]
_SUSPEND_SUM_COLS = ["AMOUNT ORI"]

# Kolom output final
FINAL_COLUMNS = [
    "CCOS_DOC_NO", "CCOS_REF_CODE",
    "RECEIPT NO", "CREDIT NOTES", "DETAIL RINCIAN NO", "RECEIPT DATE",
    "CEDANT NAME", "CEDANT SHRT NAME",
    "INSURED_ORI", "INSURED_1", "INSURED_2",
    "CURR ORI", "AMOUNT ORI", "AMOUNT_ORI_MIN1",
    "CURR PAY", "AMOUNT PAY",
    "CCOS_OR_BAL", "CCOS_BAL_DUE", "DIFERENCE",
    "FLAG_PROD",
    "POLIS", "POLIS_CLN",
    "SLIP_NO", "SLIP_NO_CLN",
    "DESC 1", "DESC 2", "DESC 3", "DESC 4",
    "STATUS", "REC_TYPE",
    "SKENARIO",
]


# =============================================================================
# STRING UTILITIES
# =============================================================================

@functools.lru_cache(maxsize=131072)
def _norm_cached(s: str) -> str:
    """Normalisasi string — cached karena dipanggil ribuan kali."""
    text = s.strip().upper().replace("S/D", "SD")
    return re.sub(r"\s+", " ", text) if "  " in text else text


def _normalize(value) -> str:
    """Uppercase + strip + collapse spasi, return '' untuk null."""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    return _norm_cached(str(value))


@functools.lru_cache(maxsize=100000)
def _split_composite(text: str) -> frozenset:
    """Pecah string berdelimiter (koma/plus/spasi/titik koma/pipe) jadi token."""
    if not text:
        return frozenset()
    return frozenset(tok.strip() for tok in _DELIMITER_RE.split(text) if tok.strip())


def _tokenize(text: str) -> frozenset:
    """Pecah string jadi frozenset token alfanumerik (min 3 karakter)."""
    if not text:
        return frozenset()
    return frozenset(t for t in _TOKEN_RE.split(text) if len(t) >= _MIN_TOKEN_LEN)


def _get_clean_cols(columns: list, prefix: str) -> list:
    """Ambil nama kolom yang berawalan prefix (case-insensitive)."""
    if not columns:
        return []
    alt_map = {
        "clean polis": [
            "clean polis", "polis_clean", "policy_clean", "polis_cln",
        ],
        "clean slip": [
            "clean slip", "slip_no_clean", "slip_clean", "slip_no_cln",
            "clsdt_slip",
        ],
        "clean insured": [
            "clean insured", "insured_clean", "insured_",
        ],
        "clean sertif": [
            "clean sertif", "sertif_clean", "certificate_", "sertif_cln",
        ],
    }
    prefixes = alt_map.get(prefix.lower(), [prefix.lower()])
    result = []
    for col in columns:
        col_lower = str(col).lower()
        if any(col_lower.startswith(p) for p in prefixes):
            result.append(col)
    return result


# =============================================================================
# ROW DICT HELPERS
# =============================================================================

def _is_excluded_ref_row(row: dict, polis_cols: list, slip_cols: list) -> bool:
    """True jika baris referensi adalah baris administratif."""
    for col in list(polis_cols) + list(slip_cols):
        val = _normalize(row.get(col, ""))
        if val and any(m in val for m in _EXCLUDED_REF_MARKERS):
            return True
    return False


def _is_lineslip_row(sus_row: dict, insured_cols: list) -> bool:
    """True jika baris suspend mengandung marker LINESLIP di kolom insured."""
    for col in list(insured_cols) + ["INSURED", "insured_ori"]:
        val = _normalize(sus_row.get(col, ""))
        if val and any(m in val for m in _LINESLIP_MARKERS):
            return True
    return False


def _collect_clean_values(row: dict, cols: list) -> list:
    """Kumpulkan nilai bersih unik (non-empty) dari kolom yang diberikan."""
    seen, result = set(), []
    for col in cols:
        val = _normalize(row.get(col, ""))
        if val and val not in seen:
            seen.add(val)
            result.append(val)
    return result


def _ref_has_value(ref_row: dict, clean_cols: list, query_values: list) -> bool:
    """True jika ref_row memiliki salah satu query_values di kolom clean_cols."""
    if not query_values:
        return False
    ref_vals = set(_collect_clean_values(ref_row, clean_cols))
    if not ref_vals:
        return False
    # Exact match
    if set(query_values) & ref_vals:
        return True
    # Composite token split match
    for rv in ref_vals:
        if any(v in _split_composite(rv) for v in query_values if v and len(v) >= 2):
            return True
    # Substring match (min 6 char, dua arah)
    for v in query_values:
        if not v or len(v) < 6:
            continue
        for rv in ref_vals:
            if not rv:
                continue
            if v in rv or (len(rv) >= 6 and rv in v):
                return True
    return False


# =============================================================================
# INDEX BUILDERS
# =============================================================================

def _build_exact_index(rows: list, cols: list, excluded: set = None) -> dict:
    """Bangun inverted index: value -> [indeks baris] untuk exact match."""
    index: dict = {}
    for i, row in enumerate(rows):
        if excluded and i in excluded:
            continue
        for col in cols:
            val = _normalize(row.get(col, ""))
            if val:
                index.setdefault(val, []).append(i)
    return index


def _build_facode_index(rows: list, col: str, excluded: set = None) -> dict:
    """Bangun inverted index: FAC_CODE -> [indeks baris]."""
    index: dict = {}
    for i, row in enumerate(rows):
        if excluded and i in excluded:
            continue
        val = _normalize(row.get(col, ""))
        if val:
            index.setdefault(val, []).append(i)
    return index


def _build_token_index(rows: list, cols: list, excluded: set = None) -> tuple:
    """Bangun inverted index token alfanumerik untuk LIKE match."""
    token_index: dict = {}
    row_tokens: list  = []
    for i, row in enumerate(rows):
        if excluded and i in excluded:
            row_tokens.append(frozenset())
            continue
        tokens: frozenset = frozenset()
        for col in cols:
            val = _normalize(row.get(col, ""))
            if val:
                tokens = tokens | _tokenize(val)
        row_tokens.append(tokens)
        for tok in tokens:
            token_index.setdefault(tok, set()).add(i)
    return token_index, row_tokens


def _build_lookup(
    rows: list,
    polis_cols: list,
    slip_cols: list,
    insured_cols: list,
    facode_col: str,
    label: str = "REF",
) -> dict:
    """Bangun kumpulan index untuk tabel referensi."""
    if not rows:
        return {}
    t0 = time.perf_counter()
    excluded = {
        i for i, r in enumerate(rows)
        if _is_excluded_ref_row(r, polis_cols, slip_cols)
    }
    lkp = {
        "exact_polis":   _build_exact_index(rows, polis_cols, excluded),
        "exact_slip":    _build_exact_index(rows, slip_cols, excluded),
        "exact_insured": _build_exact_index(rows, insured_cols, excluded),
        "token_polis":   _build_token_index(rows, polis_cols, excluded),
        "token_slip":    _build_token_index(rows, slip_cols, excluded),
        "token_insured": _build_token_index(rows, insured_cols, excluded),
        "facode":        _build_facode_index(rows, facode_col, excluded),
        "excluded":      excluded,
    }
    print(f"  {label:<6} done ({time.perf_counter()-t0:.1f}s) - {len(excluded)} dikecualikan", flush=True)
    return lkp


# =============================================================================
# MATCH ENGINE
# =============================================================================

def _exact_match(query_values: list, index: dict) -> set:
    """Kembalikan indeks baris yang cocok exact dengan salah satu nilai query."""
    return {i for v in query_values if v and v in index for i in index[v]}


def _like_match(
    query_values: list,
    token_index: dict,
    rows: list,
    ref_cols: list,
    min_len: int = 5,
) -> list:
    """Kembalikan indeks baris di mana nilai query muncul sebagai substring di ref (LIKE match)."""
    valid_queries = [v for v in query_values if v and len(v) >= min_len]
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
            if any((q in rv) or (len(rv) >= min_len and rv in q) for rv in ref_values if rv):
                matched.append(i)
                break
    return matched


# =============================================================================
# NARROWING
# =============================================================================

def _narrow_by_currency(
    suspend_row: dict, matched: list, rows: list,
) -> tuple:
    """Saring kandidat: CCOS_CURR == CURR ORI."""
    if len(matched) <= 1:
        return matched, False
    sus_curr = _normalize(suspend_row.get(SUSPEND_CURR_COL, ""))
    if not sus_curr:
        return matched, False
    filtered = [i for i in matched if _normalize(rows[i].get(OSBAL_CURR_COL, "")) == sus_curr]
    if not filtered:
        return matched, True
    return filtered, False


def _narrow_by_periode(
    suspend_row: dict, matched: list, rows: list,
    effective_sus_date=None, is_lineslip: bool = False,
) -> tuple:
    """Saring kandidat: RECEIPT DATE >= FAC_COM_DATE."""
    if len(matched) <= 1:
        return matched, False

    sus_date = effective_sus_date or suspend_row.get("_sus_date_parsed")
    if sus_date is None or pd.isna(sus_date):
        return matched, False

    valid = []
    for i in matched:
        com_date = rows[i].get("_com_date_parsed")
        if com_date is None or pd.isna(com_date):
            valid.append(i)
            continue
        try:
            if is_lineslip:
                if com_date.year == sus_date.year and com_date.month == sus_date.month:
                    valid.append(i)
            else:
                if sus_date >= com_date:
                    valid.append(i)
        except Exception:
            valid.append(i)

    if not valid:
        return matched, True
    return valid, False


def _narrow_basic(
    suspend_row: dict,
    matched: list,
    rows: list,
    match_label: str,
    polis_ref_cols: list,
    slip_ref_cols: list,
    insured_ref_cols: list,
    polis_sus_cols: list,
    slip_sus_cols: list,
    insured_sus_cols: list,
) -> tuple:
    """Narrowing basic non-destruktif (cross-field confirmation).

    Aturan narrowing basic:
      - Match via POLIS   -> narrow dengan SLIP, lalu INSURED (Polis + Slip + Insured)
      - Match via SLIP    -> narrow dengan POLIS, lalu INSURED (Slip + Polis + Insured)
      - Match via INSURED -> narrow dengan POLIS, lalu SLIP (Insured + Polis + Slip)
    """
    if not matched:
        return [], "Unmatching"
    if len(matched) <= 1:
        return matched, _map_scenario(match_label)

    current_matched = list(matched)
    current_label   = _map_scenario(match_label)

    polis_vals   = _collect_clean_values(suspend_row, polis_sus_cols)
    polis_vals   = [p for p in polis_vals if len(p) >= 5 and p not in ["-", "TBA", "VARIOUS"]]
    slip_vals    = _collect_clean_values(suspend_row, slip_sus_cols)
    slip_vals    = [s for s in slip_vals if len(s) >= 4 and s not in ["-", "TBA"]]
    insured_vals = _collect_clean_values(suspend_row, insured_sus_cols)
    insured_vals = [i for i in insured_vals if len(i) >= 3 and i not in ["-", "TBA"]]

    # Ekstrak 10-digit base suffix untuk nomor slip bertipe Askrindo (misal 0561000032 dari 2201250561000032)
    slip_queries = list(slip_vals)
    for sv in slip_vals:
        if len(sv) >= 14 and sv[-10:].isdigit():
            base10 = sv[-10:]
            if base10 not in slip_queries:
                slip_queries.append(base10)

    # Karakteristik Askrindo: nomor referensi kontrak sering dicatat di kolom polis OSBAL, dan sebaliknya
    slip_check_cols = list(slip_ref_cols) + [c for c in polis_ref_cols if c not in slip_ref_cols]
    polis_check_cols = list(polis_ref_cols) + [c for c in slip_ref_cols if c not in polis_ref_cols]

    if "POLIS" in match_label.upper():
        base = "Polis"
        has_slip = False
        if slip_queries and len(current_matched) > 1:
            confirmed = [
                i for i in current_matched
                if _ref_has_value(rows[i], slip_check_cols, slip_queries)
            ]
            if confirmed:
                current_matched = confirmed
                has_slip = True

        has_ins = False
        if insured_vals and len(current_matched) > 1:
            confirmed = [
                i for i in current_matched
                if _ref_has_value(rows[i], insured_ref_cols, insured_vals)
            ]
            if confirmed:
                current_matched = confirmed
                has_ins = True

        if has_slip and has_ins:
            current_label = f"{base} + Slip + Insured"
        elif has_slip:
            current_label = f"{base} + Slip"
        elif has_ins:
            current_label = f"{base} + Insured"
        else:
            current_label = f"{base} only"

    elif "SLIP" in match_label.upper():
        base = "Slip"
        has_pol = False
        if polis_vals and len(current_matched) > 1:
            confirmed = [
                i for i in current_matched
                if _ref_has_value(rows[i], polis_check_cols, polis_vals)
            ]
            if confirmed:
                current_matched = confirmed
                has_pol = True

        has_ins = False
        if insured_vals and len(current_matched) > 1:
            confirmed = [
                i for i in current_matched
                if _ref_has_value(rows[i], insured_ref_cols, insured_vals)
            ]
            if confirmed:
                current_matched = confirmed
                has_ins = True

        if has_pol and has_ins:
            current_label = f"{base} + Polis + Insured"
        elif has_pol:
            current_label = f"{base} + Polis"
        elif has_ins:
            current_label = f"{base} + Insured"
        else:
            current_label = f"{base} only"

    elif "INSURED" in match_label.upper():
        base = "Insured"
        has_pol = False
        if polis_vals and len(current_matched) > 1:
            confirmed = [
                i for i in current_matched
                if _ref_has_value(rows[i], polis_check_cols, polis_vals)
            ]
            if confirmed:
                current_matched = confirmed
                has_pol = True

        has_slip = False
        if slip_queries and len(current_matched) > 1:
            confirmed = [
                i for i in current_matched
                if _ref_has_value(rows[i], slip_check_cols, slip_queries)
            ]
            if confirmed:
                current_matched = confirmed
                has_slip = True

        if has_pol and has_slip:
            current_label = f"{base} + Polis + Slip"
        elif has_pol:
            current_label = f"{base} + Polis"
        elif has_slip:
            current_label = f"{base} + Slip"
        else:
            current_label = f"{base} only"

    return current_matched, current_label


def _map_scenario(label: str) -> str:
    """Petakan label internal ke nama skenario resmi."""
    if not label or label in ("Unmatching", "UNMATCHED"):
        return "Unmatching"
    if "+" in label:
        return label
    if "POLIS" in label.upper():
        return "Polis only"
    if "SLIP" in label.upper():
        return "Slip only"
    if "INSURED" in label.upper():
        return "Insured only"
    return label


# =============================================================================
# FAC CODE RESOLUTION
# =============================================================================

def _resolve_facode(
    ref_rows_matched: list,
    facode_col: str,
    facode_osbal_idx: dict,
    osbal_rows: list,
) -> tuple:
    """Resolve FAC_CODE dari FACUL ke indeks baris OSBAL."""
    fac_codes = {_normalize(r.get(facode_col, "")) for r in ref_rows_matched}
    fac_codes.discard("")
    if not fac_codes:
        return [], False
    osbal_indices: set = set()
    for fac in fac_codes:
        if fac in facode_osbal_idx:
            osbal_indices.update(facode_osbal_idx[fac])
    result = list(osbal_indices)
    return result, bool(result)


# =============================================================================
# OUTPUT BUILDERS
# =============================================================================

def _aggregate_ccos(osbal_rows: list, all_fac_osbal_rows: list = None) -> dict:
    """Agregasi nilai finansial OSBAL."""
    fac_codes = {
        str(r.get("CCOS_REF_CODE", "")).strip()
        for r in osbal_rows
        if str(r.get("CCOS_REF_CODE", "")).strip()
    }
    multi_fac = len(fac_codes) > 1
    bal_due_rows = (
        all_fac_osbal_rows if (all_fac_osbal_rows is not None and not multi_fac) else osbal_rows
    )

    def join_unique(key):
        seen, vals = set(), []
        for r in osbal_rows:
            v = str(r.get(key, "")).strip()
            if v and v not in seen:
                seen.add(v)
                vals.append(v)
        return ", ".join(vals)

    def sum_col(key, source_rows):
        if multi_fac:
            return np.nan
        total = 0.0
        for r in source_rows:
            v = r.get(key, 0)
            if pd.isna(v):
                continue
            if isinstance(v, (int, float)):
                total += float(v)
            else:
                s = str(v).strip()
                if not s:
                    continue
                if "," in s and "." in s:
                    s = (
                        s.replace(".", "").replace(",", ".")
                        if s.rfind(",") > s.rfind(".")
                        else s.replace(",", "")
                    )
                elif "," in s:
                    s = s.replace(",", ".")
                try:
                    total += float(s)
                except ValueError:
                    pass
        return total

    return {
        "CCOS_DOC_NO":   join_unique("CCOS_DOC_NO"),
        "CCOS_REF_CODE": join_unique("CCOS_REF_CODE"),
        "CCOS_OR_BAL":   np.nan if multi_fac else join_unique("CCOS_OR_BAL"),
        "CCOS_BAL_DUE":  sum_col("CCOS_BAL_DUE", bal_due_rows),
    }


def _format_final_scenario(
    base_scenario: str,
    source: str,
    is_beda_curr: bool,
    is_beda_per: bool,
    num_fac: int,
) -> str:
    """Format nama skenario akhir dengan suffix kondisi khusus."""
    if not source or not base_scenario or base_scenario.upper() in ("UNMATCHING", "UNMATCHED"):
        return "Unmatching"
    base = base_scenario.strip()
    if " only + " in base:
        base = base.replace(" only + ", " + ")
    if is_beda_curr:
        return f"{base} (Beda Currency)"
    if is_beda_per:
        return f"{base} (Beda Periode)"
    if num_fac > 1:
        return f"{base} (> 1 Fac)"
    return base


def _build_output_row(
    sus: dict,
    source: str = None,
    scenario: str = "",
    osbal_rows: list = None,
    osbal_count: int = 0,
    resolved: bool = False,
    suspend_count: int = 1,
    is_beda_curr: bool = False,
    is_beda_per: bool = False,
    all_fac_osbal_rows: list = None,
) -> dict:
    """Bangun satu baris output final."""
    if not osbal_rows or osbal_rows == [{}]:
        ccos = {"CCOS_DOC_NO": "", "CCOS_REF_CODE": "", "CCOS_OR_BAL": np.nan, "CCOS_BAL_DUE": np.nan}
        num_fac = 0
    else:
        ccos = _aggregate_ccos(osbal_rows, all_fac_osbal_rows)
        fac_codes = {str(r.get("CCOS_REF_CODE", "")).strip() for r in osbal_rows if r.get("CCOS_REF_CODE")}
        fac_codes.discard("")
        num_fac = len(fac_codes)

    final_scen = _format_final_scenario(scenario, source, is_beda_curr, is_beda_per, num_fac)

    row = {
        "CCOS_DOC_NO":        ccos["CCOS_DOC_NO"],
        "CCOS_REF_CODE":      ccos["CCOS_REF_CODE"],
        "RECEIPT NO":         sus.get("RECEIPT NO", ""),
        "CREDIT NOTES":       sus.get("CREDIT NOTES", ""),
        "DETAIL RINCIAN NO":  sus.get("DETAIL RINCIAN NO", ""),
        "RECEIPT DATE":       sus.get("RECEIPT DATE", ""),
        "CEDANT NAME":        sus.get("CEDANT NAME", ""),
        "CEDANT SHRT NAME":   sus.get("CEDANT SHRT NAME", ""),
        "INSURED_ORI":        sus.get("INSURED", ""),
        "INSURED_1":          sus.get("INSURED_CLEAN_1", ""),
        "INSURED_2":          sus.get("INSURED_CLEAN_2", ""),
        "CURR ORI":           sus.get("CURR ORI", ""),
        "AMOUNT ORI":         sus.get("AMOUNT ORI", np.nan),
        "AMOUNT_ORI_MIN1":    np.nan,
        "CURR PAY":           sus.get("CURR PAY", ""),
        "AMOUNT PAY":         sus.get("AMOUNT PAY", np.nan),
        "CCOS_OR_BAL":        ccos["CCOS_OR_BAL"],
        "CCOS_BAL_DUE":       ccos["CCOS_BAL_DUE"],
        "DIFERENCE":          np.nan,
        "FLAG_PROD":          "",
        "POLIS":              sus.get("POLIS", ""),
        "POLIS_CLN":          sus.get("POLIS_CLEAN", ""),
        "SLIP_NO":            sus.get("SLIP NO", ""),
        "SLIP_NO_CLN":        sus.get("SLIP_NO_CLEAN", ""),
        "DESC 1":             sus.get("DESC 1", ""),
        "DESC 2":             sus.get("DESC 2", ""),
        "DESC 3":             sus.get("DESC 3", ""),
        "DESC 4":             sus.get("DESC 4", ""),
        "STATUS":             sus.get("STATUS", ""),
        "REC_TYPE":           sus.get("REC_TYPE", ""),
        "SKENARIO":           final_scen,
        "_OSBAL_ROW_COUNT":   osbal_count,
        "_SUSPEND_ROW_COUNT": suspend_count,
    }
    return row


def _merge_suspend_rows(sus_list: list) -> dict:
    """Gabungkan beberapa baris suspend menjadi satu baris agregat."""
    if not sus_list:
        return {}
    if len(sus_list) == 1:
        return dict(sus_list[0])

    merged = dict(sus_list[0])
    for col in _SUSPEND_SUM_COLS:
        total = 0.0
        for s in sus_list:
            v = s.get(col, 0)
            if pd.notna(v) and v != "":
                try:
                    total += float(v)
                except ValueError:
                    pass
        merged[col] = total

    for col in _SUSPEND_JOIN_COLS:
        seen, vals = set(), []
        for s in sus_list:
            v = str(s.get(col, "")).strip()
            if v and v != "nan" and v not in seen:
                seen.add(v)
                vals.append(v)
        merged[col] = ", ".join(vals)
    return merged


def _compute_derived_cols(df: pd.DataFrame) -> pd.DataFrame:
    """Hitung AMOUNT_ORI_MIN1, DIFERENCE, dan FLAG_PROD."""
    amount_ori_base = df.get("AMOUNT ORI", pd.Series([0] * len(df), index=df.index))
    if "_MERGED_AMOUNT_ORI" in df.columns:
        amount_ori_raw = df["_MERGED_AMOUNT_ORI"].combine_first(amount_ori_base)
    else:
        amount_ori_raw = amount_ori_base

    amount_ori_clean = amount_ori_raw.replace(r"^\s*$", np.nan, regex=True)
    if amount_ori_clean.dtype == object:
        amount_ori_clean = (
            amount_ori_clean.astype(str).str.strip().str.replace(r"\s+", "", regex=True)
        )
        amount_ori_clean = amount_ori_clean.replace("nan", np.nan)
        has_comma = amount_ori_clean.str.contains(",", na=False)
        amount_ori_eu = (
            amount_ori_clean
            .str.replace(".", "", regex=False)
            .str.replace(",", ".", regex=False)
        )
        amount_ori_clean = amount_ori_eu.where(has_comma, amount_ori_clean)
    amount_ori_numeric = pd.to_numeric(amount_ori_clean, errors="coerce")

    df["AMOUNT_ORI_MIN1"] = np.where(amount_ori_numeric.notna(), amount_ori_numeric * -1, np.nan)
    amount_ori  = amount_ori_numeric.fillna(0)
    amount_neg  = df["AMOUNT_ORI_MIN1"].fillna(0)

    balance_due     = pd.to_numeric(
        df.get("CCOS_BAL_DUE", pd.Series(0, index=df.index)), errors="coerce"
    ).fillna(0)
    osbal_count_col = pd.to_numeric(
        df.get("_OSBAL_ROW_COUNT", pd.Series(0, index=df.index)), errors="coerce"
    ).fillna(0)
    sus_count_col   = pd.to_numeric(
        df.get("_SUSPEND_ROW_COUNT", pd.Series(1, index=df.index)), errors="coerce"
    ).fillna(1)
    ccos_ref        = df.get("CCOS_REF_CODE", pd.Series("", index=df.index)).fillna("").astype(str)
    skenario_col    = df.get("SKENARIO", pd.Series("", index=df.index)).fillna("")

    df["DIFERENCE"] = amount_neg - balance_due
    accumulated = (osbal_count_col > 1) | (sus_count_col > 1)
    is_equal    = amount_neg.round(2) == balance_due.round(2)

    is_unmatched          = (
        skenario_col.str.strip().str.upper().isin(["UNMATCHING", "UNMATCHED"])
        | (skenario_col == "")
    )
    is_beda_currency      = skenario_col.str.contains("Beda Currency",  na=False, regex=False)
    is_beda_periode       = skenario_col.str.contains("Beda Periode",   na=False, regex=False)
    is_gt1_fac            = (
        skenario_col.str.contains("> 1 Fac", na=False, regex=False)
        | ccos_ref.str.contains(",", na=False)
    )
    is_new_entry_total    = ~is_unmatched & (amount_ori != 0) & (balance_due == 0)
    is_adj_total          = ~is_unmatched & is_equal
    is_adj_sebagian       = ~is_unmatched & ~is_equal & (amount_neg < balance_due)
    is_new_entry_sebagian = ~is_unmatched & ~is_equal & (amount_neg > balance_due)

    df["FLAG_PROD"] = np.select(
        [
            is_beda_currency, is_beda_periode, is_gt1_fac,
            is_new_entry_total,
            is_adj_total & ~accumulated, is_adj_total & accumulated,
            is_adj_sebagian & ~accumulated, is_adj_sebagian & accumulated,
            is_new_entry_sebagian,
        ],
        [
            "Beda Currency", "Beda Periode", "Matching >1 fac code",
            "New Entry total",
            "Adjustment total tanpa akumulasi", "Adjustment total dengan akumulasi",
            "Adjustment sebagian tanpa akumulasi", "Adjustment sebagian dengan akumulasi",
            "New Entry sebagian",
        ],
        default="Unmatching",
    )

    for col in ["_OSBAL_ROW_COUNT", "_SUSPEND_ROW_COUNT", "_MERGED_AMOUNT_ORI"]:
        if col in df.columns:
            df.drop(columns=[col], inplace=True)
    return df


# =============================================================================
# DATA LOADING
# =============================================================================

def _load_excel(path: str, label: str, optional: bool = False) -> tuple:
    """Load Excel ke (list[dict], list[str]) dengan pickle cache otomatis."""
    if not os.path.exists(path):
        print(
            f"  {label}: {path} ({'tidak ada, dilewati' if optional else 'TIDAK DITEMUKAN'})",
            flush=True,
        )
        return [], []

    cache = path + ".cache.pkl"
    if os.path.exists(cache) and os.path.getmtime(cache) >= os.path.getmtime(path):
        t = time.perf_counter()
        print(f"  {label}: loading dari cache ...", flush=True)
        try:
            data, header = pd.read_pickle(cache)
            print(f"  -> {len(data):,} rows, {len(header)} cols ({time.perf_counter()-t:.2f}s)", flush=True)
            return data, header
        except Exception:
            print(f"  {label}: cache rusak, baca ulang dari Excel ...", flush=True)

    t = time.perf_counter()
    print(f"  {label}: {path} ...", flush=True)
    try:
        df = pd.read_excel(path, engine="calamine")
    except Exception:
        try:
            df = pd.read_excel(path)
        except Exception as e:
            print(f"  [RETRY] engine default gagal ({e}), coba openpyxl ...", flush=True)
            df = pd.read_excel(path, engine="openpyxl")

    header = list(df.columns)
    data   = df.to_dict("records")
    print(f"  -> {len(data):,} rows, {len(header)} cols ({time.perf_counter()-t:.1f}s)", flush=True)

    try:
        pd.to_pickle((data, header), cache)
        print(f"  -> cache disimpan: {cache}", flush=True)
    except Exception:
        pass
    return data, header


# =============================================================================
# PIPELINE UTAMA — V2 (AKUMULASI / MERGED ROWS)
# =============================================================================

def run() -> None:
    """Pipeline utama V2 matching suspend Askrindo (output diakumulasi per FAC CODE).

    Urutan matching:
      FASE 1 -> POLIS   (Exact -> LIKE)
      FASE 2 -> SLIP    (Exact -> LIKE)
      FASE 3 -> INSURED (Exact -> LIKE)

    Narrowing:
      Wajib : Currency -> Periode (non-destruktif)
      Basic : cross-field (Polis -> +Slip +Insured, Slip -> +Polis +Insured, Insured -> +Polis +Slip)
    """
    t_start = time.perf_counter()
    print("\n" + "=" * 60, flush=True)
    print("  SUSPEND MATCHING — ASKRINDO — V2 (AKUMULASI)", flush=True)
    print("  Polis -> Slip -> Insured Matching | No Bordero | No Periode-2", flush=True)
    print("=" * 60, flush=True)

    # ------------------------------------------------------------------
    # [1] Load data
    # ------------------------------------------------------------------
    print("\n[1/6] Loading data ...", flush=True)
    suspend_rows, suspend_cols = _load_excel(SUSPEND_FILE, "SUSPEND")
    osbal_rows,   osbal_cols   = _load_excel(OSBAL_FILE,   "OSBAL")
    facul_rows,   facul_cols   = _load_excel(FACUL_FILE,   "FACUL", optional=True)

    if not suspend_rows:
        print("[ERROR] File suspend kosong atau tidak ditemukan. Hentikan.", flush=True)
        return
    if not osbal_rows:
        print("[ERROR] File OSBAL kosong atau tidak ditemukan. Hentikan.", flush=True)
        return

    # Pre-parse tanggal OSBAL
    print("  Pre-parsing OSBAL FAC_COM_DATE ...", flush=True)
    for r in osbal_rows:
        raw = r.get(OSBAL_DATE_COL, "")
        r["_com_date_parsed"] = (
            pd.to_datetime(raw, errors="coerce")
            if raw and not (isinstance(raw, float) and pd.isna(raw))
            else None
        )

    # Pre-parse tanggal Suspend
    print("  Pre-parsing Suspend RECEIPT DATE ...", flush=True)
    for r in suspend_rows:
        raw = r.get(SUSPEND_DATE_COL, "")
        r["_sus_date_parsed"] = (
            pd.to_datetime(raw, errors="coerce")
            if raw and not (isinstance(raw, float) and pd.isna(raw))
            else None
        )

    # ------------------------------------------------------------------
    # [2] Deteksi kolom bersih
    # ------------------------------------------------------------------
    print("\n[2/6] Detecting clean columns ...", flush=True)

    polis_sus   = _get_clean_cols(suspend_cols, "clean polis")
    slip_sus    = _get_clean_cols(suspend_cols, "clean slip")
    insured_sus = _get_clean_cols(suspend_cols, "clean insured")

    if not polis_sus:
        polis_sus = [c for c in ["POLIS_CLEAN", "clean polis 1"] if c in suspend_cols]
    if not slip_sus:
        slip_sus = [c for c in ["SLIP_NO_CLEAN", "clean slip 1"] if c in suspend_cols]
    if not insured_sus:
        insured_sus = [c for c in ["INSURED_CLEAN_1", "INSURED_CLEAN_2", "clean insured 1"]
                       if c in suspend_cols]

    print(f"  Suspend -> polis_sus: {polis_sus}", flush=True)
    print(f"  Suspend -> slip_sus:  {slip_sus}", flush=True)
    print(f"  Suspend -> ins_sus:   {insured_sus}", flush=True)

    # Pre-parse LINESLIP flags
    print("  Pre-parsing LINESLIP flags ...", flush=True)
    lineslip_count = 0
    for r in suspend_rows:
        r["_is_lineslip"] = _is_lineslip_row(r, insured_sus)
        if r["_is_lineslip"]:
            lineslip_count += 1
            parsed = r.get("_sus_date_parsed")
            try:
                r["_sus_date_lineslip"] = (
                    (parsed - pd.DateOffset(months=1)) if parsed is not None else None
                )
            except Exception:
                r["_sus_date_lineslip"] = None
        else:
            r["_sus_date_lineslip"] = None
    print(f"  -> {lineslip_count:,} baris LINESLIP terdeteksi", flush=True)

    # OSBAL clean cols
    polis_osbal   = _get_clean_cols(osbal_cols, "clean polis")
    slip_osbal    = _get_clean_cols(osbal_cols, "clean slip")
    insured_osbal = _get_clean_cols(osbal_cols, "clean insured")

    def _prepend(col, cols_list, header):
        return ([col] + cols_list) if col in header and col not in cols_list else cols_list

    slip_osbal  = _prepend("CLSDT_SLIP_NO",   slip_osbal,  osbal_cols)
    polis_osbal = _prepend("CLSDT_POLICY_NO", polis_osbal, osbal_cols)

    print(f"  OSBAL  -> polis_osbal: {polis_osbal}", flush=True)
    print(f"  OSBAL  -> slip_osbal:  {slip_osbal}", flush=True)
    print(f"  OSBAL  -> ins_osbal:   {insured_osbal}", flush=True)

    # FACUL clean cols
    polis_facul   = _get_clean_cols(facul_cols, "clean polis")   if facul_rows else []
    slip_facul    = _get_clean_cols(facul_cols, "clean slip")    if facul_rows else []
    insured_facul = _get_clean_cols(facul_cols, "clean insured") if facul_rows else []

    if facul_rows:
        slip_facul  = _prepend("CLSDT_SLIP_NO",   slip_facul,  facul_cols)
        polis_facul = _prepend("CLSDT_POLICY_NO", polis_facul, facul_cols)
        print(f"  FACUL  -> polis_facul: {polis_facul}", flush=True)
        print(f"  FACUL  -> slip_facul:  {slip_facul}", flush=True)
        print(f"  FACUL  -> ins_facul:   {insured_facul}", flush=True)

    # ------------------------------------------------------------------
    # [3] Bangun lookup index
    # ------------------------------------------------------------------
    print("\n[3/6] Building lookup indexes ...", flush=True)
    lkp_osbal = _build_lookup(
        osbal_rows, polis_osbal, slip_osbal, insured_osbal,
        facode_col=OSBAL_FACODE_COL, label="OSBAL",
    )
    lkp_facul = (
        _build_lookup(
            facul_rows, polis_facul, slip_facul, insured_facul,
            facode_col=FACUL_FACODE_COL, label="FACUL",
        )
        if facul_rows
        else {}
    )

    facode_osbal_idx = lkp_osbal.get("facode", {})

    # ------------------------------------------------------------------
    # [4] Matching per baris
    # ------------------------------------------------------------------
    total = len(suspend_rows)
    print(f"\n[4/6] Matching {total:,} suspend rows ...", flush=True)
    print("  Strategi: POLIS -> SLIP -> INSURED (Exact -> LIKE)", flush=True)
    t4 = time.perf_counter()
    raw_results = []

    for n, sus in enumerate(suspend_rows, 1):
        if n % 500 == 0:
            print(f"    Memproses baris {n:,} ...", flush=True)

        # Skip baris yang sudah MATCHING
        if str(sus.get("STATUS", "")).strip().upper() == "MATCHING":
            continue

        _is_ls         = sus.get("_is_lineslip", False)
        effective_date = sus.get("_sus_date_lineslip") if _is_ls else sus.get("_sus_date_parsed")

        matched_osbal_idx: set = set()
        source   = None
        scenario = ""
        resolved = False

        # ----------------------------------------------------------------
        # FASE 1: POLIS (urutan ke-1)
        # ----------------------------------------------------------------
        polis_vals = _collect_clean_values(sus, polis_sus)
        polis_vals = [p for p in polis_vals if len(p) >= 5 and p not in ["-", "TBA", "VARIOUS"]]

        if polis_vals:
            # 1a. Exact polis -> OSBAL (kolom polis & cross ke kolom slip)
            hits = _exact_match(polis_vals, lkp_osbal["exact_polis"])
            if not hits:
                hits = _exact_match(polis_vals, lkp_osbal["exact_slip"])
            if hits:
                matched_osbal_idx.update(hits)
                source   = "OSBAL"
                scenario = "Polis only"
                resolved = True

            # 1b. Exact polis -> FACUL (resolve ke OSBAL via FAC_CODE)
            if not matched_osbal_idx and facul_rows:
                hits = _exact_match(polis_vals, lkp_facul.get("exact_polis", {}))
                if not hits:
                    hits = _exact_match(polis_vals, lkp_facul.get("exact_slip", {}))
                if hits:
                    res_idx, ok = _resolve_facode(
                        [facul_rows[i] for i in hits],
                        FACUL_FACODE_COL, facode_osbal_idx, osbal_rows,
                    )
                    if ok:
                        matched_osbal_idx.update(res_idx)
                        source   = "FACUL"
                        scenario = "Polis only"
                        resolved = True

            # 1c. LIKE polis -> OSBAL (min_len=5, polis & slip cols)
            if not matched_osbal_idx:
                tok_polis = lkp_osbal.get("token_polis", ({}, {}))
                if tok_polis[0] and polis_vals:
                    hits = _like_match(polis_vals, tok_polis[0], osbal_rows, polis_osbal, min_len=5)
                if not hits:
                    tok_slip = lkp_osbal.get("token_slip", ({}, {}))
                    if tok_slip[0] and polis_vals:
                        hits = _like_match(polis_vals, tok_slip[0], osbal_rows, slip_osbal, min_len=5)
                if hits:
                    matched_osbal_idx.update(hits)
                    source   = "OSBAL"
                    scenario = "Polis only"
                    resolved = True

            # 1d. LIKE polis -> FACUL (resolve ke OSBAL via FAC_CODE)
            if not matched_osbal_idx and facul_rows:
                tok_pol_f = lkp_facul.get("token_polis", ({}, {}))
                if tok_pol_f[0] and polis_vals:
                    hits = _like_match(polis_vals, tok_pol_f[0], facul_rows, polis_facul, min_len=5)
                if not hits:
                    tok_slp_f = lkp_facul.get("token_slip", ({}, {}))
                    if tok_slp_f[0] and polis_vals:
                        hits = _like_match(polis_vals, tok_slp_f[0], facul_rows, slip_facul, min_len=5)
                if hits:
                    res_idx, ok = _resolve_facode(
                        [facul_rows[i] for i in hits],
                        FACUL_FACODE_COL, facode_osbal_idx, osbal_rows,
                    )
                    if ok:
                        matched_osbal_idx.update(res_idx)
                        source   = "FACUL"
                        scenario = "Polis only"
                        resolved = True

        # ----------------------------------------------------------------
        # FASE 2: SLIP (urutan ke-2)
        # ----------------------------------------------------------------
        if not matched_osbal_idx:
            slip_vals = _collect_clean_values(sus, slip_sus)
            slip_vals = [s for s in slip_vals if len(s) >= 4 and s not in ["-", "TBA"]]

            # Tambahkan variasi query slip: base 10-digit unik jika panjang >= 14 (karakteristik slip Askrindo)
            slip_queries = list(slip_vals)
            for sv in slip_vals:
                if len(sv) >= 14 and sv[-10:].isdigit():
                    base10 = sv[-10:]
                    if base10 not in slip_queries:
                        slip_queries.append(base10)

            if slip_queries:
                # 2a. Exact slip -> OSBAL (kolom slip & cross ke kolom polis)
                hits = _exact_match(slip_queries, lkp_osbal["exact_slip"])
                if not hits:
                    hits = _exact_match(slip_queries, lkp_osbal["exact_polis"])
                if hits:
                    matched_osbal_idx.update(hits)
                    source   = "OSBAL"
                    scenario = "Slip only"
                    resolved = True

                # 2b. Exact slip -> FACUL (resolve ke OSBAL via FAC_CODE)
                if not matched_osbal_idx and facul_rows:
                    hits = _exact_match(slip_queries, lkp_facul.get("exact_slip", {}))
                    if not hits:
                        hits = _exact_match(slip_queries, lkp_facul.get("exact_polis", {}))
                    if hits:
                        res_idx, ok = _resolve_facode(
                            [facul_rows[i] for i in hits],
                            FACUL_FACODE_COL, facode_osbal_idx, osbal_rows,
                        )
                        if ok:
                            matched_osbal_idx.update(res_idx)
                            source   = "FACUL"
                            scenario = "Slip only"
                            resolved = True

                # 2c. LIKE slip -> OSBAL (min_len=8, slip & polis cols)
                if not matched_osbal_idx:
                    tok_slip = lkp_osbal.get("token_slip", ({}, {}))
                    if tok_slip[0]:
                        hits = _like_match(slip_queries, tok_slip[0], osbal_rows, slip_osbal, min_len=8)
                    if not hits:
                        tok_pol = lkp_osbal.get("token_polis", ({}, {}))
                        if tok_pol[0]:
                            hits = _like_match(slip_queries, tok_pol[0], osbal_rows, polis_osbal, min_len=8)
                    if hits:
                        matched_osbal_idx.update(hits)
                        source   = "OSBAL"
                        scenario = "Slip only"
                        resolved = True

                # 2d. LIKE slip -> FACUL (resolve ke OSBAL via FAC_CODE)
                if not matched_osbal_idx and facul_rows:
                    tok_slip_f = lkp_facul.get("token_slip", ({}, {}))
                    if tok_slip_f[0]:
                        hits = _like_match(slip_queries, tok_slip_f[0], facul_rows, slip_facul, min_len=8)
                    if not hits:
                        tok_pol_f = lkp_facul.get("token_polis", ({}, {}))
                        if tok_pol_f[0]:
                            hits = _like_match(slip_queries, tok_pol_f[0], facul_rows, polis_facul, min_len=8)
                    if hits:
                        res_idx, ok = _resolve_facode(
                            [facul_rows[i] for i in hits],
                            FACUL_FACODE_COL, facode_osbal_idx, osbal_rows,
                        )
                        if ok:
                            matched_osbal_idx.update(res_idx)
                            source   = "FACUL"
                            scenario = "Slip only"
                            resolved = True

        # ----------------------------------------------------------------
        # FASE 3: INSURED (urutan ke-3)
        # ----------------------------------------------------------------
        if not matched_osbal_idx:
            # Prioritaskan INSURED_CLEAN_1 dulu agar tidak melebar ke joint contractor (misal Pembangunan Perumahan)
            primary_cols = [c for c in insured_sus if "1" in c or c == "INSURED_CLEAN"]
            if not primary_cols and insured_sus:
                primary_cols = [insured_sus[0]]
            secondary_cols = [c for c in insured_sus if c not in primary_cols]

            prim_vals = _collect_clean_values(sus, primary_cols)
            prim_vals = [i for i in prim_vals if len(i) >= 3 and i not in ["-", "TBA"]]
            sec_vals  = _collect_clean_values(sus, secondary_cols)
            sec_vals  = [i for i in sec_vals if len(i) >= 3 and i not in ["-", "TBA"]]

            for ins_vals in [prim_vals, sec_vals]:
                if not ins_vals or matched_osbal_idx:
                    continue

                # 3a. Exact insured -> OSBAL
                hits = _exact_match(ins_vals, lkp_osbal["exact_insured"])
                if hits:
                    matched_osbal_idx.update(hits)
                    source   = "OSBAL"
                    scenario = "Insured only"
                    resolved = True
                    break

                # 3b. Exact insured -> FACUL (resolve ke OSBAL via FAC_CODE)
                if not matched_osbal_idx and facul_rows:
                    hits = _exact_match(ins_vals, lkp_facul.get("exact_insured", {}))
                    if hits:
                        res_idx, ok = _resolve_facode(
                            [facul_rows[i] for i in hits],
                            FACUL_FACODE_COL, facode_osbal_idx, osbal_rows,
                        )
                        if ok:
                            matched_osbal_idx.update(res_idx)
                            source   = "FACUL"
                            scenario = "Insured only"
                            resolved = True
                            break

                # 3c. LIKE insured -> OSBAL (min_len=6)
                if not matched_osbal_idx:
                    tok_ins = lkp_osbal.get("token_insured", ({}, {}))
                    if tok_ins[0] and ins_vals:
                        hits = _like_match(ins_vals, tok_ins[0], osbal_rows, insured_osbal, min_len=6)
                        if hits:
                            matched_osbal_idx.update(hits)
                            source   = "OSBAL"
                            scenario = "Insured only"
                            resolved = True
                            break

                # 3d. LIKE insured -> FACUL (resolve ke OSBAL via FAC_CODE)
                if not matched_osbal_idx and facul_rows:
                    tok_ins_f = lkp_facul.get("token_insured", ({}, {}))
                    if tok_ins_f[0] and ins_vals:
                        hits = _like_match(ins_vals, tok_ins_f[0], facul_rows, insured_facul, min_len=6)
                        if hits:
                            res_idx, ok = _resolve_facode(
                                [facul_rows[i] for i in hits],
                                FACUL_FACODE_COL, facode_osbal_idx, osbal_rows,
                            )
                            if ok:
                                matched_osbal_idx.update(res_idx)
                                source   = "FACUL"
                                scenario = "Insured only"
                                resolved = True
                                break

        # ----------------------------------------------------------------
        # FASE 4: NARROWING WAJIB — Currency & Periode (non-destruktif)
        # ----------------------------------------------------------------
        is_beda_curr = False
        is_beda_per  = False

        if matched_osbal_idx:
            matched_list = list(matched_osbal_idx)

            # 4a. Narrowing currency
            if len(matched_list) > 1:
                matched_list, is_beda_curr = _narrow_by_currency(sus, matched_list, osbal_rows)

            # 4b. Narrowing periode (hanya jika curr OK)
            if not is_beda_curr and len(matched_list) > 1:
                matched_list, is_beda_per = _narrow_by_periode(
                    sus, matched_list, osbal_rows,
                    effective_sus_date=effective_date, is_lineslip=_is_ls,
                )

            matched_osbal_idx = set(matched_list)

        # ----------------------------------------------------------------
        # FASE 5: NARROWING BASIC — Cross-field (non-destruktif)
        # ----------------------------------------------------------------
        if matched_osbal_idx and not is_beda_curr and not is_beda_per:
            matched_list = list(matched_osbal_idx)

            if len(matched_list) > 1:
                matched_list, scenario = _narrow_basic(
                    sus, matched_list, osbal_rows,
                    match_label=scenario,
                    polis_ref_cols=polis_osbal,
                    slip_ref_cols=slip_osbal,
                    insured_ref_cols=insured_osbal,
                    polis_sus_cols=polis_sus,
                    slip_sus_cols=slip_sus,
                    insured_sus_cols=insured_sus,
                )
            else:
                scenario = _map_scenario(scenario)

            matched_osbal_idx = set(matched_list)

        # ----------------------------------------------------------------
        # Finalisasi hasil per baris
        # ----------------------------------------------------------------
        if is_beda_curr or is_beda_per:
            osbal_count = 0
            resolved    = False
        elif not matched_osbal_idx:
            source      = None
            scenario    = "Unmatching"
            osbal_count = 0
            resolved    = False
        else:
            osbal_count = len(matched_osbal_idx)

        if scenario and " only + " in scenario:
            scenario = scenario.replace(" only + ", " + ")

        fac_codes = {
            _normalize(osbal_rows[i].get(OSBAL_FACODE_COL, ""))
            for i in matched_osbal_idx
        }
        fac_codes.discard("")

        raw_results.append({
            "suspend":       sus,
            "source":        source,
            "base_scenario": scenario,
            "is_beda_curr":  is_beda_curr,
            "is_beda_per":   is_beda_per,
            "osbal_idx":     list(matched_osbal_idx),
            "osbal_count":   osbal_count,
            "resolved":      resolved,
            "fac_codes":     fac_codes,
        })

    print(f"  Matching selesai dalam {time.perf_counter()-t4:.1f}s", flush=True)

    # ------------------------------------------------------------------
    # [4c/6] Fac code claim & accumulation (V2: Merge rows per FAC CODE)
    # ------------------------------------------------------------------
    print("\n[4c/6] Fac code claim & accumulation (V2: merge rows) ...", flush=True)

    claimed: set = set()
    for r in raw_results:
        if (
            len(r["fac_codes"]) == 1
            and r["source"] is not None
            and not r["is_beda_curr"]
            and not r["is_beda_per"]
        ):
            claimed |= r["fac_codes"]

    for r in raw_results:
        if len(r["fac_codes"]) > 1:
            remaining = r["fac_codes"] - claimed
            if (
                len(remaining) == 1
                and r["source"] is not None
                and not r["is_beda_curr"]
                and not r["is_beda_per"]
            ):
                r["fac_codes"] = remaining
                claimed |= remaining
                fac = next(iter(remaining))
                r["osbal_idx"] = [
                    i for i in r["osbal_idx"]
                    if _normalize(osbal_rows[i].get(OSBAL_FACODE_COL, "")) == fac
                ]
                r["osbal_count"] = len(r["osbal_idx"])

    # Group per fac untuk di-merge menjadi 1 baris per FAC CODE (Mode V2)
    groups: dict = {}
    final:  list = []

    for r in raw_results:
        r.setdefault("suspend_count", 1)
        if (
            len(r["fac_codes"]) == 1
            and r["source"] is not None
            and not r["is_beda_curr"]
            and not r["is_beda_per"]
        ):
            groups.setdefault(next(iter(r["fac_codes"])), []).append(r)
        else:
            final.append(r)

    for fac, group in groups.items():
        if len(group) == 1:
            group[0]["suspend_count"] = 1
            final.append(group[0])
            continue
        merged_osbal_idx = set()
        merged_fac_codes = set()
        for g in group:
            merged_osbal_idx.update(g["osbal_idx"])
            merged_fac_codes.update(g["fac_codes"])
        merged_sus = _merge_suspend_rows([g["suspend"] for g in group])
        ref = group[0]
        final.append({
            "suspend":            merged_sus,
            "source":             ref["source"],
            "base_scenario":      ref["base_scenario"],
            "is_beda_curr":       False,
            "is_beda_per":        False,
            "osbal_idx":          list(merged_osbal_idx),
            "osbal_count":        len(merged_osbal_idx),
            "resolved":           True,
            "fac_codes":          merged_fac_codes,
            "suspend_count":      len(group),
            "_merged_amount_ori": merged_sus.get("AMOUNT ORI", 0),
        })

    print(
        f"  {len(raw_results):,} suspend rows -> {len(final):,} output rows "
        f"({len(raw_results) - len(final):,} baris diserap akumulasi)",
        flush=True,
    )

    # ------------------------------------------------------------------
    # [5/6] Building output DataFrame
    # ------------------------------------------------------------------
    print(f"\n[5/6] Building output ({len(final):,} rows) ...", flush=True)
    output_rows = []
    for r in final:
        ref_osbal = [osbal_rows[i] for i in r["osbal_idx"]] if r["source"] else [{}]
        all_fac_osbal_rows = None
        osbal_cnt = r["osbal_count"]

        if r["source"] and len(r["fac_codes"]) == 1:
            fac = next(iter(r["fac_codes"]))
            all_fac_idx = facode_osbal_idx.get(fac, [])
            if len(all_fac_idx) > 0:
                all_fac_osbal_rows = [osbal_rows[i] for i in all_fac_idx]
                osbal_cnt          = len(all_fac_idx)

        out = _build_output_row(
            r["suspend"],
            source=r["source"],
            scenario=r["base_scenario"],
            osbal_rows=ref_osbal,
            osbal_count=osbal_cnt,
            resolved=r["resolved"],
            suspend_count=r.get("suspend_count", 1),
            is_beda_curr=r.get("is_beda_curr", False),
            is_beda_per=r.get("is_beda_per", False),
            all_fac_osbal_rows=all_fac_osbal_rows,
        )
        if "_merged_amount_ori" in r:
            out["_MERGED_AMOUNT_ORI"] = r["_merged_amount_ori"]
        output_rows.append(out)

    df_out = pd.DataFrame(output_rows)
    df_out = _compute_derived_cols(df_out)

    cols_to_use = [c for c in FINAL_COLUMNS if c in df_out.columns]
    df_out      = df_out[cols_to_use]

    # ------------------------------------------------------------------
    # [6/6] Save output
    # ------------------------------------------------------------------
    print(f"\n[6/6] Saving to: {OUTPUT_FILE} ...", flush=True)
    os.makedirs(os.path.dirname(os.path.abspath(OUTPUT_FILE)), exist_ok=True)
    df_out.to_excel(OUTPUT_FILE, index=False)

    try:
        from excel_styler import apply_purple_column_style
        apply_purple_column_style(OUTPUT_FILE, "matching")
    except Exception as e:
        print(f"  [WARN] Styling Excel gagal ({e}), dilewati.", flush=True)

    # Ekspor ke PostgreSQL (opsional)
    try:
        load_dotenv()
        user     = os.getenv("user")
        password = os.getenv("password")
        host     = os.getenv("host")
        port     = os.getenv("port", "5432")
        db       = os.getenv("db")
        if all([user, password, host, db]):
            engine    = create_engine(f"postgresql://{user}:{password}@{host}:{port}/{db}")
            table_name = "SUSPENSE_DATA_SUSPENSE_ASKRINDO"
            print(f"  Exporting ke PostgreSQL -> {table_name} ...", flush=True)
            with engine.connect() as conn:
                conn.execute(
                    text(f'DELETE FROM "{table_name}" WHERE "CEDANT SHRT NAME" = :cedant;'),
                    {"cedant": "ASKRINDO"},
                )
                conn.commit()
            df_out.to_sql(table_name, engine, if_exists="append", index=False)
            print("  -> PostgreSQL export berhasil.", flush=True)
    except Exception as e:
        print(f"  [WARN] PostgreSQL export dilewati / gagal: {e}", flush=True)

    # Summary
    elapsed_total = time.perf_counter() - t_start
    print("\n" + "=" * 60, flush=True)
    print(f"  Selesai dalam {elapsed_total:.1f}s ({elapsed_total/60:.1f} menit)", flush=True)
    print(f"  Rows: {len(df_out):,}  |  Cols: {len(df_out.columns)}  |  File: {OUTPUT_FILE}", flush=True)
    if "FLAG_PROD" in df_out.columns:
        print("\n  FLAG_PROD:", flush=True)
        for val, cnt in df_out["FLAG_PROD"].value_counts().items():
            print(f"    {val:<45}: {cnt:,}", flush=True)
    if "SKENARIO" in df_out.columns:
        print("\n  SKENARIO:", flush=True)
        for val, cnt in df_out["SKENARIO"].value_counts().items():
            print(f"    {val:<45}: {cnt:,}", flush=True)
    print("=" * 60, flush=True)


if __name__ == "__main__":
    run()