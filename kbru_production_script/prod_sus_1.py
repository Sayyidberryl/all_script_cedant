"""
prod_sus_1.py — Suspend Matching v1 for Cedant KBRU (Row Asli Dipertahankan)

Spesifikasi KBRU:
1. Urutan matching: SLIP -> POLIS -> INSURED
2. Tidak menggunakan sertifikat sama sekali
3. Tidak menggunakan periode 2 (tanpa offset lineslip 1 bulan, sus_date >= com_date standar)
4. Hanya basic rules (tanpa custom rules ACA atau bordero)
5. Export ke PostgreSQL (SUSPENSE_DATA_SUSPENSE_V1 & SUSPENSE_CLEAN_2026) per-cedant update
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

SUSPEND_FILE = os.path.join(DATA_DIR, "kbru_output_suspend.xlsx")
OSBAL_FILE   = os.path.join(DATA_DIR, "kbru_output_osbal.xlsx")
FACUL_FILE   = os.path.join(DATA_DIR, "kbru_output_facul.xlsx")
OUTPUT_FILE  = os.path.join(DATA_DIR, "final_output_v1.xlsx")

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

# Kolom FAC CODE
OSBAL_FACODE_COL  = "CCOS_REF_CODE"
FACUL_FACODE_COL  = "FAC_CODE"
SLIPDB_FACODE_COL = "FAC_CODE"

# Kolom currency dan tanggal
SUSPEND_CURR_COL = "CURR ORI"
OSBAL_CURR_COL   = "CCOS_CURR"
SUSPEND_DATE_COL = "RECEIPT DATE"
OSBAL_DATE_COL   = "FAC_COM_DATE"

# Marker baris administratif yang dikecualikan
_EXCLUDED_REF_MARKERS = ["HUTANG PIUTANG", "DATA SUSPENSE"]
_DUMMY_VALUES = frozenset({"-", "--", "---", "0", "00", "000", "NONE", "NULL", "TBA", "NA", "N/A", "SELISIH", "NO", "NON", "."})

_MIN_TOKEN_LEN = 3
_TOKEN_RE      = re.compile(r"[^A-Z0-9]+")
_DELIMITER_RE  = re.compile(r"[,;|+\s]+")

# Cedant KBRU identifiers
KBRU_CEDANT_NAME = "PT. KBRU REINSURANCE BROKERS (EX. SIMAS)"
KBRU_SHRT_NAME   = "KBRU"

# Kolom suspend yang di-join jika ada row merging
_SUSPEND_JOIN_COLS = [
    "RECEIPT NO", "CREDIT NOTES", "DETAIL RINCIAN NO", "RECEIPT DATE",
    "CEDANT NAME", "CEDANT SHRT NAME",
    "INSURED_ORI", "INSURED_1", "INSURED_2",
    "CURR ORI", "CURR PAY", "AMOUNT PAY",
    "POLIS", "POLICY_CLEAN_1",
    "SLIP_NO", "SLIP_NO_CLN",
    "DESC 1", "DESC 2", "DESC 3", "DESC 4",
    "STATUS", "REC_TYPE",
]
_SUSPEND_SUM_COLS = ["AMOUNT ORI"]

# 38 Kolom standar output
FINAL_COLUMNS = [
    "CCOS_DOC_NO", "CCOS_REF_CODE",
    "RECEIPT NO", "CREDIT NOTES", "DETAIL RINCIAN NO", "RECEIPT DATE",
    "CEDANT NAME", "CEDANT SHRT NAME",
    "INSURED_ORI", "INSURED_1", "INSURED_2",
    "CURR ORI", "AMOUNT ORI", "AMOUNT_ORI_MIN1",
    "CURR PAY", "AMOUNT PAY",
    "CCOS_OR_BAL", "CCOS_BAL_DUE", "DIFERENCE",
    "FLAG_PROD",
    "POLIS",
    "POLICY_CLEAN_1", "SERTIF_CLEAN_1",
    "POLICY_CLEAN_2", "SERTIF_CLEAN_2",
    "POLICY_CLEAN_3", "SERTIF_CLEAN_3",
    "POLICY_CLEAN_4", "POLICY_CLEAN_5",
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
    text = s.strip().upper().replace("S/D", "SD")
    return re.sub(r"\s+", " ", text) if "  " in text else text

def _normalize(value) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    return _norm_cached(str(value))

def _clean_alphanumeric(s: str) -> str:
    if not s:
        return ""
    return re.sub(r"[^A-Z0-9]", "", str(s).upper())

def _tokenize(text: str) -> frozenset:
    if not text:
        return frozenset()
    return frozenset(t for t in _TOKEN_RE.split(text) if len(t) >= _MIN_TOKEN_LEN)

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

def _is_valid_token(s: str) -> bool:
    if not s or (isinstance(s, float) and pd.isna(s)):
        return False
    norm = _norm_cached(str(s))
    if norm in _DUMMY_VALUES:
        return False
    c = _clean_alphanumeric(norm)
    return len(c) >= 3 and c not in _DUMMY_VALUES

def _extract_polis_variants(val: str) -> list:
    if not _is_valid_token(val):
        return []
    s = _normalize(val)
    variants = set()

    # 1. Alphanumeric polos
    c = _clean_alphanumeric(s)
    if len(c) >= 6:
        variants.add(c)
        # Bentuk tanpa kata broker (KBRURE / SRB)
        no_broker = re.sub(r'KBRURE|SRB', '', c)
        if len(no_broker) >= 6:
            variants.add(no_broker)

    # 2. Pola registrasi polis KBRU: 00004903/FEQI/054.0000002-SRB/08/2023 atau 00000481.FPAR.002.000005.SRB
    m = re.search(r'0*(\d{3,8})[/.\s]+([A-Z]{3,5})[/.\s]+0*(\d+)[/.\s]+0*(\d+)', s)
    if m:
        doc_num, cls, p3, p4 = m.groups()
        variants.add(f"{doc_num}{cls}{p3}{p4}")
        variants.add(f"{doc_num}{cls}")
    else:
        m2 = re.search(r'0*(\d{3,8})[/.\s]+([A-Z]{3,5})', s)
        if m2:
            doc_num, cls = m2.groups()
            variants.add(f"{doc_num}{cls}")

    # 3. Multi-polis dengan '+'
    if '+' in s:
        for part in s.split('+'):
            if _is_valid_token(part):
                variants.update(_extract_polis_variants(part))

    return [v for v in variants if len(v) >= 3 and v not in _DUMMY_VALUES]

def _extract_slip_variants(val: str) -> list:
    if not _is_valid_token(val):
        return []
    s = _normalize(val)
    variants = set()

    variants.add(s)
    c = _clean_alphanumeric(s)
    if len(c) >= 3:
        variants.add(c)

    # Strip -01, -02 (installment suffix)
    base = re.sub(r'-\d+$', '', s)
    if base and base != s and _is_valid_token(base):
        variants.add(base)
        c_base = _clean_alphanumeric(base)
        if len(c_base) >= 3:
            variants.add(c_base)

    if '+' in s:
        for part in s.split('+'):
            if _is_valid_token(part):
                variants.update(_extract_slip_variants(part))

    return [v for v in variants if len(v) >= 3 and v not in _DUMMY_VALUES]

def _collect_clean_values(row: dict, cols: list, is_slip: bool = False, is_polis: bool = False) -> list:
    seen, result = set(), []
    for col in cols:
        raw_val = row.get(col, "")
        if not _is_valid_token(raw_val):
            continue
        val = _normalize(raw_val)
        if not val or val in _DUMMY_VALUES:
            continue

        if is_polis:
            for pv in _extract_polis_variants(val):
                if pv not in seen:
                    seen.add(pv)
                    result.append(pv)
        elif is_slip:
            for sv in _extract_slip_variants(val):
                if sv not in seen:
                    seen.add(sv)
                    result.append(sv)
        else:
            if val not in seen:
                seen.add(val)
                result.append(val)
    return result

def _collect_all_values(row: dict, clean_cols: list, ori_col: str, is_slip: bool = False, is_polis: bool = False) -> list:
    vals = _collect_clean_values(row, clean_cols, is_slip=is_slip, is_polis=is_polis)
    if not vals and ori_col:
        raw_val = row.get(ori_col, "")
        if _is_valid_token(raw_val):
            val_ori = _normalize(raw_val)
            if val_ori and val_ori not in _DUMMY_VALUES:
                if is_polis:
                    vals = _extract_polis_variants(val_ori)
                elif is_slip:
                    vals = _extract_slip_variants(val_ori)
                else:
                    vals = [val_ori]
    return list(set(vals))

def _ref_has_value(ref_row: dict, clean_cols: list, ori_col: str, query_values: list, is_slip: bool = False, is_polis: bool = False) -> bool:
    if not query_values:
        return False
    q_set = set(query_values)
    ref_cols = list(clean_cols)
    if ori_col and ori_col not in ref_cols:
        ref_cols.append(ori_col)
    ref_vals = _collect_clean_values(ref_row, ref_cols, is_slip=is_slip, is_polis=is_polis)
    return any(v in q_set for v in ref_vals)

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
        _build_token_index(rows, slip_cols,    excluded=excluded),                # [3]
        _build_token_index(rows, polis_cols,   excluded=excluded),                # [4]
        _build_token_index(rows, insured_cols, excluded=excluded),                # [5]
        facode_index,                                                             # [6]
    )
    elapsed = time.perf_counter() - t0
    if label:
        exc_info = f" - {len(excluded):,} baris dikecualikan" if excluded else ""
        print(f"  {label:<6} done ({elapsed:.1f}s){exc_info}", flush=True)
    return lookup

def _exact_match(query_values: list, index: dict) -> set:
    return {i for v in query_values if v and v in index for i in index[v]}

def _like_match(query_values: list, token_index: dict, rows: list, ref_cols: list) -> list:
    valid_queries = [v for v in query_values if v and len(v) >= 3]
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

def _resolve_facode(ref_rows_matched: list, facode_col: str, facode_osbal_idx: dict, osbal_rows: list) -> tuple:
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

# =============================================================================
# REMATCH STRICT & OUTPUT AGGREGATION
# =============================================================================

def _rematch_strict(
    suspend_row: dict, osbal_rows: list, lookup_osbal: tuple,
    slip_sus_cols: list, polis_sus_cols: list, insured_sus_cols: list,
    slip_ref_cols: list, polis_ref_cols: list, insured_ref_cols: list,
    facode_osbal_idx: dict,
    slipdb_rows: list = None, lookup_slipdb: tuple = None,
    facul_rows: list = None, lookup_facul: tuple = None,
    effective_sus_date = None
) -> tuple:
    sus_curr = _normalize(suspend_row.get(SUSPEND_CURR_COL, ""))

    def _strict_filter(idxs):
        res = []
        for i in idxs:
            if sus_curr and _normalize(osbal_rows[i].get(OSBAL_CURR_COL, "")) != sus_curr:
                continue
            if pd.notna(effective_sus_date):
                com_date = osbal_rows[i].get("_com_date_parsed")
                if pd.notna(com_date):
                    if effective_sus_date < com_date:
                        continue
            res.append(i)
        return res

    # R1: OSBAL Slip exact (Urutan 1: Slip)
    slip_vals = _collect_clean_values(suspend_row, slip_sus_cols, is_slip=True)
    if slip_vals:
        hits = _strict_filter(_exact_match(slip_vals, lookup_osbal[0]))
        if hits:
            polis_vals = _collect_clean_values(suspend_row, polis_sus_cols, is_polis=True)
            confirmed = [i for i in hits if _ref_has_value(osbal_rows[i], polis_ref_cols, "FAC_POLICY_NO_ORI", polis_vals, is_polis=True)]
            chosen = confirmed if confirmed else hits
            return "OSBAL", ("Slip + Polis" if confirmed else "Slip only"), chosen, True

        # R2: OSBAL Slip LIKE
        if lookup_osbal[3] and lookup_osbal[3][0]:
            like_hits = _strict_filter(_like_match(slip_vals, lookup_osbal[3][0], osbal_rows, slip_ref_cols))
            if like_hits:
                return "OSBAL", "Slip only", like_hits, True

    # R3: OSBAL Polis exact (Urutan 2: Polis)
    polis_vals = _collect_clean_values(suspend_row, polis_sus_cols, is_polis=True)
    if polis_vals:
        hits = _strict_filter(_exact_match(polis_vals, lookup_osbal[1]))
        if hits:
            confirmed = [i for i in hits if _ref_has_value(osbal_rows[i], slip_ref_cols, "FAC_SLIP_ORI", slip_vals, is_slip=True)]
            chosen = confirmed if confirmed else hits
            return "OSBAL", ("Polis + Slip" if confirmed else "Polis only"), chosen, True

        # R4: OSBAL Polis LIKE
        if lookup_osbal[4] and lookup_osbal[4][0]:
            like_hits = _strict_filter(_like_match(polis_vals, lookup_osbal[4][0], osbal_rows, polis_ref_cols))
            if like_hits:
                return "OSBAL", "Polis only", like_hits, True

    # R5 & R6: SLIPDB
    if slipdb_rows and lookup_slipdb:
        if slip_vals:
            hits = _exact_match(slip_vals, lookup_slipdb[0])
            if hits:
                res_idx, ok = _resolve_facode([slipdb_rows[i] for i in hits], SLIPDB_FACODE_COL, facode_osbal_idx, osbal_rows)
                if ok:
                    curr_res = _strict_filter(res_idx)
                    if curr_res:
                        return "SLIPDB", "Slip only", curr_res, True
        if polis_vals:
            hits = _exact_match(polis_vals, lookup_slipdb[1])
            if hits:
                res_idx, ok = _resolve_facode([slipdb_rows[i] for i in hits], SLIPDB_FACODE_COL, facode_osbal_idx, osbal_rows)
                if ok:
                    curr_res = _strict_filter(res_idx)
                    if curr_res:
                        return "SLIPDB", "Polis only", curr_res, True

    # R7 & R8: FACUL
    if facul_rows and lookup_facul:
        if slip_vals:
            hits = _exact_match(slip_vals, lookup_facul[0])
            if hits:
                res_idx, ok = _resolve_facode([facul_rows[i] for i in hits], FACUL_FACODE_COL, facode_osbal_idx, osbal_rows)
                if ok:
                    curr_res = _strict_filter(res_idx)
                    if curr_res:
                        return "FACUL", "Slip only", curr_res, True
        if polis_vals:
            hits = _exact_match(polis_vals, lookup_facul[1])
            if hits:
                res_idx, ok = _resolve_facode([facul_rows[i] for i in hits], FACUL_FACODE_COL, facode_osbal_idx, osbal_rows)
                if ok:
                    curr_res = _strict_filter(res_idx)
                    if curr_res:
                        return "FACUL", "Polis only", curr_res, True

    # R9: OSBAL Insured fallback (Urutan 3: Insured)
    ins_vals = _collect_clean_values(suspend_row, insured_sus_cols, is_slip=False)
    if ins_vals:
        curr_hits = _strict_filter(_exact_match(ins_vals, lookup_osbal[2]))
        if curr_hits:
            return "OSBAL", "Insured only", curr_hits, True

    return None

def _aggregate_ccos(osbal_rows: list, all_fac_osbal_rows: list = None) -> dict:
    fac_codes = {str(r.get("CCOS_REF_CODE", "")).strip() for r in osbal_rows if str(r.get("CCOS_REF_CODE", "")).strip()}
    multi_fac = len(fac_codes) > 1
    bal_due_rows = all_fac_osbal_rows if (all_fac_osbal_rows is not None and not multi_fac) else osbal_rows

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
                s = str(v).strip()
                if not s:
                    continue
                if ',' in s and '.' in s:
                    s = s.replace('.', '').replace(',', '.') if s.rfind(',') > s.rfind('.') else s.replace(',', '')
                elif ',' in s:
                    s = s.replace(',', '.')
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

def _facode_label(osbal_rows: list, source: str) -> str:
    if not osbal_rows or not osbal_rows[0]:
        return ""
    col = OSBAL_FACODE_COL if source == "OSBAL" else (
          SLIPDB_FACODE_COL if source == "SLIPDB" else FACUL_FACODE_COL)
    codes = {str(r.get(col, "")).strip() for r in osbal_rows if str(r.get(col, "")).strip()}
    return "facode lebih dari 1" if len(codes) > 1 else next(iter(codes), "")

def _clean_base_match(scenario: str) -> str:
    s = str(scenario or "").strip()
    for to_remove in [" (Beda Currency)", " (Beda Periode)", " (> 1 Fac)"]:
        s = s.replace(to_remove, "")
    s = s.strip()
    if " only + " in s:
        s = s.replace(" only + ", " + ")
    return s

def _format_final_scenario(
    base_scenario: str, source: str,
    is_beda_curr: bool, is_beda_per: bool, num_fac: int
) -> str:
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
    osbal_count: int = 0, resolved: bool = True,
    suspend_count: int = 1, all_fac_osbal_rows: list = None
) -> dict:
    has_match = bool(source and osbal_rows and osbal_rows[0])
    has_osbal = has_match and (source in ("OSBAL", "SLIPDB", "FACUL") or resolved)

    ccos = (
        _aggregate_ccos(osbal_rows, all_fac_osbal_rows=all_fac_osbal_rows)
        if has_osbal
        else {"CCOS_DOC_NO": "", "CCOS_REF_CODE": "", "CCOS_OR_BAL": np.nan, "CCOS_BAL_DUE": np.nan}
    )
    if "Beda Currency" in str(scenario) or "Beda Periode" in str(scenario):
        ccos["CCOS_OR_BAL"] = np.nan
        ccos["CCOS_BAL_DUE"] = np.nan

    final_scenario = "Unmatching" if not source or scenario in ("Unmatching", "UNMATCHED") else scenario

    return {
        "CCOS_DOC_NO":            ccos["CCOS_DOC_NO"] if has_osbal else "",
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
        "POLIS":                  suspend_row.get("POLIS_ORI") or suspend_row.get("POLIS") or suspend_row.get("polis_ori") or "",
        "POLICY_CLEAN_1":         suspend_row.get("POLIS_CLEAN_1") or suspend_row.get("POLICY_CLEAN_1") or suspend_row.get("clean polis 1") or "",
        "SERTIF_CLEAN_1":         "", # Tidak menggunakan sertifikat
        "POLICY_CLEAN_2":         suspend_row.get("POLIS_CLEAN_2") or suspend_row.get("POLICY_CLEAN_2") or "",
        "SERTIF_CLEAN_2":         "",
        "POLICY_CLEAN_3":         suspend_row.get("POLIS_CLEAN_3") or suspend_row.get("POLICY_CLEAN_3") or "",
        "SERTIF_CLEAN_3":         "",
        "POLICY_CLEAN_4":         suspend_row.get("POLIS_CLEAN_4") or suspend_row.get("POLICY_CLEAN_4") or "",
        "POLICY_CLEAN_5":         suspend_row.get("POLIS_CLEAN_5") or suspend_row.get("POLICY_CLEAN_5") or "",
        "SLIP_NO":                suspend_row.get("SLIP_NO_ORI") or suspend_row.get("SLIP_NO") or suspend_row.get("slip_ori") or "",
        "SLIP_NO_CLN":            suspend_row.get("SLIP_CLEAN_1") or suspend_row.get("SLIP_NO_CLN") or suspend_row.get("clean slip 1") or "",
        "DESC 1":                 suspend_row.get("DESC 1",            ""),
        "DESC 2":                 suspend_row.get("DESC 2",            ""),
        "DESC 3":                 suspend_row.get("DESC 3",            ""),
        "DESC 4":                 suspend_row.get("DESC 4",            ""),
        "STATUS":                 suspend_row.get("STATUS",            ""),
        "REC_TYPE":               suspend_row.get("REC_TYPE",          ""),
        "SKENARIO":               final_scenario,
        "_OSBAL_ROW_COUNT":       osbal_count,
        "_SUSPEND_ROW_COUNT":     suspend_count,
    }

def _merge_suspend_rows(rows: list) -> dict:
    if len(rows) == 1:
        return rows[0]
    merged: dict = {}
    for col in _SUSPEND_SUM_COLS:
        total = 0.0
        for r in rows:
            v = r.get(col, 0)
            if pd.isna(v) or v == "":
                continue
            if isinstance(v, (int, float)):
                total += float(v)
            else:
                s = str(v).strip()
                if not s:
                    continue
                if ',' in s and '.' in s:
                    s = s.replace('.', '').replace(',', '.') if s.rfind(',') > s.rfind('.') else s.replace(',', '')
                elif ',' in s:
                    s = s.replace(',', '.')
                try:
                    total += float(s)
                except ValueError:
                    pass
        merged[col] = total
    for col in _SUSPEND_JOIN_COLS:
        seen, vals = set(), []
        for r in rows:
            v = str(r.get(col, "")).strip()
            if v and v not in seen:
                seen.add(v)
                vals.append(v)
        merged[col] = ", ".join(vals)
    return merged

def _compute_derived_cols(df: pd.DataFrame) -> pd.DataFrame:
    amount_ori_base = df.get("AMOUNT ORI", pd.Series([0]*len(df), index=df.index))
    if "_MERGED_AMOUNT_ORI" in df.columns:
        amount_ori_raw = df["_MERGED_AMOUNT_ORI"].combine_first(amount_ori_base)
    else:
        amount_ori_raw = amount_ori_base

    amount_ori_clean = amount_ori_raw.replace(r'^\s*$', np.nan, regex=True)
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
    sus_count_col   = pd.to_numeric(df["_SUSPEND_ROW_COUNT"] if "_SUSPEND_ROW_COUNT" in df.columns else pd.Series(1, index=df.index), errors="coerce").fillna(1)
    ccos_ref        = df.get("CCOS_REF_CODE", pd.Series("", index=df.index)).fillna("").astype(str)
    skenario_col    = df.get("SKENARIO",      pd.Series("", index=df.index)).fillna("")

    df["DIFERENCE"] = amount_neg - balance_due
    accumulated = (osbal_count_col > 1) | (sus_count_col > 1)
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
        print(f"  {label}: {path} ({'tidak ada, dilewati' if optional else 'TIDAK DITEMUKAN'})", flush=True)
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
# PIPELINE UTAMA V1
# =============================================================================

def run() -> None:
    t_start = time.perf_counter()
    print("\n" + "=" * 60, flush=True)
    print("  PRODUCTION SCRIPT — SUSPEND MATCHING (KBRU) — V1", flush=True)
    print("  Urutan Matching: SLIP -> POLIS -> INSURED", flush=True)
    print("  Basic Rules: Tanpa Sertifikat, Tanpa Periode 2 / Lineslip Offset", flush=True)
    print("=" * 60, flush=True)

    # [1] Load data
    print("\n[1/5] Loading data ...", flush=True)
    suspend_rows, suspend_cols = _load_excel(SUSPEND_FILE, "SUSPEND")
    osbal_rows,   osbal_cols   = _load_excel(OSBAL_FILE,   "OSBAL")
    facul_rows,   facul_cols   = _load_excel(FACUL_FILE,   "FACUL")
    slipdb_rows,  slipdb_cols  = _load_excel(SLIPDB_FILE,  "SLIPDB", optional=True)

    # Pre-parse tanggal OSBAL FAC_COM_DATE
    print("  Pre-parsing OSBAL FAC_COM_DATE ...", flush=True)
    for r in osbal_rows:
        raw = r.get(OSBAL_DATE_COL, "")
        r["_com_date_parsed"] = pd.to_datetime(raw, errors="coerce") \
            if raw and not (isinstance(raw, float) and pd.isna(raw)) else None

    # Pre-parse tanggal Suspend RECEIPT DATE
    print("  Pre-parsing Suspend RECEIPT DATE ...", flush=True)
    for r in suspend_rows:
        raw = r.get(SUSPEND_DATE_COL, "")
        r["_sus_date_parsed"] = pd.to_datetime(raw, errors="coerce") \
            if raw and not (isinstance(raw, float) and pd.isna(raw)) else None

    # [2] Deteksi kolom bersih
    print("\n[2/5] Detecting clean columns ...", flush=True)
    polis_sus   = _get_clean_cols(suspend_cols, "clean polis")
    slip_sus    = _get_clean_cols(suspend_cols, "clean slip")
    insured_sus = _get_clean_cols(suspend_cols, "clean insured")

    for c in ["POLIS_ORI", "POLIS"]:
        if c in suspend_cols and c not in polis_sus:
            polis_sus.append(c)
    for c in ["SLIP_NO_ORI", "SLIP_NO"]:
        if c in suspend_cols and c not in slip_sus:
            slip_sus.append(c)
    if "INSURED_ORI" in suspend_cols and "INSURED_ORI" not in insured_sus:
        insured_sus.append("INSURED_ORI")

    polis_osbal   = _get_clean_cols(osbal_cols, "clean polis")
    slip_osbal    = _get_clean_cols(osbal_cols, "clean slip")
    insured_osbal = _get_clean_cols(osbal_cols, "clean insured")

    for c in ["CLSDT_POLICY_NO", "FAC_POLICY_NO_ORI", "FAC_SLIP_ORI"]:
        if c in osbal_cols and c not in polis_osbal:
            polis_osbal.append(c)
    for c in ["CLSDT_SLIP_NO", "FAC_SLIP_ORI"]:
        if c in osbal_cols and c not in slip_osbal:
            slip_osbal.append(c)

    polis_facul   = _get_clean_cols(facul_cols, "clean polis")
    slip_facul    = _get_clean_cols(facul_cols, "clean slip")
    insured_facul = _get_clean_cols(facul_cols, "clean insured")

    for c in ["FAC_POLICY_NO"]:
        if c in facul_cols and c not in polis_facul:
            polis_facul.append(c)
    for c in ["FAC_SLIP"]:
        if c in facul_cols and c not in slip_facul:
            slip_facul.append(c)

    polis_slipdb   = _get_clean_cols(slipdb_cols, "clean polis")
    slip_slipdb    = _get_clean_cols(slipdb_cols, "clean slip")
    insured_slipdb = _get_clean_cols(slipdb_cols, "clean insured")

    # [3] Bangun lookup index
    print("\n[3/5] Building lookup indexes ...", flush=True)
    lookup_osbal  = _build_lookup(osbal_rows,  polis_osbal,  slip_osbal,  insured_osbal,  OSBAL_FACODE_COL,  label="OSBAL")
    lookup_facul  = _build_lookup(facul_rows,  polis_facul,  slip_facul,  insured_facul,  FACUL_FACODE_COL,  label="FACUL")
    lookup_slipdb = _build_lookup(slipdb_rows, polis_slipdb, slip_slipdb, insured_slipdb, SLIPDB_FACODE_COL, label="SLIPDB")

    facode_osbal_idx = lookup_osbal[6]

    # [4] Matching per baris
    total = len(suspend_rows)
    print(f"\n[4/5] Matching {total:,} suspend rows ...", flush=True)
    raw_results = []

    for n, sus in enumerate(suspend_rows, 1):
        if n % 500 == 0 or n == total:
            print(f"    Memproses baris {n:,} / {total:,} ...", flush=True)

        if str(sus.get("STATUS", "")).strip().upper() == "MATCHING":
            continue

        effective_date = sus.get("_sus_date_parsed")
        sus_curr = _normalize(sus.get(SUSPEND_CURR_COL, ""))

        matched_osbal_idx = set()
        source = None
        scenario = ""
        resolved = False

        # -------------------------------------------------------------
        # FASE 1: PENCARIAN BERDASARKAN SLIP (Prioritas 1)
        # -------------------------------------------------------------
        slip_vals = _collect_clean_values(sus, slip_sus, is_slip=True)
        if slip_vals:
            # 1. OSBAL Exact
            osb_hits = _exact_match(slip_vals, lookup_osbal[0])
            if osb_hits:
                matched_osbal_idx.update(osb_hits)
                source = "OSBAL"
                scenario = "Slip only"

            # 2. SLIPDB Exact
            if slipdb_rows and not matched_osbal_idx:
                sld_hits = _exact_match(slip_vals, lookup_slipdb[0])
                if sld_hits:
                    res_idx, ok = _resolve_facode([slipdb_rows[i] for i in sld_hits], SLIPDB_FACODE_COL, facode_osbal_idx, osbal_rows)
                    if ok:
                        matched_osbal_idx.update(res_idx)
                        source = "SLIPDB"
                        scenario = "Slip only"
                        resolved = True

            # 3. FACUL Exact
            if facul_rows and not matched_osbal_idx:
                fac_hits = _exact_match(slip_vals, lookup_facul[0])
                if fac_hits:
                    res_idx, ok = _resolve_facode([facul_rows[i] for i in fac_hits], FACUL_FACODE_COL, facode_osbal_idx, osbal_rows)
                    if ok:
                        matched_osbal_idx.update(res_idx)
                        source = "FACUL"
                        scenario = "Slip only"
                        resolved = True

            # 4. OSBAL LIKE
            if not matched_osbal_idx and lookup_osbal[3] and lookup_osbal[3][0]:
                like_hits = _like_match(slip_vals, lookup_osbal[3][0], osbal_rows, slip_osbal)
                if like_hits:
                    matched_osbal_idx.update(like_hits)
                    source = "OSBAL"
                    scenario = "Slip only"

        # -------------------------------------------------------------
        # FASE 2: PENCARIAN BERDASARKAN POLIS (Prioritas 2)
        # -------------------------------------------------------------
        if not matched_osbal_idx:
            polis_vals = _collect_clean_values(sus, polis_sus, is_polis=True)
            if polis_vals:
                # 1. OSBAL Exact
                osb_hits = _exact_match(polis_vals, lookup_osbal[1])
                if osb_hits:
                    matched_osbal_idx.update(osb_hits)
                    source = "OSBAL"
                    scenario = "Polis only"

                # 2. SLIPDB Exact
                if slipdb_rows and not matched_osbal_idx:
                    sld_hits = _exact_match(polis_vals, lookup_slipdb[1])
                    if sld_hits:
                        res_idx, ok = _resolve_facode([slipdb_rows[i] for i in sld_hits], SLIPDB_FACODE_COL, facode_osbal_idx, osbal_rows)
                        if ok:
                            matched_osbal_idx.update(res_idx)
                            source = "SLIPDB"
                            scenario = "Polis only"
                            resolved = True

                # 3. FACUL Exact
                if facul_rows and not matched_osbal_idx:
                    fac_hits = _exact_match(polis_vals, lookup_facul[1])
                    if fac_hits:
                        res_idx, ok = _resolve_facode([facul_rows[i] for i in fac_hits], FACUL_FACODE_COL, facode_osbal_idx, osbal_rows)
                        if ok:
                            matched_osbal_idx.update(res_idx)
                            source = "FACUL"
                            scenario = "Polis only"
                            resolved = True

                # 4. OSBAL LIKE
                if not matched_osbal_idx and lookup_osbal[4] and lookup_osbal[4][0]:
                    like_hits = _like_match(polis_vals, lookup_osbal[4][0], osbal_rows, polis_osbal)
                    if like_hits:
                        matched_osbal_idx.update(like_hits)
                        source = "OSBAL"
                        scenario = "Polis only"

        # -------------------------------------------------------------
        # FASE 3: PENCARIAN BERDASARKAN INSURED (Prioritas 3)
        # -------------------------------------------------------------
        if not matched_osbal_idx:
            ins_vals = _collect_clean_values(sus, insured_sus, is_slip=False)
            if ins_vals:
                # 1. OSBAL Exact
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
                            resolved = True
                elif facul_rows and not matched_osbal_idx:
                    fac_hits = _exact_match(ins_vals, lookup_facul[2])
                    if fac_hits:
                        res_idx, ok = _resolve_facode([facul_rows[i] for i in fac_hits], FACUL_FACODE_COL, facode_osbal_idx, osbal_rows)
                        if ok:
                            matched_osbal_idx.update(res_idx)
                            source = "FACUL"
                            scenario = "Insured only"
                            resolved = True

                # 4. OSBAL LIKE
                if not matched_osbal_idx and lookup_osbal[5] and lookup_osbal[5][0]:
                    like_hits = _like_match(ins_vals, lookup_osbal[5][0], osbal_rows, insured_osbal)
                    if like_hits:
                        matched_osbal_idx.update(like_hits)
                        source = "OSBAL"
                        scenario = "Insured only"

        # -------------------------------------------------------------
        # FASE 4: NARROWING WAJIB (Currency & Periode) + Rematch Strict
        # -------------------------------------------------------------
        is_beda_curr = False
        is_beda_per = False

        if matched_osbal_idx:
            matched_list = list(matched_osbal_idx)

            # Currency check
            if sus_curr:
                filtered_curr = [i for i in matched_list if _normalize(osbal_rows[i].get(OSBAL_CURR_COL, "")) == sus_curr]
                if not filtered_curr:
                    is_beda_curr = True
                elif len(filtered_curr) < len(matched_list):
                    matched_list = filtered_curr

            # Periode check (Standard: sus_date >= com_date, no lineslip offset)
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

            if is_beda_curr or is_beda_per:
                rm = _rematch_strict(
                    sus, osbal_rows, lookup_osbal,
                    slip_sus, polis_sus, insured_sus,
                    slip_osbal, polis_osbal, insured_osbal,
                    facode_osbal_idx,
                    slipdb_rows, lookup_slipdb,
                    facul_rows, lookup_facul,
                    effective_sus_date=effective_date
                )
                if rm:
                    src, scen, hits, _ = rm
                    source = src
                    scenario = scen
                    matched_list = hits
                    is_beda_curr = False
                    is_beda_per = False

            matched_osbal_idx = set(matched_list)

        # -------------------------------------------------------------
        # FASE 5: NARROWING UTAMA (Non-Destructive)
        # Urutan: Slip -> Polis -> Insured (Tanpa Sertifikat)
        # -------------------------------------------------------------
        if len(matched_osbal_idx) > 1:
            matched_list = list(matched_osbal_idx)

            if "Slip" in scenario:
                # Step 1: Polis Confirmation
                polis_vals = _collect_clean_values(sus, polis_sus, is_polis=True)
                if polis_vals:
                    confirmed = [idx for idx in matched_list if _ref_has_value(osbal_rows[idx], polis_osbal, "FAC_POLICY_NO_ORI", polis_vals, is_polis=True)]
                    if confirmed:
                        matched_list = confirmed
                        scenario = "Slip + Polis"
                # Step 2: Insured Confirmation
                ins_vals = _collect_clean_values(sus, insured_sus, is_slip=False)
                if len(matched_list) > 1 and ins_vals:
                    confirmed = [idx for idx in matched_list if _ref_has_value(osbal_rows[idx], insured_osbal, "FAC_INSURED_ORI", ins_vals)]
                    if confirmed:
                        matched_list = confirmed
                        scenario = scenario + " + Insured"

            elif "Polis" in scenario:
                # Step 1: Slip Confirmation
                slip_vals = _collect_clean_values(sus, slip_sus, is_slip=True)
                if slip_vals:
                    confirmed = [idx for idx in matched_list if _ref_has_value(osbal_rows[idx], slip_osbal, "FAC_SLIP_ORI", slip_vals, is_slip=True)]
                    if confirmed:
                        matched_list = confirmed
                        scenario = "Polis + Slip"
                # Step 2: Insured Confirmation
                ins_vals = _collect_clean_values(sus, insured_sus, is_slip=False)
                if len(matched_list) > 1 and ins_vals:
                    confirmed = [idx for idx in matched_list if _ref_has_value(osbal_rows[idx], insured_osbal, "FAC_INSURED_ORI", ins_vals)]
                    if confirmed:
                        matched_list = confirmed
                        scenario = scenario + " + Insured"

            elif "Insured" in scenario:
                # Step 1: Slip Confirmation
                slip_vals = _collect_clean_values(sus, slip_sus, is_slip=True)
                if slip_vals:
                    confirmed = [idx for idx in matched_list if _ref_has_value(osbal_rows[idx], slip_osbal, "FAC_SLIP_ORI", slip_vals, is_slip=True)]
                    if confirmed:
                        matched_list = confirmed
                        scenario = "Insured + Slip"
                # Step 2: Polis Confirmation
                polis_vals = _collect_clean_values(sus, polis_sus, is_polis=True)
                if len(matched_list) > 1 and polis_vals:
                    confirmed = [idx for idx in matched_list if _ref_has_value(osbal_rows[idx], polis_osbal, "FAC_POLICY_NO_ORI", polis_vals, is_polis=True)]
                    if confirmed:
                        matched_list = confirmed
                        scenario = scenario + " + Polis"

            matched_osbal_idx = set(matched_list)

        # -------------------------------------------------------------
        # FASE 6: RESOLUSI AKHIR & LABEL
        # -------------------------------------------------------------
        fac_codes = set()
        if matched_osbal_idx:
            for idx in matched_osbal_idx:
                fc = str(osbal_rows[idx].get("CCOS_REF_CODE", "")).strip()
                if fc:
                    fac_codes.add(fc)

        final_scenario = _format_final_scenario(
            scenario, source, is_beda_curr, is_beda_per, len(fac_codes)
        )

        raw_results.append({
            "suspend":       sus,
            "source":        source if final_scenario != "Unmatching" else None,
            "scenario":      final_scenario,
            "osbal_idx":     list(matched_osbal_idx) if final_scenario != "Unmatching" else [],
            "osbal_count":   len(matched_osbal_idx) if final_scenario != "Unmatching" else 0,
            "resolved":      resolved,
            "fac_codes":     fac_codes,
            "suspend_count": 1,
        })

    # [5] Build output DataFrame (V1: Row Asli)
    print(f"\n[5/5] Building output DataFrame ({len(raw_results):,} rows) ...", flush=True)
    output_rows = []
    for r in raw_results:
        ref_osbal = [osbal_rows[i] for i in r["osbal_idx"]] if r["source"] else [{}]
        all_fac_osbal_rows = None
        osbal_cnt = r["osbal_count"]

        if r["source"] and len(r["fac_codes"]) == 1:
            fac = next(iter(r["fac_codes"]))
            all_fac_idx = facode_osbal_idx.get(fac, [])
            if len(all_fac_idx) > 0:
                all_fac_osbal_rows = [osbal_rows[i] for i in all_fac_idx]
                osbal_cnt = len(all_fac_idx)

        out = _build_output_row(
            r["suspend"], source=r["source"], scenario=r["scenario"],
            osbal_rows=ref_osbal, osbal_count=osbal_cnt,
            resolved=r["resolved"], suspend_count=1,
            all_fac_osbal_rows=all_fac_osbal_rows,
        )
        output_rows.append(out)

    df = pd.DataFrame(output_rows)

    # Format numeric
    for col in ["AMOUNT ORI", "CCOS_OR_BAL", "CCOS_BAL_DUE"]:
        s = df[col]
        if s.dtype == object:
            s = s.astype(str).str.strip().str.replace(r'\s+', '', regex=True)
            has_comma = s.str.contains(',', na=False)
            s_eu = s.str.replace('.', '', regex=False).str.replace(',', '.', regex=False)
            s = s_eu.where(has_comma, s)
        df[col] = pd.to_numeric(s, errors="coerce")

    df = _compute_derived_cols(df)
    df.loc[df["SKENARIO"].isin(["UNMATCHED", "Unmatching"]), "FLAG_PROD"] = "Unmatching"
    df.loc[df["SKENARIO"] == "UNMATCHED", "SKENARIO"] = "Unmatching"

    # Pastikan urutan 38 kolom persis
    for col in FINAL_COLUMNS:
        if col not in df.columns:
            df[col] = ""
    df = df[FINAL_COLUMNS]

    # Simpan ke Excel
    print(f"  Saving Excel to: {OUTPUT_FILE} ...", flush=True)
    for c in df.select_dtypes(include=['object', 'string']).columns:
        df[c] = df[c].astype(str).str.replace(r'[\x00-\x08\x0B\x0C\x0E-\x1F]', '', regex=True)
    df.to_excel(OUTPUT_FILE, index=False)

    try:
        from excel_styler import apply_purple_column_style
        apply_purple_column_style(OUTPUT_FILE, "matching")
    except Exception as e:
        print(f"  [WARN] Styling gagal: {e}")

    # Export ke PostgreSQL per-cedant
    load_dotenv()
    print("  Exporting ke PostgreSQL (SUSPENSE_DATA_SUSPENSE_V1 & SUSPENSE_CLEAN_2026) ...", flush=True)
    db_user = os.environ.get("DB_USER", "postgres")
    db_pass = os.environ.get("DB_PASS", "postgres")
    db_host = os.environ.get("DB_HOST", "localhost")
    db_port = os.environ.get("DB_PORT", "5432")
    db_name = os.environ.get("DB_NAME", "postgres")
    conn_str = f"postgresql+psycopg2://{db_user}:{db_pass}@{db_host}:{db_port}/{db_name}"

    try:
        df_db = df.copy()
        df_db["RECEIPT DATE"] = pd.to_datetime(df_db["RECEIPT DATE"], errors="coerce").dt.date
        df_db["RECEIPT DATE"] = df_db["RECEIPT DATE"].where(df_db["RECEIPT DATE"].notnull(), None)

        numeric_cols = ["AMOUNT ORI", "AMOUNT_ORI_MIN1", "AMOUNT PAY", "CCOS_OR_BAL", "CCOS_BAL_DUE", "DIFERENCE"]
        for c in numeric_cols:
            df_db[c] = pd.to_numeric(df_db[c], errors="coerce")

        text_cols = [c for c in FINAL_COLUMNS if c not in numeric_cols and c != "RECEIPT DATE"]
        for c in text_cols:
            df_db[c] = df_db[c].astype(str).str.strip()
            df_db[c] = df_db[c].replace({"nan": None, "None": None, "": None})

        engine = create_engine(conn_str)
        with engine.begin() as con:
            con.execute(text('DELETE FROM "SUSPENSE_DATA_SUSPENSE_V1" WHERE "CEDANT NAME" = :name OR "CEDANT SHRT NAME" = :shrt;'),
                        {"name": KBRU_CEDANT_NAME, "shrt": KBRU_SHRT_NAME})
            con.execute(text('DELETE FROM "SUSPENSE_CLEAN_2026" WHERE "CEDANT NAME" = :name OR "CEDANT SHRT NAME" = :shrt;'),
                        {"name": KBRU_CEDANT_NAME, "shrt": KBRU_SHRT_NAME})

        df_db.to_sql("SUSPENSE_DATA_SUSPENSE_V1", con=engine, if_exists="append", index=False)
        df_db.to_sql("SUSPENSE_CLEAN_2026", con=engine, if_exists="append", index=False)
        print(f"      -> {len(df_db):,} rows berhasil di-export ke 'SUSPENSE_DATA_SUSPENSE_V1' & 'SUSPENSE_CLEAN_2026'")
    except Exception as e:
        print(f"  [WARN] PostgreSQL export gagal: {e}")

    elapsed = time.perf_counter() - t_start
    print(f"\n{'=' * 60}")
    print(f"  Selesai Suspend Matching V1 dalam {elapsed:.1f}s ({elapsed/60:.1f} menit)")
    print(f"  Rows  : {len(df):,}  |  Cols: {len(df.columns)}  |  File: {OUTPUT_FILE}")
    print(f"\n  FLAG_PROD:")
    for flag, count in df["FLAG_PROD"].value_counts().items():
        print(f"    {flag:<45}: {count:,}")
    print(f"\n  SKENARIO:")
    for sce, count in df["SKENARIO"].value_counts().items():
        print(f"    {sce:<50}: {count:,}")
    print("=" * 60)

if __name__ == "__main__":
    run()