"""
================================================================================
 DEBUG MATCHING LIPPO (REVISED PIPELINE)
================================================================================
 Script debug independen untuk mengevaluasi pembaruan logika pencocokan Lippo:
 - Mengakomodasi struktur kolom POLIS di Suspend Lippo:
     <NOMOR POLIS> - <NOMOR SLIP>
     Bagian depan  = Nomor Polis (clean digits)
     Bagian belakang = Nomor Slip / Cession No (clean digits)
 - Memprioritaskan Joint Match (Polis + Slip) secara langsung.
 - Memperbaiki penanganan narrowing slip agar menggunakan nomor slip sejati.
 - Tidak memodifikasi script production asli.
 - Menghasilkan laporan komparasi mendalam terhadap baseline (final_output_v1.xlsx).
================================================================================
"""

import sys
sys.modules['numexpr'] = None
sys.modules['bottleneck'] = None

import os
import re
import time
import functools
import numpy as np
import pandas as pd
from openpyxl import load_workbook

# =============================================================================
# KONFIGURASI PATH
# =============================================================================
SCRIPT_DIR    = os.path.dirname(os.path.abspath(__file__))
DATA_DIR      = os.path.join(SCRIPT_DIR, "..", "data")
OUTPUT_DIR    = SCRIPT_DIR

SUSPEND_FILE  = os.path.join(DATA_DIR, "lippo_output_suspense_V3.xlsx")
OSBAL_FILE    = os.path.join(DATA_DIR, "lippo_output_osbal_V2.xlsx")
FACUL_FILE    = os.path.join(DATA_DIR, "lippo_output_facul_V3.xlsx")
BASELINE_FILE = os.path.join(DATA_DIR, "final_output_v1.xlsx")
OUTPUT_FILE   = os.path.join(OUTPUT_DIR, "output_debug_lippo.xlsx")

OSBAL_FACODE_COL = "CCOS_REF_CODE"
FACUL_FACODE_COL = "FAC_CODE"
SUSPEND_CURR_COL = "CURR ORI"
OSBAL_CURR_COL   = "CCOS_CURR"
SUSPEND_DATE_COL = "RECEIPT DATE"
OSBAL_DATE_COL   = "FAC_COM_DATE"

_EXCLUDED_REF_MARKERS = ["HUTANG PIUTANG", "DATA SUSPENSE"]
_MIN_TOKEN_LEN        = 3
_TOKEN_RE             = re.compile(r"[^A-Z0-9]+")

_SUSPEND_JOIN_COLS = [
    "RECEIPT NO", "CREDIT NOTES", "DETAIL RINCIAN NO", "RECEIPT DATE",
    "CEDANT NAME", "CEDANT SHRT NAME",
    "INSURED", "INSURED_CLEAN_1", "INSURED_CLEAN_2",
    "CURR ORI", "CURR PAY", "AMOUNT PAY",
    "POLIS", "POLICY_CLEAN_1", "POLICY_CLEAN_2",
    "SLIP NO", "SLIP_NO_CLEAN",
    "DESC 1", "DESC 2", "DESC 3", "DESC 4",
    "STATUS", "REC_TYPE",
]
_SUSPEND_SUM_COLS = ["AMOUNT ORI"]

FINAL_COLUMNS = [
    "CCOS_DOC_NO", "CCOS_REF_CODE",
    "RECEIPT NO", "CREDIT NOTES", "DETAIL RINCIAN NO", "RECEIPT DATE",
    "CEDANT NAME", "CEDANT SHRT NAME",
    "INSURED_ORI", "INSURED_CLEAN_1", "INSURED_CLEAN_2",
    "CURR ORI", "AMOUNT ORI", "AMOUNT_ORI_MIN1",
    "CURR PAY", "AMOUNT PAY",
    "CCOS_OR_BAL", "CCOS_BAL_DUE", "DIFERENCE",
    "FLAG_PROD",
    "POLIS",
    "POLICY_CLEAN_1", "SERTIF_CLEAN_1",
    "POLICY_CLEAN_2", "SERTIF_CLEAN_2",
    "POLICY_CLEAN_3", "SERTIF_CLEAN_3",
    "POLICY_CLEAN_4", "POLICY_CLEAN_5",
    "SLIP_NO", "SLIP_NO_CLEAN",
    "DESC 1", "DESC 2", "DESC 3", "DESC 4",
    "STATUS", "REC_TYPE",
    "SKENARIO",
]


# =============================================================================
# HELPER FUNCTIONS
# =============================================================================
@functools.lru_cache(maxsize=131072)
def _norm_cached(s: str) -> str:
    text = s.strip().upper().replace("S/D", "SD")
    if text.endswith(".0"):
        text = text[:-2]
    return re.sub(r"\s+", " ", text) if "  " in text else text

def _normalize(value) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    return _norm_cached(str(value))

def _clean_code(val) -> str:
    s = _normalize(val)
    if not s or s == "NAN":
        return ""
    m = re.search(r"\b\d{6,}\b", s)
    return m.group(0) if m else s

def _clean_cert_str(val) -> str:
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return ""
    s = str(val).strip()
    return re.sub(r"\.0+$", "", s)

def _expand_sertif_range(val: str) -> list:
    if not val:
        return []
    s = re.sub(r"\.0+$", "", str(val).strip())
    clean = re.sub(r"\s*SD\s*|\s*S/D\s*", "-", s, flags=re.IGNORECASE)
    clean = re.sub(r"\s+", "", clean)
    if clean.isdigit() and 1 <= len(clean) <= 6:
        return [clean.zfill(6)]
    m = re.match(r"^(\d{1,6})-(\d{1,6})$", clean)
    if m:
        start, end = int(m.group(1)), int(m.group(2))
        if start <= end and (end - start) <= 5000:
            return [str(i).zfill(6) for i in range(start, end + 1)]
    return []

def _cert_in_range(query_cert: str, ref_cert_str: str) -> bool:
    q = _clean_cert_str(query_cert)
    if not q or not q.isdigit():
        return False
    q_num = int(q)
    s = _clean_cert_str(ref_cert_str)
    clean = re.sub(r"\s*SD\s*|\s*S/D\s*", "-", s, flags=re.IGNORECASE)
    clean = re.sub(r"\s+", "", clean)
    if clean.isdigit():
        return q_num == int(clean)
    m = re.match(r"^(\d{1,6})-(\d{1,6})$", clean)
    if m:
        start, end = int(m.group(1)), int(m.group(2))
        if start <= end:
            return start <= q_num <= end
        return q_num >= start or q_num <= end
    return False

def _tokenize(text: str) -> frozenset:
    if not text:
        return frozenset()
    return frozenset(t for t in _TOKEN_RE.split(text) if len(t) >= _MIN_TOKEN_LEN)

def _format_sertif(val) -> str:
    if pd.isna(val) or str(val).strip() == "":
        return ""
    if isinstance(val, (int, float)):
        return str(int(val)).zfill(6)
    s = str(val).strip()
    if s.endswith(".0"):
        s = s[:-2]
    return s.zfill(6) if s.isdigit() else s

def _is_excluded_ref_row(ref_row: dict, polis_cols: list, slip_cols: list) -> bool:
    for col in list(polis_cols) + list(slip_cols):
        val = _normalize(ref_row.get(col, ""))
        if val and any(m in val for m in _EXCLUDED_REF_MARKERS):
            return True
    return False

def _get_clean_cols(columns: list, prefix: str) -> list:
    alt_map = {
        "clean polis": ["clean polis", "policy_clean", "polis_cln"],
        "clean slip": ["clean slip", "slip_clean", "slip_no_cln"],
        "clean insured": ["clean insured", "insured_clean", "insured_cln", "insured_"],
        "clean sertif": ["clean sertif", "sertif_clean", "sertif_cln", "certificate", "clsdt_sertf"],
    }
    prefixes = alt_map.get(prefix.lower(), [prefix.lower()])
    res = []
    for c in columns:
        clow = str(c).lower()
        if any(clow.startswith(p) for p in prefixes):
            res.append(c)
    return res

def _load_excel(path: str, label: str) -> tuple:
    if not os.path.exists(path):
        print(f"  {label}: {path} (TIDAK DITEMUKAN)", flush=True)
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
            pass
    t = time.perf_counter()
    print(f"  {label}: membaca Excel {path} ...", flush=True)
    df = pd.read_excel(path)
    data = df.to_dict(orient="records")
    header = list(df.columns)
    print(f"  -> {len(data):,} rows, {len(header)} cols ({time.perf_counter()-t:.2f}s)", flush=True)
    try:
        pd.to_pickle((data, header), cache)
    except Exception:
        pass
    return data, header


# =============================================================================
# EXTRACTOR KHUSUS SUSPEND LIPPO
# =============================================================================
def extract_suspend_identifiers(row: dict) -> tuple[list[str], list[str], list[str], list[str]]:
    polis_candidates = []
    slip_candidates = []
    cert_candidates = []
    insured_candidates = []

    for p_col in ["POLIS_CLN_1", "POLICY_CLEAN_1", "POLIS"]:
        val = _clean_code(row.get(p_col))
        if val and val not in polis_candidates:
            polis_candidates.append(val)

    for s_col in ["SLIP_NO_CLN_1", "SLIP_CLEAN_1", "SLIP_NO", "SLIP NO"]:
        val = _clean_code(row.get(s_col))
        if val and val not in slip_candidates:
            slip_candidates.append(val)

    for c_col in ["CERTIFICATE", "SERTIF_CLEAN_1", "SERTIF", "CERTIFICATE_CLN_1"]:
        c_val = _clean_cert_str(row.get(c_col))
        if c_val and c_val != "NAN" and c_val not in cert_candidates:
            cert_candidates.append(c_val)

    for i in range(1, 6):
        for col_name in [f"INSURED_CLN_{i}", f"INSURED_CLEAN_{i}"]:
            v = _normalize(row.get(col_name))
            if v and v != "NAN" and v not in insured_candidates:
                insured_candidates.append(v)
    if not insured_candidates:
        raw_ins = _normalize(row.get("INSURED") or row.get("INSURED_ORI"))
        if raw_ins and raw_ins != "NAN":
            insured_candidates.append(raw_ins)

    return polis_candidates, slip_candidates, cert_candidates, insured_candidates


# =============================================================================
# INDEX BUILDERS
# =============================================================================
def _build_exact_index(rows: list, cols: list, excluded: set = None) -> dict:
    index = {}
    for i, row in enumerate(rows):
        if excluded and i in excluded:
            continue
        for col in cols:
            val = _clean_code(row.get(col, ""))
            if val:
                index.setdefault(val, []).append(i)
    return index

def _build_token_index(rows: list, clean_cols: list, excluded: set = None) -> tuple:
    token_index = {}
    for i, row in enumerate(rows):
        if excluded and i in excluded:
            continue
        tokens = frozenset()
        for c in clean_cols:
            t = _normalize(row.get(c, ""))
            if t:
                tokens = tokens | _tokenize(t)
        for tok in tokens:
            token_index.setdefault(tok, set()).add(i)
    return token_index

def _build_sertif_index(rows: list, sertif_cols: list, excluded: set = None) -> dict:
    index = {}
    for i, row in enumerate(rows):
        if excluded and i in excluded:
            continue
        for col in sertif_cols:
            val = _normalize(row.get(col, ""))
            for cert in _expand_sertif_range(val):
                index.setdefault(cert, []).append(i)
    return index


# =============================================================================
# MATCH ENGINE
# =============================================================================
def _exact_match(query_values: list, index: dict) -> set:
    hits = set()
    for v in query_values:
        clean_v = _clean_code(v)
        if clean_v and clean_v in index:
            hits.update(index[clean_v])
    return hits

def _like_match(query_values: list, token_index: dict, rows: list, ref_cols: list) -> list:
    valid_queries = [v for v in query_values if v and len(v) >= 2]
    if not valid_queries:
        return []
    query_tokens = frozenset()
    for v in valid_queries:
        query_tokens = query_tokens | _tokenize(v)
    candidates = set()
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

def _resolve_facode(matched_rows: list, facode_col: str, facode_osbal_idx: dict) -> set:
    osbal_indices = set()
    for row in matched_rows:
        fac = _normalize(row.get(facode_col, ""))
        if fac and fac in facode_osbal_idx:
            osbal_indices.update(facode_osbal_idx[fac])
    return osbal_indices

def _aggregate_ccos(rows: list, all_fac_osbal_rows: list = None) -> dict:
    if not rows or not rows[0]:
        return {"CCOS_DOC_NO": "", "CCOS_REF_CODE": "", "CCOS_OR_BAL": np.nan, "CCOS_BAL_DUE": np.nan}
    docs, refs = [], []
    for r in rows:
        d = str(r.get("CCOS_DOC_NO", "")).strip()
        rf = str(r.get(OSBAL_FACODE_COL, "")).strip()
        if d and d not in docs:
            docs.append(d)
        if rf and rf not in refs:
            refs.append(rf)
    calc_rows = all_fac_osbal_rows if all_fac_osbal_rows is not None else rows
    seen_indices = set()
    unique_calc_rows = []
    for r in calc_rows:
        r_id = id(r)
        if r_id not in seen_indices:
            seen_indices.add(r_id)
            unique_calc_rows.append(r)
    or_bal_sum  = sum(float(r.get("CCOS_OR_BAL", 0) or 0) for r in unique_calc_rows)
    bal_due_sum = sum(float(r.get("CCOS_BAL_DUE", 0) or 0) for r in unique_calc_rows)
    return {
        "CCOS_DOC_NO":   ", ".join(docs),
        "CCOS_REF_CODE": ", ".join(refs),
        "CCOS_OR_BAL":   or_bal_sum,
        "CCOS_BAL_DUE":  bal_due_sum,
    }

def _format_final_scenario(base_scenario: str, source: str, is_beda_curr: bool = False, is_beda_per: bool = False, num_fac: int = 1) -> str:
    base = base_scenario or "Unmatching"
    for to_remove in [" (Beda Currency)", " (Beda Periode)", " (> 1 Fac)"]:
        base = base.replace(to_remove, "")
    if is_beda_curr and is_beda_per:
        return f"{base} (Beda Currency) (Beda Periode)"
    elif is_beda_curr:
        return f"{base} (Beda Currency)"
    elif is_beda_per:
        return f"{base} (Beda Periode)"
    elif num_fac > 1:
        return f"{base} (> 1 Fac)"
    return base

def _merge_suspend_rows(rows: list) -> dict:
    if len(rows) == 1:
        return rows[0]
    merged = {}
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
    is_beda_currency     = skenario_col.str.contains("Beda Currency", na=False, regex=False)
    is_beda_periode      = skenario_col.str.contains("Beda Periode",  na=False, regex=False)
    is_gt1_fac           = (skenario_col.str.contains("> 1 Fac", na=False, regex=False) | ccos_ref.str.contains(",", na=False))
    is_new_entry_total   = ~is_unmatched & (amount_ori != 0) & (balance_due == 0)
    is_adj_total         = ~is_unmatched & is_equal
    is_adj_sebagian      = ~is_unmatched & ~is_equal & (amount_neg < balance_due)
    is_new_entry_sebagian= ~is_unmatched & ~is_equal & (amount_neg > balance_due)

    df["FLAG_PROD"] = np.select(
        [is_beda_currency, is_beda_periode, is_gt1_fac,
         is_new_entry_total,
         is_adj_total & ~accumulated, is_adj_total & accumulated,
         is_adj_sebagian & ~accumulated, is_adj_sebagian & accumulated,
         is_new_entry_sebagian],
        ["Beda Currency", "Beda Periode",
         "Matching >1 fac code",
         "New Entry total",
         "Adjustment total tanpa akumulasi", "Adjustment total dengan akumulasi",
         "Adjustment sebagian tanpa akumulasi", "Adjustment sebagian dengan akumulasi",
         "New Entry sebagian"],
        default="Unmatching",
    )

    for col in ["_OSBAL_ROW_COUNT", "_SUSPEND_ROW_COUNT", "_MERGED_AMOUNT_ORI"]:
        if col in df.columns:
            df.drop(columns=[col], inplace=True)
    return df

def _build_output_row(
    suspend_row: dict, source: str, scenario: str, osbal_rows: list,
    osbal_count: int = 0, resolved: bool = True,
    suspend_count: int = 1, all_fac_osbal_rows: list = None,
    clean_polis_val: str = "", clean_slip_val: str = "",
) -> dict:
    has_match = bool(source and osbal_rows and osbal_rows[0])
    has_osbal = has_match and (source in ("OSBAL", "FACUL") or resolved)

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
        "CCOS_DOC_NO":       ccos["CCOS_DOC_NO"] if has_osbal else "",
        "CCOS_REF_CODE":     ccos["CCOS_REF_CODE"],
        "RECEIPT NO":        suspend_row.get("RECEIPT NO",        ""),
        "CREDIT NOTES":      suspend_row.get("CREDIT NOTES",      ""),
        "DETAIL RINCIAN NO": suspend_row.get("DETAIL RINCIAN NO", ""),
        "RECEIPT DATE":      suspend_row.get("RECEIPT DATE",      ""),
        "CEDANT NAME":       suspend_row.get("CEDANT NAME",       ""),
        "CEDANT SHRT NAME":  suspend_row.get("CEDANT SHRT NAME",  ""),
        "INSURED_ORI":       suspend_row.get("INSURED") or suspend_row.get("insured_ori") or "",
        "INSURED_CLEAN_1":   suspend_row.get("INSURED_CLEAN_1") or suspend_row.get("clean insured 1") or suspend_row.get("INSURED_CLN_1") or "",
        "INSURED_CLEAN_2":   suspend_row.get("INSURED_CLEAN_2") or suspend_row.get("clean insured 2") or suspend_row.get("INSURED_CLN_2") or "",
        "CURR ORI":          suspend_row.get("CURR ORI",          ""),
        "AMOUNT ORI":        suspend_row.get("AMOUNT ORI",        ""),
        "CURR PAY":          suspend_row.get("CURR PAY",          ""),
        "AMOUNT PAY":        suspend_row.get("AMOUNT PAY",        ""),
        "CCOS_OR_BAL":       ccos["CCOS_OR_BAL"],
        "CCOS_BAL_DUE":      ccos["CCOS_BAL_DUE"],
        "POLIS":             suspend_row.get("POLIS") or "",
        "POLICY_CLEAN_1":    clean_polis_val or suspend_row.get("POLICY_CLEAN_1") or suspend_row.get("POLIS_CLN_1") or "",
        "SERTIF_CLEAN_1":    _format_sertif(suspend_row.get("SERTIF_CLEAN_1") or suspend_row.get("SERTIF") or suspend_row.get("CERTIFICATE") or ""),
        "POLICY_CLEAN_2":    suspend_row.get("POLICY_CLEAN_2") or suspend_row.get("POLIS_CLN_2") or "",
        "SERTIF_CLEAN_2":    _format_sertif(suspend_row.get("SERTIF_CLEAN_2") or suspend_row.get("CERTIFICATE_2") or ""),
        "POLICY_CLEAN_3":    suspend_row.get("POLICY_CLEAN_3") or suspend_row.get("POLIS_CLN_3") or "",
        "SERTIF_CLEAN_3":    _format_sertif(suspend_row.get("SERTIF_CLEAN_3") or suspend_row.get("CERTIFICATE_3") or ""),
        "POLICY_CLEAN_4":    suspend_row.get("POLICY_CLEAN_4") or suspend_row.get("POLIS_CLN_4") or "",
        "POLICY_CLEAN_5":    suspend_row.get("POLICY_CLEAN_5") or suspend_row.get("POLIS_CLN_5") or "",
        "SLIP_NO":           suspend_row.get("SLIP NO") or suspend_row.get("SLIP_NO") or "",
        "SLIP_NO_CLEAN":     clean_slip_val or suspend_row.get("SLIP_CLEAN_1") or suspend_row.get("SLIP_NO_CLEAN") or suspend_row.get("SLIP_NO_CLN_1") or "",
        "DESC 1":            suspend_row.get("DESC 1",            ""),
        "DESC 2":            suspend_row.get("DESC 2",            ""),
        "DESC 3":            suspend_row.get("DESC 3",            ""),
        "DESC 4":            suspend_row.get("DESC 4",            ""),
        "STATUS":            suspend_row.get("STATUS",            ""),
        "REC_TYPE":          suspend_row.get("REC_TYPE",          ""),
        "SKENARIO":          final_scenario,
        "_OSBAL_ROW_COUNT":  osbal_count,
        "_SUSPEND_ROW_COUNT": suspend_count,
    }


# =============================================================================
# RUN ENGINE
# =============================================================================
def run_debug():
    t_start = time.perf_counter()
    print("=" * 75)
    print(" REVISED DEBUG MATCHING PIPELINE — LIPPO PRODUCTION")
    print(" Evaluasi Pemisahan POLIS (depan) dan SLIP (belakang)")
    print("=" * 75)

    # 1. Load Data
    print("\n[1/5] Memuat dataset ...")
    suspend_rows, suspend_cols = _load_excel(SUSPEND_FILE, "SUSPENSE")
    osbal_rows,   osbal_cols   = _load_excel(OSBAL_FILE,   "OSBAL")
    facul_rows,   facul_cols   = _load_excel(FACUL_FILE,   "FACUL")

    # Date parsing
    for r in osbal_rows:
        raw = r.get(OSBAL_DATE_COL, "")
        r["_com_date_parsed"] = pd.to_datetime(raw, errors="coerce") if raw and not (isinstance(raw, float) and pd.isna(raw)) else None

    for r in suspend_rows:
        raw = r.get(SUSPEND_DATE_COL, "")
        r["_sus_date_parsed"] = pd.to_datetime(raw, errors="coerce") if raw and not (isinstance(raw, float) and pd.isna(raw)) else None

    # 2. Build Indexes
    print("\n[2/5] Membangun indeks pencarian ...")
    polis_osbal_cols = [c for c in ["CLSDT_POLICY_NO", "FAC_POLICY_NO", "POLIS_CLN_1", "POLIS_CLN_2"] if c in osbal_cols]
    slip_osbal_cols  = [c for c in ["CLSDT_SLIP_NO", "FAC_SLIP", "SLIP_NO_CLN_1", "SLIP_NO_CLN_2"] if c in osbal_cols]
    insured_osbal_cols = [c for c in osbal_cols if c == "FAC_INSURED" or c.startswith("INSURED_CLN_") or c.startswith("INSURED_CLEAN_")]
    sertif_osbal_cols = [c for c in ["CLSDT_SERTF_NO", "CERTIFICATE_CLN_1", "CERTIFICATE_CLN_2"] if c in osbal_cols]

    polis_facul_cols = [c for c in ["FAC_POLICY_NO", "POLIS_CLN_1"] if c in facul_cols]
    slip_facul_cols  = [c for c in ["FAC_SLIP", "SLIP_NO_CLN_1"] if c in facul_cols]
    insured_facul_cols = [c for c in facul_cols if c == "FAC_INSURED" or c.startswith("INSURED_CLN_") or c.startswith("INSURED_CLEAN_")]

    excluded_osbal = {i for i, r in enumerate(osbal_rows) if _is_excluded_ref_row(r, polis_osbal_cols, slip_osbal_cols)}
    print(f"  Baris OSBAL administratif dikecualikan: {len(excluded_osbal):,}")

    idx_osbal_polis   = _build_exact_index(osbal_rows, polis_osbal_cols, excluded=excluded_osbal)
    idx_osbal_slip    = _build_exact_index(osbal_rows, slip_osbal_cols,  excluded=excluded_osbal)
    idx_osbal_insured = {}
    for i, r in enumerate(osbal_rows):
        if i in excluded_osbal: continue
        for c in insured_osbal_cols:
            v = _normalize(r.get(c))
            if v: idx_osbal_insured.setdefault(v, []).append(i)

    idx_osbal_facode  = {}
    for i, r in enumerate(osbal_rows):
        if i in excluded_osbal: continue
        fc = _normalize(r.get(OSBAL_FACODE_COL))
        if fc: idx_osbal_facode.setdefault(fc, []).append(i)

    tok_osbal_polis   = _build_token_index(osbal_rows, polis_osbal_cols, excluded=excluded_osbal)
    tok_osbal_slip    = _build_token_index(osbal_rows, slip_osbal_cols,  excluded=excluded_osbal)
    idx_osbal_sertif  = _build_sertif_index(osbal_rows, sertif_osbal_cols, excluded=excluded_osbal)

    # FACUL Indexes
    idx_facul_polis   = _build_exact_index(facul_rows, polis_facul_cols)
    idx_facul_slip    = _build_exact_index(facul_rows, slip_facul_cols)
    idx_facul_insured = {}
    for i, r in enumerate(facul_rows):
        for c in insured_facul_cols:
            v = _normalize(r.get(c))
            if v: idx_facul_insured.setdefault(v, []).append(i)

    tok_facul_polis   = _build_token_index(facul_rows, polis_facul_cols)
    tok_facul_slip    = _build_token_index(facul_rows, slip_facul_cols)

    # 3. Matching Loop
    print(f"\n[3/5] Menjalankan matching pada {len(suspend_rows):,} baris suspend ...")
    raw_results = []
    t_match = time.perf_counter()

    for n, sus in enumerate(suspend_rows, 1):
        if n % 2000 == 0:
            print(f"    Memproses baris {n:,} ...", flush=True)

        polis_cands, slip_cands, cert_cands, ins_cands = extract_suspend_identifiers(sus)

        sus_date = sus.get("_sus_date_parsed")
        sus_curr = _normalize(sus.get(SUSPEND_CURR_COL, ""))

        matched_osbal_idx = set()
        source = None
        scenario = ""
        resolved = False

        # -------------------------------------------------------------
        # FASE 0: PASS A (Polis + Sertifikat)
        # -------------------------------------------------------------
        if cert_cands:
            for cert_val in cert_cands:
                if cert_val in idx_osbal_sertif and polis_cands:
                    c_idxs = idx_osbal_sertif[cert_val]
                    conf = [i for i in c_idxs if any(_clean_code(osbal_rows[i].get(c)) in polis_cands for c in polis_osbal_cols)]
                    if conf:
                        matched_osbal_idx.update(conf)
                        source = "OSBAL"
                        scenario = "Polis + Sertifikat"
                        resolved = True
                        break

        # -------------------------------------------------------------
        # FASE 1: JOINT MATCH (Polis + Slip Sekaligus) - PRIORITAS BARU
        # -------------------------------------------------------------
        if not matched_osbal_idx and polis_cands and slip_cands:
            p_hits = _exact_match(polis_cands, idx_osbal_polis)
            s_hits = _exact_match(slip_cands,  idx_osbal_slip)
            both_hits = p_hits & s_hits
            if both_hits:
                matched_osbal_idx.update(both_hits)
                source = "OSBAL"
                scenario = "Polis + Slip"
                resolved = True
            elif facul_rows:
                # Joint di FACUL -> resolve ke OSBAL
                f_p = _exact_match(polis_cands, idx_facul_polis)
                f_s = _exact_match(slip_cands,  idx_facul_slip)
                f_both = f_p & f_s
                if f_both:
                    res_idx = _resolve_facode([facul_rows[i] for i in f_both], FACUL_FACODE_COL, idx_osbal_facode)
                    if res_idx:
                        matched_osbal_idx.update(res_idx)
                        source = "FACUL"
                        scenario = "Polis + Slip"
                        resolved = True

        # -------------------------------------------------------------
        # FASE 2: POLIS MATCHING (Bila Joint belum match)
        # -------------------------------------------------------------
        if not matched_osbal_idx and polis_cands:
            p_hits = _exact_match(polis_cands, idx_osbal_polis)
            if p_hits:
                matched_osbal_idx.update(p_hits)
                source = "OSBAL"
                scenario = "Polis only"
                resolved = True
            elif facul_rows:
                f_p = _exact_match(polis_cands, idx_facul_polis)
                if f_p:
                    res_idx = _resolve_facode([facul_rows[i] for i in f_p], FACUL_FACODE_COL, idx_osbal_facode)
                    if res_idx:
                        matched_osbal_idx.update(res_idx)
                        source = "FACUL"
                        scenario = "Polis only"
                        resolved = True

            # LIKE Polis
            if not matched_osbal_idx:
                like_p = _like_match(polis_cands, tok_osbal_polis, osbal_rows, polis_osbal_cols)
                if like_p:
                    matched_osbal_idx.update(like_p)
                    source = "OSBAL"
                    scenario = "Polis only"
                    resolved = True
                elif facul_rows:
                    f_like = _like_match(polis_cands, tok_facul_polis, facul_rows, polis_facul_cols)
                    if f_like:
                        res_idx = _resolve_facode([facul_rows[i] for i in f_like], FACUL_FACODE_COL, idx_osbal_facode)
                        if res_idx:
                            matched_osbal_idx.update(res_idx)
                            source = "FACUL"
                            scenario = "Polis only"
                            resolved = True

        # -------------------------------------------------------------
        # FASE 3: SLIP MATCHING (Bila Polis tidak match)
        # -------------------------------------------------------------
        if not matched_osbal_idx and slip_cands:
            s_hits = _exact_match(slip_cands, idx_osbal_slip)
            if s_hits:
                matched_osbal_idx.update(s_hits)
                source = "OSBAL"
                scenario = "Slip only"
                resolved = True
            elif facul_rows:
                f_s = _exact_match(slip_cands, idx_facul_slip)
                if f_s:
                    res_idx = _resolve_facode([facul_rows[i] for i in f_s], FACUL_FACODE_COL, idx_osbal_facode)
                    if res_idx:
                        matched_osbal_idx.update(res_idx)
                        source = "FACUL"
                        scenario = "Slip only"
                        resolved = True

            # LIKE Slip
            if not matched_osbal_idx:
                like_s = _like_match(slip_cands, tok_osbal_slip, osbal_rows, slip_osbal_cols)
                if like_s:
                    matched_osbal_idx.update(like_s)
                    source = "OSBAL"
                    scenario = "Slip only"
                    resolved = True
                elif facul_rows:
                    f_like = _like_match(slip_cands, tok_facul_slip, facul_rows, slip_facul_cols)
                    if f_like:
                        res_idx = _resolve_facode([facul_rows[i] for i in f_like], FACUL_FACODE_COL, idx_osbal_facode)
                        if res_idx:
                            matched_osbal_idx.update(res_idx)
                            source = "FACUL"
                            scenario = "Slip only"
                            resolved = True

        # -------------------------------------------------------------
        # FASE 4: INSURED FALLBACK (Hanya jika Polis & Slip tidak ketemu)
        # -------------------------------------------------------------
        if not matched_osbal_idx and ins_cands:
            ins_hits = set()
            for iv in ins_cands:
                if iv in idx_osbal_insured:
                    ins_hits.update(idx_osbal_insured[iv])
            if ins_hits:
                matched_osbal_idx.update(ins_hits)
                source = "OSBAL"
                scenario = "Insured only"
                resolved = True
            elif facul_rows:
                f_ins = set()
                for iv in ins_cands:
                    if iv in idx_facul_insured:
                        f_ins.update(idx_facul_insured[iv])
                if f_ins:
                    res_idx = _resolve_facode([facul_rows[i] for i in f_ins], FACUL_FACODE_COL, idx_osbal_facode)
                    if res_idx:
                        matched_osbal_idx.update(res_idx)
                        source = "FACUL"
                        scenario = "Insured only"
                        resolved = True

        # -------------------------------------------------------------
        # FASE 5: NARROWING MULTI-FAC (Bila > 1 FAC_CODE)
        # -------------------------------------------------------------
        def _get_facs(indices):
            return {_normalize(osbal_rows[i].get(OSBAL_FACODE_COL)) for i in indices if _normalize(osbal_rows[i].get(OSBAL_FACODE_COL))}

        current_facs = _get_facs(matched_osbal_idx)

        if len(current_facs) > 1 and slip_cands:
            # Persempit dengan Slip Sejati
            confirmed_slip = [
                i for i in matched_osbal_idx
                if any(_clean_code(osbal_rows[i].get(c)) in slip_cands for c in slip_osbal_cols)
            ]
            conf_facs = _get_facs(confirmed_slip)
            if confirmed_slip and conf_facs and conf_facs < current_facs:
                matched_osbal_idx = set(confirmed_slip)
                current_facs = conf_facs
                if "Slip" not in scenario:
                    scenario = "Polis + Slip" if "Polis" in scenario else f"{scenario} + Slip"

        if len(current_facs) > 1 and cert_cands:
            # Persempit dengan Sertifikat
            cert_matches = []
            for idx in matched_osbal_idx:
                r_certs = [_clean_cert_str(osbal_rows[idx].get(c, "")) for c in sertif_osbal_cols if _clean_cert_str(osbal_rows[idx].get(c, ""))]
                if any(sc in r_certs for sc in cert_cands) or any(_cert_in_range(sc, rc) for sc in cert_cands for rc in r_certs):
                    cert_matches.append(idx)
            cert_facs = _get_facs(cert_matches)
            if cert_matches and cert_facs and cert_facs < current_facs:
                matched_osbal_idx = set(cert_matches)
                current_facs = cert_facs
                if "Sertifikat" not in scenario:
                    scenario = f"{scenario} + Sertifikat"

        if len(current_facs) > 1 and ins_cands and "Insured" not in scenario:
            # Persempit dengan Insured
            confirmed_ins = [
                i for i in matched_osbal_idx
                if any(_normalize(osbal_rows[i].get(c)) in ins_cands for c in insured_osbal_cols)
            ]
            ins_facs = _get_facs(confirmed_ins)
            if confirmed_ins and ins_facs and ins_facs < current_facs:
                matched_osbal_idx = set(confirmed_ins)
                current_facs = ins_facs
                scenario = f"{scenario} + Insured"

        # -------------------------------------------------------------
        # FASE 6: VALIDASI CURRENCY & PERIODE
        # -------------------------------------------------------------
        is_beda_curr = False
        is_beda_per = False
        if matched_osbal_idx:
            matched_list = list(matched_osbal_idx)
            if sus_curr:
                curr_filtered = [i for i in matched_list if _normalize(osbal_rows[i].get(OSBAL_CURR_COL)) == sus_curr]
                if not curr_filtered:
                    is_beda_curr = True
                elif len(curr_filtered) < len(matched_list):
                    matched_list = curr_filtered

            if not is_beda_curr and pd.notna(sus_date):
                per_filtered = []
                for i in matched_list:
                    com_d = osbal_rows[i].get("_com_date_parsed")
                    if pd.isna(com_d) or sus_date >= com_d:
                        per_filtered.append(i)
                if not per_filtered:
                    is_beda_per = True
                elif len(per_filtered) < len(matched_list):
                    matched_list = per_filtered

            matched_osbal_idx = set(matched_list)
            current_facs = _get_facs(matched_osbal_idx)

        osbal_count = len(matched_osbal_idx)
        osbal_ref = [osbal_rows[i] for i in matched_osbal_idx] if matched_osbal_idx else [{}]

        if scenario and " only + " in scenario:
            scenario = scenario.replace(" only + ", " + ")

        raw_results.append({
            "suspend": sus,
            "source": source,
            "base_scenario": scenario,
            "is_beda_curr": is_beda_curr,
            "is_beda_per": is_beda_per,
            "osbal_idx": list(matched_osbal_idx),
            "osbal_count": osbal_count,
            "resolved": resolved,
            "fac_codes": current_facs,
            "clean_polis_val": polis_cands[0] if polis_cands else "",
            "clean_slip_val": slip_cands[0] if slip_cands else "",
        })

    print(f"  Matching selesai dalam {time.perf_counter()-t_match:.1f} detik.")

    # 4. Akumulasi & Output Build
    print("\n[4/5] Memproses akumulasi dan membuat output DataFrame ...")
    groups = {}
    final = []

    for r in raw_results:
        r.setdefault("suspend_count", 1)
        if (len(r["fac_codes"]) == 1 and r["source"] is not None
                and not r["is_beda_curr"] and not r["is_beda_per"]):
            groups.setdefault(next(iter(r["fac_codes"])), []).append(r)
        else:
            final.append(r)

    for fac, group in groups.items():
        if len(group) == 1:
            group[0]["suspend_count"] = 1
            final.append(group[0])
            continue
        merged_osbal_idx = set()
        for g in group:
            merged_osbal_idx.update(g["osbal_idx"])
        merged_sus = _merge_suspend_rows([g["suspend"] for g in group])
        for g in group:
            g["osbal_idx"]        = list(merged_osbal_idx)
            g["osbal_count"]      = len(merged_osbal_idx)
            g["suspend_count"]    = len(group)
            g["_merged_amount_ori"] = merged_sus.get("AMOUNT ORI", 0)
            final.append(g)

    output_rows = []
    for r in final:
        ref_osbal = [osbal_rows[i] for i in r["osbal_idx"]] if r["source"] else [{}]
        all_fac_osbal_rows = None
        osbal_cnt = r["osbal_count"]
        if r["source"] and len(r["fac_codes"]) == 1:
            fac = next(iter(r["fac_codes"]))
            all_fac_idx = idx_osbal_facode.get(fac, [])
            if len(all_fac_idx) > 0:
                all_fac_osbal_rows = [osbal_rows[i] for i in all_fac_idx]
                osbal_cnt = len(all_fac_idx)

        final_scenario = _format_final_scenario(
            base_scenario=r.get("base_scenario", ""),
            source=r["source"],
            is_beda_curr=r.get("is_beda_curr", False),
            is_beda_per=r.get("is_beda_per", False),
            num_fac=len(r["fac_codes"]),
        )

        out = _build_output_row(
            r["suspend"], source=r["source"], scenario=final_scenario,
            osbal_rows=ref_osbal, osbal_count=osbal_cnt,
            resolved=r["resolved"], suspend_count=r.get("suspend_count", 1),
            all_fac_osbal_rows=all_fac_osbal_rows,
            clean_polis_val=r["clean_polis_val"], clean_slip_val=r["clean_slip_val"],
        )
        if "_merged_amount_ori" in r:
            out["_MERGED_AMOUNT_ORI"] = r["_merged_amount_ori"]
        output_rows.append(out)

    df_out = pd.DataFrame(output_rows)

    # Format numeric
    for col in ["AMOUNT ORI", "CCOS_OR_BAL", "CCOS_BAL_DUE"]:
        if col in df_out.columns:
            s = df_out[col]
            if s.dtype == object:
                s = s.astype(str).str.strip().str.replace(r'\s+', '', regex=True)
                has_comma = s.str.contains(',', na=False)
                s_eu = s.str.replace('.', '', regex=False).str.replace(',', '.', regex=False)
                s = s_eu.where(has_comma, s)
            df_out[col] = pd.to_numeric(s, errors="coerce")

    df_out = _compute_derived_cols(df_out)
    if "SKENARIO" in df_out.columns:
        df_out["SKENARIO"] = df_out["SKENARIO"].astype(str).str.replace("Sertifikat", "Cert", regex=False)
    df_out.loc[df_out["SKENARIO"].isin(["UNMATCHED", "Unmatching"]), "FLAG_PROD"] = "Unmatching"
    df_out.loc[df_out["SKENARIO"] == "UNMATCHED", "SKENARIO"] = "Unmatching"

    _blank_balance_flags = {"Unmatching", "Matching >1 fac code", "Beda Periode", "Beda Currency"}
    if "FLAG_PROD" in df_out.columns:
        _blank_mask = df_out["FLAG_PROD"].isin(_blank_balance_flags)
        for _balance_col in ("CCOS_BAL_DUE", "CCOS_OR_BAL"):
            if _balance_col in df_out.columns:
                df_out.loc[_blank_mask, _balance_col] = np.nan

    for col in FINAL_COLUMNS:
        if col not in df_out.columns:
            df_out[col] = ""
    df_out = df_out[FINAL_COLUMNS]

    # 5. Export ke Excel
    print(f"\n[5/5] Menyimpan hasil debug ke: {OUTPUT_FILE} ...")
    df_out.to_excel(OUTPUT_FILE, index=False)

    try:
        wb_fmt = load_workbook(OUTPUT_FILE)
        ws_fmt = wb_fmt.active
        header_map = {cell.value: cell.column for cell in ws_fmt[1]}
        text_cols = [
            "CCOS_DOC_NO", "CCOS_REF_CODE", "RECEIPT NO", "CREDIT NOTES",
            "DETAIL RINCIAN NO", "POLIS", "POLICY_CLEAN_1", "POLICY_CLEAN_2",
            "POLICY_CLEAN_3", "POLICY_CLEAN_4", "POLICY_CLEAN_5",
            "SERTIF_CLEAN_1", "SERTIF_CLEAN_2", "SERTIF_CLEAN_3",
            "SLIP_NO", "SLIP_NO_CLEAN"
        ]
        for cn in text_cols:
            cidx = header_map.get(cn)
            if not cidx: continue
            for r_idx in range(2, ws_fmt.max_row + 1):
                cell = ws_fmt.cell(row=r_idx, column=cidx)
                if cell.value is not None:
                    cell.number_format = "@"
                    cell.value = str(cell.value)
        wb_fmt.save(OUTPUT_FILE)
    except Exception as e:
        print(f"  [WARN] Format anti-E+ gagal: {e}")

    # =========================================================================
    # LAPORAN KOMPARASI (BASELINE vs DEBUG REVISED)
    # =========================================================================
    print("\n" + "=" * 75)
    print(" LAPORAN EVALUASI & KOMPARASI HASIL MATCHING")
    print("=" * 75)
    
    if os.path.exists(BASELINE_FILE):
        df_base = pd.read_excel(BASELINE_FILE)
        print(f"\n{'Kategori FLAG_PROD':<42} {'Baseline':>12} {'Revised (Debug)':>18} {'Selisih':>12}")
        print("-" * 86)
        
        all_flags = sorted(list(set(df_base["FLAG_PROD"].dropna().unique()) | set(df_out["FLAG_PROD"].dropna().unique())))
        for flag in all_flags:
            cnt_base = (df_base["FLAG_PROD"] == flag).sum()
            cnt_rev  = (df_out["FLAG_PROD"] == flag).sum()
            diff     = cnt_rev - cnt_base
            diff_str = f"+{diff:,}" if diff > 0 else f"{diff:,}"
            print(f"{flag:<42} {cnt_base:>12,} {cnt_rev:>18,} {diff_str:>12}")

        print("-" * 86)
        total_base = len(df_base)
        total_rev  = len(df_out)
        print(f"{'TOTAL BARIS':<42} {total_base:>12,} {total_rev:>18,} {total_rev - total_base:>12}")

        print("\n\n" + "=" * 75)
        print(" KOMPARASI SKENARIO TOP 15")
        print("=" * 75)
        print(f"{'Skenario':<45} {'Baseline':>12} {'Revised (Debug)':>15}")
        print("-" * 75)
        top_scens = df_out["SKENARIO"].value_counts().head(15)
        for sc, cnt_rev in top_scens.items():
            cnt_base = (df_base["SKENARIO"] == sc).sum() if "SKENARIO" in df_base.columns else 0
            print(f"{sc:<45} {cnt_base:>12,} {cnt_rev:>15,}")

    elapsed = time.perf_counter() - t_start
    print(f"\nEksekusi selesai dalam {elapsed:.1f} detik ({elapsed/60:.1f} menit).")
    print(f"File output debug tersimpan di: {OUTPUT_FILE}")
    print("=" * 75)

if __name__ == "__main__":
    run_debug()
