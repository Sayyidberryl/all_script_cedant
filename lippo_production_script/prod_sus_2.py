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
# KONFIGURASI LIPPO (V2: MERGE ROWS PER FAC CODE)
# =============================================================================

SCRIPT_DIR   = os.path.dirname(os.path.abspath(__file__))
DATA_DIR     = os.path.join(SCRIPT_DIR, "data")

SUSPEND_FILE = os.path.join(DATA_DIR, "lippo_output_suspense_V3.xlsx")
OSBAL_FILE   = os.path.join(DATA_DIR, "lippo_output_osbal_V2.xlsx")
FACUL_FILE   = os.path.join(DATA_DIR, "lippo_output_facul_V3.xlsx")
OUTPUT_FILE  = os.path.join(DATA_DIR, "final_output_v2.xlsx")

# Kolom FAC CODE
OSBAL_FACODE_COL = "CCOS_REF_CODE"
FACUL_FACODE_COL = "FAC_CODE"

# Kolom Currency & Periode
SUSPEND_CURR_COL = "CURR ORI"
OSBAL_CURR_COL   = "CCOS_CURR"
SUSPEND_DATE_COL = "RECEIPT DATE"
OSBAL_DATE_COL   = "FAC_COM_DATE"

# Marker baris administratif
_EXCLUDED_REF_MARKERS = ["HUTANG PIUTANG", "DATA SUSPENSE"]
_MIN_TOKEN_LEN = 3
_TOKEN_RE      = re.compile(r"[^A-Z0-9]+")
_DELIMITER_RE  = re.compile(r"[,;|+\s]+")

# Kolom join untuk akumulasi V2 (merge rows)
_SUSPEND_JOIN_COLS = [
    "RECEIPT NO", "CREDIT NOTES", "DETAIL RINCIAN NO", "RECEIPT DATE",
    "CEDANT NAME", "CEDANT SHRT NAME",
    "INSURED", "INSURED_CLN_1", "INSURED_CLN_2",
    "CURR ORI", "CURR PAY", "AMOUNT PAY",
    "POLIS", "POLIS_CLN_1", "POLIS_CLN_2", "CERTIFICATE",
    "SLIP NO", "SLIP_NO_CLN_1",
    "DESC 1", "DESC 2", "DESC 3", "DESC 4",
    "STATUS", "REC_TYPE",
]
_SUSPEND_SUM_COLS = ["AMOUNT ORI"]

# Output 38 kolom standar V2 (konsisten dengan V1 dan PostgreSQL schema)
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
    if text.endswith(".0"):
        text = text[:-2]
    return re.sub(r"\s+", " ", text) if "  " in text else text


def _normalize(value) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    return _norm_cached(str(value))


def _clean_cert_str(val) -> str:
    if val is None or (isinstance(value:=val, float) and pd.isna(value)):
        return ""
    s = str(val).strip()
    return re.sub(r'\.0+$', '', s)


def _expand_sertif_range(val: str) -> list:
    if not val:
        return []
    s = re.sub(r'\.0+$', '', str(val).strip())
    clean = re.sub(r'\s*SD\s*|\s*S/D\s*', '-', s, flags=re.IGNORECASE)
    clean = re.sub(r'\s+', '', clean)
    if clean.isdigit() and 1 <= len(clean) <= 6:
        return [clean.zfill(6)]
    m = re.match(r'^(\d{1,6})-(\d{1,6})$', clean)
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
    clean = re.sub(r'\s*SD\s*|\s*S/D\s*', '-', s, flags=re.IGNORECASE)
    clean = re.sub(r'\s+', '', clean)
    if clean.isdigit():
        return q_num == int(clean)
    m = re.match(r'^(\d{1,6})-(\d{1,6})$', clean)
    if m:
        start, end = int(m.group(1)), int(m.group(2))
        if start <= end:
            return start <= q_num <= end
        return q_num >= start or q_num <= end
    return False


def _slip_in_range(sus_slip: str, osbal_row: dict) -> bool:
    if not sus_slip:
        return False
    s_clean = re.sub(r'\D', '', str(sus_slip))
    if not s_clean:
        return False
    sl1 = re.sub(r'\D', '', str(osbal_row.get("SLIP_NO_CLN_1", osbal_row.get("clean slip 1", ""))))
    sl2 = re.sub(r'\D', '', str(osbal_row.get("SLIP_NO_CLN_2", osbal_row.get("clean slip 2", ""))))
    if sl1 and sl2 and len(sl1) == len(s_clean) and len(sl2) == len(s_clean):
        try:
            n_sus = int(s_clean)
            n_sl1 = int(sl1)
            n_sl2 = int(sl2)
            if min(n_sl1, n_sl2) <= n_sus <= max(n_sl1, n_sl2):
                return True
        except ValueError:
            pass
    return False


def _tokenize(text: str) -> frozenset:
    if not text:
        return frozenset()
    return frozenset(t for t in _TOKEN_RE.split(text) if len(t) >= _MIN_TOKEN_LEN)


def _get_clean_cols(columns: list, prefix: str) -> list:
    if not columns:
        return []
    alt_map = {
        "clean polis": ["clean polis", "policy_clean", "polis_cln"],
        "clean slip": ["clean slip", "slip_clean", "slip_no_cln"],
        "clean insured": ["clean insured", "insured_clean", "insured_"],
        "clean sertif": ["clean sertif", "sertif_clean", "sertif_cln", "certificate", "clsdt_sertf"],
    }
    prefixes = alt_map.get(prefix.lower(), [prefix.lower()])
    res = []
    for c in columns:
        clow = str(c).lower()
        if any(clow.startswith(p) for p in prefixes):
            res.append(c)
    return res


def _format_sertif(val) -> str:
    if pd.isna(val) or str(val).strip() == "":
        return ""
    if isinstance(val, (int, float)):
        return str(int(val)).zfill(6)
    s = str(val).strip()
    if s.endswith('.0'):
        s = s[:-2]
    return s.zfill(6) if s.isdigit() else s


def _is_excluded_ref_row(ref_row: dict, polis_cols: list, slip_cols: list) -> bool:
    for col in list(polis_cols) + list(slip_cols):
        val = _normalize(ref_row.get(col, ""))
        if val and any(m in val for m in _EXCLUDED_REF_MARKERS):
            return True
    return False


def _collect_clean_values(row: dict, cols: list) -> list:
    seen, result = set(), []
    for col in cols:
        val = _normalize(row.get(col, ""))
        if val and val not in seen:
            seen.add(val)
            result.append(val)
    return result


def _ref_has_value(ref_row: dict, clean_cols: list, query_values: list) -> bool:
    if not query_values:
        return False
    ref_vals = [_normalize(ref_row.get(c, "")) for c in clean_cols]
    ref_vals_nonempty = [rv for rv in ref_vals if rv]
    if any(qv in ref_vals_nonempty for qv in query_values):
        return True
    return any(v in rv for v in query_values if v and len(v) >= 10 for rv in ref_vals_nonempty)


# =============================================================================
# INDEX BUILDERS
# =============================================================================

def _build_exact_index(rows: list, cols: list, excluded: set = None) -> dict:
    index = {}
    for i, row in enumerate(rows):
        if excluded and i in excluded:
            continue
        for col in cols:
            val = _normalize(row.get(col, ""))
            if val:
                index.setdefault(val, []).append(i)
    return index


def _build_facode_index(rows: list, col: str, excluded: set = None) -> dict:
    index = {}
    for i, row in enumerate(rows):
        if excluded and i in excluded:
            continue
        val = _normalize(row.get(col, ""))
        if val:
            index.setdefault(val, []).append(i)
    return index


def _build_token_index(rows: list, clean_cols: list, excluded: set = None) -> tuple:
    token_index = {}
    token_cache = {}
    for i, row in enumerate(rows):
        if excluded and i in excluded:
            continue
        tokens = frozenset()
        for c in clean_cols:
            t = _normalize(row.get(c, ""))
            if t:
                tokens = tokens | _tokenize(t)
        token_cache[i] = tokens
        for tok in tokens:
            token_index.setdefault(tok, set()).add(i)
    return token_index, token_cache


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


def _build_lookup(rows: list, polis_cols: list, slip_cols: list,
                  insured_cols: list, facode_col: str = None, label: str = "") -> tuple:
    if not rows:
        return ({}, {}, {}, ({}, {}), ({}, {}), ({}, {}), {})

    t0 = time.perf_counter()
    excluded = {i for i, r in enumerate(rows) if _is_excluded_ref_row(r, polis_cols, slip_cols)}
    facode_index = _build_facode_index(rows, facode_col, excluded=excluded) if facode_col else {}
    lookup = (
        _build_exact_index(rows, slip_cols,    excluded=excluded),
        _build_exact_index(rows, polis_cols,   excluded=excluded),
        _build_exact_index(rows, insured_cols, excluded=excluded),
        _build_token_index(rows, slip_cols,    excluded=excluded),
        _build_token_index(rows, polis_cols,   excluded=excluded),
        _build_token_index(rows, insured_cols, excluded=excluded),
        facode_index,
    )
    elapsed = time.perf_counter() - t0
    exc_info = f" - {len(excluded):,} baris dikecualikan" if excluded else ""
    print(f"  {label:<6} done ({elapsed:.1f}s){exc_info}", flush=True)
    return lookup


# =============================================================================
# MATCH ENGINE
# =============================================================================

def _exact_match(query_values: list, index: dict) -> set:
    return {i for v in query_values if v and v in index for i in index[v]}


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


def _resolve_facode(matched_rows: list, facode_col: str, facode_osbal_idx: dict, osbal_rows: list) -> tuple:
    osbal_indices = set()
    found_fac_codes = set()
    for row in matched_rows:
        fac = _normalize(row.get(facode_col, ""))
        if fac:
            found_fac_codes.add(fac)
            if fac in facode_osbal_idx:
                osbal_indices.update(facode_osbal_idx[fac])
    return osbal_indices, found_fac_codes


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


def _build_output_row(
    suspend_row: dict, source: str, scenario: str, osbal_rows: list,
    osbal_count: int = 0, resolved: bool = True,
    suspend_count: int = 1, all_fac_osbal_rows: list = None,
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
        "INSURED_1":         suspend_row.get("INSURED_CLN_1") or suspend_row.get("clean insured 1") or "",
        "INSURED_2":         suspend_row.get("INSURED_CLN_2") or suspend_row.get("clean insured 2") or "",
        "CURR ORI":          suspend_row.get("CURR ORI",          ""),
        "AMOUNT ORI":        suspend_row.get("AMOUNT ORI",        ""),
        "CURR PAY":          suspend_row.get("CURR PAY",          ""),
        "AMOUNT PAY":        suspend_row.get("AMOUNT PAY",        ""),
        "CCOS_OR_BAL":       ccos["CCOS_OR_BAL"],
        "CCOS_BAL_DUE":      ccos["CCOS_BAL_DUE"],
        "POLIS":             suspend_row.get("POLIS") or "",
        "POLICY_CLEAN_1":    suspend_row.get("POLIS_CLN_1") or "",
        "SERTIF_CLEAN_1":    _format_sertif(suspend_row.get("CERTIFICATE") or suspend_row.get("CERTIFICATE_1") or ""),
        "POLICY_CLEAN_2":    suspend_row.get("POLIS_CLN_2") or "",
        "SERTIF_CLEAN_2":    _format_sertif(suspend_row.get("CERTIFICATE_2") or ""),
        "POLICY_CLEAN_3":    suspend_row.get("POLIS_CLN_3") or "",
        "SERTIF_CLEAN_3":    _format_sertif(suspend_row.get("CERTIFICATE_3") or ""),
        "POLICY_CLEAN_4":    suspend_row.get("POLIS_CLN_4") or "",
        "POLICY_CLEAN_5":    suspend_row.get("POLIS_CLN_5") or "",
        "SLIP_NO":           suspend_row.get("SLIP NO") or suspend_row.get("SLIP_NO") or "",
        "SLIP_NO_CLN":       suspend_row.get("SLIP_NO_CLN_1") or suspend_row.get("SLIP_NO_CLN") or "",
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


# =============================================================================
# PIPELINE EXECUTION
# =============================================================================

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


def run() -> None:
    t_start = time.perf_counter()
    print("\n" + "=" * 60, flush=True)
    print("  PRODUCTION SCRIPT — SUSPEND MATCHING (LIPPO) — V2", flush=True)
    print("=" * 60, flush=True)

    # [1] Load data
    print("\n[1/6] Loading data ...", flush=True)
    suspend_rows, suspend_cols = _load_excel(SUSPEND_FILE, "SUSPEND")
    osbal_rows,   osbal_cols   = _load_excel(OSBAL_FILE,   "OSBAL")
    facul_rows,   facul_cols   = _load_excel(FACUL_FILE,   "FACUL")

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
    print("\n[2/6] Detecting clean columns ...", flush=True)
    polis_sus   = _get_clean_cols(suspend_cols, "clean polis")
    slip_sus    = _get_clean_cols(suspend_cols, "clean slip")
    insured_sus = _get_clean_cols(suspend_cols, "clean insured")
    sertif_sus  = _get_clean_cols(suspend_cols, "clean sertif")

    polis_osbal   = _get_clean_cols(osbal_cols, "clean polis")
    slip_osbal    = _get_clean_cols(osbal_cols, "clean slip")
    insured_osbal = _get_clean_cols(osbal_cols, "clean insured")
    sertif_osbal  = _get_clean_cols(osbal_cols, "clean sertif")

    polis_facul   = _get_clean_cols(facul_cols, "clean polis")
    slip_facul    = _get_clean_cols(facul_cols, "clean slip")
    insured_facul = _get_clean_cols(facul_cols, "clean insured")
    sertif_facul  = _get_clean_cols(facul_cols, "clean sertif")

    def _prepend(col, cols_list, header):
        return ([col] + cols_list) if col in header and col not in cols_list else cols_list

    polis_osbal  = _prepend("CLSDT_POLICY_NO", polis_osbal,  osbal_cols)
    polis_osbal  = _prepend("FAC_POLICY_NO",   polis_osbal,  osbal_cols)
    slip_osbal   = _prepend("CLSDT_SLIP_NO",   slip_osbal,   osbal_cols)
    slip_osbal   = _prepend("FAC_SLIP",        slip_osbal,   osbal_cols)

    polis_facul  = _prepend("FAC_POLICY_NO",   polis_facul,  facul_cols)
    slip_facul   = _prepend("FAC_SLIP",        slip_facul,   facul_cols)

    # [3] Bangun lookup index
    print("\n[3/6] Building lookup indexes ...", flush=True)
    lookup_osbal = _build_lookup(osbal_rows, polis_osbal, slip_osbal, insured_osbal, OSBAL_FACODE_COL, label="OSBAL")
    lookup_facul = _build_lookup(facul_rows, polis_facul, slip_facul, insured_facul, FACUL_FACODE_COL, label="FACUL")

    facode_osbal_idx = lookup_osbal[6]
    excluded_osbal   = {i for i, r in enumerate(osbal_rows) if _is_excluded_ref_row(r, polis_osbal, slip_osbal)}
    sertif_osbal_idx = _build_sertif_index(osbal_rows, sertif_osbal, excluded=excluded_osbal)
    print(f"  Sertif OSBAL index: {len(sertif_osbal_idx):,} sertif unik", flush=True)

    # [4] Matching per baris
    print(f"\n[4/6] Matching {len(suspend_rows):,} baris suspend (Basic Rules, No Bordero) ...", flush=True)
    t4 = time.perf_counter()
    raw_results = []

    for n, sus in enumerate(suspend_rows, 1):
        if n % 1000 == 0:
            print(f"    Memproses baris {n:,} ...", flush=True)

        eff_polis = polis_sus
        eff_slip  = slip_sus
        sus_date  = sus.get("_sus_date_parsed")

        matched_fac_codes = set()
        matched_osbal_idx = set()
        source = None
        scenario = ""
        resolved = False

        # FASE 0: PASS A (Polis + Sertif)
        sus_cert_raw = sus.get("CERTIFICATE", "")
        sus_cert_val = _clean_cert_str(sus_cert_raw) if sus_cert_raw else ""
        if sus_cert_val and sus_cert_val in sertif_osbal_idx:
            c_idxs = sertif_osbal_idx[sus_cert_val]
            p_vals = _collect_clean_values(sus, eff_polis)
            if p_vals:
                confirmed_idxs = [i for i in c_idxs if _ref_has_value(osbal_rows[i], polis_osbal, p_vals)]
                if confirmed_idxs:
                    matched_osbal_idx.update(confirmed_idxs)
                    source = "OSBAL"
                    scenario = "Polis + Sertifikat"
                    resolved = True

        # FASE 1: POLIS
        if not matched_osbal_idx:
            polis_vals = _collect_clean_values(sus, eff_polis)
            if polis_vals:
                # R1: OSBAL polis exact
                curr_hits = _exact_match(polis_vals, lookup_osbal[1])
                if curr_hits:
                    matched_osbal_idx.update(curr_hits)
                    source = "OSBAL"
                    scenario = "Polis only"
                    resolved = True
                
                # R2: FACUL polis exact -> OSBAL
                elif facul_rows:
                    fac_hits = _exact_match(polis_vals, lookup_facul[1])
                    if fac_hits:
                        res_idx, _ = _resolve_facode([facul_rows[i] for i in fac_hits], FACUL_FACODE_COL, facode_osbal_idx, osbal_rows)
                        if res_idx:
                            matched_osbal_idx.update(res_idx)
                            source = "FACUL"
                            scenario = "Polis only"
                            resolved = True

                # R3: OSBAL polis LIKE
                if not matched_osbal_idx:
                    like_hits = _like_match(polis_vals, lookup_osbal[4][0], osbal_rows, polis_osbal)
                    if like_hits:
                        matched_osbal_idx.update(like_hits)
                        source = "OSBAL"
                        scenario = "Polis only"
                        resolved = True

                # R4: FACUL polis LIKE -> OSBAL
                elif not matched_osbal_idx and facul_rows:
                    fac_like = _like_match(polis_vals, lookup_facul[4][0], facul_rows, polis_facul)
                    if fac_like:
                        res_idx, _ = _resolve_facode([facul_rows[i] for i in fac_like], FACUL_FACODE_COL, facode_osbal_idx, osbal_rows)
                        if res_idx:
                            matched_osbal_idx.update(res_idx)
                            source = "FACUL"
                            scenario = "Polis only"
                            resolved = True

        # FASE 2: SLIP
        if not matched_osbal_idx:
            slip_vals = _collect_clean_values(sus, eff_slip)
            if slip_vals:
                # R5: OSBAL slip exact
                curr_hits = _exact_match(slip_vals, lookup_osbal[0])
                if curr_hits:
                    matched_osbal_idx.update(curr_hits)
                    source = "OSBAL"
                    scenario = "Slip only"
                    resolved = True

                # R6: FACUL slip exact -> OSBAL
                elif facul_rows:
                    fac_hits = _exact_match(slip_vals, lookup_facul[0])
                    if fac_hits:
                        res_idx, _ = _resolve_facode([facul_rows[i] for i in fac_hits], FACUL_FACODE_COL, facode_osbal_idx, osbal_rows)
                        if res_idx:
                            matched_osbal_idx.update(res_idx)
                            source = "FACUL"
                            scenario = "Slip only"
                            resolved = True

                # R7: OSBAL slip LIKE
                if not matched_osbal_idx:
                    like_hits = _like_match(slip_vals, lookup_osbal[3][0], osbal_rows, slip_osbal)
                    if like_hits:
                        matched_osbal_idx.update(like_hits)
                        source = "OSBAL"
                        scenario = "Slip only"
                        resolved = True

                # R8: FACUL slip LIKE -> OSBAL
                elif not matched_osbal_idx and facul_rows:
                    fac_like = _like_match(slip_vals, lookup_facul[3][0], facul_rows, slip_facul)
                    if fac_like:
                        res_idx, _ = _resolve_facode([facul_rows[i] for i in fac_like], FACUL_FACODE_COL, facode_osbal_idx, osbal_rows)
                        if res_idx:
                            matched_osbal_idx.update(res_idx)
                            source = "FACUL"
                            scenario = "Slip only"
                            resolved = True

        # FASE 3: INSURED FALLBACK
        if not matched_osbal_idx:
            ins_vals = _collect_clean_values(sus, insured_sus)
            if ins_vals:
                curr_hits = _exact_match(ins_vals, lookup_osbal[2])
                if curr_hits:
                    matched_osbal_idx.update(curr_hits)
                    source = "OSBAL"
                    scenario = "Insured only"
                    resolved = True
                elif facul_rows:
                    fac_hits = _exact_match(ins_vals, lookup_facul[2])
                    if fac_hits:
                        res_idx, _ = _resolve_facode([facul_rows[i] for i in fac_hits], FACUL_FACODE_COL, facode_osbal_idx, osbal_rows)
                        if res_idx:
                            matched_osbal_idx.update(res_idx)
                            source = "FACUL"
                            scenario = "Insured only"
                            resolved = True

        # FASE 4: NARROWING WAJIB (CURRENCY & PERIODE)
        is_beda_curr = False
        is_beda_per = False
        if matched_osbal_idx:
            matched_list = list(matched_osbal_idx)
            sus_curr = _normalize(sus.get(SUSPEND_CURR_COL, ""))
            if sus_curr:
                filtered_curr = [i for i in matched_list if _normalize(osbal_rows[i].get(OSBAL_CURR_COL, "")) == sus_curr]
                if not filtered_curr:
                    is_beda_curr = True
                elif len(filtered_curr) < len(matched_list):
                    matched_list = filtered_curr

            if not is_beda_curr and pd.notna(sus_date):
                filtered_per = []
                for i in matched_list:
                    com_date = osbal_rows[i].get("_com_date_parsed")
                    if pd.isna(com_date):
                        filtered_per.append(i)
                    else:
                        if sus_date >= com_date:
                            filtered_per.append(i)
                if not filtered_per:
                    is_beda_per = True
                elif len(filtered_per) < len(matched_list):
                    matched_list = filtered_per

            matched_osbal_idx = set(matched_list)

        # FASE 5: NARROWING UTAMA (Non-Destructive)
        if len(matched_osbal_idx) > 1:
            matched_list = list(matched_osbal_idx)

            # Step 1: Slip Exact Matching
            slip_vals = _collect_clean_values(sus, slip_sus)
            if slip_vals:
                confirmed = [idx for idx in matched_list if _ref_has_value(osbal_rows[idx], slip_osbal, slip_vals)]
                if confirmed:
                    matched_list = confirmed
                    scenario = "Polis + Slip" if "Polis" in scenario else scenario

            # Step 1b: Slip Numeric Range Check
            if len(matched_list) > 1 and slip_vals:
                confirmed_slip = [
                    idx for idx in matched_list
                    if any(_slip_in_range(sv, osbal_rows[idx]) for sv in slip_vals)
                ]
                if confirmed_slip:
                    matched_list = confirmed_slip
                    if "Slip" not in scenario:
                        scenario = "Polis + Slip" if "Polis" in scenario else scenario

            # Step 2: Sertifikat Exact & In-Range Matching
            sus_cert_vals = set(_clean_cert_str(v) for v in _collect_clean_values(sus, sertif_sus) if _clean_cert_str(v))
            if sus_cert_vals and len(matched_list) > 1:
                cert_matches = []
                for idx in matched_list:
                    ref_r = osbal_rows[idx]
                    r_certs = [_clean_cert_str(ref_r.get(c, "")) for c in sertif_osbal if _clean_cert_str(ref_r.get(c, ""))]
                    if any(sc in r_certs for sc in sus_cert_vals):
                        cert_matches.append(idx)
                    elif any(_cert_in_range(sc, rc) for sc in sus_cert_vals for rc in r_certs):
                        cert_matches.append(idx)
                if cert_matches:
                    matched_list = cert_matches
                    if "Sertifikat" not in scenario:
                        scenario = f"{scenario} + Sertifikat"

            # Step 3: Insured Token Matching
            if len(matched_list) > 1:
                ins_vals = _collect_clean_values(sus, insured_sus)
                if ins_vals:
                    confirmed_ins = [
                        idx for idx in matched_list
                        if _ref_has_value(osbal_rows[idx], insured_osbal, ins_vals)
                    ]
                    if confirmed_ins:
                        matched_list = confirmed_ins
                        if "Insured" not in scenario:
                            scenario = f"{scenario} + Insured"

            # Step 4: Amount Ori vs CCOS_BAL_DUE Matching
            if len(matched_list) > 1:
                amt_ori_raw = sus.get("AMOUNT ORI")
                try:
                    amt_sus = abs(float(amt_ori_raw)) if amt_ori_raw is not None and not pd.isna(amt_ori_raw) else None
                except (ValueError, TypeError):
                    amt_sus = None
                if amt_sus is not None and amt_sus > 0:
                    amt_matches = [
                        idx for idx in matched_list
                        if abs(abs(float(osbal_rows[idx].get("CCOS_BAL_DUE", 0) or 0)) - amt_sus) <= 1.0
                    ]
                    if amt_matches:
                        matched_list = amt_matches

            matched_osbal_idx = set(matched_list)

        osbal_count = len(matched_osbal_idx)
        osbal_ref = [osbal_rows[i] for i in matched_osbal_idx] if matched_osbal_idx else [{}]

        if scenario and " only + " in scenario:
            scenario = scenario.replace(" only + ", " + ")

        fac_codes = {_normalize(r.get(OSBAL_FACODE_COL, "")) for r in osbal_ref if r}
        fac_codes.discard("")

        raw_results.append({
            "suspend": sus, "source": source,
            "base_scenario": scenario,
            "is_beda_curr": is_beda_curr,
            "is_beda_per": is_beda_per,
            "osbal_idx": list(matched_osbal_idx),
            "osbal_count": osbal_count, "resolved": resolved, "fac_codes": fac_codes,
        })

    print(f"  Matching selesai dalam {time.perf_counter()-t4:.1f}s", flush=True)

    # [4c/6] Fac code claim & accumulation (V2: MERGE ROWS PER FAC CODE)
    print("\n[4c/6] Fac code claim & accumulation (V2: merge rows) ...", flush=True)

    claimed = set()
    _EXCL = {"Unmatching", "Beda Currency", "Beda Periode", "UNMATCHED"}
    for r in raw_results:
        if len(r["fac_codes"]) == 1 and not any(exc in r["base_scenario"] for exc in _EXCL) and not r["is_beda_curr"] and not r["is_beda_per"]:
            claimed |= r["fac_codes"]

    for r in raw_results:
        if len(r["fac_codes"]) > 1:
            remaining = r["fac_codes"] - claimed
            if len(remaining) == 1 and not any(exc in r["base_scenario"] for exc in _EXCL) and not r["is_beda_curr"] and not r["is_beda_per"]:
                r["fac_codes"] = remaining
                claimed |= remaining
                fac = next(iter(remaining))
                r["osbal_idx"] = [i for i in r["osbal_idx"]
                                   if _normalize(osbal_rows[i].get(OSBAL_FACODE_COL, "")) == fac]
                r["osbal_count"] = len(r["osbal_idx"])

    # Group per fac untuk di-merge menjadi 1 baris per FAC CODE (Mode V2)
    groups = {}
    final = []

    for r in raw_results:
        r.setdefault("suspend_count", 1)
        if (len(r["fac_codes"]) == 1 and r["source"] is not None
                and not any(exc in r["base_scenario"] for exc in _EXCL)
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

    print(f"  {len(raw_results):,} suspend rows -> {len(final):,} output rows "
          f"({len(raw_results) - len(final):,} baris diserap akumulasi)", flush=True)

    # [5] Build output DataFrame
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
        )
        if "_merged_amount_ori" in r:
            out["_MERGED_AMOUNT_ORI"] = r["_merged_amount_ori"]
        output_rows.append(out)

    df = pd.DataFrame(output_rows)

    # Fix format angka koma-desimal
    for col in ["AMOUNT ORI", "CCOS_OR_BAL", "CCOS_BAL_DUE"]:
        if col in df.columns:
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

    # [6] Export
    print(f"\n[6/6] Saving to: {OUTPUT_FILE} ...", flush=True)
    for col in FINAL_COLUMNS:
        if col not in df.columns:
            df[col] = ""
    df = df[FINAL_COLUMNS]

    for c in df.select_dtypes(include=['object']).columns:
        df[c] = df[c].astype(str).str.replace(r'[\x00-\x08\x0B\x0C\x0E-\x1F]', '', regex=True)
    df.to_excel(OUTPUT_FILE, index=False)

    try:
        from excel_styler import apply_purple_column_style
        apply_purple_column_style(OUTPUT_FILE, "matching")
    except Exception as e:
        print(f"  [WARN] Styling gagal: {e}")

    # Export ke PostgreSQL
    env_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), ".env")
    load_dotenv(env_path)
    load_dotenv()
    print("  Exporting ke PostgreSQL ...")
    db_user = os.environ.get("DB_USER", "viy4user")
    db_pass = os.environ.get("DB_PASS", "fild42op61nx")
    db_host = os.environ.get("DB_HOST", "10.10.125.91")
    db_port = os.environ.get("DB_PORT", "5432")
    db_name = os.environ.get("DB_NAME", "fsi_db")
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
            con.execute(text('DELETE FROM "SUSPENSE_DATA_SUSPENSE_V2" WHERE "CEDANT NAME" = \'PT LIPPO GENERAL INSURANCE\';'))
        df_db.to_sql("SUSPENSE_DATA_SUSPENSE_V2", con=engine, if_exists="append", index=False)
        print(f"      -> {len(df_db):,} rows berhasil di-export ke 'SUSPENSE_DATA_SUSPENSE_V2'")
    except Exception as e:
        print(f"  [WARN] PostgreSQL export gagal: {e}")

    elapsed = time.perf_counter() - t_start
    print(f"\n{'=' * 60}")
    print(f"  Selesai dalam {elapsed:.1f}s ({elapsed/60:.1f} menit)")
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