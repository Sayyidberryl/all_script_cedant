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

SUSPEND_FILE = os.path.join("data", "kbru_output_suspend.xlsx")
OSBAL_FILE   = os.path.join("data", "kbru_output_osbal.xlsx")
FACUL_FILE   = os.path.join("data", "kbru_output_facul.xlsx")

OSBAL_FACODE_COL = "CCOS_REF_CODE"
FACUL_FACODE_COL = "FAC_CODE"

SUSPEND_CURR_COL = "CURR ORI"
OSBAL_CURR_COL   = "CCOS_CURR"
SUSPEND_DATE_COL = "RECEIPT DATE"
OSBAL_DATE_COL   = "FAC_COM_DATE"

_MIN_TOKEN_LEN = 3
_TOKEN_RE      = re.compile(r"[^A-Z0-9]+")
_DELIMITER_RE  = re.compile(r"[,;|+\s]+")

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
        if val and any(m in val for m in ["HUTANG PIUTANG", "DATA SUSPENSE"]):
            return True
    return False

def _collect_clean_values(row: dict, cols: list, is_slip: bool = False) -> list:
    seen, result = set(), []
    for col in cols:
        val = _normalize(row.get(col, ""))
        if val and val not in seen:
            seen.add(val)
            result.append(val)
            if is_slip:
                # Add base slip with -NN stripped
                base = re.sub(r'-\d+$', '', val)
                if base and base not in seen:
                    seen.add(base)
                    result.append(base)
                # Add clean alphanumeric
                clean_num = _clean_alphanumeric(val)
                if clean_num and clean_num not in seen:
                    seen.add(clean_num)
                    result.append(clean_num)
                clean_base = _clean_alphanumeric(base)
                if clean_base and clean_base not in seen:
                    seen.add(clean_base)
                    result.append(clean_base)
    return result

def _build_exact_index(rows: list, cols: list, is_slip: bool = False, excluded: set = None) -> dict:
    index: dict = {}
    for i, row in enumerate(rows):
        if excluded and i in excluded:
            continue
        vals = _collect_clean_values(row, cols, is_slip=is_slip)
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
    matched_osbal_idx = []
    for fac in fac_codes:
        if fac in facode_osbal_idx:
            matched_osbal_idx.extend(facode_osbal_idx[fac])
    return matched_osbal_idx, len(matched_osbal_idx) > 0

print("Testing prototype index building and matching...")
t0 = time.perf_counter()

df_sus = pd.read_excel(SUSPEND_FILE)
df_osbal = pd.read_excel(OSBAL_FILE)
df_facul = pd.read_excel(FACUL_FILE)

suspend_rows = df_sus.to_dict("records")
osbal_rows   = df_osbal.to_dict("records")
facul_rows   = df_facul.to_dict("records")

# Pre-parse dates
for r in osbal_rows:
    raw = r.get(OSBAL_DATE_COL, "")
    r["_com_date_parsed"] = pd.to_datetime(raw, errors="coerce") if raw and not (isinstance(raw, float) and pd.isna(raw)) else None

for r in suspend_rows:
    raw = r.get(SUSPEND_DATE_COL, "")
    r["_sus_date_parsed"] = pd.to_datetime(raw, errors="coerce") if raw and not (isinstance(raw, float) and pd.isna(raw)) else None

polis_sus   = _get_clean_cols(df_sus.columns, "clean polis")
slip_sus    = _get_clean_cols(df_sus.columns, "clean slip")
insured_sus = _get_clean_cols(df_sus.columns, "clean insured")

polis_osbal   = _get_clean_cols(df_osbal.columns, "clean polis")
slip_osbal    = _get_clean_cols(df_osbal.columns, "clean slip")
insured_osbal = _get_clean_cols(df_osbal.columns, "clean insured")

# Prepend CLSDT
if "CLSDT_POLICY_NO" in df_osbal.columns and "CLSDT_POLICY_NO" not in polis_osbal:
    polis_osbal = ["CLSDT_POLICY_NO"] + polis_osbal
if "CLSDT_SLIP_NO" in df_osbal.columns and "CLSDT_SLIP_NO" not in slip_osbal:
    slip_osbal = ["CLSDT_SLIP_NO"] + slip_osbal

polis_facul   = _get_clean_cols(df_facul.columns, "clean polis")
slip_facul    = _get_clean_cols(df_facul.columns, "clean slip")
insured_facul = _get_clean_cols(df_facul.columns, "clean insured")

excluded_osbal = {i for i, r in enumerate(osbal_rows) if _is_excluded_ref_row(r, polis_osbal, slip_osbal)}
excluded_facul = {i for i, r in enumerate(facul_rows) if _is_excluded_ref_row(r, polis_facul, slip_facul)}

facode_osbal_idx = _build_facode_index(osbal_rows, OSBAL_FACODE_COL, excluded=excluded_osbal)

# Build lookups
# [0]: slip exact, [1]: polis exact, [2]: insured exact, [3]: tok_slip, [4]: tok_polis, [5]: tok_ins
lookup_osbal = (
    _build_exact_index(osbal_rows, slip_osbal, is_slip=True, excluded=excluded_osbal),
    _build_exact_index(osbal_rows, polis_osbal, is_slip=False, excluded=excluded_osbal),
    _build_exact_index(osbal_rows, insured_osbal, is_slip=False, excluded=excluded_osbal),
    _build_token_index(osbal_rows, slip_osbal, excluded=excluded_osbal),
    _build_token_index(osbal_rows, polis_osbal, excluded=excluded_osbal),
    _build_token_index(osbal_rows, insured_osbal, excluded=excluded_osbal),
)

lookup_facul = (
    _build_exact_index(facul_rows, slip_facul, is_slip=True, excluded=excluded_facul),
    _build_exact_index(facul_rows, polis_facul, is_slip=False, excluded=excluded_facul),
    _build_exact_index(facul_rows, insured_facul, is_slip=False, excluded=excluded_facul),
    _build_token_index(facul_rows, slip_facul, excluded=excluded_facul),
    _build_token_index(facul_rows, polis_facul, excluded=excluded_facul),
    _build_token_index(facul_rows, insured_facul, excluded=excluded_facul),
)

print(f"Indexes built in {time.perf_counter()-t0:.2f}s")

# Test matching loop
scenarios = {}
sources = {}
match_count = 0

for sus in suspend_rows:
    sus_date = sus.get("_sus_date_parsed")
    sus_curr = _normalize(sus.get(SUSPEND_CURR_COL, ""))

    matched_osbal_idx = set()
    source = None
    scenario = ""

    # FASE 1: SLIP (OSBAL -> FACUL)
    slip_vals = _collect_clean_values(sus, slip_sus, is_slip=True)
    if slip_vals:
        # OSBAL exact
        hits = _exact_match(slip_vals, lookup_osbal[0])
        if hits:
            matched_osbal_idx.update(hits)
            source = "OSBAL"
            scenario = "Slip only"
        else:
            # FACUL exact
            fac_hits = _exact_match(slip_vals, lookup_facul[0])
            if fac_hits:
                res_idx, ok = _resolve_facode([facul_rows[i] for i in fac_hits], FACUL_FACODE_COL, facode_osbal_idx, osbal_rows)
                if ok:
                    matched_osbal_idx.update(res_idx)
                    source = "FACUL"
                    scenario = "Slip only"
            if not matched_osbal_idx:
                # OSBAL LIKE
                like_hits = _like_match(slip_vals, lookup_osbal[3][0], osbal_rows, slip_osbal)
                if like_hits:
                    matched_osbal_idx.update(like_hits)
                    source = "OSBAL"
                    scenario = "Slip only"

    # FASE 2: POLIS (if not matched by slip)
    if not matched_osbal_idx:
        polis_vals = _collect_clean_values(sus, polis_sus, is_slip=False)
        if polis_vals:
            # OSBAL exact
            hits = _exact_match(polis_vals, lookup_osbal[1])
            if hits:
                matched_osbal_idx.update(hits)
                source = "OSBAL"
                scenario = "Polis only"
            else:
                # FACUL exact
                fac_hits = _exact_match(polis_vals, lookup_facul[1])
                if fac_hits:
                    res_idx, ok = _resolve_facode([facul_rows[i] for i in fac_hits], FACUL_FACODE_COL, facode_osbal_idx, osbal_rows)
                    if ok:
                        matched_osbal_idx.update(res_idx)
                        source = "FACUL"
                        scenario = "Polis only"
                if not matched_osbal_idx:
                    # OSBAL LIKE
                    like_hits = _like_match(polis_vals, lookup_osbal[4][0], osbal_rows, polis_osbal)
                    if like_hits:
                        matched_osbal_idx.update(like_hits)
                        source = "OSBAL"
                        scenario = "Polis only"

    # FASE 3: INSURED (if not matched by slip or polis)
    if not matched_osbal_idx:
        ins_vals = _collect_clean_values(sus, insured_sus, is_slip=False)
        if ins_vals:
            # OSBAL exact
            hits = _exact_match(ins_vals, lookup_osbal[2])
            if hits:
                matched_osbal_idx.update(hits)
                source = "OSBAL"
                scenario = "Insured only"
            else:
                # FACUL exact
                fac_hits = _exact_match(ins_vals, lookup_facul[2])
                if fac_hits:
                    res_idx, ok = _resolve_facode([facul_rows[i] for i in fac_hits], FACUL_FACODE_COL, facode_osbal_idx, osbal_rows)
                    if ok:
                        matched_osbal_idx.update(res_idx)
                        source = "FACUL"
                        scenario = "Insured only"
                if not matched_osbal_idx:
                    # OSBAL LIKE
                    like_hits = _like_match(ins_vals, lookup_osbal[5][0], osbal_rows, insured_osbal)
                    if like_hits:
                        matched_osbal_idx.update(like_hits)
                        source = "OSBAL"
                        scenario = "Insured only"

    # FASE 4: Narrowing Wajib (Currency & Periode)
    if matched_osbal_idx:
        matched_list = list(matched_osbal_idx)
        # Currency
        if sus_curr:
            curr_filtered = [i for i in matched_list if _normalize(osbal_rows[i].get(OSBAL_CURR_COL, "")) == sus_curr]
            if curr_filtered:
                matched_list = curr_filtered
        # Periode (Basic rule: sus_date >= com_date, no lineslip offset)
        if pd.notna(sus_date):
            per_filtered = []
            for i in matched_list:
                com_date = osbal_rows[i].get("_com_date_parsed")
                if pd.isna(com_date) or sus_date >= com_date:
                    per_filtered.append(i)
            if per_filtered:
                matched_list = per_filtered
        matched_osbal_idx = set(matched_list)

    # FASE 5: Narrowing Utama (Slip <-> Polis <-> Insured)
    if len(matched_osbal_idx) > 1:
        matched_list = list(matched_osbal_idx)
        if "Slip" in scenario:
            # Narrow by Polis
            polis_vals = _collect_clean_values(sus, polis_sus, is_slip=False)
            if polis_vals:
                conf = [i for i in matched_list if any(_normalize(osbal_rows[i].get(c, "")) in polis_vals for c in polis_osbal)]
                if conf:
                    matched_list = conf
                    scenario = "Slip + Polis"
            # Narrow by Insured
            ins_vals = _collect_clean_values(sus, insured_sus, is_slip=False)
            if len(matched_list) > 1 and ins_vals:
                conf = [i for i in matched_list if any(_normalize(osbal_rows[i].get(c, "")) in ins_vals for c in insured_osbal)]
                if conf:
                    matched_list = conf
                    scenario = scenario + " + Insured"
        elif "Polis" in scenario:
            # Narrow by Slip
            slip_vals = _collect_clean_values(sus, slip_sus, is_slip=True)
            if slip_vals:
                conf = [i for i in matched_list if any(_normalize(osbal_rows[i].get(c, "")) in slip_vals for c in slip_osbal)]
                if conf:
                    matched_list = conf
                    scenario = "Polis + Slip"
            # Narrow by Insured
            ins_vals = _collect_clean_values(sus, insured_sus, is_slip=False)
            if len(matched_list) > 1 and ins_vals:
                conf = [i for i in matched_list if any(_normalize(osbal_rows[i].get(c, "")) in ins_vals for c in insured_osbal)]
                if conf:
                    matched_list = conf
                    scenario = scenario + " + Insured"
        elif "Insured" in scenario:
            # Narrow by Slip
            slip_vals = _collect_clean_values(sus, slip_sus, is_slip=True)
            if slip_vals:
                conf = [i for i in matched_list if any(_normalize(osbal_rows[i].get(c, "")) in slip_vals for c in slip_osbal)]
                if conf:
                    matched_list = conf
                    scenario = "Insured + Slip"
            # Narrow by Polis
            polis_vals = _collect_clean_values(sus, polis_sus, is_slip=False)
            if len(matched_list) > 1 and polis_vals:
                conf = [i for i in matched_list if any(_normalize(osbal_rows[i].get(c, "")) in polis_vals for c in polis_osbal)]
                if conf:
                    matched_list = conf
                    scenario = "Insured + Polis" if "Slip" not in scenario else scenario + " + Polis"
        matched_osbal_idx = set(matched_list)

    if not matched_osbal_idx:
        scenario = "Unmatching"
        source = None
    else:
        match_count += 1

    scenarios[scenario] = scenarios.get(scenario, 0) + 1
    src_key = source or "NONE"
    sources[src_key] = sources.get(src_key, 0) + 1

print("\n--- MATCHING RESULTS SUMMARY ---")
print(f"Total Suspend rows: {len(suspend_rows)}")
print(f"Total Matched:      {match_count} ({match_count/len(suspend_rows)*100:.1f}%)")
print(f"Unmatched:          {scenarios.get('Unmatching', 0)}")
print("\nBy Scenario:")
for k, v in sorted(scenarios.items(), key=lambda x: -x[1]):
    print(f"  {k:<35}: {v:,}")
print("\nBy Source:")
for k, v in sorted(sources.items(), key=lambda x: -x[1]):
    print(f"  {k:<35}: {v:,}")
