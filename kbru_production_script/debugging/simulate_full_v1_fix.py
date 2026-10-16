import os, sys, re, time
sys.modules['numexpr'] = None
sys.modules['bottleneck'] = None
import pandas as pd
import numpy as np

DATA_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "data")
SUSPEND_FILE = os.path.join(DATA_DIR, "kbru_output_suspend.xlsx")
OSBAL_FILE   = os.path.join(DATA_DIR, "kbru_output_osbal.xlsx")
FACUL_FILE   = os.path.join(DATA_DIR, "kbru_output_facul.xlsx")

SUSPEND_CURR_COL = "CURR ORI"
OSBAL_CURR_COL   = "CCOS_CURR"
SUSPEND_DATE_COL = "RECEIPT DATE"
OSBAL_DATE_COL   = "FAC_COM_DATE"

OSBAL_FACODE_COL = "CCOS_REF_CODE"
FACUL_FACODE_COL = "FAC_CODE"

_DUMMY_VALUES = frozenset({"-", "--", "---", "0", "00", "000", "NONE", "NULL", "TBA", "NA", "N/A", "SELISIH", "NO", "NON", "."})

def _normalize(value) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    text = str(value).strip().upper().replace("S/D", "SD")
    return re.sub(r"\s+", " ", text) if "  " in text else text

def _clean_alphanumeric(s: str) -> str:
    if not s or (isinstance(s, float) and pd.isna(s)):
        return ""
    return re.sub(r"[^A-Z0-9]", "", str(s).upper())

def _is_valid_token(s: str) -> bool:
    if not s or (isinstance(s, float) and pd.isna(s)):
        return False
    norm = _normalize(s)
    if not norm or norm in _DUMMY_VALUES:
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
        no_broker = re.sub(r'KBRURE|SRB', '', c)
        if len(no_broker) >= 6:
            variants.add(no_broker)

    # 2. Pola registrasi polis KBRU
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
        raw = row.get(col, "")
        if not _is_valid_token(raw):
            continue
        val = _normalize(raw)
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

def _ref_has_value(ref_row: dict, clean_cols: list, ori_col: str, query_values: list, is_slip: bool = False, is_polis: bool = False) -> bool:
    if not query_values:
        return False
    q_set = set(query_values)
    ref_cols = list(clean_cols)
    if ori_col and ori_col not in ref_cols:
        ref_cols.append(ori_col)
    ref_vals = _collect_clean_values(ref_row, ref_cols, is_slip=is_slip, is_polis=is_polis)
    return any(v in q_set for v in ref_vals)

# Load data
suspend_df = pd.read_excel(SUSPEND_FILE)
osbal_df   = pd.read_excel(OSBAL_FILE)
facul_df   = pd.read_excel(FACUL_FILE)

suspend_rows = suspend_df.to_dict("records")
osbal_rows   = osbal_df.to_dict("records")
facul_rows   = facul_df.to_dict("records")

for r in osbal_rows:
    raw = r.get(OSBAL_DATE_COL, "")
    r["_com_date_parsed"] = pd.to_datetime(raw, errors="coerce") if raw and not (isinstance(raw, float) and pd.isna(raw)) else None

for r in suspend_rows:
    raw = r.get(SUSPEND_DATE_COL, "")
    r["_sus_date_parsed"] = pd.to_datetime(raw, errors="coerce") if raw and not (isinstance(raw, float) and pd.isna(raw)) else None

def _get_clean_cols(columns, prefix: str) -> list:
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

polis_sus   = _get_clean_cols(suspend_df.columns, "clean polis")
slip_sus    = _get_clean_cols(suspend_df.columns, "clean slip")
insured_sus = _get_clean_cols(suspend_df.columns, "clean insured")

for c in ["POLIS_ORI", "POLIS"]:
    if c in suspend_df.columns and c not in polis_sus:
        polis_sus.append(c)
for c in ["SLIP_NO_ORI", "SLIP_NO"]:
    if c in suspend_df.columns and c not in slip_sus:
        slip_sus.append(c)
if "INSURED_ORI" in suspend_df.columns and "INSURED_ORI" not in insured_sus:
    insured_sus.append("INSURED_ORI")

polis_osbal   = _get_clean_cols(osbal_df.columns, "clean polis")
slip_osbal    = _get_clean_cols(osbal_df.columns, "clean slip")
insured_osbal = _get_clean_cols(osbal_df.columns, "clean insured")

for c in ["CLSDT_POLICY_NO", "FAC_POLICY_NO_ORI", "FAC_SLIP_ORI"]:
    if c in osbal_df.columns and c not in polis_osbal:
        polis_osbal.append(c)
for c in ["CLSDT_SLIP_NO", "FAC_SLIP_ORI"]:
    if c in osbal_df.columns and c not in slip_osbal:
        slip_osbal.append(c)

polis_facul   = _get_clean_cols(facul_df.columns, "clean polis")
slip_facul    = _get_clean_cols(facul_df.columns, "clean slip")
insured_facul = _get_clean_cols(facul_df.columns, "clean insured")

for c in ["FAC_POLICY_NO"]:
    if c in facul_df.columns and c not in polis_facul:
        polis_facul.append(c)
for c in ["FAC_SLIP"]:
    if c in facul_df.columns and c not in slip_facul:
        slip_facul.append(c)

# Build Indexes
def _build_exact_index(rows, cols, is_slip=False, is_polis=False):
    idx = {}
    for i, r in enumerate(rows):
        vals = _collect_clean_values(r, cols, is_slip=is_slip, is_polis=is_polis)
        for v in vals:
            idx.setdefault(v, []).append(i)
    return idx

def _build_facode_index(rows, col):
    idx = {}
    for i, r in enumerate(rows):
        fc = _normalize(r.get(col, ""))
        if fc:
            idx.setdefault(fc, []).append(i)
    return idx

osbal_slip_idx = _build_exact_index(osbal_rows, slip_osbal, is_slip=True)
osbal_pol_idx  = _build_exact_index(osbal_rows, polis_osbal, is_polis=True)
osbal_ins_idx  = _build_exact_index(osbal_rows, insured_osbal)
facode_osbal_idx = _build_facode_index(osbal_rows, OSBAL_FACODE_COL)

facul_slip_idx = _build_exact_index(facul_rows, slip_facul, is_slip=True)
facul_pol_idx  = _build_exact_index(facul_rows, polis_facul, is_polis=True)
facul_ins_idx  = _build_exact_index(facul_rows, insured_facul)

def _resolve_facode(ref_rows_matched, facode_col):
    fac_codes = { _normalize(r.get(facode_col, "")) for r in ref_rows_matched if _normalize(r.get(facode_col, "")) }
    res = set()
    for fc in fac_codes:
        if fc in facode_osbal_idx:
            res.update(facode_osbal_idx[fc])
    return list(res), bool(res)

print("Index ready. Running match on 1833 rows...")

raw_results = []
for n, sus in enumerate(suspend_rows, 1):
    if str(sus.get("STATUS", "")).strip().upper() == "MATCHING":
        continue

    effective_date = sus.get("_sus_date_parsed")
    sus_curr = _normalize(sus.get(SUSPEND_CURR_COL, ""))

    matched_osbal_idx = set()
    source = None
    scenario = ""

    # FASE 1: SLIP (Prioritas 1)
    slip_vals = _collect_clean_values(sus, slip_sus, is_slip=True)
    if slip_vals:
        hits = {i for v in slip_vals if v in osbal_slip_idx for i in osbal_slip_idx[v]}
        if hits:
            matched_osbal_idx.update(hits)
            source = "OSBAL"
            scenario = "Slip only"
        elif facul_rows:
            f_hits = {i for v in slip_vals if v in facul_slip_idx for i in facul_slip_idx[v]}
            if f_hits:
                res_idx, ok = _resolve_facode([facul_rows[i] for i in f_hits], FACUL_FACODE_COL)
                if ok:
                    matched_osbal_idx.update(res_idx)
                    source = "FACUL"
                    scenario = "Slip only"

    # FASE 2: POLIS (Prioritas 2)
    if not matched_osbal_idx:
        polis_vals = _collect_clean_values(sus, polis_sus, is_polis=True)
        if polis_vals:
            hits = {i for v in polis_vals if v in osbal_pol_idx for i in osbal_pol_idx[v]}
            if hits:
                matched_osbal_idx.update(hits)
                source = "OSBAL"
                scenario = "Polis only"
            elif facul_rows:
                f_hits = {i for v in polis_vals if v in facul_pol_idx for i in facul_pol_idx[v]}
                if f_hits:
                    res_idx, ok = _resolve_facode([facul_rows[i] for i in f_hits], FACUL_FACODE_COL)
                    if ok:
                        matched_osbal_idx.update(res_idx)
                        source = "FACUL"
                        scenario = "Polis only"

    # FASE 3: INSURED (Prioritas 3)
    if not matched_osbal_idx:
        ins_vals = _collect_clean_values(sus, insured_sus)
        if ins_vals:
            hits = {i for v in ins_vals if v in osbal_ins_idx for i in osbal_ins_idx[v]}
            if hits:
                matched_osbal_idx.update(hits)
                source = "OSBAL"
                scenario = "Insured only"
            elif facul_rows:
                f_hits = {i for v in ins_vals if v in facul_ins_idx for i in facul_ins_idx[v]}
                if f_hits:
                    res_idx, ok = _resolve_facode([facul_rows[i] for i in f_hits], FACUL_FACODE_COL)
                    if ok:
                        matched_osbal_idx.update(res_idx)
                        source = "FACUL"
                        scenario = "Insured only"

    # FASE 4: Narrowing Currency & Periode
    is_beda_curr = False
    is_beda_per = False

    if matched_osbal_idx:
        matched_list = list(matched_osbal_idx)
        if sus_curr:
            fc = [i for i in matched_list if _normalize(osbal_rows[i].get(OSBAL_CURR_COL, "")) == sus_curr]
            if not fc:
                is_beda_curr = True
            elif len(fc) < len(matched_list):
                matched_list = fc

        if not is_beda_curr and pd.notna(effective_date):
            fp = []
            for i in matched_list:
                cd = osbal_rows[i].get("_com_date_parsed")
                if pd.isna(cd) or effective_date >= cd:
                    fp.append(i)
            if not fp:
                is_beda_per = True
            elif len(fp) < len(matched_list):
                matched_list = fp

        matched_osbal_idx = set(matched_list)

    # FASE 5: Narrowing Non-Destruktif
    if len(matched_osbal_idx) > 1:
        matched_list = list(matched_osbal_idx)
        if "Slip" in scenario:
            p_vals = _collect_clean_values(sus, polis_sus, is_polis=True)
            if p_vals:
                conf = [i for i in matched_list if _ref_has_value(osbal_rows[i], polis_osbal, "FAC_POLICY_NO_ORI", p_vals, is_polis=True)]
                if conf:
                    matched_list = conf
                    scenario = "Slip + Polis"
            ins_vals = _collect_clean_values(sus, insured_sus)
            if len(matched_list) > 1 and ins_vals:
                conf = [i for i in matched_list if _ref_has_value(osbal_rows[i], insured_osbal, "FAC_INSURED_ORI", ins_vals)]
                if conf:
                    matched_list = conf
                    scenario = scenario + " + Insured"

        elif "Polis" in scenario:
            s_vals = _collect_clean_values(sus, slip_sus, is_slip=True)
            if s_vals:
                conf = [i for i in matched_list if _ref_has_value(osbal_rows[i], slip_osbal, "FAC_SLIP_ORI", s_vals, is_slip=True)]
                if conf:
                    matched_list = conf
                    scenario = "Polis + Slip"
            ins_vals = _collect_clean_values(sus, insured_sus)
            if len(matched_list) > 1 and ins_vals:
                conf = [i for i in matched_list if _ref_has_value(osbal_rows[i], insured_osbal, "FAC_INSURED_ORI", ins_vals)]
                if conf:
                    matched_list = conf
                    scenario = scenario + " + Insured"

        elif "Insured" in scenario:
            s_vals = _collect_clean_values(sus, slip_sus, is_slip=True)
            if s_vals:
                conf = [i for i in matched_list if _ref_has_value(osbal_rows[i], slip_osbal, "FAC_SLIP_ORI", s_vals, is_slip=True)]
                if conf:
                    matched_list = conf
                    scenario = "Insured + Slip"
            p_vals = _collect_clean_values(sus, polis_sus, is_polis=True)
            if len(matched_list) > 1 and p_vals:
                conf = [i for i in matched_list if _ref_has_value(osbal_rows[i], polis_osbal, "FAC_POLICY_NO_ORI", p_vals, is_polis=True)]
                if conf:
                    matched_list = conf
                    scenario = scenario + " + Polis"

        matched_osbal_idx = set(matched_list)

    # Resolusi
    fac_codes = {str(osbal_rows[i].get("CCOS_REF_CODE", "")).strip() for i in matched_osbal_idx if str(osbal_rows[i].get("CCOS_REF_CODE", "")).strip()}
    num_fac = len(fac_codes)

    if not source or not scenario or scenario == "Unmatching" or not matched_osbal_idx:
        final_scenario = "Unmatching"
    elif is_beda_curr:
        final_scenario = f"{scenario} (Beda Currency)"
    elif is_beda_per:
        final_scenario = f"{scenario} (Beda Periode)"
    elif num_fac > 1:
        final_scenario = f"{scenario} (> 1 Fac)"
    else:
        final_scenario = scenario

    raw_results.append({
        "scenario": final_scenario,
        "num_fac": num_fac,
        "source": source if final_scenario != "Unmatching" else None
    })

res_df = pd.DataFrame(raw_results)
print("\n" + "="*60)
print("HASIL SIMULASI PERBAIKAN V1:")
print("="*60)
print("\nSKENARIO Counts:")
print(res_df["scenario"].value_counts())
