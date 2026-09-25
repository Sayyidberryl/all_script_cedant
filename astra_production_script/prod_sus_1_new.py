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
# CONFIG
# =============================================================================

# -----------------------------------------------------------------------
# PT ASURANSI ASTRA BUANA — file paths & column config
# -----------------------------------------------------------------------
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = "data" if os.path.exists("data") else os.path.join(BASE_DIR, "data")

SUSPEND_FILE = os.path.join(DATA_DIR, "astrabuana_output_suspend.xlsx")
OSBAL_FILE   = os.path.join(DATA_DIR, "astrabuana_output_osbal.xlsx")
FACUL_FILE   = os.path.join(DATA_DIR, "astrabuana_output_facul.xlsx")
OUTPUT_FILE  = os.path.join(DATA_DIR, "final_output_astrabuana_v1.xlsx")

# File Bordero (Open Cover) — memuat FIRE, MC, dan MH
BORDERO_DIR        = os.path.join(DATA_DIR, "BORDERO")
BORDERO_FIRE_FILE  = os.path.join(BORDERO_DIR, "ASTRA_OPEN_COVER_FIRE.xlsx")
BORDERO_MC_FILE    = os.path.join(BORDERO_DIR, "ASTRA_OPEN_COVER_MC.xlsx")
BORDERO_MH_FILE    = os.path.join(BORDERO_DIR, "ASTRA_OPEN_COVER_MH.xlsx")
BORDERO_FILES      = [BORDERO_FIRE_FILE, BORDERO_MC_FILE, BORDERO_MH_FILE]

OSBAL_FACODE_COL  = "CCOS_REF_CODE"
FACUL_FACODE_COL  = "FAC_CODE"

# Kolom di tabel Bordero (Open Cover)
BORDERO_FAC_COL     = "FAC CODE"
BORDERO_POLIS_COL   = "POLIS"
BORDERO_SLIP_COL    = "SLIP"
BORDERO_CERT_COL    = "CERTIFICATE"
BORDERO_INSURED_COL = "INSURED"
BORDERO_CURR_COL    = "CURR"
BORDERO_BULAN_COL   = "BULAN"
BORDERO_TAHUN_COL   = "TAHUN"
BORDERO_NET_COL     = "NET"

# Mode & toleransi matching AMOUNT ORI vs NET bordero (last choice)
NET_MATCH_MODE    = "per_row"   # pilihan: 'per_row' | 'sum_all' | 'sum_by_curr'
NET_TOLERANCE_PCT = 0.0

# Mapping bulan Indonesia ke nomor
_BULAN_MAP = {
    "JANUARI": 1, "FEBRUARI": 2, "MARET": 3, "APRIL": 4,
    "MEI": 5, "JUNI": 6, "JULI": 7, "AGUSTUS": 8,
    "SEPTEMBER": 9, "OKTOBER": 10, "NOVEMBER": 11, "DESEMBER": 12,
}

# Regex ekstrak nomor sertifikat 6-digit dari polis (contoh: '...155-001007' -> '001007')
_CERT_RE = re.compile(r'-([0-9]{6})(?:[^0-9]|$)')

# -----------------------------------------------------------------------
# Kolom "ori" (raw) di tabel Suspend — Astra Buana
# -----------------------------------------------------------------------
SUS_INSURED_ORI_COL = "INSURED"
SUS_POLIS_ORI_COL   = "POLIS"
SUS_SLIP_ORI_COL    = "SLIP NO"

# Kolom CLSDT / FAC prioritas di Suspend
CLSDT_POLIS_COL = "CLSDT_POLICY_NO"
CLSDT_SLIP_COL  = "CLSDT_SLIP_NO"
FAC_POLIS_COL   = "FAC_POLICY_NO"
FAC_SLIP_COL    = "FAC_SLIP"

# -----------------------------------------------------------------------
# Prefix kolom clean per tabel — Astra Buana
# -----------------------------------------------------------------------
_SUS_INSURED_PREFIX  = "INSURED_CLEAN"
_SUS_POLIS_PREFIX    = "POLIS_CLEAN"
_SUS_SLIP_PREFIX     = "SLIP_NO_CLEAN"
_SUS_SERTIF_PREFIX   = "CERTIFICATE_"

_OSBAL_INSURED_PREFIX  = "FAC_INSURED_CLN_"
_OSBAL_POLIS_PREFIX    = "FAC_POLICY_NO_CLEAN_"
_OSBAL_SLIP_PREFIX     = "FAC_SLIP_CLEAN_"
_OSBAL_SERTIF_PREFIX   = "CERTIFICATE_"

# FACUL: prefix polis berbeda (FAC_POLICY_CLEAN_ bukan FAC_POLICY_NO_CLEAN_)
_FACUL_INSURED_PREFIX  = "FAC_INSURED_CLN_"
_FACUL_POLIS_PREFIX    = "FAC_POLICY_CLEAN_"
_FACUL_SLIP_PREFIX     = "FAC_SLIP_CLEAN_"
_FACUL_SERTIF_PREFIX   = "CERTIFICATE_"

# Kolom currency dan periode
SUSPEND_CURR_COL = "CURR ORI"
OSBAL_CURR_COL   = "CCOS_CURR"
SUSPEND_DATE_COL = "RECEIPT DATE"
OSBAL_DATE_COL   = "FAC_COM_DATE"

# Marker baris administratif yang dikecualikan dari matching
_EXCLUDED_REF_MARKERS = ["HUTANG PIUTANG", "DATA SUSPENSE"]

# Marker LINESLIP — pakai RECEIPT DATE - 1 bulan sebagai tanggal efektif
_LINESLIP_MARKERS = ("LINESLIP", "LINE SLIP")

# Prefix kolom bersih
SUSPEND_SERTIF_PREFIX = _SUS_SERTIF_PREFIX
OSBAL_SERTIF_PREFIX   = _OSBAL_SERTIF_PREFIX
FACUL_SERTIF_PREFIX   = _FACUL_SERTIF_PREFIX

_MIN_TOKEN_LEN = 3
_TOKEN_RE      = re.compile(r"[^A-Z0-9]+")
_DELIMITER_RE  = re.compile(r"[,;|+\s]+")

# Kolom suspend yang di-join saat akumulasi (v1: digunakan untuk shared amount)
_SUSPEND_JOIN_COLS = [
    "RECEIPT NO", "CREDIT NOTES", "DETAIL RINCIAN NO", "RECEIPT DATE",
    "CEDANT NAME", "CEDANT SHRT NAME",
    "INSURED", "INSURED_CLEAN",
    "CURR ORI", "CURR PAY", "AMOUNT PAY",
    "POLIS", "POLIS_CLEAN",
    "SLIP NO", "SLIP_NO_CLEAN",
    "DESC 1", "DESC 2", "DESC 3", "DESC 4",
    "STATUS", "REC_TYPE",
]
_SUSPEND_SUM_COLS = ["AMOUNT ORI"]

# V1: output 38 kolom standar alternating Polis & Sertifikat, row asli dipertahankan
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
# HELPER FUNCTIONS: FILE DISCOVERY
# =============================================================================

def _find_bordero_files() -> list:
    """Mencari semua file Bordero (FIRE, MC, MH) di folder data/BORDERO atau fallback."""
    files = []
    for f in BORDERO_FILES:
        if os.path.exists(f) and f not in files:
            files.append(f)

    for folder_name in ["BORDERO", "bordero", "Bordero"]:
        b_dir = os.path.join(DATA_DIR, folder_name)
        if os.path.exists(b_dir) and os.path.isdir(b_dir):
            for fname in sorted(os.listdir(b_dir)):
                if fname.lower().endswith((".xlsx", ".xls")) and not fname.startswith("~$"):
                    fpath = os.path.join(b_dir, fname)
                    if fpath not in files:
                        files.append(fpath)
            if files:
                return files

    if files:
        return files

    # Fallback jika file ditaruh langsung di data/
    candidates = []
    if os.path.exists(DATA_DIR):
        for fname in os.listdir(DATA_DIR):
            if ("bordero" in fname.lower() or "open_cover" in fname.lower()) and fname.lower().endswith((".xlsx", ".xls")) and not fname.startswith("~$"):
                candidates.append(os.path.join(DATA_DIR, fname))
        candidates.sort()
    return candidates


# =============================================================================
# STRING NORMALIZATION & TOKENIZATION
# =============================================================================

@functools.lru_cache(maxsize=50000)
def _norm_cached(s: str) -> str:
    s = s.upper()
    s = re.sub(r'[\x00-\x1f\x7f-\x9f]', '', s)
    s = re.sub(r'[^A-Z0-9]', '', s)
    return s


def _normalize(value) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and pd.isna(value):
        return ""
    return _norm_cached(str(value).strip())


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


def _clean_cert_str(val) -> str:
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return ""
    s = str(val).strip()
    return re.sub(r'\.0+$', '', s)


def _cert_in_range(query_cert: str, ref_cert_str: str) -> bool:
    """Mengecek apakah nomor sertifikat berada di dalam range referensi (misal 2846 s/d 3466)."""
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
        # Rollover (misal 4598 - 18)
        return q_num >= start or q_num <= end
    return False


def _slip_in_range(sus_slip: str, osbal_row: dict) -> bool:
    """Mengecek apakah nomor slip suspend berada di dalam range slip OSBAL (antara slip 1 dan slip 2)."""
    if not sus_slip:
        return False
    s_clean = re.sub(r'\D', '', str(sus_slip))
    if not s_clean:
        return False

    sl1_val = osbal_row.get("clean slip 1") or osbal_row.get("FAC_SLIP_CLEAN_1") or ""
    sl2_val = osbal_row.get("clean slip 2") or osbal_row.get("FAC_SLIP_CLEAN_2") or ""

    sl1 = re.sub(r'\D', '', str(sl1_val))
    sl2 = re.sub(r'\D', '', str(sl2_val))
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


@functools.lru_cache(maxsize=100000)
def _split_composite(text: str) -> frozenset:
    if not text:
        return frozenset()
    return frozenset(tok.strip() for tok in _DELIMITER_RE.split(text) if tok.strip())


def _get_clean_cols(columns: list, prefix: str) -> list:
    if not columns:
        return []
    pfx = prefix.strip().lower()
    matched = [c for c in columns if c.lower().startswith(pfx)]

    def _sort_key(name: str):
        suffix = name[len(prefix):].strip(" _")
        return (0, int(suffix)) if suffix.isdigit() else (1, name)

    return sorted(matched, key=_sort_key)


def _map_scenario(label: str) -> str:
    m = {
        "POLIS_EXACT": "Polis only",
        "POLIS_LIKE":  "Polis only",
        "SLIP_EXACT":  "Slip only",
        "SLIP_LIKE":   "Slip only",
        "INSURED_EXACT": "Insured only",
        "INSURED_LIKE":  "Insured only",
    }
    return m.get(label, label)


def _format_sertif(val) -> str:
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return ""
    s = str(val).strip()
    s = re.sub(r'\.0+$', '', s)
    return s.zfill(6) if s.isdigit() and len(s) < 6 else s


def _extract_cert_from_polis(polis_val: str) -> str:
    if not polis_val:
        return ""
    m = _CERT_RE.search(str(polis_val))
    return m.group(1) if m else ""


# =============================================================================
# ROW FILTERING & VALUE COLLECTION
# =============================================================================

def _is_excluded_ref_row(row: dict, polis_cols: list, slip_cols: list) -> bool:
    for col in (polis_cols or []) + (slip_cols or []):
        val = str(row.get(col, "") or "").upper()
        if any(marker in val for marker in _EXCLUDED_REF_MARKERS):
            return True
    return False


def _is_lineslip_row(sus_row: dict, insured_cols: list) -> bool:
    for col in insured_cols:
        val = str(sus_row.get(col, "") or "").upper()
        if any(marker in val for marker in _LINESLIP_MARKERS):
            return True
    return False


def _collect_clean_values(row: dict, cols: list) -> list:
    if not cols:
        return []
    result = []
    for c in cols:
        v = _normalize(row.get(c, ""))
        if v and v not in result:
            result.append(v)
    return result


def _collect_all_values(row: dict, clean_cols: list, ori_col: str) -> list:
    vals = _collect_clean_values(row, clean_cols)
    ori = _normalize(row.get(ori_col, ""))
    if ori and ori not in vals:
        vals.append(ori)
    return vals


def _get_effective_sus_cols(sus_row: dict, base_polis_cols: list, base_slip_cols: list) -> tuple:
    clsdt_p = _normalize(sus_row.get(CLSDT_POLIS_COL, ""))
    fac_p   = _normalize(sus_row.get(FAC_POLIS_COL,   ""))
    eff_polis  = [CLSDT_POLIS_COL] if clsdt_p else base_polis_cols
    like_polis = [CLSDT_POLIS_COL] if clsdt_p else (base_polis_cols + ([FAC_POLIS_COL] if fac_p else []))

    clsdt_s = _normalize(sus_row.get(CLSDT_SLIP_COL, ""))
    fac_s   = _normalize(sus_row.get(FAC_SLIP_COL,   ""))
    eff_slip  = [CLSDT_SLIP_COL] if clsdt_s else base_slip_cols
    like_slip = [CLSDT_SLIP_COL] if clsdt_s else (base_slip_cols + ([FAC_SLIP_COL] if fac_s else []))

    return eff_polis, eff_slip, like_polis, like_slip


def _ref_has_value(ref_row: dict, clean_cols: list, ori_col: str, query_values: list) -> bool:
    if not query_values:
        return False
    ref_vals = set(_collect_clean_values(ref_row, clean_cols))
    if set(query_values) & ref_vals:
        return True

    for rv in ref_vals:
        if any(v in _split_composite(rv) for v in query_values if v and len(v) >= 2):
            return True

    return any(
        v in rv
        for v in query_values if v and len(v) >= 10
        for rv in ref_vals if rv
    )


# =============================================================================
# INDEX BUILDERS
# =============================================================================

def _build_exact_index(rows: list, cols: list, excluded: set = None) -> dict:
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
    index: dict = {}
    for i, row in enumerate(rows):
        if excluded and i in excluded:
            continue
        val = _normalize(row.get(col, ""))
        if val:
            index.setdefault(val, []).append(i)
    return index


def _build_token_index(rows: list, clean_cols: list, ori_col: str, excluded: set = None) -> tuple:
    index: dict = {}
    for i, row in enumerate(rows):
        if excluded and i in excluded:
            continue
        seen_tokens: set = set()
        for col in clean_cols:
            val = _normalize(row.get(col, ""))
            if val:
                for tok in _tokenize(val):
                    if tok not in seen_tokens:
                        seen_tokens.add(tok)
                        index.setdefault(tok, set()).append(i) if isinstance(index.get(tok), list) else index.setdefault(tok, set()).add(i)
    return index, {}


def _build_sertif_index(rows: list, sertif_cols: list, excluded: set = None) -> dict:
    index: dict = {}
    for i, row in enumerate(rows):
        if excluded and i in excluded:
            continue
        for col in sertif_cols:
            val = row.get(col, "")
            for cert in _expand_sertif_range(str(val) if val is not None else ""):
                index.setdefault(cert, []).append(i)
    return index


def _build_lookup(rows: list, polis_cols: list, slip_cols: list,
                  insured_cols: list, facode_col: str = None, label: str = "") -> tuple:
    if not rows:
        if label:
            print(f"  {label:<6} done (0.0s)", flush=True)
        return ({}, {}, {}, {}, {}, {}, ({}, {}), ({}, {}), ({}, {}), {})

    t0 = time.perf_counter()
    excluded = {i for i, r in enumerate(rows) if _is_excluded_ref_row(r, polis_cols, slip_cols)}

    facode_index = _build_facode_index(rows, facode_col, excluded=excluded) if facode_col else {}
    lookup = (
        _build_exact_index(rows, slip_cols,    excluded=excluded),
        _build_exact_index(rows, polis_cols,   excluded=excluded),
        _build_exact_index(rows, insured_cols, excluded=excluded),
        {}, {}, {},
        _build_token_index(rows, slip_cols,    "slip_ori",    excluded=excluded),
        _build_token_index(rows, polis_cols,   "polis_ori",   excluded=excluded),
        _build_token_index(rows, insured_cols, "insured_ori", excluded=excluded),
        facode_index,
    )

    elapsed = time.perf_counter() - t0
    if label:
        exc_info = f" - {len(excluded):,} baris dikecualikan" if excluded else ""
        print(f"  {label:<6} done ({elapsed:.1f}s){exc_info}", flush=True)
    return lookup


# =============================================================================
# MATCH ENGINE
# =============================================================================

def _exact_match(query_values: list, index: dict) -> set:
    return {i for v in query_values if v and v in index for i in index[v]}


def _like_match(query_values: list, token_index: dict, rows: list,
                ref_cols: list, ori_col: str = None) -> list:
    valid_queries = [v for v in query_values if v and len(v) >= 2]
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


def _rematch_strict(
    suspend_row: dict, osbal_rows: list, lookup_osbal: tuple,
    polis_sus_cols: list, slip_sus_cols: list, insured_sus_cols: list,
    polis_ref_cols: list, slip_ref_cols: list, insured_ref_cols: list,
    facode_osbal_idx: dict,
    slipdb_rows: list = None, lookup_slipdb: tuple = None,
    facul_rows: list = None, lookup_facul: tuple = None,
    polis_sus_like_cols: list = None,
    effective_sus_date=None, is_lineslip: bool = False
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
                    if is_lineslip:
                        if com_date.year != effective_sus_date.year or com_date.month != effective_sus_date.month:
                            continue
                    else:
                        if effective_sus_date < com_date:
                            continue
            res.append(i)
        return res

    polis_vals = _collect_clean_values(suspend_row, polis_sus_cols)

    # R1: OSBAL polis exact
    curr_hits = _strict_filter(_exact_match(polis_vals, lookup_osbal[1]))
    if curr_hits:
        slip_vals = _collect_clean_values(suspend_row, slip_sus_cols)
        confirmed = [i for i in curr_hits if any(
            _normalize(osbal_rows[i].get(c, "")) in slip_vals
            for c in slip_ref_cols if _normalize(osbal_rows[i].get(c, ""))
        )]
        chosen = confirmed if confirmed else curr_hits
        return "OSBAL", ("Polis + Slip" if confirmed else "Polis only"), chosen, True

    # R2: OSBAL polis LIKE
    like_cols = polis_sus_like_cols if polis_sus_like_cols is not None else polis_sus_cols
    if lookup_osbal[7] and lookup_osbal[7][0] and like_cols:
        like_vals = _collect_clean_values(suspend_row, like_cols)
        if like_vals:
            curr_hits = _strict_filter(_like_match(like_vals, lookup_osbal[7][0], osbal_rows, polis_ref_cols))
            if curr_hits:
                return "OSBAL", "POLIS_LIKE", curr_hits, True

    # R3: SLIPDB (jika ada)
    if slipdb_rows and lookup_slipdb:
        hits = _exact_match(polis_vals, lookup_slipdb[1])
        if hits:
            res_idx, ok = _resolve_facode([slipdb_rows[i] for i in hits], FACUL_FACODE_COL, facode_osbal_idx, osbal_rows)
            if ok:
                curr_res = _strict_filter(res_idx)
                if curr_res:
                    return "SLIPDB", "Polis only", curr_res, True

    # R4: FACUL
    if facul_rows and lookup_facul:
        hits = _exact_match(polis_vals, lookup_facul[1])
        if hits:
            res_idx, ok = _resolve_facode([facul_rows[i] for i in hits], FACUL_FACODE_COL, facode_osbal_idx, osbal_rows)
            if ok:
                curr_res = _strict_filter(res_idx)
                if curr_res:
                    return "FACUL", "Polis only", curr_res, True

    # R5: OSBAL insured fallback
    curr_hits = _strict_filter(_exact_match(_collect_clean_values(suspend_row, insured_sus_cols), lookup_osbal[2]))
    if curr_hits:
        return "OSBAL", "Insured only", curr_hits, True

    return None


# =============================================================================
# BORDERO NARROWING & MULTI-FILE OVERRIDE (OPEN COVER)
# =============================================================================

def _build_bordero_index(bordero_rows: list) -> dict:
    """Bangun index: fac_code -> list entri per baris bordero dengan normalisasi kolom."""
    idx: dict = {}
    for raw_row in bordero_rows:
        row = {str(k).strip().upper(): v for k, v in raw_row.items()}
        fac = _normalize(row.get(BORDERO_FAC_COL, ""))
        if not fac:
            continue
        bulan_str = str(row.get(BORDERO_BULAN_COL, "")).strip().upper()
        tahun_str = str(row.get(BORDERO_TAHUN_COL, "")).strip()
        tahun_str = re.sub(r'\.0+$', '', tahun_str)
        month = _BULAN_MAP.get(bulan_str, 0)
        ym = (int(tahun_str), month) if month and tahun_str.isdigit() else None

        cert_val = _clean_cert_str(row.get(BORDERO_CERT_COL, ""))

        idx.setdefault(fac, []).append({
            "polis":   _normalize(row.get(BORDERO_POLIS_COL, "")),
            "slip":    _normalize(row.get(BORDERO_SLIP_COL, "")),
            "cert":    cert_val,
            "insured": _normalize(row.get(BORDERO_INSURED_COL, "")),
            "curr":    _normalize(row.get(BORDERO_CURR_COL, "")),
            "period":  ym,
            "net":     row.get(BORDERO_NET_COL),
        })
    return idx


def _find_fac_codes_in_bordero(
    suspend_row: dict, bordero_idx: dict,
    polis_sus_cols: list, slip_sus_cols: list,
    bordero_polis_index: dict = None,
    sertif_sus_cols: list = None,
) -> tuple:
    if not bordero_idx:
        return set(), False

    sus_polis = set(_collect_clean_values(suspend_row, polis_sus_cols))
    for c in [CLSDT_POLIS_COL, FAC_POLIS_COL, "polis_ori", SUS_POLIS_ORI_COL, "POLIS", "POLIS_CLEAN"]:
        v = _normalize(suspend_row.get(c, ""))
        if v:
            sus_polis.add(v)

    sus_cert = set()
    if sertif_sus_cols:
        for v in _collect_clean_values(suspend_row, sertif_sus_cols):
            cv = _clean_cert_str(v)
            if cv:
                sus_cert.add(cv)
    for pv in list(sus_polis) + [_normalize(suspend_row.get(SUS_POLIS_ORI_COL, ""))]:
        c = _extract_cert_from_polis(pv)
        if c:
            cv = _clean_cert_str(c)
            if cv:
                sus_cert.add(cv)

    sus_slip = set(_collect_clean_values(suspend_row, slip_sus_cols))
    for c in [CLSDT_SLIP_COL, FAC_SLIP_COL, "slip_ori", SUS_SLIP_ORI_COL, "SLIP NO", "SLIP_NO_CLEAN"]:
        v = _normalize(suspend_row.get(c, ""))
        if v:
            sus_slip.add(v)

    polis_f = {sp for sp in sus_polis if len(sp) >= 5}
    slip_f = {ss for ss in sus_slip if len(ss) >= 5}
    sus_amount = suspend_row.get("AMOUNT ORI")
    sus_amount_f = None
    if sus_amount is not None:
        try:
            s_amt = str(sus_amount).strip().replace(' ', '')
            if ',' in s_amt and '.' in s_amt:
                s_amt = s_amt.replace('.', '').replace(',', '.') if s_amt.rfind(',') > s_amt.rfind('.') else s_amt.replace(',', '')
            elif ',' in s_amt:
                s_amt = s_amt.replace(',', '.')
            sus_amount_f = float(s_amt)
        except (ValueError, TypeError):
            sus_amount_f = None

    sus_date = suspend_row.get("_sus_date_parsed")
    sus_ym = (sus_date.year, sus_date.month) if sus_date is not None else None
    sus_curr_val = _normalize(suspend_row.get(SUSPEND_CURR_COL, ""))

    def mp(e):
        return bool(e["polis"] and len(e["polis"]) >= 5 and
                    any(sp == e["polis"] or sp in e["polis"] or e["polis"] in sp for sp in polis_f))
    def mc(e):
        if not e.get("cert"):
            return False
        ec = e["cert"]
        if ec in sus_cert:
            return True
        return any(_cert_in_range(sc, ec) for sc in sus_cert)
    def ms(e):
        return bool(e["slip"] and len(e["slip"]) >= 5 and
                    any(ss == e["slip"] or ss in e["slip"] or e["slip"] in ss for ss in slip_f))
    def mnet(e):
        if sus_amount_f is None or e.get("net") is None:
            return False
        try:
            b_net = float(e["net"])
            return abs(abs(b_net) - abs(sus_amount_f)) < 0.05
        except Exception:
            return False
    def mper(e):
        eym = e["period"]
        return True if sus_ym is None or not eym else eym <= sus_ym
    def mcurr(e):
        ec = e.get("curr", "")
        return not sus_curr_val or not ec or ec == sus_curr_val

    candidates = {}

    if bordero_polis_index is not None:
        base_polis = {_normalize(_CERT_RE.split(pv)[0].rstrip('-') if _CERT_RE.search(pv) else pv)
                      for pv in polis_f if pv}
        matched_b_polis = set()

        # 1. Exact match
        for bp in (polis_f | base_polis):
            if bp in bordero_polis_index:
                matched_b_polis.add(bp)

        # 2. Jika tidak ada exact match, gunakan like match terbatas
        if not matched_b_polis:
            matched_b_polis = {bp for bp in bordero_polis_index for sp in polis_f if sp == bp or sp in bp or bp in sp}

        for bp in matched_b_polis:
            for fac, r in bordero_polis_index[bp]:
                candidates.setdefault(fac, []).append(r)
    else:
        for fac, rows in bordero_idx.items():
            for r in rows:
                if mp(r):
                    candidates.setdefault(fac, []).append(r)

    if not candidates:
        return set(), False

    if sus_cert:
        cert_filtered = {}
        for fac, rows in candidates.items():
            valid = [r for r in rows if mc(r)]
            if valid:
                cert_filtered[fac] = valid
        if cert_filtered:
            candidates = cert_filtered

    if sus_slip:
        slip_filtered = {}
        for fac, rows in candidates.items():
            valid = [r for r in rows if ms(r)]
            if valid:
                slip_filtered[fac] = valid
        if slip_filtered:
            candidates = slip_filtered

    wajib_filtered = {}
    for fac, rows in candidates.items():
        valid = [r for r in rows if mcurr(r) and mper(r)]
        if valid:
            wajib_filtered[fac] = valid
    if wajib_filtered:
        candidates = wajib_filtered

    if sus_amount_f is not None:
        net_filtered = {}
        for fac, rows in candidates.items():
            valid = [r for r in rows if mnet(r)]
            if valid:
                net_filtered[fac] = valid
        if net_filtered:
            candidates = net_filtered

    return set(candidates.keys()), True


# =============================================================================
# FAC CODE RESOLUTION
# =============================================================================

def _resolve_facode(ref_rows_matched: list, facode_col: str,
                    facode_osbal_idx: dict, osbal_rows: list) -> tuple:
    fac_codes = set()
    for r in ref_rows_matched:
        val = _normalize(r.get(facode_col, ""))
        if val:
            fac_codes.add(val)

    if not fac_codes:
        return [], False

    osbal_indices = []
    for fc in fac_codes:
        if fc in facode_osbal_idx:
            osbal_indices.extend(facode_osbal_idx[fc])

    return list(dict.fromkeys(osbal_indices)), len(osbal_indices) > 0


# =============================================================================
# OUTPUT BUILDERS
# =============================================================================

def _aggregate_ccos(osbal_rows: list, all_fac_osbal_rows: list = None) -> dict:
    fac_codes = {str(r.get("CCOS_REF_CODE", "")).strip() for r in osbal_rows
                 if str(r.get("CCOS_REF_CODE", "")).strip()}
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
            if pd.isna(v):
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
        "CCOS_OR_BAL":   sum_col("CCOS_OR_BAL", osbal_rows),
        "CCOS_BAL_DUE":  sum_col("CCOS_BAL_DUE", bal_due_rows),
    }


def _facode_label(osbal_rows: list, source: str) -> str:
    fac_codes = {str(r.get("CCOS_REF_CODE", "")).strip() for r in osbal_rows
                 if str(r.get("CCOS_REF_CODE", "")).strip()}
    return "facode lebih dari 1" if len(fac_codes) > 1 else ""


def _clean_base_match(scenario: str) -> str:
    s = str(scenario or "").strip()
    for to_remove in [" -> Bordero OC Override", " (Bordero)", " (Beda Currency)", " (Beda Periode)", " (> 1 Fac)", " (> 1 Fac Bordero)"]:
        s = s.replace(to_remove, "")
    s = s.strip()
    if s.startswith("Slip + Polis"):
        s = s.replace("Slip + Polis", "Polis + Slip")
    if " only + " in s:
        s = s.replace(" only + ", " + ")
    return s


def _format_final_scenario(
    base_scenario: str,
    source: str,
    is_bordero: bool,
    is_beda_curr: bool,
    is_beda_per: bool,
    num_fac: int,
) -> str:
    if not source:
        return "Unmatching"

    base = _clean_base_match(base_scenario) if base_scenario else ""
    if not base or base.upper() in ("UNMATCHING", "UNMATCHED"):
        base = "Polis only" if is_bordero else "Unmatching"
    if base == "Unmatching":
        return "Unmatching"

    if is_beda_curr:
        return f"{base} (Beda Currency)"
    elif is_beda_per:
        return f"{base} (Beda Periode)"
    elif is_bordero and num_fac > 1:
        return f"{base} (> 1 Fac Bordero)"
    elif is_bordero and num_fac <= 1:
        return f"{base} (Bordero)"
    elif not is_bordero and num_fac > 1:
        return f"{base} (> 1 Fac)"
    else:
        return base


def _build_output_row(
    suspend_row:     dict,
    source:          str,
    scenario:        str,
    osbal_rows:      list,
    osbal_count:     int  = 0,
    resolved:        bool = True,
    suspend_count:   int  = 1,
    all_fac_rows:    list = None,
) -> dict:
    has_match = bool(source and osbal_rows and osbal_rows[0])
    has_osbal = has_match and (source in ("OSBAL", "SLIPDB", "FACUL") or resolved)

    ccos = (
        _aggregate_ccos(osbal_rows, all_fac_osbal_rows=all_fac_rows)
        if has_osbal
        else {"CCOS_DOC_NO": "", "CCOS_REF_CODE": "", "CCOS_OR_BAL": np.nan, "CCOS_BAL_DUE": np.nan}
    )

    if "Beda Currency" in str(scenario) or "Beda Periode" in str(scenario):
        ccos["CCOS_OR_BAL"] = np.nan
        ccos["CCOS_BAL_DUE"] = np.nan

    final_scenario = "Unmatching" if not source or scenario in ("Unmatching", "UNMATCHED") else scenario

    return {
        "CCOS_DOC_NO":   ccos["CCOS_DOC_NO"] if has_osbal else "",
        "CCOS_REF_CODE": ccos["CCOS_REF_CODE"],

        "RECEIPT NO":        suspend_row.get("RECEIPT NO",        ""),
        "CREDIT NOTES":      suspend_row.get("CREDIT NOTES",      ""),
        "DETAIL RINCIAN NO": suspend_row.get("DETAIL RINCIAN NO", ""),
        "RECEIPT DATE":      suspend_row.get("RECEIPT DATE",      ""),
        "CEDANT NAME":       suspend_row.get("CEDANT NAME",       ""),
        "CEDANT SHRT NAME":  suspend_row.get("CEDANT SHRT NAME",  ""),

        "INSURED_ORI": suspend_row.get("insured_ori") or suspend_row.get("INSURED") or suspend_row.get(SUS_INSURED_ORI_COL, ""),
        "INSURED_1":   suspend_row.get("clean insured 1") or suspend_row.get("INSURED_CLEAN", ""),
        "INSURED_2":   suspend_row.get("clean insured 2", ""),

        "CURR ORI":   suspend_row.get("CURR ORI",   ""),
        "AMOUNT ORI": suspend_row.get("AMOUNT ORI", ""),
        "CURR PAY":   suspend_row.get("CURR PAY",   ""),
        "AMOUNT PAY": suspend_row.get("AMOUNT PAY", ""),

        "CCOS_OR_BAL":            ccos["CCOS_OR_BAL"],
        "CCOS_BAL_DUE":           ccos["CCOS_BAL_DUE"],

        "POLIS":          suspend_row.get("POLIS") or suspend_row.get("polis_ori") or suspend_row.get(SUS_POLIS_ORI_COL, ""),
        "POLICY_CLEAN_1": suspend_row.get("POLICY_CLEAN_1") or suspend_row.get("clean polis 1") or suspend_row.get("POLIS_CLEAN") or suspend_row.get("POLIS_CLN", ""),
        "SERTIF_CLEAN_1": _format_sertif(suspend_row.get("SERTIF_CLEAN_1") or suspend_row.get("clean sertif 1") or suspend_row.get("CERTIFICATE_1") or suspend_row.get("SERTIF_CLN_1") or suspend_row.get("SERTIF_CLN", "")),
        "POLICY_CLEAN_2": suspend_row.get("POLICY_CLEAN_2") or suspend_row.get("clean polis 2", ""),
        "SERTIF_CLEAN_2": _format_sertif(suspend_row.get("SERTIF_CLEAN_2") or suspend_row.get("clean sertif 2") or suspend_row.get("CERTIFICATE_2") or suspend_row.get("SERTIF_CLN_2", "")),
        "POLICY_CLEAN_3": suspend_row.get("POLICY_CLEAN_3") or suspend_row.get("clean polis 3", ""),
        "SERTIF_CLEAN_3": _format_sertif(suspend_row.get("SERTIF_CLEAN_3") or suspend_row.get("clean sertif 3") or suspend_row.get("CERTIFICATE_3") or suspend_row.get("SERTIF_CLN_3", "")),
        "POLICY_CLEAN_4": suspend_row.get("POLICY_CLEAN_4") or suspend_row.get("clean polis 4", ""),
        "POLICY_CLEAN_5": suspend_row.get("POLICY_CLEAN_5") or suspend_row.get("clean polis 5", ""),
        "SLIP_NO":        suspend_row.get("SLIP_NO") or suspend_row.get("slip_ori") or suspend_row.get(SUS_SLIP_ORI_COL, ""),
        "SLIP_NO_CLN":    suspend_row.get("SLIP_NO_CLN") or suspend_row.get("clean slip 1") or suspend_row.get("SLIP_NO_CLEAN", ""),

        "DESC 1": suspend_row.get("DESC 1", ""),
        "DESC 2": suspend_row.get("DESC 2", ""),
        "DESC 3": suspend_row.get("DESC 3", ""),
        "DESC 4": suspend_row.get("DESC 4", ""),

        "STATUS":   suspend_row.get("STATUS",   ""),
        "REC_TYPE": suspend_row.get("REC_TYPE", ""),

        "SKENARIO":           final_scenario,
        "_OSBAL_ROW_COUNT":   osbal_count,
        "_SUSPEND_ROW_COUNT": suspend_count,
    }


def _merge_suspend_rows(rows: list) -> dict:
    if len(rows) == 1:
        return dict(rows[0])
    merged = dict(rows[0])
    for col in _SUSPEND_JOIN_COLS:
        vals, seen = [], set()
        for r in rows:
            v = str(r.get(col, "") or "").strip()
            if v and v not in seen:
                seen.add(v)
                vals.append(v)
        merged[col] = ", ".join(vals)

    for col in _SUSPEND_SUM_COLS:
        tot = 0.0
        for r in rows:
            v = r.get(col, 0)
            if v is None:
                continue
            s = str(v).strip().replace(' ', '')
            if ',' in s and '.' in s:
                s = s.replace('.', '').replace(',', '.') if s.rfind(',') > s.rfind('.') else s.replace(',', '')
            elif ',' in s:
                s = s.replace(',', '.')
            try:
                tot += float(s)
            except (ValueError, TypeError):
                pass
        merged[col] = tot
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
    is_gt1_ocmc          = skenario_col.str.contains("> 1 Fac Bordero",    na=False, regex=False) | (skenario_col.str.contains("Bordero", na=False, regex=False) & ccos_ref.str.contains(",", na=False))
    is_gt1_fac           = (skenario_col.str.contains("> 1 Fac", na=False, regex=False) | ccos_ref.str.contains(",", na=False)) & ~is_gt1_ocmc
    is_new_entry_total   = ~is_unmatched & (amount_ori != 0) & (balance_due == 0)
    is_adj_total         = ~is_unmatched & is_equal
    is_adj_sebagian      = ~is_unmatched & ~is_equal & (amount_neg < balance_due)
    is_new_entry_sebagian= ~is_unmatched & ~is_equal & (amount_neg > balance_due)

    df["FLAG_PROD"] = np.select(
        [is_beda_currency, is_beda_periode, is_gt1_ocmc, is_gt1_fac,
         is_new_entry_total,
         is_adj_total & ~accumulated, is_adj_total & accumulated,
         is_adj_sebagian & ~accumulated, is_adj_sebagian & accumulated,
         is_new_entry_sebagian],
        ["Beda Currency", "Beda Periode",
         "Matching >1 Fac code OC MC", "Matching >1 fac code",
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
# LOAD DATA WITH CACHING
# =============================================================================

def _load_excel(path: str, label: str, optional: bool = False) -> tuple:
    if not os.path.exists(path):
        if optional:
            print(f"  {label:<10} not found ({path}) -> skipped", flush=True)
            return [], []
        raise FileNotFoundError(f"File tidak ditemukan: {path}")

    cache_path = path + ".cache.pkl"
    excel_mtime = os.path.getmtime(path)

    if os.path.exists(cache_path):
        try:
            cache_mtime = os.path.getmtime(cache_path)
            if cache_mtime >= excel_mtime:
                t0 = time.perf_counter()
                import pickle
                with open(cache_path, "rb") as f:
                    rows, cols = pickle.load(f)
                elapsed = time.perf_counter() - t0
                print(f"  {label:<10} loaded from cache in {elapsed:.2f}s ({len(rows):,} rows)", flush=True)
                return rows, cols
        except Exception:
            pass

    t0 = time.perf_counter()
    try:
        df = pd.read_excel(path, dtype=object, engine="calamine")
    except Exception:
        df = pd.read_excel(path, dtype=object)
    elapsed = time.perf_counter() - t0
    print(f"  {label:<10} read in {elapsed:.1f}s ({len(df):,} rows)", flush=True)

    rows = df.to_dict(orient="records")
    cols = list(df.columns)

    try:
        import pickle
        with open(cache_path, "wb") as f:
            pickle.dump((rows, cols), f, protocol=pickle.HIGHEST_PROTOCOL)
    except Exception:
        pass

    return rows, cols


# =============================================================================
# MAIN PIPELINE
# =============================================================================

def run() -> None:
    t_start = time.perf_counter()

    print("\n" + "=" * 70, flush=True)
    print("  PRODUCTION SCRIPT — SUSPEND MATCHING V1 NEW", flush=True)
    print("  Cedant : PT ASURANSI ASTRA BUANA", flush=True)
    print("  Bordero: Multi-file checking (Folder data/BORDERO)", flush=True)
    print("=" * 70, flush=True)

    # [1] Load data
    print("\n[1/6] Loading reference & suspend data ...", flush=True)
    suspend_rows, suspend_cols = _load_excel(SUSPEND_FILE, "SUSPEND")
    osbal_rows,   osbal_cols   = _load_excel(OSBAL_FILE,   "OSBAL")
    facul_rows,   facul_cols   = _load_excel(FACUL_FILE,   "FACUL")
    slipdb_rows,  slipdb_cols  = [], []

    # Pre-parse OSBAL FAC_COM_DATE
    print("  Pre-parsing OSBAL FAC_COM_DATE ...", flush=True)
    for _row in osbal_rows:
        _raw = _row.get(OSBAL_DATE_COL, "")
        if _raw and not (isinstance(_raw, float) and pd.isna(_raw)):
            try:
                _row["_com_date_parsed"] = pd.to_datetime(_raw)
            except Exception:
                _row["_com_date_parsed"] = None
        else:
            _row["_com_date_parsed"] = None

    # Pre-parse Suspend RECEIPT DATE
    print("  Pre-parsing Suspend RECEIPT DATE ...", flush=True)
    for _row in suspend_rows:
        _raw = _row.get(SUSPEND_DATE_COL, "")
        if _raw and not (isinstance(_raw, float) and pd.isna(_raw)):
            try:
                _row["_sus_date_parsed"] = pd.to_datetime(_raw)
            except Exception:
                _row["_sus_date_parsed"] = None
        else:
            _row["_sus_date_parsed"] = None

    # [2] Detect clean columns
    print("\n[2/6] Detecting clean columns ...", flush=True)

    def _get_cols_for(columns: list, table: str, field: str) -> list:
        prefix_map = {
            ("suspend", "polis"):   _SUS_POLIS_PREFIX,
            ("suspend", "slip"):    _SUS_SLIP_PREFIX,
            ("suspend", "insured"): _SUS_INSURED_PREFIX,
            ("suspend", "sertif"):  _SUS_SERTIF_PREFIX,
            ("osbal",   "polis"):   _OSBAL_POLIS_PREFIX,
            ("osbal",   "slip"):    _OSBAL_SLIP_PREFIX,
            ("osbal",   "insured"): _OSBAL_INSURED_PREFIX,
            ("osbal",   "sertif"):  _OSBAL_SERTIF_PREFIX,
            ("facul",   "polis"):   _FACUL_POLIS_PREFIX,
            ("facul",   "slip"):    _FACUL_SLIP_PREFIX,
            ("facul",   "insured"): _FACUL_INSURED_PREFIX,
            ("facul",   "sertif"):  _FACUL_SERTIF_PREFIX,
        }
        prefix = prefix_map.get((table, field), "")
        return _get_clean_cols(columns, prefix) if prefix else []

    polis_sus   = _get_cols_for(suspend_cols, "suspend", "polis")
    slip_sus    = _get_cols_for(suspend_cols, "suspend", "slip")
    insured_sus = _get_cols_for(suspend_cols, "suspend", "insured")
    sertif_sus  = _get_cols_for(suspend_cols, "suspend", "sertif")

    polis_osbal   = _get_cols_for(osbal_cols, "osbal", "polis")
    slip_osbal    = _get_cols_for(osbal_cols, "osbal", "slip")
    insured_osbal = _get_cols_for(osbal_cols, "osbal", "insured")
    sertif_osbal  = _get_cols_for(osbal_cols, "osbal", "sertif")

    polis_facul   = _get_cols_for(facul_cols, "facul", "polis")
    slip_facul    = _get_cols_for(facul_cols, "facul", "slip")
    insured_facul = _get_cols_for(facul_cols, "facul", "insured")
    sertif_facul  = _get_cols_for(facul_cols, "facul", "sertif")

    # Lineslip parsing
    print("  Pre-parsing LINESLIP flags ...", flush=True)
    lineslip_count = 0
    for _row in suspend_rows:
        _row["_is_lineslip"] = _is_lineslip_row(_row, list(insured_sus) + [SUS_INSURED_ORI_COL])
        if _row["_is_lineslip"]:
            lineslip_count += 1
            _parsed = _row.get("_sus_date_parsed")
            if _parsed is not None:
                try:
                    _row["_sus_date_lineslip"] = _parsed - pd.DateOffset(months=1)
                except Exception:
                    _row["_sus_date_lineslip"] = None
            else:
                _row["_sus_date_lineslip"] = None
        else:
            _row["_sus_date_lineslip"] = None
    print(f"  -> {lineslip_count:,} baris LINESLIP terdeteksi", flush=True)

    def _prepend(col: str, cols_list: list, header: list) -> list:
        return ([col] + cols_list) if col in header and col not in cols_list else cols_list

    polis_osbal = _prepend("CLSDT_POLICY_NO", polis_osbal, osbal_cols)
    slip_osbal  = _prepend("CLSDT_SLIP_NO",   slip_osbal,  osbal_cols)
    if "CLSDT_SERTF_NO" in osbal_cols and "CLSDT_SERTF_NO" not in sertif_osbal:
        sertif_osbal.insert(0, "CLSDT_SERTF_NO")

    polis_facul = _prepend("CLSDT_POLICY_NO", polis_facul, facul_cols)
    slip_facul  = _prepend("CLSDT_SLIP_NO",   slip_facul,  facul_cols)

    # [3] Build lookup index
    print("\n[3/6] Building lookup indexes ...", flush=True)
    lookup_osbal = _build_lookup(osbal_rows, polis_osbal, slip_osbal, insured_osbal, OSBAL_FACODE_COL, label="OSBAL")
    lookup_facul = _build_lookup(facul_rows, polis_facul, slip_facul, insured_facul, FACUL_FACODE_COL, label="FACUL")
    lookup_slipdb = ({}, {}, {}, {}, {}, {}, ({}, {}), ({}, {}), ({}, {}), {})

    facode_osbal_idx = lookup_osbal[9]

    # [4] Load multi-file Bordero
    bordero_files = _find_bordero_files()
    bordero_file_items = []
    if bordero_files:
        print(f"\n[4/6] Loading {len(bordero_files)} Bordero file(s) ...", flush=True)
        for i, bf in enumerate(bordero_files, 1):
            fname = os.path.basename(bf)
            b_rows, _ = _load_excel(bf, f"BORDERO #{i}", optional=True)
            if b_rows:
                b_idx = _build_bordero_index(b_rows)
                b_polis_idx = {}
                for fac, rows in b_idx.items():
                    for r in rows:
                        p = r.get("polis")
                        if p and len(p) >= 5:
                            b_polis_idx.setdefault(p, []).append((fac, r))
                print(f"      [{i}/{len(bordero_files)}] {fname} -> {len(b_rows):,} baris, {len(b_idx):,} FAC codes, {len(b_polis_idx):,} unique polis", flush=True)
                bordero_file_items.append({
                    "filename": fname,
                    "filepath": bf,
                    "idx": b_idx,
                    "polis_idx": b_polis_idx,
                    "rows": b_rows,
                })
            else:
                print(f"      [{i}/{len(bordero_files)}] {fname} -> Kosong / gagal dimuat", flush=True)
    else:
        print("\n[4/6] Bordero: tidak ada file di folder data/BORDERO", flush=True)

    # [5] Matching per baris
    total = len(suspend_rows)
    print(f"\n[5/6] Matching {total:,} suspend rows ...", flush=True)
    t4 = time.perf_counter()
    raw_results = []

    for n, sus in enumerate(suspend_rows, 1):
        if n % 1000 == 0:
            print(f"    Memproses baris {n:,} / {total:,} ...", flush=True)

        if str(sus.get("STATUS", "")).strip().upper() == "MATCHING":
            continue

        eff_polis, eff_slip, like_polis, like_slip = _get_effective_sus_cols(sus, polis_sus, slip_sus)
        _is_ls = sus.get("_is_lineslip", False)
        effective_date = sus.get("_sus_date_lineslip") if _is_ls else sus.get("_sus_date_parsed")

        matched_fac_codes = set()
        matched_osbal_idx = set()
        source = None
        scenario = ""
        resolved = False

        # FASE 1: PENCARIAN AWAL BERDASARKAN POLIS
        polis_vals = _collect_clean_values(sus, eff_polis)
        found_by_polis = False
        if polis_vals:
            osb_hits = _exact_match(polis_vals, lookup_osbal[1])
            if osb_hits:
                matched_osbal_idx.update(osb_hits)
                found_by_polis = True

            if facul_rows:
                fac_hits = _exact_match(polis_vals, lookup_facul[1])
                if fac_hits:
                    res_idx, _ = _resolve_facode([facul_rows[idx] for idx in fac_hits], FACUL_FACODE_COL, facode_osbal_idx, osbal_rows)
                    if res_idx:
                        matched_osbal_idx.update(res_idx)
                        found_by_polis = True

        if found_by_polis:
            source = "OSBAL"
            scenario = "Polis only"

        # FASE 2: PENCARIAN AWAL BERDASARKAN SLIP
        if not found_by_polis:
            slip_vals = _collect_clean_values(sus, slip_sus)
            if slip_vals:
                osb_hits = _exact_match(slip_vals, lookup_osbal[0])
                if osb_hits:
                    matched_osbal_idx.update(osb_hits)
                    source = "OSBAL"
                    scenario = "Slip only"

                if facul_rows and not matched_osbal_idx:
                    fac_hits = _exact_match(slip_vals, lookup_facul[0])
                    if fac_hits:
                        res_idx, _ = _resolve_facode([facul_rows[idx] for idx in fac_hits], FACUL_FACODE_COL, facode_osbal_idx, osbal_rows)
                        if res_idx:
                            matched_osbal_idx.update(res_idx)
                            source = "FACUL"
                            scenario = "Slip only"
                            resolved = True

        # FASE 3: PENCARIAN AWAL BERDASARKAN INSURED
        if not matched_osbal_idx:
            ins_vals = _collect_clean_values(sus, insured_sus)
            if ins_vals:
                osb_hits = _exact_match(ins_vals, lookup_osbal[2])
                if osb_hits:
                    matched_osbal_idx.update(osb_hits)
                    source = "OSBAL"
                    scenario = "Insured only"
                elif facul_rows and not matched_osbal_idx:
                    fac_hits = _exact_match(ins_vals, lookup_facul[2])
                    if fac_hits:
                        res_idx, _ = _resolve_facode([facul_rows[idx] for idx in fac_hits], FACUL_FACODE_COL, facode_osbal_idx, osbal_rows)
                        if res_idx:
                            matched_osbal_idx.update(res_idx)
                            source = "FACUL"
                            scenario = "Insured only"
                            resolved = True

        # FASE 4: NARROWING WAJIB (Currency & Periode)
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

            if not is_beda_curr:
                sus_date = effective_date
                if pd.notna(sus_date):
                    filtered_per = []
                    for i in matched_list:
                        com_date = osbal_rows[i].get("_com_date_parsed")
                        if pd.isna(com_date):
                            filtered_per.append(i)
                        else:
                            if _is_ls:
                                if com_date.year == sus_date.year and com_date.month == sus_date.month:
                                    filtered_per.append(i)
                            else:
                                if sus_date >= com_date:
                                    filtered_per.append(i)
                    if not filtered_per:
                        is_beda_per = True
                    elif len(filtered_per) < len(matched_list):
                        matched_list = filtered_per

            if is_beda_curr or is_beda_per:
                rm = _rematch_strict(
                    sus, osbal_rows, lookup_osbal,
                    eff_polis, slip_sus, insured_sus,
                    polis_osbal, slip_osbal, insured_osbal,
                    facode_osbal_idx,
                    slipdb_rows, lookup_slipdb,
                    facul_rows, lookup_facul,
                    like_polis,
                    effective_sus_date=effective_date, is_lineslip=_is_ls
                )
                if rm:
                    src, scen, hits, _ = rm
                    source = src
                    scenario = scen
                    matched_list = hits
                    is_beda_curr = False
                    is_beda_per = False

            matched_osbal_idx = set(matched_list)

        # FASE 5: NARROWING UTAMA (Non-Destructive)
        if len(matched_osbal_idx) > 1:
            matched_list = list(matched_osbal_idx)

            # Step 1: Slip Exact Matching
            slip_vals = _collect_clean_values(sus, slip_sus)
            if slip_vals:
                confirmed = [idx for idx in matched_list if _ref_has_value(osbal_rows[idx], slip_osbal, "slip_ori", slip_vals)]
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
            for pv in [_normalize(sus.get(SUS_POLIS_ORI_COL, ""))] + [_normalize(sus.get(c, "")) for c in (eff_polis or [])]:
                if pv:
                    c = _extract_cert_from_polis(pv)
                    if c:
                        sus_cert_vals.add(_clean_cert_str(c))
            sus_cert_vals.discard("")

            if sus_cert_vals:
                confirmed = [idx for idx in matched_list if _ref_has_value(osbal_rows[idx], sertif_osbal, "", list(sus_cert_vals))]
                if not confirmed:
                    confirmed = [
                        idx for idx in matched_list
                        if any(
                            _cert_in_range(cv, osbal_rows[idx].get(col, ""))
                            for cv in sus_cert_vals
                            for col in sertif_osbal
                        )
                    ]
                if confirmed:
                    matched_list = confirmed
                    if "Cert" not in scenario:
                        scenario = scenario + " + Cert"

            # Step 3: Insured Matching
            ins_vals = _collect_all_values(sus, insured_sus, SUS_INSURED_ORI_COL)
            if ins_vals:
                confirmed = [idx for idx in matched_list if _ref_has_value(osbal_rows[idx], insured_osbal, "FAC_INSURED", ins_vals)]
                if confirmed:
                    matched_list = confirmed
                    if "Insured" not in scenario:
                        scenario = scenario + " + Insured"

            matched_osbal_idx = set(matched_list)

        # FASE 6: BORDERO (Open Cover) OVERRIDE — DICEK SATU PER SATU PER FILE (FIRE, MC, & MH)
        is_bordero = False
        if bordero_file_items and (found_by_polis or matched_osbal_idx or polis_vals):
            for b_item in bordero_file_items:
                b_idx = b_item["idx"]
                b_polis_idx = b_item["polis_idx"]
                b_name = b_item["filename"]

                bordero_facs, ok = _find_fac_codes_in_bordero(
                    sus, b_idx, eff_polis, slip_sus,
                    bordero_polis_index=b_polis_idx,
                    sertif_sus_cols=sertif_sus
                )
                if ok and bordero_facs:
                    res_idx = []
                    for fac in bordero_facs:
                        if fac in facode_osbal_idx:
                            res_idx.extend(facode_osbal_idx[fac])

                    if res_idx:
                        if set(res_idx) != matched_osbal_idx:
                            matched_osbal_idx = set(res_idx)
                        source = "OSBAL"
                        is_bordero = True
                        resolved = True
                        if not scenario or scenario == "Unmatching":
                            scenario = "Polis only"
                        # Berhasil cocok pada file bordero ini, hentikan pengecekan file berikutnya
                        break

        if is_beda_curr or is_beda_per:
            osbal_ref = [osbal_rows[i] for i in matched_list]
            osbal_count = 0
            resolved = False
        elif not matched_osbal_idx:
            source = None
            scenario = "Unmatching"
            osbal_ref = [{}]
            osbal_count = 0
            resolved = False
        else:
            osbal_ref = [osbal_rows[i] for i in matched_osbal_idx]
            osbal_count = len(osbal_ref)

        if scenario and " only + " in scenario:
            scenario = scenario.replace(" only + ", " + ")

        fac_codes = {_normalize(r.get(OSBAL_FACODE_COL, "")) for r in osbal_ref if r}
        fac_codes.discard("")

        raw_results.append({
            "suspend": sus, "source": source,
            "base_scenario": scenario,
            "is_bordero": is_bordero,
            "is_beda_curr": is_beda_curr,
            "is_beda_per": is_beda_per,
            "osbal_idx": list(matched_osbal_idx) if not (is_beda_curr or is_beda_per) else list(matched_list),
            "osbal_count": osbal_count, "resolved": resolved, "fac_codes": fac_codes,
        })

    print(f"  Matching selesai dalam {time.perf_counter()-t4:.1f}s", flush=True)

    # Fac code claim & osbal accumulation (V1: baris asli dipertahankan)
    print("\n[6/6] Fac code claim & osbal accumulation ...", flush=True)

    claimed: set = set()
    for r in raw_results:
        if len(r["fac_codes"]) == 1 and r["source"] is not None and not r["is_beda_curr"] and not r["is_beda_per"]:
            claimed |= r["fac_codes"]

    for r in raw_results:
        if len(r["fac_codes"]) > 1:
            remaining = r["fac_codes"] - claimed
            if len(remaining) == 1 and r["source"] is not None and not r["is_beda_curr"] and not r["is_beda_per"]:
                r["fac_codes"] = remaining
                claimed |= remaining
                fac = next(iter(remaining))
                r["osbal_idx"] = [i for i in r["osbal_idx"]
                                   if _normalize(osbal_rows[i].get(OSBAL_FACODE_COL, "")) == fac]
                r["osbal_count"] = len(r["osbal_idx"])

    # Group per fac untuk shared osbal
    groups: dict = {}
    final:  list = []

    for r in raw_results:
        r.setdefault("suspend_count", 1)
        if (len(r["fac_codes"]) == 1 and r["source"] is not None
                and not r["is_beda_curr"] and not r["is_beda_per"]):
            groups.setdefault(next(iter(r["fac_codes"])), []).append(r)
        else:
            final.append(r)

    for fac, grp in groups.items():
        all_fac_idx = facode_osbal_idx.get(fac, [])
        all_fac_rows = [osbal_rows[i] for i in all_fac_idx] if all_fac_idx else None

        if len(grp) == 1:
            r = grp[0]
            if all_fac_rows:
                r["all_fac_rows"] = all_fac_rows
            final.append(r)
        else:
            ref_osbal = [osbal_rows[i] for i in grp[0]["osbal_idx"]]
            all_osbal_idx = set()
            for r in grp:
                all_osbal_idx.update(r["osbal_idx"])
            all_osbal_ref = [osbal_rows[i] for i in all_osbal_idx] if all_osbal_idx else ref_osbal

            sus_rows = [r["suspend"] for r in grp]
            merged_sus = _merge_suspend_rows(sus_rows)
            shared_amount = merged_sus.get("AMOUNT ORI", 0)

            osbal_cnt = grp[0]["osbal_count"]
            if all_fac_idx and len(all_fac_idx) > 0:
                ref_osbal = all_fac_rows
                osbal_cnt = len(all_fac_idx)

            for member in grp:
                member["osbal_rows"] = ref_osbal
                member["osbal_count"] = osbal_cnt
                member["all_fac_rows"] = all_fac_rows
                member["_merged_amount_ori"] = shared_amount
                member["suspend_count"] = len(grp)
                final.append(member)

    final.sort(key=lambda r: suspend_rows.index(r["suspend"]))

    print(f"  {len(raw_results):,} suspend rows -> {len(final):,} output rows (row asli dipertahankan)", flush=True)

    # Build output DataFrame
    output_rows = []
    for r in final:
        ref_osbal = r.get("osbal_rows") or ([osbal_rows[i] for i in r["osbal_idx"]] if r["source"] else [{}])
        all_fac_rows = r.get("all_fac_rows")
        osbal_cnt = r["osbal_count"]
        if r["source"] and len(r["fac_codes"]) == 1:
            fac = next(iter(r["fac_codes"]))
            all_fac_idx = facode_osbal_idx.get(fac, [])
            if len(all_fac_idx) > 0:
                all_fac_rows = [osbal_rows[i] for i in all_fac_idx]
                osbal_cnt = len(all_fac_idx)

        final_scenario = _format_final_scenario(
            base_scenario=r.get("base_scenario", ""),
            source=r["source"],
            is_bordero=r.get("is_bordero", False),
            is_beda_curr=r.get("is_beda_curr", False),
            is_beda_per=r.get("is_beda_per", False),
            num_fac=len(r["fac_codes"]),
        )

        out_row = _build_output_row(
            r["suspend"],
            source          = r["source"],
            scenario        = final_scenario,
            osbal_rows      = ref_osbal,
            osbal_count     = osbal_cnt,
            resolved        = r["resolved"],
            suspend_count   = r.get("suspend_count", 1),
            all_fac_rows    = all_fac_rows,
        )
        if "_merged_amount_ori" in r:
            out_row["_MERGED_AMOUNT_ORI"] = r["_merged_amount_ori"]
        output_rows.append(out_row)

    df = pd.DataFrame(output_rows)

    # Format numeric
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

    df.loc[df["SKENARIO"] == "UNMATCHED",   "SKENARIO"]  = "Unmatching"
    df.loc[df["SKENARIO"] == "Unmatching",  "FLAG_PROD"] = "Unmatching"

    for col in FINAL_COLUMNS:
        if col not in df.columns:
            df[col] = ""
    df = df[FINAL_COLUMNS]

    # Export to Excel
    print(f"\nSaving to: {OUTPUT_FILE} ...", flush=True)
    for c in df.select_dtypes(include=['object']).columns:
        df[c] = (df[c]
                 .fillna('')
                 .astype(str)
                 .str.replace(r'[\x00-\x08\x0B\x0C\x0E-\x1F]', '', regex=True))

    save_path = OUTPUT_FILE
    try:
        df.to_excel(save_path, index=False)
    except PermissionError:
        save_path = OUTPUT_FILE.replace(".xlsx", "_new.xlsx")
        print(f"  [WARN] '{OUTPUT_FILE}' terkunci/sedang dibuka. Menyimpan ke '{save_path}' ...", flush=True)
        df.to_excel(save_path, index=False)

    try:
        from excel_styler import apply_purple_column_style
        apply_purple_column_style(save_path, "matching")
    except Exception as e:
        print(f"  [WARN] Styling failed: {e}")

    # Export to PostgreSQL
    env_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), ".env")
    load_dotenv(env_path)
    load_dotenv()
    print("  Exporting to PostgreSQL ...")
    db_user = os.environ.get("DB_USER", "viy4user")
    db_pass = os.environ.get("DB_PASS", "fild42op61nx")
    db_host = os.environ.get("DB_HOST", "10.10.125.91")
    db_port = os.environ.get("DB_PORT", "5432")
    db_name = os.environ.get("DB_NAME", "fsi_db")

    conn_str = f"postgresql+psycopg2://{db_user}:{db_pass}@{db_host}:{db_port}/{db_name}"
    try:
        df_db = df.copy()

        # RECEIPT DATE -> date or None
        df_db["RECEIPT DATE"] = pd.to_datetime(df_db["RECEIPT DATE"], errors="coerce").dt.date
        df_db["RECEIPT DATE"] = df_db["RECEIPT DATE"].where(df_db["RECEIPT DATE"].notnull(), None)

        # Numerics -> float or None
        numeric_cols = ["AMOUNT ORI", "AMOUNT_ORI_MIN1", "AMOUNT PAY", "CCOS_OR_BAL", "CCOS_BAL_DUE", "DIFERENCE"]
        for c in numeric_cols:
            df_db[c] = pd.to_numeric(df_db[c], errors="coerce")

        # Text columns: replace "nan", "None", np.nan with None
        text_cols = [c for c in FINAL_COLUMNS if c not in numeric_cols and c != "RECEIPT DATE"]
        for c in text_cols:
            df_db[c] = df_db[c].astype(str).str.strip()
            df_db[c] = df_db[c].replace({"nan": None, "None": None, "": None})

        engine = create_engine(conn_str)
        with engine.begin() as con:
            con.execute(text('DELETE FROM "SUSPENSE_DATA_SUSPENSE_V1" WHERE "CEDANT NAME" = \'PT ASURANSI ASTRA BUANA\';'))
            con.execute(text('DELETE FROM "SUSPENSE_CLEAN_2026" WHERE "CEDANT NAME" = \'PT ASURANSI ASTRA BUANA\';'))
        df_db.to_sql("SUSPENSE_DATA_SUSPENSE_V1", con=engine, if_exists="append", index=False)
        df_db.to_sql("SUSPENSE_CLEAN_2026", con=engine, if_exists="append", index=False)
        print(f"      -> Successfully exported {len(df_db):,} rows to table 'SUSPENSE_DATA_SUSPENSE_V1' & 'SUSPENSE_CLEAN_2026'")
    except Exception as e:
        print(f"  [WARN] PostgreSQL export failed: {e}")

    elapsed = time.perf_counter() - t_start
    print(f"\n{'=' * 70}")
    print(f"  Done in {elapsed:.1f}s ({elapsed/60:.1f} min)")
    print(f"  Rows  : {len(df):,}")
    print(f"  Cols  : {len(df.columns)}")
    print(f"  File  : {save_path}")

    print(f"\n  FLAG_PROD summary:")
    for flag, count in df["FLAG_PROD"].value_counts().items():
        print(f"    {flag:<45}: {count:,}")

    print(f"\n  SKENARIO summary:")
    for sce, count in df["SKENARIO"].value_counts().items():
        print(f"    {sce:<50}: {count:,}")

    print("=" * 70)


if __name__ == "__main__":
    run()
