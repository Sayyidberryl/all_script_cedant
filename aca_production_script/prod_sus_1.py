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

SCRIPT_DIR   = os.path.dirname(os.path.abspath(__file__))
DATA_DIR     = os.path.join(SCRIPT_DIR, "data")

def _find_slipdb_file() -> str:
    kandidat = [
        os.path.join(DATA_DIR, "slipdb_clean_aca.xlsx"),
        os.path.join(DATA_DIR, "ri slip.xlsx"),
        os.path.join(DATA_DIR, "ri_slip.xlsx"),
        os.path.join(DATA_DIR, "slipdb.xlsx"),
    ]
    for p in kandidat:
        if os.path.exists(p):
            return p
    return kandidat[0]

def _find_bordero_file() -> str:
    kandidat = [
        os.path.join(DATA_DIR, "ACA_Open_Cover_Marine_Cargo.xlsx"),
        os.path.join(DATA_DIR, "ACA_Database_Open_Cover_Marine_Cargo.xlsx"),
        os.path.join(DATA_DIR, "ACA_Database_Open_Cover_Marine_Hull.xlsx"),
    ]
    for p in kandidat:
        if os.path.exists(p):
            return p
    if os.path.exists(DATA_DIR):
        for fname in os.listdir(DATA_DIR):
            if "open_cover" in fname.lower() and fname.endswith(".xlsx"):
                return os.path.join(DATA_DIR, fname)
    return kandidat[0]


SUSPEND_FILE = os.path.join(DATA_DIR, "suspend_clean_aca.xlsx")
OSBAL_FILE   = os.path.join(DATA_DIR, "osbal_clean_aca.xlsx")
FACUL_FILE   = os.path.join(DATA_DIR, "facul_clean_aca.xlsx")
SLIPDB_FILE  = _find_slipdb_file()
BORDERO_FILE = _find_bordero_file()
OUTPUT_FILE  = os.path.join(DATA_DIR, "final_output_v1.xlsx")

# Kolom FAC CODE di masing-masing tabel
OSBAL_FACODE_COL  = "CCOS_REF_CODE"
FACUL_FACODE_COL  = "FAC_CODE"
SLIPDB_FACODE_COL = "FAC_CODE"

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

# Kolom polis/slip suspend — prioritas CLSDT, fallback FAC
CLSDT_POLIS_COL = "CLSDT_POLICY_NO"
CLSDT_SLIP_COL  = "CLSDT_SLIP_NO"
FAC_POLIS_COL   = "FAC_POLICY_NO"
FAC_SLIP_COL    = "FAC_SLIP"

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
SUSPEND_SERTIF_PREFIX = "clean sertif"
OSBAL_SERTIF_PREFIX   = "clean sertif"
FACUL_SERTIF_PREFIX   = "clean sertif"

_MIN_TOKEN_LEN = 3
_TOKEN_RE      = re.compile(r"[^A-Z0-9]+")
_DELIMITER_RE  = re.compile(r"[,;|+\s]+")

# Kolom suspend yang di-join saat akumulasi (v1: digunakan untuk shared amount)
_SUSPEND_JOIN_COLS = [
    "RECEIPT NO", "CREDIT NOTES", "DETAIL RINCIAN NO", "RECEIPT DATE",
    "CEDANT NAME", "CEDANT SHRT NAME",
    "insured_ori", "clean insured 1", "clean insured 2",
    "CURR ORI", "CURR PAY", "AMOUNT PAY",
    "polis_ori", "clean polis 1",
    "slip_ori", "clean slip 1",
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
# STRING UTILITIES
# =============================================================================

@functools.lru_cache(maxsize=131072)
def _norm_cached(s: str) -> str:
    # Normalisasi string — cached karena dipanggil jutaan kali
    text = s.strip().upper().replace("S/D", "SD")
    return re.sub(r"\s+", " ", text) if "  " in text else text


def _normalize(value) -> str:
    # Uppercase + strip + collapse spasi, return '' untuk null
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    return _norm_cached(str(value))


def _expand_sertif_range(val: str) -> list:
    # Expand range sertif: '100-110' atau '100 SD 110' -> ['000100', '000101', ...]
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
    """Mengecek apakah nomor slip suspend berada di dalam range slip OSBAL (antara clean slip 1 dan clean slip 2)."""
    if not sus_slip:
        return False
    s_clean = re.sub(r'\D', '', str(sus_slip))
    if not s_clean:
        return False

    sl1 = re.sub(r'\D', '', str(osbal_row.get("clean slip 1", "")))
    sl2 = re.sub(r'\D', '', str(osbal_row.get("clean slip 2", "")))
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
    # Pecah string jadi frozenset token alfanumerik (min 3 karakter)
    if not text:
        return frozenset()
    return frozenset(t for t in _TOKEN_RE.split(text) if len(t) >= _MIN_TOKEN_LEN)


import functools

@functools.lru_cache(maxsize=100000)
def _split_composite(text: str) -> frozenset:
    # Pecah string berdelimiter (koma/plus/spasi/titik koma/pipe) jadi token
    if not text:
        return frozenset()
    return frozenset(tok.strip() for tok in _DELIMITER_RE.split(text) if tok.strip())


def _get_clean_cols(columns: list, prefix: str) -> list:
    # Ambil nama kolom yang berawalan prefix (case-insensitive) atau variasinya
    if not columns:
        return []
    alt_map = {
        "clean polis": ["clean polis", "policy_clean", "polis_cln"],
        "clean slip": ["clean slip", "slip_clean", "slip_no_cln"],
        "clean insured": ["clean insured", "insured_clean", "insured_"],
        "clean sertif": ["clean sertif", "sertif_clean", "sertif_cln"],
    }
    prefixes = alt_map.get(prefix.lower(), [prefix.lower()])
    res = []
    for c in columns:
        clow = str(c).lower()
        if any(clow.startswith(p) for p in prefixes):
            res.append(c)
    return res


def _map_scenario(label: str) -> str:
    # Petakan label internal ke nama skenario resmi
    if not label or label in ("Unmatching", "UNMATCHED"):
        return "Unmatching"
    if "INSURED" in label:
        return "Insured only"
    if "SLIP" in label:
        return "Slip only"
    if "POLIS" in label:
        return "Polis only"
    return label


def _format_sertif(val) -> str:
    # Format nilai sertif jadi string 6-digit (contoh: 7 -> '000007')
    if pd.isna(val) or str(val).strip() == "":
        return ""
    if isinstance(val, (int, float)):
        return str(int(val)).zfill(6)
    s = str(val).strip()
    if s.endswith('.0'):
        s = s[:-2]
    return s.zfill(6) if s.isdigit() else s


def _extract_cert_from_polis(polis_val: str) -> str:
    # Ekstrak 6-digit sertif dari polis_ori (format: base-CERT6DIGIT)
    if not polis_val:
        return ""
    m = _CERT_RE.search(polis_val)
    return m.group(1) if m else ""


# =============================================================================
# ROW DICT HELPERS
# =============================================================================

def _is_excluded_ref_row(row: dict, polis_cols: list, slip_cols: list) -> bool:
    # True jika baris referensi adalah baris administratif (HUTANG PIUTANG / DATA SUSPENSE)
    for col in list(polis_cols) + ["polis_ori"] + list(slip_cols) + ["slip_ori"]:
        val = _normalize(row.get(col, ""))
        if val and any(m in val for m in _EXCLUDED_REF_MARKERS):
            return True
    return False


def _is_lineslip_row(sus_row: dict, insured_cols: list) -> bool:
    # True jika baris suspend mengandung marker LINESLIP di kolom insured
    for col in list(insured_cols) + ["insured_ori"]:
        val = _normalize(sus_row.get(col, ""))
        if val and any(m in val for m in _LINESLIP_MARKERS):
            return True
    return False


def _collect_clean_values(row: dict, cols: list) -> list:
    # Kumpulkan nilai bersih unik (non-empty) dari kolom yang diberikan
    seen, result = set(), []
    for col in cols:
        val = _normalize(row.get(col, ""))
        if val and val not in seen:
            seen.add(val)
            result.append(val)
    return result


def _collect_all_values(row: dict, clean_cols: list, ori_col: str) -> list:
    # Gabung nilai clean menjadi list unik
    return list(set(_collect_clean_values(row, clean_cols)))


def _get_effective_sus_cols(sus_row: dict, base_polis_cols: list, base_slip_cols: list) -> tuple:
    # Tentukan kolom efektif polis/slip per baris: CLSDT jika ada, else clean bawaan
    # LIKE: tambahkan FAC sebagai opsi terakhir ketika CLSDT kosong
    clsdt_polis = _normalize(sus_row.get(CLSDT_POLIS_COL, ""))
    fac_polis   = _normalize(sus_row.get(FAC_POLIS_COL,   ""))
    if clsdt_polis:
        eff_polis  = [CLSDT_POLIS_COL]
        like_polis = [CLSDT_POLIS_COL]
    else:
        eff_polis  = base_polis_cols
        like_polis = list(base_polis_cols) + ([FAC_POLIS_COL] if fac_polis else [])

    clsdt_slip = _normalize(sus_row.get(CLSDT_SLIP_COL, ""))
    fac_slip   = _normalize(sus_row.get(FAC_SLIP_COL,   ""))
    if clsdt_slip:
        eff_slip  = [CLSDT_SLIP_COL]
        like_slip = [CLSDT_SLIP_COL]
    else:
        eff_slip  = base_slip_cols
        like_slip = list(base_slip_cols) + ([FAC_SLIP_COL] if fac_slip else [])

    return eff_polis, eff_slip, like_polis, like_slip


def _ref_has_value(ref_row: dict, clean_cols: list, ori_col: str, query_values: list) -> bool:
    if not query_values:
        return False
    ref_vals = set(_collect_clean_values(ref_row, clean_cols))
    if set(query_values) & ref_vals:
        return True
    
    # Early return optimization for token check
    for rv in ref_vals:
        if any(v in _split_composite(rv) for v in query_values if v and len(v) >= 2):
            return True
            
    # Substring check min 10 char
    return any(
        v in rv
        for v in query_values if v and len(v) >= 10
        for rv in ref_vals if rv
    )


# =============================================================================
# INDEX BUILDERS
# =============================================================================

def _build_exact_index(rows: list, cols: list, excluded: set = None) -> dict:
    # Bangun inverted index: value -> [indeks baris] untuk exact match
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
    # Bangun inverted index: FAC_CODE -> [indeks baris]
    index: dict = {}
    for i, row in enumerate(rows):
        if excluded and i in excluded:
            continue
        val = _normalize(row.get(col, ""))
        if val:
            index.setdefault(val, []).append(i)
    return index


def _build_token_index(rows: list, clean_cols: list, ori_col: str, excluded: set = None) -> tuple:
    # Bangun token inverted index + cache token per baris untuk fuzzy matching
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


def _build_sertif_index(rows: list, sertif_cols: list, excluded: set = None) -> dict:
    # Bangun inverted index: sertif_value -> [indeks baris], expand range otomatis
    index: dict = {}
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
    # Bangun semua struktur lookup untuk satu tabel referensi (10-tuple)
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
        {}, {}, {},  # ori index (tidak dipakai)
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
    # Kembalikan indeks baris yang cocok exact dengan salah satu nilai query
    return {i for v in query_values if v and v in index for i in index[v]}


def _like_match(query_values: list, token_index: dict, rows: list,
                ref_cols: list, ori_col: str = None) -> list:
    # Kembalikan indeks baris di mana nilai query muncul sebagai substring di ref
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


def _match_polis_slip(
    suspend_row: dict,
    lookup_slip_cln: dict, lookup_polis_cln: dict,
    lookup_slip_ori: dict, lookup_polis_ori: dict,
    polis_sus_cols: list, slip_sus_cols: list,
    rows: list = None, slip_ref_cols: list = None, polis_ref_cols: list = None,
    token_slip: tuple = None, token_polis: tuple = None,
    polis_sus_like_cols: list = None, slip_sus_like_cols: list = None,
) -> tuple:
    # Matching polis/slip: Stage 1 exact (CLSDT/clean), Stage 2 LIKE (+ FAC fallback)
    # Urutan: POLIS dulu (lebih spesifik), baru SLIP
    for values, index, label in [
        (_collect_clean_values(suspend_row, polis_sus_cols), lookup_polis_cln, "POLIS_CLEAN"),
        (_collect_clean_values(suspend_row, slip_sus_cols),  lookup_slip_cln,  "SLIP_CLEAN"),
    ]:
        hit = _exact_match(values, index)
        if hit:
            return list(hit), label

    _slip_like  = slip_sus_like_cols  if slip_sus_like_cols  is not None else slip_sus_cols
    _polis_like = polis_sus_like_cols if polis_sus_like_cols is not None else polis_sus_cols

    if rows is not None and token_slip and token_polis:
        for values, idx_tok, ref_cols, label in [
            (_collect_clean_values(suspend_row, _polis_like), token_polis[0], polis_ref_cols or [], "POLIS_LIKE"),
            (_collect_clean_values(suspend_row, _slip_like),  token_slip[0],  slip_ref_cols  or [], "SLIP_LIKE"),
        ]:
            if values and idx_tok:
                hit = _like_match(values, idx_tok, rows, ref_cols)
                if hit:
                    return hit, label

    return [], "Unmatching"


def _match_insured(
    suspend_row: dict,
    lookup_insured_cln: dict, lookup_insured_ori: dict,
    insured_sus_cols: list,
    rows: list = None, insured_ref_cols: list = None, token_insured: tuple = None,
) -> tuple:
    # Matching insured: Stage 1 exact, Stage 2 LIKE
    hit = _exact_match(_collect_clean_values(suspend_row, insured_sus_cols), lookup_insured_cln)
    if hit:
        return list(hit), "INSURED_CLEAN"

    if rows is not None and token_insured:
        values = _collect_clean_values(suspend_row, insured_sus_cols)
        idx_tok = token_insured[0] if token_insured else None
        if values and idx_tok:
            hit = _like_match(values, idx_tok, rows, insured_ref_cols or [])
            if hit:
                return hit, "INSURED_LIKE"

    return [], "Unmatching"


def _match_polis_sertif(
    suspend_row: dict, osbal_rows: list, lookup_osbal: tuple,
    polis_sus_cols: list, sertif_sus_cols: list,
    polis_ref_cols: list, sertif_idx: dict,
) -> tuple:
    # Pass A: matching Polis + Sertifikat (lebih spesifik dari Polis+Slip)
    polis_vals = _collect_clean_values(suspend_row, polis_sus_cols)
    polis_hit = _exact_match(polis_vals, lookup_osbal[1])
    if not polis_hit:
        return [], "Unmatching"

    # Expand range sertif di suspend
    sertif_vals = []
    for v in _collect_clean_values(suspend_row, sertif_sus_cols):
        sertif_vals.extend(_expand_sertif_range(v))
    sertif_vals = list(set(sertif_vals))

    if not sertif_vals:
        return list(polis_hit), "Polis only"

    # Filter kandidat polis yang sertifnya cocok
    sertif_confirmed = [
        i for i in polis_hit
        if any(sv in sertif_idx and i in sertif_idx[sv] for sv in sertif_vals)
    ]
    return (sertif_confirmed, "Polis + Sertif") if sertif_confirmed else (list(polis_hit), "Polis only")


# =============================================================================
# NARROWING
# =============================================================================

def _narrow(
    suspend_row: dict, matched: list, rows: list, label: str,
    polis_ref_cols: list, slip_ref_cols: list, insured_ref_cols: list,
    polis_sus_cols: list, slip_sus_cols: list, insured_sus_cols: list,
    sertif_ref_cols: list = None, sertif_sus_cols: list = None,
) -> tuple:
    # Narrowing bertingkat NON-DESTRUKTIF: Step1 Polis<->Slip, Step2 Cert, Step3 Insured
    base_scenario = _map_scenario(label)
    if not matched:
        return [], "Unmatching"
    if len(matched) <= 1:
        return matched, base_scenario

    polis_values = _collect_all_values(suspend_row, polis_sus_cols, "polis_ori")
    slip_values  = _collect_all_values(suspend_row, slip_sus_cols,  "slip_ori")
    current_matched = matched
    current_label   = base_scenario

    # Step 1: konfirmasi field komplementer
    if "SLIP" in label:
        confirmed = [i for i in current_matched
                     if _ref_has_value(rows[i], polis_ref_cols, "polis_ori", polis_values)]
        if confirmed:
            current_matched = confirmed
            current_label = "Slip + Polis"
        else:
            current_label = "Slip only"

    if "POLIS" in label:
        confirmed = [i for i in current_matched
                     if _ref_has_value(rows[i], slip_ref_cols, "slip_ori", slip_values)]
        if not confirmed and slip_values:
            confirmed = [i for i in current_matched
                         if any(_slip_in_range(sv, rows[i]) for sv in slip_values)]
        if confirmed:
            current_matched = confirmed
            current_label = "Polis + Slip"
        else:
            current_label = "Polis only"

    # Step 2: narrowing sertifikat (NON-DESTRUKTIF)
    if len(current_matched) > 1 and sertif_ref_cols:
        sus_cert_vals: set = set(_clean_cert_str(v) for v in (_collect_clean_values(suspend_row, sertif_sus_cols) if sertif_sus_cols else []) if _clean_cert_str(v))
        for pv in [_normalize(suspend_row.get("polis_ori", ""))] + \
                  [_normalize(suspend_row.get(c, "")) for c in (polis_sus_cols or [])]:
            if pv:
                c = _extract_cert_from_polis(pv)
                if c:
                    sus_cert_vals.add(_clean_cert_str(c))
        sus_cert_vals.discard("")
        if sus_cert_vals:
            confirmed = [i for i in current_matched
                         if _ref_has_value(rows[i], sertif_ref_cols, "", list(sus_cert_vals))]
            if not confirmed:
                confirmed = [
                    i for i in current_matched
                    if any(
                        _cert_in_range(cv, rows[i].get(col, ""))
                        for cv in sus_cert_vals
                        for col in sertif_ref_cols
                    )
                ]
            if confirmed:
                current_matched = confirmed
                current_label   = current_label + " + Cert"

    # Step 3: narrowing insured (NON-DESTRUKTIF)
    if len(current_matched) > 1 and insured_ref_cols and insured_sus_cols:
        ins_vals = _collect_all_values(suspend_row, insured_sus_cols, "insured_ori")
        if ins_vals:
            confirmed = [i for i in current_matched
                         if _ref_has_value(rows[i], insured_ref_cols, "insured_ori", ins_vals)]
            if confirmed:
                current_matched = confirmed
                current_label   = current_label + " + Insured"

    return current_matched, current_label


def _narrow_by_currency(suspend_row: dict, matched: list, rows: list) -> list:
    # Saring kandidat: CCOS_CURR == CURR ORI. Jika semua beda, kembalikan asli.
    if len(matched) <= 1:
        return matched
    sus_curr = _normalize(suspend_row.get(SUSPEND_CURR_COL, ""))
    if not sus_curr:
        return matched
    filtered = [i for i in matched if _normalize(rows[i].get(OSBAL_CURR_COL, "")) == sus_curr]
    return filtered if filtered else matched


def _narrow_by_periode(
    suspend_row: dict, matched: list, rows: list,
    effective_sus_date=None, is_lineslip: bool = False,
) -> list:
    # Saring kandidat: RECEIPT DATE >= FAC_COM_DATE. LINESLIP: exact bulan+tahun.
    # Jika semua tersingkir, kembalikan asli (deteksi Beda Periode di run()).
    if len(matched) <= 1:
        return matched

    sus_date = effective_sus_date or suspend_row.get("_sus_date_parsed")
    if sus_date is None:
        return matched

    valid = []
    for i in matched:
        com_date = rows[i].get("_com_date_parsed")
        if com_date is None:
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

    return valid if valid else matched


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

    # R3: SLIPDB
    if slipdb_rows and lookup_slipdb:
        hits = _exact_match(polis_vals, lookup_slipdb[1])
        if hits:
            res_idx, ok = _resolve_facode([slipdb_rows[i] for i in hits], SLIPDB_FACODE_COL, facode_osbal_idx, osbal_rows)
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


def _run_polis_slip_pass(
    suspend_row: dict, rows: list, lookup: tuple,
    polis_ref_cols: list, slip_ref_cols: list, insured_ref_cols: list,
    polis_sus_cols: list, slip_sus_cols: list, insured_sus_cols: list,
    polis_sus_like_cols: list = None, slip_sus_like_cols: list = None,
    facode_col: str = None, gunakan_narrow_aca: bool = False,
    currency_narrowing: bool = False,
    effective_sus_date=None, is_lineslip: bool = False,
    sertif_ref_cols: list = None, sertif_sus_cols: list = None,
) -> tuple:
    # Jalankan satu pass matching polis/slip + narrowing currency/periode + narrowing utama
    lkp_slip_cln, lkp_polis_cln, lkp_ins_cln, _, _, _, tok_slip, tok_polis, tok_ins, _ = lookup

    matched, label = _match_polis_slip(
        suspend_row, lkp_slip_cln, lkp_polis_cln, {}, {},
        polis_sus_cols, slip_sus_cols, rows=rows,
        slip_ref_cols=slip_ref_cols, polis_ref_cols=polis_ref_cols,
        token_slip=tok_slip, token_polis=tok_polis,
        polis_sus_like_cols=polis_sus_like_cols, slip_sus_like_cols=slip_sus_like_cols,
    )

    if matched:
        # Narrowing wajib: currency dulu, lalu periode
        if currency_narrowing and len(matched) > 1:
            matched = _narrow_by_currency(suspend_row, matched, rows)
            if len(matched) > 1:
                matched = _narrow_by_periode(suspend_row, matched, rows,
                                             effective_sus_date=effective_sus_date,
                                             is_lineslip=is_lineslip)
        narrowed, scenario = _narrow(
            suspend_row, matched, rows, label,
            polis_ref_cols, slip_ref_cols, insured_ref_cols,
            polis_sus_cols, slip_sus_cols, insured_sus_cols,
            sertif_ref_cols=sertif_ref_cols, sertif_sus_cols=sertif_sus_cols,
        )
        return narrowed, scenario
    return [], "Unmatching"


def _run_insured_pass(
    suspend_row: dict, rows: list, lookup: tuple,
    insured_sus_cols: list, insured_ref_cols: list = None,
    currency_narrowing: bool = False,
    effective_sus_date=None, is_lineslip: bool = False,
    **_kwargs,
) -> tuple:
    # Jalankan satu pass matching insured (fallback) + narrowing currency/periode
    _, _, lkp_ins_cln, _, _, _, _, _, tok_ins, _ = lookup

    matched, label = _match_insured(
        suspend_row, lkp_ins_cln, {}, insured_sus_cols,
        rows=rows, insured_ref_cols=insured_ref_cols, token_insured=tok_ins,
    )
    if matched:
        if currency_narrowing and len(matched) > 1:
            matched = _narrow_by_currency(suspend_row, matched, rows)
            if len(matched) > 1:
                matched = _narrow_by_periode(suspend_row, matched, rows,
                                             effective_sus_date=effective_sus_date,
                                             is_lineslip=is_lineslip)
        return matched, _map_scenario(label)
    return [], "Unmatching"


# =============================================================================
# BORDERO NARROWING (OPEN COVER)
# =============================================================================

def _build_bordero_index(bordero_rows: list) -> dict:
    # Bangun index: fac_code -> list entri per baris bordero
    idx: dict = {}
    for row in bordero_rows:
        fac = _normalize(row.get(BORDERO_FAC_COL, ""))
        if not fac:
            continue
        bulan_str = str(row.get(BORDERO_BULAN_COL, "")).strip().upper()
        tahun_str = str(row.get(BORDERO_TAHUN_COL, "")).strip()
        month = _BULAN_MAP.get(bulan_str, 0)
        ym = (int(tahun_str), month) if month and tahun_str.isdigit() else None
        idx.setdefault(fac, []).append({
            "polis":   _normalize(row.get(BORDERO_POLIS_COL,   "")),
            "slip":    _normalize(row.get(BORDERO_SLIP_COL,    "")),
            "cert":    _normalize(row.get(BORDERO_CERT_COL,    "")),
            "insured": _normalize(row.get(BORDERO_INSURED_COL, "")),
            "period":  ym,
            "net":     row.get(BORDERO_NET_COL),
        })
    return idx


def _find_fac_codes_in_bordero(
    suspend_row: dict, bordero_idx: dict,
    polis_sus_cols: list, slip_sus_cols: list,
    bordero_polis_index: dict = None
) -> tuple:
    if not bordero_idx:
        return set(), False

    sus_polis = set(_collect_clean_values(suspend_row, polis_sus_cols))
    for c in [CLSDT_POLIS_COL, FAC_POLIS_COL, "polis_ori"]:
        v = _normalize(suspend_row.get(c, ""))
        if v:
            sus_polis.add(v)

    sus_cert = set()
    for pv in list(sus_polis) + [_normalize(suspend_row.get("polis_ori", ""))]:
        c = _extract_cert_from_polis(pv)
        if c:
            sus_cert.add(c)

    sus_slip = set(_collect_clean_values(suspend_row, slip_sus_cols))
    for c in [CLSDT_SLIP_COL, FAC_SLIP_COL, "slip_ori"]:
        v = _normalize(suspend_row.get(c, ""))
        if v:
            sus_slip.add(v)

    polis_f = {sp for sp in sus_polis if len(sp) >= 5}
    slip_f = {ss for ss in sus_slip if len(ss) >= 5}
    sus_amount = suspend_row.get("AMOUNT ORI")
    sus_amount_f = float(sus_amount) if sus_amount is not None else None

    sus_date = suspend_row.get("_sus_date_parsed")
    sus_ym = (sus_date.year, sus_date.month) if sus_date is not None else None
    sus_curr_val = _normalize(suspend_row.get("CURR ORI", ""))

    def mp(e):
        return bool(e["polis"] and len(e["polis"]) >= 5 and
                    any(sp == e["polis"] or sp in e["polis"] or e["polis"] in sp for sp in polis_f))
    def mc(e):
        return bool(e["cert"] and e["cert"] in sus_cert)
    def ms(e):
        return bool(e["slip"] and len(e["slip"]) >= 5 and
                    any(ss == e["slip"] or ss in e["slip"] or e["slip"] in ss for ss in slip_f))
    def mnet(e):
        if sus_amount_f is None or e.get("net") is None: return False
        try:
            b_net = float(e["net"])
            return abs(abs(b_net) - abs(sus_amount_f)) < 0.05
        except: return False
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
        
        # 1. Exact match (sangat cepat)
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
        # Fallback (sangat lambat)
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
            if valid: cert_filtered[fac] = valid
        if cert_filtered: candidates = cert_filtered

    if sus_slip:
        slip_filtered = {}
        for fac, rows in candidates.items():
            valid = [r for r in rows if ms(r)]
            if valid: slip_filtered[fac] = valid
        if slip_filtered: candidates = slip_filtered

    wajib_filtered = {}
    for fac, rows in candidates.items():
        valid = [r for r in rows if mcurr(r) and mper(r)]
        if valid: wajib_filtered[fac] = valid
    if wajib_filtered: candidates = wajib_filtered

    if sus_amount_f is not None:
        net_filtered = {}
        for fac, rows in candidates.items():
            valid = [r for r in rows if mnet(r)]
            if valid: net_filtered[fac] = valid
        if net_filtered: candidates = net_filtered

    return set(candidates.keys()), True

def _narrow_by_bordero_mc_period(suspend_row: dict, fac_codes: set, bordero_idx: dict) -> set:
    # Narrowing cara kedua Marine Cargo: RECEIPT DATE + 1 bulan = bordero BULAN/TAHUN
    if len(fac_codes) <= 1 or not bordero_idx:
        return fac_codes
    sus_date = suspend_row.get("_sus_date_parsed")
    if sus_date is None:
        return fac_codes
    try:
        target_ym = (sus_date + pd.DateOffset(months=1))
        target_ym = (target_ym.year, target_ym.month)
    except Exception:
        return fac_codes
    narrowed = {f for f in fac_codes
                if any(e["period"] == target_ym for e in bordero_idx.get(f, []) if e.get("period"))}
    return narrowed if narrowed and len(narrowed) < len(fac_codes) else fac_codes


def _narrow_by_amount_net(
    suspend_row: dict, fac_codes: set, bordero_idx: dict,
    tolerance_pct: float = NET_TOLERANCE_PCT, mode: str = NET_MATCH_MODE,
) -> tuple:
    # Last choice: |AMOUNT ORI| ≈ NET bordero -> resolve tepat 1 fac code
    if len(fac_codes) <= 1 or not bordero_idx:
        return fac_codes, False
    try:
        amount_ori = abs(float(suspend_row.get("AMOUNT ORI", 0) or 0))
    except (ValueError, TypeError):
        return fac_codes, False
    if amount_ori == 0:
        return fac_codes, False

    tolerance = amount_ori * tolerance_pct

    def net_f(e):
        try:
            v = e.get("net")
            return float(v) if v is not None else None
        except (ValueError, TypeError):
            return None

    sus_cert = None
    for col in ["polis_ori", "clean polis 1", "clean polis 2"]:
        v = _normalize(suspend_row.get(col, ""))
        if v:
            c = _extract_cert_from_polis(v)
            if c:
                sus_cert = c
                break

    matched_facs = set()
    for fac in fac_codes:
        entries = bordero_idx.get(fac, [])
        if not entries:
            continue
        if sus_cert:
            filtered = [e for e in entries if e.get("cert") == sus_cert]
            if filtered:
                entries = filtered
        if mode == "per_row":
            for e in entries:
                n = net_f(e)
                if n is not None and abs(abs(n) - amount_ori) <= tolerance:
                    matched_facs.add(fac)
                    break
        elif mode == "sum_all":
            nets = [net_f(e) for e in entries if net_f(e) is not None]
            if nets and abs(abs(sum(nets)) - amount_ori) <= tolerance:
                matched_facs.add(fac)
        elif mode == "sum_by_curr":
            sus_curr = _normalize(suspend_row.get(SUSPEND_CURR_COL, ""))
            nets = [net_f(e) for e in entries
                    if net_f(e) is not None and _normalize(e.get("curr", "")) == sus_curr]
            if nets and abs(abs(sum(nets)) - amount_ori) <= tolerance:
                matched_facs.add(fac)

    return (matched_facs, True) if len(matched_facs) == 1 else (fac_codes, False)


# =============================================================================
# FAC CODE RESOLUTION
# =============================================================================

def _resolve_facode(ref_rows_matched: list, facode_col: str,
                    facode_osbal_idx: dict, osbal_rows: list) -> tuple:
    # Resolve FAC_CODE dari SLIPDB/FACUL ke indeks baris OSBAL
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
    # Agregasi nilai finansial OSBAL; all_fac_osbal_rows untuk BAL_DUE akurat
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
        "CCOS_OR_BAL":   np.nan if multi_fac else join_unique("CCOS_OR_BAL"),
        "CCOS_BAL_DUE":  sum_col("CCOS_BAL_DUE", bal_due_rows),
    }


def _facode_label(osbal_rows: list, source: str) -> str:
    # Kembalikan FAC_CODE tunggal, atau 'facode lebih dari 1' jika multiple
    if not osbal_rows or not osbal_rows[0]:
        return ""
    col = OSBAL_FACODE_COL if source == "OSBAL" else (
          SLIPDB_FACODE_COL if source == "SLIPDB" else FACUL_FACODE_COL)
    codes = {str(r.get(col, "")).strip() for r in osbal_rows if str(r.get(col, "")).strip()}
    return "facode lebih dari 1" if len(codes) > 1 else next(iter(codes), "")


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
    if not source or not base_scenario or str(base_scenario).strip().upper() in ("UNMATCHING", "UNMATCHED"):
        return "Unmatching"

    base = _clean_base_match(base_scenario)
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
    suspend_row: dict, source: str, scenario: str, osbal_rows: list,
    osbal_count: int = 0, resolved: bool = True,
    suspend_count: int = 1, all_fac_osbal_rows: list = None,
) -> dict:
    # Bangun satu baris output (38 kolom standar alternating Polis & Sertifikat)
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
        "INSURED_ORI":            suspend_row.get("insured_ori") or suspend_row.get("INSURED") or suspend_row.get("INSURED_ORI", ""),
        "INSURED_1":              suspend_row.get("clean insured 1") or suspend_row.get("INSURED_CLEAN", ""),
        "INSURED_2":              suspend_row.get("clean insured 2",   ""),
        "CURR ORI":               suspend_row.get("CURR ORI",          ""),
        "AMOUNT ORI":             suspend_row.get("AMOUNT ORI",        ""),
        "CURR PAY":               suspend_row.get("CURR PAY",          ""),
        "AMOUNT PAY":             suspend_row.get("AMOUNT PAY",        ""),
        "CCOS_OR_BAL":            ccos["CCOS_OR_BAL"],
        "CCOS_BAL_DUE":           ccos["CCOS_BAL_DUE"],
        "POLIS":                  suspend_row.get("POLIS") or suspend_row.get("polis_ori") or suspend_row.get("POLIS_ORI", ""),
        "POLICY_CLEAN_1":         suspend_row.get("POLICY_CLEAN_1") or suspend_row.get("clean polis 1") or suspend_row.get("POLIS_CLEAN") or suspend_row.get("POLIS_CLN", ""),
        "SERTIF_CLEAN_1":         _format_sertif(suspend_row.get("SERTIF_CLEAN_1") or suspend_row.get("clean sertif 1") or suspend_row.get("CERTIFICATE_1") or suspend_row.get("SERTIF_CLN_1") or suspend_row.get("SERTIF_CLN", "")),
        "POLICY_CLEAN_2":         suspend_row.get("POLICY_CLEAN_2") or suspend_row.get("clean polis 2", ""),
        "SERTIF_CLEAN_2":         _format_sertif(suspend_row.get("SERTIF_CLEAN_2") or suspend_row.get("clean sertif 2") or suspend_row.get("CERTIFICATE_2") or suspend_row.get("SERTIF_CLN_2", "")),
        "POLICY_CLEAN_3":         suspend_row.get("POLICY_CLEAN_3") or suspend_row.get("clean polis 3", ""),
        "SERTIF_CLEAN_3":         _format_sertif(suspend_row.get("SERTIF_CLEAN_3") or suspend_row.get("clean sertif 3") or suspend_row.get("CERTIFICATE_3") or suspend_row.get("SERTIF_CLN_3", "")),
        "POLICY_CLEAN_4":         suspend_row.get("POLICY_CLEAN_4") or suspend_row.get("clean polis 4", ""),
        "POLICY_CLEAN_5":         suspend_row.get("POLICY_CLEAN_5") or suspend_row.get("clean polis 5", ""),
        "SLIP_NO":                suspend_row.get("SLIP_NO") or suspend_row.get("slip_ori") or suspend_row.get("SLIP NO", ""),
        "SLIP_NO_CLN":            suspend_row.get("SLIP_NO_CLN") or suspend_row.get("clean slip 1") or suspend_row.get("SLIP_NO_CLEAN", ""),
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
    # Merge beberapa baris suspend (fac code sama): sum AMOUNT ORI, join-unique sisanya
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
    # Hitung AMOUNT_ORI_MIN1, DIFERENCE, dan FLAG_PROD (10 kategori)
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
# PIPELINE
# =============================================================================

def _load_excel(path: str, label: str, optional: bool = False) -> tuple:
    # Load Excel ke (list[dict], list[str]) dengan pickle cache otomatis
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


def run() -> None:
    """Pipeline utama V1: matching suspend vs OSBAL/SLIPDB/FACUL, row asli dipertahankan."""
    t_start = time.perf_counter()
    print("\n" + "=" * 60, flush=True)
    print("  PRODUCTION SCRIPT — SUSPEND MATCHING  (ACA) — V1", flush=True)
    print("=" * 60, flush=True)

    # [1] Load data
    print("\n[1/6] Loading data ...", flush=True)
    suspend_rows, suspend_cols = _load_excel(SUSPEND_FILE, "SUSPEND")
    osbal_rows,   osbal_cols   = _load_excel(OSBAL_FILE,   "OSBAL")
    facul_rows,   facul_cols   = _load_excel(FACUL_FILE,   "FACUL")
    slipdb_rows,  slipdb_cols  = _load_excel(SLIPDB_FILE,  "SLIPDB", optional=True)
    bordero_rows, bordero_cols = _load_excel(BORDERO_FILE, "BORDERO", optional=True)

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

    # Pre-parse LINESLIP flags
    print("  Pre-parsing LINESLIP flags ...", flush=True)
    lineslip_count = 0
    for r in suspend_rows:
        r["_is_lineslip"] = _is_lineslip_row(r, insured_sus)
        if r["_is_lineslip"]:
            lineslip_count += 1
            parsed = r.get("_sus_date_parsed")
            try:
                r["_sus_date_lineslip"] = (parsed - pd.DateOffset(months=1)) if parsed is not None else None
            except Exception:
                r["_sus_date_lineslip"] = None
        else:
            r["_sus_date_lineslip"] = None
    print(f"  -> {lineslip_count:,} baris LINESLIP terdeteksi", flush=True)

    polis_osbal   = _get_clean_cols(osbal_cols, "clean polis")
    slip_osbal    = _get_clean_cols(osbal_cols, "clean slip")
    insured_osbal = _get_clean_cols(osbal_cols, "clean insured")
    sertif_osbal  = _get_clean_cols(osbal_cols, "clean sertif")

    polis_facul   = _get_clean_cols(facul_cols, "clean polis")
    slip_facul    = _get_clean_cols(facul_cols, "clean slip")
    insured_facul = _get_clean_cols(facul_cols, "clean insured")
    sertif_facul  = _get_clean_cols(facul_cols, "clean sertif")

    polis_slipdb   = _get_clean_cols(slipdb_cols, "clean polis")
    slip_slipdb    = _get_clean_cols(slipdb_cols, "clean slip")
    insured_slipdb = _get_clean_cols(slipdb_cols, "clean insured")
    sertif_slipdb  = _get_clean_cols(slipdb_cols, "clean sertif")

    if not polis_slipdb:
        polis_slipdb = [c for c in ["CLSDT_POLICY_NO","FAC_POLICY_NO","POLIS","POLICY_NO","POLICY NO"] if c in slipdb_cols]
    if not slip_slipdb:
        slip_slipdb  = [c for c in ["CLSDT_SLIP_NO","FAC_SLIP","SLIP","SLIP_NO","SLIP NO"] if c in slipdb_cols]
    if not insured_slipdb:
        insured_slipdb = [c for c in ["FAC_INSURED","INSURED","CEDANT_NAME","CEDANT NAME"] if c in slipdb_cols]

    global SLIPDB_FACODE_COL
    if   "FAC_CODE"     in slipdb_cols: SLIPDB_FACODE_COL = "FAC_CODE"
    elif "FAC CODE"     in slipdb_cols: SLIPDB_FACODE_COL = "FAC CODE"
    elif "CCOS_REF_CODE" in slipdb_cols: SLIPDB_FACODE_COL = "CCOS_REF_CODE"

    # Tambahkan CLSDT ke depan kolom referensi agar ter-index
    def _prepend(col, cols_list, header):
        return ([col] + cols_list) if col in header and col not in cols_list else cols_list

    polis_osbal  = _prepend("CLSDT_POLICY_NO", polis_osbal,  osbal_cols)
    slip_osbal   = _prepend("CLSDT_SLIP_NO",   slip_osbal,   osbal_cols)
    polis_facul  = _prepend("CLSDT_POLICY_NO", polis_facul,  facul_cols)
    slip_facul   = _prepend("CLSDT_SLIP_NO",   slip_facul,   facul_cols)
    polis_slipdb = _prepend("CLSDT_POLICY_NO", polis_slipdb, slipdb_cols)
    slip_slipdb  = _prepend("CLSDT_SLIP_NO",   slip_slipdb,  slipdb_cols)

    # [3] Bangun lookup index
    print("\n[3/6] Building lookup indexes ...", flush=True)
    lookup_osbal  = _build_lookup(osbal_rows,  polis_osbal,  slip_osbal,  insured_osbal,  OSBAL_FACODE_COL,  label="OSBAL")
    lookup_facul  = _build_lookup(facul_rows,  polis_facul,  slip_facul,  insured_facul,  FACUL_FACODE_COL,  label="FACUL")
    lookup_slipdb = _build_lookup(slipdb_rows, polis_slipdb, slip_slipdb, insured_slipdb, SLIPDB_FACODE_COL, label="SLIPDB")

    facode_osbal_idx = lookup_osbal[9]

    excluded_osbal   = {i for i, r in enumerate(osbal_rows) if _is_excluded_ref_row(r, polis_osbal, slip_osbal)}
    sertif_osbal_idx = _build_sertif_index(osbal_rows, sertif_osbal, excluded=excluded_osbal)
    print(f"  Sertif OSBAL index: {len(sertif_osbal_idx):,} sertif unik", flush=True)
    if not sertif_osbal:
        print("  [INFO] Kolom clean sertif tidak ada di OSBAL — Pass A dilewati", flush=True)

    # Kwargs tetap per tabel referensi
    kw_osbal = dict(
        rows=osbal_rows, lookup=lookup_osbal,
        polis_ref_cols=polis_osbal, slip_ref_cols=slip_osbal, insured_ref_cols=insured_osbal,
        facode_col=OSBAL_FACODE_COL, currency_narrowing=True,
        sertif_ref_cols=sertif_osbal,
    )
    kw_facul = dict(
        rows=facul_rows, lookup=lookup_facul,
        polis_ref_cols=polis_facul, slip_ref_cols=slip_facul, insured_ref_cols=insured_facul,
        facode_col=FACUL_FACODE_COL, gunakan_narrow_aca=True, sertif_ref_cols=sertif_facul,
    )
    kw_slipdb = dict(
        rows=slipdb_rows, lookup=lookup_slipdb,
        polis_ref_cols=polis_slipdb, slip_ref_cols=slip_slipdb, insured_ref_cols=insured_slipdb,
        facode_col=SLIPDB_FACODE_COL, gunakan_narrow_aca=True, sertif_ref_cols=sertif_slipdb,
    )

    # Build bordero index untuk fitur override di Fase 6
    bordero_idx = _build_bordero_index(bordero_rows) if bordero_rows else {}
    bordero_polis_index = {}
    if bordero_idx:
        for fac, rows in bordero_idx.items():
            for r in rows:
                p = r.get("polis")
                if p and len(p) >= 5:
                    bordero_polis_index.setdefault(p, []).append((fac, r))
    
    print(f"  Bordero index: {len(bordero_idx):,} FAC CODE entries, {len(bordero_polis_index):,} unique polis" if bordero_idx
          else "  Bordero: tidak tersedia, dilewati", flush=True)

    # [4] Matching per baris
    total = len(suspend_rows)
    print(f"\n[4/6] Matching {total:,} suspend rows ...", flush=True)
    t4 = time.perf_counter()
    raw_results = []

    for n, sus in enumerate(suspend_rows, 1):
        if n % 1000 == 0:
            print(f"    Memproses baris {n:,} ...", flush=True)

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
        
        kw_sus_row = dict(
            polis_sus_cols=eff_polis, slip_sus_cols=slip_sus, insured_sus_cols=insured_sus,
            polis_sus_like_cols=like_polis, slip_sus_like_cols=like_slip,
            effective_sus_date=effective_date, is_lineslip=_is_ls,
            sertif_sus_cols=sertif_sus,
        )

        # FASE 1: PENCARIAN AWAL BERDASARKAN POLIS
        polis_vals = _collect_clean_values(sus, eff_polis)
        found_by_polis = False
        if polis_vals:
            osb_hits = _exact_match(polis_vals, lookup_osbal[1])
            if osb_hits:
                matched_osbal_idx.update(osb_hits)
                found_by_polis = True
                
            if slipdb_rows:
                sld_hits = _exact_match(polis_vals, lookup_slipdb[1])
                if sld_hits:
                    res_idx, _ = _resolve_facode([slipdb_rows[idx] for idx in sld_hits], SLIPDB_FACODE_COL, facode_osbal_idx, osbal_rows)
                    if res_idx:
                        matched_osbal_idx.update(res_idx)
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
                    
                if slipdb_rows and not osb_hits:
                    sld_hits = _exact_match(slip_vals, lookup_slipdb[0])
                    if sld_hits:
                        res_idx, _ = _resolve_facode([slipdb_rows[idx] for idx in sld_hits], SLIPDB_FACODE_COL, facode_osbal_idx, osbal_rows)
                        if res_idx:
                            matched_osbal_idx.update(res_idx)
                            source = "SLIPDB"
                            scenario = "Slip only"
                            resolved = True
                            
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
                elif slipdb_rows:
                    sld_hits = _exact_match(ins_vals, lookup_slipdb[2])
                    if sld_hits:
                        res_idx, _ = _resolve_facode([slipdb_rows[idx] for idx in sld_hits], SLIPDB_FACODE_COL, facode_osbal_idx, osbal_rows)
                        if res_idx:
                            matched_osbal_idx.update(res_idx)
                            source = "SLIPDB"
                            scenario = "Insured only"
                            resolved = True
                elif facul_rows and not matched_osbal_idx:
                    fac_hits = _exact_match(ins_vals, lookup_facul[2])
                    if fac_hits:
                        res_idx, _ = _resolve_facode([facul_rows[idx] for idx in fac_hits], FACUL_FACODE_COL, facode_osbal_idx, osbal_rows)
                        if res_idx:
                            matched_osbal_idx.update(res_idx)
                            source = "FACUL"
                            scenario = "Insured only"
                            resolved = True

        # FASE 4: NARROWING WAJIB (Curr & Periode)
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

            # Step 1b: Slip Numeric Range Check (antara clean slip 1 dan clean slip 2)
            if len(matched_list) > 1 and slip_vals:
                confirmed_slip = [
                    idx for idx in matched_list
                    if any(_slip_in_range(sv, osbal_rows[idx]) for sv in slip_vals)
                ]
                if confirmed_slip:
                    matched_list = confirmed_slip
                    if "Slip" not in scenario:
                        scenario = "Polis + Slip" if "Polis" in scenario else scenario
            
            # Step 2: Sertifikat Exact & In-Range Matching (dengan pembersihan float .0)
            sus_cert_vals = set(_clean_cert_str(v) for v in _collect_clean_values(sus, sertif_sus) if _clean_cert_str(v))
            for pv in [_normalize(sus.get("polis_ori", ""))] + [_normalize(sus.get(c, "")) for c in (eff_polis or [])]:
                if pv:
                    c = _extract_cert_from_polis(pv)
                    if c: sus_cert_vals.add(_clean_cert_str(c))
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
            ins_vals = _collect_all_values(sus, insured_sus, "insured_ori")
            if ins_vals:
                confirmed = [idx for idx in matched_list if _ref_has_value(osbal_rows[idx], insured_osbal, "insured_ori", ins_vals)]
                if confirmed:
                    matched_list = confirmed
                    if "Insured" not in scenario:
                        scenario = scenario + " + Insured"
                    
            matched_osbal_idx = set(matched_list)

        # FASE 6: BORDERO (Open Cover) OVERRIDE
        is_bordero = False
        if bordero_idx and (found_by_polis or matched_osbal_idx):
            bordero_facs, ok = _find_fac_codes_in_bordero(sus, bordero_idx, eff_polis, slip_sus, bordero_polis_index=bordero_polis_index)
            if ok and bordero_facs:
                res_idx = []
                for fac in bordero_facs:
                    if fac in facode_osbal_idx:
                        res_idx.extend(facode_osbal_idx[fac])
                
                if res_idx:
                    if res_idx and list(res_idx) != list(matched_osbal_idx):
                        matched_osbal_idx = set(res_idx)
                    is_bordero = True
                    resolved = True

        if is_beda_curr or is_beda_per:
            osbal_ref = [osbal_rows[i] for i in matched_list]
            osbal_count = 0  # quantitative values clear out
            resolved = False
        elif not matched_osbal_idx:
            source = None; scenario = "Unmatching"
            osbal_ref = [{}]; osbal_count = 0; resolved = False
        else:
            osbal_ref = [osbal_rows[i] for i in matched_osbal_idx]
            osbal_count = len(osbal_ref)

        # Clean up confusing scenario names (e.g. "Polis only + Insured" -> "Polis + Insured")
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

    # [4b] Bordero narrowing
    print("\n[4c/6] Fac code claim & osbal accumulation (V1: row asli dipertahankan) ...", flush=True)

    claimed: set = set()
    for r in raw_results:
        if len(r["fac_codes"]) == 1 and r["source"] is not None and not r["is_beda_curr"] and not r["is_beda_per"]:
            claimed |= r["fac_codes"]

    for r in raw_results:
        if len(r["fac_codes"]) > 1:
            remaining = r["fac_codes"] - claimed
            if len(remaining) == 1 and r["source"] is not None and not r["is_beda_curr"] and not r["is_beda_per"]:
                r["fac_codes"] = remaining; claimed |= remaining
                fac = next(iter(remaining))
                r["osbal_idx"] = [i for i in r["osbal_idx"]
                                   if _normalize(osbal_rows[i].get(OSBAL_FACODE_COL, "")) == fac]
                r["osbal_count"] = len(r["osbal_idx"])

    # Group per fac untuk shared osbal, tapi simpan semua row suspend
    groups: dict = {}
    final:  list = []

    for r in raw_results:
        r.setdefault("suspend_count", 1)
        if (len(r["fac_codes"]) == 1 and r["source"] is not None
                and not r["is_beda_curr"] and not r["is_beda_per"]):
            groups.setdefault(next(iter(r["fac_codes"])), []).append(r)
        else:
            final.append(r)

    for fac, group in groups.items():
        if len(group) == 1:
            group[0]["suspend_count"] = 1; final.append(group[0]); continue
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

    print(f"  {len(raw_results):,} suspend rows -> {len(final):,} output rows (row asli dipertahankan)", flush=True)

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
            is_bordero=r.get("is_bordero", False),
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
            con.execute(text('DELETE FROM "SUSPENSE_DATA_SUSPENSE_V1" WHERE "CEDANT NAME" = \'PT. ASURANSI CENTRAL ASIA\';'))
            con.execute(text('DELETE FROM "SUSPENSE_CLEAN_2026" WHERE "CEDANT NAME" = \'PT. ASURANSI CENTRAL ASIA\';'))
        df_db.to_sql("SUSPENSE_DATA_SUSPENSE_V1", con=engine, if_exists="append", index=False)
        df_db.to_sql("SUSPENSE_CLEAN_2026", con=engine, if_exists="append", index=False)
        print(f"      -> {len(df_db):,} rows berhasil di-export ke 'SUSPENSE_DATA_SUSPENSE_V1' & 'SUSPENSE_CLEAN_2026'")
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