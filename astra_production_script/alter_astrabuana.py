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
from sqlalchemy import create_engine


# =============================================================================
# CONFIG — PT ASURANSI ASTRA BUANA
# =============================================================================

SUSPEND_FILE = os.path.join("data", "astrabuana_output_suspend.xlsx")
OSBAL_FILE   = os.path.join("data", "astrabuana_output_osbal.xlsx")
FACUL_FILE   = os.path.join("data", "astrabuana_output_facul.xlsx")
OUTPUT_FILE  = os.path.join("data", "final_output_astrabuana.xlsx")

OSBAL_FACODE_COL  = "CCOS_REF_CODE"
FACUL_FACODE_COL  = "FAC_CODE"

# Currency columns (ATURAN 1)
SUSPEND_CURR_COL = "CURR ORI"   # Suspend: kolom currency asal
OSBAL_CURR_COL   = "CCOS_CURR"  # OSBAL  : patokan kebenaran currency

# Periode columns (ATURAN BEDA PERIODE)
SUSPEND_DATE_COL = "RECEIPT DATE"  # Suspend: tanggal pembayaran
OSBAL_DATE_COL   = "FAC_COM_DATE"  # OSBAL  : tanggal awal pertanggungan

# Reference-row exclusion markers (ATURAN 2)
_EXCLUDED_REF_MARKERS = ["HUTANG PIUTANG", "DATA SUSPENSE"]

# LINESLIP treatment markers
_LINESLIP_MARKERS = ("LINESLIP", "LINE SLIP")

# CLSDT general terms yang harus dianggap kosong (fallback ke FAC)
_GENERAL_CLSDT_MARKERS = {"TBA", "VARIOUS", "P"}

# -----------------------------------------------------------------------
# Kolom "ori" (raw / belum dibersihkan) di masing-masing tabel
# -----------------------------------------------------------------------
# SUSPEND
SUS_INSURED_ORI_COL = "INSURED"
SUS_POLIS_ORI_COL   = "POLIS"
SUS_SLIP_ORI_COL    = "SLIP NO"

# Kolom CLSDT prioritas di Suspend (tidak ada di data Astra Buana → kosong)
CLSDT_POLIS_COL = "CLSDT_POLICY_NO"  # tidak ada → otomatis fallback
CLSDT_SLIP_COL  = "CLSDT_SLIP_NO"    # tidak ada → otomatis fallback
FAC_POLIS_COL   = "FAC_POLICY_NO"    # kolom polis fallback (tidak ada di suspend astra)
FAC_SLIP_COL    = "FAC_SLIP"          # kolom slip  fallback (tidak ada di suspend astra)

# -----------------------------------------------------------------------
# Prefix kolom clean per tabel (dipakai di _get_clean_cols_for())
# -----------------------------------------------------------------------
# Suspend
_SUS_INSURED_PREFIX  = "INSURED_CLEAN"
_SUS_POLIS_PREFIX    = "POLIS_CLEAN"
_SUS_SLIP_PREFIX     = "SLIP_NO_CLEAN"
_SUS_SERTIF_PREFIX   = "CERTIFICATE_"

# OSBAL
_OSBAL_INSURED_PREFIX  = "FAC_INSURED_CLN_"
_OSBAL_POLIS_PREFIX    = "FAC_POLICY_NO_CLEAN_"
_OSBAL_SLIP_PREFIX     = "FAC_SLIP_CLEAN_"
_OSBAL_SERTIF_PREFIX   = "CERTIFICATE_"

# FACUL  (prefix polis berbeda: FAC_POLICY_CLEAN_ bukan FAC_POLICY_NO_CLEAN_)
_FACUL_INSURED_PREFIX  = "FAC_INSURED_CLN_"
_FACUL_POLIS_PREFIX    = "FAC_POLICY_CLEAN_"
_FACUL_SLIP_PREFIX     = "FAC_SLIP_CLEAN_"
_FACUL_SERTIF_PREFIX   = "CERTIFICATE_"

_MIN_TOKEN_LEN = 3
_TOKEN_RE      = re.compile(r"[^A-Z0-9]+")
_DELIMITER_RE  = re.compile(r"[,;|+\s]+")

# Regex untuk mengekstrak nomor sertifikat 6-digit
_CERT_RE = re.compile(r'-([0-9]{6})(?:[^0-9]|$)')

# Suspend columns to comma-join when multiple rows are accumulated into one output row
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
    "SERTIF_CLN_1", "SERTIF_CLN_2", "SERTIF_CLN_3",
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
    """Cached core normalization — hanya menerima str, dipanggil oleh _normalize()."""
    text = s.strip().upper().replace("S/D", "SD")
    return re.sub(r"\s+", " ", text) if "  " in text else text


def _expand_sertif_range(val: str) -> list:
    """Expand '100-110', '110-100', '100 SD 110' into individual 6-digit strings."""
    if not val:
        return []
    clean_val = re.sub(r'\s*SD\s*|\s*S/D\s*', '-', val, flags=re.IGNORECASE)
    clean_val = re.sub(r'\s+', '', clean_val)

    if clean_val.isdigit() and 1 <= len(clean_val) <= 6:
        return [clean_val.zfill(6)]

    m = re.match(r'^(\d{1,6})-(\d{1,6})$', clean_val)
    if m:
        start = int(m.group(1))
        end = int(m.group(2))
        step = 1 if start <= end else -1
        if abs(start - end) <= 5000:
            return [str(i).zfill(6) for i in range(start, end + step, step)]
    return []


def _normalize(value) -> str:
    """Uppercase, strip, collapse whitespace, return '' for null.

    Menggunakan lru_cache via _norm_cached() untuk menghindari re-komputasi
    nilai yang sama (dipanggil jutaan kali per run).
    """
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    return _norm_cached(str(value))


def _is_excluded_ref_row(row: dict, polis_cols: list, slip_cols: list) -> bool:
    """True jika salah satu kolom Polis/Slip (clean & ori) pada baris referensi
    mengandung salah satu _EXCLUDED_REF_MARKERS (ATURAN 2).

    Cek dilakukan di: polis_cols (clean polis N), "polis_ori",
    slip_cols (clean slip N), "slip_ori" — pakai _normalize() untuk
    case-insensitive dan whitespace-safe comparison. Substring check,
    BUKAN exact match, karena variasi teks banyak (ada tahun, "IDR ENG", dll).
    """
    check_cols = list(polis_cols) + ["polis_ori", "FAC_POLICY_NO", "POLIS"] + list(slip_cols) + ["slip_ori", "FAC_SLIP", "SLIP_NO"]
    for col in check_cols:
        val = _normalize(row.get(col, ""))
        if val and any(marker in val for marker in _EXCLUDED_REF_MARKERS):
            return True
    return False


def _tokenize(text: str) -> frozenset:
    """Split string into a frozenset of alphanumeric tokens (min length 3)."""
    if not text:
        return frozenset()
    return frozenset(t for t in _TOKEN_RE.split(text) if len(t) >= _MIN_TOKEN_LEN)


def _tokens_intersect(tokens_a: frozenset, tokens_b: frozenset) -> bool:
    """True if any token pair shares an exact match or one contains the other."""
    for a in tokens_a:
        for b in tokens_b:
            if a == b or a in b or b in a:
                return True
    return False


def _is_lineslip_row(sus_row: dict, insured_cols: list) -> bool:
    """True jika salah satu kolom insured suspend mengandung 'LINESLIP' atau 'LINE SLIP'.

    Kolom yang dicek: insured_cols (clean insured N) + insured_ori.
    Treatment: RECEIPT DATE + 1 bulan digunakan sebagai tanggal efektif untuk
    perbandingan dengan FAC_COM_DATE (OSBAL) — karena pembayaran LINESLIP selalu
    datang satu bulan setelah FAC_COM_DATE.
    """
    check_cols = list(insured_cols) + ["insured_ori", "INSURED_ORI", "INSURED"]
    for col in check_cols:
        val = _normalize(sus_row.get(col, ""))
        if val and any(m in val for m in _LINESLIP_MARKERS):
            return True
    return False


def _get_clean_cols(columns: list, prefix) -> list:
    """Return column names that match prefix (case-insensitive, supports clean X and X_CLEAN)."""
    if not columns:
        return []
    pfx = str(prefix).strip().lower()
    alt_map = {
        "polis_clean": ["polis_clean", "polis_cln", "policy_clean", "clean polis"],
        "slip_no_clean": ["slip_no_clean", "slip_no_cln", "slip_clean", "clean slip"],
        "insured_clean": ["insured_clean", "fac_insured_cln", "clean insured"],
        "certificate_": ["certificate_", "sertif_clean", "sertif_cln", "clean sertif"],
        "fac_insured_cln_": ["fac_insured_cln_", "insured_clean_", "clean insured "],
        "fac_policy_no_clean_": ["fac_policy_no_clean_", "policy_clean_", "clean polis "],
        "fac_policy_clean_": ["fac_policy_clean_", "policy_clean_", "clean polis "],
        "fac_slip_clean_": ["fac_slip_clean_", "slip_clean_", "clean slip "],
    }
    prefixes = alt_map.get(pfx, [pfx])
    matched = []
    for c in columns:
        cl = str(c).strip().lower()
        if any(cl.startswith(p) for p in prefixes):
            matched.append(c)
    return matched


def _get_clean_cols_for(columns: list, table: str, field: str) -> list:
    """Return clean columns untuk tabel dan field tertentu dengan prefix yang tepat.

    table  : 'suspend', 'osbal', 'facul'
    field  : 'polis', 'slip', 'insured', 'sertif'
    """
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
    if not prefix:
        return []
    return _get_clean_cols(columns, prefix)


def _map_scenario(label: str) -> str:
    """Map internal match label to official scenario name."""
    if not label or label in ("Unmatching", "UNMATCHED"):
        return "Unmatching"
    if "INSURED" in label:
        return "Insured only"
    if "SLIP" in label:
        return "Slip only"
    if "POLIS" in label:
        return "Polis only"
    return label


# =============================================================================
# ROW DICT HELPERS
# =============================================================================

def _get_norm(row: dict, col: str) -> str:
    return _normalize(row.get(col, ""))


def _collect_clean_values(row: dict, cols: list) -> list:
    """Return unique non-empty normalized values from the given columns."""
    seen, result = set(), []
    for col in cols:
        val = _normalize(row.get(col, ""))
        if val and val not in seen:
            seen.add(val)
            result.append(val)
    return result


def _collect_all_values(row: dict, clean_cols: list, ori_col: str) -> list:
    """Combine clean column values into a unique list (ori column disabled)."""
    values = set(_collect_clean_values(row, clean_cols))
    # ori = _normalize(row.get(ori_col, ""))  # DISABLED
    return list(values)


def _get_effective_sus_cols(
    sus_row:        dict,
    base_polis_cols: list,
    base_slip_cols:  list,
) -> tuple:
    """Return (eff_polis, eff_slip, like_polis, like_slip) untuk satu baris suspend.

    Data Astra Buana tidak memiliki CLSDT_POLICY_NO / CLSDT_SLIP_NO / FAC_POLICY_NO / FAC_SLIP
    di tabel suspend — fungsi ini akan otomatis fallback ke base_polis_cols / base_slip_cols.
    """
    # --- Polis ---
    clsdt_polis = _normalize(sus_row.get(CLSDT_POLIS_COL, ""))
    if clsdt_polis in _GENERAL_CLSDT_MARKERS:
        clsdt_polis = ""
    fac_polis = _normalize(sus_row.get(FAC_POLIS_COL, ""))

    if clsdt_polis:
        eff_polis  = [CLSDT_POLIS_COL]
        like_polis = [CLSDT_POLIS_COL]
    else:
        eff_polis  = base_polis_cols
        like_polis = list(base_polis_cols) + ([FAC_POLIS_COL] if fac_polis else [])

    # --- Slip ---
    clsdt_slip = _normalize(sus_row.get(CLSDT_SLIP_COL, ""))
    if clsdt_slip in _GENERAL_CLSDT_MARKERS:
        clsdt_slip = ""
    fac_slip = _normalize(sus_row.get(FAC_SLIP_COL, ""))

    if clsdt_slip:
        eff_slip  = [CLSDT_SLIP_COL]
        like_slip = [CLSDT_SLIP_COL]
    else:
        eff_slip  = base_slip_cols
        like_slip = list(base_slip_cols) + ([FAC_SLIP_COL] if fac_slip else [])

    return eff_polis, eff_slip, like_polis, like_slip


# =============================================================================
# INDEX BUILDERS
# =============================================================================

def _build_exact_index(rows: list, cols: list, excluded: set = None) -> dict:
    """Build an inverted exact-match index: value → [row indices]."""
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
    """Build an inverted index on a single FAC_CODE column."""
    index: dict = {}
    for i, row in enumerate(rows):
        if excluded and i in excluded:
            continue
        val = _normalize(row.get(col, ""))
        if val:
            index.setdefault(val, []).append(i)
    return index


def _build_token_index(rows: list, clean_cols: list, ori_col: str, excluded: set = None) -> tuple:
    """Build a token inverted index and a per-row token cache for fuzzy matching."""
    token_index: dict = {}
    token_cache: dict = {}
    for i, row in enumerate(rows):
        if excluded and i in excluded:
            continue
        texts = [_normalize(row.get(c, "")) for c in clean_cols]
        # texts.append(_normalize(row.get(ori_col, "")))  # DISABLED
        tokens: frozenset = frozenset()
        for t in filter(None, texts):
            tokens = tokens | _tokenize(t)
        token_cache[i] = tokens
        for tok in tokens:
            token_index.setdefault(tok, set()).add(i)
    return token_index, token_cache


def _build_lookup(rows: list, polis_cols: list, slip_cols: list,
                  insured_cols: list, facode_col: str = None,
                  label: str = "") -> tuple:
    """Build all lookup structures for one reference table (10-tuple)."""
    if not rows:
        if label:
            print(f"  {label:<6} done (0.0s)", flush=True)
        return ({}, {}, {}, {}, {}, {}, ({}, {}), ({}, {}), ({}, {}), {})

    t0 = time.perf_counter()
    excluded = {i for i, r in enumerate(rows) if _is_excluded_ref_row(r, polis_cols, slip_cols)}

    facode_index = (
        _build_facode_index(rows, facode_col, excluded=excluded)
        if facode_col else {}
    )
    lookup = (
        _build_exact_index(rows, slip_cols,    excluded=excluded),
        _build_exact_index(rows, polis_cols,   excluded=excluded),
        _build_exact_index(rows, insured_cols, excluded=excluded),
        {}, # DISABLED
        {}, # DISABLED
        {}, # DISABLED
        _build_token_index(rows, slip_cols,    "slip_ori",    excluded=excluded),
        _build_token_index(rows, polis_cols,   "polis_ori",   excluded=excluded),
        _build_token_index(rows, insured_cols, "insured_ori", excluded=excluded),
        facode_index,
    )

    elapsed = time.perf_counter() - t0
    if label:
        exc_info = f" - {len(excluded):,} baris dikecualikan dari matching (HUTANG PIUTANG / DATA SUSPENSE)" if excluded else ""
        print(f"  {label:<6} done ({elapsed:.1f}s){exc_info}", flush=True)

    return lookup


# =============================================================================
# MATCH ENGINE
# =============================================================================

def _exact_match(query_values: list, index: dict) -> set:
    """Return row indices that exact-match any query value."""
    return {i for v in query_values if v and v in index for i in index[v]}


def _like_match(
    query_values: list,
    token_index:  dict,
    rows:         list,
    ref_cols:     list,
    ori_col:      str = None,
) -> list:
    """Return row indices where any query value appears as a full substring in a reference column."""
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

    check_cols = list(ref_cols)
    # DISABLED: non-aktifkan cek ke ori_col

    matched = []
    for i in candidates:
        ref_values = [_normalize(rows[i].get(c, "")) for c in check_cols]
        for q in valid_queries:
            if any(q in rv for rv in ref_values if rv):
                matched.append(i)
                break
    return matched


def _match_polis_slip(
    suspend_row:          dict,
    lookup_slip_cln:      dict,
    lookup_polis_cln:     dict,
    lookup_slip_ori:      dict,
    lookup_polis_ori:     dict,
    polis_sus_cols:       list,
    slip_sus_cols:        list,
    rows:                 list = None,
    slip_ref_cols:        list = None,
    polis_ref_cols:       list = None,
    token_slip:           tuple = None,
    token_polis:          tuple = None,
    polis_sus_like_cols:  list = None,
    slip_sus_like_cols:   list = None,
) -> tuple:
    """Match one Suspend row against a reference table on Slip & Polis fields.

    Stage 1 (Exact): Polis dicek lebih dulu (sesuai urutan Polis → Slip → Insured).
    Stage 2 (LIKE):  contains/substring match.
    Returns (matched_indices, scenario_label).
    """
    def _sus_values(like_cols, ori_key):
        vals = _collect_clean_values(suspend_row, like_cols)
        return vals

    # Stage 1: Exact match — Polis dicek lebih dulu
    for values, index, label in [
        (_collect_clean_values(suspend_row, polis_sus_cols), lookup_polis_cln, "POLIS_CLEAN"),
        (_collect_clean_values(suspend_row, slip_sus_cols),  lookup_slip_cln,  "SLIP_CLEAN"),
    ]:
        hit = _exact_match(values, index)
        if hit:
            return list(hit), label

    # Stage 2: LIKE / contains match
    _slip_like  = slip_sus_like_cols  if slip_sus_like_cols  is not None else slip_sus_cols
    _polis_like = polis_sus_like_cols if polis_sus_like_cols is not None else polis_sus_cols

    if rows is not None and token_slip and token_polis:
        for values, idx_tok, ref_cols, ori_col, label in [
            (_sus_values(_polis_like, "polis_ori"), token_polis[0], polis_ref_cols or [], None, "POLIS_LIKE"),
            (_sus_values(_slip_like,  "slip_ori"),  token_slip[0],  slip_ref_cols  or [], None, "SLIP_LIKE"),
        ]:
            if values and idx_tok:
                hit = _like_match(values, idx_tok, rows, ref_cols, ori_col=None)
                if hit:
                    return hit, label

    return [], "Unmatching"


def _match_insured(
    suspend_row:       dict,
    lookup_insured_cln: dict,
    lookup_insured_ori: dict,
    insured_sus_cols:   list,
    rows:               list = None,
    insured_ref_cols:   list = None,
    token_insured:      tuple = None,
) -> tuple:
    """Match one Suspend row against a reference table on Insured field.

    Stage 1: Exact match. Stage 2: LIKE/contains match.
    Returns (matched_indices, scenario_label).
    """
    def _sus_values():
        vals = _collect_clean_values(suspend_row, insured_sus_cols)
        return vals

    for values, index, label in [
        (_collect_clean_values(suspend_row, insured_sus_cols), lookup_insured_cln, "INSURED_CLEAN"),
    ]:
        hit = _exact_match(values, index)
        if hit:
            return list(hit), label

    if rows is not None and token_insured:
        values = _sus_values()
        idx_tok = token_insured[0] if token_insured else None
        if values and idx_tok:
            hit = _like_match(values, idx_tok, rows, insured_ref_cols or [], ori_col=None)
            if hit:
                return hit, "INSURED_LIKE"

    return [], "Unmatching"


# =============================================================================
# NARROWING
# =============================================================================

def _split_composite(text: str) -> set:
    """Split a delimiter-separated string into individual tokens."""
    if not text:
        return set()
    return {tok.strip() for tok in _DELIMITER_RE.split(text) if tok.strip()}


def _ref_has_value(ref_row: dict, clean_cols: list, ori_col: str, query_values: list) -> bool:
    """Check if a reference row contains at least one query value.

    Pass (a): direct set intersection (exact match).
    Pass (b): split each reference value by delimiter, then exact-match each token.
    Pass (c): substring check on the full (unsplit) ref value (min length 10).
    """
    if not query_values:
        return False

    ref_clean = _collect_clean_values(ref_row, clean_cols)
    ref_vals  = set(ref_clean)

    if set(query_values) & ref_vals:
        return True

    ref_tokens: set = set()
    for rv in ref_vals:
        ref_tokens |= _split_composite(rv)

    if any(v in ref_tokens for v in query_values if v and len(v) >= 2):
        return True

    return any(
        v in rv
        for v in query_values if v and len(v) >= 10
        for rv in ref_vals if rv
    )


def _narrow(
    suspend_row:      dict,
    matched:          list,
    rows:             list,
    label:            str,
    polis_ref_cols:   list,
    slip_ref_cols:    list,
    insured_ref_cols: list,
    polis_sus_cols:   list,
    slip_sus_cols:    list,
    insured_sus_cols: list,
    sertif_sus_cols:  list = None,
    sertif_idx:       dict = None,
) -> tuple:
    """Narrow matched candidates using cascading field checks.

    Urutan narrowing (semua preferential):
      Step 0: Sertifikat (paling pertama, jika tersedia)
      Step 1: Field pelengkap (Slip jika initial=POLIS, Polis jika initial=SLIP)
      Step 2: Insured (jika masih >1 setelah step 0+1)
    """
    base_scenario = _map_scenario(label)
    if not matched:
        return [], "Unmatching"
    if len(matched) <= 1:
        return matched, base_scenario

    polis_values   = _collect_all_values(suspend_row, polis_sus_cols,   SUS_POLIS_ORI_COL)
    slip_values    = _collect_all_values(suspend_row, slip_sus_cols,    SUS_SLIP_ORI_COL)
    insured_values = _collect_all_values(suspend_row, insured_sus_cols, SUS_INSURED_ORI_COL)

    current_matched = matched

    got_sertif  = False
    got_compl   = False
    got_insured = False

    # ---------- Step 0: Sertifikat ----------
    if len(current_matched) > 1 and sertif_sus_cols and sertif_idx:
        sertif_vals = _collect_clean_values(suspend_row, sertif_sus_cols)
        expanded_sertif_vals = []
        for v in sertif_vals:
            expanded_sertif_vals.extend(_expand_sertif_range(v))
        expanded_sertif_vals = list(set(expanded_sertif_vals))

        if expanded_sertif_vals:
            confirmed = [
                i for i in current_matched
                if any(sv in sertif_idx and i in sertif_idx[sv] for sv in expanded_sertif_vals)
            ]
            if confirmed:
                current_matched = confirmed
                got_sertif = True

    # ---------- Step 1: Field pelengkap ----------
    if len(current_matched) > 1:
        if "SLIP" in label:
            confirmed = [i for i in current_matched
                         if _ref_has_value(rows[i], polis_ref_cols, SUS_POLIS_ORI_COL, polis_values)]
            if confirmed:
                current_matched = confirmed
                got_compl = True
        elif "POLIS" in label:
            confirmed = [i for i in current_matched
                         if _ref_has_value(rows[i], slip_ref_cols, SUS_SLIP_ORI_COL, slip_values)]
            if confirmed:
                current_matched = confirmed
                got_compl = True

    # ---------- Step 2: Insured ----------
    if len(current_matched) > 1 and ("POLIS" in label or "SLIP" in label):
        confirmed = [i for i in current_matched
                     if _ref_has_value(rows[i], insured_ref_cols, SUS_INSURED_ORI_COL, insured_values)]
        if confirmed:
            current_matched = confirmed
            got_insured = True

    # ---------- Bangun label secara eksplisit ----------
    if "POLIS" in label:
        parts = ["Polis"]
        if got_sertif:  parts.append("Sertif")
        if got_compl:   parts.append("Slip")
        if got_insured: parts.append("Insured")
        scenario = " + ".join(parts) if len(parts) > 1 else "Polis only"
    elif "SLIP" in label:
        parts = ["Slip"]
        if got_sertif:  parts.append("Sertif")
        if got_compl:   parts.append("Polis")
        if got_insured: parts.append("Insured")
        scenario = " + ".join(parts) if len(parts) > 1 else "Slip only"
    else:
        scenario = base_scenario

    return current_matched, scenario


def _narrow_by_currency(
    suspend_row: dict,
    matched:     list,
    rows:        list,
) -> list:
    """Preferential currency narrowing: filter OSBAL candidates by CCOS_CURR == CURR ORI."""
    if len(matched) <= 1:
        return matched

    sus_curr = _normalize(suspend_row.get(SUSPEND_CURR_COL, ""))
    if not sus_curr:
        return matched

    filtered = [
        i for i in matched
        if _normalize(rows[i].get(OSBAL_CURR_COL, "")) == sus_curr
    ]
    return filtered if filtered else matched


def _narrow_by_periode(
    suspend_row: dict,
    matched:     list,
    rows:        list,
) -> list:
    """Preferential periode narrowing: filter OSBAL candidates where RECEIPT DATE < FAC_COM_DATE.

    NON-LINESLIP: RECEIPT DATE >= FAC_COM_DATE.
    LINESLIP    : FAC_COM_DATE harus TEPAT = RECEIPT DATE - 1 bulan (exact month+year).
    """
    if len(matched) <= 1:
        return matched

    sus_date_normal = suspend_row.get("_sus_date_parsed", "__MISSING__")
    sus_date_ls     = suspend_row.get("_sus_date_lineslip", None)

    if sus_date_normal == "__MISSING__":
        sus_date_raw = suspend_row.get(SUSPEND_DATE_COL, "")
        if sus_date_raw is None or (isinstance(sus_date_raw, float) and pd.isna(sus_date_raw)):
            return matched
        try:
            sus_date_normal = pd.to_datetime(sus_date_raw, errors="raise")
            sus_date_ls     = sus_date_normal - pd.DateOffset(months=1)
        except Exception:
            return matched

    if sus_date_normal is None:
        return matched

    sus_is_ls = suspend_row.get("_is_lineslip", False)

    valid = []
    for i in matched:
        com_date = rows[i].get("_com_date_parsed", "__MISSING__")
        if com_date == "__MISSING__":
            com_raw = rows[i].get(OSBAL_DATE_COL, "")
            if com_raw is None or (isinstance(com_raw, float) and pd.isna(com_raw)):
                valid.append(i)
                continue
            try:
                com_date = pd.to_datetime(com_raw, errors="raise")
            except Exception:
                valid.append(i)
                continue

        if com_date is None:
            valid.append(i)
            continue

        try:
            is_ls = sus_is_ls or rows[i].get("_is_lineslip", False)
            if is_ls and sus_date_ls is not None:
                if com_date.year == sus_date_ls.year and com_date.month == sus_date_ls.month:
                    valid.append(i)
            else:
                if sus_date_normal >= com_date:
                    valid.append(i)
        except Exception:
            valid.append(i)

    return valid if valid else matched


def _rematch_by_currency(
    suspend_row:          dict,
    osbal_rows:           list,
    lookup_osbal:         tuple,
    polis_sus_cols:       list,
    slip_sus_cols:        list,
    insured_sus_cols:     list,
    polis_ref_cols:       list,
    slip_ref_cols:        list,
    insured_ref_cols:     list,
    facode_osbal_idx:     dict,
    facul_rows:           list = None,
    lookup_facul:         tuple = None,
    polis_facul:          list = None,
    slip_facul:           list = None,
    insured_facul:        list = None,
    polis_sus_like_cols:  list = None,
    slip_sus_like_cols:   list = None,
) -> tuple:
    """Rematch a Suspend row flagged 'Beda Currency' searching for currency-matching candidates."""
    sus_curr = _normalize(suspend_row.get(SUSPEND_CURR_COL, ""))
    if not sus_curr:
        return None

    def _filter_osbal_curr(osbal_indices):
        return [i for i in osbal_indices if _normalize(osbal_rows[i].get(OSBAL_CURR_COL, "")) == sus_curr]

    # --- [Pass R1] OSBAL Polis Match with Currency Filter ---
    lkp_p_osbal = lookup_osbal[1]
    polis_vals = _collect_clean_values(suspend_row, polis_sus_cols)
    hits_p = _exact_match(polis_vals, lkp_p_osbal)
    curr_hits_p = _filter_osbal_curr(hits_p)

    if curr_hits_p:
        slip_vals = _collect_clean_values(suspend_row, slip_sus_cols)
        confirmed = [i for i in curr_hits_p if any(
            _normalize(osbal_rows[i].get(c, "")) in slip_vals
            for c in slip_ref_cols if _normalize(osbal_rows[i].get(c, ""))
        )]
        chosen = confirmed if confirmed else curr_hits_p
        scenario = "Polis + Slip" if confirmed else "Polis only"
        return "OSBAL", scenario, chosen, True

    # --- [Pass R2] OSBAL Polis LIKE Match with Currency Filter ---
    like_p_cols = polis_sus_like_cols if polis_sus_like_cols is not None else polis_sus_cols
    if lookup_osbal[7] and lookup_osbal[7][0] and like_p_cols:
        like_vals = _collect_clean_values(suspend_row, like_p_cols)
        if like_vals:
            hits_like = _like_match(like_vals, lookup_osbal[7][0], osbal_rows, polis_ref_cols)
            curr_hits_like = _filter_osbal_curr(hits_like)
            if curr_hits_like:
                return "OSBAL", "POLIS_LIKE", curr_hits_like, True

    # --- [Pass R3] FACUL Polis Match → OSBAL resolution with Currency Filter ---
    if facul_rows and lookup_facul:
        lkp_p_facul = lookup_facul[1]
        hits_p_fac = _exact_match(polis_vals, lkp_p_facul)
        if hits_p_fac:
            matched_fac = list(hits_p_fac)
            res_idx, ok = _resolve_facode([facul_rows[i] for i in matched_fac], FACUL_FACODE_COL, facode_osbal_idx, osbal_rows)
            if ok:
                curr_res = _filter_osbal_curr(res_idx)
                if curr_res:
                    return "FACUL", "Polis only", curr_res, True

    # --- [Pass R4] OSBAL Insured Fallback with Currency Filter ---
    lkp_ins_osbal = lookup_osbal[2]
    ins_vals = _collect_clean_values(suspend_row, insured_sus_cols)
    hits_ins = _exact_match(ins_vals, lkp_ins_osbal)
    curr_hits_ins = _filter_osbal_curr(hits_ins)
    if curr_hits_ins:
        return "OSBAL", "Insured only", curr_hits_ins, True

    return None


def _run_polis_slip_pass(
    suspend_row:          dict,
    rows:                 list,
    lookup:               tuple,
    polis_ref_cols:       list,
    slip_ref_cols:        list,
    insured_ref_cols:     list,
    polis_sus_cols:       list,
    slip_sus_cols:        list,
    insured_sus_cols:     list,
    polis_sus_like_cols:  list = None,
    slip_sus_like_cols:   list = None,
    facode_col:           str  = None,
    gunakan_narrow_aca:   bool = False,
    currency_narrowing:   bool = False,
    sertif_sus_cols:      list = None,
    sertif_idx:           dict = None,
) -> tuple:
    """Run one Polis/Slip matching pass + narrowing."""
    (
        lkp_slip_cln, lkp_polis_cln, lkp_insured_cln,
        lkp_slip_ori, lkp_polis_ori, lkp_insured_ori,
        tok_slip, tok_polis, tok_insured, _facode,
    ) = lookup

    matched, label = _match_polis_slip(
        suspend_row,
        lkp_slip_cln, lkp_polis_cln,
        lkp_slip_ori, lkp_polis_ori,
        polis_sus_cols, slip_sus_cols,
        rows=rows,
        slip_ref_cols=slip_ref_cols,
        polis_ref_cols=polis_ref_cols,
        token_slip=tok_slip,
        token_polis=tok_polis,
        polis_sus_like_cols=polis_sus_like_cols,
        slip_sus_like_cols=slip_sus_like_cols,
    )

    if matched:
        if currency_narrowing and len(matched) > 1:
            matched = _narrow_by_currency(suspend_row, matched, rows)
            if len(matched) > 1:
                matched = _narrow_by_periode(suspend_row, matched, rows)
        narrowed, scenario = _narrow(
            suspend_row, matched, rows, label,
            polis_ref_cols, slip_ref_cols, insured_ref_cols,
            polis_sus_cols, slip_sus_cols, insured_sus_cols,
            sertif_sus_cols=sertif_sus_cols,
            sertif_idx=sertif_idx,
        )
        return narrowed, scenario
    return [], "Unmatching"


def _run_insured_pass(
    suspend_row:        dict,
    rows:               list,
    lookup:             tuple,
    insured_sus_cols:   list,
    insured_ref_cols:   list = None,
    currency_narrowing: bool = False,
    **_kwargs,
) -> tuple:
    """Run one Insured fallback matching pass."""
    (
        _s_cln, _p_cln, lkp_ins_cln,
        _s_ori, _p_ori, lkp_ins_ori,
        _tok_s, _tok_p, tok_ins, _fac,
    ) = lookup

    matched, label = _match_insured(
        suspend_row,
        lkp_ins_cln, lkp_ins_ori,
        insured_sus_cols,
        rows=rows,
        insured_ref_cols=insured_ref_cols,
        token_insured=tok_ins,
    )

    if matched:
        if currency_narrowing and len(matched) > 1:
            matched = _narrow_by_currency(suspend_row, matched, rows)
            if len(matched) > 1:
                matched = _narrow_by_periode(suspend_row, matched, rows)
        return matched, _map_scenario(label)
    return [], "Unmatching"


# =============================================================================
# FAC CODE RESOLUTION
# =============================================================================

def _resolve_facode(
    ref_rows_matched: list,
    facode_col:       str,
    facode_osbal_idx: dict,
    osbal_rows:       list,
) -> tuple:
    """Resolve FAC_CODE from FACUL matches to their OSBAL row indices."""
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
# SERTIF INDEX
# =============================================================================

def _build_sertif_index(rows: list, sertif_cols: list, excluded: set = None) -> dict:
    """Build inverted index: sertif_value -> [row indices]."""
    index: dict = {}
    for i, row in enumerate(rows):
        if excluded and i in excluded:
            continue
        for col in sertif_cols:
            val = _normalize(row.get(col, ""))
            expanded = _expand_sertif_range(val)
            for cert in expanded:
                index.setdefault(cert, []).append(i)
    return index


# =============================================================================
# OUTPUT BUILDERS
# =============================================================================

def _aggregate_ccos(osbal_rows: list) -> dict:
    """Aggregate financial values and join text fields from matched OSBAL rows."""
    fac_codes = {str(r.get("CCOS_REF_CODE", "")).strip() for r in osbal_rows
                 if str(r.get("CCOS_REF_CODE", "")).strip()}
    multi_fac = len(fac_codes) > 1

    def _join_unique(key):
        seen, vals = set(), []
        for r in osbal_rows:
            v = str(r.get(key, "")).strip()
            if v and v not in seen:
                seen.add(v)
                vals.append(v)
        return ", ".join(vals)

    def _sum_or_nan(key):
        if multi_fac:
            return np.nan
        total = 0.0
        for r in osbal_rows:
            v = r.get(key, 0)
            if pd.isna(v): continue
            if isinstance(v, (int, float)):
                total += float(v)
            else:
                v_str = str(v).strip()
                if not v_str: continue
                if ',' in v_str and '.' in v_str:
                    if v_str.rfind(',') > v_str.rfind('.'):
                        v_str = v_str.replace('.', '').replace(',', '.')
                    else:
                        v_str = v_str.replace(',', '')
                elif ',' in v_str:
                    v_str = v_str.replace(',', '.')
                try:
                    total += float(v_str)
                except ValueError:
                    pass
        return total

    return {
        "CCOS_DOC_NO":   _join_unique("CCOS_DOC_NO"),
        "CCOS_REF_CODE": _join_unique("CCOS_REF_CODE"),
        "CCOS_OR_BAL":   _sum_or_nan("CCOS_OR_BAL"),
        "CCOS_BAL_DUE":  _sum_or_nan("CCOS_BAL_DUE"),
    }


def _facode_label(osbal_rows: list, source: str) -> str:
    """Return the single FAC_CODE or 'facode lebih dari 1' if multiple."""
    if not osbal_rows or not osbal_rows[0]:
        return ""
    col = OSBAL_FACODE_COL if source == "OSBAL" else FACUL_FACODE_COL
    codes = {str(r.get(col, "")).strip() for r in osbal_rows if str(r.get(col, "")).strip()}
    return "facode lebih dari 1" if len(codes) > 1 else next(iter(codes), "")


def _format_sertif(val):
    if pd.isna(val) or str(val).strip() == "":
        return ""
    if isinstance(val, (int, float)):
        return str(int(val)).zfill(6)
    s = str(val).strip()
    if s.endswith('.0'):
        s = s[:-2]
    if s.isdigit():
        return s.zfill(6)
    return s


def _build_output_row(
    suspend_row:    dict,
    source:         str,
    scenario:       str,
    osbal_rows:     list,
    osbal_count:    int  = 0,
    resolved:       bool = True,
    suspend_count:  int  = 1,
) -> dict:
    """Build one complete output row dict."""
    has_match   = bool(source and osbal_rows and osbal_rows[0])
    has_osbal   = has_match and (source == "OSBAL" or resolved)

    ccos = (
        _aggregate_ccos(osbal_rows)
        if has_osbal
        else {"CCOS_DOC_NO": "", "CCOS_REF_CODE": "", "CCOS_OR_BAL": np.nan, "CCOS_BAL_DUE": np.nan}
    )

    fac_label = (
        _facode_label(osbal_rows, "OSBAL") if has_osbal else
        _facode_label(osbal_rows, source)  if has_match else ""
    )

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

        # Kolom insured dari suspend Astra Buana
        "INSURED_ORI": suspend_row.get(SUS_INSURED_ORI_COL,  ""),   # "INSURED"
        "INSURED_1":   suspend_row.get("INSURED_CLEAN",       ""),
        "INSURED_2":   "",  # hanya 1 clean insured di data Astra Buana

        "CURR ORI":   suspend_row.get("CURR ORI",   ""),
        "AMOUNT ORI": suspend_row.get("AMOUNT ORI", ""),
        "CURR PAY":   suspend_row.get("CURR PAY",   ""),
        "AMOUNT PAY": suspend_row.get("AMOUNT PAY", ""),

        "CCOS_OR_BAL":  ccos["CCOS_OR_BAL"],
        "CCOS_BAL_DUE": ccos["CCOS_BAL_DUE"],

        # Kolom polis/sertif/slip dari suspend Astra Buana
        "POLIS":       suspend_row.get("POLIS") or suspend_row.get(SUS_POLIS_ORI_COL, ""),
        "POLIS_CLN":   suspend_row.get("POLIS_CLN") or suspend_row.get("POLIS_CLEAN", ""),
        "SERTIF_CLN_1": _format_sertif(suspend_row.get("SERTIF_CLN_1") or suspend_row.get("SERTIF_CLN") or suspend_row.get("CERTIFICATE_1", "")),
        "SERTIF_CLN_2": _format_sertif(suspend_row.get("SERTIF_CLN_2") or suspend_row.get("CERTIFICATE_2", "")),
        "SERTIF_CLN_3": _format_sertif(suspend_row.get("SERTIF_CLN_3") or suspend_row.get("CERTIFICATE_3", "")),
        "SLIP_NO":     suspend_row.get("SLIP_NO") or suspend_row.get(SUS_SLIP_ORI_COL, ""),
        "SLIP_NO_CLN": suspend_row.get("SLIP_NO_CLN") or suspend_row.get("SLIP_NO_CLEAN", ""),

        "DESC 1": suspend_row.get("DESC 1", ""),
        "DESC 2": suspend_row.get("DESC 2", ""),
        "DESC 3": suspend_row.get("DESC 3", ""),
        "DESC 4": suspend_row.get("DESC 4", ""),

        "STATUS":   suspend_row.get("STATUS",   ""),
        "REC_TYPE": suspend_row.get("REC_TYPE", ""),

        "CEK AMOUNT DATABASE X BAL RV": "",
        "Mark Admin Fac Code":          fac_label,
        "Mark Admin":                   "",
        "Mark Admin (Status)":          "",
        "Mark ARP":                     "",

        "SKENARIO":           final_scenario,
        "_OSBAL_ROW_COUNT":   osbal_count,
        "_SUSPEND_ROW_COUNT": suspend_count,
    }


def _merge_suspend_rows(rows: list) -> dict:
    """Merge multiple Suspend dicts (same fac code) into one: sum AMOUNT ORI, join-unique others."""
    if len(rows) == 1:
        return rows[0]

    merged: dict = {}
    for col in _SUSPEND_SUM_COLS:
        merged[col] = sum(pd.to_numeric(r.get(col, 0), errors="coerce") or 0 for r in rows)
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
    """Compute AMOUNT_ORI_MIN1, DIFERENCE, and FLAG_PROD."""
    if "_MERGED_AMOUNT_ORI" in df.columns:
        amount_ori_raw = df["_MERGED_AMOUNT_ORI"]
    else:
        amount_ori_raw = df.get("AMOUNT ORI", pd.Series([0]*len(df), index=df.index))

    amount_ori_clean = amount_ori_raw.replace(r'^\s*$', np.nan, regex=True)
    if amount_ori_clean.dtype == object:
        amount_ori_clean = amount_ori_clean.str.replace(',', '.', regex=False)
    amount_ori_numeric = pd.to_numeric(amount_ori_clean, errors="coerce")

    df["AMOUNT_ORI_MIN1"] = np.where(amount_ori_numeric.notna(), amount_ori_numeric * -1, np.nan)

    amount_ori = amount_ori_numeric.fillna(0)
    amount_neg = df["AMOUNT_ORI_MIN1"].fillna(0)

    balance_due  = pd.to_numeric(df.get("CCOS_BAL_DUE",      0), errors="coerce").fillna(0)
    osbal_count  = pd.to_numeric(df.get("_OSBAL_ROW_COUNT",  0), errors="coerce").fillna(0)
    sus_count    = pd.to_numeric(df.get("_SUSPEND_ROW_COUNT", 1), errors="coerce").fillna(1)
    fac_series   = df.get("Mark Admin Fac Code", pd.Series("", index=df.index)).fillna("")
    skenario_col = df.get("SKENARIO", pd.Series("", index=df.index)).fillna("")

    df["DIFERENCE"] = amount_neg - balance_due

    accumulated = (osbal_count > 1) | (sus_count > 1)
    is_equal = (amount_neg - balance_due).abs() <= 0.01 * amount_neg.abs()

    is_new_entry_total = (amount_ori != 0) & (balance_due == 0)
    is_matching_gt1_fac = (fac_series.str.strip() == "facode lebih dari 1")

    is_adj_total    = is_equal
    is_adj_sebagian = ~is_equal & (amount_neg < balance_due)
    is_new_entry_sebagian = ~is_equal & (amount_neg > balance_due)

    df["FLAG_PROD"] = np.select(
        [
            skenario_col == "Beda Currency",
            skenario_col == "Beda Periode",
            is_matching_gt1_fac,
            is_new_entry_total,
            is_adj_total & ~accumulated,
            is_adj_total & accumulated,
            is_adj_sebagian & ~accumulated,
            is_adj_sebagian & accumulated,
            is_new_entry_sebagian,
        ],
        [
            "Beda Currency",
            "Beda Periode",
            "Matching >1 fac code",
            "New Entry total",
            "Adjustment total tanpa akumulasi",
            "Adjustment total dengan akumulasi",
            "Adjustment sebagian tanpa akumulasi",
            "Adjustment sebagian dengan akumulasi",
            "New Entry sebagian",
        ],
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
    """Load an Excel file to (list[dict], list[str]) with automatic pickle cache."""
    if not os.path.exists(path):
        print(f"  {label}: {path} ({'not found, skipped' if optional else 'NOT FOUND'})", flush=True)
        return [], []

    cache = path + ".cache.pkl"
    if os.path.exists(cache) and os.path.getmtime(cache) >= os.path.getmtime(path):
        t = time.perf_counter()
        print(f"  {label}: loading from cache ...", flush=True)
        try:
            data, header = pd.read_pickle(cache)
            print(f"  -> {len(data):,} rows, {len(header)} cols ({time.perf_counter()-t:.2f}s)", flush=True)
            return data, header
        except Exception:
            print(f"  {label}: cache corrupt, fallback to Excel read ...", flush=True)

    t = time.perf_counter()
    print(f"  {label}: {path} ...", flush=True)
    try:
        df = pd.read_excel(path)
    except Exception as e:
        print(f"  [RETRY] Default engine failed ({e}), loading via engine='openpyxl' ...", flush=True)
        df = pd.read_excel(path, engine="openpyxl")

    header = list(df.columns)
    data   = df.to_dict("records")
    print(f"  -> {len(data):,} rows, {len(header)} cols ({time.perf_counter()-t:.1f}s)", flush=True)

    try:
        pd.to_pickle((data, header), cache)
        print(f"  -> cache saved: {cache}", flush=True)
    except Exception:
        pass

    return data, header


def run() -> None:
    """Main pipeline: match SUSPEND rows against OSBAL / FACUL and export results.

    Cedant : PT ASURANSI ASTRA BUANA
    Data   : astrabuana_output_suspend / osbal / facul.xlsx
    Bordero: TIDAK DIGUNAKAN
    """
    t_start = time.perf_counter()

    print("\n" + "=" * 60, flush=True)
    print("  PRODUCTION SCRIPT — SUSPEND MATCHING", flush=True)
    print("  Cedant : PT ASURANSI ASTRA BUANA", flush=True)
    print("=" * 60, flush=True)

    # [1] Load data
    print("\n[1/5] Loading data ...", flush=True)
    suspend_rows, suspend_cols = _load_excel(SUSPEND_FILE, "SUSPEND")
    osbal_rows,   osbal_cols   = _load_excel(OSBAL_FILE,   "OSBAL")
    facul_rows,   facul_cols   = _load_excel(FACUL_FILE,   "FACUL")

    if not suspend_rows:
        print("  [ERROR] Tidak ada data suspend. Hentikan.", flush=True)
        return
    if not osbal_rows:
        print("  [ERROR] Tidak ada data osbal. Hentikan.", flush=True)
        return

    # Pre-parse tanggal OSBAL (FAC_COM_DATE)
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

    # Pre-parse tanggal Suspend (RECEIPT DATE)
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

    # [2] Detect clean columns (per tabel dengan prefix spesifik)
    print("\n[2/5] Detecting clean columns ...", flush=True)

    polis_sus   = _get_clean_cols_for(suspend_cols, "suspend", "polis")
    slip_sus    = _get_clean_cols_for(suspend_cols, "suspend", "slip")
    insured_sus = _get_clean_cols_for(suspend_cols, "suspend", "insured")
    sertif_sus  = _get_clean_cols_for(suspend_cols, "suspend", "sertif")

    polis_osbal   = _get_clean_cols_for(osbal_cols, "osbal", "polis")
    slip_osbal    = _get_clean_cols_for(osbal_cols, "osbal", "slip")
    insured_osbal = _get_clean_cols_for(osbal_cols, "osbal", "insured")
    sertif_osbal  = _get_clean_cols_for(osbal_cols, "osbal", "sertif")

    polis_facul   = _get_clean_cols_for(facul_cols, "facul", "polis")
    slip_facul    = _get_clean_cols_for(facul_cols, "facul", "slip")
    insured_facul = _get_clean_cols_for(facul_cols, "facul", "insured")

    print(f"  Suspend  polis  : {polis_sus}", flush=True)
    print(f"  Suspend  slip   : {slip_sus}", flush=True)
    print(f"  Suspend  insured: {insured_sus}", flush=True)
    print(f"  Suspend  sertif : {sertif_sus}", flush=True)
    print(f"  OSBAL    polis  : {polis_osbal}", flush=True)
    print(f"  OSBAL    slip   : {slip_osbal}", flush=True)
    print(f"  OSBAL    insured: {insured_osbal}", flush=True)
    print(f"  OSBAL    sertif : {sertif_osbal}", flush=True)
    print(f"  FACUL    polis  : {polis_facul}", flush=True)
    print(f"  FACUL    slip   : {slip_facul}", flush=True)
    print(f"  FACUL    insured: {insured_facul}", flush=True)

    # Tambahkan CLSDT_POLICY_NO / CLSDT_SLIP_NO / CLSDT_SERTF_NO ke OSBAL jika ada
    def _prepend_if_exists(col: str, cols_list: list, header: list) -> list:
        return ([col] + cols_list) if col in header and col not in cols_list else cols_list

    polis_osbal = _prepend_if_exists("CLSDT_POLICY_NO", polis_osbal, osbal_cols)
    slip_osbal  = _prepend_if_exists("CLSDT_SLIP_NO",   slip_osbal,  osbal_cols)
    if "CLSDT_SERTF_NO" in osbal_cols and "CLSDT_SERTF_NO" not in sertif_osbal:
        sertif_osbal.insert(0, "CLSDT_SERTF_NO")

    polis_facul = _prepend_if_exists("CLSDT_POLICY_NO", polis_facul, facul_cols)
    slip_facul  = _prepend_if_exists("CLSDT_SLIP_NO",   slip_facul,  facul_cols)

    # Pre-parse LINESLIP flags untuk OSBAL
    print("  Pre-parsing OSBAL LINESLIP flags ...", flush=True)
    osbal_ls_cols = polis_osbal + slip_osbal + insured_osbal + ["FAC_POLICY_NO", "FAC_SLIP", "FAC_INSURED"]
    for _row in osbal_rows:
        _row["_is_lineslip"] = _is_lineslip_row(_row, osbal_ls_cols)

    # Pre-parse LINESLIP flags Suspend
    print("  Pre-parsing Suspend LINESLIP flags ...", flush=True)
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
    print(f"  -> {lineslip_count:,} baris Suspend LINESLIP terdeteksi", flush=True)

    # [3] Build lookup indexes
    print("\n[3/5] Building lookup indexes ...", flush=True)
    lookup_osbal = _build_lookup(osbal_rows, polis_osbal, slip_osbal, insured_osbal,
                                 OSBAL_FACODE_COL, label="OSBAL")
    lookup_facul = _build_lookup(facul_rows, polis_facul, slip_facul, insured_facul,
                                 FACUL_FACODE_COL, label="FACUL")

    facode_osbal_idx = lookup_osbal[9]

    # Build sertif index untuk OSBAL
    sertif_idx = _build_sertif_index(osbal_rows, sertif_osbal) if sertif_osbal else {}
    if sertif_idx:
        print(f"  Sertif index: {len(sertif_idx):,} entri", flush=True)

    # Shared kwargs
    kw_osbal = dict(
        rows=osbal_rows, lookup=lookup_osbal,
        polis_ref_cols=polis_osbal, slip_ref_cols=slip_osbal, insured_ref_cols=insured_osbal,
        facode_col=OSBAL_FACODE_COL,
        currency_narrowing=True,
    )
    kw_facul = dict(
        rows=facul_rows, lookup=lookup_facul,
        polis_ref_cols=polis_facul, slip_ref_cols=slip_facul, insured_ref_cols=insured_facul,
        facode_col=FACUL_FACODE_COL, gunakan_narrow_aca=True,
    )

    # [4] Per-row matching
    total = len(suspend_rows)
    print(f"\n[4/5] Matching {total:,} suspend rows ...", flush=True)
    t4 = time.perf_counter()

    raw_results = []

    for n, sus in enumerate(suspend_rows, 1):
        if n % 500 == 0:
            elapsed = time.perf_counter() - t4
            rate    = n / elapsed
            eta     = (total - n) / rate if rate > 0 else 0
            print(f"  {n:,} / {total:,}  ({rate:.0f} rows/s, ETA {eta:.0f}s)")

        source = scenario = None
        matched: list  = []
        osbal_ref: list = []
        osbal_count     = 0
        resolved        = True

        eff_polis, eff_slip, like_polis, like_slip = _get_effective_sus_cols(
            sus, polis_sus, slip_sus
        )

        kw_sus_row = dict(
            polis_sus_cols=eff_polis,
            slip_sus_cols=eff_slip,
            insured_sus_cols=insured_sus,
            polis_sus_like_cols=like_polis,
            slip_sus_like_cols=like_slip,
        )

        # Pass 1: SUSPEND → OSBAL (Polis/Slip) + sertif narrowing
        kw_osbal_sertif = dict(**kw_osbal, sertif_sus_cols=sertif_sus, sertif_idx=sertif_idx) \
            if sertif_sus and sertif_idx else kw_osbal
        m, lbl = _run_polis_slip_pass(sus, **kw_osbal_sertif, **kw_sus_row)
        if m:
            source = "OSBAL"; scenario = lbl
            matched = m; osbal_ref = [osbal_rows[i] for i in m]; osbal_count = len(osbal_ref)

        # Pass 2: SUSPEND → FACUL → OSBAL resolution
        if not matched and facul_rows:
            m, lbl = _run_polis_slip_pass(sus, **kw_facul, **kw_sus_row)
            if m:
                res_idx, ok = _resolve_facode([facul_rows[i] for i in m], FACUL_FACODE_COL, facode_osbal_idx, osbal_rows)
                if ok:
                    if len(res_idx) > 1:
                        res_idx = _narrow_by_periode(sus, res_idx, osbal_rows)
                    source = "FACUL"; scenario = lbl; resolved = True
                    matched = res_idx; osbal_ref = [osbal_rows[i] for i in res_idx]; osbal_count = len(osbal_ref)

        # Pass 3: Insured fallback → OSBAL
        if not matched:
            m, lbl = _run_insured_pass(sus, **kw_osbal, **kw_sus_row)
            if m:
                source = "OSBAL"; scenario = lbl
                matched = m; osbal_ref = [osbal_rows[i] for i in m]; osbal_count = len(osbal_ref)

        # Pass 4: Insured fallback → FACUL → OSBAL
        if not matched and facul_rows:
            m, lbl = _run_insured_pass(sus, **kw_facul, **kw_sus_row)
            if m:
                res_idx, ok = _resolve_facode([facul_rows[i] for i in m], FACUL_FACODE_COL, facode_osbal_idx, osbal_rows)
                if ok:
                    source = "FACUL"; scenario = lbl; resolved = True
                    matched = res_idx; osbal_ref = [osbal_rows[i] for i in res_idx]; osbal_count = len(osbal_ref)

        # Pass 5: Unmatching fallback
        if not matched:
            source = None; scenario = "Unmatching"
            osbal_ref = [{}]; osbal_count = 0; resolved = False

        # Currency check (ATURAN 1)
        if source is not None and osbal_ref and osbal_ref[0]:
            sus_curr = _normalize(sus.get(SUSPEND_CURR_COL, ""))
            osbal_curr_values = {
                _normalize(r.get(OSBAL_CURR_COL, ""))
                for r in osbal_ref if r
            }
            osbal_curr_values.discard("")
            if sus_curr and osbal_curr_values and sus_curr not in osbal_curr_values:
                rematch_res = _rematch_by_currency(
                    sus, osbal_rows, lookup_osbal,
                    kw_sus_row["polis_sus_cols"], kw_sus_row["slip_sus_cols"], kw_sus_row["insured_sus_cols"],
                    polis_osbal, slip_osbal, insured_osbal,
                    facode_osbal_idx,
                    facul_rows=facul_rows, lookup_facul=lookup_facul,
                    polis_facul=polis_facul, slip_facul=slip_facul, insured_facul=insured_facul,
                    polis_sus_like_cols=kw_sus_row["polis_sus_like_cols"],
                    slip_sus_like_cols=kw_sus_row["slip_sus_like_cols"],
                )
                if rematch_res:
                    source, scenario, matched, resolved = rematch_res
                    osbal_ref = [osbal_rows[i] for i in matched]
                    osbal_count = len(osbal_ref)
                else:
                    scenario = "Beda Currency"

        # Periode check (ATURAN BEDA PERIODE)
        if (source is not None and osbal_ref and osbal_ref[0]
                and scenario != "Beda Currency"):

            all_beda = True
            sus_is_ls = sus.get("_is_lineslip", False)
            sus_date_normal = sus.get("_sus_date_parsed")
            sus_date_ls = sus.get("_sus_date_lineslip")

            for r in osbal_ref:
                if not r:
                    all_beda = False
                    break

                is_ls = sus_is_ls or r.get("_is_lineslip", False)
                com_date = r.get("_com_date_parsed", "__MISSING__")
                if com_date == "__MISSING__":
                    com_raw = r.get(OSBAL_DATE_COL, "")
                    if com_raw is None or (isinstance(com_raw, float) and pd.isna(com_raw)):
                        all_beda = False
                        break
                    try:
                        com_date = pd.to_datetime(com_raw, errors="raise")
                    except Exception:
                        all_beda = False
                        break

                if com_date is None:
                    all_beda = False
                    break

                try:
                    if is_ls and sus_date_ls is not None:
                        if com_date.year == sus_date_ls.year and com_date.month == sus_date_ls.month:
                            all_beda = False
                            break
                    elif not is_ls and sus_date_normal is not None:
                        if sus_date_normal >= com_date:
                            all_beda = False
                            break
                except Exception:
                    all_beda = False
                    break

            if all_beda:
                scenario = "Beda Periode"

        fac_codes = {_normalize(r.get(OSBAL_FACODE_COL, "")) for r in osbal_ref if r}
        fac_codes.discard("")

        raw_results.append({
            "suspend":    sus,
            "source":     source,
            "scenario":   scenario,
            "osbal_idx":  list(matched) if source else [],
            "osbal_count": osbal_count,
            "resolved":   resolved,
            "fac_codes":  fac_codes,
        })

    print(f"  Matching done in {time.perf_counter()-t4:.1f}s", flush=True)

    # [4b] Fac code accumulation (post-matching)
    print("\n[4b/5] Fac code accumulation ...", flush=True)

    _EXCLUDED_SCENARIOS = {"Beda Currency", "Beda Periode"}

    claimed: set = set()
    for r in raw_results:
        if len(r["fac_codes"]) == 1 and r["scenario"] not in _EXCLUDED_SCENARIOS:
            claimed |= r["fac_codes"]

    for r in raw_results:
        if len(r["fac_codes"]) > 1:
            remaining = r["fac_codes"] - claimed
            if len(remaining) == 1 and r["scenario"] not in _EXCLUDED_SCENARIOS:
                r["fac_codes"] = remaining
                claimed |= remaining
                fac = next(iter(remaining))
                r["osbal_idx"] = [i for i in r["osbal_idx"]
                                  if _normalize(osbal_rows[i].get(OSBAL_FACODE_COL, "")) == fac]
                r["osbal_count"] = len(r["osbal_idx"])

    groups: dict  = {}
    final:  list  = []

    for r in raw_results:
        r.setdefault("suspend_count", 1)
        if (len(r["fac_codes"]) == 1
                and r["source"] is not None
                and r["scenario"] not in _EXCLUDED_SCENARIOS):
            fac = next(iter(r["fac_codes"]))
            groups.setdefault(fac, []).append(r)
        else:
            final.append(r)

    for key, group in groups.items():
        if len(group) == 1:
            group[0]["suspend_count"] = 1
            final.append(group[0])
            continue

        merged_sus = _merge_suspend_rows([g["suspend"] for g in group])
        merged_osbal_idx = set()
        merged_fac_codes = set()
        for g in group:
            merged_osbal_idx.update(g["osbal_idx"])
            merged_fac_codes.update(g["fac_codes"])

        ref = group[0]
        final.append({
            "suspend":      merged_sus,
            "source":       ref["source"],
            "scenario":     ref["scenario"],
            "osbal_idx":    list(merged_osbal_idx),
            "osbal_count":  len(merged_osbal_idx),
            "resolved":     True,
            "fac_codes":    merged_fac_codes,
            "suspend_count": len(group),
        })

    print(f"  {len(raw_results):,} suspend rows -> {len(final):,} output rows "
          f"({len(raw_results) - len(final):,} absorbed by accumulation)", flush=True)

    # [5] Build output DataFrame
    print(f"\n[5/5] Building output ({len(final):,} rows) ...", flush=True)

    output_rows = []
    for r in final:
        ref_osbal = [osbal_rows[i] for i in r["osbal_idx"]] if r["source"] else [{}]
        output_rows.append(_build_output_row(
            r["suspend"],
            source        = r["source"],
            scenario      = r["scenario"],
            osbal_rows    = ref_osbal,
            osbal_count   = r["osbal_count"],
            resolved      = r["resolved"],
            suspend_count = r.get("suspend_count", 1),
        ))

    df = pd.DataFrame(output_rows)

    for col in ["AMOUNT ORI", "CCOS_OR_BAL", "CCOS_BAL_DUE"]:
        df[col] = pd.to_numeric(df[col], errors="coerce")

    df = _compute_derived_cols(df)

    df.loc[df["SKENARIO"] == "UNMATCHED",   "SKENARIO"]  = "Unmatching"
    df.loc[df["SKENARIO"] == "Unmatching",  "FLAG_PROD"] = "Unmatching"

    for col in FINAL_COLUMNS:
        if col not in df.columns:
            df[col] = ""
    df = df[FINAL_COLUMNS]

    # Export ke Excel
    print(f"\nSaving to: {OUTPUT_FILE} ...", flush=True)

    for c in df.select_dtypes(include=['object']).columns:
        df[c] = (df[c]
                 .fillna('')
                 .astype(str)
                 .str.replace(r'[\x00-\x08\x0B\x0C\x0E-\x1F]', '', regex=True))

    df.to_excel(OUTPUT_FILE, index=False)

    try:
        from excel_styler import apply_purple_column_style
        apply_purple_column_style(OUTPUT_FILE, "matching")
    except Exception as e:
        print(f"  [WARN] Styling failed: {e}")

    # Export ke PostgreSQL
    load_dotenv()
    print("  Exporting to PostgreSQL ...")
    db_user = os.environ.get("DB_USER", "postgres")
    db_pass = os.environ.get("DB_PASS", "postgres")
    db_host = os.environ.get("DB_HOST", "localhost")
    db_port = os.environ.get("DB_PORT", "5432")
    db_name = os.environ.get("DB_NAME", "postgres")

    conn_str = f"postgresql+psycopg2://{db_user}:{db_pass}@{db_host}:{db_port}/{db_name}"
    try:
        engine = create_engine(conn_str)
        table_name = "SUSPENSE_DATA_ASTRA_BUANA"
        df.to_sql(table_name, con=engine, if_exists="append", index=False)
        print(f"      -> Successfully exported {len(df):,} rows to table '{table_name}'")
    except Exception as e:
        print(f"  [WARN] PostgreSQL export failed: {e}")

    elapsed = time.perf_counter() - t_start
    print(f"\n{'=' * 60}")
    print(f"  Done in {elapsed:.1f}s ({elapsed/60:.1f} min)")
    print(f"  Rows  : {len(df):,}")
    print(f"  Cols  : {len(df.columns)}")
    print(f"  File  : {OUTPUT_FILE}")

    print("\n  FLAG_PROD summary:")
    for flag, count in df["FLAG_PROD"].value_counts().items():
        print(f"    {flag:<45}: {count:,}")

    print("\n  SKENARIO summary:")
    for sce, count in df["SKENARIO"].value_counts().items():
        print(f"    {sce:<50}: {count:,}")

    print("=" * 60)


if __name__ == "__main__":
    run()
