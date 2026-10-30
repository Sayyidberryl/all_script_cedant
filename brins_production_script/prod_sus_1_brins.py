"""
================================================================================
 PRODUCTION SCRIPT — SUSPEND MATCHING V1 (CEDANT: BRINS GENERAL INSURANCE)
================================================================================
 Pipeline rekonsiliasi data SUSPEND vs OSBAL, SLIPDB, dan FACUL untuk
 Cedant PT. BRI Asuransi Indonesia (BRINS).

 Sesuai spesifikasi workflow:
   FASE 1: Slip    → OSBAL (exact) → FACUL → SLIPDB
   FASE 2: Polis   → OSBAL (exact) → FACUL → SLIPDB   [fallback jika Fase 1 gagal]
   FASE 3: Insured → OSBAL (exact) → FACUL → SLIPDB   [fallback jika Fase 2 gagal]
   FASE 3b: Nama Group → OSBAL (exact) → FACUL → SLIPDB
   FASE 4: Narrowing Wajib (Currency & Periode) + Strict Rematch Fallback
   FASE 5: Narrowing Non-Destruktif (Slip, Polis, Cert, Insured, Nama Group)
   POST  : Claim FAC Code, Akumulasi Suspend, FLAG_PROD 10 Kategori, Export Excel

 Output file: final_output_v1.xlsx
================================================================================
"""

import os
import re
import sys
import time
import functools

sys.modules["numexpr"] = None
sys.modules["bottleneck"] = None

import numpy as np
import pandas as pd


# ==============================================================================
# KONFIGURASI PATH & FILE
# ==============================================================================
SCRIPT_DIR     = os.path.dirname(os.path.abspath(__file__))
DATA_CLEAN_DIR = os.path.join(SCRIPT_DIR, "data_clean")
DATA_RAW_DIR   = os.path.join(SCRIPT_DIR, "data_raw")

def _find_slipdb_file() -> str:
    """Mencari path file Database Rekap RI Slip dari direktori data."""
    kandidat = [
        os.path.join(DATA_RAW_DIR, "5. Database Rekap RI Slip.xlsx"),
        os.path.join(DATA_CLEAN_DIR, "5. Database Rekap RI Slip.xlsx"),
        os.path.join(DATA_CLEAN_DIR, "Database Rekap RI Slip.xlsx"),
        os.path.join(DATA_CLEAN_DIR, "ri slip.xlsx"),
        os.path.join(DATA_CLEAN_DIR, "slipdb.xlsx"),
    ]
    for p in kandidat:
        if os.path.exists(p):
            return p
    return kandidat[0]

SUSPEND_FILE = os.path.join(DATA_CLEAN_DIR, "brins_output_suspend.xlsx")
OSBAL_FILE   = os.path.join(DATA_CLEAN_DIR, "brins_output_osbal_V2.xlsx")
FACUL_FILE   = os.path.join(DATA_CLEAN_DIR, "brins_output_facul.xlsx")
SLIPDB_FILE  = _find_slipdb_file()
OUTPUT_FILE  = os.path.join(SCRIPT_DIR, "final_output_v1.xlsx")


# ==============================================================================
# KONFIGURASI KOLOM
# ==============================================================================
OSBAL_FACODE_COL  = "CCOS_REF_CODE"
OSBAL_CURR_COL    = "CCOS_CURR"
OSBAL_DATE_COL    = "FAC_COM_DATE"
OSBAL_POLIS_ORI   = "FAC_POLICY_NO"
OSBAL_SLIP_ORI    = "FAC_SLIP"
OSBAL_INSURED_ORI = "FAC_INSURED"

FACUL_FACODE_COL  = "FAC_CODE"
FACUL_CURR_COL    = "FAC_CURRENCY"
FACUL_DATE_COL    = "FAC_COM_DATE"
FACUL_POLIS_ORI   = "FAC_POLICY_NO"
FACUL_SLIP_ORI    = "FAC_SLIP"
FACUL_INSURED_ORI = "FAC_INSURED"

SLIPDB_HEADER_ROW  = 1
SLIPDB_FACODE_COL  = "FAC. CODE"
SLIPDB_POLIS_COL   = "POLICY NO"
SLIPDB_SLIP_COL    = "SLIP"
SLIPDB_INSURED_COL = "INSURED"

SUSPEND_CURR_COL    = "CURR ORI"
SUSPEND_DATE_COL    = "RECEIPT DATE"
SUSPEND_STATUS_COL  = "STATUS"
SUSPEND_POLIS_ORI   = "POLIS"
SUSPEND_SLIP_ORI    = "SLIP NO"
SUSPEND_INSURED_ORI = "INSURED"

_EXCLUDED_REF_MARKERS = ["HUTANG PIUTANG", "DATA SUSPENSE"]
_LINESLIP_MARKERS     = ("LINESLIP", "LINE SLIP")

_MIN_TOKEN_LEN = 3
_TOKEN_RE      = re.compile(r"[^A-Z0-9]+")
_DELIMITER_RE  = re.compile(r"[,;|+\s]+")
_CERT_RE       = re.compile(r"-([0-9]{6})(?:[^0-9]|$)")

_SUSPEND_JOIN_COLS = [
    "RECEIPT NO", "CREDIT NOTES", "DETAIL RINCIAN NO", "RECEIPT DATE",
    "CEDANT NAME", "CEDANT SHRT NAME",
    "INSURED", "INSURED_CLN_1", "INSURED_CLN_2",
    "CURR ORI", "CURR PAY", "AMOUNT PAY",
    "POLIS", "POLIS_CLN_1", "CERTIFICATE",
    "SLIP NO", "SLIP_NO_CLN_1",
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
    "POLIS_ORI", "POLIS_CLN",
    "SERTIF_CLN",
    "SLIP_NO_ORI", "SLIP_NO_CLN",
    "DESC 1", "DESC 2", "DESC 3", "DESC 4",
    "STATUS", "REC_TYPE",
    "SKENARIO",
]


# ==============================================================================
# STRING UTILITIES
# ==============================================================================

@functools.lru_cache(maxsize=131072)
def _norm_cached(s: str) -> str:
    """Normalisasi string dengan huruf kapital dan menghapus spasi ganda."""
    text = s.strip().upper().replace("S/D", "SD")
    return re.sub(r"\s+", " ", text) if "  " in text else text


def _normalize(value) -> str:
    """Normalisasi nilai input menjadi string bersih terstandarisasi."""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    return _norm_cached(str(value))


def _tokenize(text: str) -> frozenset:
    """Memecah string menjadi himpunan token kata alphanumerik minimal 3 karakter."""
    if not text:
        return frozenset()
    return frozenset(t for t in _TOKEN_RE.split(text) if len(t) >= _MIN_TOKEN_LEN)


@functools.lru_cache(maxsize=100000)
def _split_composite(text: str) -> frozenset:
    """Memecah string komposit berdasarkan pemisah umum (+, /, |, koma, spasi)."""
    if not text:
        return frozenset()
    return frozenset(tok.strip() for tok in _DELIMITER_RE.split(text) if tok.strip())


def _clean_cert_str(val) -> str:
    """Membersihkan format string sertifikat dari spasi dan desimal."""
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return ""
    s = str(val).strip()
    return re.sub(r"\.0+$", "", s)


def _extract_cert_from_polis(polis_val: str) -> str:
    """Mengekstrak nomor sertifikat 6 digit dari format nomor polis (base-000000)."""
    if not polis_val:
        return ""
    m = _CERT_RE.search(polis_val)
    return m.group(1) if m else ""


def _format_sertif(val) -> str:
    """Memformat nomor sertifikat menjadi 6 digit angka dengan leading zero."""
    if pd.isna(val) or str(val).strip() == "":
        return ""
    if isinstance(val, (int, float)):
        return str(int(val)).zfill(6)
    s = str(val).strip()
    if s.endswith(".0"):
        s = s[:-2]
    return s.zfill(6) if s.isdigit() else s


def _get_clean_cols(columns: list, prefix: str) -> list:
    """Mengambil nama-nama kolom bersih yang berawalan prefix tertentu."""
    alt_map = {
        "clean polis":   ["polis_cln_", "policy_clean_"],
        "clean slip":    ["slip_no_cln_", "slip_clean_"],
        "clean insured": ["insured_cln_", "insured_clean_"],
        "clean sertif":  ["certificate_cln_", "sertif_cln_", "certificate_", "certificate", "sertif_clean_"],
    }
    prefixes = alt_map.get(prefix.lower(), [prefix.lower()])
    res = []
    for c in columns:
        clow = str(c).lower()
        if any(clow.startswith(p) for p in prefixes):
            res.append(c)
    return res


# ==============================================================================
# ROW DICT HELPERS
# ==============================================================================

def _is_excluded_ref_row(row: dict, polis_cols: list, slip_cols: list) -> bool:
    """Mengecek apakah baris referensi merupakan data administratif yang harus dikecualikan."""
    for col in list(polis_cols) + [OSBAL_POLIS_ORI] + list(slip_cols) + [OSBAL_SLIP_ORI]:
        val = _normalize(row.get(col, ""))
        if val and any(m in val for m in _EXCLUDED_REF_MARKERS):
            return True
    return False


def _is_lineslip_row(sus_row: dict, insured_cols: list) -> bool:
    """Mengecek apakah baris suspend terindikasi sebagai transaksi tipe LINESLIP."""
    for col in list(insured_cols) + [SUSPEND_INSURED_ORI]:
        val = _normalize(sus_row.get(col, ""))
        if val and any(m in val for m in _LINESLIP_MARKERS):
            return True
    return False


def _collect_clean_values(row: dict, cols: list) -> list:
    """Mengumpulkan nilai-nilai unik yang tidak kosong dari kolom bersih yang ditentukan."""
    seen, result = set(), []
    for col in cols:
        val = _normalize(row.get(col, ""))
        if val and val not in seen:
            seen.add(val)
            result.append(val)
    return result


def _collect_all_values(row: dict, clean_cols: list, ori_col: str = "") -> list:
    """Mengumpulkan nilai kolom bersih untuk pencarian rekonsiliasi."""
    return list(set(_collect_clean_values(row, clean_cols)))


def _get_group_variants(ins_vals: list, raw_insured: str = "") -> list:
    """Ekstrak variasi nama grup dari insured bersih dan insured raw."""
    results = set()
    for v in ins_vals:
        if not v:
            continue
        v_upper = v.upper().strip()
        results.add(v_upper)
        if v_upper.endswith(" GROUP"):
            base = v_upper[:-6].strip()
            if base:
                results.add(base)
        else:
            results.add(v_upper + " GROUP")

    if raw_insured:
        raw_upper = str(raw_insured).upper().strip()
        m = re.search(r"([A-Z0-9\s]{3,}?)\s+GROUP\b", raw_upper)
        if m:
            grp = m.group(1).strip()
            grp = re.sub(r"^(?:PREMI\s+)?(?:AF\s+)?(?:MARINE\s+CARGO\s+)?", "", grp).strip()
            if grp and len(grp) >= 3:
                results.add(grp)
                results.add(grp + " GROUP")

    return list(results)


def _ref_has_value(ref_row: dict, clean_cols: list, ori_col: str, query_values: list) -> bool:
    """Mengecek apakah baris referensi memuat nilai kueri (exact, token komposit, atau substring)."""
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


def _slip_in_range(sus_slip: str, osbal_row: dict) -> bool:
    """Mengecek apakah nomor slip suspend berada dalam rentang slip referensi OSBAL."""
    if not sus_slip:
        return False
    s_clean = re.sub(r"\D", "", str(sus_slip))
    if not s_clean:
        return False

    slip_vals = []
    for k, v in osbal_row.items():
        k_low = str(k).lower()
        if k_low.startswith("slip_no_cln_") and v:
            slip_vals.append(re.sub(r"\D", "", str(v)))

    if len(slip_vals) >= 2:
        sl1, sl2 = slip_vals[0], slip_vals[1]
        if sl1 and sl2 and len(sl1) == len(s_clean) and len(sl2) == len(s_clean):
            try:
                n_sus = int(s_clean)
                n_sl1, n_sl2 = int(sl1), int(sl2)
                if min(n_sl1, n_sl2) <= n_sus <= max(n_sl1, n_sl2):
                    return True
            except ValueError:
                pass
    return False


def _cert_in_range(query_cert: str, ref_cert_str: str) -> bool:
    """Mengecek apakah nomor sertifikat berada dalam rentang sertifikat referensi (misal: 100 SD 200)."""
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


def _row_has_official_cert(query_certs: set, ref_row: dict, sertif_cols: list) -> bool:
    """Cek kecocokan sertifikat pada kolom resmi sertifikat (CERTIFICATE_CLN_1, dsb)."""
    if not query_certs or not sertif_cols:
        return False
    for cv in query_certs:
        for col in sertif_cols:
            v = ref_row.get(col)
            if v and str(v).strip().lower() != "nan":
                if _cert_in_range(cv, str(v)):
                    return True
                s_v = str(v).strip()
                if str(cv).zfill(6) == s_v or str(cv) == s_v:
                    return True
    return False


def _row_has_cert_in_policy(query_certs: set, ref_row: dict, polis_cols: list = None) -> bool:
    """Fallback: Ekstrak rentang dan nomor sertifikat dari kolom polis bersih (POLIS_CLN_*, CLSDT_POLICY_NO)."""
    if not query_certs:
        return False
    check_cols = list(polis_cols) if polis_cols else []
    if "CLSDT_POLICY_NO" not in check_cols:
        check_cols.append("CLSDT_POLICY_NO")

    for col in check_cols:
        val = ref_row.get(col)
        if val is None or (isinstance(val, float) and pd.isna(val)):
            continue
        s = str(val).strip()
        if not s or s.lower() == "nan":
            continue

        clean_s = re.sub(r"\s*SD\s*|\s*S/D\s*", "-", s, flags=re.IGNORECASE)
        for m in re.finditer(r"(?<!\d)(\d{1,6})(?!\d)\s*-\s*(?<!\d)(\d{1,6})(?!\d)", clean_s):
            t1, t2 = m.group(1), m.group(2)
            st, en = int(t1), int(t2)
            if abs(en - st) < 5000:
                for qc in query_certs:
                    q_num = int(_clean_cert_str(qc)) if _clean_cert_str(qc).isdigit() else None
                    if q_num is not None and min(st, en) <= q_num <= max(st, en):
                        return True

        for tok in re.findall(r"(?:[\s\-]|^)(\d{1,5})(?=[/\s]|$)", s):
            for qc in query_certs:
                q_num = int(_clean_cert_str(qc)) if _clean_cert_str(qc).isdigit() else None
                if q_num is not None and int(tok) == q_num:
                    return True

    return False


def _row_has_cert(query_certs: set, ref_row: dict, sertif_cols: list, polis_cols: list = None) -> bool:
    """Cek kecocokan sertifikat (kolom resmi atau fallback policy)."""
    return _row_has_official_cert(query_certs, ref_row, sertif_cols) or _row_has_cert_in_policy(query_certs, ref_row, polis_cols)


# ==============================================================================
# INDEX BUILDERS
# ==============================================================================

def _expand_composite_policy(text: str) -> frozenset:
    """Ekspansi nomor polis komposit / berantai BRINS (mengekstrak 16 digit dasar).
    Contoh: '3114031023000024 071 974 123' -> '3114031023000024'
    '3114031023000024345' (polis 16 digit + cert 345 menempel) -> '3114031023000024'
    '3115010423000109 + 913 + 109...' -> '3115010423000109'
    """
    if not text or not isinstance(text, str):
        return frozenset()
    s = text.strip()
    if not s or s.lower() == "nan":
        return frozenset()
    results = set()
    m = re.match(r"^(\d{13,21})", s)
    if m:
        full = m.group(1)
        results.add(full)
        if len(full) > 16:
            results.add(full[:16])
    for tok in re.findall(r"\b\d{13,21}\b", s):
        results.add(tok)
        if len(tok) > 16:
            results.add(tok[:16])
    for tok in re.findall(r"\b[A-Z]{2}\d{10,14}\b", s, re.IGNORECASE):
        results.add(tok.upper())
    return frozenset(results)


def _build_exact_index(rows: list, cols: list, excluded: set = None, is_polis: bool = False) -> dict:
    """Membangun indeks pencarian exact match dari baris data dan kolom yang ditentukan."""
    index: dict = {}
    for i, row in enumerate(rows):
        if excluded and i in excluded:
            continue
        for col in cols:
            val = _normalize(row.get(col, ""))
            if val and val not in ("-", "--", "---", "NAN", "NONE", "TBA"):
                index.setdefault(val, []).append(i)
                if is_polis:
                    for p in _expand_composite_policy(val):
                        if p != val and p not in ("-", "--", "---", "NAN", "NONE", "TBA"):
                            index.setdefault(p, []).append(i)
    return index


def _build_facode_index(rows: list, col: str, excluded: set = None) -> dict:
    """Membangun indeks pemetaan FAC Code ke indeks baris referensi."""
    index: dict = {}
    for i, row in enumerate(rows):
        if excluded and i in excluded:
            continue
        val = _normalize(row.get(col, ""))
        if val:
            index.setdefault(val, []).append(i)
    return index


def _build_sertif_index(rows: list, sertif_cols: list, excluded: set = None) -> dict:
    """Membangun indeks pencarian nomor sertifikat dengan ekspansi format 6 digit."""
    index: dict = {}
    for i, row in enumerate(rows):
        if excluded and i in excluded:
            continue
        for col in sertif_cols:
            val = _normalize(row.get(col, ""))
            if val:
                index.setdefault(val, []).append(i)
                cert6 = _format_sertif(val)
                if cert6 and cert6 != val:
                    index.setdefault(cert6, []).append(i)
    return index


def _build_token_index(rows: list, clean_cols: list, excluded: set = None, is_polis: bool = False) -> tuple:
    """Membangun indeks token kata untuk pencocokan LIKE / parsial."""
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
                if is_polis:
                    tokens = tokens | _expand_composite_policy(t)
        token_cache[i] = tokens
        for tok in tokens:
            token_index.setdefault(tok, set()).add(i)
    return token_index, token_cache


def _build_lookup(rows: list, polis_cols: list, slip_cols: list,
                  insured_cols: list, facode_col: str = None, label: str = "") -> tuple:
    """Membangun seluruh struktur indeks lookup (exact, LIKE, facode) untuk tabel referensi."""
    if not rows:
        if label:
            print(f"  {label:<6} done (0 rows)", flush=True)
        return ({}, {}, {}, {}, {}, {}, ({}, {}), ({}, {}), ({}, {}), {})

    t0 = time.perf_counter()
    excluded = {i for i, r in enumerate(rows) if _is_excluded_ref_row(r, polis_cols, slip_cols)}

    facode_index = _build_facode_index(rows, facode_col, excluded=excluded) if facode_col else {}
    lookup = (
        _build_exact_index(rows, slip_cols,    excluded=excluded),
        _build_exact_index(rows, polis_cols,   excluded=excluded, is_polis=True),
        _build_exact_index(rows, insured_cols, excluded=excluded),
        {}, {}, {},
        _build_token_index(rows, slip_cols,    excluded=excluded),
        _build_token_index(rows, polis_cols,   excluded=excluded, is_polis=True),
        _build_token_index(rows, insured_cols, excluded=excluded),
        facode_index,
    )

    elapsed = time.perf_counter() - t0
    exc_info = f" - {len(excluded):,} dikecualikan" if excluded else ""
    if label:
        print(f"  {label:<6} done ({elapsed:.1f}s){exc_info}", flush=True)
    return lookup


# ==============================================================================
# MATCH ENGINE
# ==============================================================================

def _exact_match(query_values: list, index: dict) -> set:
    """Mencari indeks baris referensi yang cocok secara exact dengan nilai kueri."""
    return {i for v in query_values if v and v not in ("-", "--", "---", "NAN", "NONE", "TBA") and v in index for i in index[v]}


def _like_match(query_values: list, token_index: dict, rows: list, ref_cols: list) -> list:
    """Mencari indeks baris referensi yang cocok sebagian (LIKE substring)."""
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


def _resolve_facode(ref_rows_matched: list, facode_col: str,
                    facode_osbal_idx: dict, osbal_rows: list) -> tuple:
    """Menghubungkan baris referensi FACUL/SLIPDB ke baris OSBAL terkait via FAC Code."""
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


# ==============================================================================
# NARROWING
# ==============================================================================

def _narrow_by_currency(suspend_row: dict, matched: list, rows: list) -> list:
    """Menyaring baris kandidat berdasarkan kesesuaian mata uang (currency)."""
    if len(matched) <= 1:
        return matched
    sus_curr = _normalize(suspend_row.get(SUSPEND_CURR_COL, ""))
    if not sus_curr:
        return matched
    filtered = [i for i in matched if _normalize(rows[i].get(OSBAL_CURR_COL, "")) == sus_curr]
    return filtered if filtered else matched


def _narrow_by_periode(suspend_row: dict, matched: list, rows: list,
                       effective_sus_date=None, is_lineslip: bool = False) -> list:
    """Menyaring baris kandidat berdasarkan kesesuaian periode tanggal (receipt date vs com date)."""
    if len(matched) <= 1:
        return matched
    sus_date = effective_sus_date or suspend_row.get("_sus_date_parsed")
    if sus_date is None:
        return matched
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
    return valid if valid else matched


def _rematch_strict(
    suspend_row: dict, osbal_rows: list, lookup_osbal: tuple,
    polis_sus_cols: list, slip_sus_cols: list, insured_sus_cols: list,
    polis_ref_cols: list, slip_ref_cols: list, insured_ref_cols: list,
    facode_osbal_idx: dict,
    slipdb_rows: list = None, lookup_slipdb: tuple = None,
    facul_rows: list = None, lookup_facul: tuple = None,
    effective_sus_date=None, is_lineslip: bool = False,
) -> tuple:
    """Pencocokan ulang bertahap dengan filter ketat currency dan periode."""
    sus_curr = _normalize(suspend_row.get(SUSPEND_CURR_COL, ""))

    def _strict_filter(idxs):
        res = []
        for i in idxs:
            if sus_curr and _normalize(osbal_rows[i].get(OSBAL_CURR_COL, "")) != sus_curr:
                continue
            if effective_sus_date is not None and pd.notna(effective_sus_date):
                com_date = osbal_rows[i].get("_com_date_parsed")
                if com_date is not None and pd.notna(com_date):
                    if is_lineslip:
                        if com_date.year != effective_sus_date.year or com_date.month != effective_sus_date.month:
                            continue
                    else:
                        if effective_sus_date < com_date:
                            continue
            res.append(i)
        return res

    slip_vals  = _collect_clean_values(suspend_row, slip_sus_cols)
    polis_vals = _collect_clean_values(suspend_row, polis_sus_cols)

    if slip_vals:
        curr_hits = _strict_filter(list(_exact_match(slip_vals, lookup_osbal[0])))
        if curr_hits:
            confirmed = [i for i in curr_hits if any(
                _normalize(osbal_rows[i].get(c, "")) in polis_vals
                for c in polis_ref_cols if _normalize(osbal_rows[i].get(c, ""))
            )]
            chosen = confirmed if confirmed else curr_hits
            return "OSBAL", ("Slip + Polis" if confirmed else "Slip only"), chosen, True

    if polis_vals:
        curr_hits = _strict_filter(list(_exact_match(polis_vals, lookup_osbal[1])))
        if curr_hits:
            return "OSBAL", "Polis only", curr_hits, True

    if lookup_osbal[7] and lookup_osbal[7][0] and polis_vals:
        like_hits = _like_match(polis_vals, lookup_osbal[7][0], osbal_rows, polis_ref_cols)
        curr_hits = _strict_filter(like_hits)
        if curr_hits:
            return "OSBAL", "Polis only", curr_hits, True

    if slipdb_rows and lookup_slipdb:
        if slip_vals:
            hits = _exact_match(slip_vals, lookup_slipdb[0])
            if hits:
                res_idx, ok = _resolve_facode([slipdb_rows[i] for i in hits],
                                              SLIPDB_FACODE_COL, facode_osbal_idx, osbal_rows)
                if ok:
                    curr_res = _strict_filter(res_idx)
                    if curr_res:
                        return "SLIPDB", "Slip only", curr_res, True
        if polis_vals:
            hits = _exact_match(polis_vals, lookup_slipdb[1])
            if hits:
                res_idx, ok = _resolve_facode([slipdb_rows[i] for i in hits],
                                              SLIPDB_FACODE_COL, facode_osbal_idx, osbal_rows)
                if ok:
                    curr_res = _strict_filter(res_idx)
                    if curr_res:
                        return "SLIPDB", "Polis only", curr_res, True

    if facul_rows and lookup_facul:
        if slip_vals:
            hits = _exact_match(slip_vals, lookup_facul[0])
            if hits:
                res_idx, ok = _resolve_facode([facul_rows[i] for i in hits],
                                              FACUL_FACODE_COL, facode_osbal_idx, osbal_rows)
                if ok:
                    curr_res = _strict_filter(res_idx)
                    if curr_res:
                        return "FACUL", "Slip only", curr_res, True
        if polis_vals:
            hits = _exact_match(polis_vals, lookup_facul[1])
            if hits:
                res_idx, ok = _resolve_facode([facul_rows[i] for i in hits],
                                              FACUL_FACODE_COL, facode_osbal_idx, osbal_rows)
                if ok:
                    curr_res = _strict_filter(res_idx)
                    if curr_res:
                        return "FACUL", "Polis only", curr_res, True

    ins_vals = _collect_clean_values(suspend_row, insured_sus_cols)
    if ins_vals:
        curr_hits = _strict_filter(list(_exact_match(ins_vals, lookup_osbal[2])))
        if curr_hits:
            return "OSBAL", "Insured only", curr_hits, True

    return None


# ==============================================================================
# AGREGASI & OUTPUT BUILDERS
# ==============================================================================

def _aggregate_ccos(osbal_rows_matched: list) -> dict:
    """Menggabungkan data CCOS (nomor dokumen, FAC Code, balance due) dari baris OSBAL yang cocok."""
    fac_codes = {str(r.get(OSBAL_FACODE_COL, "")).strip() for r in osbal_rows_matched
                 if str(r.get(OSBAL_FACODE_COL, "")).strip()}
    multi_fac = len(fac_codes) > 1

    def join_unique(key):
        seen, vals = set(), []
        for r in osbal_rows_matched:
            v = str(r.get(key, "")).strip()
            if v and v not in seen:
                seen.add(v)
                vals.append(v)
        return ", ".join(vals)

    def sum_col(key):
        if multi_fac:
            return np.nan
        total = 0.0
        for r in osbal_rows_matched:
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
                    s = s.replace(".", "").replace(",", ".") if s.rfind(",") > s.rfind(".") else s.replace(",", "")
                elif "," in s:
                    s = s.replace(",", ".")
                try:
                    total += float(s)
                except ValueError:
                    pass
        return total

    return {
        "CCOS_DOC_NO":   join_unique("CCOS_DOC_NO"),
        "CCOS_REF_CODE": join_unique(OSBAL_FACODE_COL),
        "CCOS_OR_BAL":   np.nan if multi_fac else join_unique("CCOS_OR_BAL"),
        "CCOS_BAL_DUE":  sum_col("CCOS_BAL_DUE"),
    }


def _merge_suspend_rows(rows: list) -> dict:
    """Menggabungkan beberapa baris suspend yang mengklaim FAC Code unik yang sama."""
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
                if "," in s and "." in s:
                    s = s.replace(".", "").replace(",", ".") if s.rfind(",") > s.rfind(".") else s.replace(",", "")
                elif "," in s:
                    s = s.replace(",", ".")
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


def _build_output_row(
    suspend_row: dict, source: str, scenario: str, osbal_rows_matched: list,
    osbal_count: int = 0, suspend_count: int = 1,
) -> dict:
    """Menyusun baris dictionary output final sesuai struktur 32 kolom standar."""
    has_match = bool(source and osbal_rows_matched and osbal_rows_matched[0])
    ccos = (
        _aggregate_ccos(osbal_rows_matched)
        if has_match
        else {"CCOS_DOC_NO": "", "CCOS_REF_CODE": "", "CCOS_OR_BAL": np.nan, "CCOS_BAL_DUE": np.nan}
    )
    if "Beda Currency" in str(scenario) or "Beda Periode" in str(scenario):
        ccos["CCOS_OR_BAL"] = np.nan
        ccos["CCOS_BAL_DUE"] = np.nan

    final_scenario = "Unmatching" if not source or scenario in ("Unmatching", "UNMATCHED") else scenario

    return {
        "CCOS_DOC_NO":        ccos["CCOS_DOC_NO"] if has_match else "",
        "CCOS_REF_CODE":      ccos["CCOS_REF_CODE"],
        "RECEIPT NO":         suspend_row.get("RECEIPT NO",        ""),
        "CREDIT NOTES":       suspend_row.get("CREDIT NOTES",      ""),
        "DETAIL RINCIAN NO":  suspend_row.get("DETAIL RINCIAN NO", ""),
        "RECEIPT DATE":       suspend_row.get("RECEIPT DATE",      ""),
        "CEDANT NAME":        suspend_row.get("CEDANT NAME",       ""),
        "CEDANT SHRT NAME":   suspend_row.get("CEDANT SHRT NAME",  ""),
        "INSURED_ORI":        suspend_row.get("INSURED", ""),
        "INSURED_1":          suspend_row.get("INSURED_CLN_1", ""),
        "INSURED_2":          suspend_row.get("INSURED_CLN_2", ""),
        "CURR ORI":           suspend_row.get("CURR ORI",          ""),
        "AMOUNT ORI":         suspend_row.get("AMOUNT ORI",        ""),
        "AMOUNT_ORI_MIN1":    np.nan,
        "CURR PAY":           suspend_row.get("CURR PAY",          ""),
        "AMOUNT PAY":         suspend_row.get("AMOUNT PAY",        ""),
        "CCOS_OR_BAL":        ccos["CCOS_OR_BAL"],
        "CCOS_BAL_DUE":       ccos["CCOS_BAL_DUE"],
        "DIFERENCE":          np.nan,
        "FLAG_PROD":          "",
        "POLIS_ORI":          suspend_row.get("POLIS", ""),
        "POLIS_CLN":          suspend_row.get("POLIS_CLN_1", ""),
        "SERTIF_CLN":         _format_sertif(suspend_row.get("CERTIFICATE", "")),
        "SLIP_NO_ORI":        suspend_row.get("SLIP NO", ""),
        "SLIP_NO_CLN":        suspend_row.get("SLIP_NO_CLN_1", ""),
        "DESC 1":             suspend_row.get("DESC 1",            ""),
        "DESC 2":             suspend_row.get("DESC 2",            ""),
        "DESC 3":             suspend_row.get("DESC 3",            ""),
        "DESC 4":             suspend_row.get("DESC 4",            ""),
        "STATUS":             suspend_row.get("STATUS",            ""),
        "REC_TYPE":           suspend_row.get("REC_TYPE",          ""),
        "SKENARIO":           final_scenario,
        "_OSBAL_ROW_COUNT":   osbal_count,
        "_SUSPEND_ROW_COUNT": suspend_count,
    }


def _compute_derived_cols(df: pd.DataFrame) -> pd.DataFrame:
    """Menghitung kolom turunan (DIFERENCE, FLAG_PROD 9 kategori) berdasarkan perbandingan saldo."""
    amount_ori_base = df.get("AMOUNT ORI", pd.Series([0] * len(df), index=df.index))
    if "_MERGED_AMOUNT_ORI" in df.columns:
        amount_ori_raw = df["_MERGED_AMOUNT_ORI"].combine_first(amount_ori_base)
    else:
        amount_ori_raw = amount_ori_base

    amount_ori_clean = amount_ori_raw.replace(r"^\s*$", np.nan, regex=True)
    if amount_ori_clean.dtype == object:
        amount_ori_clean = amount_ori_clean.astype(str).str.strip().str.replace(r"\s+", "", regex=True)
        amount_ori_clean = amount_ori_clean.replace("nan", np.nan)
        has_comma = amount_ori_clean.str.contains(",", na=False)
        amount_ori_eu = (amount_ori_clean
                         .str.replace(".", "", regex=False)
                         .str.replace(",", ".", regex=False))
        amount_ori_clean = amount_ori_eu.where(has_comma, amount_ori_clean)
    amount_ori_numeric = pd.to_numeric(amount_ori_clean, errors="coerce")

    df["AMOUNT_ORI_MIN1"] = np.where(amount_ori_numeric.notna(), amount_ori_numeric * -1, np.nan)
    amount_neg = df["AMOUNT_ORI_MIN1"].fillna(0)

    balance_due     = pd.to_numeric(df.get("CCOS_BAL_DUE", pd.Series(0, index=df.index)), errors="coerce").fillna(0)
    osbal_count_col = pd.to_numeric(df.get("_OSBAL_ROW_COUNT", pd.Series(0, index=df.index)), errors="coerce").fillna(0)
    sus_count_col   = pd.to_numeric(df.get("_SUSPEND_ROW_COUNT", pd.Series(1, index=df.index)), errors="coerce").fillna(1)
    ccos_ref        = df.get("CCOS_REF_CODE", pd.Series("", index=df.index)).fillna("").astype(str)
    skenario_col    = df.get("SKENARIO", pd.Series("", index=df.index)).fillna("")

    df["DIFERENCE"] = amount_neg - balance_due
    accumulated     = (osbal_count_col > 1) | (sus_count_col > 1)
    amount_ori_num  = amount_ori_numeric.fillna(0)
    is_equal        = amount_neg.round(2) == balance_due.round(2)

    is_unmatched          = skenario_col.str.strip().str.upper().isin(["UNMATCHING", "UNMATCHED"]) | (skenario_col == "")
    is_beda_currency      = skenario_col.str.contains("Beda Currency", na=False, regex=False)
    is_beda_periode       = skenario_col.str.contains("Beda Periode", na=False, regex=False)
    is_gt1_fac            = ccos_ref.str.contains(",", na=False) & ~is_beda_currency & ~is_beda_periode
    is_new_entry_total    = ~is_unmatched & (amount_ori_num != 0) & (balance_due == 0)
    is_adj_total          = ~is_unmatched & is_equal
    is_adj_sebagian       = ~is_unmatched & ~is_equal & (amount_neg < balance_due)
    is_new_entry_sebagian = ~is_unmatched & ~is_equal & (amount_neg > balance_due)

    df["FLAG_PROD"] = np.select(
        [is_beda_currency, is_beda_periode, is_gt1_fac,
         is_new_entry_total,
         is_adj_total & ~accumulated, is_adj_total & accumulated,
         is_adj_sebagian & ~accumulated, is_adj_sebagian & accumulated,
         is_new_entry_sebagian],
        ["Beda Currency", "Beda Periode", "Matching >1 fac code",
         "New Entry total",
         "Adjustment total tanpa akumulasi", "Adjustment total dengan akumulasi",
         "Adjustment sebagian tanpa akumulasi", "Adjustment sebagian dengan akumulasi",
         "New Entry sebagian"],
        default="Unmatching",
    )

    for col in ["_OSBAL_ROW_COUNT", "_SUSPEND_ROW_COUNT", "_MERGED_AMOUNT_ORI"]:
        if col in df.columns:
            df = df.drop(columns=[col])
    return df


# ==============================================================================
# DATA LOADER
# ==============================================================================

def _load_excel(path: str, label: str, optional: bool = False,
                header: int = 0) -> tuple:
    """Membaca file Excel dengan dukungan pickle cache untuk efisiensi pemuatan data."""
    if not os.path.exists(path):
        print(f"  {label}: {path} ({'tidak ada, dilewati' if optional else 'TIDAK DITEMUKAN'})", flush=True)
        return [], []

    cache = path + ".cache.pkl"
    if os.path.exists(cache) and os.path.getmtime(cache) >= os.path.getmtime(path):
        t = time.perf_counter()
        print(f"  {label}: loading dari cache ...", flush=True)
        try:
            data, hdr = pd.read_pickle(cache)
            print(f"  -> {len(data):,} rows, {len(hdr)} cols ({time.perf_counter()-t:.2f}s)", flush=True)
            return data, hdr
        except Exception:
            print(f"  {label}: cache rusak, baca ulang ...", flush=True)

    t = time.perf_counter()
    print(f"  {label}: {os.path.basename(path)} ...", flush=True)
    try:
        df = pd.read_excel(path, header=header, engine="calamine")
    except Exception:
        try:
            df = pd.read_excel(path, header=header)
        except Exception as e:
            print(f"  [RETRY] gagal ({e}), coba openpyxl ...", flush=True)
            df = pd.read_excel(path, header=header, engine="openpyxl")

    hdr  = list(df.columns)
    data = df.to_dict("records")
    print(f"  -> {len(data):,} rows, {len(hdr)} cols ({time.perf_counter()-t:.1f}s)", flush=True)

    try:
        pd.to_pickle((data, hdr), cache)
        print(f"  -> cache disimpan: {os.path.basename(cache)}", flush=True)
    except Exception:
        pass
    return data, hdr


# ==============================================================================
# PIPELINE UTAMA
# ==============================================================================

def run() -> None:
    """Menjalankan seluruh pipeline matching rekonsiliasi suspend BRINS.

    Pipeline meliputi tahapan:
    1. Membaca data Suspend, OSBAL, FACUL, dan SLIPDB.
    2. Mendeteksi kolom-kolom bersih (clean columns).
    3. Membangun indeks pencarian cepat (lookup indexes).
    4. Melakukan matching bertingkat (Fase 1 Slip, Fase 2 Polis, Fase 3 Insured,
       Fase 4 Narrowing Currency & Periode, Fase 5 Narrowing Bertingkat).
    5. Agregasi CCOS, klaim FAC Code unik, akumulasi, dan export hasil matching.
    """
    t_start = time.perf_counter()
    print("\n" + "=" * 64, flush=True)
    print("  PRODUCTION SCRIPT — SUSPEND MATCHING  (BRINS) — V1", flush=True)
    print("=" * 64, flush=True)

    # [1] Load data
    print("\n[1/5] Loading data ...", flush=True)
    suspend_rows, suspend_cols = _load_excel(SUSPEND_FILE, "SUSPEND")
    osbal_rows,   osbal_cols   = _load_excel(OSBAL_FILE,   "OSBAL")
    facul_rows,   facul_cols   = _load_excel(FACUL_FILE,   "FACUL")
    slipdb_rows,  slipdb_cols  = _load_excel(SLIPDB_FILE,  "SLIPDB",
                                              optional=True, header=SLIPDB_HEADER_ROW)

    if not suspend_rows:
        print("[ERROR] Data suspend tidak ditemukan. Jalankan cleansing terlebih dahulu.", flush=True)
        return
    if not osbal_rows:
        print("[ERROR] Data OSBAL tidak ditemukan. Jalankan cleansing terlebih dahulu.", flush=True)
        return

    # Pre-parse tanggal OSBAL & Suspend
    print("  Pre-parsing OSBAL FAC_COM_DATE ...", flush=True)
    for r in osbal_rows:
        raw = r.get(OSBAL_DATE_COL, "")
        r["_com_date_parsed"] = (
            pd.to_datetime(raw, errors="coerce")
            if raw and not (isinstance(raw, float) and pd.isna(raw))
            else None
        )

    print("  Pre-parsing Suspend RECEIPT DATE ...", flush=True)
    for r in suspend_rows:
        raw = r.get(SUSPEND_DATE_COL, "")
        r["_sus_date_parsed"] = (
            pd.to_datetime(raw, errors="coerce")
            if raw and not (isinstance(raw, float) and pd.isna(raw))
            else None
        )

    # [2] Deteksi kolom bersih
    print("\n[2/5] Detecting clean columns ...", flush=True)
    polis_sus   = _get_clean_cols(suspend_cols, "clean polis")
    slip_sus    = _get_clean_cols(suspend_cols, "clean slip")
    insured_sus = _get_clean_cols(suspend_cols, "clean insured")
    sertif_sus  = _get_clean_cols(suspend_cols, "clean sertif")

    if not polis_sus:   polis_sus   = [c for c in ["POLIS_CLN_1"] if c in suspend_cols]
    if not slip_sus:    slip_sus    = [c for c in ["SLIP_NO_CLN_1"] if c in suspend_cols]
    if not insured_sus: insured_sus = [c for c in ["INSURED_CLN_1", "INSURED_CLN_2"] if c in suspend_cols]
    if not sertif_sus:  sertif_sus  = [c for c in ["CERTIFICATE"] if c in suspend_cols]

    print(f"  Suspend polis cols   : {polis_sus}", flush=True)
    print(f"  Suspend slip cols    : {slip_sus}", flush=True)
    print(f"  Suspend insured cols : {insured_sus}", flush=True)
    print(f"  Suspend sertif cols  : {sertif_sus}", flush=True)

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

    if not polis_osbal:   polis_osbal   = [c for c in osbal_cols if str(c).upper().startswith("POLIS_CLN_")]
    if not slip_osbal:    slip_osbal    = [c for c in osbal_cols if str(c).upper().startswith("SLIP_NO_CLN_")]
    if not insured_osbal: insured_osbal = [c for c in osbal_cols if str(c).upper().startswith("INSURED_CLN_")]
    if not sertif_osbal:  sertif_osbal  = [c for c in osbal_cols if str(c).upper().startswith("CERTIFICATE_CLN_")]

    def _prepend(col, cols_list, header):
        return ([col] + cols_list) if col in header and col not in cols_list else cols_list

    polis_osbal = _prepend("CLSDT_POLICY_NO", polis_osbal, osbal_cols)
    slip_osbal  = _prepend("CLSDT_SLIP_NO",   slip_osbal,  osbal_cols)
    if "CLSDT_SERTF_NO" in osbal_cols and "CLSDT_SERTF_NO" not in sertif_osbal:
        sertif_osbal = ["CLSDT_SERTF_NO"] + sertif_osbal

    print(f"  OSBAL polis cols     : {polis_osbal[:4]}{'...' if len(polis_osbal) > 4 else ''}", flush=True)
    print(f"  OSBAL slip cols      : {slip_osbal[:4]}{'...' if len(slip_osbal) > 4 else ''}", flush=True)
    print(f"  OSBAL insured cols   : {insured_osbal[:4]}{'...' if len(insured_osbal) > 4 else ''}", flush=True)

    polis_facul   = _get_clean_cols(facul_cols, "clean polis")
    slip_facul    = _get_clean_cols(facul_cols, "clean slip")
    insured_facul = _get_clean_cols(facul_cols, "clean insured")

    if not polis_facul:   polis_facul   = [c for c in facul_cols if str(c).upper().startswith("POLIS_CLN_")]
    if not slip_facul:    slip_facul    = [c for c in facul_cols if str(c).upper().startswith("SLIP_NO_CLN_")]
    if not insured_facul: insured_facul = [c for c in facul_cols if str(c).upper().startswith("INSURED_CLN_")]

    polis_facul = _prepend("CLSDT_POLICY_NO", polis_facul, facul_cols)
    slip_facul  = _prepend("CLSDT_SLIP_NO",   slip_facul,  facul_cols)

    polis_slipdb   = [c for c in [SLIPDB_POLIS_COL] if c in slipdb_cols]
    slip_slipdb    = [c for c in [SLIPDB_SLIP_COL]   if c in slipdb_cols]
    insured_slipdb = [c for c in [SLIPDB_INSURED_COL] if c in slipdb_cols]

    slipdb_facode_col = SLIPDB_FACODE_COL
    if slipdb_facode_col not in slipdb_cols:
        for candidate in ["FAC_CODE", "FAC CODE", "CCOS_REF_CODE"]:
            if candidate in slipdb_cols:
                slipdb_facode_col = candidate
                break

    # [3] Bangun Lookup Index
    print("\n[3/5] Building lookup indexes ...", flush=True)
    lookup_osbal  = _build_lookup(osbal_rows,  polis_osbal,  slip_osbal,  insured_osbal,  OSBAL_FACODE_COL,   label="OSBAL")
    lookup_facul  = _build_lookup(facul_rows,  polis_facul,  slip_facul,  insured_facul,  FACUL_FACODE_COL,   label="FACUL")
    lookup_slipdb = _build_lookup(slipdb_rows, polis_slipdb, slip_slipdb, insured_slipdb, slipdb_facode_col,  label="SLIPDB")

    facode_osbal_idx = lookup_osbal[9]
    excluded_osbal   = {i for i, r in enumerate(osbal_rows) if _is_excluded_ref_row(r, polis_osbal, slip_osbal)}
    sertif_osbal_idx = _build_sertif_index(osbal_rows, sertif_osbal, excluded=excluded_osbal)
    print(f"  Sertif OSBAL index   : {len(sertif_osbal_idx):,} entri unik", flush=True)

    # [4] Matching per baris Suspend
    total = len(suspend_rows)
    print(f"\n[4/5] Matching {total:,} suspend rows ...", flush=True)
    t4 = time.perf_counter()
    raw_results = []

    for n, sus in enumerate(suspend_rows, 1):
        if n % 500 == 0:
            print(f"    Memproses baris {n:,}/{total:,} ...", flush=True)

        if str(sus.get(SUSPEND_STATUS_COL, "")).strip().upper() == "MATCHING":
            continue

        _is_ls = sus.get("_is_lineslip", False)
        effective_date = sus.get("_sus_date_lineslip") if _is_ls else sus.get("_sus_date_parsed")

        matched_osbal_idx: set = set()
        source       = None
        scenario     = ""
        resolved     = False
        is_beda_curr = False
        is_beda_per  = False

        # FASE 1: SLIP
        slip_vals = _collect_clean_values(sus, slip_sus)
        slip_vals = [v for v in slip_vals if v and v not in ("-", "--", "---", "NAN", "NONE", "TBA")]

        found_by_slip = False
        if slip_vals:
            osb_hits = _exact_match(slip_vals, lookup_osbal[0])
            if osb_hits:
                matched_osbal_idx.update(osb_hits)
                found_by_slip = True
                source   = "OSBAL"
                scenario = "Slip only"

            if facul_rows and not found_by_slip:
                fac_hits = _exact_match(slip_vals, lookup_facul[0])
                if fac_hits:
                    res_idx, _ = _resolve_facode(
                        [facul_rows[i] for i in fac_hits],
                        FACUL_FACODE_COL, facode_osbal_idx, osbal_rows
                    )
                    if res_idx:
                        matched_osbal_idx.update(res_idx)
                        found_by_slip = True
                        source   = "FACUL"
                        scenario = "Slip only"
                        resolved = True

            if slipdb_rows and not found_by_slip:
                sld_hits = _exact_match(slip_vals, lookup_slipdb[0])
                if sld_hits:
                    res_idx, _ = _resolve_facode(
                        [slipdb_rows[i] for i in sld_hits],
                        slipdb_facode_col, facode_osbal_idx, osbal_rows
                    )
                    if res_idx:
                        matched_osbal_idx.update(res_idx)
                        found_by_slip = True
                        source   = "SLIPDB"
                        scenario = "Slip only"
                        resolved = True

        # FASE 2: POLIS (fallback jika Slip gagal)
        found_by_polis = False
        if not matched_osbal_idx:
            polis_vals = _collect_clean_values(sus, polis_sus)
            polis_vals = [v for v in polis_vals if v and v not in ("-", "--", "---", "NAN", "NONE", "TBA")]

            if polis_vals:
                osb_hits = _exact_match(polis_vals, lookup_osbal[1])
                if not osb_hits and lookup_osbal[7] and lookup_osbal[7][0]:
                    osb_hits = set(_like_match(polis_vals, lookup_osbal[7][0], osbal_rows, polis_osbal))
                if osb_hits:
                    matched_osbal_idx.update(osb_hits)
                    found_by_polis = True
                    source   = "OSBAL"
                    scenario = "Polis only"

                if facul_rows and not found_by_polis:
                    fac_hits = _exact_match(polis_vals, lookup_facul[1])
                    if not fac_hits and lookup_facul[7] and lookup_facul[7][0]:
                        fac_hits = set(_like_match(polis_vals, lookup_facul[7][0], facul_rows, polis_facul))
                    if fac_hits:
                        res_idx, _ = _resolve_facode(
                            [facul_rows[i] for i in fac_hits],
                            FACUL_FACODE_COL, facode_osbal_idx, osbal_rows
                        )
                        if res_idx:
                            matched_osbal_idx.update(res_idx)
                            found_by_polis = True
                            source   = "FACUL"
                            scenario = "Polis only"
                            resolved = True

                if slipdb_rows and not found_by_polis:
                    sld_hits = _exact_match(polis_vals, lookup_slipdb[1])
                    if not sld_hits and lookup_slipdb[7] and lookup_slipdb[7][0]:
                        sld_hits = set(_like_match(polis_vals, lookup_slipdb[7][0], slipdb_rows, polis_slipdb))
                    if sld_hits:
                        res_idx, _ = _resolve_facode(
                            [slipdb_rows[i] for i in sld_hits],
                            slipdb_facode_col, facode_osbal_idx, osbal_rows
                        )
                        if res_idx:
                            matched_osbal_idx.update(res_idx)
                            found_by_polis = True
                            source   = "SLIPDB"
                            scenario = "Polis only"
                            resolved = True

        # FASE 3: INSURED & FASE 3b: NAMA GROUP (fallback jika Polis gagal)
        if not matched_osbal_idx:
            ins_vals = _collect_clean_values(sus, insured_sus)
            ins_vals = [v for v in ins_vals if v and v not in ("-", "--", "---", "NAN", "NONE", "TBA")]

            if ins_vals:
                osb_hits = _exact_match(ins_vals, lookup_osbal[2])
                if osb_hits:
                    matched_osbal_idx.update(osb_hits)
                    source   = "OSBAL"
                    scenario = "Insured only"
                elif facul_rows:
                    fac_hits = _exact_match(ins_vals, lookup_facul[2])
                    if fac_hits:
                        res_idx, _ = _resolve_facode(
                            [facul_rows[i] for i in fac_hits],
                            FACUL_FACODE_COL, facode_osbal_idx, osbal_rows
                        )
                        if res_idx:
                            matched_osbal_idx.update(res_idx)
                            source   = "FACUL"
                            scenario = "Insured only"
                            resolved = True
                if not matched_osbal_idx and slipdb_rows:
                    sld_hits = _exact_match(ins_vals, lookup_slipdb[2])
                    if sld_hits:
                        res_idx, _ = _resolve_facode(
                            [slipdb_rows[i] for i in sld_hits],
                            slipdb_facode_col, facode_osbal_idx, osbal_rows
                        )
                        if res_idx:
                            matched_osbal_idx.update(res_idx)
                            source   = "SLIPDB"
                            scenario = "Insured only"
                            resolved = True

            # FASE 3b: Nama Group Fallback (jika Insured exact belum ketemu)
            if not matched_osbal_idx:
                grp_vals = _get_group_variants(ins_vals, sus.get(SUSPEND_INSURED_ORI, ""))
                grp_vals = [v for v in grp_vals if v and v not in ("-", "--", "---", "NAN", "NONE", "TBA")]
                if grp_vals:
                    osb_hits = _exact_match(grp_vals, lookup_osbal[2])
                    if osb_hits:
                        matched_osbal_idx.update(osb_hits)
                        source   = "OSBAL"
                        scenario = "Insured only"
                    elif facul_rows:
                        fac_hits = _exact_match(grp_vals, lookup_facul[2])
                        if fac_hits:
                            res_idx, _ = _resolve_facode(
                                [facul_rows[i] for i in fac_hits],
                                FACUL_FACODE_COL, facode_osbal_idx, osbal_rows
                            )
                            if res_idx:
                                matched_osbal_idx.update(res_idx)
                                source   = "FACUL"
                                scenario = "Insured only"
                                resolved = True
                    if not matched_osbal_idx and slipdb_rows:
                        sld_hits = _exact_match(grp_vals, lookup_slipdb[2])
                        if sld_hits:
                            res_idx, _ = _resolve_facode(
                                [slipdb_rows[i] for i in sld_hits],
                                slipdb_facode_col, facode_osbal_idx, osbal_rows
                            )
                            if res_idx:
                                matched_osbal_idx.update(res_idx)
                                source   = "SLIPDB"
                                scenario = "Insured only"
                                resolved = True

        # FASE 4: NARROWING WAJIB (Currency & Periode) + Rematch Strict
        if matched_osbal_idx:
            matched_list = list(matched_osbal_idx)

            sus_curr = _normalize(sus.get(SUSPEND_CURR_COL, ""))
            if sus_curr:
                filtered_curr = [i for i in matched_list
                                 if _normalize(osbal_rows[i].get(OSBAL_CURR_COL, "")) == sus_curr]
                if not filtered_curr:
                    is_beda_curr = True
                elif len(filtered_curr) < len(matched_list):
                    matched_list = filtered_curr

            if not is_beda_curr and effective_date is not None and pd.notna(effective_date):
                filtered_per = []
                for i in matched_list:
                    com_date = osbal_rows[i].get("_com_date_parsed")
                    if com_date is None or pd.isna(com_date):
                        filtered_per.append(i)
                    else:
                        try:
                            if _is_ls:
                                if com_date.year == effective_date.year and com_date.month == effective_date.month:
                                    filtered_per.append(i)
                            else:
                                if effective_date >= com_date:
                                    filtered_per.append(i)
                        except Exception:
                            filtered_per.append(i)
                if not filtered_per:
                    is_beda_per = True
                elif len(filtered_per) < len(matched_list):
                    matched_list = filtered_per

            if is_beda_curr or is_beda_per:
                rm = _rematch_strict(
                    sus, osbal_rows, lookup_osbal,
                    polis_sus, slip_sus, insured_sus,
                    polis_osbal, slip_osbal, insured_osbal,
                    facode_osbal_idx,
                    slipdb_rows, lookup_slipdb,
                    facul_rows, lookup_facul,
                    effective_sus_date=effective_date, is_lineslip=_is_ls,
                )
                if rm:
                    src, scen, hits, _ = rm
                    source       = src
                    scenario     = scen
                    matched_list = hits
                    is_beda_curr = False
                    is_beda_per  = False

            matched_osbal_idx = set(matched_list)

        # FASE 5: NARROWING NON-DESTRUKTIF
        # Urutan alur: Slip + Polis + Cert + Insured + Nama Group
        if len(matched_osbal_idx) > 1:
            matched_list = list(matched_osbal_idx)

            # 1. Polis Matching (jika base awal adalah Slip)
            polis_vals = _collect_clean_values(sus, polis_sus)
            polis_vals = [v for v in polis_vals if v and v not in ("-", "--", "---", "NAN", "NONE", "TBA")]

            if polis_vals:
                confirmed_polis = [idx for idx in matched_list
                                   if _ref_has_value(osbal_rows[idx], polis_osbal, "", polis_vals)]
                if confirmed_polis:
                    matched_list = confirmed_polis
                    if "Slip" in scenario:
                        scenario = "Slip + Polis"
                    elif "Polis" not in scenario:
                        if "only" in scenario:
                            scenario = scenario.replace(" only", "") + " + Polis"
                        else:
                            scenario = scenario + " + Polis"

            # 1b. Slip Matching (jika base awal adalah Polis atau Insured)
            slip_vals = _collect_clean_values(sus, slip_sus)
            slip_vals = [v for v in slip_vals if v and v not in ("-", "--", "---", "NAN", "NONE", "TBA")]

            if len(matched_list) > 1 and slip_vals:
                confirmed_slip = [idx for idx in matched_list
                                  if _ref_has_value(osbal_rows[idx], slip_osbal, "", slip_vals)]
                if not confirmed_slip:
                    confirmed_slip = [idx for idx in matched_list
                                      if any(_slip_in_range(sv, osbal_rows[idx]) for sv in slip_vals)]
                if confirmed_slip:
                    matched_list = confirmed_slip
                    if "Polis" in scenario and "Slip" not in scenario:
                        scenario = "Slip + Polis"

            # 2. Sertifikat (certf)
            sus_cert_vals = set(
                _clean_cert_str(v)
                for v in _collect_clean_values(sus, sertif_sus)
                if _clean_cert_str(v)
            )
            for pv in ([_normalize(sus.get(c, "")) for c in polis_sus] +
                      (polis_vals if polis_vals else [])):
                if pv:
                    c = _extract_cert_from_polis(pv)
                    if c: sus_cert_vals.add(_clean_cert_str(c))
            sus_cert_vals.discard("")

            if len(matched_list) > 1 and sus_cert_vals:
                # Prioritas 1: Kolom resmi sertifikat (CERTIFICATE_CLN_1, CLSDT_SERTF_NO)
                confirmed_cert = [idx for idx in matched_list
                                  if _row_has_official_cert(sus_cert_vals, osbal_rows[idx], sertif_osbal)]
                # Prioritas 2: Fallback nomor sertifikat di kolom polis
                if not confirmed_cert:
                    confirmed_cert = [idx for idx in matched_list
                                      if _row_has_cert_in_policy(sus_cert_vals, osbal_rows[idx], polis_osbal)]

                if confirmed_cert:
                    matched_list = confirmed_cert
                    if "Cert" not in scenario:
                        if "only" in scenario:
                            scenario = scenario.replace(" only", "") + " + Cert"
                        else:
                            scenario = scenario + " + Cert"

            # 3. Insured (insured)
            ins_vals = _collect_all_values(sus, insured_sus)
            ins_vals = [v for v in ins_vals if v and v not in ("-", "--", "---", "NAN", "NONE", "TBA")]

            if len(matched_list) > 1 and ins_vals:
                confirmed_ins = [idx for idx in matched_list
                                 if _ref_has_value(osbal_rows[idx], insured_osbal, "", ins_vals)]
                if confirmed_ins:
                    matched_list = confirmed_ins
                    if "Slip" not in scenario and "Polis" not in scenario:
                        if "Insured" not in scenario:
                            scenario = scenario + " + Insured"

            # 4. Nama Group (nama group)
            if len(matched_list) > 1:
                grp_vals = _get_group_variants(ins_vals, sus.get(SUSPEND_INSURED_ORI, ""))
                grp_vals = [v for v in grp_vals if v and v not in ("-", "--", "---", "NAN", "NONE", "TBA")]
                if grp_vals:
                    confirmed_grp = [idx for idx in matched_list
                                     if _ref_has_value(osbal_rows[idx], insured_osbal, "", grp_vals)]
                    if confirmed_grp:
                        matched_list = confirmed_grp
                        if "Slip" not in scenario and "Polis" not in scenario:
                            if "Insured" not in scenario:
                                scenario = scenario + " + Insured"

            matched_osbal_idx = set(matched_list)

        # Post-matching label
        if is_beda_curr or is_beda_per:
            osbal_ref   = [osbal_rows[i] for i in matched_list] if matched_list else [{}]
            osbal_count = 0
            resolved    = False
        elif not matched_osbal_idx:
            source      = None
            scenario    = "Unmatching"
            osbal_ref   = [{}]
            osbal_count = 0
            resolved    = False
        else:
            osbal_ref   = [osbal_rows[i] for i in matched_osbal_idx]
            osbal_count = len(osbal_ref)

        if scenario and " only + " in scenario:
            scenario = scenario.replace(" only + ", " + ")
        if "Slip" in scenario or "Polis" in scenario:
            scenario = re.sub(r"\s*\+\s*Insured", "", scenario).strip()

        fac_codes = {_normalize(r.get(OSBAL_FACODE_COL, "")) for r in osbal_ref if r}
        fac_codes.discard("")

        if is_beda_curr:
            scenario = "Beda Currency"
        elif is_beda_per:
            scenario = "Beda Periode"

        raw_results.append({
            "suspend":       sus,
            "source":        source,
            "base_scenario": scenario,
            "is_beda_curr":  is_beda_curr,
            "is_beda_per":   is_beda_per,
            "osbal_idx":     list(matched_osbal_idx) if not (is_beda_curr or is_beda_per) else (list(matched_list) if matched_list else []),
            "osbal_count":   osbal_count,
            "resolved":      resolved,
            "fac_codes":     fac_codes,
        })

    print(f"  Matching selesai dalam {time.perf_counter()-t4:.1f}s", flush=True)

    # Claim FAC Code & Akumulasi Suspend
    print("\n[4c/5] Fac code claim & akumulasi ...", flush=True)
    claimed: set = set()
    for r in raw_results:
        if (len(r["fac_codes"]) == 1 and r["source"] is not None
                and not r["is_beda_curr"] and not r["is_beda_per"]):
            claimed |= r["fac_codes"]

    for r in raw_results:
        if len(r["fac_codes"]) > 1:
            remaining = r["fac_codes"] - claimed
            if (len(remaining) == 1 and r["source"] is not None
                    and not r["is_beda_curr"] and not r["is_beda_per"]):
                r["fac_codes"] = remaining
                claimed |= remaining
                fac = next(iter(remaining))
                r["osbal_idx"] = [i for i in r["osbal_idx"]
                                   if _normalize(osbal_rows[i].get(OSBAL_FACODE_COL, "")) == fac]
                r["osbal_count"] = len(r["osbal_idx"])

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
            group[0]["suspend_count"] = 1
            final.append(group[0])
            continue
        merged_osbal_idx = set()
        for g in group:
            merged_osbal_idx.update(g["osbal_idx"])
        merged_sus = _merge_suspend_rows([g["suspend"] for g in group])
        for g in group:
            g["osbal_idx"]          = list(merged_osbal_idx)
            g["osbal_count"]        = len(merged_osbal_idx)
            g["suspend_count"]      = len(group)
            g["_merged_amount_ori"] = merged_sus.get("AMOUNT ORI", 0)
            final.append(g)

    print(f"  {len(raw_results):,} suspend rows -> {len(final):,} output rows", flush=True)

    # [5] Build Output DataFrame & Export
    print(f"\n[5/5] Building output ({len(final):,} rows) & saving ...", flush=True)
    output_rows = []
    for r in final:
        ref_osbal = [osbal_rows[i] for i in r["osbal_idx"]] if r["source"] else [{}]
        out = _build_output_row(
            r["suspend"],
            source=r["source"],
            scenario=r["base_scenario"],
            osbal_rows_matched=ref_osbal,
            osbal_count=r["osbal_count"],
            suspend_count=r.get("suspend_count", 1),
        )
        if "_merged_amount_ori" in r:
            out["_MERGED_AMOUNT_ORI"] = r["_merged_amount_ori"]
        output_rows.append(out)

    df = pd.DataFrame(output_rows)

    # Numerik cleanup
    for col in ["AMOUNT ORI", "CCOS_OR_BAL", "CCOS_BAL_DUE"]:
        if col in df.columns:
            s = df[col]
            if s.dtype == object:
                s = s.astype(str).str.strip().str.replace(r"\s+", "", regex=True)
                has_comma = s.str.contains(",", na=False)
                s_eu = s.str.replace(".", "", regex=False).str.replace(",", ".", regex=False)
                s = s_eu.where(has_comma, s)
            df[col] = pd.to_numeric(s, errors="coerce")

    df = _compute_derived_cols(df)
    df.loc[df["SKENARIO"].isin(["UNMATCHED", "Unmatching"]), "FLAG_PROD"] = "Unmatching"

    for col in FINAL_COLUMNS:
        if col not in df.columns:
            df[col] = ""
    df = df[FINAL_COLUMNS]

    # Bersihkan karakter kontrol
    for c in df.select_dtypes(include=["object", "string"]).columns:
        df[c] = df[c].astype(str).str.replace(r"[\x00-\x08\x0B\x0C\x0E-\x1F]", "", regex=True)

    # Export ke Excel (hanya sheet Output_Matching)
    save_dest = OUTPUT_FILE
    try:
        with pd.ExcelWriter(save_dest, engine="openpyxl") as writer:
            df.to_excel(writer, sheet_name="Output_Matching", index=False)
        print(f"  [OK] Tersimpan di: {save_dest} (Sheet: Output_Matching)", flush=True)
    except PermissionError:
        base, ext = os.path.splitext(OUTPUT_FILE)
        save_dest = f"{base}_latest{ext}"
        print(f"  [WARN] File {OUTPUT_FILE} sedang terbuka di Excel. Menyimpan ke: {save_dest} ...", flush=True)
        with pd.ExcelWriter(save_dest, engine="openpyxl") as writer:
            df.to_excel(writer, sheet_name="Output_Matching", index=False)
        print(f"  [OK] Tersimpan di: {save_dest} (Sheet: Output_Matching)", flush=True)

    # Summary
    elapsed = time.perf_counter() - t_start
    print(f"\n{'=' * 64}")
    print(f"  Selesai dalam {elapsed:.1f}s ({elapsed/60:.1f} menit)")
    print(f"  Rows  : {len(df):,}  |  Cols: {len(df.columns)}  |  File: {save_dest}")

    # Hitung ringkasan hasil matching
    flags = df["FLAG_PROD"].value_counts()
    n_new = flags.get("New Entry sebagian", 0) + flags.get("New Entry total", 0)
    n_adj = (
        flags.get("Adjustment sebagian dengan akumulasi", 0) +
        flags.get("Adjustment total dengan akumulasi", 0) +
        flags.get("Adjustment sebagian tanpa akumulasi", 0) +
        flags.get("Adjustment total tanpa akumulasi", 0)
    )
    n_matched = n_new + n_adj
    pct_matched = (n_matched / len(df)) * 100 if len(df) > 0 else 0
    pct_str = f"{pct_matched:.2f}".replace(".", ",")
    total_str = f"{len(df):,}".replace(",", ".")
    matched_str = f"{n_matched:,}".replace(",", ".")

    print(f"\n{'-' * 60}")
    print("Hasil Matching BRINS")
    print(f"New entry : {n_new}")
    print(f"Adjustment : {n_adj}")
    print(f"Dari {total_str} data ada {pct_str}% ({matched_str} data) yang masuk ke dalam New Entry dan Adjustment.")
    print("\nCatatan :")
    print("- ")
    print(f"{'-' * 60}")
    print(f"\n  SUMMARY FLAG_PROD:")
    summary_flag = pd.DataFrame(df["FLAG_PROD"].value_counts()).reset_index()
    summary_flag.columns = ["FLAG_PROD", "JUMLAH"]
    for _, row in summary_flag.iterrows():
        print(f"    {row['FLAG_PROD']:<45}: {row['JUMLAH']:,}")
    print(f"\n  SKENARIO:")
    for sce, count in df["SKENARIO"].value_counts().items():
        print(f"    {sce:<50}: {count:,}")
    print("=" * 64)


if __name__ == "__main__":
    run()

