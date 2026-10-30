import os
import re
import sys
import time
import functools

sys.modules['numexpr'] = None
sys.modules['bottleneck'] = None

import numpy as np
import pandas as pd


# =============================================================================
# KONFIGURASI  (CEDANT: JASA RP -- PT Jasa Raharja Putra)
#
# Sumber arsitektur : matching_tpi.py, prod_sus_1_aca.py, prod_sus_1_astra.py
# Referensi         : matching_TPI.md, teknis_matching_aca.md
# Dokumen workflow  : Workflow Teknis Suspend 2026.docx
#
# Standar Arsitektur Terintegrasi:
#   1. Ingestion Teks Murni (dtype=str) agar data normal 18-digit tidak berubah
#      menjadi notasi ilmiah (e+), serta data error E+ berjalan apa adanya.
#   2. Stopwords: Menambahkan KSO, menghapus nama cedant (JASA, RAHARJA, PUTRA).
#   3. FASE 2 Fallback Pool (Solusi C) untuk menyelamatkan slip payung.
#   4. FASE 3 Hierarki Insured: Step 3A Exact All -> Step 3B Like All (>= 65%)
#      dengan proteksi asimetris strict (_is_insured_match_strict).
#   5. FASE 4 Toleransi Periode (RECEIPT DATE >= FAC_COM_DATE) & Currency.
#   6. FASE 5 Narrowing Bertingkat Non-Destruktif:
#      - Step 1: Slip -> Polis + Slip
#      - Step 2: Sertifikat Exact/In-Range -> + Cert
#      - Step 3: Insured Strict -> + Insured
#   7. Finansial & Standarisasi: round(2) & abs(diff) < 0.5, suffix (> 1 Fac),
#      serta Defensive Export (PermissionError Safe) dengan tanggal _%d%b.
# =============================================================================

_BASE_DIR   = os.path.dirname(os.path.abspath(__file__))   # folder Jasa_Rp/
_CLEAN_DIR  = os.path.join(_BASE_DIR, "data_clean")         # hasil cleansing
_RAW_DIR    = os.path.join(_BASE_DIR, "data_mentah")        # data mentah
_OUT_DIR    = os.path.join(_BASE_DIR, "output matching")    # folder output
def _find_file_in(directory: str, basenames: list) -> str:
    """Cari file pertama yang ada dari daftar kandidat nama dalam direktori tertentu."""
    for name in basenames:
        p = os.path.join(directory, name)
        if os.path.exists(p):
            return p
    return os.path.join(directory, basenames[0])



# [JASA RP] Path file menggunakan suffix/nama Jasa RP
SUSPEND_FILE = _find_file_in(_CLEAN_DIR, [
    "JasaRp_output_suspend.xlsx",
    "jasarp_output_suspend.xlsx",
    "jasa_rp_output_suspend.xlsx",
])
OSBAL_FILE = _find_file_in(_CLEAN_DIR, [
    "JasaRp_output_osbal.xlsx",
    "jasarp_output_osbal.xlsx",
    "jasa_rp_output_osbal.xlsx",
])
FACUL_FILE = _find_file_in(_CLEAN_DIR, [
    "JasaRp_output_facul.xlsx",
    "jasarp_output_facul.xlsx",
    "jasa_rp_output_facul.xlsx",
])
SLIPDB_FILE = _find_file_in(_RAW_DIR, [
    "5. Database Rekap RI Slip.xlsx",
    "5.Database Rekap RI Slip.xlsx",
    "Database Rekap RI Slip.xlsx",
])

# Pastikan folder output ada
os.makedirs(_OUT_DIR, exist_ok=True)
OUTPUT_FILE = os.path.join(_OUT_DIR, "matching_jasarp.xlsx")

# Kolom FAC CODE di masing-masing tabel referensi
OSBAL_FACODE_COL  = "CCOS_REF_CODE"
FACUL_FACODE_COL  = "FAC_CODE"
SLIPDB_FACODE_COL = "FAC_CODE"   # dapat di-override runtime sesuai header aktual

# [JASA RP] Kolom polis/slip/insured SUSPEND
CLSDT_POLIS_COL = "POLIS_ORI"
CLSDT_SLIP_COL  = "SLIP_NO_ORI"
FAC_POLIS_COL   = "FAC_POLICY_NO"
FAC_SLIP_COL    = "FAC_SLIP"

# [JASA RP] Nama kolom currency/periode
SUSPEND_CURR_COL = "CURR ORI"
OSBAL_CURR_COL   = "CCOS_CURR"
SUSPEND_DATE_COL = "RECEIPT DATE"
OSBAL_DATE_COL   = "FAC_COM_DATE"

# Marker baris administratif yang dikecualikan dari matching
_EXCLUDED_REF_MARKERS = ["HUTANG PIUTANG", "DATA SUSPENSE"]

# Marker LINESLIP
_LINESLIP_MARKERS = ("LINESLIP", "LINE SLIP")

_MIN_TOKEN_LEN = 3
_TOKEN_RE      = re.compile(r"[^A-Z0-9]+")
_DELIMITER_RE  = re.compile(r"[,;|+\s]+")
_ENTITY_SPLIT_RE = re.compile(r"\s*(?:QQ|/|,|;|\|)\s*", re.IGNORECASE)

# Token generik yang dikecualikan saat menghitung overlap INSURED
_INSURED_STOPWORDS: frozenset = frozenset({
    "PT", "CV", "TBK", "UD", "PD", "NV", "LLC", "LTD", "INC", "KSO",
    "GROUP", "GROUPS",
    "INDONESIA", "INDO",
    "INDUSTRY", "INDUSTRIES", "INDUSTRI",
    "INTERNATIONAL", "INTERNASIONAL",
    "COMPANY", "COMPANIES", "CORP", "CORPORATION",
    "ENTERPRISE", "ENTERPRISES",
    "TRADING", "TRADE",
    "UTAMA", "MANDIRI", "MAKMUR", "SEJAHTERA",
})

# Overlap coefficient threshold
_INSURED_LIKE_THRESHOLD: float = 0.65
_POLIS_LIKE_THRESHOLD: float   = 0.80

# [JASA RP] Kolom join saat akumulasi suspend
_SUSPEND_JOIN_COLS = [
    "RECEIPT NO", "CREDIT NOTES", "DETAIL RINCIAN NO", "RECEIPT DATE",
    "CEDANT NAME", "CEDANT SHRT NAME",
    "INSURED_ORI", "INSURED_1", "INSURED_2", "INSURED_3",
    "CURR ORI", "CURR PAY", "AMOUNT PAY",
    "POLIS_ORI", "POLIS_CLEAN_1",
    "SLIP_NO_ORI", "SLIP_CLEAN_1",
    "DESC 1", "DESC 2", "DESC 3", "DESC 4",
    "STATUS", "REC_TYPE", "CEK FASE",
]
_SUSPEND_SUM_COLS = ["AMOUNT ORI"]

# [JASA RP] FINAL_COLUMNS: kolom standar output resmi
FINAL_COLUMNS = [
    "CCOS_DOC_NO", "CCOS_REF_CODE",
    "RECEIPT NO", "CREDIT NOTES", "DETAIL RINCIAN NO", "RECEIPT DATE",
    "CEDANT NAME", "CEDANT SHRT NAME",
    "INSURED_ORI", "INSURED_1", "INSURED_2", "INSURED_3",
    "CURR ORI", "AMOUNT ORI", "AMOUNT_ORI_MIN1",
    "CURR PAY", "AMOUNT PAY",
    "CCOS_OR_BAL", "CCOS_BAL_DUE", "DIFERENCE",
    "FLAG_PROD",
    "POLIS",
    "POLICY_CLEAN_1", "SERTIF_CLEAN_1",
    "POLICY_CLEAN_2",
    "POLICY_CLEAN_3",
    "POLICY_CLEAN_4",
    "SLIP_NO", "SLIP_NO_CLN",
    "DESC 1", "DESC 2", "DESC 3", "DESC 4",
    "STATUS", "REC_TYPE", "CEK FASE",
    "SKENARIO",
]


# =============================================================================
# STRING UTILITIES & STRICT INSURED LOGIC
# =============================================================================

@functools.lru_cache(maxsize=131072)
def _norm_cached(s: str) -> str:
    text = s.strip().upper().replace("S/D", "SD")
    return re.sub(r"\s+", " ", text) if "  " in text else text


def _normalize(value) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    return _norm_cached(str(value))


def _clean_tokens_for_insured(s: str) -> list:
    """Bersihkan string nama insured dan tokenize menjadi list kata signifikan."""
    if not s:
        return []
    return [
        w for w in re.split(r'[\s\-_/,\.]+', str(s).upper().strip())
        if w and w not in _INSURED_STOPWORDS and len(w) >= 2
    ]


def _is_insured_match_strict(sus_str: str, ref_str: str) -> bool:
    """Cek kecocokan insured dengan aturan asimetris ketat (mengadopsi TPI):
    - Jika suspend >= 2 kata signifikan:
      * Reference DILARANG match jika hanya 1 kata (holding/entitas induk).
      * Reference harus cocok 2 kata awal identik / progresif (>= 4 char), ATAU
        memiliki irisan minimal 2 kata signifikan identik.
    - Jika suspend 1 kata: kata harus identik persis.
    """
    s = str(sus_str).upper().strip()
    r = str(ref_str).upper().strip()
    if not s or not r:
        return False
    if s == r or _normalize(s) == _normalize(r):
        return True

    st = _clean_tokens_for_insured(s)
    rt = _clean_tokens_for_insured(r)
    if not st or not rt:
        return False

    if len(st) >= 2:
        if len(rt) < 2:
            return False
        w1 = (st[0] == rt[0]) or (len(st[0]) >= 4 and (st[0].startswith(rt[0]) or rt[0].startswith(st[0])))
        w2 = (len(st) > 1 and len(rt) > 1) and ((st[1] == rt[1]) or (len(st[1]) >= 4 and (st[1].startswith(rt[1]) or rt[1].startswith(st[1]))))
        if w1 and w2:
            return True
        return len(set(st) & set(rt)) >= 2
    else:
        return st[0] == rt[0]


def _clean_cert_str(val) -> str:
    """Membersihkan nomor sertifikat: hapus format desimal .0 dan leading zeros."""
    if val is None or pd.isna(val):
        return ""
    s = str(val).strip()
    if s.lower() in ("nan", "none", "-", ""):
        return ""
    s = re.sub(r'\.0+$', '', s)
    s_digits = re.sub(r'^[0\s]+', '', s)
    return s_digits if s_digits else "0"


def _cert_in_range(query_cert: str, ref_cert_str: str) -> bool:
    """Mengecek apakah nomor sertifikat suspend (query_cert) cocok dengan referensi (ref_cert_str).
    Mendukung:
      1. Nilai tunggal exact: '000018' == '18'
      2. Nilai dipisah koma/titik-koma/slash: '000146, 000147' atau '000228, 000230, 000232'
      3. Rentang SD / S/D / s.d. / TO / strip: '000228 SD 000232' (228 s/d 232)
      4. Kombinasi koma dan rentang: '000148, 000169 SD 000198'
    """
    q_clean = _clean_cert_str(query_cert)
    if not q_clean or not q_clean.isdigit():
        return False
    q_num = int(q_clean)

    s_ref = str(ref_cert_str or "").strip()
    if not s_ref or s_ref.lower() in ("nan", "none", "-"):
        return False

    # Normalisasi 'S/D' ke 'SD' sebelum pemisahan koma/delimiter
    s_norm = re.sub(r'\s*S/D\s*', ' SD ', s_ref, flags=re.IGNORECASE)
    segments = re.split(r'[,;|\n]+', s_norm)

    for seg in segments:
        seg = seg.strip()
        if not seg:
            continue

        # Cek apakah segmen merupakan rentang SD (misal: '000228 SD 000232' atau '100 - 200' atau '100 TO 200')
        m_range = re.search(r'(\d+)\s*(?:SD|S\.D\.|\bTO\b|-)\s*(\d+)', seg, re.IGNORECASE)
        if m_range:
            try:
                start_c = int(_clean_cert_str(m_range.group(1)))
                end_c   = int(_clean_cert_str(m_range.group(2)))
                if min(start_c, end_c) <= q_num <= max(start_c, end_c):
                    return True
            except (ValueError, TypeError):
                pass
            continue

        # Cek apakah segmen merupakan angka tunggal exact
        c_clean = _clean_cert_str(seg)
        if c_clean.isdigit():
            if int(c_clean) == q_num:
                return True

    return False

def _format_sertif(val) -> str:
    if pd.isna(val) or str(val).strip() == "":
        return ""
    if isinstance(val, (int, float)):
        return str(int(val)).zfill(6)
    s = str(val).strip()
    if s.endswith('.0'):
        s = s[:-2]
    return s.zfill(6) if s.isdigit() else s


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
    return [c for c in columns if str(c).lower().startswith(prefix.lower())]


def _clean_base_match(scenario: str) -> str:
    """Bersihkan suffix tambahan dari skenario dasar."""
    s = str(scenario or "").strip()
    for to_remove in [
        " (Beda Currency)", " (Beda Periode)", " (> 1 Fac)", " (Matching >1 fac code)",
        " (> 1 fac code)",
    ]:
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
    is_beda_curr: bool,
    is_beda_per: bool,
    num_fac: int,
) -> str:
    """Format skenario final sesuai standar (> 1 Fac, Beda Periode, Beda Currency)."""
    if not source or not base_scenario or str(base_scenario).strip().upper() in ("UNMATCHING", "UNMATCHED"):
        return "Unmatching"

    base = _clean_base_match(base_scenario)
    if not base or base.upper() in ("UNMATCHING", "UNMATCHED"):
        return "Unmatching"
    if base == "Unmatching":
        return "Unmatching"

    if is_beda_curr:
        return f"{base} (Beda Currency)"
    elif is_beda_per:
        return f"{base} (Beda Periode)"
    elif num_fac > 1:
        return f"{base} (> 1 Fac)"
    else:
        return base


# =============================================================================
# ROW DICT HELPERS
# =============================================================================

def _is_excluded_ref_row(ref_row: dict) -> bool:
    for marker in _EXCLUDED_REF_MARKERS:
        for val in ref_row.values():
            if val and marker in str(val).upper():
                return True
    return False


def _is_lineslip_row(suspend_row: dict, insured_cols: list) -> bool:
    for col in insured_cols:
        val = str(suspend_row.get(col, "")).upper()
        if any(marker in val for marker in _LINESLIP_MARKERS):
            return True
    return False


def _collect_clean_values(row_dict: dict, cols: list) -> list:
    results, seen = [], set()
    for col in cols:
        val = _normalize(row_dict.get(col, ""))
        if not val or val in seen:
            continue
        seen.add(val)
        results.append(val)
        for tok in _split_composite(val):
            tok_norm = _normalize(tok)
            if tok_norm and tok_norm not in seen:
                seen.add(tok_norm)
                results.append(tok_norm)
    return results


def _collect_all_values(row_dict: dict, clean_cols: list, ori_col: str) -> list:
    vals = _collect_clean_values(row_dict, clean_cols)
    ori_val = _normalize(row_dict.get(ori_col, ""))
    if ori_val and ori_val not in vals:
        vals.append(ori_val)
        for tok in _split_composite(ori_val):
            tok_norm = _normalize(tok)
            if tok_norm and tok_norm not in vals:
                vals.append(tok_norm)
    return vals


def _collect_insured_values(row_dict: dict, clean_cols: list, ori_col: str = None) -> list:
    """Kumpulkan frasa/entitas insured utuh tanpa memecah kata per-spasi."""
    results, seen = [], set()
    for col in clean_cols:
        val = _normalize(row_dict.get(col, ""))
        if not val or val in seen:
            continue
        seen.add(val)
        results.append(val)
        for ent in _ENTITY_SPLIT_RE.split(val):
            ent_norm = _normalize(ent)
            if ent_norm and ent_norm not in seen:
                seen.add(ent_norm)
                results.append(ent_norm)
    if ori_col:
        ori_val = _normalize(row_dict.get(ori_col, ""))
        if ori_val and ori_val not in seen:
            seen.add(ori_val)
            results.append(ori_val)
            for ent in _ENTITY_SPLIT_RE.split(ori_val):
                ent_norm = _normalize(ent)
                if ent_norm and ent_norm not in seen:
                    seen.add(ent_norm)
                    results.append(ent_norm)
    return results


def _get_effective_sus_cols(sus_cols: list, is_lineslip: bool, prefix: str) -> list:
    if not is_lineslip:
        return [c for c in sus_cols if not c.lower().endswith("_min1")]
    min1_cols = [c for c in sus_cols if c.lower().endswith("_min1")]
    return min1_cols if min1_cols else sus_cols


def _ref_has_value(ref_row: dict, ref_cols: list, ori_col: str, query_values: list, is_insured: bool = False) -> bool:
    qset = set(query_values)
    for col in ref_cols:
        val = _normalize(ref_row.get(col, ""))
        if not val:
            continue
        if val in qset:
            return True
        if is_insured:
            for ent in _ENTITY_SPLIT_RE.split(val):
                if _normalize(ent) in qset:
                    return True
        else:
            for tok in _split_composite(val):
                if _normalize(tok) in qset:
                    return True
    if ori_col:
        ori_val = _normalize(ref_row.get(ori_col, ""))
        if ori_val:
            if ori_val in qset:
                return True
            if is_insured:
                for ent in _ENTITY_SPLIT_RE.split(ori_val):
                    if _normalize(ent) in qset:
                        return True
            else:
                for tok in _split_composite(ori_val):
                    if _normalize(tok) in qset:
                        return True
    return False


def _slip_in_range(slip_val: str, osbal_row: dict) -> bool:
    if not slip_val or not slip_val.isdigit():
        return False
    slip_num = int(slip_val)
    for col, val in osbal_row.items():
        if not val or not isinstance(val, str):
            continue
        col_upper = col.upper()
        if "SLIP" not in col_upper and "DESC" not in col_upper:
            continue
        val_upper = val.upper()
        if "SD" not in val_upper and "S/D" not in val_upper:
            continue
        parts = re.split(r"\s*(?:SD|S/D)\s*", val_upper)
        if len(parts) == 2:
            left  = re.sub(r"[^0-9]", "", parts[0])
            right = re.sub(r"[^0-9]", "", parts[1])
            if left.isdigit() and right.isdigit():
                low, high = int(left), int(right)
                if low > high:
                    low, high = high, low
                if low <= slip_num <= high:
                    return True
    return False


# =============================================================================
# INVERTED INDEX BUILDERS
# =============================================================================

def _build_exact_index(rows: list, cols: list, is_insured: bool = False) -> dict:
    index = {}
    for i, r in enumerate(rows):
        if _is_excluded_ref_row(r):
            continue
        for col in cols:
            val = _normalize(r.get(col, ""))
            if not val:
                continue
            index.setdefault(val, []).append(i)
            if is_insured:
                for ent in _ENTITY_SPLIT_RE.split(val):
                    ent_norm = _normalize(ent)
                    if ent_norm and ent_norm != val:
                        index.setdefault(ent_norm, []).append(i)
            else:
                for tok in _split_composite(val):
                    tok_norm = _normalize(tok)
                    if tok_norm:
                        index.setdefault(tok_norm, []).append(i)
    return index


def _build_facode_index(rows: list, facode_col: str) -> dict:
    index = {}
    for i, r in enumerate(rows):
        fac = _normalize(r.get(facode_col, ""))
        if fac:
            index.setdefault(fac, []).append(i)
    return index


def _build_token_index(rows: list, cols: list, stopwords: frozenset = frozenset()) -> tuple:
    index = {}
    row_tokens = []
    for i, r in enumerate(rows):
        if _is_excluded_ref_row(r):
            row_tokens.append(frozenset())
            continue
        tokens = set()
        for col in cols:
            val = r.get(col, "")
            if val and not (isinstance(val, float) and pd.isna(val)):
                tokens.update(_tokenize(str(val)))
        tokens -= stopwords
        frozen = frozenset(tokens)
        row_tokens.append(frozen)
        for tok in frozen:
            index.setdefault(tok, []).append(i)
    return index, row_tokens


def _build_lookup(rows: list, polis_cols: list, slip_cols: list, insured_cols: list,
                  facode_col: str, label: str = "") -> tuple:
    t0 = time.perf_counter()
    exact_slip    = _build_exact_index(rows, slip_cols, is_insured=False)
    exact_polis   = _build_exact_index(rows, polis_cols, is_insured=False)
    exact_insured = _build_exact_index(rows, insured_cols, is_insured=True)
    facode_idx    = _build_facode_index(rows, facode_col)
    tok_polis     = _build_token_index(rows, polis_cols)
    tok_slip      = _build_token_index(rows, slip_cols)
    tok_insured   = _build_token_index(rows, insured_cols, stopwords=_INSURED_STOPWORDS)

    slip_osbal_idx = exact_slip
    polis_osbal_idx = exact_polis
    insured_osbal_idx = exact_insured

    elapsed = time.perf_counter() - t0
    print(f"  [{label}] Index selesai dalam {elapsed:.2f}s | "
          f"Polis: {len(exact_polis):,} keys | Slip: {len(exact_slip):,} keys | "
          f"Insured: {len(exact_insured):,} keys", flush=True)

    return (
        exact_slip, exact_polis, exact_insured,
        slip_osbal_idx, polis_osbal_idx, insured_osbal_idx,
        tok_slip, tok_polis, tok_insured,
        facode_idx,
    )


def _ref_matches_insured(ref_row: dict, insured_ref_cols: list, sus_ins_vals: list) -> bool:
    """Cek apakah insured suspend cocok dengan referensi (INSURED_CLEAN atau FAC_INSURED_ORI / QQ / /).
    Menggunakan _is_insured_match_strict agar tidak terjadi false match multi-kata ke holding 1 kata.
    """
    if not sus_ins_vals:
        return False
    # 1. Cek exact lewat kolom clean standar
    if _ref_has_value(ref_row, insured_ref_cols, "insured_ori", sus_ins_vals, is_insured=True):
        return True

    for sus_v in sus_ins_vals:
        if not sus_v:
            continue
        # 2. Cek lewat kolom-kolom INSURED_CLEAN
        for c in insured_ref_cols:
            val = ref_row.get(c, "")
            if val and str(val) != "nan":
                if _is_insured_match_strict(sus_v, str(val)):
                    return True
        # 3. Cek FAC_INSURED_ORI (termasuk subsidiary setelah 'QQ' / 'C/O' / '/')
        fac_ori = ref_row.get("FAC_INSURED_ORI", "")
        if fac_ori and str(fac_ori) != "nan":
            parts = [p.strip() for p in str(fac_ori).replace('/', ' QQ ').replace('C/O', ' QQ ').split('QQ') if p.strip()]
            parts.append(str(fac_ori))
            for p in parts:
                if _is_insured_match_strict(sus_v, p):
                    return True
    return False


def _exact_match(query_values: list, index: dict) -> set:
    return {i for v in query_values if v and v in index for i in index[v]}


def _shared_tokens(toks_a: set, toks_b: set) -> set:
    """Hitung shared tokens dengan toleransi pemotongan karakter (>= 4 char)."""
    shared = set()
    for a in toks_a:
        for b in toks_b:
            if a == b or (len(a) >= 4 and len(b) >= 4 and (a.startswith(b) or b.startswith(a))):
                shared.add(a)
                break
    return shared


def _like_match(
    query_values: list,
    token_index: dict,
    rows: list,
    ref_cols: list,
    stopwords: frozenset = frozenset(),
    overlap_threshold: float = 0.65,
    match_log: list = None,
    min_shared_tokens: int = 1,
) -> set:
    query_tokens = set()
    for q in query_values:
        if q:
            query_tokens.update(_tokenize(q))
    query_tokens -= stopwords
    if not query_tokens:
        return set()

    clean_queries = [_normalize(q) for q in query_values if _normalize(q)]

    # [1] SUBSTR
    substr_hits = set()
    for cq in clean_queries:
        if len(cq) < 4:
            continue
        for idx, r in enumerate(rows):
            if _is_excluded_ref_row(r):
                continue
            for col in ref_cols:
                rv = _normalize(r.get(col, ""))
                if rv and len(rv) >= 4:
                    if cq in rv or rv in cq:
                        if min_shared_tokens > 1:
                            cq_toks = set(_tokenize(cq)) - stopwords
                            rv_toks = set(_tokenize(rv)) - stopwords
                            if len(_shared_tokens(cq_toks, rv_toks)) < min_shared_tokens:
                                continue
                        substr_hits.add(idx)
                        break
    if substr_hits:
        if match_log is not None:
            match_log.append((substr_hits, "SUBSTR"))
        return substr_hits

    # [2] SUBSET
    candidate_rows: set = set()
    for tok in query_tokens:
        if tok in token_index:
            candidate_rows.update(token_index[tok])
    if not candidate_rows:
        return set()

    subset_hits = set()
    for idx in candidate_rows:
        row_toks = set()
        for col in ref_cols:
            val = rows[idx].get(col, "")
            if val and not (isinstance(val, float) and pd.isna(val)):
                row_toks.update(_tokenize(str(val)))
        row_toks -= stopwords
        if not row_toks:
            continue
        shared = _shared_tokens(query_tokens, row_toks)
        if len(shared) < min_shared_tokens:
            continue
        if query_tokens.issubset(row_toks) or row_toks.issubset(query_tokens):
            subset_hits.add(idx)

    if subset_hits:
        if match_log is not None:
            match_log.append((subset_hits, "SUBSET"))
        return subset_hits

    # [3] OVERLAP
    overlap_hits = set()
    for idx in candidate_rows:
        row_toks = set()
        for col in ref_cols:
            val = rows[idx].get(col, "")
            if val and not (isinstance(val, float) and pd.isna(val)):
                row_toks.update(_tokenize(str(val)))
        row_toks -= stopwords
        if not row_toks:
            continue
        shared = _shared_tokens(query_tokens, row_toks)
        if len(shared) < min_shared_tokens:
            continue
        intersection = len(shared)
        overlap = intersection / min(len(query_tokens), len(row_toks))
        if overlap >= overlap_threshold:
            overlap_hits.add(idx)

    if overlap_hits:
        if match_log is not None:
            match_log.append((overlap_hits, "OVERLAP"))
        return overlap_hits

    return set()


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
    """Rematch ketat dengan filter currency + periode (R1-R5 standar)."""
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
    like_vals = _collect_clean_values(suspend_row, like_cols)
    if lookup_osbal[7] and lookup_osbal[7][0] and like_vals:
        like_hits = _strict_filter(_like_match(
            like_vals, lookup_osbal[7][0], osbal_rows, polis_ref_cols,
            stopwords=frozenset(), overlap_threshold=_POLIS_LIKE_THRESHOLD,
        ))
        if like_hits:
            return "OSBAL", "Polis only", like_hits, True

    # R3: SLIPDB polis exact
    if slipdb_rows and lookup_slipdb:
        sld_hits = _exact_match(polis_vals, lookup_slipdb[1])
        if sld_hits:
            res_idx, ok = _resolve_facode([slipdb_rows[i] for i in sld_hits], SLIPDB_FACODE_COL, facode_osbal_idx, osbal_rows)
            curr_res = _strict_filter(res_idx)
            if curr_res:
                return "SLIPDB", "Polis only", curr_res, True

    # R4: FACUL polis exact
    if facul_rows and lookup_facul:
        fac_hits = _exact_match(polis_vals, lookup_facul[1])
        if fac_hits:
            res_idx, ok = _resolve_facode([facul_rows[i] for i in fac_hits], FACUL_FACODE_COL, facode_osbal_idx, osbal_rows)
            curr_res = _strict_filter(res_idx)
            if curr_res:
                return "FACUL", "Polis only", curr_res, True

    # R5: OSBAL slip exact
    slip_vals = _collect_clean_values(suspend_row, slip_sus_cols)
    if slip_vals:
        curr_slip_hits = _strict_filter(_exact_match(slip_vals, lookup_osbal[0]))
        if curr_slip_hits:
            return "OSBAL", "Slip only", curr_slip_hits, True

    return None


def _resolve_facode(matched_ref_rows: list, facode_col: str, facode_osbal_idx: dict, osbal_rows: list) -> tuple:
    osbal_indices = []
    found_any = False
    for r in matched_ref_rows:
        fac = _normalize(r.get(facode_col, ""))
        if fac and fac in facode_osbal_idx:
            osbal_indices.extend(facode_osbal_idx[fac])
            found_any = True
    seen, unique_indices = set(), []
    for i in osbal_indices:
        if i not in seen:
            seen.add(i)
            unique_indices.append(i)
    return unique_indices, found_any


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
        "CCOS_OR_BAL":   sum_col("CCOS_OR_BAL",   osbal_rows),
        "CCOS_BAL_DUE":  sum_col("CCOS_BAL_DUE",  bal_due_rows),
    }


def _build_output_row(
    suspend_row: dict, source: str, scenario: str, osbal_rows: list,
    osbal_count: int = 0, resolved: bool = True,
    suspend_count: int = 1, all_fac_osbal_rows: list = None,
    is_beda_curr: bool = False, is_beda_per: bool = False,
) -> dict:
    """Bangun satu baris output untuk JASA RP."""
    has_match = bool(source and osbal_rows and osbal_rows[0])
    has_osbal = has_match and (source in ("OSBAL", "SLIPDB", "FACUL") or resolved)

    ccos = (
        _aggregate_ccos(osbal_rows, all_fac_osbal_rows=all_fac_osbal_rows)
        if has_osbal
        else {"CCOS_DOC_NO": "", "CCOS_REF_CODE": "", "CCOS_OR_BAL": np.nan, "CCOS_BAL_DUE": np.nan}
    )
    if is_beda_curr or is_beda_per or "Beda Currency" in str(scenario) or "Beda Periode" in str(scenario):
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
        "INSURED_ORI":       suspend_row.get("INSURED_ORI",       ""),
        "INSURED_1":         suspend_row.get("INSURED_1",         ""),
        "INSURED_2":         suspend_row.get("INSURED_2",         ""),
        "INSURED_3":         suspend_row.get("INSURED_3",         ""),
        "CURR ORI":          suspend_row.get("CURR ORI",          ""),
        "AMOUNT ORI":        suspend_row.get("AMOUNT ORI",        ""),
        "CURR PAY":          suspend_row.get("CURR PAY",          ""),
        "AMOUNT PAY":        suspend_row.get("AMOUNT PAY",        ""),
        "CCOS_OR_BAL":       ccos["CCOS_OR_BAL"],
        "CCOS_BAL_DUE":      ccos["CCOS_BAL_DUE"],
        "POLIS":             suspend_row.get("POLIS_ORI",         ""),
        "POLICY_CLEAN_1":    suspend_row.get("POLIS_CLEAN_1",     ""),
        "SERTIF_CLEAN_1":    _format_sertif(suspend_row.get("CERTIFICATE_1", "")),
        "POLICY_CLEAN_2":    suspend_row.get("clean polis 2",     ""),
        "POLICY_CLEAN_3":    suspend_row.get("clean polis 3",     ""),
        "POLICY_CLEAN_4":    suspend_row.get("clean polis 4",     ""),
        "SLIP_NO":           suspend_row.get("SLIP_NO_ORI",       ""),
        "SLIP_NO_CLN":       suspend_row.get("SLIP_CLEAN_1",      ""),
        "DESC 1":            suspend_row.get("DESC 1",            ""),
        "DESC 2":            suspend_row.get("DESC 2",            ""),
        "DESC 3":            suspend_row.get("DESC 3",            ""),
        "DESC 4":            suspend_row.get("DESC 4",            ""),
        "STATUS":            suspend_row.get("STATUS",            ""),
        "REC_TYPE":          suspend_row.get("REC_TYPE",          ""),
        "CEK FASE":          suspend_row.get("CEK FASE",          ""),
        "SKENARIO":          final_scenario,
        "_OSBAL_ROW_COUNT":   osbal_count,
        "_SUSPEND_ROW_COUNT": suspend_count,
    }


def _merge_suspend_rows(suspend_rows_list: list) -> dict:
    base = dict(suspend_rows_list[0])
    for col in _SUSPEND_SUM_COLS:
        total = 0.0
        for r in suspend_rows_list:
            v = r.get(col, 0)
            if pd.notna(v):
                try:
                    total += float(str(v).replace(',', ''))
                except ValueError:
                    pass
        base[col] = total
    for col in _SUSPEND_JOIN_COLS:
        vals = []
        for r in suspend_rows_list:
            v = str(r.get(col, "")).strip()
            if v and v not in vals:
                vals.append(v)
        base[col] = ", ".join(vals)
    return base


def _compute_derived_cols(df: pd.DataFrame) -> pd.DataFrame:
    """Hitung AMOUNT_ORI_MIN1, DIFERENCE, dan FLAG_PROD (10 kategori)."""
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
    amount_ori  = amount_ori_numeric.fillna(0)
    amount_neg  = df["AMOUNT_ORI_MIN1"].fillna(0)

    balance_due     = pd.to_numeric(df.get("CCOS_BAL_DUE",      0), errors="coerce").fillna(0)
    osbal_count_col = pd.to_numeric(df.get("_OSBAL_ROW_COUNT",  0), errors="coerce").fillna(0)
    sus_count_col   = pd.to_numeric(df.get("_SUSPEND_ROW_COUNT",1), errors="coerce").fillna(1)
    ccos_ref        = df.get("CCOS_REF_CODE", pd.Series("", index=df.index)).fillna("").astype(str)
    skenario_col    = df.get("SKENARIO",      pd.Series("", index=df.index)).fillna("")

    df["DIFERENCE"] = amount_neg - balance_due
    accumulated = (osbal_count_col > 1) | (sus_count_col > 1)
    # Toleransi finansial: round(2) dan selisih sen < 0.5 kebal floating point drift
    is_equal    = (amount_neg.round(2) == balance_due.round(2)) | (abs(df["DIFERENCE"]) < 0.5)

    if "_IS_BEDA_PER" in df.columns:
        is_beda_periode = df["_IS_BEDA_PER"].fillna(False).astype(bool)
    else:
        is_beda_periode = skenario_col.str.contains("Beda Periode",  na=False, regex=False)

    if "_IS_BEDA_CURR" in df.columns:
        is_beda_currency = df["_IS_BEDA_CURR"].fillna(False).astype(bool)
    else:
        is_beda_currency = skenario_col.str.contains("Beda Currency", na=False, regex=False)

    is_gt1_fac           = (
        skenario_col.str.contains("> 1 Fac", na=False, regex=False)
        | skenario_col.str.contains("Matching >1 fac code", na=False, regex=False)
        | (ccos_ref.str.contains(",", na=False))
    )
    is_new_entry_total    = (amount_ori != 0) & (balance_due == 0)
    is_adj_total          = is_equal
    is_adj_sebagian       = ~is_equal & (amount_neg < balance_due)
    is_new_entry_sebagian = ~is_equal & (amount_neg > balance_due)

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
        default="Unmatching"
    )

    df.loc[df["SKENARIO"].isin(["UNMATCHED", "Unmatching"]), "FLAG_PROD"] = "Unmatching"
    return df


def _load_excel(path: str, label: str, optional: bool = False) -> tuple:
    """Load Excel ke (list[dict], list[str]) dengan pickle cache otomatis dan pembacaan string murni."""
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
    header_row = 1 if "Slip" in os.path.basename(path) else 0
    try:
        df = pd.read_excel(path, header=header_row, dtype=str)
    except Exception as e:
        print(f"  [RETRY] engine default gagal ({e}), coba openpyxl ...", flush=True)
        df = pd.read_excel(path, engine="openpyxl", header=header_row, dtype=str)

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
    """Pipeline utama JASA RP: matching bertingkat 5 fase standar."""
    t_start = time.perf_counter()
    print("\n" + "=" * 65, flush=True)
    print("  PRODUCTION SCRIPT -- SUSPEND MATCHING  (JASA RP) -- V2", flush=True)
    print("  Cedant: PT Jasa Raharja Putra", flush=True)
    print("=" * 65, flush=True)

    # [1] Load data
    print("\n[1/6] Loading data ...", flush=True)
    suspend_rows, suspend_cols = _load_excel(SUSPEND_FILE, "SUSPEND")
    osbal_rows,   osbal_cols   = _load_excel(OSBAL_FILE,   "OSBAL")
    facul_rows,   facul_cols   = _load_excel(FACUL_FILE,   "FACUL")
    slipdb_rows,  slipdb_cols  = _load_excel(SLIPDB_FILE,  "SLIPDB", optional=True)

    print(f"  Output folder    : {_OUT_DIR}", flush=True)
    print(f"  Output file      : {OUTPUT_FILE}", flush=True)

    # Diagnostik kolom kunci
    print("\n  --- Cek kolom SUSPEND ---", flush=True)
    for col in ["INSURED_ORI", "INSURED_1", "INSURED_2", "INSURED_3",
                "POLIS_ORI", "POLIS_CLEAN_1", "CERTIFICATE_1",
                "SLIP_NO_ORI", "SLIP_CLEAN_1",
                SUSPEND_CURR_COL, SUSPEND_DATE_COL, "CEK FASE"]:
        status = "OK" if col in suspend_cols else "[MISSING]"
        print(f"    {col:<22}: {status}", flush=True)

    print("\n  --- Cek kolom OSBAL ---", flush=True)
    for col in [OSBAL_FACODE_COL, "FAC_POLICY_NO_ORI", "POLIS_CLEAN_1",
                "FAC_SLIP_ORI", "SLIP_CLEAN_1",
                "FAC_INSURED_ORI", "INSURED_CLEAN_1", "CERTIFICATE_1",
                OSBAL_CURR_COL, OSBAL_DATE_COL]:
        status = "OK" if col in osbal_cols else "[MISSING]"
        print(f"    {col:<22}: {status}", flush=True)

    # Pre-parse tanggal
    print("\n  Pre-parsing dates ...", flush=True)
    for r in osbal_rows:
        raw = r.get(OSBAL_DATE_COL, "")
        r["_com_date_parsed"] = pd.to_datetime(raw, errors="coerce")             if raw and not (isinstance(raw, float) and pd.isna(raw)) else None
    for r in suspend_rows:
        raw = r.get(SUSPEND_DATE_COL, "")
        r["_sus_date_parsed"] = pd.to_datetime(raw, errors="coerce")             if raw and not (isinstance(raw, float) and pd.isna(raw)) else None

    # [2] Deteksi kolom bersih
    print("\n[2/6] Detecting clean columns ...", flush=True)

    polis_sus   = _get_clean_cols(suspend_cols, "POLIS_CLEAN")   or _get_clean_cols(suspend_cols, "clean polis")   or []
    slip_sus    = _get_clean_cols(suspend_cols, "SLIP_CLEAN")    or _get_clean_cols(suspend_cols, "clean slip")    or []
    insured_sus = _get_clean_cols(suspend_cols, "INSURED_CLEAN") or _get_clean_cols(suspend_cols, "clean insured") or []
    for extra in ["INSURED_1", "INSURED_2", "INSURED_3"]:
        if extra in suspend_cols and extra not in insured_sus:
            insured_sus.append(extra)

    sertif_sus = [c for c in suspend_cols if "CERTIFICATE" in c.upper() or "SERTIF" in c.upper()]

    print(f"  Kolom polis suspend   : {polis_sus}", flush=True)
    print(f"  Kolom slip suspend    : {slip_sus}", flush=True)
    print(f"  Kolom insured suspend : {insured_sus}", flush=True)
    print(f"  Kolom sertif suspend  : {sertif_sus}", flush=True)

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

    # OSBAL columns
    polis_osbal   = _get_clean_cols(osbal_cols, "POLIS_CLEAN")   or _get_clean_cols(osbal_cols, "clean polis")   or []
    slip_osbal    = _get_clean_cols(osbal_cols, "SLIP_CLEAN")    or _get_clean_cols(osbal_cols, "clean slip")    or []
    insured_osbal = _get_clean_cols(osbal_cols, "INSURED_CLEAN") or _get_clean_cols(osbal_cols, "clean insured") or []
    sertif_osbal  = [c for c in osbal_cols if "CERTIFICATE" in c.upper() or "SERTIF" in c.upper()]

    # FACUL columns
    polis_facul   = _get_clean_cols(facul_cols, "POLIS_CLEAN") or _get_clean_cols(facul_cols, "clean polis") or []
    slip_facul    = []
    for candidate in ["FAC_SLIP", "FAC_SLIP_1", "FAC_SLIP_2", "FAC_SLIP_3", "FAC_SLIP_4", "FAC_SLIP_5"]:
        if candidate in facul_cols and candidate not in slip_facul:
            slip_facul.append(candidate)
    if not slip_facul:
        slip_facul = _get_clean_cols(facul_cols, "SLIP_CLEAN") or _get_clean_cols(facul_cols, "clean slip") or []

    insured_facul = []
    for candidate in ["FAC_INSURED_ORI", "FAC_INSURED_1", "FAC_INSURED_2", "FAC_INSURED_3", "FAC_INSURED_4", "FAC_INSURED_5"]:
        if candidate in facul_cols and candidate not in insured_facul:
            insured_facul.append(candidate)
    if not insured_facul:
        insured_facul = _get_clean_cols(facul_cols, "INSURED_CLEAN") or _get_clean_cols(facul_cols, "clean insured") or []

    sertif_facul = [c for c in facul_cols if "CERTIFICATE" in c.upper() or "SERTIF" in c.upper()]

    # SLIPDB columns
    polis_slipdb   = _get_clean_cols(slipdb_cols, "POLIS_CLEAN")   or _get_clean_cols(slipdb_cols, "clean polis")   or []
    slip_slipdb    = _get_clean_cols(slipdb_cols, "SLIP_CLEAN")    or _get_clean_cols(slipdb_cols, "clean slip")    or []
    insured_slipdb = _get_clean_cols(slipdb_cols, "INSURED_CLEAN") or _get_clean_cols(slipdb_cols, "clean insured") or []

    if not polis_slipdb:
        polis_slipdb = [c for c in ["CLSDT_POLICY_NO", "FAC_POLICY_NO", "POLIS", "POLICY_NO", "POLICY NO"] if c in slipdb_cols]
    if not slip_slipdb:
        slip_slipdb  = [c for c in ["CLSDT_SLIP_NO", "FAC_SLIP", "SLIP", "SLIP_NO", "SLIP NO"] if c in slipdb_cols]
    if not insured_slipdb:
        insured_slipdb = [c for c in ["FAC_INSURED", "INSURED", "CEDANT_NAME", "CEDANT NAME"] if c in slipdb_cols]

    global SLIPDB_FACODE_COL
    if   "FAC_CODE"      in slipdb_cols: SLIPDB_FACODE_COL = "FAC_CODE"
    elif "FAC CODE"      in slipdb_cols: SLIPDB_FACODE_COL = "FAC CODE"
    elif "CCOS_REF_CODE" in slipdb_cols: SLIPDB_FACODE_COL = "CCOS_REF_CODE"

    def _prepend(col, cols_list, header):
        return ([col] + cols_list) if col in header and col not in cols_list else cols_list

    polis_osbal   = _prepend("FAC_POLICY_NO_ORI", polis_osbal,   osbal_cols)
    slip_osbal    = _prepend("FAC_SLIP_ORI",       slip_osbal,    osbal_cols)
    insured_osbal = _prepend("FAC_INSURED_ORI",    insured_osbal, osbal_cols)
    polis_facul   = _prepend("FAC_POLICY_NO",      polis_facul,   facul_cols)
    insured_facul = _prepend("FAC_INSURED_ORI",    insured_facul, facul_cols)
    polis_slipdb  = _prepend("CLSDT_POLICY_NO",    polis_slipdb,  slipdb_cols)
    slip_slipdb   = _prepend("CLSDT_SLIP_NO",      slip_slipdb,   slipdb_cols)

    # [3] Bangun lookup index
    print("\n[3/6] Building lookup indexes ...", flush=True)
    lookup_osbal  = _build_lookup(osbal_rows,  polis_osbal,  slip_osbal,  insured_osbal,  OSBAL_FACODE_COL,  label="OSBAL")
    lookup_facul  = _build_lookup(facul_rows,  polis_facul,  slip_facul,  insured_facul,  FACUL_FACODE_COL,  label="FACUL")
    lookup_slipdb = _build_lookup(slipdb_rows, polis_slipdb, slip_slipdb, insured_slipdb, SLIPDB_FACODE_COL, label="SLIPDB")

    facode_osbal_idx = lookup_osbal[9]

    # [4] Matching per baris
    total = len(suspend_rows)
    print(f"\n[4/6] Matching {total:,} suspend rows ...", flush=True)
    t4 = time.perf_counter()

    raw_results = []
    stat_like_substr  = 0
    stat_like_subset  = 0
    stat_like_overlap = 0

    for idx, sus in enumerate(suspend_rows):
        if (idx + 1) % 200 == 0 or (idx + 1) == total:
            pct = (idx + 1) / total * 100
            print(f"  Processed {idx+1:,}/{total:,} ({pct:.1f}%) ...", flush=True)

        _like_log = []
        _is_ls = sus.get("_is_lineslip", False)
        effective_date = sus.get("_sus_date_lineslip") if _is_ls else sus.get("_sus_date_parsed")

        eff_polis = _get_effective_sus_cols(polis_sus, _is_ls, "POLIS")
        eff_slip  = _get_effective_sus_cols(slip_sus,  _is_ls, "SLIP")
        like_polis = polis_sus
        like_slip  = slip_sus

        matched_osbal_idx = set()
        source   = None
        scenario = None
        resolved = False

        fallback_slip_hits = set()
        fallback_slip_source = None
        fallback_slip_resolved = False

        # =================================================================
        # FASE 1: Pencarian berdasarkan POLIS
        # OSBAL -> SLIPDB -> FACUL -> LIKE match polis
        # =================================================================
        found_by_polis = False
        polis_vals = _collect_clean_values(sus, eff_polis)
        if polis_vals:
            # OSBAL exact
            osb_hits = _exact_match(polis_vals, lookup_osbal[1])
            if osb_hits:
                matched_osbal_idx.update(osb_hits)
                found_by_polis = True

            # SLIPDB exact
            if slipdb_rows:
                sld_hits = _exact_match(polis_vals, lookup_slipdb[1])
                if sld_hits:
                    res_idx, _ = _resolve_facode([slipdb_rows[i] for i in sld_hits], SLIPDB_FACODE_COL, facode_osbal_idx, osbal_rows)
                    if res_idx:
                        matched_osbal_idx.update(res_idx)
                        found_by_polis = True

            # FACUL exact
            if facul_rows:
                fac_hits = _exact_match(polis_vals, lookup_facul[1])
                if fac_hits:
                    res_idx, _ = _resolve_facode([facul_rows[i] for i in fac_hits], FACUL_FACODE_COL, facode_osbal_idx, osbal_rows)
                    if res_idx:
                        matched_osbal_idx.update(res_idx)
                        found_by_polis = True

            # LIKE match polis (jika exact miss)
            if not found_by_polis and lookup_osbal[7] and lookup_osbal[7][0] and like_polis:
                like_vals = _collect_clean_values(sus, like_polis)
                if like_vals:
                    like_hits = _like_match(
                        like_vals, lookup_osbal[7][0], osbal_rows, polis_osbal,
                        stopwords=frozenset(), overlap_threshold=_POLIS_LIKE_THRESHOLD, match_log=_like_log,
                    )
                    if like_hits:
                        matched_osbal_idx.update(like_hits)
                        found_by_polis = True

            if found_by_polis:
                source   = "OSBAL"
                scenario = "Polis only"

        # =================================================================
        # FASE 2: Pencarian berdasarkan SLIP (fallback)
        # Dilengkapi Mekanisme Fallback Pool (Solusi C) untuk Slip Payung
        # =================================================================
        if not found_by_polis:
            slip_vals = _collect_clean_values(sus, eff_slip)
            if slip_vals:
                # 1. OSBAL exact
                osb_hits = _exact_match(slip_vals, lookup_osbal[0])
                if osb_hits:
                    sub_facs = {str(osbal_rows[i].get(OSBAL_FACODE_COL, "")).strip() for i in osb_hits}
                    sus_ins_curr = _collect_insured_values(sus, insured_sus, "insured_ori")
                    sub_ins_confirmed = [i for i in osb_hits if _ref_matches_insured(osbal_rows[i], insured_osbal, sus_ins_curr)]
                    if sub_ins_confirmed:
                        matched_osbal_idx.update(sub_ins_confirmed)
                        source = "OSBAL"
                        scenario = "Slip + Insured"
                    elif len(sub_facs) == 1:
                        matched_osbal_idx.update(osb_hits)
                        source = "OSBAL"
                        scenario = "Slip only"
                    else:
                        # Slip Payung multi-FAC: simpan ke pool fallback
                        fallback_slip_hits = set(osb_hits)
                        fallback_slip_source = "OSBAL"
                        fallback_slip_resolved = False

                # 2. SLIPDB exact
                if not matched_osbal_idx and slipdb_rows:
                    sld_hits = _exact_match(slip_vals, lookup_slipdb[0])
                    if sld_hits:
                        sus_ins_curr = _collect_insured_values(sus, insured_sus, "insured_ori")
                        sld_ins_confirmed = [i for i in sld_hits if _ref_matches_insured(slipdb_rows[i], insured_slipdb, sus_ins_curr)]
                        sub_sld_facs = {str(slipdb_rows[i].get(SLIPDB_FACODE_COL, "")).strip() for i in sld_hits}
                        if sld_ins_confirmed:
                            res_idx, ok = _resolve_facode([slipdb_rows[i] for i in sld_ins_confirmed], SLIPDB_FACODE_COL, facode_osbal_idx, osbal_rows)
                            if ok:
                                matched_osbal_idx.update(res_idx)
                                source = "SLIPDB"
                                scenario = "Slip + Insured"
                                resolved = True
                        else:
                            res_idx, ok = _resolve_facode([slipdb_rows[i] for i in sld_hits], SLIPDB_FACODE_COL, facode_osbal_idx, osbal_rows)
                            if ok:
                                sub_facs = {str(osbal_rows[x].get(OSBAL_FACODE_COL, "")).strip() for x in res_idx}
                                if len(sub_facs) == 1:
                                    matched_osbal_idx.update(res_idx)
                                    source = "SLIPDB"
                                    scenario = "Slip only"
                                    resolved = True
                                else:
                                    fallback_slip_hits = set(res_idx)
                                    fallback_slip_source = "SLIPDB"
                                    fallback_slip_resolved = True

                # 3. FACUL exact
                if not matched_osbal_idx and facul_rows:
                    fac_hits = _exact_match(slip_vals, lookup_facul[0])
                    if fac_hits:
                        sus_ins_curr = _collect_insured_values(sus, insured_sus, "insured_ori")
                        fac_ins_confirmed = [i for i in fac_hits if _ref_matches_insured(facul_rows[i], insured_facul, sus_ins_curr)]
                        sub_fac_facs = {str(facul_rows[i].get(FACUL_FACODE_COL, "")).strip() for i in fac_hits}
                        if fac_ins_confirmed:
                            res_idx, ok = _resolve_facode([facul_rows[i] for i in fac_ins_confirmed], FACUL_FACODE_COL, facode_osbal_idx, osbal_rows)
                            if ok:
                                matched_osbal_idx.update(res_idx)
                                source = "FACUL"
                                scenario = "Slip + Insured"
                                resolved = True
                        else:
                            res_idx, ok = _resolve_facode([facul_rows[i] for i in fac_hits], FACUL_FACODE_COL, facode_osbal_idx, osbal_rows)
                            if ok:
                                sub_facs = {str(osbal_rows[x].get(OSBAL_FACODE_COL, "")).strip() for x in res_idx}
                                if len(sub_facs) == 1:
                                    matched_osbal_idx.update(res_idx)
                                    source = "FACUL"
                                    scenario = "Slip only"
                                    resolved = True
                                else:
                                    fallback_slip_hits = set(res_idx)
                                    fallback_slip_source = "FACUL"
                                    fallback_slip_resolved = True

                # 4. LIKE match slip (fallback)
                if not matched_osbal_idx and not fallback_slip_hits and not osb_hits and lookup_osbal[6] and lookup_osbal[6][0] and like_slip:
                    like_slip_vals = _collect_clean_values(sus, like_slip)
                    if like_slip_vals:
                        like_hits = _like_match(
                            like_slip_vals, lookup_osbal[6][0], osbal_rows, slip_osbal,
                            stopwords=frozenset(), overlap_threshold=_POLIS_LIKE_THRESHOLD, match_log=_like_log,
                        )
                        if like_hits:
                            sus_ins_curr = _collect_insured_values(sus, insured_sus, "insured_ori")
                            sub_ins_confirmed = [i for i in like_hits if _ref_matches_insured(osbal_rows[i], insured_osbal, sus_ins_curr)]
                            if sub_ins_confirmed:
                                matched_osbal_idx.update(sub_ins_confirmed)
                                source = "OSBAL"
                                scenario = "Slip + Insured"
                            else:
                                sub_facs = {str(osbal_rows[i].get(OSBAL_FACODE_COL, "")).strip() for i in like_hits}
                                if len(sub_facs) == 1:
                                    matched_osbal_idx.update(like_hits)
                                    source = "OSBAL"
                                    scenario = "Slip only"
                                else:
                                    fallback_slip_hits = set(like_hits)
                                    fallback_slip_source = "OSBAL"
                                    fallback_slip_resolved = False

        # =================================================================
        # FASE 3: Pencarian berdasarkan INSURED (fallback)
        # Prioritas: Step 3A Exact All -> Step 3B Strict Like All
        # =================================================================
        if not matched_osbal_idx:
            ins_vals = _collect_insured_values(sus, insured_sus, "insured_ori")
            if ins_vals:
                slip_check_vals = _collect_clean_values(sus, eff_slip)

                # --- STEP 3A: EXACT MATCH DULU DI SEMUA SUMBER ---
                # 1. OSBAL Exact
                osb_exact = _exact_match(ins_vals, lookup_osbal[2])
                if osb_exact:
                    matched_osbal_idx.update(osb_exact)
                    source = "OSBAL"
                    osb_slip_conf = any(
                        any(_normalize(osbal_rows[i].get(c, "")) in slip_check_vals for c in slip_osbal if _normalize(osbal_rows[i].get(c, "")))
                        for i in osb_exact
                    ) if slip_check_vals else False
                    scenario = "Slip + Insured" if osb_slip_conf else "Insured only"

                # 2. SLIPDB Exact
                elif slipdb_rows and not matched_osbal_idx:
                    sld_exact = _exact_match(ins_vals, lookup_slipdb[2])
                    if sld_exact:
                        res_idx, _ = _resolve_facode([slipdb_rows[i] for i in sld_exact], SLIPDB_FACODE_COL, facode_osbal_idx, osbal_rows)
                        if res_idx:
                            matched_osbal_idx.update(res_idx)
                            source = "SLIPDB"
                            resolved = True
                            sld_slip_conf = any(
                                any(_normalize(slipdb_rows[i].get(c, "")) in slip_check_vals for c in slip_slipdb if _normalize(slipdb_rows[i].get(c, "")))
                                for i in sld_exact
                            ) if slip_check_vals else False
                            scenario = "Slip + Insured" if sld_slip_conf else "Insured only"

                # 3. FACUL Exact
                elif facul_rows and not matched_osbal_idx:
                    fac_exact = _exact_match(ins_vals, lookup_facul[2])
                    if fac_exact:
                        res_idx, _ = _resolve_facode([facul_rows[i] for i in fac_exact], FACUL_FACODE_COL, facode_osbal_idx, osbal_rows)
                        if res_idx:
                            matched_osbal_idx.update(res_idx)
                            source = "FACUL"
                            resolved = True
                            fac_slip_conf = any(
                                any(_normalize(facul_rows[i].get(c, "")) in slip_check_vals for c in slip_facul if _normalize(facul_rows[i].get(c, "")))
                                for i in fac_hits
                            ) if slip_check_vals else False
                            scenario = "Slip + Insured" if fac_slip_conf else "Insured only"

                # --- STEP 3B: JIKA SEMUA EXACT NIHIL, BARU JALANKAN STRICT LIKE MATCH ---
                if not matched_osbal_idx:
                    # 1. OSBAL Like
                    if lookup_osbal[8] and lookup_osbal[8][0]:
                        osb_like = _like_match(
                            ins_vals, lookup_osbal[8][0], osbal_rows, insured_osbal,
                            stopwords=_INSURED_STOPWORDS, overlap_threshold=_INSURED_LIKE_THRESHOLD, match_log=_like_log,
                            min_shared_tokens=2,
                        )
                        if osb_like:
                            matched_osbal_idx.update(osb_like)
                            source = "OSBAL"
                            scenario = "Insured only"

                    # 2. SLIPDB Like
                    if not matched_osbal_idx and slipdb_rows and lookup_slipdb[8] and lookup_slipdb[8][0]:
                        sld_like = _like_match(
                            ins_vals, lookup_slipdb[8][0], slipdb_rows, insured_slipdb,
                            stopwords=_INSURED_STOPWORDS, overlap_threshold=_INSURED_LIKE_THRESHOLD, match_log=_like_log,
                            min_shared_tokens=2,
                        )
                        if sld_like:
                            res_idx, _ = _resolve_facode([slipdb_rows[i] for i in sld_like], SLIPDB_FACODE_COL, facode_osbal_idx, osbal_rows)
                            if res_idx:
                                matched_osbal_idx.update(res_idx)
                                source = "SLIPDB"
                                scenario = "Insured only"
                                resolved = True

                    # 3. FACUL Like
                    if not matched_osbal_idx and facul_rows and lookup_facul[8] and lookup_facul[8][0]:
                        fac_like = _like_match(
                            ins_vals, lookup_facul[8][0], facul_rows, insured_facul,
                            stopwords=_INSURED_STOPWORDS, overlap_threshold=_INSURED_LIKE_THRESHOLD, match_log=_like_log,
                            min_shared_tokens=2,
                        )
                        if fac_like:
                            res_idx, _ = _resolve_facode([facul_rows[i] for i in fac_like], FACUL_FACODE_COL, facode_osbal_idx, osbal_rows)
                            if res_idx:
                                matched_osbal_idx.update(res_idx)
                                source = "FACUL"
                                scenario = "Insured only"
                                resolved = True

        # Fallback Slip: jika FASE 3 (Insured) tidak menghasilkan match,
        # gunakan kandidat slip FASE 2 agar tidak menjadi Unmatching (termasuk slip payung)
        if not matched_osbal_idx and fallback_slip_hits:
            matched_osbal_idx.update(fallback_slip_hits)
            source = fallback_slip_source
            scenario = "Slip only"
            resolved = fallback_slip_resolved

        # =================================================================
        # FASE 4: Narrowing Wajib (Currency & Periode) + Strict Rematch
        # =================================================================
        is_beda_curr = False
        is_beda_per  = False
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
                    eff_polis, eff_slip, insured_sus,
                    polis_osbal, slip_osbal, insured_osbal,
                    facode_osbal_idx,
                    slipdb_rows, lookup_slipdb,
                    facul_rows, lookup_facul,
                    like_polis,
                    effective_sus_date=effective_date, is_lineslip=_is_ls
                )
                if rm:
                    src, scen, hits, res = rm
                    source   = src
                    scenario = scen
                    matched_list = hits
                    resolved = res
                    is_beda_curr = False
                    is_beda_per  = False

            matched_osbal_idx = set(matched_list)

        # =================================================================
        # FASE 5: Narrowing Bertingkat Non-Destruktif
        # Hierarki: Polis -> Cert -> Slip -> Insured
        # Aturan Khusus:
        #   1. + Cert HANYA boleh aktif jika base matching adalah POLIS
        #   2. Urutan label tersusun rapi: Polis + Cert + Slip
        # =================================================================
        if len(matched_osbal_idx) > 1:
            matched_list = list(matched_osbal_idx)

            has_polis_base = "Polis" in scenario
            has_cert_conf  = False
            has_slip_conf  = False

            # Step 1: Konfirmasi Sertifikat (KHUSUS jika base matching memiliki POLIS)
            # Karena sertifikat secara hierarki asuransi melekat langsung pada Polis Induk
            if has_polis_base:
                sus_cert_vals = set()
                for cv in _collect_clean_values(sus, sertif_sus):
                    c_clean = _clean_cert_str(cv)
                    if c_clean and c_clean != "0":
                        sus_cert_vals.add(c_clean)

                # Ekstrak sertifikat dari suffix nomor polis jika ada (misal: 10901...-000013)
                for pv in [_normalize(sus.get("POLIS_ORI", ""))] + [_normalize(sus.get(c, "")) for c in (eff_polis or [])]:
                    if pv and "-" in pv:
                        m_cert = re.search(r'-(\d{3,8})(?:/|\b|$)', pv)
                        if m_cert:
                            cl = _clean_cert_str(m_cert.group(1))
                            if cl and cl != "0":
                                sus_cert_vals.add(cl)

                if sus_cert_vals and sertif_osbal:
                    confirmed_cert = [
                        i for i in matched_list
                        if any(
                            _cert_in_range(cv, osbal_rows[i].get(c, ""))
                            for cv in sus_cert_vals
                            for c in sertif_osbal
                        )
                    ]
                    if confirmed_cert:
                        matched_list = confirmed_cert
                        has_cert_conf = True

            # Step 2: Konfirmasi Slip (Exact & Range)
            slip_confirm_vals = _collect_clean_values(sus, eff_slip)
            if slip_confirm_vals:
                confirmed_slip = [i for i in matched_list if _ref_has_value(osbal_rows[i], slip_osbal, "slip_ori", slip_confirm_vals)]
                if confirmed_slip:
                    matched_list = confirmed_slip
                    has_slip_conf = True
                elif len(matched_list) > 1:
                    confirmed_slip_range = [i for i in matched_list if any(_slip_in_range(sv, osbal_rows[i]) for sv in slip_confirm_vals)]
                    if confirmed_slip_range:
                        matched_list = confirmed_slip_range
                        has_slip_conf = True

            # Step 2b: Konfirmasi Polis (jika match awal via Slip)
            if not has_polis_base and "Slip" in scenario:
                polis_confirm_vals = _collect_all_values(sus, eff_polis, "polis_ori")
                if polis_confirm_vals:
                    confirmed_pol = [i for i in matched_list if _ref_has_value(osbal_rows[i], polis_osbal, "polis_ori", polis_confirm_vals)]
                    if confirmed_pol:
                        matched_list = confirmed_pol
                        has_polis_base = True
                        has_slip_conf  = True
                        # Karena polis terkonfirmasi, periksa sertifikat
                        sus_cert_vals = set()
                        for cv in _collect_clean_values(sus, sertif_sus):
                            c_clean = _clean_cert_str(cv)
                            if c_clean and c_clean != "0":
                                sus_cert_vals.add(c_clean)
                        for pv in polis_confirm_vals:
                            if pv and "-" in pv:
                                m_cert = re.search(r'-(\d{3,8})(?:/|\b|$)', pv)
                                if m_cert:
                                    cl = _clean_cert_str(m_cert.group(1))
                                    if cl and cl != "0":
                                        sus_cert_vals.add(cl)
                        if sus_cert_vals and sertif_osbal:
                            confirmed_cert = [
                                i for i in matched_list
                                if any(
                                    _cert_in_range(cv, osbal_rows[i].get(c, ""))
                                    for cv in sus_cert_vals
                                    for c in sertif_osbal
                                )
                            ]
                            if confirmed_cert:
                                matched_list = confirmed_cert
                                has_cert_conf = True

            # Susun label skenario Polis / Cert / Slip dengan urutan terstruktur: Polis -> Cert -> Slip
            if has_polis_base:
                parts = ["Polis"]
                if has_cert_conf:
                    parts.append("Cert")
                if has_slip_conf:
                    parts.append("Slip")
                if len(parts) > 1:
                    scenario = " + ".join(parts)
            elif has_slip_conf and "Slip" in scenario:
                scenario = "Slip only"

            # Step 3: Konfirmasi Insured Strict (Non-Destruktif)
            ins_vals = _collect_insured_values(sus, insured_sus, "insured_ori")
            if ins_vals and len(matched_list) > 1:
                confirmed = [i for i in matched_list if _ref_matches_insured(osbal_rows[i], insured_osbal, ins_vals)]
                if confirmed:
                    matched_list = confirmed
                    if "Insured" not in scenario:
                        insured_vals_before: set = set()
                        for _idx in list(matched_osbal_idx):
                            for _c in insured_osbal:
                                if _c != "FAC_INSURED_ORI":
                                    _v = _normalize(osbal_rows[_idx].get(_c, ""))
                                    if _v:
                                        insured_vals_before.add(_v)
                        if len(insured_vals_before) > 1:
                            base = scenario.replace(" only", "") if not scenario.endswith("only") else scenario
                            scenario = base + " + Insured"

            matched_osbal_idx = set(matched_list)
        # Post-matching normalisasi label
        if is_beda_curr or is_beda_per:
            osbal_ref   = [osbal_rows[i] for i in list(matched_osbal_idx)] if matched_osbal_idx else [{}]
            osbal_count = 0
            resolved = False
        elif not matched_osbal_idx:
            source = None
            scenario = "Unmatching"
            osbal_ref = [{}]
            osbal_count = 0
            resolved = False
        else:
            osbal_ref   = [osbal_rows[i] for i in matched_osbal_idx]
            osbal_count = len(osbal_ref)

        if scenario == "INSURED_LIKE":
            scenario = "Insured only"
        elif scenario and " only + " in scenario:
            scenario = scenario.replace(" only + ", " + ")

        fac_codes = {_normalize(r.get(OSBAL_FACODE_COL, "")) for r in osbal_ref if r}
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
            "_like_log":     [m for _, m in _like_log],
        })

    print(f"  Matching selesai dalam {time.perf_counter()-t4:.1f}s", flush=True)
    for _row in raw_results:
        for _method in _row.get("_like_log", []):
            if _method == "SUBSTR":    stat_like_substr  += 1
            elif _method == "SUBSET":  stat_like_subset  += 1
            elif _method == "OVERLAP": stat_like_overlap += 1

    def _is_invalid_match(r):
        return (
            r.get("is_beda_curr", False)
            or r.get("is_beda_per", False)
            or r.get("source") is None
            or r.get("base_scenario", "") in ("Unmatching", "UNMATCHED")
        )

    # [4b] Fac code claim & osbal accumulation (row asli dipertahankan)
    print("\n[4b/6] Fac code claim & osbal accumulation ...", flush=True)

    claimed: set = set()

    for r in raw_results:
        if len(r["fac_codes"]) == 1 and not _is_invalid_match(r):
            claimed |= r["fac_codes"]

    for r in raw_results:
        if len(r["fac_codes"]) > 1:
            remaining = r["fac_codes"] - claimed
            if len(remaining) == 1 and not _is_invalid_match(r):
                r["_originally_multi_fac"] = True
                claimed |= remaining

    groups: dict = {}
    final:  list = []

    for r in raw_results:
        r.setdefault("suspend_count", 1)
        if len(r["fac_codes"]) == 1 and not _is_invalid_match(r):
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

    print(f"  {len(raw_results):,} suspend rows -> {len(final):,} output rows (row asli dipertahankan)", flush=True)
    total_like = stat_like_substr + stat_like_subset + stat_like_overlap
    print(f"  [DEBUG] _like_match total hits: {total_like}", flush=True)
    print(f"    SUBSTR  : {stat_like_substr}", flush=True)
    print(f"    SUBSET  : {stat_like_subset}", flush=True)
    print(f"    OVERLAP (>={_INSURED_LIKE_THRESHOLD:.0%}): {stat_like_overlap}", flush=True)

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

        # Format skenario final sesuai standar ACA/Astra/TPI
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
            is_beda_curr=r.get("is_beda_curr", False),
            is_beda_per=r.get("is_beda_per", False),
        )
        if "_merged_amount_ori" in r:
            out["_MERGED_AMOUNT_ORI"] = r["_merged_amount_ori"]
        out["_IS_BEDA_CURR"] = r.get("is_beda_curr", False)
        out["_IS_BEDA_PER"]  = r.get("is_beda_per", False)
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
    df["SKENARIO"] = df["SKENARIO"].replace({"INSURED_LIKE": "Insured only"}).astype(str).str.strip()
    df.loc[df["SKENARIO"].isin(["UNMATCHED", "Unmatching"]), "FLAG_PROD"] = "Unmatching"
    df.loc[df["SKENARIO"] == "UNMATCHED", "SKENARIO"] = "Unmatching"

    for col in FINAL_COLUMNS:
        if col not in df.columns:
            df[col] = ""
    df = df[FINAL_COLUMNS]

    # [6] Export ke Excel di folder "output matching" (Defensive Export Safe)
    _out_file = OUTPUT_FILE
    print(f"\n[6/6] Saving to: {_out_file} ...", flush=True)
    for c in df.select_dtypes(include=['object', 'string']).columns:
        df[c] = df[c].astype(str).str.replace(r'[\x00-\x08\x0B\x0C\x0E-\x1F]', '', regex=True)

    try:
        df.to_excel(_out_file, index=False)
    except PermissionError:
        date_str = time.strftime("_%d%b").lower()  # contoh: _27oct
        _out_file = OUTPUT_FILE.replace(".xlsx", f"{date_str}.xlsx")
        print(f"  [WARN] File sedang terbuka di Excel -> menyimpan ke: {_out_file}", flush=True)
        df.to_excel(_out_file, index=False)

    # Format kolom teks sebagai "@" di Excel (cegah auto-konversi numerik Excel)
    try:
        from openpyxl import load_workbook
        workbook = load_workbook(_out_file)
        worksheet = workbook.active
        text_cols = [
            "POLIS", "POLICY_CLEAN_1", "POLICY_CLEAN_2", "POLICY_CLEAN_3", "POLICY_CLEAN_4",
            "SERTIF_CLEAN_1", "SLIP_NO", "SLIP_NO_CLN",
            "RECEIPT NO", "CREDIT NOTES", "DETAIL RINCIAN NO",
        ]
        text_col_indices = [idx for idx, col in enumerate(df.columns, start=1) if col in text_cols]
        for col_idx in text_col_indices:
            for row in range(2, worksheet.max_row + 1):
                cell = worksheet.cell(row=row, column=col_idx)
                cell.number_format = "@"
                if cell.value is not None:
                    cell.value = str(cell.value)
        workbook.save(_out_file)
    except Exception as e:
        print(f"  [WARN] Format text gagal: {e}")

    # Optional: apply styling jika excel_styler tersedia
    try:
        # pyrefly: ignore [missing-import]
        from excel_styler import apply_purple_column_style
        apply_purple_column_style(_out_file, "matching")
    except Exception:
        pass

    elapsed = time.perf_counter() - t_start
    print(f"\n{'=' * 60}")
    print(f"  Selesai dalam {elapsed:.1f}s ({elapsed/60:.1f} menit)")
    print(f"  Rows  : {len(df):,}  |  Cols: {len(df.columns)}  |  File: {_out_file}")
    print(f"\n  FLAG_PROD:")
    flag_counts = df["FLAG_PROD"].astype(str).str.strip().value_counts()
    for flag, count in flag_counts.items():
        print(f"    {str(flag):<45}: {count:,}")
    print(f"\n  SKENARIO:")
    scen_counts = df["SKENARIO"].astype(str).str.strip().value_counts()
    for sce, count in scen_counts.items():
        print(f"    {str(sce):<50}: {count:,}")
    print("=" * 60)


if __name__ == "__main__":
    run()
