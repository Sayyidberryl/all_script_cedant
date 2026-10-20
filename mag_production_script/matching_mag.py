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
# KONFIGURASI  (CEDANT: MAG)
#
# Sumber arsitektur : matching_ibs.py  (IBS reference, berbasis prod_sus_1_new.py ACA)
# Dokumen aturan    : Workflow Teknis Suspend 2026.docx, mag.txt
# Nama kolom        : disesuaikan dengan output cleansing MAG
#
# Delta utama vs IBS:
#   1. BORDERO / Open Cover Marine Cargo - TIDAK ADA untuk MAG.
#      Seluruh kode bordero dihapus dari pipeline (sama dengan IBS).
#   2. FASE 1 MAG: pencarian berdasarkan POLIS (bukan DESC_3_CLEAN seperti IBS).
#      Urutan fase: Polis -> Slip -> Insured -> Currency & Periode -> Narrowing.
#      (Mengikuti urutan ACA, bukan IBS).
#   3. Tidak ada rule tambahan pengecekan DESC 3 di data suspend.
#   4. Kolom-kolom MAG mengikuti naming convention dari data riil MAG (mag.txt).
# =============================================================================

_BASE_DIR  = os.path.dirname(os.path.abspath(__file__))   # folder MAG/
_DATA_DIR  = os.path.join(_BASE_DIR, "data")
_CLEAN_DIR = os.path.join(_BASE_DIR, "data_cleaned") if os.path.exists(os.path.join(_BASE_DIR, "data_cleaned")) else _DATA_DIR
_RAW_DIR   = os.path.join(_BASE_DIR, "data_raw") if os.path.exists(os.path.join(_BASE_DIR, "data_raw")) else _DATA_DIR


def _find_file_in(directory: str, basenames: list) -> str:
    """Cari file pertama yang ada dari daftar kandidat nama dalam direktori tertentu."""
    for name in basenames:
        for d in [directory, _DATA_DIR, _BASE_DIR]:
            p = os.path.join(d, name)
            if os.path.exists(p):
                return p
    return os.path.join(directory, basenames[0])


# Path file MAG
SUSPEND_FILE = _find_file_in(_CLEAN_DIR, ["mag_output_suspend.xlsx"])
OSBAL_FILE   = _find_file_in(_CLEAN_DIR, ["mag_output_osbal.xlsx"])
FACUL_FILE   = _find_file_in(_CLEAN_DIR, ["mag_output_facul.xlsx"])
SLIPDB_FILE  = _find_file_in(_RAW_DIR, [
    "5. Database Rekap RI Slip.xlsx",
    "5.Database Rekap RI Slip.xlsx",
    "Database Rekap RI Slip.xlsx",
])
OUTPUT_FILE  = os.path.join(_BASE_DIR, "matching_mag.xlsx")

# Tidak ada BORDERO_FILE -- MAG tidak punya Open Cover Marine Cargo.

# Kolom FAC CODE di masing-masing tabel referensi
OSBAL_FACODE_COL  = "CCOS_REF_CODE"
FACUL_FACODE_COL  = "FAC_CODE"
SLIPDB_FACODE_COL = "FAC_CODE"   # dapat di-override runtime sesuai header aktual

# Kolom polis/slip suspend
CLSDT_POLIS_COL = "CLSDT_POLICY_NO"
CLSDT_SLIP_COL  = "CLSDT_SLIP_NO"
FAC_POLIS_COL   = "FAC_POLICY_NO"
FAC_SLIP_COL    = "FAC_SLIP"

# Kolom currency/periode sesuai data riil MAG
SUSPEND_CURR_COL = "CURR ORI"
OSBAL_CURR_COL   = "CCOS_CURR"  # kolom currency di tabel OSBAL IBS
SUSPEND_DATE_COL = "RECEIPT DATE"
OSBAL_DATE_COL   = "FAC_COM_DATE"

# Marker baris administratif yang dikecualikan dari matching
_EXCLUDED_REF_MARKERS = ["HUTANG PIUTANG", "DATA SUSPENSE"]

# Marker LINESLIP - pakai RECEIPT DATE - 1 bulan sebagai tanggal efektif
_LINESLIP_MARKERS = ("LINESLIP", "LINE SLIP")

_MIN_TOKEN_LEN = 3
_TOKEN_RE      = re.compile(r"[^A-Z0-9]+")
_DELIMITER_RE  = re.compile(r"[,;|+\s]+")
_QQ_STUCK_RE   = re.compile(r'([A-Za-z0-9])QQ\b', re.IGNORECASE)

# Token generik yang dikecualikan saat menghitung overlap INSURED (kurangi false positive).
# Kata-kata ini sangat umum di nama perusahaan Indonesia sehingga tidak cukup diskriminatif
# untuk dijadikan sinyal kecocokan.
_INSURED_STOPWORDS: frozenset = frozenset({
    # Badan usaha & hukum
    "PT", "CV", "TBK", "UD", "PD", "NV", "LLC", "LTD", "INC", "PERSERO", "CORP", "CORPORATION", "CO", "FIRMA",
    # Entitas grup / holding / afiliasi
    "GROUP", "GROUPS", "GRUP", "HOLDING", "HOLDINGS", "SUBSIDIARY", "SUBSIDIARIES", "AFFILIATE", "AFFILIATES",
    # Sektor industri & keuangan umum (agar tidak match hanya karena sama-sama 'BANK' atau 'ASURANSI')
    "BANK", "ASURANSI", "INSURANCE", "REASURANSI", "REINSURANCE", "FINANCE", "MULTIFINANCE", "LEASING",
    "SEKURITAS", "SECURITIES", "CAPITAL", "ASSET", "MANAGEMENT", "INVESTAMA", "INVESTASI", "INVESTMENT",
    "LIFE", "GENERAL", "SYARIAH", "VENTURA", "VENTURES", "BROKER", "BROKERS", "DEVELOPMENT",
    # Wilayah geografis umum
    "INDONESIA", "INDO", "JAKARTA", "ASIA", "GLOBAL", "NUSANTARA", "INTERNATIONAL", "INTERNASIONAL", "NASIONAL",
    # Kata umum / modifier perusahaan Indonesia (mencegah false match seperti 'JAYA', 'MAKMUR')
    "COMPANY", "COMPANIES", "ENTERPRISE", "ENTERPRISES", "INDUSTRY", "INDUSTRIES", "INDUSTRI", "TRADING", "TRADE",
    "UTAMA", "MANDIRI", "MAKMUR", "SEJAHTERA", "JAYA", "ABADI", "SUKSES", "BERSAMA", "SENTOSA", "PRIMA", "KARYA",
    "MITRA", "ANUGERAH", "MEGAH", "LESTARI", "BERKAH", "INDAH", "PERSADA", "BINTANG", "SINAR",
    # Kata umum lain yang menyebabkan false positive chimeric co-insured
    "RAYA", "MULIA", "AGUNG", "CENTRAL", "USAHA", "NIAGA",
})

# Penanda entitas grup / holding induk (wajib ada di salah satu nama untuk mengaktifkan GROUP_LIKE)
_GROUP_MARKERS: frozenset = frozenset({"GRUP", "GROUP", "HOLDING", "HOLDINGS"})

# Overlap coefficient = |A∩B| / min(|A|, |B|) setelah exclude stopword.
# 0.65 = kalau 2 dari 3 token signifikan overlap → match.
# Naikkan ke 0.75-0.80 jika false positive terlalu banyak.
_INSURED_LIKE_THRESHOLD: float = 0.65

# Threshold ketat untuk POLIS/SLIP (identifier dokumen — harus lebih presisi).
_POLIS_LIKE_THRESHOLD: float  = 0.80

# Kolom join saat akumulasi suspend - nama sesuai file riil MAG
# Ref: mag.txt -- kolom clean di suspend: polis_ori, clean polis 1, slip_ori, clean slip 1,
#                                          insured_ori, clean insured 1..5
_SUSPEND_JOIN_COLS = [
    "RECEIPT NO", "CREDIT NOTES", "DETAIL RINCIAN NO", "RECEIPT DATE",
    "CEDANT NAME", "CEDANT SHRT NAME",
    "insured_ori", "clean insured 1", "clean insured 2", "clean insured 3", "clean insured 4", "clean insured 5",
    "CURR ORI", "CURR PAY", "AMOUNT PAY",
    "polis_ori", "clean polis 1", "CERTIFICATE", "clean sertif 1",
    "slip_ori", "clean slip 1",
    "DESC 1", "DESC 2", "DESC 3", "DESC 4",
    "STATUS", "REC_TYPE",
]
_SUSPEND_SUM_COLS = ["AMOUNT ORI"]

# FINAL_COLUMNS: 32 kolom standar output (sesuai readme/final_output_v1.xlsx)
FINAL_COLUMNS = [
    "CCOS_DOC_NO", "CCOS_REF_CODE",
    "RECEIPT NO", "CREDIT NOTES", "DETAIL RINCIAN NO", "RECEIPT DATE",
    "CEDANT NAME", "CEDANT SHRT NAME",
    "INSURED_ORI", "INSURED_1", "INSURED_2",
    "CURR ORI", "AMOUNT ORI", "AMOUNT_ORI_MIN1",
    "CURR PAY", "AMOUNT PAY",
    "CCOS_OR_BAL", "CCOS_BAL_DUE", "DIFERENCE",
    "FLAG_PROD",
    "POLIS_ORI", "POLIS_CLN", "SERTIF_CLN",
    "SLIP_NO_ORI", "SLIP_NO_CLN",
    "DESC 1", "DESC 2", "DESC 3", "DESC 4",
    "STATUS", "REC_TYPE",
    "SKENARIO",
]



# =============================================================================
# STRING UTILITIES
# =============================================================================

@functools.lru_cache(maxsize=131072)
def _norm_cached(s: str) -> str:
    # Normalisasi string - cached karena dipanggil jutaan kali
    text = s.strip().upper().replace("S/D", "SD")
    return re.sub(r"\s+", " ", text) if "  " in text else text


def _normalize(value) -> str:
    # Uppercase + strip + collapse spasi, return '' untuk null
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    if isinstance(value, (float, np.floating)):
        if np.isnan(value):
            return ""
        if value.is_integer():
            return _norm_cached(str(int(value)))
    s = str(value).strip()
    if s.endswith('.0') and s[:-2].replace('-', '').isdigit():
        s = s[:-2]
    return _norm_cached(s)


def _format_sertif(val) -> str:
    # Format nilai sertif jadi string 6-digit (contoh: 7 -> '000007', 1623.0 -> '001623')
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return ""
    s = str(val).strip()
    if not s or s.lower() == "nan":
        return ""
    if isinstance(val, (int, float)):
        return str(int(val)).zfill(6)
    # Jika berisi beberapa sertif dipisah koma (hasil merge)
    if "," in s:
        parts = [p.strip() for p in s.split(",") if p.strip()]
        formatted_parts = []
        for p in parts:
            if p.endswith('.0'):
                p = p[:-2]
            formatted_parts.append(p.zfill(6) if p.isdigit() else p)
        return ", ".join(dict.fromkeys(formatted_parts))
    if s.endswith('.0'):
        s = s[:-2]
    return s.zfill(6) if s.isdigit() else s


def _extract_sertif_from_polis(polis_val) -> str:
    """Ekstraksi suffix certificate numeric (1-6 digit) dari polis_ori.
    Contoh: 02031124000016-001623 -> 001623
    """
    if polis_val is None or (isinstance(polis_val, float) and pd.isna(polis_val)):
        return ""
    polis = str(polis_val).strip().upper()
    if not polis or polis == "-" or re.search(r"\bSUSPENSE\b|\bENDORSEMENT\b", polis):
        return ""
    m = re.search(r"^.*?-(\d{1,6})(?:-\d+/\d+)?$", polis)
    if m:
        return m.group(1).zfill(6)
    return ""


def _extract_fac_year(fac: str):
    """Ambil tahun (4 digit) dari 2 digit prefix fac code (contoh: 24FYA2A0 -> 2024)."""
    if not fac:
        return None
    m = re.match(r"^(\d{2})", str(fac).strip())
    if m:
        yy = int(m.group(1))
        return 2000 + yy if yy < 70 else 1900 + yy
    return None


def _parse_osbal_cert_ranges(cert_val) -> list:
    """Parse nilai certificate dari OSBAL menjadi list tuple (start, end).
    Mendukung format:
      - '000001 SD 000084' / '000455 S/D 000536' -> [(1, 84), (455, 536)]
      - '001199, 001210' (2 endpoint dari S/D di cleaning_osbal) -> [(1199, 1210)]
      - '000108' (angka tunggal) -> [(108, 108)]
    """
    if cert_val is None or (isinstance(cert_val, float) and pd.isna(cert_val)):
        return []
    s = str(cert_val).strip()
    if not s or s.lower() == "nan":
        return []

    ranges = []

    # 1. Format eksplisit 'SD' atau 'S/D'
    m_sd = re.search(r"(\d+)\s*(?:SD|S/D|-)\s*(\d+)", s, re.IGNORECASE)
    if m_sd:
        ranges.append((int(m_sd.group(1)), int(m_sd.group(2))))
        return ranges

    # 2. Format 2 endpoint koma (hasil cleaning_osbal.py: '001199, 001210')
    parts = [p.strip() for p in s.split(",") if p.strip()]
    if len(parts) == 2 and all(p.isdigit() for p in parts):
        n1, n2 = int(parts[0]), int(parts[1])
        if n1 <= n2:
            ranges.append((n1, n2))
            return ranges

    # 3. Fallback: nomor sertifikat tunggal / daftar angka
    for p in parts:
        m = re.search(r"\b(\d+)\b", p)
        if m:
            val = int(m.group(1))
            ranges.append((val, val))

    return ranges


def _is_sertif_in_ranges(sertif_num: int, ranges: list) -> bool:
    """Cek apakah angka sertifikat berada di dalam salah satu rentang."""
    return any(lo <= sertif_num <= hi for lo, hi in ranges)


def _tokenize(text: str) -> frozenset:
    # Pecah string jadi frozenset token alfanumerik (min 3 karakter)
    if not text:
        return frozenset()
    text = _QQ_STUCK_RE.sub(r'\1 QQ', text)
    return frozenset(t for t in _TOKEN_RE.split(text) if len(t) >= _MIN_TOKEN_LEN)


@functools.lru_cache(maxsize=100000)
def _split_composite(text: str) -> frozenset:
    # Pecah string berdelimiter (koma/plus/spasi/titik koma/pipe) jadi token
    if not text:
        return frozenset()
    return frozenset(tok.strip() for tok in _DELIMITER_RE.split(text) if tok.strip())


def _get_clean_cols(columns: list, prefix: str) -> list:
    # Ambil nama kolom yang berawalan prefix (case-insensitive)
    if not columns:
        return []
    return [c for c in columns if str(c).lower().startswith(prefix.lower())]


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


# =============================================================================
# ROW DICT HELPERS
# =============================================================================

def _is_excluded_ref_row(row: dict, polis_cols: list, slip_cols: list) -> bool:
    # True jika baris referensi adalah baris administratif
    for col in list(polis_cols) + ["polis_ori"] + list(slip_cols) + ["slip_ori"]:
        val = _normalize(row.get(col, ""))
        if val and any(m in val for m in _EXCLUDED_REF_MARKERS):
            return True
    return False


def _is_lineslip_row(sus_row: dict, insured_cols: list) -> bool:
    # True jika baris suspend mengandung marker LINESLIP di kolom insured
    # MAG: kolom ori adalah "insured_ori"
    for col in list(insured_cols) + ["insured_ori", "INSURED_ORI"]:
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
            if "polis" in col and val.isdigit() and len(val) == 13:
                z = val.zfill(14)
                if z not in seen:
                    seen.add(z)
                    result.append(z)
    return result


def _collect_all_values(row: dict, clean_cols: list, ori_col: str) -> list:
    # Gabung nilai clean menjadi list unik
    return list(set(_collect_clean_values(row, clean_cols)))


def _get_effective_sus_cols(sus_row: dict, base_polis_cols: list, base_slip_cols: list) -> tuple:
    """Tentukan kolom efektif polis/slip per baris.

    MAG mengikuti pola ACA:
    - FASE 1: Polis (CLSDT_POLICY_NO jika terisi, fallback clean polis cols)
    - FASE 2: Slip  (CLSDT_SLIP_NO jika terisi, fallback clean slip cols)
    """
    clsdt_polis = _normalize(sus_row.get(CLSDT_POLIS_COL, ""))
    fac_polis   = _normalize(sus_row.get(FAC_POLIS_COL,   ""))
    if clsdt_polis:
        eff_polis  = [CLSDT_POLIS_COL]
        like_polis = [CLSDT_POLIS_COL]
    else:
        eff_polis  = base_polis_cols
        like_polis = list(base_polis_cols) + ([FAC_POLIS_COL] if fac_polis else [])

    clsdt_slip = _normalize(sus_row.get(CLSDT_SLIP_COL, ""))
    fac_slip   = _normalize(sus_row.get(FAC_SLIP_COL, ""))
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
    for rv in ref_vals:
        if any(v in _split_composite(rv) for v in query_values if v and len(v) >= 2):
            return True
    if any(
        v in rv
        for v in query_values if v and len(v) >= 10
        for rv in ref_vals if rv
    ):
        return True

    # Khusus INSURED: periksa kesamaan minimal 1 kata grup & anak perusahaan
    # Syarat: salah satu pihak harus secara eksplisit merupakan entitas grup/holding (GRUP/GROUP/HOLDING)
    if ori_col == "insured_ori":
        for qv in query_values:
            q_raw = _tokenize(qv)
            q_sig = q_raw - _INSURED_STOPWORDS
            if not q_sig:
                continue
            for rv in ref_vals:
                r_raw = _tokenize(rv)
                r_sig = r_raw - _INSURED_STOPWORDS
                has_group = bool(q_raw & _GROUP_MARKERS or r_raw & _GROUP_MARKERS)
                if has_group and (q_sig & r_sig):
                    return True

    return False


def _slip_in_range(sus_slip: str, osbal_row: dict) -> bool:
    """Cek apakah nomor slip suspend berada di dalam range slip OSBAL."""
    if not sus_slip:
        return False
    s_clean = re.sub(r'\D', '', str(sus_slip))
    if not s_clean:
        return False
    slip_ori = str(osbal_row.get("slip_ori", "")).upper()
    if "+" in slip_ori or "," in slip_ori:
        if not re.search(r'\b(SD|S/D|TO)\b', slip_ori):
            return False
    elif not re.search(r'\b(SD|S/D|TO|-)\b', slip_ori):
        return False

    sl1 = re.sub(r'\D', '', str(osbal_row.get("clean slip 1", "")))
    sl2 = re.sub(r'\D', '', str(osbal_row.get("clean slip 2", "")))
    if sl1 and sl2 and len(sl1) == len(s_clean) and len(sl2) == len(s_clean):
        if sl1[:6] != sl2[:6] or s_clean[:6] != sl1[:6]:
            return False
        try:
            n_sus = int(s_clean)
            n_sl1 = int(sl1)
            n_sl2 = int(sl2)
            if min(n_sl1, n_sl2) <= n_sus <= max(n_sl1, n_sl2):
                return True
        except ValueError:
            pass
    return False


_BDO_MONTH_MAP = {
    'JAN': 1, 'JANUARI': 1, 'JANUARY': 1,
    'FEB': 2, 'FEBRUARI': 2, 'FEBRUARY': 2,
    'MAR': 3, 'MARET': 3, 'MARCH': 3,
    'APR': 4, 'APRIL': 4,
    'MEI': 5, 'MAY': 5,
    'JUN': 6, 'JUNI': 6, 'JUNE': 6,
    'JUL': 7, 'JULI': 7, 'JULY': 7,
    'AGS': 8, 'AGT': 8, 'AGUSTUS': 8, 'AUG': 8, 'AUGUST': 8,
    'SEP': 9, 'SEPT': 9, 'SEPTEMBER': 9,
    'OKT': 10, 'OKTOBER': 10, 'OCT': 10, 'OCTOBER': 10,
    'NOV': 11, 'NOVEMBER': 11,
    'DES': 12, 'DEC': 12, 'DESEMBER': 12, 'DECEMBER': 12,
}

def _parse_bordero_slip(text: str) -> set:
    """Ekstraksi pasangan (bulan, tahun) dari teks slip bertipe bordero."""
    if not text:
        return set()
    s = str(text).upper()
    if not any(k in s for k in ['BDO', 'BORDERO']):
        return set()
    res = set()
    for m in re.finditer(r'\b(0?[1-9]|1[0-2])\s+(\d{4})\b', s):
        res.add((int(m.group(1)), int(m.group(2))))
    for word, m_num in _BDO_MONTH_MAP.items():
        for m in re.finditer(r'\b' + word + r'\s+(\d{4})\b', s):
            res.add((m_num, int(m.group(1))))
    return res


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


def _build_lookup(rows: list, polis_cols: list, slip_cols: list,
                  insured_cols: list, facode_col: str = None, label: str = "") -> tuple:
    """Bangun semua struktur lookup untuk satu tabel referensi (10-tuple).

    Index tuple identik dengan prod_sus_1_new.py:
    [0] slip exact  [1] polis exact  [2] insured exact
    [3][4][5] unused
    [6] tok_slip  [7] tok_polis  [8] tok_insured  [9] facode_index
    """
    if not rows:
        if label:
            print(f"  {label:<6} done (0.0s)", flush=True)
        return ({}, {}, {}, {}, {}, {}, ({}, {}), ({}, {}), ({}, {}), {})

    t0 = time.perf_counter()
    excluded = {i for i, r in enumerate(rows) if _is_excluded_ref_row(r, polis_cols, slip_cols)}

    facode_index = _build_facode_index(rows, facode_col, excluded=excluded) if facode_col else {}
    lookup = (
        _build_exact_index(rows, slip_cols,    excluded=excluded),   # [0] slip
        _build_exact_index(rows, polis_cols,   excluded=excluded),   # [1] polis
        _build_exact_index(rows, insured_cols, excluded=excluded),   # [2] insured
        {}, {}, {},                                                   # [3][4][5] unused
        _build_token_index(rows, slip_cols,    "slip_ori",    excluded=excluded),  # [6]
        _build_token_index(rows, polis_cols,   "polis_ori",   excluded=excluded),  # [7]
        _build_token_index(rows, insured_cols, "insured_ori", excluded=excluded),  # [8]
        facode_index,                                                 # [9]
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


def _like_match(
    query_values: list,
    token_index: dict,
    rows: list,
    ref_cols: list,
    ori_col: str = None,
    stopwords: frozenset = frozenset(),
    overlap_threshold: float = _INSURED_LIKE_THRESHOLD,
    allow_group_match: bool = False,
    match_log: list = None,
) -> list:
    """Implementasi murni ACA (prod_sus_1_new.py): pencocokan substring literal (q in rv)."""
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
        ref_values = [_normalize(rows[i].get(c, "")) for c in ref_cols if rows[i].get(c)]
        for q in valid_queries:
            if any(q in rv for rv in ref_values if rv):
                matched.append(i)
                if match_log is not None:
                    match_log.append((i, "SUBSTR"))
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
    # Matching polis/slip: Stage 1 exact, Stage 2 LIKE
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
                hit = _like_match(
                    values, idx_tok, rows, ref_cols,
                    stopwords=frozenset(),
                    overlap_threshold=_POLIS_LIKE_THRESHOLD,
                )
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
            hit = _like_match(
                values, idx_tok, rows, insured_ref_cols or [],
                stopwords=_INSURED_STOPWORDS,
                overlap_threshold=_INSURED_LIKE_THRESHOLD,
            )
            if hit:
                return hit, "INSURED_LIKE"

    return [], "Unmatching"


# =============================================================================
# NARROWING
# =============================================================================

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
    """Rematch ketat dengan filter currency + periode (R1-R5 standar ACA).

    [IBS DELTA] S6 poin 3: apakah perlu jalur R0 berbasis DESC_3_CLEAN?
    Asumsi #2: Implementasi konservatif -- R1-R5 standard ACA dipakai apa adanya.
    Konfirmasi tim diperlukan sebelum menambah jalur R0.
    """
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
            curr_hits = _strict_filter(_like_match(
                like_vals, lookup_osbal[7][0], osbal_rows, polis_ref_cols,
                stopwords=frozenset(),
                overlap_threshold=_POLIS_LIKE_THRESHOLD,
            ))
            if curr_hits:
                return "OSBAL", "Polis only", curr_hits, True

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

    # R5: OSBAL insured fallback — exact, lalu LIKE
    insured_vals_r5 = _collect_clean_values(suspend_row, insured_sus_cols)
    curr_hits = _strict_filter(_exact_match(insured_vals_r5, lookup_osbal[2]))
    if curr_hits:
        return "OSBAL", "Insured only", curr_hits, True

    if lookup_osbal[8] and lookup_osbal[8][0] and insured_vals_r5:
        curr_hits = _strict_filter(_like_match(
            insured_vals_r5, lookup_osbal[8][0], osbal_rows, insured_ref_cols,
            stopwords=_INSURED_STOPWORDS,
            overlap_threshold=_INSURED_LIKE_THRESHOLD,
        ))
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
    # Jalankan satu pass matching polis/slip + narrowing currency/periode
    lkp_slip_cln, lkp_polis_cln, lkp_ins_cln, _, _, _, tok_slip, tok_polis, tok_ins, _ = lookup

    matched, label = _match_polis_slip(
        suspend_row, lkp_slip_cln, lkp_polis_cln, {}, {},
        polis_sus_cols, slip_sus_cols, rows=rows,
        slip_ref_cols=slip_ref_cols, polis_ref_cols=polis_ref_cols,
        token_slip=tok_slip, token_polis=tok_polis,
        polis_sus_like_cols=polis_sus_like_cols, slip_sus_like_cols=slip_sus_like_cols,
    )

    if matched:
        if currency_narrowing and len(matched) > 1:
            matched = _narrow_by_currency(suspend_row, matched, rows)
            if len(matched) > 1:
                matched = _narrow_by_periode(suspend_row, matched, rows,
                                             effective_sus_date=effective_sus_date,
                                             is_lineslip=is_lineslip)
        # Narrowing komplementer POLIS <-> SLIP
        narrowed = matched
        scenario = _map_scenario(label)
        if "SLIP" in label and len(matched) > 1:
            polis_vals = _collect_all_values(suspend_row, polis_sus_cols, "polis_ori")
            confirmed = [i for i in matched if _ref_has_value(rows[i], polis_ref_cols, "polis_ori", polis_vals)]
            if confirmed:
                narrowed = confirmed
                scenario = "Slip + Polis"
            else:
                scenario = "Slip only"
        elif "POLIS" in label and len(matched) > 1:
            slip_vals = _collect_all_values(suspend_row, slip_sus_cols, "slip_ori")
            confirmed = [i for i in matched if _ref_has_value(rows[i], slip_ref_cols, "slip_ori", slip_vals)]
            if not confirmed and slip_vals:
                confirmed = [i for i in matched if any(_slip_in_range(sv, rows[i]) for sv in slip_vals)]
            if confirmed:
                narrowed = confirmed
                scenario = "Polis + Slip"
            else:
                scenario = "Polis only"
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
    for to_remove in [
        " -> Bordero OC Override", " (Bordero)", " (Beda Currency)",
        " (Beda Periode)", " (> 1 Fac)", " (> 1 Fac Bordero)",
        " (DESC 1)", " + Amount",
    ]:
        s = s.replace(to_remove, "")
    s = s.strip()
    if s.startswith("Slip + Polis"):
        s = s.replace("Slip + Polis", "Polis + Slip")
    if " only + " in s:
        s = s.replace(" only + ", " + ")
    if s == "Polis + Sertif":
        s = "Polis + Cert"
    if s in ("Insured", "Polis", "Slip"):
        s = s + " only"
    if s in ("Insured + Cert", "Insured + Cert only"):
        s = "Insured only"
    if s in ("Slip + Cert", "Slip + Cert only"):
        s = "Slip only"
    return s


def _format_final_scenario(
    base_scenario: str,
    source: str,
    is_bordero: bool = False,
    is_beda_curr: bool = False,
    is_beda_per: bool = False,
    num_fac: int = 1,
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
    """Bangun satu baris output.

    MAG: Nama kolom sesuai data riil MAG (insured_ori, clean insured 1/2/3,
    polis_ori, clean polis 1, slip_ori, clean slip 1, dst.)
    """
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

    fac_label = (
        _facode_label(osbal_rows, "OSBAL") if has_osbal else
        _facode_label(osbal_rows, source)  if has_match else ""
    )

    final_scenario = "Unmatching" if not source or scenario in ("Unmatching", "UNMATCHED") else scenario

    # Ambil nilai sertifikat dari berbagai kandidat kolom, fallback ke ekstraksi polis_ori
    _sertif_raw = ""
    for _k in ("CERTIFICATE", "clean sertif 1", "certificate 1", "SERTIF_CLN"):
        _v = suspend_row.get(_k)
        if _v is not None and not (isinstance(_v, float) and pd.isna(_v)) and str(_v).strip() != "":
            _sertif_raw = _v
            break
    _sertif_formatted = _format_sertif(_sertif_raw)
    if not _sertif_formatted:
        _sertif_formatted = _extract_sertif_from_polis(
            suspend_row.get("polis_ori", "") or suspend_row.get("POLIS_ORI", "")
        )

    return {
        "CCOS_DOC_NO":       ccos["CCOS_DOC_NO"] if has_osbal else "",
        "CCOS_REF_CODE":     ccos["CCOS_REF_CODE"],
        "RECEIPT NO":        suspend_row.get("RECEIPT NO",        ""),
        "CREDIT NOTES":      suspend_row.get("CREDIT NOTES",      ""),
        "DETAIL RINCIAN NO": suspend_row.get("DETAIL RINCIAN NO", ""),
        "RECEIPT DATE":      suspend_row.get("RECEIPT DATE",      ""),
        "CEDANT NAME":       suspend_row.get("CEDANT NAME",       ""),
        "CEDANT SHRT NAME":  suspend_row.get("CEDANT SHRT NAME",  ""),

        # MAG: nama kolom insured dari file riil MAG (mag.txt: insured_ori, clean insured 1..)
        "INSURED_ORI": suspend_row.get("insured_ori",     "") or suspend_row.get("INSURED_ORI", ""),
        "INSURED_1":   suspend_row.get("clean insured 1", "") or suspend_row.get("INSURED_1",   ""),
        "INSURED_2":   suspend_row.get("clean insured 2", "") or suspend_row.get("INSURED_2",   ""),

        "CURR ORI":   suspend_row.get("CURR ORI",   ""),
        "AMOUNT ORI": suspend_row.get("AMOUNT ORI", ""),
        "CURR PAY":   suspend_row.get("CURR PAY",   ""),
        "AMOUNT PAY": suspend_row.get("AMOUNT PAY", ""),

        "CCOS_OR_BAL":  ccos["CCOS_OR_BAL"],
        "CCOS_BAL_DUE": ccos["CCOS_BAL_DUE"],

        # Kolom sesuai template readme/final_output_v1.xlsx
        "POLIS_ORI":   suspend_row.get("polis_ori",    "") or suspend_row.get("POLIS_ORI", "") or suspend_row.get("POLIS", ""),
        "POLIS_CLN":   suspend_row.get("clean polis 1","") or suspend_row.get("POLIS_CLN", "") or suspend_row.get("POLICY_CLEAN_1", ""),
        "SERTIF_CLN":  _sertif_formatted,
        "SLIP_NO_ORI": suspend_row.get("slip_ori",     "") or suspend_row.get("SLIP_NO_ORI", "") or suspend_row.get("SLIP_NO", ""),
        "SLIP_NO_CLN": suspend_row.get("clean slip 1", "") or suspend_row.get("SLIP_NO_CLN", "") or suspend_row.get("SLIP_NO_CLEAN", ""),

        "DESC 1": suspend_row.get("DESC 1", ""),
        "DESC 2": suspend_row.get("DESC 2", ""),
        "DESC 3": suspend_row.get("DESC 3", ""),
        "DESC 4": suspend_row.get("DESC 4", ""),

        "STATUS":   suspend_row.get("STATUS",   ""),
        "REC_TYPE": suspend_row.get("REC_TYPE", ""),

        "SKENARIO":            final_scenario,
        "_OSBAL_ROW_COUNT":    osbal_count,
        "_SUSPEND_ROW_COUNT":  suspend_count,
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
    """Hitung AMOUNT_ORI_MIN1, DIFERENCE, dan FLAG_PROD (10 kategori).

    Matriks np.select identik dengan ACA.
    Kondisi is_gt1_ocmc (Bordero OC Override) selalu False karena MAG tidak punya bordero.
    """
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

    df["DIFERENCE"] = (amount_neg - balance_due).round(2)
    # Fix floating point micro-residual (e.g. -2.84e-14 instead of 0.0)
    df.loc[df["DIFERENCE"].abs() < 1e-5, "DIFERENCE"] = 0.0
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

def _load_excel(path: str, label: str, optional: bool = False, header_row: int = 0) -> tuple:
    # Load Excel ke (list[dict], list[str]) dengan pickle cache otomatis.
    # header_row: baris ke-n (0-indexed) yang menjadi nama kolom (default=0).
    # Untuk SLIPDB (Database Rekap RI Slip.xlsx): header ada di baris ke-2 -> header_row=1.
    if not os.path.exists(path):
        print(f"  {label}: {path} ({'tidak ada, dilewati' if optional else 'TIDAK DITEMUKAN'})", flush=True)
        return [], []

    # Cache key menyertakan header_row agar cache dari header berbeda tidak tertukar
    cache = path + f".h{header_row}.cache.pkl"
    # Hapus cache lama (tanpa suffix .h*) jika masih ada agar tidak terbaca
    _old_cache = path + ".cache.pkl"
    if os.path.exists(_old_cache) and not os.path.exists(cache):
        try:
            os.remove(_old_cache)
        except Exception:
            pass

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
    print(f"  {label}: {path} (header_row={header_row}) ...", flush=True)
    try:
        df = pd.read_excel(path, header=header_row)
    except Exception as e:
        print(f"  [RETRY] engine default gagal ({e}), coba openpyxl ...", flush=True)
        df = pd.read_excel(path, header=header_row, engine="openpyxl")

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
    """Pipeline utama MAG V1: matching 5 fase (tanpa Bordero).

    Urutan fase (mengikuti ACA, bukan IBS):
      FASE 1: Polis  -> cross-check OSBAL -> SLIPDB -> FACUL
      FASE 2: Slip   -> cross-check OSBAL -> SLIPDB -> FACUL (fallback jika FASE 1 miss)
      FASE 3: Insured                                         (fallback jika FASE 1 & 2 miss)
      FASE 4: Narrowing Wajib (Currency & Periode) + Strict Rematch
      FASE 5: Narrowing Non-Destruktif (Slip Exact -> Slip Range -> Insured)

    Tidak ada FASE Bordero Open Cover -- MAG tidak memiliki data OC Marine Cargo.
    Tidak ada pengecekan DESC 3 (khusus IBS, tidak berlaku untuk MAG).
    """
    t_start = time.perf_counter()
    print("\n" + "=" * 60, flush=True)
    print("  PRODUCTION SCRIPT -- SUSPEND MATCHING  (MAG) -- V1", flush=True)
    print("=" * 60, flush=True)

    # Cek apakah OUTPUT_FILE sedang dibuka di program lain (misal: Excel)
    if os.path.exists(OUTPUT_FILE):
        try:
            with open(OUTPUT_FILE, "r+"):
                pass
        except PermissionError:
            print(f"\n[PERHATIAN] File '{OUTPUT_FILE}' sedang dibuka di Microsoft Excel atau program lain!", flush=True)
            print("Agar file utama dapat langsung diperbarui, silakan TUTUP file tersebut di Excel terlebih dahulu.\n", flush=True)

    # [1] Load data
    print("\n[1/6] Loading data ...", flush=True)
    suspend_rows, suspend_cols = _load_excel(SUSPEND_FILE, "SUSPEND")
    osbal_rows,   osbal_cols   = _load_excel(OSBAL_FILE,   "OSBAL")
    facul_rows,   facul_cols   = _load_excel(FACUL_FILE,   "FACUL")
    # [MEK 3] SLIPDB: header ada di baris ke-2 Excel (row index 1) → header_row=1
    # Sebelumnya dibaca dengan header=0 (default) sehingga kolom menjadi angka
    # (2, 0, '0.1', ...) dan data baris pertama malah berisi nama kolom asli.
    slipdb_rows,  slipdb_cols  = _load_excel(SLIPDB_FILE,  "SLIPDB", optional=True, header_row=1)
    # Bordero tidak dimuat -- MAG tidak punya Open Cover Marine Cargo.

    # Diagnostik kolom kunci
    print("\n  --- Cek kolom SUSPEND ---", flush=True)
    print(f"  Semua kolom SUSPEND: {suspend_cols}", flush=True)
    # Kolom MAG berdasarkan mag.txt: polis_ori, clean polis 1, slip_ori, clean slip 1,
    #                                 insured_ori, clean insured 1..5
    for col in ["polis_ori", "clean polis 1", "slip_ori", "clean slip 1",
                "insured_ori", "clean insured 1", "clean insured 2",
                SUSPEND_CURR_COL, SUSPEND_DATE_COL]:
        status = "OK" if col in suspend_cols else "[MISSING]"
        print(f"    {col:<25}: {status}", flush=True)

    print("\n  --- Cek kolom OSBAL ---", flush=True)
    print(f"  Semua kolom OSBAL: {osbal_cols}", flush=True)
    # Kolom MAG OSBAL berdasarkan mag.txt:
    # insured_ori, clean insured 1..5, polis_ori, clean polis 1, certificate 1,
    # clean polis 2, certificate 2, ..., clean polis 5, slip_ori, clean slip 1..5
    for col in [OSBAL_FACODE_COL, OSBAL_CURR_COL, OSBAL_DATE_COL,
                "CCOS_DOC_NO", "CCOS_OR_BAL", "CCOS_BAL_DUE",
                "polis_ori", "clean polis 1", "slip_ori", "clean slip 1",
                "insured_ori", "clean insured 1"]:
        status = "OK" if col in osbal_cols else "[MISSING]"
        print(f"    {col:<25}: {status}", flush=True)

    # Pre-parse bordero slip periods
    print("  Pre-parsing bordero slip periods ...", flush=True)
    for r in osbal_rows:
        r["_bdo_periods"] = _parse_bordero_slip(r.get("slip_ori", ""))
    for r in suspend_rows:
        r["_bdo_periods"] = _parse_bordero_slip(r.get("slip_ori", ""))

    # Pre-parse tanggal OSBAL FAC_COM_DATE
    print("  Pre-parsing OSBAL FAC_COM_DATE ...", flush=True)
    for r in osbal_rows:
        raw = r.get(OSBAL_DATE_COL, "")
        r["_com_date_parsed"] = pd.to_datetime(raw, errors="coerce") \
            if raw and not (isinstance(raw, float) and pd.isna(raw)) else None

    # Pre-parse OSBAL certificate ranges & fac code years
    print("  Pre-parsing OSBAL certificate ranges & fac code years ...", flush=True)
    cert_cols_osbal = [c for c in osbal_cols if any(k in str(c).lower() for k in ["certificate", "sertif", "certif", "sertf"])]
    for r in osbal_rows:
        r_ranges = []
        for ccol in cert_cols_osbal:
            cval = r.get(ccol)
            if pd.notna(cval):
                r_ranges.extend(_parse_osbal_cert_ranges(cval))
        r["_cert_ranges"] = r_ranges
        r["_fac_year"] = _extract_fac_year(r.get(OSBAL_FACODE_COL, ""))

    # Pre-parse tanggal Suspend RECEIPT DATE
    print("  Pre-parsing Suspend RECEIPT DATE ...", flush=True)
    for r in suspend_rows:
        raw = r.get(SUSPEND_DATE_COL, "")
        r["_sus_date_parsed"] = pd.to_datetime(raw, errors="coerce") \
            if raw and not (isinstance(raw, float) and pd.isna(raw)) else None

    # [2] Deteksi kolom bersih
    print("\n[2/6] Detecting clean columns ...", flush=True)

    # MAG Suspend columns (dari mag.txt):
    #   polis: polis_ori, clean polis 1
    #   slip:  slip_ori, clean slip 1
    #   insured: insured_ori, clean insured 1..5
    polis_sus = _get_clean_cols(suspend_cols, "clean polis") or []
    slip_sus  = _get_clean_cols(suspend_cols, "clean slip")  or []
    insured_sus = _get_clean_cols(suspend_cols, "clean insured") or []

    # Fallback jika kolom clean tidak ditemukan
    if not polis_sus:
        polis_sus = [c for c in ["POLIS_CLN_1", "POLIS_CLN", "clean polis 1"] if c in suspend_cols]
    if not slip_sus:
        slip_sus  = [c for c in ["SLIP_NO_CLN_1", "SLIP_NO_CLN", "clean slip 1"] if c in suspend_cols]
    if not insured_sus:
        insured_sus = [c for c in ["INSURED_CLN_1", "clean insured 1"] if c in suspend_cols]

    print(f"  polis_sus    : {polis_sus}", flush=True)
    print(f"  slip_sus     : {slip_sus}", flush=True)
    print(f"  insured_sus  : {insured_sus}", flush=True)

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

    polis_osbal   = _get_clean_cols(osbal_cols, "POLIS_CLN") or _get_clean_cols(osbal_cols, "clean polis") or []
    slip_osbal    = _get_clean_cols(osbal_cols, "SLIP_NO_CLN") or _get_clean_cols(osbal_cols, "clean slip") or []
    insured_osbal = _get_clean_cols(osbal_cols, "INSURED_CLN") or _get_clean_cols(osbal_cols, "clean insured") or []
    sertif_osbal  = _get_clean_cols(osbal_cols, "clean sertif") or []

    polis_facul   = _get_clean_cols(facul_cols, "POLIS_CLN") or _get_clean_cols(facul_cols, "clean polis") or []
    slip_facul    = _get_clean_cols(facul_cols, "SLIP_NO_CLN") or _get_clean_cols(facul_cols, "clean slip") or []
    insured_facul = _get_clean_cols(facul_cols, "INSURED_CLN") or _get_clean_cols(facul_cols, "clean insured") or []

    polis_slipdb   = _get_clean_cols(slipdb_cols, "POLIS_CLN") or _get_clean_cols(slipdb_cols, "clean polis") or []
    slip_slipdb    = _get_clean_cols(slipdb_cols, "SLIP_NO_CLN") or _get_clean_cols(slipdb_cols, "clean slip") or []
    insured_slipdb = _get_clean_cols(slipdb_cols, "INSURED_CLN") or _get_clean_cols(slipdb_cols, "clean insured") or []

    if not polis_slipdb:
        polis_slipdb = [c for c in ["CLSDT_POLICY_NO", "FAC_POLICY_NO", "POLIS", "POLICY_NO", "POLICY NO"] if c in slipdb_cols]
    if not slip_slipdb:
        slip_slipdb  = [c for c in ["CLSDT_SLIP_NO", "FAC_SLIP", "SLIP", "SLIP_NO", "SLIP NO"] if c in slipdb_cols]
    if not insured_slipdb:
        insured_slipdb = [c for c in ["FAC_INSURED", "INSURED", "CEDANT_NAME", "CEDANT NAME"] if c in slipdb_cols]

    global SLIPDB_FACODE_COL
    if   "FAC_CODE"      in slipdb_cols: SLIPDB_FACODE_COL = "FAC_CODE"
    elif "FAC CODE"      in slipdb_cols: SLIPDB_FACODE_COL = "FAC CODE"
    elif "FAC. CODE"     in slipdb_cols: SLIPDB_FACODE_COL = "FAC. CODE"  # Database Rekap RI Slip.xlsx
    elif "CCOS_REF_CODE" in slipdb_cols: SLIPDB_FACODE_COL = "CCOS_REF_CODE"

    def _prepend(col, cols_list, header):
        return ([col] + cols_list) if col in header and col not in cols_list else cols_list

    polis_osbal  = _prepend("CLSDT_POLICY_NO", polis_osbal,  osbal_cols)
    slip_osbal   = _prepend("CLSDT_SLIP_NO",   slip_osbal,   osbal_cols)
    polis_facul  = _prepend("CLSDT_POLICY_NO", polis_facul,  facul_cols)
    slip_facul   = _prepend("CLSDT_SLIP_NO",   slip_facul,   facul_cols)
    polis_slipdb = _prepend("CLSDT_POLICY_NO", polis_slipdb, slipdb_cols)
    slip_slipdb  = _prepend("CLSDT_SLIP_NO",   slip_slipdb,  slipdb_cols)

    # [3] Bangun lookup index
    # [IBS DELTA] Tidak ada bordero_idx / bordero_polis_index
    print("\n[3/6] Building lookup indexes ...", flush=True)
    lookup_osbal  = _build_lookup(osbal_rows,  polis_osbal,  slip_osbal,  insured_osbal,  OSBAL_FACODE_COL,  label="OSBAL")
    lookup_facul  = _build_lookup(facul_rows,  polis_facul,  slip_facul,  insured_facul,  FACUL_FACODE_COL,  label="FACUL")
    lookup_slipdb = _build_lookup(slipdb_rows, polis_slipdb, slip_slipdb, insured_slipdb, SLIPDB_FACODE_COL, label="SLIPDB")

    facode_osbal_idx = lookup_osbal[9]

    kw_osbal = dict(
        rows=osbal_rows, lookup=lookup_osbal,
        polis_ref_cols=polis_osbal, slip_ref_cols=slip_osbal, insured_ref_cols=insured_osbal,
        facode_col=OSBAL_FACODE_COL, currency_narrowing=True,
        sertif_ref_cols=sertif_osbal,
    )
    kw_facul = dict(
        rows=facul_rows, lookup=lookup_facul,
        polis_ref_cols=polis_facul, slip_ref_cols=slip_facul, insured_ref_cols=insured_facul,
        facode_col=FACUL_FACODE_COL, gunakan_narrow_aca=True,
    )
    kw_slipdb = dict(
        rows=slipdb_rows, lookup=lookup_slipdb,
        polis_ref_cols=polis_slipdb, slip_ref_cols=slip_slipdb, insured_ref_cols=insured_slipdb,
        facode_col=SLIPDB_FACODE_COL, gunakan_narrow_aca=True,
    )

    # [IBS DELTA] Bordero index tidak dibangun -- IBS tidak punya Open Cover MC
    print("  Bordero: TIDAK DIGUNAKAN untuk IBS (fase dihapus dari pipeline)", flush=True)

    # [3.1] Bangun sertif lookup OSBAL (sekali di luar loop untuk performa)
    _cert_osbal_cols = [c for c in osbal_cols if "certificate" in str(c).lower() or "sertif" in str(c).lower()]
    _sertif_osbal_idx: dict = {}
    for _ii, _rr in enumerate(osbal_rows):
        for _cc in _cert_osbal_cols:
            _cv = _rr.get(_cc)
            if _cv is not None and not (isinstance(_cv, float) and pd.isna(_cv)):
                _cv_str = re.sub(r'\.0$', '', str(_cv).strip())
                if _cv_str and _cv_str.lower() != 'nan':
                    _sertif_osbal_idx.setdefault(_cv_str.zfill(6), []).append(_ii)
                    _sertif_osbal_idx.setdefault(_cv_str, []).append(_ii)

    # Deteksi kolom sertifikat suspend sekali
    _sertif_sus_cols = []
    if "CERTIFICATE" in suspend_cols:
        _sertif_sus_cols.append("CERTIFICATE")
    for _sc in suspend_cols:
        if "clean sertif" in str(_sc).lower() and _sc not in _sertif_sus_cols:
            _sertif_sus_cols.append(_sc)

    # [4] Matching per baris
    total = len(suspend_rows)
    print(f"\n[4/6] Matching {total:,} suspend rows ...", flush=True)
    t4 = time.perf_counter()
    raw_results = []
    stat_like_substr  = 0   # hit karena whole-string substring (backward compat)
    stat_like_subset  = 0   # hit karena token subset match
    stat_like_overlap = 0   # hit karena overlap coefficient >= threshold
    stat_like_group   = 0   # hit karena kesamaan minimal 1 kata grup/anak perusahaan

    for n, sus in enumerate(suspend_rows, 1):
        if n % 1000 == 0:
            elapsed_so_far = time.perf_counter() - t4
            rate = n / elapsed_so_far if elapsed_so_far > 0 else 0
            eta  = (total - n) / rate if rate > 0 else 0
            print(f"    Memproses baris {n:,} / {total:,}  ({rate:.0f} rows/s, ETA {eta:.0f}s)", flush=True)

        if str(sus.get("STATUS", "")).strip().upper() == "MATCHING":
            continue

        eff_polis, eff_slip, like_polis, like_slip = _get_effective_sus_cols(sus, polis_sus, slip_sus)
        _is_ls = sus.get("_is_lineslip", False)
        effective_date = sus.get("_sus_date_lineslip") if _is_ls else sus.get("_sus_date_parsed")

        matched_osbal_idx = set()
        source = None
        scenario = ""
        resolved = False
        _like_log: list = []   # audit trail match method per baris ini

        kw_sus_row = dict(
            polis_sus_cols=eff_polis, slip_sus_cols=eff_slip, insured_sus_cols=insured_sus,
            polis_sus_like_cols=like_polis, slip_sus_like_cols=like_slip,
            effective_sus_date=effective_date, is_lineslip=_is_ls,
            sertif_sus_cols=_sertif_sus_cols,
        )

        # =================================================================
        # FASE 1: Pencarian awal berdasarkan POLIS
        # MAG mengikuti urutan ACA: Polis dulu, baru Slip, baru Insured.
        # Cross-check: OSBAL -> SLIPDB -> FACUL
        # =================================================================
        found_by_polis = False
        polis_vals = _collect_clean_values(sus, eff_polis)
        if polis_vals:
            osb_hits = _exact_match(polis_vals, lookup_osbal[1])
            if osb_hits:
                matched_osbal_idx.update(osb_hits)
                found_by_polis = True

                # [Pass A] Polis + Sertifikat: jika suspend punya nomor sertif,
                # coba saring kandidat polis ke baris OSBAL yang sertifnya cocok.
                # Lebih spesifik dari Polis only — langsung mengunci 1 baris definitif
                # untuk master policy dengan banyak sertifikat (misal OLAM GLOBAL AGRI).
                if _sertif_sus_cols:
                    _sus_cert_raw = ""
                    for _ck in _sertif_sus_cols:
                        _cv = sus.get(_ck)
                        if _cv is not None and not (isinstance(_cv, float) and pd.isna(_cv)):
                            _s = re.sub(r'\.0$', '', str(_cv).strip())
                            if _s and _s.lower() != 'nan':
                                _sus_cert_raw = _s
                                break
                    if _sus_cert_raw:
                        _cert_variants = [_sus_cert_raw, _sus_cert_raw.zfill(6)]
                        _cert_confirmed = [
                            idx for idx in osb_hits
                            if any(
                                idx in _sertif_osbal_idx.get(cv, [])
                                for cv in _cert_variants
                            )
                        ]
                        if _cert_confirmed:
                            matched_osbal_idx = set(_cert_confirmed)
                            scenario = "Polis + Cert"

            if slipdb_rows and not matched_osbal_idx:
                sld_hits = _exact_match(polis_vals, lookup_slipdb[1])
                if sld_hits:
                    res_idx, _ = _resolve_facode([slipdb_rows[idx] for idx in sld_hits], SLIPDB_FACODE_COL, facode_osbal_idx, osbal_rows)
                    if res_idx:
                        matched_osbal_idx.update(res_idx)
                        found_by_polis = True
                        resolved = True

            if facul_rows and not matched_osbal_idx:
                fac_hits = _exact_match(polis_vals, lookup_facul[1])
                if fac_hits:
                    res_idx, _ = _resolve_facode([facul_rows[idx] for idx in fac_hits], FACUL_FACODE_COL, facode_osbal_idx, osbal_rows)
                    if res_idx:
                        matched_osbal_idx.update(res_idx)
                        found_by_polis = True
                        resolved = True

        if found_by_polis:
            source = "OSBAL" if not resolved else ("SLIPDB" if slipdb_rows and matched_osbal_idx else "FACUL")
            # Override source if it was found in OSBAL
            if osb_hits: source = "OSBAL"
            if not scenario:
                scenario = "Polis only"

        # =================================================================
        # FASE 2: Pencarian berdasarkan SLIP (fallback jika FASE 1 miss)
        # Cross-check: OSBAL -> SLIPDB -> FACUL
        # =================================================================
        if not found_by_polis:
            slip_vals = _collect_clean_values(sus, eff_slip)
            if slip_vals:
                osb_hits = _exact_match(slip_vals, lookup_osbal[0])
                if osb_hits:
                    matched_osbal_idx.update(osb_hits)
                    source = "OSBAL"
                    scenario = "Slip only"
                else:
                    if slipdb_rows:
                        sld_hits = _exact_match(slip_vals, lookup_slipdb[0])
                        if sld_hits:
                            res_idx, _ = _resolve_facode([slipdb_rows[idx] for idx in sld_hits], SLIPDB_FACODE_COL, facode_osbal_idx, osbal_rows)
                            if res_idx:
                                matched_osbal_idx.update(res_idx)
                                source = "SLIPDB"
                                scenario = "Slip only"
                                resolved = True
                    if not matched_osbal_idx and facul_rows:
                        fac_hits = _exact_match(slip_vals, lookup_facul[0])
                        if fac_hits:
                            res_idx, _ = _resolve_facode([facul_rows[idx] for idx in fac_hits], FACUL_FACODE_COL, facode_osbal_idx, osbal_rows)
                            if res_idx:
                                matched_osbal_idx.update(res_idx)
                                source = "FACUL"
                                scenario = "Slip only"
                                resolved = True

        # =================================================================
        # FASE 3: Pencarian berdasarkan INSURED (fallback FASE 1 & 2 miss)
        # =================================================================
        if not matched_osbal_idx:
            ins_vals = _collect_clean_values(sus, insured_sus)
            if ins_vals:
                # 1. Coba OSBAL Exact & Like
                osb_hits = set(_exact_match(ins_vals, lookup_osbal[2]))
                if not osb_hits and lookup_osbal[8] and lookup_osbal[8][0]:
                    like_hits = _like_match(
                        ins_vals, lookup_osbal[8][0], osbal_rows, insured_osbal,
                        stopwords=_INSURED_STOPWORDS,
                        overlap_threshold=_INSURED_LIKE_THRESHOLD,
                        allow_group_match=False,  # [Opsi2] GROUP_LIKE dinonaktifkan
                        match_log=_like_log,
                    )
                    if like_hits:
                        osb_hits.update(like_hits)
                if osb_hits:
                    matched_osbal_idx.update(osb_hits)
                    source = "OSBAL"
                    scenario = "Insured only"

                # [DESC 1 Fallback] Jika insured primary tidak ketemu, coba DESC 1
                # (berisi nama grup/holding induk) sebagai sumber insured tambahan
                if not osb_hits:
                    desc1_val = _normalize(sus.get("DESC 1", ""))
                    if desc1_val and desc1_val != "-" and desc1_val not in ins_vals:
                        desc1_hits = set(_exact_match([desc1_val], lookup_osbal[2]))
                        if not desc1_hits and lookup_osbal[8] and lookup_osbal[8][0]:
                            desc1_like = _like_match(
                                [desc1_val], lookup_osbal[8][0], osbal_rows, insured_osbal,
                                stopwords=_INSURED_STOPWORDS,
                                overlap_threshold=_INSURED_LIKE_THRESHOLD,
                                allow_group_match=False,
                                match_log=_like_log,
                            )
                            if desc1_like:
                                desc1_hits.update(desc1_like)
                        if desc1_hits:
                            matched_osbal_idx.update(desc1_hits)
                            source = "OSBAL"
                            scenario = "Insured only"
                
                # 2. Coba SLIPDB Exact & Like
                elif slipdb_rows and not matched_osbal_idx:
                    sld_hits = set(_exact_match(ins_vals, lookup_slipdb[2]))
                    if not sld_hits and lookup_slipdb[8] and lookup_slipdb[8][0]:
                        like_hits = _like_match(
                            ins_vals, lookup_slipdb[8][0], slipdb_rows, insured_slipdb,
                            stopwords=_INSURED_STOPWORDS,
                            overlap_threshold=_INSURED_LIKE_THRESHOLD,
                            allow_group_match=False,  # [Opsi2] GROUP_LIKE dinonaktifkan
                            match_log=_like_log,
                        )
                        if like_hits:
                            sld_hits.update(like_hits)
                    if sld_hits:
                        res_idx, _ = _resolve_facode([slipdb_rows[idx] for idx in sld_hits], SLIPDB_FACODE_COL, facode_osbal_idx, osbal_rows)
                        if res_idx:
                            matched_osbal_idx.update(res_idx)
                            source = "SLIPDB"
                            scenario = "Insured only"
                            resolved = True
                
                # 3. Coba FACUL Exact & Like
                elif facul_rows and not matched_osbal_idx:
                    fac_hits = set(_exact_match(ins_vals, lookup_facul[2]))
                    if not fac_hits and lookup_facul[8] and lookup_facul[8][0]:
                        like_hits = _like_match(
                            ins_vals, lookup_facul[8][0], facul_rows, insured_facul,
                            stopwords=_INSURED_STOPWORDS,
                            overlap_threshold=_INSURED_LIKE_THRESHOLD,
                            allow_group_match=False,  # [Opsi2] GROUP_LIKE dinonaktifkan
                            match_log=_like_log,
                        )
                        if like_hits:
                            fac_hits.update(like_hits)
                    if fac_hits:
                        res_idx, _ = _resolve_facode([facul_rows[idx] for idx in fac_hits], FACUL_FACODE_COL, facode_osbal_idx, osbal_rows)
                        if res_idx:
                            matched_osbal_idx.update(res_idx)
                            source = "FACUL"
                            scenario = "Insured only"
                            resolved = True

        # =================================================================
        # FASE 4: Narrowing Wajib (Currency & Periode) + Strict Rematch
        # Identik dengan ACA
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
                # _rematch_strict: R1-R5 standard ACA
                # Tidak ada jalur R0 berbasis DESC_3_CLEAN untuk MAG
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
                    src, scen, hits, _ = rm
                    source = src
                    scenario = _map_scenario(scen)
                    matched_list = hits
                    is_beda_curr = False
                    is_beda_per  = False

            matched_osbal_idx = set(matched_list)

        # =================================================================
        # FASE 5: Narrowing Utama Bertingkat Non-Destruktif
        # Identik dengan ACA
        # =================================================================
        if len(matched_osbal_idx) > 1:
            matched_list = list(matched_osbal_idx)

            # Step 1: Slip Exact Matching
            slip_vals = _collect_clean_values(sus, slip_sus if slip_sus else desc3_slip_vals)
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

            # Step 1c: Bordero Slip Month-Year Check (khusus bordero slip)
            if len(matched_list) > 1:
                sus_bdo = sus.get("_bdo_periods")
                if sus_bdo:
                    confirmed_bdo = [
                        idx for idx in matched_list
                        if osbal_rows[idx].get("_bdo_periods") and (sus_bdo & osbal_rows[idx]["_bdo_periods"])
                    ]
                    if confirmed_bdo:
                        matched_list = confirmed_bdo
                        if "Slip" not in scenario:
                            scenario = "Polis + Slip" if "Polis" in scenario else scenario

            # Step 2: Certificate Range Narrowing (khusus jika suspend punya nomor sertifikat)
            # Dijalankan sebelum Insured (mengikuti urutan ACA prod_sus_1_new.py Step 2)
            if len(matched_list) > 1:
                _sus_sertif_raw = ""
                for _k in ("CERTIFICATE", "clean sertif 1", "certificate 1", "SERTIF_CLN"):
                    _v = sus.get(_k)
                    if _v is not None and not (isinstance(_v, float) and pd.isna(_v)) and str(_v).strip() not in ("", "nan"):
                        _sus_sertif_raw = str(_v).strip()
                        break
                if not _sus_sertif_raw:
                    _sus_sertif_raw = _extract_sertif_from_polis(
                        sus.get("polis_ori", "") or sus.get("POLIS_ORI", "")
                    )

                if _sus_sertif_raw:
                    _s_clean = re.sub(r"\.0$", "", str(_sus_sertif_raw).strip())
                    if _s_clean.isdigit():
                        _s_num = int(_s_clean)
                        confirmed_cert = [
                            idx for idx in matched_list
                            if _is_sertif_in_ranges(_s_num, osbal_rows[idx].get("_cert_ranges", []))
                        ]
                        if confirmed_cert and len(confirmed_cert) < len(matched_list):
                            matched_list = confirmed_cert
                            if "Polis" in scenario:
                                scenario = "Polis + Cert"

            # Step 3: Insured Matching
            # Sesuai instruksi & acuan prod_sus_1_new.py:
            # - Insured menyempurnakan kandidat jika masih >1 kandidat.
            # - Jika sudah ketemu via Polis only, Polis + Cert, Polis + Slip, Slip only, dll.,
            #   label skenario TIDAK ditambahkan "+ Insured".
            # - Tambah "+ Insured" HANYA jika base skenario belum memiliki identifier primer (Polis, Slip, Cert).
            if len(matched_list) > 1:
                ins_vals = _collect_all_values(sus, insured_sus, "insured_ori")
                if ins_vals:
                    confirmed = [idx for idx in matched_list if _ref_has_value(osbal_rows[idx], insured_osbal, "insured_ori", ins_vals)]
                    if confirmed:
                        matched_list = confirmed
                        _has_primary = any(k in scenario for k in ("Polis", "Slip", "Cert"))
                        if not _has_primary and "Insured" not in scenario:
                            base = scenario.replace(" only", "") if " only" in scenario else scenario
                            scenario = base + " + Insured"

            # ----------------------------------------------------------------
            # Step 4: Amount Narrowing (Mekanisme 1 - adopsi ACA _narrow_by_amount_net)
            # Jika masih multi-FAC setelah Cert/Insured Narrowing, coba narrow
            # dengan mencocokkan AMOUNT ORI suspend vs CCOS_BAL_DUE / CCOS_OR_BAL tiap FAC.
            # CATATAN: Menyempitkan kandidat di memori tanpa menambah kata baru "+ Amount" ke SKENARIO.
            # ----------------------------------------------------------------
            if len(matched_list) > 1:
                amt_ori = sus.get("AMOUNT ORI") or sus.get("AMOUNT_ORI")
                try:
                    amt_ori_f = abs(float(str(amt_ori).replace(",", ".").strip())) if amt_ori is not None and str(amt_ori).strip() not in ("", "nan") else None
                except (ValueError, TypeError):
                    amt_ori_f = None

                if amt_ori_f is not None and amt_ori_f > 0:
                    fac_group: dict = {}
                    for idx in matched_list:
                        fac = _normalize(osbal_rows[idx].get(OSBAL_FACODE_COL, ""))
                        if fac:
                            fac_group.setdefault(fac, []).append(idx)

                    if len(fac_group) > 1:
                        tol = 0.01  # toleransi 1 sen
                        matched_facs = []
                        for fac, idxs in fac_group.items():
                            row_match = False
                            for i in idxs:
                                bd = osbal_rows[i].get("CCOS_BAL_DUE", 0)
                                ob = osbal_rows[i].get("CCOS_OR_BAL", 0)
                                try:
                                    if abs(abs(float(bd)) - amt_ori_f) <= tol or abs(abs(float(ob)) - amt_ori_f) <= tol:
                                        row_match = True
                                        break
                                except:
                                    pass

                            all_idx = facode_osbal_idx.get(fac, idxs)
                            bal_due_total = 0.0
                            or_bal_total  = 0.0
                            for i in all_idx:
                                bd = osbal_rows[i].get("CCOS_BAL_DUE", 0)
                                ob = osbal_rows[i].get("CCOS_OR_BAL", 0)
                                try: bal_due_total += float(bd) if not pd.isna(bd) else 0.0
                                except: pass
                                try: or_bal_total  += float(ob) if not pd.isna(ob) else 0.0
                                except: pass

                            if row_match or abs(abs(bal_due_total) - amt_ori_f) <= tol or abs(abs(or_bal_total) - amt_ori_f) <= tol:
                                matched_facs.append(fac)

                        if len(matched_facs) == 1:
                            narrow_fac = matched_facs[0]
                            narrow_idxs = [i for i in fac_group[narrow_fac] if (
                                abs(abs(float(osbal_rows[i].get("CCOS_BAL_DUE", 0) or 0)) - amt_ori_f) <= tol or
                                abs(abs(float(osbal_rows[i].get("CCOS_OR_BAL", 0) or 0)) - amt_ori_f) <= tol
                            )]
                            if not narrow_idxs:
                                narrow_idxs = fac_group[narrow_fac]

                            if len(narrow_idxs) < len(matched_list):
                                matched_list = narrow_idxs

            matched_osbal_idx = set(matched_list)

        # =================================================================
        # FASE Bordero Open Cover Override -- DIHAPUS UNTUK MAG
        # MAG tidak memiliki data Open Cover Marine Cargo.
        # "Bordero OC Override" dan "Matching >1 Fac code OC MC" tidak muncul.
        # Fase ini dihapus bersih.
        # =================================================================

        # Post-matching normalisasi label
        # Mengikuti prod_sus (teknis_matching_aca.md §2.1 & prod_sus_2.py):
        # Jika is_beda_curr -> scenario = "Beda Currency", osbal_count = 0.
        # Jika is_beda_per  -> scenario = "Beda Periode", osbal_count = 0.
        # Jika not matched_osbal_idx -> scenario = "Unmatching".
        if is_beda_curr or is_beda_per:
            osbal_ref = [osbal_rows[i] for i in matched_list] if matched_list else [{}]
            osbal_count = 0
            resolved = False
        elif not matched_osbal_idx:
            source = None; scenario = "Unmatching"
            osbal_ref = [{}]; osbal_count = 0; resolved = False
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
            "is_bordero": False,
            "is_beda_curr": is_beda_curr,
            "is_beda_per": is_beda_per,
            "osbal_idx": list(matched_osbal_idx) if not (is_beda_curr or is_beda_per) else list(matched_list),
            "osbal_count": osbal_count, "resolved": resolved, "fac_codes": fac_codes,
            "_like_log": [m for _, m in _like_log],
            "_orig_idx": n,
        })

    print(f"  Matching selesai dalam {time.perf_counter()-t4:.1f}s", flush=True)
    # Akumulasi counter method dari semua baris
    for _row in raw_results:
        for _method in _row.get("_like_log", []):
            if _method == "SUBSTR":        stat_like_substr  += 1
            elif _method == "SUBSET":      stat_like_subset  += 1
            elif _method == "OVERLAP":     stat_like_overlap += 1
            elif _method == "GROUP_LIKE":  stat_like_group   += 1

    # [4b] Fac code claim & osbal accumulation (identik dengan prod_sus_1_new.py)
    print("\n[4c/6] Fac code claim & osbal accumulation (row asli dipertahankan) ...", flush=True)

    claimed: set = set()
    for r in raw_results:
        if len(r["fac_codes"]) == 1 and r["source"] is not None and not r.get("is_beda_curr", False) and not r.get("is_beda_per", False):
            claimed |= r["fac_codes"]

    for r in raw_results:
        if len(r["fac_codes"]) > 1:
            remaining = r["fac_codes"] - claimed
            if len(remaining) == 1 and r["source"] is not None and not r.get("is_beda_curr", False) and not r.get("is_beda_per", False):
                # [Fac Code Claim] Narrow baris GT1 ke fac yang tersisa.
                # Mengikuti logika prod_sus_1_new.py:
                #   fac code yang sudah di-claim suspend lain dikeluarkan,
                #   jika sisa tepat 1 fac → baris ini di-narrow ke 1 fac (exact).
                # Ini mengkonversi ">1 Fac Code" → resolved exact 1 fac secara otomatis.
                fac = next(iter(remaining))
                r["fac_codes"] = remaining
                r["osbal_idx"] = [
                    i for i in r["osbal_idx"]
                    if _normalize(osbal_rows[i].get(OSBAL_FACODE_COL, "")) == fac
                ]
                r["osbal_count"] = len(r["osbal_idx"])
                claimed |= remaining  # klaim fac code sisa

    groups: dict = {}
    final:  list = []

    for r in raw_results:
        r.setdefault("suspend_count", 1)
        if (len(r["fac_codes"]) == 1 and r["source"] is not None
                and not r.get("is_beda_curr", False) and not r.get("is_beda_per", False)):
            fac = next(iter(r["fac_codes"]))
            groups.setdefault(fac, []).append(r)
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
            g["osbal_idx"]          = list(merged_osbal_idx)
            g["osbal_count"]        = len(merged_osbal_idx)
            g["suspend_count"]      = len(group)
            g["_merged_amount_ori"] = merged_sus.get("AMOUNT ORI", 0)
            final.append(g)

    # Urutkan kembali sesuai urutan baris suspend asli (identik dengan standar ASTRA & ACA)
    final.sort(key=lambda r: r.get("_orig_idx", 0))

    print(f"  {len(raw_results):,} suspend rows -> {len(final):,} output rows (row asli dipertahankan)", flush=True)
    total_like = stat_like_substr + stat_like_subset + stat_like_overlap + stat_like_group
    print(f"  [DEBUG] _like_match total hits    : {total_like}", flush=True)
    print(f"    SUBSTR      (substring lama)                      : {stat_like_substr}", flush=True)
    print(f"    SUBSET      (token subset)                        : {stat_like_subset}", flush=True)
    print(f"    OVERLAP     (overlap coeff >={_INSURED_LIKE_THRESHOLD:.0%}): {stat_like_overlap}", flush=True)
    print(f"    GROUP_LIKE  (kesamaan min 1 kata nama grup/anak)  : {stat_like_group}", flush=True)

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

    for col in FINAL_COLUMNS:
        if col not in df.columns:
            df[col] = ""
    df = df[FINAL_COLUMNS]

    # [6] Export
    print(f"\n[6/6] Saving to: {OUTPUT_FILE} ...", flush=True)
    for c in df.select_dtypes(include=['object', 'string']).columns:
        df[c] = df[c].astype(str).str.replace(r'[\x00-\x08\x0B\x0C\x0E-\x1F]', '', regex=True)

    target_out = OUTPUT_FILE
    try:
        df.to_excel(target_out, index=False)
    except PermissionError:
        print(f"\n[PERINGATAN] File '{OUTPUT_FILE}' terkunci karena sedang dibuka di Microsoft Excel.", flush=True)
        target_out = os.path.join(_BASE_DIR, "matching_mag_new.xlsx")
        print(f"Hasil berhasil disimpan ke file alternatif: {target_out}", flush=True)
        print("Silakan tutup Excel jika ingin menimpa file utama matching_mag.xlsx.\n", flush=True)
        df.to_excel(target_out, index=False)

    try:
        from openpyxl import load_workbook
        workbook = load_workbook(target_out)
        worksheet = workbook.active
        text_cols = [
            "POLIS_ORI", "POLIS_CLN", "SERTIF_CLN", "SLIP_NO_ORI", "SLIP_NO_CLN", 
            "RECEIPT NO", "CREDIT NOTES", "DETAIL RINCIAN NO"
        ]
        text_col_indices = [idx for idx, col in enumerate(df.columns, start=1) if col in text_cols]
        for col_idx in text_col_indices:
            for row in range(2, worksheet.max_row + 1):
                cell = worksheet.cell(row=row, column=col_idx)
                cell.number_format = "@"
                if cell.value is not None:
                    cell.value = str(cell.value)
        workbook.save(target_out)
    except Exception as e:
        print(f"  [WARN] Format text gagal: {e}")


    try:
        from excel_styler import apply_purple_column_style
        apply_purple_column_style(target_out, "matching")
    except Exception as e:
        print(f"  [WARN] Styling gagal: {e}")

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
