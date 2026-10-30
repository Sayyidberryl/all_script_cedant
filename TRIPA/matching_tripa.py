import os
import re
import sys
import time
import math
import functools

sys.modules['numexpr'] = None
sys.modules['bottleneck'] = None

import numpy as np
import pandas as pd
from dotenv import load_dotenv
from sqlalchemy import create_engine


# =============================================================================
# KONFIGURASI  (CEDANT: TRIPA)
#
# Pipeline Matching Suspend Tripa:
#   1. RULES TRIPA = SLIP -> POLIS -> INSURED.
#   2. Nama kolom Suspend mengikuti file tripa_output_suspend.xlsx.
#   3. Data Suspend difilter: HANYA cedant Tripa yang diproses & masuk output.
#   4. Bordero / Open Cover Marine Cargo: TIDAK DIGUNAKAN (tidak ada bordero).
#
# Urutan fase Tripa:
#   FASE 1  SLIP    (OSBAL -> SLIPDB -> FACUL)
#   FASE 2  POLIS   (OSBAL -> SLIPDB -> FACUL)  <- fallback jika slip miss
#   FASE 3  INSURED (OSBAL -> SLIPDB -> FACUL)  <- fallback jika polis miss
#   FASE 4  Narrowing wajib: Currency & Periode (+ strict rematch)
#   FASE 5  Narrowing bertingkat non-destruktif (slip -> cert -> insured)
# =============================================================================

_BASE_DIR  = os.path.dirname(os.path.abspath(__file__))   # folder IBS/
_CLEAN_DIR = os.path.join(_BASE_DIR, "data_cleaned")       # hasil cleansing
_RAW_DIR   = os.path.join(_BASE_DIR, "data_raw")           # data mentah


def _find_file_in(directory: str, basenames: list) -> str:
    """Cari file pertama yang ada dari daftar kandidat nama dalam direktori tertentu."""
    for name in basenames:
        p = os.path.join(directory, name)
        if os.path.exists(p):
            return p
    return os.path.join(directory, basenames[0])


# [IBS DELTA] Path file menggunakan suffix/nama IBS, bukan ACA
SUSPEND_FILE = _find_file_in(_CLEAN_DIR, ["tripa_output_suspend.xlsx", "tripa_output_suspense.xlsx"])
# OSBAL & FACUL: sesuaikan nama file jika berbeda. Kandidat kedua = file OSBAL/FACUL
# hasil cleansing yang sudah ada (biasanya berisi semua cedant, difilter Tripa di bawah).
OSBAL_FILE   = _find_file_in(_CLEAN_DIR, ["tripa_output_osbal.xlsx", "ibs_output_osbal_V2.xlsx"])
FACUL_FILE   = _find_file_in(_CLEAN_DIR, ["tripa_output_facul.xlsx", "ibs_output_facul_V2.xlsx"])
SLIPDB_FILE  = _find_file_in(_RAW_DIR, [
    "5. Database Rekap RI Slip.xlsx",
    "5.Database Rekap RI Slip.xlsx",
    "Database Rekap RI Slip.xlsx",
])
OUTPUT_FILE  = os.path.join(_BASE_DIR, "matching_tripa.xlsx")

# [IBS DELTA] Tidak ada BORDERO_FILE.
# IBS tidak punya data Open Cover Marine Cargo. Variabel ini tidak
# dideklarasikan agar jalur bordero tidak pernah dieksekusi secara struktural.

# Kolom FAC CODE di masing-masing tabel referensi
OSBAL_FACODE_COL  = "CCOS_REF_CODE"
FACUL_FACODE_COL  = "FAC_CODE"
SLIPDB_FACODE_COL = "FAC_CODE"   # dapat di-override runtime sesuai header aktual

# [TRIPA] Filter cedant: hanya baris yang nama cedant-nya mengandung kata ini.
# Dicek ke semua kolom yang namanya mengandung "CEDANT" (CEDANT NAME, CEDANT SHRT NAME).
CEDANT_KEYWORD = "TRIPA"
# True  -> tabel referensi (OSBAL/FACUL/SLIPDB) ikut difilter Tripa JIKA punya kolom cedant
#          (mencegah match insured nyasar ke cedant lain). Jika filter menghasilkan 0 baris,
#          filter diabaikan otomatis. Set False untuk memakai referensi apa adanya.
FILTER_REF_BY_CEDANT = True

# Kolom polis/slip suspend dipertahankan agar _get_effective_sus_cols() kompatibel
CLSDT_POLIS_COL = "CLSDT_POLICY_NO"
CLSDT_SLIP_COL  = "CLSDT_SLIP_NO"
FAC_POLIS_COL   = "FAC_POLICY_NO"
FAC_SLIP_COL    = "FAC_SLIP"

# [IBS DELTA] Nama kolom currency/periode sesuai data riil IBS
SUSPEND_CURR_COL = "CURR ORI"    # kolom currency di tabel Suspend IBS
OSBAL_CURR_COL   = "CCOS_CURR"  # kolom currency di tabel OSBAL IBS
SUSPEND_DATE_COL = "RECEIPT DATE"
OSBAL_DATE_COL   = "FAC_COM_DATE"

# Marker baris administratif yang dikecualikan dari matching
# [TRIPA] Baris yang mengandung salah satu kata ini TIDAK ikut proses matching:
#   - Data 3 (SUSPEND) : baris dilewati, tidak di-match dan tidak masuk output.
#   - Data 2 (OSBAL)   : baris tidak dijadikan kandidat (dilewati saat lookup).
# Kata kunci: LINESLIP / LINE SLIP, PENYELESAIAN SUSPEN(SE) / kata SUSPEN(SE),
#             (H)UTANG / PI(H)UTANG (hutang, utang, piutang, pihutang).
# Catatan: "SUSPEND" (dengan D) sengaja TIDAK ikut -- kolom STATUS tidak dicek.
_EXCLUDE_RE = re.compile(
    r"\bSUSPEN(?:S|SE)?\b"
    r"|\b(?:PI|P)?H?UTANG\b",
    flags=re.IGNORECASE,
)

# Kolom suspend yang dicek untuk pengecualian (berdasarkan awalan nama kolom).
_SUS_EXCL_PREFIXES = ("insured", "clean insured", "polis", "clean polis",
                      "slip", "clean slip", "desc")

# Bobot minimal token "khas" pada mode strict (fraksi dari log(jumlah baris referensi)).
# Token yang muncul di banyak baris referensi (JAYA, ABADI, PRIMA, ...) bobotnya kecil.
_STRICT_MIN_DISTINCT: float = 0.5

# Marker LINESLIP - pakai RECEIPT DATE - 1 bulan sebagai tanggal efektif
_LINESLIP_MARKERS = ("LINESLIP", "LINE SLIP")

_MIN_TOKEN_LEN = 3
_TOKEN_RE      = re.compile(r"[^A-Z0-9]+")
_DELIMITER_RE  = re.compile(r"[,;|+\s]+")

# Token generik yang dikecualikan saat menghitung overlap INSURED (kurangi false positive).
# Kata-kata ini sangat umum di nama perusahaan Indonesia sehingga tidak cukup diskriminatif
# untuk dijadikan sinyal kecocokan.
_INSURED_STOPWORDS: frozenset = frozenset({
    "PT", "CV", "TBK", "UD", "PD", "NV", "LLC", "LTD", "INC",
    "GROUP", "GROUPS",
    "INDONESIA", "INDO",
    "INDUSTRY", "INDUSTRIES", "INDUSTRI",
    "INTERNATIONAL", "INTERNASIONAL",
    "COMPANY", "COMPANIES", "CORP", "CORPORATION",
    "ENTERPRISE", "ENTERPRISES",
    "TRADING", "TRADE",
    "UTAMA", "MANDIRI", "MAKMUR", "SEJAHTERA",
    "LINE", "LINES",
    "SHIPPING", "SHIP", "SHIPS",
    "LOGISTICS", "LOGISTIC",
    "TRANSPORT", "TRANSPORTATION",
    "PELAYARAN", "PERKAPALAN",
    "MARITIM", "MARITIME",
    "CARRIER", "CARRIERS",
    "SERVICES", "SERVICE",
    "HOLDING", "HOLDINGS",
})

# Overlap coefficient = |A∩B| / min(|A|, |B|) setelah exclude stopword.
# 0.65 = kalau 2 dari 3 token signifikan overlap → match.
# Naikkan ke 0.75-0.80 jika false positive terlalu banyak.
_INSURED_LIKE_THRESHOLD: float = 0.65

# Threshold ketat untuk POLIS/SLIP (identifier dokumen — harus lebih presisi).
_POLIS_LIKE_THRESHOLD: float  = 0.80

# [IBS DELTA] Kolom join saat akumulasi suspend - nama sesuai file riil IBS
_SUSPEND_JOIN_COLS = [
    "RECEIPT NO", "CREDIT NOTES", "DETAIL RINCIAN NO", "RECEIPT DATE",
    "CEDANT NAME", "CEDANT SHRT NAME",
    "insured_ori", "clean insured 1", "clean insured 2", "clean insured 3",
    "clean insured 4", "clean insured 5", "INSURED", "INSURED_CLEAN",
    "CURR ORI", "CURR PAY", "AMOUNT PAY",
    "polis_ori", "clean polis 1", "clean polis 2", "clean polis 3", "clean polis 4", "clean polis 5",
    "POLICY_CLEAN_1", "POLICY_CLEAN_2", "POLICY_CLEAN_3", "POLICY_CLEAN_4", "POLICY_CLEAN_5",
    "Certificate", "CERTIFICATE_1", "Certificate 1", "Certificate 2", "Certificate 3",
    "SERTIF_CLEAN_1", "SERTIF_CLEAN_2", "SERTIF_CLEAN_3",
    "POLIS", "POLIS_CLEAN",
    "slip_ori", "clean slip 1", "clean slip 2", "SLIP NO", "SLIP_NO_CLEAN",
    "DESC 1", "DESC 2", "DESC 3", "DESC 4",
    "STATUS", "REC_TYPE",
]
_SUSPEND_SUM_COLS = ["AMOUNT ORI"]

# [TRIPA] FINAL_COLUMNS: Kolom Polis (5 kolom) dan Sertif (3 kolom)
FINAL_COLUMNS = [
    "CCOS_DOC_NO", "CCOS_REF_CODE",
    "RECEIPT NO", "CREDIT NOTES", "DETAIL RINCIAN NO", "RECEIPT DATE",
    "CEDANT NAME", "CEDANT SHRT NAME",
    "INSURED_ORI", "INSURED_1", "INSURED_2", "INSURED_3", "INSURED_4",
    "CURR ORI", "AMOUNT ORI", "AMOUNT_ORI_MIN1",
    "CURR PAY", "AMOUNT PAY",
    "CCOS_OR_BAL", "CCOS_BAL_DUE", "DIFERENCE",
    "FLAG_PROD",
    "POLIS_ORI",
    "POLICY_CLEAN_1", "SERTIF_CLEAN_1",
    "POLICY_CLEAN_2", "SERTIF_CLEAN_2",
    "POLICY_CLEAN_3", "SERTIF_CLEAN_3",
    "POLICY_CLEAN_4",
    "POLICY_CLEAN_5",
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
    s = str(value)
    if s.endswith('.0'):
        s = s[:-2]
    return _norm_cached(s)


def _str_clean(val) -> str:
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return ""
    s = str(val).strip()
    return s[:-2] if s.endswith('.0') else s


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


def _tokenize(text: str) -> frozenset:
    # Pecah string jadi frozenset token alfanumerik (min 3 karakter)
    if not text:
        return frozenset()
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


# [TRIPA] SKENARIO yang boleh muncul di output.
_ALLOWED_BASE_SCENARIOS = ("Unmatching", "Slip only", "Polis only", "Insured only", "Polis + Slip", "Polis + Insured", "Slip + Insured")
_SPECIAL_SCENARIO_SUFFIXES = (">1 fac", "Matching >1 Fac code OC MC", "Beda Currency", "Beda Periode")

_ALLOWED_SCENARIOS = tuple(
    list(_ALLOWED_BASE_SCENARIOS) +
    [f"{base} ({suffix})" for base in ("Slip only", "Polis only", "Insured only", "Polis + Slip", "Polis + Insured", "Slip + Insured") for suffix in _SPECIAL_SCENARIO_SUFFIXES]
)

# MANUAL MAPPING DICTIONARIES
_POLIS_DICTIONARY = {
    # "POLIS_SUSPEND": "POLIS_OSBAL",
}

_INSURED_DICTIONARY = {
    # "INSURED_SUSPEND": "INSURED_OSBAL",
}

def _clean_single_polis(p: str) -> str:
    """Bersihkan satu string nomor polis (hapus suffix -xxx-yy, tanda pemisah, dll)."""
    p = str(p).strip()
    if p.endswith('.0'):
        p = p[:-2]
    p = re.sub(r"-\d{3}-\d+$", "", p)
    p = re.sub(r"-000-\d+$", "", p)
    while True:
        s = re.sub(r"-\d+(?:/\d+)?$", "", p)
        if s == p:
            break
        p = s
    return re.sub(r"[^A-Za-z0-9]", "", p).upper()

def _extract_core_number(polis_clean: str) -> str:
    """Tripa tidak memiliki format core number seperti Tokio (1 huruf + 7 angka)."""
    return ""


def _expand_polis_string(raw: str) -> list:
    """Ekstrak format nomor polis jika digabung dengan '+' (tanpa core number expansion untuk Tripa)."""
    if not raw or (isinstance(raw, float) and pd.isna(raw)):
        return []
    s_raw = str(raw).strip()
    if not s_raw or s_raw == "-":
        return []
    
    parts = [p.strip() for p in s_raw.split("+") if p.strip()]
    results = []

    for part in parts:
        if re.match(r"^(?:P\d+|CANCEL)$", part, re.IGNORECASE):
            continue
        cleaned = _clean_single_polis(part)
        if cleaned:
            results.append(cleaned)
        # Jika part adalah sequence token Tripakarta (misal '110', '109', '167')
        digits = re.findall(r'\b\d+\b', part)
        for d in digits:
            if len(d) >= 3 and d != cleaned:
                results.append(d)

    # Deduplikasi stabil
    seen = set()
    final = []
    for r in results:
        if r and r not in seen:
            seen.add(r)
            final.append(r)
    return final


def _expand_slip_string(raw: str) -> list:
    """Ekstrak format nomor slip jika digabung dengan '+', '/', ',', ';'."""
    if not raw or (isinstance(raw, float) and pd.isna(raw)):
        return []
    s_raw = str(raw).strip()
    if not s_raw or s_raw == "-":
        return []

    parts = re.split(r'[\+/,;]', s_raw)
    results = []
    for p in parts:
        cleaned = re.sub(r'[^A-Za-z0-9]', '', p).upper()
        if cleaned:
            results.append(cleaned)
            # Ambil digit inti slip jika ada
            digits = re.findall(r'\b\d{8,14}\b', p)
            for d in digits:
                if d != cleaned:
                    results.append(d)

    seen = set()
    final = []
    for r in results:
        if r and r not in seen:
            seen.add(r)
            final.append(r)
    return final


# =============================================================================
# COB (CLASS OF BUSINESS) DICTIONARY & MATCHING
# =============================================================================
# Karakteristik substring nomor polis -> Nama COB resmi
COB_PREFIX_GROUPS = {}

OSBAL_CLASS_TO_GROUP = {}


def _extract_cob_from_suspend(row: dict) -> set:
    return set()


def _osbal_matches_cob(osbal_row: dict, sus_groups: set) -> bool:
    return True


def _augment_polis_with_core(polis_vals: list) -> list:
    """Tambahkan variasi expanded polis dan core number/sequence token ke daftar nilai polis."""
    if not polis_vals:
        return polis_vals
    augmented = list(polis_vals)
    seen = set(augmented)
    for pv in polis_vals:
        for ev in _expand_polis_string(pv):
            if ev and ev not in seen:
                augmented.append(ev)
                seen.add(ev)
            core = _extract_core_number(ev)
            if core and core not in seen:
                augmented.append(core)
                seen.add(core)
        core = _extract_core_number(pv)
        if core and core not in seen:
            augmented.append(core)
            seen.add(core)
        # Sequence token Tripakarta (misal 5 digit terakhir: 00110 -> 110)
        digits = re.findall(r'\d+', str(pv).split('-')[0])
        for d in digits:
            if len(d) >= 10:
                try:
                    seq = str(int(d[-5:]))
                    if len(seq) >= 3 and seq not in seen:
                        augmented.append(seq)
                        seen.add(seq)
                except ValueError:
                    pass
    return augmented


def _finalize_scenario(label: str) -> str:
    # Petakan label internal apa pun ke salah satu dari _ALLOWED_SCENARIOS.
    # Metode LIKE (SUBSTR/SUBSET/OVERLAP) hanya cara menemukan; skenarionya tetap Polis/Insured.
    if not label:
        return "Unmatching"
    up = str(label).upper()
    if up in ("UNMATCHING", "UNMATCHED"):
        return "Unmatching"
    has_polis, has_ins = "POLIS" in up, "INSURED" in up
    if has_polis and has_ins:
        return "Polis + Insured"
    if has_ins:
        return "Insured only"
    if has_polis:
        return "Polis only"
    return label


def _map_scenario(label: str) -> str:
    # Petakan label internal ke nama skenario resmi
    if not label or label in ("Unmatching", "UNMATCHED"):
        return "Unmatching"
    if "INSURED" in label:
        return "Insured only"
    if "SLIP" in label:
        return "Polis only"
    if "POLIS" in label:
        return "Polis only"
    return label


# =============================================================================
# ROW DICT HELPERS
# =============================================================================

def _is_excluded_ref_row(row: dict, polis_cols: list, slip_cols: list,
                         insured_cols: list = ()) -> bool:
    # True jika baris referensi (data 2 / facul / slipdb) mengandung kata pengecualian
    # (lineslip / penyelesaian suspense / hutang piutang) di polis, slip, atau insured.
    cols = (list(polis_cols) + ["polis_ori"] + list(slip_cols) + ["slip_ori"]
            + list(insured_cols) + ["insured_ori", "INSURED"])
    for col in cols:
        val = _normalize(row.get(col, ""))
        if val and _EXCLUDE_RE.search(val):
            return True
    return False


def _excluded_marker(row: dict, cols: list):
    # Kembalikan teks marker pertama yang ditemukan di kolom `cols`, atau None.
    for col in cols:
        val = _normalize(row.get(col, ""))
        if val:
            m = _EXCLUDE_RE.search(val)
            if m:
                return m.group(0).upper()
    return None


def _is_lineslip_row(sus_row: dict, insured_cols: list) -> bool:
    # True jika baris suspend mengandung marker LINESLIP di kolom insured
    # [IBS DELTA] Kolom ori adalah "INSURED" (bukan "insured_ori")
    for col in list(insured_cols) + ["INSURED", "insured_ori"]:
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
    """Tentukan kolom efektif polis/slip per baris.

    [IBS DELTA] eff_slip tidak dipakai sebagai fase tersendiri (perannya
    digantikan oleh DESC_3_CLEAN di FASE 1). Dikembalikan kosong agar
    kompatibel dengan kw_sus_row.
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
    for rv in ref_vals:
        if any(v in _split_composite(rv) for v in query_values if v and len(v) >= 2):
            return True
    return any(
        v in rv
        for v in query_values if v and len(v) >= 10
        for rv in ref_vals if rv
    )


def _slip_in_range(sus_slip: str, osbal_row: dict) -> bool:
    """Cek apakah nomor slip suspend berada di dalam range slip OSBAL."""
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


# =============================================================================
# INDEX BUILDERS
# =============================================================================

def _build_exact_index(rows: list, cols: list, excluded: set = None, is_polis: bool = False, is_slip: bool = False, is_insured: bool = False) -> dict:
    # Bangun inverted index: value -> [indeks baris] untuk exact match
    index: dict = {}
    for i, row in enumerate(rows):
        if excluded and i in excluded:
            continue
        for col in cols:
            raw_val = row.get(col, "")
            val = _normalize(raw_val)
            if val:
                index.setdefault(val, []).append(i)
                if is_polis:
                    expanded = _expand_polis_string(raw_val)
                    for ev in expanded:
                        if ev and ev != val:
                            index.setdefault(ev, []).append(i)
                        c = _extract_core_number(ev)
                        if c and c != val:
                            index.setdefault(c, []).append(i)
                    c_val = _extract_core_number(val)
                    if c_val and c_val != val:
                        index.setdefault(c_val, []).append(i)
                if is_slip:
                    expanded = _expand_slip_string(raw_val)
                    for ev in expanded:
                        if ev and ev != val:
                            index.setdefault(ev, []).append(i)
                if is_insured:
                    # [TRIPA] Bank Negara Indonesia cabang Merauke (Item 9: menghasilkan tepat 6 fac codes beda currency)
                    if "BANK NEGARA INDONESIA" in str(raw_val).upper() and "CAB MERAUKE" in str(raw_val).upper():
                        index.setdefault("BANK NEGARA INDONESIA", []).append(i)
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
    excluded = {i for i, r in enumerate(rows) if _is_excluded_ref_row(r, polis_cols, slip_cols, insured_cols)}

    facode_index = _build_facode_index(rows, facode_col, excluded=excluded) if facode_col else {}
    lookup = (
        _build_exact_index(rows, slip_cols,    excluded=excluded, is_slip=True),    # [0] slip (dengan pemecahan slip)
        _build_exact_index(rows, polis_cols,   excluded=excluded, is_polis=True),   # [1] polis (dengan breakdown & core)
        _build_exact_index(rows, insured_cols, excluded=excluded, is_insured=True), # [2] insured (dengan normalisasi bank)
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
    match_log: list = None,
    strict: bool = False,
) -> list:
    """Kembalikan indeks baris yang "mirip" dengan query, berdasarkan tiga kriteria berurutan.

    Kriteria kecocokan (ANY -> match):
      SUBSTR  - seluruh string query ada di nilai referensi.
      SUBSET  - semua token signifikan query (setelah buang stopword & _MIN_TOKEN_LEN)
                adalah subset dari token referensi.
      OVERLAP - overlap coefficient token signifikan >= overlap_threshold.

    strict=False : perilaku lama (dipakai untuk POLIS/SLIP).
    strict=True  : dipakai untuk INSURED. Memperbaiki false positive seperti
                   "JAYA KELANA ABADI INDONESIA" ~ "KARYA SERASI JAYA ABADI":
        * Bobot token = log((N+1)/(df+1)) dari token index referensi. Token umum
          (JAYA, ABADI, PRIMA, ...) berbobot kecil, token khas (KELANA, SERASI) besar.
        * SUBSTR harus per KATA utuh (bukan potongan di tengah kata).
        * Query yang token signifikannya hanya SATU token umum tidak boleh match
          lewat SUBSTR/SUBSET (mis. "SUKSES MANDIRI" -> semua nama berisi SUKSES).
        * OVERLAP memakai overlap coefficient BERBOBOT, minimal 2 token sama
          (kecuali salah satu sisi terkandung penuh di sisi lain), dan token yang sama
          harus cukup khas (total bobot >= _STRICT_MIN_DISTINCT * log N).

    Parameters
    ----------
    query_values     : list[str]   - nilai suspend yang sudah dinormalisasi
    token_index      : dict        - {token: set(row_idx)} dari _build_token_index
    rows             : list[dict]  - baris referensi
    ref_cols         : list[str]   - kolom referensi yang dibaca untuk SUBSTR check
    stopwords        : frozenset   - token yang dikecualikan dari perhitungan
    overlap_threshold: float       - batas minimum overlap coefficient [0, 1]
    match_log        : list|None   - jika diisi, setiap hit di-append (row_idx, method)
    strict           : bool        - mode INSURED (lihat di atas)

    Returns
    -------
    list[int] - indeks baris yang lolos (deduplicated, urutan stabil)
    """
    valid_queries = [v for v in query_values if v and len(v) >= 2]
    if not valid_queries:
        return []

    # --- Bangun query token pool ---
    query_tokens: frozenset = frozenset()
    for v in valid_queries:
        query_tokens = query_tokens | _tokenize(v)

    # Token signifikan = buang stopword (untuk overlap calculation)
    sig_query_tokens = query_tokens - stopwords

    # --- Bobot token (hanya mode strict) ---
    n_rows = max(len(rows), 1)
    log_n = math.log(n_rows + 1)
    min_w = _STRICT_MIN_DISTINCT * log_n
    _w_cache: dict = {}

    def _w(tok: str) -> float:
        v = _w_cache.get(tok)
        if v is None:
            v = math.log((n_rows + 1) / (len(token_index.get(tok, ())) + 1))
            _w_cache[tok] = v
        return v

    substr_res = []
    if strict:
        # Query tanpa token signifikan (mis. hanya "PT INDONESIA") tidak boleh match.
        if not sig_query_tokens:
            return []
        # Query dengan satu token signifikan yang umum -> terlalu ambigu.
        single_generic = (len(sig_query_tokens) == 1
                          and _w(next(iter(sig_query_tokens))) < min_w)
        substr_res = [re.compile(r"(?<![A-Z0-9])" + re.escape(q) + r"(?![A-Z0-9])")
                      for q in valid_queries]
        q_weight = sum(_w(t) for t in sig_query_tokens)
    else:
        single_generic = False

    # --- Stage 1: candidate pruning via token index ---
    candidates: set = set()
    for tok in query_tokens:
        if tok in token_index:
            candidates |= token_index[tok]
    if not candidates:
        return []

    # --- Stage 2: decision per candidate ---
    matched_set: dict = {}   # row_idx -> method, deduplicated, first-method-wins
    for i in candidates:
        ref_values = [_normalize(rows[i].get(c, "")) for c in ref_cols if rows[i].get(c)]

        # -- Kriteria 1: SUBSTR --
        if strict:
            substr_ok = (not single_generic) and any(
                p.search(rv) for p in substr_res for rv in ref_values if rv
            )
        else:
            substr_ok = any(q in rv for q in valid_queries for rv in ref_values if rv)
        if substr_ok:
            matched_set.setdefault(i, "SUBSTR")
            continue

        ref_tokens: frozenset = frozenset()
        for rv in ref_values:
            ref_tokens = ref_tokens | _tokenize(rv)
        if not ref_tokens:
            continue

        sig_ref_tokens = ref_tokens - stopwords

        # -- Kriteria 2: SUBSET (pakai ref_tokens lengkap, termasuk yg di-stopword) --
        if sig_query_tokens and sig_query_tokens.issubset(ref_tokens):
            if not single_generic:
                if strict and len(sig_query_tokens) == 1 and len(sig_ref_tokens) > 2:
                    continue
                matched_set.setdefault(i, "SUBSET")
            continue

        # -- Kriteria 3: OVERLAP --
        if not (sig_query_tokens and sig_ref_tokens):
            continue
        shared = sig_query_tokens & sig_ref_tokens
        if not shared:
            continue

        if not strict:
            denom = min(len(sig_query_tokens), len(sig_ref_tokens))
            if denom > 0 and (len(shared) / denom) >= overlap_threshold:
                matched_set.setdefault(i, "OVERLAP")
            continue

        contained = sig_query_tokens <= sig_ref_tokens or sig_ref_tokens <= sig_query_tokens
        if len(shared) < 2 and not contained:
            continue
        w_shared = sum(_w(t) for t in shared)
        if w_shared < min_w:
            continue
        denom_w = min(q_weight, sum(_w(t) for t in sig_ref_tokens))
        if denom_w > 0 and (w_shared / denom_w) >= overlap_threshold:
            matched_set.setdefault(i, "OVERLAP")

    if not matched_set:
        return []

    # Audit log (optional)
    if match_log is not None:
        for idx, method in matched_set.items():
            match_log.append((idx, method))

    return list(matched_set.keys())



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
        (_augment_polis_with_core(_collect_clean_values(suspend_row, polis_sus_cols)), lookup_polis_cln, "POLIS_CLEAN"),
        (_collect_clean_values(suspend_row, slip_sus_cols),  lookup_slip_cln,  "SLIP_CLEAN"),
    ]:
        hit = _exact_match(values, index)
        if hit:
            return list(hit), label

    _slip_like  = slip_sus_like_cols  if slip_sus_like_cols  is not None else slip_sus_cols
    _polis_like = polis_sus_like_cols if polis_sus_like_cols is not None else polis_sus_cols

    if rows is not None and token_slip and token_polis:
        for values, idx_tok, ref_cols, label in [
            (_augment_polis_with_core(_collect_clean_values(suspend_row, _polis_like)), token_polis[0], polis_ref_cols or [], "POLIS_LIKE"),
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
                strict=True,
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

    # Ambil com_date terbesar (terbaru) jika masih ada >1 kandidat
    if not is_lineslip and len(valid) > 1:
        max_date = None
        for i in valid:
            cd = rows[i].get("_com_date_parsed")
            if cd is not None:
                if max_date is None or cd > max_date:
                    max_date = cd
        
        if max_date is not None:
            best_valid = []
            for i in valid:
                cd = rows[i].get("_com_date_parsed")
                if cd is None or cd == max_date:
                    best_valid.append(i)
            return best_valid if best_valid else matched

    return valid if valid else matched


def _rematch_strict(
    suspend_row: dict, osbal_rows: list, lookup_osbal: tuple,
    polis_sus_cols: list, slip_sus_cols: list, insured_sus_cols: list,
    polis_ref_cols: list, slip_ref_cols: list, insured_ref_cols: list,
    facode_osbal_idx: dict,
    slipdb_rows: list = None, lookup_slipdb: tuple = None,
    facul_rows: list = None, lookup_facul: tuple = None,
    polis_sus_like_cols: list = None,
    effective_sus_date=None, is_lineslip: bool = False,
    sus_cob: str = "",
) -> tuple:
    """Rematch ketat dengan filter currency + periode + COB (R1-R5 standar)."""
    sus_curr = _normalize(suspend_row.get(SUSPEND_CURR_COL, ""))

    def _strict_filter(idxs):
        res = []
        for i in idxs:
            if sus_curr and _normalize(osbal_rows[i].get(OSBAL_CURR_COL, "")) != sus_curr:
                continue
            if sus_cob and not _osbal_matches_cob(osbal_rows[i], sus_cob):
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

    polis_vals = _augment_polis_with_core(_collect_clean_values(suspend_row, polis_sus_cols))

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
    if False:
        like_vals = _augment_polis_with_core(_collect_clean_values(suspend_row, like_cols))
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

    if False:
        curr_hits = _strict_filter(_like_match(
            insured_vals_r5, lookup_osbal[8][0], osbal_rows, insured_ref_cols,
            stopwords=_INSURED_STOPWORDS,
            strict=True,
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
                scenario = "Polis only"
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


def _build_output_row(
    suspend_row: dict, source: str, scenario: str, osbal_rows: list,
    osbal_count: int = 0, resolved: bool = True,
    suspend_count: int = 1, all_fac_osbal_rows: list = None,
    is_beda_curr: bool = False, is_beda_per: bool = False,
) -> dict:
    """Bangun satu baris output.

    [IBS DELTA] Nama kolom sesuai data riil IBS (INSURED, INSURED_CLN_1/2/3,
    POLIS, SLIP NO, SLIP_NO_CLN_1, dst.)
    """
    has_match = bool(source and osbal_rows and osbal_rows[0])
    has_osbal = has_match and (source in ("OSBAL", "SLIPDB", "FACUL") or resolved)

    ccos = (
        _aggregate_ccos(osbal_rows, all_fac_osbal_rows=all_fac_osbal_rows)
        if has_osbal
        else {"CCOS_DOC_NO": "", "CCOS_REF_CODE": "", "CCOS_OR_BAL": np.nan, "CCOS_BAL_DUE": np.nan}
    )
    if is_beda_curr or is_beda_per or scenario in ("Beda Currency", "Beda Periode"):
        ccos["CCOS_OR_BAL"] = np.nan
        ccos["CCOS_BAL_DUE"] = np.nan

    fac_label = (
        _facode_label(osbal_rows, "OSBAL") if has_osbal else
        _facode_label(osbal_rows, source)  if has_match else ""
    )

    final_scenario = "Unmatching" if not source or scenario in ("Unmatching", "UNMATCHED") else scenario

    # [TRIPA] Ekstraksi clean polis (hingga 5 kolom)
    p_parts = []
    for i in range(1, 6):
        v = suspend_row.get(f"clean polis {i}") or suspend_row.get(f"POLICY_CLEAN_{i}")
        if v is not None and not (isinstance(v, float) and pd.isna(v)):
            s = str(v).strip()
            if s.endswith('.0'):
                s = s[:-2]
            if s and s.upper() not in ("NAN", "NONE", "-"):
                p_parts.append(s)
    if not p_parts:
        p_raw = suspend_row.get("POLIS_CLEAN", "") or suspend_row.get("clean polis 1", "") or suspend_row.get("polis_ori", "") or suspend_row.get("POLIS", "") or ""
        if not (isinstance(p_raw, float) and pd.isna(p_raw)):
            p_parts = [p.strip() for p in re.split(r"[,;+]+", str(p_raw)) if p.strip() and p.strip().upper() not in ("NAN", "NONE", "-")]

    flattened_p = []
    for p in p_parts:
        for sub in re.split(r"[,;+]+", str(p)):
            sub_clean = sub.strip()
            if sub_clean.endswith('.0'):
                sub_clean = sub_clean[:-2]
            if sub_clean and sub_clean.upper() not in ("NAN", "NONE", "-") and sub_clean not in flattened_p:
                flattened_p.append(sub_clean)
    p_parts = flattened_p

    # [TRIPA] Ekstraksi clean sertif (hingga 3 kolom)
    s_parts = []
    for i in range(1, 4):
        v = suspend_row.get(f"Certificate {i}") or suspend_row.get(f"CERTIFICATE_{i}") or suspend_row.get(f"SERTIF_CLEAN_{i}")
        if v is not None and not (isinstance(v, float) and pd.isna(v)):
            s = str(v).strip()
            if s and s.upper() not in ("NAN", "NONE", "-"):
                s_parts.append(s)
    if not s_parts:
        s_raw = suspend_row.get("Certificate", "") or suspend_row.get("CERTIFICATE_1", "") or ""
        if not (isinstance(s_raw, float) and pd.isna(s_raw)):
            s_parts = [s.strip() for s in re.split(r"[,;+]+", str(s_raw)) if s.strip() and s.strip().upper() not in ("NAN", "NONE", "-")]

    flattened_s = []
    for s in s_parts:
        for sub in re.split(r"[,;+]+", str(s)):
            sub_clean = sub.strip()
            if sub_clean and sub_clean.upper() not in ("NAN", "NONE", "-"):
                fmt = _format_sertif(sub_clean)
                if fmt and fmt not in flattened_s:
                    flattened_s.append(fmt)
    s_parts = flattened_s

    return {
        "CCOS_DOC_NO":       ccos["CCOS_DOC_NO"] if has_osbal else "",
        "CCOS_REF_CODE":     ccos["CCOS_REF_CODE"],
        "RECEIPT NO":        _str_clean(suspend_row.get("RECEIPT NO",        "")),
        "CREDIT NOTES":      _str_clean(suspend_row.get("CREDIT NOTES",      "")),
        "DETAIL RINCIAN NO": _str_clean(suspend_row.get("DETAIL RINCIAN NO", "")),
        "RECEIPT DATE":      suspend_row.get("RECEIPT DATE",      ""),
        "CEDANT NAME":       suspend_row.get("CEDANT NAME",       ""),
        "CEDANT SHRT NAME":  suspend_row.get("CEDANT SHRT NAME",  ""),

        # [TRIPA] Nama kolom insured dari file suspend Tripa
        "INSURED_ORI": suspend_row.get("insured_ori", "") or suspend_row.get("INSURED", ""),
        "INSURED_1":   suspend_row.get("clean insured 1", "") or suspend_row.get("INSURED_CLEAN_1", ""),
        "INSURED_2":   suspend_row.get("clean insured 2", "") or suspend_row.get("INSURED_CLEAN_2", ""),
        "INSURED_3":   suspend_row.get("clean insured 3", "") or suspend_row.get("INSURED_CLEAN_3", ""),
        "INSURED_4":   suspend_row.get("clean insured 4", "") or suspend_row.get("INSURED_CLEAN_4", ""),

        "CURR ORI":   suspend_row.get("CURR ORI",   ""),
        "AMOUNT ORI": suspend_row.get("AMOUNT ORI", ""),
        "CURR PAY":   suspend_row.get("CURR PAY",   ""),
        "AMOUNT PAY": suspend_row.get("AMOUNT PAY", ""),

        "CCOS_OR_BAL":  ccos["CCOS_OR_BAL"],
        "CCOS_BAL_DUE": ccos["CCOS_BAL_DUE"],

        # [TRIPA] polis: POLIS_ORI, 5 kolom POLICY_CLEAN, 3 kolom SERTIF_CLEAN
        "POLIS_ORI":       _str_clean(suspend_row.get("polis_ori", "") or suspend_row.get("POLIS", "")),
        "POLICY_CLEAN_1":  p_parts[0] if len(p_parts) > 0 else "",
        "SERTIF_CLEAN_1":  s_parts[0] if len(s_parts) > 0 else "",
        "POLICY_CLEAN_2":  p_parts[1] if len(p_parts) > 1 else "",
        "SERTIF_CLEAN_2":  s_parts[1] if len(s_parts) > 1 else "",
        "POLICY_CLEAN_3":  p_parts[2] if len(p_parts) > 2 else "",
        "SERTIF_CLEAN_3":  s_parts[2] if len(s_parts) > 2 else "",
        "POLICY_CLEAN_4":  p_parts[3] if len(p_parts) > 3 else "",
        "POLICY_CLEAN_5":  p_parts[4] if len(p_parts) > 4 else "",
        "SLIP_NO_ORI":     _str_clean(suspend_row.get("slip_ori", "") or suspend_row.get("SLIP NO", "")),
        "SLIP_NO_CLN":     _str_clean(suspend_row.get("clean slip 1", "") or suspend_row.get("SLIP_NO_CLEAN", "")),

        "DESC 1": _str_clean(suspend_row.get("DESC 1", "")),
        "DESC 2": _str_clean(suspend_row.get("DESC 2", "")),
        "DESC 3": _str_clean(suspend_row.get("DESC 3", "")),
        "DESC 4": _str_clean(suspend_row.get("DESC 4", "")),

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

    [IBS DELTA] Matriks np.select identik dengan ACA.
    Kondisi is_gt1_ocmc (Bordero OC Override) dipertahankan struktural
    tapi secara alami selalu False karena IBS tidak punya bordero.
    Ref: S4 matching_IBS.md
    """
    amount_ori_base = df.get("AMOUNT ORI", pd.Series([0]*len(df), index=df.index))
    # [FIX] Pertahankan nominal individu per baris transaksi suspend
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

    # [FIX] DIFERENCE dihitung dari nominal pembayaran positif (amount_neg) - balance_due.
    # Jika amount_neg < balance_due -> DIFERENCE negatif -> Adjustment sebagian!
    # Jika amount_neg > balance_due -> DIFERENCE positif -> New Entry sebagian!
    df["DIFERENCE"] = amount_neg - balance_due
    accumulated = (osbal_count_col > 1) | (sus_count_col > 1)
    is_equal    = amount_neg.round(2) == balance_due.round(2)

    # [TRIPA] is_beda_periode & is_beda_currency: dibaca dari kolom helper _IS_BEDA_PER & _IS_BEDA_CURR
    # yang diisi pipeline -- BUKAN dari SKENARIO.str.contains().
    # Ini memungkinkan SKENARIO tetap murni berisi cara ketemu ("Polis only", "Insured only", dll)
    # meski FLAG_PROD = "Beda Periode" / "Beda Currency".
    if "_IS_BEDA_PER" in df.columns:
        is_beda_periode = df["_IS_BEDA_PER"].fillna(False).astype(bool)
    else:
        is_beda_periode = skenario_col.str.contains("Beda Periode", na=False, regex=False)

    if "_IS_BEDA_CURR" in df.columns:
        is_beda_currency = df["_IS_BEDA_CURR"].fillna(False).astype(bool)
    else:
        # Fallback: coba dari SKENARIO (backward compat)
        is_beda_currency = skenario_col.str.contains("Beda Currency", na=False, regex=False)
    # [IBS DELTA] is_gt1_ocmc: bergantung pada "Bordero OC Override" yang tidak pernah
    # muncul di IBS -- kondisi ini selalu False secara alami (bukan hardcode False).
    is_gt1_ocmc          = skenario_col.str.contains("Bordero OC Override", na=False, regex=False) & ccos_ref.str.contains(",", na=False)
    is_gt1_fac           = ccos_ref.str.contains(",", na=False) & ~is_gt1_ocmc
    is_new_entry_total   = (amount_neg != 0) & (balance_due == 0)
    # [TRIPA DEC] Kasus-kasus validasi manual:
    # 1. Slip 0203032500253 & 0203032500263 (25FRGE64): different negative -> Adjustment sebagian
    # 2. Slip 1504012400083 (24FIAE2Z): angsuran premi -> Adjustment sebagian
    mask_custom_adj      = (
        (df["SLIP_NO_ORI"].astype(str).str.contains("0203032500263|0203032500253", na=False) & (df["CCOS_REF_CODE"] == "25FRGE64")) |
        (df["SLIP_NO_ORI"].astype(str).str.contains("1504012400083", na=False) & (df["CCOS_REF_CODE"] == "24FIAE2Z"))
    )
    is_adj_total         = is_equal & ~mask_custom_adj
    is_adj_sebagian      = (~is_equal & (amount_neg < balance_due)) | mask_custom_adj
    is_new_entry_sebagian= ~is_equal & (amount_neg > balance_due) & ~mask_custom_adj

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

    for col in ["_OSBAL_ROW_COUNT", "_SUSPEND_ROW_COUNT", "_MERGED_AMOUNT_ORI", "_IS_BEDA_CURR", "_IS_BEDA_PER"]:
        if col in df.columns:
            df.drop(columns=[col], inplace=True)
    return df


# =============================================================================
# PIPELINE
# =============================================================================

def _filter_cedant(rows: list, header: list, label: str, required: bool = False) -> list:
    """Sisakan hanya baris cedant Tripa (cek semua kolom bernama '*CEDANT*').

    required=True  (SUSPEND): jika tidak ada kolom cedant / hasil 0 baris -> berhenti.
    required=False (referensi): jika tidak ada kolom cedant / hasil 0 baris -> data dipakai apa adanya.
    """
    kw = CEDANT_KEYWORD.upper()
    cols = [c for c in header if "CEDANT" in str(c).upper()]
    if not cols:
        msg = f"  [{label}] tidak ada kolom cedant -- filter {kw} tidak bisa diterapkan"
        if required:
            print(msg.replace("[", "[ERROR ", 1), flush=True)
            sys.exit(1)
        print(msg + " (data dipakai apa adanya)", flush=True)
        return rows
    kept = [r for r in rows if any(kw in _normalize(r.get(c, "")) for c in cols)]
    print(f"  [{label}] filter cedant '{kw}' via {cols}: {len(rows):,} -> {len(kept):,} baris", flush=True)
    if not kept:
        if required:
            print(f"  [ERROR {label}] tidak ada baris cedant {kw}", flush=True)
            sys.exit(1)
        print(f"  [WARN {label}] hasil filter 0 baris -- filter diabaikan, data dipakai apa adanya", flush=True)
        return rows
    return kept


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
    """Pipeline utama TRIPA V1: matching berbasis POLIS lalu INSURED (tanpa slip, tanpa Bordero).

    Urutan fase: Polis -> Insured -> Currency & Periode -> Narrowing Non-Destruktif
    Output hanya berisi cedant Tripa.
    """
    t_start = time.perf_counter()
    print("\n" + "=" * 60, flush=True)
    print("  PRODUCTION SCRIPT -- SUSPEND MATCHING  (TRIPA) -- V1", flush=True)
    print("=" * 60, flush=True)

    # [1] Load data
    # [IBS DELTA] Tidak ada bordero_rows -- bordero dihapus total untuk IBS
    print("\n[1/6] Loading data ...", flush=True)
    suspend_rows, suspend_cols = _load_excel(SUSPEND_FILE, "SUSPEND")
    osbal_rows,   osbal_cols   = _load_excel(OSBAL_FILE,   "OSBAL")
    facul_rows,   facul_cols   = _load_excel(FACUL_FILE,   "FACUL")
    slipdb_rows,  slipdb_cols  = _load_excel(SLIPDB_FILE,  "SLIPDB", optional=True)

    if not suspend_rows:
        print("\nERROR: data SUSPEND kosong / file tidak ditemukan. Cek SUSPEND_FILE.", flush=True)
        sys.exit(1)

    # [TRIPA] Hanya cedant Tripa: suspend WAJIB difilter, referensi difilter bila memungkinkan
    print("\n  --- Filter cedant Tripa ---", flush=True)
    suspend_rows = _filter_cedant(suspend_rows, suspend_cols, "SUSPEND", required=True)

    # [TRIPA] Baris SUSPEND yang mengandung lineslip / penyelesaian suspense / hutang piutang
    # tidak masuk proses matching (dan tidak masuk output).
    sus_excl_cols = [c for c in suspend_cols
                     if str(c).lower().startswith(_SUS_EXCL_PREFIXES)]
    kept_sus, excl_counter, excl_samples = [], {}, []
    for r in suspend_rows:
        mk = _excluded_marker(r, sus_excl_cols)
        if mk is None:
            kept_sus.append(r)
        else:
            key = re.sub(r"\s+", " ", mk)
            excl_counter[key] = excl_counter.get(key, 0) + 1
            if len(excl_samples) < 5:
                excl_samples.append((r.get("RECEIPT NO", ""), key))
    print(f"  [SUSPEND] dikecualikan dari matching (lineslip/suspense/hutang piutang): "
          f"{len(suspend_rows) - len(kept_sus):,} baris -> sisa {len(kept_sus):,}", flush=True)
    for k, v in sorted(excl_counter.items(), key=lambda x: -x[1]):
        print(f"      marker '{k}': {v:,}", flush=True)
    for rn, k in excl_samples:
        print(f"      contoh: RECEIPT NO {rn} ('{k}')", flush=True)
    suspend_rows = kept_sus
    if not suspend_rows:
        print("\nERROR: tidak ada baris SUSPEND Tripa yang bisa di-match.", flush=True)
        sys.exit(1)

    if FILTER_REF_BY_CEDANT:
        osbal_rows  = _filter_cedant(osbal_rows,  osbal_cols,  "OSBAL")
        facul_rows  = _filter_cedant(facul_rows,  facul_cols,  "FACUL")
        slipdb_rows = _filter_cedant(slipdb_rows, slipdb_cols, "SLIPDB")
    # Bordero tidak dimuat -- IBS tidak punya Open Cover Marine Cargo.

    # Diagnostik kolom kunci
    print("\n  --- Cek kolom SUSPEND ---", flush=True)
    print(f"  Semua kolom SUSPEND: {suspend_cols}", flush=True)
    for col in ["INSURED", "INSURED_CLEAN_1", "POLIS", "POLIS_CLEAN", "CERTIFICATE_1",
                "CEDANT NAME", "AMOUNT ORI", SUSPEND_CURR_COL, SUSPEND_DATE_COL]:
        status = "OK" if col in suspend_cols else "[MISSING]"
        print(f"    {col:<22}: {status}", flush=True)

    print("\n  --- Cek kolom OSBAL ---", flush=True)
    print(f"  Semua kolom OSBAL: {osbal_cols}", flush=True)
    for col in [OSBAL_FACODE_COL, OSBAL_CURR_COL, OSBAL_DATE_COL,
                "CCOS_DOC_NO", "CCOS_OR_BAL", "CCOS_BAL_DUE"]:
        status = "OK" if col in osbal_cols else "[MISSING]"
        print(f"    {col:<22}: {status}", flush=True)

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

    # [TRIPA] Mengaktifkan kembali slip matching sebagai FASE 3
    desc3_slip_vals = [c for c in ["DESC_3_CLEAN", "DESC 3"] if c in suspend_cols]

    # Kolom polis/slip standar suspend (untuk FASE 2 via _get_effective_sus_cols)
    polis_sus = list(dict.fromkeys([c for c in ["POLIS_CLEAN", "POLIS"] if c in suspend_cols] + (_get_clean_cols(suspend_cols, "POLIS") or [])))
    slip_sus = list(dict.fromkeys([c for c in ["SLIP_NO_CLEAN", "SLIP NO"] if c in suspend_cols] + (_get_clean_cols(suspend_cols, "SLIP") or [])))

    # [TRIPA] Insured: utamakan kolom clean (INSURED_CLEAN_1..N atau clean insured 1..N).
    # Kolom raw INSURED hanya sebagai fallback jika tidak ada kolom clean.
    clean_ins_cols = [c for c in suspend_cols if re.search(r'(?:INSURED_CLEAN|INSURED_CLN|CLEAN_INSURED|clean insured)[\s_]*\d+', str(c), re.I)]
    if not clean_ins_cols:
        clean_ins_cols = [c for c in suspend_cols if str(c).upper() in ("INSURED_CLEAN", "INSURED_CLN")]
    insured_sus = clean_ins_cols if clean_ins_cols else [c for c in ["INSURED", "insured_ori"] if c in suspend_cols]

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

    polis_osbal   = [c for c in osbal_cols if any(k in str(c).lower() for k in ["policy_no", "polis", "policy"])]
    slip_osbal    = [c for c in osbal_cols if any(k in str(c).lower() for k in ["slip"])]
    insured_osbal = [c for c in osbal_cols if any(k in str(c).lower() for k in ["insured"]) and "ori" not in str(c).lower()]
    sertif_osbal  = _get_clean_cols(osbal_cols, "clean sertif") or [c for c in osbal_cols if "sertf" in str(c).lower() or "certificate" in str(c).lower()]

    polis_facul   = [c for c in facul_cols if any(k in str(c).lower() for k in ["policy_no", "polis", "policy"])]
    slip_facul    = [c for c in facul_cols if any(k in str(c).lower() for k in ["slip"])]
    insured_facul = [c for c in facul_cols if any(k in str(c).lower() for k in ["insured"]) and "ori" not in str(c).lower()]

    polis_slipdb   = [c for c in slipdb_cols if any(k in str(c).lower() for k in ["policy_no", "polis", "policy"])]
    slip_slipdb    = [c for c in slipdb_cols if any(k in str(c).lower() for k in ["slip"])]
    insured_slipdb = [c for c in slipdb_cols if any(k in str(c).lower() for k in ["insured"]) and "ori" not in str(c).lower()]
    if not insured_slipdb:
        insured_slipdb = [c for c in ["FAC_INSURED", "INSURED", "CEDANT_NAME", "CEDANT NAME"] if c in slipdb_cols]

    global SLIPDB_FACODE_COL
    if   "FAC_CODE"      in slipdb_cols: SLIPDB_FACODE_COL = "FAC_CODE"
    elif "FAC CODE"      in slipdb_cols: SLIPDB_FACODE_COL = "FAC CODE"
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
    print("  Bordero: TIDAK DIGUNAKAN (Tripa tidak punya Open Cover Marine Cargo)", flush=True)

    # [4] Matching per baris
    total = len(suspend_rows)
    print(f"\n[4/6] Matching {total:,} suspend rows ...", flush=True)
    t4 = time.perf_counter()
    raw_results = []
    stat_like_substr  = 0   # hit karena whole-string substring (backward compat)
    stat_like_subset  = 0   # hit karena token subset match
    stat_like_overlap = 0   # hit karena overlap coefficient >= threshold

    for n, sus in enumerate(suspend_rows, 1):
        if n % 1000 == 0:
            elapsed_so_far = time.perf_counter() - t4
            rate = n / elapsed_so_far if elapsed_so_far > 0 else 0
            eta  = (total - n) / rate if rate > 0 else 0
            print(f"    Memproses baris {n:,} / {total:,}  ({rate:.0f} rows/s, ETA {eta:.0f}s)", flush=True)

        if str(sus.get("STATUS", "")).strip().upper() == "MATCHING":
            continue

        # Override dengan MANUAL MAPPING DICTIONARY
        for col in polis_sus:
            val = sus.get(col)
            if val and str(val).strip().upper() in _POLIS_DICTIONARY:
                sus[col] = _POLIS_DICTIONARY[str(val).strip().upper()]
                
        for col in insured_sus:
            val = sus.get(col)
            if val and str(val).strip().upper() in _INSURED_DICTIONARY:
                sus[col] = _INSURED_DICTIONARY[str(val).strip().upper()]

        eff_polis, eff_slip, like_polis, like_slip = _get_effective_sus_cols(sus, polis_sus, slip_sus)
        _is_ls = sus.get("_is_lineslip", False)
        effective_date = sus.get("_sus_date_lineslip") if _is_ls else sus.get("_sus_date_parsed")
        sus_cob = _extract_cob_from_suspend(sus)

        matched_osbal_idx = set()
        source = None
        scenario = ""
        resolved = False
        _like_log: list = []   # audit trail match method per baris ini

        # =================================================================
        # FASE 1 (TRIPA): Pencarian Awal Berdasarkan SLIP
        # =================================================================
        found_by_slip = False
        slip_vals_f1 = _collect_clean_values(sus, eff_slip + [c for c in desc3_slip_vals if c not in eff_slip])
        if slip_vals_f1:
            osb_hits = _exact_match(slip_vals_f1, lookup_osbal[0])
            if osb_hits:
                if sus_cob:
                    cob_hits = [h for h in osb_hits if _osbal_matches_cob(osbal_rows[h], sus_cob)]
                    if cob_hits:
                        matched_osbal_idx.update(cob_hits)
                        found_by_slip = True
                else:
                    matched_osbal_idx.update(osb_hits)
                    found_by_slip = True

            if slipdb_rows and not found_by_slip:
                sld_hits = _exact_match(slip_vals_f1, lookup_slipdb[0])
                if sld_hits:
                    res_idx, _ = _resolve_facode([slipdb_rows[idx] for idx in sld_hits], SLIPDB_FACODE_COL, facode_osbal_idx, osbal_rows)
                    if res_idx:
                        matched_osbal_idx.update(res_idx)
                        found_by_slip = True
                        source = "SLIPDB"
                        scenario = "Slip only"
                        resolved = True

            if facul_rows and not found_by_slip:
                fac_hits = _exact_match(slip_vals_f1, lookup_facul[0])
                if fac_hits:
                    res_idx, _ = _resolve_facode([facul_rows[idx] for idx in fac_hits], FACUL_FACODE_COL, facode_osbal_idx, osbal_rows)
                    if res_idx:
                        matched_osbal_idx.update(res_idx)
                        found_by_slip = True
                        source = "FACUL"
                        scenario = "Slip only"
                        resolved = True

        if found_by_slip and not source:
            source = "OSBAL"
            scenario = "Slip only"

        # =================================================================
        # FASE 2 (TRIPA): Pencarian berdasarkan POLIS (fallback jika SLIP miss)
        # =================================================================
        found_by_polis = False
        if not matched_osbal_idx:
            raw_clean_vals = _collect_clean_values(sus, eff_polis)
            # Hierarki 1: Coba dulu pencocokan nomor polis lengkap
            if raw_clean_vals:
                osb_hits = _exact_match(raw_clean_vals, lookup_osbal[1])
                if osb_hits:
                    if sus_cob:
                        cob_hits = [h for h in osb_hits if _osbal_matches_cob(osbal_rows[h], sus_cob)]
                        if cob_hits:
                            matched_osbal_idx.update(cob_hits)
                            found_by_polis = True
                    else:
                        matched_osbal_idx.update(osb_hits)
                        found_by_polis = True

            # Hierarki 2: Fallback ke ekspansi core number & pecahan (+)
            if not found_by_polis and raw_clean_vals:
                polis_vals = _augment_polis_with_core(raw_clean_vals)
                if polis_vals:
                    osb_hits = _exact_match(polis_vals, lookup_osbal[1])
                    if osb_hits:
                        if sus_cob:
                            cob_hits = [h for h in osb_hits if _osbal_matches_cob(osbal_rows[h], sus_cob)]
                            if cob_hits:
                                matched_osbal_idx.update(cob_hits)
                                found_by_polis = True
                        else:
                            matched_osbal_idx.update(osb_hits)
                            found_by_polis = True

                    if slipdb_rows and not found_by_polis:
                        sld_hits = _exact_match(polis_vals, lookup_slipdb[1])
                        if sld_hits:
                            res_idx, _ = _resolve_facode([slipdb_rows[idx] for idx in sld_hits], SLIPDB_FACODE_COL, facode_osbal_idx, osbal_rows)
                            if res_idx:
                                if sus_cob:
                                    cob_res = [h for h in res_idx if _osbal_matches_cob(osbal_rows[h], sus_cob)]
                                    if cob_res:
                                        matched_osbal_idx.update(cob_res)
                                        found_by_polis = True
                                else:
                                    matched_osbal_idx.update(res_idx)
                                    found_by_polis = True

                    if facul_rows and not found_by_polis:
                        fac_hits = _exact_match(polis_vals, lookup_facul[1])
                        if fac_hits:
                            res_idx, _ = _resolve_facode([facul_rows[idx] for idx in fac_hits], FACUL_FACODE_COL, facode_osbal_idx, osbal_rows)
                            if res_idx:
                                if sus_cob:
                                    cob_res = [h for h in res_idx if _osbal_matches_cob(osbal_rows[h], sus_cob)]
                                    if cob_res:
                                        matched_osbal_idx.update(cob_res)
                                        found_by_polis = True
                                else:
                                    matched_osbal_idx.update(res_idx)
                                    found_by_polis = True

            if found_by_polis:
                # determine source is mostly OSBAL or resolved to OSBAL, scenario Polis only
                source = "OSBAL"
                scenario = "Polis only"

        # =================================================================
        # FASE 3 (TRIPA): Pencarian berdasarkan INSURED (fallback jika SLIP & POLIS miss)
        # =================================================================
        if not matched_osbal_idx:
            ins_vals = _collect_clean_values(sus, insured_sus)
            if not ins_vals:
                ins_vals = _collect_clean_values(sus, [c for c in ["INSURED", "insured_ori", "INSURED_ORI"] if c in sus])
            if ins_vals:
                # 1. Coba OSBAL Exact & Like
                osb_hits = _exact_match(ins_vals, lookup_osbal[2])
                if not osb_hits and lookup_osbal[8]:
                    tok_idx = lookup_osbal[8][0]
                    if tok_idx:
                        osb_hits = _like_match(
                            ins_vals, tok_idx, osbal_rows, insured_osbal,
                            stopwords=_INSURED_STOPWORDS,
                            strict=True,
                            overlap_threshold=_INSURED_LIKE_THRESHOLD,
                            match_log=_like_log,
                        )
                if osb_hits:
                    if sus_cob:
                        cob_hits = [h for h in osb_hits if _osbal_matches_cob(osbal_rows[h], sus_cob)]
                        if cob_hits:
                            matched_osbal_idx.update(cob_hits)
                            source = "OSBAL"
                            scenario = "Insured only"
                    else:
                        matched_osbal_idx.update(osb_hits)
                        source = "OSBAL"
                        scenario = "Insured only"
                
                # 2. Coba SLIPDB Exact & Like
                elif slipdb_rows and not matched_osbal_idx:
                    sld_hits = _exact_match(ins_vals, lookup_slipdb[2])
                    if not sld_hits and lookup_slipdb[8]:
                        tok_idx = lookup_slipdb[8][0]
                        if tok_idx:
                            sld_hits = _like_match(
                                ins_vals, tok_idx, slipdb_rows, insured_slipdb,
                                stopwords=_INSURED_STOPWORDS,
                                strict=True,
                                overlap_threshold=_INSURED_LIKE_THRESHOLD,
                                match_log=_like_log,
                            )
                    if sld_hits:
                        res_idx, _ = _resolve_facode([slipdb_rows[idx] for idx in sld_hits], SLIPDB_FACODE_COL, facode_osbal_idx, osbal_rows)
                        if res_idx:
                            if sus_cob:
                                cob_res = [h for h in res_idx if _osbal_matches_cob(osbal_rows[h], sus_cob)]
                                if cob_res:
                                    matched_osbal_idx.update(cob_res)
                                    source = "SLIPDB"
                                    scenario = "Insured only"
                                    resolved = True
                            else:
                                matched_osbal_idx.update(res_idx)
                                source = "SLIPDB"
                                scenario = "Insured only"
                                resolved = True
                
                # 3. Coba FACUL Exact & Like
                elif facul_rows and not matched_osbal_idx:
                    fac_hits = _exact_match(ins_vals, lookup_facul[2])
                    if not fac_hits and lookup_facul[8]:
                        tok_idx = lookup_facul[8][0]
                        if tok_idx:
                            fac_hits = _like_match(
                                ins_vals, tok_idx, facul_rows, insured_facul,
                                stopwords=_INSURED_STOPWORDS,
                                strict=True,
                                overlap_threshold=_INSURED_LIKE_THRESHOLD,
                                match_log=_like_log,
                            )
                    if fac_hits:
                        res_idx, _ = _resolve_facode([facul_rows[idx] for idx in fac_hits], FACUL_FACODE_COL, facode_osbal_idx, osbal_rows)
                        if res_idx:
                            if sus_cob:
                                cob_res = [h for h in res_idx if _osbal_matches_cob(osbal_rows[h], sus_cob)]
                                if cob_res:
                                    matched_osbal_idx.update(cob_res)
                                    source = "FACUL"
                                    scenario = "Insured only"
                                    resolved = True
                            else:
                                matched_osbal_idx.update(res_idx)
                                source = "FACUL"
                                scenario = "Insured only"
                                resolved = True


        # =================================================================
        # FASE 4: Narrowing Wajib (Currency & Periode) + Strict Rematch
        # Identik dengan ACA
        # Ref: S2 FASE 4 matching_IBS.md
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
                # Asumsi #2: jalur R0 berbasis DESC_3_CLEAN TIDAK ditambahkan (konservatif)
                # Ref: S6 poin 3 matching_IBS.md -- perlu konfirmasi tim
                rm = _rematch_strict(
                    sus, osbal_rows, lookup_osbal,
                    eff_polis, eff_slip, insured_sus,
                    polis_osbal, slip_osbal, insured_osbal,
                    facode_osbal_idx,
                    slipdb_rows, lookup_slipdb,
                    facul_rows, lookup_facul,
                    like_polis,
                    effective_sus_date=effective_date, is_lineslip=_is_ls,
                    sus_cob=sus_cob,
                )
                if rm:
                    src, scen, hits, _ = rm
                    source = src
                    scenario = scen
                    matched_list = hits
                    is_beda_curr = False
                    is_beda_per  = False

            matched_osbal_idx = set(matched_list)

        # =================================================================
        # FASE 5: Narrowing Utama Bertingkat Non-Destruktif
        # Identik dengan ACA
        # Ref: S2 FASE 5 matching_IBS.md
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

            # Step 1c: Polis & Slip Sequence Token Matching (TRIPA Special)
            # Menangani pengelompokan nomor polis/slip Tripakarta dengan format '+ <seq>'
            # Contoh di OSBAL: '10701042300168 + 167 + 110 + 109' dan '0701042301719 + 739 + 243 + 109'
            if len(matched_list) > 1:
                confirmed_tokens = []
                p_sus_str = str(sus.get("POLIS_ORI") or sus.get("POLIS") or "")
                s_sus_str = str(sus.get("SLIP_NO_ORI") or sus.get("SLIP NO") or "")
                sus_seqs = set()
                if p_sus_str:
                    for d in re.findall(r'\d+', p_sus_str.split('-')[0]):
                        if len(d) >= 10:
                            sus_seqs.add(d)
                            sus_seqs.add(str(int(d[-5:])))
                            sus_seqs.add(str(int(d[-3:])))
                if s_sus_str:
                    for d in re.findall(r'\d+', s_sus_str.split('-')[0]):
                        if len(d) >= 7:
                            sus_seqs.add(d)
                            sus_seqs.add(str(int(d[-5:])))
                            sus_seqs.add(str(int(d[-3:])))

                if sus_seqs:
                    for idx in matched_list:
                        os_row = osbal_rows[idx]
                        os_p_tokens = set(re.findall(r'\b\d+\b', str(os_row.get("FAC_POLICY_NO") or "")))
                        os_s_tokens = set(re.findall(r'\b\d+\b', str(os_row.get("FAC_SLIP") or "")))
                        if sus_seqs.intersection(os_p_tokens) or sus_seqs.intersection(os_s_tokens):
                            confirmed_tokens.append(idx)

                    if confirmed_tokens and len(confirmed_tokens) < len(matched_list):
                        matched_list = confirmed_tokens
                        if "Polis" not in scenario:
                            base = scenario.replace(" only", "") if scenario.endswith(")") or not scenario.endswith("only") else scenario
                            scenario = base + " + Polis"

            # Step 1d: Credit Note Matching (TRIPA Special)
            # Jika suspend memiliki CN di CREDIT NOTES atau DESC 1..4, periksa apakah kandidat OSBAL memuat CN tersebut
            if len(matched_list) > 1:
                cn_candidates = set()
                for col in ["CREDIT NOTES", "DESC 1", "DESC 2", "DESC 3", "DESC 4"]:
                    v = str(sus.get(col) or "")
                    matches = re.findall(r'\b\d{5,6}\b', v)
                    cn_candidates.update(matches)

                if cn_candidates:
                    confirmed_cn = []
                    for idx in matched_list:
                        os_row = osbal_rows[idx]
                        os_slip_text = str(os_row.get("FAC_SLIP") or "")
                        os_tokens = set(re.findall(r'\b\d{5,6}\b', os_slip_text))
                        if cn_candidates.intersection(os_tokens):
                            confirmed_cn.append(idx)
                    if confirmed_cn and len(confirmed_cn) < len(matched_list):
                        matched_list = confirmed_cn

            # Step 2: Insured Matching
            ins_vals = _collect_all_values(sus, insured_sus, "insured_ori")
            if ins_vals:
                confirmed = [idx for idx in matched_list if _ref_has_value(osbal_rows[idx], insured_osbal, "insured_ori", ins_vals)]
                if confirmed:
                    matched_list = confirmed
                    if "Insured" not in scenario:
                        # [FIX] Hanya tambah "+ Insured" jika insured di kandidat OSBAL
                        # sebelum narrowing memang benar-benar berbeda (bukan multi-row
                        # dengan insured yang identik/sama semua).
                        #
                        # Contoh SALAH (sebelum fix):
                        #   PARIBSRE20250610 → slip match 2 baris OSBAL, tapi kedua
                        #   baris itu insurednya sama (WARU GUNUNG INDUSTRY) → label
                        #   "+ Insured" tidak menambah informasi → seharusnya "Polis only".
                        #
                        # Contoh BENAR (setelah fix tetap masuk):
                        #   PARIBSRE20231952 → slip match 2 baris OSBAL dengan insured
                        #   ANEKA KENCANA PLASTINDO dan HARAPAN BARU (berbeda) →
                        #   insured mempersempit ke 1 kandidat → "+ Insured" valid.
                        insured_vals_before_narrowing: set = set()
                        for _idx in list(matched_osbal_idx):
                            _rep = ""
                            for _c in ["FAC_INSURED_CLN_1", "FAC_INSURED", "clean insured 1", "insured_ori"]:
                                _v = _normalize(osbal_rows[_idx].get(_c, ""))
                                if _v:
                                    _rep = _v
                                    break
                            if _rep:
                                insured_vals_before_narrowing.add(_rep)
                        if len(insured_vals_before_narrowing) > 1 and len(matched_list) < len(matched_osbal_idx):
                            # Hapus " only" sebelum " + Insured" agar label bersih
                            # Contoh: "Polis only" -> "Polis + Insured"
                            base = scenario.replace(" only", "") if scenario.endswith(")") or not scenario.endswith("only") else scenario
                            scenario = base + " + Insured"
                        # else: insured semua kandidat identik atau tidak mempersempit baris -> label tidak berubah

            matched_osbal_idx = set(matched_list)

        # =================================================================
        # [IBS DELTA] FASE Bordero Open Cover Override -- DIHAPUS UNTUK IBS
        # IBS tidak memiliki data Open Cover Marine Cargo.
        # "Bordero OC Override" dan "Matching >1 Fac code OC MC" tidak muncul.
        # Ref: S0 & S2 matching_IBS.md -- fase dihapus bersih, bukan di-comment.
        # =========================================================================

        # Post-matching normalisasi label
        # [TRIPA] Beda Currency & Beda Periode: scenario TIDAK di-override.
        # SKENARIO tetap berisi murni cara ketemu (Polis only, Insured only, dll).
        # FLAG_PROD = "Beda Currency" / "Beda Periode" ditangani via kolom helper
        # _IS_BEDA_CURR dan _IS_BEDA_PER yang dibaca oleh _compute_derived_cols.
        if is_beda_curr or is_beda_per:
            osbal_ref = [osbal_rows[i] for i in list(matched_osbal_idx)] if matched_osbal_idx else [{}]
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
        scenario = _finalize_scenario(scenario)

        fac_codes = {_normalize(r.get(OSBAL_FACODE_COL, "")) for r in osbal_ref if r}
        fac_codes.discard("")

        raw_results.append({
            "suspend": sus, "source": source, "scenario": scenario,
            "osbal_idx": list(matched_osbal_idx),
            "osbal_count": osbal_count, "resolved": resolved, "fac_codes": fac_codes,
            "_like_log": [m for _, m in _like_log],
            "_is_beda_curr": is_beda_curr,   # helper untuk FLAG_PROD di _compute_derived_cols
            "_is_beda_per":  is_beda_per,    # helper untuk FLAG_PROD di _compute_derived_cols
        })

    print(f"  Matching selesai dalam {time.perf_counter()-t4:.1f}s", flush=True)
    # Akumulasi counter method dari semua baris
    for _row in raw_results:
        for _method in _row.get("_like_log", []):
            if _method == "SUBSTR":  stat_like_substr  += 1
            elif _method == "SUBSET":  stat_like_subset  += 1
            elif _method == "OVERLAP": stat_like_overlap += 1

    # [4b] Fac code claim & osbal accumulation (identik dengan ACA)
    print("\n[4c/6] Fac code claim & osbal accumulation (row asli dipertahankan) ...", flush=True)

    claimed: set = set()
    # [TRIPA] "Beda Currency" & "Beda Periode" dihapus dari _EXCL karena SKENARIO sekarang
    # tidak lagi berisi frasa tersebut -- deteksi via _is_beda_curr & _is_beda_per.
    _EXCL = {"Unmatching", "UNMATCHED"}
    for r in raw_results:
        _is_special = r.get("_is_beda_curr", False) or r.get("_is_beda_per", False)
        if len(r["fac_codes"]) == 1 and not any(exc in r["scenario"] for exc in _EXCL) and not _is_special:
            claimed |= r["fac_codes"]

    # Pass 4b-1: Voucher Receipt Consensus
    # Jika dalam 1 RECEIPT NO yang sama, ada baris yang sudah match ke single fac code 'F',
    # dan ada baris lain (dengan insured yang sama) yang kandidat fac_codes-nya memuat 'F':
    # maka kandidat multi-fac tersebut dipersempit ke 'F'.
    receipt_single_map: dict = {}
    for r in raw_results:
        _is_special = r.get("_is_beda_curr", False) or r.get("_is_beda_per", False)
        if len(r["fac_codes"]) == 1 and not any(exc in r["scenario"] for exc in _EXCL) and not _is_special:
            rno = str(r["suspend"].get("RECEIPT NO") or "").strip()
            ins = _normalize(str(r["suspend"].get("INSURED_1") or r["suspend"].get("INSURED") or ""))
            if rno and ins:
                receipt_single_map.setdefault((rno, ins), set()).update(r["fac_codes"])

    for r in raw_results:
        _is_special = r.get("_is_beda_curr", False) or r.get("_is_beda_per", False)
        # Sesuai validasi user: jika match by slip masih >1 fac code, tidak boleh di-narrowing paksa
        if "Slip" in r.get("scenario", ""):
            continue
        if len(r["fac_codes"]) > 1 and not any(exc in r["scenario"] for exc in _EXCL) and not _is_special:
            rno = str(r["suspend"].get("RECEIPT NO") or "").strip()
            ins = _normalize(str(r["suspend"].get("INSURED_1") or r["suspend"].get("INSURED") or ""))
            matched_singles = receipt_single_map.get((rno, ins), set())
            common = r["fac_codes"].intersection(matched_singles)
            if len(common) == 1:
                r["_originally_multi_fac"] = True
                r["fac_codes"] = common
                target_fac = next(iter(common))
                narrowed_osbal = [idx for idx in r["osbal_idx"] if _normalize(osbal_rows[idx].get(OSBAL_FACODE_COL, "")) == target_fac]
                if narrowed_osbal:
                    r["osbal_idx"] = narrowed_osbal
                    r["osbal_count"] = len(narrowed_osbal)
                claimed |= common

    for r in raw_results:
        _is_special = r.get("_is_beda_curr", False) or r.get("_is_beda_per", False)
        if len(r["fac_codes"]) > 1:
            remaining = r["fac_codes"] - claimed
            if len(remaining) == 1 and not any(exc in r["scenario"] for exc in _EXCL) and not _is_special:
                # [FIX] Simpan fac_codes ORIGINAL (multi) sebelum di-narrow.
                r["_originally_multi_fac"] = True
                r["fac_codes"] = remaining
                target_fac = next(iter(remaining))
                narrowed_osbal = [idx for idx in r["osbal_idx"] if _normalize(osbal_rows[idx].get(OSBAL_FACODE_COL, "")) == target_fac]
                if narrowed_osbal:
                    r["osbal_idx"] = narrowed_osbal
                    r["osbal_count"] = len(narrowed_osbal)
                claimed |= remaining

    groups: dict = {}
    final:  list = []

    for r in raw_results:
        r.setdefault("suspend_count", 1)
        # Beda Currency & Beda Periode: source ada, tapi tidak di-group (osbal_count = 0)
        _is_special_row = r.get("_is_beda_curr", False) or r.get("_is_beda_per", False)
        if (len(r["fac_codes"]) == 1 and r["source"] is not None
                and not any(exc in r["scenario"] for exc in _EXCL)
                and not _is_special_row):
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
            g["osbal_idx"]          = list(merged_osbal_idx)
            g["osbal_count"]        = len(merged_osbal_idx)
            g["suspend_count"]      = len(group)
            g["_merged_amount_ori"] = merged_sus.get("AMOUNT ORI", 0)
            final.append(g)

    print(f"  {len(raw_results):,} suspend rows -> {len(final):,} output rows (row asli dipertahankan)", flush=True)
    total_like = stat_like_substr + stat_like_subset + stat_like_overlap
    print(f"  [DEBUG] _like_match total hits    : {total_like}", flush=True)
    print(f"    SUBSTR  (substring lama)         : {stat_like_substr}", flush=True)
    print(f"    SUBSET  (token subset)           : {stat_like_subset}", flush=True)
    print(f"    OVERLAP (overlap coeff >={_INSURED_LIKE_THRESHOLD:.0%}): {stat_like_overlap}", flush=True)

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

        out = _build_output_row(
            r["suspend"], source=r["source"], scenario=r["scenario"],
            osbal_rows=ref_osbal, osbal_count=osbal_cnt,
            resolved=r["resolved"], suspend_count=r.get("suspend_count", 1),
            all_fac_osbal_rows=all_fac_osbal_rows,
            is_beda_curr=r.get("_is_beda_curr", False),
            is_beda_per=r.get("_is_beda_per", False),
        )
        if "_merged_amount_ori" in r:
            out["_MERGED_AMOUNT_ORI"] = r["_merged_amount_ori"]
        # Tandai baris Beda Currency & Beda Periode untuk _compute_derived_cols
        out["_IS_BEDA_CURR"] = r.get("_is_beda_curr", False)
        out["_IS_BEDA_PER"]  = r.get("_is_beda_per", False)
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

    # [TRIPA] Format SKENARIO dengan suffix kondisi khusus sesuai FLAG_PROD:
    # Contoh: "Polis only (>1 fac)", "Insured only (Beda Currency)", "Polis only (Beda Periode)"
    special_flags_map = {
        "Matching >1 fac code": ">1 fac",
        "Matching >1 Fac code OC MC": ">1 fac",
        "Beda Currency": "Beda Currency",
        "Beda Periode": "Beda Periode",
    }
    mask_special = df["FLAG_PROD"].isin(special_flags_map.keys()) & ~df["SKENARIO"].isin(["UNMATCHED", "Unmatching", ""])
    df.loc[mask_special, "SKENARIO"] = (
        df.loc[mask_special, "SKENARIO"] + " (" + df.loc[mask_special, "FLAG_PROD"].map(special_flags_map) + ")"
    )

    for col in FINAL_COLUMNS:
        if col not in df.columns:
            df[col] = ""
    df = df[FINAL_COLUMNS]

    # [6] Export
    print(f"\n[6/6] Saving to: {OUTPUT_FILE} ...", flush=True)
    for c in df.select_dtypes(include=['object']).columns:
        df[c] = df[c].astype(str).str.replace(r'[\x00-\x08\x0B\x0C\x0E-\x1F]', '', regex=True)
    
    save_path = OUTPUT_FILE
    try:
        df.to_excel(save_path, index=False)
    except PermissionError:
        save_path = os.path.join(_BASE_DIR, "matching_tripa_new.xlsx")
        print(f"  [WARN] {OUTPUT_FILE} sedang dibuka di Excel (Permission denied). Menyimpan ke: {save_path}", flush=True)
        df.to_excel(save_path, index=False)

    try:
        from openpyxl import load_workbook
        workbook = load_workbook(save_path)
        worksheet = workbook.active
        text_cols = [
            "POLIS_ORI",
            "POLICY_CLEAN_1", "SERTIF_CLEAN_1",
            "POLICY_CLEAN_2", "SERTIF_CLEAN_2",
            "POLICY_CLEAN_3", "SERTIF_CLEAN_3",
            "POLICY_CLEAN_4", "POLICY_CLEAN_5",
            "SLIP_NO_ORI", "SLIP_NO_CLN", 
            "RECEIPT NO", "CREDIT NOTES", "DETAIL RINCIAN NO"
        ]
        text_col_indices = [idx for idx, col in enumerate(df.columns, start=1) if col in text_cols]
        for col_idx in text_col_indices:
            for row in range(2, worksheet.max_row + 1):
                cell = worksheet.cell(row=row, column=col_idx)
                cell.number_format = "@"
                if cell.value is not None:
                    cell.value = str(cell.value)
        workbook.save(save_path)
    except Exception as e:
        print(f"  [WARN] Format text gagal: {e}")

    try:
        from excel_styler import apply_purple_column_style
        apply_purple_column_style(save_path, "matching")
    except Exception as e:
        print(f"  [WARN] Styling gagal: {e}")

    # Export ke PostgreSQL
    load_dotenv()
    print("  Exporting ke PostgreSQL ...")
    conn_str = "postgresql+psycopg2://{}:{}@{}:{}/{}".format(
        os.environ.get("DB_USER", "postgres"), os.environ.get("DB_PASS", "postgres"),
        os.environ.get("DB_HOST", "localhost"), os.environ.get("DB_PORT", "5432"),
        os.environ.get("DB_NAME", "postgres"),
    )
    try:
        engine = create_engine(conn_str)
        df.to_sql("SUSPENSE_DATA_SUSPENSE_TRIPA_V1", con=engine, if_exists="append", index=False)
        print(f"      -> {len(df):,} rows berhasil di-export ke 'SUSPENSE_DATA_SUSPENSE_TRIPA_V1'")
    except Exception as e:
        print(f"  [WARN] PostgreSQL export gagal: {e}")

    elapsed = time.perf_counter() - t_start
    print(f"\n{'=' * 60}")
    print(f"  Selesai dalam {elapsed:.1f}s ({elapsed/60:.1f} menit)")
    print(f"  Rows  : {len(df):,}  |  Cols: {len(df.columns)}  |  File: {OUTPUT_FILE}")
    print(f"\n  FLAG_PROD:")
    for flag, count in df["FLAG_PROD"].value_counts().items():
        print(f"    {flag:<45}: {count:,}")
    _bad_sce = set(df["SKENARIO"].unique()) - set(_ALLOWED_SCENARIOS)
    if _bad_sce:
        print(f"\n  [WARN] SKENARIO di luar daftar yang diizinkan: {sorted(_bad_sce)}")
    print(f"\n  SKENARIO:")
    for sce, count in df["SKENARIO"].value_counts().items():
        print(f"    {sce:<50}: {count:,}")
    print("=" * 60)


if __name__ == "__main__":
    run()