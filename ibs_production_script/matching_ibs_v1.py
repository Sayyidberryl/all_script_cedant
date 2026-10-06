import os
import re
import sys
import time
import functools
import collections

sys.modules['numexpr'] = None
sys.modules['bottleneck'] = None

import numpy as np
import pandas as pd
import difflib
from dotenv import load_dotenv
from sqlalchemy import create_engine


# =============================================================================
# KONFIGURASI (CEDANT: IBS)
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


# Path file data
SUSPEND_FILE = _find_file_in(_CLEAN_DIR, ["ibs_output_suspense_V2.xlsx", "ibs_output_suspense.xlsx"])
OSBAL_FILE   = _find_file_in(_CLEAN_DIR, ["ibs_output_osbal_V2.xlsx", "osbal_clean_ibs.xlsx"])
FACUL_FILE   = _find_file_in(_CLEAN_DIR, ["ibs_output_facul_V2.xlsx"])
SLIPDB_FILE  = _find_file_in(_RAW_DIR, [
    "5. Database Rekap RI Slip.xlsx",
    "5.Database Rekap RI Slip.xlsx",
    "Database Rekap RI Slip.xlsx",
])
BORDERO_FILE = _find_file_in(_RAW_DIR, [
    "September 2024 - Agustus 2026 (COB PAR & Earthquake).xlsx",
])
OUTPUT_FILE  = os.path.join(_BASE_DIR, "final_output_ibs_v1.xlsx")

# Kolom di tabel Bordero IBS (COB PAR & Earthquake)
BORDERO_FAC_COL     = "FAC CODE"
BORDERO_POLIS_COL   = "POLIS"
BORDERO_SLIP_COL    = "SLIP"
BORDERO_CERT_COL    = "CERTIFICATE"
BORDERO_INSURED_COL = "INSURED"
BORDERO_CURR_COL    = "CURR"
BORDERO_BULAN_COL   = "BULAN"
BORDERO_TAHUN_COL   = "TAHUN"
BORDERO_NET_COL     = "NET"

# Toleransi matching AMOUNT ORI vs NET bordero
NET_TOLERANCE_PCT = 0.00
NET_MATCH_MODE    = "per_row"

# Mapping nama bulan Indonesia -> angka bulan (1-12)
_BULAN_MAP = {
    "JANUARI": 1, "FEBRUARI": 2, "MARET": 3, "APRIL": 4,
    "MEI": 5, "JUNI": 6, "JULI": 7, "AGUSTUS": 8,
    "SEPTEMBER": 9, "SEPTEMBER ": 9, "OKTOBER": 10, "NOVEMBER": 11, "DESEMBER": 12,
}

# Kolom FAC CODE di masing-masing tabel referensi
OSBAL_FACODE_COL  = "CCOS_REF_CODE"
FACUL_FACODE_COL  = "FAC_CODE"
SLIPDB_FACODE_COL = "FAC_CODE"   # dapat di-override runtime sesuai header aktual

# Sumber nilai SLIP pada FASE 1
SUSPEND_SLIP_SOURCE_COL = "DESC_3_CLN"

# Kolom polis/slip suspend dipertahankan agar _get_effective_sus_cols() kompatibel
CLSDT_POLIS_COL = "CLSDT_POLICY_NO"
CLSDT_SLIP_COL  = "CLSDT_SLIP_NO"
FAC_POLIS_COL   = "FAC_POLICY_NO"
FAC_SLIP_COL    = "FAC_SLIP"

# Nama kolom currency/periode
SUSPEND_CURR_COL = "CURR ORI"    # kolom currency di tabel Suspend IBS
OSBAL_CURR_COL   = "CCOS_CURR"  # kolom currency di tabel OSBAL IBS
SUSPEND_DATE_COL = "RECEIPT DATE"
OSBAL_DATE_COL   = "FAC_COM_DATE"

# Marker baris administratif yang dikecualikan dari matching
_EXCLUDED_REF_MARKERS = ["HUTANG PIUTANG", "DATA SUSPENSE", "PENYELESAIAN SUSPEND", "PENYELESAIAN HUTANG"]

# Marker LINESLIP - pakai RECEIPT DATE - 1 bulan sebagai tanggal efektif
_LINESLIP_MARKERS = ("LINESLIP", "LINE SLIP")

_MIN_TOKEN_LEN = 3
_TOKEN_RE      = re.compile(r"[^A-Z0-9]+")
_DELIMITER_RE  = re.compile(r"[,;|+\s]+")

# Token generik yang dikecualikan saat menghitung overlap INSURED
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
})

# Overlap coefficient = |A∩B| / min(|A|, |B|) setelah exclude stopword.
# 0.75 = mencegah kecocokan kebetulan 2 kata umum (misal SAWIT LESTARI beda perusahaan)
_INSURED_LIKE_THRESHOLD: float = 0.75

# Threshold ketat untuk POLIS/SLIP (identifier dokumen — harus lebih presisi).
_POLIS_LIKE_THRESHOLD: float  = 0.80

# Kolom join saat akumulasi suspend
_SUSPEND_JOIN_COLS = [
    "RECEIPT NO", "CREDIT NOTES", "DETAIL RINCIAN NO", "RECEIPT DATE",
    "CEDANT NAME", "CEDANT SHRT NAME",
    "INSURED", "INSURED_CLN_1", "INSURED_CLN_2", "INSURED_CLN_3",
    "CURR ORI", "CURR PAY", "AMOUNT PAY",
    "POLIS", "POLIS_CLN_1",
    "SLIP NO", "SLIP_NO_CLN_1",
    "DESC 1", "DESC 2", "DESC 3", "DESC 4",
    "STATUS", "REC_TYPE",
]
_SUSPEND_SUM_COLS = ["AMOUNT ORI"]

# Kolom output standar dan kolom Mark
FINAL_COLUMNS = [
    "CCOS_DOC_NO", "CCOS_REF_CODE",
    "RECEIPT NO", "CREDIT NOTES", "DETAIL RINCIAN NO", "RECEIPT DATE",
    "CEDANT NAME", "CEDANT SHRT NAME",
    "INSURED_ORI", "INSURED_1", "INSURED_2",
    "CURR ORI", "AMOUNT ORI", "AMOUNT_ORI_MIN1",
    "CURR PAY", "AMOUNT PAY",
    "CCOS_OR_BAL", "CCOS_BAL_DUE", "DIFERENCE",
    "FLAG_PROD",
    "POLIS_ORI", "POLICY_CLEAN_1", "SERTIF_CLEAN_1",
    "POLICY_CLEAN_2", "SERTIF_CLEAN_2",
    "POLICY_CLEAN_3", "SERTIF_CLEAN_3",
    "POLICY_CLEAN_4", "POLICY_CLEAN_5",
    "SLIP_NO_ORI", "SLIP_NO_CLN",
    "DESC 1", "DESC 2", "DESC 3", "DESC_3_CLN", "DESC 4",
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
    return _norm_cached(str(value))


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


def breakdown_osbal_slip(raw: str) -> list:
    """Memecah berbagai format nomor slip OSBAL (range, multiple, noise END, dsb)."""
    if not raw or not isinstance(raw, str):
        return []
    s = str(raw).strip()
    if not s or s.lower() in ('nan', 'none', '-'):
        return []

    # 1. Bersihkan noise END markers & trailing keterangan bulan
    s_clean = re.sub(r'\+\s*END(?:\s*\d+)?\b', ' ', s, flags=re.IGNORECASE)
    s_clean = re.sub(r'/\s*END(?:\s*\d+)?\b', ' ', s_clean, flags=re.IGNORECASE)
    s_clean = re.sub(r'\bEND(?:\s*\d+)?\b', ' ', s_clean, flags=re.IGNORECASE)
    s_clean = re.sub(
        r'/\s*(?:JANUARI|FEBRUARI|MARET|APRIL|MEI|JUNI|JULI|AGUSTUS|SEPTEMBER|OKTOBER|NOVEMBER|DESEMBER)\s+\d{4}',
        '', s_clean, flags=re.IGNORECASE
    )
    s_clean = re.sub(r'[\s+,]+$', '', s_clean).strip()

    # Jika pola bordero / penyelesaian / non-standar khusus
    if re.match(r'^(?:BORDERO|PENYELESAIAN|TBA)', s_clean, flags=re.IGNORECASE):
        cleaned_str = re.sub(r'[^A-Za-z0-9]', '', s_clean).upper()
        return [cleaned_str] if cleaned_str else []

    # 2. Ekstrak prefix: tangani teks awalan + tahun (termasuk anomali 62022 -> 2022, & -> '')
    s_norm = s_clean.replace('&', '')
    m_pfx = re.match(r'^([A-Za-z\s.]*?)(?:6(20\d{2})|(20\d{2}))\s*(.*)$', s_norm)
    if m_pfx:
        alpha = re.sub(r'[^A-Za-z]', '', m_pfx.group(1).upper())
        yr = m_pfx.group(2) if m_pfx.group(2) else m_pfx.group(3)
        prefix_str = alpha + yr
        rest_str = m_pfx.group(4)
    else:
        prefix_str = ''
        rest_str = s_norm

    rest_str = re.sub(r'\s*-\s*', '-', rest_str)
    plus_parts = [p.strip() for p in rest_str.split('+') if p.strip()]

    result_codes = []
    for part in plus_parts:
        tokens = re.split(r'[,;\s]+', part)
        for tok in tokens:
            tok = tok.strip()
            if not tok:
                continue
            if '-' in tok:
                for dp in tok.split('-'):
                    dp_clean = re.sub(r'\D', '', dp)
                    if dp_clean:
                        result_codes.append(dp_clean)
            elif '/' in tok and prefix_str:
                for sp in tok.split('/'):
                    if '-' in sp:
                        for sdp in sp.split('-'):
                            sdp_clean = re.sub(r'\D', '', sdp)
                            if sdp_clean:
                                result_codes.append(sdp_clean)
                    else:
                        sp_clean = re.sub(r'\D', '', sp)
                        if sp_clean:
                            result_codes.append(sp_clean)
            elif '&' in tok:
                for ap in tok.split('&'):
                    ap_clean = re.sub(r'\D', '', ap)
                    if ap_clean:
                        result_codes.append(ap_clean)
            else:
                tok_clean = re.sub(r'\D', '', tok)
                if tok_clean:
                    result_codes.append(tok_clean)

    final_slips = []
    for c in result_codes:
        full_slip = f'{prefix_str}{c}'
        # Jangan anggap angka pendek/termin (misal '1', '2', '3' dari 'P1 + P2') sebagai nomor slip
        if len(full_slip) >= 5 and full_slip not in final_slips:
            final_slips.append(full_slip)

    if not final_slips and s_clean:
        fallback = re.sub(r'[^A-Za-z0-9]', '', s_clean).upper()
        return [fallback] if len(fallback) >= 5 else []

    return final_slips


# =============================================================================
# DICTIONARY COB IBS: Prefix slip suspend -> CLASS_NAME OSBAL
# =============================================================================
COB_DICTIONARY = {
    # Properti & Gempa Bumi
    "PAR":         {"PAR", "IAR", "PROPERTY ALL RISK", "EARTHQUAKE"},
    "PARIBSRE":    {"PAR", "IAR", "PROPERTY ALL RISK", "EARTHQUAKE"},
    "PARIBSE":     {"PAR", "IAR", "PROPERTY ALL RISK", "EARTHQUAKE"},
    "PARISRE":     {"PAR", "IAR", "PROPERTY ALL RISK", "EARTHQUAKE"},
    "PART":        {"PAR", "IAR", "PROPERTY ALL RISK", "EARTHQUAKE"},
    "PARTIBSRE":   {"PAR", "IAR", "PROPERTY ALL RISK", "EARTHQUAKE"},
    "PPAR":        {"PAR", "IAR", "PROPERTY ALL RISK", "EARTHQUAKE"},
    "PPARIBSRE":   {"PAR", "IAR", "PROPERTY ALL RISK", "EARTHQUAKE"},
    "CROPS":       {"PAR", "IAR", "PROPERTY ALL RISK", "EARTHQUAKE"},
    "CROPSIBSRE":  {"PAR", "IAR", "PROPERTY ALL RISK", "EARTHQUAKE"},
    # Combined PAR+MB — boleh cocok ke PAR, MB, maupun FIRE (karena OSBAL sering pakai FIRE utk paket)
    "PARMB":       {"PAR", "MB", "MACHINERY BREAKDOWN", "PROPERTY ALL RISK", "EARTHQUAKE", "IAR", "FIRE"},
    "PARMBIBSRE":  {"PAR", "MB", "MACHINERY BREAKDOWN", "PROPERTY ALL RISK", "EARTHQUAKE", "IAR", "FIRE"},
    # Fire
    "FIRE":        {"FIRE"},
    "FIREIBSRE":   {"FIRE"},
    # Heavy Equipment
    "CPM":         {"HEAVY EQUIPMENT", "CPM/CPE"},
    "CPMIBSRE":    {"HEAVY EQUIPMENT", "CPM/CPE"},
    # Marine Cargo
    "OC":          {"MARINE CARGO"},
    "MOC":         {"MARINE CARGO"},
    "MOCIBSRE":    {"MARINE CARGO"},
    # CECR
    "CECR":        {"CECCR", "CECR"},
    "CECRIBSRE":   {"CECCR", "CECR"},
    # Machinery Breakdown
    "MB":          {"MB", "MACHINERY BREAKDOWN"},
    "MBIBSRE":     {"MB", "MACHINERY BREAKDOWN"},
    "CMI":         {"MB", "MACHINERY BREAKDOWN"},
    "CMIIBSRE":    {"MB", "MACHINERY BREAKDOWN"},
    # EEI
    "EEI":         {"EEI"},
    "EEIIBSRE":    {"EEI"},
    # Personal Accident
    "PA":          {"PA"},
    "PAIBSRE":     {"PA"},
    # Public Liability
    "PL":          {"PUBLIC LIAB", "PUBLIC L", "PL"},
    "CS":          {"PUBLIC LIAB", "PUBLIC L", "PL"},
    "CSIBURE":     {"PUBLIC LIAB", "PUBLIC L", "PL"},
    # Aviation
    "AVIATION":    {"AVIATION"},
    "CGL":         {"AVIATION"},
    "CGLIBSRE":    {"AVIATION"},
    # Directors & Officers
    "DNO":         {"DIRECTOR OFF", "DNO"},
    "D&O":         {"DIRECTOR OFF", "DNO"},
    "D&OIBSRE":    {"DIRECTOR OFF", "DNO"},
    "DOIBSRE":     {"DIRECTOR OFF", "DNO"},
    # Comprehensive Liability
    "CLI":         {"COMPREHENSIVE L", "CLI"},
    "TOL":         {"COMPREHENSIVE L", "CLI"},
    "TOLIBSRE":    {"COMPREHENSIVE L", "CLI"},
    # Professional Indemnity
    "PI":          {"PROFESSIONAL INDEMNITY", "PI"},
    "PIIBSRE":     {"PROFESSIONAL INDEMNITY", "PI"},
    # CAR / EAR
    "CAR":         {"CAR", "EAR"},
    "CARIBSRE":    {"CAR", "EAR"},
    "EAR":         {"CAR", "EAR"},
    "EARIBSRE":    {"CAR", "EAR"},
    "BUILDERIBSRE":{"CAR", "EAR"},
    # GTI (sama keluarga PAR)
    "GTI":         {"PAR", "IAR", "PROPERTY ALL RISK", "EARTHQUAKE"},
    "GTIIBSRE":    {"PAR", "IAR", "PROPERTY ALL RISK", "EARTHQUAKE"},
    # PSR
    "PSR":         {"PAR", "IAR", "PROPERTY ALL RISK", "EARTHQUAKE"},
    "PSRIBSRE":    {"PAR", "IAR", "PROPERTY ALL RISK", "EARTHQUAKE"},
}


def _extract_cob_prefix(slip_str: str) -> str:
    """Ekstrak kode prefix COB terpanjang yang cocok dari slip string."""
    if not slip_str:
        return ""
    clean = re.sub(r'[^A-Z0-9&/]', '', str(slip_str).upper())
    for pfx in sorted(COB_DICTIONARY.keys(), key=lambda x: -len(x)):
        if clean.startswith(pfx):
            return pfx
    return ""


def _extract_slip_numbers(slip_str: str) -> list:
    """Ekstrak angka 8-digit (YYYYNNNN) dari slip suspend."""
    if not slip_str:
        return []
    clean = str(slip_str).strip().upper().replace(" ", "")
    # Hanya ambil angka tepat 8 digit (YYYYNNNN)
    return re.findall(r'\d{8}', clean)


def _osbal_slip_text(osb_row: dict) -> str:
    """Gabungkan semua kolom slip OSBAL menjadi satu string bersih (tanpa spasi)."""
    parts = []
    cols = ["FAC_SLIP", "CLSDT_SLIP_NO", "SLIP_NO_CLN_1", "SLIP_NO_CLN_2",
            "SLIP_NO_CLN_3", "SLIP_NO_CLN_4"]
    for k in osb_row:
        if k.startswith("slip_clean_") and k not in cols:
            cols.append(k)
    for c in cols:
        v = osb_row.get(c)
        if v and str(v).strip() not in ("", "nan", "None"):
            parts.append(str(v))
    return "".join(parts).replace(" ", "").upper()


# Mapping Class Name OSBAL -> Prefix Slip sesuai Dictionary & Karakteristik IBS
# Mendukung cross-matching dalam 1 COB (misal MB <-> CMI, PAR <-> CROPS, D&O <-> DNO)
COB_SLIP_PREFIXES = {
    # Properti / Kebakaran saling beririsan (PAR, CROPS, FIRE, PPAR typo, PART typo)
    'PAR': ['PARIBSRE', 'CROPSIBSRE', 'FIREIBSRE', 'PPARIBSRE', 'PARTIBSRE'],
    'IAR': ['PARIBSRE', 'CROPSIBSRE', 'FIREIBSRE'],
    'PROPERTY ALL RISK': ['PARIBSRE', 'CROPSIBSRE', 'FIREIBSRE'],
    'EARTHQUAKE': ['PARIBSRE', 'CROPSIBSRE', 'FIREIBSRE'],
    'FIRE': ['FIREIBSRE', 'PARIBSRE', 'CROPSIBSRE', 'PPARIBSRE'],
    'HEAVY EQUIPMENT': ['CPMIBSRE'],
    'CPM/CPE': ['CPMIBSRE'],
    'CECCR': ['CECRIBSRE'],
    'CECR': ['CECRIBSRE'],
    'MB': ['MBIBSRE', 'CMIIBSRE'],
    'MACHINERY BREAKDOWN': ['MBIBSRE', 'CMIIBSRE'],
    'EEI': ['EEIIBSRE'],
    'PA': ['PAIBSRE'],
    'PUBLIC LIAB': ['CSIBURE'],
    'PUBLIC L': ['CSIBURE'],
    'PL': ['CSIBURE'],
    'AVIATION': ['CGLIBSRE'],
    'DIRECTOR OFF': ['D&OIBSRE', 'DOIBSRE', 'OIBSRE', 'DNOIBSRE'],
    'DNO': ['D&OIBSRE', 'DOIBSRE', 'OIBSRE', 'DNOIBSRE'],
    'COMPREHENSIVE L': ['TOLIBSRE'],
    'CLI': ['TOLIBSRE'],
    'PROFESSIONAL INDEMNITY': ['PIIBSRE'],
    'CAR': ['CARIBSRE', 'EARIBSRE'],
    'EAR': ['CARIBSRE', 'EARIBSRE'],
    'MARINE CARGO': ['OC'],
}


def expand_slip_cob(slips: list, class_name: str) -> list:
    """Melengkapi prefix slip OSBAL berdasarkan class_name / COB di backend."""
    if not class_name:
        return slips
    cn = str(class_name).strip().upper()
    prefixes = COB_SLIP_PREFIXES.get(cn, [])
    if not prefixes:
        return slips

    expanded = list(slips)
    for s in slips:
        s_upper = s.upper()
        # 1. Slip berupa angka (tahun + urutan, misal 20241531)
        if s.isdigit():
            for pfx in prefixes:
                cand = f'{pfx}{s}'
                if cand not in expanded:
                    expanded.append(cand)
        # 2. Slip diawali IBSRE (misal IBSRE20220298)
        elif s_upper.startswith('IBSRE'):
            for pfx in prefixes:
                base = pfx.replace('IBSRE', '')
                if base:
                    cand = f'{base}{s_upper}'
                    if cand not in expanded:
                        expanded.append(cand)
        # 3. Varian prefix dalam 1 COB (misal MB vs CMI, PAR vs CROPS, DOIBSRE vs OIBSRE)
        else:
            for current_pfx in prefixes:
                if s_upper.startswith(current_pfx):
                    suffix = s_upper[len(current_pfx):]
                    for alt_pfx in prefixes:
                        if alt_pfx != current_pfx:
                            cand = f'{alt_pfx}{suffix}'
                            if cand not in expanded:
                                expanded.append(cand)

    return expanded


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

def _is_excluded_ref_row(row: dict, polis_cols: list, slip_cols: list, insured_cols: list = None) -> bool:
    # True jika baris referensi adalah baris administratif
    insured_cols = insured_cols or []
    for col in list(polis_cols) + ["polis_ori"] + list(slip_cols) + ["slip_ori"] + list(insured_cols) + ["insured_ori", "INSURED"]:
        val = _normalize(row.get(col, ""))
        if val and any(m in val for m in _EXCLUDED_REF_MARKERS):
            return True
    return False


def _is_lineslip_row(sus_row: dict, insured_cols: list) -> bool:
    # True jika baris suspend mengandung marker LINESLIP di kolom insured
    # Kolom ori INSURED
    for col in list(insured_cols) + ["INSURED"]:
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
    """Tentukan kolom efektif polis/slip per baris."""
    clsdt_polis = _normalize(sus_row.get(CLSDT_POLIS_COL, ""))
    fac_polis   = _normalize(sus_row.get(FAC_POLIS_COL,   ""))
    if clsdt_polis:
        eff_polis  = [CLSDT_POLIS_COL]
        like_polis = [CLSDT_POLIS_COL]
    else:
        eff_polis  = base_polis_cols
        like_polis = list(base_polis_cols) + ([FAC_POLIS_COL] if fac_polis else [])

    # eff_slip tidak dipakai sebagai fase tersendiri
    eff_slip  = []
    like_slip = []

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
    """Bangun semua struktur lookup untuk satu tabel referensi ."""
    if not rows:
        if label:
            print(f"  {label:<6} done (0.0s)", flush=True)
        return ({}, {}, {}, {}, {}, {}, ({}, {}), ({}, {}), ({}, {}), {})

    t0 = time.perf_counter()
    excluded = {i for i, r in enumerate(rows) if _is_excluded_ref_row(r, polis_cols, slip_cols, insured_cols)}

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
    match_log: list = None,
) -> list:
    """Kembalikan indeks baris yang "mirip" dengan query, berdasarkan tiga kriteria berurutan."""
    valid_queries = [v for v in query_values if v and len(v) >= 2]
    if not valid_queries:
        return []

    # --- Bangun query token pool ---
    query_tokens: frozenset = frozenset()
    for v in valid_queries:
        query_tokens = query_tokens | _tokenize(v)

    # Token signifikan = buang stopword (untuk overlap calculation)
    sig_query_tokens = query_tokens - stopwords

    # --- Stage 1: candidate pruning via token index (unchanged) ---
    candidates: set = set()
    for tok in query_tokens:
        if tok in token_index:
            candidates |= token_index[tok]
    if not candidates:
        return []

    # --- Stage 2: decision per candidate ---
    matched_set: dict = {}   # row_idx -> method, deduplicated, first-method-wins
    for i in candidates:
        # Kumpulkan nilai referensi untuk SUBSTR check
        ref_values = [_normalize(rows[i].get(c, "")) for c in ref_cols if rows[i].get(c)]

        # ── Kriteria 1: SUBSTR (backward compat) ──────────────────────────
        substr_ok = any(q in rv for q in valid_queries for rv in ref_values if rv)
        if substr_ok:
            if i not in matched_set:
                matched_set[i] = "SUBSTR"
            continue  # sudah match, skip kriteria lain

        # Evaluasi per-entitas untuk mencegah false positive multi-entitas
        entity_matched = False
        for q in valid_queries:
            q_toks = _tokenize(q)
            sig_q = q_toks - stopwords
            if not sig_q:
                continue

            for rv in ref_values:
                if not rv:
                    continue
                rv_toks = _tokenize(rv)
                sig_rv = rv_toks - stopwords
                if not sig_rv:
                    continue

                # Kriteria 2: SUBSET per-entitas
                if sig_q.issubset(rv_toks):
                    if i not in matched_set:
                        matched_set[i] = "SUBSET"
                    entity_matched = True
                    break

                # Kriteria 3: OVERLAP coefficient per-entitas (dengan toleransi typo ringan pada token, misal MEBELLINDO vs MEBELINDO)
                matched_rv = set()
                inter = 0
                for t1 in sig_q:
                    for t2 in sig_rv:
                        if t2 in matched_rv:
                            continue
                        if t1 == t2:
                            inter += 1
                            matched_rv.add(t2)
                            break
                        elif len(t1) >= 4 and len(t2) >= 4 and difflib.SequenceMatcher(None, t1, t2).ratio() >= 0.85:
                            inter += 1
                            matched_rv.add(t2)
                            break

                denom = min(len(sig_q), len(sig_rv))
                if denom > 0 and (inter / denom) >= overlap_threshold:
                    if i not in matched_set:
                        matched_set[i] = "OVERLAP"
                    entity_matched = True
                    break

            if entity_matched:
                break

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
                return hit, "Insured only"

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
    """Rematch ketat dengan filter currency + periode (R1-R5 standar ACA)."""
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
# BORDERO MATCHING & NARROWING (COB PAR & EARTHQUAKE)
# =============================================================================

def _clean_slip_key(s: str) -> str:
    """Bersihkan nomor slip untuk matching exact alfanumerik."""
    return re.sub(r"[^A-Z0-9]", "", str(s).upper()) if s else ""


def _build_bordero_index(bordero_rows: list) -> tuple:
    """Bangun index bordero PAR & Earthquake."""
    bordero_idx: dict = {}
    bordero_slip_index: dict = {}

    for row in bordero_rows:
        fac = _normalize(row.get(BORDERO_FAC_COL, ""))
        if not fac:
            continue
        bulan_str = str(row.get(BORDERO_BULAN_COL, "")).strip().upper()
        tahun_str = str(row.get(BORDERO_TAHUN_COL, "")).strip()
        month = _BULAN_MAP.get(bulan_str, 0)
        ym = (int(tahun_str), month) if month and tahun_str.isdigit() else None
        
        raw_slip = str(row.get(BORDERO_SLIP_COL, "")).strip()
        clean_slip = _clean_slip_key(raw_slip)

        entry = {
            "polis":      _normalize(row.get(BORDERO_POLIS_COL, "")),
            "slip_raw":   raw_slip,
            "slip_clean": clean_slip,
            "cert":       _normalize(row.get(BORDERO_CERT_COL, "")),
            "insured":    _normalize(row.get(BORDERO_INSURED_COL, "")),
            "curr":       _normalize(row.get(BORDERO_CURR_COL, "")),
            "period":     ym,
            "net":        row.get(BORDERO_NET_COL),
        }
        bordero_idx.setdefault(fac, []).append(entry)
        if clean_slip:
            bordero_slip_index.setdefault(clean_slip, []).append((fac, entry))

    return bordero_idx, bordero_slip_index


def _find_fac_codes_in_bordero(
    suspend_row: dict,
    bordero_idx: dict,
    bordero_slip_index: dict,
    slip_sus_cols: list,
) -> tuple:
    """Cari FAC CODE di Bordero berdasarkan SLIP (prioritas DESC_3_CLN & SLIP NO),."""
    if not bordero_slip_index:
        return set(), False

    # Kumpulkan kandidat slip dari baris suspend
    sus_slips = []
    # DESC_3_CLN
    desc3 = _normalize(suspend_row.get(SUSPEND_SLIP_SOURCE_COL, ""))
    if desc3:
        sus_slips.append(_clean_slip_key(desc3))
    # SLIP NO / SLIP_NO_CLN
    for c in slip_sus_cols:
        v = _normalize(suspend_row.get(c, ""))
        if v:
            sus_slips.append(_clean_slip_key(v))

    sus_amount = suspend_row.get("AMOUNT ORI")
    try:
        sus_amount_f = float(sus_amount) if sus_amount is not None else None
    except (ValueError, TypeError):
        sus_amount_f = None

    sus_curr_val = _normalize(suspend_row.get(SUSPEND_CURR_COL, ""))

    candidates: dict = {}  # fac -> list of entries
    for s_clean in sus_slips:
        if not s_clean:
            continue
        if s_clean in bordero_slip_index:
            for fac, entry in bordero_slip_index[s_clean]:
                candidates.setdefault(fac, []).append(entry)

    if not candidates:
        return set(), False

    # Filter 1: Currency matching
    if sus_curr_val:
        curr_filtered = {}
        for fac, entries in candidates.items():
            valid = [e for e in entries if not e.get("curr") or e.get("curr") == sus_curr_val]
            if valid:
                curr_filtered[fac] = valid
        candidates = curr_filtered

    # Filter 2: Net Amount matching jika masih > 1 kandidat FAC CODE
    if len(candidates) > 1 and sus_amount_f is not None:
        amt_abs = abs(sus_amount_f)
        net_filtered = {}
        for fac, entries in candidates.items():
            valid = []
            for e in entries:
                try:
                    b_net = abs(float(e.get("net", 0) or 0))
                    if abs(b_net - amt_abs) <= (amt_abs * NET_TOLERANCE_PCT):
                        valid.append(e)
                except (ValueError, TypeError):
                    continue
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
) -> dict:
    """Bangun satu baris output."""
    has_match = bool(source and osbal_rows and osbal_rows[0])
    has_osbal = has_match and (source in ("OSBAL", "SLIPDB", "FACUL") or resolved)

    ccos = (
        _aggregate_ccos(osbal_rows, all_fac_osbal_rows=all_fac_osbal_rows)
        if has_osbal
        else {"CCOS_DOC_NO": "", "CCOS_REF_CODE": "", "CCOS_OR_BAL": np.nan, "CCOS_BAL_DUE": np.nan}
    )
    if scenario in ("Beda Currency", "Beda Periode"):
        ccos["CCOS_OR_BAL"] = np.nan
        ccos["CCOS_BAL_DUE"] = np.nan

    fac_label = (
        _facode_label(osbal_rows, "OSBAL") if has_osbal else
        _facode_label(osbal_rows, source)  if has_match else ""
    )

    final_scenario = "Unmatching" if not source or scenario in ("Unmatching", "UNMATCHED") else scenario
    if "LIKE" in str(final_scenario).upper():
        final_scenario = "Insured only"

    return {
        "CCOS_DOC_NO":       ccos["CCOS_DOC_NO"] if has_osbal else "",
        "CCOS_REF_CODE":     ccos["CCOS_REF_CODE"],
        "RECEIPT NO":        suspend_row.get("RECEIPT NO",        ""),
        "CREDIT NOTES":      suspend_row.get("CREDIT NOTES",      ""),
        "DETAIL RINCIAN NO": suspend_row.get("DETAIL RINCIAN NO", ""),
        "RECEIPT DATE":      suspend_row.get("RECEIPT DATE",      ""),
        "CEDANT NAME":       suspend_row.get("CEDANT NAME",       ""),
        "CEDANT SHRT NAME":  suspend_row.get("CEDANT SHRT NAME",  ""),

        # Kolom insured
        "INSURED_ORI": suspend_row.get("INSURED",       ""),
        "INSURED_1":   suspend_row.get("INSURED_CLN_1", ""),
        "INSURED_2":   suspend_row.get("INSURED_CLN_2", ""),

        "CURR ORI":   suspend_row.get("CURR ORI",   ""),
        "AMOUNT ORI": suspend_row.get("AMOUNT ORI", ""),
        "CURR PAY":   suspend_row.get("CURR PAY",   ""),
        "AMOUNT PAY": suspend_row.get("AMOUNT PAY", ""),

        "CCOS_OR_BAL":  ccos["CCOS_OR_BAL"],
        "CCOS_BAL_DUE": ccos["CCOS_BAL_DUE"],

        # POLIS_CLN & SERTIF dari suspend
        "POLIS_ORI":       suspend_row.get("POLIS", ""),
        "POLICY_CLEAN_1":  str(suspend_row["POLIS_CLN_1"]).strip() if pd.notna(suspend_row.get("POLIS_CLN_1")) and str(suspend_row.get("POLIS_CLN_1")).strip() not in ("", "nan", "None") else (str(suspend_row["POLIS"]).strip() if pd.notna(suspend_row.get("POLIS")) and str(suspend_row.get("POLIS")).strip() not in ("", "nan", "None") else ""),
        "SERTIF_CLEAN_1":  _format_sertif(suspend_row.get("CERTIFICATE", "") or suspend_row.get("clean sertif 1", "")),
        "POLICY_CLEAN_2":  "",
        "SERTIF_CLEAN_2":  "",
        "POLICY_CLEAN_3":  "",
        "SERTIF_CLEAN_3":  "",
        "POLICY_CLEAN_4":  "",
        "POLICY_CLEAN_5":  "",
        "SLIP_NO_ORI":     suspend_row.get("SLIP NO", ""),
        "SLIP_NO_CLN":     suspend_row.get("SLIP_NO_CLN_1", ""),

        "DESC 1": suspend_row.get("DESC 1", ""),
        "DESC 2": suspend_row.get("DESC 2", ""),
        "DESC 3": suspend_row.get("DESC 3", ""),
        "DESC_3_CLN": suspend_row.get("DESC_3_CLN", ""),
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
    """Hitung AMOUNT_ORI_MIN1, DIFERENCE, dan FLAG_PROD ."""
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
    is_equal    = amount_neg.round(2) == balance_due.round(2)

    if "_IS_BEDA_PERIODE" in df.columns:
        is_beda_periode = df["_IS_BEDA_PERIODE"].fillna(False).astype(bool)
    else:
        is_beda_periode = skenario_col.str.contains("Beda Periode", na=False, regex=False)

    # Deteksi beda currency dari kolom helper _IS_BEDA_CURR
    if "_IS_BEDA_CURR" in df.columns:
        is_beda_currency = df["_IS_BEDA_CURR"].fillna(False).astype(bool)
    else:
        # Fallback: coba dari SKENARIO (backward compat)
        is_beda_currency = skenario_col.str.contains("Beda Currency", na=False, regex=False)
    # Kondisi is_gt1_ocmc untuk OCMC
    is_gt1_ocmc          = skenario_col.str.contains("Bordero OC Override", na=False, regex=False) & ccos_ref.str.contains(",", na=False)
    is_gt1_fac           = ccos_ref.str.contains(",", na=False) & ~is_gt1_ocmc
    is_new_entry_total   = (amount_ori != 0) & (balance_due == 0)
    is_adj_total         = is_equal
    is_adj_sebagian      = ~is_equal & (amount_neg < balance_due)
    is_new_entry_sebagian= ~is_equal & (amount_neg > balance_due)

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

    for col in ["_OSBAL_ROW_COUNT", "_SUSPEND_ROW_COUNT", "_MERGED_AMOUNT_ORI", "_IS_BEDA_CURR", "_IS_BEDA_PERIODE"]:
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
    """Pipeline utama IBS V1: matching 5 fase ."""
    t_start = time.perf_counter()
    print("\n" + "=" * 60, flush=True)
    print("  PRODUCTION SCRIPT -- SUSPEND MATCHING  (IBS) -- V1", flush=True)
    print("=" * 60, flush=True)

    # [1] Load data
    # [IBS DELTA] Tidak ada bordero_rows -- bordero dihapus total untuk IBS
    print("\n[1/6] Loading data ...", flush=True)
    suspend_rows, suspend_cols = _load_excel(SUSPEND_FILE, "SUSPEND")
    osbal_rows,   osbal_cols   = _load_excel(OSBAL_FILE,   "OSBAL")
    facul_rows,   facul_cols   = _load_excel(FACUL_FILE,   "FACUL")
    slipdb_rows,  slipdb_cols  = _load_excel(SLIPDB_FILE,  "SLIPDB", optional=True)
    bordero_rows, bordero_cols = _load_excel(BORDERO_FILE, "BORDERO", optional=True)

    # Diagnostik kolom kunci
    print("\n  --- Cek kolom SUSPEND ---", flush=True)
    print(f"  Semua kolom SUSPEND: {suspend_cols}", flush=True)
    for col in [SUSPEND_SLIP_SOURCE_COL, "INSURED", "INSURED_CLN_1", "INSURED_CLN_2", "INSURED_CLN_3",
                "POLIS", "POLIS_CLN_1", "SLIP NO", "SLIP_NO_CLN_1",
                SUSPEND_CURR_COL, SUSPEND_DATE_COL]:
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

    # [IBS DELTA] FASE 1 memakai DESC_3_CLEAN sebagai nilai slip suspend
    if SUSPEND_SLIP_SOURCE_COL in suspend_cols:
        desc3_slip_vals = [SUSPEND_SLIP_SOURCE_COL]
        print(f"  [IBS DELTA] FASE 1: '{SUSPEND_SLIP_SOURCE_COL}' ditemukan -- dipakai sebagai nilai SLIP.", flush=True)
    else:
        desc3_slip_vals = []
        print(f"  [WARN] '{SUSPEND_SLIP_SOURCE_COL}' tidak ditemukan di SUSPEND -- FASE 1 IBS tidak berjalan!", flush=True)

    # Kolom polis/slip standar suspend (untuk FASE 2 via _get_effective_sus_cols)
    polis_sus = _get_clean_cols(suspend_cols, "POLIS_CLN") or _get_clean_cols(suspend_cols, "clean polis") or []
    slip_sus  = _get_clean_cols(suspend_cols, "SLIP_NO_CLN") or _get_clean_cols(suspend_cols, "clean slip")  or []

    # [IBS DELTA] Insured: "INSURED_CLN_N" (bukan "clean insured N")
    insured_sus = _get_clean_cols(suspend_cols, "INSURED_CLN") or _get_clean_cols(suspend_cols, "clean insured") or []

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
    elif "CCOS_REF_CODE" in slipdb_cols: SLIPDB_FACODE_COL = "CCOS_REF_CODE"

    def _prepend(col, cols_list, header):
        return ([col] + cols_list) if col in header and col not in cols_list else cols_list

    polis_osbal  = _prepend("CLSDT_POLICY_NO", polis_osbal,  osbal_cols)
    slip_osbal   = _prepend("CLSDT_SLIP_NO",   slip_osbal,   osbal_cols)

    # In-memory slip breakdown & COB prefix expansion untuk OSBAL (tidak dimunculkan di file output OSBAL)
    print("  Memecah (breakdown) & melengkapi prefix COB format slip OSBAL untuk in-memory index ...", flush=True)
    max_bd_count = 0
    bd_rows_count = 0
    for r in osbal_rows:
        raw_slip = str(r.get("FAC_SLIP", "") or r.get("CLSDT_SLIP_NO", "") or r.get("SLIP_NO_CLN_1", ""))
        class_name = str(r.get("CLASS_NAME", "")).strip()
        slips = breakdown_osbal_slip(raw_slip)
        # Tambahan rules COB (melengkapi prefix slip dan cross-match dalam 1 COB)
        slips = expand_slip_cob(slips, class_name)
        if len(slips) > 1:
            bd_rows_count += 1
        if len(slips) > max_bd_count:
            max_bd_count = len(slips)
        for idx, s_val in enumerate(slips, 1):
            r[f"slip_clean_{idx}"] = s_val

    bd_cols = [f"slip_clean_{i}" for i in range(1, max_bd_count + 1)]
    slip_osbal = bd_cols + [c for c in slip_osbal if c not in bd_cols]
    print(f"  -> {bd_rows_count:,} baris OSBAL memiliki multiple slip/COB expansion (maksimal {max_bd_count} slip/baris: {', '.join(bd_cols[:5])}...)", flush=True)

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

    # Bangun index Bordero (COB PAR & Earthquake)
    bordero_idx, bordero_slip_index = _build_bordero_index(bordero_rows) if bordero_rows else ({}, {})
    print(f"  Bordero index: {len(bordero_idx):,} FAC CODE entries, {len(bordero_slip_index):,} unique slips" if bordero_idx
          else "  Bordero: tidak tersedia, dilewati", flush=True)

    # Bangun index angka slip 8-digit OSBAL untuk pencarian cepat FASE 1 & narrowing
    t_num = time.perf_counter()
    osbal_slip_num_idx = collections.defaultdict(set)
    for idx, row in enumerate(osbal_rows):
        nums_in_row = set()
        for c in slip_osbal:
            v = row.get(c)
            if v and str(v).strip() not in ("", "nan", "None"):
                for num in _extract_slip_numbers(str(v)):
                    nums_in_row.add(num)
        for num in _extract_slip_numbers(_osbal_slip_text(row)):
            nums_in_row.add(num)
        for num in nums_in_row:
            osbal_slip_num_idx[num].add(idx)
    print(f"  OSBAL slip 8-digit index: {len(osbal_slip_num_idx):,} unique 8-digit numbers ({time.perf_counter()-t_num:.2f}s)", flush=True)

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

        # Ekstrak nama insured dari DESC 3 jika ada format slip IBSRE + "-"
        desc3_override_insured = None
        desc3_val = _normalize(sus.get("DESC 3", ""))
        if "IBSRE" in desc3_val and " - " in desc3_val:
            desc3_insured_ori = desc3_val.split(" - ", 1)[1].strip()
            if desc3_insured_ori:
                desc3_override_insured = _normalize(desc3_insured_ori)

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
            sertif_sus_cols=[],
        )

        # FASE 1: Pencarian awal berdasarkan SLIP via DESC_3_CLN & angka 8 digit
        found_by_slip_desc3 = False
        if desc3_slip_vals:
            desc3_val = _normalize(sus.get(SUSPEND_SLIP_SOURCE_COL, ""))
            # Abaikan jika hanya berupa 1 angka atau string pendek non-slip (misal '4', '2', '5')
            if desc3_val and len(desc3_val) >= 5:
                osb_slip_hits = _exact_match([desc3_val], lookup_osbal[0])
                if osb_slip_hits:
                    matched_osbal_idx.update(osb_slip_hits)
                    found_by_slip_desc3 = True
                    source = "OSBAL"
                    scenario = "Slip only"

        # Coba Like matching berbasis angka 8 digit jika exact slip belum ketemu
        if not found_by_slip_desc3:
            sus_slip_src = (sus.get(SUSPEND_SLIP_SOURCE_COL, "")
                            or sus.get("SLIP_NO_CLN_1", "")
                            or sus.get("SLIP NO", ""))
            sus_slip_str = str(sus_slip_src).strip()
            nums = _extract_slip_numbers(sus_slip_str)
            if nums:
                hits = {i for num in nums if num in osbal_slip_num_idx for i in osbal_slip_num_idx[num]}
                if hits:
                    # Jika > 1 kandidat, coba saring dengan COB prefix jika ada
                    if len(hits) > 1:
                        pfx = _extract_cob_prefix(sus_slip_str)
                        if pfx and pfx in COB_DICTIONARY:
                            allowed_classes = COB_DICTIONARY[pfx]
                            filtered = {
                                i for i in hits
                                if str(osbal_rows[i].get("CLASS_NAME", "")).strip().upper() in allowed_classes
                            }
                            if 0 < len(filtered) < len(hits):
                                hits = filtered
                    matched_osbal_idx.update(hits)
                    found_by_slip_desc3 = True
                    source = "OSBAL"
                    scenario = "Slip only"

        # FASE 1b: Fallback slip dari DESC 4
        if not found_by_slip_desc3:
            desc4_raw = _normalize(sus.get("DESC 4", ""))
            if desc4_raw and re.search(r'[A-Z]', desc4_raw):
                desc4_compact = desc4_raw.replace(" ", "")
                nums_d4 = _extract_slip_numbers(desc4_compact)
                if nums_d4:
                    hits_d4 = {i for num in nums_d4 if num in osbal_slip_num_idx
                               for i in osbal_slip_num_idx[num]}
                    if hits_d4:
                        if len(hits_d4) > 1:
                            pfx_d4 = _extract_cob_prefix(desc4_compact)
                            if pfx_d4 and pfx_d4 in COB_DICTIONARY:
                                allowed_d4 = COB_DICTIONARY[pfx_d4]
                                filtered_d4 = {
                                    i for i in hits_d4
                                    if str(osbal_rows[i].get("CLASS_NAME", "")).strip().upper()
                                    in allowed_d4
                                }
                                if 0 < len(filtered_d4) < len(hits_d4):
                                    hits_d4 = filtered_d4
                        matched_osbal_idx.update(hits_d4)
                        found_by_slip_desc3 = True
                        source = "OSBAL"
                        scenario = "Slip only"

        # FASE 2: Pencarian berdasarkan POLIS
        found_by_polis = False
        if not found_by_slip_desc3:
            polis_vals = _collect_clean_values(sus, eff_polis)
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

        # FASE 3: Pencarian berdasarkan INSURED
        if not matched_osbal_idx:
            # Jika ada nama insured spesifik di DESC 3, prioritaskan nilai tersebut untuk matching
            if desc3_override_insured:
                ins_vals = [desc3_override_insured]
            else:
                ins_vals = _collect_clean_values(sus, insured_sus)
            if ins_vals:
                # 1. Coba OSBAL Exact & Like
                osb_hits = _exact_match(ins_vals, lookup_osbal[2])
                if not osb_hits and lookup_osbal[8] and lookup_osbal[8][0]:
                    osb_hits = _like_match(
                        ins_vals, lookup_osbal[8][0], osbal_rows, insured_osbal,
                        stopwords=_INSURED_STOPWORDS,
                        overlap_threshold=_INSURED_LIKE_THRESHOLD,
                        match_log=_like_log,
                    )
                if osb_hits:
                    matched_osbal_idx.update(osb_hits)
                    source = "OSBAL"
                    scenario = "Insured only"
                
                # 2. Coba SLIPDB Exact & Like
                elif slipdb_rows and not matched_osbal_idx:
                    sld_hits = _exact_match(ins_vals, lookup_slipdb[2])
                    if not sld_hits and lookup_slipdb[8] and lookup_slipdb[8][0]:
                        sld_hits = _like_match(
                            ins_vals, lookup_slipdb[8][0], slipdb_rows, insured_slipdb,
                            stopwords=_INSURED_STOPWORDS,
                            overlap_threshold=_INSURED_LIKE_THRESHOLD,
                            match_log=_like_log,
                        )
                    if sld_hits:
                        res_idx, _ = _resolve_facode([slipdb_rows[idx] for idx in sld_hits], SLIPDB_FACODE_COL, facode_osbal_idx, osbal_rows)
                        if res_idx:
                            matched_osbal_idx.update(res_idx)
                            source = "SLIPDB"
                            scenario = "Insured only"
                            resolved = True
                
                # 3. Coba FACUL Exact & Like
                elif facul_rows and not matched_osbal_idx:
                    fac_hits = _exact_match(ins_vals, lookup_facul[2])
                    if not fac_hits and lookup_facul[8] and lookup_facul[8][0]:
                        fac_hits = _like_match(
                            ins_vals, lookup_facul[8][0], facul_rows, insured_facul,
                            stopwords=_INSURED_STOPWORDS,
                            overlap_threshold=_INSURED_LIKE_THRESHOLD,
                            match_log=_like_log,
                        )
                    if fac_hits:
                        res_idx, _ = _resolve_facode([facul_rows[idx] for idx in fac_hits], FACUL_FACODE_COL, facode_osbal_idx, osbal_rows)
                        if res_idx:
                            matched_osbal_idx.update(res_idx)
                            source = "FACUL"
                            scenario = "Insured only"
                            resolved = True

        # FASE 4: Narrowing utama bertingkat non-destruktif
        if len(matched_osbal_idx) > 1:
            matched_list = list(matched_osbal_idx)

            if "Slip" in scenario:
                # 1. Slip -> coba + Polis
                polis_vals = _collect_clean_values(sus, eff_polis)
                if polis_vals:
                    confirmed = [idx for idx in matched_list if _ref_has_value(osbal_rows[idx], polis_osbal, "polis_ori", polis_vals)]
                    if 0 < len(confirmed) < len(matched_list):
                        matched_list = confirmed
                        scenario = "Polis + Slip"

                # 2. Jika masih >1, coba + Insured
                if len(matched_list) > 1:
                    ins_vals = [desc3_override_insured] if desc3_override_insured else _collect_all_values(sus, insured_sus, "insured_ori")
                    if ins_vals:
                        confirmed = [idx for idx in matched_list if _ref_has_value(osbal_rows[idx], insured_osbal, "insured_ori", ins_vals)]
                        if 0 < len(confirmed) < len(matched_list):
                            insured_vals_before: set = set()
                            for _idx in matched_list:
                                for _c in insured_osbal:
                                    _v = _normalize(osbal_rows[_idx].get(_c, ""))
                                    if _v:
                                        insured_vals_before.add(_v)
                            matched_list = confirmed
                            if len(insured_vals_before) > 1 and "Insured" not in scenario:
                                base = scenario.replace(" only", "") if scenario.endswith(")") or not scenario.endswith("only") else scenario
                                scenario = base + " + Insured"

            elif "Polis" in scenario:
                # 1. Polis -> coba + Slip
                slip_vals = _collect_clean_values(sus, slip_sus if slip_sus else desc3_slip_vals)
                confirmed = [idx for idx in matched_list if _ref_has_value(osbal_rows[idx], slip_osbal, "slip_ori", slip_vals)] if slip_vals else []
                if not confirmed:
                    # Cek angka 8 digit jika exact slip tidak ketemu
                    sus_slip_src = (sus.get(SUSPEND_SLIP_SOURCE_COL, "") or sus.get("SLIP_NO_CLN_1", "") or sus.get("SLIP NO", ""))
                    nums = _extract_slip_numbers(str(sus_slip_src).strip())
                    if nums:
                        confirmed = [idx for idx in matched_list if any(num in _osbal_slip_text(osbal_rows[idx]) for num in nums)]
                if 0 < len(confirmed) < len(matched_list):
                    matched_list = confirmed
                    scenario = "Polis + Slip"

                # 2. Jika masih >1, coba + Insured
                if len(matched_list) > 1:
                    ins_vals = [desc3_override_insured] if desc3_override_insured else _collect_all_values(sus, insured_sus, "insured_ori")
                    if ins_vals:
                        confirmed = [idx for idx in matched_list if _ref_has_value(osbal_rows[idx], insured_osbal, "insured_ori", ins_vals)]
                        if 0 < len(confirmed) < len(matched_list):
                            insured_vals_before: set = set()
                            for _idx in matched_list:
                                for _c in insured_osbal:
                                    _v = _normalize(osbal_rows[_idx].get(_c, ""))
                                    if _v:
                                        insured_vals_before.add(_v)
                            matched_list = confirmed
                            if len(insured_vals_before) > 1 and "Insured" not in scenario:
                                base = scenario.replace(" only", "") if scenario.endswith(")") or not scenario.endswith("only") else scenario
                                scenario = base + " + Insured"

            elif "Insured" in scenario:
                # 1. Insured -> coba + Slip
                slip_vals = _collect_clean_values(sus, slip_sus if slip_sus else desc3_slip_vals)
                confirmed = [idx for idx in matched_list if _ref_has_value(osbal_rows[idx], slip_osbal, "slip_ori", slip_vals)] if slip_vals else []
                if not confirmed:
                    # Cek angka 8 digit
                    sus_slip_src = (sus.get(SUSPEND_SLIP_SOURCE_COL, "") or sus.get("SLIP_NO_CLN_1", "") or sus.get("SLIP NO", ""))
                    nums = _extract_slip_numbers(str(sus_slip_src).strip())
                    if nums:
                        confirmed = [idx for idx in matched_list if any(num in _osbal_slip_text(osbal_rows[idx]) for num in nums)]
                if 0 < len(confirmed) < len(matched_list):
                    matched_list = confirmed
                    scenario = "Slip + Insured"

                # 2. Jika masih >1, coba + Polis
                if len(matched_list) > 1:
                    polis_vals = _collect_clean_values(sus, eff_polis)
                    if polis_vals:
                        confirmed = [idx for idx in matched_list if _ref_has_value(osbal_rows[idx], polis_osbal, "polis_ori", polis_vals)]
                        if 0 < len(confirmed) < len(matched_list):
                            matched_list = confirmed
                            if "Polis" not in scenario:
                                scenario = "Polis + " + scenario

            # Step Tambahan: Jika masih >1 kandidat, coba saring dengan COB dictionary
            if len(matched_list) > 1:
                sus_slip_src = (sus.get(SUSPEND_SLIP_SOURCE_COL, "") or sus.get("SLIP_NO_CLN_1", "") or sus.get("SLIP NO", ""))
                pfx = _extract_cob_prefix(str(sus_slip_src).strip())
                if pfx and pfx in COB_DICTIONARY:
                    allowed_classes = COB_DICTIONARY[pfx]
                    kept_idx = [idx for idx in matched_list if str(osbal_rows[idx].get("CLASS_NAME", "")).strip().upper() in allowed_classes]
                    if 0 < len(kept_idx) < len(matched_list):
                        matched_list = kept_idx
            # Kasus khusus Wintrad Jaya: filter kandidat by insured ke FAC CODE aktif (25FWA3L0)
            if len(matched_list) > 1:
                sus_ins_name = _normalize(sus.get("INSURED_CLN_1", "") or sus.get("INSURED", ""))
                if "WINTRAD" in sus_ins_name:
                    w_cand = [idx for idx in matched_list if "25FWA3L0" in str(osbal_rows[idx].get("CCOS_REF_CODE", "")) or float(osbal_rows[idx].get("CCOS_BAL_DUE", 0) or 0) > 0]
                    if w_cand:
                        matched_list = w_cand
                        scenario = "Insured only"

            matched_osbal_idx = set(matched_list)

        # FASE 5: Validasi Wajib (Currency & Periode) + Strict Rematch
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
                    scenario = scen
                    matched_list = hits
                    is_beda_curr = False
                    is_beda_per  = False

            matched_osbal_idx = set(matched_list)

        # FASE 6: Matching Bordero (COB PAR & Earthquake)
        is_bordero = False
        if bordero_slip_index:
            bordero_facs, ok = _find_fac_codes_in_bordero(
                sus, bordero_idx, bordero_slip_index, eff_slip
            )
            if ok and bordero_facs:
                res_idx = []
                for fac in bordero_facs:
                    if fac in facode_osbal_idx:
                        res_idx.extend(facode_osbal_idx[fac])

                if res_idx:
                    # Validasi Currency & Periode untuk kandidat Bordero
                    b_curr_diff = False
                    b_per_diff  = False

                    sus_curr = _normalize(sus.get(SUSPEND_CURR_COL, ""))
                    if sus_curr:
                        filtered_curr = [i for i in res_idx if _normalize(osbal_rows[i].get(OSBAL_CURR_COL, "")) == sus_curr]
                        if not filtered_curr:
                            b_curr_diff = True
                        elif len(filtered_curr) < len(res_idx):
                            res_idx = filtered_curr

                    if not b_curr_diff:
                        sus_date = effective_date
                        if pd.notna(sus_date):
                            filtered_per = []
                            for i in res_idx:
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
                                b_per_diff = True
                            elif len(filtered_per) < len(res_idx):
                                res_idx = filtered_per

                    # SELALU override hasil sebelumnya karena Bordero dianggap lebih valid/benar
                    matched_osbal_idx = set(res_idx)
                    source = "OSBAL"
                    scenario = "Slip only (Bordero)"
                    is_bordero = True
                    resolved = True
                    is_beda_curr = b_curr_diff
                    is_beda_per  = b_per_diff


        # Post-matching normalisasi label
        if is_beda_curr or is_beda_per:
            # Beda Periode dan Beda Currency ditangani sama (scenario awal dipertahankan)
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

        fac_codes = {_normalize(r.get(OSBAL_FACODE_COL, "")) for r in osbal_ref if r}
        fac_codes.discard("")

        raw_results.append({
            "suspend": sus, "source": source, "scenario": scenario,
            "is_bordero": is_bordero,
            "osbal_idx": list(matched_osbal_idx),
            "osbal_count": osbal_count, "resolved": resolved, "fac_codes": fac_codes,
            "_like_log": [m for _, m in _like_log],
            "_is_beda_curr": is_beda_curr,
            "_is_beda_periode": is_beda_per,
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
    _EXCL = {"Unmatching", "Beda Periode", "UNMATCHED"}
    for r in raw_results:
        if len(r["fac_codes"]) == 1 and not any(exc in r["scenario"] for exc in _EXCL):
            claimed |= r["fac_codes"]

    for r in raw_results:
        if len(r["fac_codes"]) > 1:
            remaining = r["fac_codes"] - claimed
            if len(remaining) == 1 and not any(exc in r["scenario"] for exc in _EXCL):
                # Pertahankan multi fac_codes agar tetap terdeteksi matching >1 fac code
                r["_originally_multi_fac"] = True
                claimed |= remaining

    groups: dict = {}
    final:  list = []

    for r in raw_results:
        r.setdefault("suspend_count", 1)
        # Beda Currency: source ada, tapi tidak di-group (osbal_count = 0)
        _is_beda_curr_row = r.get("_is_beda_curr", False)
        if (len(r["fac_codes"]) == 1 and r["source"] is not None
                and not any(exc in r["scenario"] for exc in _EXCL)
                and not _is_beda_curr_row):
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
        )
        if "_merged_amount_ori" in r:
            out["_MERGED_AMOUNT_ORI"] = r["_merged_amount_ori"]
        # Tandai baris Beda Currency dan Periode untuk _compute_derived_cols
        out["_IS_BEDA_CURR"] = r.get("_is_beda_curr", False)
        out["_IS_BEDA_PERIODE"] = r.get("_is_beda_periode", False)
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

    # Append FLAG_PROD ke SKENARIO if FLAG_PROD matches specific scenarios
    override_flags = ["Beda Currency", "Beda Periode", "Matching >1 fac code", "Matching >1 Fac code OC MC"]
    cond = df["FLAG_PROD"].isin(override_flags)
    def _format_skenario(row):
        sken = str(row["SKENARIO"])
        flag = str(row["FLAG_PROD"])
        
        if not sken or sken in ["UNMATCHED", "Unmatching"]:
            return flag
            
        if " (Bordero)" in sken:
            base = sken.replace(" (Bordero)", "").strip()
            if flag == "Matching >1 fac code":
                return f"{base} (> 1 Fac Bordero)"
            else:
                return f"{base} ({flag} Bordero)"
        else:
            if flag == "Matching >1 fac code":
                return f"{sken} (> 1 Fac)"
            return f"{sken} ({flag})"

    df.loc[cond, "SKENARIO"] = df.loc[cond].apply(_format_skenario, axis=1)

    for col in FINAL_COLUMNS:
        if col not in df.columns:
            df[col] = ""
    df = df[FINAL_COLUMNS]

    # [6] Export
    print(f"\n[6/6] Saving to: {OUTPUT_FILE} ...", flush=True)
    for c in df.select_dtypes(include=['object', 'string']).columns:
        df[c] = df[c].astype(str).str.replace(r'[\x00-\x08\x0B\x0C\x0E-\x1F]', '', regex=True)

    save_path = OUTPUT_FILE
    try:
        df.to_excel(save_path, index=False)
    except PermissionError:
        save_path = os.path.join(_BASE_DIR, "final_output_ibs_v1_new.xlsx")
        print(f"  [PERINGATAN] File {OUTPUT_FILE} sedang dibuka di Excel! Menyimpan ke: {save_path} ...", flush=True)
        df.to_excel(save_path, index=False)

    try:
        from openpyxl import load_workbook
        workbook = load_workbook(save_path)
        worksheet = workbook.active
        text_cols = [
            "POLIS_ORI", "POLICY_CLEAN_1", "SERTIF_CLEAN_1",
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
        df.to_sql("SUSPENSE_DATA_SUSPENSE_IBS_V1", con=engine, if_exists="append", index=False)
        print(f"      -> {len(df):,} rows berhasil di-export ke 'SUSPENSE_DATA_SUSPENSE_IBS_V1'")
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
