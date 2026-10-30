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
# KONFIGURASI  (CEDANT: TOKIO MARINE)
#
# Diturunkan dari matching_ibs.py. Delta utama vs IBS:
#   1. RULES TOKIO = POLIS lalu INSURED saja. Slip TIDAK dipakai
#      (tidak ada FASE 1 via DESC_3_CLN, tidak ada narrowing via slip).
#   2. Nama kolom Suspend mengikuti file tokio_output_suspend.xlsx
#      ("clean polis N", "clean insured N", "Certificate", dst).
#   3. Data Suspend difilter: HANYA cedant Tokio yang diproses & masuk output
#      (baris cedant lain dibuang).
#   4. Bordero / Open Cover Marine Cargo: tidak ada (sama seperti IBS).
#
# Urutan fase Tokio:
#   FASE 1  POLIS   (OSBAL -> SLIPDB -> FACUL)
#   FASE 2  INSURED (OSBAL -> SLIPDB -> FACUL)  <- fallback jika polis miss
#   FASE 3  Narrowing wajib: Currency & Periode (+ strict rematch)
#   FASE 4  Narrowing bertingkat non-destruktif (polis -> insured)
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
SUSPEND_FILE = _find_file_in(_CLEAN_DIR, ["tokio_output_suspend_new.xlsx", "tokio_output_suspend.xlsx", "tokio_output_suspense.xlsx"])
# OSBAL & FACUL: sesuaikan nama file jika berbeda. Kandidat kedua = file OSBAL/FACUL
# hasil cleansing yang sudah ada (biasanya berisi semua cedant, difilter Tokio di bawah).
OSBAL_FILE   = _find_file_in(_CLEAN_DIR, ["tokio_output_osbal_new.xlsx", "tokio_output_osbal.xlsx", "ibs_output_osbal_V2.xlsx"])
FACUL_FILE   = _find_file_in(_CLEAN_DIR, ["tokio_output_facul.xlsx", "ibs_output_facul_V2.xlsx"])
SLIPDB_FILE  = _find_file_in(_RAW_DIR, [
    "5. Database Rekap RI Slip.xlsx",
    "5.Database Rekap RI Slip.xlsx",
    "Database Rekap RI Slip.xlsx",
])
OUTPUT_FILE  = os.path.join(_BASE_DIR, "matching_tokio.xlsx")

# [TOKIO] BORDERO_FILE Open Cover Marine Cargo
BORDERO_FILE = _find_file_in(_CLEAN_DIR, ["tokio_output_bordero.xlsx", "bordero_tokio.xlsx"])

# Kolom FAC CODE di masing-masing tabel referensi
OSBAL_FACODE_COL  = "CCOS_REF_CODE"
FACUL_FACODE_COL  = "FAC_CODE"
SLIPDB_FACODE_COL = "FAC_CODE"   # dapat di-override runtime sesuai header aktual

# [TOKIO] Filter cedant: hanya baris yang nama cedant-nya mengandung kata ini.
# Dicek ke semua kolom yang namanya mengandung "CEDANT" (CEDANT NAME, CEDANT SHRT NAME).
CEDANT_KEYWORD = "TOKIO"
# True  -> tabel referensi (OSBAL/FACUL/SLIPDB) ikut difilter Tokio JIKA punya kolom cedant
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
# [TOKIO] Baris yang mengandung salah satu kata ini TIDAK ikut proses matching:
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

# [TOKIO] FINAL_COLUMNS: Kolom Polis (5 kolom) dan Sertif (3 kolom)
FINAL_COLUMNS = [
    "CCOS_DOC_NO", "CCOS_REF_CODE",
    "RECEIPT NO", "CREDIT NOTES", "DETAIL RINCIAN NO", "RECEIPT DATE",
    "CEDANT NAME", "CEDANT SHRT NAME",
    "INSURED_ORI", "INSURED_1", "INSURED_2",
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


# [TOKIO] SKENARIO yang boleh muncul di output.
_ALLOWED_BASE_SCENARIOS = (
    "Unmatching", "Insured only", "Polis only", "Polis + Insured", "Polis + Slip", "Slip only",
    "Polis only -> Bordero OC Override", "Insured only -> Bordero OC Override",
    "Polis + Insured -> Bordero OC Override", "Polis + Slip -> Bordero OC Override",
    "Slip only -> Bordero OC Override",
)
_SPECIAL_SCENARIO_SUFFIXES = (
    "> 1 Fac", "Beda Currency", "Beda Periode"
)

_ALLOWED_SCENARIOS = tuple(
    list(_ALLOWED_BASE_SCENARIOS) +
    [f"{base} ({suffix})" for base in _ALLOWED_BASE_SCENARIOS for suffix in _SPECIAL_SCENARIO_SUFFIXES]
)


# Regex untuk mengekstrak nomor polis Tokio Marine
_POLIS_FULL_RE = re.compile(r'^(?:TMD|PLD|MDD|BTD|TMF)?([A-Z]{3,4})?(\d{2})([A-Z]\d{7})$')
_POLIS_YEAR_CORE_RE = re.compile(r'^(\d{2})([A-Z]\d{7})$')
_CORE_NUMBER_RE = re.compile(r'([A-Z]\d{7})$')


def _clean_single_polis(p: str) -> str:
    """Bersihkan satu string nomor polis (hapus suffix -xxx-yy, tanda pemisah, dll)."""
    p = str(p).strip()
    p = re.sub(r"-\d{3}-\d+$", "", p)
    p = re.sub(r"-000-\d+$", "", p)
    while True:
        s = re.sub(r"-\d+(?:/\d+)?$", "", p)
        if s == p:
            break
        p = s
    return re.sub(r"[^A-Za-z0-9]", "", p).upper()


def _get_polis_year(polis_str: str) -> str:
    """Ekstrak 2 digit tahun dari nomor polis (mis. 'TMDEMCB24E0006738' -> '24')."""
    if not polis_str or (isinstance(polis_str, float) and pd.isna(polis_str)):
        return ""
    clean = re.sub(r'[^A-Z0-9]', '', str(polis_str).upper())
    m = _POLIS_FULL_RE.match(clean)
    if m and m.group(2):
        return m.group(2)
    m2 = _POLIS_YEAR_CORE_RE.match(clean)
    if m2:
        return m2.group(1)
    m3 = re.search(r'[\.\-\/](\d{2})[\.\-\/]', str(polis_str).upper())
    if m3:
        return m3.group(1)
    return ""


def _get_row_year(row: dict) -> str:
    """Ekstrak 2 digit tahun dari baris referensi (OSBAL/FACUL/SLIPDB)."""
    for col in ['FAC_POLICY_NO', 'FAC_POLICY_NO_CLEAN_1', 'CLSDT_POLICY_NO', 'POLIS', 'POLIS_CLEAN']:
        val = row.get(col, '')
        yr = _get_polis_year(val)
        if yr:
            return yr
    fac_date = str(row.get('FAC_COM_DATE', '')).strip()
    if len(fac_date) >= 4:
        if fac_date[:4].isdigit():
            return fac_date[2:4]
        elif '/' in fac_date:
            parts = fac_date.split('/')
            if len(parts) == 3 and len(parts[2]) >= 4 and parts[2][:4].isdigit():
                return parts[2][2:4]
    ref = str(row.get('CCOS_REF_CODE', '')).strip()
    if len(ref) >= 2 and ref[:2].isdigit():
        return ref[:2]
    return ""


def _get_polis_variations(clean_val: str) -> list:
    """Hasilkan variasi penulisan nomor polis dengan tetap menjaga informasi tahun."""
    if not clean_val:
        return []
    res = [clean_val]
    m = _POLIS_FULL_RE.match(clean_val)
    if m:
        cob, yr, core = m.groups()
        if cob and yr and core:
            res.append(f"{cob}{yr}{core}")
            res.append(f"TMD{cob}{yr}{core}")
        if yr and core:
            res.append(f"{yr}{core}")
    seen = set()
    final = []
    for r in res:
        if r and r not in seen:
            seen.add(r)
            final.append(r)
    return final


def _extract_core_number(polis_clean: str) -> str:
    """Ekstrak nomor inti dari polis bersih, mis. 'TMDMINL24M5751934' -> 'M5751934'.

    Juga mengembalikan nomor inti jika string sudah berupa nomor inti (1 huruf + 7 digit),
    misalnya 'F5011636'.
    """
    if not polis_clean:
        return ""
    m = _CORE_NUMBER_RE.search(str(polis_clean).strip().upper())
    return m.group(1) if m else ""


def _expand_polis_string(raw: str) -> list:
    """Ekstrak dan breakdown format nomor polis dengan '+' atau format multi-polis '/'.

    Contoh input:
      'F5011636 + 637 + 630 + 631 + 737 + 738 + 410 + 411'
      'A0873493/A0885832/A0967466'
    """
    if not raw or (isinstance(raw, float) and pd.isna(raw)):
        return []
    s_raw = str(raw).strip()
    if not s_raw or s_raw == "-":
        return []

    # Cek apakah raw berisi multiple polis yang dipisah '+'
    plus_parts = [p.strip() for p in s_raw.split('+') if p.strip()]
    parts = []
    for pp in plus_parts:
        # Cek jika dipisah '/' seperti 'A0873493/A0885832/A0967466' (bukan prefix seperti TMD/EMCB/...)
        if '/' in pp and not re.match(r'^(?:TMD|PLD|MDD|BTD|TMF)\/', pp, re.IGNORECASE):
            slash_parts = [sp.strip() for sp in pp.split('/') if sp.strip()]
            if all(len(sp) >= 7 for sp in slash_parts):
                parts.extend(slash_parts)
            else:
                parts.append(pp)
        else:
            parts.append(pp)

    results = []
    base_core = ""

    for part in parts:
        if re.match(r"^(?:VAR|VARIOUS|P\d+|CANCEL|NONE|NAN|-)$", part, re.IGNORECASE):
            continue
        cleaned = _clean_single_polis(part)
        core = _extract_core_number(cleaned) or _extract_core_number(part)
        if core:
            results.append(core)
            base_core = core
            if cleaned and cleaned != core:
                results.append(cleaned)
            continue
        if part.isdigit() and base_core:
            k = len(part)
            if k < len(base_core):
                expanded = base_core[:-k] + part
                results.append(expanded)
                continue
            else:
                results.append(part)
                continue
        if cleaned and len(cleaned) >= 6:
            results.append(cleaned)

    # Filter token yang terlalu pendek atau hanya prefix/COB
    filtered = []
    prefixes = {"TMD", "PLD", "MDD", "BTD", "TMF"}
    for r in results:
        if len(r) < 6:
            continue
        if r in prefixes or r in COB_PREFIX_GROUPS:
            continue
        filtered.append(r)

    # Deduplikasi stabil
    seen = set()
    final = []
    for r in filtered:
        if r and r not in seen:
            seen.add(r)
            final.append(r)
    return final


# =============================================================================
# COB (CLASS OF BUSINESS) DICTIONARY & MATCHING
# =============================================================================
# Karakteristik substring nomor polis -> Nama COB resmi
COB_PREFIX_GROUPS = {
    # Property
    "FPAR": "PROPERTY", "FIAR": "PROPERTY", "FISP": "PROPERTY",
    # Earthquake
    "FEAQ": "EARTHQUAKE", "FEQI": "EARTHQUAKE", "FEAW": "EARTHQUAKE",
    # Terrorism
    "FRCC": "TERRORISM",
    # Marine
    "MINL": "MARINE_CARGO", "MIMP": "MARINE_CARGO", "MEXP": "MARINE_CARGO", "MFFL": "MARINE_CARGO",
    "HMAR": "MARINE_HULL", "CMHL": "MARINE_HULL",
    # Motor
    "AMBL": "MOTOR",
    # Engineering
    "EMCB": "MB", "FMBI": "MB", "ECAR": "EAR", "EEAR": "CAR", "ECPM": "CPM",
    # Casualty
    "VMNY": "CASUALTY", "LGEN": "CASUALTY", "LPRO": "CASUALTY", "LPIC": "CASUALTY", "LDNO": "CASUALTY",
}

OSBAL_CLASS_TO_GROUP = {
    "PAR": "PROPERTY", "IAR": "PROPERTY", "FIRE": "PROPERTY",
    "PROPERTY ALL RISK": "PROPERTY", "INDUSTRIAL ALL RISK": "PROPERTY",
    "EARTHQUAKE": "EARTHQUAKE", "EQ": "EARTHQUAKE",
    "TERRORISM": "TERRORISM", "SABOTAGE": "TERRORISM", "TERRORISM AND SABOTAGE": "TERRORISM",
    "MARINE CARGO": "MARINE_CARGO", "MARINE HULL": "MARINE_HULL", "MOTOR": "MOTOR",
    "MB": "MB", "MACHINERY BREAKDOWN": "MB",
    "EAR": "EAR", "CAR": "CAR", "CPM/CPE": "CPM", "HEAVY EQUIPMENT": "CPM",
    "CASUALTY": "CASUALTY", "MONEY INSURANCE": "CASUALTY", "CASH IN SAFE": "CASUALTY",
    "COMPREHENSIVE L": "CASUALTY", "PUBLIC LIABILITY": "CASUALTY", "PRODUCT LIAB": "CASUALTY",
    "PROF INDEMNITY": "CASUALTY", "DIRECTOR OFF": "CASUALTY",
}

# =============================================================================
# PELENGKAPAN FORMAT POLIS OSBAL (Dictionary Tokio M)
# =============================================================================
# Mapping CLASS_NAME OSBAL -> prefix TMD berdasarkan dictionary tokio m.
# Digunakan untuk melengkapi nomor polis OSBAL yang tidak lengkap formatnya
# (tidak ada TMD/XXX/YY- prefix) agar bisa dicocokkan dengan polis Suspend
# yang sudah dalam format lengkap.
# Ref: readme/dictionary tokio m.txt
_OSBAL_CLASS_TO_TMD_PREFIX: dict = {
    "MARINE CARGO":    "TMDMINL",
    "PAR":             "TMDFPAR",
    "IAR":             "TMDFIAR",
    "EAR":             "TMDECAR",
    "COMPREHENSIVE L": "TMDLGEN",
    "MB":              "TMDEMCB",
    "CAR":             "TMDEEAR",
    "PROF INDEMNITY":  "TMDLPIC",
    "MARINE HULL":     "TMDHMAR",
}

# Pola format polis pendek yang akan dilengkapi:
#   Case 1: 2-digit tahun + 1 huruf + 7 angka  (10 char)  → prefix + '26X0000000'
#   Case 2: 1 huruf + 7 angka                  ( 8 char)  → prefix + tahun + 'X0000000'
_RE_POLIS_YEAR_CORE = re.compile(r'^(\d{2})([A-Z]\d{7})$')   # Case 1: 26F0003746
_RE_POLIS_CORE_ONLY = re.compile(r'^([A-Z]\d{7})$')            # Case 2: F0003746


def _extract_cob_from_suspend(row: dict) -> set:
    """Ekstrak set group COB dari Suspend berdasarkan substring karakteristik polis."""
    sus_groups = set()
    for col in ["POLIS", "POLIS_CLEAN", "polis_ori", "clean polis 1"]:
        val = str(row.get(col, "")).upper()
        if val and val != "NAN" and val != "NONE" and val != "-":
            for prefix, group in COB_PREFIX_GROUPS.items():
                if prefix in val:
                    sus_groups.add(group)
    return sus_groups


_COB_COMPATIBILITY = {
    "PROPERTY":   {"PROPERTY", "EARTHQUAKE", "TERRORISM"},
    "EARTHQUAKE": {"PROPERTY", "EARTHQUAKE"},
    "TERRORISM":  {"PROPERTY", "TERRORISM"},
    "CAR":        {"CAR", "EAR"},
    "EAR":        {"EAR", "CAR"},
}

def _osbal_matches_cob(osbal_row: dict, sus_groups: set) -> bool:
    """Cek apakah CLASS_NAME atau Prefix baris OSBAL sesuai dengan group COB Suspend."""
    if not sus_groups:
        return True
        
    osbal_groups = set()
    for col in ["FAC_POLICY_NO", "FAC_POLICY_NO_CLEAN_1", "FAC_POLICY_NO_CLEAN_2", "FAC_POLICY_NO_CLEAN_3"]:
        val = str(osbal_row.get(col, "")).upper()
        if val and val != "NAN":
            for prefix, group in COB_PREFIX_GROUPS.items():
                if prefix in val:
                    osbal_groups.add(group)
                    
    cname = str(osbal_row.get("CLASS_NAME", "")).strip().upper()
    if cname in OSBAL_CLASS_TO_GROUP:
        osbal_groups.add(OSBAL_CLASS_TO_GROUP[cname])
        
    if not osbal_groups:
        return True
        
    if sus_groups.intersection(osbal_groups):
        return True

    for sg in sus_groups:
        allowed = _COB_COMPATIBILITY.get(sg, {sg})
        if allowed.intersection(osbal_groups):
            return True

    return False


def _augment_polis_with_core(polis_vals: list) -> list:
    """Tambahkan variasi expanded polis dan core number ke daftar nilai polis."""
    if not polis_vals:
        return polis_vals
    augmented = list(polis_vals)
    seen = set(augmented)
    for pv in polis_vals:
        for var in _get_polis_variations(pv):
            if var and var not in seen:
                augmented.append(var)
                seen.add(var)
        for ev in _expand_polis_string(pv):
            if ev and ev not in seen:
                augmented.append(ev)
                seen.add(ev)
            for ev_var in _get_polis_variations(ev):
                if ev_var and ev_var not in seen:
                    augmented.append(ev_var)
                    seen.add(ev_var)
            core = _extract_core_number(ev)
            if core and core not in seen:
                augmented.append(core)
                seen.add(core)
        core = _extract_core_number(pv)
        if core and core not in seen:
            augmented.append(core)
            seen.add(core)
    return augmented


def _get_osbal_polis_augmented(row: dict) -> list:
    """Kembalikan variasi format polis lengkap untuk baris OSBAL yang nomor polisnya pendek.

    Rules dari 'readme/dictionary tokio m.txt':
      Jika FAC_POLICY_NO di OSBAL tidak lengkap (tidak ada prefix TMD/XXX/YY-),
      tambahkan prefix berdasarkan CLASS_NAME sesuai dictionary, dengan dua case:
        Case 1: format '10-char' (YY + 1 huruf + 7 angka, e.g. '26F0003746')
                → ditambah prefix saja   : TMDFPAR + 26F0003746 = TMDFPAR26F0003746
        Case 2: format '8-char'  (1 huruf + 7 angka, e.g. 'F0003746')
                → ditambah prefix + tahun: TMDFPAR + 26 + F0003746 = TMDFPAR26F0003746
                  (tahun 2-digit diambil dari kolom FAC_COM_DATE)

    Hasil dikembalikan sebagai list string; kosong jika tidak ada augmentasi.
    Tidak mengubah nilai asli, hanya menambahkan variasi untuk keperluan indexing.
    """
    raw_pol = str(row.get("FAC_POLICY_NO", "")).strip()
    if not raw_pol or raw_pol.upper() in ("NAN", "", "-"):
        return []
    cls = str(row.get("CLASS_NAME", "")).strip().upper()
    prefix = _OSBAL_CLASS_TO_TMD_PREFIX.get(cls, "")
    if not prefix:
        return []
    clean = re.sub(r"[^A-Z0-9]", "", raw_pol.upper())
    augmented = []
    m1 = _RE_POLIS_YEAR_CORE.match(clean)
    if m1:
        # Case 1: 10-char (YY + letter + 7 digits)
        augmented.append(prefix + clean)
    else:
        m2 = _RE_POLIS_CORE_ONLY.match(clean)
        if m2:
            # Case 2: 8-char, butuh tahun dari FAC_COM_DATE
            fac_date = str(row.get("FAC_COM_DATE", "")).strip()
            # FAC_COM_DATE bisa berupa 'YYYY-MM-DD', 'YYYY-MM-DD HH:MM:SS', atau 'DD/MM/YYYY'
            year2 = ""
            if len(fac_date) >= 4:
                if fac_date[:4].isdigit():          # YYYY-... format
                    year2 = fac_date[2:4]
                elif "/" in fac_date:               # DD/MM/YYYY
                    parts = fac_date.split("/")
                    if len(parts) == 3 and len(parts[2]) >= 4 and parts[2][:4].isdigit():
                        year2 = parts[2][2:4]
            if year2:
                augmented.append(prefix + year2 + clean)
    return augmented


def _clean_base_match(scenario: str) -> str:
    s = str(scenario or "").strip()
    for to_remove in [
        " -> Bordero OC Override", " (Bordero)", " (Beda Currency)",
        " (Beda Periode)", " (> 1 Fac)", " (> 1 Fac Bordero)",
        " (>1 fac)", " (>1 fac OC MC)", " (DESC 1)", " + Amount",
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

    scen = f"{base} -> Bordero OC Override" if is_bordero else base

    if is_beda_curr:
        return f"{scen} (Beda Currency)"
    elif is_beda_per:
        return f"{scen} (Beda Periode)"
    elif num_fac > 1:
        return f"{scen} (> 1 Fac)"
    else:
        return scen


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
        raw_val = str(row.get(col, "")).strip()
        if not raw_val or raw_val.upper() in ("NAN", "NONE", "-"):
            continue
        if ";" in raw_val:
            for sub_val in raw_val.split(";"):
                norm = _normalize(sub_val)
                if norm and norm not in seen:
                    seen.add(norm)
                    result.append(norm)
        val = _normalize(raw_val)
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

    # [IBS DELTA] slip standar tidak dipakai sebagai fase independent
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

def _build_exact_index(rows: list, cols: list, excluded: set = None, is_polis: bool = False) -> dict:
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
                    for v_var in _get_polis_variations(val):
                        if v_var and v_var != val:
                            index.setdefault(v_var, []).append(i)
                    expanded = _expand_polis_string(raw_val)
                    for ev in expanded:
                        if ev and ev != val:
                            index.setdefault(ev, []).append(i)
                        for ev_var in _get_polis_variations(ev):
                            if ev_var and ev_var != ev:
                                index.setdefault(ev_var, []).append(i)
                        c = _extract_core_number(ev)
                        if c and c != val:
                            index.setdefault(c, []).append(i)
                    c_val = _extract_core_number(val)
                    if c_val and c_val != val:
                        index.setdefault(c_val, []).append(i)
                # [DICT TOKIO M] Pelengkapan format polis OSBAL yang pendek:
                # Jika FAC_POLICY_NO tidak punya prefix TMD, lengkapi menggunakan
                # CLASS_NAME + FAC_COM_DATE sesuai kaidah dictionary tokio m.
                if col == "FAC_POLICY_NO":
                    for aug_val in _get_osbal_polis_augmented(row):
                        aug_norm = _normalize(aug_val)
                        if aug_norm and aug_norm not in (val, ""):
                            index.setdefault(aug_norm, []).append(i)
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
        _build_exact_index(rows, slip_cols,    excluded=excluded),   # [0] slip
        _build_exact_index(rows, polis_cols,   excluded=excluded, is_polis=True),   # [1] polis (dengan breakdown & core)
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


def _build_bordero_lookup(rows: list) -> tuple:
    """Bangun lookup index untuk data Bordero Tokio Marine: (core_idx, polis_idx)."""
    core_idx = {}
    polis_idx = {}
    for r in rows:
        c = r.get("CORE_M")
        if not c:
            c = _extract_core_number(r.get("POLIS", ""))
        c_norm = _normalize(c)
        if c_norm:
            core_idx.setdefault(c_norm, []).append(r)

        p = _normalize(r.get("POLIS_CLEAN") or r.get("POLIS", ""))
        if p:
            polis_idx.setdefault(p, []).append(r)
    return core_idx, polis_idx


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

        matched_method = None
        # Evaluasi per-item ref_values (untuk kolom insured clean terpisah/QQ), lalu gabungan seluruhnya
        candidate_eval_lists = [[rv] for rv in ref_values if rv] + ([ref_values] if len(ref_values) > 1 else [])
        for r_vals in candidate_eval_lists:
            ref_tokens: frozenset = frozenset()
            for rv in r_vals:
                ref_tokens = ref_tokens | _tokenize(rv)
            if not ref_tokens:
                continue

            sig_ref_tokens = ref_tokens - stopwords

            # -- Kriteria 2: SUBSET (pakai ref_tokens lengkap, termasuk yg di-stopword) --
            if sig_query_tokens and (sig_query_tokens.issubset(ref_tokens) or (strict and sig_ref_tokens and sig_ref_tokens <= sig_query_tokens)):
                if not single_generic:
                    if strict and len(sig_query_tokens) == 1 and len(sig_ref_tokens) > 2:
                        pass
                    else:
                        matched_method = "SUBSET"
                        break

            # -- Kriteria 3: OVERLAP --
            if not (sig_query_tokens and sig_ref_tokens):
                continue
            shared = sig_query_tokens & sig_ref_tokens
            if not shared:
                continue

            if not strict:
                denom = min(len(sig_query_tokens), len(sig_ref_tokens))
                if denom > 0 and (len(shared) / denom) >= overlap_threshold:
                    matched_method = "OVERLAP"
                    break
                continue

            contained = sig_query_tokens <= sig_ref_tokens or sig_ref_tokens <= sig_query_tokens
            if len(shared) < 2 and not contained:
                continue
            w_shared = sum(_w(t) for t in shared)
            if w_shared < min_w:
                continue
            denom_w = min(q_weight, sum(_w(t) for t in sig_ref_tokens))
            if denom_w > 0 and (w_shared / denom_w) >= overlap_threshold:
                matched_method = "OVERLAP"
                break

        if matched_method:
            matched_set.setdefault(i, matched_method)

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
    sus_yr: str = "",
) -> tuple:
    """Rematch ketat dengan filter currency + periode + COB + Tahun (R1-R5 standar)."""
    sus_curr = _normalize(suspend_row.get(SUSPEND_CURR_COL, ""))

    def _strict_filter(idxs):
        res = []
        for i in idxs:
            if sus_curr and _normalize(osbal_rows[i].get(OSBAL_CURR_COL, "")) != sus_curr:
                continue
            if sus_cob and not _osbal_matches_cob(osbal_rows[i], sus_cob):
                continue
            if sus_yr:
                row_yr = _get_row_year(osbal_rows[i])
                if row_yr and row_yr != sus_yr:
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
    if lookup_osbal[7] and lookup_osbal[7][0] and like_cols:
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

    if lookup_osbal[8] and lookup_osbal[8][0] and insured_vals_r5:
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
# BORDERO / OPEN COVER ENGINE (OVERRIDE & NARROWING)
# =============================================================================

_MONTH_NAME_TO_INT = {
    "JANUARI": 1, "JANUARY": 1, "JAN": 1,
    "FEBRUARI": 2, "FEBRUARY": 2, "FEB": 2,
    "MARET": 3, "MARCH": 3, "MAR": 3,
    "APRIL": 4, "APR": 4,
    "MEI": 5, "MAY": 5,
    "JUNI": 6, "JUNE": 6, "JUN": 6,
    "JULI": 7, "JULY": 7, "JUL": 7,
    "AGUSTUS": 8, "AUGUST": 8, "AGU": 8, "AUG": 8,
    "SEPTEMBER": 9, "SEP": 9,
    "OKTOBER": 10, "OCTOBER": 10, "OKT": 10, "OCT": 10,
    "NOVEMBER": 11, "NOV": 11,
    "DESEMBER": 12, "DECEMBER": 12, "DES": 12, "DEC": 12,
}


def _get_bordero_period(b: dict) -> tuple:
    """Return (year: int|None, month: int|None) from a Bordero row."""
    yr, mo = None, None
    raw_yr = b.get("TAHUN")
    if raw_yr is not None and not (isinstance(raw_yr, float) and math.isnan(raw_yr)):
        try:
            yr = int(float(raw_yr))
        except (ValueError, TypeError):
            yr = None
    raw_mo = str(b.get("BULAN", "")).strip().upper()
    if raw_mo:
        mo = _MONTH_NAME_TO_INT.get(raw_mo)
    return yr, mo


def _extract_suspend_slips(sus: dict) -> list:
    slips = []
    for c in ["SLIP NO", "SLIP_NO_CLEAN", "Certificate", "CERTIFICATE_1", "clean slip 1"]:
        v = sus.get(c)
        if v is not None and not (isinstance(v, float) and pd.isna(v)):
            s = _normalize(str(v))
            if s and s not in ("NAN", "NONE", "-"):
                slips.append(s)
    return list(dict.fromkeys(slips))


def _extract_bordero_slips(b: dict) -> list:
    slips = []
    for c in ["SLIP", "CERTIFICATE"]:
        v = b.get(c)
        if v is not None and not (isinstance(v, float) and pd.isna(v)):
            s = _normalize(str(v))
            if s and s not in ("NAN", "NONE", "-"):
                slips.append(s)
    return list(dict.fromkeys(slips))


def _bordero_matches_insured(b: dict, sus_ins_vals: list) -> bool:
    if not sus_ins_vals:
        return True
    b_ins = _normalize(b.get("INSURED_CLEAN") or b.get("INSURED", ""))
    if not b_ins or b_ins in ("NAN", "NONE", "-"):
        return True
    b_tokens = _tokenize(b_ins)
    for s_val in sus_ins_vals:
        s_norm = _normalize(s_val)
        if not s_norm:
            continue
        if s_norm in b_ins or b_ins in s_norm:
            return True
        s_tokens = _tokenize(s_norm)
        if s_tokens and b_tokens:
            common = s_tokens.intersection(b_tokens)
            sig_common = [t for t in common if t not in _INSURED_STOPWORDS and len(t) >= 3]
            if len(sig_common) >= 2 or (len(sig_common) == 1 and len(s_tokens) == 1):
                return True
    return False


def _get_suspend_amount(sus: dict) -> float:
    for c in ["AMOUNT ORI", "AMOUNT PAY"]:
        v = sus.get(c)
        if v is not None and not (isinstance(v, float) and pd.isna(v)):
            try:
                amt = abs(float(v))
                if amt > 0:
                    return amt
            except (ValueError, TypeError):
                pass
    return 0.0


def _get_bordero_net(b: dict) -> float:
    v = b.get("NET")
    if v is not None and not (isinstance(v, float) and pd.isna(v)):
        try:
            return abs(float(v))
        except (ValueError, TypeError):
            pass
    return 0.0


def _apply_bordero_override(
    sus: dict,
    bordero_core_idx: dict,
    bordero_polis_idx: dict,
    facode_osbal_idx: dict,
    osbal_rows: list,
    sus_curr: str,
    effective_date,
    is_lineslip: bool,
    insured_sus: list,
) -> tuple:
    """Teknik 2: Override Bordero.
    
    Menemukan kemungkinan match secara exact berdasarkan bordero dengan memulai
    kembali matching dari awal:
      polis -> periode -> curr -> slip -> insured -> net
      dari suspend -> ke bordero -> osbal -> suspend.
      
    Jika dapat exact match definitif, mengabaikan hasil matching sebelumnya.
    Returns: (osbal_indices: list, scenario: str) atau (None, None)
    """
    sus_p = sus.get("POLIS_CLEAN") or sus.get("clean polis 1") or sus.get("POLIS") or ""
    sus_core = _extract_core_number(sus_p) or _extract_core_number(sus.get("POLIS", ""))
    sus_p_norm = _normalize(sus_p)

    bor_hits = []
    if sus_core and sus_core in bordero_core_idx:
        bor_hits = bordero_core_idx[sus_core]
    elif sus_p_norm and sus_p_norm in bordero_polis_idx:
        bor_hits = bordero_polis_idx[sus_p_norm]

    if not bor_hits:
        return None, None, False, False

    candidates = list(bor_hits)

    # 1. Filter Periode (jika ada data tanggal di suspend dan bordero)
    sus_date = effective_date
    if pd.notna(sus_date):
        sus_yr = sus_date.year
        sus_mo = sus_date.month
        exact_ym = [b for b in candidates if _get_bordero_period(b) == (sus_yr, sus_mo)]
        if exact_ym:
            candidates = exact_ym
        else:
            exact_y = [b for b in candidates if _get_bordero_period(b)[0] == sus_yr]
            if exact_y:
                candidates = exact_y

    # 2. Filter Currency (wajib exact)
    if sus_curr:
        curr_hits = [b for b in candidates if _normalize(b.get("CURR", "")) == sus_curr]
        if curr_hits:
            candidates = curr_hits

    # 3. Filter Slip / Certificate (jika ada di kedua sisi)
    sus_slips = _extract_suspend_slips(sus)
    if sus_slips:
        slip_matched = [
            b for b in candidates
            if any(s in _extract_bordero_slips(b) for s in sus_slips)
        ]
        if slip_matched:
            candidates = slip_matched

    # 4. Filter Insured (jika ada nama insured)
    ins_vals = _collect_clean_values(sus, insured_sus)
    if ins_vals:
        ins_matched = [b for b in candidates if _bordero_matches_insured(b, ins_vals)]
        if ins_matched:
            candidates = ins_matched

    # 5. Filter Net (cuma berfungsi kalau nilainya exact, kalau engga exact abaikan)
    amt_sus = _get_suspend_amount(sus)
    if amt_sus > 0:
        net_exact_hits = [
            b for b in candidates
            if abs(_get_bordero_net(b) - amt_sus) < 0.05
        ]
        if net_exact_hits:
            candidates = net_exact_hits

    # Ambil FACCODE definitif dari hasil filter Bordero
    bor_facs = {
        _normalize(b.get("FACCODE", ""))
        for b in candidates
        if b.get("FACCODE") and _normalize(b.get("FACCODE", "")) not in ("", "NAN")
    }

    if len(bor_facs) == 1:
        single_fac = list(bor_facs)[0]
        if single_fac in facode_osbal_idx:
            osbal_target_indices = list(facode_osbal_idx[single_fac])
            is_beda_curr = False
            if sus_curr:
                curr_osbal = [
                    idx for idx in osbal_target_indices
                    if _normalize(osbal_rows[idx].get(OSBAL_CURR_COL, "")) == sus_curr
                ]
                if curr_osbal:
                    osbal_target_indices = curr_osbal
                else:
                    is_beda_curr = True

            is_beda_per = False
            if not is_beda_curr:
                if len(osbal_target_indices) > 1:
                    period_osbal = _narrow_by_periode(
                        sus, osbal_target_indices, osbal_rows,
                        effective_sus_date=effective_date, is_lineslip=is_lineslip
                    )
                    if period_osbal:
                        osbal_target_indices = period_osbal
                if pd.notna(effective_date):
                    has_valid_per = any(
                        osbal_rows[i].get("_com_date_parsed") is None or
                        (effective_date.year == osbal_rows[i]["_com_date_parsed"].year and effective_date.month == osbal_rows[i]["_com_date_parsed"].month if is_lineslip else effective_date >= osbal_rows[i]["_com_date_parsed"])
                        for i in osbal_target_indices
                    )
                    if not has_valid_per:
                        is_beda_per = True

            if osbal_target_indices:
                return osbal_target_indices, "Polis only", is_beda_curr, is_beda_per

    return None, None, False, False


def _apply_bordero_narrowing(
    matched_osbal_idx: set,
    sus: dict,
    bordero_core_idx: dict,
    bordero_polis_idx: dict,
    osbal_rows: list,
    sus_curr: str,
    effective_date,
    insured_sus: list,
) -> tuple:
    """Teknik 1: Narrowing Bordero.
    
    Menyempitkan kandidat yang memiliki > 1 fac code menggunakan data Bordero
    secara bertahap non-destruktif:
      polis -> periode -> curr -> slip -> insured -> net (exact only).
      
    Hanya mempersempit kemungkinan kandidat OSBAL yang ada.
    Returns: (narrowed_osbal_idx: set, was_narrowed: bool)
    """
    if len(matched_osbal_idx) <= 1:
        return matched_osbal_idx, False

    sus_p = sus.get("POLIS_CLEAN") or sus.get("clean polis 1") or sus.get("POLIS") or ""
    sus_core = _extract_core_number(sus_p) or _extract_core_number(sus.get("POLIS", ""))
    sus_p_norm = _normalize(sus_p)

    bor_hits = []
    if sus_core and sus_core in bordero_core_idx:
        bor_hits = bordero_core_idx[sus_core]
    elif sus_p_norm and sus_p_norm in bordero_polis_idx:
        bor_hits = bordero_polis_idx[sus_p_norm]

    if not bor_hits:
        return matched_osbal_idx, False

    candidates = list(bor_hits)

    # 1. Periode (non-destruktif)
    sus_date = effective_date
    if pd.notna(sus_date):
        sus_yr = sus_date.year
        sus_mo = sus_date.month
        exact_ym = [b for b in candidates if _get_bordero_period(b) == (sus_yr, sus_mo)]
        if exact_ym:
            candidates = exact_ym
        else:
            exact_y = [b for b in candidates if _get_bordero_period(b)[0] == sus_yr]
            if exact_y:
                candidates = exact_y

    # 2. Currency (non-destruktif)
    if sus_curr:
        curr_hits = [b for b in candidates if _normalize(b.get("CURR", "")) == sus_curr]
        if curr_hits:
            candidates = curr_hits

    # 3. Slip (non-destruktif)
    sus_slips = _extract_suspend_slips(sus)
    if sus_slips:
        slip_hits = [b for b in candidates if any(s in _extract_bordero_slips(b) for s in sus_slips)]
        if slip_hits:
            candidates = slip_hits

    # 4. Insured (non-destruktif)
    ins_vals = _collect_clean_values(sus, insured_sus)
    if ins_vals:
        ins_hits = [b for b in candidates if _bordero_matches_insured(b, ins_vals)]
        if ins_hits:
            candidates = ins_hits

    # 5. Net (cuma berfungsi kalau nilainya exact, kalau engga exact abaikan)
    amt_sus = _get_suspend_amount(sus)
    if amt_sus > 0:
        net_hits = [b for b in candidates if abs(_get_bordero_net(b) - amt_sus) < 0.05]
        if net_hits:
            candidates = net_hits

    bor_facs = {
        _normalize(b.get("FACCODE", ""))
        for b in candidates
        if b.get("FACCODE") and _normalize(b.get("FACCODE", "")) not in ("", "NAN")
    }

    if bor_facs:
        cand_in_bor = [
            idx for idx in matched_osbal_idx
            if _normalize(osbal_rows[idx].get(OSBAL_FACODE_COL, "")) in bor_facs
        ]
        if cand_in_bor and len(cand_in_bor) < len(matched_osbal_idx):
            return set(cand_in_bor), True

    return matched_osbal_idx, False


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

    # [TOKIO] Ekstraksi clean polis (hingga 5 kolom)
    p_parts = []
    for i in range(1, 6):
        v = suspend_row.get(f"clean polis {i}") or suspend_row.get(f"POLICY_CLEAN_{i}")
        if v is not None and not (isinstance(v, float) and pd.isna(v)):
            s = str(v).strip()
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
            if sub_clean and sub_clean.upper() not in ("NAN", "NONE", "-") and sub_clean not in flattened_p:
                flattened_p.append(sub_clean)
    p_parts = flattened_p

    # [TOKIO] Ekstraksi clean sertif (hingga 3 kolom)
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

    # [TOKIO] Ekstraksi clean insured (hingga 5 bagian atau split ;)
    raw_ins_cln = str(suspend_row.get("INSURED_CLEAN", "")).strip()
    ins_parts = []
    for c in ["clean insured 1", "clean insured 2", "clean insured 3", "clean insured 4", "clean insured 5"]:
        v = suspend_row.get(c, "")
        if v and str(v).strip() and str(v).strip().upper() not in ("NAN", "NONE"):
            ins_parts.append(str(v).strip())
    if not ins_parts and raw_ins_cln and raw_ins_cln.upper() not in ("NAN", "NONE"):
        ins_parts = [p.strip() for p in raw_ins_cln.split(";") if p.strip()]

    return {
        "CCOS_DOC_NO":       ccos["CCOS_DOC_NO"] if has_osbal else "",
        "CCOS_REF_CODE":     ccos["CCOS_REF_CODE"],
        "RECEIPT NO":        suspend_row.get("RECEIPT NO",        ""),
        "CREDIT NOTES":      suspend_row.get("CREDIT NOTES",      ""),
        "DETAIL RINCIAN NO": suspend_row.get("DETAIL RINCIAN NO", ""),
        "RECEIPT DATE":      suspend_row.get("RECEIPT DATE",      ""),
        "CEDANT NAME":       suspend_row.get("CEDANT NAME",       ""),
        "CEDANT SHRT NAME":  suspend_row.get("CEDANT SHRT NAME",  ""),

        # [TOKIO] Nama kolom insured dari file suspend Tokio
        "INSURED_ORI": suspend_row.get("insured_ori", "") or suspend_row.get("INSURED", ""),
        "INSURED_1":   ins_parts[0] if len(ins_parts) > 0 else (suspend_row.get("INSURED_CLEAN", "") or ""),
        "INSURED_2":   ins_parts[1] if len(ins_parts) > 1 else "",

        "CURR ORI":   suspend_row.get("CURR ORI",   ""),
        "AMOUNT ORI": suspend_row.get("AMOUNT ORI", ""),
        "CURR PAY":   suspend_row.get("CURR PAY",   ""),
        "AMOUNT PAY": suspend_row.get("AMOUNT PAY", ""),

        "CCOS_OR_BAL":  ccos["CCOS_OR_BAL"],
        "CCOS_BAL_DUE": ccos["CCOS_BAL_DUE"],

        # [TOKIO] polis: POLIS_ORI, 5 kolom POLICY_CLEAN, 3 kolom SERTIF_CLEAN
        "POLIS_ORI":       suspend_row.get("polis_ori", "") or suspend_row.get("POLIS", ""),
        "POLICY_CLEAN_1":  p_parts[0] if len(p_parts) > 0 else "",
        "SERTIF_CLEAN_1":  s_parts[0] if len(s_parts) > 0 else "",
        "POLICY_CLEAN_2":  p_parts[1] if len(p_parts) > 1 else "",
        "SERTIF_CLEAN_2":  s_parts[1] if len(s_parts) > 1 else "",
        "POLICY_CLEAN_3":  p_parts[2] if len(p_parts) > 2 else "",
        "SERTIF_CLEAN_3":  s_parts[2] if len(s_parts) > 2 else "",
        "POLICY_CLEAN_4":  p_parts[3] if len(p_parts) > 3 else "",
        "POLICY_CLEAN_5":  p_parts[4] if len(p_parts) > 4 else "",
        "SLIP_NO_ORI":     suspend_row.get("slip_ori", "") or suspend_row.get("SLIP NO", ""),
        "SLIP_NO_CLN":     suspend_row.get("clean slip 1", "") or suspend_row.get("SLIP_NO_CLEAN", ""),

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

    [IBS DELTA] Matriks np.select identik dengan ACA.
    Kondisi is_gt1_ocmc (Bordero OC Override) dipertahankan struktural
    tapi secara alami selalu False karena IBS tidak punya bordero.
    Ref: S4 matching_IBS.md
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
    is_gt1_ocmc          = skenario_col.str.contains("Bordero", na=False, regex=False) & (skenario_col.str.contains("> 1 Fac", na=False, regex=False) | ccos_ref.str.contains(",", na=False))
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

def _filter_cedant(rows: list, header: list, label: str, required: bool = False) -> list:
    """Sisakan hanya baris cedant Tokio (cek semua kolom bernama '*CEDANT*').

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
    """Pipeline utama TOKIO V1: matching berbasis POLIS lalu INSURED (tanpa slip, tanpa Bordero).

    Urutan fase: Polis -> Insured -> Currency & Periode -> Narrowing Non-Destruktif
    Output hanya berisi cedant Tokio.
    """
    t_start = time.perf_counter()
    print("\n" + "=" * 60, flush=True)
    print("  PRODUCTION SCRIPT -- SUSPEND MATCHING  (TOKIO) -- V1", flush=True)
    print("=" * 60, flush=True)

    # [1] Load data
    print("\n[1/6] Loading data ...", flush=True)
    suspend_rows, suspend_cols = _load_excel(SUSPEND_FILE, "SUSPEND")
    osbal_rows,   osbal_cols   = _load_excel(OSBAL_FILE,   "OSBAL")
    facul_rows,   facul_cols   = _load_excel(FACUL_FILE,   "FACUL")
    slipdb_rows,  slipdb_cols  = _load_excel(SLIPDB_FILE,  "SLIPDB", optional=True)
    bordero_rows, bordero_cols = _load_excel(BORDERO_FILE, "BORDERO", optional=True)

    if not suspend_rows:
        print("\nERROR: data SUSPEND kosong / file tidak ditemukan. Cek SUSPEND_FILE.", flush=True)
        sys.exit(1)

    # [TOKIO] Hanya cedant Tokio: suspend WAJIB difilter, referensi difilter bila memungkinkan
    print("\n  --- Filter cedant Tokio ---", flush=True)
    suspend_rows = _filter_cedant(suspend_rows, suspend_cols, "SUSPEND", required=True)

    # [TOKIO] Baris SUSPEND yang mengandung lineslip / penyelesaian suspense / hutang piutang
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
        print("\nERROR: tidak ada baris SUSPEND Tokio yang bisa di-match.", flush=True)
        sys.exit(1)

    if FILTER_REF_BY_CEDANT:
        osbal_rows  = _filter_cedant(osbal_rows,  osbal_cols,  "OSBAL")
        facul_rows  = _filter_cedant(facul_rows,  facul_cols,  "FACUL")
        slipdb_rows = _filter_cedant(slipdb_rows, slipdb_cols, "SLIPDB")
    # Bordero tidak dimuat -- IBS tidak punya Open Cover Marine Cargo.

    # Diagnostik kolom kunci
    print("\n  --- Cek kolom SUSPEND ---", flush=True)
    print(f"  Semua kolom SUSPEND: {suspend_cols}", flush=True)
    for col in ["INSURED", "INSURED_CLEAN", "POLIS", "POLIS_CLEAN", "CERTIFICATE_1",
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

    # [TOKIO] Slip tidak dipakai sebagai kunci matching (rules Tokio: polis & insured).
    desc3_slip_vals = []

    # Kolom polis/slip standar suspend (untuk FASE 2 via _get_effective_sus_cols)
    polis_sus = [c for c in ["POLIS_CLEAN", "POLIS"] if c in suspend_cols] or _get_clean_cols(suspend_cols, "POLIS") or []
    slip_sus  = []   # [TOKIO] slip tidak dipakai (kolom slip suspend tetap ditampilkan di output)

    # [IBS DELTA] Insured
    insured_sus = [c for c in ["INSURED_CLEAN", "INSURED"] if c in suspend_cols] or _get_clean_cols(suspend_cols, "INSURED") or []

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

    # [TOKIO] Bangun index Bordero Open Cover Marine Cargo
    if bordero_rows:
        t_b0 = time.perf_counter()
        bordero_core_idx, bordero_polis_idx = _build_bordero_lookup(bordero_rows)
        print(f"  Bordero done ({time.perf_counter() - t_b0:.1f}s) - {len(bordero_core_idx):,} unique core M, {len(bordero_polis_idx):,} unique polis", flush=True)
    else:
        bordero_core_idx, bordero_polis_idx = {}, {}
        print("  Bordero: TIDAK DITEMUKAN", flush=True)

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

        eff_polis, eff_slip, like_polis, like_slip = _get_effective_sus_cols(sus, polis_sus, slip_sus)
        _is_ls = sus.get("_is_lineslip", False)
        effective_date = sus.get("_sus_date_lineslip") if _is_ls else sus.get("_sus_date_parsed")
        sus_cob = _extract_cob_from_suspend(sus)

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

        found_by_slip_desc3 = False   # [TOKIO] tidak ada fase slip

        # =================================================================
        # FASE 1 (TOKIO): Pencarian berdasarkan POLIS
        # Logika identik dengan FASE 1 ACA: cross-check OSBAL -> SLIPDB -> FACUL
        # =================================================================
        found_by_polis = False
        sus_p_val = sus.get("POLIS_CLEAN") or sus.get("POLIS") or ""
        sus_yr = _get_polis_year(sus_p_val)

        if not found_by_slip_desc3:
            raw_clean_vals = _collect_clean_values(sus, eff_polis)
            # Hierarki 1: Coba dulu pencocokan variasi nomor polis dengan tahun (mencegah bentrok antar-tahun)
            if raw_clean_vals:
                year_vars = []
                for rv in raw_clean_vals:
                    year_vars.extend(_get_polis_variations(rv))
                year_vars = list(dict.fromkeys(year_vars))

                osb_hits = _exact_match(year_vars, lookup_osbal[1])
                if osb_hits:
                    if sus_yr:
                        osb_hits = [h for h in osb_hits if not _get_row_year(osbal_rows[h]) or _get_row_year(osbal_rows[h]) == sus_yr]
                    if sus_cob:
                        cob_hits = [h for h in osb_hits if _osbal_matches_cob(osbal_rows[h], sus_cob)]
                        if cob_hits:
                            matched_osbal_idx.update(cob_hits)
                            found_by_polis = True
                    elif osb_hits:
                        matched_osbal_idx.update(osb_hits)
                        found_by_polis = True

            # Hierarki 2: Fallback ke ekspansi core number & pecahan (+) bila polis dengan tahun tidak ketemu
            if not found_by_polis and raw_clean_vals:
                polis_vals = _augment_polis_with_core(raw_clean_vals)
                if polis_vals:
                    osb_hits = _exact_match(polis_vals, lookup_osbal[1])
                    if osb_hits:
                        if sus_yr:
                            osb_hits = [h for h in osb_hits if not _get_row_year(osbal_rows[h]) or _get_row_year(osbal_rows[h]) == sus_yr]
                        if sus_cob:
                            cob_hits = [h for h in osb_hits if _osbal_matches_cob(osbal_rows[h], sus_cob)]
                            if cob_hits:
                                matched_osbal_idx.update(cob_hits)
                                found_by_polis = True
                        elif osb_hits:
                            matched_osbal_idx.update(osb_hits)
                            found_by_polis = True

                    if slipdb_rows and not found_by_polis:
                        sld_hits = _exact_match(polis_vals, lookup_slipdb[1])
                        if sld_hits:
                            res_idx, _ = _resolve_facode([slipdb_rows[idx] for idx in sld_hits], SLIPDB_FACODE_COL, facode_osbal_idx, osbal_rows)
                            if res_idx:
                                if sus_yr:
                                    res_idx = [h for h in res_idx if not _get_row_year(osbal_rows[h]) or _get_row_year(osbal_rows[h]) == sus_yr]
                                if sus_cob:
                                    cob_res = [h for h in res_idx if _osbal_matches_cob(osbal_rows[h], sus_cob)]
                                    if cob_res:
                                        matched_osbal_idx.update(cob_res)
                                        found_by_polis = True
                                elif res_idx:
                                    matched_osbal_idx.update(res_idx)
                                    found_by_polis = True

                    if facul_rows and not found_by_polis:
                        fac_hits = _exact_match(polis_vals, lookup_facul[1])
                        if fac_hits:
                            res_idx, _ = _resolve_facode([facul_rows[idx] for idx in fac_hits], FACUL_FACODE_COL, facode_osbal_idx, osbal_rows)
                            if res_idx:
                                if sus_yr:
                                    res_idx = [h for h in res_idx if not _get_row_year(osbal_rows[h]) or _get_row_year(osbal_rows[h]) == sus_yr]
                                if sus_cob:
                                    cob_res = [h for h in res_idx if _osbal_matches_cob(osbal_rows[h], sus_cob)]
                                    if cob_res:
                                        matched_osbal_idx.update(cob_res)
                                        found_by_polis = True
                                elif res_idx:
                                    matched_osbal_idx.update(res_idx)
                                    found_by_polis = True

            if found_by_polis:
                source = "OSBAL"
                scenario = "Polis only"

        # =================================================================
        # FASE 2 (TOKIO): Pencarian berdasarkan INSURED (fallback jika POLIS miss)
        # =================================================================
        if not matched_osbal_idx:
            ins_vals = _collect_clean_values(sus, insured_sus)
            if ins_vals:
                def _filter_valid_insured_hits(hit_indices):
                    if not sus_yr or len(hit_indices) <= 1:
                        return hit_indices
                    filtered = [
                        h for h in hit_indices
                        if not (_get_row_year(osbal_rows[h]) and _get_row_year(osbal_rows[h]) != sus_yr and float(osbal_rows[h].get("CCOS_BAL_DUE") or 0) <= 0.05)
                    ]
                    return filtered if filtered else hit_indices

                # 1. Coba OSBAL Exact & Like
                osb_hits = _exact_match(ins_vals, lookup_osbal[2])
                if not osb_hits and lookup_osbal[8] and lookup_osbal[8][0]:
                    osb_hits = _like_match(
                        ins_vals, lookup_osbal[8][0], osbal_rows, insured_osbal,
                        stopwords=_INSURED_STOPWORDS,
                        strict=True,
                        overlap_threshold=_INSURED_LIKE_THRESHOLD,
                        match_log=_like_log,
                    )
                if osb_hits:
                    osb_hits = _filter_valid_insured_hits(osb_hits)
                    if sus_cob:
                        cob_hits = [h for h in osb_hits if _osbal_matches_cob(osbal_rows[h], sus_cob)]
                        if cob_hits:
                            matched_osbal_idx.update(cob_hits)
                            source = "OSBAL"
                            scenario = "Insured only"
                    elif osb_hits:
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
                            strict=True,
                            overlap_threshold=_INSURED_LIKE_THRESHOLD,
                            match_log=_like_log,
                        )
                    if sld_hits:
                        res_idx, _ = _resolve_facode([slipdb_rows[idx] for idx in sld_hits], SLIPDB_FACODE_COL, facode_osbal_idx, osbal_rows)
                        if res_idx:
                            res_idx = _filter_valid_insured_hits(res_idx)
                            if sus_cob:
                                cob_res = [h for h in res_idx if _osbal_matches_cob(osbal_rows[h], sus_cob)]
                                if cob_res:
                                    matched_osbal_idx.update(cob_res)
                                    source = "SLIPDB"
                                    scenario = "Insured only"
                                    resolved = True
                            elif res_idx:
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
                            strict=True,
                            overlap_threshold=_INSURED_LIKE_THRESHOLD,
                            match_log=_like_log,
                        )
                    if fac_hits:
                        res_idx, _ = _resolve_facode([facul_rows[idx] for idx in fac_hits], FACUL_FACODE_COL, facode_osbal_idx, osbal_rows)
                        if res_idx:
                            res_idx = _filter_valid_insured_hits(res_idx)
                            if sus_cob:
                                cob_res = [h for h in res_idx if _osbal_matches_cob(osbal_rows[h], sus_cob)]
                                if cob_res:
                                    matched_osbal_idx.update(cob_res)
                                    source = "FACUL"
                                    scenario = "Insured only"
                                    resolved = True
                            elif res_idx:
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
                    sus_yr=sus_yr,
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
                        #   "+ Insured" tidak menambah informasi → seharusnya "Slip only".
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
        # FASE 6: Bordero Open Cover (Override & Narrowing)
        # Sesuai 2 Teknik Open Cover:
        # 1. Override: Suspend -> Bordero (exact 6 param: polis, periode, curr, slip, insured, net)
        #              -> FAC CODE -> OSBAL -> Override match sebelumnya.
        # 2. Narrowing: Jika >1 fac code -> sempitkan bertahap non-destruktif
        #               (polis -> periode -> curr -> slip -> insured -> net exact only).
        # =================================================================
        is_bordero = False
        if bordero_rows:
            sus_curr = _normalize(sus.get(SUSPEND_CURR_COL, ""))

            # TEKNIK 2: OVERRIDE (Mulai dari awal Suspend -> Bordero -> OSBAL)
            ov_indices, ov_scen, ov_bc, ov_bp = _apply_bordero_override(
                sus, bordero_core_idx, bordero_polis_idx, facode_osbal_idx,
                osbal_rows, sus_curr, effective_date, _is_ls, insured_sus
            )
            if ov_indices:
                matched_osbal_idx = set(ov_indices)
                source = "OSBAL"
                scenario = ov_scen or "Polis only"
                is_bordero = True
                is_beda_curr = ov_bc
                is_beda_per = ov_bp
                resolved = True
            else:
                # TEKNIK 1: NARROWING (>1 Fac Code -> sempitkan kemungkinan)
                current_facs = {
                    _normalize(osbal_rows[i].get(OSBAL_FACODE_COL, ""))
                    for i in matched_osbal_idx
                    if osbal_rows[i].get(OSBAL_FACODE_COL)
                }
                if len(current_facs) > 1:
                    narrowed_idx, was_narrowed = _apply_bordero_narrowing(
                        matched_osbal_idx, sus, bordero_core_idx, bordero_polis_idx,
                        osbal_rows, sus_curr, effective_date, insured_sus
                    )
                    if was_narrowed:
                        matched_osbal_idx = narrowed_idx
                        is_bordero = True

        # Post-matching normalisasi label
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

        fac_codes = {_normalize(r.get(OSBAL_FACODE_COL, "")) for r in osbal_ref if r}
        fac_codes.discard("")

        raw_results.append({
            "suspend": sus, "source": source,
            "base_scenario": scenario,
            "is_bordero": is_bordero,
            "is_beda_curr": is_beda_curr,
            "is_beda_per": is_beda_per,
            "osbal_idx": list(matched_osbal_idx) if not (is_beda_curr or is_beda_per) else (list(matched_list) if 'matched_list' in locals() and matched_list else list(matched_osbal_idx)),
            "osbal_count": osbal_count, "resolved": resolved, "fac_codes": fac_codes,
            "_like_log": [m for _, m in _like_log],
            "_orig_idx": n,
        })

    print(f"  Matching selesai dalam {time.perf_counter()-t4:.1f}s", flush=True)
    # Akumulasi counter method dari semua baris
    for _row in raw_results:
        for _method in _row.get("_like_log", []):
            if _method == "SUBSTR":  stat_like_substr  += 1
            elif _method == "SUBSET":  stat_like_subset  += 1
            elif _method == "OVERLAP": stat_like_overlap += 1

    # [4b] Fac code claim & osbal accumulation (identik dengan ACA/MAG)
    print("\n[4c/6] Fac code claim & osbal accumulation (row asli dipertahankan) ...", flush=True)

    claimed: set = set()
    for r in raw_results:
        if len(r["fac_codes"]) == 1 and r["source"] is not None and not r.get("is_beda_curr", False) and not r.get("is_beda_per", False):
            claimed |= r["fac_codes"]

    for r in raw_results:
        if len(r["fac_codes"]) > 1:
            remaining = r["fac_codes"] - claimed
            if len(remaining) == 1 and r["source"] is not None and not r.get("is_beda_curr", False) and not r.get("is_beda_per", False):
                fac = next(iter(remaining))
                r["fac_codes"] = remaining
                r["osbal_idx"] = [
                    i for i in r["osbal_idx"]
                    if _normalize(osbal_rows[i].get(OSBAL_FACODE_COL, "")) == fac
                ]
                r["osbal_count"] = len(r["osbal_idx"])
                claimed |= remaining

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

    # Urutkan kembali sesuai urutan baris suspend asli (identik dengan standar MAG & ACA)
    final.sort(key=lambda r: r.get("_orig_idx", 0))

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
            is_beda_curr=r.get("is_beda_curr", False),
            is_beda_per=r.get("is_beda_per", False),
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
    for c in df.select_dtypes(include=['object']).columns:
        df[c] = df[c].astype(str).str.replace(r'[\x00-\x08\x0B\x0C\x0E-\x1F]', '', regex=True)
    
    save_path = OUTPUT_FILE
    try:
        df.to_excel(save_path, index=False)
    except PermissionError:
        save_path = os.path.join(_BASE_DIR, "matching_tokio_new.xlsx")
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
        df.to_sql("SUSPENSE_DATA_SUSPENSE_TOKIO_V1", con=engine, if_exists="append", index=False)
        print(f"      -> {len(df):,} rows berhasil di-export ke 'SUSPENSE_DATA_SUSPENSE_TOKIO_V1'")
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