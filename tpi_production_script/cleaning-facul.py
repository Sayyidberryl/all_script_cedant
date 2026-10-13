"""
Cleaning Data 1 - Facultative | Cedant: TPI (PT Asuransi Tugu Pratama Indonesia)

UPDATE BARU -- KOLOM CERTIFICATE (rule sama persis dgn mesin JASA R.P, lihat
build_certificate()):
- SUMBER: diambil dari FAC_POLICY_NO (nilai mentah, SEBELUM clean_polis
  menjalankan mesin repetition-chain-nya), diposisikan tepat setelah
  FAC_POLICY_NO / sebelum FAC_POLICY_NO_1.
- Pola yang dikonfirmasi user dari eksplorasi TPI Data 1:
    * "1911000022019000006 - 001364"   -> CERTIFICATE = "001364"
    * "04240000000099-000001 / 2024"   -> CERTIFICATE = "000001"
  Jadi pola umum: BASE(10+ digit) [spasi opsional] "-" [spasi opsional]
  CERT(3-7 digit) [opsional "/ TAHUN"], lalu akhir string.
- PENTING (harus diketahui): clean_polis()/mesin repetition-chain TIDAK
  diubah sama sekali oleh update ini -- artinya untuk pola BASE-CERT di
  atas, FAC_POLICY_NO_1 dst TETAP mengikuti perilaku lama (merekonstruksi
  suffix pendek sebagai pengganti digit akhir base, BUKAN membuang cert-nya
  seperti di Data 1/3 JRP). CERTIFICATE murni kolom tambahan yang diambil
  independen dari nilai mentah, tidak memengaruhi FAC_POLICY_NO_1 dst.
  Kalau ternyata FAC_POLICY_NO_1 juga perlu diperbaiki supaya konsisten
  (mis. cert dibuang bersih dari base), itu perubahan terpisah -- beri tahu
  supaya bisa disesuaikan.
- Rule format build_certificate() SAMA seperti mesin JRP:
    * range 2-angka ("-" / S/D / SD / S): >3 anggota -> "AWAL SD AKHIR",
      <=3 anggota -> breakdown penuh dipisah koma.
    * daftar eksplisit koma/campur S-D: >3 item -> kompres
      "ITEM_PERTAMA SD ITEM_TERAKHIR" (tanpa isi celah), <=3 -> apa adanya.
    * semua angka di-pad ke 6 digit.
    * "S/D"/"s/d"/"SD"/"S" di output SELALU distandarisasi jadi "SD".
    * pola tidak masuk akal (bukan angka murni) -> CERTIFICATE kosong.
  Catatan: baru 2 contoh yang terverifikasi user (single value, bukan
  range/list) -- cabang range/list di build_certificate() belum tervalidasi
  ke data TPI asli, ini perlu verifikasi lanjut kalau ditemukan kasusnya.
"""

import os
import re

import pandas as pd


# ─────────────────────────────────────────────────────────────────────────────
# KONFIGURASI — ubah bagian ini kalau nama file/kolom berbeda
# ─────────────────────────────────────────────────────────────────────────────

INPUT_FILE  = os.path.join("input", "1b. Transaksi Facul 01.01.23 - 17.08.2026.xlsx")
OUTPUT_FILE = os.path.join("output", "tpi_output_facul.xlsx")

CEDANT_COL   = "COMP_NAME"
CEDANT_VALUE = "PT ASURANSI TUGU PRATAMA INDONESIA"

BROKER_COL  = "COMP_NAME.1"
POLIS_COL   = "FAC_POLICY_NO"
SLIP_COL    = "FAC_SLIP"
INSURED_COL = "FAC_INSURED"

MAX_SPLIT_COLS = 5   # batas breakdown maksimal, sisanya digabung koma

# ─────────────────────────────────────────────────────────────────────────────
# REGEX & KAMUS PEMBANTU (diadaptasi dari referensi)
# ─────────────────────────────────────────────────────────────────────────────

POLIS_EXCEPTION_RE = re.compile(
    r"""
    MOP\s*MARINE
  | (?:LINE\s*SLIP|LINESLIP)
  | \bP[1-5]\s+CANCEL
  | \bP[1-5]\b
  | \bCANCEL\b
  | PENYELESAIAN
  | HUTANG\s*PIUTANG
    """,
    re.IGNORECASE | re.VERBOSE,
)

SLIP_EXCEPTION_RE = re.compile(
    r"""
    \bSUMMARY\b
  | \bBORDER[OA]\b
  | \bBORDRO\b
  | \bSINGGLESHIPMENT\b
    """,
    re.IGNORECASE | re.VERBOSE,
)

INSURED_JUNK_WORDS = frozenset({
    "PT", "CV", "TBK", "PERSERO", "LTD", "INC", "LLC",
    "AND", "OR", "THE", "OF", "AS",
    "NON FOOD", "DIV",
})

INSURED_SUFFIX_RE = re.compile(
    r"""
    ,?\s*\bTBK\b\s*(?:,?\s*PT\.?)?
  | ,?\s*\(\s*PERSERO\s*\)\s*
  | ,?\s*\bPERSERO\b\s*
  | ,?\s*\bLTD\.?\b\s*
  | ,?\s*\bPT\.?\s*$
  | ,?\s*\bCV\.?\s*$
    """,
    re.IGNORECASE | re.VERBOSE,
)

_SLIP_NOISE_RE = re.compile(
    r"""
    \b(?:JANUARI|FEBRUARI|MARET|APRIL|MEI|JUNI|JULI|AGUSTUS
        |SEPTEMBER|OKTOBER|NOVEMBER|DESEMBER)\b
  | \b(?:JANUARY|FEBRUARY|MARCH|APRIL|MAY|JUNE|JULY|AUGUST
        |SEPTEMBER|OCTOBER|NOVEMBER|DECEMBER)\b
  | \b(?:IDR|USD|ENG)\b
  | \bP[1-5]\b
  | \bNEW\b
  | \bVARIOUS\b
  | \b20[0-9]{2}\b
    """,
    re.IGNORECASE | re.VERBOSE,
)

_SLIP_TOKEN_RE = re.compile(r"[A-Z0-9][A-Z0-9\-]{6,}", re.IGNORECASE)


# ─────────────────────────────────────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────────────────────────────────────

def _normalize_spaces(text: str) -> str:
    return re.sub(r"\s{2,}", " ", text).strip()


def _is_valid_polis_token(tok: str) -> bool:
    tok = tok.strip()
    return len(tok) >= 5 and not re.search(r"TBA$", tok, re.IGNORECASE) and bool(re.search(r"\d", tok))


def _extract_polis_tokens(text: str) -> list:
    tokens = []
    for block in re.split(r"\s{2,}", text.strip()):
        for tok in re.split(r"\+|\s+", block.strip()):
            tok = tok.strip().strip("-/")
            if _is_valid_polis_token(tok):
                tokens.append(tok)
    return tokens


def _expand_dash_chain(segment: str) -> list:
    results = []
    current_base = None
    for part in segment.split("-"):
        part = part.strip()
        if not part:
            continue
        if len(part) >= 8 and re.search(r"\d", part):
            current_base = part
            results.append(part)
        elif current_base and len(part) >= 2 and re.match(r"^\d+$", part) and len(part) < len(current_base):
            n = len(part)
            results.append(current_base[:-n] + part)
        else:
            if _is_valid_polis_token(part):
                results.append(part)
    return results or ([segment] if _is_valid_polis_token(segment) else [])


def _is_valid_slip_token(tok: str) -> bool:
    tok = tok.strip()
    return len(tok) >= 7 and bool(re.search(r"\d", tok)) and not re.match(r"^\d{4}$", tok)


def _strip_slip_suffix(tok: str) -> str:
    m = re.match(r"^(.+?)-(\d{1,6})$", tok)
    if m and len(m.group(1)) > len(m.group(2)):
        return m.group(1)
    return tok


def _extract_slip_tokens(text: str) -> list:
    results = []
    for block in re.split(r"\s{2,}", text.strip()):
        clean = _SLIP_NOISE_RE.sub(" ", block)
        clean = re.sub(r"^[\s\-/+,]+|[\s\-/+,]+$", "", clean).strip()
        for cand in _SLIP_TOKEN_RE.findall(clean):
            cand = cand.strip("-")
            parts = cand.split("-")
            if len(parts) == 2:
                a, b = parts
                if _is_valid_slip_token(a) and _is_valid_slip_token(b) and abs(len(a) - len(b)) <= 2:
                    results.extend([a, b])
                    continue
            if _is_valid_slip_token(cand):
                results.append(_strip_slip_suffix(cand))
    return results


_INSURED_PAREN_JUNK = {"PERSERO", "TBK", "SMALL SHIP FLEET"}


def _handle_insured_paren(match: "re.Match") -> str:
    content_ = match.group(1).strip()
    if content_.upper() in _INSURED_PAREN_JUNK or not content_:
        return ""
    kept = content_.split(",")[0].strip()
    return " " + kept + " " if kept else ""

_REGION_PREFIX_WORDS = {"KALTIM", "KALSEL", "KALTENG", "KALBAR", "KALUT",
                         "SULSEL", "SULUT", "SULTENG", "SULBAR", "SULTRA",
                         "JATIM", "JATENG", "JABAR", "SUMUT", "SUMSEL",
                         "SUMBAR", "NTB", "NTT", "DKI"}

def _merge_region_prefix_parts(parts: list) -> list:
    merged = []
    skip_next_prefix = False
    for i, p in enumerate(parts):
        p_stripped = p.strip()
        if p_stripped.upper() in _REGION_PREFIX_WORDS and i + 1 < len(parts):
            merged.append(p_stripped + " " + parts[i + 1].strip())
            skip_next_prefix = True
        elif skip_next_prefix:
            skip_next_prefix = False
            continue
        else:
            merged.append(p_stripped)
    return merged

def _merge_digit_start_parts(parts: list) -> list:
    merged = []
    for p in parts:
        p_stripped = p.strip()
        if merged and re.match(r"^\d", p_stripped):
            merged[-1] = merged[-1] + " " + p_stripped
        else:
            merged.append(p_stripped)
    return merged


def _name_initials_variants(name: str) -> set:
    words = [w for w in re.split(r"[\s\-]+", name.upper()) if w]
    if not words:
        return set()
    variants = {"".join(w[0] for w in words if w[0].isalpha())}
    if len(words) >= 2:
        variants.add("".join(w[0] for w in words[:-1] if w[0].isalpha()) + words[-1])
    return variants


def _is_abbreviation_of(short: str, long_name: str) -> bool:
    short_clean = re.sub(r"[^A-Z0-9]", "", short.upper())
    if len(short_clean) < 2 or len(short_clean) > 8:
        return False
    return short_clean in _name_initials_variants(long_name)


def _strip_trailing_abbrev_word(name: str, priors: list) -> str:
    words = name.split()
    if len(words) >= 2 and any(_is_abbreviation_of(words[-1], p) for p in priors):
        return " ".join(words[:-1])
    return name


def _clean_insured_name(name: str) -> str:
    name = _normalize_spaces(name)
    name = _INSURED_PREFIX_RE.sub("", name)
    name = _INSURED_GELAR_RE.sub("", name)
    name = re.sub(r"\bPTE\.?\b", "", name, flags=re.IGNORECASE)
    name = re.sub(r"\(([^()]*)\)", _handle_insured_paren, name)
    name = name.replace("(", " ").replace(")", " ")
    name = re.sub(r"/", " ", name)
    for _ in range(3):
        cleaned = _normalize_spaces(INSURED_SUFFIX_RE.sub("", name).strip().strip(","))
        if cleaned == name:
            break
        name = cleaned
    name = name.replace(".", "")
    name = _normalize_spaces(name).strip()
    name = name.upper()
    return name


def _cap_or_join(items: list) -> list:
    if len(items) > MAX_SPLIT_COLS:
        return [",".join(items)]
    return items


# ─────────────────────────────────────────────────────────────────────────────
# KAMUS POLA STANDAR TPI
# ─────────────────────────────────────────────────────────────────────────────

KNOWN_POLIS_PATTERNS = [
    re.compile(r"^\d{14}\s*/\s*\d{4}$"),          # 03250000007295 / 2025
    re.compile(r"^\d{14}$"),                       # 03250000009340
    re.compile(r"^[A-Za-z]{3}\d{7}$"),             # PVE1900115
    re.compile(r"^[A-Za-z]{3}\d{7}\s*/\s*\d{4}$"), # PVV2000087 / 2020
    re.compile(r"^\d{19,20}$"),                    # 1913102022120000372 (Heavy Equipment)
    re.compile(r"^\d{14}-\d{6}\s*/\s*\d{4}$"),     # 04250000000099-000174 / 2026 (Marine Cargo)
    re.compile(r"^\d{6}$"),                        # 140057
]

KNOWN_SLIP_PATTERNS = [
    re.compile(r"^\d{10}$"),                       # 5100486892
    re.compile(r"^[A-Za-z]{3}\d{7}$"),              # FVV2000148
    re.compile(r"^\d{17,19}$"),                     # 29133060324000000 / 891310202212000008
    re.compile(r"^\d{10}\s*/\s*PAYMENT NUMBER\s*\d{10}$", re.IGNORECASE),
    re.compile(r"^TIDAK ADA$", re.IGNORECASE),
]

_COMPOSITE_SLIP_RE = re.compile(
    r"(?:IDR|USD|EUR|SGD|JPY|GBP|AUD|CNY|MYR|CHF|THB)[\s\-]*"
    r"(?:[A-Za-z][A-Za-z0-9]{1,3}|[A-Za-z0-9]{1,3}[A-Za-z])[\s\-]*"
    r"(\d{10})(?!\d)",
    re.IGNORECASE,
)

_COMPOSITE_POLIS_RE = re.compile(
    r"^([A-Za-z]{3}\d{7}|\d{14})[\s\-]*(?:19|20)\d{2}.*(?:IDR|USD|EUR|SGD|JPY|GBP|AUD|CNY|MYR|CHF|THB)",
    re.IGNORECASE,
)


def _matches_known_pattern(value: str, patterns) -> bool:
    return any(p.match(value) for p in patterns)


def refine_with_known_pattern(value: str, patterns) -> str:
    if not value:
        return value
    if _matches_known_pattern(value, patterns):
        return value

    candidates = [
        re.sub(r"\s+", "", value),
        re.sub(r"\s*/\s*", " / ", value).strip(),
        value.strip(" .-/"),
        re.sub(r"\.0+$", "", value),
        "0" + value if value.isdigit() else value,
        re.sub(r"^(\d+)(\s*/\s*(?:19|20)\d{2})$", r"0\1\2", value),
    ]
    for c in candidates:
        if _matches_known_pattern(c, patterns):
            return c

    return value


def _refine_list(items: list, patterns) -> list:
    return [refine_with_known_pattern(v, patterns) for v in items]


# ─────────────────────────────────────────────────────────────────────────────
# KOLOM CERTIFICATE (BARU -- rule sama persis dengan mesin JASA R.P)
# ─────────────────────────────────────────────────────────────────────────────

_CERT_RANGE_PAIR_RE = re.compile(
    r"^\s*(\d{1,7})\s*(?:-|S\s*/\s*D|SD|S(?!\d))\s*(\d{1,7})\s*$",
    re.IGNORECASE,
)
_CERT_SPLIT_RE = re.compile(r"\s*(?:S\s*/\s*D|SD|S(?!\d)|,)\s*", re.IGNORECASE)
_CERT_TOKEN_RE = re.compile(r"^\d{1,7}$")


def build_certificate(cert_raw) -> str:
    """Format nomor sertifikat sesuai rule baru (sama dipakai di mesin JRP):
    - range 2-angka ("-" / S/D / SD / S): >3 anggota -> "AWAL SD AKHIR",
      <=3 anggota -> breakdown penuh dipisah koma.
    - daftar eksplisit (koma dan/atau campur S/D/SD/S, bukan range murni):
      >3 item -> dikompres "ITEM_PERTAMA SD ITEM_TERAKHIR" (tanpa isi celah),
      <=3 item -> tetap apa adanya, dipisah koma.
    - semua angka di-pad ke 6 digit.
    - "S/D"/"s/d"/"SD"/"S" di output SELALU ditulis "SD".
    - kalau pola tidak masuk akal (bukan angka murni) -> kosong.
    """
    if cert_raw is None:
        return ""
    text = str(cert_raw).strip()
    if not text:
        return ""

    m = _CERT_RANGE_PAIR_RE.match(text)
    if m:
        lo, hi = int(m.group(1)), int(m.group(2))
        if hi < lo:
            lo, hi = hi, lo
        count = hi - lo + 1
        if count > 3:
            return f"{str(lo).zfill(6)} SD {str(hi).zfill(6)}"
        return ", ".join(str(n).zfill(6) for n in range(lo, hi + 1))

    raw_tokens = [t for t in _CERT_SPLIT_RE.split(text) if t.strip()]
    if not raw_tokens or not all(_CERT_TOKEN_RE.match(t.strip()) for t in raw_tokens):
        return ""
    nums = [int(t) for t in raw_tokens]
    if len(nums) > 3:
        return f"{str(nums[0]).zfill(6)} SD {str(nums[-1]).zfill(6)}"
    return ", ".join(str(n).zfill(6) for n in nums)


# BUG TPI NEW -- kriteria sertifikat DIPERSEMPIT. Dari feedback user,
# "BASE-angka_pendek" TANPA spasi di sekitar dash DAN TANPA "/ TAHUN"
# menempel (mis. "11401101900111-201", "1911000022019000006-002946")
# TERNYATA BUKAN sertifikat -- itu pola breakdown perulangan biasa (base
# yang N digit terakhirnya diganti). Sertifikat ASLI selalu punya salah
# satu dari dua ciri berikut:
#   Pola A: BASE(10+digit) SPASI "-" SPASI CERT(3-7digit) [/ TAHUN opsional]
#     contoh: "1911000022019000006 - 001364" -> cert "001364"
#   Pola B: BASE(10+digit)"-"CERT(3-7digit)" / "TAHUN (tanpa spasi di
#           dash, TAPI wajib ada "/ TAHUN" di akhir)
#     contoh: "04240000000099-000001 / 2024"  -> cert "000001"
#             "04250000000099-000035 / 2025"  -> cert "000035"
# Kalau TIDAK ada spasi di dash DAN TIDAK ada "/ TAHUN" -> BUKAN sertifikat,
# dikembalikan kosong (biar diproses jalur breakdown/repetition biasa).
_CERT_TAIL_SPACED_RE = re.compile(
    r"^\s*(\d{10,})\s+-\s+(\d{3,7})\s*(?:/\s*(?:19|20)\d{2})?\s*$"
)
_CERT_TAIL_YEAR_RE = re.compile(
    r"^\s*(\d{10,})-(\d{3,7})\s*/\s*(?:19|20)\d{2}\s*$"
)


def extract_certificate(polis_raw) -> str:
    """Ambil blok sertifikat dari FAC_POLICY_NO mentah (Pola A atau Pola B
    di atas), lalu format lewat build_certificate(). Return '' kalau nilai
    polis tidak cocok kriteria ini -- termasuk pola "BASE-angka_pendek"
    polos tanpa spasi/tahun, yang BUKAN sertifikat (lihat catatan di atas)."""
    if pd.isna(polis_raw):
        return ""
    val = str(polis_raw).strip()
    if not val:
        return ""

    m = _CERT_TAIL_SPACED_RE.match(val)
    if m:
        return build_certificate(m.group(2))

    m = _CERT_TAIL_YEAR_RE.match(val)
    if m:
        return build_certificate(m.group(2))

    return ""


# ─────────────────────────────────────────────────────────────────────────────
# MESIN "REPETITION CHAIN" (TIDAK DIUBAH oleh update CERTIFICATE ini)
# ─────────────────────────────────────────────────────────────────────────────

_MONTH_NAMES_RE = (
    r"JANUARI|FEBRUARI|SEPTEMBER|NOVEMBER|DESEMBER|OKTOBER|AGUSTUS|MARET"
    r"|JANUARY|FEBRUARY|SEPTEMBER|NOVEMBER|DECEMBER|OCTOBER|AUGUST|MARCH"
    r"|APRIL|JUNI|JULI|JUNE|JULY|MEI|MAY"
    r"|SEPT|AUGUS|AGUS|OKT|DES|JAN|FEB|MAR|APR|JUN|JUL|AUG|SEP|OCT|NOV|DEC"
)

_JUNK_PHRASE_RE = re.compile(
    r"""
    SLIP\s+FOR\s+NEXT\s+\d+\s+MONTHS
  | TANPA\s+INDUK\s+SLIP
  | TANPA\s+SLIP\s+INDUK
  | SLIP\s+MASIH\s+PAKAI\s+SLIP\s+SEMENTARA
  | PAKAI\s+SLIP\s+SEMENTARA
  | SUPEN\s+SLIP
  | SUSPEN\s+SLIP
  | PAYMENT\s+NUMBER\s*:?
  | BINDER\s+NO\.?
  | REN\.?\s*POL\.?\s*NO\.?
  | REPLACEMENT
  | FINAL\s+ADJUSTMENT
  | ACTUAL\s+PREMIUM\s+FOR
  | MINDEPPREM\.?\s*FOR
  | MINDEP\s+PREMI(?:UM)?
  | MINDEP\s+FOR
  | \bMINDEP\b
  | \bMIN\.?
  | \bPERIOD\b
  | \bFOR\b
  | DATE\s+\d{1,2}\.\d{1,2}\.\d{2,4}
  | (?:19|20)\d{2}\s*-\s*\d{2}\b
  | \.?\s*\+?\s*SLIP\s+END\s*\d*(?:\s*\+\s*\d+\.?)*
  | \.?\s*\+?\s*END\.?\s*\d*(?:\s*\+\s*\d+\.?)*
  | DE(?:KLARASI|CLARASI)\s*:?
  | DECL\.?\s*:?
  | INTER\s+ISLAND
  | \bCOL\.?\s*NUMBER\b
  | \bCREDIT\s+NOTE\b
  | \bNPWN\b
  | BIND\.?\s*NO\.?
  | \bPN\.?\s+(?=\S)
  | \bR\.?\s*PREMIUM\b
  | \bMENJADI\b
  | \bS\s*/\s*D\b
  | \bDEPOSIT\b.*?\bFOR\b
  | AND\s+AC\.?\s*PREMI\.?\s*FOR
  | \bFROM\b
  | \bLIHAT\s+PAGE\s*\d*\b
  | \bEX\s*:?
  | \bLOL\b
  | \bPA\b
  | TANPA\s+ADA\s+SLIP\s+INDUK
  | \bIDR\b
  | \bUSD\b
  | \bCN\b\s*:?
  | \bSLIP\b
  | \bID\s*:
  | \bAND\b
  | \bPR\.
  | \bACTUAL\s+PREMIUM\b
  | \bEQ\b
  | \bNO\b\s*:?
  | \bMONTHS\b
  | \bPOLIS\b
  | \bVARIOUS\b
    """
    + r"| (?:" + _MONTH_NAMES_RE + r")\.?\s*\d{0,4}",
    re.IGNORECASE | re.VERBOSE,
)

_STANDALONE_NARRATIVE_RE = re.compile(r"SUSPENSE.*CEDANT", re.IGNORECASE)

_STANDALONE_P_RE = re.compile(
    r"^\s*(?:P[1-5]|TBA|VAR)(?:\s*[+,]\s*(?:P[1-5]|TBA|VAR))*\s*$", re.IGNORECASE
)

_YEAR_RE = re.compile(r"/\s*(?:19|20)\d{2}\b")
_YEAR_2DIGIT_RE = re.compile(r"/\s*\d{2}$")
_DOT_YEAR_RE = re.compile(r"^(.*)\.(?:19|20)?\d{2,4}$")

_MAX_SUFFIX_LEN = 7

_LITERAL_WHITELIST = {
    "00335-338",
    "**00335-338",
    "197-2006",
    "COMB012/13TPI X",
    "COMB-012/TPI/XII/10/REVISED 1",
    "COMB-012/013/TPI/XII/10/REVISED1",
    "#SPN-2011-2013#",
    "#SPN-2014#",
    "FSF1700024&22&20&18&16&14&12&10-085087075067069071073",
    "BBM/2014-160,180,135 / FEBRUARI 2014",
    "017.PAR.EQ.6.2025",
    "CN : 148/SRE-CM/CN/XII/2024",
    "282S/D286",
    "140051-053",
    "FUC1700060/C00455/17 1700002",
    # BUG TPI NEW -- biarkan apa adanya
    "P1 / SLIP ADA 4 SUDAH DI ELO",
    "0755/XOL-FAC-INDORE/VII/2026",
    "P1 / ADA 6 SLIP",
}


def _is_standalone_exception(val: str) -> bool:
    if val.strip() in _LITERAL_WHITELIST:
        return True
    return bool(_STANDALONE_P_RE.match(val)) or bool(_STANDALONE_NARRATIVE_RE.search(val))


# ── BUG TPI NEW: pola tambahan yang perlu dibuang sebelum tokenisasi ──
# "- 2023" / "- 2024" dst di akhir (tahun BERDIRI SENDIRI setelah dash,
# bukan bagian dari "/ TAHUN" yang sudah ditangani _YEAR_RE) -> dihapus.
_DASH_YEAR_TAIL_RE = re.compile(r"\s*-\s*(?:19|20)\d{2}\s*$")
# "-N/0" (mis. "-2/0", "-3/0") -> penanda endorsement/nol, BUKAN sertifikat
# ataupun suffix perulangan -> dibuang total.
_ENDORSEMENT_TAIL_RE = re.compile(r"-\d{1,3}\s*/\s*0\b")
# "/END N S/D M , X , Y" (deklarasi endorsement range) di akhir -> dibuang
# seluruhnya (base di depannya tetap dipertahankan).
_END_DECLARATION_TAIL_RE = re.compile(
    r"/?\s*END\s+\d+\s*S\s*/\s*D\s*\d+(?:\s*,\s*\d+)*\s*$", re.IGNORECASE
)
# "/END" polos di akhir (tanpa deklarasi S/D setelahnya) -> dibuang.
_BARE_END_TAIL_RE = re.compile(r"/\s*END\.?\s*$", re.IGNORECASE)
# "/ SA" di akhir -> dibuang (kode referensi tambahan, bukan bagian nomor).
_TRAILING_SA_RE = re.compile(r"/\s*SA\s*$", re.IGNORECASE)


def _strip_junk_and_year(text: str) -> str:
    text = _END_DECLARATION_TAIL_RE.sub("", text)
    text = _ENDORSEMENT_TAIL_RE.sub("", text)
    text = _BARE_END_TAIL_RE.sub("", text)
    text = _TRAILING_SA_RE.sub("", text)
    text = _JUNK_PHRASE_RE.sub("", text)
    text = _YEAR_RE.sub("", text)
    text = _DASH_YEAR_TAIL_RE.sub("", text)
    return text


def _strip_attached_noise(text: str) -> str:
    text = re.sub(
        r"(?:^|[\s/+,\-])P[1-5](?=\s|$|[+/,])",
        "",
        text,
        flags=re.IGNORECASE,
    )
    text = re.sub(
        r"\bTBA\b\s*/?\s*",
        "",
        text,
        flags=re.IGNORECASE,
    )
    text = re.sub(
        r"\bVAR\b\s*/?\s*",
        "",
        text,
        flags=re.IGNORECASE,
    )

    text = re.sub(r"^P-(?=\d)", "", text, flags=re.IGNORECASE)

    return text


def _clean_token(tok: str) -> str:
    tok = tok.strip(" :;/")
    if not tok:
        return tok
    if not re.search(r"[A-Za-z]", tok):
        tok = tok.replace(".", "")
    elif re.match(r"^[A-Za-z]+[\d.]+$", tok):
        tok = tok.replace(".", "")
    return tok


# Panjang kode standar TPI "LLL+DDDDDDD" (3 huruf + 7 digit = 10 karakter),
# mis. "PVF2300041", "FVF2300113" -- dipakai untuk mendeteksi base huruf yang
# SENGAJA dipotong pendek (mis. "PVF23000" cuma 8 char) dan perlu DISAMBUNG
# (konkatenasi), bukan digantikan digit akhirnya seperti base yang sudah utuh.
_FULL_CODE_LEN = 10


def _classify_and_reconstruct(tokens: list, year_tail: str = None) -> list:
    results = []
    current_base = None
    # Token huruf+angka yang "belum lengkap" (lebih pendek dari
    # _FULL_CODE_LEN) ditahan dulu (belum di-append) -- kalau token berikutnya
    # ternyata melengkapinya jadi persis _FULL_CODE_LEN karakter, base mentah
    # ini TIDAK ikut muncul di hasil (hanya versi lengkapnya). Kalau ternyata
    # tidak dilengkapi (token berikutnya bukan pelengkap), base mentah ini
    # tetap dimunculkan apa adanya.
    pending_incomplete = None

    for tok in tokens:
        tok = _clean_token(tok)
        if not tok:
            continue
        digits_only = re.match(r"^\d+$", tok) is not None

        # BUG TPI NEW -- suffix angka kadang kebawa 2-digit sisa TAHUN yang
        # baru saja dihapus (mis. "PVF1800092/2018+18091" -> "18" di depan
        # "18091" itu sisa "2018", bukan bagian nomor sertifikat/urutan).
        # Kalau suffix diawali persis 2 digit ekor tahun tsb, buang dulu 2
        # digit itu sebelum direkonstruksi -- HANYA kalau sisanya masih
        # panjang suffix yang masuk akal (1-7 digit).
        if (
            digits_only
            and year_tail
            and tok.startswith(year_tail)
            and 1 <= len(tok) - 2 <= _MAX_SUFFIX_LEN
        ):
            tok = tok[2:]

        base_has_letter = bool(current_base and re.search(r"[A-Za-z]", current_base))

        if digits_only and current_base is not None:
            # Base huruf yang masih PENDEK -> suffix ini MELENGKAPI base lewat
            # konkatenasi langsung (mis. "PVF23000"+"41" -> "PVF2300041"),
            # BUKAN mengganti digit akhir.
            if (
                base_has_letter
                and len(current_base) < _FULL_CODE_LEN
                and len(current_base) + len(tok) == _FULL_CODE_LEN
            ):
                completed = current_base + tok
                pending_incomplete = None  # base mentah tidak usah dimunculkan
                results.append(completed)
                current_base = completed
                continue
            # Base sudah cukup panjang -> suffix pendek MENGGANTI N digit
            # terakhir base (perilaku lama, dipertahankan).
            if (
                len(tok) < len(current_base)
                and len(tok) <= _MAX_SUFFIX_LEN
                and (base_has_letter or len(current_base) >= 7)
            ):
                if pending_incomplete:
                    results.append(pending_incomplete)
                    pending_incomplete = None
                results.append(current_base[: -len(tok)] + tok)
                continue

        if pending_incomplete:
            results.append(pending_incomplete)
            pending_incomplete = None

        # Token huruf+angka yang lebih pendek dari _FULL_CODE_LEN -> mungkin
        # base "terpotong" yang akan disambung suffix berikutnya. Tahan dulu,
        # jangan langsung di-append.
        if not digits_only and re.search(r"\d{4,}$", tok) and len(tok) < _FULL_CODE_LEN:
            pending_incomplete = tok
            current_base = tok
            continue

        results.append(tok)
        # BUG FIX: current_base HARUS diperbarui juga untuk token huruf+angka
        # (bukan cuma angka murni) -- sebelumnya suffix pendek sesudah kode
        # seperti "PFV2100285" tidak pernah direkonstruksi karena current_base
        # tidak pernah di-set dari token semacam itu.
        if digits_only and len(tok) >= 7:
            current_base = tok
        elif not digits_only and re.search(r"\d{4,}$", tok):
            current_base = tok

    if pending_incomplete:
        results.append(pending_incomplete)

    return results


def _dedup_preserve_order(items: list) -> list:
    seen = set()
    deduped = []
    for r in items:
        if r not in seen:
            seen.add(r)
            deduped.append(r)
    return deduped


def expand_repetition_chain(val, patterns=None) -> list:
    if pd.isna(val):
        return []
    val = str(val).strip()
    if not val:
        return []

    if _is_standalone_exception(val):
        return [_normalize_spaces(val)]

    # BUG TPI NEW -- kalau val cocok kriteria SERTIFIKAT (Pola A/B, lihat
    # extract_certificate()), JANGAN di-breakdown jadi 2 nilai. Sertifikat
    # cukup muncul di kolom CERTIFICATE saja; POLIS cukup base-nya saja.
    # (feedback user 2x: breakdown untuk kasus ini adalah bug.)
    _cert_check_m = _CERT_TAIL_SPACED_RE.match(val) or _CERT_TAIL_YEAR_RE.match(val)
    if _cert_check_m:
        return [_cert_check_m.group(1)]

    _tba_prefix_m = re.match(r"^TBA\s*/\s*(.+)$", val, re.IGNORECASE)
    if _tba_prefix_m:
        return [_normalize_spaces(_tba_prefix_m.group(1))]

    _had_narrative = bool(_JUNK_PHRASE_RE.search(val))

    # BUG TPI NEW -- tangkap 2 digit ekor tahun yang akan dibuang (mis.
    # "/2018" -> "18"), dipakai nanti untuk membuang sisa digit tahun yang
    # nyangkut di depan suffix pendek (lihat _classify_and_reconstruct).
    _year_m = _YEAR_RE.search(val)
    _year_tail = _year_m.group(0)[-2:] if _year_m else None

    cleaned_text = _strip_junk_and_year(val)
    cleaned_text = _strip_attached_noise(cleaned_text)
    cleaned_text = re.sub(
        r"(?<![A-Za-z0-9])P[1-5](?![A-Za-z0-9])",
        "",
        cleaned_text,
        flags=re.IGNORECASE,
    )
    cleaned_text = _YEAR_2DIGIT_RE.sub("", cleaned_text)
    if cleaned_text.count(".") == 1:
        m = _DOT_YEAR_RE.match(cleaned_text)
        if m:
            cleaned_text = m.group(1)

    if "." in cleaned_text and not re.search(r"[A-Za-z]", cleaned_text):
        _dot_parts = cleaned_text.split(".")
        if len(_dot_parts) > 1 and all(p.isdigit() for p in _dot_parts):
            _rest_lens = {len(p) for p in _dot_parts[1:]}
            if len(_rest_lens) == 1 and next(iter(_rest_lens)) < len(_dot_parts[0]):
                cleaned_text = "+".join(_dot_parts)
    cleaned_text = re.sub(r"[/\-+,]\s*[A-Za-z]\s*$", "", cleaned_text)
    cleaned_text = cleaned_text.replace(":", "").replace(";", "").replace("(", "").replace(")", "")
    cleaned_text = cleaned_text.strip()
    if not cleaned_text:
        return [_normalize_spaces(val)]

    cleaned_text = re.sub(r"\b([A-Za-z]{2,5})\s+(\d{5,})\b", r"\1\2", cleaned_text)
    cleaned_text = cleaned_text.replace("*", "")

    raw_tokens = [t for t in re.split(r"\s+|[+\-,/&]", cleaned_text) if t.strip()]
    raw_tokens = [_clean_token(t) for t in raw_tokens]
    raw_tokens = [t for t in raw_tokens if t]
    if not raw_tokens:
        return [_normalize_spaces(val)]

    if len(raw_tokens) > 1 and all(
        re.match(r"^\d+$", t) and len(t) <= 4 for t in raw_tokens
    ):
        return [_normalize_spaces(val)]

    reconstructed = _classify_and_reconstruct(raw_tokens, year_tail=_year_tail)

    if patterns:
        reconstructed = [refine_with_known_pattern(r, patterns) for r in reconstructed]

    reconstructed = _dedup_preserve_order(reconstructed)

    if len(reconstructed) == 1:
        return [_normalize_spaces(reconstructed[0])]

    if len(reconstructed) <= MAX_SPLIT_COLS:
        return [_normalize_spaces(r) for r in reconstructed]

    if _had_narrative:
        return [_normalize_spaces(reconstructed[0])]
    return [_normalize_spaces(cleaned_text).strip(" +,-/")]


# ─────────────────────────────────────────────────────────────────────────────
# CLEAN POLIS
# ─────────────────────────────────────────────────────────────────────────────

def clean_polis(val) -> list:
    if pd.isna(val):
        return []
    val = str(val).strip()
    if not val:
        return []

    comp_polis = _COMPOSITE_POLIS_RE.search(val)
    if comp_polis:
        return [comp_polis.group(1)]

    _trigger = (
        "+" in val
        or "/" in val
        or "-" in val
        or "&" in val
        or "." in val
        or re.search(r"\s{2,}", val)
        or _YEAR_RE.search(val)
        or re.search(r"\bP[1-5]\b", val, re.IGNORECASE)
        or re.search(r"\bTBA\b", val, re.IGNORECASE)
        or re.search(r"\bVAR\b", val, re.IGNORECASE)
        or _JUNK_PHRASE_RE.search(val)
    )
    if _trigger:
        return expand_repetition_chain(val, KNOWN_POLIS_PATTERNS)

    if POLIS_EXCEPTION_RE.search(val):
        return [_normalize_spaces(val)]

    val_upper = val.upper().strip()
    if re.match(r"^VARIOUS\s*$", val_upper):
        return [val.strip()]
    if re.match(r"^VARIOUS\s*-\s*SEE\s+ATTACH", val_upper):
        return [val.strip()]
    if re.match(r"^TBA\s*$", val_upper):
        return [val.strip()]

    m = re.match(r"^(\d[\d\-]+)\.\d{1,6}$", val.strip())
    if m:
        base = m.group(1).strip()
        if _is_valid_polis_token(base):
            return [base]

    if re.search(r"S/D", val, re.IGNORECASE):
        sd_parts = re.split(r"\s{2,}", val.strip())
        left = sd_parts[0] if sd_parts else val
        base_m = re.match(r"^(\d[\d\-]+?)-\d+S/D\d+", left, re.IGNORECASE) or \
                 re.match(r"^(\d[\d\-]+)S/D\d+", left, re.IGNORECASE)
        results = [base_m.group(1).strip()] if base_m else []
        for part in sd_parts[1:]:
            part = re.sub(r"\+?\s*\bTBA\b\s*", "", part, flags=re.IGNORECASE).strip().strip("+")
            results += [t for t in part.split() if _is_valid_polis_token(t.strip())]
        if results:
            return _cap_or_join(results)

    if re.match(r"^\s*TBA\s{2,}", val, re.IGNORECASE):
        rest = re.sub(r"^\s*TBA\s+", "", val, flags=re.IGNORECASE).strip()
        tokens = _extract_polis_tokens(rest)
        if tokens:
            return _cap_or_join(tokens)

    if "+" in val:
        blocks = re.split(r"\s{2,}", val.strip())
        left_parts = [p.strip() for p in blocks[0].split("+")
                      if p.strip() and not re.match(r"^TBA$", p.strip(), re.IGNORECASE)]
        right_blocks = blocks[1:]
        if left_parts:
            first = left_parts[0]
            rest_parts = left_parts[1:]
            any_long = any(len(r) >= 10 for r in rest_parts)
            all_suffix = rest_parts and all(re.match(r"^\d{1,6}$", r) for r in rest_parts)
            if all_suffix and len(first) >= 10 and not any_long:
                results = [first[:-len(s)] + s if len(first) > len(s) else first + s for s in rest_parts]
                for rb in right_blocks:
                    results += [t for t in rb.split() if _is_valid_polis_token(t.strip())]
                return _cap_or_join(results) if results else [first]
            all_tokens = _extract_polis_tokens(val)
            if all_tokens:
                return _cap_or_join(all_tokens)

    if re.search(r"\s{2,}", val):
        tokens = _extract_polis_tokens(val)
        if tokens:
            return _cap_or_join(tokens)

    if "-" in val and re.match(r"^[A-Za-z0-9\-]+$", val):
        dash_parts = val.split("-")
        if any(len(p.strip()) >= 8 and re.search(r"\d", p) for p in dash_parts):
            expanded = _expand_dash_chain(val)
            if len(expanded) > 1:
                return expanded

    m = re.match(r"^(\d[\d\-]+)\s+((?:(?:\d{1,2})|¿)(?:/(?:(?:\d{1,2})|¿))+)\s*$", val)
    if m:
        base = m.group(1).strip()
        results = [base[:-2] + s.zfill(2) if len(base) >= 2 else base + s.zfill(2)
                   for s in m.group(2).split("/") if re.match(r"^\d{1,2}$", s.strip())]
        if results:
            return _cap_or_join(results)

    val = re.sub(r"\bVARIOUS\b\s*", "", val, flags=re.IGNORECASE).strip()
    val = re.sub(r"\bVAR\b\s*", "", val, flags=re.IGNORECASE).strip()
    val = val.replace("¿", "")

    parts = re.split(r"\s*/\s*", val)
    cleaned = [_normalize_spaces(p).strip(" -/") for p in parts if _normalize_spaces(p).strip(" -/")]

    if len(cleaned) > 1 and all(len(c) < 7 for c in cleaned):
        return [val.strip()] if val.strip() else []

    return _cap_or_join(cleaned) if cleaned else []


# ─────────────────────────────────────────────────────────────────────────────
# CLEAN SLIP
# ─────────────────────────────────────────────────────────────────────────────

def clean_slip(val) -> list:
    if pd.isna(val):
        return []
    val = str(val).strip()
    if not val:
        return []

    comp_matches = _COMPOSITE_SLIP_RE.findall(val)
    if comp_matches:
        return comp_matches

    _trigger = (
        "+" in val
        or "/" in val
        or "-" in val
        or "&" in val
        or "." in val
        or re.search(r"\s{2,}", val)
        or _YEAR_RE.search(val)
        or re.search(r"\bP[1-5]\b", val, re.IGNORECASE)
        or re.search(r"\bTBA\b", val, re.IGNORECASE)
        or re.search(r"\bVAR\b", val, re.IGNORECASE)
        or _JUNK_PHRASE_RE.search(val)
    )
    if _trigger:
        return expand_repetition_chain(val, KNOWN_SLIP_PATTERNS)

    blocks = re.split(r"\s{2,}", val.strip())
    if len(blocks) >= 2:
        left = blocks[0]
        left_check = re.sub(
            r"\b(?:SUMMARY|BORDERO?|BORDERA|BORDRO|SINGGLESHIPMENT|VARIOUS)\b",
            "", left, flags=re.IGNORECASE,
        )
        left_check = _SLIP_NOISE_RE.sub("", left_check)
        left_check = re.sub(r"[\s\+\-/]+", "", left_check).strip()
        if not _is_valid_slip_token(left_check):
            tokens = _extract_slip_tokens("  ".join(blocks[1:]))
            if tokens:
                return _cap_or_join(tokens)

    tokens = _extract_slip_tokens(val)
    if tokens:
        return _cap_or_join(tokens)

    cleaned = re.sub(r"\bVARIOUS\b\s*", "", val, flags=re.IGNORECASE)
    cleaned = re.sub(r"^\s*-\s*", "", cleaned).strip()
    cleaned = _normalize_spaces(cleaned)
    return [cleaned] if cleaned else []


# ─────────────────────────────────────────────────────────────────────────────
# CLEAN INSURED
# ─────────────────────────────────────────────────────────────────────────────

_INSURED_PREFIX_RE = re.compile(r"^\s*(?:PT|CV)\.?\s+", re.IGNORECASE)

_INSURED_GELAR_RE = re.compile(
    r"""
    \b(?:BAPAK|IBU|SDRI?|NYONYA|NY|TUAN|MR|MRS|MS|
        DR|IR|PROF|HJ|H)\b\.?\s*
    """,
    re.IGNORECASE | re.VERBOSE,
)

_INSURED_SPLIT_RE = r"\bQQ\b|\bAND\s*/\s*OR\b|/|,|–|¿"

# BUG TPI NEW -- singkatan dikenal yang harus dibuang kalau bentuk
# panjangnya JUGA muncul di nilai INSURED yang sama (baik di segmen terpisah
# via "/", "," dst, MAUPUN menempel langsung tanpa pemisah di string yang
# sama). Ini hanya menyentuh Data 1, sesuai instruksi user ("tanpa mengubah
# rules pada insured lain").
#   - "VICO" = singkatan "VIRGINIA INDONESIA COMPANY"
#   - "STPI" = singkatan "SEKOLAH TINGGI PENERBANGAN INDONESIA"
#   - "ANTAM" = singkatan "ANEKA TAMBANG"
_KNOWN_INSURED_ABBREV = {
    "VICO": "VIRGINIA INDONESIA",
    "STPI": "SEKOLAH TINGGI PENERBANGAN INDONESIA",
    "ANTAM": "ANEKA TAMBANG",
}


def _strip_embedded_known_abbrev(text: str) -> str:
    """Buang token singkatan dikenal (mis. 'STPI', 'VICO', 'ANTAM') dari
    `text` KALAU bentuk panjangnya juga ada di `text` yang sama -- berlaku
    baik singkatan itu berdiri sebagai segmen terpisah (dipisah '/' atau
    ',') maupun menempel langsung tanpa pemisah di string yang sama (mis.
    "SEKOLAH TINGGI PENERBANGAN INDONESIA STPI PLP CURUG")."""
    upper = text.upper()
    for abbr, full in _KNOWN_INSURED_ABBREV.items():
        if full in upper:
            text = re.sub(rf"\b{abbr}\b\.?\s*", " ", text, flags=re.IGNORECASE)
    return _normalize_spaces(text)


def _remove_polis_slip_from_insured(text: str, polis_ori, slip_ori) -> str:
    for token in (polis_ori, slip_ori):
        if pd.notna(token):
            t = str(token).strip()
            if t and t != "-":
                text = text.replace(t, "")
    return text


def clean_insured(val, polis_ori=None, slip_ori=None) -> list:
    if pd.isna(val):
        return []
    val = str(val).strip()
    if not val:
        return []

    val = _remove_polis_slip_from_insured(val, polis_ori, slip_ori)
    val = _strip_embedded_known_abbrev(val)

    val = re.sub(r",\s*(?:PT|CV)\.?\s+", " ", val, flags=re.IGNORECASE)
    val = re.sub(r",\s*KSO\.?\s*$", " KSO", val, flags=re.IGNORECASE)

    raw_parts = re.split(_INSURED_SPLIT_RE, val, flags=re.IGNORECASE)
    raw_parts = _merge_digit_start_parts(raw_parts)
    raw_parts = _merge_region_prefix_parts(raw_parts)

    cleaned = []
    for p in raw_parts:
        p = _normalize_spaces(p.strip())
        if len(p) <= 2:
            continue
        if p.upper().strip() in INSURED_JUNK_WORDS:
            continue
        if re.match(r"^[^a-zA-Z0-9]+$", p):
            continue
        p_clean = _clean_insured_name(p)
        if not p_clean or len(p_clean) <= 2:
            continue
        if p_clean.upper().strip() in INSURED_JUNK_WORDS:
            continue
        cleaned.append(p_clean)

    if not cleaned:
        fallback = _clean_insured_name(_normalize_spaces(val))
        return [fallback] if fallback else []

    final = []
    for name in cleaned:
        name = _strip_trailing_abbrev_word(name, final)
        if any(_is_abbreviation_of(name, prior) for prior in final):
            continue
        final.append(name)

    return _cap_or_join(final)


# ─────────────────────────────────────────────────────────────────────────────
# MITRA BISNIS
# ─────────────────────────────────────────────────────────────────────────────

def get_business_partner(comp_name2, comp_name) -> str:
    broker = "" if pd.isna(comp_name2) else str(comp_name2).strip()
    if not broker or broker.upper() == "DIRECT":
        result = "" if pd.isna(comp_name) else str(comp_name).strip()
    else:
        result = broker
    result = re.sub(r"\.", " ", result)
    return re.sub(r"\s+", " ", result).upper().strip()


# ─────────────────────────────────────────────────────────────────────────────
# PROSES UTAMA
# ─────────────────────────────────────────────────────────────────────────────

def _insert_clean_columns(df, all_lists, prefix, max_cols):
    added = []
    for i in range(1, max_cols + 1):
        col_name = f"{prefix}_{i}"
        df[col_name] = [lst[i - 1] if i - 1 < len(lst) else None for lst in all_lists]
        added.append(col_name)
    return added


def process_data(input_file: str, output_file: str) -> None:
    print(f"[1/5] Membaca data dari: {input_file} ...")
    df = pd.read_excel(input_file, header=0)
    print(f"      Total baris keseluruhan: {len(df):,}")

    if CEDANT_COL not in df.columns:
        print(f"\n[ERROR] Kolom '{CEDANT_COL}' tidak ditemukan!")
        print(f"        Kolom tersedia: {list(df.columns)}")
        return

    df[CEDANT_COL] = df[CEDANT_COL].astype(str).str.strip()
    df = df[df[CEDANT_COL] == CEDANT_VALUE].copy()
    print(f"[2/5] Filter cedant '{CEDANT_VALUE}': {len(df):,} baris ditemukan.")

    if df.empty:
        print("\n[WARN] Tidak ada data setelah filter. Proses dihentikan.")
        return

    for col in [POLIS_COL, SLIP_COL, INSURED_COL, BROKER_COL]:
        if col not in df.columns:
            print(f"\n[ERROR] Kolom '{col}' tidak ditemukan di file!")
            print(f"        Kolom tersedia: {list(df.columns)}")
            return

    business_partner_values = [get_business_partner(row[BROKER_COL], row[CEDANT_COL]) for _, row in df.iterrows()]

    df.rename(columns={INSURED_COL: "FAC_INSURED_ORI"}, inplace=True)

    insert_pos = list(df.columns).index(BROKER_COL) + 1 if BROKER_COL in df.columns else len(df.columns)
    df.insert(insert_pos, "BUSINESS_PARTNER", business_partner_values)

    print("[3/5] Menjalankan proses cleaning ...")

    all_clean_polis, all_clean_slip, all_clean_ins, all_certificate = [], [], [], []
    max_polis = max_slip = max_ins = 1

    for _, row in df.iterrows():
        c_polis = _refine_list(clean_polis(row.get(POLIS_COL, "")), KNOWN_POLIS_PATTERNS)
        c_slip  = _refine_list(clean_slip(row.get(SLIP_COL, "")), KNOWN_SLIP_PATTERNS)
        c_ins   = clean_insured(row.get("FAC_INSURED_ORI", ""), row.get(POLIS_COL, ""), row.get(SLIP_COL, ""))
        cert    = extract_certificate(row.get(POLIS_COL, ""))

        max_polis = max(max_polis, len(c_polis))
        max_slip  = max(max_slip, len(c_slip))
        max_ins   = max(max_ins, len(c_ins))

        all_clean_polis.append(c_polis)
        all_clean_slip.append(c_slip)
        all_clean_ins.append(c_ins)
        all_certificate.append(cert)

    print("[4/5] Menyusun kolom output ...")

    polis_clean_cols = []

    for i in range(1, max_polis + 1):

        col_name = f"POLIS_CLEAN_{i}"

        df[col_name] = [
            x[i - 1] if i - 1 < len(x) else None
            for x in all_clean_polis
        ]

        polis_clean_cols.append(col_name)

    certificate_cols = [
        "CERTIFICATE_1"
    ]

    df["CERTIFICATE_1"] = [
        x if x not in (None, "") else None
        for x in all_certificate
    ]


    slip_clean_cols = []

    for i in range(1, max_slip + 1):

        col_name = f"SLIP_CLEAN_{i}"

        df[col_name] = [
            x[i - 1] if i - 1 < len(x) else None
            for x in all_clean_slip
        ]

        slip_clean_cols.append(col_name)

    insured_cols = []

    for i in range(1, max_ins + 1):

        col_name = f"INSURED_CLEAN_{i}"

        df[col_name] = [
            x[i - 1] if i - 1 < len(x) else None
            for x in all_clean_ins
        ]

        insured_cols.append(col_name)

    new_columns = []

    cleaning_cols = set(
        polis_clean_cols
        + certificate_cols
        + slip_clean_cols
        + insured_cols
    )

    for col in df.columns:

        if col in cleaning_cols:
            continue

        new_columns.append(col)


        if col == "FAC_POLICY_NO":

            for i in range(1, max_polis + 1):

                new_columns.append(f"POLIS_CLEAN_{i}")

                # Certificate pertama SELALU ada.
                if i == 1:
                    new_columns.append("CERTIFICATE_1")

        elif col == "FAC_SLIP":

            new_columns.extend(slip_clean_cols)

        elif col == "FAC_INSURED_ORI":

            new_columns.extend(insured_cols)


    df = df[new_columns]

    os.makedirs(os.path.dirname(output_file), exist_ok=True)
    print(f"[5/5] Menyimpan hasil ke: {output_file} ...")
    df.to_excel(output_file, index=False)

    print(f"\n{'=' * 55}")
    print("  [OK] Selesai!")
    print(f"  Total baris output  : {len(df):,}")
    print(f"  Total kolom output  : {len(df.columns)}")
    print(f"  File disimpan di    : {output_file}")
    print(f"{'=' * 55}")


if __name__ == "__main__":
    process_data(INPUT_FILE, OUTPUT_FILE)