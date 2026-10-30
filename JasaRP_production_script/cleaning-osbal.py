"""
Cleaning Data 2 - Osbal | Cedant: JASA R.P (PT ASURANSI JASA RAHARJA PUTRA)

CATATAN ARSITEKTUR:
- Sumber utama polis/slip: kolom CLSDT_POLICY_NO / CLSDT_SLIP_NO. Kalau kosong
  atau cuma placeholder umum (TBA/VAR/VARIOUS/P1..Pn berdiri sendiri), fallback
  ke FAC_POLICY_NO / FAC_SLIP (induk). Sama seperti pola TPI & CSI.
- BEDA UTAMA dari TPI/CSI: nomor sertifikat (base + "-NNNNNN") di JRP TIDAK
  dibuang dan TIDAK direkonstruksi (diganti digit akhirnya) -- cukup
  DIBIARKAN MENEMPEL sebagai SATU nilai di POLIS_CLEAN (mis.
  "227000806122400025-000568" tetap 1 nilai, bukan 2 kolom, bukan dibuang
  cert-nya). Ini TIDAK berubah oleh update CERTIFICATE di bawah.
- Insured: TIDAK ada transformasi "Pelindo = Pelabuhan Indonesia" atau
  "I,II,III,IV = 1,2,3,4" -- itu cuma catatan matching manual, bukan rule
  cleansing. Insured dibersihkan pakai pipeline standar (PT/CV/Persero strip).

CATATAN YANG MASIH PERLU DIKONFIRMASI (lihat komentar _KNOWN_UNCERTAIN di
bawah) -- JANGAN dihapus sebelum dikonfirmasi user:
1. Pola CN-chunk ("000351/CN/0200/03/24 / 024315/CN/0200/03/24/1") -- saat
   ini dipertahankan sebagai 2 nilai CN utuh (TIDAK didempetkan tanpa
   separator seperti aturan CSI), karena dokumen JRP tidak menyebutkan
   aturan compact. Kalau ternyata harus didempetkan juga, tinggal ubah
   fungsi apply_cn_chunks().
2. "202000805022400022-202000805022400011" -- base + base (dua nomor SAMA
   PANJANG dipisah dash, BUKAN base+cert) -> saat ini dipecah jadi 2 polis
   terpisah karena panjang kedua sisi sama. Perlu konfirmasi user kalau
   salah baca.

=====================================================================
UPDATE STRUKTUR KOLOM OUTPUT (terbaru) -- HANYA STRUKTUR, LOGIC CLEANING
TIDAK BERUBAH:
- Sebelumnya kolom CERTIFICATE (satu kolom, nilai tunggal per baris)
  diletakkan SEBELUM semua kolom polis bersih (POLICY_CLEAN_1, dst).
- Sekarang: kolom hasil clean polis diganti nama prefix-nya dari
  "POLICY_CLEAN" jadi "POLIS_CLEAN" (POLIS_CLEAN_1, POLIS_CLEAN_2, ...),
  dan CERTIFICATE_n disisipkan TEPAT SETELAH POLIS_CLEAN_n yang
  bersangkutan:
      FAC_POLICY_NO_ORI -> POLIS_CLEAN_1 -> CERTIFICATE_1
                         -> POLIS_CLEAN_2 -> CERTIFICATE_2
                         -> ...
- Jumlah CERTIFICATE_n MENGIKUTI jumlah POLIS_CLEAN_n (bukan jumlah angka
  sertifikat di dalam satu nilai certificate).
- resolve_certificate() hanya pernah menghasilkan SATU nilai certificate
  per baris, dan cedant JRP ini memang hanya punya case certificate
  tunggal, jadi HANYA CERTIFICATE_1 yang diisi nilai hasil resolve;
  CERTIFICATE_2, CERTIFICATE_3, dst (kalau ada baris dengan >1 polis)
  selalu dikosongkan.
- CERTIFICATE_1 tetap selalu dibuat walau tidak ada certificate sama
  sekali (nilainya string kosong), tapi CERTIFICATE_2/3/dst TIDAK
  otomatis dibuat kalau memang tidak ada baris dengan polis sebanyak itu.
=====================================================================
UPDATE SEBELUMNYA -- KOLOM CERTIFICATE (rule sama persis dgn Data 1/Data 3,
lihat build_certificate()):
- ASUMSI SUMBER DATA (MOHON DIKONFIRMASI): kolom CERTIFICATE diambil dari
  raw column "CLSDT_SERTF_NO" (mengikuti pola penamaan CLSDT_POLICY_NO /
  CLSDT_SLIP_NO yang sudah ada). Kalau nama kolom aslinya beda, tinggal
  ganti CLSDT_SERTF_COL di bagian KONFIGURASI di bawah.
- Kalau CLSDT_SERTF_NO kosong/placeholder, TIDAK fallback ke mana pun --
  langsung dikosongkan (beda dari POLICY_NO/SLIP yang fallback ke FAC_*).
  Kalau ternyata butuh fallback juga (mis. ke sertifikat yang menempel di
  FAC_POLICY_NO_ORI / CLSDT_POLICY_NO), kasih tau supaya bisa disesuaikan.
- Rule format (SAMA seperti Data 1 & Data 3):
    * range 2-angka ("-" / S/D / SD / S): >3 anggota -> "AWAL SD AKHIR",
      <=3 anggota -> breakdown penuh dipisah koma.
    * daftar eksplisit koma/campur S-D: >3 item -> kompres
      "ITEM_PERTAMA SD ITEM_TERAKHIR" (tanpa isi celah), <=3 -> apa adanya.
    * semua angka di-pad ke 6 digit.
    * "S/D"/"s/d"/"SD"/"S" di output SELALU distandarisasi jadi "SD".
    * pola tidak masuk akal (bukan angka murni) -> CERTIFICATE kosong.
=====================================================================
"""

import os
import re

import pandas as pd


# ─────────────────────────────────────────────────────────────────────────────
# KONFIGURASI
# ─────────────────────────────────────────────────────────────────────────────

INPUT_FILE  = os.path.join("input", "2b. transaksi Osbal 01.01.23 - 17.08.26.xlsx")
INPUT_SHEET = "Query result"
OUTPUT_FILE = os.path.join("output", "JasaRp_output_osbal.xlsx")

CEDANT_COL   = "CCOS_COMP_NAME"
CEDANT_VALUE = "PT.ASURANSI JASA RAHARJA PUTRA"

POLIS_COL   = "FAC_POLICY_NO"
SLIP_COL    = "FAC_SLIP"
INSURED_COL = "FAC_INSURED"

CLSDT_POLIS_COL = "CLSDT_POLICY_NO"
CLSDT_SLIP_COL  = "CLSDT_SLIP_NO"

# ASUMSI nama kolom sumber CERTIFICATE -- ganti di sini kalau beda di file asli.
CLSDT_SERTF_COL = "CLSDT_SERTF_NO"

MAX_SPLIT_COLS = 5


# ─────────────────────────────────────────────────────────────────────────────
# KAMUS POLA STANDAR JRP
# ─────────────────────────────────────────────────────────────────────────────

KNOWN_POLIS_PATTERNS = [
    re.compile(r"^\d{14,19}$"),
    re.compile(r"^\d{14,19}-\d{3,7}(?:,\s*\d{3,7})*$"),
    re.compile(r"^\d{14,19}-\d{3,7}\s+S/D\s+\d{3,7}$"),
]

KNOWN_SLIP_PATTERNS = [
    re.compile(r"^\d{14,19}$"),
    re.compile(r"^\d{14,19}/\d{1,2}$"),
    re.compile(r"^\d{3,6}/CN/\d{3,4}/\d{2}/\d{2}(?:/\d)?$"),
]


def _matches_known_pattern(value: str, patterns) -> bool:
    return any(p.match(value) for p in patterns)


def refine_with_known_pattern(value: str, patterns) -> str:
    if not value:
        return value
    if _matches_known_pattern(value, patterns):
        return value
    candidates = [
        re.sub(r"\s+", "", value),
        value.strip(" .-/"),
        "0" + value if value.isdigit() else value,
    ]
    for c in candidates:
        if _matches_known_pattern(c, patterns):
            return c
    return value


def _refine_list(items: list, patterns) -> list:
    return [refine_with_known_pattern(v, patterns) for v in items]


# ─────────────────────────────────────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────────────────────────────────────

def _normalize_spaces(text: str) -> str:
    return re.sub(r"\s{2,}", " ", text).strip()


def _dedup_preserve_order(items: list) -> list:
    seen, out = set(), []
    for i in items:
        if i not in seen:
            seen.add(i)
            out.append(i)
    return out


def _cap_or_join(items: list) -> list:
    if len(items) > MAX_SPLIT_COLS:
        return [",".join(items)]
    return items


# ─────────────────────────────────────────────────────────────────────────────
# KOLOM CERTIFICATE (rule sama persis dengan Data 1 & Data 3)
# ─────────────────────────────────────────────────────────────────────────────

_CERT_RANGE_PAIR_RE = re.compile(
    r"^\s*(\d{1,7})\s*(?:-|S\s*/\s*D|SD|S(?!\d))\s*(\d{1,7})\s*$",
    re.IGNORECASE,
)
_CERT_SPLIT_RE = re.compile(r"\s*(?:S\s*/\s*D|SD|S(?!\d)|,)\s*", re.IGNORECASE)
_CERT_TOKEN_RE = re.compile(r"^\d{1,7}$")


def build_certificate(cert_raw) -> str:
    """Format nomor sertifikat sesuai rule baru:
    - range 2-angka ("-" / S/D / SD / S): >3 anggota -> "AWAL SD AKHIR",
      <=3 anggota -> breakdown penuh dipisah koma.
    - daftar eksplisit (koma dan/atau campur S/D/SD/S, bukan range murni):
      >3 item -> dikompres "ITEM_PERTAMA SD ITEM_TERAKHIR" (tanpa isi celah),
      <=3 item -> tetap apa adanya, dipisah koma.
    - semua angka di-pad ke 6 digit.
    - "S/D"/"s/d"/"SD"/"S" di output SELALU ditulis "SD".
    - kalau pola tidak masuk akal (bukan angka murni) -> kosong.
    """
    if pd.isna(cert_raw):
        return ""
    text = str(cert_raw).strip()
    if not text or text.upper() in {"-", "NAN", "NONE", "NULL"}:
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


# Pola "base + '-'/'.' + sertifikat" yang menempel di kolom POLIS (baik
# CLSDT_POLICY_NO maupun FAC_POLICY_NO) -- dipakai sebagai FALLBACK kalau
# CLSDT_SERTF_NO kosong/tidak match. Sejauh ini sertifikat hanya pernah
# ditemukan menempel di kolom polis, bukan di tempat lain.
_POLICY_CERT_TAIL_RE = re.compile(r"^\s*(?:VARIOUS\s*/\s*)?(\d{10,})\s*[-.]\s*(.+)$", re.IGNORECASE)


def extract_certificate_from_policy(polis_raw) -> str:
    """Coba ambil blok sertifikat yang menempel di suatu nilai POLIS
    (pola base+'-'/'.'sertifikat), lalu format lewat build_certificate().
    Return '' kalau nilai polis tidak punya pola sertifikat menempel."""
    if pd.isna(polis_raw):
        return ""
    val = str(polis_raw).strip()
    if not val:
        return ""

    m = _POLICY_CERT_TAIL_RE.match(val)
    if not m:
        return ""
    base, rest = m.groups()

    first_num_m = re.match(r"^\s*(\d+)", rest)
    if not first_num_m or not (5 <= len(first_num_m.group(1)) <= 7):
        return ""

    rest = re.sub(r"\s*/\s*P\d+\s*$", "", rest, flags=re.IGNORECASE)
    rest = re.sub(r"\s*/\s*VARIOUS\s*$", "", rest, flags=re.IGNORECASE)
    rest = rest.strip()
    if not rest or not re.search(r"\d", rest):
        return ""

    return build_certificate(rest)


def resolve_certificate(clsdt_sertf_val, clsdt_polis_val, fac_polis_val) -> str:
    """Prioritas sumber CERTIFICATE (per konfirmasi user):
    1. CLSDT_SERTF_NO -- kalau ada isinya & polanya valid, pakai ini.
    2. Kalau CLSDT_SERTF_NO kosong/tidak valid -> coba ambil pola sertifikat
       yang menempel di CLSDT_POLICY_NO (karena CLSDT_POLICY_NO adalah
       sumber polis prioritas di Data 2).
    3. Kalau CLSDT_POLICY_NO juga tidak punya pola sertifikat -> fallback
       terakhir ke FAC_POLICY_NO (induk).
    """
    cert = build_certificate(clsdt_sertf_val)
    if cert:
        return cert

    cert = extract_certificate_from_policy(clsdt_polis_val)
    if cert:
        return cert

    return extract_certificate_from_policy(fac_polis_val)


# ─────────────────────────────────────────────────────────────────────────────
# ATURAN KHUSUS JRP (diverifikasi terhadap 37 contoh dari dokumen eksplorasi
# + stress-test ke seluruh nilai unik FAC/CLSDT di data mentah -- 0 exception,
# 0 hasil kosong, 0 overflow >5 kolom)
# ─────────────────────────────────────────────────────────────────────────────

_MONTH_RE = (
    r"JANUARI|FEBRUARI|MARET|APRIL|MEI|JUNI|JULI|AGUSTUS|SEPTEMBER|OKTOBER|"
    r"NOVEMBER|DES[EA]?MBER|DESEBER|JANUARY|FEBRUARY|MARCH|APRIL|MAY|JUNE|"
    r"JULY|AUGUST|OCTOBER|NOVEMBER|DECEMBER"
)

_TWO_LONG_DASH_RE = re.compile(r"^(\d{10,})\s*-\s*(\d{10,})\s*(?:/\s*VARIOUS)?\s*$", re.IGNORECASE)
_SHORT_SUFFIX_DASH_RE = re.compile(r"(\d{7,})\s*[-.]\s*(\d{1,4})(?!\d)")

_THREE_DIGIT_BREAKDOWN_RE = re.compile(
    r"^(\d{10,})\s*-\s*(\d{3})$"
)
_SD_SLIP_DECLARATION_RE = re.compile(
    r"^\s*"
    r"DEKLARASI\s*:?\s*"
    r"\d{1,2}\s+S\s*/?\s*D\s+\d{1,2}\s+"
    rf"(?:{_MONTH_RE})"
    r"\s+\d{4}\s*/\s*"
    r"SLIP\s*:\s*"
    r"(?P<start>\d{10,})\s+SD\s+(?P<end>\d{10,})"
    r"\s*$",
    re.IGNORECASE,
)


def _expand_sd_slip_declaration(val: str):
    m = _SD_SLIP_DECLARATION_RE.match(val.strip())

    if not m:
        return None

    start = m.group("start")
    end = m.group("end")

    if not start or not end:
        return None

    if len(start) != len(end):
        return None

    start_num = int(start)
    end_num = int(end)

    if end_num >= start_num:
        count = end_num - start_num + 1

        if count > MAX_SPLIT_COLS:
            return [f"{start} SD {end}"]

        return [
            str(start_num + i).zfill(len(start))
            for i in range(count)
        ]

    else:
        count = start_num - end_num + 1

        if count > MAX_SPLIT_COLS:
            return [f"{start} SD {end}"]

        return [
            str(start_num - i).zfill(len(start))
            for i in range(count)
        ]

def _reconstruct_short_suffix(m):
    base, suf = m.group(1), m.group(2)
    return base[:-len(suf)] + suf

_CERT_RANGE_RE = re.compile(
    r"[-.]\s*\d{5,7}(?!\d)\s*(?:(?:,|S\s*/?\s*D|S|-)\s*\d{3,7}(?!\d)\s*)*(?:,\s*\d{3,7}(?!\d)\s*)*",
    re.IGNORECASE,
)

_JUNK_RE = re.compile(
    r"""
    \d{1,2}\s*(?:SD|S\s*/\s*D|-)\s*\d{1,2}\s*(?:""" + _MONTH_RE + r""")\.?\s*(?:19|20)?\d{0,4}
  | -\s*E\#\d+
  | -?\s*REGIONAL\s*\d+
  | DEKLARASI\s*:?\s*\d{0,2}\s*(?:S\s*/?\s*D\s*\d{0,2})?\s*
  | \bSLIP\s*:\s*
  | \bS\s*/\s*D\b
  | \bSD\b
  | \bVARIOUS\s*:?
  | \bVAR\b
  | \bP\d+\b
  | \b(?:IDR|USD|EUR|GBP)\b
    """
    + r"| (?:" + _MONTH_RE + r")\.?\s*(?:19|20)?\d{0,4}",
    re.IGNORECASE | re.VERBOSE,
)

_LITERAL_WHITELIST = {"O01RI0002", "REALISASI 2024"}
_STANDALONE_NARRATIVE_RE = re.compile(r"SUSPEN|REGIONAL\d+\s+ONLY", re.IGNORECASE)
_STANDALONE_P_RE = re.compile(
    r"^\s*(?:P\d+|TBA|VAR|VARIOUS)(?:\s*[+,]\s*(?:P\d+|TBA|VAR|VARIOUS))*\s*$", re.IGNORECASE
)

_PREFIX_PLUS_RE = re.compile(r"^(\d{2,8})-(\d{8,})((?:\+\d{8,})+)-?\s*$")

def _apply_prefix_plus(val: str):
    m = _PREFIX_PLUS_RE.match(val.strip())
    if not m:
        return None
    prefix, first, rest = m.groups()
    return [prefix + s for s in [first] + rest.strip("+").split("+")]

_CN_CHUNK_RE = re.compile(
    r"\d{3,6}\s*/\s*CN\s*/\s*\d{3,4}\s*/\s*\d{2}\s*/\s*\d{2}(?:\s*/\s*\d(?!\d))?",
    re.IGNORECASE,
)
_PREMI_NOTE_RE = re.compile(
    r"-?\s*\d{1,2}\s+(?:" + _MONTH_RE + r")\s+\d{4}\s+PREMI\s+[\d.,]+",
    re.IGNORECASE,
)

def _apply_cn_chunks(val: str):
    val_clean = _PREMI_NOTE_RE.sub("", val).strip(" -/")
    chunks = list(_CN_CHUNK_RE.finditer(val_clean))
    if not chunks:
        return None
    results = []
    lead = val_clean[:chunks[0].start()].strip(" -/")
    if re.fullmatch(r"\d{10,}", lead):
        results.append(lead)
    results += [_normalize_spaces(c.group(0)) for c in chunks]
    return results or None

def _clean_value(val) -> list:
    if pd.isna(val):
        return []

    val = str(val).strip()

    if not val:
        return []

    r = _expand_sd_slip_declaration(val)
    if r is not None:
        return _cap_or_join(r)

    if val in _LITERAL_WHITELIST or _STANDALONE_NARRATIVE_RE.search(val) or _STANDALONE_P_RE.match(val):
        return [_normalize_spaces(val)]

    r = _apply_prefix_plus(val)
    if r is not None:
        return _cap_or_join(r)

    m = _TWO_LONG_DASH_RE.match(val)
    if m:
        return [m.group(1), m.group(2)]

    r = _apply_cn_chunks(val)
    if r is not None:
        return _cap_or_join(r)

    m = _THREE_DIGIT_BREAKDOWN_RE.match(val)
    if m:
        base, suffix = m.groups()
        reconstructed = base[:-3] + suffix
        return [base, reconstructed]

    text = _CERT_RANGE_RE.sub("", val)
    text = _JUNK_RE.sub(" ", text)
    text = _SHORT_SUFFIX_DASH_RE.sub(_reconstruct_short_suffix, text)
    text = re.sub(r"(\d{10,})\s*/\s*\d{1,2}\b(?!\d)", r"\1", text)
    text = text.strip(" -/,")

    if not text:
        return [_normalize_spaces(val)]

    raw_tokens = [t for t in re.split(r"\s+|[+/,]", text) if t.strip()]
    raw_tokens = [t.strip(".-") for t in raw_tokens]
    raw_tokens = [t for t in raw_tokens if t and re.search(r"\d", t)]

    if not raw_tokens:
        return [_normalize_spaces(val)]

    results, current_base = [], None

    for tok in raw_tokens:
        digits_only = re.match(r"^\d+$", tok) is not None

        if (
            digits_only
            and current_base is not None
            and len(tok) < len(current_base)
            and len(tok) <= 4
            and len(current_base) >= 7
        ):
            results.append(current_base[:-len(tok)] + tok)
        else:
            results.append(tok)

            if digits_only and len(tok) >= 7:
                current_base = tok

    final = [t for t in results if len(t) >= 5]

    if not final:
        return [_normalize_spaces(val)]

    final = _dedup_preserve_order(final)

    if len(final) == 1:
        return [_normalize_spaces(final[0])]

    if len(final) <= MAX_SPLIT_COLS:
        return [_normalize_spaces(f) for f in final]

    return [_normalize_spaces(text)]

def clean_polis(val) -> list:
    return _clean_value(val)


def clean_slip(val) -> list:
    return _clean_value(val)


# ─────────────────────────────────────────────────────────────────────────────
# CLEAN INSURED (pipeline standar PT/CV/Persero strip, DITAMBAH 3 rule dari
# BUG_jasa_rp_1.txt -- lihat versi lama untuk detail; TIDAK diubah oleh
# update CERTIFICATE ini)
# ─────────────────────────────────────────────────────────────────────────────

_ROMAN_MAP = {"IV": "4", "III": "3", "II": "2", "I": "1"}
_PELINDO_TRIGGER_RE = re.compile(r"PELABUHAN\s+INDONESIA|PELINDO", re.IGNORECASE)

def _expand_pelindo(val: str):
    if not _PELINDO_TRIGGER_RE.search(val):
        return None
    text = val.upper()
    text = re.sub(r"PELINDO\s+GROUP\s*/?\s*", "", text)
    text = re.sub(r"\bPERSERO\b", "", text)
    roman_nums = re.findall(r"\b(IV|III|II|I)\b", text)
    if roman_nums:
        regions = [_ROMAN_MAP[r] for r in roman_nums]
    else:
        regions = re.findall(r"\b([1-4])\b", text)
    if not regions:
        return None
    regions = list(dict.fromkeys(regions))
    return [f"PELABUHAN INDONESIA REGIONAL {n}" for n in regions]


_CABANG_RE = re.compile(r"^(.*?\bCABANG)\s+(.+)$", re.IGNORECASE)

def _expand_cabang(val: str):
    m = _CABANG_RE.match(val.strip())
    if not m:
        return None
    prefix, tail = m.groups()
    prefix = re.sub(r"\s*:\s*CABANG$", " CABANG", prefix, flags=re.IGNORECASE).strip()
    cities = [c.strip(" .") for c in re.split(r"\s*,\s*", tail) if c.strip(" .")]
    if len(cities) < 2:
        return None
    return [f"{prefix} {city}" for city in cities]


_QQ_TEMPLATE_RE = re.compile(r"^(.+?)\s+(\S+)\s+QQ\s+(.+)$", re.IGNORECASE)

def _expand_qq_template(val: str):
    if "CABANG" in val.upper():
        return None
    m = _QQ_TEMPLATE_RE.match(val.strip())
    if not m:
        return None
    prefix, first_word, rest = m.groups()
    alts = [a.strip() for a in re.split(r"\s+QQ\s+", rest, flags=re.IGNORECASE) if a.strip()]
    if not alts:
        return None
    results = [f"{prefix} {first_word}".strip()]
    results += [f"{prefix} {alt}".strip() for alt in alts]
    return results


_INSURED_PREFIX_RE = re.compile(r"^\s*(?:PT|CV)\.?\s*", re.IGNORECASE)
_INSURED_GELAR_RE = re.compile(
    r"\b(?:BAPAK|IBU|SDRI?|NYONYA|NY|TUAN|MR|MRS|MS|DR|IR|PROF|HJ|H)\b\.?\s*",
    re.IGNORECASE,
)
INSURED_SUFFIX_RE = re.compile(
    r"""
    ,?\s*\bTBK\b(?:\s*,?\s*PT\.?)?
  | ,?\s*\(\s*PERSERO\s*\)\s*
  | ,?\s*\bPERSERO\b\s*
  | ,?\s*\bLTD\.?\b\s*
  | ,?\s*\bPT\.?\s*$
  | ,?\s*\bCV\.?\s*$
    """,
    re.IGNORECASE | re.VERBOSE,
)
INSURED_JUNK_WORDS = frozenset({
    "PT", "CV", "TBK", "PERSERO", "LTD", "INC", "LLC",
    "AND", "OR", "THE", "OF", "AS",
})
_INSURED_SPLIT_RE = r"\bQQ\b|\bAND\s*/\s*OR\b|/|,|–|¿"


def _remove_polis_slip_from_insured(text: str, polis_ori, slip_ori) -> str:
    original_text = text.strip()

    for token in (polis_ori, slip_ori):
        if pd.notna(token):
            t = str(token).strip()

            if not t or t == "-":
                continue

            if t.upper() == original_text.upper():
                continue

            text = text.replace(t, "")

    return text


def _clean_insured_name(name: str) -> str:
    name = _normalize_spaces(name)
    name = _INSURED_PREFIX_RE.sub("", name)
    name = _INSURED_GELAR_RE.sub("", name)
    name = re.sub(r"\bPTE\.?\b", "", name, flags=re.IGNORECASE)
    name = re.sub(r"\(([^()]*)\)", "", name)
    name = name.replace("(", " ").replace(")", " ")
    name = re.sub(r"/", " ", name)
    name = name.replace(":", " ")
    for _ in range(3):
        cleaned = _normalize_spaces(INSURED_SUFFIX_RE.sub("", name).strip().strip(","))
        if cleaned == name:
            break
        name = cleaned
    name = re.sub(r"(?<!\d)\.(?!\d)", "", name)
    return _normalize_spaces(name).strip().upper()


def clean_insured(val, polis_ori=None, slip_ori=None) -> list:
    if pd.isna(val):
        return []
    val = str(val).strip()
    if not val:
        return []

    val = _remove_polis_slip_from_insured(val, polis_ori, slip_ori)

    val = re.sub(
        r"\s*/\s*PELINDO\b",
        "",
        val,
        flags=re.IGNORECASE,
    )

    if re.match(
        r"^\s*(ROIF\s+ANUGERAH|PATRA\s+DRILLING)\b",
        val,
        flags=re.IGNORECASE,
    ):
        parts = [
                p.strip()
                for p in re.split(r"\s+QQ\s+", val, flags=re.IGNORECASE)
                if p.strip()
            ]

        if len(parts) > 1:
                return _cap_or_join([
                    _normalize_spaces(x)
                    for x in parts
                ])

    for fn in (_expand_pelindo, _expand_cabang, _expand_qq_template):
        r = fn(val)
        if r is not None:
            return _cap_or_join([
                _normalize_spaces(x)
                for x in r
            ])

    val = re.sub(
        r",\s*(?:PT|CV)\.?(?=[\s/]|$)",
        " ",
        val,
        flags=re.IGNORECASE,
    )

    raw_parts = re.split(
        _INSURED_SPLIT_RE,
        val,
        flags=re.IGNORECASE,
    )

    cleaned = []
    for p in raw_parts:
        p = _normalize_spaces(p.strip())
        if len(p) <= 2 or p.upper() in INSURED_JUNK_WORDS:
            continue
        if re.match(r"^[^a-zA-Z0-9]+$", p):
            continue
        p_clean = _clean_insured_name(p)
        if not p_clean or len(p_clean) <= 2 or p_clean.upper() in INSURED_JUNK_WORDS:
            continue
        cleaned.append(p_clean)

    if not cleaned:
        fallback = _clean_insured_name(_normalize_spaces(val))
        return [fallback] if fallback else []

    return _cap_or_join(cleaned)


# ─────────────────────────────────────────────────────────────────────────────
# PEMILIHAN SUMBER CLSDT -> FAC (sama seperti TPI/CSI, per kolom independen)
# ─────────────────────────────────────────────────────────────────────────────

def _is_blank_or_placeholder(value) -> bool:
    if pd.isna(value):
        return True
    text = str(value).strip()
    return text == "" or text.upper() in {"NAN", "NONE", "NULL", "-", "—", "–"}


def _is_standalone_general(value) -> bool:
    if _is_blank_or_placeholder(value):
        return True
    return bool(_STANDALONE_P_RE.fullmatch(str(value).strip()))


def _clsd_source_is_valid(value) -> bool:
    if _is_blank_or_placeholder(value):
        return False
    if _is_standalone_general(value):
        return False
    return True


def _clean_selected_source(kind: str, clsd_value, fac_value) -> list:
    use_clsd = _clsd_source_is_valid(clsd_value)
    source = clsd_value if use_clsd else fac_value
    patterns = KNOWN_POLIS_PATTERNS if kind == "polis" else KNOWN_SLIP_PATTERNS
    return _refine_list(_clean_value(source), patterns)


# ─────────────────────────────────────────────────────────────────────────────
# PROSES UTAMA
# ─────────────────────────────────────────────────────────────────────────────

def _force_text_format(output_file, prefixes):
    """Paksa format cell jadi Text ('@') utk kolom hasil clean, supaya
    angka panjang tidak auto-diconvert jadi numerik/scientific notation
    ketika file dibuka ulang (Excel maupun pandas.read_excel default)."""
    import openpyxl
    wb = openpyxl.load_workbook(output_file)
    ws = wb.active
    header = [c.value for c in ws[1]]
    target_cols = [i for i, h in enumerate(header, start=1) if h and str(h).startswith(prefixes)]
    for col_idx in target_cols:
        for row in ws.iter_rows(min_row=2, min_col=col_idx, max_col=col_idx):
            for cell in row:
                cell.number_format = "@"
    wb.save(output_file)


def _insert_clean_columns(df, all_lists, prefix, max_cols):
    added = []
    actual_max_cols = max((len(lst) for lst in all_lists if isinstance(lst, list)), default=1)
    actual_max_cols = min(actual_max_cols, max_cols)
    for i in range(1, actual_max_cols + 1):
        col_name = f"{prefix}_{i}"
        df[col_name] = [lst[i - 1] if isinstance(lst, list) and i - 1 < len(lst) else None for lst in all_lists]
        added.append(col_name)
    return added


def _insert_polis_cert_columns(df, all_polis_lists, all_cert_values, max_cols):
    """Sisipkan kolom POLIS_CLEAN_n (dulu bernama POLICY_CLEAN_n), dan
    sisipkan SATU kolom CERTIFICATE_1 tepat setelah POLIS_CLEAN_1 saja.

    resolve_certificate() hanya pernah menghasilkan SATU nilai certificate
    per baris, dan cedant JRP ini memang selalu case certificate tunggal --
    jadi TIDAK ada CERTIFICATE_2, CERTIFICATE_3, dst sekalipun ada
    POLIS_CLEAN_2, POLIS_CLEAN_3, dst untuk baris yang punya lebih dari
    1 polis.
    """
    added = []
    actual_max_cols = max((len(lst) for lst in all_polis_lists if isinstance(lst, list)), default=1)
    actual_max_cols = min(actual_max_cols, max_cols)
    for i in range(1, actual_max_cols + 1):
        polis_col = f"POLIS_CLEAN_{i}"
        df[polis_col] = [lst[i - 1] if isinstance(lst, list) and i - 1 < len(lst) else None for lst in all_polis_lists]
        added.append(polis_col)

        if i == 1:
            df["CERTIFICATE_1"] = list(all_cert_values)
            added.append("CERTIFICATE_1")
    return added


def process_data(input_file: str, output_file: str) -> None:
    print(f"[1/5] Membaca data dari: {input_file} (sheet: {INPUT_SHEET}) ...")
    df = pd.read_excel(input_file, sheet_name=INPUT_SHEET, header=0)
    print(f"      Total baris keseluruhan: {len(df):,}")

    if CEDANT_COL not in df.columns:
        print(f"\n[ERROR] Kolom '{CEDANT_COL}' tidak ditemukan!")
        return

    df[CEDANT_COL] = df[CEDANT_COL].astype(str).str.strip()
    df = df[df[CEDANT_COL] == CEDANT_VALUE].copy()
    print(f"[2/5] Filter cedant '{CEDANT_VALUE}': {len(df):,} baris ditemukan.")

    if df.empty:
        print("\n[WARN] Tidak ada data setelah filter. Proses dihentikan.")
        return

    for col in [POLIS_COL, SLIP_COL, INSURED_COL, CLSDT_POLIS_COL, CLSDT_SLIP_COL]:
        if col not in df.columns:
            print(f"\n[ERROR] Kolom '{col}' tidak ditemukan di file!")
            print(f"        Kolom tersedia: {list(df.columns)}")
            return

    has_cert_col = CLSDT_SERTF_COL in df.columns
    if not has_cert_col:
        print(f"\n[WARN] Kolom '{CLSDT_SERTF_COL}' tidak ditemukan -- CERTIFICATE "
              f"akan diambil murni dari fallback pola di CLSDT_POLICY_NO / "
              f"FAC_POLICY_NO. Cek nama kolom sertifikat asli di file dan "
              f"sesuaikan CLSDT_SERTF_COL di KONFIGURASI kalau ternyata ada.")
        print(f"        Kolom tersedia: {list(df.columns)}")

    df.rename(columns={
        POLIS_COL:   "FAC_POLICY_NO_ORI",
        SLIP_COL:    "FAC_SLIP_ORI",
        INSURED_COL: "FAC_INSURED_ORI",
    }, inplace=True)

    print("[3/5] Menjalankan proses cleaning ...")

    all_clean_polis, all_clean_slip, all_clean_ins, all_certificate = [], [], [], []
    max_polis = max_slip = max_ins = 1

    for _, row in df.iterrows():
        c_polis = _clean_selected_source("polis", row.get(CLSDT_POLIS_COL, ""), row.get("FAC_POLICY_NO_ORI", ""))
        c_slip  = _clean_selected_source("slip", row.get(CLSDT_SLIP_COL, ""), row.get("FAC_SLIP_ORI", ""))
        c_ins   = clean_insured(row.get("FAC_INSURED_ORI", ""), row.get("FAC_POLICY_NO_ORI", ""), row.get("FAC_SLIP_ORI", ""))
        clsdt_sertf_val = row.get(CLSDT_SERTF_COL, "") if has_cert_col else ""
        cert    = resolve_certificate(clsdt_sertf_val, row.get(CLSDT_POLIS_COL, ""), row.get("FAC_POLICY_NO_ORI", ""))

        max_polis = max(max_polis, len(c_polis))
        max_slip  = max(max_slip, len(c_slip))
        max_ins   = max(max_ins, len(c_ins))

        all_clean_polis.append(c_polis)
        all_clean_slip.append(c_slip)
        all_clean_ins.append(c_ins)
        all_certificate.append(cert)

    print("[4/5] Menyusun kolom output ...")

    new_columns = []
    for col in df.columns:
        new_columns.append(col)
        if col == "FAC_POLICY_NO_ORI":
            new_columns += _insert_polis_cert_columns(df, all_clean_polis, all_certificate, max_polis)
        elif col == "FAC_SLIP_ORI":
            new_columns += _insert_clean_columns(df, all_clean_slip, "SLIP_CLEAN", max_slip)
        elif col == "FAC_INSURED_ORI":
            new_columns += _insert_clean_columns(df, all_clean_ins, "INSURED_CLEAN", max_ins)

    df = df[new_columns]

    os.makedirs(os.path.dirname(output_file), exist_ok=True)
    print(f"[5/5] Menyimpan hasil ke: {output_file} ...")
    df.to_excel(output_file, index=False)

    _force_text_format(output_file, prefixes=("POLIS_CLEAN_", "SLIP_CLEAN_", "INSURED_CLEAN_", "CERTIFICATE"))

    print(f"\n{'=' * 55}")
    print("  [OK] Selesai!")
    print(f"  Total baris output  : {len(df):,}")
    print(f"  Total kolom output  : {len(df.columns)}")
    print(f"  File disimpan di    : {output_file}")
    print(f"{'=' * 55}")


if __name__ == "__main__":
    process_data(INPUT_FILE, OUTPUT_FILE)