"""
Cleaning Data 1 - Facultative | Cedant: JASA R.P (PT ASURANSI JASA RAHARJA PUTRA)

CATATAN ARSITEKTUR:
- Mesin polis/slip diadaptasi dari Data 2 (Osbal), TAPI bukan salinan 1:1 --
  beberapa rule beda, diverifikasi lewat dokumen eksplorasi Data 1 (bukan
  tebakan). Lihat komentar tiap rule.
- Kolom CERTIFICATE_n, diisi dari nomor sertifikat yang ada di pola
  POLIS (base + "-"/"." + sertifikat). Nilai certificate SELALU diambil
  dari FAC_POLICY_NO ORI (bukan hasil clean). Kosong kalau polis tidak
  bersertifikat atau pola tidak cocok kriteria sertifikat.
- Insured: Pelindo/CABANG/singkatan/gelar/dll -- lihat komentar tiap rule.

=====================================================================
UPDATE STRUKTUR KOLOM OUTPUT (terbaru) -- HANYA STRUKTUR, LOGIC CLEANING
TIDAK BERUBAH:
- Sebelumnya kolom CERTIFICATE (satu kolom, nilai tunggal per baris)
  diletakkan SEBELUM semua kolom POLIS bersih (FAC_POLICY_NO_1, dst).
- Sekarang: kolom hasil clean polis diseragamkan namanya jadi
  POLIS_CLEAN_n (bukan lagi memakai nama kolom asal FAC_POLICY_NO_n),
  dan CERTIFICATE_n disisipkan TEPAT SETELAH POLIS_CLEAN_n yang
  bersangkutan:
      FAC_POLICY_NO -> POLIS_CLEAN_1 -> CERTIFICATE_1
                     -> POLIS_CLEAN_2 -> CERTIFICATE_2
                     -> ...
- Jumlah CERTIFICATE_n MENGIKUTI jumlah POLIS_CLEAN_n (bukan jumlah
  angka sertifikat di dalam satu nilai certificate -- satu certificate
  boleh berisi "000121 SD 000122" atau beberapa nomor sekaligus, tetap
  1 kolom CERTIFICATE_n).
- Karena extract_certificate() hanya pernah menghasilkan SATU nilai
  certificate per baris (diambil dari FAC_POLICY_NO ORI utuh, bukan per
  pecahan polis) dan cedant JRP ini memang hanya punya case certificate
  tunggal, maka HANYA CERTIFICATE_1 yang diisi nilai hasil ekstraksi;
  CERTIFICATE_2, CERTIFICATE_3, dst (kalau ada baris dengan >1 polis)
  selalu dikosongkan.
- CERTIFICATE_1 tetap selalu dibuat walau tidak ada certificate sama
  sekali (nilainya string kosong), tapi CERTIFICATE_2/3/dst TIDAK
  otomatis dibuat kalau memang tidak ada baris dengan polis sebanyak
  itu.
=====================================================================

BUG FIX LOG (BUG_data_1.docx) -- 24/09 kasus insured baru diverifikasi
lolos + 2 kasus POLIS/SLIP comma-list yang salah rekonstruksi:

1) INSURED - mesin PELINDO/PELABUHAN INDONESIA ditulis ULANG TOTAL.
   Versi lama cuma menangani SATU token "PELINDO <angka romawi>" dan
   gagal total untuk: list region campur romawi/angka (I, II, III, IV /
   1 2 3 4 / 1,2,3,4), pemisah "DAN", duplikat penyebutan region yang
   sama (harus di-dedup jadi 1), literal company name yang KEBETULAN
   diawali "PELINDO" tapi bukan indikasi region (mis. "PELINDO JASA
   MARITIM" -> singkatan "SPJM" di sebelahnya harus dibuang, BUKAN
   diperlakukan sebagai insured terpisah).
   Mesin baru: cari SEMUA pasangan (PELABUHAN INDONESIA [PERSERO]
   [REGIONAL] | PELINDO) + daftar romawi/angka apa pun pemisahnya
   (koma, &, /, DAN, spasi), kumpulkan semua region unik yang
   ditemukan di SELURUH teks, lalu breakdown jadi
   "PELABUHAN INDONESIA REGIONAL <n>" terurut -- ini otomatis
   menghasilkan dedup (region yang sama disebut 2x tetap 1 hasil) dan
   otomatis mengabaikan teks lain di sekitarnya (GROUP, ASET, &, dll).
   Kalau tidak ada region yang match sama sekali (mis. "PELINDO JASA
   MARITIM"), fallback ke split umum.

2) INSURED - split umum diperbaiki:
   - Separator QQ dan "/" dan "," dkk SEKARANG diproses BERSAMAAN dalam
     satu pass (dulu: kalau ada QQ, split HANYA by QQ, bagian lain yang
     masih ada "/" tidak ikut kepisah -- itu sebabnya
     "HIO JRP VARIOUS/PT CEPA JASA INDONESIA QQ PT GABUNGAN" dulu cuma
     jadi 2 bagian, seharusnya 3).
   - Kata "VARIOUS"/"TBA" yang menempel di AWAL/AKHIR satu entitas
     (bukan whole-value) sekarang dibuang juga (dulu hanya value yang
     PERSIS "VARIOUS" yang dibuang, jadi "HIO JRP VARIOUS" dulu lolos
     apa adanya).
   - Separator " - " (spasi-dash-spasi) ditambahkan sebagai pemisah
     entitas insured (mis. "PP - WIKA" -> 2 insured: PP, WIKA;
     "PP - ADHI KSO" -> PP, ADHI KSO). Dulu dash jenis ini tidak
     dianggap pemisah sama sekali.
   - Ambang batas panjang minimal nama diturunkan dari <=2 ke <2 supaya
     nama 2 karakter yang valid (mis. "PP") tidak ikut terbuang sebagai
     "junk".
   - Kamus singkatan dikenal (_KNOWN_INSURED_ABBREV) ditambahkan untuk
     kasus "SPJM" = singkatan "PELINDO JASA MARITIM" (bukan pola
     inisial huruf pertama biasa, jadi perlu whitelist eksplisit).

3) POLIS/SLIP (clean_polis/clean_slip, sama2 lewat _clean_value) -
   _apply_comma_list_dash SALAH merekonstruksi SEMUA item dalam comma
   list saat totalnya > MAX_SPLIT_COLS (>5). Seharusnya: kalau item
   >5, JANGAN direkonstruksi sama sekali (dibiarkan mentah apa adanya,
   cuma item PERTAMA yang digabung ke base), digabung jadi SATU string
   dengan koma. Contoh nyata dari bug report:
     "...23001203052100-126,128,130,...,14" (11 item, >5)
     -> SEHARUSNYA: "23001203052100126,128,130,...,14" (base digabung
        cuma ke item pertama, sisanya tetap mentah)
     -> BUG LAMA: semua 11 item direkonstruksi jadi nomor panjang penuh
        (salah, dan otomatis error tambahan: reconstruct pakai basis
        yg sudah 17 digit, bukan basis asli 14 digit -> hasil ambigu)
   Fix: cek dulu apakah item > MAX_SPLIT_COLS SEBELUM rekonstruksi --
   kalau ya, cukup gabung base+item[0] lalu sisanya (item[1:]) mentah,
   join koma jadi satu value. Kalau item <= MAX_SPLIT_COLS, baru
   direkonstruksi semua seperti versi lama (breakdown ke banyak kolom).

4) CERTIFICATE - mesin _normalize_cert_block DIGANTI TOTAL dengan
   build_certificate() (rule baru, sama dipakai di Data 2 & Data 3):
   - Kalau berupa RANGE 2-angka (dipisah "-" ATAU "S/D"/"SD"/"S", BUKAN
     koma): commit ke jumlah anggota range (hi-lo+1).
       jumlah > 3  -> notasi "AWAL SD AKHIR" (breakdown TIDAK dijalankan)
       jumlah <= 3 -> breakdown penuh, gabung koma "a, b, c"
   - Kalau berupa daftar eksplisit (dipisah koma, dan/atau campur
     S/D/SD/S) TANPA pola range 2-angka murni: dianggap daftar diskrit
     apa adanya (TIDAK diisi celah), tapi kalau jumlah item > 3 tetap
     dikompres ke notasi "ITEM_PERTAMA SD ITEM_TERAKHIR".
   - Semua angka di-pad ke 6 digit (zfill 6).
   - Semua varian penulisan "S/D", "s/d", "SD", "S" -> distandarisasi
     jadi "SD" di output.
   - Kalau ada token yang bukan angka murni (pola tidak masuk akal) ->
     CERTIFICATE dikosongkan (bukan dipaksakan).
   Bug lama: separator asli (S/D vs SD vs S vs -) dipertahankan APA
   ADANYA di output (tidak distandarisasi ke "SD"), dan tidak ada
   logika ambang batas jumlah (>3 pakai SD, <=3 breakdown) sama sekali
   -- semua item HANYA di-pad6 tanpa keputusan format berbasis count.
=====================================================================
"""

import os
import re

import pandas as pd


# ─────────────────────────────────────────────────────────────────────────────
# KONFIGURASI
# ─────────────────────────────────────────────────────────────────────────────

INPUT_FILE  = os.path.join("input", "1b. Transaksi Facul 01.01.23 - 17.08.2026.xlsx")
INPUT_SHEET = "Query result"
OUTPUT_FILE = os.path.join("output", "JasaRp_output_facul.xlsx")

CEDANT_COL  = "COMP_NAME"
CEDANT_VALUE = "PT.ASURANSI JASA RAHARJA PUTRA"

# Kolom broker: file terbaru punya 2 kolom "COMP_NAME" duplikat (satu utk
# cedant, satu utk broker) -- pandas otomatis rename yg kedua jadi
# "COMP_NAME.1". Kalau file yang dipakai versi lama (kolom "COMP_NAME2"),
# ganti manual di sini.
BROKER_COL  = "COMP_NAME.1"
POLIS_COL   = "FAC_POLICY_NO"
SLIP_COL    = "FAC_SLIP"
INSURED_COL = "FAC_INSURED"

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
# KOLOM CERTIFICATE (rule baru -- lihat BUG FIX LOG #4 di atas)
# ─────────────────────────────────────────────────────────────────────────────

_CERT_RANGE_PAIR_RE = re.compile(
    r"^\s*(\d{1,7})\s*(?:-|S\s*/\s*D|SD|S(?!\d))\s*(\d{1,7})\s*$",
    re.IGNORECASE,
)
_CERT_SPLIT_RE = re.compile(r"\s*(?:S\s*/\s*D|SD|S(?!\d)|,)\s*", re.IGNORECASE)
_CERT_TOKEN_RE = re.compile(r"^\d{1,7}$")


def build_certificate(cert_raw) -> str:
    """Format nomor sertifikat sesuai rule baru (dipakai sama di Data 1/2/3):
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


def extract_certificate(polis_raw) -> str:
    """Ambil blok sertifikat dari FAC_POLICY_NO_ORI (pola base+'-'/'.'
    +sertifikat), lalu format lewat build_certificate()."""
    if pd.isna(polis_raw):
        return ""
    val = str(polis_raw).strip()
    if not val:
        return ""

    m = re.match(r"^\s*(?:VARIOUS\s*/\s*)?(\d{10,})\s*[-.]\s*(.+)$", val, re.IGNORECASE)
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


# ─────────────────────────────────────────────────────────────────────────────
# MESIN POLIS/SLIP
# ─────────────────────────────────────────────────────────────────────────────

_MONTH_RE = (
    r"JANUARI|FEBRUARI|MARET|APRIL|MEI|JUNI|JULI|AGUSTUS|SEPTEMBER|OKTOBER|"
    r"NOVEMBER|DES[EA]?MBER|DESEBER|JANUARY|FEBRUARY|MARCH|APRIL|MAY|JUNE|"
    r"JULY|AUGUST|OCTOBER|NOVEMBER|DECEMBER"
)

_TWO_LONG_DASH_RE = re.compile(r"^(\d{10,})\s*-\s*(\d{10,})\s*(?:/\s*VARIOUS)?\s*$", re.IGNORECASE)
_SHORT_SUFFIX_DASH_RE = re.compile(r"(\d{7,})\s*[-.]\s*(\d{1,4})(?!\d)")

def _reconstruct_short_suffix(m):
    base, suf = m.group(1), m.group(2)
    return base[:-len(suf)] + suf

_COMMA_LIST_DASH_RE = re.compile(r"^(\d{7,})\s*-\s*(\d{1,4})\s*(?:,\s*\d{1,4}\s*)+$")

def _apply_comma_list_dash(val: str):
    m = _COMMA_LIST_DASH_RE.match(val.strip())
    if not m:
        return None
    base = m.group(1)
    items = re.findall(r"\d{1,4}", val.split("-", 1)[1])
    first_full = base + items[0]
    # BUG FIX #3: kalau item > MAX_SPLIT_COLS, JANGAN rekonstruksi semua --
    # cuma item pertama yang digabung ke base, sisanya tetap mentah, semua
    # digabung jadi SATU value dipisah koma.
    if len(items) > MAX_SPLIT_COLS:
        return [",".join([first_full] + items[1:])]
    results = [first_full]
    for it in items[1:]:
        results.append(first_full[:-len(it)] + it)
    return results

_CERT_RANGE_RE = re.compile(
    r"[-.]\s*\d{5,7}(?!\d)\s*(?:(?:,|S\s*/?\s*D|S|-)\s*\d{2,7}(?!\d)\s*)*(?:,\s*\d{2,7}(?!\d)\s*)*",
    re.IGNORECASE,
)

_SD_FULL_RANGE_RE = re.compile(r"(\d{10,})\s*(?:S\s*/\s*D|SD)\s*(\d{10,})", re.IGNORECASE)

def _apply_sd_full_range(text: str):
    m = _SD_FULL_RANGE_RE.search(text)
    if not m:
        return None
    base1, base2 = m.group(1), m.group(2)
    if len(base1) != len(base2):
        return None
    diff = abs(int(base2) - int(base1)) + 1
    if diff > 5:
        return [f"{base1} SD {base2}"]
    lo, hi = sorted([int(base1), int(base2)])
    width = len(base1)
    return [str(n).zfill(width) for n in range(lo, hi + 1)]

_DOT_CHAIN_DIGIT_RE = re.compile(r"^[\d.]+$")
_PAREN_RE = re.compile(r"\([^()]*\)")

_PROD_YEAR_RANGE_RE = re.compile(
    r"\bPROD\.?\s*(?:19|20)\d{2}\s*/\s*(?:19|20)?\d{2,4}\b", re.IGNORECASE
)

_JUNK_RE = re.compile(
    r"""
    \d{1,2}\s*(?:SD|S\s*/\s*D|-)\s*\d{1,2}\s*(?:""" + _MONTH_RE + r""")\.?\s*(?:19|20)?\d{0,4}
  | -\s*E\#\d+
  | -?\s*END\s*\d+
  | /\s*0\b
  | /\s*(?:19|20)\d{2}\b
  | -?\s*REGIONAL\s*\d+
  | DEKLARASI\s*:?\s*\d{0,2}\s*(?:S\s*/?\s*D\s*\d{0,2})?\s*
  | \bSLIP\s*:\s*
  | \bS\s*/\s*D\b
  | \bSD\b
  | \bVARIOUS\s*:?
  | \bVAR\b
  | \bP\d+\b
  | \bSEE\s+ATTACHMENT\b\s*/?
  | \bPROD\.?\s*
  | \bSP\b
  | \b(?:IDR|USD|EUR|GBP)\b
    """
    + r"| (?:" + _MONTH_RE + r")\.?\s*(?:19|20)?\d{0,4}",
    re.IGNORECASE | re.VERBOSE,
)

_CN_CHUNK_STRIP_RE = re.compile(
    r"/?\s*\d{3,6}\s*/\s*CN\s*/\s*\d{3,4}\s*/\s*\d{2}\s*/\s*\d{2}(?:\s*/\s*\d(?!\d))?",
    re.IGNORECASE,
)

_LITERAL_WHITELIST = {"O01RI0002", "REALISASI 2024", "AIR COMPRESSOR PDS 175 S/C 240"}
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

_TBA_PREFIX_RE = re.compile(r"^\s*TBA\s*/\s*", re.IGNORECASE)


def _clean_value(val) -> list:
    if pd.isna(val):
        return []
    val = str(val).strip()
    if not val:
        return []

    if val in _LITERAL_WHITELIST or _STANDALONE_P_RE.match(val):
        return [_normalize_spaces(val)]
    if _STANDALONE_NARRATIVE_RE.search(val):
        return [_normalize_spaces(val.replace("/", " "))]

    if "000351/CN/0200/03/24" in val and "024315/CN/0200/03/24/1" in val:
        return ["000351/CN/0200/03/24", "024315/CN/0200/03/24/1"]

    if _TBA_PREFIX_RE.match(val) and not _STANDALONE_P_RE.match(val):
        val = _TBA_PREFIX_RE.sub("", val).strip()

    r = _apply_prefix_plus(val)
    if r is not None:
        return _cap_or_join(r)

    m = _TWO_LONG_DASH_RE.match(val)
    if m:
        return [m.group(1), m.group(2)]

    if _DOT_CHAIN_DIGIT_RE.match(val) and val.count(".") >= 2:
        return [val.replace(".", "")]

    val = _PAREN_RE.sub(" ", val)
    val = _CN_CHUNK_STRIP_RE.sub("", val)
    val = _PROD_YEAR_RANGE_RE.sub("", val)

    sd_result = _apply_sd_full_range(val)
    if sd_result is not None:
        return _cap_or_join(sd_result)

    comma_result = _apply_comma_list_dash(_JUNK_RE.sub(" ", val).strip(" -/,"))
    if comma_result is not None:
        return _cap_or_join(comma_result)

    text = _CERT_RANGE_RE.sub("", val)
    text = _JUNK_RE.sub(" ", text)
    text = _SHORT_SUFFIX_DASH_RE.sub(_reconstruct_short_suffix, text)
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
        if (digits_only and current_base is not None and len(tok) < len(current_base)
                and len(tok) <= 4 and len(current_base) >= 7):
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
# CLEAN INSURED (mesin PELINDO ditulis ulang total -- lihat BUG FIX LOG #1/#2)
# ─────────────────────────────────────────────────────────────────────────────

_ROMAN_MAP = {"IV": "4", "III": "3", "II": "2", "I": "1"}

# Kamus singkatan yang dikenal secara eksplisit (bukan lewat pola inisial
# huruf pertama biasa) -- dipakai untuk membuang duplikat singkatan.
_KNOWN_INSURED_ABBREV = {
    "SPJM": "PELINDO JASA MARITIM",
    "ANTAM": "ANEKA TAMBANG",
}


def _initials_windows(name: str) -> set:
    words = [w for w in re.split(r"[\s\-]+", name.upper()) if w]
    variants = set()
    n = len(words)
    for start in range(n):
        for end in range(start + 1, n + 1):
            window = words[start:end]
            variants.add("".join(w[0] for w in window if w[0].isalpha()))
    return variants


def _is_abbreviation_of(short: str, long_name: str) -> bool:
    short_clean = re.sub(r"[^A-Z0-9]", "", short.upper())
    if not short_clean or len(short_clean) < 2 or len(short_clean) > 8:
        return False
    if short_clean in _initials_windows(long_name):
        return True
    if short_clean in _KNOWN_INSURED_ABBREV and _KNOWN_INSURED_ABBREV[short_clean] in long_name.upper():
        return True
    return False


# Cari SEMUA pasangan "(PELABUHAN INDONESIA [PERSERO] [REGIONAL] | PELINDO)
# <daftar romawi/angka>" di seluruh teks -- pemisah dalam daftar boleh koma,
# &, /, "DAN", atau spasi polos, campur sekalipun. Semua region yang
# ditemukan dikumpulkan sbg SET (otomatis dedup), lalu breakdown terurut.
# Kalau tidak ada region match sama sekali (mis. "PELINDO JASA MARITIM" --
# bukan indikasi region), fungsi return None dan fallback ke split umum.
_PELINDO_REGION_RE = re.compile(
    r"(?:PELABUHAN\s+INDONESIA(?:\s*\(?PERSERO\)?)?\s*(?:REGIONAL)?|PELINDO)\s+"
    r"((?:(?:IV|III|II|I|[1-4])\s*(?:,|&|/|DAN|\s)\s*)*(?:IV|III|II|I|[1-4]))",
    re.IGNORECASE,
)
_REGION_TOKEN_RE = re.compile(r"IV|III|II|I|[1-4]", re.IGNORECASE)


def _expand_pelindo(val: str):
    regions = set()
    for m in _PELINDO_REGION_RE.finditer(val):
        for tok in _REGION_TOKEN_RE.findall(m.group(1)):
            regions.add(int(_ROMAN_MAP.get(tok.upper(), tok)))
    if not regions:
        return None
    return [f"PELABUHAN INDONESIA REGIONAL {n}" for n in sorted(regions)]


_CABANG_RE = re.compile(r"^(.*?\bCABANG)\s+(.+)$", re.IGNORECASE)

def _expand_cabang(val: str):
    m = _CABANG_RE.match(val.strip())
    if not m:
        return None
    prefix, tail = m.groups()
    prefix = re.sub(r"\s*:\s*CABANG$", " CABANG", prefix, flags=re.IGNORECASE).strip()
    cities = [c.strip(" .") for c in re.split(r"\s*,\s*", tail) if c.strip(" .")]
    cities = [c for c in cities if c.upper() not in {"PT", "CV"}]
    if len(cities) < 2:
        return None
    return [f"{prefix} {city}" for city in cities]


_QQ_TEMPLATE_RE = re.compile(r"^(.+?)\s+(\S+)\s+QQ\s+(.+)$", re.IGNORECASE)

def _expand_qq_template(val: str):
    if "CABANG" in val.upper() or not re.match(r"^\s*KASTIP\b", val, re.IGNORECASE):
        return None
    m = _QQ_TEMPLATE_RE.match(val.strip())
    if not m:
        return None
    prefix, first_word, rest = m.groups()
    alts = [a.strip() for a in re.split(r"\s+QQ\s+", rest, flags=re.IGNORECASE) if a.strip()]
    if not alts:
        return None
    return [f"{prefix} {first_word}".strip()] + [f"{prefix} {alt}".strip() for alt in alts]


_REGION_WORDS = {"KALTIM", "KALSEL", "KALTENG", "KALBAR", "KALUT",
                 "SULSEL", "SULUT", "SULTENG", "SULBAR", "SULTRA",
                 "JATIM", "JATENG", "JABAR", "SUMUT", "SUMSEL",
                 "SUMBAR", "NTB", "NTT", "DKI"}

_JUNK_BOILERPLATE_RE = re.compile(
    r"^\s*(?:HIS|ITS|THEIR)\s+SUBSIDIAR(?:Y|IES)\b.*$|^\s*BELUM\s+RENEWAL\s*$",
    re.IGNORECASE,
)

_GELAR_PREFIX_RE = re.compile(
    r"^\s*(?:BAPAK|IBU|SDRI?|NYONYA|NY|TUAN|TN|MR|MRS|MS|DR|IR|PROF|HJ|H)\b\.?\s*",
    re.IGNORECASE,
)
_GELAR_SUFFIX_RE = re.compile(
    r"\s*,?\s*\b(?:S\.?E|S\.?H|S\.?SI|S\.?T|S\.?SOS|S\.?KOM|S\.?PD|"
    r"M\.?SI|M\.?M|M\.?H|M\.?KOM|SH|SE|TN|IR)\b\.?\s*$",
    re.IGNORECASE,
)

# BUG FIX #2: buang "VARIOUS"/"TBA" yang menempel di AWAL/AKHIR SATU entitas
# (bukan cuma value yang PERSIS "VARIOUS")
_JUNK_EDGE_WORDS_RE = re.compile(
    r"^\s*(?:VARIOUS|TBA)\b\s*|\s*\b(?:VARIOUS|TBA)\b\s*$", re.IGNORECASE
)

_LTD_SUFFIX_RE = re.compile(r"\bLTD\b\.?\s*$", re.IGNORECASE)
_AMP_SPACED_RE = re.compile(r"\s+&\s+")

INSURED_JUNK_WORDS = frozenset({
    "PT", "CV", "TBK", "PERSERO", "LTD", "INC", "LLC",
    "AND", "OR", "THE", "OF", "AS", "VARIOUS",
})


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
    for _ in range(2):
        stripped = _JUNK_EDGE_WORDS_RE.sub("", name).strip()
        if stripped == name:
            break
        name = stripped
    name = _GELAR_PREFIX_RE.sub("", name)
    for _ in range(3):
        stripped = _GELAR_SUFFIX_RE.sub("", name).strip()
        if stripped == name:
            break
        name = stripped
    name = re.sub(r"\([^()]*\)", " ", name)
    name = _LTD_SUFFIX_RE.sub("", name)
    name = _AMP_SPACED_RE.sub(" ", name)
    name = re.sub(r",?\s*\(PERSERO\)\s*", " ", name, flags=re.IGNORECASE)
    name = re.sub(r"\bTBK\b\s*,\s*PT\.?\s*", " ", name, flags=re.IGNORECASE)
    name = re.sub(r",?\s*\bTBK\b\s*", " ", name, flags=re.IGNORECASE)
    name = re.sub(r",\s*PT\.?\s*$", "", name, flags=re.IGNORECASE)
    name = re.sub(r",\s*CV\.?\s*$", "", name, flags=re.IGNORECASE)
    name = re.sub(r"^\s*PT\.?\s+", "", name, flags=re.IGNORECASE)
    name = re.sub(r"^\s*CV\.?\s+", "", name, flags=re.IGNORECASE)
    name = re.sub(r"\bPT\.?\s*$", "", name, flags=re.IGNORECASE)
    name = re.sub(r"\bCV\.?\s*$", "", name, flags=re.IGNORECASE)
    name = re.sub(r"(?<!\d)\.(?!\d)", " ", name)
    return _normalize_spaces(name).strip().upper()


_LITERAL_WHITELIST_INSURED = {
    "AIR COMPRESSOR PDS 175 S/C 240",
    "INDORAMA SYNTHETICS/MEDISAFE/TECHNOLOGIES/TIGADAYA/INDRADHANUSA INDONESIA/INDO RAMA",
}

# BUG FIX #2: QQ, AND/OR, "/", ",", "–", "¿", DAN " - " (spasi-dash-spasi)
# SEKARANG semua diproses BERSAMAAN dalam satu regex split (dulu: kalau ada
# QQ, split HANYA by QQ, "/" di bagian lain tidak ikut kepisah).
_SPLIT_RE = re.compile(r"\bQQ\b|\bAND\s*/\s*OR\b|/|,|–|¿|\s-\s", re.IGNORECASE)


def clean_insured(val, polis_ori=None, slip_ori=None) -> list:
    if pd.isna(val):
        return []
    val = str(val).strip()
    if not val:
        return []
    if val.upper() in _LITERAL_WHITELIST_INSURED:
        return [val]
    if _JUNK_BOILERPLATE_RE.match(val):
        return []

    val = _remove_polis_slip_from_insured(val, polis_ori, slip_ori)

    for fn in (_expand_pelindo_cabang, _expand_pelindo, _expand_cabang, _expand_qq_template):
        r = fn(val)
        if r is not None:
            return _cap_or_join([_normalize_spaces(x) for x in r])

    val = re.sub(r"-\s*PT\s+", "  /  ", val, flags=re.IGNORECASE)
    val = re.sub(r"\bO[QO]\b", "QQ", val, flags=re.IGNORECASE)
    val = re.sub(r",\s*(?:PT|CV)\.?\b", "", val, flags=re.IGNORECASE)

    parts = [p.strip() for p in _SPLIT_RE.split(val) if p.strip()]

    merged = []
    for p in parts:
        p_stripped = p.strip()
        if p_stripped.upper() in _REGION_WORDS and merged:
            merged[-1] = merged[-1] + " " + p_stripped
        elif re.fullmatch(r"KSO\.?", p_stripped, flags=re.IGNORECASE) and merged:
            merged[-1] = merged[-1] + " KSO"
        else:
            merged.append(p_stripped)

    cleaned = []
    for p in merged:
        if _JUNK_BOILERPLATE_RE.match(p):
            continue
        p_clean = _clean_insured_name(p)
        # BUG FIX #2: ambang batas panjang diturunkan dari <=2 ke <2 supaya
        # nama pendek yang valid (mis. "PP") tidak ikut terbuang.
        if not p_clean or len(p_clean) < 2 or p_clean in INSURED_JUNK_WORDS:
            continue
        # BUG FIX #1: buang duplikat kalau segmen ini singkatan dikenal dari
        # entitas yang sudah ada (mis. "SPJM" utk "PELINDO JASA MARITIM")
        if p_clean in _KNOWN_INSURED_ABBREV:
            full = _KNOWN_INSURED_ABBREV[p_clean]
            if any(full in prior or prior in full for prior in cleaned):
                continue
        cleaned.append(p_clean)

    if not cleaned:
        fallback = _clean_insured_name(_normalize_spaces(val))
        return [fallback] if fallback else []

    final = []
    for name in cleaned:
        words = name.split()
        if len(words) >= 2 and any(_is_abbreviation_of(words[-1], prior) for prior in final):
            name = " ".join(words[:-1])
        if any(_is_abbreviation_of(name, prior) for prior in final):
            continue
        final.append(name)

    return _cap_or_join(final)


def _expand_pelindo_cabang(val: str):
    """
    Khusus PELINDO/PELABUHAN INDONESIA yang memiliki CAB atau CABANG.

    Contoh:
    PELINDO IV PERSERO CABANG BALIKPAPAN, PT
        -> PELABUHAN INDONESIA REGIONAL 4 CABANG BALIKPAPAN

    PELINDO IV (PERSERO) CAB SAMARINDA
        -> PELABUHAN INDONESIA REGIONAL 4 CAB SAMARINDA

    Rule ini hanya aktif jika:
    - ada PELINDO / PELABUHAN INDONESIA
    - ada nomor regional I-IV / 1-4
    - ada CAB / CABANG
    """

    text = val.strip()

    # Harus merupakan PELINDO / PELABUHAN INDONESIA
    if not re.search(
        r"\b(?:PELINDO|PELABUHAN\s+INDONESIA)\b",
        text,
        re.IGNORECASE,
    ):
        return None

    # Cari regional
    m_region = re.search(
        r"\b(?:PELINDO|PELABUHAN\s+INDONESIA)"
        r"(?:\s*\(?\s*PERSERO\s*\)?)?"
        r"\s*(?:REGIONAL\s*)?"
        r"(IV|III|II|I|[1-4])\b",
        text,
        re.IGNORECASE,
    )

    if not m_region:
        return None

    region_token = m_region.group(1).upper()
    region = _ROMAN_MAP.get(region_token, region_token)

    # Cari CAB / CABANG dan ambil semua setelahnya
    m_cab = re.search(
        r"\b(CABANG|CAB)\b\s+(.+)",
        text,
        re.IGNORECASE,
    )

    if not m_cab:
        return None

    cab_type = m_cab.group(1).upper()
    cab_name = m_cab.group(2).strip()

    # Buang PT / CV di akhir
    cab_name = re.sub(
        r",?\s*\b(?:PT|CV)\.?\s*$",
        "",
        cab_name,
        flags=re.IGNORECASE,
    ).strip(" ,.")

    if not cab_name:
        return None

    # Pertahankan kata CABANG atau CAB sesuai data asal
    return [
        f"PELABUHAN INDONESIA REGIONAL {region} "
        f"{cab_type} {cab_name}"
    ]


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

def _force_text_format(output_file, prefixes):
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
    """Sisipkan kolom POLIS_CLEAN_n, dan sisipkan SATU kolom CERTIFICATE_1
    tepat setelah POLIS_CLEAN_1 saja.

    extract_certificate() hanya pernah menghasilkan SATU nilai certificate
    per baris (dari FAC_POLICY_NO ORI utuh, bukan per pecahan polis), dan
    cedant JRP ini memang selalu case certificate tunggal -- jadi TIDAK ada
    CERTIFICATE_2, CERTIFICATE_3, dst sekalipun ada POLIS_CLEAN_2,
    POLIS_CLEAN_3, dst untuk baris yang punya lebih dari 1 polis.
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

    new_columns = []
    for col in df.columns:
        new_columns.append(col)
        if col == POLIS_COL:
            new_columns += _insert_polis_cert_columns(df, all_clean_polis, all_certificate, max_polis)
        elif col == SLIP_COL:
            new_columns += _insert_clean_columns(df, all_clean_slip, SLIP_COL, max_slip)
        elif col == "FAC_INSURED_ORI":
            new_columns += _insert_clean_columns(df, all_clean_ins, "FAC_INSURED", max_ins)

    df = df[new_columns]

    os.makedirs(os.path.dirname(output_file), exist_ok=True)
    print(f"[5/5] Menyimpan hasil ke: {output_file} ...")
    df.to_excel(output_file, index=False)
    _force_text_format(output_file, prefixes=("POLIS_CLEAN_", f"{SLIP_COL}_", "FAC_INSURED_", "CERTIFICATE"))

    print(f"\n{'=' * 55}")
    print("  [OK] Selesai!")
    print(f"  Total baris output  : {len(df):,}")
    print(f"  Total kolom output  : {len(df.columns)}")
    print(f"  File disimpan di    : {output_file}")
    print(f"{'=' * 55}")


if __name__ == "__main__":
    process_data(INPUT_FILE, OUTPUT_FILE)