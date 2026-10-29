"""
Cleaning Data 3 - Suspend | Cedant: KBRU (PT. KBRU REINSURANCE BROKERS
(EX. SIMAS))

CATATAN VERSI: v5 -- rewrite setelah bug report ke-4 (15/09/26, file:
KBRU_BUG_150926.md). Perubahan dari v4:

  POLIS/SLIP:
  1. [BARU] Narasi/deskripsi MURNI yang nyasar di kolom polis/slip --
     (a) mengandung SUSPEN/UTANG PIUTANG (mis. "PENYELESAIAN HUTANG
     PIUTANG KBRU 2021 USD", "SIMASRB - PENYELESAIAN SUSPENSE 2018 -
     2019 VARIOUS CEDING"), atau (b) narasi murni BORDERO/SUMMARY SIMAS
     RE/nama bulan/tahun tanpa angka lain (mis. "BORDERO NOVEMBER 2017",
     "MEI - JUNI 2018") -> DIBIARKAN APA ADANYA, spasi TIDAK dihapus.
  2. [BARU] "TBA/" (placeholder + garis miring doang) -> garis miringnya
     dibuang, hasil cukup "TBA".
  3. [FIX] "No." SEKARANG dibuang di MANA PUN posisinya (sebelumnya cuma
     di AWAL value) -- mis. "0061-00/SFGI/...-0001/No.00087/S/06/2018/AB"
     -> "006100SFGI0006291SRB062018000100087S062018AB".
  4. [FIX] Nilai yang HANYA "CN : <nomor>" (tanpa nilai lain) -> cuma
     label "CN :"-nya yg dibuang, NOMORNYA dipertahankan sbg hasil
     (sebelumnya seluruhnya ikut terbuang jadi kosong). Notasi CN yang
     BERSANDINGAN dgn nilai lain tetap dibuang total spt v3/v4.
  5. [FIX] Pemangkasan ekor "/segmen-tanpa-digit" hanya dipakai kalau
     value MEMANG punya rantai '+' -- supaya ekor sah spt "/AB", "/BB",
     "/AC" (kode 2 huruf) pada value tunggal tidak ikut kebuang.
  6. [BARU] Dash NYASAR di dalam pola F/S 2-segmen (mis.
     "11055129F00000-275") dilebur jadi "11055129F00000275" sebelum
     diproses.
  7. [FIX] Target substitusi 'pengulangan' saat ada BEBERAPA segmen digit
     berpanjang SAMA -> pilih yg pertama yang BUKAN nol semua (mis.
     "0000-5747": target "5747", bukan "0000").
  8. [BARU] Ambang 'berstruktur' utk substitusi diperluas: >=1 pemisah
     (./-) DAN total digit >=7 juga memenuhi (mis. "0000-5747"), selain
     3 syarat lama.
  9. [BARU] Token substitusi yang SANGAT pendek (<=2 digit, mis.
     "+97+07+43") -> pola pengulangan dianggap tidak jelas, seluruh
     nilai dibiarkan apa adanya.
 10. [BARU] Token tambahan tanpa huruf tapi ada "/" atau "." nempel
     (mis. "04820/0001189") -> diambil potongan PERTAMA saja sbg
     kandidat substitusi; sisanya dibuang (KBRU tidak punya kolom
     sertifikat).
 11. [BARU] Suffix "PT"/"CV" ikut dibuang dari hasil yg berupa nama
     perusahaan (mis. "TBA / ASURANSI FPG INSURANCE, PT" -> "ASURANSI
     FPG INSURANCE").

  INSURED:
 12. [BARU] "PELINDO" dikenali sbg SINONIM "PELABUHAN INDONESIA", dan
     penyebutan berulang dgn pemisah apa pun (mis. "PELABUHAN INDONESIA
     I QQ PELINDO I QQ PELINDO II QQ ...") dikumpulkan + dinormalisasi
     ke "PELABUHAN INDONESIA <arab>" + dedup.
 13. [FIX] Pola inisial huruf-tunggal-bertitik (mis. "S.A.R.I") -> titik
     dibuang TANPA jadi spasi ("SARI"), bukan "S A R I".
 14. [BARU] Kata kunci ekspansi ditambah UNIT & MALL (mis. "DELTA DUNIA
     TEKSTIL/UNIT IV/...", "GALAXY MALL I & II").
 15. [FIX] Ekspansi keyword+daftar sekarang juga dicoba PER BAGIAN hasil
     split (bukan cuma whole-value), supaya entitas lain di depan tidak
     ikut kebawa jadi prefix (mis. "SINAR GALAXY, PT / GALAXY MALL I &
     II" -> 3 insured, bukan 2 dgn prefix salah).
 16. [BARU] "QQ" yang nempel langsung ke huruf lain tanpa spasi (mis.
     "QQBALI", "INDAHQQ") -> tanda data rusak/ambigu, SELURUH value
     dibiarkan apa adanya. (Rantai Q murni spt "QQQ" TIDAK dihitung
     rusak -- itu sudah ditangani sbg varian pemisah QQ.)

CATATAN VERSI (v4) -- bug report ke-3 (14/09/26, file:
KBRU_BUG_140926.md). Perubahan dari v3:
  1. [FIX INSURED, semua data] "SINAR MAS"/"SINARMAS" dikonsistenkan jadi
     "SINARMAS" (tanpa spasi) -- berlaku di kolom INSURED maupun di
     hasil POLIS/SLIP yang kebetulan berisi nama perusahaan (lihat poin
     2 & 3).
  2. [FIX POLIS/SLIP] Hasil yang ternyata berupa TEKS/NAMA (tanpa digit
     sama sekali, muncul lewat mekanisme buang TBA/P<n> yg bersandingan,
     mis. "TBA/MERITZ KORINDO INSURANCE") SEBELUMNYA ikut kena "hapus
     karakter" penuh (spasi antar kata ikut hilang, jadi
     "MERITZKORINDOINSURANCE") -- SEKARANG utk hasil semacam ini spasi
     DIPERTAHANKAN (cuma dirapikan + SINAR MAS->SINARMAS), BUKAN
     dihapus semua.
  3. [BARU] Rantai P<n> yang dipisah SPASI (bukan "+", mis. "SINARMAS P2
     P3 P4 P5 P6 P7") sekarang ikut dibuang (sebelumnya cuma pola
     berpisah "+" yg dikenali). Rantai P<n> yg diselingi anotasi
     "(ADD n)" (mis. "P2 (ADD 31) + P3 (ADD 32) + ...") juga sekarang
     dikenali & dibuang bersama anotasinya.
  4. [BARU] Segmen depan (sebelum "/" pertama) berisi RANTAI angka
     pendek digabung "-" (bukan "+") sebanyak >=5 kali -> dianggap
     'pengulangan berlebih', DIBIARKAN APA ADANYA (sama spt aturan >5
     via "+" yg sudah ada) -- sebelumnya pola dash-chain ini tidak
     terdeteksi (cuma pola "+"-chain yg dicek).
  5. [FIX] Target substitusi 'pengulangan' utk pola PERSIS "ANGKA+1-2
     HURUF+ANGKA" (mis. "F00000473" pada "110551216F00000473", tanpa
     karakter lain) SEKARANG benar menyasar SEGMEN TERAKHIR (setelah
     huruf) -- sebelumnya salah menyasar segmen PERTAMA krn kebetulan
     lebih panjang (heuristik "segmen terpanjang" cocok utk pola FPAR-
     dot-style tapi TIDAK utk pola F/S 2-segmen ini).
  (Perubahan khusus Osbal terkait file ini ada di CATATAN ARSITEKTUR
  KBRU_cleaning_osbal.py poin seleksi sumber SLIP_CLEAN.)

CATATAN VERSI (v3) -- rewrite setelah bug report ke-2 (10/09/26, file:
filter_slip_clean_1_yang_masih_belu.md, Filter_clsdt_polis_yang_sesuai_
kara.md, filter_polis_clean_yang_sesuai_kara.md, KBRU_BUG.md). File ini
sendiri (Suspend) TIDAK kena bug seleksi sumber CLSDT/FAC (itu khusus
Osbal), TAPI mesin cleaning karakter POLIS/SLIP & INSURED yang dipakai
DIROMBAK CUKUP BESAR, jadi versi ini ikut naik ke v3 supaya konsisten
dgn 2 file lain. Perubahan dari v2:

  MESIN POLIS/SLIP:
  1. [FIX] "TBA"/"P<n>" yang BERSANDINGAN dgn konten lain (mis. "TBA/
     SINARMAS", "SINARMAS/TBA") -- TBA/P<n>-nya DIBUANG, sisanya (SINARMAS)
     yang dipakai sbg nilai (SEBELUMNYA di v2 seluruh nilai malah
     dikembalikan apa adanya/"abaikan" -- ini salah). TBA/P<n> yang
     BERDIRI SENDIRI (tanpa konten lain) tetap "abaikan" apa adanya.
  2. [FIX] Notasi "CN : <nomor>" sekarang DIBUANG TOTAL bersama nilainya
     (bukan lagi dijadikan nilai tambahan di kolom dinamis spt v2).
  3. [FIX] 'Pengulangan' digit sekarang bisa mendeteksi suffix
     berstruktur yang nempel di TOKEN MANAPUN dalam rantai '+' (v2 cuma
     mendukung suffix di token PERTAMA), mis. "23018+23019+23023+23022/
     FIAR/001.0000344-SRB/06/2022" -> 4 nilai, suffix "/FIAR/...2022"
     diterapkan ke semua basis angka (23018/23019/23023/23022).
  4. [BARU] Substitusi 'pengulangan' HANYA jalan kalau nilai utama cukup
     'berstruktur' (ada huruf, ATAU >=2 pemisah '.'//', ATAU total
     digitnya panjang >=10) -- angka pendek polos yg digabung '+' tanpa
     struktur (mis. "3522 + 3846", "0011477 + 0012959") TETAP dibiarkan
     apa adanya (ambigu, BUKAN disubstitusi jadi 2 nilai terpisah).
  5. [FIX] Extra token yg kena tempelan tanda baca (mis. "9877." dgn
     titik nyempil) sekarang benar dikenali sbg token digit polos,
     bukan salah dianggap "nilai lengkap" krn ada titiknya.

  INSURED: lihat CATATAN ARSITEKTUR di bawah utk daftar lengkap
  perubahan (pola PELABUHAN INDONESIA/CLUSTER/TOWER/REGIONAL/DISTRICT,
  suffix UD/PD/PT/CV/gelar, delimiter QQQ/OO/*, dst).

CATATAN ARSITEKTUR:
- KBRU adalah BROKER, nomor polis tidak selalu ada & formatnya
  tergantung cedant aslinya. TIDAK ADA penukaran isi/posisi kolom POLIS
  & SLIP.
- Mesin cleaning POLIS & SLIP (clean_value(), 1 fungsi utk KEDUA kolom):
    1. Placeholder murni BERDIRI SENDIRI (TBA, P<n>, VARIOUS/VAR,
       SELISIH, END, "-", "&"-joined 2 angka polos) -> ABAIKAN = teks
       ORI PERSIS.
    2. Pola khusus "NNNNNNNN[A-Z]NNNNNNNN-NN" (mis. "11055124F00000144-
       03") -> ABAIKAN juga (dash dipertahankan), di mana pun ditemukan.
    3. TBA/P<n> yang bersandingan dgn konten lain -> dibuang, sisanya
       yang diproses (baik itu angka asli MAUPUN nama tanpa digit, mis.
       "SINARMAS").
    4. Notasi "CN : <nomor>" -> dibuang total.
    5. "Pengulangan" digit: token digit pendek (di rantai '+') meng-
       gantikan N digit terakhir dari SEGMEN DIGIT TERPANJANG pada
       bagian yang paling berstruktur (bisa di posisi mana pun dalam
       rantai '+', termasuk kalau suffix-nya nempel di token terakhir).
       Hanya jalan kalau bagian utama cukup berstruktur; kalau tidak
       (angka pendek polos) -> abaikan (seluruh rantai '+' dibiarkan
       apa adanya).
    6. TERAKHIR ("hapus karakter"): tiap hasil dibuang SEMUA karakter
       selain huruf/angka -- KECUALI kalau hasilnya sendiri cocok pola
       khusus poin 2 (dash dipertahankan).
- INSURED -- delimiter: QQ/QQQ (typo "OO" juga dikenali), AND/OR, "/",
  "//", ",", en-dash "–", "¿", "*", dan "&" HANYA kalau diikuti PT/CV
  (mis. "... & PT OKI PULP..."). "&" polos (mis. "PULP & PAPER") TETAP
  bukan pemisah. "+" dan " - " (spasi-dash-spasi) BUKAN pemisah (supaya
  "SK+SUMATERA" & "PULAU - PULAU MEDIA" tidak salah pecah).
  * ", PT"/", CV" (di mana pun, bukan cuma di akhir) dibuang total
    bersama komanya (cegah split palsu kalau ada isi lain menyertai).
  * ", UD"/", PD" (jenis badan usaha) TIDAK memicu split -- koma
    dibuang, kode dipertahankan menyatu (mis. "GAJAH MAS, UD" -> "GAJAH
    MAS UD").
  * Gelar/sapaan SUFFIX (BP, MR, MRS, NY, TN, DRS, S.E, S.H, dst) yg
    nempel via koma di belakang nama -> dibuang.
  * Pola khusus Pelabuhan Indonesia/Pelindo: "PELABUHAN INDONESIA
    <daftar romawi/angka>" -> pecah per nomor (romawi dikonversi ke
    angka arab).
  * Pola "PREFIX [/] KEYWORD tok1 tok2 ..." (KEYWORD = FACTORY/CLUSTER/
    TOWER/REGIONAL/DISTRICT, daftar >=2 token dipisah koma/&/DAN/spasi)
    -> pecah per token jadi "PREFIX KEYWORD tok". Kalau cuma 1 token
    (mis. "... / DISTRICT 8 / ...") -> TIDAK dipecah jadi entitas baru,
    malah DIGABUNG ke entitas SEBELUMNYA ("... DISTRICT 8").
  * Isi dalam kurung: singkatan dari teks sebelumnya / kata boilerplate
    peran (PERSERO/PRINCIPAL/OWNER/INSURED/CONTRACTOR) / mengandung
    PT-CV -> DIBUANG total. Isi kurung mengandung KOMA (daftar
    deskriptif, mis. "(REFINERY, PLANTATION, KEMITRAAN)") -> kurung
    DIPERTAHANKAN UTUH literal. Isi kurung lain (identitas cabang/pihak
    lain, mis. "(KSO)") -> kurung dibuang, isi dipertahankan.
  * Pola "(AS (THE) OWNER/PRINCIPAL OF <NAMA>)" -> <NAMA> jadi insured
    TERPISAH.
  * Data rusak (kata terduplikasi berturut-turut, mis. "WASKITA WASKITA
    KARYA KARYA", atau kurung tak seimbang) dinormalisasi dulu di awal.
  * Baris SUSPEN(SE)/(H) UTANG PIUTANG / deskripsi treaty-struktur
    (TREATY/STRUCTURE/EXCESS OF LOSS/WHOLE ACCOUNT) -> dibiarkan apa
    adanya (tidak diproses/di-split sama sekali).
  * TIDAK ADA lagi langkah dedup-abbreviation (beda dari KB INS) --
    singkatan yg muncul via delimiter (mis. "SINARMAS LAND GROUP / SML")
    tetap di-breakdown jadi entitas terpisah.
- Filter baris: CEDANT SHRT NAME == "KBRU", STATUS == "SUSPENSE".
- CERTIFICATE_1 SELALU KOSONG (cedant ini tidak punya konsep sertifikat),
  posisinya TEPAT setelah POLIS_CLEAN_1.
"""

import os
import re

import pandas as pd


# ─────────────────────────────────────────────────────────────────────────────
# KONFIGURASI
# ─────────────────────────────────────────────────────────────────────────────

INPUT_FILE  = os.path.join("input", "3b. Database Suspense 150826.xlsx")
INPUT_SHEET = "Sheet1"
HEADER_ROW  = 2

OUTPUT_FILE = os.path.join("output", "Kbru_output_suspend.xlsx")

CEDANT_FILTER_COL   = "CEDANT SHRT NAME"
CEDANT_FILTER_VALUE = "KBRU"

STATUS_COL          = "STATUS"
STATUS_KEEP_VALUE   = "SUSPENSE"

INSURED_COL = "INSURED"
POLIS_COL   = "POLIS"
SLIP_COL    = "SLIP NO"

MAX_SPLIT_COLS = 5


# ─────────────────────────────────────────────────────────────────────────────
# HELPER UMUM
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
        return [", ".join(items)]
    return items


# ─────────────────────────────────────────────────────────────────────────────
# MESIN POLIS/SLIP KBRU
# ─────────────────────────────────────────────────────────────────────────────

_JUNK_STANDALONE_WORDS = {"TBA", "VAR", "VARIOUS", "SELISIH", "END", "-", "NAN", "NONE", "NULL"}
_P_CODE_RE = re.compile(r"^P\d+$", re.IGNORECASE)

_MONTH_RE = (
    r"JANUARI|FEBRUARI|MARET|APRIL|MEI|JUNI|JULI|AGUSTUS|SEPTEMBER|OKTOBER|"
    r"NOVEMBER|DES[EA]?MBER|JANUARY|FEBRUARY|MARCH|MAY|JUNE|JULY|AUGUST|OCTOBER|DECEMBER"
)
_MONTH_YEAR_TAIL_RE = re.compile(
    r"\s*/?\s*(?:" + _MONTH_RE + r")\.?\s*(?:19|20)?\d{0,4}\s*$", re.IGNORECASE
)

_CN_MARKER_STRIP_RE_UNUSED = None
_NO_ANYWHERE_RE = re.compile(r"\bNO\.\s*", re.IGNORECASE)
_LEADING_ALPHA_PLUS_RE = re.compile(r"^[^\d/]{2,}?\+\s*")
_TRAILING_JUNK_WORD_RE = re.compile(r"[\s/]+(?:TBA|VAR|VARIOUS)\s*$", re.IGNORECASE)

# [BARU 15/09] deskripsi murni SUSPENSE/UTANG PIUTANG yang nyasar di
# kolom polis/slip -- dibiarkan apa adanya (spasi TIDAK dihapus), sama
# spt perlakuan di kolom INSURED.
_SUSPENSE_LIKE_RE = re.compile(r"SUSPEN|\(?H\)?\s*UTANG\s*PIUTANG", re.IGNORECASE)

# [BARU 15/09] teks narasi MURNI (BORDERO/SUMMARY SIMAS RE/nama bulan/
# tahun/pemisah saja, TANPA angka lain) -- dibiarkan apa adanya.
_NARRATIVE_ONLY_RE = re.compile(
    r"^(?:BORDERO|SUMMARY|BULAN|SIMAS\s*RE|" + _MONTH_RE + r"|(?:19|20)\d{2}|[-\s,])+$",
    re.IGNORECASE,
)

# [BARU] rantai P<n> (+ anotasi opsional "(ADD n)") yang digabung "+" --
# dibuang SELURUHNYA (termasuk "+"-nya), di mana pun posisinya (bukan
# cuma di akhir).
_STRAY_P_CHAIN_RE = re.compile(
    r"(?:\bP\d+\b\s*(?:\(ADD\s*\d+\))?\s*\+\s*)+", re.IGNORECASE
)
# [BARU] rantai P<n> yang dipisah SPASI (bukan "+") yang nempel di
# BELAKANG nilai -- dibuang.
_TRAILING_P_CHAIN_RE = re.compile(r"(?:\s+P\d+\b)+\s*$", re.IGNORECASE)

# [BARU] segmen pertama (sebelum "/" pertama) yang berisi RANTAI angka
# pendek digabung "-" (bukan "+") sebanyak >=5 kali -- dianggap
# 'pengulangan berlebih', DIBIARKAN APA ADANYA (raw), sama seperti
# aturan pengulangan >5 via "+".
def _has_excessive_dash_chain(value: str) -> bool:
    first_seg = value.split("/", 1)[0]
    return first_seg.count("-") >= 5


KEEP_AS_IS_PATTERN_RE = re.compile(r"^\d{7,10}[A-Z]{1,2}\d{6,9}-\d{2}$", re.IGNORECASE)


def _is_junk_token(tok: str) -> bool:
    t = tok.strip().upper()
    if not t:
        return True
    if t in _JUNK_STANDALONE_WORDS:
        return True
    if _P_CODE_RE.match(t):
        return True
    return False


def _value_has_real_digits(text: str) -> bool:
    return bool(re.search(r"\d{3,}", text))


def _has_junk_marker(text: str) -> bool:
    return bool(re.search(r"\bTBA\b|\bP\d+\b", text, re.IGNORECASE))


def _split_long_space_tail(val: str):
    parts = re.split(r"\s{2,}", val, maxsplit=1)
    if len(parts) == 1:
        return val, None
    return parts[0].strip(), parts[1].strip()


def _strip_leading_junk_wrapper(val: str):
    """Buang wrapper di depan via '/' SAMPAI ketemu nilai asli:
    - PASS 1: cari segmen PERTAMA yang punya digit -> itu awal nilai
      asli (buang semua sebelumnya, apa pun isinya -- nama cedant/broker
      mis. 'SINARMAS', placeholder TBA/P<n>, dst).
    - PASS 2 (kalau TIDAK ADA segmen berdigit sama sekali): hanya buang
      segmen placeholder BAKU (TBA/P<n>) di depan; segmen bukan-baku
      lain (mis. nama perusahaan tanpa digit) TETAP DIPERTAHANKAN kalau
      itu satu-satunya isi yang tersisa (mis. 'TBA/SINARMAS' -> 'SINARMAS').
    Kalau SEMUA segmen placeholder baku -> None (murni placeholder)."""
    if "/" not in val:
        return None if _is_junk_token(val) else val
    segs = val.split("/")
    for i, s in enumerate(segs):
        if re.search(r"\d", s):
            return "/".join(segs[i:]).strip()
    start_idx = None
    for i, s in enumerate(segs):
        if not _is_junk_token(s.strip()):
            start_idx = i
            break
    if start_idx is None:
        return None
    return "/".join(segs[start_idx:]).strip()


def _strip_trailing_junk_wrapper(val: str) -> str:
    """Buang segmen TANPA digit (nama cedant/broker/kode singkat, mis.
    'ACA', 'SINAR MAS') yang nempel di BELAKANG via '/', SAMPAI ketemu
    segmen yang punya digit (dianggap bagian nomor asli). BEDA dgn
    _strip_leading_junk_wrapper (berbasis _is_junk_token) -- di
    belakang, kriterianya 'ada digit atau tidak', krn tag broker/cedant
    yg nempel di ekor biasanya bukan kode baku (TBA/P<n>) tapi nama
    bebas apa saja (mis. 'ACA')."""
    if "/" not in val:
        return val
    segs = val.split("/")
    end_idx = None
    for i in range(len(segs) - 1, -1, -1):
        if re.search(r"\d", segs[i]):
            end_idx = i + 1
            break
    if end_idx is None:
        return val
    return "/".join(segs[:end_idx]).strip()


def _find_substitution_target(text: str):
    """Cari segmen digit yang jadi TARGET substitusi 'pengulangan'.
    Kasus khusus: pola 'ANGKA + 1-2 HURUF + ANGKA' PERSIS (mis.
    'S00000271' pada '110551216F00000473', tanpa karakter lain) ->
    target SEGMEN TERAKHIR (setelah huruf) -- meski segmen pertama
    (di depan huruf) sering lebih panjang, BUKAN itu yg dimaksud
    'pengulangan'. Selain pola ini -> segmen digit TERPANJANG; kalau
    ada BEBERAPA segmen dgn panjang SAMA (tie), pilih yg PERTAMA di
    antaranya yang BUKAN nol semua (mis. '0000-5747': '0000' dan '5747'
    sama panjang, tapi '0000' cuma nol -> target '5747')."""
    m = re.match(r"^(\d+)([A-Za-z]{1,2})(\d+)$", text)
    if m:
        return m.group(3), m.start(3), m.end(3)
    candidates = list(re.finditer(r"\d+", text))
    if not candidates:
        return None, None, None
    max_len = max(len(c.group(0)) for c in candidates)
    tied = [c for c in candidates if len(c.group(0)) == max_len]
    non_zero = [c for c in tied if set(c.group(0)) != {"0"}]
    best = non_zero[0] if non_zero else tied[0]
    return best.group(0), best.start(), best.end()


def _qualifies_for_substitution(main: str) -> bool:
    """Nilai utama cukup 'berstruktur' utk dijadikan target substitusi
    'pengulangan' kalau: ada huruf (pola FPAR-style), ATAU >=2 pemisah
    '.'/'/' (pola dot-style spt '12.100.0001.82288'), ATAU total digitnya
    panjang (>=10, pola digit polos panjang spt '0124012000618'), ATAU
    ada >=1 pemisah '.'/'/'/'-' DAN total digit >=7 (pola spt '0000-
    5747', beda dari angka pendek polos TANPA pemisah spt '3522' yang
    TETAP tidak memenuhi)."""
    if re.search(r"[A-Za-z]", main):
        return True
    sep_count = len(re.findall(r"[./]", main))
    if sep_count >= 2:
        return True
    digit_len = len(re.sub(r"\D", "", main))
    if digit_len >= 10:
        return True
    if len(re.findall(r"[.\-/]", main)) >= 1 and digit_len >= 7:
        return True
    return False


def _apply_segment_substitution(primary: str, extra_tokens: list) -> list:
    seg, start, end = _find_substitution_target(primary)
    if seg is None:
        return [primary]
    prefix, suffix = primary[:start], primary[end:]
    out = [primary]
    for tok in extra_tokens:
        tok = tok.strip()
        if not tok.isdigit():
            continue
        if len(tok) >= len(seg):
            new_seg = tok
        else:
            new_seg = seg[: -len(tok)] + tok
        out.append(prefix + new_seg + suffix)
    return _dedup_preserve_order(out)


# [BARU 15/09] dash NYASAR di dalam pola F/S 2-segmen (mis.
# '11055129F00000-275' -- seharusnya '11055129F00000275') -- dash
# dianggap noise/typo, dilebur.
_FS_STRAY_DASH_RE = re.compile(r"^(\d+[A-Za-z]{1,2}\d+)-(\d+)$")


def _merge_fs_stray_dash(token: str) -> str:
    m = _FS_STRAY_DASH_RE.match(token)
    if m:
        return m.group(1) + m.group(2)
    return token


_CN_TRAILING_RE = re.compile(r"\s*/?\s*CN\s*:\s*.*$", re.IGNORECASE)
_CN_LEADING_WITH_SLASH_RE = re.compile(r"^\s*CN\s*:\s*[^/]*/\s*", re.IGNORECASE)
_CN_LABEL_ONLY_RE = re.compile(r"^\s*CN\s*:\s*", re.IGNORECASE)
_CN_STARTS_RE = re.compile(r"^CN\s*:", re.IGNORECASE)


def _strip_cn_suffix(val: str) -> str:
    """Notasi 'CN : <nomor>':
    - Muncul di TENGAH/BELAKANG ('main / CN : cn') -> dibuang total
      bersama nilainya, sisakan main di depan.
    - Muncul di AWAL DIIKUTI nilai lain setelah '/' ('CN : cn / main')
      -> label+nilai CN dibuang, sisakan main setelah '/'.
    - HANYA 'CN : cn' saja (seluruh value, tanpa '/' sama sekali) ->
      cuma label 'CN :'-nya yg dibuang, nilai CN dipertahankan sbg
      hasil utama (BUKAN dibuang semua)."""
    stripped = val.strip()
    if not _CN_STARTS_RE.match(stripped):
        return _CN_TRAILING_RE.sub("", val).strip(" /-")
    m = _CN_LEADING_WITH_SLASH_RE.match(val)
    if m:
        return val[m.end():].strip(" /-")
    return _CN_LABEL_ONLY_RE.sub("", val).strip(" /-")


_EMBEDDED_SUFFIX_RE = re.compile(r"^(\d+)([./].*[A-Za-z].*)$")


def _split_embedded_suffix(token: str):
    """Deteksi token yg formatnya 'ANGKA' + 'SUFFIX BERSTRUKTUR' nempel
    jadi satu (mis. '23022/FIAR/001.0000344-SRB/06/2022' -> ('23022',
    '/FIAR/001.0000344-SRB/06/2022')). Dipakai utk kasus 'pengulangan'
    di mana suffix format cuma nempel di SALAH SATU token dlm rantai
    '+' (biasanya yg terakhir), bukan di token pertama."""
    m = _EMBEDDED_SUFFIX_RE.match(token)
    if m:
        return m.group(1), m.group(2)
    return token, None


def _strip_all_characters(value: str) -> str:
    return re.sub(r"[^A-Za-z0-9]", "", value)


_SINAR_MAS_RE = re.compile(r"\bSINAR\s+MAS\b", re.IGNORECASE)
_TEXT_PIECE_PTCV_RE = re.compile(r",?\s*\b(?:P\s*T|CV)\b\.?\s*$", re.IGNORECASE)


def _clean_text_piece(value: str) -> str:
    """Utk hasil yang ternyata berupa TEKS/NAMA (tanpa digit sama
    sekali, muncul lewat mekanisme buang TBA/P<n> yang bersandingan) --
    JANGAN di-'hapus karakter' penuh (spasi antar kata dipertahankan,
    beda dari nilai numerik biasa), cuma normalisasi spasi + standardisasi
    'SINAR MAS' -> 'SINARMAS' (konsisten dgn kolom INSURED) + buang
    suffix "PT"/"CV" di akhir (mis. "ASURANSI FPG INSURANCE, PT" ->
    "ASURANSI FPG INSURANCE")."""
    value = _SINAR_MAS_RE.sub("SINARMAS", value)
    value = _TEXT_PIECE_PTCV_RE.sub("", value)
    return _normalize_spaces(value).upper()


def _simplify_noletter_extra(e: str) -> str:
    """[BARU 15/09] Token tambahan (di rantai '+') yang TIDAK punya
    huruf tapi masih ada '/' atau '.' nempel (mis. '04820/0001189' dari
    '...+04820/0001189.') -- ambil cuma potongan PERTAMA sbg kandidat
    substitusi digit, sisanya (mis. sertifikat/nomor duplikat) dibuang
    (KBRU tidak punya kolom sertifikat, jadi info ini tidak tertampung)."""
    if re.search(r"[A-Za-z]", e):
        return e
    return re.split(r"[./]", e)[0]


def clean_value(val) -> list:
    if pd.isna(val):
        return []
    raw = str(val).strip()
    if not raw:
        return []

    if raw.upper() in _JUNK_STANDALONE_WORDS:
        return [raw]
    # [BARU 15/09] "TBA/" (placeholder + garis miring doang, tanpa isi
    # lain) -> garis miringnya dibuang, cukup "TBA".
    raw_no_slash = raw.strip("/ ")
    if raw_no_slash.upper() in _JUNK_STANDALONE_WORDS:
        return [raw_no_slash]
    if KEEP_AS_IS_PATTERN_RE.match(raw):
        return [raw]
    if re.fullmatch(r"\d+&\d+", raw):
        return [raw]
    if _has_excessive_dash_chain(raw):
        # segmen depan berisi >=5 angka digabung '-' -- pengulangan
        # berlebih, dibiarkan apa adanya (sama spt aturan >5 via '+')
        return [raw]
    if _SUSPENSE_LIKE_RE.search(raw) or _NARRATIVE_ONLY_RE.match(raw.strip()):
        # [BARU 15/09] deskripsi SUSPENSE/UTANG PIUTANG, atau narasi
        # murni (BORDERO/bulan/tahun/SUMMARY SIMAS RE) yg nyasar di
        # kolom polis/slip -- dibiarkan apa adanya, spasi TIDAK dihapus
        return [raw]

    if not _has_junk_marker(raw) and not _value_has_real_digits(raw):
        # teks tanpa placeholder (TBA/P<n>) & tanpa digit sama sekali
        # (mis. nama perusahaan nyasar, "ASURANSI SINAR MAS") -> tetap
        # "abaikan" (bukan diproses spt polis/slip), TAPI tetap lewat
        # normalisasi teks (spasi + SINAR MAS->SINARMAS) spy konsisten
        # dgn kasus yg lewat jalur TBA-adjacent.
        return [_clean_text_piece(raw)]

    val2 = _NO_ANYWHERE_RE.sub("", raw)
    val2 = _strip_cn_suffix(val2)
    val2 = _LEADING_ALPHA_PLUS_RE.sub("", val2)
    val2 = _STRAY_P_CHAIN_RE.sub("", val2)
    val2 = _TRAILING_P_CHAIN_RE.sub("", val2)

    pieces = []
    stripped = _strip_leading_junk_wrapper(val2)
    if stripped is not None:
        val2 = stripped
        val2 = _TRAILING_P_CHAIN_RE.sub("", val2)
        primary, _tail = _split_long_space_tail(val2)
        primary2 = _MONTH_YEAR_TAIL_RE.sub("", primary).strip(" -/,")
        primary = primary2 if primary2 else primary
        primary = _TRAILING_JUNK_WORD_RE.sub("", primary).strip()

        raw_parts = [p.strip() for p in re.split(r"\s*\+\s*", primary) if p.strip()]
        raw_parts = [_merge_fs_stray_dash(p) for p in raw_parts]
        # [FIX 15/09] "_strip_trailing_junk_wrapper" cuma dipakai kalau
        # MEMANG ada rantai '+' (>1 bagian) -- kalau cuma 1 bagian
        # (whole value, mis. hasil gabungan lewat "No." di tengah),
        # JANGAN dipangkas, supaya ekor sah spt "/AB"/"/BB" (kode akhir
        # 2 huruf tanpa digit) tidak ikut kebuang.
        apply_trailing_strip = len(raw_parts) > 1
        parts = []
        for p in raw_parts:
            if _is_junk_token(p):
                continue
            p2 = _strip_trailing_junk_wrapper(p) if apply_trailing_strip else p
            p2 = p2.strip()
            if p2 and not _is_junk_token(p2):
                parts.append(p2)

        if parts:
            if len(parts) > 1:
                # kalau ADA campuran (sebagian bagian punya digit,
                # sebagian tidak sama sekali, mis. 'UMUM MEGA' vs
                # '2010208012400001') -> bagian tanpa digit dibuang
                # (nama cedant/broker yg nyempil, bukan bagian nilai).
                digit_parts = [p for p in parts if re.search(r"\d", p)]
                if digit_parts:
                    parts = digit_parts

            # [CEK 1] suffix berstruktur nempel di SALAH SATU token
            # (biasanya token TERAKHIR, mis. '23018+23019+23023+23022/
            # FIAR/001.0000344-SRB/06/2022') -- kalau ketemu & SISANYA
            # semua digit polos, terapkan suffix yg sama ke semua nomor.
            suffix_found = None
            base_numbers = []
            embed_ok = True
            for p in parts:
                num_part, suf = _split_embedded_suffix(p)
                if suf:
                    if suffix_found is None:
                        suffix_found = suf
                    else:
                        embed_ok = False
                if not num_part.isdigit():
                    embed_ok = False
                base_numbers.append(num_part)

            if embed_ok and suffix_found and len(base_numbers) > 1:
                ref = max(base_numbers, key=len)
                combined = []
                for n in base_numbers:
                    if len(n) >= len(ref):
                        combined.append(n + suffix_found)
                    else:
                        combined.append(ref[: -len(n)] + n + suffix_found)
                pieces = _dedup_preserve_order(combined)
            else:
                # [CEK 2] main = bagian yg PALING berstruktur (ada digit
                # DAN (huruf ATAU >=2 pemisah)), di mana pun posisinya
                # dalam rantai '+' -- sisanya jadi kandidat token
                # substitusi/nilai lain.
                structured_idx = None
                for i, p in enumerate(parts):
                    if re.search(r"\d", p) and (
                        re.search(r"[A-Za-z]", p) or len(re.findall(r"[./]", p)) >= 2
                    ):
                        structured_idx = i
                        break
                if structured_idx is not None:
                    main = parts[structured_idx]
                    extras_raw = parts[:structured_idx] + parts[structured_idx + 1:]
                else:
                    main = parts[0]
                    extras_raw = parts[1:]

                extras_clean = [_simplify_noletter_extra(e.strip().rstrip(".").strip()) for e in extras_raw]
                whole_extras = [e for e in extras_clean if "/" in e or re.search(r"\d\.\d", e)]
                digit_extras = [e for e in extras_clean if e.isdigit()]

                # [BARU 15/09] token substitusi yg SANGAT pendek (<=2
                # digit, mis. "+97+07+43") -- pola pengulangannya tidak
                # jelas (berapa digit ekor yg diganti jadi ambigu) ->
                # seluruh nilai dibiarkan apa adanya.
                if digit_extras and any(len(e) < 3 for e in digit_extras):
                    pieces = []
                elif digit_extras and not _qualifies_for_substitution(main):
                    pieces = []  # ambigu -> jatuh ke abaikan (raw) di bawah
                elif digit_extras:
                    pieces = _apply_segment_substitution(main, digit_extras)
                    pieces += whole_extras
                else:
                    pieces = [main] + whole_extras

    pieces = _dedup_preserve_order([_normalize_spaces(p) for p in pieces if p.strip()])

    if not pieces or len(pieces) > MAX_SPLIT_COLS:
        return [raw]

    final = []
    for p in pieces:
        if KEEP_AS_IS_PATTERN_RE.match(p):
            final.append(p)
        elif not re.search(r"\d", p):
            final.append(_clean_text_piece(p))
        else:
            final.append(_strip_all_characters(p))
    return final


def clean_polis(val):
    return clean_value(val)


def clean_slip(val):
    return clean_value(val)


# ─────────────────────────────────────────────────────────────────────────────
# CLEAN INSURED
# ─────────────────────────────────────────────────────────────────────────────

_GELAR_PREFIX_RE = re.compile(
    r"^\s*(?:BAPAK|IBU|SDRI?|NYONYA|NY|TUAN|TN|MR|MRS|MS|DR|IR|PROF|HJ|H)\b\.?\s*",
    re.IGNORECASE,
)
_GELAR_SUFFIX_RE = re.compile(
    r"\s*,?\s*\b(?:S\.?E|S\.?H|S\.?SI|S\.?T|S\.?SOS|S\.?KOM|S\.?PD|"
    r"M\.?SI|M\.?M|M\.?H|M\.?KOM|SH|SE|BP|MRS|MR|NY|TN|DRS)\b\.?\s*$",
    re.IGNORECASE,
)
_SUSPENSE_KEEP_RE = re.compile(r"SUSPEN|\(?H\)?\s*UTANG\s*PIUTANG", re.IGNORECASE)
_TREATY_KEEP_RE = re.compile(
    r"\bTREATY\b|\bSTRUCTURE\b|EXCESS\s+OF\s+LOSS|WHOLE\s+ACCOUNT", re.IGNORECASE
)


def _is_keep_as_is_insured(val: str) -> bool:
    if _SUSPENSE_KEEP_RE.search(val):
        return True
    if _TREATY_KEEP_RE.search(val):
        return True
    return False


def _collapse_duplicate_words(text: str) -> str:
    """Normalisasi data rusak: kata terduplikasi persis berturut-turut
    (mis. 'WASKITA WASKITA KARYA KARYA' -> 'WASKITA KARYA'), kata
    gandeng tanpa spasi ('PERSEROPERSERO' -> 'PERSERO'), kurung ganda/
    tak seimbang."""
    text = re.sub(r"\(\s*\(", "(", text)
    text = re.sub(r"\)\s*\)", ")", text)
    text = re.sub(r"\b(\w{3,})\1\b", r"\1", text, flags=re.IGNORECASE)
    text = re.sub(r"\b(\w+)(?:\s+\1\b)+", r"\1", text, flags=re.IGNORECASE)
    if text.count("(") > text.count(")"):
        text = text + ")"
    return text


_KEEP_SUFFIX_COMMA_RE = re.compile(r",\s*(UD|PD)\b", re.IGNORECASE)
_REMOVE_COMMA_PTCV_RE = re.compile(r",\s*(?:P\s*T|CV)\b\.?", re.IGNORECASE)

# pemisah: QQ/QQQ, OO (typo QQ), AND/OR, "/", "//", ",", en-dash, "¿",
# "*", dan "&" HANYA kalau diikuti PT/CV. "+" dan " - " BUKAN pemisah.
_SPLIT_RE = re.compile(
    r"\bQ{2,3}\b|\bOO\b|\bAND\s*/\s*OR\.?\b|/{1,2}|,|–|¿|\*|&\s*(?=(?:PT|CV)\b)",
    re.IGNORECASE,
)

INSURED_JUNK_WORDS = frozenset({
    "PT", "CV", "TBK", "PERSERO", "LTD", "INC", "LLC",
    "AND", "OR", "THE", "OF", "AS", "VARIOUS", "VAR", "QQ", "TBA",
    "ASSOCIATED",
})

_MONTH_STANDALONE_RE = re.compile(
    r"^(?:" + _MONTH_RE + r")\.?\s*(?:19|20)?\d{2,4}$", re.IGNORECASE
)

_GELAR_PREFIX_RE = re.compile(
    r"^\s*(?:BAPAK|IBU|SDRI?|NYONYA|NY|TUAN|TN|MR|MRS|MS|DR|IR|PROF|HJ|H)\b\.?\s*",
    re.IGNORECASE,
)
# [BARU] tambah BP/MR/MRS/NY/TN/DRS sbg gelar SUFFIX (nempel via koma di
# BELAKANG nama, mis. "SUTJITA, BP" / "RUDY AKILI, MR") -- dibuang.
_GELAR_SUFFIX_RE = re.compile(
    r"\s*,?\s*\b(?:S\.?E|S\.?H|S\.?SI|S\.?T|S\.?SOS|S\.?KOM|S\.?PD|"
    r"M\.?SI|M\.?M|M\.?H|M\.?KOM|SH|SE|BP|MRS|MR|NY|TN|DRS)\b\.?\s*$",
    re.IGNORECASE,
)
_SUSPENSE_KEEP_RE = re.compile(r"SUSPEN|\(?H\)?\s*UTANG\s*PIUTANG", re.IGNORECASE)
_TREATY_KEEP_RE = re.compile(
    r"\bTREATY\b|\bSTRUCTURE\b|EXCESS\s+OF\s+LOSS|WHOLE\s+ACCOUNT", re.IGNORECASE
)


def _is_keep_as_is_insured(val: str) -> bool:
    if _SUSPENSE_KEEP_RE.search(val):
        return True
    if _TREATY_KEEP_RE.search(val):
        return True
    if _MALFORMED_QQ_RE.search(val):
        return True
    return False


# [BARU] normalisasi data rusak: kata terduplikasi persis berturut-turut
# (mis. "WASKITA WASKITA KARYA KARYA" -> "WASKITA KARYA"), kata gandeng
# tanpa spasi ("PERSEROPERSERO" -> "PERSERO"), kurung ganda/tak seimbang.
def _collapse_duplicate_words(text: str) -> str:
    text = re.sub(r"\(\s*\(", "(", text)
    text = re.sub(r"\)\s*\)", ")", text)
    text = re.sub(r"\b(\w{3,})\1\b", r"\1", text, flags=re.IGNORECASE)
    text = re.sub(r"\b(\w+)(?:\s+\1\b)+", r"\1", text, flags=re.IGNORECASE)
    if text.count("(") > text.count(")"):
        text = text + ")"
    return text


# [BARU] ", PT" / ", CV" (di mana pun, bukan cuma di akhir) DIBUANG
# TOTAL bersama komanya -- supaya tidak memicu split palsu saat ada isi
# lain menyertai PT/CV di segmen yang sama (mis. "SMAILING TOUR, PT &
# TRAVEL SERVICE, PT" -> "SMAILING TOUR & TRAVEL SERVICE", bukan
# terpecah jadi "SMAILING TOUR" + "& TRAVEL SERVICE").
_REMOVE_COMMA_PTCV_RE = re.compile(r",\s*(?:P\s*T|CV)\b\.?", re.IGNORECASE)

# ", UD" / ", PD" (jenis badan usaha) TIDAK dianggap pemisah entitas
# baru -- koma dibuang, kode dipertahankan menyatu dgn nama sebelumnya.
_KEEP_SUFFIX_COMMA_RE = re.compile(r",\s*(UD|PD)\b", re.IGNORECASE)

# pemisah: QQ/QQQ, OO (typo QQ), AND/OR, "/", "//", ",", en-dash, "¿",
# "*", dan "&" HANYA kalau diikuti PT/CV (JV baru, mis. "... & PT OKI
# PULP..."). "&" polos (mis. "PULP & PAPER") TETAP bukan pemisah. "+"
# TIDAK dipakai sbg pemisah (mis. "SK+SUMATERA" -> 1 entitas "SK
# SUMATERA", bukan 2) -- beda dgn mesin polis/slip. " - " (spasi-dash-
# spasi) juga TIDAK dipakai (bikin false split di nama spt "PULAU -
# PULAU MEDIA").
_SPLIT_RE = re.compile(
    r"\bQ{2,3}\b|\bOO\b|\bAND\s*/\s*OR\.?\b|/{1,2}|,|–|¿|\*|&\s*(?=(?:PT|CV)\b)",
    re.IGNORECASE,
)

INSURED_JUNK_WORDS = frozenset({
    "PT", "CV", "TBK", "PERSERO", "LTD", "INC", "LLC",
    "AND", "OR", "THE", "OF", "AS", "VARIOUS", "VAR", "QQ", "TBA",
    "ASSOCIATED",
})

_OWNER_OF_RE = re.compile(
    r"\(\s*AS\s+(?:THE\s+)?(?:OWNER|PRINCIPAL)\s+OF\s+([^()]+)\)", re.IGNORECASE
)

_ROLE_KEYWORDS_RE = re.compile(
    r"^(?:PERSERO|PRINCIPAL|OWNER|INSURED|CONTRACTOR)$", re.IGNORECASE
)

# [DIPERLONGGAR] toleransi "P T" (spasi nyempil di antara P & T, data
# rusak) selain "PT" biasa.
_PT_CV_PREFIX_RE = re.compile(r"^\s*(?:P\s*T|CV)(?:\.|(?=\s|$))\.?\s*", re.IGNORECASE)
_PT_CV_SUFFIX_RE = re.compile(r",?\s*(?:P\s*T|CV)\.?\s*$", re.IGNORECASE)

_MASK_COMMA = "\x01"
_MASK_SLASH = "\x02"


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
    words = [w for w in re.split(r"[\s\-]+", long_name.upper()) if w]
    if any(w.startswith(short_clean) and w != short_clean for w in words):
        return True
    return False


def _extract_owner_of_entities(text: str):
    extra = []

    def repl(m):
        extra.append(m.group(1).strip())
        return " "

    new_text = _OWNER_OF_RE.sub(repl, text)
    return new_text, extra


def _strip_parenthetical(text: str) -> str:
    """Isi dalam kurung yang singkatan dari teks sebelumnya, ATAU kata
    boilerplate peran (PERSERO/PRINCIPAL/OWNER/INSURED/CONTRACTOR), ATAU
    mengandung PT/CV -> kurung + isi DIBUANG total. Isi kurung yang
    mengandung KOMA (daftar deskriptif, mis. "(REFINERY, PLANTATION,
    KEMITRAAN)") -> kurung DIPERTAHANKAN UTUH (literal, koma di-mask
    dulu spy tdk ikut split). Isi kurung lain (identitas cabang/pihak
    lain, mis. "(KSO)") -> kurung dibuang, ISI dipertahankan."""

    def repl(m):
        inner = m.group(1).strip()
        if re.search(r"\bPT\b|\bCV\b", inner, re.IGNORECASE):
            return " "
        if _ROLE_KEYWORDS_RE.match(inner):
            return " "
        inner_letters = re.sub(r"[^A-Za-z]", "", inner)
        before = text[: m.start()].strip()
        if before and inner_letters and _is_abbreviation_of(inner_letters, before):
            return " "
        if "," in inner:
            masked = inner.replace(",", _MASK_COMMA).replace("/", _MASK_SLASH)
            return f" ({masked}) "
        masked = inner.replace(",", _MASK_COMMA).replace("/", _MASK_SLASH)
        return f" {masked} "

    return re.sub(r"\(([^()]*)\)", repl, text)


# ─────────────────────────────────────────────────────────────────────────────
# [BARU] EKSPANSI POLA "KEYWORD + DAFTAR" (FACTORY/CLUSTER/TOWER/
# REGIONAL/DISTRICT) & KHUSUS PELABUHAN INDONESIA/PELINDO
# ─────────────────────────────────────────────────────────────────────────────

_ROMAN_MAP = {"I": 1, "II": 2, "III": 3, "IV": 4, "V": 5, "VI": 6, "VII": 7, "VIII": 8, "IX": 9, "X": 10}
_ROMAN_ARABIC_ALT = r"I|II|III|IV|V|VI|VII|VIII|IX|X|\d+"

_PELINDO_RE = re.compile(
    r"^(PELABUHAN\s+INDONESIA)\s+((?:" + _ROMAN_ARABIC_ALT + r")"
    r"(?:\s*(?:,|&|DAN)\s*(?:" + _ROMAN_ARABIC_ALT + r"))+)\s*$",
    re.IGNORECASE,
)

# [BARU 15/09] "PELINDO" adalah SINONIM "PELABUHAN INDONESIA" -- bisa
# disebut BERULANG-ULANG dgn pemisah apa saja (biasanya QQ), mis.
# "PELABUHAN INDONESIA I QQ PELINDO I QQ PELINDO II QQ PELINDO III QQ
# PELINDO IV" -> semua kemunculan "PELABUHAN INDONESIA <n>"/"PELINDO <n>"
# dikumpulkan, dinormalisasi ke "PELABUHAN INDONESIA <n>" (arab), dedup.
_PELINDO_MENTION_RE = re.compile(
    r"\b(?:PELABUHAN\s+INDONESIA|PELINDO)\s+(" + _ROMAN_ARABIC_ALT + r")\b",
    re.IGNORECASE,
)


def _expand_pelindo(val: str):
    m = _PELINDO_RE.match(val.strip())
    if not m:
        return None
    prefix, nums_part = m.groups()
    tokens = [t.strip() for t in re.split(r"\s*(?:,|&|DAN)\s*", nums_part, flags=re.IGNORECASE) if t.strip()]
    out = []
    for t in tokens:
        num = _ROMAN_MAP.get(t.upper(), t)
        out.append(f"{prefix.upper()} {num}")
    return out


def _expand_pelindo_mentions(val: str):
    matches = _PELINDO_MENTION_RE.findall(val)
    if len(matches) < 2:
        return None
    seen = []
    for t in matches:
        num = _ROMAN_MAP.get(t.upper(), t)
        entry = f"PELABUHAN INDONESIA {num}"
        if entry not in seen:
            seen.append(entry)
    return seen if len(seen) >= 2 else None


_KEYWORD_MULTI_RE = re.compile(
    r"^([^/]*?)[/\s]*\b(FACTORY|CLUSTER|TOWER|REGIONAL|DISTRICT|UNIT|MALL)\s+([^/]+)$",
    re.IGNORECASE,
)


def _expand_keyword_multi(val: str):
    """'PREFIX [/] KEYWORD tok1 tok2 ...' atau 'PREFIX [/] KEYWORD tok1,
    tok2 & tok3' -> daftar 'PREFIX KEYWORD tok'. Butuh >=2 token supaya
    tidak salah nembak segmen tunggal (yg ditangani _merge_keyword_single
    di tempat lain)."""
    m = _KEYWORD_MULTI_RE.match(val.strip())
    if not m:
        return None
    prefix, keyword, tail = m.groups()
    tokens = [t for t in re.split(r"\s*(?:,|&|DAN)\s*|\s+", tail.strip(), flags=re.IGNORECASE) if t.strip()]
    tokens = [t.strip() for t in tokens if re.fullmatch(r"[A-Za-z0-9]{1,4}", t.strip())]
    if len(tokens) < 2:
        return None
    prefix = prefix.strip().rstrip(",").strip()
    prefix = re.sub(r",?\s*(?:P\s*T|CV)\.?\s*$", "", prefix, flags=re.IGNORECASE).strip()
    return [f"{prefix} {keyword.upper()} {t.upper()}".strip() for t in tokens]


_KEYWORD_SINGLE_RE = re.compile(
    r"^(?:FACTORY|CLUSTER|TOWER|REGIONAL|DISTRICT|UNIT|MALL)\s+[A-Za-z0-9]{1,4}$", re.IGNORECASE
)

# [BARU 15/09] "QQ" yang NYASAR nempel langsung ke huruf lain tanpa
# spasi/pemisah (mis. "QQBALI", "INDAHQQ") -- tanda data rusak/typo yg
# terlalu ambigu utk diproses -> seluruh value dibiarkan apa adanya.
# Rantai Q murni ("QQQ") TIDAK dihitung rusak -- itu cuma varian QQ
# dgn Q berlebih, sudah ditangani _SPLIT_RE (\bQ{2,3}\b).
_MALFORMED_QQ_RE = re.compile(r"[A-PR-Za-pr-z]QQ|QQ[A-PR-Za-pr-z]")


def _clean_insured_name(name: str) -> str:
    name = _normalize_spaces(name)
    name = _GELAR_PREFIX_RE.sub("", name)
    for _ in range(3):
        stripped = _GELAR_SUFFIX_RE.sub("", name).strip()
        if stripped == name:
            break
        name = stripped
    name = re.sub(r",?\s*\(\s*PERSERO\s*\)\s*", " ", name, flags=re.IGNORECASE)
    name = re.sub(r"\bTBK\b\s*,\s*P\s*T\.?\s*", " ", name, flags=re.IGNORECASE)
    name = re.sub(r",?\s*\bTBK\b\s*", " ", name, flags=re.IGNORECASE)
    name = re.sub(r"\bLTD\b\.?\s*$", "", name, flags=re.IGNORECASE)
    name = re.sub(r"\bPTE\b\.?\s*", "", name, flags=re.IGNORECASE)
    name = _PT_CV_SUFFIX_RE.sub("", name)
    name = _PT_CV_PREFIX_RE.sub("", name)
    # [FIX 15/09] pola inisial huruf-tunggal-bertitik berturut-turut
    # (mis. "S.A.R.I") -> titik dibuang TANPA jadi spasi ("SARI"),
    # BUKAN "S A R I". Titik lain (bukan pola ini, bukan antar digit)
    # baru diganti spasi spt biasa.
    name = re.sub(r"\b(?:[A-Za-z]\.){2,}[A-Za-z]?\b", lambda mo: mo.group(0).replace(".", ""), name)
    name = re.sub(r"(?<!\d)\.(?!\d)", " ", name)
    name = re.sub(r"\s*\+\s*", " ", name)
    name = _SINAR_MAS_RE.sub("SINARMAS", name)
    name = name.replace(_MASK_COMMA, ",").replace(_MASK_SLASH, "/")
    return _normalize_spaces(name).strip().upper()


def _try_expansions(segment: str):
    """Coba semua mekanisme ekspansi (pelindo list, pelindo scattered
    mentions, keyword+daftar) pada SATU string. None kalau tidak ada yg
    cocok."""
    r = _expand_pelindo(segment)
    if r is not None:
        return r
    r = _expand_pelindo_mentions(segment)
    if r is not None:
        return r
    r = _expand_keyword_multi(segment)
    if r is not None:
        return r
    return None


def clean_insured(val) -> list:
    if pd.isna(val):
        return []
    val = str(val).strip()
    if not val:
        return []

    val = _collapse_duplicate_words(val)

    if _is_keep_as_is_insured(val):
        return [_normalize_spaces(val).upper()]

    val, owner_extra = _extract_owner_of_entities(val)
    val = _strip_parenthetical(val)
    val = _KEEP_SUFFIX_COMMA_RE.sub(r" \1", val)
    val = _REMOVE_COMMA_PTCV_RE.sub("", val)

    # coba ekspansi pada SELURUH value dulu (kasus umum: satu pola
    # keyword/pelindo mencakup semua, mis. "BIO FARMA/CLUSTER B C D",
    # "PELABUHAN INDONESIA I QQ PELINDO II ...").
    whole = _try_expansions(val)
    if whole is not None:
        return [_normalize_spaces(x).upper() for x in whole]

    parts = [p.strip() for p in _SPLIT_RE.split(val) if p.strip()]
    cleaned = []
    for p in parts:
        # [FIX 15/09] kalau SATU bagian hasil split (bukan whole value)
        # sendiri cocok pola keyword+daftar (mis. "GALAXY MALL I & II"
        # setelah "SINAR GALAXY /" ikut terpisah lebih dulu) -> expand
        # jadi beberapa entitas, JANGAN diproses sbg 1 entitas biasa.
        r = _expand_keyword_multi(p)
        if r is not None and len(r) >= 2:
            cleaned.extend(_normalize_spaces(x).upper() for x in r)
            continue
        p_clean = _clean_insured_name(p)
        if not p_clean:
            continue
        if _MONTH_STANDALONE_RE.match(p_clean):
            continue
        if _KEYWORD_SINGLE_RE.match(p_clean) and cleaned:
            cleaned[-1] = _normalize_spaces(cleaned[-1] + " " + p_clean)
            continue
        if len(p_clean) < 2 or p_clean in INSURED_JUNK_WORDS:
            continue
        cleaned.append(p_clean)

    for extra in owner_extra:
        extra_clean = _clean_insured_name(extra)
        if extra_clean and extra_clean not in cleaned:
            cleaned.append(extra_clean)

    if not cleaned:
        fallback = _clean_insured_name(val)
        return [fallback] if fallback else []

    return _cap_or_join(cleaned)


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


def _insert_polis_cert_columns(df, all_polis_lists, max_cols):
    """POLIS_CLEAN_i dgn CERTIFICATE_1 TEPAT setelah POLIS_CLEAN_1
    (SELALU kosong -- cedant ini tidak punya konsep sertifikat)."""
    added = []
    actual_max_cols = max((len(lst) for lst in all_polis_lists if isinstance(lst, list)), default=1)
    actual_max_cols = min(actual_max_cols, max_cols)
    for i in range(1, actual_max_cols + 1):
        polis_col = f"POLIS_CLEAN_{i}"
        df[polis_col] = [lst[i - 1] if isinstance(lst, list) and i - 1 < len(lst) else None for lst in all_polis_lists]
        added.append(polis_col)
        if i == 1:
            df["CERTIFICATE_1"] = None
            added.append("CERTIFICATE_1")
    return added


def process_data(input_file: str, output_file: str) -> None:
    print(f"[1/5] Membaca data dari: {input_file} (sheet: {INPUT_SHEET}) ...")
    df = pd.read_excel(input_file, sheet_name=INPUT_SHEET, header=HEADER_ROW)
    print(f"      Total baris keseluruhan: {len(df):,}")

    if CEDANT_FILTER_COL not in df.columns:
        print(f"\n[ERROR] Kolom '{CEDANT_FILTER_COL}' tidak ditemukan!")
        print(f"        Kolom tersedia: {list(df.columns)}")
        return

    df[CEDANT_FILTER_COL] = df[CEDANT_FILTER_COL].astype(str).str.strip()
    df = df[df[CEDANT_FILTER_COL] == CEDANT_FILTER_VALUE].copy()
    print(f"[2/5] Filter cedant '{CEDANT_FILTER_VALUE}': {len(df):,} baris ditemukan.")

    if df.empty:
        print("\n[WARN] Tidak ada data setelah filter. Proses dihentikan.")
        return

    if STATUS_COL not in df.columns:
        print(f"\n[ERROR] Kolom '{STATUS_COL}' tidak ditemukan!")
        print(f"        Kolom tersedia: {list(df.columns)}")
        return

    before_status = len(df)
    df[STATUS_COL] = df[STATUS_COL].astype(str).str.strip()
    df = df[df[STATUS_COL] == STATUS_KEEP_VALUE].copy()
    print(f"[2b/5] Filter status '{STATUS_KEEP_VALUE}': {len(df):,} baris "
          f"(dibuang {before_status - len(df):,} baris ADJUSTED).")

    if df.empty:
        print("\n[WARN] Tidak ada data setelah filter status. Proses dihentikan.")
        return

    for col in [INSURED_COL, POLIS_COL, SLIP_COL]:
        if col not in df.columns:
            print(f"\n[ERROR] Kolom '{col}' tidak ditemukan di file!")
            print(f"        Kolom tersedia: {list(df.columns)}")
            return

    df.rename(columns={
        INSURED_COL: "INSURED_ORI",
        POLIS_COL:   "POLIS_ORI",
        SLIP_COL:    "SLIP_NO_ORI",
    }, inplace=True)

    print("[3/5] Menjalankan proses cleaning ...")

    all_polis, all_slip, all_insured = [], [], []
    max_polis = max_slip = max_ins = 1

    for _, row in df.iterrows():
        p_clean = clean_polis(row.get("POLIS_ORI", ""))
        s_clean = clean_slip(row.get("SLIP_NO_ORI", ""))
        parts = clean_insured(row.get("INSURED_ORI", ""))

        max_polis = max(max_polis, len(p_clean))
        max_slip = max(max_slip, len(s_clean))
        max_ins = max(max_ins, len(parts))

        all_polis.append(p_clean)
        all_slip.append(s_clean)
        all_insured.append(parts)

    print("[4/5] Menyusun kolom output ...")

    insured_cols = _insert_clean_columns(df, all_insured, "INSURED", max_ins)
    polis_cols = _insert_polis_cert_columns(df, all_polis, max_polis)
    slip_cols = _insert_clean_columns(df, all_slip, "SLIP_CLEAN", max_slip)

    new_columns = []
    for col in df.columns:
        if col in insured_cols + polis_cols + slip_cols:
            continue
        new_columns.append(col)
        if col == "INSURED_ORI":
            new_columns += insured_cols
        elif col == "POLIS_ORI":
            new_columns += polis_cols
        elif col == "SLIP_NO_ORI":
            new_columns += slip_cols

    df = df[new_columns]

    os.makedirs(os.path.dirname(output_file), exist_ok=True)
    print(f"[5/5] Menyimpan hasil ke: {output_file} ...")
    df.to_excel(output_file, index=False)
    _force_text_format(output_file, prefixes=("POLIS_CLEAN_", "CERTIFICATE_", "SLIP_CLEAN_"))

    print(f"\n{'=' * 55}")
    print("  [OK] Selesai!")
    print(f"  Total baris output  : {len(df):,}")
    print(f"  Total kolom output  : {len(df.columns)}")
    print(f"  File disimpan di    : {output_file}")
    print(f"{'=' * 55}")


if __name__ == "__main__":
    process_data(INPUT_FILE, OUTPUT_FILE)
    
"""
Cleaning Data 2 - Osbal | Cedant: KBRU (PT. KBRU REINSURANCE BROKERS
(EX. SIMAS))

CATATAN VERSI: v5 -- rewrite setelah bug report ke-4 (15/09/26, file:
KBRU_BUG_150926.md). Mesin POLIS/SLIP & INSURED SAMA PERSIS dgn
KBRU_cleaning_suspend.py v5 (lihat docstring file itu utk daftar lengkap
16 poin perubahan). Logika seleksi sumber SLIP_CLEAN di file ini tidak
berubah dari v4.

CATATAN VERSI (v4) -- bug report ke-3 (14/09/26, file:
KBRU_BUG_140926.md). Perubahan mesin POLIS/SLIP & INSURED SAMA PERSIS
dgn KBRU_cleaning_suspend.py v4 (lihat docstring file itu utk daftar
lengkap). Tambahan KHUSUS Osbal di versi ini:
  6. [FIX] Regex "sesuai karakteristik" (_matches_slip_characteristic)
     diperlonggar utk toleransi kode cedant "KBRU" polos (tanpa "RE")
     dan segmen kode produk berawalan huruf (mis. "A98" pada
     "A98.514-KBRU"), serta panjang segmen akhir 3-7 digit (sebelumnya
     wajib 5-7, kelewat ketat utk kasus spt ini).
  7. [BARU] Fallback ke FAC_POLICY_NO (fallback ke-4) sekarang pakai
     pengecekan LONGGAR (_matches_loose_characteristic -- cukup ada
     kode cedant KBRURE/KBURE/KBRU/SRB + digit nyata di mana pun, tanpa
     syarat struktur segmen lengkap), BUKAN pengecekan ketat yg sama
     dgn CLSDT/FAC_SLIP. Alasan: di titik ini sumber lain sudah habis,
     jadi ada kode cedant + digit asli sudah cukup jadi indikasi nilai
     valid meski variasi formatnya tidak persis cocok pola baku.

CATATAN VERSI (v3) -- rewrite setelah bug report ke-2 (10/09/26). Mesin
cleaning karakter POLIS/SLIP & INSURED SAMA PERSIS dgn
KBRU_cleaning_suspend.py v3 (lihat docstring file itu utk daftar
lengkap perubahan mesin & contoh). File ini sendiri MENDAPAT TAMBAHAN
FIX KHUSUS di logika SELEKSI SUMBER SLIP_CLEAN (temuan dari 3 file bug
report: filter_slip_clean_1_yang_masih_belu.md, Filter_clsdt_polis_
yang_sesuai_kara.md, filter_polis_clean_yang_sesuai_kara.md):

1. [FIX] "Sesuai karakteristik slip" (dipakai utk seleksi sumber)
   SEKARANG HANYA merujuk ke pola baku No Slip KBRU (format multi-segmen
   spt "NNNNN/XXXX/NNN.NNNNNNN-(KBRURE|SRB)/MM/YYYY", boleh beda
   pemisah . / -) -- pola khusus "NNNNNNNN[A-Z]NNNNNNNN-NN" (mis.
   "11055124F00000144-03") TIDAK LAGI dihitung sbg "sesuai karakteristik"
   utk keperluan SELEKSI SUMBER (meskipun nilainya sendiri tetap valid
   & "dibiarkan apa adanya" kalau akhirnya terpakai). Sebelumnya (v2)
   kedua pola dianggap setara utk seleksi sumber -- ini yang bikin
   CLSDT_SLIP_NO ber-pola F/S salah didahulukan drpd CLSDT_POLICY_NO yg
   sebenarnya lebih valid (format baku), sesuai temuan bug report.
2. [BARU] Fallback ke-4: kalau CLSDT_SLIP_NO, CLSDT_POLICY_NO, MAUPUN
   FAC_SLIP semuanya TIDAK sesuai karakteristik (kosong/placeholder/kata
   umum spt END, P1, P1+P3+P4+P5) -> coba FAC_POLICY_NO; kalau ITU
   sesuai karakteristik -> dipakai. Kalau tetap tidak -> fallback
   terakhir ke FAC_SLIP apa adanya (spt semula).

CATATAN ARSITEKTUR:
- POLIS_CLEAN: SELALU dari FAC_POLICY_NO saja (CLSDT_POLICY_NO tidak
  dipakai sbg sumber POLIS_CLEAN sama sekali -- ini TIDAK berubah dari
  v2).
- SLIP_CLEAN: CLSDT_SLIP_NO -> CLSDT_POLICY_NO -> FAC_SLIP -> FAC_POLICY_NO
  (masing2 dicek "sesuai karakteristik" secara berurutan, pakai definisi
  yg SUDAH DIPERBAIKI di poin 1; begitu ketemu yg sesuai, langsung
  dipakai & diproses lewat clean_value() yg sama).
- CERTIFICATE_1 SELALU KOSONG, posisinya TEPAT setelah POLIS_CLEAN_1.
  TIDAK ADA kolom CN_1 -- notasi "CN : ..." dibuang total oleh
  clean_value() (lihat KBRU_cleaning_suspend.py).
- INSURED: sumbernya kolom FAC_INSURED. Rule SAMA PERSIS dgn Suspend
  (lihat docstring file itu utk daftar lengkap: pola PELABUHAN
  INDONESIA/CLUSTER/TOWER/REGIONAL/DISTRICT, suffix UD/PD/PT/CV/gelar,
  delimiter QQQ/OO/*, parenthetical, treaty-keep, dst).

"""

import os
import re

import pandas as pd


# ─────────────────────────────────────────────────────────────────────────────
# KONFIGURASI
# ─────────────────────────────────────────────────────────────────────────────

INPUT_FILE  = os.path.join("input", "2b. transaksi Osbal 01.01.23 - 17.08.26.xlsx")
INPUT_SHEET = "Query result"
OUTPUT_FILE = os.path.join("output", "Kbru_output_osbal.xlsx")

CEDANT_COL   = "CCOS_COMP_NAME"
CEDANT_VALUE = "PT. KBRU REINSURANCE BROKERS (EX. SIMAS)"

POLIS_COL   = "FAC_POLICY_NO"
SLIP_COL    = "FAC_SLIP"
INSURED_COL = "FAC_INSURED"

CLSDT_POLIS_COL = "CLSDT_POLICY_NO"
CLSDT_SLIP_COL  = "CLSDT_SLIP_NO"

MAX_SPLIT_COLS = 5


# ─────────────────────────────────────────────────────────────────────────────
# HELPER UMUM
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
        return [", ".join(items)]
    return items


# ─────────────────────────────────────────────────────────────────────────────
# MESIN POLIS/SLIP KBRU (identik dgn KBRU_cleaning_suspend.py v3)
# ─────────────────────────────────────────────────────────────────────────────

_JUNK_STANDALONE_WORDS = {"TBA", "VAR", "VARIOUS", "SELISIH", "END", "-", "NAN", "NONE", "NULL"}
_P_CODE_RE = re.compile(r"^P\d+$", re.IGNORECASE)

_MONTH_RE = (
    r"JANUARI|FEBRUARI|MARET|APRIL|MEI|JUNI|JULI|AGUSTUS|SEPTEMBER|OKTOBER|"
    r"NOVEMBER|DES[EA]?MBER|JANUARY|FEBRUARY|MARCH|MAY|JUNE|JULY|AUGUST|OCTOBER|DECEMBER"
)
_MONTH_YEAR_TAIL_RE = re.compile(
    r"\s*/?\s*(?:" + _MONTH_RE + r")\.?\s*(?:19|20)?\d{0,4}\s*$", re.IGNORECASE
)

_CN_MARKER_STRIP_RE_UNUSED = None
_NO_ANYWHERE_RE = re.compile(r"\bNO\.\s*", re.IGNORECASE)
_LEADING_ALPHA_PLUS_RE = re.compile(r"^[^\d/]{2,}?\+\s*")
_TRAILING_JUNK_WORD_RE = re.compile(r"[\s/]+(?:TBA|VAR|VARIOUS)\s*$", re.IGNORECASE)

# [BARU 15/09] deskripsi murni SUSPENSE/UTANG PIUTANG yang nyasar di
# kolom polis/slip -- dibiarkan apa adanya (spasi TIDAK dihapus), sama
# spt perlakuan di kolom INSURED.
_SUSPENSE_LIKE_RE = re.compile(r"SUSPEN|\(?H\)?\s*UTANG\s*PIUTANG", re.IGNORECASE)

# [BARU 15/09] teks narasi MURNI (BORDERO/SUMMARY SIMAS RE/nama bulan/
# tahun/pemisah saja, TANPA angka lain) -- dibiarkan apa adanya.
_NARRATIVE_ONLY_RE = re.compile(
    r"^(?:BORDERO|SUMMARY|BULAN|SIMAS\s*RE|" + _MONTH_RE + r"|(?:19|20)\d{2}|[-\s,])+$",
    re.IGNORECASE,
)

# [BARU] rantai P<n> (+ anotasi opsional "(ADD n)") yang digabung "+" --
# dibuang SELURUHNYA (termasuk "+"-nya), di mana pun posisinya (bukan
# cuma di akhir).
_STRAY_P_CHAIN_RE = re.compile(
    r"(?:\bP\d+\b\s*(?:\(ADD\s*\d+\))?\s*\+\s*)+", re.IGNORECASE
)
# [BARU] rantai P<n> yang dipisah SPASI (bukan "+") yang nempel di
# BELAKANG nilai -- dibuang.
_TRAILING_P_CHAIN_RE = re.compile(r"(?:\s+P\d+\b)+\s*$", re.IGNORECASE)

# [BARU] segmen pertama (sebelum "/" pertama) yang berisi RANTAI angka
# pendek digabung "-" (bukan "+") sebanyak >=5 kali -- dianggap
# 'pengulangan berlebih', DIBIARKAN APA ADANYA (raw), sama seperti
# aturan pengulangan >5 via "+".
def _has_excessive_dash_chain(value: str) -> bool:
    first_seg = value.split("/", 1)[0]
    return first_seg.count("-") >= 5


KEEP_AS_IS_PATTERN_RE = re.compile(r"^\d{7,10}[A-Z]{1,2}\d{6,9}-\d{2}$", re.IGNORECASE)


def _is_junk_token(tok: str) -> bool:
    t = tok.strip().upper()
    if not t:
        return True
    if t in _JUNK_STANDALONE_WORDS:
        return True
    if _P_CODE_RE.match(t):
        return True
    return False


def _value_has_real_digits(text: str) -> bool:
    return bool(re.search(r"\d{3,}", text))


def _has_junk_marker(text: str) -> bool:
    return bool(re.search(r"\bTBA\b|\bP\d+\b", text, re.IGNORECASE))


def _split_long_space_tail(val: str):
    parts = re.split(r"\s{2,}", val, maxsplit=1)
    if len(parts) == 1:
        return val, None
    return parts[0].strip(), parts[1].strip()


def _strip_leading_junk_wrapper(val: str):
    """Buang wrapper di depan via '/' SAMPAI ketemu nilai asli:
    - PASS 1: cari segmen PERTAMA yang punya digit -> itu awal nilai
      asli (buang semua sebelumnya, apa pun isinya -- nama cedant/broker
      mis. 'SINARMAS', placeholder TBA/P<n>, dst).
    - PASS 2 (kalau TIDAK ADA segmen berdigit sama sekali): hanya buang
      segmen placeholder BAKU (TBA/P<n>) di depan; segmen bukan-baku
      lain (mis. nama perusahaan tanpa digit) TETAP DIPERTAHANKAN kalau
      itu satu-satunya isi yang tersisa (mis. 'TBA/SINARMAS' -> 'SINARMAS').
    Kalau SEMUA segmen placeholder baku -> None (murni placeholder)."""
    if "/" not in val:
        return None if _is_junk_token(val) else val
    segs = val.split("/")
    for i, s in enumerate(segs):
        if re.search(r"\d", s):
            return "/".join(segs[i:]).strip()
    start_idx = None
    for i, s in enumerate(segs):
        if not _is_junk_token(s.strip()):
            start_idx = i
            break
    if start_idx is None:
        return None
    return "/".join(segs[start_idx:]).strip()


def _strip_trailing_junk_wrapper(val: str) -> str:
    """Buang segmen TANPA digit (nama cedant/broker/kode singkat, mis.
    'ACA', 'SINAR MAS') yang nempel di BELAKANG via '/', SAMPAI ketemu
    segmen yang punya digit (dianggap bagian nomor asli). BEDA dgn
    _strip_leading_junk_wrapper (berbasis _is_junk_token) -- di
    belakang, kriterianya 'ada digit atau tidak', krn tag broker/cedant
    yg nempel di ekor biasanya bukan kode baku (TBA/P<n>) tapi nama
    bebas apa saja (mis. 'ACA')."""
    if "/" not in val:
        return val
    segs = val.split("/")
    end_idx = None
    for i in range(len(segs) - 1, -1, -1):
        if re.search(r"\d", segs[i]):
            end_idx = i + 1
            break
    if end_idx is None:
        return val
    return "/".join(segs[:end_idx]).strip()


def _find_substitution_target(text: str):
    """Cari segmen digit yang jadi TARGET substitusi 'pengulangan'.
    Kasus khusus: pola 'ANGKA + 1-2 HURUF + ANGKA' PERSIS (mis.
    'S00000271' pada '110551216F00000473', tanpa karakter lain) ->
    target SEGMEN TERAKHIR (setelah huruf) -- meski segmen pertama
    (di depan huruf) sering lebih panjang, BUKAN itu yg dimaksud
    'pengulangan'. Selain pola ini -> segmen digit TERPANJANG; kalau
    ada BEBERAPA segmen dgn panjang SAMA (tie), pilih yg PERTAMA di
    antaranya yang BUKAN nol semua (mis. '0000-5747': '0000' dan '5747'
    sama panjang, tapi '0000' cuma nol -> target '5747')."""
    m = re.match(r"^(\d+)([A-Za-z]{1,2})(\d+)$", text)
    if m:
        return m.group(3), m.start(3), m.end(3)
    candidates = list(re.finditer(r"\d+", text))
    if not candidates:
        return None, None, None
    max_len = max(len(c.group(0)) for c in candidates)
    tied = [c for c in candidates if len(c.group(0)) == max_len]
    non_zero = [c for c in tied if set(c.group(0)) != {"0"}]
    best = non_zero[0] if non_zero else tied[0]
    return best.group(0), best.start(), best.end()


def _qualifies_for_substitution(main: str) -> bool:
    """Nilai utama cukup 'berstruktur' utk dijadikan target substitusi
    'pengulangan' kalau: ada huruf (pola FPAR-style), ATAU >=2 pemisah
    '.'/'/' (pola dot-style spt '12.100.0001.82288'), ATAU total digitnya
    panjang (>=10, pola digit polos panjang spt '0124012000618'), ATAU
    ada >=1 pemisah '.'/'/'/'-' DAN total digit >=7 (pola spt '0000-
    5747', beda dari angka pendek polos TANPA pemisah spt '3522' yang
    TETAP tidak memenuhi)."""
    if re.search(r"[A-Za-z]", main):
        return True
    sep_count = len(re.findall(r"[./]", main))
    if sep_count >= 2:
        return True
    digit_len = len(re.sub(r"\D", "", main))
    if digit_len >= 10:
        return True
    if len(re.findall(r"[.\-/]", main)) >= 1 and digit_len >= 7:
        return True
    return False


def _apply_segment_substitution(primary: str, extra_tokens: list) -> list:
    seg, start, end = _find_substitution_target(primary)
    if seg is None:
        return [primary]
    prefix, suffix = primary[:start], primary[end:]
    out = [primary]
    for tok in extra_tokens:
        tok = tok.strip()
        if not tok.isdigit():
            continue
        if len(tok) >= len(seg):
            new_seg = tok
        else:
            new_seg = seg[: -len(tok)] + tok
        out.append(prefix + new_seg + suffix)
    return _dedup_preserve_order(out)


# [BARU 15/09] dash NYASAR di dalam pola F/S 2-segmen (mis.
# '11055129F00000-275' -- seharusnya '11055129F00000275') -- dash
# dianggap noise/typo, dilebur.
_FS_STRAY_DASH_RE = re.compile(r"^(\d+[A-Za-z]{1,2}\d+)-(\d+)$")


def _merge_fs_stray_dash(token: str) -> str:
    m = _FS_STRAY_DASH_RE.match(token)
    if m:
        return m.group(1) + m.group(2)
    return token


_CN_TRAILING_RE = re.compile(r"\s*/?\s*CN\s*:\s*.*$", re.IGNORECASE)
_CN_LEADING_WITH_SLASH_RE = re.compile(r"^\s*CN\s*:\s*[^/]*/\s*", re.IGNORECASE)
_CN_LABEL_ONLY_RE = re.compile(r"^\s*CN\s*:\s*", re.IGNORECASE)
_CN_STARTS_RE = re.compile(r"^CN\s*:", re.IGNORECASE)


def _strip_cn_suffix(val: str) -> str:
    """Notasi 'CN : <nomor>':
    - Muncul di TENGAH/BELAKANG ('main / CN : cn') -> dibuang total
      bersama nilainya, sisakan main di depan.
    - Muncul di AWAL DIIKUTI nilai lain setelah '/' ('CN : cn / main')
      -> label+nilai CN dibuang, sisakan main setelah '/'.
    - HANYA 'CN : cn' saja (seluruh value, tanpa '/' sama sekali) ->
      cuma label 'CN :'-nya yg dibuang, nilai CN dipertahankan sbg
      hasil utama (BUKAN dibuang semua)."""
    stripped = val.strip()
    if not _CN_STARTS_RE.match(stripped):
        return _CN_TRAILING_RE.sub("", val).strip(" /-")
    m = _CN_LEADING_WITH_SLASH_RE.match(val)
    if m:
        return val[m.end():].strip(" /-")
    return _CN_LABEL_ONLY_RE.sub("", val).strip(" /-")


_EMBEDDED_SUFFIX_RE = re.compile(r"^(\d+)([./].*[A-Za-z].*)$")


def _split_embedded_suffix(token: str):
    """Deteksi token yg formatnya 'ANGKA' + 'SUFFIX BERSTRUKTUR' nempel
    jadi satu (mis. '23022/FIAR/001.0000344-SRB/06/2022' -> ('23022',
    '/FIAR/001.0000344-SRB/06/2022')). Dipakai utk kasus 'pengulangan'
    di mana suffix format cuma nempel di SALAH SATU token dlm rantai
    '+' (biasanya yg terakhir), bukan di token pertama."""
    m = _EMBEDDED_SUFFIX_RE.match(token)
    if m:
        return m.group(1), m.group(2)
    return token, None


def _strip_all_characters(value: str) -> str:
    return re.sub(r"[^A-Za-z0-9]", "", value)


_SINAR_MAS_RE = re.compile(r"\bSINAR\s+MAS\b", re.IGNORECASE)
_TEXT_PIECE_PTCV_RE = re.compile(r",?\s*\b(?:P\s*T|CV)\b\.?\s*$", re.IGNORECASE)


def _clean_text_piece(value: str) -> str:
    """Utk hasil yang ternyata berupa TEKS/NAMA (tanpa digit sama
    sekali, muncul lewat mekanisme buang TBA/P<n> yang bersandingan) --
    JANGAN di-'hapus karakter' penuh (spasi antar kata dipertahankan,
    beda dari nilai numerik biasa), cuma normalisasi spasi + standardisasi
    'SINAR MAS' -> 'SINARMAS' (konsisten dgn kolom INSURED) + buang
    suffix "PT"/"CV" di akhir (mis. "ASURANSI FPG INSURANCE, PT" ->
    "ASURANSI FPG INSURANCE")."""
    value = _SINAR_MAS_RE.sub("SINARMAS", value)
    value = _TEXT_PIECE_PTCV_RE.sub("", value)
    return _normalize_spaces(value).upper()


def _simplify_noletter_extra(e: str) -> str:
    """[BARU 15/09] Token tambahan (di rantai '+') yang TIDAK punya
    huruf tapi masih ada '/' atau '.' nempel (mis. '04820/0001189' dari
    '...+04820/0001189.') -- ambil cuma potongan PERTAMA sbg kandidat
    substitusi digit, sisanya (mis. sertifikat/nomor duplikat) dibuang
    (KBRU tidak punya kolom sertifikat, jadi info ini tidak tertampung)."""
    if re.search(r"[A-Za-z]", e):
        return e
    return re.split(r"[./]", e)[0]


def clean_value(val) -> list:
    if pd.isna(val):
        return []
    raw = str(val).strip()
    if not raw:
        return []

    if raw.upper() in _JUNK_STANDALONE_WORDS:
        return [raw]
    # [BARU 15/09] "TBA/" (placeholder + garis miring doang, tanpa isi
    # lain) -> garis miringnya dibuang, cukup "TBA".
    raw_no_slash = raw.strip("/ ")
    if raw_no_slash.upper() in _JUNK_STANDALONE_WORDS:
        return [raw_no_slash]
    if KEEP_AS_IS_PATTERN_RE.match(raw):
        return [raw]
    if re.fullmatch(r"\d+&\d+", raw):
        return [raw]
    if _has_excessive_dash_chain(raw):
        # segmen depan berisi >=5 angka digabung '-' -- pengulangan
        # berlebih, dibiarkan apa adanya (sama spt aturan >5 via '+')
        return [raw]
    if _SUSPENSE_LIKE_RE.search(raw) or _NARRATIVE_ONLY_RE.match(raw.strip()):
        # [BARU 15/09] deskripsi SUSPENSE/UTANG PIUTANG, atau narasi
        # murni (BORDERO/bulan/tahun/SUMMARY SIMAS RE) yg nyasar di
        # kolom polis/slip -- dibiarkan apa adanya, spasi TIDAK dihapus
        return [raw]

    if not _has_junk_marker(raw) and not _value_has_real_digits(raw):
        # teks tanpa placeholder (TBA/P<n>) & tanpa digit sama sekali
        # (mis. nama perusahaan nyasar, "ASURANSI SINAR MAS") -> tetap
        # "abaikan" (bukan diproses spt polis/slip), TAPI tetap lewat
        # normalisasi teks (spasi + SINAR MAS->SINARMAS) spy konsisten
        # dgn kasus yg lewat jalur TBA-adjacent.
        return [_clean_text_piece(raw)]

    val2 = _NO_ANYWHERE_RE.sub("", raw)
    val2 = _strip_cn_suffix(val2)
    val2 = _LEADING_ALPHA_PLUS_RE.sub("", val2)
    val2 = _STRAY_P_CHAIN_RE.sub("", val2)
    val2 = _TRAILING_P_CHAIN_RE.sub("", val2)

    pieces = []
    stripped = _strip_leading_junk_wrapper(val2)
    if stripped is not None:
        val2 = stripped
        val2 = _TRAILING_P_CHAIN_RE.sub("", val2)
        primary, _tail = _split_long_space_tail(val2)
        primary2 = _MONTH_YEAR_TAIL_RE.sub("", primary).strip(" -/,")
        primary = primary2 if primary2 else primary
        primary = _TRAILING_JUNK_WORD_RE.sub("", primary).strip()

        raw_parts = [p.strip() for p in re.split(r"\s*\+\s*", primary) if p.strip()]
        raw_parts = [_merge_fs_stray_dash(p) for p in raw_parts]
        # [FIX 15/09] "_strip_trailing_junk_wrapper" cuma dipakai kalau
        # MEMANG ada rantai '+' (>1 bagian) -- kalau cuma 1 bagian
        # (whole value, mis. hasil gabungan lewat "No." di tengah),
        # JANGAN dipangkas, supaya ekor sah spt "/AB"/"/BB" (kode akhir
        # 2 huruf tanpa digit) tidak ikut kebuang.
        apply_trailing_strip = len(raw_parts) > 1
        parts = []
        for p in raw_parts:
            if _is_junk_token(p):
                continue
            p2 = _strip_trailing_junk_wrapper(p) if apply_trailing_strip else p
            p2 = p2.strip()
            if p2 and not _is_junk_token(p2):
                parts.append(p2)

        if parts:
            if len(parts) > 1:
                # kalau ADA campuran (sebagian bagian punya digit,
                # sebagian tidak sama sekali, mis. 'UMUM MEGA' vs
                # '2010208012400001') -> bagian tanpa digit dibuang
                # (nama cedant/broker yg nyempil, bukan bagian nilai).
                digit_parts = [p for p in parts if re.search(r"\d", p)]
                if digit_parts:
                    parts = digit_parts

            # [CEK 1] suffix berstruktur nempel di SALAH SATU token
            # (biasanya token TERAKHIR, mis. '23018+23019+23023+23022/
            # FIAR/001.0000344-SRB/06/2022') -- kalau ketemu & SISANYA
            # semua digit polos, terapkan suffix yg sama ke semua nomor.
            suffix_found = None
            base_numbers = []
            embed_ok = True
            for p in parts:
                num_part, suf = _split_embedded_suffix(p)
                if suf:
                    if suffix_found is None:
                        suffix_found = suf
                    else:
                        embed_ok = False
                if not num_part.isdigit():
                    embed_ok = False
                base_numbers.append(num_part)

            if embed_ok and suffix_found and len(base_numbers) > 1:
                ref = max(base_numbers, key=len)
                combined = []
                for n in base_numbers:
                    if len(n) >= len(ref):
                        combined.append(n + suffix_found)
                    else:
                        combined.append(ref[: -len(n)] + n + suffix_found)
                pieces = _dedup_preserve_order(combined)
            else:
                # [CEK 2] main = bagian yg PALING berstruktur (ada digit
                # DAN (huruf ATAU >=2 pemisah)), di mana pun posisinya
                # dalam rantai '+' -- sisanya jadi kandidat token
                # substitusi/nilai lain.
                structured_idx = None
                for i, p in enumerate(parts):
                    if re.search(r"\d", p) and (
                        re.search(r"[A-Za-z]", p) or len(re.findall(r"[./]", p)) >= 2
                    ):
                        structured_idx = i
                        break
                if structured_idx is not None:
                    main = parts[structured_idx]
                    extras_raw = parts[:structured_idx] + parts[structured_idx + 1:]
                else:
                    main = parts[0]
                    extras_raw = parts[1:]

                extras_clean = [_simplify_noletter_extra(e.strip().rstrip(".").strip()) for e in extras_raw]
                whole_extras = [e for e in extras_clean if "/" in e or re.search(r"\d\.\d", e)]
                digit_extras = [e for e in extras_clean if e.isdigit()]

                # [BARU 15/09] token substitusi yg SANGAT pendek (<=2
                # digit, mis. "+97+07+43") -- pola pengulangannya tidak
                # jelas (berapa digit ekor yg diganti jadi ambigu) ->
                # seluruh nilai dibiarkan apa adanya.
                if digit_extras and any(len(e) < 3 for e in digit_extras):
                    pieces = []
                elif digit_extras and not _qualifies_for_substitution(main):
                    pieces = []  # ambigu -> jatuh ke abaikan (raw) di bawah
                elif digit_extras:
                    pieces = _apply_segment_substitution(main, digit_extras)
                    pieces += whole_extras
                else:
                    pieces = [main] + whole_extras

    pieces = _dedup_preserve_order([_normalize_spaces(p) for p in pieces if p.strip()])

    if not pieces or len(pieces) > MAX_SPLIT_COLS:
        return [raw]

    final = []
    for p in pieces:
        if KEEP_AS_IS_PATTERN_RE.match(p):
            final.append(p)
        elif not re.search(r"\d", p):
            final.append(_clean_text_piece(p))
        else:
            final.append(_strip_all_characters(p))
    return final


# ─────────────────────────────────────────────────────────────────────────────
# [KHUSUS DATA 2] SUMBER POLIS & SLIP
# ─────────────────────────────────────────────────────────────────────────────

# pola BAKU "No Slip" KBRU (sheet Karakteristik_Polis_dan_Slip.xlsx),
# toleran beda karakter pemisah (. / -) DAN toleran typo "KBURE" (tanpa
# huruf R kedua) & kode cedant polos "KBRU" (tanpa "RE") yang ditemukan
# di data mentah selain "KBRURE" baku. Pola khusus "NNNNNNNN[A-Z]
# NNNNNNNN-NN" (KEEP_AS_IS_PATTERN_RE) SENGAJA TIDAK termasuk di sini
# -- itu bukan "karakteristik" baku, cuma nilai yg tetap dipertahankan
# apa adanya KALAU akhirnya terpilih/dipakai.
_SLIP_CHARACTERISTIC_RE = re.compile(
    r"\d{3,9}[./-]{1,2}[A-Z]{3,5}[./-]{1,2}[A-Z]?\d{2,4}[./-]\d{3,7}[-./]{1,2}(?:KBRURE|KBURE|KBRU|SRB)\b",
    re.IGNORECASE,
)

# versi LONGGAR -- cuma cek ada kode cedant (KBRURE/KBURE/KBRU/SRB) +
# digit nyata di mana pun, TANPA syarat struktur segmen lengkap. Dipakai
# KHUSUS utk fallback terakhir ke FAC_POLICY_NO -- pada titik itu kita
# sudah kehabisan sumber lain, jadi ada kode cedant + digit asli sudah
# cukup jadi indikasi nilai valid, meski variasi formatnya (mis.
# "A98.514-KBRU") tidak persis cocok pola baku.
_LOOSE_CEDANT_CODE_RE = re.compile(r"(?:KBRURE|KBURE|KBRU|SRB)\b", re.IGNORECASE)


def _matches_slip_characteristic(value) -> bool:
    if pd.isna(value):
        return False
    v = str(value).strip()
    if not v:
        return False
    return bool(_SLIP_CHARACTERISTIC_RE.search(v))


def _matches_loose_characteristic(value) -> bool:
    if pd.isna(value):
        return False
    v = str(value).strip()
    if not v:
        return False
    return bool(_LOOSE_CEDANT_CODE_RE.search(v)) and _value_has_real_digits(v)


def select_slip_source(fac_slip, clsdt_slip, clsdt_polis, fac_polis):
    """SLIP_CLEAN: CLSDT_SLIP_NO -> CLSDT_POLICY_NO -> FAC_SLIP (dicek
    KETAT) -> FAC_POLICY_NO (dicek LONGGAR, fallback terakhir),
    berhenti di sumber PERTAMA yg sesuai."""
    for candidate in (clsdt_slip, clsdt_polis, fac_slip):
        if _matches_slip_characteristic(candidate):
            return clean_value(candidate)
    if _matches_loose_characteristic(fac_polis):
        return clean_value(fac_polis)
    return clean_value(fac_slip)


def select_polis_source(fac_polis):
    """POLIS_CLEAN: SELALU dari FAC_POLICY_NO saja."""
    return clean_value(fac_polis)


# ─────────────────────────────────────────────────────────────────────────────
# CLEAN INSURED (identik dgn KBRU_cleaning_suspend.py v3)
# ─────────────────────────────────────────────────────────────────────────────

_GELAR_PREFIX_RE = re.compile(
    r"^\s*(?:BAPAK|IBU|SDRI?|NYONYA|NY|TUAN|TN|MR|MRS|MS|DR|IR|PROF|HJ|H)\b\.?\s*",
    re.IGNORECASE,
)
_GELAR_SUFFIX_RE = re.compile(
    r"\s*,?\s*\b(?:S\.?E|S\.?H|S\.?SI|S\.?T|S\.?SOS|S\.?KOM|S\.?PD|"
    r"M\.?SI|M\.?M|M\.?H|M\.?KOM|SH|SE|BP|MRS|MR|NY|TN|DRS)\b\.?\s*$",
    re.IGNORECASE,
)
_SUSPENSE_KEEP_RE = re.compile(r"SUSPEN|\(?H\)?\s*UTANG\s*PIUTANG", re.IGNORECASE)
_TREATY_KEEP_RE = re.compile(
    r"\bTREATY\b|\bSTRUCTURE\b|EXCESS\s+OF\s+LOSS|WHOLE\s+ACCOUNT", re.IGNORECASE
)


def _is_keep_as_is_insured(val: str) -> bool:
    if _SUSPENSE_KEEP_RE.search(val):
        return True
    if _TREATY_KEEP_RE.search(val):
        return True
    return False


def _collapse_duplicate_words(text: str) -> str:
    text = re.sub(r"\(\s*\(", "(", text)
    text = re.sub(r"\)\s*\)", ")", text)
    text = re.sub(r"\b(\w{3,})\1\b", r"\1", text, flags=re.IGNORECASE)
    text = re.sub(r"\b(\w+)(?:\s+\1\b)+", r"\1", text, flags=re.IGNORECASE)
    if text.count("(") > text.count(")"):
        text = text + ")"
    return text


_KEEP_SUFFIX_COMMA_RE = re.compile(r",\s*(UD|PD)\b", re.IGNORECASE)
_REMOVE_COMMA_PTCV_RE = re.compile(r",\s*(?:P\s*T|CV)\b\.?", re.IGNORECASE)

_SPLIT_RE = re.compile(
    r"\bQ{2,3}\b|\bOO\b|\bAND\s*/\s*OR\.?\b|/{1,2}|,|–|¿|\*|&\s*(?=(?:PT|CV)\b)",
    re.IGNORECASE,
)

INSURED_JUNK_WORDS = frozenset({
    "PT", "CV", "TBK", "PERSERO", "LTD", "INC", "LLC",
    "AND", "OR", "THE", "OF", "AS", "VARIOUS", "VAR", "QQ", "TBA",
    "ASSOCIATED",
})

_MONTH_STANDALONE_RE = re.compile(
    r"^(?:" + _MONTH_RE + r")\.?\s*(?:19|20)?\d{2,4}$", re.IGNORECASE
)

_GELAR_PREFIX_RE = re.compile(
    r"^\s*(?:BAPAK|IBU|SDRI?|NYONYA|NY|TUAN|TN|MR|MRS|MS|DR|IR|PROF|HJ|H)\b\.?\s*",
    re.IGNORECASE,
)
# [BARU] tambah BP/MR/MRS/NY/TN/DRS sbg gelar SUFFIX (nempel via koma di
# BELAKANG nama, mis. "SUTJITA, BP" / "RUDY AKILI, MR") -- dibuang.
_GELAR_SUFFIX_RE = re.compile(
    r"\s*,?\s*\b(?:S\.?E|S\.?H|S\.?SI|S\.?T|S\.?SOS|S\.?KOM|S\.?PD|"
    r"M\.?SI|M\.?M|M\.?H|M\.?KOM|SH|SE|BP|MRS|MR|NY|TN|DRS)\b\.?\s*$",
    re.IGNORECASE,
)
_SUSPENSE_KEEP_RE = re.compile(r"SUSPEN|\(?H\)?\s*UTANG\s*PIUTANG", re.IGNORECASE)
_TREATY_KEEP_RE = re.compile(
    r"\bTREATY\b|\bSTRUCTURE\b|EXCESS\s+OF\s+LOSS|WHOLE\s+ACCOUNT", re.IGNORECASE
)


def _is_keep_as_is_insured(val: str) -> bool:
    if _SUSPENSE_KEEP_RE.search(val):
        return True
    if _TREATY_KEEP_RE.search(val):
        return True
    if _MALFORMED_QQ_RE.search(val):
        return True
    return False


# [BARU] normalisasi data rusak: kata terduplikasi persis berturut-turut
# (mis. "WASKITA WASKITA KARYA KARYA" -> "WASKITA KARYA"), kata gandeng
# tanpa spasi ("PERSEROPERSERO" -> "PERSERO"), kurung ganda/tak seimbang.
def _collapse_duplicate_words(text: str) -> str:
    text = re.sub(r"\(\s*\(", "(", text)
    text = re.sub(r"\)\s*\)", ")", text)
    text = re.sub(r"\b(\w{3,})\1\b", r"\1", text, flags=re.IGNORECASE)
    text = re.sub(r"\b(\w+)(?:\s+\1\b)+", r"\1", text, flags=re.IGNORECASE)
    if text.count("(") > text.count(")"):
        text = text + ")"
    return text


# [BARU] ", PT" / ", CV" (di mana pun, bukan cuma di akhir) DIBUANG
# TOTAL bersama komanya -- supaya tidak memicu split palsu saat ada isi
# lain menyertai PT/CV di segmen yang sama (mis. "SMAILING TOUR, PT &
# TRAVEL SERVICE, PT" -> "SMAILING TOUR & TRAVEL SERVICE", bukan
# terpecah jadi "SMAILING TOUR" + "& TRAVEL SERVICE").
_REMOVE_COMMA_PTCV_RE = re.compile(r",\s*(?:P\s*T|CV)\b\.?", re.IGNORECASE)

# ", UD" / ", PD" (jenis badan usaha) TIDAK dianggap pemisah entitas
# baru -- koma dibuang, kode dipertahankan menyatu dgn nama sebelumnya.
_KEEP_SUFFIX_COMMA_RE = re.compile(r",\s*(UD|PD)\b", re.IGNORECASE)

# pemisah: QQ/QQQ, OO (typo QQ), AND/OR, "/", "//", ",", en-dash, "¿",
# "*", dan "&" HANYA kalau diikuti PT/CV (JV baru, mis. "... & PT OKI
# PULP..."). "&" polos (mis. "PULP & PAPER") TETAP bukan pemisah. "+"
# TIDAK dipakai sbg pemisah (mis. "SK+SUMATERA" -> 1 entitas "SK
# SUMATERA", bukan 2) -- beda dgn mesin polis/slip. " - " (spasi-dash-
# spasi) juga TIDAK dipakai (bikin false split di nama spt "PULAU -
# PULAU MEDIA").
_SPLIT_RE = re.compile(
    r"\bQ{2,3}\b|\bOO\b|\bAND\s*/\s*OR\.?\b|/{1,2}|,|–|¿|\*|&\s*(?=(?:PT|CV)\b)",
    re.IGNORECASE,
)

INSURED_JUNK_WORDS = frozenset({
    "PT", "CV", "TBK", "PERSERO", "LTD", "INC", "LLC",
    "AND", "OR", "THE", "OF", "AS", "VARIOUS", "VAR", "QQ", "TBA",
    "ASSOCIATED",
})

_OWNER_OF_RE = re.compile(
    r"\(\s*AS\s+(?:THE\s+)?(?:OWNER|PRINCIPAL)\s+OF\s+([^()]+)\)", re.IGNORECASE
)

_ROLE_KEYWORDS_RE = re.compile(
    r"^(?:PERSERO|PRINCIPAL|OWNER|INSURED|CONTRACTOR)$", re.IGNORECASE
)

# [DIPERLONGGAR] toleransi "P T" (spasi nyempil di antara P & T, data
# rusak) selain "PT" biasa.
_PT_CV_PREFIX_RE = re.compile(r"^\s*(?:P\s*T|CV)(?:\.|(?=\s|$))\.?\s*", re.IGNORECASE)
_PT_CV_SUFFIX_RE = re.compile(r",?\s*(?:P\s*T|CV)\.?\s*$", re.IGNORECASE)

_MASK_COMMA = "\x01"
_MASK_SLASH = "\x02"


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
    words = [w for w in re.split(r"[\s\-]+", long_name.upper()) if w]
    if any(w.startswith(short_clean) and w != short_clean for w in words):
        return True
    return False


def _extract_owner_of_entities(text: str):
    extra = []

    def repl(m):
        extra.append(m.group(1).strip())
        return " "

    new_text = _OWNER_OF_RE.sub(repl, text)
    return new_text, extra


def _strip_parenthetical(text: str) -> str:
    """Isi dalam kurung yang singkatan dari teks sebelumnya, ATAU kata
    boilerplate peran (PERSERO/PRINCIPAL/OWNER/INSURED/CONTRACTOR), ATAU
    mengandung PT/CV -> kurung + isi DIBUANG total. Isi kurung yang
    mengandung KOMA (daftar deskriptif, mis. "(REFINERY, PLANTATION,
    KEMITRAAN)") -> kurung DIPERTAHANKAN UTUH (literal, koma di-mask
    dulu spy tdk ikut split). Isi kurung lain (identitas cabang/pihak
    lain, mis. "(KSO)") -> kurung dibuang, ISI dipertahankan."""

    def repl(m):
        inner = m.group(1).strip()
        if re.search(r"\bPT\b|\bCV\b", inner, re.IGNORECASE):
            return " "
        if _ROLE_KEYWORDS_RE.match(inner):
            return " "
        inner_letters = re.sub(r"[^A-Za-z]", "", inner)
        before = text[: m.start()].strip()
        if before and inner_letters and _is_abbreviation_of(inner_letters, before):
            return " "
        if "," in inner:
            masked = inner.replace(",", _MASK_COMMA).replace("/", _MASK_SLASH)
            return f" ({masked}) "
        masked = inner.replace(",", _MASK_COMMA).replace("/", _MASK_SLASH)
        return f" {masked} "

    return re.sub(r"\(([^()]*)\)", repl, text)


# ─────────────────────────────────────────────────────────────────────────────
# [BARU] EKSPANSI POLA "KEYWORD + DAFTAR" (FACTORY/CLUSTER/TOWER/
# REGIONAL/DISTRICT) & KHUSUS PELABUHAN INDONESIA/PELINDO
# ─────────────────────────────────────────────────────────────────────────────

_ROMAN_MAP = {"I": 1, "II": 2, "III": 3, "IV": 4, "V": 5, "VI": 6, "VII": 7, "VIII": 8, "IX": 9, "X": 10}
_ROMAN_ARABIC_ALT = r"I|II|III|IV|V|VI|VII|VIII|IX|X|\d+"

_PELINDO_RE = re.compile(
    r"^(PELABUHAN\s+INDONESIA)\s+((?:" + _ROMAN_ARABIC_ALT + r")"
    r"(?:\s*(?:,|&|DAN)\s*(?:" + _ROMAN_ARABIC_ALT + r"))+)\s*$",
    re.IGNORECASE,
)

# [BARU 15/09] "PELINDO" adalah SINONIM "PELABUHAN INDONESIA" -- bisa
# disebut BERULANG-ULANG dgn pemisah apa saja (biasanya QQ), mis.
# "PELABUHAN INDONESIA I QQ PELINDO I QQ PELINDO II QQ PELINDO III QQ
# PELINDO IV" -> semua kemunculan "PELABUHAN INDONESIA <n>"/"PELINDO <n>"
# dikumpulkan, dinormalisasi ke "PELABUHAN INDONESIA <n>" (arab), dedup.
_PELINDO_MENTION_RE = re.compile(
    r"\b(?:PELABUHAN\s+INDONESIA|PELINDO)\s+(" + _ROMAN_ARABIC_ALT + r")\b",
    re.IGNORECASE,
)


def _expand_pelindo(val: str):
    m = _PELINDO_RE.match(val.strip())
    if not m:
        return None
    prefix, nums_part = m.groups()
    tokens = [t.strip() for t in re.split(r"\s*(?:,|&|DAN)\s*", nums_part, flags=re.IGNORECASE) if t.strip()]
    out = []
    for t in tokens:
        num = _ROMAN_MAP.get(t.upper(), t)
        out.append(f"{prefix.upper()} {num}")
    return out


def _expand_pelindo_mentions(val: str):
    matches = _PELINDO_MENTION_RE.findall(val)
    if len(matches) < 2:
        return None
    seen = []
    for t in matches:
        num = _ROMAN_MAP.get(t.upper(), t)
        entry = f"PELABUHAN INDONESIA {num}"
        if entry not in seen:
            seen.append(entry)
    return seen if len(seen) >= 2 else None


_KEYWORD_MULTI_RE = re.compile(
    r"^([^/]*?)[/\s]*\b(FACTORY|CLUSTER|TOWER|REGIONAL|DISTRICT|UNIT|MALL)\s+([^/]+)$",
    re.IGNORECASE,
)


def _expand_keyword_multi(val: str):
    """'PREFIX [/] KEYWORD tok1 tok2 ...' atau 'PREFIX [/] KEYWORD tok1,
    tok2 & tok3' -> daftar 'PREFIX KEYWORD tok'. Butuh >=2 token supaya
    tidak salah nembak segmen tunggal (yg ditangani _merge_keyword_single
    di tempat lain)."""
    m = _KEYWORD_MULTI_RE.match(val.strip())
    if not m:
        return None
    prefix, keyword, tail = m.groups()
    tokens = [t for t in re.split(r"\s*(?:,|&|DAN)\s*|\s+", tail.strip(), flags=re.IGNORECASE) if t.strip()]
    tokens = [t.strip() for t in tokens if re.fullmatch(r"[A-Za-z0-9]{1,4}", t.strip())]
    if len(tokens) < 2:
        return None
    prefix = prefix.strip().rstrip(",").strip()
    prefix = re.sub(r",?\s*(?:P\s*T|CV)\.?\s*$", "", prefix, flags=re.IGNORECASE).strip()
    return [f"{prefix} {keyword.upper()} {t.upper()}".strip() for t in tokens]


_KEYWORD_SINGLE_RE = re.compile(
    r"^(?:FACTORY|CLUSTER|TOWER|REGIONAL|DISTRICT|UNIT|MALL)\s+[A-Za-z0-9]{1,4}$", re.IGNORECASE
)

# [BARU 15/09] "QQ" yang NYASAR nempel langsung ke huruf lain tanpa
# spasi/pemisah (mis. "QQBALI", "INDAHQQ") -- tanda data rusak/typo yg
# terlalu ambigu utk diproses -> seluruh value dibiarkan apa adanya.
# Rantai Q murni ("QQQ") TIDAK dihitung rusak -- itu cuma varian QQ
# dgn Q berlebih, sudah ditangani _SPLIT_RE (\bQ{2,3}\b).
_MALFORMED_QQ_RE = re.compile(r"[A-PR-Za-pr-z]QQ|QQ[A-PR-Za-pr-z]")


def _clean_insured_name(name: str) -> str:
    name = _normalize_spaces(name)
    name = _GELAR_PREFIX_RE.sub("", name)
    for _ in range(3):
        stripped = _GELAR_SUFFIX_RE.sub("", name).strip()
        if stripped == name:
            break
        name = stripped
    name = re.sub(r",?\s*\(\s*PERSERO\s*\)\s*", " ", name, flags=re.IGNORECASE)
    name = re.sub(r"\bTBK\b\s*,\s*P\s*T\.?\s*", " ", name, flags=re.IGNORECASE)
    name = re.sub(r",?\s*\bTBK\b\s*", " ", name, flags=re.IGNORECASE)
    name = re.sub(r"\bLTD\b\.?\s*$", "", name, flags=re.IGNORECASE)
    name = re.sub(r"\bPTE\b\.?\s*", "", name, flags=re.IGNORECASE)
    name = _PT_CV_SUFFIX_RE.sub("", name)
    name = _PT_CV_PREFIX_RE.sub("", name)
    # [FIX 15/09] pola inisial huruf-tunggal-bertitik berturut-turut
    # (mis. "S.A.R.I") -> titik dibuang TANPA jadi spasi ("SARI"),
    # BUKAN "S A R I". Titik lain (bukan pola ini, bukan antar digit)
    # baru diganti spasi spt biasa.
    name = re.sub(r"\b(?:[A-Za-z]\.){2,}[A-Za-z]?\b", lambda mo: mo.group(0).replace(".", ""), name)
    name = re.sub(r"(?<!\d)\.(?!\d)", " ", name)
    name = re.sub(r"\s*\+\s*", " ", name)
    name = _SINAR_MAS_RE.sub("SINARMAS", name)
    name = name.replace(_MASK_COMMA, ",").replace(_MASK_SLASH, "/")
    return _normalize_spaces(name).strip().upper()


def _try_expansions(segment: str):
    """Coba semua mekanisme ekspansi (pelindo list, pelindo scattered
    mentions, keyword+daftar) pada SATU string. None kalau tidak ada yg
    cocok."""
    r = _expand_pelindo(segment)
    if r is not None:
        return r
    r = _expand_pelindo_mentions(segment)
    if r is not None:
        return r
    r = _expand_keyword_multi(segment)
    if r is not None:
        return r
    return None


def clean_insured(val) -> list:
    if pd.isna(val):
        return []
    val = str(val).strip()
    if not val:
        return []

    val = _collapse_duplicate_words(val)

    if _is_keep_as_is_insured(val):
        return [_normalize_spaces(val).upper()]

    val, owner_extra = _extract_owner_of_entities(val)
    val = _strip_parenthetical(val)
    val = _KEEP_SUFFIX_COMMA_RE.sub(r" \1", val)
    val = _REMOVE_COMMA_PTCV_RE.sub("", val)

    # coba ekspansi pada SELURUH value dulu (kasus umum: satu pola
    # keyword/pelindo mencakup semua, mis. "BIO FARMA/CLUSTER B C D",
    # "PELABUHAN INDONESIA I QQ PELINDO II ...").
    whole = _try_expansions(val)
    if whole is not None:
        return [_normalize_spaces(x).upper() for x in whole]

    parts = [p.strip() for p in _SPLIT_RE.split(val) if p.strip()]
    cleaned = []
    for p in parts:
        # [FIX 15/09] kalau SATU bagian hasil split (bukan whole value)
        # sendiri cocok pola keyword+daftar (mis. "GALAXY MALL I & II"
        # setelah "SINAR GALAXY /" ikut terpisah lebih dulu) -> expand
        # jadi beberapa entitas, JANGAN diproses sbg 1 entitas biasa.
        r = _expand_keyword_multi(p)
        if r is not None and len(r) >= 2:
            cleaned.extend(_normalize_spaces(x).upper() for x in r)
            continue
        p_clean = _clean_insured_name(p)
        if not p_clean:
            continue
        if _MONTH_STANDALONE_RE.match(p_clean):
            continue
        if _KEYWORD_SINGLE_RE.match(p_clean) and cleaned:
            cleaned[-1] = _normalize_spaces(cleaned[-1] + " " + p_clean)
            continue
        if len(p_clean) < 2 or p_clean in INSURED_JUNK_WORDS:
            continue
        cleaned.append(p_clean)

    for extra in owner_extra:
        extra_clean = _clean_insured_name(extra)
        if extra_clean and extra_clean not in cleaned:
            cleaned.append(extra_clean)

    if not cleaned:
        fallback = _clean_insured_name(val)
        return [fallback] if fallback else []

    return _cap_or_join(cleaned)


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


def _insert_polis_cert_columns(df, all_polis_lists, max_cols):
    added = []
    actual_max_cols = max((len(lst) for lst in all_polis_lists if isinstance(lst, list)), default=1)
    actual_max_cols = min(actual_max_cols, max_cols)
    for i in range(1, actual_max_cols + 1):
        polis_col = f"POLIS_CLEAN_{i}"
        df[polis_col] = [lst[i - 1] if isinstance(lst, list) and i - 1 < len(lst) else None for lst in all_polis_lists]
        added.append(polis_col)
        if i == 1:
            df["CERTIFICATE_1"] = None
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

    df.rename(columns={
        POLIS_COL:   "FAC_POLICY_NO_ORI",
        SLIP_COL:    "FAC_SLIP_ORI",
        INSURED_COL: "FAC_INSURED_ORI",
    }, inplace=True)

    print("[3/5] Menjalankan proses cleaning ...")

    all_polis, all_slip, all_insured = [], [], []
    max_polis = max_slip = max_ins = 1

    for _, row in df.iterrows():
        p_clean = select_polis_source(row.get("FAC_POLICY_NO_ORI", ""))
        s_clean = select_slip_source(
            row.get("FAC_SLIP_ORI", ""),
            row.get(CLSDT_SLIP_COL, ""),
            row.get(CLSDT_POLIS_COL, ""),
            row.get("FAC_POLICY_NO_ORI", ""),
        )
        c_ins = clean_insured(row.get("FAC_INSURED_ORI", ""))

        max_polis = max(max_polis, len(p_clean))
        max_slip = max(max_slip, len(s_clean))
        max_ins = max(max_ins, len(c_ins))

        all_polis.append(p_clean)
        all_slip.append(s_clean)
        all_insured.append(c_ins)

    print("[4/5] Menyusun kolom output ...")

    new_columns = []
    for col in df.columns:
        new_columns.append(col)
        if col == "FAC_POLICY_NO_ORI":
            new_columns += _insert_polis_cert_columns(df, all_polis, max_polis)
        elif col == "FAC_SLIP_ORI":
            new_columns += _insert_clean_columns(df, all_slip, "SLIP_CLEAN", max_slip)
        elif col == "FAC_INSURED_ORI":
            new_columns += _insert_clean_columns(df, all_insured, "INSURED_CLEAN", max_ins)

    df = df[new_columns]

    # FORMAT OUTPUT OSBAL: susunan dan nama kolom mengikuti file tabel osbal.xlsx.
    # Ini hanya pemetaan kolom output; nilai hasil cleaning tidak dihitung ulang.
    df = df.rename(columns={'FAC_POLICY_NO_ORI': 'FAC_POLICY_NO', 'FAC_SLIP_ORI': 'FAC_SLIP', 'FAC_INSURED_ORI': 'FAC_INSURED'})
    df = df.rename(columns={
        **{f"POLIS_CLEAN_{i}": f"POLICY_CLEAN_{i}" for i in range(1, 6)},
        **{f"CERTIFICATE_{i}": f"SERTIF_CLEAN_{i}" for i in range(1, 6)},
    })
    # Kolom dari template yang tidak ada di sumber dibuat kosong.
    # Susunan dan ejaan header diambil persis dari dua template yang diberikan.
    output_columns = ['CCOS_DOC_NO', 'CCOS_DATE', 'CCOS_REF_CODE', 'CCOS_COMP', 'CCOS_COMP_NAME', 'CCOS_REF_COMP', 'CCOS_REF_COMP_NAME', 'FAC_INSURED', 'INSURED_CLEAN_1', 'INSURED_CLEAN_2', 'INSURED_CLEAN_3', 'INSURED_CLEAN_4', 'INSURED_CLEAN_5', 'CCOS_CURR', 'CCOS_OR_BAL', 'CCOS_BAL_DUE', 'CCOS_OR_BAL_IN_IDR', 'CCOS_BAL_DUE_IN_IDR', 'FAC_COM_DATE', 'FAC_EXP_DATE', 'FAC_DUE_DATES', 'FAC_SUB_CLASS', 'FAC_POLICY_NO', 'POLICY_CLEAN_1', 'SERTIF_CLEAN_1', 'POLICY_CLEAN_2', 'SERTIF_CLEAN_2', 'POLICY_CLEAN_3', 'SERTIF_CLEAN_3', 'POLICY_CLEAN_4', 'POLICY_CLEAN_5', 'FAC_SLIP', 'SLIP_CLEAN_1', 'SLIP_CLEAN_2', 'SLIP_CLEAN_3', 'SLIP_CLEAN_4', 'SLIP_CLEAN_5', 'CLASS_CODE', 'CLASS_NAME', 'CLSDT_POLICY_NO', 'CLSDT_SLIP_NO', 'CLSDT_SERTF_NO']
    df = df.reindex(columns=output_columns)

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

"""
Cleaning Data 1 - Facultative | Cedant: KBRU (PT. KBRU REINSURANCE
BROKERS (EX. SIMAS))

CATATAN VERSI: v5 -- rewrite setelah bug report ke-4 (15/09/26, file:
KBRU_BUG_150926.md). Mesin POLIS/SLIP & INSURED SAMA PERSIS dgn
KBRU_cleaning_suspend.py v5 (lihat docstring file itu utk daftar lengkap
16 poin perubahan). Struktur file ini tidak berubah dari v4.

CATATAN VERSI (v4) -- bug report ke-3 (14/09/26, file:
KBRU_BUG_140926.md). Mesin POLIS/SLIP & INSURED SAMA PERSIS dgn
KBRU_cleaning_suspend.py v4 (lihat docstring file itu utk daftar
lengkap perubahan: normalisasi SINAR MAS->SINARMAS, hasil teks/nama
tidak lagi kena hapus-karakter penuh, rantai P<n> berpisah spasi/anotasi
ADD ikut dibuang, rantai dash>=5 dibiarkan apa adanya, fix target
substitusi pola F/S 2-segmen). File ini sendiri tidak ada perubahan
struktural lain (tetap tidak ada seleksi sumber, sama spt v3).

CATATAN VERSI (v3) -- rewrite setelah bug report ke-2 (10/09/26). Mesin
cleaning karakter POLIS/SLIP & INSURED SAMA PERSIS dengan
KBRU_cleaning_suspend.py v3 (lihat docstring file itu utk daftar
lengkap perubahan mesin, contoh, dan CATATAN YANG MASIH PERLU
DIKONFIRMASI). File ini TIDAK kena bug seleksi sumber CLSDT/FAC (itu
khusus Osbal, Data 1 tidak punya kolom CLSDT). Yang BEDA khusus di file
ini:
1. Data 1 tidak punya kolom CLSDT -- FAC_POLICY_NO/FAC_SLIP langsung
   dipakai & di-clean spt biasa (TIDAK ada seleksi sumber, TIDAK ada
   penukaran isi/posisi kolom).
2. Ada kolom BUSINESS_PARTNER (dari 2 kolom "COMP_NAME" -- cedant &
   broker), tidak berubah dari v1/v2.
3. CERTIFICATE_1 SELALU KOSONG & posisinya TEPAT setelah POLIS_CLEAN_1.
   TIDAK ADA kolom CN_1 terpisah -- notasi "CN : ..." dibuang total oleh
   clean_value().

CATATAN ARSITEKTUR (sama dgn KBRU_cleaning_suspend.py v3, lihat file itu
utk detail lengkap tiap rule POLIS/SLIP/INSURED):
- Placeholder murni BERDIRI SENDIRI -> ABAIKAN (teks ORI persis). TBA/
  P<n> yang BERSANDINGAN dgn konten lain -> dibuang, sisanya dipakai.
- Pola "NNNNNNNN[A-Z]NNNNNNNN-NN" -> ABAIKAN juga (dash dipertahankan).
- "Pengulangan" digit: substitusi jalan di bagian PALING BERSTRUKTUR di
  rantai '+' (posisi mana pun), HANYA kalau bagian itu cukup berstruktur
  (ada huruf/>=2 pemisah/digit panjang >=10) -- kalau tidak, seluruh
  rantai dibiarkan apa adanya.
- Notasi "CN : ..." dibuang total (bukan lagi anggota tambahan).
- "Hapus karakter" -- semua karakter non-alfanumerik dibuang dari hasil
  akhir (kecuali pola khusus di atas).
- INSURED: rule sama dgn Suspend, termasuk pola PELABUHAN INDONESIA/
  CLUSTER/TOWER/REGIONAL/DISTRICT, suffix UD/PD/PT/CV/gelar, delimiter
  QQQ/OO/*, parenthetical dgn daftar koma dipertahankan, dst.

=====================================================================
CATATAN YANG MASIH PERLU DIKONFIRMASI:
(sama seperti KBRU_cleaning_suspend.py v3)
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
OUTPUT_FILE = os.path.join("output", "Kbru_output_facul.xlsx")

CEDANT_COL   = "COMP_NAME"
CEDANT_VALUE = "PT. KBRU REINSURANCE BROKERS (EX. SIMAS)"

BROKER_COL  = "COMP_NAME.1"
POLIS_COL   = "FAC_POLICY_NO"
SLIP_COL    = "FAC_SLIP"
INSURED_COL = "FAC_INSURED"

MAX_SPLIT_COLS = 5


# ─────────────────────────────────────────────────────────────────────────────
# HELPER UMUM
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
        return [", ".join(items)]
    return items


# ─────────────────────────────────────────────────────────────────────────────
# MESIN POLIS/SLIP KBRU (identik dgn KBRU_cleaning_suspend.py v3)
# ─────────────────────────────────────────────────────────────────────────────

_JUNK_STANDALONE_WORDS = {"TBA", "VAR", "VARIOUS", "SELISIH", "END", "-", "NAN", "NONE", "NULL"}
_P_CODE_RE = re.compile(r"^P\d+$", re.IGNORECASE)

_MONTH_RE = (
    r"JANUARI|FEBRUARI|MARET|APRIL|MEI|JUNI|JULI|AGUSTUS|SEPTEMBER|OKTOBER|"
    r"NOVEMBER|DES[EA]?MBER|JANUARY|FEBRUARY|MARCH|MAY|JUNE|JULY|AUGUST|OCTOBER|DECEMBER"
)
_MONTH_YEAR_TAIL_RE = re.compile(
    r"\s*/?\s*(?:" + _MONTH_RE + r")\.?\s*(?:19|20)?\d{0,4}\s*$", re.IGNORECASE
)

_CN_MARKER_STRIP_RE_UNUSED = None
_NO_ANYWHERE_RE = re.compile(r"\bNO\.\s*", re.IGNORECASE)
_LEADING_ALPHA_PLUS_RE = re.compile(r"^[^\d/]{2,}?\+\s*")
_TRAILING_JUNK_WORD_RE = re.compile(r"[\s/]+(?:TBA|VAR|VARIOUS)\s*$", re.IGNORECASE)

# [BARU 15/09] deskripsi murni SUSPENSE/UTANG PIUTANG yang nyasar di
# kolom polis/slip -- dibiarkan apa adanya (spasi TIDAK dihapus), sama
# spt perlakuan di kolom INSURED.
_SUSPENSE_LIKE_RE = re.compile(r"SUSPEN|\(?H\)?\s*UTANG\s*PIUTANG", re.IGNORECASE)

# [BARU 15/09] teks narasi MURNI (BORDERO/SUMMARY SIMAS RE/nama bulan/
# tahun/pemisah saja, TANPA angka lain) -- dibiarkan apa adanya.
_NARRATIVE_ONLY_RE = re.compile(
    r"^(?:BORDERO|SUMMARY|BULAN|SIMAS\s*RE|" + _MONTH_RE + r"|(?:19|20)\d{2}|[-\s,])+$",
    re.IGNORECASE,
)

# [BARU] rantai P<n> (+ anotasi opsional "(ADD n)") yang digabung "+" --
# dibuang SELURUHNYA (termasuk "+"-nya), di mana pun posisinya (bukan
# cuma di akhir).
_STRAY_P_CHAIN_RE = re.compile(
    r"(?:\bP\d+\b\s*(?:\(ADD\s*\d+\))?\s*\+\s*)+", re.IGNORECASE
)
# [BARU] rantai P<n> yang dipisah SPASI (bukan "+") yang nempel di
# BELAKANG nilai -- dibuang.
_TRAILING_P_CHAIN_RE = re.compile(r"(?:\s+P\d+\b)+\s*$", re.IGNORECASE)

# [BARU] segmen pertama (sebelum "/" pertama) yang berisi RANTAI angka
# pendek digabung "-" (bukan "+") sebanyak >=5 kali -- dianggap
# 'pengulangan berlebih', DIBIARKAN APA ADANYA (raw), sama seperti
# aturan pengulangan >5 via "+".
def _has_excessive_dash_chain(value: str) -> bool:
    first_seg = value.split("/", 1)[0]
    return first_seg.count("-") >= 5


KEEP_AS_IS_PATTERN_RE = re.compile(r"^\d{7,10}[A-Z]{1,2}\d{6,9}-\d{2}$", re.IGNORECASE)


def _is_junk_token(tok: str) -> bool:
    t = tok.strip().upper()
    if not t:
        return True
    if t in _JUNK_STANDALONE_WORDS:
        return True
    if _P_CODE_RE.match(t):
        return True
    return False


def _value_has_real_digits(text: str) -> bool:
    return bool(re.search(r"\d{3,}", text))


def _has_junk_marker(text: str) -> bool:
    return bool(re.search(r"\bTBA\b|\bP\d+\b", text, re.IGNORECASE))


def _split_long_space_tail(val: str):
    parts = re.split(r"\s{2,}", val, maxsplit=1)
    if len(parts) == 1:
        return val, None
    return parts[0].strip(), parts[1].strip()


def _strip_leading_junk_wrapper(val: str):
    """Buang wrapper di depan via '/' SAMPAI ketemu nilai asli:
    - PASS 1: cari segmen PERTAMA yang punya digit -> itu awal nilai
      asli (buang semua sebelumnya, apa pun isinya -- nama cedant/broker
      mis. 'SINARMAS', placeholder TBA/P<n>, dst).
    - PASS 2 (kalau TIDAK ADA segmen berdigit sama sekali): hanya buang
      segmen placeholder BAKU (TBA/P<n>) di depan; segmen bukan-baku
      lain (mis. nama perusahaan tanpa digit) TETAP DIPERTAHANKAN kalau
      itu satu-satunya isi yang tersisa (mis. 'TBA/SINARMAS' -> 'SINARMAS').
    Kalau SEMUA segmen placeholder baku -> None (murni placeholder)."""
    if "/" not in val:
        return None if _is_junk_token(val) else val
    segs = val.split("/")
    for i, s in enumerate(segs):
        if re.search(r"\d", s):
            return "/".join(segs[i:]).strip()
    start_idx = None
    for i, s in enumerate(segs):
        if not _is_junk_token(s.strip()):
            start_idx = i
            break
    if start_idx is None:
        return None
    return "/".join(segs[start_idx:]).strip()


def _strip_trailing_junk_wrapper(val: str) -> str:
    """Buang segmen TANPA digit (nama cedant/broker/kode singkat, mis.
    'ACA', 'SINAR MAS') yang nempel di BELAKANG via '/', SAMPAI ketemu
    segmen yang punya digit (dianggap bagian nomor asli). BEDA dgn
    _strip_leading_junk_wrapper (berbasis _is_junk_token) -- di
    belakang, kriterianya 'ada digit atau tidak', krn tag broker/cedant
    yg nempel di ekor biasanya bukan kode baku (TBA/P<n>) tapi nama
    bebas apa saja (mis. 'ACA')."""
    if "/" not in val:
        return val
    segs = val.split("/")
    end_idx = None
    for i in range(len(segs) - 1, -1, -1):
        if re.search(r"\d", segs[i]):
            end_idx = i + 1
            break
    if end_idx is None:
        return val
    return "/".join(segs[:end_idx]).strip()


def _find_substitution_target(text: str):
    """Cari segmen digit yang jadi TARGET substitusi 'pengulangan'.
    Kasus khusus: pola 'ANGKA + 1-2 HURUF + ANGKA' PERSIS (mis.
    'S00000271' pada '110551216F00000473', tanpa karakter lain) ->
    target SEGMEN TERAKHIR (setelah huruf) -- meski segmen pertama
    (di depan huruf) sering lebih panjang, BUKAN itu yg dimaksud
    'pengulangan'. Selain pola ini -> segmen digit TERPANJANG; kalau
    ada BEBERAPA segmen dgn panjang SAMA (tie), pilih yg PERTAMA di
    antaranya yang BUKAN nol semua (mis. '0000-5747': '0000' dan '5747'
    sama panjang, tapi '0000' cuma nol -> target '5747')."""
    m = re.match(r"^(\d+)([A-Za-z]{1,2})(\d+)$", text)
    if m:
        return m.group(3), m.start(3), m.end(3)
    candidates = list(re.finditer(r"\d+", text))
    if not candidates:
        return None, None, None
    max_len = max(len(c.group(0)) for c in candidates)
    tied = [c for c in candidates if len(c.group(0)) == max_len]
    non_zero = [c for c in tied if set(c.group(0)) != {"0"}]
    best = non_zero[0] if non_zero else tied[0]
    return best.group(0), best.start(), best.end()


def _qualifies_for_substitution(main: str) -> bool:
    """Nilai utama cukup 'berstruktur' utk dijadikan target substitusi
    'pengulangan' kalau: ada huruf (pola FPAR-style), ATAU >=2 pemisah
    '.'/'/' (pola dot-style spt '12.100.0001.82288'), ATAU total digitnya
    panjang (>=10, pola digit polos panjang spt '0124012000618'), ATAU
    ada >=1 pemisah '.'/'/'/'-' DAN total digit >=7 (pola spt '0000-
    5747', beda dari angka pendek polos TANPA pemisah spt '3522' yang
    TETAP tidak memenuhi)."""
    if re.search(r"[A-Za-z]", main):
        return True
    sep_count = len(re.findall(r"[./]", main))
    if sep_count >= 2:
        return True
    digit_len = len(re.sub(r"\D", "", main))
    if digit_len >= 10:
        return True
    if len(re.findall(r"[.\-/]", main)) >= 1 and digit_len >= 7:
        return True
    return False


def _apply_segment_substitution(primary: str, extra_tokens: list) -> list:
    seg, start, end = _find_substitution_target(primary)
    if seg is None:
        return [primary]
    prefix, suffix = primary[:start], primary[end:]
    out = [primary]
    for tok in extra_tokens:
        tok = tok.strip()
        if not tok.isdigit():
            continue
        if len(tok) >= len(seg):
            new_seg = tok
        else:
            new_seg = seg[: -len(tok)] + tok
        out.append(prefix + new_seg + suffix)
    return _dedup_preserve_order(out)


# [BARU 15/09] dash NYASAR di dalam pola F/S 2-segmen (mis.
# '11055129F00000-275' -- seharusnya '11055129F00000275') -- dash
# dianggap noise/typo, dilebur.
_FS_STRAY_DASH_RE = re.compile(r"^(\d+[A-Za-z]{1,2}\d+)-(\d+)$")


def _merge_fs_stray_dash(token: str) -> str:
    m = _FS_STRAY_DASH_RE.match(token)
    if m:
        return m.group(1) + m.group(2)
    return token


_CN_TRAILING_RE = re.compile(r"\s*/?\s*CN\s*:\s*.*$", re.IGNORECASE)
_CN_LEADING_WITH_SLASH_RE = re.compile(r"^\s*CN\s*:\s*.*?\s+/\s+", re.IGNORECASE)
_CN_LABEL_ONLY_RE = re.compile(r"^\s*CN\s*:\s*", re.IGNORECASE)
_CN_STARTS_RE = re.compile(r"^CN\s*:", re.IGNORECASE)


def _strip_cn_suffix(val: str) -> str:
    """Notasi 'CN : <nomor>':
    - Muncul di TENGAH/BELAKANG ('main / CN : cn') -> dibuang total
      bersama nilainya, sisakan main di depan.
    - Muncul di AWAL DIIKUTI nilai lain setelah '/' ('CN : cn / main')
      -> label+nilai CN dibuang, sisakan main setelah '/'.
    - HANYA 'CN : cn' saja (seluruh value, tanpa '/' sama sekali) ->
      cuma label 'CN :'-nya yg dibuang, nilai CN dipertahankan sbg
      hasil utama (BUKAN dibuang semua)."""
    stripped = val.strip()
    if not _CN_STARTS_RE.match(stripped):
        return _CN_TRAILING_RE.sub("", val).strip(" /-")
    m = _CN_LEADING_WITH_SLASH_RE.match(val)
    if m:
        return val[m.end():].strip(" /-")
    return _CN_LABEL_ONLY_RE.sub("", val).strip(" /-")


_EMBEDDED_SUFFIX_RE = re.compile(r"^(\d+)([./].*[A-Za-z].*)$")


def _split_embedded_suffix(token: str):
    """Deteksi token yg formatnya 'ANGKA' + 'SUFFIX BERSTRUKTUR' nempel
    jadi satu (mis. '23022/FIAR/001.0000344-SRB/06/2022' -> ('23022',
    '/FIAR/001.0000344-SRB/06/2022')). Dipakai utk kasus 'pengulangan'
    di mana suffix format cuma nempel di SALAH SATU token dlm rantai
    '+' (biasanya yg terakhir), bukan di token pertama."""
    m = _EMBEDDED_SUFFIX_RE.match(token)
    if m:
        return m.group(1), m.group(2)
    return token, None


def _strip_all_characters(value: str) -> str:
    return re.sub(r"[^A-Za-z0-9]", "", value)


_SINAR_MAS_RE = re.compile(r"\bSINAR\s+MAS\b", re.IGNORECASE)
_TEXT_PIECE_PTCV_RE = re.compile(r",?\s*\b(?:P\s*T|CV)\b\.?\s*$", re.IGNORECASE)


def _clean_text_piece(value: str) -> str:
    """Utk hasil yang ternyata berupa TEKS/NAMA (tanpa digit sama
    sekali, muncul lewat mekanisme buang TBA/P<n> yang bersandingan) --
    JANGAN di-'hapus karakter' penuh (spasi antar kata dipertahankan,
    beda dari nilai numerik biasa), cuma normalisasi spasi + standardisasi
    'SINAR MAS' -> 'SINARMAS' (konsisten dgn kolom INSURED) + buang
    suffix "PT"/"CV" di akhir (mis. "ASURANSI FPG INSURANCE, PT" ->
    "ASURANSI FPG INSURANCE")."""
    value = _SINAR_MAS_RE.sub("SINARMAS", value)
    value = _TEXT_PIECE_PTCV_RE.sub("", value)
    return _normalize_spaces(value).upper()


def _simplify_noletter_extra(e: str) -> str:
    """[BARU 15/09] Token tambahan (di rantai '+') yang TIDAK punya
    huruf tapi masih ada '/' atau '.' nempel (mis. '04820/0001189' dari
    '...+04820/0001189.') -- ambil cuma potongan PERTAMA sbg kandidat
    substitusi digit, sisanya (mis. sertifikat/nomor duplikat) dibuang
    (KBRU tidak punya kolom sertifikat, jadi info ini tidak tertampung)."""
    if re.search(r"[A-Za-z]", e):
        return e
    return re.split(r"[./]", e)[0]


def clean_value(val) -> list:
    if pd.isna(val):
        return []
    raw = str(val).strip()
    if not raw:
        return []

    if raw.upper() in _JUNK_STANDALONE_WORDS:
        return [raw]
    # [BARU 15/09] "TBA/" (placeholder + garis miring doang, tanpa isi
    # lain) -> garis miringnya dibuang, cukup "TBA".
    raw_no_slash = raw.strip("/ ")
    if raw_no_slash.upper() in _JUNK_STANDALONE_WORDS:
        return [raw_no_slash]
    if KEEP_AS_IS_PATTERN_RE.match(raw):
        return [raw]
    if re.fullmatch(r"\d+&\d+", raw):
        return [raw]
    if _has_excessive_dash_chain(raw):
        # segmen depan berisi >=5 angka digabung '-' -- pengulangan
        # berlebih, dibiarkan apa adanya (sama spt aturan >5 via '+')
        return [raw]
    if _SUSPENSE_LIKE_RE.search(raw) or _NARRATIVE_ONLY_RE.match(raw.strip()):
        # [BARU 15/09] deskripsi SUSPENSE/UTANG PIUTANG, atau narasi
        # murni (BORDERO/bulan/tahun/SUMMARY SIMAS RE) yg nyasar di
        # kolom polis/slip -- dibiarkan apa adanya, spasi TIDAK dihapus
        return [raw]

    if not _has_junk_marker(raw) and not _value_has_real_digits(raw):
        # teks tanpa placeholder (TBA/P<n>) & tanpa digit sama sekali
        # (mis. nama perusahaan nyasar, "ASURANSI SINAR MAS") -> tetap
        # "abaikan" (bukan diproses spt polis/slip), TAPI tetap lewat
        # normalisasi teks (spasi + SINAR MAS->SINARMAS) spy konsisten
        # dgn kasus yg lewat jalur TBA-adjacent.
        return [_clean_text_piece(raw)]

    val2 = _NO_ANYWHERE_RE.sub("", raw)
    val2 = _strip_cn_suffix(val2)
    val2 = _LEADING_ALPHA_PLUS_RE.sub("", val2)
    val2 = _STRAY_P_CHAIN_RE.sub("", val2)
    val2 = _TRAILING_P_CHAIN_RE.sub("", val2)

    pieces = []
    stripped = _strip_leading_junk_wrapper(val2)
    if stripped is not None:
        val2 = stripped
        val2 = _TRAILING_P_CHAIN_RE.sub("", val2)
        primary, _tail = _split_long_space_tail(val2)
        primary2 = _MONTH_YEAR_TAIL_RE.sub("", primary).strip(" -/,")
        primary = primary2 if primary2 else primary
        primary = _TRAILING_JUNK_WORD_RE.sub("", primary).strip()

        raw_parts = [p.strip() for p in re.split(r"\s*\+\s*", primary) if p.strip()]
        raw_parts = [_merge_fs_stray_dash(p) for p in raw_parts]
        # [FIX 15/09] "_strip_trailing_junk_wrapper" cuma dipakai kalau
        # MEMANG ada rantai '+' (>1 bagian) -- kalau cuma 1 bagian
        # (whole value, mis. hasil gabungan lewat "No." di tengah),
        # JANGAN dipangkas, supaya ekor sah spt "/AB"/"/BB" (kode akhir
        # 2 huruf tanpa digit) tidak ikut kebuang.
        apply_trailing_strip = len(raw_parts) > 1
        parts = []
        for p in raw_parts:
            if _is_junk_token(p):
                continue
            p2 = _strip_trailing_junk_wrapper(p) if apply_trailing_strip else p
            p2 = p2.strip()
            if p2 and not _is_junk_token(p2):
                parts.append(p2)

        if parts:
            if len(parts) > 1:
                # kalau ADA campuran (sebagian bagian punya digit,
                # sebagian tidak sama sekali, mis. 'UMUM MEGA' vs
                # '2010208012400001') -> bagian tanpa digit dibuang
                # (nama cedant/broker yg nyempil, bukan bagian nilai).
                digit_parts = [p for p in parts if re.search(r"\d", p)]
                if digit_parts:
                    parts = digit_parts

            # [CEK 1] suffix berstruktur nempel di SALAH SATU token
            # (biasanya token TERAKHIR, mis. '23018+23019+23023+23022/
            # FIAR/001.0000344-SRB/06/2022') -- kalau ketemu & SISANYA
            # semua digit polos, terapkan suffix yg sama ke semua nomor.
            suffix_found = None
            base_numbers = []
            embed_ok = True
            for p in parts:
                num_part, suf = _split_embedded_suffix(p)
                if suf:
                    if suffix_found is None:
                        suffix_found = suf
                    else:
                        embed_ok = False
                if not num_part.isdigit():
                    embed_ok = False
                base_numbers.append(num_part)

            if embed_ok and suffix_found and len(base_numbers) > 1:
                ref = max(base_numbers, key=len)
                combined = []
                for n in base_numbers:
                    if len(n) >= len(ref):
                        combined.append(n + suffix_found)
                    else:
                        combined.append(ref[: -len(n)] + n + suffix_found)
                pieces = _dedup_preserve_order(combined)
            else:
                # [CEK 2] main = bagian yg PALING berstruktur (ada digit
                # DAN (huruf ATAU >=2 pemisah)), di mana pun posisinya
                # dalam rantai '+' -- sisanya jadi kandidat token
                # substitusi/nilai lain.
                structured_idx = None
                for i, p in enumerate(parts):
                    if re.search(r"\d", p) and (
                        re.search(r"[A-Za-z]", p) or len(re.findall(r"[./]", p)) >= 2
                    ):
                        structured_idx = i
                        break
                if structured_idx is not None:
                    main = parts[structured_idx]
                    extras_raw = parts[:structured_idx] + parts[structured_idx + 1:]
                else:
                    main = parts[0]
                    extras_raw = parts[1:]

                extras_clean = [_simplify_noletter_extra(e.strip().rstrip(".").strip()) for e in extras_raw]
                whole_extras = [e for e in extras_clean if "/" in e or re.search(r"\d\.\d", e)]
                digit_extras = [e for e in extras_clean if e.isdigit()]

                # [BARU 15/09] token substitusi yg SANGAT pendek (<=2
                # digit, mis. "+97+07+43") -- pola pengulangannya tidak
                # jelas (berapa digit ekor yg diganti jadi ambigu) ->
                # seluruh nilai dibiarkan apa adanya.
                if digit_extras and any(len(e) < 3 for e in digit_extras):
                    pieces = []
                elif digit_extras and not _qualifies_for_substitution(main):
                    pieces = []  # ambigu -> jatuh ke abaikan (raw) di bawah
                elif digit_extras:
                    pieces = _apply_segment_substitution(main, digit_extras)
                    pieces += whole_extras
                else:
                    pieces = [main] + whole_extras

    pieces = _dedup_preserve_order([_normalize_spaces(p) for p in pieces if p.strip()])

    if not pieces or len(pieces) > MAX_SPLIT_COLS:
        return [raw]

    final = []
    for p in pieces:
        if KEEP_AS_IS_PATTERN_RE.match(p):
            final.append(p)
        elif not re.search(r"\d", p):
            final.append(_clean_text_piece(p))
        else:
            final.append(_strip_all_characters(p))
    return final


def clean_polis(val):
    return clean_value(val)


def clean_slip(val):
    return clean_value(val)


# ─────────────────────────────────────────────────────────────────────────────
# CLEAN INSURED (identik dgn KBRU_cleaning_suspend.py v3)
# ─────────────────────────────────────────────────────────────────────────────

_GELAR_PREFIX_RE = re.compile(
    r"^\s*(?:BAPAK|IBU|SDRI?|NYONYA|NY|TUAN|TN|MR|MRS|MS|DR|IR|PROF|HJ|H)\b\.?\s*",
    re.IGNORECASE,
)
_GELAR_SUFFIX_RE = re.compile(
    r"\s*,?\s*\b(?:S\.?E|S\.?H|S\.?SI|S\.?T|S\.?SOS|S\.?KOM|S\.?PD|"
    r"M\.?SI|M\.?M|M\.?H|M\.?KOM|SH|SE|BP|MRS|MR|NY|TN|DRS)\b\.?\s*$",
    re.IGNORECASE,
)
_SUSPENSE_KEEP_RE = re.compile(r"SUSPEN|\(?H\)?\s*UTANG\s*PIUTANG", re.IGNORECASE)
_TREATY_KEEP_RE = re.compile(
    r"\bTREATY\b|\bSTRUCTURE\b|EXCESS\s+OF\s+LOSS|WHOLE\s+ACCOUNT", re.IGNORECASE
)


def _is_keep_as_is_insured(val: str) -> bool:
    if _SUSPENSE_KEEP_RE.search(val):
        return True
    if _TREATY_KEEP_RE.search(val):
        return True
    return False


def _collapse_duplicate_words(text: str) -> str:
    text = re.sub(r"\(\s*\(", "(", text)
    text = re.sub(r"\)\s*\)", ")", text)
    text = re.sub(r"\b(\w{3,})\1\b", r"\1", text, flags=re.IGNORECASE)
    text = re.sub(r"\b(\w+)(?:\s+\1\b)+", r"\1", text, flags=re.IGNORECASE)
    if text.count("(") > text.count(")"):
        text = text + ")"
    return text


_KEEP_SUFFIX_COMMA_RE = re.compile(r",\s*(UD|PD)\b", re.IGNORECASE)
_REMOVE_COMMA_PTCV_RE = re.compile(r",\s*(?:P\s*T|CV)\b\.?", re.IGNORECASE)

_SPLIT_RE = re.compile(
    r"\bQ{2,3}\b|\bOO\b|\bAND\s*/\s*OR\.?\b|/{1,2}|,|–|¿|\*|&\s*(?=(?:PT|CV)\b)",
    re.IGNORECASE,
)

INSURED_JUNK_WORDS = frozenset({
    "PT", "CV", "TBK", "PERSERO", "LTD", "INC", "LLC",
    "AND", "OR", "THE", "OF", "AS", "VARIOUS", "VAR", "QQ", "TBA",
    "ASSOCIATED",
})

_MONTH_STANDALONE_RE = re.compile(
    r"^(?:" + _MONTH_RE + r")\.?\s*(?:19|20)?\d{2,4}$", re.IGNORECASE
)

_GELAR_PREFIX_RE = re.compile(
    r"^\s*(?:BAPAK|IBU|SDRI?|NYONYA|NY|TUAN|TN|MR|MRS|MS|DR|IR|PROF|HJ|H)\b\.?\s*",
    re.IGNORECASE,
)
# [BARU] tambah BP/MR/MRS/NY/TN/DRS sbg gelar SUFFIX (nempel via koma di
# BELAKANG nama, mis. "SUTJITA, BP" / "RUDY AKILI, MR") -- dibuang.
_GELAR_SUFFIX_RE = re.compile(
    r"\s*,?\s*\b(?:S\.?E|S\.?H|S\.?SI|S\.?T|S\.?SOS|S\.?KOM|S\.?PD|"
    r"M\.?SI|M\.?M|M\.?H|M\.?KOM|SH|SE|BP|MRS|MR|NY|TN|DRS)\b\.?\s*$",
    re.IGNORECASE,
)
_SUSPENSE_KEEP_RE = re.compile(r"SUSPEN|\(?H\)?\s*UTANG\s*PIUTANG", re.IGNORECASE)
_TREATY_KEEP_RE = re.compile(
    r"\bTREATY\b|\bSTRUCTURE\b|EXCESS\s+OF\s+LOSS|WHOLE\s+ACCOUNT", re.IGNORECASE
)


def _is_keep_as_is_insured(val: str) -> bool:
    if _SUSPENSE_KEEP_RE.search(val):
        return True
    if _TREATY_KEEP_RE.search(val):
        return True
    if _MALFORMED_QQ_RE.search(val):
        return True
    return False


# [BARU] normalisasi data rusak: kata terduplikasi persis berturut-turut
# (mis. "WASKITA WASKITA KARYA KARYA" -> "WASKITA KARYA"), kata gandeng
# tanpa spasi ("PERSEROPERSERO" -> "PERSERO"), kurung ganda/tak seimbang.
def _collapse_duplicate_words(text: str) -> str:
    text = re.sub(r"\(\s*\(", "(", text)
    text = re.sub(r"\)\s*\)", ")", text)
    text = re.sub(r"\b(\w{3,})\1\b", r"\1", text, flags=re.IGNORECASE)
    text = re.sub(r"\b(\w+)(?:\s+\1\b)+", r"\1", text, flags=re.IGNORECASE)
    if text.count("(") > text.count(")"):
        text = text + ")"
    return text


# [BARU] ", PT" / ", CV" (di mana pun, bukan cuma di akhir) DIBUANG
# TOTAL bersama komanya -- supaya tidak memicu split palsu saat ada isi
# lain menyertai PT/CV di segmen yang sama (mis. "SMAILING TOUR, PT &
# TRAVEL SERVICE, PT" -> "SMAILING TOUR & TRAVEL SERVICE", bukan
# terpecah jadi "SMAILING TOUR" + "& TRAVEL SERVICE").
_REMOVE_COMMA_PTCV_RE = re.compile(r",\s*(?:P\s*T|CV)\b\.?", re.IGNORECASE)

# ", UD" / ", PD" (jenis badan usaha) TIDAK dianggap pemisah entitas
# baru -- koma dibuang, kode dipertahankan menyatu dgn nama sebelumnya.
_KEEP_SUFFIX_COMMA_RE = re.compile(r",\s*(UD|PD)\b", re.IGNORECASE)

# pemisah: QQ/QQQ, OO (typo QQ), AND/OR, "/", "//", ",", en-dash, "¿",
# "*", dan "&" HANYA kalau diikuti PT/CV (JV baru, mis. "... & PT OKI
# PULP..."). "&" polos (mis. "PULP & PAPER") TETAP bukan pemisah. "+"
# TIDAK dipakai sbg pemisah (mis. "SK+SUMATERA" -> 1 entitas "SK
# SUMATERA", bukan 2) -- beda dgn mesin polis/slip. " - " (spasi-dash-
# spasi) juga TIDAK dipakai (bikin false split di nama spt "PULAU -
# PULAU MEDIA").
_SPLIT_RE = re.compile(
    r"\bQ{2,3}\b|\bOO\b|\bAND\s*/\s*OR\.?\b|/{1,2}|,|–|¿|\*|&\s*(?=(?:PT|CV)\b)",
    re.IGNORECASE,
)

INSURED_JUNK_WORDS = frozenset({
    "PT", "CV", "TBK", "PERSERO", "LTD", "INC", "LLC",
    "AND", "OR", "THE", "OF", "AS", "VARIOUS", "VAR", "QQ", "TBA",
    "ASSOCIATED",
})

_OWNER_OF_RE = re.compile(
    r"\(\s*AS\s+(?:THE\s+)?(?:OWNER|PRINCIPAL)\s+OF\s+([^()]+)\)", re.IGNORECASE
)

_ROLE_KEYWORDS_RE = re.compile(
    r"^(?:PERSERO|PRINCIPAL|OWNER|INSURED|CONTRACTOR)$", re.IGNORECASE
)

# [DIPERLONGGAR] toleransi "P T" (spasi nyempil di antara P & T, data
# rusak) selain "PT" biasa.
_PT_CV_PREFIX_RE = re.compile(r"^\s*(?:P\s*T|CV)(?:\.|(?=\s|$))\.?\s*", re.IGNORECASE)
_PT_CV_SUFFIX_RE = re.compile(r",?\s*(?:P\s*T|CV)\.?\s*$", re.IGNORECASE)

_MASK_COMMA = "\x01"
_MASK_SLASH = "\x02"


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
    words = [w for w in re.split(r"[\s\-]+", long_name.upper()) if w]
    if any(w.startswith(short_clean) and w != short_clean for w in words):
        return True
    return False


def _extract_owner_of_entities(text: str):
    extra = []

    def repl(m):
        extra.append(m.group(1).strip())
        return " "

    new_text = _OWNER_OF_RE.sub(repl, text)
    return new_text, extra


def _strip_parenthetical(text: str) -> str:
    """Isi dalam kurung yang singkatan dari teks sebelumnya, ATAU kata
    boilerplate peran (PERSERO/PRINCIPAL/OWNER/INSURED/CONTRACTOR), ATAU
    mengandung PT/CV -> kurung + isi DIBUANG total. Isi kurung yang
    mengandung KOMA (daftar deskriptif, mis. "(REFINERY, PLANTATION,
    KEMITRAAN)") -> kurung DIPERTAHANKAN UTUH (literal, koma di-mask
    dulu spy tdk ikut split). Isi kurung lain (identitas cabang/pihak
    lain, mis. "(KSO)") -> kurung dibuang, ISI dipertahankan."""

    def repl(m):
        inner = m.group(1).strip()
        if re.search(r"\bPT\b|\bCV\b", inner, re.IGNORECASE):
            return " "
        if _ROLE_KEYWORDS_RE.match(inner):
            return " "
        inner_letters = re.sub(r"[^A-Za-z]", "", inner)
        before = text[: m.start()].strip()
        if before and inner_letters and _is_abbreviation_of(inner_letters, before):
            return " "
        if "," in inner:
            masked = inner.replace(",", _MASK_COMMA).replace("/", _MASK_SLASH)
            return f" ({masked}) "
        masked = inner.replace(",", _MASK_COMMA).replace("/", _MASK_SLASH)
        return f" {masked} "

    return re.sub(r"\(([^()]*)\)", repl, text)


# ─────────────────────────────────────────────────────────────────────────────
# [BARU] EKSPANSI POLA "KEYWORD + DAFTAR" (FACTORY/CLUSTER/TOWER/
# REGIONAL/DISTRICT) & KHUSUS PELABUHAN INDONESIA/PELINDO
# ─────────────────────────────────────────────────────────────────────────────

_ROMAN_MAP = {"I": 1, "II": 2, "III": 3, "IV": 4, "V": 5, "VI": 6, "VII": 7, "VIII": 8, "IX": 9, "X": 10}
_ROMAN_ARABIC_ALT = r"I|II|III|IV|V|VI|VII|VIII|IX|X|\d+"

_PELINDO_RE = re.compile(
    r"^(PELABUHAN\s+INDONESIA)\s+((?:" + _ROMAN_ARABIC_ALT + r")"
    r"(?:\s*(?:,|&|DAN)\s*(?:" + _ROMAN_ARABIC_ALT + r"))+)\s*$",
    re.IGNORECASE,
)

# [BARU 15/09] "PELINDO" adalah SINONIM "PELABUHAN INDONESIA" -- bisa
# disebut BERULANG-ULANG dgn pemisah apa saja (biasanya QQ), mis.
# "PELABUHAN INDONESIA I QQ PELINDO I QQ PELINDO II QQ PELINDO III QQ
# PELINDO IV" -> semua kemunculan "PELABUHAN INDONESIA <n>"/"PELINDO <n>"
# dikumpulkan, dinormalisasi ke "PELABUHAN INDONESIA <n>" (arab), dedup.
_PELINDO_MENTION_RE = re.compile(
    r"\b(?:PELABUHAN\s+INDONESIA|PELINDO)\s+(" + _ROMAN_ARABIC_ALT + r")\b",
    re.IGNORECASE,
)


def _expand_pelindo(val: str):
    m = _PELINDO_RE.match(val.strip())
    if not m:
        return None
    prefix, nums_part = m.groups()
    tokens = [t.strip() for t in re.split(r"\s*(?:,|&|DAN)\s*", nums_part, flags=re.IGNORECASE) if t.strip()]
    out = []
    for t in tokens:
        num = _ROMAN_MAP.get(t.upper(), t)
        out.append(f"{prefix.upper()} {num}")
    return out


def _expand_pelindo_mentions(val: str):
    matches = _PELINDO_MENTION_RE.findall(val)
    if len(matches) < 2:
        return None
    seen = []
    for t in matches:
        num = _ROMAN_MAP.get(t.upper(), t)
        entry = f"PELABUHAN INDONESIA {num}"
        if entry not in seen:
            seen.append(entry)
    return seen if len(seen) >= 2 else None


_KEYWORD_MULTI_RE = re.compile(
    r"^([^/]*?)[/\s]*\b(FACTORY|CLUSTER|TOWER|REGIONAL|DISTRICT|UNIT|MALL)\s+([^/]+)$",
    re.IGNORECASE,
)


def _expand_keyword_multi(val: str):
    """'PREFIX [/] KEYWORD tok1 tok2 ...' atau 'PREFIX [/] KEYWORD tok1,
    tok2 & tok3' -> daftar 'PREFIX KEYWORD tok'. Butuh >=2 token supaya
    tidak salah nembak segmen tunggal (yg ditangani _merge_keyword_single
    di tempat lain)."""
    m = _KEYWORD_MULTI_RE.match(val.strip())
    if not m:
        return None
    prefix, keyword, tail = m.groups()
    tokens = [t for t in re.split(r"\s*(?:,|&|DAN)\s*|\s+", tail.strip(), flags=re.IGNORECASE) if t.strip()]
    tokens = [t.strip() for t in tokens if re.fullmatch(r"[A-Za-z0-9]{1,4}", t.strip())]
    if len(tokens) < 2:
        return None
    prefix = prefix.strip().rstrip(",").strip()
    prefix = re.sub(r",?\s*(?:P\s*T|CV)\.?\s*$", "", prefix, flags=re.IGNORECASE).strip()
    return [f"{prefix} {keyword.upper()} {t.upper()}".strip() for t in tokens]


_KEYWORD_SINGLE_RE = re.compile(
    r"^(?:FACTORY|CLUSTER|TOWER|REGIONAL|DISTRICT|UNIT|MALL)\s+[A-Za-z0-9]{1,4}$", re.IGNORECASE
)

# [BARU 15/09] "QQ" yang NYASAR nempel langsung ke huruf lain tanpa
# spasi/pemisah (mis. "QQBALI", "INDAHQQ") -- tanda data rusak/typo yg
# terlalu ambigu utk diproses -> seluruh value dibiarkan apa adanya.
# Rantai Q murni ("QQQ") TIDAK dihitung rusak -- itu cuma varian QQ
# dgn Q berlebih, sudah ditangani _SPLIT_RE (\bQ{2,3}\b).
_MALFORMED_QQ_RE = re.compile(r"[A-PR-Za-pr-z]QQ|QQ[A-PR-Za-pr-z]")


def _clean_insured_name(name: str) -> str:
    name = _normalize_spaces(name)
    name = _GELAR_PREFIX_RE.sub("", name)
    for _ in range(3):
        stripped = _GELAR_SUFFIX_RE.sub("", name).strip()
        if stripped == name:
            break
        name = stripped
    name = re.sub(r",?\s*\(\s*PERSERO\s*\)\s*", " ", name, flags=re.IGNORECASE)
    name = re.sub(r"\bTBK\b\s*,\s*P\s*T\.?\s*", " ", name, flags=re.IGNORECASE)
    name = re.sub(r",?\s*\bTBK\b\s*", " ", name, flags=re.IGNORECASE)
    name = re.sub(r"\bLTD\b\.?\s*$", "", name, flags=re.IGNORECASE)
    name = re.sub(r"\bPTE\b\.?\s*", "", name, flags=re.IGNORECASE)
    name = _PT_CV_SUFFIX_RE.sub("", name)
    name = _PT_CV_PREFIX_RE.sub("", name)
    # [FIX 15/09] pola inisial huruf-tunggal-bertitik berturut-turut
    # (mis. "S.A.R.I") -> titik dibuang TANPA jadi spasi ("SARI"),
    # BUKAN "S A R I". Titik lain (bukan pola ini, bukan antar digit)
    # baru diganti spasi spt biasa.
    name = re.sub(r"\b(?:[A-Za-z]\.){2,}[A-Za-z]?\b", lambda mo: mo.group(0).replace(".", ""), name)
    name = re.sub(r"(?<!\d)\.(?!\d)", " ", name)
    name = re.sub(r"\s*\+\s*", " ", name)
    name = _SINAR_MAS_RE.sub("SINARMAS", name)
    name = name.replace(_MASK_COMMA, ",").replace(_MASK_SLASH, "/")
    return _normalize_spaces(name).strip().upper()


def _try_expansions(segment: str):
    """Coba semua mekanisme ekspansi (pelindo list, pelindo scattered
    mentions, keyword+daftar) pada SATU string. None kalau tidak ada yg
    cocok."""
    r = _expand_pelindo(segment)
    if r is not None:
        return r
    r = _expand_pelindo_mentions(segment)
    if r is not None:
        return r
    r = _expand_keyword_multi(segment)
    if r is not None:
        return r
    return None


def clean_insured(val) -> list:
    if pd.isna(val):
        return []
    val = str(val).strip()
    if not val:
        return []

    val = _collapse_duplicate_words(val)

    if _is_keep_as_is_insured(val):
        return [_normalize_spaces(val).upper()]

    val, owner_extra = _extract_owner_of_entities(val)
    val = _strip_parenthetical(val)
    val = _KEEP_SUFFIX_COMMA_RE.sub(r" \1", val)
    val = _REMOVE_COMMA_PTCV_RE.sub("", val)

    # coba ekspansi pada SELURUH value dulu (kasus umum: satu pola
    # keyword/pelindo mencakup semua, mis. "BIO FARMA/CLUSTER B C D",
    # "PELABUHAN INDONESIA I QQ PELINDO II ...").
    whole = _try_expansions(val)
    if whole is not None:
        return [_normalize_spaces(x).upper() for x in whole]

    parts = [p.strip() for p in _SPLIT_RE.split(val) if p.strip()]
    cleaned = []
    for p in parts:
        # [FIX 15/09] kalau SATU bagian hasil split (bukan whole value)
        # sendiri cocok pola keyword+daftar (mis. "GALAXY MALL I & II"
        # setelah "SINAR GALAXY /" ikut terpisah lebih dulu) -> expand
        # jadi beberapa entitas, JANGAN diproses sbg 1 entitas biasa.
        r = _expand_keyword_multi(p)
        if r is not None and len(r) >= 2:
            cleaned.extend(_normalize_spaces(x).upper() for x in r)
            continue
        p_clean = _clean_insured_name(p)
        if not p_clean:
            continue
        if _MONTH_STANDALONE_RE.match(p_clean):
            continue
        if _KEYWORD_SINGLE_RE.match(p_clean) and cleaned:
            cleaned[-1] = _normalize_spaces(cleaned[-1] + " " + p_clean)
            continue
        if len(p_clean) < 2 or p_clean in INSURED_JUNK_WORDS:
            continue
        cleaned.append(p_clean)

    for extra in owner_extra:
        extra_clean = _clean_insured_name(extra)
        if extra_clean and extra_clean not in cleaned:
            cleaned.append(extra_clean)

    if not cleaned:
        fallback = _clean_insured_name(val)
        return [fallback] if fallback else []

    return _cap_or_join(cleaned)


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


def _insert_polis_cert_columns(df, all_polis_lists, max_cols):
    added = []
    actual_max_cols = max((len(lst) for lst in all_polis_lists if isinstance(lst, list)), default=1)
    actual_max_cols = min(actual_max_cols, max_cols)
    for i in range(1, actual_max_cols + 1):
        polis_col = f"POLIS_CLEAN_{i}"
        df[polis_col] = [lst[i - 1] if isinstance(lst, list) and i - 1 < len(lst) else None for lst in all_polis_lists]
        added.append(polis_col)
        if i == 1:
            df["CERTIFICATE_1"] = None
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

    all_polis, all_slip, all_insured = [], [], []
    max_polis = max_slip = max_ins = 1

    for _, row in df.iterrows():
        p_clean = clean_polis(row.get(POLIS_COL, ""))
        s_clean = clean_slip(row.get(SLIP_COL, ""))
        c_ins = clean_insured(row.get("FAC_INSURED_ORI", ""))

        max_polis = max(max_polis, len(p_clean))
        max_slip = max(max_slip, len(s_clean))
        max_ins = max(max_ins, len(c_ins))

        all_polis.append(p_clean)
        all_slip.append(s_clean)
        all_insured.append(c_ins)

    print("[4/5] Menyusun kolom output ...")

    new_columns = []
    for col in df.columns:
        new_columns.append(col)
        if col == POLIS_COL:
            new_columns += _insert_polis_cert_columns(df, all_polis, max_polis)
        elif col == SLIP_COL:
            new_columns += _insert_clean_columns(df, all_slip, f"{SLIP_COL}_CLEAN", max_slip)
        elif col == "FAC_INSURED_ORI":
            new_columns += _insert_clean_columns(df, all_insured, "FAC_INSURED", max_ins)

    df = df[new_columns]

    # FORMAT OUTPUT FACUL: susunan dan nama kolom mengikuti file tabel facul.xlsx.
    # Ini hanya pemetaan kolom output; nilai hasil cleaning tidak dihitung ulang.
    df = df.rename(columns={'FAC_INSURED_ORI': 'FAC_INSURED', 'COMP_NAME.1': 'COMP_NAME2', 'BUSINESS_PARTNER': 'BUSSINESS_PATNER'})
    df = df.rename(columns={
        **{f"POLIS_CLEAN_{i}": f"POLICY_CLEAN_{i}" for i in range(1, 6)},
        **{f"CERTIFICATE_{i}": f"SERTIF_CLEAN_{i}" for i in range(1, 6)},
        **{f"FAC_SLIP_CLEAN_{i}": f"SLIP_CLEAN_{i}" for i in range(1, 6)},
        **{f"FAC_INSURED_{i}": f"INSURED_CLEAN_{i}" for i in range(1, 6)},
    })
    # Kolom dari template yang tidak ada di sumber dibuat kosong.
    # Susunan dan ejaan header diambil persis dari dua template yang diberikan.
    output_columns = ['FAC_CODE', 'FAC_CEDANT', 'COMP_NAME', 'FAC_BROKER', 'COMP_NAME2', 'BUSSINESS_PATNER', 'FAC_INSURED', 'INSURED_CLEAN_1', 'INSURED_CLEAN_2', 'INSURED_CLEAN_3', 'INSURED_CLEAN_4', 'INSURED_CLEAN_5', 'FAC_POLICY_NO', 'POLICY_CLEAN_1', 'SERTIF_CLEAN_1', 'POLICY_CLEAN_2', 'SERTIF_CLEAN_2', 'POLICY_CLEAN_3', 'SERTIF_CLEAN_3', 'POLICY_CLEAN_4', 'POLICY_CLEAN_5', 'FAC_SLIP', 'SLIP_CLEAN_1', 'SLIP_CLEAN_2', 'SLIP_CLEAN_3', 'SLIP_CLEAN_4', 'SLIP_CLEAN_5', 'FAC_CURRENCY', 'FAC_INP_DATE', 'FAC_COM_DATE', 'FAC_EXP_DATE', 'FAC_SUB_CLASS', 'FAC_RISK', 'FAC_DESC', 'FAC_ACC_STS', 'FAC_STS_SLIP']
    df = df.reindex(columns=output_columns)

    os.makedirs(os.path.dirname(output_file), exist_ok=True)
    print(f"[5/5] Menyimpan hasil ke: {output_file} ...")
    df.to_excel(output_file, index=False)
    _force_text_format(output_file, prefixes=("POLIS_CLEAN_", f"{SLIP_COL}_CLEAN_", "FAC_INSURED_", "CERTIFICATE"))

    print(f"\n{'=' * 55}")
    print("  [OK] Selesai!")
    print(f"  Total baris output  : {len(df):,}")
    print(f"  Total kolom output  : {len(df.columns)}")
    print(f"  File disimpan di    : {output_file}")
    print(f"{'=' * 55}")


if __name__ == "__main__":
    process_data(INPUT_FILE, OUTPUT_FILE)

"""
================================================================================
 MATCHING KBRU
================================================================================
"""


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
from sqlalchemy import create_engine, text

# =============================================================================
# KONFIGURASI
# =============================================================================

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR   = os.path.join(SCRIPT_DIR, "output")

SUSPEND_FILE = os.path.join(DATA_DIR, "kbru_output_suspend.xlsx")
OSBAL_FILE   = os.path.join(DATA_DIR, "kbru_output_osbal.xlsx")
FACUL_FILE   = os.path.join(DATA_DIR, "kbru_output_facul.xlsx")
OUTPUT_FILE  = os.path.join(DATA_DIR, "final_output_v1.xlsx")

def _find_slipdb_file() -> str:
    kandidat = [
        os.path.join(DATA_DIR, "kbru_output_slipdb.xlsx"),
        os.path.join(DATA_DIR, "slipdb_clean_kbru.xlsx"),
        os.path.join(DATA_DIR, "ri slip.xlsx"),
        os.path.join(DATA_DIR, "slipdb.xlsx"),
    ]
    for p in kandidat:
        if os.path.exists(p):
            return p
    return kandidat[0]

SLIPDB_FILE = _find_slipdb_file()

# Kolom FAC CODE
OSBAL_FACODE_COL  = "CCOS_REF_CODE"
FACUL_FACODE_COL  = "FAC_CODE"
SLIPDB_FACODE_COL = "FAC_CODE"

# Kolom currency dan tanggal
SUSPEND_CURR_COL = "CURR ORI"
OSBAL_CURR_COL   = "CCOS_CURR"
SUSPEND_DATE_COL = "RECEIPT DATE"
OSBAL_DATE_COL   = "FAC_COM_DATE"

# Marker baris administratif yang dikecualikan
_EXCLUDED_REF_MARKERS = ["HUTANG PIUTANG", "DATA SUSPENSE"]
_DUMMY_VALUES = frozenset({"-", "--", "---", "0", "00", "000", "NONE", "NULL", "TBA", "NA", "N/A", "SELISIH", "NO", "NON", "."})

_MIN_TOKEN_LEN = 3
_TOKEN_RE      = re.compile(r"[^A-Z0-9]+")
_DELIMITER_RE  = re.compile(r"[,;|+\s]+")

# Cedant KBRU identifiers
KBRU_CEDANT_NAME = "PT. KBRU REINSURANCE BROKERS (EX. SIMAS)"
KBRU_SHRT_NAME   = "KBRU"

# Kolom suspend yang di-join jika ada row merging
_SUSPEND_JOIN_COLS = [
    "RECEIPT NO", "CREDIT NOTES", "DETAIL RINCIAN NO", "RECEIPT DATE",
    "CEDANT NAME", "CEDANT SHRT NAME",
    "INSURED_ORI", "INSURED_1", "INSURED_2",
    "CURR ORI", "CURR PAY", "AMOUNT PAY",
    "POLIS", "POLICY_CLEAN_1",
    "SLIP_NO", "SLIP_NO_CLN",
    "DESC 1", "DESC 2", "DESC 3", "DESC 4",
    "STATUS", "REC_TYPE",
]
_SUSPEND_SUM_COLS = ["AMOUNT ORI"]

# 38 Kolom standar output
FINAL_COLUMNS = [
    "CCOS_DOC_NO", "CCOS_REF_CODE",
    "RECEIPT NO", "CREDIT NOTES", "DETAIL RINCIAN NO", "RECEIPT DATE",
    "CEDANT NAME", "CEDANT SHRT NAME",
    "INSURED_ORI", "INSURED_1", "INSURED_2",
    "CURR ORI", "AMOUNT ORI", "AMOUNT_ORI_MIN1",
    "CURR PAY", "AMOUNT PAY",
    "CCOS_OR_BAL", "CCOS_BAL_DUE", "DIFERENCE",
    "FLAG_PROD",
    "POLIS",
    "POLICY_CLEAN_1", "SERTIF_CLEAN_1",
    "POLICY_CLEAN_2", "SERTIF_CLEAN_2",
    "POLICY_CLEAN_3", "SERTIF_CLEAN_3",
    "POLICY_CLEAN_4", "POLICY_CLEAN_5",
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
    text = s.strip().upper().replace("S/D", "SD")
    return re.sub(r"\s+", " ", text) if "  " in text else text

def _normalize(value) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    if isinstance(value, (float, np.floating)) and float(value).is_integer():
        return str(int(value))
    return _norm_cached(str(value))

def _clean_alphanumeric(s: str) -> str:
    if not s:
        return ""
    return re.sub(r"[^A-Z0-9]", "", str(s).upper())

def _tokenize(text: str) -> frozenset:
    if not text:
        return frozenset()
    return frozenset(t for t in _TOKEN_RE.split(text) if len(t) >= _MIN_TOKEN_LEN)

def _get_clean_cols(columns, prefix: str) -> list:
    if columns is None or len(columns) == 0:
        return []
    alt_map = {
        "clean polis": ["clean polis", "policy_clean", "polis_clean", "polis_cln", "policy_cln"],
        "clean slip": ["clean slip", "slip_clean", "fac_slip_clean", "slip_no_cln", "slip_cln"],
        "clean insured": ["clean insured", "insured_clean", "fac_insured_", "insured_1", "insured_2", "insured_3", "insured_4", "insured_5"],
    }
    prefixes = alt_map.get(prefix.lower(), [prefix.lower()])
    res = []
    for c in columns:
        clow = str(c).lower().strip()
        if any(clow.startswith(p) for p in prefixes):
            res.append(c)
    return res

def _is_excluded_ref_row(row: dict, polis_cols: list, slip_cols: list) -> bool:
    for col in list(polis_cols) + ["polis_ori", "FAC_POLICY_NO_ORI"] + list(slip_cols) + ["slip_ori", "FAC_SLIP_ORI"]:
        val = _normalize(row.get(col, ""))
        if val and any(m in val for m in _EXCLUDED_REF_MARKERS):
            return True
    return False

def _is_valid_token(s: str) -> bool:
    if not s or (isinstance(s, float) and pd.isna(s)):
        return False
    norm = _normalize(s)
    if norm in _DUMMY_VALUES:
        return False
    c = _clean_alphanumeric(norm)
    return len(c) >= 3 and c not in _DUMMY_VALUES

def _extract_polis_variants(val: str) -> list:
    if not _is_valid_token(val):
        return []
    s = _normalize(val)
    variants = set()

    # 1. Alphanumeric polos
    c = _clean_alphanumeric(s)
    if len(c) >= 6:
        variants.add(c)
        # Bentuk tanpa kata broker (KBRURE / SRB)
        no_broker = re.sub(r'KBRURE|SRB', '', c)
        if len(no_broker) >= 6:
            variants.add(no_broker)

    # 2. Pola registrasi polis KBRU: 00004903/FEQI/054.0000002-SRB/08/2023 atau 00000481.FPAR.002.000005.SRB
    m = re.search(r'0*(\d{3,8})[/.\s]+([A-Z]{3,5})[/.\s]+0*(\d+)[/.\s]+0*(\d+)', s)
    if m:
        doc_num, cls, p3, p4 = m.groups()
        variants.add(f"{doc_num}{cls}{p3}{p4}")
        variants.add(f"{doc_num}{cls}")
    else:
        m2 = re.search(r'0*(\d{3,8})[/.\s]+([A-Z]{3,5})', s)
        if m2:
            doc_num, cls = m2.groups()
            variants.add(f"{doc_num}{cls}")

    # 3. Multi-polis dengan '+'
    if '+' in s:
        for part in s.split('+'):
            if _is_valid_token(part):
                variants.update(_extract_polis_variants(part))

    return [v for v in variants if len(v) >= 3 and v not in _DUMMY_VALUES]

def _extract_slip_variants(val: str) -> list:
    if not _is_valid_token(val):
        return []
    s = _normalize(val)
    variants = set()

    variants.add(s)
    c = _clean_alphanumeric(s)
    if len(c) >= 3:
        variants.add(c)

    # Strip -01, -02 (installment suffix)
    base = re.sub(r'-\d+$', '', s)
    if base and base != s and _is_valid_token(base):
        variants.add(base)
        c_base = _clean_alphanumeric(base)
        if len(c_base) >= 3:
            variants.add(c_base)

    if '+' in s:
        for part in s.split('+'):
            if _is_valid_token(part):
                variants.update(_extract_slip_variants(part))

    return [v for v in variants if len(v) >= 3 and v not in _DUMMY_VALUES]

def _collect_clean_values(row: dict, cols: list, is_slip: bool = False, is_polis: bool = False) -> list:
    seen, result = set(), []
    for col in cols:
        raw_val = row.get(col, "")
        if not _is_valid_token(raw_val):
            continue
        val = _normalize(raw_val)
        if not val or val in _DUMMY_VALUES:
            continue

        if is_polis:
            for pv in _extract_polis_variants(val):
                if pv not in seen:
                    seen.add(pv)
                    result.append(pv)
        elif is_slip:
            for sv in _extract_slip_variants(val):
                if sv not in seen:
                    seen.add(sv)
                    result.append(sv)
        else:
            if val not in seen:
                seen.add(val)
                result.append(val)
    return result

def _collect_all_values(row: dict, clean_cols: list, ori_col: str, is_slip: bool = False, is_polis: bool = False) -> list:
    vals = _collect_clean_values(row, clean_cols, is_slip=is_slip, is_polis=is_polis)
    if not vals and ori_col:
        raw_val = row.get(ori_col, "")
        if _is_valid_token(raw_val):
            val_ori = _normalize(raw_val)
            if val_ori and val_ori not in _DUMMY_VALUES:
                if is_polis:
                    vals = _extract_polis_variants(val_ori)
                elif is_slip:
                    vals = _extract_slip_variants(val_ori)
                else:
                    vals = [val_ori]
    return list(set(vals))

def _ref_has_value(ref_row: dict, clean_cols: list, ori_col: str, query_values: list, is_slip: bool = False, is_polis: bool = False) -> bool:
    if not query_values:
        return False
    q_set = set(query_values)
    ref_cols = list(clean_cols)
    if ori_col and ori_col not in ref_cols:
        ref_cols.append(ori_col)
    ref_vals = _collect_clean_values(ref_row, ref_cols, is_slip=is_slip, is_polis=is_polis)
    return any(v in q_set for v in ref_vals)

# =============================================================================
# INDEXING & MATCH ENGINE
# =============================================================================


# =============================================================================
# INSURED MATCHING — WAHANA RULES (original six + three additional)
# KBRU cleaning remains separate. These functions only decide candidate matches.
# =============================================================================
def normalize_text(value):
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except (TypeError, ValueError):
        pass
    return re.sub(r"\s+", " ", str(value).strip().upper())

def clean_values(values):
    return list(dict.fromkeys(x for v in values if (x := normalize_text(v))))

INSURED_GENERIC_TOKENS = {
    "BANK", "PT", "CV", "TBK", "LTD", "LIMITED", "INC", "CORP",
    "CORPORATION", "COMPANY", "CO", "GROUP", "HOLDING", "INDONESIA", "PLANTATION", "REFINERY", "KEMITRAAN",
    "INDONESIAN", "INDUSTRY", "INDUSTRIES", "INTERNATIONAL", "PERSERO",
    "PERSEROAN", "TERBATAS", "FINANCE", "FINANCIAL", "ASURANSI"
}


def _insured_tokens(value):
    text = normalize_text(value)
    if not text:
        return []
    return [t for t in re.findall(r"[A-Z0-9]+", text) if t]


def _insured_meaningful(value):
    return {t for t in _insured_tokens(value)
            if len(t) >= 3 and t not in INSURED_GENERIC_TOKENS}


def _insured_compact(value):
    return "".join(_insured_tokens(value))


def _insured_bank_phrase_match(query, reference):
    """Nama bank utuh sebagai frasa, bukan kecocokan kata BANK saja.

    Berlaku dua arah: nama bank pendek dapat berada dalam insured panjang.
    Minimal BANK + dua token nama untuk menghindari false positive.
    """
    qtokens, rtokens = _insured_tokens(query), _insured_tokens(reference)
    for short, long in ((qtokens, rtokens), (rtokens, qtokens)):
        if len(short) < 3 or "BANK" not in short or len(short) > len(long):
            continue
        for start in range(len(long) - len(short) + 1):
            if long[start:start + len(short)] == short:
                return True
    return False


def insured_match_methods(query, reference):
    """Enam aturan Insured yang disepakati; tidak memakai TF-IDF untuk keputusan."""
    from difflib import SequenceMatcher
    q, r = normalize_text(query), normalize_text(reference)
    if not q or not r:
        return set()
    methods = set()
    if q == r:
        methods.add("INSURED_EXACT")
    qt, rt = _insured_meaningful(q), _insured_meaningful(r)
    common = qt & rt
    if len(common) >= 2 or _insured_bank_phrase_match(q, r):
        methods.add("INSURED_LIKE")
    if len(common) >= 2 and ("GROUP" in _insured_tokens(q) or "GROUP" in _insured_tokens(r)):
        methods.add("INSURED_GROUP")
    if _insured_compact(q) == _insured_compact(r):
        methods.add("INSURED_WORD_JOIN")

    # WORD_JOIN parsial: dua kata berurutan di satu sisi menjadi satu di sisi lain.
    qtokens, rtokens = _insured_tokens(q), _insured_tokens(r)
    joined = any(a+b in rtokens for a,b in zip(qtokens,qtokens[1:])) or any(
        a+b in qtokens for a,b in zip(rtokens,rtokens[1:]))
    if joined and common:
        methods.add("INSURED_LIKE + WORD_JOIN")

    # TYPO: satu substitusi/penambahan/penghapusan huruf pada satu token khas,
    # dengan token khas lain tetap identik ATAU kedua nama hanya satu token khas.
    def one_edit(x, y):
        if x == y or abs(len(x)-len(y)) > 1 or min(len(x),len(y)) < 5:
            return False
        if len(x) == len(y):
            return sum(a != b for a,b in zip(x,y)) == 1
        short, long = (x,y) if len(x) < len(y) else (y,x)
        return any(short == long[:i]+long[i+1:] for i in range(len(long)))
    qonly, ronly = qt-common, rt-common
    if (len(qonly) == 1 and len(ronly) == 1 and
        (common or (len(qt) == 1 and len(rt) == 1)) and
        one_edit(next(iter(qonly)), next(iter(ronly)))):
        methods.add("INSURED_TYPO")
    return methods


# ============================================================
# THREE APPROVED, STRICT ADDITIONAL INSURED RULES
# Keep insured_match_methods() (the original six rules) UNCHANGED.
# ============================================================
INSURED_VERIFIED_ALIASES = {
    frozenset(("TEMAS", "TEMAS SHIPPING")),
}

# Controlled location suffixes only. Extend only after business verification.
# Whole suffix must match; arbitrary appended words never qualify.
INSURED_VERIFIED_LOCATIONS = {
    "JAKARTA", "TANGERANG", "JAKARTA TANGERANG", "SURABAYA",
    "BANDUNG", "BEKASI", "DEPOK", "BOGOR", "SEMARANG", "MEDAN",
    "MAKASSAR", "DENPASAR", "BALIKPAPAN", "PALEMBANG",
    "YOGYAKARTA", "BATAM", "PEKANBARU", "MALANG", "SOLO",
    "JAKARTA BARAT", "JAKARTA TIMUR", "JAKARTA SELATAN",
    "JAKARTA UTARA", "JAKARTA PUSAT",
}


def insured_additional_match_methods(query, reference):
    """Extra rules, isolated from the six diagnostic methods."""
    q, r = normalize_text(query), normalize_text(reference)
    if not q or not r:
        return set()
    methods = set()
    qtokens, rtokens = _insured_tokens(q), _insured_tokens(r)

    if frozenset((" ".join(qtokens), " ".join(rtokens))) in INSURED_VERIFIED_ALIASES:
        methods.add("INSURED_ALIAS_VERIFIED")

    # Only full name + a strictly verified location suffix; prefix must be exact.
    for shorter, longer in ((qtokens, rtokens), (rtokens, qtokens)):
        if len(longer) <= len(shorter) or longer[:len(shorter)] != shorter:
            continue
        suffix = " ".join(longer[len(shorter):])
        if suffix in INSURED_VERIFIED_LOCATIONS:
            methods.add("INSURED_FAMILY_NAME")

    # Narrow insurance phrase: ASURANSI + identical distinctive token(s),
    # with only the trailing INDONESIA token additionally permitted.
    for short, long in ((qtokens, rtokens), (rtokens, qtokens)):
        if (len(short) >= 2 and short[0] == "ASURANSI" and
            len(long) == len(short) + 1 and
            long[:-1] == short and long[-1] == "INDONESIA" and
            any(token not in INSURED_GENERIC_TOKENS for token in short[1:])):
            methods.add("INSURED_INSURANCE_PHRASE")
    return methods


def insured_all_match_methods(query, reference):
    """Union of all original six + approved extras; never first-hit wins."""
    return insured_match_methods(query, reference) | insured_additional_match_methods(query, reference)


def insured_like_match(query, reference):
    """Compatibility wrapper: true bila salah satu metode Insured cocok."""
    return bool(insured_match_methods(query, reference))


def _insured_candidate_rows(query, token_rows, compact_rows, reference_matrix):
    """Cari kandidat melalui token, word join, dan typo tanpa full scan per receipt."""
    from difflib import SequenceMatcher
    candidates = set(compact_rows.get(_insured_compact(query), ()))
    tokens = _insured_meaningful(query)
    for token in tokens:
        candidates.update(token_rows.get(token, ()))
        # Token ejaan mirip; batasi dengan huruf awal dan selisih panjang.
        for ref_token, rows in token_rows.items():
            if (ref_token != token and ref_token[:1] == token[:1]
                and abs(len(ref_token) - len(token)) <= 2
                and min(len(ref_token), len(token)) >= 5
                and SequenceMatcher(None, token, ref_token).ratio() >= 0.80):
                candidates.update(rows)
    # Kasus join parsial: TRI SAKTI vs TRISAKTI.
    qtokens = _insured_tokens(query)
    # BANK NEGARA INDONESIA hanya punya satu token khas (NEGARA),
    # sehingga indeks token tetap harus bisa mengambil kandidatnya.
    # Bila tidak ada token khas sama sekali, gunakan BANK sebagai anchor;
    # validasi frasa lengkap tetap dilakukan oleh insured_match_methods.
    if "BANK" in qtokens and not tokens:
        candidates.update(token_rows.get("BANK", ()))
    for a, b in zip(qtokens, qtokens[1:]):
        candidates.update(token_rows.get(a + b, ()))
    return [idx for idx in candidates
            if any(insured_all_match_methods(query, r) for r in clean_values(reference_matrix[idx]))]



def _kbru_insured_search_index(rows, insured_cols, polis_cols, slip_cols):
    """Index inverted token + compact names for WAHANA matching; exclude invalid refs."""
    token_rows, compact_rows, name_rows, first_letter_rows = {}, {}, {}, {}
    for idx, row in enumerate(rows):
        if _is_excluded_ref_row(row, polis_cols, slip_cols):
            continue
        names = clean_values(row.get(c, "") for c in insured_cols)
        if not names:
            continue
        name_rows[idx] = names
        for name in names:
            compact_rows.setdefault(_insured_compact(name), set()).add(idx)
            for token in set(_insured_tokens(name)):
                token_rows.setdefault(token, set()).add(idx)
                if len(token) >= 5:
                    first_letter_rows.setdefault(token[0], set()).add(token)
    return token_rows, compact_rows, name_rows, first_letter_rows

def _kbru_insured_hits(queries, search_index, mode):
    """Exact or WAHANA's complete nine-rule decision over indexed candidates."""
    token_rows, compact_rows, name_rows, first_letter_rows = search_index
    hits = set()
    for q in clean_values(queries):
        candidates = set(compact_rows.get(_insured_compact(q), ()))
        if mode != "EXACT":
            qtokens = _insured_tokens(q)
            for tok in qtokens:
                candidates.update(token_rows.get(tok, ()))
                if len(tok) >= 5:
                    for ref_tok in first_letter_rows.get(tok[0], ()):
                        if abs(len(tok) - len(ref_tok)) <= 1 and _wahana_one_edit(tok, ref_tok):
                            candidates.update(token_rows.get(ref_tok, ()))
            for a, b in zip(qtokens, qtokens[1:]):
                candidates.update(token_rows.get(a+b, ()))
        for idx in candidates:
            if any((q == r if mode == "EXACT" else bool(insured_all_match_methods(q, r)))
                   for r in name_rows[idx]):
                hits.add(idx)
    return hits

def _wahana_one_edit(x, y):
    """Candidate retrieval only; final decision is insured_all_match_methods."""
    if x == y or abs(len(x)-len(y)) > 1 or min(len(x),len(y)) < 5:
        return False
    if len(x) == len(y):
        return sum(a != b for a,b in zip(x,y)) == 1
    short,long = (x,y) if len(x)<len(y) else (y,x)
    return any(short == long[:i]+long[i+1:] for i in range(len(long)))

def _build_exact_index(rows: list, cols: list, is_slip: bool = False, is_polis: bool = False, excluded: set = None) -> dict:
    index: dict = {}
    for i, row in enumerate(rows):
        if excluded and i in excluded:
            continue
        vals = _collect_clean_values(row, cols, is_slip=is_slip, is_polis=is_polis)
        for val in vals:
            index.setdefault(val, []).append(i)
    return index

def _build_facode_index(rows: list, col: str, excluded: set = None) -> dict:
    index: dict = {}
    for i, row in enumerate(rows):
        if excluded and i in excluded:
            continue
        fac = _normalize(row.get(col, ""))
        if fac:
            index.setdefault(fac, []).append(i)
    return index

def _build_token_index(rows: list, clean_cols: list, excluded: set = None) -> tuple:
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

def _build_lookup(rows: list, polis_cols: list, slip_cols: list, insured_cols: list,
                  facode_col: str = None, label: str = "") -> tuple:
    if not rows:
        if label:
            print(f"  {label:<6} done (0.0s) [kosong]", flush=True)
        return ({}, {}, {}, ({}, {}), ({}, {}), ({}, {}), {})

    t0 = time.perf_counter()
    excluded = {i for i, r in enumerate(rows) if _is_excluded_ref_row(r, polis_cols, slip_cols)}
    facode_index = _build_facode_index(rows, facode_col, excluded=excluded) if facode_col else {}
    lookup = (
        _build_exact_index(rows, slip_cols,    is_slip=True,  is_polis=False, excluded=excluded), # [0]
        _build_exact_index(rows, polis_cols,   is_slip=False, is_polis=True,  excluded=excluded), # [1]
        _build_exact_index(rows, insured_cols, is_slip=False, is_polis=False, excluded=excluded), # [2]
        _build_token_index(rows, slip_cols,    excluded=excluded),                # [3]
        _build_token_index(rows, polis_cols,   excluded=excluded),                # [4]
        _build_token_index(rows, insured_cols, excluded=excluded),                # [5]
        facode_index,                                                             # [6]
    )
    elapsed = time.perf_counter() - t0
    if label:
        exc_info = f" - {len(excluded):,} baris dikecualikan" if excluded else ""
        print(f"  {label:<6} done ({elapsed:.1f}s){exc_info}", flush=True)
    return lookup

def _exact_match(query_values: list, index: dict) -> set:
    return {i for v in query_values if v and v in index for i in index[v]}

def _resolve_facode(ref_rows_matched: list, facode_col: str, facode_osbal_idx: dict, osbal_rows: list) -> tuple:
    fac_codes = set()
    for r in ref_rows_matched:
        fac = _normalize(r.get(facode_col, ""))
        if fac:
            fac_codes.add(fac)
    if not fac_codes:
        return [], False
    osbal_indices: set = set()
    for fac in fac_codes:
        if fac in facode_osbal_idx:
            osbal_indices.update(facode_osbal_idx[fac])
    res = list(osbal_indices)
    return res, bool(res)

# =============================================================================
# REMATCH STRICT & OUTPUT AGGREGATION
# =============================================================================

def _aggregate_ccos(osbal_rows: list, all_fac_osbal_rows: list = None) -> dict:
    fac_codes = {str(r.get("CCOS_REF_CODE", "")).strip() for r in osbal_rows if str(r.get("CCOS_REF_CODE", "")).strip()}
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
        return total

    return {
        "CCOS_DOC_NO":   join_unique("CCOS_DOC_NO"),
        "CCOS_REF_CODE": join_unique("CCOS_REF_CODE"),
        "CCOS_OR_BAL":   np.nan if multi_fac else join_unique("CCOS_OR_BAL"),
        "CCOS_BAL_DUE":  sum_col("CCOS_BAL_DUE", bal_due_rows),
    }

def _facode_label(osbal_rows: list, source: str) -> str:
    if not osbal_rows or not osbal_rows[0]:
        return ""
    col = OSBAL_FACODE_COL if source == "OSBAL" else (
          SLIPDB_FACODE_COL if source == "SLIPDB" else FACUL_FACODE_COL)
    codes = {str(r.get(col, "")).strip() for r in osbal_rows if str(r.get(col, "")).strip()}
    return "facode lebih dari 1" if len(codes) > 1 else next(iter(codes), "")

def _clean_base_match(scenario: str) -> str:
    s = str(scenario or "").strip()
    for to_remove in [" (Beda Currency)", " (Beda Periode)", " (> 1 Fac)"]:
        s = s.replace(to_remove, "")
    s = s.strip()
    if " only + " in s:
        s = s.replace(" only + ", " + ")
    return s

def _format_final_scenario(
    base_scenario: str, source: str,
    is_beda_curr: bool, is_beda_per: bool, num_fac: int
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
    elif num_fac > 1:
        return f"{base} (> 1 Fac)"
    else:
        return base

def _build_output_row(
    suspend_row: dict, source: str, scenario: str, osbal_rows: list,
    osbal_count: int = 0, resolved: bool = True,
    suspend_count: int = 1, all_fac_osbal_rows: list = None
) -> dict:
    has_match = bool(source and osbal_rows and osbal_rows[0])
    has_osbal = has_match and (source in ("OSBAL", "SLIPDB", "FACUL") or resolved)

    ccos = (
        _aggregate_ccos(osbal_rows, all_fac_osbal_rows=all_fac_osbal_rows)
        if has_osbal
        else {"CCOS_DOC_NO": "", "CCOS_REF_CODE": "", "CCOS_OR_BAL": np.nan, "CCOS_BAL_DUE": np.nan}
    )
    # Tetap tampilkan CCOS_BAL_DUE yang berasal dari OSBAL untuk audit,
    # walaupun flag akhirnya Beda Currency / Beda Periode.

    final_scenario = "Unmatching" if not source or scenario in ("Unmatching", "UNMATCHED") else scenario

    return {
        "CCOS_DOC_NO":            ccos["CCOS_DOC_NO"] if has_osbal else "",
        "CCOS_REF_CODE":          ccos["CCOS_REF_CODE"],
        "RECEIPT NO":             suspend_row.get("RECEIPT NO",        ""),
        "CREDIT NOTES":           suspend_row.get("CREDIT NOTES",      ""),
        "DETAIL RINCIAN NO":      suspend_row.get("DETAIL RINCIAN NO", ""),
        "RECEIPT DATE":           suspend_row.get("RECEIPT DATE",      ""),
        "CEDANT NAME":            suspend_row.get("CEDANT NAME",       ""),
        "CEDANT SHRT NAME":       suspend_row.get("CEDANT SHRT NAME",  ""),
        "INSURED_ORI":            suspend_row.get("INSURED_ORI") or suspend_row.get("insured_ori") or "",
        "INSURED_1":              suspend_row.get("INSURED_1") or suspend_row.get("clean insured 1") or "",
        "INSURED_2":              suspend_row.get("INSURED_2") or suspend_row.get("clean insured 2") or "",
        "CURR ORI":               suspend_row.get("CURR ORI",          ""),
        "AMOUNT ORI":             suspend_row.get("AMOUNT ORI",        ""),
        "CURR PAY":               suspend_row.get("CURR PAY",          ""),
        "AMOUNT PAY":             suspend_row.get("AMOUNT PAY",        ""),
        "CCOS_OR_BAL":            ccos["CCOS_OR_BAL"],
        "CCOS_BAL_DUE":           ccos["CCOS_BAL_DUE"],
        "POLIS":                  suspend_row.get("POLIS_ORI") or suspend_row.get("POLIS") or suspend_row.get("polis_ori") or "",
        "POLICY_CLEAN_1":         suspend_row.get("POLIS_CLEAN_1") or suspend_row.get("POLICY_CLEAN_1") or suspend_row.get("clean polis 1") or "",
        "SERTIF_CLEAN_1":         "", # Tidak menggunakan sertifikat
        "POLICY_CLEAN_2":         suspend_row.get("POLIS_CLEAN_2") or suspend_row.get("POLICY_CLEAN_2") or "",
        "SERTIF_CLEAN_2":         "",
        "POLICY_CLEAN_3":         suspend_row.get("POLIS_CLEAN_3") or suspend_row.get("POLICY_CLEAN_3") or "",
        "SERTIF_CLEAN_3":         "",
        "POLICY_CLEAN_4":         suspend_row.get("POLIS_CLEAN_4") or suspend_row.get("POLICY_CLEAN_4") or "",
        "POLICY_CLEAN_5":         suspend_row.get("POLIS_CLEAN_5") or suspend_row.get("POLICY_CLEAN_5") or "",
        "SLIP_NO":                suspend_row.get("SLIP_NO_ORI") or suspend_row.get("SLIP_NO") or suspend_row.get("slip_ori") or "",
        "SLIP_NO_CLN":            suspend_row.get("SLIP_CLEAN_1") or suspend_row.get("SLIP_NO_CLN") or suspend_row.get("clean slip 1") or "",
        "DESC 1":                 suspend_row.get("DESC 1",            ""),
        "DESC 2":                 suspend_row.get("DESC 2",            ""),
        "DESC 3":                 suspend_row.get("DESC 3",            ""),
        "DESC 4":                 suspend_row.get("DESC 4",            ""),
        "STATUS":                 suspend_row.get("STATUS",            ""),
        "REC_TYPE":               suspend_row.get("REC_TYPE",          ""),
        "SKENARIO":               final_scenario,
        "_OSBAL_ROW_COUNT":       osbal_count,
        "_SUSPEND_ROW_COUNT":     suspend_count,
    }

def _merge_suspend_rows(rows: list) -> dict:
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

def _apply_wahana_fac_accumulation(df, raw_results, osbal_rows, facode_osbal_idx):
    """WAHANA: untuk tepat satu FAC CODE yang lolos currency/periode,
    agregasikan semua baris Suspend resolved dan SELURUH baris OSBAL
    pada FAC CODE itu, dengan indeks OSBAL unik (tanpa double count).
    Output tetap satu baris per transaksi Suspend asli.
    """
    from collections import defaultdict
    from decimal import Decimal, InvalidOperation

    def as_decimal(value):
        if value is None or pd.isna(value):
            return None
        raw = str(value).strip().replace(" ", "")
        if not raw or raw.lower() in ("nan", "none", "null"):
            return None
        if "," in raw and "." in raw:
            raw = raw.replace(".", "").replace(",", ".") if raw.rfind(",") > raw.rfind(".") else raw.replace(",", "")
        elif "," in raw:
            raw = raw.replace(",", ".")
        try:
            result = Decimal(raw)
            return result if result.is_finite() else None
        except (InvalidOperation, TypeError, ValueError):
            return None

    suspend_group = defaultdict(list)
    for i, result in enumerate(raw_results):
        scenario = str(result.get("scenario", ""))
        codes = result.get("fac_codes") or set()
        if (result.get("source") and len(codes) == 1
                and "Beda Currency" not in scenario
                and "Beda Periode" not in scenario
                and scenario.lower() not in ("unmatching", "unmatched")):
            suspend_group[next(iter(codes))].append(i)

    df["_GROUP_AMOUNT_TOTAL"] = None
    for fac, suspend_indices in suspend_group.items():
        # Nilai Suspend dihitung dari baris SUMBER, bukan dari baris output
        suspend_values = [as_decimal(raw_results[i]["suspend"].get("AMOUNT ORI")) for i in suspend_indices]
        suspend_values = [-v for v in suspend_values if v is not None]
        amount_total = sum(suspend_values, Decimal("0")) if suspend_values else None

        # Ambil SELURUH baris OSBAL dari indeks FAC CODE, bukan subset matching
        osbal_indices = sorted(set(facode_osbal_idx.get(fac, [])))
        osbal_values = [as_decimal(osbal_rows[j].get("CCOS_BAL_DUE")) for j in osbal_indices]
        osbal_values = [v for v in osbal_values if v is not None]
        bal_due_total = sum(osbal_values, Decimal("0")) if osbal_values else None

        for i in suspend_indices:
            df.at[i, "_GROUP_AMOUNT_TOTAL"] = amount_total
            df.at[i, "CCOS_BAL_DUE"] = float(bal_due_total) if bal_due_total is not None else float("nan")
            df.at[i, "_SUSPEND_ROW_COUNT"] = len(suspend_indices)
            df.at[i, "_OSBAL_ROW_COUNT"] = len(osbal_indices)
    print(f"  [AKUMULASI] {len(suspend_group):,} FAC CODE resolved; "
          f"{sum(len(v) for v in suspend_group.values()):,} baris Suspend dihitung per FAC CODE.", flush=True)
    return df


def _compute_derived_cols(df: pd.DataFrame) -> pd.DataFrame:
    """Financial WAHANA-like: status prioritas lebih tinggi, nominal dihitung setelah valid FAC."""
    if df.empty:
        for col in ("AMOUNT_ORI_MIN1", "DIFERENCE", "FLAG_PROD"):
            df[col] = pd.Series(dtype="object")
        return df

    from decimal import Decimal, InvalidOperation

    def money(v):
        if pd.isna(v) or str(v).strip().lower() in ("", "nan", "none"):
            return None
        raw = str(v).strip().replace(" ", "")
        if "," in raw and "." in raw:
            raw = raw.replace(".", "").replace(",", ".") if raw.rfind(",") > raw.rfind(".") else raw.replace(",", "")
        elif "," in raw:
            raw = raw.replace(",", ".")
        try:
            x = Decimal(raw)
            return x if x.is_finite() else None
        except (InvalidOperation, ValueError):
            return None

    amounts = [money(v) for v in df["AMOUNT ORI"]]
    balances = [money(v) for v in df["CCOS_BAL_DUE"]]
    # AMOUNT_ORI_MIN1 = minus amount asli, sama dengan rumus lama KBRU.
    negative = [(-x if x is not None else None) for x in amounts]
    # Hasil akumulasi FAC CODE menggantikan nilai per-baris HANYA pada grup resolved.
    if "_GROUP_AMOUNT_TOTAL" in df.columns:
        negative = [group if group is not None else per_row
                    for group, per_row in zip(df["_GROUP_AMOUNT_TOTAL"], negative)]
    df["AMOUNT_ORI_MIN1"] = [float(x) if x is not None else float('nan') for x in negative]
    df["DIFERENCE"] = [float(a-b) if a is not None and b is not None else float('nan') for a,b in zip(negative, balances)]
    flags = []
    for i, (a, b) in enumerate(zip(negative, balances)):
        row = df.iloc[i]
        scenario = str(row.get("SKENARIO", "") or "").strip()
        refs = str(row.get("CCOS_REF_CODE", "") or "").strip()
        if scenario.upper() in ("", "UNMATCHED", "UNMATCHING") or not refs:
            flags.append("Unmatching")
        elif "Beda Currency" in scenario:
            flags.append("Beda Currency")
        elif "Beda Periode" in scenario:
            flags.append("Beda Periode")
        elif "> 1 Fac" in scenario or len([x for x in refs.split(",") if x.strip()]) > 1:
            flags.append("Matching >1 fac code")
        elif a is None:
            flags.append("Amount ORI tidak dapat dihitung")
        elif b is None:
            flags.append("CCOS_BAL_DUE tidak dapat dihitung")
        elif a != 0 and b == 0:
            flags.append("New Entry total")
        elif a > b:
            flags.append("New Entry sebagian")
        else:
            accumulated = int(row.get("_OSBAL_ROW_COUNT", 0) or 0) > 1 or int(row.get("_SUSPEND_ROW_COUNT", 1) or 1) > 1
            if a == b:
                flags.append("Adjustment total dengan akumulasi" if accumulated else "Adjustment total tanpa akumulasi")
            else:
                flags.append("Adjustment sebagian dengan akumulasi" if accumulated else "Adjustment sebagian tanpa akumulasi")
    df["FLAG_PROD"] = flags
    return df

def _load_excel(path: str, label: str, optional: bool = False) -> tuple:
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

# =============================================================================
# PIPELINE UTAMA V1
# =============================================================================

def run() -> None:
    t_start = time.perf_counter()
    print("\n" + "=" * 60, flush=True)
    print("  PRODUCTION SCRIPT — SUSPEND MATCHING (KBRU) — V3", flush=True)
    print("  Urutan Matching: SLIP -> POLIS -> INSURED", flush=True)
    print("  Basic Rules: Tanpa Sertifikat, Tanpa Periode 2 / Lineslip Offset", flush=True)
    print("=" * 60, flush=True)

    # [1] Load data
    print("\n[1/5] Loading data ...", flush=True)
    suspend_rows, suspend_cols = _load_excel(SUSPEND_FILE, "SUSPEND")
    osbal_rows,   osbal_cols   = _load_excel(OSBAL_FILE,   "OSBAL")
    facul_rows,   facul_cols   = _load_excel(FACUL_FILE,   "FACUL")
    slipdb_rows,  slipdb_cols  = _load_excel(SLIPDB_FILE,  "SLIPDB", optional=True)

    for label, rows in (("SUSPEND", suspend_rows), ("OSBAL", osbal_rows), ("FACUL", facul_rows)):
        if not rows:
            raise ValueError(f"{label} tidak ada/kosong. Periksa file Excel pada {DATA_DIR}")

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
    print("\n[2/5] Detecting clean columns ...", flush=True)
    polis_sus   = _get_clean_cols(suspend_cols, "clean polis")
    slip_sus    = _get_clean_cols(suspend_cols, "clean slip")
    insured_sus = _get_clean_cols(suspend_cols, "clean insured")

    for c in ["POLIS_ORI", "POLIS"]:
        if c in suspend_cols and c not in polis_sus:
            polis_sus.append(c)
    for c in ["SLIP_NO_ORI", "SLIP_NO"]:
        if c in suspend_cols and c not in slip_sus:
            slip_sus.append(c)
    if "INSURED_ORI" in suspend_cols and "INSURED_ORI" not in insured_sus:
        insured_sus.append("INSURED_ORI")

    polis_osbal   = _get_clean_cols(osbal_cols, "clean polis")
    slip_osbal    = _get_clean_cols(osbal_cols, "clean slip")
    insured_osbal = _get_clean_cols(osbal_cols, "clean insured")

    for c in ["CLSDT_POLICY_NO", "FAC_POLICY_NO_ORI", "FAC_POLICY_NO"]:
        if c in osbal_cols and c not in polis_osbal:
            polis_osbal.append(c)
    for c in ["CLSDT_SLIP_NO", "FAC_SLIP_ORI", "FAC_SLIP"]:
        if c in osbal_cols and c not in slip_osbal:
            slip_osbal.append(c)

    polis_facul   = _get_clean_cols(facul_cols, "clean polis")
    slip_facul    = _get_clean_cols(facul_cols, "clean slip")
    insured_facul = _get_clean_cols(facul_cols, "clean insured")

    for c in ["FAC_POLICY_NO"]:
        if c in facul_cols and c not in polis_facul:
            polis_facul.append(c)
    for c in ["FAC_SLIP"]:
        if c in facul_cols and c not in slip_facul:
            slip_facul.append(c)

    polis_slipdb   = _get_clean_cols(slipdb_cols, "clean polis")
    slip_slipdb    = _get_clean_cols(slipdb_cols, "clean slip")
    insured_slipdb = _get_clean_cols(slipdb_cols, "clean insured")

    # [3] Bangun lookup index
    print("\n[3/5] Building lookup indexes ...", flush=True)
    lookup_osbal  = _build_lookup(osbal_rows,  polis_osbal,  slip_osbal,  insured_osbal,  OSBAL_FACODE_COL,  label="OSBAL")
    lookup_facul  = _build_lookup(facul_rows,  polis_facul,  slip_facul,  insured_facul,  FACUL_FACODE_COL,  label="FACUL")
    lookup_slipdb = _build_lookup(slipdb_rows, polis_slipdb, slip_slipdb, insured_slipdb, SLIPDB_FACODE_COL, label="SLIPDB")

    facode_osbal_idx = lookup_osbal[6]

    # Index INSURED separately from KBRU's ordinary slip/polis indexes.
    insured_search = {
        "OSBAL": _kbru_insured_search_index(osbal_rows, insured_osbal, polis_osbal, slip_osbal),
        "SLIPDB": _kbru_insured_search_index(slipdb_rows, insured_slipdb, polis_slipdb, slip_slipdb),
        "FACUL": _kbru_insured_search_index(facul_rows, insured_facul, polis_facul, slip_facul),
    }

    # [4] Matching per baris
    total = len(suspend_rows)
    print(f"\n[4/5] Matching {total:,} suspend rows ...", flush=True)
    raw_results = []

    for n, sus in enumerate(suspend_rows, 1):
        if n % 500 == 0 or n == total:
            print(f"    Memproses baris {n:,} / {total:,} ...", flush=True)

        if str(sus.get("STATUS", "")).strip().upper() == "MATCHING":
            continue

        effective_date = sus.get("_sus_date_parsed")
        sus_curr = _normalize(sus.get(SUSPEND_CURR_COL, ""))

        matched_osbal_idx = set()
        source = None
        scenario = ""
        resolved = False

        # KBRU: SUSPEND POLIS -> referensi SLIP; SUSPEND SLIP -> referensi POLIS.
        insured_vals = _collect_clean_values(sus, insured_sus)
        sus_slip_as_polis = _collect_clean_values(sus, slip_sus, is_polis=True)
        # Khusus pencarian SLIP dari Suspend: hapus akhiran EXT + angka (EXT1, EXT2, dst).
        # Nilai asli Suspend/output tidak diubah; aturan EXACT/EXTENSION tetap sama.
        sus_polis_as_slip = list(dict.fromkeys(
            re.sub(r"EXT\d+$", "", v)
            for v in _collect_clean_values(sus, polis_sus, is_slip=True)
        ))
        sources = [
            ("OSBAL", osbal_rows, lookup_osbal, OSBAL_FACODE_COL),
            ("SLIPDB", slipdb_rows, lookup_slipdb, SLIPDB_FACODE_COL),
            ("FACUL", facul_rows, lookup_facul, FACUL_FACODE_COL),
        ]
        methods = [("Slip", "EXACT"), ("Slip", "EXTENSION"),
                   ("Polis", "EXACT"), ("Insured", "WAHANA_RULES")]

        def get_code_hits(src, method, mode):
            name, rows, lookup, fac_col = src
            if not rows:
                return set()
            if method == "Insured":
                hits = _kbru_insured_hits(insured_vals, insured_search[name], "WAHANA_RULES")
            elif method == "Slip":
                hits = _exact_match(sus_polis_as_slip, lookup[0]) if mode == "EXACT" else set()
                if mode == "EXTENSION":
                    # Prefix lengkap, tambahan HARUS digit, minimal 8 karakter.
                    # Tidak mencocokkan substring di posisi tengah nomor.
                    from bisect import bisect_left
                    keys = sorted(lookup[0])
                    for q in sus_polis_as_slip:
                        if len(q) < 8 or not any(ch.isdigit() for ch in q):
                            continue
                        at = bisect_left(keys, q)
                        while at < len(keys) and keys[at].startswith(q):
                            ref = keys[at]
                            if len(ref) > len(q) and ref[len(q):].isdigit():
                                hits.update(lookup[0][ref])
                            at += 1
            elif method == "Polis":
                hits = _exact_match(sus_slip_as_polis, lookup[1]) if mode == "EXACT" else set()
            else:
                hits = set()
            # Kode yang berasal dari RI/Slip ataupun FACUL WAJIB ada di OSBAL.
            return {fc for i in hits
                    if (fc := _normalize(rows[i].get(fac_col, ""))) in facode_osbal_idx}

        # =========================================================
        # FASE 1: BASE MATCHING
        # Prioritas SUMBER: OSBAL -> RI/SLIP -> FACUL
        # Prioritas METODE: SLIP EXACT -> SLIP EXTENSION
        #                  -> POLIS EXACT -> INSURED
        # =========================================================
        # BASE MATCHING: SOURCE FIRST, METHOD SECOND. Begitu ada metode yang
        # menghasilkan kode valid, berhenti mencari BASE dan lakukan wajib.
        active = set()
        base_source = base_method = None
        for src in sources:
            if not src[1]:
                continue
            for method, mode in methods:
                found = get_code_hits(src, method, mode)
                if found:
                    active = set(found)
                    base_source, base_method = src[0], method
                    break
            if active:
                break

        is_beda_curr = False
        is_beda_per = False
        contributing = []
        source = base_source
        if active:
            # =========================================================
            # FASE 2: NARROWING WAJIB
            # =========================================================
            # NARROWING WAJIB: semua atribut Currency/Periode HANYA dari OSBAL.
            # Satu FAC CODE dapat memiliki beberapa baris OSBAL: penilaian
            # Currency dan Periode harus menggunakan baris yang konsisten.
            # Jika sebuah tahap tidak menyisakan baris, simpan diagnosis,
            # pertahankan kandidat sebelum tahap gagal, dan lanjutkan narrowing.
            candidate_rows = {i for fc in active for i in facode_osbal_idx.get(fc, ())}
            assert candidate_rows, "FAC CODE base tidak ditemukan pada indeks OSBAL"

            if sus_curr:
                currency_rows = {i for i in candidate_rows
                                 if _normalize(osbal_rows[i].get(OSBAL_CURR_COL, "")) == sus_curr}
                if currency_rows:
                    active = {_normalize(osbal_rows[i].get(OSBAL_FACODE_COL, ""))
                              for i in currency_rows}
                    candidate_rows = currency_rows
                else:
                    is_beda_curr = True
                    # Tidak membuang kandidat; tetap cek Periode ke OSBAL.

            if pd.notna(effective_date):
                period_rows = {i for i in candidate_rows
                               if (pd.isna(osbal_rows[i].get("_com_date_parsed"))
                                   or effective_date >= osbal_rows[i].get("_com_date_parsed"))}
                if period_rows:
                    active = {_normalize(osbal_rows[i].get(OSBAL_FACODE_COL, ""))
                              for i in period_rows}
                    candidate_rows = period_rows
                else:
                    is_beda_per = True
                    # Pertahankan kandidat untuk narrowing tambahan/lintas sumber.

            def narrow(src_name, method, mode):
                src = next(x for x in sources if x[0] == src_name)
                overlap = active.intersection(get_code_hits(src, method, mode))
                return overlap if 0 < len(overlap) < len(active) else None

            if len(active) > 1:
                # =========================================================
                # FASE 3: NARROWING TAMBAHAN
                # =========================================================
                # Tambahan OSBAL: POLIS EXACT -> INSURED (tidak mengulang SLIP).
                plan = [("OSBAL", "Polis", "EXACT"),
                        ("OSBAL", "Insured", "WAHANA_RULES")]
                # =========================================================
                # FASE 4: NARROWING LINTAS SUMBER
                # =========================================================
                # Lintas sumber hanya jika masih >1: RI/Slip SEBELUM FACUL.
                for src_name in ("SLIPDB", "FACUL"):
                    plan.extend((src_name, met, mode) for met, mode in methods)
                for name, method, mode in plan:
                    if len(active) == 1:
                        break
                    shrunk = narrow(name, method, mode)
                    if shrunk is not None:
                        active = shrunk
                        if method != base_method and method not in contributing:
                            contributing.append(method)

            # Fac code final harus didapat dari ACTIVE terakhir, bukan kandidat awal.
            matched_osbal_idx = {i for fc in active for i in facode_osbal_idx[fc]}
            scenario_methods = [base_method] + contributing
            scenario_methods = [x for x in ("Slip", "Polis", "Insured") if x in scenario_methods]
            scenario = " + ".join(scenario_methods)
            if len(scenario_methods) == 1:
                scenario += " only"
        else:
            matched_osbal_idx = set()
            scenario = "Unmatching"

        fac_codes = set(active)
        if not fac_codes.issubset(facode_osbal_idx):
            raise AssertionError("Terdapat FAC CODE final yang tidak ada di OSBAL")
        final_scenario = _format_final_scenario(
            scenario, source, is_beda_curr, is_beda_per, len(fac_codes)
        )

        raw_results.append({
            "suspend":       sus,
            "source":        source if final_scenario != "Unmatching" else None,
            "scenario":      final_scenario,
            "osbal_idx":     list(matched_osbal_idx) if final_scenario != "Unmatching" else [],
            "osbal_count":   len(matched_osbal_idx) if final_scenario != "Unmatching" else 0,
            "resolved":      bool(len(fac_codes) == 1 and not is_beda_curr and not is_beda_per),
            "fac_codes":     fac_codes,
            "suspend_count": 1,
        })

    # [5] Build output DataFrame: seluruh CCOS_BAL_DUE dari OSBAL, bukan FACUL/RI
    print(f"\n[5/5] Building output DataFrame ({len(raw_results):,} rows) ...", flush=True)
    output_rows = []
    for r in raw_results:
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
            resolved=r["resolved"], suspend_count=1,
            all_fac_osbal_rows=all_fac_osbal_rows,
        )
        output_rows.append(out)

    df = pd.DataFrame(output_rows)
    if df.empty:
        raise ValueError("Tidak ada output matching; cek data SUSPEND dan mapping kolom.")

    # Format numeric
    for col in ["AMOUNT ORI", "CCOS_OR_BAL", "CCOS_BAL_DUE"]:
        s = df[col]
        if s.dtype == object:
            s = s.astype(str).str.strip().str.replace(r'\s+', '', regex=True)
            has_comma = s.str.contains(',', na=False)
            s_eu = s.str.replace('.', '', regex=False).str.replace(',', '.', regex=False)
            s = s_eu.where(has_comma, s)
        df[col] = pd.to_numeric(s, errors="coerce")

    # =========================================================
    # FASE 5: PAYMENT RECONCILIATION
    # =========================================================
    df = _apply_wahana_fac_accumulation(df, raw_results, osbal_rows, facode_osbal_idx)
    # =========================================================
    # FASE 6: RESOLUSI AKHIR & LABEL
    # =========================================================
    df = _compute_derived_cols(df)
    df.loc[df["SKENARIO"].isin(["UNMATCHED", "Unmatching"]), "FLAG_PROD"] = "Unmatching"
    df.loc[df["SKENARIO"] == "UNMATCHED", "SKENARIO"] = "Unmatching"

    # Hanya untuk OUTPUT: kosongkan tiga kolom nominal pada flag non-financial.
    # Diterapkan sesudah FLAG_PROD final agar tidak mengubah perhitungan/flagging.
    blank_flags = {
        "unmatching", "beda periode", "beda currency", "matching >1 fac code"
    }
    blank_mask = df["FLAG_PROD"].astype("string").str.strip().str.casefold().isin(blank_flags)
    df.loc[blank_mask, ["CCOS_OR_BAL", "CCOS_BAL_DUE", "DIFERENCE"]] = np.nan

    # Pastikan urutan 38 kolom persis
    for col in FINAL_COLUMNS:
        if col not in df.columns:
            df[col] = ""
    df = df[FINAL_COLUMNS]

    # Simpan ke Excel
    print(f"  Saving Excel to: {OUTPUT_FILE} ...", flush=True)
    for c in df.select_dtypes(include=['object', 'string']).columns:
        df[c] = df[c].astype(str).str.replace(r'[\x00-\x08\x0B\x0C\x0E-\x1F]', '', regex=True)
    df.to_excel(OUTPUT_FILE, index=False)

    try:
        from excel_styler import apply_purple_column_style
        apply_purple_column_style(OUTPUT_FILE, "matching")
    except Exception as e:
        print(f"  [WARN] Styling gagal: {e}")


    elapsed = time.perf_counter() - t_start
    print(f"\n{'=' * 60}")
    print(f"  Selesai Suspend Matching V1 dalam {elapsed:.1f}s ({elapsed/60:.1f} menit)")
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