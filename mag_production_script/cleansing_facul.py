import os
import re
import numpy as np
import pandas as pd

# ============================================================
# CONFIG
# ============================================================
INPUT_FILE = os.path.join("input", "1b. Transaksi Facul 01.01.23 - 17.08.2026.xlsx")
OUTPUT_FILE = os.path.join("output", "mag_output_facul.xlsx")

CEDANT_COL = "COMP_NAME"
CEDANT_VALUE = "PT ASURANSI MULTI ARTHA GUNA"   # setelah normalisasi (upper, buang '.', collapse spasi)
POLIS_COL = "FAC_POLICY_NO"
SLIP_COL = "FAC_SLIP"
INSURED_COL = "FAC_INSURED"
BROKER_NAME_COL = "COMP_NAME_1"
BROKER_CODE_COL = "FAC_BROKER"
MAX_SPLIT_COLS = 5

# NOTE:
# Tidak ada logic CLSDT_POLICY_NO / CLSDT_SLIP_NO / CLSDT_SERTF_NO / certificate
# di script ini. Ini murni Data 1 - FACUL.

POLIS_EXCEPTION_RE = re.compile(
    r"MOP\s*MARINE|(?:LINE\s*SLIP|LINESLIP)|\b(?:P1|P2|P3|P73)\s*CANCEL\b|\b(?:P1|P2|P3|P73)\b|\bCANCEL\b|PENYELESAIAN(?:\s+SUSPENSE)?|HUTANG|UTANG",
    re.I,
)

SLIP_EXCEPTION_RE = re.compile(
    r"\bSUMMARY\b|\bBORDER[OA]\b|\bBORDRO\b|\bSINGGLESHIPMENT\b|PENYELESAIAN(?:\s+SUSPENSE)?|HUTANG|UTANG",
    re.I,
)

# --------------------------------------------------------------
# POLA BARU KHUSUS MULTI ARTHA GUNA (ditemukan dari data asli):
# "<POLIS AKTIF> EX[.] [POLICY NO] [:] <POLIS LAMA/REFERENSI>"
# Contoh nyata:
#   45040118001473 Ex. Policy No : 45040117001948
#   36040119000076 EX. POLICY NO : 36040118000085
#   45090119000635 Ex 45090118000303
#   45040518000083 Ex.45040517000081
# Polis aktif adalah angka PERTAMA (sebelum EX). Bagian setelah EX
# adalah referensi ke polis periode sebelumnya, bukan polis tambahan.
# --------------------------------------------------------------
EX_POLICY_RE = re.compile(
    r"^\s*(\d{8,})\s*EX\.?\s*(?:POLICY\s*NO\.?)?\s*:?\s*\d{4,}\s*$",
    re.I,
)

MONTH_NAMES = frozenset({
    "JANUARI", "FEBRUARI", "MARET", "APRIL", "MEI", "JUNI", "JULI", "AGUSTUS",
    "SEPTEMBER", "OKTOBER", "NOVEMBER", "DESEMBER", "JANUARY", "FEBRUARY", "MARCH",
    "APRIL", "MAY", "JUNE", "JULY", "AUGUST", "SEPTEMBER", "OCTOBER", "NOVEMBER", "DECEMBER"
})

INSURED_REMOVE_RE = re.compile(
    r"\b(PT|CV|TBK|PERSERO|LTD|PTE|INC|LLC|MR|MRS|MS|BAPAK|BPK|IBU|NY|DR|DRS|DRA|IR|H|HJ)\b",
    re.I,
)
INSURED_POLIS_RE = re.compile(r"(POLIS\s*NO\.? .*)|(POLICY\s*NO\.? .*)|(SLIP\s*NO\.? .*)", re.I | re.X)
INSURED_SPLIT_RE = re.compile(r"\s*,\s*|\s*/\s*|\s+QQ\s+|\s+AND/OR\s+|\s*\+\s*", re.I)
INSURED_JUNK_WORDS = {"", "AND", "OR", "THE", "OF", "AS"}

# Angka romawi (I..XX) yang muncul sebagai kata utuh pada nama insured
# (mis. "BUANA I & II", "TOWER III") -> diubah jadi angka biasa.
# Urutan dari terpanjang ke terpendek supaya alternation regex match
# yang benar (mis. "III" tidak keburu cocok sebagai "II"+"I").
_ROMAN_NUMERAL_MAP = {
    "XX": "20", "XIX": "19", "XVIII": "18", "XVII": "17", "XVI": "16",
    "XV": "15", "XIV": "14", "XIII": "13", "XII": "12", "XI": "11",
    "X": "10", "IX": "9", "VIII": "8", "VII": "7", "VI": "6",
    "V": "5", "IV": "4", "III": "3", "II": "2", "I": "1",
}
_ROMAN_NUMERAL_RE = re.compile(
    r"\b(" + "|".join(sorted(_ROMAN_NUMERAL_MAP, key=len, reverse=True)) + r")\b"
)


def _convert_roman_numerals(text):
    return _ROMAN_NUMERAL_RE.sub(lambda m: _ROMAN_NUMERAL_MAP[m.group(1)], text)


# Panjang digit "wajar" untuk nomor polis Multi Artha Guna, berdasarkan
# observasi data asli (umumnya 8-20 digit). Kandidat digit di atas ini
# dianggap hasil data korup/gabungan tanpa separator -> jangan diambil.
MAX_REASONABLE_POLIS_DIGITS = 20


def _normalize_spaces(text):
    return re.sub(r"\s{2,}", " ", str(text)).strip()


def _cap_or_join(items):
    items = list(dict.fromkeys([str(x) for x in items if str(x).strip()]))
    return [",".join(items)] if len(items) > MAX_SPLIT_COLS else items


def _cap_or_original(items, original):
    items = list(dict.fromkeys([str(x) for x in items if str(x).strip()]))
    return [original] if len(items) > MAX_SPLIT_COLS else items


# ============================================================
# INSURED  (reused dari script Etiqa, generic - tidak diubah)
# ============================================================
def _clean_insured_name(name):
    if pd.isna(name):
        return ""
    name = str(name).upper()
    name = re.sub(r"\s*/\s*BORD(?:ER|ERO)?\b.*$", "", name, flags=re.I)
    name = re.sub(r"POLIS\s*NO\.?\s*.*$|POLICY\s*NO\.?\s*.*$|SLIP\s*NO\.?\s*.*$", "", name, flags=re.I)
    name = INSURED_REMOVE_RE.sub(" ", name)
    name = re.sub(r"[()]", " ", name)
    name = name.replace("/", " ").replace("-", " ")
    name = _convert_roman_numerals(name)
    return _normalize_spaces(name)


def split_insured(name):
    if pd.isna(name):
        return []
    hasil = []
    for p in INSURED_SPLIT_RE.split(str(name)):
        p = _clean_insured_name(p)
        if p and p not in INSURED_JUNK_WORDS:
            hasil.append(p)
    return hasil


def clean_insured(val):
    if pd.isna(val):
        return []
    val = str(val).strip()
    if not val:
        return []
    upper = val.upper()
    if re.match(r"^BIO\s+FARMA\b", upper):
        return [_normalize_spaces(re.sub(r"\s*[/&]\s*", " ", upper))]
    cleaned = []
    for ins in split_insured(val):
        ins = _normalize_spaces(ins).upper()
        if len(ins) <= 2 or ins in INSURED_JUNK_WORDS:
            continue
        if re.fullmatch(
            r"BORD(?:ER|ERO)?\s+(?:JANUARI|FEBRUARI|MARET|APRIL|MEI|JUNI|JULI|AGUSTUS|SEPTEMBER|OKTOBER|NOVEMBER|DESEMBER|JANUARY|FEBRUARY|MARCH|MAY|JUNE|JULY|AUGUST|SEPTEMBER|OCTOBER|NOVEMBER|DECEMBER)(?:\s+20\d{2})?",
            ins,
            re.I,
        ):
            continue
        if ins not in cleaned:
            cleaned.append(ins)
    if not cleaned:
        # Jangan sampai hasil akhir kosong padahal input tidak kosong
        # (mis. "BAPAK/IBU" -> semua kata masuk daftar gelar yang dihapus).
        fallback = _clean_insured_name(val)
        if not fallback:
            fallback = _normalize_spaces(val)
        cleaned = [fallback.upper()]
    return _cap_or_join(cleaned)


# ============================================================
# CERTIFICATE (reuse PERSIS dari script Etiqa, tidak diubah)
# ============================================================
def _normalize_certificate_number(value):
    """
    Normalisasi certificate menjadi 6 digit.
    Contoh: 176 -> 000176, 000176 -> 000176, 0000131 -> 000131
    """
    value = str(value).strip()
    if not value.isdigit():
        return None
    if len(value) > 6:
        value = value[-6:]
    if int(value) == 0:
        return None
    return value.zfill(6)


def _format_certificate_list(values):
    """
    <= 3 certificate -> "000001, 000002, 000003"
    >  3 certificate -> "000001 SD 000005"
    """
    cleaned = []
    for value in values:
        value = str(value).strip()
        if not value:
            continue
        if not re.fullmatch(r"\d{1,6}", value):
            continue
        if int(value) == 0:
            continue
        value = value.zfill(6)
        if value not in cleaned:
            cleaned.append(value)
    if not cleaned:
        return ""
    if len(cleaned) <= 3:
        return ", ".join(cleaned)
    return f"{cleaned[0]} SD {cleaned[-1]}"


def _expand_certificate_range(start, end):
    if not (re.fullmatch(r"\d{1,6}", start) and re.fullmatch(r"\d{1,6}", end)):
        return ""
    a = int(start)
    b = int(end)
    if a == 0 or b == 0:
        return ""
    count = abs(b - a) + 1
    if count > 3:
        return f"{start.zfill(6)} SD {end.zfill(6)}"
    step = 1 if a <= b else -1
    values = []
    for number in range(a, b + step, step):
        if number == 0:
            continue
        values.append(str(number).zfill(6))
    return ", ".join(values)


def clean_certificate(polis_ori):
    """
    Ekstrak certificate dari FAC_POLICY_NO, PERSIS logic Etiqa.
    Hanya aktif untuk pola "POLICY - CERTIFICATE[...]" murni
    (fullmatch), jadi tidak akan aktif pada chain multi-polis
    (comma-chain/plus-chain) yang sudah ditangani clean_polis.
    """
    if pd.isna(polis_ori):
        return []

    raw = _normalize_spaces(str(polis_ori).upper())
    if not raw:
        return []

    # 1. POLICY - CERTIFICATE S/D CERTIFICATE
    m = re.fullmatch(r"\s*(\d{8,})\s*-\s*(\d+)\s*S\s*/?\s*D\s*(\d+)\s*", raw, re.I)
    if m:
        cert_start = _normalize_certificate_number(m.group(2))
        cert_end = _normalize_certificate_number(m.group(3))
        if cert_start and cert_end:
            return [f"{cert_start} SD {cert_end}"]
        return []

    # 2. POLICY - CERTIFICATE - CERTIFICATE
    m = re.fullmatch(r"\s*(\d{8,})\s*-\s*(\d+)\s*-\s*(\d+)\s*", raw, re.I)
    if m:
        c2, c3 = m.group(2), m.group(3)
        # MAG-specific: certificate asli PERSIS 6 digit dari sananya.
        # Kalau kurang dari 6 -> kemungkinan perulangan/lanjutan polis.
        # Kalau lebih dari 6 -> kemungkinan itu polis penuh lain, BUKAN
        # certificate (jangan dipotong 6 digit terakhir, itu menebak).
        if len(c2) != 6 or len(c3) != 6:
            return []
        certificates = []
        for value in (c2, c3):
            cert = _normalize_certificate_number(value)
            if cert:
                certificates.append(cert)
        return certificates[:3]

    # 3. POLICY - CERTIFICATE
    m = re.fullmatch(r"\s*(\d{8,})\s*-\s*(\d+)\s*", raw, re.I)
    if m:
        suffix = m.group(2)
        # Sama seperti RULE 2: certificate asli harus PERSIS 6 digit.
        if len(suffix) != 6:
            return []
        cert = _normalize_certificate_number(suffix)
        if cert:
            return [cert]
        return []

    # 4. POLICY - CERTIFICATE/CERTIFICATE/CERTIFICATE
    m = re.fullmatch(r"\s*(\d{8,})\s*-\s*(.+?)\s*", raw, re.I)
    if m:
        suffix = m.group(2).strip()
        suffix = re.sub(r"\b(?:VARIOUS|VAR|TBA)\b", "", suffix, flags=re.I).strip(" ,/-")
        parts = re.split(r"\s*[/,]\s*", suffix)
        certificates = []
        for part in parts:
            part = part.strip()
            if not part:
                continue
            range_match = re.fullmatch(r"(\d+)\s*-\s*(\d+)", part)
            dummy_prefix_match = re.fullmatch(r"0+\s*-\s*(\d+)", part)
            if dummy_prefix_match:
                digits = dummy_prefix_match.group(1)
                # >=8 digit berarti itu polis penuh lain, BUKAN certificate.
                if len(digits) >= 8:
                    certificates = []
                    break
                cert = _normalize_certificate_number(digits)
                if cert:
                    certificates.append(cert)
                continue
            if range_match:
                start_d, end_d = range_match.group(1), range_match.group(2)
                if len(start_d) >= 8 or len(end_d) >= 8:
                    certificates = []
                    break
                start = _normalize_certificate_number(start_d)
                end = _normalize_certificate_number(end_d)
                if start and end:
                    certificates.extend([start, end])
                continue
            # Bagian polos (tanpa dash internal): kalau >=8 digit, itu
            # polis penuh lain (BUKAN certificate) -> batalkan seluruh
            # RULE 4 untuk value ini (ambigu, jangan ditebak sebagian).
            if len(part) >= 8 and part.isdigit():
                certificates = []
                break
            cert = _normalize_certificate_number(part)
            if cert:
                certificates.append(cert)

        unique_certificates = []
        for cert in certificates:
            if cert not in unique_certificates:
                unique_certificates.append(cert)
        certificates = unique_certificates

        if not certificates:
            return []
        if len(certificates) > 3:
            return [f"{certificates[0]} SD {certificates[-1]}"]
        return certificates[:3]

    # 5. FORMAT LAIN / RANDOM
    return []


# ============================================================
# RULE BARU (TAMBAHAN): PEMISAHAN POLIS & CERTIFICATE BERDASARKAN
# NOMOR POLIS DASAR YANG SAMA.
#
# INI TIDAK MENGUBAH clean_polis / clean_certificate DI ATAS SAMA
# SEKALI. Ini adalah LAYER TAMBAHAN yang dicoba LEBIH DULU; kalau
# tidak cocok/tidak yakin, otomatis fallback ke clean_polis +
# clean_certificate yang sudah ada (lihat get_policy_certificate_pairs
# di bawah).
#
# Prinsip: pecah value berdasarkan separator +, /, , (level atas).
# Tiap bagian ("chunk") harus jelas berupa salah satu dari:
#   - BASE-SUFFIX (mis. "36040118000154-000001")
#   - BASE-SUFFIX S/D SUFFIX (mis. "...-000001 S/D 000006")
#   - BASE- (dash tanpa suffix, mis. "08080517000116-")
#   - BASE saja (tanpa dash)
#   - angka pendek <8 digit tanpa base sendiri (mis. "00002") -> ini
#     dianggap suffix/certificate lanjutan dari BASE TERAKHIR yang
#     baru saja dikenali (elision).
# Kalau ADA SATU SAJA chunk yang tidak cocok pola di atas -> seluruh
# value dianggap TIDAK YAKIN -> return None (fallback ke rule lama).
#
# Chunk dengan BASE yang SAMA (persis sama string-nya) digabung jadi
# SATU polis, certificate-nya dijadikan list (duplicate dibuang).
# ============================================================
_CHUNK_SD_RE = re.compile(r"^(\d{8,})\s*-\s*(\d+)\s*SDPLACEHOLDER\s*(\d+)$")
_CHUNK_BASE_SUFFIX_RE = re.compile(r"^(\d{8,})\s*-\s*(\d+)$")
_CHUNK_BASE_DASH_EMPTY_RE = re.compile(r"^(\d{8,})\s*-\s*$")
_CHUNK_BASE_ONLY_RE = re.compile(r"^(\d{8,})$")
_CHUNK_SHORT_SUFFIX_RE = re.compile(r"^(\d{1,7})$")


def _try_new_certificate_grouping(original):
    """
    Return list of (polis, certificate_or_None) kalau berhasil
    diklasifikasi dengan yakin, atau None kalau tidak (-> fallback
    ke clean_polis + clean_certificate seperti biasa, TIDAK berubah).

    Aturan tambahan (dari instruksi terbaru):
      - Certificate WAJIB diawali angka "0" (kalau tidak, bukan
        certificate -> batalkan seluruh grouping, ambigu).
      - Angka pendek yang menempel ke BASE TANPA DASH (mis.
        "45013018008305+8007321", base-nya bare tanpa "-") BUKAN
        certificate: kalau cuma 1 -> serahkan ke rule lama (yang akan
        merekonstruksinya sebagai "perulangan"/lanjutan polis);
        kalau 2 atau lebih -> pola dianggap terlalu ambigu untuk
        direkonstruksi rule lama -> paksa keep original.
    """
    # Buang descriptor VARIOUS/VAR/TBA (mengikuti konvensi clean_polis)
    # SEBELUM cek huruf, supaya "/VARIOUS" di akhir tidak menggagalkan
    # grouping base-yang-sama.
    without_descriptor = re.sub(
        r"[+/,]?\s*\b(?:VARIOUS|VAR|TBA)\b", "", original, flags=re.I
    ).strip(" +/,")

    # Tolak dulu kalau ada huruf selain yang membentuk pola "S/D"
    # (supaya CANCEL/PENYELESAIAN/P1/P2/P3/dll tetap ditangani rule
    # lama, tidak diganggu rule baru ini).
    check_str = re.sub(r"S\s*/\s*D", "", without_descriptor, flags=re.I)
    if re.search(r"[A-Z]", check_str):
        return None
    if "&" in without_descriptor:
        return None

    protected = re.sub(r"S\s*/\s*D", "SDPLACEHOLDER", without_descriptor, flags=re.I)
    chunks = [c.strip() for c in re.split(r"\s*[+/,]\s*", protected) if c.strip()]
    if not chunks:
        return None

    # Kalau cuma ADA SATU chunk (tidak ada +/,/ sama sekali di value),
    # rule baru ini HANYA boleh aktif untuk pola S/D yang memang jelas
    # (mis. "POLIS-000001 S/D 000006"). Untuk single "BASE-SUFFIX"
    # polos TANPA pengulangan base (tidak ada bukti itu certificate,
    # bisa juga cuma perulangan/lanjutan nomor polis), JANGAN diambil
    # alih di sini - biarkan rule LAMA yang menentukan (certificate
    # cuma diakui kalau suffix-nya PERSIS 6 digit, sesuai instruksi
    # sebelumnya). Ini mencegah rule baru menabrak fix yang sudah ada.
    if len(chunks) == 1 and not _CHUNK_SD_RE.match(chunks[0]):
        return None

    groups = []          # list of dict {base, certs: [..]}
    base_to_group = {}   # base -> group dict (untuk lookup cepat)
    current_base = None
    current_dash_established = False
    bare_elision_count = 0

    for chunk in chunks:
        m = _CHUNK_SD_RE.match(chunk)
        if m:
            base, s1, s2 = m.group(1), m.group(2), m.group(3)
            if not (s1.startswith("0") and s2.startswith("0")):
                return None
            suffix = f"{s1} S/D {s2}"
            dash_established = True
        else:
            m = _CHUNK_BASE_SUFFIX_RE.match(chunk)
            if m:
                base, suffix = m.group(1), m.group(2)
                if not suffix.startswith("0"):
                    return None
                dash_established = True
            else:
                m = _CHUNK_BASE_DASH_EMPTY_RE.match(chunk)
                if m:
                    base, suffix = m.group(1), None
                    dash_established = True
                else:
                    m = _CHUNK_BASE_ONLY_RE.match(chunk)
                    if m:
                        base, suffix = m.group(1), None
                        dash_established = False
                    else:
                        m = _CHUNK_SHORT_SUFFIX_RE.match(chunk)
                        if m:
                            if current_base is None:
                                return None
                            if not current_dash_established:
                                # Base terakhir BARE (tanpa dash) -> angka
                                # pendek ini BUKAN certificate. Jangan
                                # putuskan sekarang - hitung dulu semua,
                                # baru diputuskan setelah loop selesai
                                # (1 elision -> decline; 2+ -> keep original).
                                bare_elision_count += 1
                                continue
                            base, suffix = current_base, m.group(1)
                            if not suffix.startswith("0"):
                                return None
                            dash_established = True
                        else:
                            # Pola tidak dikenali sama sekali -> ambigu,
                            # jangan tebak, batalkan seluruh grouping.
                            return None

        current_base = base
        current_dash_established = dash_established
        grp = base_to_group.get(base)
        if grp is None:
            grp = {"base": base, "certs": []}
            base_to_group[base] = grp
            groups.append(grp)
        if suffix is not None and suffix not in grp["certs"]:
            grp["certs"].append(suffix)

    if bare_elision_count == 1:
        return None
    if bare_elision_count >= 2:
        return [(original, None)]

    if not groups or len(groups) > MAX_SPLIT_COLS:
        return None

    return [(g["base"], ",".join(g["certs"]) if g["certs"] else None) for g in groups]


def get_policy_certificate_pairs(raw_polis):
    """
    Fungsi utama BARU yang dipakai process_data untuk menghasilkan
    pasangan (clean polis, certificate). Prioritas:
      1. Rule BARU (_try_new_certificate_grouping) - hanya aktif untuk
         pola "beberapa chunk BASE(-SUFFIX)? yang digabung +/,/ /".
      2. Kalau rule baru tidak yakin -> fallback PERSIS ke clean_polis()
         + clean_certificate() yang SUDAH ADA (tidak diubah sama sekali).
    """
    if pd.isna(raw_polis):
        return []
    original = _normalize_spaces(str(raw_polis).upper())
    if not original:
        return []

    new_result = _try_new_certificate_grouping(original)
    if new_result is not None:
        return new_result

    # --- fallback: rule lama, TIDAK DIUBAH ---
    c_polis = clean_polis(raw_polis)
    if len(c_polis) == 1:
        raw_str = original
        if "," not in raw_str and "&" not in raw_str and "+" not in raw_str:
            c_cert = clean_certificate(raw_polis)
            if c_cert:
                return [(c_polis[0], ",".join(c_cert))]
        # Safety-net BARU (tidak mengubah clean_polis itu sendiri):
        # kalau clean_polis mengembalikan 1 token yang ternyata cuma
        # SEBAGIAN dari original, dan sisanya murni angka/spasi/dash
        # (bukan descriptor huruf seperti VAR/TBA/EX POLICY NO yang
        # memang sengaja dibuang rule lama) -> berarti ada digit asli
        # yang diam-diam hilang -> JANGAN dipercaya, kembalikan
        # original utuh (tidak menebak), sesuai instruksi terbaru.
        if c_polis[0] != original:
            leftover = original.replace(c_polis[0], "", 1)
            if re.fullmatch(r"[\d\s\-,/]+", leftover) and re.search(r"\d", leftover):
                return [(original, None)]
    return [(p, None) for p in c_polis]


# ============================================================
# POLIS  (adaptasi FACUL - Multi Artha Guna, TANPA certificate)
# ============================================================
def _extract_main_policy(val):
    m = re.search(r"\b\d{8,}\b", val)
    return m.group(0) if m else ""


_TOKEN_FULL_DASH_RE = re.compile(r"^(\d{8,})\s*-\s*(\d{1,14})$")


def _classify_policy_token(tok):
    """
    Klasifikasi satu token polis (hasil split dari '-', '+', atau ',').
      - "full_dash": token sendiri sudah lengkap berupa BASE-SUFFIX
        (dipertahankan literal apa adanya, TIDAK direkonstruksi ulang).
      - "full": token berupa angka polos >=8 digit, berdiri sendiri.
      - "short": angka pendek (<=7 digit) yang mengacu ke base sebelumnya.
      - None: tidak bisa diklasifikasi dengan aman -> ambigu.
    """
    tok = tok.strip()
    m = _TOKEN_FULL_DASH_RE.match(tok)
    if m:
        return "full_dash", tok.replace(" ", ""), m.group(1)
    if re.fullmatch(r"\d{8,}", tok):
        return "full", tok, tok
    if re.fullmatch(r"\d{1,7}", tok):
        return "short", tok, None
    return None, None, None


def _build_policy_chain(tokens_raw, original, drop_words=frozenset()):
    """
    Fungsi rekonstruksi tunggal, dipakai bersama untuk chain '-', '+', dan ','
    supaya perlakuan konsisten (tidak ada logic ganda/beda antar separator).

    Aturan:
      - Token "full" / "full_dash" -> dipertahankan apa adanya (tidak diotak-atik).
        Setelah token "full" (angka polos, TANPA dash internal), base
        angka tsb boleh dipakai untuk merekonstruksi token "short" berikutnya.
      - Setelah token "full_dash" (sudah mengandung dash sendiri), base
        DIRESET (base=None). Ini supaya token "short" berikutnya yang
        mengacu ke base yang sudah punya dash internal TIDAK direkonstruksi
        secara otomatis (pola ambigu -> lebih aman keep original) daripada
        menebak makna suffix tersebut.
      - Kalau ada token yang tidak bisa diklasifikasi (huruf, S/D, dsb),
        atau token "short" muncul tanpa base yang valid -> ambigu -> keep original.
    """
    tokens_out = []
    base = None
    for raw in tokens_raw:
        raw_s = raw.strip()
        if not raw_s:
            continue
        if raw_s.upper() in drop_words or re.fullmatch(r"P[123]", raw_s, re.I):
            continue
        kind, value, base_digits = _classify_policy_token(raw_s)
        if kind is None:
            return [original]
        if kind == "full":
            tokens_out.append(value)
            base = base_digits
        elif kind == "full_dash":
            tokens_out.append(value)
            base = None  # ambigu untuk rekonstruksi short berikutnya
        else:  # short
            if base is None:
                return [original]
            suf = value
            full_val = base[:len(base) - len(suf)] + suf if len(suf) < len(base) else suf
            tokens_out.append(full_val)
            base = full_val

    tokens_out = list(dict.fromkeys(tokens_out))
    if not tokens_out:
        return [original]
    return _cap_or_original(tokens_out, original)


def clean_polis(val):
    if pd.isna(val):
        return []
    original = _normalize_spaces(str(val).upper())
    if not original:
        return []

    # --------------------------------------------------------
    # RULE BARU KHUSUS MAG: "POLIS EX POLIS_LAMA"
    # Ambil polis aktif (bagian pertama) saja.
    # --------------------------------------------------------
    m = EX_POLICY_RE.match(original)
    if m:
        return [m.group(1)]

    # Exception wajib dipertahankan (PENYELESAIAN/HUTANG/CANCEL/P1-P3/dll).
    if POLIS_EXCEPTION_RE.search(original):
        m = re.search(r"\bP[123P]?\s*/\s*(\d{8,})", original, re.I)
        if m:
            return [m.group(1)]
        if re.fullmatch(r"P[123](?:\s*\+\s*P[123])?", original, re.I):
            return [original]
        if re.fullmatch(r"TBA\s*/.*", original, re.I):
            return [original]
        if re.fullmatch(r"(?:P1|P2|P3|P73)", original, re.I):
            return [original]

    if re.fullmatch(r"VARIOUS", original, re.I):
        return ["VARIOUS"]

    # --------------------------------------------------------
    # PRIORITASKAN pola "POLICY - CERTIFICATE[s]" ala Etiqa untuk
    # kasus SEDERHANA (tanpa koma/&/+ tercampur, supaya tidak bentrok
    # dengan comma-chain/plus-chain multi-polis yang sudah tervalidasi
    # di atas). Kalau cocok, clean_polis hanya mengembalikan BASE
    # policy-nya saja (satu token) - certificate-nya diekstrak terpisah
    # lewat clean_certificate() di process_data.
    # --------------------------------------------------------
    if "," not in original and "&" not in original and "+" not in original:
        if clean_certificate(original):
            m_base = re.match(r"^\s*(\d{8,})", original)
            if m_base:
                return [m_base.group(1)]

    # Hapus descriptor VARIOUS/VAR/TBA setelah policy valid.
    val2 = re.sub(r"\s*/\s*(?:VAR|VARIOUS|TBA)\b.*$", "", original, flags=re.I)
    val2 = re.sub(r"\s+(?:VAR|VARIOUS)\b.*$", "", val2, flags=re.I)
    val2 = _normalize_spaces(val2)

    # P1 / policy atau P2 / policy.
    m = re.fullmatch(r"P[123]\s*/\s*(\d{8,})(?:\s+.*)?", val2, re.I)
    if m:
        return [m.group(1)]

    # Catatan: shortcut "ambil main policy saja jika ada S/D" HANYA berlaku
    # kalau tidak ada koma. Kalau ada koma + S/D tercampur (pola langka,
    # <1% data), pola dianggap tidak cukup jelas -> ditangani di
    # comma-chain di bawah (yang akan fallback ke original bila ambigu).
    main = _extract_main_policy(val2)
    if main and "," not in val2 and re.search(r"\bS\s*\.?\s*/?\s*D\b", val2, re.I):
        return [main]

    # --------------------------------------------------------
    # COMMA CHAIN - beberapa polis dipisah ',' (BUKAN certificate).
    # Contoh: 01031118000012-000006,01031118000012-000007
    #         05030319049722,05030319049733,05030319049744,49755
    # Cek DULU sebelum dash-chain, karena comma adalah separator utama.
    # --------------------------------------------------------
    if "," in val2:
        parts = [p.strip() for p in val2.split(",") if p.strip()]
        if len(parts) >= 2:
            return _build_policy_chain(parts, original)

    # --------------------------------------------------------
    # DASH CHAIN - perulangan polis (BUKAN certificate).
    # Menangani baik suffix pendek maupun beberapa polis penuh sekaligus.
    # Contoh: 13010921002655-13010921002644-13010921002677-2666 (mixed)
    #         00940502012023001276-136-135-1275 (short suffix)
    # --------------------------------------------------------
    if re.fullmatch(r"\d{8,}(?:\s*-\s*\d+)+", val2):
        parts = [p for p in re.split(r"\s*-\s*", val2) if p]
        return _build_policy_chain(parts, original)

    # Numeric S/D range tanpa policy prefix: expand jika <=5, selain itu pertahankan.
    m = re.fullmatch(r"(\d+)\s*S\s*/?\s*D\s*(\d+)", val2, re.I)
    if m and len(m.group(1)) == len(m.group(2)):
        a, b = int(m.group(1)), int(m.group(2))
        count = abs(b - a) + 1
        if count > MAX_SPLIT_COLS:
            return [original]
        step = 1 if a <= b else -1
        return [str(n).zfill(len(m.group(1))) for n in range(a, b + step, step)]

    # --------------------------------------------------------
    # AMPERSAND - diperketat untuk MAG.
    # Hanya displit jika SETIAP bagian setelah dibersihkan
    # sudah terlihat seperti polis penuh (>=8 digit alfanumerik).
    # Jika ada bagian yang pendek/ambigu (mis. sisa certificate-like
    # suffix: "45031125000863 - 000001 & 000002"), JANGAN displit,
    # pertahankan original.
    # --------------------------------------------------------
    if "&" in val2:
        raw_parts = [p.strip() for p in re.split(r"\s*&\s*", val2)]
        cleaned_parts = [re.sub(r"[^A-Z0-9]", "", p) for p in raw_parts]
        cleaned_parts = [p for p in cleaned_parts if p]
        if cleaned_parts and all(len(p) >= 8 for p in cleaned_parts):
            return _cap_or_original(cleaned_parts, original)
        return [original]

    # Plus = separate policy. Suffix pendek dibentuk dari policy sebelumnya.
    if "+" in val2:
        parts = [p.strip() for p in re.split(r"\s*\+\s*", val2) if p.strip()]
        return _build_policy_chain(parts, original, drop_words={"VAR", "VARIOUS", "TBA"})

    # Slash yang hanya memisahkan policy valid.
    if "/" in val2:
        result = _build_policy_chain(
            re.split(r"\s*/\s*", val2), original, drop_words={"VAR", "VARIOUS", "TBA"}
        )
        if result != [original]:
            return result

    # Policy numerik murni / dotted.
    if re.fullmatch(r"[0-9.]+", val2):
        return [val2.replace(".", "")]

    # Dash yang bukan certificate: fallback bersih tanpa separator.
    # Catatan: hanya aman kalau ada SATU bagian yang jelas berupa
    # polis penuh (>=8 digit). Kalau semua bagian pendek (mis.
    # "2485-2474-2452-2463", tidak ada basis polis yang jelas),
    # JANGAN digabung paksa jadi satu angka -> keep original.
    if re.fullmatch(r"[0-9-]+", val2):
        parts = [p for p in val2.split("-") if p]
        if len(parts) == 1:
            return [parts[0]]
        if len(parts[0]) >= 8:
            return [parts[0]]
        return [original]

    # --------------------------------------------------------
    # Fallback umum: ambil semua nomor panjang (>=8 digit).
    # Tambahan safety-cap khusus MAG: kandidat yang terlalu panjang
    # (> MAX_REASONABLE_POLIS_DIGITS) dianggap data korup/gabungan
    # tanpa separator -> jangan diambil, keep original.
    # --------------------------------------------------------
    candidates = re.findall(r"\b\d{8,}\b", val2)
    candidates = [c for c in candidates if len(c) <= MAX_REASONABLE_POLIS_DIGITS]
    if candidates:
        return _cap_or_original(list(dict.fromkeys(candidates)), original)

    return [val2] if val2 else []


# ============================================================
# SLIP  (reused dari script Etiqa, generic - tidak diubah)
# ============================================================
def clean_slip(val):
    if pd.isna(val):
        return []
    original = _normalize_spaces(str(val).upper())
    if not original:
        return []

    val2 = original.replace(".", "")

    # Exception: P1/P2/P3 murni harus dipertahankan (jangan sampai
    # tersapu oleh noise-removal generic di bagian bawah fungsi ini).
    if re.fullmatch(r"P[123]", val2, re.I):
        return [val2]

    # --------------------------------------------------------
    # Kombinasi placeholder layer P1/P2/P3/P4/dst (mis. "P1 + P2",
    # "P1 + P2 + P3 + P4", "P1 (P2)", "P1 / END").
    # Regex noise-removal generic di bawah bersifat destruktif untuk
    # pola ini (bisa menghapus seluruh string). Tangani lebih awal:
    #   - kalau ada nomor slip asli (>=7 digit) di dalamnya, ambil nomor
    #     itu saja (placeholder P-nya dibuang, bukan bagian dari slip).
    #   - kalau murni kombinasi P1/P2/P3/dst tanpa nomor asli,
    #     pertahankan original (jangan sampai jadi kosong).
    # --------------------------------------------------------
    if re.search(r"\bP\s*[1-9]\b", val2, re.I):
        numbers = re.findall(r"\b\d{7,}\b", val2)
        if numbers:
            numbers = list(dict.fromkeys(numbers))
            return numbers if len(numbers) <= MAX_SPLIT_COLS else [val2]
        return [val2]

    if SLIP_EXCEPTION_RE.search(val2):
        return [_normalize_spaces(val2)]

    m = re.fullmatch(r"([A-Z]+)\s+(20\d{2})\s*/\s*VARIOUS", val2, re.I)
    if m and m.group(1) in MONTH_NAMES:
        return [f"{m.group(1)} {m.group(2)}"]

    m = re.fullmatch(r"P[123]\s*/\s*(\d{7,})", val2, re.I)
    if m:
        return [m.group(1)]

    m = re.fullmatch(r"(\d{7,})\s*/\s*(\d{7,})", val2)
    if m:
        return [m.group(1), m.group(2)]

    m = re.fullmatch(r"(\d{7,})\s*/\s*0", val2)
    if m:
        return [m.group(1)]

    m = re.match(r"^(\d{7,})\s*-\s*\d+/CN/", val2, re.I)
    if m:
        return [m.group(1)]

    m = re.search(r"(\d{7,})\s*S\s*\.?\s*/?\s*D\s*(\d{7,})", val2, re.I)
    if m and len(m.group(1)) == len(m.group(2)):
        a, b = int(m.group(1)), int(m.group(2))
        count = abs(b - a) + 1
        if count > MAX_SPLIT_COLS:
            return [f"{m.group(1)} S/D {m.group(2)}"]
        step = 1 if a <= b else -1
        return [str(n).zfill(len(m.group(1))) for n in range(a, b + step, step)]

    m = re.fullmatch(r"(\d{7,})\s*/\s*(\d{1,6})", val2)
    if m:
        base, suffix = m.groups()
        if len(suffix) < len(base):
            return [base, base[:len(base) - len(suffix)] + suffix]

    if "+" in val2:
        parts = [p.strip() for p in val2.split("+") if p.strip()]
        if parts and re.fullmatch(r"\d{7,}", parts[0]):
            base = parts[0]
            hasil = [base]
            for suffix in parts[1:]:
                if re.fullmatch(r"\d{3,6}", suffix):
                    full = base[:len(base) - len(suffix)] + suffix if len(suffix) < len(base) else suffix
                    if full not in hasil:
                        hasil.append(full)
                elif suffix.upper() not in {"VAR", "VARIOUS", "TBA", "P1", "P2", "P3"}:
                    if len(suffix) >= 7 and suffix not in hasil:
                        hasil.append(suffix)
            return _cap_or_join(hasil)

    if re.fullmatch(r"\d{7,}(?:\s*-\s*\d{1,6})+", val2):
        parts = [p.strip() for p in re.split(r"\s*-\s*", val2)]
        base = parts[0]
        hasil = [base]
        for suffix in parts[1:]:
            if re.fullmatch(r"\d{1,6}", suffix):
                full = base[:len(base) - len(suffix)] + suffix if len(suffix) < len(base) else suffix
                if full not in hasil:
                    hasil.append(full)
        return _cap_or_join(hasil)

    val3 = re.sub(r"^SEE\s+ATTACHMENT\s*/\s*", "", val2, flags=re.I).strip()
    val3 = re.sub(r"\bEND(?:\.?1)?\b", "", val3, flags=re.I)
    val3 = re.sub(r"\b(?:USD|IDR|SGD|EUR|JPY|AUD|GBP|ORI|ORIGINAL|COPY|REALISASI|REALIZATION|CANCEL(?:LED)?|ENDORSEMENT|SA|P1|P2|P3|VAR|REVISI|REV)\b.*$", "", val3, flags=re.I)
    val3 = _normalize_spaces(val3).strip("-_,.; ")

    candidates = re.findall(r"\b\d{7,}\b", val3)
    if candidates:
        candidates = list(dict.fromkeys(candidates))
        return candidates if len(candidates) <= MAX_SPLIT_COLS else [val3]

    if "/CN/" in val3 or re.search(r"[A-Z]{2,}/\d", val3):
        return [val3]

    if "," in val3:
        parts = [p.strip() for p in val3.split(",") if p.strip()]
        return _cap_or_original(parts, original)

    if re.fullmatch(r"\d{1,6}(?:\s*[&+-]\s*\d{1,6})+", val3):
        return [val3]

    return [val3] if val3 else []


# ============================================================
# BUSINESS PARTNERS  (reused dari script Etiqa)
# ============================================================
def clean_business_partners(val):
    if pd.isna(val):
        return ""
    val = str(val).strip().upper()
    val = re.sub(r"\bPT\.\s*", "PT ", val, flags=re.I)
    return _normalize_spaces(val)


def get_mitra_bisnis(broker_name, broker_code, cedant):
    broker_name = "" if pd.isna(broker_name) else str(broker_name).strip()
    broker_code = "" if pd.isna(broker_code) else str(broker_code).strip()
    cedant = "" if pd.isna(cedant) else str(cedant).strip()
    if broker_name and broker_name.upper() != "DIRECT":
        return clean_business_partners(broker_name)
    if broker_code and broker_code.upper() != "DIRECT":
        return clean_business_partners(broker_code)
    return clean_business_partners(cedant)


def get_mitra_bisnis(broker_name, broker_code, cedant):
    broker_name = "" if pd.isna(broker_name) else str(broker_name).strip()
    broker_code = "" if pd.isna(broker_code) else str(broker_code).strip()
    cedant = "" if pd.isna(cedant) else str(cedant).strip()
    if broker_name and broker_name.upper() != "DIRECT":
        return clean_business_partners(broker_name)
    if broker_code and broker_code.upper() != "DIRECT":
        return clean_business_partners(broker_code)
    return clean_business_partners(cedant)


def _insert_clean_columns(df, all_lists, prefix, max_cols):
    added = []
    for i in range(1, max_cols + 1):
        col_name = f"clean {prefix} {i}"
        df[col_name] = [lst[i - 1] if i - 1 < len(lst) else None for lst in all_lists]
        added.append(col_name)
    return added


def _fast_read_excel(path, sheet_name=None, header=0):
    import openpyxl
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb[sheet_name] if sheet_name else wb.worksheets[0]
    rows = ws.iter_rows(values_only=True)
    for _ in range(header):
        next(rows)
    cols = list(next(rows))
    seen = {}
    new_cols = []
    for col in cols:
        col = "" if col is None else str(col).strip()
        if col not in seen:
            seen[col] = 0
            new_cols.append(col)
        else:
            seen[col] += 1
            new_cols.append(f"{col}_{seen[col]}")
    data = list(rows)
    wb.close()
    return pd.DataFrame(data, columns=new_cols)


def _fast_write_excel(df, path):
    import openpyxl
    from openpyxl.cell import WriteOnlyCell
    output_dir = os.path.dirname(path)
    if output_dir:
        os.makedirs(output_dir, exist_ok=True)
    wb = openpyxl.Workbook(write_only=True)
    ws = wb.create_sheet("Sheet1")
    ws.append(list(df.columns))
    for row in df.itertuples(index=False, name=None):
        out = []
        for value in row:
            cell = WriteOnlyCell(ws, value=None if pd.isna(value) else value)
            if isinstance(value, str):
                cell.number_format = "@"
            out.append(cell)
        ws.append(out)
    wb.save(path)


# ============================================================
# PROCESS DATA  (Data 1 - FACUL, tanpa kolom Certificate)
# ============================================================
def process_data(input_file, output_file):
    print(f"[1/5] Membaca data dari: {input_file} ...")
    df = _fast_read_excel(input_file, header=0)
    total_input_rows = len(df)
    print(f"      Total baris keseluruhan (semua cedant): {total_input_rows:,}")

    for col in (CEDANT_COL, POLIS_COL, SLIP_COL, INSURED_COL, BROKER_NAME_COL, BROKER_CODE_COL):
        if col not in df.columns:
            print(f"[ERROR] Kolom '{col}' tidak ditemukan di file!")
            return

    cedant_norm = (
        df[CEDANT_COL].fillna("").astype(str).str.upper().str.strip()
        .str.replace(r"\.", " ", regex=True).str.replace(r"\s+", " ", regex=True).str.strip()
    )
    cedant_target_norm = re.sub(r"\s+", " ", CEDANT_VALUE.upper().replace(".", " ")).strip()
    is_target = cedant_norm.eq(cedant_target_norm)
    total_target = int(is_target.sum())
    total_other = total_input_rows - total_target
    print(f"[2/5] Filter cedant '{CEDANT_VALUE}': {total_target:,} baris target dari total {total_input_rows:,} baris "
          f"({total_other:,} baris cedant lain, TIDAK diubah).")

    broker_name_s = df[BROKER_NAME_COL].fillna("").astype(str).str.strip()
    broker_code_s = df[BROKER_CODE_COL].fillna("").astype(str).str.strip()
    mitra_values = [
        get_mitra_bisnis(a, b, c) for a, b, c in zip(broker_name_s, broker_code_s, cedant_norm)
    ]

    df.rename(columns={POLIS_COL: "polis_ori", SLIP_COL: "slip_ori", INSURED_COL: "insured_ori"}, inplace=True)

    insert_pos = list(df.columns).index(BROKER_NAME_COL) + 1 if BROKER_NAME_COL in df.columns else len(df.columns)
    df.insert(insert_pos, "BUSINESS PARTNERS", mitra_values)

    print("[3/5] Menjalankan cleaning HANYA untuk baris target (baris lain dibiarkan apa adanya) ...")
    n = len(df)
    all_polis_cert_pairs = [[] for _ in range(n)]   # list of (polis, cert_or_None)
    all_clean_slip = [[] for _ in range(n)]
    all_clean_ins = [[] for _ in range(n)]
    max_polis = max_slip = max_ins = 1

    target_idx = np.flatnonzero(is_target.to_numpy())
    polis_vals = df["polis_ori"].to_numpy()
    slip_vals = df["slip_ori"].to_numpy()
    insured_vals = df["insured_ori"].to_numpy()

    changed_polis = changed_slip = changed_ins = 0
    rows_with_certificate = 0

    for n_done, pos in enumerate(target_idx, 1):
        if n_done % 5000 == 0:
            print(f"      Progress: {n_done:,} / {total_target:,} baris target diproses...")

        raw_polis = polis_vals[pos]
        raw_slip = slip_vals[pos]
        raw_ins = insured_vals[pos]

        pairs = get_policy_certificate_pairs(raw_polis)
        c_slip = clean_slip(raw_slip)
        c_ins = clean_insured(raw_ins)

        if pairs and any(c for _, c in pairs):
            rows_with_certificate += 1

        if not pd.isna(raw_polis):
            orig_norm = _normalize_spaces(str(raw_polis).upper())
            c_polis_only = [p for p, _ in pairs]
            if c_polis_only != [orig_norm] or any(c for _, c in pairs):
                changed_polis += 1
        if not pd.isna(raw_slip):
            orig_norm = _normalize_spaces(str(raw_slip).upper())
            if c_slip != [orig_norm]:
                changed_slip += 1
        if not pd.isna(raw_ins):
            if c_ins != [_normalize_spaces(str(raw_ins)).upper()]:
                changed_ins += 1

        max_polis = max(max_polis, len(pairs))
        max_slip = max(max_slip, len(c_slip))
        max_ins = max(max_ins, len(c_ins))

        all_polis_cert_pairs[pos] = pairs
        all_clean_slip[pos] = c_slip
        all_clean_ins[pos] = c_ins

    print(f"      Selesai diproses! max kolom -> polis={max_polis}, slip={max_slip}, insured={max_ins}")
    print(f"      Baris dengan certificate ditemukan: {rows_with_certificate:,}")

    # Kolom interleaved: clean polis N | certificate N (bukan blok terpisah).
    polis_col_names = []
    for i in range(1, max_polis + 1):
        polis_col = f"clean polis {i}"
        cert_col = f"certificate {i}"
        df[polis_col] = [pairs[i - 1][0] if i - 1 < len(pairs) else None for pairs in all_polis_cert_pairs]
        df[cert_col] = [pairs[i - 1][1] if i - 1 < len(pairs) else None for pairs in all_polis_cert_pairs]
        polis_col_names.append((polis_col, cert_col))

    new_columns = []
    for col in df.columns:
        if col.startswith("clean polis ") or col.startswith("certificate "):
            continue
        new_columns.append(col)
        if col == "polis_ori":
            for polis_col, cert_col in polis_col_names:
                new_columns.append(polis_col)
                new_columns.append(cert_col)
        elif col == "slip_ori":
            new_columns += _insert_clean_columns(df, all_clean_slip, "slip", max_slip)
        elif col == "insured_ori":
            new_columns += _insert_clean_columns(df, all_clean_ins, "insured", max_ins)

    df = df[new_columns]
    total_output_rows = len(df)

    print(f"[4/5] Validasi baris ...")
    print("-" * 60)
    print(f"INPUT ROWS        : {total_input_rows:,}")
    print(f"OUTPUT ROWS       : {total_output_rows:,}")
    print(f"TARGET CEDANT     : {CEDANT_VALUE}")
    print(f"TARGET ROWS       : {total_target:,}")
    print(f"OTHER CEDANT ROWS : {total_other:,}")
    print("-" * 60)
    print(f"CLEANED (polis)   : {changed_polis:,}  | UNCHANGED: {total_target - changed_polis:,}")
    print(f"CLEANED (slip)    : {changed_slip:,}  | UNCHANGED: {total_target - changed_slip:,}")
    print(f"CLEANED (insured) : {changed_ins:,}  | UNCHANGED: {total_target - changed_ins:,}")
    print("-" * 60)
    assert total_input_rows == total_output_rows, "JUMLAH BARIS INPUT != OUTPUT! PROSES DIHENTIKAN."
    assert total_target + total_other == total_input_rows
    print("VALIDASI OK: INPUT ROWS == OUTPUT ROWS, tidak ada baris yang hilang/bertambah.")

    print(f"[5/5] Menyimpan hasil ke: {output_file} ...")
    _fast_write_excel(df, output_file)
    print(f"[OK] Selesai. Total baris={total_output_rows:,}, kolom={len(df.columns)}, target={total_target:,}")


if __name__ == "__main__":
    print("=" * 80)
    print("TEST ENTRY POINT - CLEANING FACUL - MULTI ARTHA GUNA")
    print("(TANPA logic certificate / CLSDT - hanya untuk Data 1 FACUL)")
    print("=" * 80)

    print("\n--- TEST FAC_POLICY_NO ---")
    polis_tests = [
        # Pola EX Policy No (spesifik MAG, harus ambil polis aktif saja)
        "45040118001473 Ex. Policy No : 45040117001948",
        "36040119000076 EX. POLICY NO : 36040118000085",
        "45090119000635 Ex 45090118000303",
        "45040518000083 Ex.45040517000081",
        # TBA sebagai placeholder (harus di-skip)
        "05010921011067 + TBA",
        "01013025000694+48013025002218+01013225000608+TBA",
        # TBA murni / ambigu -> keep original
        "TBA/MAG",
        "TBA - 40010917014875.000582 - 000588",
        # Dash chain perulangan polis (bukan certificate)
        "13010921002655-13010921002644-13010921002677-2666",
        # Plus chain > MAX_SPLIT_COLS -> keep original
        "40012121020589 + 143 + 154 + 934 + 945 + 578",
        # Ampersand dua polis penuh -> boleh displit
        "05010919008355&05010519003327",
        # Ampersand ambigu (certificate-like suffix pendek) -> keep original
        "45031125000863 - 000001 & 000002",
        # Comma dua polis penuh -> boleh displit (bukan certificate)
        "01031118000012-000006,01031118000012-000007",
        # Comma dengan suffix pendek ambigu (>5 hasil) -> keep original
        "36080519000-369,371,382,405,416,484,495,507",
        # P1 murni
        "P1",
        # PENYELESAIAN (deskriptif, keep original)
        "PENYELESAIAN HUTANG PIUTANG MAG 2024 IDR TW1",
        # Data korup/gabungan tanpa separator -> keep original (safety cap)
        "28218919120321422523624725826927129330531632733TBA",
        # Polis biasa
        "45040120001088",
        # Comma chain: increasing suffix pattern (harus direkonstruksi)
        "05030319049722,05030319049733,05030319049744,49755",
        # Comma chain: dua polis penuh dengan dash-suffix masing2
        "08080519000166 - 000001,08080519000166 - 000002",
        # Comma chain dengan S/D di dalamnya -> ambigu, keep original
        "01020118000018 - 000001 S/D 000003, 05020118001155",
    ]
    for i, t in enumerate(polis_tests, 1):
        print(f"[{i:2d}] ORI : {t!r}")
        print(f"      HASIL: {clean_polis(t)}")

    print("\n--- TEST FAC_SLIP ---")
    slip_tests = [
        "BORDERO SEPTEMBER 2019",
        "BORDERO AGUSTUS 2023 - USD",
        "4503111800016 S/D 4503112001706",
        "5001302100046 + 041 + 044",
        "4609012300003  /  END",
        "3601102200029 - 3601102200028",
        "P1",
        "4501301957947-0487-9669-0491-9673-8858",
    ]
    for i, t in enumerate(slip_tests, 1):
        print(f"[{i:2d}] ORI : {t!r}")
        print(f"      HASIL: {clean_slip(t)}")

    print("\n--- TEST FAC_INSURED ---")
    insured_tests = [
        "SRIBOGA FLOUR MILL,PT",
        "HARAPAN LANGGENG ABADI/THERESIA MONITA PAIMANTA/ANDY SUTRISNO",
        "BANK INDEX SELINDO QQ ARTHALAUT BUMIJASA",
        "HUTAMA KARYA (PERSERO) AS PRINCIPAL AND/OR PT ADHI KARYA (PERSERO)",
        "MAHAMERU TEGAR SENTOSA QQ THIO RIANDY THEONARDO, PT",
        "LAUTAN REZEKI/FURNILUX INDONESIA/CIA SENG/MURTONO/ARYA CHANDRA",
    ]
    for i, t in enumerate(insured_tests, 1):
        print(f"[{i:2d}] ORI : {t!r}")
        print(f"      HASIL: {clean_insured(t)}")

    print("\n--- TEST CERTIFICATE (reuse persis logic Etiqa) ---")
    cert_tests = [
        "1071031117000018 - 000176",
        "1071031117000018 - 000162-000170",
        "1071031124000024 - 000251 S/D 0000131",
        "1010031116000176 - 000118/1119/1120",
        "1010031116000187 - 000379/380/381/382/VARIOUS",
        # Comma-chain multi-polis MAG (harus TETAP [] karena clean_polis
        # sudah memecahnya jadi 2 token -> tidak boleh dobel jadi certificate)
        "01031118000012-000006,01031118000012-000007",
        # Polis biasa tanpa certificate
        "45040120001088",
    ]
    for i, t in enumerate(cert_tests, 1):
        cp = clean_polis(t)
        cc = clean_certificate(t) if len(cp) == 1 else []
        print(f"[{i}] ORI          : {t!r}")
        print(f"    clean_polis  : {cp}")
        print(f"    certificate  : {cc}")

    print("\n--- TEST RULE BARU: PEMISAHAN POLIS & CERTIFICATE (get_policy_certificate_pairs) ---")
    new_rule_tests = [
        # PRINSIP UTAMA
        ("36040118000154-000001+000002", [("36040118000154", "000001,000002")]),
        ("45013024030214-000001+124002553-000001", [("45013024030214", "000001"), ("124002553", "000001")]),
        ("45013024018823-0001+45013224012546-0001", [("45013024018823", "0001"), ("45013224012546", "0001")]),
        ("45013024027219-000001+45013124002245-000001", [("45013024027219", "000001"), ("45013124002245", "000001")]),
        ("40013024000513 - 000005+40013224000574 - 000005", [("40013024000513", "000005"), ("40013224000574", "000005")]),
        # POLIS SAMA + BANYAK CERTIFICATE
        ("45040118000745 - 000001 / 45040118000745 - 000002", [("45040118000745", "000001,000002")]),
        ("45040117010905 - 000001+45040117010905 - 000002", [("45040117010905", "000001,000002")]),
        ("45040117009628 - 000001+45040117009628 - 000002", [("45040117009628", "000001,000002")]),
        ("06100418000013-000001,06100418000013-000002", [("06100418000013", "000001,000002")]),
        ("36040118000336-000001+36040118000336-000002", [("36040118000336", "000001,000002")]),
        ("36040118000405 - 000001+36040118000405 - 000002", [("36040118000405", "000001,000002")]),
        ("36040119000032 - 000001,36040119000032 - 000002", [("36040119000032", "000001,000002")]),
        ("08080519000166 - 000001,08080519000166 - 000002", [("08080519000166", "000001,000002")]),
        ("01031118000012-000006,01031118000012-000007", [("01031118000012", "000006,000007")]),
        ("40030321000148-000001/40030321000148-000001", [("40030321000148", "000001")]),  # dedup
        # CERTIFICATE TANPA POLIS DIULANG (elision)
        ("07080518000014-00001,00002,00004", [("07080518000014", "00001,00002,00004")]),
        ("40010520009476-000002 + 003 + 004 + 005 + 006", [("40010520009476", "000002,003,004,005,006")]),
        ("03010924000363-000001+000002+000003+000004", [("03010924000363", "000001,000002,000003,000004")]),
        # JANGAN SALAH TEBAK -> None dari rule baru, fallback ke rule lama (keep original)
        ("05012920001475-987-998-009-022-033-044-146", None),
        ("40012120055055-491-033-044-009-088-022077011066998", None),
        ("02031125000018 000455-000536,000467/VARIOUS", None),
        ("02031125000018 000134,000221-000409", None),
        # RULE S/D
        ("08080517000105-000001 S/D 000006", [("08080517000105", "000001 S/D 000006")]),
        # 2 POLIS + CERTIFICATE (base kedua tanpa suffix)
        ("08080517000105-000001 S/D 000006 , 08080517000116-",
         [("08080517000105", "000001 S/D 000006"), ("08080517000116", None)]),
        # RULE BARU LAGI: certificate wajib mulai dari "0"; bare-base
        # (tanpa dash) + angka pendek BUKAN certificate.
        ("0503031704-0614,0716,0727,2319,3481,3594", None),  # tidak semua cert mulai "0" -> keep original
        ("45013018008305+8007321", "PERULANGAN"),  # bare base + 1 elision -> fallback ke rule lama
        ("01020117000803-000001/01020117000803-000004/VARIOUS",
         [("01020117000803", "000001,000004")]),  # base sama + VARIOUS di akhir
        ("36011018000123+5180737+5180737+180123+180123+18073",
         [("36011018000123+5180737+5180737+180123+180123+18073", None)]),  # bare base + 2+ elision -> keep original
    ]
    all_ok = True
    for ori, expected in new_rule_tests:
        result = get_policy_certificate_pairs(ori)
        if expected == "PERULANGAN":
            ok = len(result) == 2 and result[0][1] is None and result[1][1] is None
        else:
            ok = (expected is None) or (result == expected)
        status = "OK" if ok else "MISMATCH"
        if not ok:
            all_ok = False
        print(f"[{status}] ORI: {ori!r}")
        for i, (p, c) in enumerate(result, 1):
            print(f"        clean polis {i} = {p!r:20s} certificate {i} = {c!r}")
        if expected is not None and not ok:
            print(f"        EXPECTED       : {expected}")
    print("\nSEMUA TEST RULE BARU:", "PASS" if all_ok else "ADA YANG GAGAL, CEK DI ATAS")

    print("\n" + "=" * 80)
    print("TEST SELESAI - LOGIC SUDAH DIVALIDASI TERHADAP DATA ASLI")
    print("=" * 80)

    process_data(INPUT_FILE, OUTPUT_FILE)