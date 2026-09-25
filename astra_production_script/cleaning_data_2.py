import os
import re
from typing import Any, List, Optional

import numpy as np
import pandas as pd

DEFAULT_RAW = os.path.join("dataExcel", "raw", "2b. transaksi Osbal 01.01.23 - 17.08.26.xlsx")
INPUT_FILE = DEFAULT_RAW if os.path.exists(DEFAULT_RAW) else os.path.join("data", "osbal.xlsx")
SHEET_NAME = "Query result"
DEFAULT_OUT = os.path.join("dataExcel", "processed", "astrabuana_output_osbal.xlsx")
OUTPUT_FILE = DEFAULT_OUT if os.path.exists(os.path.dirname(DEFAULT_OUT)) else os.path.join("data", "astrabuana_output_osbal.xlsx")

CEDANT_COL = "CCOS_COMP_NAME"
CEDANT_VALUE = "PT ASURANSI ASTRA BUANA"

# Kolom baru dari update data 2 (CLSDT). Prioritas: CLSDT -> fallback ke FAC
# jika CLSDT kosong atau isinya generik (mis. "SISA OSBAL" / "SEPTEMBER 2025 EUR").
CLSDT_POLICY_COL = "CLSDT_POLICY_NO"
CLSDT_SLIP_COL = "CLSDT_SLIP_NO"
FAC_POLICY_COL = "FAC_POLICY_NO"
FAC_SLIP_COL = "FAC_SLIP"

# Kolom sumber untuk CERTIFICATE (diambil dari data 2, rules_certificate1.txt)
CLSDT_SERTF_COL = "CLSDT_SERTF_NO"
CERTIFICATE_COL = "CERTIFICATE_1"

MINIMAL_DIGIT_NOMOR_POLIS = 11
MINIMAL_PANJANG_KODE = 6
MINIMAL_PANJANG_KODE_RANGE = 8
MAX_HASIL_BREAKDOWN = 5

# regex label & pembersihan teks umum
LABEL_WORDS_RE = re.compile(
    r"""
    \bTBA\b|\bVARIOUS\b|\bVAR\b|\bBORD\b|\bBORDERO\b|\bBORDEROUX\b|\bSUMMARY\b|\bASTRA\b|\bSISA\b|\bOSBAL\b
    | \b(?:JANUARI|FEBRUARI|MARET|APRIL|MEI|JUNI|JULI|AGUSTUS|SEPTEMBER|OKTOBER|NOVEMBER|DESEMBER)\b
    | \b(?:JANUARY|FEBRUARY|MARCH|APRIL|MAY|JUNE|JULY|AUGUST|OCTOBER|DECEMBER)\b
    | \b(?:JAN|FEB|MAR|APR|MEI|JUN|JUL|AGU|AGS|SEP|OKT|NOV|DES|AUG|OCT|DEC)\b
    | \b(?:IDR|USD|JPY|SGD|EUR|GBP)\b
    | \b20[0-9]{2}\b
    """,
    re.IGNORECASE | re.VERBOSE,
)

PLACEHOLDER_TOKENS_RE = re.compile(r"\+?\s*\bP\d{1,3}\b\s*\+?", re.IGNORECASE)
VARIOUS_STANDALONE_RE = re.compile(
    r"^\s*VARIOUS(?:\s*\([^)]*\))?(?:\s+(?:INCLUDE\s+)?FISHING\s+VESSEL)?\s*$", re.IGNORECASE
)
VARIOUS_PREFIX_RE = re.compile(
    r"^\s*VARIOUS(?:\s*\([^)]*\))?(?:\s+(?:INCLUDE\s+)?FISHING\s+VESSEL)?\s*/\s*", re.IGNORECASE
)
VARIOUS_SUFFIX_RE = re.compile(
    r"(?:\s*/\s*|\s*,\s*PT\s*/\s*)VARIOUS(?:\s*\([^)]*\))?(?:\s+(?:INCLUDE\s+)?FISHING\s+VESSEL)?\s*$",
    re.IGNORECASE,
)
BORDERO_RE = re.compile(r"\s*\b(?:BORDERO|BORDEROUX)\b.*$", re.IGNORECASE)
INSURED_ENTITIES_RE = re.compile(
    r"[,.\s]*\b(PT|CV|TBK|PTE|LTD|PERSERO|PELAYARAN)\b(?!\w)[,.\s]*", re.IGNORECASE
)
TRAILING_SINGLE_LETTER_TYPO_RE = re.compile(r"[,.\s]+\b[A-Za-z]\b$", re.IGNORECASE)
BRANCH_ANNOTATION_RE = re.compile(
    r"\bCAB(?:ANG)?(?:\s*\.\s*|\s+)[A-Z0-9_.]+\b|\bCAB(?:ANG)?\b", re.IGNORECASE
)
PAREN_ANNOTATIONS_RE = re.compile(
    r"\([^)]*KNOWN\s+AS[^)]*\)|\(\s*APARTEMEN\b[^\)]*\)|\(\s*PERSERO\s*\)|\(\s*[A-Z0-9]\s*\)",
    re.IGNORECASE,
)
INSURED_KEEP_AS_IS_RE = re.compile(
    r"FACILITY|LINESLIP|OPEN COVER|ASTRA OTOPARTS|DIV\b|DIVISI\b", re.IGNORECASE
)
QQ_RE = re.compile(r"\bQ\.?Q\.?\b", re.IGNORECASE)
SD_RE = re.compile(r"(?:\b|(?<=\d))S\s*/\s*D(?:\b|(?=\d))", re.IGNORECASE)
KATA_SD_RE = re.compile(r"\b(?:S\s*/\s*D|SD)\b|(?<=\d)SD(?=\d)", re.IGNORECASE)
STRIP_LABEL_CLEAN_RE = re.compile(r"^[\s/\-:,.+'’‘]+|[\s/\-:,.+'’‘]+$")
TBA_VAR_RE = re.compile(r"\b(TBA|VARIOUS|VAR)\b", re.IGNORECASE)
NON_ALPHANUM_RE = re.compile(r"[^A-Za-z0-9]")
MULTIPLE_SPACES_RE = re.compile(r"\s+")

POLIS_VALID_FORMAT_RE = re.compile(r"\d{11,}")
POLIS_SUFFIX_FORMAT_RE = re.compile(r"\d{7,10}[+,\-]\d{2,4}")
RANGE_PATTERN_RE = re.compile(r"[A-Za-z]{2,8}\d{8,}")
ABBREVIATION_RE = re.compile(r"\b(?:[A-Za-z]\.){2,}")

# --- Aturan kolom CERTIFICATE (lihat Rules_certificate1.txt) ---
CERTIFICATE_SD_RE = re.compile(r"\bS\s*/\s*D\b", re.IGNORECASE)
CERTIFICATE_RANGE_RE = re.compile(r"^\s*(\d{4,6})\s*(?:-|SD)\s*(\d{4,6})\s*$", re.IGNORECASE)
CERTIFICATE_SINGLE_RE = re.compile(r"^\s*(\d{4,6})\s*$")
CERTIFICATE_DIGIT_LEN = 6
CERTIFICATE_MAX_RANGE_UNTUK_DIPECAH = (
    3  # jumlah anggota range <= 3 -> dipecah pakai koma, > 3 -> pakai "SD"
)

# Dideteksi otomatis -> mapping manual per kasus
MANUAL_INSURED_SPLIT = {
    "CEMERLANG COKE INDUSTRIAL SDN BHD AMBANK ISLAMIC BERHAD F.T.": [
        "CEMERLANG COKE INDUSTRIAL SDN BHD",
        "AMBANK ISLAMIC BERHAD F.T.",
    ],
}


# helper pembersihan teks & pemilihan sumber data (CLSDT vs FAC)
def _hapus_petik(text: Any) -> str:
    """Menghapus berbagai jenis karakter tanda petik tunggal (' / ’ / ‘)."""
    return str(text).replace("'", "").replace("’", "").replace("‘", "")


def _clean_code(text: Any) -> str:
    return NON_ALPHANUM_RE.sub("", _hapus_petik(text)).upper()


def _hapus_kata_label(text: Any) -> str:
    text_str = _hapus_petik(text).upper()
    text_str = LABEL_WORDS_RE.sub("", text_str)
    text_str = PLACEHOLDER_TOKENS_RE.sub("", text_str)
    return text_str


def _rapikan_teks_label(text: Any) -> str:
    text_str = _hapus_petik(text).upper()
    text_str = TBA_VAR_RE.sub("", text_str)
    text_str = PLACEHOLDER_TOKENS_RE.sub("", text_str)
    text_str = STRIP_LABEL_CLEAN_RE.sub("", text_str)
    return MULTIPLE_SPACES_RE.sub(" ", text_str).strip()


def _standarisasi_sd(text: Any) -> str:
    """Mengubah variasi S/D, s/d, S / D menjadi SD dengan spasi yang rapi."""
    if pd.isna(text) or text is None:
        return ""
    text_str = _hapus_petik(text)
    res = re.sub(r"\s*S\s*/\s*D\s*", " SD ", text_str, flags=re.IGNORECASE)
    res = re.sub(r"(?<=\w)SD(?=\w)", " SD ", res)
    return MULTIPLE_SPACES_RE.sub(" ", res).strip()


def _pisah_jadi_segmen(text: str) -> List[str]:
    text_sd = _standarisasi_sd(text)
    penanda_sd = "§SD§"
    teks_terlindungi = re.sub(r"\bSD\b", penanda_sd, text_sd)
    segmen_mentah = re.split(r"\s{2,}|/", teks_terlindungi)

    segmen_hasil = []
    for segmen in segmen_mentah:
        segmen = segmen.replace(penanda_sd, "SD").strip()
        if segmen and re.search(r"[A-Za-z0-9]", segmen):
            segmen_hasil.append(segmen)
    return segmen_hasil


def _mengandung_kode_panjang(text: str) -> bool:
    text_clean = _hapus_kata_label(text)
    pattern = rf"[A-Za-z0-9]{{{MINIMAL_PANJANG_KODE},}}"
    return bool(re.search(pattern, text_clean)) and bool(re.search(r"\d", text_clean))


def _clsdt_layak_dipakai(value: Any) -> bool:
    if pd.isna(value):
        return False
    text = _hapus_petik(value).strip()
    if not text:
        return False
    return _mengandung_kode_panjang(text)


def _pilih_sumber_data(row: pd.Series, clsdt_col: str, fac_col: str) -> Any:
    clsdt_val = row[clsdt_col] if clsdt_col in row.index else np.nan
    if _clsdt_layak_dipakai(clsdt_val):
        return clsdt_val
    return row[fac_col] if fac_col in row.index else np.nan


# deteksi kasus khusus: penyelesaian suspend & pola bulan/tahun/currency standalone
POLIS_DASH_SATU_DIGIT_RE = re.compile(r"^(\d{6,})\s*-\s*\d$")

_BULAN_ALTERNATION = (
    r"(?:JANUARI|FEBRUARI|MARET|APRIL|MEI|JUNI|JULI|AGUSTUS|SEPTEMBER|OKTOBER|NOVEMBER|DESEMBER"
    r"|JANUARY|FEBRUARY|MARCH|MAY|JUNE|JULY|AUGUST|OCTOBER|DECEMBER"
    r"|JAN|FEB|MAR|APR|AGU|AGS|SEP|OKT|NOV|DES|AUG|OCT|DEC|JUN|JUL)"
)
_CURRENCY_ALTERNATION = r"(?:IDR|USD|JPY|SGD|EUR|GBP)"

BULAN_TAHUN_CURRENCY_STANDALONE_RE = re.compile(
    rf"^\s*{_BULAN_ALTERNATION}\s+20\d{{2}}\s*-\s*{_CURRENCY_ALTERNATION}\s*$",
    re.IGNORECASE,
)

PENYELESAIAN_SUSPEND_RE = re.compile(
    rf"^PENYELESAIAN\s+SUSPEND(?:\s+[A-Za-z]+)?\s+20\d{{2}}\s+{_CURRENCY_ALTERNATION}$",
    re.IGNORECASE,
)


def _bersihkan_pola_dash_satu_digit(value: Any) -> Any:
    if pd.isna(value):
        return value
    text = _hapus_petik(value).strip()
    m = POLIS_DASH_SATU_DIGIT_RE.match(text)
    return m.group(1) if m else text


def _cek_pola_penyelesaian_suspend(fac_value: Any, clsdt_value: Any) -> Optional[str]:
    if pd.isna(fac_value) or pd.isna(clsdt_value):
        return None
    fac_text = _hapus_petik(fac_value).strip()
    clsdt_text = _hapus_petik(clsdt_value).strip()
    if not fac_text or not clsdt_text:
        return None
    if PENYELESAIAN_SUSPEND_RE.match(fac_text):
        return f"{fac_text.upper()} {clsdt_text.upper()}"
    return None


def _cek_pola_bulan_tahun_currency_standalone(value: Any) -> bool:
    if pd.isna(value):
        return False
    text = _hapus_petik(value).strip()
    if not text:
        return False
    return bool(BULAN_TAHUN_CURRENCY_STANDALONE_RE.match(text))


def _tentukan_nilai_clean_khusus(
    row: pd.Series, fac_col: str, clsdt_col: str, kind: str
) -> Optional[str]:
    fac_value = row[fac_col] if fac_col in row.index else np.nan
    clsdt_value = row[clsdt_col] if clsdt_col in row.index else np.nan

    gabungan = _cek_pola_penyelesaian_suspend(fac_value, clsdt_value)
    if gabungan:
        return _standarisasi_sd(gabungan)

    if kind == "slip" and _cek_pola_bulan_tahun_currency_standalone(fac_value):
        if pd.notna(clsdt_value) and str(clsdt_value).strip():
            return _standarisasi_sd(_hapus_petik(clsdt_value).strip().upper())

    return None


# breakdown polis/slip: deteksi pola & pemecahan segmen
def _mengandung_kata_sd(text: str) -> bool:
    return KATA_SD_RE.search(text) is not None


def _sepertinya_pola_range(text: str) -> bool:
    kode_kode = RANGE_PATTERN_RE.findall(text)
    if len(kode_kode) != 2:
        return False

    kode_awal, kode_akhir = kode_kode
    if kode_awal.upper() == kode_akhir.upper():
        return False

    prefix_awal = re.match(r"[A-Za-z]*", kode_awal).group().upper()
    prefix_akhir = re.match(r"[A-Za-z]*", kode_akhir).group().upper()
    if prefix_awal != prefix_akhir:
        return False

    gabungan = f"{kode_awal}-{kode_akhir}".upper()
    return gabungan in re.sub(r"\s+", "", text).upper()


def _berisi_banyak_nomor_utuh(text: str) -> bool:
    if re.search(r"[+,\-]", text):
        return False
    token_token = text.split()
    if len(token_token) < 2:
        return False
    return all(len(re.sub(r"\D", "", t)) >= MINIMAL_DIGIT_NOMOR_POLIS for t in token_token)


def _split_long_digit_token(t_clean: str) -> List[str]:
    L = len(t_clean)
    if L < 22:
        return [t_clean]
    if L % 12 == 0:
        return [t_clean[i : i + 12] for i in range(0, L, 12)]
    if L % 11 == 0:
        return [t_clean[i : i + 11] for i in range(0, L, 11)]

    for split_pos in (11, 12):
        part1 = t_clean[:split_pos]
        part2 = t_clean[split_pos:]
        if len(part2) in (11, 12):
            if part2.startswith(part1[:2]) or part2.startswith(("0", "1", "6", "7", "8", "9")):
                return [part1, part2]

    res = []
    curr = 0
    while curr < L:
        rem = L - curr
        if rem in (11, 12):
            res.append(t_clean[curr:])
            break
        elif rem >= 23:
            res.append(t_clean[curr : curr + 12])
            curr += 12
        elif rem == 22:
            res.append(t_clean[curr : curr + 11])
            curr += 11
        else:
            res.append(t_clean[curr:])
            break
    return res


def _rekonstruksi_nomor_dari_pengulangan(
    teks_tanpa_label: str, acuan_awal: Optional[str] = None
) -> List[str]:
    raw_tokens = [t.strip() for t in re.split(r"[+,\-]", teks_tanpa_label) if t.strip()]

    tokens_expanded = []
    for t in raw_tokens:
        t_clean = _clean_code(t)
        if not t_clean:
            continue
        if t_clean.isdigit() and len(t_clean) >= 22:
            tokens_expanded.extend(_split_long_digit_token(t_clean))
        else:
            tokens_expanded.append(t_clean)

    if not tokens_expanded:
        return []

    hasil = []
    nomor_acuan_terakhir = acuan_awal

    for tok in tokens_expanded:
        if not tok.isdigit():
            hasil.append(tok)
            nomor_acuan_terakhir = tok
            continue

        if len(tok) >= MINIMAL_DIGIT_NOMOR_POLIS:
            hasil.append(tok)
            nomor_acuan_terakhir = tok
        elif len(tok) in (7, 8, 9) and (
            nomor_acuan_terakhir is None or len(nomor_acuan_terakhir) < 11
        ):
            nomor_acuan_terakhir = tok
        elif nomor_acuan_terakhir is not None:
            if len(nomor_acuan_terakhir) in (7, 8, 9) and (
                len(nomor_acuan_terakhir) + len(tok) in (11, 12)
            ):
                full_num = nomor_acuan_terakhir + tok
                hasil.append(full_num)
                nomor_acuan_terakhir = full_num
            elif len(tok) < len(nomor_acuan_terakhir):
                reconstructed = nomor_acuan_terakhir[: -len(tok)] + tok
                hasil.append(reconstructed)
                nomor_acuan_terakhir = reconstructed
            else:
                hasil.append(tok)
                nomor_acuan_terakhir = tok
        else:
            hasil.append(tok)
            nomor_acuan_terakhir = tok

    return hasil


def _proses_satu_segmen(
    text: str, wajib_11_digit: bool, acuan_awal: Optional[str] = None
) -> List[str]:
    if not _mengandung_kode_panjang(text):
        cleaned = _rapikan_teks_label(text)
        return [_standarisasi_sd(cleaned)]

    if _mengandung_kata_sd(text) or _sepertinya_pola_range(text):
        return [_standarisasi_sd(text)]

    if _berisi_banyak_nomor_utuh(text):
        hasil = [_clean_code(t) for t in text.split()]
        return hasil if len(hasil) <= MAX_HASIL_BREAKDOWN else [_standarisasi_sd(text)]

    if wajib_11_digit:
        has_valid_polis_format = bool(POLIS_VALID_FORMAT_RE.search(text)) or bool(
            POLIS_SUFFIX_FORMAT_RE.search(text)
        )
        if not has_valid_polis_format:
            kode_pertama = re.split(r"[+,\-]", text)[0]
            if len(re.sub(r"\D", "", kode_pertama)) < MINIMAL_DIGIT_NOMOR_POLIS:
                return [_standarisasi_sd(text)]

    teks_tanpa_placeholder = PLACEHOLDER_TOKENS_RE.sub("", text)
    hasil_breakdown = _rekonstruksi_nomor_dari_pengulangan(teks_tanpa_placeholder, acuan_awal)

    if not hasil_breakdown or len(hasil_breakdown) > MAX_HASIL_BREAKDOWN:
        return [_standarisasi_sd(text)]

    return [_standarisasi_sd(h) for h in hasil_breakdown]


# insured
def clean_insured(name: Any) -> str:
    if pd.isna(name):
        return ""

    name_str = _hapus_petik(name)
    name_str = PAREN_ANNOTATIONS_RE.sub("", name_str)
    name_str = BRANCH_ANNOTATION_RE.sub("", name_str)
    name_str = TRAILING_SINGLE_LETTER_TYPO_RE.sub("", name_str)
    name_str = INSURED_ENTITIES_RE.sub(" ", name_str)

    name_str = re.sub(r",", " ", name_str)
    name_str = re.sub(r"\.\s+$", "", name_str)
    name_str = re.sub(r"^[\s/\-:,.+&;]+|[\s/\-:,.+&;]+$", "", name_str)

    cleaned = MULTIPLE_SPACES_RE.sub(" ", name_str).strip()
    return cleaned.upper() if len(cleaned) > 1 else ""


def preprocess_insured_text(raw_str: str) -> str:
    text = VARIOUS_PREFIX_RE.sub("", raw_str)
    text = VARIOUS_SUFFIX_RE.sub("", text)
    text = BORDERO_RE.sub("", text)
    text = PAREN_ANNOTATIONS_RE.sub("", text)
    return text.strip()


def breakdown_insured(value: Any) -> List[str]:
    if pd.isna(value):
        return [value]

    raw_str = _hapus_petik(value).strip()
    if not raw_str:
        return [raw_str]

    if VARIOUS_STANDALONE_RE.match(raw_str):
        return [raw_str.upper()]

    text = preprocess_insured_text(raw_str)
    if not text:
        return [raw_str.upper()]

    manual_split = MANUAL_INSURED_SPLIT.get(text.upper())
    if manual_split:
        return [clean_insured(p) for p in manual_split if clean_insured(p)]

    if INSURED_KEEP_AS_IS_RE.search(text) or "POLYCHEM" in text.upper():
        cleaned = clean_insured(text)
        return [cleaned] if cleaned else [raw_str.upper()]

    split_pattern = r"\bQ\.?Q\.?\b|/|:|\s+-\s+|\s*,\s*"
    cleaned_parts = [clean_insured(p) for p in re.split(split_pattern, text) if clean_insured(p)]

    if len(cleaned_parts) > 5:
        return [raw_str.upper()]

    return cleaned_parts if cleaned_parts else [clean_insured(raw_str)]


# polis & slip
def breakdown_polis_atau_slip(value: Any, wajib_11_digit: bool) -> List[str]:
    """Fungsi generic untuk memecah Polis atau Slip."""
    if pd.isna(value):
        return [value]

    raw = _hapus_petik(value).strip().upper()
    if not raw:
        return [_standarisasi_sd(raw)]
    if not _mengandung_kode_panjang(raw):
        return [_standarisasi_sd(raw)]

    text = _hapus_kata_label(raw)
    if not text.strip():
        return [_standarisasi_sd(raw)]

    segmen_segmen = _pisah_jadi_segmen(text)
    if not segmen_segmen:
        return [_standarisasi_sd(raw)]

    ada_segmen_berkode = any(_mengandung_kode_panjang(s) for s in segmen_segmen)
    semua_hasil = []
    acuan_terakhir = None

    for segmen in segmen_segmen:
        berkode = _mengandung_kode_panjang(segmen)
        if ada_segmen_berkode and not berkode:
            continue

        hasil_segmen = _proses_satu_segmen(segmen, wajib_11_digit, acuan_terakhir)
        semua_hasil.extend(hasil_segmen)

        if hasil_segmen and re.fullmatch(r"[A-Za-z0-9]+", hasil_segmen[-1]):
            acuan_terakhir = hasil_segmen[-1]

    hasil_final = []
    for item in semua_hasil:
        if not hasil_final or item != hasil_final[-1]:
            hasil_final.append(item)

    hasil_target = hasil_final if hasil_final and len(hasil_final) <= MAX_HASIL_BREAKDOWN else [raw]
    return [_standarisasi_sd(x) if isinstance(x, str) else x for x in hasil_target]


def clean_polis(value: Any) -> Any:
    if pd.isna(value):
        return value
    res = _rapikan_teks_label(value)
    return _standarisasi_sd(res)


def breakdown_polis(value: Any) -> List[str]:
    return breakdown_polis_atau_slip(value, wajib_11_digit=True)


def clean_slip(value: Any) -> Any:
    if pd.isna(value):
        return value
    res = _rapikan_teks_label(value)
    return _standarisasi_sd(res)


def breakdown_slip(value: Any) -> List[str]:
    return breakdown_polis_atau_slip(value, wajib_11_digit=False)


# certificate
def _pad_certificate_digit(digit_str: str) -> str:
    """Tambahkan nol di depan agar menjadi 6 digit, contoh 00012 -> 000012."""
    return digit_str.zfill(CERTIFICATE_DIGIT_LEN)


def process_certificate(value: Any) -> Any:
    """
    Mengisi kolom CERTIFICATE dari kolom CLSDT_SERTF_NO sesuai
    Rules_certificate1.txt:
    - "S/D" atau "s/d" disamakan jadi "SD"
    - format kurang dari 6 digit ditambahkan nol di depan jadi 6 digit
    - range dengan jumlah anggota > 3 ditulis "AWAL SD AKHIR"
    - range dengan jumlah anggota <= 3 dipecah dengan koma
    - jika CLSDT_SERTF_NO tidak berupa angka 5/6 digit (tunggal atau range),
      kolom CERTIFICATE dikosongkan walau CLSDT_SERTF_NO tidak kosong
    """
    if pd.isna(value):
        return np.nan

    text = _hapus_petik(value).strip()
    if not text:
        return np.nan

    # samakan "s/d" / "S/D" (dengan atau tanpa spasi di sekitar "/") jadi "SD"
    text_norm = CERTIFICATE_SD_RE.sub("SD", text).upper()
    text_norm = MULTIPLE_SPACES_RE.sub(" ", text_norm).strip()

    range_match = CERTIFICATE_RANGE_RE.match(text_norm)
    if range_match:
        awal_str, akhir_str = range_match.groups()
        awal_num, akhir_num = int(awal_str), int(akhir_str)

        if akhir_num < awal_num:
            awal_num, akhir_num = akhir_num, awal_num

        jumlah_anggota = akhir_num - awal_num + 1

        if jumlah_anggota > CERTIFICATE_MAX_RANGE_UNTUK_DIPECAH:
            awal_padded = _pad_certificate_digit(str(awal_num))
            akhir_padded = _pad_certificate_digit(str(akhir_num))
            return f"{awal_padded} SD {akhir_padded}"
        else:
            daftar_nomor = [_pad_certificate_digit(str(n)) for n in range(awal_num, akhir_num + 1)]
            return ", ".join(daftar_nomor)

    single_match = CERTIFICATE_SINGLE_RE.match(text_norm)
    if single_match:
        return _pad_certificate_digit(single_match.group(1))

    # format tidak sesuai kriteria certificate (bukan angka 5/6 digit tunggal maupun range)
    return np.nan


# helper penyisipan kolom ke dataframe
def _insert_certificate_column(
    df: pd.DataFrame,
    source_col: str,
    anchor_col: str,  # <--- Ubah nama parameter menjadi anchor_col
    is_astrabuana_mask: pd.Series,
) -> None:
    """Sisipkan kolom CERTIFICATE tepat setelah kolom anchor (mis. FAC_POLICY_NO_CLEAN_1)."""
    if source_col not in df.columns:
        print(f"[WARNING] Kolom '{source_col}' tidak ditemukan, kolom CERTIFICATE dikosongkan.")
        hasil_certificate = pd.Series(np.nan, index=df.index)
    else:
        hasil_certificate = df.apply(
            lambda row: (
                process_certificate(row[source_col]) if is_astrabuana_mask[row.name] else np.nan
            ),
            axis=1,
        )

    # Dapatkan posisi dari kolom yang diinginkan (anchor_col)
    posisi = df.columns.get_loc(anchor_col)
    df.insert(posisi + 1, CERTIFICATE_COL, hasil_certificate)


def _insert_breakdown_columns(
    df: pd.DataFrame,
    source_col: str,
    fn_breakdown: callable,
    prefix: str,
    is_astrabuana_mask: pd.Series,
    max_cols_target: int = MAX_HASIL_BREAKDOWN,
) -> None:
    hasil_breakdown = df.apply(
        lambda row: fn_breakdown(row[source_col]) if is_astrabuana_mask[row.name] else [], axis=1
    )

    lengths = hasil_breakdown.apply(len)
    max_in_data = int(lengths.max()) if not lengths.empty and lengths.max() > 0 else 1
    max_cols = max(max_in_data, max_cols_target)

    padded_hasil = hasil_breakdown.apply(lambda lst: lst + [np.nan] * (max_cols - len(lst)))

    df_breakdown = pd.DataFrame(
        padded_hasil.tolist(),
        columns=[f"{prefix}_{i + 1}" for i in range(max_cols)],
        index=df.index,
    )

    posisi = df.columns.get_loc(source_col)
    for i, col in enumerate(df_breakdown.columns):
        df.insert(posisi + 1 + i, col, df_breakdown[col])


def _insert_breakdown_columns_combined(
    df: pd.DataFrame,
    fac_col: str,
    clsdt_col: str,
    fn_breakdown: callable,
    prefix: str,
    is_astrabuana_mask: pd.Series,
    max_cols_target: int = MAX_HASIL_BREAKDOWN,
    kind: str = "",
) -> None:

    def _proses_baris(row: pd.Series) -> List[Any]:
        if not is_astrabuana_mask[row.name]:
            return []

        nilai_khusus = _tentukan_nilai_clean_khusus(row, fac_col, clsdt_col, kind)
        if nilai_khusus is not None:
            return [_standarisasi_sd(nilai_khusus)]

        sumber = _pilih_sumber_data(row, clsdt_col, fac_col)
        if pd.isna(sumber):
            return []

        if kind == "polis":
            sumber = _bersihkan_pola_dash_satu_digit(sumber)

        hasil = fn_breakdown(sumber)
        return [_standarisasi_sd(x) if isinstance(x, str) else x for x in hasil]

    hasil_breakdown = df.apply(_proses_baris, axis=1)

    lengths = hasil_breakdown.apply(len)
    max_in_data = int(lengths.max()) if not lengths.empty and lengths.max() > 0 else 1
    max_cols = max(max_in_data, max_cols_target)

    padded_hasil = hasil_breakdown.apply(lambda lst: lst + [np.nan] * (max_cols - len(lst)))

    df_breakdown = pd.DataFrame(
        padded_hasil.tolist(),
        columns=[f"{prefix}_{i + 1}" for i in range(max_cols)],
        index=df.index,
    )

    posisi = df.columns.get_loc(fac_col)
    for i, col in enumerate(df_breakdown.columns):
        df.insert(posisi + 1 + i, col, df_breakdown[col])


# proses data
def process_data(input_file: str, sheet_name: str, output_file: str) -> None:
    print(f"Membaca data dari: {input_file} (Sheet: {sheet_name}) ...")

    if not os.path.exists(input_file):
        print(f"[ERROR] File tidak ditemukan di lokasi: {input_file}")
        return

    df = pd.read_excel(input_file, sheet_name=sheet_name, header=0)
    df.columns = df.columns.str.strip()

    if CEDANT_COL not in df.columns:
        print(f"[ERROR] Kolom '{CEDANT_COL}' tidak ditemukan di file!")
        return

    is_astrabuana = df[CEDANT_COL].astype(str).str.strip() == CEDANT_VALUE

    kolom_sisa = [
        col
        for col in df.columns
        if col.startswith(
            ("INSURED_CLEAN", "POLICY_CLEAN", "SLIP_CLEAN", "SERTIF_CLEAN", "FAC_INSURED_CLN", "FAC_POLICY_CLEAN", "FAC_POLICY_NO_CLEAN", "FAC_SLIP_CLEAN")
        )
        or col.endswith("_SOURCE")
    ]
    if kolom_sisa:
        df.drop(columns=kolom_sisa, inplace=True)

    print("Menjalankan pembersihan dan breakdown kolom...")
    _insert_breakdown_columns(
        df, "FAC_INSURED", breakdown_insured, "INSURED_CLEAN", is_astrabuana
    )

    if CLSDT_POLICY_COL not in df.columns:
        print(
            f"[WARNING] Kolom '{CLSDT_POLICY_COL}' tidak ditemukan, hanya memakai '{FAC_POLICY_COL}'."
        )
        df[CLSDT_POLICY_COL] = np.nan
    if CLSDT_SLIP_COL not in df.columns:
        print(
            f"[WARNING] Kolom '{CLSDT_SLIP_COL}' tidak ditemukan, hanya memakai '{FAC_SLIP_COL}'."
        )
        df[CLSDT_SLIP_COL] = np.nan

    _insert_breakdown_columns_combined(
        df,
        FAC_POLICY_COL,
        CLSDT_POLICY_COL,
        breakdown_polis,
        "POLICY_CLEAN",
        is_astrabuana,
        kind="polis",
    )

    # Kolom SERTIF_CLEAN_1 diisi dari CLSDT_SERTF_NO, diletakkan tepat setelah POLICY_CLEAN_1
    _insert_certificate_column(df, CLSDT_SERTF_COL, "POLICY_CLEAN_1", is_astrabuana)
    if "CERTIFICATE_1" in df.columns:
        df.rename(columns={"CERTIFICATE_1": "SERTIF_CLEAN_1"}, inplace=True)

    _insert_breakdown_columns_combined(
        df,
        FAC_SLIP_COL,
        CLSDT_SLIP_COL,
        breakdown_slip,
        "SLIP_CLEAN",
        is_astrabuana,
        kind="slip",
    )

    df = df.loc[is_astrabuana].reset_index(drop=True)

    # Standarisasi urutan kolom OSBAL CLEAN (Rule 1: SERTIF berurutan setelah POLIS)
    OSBAL_TARGET_COLS = [
        "CCOS_DOC_NO", "CCOS_DATE", "CCOS_REF_CODE", "CCOS_COMP", "CCOS_COMP_NAME", "CCOS_REF_COMP", "CCOS_REF_COMP_NAME",
        "FAC_INSURED", "INSURED_CLEAN_1", "INSURED_CLEAN_2", "INSURED_CLEAN_3", "INSURED_CLEAN_4", "INSURED_CLEAN_5",
        "CCOS_CURR", "CCOS_OR_BAL", "CCOS_BAL_DUE", "CCOS_OR_BAL_IN_IDR", "CCOS_BAL_DUE_IN_IDR",
        "FAC_COM_DATE", "FAC_EXP_DATE", "FAC_DUE_DATES", "FAC_SUB_CLASS",
        "FAC_POLICY_NO", "POLICY_CLEAN_1", "SERTIF_CLEAN_1", "POLICY_CLEAN_2", "SERTIF_CLEAN_2", "POLICY_CLEAN_3", "SERTIF_CLEAN_3", "POLICY_CLEAN_4", "POLICY_CLEAN_5",
        "FAC_SLIP", "SLIP_CLEAN_1", "SLIP_CLEAN_2", "SLIP_CLEAN_3", "SLIP_CLEAN_4", "SLIP_CLEAN_5",
        "CLASS_CODE", "CLASS_NAME", "CLSDT_POLICY_NO", "CLSDT_SLIP_NO", "CLSDT_SERTF_NO"
    ]
    for c in OSBAL_TARGET_COLS:
        if c not in df.columns:
            df[c] = ""
    extra_cols = [c for c in df.columns if c not in OSBAL_TARGET_COLS]
    df = df[OSBAL_TARGET_COLS + extra_cols]

    # Pastikan pada kolom polis clean, slip clean, dan sertif clean semua S/D diganti menjadi SD
    clean_cols_to_check = [
        c for c in df.columns 
        if any(k in c.lower() for k in ["polis", "policy", "slip", "sertif"]) and "clean" in c.lower()
    ]
    for c in clean_cols_to_check:
        df[c] = df[c].map(
            lambda x: _standarisasi_sd(x)
            if pd.notna(x) and re.search(r"S\s*/\s*D", str(x), re.IGNORECASE)
            else x
        )

    print(f"Menyimpan hasil ke Excel: {output_file} ...")

    output_dir = os.path.dirname(output_file)
    if output_dir:
        os.makedirs(output_dir, exist_ok=True)
    df.to_excel(output_file, index=False)

    print("Pemrosesan data 2 selesai!")


if __name__ == "__main__":
    process_data(INPUT_FILE, SHEET_NAME, OUTPUT_FILE)