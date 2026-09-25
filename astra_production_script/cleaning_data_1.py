import os
import re
import pandas as pd
from typing import List, Optional, Union

INPUT_FILE = os.path.join("dataExcel", "raw", "1b. Transaksi Facul 01.01.23 - 17.08.2026.xlsx")
SHEET_NAME = "Query result"
OUTPUT_FILE = os.path.join("dataExcel", "processed", "astrabuana_output_facul.xlsx")

CEDANT_COL = "COMP_NAME"
CEDANT_VALUE = "PT ASURANSI ASTRA BUANA"

MINIMAL_DIGIT_NOMOR_POLIS = 11
MINIMAL_PANJANG_KODE = 6
MINIMAL_PANJANG_KODE_RANGE = 8
MAX_HASIL_BREAKDOWN = 5
VALID_FULL_LENGTHS_POLIS = (11, 12, 13, 16)
SLIP_SUFFIX_ALLOWED_LENGTHS = {2, 3, 4, 6, 7}


class Patterns:
    AS_PER_LIST = re.compile(r"\bA\s*S\s*P?\s*E?\s*R?\s*L\s*I\s*S\s*T\b", re.IGNORECASE)
    CEDANT_REMOVE = re.compile(
        r"\b(?:PT\s+)?ASURANSI\s+ASTRA\s+BUANA\b|\bASTRA\s+BUANA\b", re.IGNORECASE
    )
    PLACEHOLDER_TOKENS = re.compile(r"\+?\s*\bP\d{1,3}\b\s*\+?", re.IGNORECASE)
    SD = re.compile(r"S\s*/\s*D", re.IGNORECASE)
    STRIP_LABEL_CLEAN = re.compile(r"^[\s/\-:,.+&;]+|[\s/\-:,.+&;]+$")
    TBA_VAR = re.compile(r"\b(TBA|VARIOUS|VAR)\b", re.IGNORECASE)
    NON_ALPHANUM = re.compile(r"[^A-Za-z0-9]")
    MULTIPLE_SPACES = re.compile(r"\s{2,}")

    # Regex khusus menghapus titik setelah PT (e.g., "PT." -> "PT")
    CLEAN_PT_DOT = re.compile(r"\bPT\s*\.\s*", re.IGNORECASE)

    # Pola khusus ASTRA EX LS
    ASTRA_EX_LS_PATTERN = re.compile(r"\bASTRA\s+EX\s+LS\b", re.IGNORECASE)

    # Policy Specific
    POLIS_VALID_FORMAT = re.compile(r"\d{11,}")
    POLIS_SUFFIX_FORMAT = re.compile(r"\d{7,10}[+,\-]\d{2,4}")
    RANGE_PATTERN = re.compile(r"[A-Za-z]{2,8}\d{8,}")
    MASTER_POLICY_ANNOTATION = re.compile(r"\(\s*MASTER\b", re.IGNORECASE)
    DOT_GROUPED_CODE = re.compile(r"^[A-Za-z0-9]+(?:\.[A-Za-z0-9]+)+$")
    TAGGED_POLIS = re.compile(r"^\d{8,}(?:/[A-Za-z]{1,5})?$")
    INFO_NOTE_PAREN = re.compile(r"\(\s*LANJUTAN\b[^\)]*\)", re.IGNORECASE)

    # Dates & Slip Patterns
    _BULAN_ALT = (
        r"(?:JAN(?:UARY|UARI)?|FEB(?:RUARY|RUARI)?|MAR(?:CH|ET)?|APR(?:IL)?|MAY|MEI|"
        r"JUN(?:E|I)?|JUL(?:Y|I)?|AUG(?:UST)?|AGU(?:STUS)?|SEP(?:TEMBER)?|OCT(?:OBER)?|"
        r"OKT(?:OBER)?|NOV(?:EMBER)?|DEC(?:EMBER)?|DES(?:EMBER)?)"
    )
    MONTH_YEAR_PREFIX = re.compile(
        rf"^(?:{_BULAN_ALT}\s+)?{_BULAN_ALT}\s+\d{{4}}\s*[-/]?\s*", re.IGNORECASE
    )
    # Unit "BULAN TAHUN - CCY/" yang bisa berulang di awal teks, contoh:
    # "MEI 2023 - JPY/MEI 2023 - JPY/-PMCGP23001060101" (label bulan+tahun+
    MONTH_YEAR_CCY_SLASH_PREFIX = re.compile(
        rf"^{_BULAN_ALT}\s+\d{{4}}\s*-\s*(?:IDR|USD|JPY|SGD|EUR|GBP)\s*/\s*", re.IGNORECASE
    )
    MONTH_YEAR_SUFFIX = re.compile(
        r"\s*[-/]\s*(?:JAN(?:UARY|UARI)?|FEB(?:RUARY|RUARI)?|MAR(?:CH|ET)?|APR(?:IL)?|MAY|MEI|JUN(?:E|I)?|JUL(?:Y|I)?|AUG(?:UST)?|AGU(?:STUS)?|SEP(?:TEMBERI|TEMBER)?|OCT(?:OBER)?|OKT(?:OBER)?|NOV(?:EMBER)?|DEC(?:EMBER)?|DES(?:EMBER)?)\s+\d{4}\s*$",
        re.IGNORECASE,
    )
    TRAILING_NOTE = re.compile(
        r"\s*/\s*(?:P\d+|\bSLIP\b|\bBLM\b|\bBELUM\b|\bDATANG\b).*$", re.IGNORECASE
    )

    _SLIP_MATA_UANG = r"(?:IDR|USD|JPY|SGD|EUR|GBP)"
    _SLIP_KATA_BORDER = r"(?:BORDEREAUX|BORDEROUX|BORDERO|BORDX|BORDR|BORD|BODR|BD)"
    SLIP_CATATAN_EKOR = re.compile(
        rf"\s*/?\s*(?:{_SLIP_MATA_UANG}\s*/\s*)?(?:{_SLIP_KATA_BORDER}\s+)?(?:{_BULAN_ALT}\s+)?{_BULAN_ALT}(?:\.\d{{1,2}}|\s+\d{{4}})?\s*-?\s*(?:{_SLIP_MATA_UANG})?\s*$",
        re.IGNORECASE,
    )
    SLIP_MATA_UANG_EKOR = re.compile(rf"\s*/?\s*-?\s*{_SLIP_MATA_UANG}\s*$", re.IGNORECASE)

    _KATA_LABEL_UNTUK_SLASH = (
        r"(?:TBA|VAR|VARIOUS|SUMMARY|ASTRA|IDR|USD|JPY|SGD|EUR|GBP"
        r"|JAN|FEB|MAR|APR|MEI|JUN|JUL|AGU|AGS|SEP|OKT|NOV|DES|AUG|OCT|DEC)"
    )
    SLIP_KEEP_AS_IS = re.compile(
        r"(?:^\d{11}\s+\d{2,3}(?:\s+\d{2,3})+$|\bFISHVSL\b|\b\d{11}\s*\+\s*\d{3}\s+\d{4}\b|^\d{25,}$|^\d{2,3}(?:&\d{2,3})+$)",
        re.IGNORECASE,
    )

    END_TOKEN = re.compile(r"^(?:END|EN)\.?\d*$", re.IGNORECASE)
    END_TRAILING_MARK = re.compile(r"[\s/.]+(?:END|EN)\b\.?\d*$", re.IGNORECASE)
    END_OR_EN_PRESENT = re.compile(r"\b(?:END|EN)\b", re.IGNORECASE)

    # Pola khusus slip/polis yang harus dibiarkan apa adanya (fac = clean)
    POLIS_KEEP_AS_IS = re.compile(r"^\d{10,13}(?:/\d{2,3}){1,6}/\d{4,8}$")
    # Pola "P1 (P2)" -> dibiarkan apa adanya, JANGAN dihapus oleh PLACEHOLDER_TOKENS
    P_PAREN_P_KEEP_AS_IS = re.compile(r"^P\d{1,3}\s*\(\s*P\d{1,3}\s*\)$", re.IGNORECASE)
    SLIP_JUNK_LABEL_NO_KODE = re.compile(r"^[A-Za-z]+$")
    PREFIX_DASH_DIGIT = re.compile(r"^([A-Za-z]{1,8})-(\d{7,})$")

    # Insured Patterns
    INSURED_ENTITIES = re.compile(
        r"[,.\s]*\b(PT|CV|TBK|PTE|LTD|PELAYARAN)\b(?!\w)[,.\s]*", re.IGNORECASE
    )
    INSURED_PERSERO = re.compile(r"[,.\s]*\(\s*PERSERO\s*\)[,.\s]*|\bPERSERO\b", re.IGNORECASE)
    STRIP_INSURED_BOUNDARIES = re.compile(r"^[\s/\\\-:,.+&]+|[\s/\\\-:,.+&]+$")

    # Rule tambahan insured (lihat: Tambahan insured data 1.txt)
    # 1) Hapus kata "AS CONTRACTOR" / "AS PRINCIPAL"
    AS_CONTRACTOR_PRINCIPAL = re.compile(r"\bAS\s+(?:CONTRACTOR|PRINCIPAL)\b", re.IGNORECASE)
    # 2) Ekor "LTD AND/OR SUBSIDIARIES" ikut dihapus (bukan jadi pemisah breakdown)
    LTD_AND_OR_SUBSIDIARIES = re.compile(r"\bLTD\s+AND\s*/\s*OR\s+SUBSIDIARIES\b", re.IGNORECASE)
    # 3) Pemisah "AND/OR" -> breakdown jadi 2 insured (mis. pola AS PRINCIPAL .. AND/OR .. AS CONTRACTOR)
    AND_OR_SEPARATOR = re.compile(r"\s*\bAND\s*/\s*OR\b\s*", re.IGNORECASE)
    # 4) Prefix "AUTO FACILITY/", "CARGO FACILITY/", "LINESLIP/" (boleh berulang di awal teks)
    FAC_LEADING_PREFIX = re.compile(
        r"^(?:(?:AUTO\s+FACILITY|CARGO\s+FACILITY|LINESLIP)\s*/\s*)+", re.IGNORECASE
    )
    # 5) Penanda entitas yang "menempel" tanpa spasi ke nama berikutnya, mis. "PTSUMBER..."
    GLUED_ENTITY_MARKER = re.compile(r"^(PT|CV|TBK|PTE|LTD)(?=[A-Za-z])", re.IGNORECASE)
    # 6) Kode referensi/bank yang menempel di ekor (mis. "BANK0031") -> bukan insured terpisah
    BANK_CODE_TOKEN = re.compile(r"^[A-Za-z]{2,}\d{2,}$")

    #    prefix AUTO FACILITY dkk dihapus.
    INSURED_EXACT_OVERRIDES = {
        re.sub(r"\s+", " ", s.strip().upper()): hasil
        for s, hasil in [
            (
                "AUTO FACILITY/CARGO FACILITY/LINESLIP/SURYA INDAH NUSANTARA PAGI/ PERKEBUNAN LEMBAH BHAKTI",
                ["SURYA INDAH NUSANTARA", "PERKEBUNAN LEMBAH BHAKTI"],
            ),
            (
                "AUTO FACILITY/CARGO FACILITY/LINESLIP/TUNGGAL PERKASA PLANTATIONS, PT/ SARI LEMBAH SUBUR",
                ["TUNGGAL PERKASA", "SARI LEMBAH SUBUR PLANTATIONS"],
            ),
        ]
    }
    BORDERO_TAIL = re.compile(r"\b(?:BORDEROUX|BORDERO|BORD)\b.*$", re.IGNORECASE)
    SLIP_TAIL_VARIOUS = re.compile(r"\s*/\s*VARIOUS\b.*$", re.IGNORECASE)
    SLIP_TAIL_END_DOT = re.compile(r"\.END(?:\.\d+)?(?:\s*\+\s*\d+)*\s*$", re.IGNORECASE)
    SLIP_TAIL_END_PLUS = re.compile(r"(?:\s*\+\s*END\.\d+)+\s*$", re.IGNORECASE)
    SLIP_FIRST_SEGMENT_CODE = re.compile(r"^[A-Za-z]{2,8}\d[A-Za-z0-9\-]{2,}$")

    SLIP_KOMPLEKS_CONNECTORS = [
        re.compile(r"\s*;\s*"),
        re.compile(r"\s*/\s*"),
        re.compile(r"\s*&\s*"),
        re.compile(r"\s*\+\s*"),
    ]
    SLIP_FIRST_TOKEN_DASH = re.compile(r"^([A-Za-z]{0,8})-?(\d{7,})$")
    STANDALONE_BULAN_TAHUN = re.compile(rf"^{_BULAN_ALT}(?:\s+\d{{4}})?$", re.IGNORECASE)
    SLIP_JUNK_TOKEN = re.compile(
        r"^(?:END(?:\.\d+)?|VAR|TBA|VARIOUS|AS ATTACHMENT|P\d{1,3})\b", re.IGNORECASE
    )
    SLIP_UPLOADED_FILES = re.compile(r"\bSEE\s+UPLOADED\s+FILES\b", re.IGNORECASE)
    SLIP_FIRST_TOKEN = re.compile(r"^([A-Za-z]{2,8})(\d{2,})$")

    SLIP_PENGULANGAN_CONNECTORS = [
        re.compile(r"\s*/\s*"),
        re.compile(r"\s*\+\s*"),
        re.compile(r"\s*-\s*"),
        re.compile(r"\s+"),
    ]
    VARIOUS_PHRASES_PATTERN = (
        r"VARIOUS(?:\s+INCLUDE\s+FISHING\s+VESSEL|\s+FISHING\s+VESSEL|\s*\(\s*LINE\s*SLIP\s*\))?"
    )
    VARIOUS_EXACT = re.compile(rf"^\s*{VARIOUS_PHRASES_PATTERN}\s*$", re.IGNORECASE)
    TARGET_PARENS = re.compile(
        r"\(\s*(?:PERSERO|ALL\s+THE\s+OWNERS\s+OF\s+UNIT\s+APARTEMENT\s+DHARMAWANGSA\s+1|A|B|APARTEMEN\s+GRAND\s+DHIKA\s+BEKASI)\s*\)",
        re.IGNORECASE,
    )
    # Tanda kurung buka ganda yang berlebih, misal "( (KNOWN AS ...)" -> "(KNOWN AS ...)"
    DOUBLE_OPEN_PAREN = re.compile(r"\(\s*\(")
    PAREN_CONTENT = re.compile(r"\([^()]*\)")
    PAREN_ONLY_SEGMENT = re.compile(r"^§PAREN\d+§$")
    MASK_FAC = re.compile(r"\b(AUTO\s+FACILITY|CARGO\s+FACILITY|LINESLIP)\s*/", re.IGNORECASE)

    BORDERO_KEYWORD_PRESENT = re.compile(_SLIP_KATA_BORDER, re.IGNORECASE)
    BORDERO_VOCAB = re.compile(
        rf"\b(?:ASURANSI\s+ASTRA\s+BUANA|ASTRA\s+BUANA|ASTRA|{_SLIP_KATA_BORDER}|FACILITY|SLIP|DI|{_BULAN_ALT}|{_SLIP_MATA_UANG})\b",
        re.IGNORECASE,
    )
    BORDERO_DATE_RANGE = re.compile(r"\b\d{6}\s*-\s*\d{6}\b")
    TOKEN_SAMPAH_SPASI = re.compile(r"^(?:VAR|TBA|VARIOUS|END(?:\.\d+)?|P\d{1,3})$", re.IGNORECASE)

    # Pola-pola eksplisit dari user yang harus dibiarkan apa adanya (fac = clean),
    # tidak boleh di-breakdown walaupun terlihat mirip pola lain yang di-breakdown.
    SLIP_KEEP_AS_IS_EXACT = {
        re.sub(r"\s+", " ", s.strip().upper())
        for s in [
            "PMAHL20021800301, 20021810301,20022060301/ VARIOUS",
            "PMAHL20017290301/20021730301,40301,790301, 21910301/VARIOUS",
        ]
    }


def get_business_partners(comp_name2, comp_name) -> str:

    broker = "" if pd.isna(comp_name2) else str(comp_name2).strip()
    if not broker or broker.upper() == "DIRECT":
        val = "" if pd.isna(comp_name) else str(comp_name).strip()
    else:
        val = broker

    # Hapus titik setelah PT dan rapikan spasi berlebih
    val = Patterns.CLEAN_PT_DOT.sub("PT ", val)
    return Patterns.MULTIPLE_SPACES.sub(" ", val).strip()


def _normalisasi_as_per_list(text: str) -> str:
    return Patterns.AS_PER_LIST.sub("AS PER LIST", text)


def _clean_code(text: str) -> str:
    return Patterns.NON_ALPHANUM.sub("", str(text)).upper()


def _hapus_kata_label(text: str) -> str:
    text = str(text).upper()
    text = Patterns.CEDANT_REMOVE.sub("", text)
    return Patterns.PLACEHOLDER_TOKENS.sub("", text)


def _hapus_single_char_suffix(text: str) -> str:
    pattern = re.compile(r"(?<=[A-Za-z0-9]{3})\s*[/\-.]\s*[A-Za-z0-9]\b(?!\w)")
    prev = ""
    curr = text.strip()
    while curr != prev:
        prev = curr
        curr = pattern.sub("", curr)
    return curr


def _handle_plus_minus_symmetry(text: str) -> str:
    def _replace_match(match):
        left, op, right = match.group(1), match.group(2), match.group(3)
        if len(left) == len(right):
            if left.isdigit() and right.isdigit():
                return f"{left}{op}{right}"
            protected_op = "XXXXPLUSXXXX" if op == "+" else "XXXXMINUSXXXX"
            return f"{left}{protected_op}{right}"
        if re.search(r"[A-Za-z]", left) or re.search(r"[A-Za-z]", right):
            return f"{left}{right}"
        return f"{left}{op}{right}"

    pattern = re.compile(r"([A-Za-z0-9]+)\s*([\+\-])\s*([A-Za-z0-9]+)")
    prev_text = ""
    curr_text = text
    while curr_text != prev_text:
        prev_text = curr_text
        curr_text = pattern.sub(_replace_match, curr_text)
    return curr_text


def _rapikan_teks_label(text: str) -> str:
    text = str(text).upper()
    text = _normalisasi_as_per_list(text)
    text = Patterns.INFO_NOTE_PAREN.sub("", text)

    protect_var = bool(re.match(r"^\s*(TBA|VARIOUS|VAR)\b", text))
    teks_tanpa_tba = text if protect_var else Patterns.TBA_VAR.sub("", text)

    teks_tanpa_tba = Patterns.PLACEHOLDER_TOKENS.sub("", teks_tanpa_tba)
    teks_tanpa_tba = Patterns.STRIP_LABEL_CLEAN.sub("", teks_tanpa_tba).strip()

    if not teks_tanpa_tba:
        text = Patterns.STRIP_LABEL_CLEAN.sub("", text)
        return Patterns.MULTIPLE_SPACES.sub(" ", text).strip()

    text_final = text if protect_var else Patterns.TBA_VAR.sub("", text)
    text_final = Patterns.PLACEHOLDER_TOKENS.sub("", text_final)
    text_final = Patterns.STRIP_LABEL_CLEAN.sub("", text_final)
    return Patterns.MULTIPLE_SPACES.sub(" ", text_final).strip()


def _gabung_digit_terpisah(text: str) -> str:
    cleaned = text.strip()
    if re.fullmatch(r"[\d\s]+", cleaned):
        only_digits = re.sub(r"\s+", "", cleaned)
        if len(only_digits) >= 7:
            return only_digits
    return text


def _pre_clean_polis_or_slip(text: str) -> str:
    text = Patterns.PLACEHOLDER_TOKENS.sub("", text)
    text = Patterns.INFO_NOTE_PAREN.sub("", text)
    text = Patterns.TRAILING_NOTE.sub("", text)
    text = Patterns.MONTH_YEAR_PREFIX.sub("", text)
    text = _hapus_single_char_suffix(text)
    text = _gabung_digit_terpisah(text)

    pattern_dup = re.compile(r"\b([A-Za-z0-9]+)\s*[-/]\s*\1\b", re.IGNORECASE)
    prev = ""
    while prev != text:
        prev = text
        text = pattern_dup.sub(r"\1", text)

    return _handle_plus_minus_symmetry(text).strip()


def _pisah_jadi_segmen(text: str) -> List[str]:
    penanda_sd = "§SD§"
    teks_terlindungi = Patterns.SD.sub(penanda_sd, text)
    teks_terlindungi = re.sub(
        rf"/(?!{Patterns._KATA_LABEL_UNTUK_SLASH}\b)([A-Za-z]{{2,5}})\b",
        r"§SLASH§\1",
        teks_terlindungi,
        flags=re.IGNORECASE,
    )

    segmen_mentah = re.split(r"\s{2,}|/", teks_terlindungi)
    segmen_hasil = []
    for segmen in segmen_mentah:
        segmen = segmen.replace(penanda_sd, "SD").replace("§SLASH§", "/").strip()
        if segmen != "" and re.search(r"[A-Za-z0-9]", segmen):
            segmen_hasil.append(segmen)
    return segmen_hasil


def hapus_kurung(text: str) -> str:
    return Patterns.TARGET_PARENS.sub(" ", text)


def _rapikan_kurung_ganda(text: str) -> str:
    """Hapus tanda kurung buka yang berlebih, contoh: 'X ( (Y, Z)' -> 'X (Y, Z)'."""
    return Patterns.DOUBLE_OPEN_PAREN.sub("(", text)


def _mask_isi_kurung(text: str):
    """Lindungi isi di dalam tanda kurung (termasuk koma di dalamnya) agar tidak
    ikut menjadi pemisah saat breakdown, sekaligus tidak menghapus isinya."""
    masks = []

    def _replace(m):
        masks.append(m.group(0))
        return f"§PAREN{len(masks) - 1}§"

    masked = Patterns.PAREN_CONTENT.sub(_replace, text)
    return masked, masks


def _unmask_isi_kurung(text: str, masks: List[str]) -> str:
    for i, val in enumerate(masks):
        text = text.replace(f"§PAREN{i}§", val)
    return text


def clean_insured(name: Union[str, float]) -> Union[str, float]:
    if pd.isna(name):
        return name

    text = str(name).strip()
    text = _rapikan_kurung_ganda(text)
    text = Patterns.BORDERO_TAIL.sub("", text)
    text = hapus_kurung(text)

    text_masked, masks = _mask_isi_kurung(text)
    text_masked = Patterns.INSURED_PERSERO.sub(" ", text_masked)
    text_masked = Patterns.INSURED_ENTITIES.sub(" ", text_masked)
    text_masked = re.sub(r"\s*,\s*", " ", text_masked)
    text_masked = re.sub(r"\s*/\s*(?=[A-Za-z\s]+DIVISION)", " - ", text_masked, flags=re.IGNORECASE)
    # Bersihkan sisa tanda titik pemisah singkatan (mis. "DIV." -> "DIV") yang bukan bagian angka
    text_masked = re.sub(r"(?<=[A-Za-z])\.(?=\s|$)", "", text_masked)
    text_masked = Patterns.STRIP_INSURED_BOUNDARIES.sub("", text_masked)
    text_masked = Patterns.MULTIPLE_SPACES.sub(" ", text_masked).strip()
    return _unmask_isi_kurung(text_masked, masks)


def _finalize_insured_hasil(hasil: List[str]) -> List[str]:
    hasil = [h for h in hasil if h and h.strip()]
    if len(hasil) == 2 and Patterns.BANK_CODE_TOKEN.fullmatch(hasil[-1].strip()):
        gabungan = Patterns.MULTIPLE_SPACES.sub(" ", f"{hasil[0]} {hasil[-1]}").strip()
        return [gabungan]
    return hasil


def _handle_and_or_insured(text: str) -> Optional[List[str]]:
    """Breakdown insured berdasarkan pemisah 'AND/OR', mis."""
    parts = [p.strip() for p in Patterns.AND_OR_SEPARATOR.split(text) if p.strip()]
    if len(parts) < 2:
        return None

    hasil = [clean_insured(p) for p in parts]
    hasil = [h for h in hasil if h]
    if not hasil:
        return None
    if len(hasil) > MAX_HASIL_BREAKDOWN:
        return [text]
    return hasil


def _handle_fac_prefix_insured(text: str) -> Optional[List[str]]:
    m = Patterns.FAC_LEADING_PREFIX.match(text)
    if not m:
        return None

    remainder = text[m.end() :].strip()
    remainder = Patterns.LTD_AND_OR_SUBSIDIARIES.sub("", remainder).strip()

    if not remainder:
        cln_all = clean_insured(text)
        return [cln_all] if cln_all else [text]

    remainder_slash_parts = [p for p in remainder.split("/") if p.strip()]
    if len(remainder_slash_parts) > MAX_HASIL_BREAKDOWN:
        return [text]

    # Breakdown by QQ
    if re.search(r"\bQQ\b", remainder, re.IGNORECASE):
        qq_parts = [
            p.strip() for p in re.split(r"\bQQ\b", remainder, flags=re.IGNORECASE) if p.strip()
        ]
        hasil = [clean_insured(p) for p in qq_parts]
        hasil = [h for h in hasil if h]
        if hasil:
            return hasil

    # Breakdown by "/" pada sisa teks (setelah prefix dihapus), selama hasilnya
    # masih dalam batas MAX_HASIL_BREAKDOWN. Mis. "SURYA INDAH NUSANTARA PAGI/
    # PERKEBUNAN LEMBAH BHAKTI" -> 2 insured terpisah.
    if "/" in remainder:
        slash_parts = [p.strip() for p in remainder.split("/") if p.strip()]
        if 2 <= len(slash_parts) <= MAX_HASIL_BREAKDOWN:
            hasil = [clean_insured(p) for p in slash_parts]
            hasil = [h for h in hasil if h]
            if hasil:
                return hasil

    # Breakdown by koma, hanya bila segmen setelah koma diawali penanda
    # entitas yang menempel tanpa spasi ke nama berikutnya (indikasi 2 insured berbeda)
    if "," in remainder:
        koma_parts = [p.strip() for p in remainder.split(",")]
        if len(koma_parts) == 2 and Patterns.GLUED_ENTITY_MARKER.match(koma_parts[1]):
            insured2_raw = Patterns.GLUED_ENTITY_MARKER.sub("", koma_parts[1], count=1).strip()
            hasil = [clean_insured(koma_parts[0]), clean_insured(insured2_raw)]
            hasil = [h for h in hasil if h]
            if hasil:
                return hasil

    # Default: satu insured saja (bukan di-breakdown), rapikan penanda entitas biasa
    cln = clean_insured(remainder)
    return [cln] if cln else [text]


def breakdown_insured(value: Union[str, float]) -> List:
    if pd.isna(value):
        return [value]

    original_text = str(value).strip()
    if not original_text or Patterns.VARIOUS_EXACT.fullmatch(original_text):
        return [original_text]

    # Rule tambahan: override exact-match untuk kasus yang butuh pengetahuan
    # nama perusahaan asli (lihat Patterns.INSURED_EXACT_OVERRIDES)
    key_override = re.sub(r"\s+", " ", original_text.strip().upper())
    if key_override in Patterns.INSURED_EXACT_OVERRIDES:
        return list(Patterns.INSURED_EXACT_OVERRIDES[key_override])

    # Rule tambahan: hapus kata "AS CONTRACTOR" / "AS PRINCIPAL"
    original_text = Patterns.AS_CONTRACTOR_PRINCIPAL.sub("", original_text)
    original_text = Patterns.MULTIPLE_SPACES.sub(" ", original_text).strip()

    text = _rapikan_kurung_ganda(original_text)
    text = Patterns.BORDERO_TAIL.sub("", text).strip()

    if re.search(r"POLYCHEM", text, re.IGNORECASE):
        cln = clean_insured(text)
        return [cln] if cln else [original_text]

    # Rule tambahan: prefix "AUTO FACILITY/CARGO FACILITY/LINESLIP/" dkk.
    hasil_fac = _handle_fac_prefix_insured(text)
    if hasil_fac is not None:
        return _finalize_insured_hasil(hasil_fac)

    # Rule tambahan: breakdown by "AND/OR" (di luar pola "LTD AND/OR SUBSIDIARIES")
    if re.search(
        r"\bAND\s*/\s*OR\b", text, re.IGNORECASE
    ) and not Patterns.LTD_AND_OR_SUBSIDIARIES.search(text):
        hasil_and_or = _handle_and_or_insured(text)
        if hasil_and_or is not None:
            return _finalize_insured_hasil(hasil_and_or)

    text_no_parens = hapus_kurung(text)
    text_masked, masks = _mask_isi_kurung(text_no_parens)
    text_masked = Patterns.MASK_FAC.sub(r"\1_SLASH_", text_masked)

    SPLIT_RE = re.compile(r"\s*(?:/|,|\bQQ\b|\bOR\b)\s*|\s+-\s+", re.IGNORECASE)
    raw_parts = SPLIT_RE.split(text_masked)
    cleaned_parts = []

    for p in raw_parts:
        p = p.replace("_SLASH_", "/")
        p = _unmask_isi_kurung(p, masks)
        if Patterns.VARIOUS_EXACT.fullmatch(p.strip()):
            continue
        cln = clean_insured(p)
        if not cln:
            continue
        # Segmen yang isinya cuma "(anotasi)" (mis. hasil dari "NAMA,PT (ANOTASI)")
        # digabung ke entitas sebelumnya, bukan jadi hasil breakdown terpisah.
        if re.fullmatch(r"\(.*\)", cln) and cleaned_parts:
            cleaned_parts[-1] = f"{cleaned_parts[-1]} {cln}".strip()
        else:
            cleaned_parts.append(cln)

    if len(cleaned_parts) > MAX_HASIL_BREAKDOWN:
        return [original_text]

    if not cleaned_parts:
        cln_all = clean_insured(original_text)
        return [cln_all] if cln_all else [original_text]

    return _finalize_insured_hasil(cleaned_parts)


def _hapus_end_trailing(text: str) -> str:
    if not Patterns.END_OR_EN_PRESENT.search(text):
        return text

    stripped_full = text.strip()
    if Patterns.END_TOKEN.match(stripped_full):
        return text

    segmen_plus = re.split(r"(\+)", text)
    token_token = segmen_plus[0::2]

    idx_mulai_end = None
    for i, tok in enumerate(token_token):
        tok_bersih = tok.strip()
        if not tok_bersih:
            continue
        if Patterns.END_TOKEN.match(tok_bersih):
            idx_mulai_end = i
            break
        m = Patterns.END_TRAILING_MARK.search(tok_bersih)
        if m:
            token_token[i] = tok_bersih[: m.start()]
            idx_mulai_end = i + 1
            break

    if idx_mulai_end is None:
        return text

    token_token = token_token[:idx_mulai_end]
    hasil = "+".join(t for t in token_token if t.strip() != "")
    return hasil if hasil.strip() != "" else text


def _strip_bordero_vocab(text: str) -> str:
    probe = Patterns.BORDERO_DATE_RANGE.sub(" ", text)
    probe = Patterns.BORDERO_VOCAB.sub(" ", probe)
    probe = re.sub(r"\b\d{1,2}\b", " ", probe)
    probe = re.sub(r"\b\d{4}\b", " ", probe)
    return re.sub(r"[\s\-/.,]+", "", probe)


def _is_segmen_bordero_note(segmen: str) -> bool:
    return segmen.strip() == "" or _strip_bordero_vocab(segmen) == ""


def _handle_bordero_note(original_text: str):
    if not Patterns.BORDERO_KEYWORD_PRESENT.search(original_text):
        return None
    if _strip_bordero_vocab(original_text) == "":
        return [original_text]

    segmen_segmen = [s.strip() for s in re.split(r"[/\-]", original_text) if s.strip() != ""]
    if len(segmen_segmen) < 2:
        return None

    segmen_kode = [s for s in segmen_segmen if not _is_segmen_bordero_note(s)]
    segmen_note = [s for s in segmen_segmen if _is_segmen_bordero_note(s)]

    if not segmen_note or len(segmen_kode) != 1:
        return None

    return [segmen_kode[0].strip().upper()]


def _hapus_catatan_ekor_slip(text: str) -> str:
    hasil = text
    while True:
        dihapus = Patterns.SLIP_CATATAN_EKOR.sub("", hasil).strip()
        if dihapus == hasil:
            dihapus = Patterns.SLIP_MATA_UANG_EKOR.sub("", hasil).strip()
        if dihapus == "" or dihapus == hasil:
            break
        hasil = dihapus
    return hasil


def _gabungkan_prefix_dash_token(token: str) -> str:
    """'PMOCM-16067080102' -> 'PMOCM16067080102' agar tidak terpecah saat
    direkonstruksi bersama token digit murni lain di sebelahnya."""
    m = Patterns.PREFIX_DASH_DIGIT.match(token.strip())
    if m:
        return m.group(1) + m.group(2)
    return token.strip()


def _breakdown_slip_garis_miring_koma(text: str):
    if "/" not in text or "," not in text:
        return None
    if _mengandung_kata_sd(text) or _sepertinya_pola_range(text):
        return None

    top_parts = [p.strip() for p in text.split("/") if p.strip()]
    if len(top_parts) < 2:
        return None

    hasil = []
    acuan_terakhir = None
    for part in top_parts:
        if Patterns.STANDALONE_BULAN_TAHUN.match(part):
            continue
        if Patterns.VARIOUS_EXACT.fullmatch(part):
            continue

        if "," in part:
            sub_parts = [s.strip() for s in part.split(",") if s.strip()]
            if len(sub_parts) >= 2 and _mengandung_kode_panjang(sub_parts[0]):
                sub_parts_gabung = [_gabungkan_prefix_dash_token(sp) for sp in sub_parts]
                hasil_rekon = _rekonstruksi_nomor_dari_pengulangan(
                    ",".join(sub_parts_gabung), acuan_terakhir
                )
                if hasil_rekon:
                    hasil.extend(hasil_rekon)
                    acuan_terakhir = hasil_rekon[-1]
                    continue
            hasil.append(part.strip().upper())
        else:
            hasil.append(_gabungkan_prefix_dash_token(part).upper())

        if hasil:
            acuan_terakhir = hasil[-1]

    if not hasil or len(hasil) > MAX_HASIL_BREAKDOWN:
        return [text.strip()]

    return hasil


def _breakdown_slip_khusus(raw: str) -> List[str]:
    if Patterns.SLIP_KEEP_AS_IS.search(raw):
        return [_rapikan_teks_label(raw)]

    m_dot = re.match(r"^(\d{11})\.(\d{2,4})$", raw)
    if m_dot:
        base, suffix = m_dot.groups()
        return [base, base[: -len(suffix)] + suffix]

    m_plus = re.match(
        r"^([A-Z]{3,5})(\d{2}[A-Z0-9]{4,5})\s*([+&])\s*(\d{2}[A-Z0-9]{4,5})(?:\s*([-\s]\s*\d{4}))?$",
        raw,
        re.IGNORECASE,
    )
    if m_plus:
        prefix, code1, sep, code2, suffix = m_plus.groups()
        slip1 = f"{prefix}{code1}".upper()
        slip2 = f"{prefix}{code2}".upper()
        if suffix:
            suffix_clean = re.sub(r"\s+", "", suffix).upper()
            slip2 = f"{slip2}{suffix_clean}"
        return [slip1, slip2]

    if ";" in raw:
        return [p.strip().upper() for p in raw.split(";") if p.strip()]

    if "&" in raw:
        parts_mentah = [p.strip() for p in raw.split("&") if p.strip()]
        if parts_mentah:
            digit_pertama = re.sub(r"\D", "", parts_mentah[0])
            if len(digit_pertama) < MINIMAL_DIGIT_NOMOR_POLIS:
                return [raw.strip()]
        return [p.upper() for p in parts_mentah]

    return []


def _breakdown_slip_pengulangan_digit(original_text: str):
    text = original_text.strip()
    if not text or _mengandung_kata_sd(text):
        return None

    if Patterns.SLIP_UPLOADED_FILES.search(text):
        return [original_text]

    if Patterns.SLIP_KEEP_AS_IS.search(text) and not re.match(
        r"^\d{11}\s+\d{2,3}(?:\s+\d{2,3})+$", text
    ):
        return None

    text = Patterns.MONTH_YEAR_PREFIX.sub("", text).strip().upper()
    text = re.sub(r"\.+$", "", text).strip()

    for connector_re in Patterns.SLIP_PENGULANGAN_CONNECTORS:
        if not connector_re.search(text):
            continue

        raw_parts = [p.strip() for p in connector_re.split(text) if p.strip()]
        if len(raw_parts) < 2:
            continue

        m_first = Patterns.SLIP_FIRST_TOKEN.match(raw_parts[0])
        if not m_first:
            continue

        prefix, first_digits = m_first.groups()
        hasil = [prefix + first_digits]
        gagal = False

        for tok in raw_parts[1:]:
            if Patterns.SLIP_JUNK_TOKEN.match(tok):
                break
            if tok[0].isalpha() and Patterns.SLIP_FIRST_TOKEN.match(tok):
                hasil.append(tok)
                continue
            if not tok.isdigit():
                gagal = True
                break

            if len(tok) == len(first_digits):
                new_digits = tok
            elif len(tok) in SLIP_SUFFIX_ALLOWED_LENGTHS and len(tok) < len(first_digits):
                new_digits = first_digits[: -len(tok)] + tok
            else:
                gagal = True
                break

            hasil.append(prefix + new_digits)

        if gagal or len(hasil) > MAX_HASIL_BREAKDOWN:
            return [original_text]

        return hasil

    return None


def _breakdown_slip_kode_kompleks(original_text: str):
    if Patterns.SLIP_KEEP_AS_IS.search(original_text):
        return None

    text = original_text.strip()
    if not text or _mengandung_kata_sd(text):
        return None

    text = Patterns.SLIP_TAIL_VARIOUS.sub("", text).strip()
    text = Patterns.SLIP_TAIL_END_DOT.sub("", text).strip()
    text = Patterns.SLIP_TAIL_END_PLUS.sub("", text).strip()

    for connector_re in Patterns.SLIP_KOMPLEKS_CONNECTORS:
        if not connector_re.search(text):
            continue

        parts = [p.strip() for p in connector_re.split(text) if p.strip()]
        if len(parts) < 2 or len(parts) > MAX_HASIL_BREAKDOWN:
            continue

        # Buang segmen catatan/label yang tidak mengandung kode sama sekali,
        # contoh: "FACP05Q3FV-0701 / SUSPN" -> hanya "FACP05Q3FV-0701" dipakai.
        parts_berkode = [p for p in parts if _mengandung_kode_panjang(p)]
        if parts_berkode and len(parts_berkode) < len(parts):
            parts = parts_berkode

        # Jika hasil filter di atas cuma nyisain 1 bagian tapi bagian itu
        # sendiri masih berupa gabungan 2 (atau lebih) kode lengkap yang
        # dipisah "-" (mis. "PMOCM22115560101-PMOCM22119900101"), pecah lagi
        # supaya tidak ke-concat jadi satu string panjang tanpa pemisah.
        if len(parts) == 1 and "-" in parts[0]:
            sub_parts = [sp.strip() for sp in parts[0].split("-") if sp.strip()]
            if len(sub_parts) >= 2 and all(
                Patterns.SLIP_FIRST_SEGMENT_CODE.match(sp) for sp in sub_parts
            ):
                deduped = []
                for sp in sub_parts:
                    if sp.upper() not in [d.upper() for d in deduped]:
                        deduped.append(sp)
                parts = deduped

        if not Patterns.SLIP_FIRST_SEGMENT_CODE.match(parts[0]):
            continue

        if all(re.match(r"^[A-Za-z]", p) for p in parts):
            return [_clean_code(p) for p in parts]

        return [original_text]

    return None


def _breakdown_slip_titik_group(original_text: str):
    text = original_text.strip()
    if "." not in text:
        return None

    parts = [p.strip() for p in text.split("+") if p.strip()] if "+" in text else [text]
    hasil_total = []

    for part in parts:
        segs = [s.strip() for s in part.split(".")]
        if len(segs) < 2 or not all(s.isdigit() for s in segs):
            return None

        has_long_seg = any(len(s) >= 9 for s in segs)
        total_len = sum(len(s) for s in segs)

        if has_long_seg:
            hasil_part = _rekonstruksi_nomor_dari_pengulangan(".".join(segs))
            if not hasil_part:
                return [original_text]
            hasil_total.extend(hasil_part)
        elif 10 <= total_len <= 13:
            hasil_total.append("".join(segs))
        else:
            return [original_text]

    if not hasil_total or len(hasil_total) > MAX_HASIL_BREAKDOWN:
        return [original_text]

    return hasil_total


def _mengandung_kode_panjang(text: str) -> bool:
    text_clean = _hapus_kata_label(text)
    for m in re.finditer(rf"[A-Za-z0-9]{{{MINIMAL_PANJANG_KODE},}}", text_clean):
        if re.search(r"\d", m.group()):
            return True
    return False


def _mengandung_kata_sd(text: str) -> bool:
    return Patterns.SD.search(text) is not None


def _sepertinya_pola_range(text: str) -> bool:
    kode_kode = Patterns.RANGE_PATTERN.findall(text)
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


def _berisi_banyak_polis_bertag(text: str) -> bool:
    if "," not in text:
        return False
    parts = [p.strip() for p in text.split(",") if p.strip()]
    if len(parts) < 2:
        return False
    return all(Patterns.TAGGED_POLIS.match(p) for p in parts)


def _rekonstruksi_nomor_dari_pengulangan(
    teks_tanpa_label: str, acuan_awal: str = None
) -> List[str]:
    raw_tokens = [t.strip() for t in re.split(r"[+,\-./]", teks_tanpa_label) if t.strip() != ""]
    tokens_expanded = [_clean_code(t) for t in raw_tokens if _clean_code(t)]

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
            prefix_alpha = (
                re.match(r"^[A-Za-z]+", nomor_acuan_terakhir).group()
                if nomor_acuan_terakhir and re.match(r"^[A-Za-z]+", nomor_acuan_terakhir)
                else ""
            )
            digit_acuan = (
                re.sub(r"^[A-Za-z]+", "", nomor_acuan_terakhir) if nomor_acuan_terakhir else ""
            )
            if prefix_alpha and len(tok) == len(digit_acuan):
                tok_lengkap = prefix_alpha + tok
                hasil.append(tok_lengkap)
                nomor_acuan_terakhir = tok_lengkap
            else:
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


def _proses_satu_segmen(text: str, wajib_11_digit: bool, acuan_awal: str = None) -> List[str]:
    if not _mengandung_kode_panjang(text):
        return [_rapikan_teks_label(text)]

    if _mengandung_kata_sd(text) or _sepertinya_pola_range(text):
        return [re.sub(r"\s+", " ", text).strip()]

    if _berisi_banyak_polis_bertag(text):
        hasil = [p.strip().upper() for p in text.split(",") if p.strip()]
        return hasil if len(hasil) <= MAX_HASIL_BREAKDOWN else [text.strip()]

    if _berisi_banyak_nomor_utuh(text):
        hasil = [_clean_code(t) for t in text.split()]
        return hasil if len(hasil) <= MAX_HASIL_BREAKDOWN else [text.strip()]

    if wajib_11_digit:
        has_valid_polis_format = bool(Patterns.POLIS_VALID_FORMAT.search(text)) or bool(
            Patterns.POLIS_SUFFIX_FORMAT.search(text)
        )
        if not has_valid_polis_format:
            kode_pertama = re.split(r"[+,\-]", text)[0]
            if len(re.sub(r"\D", "", kode_pertama)) < MINIMAL_DIGIT_NOMOR_POLIS:
                return [text.strip()]

    teks_tanpa_placeholder = Patterns.PLACEHOLDER_TOKENS.sub("", text)
    hasil_breakdown = _rekonstruksi_nomor_dari_pengulangan(teks_tanpa_placeholder, acuan_awal)

    if not hasil_breakdown or len(hasil_breakdown) > MAX_HASIL_BREAKDOWN:
        return [text.strip()]

    return hasil_breakdown


def _chunk_digit_string(digit_str: str, preferred_len: int = None):
    kandidat_panjang = []
    if preferred_len and len(digit_str) % preferred_len == 0:
        kandidat_panjang.append(preferred_len)
    for panjang in VALID_FULL_LENGTHS_POLIS:
        if panjang not in kandidat_panjang and len(digit_str) % panjang == 0:
            kandidat_panjang.append(panjang)

    for panjang in kandidat_panjang:
        jumlah_potongan = len(digit_str) // panjang
        if jumlah_potongan >= 1:
            return [digit_str[i * panjang : (i + 1) * panjang] for i in range(jumlah_potongan)]
    return None


def _breakdown_ampersand_polis(original_text: str):
    if "&" not in original_text:
        return None

    parts = [p.strip() for p in original_text.split("&")]
    if len(parts) < 2:
        return None

    digit_parts = []
    for p in parts:
        if not p or not p.replace(" ", "").isdigit():
            return None
        digit_parts.append(re.sub(r"\D", "", p))

    panjang_panjang = [len(d) for d in digit_parts]
    if any(panjang < MINIMAL_DIGIT_NOMOR_POLIS for panjang in panjang_panjang):
        return [original_text]

    panjang_valid = [p for p in panjang_panjang if p in VALID_FULL_LENGTHS_POLIS]
    preferred_len = panjang_valid[0] if panjang_valid else None

    hasil = []
    for d in digit_parts:
        if len(d) in VALID_FULL_LENGTHS_POLIS:
            hasil.append(d)
            continue
        potongan = _chunk_digit_string(d, preferred_len)
        if not potongan:
            return [original_text]
        hasil.extend(potongan)

    if len(hasil) > MAX_HASIL_BREAKDOWN:
        return [original_text]

    return hasil


def _cek_pisah_spasi_nomor_utuh(text: str):
    groups = [g for g in text.split() if g]
    if len(groups) < 2:
        return None

    while groups and Patterns.TOKEN_SAMPAH_SPASI.match(groups[-1]):
        groups.pop()

    if len(groups) < 2 or not all(re.fullmatch(r"\d+", g) for g in groups):
        return None

    if (
        not all(len(g) >= MINIMAL_DIGIT_NOMOR_POLIS for g in groups)
        or len(groups) > MAX_HASIL_BREAKDOWN
    ):
        return None

    return groups


def _cek_slash_tanpa_digit_lengkap(text: str, wajib_11_digit: bool):
    if (
        not wajib_11_digit
        or "/" not in text
        or _mengandung_kata_sd(text)
        or _sepertinya_pola_range(text)
    ):
        return None

    segmen_segmen = [s.strip() for s in text.split("/") if s.strip() != ""]
    if len(segmen_segmen) < 2:
        return None

    for segmen in segmen_segmen:
        if len(re.sub(r"\D", "", segmen)) >= MINIMAL_DIGIT_NOMOR_POLIS:
            return None

    return [text.strip()]


def _cek_various_tanpa_nomor_asli(original_text: str):
    m = re.match(r"^\s*(?:TBA|VARIOUS|VAR)\b", original_text, re.IGNORECASE)
    if not m:
        return None

    sisa = original_text[m.end() :].strip()
    if sisa == "":
        return None

    digit_runs = re.findall(r"\d+", sisa)
    if any(len(d) >= MINIMAL_DIGIT_NOMOR_POLIS for d in digit_runs):
        return None

    return [original_text]


def _cek_awalan_pendek_polis(text: str):
    if not re.search(r"[+,\-]", text):
        return None
    parts = re.split(r"[+,\-]", text)
    if not parts:
        return None

    digit_awal = re.sub(r"\D", "", parts[0])
    if 0 < len(digit_awal) < MINIMAL_DIGIT_NOMOR_POLIS:
        # Kecuali: awalan pendek ini bisa digabung dengan segmen berikutnya
        # (setelah tanda - pertama) membentuk format polis penuh yang valid,
        # mis. "012400131-185+187+210+189" -> "012400131" (9 digit) + "185"
        # (3 digit) = 12 digit (valid). Dalam kasus ini JANGAN bailout,
        # biarkan diproses oleh mesin rekonstruksi nomor.
        if len(parts) > 1:
            digit_next = re.sub(r"\D", "", parts[1])
            if digit_next and (len(digit_awal) + len(digit_next)) in VALID_FULL_LENGTHS_POLIS:
                return None
        return [text.strip()]
    return None


def _handle_pola_sd_tambahan(original_text: str, wajib_11_digit: bool) -> Optional[List[str]]:
    text = original_text.strip()
    
    # Hapus awalan seperti "JUNE 2021 -", "AGUSTUS 2021/IDR -", dll
    prefix_pattern = rf"^(?:{Patterns._BULAN_ALT}\s+\d{{4}}\s*[/\-]?\s*(?:IDR|USD|JPY|SGD|EUR|GBP)?\s*[/\-]?\s*)+"
    text = re.sub(prefix_pattern, "", text, flags=re.IGNORECASE).strip()
    
    # Hapus juga awalan MONTH_YEAR_PREFIX standar jika masih ada sisa
    text = Patterns.MONTH_YEAR_PREFIX.sub("", text).strip()
    text = re.sub(r"\s*S\s*/\s*D\s*", " SD ", text, flags=re.IGNORECASE).strip()
    
    if wajib_11_digit:
        parts = text.split(" SD ")
        if len(parts) >= 2:
            left_part = parts[0]
            if len(left_part) >= 23:
                rest = " SD ".join(parts[1:])
                return [left_part[:12].upper(), (left_part[12:] + " SD " + rest).upper()]
    else:
        parts = text.split(" SD ")
        if len(parts) >= 2:
            left_part = parts[0]
            if len(left_part) >= 31:
                rest = " SD ".join(parts[1:])
                return [left_part[:16].upper(), (left_part[16:] + " SD " + rest).upper()]
                
    return [text.upper()]


def _breakdown_polis_atau_slip_internal(value: Union[str, float], wajib_11_digit: bool) -> List:
    if pd.isna(value):
        return [value]

    original_text = str(value).strip()

    # --- TAMBAHAN BARU: Bypass pembersihan jika slip mengandung ASTRA EX LS ---
    if not wajib_11_digit and Patterns.ASTRA_EX_LS_PATTERN.search(original_text):
        return [original_text]
    # --------------------------------------------------------------------------

    # Untuk slip: buang ekor catatan "/VARIOUS/<BULAN TAHUN>/<CCY>" (atau
    # sebagian darinya) SEBELUM cek pola S/D, supaya "S" di akhir kata
    # "VARIOUS" dan "D" di awal kata "DECEMBER" dst tidak salah kebaca
    # sebagai penanda S/D.
    if not wajib_11_digit:
        original_text = Patterns.SLIP_TAIL_VARIOUS.sub("", original_text).strip()

    if Patterns.SD.search(original_text):
        return _handle_pola_sd_tambahan(original_text, wajib_11_digit)

    if (
        original_text == ""
        or original_text.startswith("*")
        or Patterns.END_TOKEN.match(original_text)
    ):
        return [original_text]

    if Patterns.P_PAREN_P_KEEP_AS_IS.match(original_text):
        return [original_text.strip()]

    if wajib_11_digit and Patterns.POLIS_KEEP_AS_IS.match(original_text.upper()):
        return [original_text.upper()]

    if not wajib_11_digit:
        _normalized_check = re.sub(r"\s+", " ", original_text.strip().upper())
        if _normalized_check in Patterns.SLIP_KEEP_AS_IS_EXACT:
            return [original_text.strip()]

        hasil_bordero_note = _handle_bordero_note(original_text)
        if hasil_bordero_note is not None:
            return hasil_bordero_note

    original_text = _hapus_end_trailing(original_text).strip()
    if original_text == "":
        return [str(value).strip()]

    _teks_tanpa_bulan_ekor = Patterns.MONTH_YEAR_SUFFIX.sub("", original_text).strip()
    if _teks_tanpa_bulan_ekor != "":
        original_text = _teks_tanpa_bulan_ekor

    if not wajib_11_digit:
        _teks_tanpa_catatan_ekor = _hapus_catatan_ekor_slip(original_text)
        if _teks_tanpa_catatan_ekor != "":
            original_text = _teks_tanpa_catatan_ekor

        # Hapus label "BULAN TAHUN - CCY/" berulang (mis. label bulan/tahun/
        # currency yang ke-copy 2x sebelum kode slip aslinya) secara loop,
        # bukan cuma sekali, supaya semua pengulangannya ikut terhapus.
        _teks_iter = original_text
        while True:
            _next_iter = Patterns.MONTH_YEAR_CCY_SLASH_PREFIX.sub("", _teks_iter).strip()
            if _next_iter == _teks_iter:
                _next_iter = Patterns.MONTH_YEAR_PREFIX.sub("", _teks_iter).strip()
            if _next_iter == _teks_iter or _next_iter == "":
                break
            _teks_iter = _next_iter
        _teks_tanpa_prefix_bulan = _teks_iter
        if (
            _teks_tanpa_prefix_bulan != ""
            and _teks_tanpa_prefix_bulan != original_text
            and _mengandung_kode_panjang(_teks_tanpa_prefix_bulan)
        ):
            original_text = _teks_tanpa_prefix_bulan

    if re.fullmatch(r"\d{3}(?:-\d{3}){2,}", original_text.replace(" ", "")):
        return [original_text]

    hasil_various_tanpa_nomor = _cek_various_tanpa_nomor_asli(original_text)
    if hasil_various_tanpa_nomor is not None:
        return hasil_various_tanpa_nomor

    if wajib_11_digit:
        hasil_awalan_pendek = _cek_awalan_pendek_polis(original_text)
        if hasil_awalan_pendek is not None:
            return hasil_awalan_pendek

        hasil_slash_pendek = _cek_slash_tanpa_digit_lengkap(original_text, wajib_11_digit)
        if hasil_slash_pendek is not None:
            return hasil_slash_pendek

        hasil_ampersand = _breakdown_ampersand_polis(original_text)
        if hasil_ampersand is not None:
            return hasil_ampersand

        hasil_spasi_utuh = _cek_pisah_spasi_nomor_utuh(original_text)
        if hasil_spasi_utuh is not None:
            return hasil_spasi_utuh

    if not wajib_11_digit:
        _probe = Patterns.TBA_VAR.sub("", original_text.upper())
        _probe = Patterns.PLACEHOLDER_TOKENS.sub("", _probe)
        _probe = re.sub(r"[+\-/,.\s]", "", _probe)
        if _probe == "":
            return [original_text]

        _tanpa_bulan_probe = Patterns.MONTH_YEAR_PREFIX.sub("", original_text.upper()).strip()
        if _tanpa_bulan_probe != original_text.upper().strip() and not _mengandung_kode_panjang(
            _tanpa_bulan_probe
        ):
            return [original_text]

        hasil_garis_koma = _breakdown_slip_garis_miring_koma(original_text)
        if hasil_garis_koma is not None:
            return hasil_garis_koma

        hasil_pengulangan = _breakdown_slip_pengulangan_digit(original_text)
        if hasil_pengulangan is not None:
            return hasil_pengulangan

        hasil_kompleks = _breakdown_slip_kode_kompleks(original_text)
        if hasil_kompleks is not None:
            return hasil_kompleks

        hasil_titik = _breakdown_slip_titik_group(original_text)
        if hasil_titik is not None:
            return hasil_titik

    seps = set(re.findall(r"(?<=\d)\s*([.,/+\-])\s*(?=\d)", original_text))
    if len(seps) > 1:
        # Kecuali: teks murni angka + tanda pemisah (tanpa huruf) pada kolom
        # polis (wajib_11_digit), mis. "031900007968-4/031900007971-4" atau
        # "012400125843+012400125843-1". Pola semacam ini tetap aman
        # diproses oleh mesin pembersihan/rekonstruksi di bawah; yang perlu
        # di-bailout hanya kode alfanumerik kompleks yang ambigu.
        murni_angka_dan_pemisah = bool(re.fullmatch(r"[\d\s.,/+\-]+", original_text))
        if not (wajib_11_digit and murni_angka_dan_pemisah):
            return [original_text]

    raw = original_text.upper()
    if re.fullmatch(r"(TBA|VARIOUS|VAR|P1|P2|P3)", raw.strip()):
        return [raw.strip()]

    raw = re.sub(r"\s*END\s*$", "", raw).strip()
    raw = re.sub(r"\s*[+\-]\s*$", "", raw).strip()
    raw = Patterns.TBA_VAR.sub("", raw).strip()

    parts_awal = [p for p in re.split(r"[+,\-./]", raw) if p.strip()]
    if len(parts_awal) > MAX_HASIL_BREAKDOWN:
        return [original_text]

    if wajib_11_digit:
        # Hanya suffix 1 digit setelah titik di akhir yang DIHAPUS total
        # (mis. "012300112129.3" -> "012300112129"). Suffix 2+ digit
        # (mis. "051800213144.10") harus di-breakdown (lihat mesin
        # rekonstruksi di bawah), bukan dihapus begitu saja.
        m_dot_tunggal = re.match(r"^(\d{11,16})\.(\d{1})$", raw)
        if m_dot_tunggal:
            return [m_dot_tunggal.group(1)]

    is_slip_dot_suffix_pattern = (not wajib_11_digit) and bool(re.match(r"^\d{11}\.\d{2,4}$", raw))
    is_polis_dot_suffix_pattern = wajib_11_digit and bool(re.match(r"^\d{11,16}\.\d{1,4}$", raw))
    if (
        Patterns.DOT_GROUPED_CODE.fullmatch(raw)
        and not is_slip_dot_suffix_pattern
        and not is_polis_dot_suffix_pattern
    ):
        return [raw.replace(".", "")]

    if re.fullmatch(r"\d{2,3}-\d{2,3}-\d{2,3}-\d{4,5}", raw):
        return [raw.replace("-", "")]

    raw = _pre_clean_polis_or_slip(raw)
    parts = [p for p in re.split(r"[+,\-./]", raw) if p.strip()]
    if len(parts) > MAX_HASIL_BREAKDOWN:
        return [raw]

    if not wajib_11_digit:
        hasil_khusus = _breakdown_slip_khusus(raw)
        if hasil_khusus:
            return [raw] if len(hasil_khusus) > MAX_HASIL_BREAKDOWN else hasil_khusus

    if Patterns.MASTER_POLICY_ANNOTATION.search(raw) or not _mengandung_kode_panjang(raw):
        return [_rapikan_teks_label(raw)]

    text = _hapus_kata_label(raw)
    if text.strip() == "":
        return [_rapikan_teks_label(raw)]

    segmen_segmen = _pisah_jadi_segmen(text)
    if not segmen_segmen:
        return [_rapikan_teks_label(raw)]

    ada_segmen_berkode = any(_mengandung_kode_panjang(s) for s in segmen_segmen)
    semua_hasil = []
    acuan_terakhir = None

    for segmen in segmen_segmen:
        berkode = _mengandung_kode_panjang(segmen)
        if ada_segmen_berkode and not berkode:
            if (
                re.fullmatch(r"\d+", segmen)
                and acuan_terakhir
                and len(acuan_terakhir) > len(segmen)
            ):
                hasil_segmen = [acuan_terakhir[: -len(segmen)] + segmen]
            else:
                continue
        else:
            hasil_segmen = _proses_satu_segmen(segmen, wajib_11_digit, acuan_terakhir)

        semua_hasil.extend(hasil_segmen)
        if hasil_segmen and re.fullmatch(r"[A-Za-z0-9]+", hasil_segmen[-1]):
            acuan_terakhir = hasil_segmen[-1]

    hasil_final = []
    for item in semua_hasil:
        item_clean = item.rstrip("&").strip()
        if item_clean:
            hasil_final.append(item_clean)

    if len(hasil_final) > MAX_HASIL_BREAKDOWN:
        return [raw]

    return hasil_final if hasil_final else [_rapikan_teks_label(raw)]


def breakdown_polis_atau_slip(value: Union[str, float], wajib_11_digit: bool) -> List:
    hasil_internal = _breakdown_polis_atau_slip_internal(value, wajib_11_digit)
    if not isinstance(hasil_internal, list):
        hasil_internal = [hasil_internal]

    hasil_final = []
    seen_normalized = set()
    for item in hasil_internal:
        if isinstance(item, str):
            item = item.replace("XXXXPLUSXXXX", "+").replace("XXXXMINUSXXXX", "-")
            item_clean = Patterns.MULTIPLE_SPACES.sub(" ", item.strip()).upper()
            if item_clean and item_clean not in seen_normalized:
                hasil_final.append(item_clean)
                seen_normalized.add(item_clean)
        else:
            if item not in hasil_final:
                hasil_final.append(item)

    return hasil_final


def clean_polis(value: Union[str, float]):
    return value if pd.isna(value) else _rapikan_teks_label(str(value))


def breakdown_polis(value: Union[str, float]) -> List:
    return breakdown_polis_atau_slip(value, wajib_11_digit=True)


def clean_slip(value: Union[str, float]):
    if pd.isna(value):
        return value
    original_text = str(value).strip()
    hasil_bordero_note = _handle_bordero_note(original_text)
    if hasil_bordero_note is not None:
        return hasil_bordero_note[0]
    return _rapikan_teks_label(original_text)


def breakdown_slip(value: Union[str, float]) -> List:
    return breakdown_polis_atau_slip(value, wajib_11_digit=False)


def _insert_breakdown_columns(
    df: pd.DataFrame, source_col: str, fn_breakdown, prefix: str, mask=None
) -> None:
    if mask is None:
        hasil_breakdown = df[source_col].apply(fn_breakdown)
    else:
        hasil_breakdown = pd.Series([[] for _ in range(len(df))], index=df.index, dtype=object)
        hasil_breakdown.loc[mask] = df.loc[mask, source_col].apply(fn_breakdown)

    panjang_terpanjang = int(hasil_breakdown.apply(len).max() or 0)
    max_cols = min(panjang_terpanjang, MAX_HASIL_BREAKDOWN) if panjang_terpanjang > 0 else 1

    hasil_breakdown_trimmed = hasil_breakdown.apply(
        lambda x: x[:max_cols] if isinstance(x, list) else x
    )

    df_breakdown = pd.DataFrame(
        hasil_breakdown_trimmed.tolist(),
        columns=[f"{prefix}_{i + 1}" for i in range(max_cols)],
        index=df.index,
    )

    posisi = df.columns.get_loc(source_col)
    for i, col in enumerate(df_breakdown.columns):
        df.insert(posisi + 1 + i, col, df_breakdown[col])


def process_data(input_file: str, sheet_name: str, output_file: str) -> None:
    if not os.path.exists(input_file):
        print(f"[ERROR] File input tidak ditemukan: {input_file}")
        return

    print(f"Membaca data dari: {input_file} (Sheet: {sheet_name}) ...")
    df = pd.read_excel(input_file, sheet_name=sheet_name, header=0)
    df.columns = df.columns.str.strip()

    if CEDANT_COL not in df.columns:
        print(f"[ERROR] Kolom '{CEDANT_COL}' tidak ditemukan di file!")
        return

    df = df[df[CEDANT_COL].astype(str).str.strip() == CEDANT_VALUE].copy()
    if df.empty:
        print("[WARN] Data kosong setelah filter. Proses dihentikan.")
        return

    # --- RENAME COMP_NAME.1 -> COMP_NAME2 ---
    if "COMP_NAME.1" in df.columns:
        df.rename(columns={"COMP_NAME.1": "COMP_NAME2"}, inplace=True)

    # --- PENAMBAHAN KOLOM BUSSINESS_PATNER ---
    if "COMP_NAME2" in df.columns:
        print("Menambahkan kolom BUSSINESS_PATNER ...")
        bp_series = df.apply(
            lambda row: get_business_partners(row.get("COMP_NAME2"), row.get(CEDANT_COL)), axis=1
        )
        pos_comp2 = df.columns.get_loc("COMP_NAME2")
        df.insert(pos_comp2 + 1, "BUSSINESS_PATNER", bp_series)
    else:
        print("[WARN] Kolom COMP_NAME2 tidak ditemukan. Kolom BUSSINESS_PATNER dilewati.")

    print("Menjalankan pembersihan dan breakdown kolom...")
    column_mappings = [
        ("FAC_INSURED", breakdown_insured, "INSURED_CLEAN"),
        ("FAC_POLICY_NO", breakdown_polis, "POLICY_CLEAN"),
        ("FAC_SLIP", breakdown_slip, "SLIP_CLEAN"),
    ]

    for source_col, fn, prefix in column_mappings:
        if source_col in df.columns:
            _insert_breakdown_columns(df, source_col, fn, prefix)
        else:
            print(f"[WARN] Kolom {source_col} tidak ditemukan, dilewati.")

    # --- PENAMBAHAN KOLOM SERTIFIKAT (Rule 1: sertif selalu sehabis polis) ---
    for i in range(1, 4):
        sertif_col = f"SERTIF_CLEAN_{i}"
        pol_col = f"POLICY_CLEAN_{i}"
        if pol_col in df.columns:
            pos = df.columns.get_loc(pol_col)
            df.insert(pos + 1, sertif_col, "")
        else:
            df[sertif_col] = ""

    # Standarisasi urutan kolom FACUL CLEAN
    FACUL_TARGET_COLS = [
        "FAC_CODE", "FAC_CEDANT", "COMP_NAME", "FAC_BROKER", "COMP_NAME2", "BUSSINESS_PATNER",
        "FAC_INSURED", "INSURED_CLEAN_1", "INSURED_CLEAN_2", "INSURED_CLEAN_3", "INSURED_CLEAN_4", "INSURED_CLEAN_5",
        "FAC_POLICY_NO", "POLICY_CLEAN_1", "SERTIF_CLEAN_1", "POLICY_CLEAN_2", "SERTIF_CLEAN_2", "POLICY_CLEAN_3", "SERTIF_CLEAN_3",
        "POLICY_CLEAN_4", "POLICY_CLEAN_5",
        "FAC_SLIP", "SLIP_CLEAN_1", "SLIP_CLEAN_2", "SLIP_CLEAN_3", "SLIP_CLEAN_4", "SLIP_CLEAN_5",
        "FAC_CURRENCY", "FAC_INP_DATE", "FAC_COM_DATE", "FAC_EXP_DATE", "FAC_SUB_CLASS", "FAC_RISK", "FAC_DESC", "FAC_ACC_STS", "FAC_STS_SLIP"
    ]
    for c in FACUL_TARGET_COLS:
        if c not in df.columns:
            df[c] = ""
    # Susun kolom sesuai urutan target, sertakan sisa kolom lainnya di akhir jika ada
    extra_cols = [c for c in df.columns if c not in FACUL_TARGET_COLS]
    df = df[FACUL_TARGET_COLS + extra_cols]

    output_dir = os.path.dirname(output_file)
    if output_dir and not os.path.exists(output_dir):
        os.makedirs(output_dir)

    # Pastikan pada kolom polis clean dan slip clean semua S/D diganti menjadi SD
    clean_cols_to_check = [
        c for c in df.columns 
        if any(k in c.lower() for k in ["polis", "policy", "slip", "sertif"]) and "clean" in c.lower()
    ]
    for c in clean_cols_to_check:
        df[c] = df[c].map(
            lambda x: re.sub(r"\s*S\s*/\s*D\s*", " SD ", str(x), flags=re.IGNORECASE).strip()
            if pd.notna(x) and re.search(r"S\s*/\s*D", str(x), re.IGNORECASE)
            else x
        )

    print(f"Menyimpan hasil ke Excel: {output_file} ...")
    try:
        df.to_excel(output_file, index=False)
        print("Pemrosesan data selesai!")
        print(f"File disimpan di : {output_file}")
    except PermissionError:
        print(
            f"[ERROR] Tidak dapat menyimpan file. Pastikan file {output_file} tidak sedang dibuka di Excel!"
        )


if __name__ == "__main__":
    process_data(INPUT_FILE, SHEET_NAME, OUTPUT_FILE)