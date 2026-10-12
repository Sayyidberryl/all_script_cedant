import os
import re
from typing import Any, List
import numpy as np
import pandas as pd

INPUT_FILE = os.path.join("data_excel", "raw", "1b. Transaksi Facul 01.01.23 - 17.08.2026.xlsx")
SHEET_NAME = "Query result"
OUTPUT_FILE = os.path.join("data_excel", "processed", "mitrautama_output_facul.xlsx")

CEDANT_COL = "COMP_NAME"
CEDANT_VALUE = "PT. MITRA UTAMA REASURANSI"

INSURED_COL = "FAC_INSURED"
POLICY_COL = "FAC_POLICY_NO"
SLIP_COL = "FAC_SLIP"

MAX_HASIL_BREAKDOWN = 5


# POLIS
def is_exception(raw: str) -> bool:
    norm = re.sub(r"\s+", " ", str(raw).strip().upper())
    if "TBA / GEGI + P2 + P3" in norm:
        return True
    if "TBA NEW" in norm:
        return True
    if "TBA / SUNDAY+P2I" in norm:
        return True
    if norm == "TBA":
        return True
    if norm == "P1":
        return True
    return False


def clean_polis(value: Any) -> Any:
    if pd.isna(value):
        return value

    raw = str(value).strip()
    if not raw:
        return raw

    if is_exception(raw):
        return raw

    if raw.upper() == "RBA":
        return raw

    cleaned = raw

    temp = re.sub(r"\b(TBA|RBA|PT|TBK|0|P1)\b", "", cleaned, flags=re.IGNORECASE)
    temp = re.sub(r"[.\/\-\"]", "", temp).strip()

    if temp:
        cleaned = re.sub(r"\b(TBA|RBA|P1)\s*/?\s*", "", cleaned, flags=re.IGNORECASE)

    cleaned = re.sub(r"\s*\([^)]*\)\s*", " ", cleaned).strip()

    policy_pattern = r"\d{4,5}[/A-Z0-9.\-]+\d{2,4}"
    cleaned = re.sub(r"(" + policy_pattern + r")\s*SISA OSBAL", r"\1", cleaned, flags=re.IGNORECASE)
    cleaned = re.sub(r"(" + policy_pattern + r")\s*/\s*[A-Za-z\s]+$", r"\1", cleaned, flags=re.IGNORECASE)

    company_keywords = [
        "VICTORIA INSURANCE", "FPG", "ASURANSI", "MPM", "INDONESIA", "MAG",
        "UMUM MEGA", "ZURICH", "GREAT EASTERN", "GENERAL INSURANCE INDONESIA",
        "ACA", "SUNDAY INSURANCE INDONESIA,PT"
    ]
    company_keywords.sort(key=len, reverse=True)
    
    for comp in company_keywords:
        comp_pattern = re.escape(comp) + r"\s*(?:/\s*)?(?=\d{4,5}[/A-Z0-9.\-])"
        cleaned = re.sub(comp_pattern, "", cleaned, flags=re.IGNORECASE)

    cleaned = re.sub(r"^[A-Za-z\s,]+/\s*(?=\d{4,5}[/A-Z0-9.\-])", "", cleaned)
    cleaned = re.sub(r"[.\/\-\"]", "", cleaned)
    cleaned = re.sub(r"\b(PT|Tbk|0)\b", "", cleaned, flags=re.IGNORECASE)
    cleaned = re.sub(r"\s{2,}", " ", cleaned).strip()

    return cleaned


def breakdown_polis(value: Any) -> List[str]:
    if pd.isna(value):
        return [value]

    raw = str(value).strip()
    if not raw:
        return [raw]

    if is_exception(raw):
        return [raw]
        
    if raw.upper() == "RBA":
        return [raw]

    text = raw
    text = re.sub(r",\s*(PT|TBK)\b\.?", "", text, flags=re.IGNORECASE)

    match = re.match(r"^(.+?)(?:\s{2,}\1)+$", text, flags=re.IGNORECASE)
    if match:
        text = match.group(1)

    parts = []

    if "+" in text and re.search(r"\d{4,5}[/A-Z0-9.\-]+\s*\+\s*\d{4,5}", text):
        raw_parts = [p.strip() for p in text.split("+")]
        base_policy = raw_parts[0]
        match_suffix = re.search(r"^\d{4,5}(.*)", base_policy)
        suffix = match_suffix.group(1) if match_suffix else ""

        for i, p in enumerate(raw_parts):
            if i == 0:
                parts.append(base_policy)
            else:
                if re.match(r"^\d{4,5}$", p):
                    parts.append(p + suffix)
                else:
                    parts.append(p)
    elif "," in text:
        parts = [p.strip() for p in text.split(",") if p.strip()]
    else:
        parts = [text]

    cleaned_parts = []
    for p in parts:
        c = clean_polis(p)
        if c:
            cleaned_parts.append(c)

    if not cleaned_parts:
        return [raw]

    cleaned_parts = list(dict.fromkeys(cleaned_parts))

    if len(cleaned_parts) > 1:
        cleaned_parts = [x for x in cleaned_parts if x.upper() not in {"TBA", "RBA"}]

    if len(cleaned_parts) > MAX_HASIL_BREAKDOWN:
        return [raw]

    return cleaned_parts


# CERTIFICATE
# Untuk MITRAUTAMA DATA 1 (Facul), setelah dicek manual, tidak ada pola
# FAC_POLICY_NO yang benar-benar memenuhi kriteria certificate (2 digit
# paling akhir setelah "-"/"/" seringkali ternyata bagian lain dari nomor
# polis, mis. "MM/YY", bukan certificate sungguhan). Karena itu kolom
# CERTIFICATE tetap dibuat di posisinya, tapi untuk saat ini selalu
# dikosongkan. Kalau nanti ditemukan pola yang benar-benar valid, logika
# ekstraksinya tinggal ditambahkan di sini.
def extract_certificate(value: Any) -> Any:
    return np.nan


# SLIP
STANDALONE_SLIP_LABELS = {"TBA", "EMPTY"}
SLIP_LABEL_RE = re.compile(r"\bTBA\b|\bEMPTY\b", re.IGNORECASE)
P_CODE_RE = re.compile(r"^P\d+[A-Za-z]*(?:\(P\d+\))?$", re.IGNORECASE)
# Pola untuk mendeteksi rentang tanggal seperti 30/11/2023-30/11/2024
DATE_RANGE_SLIP_RE = re.compile(r"^\d{1,2}/\d{1,2}/\d{2,4}\s*-\s*\d{1,2}/\d{1,2}/\d{2,4}$")


def _is_placeholder_p_code(segment: str) -> bool:
    return bool(P_CODE_RE.match(segment.strip()))


def _is_date_range_slip(raw: str) -> bool:
    return bool(DATE_RANGE_SLIP_RE.match(raw.strip()))


def clean_slip(value: Any) -> Any:
    if pd.isna(value):
        return value

    raw = str(value).strip()
    if not raw:
        return raw
        
    if raw.upper() in STANDALONE_SLIP_LABELS:
        return raw
        
    # Exception untuk format rentang tanggal
    if _is_date_range_slip(raw):
        return raw

    cleaned = SLIP_LABEL_RE.sub("", raw)
    cleaned = re.sub(r"[.\/\-,]", "", cleaned)
    cleaned = re.sub(r"^[\s+:]+|[\s+:]+$", "", cleaned)
    cleaned = re.sub(r"\s{2,}", " ", cleaned).strip()
    
    return cleaned if cleaned else raw


def _leading_digits(segment: str) -> "re.Match[str] | None":
    return re.match(r"^\d+", segment.strip())


def _segment_has_suffix(segment: str) -> bool:
    m = re.match(r"^\d+(.*)$", segment.strip())
    if not m:
        return False
    return bool(m.group(1).strip())


def _breakdown_numeric_segments(raw: str, segments: List[str]) -> "List[str] | None":
    if len(segments) <= 1 or not all(_leading_digits(s) for s in segments):
        return None

    has_suffix = [_segment_has_suffix(s) for s in segments]

    if not any(has_suffix):
        return [raw]

    cleaned_segments = [clean_slip(s) for s in segments]

    if all(has_suffix):
        if len(cleaned_segments) <= MAX_HASIL_BREAKDOWN:
            return cleaned_segments
        return [raw]

    leading_lengths = [len(_leading_digits(s).group(0)) for s in segments]
    if len(set(leading_lengths)) > 1:
        return [raw]

    hasil: List[Any] = [None] * len(segments)
    last_suffix = None
    
    for i, s in enumerate(cleaned_segments):
        if has_suffix[i]:
            m = re.match(r"^\d+(.*)$", s)
            last_suffix = m.group(1) if m else None
            hasil[i] = s
        elif last_suffix is not None:
            hasil[i] = s + last_suffix

    next_suffix = None
    for i in range(len(segments) - 1, -1, -1):
        if has_suffix[i]:
            m = re.match(r"^\d+(.*)$", cleaned_segments[i])
            next_suffix = m.group(1) if m else None
            continue
        if hasil[i] is None and next_suffix is not None:
            hasil[i] = cleaned_segments[i] + next_suffix

    hasil = [h for h in hasil if h]
    if hasil and len(hasil) <= MAX_HASIL_BREAKDOWN:
        return hasil
        
    return [raw]


def breakdown_slip(value: Any) -> List[str]:
    if pd.isna(value):
        return [value]

    raw = str(value).strip()
    if not raw:
        return [raw]
        
    if raw.upper() in STANDALONE_SLIP_LABELS:
        return [raw]
        
    # Exception untuk format rentang tanggal (langsung kembalikan raw)
    if _is_date_range_slip(raw):
        return [raw]

    segments = [s.strip() for s in raw.split("+") if s.strip()]
    non_label_segments = [s for s in segments if s.upper() not in STANDALONE_SLIP_LABELS]

    if non_label_segments and all(_is_placeholder_p_code(s) for s in non_label_segments):
        hasil = re.sub(r"^\s*(TBA|EMPTY)\s*\+\s*", "", raw, flags=re.IGNORECASE)
        hasil = re.sub(r"\s*\+\s*(TBA|EMPTY)\s*$", "", hasil, flags=re.IGNORECASE)
        return [hasil.strip()]

    if non_label_segments and any(_is_placeholder_p_code(s) for s in non_label_segments):
        return [raw]

    if "+" in raw:
        hasil = _breakdown_numeric_segments(raw, segments)
        if hasil is not None:
            return hasil
    elif "-" in raw:
        dash_segments = [s.strip() for s in raw.split("-") if s.strip()]
        hasil = _breakdown_numeric_segments(raw, dash_segments)
        if hasil is not None:
            return hasil

    cleaned = clean_slip(raw)
    if not isinstance(cleaned, str) or not cleaned:
        return [raw]

    parts = [p.strip() for p in cleaned.split("+") if p.strip()]
    
    if not parts:
        return [raw]
    if len(parts) > MAX_HASIL_BREAKDOWN:
        return [raw]
        
    return parts


# INSURED
def _norm(s: str) -> str:
    return re.sub(r"\s+", " ", str(s).strip()).upper()


INSURED_EXCEPTIONS = {
    "JAPFA COMFEED INDONESIA - LAMONGAN",
}


def is_insured_exception(raw: str) -> bool:
    return _norm(raw) in INSURED_EXCEPTIONS


INSURED_NOISE_EXACT_RE = re.compile(
    r"\(\s*PERSERO\s*\)|\(\s*NON\s+TOBACCO\s*\)|\bNON\s+TOBACCO\b|\(\s*ONLY\s*\)",
    re.IGNORECASE,
)

INSURED_WORD_NOISE_RE = re.compile(
    r"\bPT\b|\bTBK\b|\bLTD\b|\bCO\b|\bPTD\b|\bSAU\b|\bMEI\b|\bFG\b|\bVAR\b|\bSDN\s+BHD\b",
    re.IGNORECASE,
)

SUCACO_PATTERN_RE = re.compile(r"SUCACO\s*-?\s*SIBALEC(\s+GROUP)?", re.IGNORECASE)


def apply_sucaco_rule(text: str) -> str:
    def _repl(m: "re.Match[str]") -> str:
        suffix = m.group(1) or "" 
        return f"SUCACO,SIBALEC{suffix}"

    return SUCACO_PATTERN_RE.sub(_repl, text)


JAYA_GROUP_PATTERN_RE = re.compile(
    r"JAYA\s+GROUP\s*/\s*JAYA\s+REAL\s+PROPERTY\s*&\s*JAWA\s+POS\s*/\s*"
    r"TEMPRINA\s+MEDIA\s+GRAFIKA\s+GROUP\s*/\s*ADIPRIMA\s+SURAPRINTA",
    re.IGNORECASE,
)


def apply_jaya_group_rule(text: str) -> str:
    if JAYA_GROUP_PATTERN_RE.search(text):
        return "JAYA GROUP,JAYA REAL PROPERTY,JAWA POS,TEMPRINA MEDIA GRAFIKA GROUP,ADIPRIMA SURAPRINTA"
    return text


METROPOLITAN_METLAND_RE = re.compile(
    r"METROPOLITAN\s+LAND\s+GROUP\s*/\s*METLAND\s+GROUP", re.IGNORECASE
)


def apply_metland_rule(text: str) -> str:
    return METROPOLITAN_METLAND_RE.sub("METROPOLITAN LAND GROUP", text)


PAREN_NO_BREAKDOWN_KEYWORDS = {"BGA", "APJT", "BUMA", "GROUP"}
PAREN_RE = re.compile(r"\(\s*([^()]*?)\s*\)")


def handle_parens(text: str) -> str:
    def _repl(m: "re.Match[str]") -> str:
        inner = m.group(1)
        if _norm(inner) in PAREN_NO_BREAKDOWN_KEYWORDS:
            return m.group(0)  
        return "," + inner  

    return PAREN_RE.sub(_repl, text)


INSURED_DELIM_RE = re.compile(r",|-|/|\bQQ\b|\bAND\b|\bOR\b", re.IGNORECASE)


def _clean_insured_part(part: str) -> str:
    part = INSURED_WORD_NOISE_RE.sub(" ", part)
    part = re.sub(r"[^\w()&]+$", "", part.strip())  
    part = re.sub(r"^[^\w()&]+", "", part.strip())  
    part = re.sub(r"\s{2,}", " ", part).strip()
    return part.upper()


def breakdown_insured(value: Any) -> List[str]:
    if pd.isna(value):
        return [value]

    raw = str(value).strip()
    if not raw:
        return [raw]

    if is_insured_exception(raw):
        return [raw]

    text = re.sub(r"\s+", " ", raw)
    text = apply_jaya_group_rule(text)
    text = apply_sucaco_rule(text)
    text = apply_metland_rule(text)
    
    text = INSURED_NOISE_EXACT_RE.sub("", text)
    text = handle_parens(text)
    text = INSURED_WORD_NOISE_RE.sub(" ", text)

    raw_parts = [p for p in INSURED_DELIM_RE.split(text) if p.strip()]
    hasil = []
    
    for p in raw_parts:
        c = _clean_insured_part(p)
        if not c:
            continue
        if not re.search(r"[A-Za-z0-9]", c):
            continue
        hasil.append(c)

    # Dedup, pertahankan urutan
    hasil = list(dict.fromkeys(hasil))

    if not hasil:
        return [raw]

    if len(hasil) > MAX_HASIL_BREAKDOWN:
        return [raw]

    return hasil


# PIPELINE UTAMA
def _insert_breakdown_columns(
    df: pd.DataFrame,
    source_col: str,
    fn_breakdown: callable,
    prefix: str,
    is_cedant_mask: pd.Series,
    max_cols_target: int = MAX_HASIL_BREAKDOWN,
    anchor_col: str = None,
) -> None:
    hasil_breakdown = df.apply(
        lambda row: fn_breakdown(row[source_col]) if is_cedant_mask[row.name] else [],
        axis=1,
    )

    lengths = hasil_breakdown.apply(len)
    max_in_data = int(lengths.max()) if not lengths.empty and lengths.max() > 0 else 1
    max_cols = min(max_in_data, max_cols_target)

    padded_hasil = hasil_breakdown.apply(lambda lst: lst + [np.nan] * (max_cols - len(lst)))

    df_breakdown = pd.DataFrame(
        padded_hasil.tolist(),
        columns=[f"{prefix}_{i + 1}" for i in range(max_cols)],
        index=df.index,
    )

    posisi = df.columns.get_loc(anchor_col or source_col)
    for i, col in enumerate(df_breakdown.columns):
        df.insert(posisi + 1 + i, col, df_breakdown[col])


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

    print(f"Memfilter data, hanya menyimpan {CEDANT_VALUE}...")
    is_cedant = df[CEDANT_COL].astype(str).str.strip() == CEDANT_VALUE
    df = df[is_cedant].copy()

    if df.empty:
        print(f"[INFO] Tidak ada data dengan cedant '{CEDANT_VALUE}'. Proses dihentikan.")
        return
        
    # --- MENAMBAHKAN KOLOM BUSINESS_PARTNERS ---
    if "COMP_NAME.1" in df.columns and CEDANT_COL in df.columns and INSURED_COL in df.columns:
        posisi_insert = df.columns.get_loc(INSURED_COL)

        kondisi_direct = df["COMP_NAME.1"].astype(str).str.strip().str.upper() == "DIRECT"
        nilai_business_partners = np.where(kondisi_direct, df[CEDANT_COL], df["COMP_NAME.1"])

        df.insert(posisi_insert, "BUSINESS_PARTNERS", nilai_business_partners)

    is_cedant_mask = pd.Series(True, index=df.index)

    kolom_sisa = [
        col for col in df.columns
        if col.startswith((f"{INSURED_COL}_CLN", f"{POLICY_COL}_CLEAN", f"{SLIP_COL}_CLEAN"))
    ]
    if kolom_sisa:
        df.drop(columns=kolom_sisa, inplace=True)

    print("Menjalankan pembersihan dan breakdown kolom...")
    _insert_breakdown_columns(df, INSURED_COL, breakdown_insured, f"{INSURED_COL}_CLN", is_cedant_mask)

    # 1. Jalankan breakdown polis DULU agar kolom FAC_POLICY_NO_CLEAN_1 terbuat
    _insert_breakdown_columns(
        df, POLICY_COL, breakdown_polis, f"{POLICY_COL}_CLEAN", is_cedant_mask,
        anchor_col=POLICY_COL,  # Target anchor diubah kembali ke POLICY_COL
    )

    # 2. Sisipkan kolom CERTIFICATE_1 tepat SETELAH FAC_POLICY_NO_CLEAN_1
    # Kita asumsikan format nama kolom breakdown polis adalah "{POLICY_COL}_CLEAN_1"
    nama_kolom_breakdown_pertama = f"{POLICY_COL}_CLEAN_1"
    
    # Cek apakah kolom hasil breakdown pertama berhasil dibuat 
    # (berjaga-jaga jika semua hasil breakdown polis kosong)
    if nama_kolom_breakdown_pertama in df.columns:
        posisi_policy_clean = df.columns.get_loc(nama_kolom_breakdown_pertama)
        df.insert(posisi_policy_clean + 1, "CERTIFICATE_1", df[POLICY_COL].apply(extract_certificate))
    else:
        # Fallback jika tidak ada hasil breakdown, sisipkan setelah POLICY_COL
        posisi_policy = df.columns.get_loc(POLICY_COL)
        df.insert(posisi_policy + 1, "CERTIFICATE_1", df[POLICY_COL].apply(extract_certificate))


    _insert_breakdown_columns(df, SLIP_COL, breakdown_slip, f"{SLIP_COL}_CLEAN", is_cedant_mask)

    print(f"Menyimpan hasil ke Excel: {output_file} ...")
    os.makedirs(os.path.dirname(output_file), exist_ok=True)
    df.to_excel(output_file, index=False)

    print("Pemrosesan data MITRAUTAMA DATA 1 - FACUL selesai!")


if __name__ == "__main__":
    process_data(INPUT_FILE, SHEET_NAME, OUTPUT_FILE)