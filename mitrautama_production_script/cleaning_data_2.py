import os
import re
import warnings
from typing import Any, Callable, List

import numpy as np
import pandas as pd

warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

INPUT_FILE = os.path.join("data_excel", "raw", "2b. transaksi Osbal 01.01.23 - 17.08.26.xlsx")
SHEET_NAME = "Query result"
OUTPUT_FILE = os.path.join("data_excel", "processed", "mitrautama_output_osbal.xlsx")

CEDANT_COL = "CCOS_COMP_NAME"
CEDANT_VALUE = "PT. MITRA UTAMA REASURANSI"

INSURED_COL = "FAC_INSURED"
POLICY_COL = "FAC_POLICY_NO"
SLIP_COL = "FAC_SLIP"

CLSDT_POLICY_COL = "CLSDT_POLICY_NO"
CLSDT_SLIP_COL = "CLSDT_SLIP_NO"
CLSDT_SERTF_COL = "CLSDT_SERTF_NO"
CERTIFICATE_COL = "CERTIFICATE_1"

SLIP_CLEAN_PREFIX = "SLIP_CLEAN"
POLICY_CLEAN_PREFIX = "POLICY_CLEAN"

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
    if "ASWATA/MSIG" in norm:
        return True
    if "TBA + P2" in norm:
        return True
    if norm == "TBA":
        return True
    if "PENYELESAIAN SUSPENSE" in norm:
        return True
    return False


def clean_polis(value: Any) -> Any:
    if pd.isna(value):
        return value

    raw = str(value).strip().upper()
    if not raw:
        return raw

    if is_exception(raw):
        return raw

    if raw.upper() == "RBA":
        return raw

    cleaned = raw
    cleaned = re.sub(r"^(TBA|RBA)\s*/?\s*", "", cleaned, flags=re.IGNORECASE)
    cleaned = re.sub(r"\b(TBA|RBA)\s*/\s*", "", cleaned, flags=re.IGNORECASE)
    cleaned = re.sub(r"\b(TBA|RBA)\b", "", cleaned, flags=re.IGNORECASE)

    policy_pattern = r"\d{4,5}[/A-Z0-9.\-]+\d{2,4}"
    cleaned = re.sub(r"(" + policy_pattern + r")\s*SISA OSBAL", r"\1", cleaned, flags=re.IGNORECASE)
    cleaned = re.sub(r"(" + policy_pattern + r")\s*/\s*[A-Za-z\s]+$", r"\1", cleaned, flags=re.IGNORECASE)

    company_keywords = [
        "VICTORIA INSURANCE", "FPG", "ASURANSI", "MPM", "INDONESIA", "MAG",
        "UMUM MEGA", "ZURICH", "GREAT EASTERN", "GENERAL INSURANCE INDONESIA",
        "ACA", "SUNDAY INSURANCE INDONESIA,PT", "ASURANSI FPG INDONESIA", "SUNDAY INSURANCE",
        "KSK INSURANCE", "( SMU WINGS )"
    ]
    company_keywords.sort(key=len, reverse=True)
    for comp in company_keywords:
        comp_pattern = re.escape(comp) + r"\s*(?:/\s*)?(?=\d{4,5}[/A-Z0-9.\-])"
        cleaned = re.sub(comp_pattern, "", cleaned, flags=re.IGNORECASE)

    cleaned = re.sub(r"^[A-Za-z\s,]+/\s*(?=\d{4,5}[/A-Z0-9.\-])", "", cleaned)
    cleaned = re.sub(r"[.\/\-\"]", "", cleaned)
    cleaned = re.sub(r"\b(PT|Tbk|0)\b", "", cleaned, flags=re.IGNORECASE)
    cleaned = re.sub(r"\s{2,}", " ", cleaned).strip()

    return cleaned if cleaned else raw


def breakdown_polis(value: Any) -> List[str]:
    if pd.isna(value):
        return [value]

    raw = str(value).strip()
    raw = raw.replace('"', '')

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

        match_suffix = re.search(r"\b\d{4,5}([/A-Z0-9.\-]+)", base_policy)
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


# SLIP
STANDALONE_SLIP_LABELS = {"TBA", "EMPTY"}
SLIP_LABEL_RE = re.compile(r"\bTBA\b|\bEMPTY\b", re.IGNORECASE)
P_CODE_RE = re.compile(r"^P\d+[A-Za-z]*(?:\(P\d+\))?$", re.IGNORECASE)

GROUP_NOISE_WORDS = [
    "SAYAP MAS WINGS GROUP",
    "PROTELINDO GROUP",
    "JAYA TRADE GROUP",
    "METROPOLITAN KENTJANA GROUP",
    "SEKAR LAUT GROUP",
    "SISA OSBAL"
]

SLIP_FORMAT_RE = re.compile(r"\d{3,6}(?:[./\-]+[A-Za-z0-9]+){2,8}")
SHORT_NUM_TOKEN_RE = re.compile(r"^\d{4,6}$")
SLIP_EXCEPTION_RE = re.compile(r"^\d{5,}\s*\+\s*P\d", re.IGNORECASE)
TBA_DASH_RE = re.compile(r"^TBA\s*-+\s*$", re.IGNORECASE)


def _is_placeholder_p_code(segment: str) -> bool:
    return bool(P_CODE_RE.match(segment.strip()))


def _is_pure_group_noise(token: str) -> bool:
    t = token.strip()
    for noise in GROUP_NOISE_WORDS:
        if re.fullmatch(re.escape(noise), t, flags=re.IGNORECASE):
            return True
    return False


STANDALONE_TOKEN_RE = re.compile(r"^[A-Za-z]{0,4}\d{3,}$")


def _try_split_comma_amp(raw: str) -> Any:
    if "," not in raw and "&" not in raw:
        return None

    parts = [p.strip() for p in re.split(r"\s*,\s*|\s*&\s*", raw) if p.strip()]
    if len(parts) < 2:
        return None

    if all(STANDALONE_TOKEN_RE.match(p) for p in parts):
        return parts

    return None


LEADING_SLIP_NOISE_RE = re.compile(r"^(?:TBA|EMPTY)\b\s*/?\s*", re.IGNORECASE)
LEADING_SHORT_CODE_RE = re.compile(r"^[A-Za-z]{2,10}\s*/\s*(?=\d{3,6})", re.IGNORECASE)


def _strip_leading_slip_noise(t: str) -> str:
    text = t.strip()
    prev = None
    while text != prev:
        prev = text
        text = LEADING_SLIP_NOISE_RE.sub("", text).strip()
        text = LEADING_SHORT_CODE_RE.sub("", text).strip()
    return text


def clean_slip(value: Any) -> Any:
    if pd.isna(value):
        return value

    raw = str(value).strip()
    if not raw or raw.upper() in STANDALONE_SLIP_LABELS or TBA_DASH_RE.match(raw):
        return raw

    if raw.upper() == "TBA SISA OSBAL":
        return raw

    label_match = SLIP_LABEL_RE.search(raw)
    cleaned = SLIP_LABEL_RE.sub("", raw)
    cleaned = re.sub(r"^[\s+/\-:,.]+|[\s+/\-:,.]+$", "", cleaned)
    cleaned = LEADING_SHORT_CODE_RE.sub("", cleaned).strip()

    for noise in GROUP_NOISE_WORDS:
        cleaned = re.sub(re.escape(noise), "", cleaned, flags=re.IGNORECASE)

    cleaned = re.sub(r"-\d{2}\b", "", cleaned)

    stripped_remainder = cleaned.strip()
    if label_match and stripped_remainder and not re.search(r"\d", stripped_remainder):
        if len(stripped_remainder.split()) >= 2:
            return label_match.group(0).upper()

    cleaned = re.sub(r"[./\-\"]", "", cleaned)
    cleaned = re.sub(r"\b(PT|Tbk)\b\.?", "", cleaned, flags=re.IGNORECASE)
    cleaned = re.sub(r",", "", cleaned)
    cleaned = re.sub(r"\s{2,}", " ", cleaned).strip()

    return cleaned if cleaned else raw


def _short_and_full_digit_lengths_match(tokens: List[str]) -> bool:
    full_len = None
    for t in tokens:
        if SLIP_FORMAT_RE.search(t):
            m = re.match(r"^(\d{3,6})", t)
            if m:
                full_len = len(m.group(1))
                break

    if full_len is None:
        return True

    for t in tokens:
        if SHORT_NUM_TOKEN_RE.match(t) and len(t) != full_len:
            return False

    return True


def _process_plus_group(tokens: List[str]) -> List[str]:
    filtered: List[str] = []
    for t in tokens:
        if t.upper() in STANDALONE_SLIP_LABELS or _is_placeholder_p_code(t) or _is_pure_group_noise(t):
            continue
        filtered.append(t)

    if not filtered:
        return []

    full_idx = None
    suffix = None
    for i, t in enumerate(filtered):
        t_stripped = _strip_leading_slip_noise(t)
        if SLIP_FORMAT_RE.search(t_stripped):
            m = re.match(r"^(\d{3,6})(.*)$", t_stripped)
            if m:
                full_idx = i
                suffix = m.group(2)
                break

    if full_idx is None:
        result: List[str] = []
        for t in filtered:
            c = clean_slip(t)
            if isinstance(c, str) and c:
                result.append(c)
        return result

    ordered_tokens = [filtered[full_idx]]
    ordered_tokens += list(reversed(filtered[:full_idx]))
    ordered_tokens += filtered[full_idx + 1:]

    result = []
    for i, t in enumerate(ordered_tokens):
        if i == 0:
            c = clean_slip(t)
        elif SHORT_NUM_TOKEN_RE.match(t) and suffix:
            c = clean_slip(t + suffix)
        else:
            c = clean_slip(t)

        if isinstance(c, str) and c:
            result.append(c)

    return result


def _is_pure_p_code_result(parts: List[str]) -> bool:
    if not parts:
        return False
    for p in parts:
        subtokens = [s.strip() for s in p.split("+") if s.strip()]
        if not subtokens or not all(_is_placeholder_p_code(s) for s in subtokens):
            return False
    return True


def _process_segment(seg: str) -> List[str]:
    seg = seg.strip()
    if not seg or SLIP_EXCEPTION_RE.match(seg) or seg.upper() in STANDALONE_SLIP_LABELS:
        return [seg]

    segments = [s.strip() for s in seg.split("+") if s.strip()]
    if not segments:
        return []

    non_label_segments = [s for s in segments if s.upper() not in STANDALONE_SLIP_LABELS]

    if non_label_segments and all(_is_placeholder_p_code(s) for s in non_label_segments):
        hasil = re.sub(r"^\s*(TBA|EMPTY)\s*\+\s*", "", seg, flags=re.IGNORECASE)
        hasil = re.sub(r"\s*\+\s*(TBA|EMPTY)\s*$", "", hasil, flags=re.IGNORECASE)
        hasil = hasil.strip()
        return [hasil] if hasil else []

    if not _short_and_full_digit_lengths_match(segments):
        return [seg]

    return _process_plus_group(segments)


def breakdown_slip(value: Any) -> List[str]:
    if pd.isna(value):
        return [value]

    raw = str(value).strip()
    raw = raw.replace('"', '')

    if not raw or SLIP_EXCEPTION_RE.match(raw) or raw.upper() in STANDALONE_SLIP_LABELS or TBA_DASH_RE.match(raw):
        return [raw]

    if raw.upper() == "TBA SISA OSBAL":
        return [raw]

    comma_amp_parts = _try_split_comma_amp(raw)
    if comma_amp_parts is not None:
        gap_segments = comma_amp_parts
    else:
        gap_segments = [s.strip() for s in re.split(r"\s{2,}", raw) if s.strip()]
        if not gap_segments:
            gap_segments = [raw]

    segment_results = [_process_segment(seg) for seg in gap_segments]
    has_real_slip = any(not _is_pure_p_code_result(res) for res in segment_results)

    all_parts: List[str] = []
    for res in segment_results:
        if len(gap_segments) > 1 and has_real_slip and _is_pure_p_code_result(res):
            continue
        all_parts.extend(res)

    all_parts = [p for p in all_parts if p]

    seen = set()
    deduped: List[str] = []
    for p in all_parts:
        key = p.upper()
        if key in seen:
            continue
        seen.add(key)
        deduped.append(p)

    if len(deduped) > 1:
        deduped = [p for p in deduped if p.upper() not in STANDALONE_SLIP_LABELS]

    if not deduped:
        return [raw]

    if len(deduped) > MAX_HASIL_BREAKDOWN:
        return [raw]

    return deduped


# INSURED
INSURED_NOISE_PAREN_RE = re.compile(r"\(\s*(PERSERO|ONLY)\s*\)", re.IGNORECASE)
INSURED_LEGAL_RE = re.compile(r"\b(PT|TBK|PERSERO)\b\.?", re.IGNORECASE)
INSURED_SPLIT_RE = re.compile(r",|-|&|/|\(|\)|\bQQ\b", re.IGNORECASE)
INSURED_FG_NOISE_RE = re.compile(r"^\s*FG\s*(DISTRIBUTOR)?\s*$", re.IGNORECASE)


def clean_insured(name: Any) -> str:
    if pd.isna(name):
        return ""
    text = INSURED_LEGAL_RE.sub("", str(name))
    text = re.sub(r"\s{2,}", " ", text).strip(" ,./-")
    return text.upper()


def breakdown_insured(value: Any) -> List[str]:
    if pd.isna(value):
        return [value]

    raw = str(value).strip()
    if not raw:
        return [raw]

    text = INSURED_NOISE_PAREN_RE.sub("", raw)
    parts_raw = [p.strip() for p in INSURED_SPLIT_RE.split(text) if p.strip()]

    hasil = []
    for p in parts_raw:
        p_clean = clean_insured(p)
        if not p_clean or INSURED_FG_NOISE_RE.match(p_clean):
            continue
        hasil.append(p_clean)

    if not hasil:
        return [raw]

    if len(hasil) > MAX_HASIL_BREAKDOWN:
        return [raw]

    return hasil


def _is_blank(value: Any) -> bool:
    if pd.isna(value):
        return True
    return not str(value).strip()


SLIP_NUMBER_FORMAT_RE = re.compile(r"^\d{3,6}(?:/[A-Za-z0-9]+){3,}(?:-\d{1,3})?$")


def is_slip_number_format(value: Any) -> bool:
    if _is_blank(value):
        return False
    norm = re.sub(r"\s+", "", str(value).strip().upper())
    return bool(SLIP_NUMBER_FORMAT_RE.fullmatch(norm))


def pick_slip_source(row: pd.Series) -> Any:
    fac_slip = str(row.get(SLIP_COL, "")).strip().upper()
    clsdt_slip = row.get(CLSDT_SLIP_COL, np.nan)
    clsdt_polis = row.get(CLSDT_POLICY_COL, np.nan)

    slip_blank = _is_blank(clsdt_slip)
    polis_blank = _is_blank(clsdt_polis)
    slip_norm = "" if slip_blank else re.sub(r"\s+", " ", str(clsdt_slip).strip().upper())
    polis_norm = "" if polis_blank else re.sub(r"\s+", " ", str(clsdt_polis).strip().upper())

    if fac_slip == "TBA" and ("SISA OSBAL" in slip_norm or "SISA OSBAL" in polis_norm):
        return "TBA SISA OSBAL"

    if is_slip_number_format(clsdt_slip):
        return clsdt_slip

    fallback_ke_fac_slip = (
        (slip_blank and polis_blank)
        or (slip_norm == "-" and polis_norm == "-")
        or (slip_norm == "SISA OSBAL" and polis_norm == "SISA OSBAL")
        or (slip_blank and polis_norm == "0")
    )
    if fallback_ke_fac_slip:
        return row.get(SLIP_COL, np.nan)

    return clsdt_polis


def pick_policy_source(row: pd.Series) -> Any:
    policy = row.get(POLICY_COL, np.nan)
    insured = str(row.get(INSURED_COL, "")).strip()

    if "PENYELESAIAN SUSPENSE" in insured.upper():
        pol_str = str(policy).strip() if pd.notna(policy) else ""
        if pol_str:
            return f"{pol_str} {insured}"
        return insured

    return policy


# CERTIFICATE
CERT_NUM_RE = re.compile(r"\b\d{5,6}\b")
CERT_RANGE_RE = re.compile(
    r"(\d{5,6})\s*(?:-{1,2}|S\s*/\s*D|SD)\s*(\d{5,6})",
    re.IGNORECASE,
)


def _pad_cert(num: Any) -> str:
    return str(num).zfill(6)


def _process_cert_segment(seg: str) -> Any:
    seg = seg.strip()
    if not seg:
        return None

    m = CERT_RANGE_RE.search(seg)
    if m:
        start_raw, end_raw = m.group(1), m.group(2)
        start, end = int(start_raw), int(end_raw)
        if end < start:
            start, end = end, start

        jumlah = end - start + 1
        if jumlah > 3:
            return f"{_pad_cert(start_raw)} SD {_pad_cert(end_raw)}"
        else:
            angka_list = [_pad_cert(n) for n in range(start, end + 1)]
            return ", ".join(angka_list)

    angka_ditemukan = CERT_NUM_RE.findall(seg)
    if not angka_ditemukan:
        return None

    return ", ".join(_pad_cert(n) for n in angka_ditemukan)


def build_certificate(value: Any) -> Any:
    if pd.isna(value):
        return np.nan

    raw = str(value).strip()
    if not raw:
        return np.nan

    segmen_list = [s.strip() for s in raw.split(",") if s.strip()]
    if not segmen_list:
        segmen_list = [raw]

    hasil_segmen = []
    for seg in segmen_list:
        hasil = _process_cert_segment(seg)
        if hasil:
            hasil_segmen.append(hasil)

    if not hasil_segmen:
        return np.nan

    return ", ".join(hasil_segmen)


# PIPELINE UTAMA
def _insert_breakdown_columns(
    df: pd.DataFrame,
    insert_after_col: str,
    get_source_value: Callable[[pd.Series], Any],
    fn_breakdown: callable,
    prefix: str,
) -> None:
    hasil_breakdown = df.apply(
        lambda row: fn_breakdown(get_source_value(row)),
        axis=1,
    )

    lengths = hasil_breakdown.apply(len)
    max_cols = int(lengths.max()) if not lengths.empty and lengths.max() > 0 else 1

    padded_hasil = hasil_breakdown.apply(lambda lst: lst + [np.nan] * (max_cols - len(lst)))

    df_breakdown = pd.DataFrame(
        padded_hasil.tolist(),
        columns=[f"{prefix}_{i + 1}" for i in range(max_cols)],
        index=df.index,
    )

    posisi = df.columns.get_loc(insert_after_col)
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

    kolom_sisa = [
        col for col in df.columns
        if col.startswith((f"{INSURED_COL}_CLN", f"{POLICY_COL}_CLEAN", f"{SLIP_COL}_CLEAN", SLIP_CLEAN_PREFIX, POLICY_CLEAN_PREFIX, CERTIFICATE_COL))
    ]
    if kolom_sisa:
        df.drop(columns=kolom_sisa, inplace=True)

    for clsdt_col in (CLSDT_POLICY_COL, CLSDT_SLIP_COL, CLSDT_SERTF_COL):
        if clsdt_col not in df.columns:
            print(f"[WARNING] Kolom '{clsdt_col}' tidak ditemukan, akan selalu memakai data FAC sebagai fallback.")
            df[clsdt_col] = np.nan

    print("Menjalankan pembersihan dan breakdown kolom...")
    _insert_breakdown_columns(df, INSURED_COL, lambda row: row[INSURED_COL], breakdown_insured, f"{INSURED_COL}_CLN")
    _insert_breakdown_columns(df, POLICY_COL, pick_policy_source, breakdown_polis, POLICY_CLEAN_PREFIX)

    print("Membuat kolom CERTIFICATE dari CLSDT_SERTF_NO...")
    nilai_certificate = df[CLSDT_SERTF_COL].apply(build_certificate)

    nama_kolom_policy_clean_1 = f"{POLICY_CLEAN_PREFIX}_1"
    if nama_kolom_policy_clean_1 in df.columns:
        posisi_policy_clean = df.columns.get_loc(nama_kolom_policy_clean_1)
        df.insert(posisi_policy_clean + 1, CERTIFICATE_COL, nilai_certificate)
    else:
        posisi_policy = df.columns.get_loc(POLICY_COL)
        df.insert(posisi_policy + 1, CERTIFICATE_COL, nilai_certificate)

    _insert_breakdown_columns(df, SLIP_COL, pick_slip_source, breakdown_slip, SLIP_CLEAN_PREFIX)

    print(f"Menyimpan hasil ke Excel: {output_file} ...")
    os.makedirs(os.path.dirname(output_file), exist_ok=True)
    df.to_excel(output_file, index=False)

    print("Pemrosesan data MITRAUTAMA DATA 2 selesai!")


if __name__ == "__main__":
    process_data(INPUT_FILE, SHEET_NAME, OUTPUT_FILE)