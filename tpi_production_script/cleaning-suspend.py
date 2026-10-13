"""
Cleaning Data 3 - Suspend | Cedant: TPI (PT Asuransi Tugu Pratama Indonesia)

UPDATE BARU -- KOLOM CERTIFICATE:
- Berdasarkan eksplorasi pola polis TPI Data 3, TIDAK ditemukan pola
  sertifikat (base+"-"+cert) yang menempel di kolom POLIS -- berbeda dari
  Data 1 & Data 2. Kolom CERTIFICATE tetap DITAMBAHKAN (untuk konsistensi
  struktur output antar cedant/data), tapi SELALU DIKOSONGKAN.
- Kalau di kemudian hari ditemukan pola sertifikat di Data 3 TPI, tinggal
  isi fungsi extract_certificate() di bawah (saat ini return "" selalu).
"""

import os
import re

import pandas as pd


# ─────────────────────────────────────────────────────────────────────────────
# KONFIGURASI — ubah bagian ini kalau nama file/kolom berbeda
# ─────────────────────────────────────────────────────────────────────────────

INPUT_FILE  = os.path.join("input", "3b. Database Suspense 150826.xlsx")
INPUT_SHEET = "Sheet1"
HEADER_ROW  = 2          # header data ada di baris ke-3 file (index 2, 0-based)

OUTPUT_FILE = os.path.join("output", "tpi_output_suspend.xlsx")

CEDANT_FILTER_COL   = "CEDANT SHRT NAME"
CEDANT_FILTER_VALUE = "TPI"

# Rule baru: hanya baris berstatus SUSPENSE yang dicleaning & ditampilkan;
# baris ADJUSTED dibuang total sebelum proses cleaning.
STATUS_COL          = "STATUS"
STATUS_KEEP_VALUE   = "SUSPENSE"

INSURED_COL = "INSURED"
POLIS_COL   = "POLIS"
SLIP_COL    = "SLIP NO"

MAX_SPLIT_COLS = 5   # batas maksimal breakdown insured, sisanya digabung koma

# ─────────────────────────────────────────────────────────────────────────────
# REGEX & KAMUS PEMBANTU (diadaptasi dari referensi)
# ─────────────────────────────────────────────────────────────────────────────

_INSURED_TAIL_RE = re.compile(
    r"""
    \bAS\b\s+(?:THE\s+)?(?:PRINCIPAL|OFF-TAKER|MAINTENANCE|CONTRACTOR).*
  | \bBEING\s+(?:THE\s+)?(?:PRINCIPAL|OFF-TAKER).*
  | \bAND\s+ALL\s+SUBSIDIARI.*
  | \bINCLUDING\s+ALL\s+SUBSIDIARI.*
  | \bINCLUDING\s+ANY\s+SUBSIDIAR.*
  | \bCOMPRISING\s+OF.*
  | \bINSTALLMENT\b.*
  | \bRELATED\s+COMPANY\b.*
  | \bPURCHASED\s+OR\s+OTHERWISE\b.*
    """,
    re.IGNORECASE | re.VERBOSE,
)

_INSURED_JUNK_WORDS = frozenset({
    "SHANGHAI", "PR OF CHINA", "CHINA", "INDONESIA", "JAKARTA",
    "OFFICERS", "EMPLOYEES", "ALL OTHER CONTRACTORS",
    "SUB-CONTRACTORS", "SUB CONTRACTORS",
    "COMPANIES", "AFFILIATED", "AFFILIATES",
    "CORPORATIONS AND INCLUDING PARTNERSHIP",
    "JOINT VENTURES AND AGREEMENT OR BY LAW",
    "AS THEIR RESPECTIVE INTEREST MAY APPEAR",
    "AS THEIR RESPECTIVE INTERESTS MAY APPEAR",
    "SUBSIDIARY", "SUBSIDIARIES", "ANY SUBSIDIARY COMPANY",
    "RELATED COMPANY",
    "FOR THEIRS RESPECTIVE RIGHTS AND INTEREST",
    "MIGRASI AS400", "THE PRINCIPAL", "PRINCIPAL", "OWNER", "REFUND PREMI",
})

# P1 / P2 / P3 tetap dipertahankan apa adanya (sesuai catatan eksplorasi TPI)
_POLIS_KEEP_AS_IS_RE = re.compile(r"^\s*P[123]\s*/", re.IGNORECASE)


# ─────────────────────────────────────────────────────────────────────────────
# KAMUS POLA STANDAR TPI (dari Karakteristik_Polis_dan_Slip.xlsx sheet "TPI"
# + Bismillah_TPI_-_Eksplorasi_Cedant.docx)
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
    """Kalau hasil cleaning belum cocok kamus pola standar TPI, coba beberapa
    perbaikan ringan (spasi, simbol liar) sebelum menyerah dan mengembalikan
    hasil cleaning apa adanya. Ini TIDAK mengubah makna/isi data, hanya
    merapikan format supaya cocok pola standar bila memungkinkan."""
    if not value:
        return value
    if _matches_known_pattern(value, patterns):
        return value

    candidates = [
        re.sub(r"\s+", "", value),                       # hapus semua spasi
        re.sub(r"\s*/\s*", " / ", value).strip(),         # normalisasi spasi di sekitar '/'
        value.strip(" .-/"),                              # buang simbol liar di ujung
        re.sub(r"\.0+$", "", value),                      # sisa ".0" dari konversi angka Excel
        "0" + value if value.isdigit() else value,        # kembalikan angka 0 di depan yang hilang
        re.sub(r"^(\d+)(\s*/\s*(?:19|20)\d{2})$", r"0\1\2", value),  # "angka / tahun" kehilangan 0 di depan
    ]
    for c in candidates:
        if _matches_known_pattern(c, patterns):
            return c

    return value  # tidak cocok pola manapun -> tetap kembalikan hasil cleaning awal


# ─────────────────────────────────────────────────────────────────────────────
# KOLOM CERTIFICATE (BARU) -- TIDAK ADA POLA SERTIFIKAT DI DATA 3 TPI,
# SELALU DIKOSONGKAN. Lihat catatan di docstring atas file.
# ─────────────────────────────────────────────────────────────────────────────

def extract_certificate(polis_raw) -> str:
    """Data 3 TPI tidak punya pola sertifikat yang menempel di POLIS
    (berbeda dari Data 1 & Data 2 TPI) -- selalu kosong."""
    return ""


# ─────────────────────────────────────────────────────────────────────────────
# FUNGSI CLEANING
# ─────────────────────────────────────────────────────────────────────────────

def clean_polis(val) -> str:
    """Bersihkan nomor polis. P1/P2/P3 di depan dibiarkan apa adanya."""
    if pd.isna(val):
        return ""
    val = str(val).strip()
    if not val:
        return ""

    # Sesuai catatan TPI: "abaikan P1, biarkan apa adanya" / "abaikan P2, biarkan apa adanya"
    if _POLIS_KEEP_AS_IS_RE.match(val):
        return val

    # Jika berupa string gabungan (Polis+Tahun+Currency+Kode+Slip+Suffix), ambil nomor polisnya
    comp_polis = _COMPOSITE_POLIS_RE.search(val)
    if comp_polis:
        return comp_polis.group(1)

    # Hanya hapus suffix -EXT(angka) di bagian paling akhir.
    p = re.sub(r"-EXT\(\d+\)$", "", val, flags=re.IGNORECASE)

    # Hapus suffix "-NNN" atau "-NNN/NNN" berulang di akhir
    while True:
        stripped = re.sub(r"-\d+(?:/\d+)?$", "", p)
        if stripped == p:
            break
        p = stripped

    return p if p else val


def clean_slip(val) -> str:
    """Bersihkan nomor slip. Jika berbentuk string gabungan
    (Polis+Tahun+Currency+Kode+Slip+Suffix), ekstrak slip 10-digit."""
    if pd.isna(val):
        return ""
    val_str = str(val).strip()
    if not val_str:
        return ""

    m = _COMPOSITE_SLIP_RE.search(val_str)
    if m:
        return m.group(1)

    return val_str


def _remove_polis_slip_from_text(text: str, polis_ori, slip_ori) -> str:
    if pd.notna(polis_ori):
        token = str(polis_ori).strip()
        if token and token != "-":
            text = text.replace(token, "")
    if pd.notna(slip_ori):
        token = str(slip_ori).strip()
        if token and token != "-":
            text = text.replace(token, "")
    return text


def _normalize_insured_part(p: str) -> str:
    p = _INSURED_TAIL_RE.sub("", p)
    p = re.sub(r"\bKB\b", "", p, flags=re.IGNORECASE)
    p = re.sub(r"\bA\.?W\.?\b", "", p, flags=re.IGNORECASE)
    p = re.sub(
        r"\b(?:BAPAK|IBU|SDRI?|NYONYA|NY|TUAN|MR|MRS|MS|"
        r"DR|IR|PROF|HJ|H)\b\.?\s*",
        "", p, flags=re.IGNORECASE,
    )
    p = re.sub(r"\(\s*\)", "", p)
    p = re.sub(r"^(?:AND|OR)\b\s*", "", p, flags=re.IGNORECASE)
    p = re.sub(r"\s*\b(?:AND|OR)$", "", p, flags=re.IGNORECASE)
    p = re.sub(r"^[^a-zA-Z0-9(]+", "", p)
    p = re.sub(r"[^a-zA-Z0-9)]+$", "", p)
    return p.strip()


def _is_valid_insured_part(p: str) -> bool:
    if len(p) <= 2:
        return False
    up = p.upper()
    if re.match(r"^[\d\/\-\.]+$", up):
        return False
    if re.search(r"\b(?:NO\.\s*\d+|BUILDING|FLOOR|ROOM|ROAD|STREET|TOWER|KAV\.?|BLOK)\b", up):
        return False
    return up not in _INSURED_JUNK_WORDS

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
    """Segmen yang diawali angka (mis. "7 SUPPLY VESSEL") itu nama
    unit/armada, bukan perusahaan baru -> digabung ke segmen sebelumnya."""
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
    """Kalau kata TERAKHIR di segmen gabungan ternyata singkatan dari segmen
    sebelumnya (mis. "PDBI DSLNG" -> "DSLNG" singkatan "DONGGI-SENORO LNG"),
    buang kata itu saja, sisanya tetap."""
    words = name.split()
    if len(words) >= 2 and any(_is_abbreviation_of(words[-1], p) for p in priors):
        return " ".join(words[:-1])
    return name


def clean_insured(val, polis_ori, slip_ori) -> list:
    """Pecah nama insured menjadi list bagian-bagian bersih (breakdown)."""
    if pd.isna(val):
        return []
    val = str(val).strip()
    if not val:
        return []

    val = _remove_polis_slip_from_text(val, polis_ori, slip_ori)

    val = re.sub(r"\(\s*[\d\.\/\-]+\s*\)", "", val)
    val = re.sub(r"\b(?:AND|AN|OR)\s*/\s*(?:AND|OR)\b", ",", val, flags=re.IGNORECASE)
    val = re.sub(r"\bAND\s+OR\b", ",", val, flags=re.IGNORECASE)
    val = re.sub(r"\bCO\.,?\s*LTD\.?\b", ",", val, flags=re.IGNORECASE)
    for pattern in [r"\bTBK\.?\b", r"\(PERSERO\)", r"\bPERSERO\b", r"\bLTD\.?\b"]:
        val = re.sub(pattern, "", val, flags=re.IGNORECASE)

    split_pattern = r"\bQQ\b|/|,|\d+\.|\bPT\.?\b|\bCV\.?\b|:|;"
    parts = re.split(split_pattern, val, flags=re.IGNORECASE)
    parts = _merge_digit_start_parts(parts)
    parts = _merge_region_prefix_parts(parts)

    cleaned = []
    for p in parts:
        p = _normalize_insured_part(p)
        if _is_valid_insured_part(p):
            cleaned.append(p.replace(".", "").upper())

    if not cleaned:
        fallback = _normalize_insured_part(val).replace(".", "").upper()
        return [fallback] if fallback else []

    final = []
    for name in cleaned:
        name = _strip_trailing_abbrev_word(name, final)
        if any(_is_abbreviation_of(name, prior) for prior in final):
            continue
        final.append(name)
    cleaned = final

    if len(cleaned) > MAX_SPLIT_COLS:
        return [", ".join(cleaned)]

    return cleaned


# ─────────────────────────────────────────────────────────────────────────────
# PROSES UTAMA
# ─────────────────────────────────────────────────────────────────────────────

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

    # Rename kolom asli -> _ORI (gaya Wahana)
    df.rename(columns={
        INSURED_COL: "INSURED_ORI",
        POLIS_COL:   "POLIS_ORI",
        SLIP_COL:    "SLIP_NO_ORI",
    }, inplace=True)

    print("[3/5] Menjalankan proses cleaning ...")

    polis_cln, slip_cln, all_certificate, all_insured_parts = [], [], [], []
    max_insured_parts = 1

    for _, row in df.iterrows():
        p_ori = row.get("POLIS_ORI", "")
        s_ori = row.get("SLIP_NO_ORI", "")
        i_ori = row.get("INSURED_ORI", "")

        polis_cln.append(refine_with_known_pattern(clean_polis(p_ori), KNOWN_POLIS_PATTERNS))
        slip_cln.append(refine_with_known_pattern(clean_slip(s_ori), KNOWN_SLIP_PATTERNS))
        all_certificate.append(extract_certificate(p_ori))

        parts = clean_insured(i_ori, p_ori, s_ori)
        max_insured_parts = max(max_insured_parts, len(parts))
        all_insured_parts.append(parts)

    print("[4/5] Menyusun kolom output ...")

    df["POLIS_CLEAN_1"] = [
        x if x not in (None, "") else None
        for x in polis_cln
    ]

    df["CERTIFICATE_1"] = [
        x if x not in (None, "") else None
        for x in all_certificate
    ]

    df["SLIP_CLEAN_1"] = [
        x if x not in (None, "") else None
        for x in slip_cln
    ]

    insured_cols = []

    for i in range(1, max_insured_parts + 1):
        col_name = f"INSURED_CLEAN_{i}"

        df[col_name] = [
            p[i - 1] if i - 1 < len(p) else None
            for p in all_insured_parts
        ]

        insured_cols.append(col_name)

    new_columns = []

    cleaning_cols = {
        "POLIS_CLEAN_1",
        "CERTIFICATE_1",
        "SLIP_CLEAN_1",
        *insured_cols,
    }

    for col in df.columns:

        # Jangan masukkan kolom cleaning secara otomatis.
        # Kolom tersebut akan ditempatkan setelah kolom ORI masing-masing.
        if col in cleaning_cols:
            continue

        new_columns.append(col)

        if col == "POLIS_ORI":
            new_columns.append("POLIS_CLEAN_1")
            new_columns.append("CERTIFICATE_1")

        elif col == "SLIP_NO_ORI":
            new_columns.append("SLIP_CLEAN_1")

        elif col == "INSURED_ORI":
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