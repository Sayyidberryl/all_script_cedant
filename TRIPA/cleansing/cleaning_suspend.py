import os
import re

import numpy as np
import pandas as pd

print("SCRIPT BERJALAN")

# ─────────────────────────────────────────────────────────────────────────────
# KONFIGURASI
# ─────────────────────────────────────────────────────────────────────────────

INPUT_FILE  = os.path.join("data_raw", "3b. Database Suspense 150826.xlsx")
OUTPUT_FILE = os.path.join("data_cleaned", "tripa_output_suspend.xlsx")
SHEET_NAME  = 0

# Filter cedant: hanya baris cedant ini yang dipertahankan.
# Sesuaikan CEDANT_COL dengan nama kolom cedant di file Suspense.
CEDANT_COL   = "CEDANT NAME"
CEDANT_VALUE = "PT ASURANSI TRIPAKARTA"

# Business rule: kalau hasil breakdown (insured/polis/slip) > 5 bagian,
# tidak usah dipecah per kolom -> digabung lagi jadi satu kolom.
MAX_BREAKDOWN = 5

# Regex untuk sub-removal dalam nama insured
_INSURED_TAIL_RE = re.compile(
    r"""
    \bAS\b\s+(?:THE\s+)?(?:PRINCIPAL|OFF-TAKER|MAINTENANCE|CONTRACTOR).*
  | \bBEING\s+(?:THE\s+)?(?:PRINCIPAL|OFF-TAKER).*
  | \bAND\s+ALL\s+SUBSIDIARI.*
  | \bINCLUDING\s+ALL\s+SUBSIDIARI.*
  | \bINCLUDING\s+ANY\s+SUBSIDIAR.*
  | \bINCLUDING\s+ANY\s+SUB\.?.*
  | \bCOMPRISING\s+OF.*
  | \bINSTALLMENT\b.*
  | \bRELATED\s+COMPANY\b.*
  | \bPURCHASED\s+OR\s+OTHERWISE\b.*
  | \bWPC\s+DATE\s*:.* 
    """,
    re.IGNORECASE | re.VERBOSE,
)

# Regex pembersih catatan mutasi / rekening internal perbankan
_BANK_NOTES_RE = re.compile(
    r"[-–/\s]*(?:rekening\s+antar\s+(?:kantor|cabang)|RAK\b).*",
    re.IGNORECASE,
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
    "MIGRASI AS400", "THE PRINCIPAL", "PRINCIPAL", "OWNER",
})

# Gelar / sapaan yang ikut dihapus dari nama insured
_INSURED_TITLE_RE = re.compile(
    r"\b(?:BAPAK|IBU|BPK|NY\.?|SDR\.?|SDRI\.?|TN\.?|MR\.?|MRS\.?|MS\.?)\b\s*",
    re.IGNORECASE,
)


# ─────────────────────────────────────────────────────────────────────────────
# CLEAN FUNCTIONS
# ─────────────────────────────────────────────────────────────────────────────

def _cap_breakdown(parts: list, sep: str = "; ") -> list:
    """
    Business rule: jika jumlah bagian hasil breakdown > MAX_BREAKDOWN,
    jangan dipecah per kolom -> gabungkan lagi jadi satu nilai.
    """
    parts = [p for p in parts if p]
    if len(parts) > MAX_BREAKDOWN:
        return [sep.join(parts)]
    return parts


def _strip_polis_suffix(cur: str) -> str:
    """
    Ambil nomor polis induk saja: buang suffix fraksi & certificate di belakang.

    Contoh:
    - 10703082400154-000001           -> 10703082400154
    - 11112032500133-1/0              -> 11112032500133
    - 11202012500246-000001-1/1       -> 11202012500246
    - 10202012500292 000002 2/0       -> 10202012500292
    - 10302012400168 000003 4/1       -> 10302012400168
    - 10202012500292-000001 S/D 000005 -> 10202012500292
    """
    cur = cur.strip().upper()

    # Buang suffix / VAR / VARIOUS / TBA
    cur = re.sub(r"\s*/\s*(?:VAR|VARIOUS|TBA).*$", "", cur).strip()

    # Range S/D -> ambil bagian kiri saja
    cur = re.split(r"\s+S\s*/?\s*D\s+", cur, maxsplit=1)[0].strip()

    # 1. Suffix fraksi di belakang (pemisah dash / spasi): -1/1, " 2/0", " 4/1"
    cur = re.sub(r"[\s-]+\d+/\d+$", "", cur).strip()

    # 2. Suffix certificate dengan dash: -000001, -01
    cur = re.sub(r"-\d+$", "", cur).strip()

    # 3. Suffix certificate dengan spasi: " 000002" (hanya jika didahului
    #    nomor polis >= 10 digit, supaya nomor lain tidak ikut terpotong)
    cur = re.sub(r"(?<=\d{10})\s+\d{1,6}$", "", cur).strip()

    return cur


def clean_polis(val) -> list:
    """
    Clean nomor polis TRIPA.

    Rule:
    - 10203082500560 -> tetap
    - 11112032500133-1/0 -> 11112032500133
    - 11202012500246-000001-1/1 -> 11202012500246
    - 10202012500292 000002 2/0 -> 10202012500292
    """
    if pd.isna(val):
        return []

    val = str(val).strip()
    if not val or val == "-":
        return []

    val = val.replace(",", "")
    raw_parts = [p.strip() for p in val.split("+") if p.strip()]
    cleaned_parts = []

    for p in raw_parts:
        cur = _strip_polis_suffix(p)

        # Hapus semua separator dari hasil clean.
        cur = cur.replace(".", "")
        cur = cur.replace("/", "")
        cur = cur.replace("-", "")
        cur = cur.replace(" ", "")

        if cur:
            cleaned_parts.append(cur)

    return _cap_breakdown(cleaned_parts)


def clean_Certificate(polis_ori):
    """
    Mengambil Certificate dari POLIS ORI.
    
    Menangani pola:
    - 10602012500418-000001-1/1  -> 000001
    - 10302012400168 000003 4/1  -> 000003
    - Range S/D, / VARIOUS, dll.
    """
    if pd.isna(polis_ori):
        return ""

    polis = str(polis_ori).strip().upper()

    if not polis:
        return ""

    # Hapus data kotor / suspense
    if re.search(r"\bSUSPENSE\b|\bENDORSEMENT\b", polis):
        return ""

    # 1. Hapus suffix fraksi di paling belakang seperti -1/1, -2/1, 1/1, 4/1, dst.
    polis_clean = re.sub(r"[\s-]+\d+/\d+$", "", polis).strip()

    # 2. Hapus suffix / VAR / VARIOUS / TBA di bagian belakang
    polis_clean = re.sub(
        r"\s*/\s*(?:VAR|VARIOUS|TBA).*$",
        "",
        polis_clean,
        flags=re.IGNORECASE,
    ).strip()

    # 3. Handling Pola RANGE S/D
    if re.search(r"\bS\s*/?\s*D\b", polis_clean):
        parts_sd = re.split(r"\s*S\s*/?\s*D\s*", polis_clean)
        if len(parts_sd) == 2:
            left, right = parts_sd[0].strip(), parts_sd[1].strip()

            m_left = re.search(r"(?:^|[\s-])(\d{1,6})$", left)
            cert_left = m_left.group(1).zfill(6) if m_left else ""

            m_right = re.search(r"(?:^|[\s-])(\d{1,6})$", right)
            cert_right = m_right.group(1).zfill(6) if m_right else ""

            if cert_left and cert_right:
                return f"{cert_left} SD {cert_right}"

    # 4. Handling Single Certificate (Pemisah Dash/Tanda Hubung)
    # Contoh: 10602012500418-000001
    m_single_dash = re.search(r"-(\d{1,6})$", polis_clean)
    if m_single_dash:
        return m_single_dash.group(1).zfill(6)

    # 5. Handling Single Certificate (Pemisah Spasi)
    # Contoh: 10302012400168 000003
    m_single_space = re.search(r"(?<=\s)(\d{1,6})$", polis_clean)
    if m_single_space:
        return m_single_space.group(1).zfill(6)

    return ""

def clean_slip(val) -> list:
    """
    Clean nomor slip TRIPA.

    Rule:
    - AF-2024-01 -> tetap AF-2024-01
    - AF-2025-01 -> tetap AF-2025-01
    - Nomor slip biasa -> tetap
    - Jika ada beberapa slip dengan + -> tetap dipisah
    """
    if pd.isna(val):
        return []

    val = str(val).strip()

    if not val or val == "-":
        return []

    raw_parts = [
        p.strip()
        .replace(".", "")
        .replace("/", "")
        for p in val.split("+")
        if p.strip()
    ]

    cleaned_parts = []

    for p in raw_parts:
        p = p.strip()

        # AUTO FAC: Jangan hapus tanda "-"
        if re.fullmatch(r"AF-\d{4}-\d{2}", p, flags=re.IGNORECASE):
            cleaned_parts.append(p.upper())
            continue

        p = p.replace("-", "")

        if p:
            cleaned_parts.append(p)

    return _cap_breakdown(cleaned_parts)


def _remove_polis_slip_from_text(text: str, polis_ori, slip_ori) -> str:
    """Hapus nomor polis & slip dari teks insured secara agresif."""
    if pd.notna(polis_ori):
        for token in [str(polis_ori).strip()] + clean_polis(polis_ori):
            if token and token != "-":
                text = text.replace(token, "")

    if pd.notna(slip_ori):
        for token in [str(slip_ori).strip()] + clean_slip(slip_ori):
            if token and token != "-":
                text = text.replace(token, "")

    return text


def _normalize_insured_part(p: str) -> str:
    """Terapkan sub-removal dan normalisasi pada satu bagian nama insured."""
    p = _BANK_NOTES_RE.sub("", p)
    p = _INSURED_TAIL_RE.sub("", p)
    p = _INSURED_TITLE_RE.sub("", p)
    p = re.sub(r"\bKB\b", "", p, flags=re.IGNORECASE)
    p = re.sub(r"\bA\.?W\.?\b", "", p, flags=re.IGNORECASE)
    p = re.sub(r"\(\s*\)", "", p)
    p = re.sub(r"\*\)", "", p)  
    p = p.replace("*", "")

    # Trim leading/trailing AND/OR
    p = re.sub(r"^(?:AND|OR)\b\s*", "", p, flags=re.IGNORECASE)
    p = re.sub(r"\s*\b(?:AND|OR)$", "", p, flags=re.IGNORECASE)

    # Trim leading/trailing non-alphanumeric
    p = re.sub(r"^[^a-zA-Z0-9(]+", "", p)
    p = re.sub(r"[^a-zA-Z0-9)]+$", "", p)

    return p.strip()


def _is_valid_insured_part(p: str) -> bool:
    """Return True jika bagian insured layak dipertahankan."""
    if len(p) <= 2:
        return False

    up = p.upper()

    if re.match(r"^[\d\/\-\.]+$", up):
        return False

    if re.search(
        r"\b(?:NO\.\s*\d+|BUILDING|FLOOR|ROOM|ROAD|STREET|TOWER|KAV\.?|BLOK)\b",
        up,
    ):
        return False

    if re.search(
        r"\b(?:PLTGU|PLTMH|PLTU|POWER PLANT|COMBINED CYCLE|MW|HYDRO ELECTRIC)\b",
        up,
    ):
        return False

    return up not in _INSURED_JUNK_WORDS


def clean_insured(val, polis_ori, slip_ori) -> list:
    """Bersihkan insured berdasarkan rule TRIPA."""

    if pd.isna(val):
        return []

    val = str(val).strip()
    if not val:
        return []

    up = re.sub(r"\s+", " ", val).strip().upper()

    # AUTO FAC: jangan di-split karena / dan - merupakan bagian pola.
    if up.startswith("AUTO FAC "):
        return [up]

    # BNI berdiri sendiri = 1 insured.
    if re.fullmatch(
        r"PT\.?\s*BANK\s+NEGARA\s+INDONESIA\s*"
        r"(?:\(PERSERO\))?\s*TBK\.?",
        up,
        flags=re.IGNORECASE,
    ):
        return ["BANK NEGARA INDONESIA"]

    # CV AP CARGO ... QQ ... = 2 insured.
    if "CV AP CARGO NUSANTARA" in up and re.search(r"\bQQ\b", up):
        parts = re.split(r"\bQQ\b", up, maxsplit=1, flags=re.IGNORECASE)
        cleaned = []

        first = _normalize_insured_part(parts[0])
        first = re.sub(r"^CV\.?\s*", "", first, flags=re.IGNORECASE).strip()
        if _is_valid_insured_part(first):
            cleaned.append(first.upper())

        second = _normalize_insured_part(parts[1])
        if _is_valid_insured_part(second):
            cleaned.append(second.upper())

        return _cap_breakdown(cleaned)

    # BRI ... QQ ... = 1 insured: ambil insured utama sebelum QQ.
    if "PT BANK BRI KANTOR CABANG YOGYAKARTA KATAMSO" in up:
        first = re.split(r"\bQQ\b", up, maxsplit=1, flags=re.IGNORECASE)[0]
        first = re.sub(r"^PT\.?\s*", "", first, flags=re.IGNORECASE).strip()
        return [re.sub(r"\s+", " ", first)]

    # General cleaning
    val = _remove_polis_slip_from_text(val, polis_ori, slip_ori)
    val = _BANK_NOTES_RE.sub("", val)
    val = _INSURED_TAIL_RE.sub("", val)
    val = re.sub(r"\([^)]*\)", "", val)

    val = re.sub(
        r"\b(?:AND|AN|OR)\s*/\s*(?:AND|OR)\b",
        ",",
        val,
        flags=re.IGNORECASE,
    )
    val = re.sub(r"\bAND\s+OR\b", ",", val, flags=re.IGNORECASE)
    val = re.sub(r"\bCO\.,?\s*LTD\.?\b", ",", val, flags=re.IGNORECASE)

    for pattern in [
        r"\bTBK\.?\b",
        r"\(PERSERO\)",
        r"\bPERSERO\b",
        r"\bLTD\.?\b",
        r"\(FCI\.\s*I\.\)",
    ]:
        val = re.sub(pattern, "", val, flags=re.IGNORECASE)

    split_pattern = (
        r"\bQQ\b|/|,|&|\+"
        r"|-(?!\s*(?:19|20)\d{2}\b)"
        r"|\d+\."
        r"|\bPT\.?\b"
        r"|\bCV\.?\b"
        r"|:"
        r"|;"
    )

    parts = re.split(split_pattern, val, flags=re.IGNORECASE)
    cleaned = []

    for p in parts:
        p = _normalize_insured_part(p)
        if _is_valid_insured_part(p):
            cleaned.append(p.upper())

    return _cap_breakdown(cleaned)


# ─────────────────────────────────────────────────────────────────────────────
# COLUMN EXPANSION HELPER
# ─────────────────────────────────────────────────────────────────────────────

def _expand_clean_columns(
    df: pd.DataFrame,
    all_lists: list,
    prefix: str,
    max_cols: int,
) -> list:
    """
    Tambahkan kolom clean_{prefix}_1 .. N langsung setelah kolom ori.
    max_cols otomatis dibatasi MAX_BREAKDOWN karena _cap_breakdown()
    sudah menggabungkan hasil > MAX_BREAKDOWN jadi 1 kolom.
    Returns daftar nama kolom baru yang ditambahkan.
    """
    added = []

    for i in range(1, max_cols + 1):
        col_name = f"{prefix}" if max_cols == 1 else f"{prefix}_{i}"
        df[col_name] = [
            lst[i - 1] if i - 1 < len(lst) else None
            for lst in all_lists
        ]
        added.append(col_name)

    return added


# ─────────────────────────────────────────────────────────────────────────────
# PROCESS DATA
# ─────────────────────────────────────────────────────────────────────────────

def process_data(input_file: str, output_file: str) -> None:
    print(f"[1/5] Membaca data dari: {input_file} (sheet: {SHEET_NAME}) ...")
    df = pd.read_excel(input_file, sheet_name=SHEET_NAME, header=2)
    print(f"      Total baris keseluruhan: {len(df):,}")

    df = pd.read_excel(input_file, sheet_name=SHEET_NAME, header=2)
    print(f"      Total baris keseluruhan: {len(df):,}")

    # ========================================================
    # FILTER STATUS - HANYA SUSPENSE
    # ADJUSTED DIBUANG PERMANEN
    # ========================================================

    if "STATUS" not in df.columns:
        raise ValueError("Kolom STATUS tidak ditemukan di data.")

    total_sebelum = len(df)

    df = df[
        df["STATUS"]
        .fillna("")
        .astype(str)
        .str.strip()
        .str.upper()
        .eq("SUSPENSE")
    ].copy()

    total_sesudah = len(df)
    total_adjusted = total_sebelum - total_sesudah

    print(f"      Data SUSPENSE : {total_sesudah:,} baris")
    print(f"      Data ADJUSTED : {total_adjusted:,} baris dibuang")


    # ========================================================
    # FILTER CEDANT - HANYA TRIPA
    # ========================================================
    if CEDANT_COL not in df.columns:
        raise ValueError(
            f"Kolom '{CEDANT_COL}' tidak ditemukan. "
            f"Kolom tersedia: {list(df.columns)}"
        )

    def _norm_cedant(x) -> str:
        x = str(x).upper().replace(".", " ")
        return re.sub(r"\s+", " ", x).strip()

    total_sebelum_cedant = len(df)
    df = df[
        df[CEDANT_COL].fillna("").map(_norm_cedant).eq(_norm_cedant(CEDANT_VALUE))
    ].copy()
    print(
        f"      Cedant '{CEDANT_VALUE}': {len(df):,} baris "
        f"({total_sebelum_cedant - len(df):,} baris cedant lain dibuang)"
    )

    print(f"[2/5] Total baris yang akan diproses: {len(df):,}")

    print(
        "[3/5] Menjalankan proses cleaning untuk rule TRIPA ..."
    )

    all_clean_polis = []
    all_clean_slip = []
    all_clean_ins = []
    all_Certificates = []

    max_polis = 1
    max_slip = 1
    max_ins = 1

    for idx, (_, row) in enumerate(df.iterrows()):
        p_ori = row.get("POLIS", "")
        s_ori = row.get("SLIP NO", "")
        i_ori = row.get("INSURED", "")

        c_polis = clean_polis(p_ori)
        c_slip = clean_slip(s_ori)
        c_ins = clean_insured(i_ori, p_ori, s_ori)
        c_Certificate = clean_Certificate(p_ori)

        max_polis = max(max_polis, len(c_polis))
        max_slip = max(max_slip, len(c_slip))
        max_ins = max(max_ins, len(c_ins))

        all_clean_polis.append(c_polis)
        all_clean_slip.append(c_slip)
        all_clean_ins.append(c_ins)
        all_Certificates.append(c_Certificate)

    # Safety net
    max_polis = min(max_polis, MAX_BREAKDOWN)
    max_slip = min(max_slip, MAX_BREAKDOWN)
    max_ins = min(max_ins, MAX_BREAKDOWN)

    print("[4/5] Menyusun kolom output ...")

    # Masukkan kolom Certificate ke DataFrame
    df["CERTIFICATE_1"] = all_Certificates

    new_columns = []
    for col in df.columns:
        if col == "CERTIFICATE_1":
            continue

        if col == "POLIS":
            new_columns.append(col)

            clean_polis_columns = _expand_clean_columns(
                df,
                all_clean_polis,
                "POLIS_CLEAN",
                max_polis,
            )

            if clean_polis_columns:
                new_columns.append(clean_polis_columns[0])
                new_columns.append("CERTIFICATE_1")
                new_columns += clean_polis_columns[1:]
            else:
                new_columns.append("CERTIFICATE_1")

        elif col == "SLIP NO":
            new_columns.append(col)
            new_columns += _expand_clean_columns(
                df,
                all_clean_slip,
                "SLIP_NO_CLEAN",
                max_slip,
            )

        elif col == "INSURED":
            new_columns.append(col)
            new_columns += _expand_clean_columns(
                df,
                all_clean_ins,
                "INSURED_CLEAN",
                max_ins,
            )

        else:
            new_columns.append(col)

    df = df[new_columns]

    print(f"[5/5] Menyimpan hasil ke: {output_file} ...")

    os.makedirs(os.path.dirname(output_file), exist_ok=True)
    df.to_excel(output_file, index=False)

    print(f"\n{'=' * 55}")
    print("  [OK] Selesai!")
    print(f"  Total baris output  : {len(df):,}")
    print(f"  Total kolom output  : {len(df.columns)}")
    print(f"  File disimpan di    : {output_file}")
    print(f"{'=' * 55}")


# ─────────────────────────────────────────────────────────────────────────────
# ENTRY POINT
# ─────────────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    print("MASUK MAIN")
    process_data(INPUT_FILE, OUTPUT_FILE)