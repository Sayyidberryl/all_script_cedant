"""Cleansing script untuk data Bordero Marine Cargo Tokio Marine.

Membaca data raw bordero:
  data_raw/[2020, 2022 - 2026] Data Marine Cargo TOKIO M (data bulanan tidak lengkap per tahunnya).xlsx

Menghasilkan file bersih:
  data_cleaned/tokio_output_bordero.xlsx
  data_cleaned/tokio_output_bordero.xlsx.cache.pkl
"""
import os
import re
import unicodedata
import numpy as np
import pandas as pd

print("=== START CLEANSING BORDERO TOKIO MARINE ===", flush=True)

# ─────────────────────────────────────────────────────────────────────────────
# PATH KONFIGURASI
# ─────────────────────────────────────────────────────────────────────────────
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
RAW_DIR = os.path.join(BASE_DIR, "data_raw")
CLEAN_DIR = os.path.join(BASE_DIR, "data_cleaned")

RAW_FILENAME = "[2020, 2022 - 2026] Data Marine Cargo TOKIO M (data bulanan tidak lengkap per tahunnya).xlsx"
INPUT_FILE = os.path.join(RAW_DIR, RAW_FILENAME)
OUTPUT_FILE = os.path.join(CLEAN_DIR, "tokio_output_bordero.xlsx")
CACHE_FILE = OUTPUT_FILE + ".cache.pkl"

os.makedirs(CLEAN_DIR, exist_ok=True)

# ─────────────────────────────────────────────────────────────────────────────
# NORMALISASI TEKS & PEMBERSIHAN
# ─────────────────────────────────────────────────────────────────────────────
COMMON_ENTITIES = {
    "PT", "CV", "UD", "TBK", "PERSERO", "LTD", "INC", "LLC", "CO", "CORP",
    "CORPORATION", "COMPANY", "THE", "AND", "OR", "OF"
}

_CORE_M_RE = re.compile(r'([A-Z]\d{7})')


def clean_polis_str(val) -> str:
    """Bersihkan nomor polis (hilangkan karakter khusus, suffix, dll)."""
    if pd.isna(val):
        return ""
    s = str(val).strip().upper()
    if not s or s == "-":
        return ""
    # Hapus suffix seperti -000-01 atau -001-00
    s = re.sub(r'-\d{3}-\d+$', '', s)
    s = re.sub(r'-000-\d+$', '', s)
    # Hapus delimiter
    s = re.sub(r'[^A-Z0-9]', '', s)
    return s


def extract_core_m(val) -> str:
    """Ekstrak nomor core Tokio Marine (huruf M + 7 digit), misal 'M5751934'."""
    if pd.isna(val):
        return ""
    s = str(val).strip().upper()
    m = _CORE_M_RE.search(s)
    return m.group(1) if m else ""


def clean_insured_name(val) -> str:
    """Normalisasi nama insured (hapus entity legal PT/CV, tanda baca, dll)."""
    if pd.isna(val):
        return ""
    s = unicodedata.normalize("NFKC", str(val)).upper()
    # Buang karakter non alfanumerik menjadi spasi
    s = re.sub(r'[^A-Z0-9]+', ' ', s)
    tokens = [t for t in s.split() if t not in COMMON_ENTITIES]
    return " ".join(tokens)


def run_cleansing():
    if not os.path.exists(INPUT_FILE):
        raise FileNotFoundError(f"File raw bordero tidak ditemukan: {INPUT_FILE}")

    print(f"Membaca raw file: {INPUT_FILE} ...", flush=True)
    df = pd.read_excel(INPUT_FILE)
    print(f"Data raw terbaca: {len(df):,} baris, {len(df.columns)} kolom.", flush=True)

    # 1. Normalisasi nama kolom
    df.columns = [str(c).strip().upper() for c in df.columns]

    # Pastikan kolom wajib ada
    required_cols = ["BULAN", "TAHUN", "FACCODE", "POLIS", "SLIP", "CERTIFICATE", "INSURED", "CURR", "NET"]
    for rc in required_cols:
        if rc not in df.columns:
            df[rc] = np.nan

    # 2. Pembersihan POLIS & ekstraksi CORE_M
    print("Memproses kolom POLIS dan ekstraksi CORE_M ...", flush=True)
    df["POLIS"] = df["POLIS"].astype(str).replace(r'^\s*$', np.nan, regex=True)
    df["POLIS_CLEAN"] = df["POLIS"].apply(clean_polis_str)
    df["CORE_M"] = df["POLIS"].apply(extract_core_m)

    # 3. Pembersihan INSURED
    print("Memproses kolom INSURED ...", flush=True)
    df["INSURED"] = df["INSURED"].astype(str).replace(r'^\s*$', np.nan, regex=True)
    df["INSURED_CLEAN"] = df["INSURED"].apply(clean_insured_name)

    # 4. Pembersihan FACCODE & CURR
    print("Memproses kolom FACCODE & CURR ...", flush=True)
    df["FACCODE"] = df["FACCODE"].astype(str).str.strip().str.upper().replace({"NAN": np.nan, "": np.nan})
    df["CURR"] = df["CURR"].astype(str).str.strip().str.upper().replace({"NAN": np.nan, "": np.nan})

    # 5. Pembersihan NET
    print("Memproses kolom NET ...", flush=True)
    df["NET"] = pd.to_numeric(df["NET"], errors="coerce").fillna(0.0)

    # Susun kolom final
    final_cols = [
        "BULAN", "TAHUN", "FACCODE", "POLIS", "POLIS_CLEAN", "CORE_M",
        "SLIP", "CERTIFICATE", "INSURED", "INSURED_CLEAN", "CURR", "NET"
    ]
    df_out = df[final_cols].copy()

    # Simpan Excel & cache pickle
    print(f"Menyimpan hasil cleansing ke {OUTPUT_FILE} ...", flush=True)
    df_out.to_excel(OUTPUT_FILE, index=False)

    print(f"Menyimpan cache pickle ke {CACHE_FILE} ...", flush=True)
    df_out.to_pickle(CACHE_FILE)

    print(f"=== SELESAI: {len(df_out):,} baris berhasil disimpan ===", flush=True)


if __name__ == "__main__":
    run_cleansing()
