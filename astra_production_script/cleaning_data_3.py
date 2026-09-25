import os
import re

import numpy as np
import pandas as pd

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

INPUT_FILE = os.path.join(BASE_DIR, "dataExcel", "raw", "3b. Database Suspense 150826 (2).xlsx")
SHEET_NAME = "Sheet1"
OUTPUT_FILE = os.path.join(BASE_DIR, "dataExcel", "processed", "astrabuana_output_suspend.xlsx")

CEDANT_COL = "CEDANT SHRT NAME"
CEDANT_VALUE = "ASTRABUANA"

STATUS_COL = "STATUS"
STATUS_VALUE = "SUSPENSE"

ENTITIES_RE = re.compile(r"\b(PT|CV|TBK|PTE|LTD)\b", re.IGNORECASE)
PUNCT_RE = re.compile(r"[.,]")
MULTIPLE_SPACES_RE = re.compile(r"\s+")
EXT_RE = re.compile(r"-EXT\(\d+\)", re.IGNORECASE)
# ASUMSI: suffix "- New" (atau "-New", spasi bebas) di akhir SLIP NO cuma keterangan, bukan bagian nomor
SLIP_NEW_SUFFIX_RE = re.compile(r"\s*-\s*New\s*$", re.IGNORECASE)

# certificate
SD_RE = re.compile(r"s\s*/\s*d", re.IGNORECASE)
# ASUMSI: pola range ada di paling akhir SLIP NO, berupa "-angka-angka" atau "-angka SD angka"
# (dash di depan angka pertama WAJIB ada, supaya nomor lain di SLIP NO tidak ikut kebaca sebagai awal range)
CERT_RANGE_RE = re.compile(r"-(\d{1,4})\s*(?:-|SD)\s*(\d{1,4})\s*$", re.IGNORECASE)
# ASUMSI: pola single ada di paling akhir SLIP NO, berupa "-XX" (2 digit setelah dash terakhir)
CERT_SINGLE_RE = re.compile(r"-(\d{2})\s*$")


# cleaning dasar
def _clean_text_basic(text):
    if pd.isna(text):
        return ""
    text_str = PUNCT_RE.sub("", str(text))
    return MULTIPLE_SPACES_RE.sub(" ", text_str).strip()


# insured
def clean_insured(name):
    if pd.isna(name):
        return ""
    name_str = ENTITIES_RE.sub("", str(name))
    name_str = PUNCT_RE.sub("", name_str)
    return MULTIPLE_SPACES_RE.sub(" ", name_str).strip().upper()


# polis
def clean_polis(value):
    if pd.isna(value):
        return ""
    text_str = EXT_RE.sub("", str(value))
    text_str = SD_RE.sub("SD", text_str)
    return _clean_text_basic(text_str)


# slip
def clean_slip(value):
    if pd.isna(value):
        return ""
    text_str = SLIP_NEW_SUFFIX_RE.sub("", str(value))
    text_str = SD_RE.sub("SD", text_str)
    return _clean_text_basic(text_str)


# certificate
def clean_certificate(slip_no):
    if pd.isna(slip_no):
        return ""

    text = str(slip_no).strip()
    text = SD_RE.sub("SD", text)  # samakan s/d atau S/D -> SD

    # 1) coba pola RANGE dulu (angka - angka, atau angka SD angka) di akhir teks
    range_match = CERT_RANGE_RE.search(text)
    if range_match:
        start_raw, end_raw = range_match.groups()

        # pengecualian: kalau bagian akhir cuma 2 digit dan nilainya "00" -> blank
        if len(end_raw) <= 2 and end_raw.zfill(2) == "00":
            return ""

        start_num = int(start_raw)
        end_num = int(end_raw)
        if end_num < start_num:
            start_num, end_num = end_num, start_num

        jumlah_anggota = end_num - start_num + 1
        start_padded = str(start_num).zfill(6)
        end_padded = str(end_num).zfill(6)

        if jumlah_anggota > 3:
            return f"{start_padded} SD {end_padded}"
        else:
            return ", ".join(str(n).zfill(6) for n in range(start_num, end_num + 1))

    # 2) kalau bukan range, coba pola SINGLE "-XX" di paling akhir
    single_match = CERT_SINGLE_RE.search(text)
    if single_match:
        num_raw = single_match.group(1)
        if num_raw == "00":
            return ""
        return num_raw.zfill(6)

    # 3) tidak ada pola '-' + 2 digit sama sekali -> blank
    return ""


# proses data
def process_data(input_file: str, sheet_name: str, output_file: str) -> None:
    if not os.path.exists(input_file):
        print(f"[ERROR] File tidak ditemukan di lokasi: {input_file}")
        return

    df = pd.read_excel(input_file, sheet_name=sheet_name, header=2)
    df.columns = df.columns.str.strip()

    if CEDANT_COL not in df.columns:
        print(f"[ERROR] Kolom '{CEDANT_COL}' tidak ditemukan di Excel!")
        return

    if STATUS_COL not in df.columns:
        print(f"[ERROR] Kolom '{STATUS_COL}' tidak ditemukan di Excel!")
        return

    # filter baris ASTRABUANA yang statusnya SUSPENSE saja
    # (dibersihkan dulu dari spasi & beda kapital, jaga2 data mentahnya tidak konsisten)
    status_clean = df[STATUS_COL].astype(str).str.strip().str.upper()
    is_astrabuana = (df[CEDANT_COL] == CEDANT_VALUE) & (status_clean == STATUS_VALUE)

    # 1. Bersihkan Insured
    insured_cleaned = df["INSURED"].apply(clean_insured)
    df.insert(
        df.columns.get_loc("INSURED") + 1,
        "INSURED_CLEAN",
        np.where(is_astrabuana, insured_cleaned, None),
    )

    # 2. Bersihkan Polis (Sisipkan POLIS_CLN setelah POLIS)
    polis_cleaned = df["POLIS"].apply(clean_polis)
    df.insert(
        df.columns.get_loc("POLIS") + 1,
        "POLIS_CLN",
        np.where(is_astrabuana, polis_cleaned, None),
    )

    # 3. Proses Certificate (Sisipkan SERTIF_CLN_1..3 setelah POLIS_CLN)
    certificate_cleaned = df["SLIP NO"].apply(clean_certificate)
    df.insert(
        df.columns.get_loc("POLIS_CLN") + 1,
        "SERTIF_CLN_1",
        np.where(is_astrabuana, certificate_cleaned, None),
    )
    df.insert(df.columns.get_loc("SERTIF_CLN_1") + 1, "SERTIF_CLN_2", None)
    df.insert(df.columns.get_loc("SERTIF_CLN_2") + 1, "SERTIF_CLN_3", None)

    # 4. Bersihkan Slip No
    slip_cleaned = df["SLIP NO"].apply(clean_slip)
    df.insert(
        df.columns.get_loc("SLIP NO") + 1,
        "SLIP_NO_CLN",
        np.where(is_astrabuana, slip_cleaned, None),
    )

    # 5. Rename kolom agar sesuai standar: SLIP NO -> SLIP_NO
    df.rename(columns={"SLIP NO": "SLIP_NO"}, inplace=True)

    jumlah_sebelum = len(df)
    df = df[is_astrabuana].reset_index(drop=True)
    print(
        f"Memfilter baris '{CEDANT_COL}' = '{CEDANT_VALUE}' & "
        f"'{STATUS_COL}' = '{STATUS_VALUE}': {jumlah_sebelum} -> {len(df)} baris"
    )

    # rapikan penulisan STATUS (baris yg lolos filter sudah pasti SUSPENSE,
    # ini cuma menyeragamkan format teksnya, bukan mengubah maknanya)
    df[STATUS_COL] = STATUS_VALUE

    df.to_excel(output_file, index=False)
    print("Pemrosesan Data 3 Selesai!")


if __name__ == "__main__":
    process_data(INPUT_FILE, SHEET_NAME, OUTPUT_FILE)