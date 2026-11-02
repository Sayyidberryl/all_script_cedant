import os
import re
from typing import Any, List

import numpy as np
import pandas as pd


# KONFIGURASI UMUM
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

INPUT_FILE = os.path.join(BASE_DIR, "data_excel", "raw", "3b. Database Suspense 150826.xlsx")
SHEET_NAME = "Sheet1"
OUTPUT_FILE = os.path.join(BASE_DIR, "data_excel", "processed", "mitrautama_output_suspend.xlsx")

CEDANT_COL = "CEDANT SHRT NAME"
CEDANT_VALUE = "MITRAUTAMA"

# Tambahan konfigurasi untuk kolom status
STATUS_COL = "STATUS" 

MAX_HASIL_BREAKDOWN = 5


# INSURED
MULTIPLE_SPACES_RE = re.compile(r"\s+")

ENTITY_RE = re.compile(
    r"\(?\s*\b(PT|CV|TBK|PTE|LTD|PERSERO|SDN|BHD|KSO|KUD|HOUSING|ESTATE|MILL|HOLDING|FACTORY)\b\s*\.?\s*\)?",
    re.IGNORECASE,
)

INSURED_PUNCT_RE = re.compile(r"[.,\-/\"]")

INSURED_PHRASE_REMOVE = [
    "Perhimpunan Pemilik dan Penghuni Satuan Rumah Susun (P3SRS)",
    "PERHIMPUNAN PEMILIK DAN PENGHUNI SATUAN RUMAH SUSUN (PPPSRS)"
]

INSURED_PHRASE_RE = re.compile(
    r"(?:" + "|".join(re.escape(p) for p in INSURED_PHRASE_REMOVE) + r")",
    re.IGNORECASE,
)

ENTITY_START_RE = re.compile(
    r"^\.?\s*\(?\s*(PT|CV|TBK|PTE|LTD|PERSERO|SDN|BHD)\b", re.IGNORECASE
)

# Pemisah "keras": selalu memecah segmen (QQ, ANDOR, AND/OR, AND OR)
HARD_SEP_RE = re.compile(
    r"\bQQ\b|\bANDOR\b|\bAND\s*/\s*OR\b|\bAND\s+OR\b",
    re.IGNORECASE,
)

SEP_TOKEN_RE = re.compile(
    r"(?P<space_before>\s*)(?P<sep>,|-|&|\bQQ\b|\bANDOR\b|\bAND\s*/\s*OR\b|\bAND\s+OR\b)(?P<space_after>\s*)",
    re.IGNORECASE,
)

PAREN_RE = re.compile(r"\([^)]*\)")


def _mask_parens(text: str):
    stash: List[str] = []

    def _repl(m: "re.Match[str]") -> str:
        stash.append(m.group(0))
        return f"\x00{len(stash) - 1}\x00"

    return PAREN_RE.sub(_repl, text), stash


def _unmask_parens(text: str, stash: List[str]) -> str:
    def _repl(m: "re.Match[str]") -> str:
        return stash[int(m.group(1))]

    return re.sub(r"\x00(\d+)\x00", _repl, text)


def _split_insured_segments(raw: str) -> List[str]:
    protected, stash = _mask_parens(raw)

    segments: List[str] = []
    current = ""
    last_end = 0

    for m in SEP_TOKEN_RE.finditer(protected):
        current += protected[last_end:m.start()]
        sep = m.group("sep")
        is_hard = bool(HARD_SEP_RE.fullmatch(sep))

        remainder = protected[m.end():].lstrip()
        starts_new_entity = bool(ENTITY_START_RE.match(remainder))

        if is_hard or starts_new_entity:
            segments.append(current)
            current = ""
        elif sep == ",":
            current += " "
        elif sep == "-":
            if m.group("space_before") and m.group("space_after"):
                current += " "
            else:
                current += m.group(0)
        else: 
            current += m.group(0)

        last_end = m.end()

    current += protected[last_end:]
    segments.append(current)

    return [_unmask_parens(seg, stash) for seg in segments]


def _clean_insured_segment(segment: str) -> str:
    s = INSURED_PHRASE_RE.sub(" ", segment)
    s, stash = _mask_parens(s)
    s = ENTITY_RE.sub(" ", s)
    s = _unmask_parens(s, stash)
    
    # Hapus tanda baca (termasuk '-')
    s = INSURED_PUNCT_RE.sub(" ", s)
    return MULTIPLE_SPACES_RE.sub(" ", s).strip().upper()


def breakdown_insured(value: Any) -> List[Any]:
    if pd.isna(value):
        return [value]

    raw = str(value).strip()
    if not raw:
        return [raw]

    segments = _split_insured_segments(raw)
    cleaned_parts = [_clean_insured_segment(s) for s in segments]
    cleaned_parts = [p for p in cleaned_parts if p]

    if not cleaned_parts:
        return [raw]

    if len(cleaned_parts) > MAX_HASIL_BREAKDOWN:
        return [raw]

    return cleaned_parts


# POLIS
EXT_RE = re.compile(r"-EXT\(\d+\)", re.IGNORECASE)
POLIS_SUFFIX_RE = re.compile(r"\s+[A-Za-z]+$")
POLIS_COMPLETE_RE = re.compile(r"^[A-Za-z0-9]{14}$")

# Hapus koma dari regex karena koma ditangani saat memecah string
POLIS_PUNCT_RE = re.compile(r"[./+\-]")


def _clean_polis_punct(text: str) -> str:
    return POLIS_PUNCT_RE.sub("", text).strip()


def breakdown_polis(value: Any) -> List[Any]:
    if pd.isna(value):
        return [value]

    raw_ori = str(value).strip()
    if not raw_ori:
        return [raw_ori]
        
    # Rule 1: Jika hanya "-", biarkan tetap "-"
    if raw_ori == "-":
        return ["-"]

    # Rule 2: Hapus karakter "Â" di awal sebelum memproses
    text = raw_ori.replace("Â", "").strip()

    # Eksekusi regex bawaan untuk hapus EXT atau suffix lainnya
    text = EXT_RE.sub("", text)
    text = POLIS_SUFFIX_RE.sub("", text).strip()

    # Pisahkan dulu dengan ';' sebagai pembatas blok paling luar (Rule 6)
    blocks = [b.strip() for b in text.split(";") if b.strip()]
    hasil_sementara = []
    
    for block in blocks:
        if "&" in block:
            # Rule 4: Breakdown berdasarkan '&'
            parts = [p.strip() for p in block.split("&")]
            hasil_sementara.extend(parts)
        elif "," in block:
            # Breakdown berdasarkan koma
            parts = [p.strip() for p in block.split(",")]
            base_polis = parts[0]
            hasil_sementara.append(base_polis)
            
            for part in parts[1:]:
                # Rule 6: Jika setelah koma hanya 3 angka (shorthand)
                if part.isdigit() and len(part) <= 5: 
                    # Ganti angka belakang sesuai panjang shorthand-nya
                    new_polis = base_polis[:-len(part)] + part
                    hasil_sementara.append(new_polis)
                else:
                    # Rule 2 & 3: Breakdown sebagai polis normal jika bukan shorthand
                    hasil_sementara.append(part)
        elif " " in block:
            # Rule 5: Breakdown berdasarkan spasi
            parts = [p.strip() for p in block.split(" ")]
            hasil_sementara.extend(parts)
        else:
            # Polis tunggal dalam blok ini
            hasil_sementara.append(block)

    # Bersihkan tanda baca termasuk "-" untuk setiap polis hasil breakdown (Rule 3)
    hasil_akhir = []
    for p in hasil_sementara:
        p_clean = _clean_polis_punct(p)
        if p_clean:
            hasil_akhir.append(p_clean)

    # Validasi limit sesuai aturan MAX_HASIL_BREAKDOWN
    if len(hasil_akhir) > MAX_HASIL_BREAKDOWN:
        return [raw_ori]
        
    return hasil_akhir if hasil_akhir else [raw_ori]


# SLIP
SLIP_PUNCT_RE = re.compile(r"[./+\-]")


def clean_slip(value: Any) -> Any:
    if pd.isna(value):
        return ""
    return SLIP_PUNCT_RE.sub("", str(value)).strip()


# CERTIFICATE
# Diambil dari kolom POLIS, mencari pola angka setelah "-" atau "/" di paling akhir string.
# SYARAT PENTING: token terakhir setelah "-" atau "/" itu HARUS TEPAT 2 digit angka.
# Kalau token terakhir bukan tepat 2 digit (lebih pendek/panjang, misal 4 digit tahun
# atau 6 digit nomor urut), maka itu BUKAN pola certificate -> dikosongkan.
# Contoh: "0101-0109-24-000704" -> token terakhir "000704" (6 digit) -> BUKAN certificate -> blank
#         "011/CS/CAR/MUR31/DH/05/2023" -> token terakhir "2023" (4 digit) -> BUKAN certificate -> blank
#
# - Kalau token terakhir persis 2 digit -> dipadding jadi 6 digit (tambah 0 di depan).
#   Kecuali kalau nilainya "00" -> dikosongkan.
# - Range: kalau token SEBELUM token terakhir itu JUGA persis 2 digit (dipisah "-", "/",
#   atau "s/d"/"S/D"), maka dianggap range dari token sebelum -> token terakhir:
#     * jika jumlah anggota range (end - start + 1) > 3  -> ditulis "AWAL SD AKHIR"
#     * jika jumlah anggota range <= 3                    -> ditulis "AWAL, AKHIR, ..." (list lengkap, koma)
#   Kalau token sebelum terakhir BUKAN 2 digit (mis. "12345-12"), maka bukan range,
#   dianggap certificate tunggal dari token terakhir saja.
# - Penulisan "s/d" / "S/D" pada data disamakan menjadi "SD" pada output.
#
# CATATAN: kalau ternyata masih ada pola POLIS lain di data asli yang belum sesuai
# dengan asumsi ini, kirim contohnya supaya regex bisa disesuaikan lagi.

CERT_RANGE_RE = re.compile(
    r"[-/]\s*(\d{2})\s*(?:[-/]|s\s*/?\s*d)\s*(\d{2})\s*$",
    re.IGNORECASE,
)

CERT_SINGLE_RE = re.compile(
    r"[-/]\s*(\d{2})\s*$"
)


def generate_certificate(value: Any) -> str:
    if pd.isna(value):
        return ""

    text = str(value).strip()
    if not text:
        return ""

    # Coba pola range dulu: dua angka TEPAT 2 digit di akhir, dipisah "-"/"/"/"s/d"
    m_range = CERT_RANGE_RE.search(text)
    if m_range:
        start = int(m_range.group(1))
        end = int(m_range.group(2))

        # kalau kebalik (end < start) tetap urutkan naik
        if end < start:
            start, end = end, start

        if start == 0 and end == 0:
            return ""

        count = end - start + 1
        if count > 3:
            return f"{str(start).zfill(6)} SD {str(end).zfill(6)}"
        else:
            return ", ".join(str(n).zfill(6) for n in range(start, end + 1))

    # Kalau bukan range, cek angka TEPAT 2 digit di paling akhir setelah "-" atau "/"
    m_single = CERT_SINGLE_RE.search(text)
    if m_single:
        num_raw = m_single.group(1)
        if int(num_raw) == 0:
            return ""
        return num_raw.zfill(6)

    # Token terakhir bukan tepat 2 digit (atau tidak ada "-"/"/" sama sekali) -> kosongkan
    return ""


# PIPELINE UTAMA
def _insert_breakdown_columns(
    df: pd.DataFrame,
    source_col: str,
    fn_breakdown: callable,
    prefix: str,
    mask: pd.Series,
) -> None:
    hasil_breakdown = df.apply(
        lambda row: fn_breakdown(row[source_col]) if mask[row.name] else [],
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

    posisi = df.columns.get_loc(source_col)
    for i, col in enumerate(df_breakdown.columns):
        df.insert(posisi + 1 + i, col, df_breakdown[col])


def process_data(input_file: str, sheet_name: str, output_file: str) -> None:
    if not os.path.exists(input_file):
        print(f"[ERROR] File tidak ditemukan di lokasi: {input_file}")
        return

    df = pd.read_excel(input_file, sheet_name=sheet_name, header=2)
    df.columns = df.columns.str.strip()

    # --- FILTER STATUS (SUSPENSE ONLY) ---
    if STATUS_COL in df.columns:
        print(f"[INFO] Memeriksa kolom status '{STATUS_COL}'...")
        # Normalisasi ke lowercase untuk pengecekan aman
        status_series = df[STATUS_COL].astype(str).str.strip().str.lower()
        
        # Ambil hanya yang "suspense" (otomatis mengabaikan "adjusted" dan yang lainnya)
        df = df[status_series == "suspense"].copy()
        
        if df.empty:
            print("[PERINGATAN] Tidak ada data dengan status 'suspense'. Proses dihentikan.")
            return
            
        print(f"[INFO] Sisa data setelah filter 'suspense': {len(df)} baris")
    else:
        print(f"[PERINGATAN] Kolom '{STATUS_COL}' tidak ditemukan di dataset mentah. Lewati tahap filter.")
    # -------------------------------------

    if CEDANT_COL not in df.columns:
        print(f"[ERROR] Kolom '{CEDANT_COL}' tidak ditemukan di Excel!")
        return

    # normalisasi nilai kolom cedant (hapus spasi berlebih, samakan huruf besar/kecil)
    # supaya perbandingan tidak gagal hanya karena "Mitrautama ", "MITRA UTAMA", dsb.
    cedant_normalized = (
        df[CEDANT_COL]
        .astype(str)
        .str.strip()
        .str.upper()
        .str.replace(r"\s+", "", regex=True)
    )
    cedant_target = CEDANT_VALUE.strip().upper().replace(" ", "")

    is_mitrautama = cedant_normalized == cedant_target

    print(f"[DEBUG] Nilai unik di kolom '{CEDANT_COL}':")
    print(df[CEDANT_COL].astype(str).str.strip().unique())
    print(f"[DEBUG] Jumlah baris cocok dengan '{CEDANT_VALUE}': {is_mitrautama.sum()} dari {len(df)} baris")

    # --- FILTER DATAFRAME HANYA UNTUK MITRAUTAMA ---
    df = df[is_mitrautama].copy()

    if df.empty:
        print(
            "[PERINGATAN] Tidak ada baris yang cocok dengan CEDANT_VALUE. "
            "Cek daftar nilai unik di atas, lalu sesuaikan CEDANT_VALUE jika perlu. Proses dihentikan."
        )
        return

    print("Menjalankan pembersihan dan breakdown kolom...")
    
    # Karena DataFrame sudah difilter, semua baris pasti MITRAUTAMA.
    # Kita buat mask berisi True untuk semua sisa baris
    mask_all_true = pd.Series(True, index=df.index)

    _insert_breakdown_columns(df, "INSURED", breakdown_insured, "INSURED_CLEAN", mask_all_true)

    # CERTIFICATE dihitung dari nilai POLIS asli (sebelum breakdown),
    # jadi harus diambil sebelum kolom POLIS dipecah jadi POLIS_CLEAN_1, dst.
    certificate_values = df["POLIS"].apply(generate_certificate)

    # 1. Jalankan breakdown POLIS terlebih dahulu
    _insert_breakdown_columns(df, "POLIS", breakdown_polis, "POLIS_CLEAN", mask_all_true)

    # 2. Sisipkan CERTIFICATE_1 setelah POLIS_CLEAN_1
    nama_kolom_polis_clean_1 = "POLIS_CLEAN_1"
    
    if nama_kolom_polis_clean_1 in df.columns:
        # Jika kolom POLIS_CLEAN_1 ada, sisipkan tepat di kanannya
        posisi_polis_clean = df.columns.get_loc(nama_kolom_polis_clean_1)
        df.insert(posisi_polis_clean + 1, "CERTIFICATE_1", certificate_values)
    else:
        # Fallback: jika tidak ada hasil breakdown, sisipkan setelah POLIS utama
        posisi_polis = df.columns.get_loc("POLIS")
        df.insert(posisi_polis + 1, "CERTIFICATE_1", certificate_values)

    slip_cleaned = df["SLIP NO"].apply(clean_slip)
    df.insert(
        df.columns.get_loc("SLIP NO") + 1,
        "SLIP_NO_CLEAN",
        slip_cleaned,
    )

    # pastikan kolom STATUS seragam ke "SUSPENSE" di file akhir
    df["STATUS"] = "SUSPENSE"

    os.makedirs(os.path.dirname(output_file), exist_ok=True)
    df.to_excel(output_file, index=False)
    print("Pemrosesan Data 3 Selesai!")


if __name__ == "__main__":
    process_data(INPUT_FILE, SHEET_NAME, OUTPUT_FILE)