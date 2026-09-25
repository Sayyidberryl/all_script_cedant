# Production Script — Suspend Matching (ACA)

Pipeline rekonsiliasi data **SUSPEND** ke tabel referensi **OSBAL**, **R/I Slip DB**, dan **FACUL** untuk cedant PT. Asuransi Central Asia (ACA).

---

## Struktur Direktori

```
production_script/
├── main.py                    # Master orchestrator (menjalankan seluruh pipeline berurutan)
├── cleaning_facul.py          # Step 1: Cleaning data FACUL
├── cleaning_osbal.py          # Step 2: Cleaning data OSBAL
├── cleaning_suspend.py        # Step 3: Cleaning data Suspend
├── prod_sus_1.py              # Step 4: Suspend Matching v1 (Row Asli + Flagging)
├── prod_sus_2.py              # Step 5: Suspend Matching v2 (Akumulasi)
├── alter.py                   # Step 6: Analisis Tambahan & Output Final -> hanya untuk developer meriview hasil (tidak digunakan untuk produksi)
├── excel_styler.py            # Helper styling & penataan warna Excel output
│
├── data/
│   ├── suspend.xlsx           # Input: raw suspend data (header di baris ke-3)
│   ├── osbal.xlsx             # Input: OS Balance cedant ACA
│   ├── facul.xlsx             # Input: FACUL cedant ACA
│   ├── slipdb.xlsx            # Input: R/I Slip DB (opsional)
│   ├── suspend_clean_aca.xlsx # Output cleaning: suspend
│   ├── osbal_clean_aca.xlsx   # Output cleaning: OSBAL
│   ├── facul_clean_aca.xlsx   # Output cleaning: FACUL
│   ├── slipdb_clean_aca.xlsx  # Output cleaning: SLIPDB (opsional)
│   ├── final_output_v1.xlsx   # Output matching v1
│   ├── final_output_v2.xlsx   # Output matching v2
│   └── final_output.xlsx      # Output akhir pipeline matching & analisis
│
├── debugging/                 # Folder khusus script debugging, diagnostik & testing
│   ├── debug_trace_suspend.py # Diagnostic: trace step-by-step 1 baris Suspend
│   ├── debug_trace.py         # Trace matching dasar
│   ├── debug_24FAS9P5.py      # Debug spesifik polis / facode 24FAS9P5
│   ├── deep_dive_facode.py    # Deep dive analisis korelasi facode
│   ├── investigate_facode.py  # Investigasi pencarian facode
│   ├── mass_sample_facode.py  # Mass sampling verifikasi facode
│   ├── check_slip_raw.py      # Cek keberadaan slip suspend di OSBAL raw
│   ├── check_suspend_fields.py# Cek kelengkapan field suspend
│   ├── check_various.py       # Cek variasi data
│   └── test_clean_polis_plus.py# Unit test logika plus suffix & clean polis
│
└── sas_ready/                 # Standalone script set untuk lingkungan SAS
    └── main.py
```

> File `.cache.pkl` dibuat otomatis di folder `data/` untuk mempercepat eksekusi ulang. Hapus jika data input berubah (akan di-regenerate otomatis saat file input lebih baru dari cache).

---

## Setup

Pastikan Python ≥ 3.9 dan dependensi berikut terinstal:

```bash
pip install pandas openpyxl numpy
```

---

## Cara Menjalankan

### Menjalankan Seluruh Pipeline (Orchestrator)

Cukup jalankan master orchestrator:

```bash
python main.py
```

Script ini akan men-trigger 6 step berurutan secara otomatis:
1. `cleaning_facul.py` -> `data/facul_clean_aca.xlsx`
2. `cleaning_osbal.py` -> `data/osbal_clean_aca.xlsx`
3. `cleaning_suspend.py` -> `data/suspend_clean_aca.xlsx`
4. `prod_sus_1.py` -> `data/final_output_v1.xlsx`
5. `prod_sus_2.py` -> `data/final_output_v2.xlsx`
6. `alter.py` -> `data/final_output.xlsx`

### Menjalankan Script Secara Individual (Opsional)

Jika hanya ingin menjalankan tahap tertentu:

```bash
# Cleaning
python cleaning_suspend.py
python cleaning_osbal.py
python cleaning_facul.py

# Matching & Analisis
python prod_sus_1.py
python prod_sus_2.py
python alter.py
```

---

## Alur Pipeline (`main.py`)

Setiap baris SUSPEND dicocokkan secara sekuensial melewati pass berikut:

| Pass | Sumber | Field | Keterangan |
|------|--------|-------|------------|
| 1 | OSBAL | Polis / Slip | Direct match (Exact & LIKE substring) |
| 3 | SLIPDB → OSBAL | Polis / Slip | Indirect: match SLIPDB, resolve FAC_CODE ke OSBAL |
| 5 | FACUL → OSBAL | Polis / Slip | Indirect: match FACUL, resolve FAC_CODE ke OSBAL |
| 7a | OSBAL | Insured | Fallback insured (jika Pass 1-5 gagal semua) |
| 7b | SLIPDB → OSBAL | Insured | Fallback insured indirect via SLIPDB |
| 7c | FACUL → OSBAL | Insured | Fallback insured indirect via FACUL |
| 8 | — | — | Unmatching (semua pass gagal) |

**Narrowing**: Setelah match ditemukan via satu field (Slip atau Polis), field pelengkap dipakai sebagai konfirmasi untuk menyempitkan kandidat. Ini menghasilkan skenario: `Slip only`, `Slip + Polis`, `Polis only`, `Polis + Slip`, `Insured only`.

**Akumulasi Fac Code**: Jika beberapa baris Suspend terpisah match ke fac code tunggal yang sama, baris-baris tersebut digabung menjadi 1 baris output dengan `AMOUNT ORI` diakumulasikan.

---

## Kolom Output Penting

| Kolom | Keterangan |
|-------|------------|
| `CCOS_DOC_NO` | Nomor dokumen dari OSBAL |
| `CCOS_REF_CODE` | FAC Code dari OSBAL |
| `CCOS_BAL_DUE` | Balance due dari OSBAL (akumulasi jika multi-row) |
| `AMOUNT ORI` | Amount dari Suspend (diakumulasi jika multi-row) |
| `AMOUNT_ORI_MIN1` | `AMOUNT ORI × -1` |
| `DIFERENCE` | `AMOUNT_ORI_MIN1 - CCOS_BAL_DUE` |
| `FLAG_PROD` | Kategori hasil matching (lihat tabel di bawah) |
| `Mark Admin Fac Code` | FAC Code tunggal, atau `facode lebih dari 1` jika ambiguous |
| `SKENARIO` | Label skenario matching final |

### Nilai `FLAG_PROD`

| Nilai | Kondisi |
|-------|---------|
| `Adjustment total tanpa akumulasi` | `AMOUNT_ORI_MIN1 == CCOS_BAL_DUE`, tanpa akumulasi |
| `Adjustment total dengan akumulasi` | `AMOUNT_ORI_MIN1 == CCOS_BAL_DUE`, dengan akumulasi |
| `Adjustment sebagian tanpa akumulasi` | `AMOUNT_ORI_MIN1 < CCOS_BAL_DUE`, tanpa akumulasi |
| `Adjustment sebagian dengan akumulasi` | `AMOUNT_ORI_MIN1 < CCOS_BAL_DUE`, dengan akumulasi |
| `New Entry` | `AMOUNT_ORI_MIN1 > CCOS_BAL_DUE` |
| `Matching >1 fac code` | Match ambiguous ke lebih dari 1 FAC Code |
| `Unmatching` | Tidak ditemukan match apapun |

---

## Diagnostik & Unit Testing

Semua file diagnostik, investigasi, dan unit test tersimpan di folder `debugging/`:

### Trace Detail Baris Suspend
Untuk men-trace hasil matching satu baris Suspend secara detail:

```python
# Edit TARGET_SLIP di debugging/debug_trace_suspend.py
TARGET_SLIP = "70009031911000039"
```

```bash
python debugging/debug_trace_suspend.py
```

Output menampilkan hasil setiap pass secara lengkap: raw candidates, narrowed candidates, FAC code sebelum/sesudah narrowing, dan resolusi ke OSBAL.

### Unit Testing
Untuk menjalankan unit test logika cleaning & plus suffix expansion:

```bash
python debugging/test_clean_polis_plus.py
```
