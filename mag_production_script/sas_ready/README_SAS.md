# Panduan Penggunaan Script SAS Viya — Cedant MAG (PT Asuransi Multi Artha Guna)

Folder ini berisi implementasi SAS Viya tingkat produksi (*production-grade*) untuk seluruh pipeline pembersihan dan rekonsiliasi data Suspend, OSBAL, dan FACUL cedant PT Asuransi Multi Artha Guna (MAG), yang diselaraskan **100% identik dengan alur dan output Python (matching_mag.py)**.

---

## 🗄️ Status Verifikasi Database PostgreSQL (`fsi_db`)

Koneksi langsung ke server database PostgreSQL telah diverifikasi aktif dan terintegrasi:
* **Host & Port:** `10.10.125.91:5432`
* **User & Database:** `viy4user` @ `fsi_db`
* **Sumber Data Raw:** `SUSPESE_FACUL_ORI_2026`, `SUSPESE_OSBAL_ORI_2026`, `SUSPESE_SUSPENSE_ORI_2026`.

---

## 🚀 Arsitektur Modular (Clean Code & Production-Grade)

Sesuai best practice enterprise aca_prod, pipeline dipisahkan menjadi **2 script mandiri** agar clean, modular, mudah di-debug, dan hemat komputasi:

```text
sas_ready/
├── mag_cleaning.sas         ← [ETL Pipeline] Clean OSBAL, Facul, Suspend -> DB & CAS (~600 baris)
├── mag_matching.sas         ← [Matching Engine] Read Clean Tables -> Output V1 & V2 -> DB & CAS (~950 baris)
├── cleaning_facul.sas       ← [Standalone Facul] Pembersihan data fakultatif lokal / batch
├── cleaning_osbal.sas       ← [Standalone OSBAL] Pembersihan data osbal lokal / batch
├── cleaning_suspend.sas     ← [Standalone Suspend] Pembersihan data suspend lokal / batch
├── macros_cleaning.sas      ← [Shared Macro Library] Tokenizing, PRX regex patterns, normalisasi
├── prod_matching.sas        ← [Standalone Engine] Multi-pass matching engine (V1/V2)
└── main.py                  ← [Runner Orchestrator] Python wrapper eksekusi lokal
```

### 1. File 1: `mag_cleaning.sas`
* **Tanggung Jawab**:
  - Ambil data raw dari PostgreSQL `FSI_DB` via `PROC FEDSQL sessref=mycas;` dengan filter `COMP_NAME like 'PT ASURANSI MULTI ARTHA GUNA%'`.
  - Normalisasi string, regex, validasi token (polis, slip, sertifikat, insured).
  - Khusus MAG: penanganan pola `<POLIS AKTIF> EX[.] [POLICY NO] [:] <POLIS LAMA>` dan sertifikat format dash/range.
  - Hasilkan dataset bersih:
    - **Clean OSBAL** (42 kolom standar)
    - **Clean Facul** (36 kolom standar)
    - **Clean Suspend** (in-memory CAS)
  - Simpan ke in-memory CAS `casuser` & `Public` dan sinkronisasi ke PostgreSQL:
    - `SUSPENSE_DATA_OSBAL_CLEAN`
    - `SUSPENSE_DATA_FACUL_CLEAN`

### 2. File 2: `mag_matching.sas`
* **Tanggung Jawab**:
  - Membaca tabel yang sudah bersih dari CAS staging / PostgreSQL.
  - Menjalankan mesin pencocokan multi-pass (Pass Polis -> Pass Slip -> Pass Insured -> Currency & Periode -> Narrowing).
  - Menghasilkan dua output standar secara simultan:
    - **Suspense V1** (32 kolom standar MAG / `KEEP_ROWS`)
    - **Suspense V2** (32 kolom standar MAG / `MERGE_ROWS`)
  - Sinkronisasi idempoten ke PostgreSQL:
    - `SUSPENSE_DATA_SUSPENSE_V1` & `SUSPENSE_CLEAN_2026`
    - `SUSPENSE_DATA_SUSPENSE_V2`
  - Cetak ringkasan eksekutif dan tabel frekuensi (`FLAG_PROD * SKENARIO`) ke SAS Log.

---

## 📋 Ringkasan Pemetaan Destinasi Tabel

| Entitas Data | Kolom | CASLIB In-Memory (`casuser` & `Public`) | Tabel PostgreSQL (`FSI_DB`) | Script Python Padanan |
| :--- | :---: | :--- | :--- | :--- |
| **Clean OSBAL** | 42 | `casuser.MAG_OSBAL_CLEAN` & `Public.MAG_OSBAL_CLEAN` | `SUSPENSE_DATA_OSBAL_CLEAN` | `cleansing_osbal.py` |
| **Clean Facul** | 36 | `casuser.MAG_FACUL_CLEAN` & `Public.MAG_FACUL_CLEAN` | `SUSPENSE_DATA_FACUL_CLEAN` | `cleansing_facul.py` |
| **Clean Suspend** | — | `casuser.MAG_SUSPEND_CLEAN` & `Public.MAG_SUSPEND_CLEAN` | *(Staging In-Memory CAS)* | `cleansing_suspend.py` |
| **Suspense V1** | 32 | `casuser.MAG_MATCHING_V1` & `Public.MAG_MATCHING_V1` | `SUSPENSE_DATA_SUSPENSE_V1` & `SUSPENSE_CLEAN_2026` | `matching_mag.py` (V1) |
| **Suspense V2** | 32 | `casuser.MAG_MATCHING_V2` & `Public.MAG_MATCHING_V2` | `SUSPENSE_DATA_SUSPENSE_V2` | `matching_mag.py` (V2) |

---

## ⚙️ Cara Menjalankan di SAS Viya

1. **Jalankan Pembersihan**: Buka dan jalankan `mag_cleaning.sas`.
2. **Jalankan Rekonsiliasi**: Buka dan jalankan `mag_matching.sas`.
