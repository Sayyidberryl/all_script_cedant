# Panduan Penggunaan Script SAS Viya — Cedant ACA (PT. Asuransi Central Asia)

Folder ini berisi implementasi SAS Viya tingkat produksi (*production-grade*) untuk seluruh pipeline rekonsiliasi data Suspend, OSBAL, FACUL, dan Bordero Open Cover cedant PT. Asuransi Central Asia (ACA), yang telah diselaraskan **100% identik dengan alur dan output Python**.

---

## 🗄️ Status Verifikasi Database PostgreSQL (`fsi_db`)

Koneksi langsung ke server database telah diverifikasi aktif dan berhasil terhubung:
* **Host & Port:** `10.10.125.91:5432`
* **User & Database:** `viy4user` @ `fsi_db`
* **Sumber Data Raw:** `SUSPENSE_FACUL_ORI_2026`, `SUSPENSE_OSBAL_ORI_2026`, `SUSPENSE_ORI_2026`, `ACA_OPEN_COVER_MC`.

---

## 🚀 Arsitektur Modular (Clean Code & Production-Grade)

Sesuai best practice enterprise, pipeline dipisahkan menjadi **2 script mandiri** agar clean, modular, mudah di-debug, dan hemat komputasi:

```text
sas_ready/
├── aca_cleaning.sas         ← [ETL Pipeline] Clean OSBAL, Facul, Suspend -> DB & CAS (~590 baris)
├── aca_matching.sas         ← [Matching Engine] Read Clean Tables -> Output V1 & V2 -> DB & CAS (~980 baris)
└── aca_pipeline_viya.sas    ← [All-in-One Fallback] Seluruh proses dalam 1 file (opsional)
```

### 1. File 1: [**`aca_cleaning.sas`**](file:///c:/Users/berryl/Desktop/starcore/indore/all_script_cedant/aca_production_script/sas_ready/aca_cleaning.sas)
* **Tanggung Jawab**:
  - Ambil data raw dari PostgreSQL `FSI_DB` via `PROC FEDSQL sessref=mycas;`.
  - Normalisasi string, regex, validasi token (polis, slip, sertifikat, insured).
  - Hasilkan dataset bersih:
    - **Clean OSBAL** (42 kolom standar)
    - **Clean Facul** (36 kolom standar)
    - **Clean Suspend** (in-memory CAS)
  - Simpan ke in-memory CAS `casuser` & `Public` (`.sashdat`) dan sinkronisasi ke PostgreSQL:
    - `SUSPENSE_DATA_OSBAL_CLEAN` & `SUSPENSE_OSBAL_CLEAN_2026`
    - `SUSPENSE_DATA_FACUL_CLEAN` & `SUSPENSE_FACUL_CLEAN_2026`
* **Kapan Dijalankan**: Cukup dijalankan ketika ada data mentah baru dari cedant.

### 2. File 2: [**`aca_matching.sas`**](file:///c:/Users/berryl/Desktop/starcore/indore/all_script_cedant/aca_production_script/sas_ready/aca_matching.sas)
* **Tanggung Jawab**:
  - Membaca tabel yang sudah bersih dari CAS staging (atau fallback dari PostgreSQL jika sesi CAS baru).
  - Menjalankan mesin pencocokan multi-pass (Pass 1 Polis, Pass 2 Slip, Pass 3 Insured, Pass 4 Bordero Open Cover).
  - Menghasilkan dua output standar secara simultan:
    - **Suspense V1** (38 kolom, row-level / `KEEP_ROWS`)
    - **Suspense V2** (38 kolom, agregasi per FAC CODE / `MERGE_ROWS`)
  - Sinkronisasi idempoten ke PostgreSQL:
    - `SUSPENSE_DATA_SUSPENSE_V1` & `SUSPENSE_CLEAN_2026`
    - `SUSPENSE_DATA_SUSPENSE_V2`
  - Cetak ringkasan eksekutif (identik format Python) ke SAS Log dan tabel frekuensi (`FLAG_PROD * SKENARIO`) di tab Results SAS Studio.
* **Kapan Dijalankan**: Dapat dijalankan berkali-kali secara instan dalam hitungan detik tanpa perlu mengulang pembersihan 500k data OSBAL.

---

## 📋 Ringkasan Pemetaan Destinasi Tabel

| Entitas Data | Kolom | CASLIB In-Memory (`casuser` & `Public`) | Tabel PostgreSQL (`FSI_DB`) | Script Python Padanan |
| :--- | :---: | :--- | :--- | :--- |
| **Clean OSBAL** | 42 | `casuser.ACA_OSBAL_CLEAN` & `Public.ACA_OSBAL_CLEAN` | `SUSPENSE_DATA_OSBAL_CLEAN` | `cleaning_osbal.py` |
| **Clean Facul** | 36 | `casuser.ACA_FACUL_CLEAN` & `Public.ACA_FACUL_CLEAN` | `SUSPENSE_DATA_FACUL_CLEAN` | `cleaning_facul.py` |
| **Clean Suspend** | — | `casuser.ACA_SUSPEND_CLEAN` & `Public.ACA_SUSPEND_CLEAN` | *(Staging In-Memory CAS)* | `cleaning_suspend.py` |
| **Suspense V1** | 38 | `casuser.ACA_MATCHING_V1` & `Public.ACA_MATCHING_V1` | `SUSPENSE_DATA_SUSPENSE_V1` & `SUSPENSE_CLEAN_2026` | `prod_sus_1.py` |
| **Suspense V2** | 38 | `casuser.ACA_MATCHING_V2` & `Public.ACA_MATCHING_V2` | `SUSPENSE_DATA_SUSPENSE_V2` | `prod_sus_2.py` |

---

## 🌐 Arsitektur Multi-Cedant (Tabel Global di Database)

Tabel-tabel di database PostgreSQL `FSI_DB` (`SUSPENSE_DATA_OSBAL_CLEAN`, `SUSPENSE_DATA_FACUL_CLEAN`, `SUSPENSE_DATA_SUSPENSE_V1`, `SUSPENSE_DATA_SUSPENSE_V2`, dan `SUSPENSE_CLEAN_2026`) bersifat **tabel global terkonsolidasi** yang menampung seluruh cedant (Astra, Lippo, ACA, hingga puluhan cedant lainnya).

Mekanisme sinkronisasi database di script SAS ini bersifat **idempoten dan terisolasi**:
1. **Targeted Deletion**: Sebelum menyisipkan data baru, script **hanya menghapus baris milik ACA** (`WHERE "CEDANT NAME" = 'PT. ASURANSI CENTRAL ASIA'`).
2. **Preservation**: Data cedant lain yang ada di tabel yang sama (misal Astra, Lippo) **tidak tersentuh dan tetap utuh**.
3. **Safe Append**: Data bersih & matching terbaru ACA di-*append* ke tabel global tersebut.

---

## ⚙️ Cara Menjalankan di SAS Viya

1. **Jalankan Pembersihan**: Buka dan jalankan [**`aca_cleaning.sas`**](file:///c:/Users/berryl/Desktop/starcore/indore/all_script_cedant/aca_production_script/sas_ready/aca_cleaning.sas).
2. **Jalankan Rekonsiliasi**: Buka dan jalankan [**`aca_matching.sas`**](file:///c:/Users/berryl/Desktop/starcore/indore/all_script_cedant/aca_production_script/sas_ready/aca_matching.sas).
3. *(Opsional)* Jika menginginkan eksekusi tunggal sekali jalan, file [**`aca_pipeline_viya.sas`**](file:///c:/Users/berryl/Desktop/starcore/indore/all_script_cedant/aca_production_script/sas_ready/aca_pipeline_viya.sas) tetap tersedia.
