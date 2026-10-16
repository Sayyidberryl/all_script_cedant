# Panduan Penggunaan Script SAS Viya — Cedant PT Asuransi Astra Buana

Folder ini berisi implementasi SAS Viya / SAS 9.4 Studio untuk seluruh pipeline rekonsiliasi data Suspend, OSBAL, FACUL, dan Bordero Open Cover cedant PT Asuransi Astra Buana.

---

## 🚀 Cara Cepat (Langsung Copy-Paste)

Gunakan file **`astra_pipeline_viya.sas`**:
1. Buka SAS Studio di web browser SAS Viya Anda.
2. Buat program baru (**New SAS Program**).
3. Copy seluruh isi file [`astra_pipeline_viya.sas`](file:///c:/Users/berryl/Desktop/starcore/indore/all_script_cedant/astra_production_script/sas_ready/astra_pipeline_viya.sas).
4. Paste ke dalam editor program SAS Studio.
5. Cek baris 21 untuk macro variable `%let BASE_DIR = ...;`:
   - Ubah sesuai direktori folder kerja SAS server Anda (atau biarkan default jika dijalankan di PC lokal).
6. Klik tombol **Run (F3)**.

Seluruh tahapan pembersihan data (Cleaning Facul, OSBAL, Suspend) dan algoritma matching multi-pass, narrowing mata uang & periode, pengecekan 3 file Bordero Open Cover (FIRE, MC, MH), kalkulasi `FLAG_PROD`, serta export file Excel `final_output_astrabuana_v1.xlsx` akan dieksekusi secara otomatis dan berurutan!

---

## 📁 Daftar File & Struktur Modular

Jika ingin menjalankan tahapan secara terpisah atau per modul:

| File | Fungsi | Input | Output |
| :--- | :--- | :--- | :--- |
| **`macros_cleaning_astra.sas`** | Library macro bersama (Regex format polis Astra minimal 11 digit, parser range sertifikat, pembersih slip & insured Astra, standarisasi S/D -> SD) | - | Macro functions & formats |
| **`cleaning_facul.sas`** | Cleaning data transaksi Facul Astra Buana | `data/facul.xlsx` (atau `1b. Transaksi Facul...xlsx`) | `data/astrabuana_output_facul.xlsx` |
| **`cleaning_osbal.sas`** | Cleaning data tagihan OSBAL Astra Buana (dengan prioritas CLSDT -> FAC) | `data/osbal.xlsx` (atau `2b. transaksi Osbal...xlsx`) | `data/astrabuana_output_osbal.xlsx` |
| **`cleaning_suspend.sas`** | Cleaning data Suspend Astra Buana (header baris ke-3, filter status SUSPENSE) | `data/suspend.xlsx` (atau `3b. Database Suspense...xlsx`) | `data/astrabuana_output_suspend.xlsx` |
| **`prod_matching.sas`** | Mesin matching Suspend vs OSBAL / FACUL / Bordero (FIRE, MC, MH) | Output cleaning + Folder `data/BORDERO` | `data/final_output_astrabuana_v1.xlsx` (atau `v2.xlsx`) |
| **`astra_pipeline_viya.sas`** | **Master script All-in-One** (berisi seluruh logika tanpa dependensi file eksternal) | Excel mentah di `data/` | Semua file Excel clean & matching output 38 kolom |

---

## ⚙️ Parameter Penting di Script Matching

Pada script `prod_matching.sas` atau `astra_pipeline_viya.sas`, Anda dapat mengatur versi matching:
- **Versi 1 (`%let VERSION = 1; %let ACCUMULATE_MODE = KEEP_ROWS;`)**:
  Setiap baris transaksi suspend asli tetap dipertahankan utuh pada file output (`final_output_astrabuana_v1.xlsx`).
- **Versi 2 (`%let VERSION = 2; %let ACCUMULATE_MODE = MERGE_ROWS;`)**:
  Baris suspend yang mengarah pada fac code tunggal yang sama diakumulasikan menjadi 1 baris per grup kode (`final_output_astrabuana_v2.xlsx`).

---

## 📋 Struktur 38 Kolom Output Final Astra Buana

Output akhir disusun tepat menjadi 38 kolom alternating Polis & Sertifikat sesuai spesifikasi produksi:
1. `CCOS_DOC_NO`
2. `CCOS_REF_CODE`
3. `RECEIPT NO`
4. `CREDIT NOTES`
5. `DETAIL RINCIAN NO`
6. `RECEIPT DATE`
7. `CEDANT NAME`
8. `CEDANT SHRT NAME`
9. `INSURED_ORI`
10. `INSURED_1`
11. `INSURED_2`
12. `CURR ORI`
13. `AMOUNT ORI`
14. `AMOUNT_ORI_MIN1`
15. `CURR PAY`
16. `AMOUNT PAY`
17. `CCOS_OR_BAL`
18. `CCOS_BAL_DUE`
19. `DIFERENCE`
20. `FLAG_PROD`
21. `POLIS`
22. `POLICY_CLEAN_1`
23. `SERTIF_CLEAN_1`
24. `POLICY_CLEAN_2`
25. `SERTIF_CLEAN_2`
26. `POLICY_CLEAN_3`
27. `SERTIF_CLEAN_3`
28. `POLICY_CLEAN_4`
29. `POLICY_CLEAN_5`
30. `SLIP_NO`
31. `SLIP_NO_CLN`
32. `DESC 1`
33. `DESC 2`
34. `DESC 3`
35. `DESC 4`
36. `STATUS`
37. `REC_TYPE`
38. `SKENARIO`

---

## 📊 Kategori Status Produksi (`FLAG_PROD`)

Hasil matching diklasifikasikan ke dalam kategori resmi:
1. `Beda Currency`
2. `Beda Periode`
3. `Matching >1 Fac code OC MC`
4. `Matching >1 fac code`
5. `New Entry total`
6. `Adjustment total tanpa akumulasi`
7. `Adjustment total dengan akumulasi`
8. `Adjustment sebagian tanpa akumulasi`
9. `Adjustment sebagian dengan akumulasi`
10. `New Entry sebagian`
11. `Unmatching` (default fallback jika tidak ditemukan kecocokan)
