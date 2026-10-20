# Panduan Penggunaan Script SAS Viya — Cedant MITRA UTAMA (PT. MITRA UTAMA REASURANSI)

Folder ini berisi implementasi SAS Viya tingkat produksi (*production-grade*) untuk pipeline rekonsiliasi data Suspend, OSBAL, dan FACUL cedant PT. MITRA UTAMA REASURANSI, yang telah diselaraskan **100% identik dengan alur dan output Python**.

---

## 🗄️ Status Verifikasi Database PostgreSQL (`fsi_db`)
* **Host & Port:** `10.10.125.91:5432`
* **User & Database:** `viy4user` @ `fsi_db`
* **Sumber Data Raw:** `SUSPESE_FACUL_ORI_2026`, `SUSPESE_OSBAL_ORI_2026`, `SUSPESE_SUSPENSE_ORI_2026`.

---

## 🚀 Arsitektur Modular
```text
sas_ready/
├── mitrautama_cleaning.sas         ← [ETL Pipeline] Clean OSBAL, Facul, Suspend -> DB & CAS
├── mitrautama_matching.sas         ← [Matching Engine] Read Clean Tables -> Output V1 & V2
├── cleaning_facul.sas       ← [Standalone Facul Cleaner]
├── cleaning_osbal.sas       ← [Standalone OSBAL Cleaner]
├── cleaning_suspend.sas     ← [Standalone Suspend Cleaner]
├── macros_cleaning.sas      ← [Shared Macro Library]
├── prod_matching.sas        ← [Standalone Engine]
└── main.py                  ← [Runner Orchestrator]
```

### Catatan Khusus Cedant:
Urutan Polis -> Slip -> Insured. Narrowing Group DESC 2.
