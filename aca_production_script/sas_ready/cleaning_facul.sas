/* =============================================================================
   CLEANING_FACUL.SAS
   Konversi dari: cleaning_facul.py

   Fungsi:
   - Load data/facul.xlsx
   - Filter cedant = "PT. ASURANSI CENTRAL ASIA"
   - Clean polis, slip, insured (maks 5 kolom masing-masing)
   - Ekstrak sertifikat 6-digit dari polis_ori
   - Output Excel: data/facul_clean_aca.xlsx
   - Export ke PostgreSQL via ODBC (tabel SUSPENSE_DATA_FACUL_CLEAN)

   Optimasi vs Python:
   - Menggunakan SAS DATA step array (lebih efisien dari pandas iterrows)
   - PRXCHANGE/PRXMATCH menggantikan Python re module
   - Permanent SAS dataset (.sas7bdat) sebagai cache ganti .cache.pkl
   ============================================================================= */

options mprint mlogic symbolgen;

/* ── 0. Path setup ─────────────────────────────────────────────────────────── */
%let BASE_DIR   = %str(C:\Users\berryl\Desktop\starcore\indore\production_script);
%let DATA_DIR   = %str(&BASE_DIR.\data);
%let INPUT_FILE = %str(&DATA_DIR.\facul.xlsx);
%let OUT_FILE   = %str(&DATA_DIR.\facul_clean_aca.xlsx);
%let MACRO_INC  = %str(&BASE_DIR.\sas_ready\macros_cleaning.sas);

/* ── 1. Include shared macros ──────────────────────────────────────────────── */
%include "&MACRO_INC.";

/* ── 2. Definisi konstanta bisnis ─────────────────────────────────────────── */
%let CEDANT_COL   = COMP_NAME;
%let CEDANT_VALUE = %str(PT. ASURANSI CENTRAL ASIA);
%let POLIS_COL    = FAC_POLICY_NO;
%let SLIP_COL     = FAC_SLIP;
%let INSURED_COL  = FAC_INSURED;
%let BROKER_COL   = COMP_NAME2;
%let MAX_COLS     = 5;
%let TABLE_NAME   = SUSPENSE_DATA_FACUL_CLEAN;

/* Compile regex formats */
%compile_cleaning_regex();
%define_bulan_format();


/* =============================================================================
   STEP 1: Baca file Excel
   Setara dengan: df = pd.read_excel(INPUT_FILE, header=0)
   ============================================================================= */
%put [1/5] Reading: &INPUT_FILE.;

proc import
    datafile="&INPUT_FILE."
    out=work.facul_raw
    dbms=xlsx
    replace;
    getnames=yes;   /* header=0 → baris pertama sebagai nama kolom */
    datarow=2;
run;

%put       Total rows: %sysfunc(attrn(%sysfunc(open(work.facul_raw)), nobs));


/* =============================================================================
   STEP 2: Filter cedant = "PT. ASURANSI CENTRAL ASIA"
   Setara dengan: df = df[df[CEDANT_COL] == CEDANT_VALUE].copy()
   ============================================================================= */
%put [2/5] Filtering cedant = &CEDANT_VALUE.;

data work.facul_filtered;
    set work.facul_raw;
    /* Rename kolom jika perlu (COMP_NAME.1 → COMP_NAME2) */
    %if %varexist(work.facul_raw, COMP_NAME_1) %then %do;
        if missing(COMP_NAME2) then COMP_NAME2 = COMP_NAME_1;
    %end;

    /* Strip dan filter */
    _cedant_chk_ = strip(&CEDANT_COL.);
    if _cedant_chk_ = "&CEDANT_VALUE." then output;
    drop _cedant_chk_;
run;

%put       Rows setelah filter: %sysfunc(attrn(%sysfunc(open(work.facul_filtered)), nobs));


/* =============================================================================
   STEP 3: Rename kolom utama dan buat business_partner
   Setara dengan: df.rename(...) dan get_business_partner(...)
   ============================================================================= */
data work.facul_renamed;
    set work.facul_filtered;
    rename &POLIS_COL.   = polis_ori
           &SLIP_COL.    = slip_ori
           &INSURED_COL. = insured_ori;
run;

data work.facul_bp;
    set work.facul_renamed;
    length business_partner $500;

    /* get_business_partner: broker jika ada & bukan DIRECT; else cedant */
    _broker_ = strip(&BROKER_COL.);
    if missing(_broker_) or upcase(_broker_) = 'DIRECT' then
        business_partner = strip(&CEDANT_COL.);
    else
        business_partner = _broker_;
    drop _broker_;
run;


/* =============================================================================
   STEP 4: Cleaning — polis, slip, insured, sertif
   Setara dengan: map(clean_polis), map(clean_slip), map(clean_insured)

   Optimasi: DATA step single-pass dengan array, jauh lebih efisien
   dari iterrows(). Setara dengan pandas map() vectorized.
   ============================================================================= */
%put [3/5] Cleaning ...;

data work.facul_cleaned;
    set work.facul_bp;

    /* Deklarasi variabel output */
    length clean_polis_1   - clean_polis_5   $500
           clean_slip_1    - clean_slip_5    $500
           clean_insured_1 - clean_insured_5 $500
           clean_sertif_1  - clean_sertif_5  $10;

    /* -- 4a. Clean polis ---------------------------------------------------- */
    %apply_clean_polis(polis_ori);

    /* -- 4b. Ekstrak sertifikat dari polis_ori ------------------------------ */
    %apply_extract_sertif(polis_ori);

    /* -- 4c. Clean slip ----------------------------------------------------- */
    %apply_clean_slip(slip_ori);

    /* -- 4d. Clean insured -------------------------------------------------- */
    %apply_clean_insured(insured_ori);

    /* -- 4e. Hapus karakter kontrol XML ilegal dari semua kolom string ------- */
    array _str_cols{*} _character_;
    do _sci_ = 1 to dim(_str_cols);
        if not missing(_str_cols{_sci_}) then
            %clean_control_chars(_str_cols{_sci_});
    end;
    drop _sci_;

run;

%put       Cleaning selesai.;


/* =============================================================================
   STEP 5: Susun urutan kolom output
   Setara dengan: df = df[new_columns]

   Kolom urutan:
   [kolom asli] → polis_ori, clean_polis_1..5, clean_sertif_1..5,
                  slip_ori,  clean_slip_1..5,
                  insured_ori, clean_insured_1..5
   ============================================================================= */
%put [4/5] Building output columns ...;

/* Buat dataset dengan urutan kolom yang benar */
data work.facul_output;
    retain &CEDANT_COL. business_partner
           polis_ori clean_polis_1-clean_polis_5 clean_sertif_1-clean_sertif_5
           slip_ori  clean_slip_1-clean_slip_5
           insured_ori clean_insured_1-clean_insured_5;
    set work.facul_cleaned;
run;


/* =============================================================================
   STEP 6: Simpan ke Excel
   Setara dengan: df.to_excel(OUTPUT_FILE, index=False)
   ============================================================================= */
%put [5/5] Saving to: &OUT_FILE.;

proc export
    data=work.facul_output
    outfile="&OUT_FILE."
    dbms=xlsx
    replace;
    sheet="facul_clean";
run;

%put Done. Output: &OUT_FILE.;


/* =============================================================================
   STEP 7: Export ke PostgreSQL via ODBC / CAS (opsional)
   Setara dengan: df_export.to_sql(table_name, ...)

   Ganti &ODBC_DSN. dengan nama DSN PostgreSQL Anda.
   Jika menggunakan SAS Viya CAS, gunakan proc casutil.
   ============================================================================= */

/* --- Opsi A: Via ODBC (SAS 9.4 / Viya on-premise) --- */
/*
libname pg_lib odbc
    datasrc="&ODBC_DSN."
    user="&DB_USER."
    password="&DB_PASS."
    schema="public";

proc append
    base=pg_lib.&TABLE_NAME.
    data=work.facul_output
    force;
run;

libname pg_lib clear;
*/

/* --- Opsi B: Via SAS Viya CAS session --- */
/*
cas mySession;
libname caslib cas caslib="Public";
proc casutil;
    load data=work.facul_output
         outcaslib="Public"
         casout="&TABLE_NAME."
         replace;
run;
quit;
*/

%put NOTE: Export PostgreSQL/CAS - uncomment sesuai environment Anda.;
%put [FACUL CLEANING] Selesai.;

/* Cleanup */
proc datasets library=work nolist;
    delete facul_raw facul_filtered facul_renamed facul_bp facul_cleaned;
run;
quit;
