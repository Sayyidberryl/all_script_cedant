/* =============================================================================
   PROGRAM   : CLEANING_FACUL.SAS
   CEDANT    : PT ASURANSI ASTRA BUANA
   FUNGSI    : Pembersihan data transaksi Facul Astra Buana
   INPUT     : data/facul.xlsx (atau 1b. Transaksi Facul 01.01.23 - 17.08.2026.xlsx)
   OUTPUT    : data/astrabuana_output_facul.xlsx
   KOMPATIBEL: SAS Viya / SAS 9.4 Studio
   ============================================================================= */

options validvarname=any validmemname=extend mprint symbolgen compress=yes;

/* ── 0. Path setup ─────────────────────────────────────────────────────────── */
%if not %symexist(BASE_DIR) %then %do;
    %let BASE_DIR = %str(c:\Users\berryl\Desktop\starcore\indore\all_script_cedant\astra_production_script);
%end;
%let DATA_DIR   = %str(&BASE_DIR.\data);
%let OUT_FILE   = %str(&DATA_DIR.\astrabuana_output_facul.xlsx);
%let MACRO_INC  = %str(&BASE_DIR.\sas_ready\macros_cleaning_astra.sas);

/* Cari file input yang tersedia */
%let INPUT_FILE = %str(&DATA_DIR.\facul.xlsx);
%macro detect_input();
    %if not %sysfunc(fileexist("&INPUT_FILE.")) %then %do;
        %let ALT_FILE = %str(&BASE_DIR.\dataExcel\raw\1b. Transaksi Facul 01.01.23 - 17.08.2026.xlsx);
        %if %sysfunc(fileexist("&ALT_FILE.")) %then %let INPUT_FILE = &ALT_FILE.;
    %end;
%mend detect_input;
%detect_input();

/* ── 1. Include shared macros ──────────────────────────────────────────────── */
%if %sysfunc(fileexist("&MACRO_INC.")) %then %do;
    %include "&MACRO_INC.";
%end;

/* ── 2. Definisi konstanta bisnis ─────────────────────────────────────────── */
%let CEDANT_COL   = COMP_NAME;
%let CEDANT_VALUE = %str(PT ASURANSI ASTRA BUANA);
%let TABLE_NAME   = SUSPENSE_DATA_FACUL_CLEAN;

%define_bulan_format();


/* =============================================================================
   STEP 1: Baca file Excel Facul
   ============================================================================= */
%put [1/4] Reading: &INPUT_FILE. ...;

proc import
    datafile="&INPUT_FILE."
    out=work.facul_raw
    dbms=xlsx
    replace;
    getnames=yes;
    datarow=2;
run;

%put       Total rows dibaca: %sysfunc(attrn(%sysfunc(open(work.facul_raw)), nobs));


/* =============================================================================
   STEP 2: Filter Cedant PT ASURANSI ASTRA BUANA
   ============================================================================= */
%put [2/4] Filtering cedant: &CEDANT_VALUE. ...;

data work.facul_filtered;
    set work.facul_raw;
    _cedant_chk_ = strip(&CEDANT_COL.);
    if _cedant_chk_ = "&CEDANT_VALUE." then output;
    drop _cedant_chk_;
run;

%put       Rows setelah filter: %sysfunc(attrn(%sysfunc(open(work.facul_filtered)), nobs));


/* =============================================================================
   STEP 3: Cleaning & Breakdown (Insured, Polis, Sertif, Slip)
   ============================================================================= */
%put [3/4] Cleaning & breakdown columns ...;

data work.facul_cleaned;
    set work.facul_filtered;

    length INSURED_CLEAN_1 - INSURED_CLEAN_5 $500
           POLICY_CLEAN_1  - POLICY_CLEAN_5  $200
           SERTIF_CLEAN_1  - SERTIF_CLEAN_3  $50
           SLIP_CLEAN_1    - SLIP_CLEAN_5    $200;

    /* 1. Clean Insured */
    %apply_clean_insured_astra(FAC_INSURED);

    /* 2. Clean Polis */
    %apply_clean_polis_astra(FAC_POLICY_NO);

    /* 3. Sertif Clean (Rule 1 Astra: sertif berurutan sehabis polis) */
    SERTIF_CLEAN_1 = ''; SERTIF_CLEAN_2 = ''; SERTIF_CLEAN_3 = '';

    /* 4. Clean Slip */
    %apply_clean_slip_astra(FAC_SLIP);

    /* Clean control chars */
    array _str_cols{*} _character_;
    do _sci_ = 1 to dim(_str_cols);
        if not missing(_str_cols{_sci_}) then
            %clean_control_chars(_str_cols{_sci_});
    end;
    drop _sci_;
run;


/* =============================================================================
   STEP 4: Standarisasi Urutan Kolom & Export ke Excel
   ============================================================================= */
%put [4/4] Building output columns & saving to: &OUT_FILE. ...;

data work.facul_output;
    retain FAC_CODE FAC_CEDANT COMP_NAME FAC_BROKER COMP_NAME2 BUSSINESS_PATNER
           FAC_INSURED INSURED_CLEAN_1 INSURED_CLEAN_2 INSURED_CLEAN_3 INSURED_CLEAN_4 INSURED_CLEAN_5
           FAC_POLICY_NO POLICY_CLEAN_1 SERTIF_CLEAN_1 POLICY_CLEAN_2 SERTIF_CLEAN_2 POLICY_CLEAN_3 SERTIF_CLEAN_3
           POLICY_CLEAN_4 POLICY_CLEAN_5
           FAC_SLIP SLIP_CLEAN_1 SLIP_CLEAN_2 SLIP_CLEAN_3 SLIP_CLEAN_4 SLIP_CLEAN_5
           FAC_CURRENCY FAC_INP_DATE FAC_COM_DATE FAC_EXP_DATE FAC_SUB_CLASS FAC_RISK FAC_DESC FAC_ACC_STS FAC_STS_SLIP;
    set work.facul_cleaned;
run;

proc export
    data=work.facul_output
    outfile="&OUT_FILE."
    dbms=xlsx
    replace;
    sheet="Query result";
run;

%put [FACUL CLEANING ASTRA] Selesai: &OUT_FILE.;
