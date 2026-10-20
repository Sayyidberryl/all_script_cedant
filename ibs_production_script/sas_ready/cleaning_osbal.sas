/* =============================================================================
   CLEANING_OSBAL.SAS
   Konversi dari: cleaning_osbal.py

   Fungsi:
   - Load data/osbal.xlsx
   - Filter cedant = "PT ASURANSI IBS" ATAU "PT ASURANSI IBS SYARIAH"
   - Kolom CLSDT (opsional): CLSDT_POLICY_NO, CLSDT_SLIP_NO, CLSDT_SERTF_NO
   - Logika prioritas CLSDT: jika CLSDT valid -> gunakan CLSDT, else fallback ke FAC
   - Clean polis, slip, insured (maks 5 kolom masing-masing)
   - Clean sertif dari kolom CLSDT_SERTF_NO (1-6 digit numerik -> zero-pad 6 digit)
   - Output Excel: data/ibs_output_osbal_V2.xlsx
   - Export ke PostgreSQL (tabel SUSPENSE_DATA_OSBAL_CLEAN)

   Perbedaan kunci vs cleaning_facul:
   1. Filter ISIN (2 nilai cedant, bukan 1)
   2. Ada kolom CLSDT dengan logika prioritas
   3. clean_sertif berasal dari CLSDT_SERTF_NO (bukan ekstrak dari polis)
   4. clean_polis OSBAL menggunakan validasi lebih ketat (>= 10 digit berturutan)
   ============================================================================= */

options validvarname=any validmemname=extend mprint symbolgen;

/* ── 0. Path setup ─────────────────────────────────────────────────────────── */
%if not %symexist(BASE_DIR) %then %do;
    %let BASE_DIR = %str(c:\Users\berryl\Desktop\starcore\indore\all_script_cedant\ibs_production_script);
%end;
%let DATA_DIR   = %str(&BASE_DIR.\data);
%let INPUT_FILE = %str(&DATA_DIR.\osbal.xlsx);
%let OUT_FILE   = %str(&DATA_DIR.\ibs_output_osbal_V2.xlsx);
%let MACRO_INC  = %str(&BASE_DIR.\sas_ready\macros_cleaning.sas);

/* ── 1. Include shared macros ──────────────────────────────────────────────── */
%if %sysfunc(fileexist("&MACRO_INC.")) %then %do;
    %include "&MACRO_INC.";
%end;

/* ── 2. Konstanta bisnis ──────────────────────────────────────────────────── */
%let CEDANT_COL  = CCOS_COMP_NAME;
%let POLIS_COL   = FAC_POLICY_NO;
%let SLIP_COL    = FAC_SLIP;
%let INSURED_COL = FAC_INSURED;
/* Kolom CLSDT */
%let CLSDT_POLIS = CLSDT_POLICY_NO;
%let CLSDT_SLIP  = CLSDT_SLIP_NO;
%let CLSDT_SERTF = CLSDT_SERTF_NO;
%let TABLE_NAME  = SUSPENSE_DATA_OSBAL_CLEAN;

%compile_cleaning_regex();
%define_bulan_format();


/* =============================================================================
   STEP 1: Baca file Excel
   ============================================================================= */
%put [1/5] Reading: &INPUT_FILE.;

proc import
    datafile="&INPUT_FILE."
    out=work.osbal_raw
    dbms=xlsx
    replace;
    getnames=yes;
    datarow=2;
run;

%put       Total rows: %sysfunc(attrn(%sysfunc(open(work.osbal_raw)), nobs));


/* =============================================================================
   STEP 2: Filter cedant (2 nilai)
   Setara dengan: df[df[CEDANT_COL].isin(CEDANT_VALUES)]
   ============================================================================= */
%put [2/5] Filtering cedant (ACA / ACA Syariah) ...;

data work.osbal_filtered;
    set work.osbal_raw;
    _cedant_ = strip(&CEDANT_COL.);
    if _cedant_ in ('PT ASURANSI IBS',
                    'PT ASURANSI IBS SYARIAH') then output;
    drop _cedant_;
run;

%put       Rows setelah filter: %sysfunc(attrn(%sysfunc(open(work.osbal_filtered)), nobs));


/* =============================================================================
   STEP 3: Rename kolom utama
   ============================================================================= */
data work.osbal_renamed;
    set work.osbal_filtered;
    rename &POLIS_COL.   = polis_ori
           &SLIP_COL.    = slip_ori
           &INSURED_COL. = insured_ori;
run;


/* =============================================================================
   STEP 4: Cek ketersediaan kolom CLSDT (opsional)
   Setara dengan: has_clsdt_polis = CLSDT_POLIS_COL in df.columns
   ============================================================================= */

/* Macro helper: cek apakah variabel ada di dataset */
%macro varexist(ds, var);
    %local dsid rc;
    %let dsid = %sysfunc(open(&ds.));
    %if &dsid. %then %do;
        %let rc = %sysfunc(varnum(&dsid., &var.));
        %let dsid = %sysfunc(close(&dsid.));
        &rc.
    %end;
    %else 0;
%mend varexist;

%let HAS_CLSDT_POLIS = %eval(%varexist(work.osbal_renamed, &CLSDT_POLIS.) > 0);
%let HAS_CLSDT_SLIP  = %eval(%varexist(work.osbal_renamed, &CLSDT_SLIP.)  > 0);
%let HAS_CLSDT_SERTF = %eval(%varexist(work.osbal_renamed, &CLSDT_SERTF.) > 0);

%put       CLSDT Polis tersedia: &HAS_CLSDT_POLIS.;
%put       CLSDT Slip tersedia : &HAS_CLSDT_SLIP.;
%put       CLSDT Sertf tersedia: &HAS_CLSDT_SERTF.;


/* =============================================================================
   STEP 5: Macro untuk validasi hasil CLSDT
   Setara dengan: _clsdt_result_is_valid(tokens) di Python

   Logika:
   - Mengandung >= 10 digit berturutan -> polis ACA / slip panjang -> VALID
   - Mengandung >= 7 digit berturutan DAN bukan TBA -> slip pendek -> VALID
   - Lainnya -> TIDAK VALID (gunakan fallback FAC)
   ============================================================================= */
%macro is_clsdt_valid(val, result_var);
    &result_var = 0;
    if not missing(&val.) and strip(&val.) ne '' then do;
        if prxmatch('/\d{10,}/', strip(&val.)) then &result_var = 1;
        else if prxmatch('/\d{7,}/', strip(&val.))
             and not prxmatch('/\bTBA\b/oi', strip(&val.)) then &result_var = 1;
    end;
%mend is_clsdt_valid;


/* =============================================================================
   STEP 6: Cleaning -- polis, slip, insured, sertif (dengan logika CLSDT)
   ============================================================================= */
%put [3/5] Cleaning (dengan prioritas CLSDT) ...;

%macro run_osbal_clean_step();
data work.osbal_cleaned;
    set work.osbal_renamed;

    length clean_polis_1   - clean_polis_5   $500
           clean_slip_1    - clean_slip_5    $500
           clean_insured_1 - clean_insured_5 $500
           clean_sertif_1  - clean_sertif_5  $10;

    /* -- 6a. Clean POLIS dengan prioritas CLSDT ----------------------------- */
    %if &HAS_CLSDT_POLIS. %then %do;
        %apply_clean_polis(&CLSDT_POLIS., pfx=_cp_clsdt, tag=clsdt_p);
        %apply_clean_polis(polis_ori, pfx=clean_polis, tag=fac_p);

        _clsdt_polis_valid_ = 0;
        %is_clsdt_valid_array(_cp_clsdt, _clsdt_polis_valid_);
        if _clsdt_polis_valid_ then do;
            clean_polis_1 = _cp_clsdt_1; clean_polis_2 = _cp_clsdt_2;
            clean_polis_3 = _cp_clsdt_3; clean_polis_4 = _cp_clsdt_4;
            clean_polis_5 = _cp_clsdt_5;
        end;
        drop _cp_clsdt_1 - _cp_clsdt_5 _clsdt_polis_valid_;
    %end;
    %else %do;
        %apply_clean_polis(polis_ori, pfx=clean_polis, tag=fac_p);
    %end;

    /* -- 6b. Clean SLIP dengan prioritas CLSDT ------------------------------ */
    %if &HAS_CLSDT_SLIP. %then %do;
        %apply_clean_slip(&CLSDT_SLIP., pfx=_csl_clsdt, tag=clsdt_s);
        %apply_clean_slip(slip_ori, pfx=clean_slip, tag=fac_s);

        _clsdt_slip_valid_ = 0;
        %is_clsdt_valid_array(_csl_clsdt, _clsdt_slip_valid_);
        if _clsdt_slip_valid_ then do;
            clean_slip_1 = _csl_clsdt_1; clean_slip_2 = _csl_clsdt_2;
            clean_slip_3 = _csl_clsdt_3; clean_slip_4 = _csl_clsdt_4;
            clean_slip_5 = _csl_clsdt_5;
        end;
        drop _csl_clsdt_1 - _csl_clsdt_5 _clsdt_slip_valid_;
    %end;
    %else %do;
        %apply_clean_slip(slip_ori, pfx=clean_slip, tag=fac_s);
    %end;

    /* -- 6c. Clean INSURED (selalu dari FAC, CLSDT tidak berpengaruh) ------- */
    %apply_clean_insured(insured_ori);

    /* -- 6d. Clean SERTIFIKAT (equiv Python extract_cert_from_polis_osbal + clean_sertif) -- */
    %if &HAS_CLSDT_SERTF. %then %do;
        %apply_extract_sertif_osbal(polis_ori, &CLSDT_SERTF.);
    %end;
    %else %do;
        _dummy_sertf_ = '';
        %apply_extract_sertif_osbal(polis_ori, _dummy_sertf_);
        drop _dummy_sertf_;
    %end;

    /* -- 6e. Clean control chars -------------------------------------------- */
    array _str_all{*} _character_;
    do _sci_ = 1 to dim(_str_all);
        if not missing(_str_all{_sci_}) then
            %clean_control_chars(_str_all{_sci_});
    end;
    drop _sci_;

run;
%mend run_osbal_clean_step;
%run_osbal_clean_step();

%put       Cleaning selesai.;


/* =============================================================================
   STEP 7: Susun urutan kolom output
   ============================================================================= */
%put [4/5] Building output columns ...;

data work.osbal_output;
    retain &CEDANT_COL.
           polis_ori    clean_polis_1-clean_polis_5   clean_sertif_1-clean_sertif_5
           slip_ori     clean_slip_1-clean_slip_5
           insured_ori  clean_insured_1-clean_insured_5;
    set work.osbal_cleaned;
run;


/* =============================================================================
   STEP 8: Simpan ke Excel
   ============================================================================= */
%put [5/5] Saving to: &OUT_FILE.;

proc export
    data=work.osbal_output
    outfile="&OUT_FILE."
    dbms=xlsx
    replace;
    sheet="osbal_clean";
run;

%put Done. Output: &OUT_FILE.;


/* =============================================================================
   STEP 9: Export ke PostgreSQL (opsional -- uncomment sesuai environment)
   ============================================================================= */

/* --- Via ODBC --- */
/*
libname pg_lib odbc
    datasrc="&ODBC_DSN."
    user="&DB_USER."
    password="&DB_PASS."
    schema="public";

proc append
    base=pg_lib.&TABLE_NAME.
    data=work.osbal_output
    force;
run;

libname pg_lib clear;
*/

%put NOTE: Export PostgreSQL/CAS - uncomment sesuai environment Anda.;
%put [OSBAL CLEANING] Selesai.;

/* Cleanup */
proc datasets library=work nolist;
    delete osbal_raw osbal_filtered osbal_renamed osbal_cleaned;
run;
quit;
