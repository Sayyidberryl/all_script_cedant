/* =============================================================================
   PROGRAM   : CLEANING_OSBAL.SAS
   CEDANT    : PT ASURANSI ASTRA BUANA
   FUNGSI    : Pembersihan data tagihan OSBAL Astra Buana (dengan prioritas CLSDT)
   INPUT     : data/osbal.xlsx (atau 2b. transaksi Osbal 01.01.23 - 17.08.26.xlsx)
   OUTPUT    : data/astrabuana_output_osbal.xlsx
   KOMPATIBEL: SAS Viya / SAS 9.4 Studio
   ============================================================================= */

options validvarname=any validmemname=extend mprint symbolgen compress=yes;

/* ── 0. Path setup ─────────────────────────────────────────────────────────── */
%if not %symexist(BASE_DIR) %then %do;
    %let BASE_DIR = %str(c:\Users\berryl\Desktop\starcore\indore\all_script_cedant\astra_production_script);
%end;
%let DATA_DIR   = %str(&BASE_DIR.\data);
%let OUT_FILE   = %str(&DATA_DIR.\astrabuana_output_osbal.xlsx);
%let MACRO_INC  = %str(&BASE_DIR.\sas_ready\macros_cleaning_astra.sas);

/* Cari file input yang tersedia */
%let INPUT_FILE = %str(&DATA_DIR.\osbal.xlsx);
%macro detect_input();
    %if not %sysfunc(fileexist("&INPUT_FILE.")) %then %do;
        %let ALT_FILE = %str(&BASE_DIR.\dataExcel\raw\2b. transaksi Osbal 01.01.23 - 17.08.26.xlsx);
        %if %sysfunc(fileexist("&ALT_FILE.")) %then %let INPUT_FILE = &ALT_FILE.;
    %end;
%mend detect_input;
%detect_input();

/* ── 1. Include shared macros ──────────────────────────────────────────────── */
%if %sysfunc(fileexist("&MACRO_INC.")) %then %do;
    %include "&MACRO_INC.";
%end;

/* ── 2. Definisi konstanta bisnis ─────────────────────────────────────────── */
%let CEDANT_COL  = CCOS_COMP_NAME;
%let CEDANT_VAL  = %str(PT ASURANSI ASTRA BUANA);
%let TABLE_NAME  = SUSPENSE_DATA_OSBAL_CLEAN;

%define_bulan_format();


/* =============================================================================
   STEP 1: Baca file Excel OSBAL
   ============================================================================= */
%put [1/4] Reading: &INPUT_FILE. ...;

proc import
    datafile="&INPUT_FILE."
    out=work.osbal_raw
    dbms=xlsx
    replace;
    getnames=yes;
    datarow=2;
run;

%put       Total rows dibaca: %sysfunc(attrn(%sysfunc(open(work.osbal_raw)), nobs));


/* =============================================================================
   STEP 2: Filter Cedant PT ASURANSI ASTRA BUANA
   ============================================================================= */
%put [2/4] Filtering cedant: &CEDANT_VAL. ...;

data work.osbal_filtered;
    set work.osbal_raw;
    _cedant_ = strip(&CEDANT_COL.);
    if _cedant_ = "&CEDANT_VAL." then output;
    drop _cedant_;
run;

%put       Rows setelah filter: %sysfunc(attrn(%sysfunc(open(work.osbal_filtered)), nobs));


/* =============================================================================
   STEP 3: Cek Kolom CLSDT & Cleaning dengan Prioritas CLSDT -> FAC
   ============================================================================= */
%let HAS_CLSDT_POLIS = %eval(%varexist(work.osbal_filtered, CLSDT_POLICY_NO) > 0);
%let HAS_CLSDT_SLIP  = %eval(%varexist(work.osbal_filtered, CLSDT_SLIP_NO)  > 0);
%let HAS_CLSDT_SERTF = %eval(%varexist(work.osbal_filtered, CLSDT_SERTF_NO) > 0);

%macro is_clsdt_valid_astra(val, result_var);
    &result_var = 0;
    if not missing(&val.) and strip(&val.) ne '' then do;
        /* Abaikan kata generik seperti SISA OSBAL, TBA, VARIOUS */
        if not prxmatch('/\b(?:SISA|OSBAL|TBA|VARIOUS|VAR)\b/oi', strip(&val.)) then do;
            if prxmatch('/\d{6,}/', strip(&val.)) or lengthn(strip(&val.)) >= 6 then &result_var = 1;
        end;
    end;
%mend is_clsdt_valid_astra;

%macro run_osbal_cleaning();
data work.osbal_cleaned;
    set work.osbal_filtered;

    length INSURED_CLEAN_1 - INSURED_CLEAN_5 $500
           POLICY_CLEAN_1  - POLICY_CLEAN_5  $200
           SERTIF_CLEAN_1  - SERTIF_CLEAN_3  $50
           SLIP_CLEAN_1    - SLIP_CLEAN_5    $200;

    /* 1. Clean Insured */
    %apply_clean_insured_astra(FAC_INSURED);

    /* 2. Clean Polis: prioritas CLSDT -> fallback FAC */
    %if &HAS_CLSDT_POLIS. %then %do;
        %apply_clean_polis_astra(CLSDT_POLICY_NO);
        _cls_pol_valid_ = 0;
        %is_clsdt_valid_astra(POLICY_CLEAN_1, _cls_pol_valid_);

        array _pol_cls{5} $200 _pol_cls_1 - _pol_cls_5;
        _pol_cls_1 = POLICY_CLEAN_1; _pol_cls_2 = POLICY_CLEAN_2;
        _pol_cls_3 = POLICY_CLEAN_3; _pol_cls_4 = POLICY_CLEAN_4;
        _pol_cls_5 = POLICY_CLEAN_5;

        %apply_clean_polis_astra(FAC_POLICY_NO);

        if _cls_pol_valid_ then do;
            POLICY_CLEAN_1 = _pol_cls_1; POLICY_CLEAN_2 = _pol_cls_2;
            POLICY_CLEAN_3 = _pol_cls_3; POLICY_CLEAN_4 = _pol_cls_4;
            POLICY_CLEAN_5 = _pol_cls_5;
        end;
        drop _pol_cls_1 - _pol_cls_5 _cls_pol_valid_;
    %end;
    %else %do;
        %apply_clean_polis_astra(FAC_POLICY_NO);
    %end;

    /* 3. Clean Sertifikat dari CLSDT_SERTF_NO */
    %if &HAS_CLSDT_SERTF. %then %do;
        _sertf_raw_ = strip(CLSDT_SERTF_NO);
        if prxmatch('/\.0$/', _sertf_raw_) then
            _sertf_raw_ = substr(_sertf_raw_, 1, lengthn(_sertf_raw_) - 2);
        SERTIF_CLEAN_1 = '';
        if not missing(_sertf_raw_) and _sertf_raw_ notin ('', '#', '-') then do;
            if prxmatch('/^\d{1,6}$/', strip(_sertf_raw_)) then
                SERTIF_CLEAN_1 = put(input(strip(_sertf_raw_), ??8.), z6.);
            else
                SERTIF_CLEAN_1 = strip(_sertf_raw_);
        end;
        drop _sertf_raw_;
    %end;
    %else %do;
        SERTIF_CLEAN_1 = '';
    %end;
    SERTIF_CLEAN_2 = ''; SERTIF_CLEAN_3 = '';

    /* 4. Clean Slip: prioritas CLSDT -> fallback FAC */
    %if &HAS_CLSDT_SLIP. %then %do;
        %apply_clean_slip_astra(CLSDT_SLIP_NO);
        _cls_slp_valid_ = 0;
        %is_clsdt_valid_astra(SLIP_CLEAN_1, _cls_slp_valid_);

        array _slp_cls{5} $200 _slp_cls_1 - _slp_cls_5;
        _slp_cls_1 = SLIP_CLEAN_1; _slp_cls_2 = SLIP_CLEAN_2;
        _slp_cls_3 = SLIP_CLEAN_3; _slp_cls_4 = SLIP_CLEAN_4;
        _slp_cls_5 = SLIP_CLEAN_5;

        %apply_clean_slip_astra(FAC_SLIP);

        if _cls_slp_valid_ then do;
            SLIP_CLEAN_1 = _slp_cls_1; SLIP_CLEAN_2 = _slp_cls_2;
            SLIP_CLEAN_3 = _slp_cls_3; SLIP_CLEAN_4 = _slp_cls_4;
            SLIP_CLEAN_5 = _slp_cls_5;
        end;
        drop _slp_cls_1 - _slp_cls_5 _cls_slp_valid_;
    %end;
    %else %do;
        %apply_clean_slip_astra(FAC_SLIP);
    %end;

    /* Clean control characters */
    array _all_str_{*} _character_;
    do _si_ = 1 to dim(_all_str_);
        if not missing(_all_str_{_si_}) then
            %clean_control_chars(_all_str_{_si_});
    end;
    drop _si_;
run;
%mend run_osbal_cleaning;
%run_osbal_cleaning();


/* =============================================================================
   STEP 4: Standarisasi Urutan Kolom & Export ke Excel
   ============================================================================= */
%put [4/4] Building output columns & saving to: &OUT_FILE. ...;

data work.osbal_output;
    retain CCOS_DOC_NO CCOS_DATE CCOS_REF_CODE CCOS_COMP CCOS_COMP_NAME CCOS_REF_COMP CCOS_REF_COMP_NAME
           FAC_INSURED INSURED_CLEAN_1 INSURED_CLEAN_2 INSURED_CLEAN_3 INSURED_CLEAN_4 INSURED_CLEAN_5
           CCOS_CURR CCOS_OR_BAL CCOS_BAL_DUE CCOS_OR_BAL_IN_IDR CCOS_BAL_DUE_IN_IDR
           FAC_COM_DATE FAC_EXP_DATE FAC_DUE_DATES FAC_SUB_CLASS
           FAC_POLICY_NO POLICY_CLEAN_1 SERTIF_CLEAN_1 POLICY_CLEAN_2 SERTIF_CLEAN_2 POLICY_CLEAN_3 SERTIF_CLEAN_3
           POLICY_CLEAN_4 POLICY_CLEAN_5
           FAC_SLIP SLIP_CLEAN_1 SLIP_CLEAN_2 SLIP_CLEAN_3 SLIP_CLEAN_4 SLIP_CLEAN_5
           CLASS_CODE CLASS_NAME CLSDT_POLICY_NO CLSDT_SLIP_NO CLSDT_SERTF_NO;
    set work.osbal_cleaned;
run;

proc export
    data=work.osbal_output
    outfile="&OUT_FILE."
    dbms=xlsx
    replace;
    sheet="Query result";
run;

%put [OSBAL CLEANING ASTRA] Selesai: &OUT_FILE.;
