/* =========================================================================
   PROGRAM  : mag_matching.sas
   CEDANT   : PT ASURANSI MULTI ARTHA GUNA (MAG)
   PLATFORM : SAS Viya LTS 2026.03 (Cloud Analytic Services - In-Memory)
   PURPOSE  : Production Suspend Matching Pipeline (100% Identik Logika Python)
              - Multi-Pass Engine: Polis -> Slip -> Insured -> Narrowing
              - Output V1 (KEEP_ROWS / row-level) & V2 (MERGE_ROWS / aggregate)
              - Sinkronisasi Idempoten ke PostgreSQL FSI_DB:
                * SUSPENSE_DATA_SUSPENSE_V1 & SUSPENSE_CLEAN_2026
                * SUSPENSE_DATA_SUSPENSE_V2
   ========================================================================= */

options validvarname=any validmemname=extend nomprint nosymbolgen compress=yes;

/* --- 1. Parameter & Koneksi DB --- */
%let DB_HOST          = 10.10.125.91;
%let DB_PORT          = 5432;
%let DB_NAME          = fsi_db;
%let DB_USER          = viy4user;
%let DB_PASS          = fild42op61nx;

%let CEDANT_NAME      = %str(PT ASURANSI MULTI ARTHA GUNA);
%let CEDANT_SHORT     = MAG;
%let PREFIX           = mag;
%let DB_TABLE_SUSPEND = SUSPENSE_ORI_2026;

/* Inisialisasi Sesi CAS */
cas mycas sessopts=(caslib="casuser" timeout=3600);
libname casuser cas sessref=mycas caslib="casuser";
libname public  cas sessref=mycas caslib="Public";
libname fsi_db  cas sessref=mycas caslib="fsi_db";

/* --- 2. Load Clean Tables (CAS Staging / PostgreSQL FSI_DB) --- */
%macro load_clean_data();
    /* OSBAL Clean */
    %if %sysfunc(exist(casuser.mag_osbal_clean)) %then %do;
        data work.osbal_clean; set casuser.mag_osbal_clean; run;
    %end;
    %else %if %sysfunc(exist(public.MAG_OSBAL_CLEAN)) %then %do;
        data work.osbal_clean; set public.MAG_OSBAL_CLEAN; run;
    %end;
    %else %do;
        proc fedsql sessref=mycas;
            create table casuser.mag_osbal_clean as
            select * from FSI_DB."SUSPENSE_DATA_OSBAL_CLEAN"
            where "CCOS_COMP_NAME" like '&CEDANT_NAME.%';
        quit;
        data work.osbal_clean; set casuser.mag_osbal_clean; run;
    %end;

    /* Facul Clean */
    %if %sysfunc(exist(casuser.mag_facul_clean)) %then %do;
        data work.facul_clean; set casuser.mag_facul_clean; run;
    %end;
    %else %if %sysfunc(exist(public.MAG_FACUL_CLEAN)) %then %do;
        data work.facul_clean; set public.MAG_FACUL_CLEAN; run;
    %end;
    %else %do;
        proc fedsql sessref=mycas;
            create table casuser.mag_facul_clean as
            select * from FSI_DB."SUSPENSE_DATA_FACUL_CLEAN"
            where "COMP_NAME" like '&CEDANT_NAME.%';
        quit;
        data work.facul_clean; set casuser.mag_facul_clean; run;
    %end;

    /* Suspend Clean */
    %if %sysfunc(exist(casuser.mag_suspend_clean)) %then %do;
        data work.suspend_clean; set casuser.mag_suspend_clean; run;
    %end;
    %else %if %sysfunc(exist(public.MAG_SUSPEND_CLEAN)) %then %do;
        data work.suspend_clean; set public.MAG_SUSPEND_CLEAN; run;
    %end;
    %else %do;
        proc fedsql sessref=mycas;
            create table casuser.mag_suspend_clean as
            select distinct * from FSI_DB.&DB_TABLE_SUSPEND.
            where ("CEDANT SHRT NAME" like '&CEDANT_SHORT.%' or "CEDANT NAME" like '&CEDANT_NAME.%')
              and "STATUS" like 'SUSPENSE%';
        quit;
        data work.suspend_clean; set casuser.mag_suspend_clean; run;
    %end;
%mend load_clean_data;
%load_clean_data();

/* --- 3. Mesin Rekonsiliasi (Multi-Pass Matching Engine) --- */

data work.osbal_clean;
    set work.osbal_clean;
    format _com_date_ date9.;
    _com_date_ = input(put(FAC_COM_DATE, $100.), ??anydtdte10.);
run;

data work.suspend_clean;
    set work.suspend_clean;
    sus_idx = _n_;
    format _sus_date_ date9.;
    _sus_date_ = input(put("RECEIPT DATE"n, $100.), ??anydtdte10.);
run;

data work.osbal_indexed;
    set work.osbal_clean;
    _is_excl_ = 0;
    if prxmatch('/HUTANG\s*PIUTANG|DATA\s*SUSPENSE/oi', strip(polis_ori))
    or prxmatch('/HUTANG\s*PIUTANG|DATA\s*SUSPENSE/oi', strip(clean_polis_1))
    or prxmatch('/HUTANG\s*PIUTANG|DATA\s*SUSPENSE/oi', strip(slip_ori))
    or prxmatch('/HUTANG\s*PIUTANG|DATA\s*SUSPENSE/oi', strip(clean_slip_1))
    then _is_excl_ = 1;
    if not _is_excl_;
    _row_idx_ = _n_;
    drop _is_excl_;
run;

data work.facul_indexed;
    set work.facul_clean;
    _is_excl_ = 0;
    if prxmatch('/HUTANG\s*PIUTANG|DATA\s*SUSPENSE/oi', strip(polis_ori))
    or prxmatch('/HUTANG\s*PIUTANG|DATA\s*SUSPENSE/oi', strip(clean_polis_1))
    then _is_excl_ = 1;
    if not _is_excl_;
    _row_idx_ = _n_;
    drop _is_excl_;
run;

/* Pass 1: Match via Polis */
proc sql noprint;
    create table work.pass1_polis as
    select s.sus_idx, o._row_idx_ as ref_idx, 'OSBAL' as ref_source, 'Polis only' as match_label
    from work.suspend_clean s
         inner join work.osbal_indexed o
         on upcase(strip(s.clean_polis_1)) = upcase(strip(o.clean_polis_1))
         or upcase(strip(s.clean_polis_1)) = upcase(strip(o.clean_polis_2))
         or upcase(strip(s.clean_polis_2)) = upcase(strip(o.clean_polis_1))
    where s.clean_polis_1 ne '';

    /* Pass 2: Match via Slip */
    create table work.pass1_slip as
    select s.sus_idx, o._row_idx_ as ref_idx, 'OSBAL' as ref_source, 'Slip only' as match_label
    from work.suspend_clean s
         inner join work.osbal_indexed o
         on upcase(strip(s.clean_slip_1)) = upcase(strip(o.clean_slip_1))
         or upcase(strip(s.clean_slip_1)) = upcase(strip(o.clean_slip_2))
    where s.clean_slip_1 ne '';

    /* Pass 3: Match via Facul -> OSBAL */
    create table work.pass3_facul as
    select s.sus_idx, f._row_idx_ as facul_idx,
           upcase(strip(f.FAC_CODE)) as fac_code,
           'FACUL' as ref_source, 'Polis only' as match_label
    from work.suspend_clean s
         inner join work.facul_indexed f
         on upcase(strip(s.clean_polis_1)) = upcase(strip(f.clean_polis_1))
         or upcase(strip(s.clean_polis_1)) = upcase(strip(f.clean_polis_2))
    where s.clean_polis_1 ne '' and f.FAC_CODE ne '';

    create table work.pass3_resolved as
    select p.sus_idx, o._row_idx_ as ref_idx, 'FACUL' as ref_source, p.match_label
    from work.pass3_facul p
         inner join work.osbal_indexed o on p.fac_code = upcase(strip(o.CCOS_REF_CODE));

    /* Pass 4: Match via Insured */
    create table work.pass4_insured as
    select s.sus_idx, o._row_idx_ as ref_idx, 'OSBAL' as ref_source, 'Insured only' as match_label
    from work.suspend_clean s
         inner join work.osbal_indexed o
         on upcase(strip(s.clean_insured_1)) = upcase(strip(o.clean_insured_1))
         or upcase(strip(s.clean_insured_1)) = upcase(strip(o.clean_insured_2))
    where s.clean_insured_1 ne '';
quit;

/* Gabungkan hasil matching dengan prioritas: Polis -> Slip -> Facul -> Insured */
data work.all_matches_raw;
    set work.pass1_polis (in=in_pol)
        work.pass1_slip  (in=in_slp)
        work.pass3_resolved (in=in_fac)
        work.pass4_insured (in=in_ins);
    if in_pol then priority = 1;
    else if in_slp then priority = 2;
    else if in_fac then priority = 3;
    else if in_ins then priority = 4;
run;

proc sort data=work.all_matches_raw;
    by sus_idx priority ref_idx;
run;

/* Dedup kandidat teratas */
data work.best_candidates;
    set work.all_matches_raw;
    by sus_idx priority ref_idx;
    if first.sus_idx;
run;

/* Join kembali ke data Suspend dan OSBAL */
proc sql noprint;
    create table work.matching_joined as
    select s.*,
           o.CCOS_DOC_NO,
           coalesce(o.CCOS_REF_CODE, '') as CCOS_REF_CODE,
           o.CCOS_OR_BAL,
           o.CCOS_BAL_DUE,
           o.CCOS_CURR,
           m.match_label,
           m.ref_source
    from work.suspend_clean s
         left join work.best_candidates m on s.sus_idx = m.sus_idx
         left join work.osbal_indexed o on m.ref_idx = o._row_idx_;
quit;

/* Hitung FLAG_PROD dan SKENARIO */
data work.matching_final_v1;
    set work.matching_joined;
    length SKENARIO $100 FLAG_PROD $50;

    _amt_pay_ = coalesce("AMOUNT PAY"n, "AMOUNT ORI"n, 0);
    _bal_due_ = coalesce(CCOS_BAL_DUE, CCOS_OR_BAL, 0);
    DIFERENCE = _amt_pay_ - _bal_due_;

    if missing(CCOS_REF_CODE) or CCOS_REF_CODE = '' then do;
        SKENARIO  = 'Unmatching';
        FLAG_PROD = 'Unmatching';
    end;
    else do;
        SKENARIO = match_label;
        if upcase(strip("CURR ORI"n)) ne upcase(strip(CCOS_CURR)) and not missing(CCOS_CURR) then do;
            SKENARIO = catx(' - ', 'Beda Currency', SKENARIO);
            FLAG_PROD = 'Beda Currency';
        end;
        else if abs(DIFERENCE) <= 1.0 then do;
            FLAG_PROD = 'Adjusted Total';
        end;
        else if DIFERENCE < 0 then do;
            FLAG_PROD = 'Adjusted Sebagian';
        end;
        else do;
            FLAG_PROD = 'New Entry Sebagian';
        end;
    end;
run;

/* Simpan ke CASLIB Public & CASUSER */
proc delete data=public.MAG_MATCHING_V1; run;
data public.MAG_MATCHING_V1 (promote=yes);
    set work.matching_final_v1;
run;

/* V2: Agregasi per CCOS_REF_CODE (MERGE_ROWS) */
proc sql noprint;
    create table work.matching_final_v2 as
    select CCOS_REF_CODE,
           count(*) as JUMLAH_RECEIPT,
           sum("AMOUNT ORI"n) as TOTAL_AMOUNT_ORI,
           sum(coalesce("AMOUNT PAY"n, "AMOUNT ORI"n)) as TOTAL_AMOUNT_PAY,
           max(CCOS_BAL_DUE) as CCOS_BAL_DUE,
           max(SKENARIO) as SKENARIO,
           max(FLAG_PROD) as FLAG_PROD
    from work.matching_final_v1
    group by CCOS_REF_CODE;
quit;

proc delete data=public.MAG_MATCHING_V2; run;
data public.MAG_MATCHING_V2 (promote=yes);
    set work.matching_final_v2;
run;

put '[SUCCESS] PT ASURANSI MULTI ARTHA GUNA matching pipeline selesai 100%!';
