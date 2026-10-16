/* =============================================================================
   PROGRAM   : PROD_MATCHING.SAS
   CEDANT    : PT ASURANSI ASTRA BUANA
   FUNGSI    : Rekonsiliasi & Matching data SUSPEND terhadap OSBAL / FACUL / BORDERO
   INPUT     : data/astrabuana_output_suspend.xlsx
               data/astrabuana_output_osbal.xlsx
               data/astrabuana_output_facul.xlsx
               data/BORDERO (ASTRA_OPEN_COVER_FIRE.xlsx, MC.xlsx, MH.xlsx)
   OUTPUT    : data/final_output_astrabuana_v1.xlsx (atau v2 jika VERSION=2)
   KOMPATIBEL: SAS Viya / SAS 9.4 Studio
   ============================================================================= */

options validvarname=any validmemname=extend mprint symbolgen compress=yes;

/* =============================================================================
   SECTION 0: PARAMETER & PATH SETUP
   ============================================================================= */
%let VERSION         = 1;            /* 1 = V1 (row asli dipertahankan) | 2 = V2 (gabung baris) */
%let ACCUMULATE_MODE = KEEP_ROWS;   /* KEEP_ROWS (V1) atau MERGE_ROWS (V2) */

%if not %symexist(BASE_DIR) %then %do;
    %let BASE_DIR = %str(c:\Users\berryl\Desktop\starcore\indore\all_script_cedant\astra_production_script);
%end;
%let DATA_DIR    = %str(&BASE_DIR.\data);
%let MACRO_INC   = %str(&BASE_DIR.\sas_ready\macros_cleaning_astra.sas);

%let SUSPEND_FILE = %str(&DATA_DIR.\astrabuana_output_suspend.xlsx);
%let OSBAL_FILE   = %str(&DATA_DIR.\astrabuana_output_osbal.xlsx);
%let FACUL_FILE   = %str(&DATA_DIR.\astrabuana_output_facul.xlsx);
%let OUTPUT_FILE  = %str(&DATA_DIR.\final_output_astrabuana_v&VERSION..xlsx);

/* Folder Bordero */
%let BORDERO_DIR   = %str(&DATA_DIR.\BORDERO);
%let BDR_FIRE_FILE = %str(&BORDERO_DIR.\ASTRA_OPEN_COVER_FIRE.xlsx);
%let BDR_MC_FILE   = %str(&BORDERO_DIR.\ASTRA_OPEN_COVER_MC.xlsx);
%let BDR_MH_FILE   = %str(&BORDERO_DIR.\ASTRA_OPEN_COVER_MH.xlsx);

/* Include shared macros */
%if %sysfunc(fileexist("&MACRO_INC.")) %then %do;
    %include "&MACRO_INC.";
%end;
%define_bulan_format();

%let OSBAL_FACODE_COL = CCOS_REF_CODE;
%let FACUL_FACODE_COL = FAC_CODE;
%let OSBAL_CURR_COL   = CCOS_CURR;


/* =============================================================================
   SECTION 1: LOAD CLEANED DATA
   ============================================================================= */
%put [1/6] Loading cleaned reference data ...;

proc import datafile="&SUSPEND_FILE." out=work.suspend_raw dbms=xlsx replace; getnames=yes; datarow=2; run;
proc import datafile="&OSBAL_FILE."   out=work.osbal_raw   dbms=xlsx replace; getnames=yes; datarow=2; run;
proc import datafile="&FACUL_FILE."   out=work.facul_raw   dbms=xlsx replace; getnames=yes; datarow=2; run;

/* Load Bordero (FIRE, MC, MH) jika tersedia */
%macro load_bordero_if_exists(filepath, outds);
    %if %sysfunc(fileexist("&filepath.")) %then %do;
        proc import datafile="&filepath." out=&outds. dbms=xlsx replace; getnames=yes; datarow=2; run;
    %end;
    %else %do;
        data &outds.;
            length "FAC CODE"n $50 POLIS $200 SLIP $200 CERTIFICATE $50 INSURED $500 CURR $10 BULAN $50 TAHUN 8 NET 8;
            stop;
        run;
    %end;
%mend load_bordero_if_exists;

%load_bordero_if_exists(&BDR_FIRE_FILE., work.bdr_fire);
%load_bordero_if_exists(&BDR_MC_FILE.,   work.bdr_mc);
%load_bordero_if_exists(&BDR_MH_FILE.,   work.bdr_mh);

/* Gabungkan Bordero files */
data work.bordero_combined;
    set work.bdr_fire (in=f) work.bdr_mc (in=mc) work.bdr_mh (in=mh);
    length bdr_src $10 fac_code $50 bdr_polis $200 bdr_slip $200 bdr_cert $50 bdr_insured $500 bdr_curr $10;
    if f then bdr_src = 'FIRE';
    else if mc then bdr_src = 'MC';
    else if mh then bdr_src = 'MH';

    fac_code    = upcase(strip("FAC CODE"n));
    bdr_polis   = upcase(strip(POLIS));
    bdr_slip    = upcase(strip(SLIP));
    bdr_cert    = upcase(strip(CERTIFICATE));
    bdr_insured = upcase(strip(INSURED));
    bdr_curr    = upcase(strip(CURR));

    /* Parse Bulan & Tahun ke Date */
    _b_no_ = input(put(upcase(strip(BULAN)), $bulan_fmt.), 8.);
    _t_no_ = input(strip(put(TAHUN, 8.)), ??8.);
    if _b_no_ > 0 and _t_no_ > 0 then bdr_period = mdy(_b_no_, 1, _t_no_);
    else bdr_period = .;
    format bdr_period date9.;

    if fac_code ne '';
    drop _b_no_ _t_no_;
run;


/* =============================================================================
   SECTION 2: PRE-PROCESSING & INDEXING
   ============================================================================= */
%put [2/6] Pre-processing & parsing dates ...;

/* Suspend pre-parse: tambahkan sus_idx dan deteksi LINESLIP */
data work.suspend_clean;
    set work.suspend_raw;
    sus_idx = _n_;

    /* Parse RECEIPT DATE */
    format _sus_date_ date9. _sus_date_ls_ date9.;
    _sus_date_ = input(put("RECEIPT DATE"n, $100.), anydtdte10.);

    /* Deteksi LINESLIP di INSURED */
    _is_lineslip_ = 0;
    if prxmatch('/LINESLIP|LINE\s*SLIP/oi', strip(INSURED)) or
       prxmatch('/LINESLIP|LINE\s*SLIP/oi', strip(INSURED_CLEAN)) then
        _is_lineslip_ = 1;

    if _is_lineslip_ and _sus_date_ ne . then
        _sus_date_ls_ = intnx('month', _sus_date_, -1, 'same');
    else _sus_date_ls_ = .;
run;

/* OSBAL pre-parse */
data work.osbal_indexed;
    set work.osbal_raw;
    _row_idx_ = _n_;

    /* Filter marker exclude */
    _is_excl_ = 0;
    if prxmatch('/HUTANG\s*PIUTANG|DATA\s*SUSPENSE/oi', strip(FAC_POLICY_NO)) or
       prxmatch('/HUTANG\s*PIUTANG|DATA\s*SUSPENSE/oi', strip(POLICY_CLEAN_1)) or
       prxmatch('/HUTANG\s*PIUTANG|DATA\s*SUSPENSE/oi', strip(FAC_SLIP)) or
       prxmatch('/HUTANG\s*PIUTANG|DATA\s*SUSPENSE/oi', strip(SLIP_CLEAN_1)) then
        _is_excl_ = 1;
    if not _is_excl_;

    /* Parse FAC_COM_DATE */
    format _com_date_ date9.;
    _com_date_ = input(put(FAC_COM_DATE, $100.), anydtdte10.);
    drop _is_excl_;
run;

/* FACUL pre-parse */
data work.facul_indexed;
    set work.facul_raw;
    _row_idx_ = _n_;

    _is_excl_ = 0;
    if prxmatch('/HUTANG\s*PIUTANG|DATA\s*SUSPENSE/oi', strip(FAC_POLICY_NO)) or
       prxmatch('/HUTANG\s*PIUTANG|DATA\s*SUSPENSE/oi', strip(POLICY_CLEAN_1)) then
        _is_excl_ = 1;
    if not _is_excl_;
    drop _is_excl_;
run;


/* =============================================================================
   SECTION 3: MATCHING PASSES (POLIS, SLIP, INSURED, FACUL)
   ============================================================================= */
%put [3/6] Running matching passes ...;

proc sql noprint;
    /* Pass 1: Polis Exact Match -> OSBAL */
    create table work.match_p1_polis as
    select s.sus_idx, o._row_idx_ as ref_idx, 'OSBAL' as ref_source, 'Polis only' as match_label
    from work.suspend_clean s
         inner join work.osbal_indexed o
         on (upcase(strip(s.POLIS_CLN)) ne '' and (
             upcase(strip(s.POLIS_CLN)) = upcase(strip(o.POLICY_CLEAN_1)) or
             upcase(strip(s.POLIS_CLN)) = upcase(strip(o.POLICY_CLEAN_2)) or
             upcase(strip(s.POLIS_CLN)) = upcase(strip(o.POLICY_CLEAN_3)) or
             upcase(strip(s.POLIS_CLN)) = upcase(strip(o.POLICY_CLEAN_4)) or
             upcase(strip(s.POLIS_CLN)) = upcase(strip(o.POLICY_CLEAN_5)) or
             upcase(strip(s.POLIS_CLN)) = upcase(strip(o.FAC_POLICY_NO)) or
             upcase(strip(s.POLIS_CLN)) = upcase(strip(o.CLSDT_POLICY_NO))
         ));

    /* Pass 2: Slip Exact Match -> OSBAL */
    create table work.match_p2_slip as
    select s.sus_idx, o._row_idx_ as ref_idx, 'OSBAL' as ref_source, 'Slip only' as match_label
    from work.suspend_clean s
         inner join work.osbal_indexed o
         on (upcase(strip(s.SLIP_NO_CLN)) ne '' and (
             upcase(strip(s.SLIP_NO_CLN)) = upcase(strip(o.SLIP_CLEAN_1)) or
             upcase(strip(s.SLIP_NO_CLN)) = upcase(strip(o.SLIP_CLEAN_2)) or
             upcase(strip(s.SLIP_NO_CLN)) = upcase(strip(o.SLIP_CLEAN_3)) or
             upcase(strip(s.SLIP_NO_CLN)) = upcase(strip(o.FAC_SLIP)) or
             upcase(strip(s.SLIP_NO_CLN)) = upcase(strip(o.CLSDT_SLIP_NO))
         ));

    /* Pass 3: Insured Exact Match -> OSBAL */
    create table work.match_p3_insured as
    select s.sus_idx, o._row_idx_ as ref_idx, 'OSBAL' as ref_source, 'Insured only' as match_label
    from work.suspend_clean s
         inner join work.osbal_indexed o
         on (upcase(strip(s.INSURED_CLEAN)) ne '' and (
             upcase(strip(s.INSURED_CLEAN)) = upcase(strip(o.INSURED_CLEAN_1)) or
             upcase(strip(s.INSURED_CLEAN)) = upcase(strip(o.INSURED_CLEAN_2)) or
             upcase(strip(s.INSURED_CLEAN)) = upcase(strip(o.FAC_INSURED))
         ));

    /* Pass 4: Polis Match -> FACUL -> Resolve FAC_CODE -> OSBAL */
    create table work.facul_hit_polis as
    select s.sus_idx, upcase(strip(f.FAC_CODE)) as fac_code, 'Polis only' as match_label
    from work.suspend_clean s
         inner join work.facul_indexed f
         on (upcase(strip(s.POLIS_CLN)) ne '' and (
             upcase(strip(s.POLIS_CLN)) = upcase(strip(f.POLICY_CLEAN_1)) or
             upcase(strip(s.POLIS_CLN)) = upcase(strip(f.POLICY_CLEAN_2)) or
             upcase(strip(s.POLIS_CLN)) = upcase(strip(f.FAC_POLICY_NO))
         ));

    create table work.match_p4_facul_polis as
    select p.sus_idx, o._row_idx_ as ref_idx, 'FACUL' as ref_source, p.match_label
    from work.facul_hit_polis p
         inner join work.osbal_indexed o
         on p.fac_code = upcase(strip(o.CCOS_REF_CODE));

    /* Pass 5: Slip Match -> FACUL -> Resolve FAC_CODE -> OSBAL */
    create table work.facul_hit_slip as
    select s.sus_idx, upcase(strip(f.FAC_CODE)) as fac_code, 'Slip only' as match_label
    from work.suspend_clean s
         inner join work.facul_indexed f
         on (upcase(strip(s.SLIP_NO_CLN)) ne '' and (
             upcase(strip(s.SLIP_NO_CLN)) = upcase(strip(f.SLIP_CLEAN_1)) or
             upcase(strip(s.SLIP_NO_CLN)) = upcase(strip(f.FAC_SLIP))
         ));

    create table work.match_p5_facul_slip as
    select p.sus_idx, o._row_idx_ as ref_idx, 'FACUL' as ref_source, p.match_label
    from work.facul_hit_slip p
         inner join work.osbal_indexed o
         on p.fac_code = upcase(strip(o.CCOS_REF_CODE));

    /* Pass 6: Insured Match -> FACUL -> Resolve FAC_CODE -> OSBAL */
    create table work.facul_hit_insured as
    select s.sus_idx, upcase(strip(f.FAC_CODE)) as fac_code, 'Insured only' as match_label
    from work.suspend_clean s
         inner join work.facul_indexed f
         on (upcase(strip(s.INSURED_CLEAN)) ne '' and (
             upcase(strip(s.INSURED_CLEAN)) = upcase(strip(f.INSURED_CLEAN_1)) or
             upcase(strip(s.INSURED_CLEAN)) = upcase(strip(f.FAC_INSURED))
         ));

    create table work.match_p6_facul_insured as
    select p.sus_idx, o._row_idx_ as ref_idx, 'FACUL' as ref_source, p.match_label
    from work.facul_hit_insured p
         inner join work.osbal_indexed o
         on p.fac_code = upcase(strip(o.CCOS_REF_CODE));
quit;

/* Gabungkan semua pass dengan hierarki prioritas */
data work.all_initial_matches;
    length sus_idx 8 ref_idx 8 ref_source $10 match_label $50 pass_priority 8;
    set work.match_p1_polis         (in=p1)
        work.match_p4_facul_polis   (in=p4)
        work.match_p2_slip          (in=p2)
        work.match_p5_facul_slip    (in=p5)
        work.match_p3_insured       (in=p3)
        work.match_p6_facul_insured (in=p6);

    if p1 then pass_priority = 10;
    else if p4 then pass_priority = 15;
    else if p2 then pass_priority = 20;
    else if p5 then pass_priority = 25;
    else if p3 then pass_priority = 30;
    else if p6 then pass_priority = 35;
run;

/* Pilih pass_priority terbaik untuk setiap sus_idx */
proc sort data=work.all_initial_matches; by sus_idx pass_priority ref_idx; run;

data work.best_initial_matches;
    set work.all_initial_matches;
    by sus_idx pass_priority;
    if first.sus_idx then _best_prio_ = pass_priority;
    retain _best_prio_;
    if pass_priority = _best_prio_;
    drop _best_prio_;
run;


/* =============================================================================
   SECTION 4: NARROWING (CURRENCY, PERIODE, SLIP, SERTIFIKAT, INSURED)
   ============================================================================= */
%put [4/6] Currency & Periode Validation and Narrowing ...;

proc sql noprint;
    create table work.matches_with_attrs as
    select m.*,
           upcase(strip(o.CCOS_CURR)) as osbal_curr,
           upcase(strip(s."CURR ORI"n)) as sus_curr,
           o._com_date_,
           s._sus_date_,
           s._sus_date_ls_,
           s._is_lineslip_,
           o.CCOS_REF_CODE,
           s.POLIS_CLN,
           s.SLIP_NO_CLN,
           s.SERTIF_CLN_1,
           s.INSURED_CLEAN,
           o.SLIP_CLEAN_1,
           o.SERTIF_CLEAN_1 as osb_sertif_1,
           o.CLSDT_SERTF_NO,
           o.INSURED_CLEAN_1 as osb_ins_1
    from work.best_initial_matches m
         inner join work.osbal_indexed o on m.ref_idx = o._row_idx_
         inner join work.suspend_clean s on m.sus_idx = s.sus_idx;

    /* Hitung kecocokan currency */
    create table work.curr_counts as
    select sus_idx,
           count(*) as n_cand,
           sum(case when osbal_curr = sus_curr and osbal_curr ne '' then 1 else 0 end) as n_curr_match
    from work.matches_with_attrs
    group by sus_idx;

    create table work.curr_narrowed as
    select m.*
    from work.matches_with_attrs m
         inner join work.curr_counts c on m.sus_idx = c.sus_idx
    where (c.n_curr_match > 0 and m.osbal_curr = m.sus_curr)
       or (c.n_curr_match = 0);
quit;

/* Periode Filter */
data work.periode_narrowed;
    set work.curr_narrowed;
    _per_ok_ = 1;
    if _com_date_ ne . and _sus_date_ ne . then do;
        if _is_lineslip_ = 1 then do;
            if year(_com_date_) ne year(_sus_date_ls_) or month(_com_date_) ne month(_sus_date_ls_) then _per_ok_ = 0;
        end;
        else do;
            if _sus_date_ < _com_date_ then _per_ok_ = 0;
        end;
    end;
run;

/* Narrowing konfirmasi Slip, Sertifikat, Insured */
data work.narrowed_confirmed;
    set work.periode_narrowed;
    length final_scenario $100;
    final_scenario = match_label;

    /* Konfirmasi Slip */
    if match_label = 'Polis only' and SLIP_NO_CLN ne '' and SLIP_CLEAN_1 ne '' then do;
        if SLIP_NO_CLN = SLIP_CLEAN_1 then final_scenario = 'Polis + Slip';
    end;

    /* Konfirmasi Sertifikat */
    _cert_match_ = 0;
    if SERTIF_CLN_1 ne '' then do;
        if osb_sertif_1 ne '' and SERTIF_CLN_1 = osb_sertif_1 then _cert_match_ = 1;
        else if CLSDT_SERTF_NO ne '' and SERTIF_CLN_1 = CLSDT_SERTF_NO then _cert_match_ = 1;
    end;
    if _cert_match_ then final_scenario = catx(' ', final_scenario, '+ Cert');

    /* Konfirmasi Insured */
    if INSURED_CLEAN ne '' and osb_ins_1 ne '' and index(osb_ins_1, strip(INSURED_CLEAN)) > 0 then do;
        if index(final_scenario, 'Insured') = 0 then final_scenario = catx(' ', final_scenario, '+ Insured');
    end;

    /* Beda Currency / Beda Periode tags */
    if osbal_curr ne sus_curr and osbal_curr ne '' and sus_curr ne '' then final_scenario = 'Beda Currency';
    else if not _per_ok_ then final_scenario = 'Beda Periode';
run;


/* =============================================================================
   SECTION 5: BORDERO OPEN COVER OVERRIDE (FIRE, MC, MH)
   ============================================================================= */
%put [5/6] Checking Bordero Open Cover matches ...;

proc sql noprint;
    /* Match Bordero ke Suspend */
    create table work.sus_bdr_matches as
    select distinct s.sus_idx, b.fac_code, o._row_idx_ as ref_idx
    from work.suspend_clean s
         inner join work.bordero_combined b
         on (upcase(strip(s.POLIS_CLN)) ne '' and (
             upcase(strip(s.POLIS_CLN)) = b.bdr_polis or
             index(b.bdr_polis, upcase(strip(s.POLIS_CLN))) > 0 or
             index(upcase(strip(s.POLIS_CLN)), b.bdr_polis) > 0
         ))
         inner join work.osbal_indexed o
         on b.fac_code = upcase(strip(o.CCOS_REF_CODE));
quit;

/* Gabungkan hasil Bordero jika ada */
%let n_bdr_hits = %sysfunc(attrn(%sysfunc(open(work.sus_bdr_matches)), nobs));
%if &n_bdr_hits. > 0 %then %do;
    proc sql noprint;
        update work.narrowed_confirmed m
        set final_scenario = 'Polis only (Bordero)'
        where m.sus_idx in (select distinct sus_idx from work.sus_bdr_matches)
          and m.final_scenario not in ('Beda Currency', 'Beda Periode');
    quit;
%end;

/* Tambahkan baris Suspend yang Unmatching */
proc sql noprint;
    create table work.unmatched_sus as
    select s.sus_idx,
           . as ref_idx,
           'NONE' as ref_source,
           'Unmatching' as match_label,
           999 as pass_priority,
           'Unmatching' as final_scenario
    from work.suspend_clean s
    where s.sus_idx not in (select distinct sus_idx from work.narrowed_confirmed);
quit;

data work.all_matches_combined;
    set work.narrowed_confirmed (keep=sus_idx ref_idx ref_source match_label final_scenario)
        work.unmatched_sus;
run;


/* =============================================================================
   SECTION 6: FAC CODE CLAIM, ACCUMULATION, DAN OUTPUT 38 KOLOM
   ============================================================================= */
%put [6/6] Computing FLAG_PROD and formatting 38 columns ...;

proc sql noprint;
    create table work.match_full_details as
    select m.sus_idx,
           m.ref_idx,
           m.ref_source,
           m.final_scenario,
           s."RECEIPT NO"n,
           s."CREDIT NOTES"n,
           s."DETAIL RINCIAN NO"n,
           s."RECEIPT DATE"n,
           s."CEDANT NAME"n,
           s."CEDANT SHRT NAME"n,
           s.INSURED as INSURED_ORI,
           s.INSURED_CLEAN as INSURED_1,
           '' as INSURED_2 length=500,
           s."CURR ORI"n,
           s."AMOUNT ORI"n,
           s."CURR PAY"n,
           s."AMOUNT PAY"n,
           s.POLIS,
           s.POLIS_CLN as POLICY_CLEAN_1,
           s.SERTIF_CLN_1,
           '' as POLICY_CLEAN_2 length=200,
           s.SERTIF_CLN_2,
           '' as POLICY_CLEAN_3 length=200,
           s.SERTIF_CLN_3,
           '' as POLICY_CLEAN_4 length=200,
           '' as POLICY_CLEAN_5 length=200,
           s.SLIP_NO,
           s.SLIP_NO_CLN,
           s."DESC 1"n,
           s."DESC 2"n,
           s."DESC 3"n,
           s."DESC 4"n,
           s.STATUS,
           s.REC_TYPE,
           o.CCOS_DOC_NO,
           o.CCOS_REF_CODE,
           o.CCOS_OR_BAL,
           o.CCOS_BAL_DUE
    from work.all_matches_combined m
         left join work.osbal_indexed o on m.ref_idx = o._row_idx_
         left join work.suspend_clean s on m.sus_idx = s.sus_idx;

    /* Hitung jumlah baris per sus_idx dan fac_code */
    create table work.fac_summary as
    select sus_idx,
           count(distinct CCOS_REF_CODE) as n_fac_codes,
           count(*) as n_osbal_rows
    from work.match_full_details
    where CCOS_REF_CODE ne ''
    group by sus_idx;
quit;

data work.final_calc;
    merge work.match_full_details (in=a)
          work.fac_summary (in=b);
    by sus_idx;
    if a;

    /* Format dan Parse Numerik */
    _amt_raw_ = strip(put("AMOUNT ORI"n, best20.));
    if index(_amt_raw_, ',') > 0 then do;
        _amt_raw_ = compress(_amt_raw_, '.', 'kd');
        _amt_raw_ = translate(_amt_raw_, '.', ',');
    end;
    _amt_num_ = input(_amt_raw_, ??best32.);
    if missing(_amt_num_) then _amt_num_ = 0;

    if _amt_num_ ne 0 then AMOUNT_ORI_MIN1 = _amt_num_ * -1;
    else AMOUNT_ORI_MIN1 = .;

    _amount_neg_ = coalesce(AMOUNT_ORI_MIN1, 0);
    _bal_due_    = coalesce(CCOS_BAL_DUE, 0);

    /* Jika > 1 fac code atau beda curr/per, kosongkan CCOS balance */
    if n_fac_codes > 1 or final_scenario in ('Beda Currency', 'Beda Periode', 'Unmatching') then do;
        CCOS_OR_BAL = .;
        CCOS_BAL_DUE = .;
        _bal_due_ = 0;
    end;

    DIFERENCE = _amount_neg_ - _bal_due_;

    /* Logic Flags */
    _is_accumulated_ = (n_osbal_rows > 1);
    _is_equal_       = (round(_amount_neg_, 0.01) = round(_bal_due_, 0.01));
    _is_new_tot_     = (_amt_num_ ne 0 and _bal_due_ = 0);
    _is_gt1_ocmc_    = (final_scenario = 'Polis only (> 1 Fac Bordero)' or (index(final_scenario, 'Bordero') > 0 and n_fac_codes > 1));
    _is_gt1_fac_     = (n_fac_codes > 1 and not _is_gt1_ocmc_);
    _is_adj_tot_     = _is_equal_;
    _is_adj_seb_     = (not _is_equal_ and _amount_neg_ < _bal_due_);
    _is_new_seb_     = (not _is_equal_ and _amount_neg_ > _bal_due_);

    length FLAG_PROD $50;
    select;
        when (final_scenario = 'Beda Currency')              FLAG_PROD = 'Beda Currency';
        when (final_scenario = 'Beda Periode')               FLAG_PROD = 'Beda Periode';
        when (_is_gt1_ocmc_)                                 FLAG_PROD = 'Matching >1 Fac code OC MC';
        when (_is_gt1_fac_)                                  FLAG_PROD = 'Matching >1 fac code';
        when (_is_new_tot_)                                  FLAG_PROD = 'New Entry total';
        when (_is_adj_tot_ and not _is_accumulated_)         FLAG_PROD = 'Adjustment total tanpa akumulasi';
        when (_is_adj_tot_ and _is_accumulated_)             FLAG_PROD = 'Adjustment total dengan akumulasi';
        when (_is_adj_seb_ and not _is_accumulated_)         FLAG_PROD = 'Adjustment sebagian tanpa akumulasi';
        when (_is_adj_seb_ and _is_accumulated_)             FLAG_PROD = 'Adjustment sebagian dengan akumulasi';
        when (_is_new_seb_)                                  FLAG_PROD = 'New Entry sebagian';
        otherwise                                            FLAG_PROD = 'Unmatching';
    end;

    if final_scenario in ('Unmatching', 'UNMATCHED', '') then do;
        final_scenario = 'Unmatching';
        FLAG_PROD = 'Unmatching';
        CCOS_DOC_NO = '';
        CCOS_REF_CODE = '';
    end;

    /* CCOS_DOC_NO hanya diisi untuk Adjustment total tanpa akumulasi */
    if FLAG_PROD ne 'Adjustment total tanpa akumulasi' then CCOS_DOC_NO = '';

    rename final_scenario = SKENARIO;
    drop _amt_raw_ _amt_num_ _amount_neg_ _bal_due_ _is_accumulated_ _is_equal_
         _is_new_tot_ _is_gt1_ocmc_ _is_gt1_fac_ _is_adj_tot_ _is_adj_seb_ _is_new_seb_;
run;

/* Susun Tepat 38 Kolom Sesuai Spesifikasi Astra Buana */
data work.final_output_astrabuana;
    retain CCOS_DOC_NO CCOS_REF_CODE
           "RECEIPT NO"n "CREDIT NOTES"n "DETAIL RINCIAN NO"n "RECEIPT DATE"n
           "CEDANT NAME"n "CEDANT SHRT NAME"n
           INSURED_ORI INSURED_1 INSURED_2
           "CURR ORI"n "AMOUNT ORI"n AMOUNT_ORI_MIN1
           "CURR PAY"n "AMOUNT PAY"n
           CCOS_OR_BAL CCOS_BAL_DUE DIFERENCE
           FLAG_PROD
           POLIS
           POLICY_CLEAN_1 SERTIF_CLEAN_1
           POLICY_CLEAN_2 SERTIF_CLEAN_2
           POLICY_CLEAN_3 SERTIF_CLEAN_3
           POLICY_CLEAN_4 POLICY_CLEAN_5
           SLIP_NO SLIP_NO_CLN
           "DESC 1"n "DESC 2"n "DESC 3"n "DESC 4"n
           STATUS REC_TYPE
           SKENARIO;
    set work.final_calc;
    keep CCOS_DOC_NO CCOS_REF_CODE
         "RECEIPT NO"n "CREDIT NOTES"n "DETAIL RINCIAN NO"n "RECEIPT DATE"n
         "CEDANT NAME"n "CEDANT SHRT NAME"n
         INSURED_ORI INSURED_1 INSURED_2
         "CURR ORI"n "AMOUNT ORI"n AMOUNT_ORI_MIN1
         "CURR PAY"n "AMOUNT PAY"n
         CCOS_OR_BAL CCOS_BAL_DUE DIFERENCE
         FLAG_PROD
         POLIS
         POLICY_CLEAN_1 SERTIF_CLEAN_1
         POLICY_CLEAN_2 SERTIF_CLEAN_2
         POLICY_CLEAN_3 SERTIF_CLEAN_3
         POLICY_CLEAN_4 POLICY_CLEAN_5
         SLIP_NO SLIP_NO_CLN
         "DESC 1"n "DESC 2"n "DESC 3"n "DESC 4"n
         STATUS REC_TYPE
         SKENARIO;
run;

/* Export ke Excel */
proc export
    data=work.final_output_astrabuana
    outfile="&OUTPUT_FILE."
    dbms=xlsx
    replace;
    sheet="Sheet1";
run;

%put =============================================================================;
%put [PROD MATCHING ASTRA] Selesai dieksekusi!;
%put Output file: &OUTPUT_FILE.;
%put =============================================================================;

/* Tampilkan Ringkasan Distribusi Hasil */
title "Distribusi FLAG_PROD - PT ASURANSI ASTRA BUANA";
proc freq data=work.final_output_astrabuana;
    tables FLAG_PROD / nocum;
run;

title "Distribusi SKENARIO - PT ASURANSI ASTRA BUANA";
proc freq data=work.final_output_astrabuana;
    tables SKENARIO / nocum;
run;
title;
