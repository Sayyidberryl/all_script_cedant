options validvarname=any validmemname=extend nomprint nosymbolgen compress=yes;

/* --- 1. Parameter & Koneksi DB --- */
%let DB_HOST          = 10.10.125.91;
%let DB_PORT          = 5432;
%let DB_NAME          = fsi_db;
%let DB_USER          = viy4user;
%let DB_PASS          = fild42op61nx;

%let CAS_TARGET_LIB   = Public;
%let DB_TABLE_BORDERO = ACA_OPEN_COVER_MC;
%let DB_TABLE_SUSPEND = SUSPENSE_ORI_2026;

/* Inisialisasi Sesi CAS (Identik Pola Developer Lama) */
cas mycas sessopts=(caslib="casuser" timeout=3600);
libname casuser cas sessref=mycas caslib="casuser";
libname public  cas sessref=mycas caslib="Public";
libname fsi_db  cas sessref=mycas caslib="fsi_db";


/* --- 2. Load Clean Tables (CAS Staging / PostgreSQL FSI_DB) --- */

%macro load_clean_data();
    /* OSBAL Clean */
    %if %sysfunc(exist(casuser.aca_osbal_clean)) %then %do;
        data work.osbal_clean; set casuser.aca_osbal_clean; run;
    %end;
    %else %if %sysfunc(exist(PG_FSI."SUSPENSE_DATA_OSBAL_CLEAN"n)) %then %do;
        proc delete data=casuser.aca_osbal_clean; run;
        proc fedsql sessref=mycas;
            create table casuser.aca_osbal_clean as
            select * from FSI_DB."SUSPENSE_DATA_OSBAL_CLEAN"
            where "CCOS_COMP_NAME" like 'PT. ASURANSI CENTRAL ASIA%';
        quit;
        data work.osbal_clean; set casuser.aca_osbal_clean; run;
    %end;

    /* Facul Clean */
    %if %sysfunc(exist(casuser.aca_facul_clean)) %then %do;
        data work.facul_clean; set casuser.aca_facul_clean; run;
    %end;
    %else %if %sysfunc(exist(PG_FSI."SUSPENSE_DATA_FACUL_CLEAN"n)) %then %do;
        proc delete data=casuser.aca_facul_clean; run;
        proc fedsql sessref=mycas;
            create table casuser.aca_facul_clean as
            select * from FSI_DB."SUSPENSE_DATA_FACUL_CLEAN"
            where "COMP_NAME" like 'PT. ASURANSI CENTRAL ASIA%';
        quit;
        data work.facul_clean; set casuser.aca_facul_clean; run;
    %end;

    /* Suspend Clean */
    %if %sysfunc(exist(casuser.aca_suspend_clean)) %then %do;
        data work.suspend_clean; set casuser.aca_suspend_clean; run;
    %end;
    %else %if %sysfunc(exist(Public.ACA_SUSPEND_CLEAN)) %then %do;
        data work.suspend_clean; set Public.ACA_SUSPEND_CLEAN; run;
    %end;
    %else %do;
        proc delete data=casuser.aca_suspend_raw; run;
        proc fedsql sessref=mycas;
            create table casuser.aca_suspend_raw as 
            select distinct * 
            from FSI_DB.&DB_TABLE_SUSPEND.
            where "CEDANT SHRT NAME" like 'CENTRAL%' and "STATUS" like 'SUSPENSE%';
        quit;

        data work.suspend_filtered;
            set casuser.aca_suspend_raw;
            rename INSURED   = insured_ori
                   POLIS     = polis_ori
                   "SLIP NO"n = slip_ori;
        run;

        data work.suspend_clean;
            set work.suspend_filtered;

            length clean_polis_1-clean_polis_5 $500
                   clean_slip_1-clean_slip_5 $500
                   clean_insured_1-clean_insured_5 $500
                   clean_sertif_1-clean_sertif_5 $10;

            clean_polis_1 = strip(polis_ori);
            clean_polis_1 = prxchange('s/-(?:\d+(?:\/\d+)?|EXT\(\d+\))$//oi', -1, strip(clean_polis_1));
            clean_sertif_1 = '';
            retain _rx_sfx_ 0;
            if _rx_sfx_ = 0 then _rx_sfx_ = prxparse('/[-\s]+\s*(\d{1,6})(?:[^0-9]|$)/o');
            if prxmatch(_rx_sfx_, strip(polis_ori)) then do;
                _craw_ = prxposn(_rx_sfx_, 1, strip(polis_ori));
                if prxmatch('/^\d{1,6}$/', strip(_craw_)) then clean_sertif_1 = put(input(strip(_craw_), ??8.), z6.);
            end;

            clean_slip_1 = strip(slip_ori);
            clean_insured_1 = upcase(strip(insured_ori));

            drop _craw_ _rx_sfx_;
        run;
        data casuser.aca_suspend_clean; set work.suspend_clean; run;
    %end;

    /* Bordero Raw */
    %if %sysfunc(exist(FSI_DB."&DB_TABLE_BORDERO."n)) or %sysfunc(exist(FSI_DB.&DB_TABLE_BORDERO.)) %then %do;
        proc delete data=casuser.aca_bordero_raw; run;
        proc fedsql sessref=mycas;
            create table casuser.aca_bordero_raw as
            select * from FSI_DB.&DB_TABLE_BORDERO.;
        quit;
        data work.bordero_raw; set casuser.aca_bordero_raw; run;
    %end;
    %else %do;
        data work.bordero_raw; stop; run;
    %end;
%mend load_clean_data;
%load_clean_data();

/* --- 3. Mesin Rekonsiliasi (Multi-Pass Matching) --- */
data work.slipdb_clean; stop; run;

%macro load_bordero_source();
    %if %sysfunc(libref(FSI_DB)) = 0 %then %do;
        %if %sysfunc(exist(FSI_DB.&DB_TABLE_BORDERO.)) or %sysfunc(exist(FSI_DB."&DB_TABLE_BORDERO."n)) %then %do;
            %load_db_table(aca_bordero_raw, work.bordero_raw, &DB_TABLE_BORDERO.);
            %put [BORDERO] Loaded from database FSI_DB.&DB_TABLE_BORDERO.: %sysfunc(attrn(%sysfunc(open(work.bordero_raw)), nobs)) rows.;
        %end;
        %else %do;
            data work.bordero_raw; stop; run;
            %put [BORDERO WARNING] Table FSI_DB.&DB_TABLE_BORDERO. not found.;
        %end;
    %end;
    %else %do;
        data work.bordero_raw; stop; run;
        %put [BORDERO WARNING] FSI_DB library not connected.;
    %end;
%mend load_bordero_source;
%load_bordero_source();

data work.osbal_clean;
    set work.osbal_clean;
    format _com_date_ date9.;
    _com_date_ = input(put(FAC_COM_DATE, $100.), ??anydtdte10.);
run;

data work.suspend_clean;
    set work.suspend_clean;
    sus_idx = _n_;
    format _sus_date_ date9. _sus_date_ls_ date9.;
    _sus_date_ = input(put("RECEIPT DATE"n, $100.), ??anydtdte10.);

    _is_lineslip_ = 0;
    if prxmatch('/LINESLIP|LINE\s*SLIP/oi', strip(clean_insured_1))
    or prxmatch('/LINESLIP|LINE\s*SLIP/oi', strip(insured_ori)) then
        _is_lineslip_ = 1;

    if _is_lineslip_ and _sus_date_ ne . then
        _sus_date_ls_ = intnx('month', _sus_date_, -1, 'same');
    else
        _sus_date_ls_ = .;
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

data work.slipdb_indexed;
    set work.slipdb_clean;
    _is_excl_ = 0;
    if prxmatch('/HUTANG\s*PIUTANG|DATA\s*SUSPENSE/oi', strip(polis_ori))
    then _is_excl_ = 1;
    if not _is_excl_;
    _row_idx_ = _n_;
    drop _is_excl_;
run;

proc sql noprint;
    create table work.pass1_polis as
    select s.sus_idx, o._row_idx_ as ref_idx, 'OSBAL' as ref_source, 'Polis only' as match_label
    from work.suspend_clean s
         inner join work.osbal_indexed o
         on upcase(strip(s.clean_polis_1)) = upcase(strip(o.clean_polis_1))
         or upcase(strip(s.clean_polis_1)) = upcase(strip(o.clean_polis_2))
         or upcase(strip(s.clean_polis_1)) = upcase(strip(o.clean_polis_3))
         or upcase(strip(s.clean_polis_2)) = upcase(strip(o.clean_polis_1))
    where s.clean_polis_1 ne '';

    create table work.pass1_slip as
    select s.sus_idx, o._row_idx_ as ref_idx, 'OSBAL' as ref_source, 'Slip only' as match_label
    from work.suspend_clean s
         inner join work.osbal_indexed o
         on upcase(strip(s.clean_slip_1)) = upcase(strip(o.clean_slip_1))
         or upcase(strip(s.clean_slip_1)) = upcase(strip(o.clean_slip_2))
    where s.clean_slip_1 ne '';

    create table work.pass3_slipdb as
    select s.sus_idx, sl._row_idx_ as slipdb_idx,
           upcase(strip(sl.FAC_CODE)) as fac_code,
           'SLIPDB' as ref_source, 'Polis only' as match_label
    from work.suspend_clean s
         inner join work.slipdb_indexed sl
         on upcase(strip(s.clean_polis_1)) = upcase(strip(sl.clean_polis_1))
         or upcase(strip(s.clean_polis_1)) = upcase(strip(sl.clean_polis_2))
    where s.clean_polis_1 ne '' and sl.FAC_CODE ne '';

    create table work.pass3_resolved as
    select p.sus_idx, o._row_idx_ as ref_idx, 'SLIPDB' as ref_source, p.match_label
    from work.pass3_slipdb p
         inner join work.osbal_indexed o on p.fac_code = upcase(strip(o.CCOS_REF_CODE));

    create table work.pass5_facul as
    select s.sus_idx, f._row_idx_ as facul_idx,
           upcase(strip(f.FAC_CODE)) as fac_code,
           'FACUL' as ref_source, 'Polis only' as match_label
    from work.suspend_clean s
         inner join work.facul_indexed f
         on upcase(strip(s.clean_polis_1)) = upcase(strip(f.clean_polis_1))
         or upcase(strip(s.clean_polis_1)) = upcase(strip(f.clean_polis_2))
    where s.clean_polis_1 ne '' and f.FAC_CODE ne '';

    create table work.pass5_resolved as
    select p.sus_idx, o._row_idx_ as ref_idx, 'FACUL' as ref_source, p.match_label
    from work.pass5_facul p
         inner join work.osbal_indexed o on p.fac_code = upcase(strip(o.CCOS_REF_CODE));

    create table work.pass7_insured as
    select s.sus_idx, o._row_idx_ as ref_idx, 'OSBAL' as ref_source, 'Insured only' as match_label
    from work.suspend_clean s
         inner join work.osbal_indexed o
         on upcase(strip(s.clean_insured_1)) = upcase(strip(o.clean_insured_1))
         or upcase(strip(s.clean_insured_1)) = upcase(strip(o.clean_insured_2))
    where s.clean_insured_1 ne '';
quit;

data work.all_matches;
    length sus_idx 8 ref_idx 8 ref_source $10 match_label $50 pass_priority 8;
    set work.pass1_polis (in=p1p)
        work.pass1_slip (in=p1s)
        work.pass3_resolved (in=p3)
        work.pass5_resolved (in=p5)
        work.pass7_insured (in=p7);
    if p1p then pass_priority = 10;
    else if p1s then pass_priority = 20;
    else if p3 then pass_priority = 30;
    else if p5 then pass_priority = 40;
    else if p7 then pass_priority = 50;
run;

proc sort data=work.all_matches; by sus_idx pass_priority ref_idx; run;

data work.best_pass_per_sus;
    set work.all_matches;
    by sus_idx pass_priority;
    if first.sus_idx then _best_prio_ = pass_priority;
    retain _best_prio_;
    if pass_priority = _best_prio_;
    drop _best_prio_;
run;

proc sql noprint;
    create table work.matched_with_curr as
    select m.*,
           upcase(strip(o.CCOS_CURR)) as osbal_curr,
           upcase(strip(s."CURR ORI"n)) as sus_curr
    from work.best_pass_per_sus m
         left join work.osbal_indexed o on m.ref_idx = o._row_idx_
         left join work.suspend_clean s on m.sus_idx = s.sus_idx;

    create table work.curr_counts as
    select sus_idx,
           count(*) as n_cand,
           sum(case when osbal_curr = sus_curr and osbal_curr ne '' then 1 else 0 end) as n_curr_match
    from work.matched_with_curr
    group by sus_idx;

    create table work.curr_narrowed as
    select m.sus_idx, m.ref_idx, m.ref_source, m.match_label, m.pass_priority, m.osbal_curr, m.sus_curr
    from work.matched_with_curr m
         inner join work.curr_counts c on m.sus_idx = c.sus_idx
    where (c.n_curr_match > 0 and m.osbal_curr = m.sus_curr)
       or (c.n_curr_match = 0);
quit;

proc sql noprint;
    create table work.periode_narrowed as
    select m.*
    from work.curr_narrowed m
         left join work.osbal_indexed o on m.ref_idx = o._row_idx_
         left join work.suspend_clean s on m.sus_idx = s.sus_idx
    where
        (o._com_date_ = . or s._sus_date_ = .)
        or (s._is_lineslip_ = 1
            and year(o._com_date_) = year(s._sus_date_ls_)
            and month(o._com_date_) = month(s._sus_date_ls_))
        or (s._is_lineslip_ = 0 and s._sus_date_ >= o._com_date_);
quit;

proc sql noprint;
    create table work.cand_counts as
    select sus_idx, count(*) as n_cand_final
    from work.periode_narrowed
    group by sus_idx;

    create table work.narrowed_final as
    select m.*,
           c.n_cand_final,
           m.match_label as final_scenario length=100
    from work.periode_narrowed m
         inner join work.cand_counts c on m.sus_idx = c.sus_idx;
quit;

proc sql noprint;
    create table work.unmatched_sus as
    select s.sus_idx, . as ref_idx, 'NONE' as ref_source,
           'Unmatching' as match_label, 999 as pass_priority,
           'Unmatching' as final_scenario length=100
    from work.suspend_clean s
    where s.sus_idx not in (select distinct sus_idx from work.narrowed_final);
quit;

data work.all_results;
    set work.narrowed_final (keep=sus_idx ref_idx ref_source match_label final_scenario)
        work.unmatched_sus;
run;

proc sql noprint;
    create table work.currency_check as
    select m.sus_idx,
           count(distinct upcase(strip(o.CCOS_CURR))) as n_distinct_curr,
           max(case when upcase(strip(s."CURR ORI"n)) = upcase(strip(o.CCOS_CURR)) then 1 else 0 end) as has_curr_match
    from work.all_results m
         left join work.osbal_indexed o on m.ref_idx = o._row_idx_
         left join work.suspend_clean s on m.sus_idx = s.sus_idx
    where m.ref_source ne 'NONE'
    group by m.sus_idx;
quit;

data work.all_results;
    merge work.all_results (in=a)
          work.currency_check (in=b);
    by sus_idx;
    if a;
    if b and n_distinct_curr > 0 and not has_curr_match then
        final_scenario = 'Beda Currency';
    drop n_distinct_curr has_curr_match;
run;

%macro run_bordero();
    %local n_bdr;
    %let n_bdr = 0;
    %if %sysfunc(exist(work.bordero_raw)) %then %do;
        %let n_bdr = %sysfunc(attrn(%sysfunc(open(work.bordero_raw)), nobs));
    %end;

    %if &n_bdr. > 0 %then %do;
        data work.bordero_idx;
            set work.bordero_raw;
            length fac_code $50 bdr_polis $500 bdr_slip $500 bdr_cert $10 bdr_insured $500;
            fac_code   = upcase(strip("FAC CODE"n));
            bdr_polis  = upcase(strip(put(POLIS, $500.)));
            bdr_slip   = upcase(strip(put(SLIP, $500.)));
            bdr_cert   = upcase(strip(put(CERTIFICATE, $10.)));
            bdr_insured= upcase(strip(put(INSURED, $500.)));
            _bl_ = input(put(upcase(strip(put(BULAN, $20.))), $bulan_fmt.), ??8.);
            _th_ = input(strip(put(TAHUN, $20.)), ??8.);
            if _bl_ > 0 and _th_ > 0 then bdr_period = mdy(_bl_, 1, _th_);
            else bdr_period = .;
            format bdr_period date9.;
            if fac_code ne '';
            keep fac_code bdr_polis bdr_slip bdr_cert bdr_insured bdr_period NET;
        run;

        proc sql noprint;
            create table work.fac_counts as
            select m.sus_idx,
                   count(distinct upcase(strip(o.CCOS_REF_CODE))) as n_fac_codes
            from work.all_results m
                 left join work.osbal_indexed o on m.ref_idx = o._row_idx_
            where m.ref_source ne 'NONE'
              and m.final_scenario not in ('Beda Currency','Beda Periode')
            group by m.sus_idx;

            create table work.multi_fac_sus as
            select f.sus_idx,
                   upcase(strip(o.CCOS_REF_CODE)) as fac_code,
                   m.ref_idx,
                   upcase(strip(s.clean_polis_1)) as sus_polis,
                   upcase(strip(s.clean_slip_1))  as sus_slip,
                   upcase(strip(s.clean_sertif_1)) as sus_cert
            from work.fac_counts f
                 inner join work.all_results m on f.sus_idx = m.sus_idx
                 inner join work.osbal_indexed o on m.ref_idx = o._row_idx_
                 left join work.suspend_clean s on f.sus_idx = s.sus_idx
            where f.n_fac_codes > 1;

            create table work.bordero_matched as
            select distinct a.sus_idx, a.fac_code, a.ref_idx
            from work.multi_fac_sus a
                 inner join work.bordero_idx b on a.fac_code = b.fac_code
                 and (a.sus_polis = b.bdr_polis or a.sus_slip = b.bdr_slip or a.sus_cert = b.bdr_cert);

            update work.all_results m
            set final_scenario = 'Matching >1 Fac code OC MC'
            where m.sus_idx in (select distinct sus_idx from work.bordero_matched);
        quit;
    %end;
%mend run_bordero;
%run_bordero();

proc sql noprint;
    create table work.sus_with_fac as
    select m.sus_idx,
           m.ref_idx,
           m.ref_source,
           m.final_scenario,
           upcase(strip(o.CCOS_REF_CODE)) as fac_code,
           s."RECEIPT NO"n,
           s."CREDIT NOTES"n,
           s."DETAIL RINCIAN NO"n,
           s."RECEIPT DATE"n as receipt_date,
           s."CEDANT SHRT NAME"n as cedant_shrt_name,
           s."CURR ORI"n as curr_ori,
           s."AMOUNT ORI"n as amount_ori,
           s."CURR PAY"n as curr_pay,
           s."AMOUNT PAY"n as amount_pay,
           s.insured_ori,
           s.clean_insured_1,
           s.clean_insured_2,
           s.polis_ori,
           s.clean_polis_1,
           s.clean_sertif_1,
           s.clean_polis_2,
           s.clean_sertif_2,
           s.clean_polis_3,
           s.clean_sertif_3,
           s.clean_polis_4,
           s.clean_polis_5,
           s.slip_ori,
           s.clean_slip_1,
           s."DESC 1"n, s."DESC 2"n, s."DESC 3"n, s."DESC 4"n,
           s.STATUS, s.REC_TYPE,
           o.CCOS_DOC_NO,
           o.CCOS_REF_CODE,
           o.CCOS_OR_BAL,
           o.CCOS_BAL_DUE,
           o.CCOS_CURR as osbal_curr
    from work.all_results m
         left join work.osbal_indexed o on m.ref_idx = o._row_idx_
         left join work.suspend_clean s on m.sus_idx = s.sus_idx;
quit;

/* --- 5. DUAL OUTPUT GENERATION: SUSPENSE V1 & SUSPENSE V2 (38 KOLOM STANDAR) --- */

/* --- 5A. Generate Suspense V1 (Row-Level / KEEP_ROWS) --- */
data work.v1_prepared;
    set work.sus_with_fac;

    if vtype(amount_ori) = 'N' then do;
        AMOUNT_ORI_NUM = coalesce(amount_ori, 0);
    end;
    else do;
        _amt_str_ = strip(put(amount_ori, $50.));
        _amt_str_ = compress(_amt_str_, ' $');
        if countc(_amt_str_, ',') = 1 and countc(_amt_str_, '.') = 0 then
            _amt_str_ = translate(_amt_str_, '.', ',');
        else if countc(_amt_str_, ',') > 0 and countc(_amt_str_, '.') > 0 then
            _amt_str_ = compress(_amt_str_, ',');
        AMOUNT_ORI_NUM = input(_amt_str_, ??best32.);
        if missing(AMOUNT_ORI_NUM) then AMOUNT_ORI_NUM = 0;
    end;

    if not missing(AMOUNT_ORI_NUM) and AMOUNT_ORI_NUM ne 0 then
        AMOUNT_ORI_MIN1 = AMOUNT_ORI_NUM * -1;
    else AMOUNT_ORI_MIN1 = .;

    _amount_neg_ = coalesce(AMOUNT_ORI_MIN1, 0);
    _bal_due_    = coalesce(CCOS_BAL_DUE, 0);
    DIFERENCE    = _amount_neg_ - _bal_due_;

    _accumulated_     = 0;  /* V1 row asli dipertahankan tanpa akumulasi suspend */
    _is_equal_        = (round(_amount_neg_, 0.01) = round(_bal_due_, 0.01));
    _is_new_entry_tot = (AMOUNT_ORI_NUM ne 0 and _bal_due_ = 0);
    _is_gt1_ocmc_     = (final_scenario = 'Matching >1 Fac code OC MC');
    _is_gt1_fac_      = (index(strip(CCOS_REF_CODE), ',') > 0 or CCOS_REF_CODE = 'facode lebih dari 1') and not _is_gt1_ocmc_;
    _is_adj_tot_      = _is_equal_;
    _is_adj_seb_      = not _is_equal_ and (_amount_neg_ < _bal_due_);
    _is_new_seb_      = not _is_equal_ and (_amount_neg_ > _bal_due_);

    length FLAG_PROD $50;
    select;
        when (final_scenario = 'Beda Currency')              FLAG_PROD = 'Beda Currency';
        when (final_scenario = 'Beda Periode')               FLAG_PROD = 'Beda Periode';
        when (_is_gt1_ocmc_)                                 FLAG_PROD = 'Matching >1 Fac code OC MC';
        when (_is_gt1_fac_)                                  FLAG_PROD = 'Matching >1 fac code';
        when (_is_new_entry_tot)                             FLAG_PROD = 'New Entry total';
        when (_is_adj_tot_)                                  FLAG_PROD = 'Adjustment total tanpa akumulasi';
        when (_is_adj_seb_)                                  FLAG_PROD = 'Adjustment sebagian tanpa akumulasi';
        when (_is_new_seb_)                                  FLAG_PROD = 'New Entry sebagian';
        otherwise                                            FLAG_PROD = 'Unmatching';
    end;

    if final_scenario in ('Unmatching','UNMATCHED') then do;
        final_scenario = 'Unmatching';
        FLAG_PROD = 'Unmatching';
    end;

    if FLAG_PROD ne 'Adjustment total tanpa akumulasi' then CCOS_DOC_NO = '';

    rename final_scenario = SKENARIO;

    drop _amount_neg_ _bal_due_ _accumulated_ _is_equal_
         _is_new_entry_tot _is_gt1_ocmc_ _is_gt1_fac_ _is_adj_tot_
         _is_adj_seb_ _is_new_seb_;
run;

data work.aca_matching_v1;
    length CCOS_DOC_NO $100 CCOS_REF_CODE $100
           "RECEIPT NO"n $100 "CREDIT NOTES"n $100 "DETAIL RINCIAN NO"n $100
           "RECEIPT DATE"n 8
           "CEDANT NAME"n $100 "CEDANT SHRT NAME"n $50
           INSURED_ORI $500 INSURED_1 $500 INSURED_2 $500
           "CURR ORI"n $10 "AMOUNT ORI"n 8 AMOUNT_ORI_MIN1 8
           "CURR PAY"n $10 "AMOUNT PAY"n 8
           CCOS_OR_BAL 8 CCOS_BAL_DUE 8 DIFERENCE 8
           FLAG_PROD $50
           POLIS $500
           POLICY_CLEAN_1 $500 SERTIF_CLEAN_1 $10
           POLICY_CLEAN_2 $500 SERTIF_CLEAN_2 $10
           POLICY_CLEAN_3 $500 SERTIF_CLEAN_3 $10
           POLICY_CLEAN_4 $500 POLICY_CLEAN_5 $500
           SLIP_NO $500 SLIP_NO_CLN $500
           "DESC 1"n "DESC 2"n "DESC 3"n "DESC 4"n $500
           STATUS $50 REC_TYPE $50
           SKENARIO $100;
    format "RECEIPT DATE"n yymmdd10.;

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

    set work.v1_prepared;

    "CEDANT NAME"n      = 'PT. ASURANSI CENTRAL ASIA';
    "CEDANT SHRT NAME"n = coalesce(strip(cedant_shrt_name), 'CENTRAL');
    "RECEIPT DATE"n     = receipt_date;
    "CURR ORI"n         = curr_ori;
    "AMOUNT ORI"n       = AMOUNT_ORI_NUM;
    "CURR PAY"n         = curr_pay;
    "AMOUNT PAY"n       = amount_pay;

    INSURED_ORI         = strip(insured_ori);
    INSURED_1           = strip(clean_insured_1);
    INSURED_2           = strip(clean_insured_2);

    POLIS               = strip(polis_ori);
    POLICY_CLEAN_1      = strip(clean_polis_1);
    SERTIF_CLEAN_1      = strip(clean_sertif_1);
    POLICY_CLEAN_2      = strip(clean_polis_2);
    SERTIF_CLEAN_2      = strip(clean_sertif_2);
    POLICY_CLEAN_3      = strip(clean_polis_3);
    SERTIF_CLEAN_3      = strip(clean_sertif_3);
    POLICY_CLEAN_4      = strip(clean_polis_4);
    POLICY_CLEAN_5      = strip(clean_polis_5);

    SLIP_NO             = strip(slip_ori);
    SLIP_NO_CLN         = strip(clean_slip_1);
run;
%put [Phase 5A] Suspense V1 generated: %sysfunc(attrn(%sysfunc(open(work.aca_matching_v1)), nobs)) rows.;

/* --- 5B. Generate Suspense V2 (Aggregated / MERGE_ROWS per FAC CODE) --- */
proc sql noprint;
    /* Grup A: Baris yang cocok dengan 1 fac code valid di-merge menjadi 1 baris */
    create table work.v2_group_a as
    select fac_code,
           max(ref_source) as ref_source,
           max(final_scenario) as final_scenario,
           sum(amount_ori) as amount_ori,
           max(curr_ori) as curr_ori,
           max(curr_pay) as curr_pay,
           sum(amount_pay) as amount_pay,
           max(CCOS_DOC_NO) as CCOS_DOC_NO,
           max(CCOS_REF_CODE) as CCOS_REF_CODE,
           sum(CCOS_OR_BAL) as CCOS_OR_BAL,
           sum(CCOS_BAL_DUE) as CCOS_BAL_DUE,
           max(osbal_curr) as osbal_curr,
           min(sus_idx) as sus_idx,
           count(*) as n_sus_rows
    from work.sus_with_fac
    where ref_source ne '' 
      and fac_code ne '' 
      and index(fac_code, ',') = 0 
      and final_scenario not in ('Unmatching', 'Beda Currency', 'Beda Periode', 'UNMATCHED')
    group by fac_code;
quit;

proc sql noprint;
    create table work.v2_merged_a as
    select a.fac_code,
           a.ref_source,
           a.final_scenario,
           a.amount_ori,
           a.curr_ori,
           a.curr_pay,
           a.amount_pay,
           a.CCOS_DOC_NO,
           a.CCOS_REF_CODE,
           a.CCOS_OR_BAL,
           a.CCOS_BAL_DUE,
           a.osbal_curr,
           a.n_sus_rows,
           s."RECEIPT NO"n,
           s."CREDIT NOTES"n,
           s."DETAIL RINCIAN NO"n,
           s."RECEIPT DATE"n as receipt_date,
           s."CEDANT SHRT NAME"n as cedant_shrt_name,
           s.insured_ori,
           s.clean_insured_1,
           s.clean_insured_2,
           s.polis_ori,
           s.clean_polis_1,
           s.clean_sertif_1,
           s.clean_polis_2,
           s.clean_sertif_2,
           s.clean_polis_3,
           s.clean_sertif_3,
           s.clean_polis_4,
           s.clean_polis_5,
           s.slip_ori,
           s.clean_slip_1,
           s."DESC 1"n, s."DESC 2"n, s."DESC 3"n, s."DESC 4"n,
           s.STATUS, s.REC_TYPE
    from work.v2_group_a a
         left join work.suspend_clean s on a.sus_idx = s.sus_idx;
quit;

/* Grup B: Baris unmatching / multi-fac / beda currency / beda periode tetap 1-to-1 */
data work.v2_merged_b;
    set work.sus_with_fac;
    where ref_source = '' 
       or fac_code = '' 
       or index(fac_code, ',') > 0 
       or final_scenario in ('Unmatching', 'Beda Currency', 'Beda Periode', 'UNMATCHED');
    n_sus_rows = 1;
run;

data work.v2_accumulated;
    set work.v2_merged_a work.v2_merged_b;
run;

data work.v2_prepared;
    set work.v2_accumulated;

    if vtype(amount_ori) = 'N' then do;
        AMOUNT_ORI_NUM = coalesce(amount_ori, 0);
    end;
    else do;
        _amt_str_ = strip(put(amount_ori, $50.));
        _amt_str_ = compress(_amt_str_, ' $');
        if countc(_amt_str_, ',') = 1 and countc(_amt_str_, '.') = 0 then
            _amt_str_ = translate(_amt_str_, '.', ',');
        else if countc(_amt_str_, ',') > 0 and countc(_amt_str_, '.') > 0 then
            _amt_str_ = compress(_amt_str_, ',');
        AMOUNT_ORI_NUM = input(_amt_str_, ??best32.);
        if missing(AMOUNT_ORI_NUM) then AMOUNT_ORI_NUM = 0;
    end;

    if not missing(AMOUNT_ORI_NUM) and AMOUNT_ORI_NUM ne 0 then
        AMOUNT_ORI_MIN1 = AMOUNT_ORI_NUM * -1;
    else AMOUNT_ORI_MIN1 = .;

    _amount_neg_ = coalesce(AMOUNT_ORI_MIN1, 0);
    _bal_due_    = coalesce(CCOS_BAL_DUE, 0);
    DIFERENCE    = _amount_neg_ - _bal_due_;

    _accumulated_     = (n_sus_rows > 1);
    _is_equal_        = (round(_amount_neg_, 0.01) = round(_bal_due_, 0.01));
    _is_new_entry_tot = (AMOUNT_ORI_NUM ne 0 and _bal_due_ = 0);
    _is_gt1_ocmc_     = (final_scenario = 'Matching >1 Fac code OC MC');
    _is_gt1_fac_      = (index(strip(CCOS_REF_CODE), ',') > 0 or CCOS_REF_CODE = 'facode lebih dari 1') and not _is_gt1_ocmc_;
    _is_adj_tot_      = _is_equal_;
    _is_adj_seb_      = not _is_equal_ and (_amount_neg_ < _bal_due_);
    _is_new_seb_      = not _is_equal_ and (_amount_neg_ > _bal_due_);

    length FLAG_PROD $50;
    select;
        when (final_scenario = 'Beda Currency')              FLAG_PROD = 'Beda Currency';
        when (final_scenario = 'Beda Periode')               FLAG_PROD = 'Beda Periode';
        when (_is_gt1_ocmc_)                                 FLAG_PROD = 'Matching >1 Fac code OC MC';
        when (_is_gt1_fac_)                                  FLAG_PROD = 'Matching >1 fac code';
        when (_is_new_entry_tot)                             FLAG_PROD = 'New Entry total';
        when (_is_adj_tot_ and not _accumulated_)            FLAG_PROD = 'Adjustment total tanpa akumulasi';
        when (_is_adj_tot_ and _accumulated_)                FLAG_PROD = 'Adjustment total dengan akumulasi';
        when (_is_adj_seb_ and not _accumulated_)            FLAG_PROD = 'Adjustment sebagian tanpa akumulasi';
        when (_is_adj_seb_ and _accumulated_)                FLAG_PROD = 'Adjustment sebagian dengan akumulasi';
        when (_is_new_seb_)                                  FLAG_PROD = 'New Entry sebagian';
        otherwise                                            FLAG_PROD = 'Unmatching';
    end;

    if final_scenario in ('Unmatching','UNMATCHED') then do;
        final_scenario = 'Unmatching';
        FLAG_PROD = 'Unmatching';
    end;

    if FLAG_PROD ne 'Adjustment total tanpa akumulasi' then CCOS_DOC_NO = '';

    rename final_scenario = SKENARIO;

    drop _amount_neg_ _bal_due_ _accumulated_ _is_equal_
         _is_new_entry_tot _is_gt1_ocmc_ _is_gt1_fac_ _is_adj_tot_
         _is_adj_seb_ _is_new_seb_;
run;

data work.aca_matching_v2;
    length CCOS_DOC_NO $100 CCOS_REF_CODE $100
           "RECEIPT NO"n $100 "CREDIT NOTES"n $100 "DETAIL RINCIAN NO"n $100
           "RECEIPT DATE"n 8
           "CEDANT NAME"n $100 "CEDANT SHRT NAME"n $50
           INSURED_ORI $500 INSURED_1 $500 INSURED_2 $500
           "CURR ORI"n $10 "AMOUNT ORI"n 8 AMOUNT_ORI_MIN1 8
           "CURR PAY"n $10 "AMOUNT PAY"n 8
           CCOS_OR_BAL 8 CCOS_BAL_DUE 8 DIFERENCE 8
           FLAG_PROD $50
           POLIS $500
           POLICY_CLEAN_1 $500 SERTIF_CLEAN_1 $10
           POLICY_CLEAN_2 $500 SERTIF_CLEAN_2 $10
           POLICY_CLEAN_3 $500 SERTIF_CLEAN_3 $10
           POLICY_CLEAN_4 $500 POLICY_CLEAN_5 $500
           SLIP_NO $500 SLIP_NO_CLN $500
           "DESC 1"n "DESC 2"n "DESC 3"n "DESC 4"n $500
           STATUS $50 REC_TYPE $50
           SKENARIO $100;
    format "RECEIPT DATE"n yymmdd10.;

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

    set work.v2_prepared;

    "CEDANT NAME"n      = 'PT. ASURANSI CENTRAL ASIA';
    "CEDANT SHRT NAME"n = coalesce(strip(cedant_shrt_name), 'CENTRAL');
    "RECEIPT DATE"n     = receipt_date;
    "CURR ORI"n         = curr_ori;
    "AMOUNT ORI"n       = AMOUNT_ORI_NUM;
    "CURR PAY"n         = curr_pay;
    "AMOUNT PAY"n       = amount_pay;

    INSURED_ORI         = strip(insured_ori);
    INSURED_1           = strip(clean_insured_1);
    INSURED_2           = strip(clean_insured_2);

    POLIS               = strip(polis_ori);
    POLICY_CLEAN_1      = strip(clean_polis_1);
    SERTIF_CLEAN_1      = strip(clean_sertif_1);
    POLICY_CLEAN_2      = strip(clean_polis_2);
    SERTIF_CLEAN_2      = strip(clean_sertif_2);
    POLICY_CLEAN_3      = strip(clean_polis_3);
    SERTIF_CLEAN_3      = strip(clean_sertif_3);
    POLICY_CLEAN_4      = strip(clean_polis_4);
    POLICY_CLEAN_5      = strip(clean_polis_5);

    SLIP_NO             = strip(slip_ori);
    SLIP_NO_CLN         = strip(clean_slip_1);
run;
%put [Phase 5B] Suspense V2 generated: %sysfunc(attrn(%sysfunc(open(work.aca_matching_v2)), nobs)) rows.;

/* --- 5. Persistensi ke CAS & PostgreSQL FSI_DB --- */

%macro sync_to_cas();
    %if %sysfunc(libref(casuser)) = 0 %then %do;
        proc delete data=casuser.ACA_MATCHING_V1; run;
        proc delete data=casuser.ACA_MATCHING_V2; run;

        data casuser.ACA_MATCHING_V1; set work.aca_matching_v1; run;
        data casuser.ACA_MATCHING_V2; set work.aca_matching_v2; run;

        proc cas;
            table.dropTable / caslib="&CAS_TARGET_LIB." name="ACA_MATCHING_V1" quiet=true;
            table.promote   / caslib="casuser" name="ACA_MATCHING_V1" targetLib="&CAS_TARGET_LIB.";
            table.save      / caslib="&CAS_TARGET_LIB." table="ACA_MATCHING_V1" name="ACA_MATCHING_V1.sashdat" replace=true;

            table.dropTable / caslib="&CAS_TARGET_LIB." name="ACA_MATCHING_V2" quiet=true;
            table.promote   / caslib="casuser" name="ACA_MATCHING_V2" targetLib="&CAS_TARGET_LIB.";
            table.save      / caslib="&CAS_TARGET_LIB." table="ACA_MATCHING_V2" name="ACA_MATCHING_V2.sashdat" replace=true;
        quit;

        /* Pastikan tabel juga tetap standby di casuser (untuk dashboard pribadi) */
        data casuser.ACA_MATCHING_V1; set work.aca_matching_v1; run;
        data casuser.ACA_MATCHING_V2; set work.aca_matching_v2; run;
        data casuser.SUSPENSE_DATA_SUSPENSE_V1; set work.aca_matching_v1; run;
        data casuser.SUSPENSE_DATA_SUSPENSE_V2; set work.aca_matching_v2; run;
    %end;
%mend sync_to_cas;
%sync_to_cas();

%macro sync_to_postgresql();
    %local fsidb_ok;
    %let fsidb_ok = 0;
    proc sql noprint;
        select count(*) into :fsidb_ok trimmed
        from dictionary.libnames
        where 1=1;
    quit;

    %if &fsidb_ok. > 0 %then %do;
        /* Hapus data lama ACA saja (idempoten & aman untuk multi-cedant) */
        proc sql;
            delete from FSI_DB."SUSPENSE_DATA_SUSPENSE_V1"n
            where "CEDANT NAME"n like 'PT. ASURANSI CENTRAL ASIA%' or "CEDANT SHRT NAME"n like 'CENTRAL%';

            delete from FSI_DB."SUSPENSE_CLEAN_2026"n
            where "CEDANT NAME"n like 'PT. ASURANSI CENTRAL ASIA%' or "CEDANT SHRT NAME"n like 'CENTRAL%';

            delete from FSI_DB."SUSPENSE_DATA_SUSPENSE_V2"n
            where "CEDANT NAME"n like 'PT. ASURANSI CENTRAL ASIA%' or "CEDANT SHRT NAME"n like 'CENTRAL%';
        quit;

        /* Append Suspense V1 & V2 */
        proc append base=FSI_DB."SUSPENSE_DATA_SUSPENSE_V1"n data=work.aca_matching_v1 force; run;
        proc append base=FSI_DB."SUSPENSE_CLEAN_2026"n data=work.aca_matching_v1 force; run;
        proc append base=FSI_DB."SUSPENSE_DATA_SUSPENSE_V2"n data=work.aca_matching_v2 force; run;
    %end;
%mend sync_to_postgresql;
%sync_to_postgresql();

/* --- 6. Ringkasan Eksekutif (Identik Python & Results Tab) --- */

%macro print_python_summary();
    %local dsid1 n_v1 c_v1 dsid2 n_v2 c_v2;
    %let dsid1 = %sysfunc(open(work.aca_matching_v1));
    %let n_v1  = %sysfunc(attrn(&dsid1., nobs));
    %let c_v1  = %sysfunc(attrn(&dsid1., nvars));
    %let rc    = %sysfunc(close(&dsid1.));

    %let dsid2 = %sysfunc(open(work.aca_matching_v2));
    %let n_v2  = %sysfunc(attrn(&dsid2., nobs));
    %let c_v2  = %sysfunc(attrn(&dsid2., nvars));
    %let rc    = %sysfunc(close(&dsid2.));

    proc freq data=work.aca_matching_v1 noprint;
        tables FLAG_PROD / out=work._sum_v1_flag(rename=(COUNT=_cnt_));
        tables SKENARIO  / out=work._sum_v1_ske(rename=(COUNT=_cnt_));
    run;
    proc freq data=work.aca_matching_v2 noprint;
        tables FLAG_PROD / out=work._sum_v2_flag(rename=(COUNT=_cnt_));
        tables SKENARIO  / out=work._sum_v2_ske(rename=(COUNT=_cnt_));
    run;

    data _null_;
        put / "============================================================";
        put "  RINGKASAN EKSEKUSI REKONSILIASI SUSPENSE ACA (OUTPUT V1)";
        put "============================================================";
        put "  Rows  : " &n_v1. comma12. "  |  Cols: &c_v1.  |  Mode: Row-Level (KEEP_ROWS)";
        put ;
        put "  FLAG_PROD:";
    run;
    data _null_;
        set work._sum_v1_flag;
        put "    " FLAG_PROD $40. " : " _cnt_ comma10.;
    run;
    data _null_;
        put ;
        put "  SKENARIO:";
    run;
    data _null_;
        set work._sum_v1_ske;
        put "    " SKENARIO $45. " : " _cnt_ comma10.;
    run;
    data _null_;
        put "============================================================";
        put ;
        put "============================================================";
        put "  RINGKASAN EKSEKUSI REKONSILIASI SUSPENSE ACA (OUTPUT V2)";
        put "============================================================";
        put "  Rows  : " &n_v2. comma12. "  |  Cols: &c_v2.  |  Mode: Aggregated (MERGE_ROWS)";
        put ;
        put "  FLAG_PROD:";
    run;
    data _null_;
        set work._sum_v2_flag;
        put "    " FLAG_PROD $40. " : " _cnt_ comma10.;
    run;
    data _null_;
        put ;
        put "  SKENARIO:";
    run;
    data _null_;
        set work._sum_v2_ske;
        put "    " SKENARIO $45. " : " _cnt_ comma10.;
    run;
    data _null_;
        put "============================================================";
        put "  Semua output (CAS in-memory & PostgreSQL FSI_DB) terupdate!";
        put "============================================================";
        put ;
    run;

    /* Tampilkan juga di tab Results SAS Studio */
    title1 "ACA Suspense Reconciliation - Output V1 Summary (Row-Level / KEEP_ROWS)";
    proc freq data=work.aca_matching_v1;
        tables FLAG_PROD * SKENARIO / nocum nopercent norow nocol;
    run;

    title1 "ACA Suspense Reconciliation - Output V2 Summary (Aggregated / MERGE_ROWS)";
    proc freq data=work.aca_matching_v2;
        tables FLAG_PROD * SKENARIO / nocum nopercent norow nocol;
    run;
    title1;
%mend print_python_summary;
%print_python_summary();

%put [COMPLETE] aca_matching.sas execution completed successfully!;
