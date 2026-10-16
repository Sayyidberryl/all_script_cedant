/* =============================================================================
   PROGRAM   : ASTRA_PIPELINE_VIYA.SAS
   CEDANT    : PT ASURANSI ASTRA BUANA
   TARGET    : SAS Viya / SAS 9.4 Studio
   DESKRIPSI : Script Master All-In-One (Langsung Copy-Paste)
               Menjalankan seluruh pipeline data:
               Phase 1: Data Cleaning (Facul, OSBAL, Suspend)
               Phase 2: Suspend Matching & Flagging (V1: keep rows, V2: merge rows)
               Phase 3: Export Excel & Summary Statistics
   CARA PAKAI:
     1. Buka SAS Studio di SAS Viya.
     2. Buat Program baru, paste seluruh isi file ini.
     3. Sesuaikan macro variable BASE_DIR di bawah jika diperlukan.
     4. Klik tombol "Run" (atau tekan F3).
   ============================================================================= */

options validvarname=any validmemname=extend mprint symbolgen compress=yes;

/* ═════════════════════════════════════════════════════════════════════════════
   0. KONFIGURASI PARAMETER & PATH
   ═════════════════════════════════════════════════════════════════════════════ */
%let BASE_DIR     = %str(c:\Users\berryl\Desktop\starcore\indore\all_script_cedant\astra_production_script);
%let DATA_DIR     = %str(&BASE_DIR.\data);

/* Parameter Output Matching: 1 (row asli dipertahankan) atau 2 (merge per fac) */
%let VERSION         = 1;
%let ACCUMULATE_MODE = KEEP_ROWS; /* KEEP_ROWS untuk V1 | MERGE_ROWS untuk V2 */

/* Path File Input Mentah */
%let FACUL_INPUT   = %str(&DATA_DIR.\facul.xlsx);
%let OSBAL_INPUT   = %str(&DATA_DIR.\osbal.xlsx);
%let SUSPEND_INPUT = %str(&DATA_DIR.\suspend.xlsx);

/* Deteksi alternatif raw path jika file di data/ belum diekstrak */
%macro detect_raw_inputs();
    %if not %sysfunc(fileexist("&FACUL_INPUT.")) %then %do;
        %let ALT_F = %str(&BASE_DIR.\dataExcel\raw\1b. Transaksi Facul 01.01.23 - 17.08.2026.xlsx);
        %if %sysfunc(fileexist("&ALT_F.")) %then %let FACUL_INPUT = &ALT_F.;
    %end;
    %if not %sysfunc(fileexist("&OSBAL_INPUT.")) %then %do;
        %let ALT_O = %str(&BASE_DIR.\dataExcel\raw\2b. transaksi Osbal 01.01.23 - 17.08.26.xlsx);
        %if %sysfunc(fileexist("&ALT_O.")) %then %let OSBAL_INPUT = &ALT_O.;
    %end;
    %if not %sysfunc(fileexist("&SUSPEND_INPUT.")) %then %do;
        %let ALT_S = %str(&BASE_DIR.\dataExcel\raw\3b. Database Suspense 150826 (2).xlsx);
        %if %sysfunc(fileexist("&ALT_S.")) %then %let SUSPEND_INPUT = &ALT_S.;
    %end;
%mend detect_raw_inputs;
%detect_raw_inputs();

/* Path Folder Bordero */
%let BORDERO_DIR   = %str(&DATA_DIR.\BORDERO);
%let BDR_FIRE_FILE = %str(&BORDERO_DIR.\ASTRA_OPEN_COVER_FIRE.xlsx);
%let BDR_MC_FILE   = %str(&BORDERO_DIR.\ASTRA_OPEN_COVER_MC.xlsx);
%let BDR_MH_FILE   = %str(&BORDERO_DIR.\ASTRA_OPEN_COVER_MH.xlsx);

/* Path File Output Clean */
%let FACUL_CLEAN_OUT   = %str(&DATA_DIR.\astrabuana_output_facul.xlsx);
%let OSBAL_CLEAN_OUT   = %str(&DATA_DIR.\astrabuana_output_osbal.xlsx);
%let SUSPEND_CLEAN_OUT = %str(&DATA_DIR.\astrabuana_output_suspend.xlsx);
%let FINAL_OUTPUT_OUT  = %str(&DATA_DIR.\final_output_astrabuana_v&VERSION..xlsx);

%put =========================================================================;
%put   ASTRA BUANA RECONCILIATION PIPELINE — SAS VIYA ALL-IN-ONE MASTER;
%put   BASE_DIR : &BASE_DIR.;
%put   VERSION  : &VERSION. (&ACCUMULATE_MODE.);
%put =========================================================================;


/* ═════════════════════════════════════════════════════════════════════════════
   1. MACRO LIBRARY & FORMAT DEFINITIONS
   ═════════════════════════════════════════════════════════════════════════════ */

/* Macro helper cek keberadaan variabel di dataset */
%macro varexist(ds, var);
    %local dsid rc varnum;
    %let dsid = %sysfunc(open(&ds.));
    %if &dsid. %then %do;
        %let varnum = %sysfunc(varnum(&dsid., &var.));
        %let rc = %sysfunc(close(&dsid.));
        &varnum.
    %end;
    %else 0;
%mend varexist;

/* Format mapping bulan Indonesia ke nomor bulan */
proc format;
    value $bulan_fmt
        'JANUARI'   = '1'  'JANUARY'   = '1'  'JAN' = '1'
        'FEBRUARI'  = '2'  'FEBRUARY'  = '2'  'FEB' = '2'
        'MARET'     = '3'  'MARCH'     = '3'  'MAR' = '3'
        'APRIL'     = '4'  'APR'       = '4'
        'MEI'       = '5'  'MAY'       = '5'
        'JUNI'      = '6'  'JUNE'      = '6'  'JUN' = '6'
        'JULI'      = '7'  'JULY'      = '7'  'JUL' = '7'
        'AGUSTUS'   = '8'  'AUGUST'    = '8'  'AGU' = '8' 'AGS' = '8'
        'SEPTEMBER' = '9'  'SEP'       = '9'
        'OKTOBER'   = '10' 'OCTOBER'   = '10' 'OKT' = '10' 'OCT' = '10'
        'NOVEMBER'  = '11' 'NOV'       = '11'
        'DESEMBER'  = '12' 'DECEMBER'  = '12' 'DES' = '12' 'DEC' = '12'
        other       = '0';
run;

/* Utility Macros */
%macro normalize_spaces(val);
    strip(prxchange('s/\s{2,}/ /', -1, strip(&val)))
%mend normalize_spaces;

%macro clean_control_chars(var);
    &var = prxchange('s/[\x00-\x08\x0B\x0C\x0E-\x1F]//o', -1, &var);
%mend clean_control_chars;

%macro standarisasi_sd(var);
    if index(upcase(&var.), 'S/D') > 0 or index(upcase(&var.), 'S / D') > 0 then
        &var = prxchange('s/\s*S\s*\/\s*D\s*/ SD /oi', -1, &var);
%mend standarisasi_sd;

/* Clean Polis Khusus Astra Buana */
%macro apply_clean_polis_astra(in_var);
    array _pol_out{5} $200 POLICY_CLEAN_1-POLICY_CLEAN_5;
    do _pi_ = 1 to 5; _pol_out{_pi_} = ''; end;
    _pwrk_ = strip(&in_var.);
    _pn_ = 0;

    if missing(_pwrk_) or _pwrk_ = '' then goto _polis_astra_done_;

    %standarisasi_sd(_pwrk_);

    if prxmatch('/AS\s*PER\s*LIST|ASTRA\s*EX\s*LS|\bTBA\b|\bVARIOUS\b/oi', _pwrk_) then do;
        _pn_ = 1; POLICY_CLEAN_1 = %normalize_spaces(_pwrk_);
        goto _polis_astra_done_;
    end;

    _pwrk_ = prxchange('s/\(\s*LANJUTAN\b[^\)]*\)//oi', -1, _pwrk_);
    _pwrk_ = %normalize_spaces(_pwrk_);

    _np_parts_ = countw(_pwrk_, '/,', 'm');
    if _np_parts_ > 1 then do;
        do _si_ = 1 to _np_parts_;
            _tok_ = strip(scan(_pwrk_, _si_, '/,', 'm'));
            if prxmatch('/\d{11,}/', _tok_) and _pn_ < 5 then do;
                _pn_ + 1;
                _pol_out{_pn_} = _tok_;
            end;
        end;
    end;

    if _pn_ = 0 then do;
        if prxmatch('/\d{11,}/', _pwrk_) or lengthn(_pwrk_) >= 6 then do;
            _pn_ = 1; POLICY_CLEAN_1 = _pwrk_;
        end;
    end;

    _polis_astra_done_: ;
    drop _pwrk_ _pn_ _pi_ _si_ _np_parts_ _tok_;
%mend apply_clean_polis_astra;

/* Clean Slip Khusus Astra Buana */
%macro apply_clean_slip_astra(in_var);
    array _slp_out{5} $200 SLIP_CLEAN_1-SLIP_CLEAN_5;
    do _sli_ = 1 to 5; _slp_out{_sli_} = ''; end;
    _swrk_ = strip(&in_var.);
    _sln_ = 0;

    if missing(_swrk_) or _swrk_ = '' then goto _slip_astra_done_;

    %standarisasi_sd(_swrk_);

    _swrk_ = prxchange('s/^(?:JANUARI|FEBRUARI|MARET|APRIL|MEI|JUNI|JULI|AGUSTUS|SEPTEMBER|OKTOBER|NOVEMBER|DESEMBER|\d{4}|IDR|USD|JPY|EUR|SGD|GBP|\s*[-\/])+//oi', -1, _swrk_);
    _swrk_ = prxchange('s/\b(?:BORDEREAUX|BORDEROUX|BORDERO|BORD|SUMMARY|ASTRA|SISA|OSBAL)\b/ /oi', -1, _swrk_);
    _swrk_ = %normalize_spaces(_swrk_);

    _rx_stok_ = prxparse('/[A-Z0-9][A-Z0-9\-]{5,}/oi');
    _s_start_ = 1; _s_stop_ = 0; _s_pos_ = 1;
    call prxnext(_rx_stok_, _s_start_, length(_swrk_), _swrk_, _s_pos_, _s_len_);
    do while (_s_pos_ > 0 and _sln_ < 5);
        _cand_s_ = substr(_swrk_, _s_pos_, _s_len_);
        _cand_s_ = prxchange('s/^-|-$//o', -1, _cand_s_);
        if lengthn(_cand_s_) >= 6 and not prxmatch('/^\d{4}$/', _cand_s_) then do;
            _sln_ + 1;
            _slp_out{_sln_} = strip(_cand_s_);
        end;
        call prxnext(_rx_stok_, _s_start_, length(_swrk_), _swrk_, _s_pos_, _s_len_);
    end;
    call prxfree(_rx_stok_);

    if _sln_ = 0 and _swrk_ ne '' then do;
        _sln_ = 1; SLIP_CLEAN_1 = _swrk_;
    end;

    _slip_astra_done_: ;
    drop _swrk_ _sln_ _sli_ _cand_s_ _s_start_ _s_stop_ _s_pos_ _s_len_;
%mend apply_clean_slip_astra;

/* Clean Insured Khusus Astra Buana */
%macro apply_clean_insured_astra(in_var);
    array _ins_out{5} $500 INSURED_CLEAN_1-INSURED_CLEAN_5;
    do _ii_ = 1 to 5; _ins_out{_ii_} = ''; end;
    _iwrk_ = strip(&in_var.);
    _inn_ = 0;

    if missing(_iwrk_) or _iwrk_ = '' then goto _insured_astra_done_;

    _iwrk_ = prxchange('s/^(?:(?:AUTO\s+FACILITY|CARGO\s+FACILITY|LINESLIP)\s*\/\s*)+//oi', -1, _iwrk_);
    _iwrk_ = prxchange('s/\b(?:PT|CV|TBK|PTE|LTD|PELAYARAN|\(PERSERO\)|\bPERSERO\b)\b\.?/ /oi', -1, _iwrk_);
    _iwrk_ = prxchange('s/\s*\bAND\s*\/\s*OR\b\s*/|/oi', -1, _iwrk_);
    _iwrk_ = prxchange('s/\bQ\.?Q\.?\b/|/oi', -1, _iwrk_);
    _iwrk_ = prxchange('s/\bC\s*\/\s*Q\b/|/oi', -1, _iwrk_);
    _iwrk_ = prxchange('s/[,\/]/|/o', -1, _iwrk_);

    _n_in_parts_ = countw(_iwrk_, '|', 'm');
    do _p_ = 1 to _n_in_parts_;
        _part_ = strip(scan(_iwrk_, _p_, '|', 'm'));
        _part_ = compress(_part_, '.,');
        _part_ = %normalize_spaces(_part_);
        _part_ = upcase(_part_);
        if lengthn(_part_) > 2 and _inn_ < 5 then do;
            _inn_ + 1;
            _ins_out{_inn_} = _part_;
        end;
    end;

    if _inn_ = 0 and lengthn(_iwrk_) > 2 then do;
        _inn_ = 1; INSURED_CLEAN_1 = upcase(%normalize_spaces(compress(_iwrk_, '.,')));
    end;

    _insured_astra_done_: ;
    drop _iwrk_ _inn_ _ii_ _p_ _n_in_parts_ _part_;
%mend apply_clean_insured_astra;


/* ═════════════════════════════════════════════════════════════════════════════
   PHASE 1: DATA CLEANING (FACUL, OSBAL, SUSPEND)
   ═════════════════════════════════════════════════════════════════════════════ */

/* ── 1A. CLEANING FACUL ────────────────────────────────────────────────────── */
%put [PHASE 1A] Cleaning Facul Data ...;

proc import datafile="&FACUL_INPUT." out=work.facul_raw dbms=xlsx replace; getnames=yes; datarow=2; run;

data work.facul_cleaned;
    set work.facul_raw;
    where COMP_NAME = 'PT ASURANSI ASTRA BUANA';

    length INSURED_CLEAN_1 - INSURED_CLEAN_5 $500
           POLICY_CLEAN_1  - POLICY_CLEAN_5  $200
           SERTIF_CLEAN_1  - SERTIF_CLEAN_3  $50
           SLIP_CLEAN_1    - SLIP_CLEAN_5    $200;

    %apply_clean_insured_astra(FAC_INSURED);
    %apply_clean_polis_astra(FAC_POLICY_NO);
    SERTIF_CLEAN_1 = ''; SERTIF_CLEAN_2 = ''; SERTIF_CLEAN_3 = '';
    %apply_clean_slip_astra(FAC_SLIP);

    array _str_cols{*} _character_;
    do _sci_ = 1 to dim(_str_cols);
        if not missing(_str_cols{_sci_}) then %clean_control_chars(_str_cols{_sci_});
    end;
    drop _sci_;
run;

data work.facul_output;
    retain FAC_CODE FAC_CEDANT COMP_NAME FAC_BROKER COMP_NAME2 BUSSINESS_PATNER
           FAC_INSURED INSURED_CLEAN_1 INSURED_CLEAN_2 INSURED_CLEAN_3 INSURED_CLEAN_4 INSURED_CLEAN_5
           FAC_POLICY_NO POLICY_CLEAN_1 SERTIF_CLEAN_1 POLICY_CLEAN_2 SERTIF_CLEAN_2 POLICY_CLEAN_3 SERTIF_CLEAN_3
           POLICY_CLEAN_4 POLICY_CLEAN_5
           FAC_SLIP SLIP_CLEAN_1 SLIP_CLEAN_2 SLIP_CLEAN_3 SLIP_CLEAN_4 SLIP_CLEAN_5
           FAC_CURRENCY FAC_INP_DATE FAC_COM_DATE FAC_EXP_DATE FAC_SUB_CLASS FAC_RISK FAC_DESC FAC_ACC_STS FAC_STS_SLIP;
    set work.facul_cleaned;
run;

proc export data=work.facul_output outfile="&FACUL_CLEAN_OUT." dbms=xlsx replace; sheet="Query result"; run;


/* ── 1B. CLEANING OSBAL ────────────────────────────────────────────────────── */
%put [PHASE 1B] Cleaning OSBAL Data ...;

proc import datafile="&OSBAL_INPUT." out=work.osbal_raw dbms=xlsx replace; getnames=yes; datarow=2; run;

%let HAS_CLSDT_POLIS = %eval(%varexist(work.osbal_raw, CLSDT_POLICY_NO) > 0);
%let HAS_CLSDT_SLIP  = %eval(%varexist(work.osbal_raw, CLSDT_SLIP_NO)  > 0);
%let HAS_CLSDT_SERTF = %eval(%varexist(work.osbal_raw, CLSDT_SERTF_NO) > 0);

%macro is_clsdt_valid_astra(val, result_var);
    &result_var = 0;
    if not missing(&val.) and strip(&val.) ne '' then do;
        if not prxmatch('/\b(?:SISA|OSBAL|TBA|VARIOUS|VAR)\b/oi', strip(&val.)) then do;
            if prxmatch('/\d{6,}/', strip(&val.)) or lengthn(strip(&val.)) >= 6 then &result_var = 1;
        end;
    end;
%mend is_clsdt_valid_astra;

data work.osbal_cleaned;
    set work.osbal_raw;
    where CCOS_COMP_NAME = 'PT ASURANSI ASTRA BUANA';

    length INSURED_CLEAN_1 - INSURED_CLEAN_5 $500
           POLICY_CLEAN_1  - POLICY_CLEAN_5  $200
           SERTIF_CLEAN_1  - SERTIF_CLEAN_3  $50
           SLIP_CLEAN_1    - SLIP_CLEAN_5    $200;

    %apply_clean_insured_astra(FAC_INSURED);

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

    %if &HAS_CLSDT_SERTF. %then %do;
        _sertf_raw_ = strip(CLSDT_SERTF_NO);
        if prxmatch('/\.0$/', _sertf_raw_) then _sertf_raw_ = substr(_sertf_raw_, 1, lengthn(_sertf_raw_) - 2);
        SERTIF_CLEAN_1 = '';
        if not missing(_sertf_raw_) and _sertf_raw_ notin ('', '#', '-') then do;
            if prxmatch('/^\d{1,6}$/', strip(_sertf_raw_)) then
                SERTIF_CLEAN_1 = put(input(strip(_sertf_raw_), ??8.), z6.);
            else SERTIF_CLEAN_1 = strip(_sertf_raw_);
        end;
        drop _sertf_raw_;
    %end;
    %else %do;
        SERTIF_CLEAN_1 = '';
    %end;
    SERTIF_CLEAN_2 = ''; SERTIF_CLEAN_3 = '';

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

    array _all_str_{*} _character_;
    do _si_ = 1 to dim(_all_str_);
        if not missing(_all_str_{_si_}) then %clean_control_chars(_all_str_{_si_});
    end;
    drop _si_;
run;

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

proc export data=work.osbal_output outfile="&OSBAL_CLEAN_OUT." dbms=xlsx replace; sheet="Query result"; run;


/* ── 1C. CLEANING SUSPEND ──────────────────────────────────────────────────── */
%put [PHASE 1C] Cleaning Suspend Data ...;

proc import datafile="&SUSPEND_INPUT." out=work.suspend_raw dbms=xlsx replace; getnames=yes; datarow=4; run;

data work.suspend_cleaned;
    set work.suspend_raw;
    where upcase(strip("CEDANT SHRT NAME"n)) = 'ASTRABUANA' and upcase(strip(STATUS)) = 'SUSPENSE';
    STATUS = 'SUSPENSE';

    length INSURED_CLEAN $500 POLIS_CLN $200 SERTIF_CLN_1 - SERTIF_CLN_3 $100 SLIP_NO_CLN $200 SLIP_NO $200;

    _ins_txt_ = prxchange('s/\b(?:PT|CV|TBK|PTE|LTD)\b//oi', -1, strip(INSURED));
    _ins_txt_ = compress(_ins_txt_, '.,');
    INSURED_CLEAN = upcase(strip(prxchange('s/\s+/ /', -1, _ins_txt_)));

    _pol_txt_ = prxchange('s/-EXT\(\d+\)//oi', -1, strip(POLIS));
    if index(upcase(_pol_txt_), 'S/D') > 0 or index(upcase(_pol_txt_), 'S / D') > 0 then
        _pol_txt_ = prxchange('s/\s*S\s*\/\s*D\s*/SD/oi', -1, _pol_txt_);
    _pol_txt_ = compress(_pol_txt_, '.,');
    POLIS_CLN = strip(prxchange('s/\s+/ /', -1, _pol_txt_));

    _slip_orig_ = strip("SLIP NO"n);
    SLIP_NO     = _slip_orig_;

    _cert_txt_ = _slip_orig_;
    if index(upcase(_cert_txt_), 'S/D') > 0 or index(upcase(_cert_txt_), 'S / D') > 0 then
        _cert_txt_ = prxchange('s/\s*S\s*\/\s*D\s*/SD/oi', -1, _cert_txt_);

    SERTIF_CLN_1 = ''; SERTIF_CLN_2 = ''; SERTIF_CLN_3 = '';

    _rx_rng_ = prxparse('/-(\d{1,4})\s*(?:-|SD)\s*(\d{1,4})\s*$/oi');
    if prxmatch(_rx_rng_, _cert_txt_) then do;
        call prxposn(_rx_rng_, 1, _r_s1_, _r_l1_);
        call prxposn(_rx_rng_, 2, _r_s2_, _r_l2_);
        _raw1_ = substr(_cert_txt_, _r_s1_, _r_l1_);
        _raw2_ = substr(_cert_txt_, _r_s2_, _r_l2_);

        if not (lengthn(strip(_raw2_)) <= 2 and put(input(strip(_raw2_), ??8.), z2.) = '00') then do;
            _n1_ = input(_raw1_, 8.); _n2_ = input(_raw2_, 8.);
            if _n2_ < _n1_ then do; _tmp_n_ = _n1_; _n1_ = _n2_; _n2_ = _tmp_n_; end;
            _cnt_ = _n2_ - _n1_ + 1;
            _p1_ = put(_n1_, z6.); _p2_ = put(_n2_, z6.);
            if _cnt_ > 3 then SERTIF_CLN_1 = catx(' ', _p1_, 'SD', _p2_);
            else do;
                SERTIF_CLN_1 = put(_n1_, z6.);
                do _curr_n_ = _n1_ + 1 to _n2_;
                    SERTIF_CLN_1 = catx(', ', SERTIF_CLN_1, put(_curr_n_, z6.));
                end;
            end;
        end;
    end;
    else do;
        _rx_sgl_ = prxparse('/-(\d{2})\s*$/o');
        if prxmatch(_rx_sgl_, _cert_txt_) then do;
            call prxposn(_rx_sgl_, 1, _s_s1_, _s_l1_);
            _raw_s_ = substr(_cert_txt_, _s_s1_, _s_l1_);
            if strip(_raw_s_) ne '00' then SERTIF_CLN_1 = put(input(strip(_raw_s_), 8.), z6.);
        end;
        call prxfree(_rx_sgl_);
    end;
    call prxfree(_rx_rng_);

    _slp_txt_ = prxchange('s/\s*-\s*New\s*$//oi', -1, _slip_orig_);
    if index(upcase(_slp_txt_), 'S/D') > 0 or index(upcase(_slp_txt_), 'S / D') > 0 then
        _slp_txt_ = prxchange('s/\s*S\s*\/\s*D\s*/SD/oi', -1, _slp_txt_);
    _slp_txt_ = compress(_slp_txt_, '.,');
    SLIP_NO_CLN = strip(prxchange('s/\s+/ /', -1, _slp_txt_));

    array _str_cols{*} _character_;
    do _sci_ = 1 to dim(_str_cols);
        if not missing(_str_cols{_sci_}) then %clean_control_chars(_str_cols{_sci_});
    end;
    drop _sci_ _ins_txt_ _pol_txt_ _slip_orig_ _cert_txt_ _slp_txt_
         _r_s1_ _r_l1_ _r_s2_ _r_l2_ _raw1_ _raw2_ _n1_ _n2_ _tmp_n_ _cnt_ _p1_ _p2_ _curr_n_
         _s_s1_ _s_l1_ _raw_s_;
run;

proc export data=work.suspend_cleaned outfile="&SUSPEND_CLEAN_OUT." dbms=xlsx replace; sheet="Sheet1"; run;


/* ═════════════════════════════════════════════════════════════════════════════
   PHASE 2: RECONCILIATION & MATCHING ENGINE
   ═════════════════════════════════════════════════════════════════════════════ */
%put [PHASE 2] Matching Engine ...;

/* Load Bordero (FIRE, MC, MH) jika tersedia */
%macro load_bordero_all();
    %if %sysfunc(fileexist("&BDR_FIRE_FILE.")) %then %do;
        proc import datafile="&BDR_FIRE_FILE." out=work.bdr_fire dbms=xlsx replace; getnames=yes; datarow=2; run;
    %end;
    %else %do; data work.bdr_fire; stop; run; %end;

    %if %sysfunc(fileexist("&BDR_MC_FILE.")) %then %do;
        proc import datafile="&BDR_MC_FILE." out=work.bdr_mc dbms=xlsx replace; getnames=yes; datarow=2; run;
    %end;
    %else %do; data work.bdr_mc; stop; run; %end;

    %if %sysfunc(fileexist("&BDR_MH_FILE.")) %then %do;
        proc import datafile="&BDR_MH_FILE." out=work.bdr_mh dbms=xlsx replace; getnames=yes; datarow=2; run;
    %end;
    %else %do; data work.bdr_mh; stop; run; %end;
%mend load_bordero_all;
%load_bordero_all();

data work.bordero_combined;
    set work.bdr_fire (in=f) work.bdr_mc (in=mc) work.bdr_mh (in=mh);
    length bdr_src $10 fac_code $50 bdr_polis $200 bdr_slip $200 bdr_cert $50 bdr_insured $500 bdr_curr $10;
    if f then bdr_src = 'FIRE'; else if mc then bdr_src = 'MC'; else if mh then bdr_src = 'MH';
    fac_code    = upcase(strip("FAC CODE"n));
    bdr_polis   = upcase(strip(POLIS));
    bdr_slip    = upcase(strip(SLIP));
    bdr_cert    = upcase(strip(CERTIFICATE));
    bdr_insured = upcase(strip(INSURED));
    bdr_curr    = upcase(strip(CURR));

    _b_no_ = input(put(upcase(strip(BULAN)), $bulan_fmt.), 8.);
    _t_no_ = input(strip(put(TAHUN, 8.)), ??8.);
    if _b_no_ > 0 and _t_no_ > 0 then bdr_period = mdy(_b_no_, 1, _t_no_);
    else bdr_period = .;
    format bdr_period date9.;
    if fac_code ne '';
    drop _b_no_ _t_no_;
run;

/* Pre-parse Suspend */
data work.suspend_prep;
    set work.suspend_cleaned;
    sus_idx = _n_;

    format _sus_date_ date9. _sus_date_ls_ date9.;
    _sus_date_ = input(put("RECEIPT DATE"n, $100.), anydtdte10.);

    _is_lineslip_ = 0;
    if prxmatch('/LINESLIP|LINE\s*SLIP/oi', strip(INSURED)) or
       prxmatch('/LINESLIP|LINE\s*SLIP/oi', strip(INSURED_CLEAN)) then
        _is_lineslip_ = 1;

    if _is_lineslip_ and _sus_date_ ne . then
        _sus_date_ls_ = intnx('month', _sus_date_, -1, 'same');
    else _sus_date_ls_ = .;
run;

/* Pre-parse OSBAL */
data work.osbal_prep;
    set work.osbal_cleaned;
    _row_idx_ = _n_;

    _is_excl_ = 0;
    if prxmatch('/HUTANG\s*PIUTANG|DATA\s*SUSPENSE/oi', strip(FAC_POLICY_NO)) or
       prxmatch('/HUTANG\s*PIUTANG|DATA\s*SUSPENSE/oi', strip(POLICY_CLEAN_1)) or
       prxmatch('/HUTANG\s*PIUTANG|DATA\s*SUSPENSE/oi', strip(FAC_SLIP)) or
       prxmatch('/HUTANG\s*PIUTANG|DATA\s*SUSPENSE/oi', strip(SLIP_CLEAN_1)) then
        _is_excl_ = 1;
    if not _is_excl_;

    format _com_date_ date9.;
    _com_date_ = input(put(FAC_COM_DATE, $100.), anydtdte10.);
    drop _is_excl_;
run;

/* Pre-parse FACUL */
data work.facul_prep;
    set work.facul_cleaned;
    _row_idx_ = _n_;
    _is_excl_ = 0;
    if prxmatch('/HUTANG\s*PIUTANG|DATA\s*SUSPENSE/oi', strip(FAC_POLICY_NO)) or
       prxmatch('/HUTANG\s*PIUTANG|DATA\s*SUSPENSE/oi', strip(POLICY_CLEAN_1)) then
        _is_excl_ = 1;
    if not _is_excl_;
    drop _is_excl_;
run;

/* Matching Passes */
proc sql noprint;
    /* P1: Polis Exact -> OSBAL */
    create table work.match_p1_polis as
    select s.sus_idx, o._row_idx_ as ref_idx, 'OSBAL' as ref_source, 'Polis only' as match_label
    from work.suspend_prep s
         inner join work.osbal_prep o
         on (upcase(strip(s.POLIS_CLN)) ne '' and (
             upcase(strip(s.POLIS_CLN)) = upcase(strip(o.POLICY_CLEAN_1)) or
             upcase(strip(s.POLIS_CLN)) = upcase(strip(o.POLICY_CLEAN_2)) or
             upcase(strip(s.POLIS_CLN)) = upcase(strip(o.POLICY_CLEAN_3)) or
             upcase(strip(s.POLIS_CLN)) = upcase(strip(o.POLICY_CLEAN_4)) or
             upcase(strip(s.POLIS_CLN)) = upcase(strip(o.POLICY_CLEAN_5)) or
             upcase(strip(s.POLIS_CLN)) = upcase(strip(o.FAC_POLICY_NO)) or
             upcase(strip(s.POLIS_CLN)) = upcase(strip(o.CLSDT_POLICY_NO))
         ));

    /* P2: Slip Exact -> OSBAL */
    create table work.match_p2_slip as
    select s.sus_idx, o._row_idx_ as ref_idx, 'OSBAL' as ref_source, 'Slip only' as match_label
    from work.suspend_prep s
         inner join work.osbal_prep o
         on (upcase(strip(s.SLIP_NO_CLN)) ne '' and (
             upcase(strip(s.SLIP_NO_CLN)) = upcase(strip(o.SLIP_CLEAN_1)) or
             upcase(strip(s.SLIP_NO_CLN)) = upcase(strip(o.SLIP_CLEAN_2)) or
             upcase(strip(s.SLIP_NO_CLN)) = upcase(strip(o.SLIP_CLEAN_3)) or
             upcase(strip(s.SLIP_NO_CLN)) = upcase(strip(o.FAC_SLIP)) or
             upcase(strip(s.SLIP_NO_CLN)) = upcase(strip(o.CLSDT_SLIP_NO))
         ));

    /* P3: Insured Exact -> OSBAL */
    create table work.match_p3_insured as
    select s.sus_idx, o._row_idx_ as ref_idx, 'OSBAL' as ref_source, 'Insured only' as match_label
    from work.suspend_prep s
         inner join work.osbal_prep o
         on (upcase(strip(s.INSURED_CLEAN)) ne '' and (
             upcase(strip(s.INSURED_CLEAN)) = upcase(strip(o.INSURED_CLEAN_1)) or
             upcase(strip(s.INSURED_CLEAN)) = upcase(strip(o.INSURED_CLEAN_2)) or
             upcase(strip(s.INSURED_CLEAN)) = upcase(strip(o.FAC_INSURED))
         ));

    /* P4: Polis -> FACUL -> OSBAL */
    create table work.facul_hit_polis as
    select s.sus_idx, upcase(strip(f.FAC_CODE)) as fac_code, 'Polis only' as match_label
    from work.suspend_prep s
         inner join work.facul_prep f
         on (upcase(strip(s.POLIS_CLN)) ne '' and (
             upcase(strip(s.POLIS_CLN)) = upcase(strip(f.POLICY_CLEAN_1)) or
             upcase(strip(s.POLIS_CLN)) = upcase(strip(f.POLICY_CLEAN_2)) or
             upcase(strip(s.POLIS_CLN)) = upcase(strip(f.FAC_POLICY_NO))
         ));

    create table work.match_p4_facul_polis as
    select p.sus_idx, o._row_idx_ as ref_idx, 'FACUL' as ref_source, p.match_label
    from work.facul_hit_polis p
         inner join work.osbal_prep o
         on p.fac_code = upcase(strip(o.CCOS_REF_CODE));

    /* P5: Slip -> FACUL -> OSBAL */
    create table work.facul_hit_slip as
    select s.sus_idx, upcase(strip(f.FAC_CODE)) as fac_code, 'Slip only' as match_label
    from work.suspend_prep s
         inner join work.facul_prep f
         on (upcase(strip(s.SLIP_NO_CLN)) ne '' and (
             upcase(strip(s.SLIP_NO_CLN)) = upcase(strip(f.SLIP_CLEAN_1)) or
             upcase(strip(s.SLIP_NO_CLN)) = upcase(strip(f.FAC_SLIP))
         ));

    create table work.match_p5_facul_slip as
    select p.sus_idx, o._row_idx_ as ref_idx, 'FACUL' as ref_source, p.match_label
    from work.facul_hit_slip p
         inner join work.osbal_prep o
         on p.fac_code = upcase(strip(o.CCOS_REF_CODE));

    /* P6: Insured -> FACUL -> OSBAL */
    create table work.facul_hit_insured as
    select s.sus_idx, upcase(strip(f.FAC_CODE)) as fac_code, 'Insured only' as match_label
    from work.suspend_prep s
         inner join work.facul_prep f
         on (upcase(strip(s.INSURED_CLEAN)) ne '' and (
             upcase(strip(s.INSURED_CLEAN)) = upcase(strip(f.INSURED_CLEAN_1)) or
             upcase(strip(s.INSURED_CLEAN)) = upcase(strip(f.FAC_INSURED))
         ));

    create table work.match_p6_facul_insured as
    select p.sus_idx, o._row_idx_ as ref_idx, 'FACUL' as ref_source, p.match_label
    from work.facul_hit_insured p
         inner join work.osbal_prep o
         on p.fac_code = upcase(strip(o.CCOS_REF_CODE));
quit;

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

proc sort data=work.all_initial_matches; by sus_idx pass_priority ref_idx; run;

data work.best_initial_matches;
    set work.all_initial_matches;
    by sus_idx pass_priority;
    if first.sus_idx then _best_prio_ = pass_priority;
    retain _best_prio_;
    if pass_priority = _best_prio_;
    drop _best_prio_;
run;

/* Narrowing Atribut */
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
         inner join work.osbal_prep o on m.ref_idx = o._row_idx_
         inner join work.suspend_prep s on m.sus_idx = s.sus_idx;

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

data work.narrowed_confirmed;
    set work.periode_narrowed;
    length final_scenario $100;
    final_scenario = match_label;

    if match_label = 'Polis only' and SLIP_NO_CLN ne '' and SLIP_CLEAN_1 ne '' then do;
        if SLIP_NO_CLN = SLIP_CLEAN_1 then final_scenario = 'Polis + Slip';
    end;

    _cert_match_ = 0;
    if SERTIF_CLN_1 ne '' then do;
        if osb_sertif_1 ne '' and SERTIF_CLN_1 = osb_sertif_1 then _cert_match_ = 1;
        else if CLSDT_SERTF_NO ne '' and SERTIF_CLN_1 = CLSDT_SERTF_NO then _cert_match_ = 1;
    end;
    if _cert_match_ then final_scenario = catx(' ', final_scenario, '+ Cert');

    if INSURED_CLEAN ne '' and osb_ins_1 ne '' and index(osb_ins_1, strip(INSURED_CLEAN)) > 0 then do;
        if index(final_scenario, 'Insured') = 0 then final_scenario = catx(' ', final_scenario, '+ Insured');
    end;

    if osbal_curr ne sus_curr and osbal_curr ne '' and sus_curr ne '' then final_scenario = 'Beda Currency';
    else if not _per_ok_ then final_scenario = 'Beda Periode';
run;

/* Bordero Match */
proc sql noprint;
    create table work.sus_bdr_matches as
    select distinct s.sus_idx, b.fac_code, o._row_idx_ as ref_idx
    from work.suspend_prep s
         inner join work.bordero_combined b
         on (upcase(strip(s.POLIS_CLN)) ne '' and (
             upcase(strip(s.POLIS_CLN)) = b.bdr_polis or
             index(b.bdr_polis, upcase(strip(s.POLIS_CLN))) > 0 or
             index(upcase(strip(s.POLIS_CLN)), b.bdr_polis) > 0
         ))
         inner join work.osbal_prep o
         on b.fac_code = upcase(strip(o.CCOS_REF_CODE));
quit;

%let n_bdr_hits = %sysfunc(attrn(%sysfunc(open(work.sus_bdr_matches)), nobs));
%if &n_bdr_hits. > 0 %then %do;
    proc sql noprint;
        update work.narrowed_confirmed m
        set final_scenario = 'Polis only (Bordero)'
        where m.sus_idx in (select distinct sus_idx from work.sus_bdr_matches)
          and m.final_scenario not in ('Beda Currency', 'Beda Periode');
    quit;
%end;

proc sql noprint;
    create table work.unmatched_sus as
    select s.sus_idx, . as ref_idx, 'NONE' as ref_source, 'Unmatching' as match_label,
           999 as pass_priority, 'Unmatching' as final_scenario
    from work.suspend_prep s
    where s.sus_idx not in (select distinct sus_idx from work.narrowed_confirmed);
quit;

data work.all_matches_combined;
    set work.narrowed_confirmed (keep=sus_idx ref_idx ref_source match_label final_scenario)
        work.unmatched_sus;
run;


/* ═════════════════════════════════════════════════════════════════════════════
   PHASE 3: DERIVED COLUMNS, FLAG_PROD, FORMATTING 38 KOLOM & EXPORT
   ═════════════════════════════════════════════════════════════════════════════ */
%put [PHASE 3] Final Calculation & Export ...;

proc sql noprint;
    create table work.match_full_details as
    select m.sus_idx, m.ref_idx, m.ref_source, m.final_scenario,
           s."RECEIPT NO"n, s."CREDIT NOTES"n, s."DETAIL RINCIAN NO"n, s."RECEIPT DATE"n,
           s."CEDANT NAME"n, s."CEDANT SHRT NAME"n,
           s.INSURED as INSURED_ORI, s.INSURED_CLEAN as INSURED_1, '' as INSURED_2 length=500,
           s."CURR ORI"n, s."AMOUNT ORI"n, s."CURR PAY"n, s."AMOUNT PAY"n,
           s.POLIS, s.POLIS_CLN as POLICY_CLEAN_1, s.SERTIF_CLN_1,
           '' as POLICY_CLEAN_2 length=200, s.SERTIF_CLN_2,
           '' as POLICY_CLEAN_3 length=200, s.SERTIF_CLN_3,
           '' as POLICY_CLEAN_4 length=200, '' as POLICY_CLEAN_5 length=200,
           s.SLIP_NO, s.SLIP_NO_CLN,
           s."DESC 1"n, s."DESC 2"n, s."DESC 3"n, s."DESC 4"n,
           s.STATUS, s.REC_TYPE,
           o.CCOS_DOC_NO, o.CCOS_REF_CODE, o.CCOS_OR_BAL, o.CCOS_BAL_DUE
    from work.all_matches_combined m
         left join work.osbal_prep o on m.ref_idx = o._row_idx_
         left join work.suspend_prep s on m.sus_idx = s.sus_idx;

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

    if n_fac_codes > 1 or final_scenario in ('Beda Currency', 'Beda Periode', 'Unmatching') then do;
        CCOS_OR_BAL = .;
        CCOS_BAL_DUE = .;
        _bal_due_ = 0;
    end;

    DIFERENCE = _amount_neg_ - _bal_due_;

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
    outfile="&FINAL_OUTPUT_OUT."
    dbms=xlsx
    replace;
    sheet="Sheet1";
run;

%put =========================================================================;
%put [PIPELINE SELESAI] Hasil tersimpan di: &FINAL_OUTPUT_OUT.;
%put =========================================================================;

/* Ringkasan Frekuensi */
title "Distribusi FLAG_PROD - PT ASURANSI ASTRA BUANA";
proc freq data=work.final_output_astrabuana;
    tables FLAG_PROD / nocum;
run;

title "Distribusi SKENARIO - PT ASURANSI ASTRA BUANA";
proc freq data=work.final_output_astrabuana;
    tables SKENARIO / nocum;
run;
title;
