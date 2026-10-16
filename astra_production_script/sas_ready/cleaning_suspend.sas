/* =============================================================================
   PROGRAM   : CLEANING_SUSPEND.SAS
   CEDANT    : PT ASURANSI ASTRA BUANA
   FUNGSI    : Pembersihan data transaksi Suspend Astra Buana
   INPUT     : data/suspend.xlsx (atau 3b. Database Suspense 150826 (2).xlsx)
               Catatan: Header Excel berada pada baris ke-3 (datarow=4)
   OUTPUT    : data/astrabuana_output_suspend.xlsx
   KOMPATIBEL: SAS Viya / SAS 9.4 Studio
   ============================================================================= */

options validvarname=any validmemname=extend mprint symbolgen compress=yes;

/* ── 0. Path setup ─────────────────────────────────────────────────────────── */
%if not %symexist(BASE_DIR) %then %do;
    %let BASE_DIR = %str(c:\Users\berryl\Desktop\starcore\indore\all_script_cedant\astra_production_script);
%end;
%let DATA_DIR   = %str(&BASE_DIR.\data);
%let OUT_FILE   = %str(&DATA_DIR.\astrabuana_output_suspend.xlsx);
%let MACRO_INC  = %str(&BASE_DIR.\sas_ready\macros_cleaning_astra.sas);

/* Cari file input yang tersedia */
%let INPUT_FILE = %str(&DATA_DIR.\suspend.xlsx);
%macro detect_input();
    %if not %sysfunc(fileexist("&INPUT_FILE.")) %then %do;
        %let ALT_FILE = %str(&BASE_DIR.\dataExcel\raw\3b. Database Suspense 150826 (2).xlsx);
        %if %sysfunc(fileexist("&ALT_FILE.")) %then %let INPUT_FILE = &ALT_FILE.;
    %end;
%mend detect_input;
%detect_input();

/* ── 1. Include shared macros ──────────────────────────────────────────────── */
%if %sysfunc(fileexist("&MACRO_INC.")) %then %do;
    %include "&MACRO_INC.";
%end;

/* ── 2. Definisi konstanta bisnis ─────────────────────────────────────────── */
%let CEDANT_COL   = %str(CEDANT SHRT NAME);
%let CEDANT_VALUE = ASTRABUANA;
%let STATUS_COL   = STATUS;
%let STATUS_VALUE = SUSPENSE;


/* =============================================================================
   STEP 1: Baca file Excel Suspend
   Header di baris ke-3 (0-indexed header=2 di Pandas -> datarow=4 di SAS)
   ============================================================================= */
%put [1/4] Reading: &INPUT_FILE. (Header di baris ke-3) ...;

proc import
    datafile="&INPUT_FILE."
    out=work.suspend_raw
    dbms=xlsx
    replace;
    getnames=yes;
    datarow=4;
run;

%put       Total rows dibaca: %sysfunc(attrn(%sysfunc(open(work.suspend_raw)), nobs));


/* =============================================================================
   STEP 2: Filter Cedant ASTRABUANA & Status SUSPENSE
   ============================================================================= */
%put [2/4] Filtering cedant: &CEDANT_VALUE. & status: &STATUS_VALUE. ...;

data work.suspend_filtered;
    set work.suspend_raw;
    _cedant_chk_ = upcase(strip("CEDANT SHRT NAME"n));
    _status_chk_ = upcase(strip(STATUS));

    if _cedant_chk_ = "&CEDANT_VALUE." and _status_chk_ = "&STATUS_VALUE." then do;
        STATUS = "&STATUS_VALUE.";
        output;
    end;
    drop _cedant_chk_ _status_chk_;
run;

%put       Rows setelah filter: %sysfunc(attrn(%sysfunc(open(work.suspend_filtered)), nobs));


/* =============================================================================
   STEP 3: Cleaning Spesifik Astra Buana: Insured, Polis, Slip & Sertifikat
   ============================================================================= */
%put [3/4] Cleaning columns: INSURED_CLEAN, POLIS_CLN, SLIP_NO_CLN, SERTIF_CLN_1..3 ...;

data work.suspend_cleaned;
    set work.suspend_filtered;

    length INSURED_CLEAN $500
           POLIS_CLN     $200
           SERTIF_CLN_1  $100
           SERTIF_CLN_2  $100
           SERTIF_CLN_3  $100
           SLIP_NO_CLN   $200
           SLIP_NO       $200;

    /* 1. Bersihkan Insured */
    _ins_txt_ = prxchange('s/\b(?:PT|CV|TBK|PTE|LTD)\b//oi', -1, strip(INSURED));
    _ins_txt_ = compress(_ins_txt_, '.,');
    INSURED_CLEAN = upcase(strip(prxchange('s/\s+/ /', -1, _ins_txt_)));

    /* 2. Bersihkan Polis */
    _pol_txt_ = prxchange('s/-EXT\(\d+\)//oi', -1, strip(POLIS));
    if index(upcase(_pol_txt_), 'S/D') > 0 or index(upcase(_pol_txt_), 'S / D') > 0 then
        _pol_txt_ = prxchange('s/\s*S\s*\/\s*D\s*/SD/oi', -1, _pol_txt_);
    _pol_txt_ = compress(_pol_txt_, '.,');
    POLIS_CLN = strip(prxchange('s/\s+/ /', -1, _pol_txt_));

    /* 3. Ambil nilai slip mentah dan standardisasi penamaan kolom */
    _slip_orig_ = strip("SLIP NO"n);
    SLIP_NO     = _slip_orig_;

    /* 4. Ekstrak Sertifikat dari Slip No */
    _cert_txt_ = _slip_orig_;
    if index(upcase(_cert_txt_), 'S/D') > 0 or index(upcase(_cert_txt_), 'S / D') > 0 then
        _cert_txt_ = prxchange('s/\s*S\s*\/\s*D\s*/SD/oi', -1, _cert_txt_);

    SERTIF_CLN_1 = '';
    SERTIF_CLN_2 = '';
    SERTIF_CLN_3 = '';

    /* Pola 1: Range di akhir Slip No: -(\d{1,4})\s*(?:-|SD)\s*(\d{1,4})\s*$ */
    _rx_rng_ = prxparse('/-(\d{1,4})\s*(?:-|SD)\s*(\d{1,4})\s*$/oi');
    if prxmatch(_rx_rng_, _cert_txt_) then do;
        call prxposn(_rx_rng_, 1, _r_s1_, _r_l1_);
        call prxposn(_rx_rng_, 2, _r_s2_, _r_l2_);
        _raw1_ = substr(_cert_txt_, _r_s1_, _r_l1_);
        _raw2_ = substr(_cert_txt_, _r_s2_, _r_l2_);

        /* Pengecualian: jika 2 digit dan '00' maka blank */
        if not (lengthn(strip(_raw2_)) <= 2 and put(input(strip(_raw2_), ??8.), z2.) = '00') then do;
            _n1_ = input(_raw1_, 8.);
            _n2_ = input(_raw2_, 8.);
            if _n2_ < _n1_ then do;
                _tmp_n_ = _n1_; _n1_ = _n2_; _n2_ = _tmp_n_;
            end;
            _cnt_ = _n2_ - _n1_ + 1;
            _p1_ = put(_n1_, z6.);
            _p2_ = put(_n2_, z6.);
            if _cnt_ > 3 then do;
                SERTIF_CLN_1 = catx(' ', _p1_, 'SD', _p2_);
            end;
            else do;
                SERTIF_CLN_1 = put(_n1_, z6.);
                do _curr_n_ = _n1_ + 1 to _n2_;
                    SERTIF_CLN_1 = catx(', ', SERTIF_CLN_1, put(_curr_n_, z6.));
                end;
            end;
        end;
    end;
    else do;
        /* Pola 2: Single di akhir Slip No: -(\d{2})\s*$ */
        _rx_sgl_ = prxparse('/-(\d{2})\s*$/o');
        if prxmatch(_rx_sgl_, _cert_txt_) then do;
            call prxposn(_rx_sgl_, 1, _s_s1_, _s_l1_);
            _raw_s_ = substr(_cert_txt_, _s_s1_, _s_l1_);
            if strip(_raw_s_) ne '00' then do;
                SERTIF_CLN_1 = put(input(strip(_raw_s_), 8.), z6.);
            end;
        end;
        call prxfree(_rx_sgl_);
    end;
    call prxfree(_rx_rng_);

    /* 5. Bersihkan Slip No -> SLIP_NO_CLN */
    _slp_txt_ = prxchange('s/\s*-\s*New\s*$//oi', -1, _slip_orig_);
    if index(upcase(_slp_txt_), 'S/D') > 0 or index(upcase(_slp_txt_), 'S / D') > 0 then
        _slp_txt_ = prxchange('s/\s*S\s*\/\s*D\s*/SD/oi', -1, _slp_txt_);
    _slp_txt_ = compress(_slp_txt_, '.,');
    SLIP_NO_CLN = strip(prxchange('s/\s+/ /', -1, _slp_txt_));

    /* Clean control chars */
    array _str_cols{*} _character_;
    do _sci_ = 1 to dim(_str_cols);
        if not missing(_str_cols{_sci_}) then
            _str_cols{_sci_} = prxchange('s/[\x00-\x08\x0B\x0C\x0E-\x1F]//o', -1, _str_cols{_sci_});
    end;
    drop _sci_ _ins_txt_ _pol_txt_ _slip_orig_ _cert_txt_ _slp_txt_
         _r_s1_ _r_l1_ _r_s2_ _r_l2_ _raw1_ _raw2_ _n1_ _n2_ _tmp_n_ _cnt_ _p1_ _p2_ _curr_n_
         _s_s1_ _s_l1_ _raw_s_;
run;


/* =============================================================================
   STEP 4: Export ke Excel Output
   ============================================================================= */
%put [4/4] Saving output to: &OUT_FILE. ...;

proc export
    data=work.suspend_cleaned
    outfile="&OUT_FILE."
    dbms=xlsx
    replace;
    sheet="Sheet1";
run;

%put [SUSPEND CLEANING ASTRA] Selesai: &OUT_FILE.;
