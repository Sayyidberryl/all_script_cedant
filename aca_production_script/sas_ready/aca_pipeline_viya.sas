/* =========================================================================
   PROGRAM   : ACA_PIPELINE_VIYA.SAS
   CEDANT    : PT. ASURANSI CENTRAL ASIA (ACA)
   PLATFORM  : SAS Viya / SAS 9.4 Studio (Native Cloud Analytics Services)
   PURPOSE   : End-to-End Suspense Reconciliation Pipeline (Full Viya / DB Mode)
               - Automated Diagnostics (Active Libnames, CASLIBs, PostgreSQL)
               - In-Memory CAS Staging & Caching (PROC FEDSQL sessref=mycas)
               - Production Data Cleaning (Clean OSBAL, Clean Facul, Clean Suspend)
               - Multi-Pass Suspense Matching Engine & Dual Output Generation:
                 * Clean OSBAL (42 kolom standar)
                 * Clean Facul (36 kolom standar)
                 * Clean Suspend (In-Memory Staging CAS)
                 * Suspense V1 (38 kolom standar, KEEP_ROWS row-level)
                 * Suspense V2 (38 kolom standar, MERGE_ROWS aggregated)
               - Native Sync 100% Identik Python:
                 * In-Memory CAS & .sashdat (ACA_OSBAL_CLEAN, ACA_FACUL_CLEAN,
                   ACA_SUSPEND_CLEAN, ACA_MATCHING_V1, ACA_MATCHING_V2)
                 * PostgreSQL FSI_DB (SUSPENSE_DATA_OSBAL_CLEAN, SUSPENSE_DATA_FACUL_CLEAN,
                   SUSPENSE_DATA_SUSPENSE_V1, SUSPENSE_CLEAN_2026, SUSPENSE_DATA_SUSPENSE_V2)
   ========================================================================= */

options validvarname=any validmemname=extend nomprint nosymbolgen compress=yes;

/* --- 0. CONFIGURATION & PARAMETERS --- */

/* In-memory CAS & caching settings */
%let USE_CAS_CACHE     = 1;         /* 1 = gunakan cache staging di CAS jika ada, 0 = load ulang */
%let FORCE_REFRESH     = 0;         /* 1 = paksa re-query dan clean ulang dari database */
%let CAS_TARGET_LIB    = Public;    /* Target CASLIB untuk SAS Visual Analytics & persistensi disk */

/* PostgreSQL connection parameters (FSI_DB) */
%let DB_HOST           = 10.10.125.91;
%let DB_PORT           = 5432;
%let DB_NAME           = fsi_db;
%let DB_USER           = viy4user;
%let DB_PASS           = fild42op61nx;

/* Source table names in PostgreSQL (FSI_DB) */
%let DB_TABLE_FACUL    = SUSPENSE_FACUL_ORI_2026;
%let DB_TABLE_OSBAL    = SUSPENSE_OSBAL_ORI_2026;
%let DB_TABLE_SUSPEND  = SUSPENSE_ORI_2026;
%let DB_TABLE_BORDERO  = ACA_OPEN_COVER_MC;

%put [INIT] =========================================================================;
%put [INIT] ACA Suspense Pipeline initialized (Full SAS Viya Environment Mode);
%put [INIT] Database Source : &DB_HOST.:&DB_PORT./&DB_NAME. (Libname: FSI_DB);
%put [INIT] CAS Target      : &CAS_TARGET_LIB. (Session: mycas);
%put [INIT] Outputs         : Clean OSBAL, Clean Facul, Clean Suspend, Suspense V1 & V2;
%put [INIT] DB Sync Tables  : SUSPENSE_DATA_OSBAL_CLEAN, SUSPENSE_DATA_FACUL_CLEAN,;
%put [INIT]                   SUSPENSE_DATA_SUSPENSE_V1, SUSPENSE_CLEAN_2026,;
%put [INIT]                   SUSPENSE_DATA_SUSPENSE_V2;
%put [INIT] =========================================================================;

/* --- 1. ENVIRONMENT DIAGNOSTICS & LIBNAME SETUP --- */

proc sql noprint;
    create table work._diag_libnames as
    select distinct libname, engine, path
    from dictionary.libnames
    order by libname;
quit;

data _null_;
    set work._diag_libnames;
    put '[LIBNAME DIAGNOSTIC] ' libname $8. ' | Engine: ' engine $10. ' | Path: ' path;
run;

%macro init_and_diag_cas();
    %if not %symexist(_sessid_) %then %do;
        %put [CAS DIAGNOSTIC] Initializing CAS session (casauto)...;
        cas casauto sessopts=(caslib="casuser" timeout=3600);
    %end;
    caslib _all_ assign;
    %if %sysfunc(libref(casuser)) ne 0 %then %do;
        libname casuser cas caslib="CASUSER";
    %end;
    %if %sysfunc(libref(public)) ne 0 %then %do;
        libname public cas caslib="&CAS_TARGET_LIB.";
    %end;
    caslib _all_ list;
%mend init_and_diag_cas;
%init_and_diag_cas();

%macro check_assign_fsidb();
    %if %sysfunc(libref(fsi_db)) ne 0 %then %do;
        %put [DB DIAGNOSTIC] Assigning Libname FSI_DB to PostgreSQL (&DB_HOST.:&DB_PORT./&DB_NAME.)...;
        libname FSI_DB postgres
            server="&DB_HOST."
            port=&DB_PORT.
            user="&DB_USER."
            password="&DB_PASS."
            database="&DB_NAME."
            schema=public
            preserve_tab_names=yes;
    %end;
    %if %sysfunc(libref(fsi_db)) = 0 %then %do;
        %put [DB DIAGNOSTIC] Libname FSI_DB is active and ready.;
    %end;
    %else %do;
        %put [DB WARNING] Libname FSI_DB could not be assigned automatically.;
    %end;
%mend check_assign_fsidb;
%check_assign_fsidb();


/* --- 2. MACRO UTILITIES & FORMAT DEFINITIONS --- */

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

proc format;
    value $bulan_fmt
        'JANUARI'   = '1'
        'FEBRUARI'  = '2'
        'MARET'     = '3'
        'APRIL'     = '4'
        'MEI'       = '5'
        'JUNI'      = '6'
        'JULI'      = '7'
        'AGUSTUS'   = '8'
        'SEPTEMBER' = '9'
        'OKTOBER'   = '10'
        'NOVEMBER'  = '11'
        'DESEMBER'  = '12'
        other       = '0';
run;

%macro normalize_spaces(val);
    strip(prxchange('s/\s{2,}/ /', -1, strip(&val)))
%mend normalize_spaces;

%macro clean_control_chars(var);
    &var = compress(&var, '0001020304050607080B0C0E0F101112131415161718191A1B1C1D1E1F'x);
%mend clean_control_chars;

%macro is_valid_polis_token(tok);
    (lengthn(strip(&tok)) >= 5
     and anydigit(strip(&tok)) > 0
     and not prxmatch('/TBA$/oi', strip(&tok)))
%mend is_valid_polis_token;

%macro is_valid_slip_token(tok);
    (lengthn(strip(&tok)) >= 7 and anydigit(strip(&tok)) > 0)
%mend is_valid_slip_token;

%macro strip_polis_base(in_val, out_val);
    do;
        &out_val = strip(&in_val);
        _spb_wrk_ = prxchange('s/\s*-\s*/-/o', -1, strip(&in_val));
        _spb_nparts_ = countw(_spb_wrk_, '-', 'm');
        if _spb_nparts_ >= 2 then do;
            _spb_base_ = strip(scan(_spb_wrk_, 1, '-', 'm'));
            _spb_bdigits_ = lengthn(compress(_spb_base_, '0123456789', 'k'));
            if _spb_bdigits_ >= 10 then do;
                _spb_all_suf_ok_ = 1;
                do _spb_i_ = 2 to _spb_nparts_;
                    _spb_suf_ = strip(scan(_spb_wrk_, _spb_i_, '-', 'm'));
                    if lengthn(compress(_spb_suf_, '0123456789', 'k')) > 6 then do;
                        _spb_all_suf_ok_ = 0;
                        leave;
                    end;
                end;
                if _spb_all_suf_ok_ then &out_val = strip(_spb_base_);
            end;
        end;
        if &out_val = strip(&in_val) and index(&out_val, '-') > 0 then do;
            _spb_dp_ = find(&out_val, '-', -lengthn(&out_val));
            if _spb_dp_ > 1 then do;
                _spb_p1_ = substr(&out_val, 1, _spb_dp_ - 1);
                _spb_p2_ = substr(&out_val, _spb_dp_ + 1);
                if (lengthn(_spb_p2_) >= 1 and lengthn(_spb_p2_) <= 6) and notdigit(strip(_spb_p2_)) = 0
                   and lengthn(_spb_p1_) > lengthn(_spb_p2_) then
                    &out_val = strip(_spb_p1_);
            end;
        end;
    end;
%mend strip_polis_base;

/* apply_extract_sertif: matches Python extract_cert_from_polis() + S/D range
   PERF: retain _rxse_sd_ & _rxse_sfx_ so prxparse() runs ONCE per DATA step. */
%macro apply_extract_sertif(in_var);
    retain _rxse_sd_ _rxse_sfx_ 0;
    array _cs_out{5} $10 clean_sertif_1-clean_sertif_5;
    do _i_ = 1 to 5; _cs_out{_i_} = ''; end;
    _sertif_wrk_ = strip(&in_var);
    _cs_n_ = 0;

    if missing(_sertif_wrk_) or _sertif_wrk_ = '' then goto _sertif_done_;

    /* Lazy-compile ONCE (first row only), reuse retained ID untuk semua row */
    if _rxse_sd_  = 0 then _rxse_sd_  = prxparse('/(\d{1,6})\s*S\/D\s*(\d{1,6})/oi');
    if _rxse_sfx_ = 0 then _rxse_sfx_ = prxparse('/[-\s]+\s*(\d{1,6})(?:[^0-9]|$)/o');

    /* S/D range: '000211 S/D 000215' -> ['000211','000212',...] */
    if prxmatch(_rxse_sd_, _sertif_wrk_) then do;
        _sd_start_ = input(prxposn(_rxse_sd_, 1, _sertif_wrk_), 8.);
        _sd_end_   = input(prxposn(_rxse_sd_, 2, _sertif_wrk_), 8.);
        if _sd_start_ > _sd_end_ then do;
            _sd_tmp_ = _sd_start_; _sd_start_ = _sd_end_; _sd_end_ = _sd_tmp_;
        end;
        do _sd_i_ = _sd_start_ to min(_sd_end_, _sd_start_ + 4);
            _cs_n_ + 1;
            if _cs_n_ <= 5 then _cs_out{_cs_n_} = put(_sd_i_, z6.);
        end;
        goto _sertif_done_;
    end;

    /* Suffix digit 1-6 setelah '-' (Python: _CERT_FROM_POLIS_RE) */
    if prxmatch(_rxse_sfx_, _sertif_wrk_) then do;
        _cert_raw_ = prxposn(_rxse_sfx_, 1, _sertif_wrk_);
        if prxmatch('/^\d{1,6}$/o', strip(_cert_raw_)) then do;
            _cs_n_ = 1;
            clean_sertif_1 = put(input(strip(_cert_raw_), 8.), z6.);
        end;
    end;

    _sertif_done_: ;
    drop _sertif_wrk_ _cs_n_ _sd_start_ _sd_end_ _sd_tmp_ _sd_i_ _cert_raw_ _i_
         _rxse_sd_ _rxse_sfx_;
%mend apply_extract_sertif;

/* ==========================================================================
   HELPER: %_extract_polis_tokens(text_var, pfx=clean_polis, tag=cp)
   Equiv Python: _extract_polis_tokens(text) in cleaning_facul.py
   Splits by 2+ spaces, then by + / whitespace; validates & strips each token.
   Writes into _&tag._out{} / _cp_n_&tag._ (must be in scope from caller).
   ========================================================================== */
%macro _extract_polis_tokens(text_var, pfx=clean_polis, tag=cp);
    _ept_work_ = prxchange('s/\s{2,}/|/o', -1, strip(&text_var));
    _ept_nblk_ = countw(_ept_work_, '|', 'm');
    do _ept_bi_ = 1 to _ept_nblk_;
        if _cp_n_&tag._ >= 5 then leave;
        _ept_blk_  = scan(_ept_work_, _ept_bi_, '|', 'm');
        _ept_blk_  = prxchange('s/\+/ /o', -1, _ept_blk_);
        _ept_blk_  = prxchange('s/\s{2,}/ /o', -1, strip(_ept_blk_));
        _ept_ntok_ = countw(strip(_ept_blk_), ' ');
        do _ept_ti_ = 1 to _ept_ntok_;
            if _cp_n_&tag._ >= 5 then leave;
            _ept_tok_ = strip(scan(strip(_ept_blk_), _ept_ti_, ' '));
            _ept_tok_ = prxchange('s/^[-\/]+|[-\/]+$//o', -1, _ept_tok_);
            _ept_tok_ = strip(_ept_tok_);
            if %is_valid_polis_token(_ept_tok_) then do;
                _ept_ndig_ = lengthn(compress(_ept_tok_, '', 'kd'));
                if _ept_ndig_ >= 10 then %strip_polis_base(_ept_tok_, _ept_tok_);
                _cp_n_&tag._ + 1; _&tag._out{_cp_n_&tag._} = strip(_ept_tok_);
            end;
        end;
    end;
%mend _extract_polis_tokens;

/* ==========================================================================
   HELPER: %_expand_plus_suffix(text_var, pfx=clean_polis, tag=cp)
   Equiv Python: _expand_plus_suffix(text) in cleaning_facul.py
   Expands BASE+s1+s2 where BASE is pure digits >=10 and si replaces last digits.
   Sets _eps_ok_=1 and writes into _&tag._out{}/_cp_n_&tag._ if expansion succeeds.
   ========================================================================== */
%macro _expand_plus_suffix(text_var, pfx=clean_polis, tag=cp);
    _eps_ok_     = 0;
    _eps_nparts_ = countw(strip(&text_var), '+', 'm');
    if _eps_nparts_ >= 2 then do;
        _eps_base_ = strip(scan(strip(&text_var), 1, '+', 'm'));
        if prxmatch('/^\d{10,}$/o', _eps_base_) then do;
            _eps_valid_ = 1;
            do _eps_si_ = 2 to _eps_nparts_;
                _eps_seg_ = strip(scan(strip(&text_var), _eps_si_, '+', 'm'));
                if not prxmatch('/^\d+$/o', _eps_seg_)
                   or lengthn(_eps_seg_) >= lengthn(_eps_base_)
                then do; _eps_valid_ = 0; leave; end;
            end;
            if _eps_valid_ then do;
                _eps_ok_ = 1;
                _cp_n_&tag._ + 1; if _cp_n_&tag._ <= 5 then _&tag._out{_cp_n_&tag._} = _eps_base_;
                do _eps_si_ = 2 to _eps_nparts_;
                    _eps_seg_      = strip(scan(strip(&text_var), _eps_si_, '+', 'm'));
                    _eps_expanded_ = substr(_eps_base_, 1, lengthn(_eps_base_) - lengthn(_eps_seg_))
                                     || strip(_eps_seg_);
                    _cp_n_&tag._ + 1; if _cp_n_&tag._ <= 5 then _&tag._out{_cp_n_&tag._} = strip(_eps_expanded_);
                end;
            end;
        end;
    end;
%mend _expand_plus_suffix;

/* ==========================================================================
   apply_clean_polis: 100% faithful to Python cleaning_facul.clean_polis()
   All scenarios ported: exception, dot-suffix, S/D, TBA+content, &, +expand,
   double-space multi-token, multi-slash with all-short check, single fallback.
   PERF: retain _rxcp_dot_ & _rxcp_sd_ -- prxparse() runs ONCE per DATA step.
   Supports custom pfx & tag so multiple calls in one DATA step do not collide.
   ========================================================================== */
%macro apply_clean_polis(in_var, pfx=clean_polis, tag=cp);
    retain _rxcp_dot_ _rxcp_sd_ 0;
    array _&tag._out{5} $500 &pfx._1-&pfx._5;
    do _i_&tag._ = 1 to 5; _&tag._out{_i_&tag._} = ''; end;
    _polis_wrk_ = strip(&in_var);
    _cp_n_&tag._ = 0;

    if missing(_polis_wrk_) or _polis_wrk_ in ('', '0', '-') then goto _pdone_&tag.;

    _polis_up_ = upcase(strip(_polis_wrk_));

    /* 1. POLIS_EXCEPTION_RE -> return normalize(val) [Python lines 341-342] */
    if prxmatch('/MOP\s*MARINE|LINE\s*SLIP|LINESLIP|\bP[123]\b|\bP73\b|\bCANCEL\b|PENYELESAIAN|HUTANG\s*PIUTANG/oi', _polis_wrk_) then do;
        _cp_n_&tag._ = 1; &pfx._1 = %normalize_spaces(_polis_wrk_); goto _pdone_&tag.;
    end;

    /* 2. digit+TBA+digit -> return normalize(val) [Python lines 344-345] */
    if prxmatch('/\d+TBA\d+/oi', _polis_wrk_) then do;
        _cp_n_&tag._ = 1; &pfx._1 = %normalize_spaces(_polis_wrk_); goto _pdone_&tag.;
    end;

    /* 3. VARIOUS / TBA only -> return as-is [Python lines 347-353] */
    if prxmatch('/^VARIOUS\s*$/oi', _polis_up_)
       or prxmatch('/^VARIOUS\s*-\s*SEE\s+ATTACH/oi', _polis_up_)
       or prxmatch('/^TBA\s*$/oi', _polis_up_)
    then do;
        _cp_n_&tag._ = 1; &pfx._1 = strip(_polis_wrk_); goto _pdone_&tag.;
    end;

    /* Lazy-compile patterns ONCE */
    if _rxcp_dot_ = 0 then _rxcp_dot_ = prxparse('/^(\d[\d\-]+)\.\d{1,6}$/');
    if _rxcp_sd_  = 0 then _rxcp_sd_  = prxparse('/^(\d[\d\-]+?)[-]?\d+S\/D\d+/oi');

    /* 4. Dot-suffix: '100030817120000016.000149' -> base [Python lines 355-360] */
    if prxmatch('/^\d[\d\-]+\.\d{1,6}$/o', _polis_wrk_) then do;
        if prxmatch(_rxcp_dot_, _polis_wrk_) then do;
            _base_dot_ = prxposn(_rxcp_dot_, 1, _polis_wrk_);
            if %is_valid_polis_token(_base_dot_) then do;
                _cp_n_&tag._ = 1; &pfx._1 = strip(_base_dot_); goto _pdone_&tag.;
            end;
        end;
    end;

    /* 5. S/D range: '100170221120000071 - 000001 S/D 001127' -> base [Python lines 362-367] */
    if prxmatch('/S\/D/oi', _polis_wrk_) then do;
        if prxmatch(_rxcp_sd_, _polis_wrk_) then do;
            _base_sd_ = prxposn(_rxcp_sd_, 1, _polis_wrk_);
            if %is_valid_polis_token(_base_sd_) then do;
                _cp_n_&tag._ = 1; &pfx._1 = strip(_base_sd_); goto _pdone_&tag.;
            end;
        end;
    end;

    /* 6. TBA followed by content [Python lines 369-373] */
    if prxmatch('/^\s*TBA\s{2,}/oi', _polis_wrk_) then do;
        _tba_rest_ = strip(prxchange('s/^\s*TBA\s+//oi', 1, _polis_wrk_));
        %_extract_polis_tokens(_tba_rest_, pfx=&pfx., tag=&tag.);
        if _cp_n_&tag._ > 0 then goto _pdone_&tag.;
    end;

    /* 7. Contains '&' -> extract tokens [Python lines 375-378] */
    if index(_polis_wrk_, '&') > 0 then do;
        %_extract_polis_tokens(_polis_wrk_, pfx=&pfx., tag=&tag.);
        if _cp_n_&tag._ > 0 then goto _pdone_&tag.;
    end;

    /* 8. Contains '+' [Python lines 380-401] */
    if index(_polis_wrk_, '+') > 0 then do;
        /* Double-space check: entry blocks */
        if prxmatch('/\s{2,}/o', _polis_wrk_) then do;
            _pls_norm_ = prxchange('s/\s{2,}/|/o', -1, strip(_polis_wrk_));
            _pls_nblk_ = countw(_pls_norm_, '|', 'm');
            _pls_entry_ = 0;
            do _pbi_ = 1 to _pls_nblk_;
                _pblk_ = scan(_pls_norm_, _pbi_, '|', 'm');
                if index(_pblk_, '+') = 0 and strip(_pblk_) ne '' then _pls_entry_ = 1;
            end;
            if _pls_entry_ then do;
                do _pbi_ = 1 to _pls_nblk_;
                    if _cp_n_&tag._ >= 5 then leave;
                    _pblk_ = scan(_pls_norm_, _pbi_, '|', 'm');
                    if index(_pblk_, '+') = 0 and strip(_pblk_) ne '' then do;
                        _pblk_ = strip(_pblk_);
                        _pntok_ = countw(strip(_pblk_), ' ');
                        do _pti_ = 1 to _pntok_;
                            if _cp_n_&tag._ >= 5 then leave;
                            _ptok_ = strip(scan(strip(_pblk_), _pti_, ' '));
                            _ptok_ = prxchange('s/^[-\/]+|[-\/]+$//o', -1, _ptok_);
                            _ptok_ = strip(_ptok_);
                            if %is_valid_polis_token(_ptok_) then do;
                                _ptok_nd_ = lengthn(compress(_ptok_, '', 'kd'));
                                if _ptok_nd_ >= 10 then %strip_polis_base(_ptok_, _ptok_);
                                _cp_n_&tag._ + 1; _&tag._out{_cp_n_&tag._} = strip(_ptok_);
                            end;
                        end;
                    end;
                end;
                if _cp_n_&tag._ > 0 then goto _pdone_&tag.;
            end;
        end;
        /* Try _expand_plus_suffix */
        %_expand_plus_suffix(_polis_wrk_, pfx=&pfx., tag=&tag.);
        if _eps_ok_ then goto _pdone_&tag.;
        /* Fallback: tokenize */
        %_extract_polis_tokens(_polis_wrk_, pfx=&pfx., tag=&tag.);
        if _cp_n_&tag._ > 0 then goto _pdone_&tag.;
    end;

    /* 9. Double-space without '+' [Python lines 403-406] */
    if prxmatch('/\s{2,}/o', _polis_wrk_) then do;
        %_extract_polis_tokens(_polis_wrk_, pfx=&pfx., tag=&tag.);
        if _cp_n_&tag._ > 0 then goto _pdone_&tag.;
    end;

    /* 10. Strip VARIOUS / VAR [Python lines 408-410] */
    _polis_wrk_ = prxchange('s/\bVARIOUS\b\s*//oi', -1, _polis_wrk_);
    _polis_wrk_ = prxchange('s/\bVAR\b\s*//oi', -1, _polis_wrk_);
    _polis_wrk_ = strip(_polis_wrk_);
    if _polis_wrk_ in ('', '0', '-') then goto _pdone_&tag.;

    /* 11. Multi-slash with all-short check [Python lines 412-426] */
    _slash_n_ = countw(_polis_wrk_, '/', 'm');
    if _slash_n_ > 1 then do;
        _all_short_ = 1; _nclean_ = 0;
        do _si_ = 1 to _slash_n_;
            _tok_ = prxchange('s/^[ \-\/]+|[ \-\/]+$//o', -1,
                              prxchange('s/\s{2,}/ /o', -1, strip(scan(_polis_wrk_, _si_, '/', 'm'))));
            _tok_ = strip(_tok_);
            if _tok_ ne '' then do; _nclean_ + 1; if lengthn(_tok_) >= 7 then _all_short_ = 0; end;
        end;
        if _nclean_ > 1 and _all_short_ then do; _cp_n_&tag._ = 1; &pfx._1 = strip(_polis_wrk_); end;
        else do;
            do _si_ = 1 to _slash_n_;
                if _cp_n_&tag._ >= 5 then leave;
                _tok_ = prxchange('s/^[ \-\/]+|[ \-\/]+$//o', -1,
                                  prxchange('s/\s{2,}/ /o', -1, strip(scan(_polis_wrk_, _si_, '/', 'm'))));
                _tok_ = strip(_tok_);
                if _tok_ ne '' and %is_valid_polis_token(_tok_) then do;
                    _cp_n_&tag._ + 1; _&tag._out{_cp_n_&tag._} = _tok_;
                end;
            end;
        end;
    end;

    /* 12. Fallback single value */
    if _cp_n_&tag._ = 0 then do;
        _tok_ = strip(_polis_wrk_);
        if %is_valid_polis_token(_tok_) then do;
            _ndig_tok_ = lengthn(compress(_tok_, '', 'kd'));
            if _ndig_tok_ >= 10 then %strip_polis_base(_tok_, _tok_);
            _cp_n_&tag._ = 1; &pfx._1 = _tok_;
        end;
    end;

    _pdone_&tag.: ;
    if _cp_n_&tag._ > 5 then do;
        _joined_ = catx(',', of &pfx._1-&pfx._5);
        do _i_&tag._ = 1 to 5; _&tag._out{_i_&tag._} = ''; end;
        &pfx._1 = strip(_joined_);
    end;
    drop _polis_wrk_ _polis_up_ _cp_n_&tag._ _si_ _slash_n_ _tok_ _ndig_tok_
         _base_dot_ _base_sd_ _joined_ _i_&tag._ _rxcp_dot_ _rxcp_sd_
         _tba_rest_ _all_short_ _nclean_
         _ept_work_ _ept_nblk_ _ept_bi_ _ept_blk_ _ept_ntok_ _ept_ti_ _ept_tok_ _ept_ndig_
         _eps_ok_ _eps_nparts_ _eps_base_ _eps_valid_ _eps_si_ _eps_seg_ _eps_expanded_
         _pls_norm_ _pls_nblk_ _pls_entry_ _pbi_ _pblk_ _pntok_ _pti_ _ptok_ _ptok_nd_
         _spb_wrk_ _spb_nparts_ _spb_base_ _spb_bdigits_ _spb_all_suf_ok_ _spb_i_ _spb_suf_
         _spb_dp_ _spb_p1_ _spb_p2_;
%mend apply_clean_polis;

/* ==========================================================================
   HELPER: %_extract_slip_tokens(text_var, pfx=clean_slip, tag=csl)
   Equiv Python: _extract_slip_tokens(text) in cleaning_facul.py
   Splits by 2+ spaces; per block: noise removal -> token extraction with
   dash-split logic and _strip_slip_suffix for single tokens.
   Writes into _&tag._out{} / _csl_n_&tag._ (must be in scope from caller).
   ========================================================================== */
%macro _extract_slip_tokens(text_var, pfx=clean_slip, tag=csl);
    _est_norm_ = prxchange('s/\s{2,}/|/o', -1, strip(&text_var));
    _est_nblk_ = countw(_est_norm_, '|', 'm');
    do _est_bi_ = 1 to _est_nblk_;
        if _csl_n_&tag._ >= 5 then leave;
        _est_blk_   = scan(_est_norm_, _est_bi_, '|', 'm');
        /* Remove noise words (equiv Python: _SLIP_NOISE_RE.sub(" ", block)) */
        _est_clean_ = prxchange('s/\b(?:JANUARI|FEBRUARI|MARET|APRIL|MEI|JUNI|JULI|AGUSTUS|SEPTEMBER|OKTOBER|NOVEMBER|DESEMBER)\b/ /oi', -1, _est_blk_);
        _est_clean_ = prxchange('s/\b(?:JANUARY|FEBRUARY|MARCH|APRIL|MAY|JUNE|JULY|AUGUST|SEPTEMBER|OCTOBER|NOVEMBER|DECEMBER)\b/ /oi', -1, _est_clean_);
        _est_clean_ = prxchange('s/\b(?:IDR|USD|ENG|P1|P2|P3|P73|NEW|VARIOUS|20\d{2})\b/ /oi', -1, _est_clean_);
        _est_clean_ = prxchange('s/^[\s\-\/+,]+|[\s\-\/+,]+$//o', -1, strip(_est_clean_));
        _est_clean_ = strip(_est_clean_);
        if lengthn(_est_clean_) < 7 then continue;
        /* Tokenize: split on non-[A-Z0-9-] (equiv Python: _SLIP_TOKEN_RE.findall) */
        _est_tstr_  = prxchange('s/[^A-Za-z0-9\-]/ /oi', -1, _est_clean_);
        _est_tstr_  = prxchange('s/\s{2,}/ /o', -1, _est_tstr_);
        _est_tstr_  = strip(_est_tstr_);
        _est_ntok_  = countw(strip(_est_tstr_), ' ');
        do _est_ti_ = 1 to _est_ntok_;
            if _csl_n_&tag._ >= 5 then leave;
            _est_cand_ = strip(scan(strip(_est_tstr_), _est_ti_, ' '));
            /* Must match [A-Za-z0-9][A-Za-z0-9-]{6,} */
            if not prxmatch('/^[A-Za-z0-9][A-Za-z0-9\-]{6,}$/o', _est_cand_) then continue;
            /* Strip leading/trailing dashes (Python: cand.strip("-")) */
            _est_cand_ = prxchange('s/^-+|-+$//o', -1, _est_cand_);
            /* Dash-split: if exactly ONE dash and BOTH parts are valid with len diff <= 2 */
            if countc(_est_cand_, '-') = 1 then do;
                _est_dp_  = index(_est_cand_, '-');
                _est_pa_  = strip(substr(_est_cand_, 1, _est_dp_ - 1));
                _est_pb_  = strip(substr(_est_cand_, _est_dp_ + 1));
                if %is_valid_slip_token(_est_pa_) and %is_valid_slip_token(_est_pb_)
                   and abs(lengthn(_est_pa_) - lengthn(_est_pb_)) <= 2
                then do;
                    if _csl_n_&tag._ < 5 then do; _csl_n_&tag._ + 1; _&tag._out{_csl_n_&tag._} = _est_pa_; end;
                    if _csl_n_&tag._ < 5 then do; _csl_n_&tag._ + 1; _&tag._out{_csl_n_&tag._} = _est_pb_; end;
                    continue;
                end;
            end;
            /* Single token: apply _strip_slip_suffix then validate */
            if index(_est_cand_, '-') > 0 then do;
                _est_dp_ = find(_est_cand_, '-', -lengthn(_est_cand_));
                if _est_dp_ > 1 then do;
                    _est_p1_ = substr(_est_cand_, 1, _est_dp_ - 1);
                    _est_p2_ = substr(_est_cand_, _est_dp_ + 1);
                    if (lengthn(_est_p2_) >= 1 and lengthn(_est_p2_) <= 6) and notdigit(strip(_est_p2_)) = 0
                       and lengthn(_est_p1_) > lengthn(_est_p2_) then
                        _est_cand_ = strip(_est_p1_);
                end;
            end;
            if %is_valid_slip_token(_est_cand_) then do;
                if _csl_n_&tag._ < 5 then do;
                    _csl_n_&tag._ + 1; _&tag._out{_csl_n_&tag._} = strip(_est_cand_);
                end;
            end;
        end;
    end;
%mend _extract_slip_tokens;

/* ==========================================================================
   apply_clean_slip: 100% faithful to Python cleaning_facul.clean_slip()
   All scenarios ported: + with double-space entry blocks, multi-block left check,
   full _extract_slip_tokens with dash-split & _strip_slip_suffix, fallback.
   PERF: No PRXNEXT -- uses helper macro with SCAN.
   Supports custom pfx & tag so multiple calls in one DATA step do not collide.
   ========================================================================== */
%macro apply_clean_slip(in_var, pfx=clean_slip, tag=csl);
    array _&tag._out{5} $200 &pfx._1-&pfx._5;
    do _i_&tag._ = 1 to 5; _&tag._out{_i_&tag._} = ''; end;
    _slip_wrk_ = strip(&in_var);
    _csl_n_&tag._ = 0;

    if missing(_slip_wrk_) or _slip_wrk_ in ('', '0', '-') then goto _sdone_&tag.;

    /* 1. + AND double-space: entry blocks (without +) are the real slip tokens */
    if index(_slip_wrk_, '+') > 0 and prxmatch('/\s{2,}/o', _slip_wrk_) then do;
        _sla_norm_ = prxchange('s/\s{2,}/|/o', -1, strip(_slip_wrk_));
        _sla_nblk_ = countw(_sla_norm_, '|', 'm');
        _sla_entry_ = 0;
        do _sla_bi_ = 1 to _sla_nblk_;
            _sla_blk_ = scan(_sla_norm_, _sla_bi_, '|', 'm');
            if index(_sla_blk_, '+') = 0 and strip(_sla_blk_) ne '' then _sla_entry_ = 1;
        end;
        if _sla_entry_ then do;
            do _sla_bi_ = 1 to _sla_nblk_;
                if _csl_n_&tag._ >= 5 then leave;
                _sla_blk_ = scan(_sla_norm_, _sla_bi_, '|', 'm');
                if index(_sla_blk_, '+') = 0 and strip(_sla_blk_) ne '' then do;
                    %_extract_slip_tokens(_sla_blk_, pfx=&pfx., tag=&tag.);
                end;
            end;
            if _csl_n_&tag._ > 0 then goto _sdone_&tag.;
        end;
    end;

    /* 2. Multi-block left check: if left block is not a valid slip token,
       skip it and use the right blocks only */
    if prxmatch('/\s{2,}/o', _slip_wrk_) then do;
        _slb_norm_ = prxchange('s/\s{2,}/|/o', -1, strip(_slip_wrk_));
        _slb_nblk_ = countw(_slb_norm_, '|', 'm');
        if _slb_nblk_ >= 2 then do;
            _slb_left_ = scan(_slb_norm_, 1, '|', 'm');
            _slb_lchk_ = prxchange('s/\b(?:SUMMARY|BORDER[OA]?|BORDRO|SINGGLESHIPMENT|VARIOUS)\b//oi', -1, _slb_left_);
            _slb_lchk_ = prxchange('s/\b(?:JANUARI|FEBRUARI|MARET|APRIL|MEI|JUNI|JULI|AGUSTUS|SEPTEMBER|OKTOBER|NOVEMBER|DESEMBER)\b//oi', -1, _slb_lchk_);
            _slb_lchk_ = prxchange('s/\b(?:JANUARY|FEBRUARY|MARCH|APRIL|MAY|JUNE|JULY|AUGUST|SEPTEMBER|OCTOBER|NOVEMBER|DECEMBER)\b//oi', -1, _slb_lchk_);
            _slb_lchk_ = prxchange('s/\b(?:IDR|USD|ENG|P1|P2|P3|P73|NEW|VARIOUS|20\d{2})\b//oi', -1, _slb_lchk_);
            _slb_lchk_ = prxchange('s/[\s\+\-\/]+//o', -1, _slb_lchk_);
            _slb_lchk_ = strip(_slb_lchk_);
            if not %is_valid_slip_token(_slb_lchk_) then do;
                _slb_rest_ = '';
                do _slb_bi_ = 2 to _slb_nblk_;
                    if _slb_rest_ = '' then _slb_rest_ = strip(scan(_slb_norm_, _slb_bi_, '|', 'm'));
                    else _slb_rest_ = strip(_slb_rest_) || '  '
                                      || strip(scan(_slb_norm_, _slb_bi_, '|', 'm'));
                end;
                %_extract_slip_tokens(_slb_rest_, pfx=&pfx., tag=&tag.);
                if _csl_n_&tag._ > 0 then goto _sdone_&tag.;
            end;
        end;
    end;

    /* 3. Main extraction */
    %_extract_slip_tokens(_slip_wrk_, pfx=&pfx., tag=&tag.);
    if _csl_n_&tag._ > 0 then goto _sdone_&tag.;

    /* 4. Fallback: strip VARIOUS and leading dash */
    _slip_fb_ = prxchange('s/\bVARIOUS\b\s*//oi', -1, _slip_wrk_);
    _slip_fb_ = prxchange('s/^\s*-\s*//o', -1, strip(_slip_fb_));
    _slip_fb_ = prxchange('s/\s{2,}/ /o', -1, strip(_slip_fb_));
    _slip_fb_ = strip(_slip_fb_);
    if _slip_fb_ ne '' then do; _csl_n_&tag._ = 1; &pfx._1 = _slip_fb_; end;

    _sdone_&tag.: ;
    if _csl_n_&tag._ > 5 then do;
        _slipjoin_ = catx(',', of &pfx._1-&pfx._5);
        do _i_&tag._ = 1 to 5; _&tag._out{_i_&tag._} = ''; end;
        &pfx._1 = strip(_slipjoin_);
    end;
    drop _slip_wrk_ _csl_n_&tag._ _slip_fb_ _slipjoin_ _i_&tag._
         _est_norm_ _est_nblk_ _est_bi_ _est_blk_ _est_clean_
         _est_tstr_ _est_ntok_ _est_ti_ _est_cand_ _est_dp_
         _est_pa_ _est_pb_ _est_g1_ _est_g2_ _est_p1_ _est_p2_
         _sla_norm_ _sla_nblk_ _sla_bi_ _sla_blk_ _sla_entry_
         _slb_norm_ _slb_nblk_ _slb_left_ _slb_lchk_ _slb_bi_ _slb_rest_;
%mend apply_clean_slip;

/* Validasi hasil CLSDT tokens array (equiv Python: _clsdt_result_is_valid) */
%macro is_clsdt_valid_array(arr_pfx, result_var);
    &result_var = 0;
    array _chk_&arr_pfx._{5} $500 &arr_pfx._1-&arr_pfx._5;
    do _ci_&arr_pfx._ = 1 to 5;
        if &result_var = 0 and not missing(_chk_&arr_pfx._{_ci_&arr_pfx._}) and _chk_&arr_pfx._{_ci_&arr_pfx._} ne '' then do;
            if prxmatch('/\d{10,}/', strip(_chk_&arr_pfx._{_ci_&arr_pfx._})) then &result_var = 1;
            else if prxmatch('/\d{7,}/', strip(_chk_&arr_pfx._{_ci_&arr_pfx._}))
                 and not prxmatch('/\bTBA\b/oi', strip(_chk_&arr_pfx._{_ci_&arr_pfx._})) then &result_var = 1;
        end;
    end;
    drop _ci_&arr_pfx._;
%mend is_clsdt_valid_array;

/* Ekstrak sertifikat OSBAL (equiv Python: extract_cert_from_polis_osbal + clean_sertif) */
%macro apply_extract_sertif_osbal(polis_var, clsdt_sertf_var);
    retain _rxso_sd1_ _rxso_sd2_ 0;
    if _rxso_sd1_ = 0 then _rxso_sd1_ = prxparse('/^(\d{10,})\s*-\s*(\d+)\s*(?:S\/D|SD)\s*(\d+)$/oi');
    if _rxso_sd2_ = 0 then _rxso_sd2_ = prxparse('/^(\d{10,})\s+(\d+)\s*(?:S\/D|SD)\s*(\d+)$/oi');

    clean_sertif_1 = '';
    clean_sertif_2 = '';
    clean_sertif_3 = '';
    clean_sertif_4 = '';
    clean_sertif_5 = '';

    _p_osb_ = strip(&polis_var.);
    if not missing(_p_osb_) and _p_osb_ ne '' then do;
        if prxmatch(_rxso_sd1_, _p_osb_) then do;
            _st_ = strip(prxposn(_rxso_sd1_, 2, _p_osb_));
            _en_ = strip(prxposn(_rxso_sd1_, 3, _p_osb_));
            clean_sertif_1 = put(input(_st_, 8.), z6.) || ' SD ' || put(input(_en_, 8.), z6.);
        end;
        else if prxmatch(_rxso_sd2_, _p_osb_) then do;
            _st_ = strip(prxposn(_rxso_sd2_, 2, _p_osb_));
            _en_ = strip(prxposn(_rxso_sd2_, 3, _p_osb_));
            clean_sertif_1 = put(input(_st_, 8.), z6.) || ' SD ' || put(input(_en_, 8.), z6.);
        end;
        else if index(_p_osb_, '/') > 0 then do;
            _posb_cl_ = translate(_p_osb_, ' ', byte(191));
            _n_slash_parts_ = countw(_posb_cl_, ' /', 'm');
            if _n_slash_parts_ >= 2 then do;
                _base_p_ = scan(_posb_cl_, 1, ' /', 'm');
                if lengthn(compress(_base_p_, '', 'kd')) = lengthn(_base_p_) and lengthn(_base_p_) >= 10 then do;
                    _is_v_slash_ = 1;
                    do _spi_ = 2 to _n_slash_parts_;
                        _seg_ = scan(_posb_cl_, _spi_, ' /', 'm');
                        if lengthn(compress(_seg_, '', 'kd')) ne lengthn(_seg_) or lengthn(_seg_) >= lengthn(_base_p_) then do;
                            _is_v_slash_ = 0;
                            leave;
                        end;
                    end;
                    if _is_v_slash_ then do;
                        _st_ = scan(_posb_cl_, 2, ' /', 'm');
                        _en_ = scan(_posb_cl_, _n_slash_parts_, ' /', 'm');
                        clean_sertif_1 = put(input(_st_, 8.), z6.) || ' SD ' || put(input(_en_, 8.), z6.);
                    end;
                end;
            end;
        end;
    end;

    /* Fallback ke CLSDT_SERTF_NO jika polis_ori tidak menghasilkan sertif (equiv Python clean_sertif) */
    if missing(clean_sertif_1) or clean_sertif_1 = '' then do;
        _sertf_raw_ = strip(&clsdt_sertf_var.);
        if prxmatch('/\.0$/', _sertf_raw_) then
            _sertf_raw_ = substr(_sertf_raw_, 1, lengthn(_sertf_raw_) - 2);
        if not missing(_sertf_raw_) and not (_sertf_raw_ in ('', '#')) then do;
            if prxmatch('/^\d{1,6}$/', strip(_sertf_raw_)) then
                clean_sertif_1 = put(input(strip(_sertf_raw_), 8.), z6.);
        end;
    end;
    drop _rxso_sd1_ _rxso_sd2_ _p_osb_ _st_ _en_ _posb_cl_ _n_slash_parts_
         _base_p_ _is_v_slash_ _spi_ _seg_ _sertf_raw_;
%mend apply_extract_sertif_osbal;

%macro apply_clean_insured(in_var);
    array _cin_out{5} $500 clean_insured_1-clean_insured_5;
    do _i_ = 1 to 5; _cin_out{_i_} = ''; end;
    _ins_in_  = strip(&in_var);
    _cin_n_   = 0;

    if missing(_ins_in_) or _ins_in_ = '' then goto _insured_done_;

    /* Strip outer quotes */
    _ins_in_ = prxchange('s/^[\s"''""]+|[\s"''""]+$//o', -1, _ins_in_);

    /* Split berdasarkan AND/OR, C/Q, Q/Q, QQ, koma, slash */
    _ins_wrk_ = _ins_in_;
    _ins_wrk_ = prxchange('s/\bAND\s*\/\s*OR\b/|/oi', -1, _ins_wrk_);
    _ins_wrk_ = prxchange('s/\bC\s*\/\s*Q\b/|/oi',   -1, _ins_wrk_);
    _ins_wrk_ = prxchange('s/\bQ\s*\/\s*Q\b/|/oi',   -1, _ins_wrk_);
    _ins_wrk_ = prxchange('s/\bQ\.?Q\.?\b/|/oi',      -1, _ins_wrk_);
    _ins_wrk_ = prxchange('s/,/|/o', -1, _ins_wrk_);
    _ins_wrk_ = prxchange('s/\//|/o', -1, _ins_wrk_);

    _n_parts_ = countw(_ins_wrk_, '|', 'm');
    do _pi_ = 1 to _n_parts_;
        _part_ = strip(scan(_ins_wrk_, _pi_, '|', 'm'));
        _part_ = %normalize_spaces(_part_);

        /* Skip junk: panjang <= 2, atau pure junk words */
        if lengthn(_part_) <= 2 then continue;
        _part_up_ = upcase(strip(_part_));
        if _part_up_ in ('PT','CV','TBK','PERSERO','LTD','INC','LLC','PTE',
                         'AND','OR','THE','OF','AS','NON FOOD','DIV',
                         'BAPAK','BPK','IBU','NYONYA','NY','MR','MRS','MS',
                         'SDR','SAUDARA','SAUDARI') then continue;
        if prxmatch('/^[^a-zA-Z0-9]+$/', _part_) then continue;

        /* Apply clean_insured_name */
        _pclean_ = _part_;
        do _loop_ = 1 to 4;
            _pclean_ = prxchange('s/(?:\b|\.|,)\s*(?:PT|CV|TBK|PERSERO|\(PERSERO\)|LTD|PTE(?:\s+LTD)?|INC|LLC|UD|PD|NV|BV|GMBH|SDN\s+BHD|BHD)\b\.?\s*/ /oi', -1, _pclean_);
            _pclean_ = prxchange('s/\b(?:BAPAK|BPK|IBU|NYONYA|NY|MR|MRS|MS|SDR|SAUDARA|SAUDARI|S\.?KOM|S\.?E|S\.?T|S\.?H|S\.?SI|M\.?M|M\.?B\.?A|M\.?SI|M\.?T|DRS?|DRA?|IR|PROF|PH\.?D|HJ?)\b\.?\s*/ /oi', -1, _pclean_);
            _pclean_ = prxchange('s/^\s*(?:PT\.?|CV\.?|TBK\.?|\(PERSERO\)|\bPERSERO\b|LTD\.?|INC\.?|LLC\.?)\s*//oi', -1, _pclean_);
            _pclean_ = prxchange('s/\s*,?\s*\bLTD\.?\s*$//oi', -1, _pclean_);
            _pclean_ = prxchange('s/^\s*(?:[A-Z]\.){1,}[A-Z]?\s*//o', -1, _pclean_);
            _pclean_ = prxchange('s/\(\s*\)/ /o', -1, _pclean_);
            _pclean_ = prxchange('s/[\/\-]/ /o', -1, _pclean_);
            _pclean_ = %normalize_spaces(_pclean_);
            _pclean_ = prxchange('s/^[\s"''""\-\.,\/;:]+|[\s"''""\-\.,\/;:]+$//o', -1, _pclean_);
            _pclean_ = strip(_pclean_);
        end;

        if lengthn(_pclean_) <= 2 then continue;
        _pclean_up_ = upcase(strip(_pclean_));
        if _pclean_up_ in ('PT','CV','TBK','PERSERO','LTD','INC','LLC','PTE',
                           'AND','OR','THE','OF','AS') then continue;

        if _cin_n_ < 5 then do;
            _cin_n_ + 1; _cin_out{_cin_n_} = upcase(_pclean_);
        end;
    end;

    /* Fallback: jika semua token di-skip, proses seluruh string */
    if _cin_n_ = 0 then do;
        _fallback_ = _ins_in_;
        _fallback_ = prxchange('s/(?:\b|\.|,)\s*(?:PT|CV|TBK|PERSERO|\(PERSERO\)|LTD|PTE|INC|LLC)\b\.?\s*/ /oi', -1, _fallback_);
        _fallback_ = %normalize_spaces(_fallback_);
        _fallback_ = strip(_fallback_);
        if lengthn(_fallback_) > 2 then do; _cin_n_ = 1; clean_insured_1 = upcase(_fallback_); end;
    end;

    _insured_done_: ;
    drop _ins_in_ _ins_wrk_ _cin_n_ _part_ _part_up_ _pclean_ _pclean_up_
         _fallback_ _pi_ _n_parts_ _loop_ _i_;
%mend apply_clean_insured;

/* Macro ingesti tabel database PostgreSQL ke in-memory CAS menggunakan PROC FEDSQL / DATA step */
%macro load_db_table(cas_tbl, work_tbl, db_tbl);
    %local loaded;
    %let loaded = 0;

    %if %sysfunc(libref(casuser)) = 0 %then %do;
        proc casutil incaslib="casuser";
            droptable casdata="&cas_tbl." quiet;
        run;
        %if %symexist(_sessid_) %then %do;
            %put [INGEST] Memuat FSI_DB.&db_tbl. via PROC FEDSQL ke casuser.&cas_tbl. ...;
            proc fedsql;
                create table casuser.&cas_tbl. as
                select * from FSI_DB."&db_tbl.";
            quit;
            %if %sysfunc(exist(casuser.&cas_tbl.)) %then %do;
                data &work_tbl.;
                    set casuser.&cas_tbl.;
                run;
                %let loaded = 1;
            %end;
        %end;
    %end;

    %if &loaded. = 0 %then %do;
        %put [INGEST] Memuat FSI_DB.&db_tbl. via DATA Step ke &work_tbl. ...;
        data &work_tbl.;
            set FSI_DB."&db_tbl."n;
        run;
    %end;
%mend load_db_table;

/* --- 3. DATA CLEANING & CAS CACHING (FACUL, OSBAL, SUSPEND) --- */

%macro execute_staging_phase();
    %local cache_hit;
    %let cache_hit = 0;

    %if &USE_CAS_CACHE. = 1 and &FORCE_REFRESH. = 0 %then %do;
        %if %sysfunc(exist(casuser.aca_facul_clean)) and
            %sysfunc(exist(casuser.aca_osbal_clean)) and
            %sysfunc(exist(casuser.aca_suspend_clean)) %then %do;
            %let cache_hit = 1;
            %put [CACHE HIT] Using cached clean staging tables from in-memory CAS (casuser).;
            data work.facul_clean; set casuser.aca_facul_clean; run;
            data work.osbal_clean; set casuser.aca_osbal_clean; run;
            data work.suspend_clean; set casuser.aca_suspend_clean; run;
        %end;
    %end;

    %if &cache_hit. = 0 %then %do;
        %put [STAGING] Ingesting and cleaning raw data from FSI_DB...;

        /* Phase 1A: Clean Facul Data */
        %load_db_table(aca_facul_raw, work.facul_raw, &DB_TABLE_FACUL.);

        %macro filter_facul_step();
        data work.facul_filtered;
            set work.facul_raw;
            %if %varexist(work.facul_raw, COMP_NAME_1) %then %do;
                if missing(COMP_NAME2) then COMP_NAME2 = COMP_NAME_1;
            %end;
            _cedant_chk_ = strip(COMP_NAME);
            if _cedant_chk_ = "PT. ASURANSI CENTRAL ASIA" then output;
            drop _cedant_chk_;
        run;
        %mend filter_facul_step;
        %filter_facul_step();

        data work.facul_bp;
            set work.facul_filtered;
            rename FAC_POLICY_NO = polis_ori
                   FAC_SLIP      = slip_ori
                   FAC_INSURED   = insured_ori;
        run;

        data work.facul_cleaned;
            set work.facul_bp;
            length business_partner $500
                   clean_polis_1   - clean_polis_5   $500
                   clean_slip_1    - clean_slip_5    $500
                   clean_insured_1 - clean_insured_5 $500
                   clean_sertif_1  - clean_sertif_5  $10;

            _broker_ = strip(COMP_NAME2);
            if missing(_broker_) or upcase(_broker_) = 'DIRECT' then
                business_partner = strip(COMP_NAME);
            else
                business_partner = _broker_;
            drop _broker_;

            %apply_clean_polis(polis_ori);
            %apply_extract_sertif(polis_ori);
            %apply_clean_slip(slip_ori);
            %apply_clean_insured(insured_ori);

            array _str_cols{*} _character_;
            do _sci_ = 1 to dim(_str_cols);
                if not missing(_str_cols{_sci_}) then
                    %clean_control_chars(_str_cols{_sci_});
            end;
            drop _sci_;
        run;

        data work.facul_clean;
            length FAC_CODE $50 FAC_CEDANT $50 COMP_NAME $200 FAC_BROKER $50 COMP_NAME2 $200 BUSSINESS_PATNER $500
                   FAC_INSURED $500 INSURED_CLEAN_1 INSURED_CLEAN_2 INSURED_CLEAN_3 INSURED_CLEAN_4 INSURED_CLEAN_5 $500
                   FAC_POLICY_NO $500 POLICY_CLEAN_1 POLICY_CLEAN_2 POLICY_CLEAN_3 POLICY_CLEAN_4 POLICY_CLEAN_5 $500
                   SERTIF_CLEAN_1 SERTIF_CLEAN_2 SERTIF_CLEAN_3 $10
                   FAC_SLIP $500 SLIP_CLEAN_1 SLIP_CLEAN_2 SLIP_CLEAN_3 SLIP_CLEAN_4 SLIP_CLEAN_5 $500
                   FAC_CURRENCY $10 FAC_INP_DATE 8 FAC_COM_DATE 8 FAC_EXP_DATE 8
                   FAC_SUB_CLASS $50 FAC_RISK $50 FAC_DESC $500 FAC_ACC_STS $50 FAC_STS_SLIP $50;
            format FAC_INP_DATE FAC_COM_DATE FAC_EXP_DATE yymmdd10.;

            retain FAC_CODE FAC_CEDANT COMP_NAME FAC_BROKER COMP_NAME2 BUSSINESS_PATNER
                   FAC_INSURED INSURED_CLEAN_1 INSURED_CLEAN_2 INSURED_CLEAN_3 INSURED_CLEAN_4 INSURED_CLEAN_5
                   FAC_POLICY_NO POLICY_CLEAN_1 SERTIF_CLEAN_1 POLICY_CLEAN_2 SERTIF_CLEAN_2 POLICY_CLEAN_3 SERTIF_CLEAN_3 POLICY_CLEAN_4 POLICY_CLEAN_5
                   FAC_SLIP SLIP_CLEAN_1 SLIP_CLEAN_2 SLIP_CLEAN_3 SLIP_CLEAN_4 SLIP_CLEAN_5
                   FAC_CURRENCY FAC_INP_DATE FAC_COM_DATE FAC_EXP_DATE FAC_SUB_CLASS FAC_RISK FAC_DESC FAC_ACC_STS FAC_STS_SLIP;

            set work.facul_cleaned;

            BUSSINESS_PATNER = business_partner;
            FAC_INSURED      = insured_ori;
            INSURED_CLEAN_1  = clean_insured_1;
            INSURED_CLEAN_2  = clean_insured_2;
            INSURED_CLEAN_3  = clean_insured_3;
            INSURED_CLEAN_4  = clean_insured_4;
            INSURED_CLEAN_5  = clean_insured_5;

            FAC_POLICY_NO    = polis_ori;
            POLICY_CLEAN_1   = clean_polis_1;
            POLICY_CLEAN_2   = clean_polis_2;
            POLICY_CLEAN_3   = clean_polis_3;
            POLICY_CLEAN_4   = clean_polis_4;
            POLICY_CLEAN_5   = clean_polis_5;
            SERTIF_CLEAN_1   = clean_sertif_1;
            SERTIF_CLEAN_2   = clean_sertif_2;
            SERTIF_CLEAN_3   = clean_sertif_3;

            FAC_SLIP         = slip_ori;
            SLIP_CLEAN_1     = clean_slip_1;
            SLIP_CLEAN_2     = clean_slip_2;
            SLIP_CLEAN_3     = clean_slip_3;
            SLIP_CLEAN_4     = clean_slip_4;
            SLIP_CLEAN_5     = clean_slip_5;
        run;
        %put [Phase 1A] Facul Clean rows: %sysfunc(attrn(%sysfunc(open(work.facul_clean)), nobs));

        /* Phase 1B: Clean OSBAL Data */
        %load_db_table(aca_osbal_raw, work.osbal_raw, &DB_TABLE_OSBAL.);

        data work.osbal_filtered;
            set work.osbal_raw;
            _cedant_ = strip(CCOS_COMP_NAME);
            if _cedant_ in ('PT. ASURANSI CENTRAL ASIA', 'PT. ASURANSI CENTRAL ASIA SYARIAH') then output;
            drop _cedant_;
        run;

        data work.osbal_renamed;
            set work.osbal_filtered;
            rename FAC_POLICY_NO = polis_ori
                   FAC_SLIP      = slip_ori
                   FAC_INSURED   = insured_ori;
        run;

        %let HAS_CLSDT_POLIS = %eval(%varexist(work.osbal_renamed, CLSDT_POLICY_NO) > 0);
        %let HAS_CLSDT_SLIP  = %eval(%varexist(work.osbal_renamed, CLSDT_SLIP_NO)  > 0);
        %let HAS_CLSDT_SERTF = %eval(%varexist(work.osbal_renamed, CLSDT_SERTF_NO) > 0);

        %macro is_clsdt_valid(val, result_var);
            &result_var = 0;
            if not missing(&val.) and strip(&val.) ne '' then do;
                if prxmatch('/\d{10,}/', strip(&val.)) then &result_var = 1;
                else if prxmatch('/\d{7,}/', strip(&val.))
                     and not prxmatch('/\bTBA\b/oi', strip(&val.)) then &result_var = 1;
            end;
        %mend is_clsdt_valid;

        %macro run_osbal_clean_step();
        data work.osbal_cleaned;
            set work.osbal_renamed;

            length clean_polis_1   - clean_polis_5   $500
                   clean_slip_1    - clean_slip_5    $500
                   clean_insured_1 - clean_insured_5 $500
                   clean_sertif_1  - clean_sertif_5  $10;

            /* 1. Clean POLIS dengan prioritas CLSDT */
            %if &HAS_CLSDT_POLIS. %then %do;
                %apply_clean_polis(CLSDT_POLICY_NO, pfx=_cp_clsdt, tag=clsdt_p);
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

            /* 2. Clean SLIP dengan prioritas CLSDT */
            %if &HAS_CLSDT_SLIP. %then %do;
                %apply_clean_slip(CLSDT_SLIP_NO, pfx=_csl_clsdt, tag=clsdt_s);
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

            /* 3. Clean INSURED */
            %apply_clean_insured(insured_ori);

            /* 4. Clean SERTIFIKAT (equiv Python extract_cert_from_polis_osbal + clean_sertif) */
            %if &HAS_CLSDT_SERTF. %then %do;
                %apply_extract_sertif_osbal(polis_ori, CLSDT_SERTF_NO);
            %end;
            %else %do;
                _dummy_sertf_ = '';
                %apply_extract_sertif_osbal(polis_ori, _dummy_sertf_);
                drop _dummy_sertf_;
            %end;

            array _str_all{*} _character_;
            do _sci_ = 1 to dim(_str_all);
                if not missing(_str_all{_sci_}) then
                    %clean_control_chars(_str_all{_sci_});
            end;
            drop _sci_;
        run;
        %mend run_osbal_clean_step;
        %run_osbal_clean_step();

        data work.osbal_clean;
            length CCOS_DOC_NO $100 CCOS_DATE 8 CCOS_REF_CODE $100 CCOS_COMP $50 CCOS_COMP_NAME $200 CCOS_REF_COMP $50 CCOS_REF_COMP_NAME $200
                   FAC_INSURED $500 INSURED_CLEAN_1 INSURED_CLEAN_2 INSURED_CLEAN_3 INSURED_CLEAN_4 INSURED_CLEAN_5 $500
                   CCOS_CURR $10 CCOS_OR_BAL 8 CCOS_BAL_DUE 8 CCOS_OR_BAL_IN_IDR 8 CCOS_BAL_DUE_IN_IDR 8
                   FAC_COM_DATE 8 FAC_EXP_DATE 8 FAC_DUE_DATES $50 FAC_SUB_CLASS $50
                   FAC_POLICY_NO $500 POLICY_CLEAN_1 POLICY_CLEAN_2 POLICY_CLEAN_3 POLICY_CLEAN_4 POLICY_CLEAN_5 $500
                   SERTIF_CLEAN_1 SERTIF_CLEAN_2 SERTIF_CLEAN_3 $10
                   FAC_SLIP $500 SLIP_CLEAN_1 SLIP_CLEAN_2 SLIP_CLEAN_3 SLIP_CLEAN_4 SLIP_CLEAN_5 $500
                   CLASS_CODE $50 CLASS_NAME $200 CLSDT_POLICY_NO $500 CLSDT_SLIP_NO $500 CLSDT_SERTF_NO $500;
            format CCOS_DATE FAC_COM_DATE FAC_EXP_DATE yymmdd10.;

            retain CCOS_DOC_NO CCOS_DATE CCOS_REF_CODE CCOS_COMP CCOS_COMP_NAME CCOS_REF_COMP CCOS_REF_COMP_NAME
                   FAC_INSURED INSURED_CLEAN_1 INSURED_CLEAN_2 INSURED_CLEAN_3 INSURED_CLEAN_4 INSURED_CLEAN_5
                   CCOS_CURR CCOS_OR_BAL CCOS_BAL_DUE CCOS_OR_BAL_IN_IDR CCOS_BAL_DUE_IN_IDR
                   FAC_COM_DATE FAC_EXP_DATE FAC_DUE_DATES FAC_SUB_CLASS
                   FAC_POLICY_NO POLICY_CLEAN_1 SERTIF_CLEAN_1 POLICY_CLEAN_2 SERTIF_CLEAN_2 POLICY_CLEAN_3 SERTIF_CLEAN_3 POLICY_CLEAN_4 POLICY_CLEAN_5
                   FAC_SLIP SLIP_CLEAN_1 SLIP_CLEAN_2 SLIP_CLEAN_3 SLIP_CLEAN_4 SLIP_CLEAN_5
                   CLASS_CODE CLASS_NAME CLSDT_POLICY_NO CLSDT_SLIP_NO CLSDT_SERTF_NO;

            set work.osbal_cleaned;

            FAC_INSURED      = insured_ori;
            INSURED_CLEAN_1  = clean_insured_1;
            INSURED_CLEAN_2  = clean_insured_2;
            INSURED_CLEAN_3  = clean_insured_3;
            INSURED_CLEAN_4  = clean_insured_4;
            INSURED_CLEAN_5  = clean_insured_5;

            FAC_POLICY_NO    = polis_ori;
            POLICY_CLEAN_1   = clean_polis_1;
            POLICY_CLEAN_2   = clean_polis_2;
            POLICY_CLEAN_3   = clean_polis_3;
            POLICY_CLEAN_4   = clean_polis_4;
            POLICY_CLEAN_5   = clean_polis_5;
            SERTIF_CLEAN_1   = clean_sertif_1;
            SERTIF_CLEAN_2   = clean_sertif_2;
            SERTIF_CLEAN_3   = clean_sertif_3;

            FAC_SLIP         = slip_ori;
            SLIP_CLEAN_1     = clean_slip_1;
            SLIP_CLEAN_2     = clean_slip_2;
            SLIP_CLEAN_3     = clean_slip_3;
            SLIP_CLEAN_4     = clean_slip_4;
            SLIP_CLEAN_5     = clean_slip_5;
        run;
        %put [Phase 1B] OSBAL Clean rows: %sysfunc(attrn(%sysfunc(open(work.osbal_clean)), nobs));

        /* Phase 1C: Clean Suspend Data */
        %load_db_table(aca_suspend_raw, work.suspend_raw, &DB_TABLE_SUSPEND.);

        data work.suspend_filtered;
            set work.suspend_raw;
            _ced_chk_ = strip("CEDANT SHRT NAME"n);
            _sts_chk_ = upcase(strip(STATUS));
            if _ced_chk_ = "CENTRAL" and _sts_chk_ = 'SUSPENSE' then output;
            drop _ced_chk_ _sts_chk_;
        run;

        data work.suspend_renamed;
            set work.suspend_filtered;
            rename INSURED = insured_ori
                   POLIS   = polis_ori
                   "SLIP NO"n = slip_ori;
        run;

        %macro apply_clean_polis_suspend(in_var);
            length clean_polis_1 - clean_polis_5 $500;
            clean_polis_1 = '';
            clean_polis_2 = '';
            clean_polis_3 = '';
            clean_polis_4 = '';
            clean_polis_5 = '';

            _polis_sus_ = strip(&in_var.);
            if not missing(_polis_sus_) and not (_polis_sus_ in ('', '0', '-')) then do;
                _prev_sus_ = '';
                do while (_polis_sus_ ne _prev_sus_);
                    _prev_sus_ = _polis_sus_;
                    _polis_sus_ = prxchange('s/-(?:\d+(?:\/\d+)?|EXT\(\d+\))$//oi', -1, strip(_polis_sus_));
                    _polis_sus_ = strip(_polis_sus_);
                end;
                if _polis_sus_ ne '' then clean_polis_1 = _polis_sus_;
            end;
            drop _polis_sus_ _prev_sus_;
        %mend apply_clean_polis_suspend;

        %macro apply_clean_insured_suspend(in_var, polis_in, slip_in);
            length clean_insured_1 - clean_insured_5 $500;
            array _cis_out{5} $500 clean_insured_1-clean_insured_5;
            do _i_ = 1 to 5; _cis_out{_i_} = ''; end;
            _ins_sus_ = strip(&in_var.);
            _cis_n_   = 0;

            if missing(_ins_sus_) or _ins_sus_ in ('', '0', '-') then goto _insus_done_;

            if not missing(&polis_in.) then do;
                _praw_ = strip(put(&polis_in., $500.));
                if _praw_ ne '' and _praw_ ne '-' then
                    _ins_sus_ = tranwrd(_ins_sus_, _praw_, '');
                if clean_polis_1 ne '' and clean_polis_1 ne '-' then
                    _ins_sus_ = tranwrd(_ins_sus_, clean_polis_1, '');
            end;

            if not missing(&slip_in.) then do;
                _sraw_ = strip(put(&slip_in., $500.));
                if _sraw_ ne '' and _sraw_ ne '-' then
                    _ins_sus_ = tranwrd(_ins_sus_, _sraw_, '');
                if clean_slip_1 ne '' and clean_slip_1 ne '-' then
                    _ins_sus_ = tranwrd(_ins_sus_, clean_slip_1, '');
            end;

            _ins_sus_ = prxchange('s/\(\s*[\d\.\/\-]+\s*\)//o', -1, _ins_sus_);
            _ins_sus_ = prxchange('s/\b(?:AND|AN|OR)\s*\/\s*(?:AND|OR)\b/,/oi', -1, _ins_sus_);
            _ins_sus_ = prxchange('s/\bAND\s+OR\b/,/oi', -1, _ins_sus_);
            _ins_sus_ = prxchange('s/\bCO\.,?\s*LTD\.?\b/,/oi', -1, _ins_sus_);
            _ins_sus_ = prxchange('s/\bTBK\.?\b//oi',   -1, _ins_sus_);
            _ins_sus_ = prxchange('s/\(PERSERO\)//oi',   -1, _ins_sus_);
            _ins_sus_ = prxchange('s/\bPERSERO\b//oi',   -1, _ins_sus_);
            _ins_sus_ = prxchange('s/\bLTD\.?\b//oi',    -1, _ins_sus_);

            _ins_sus_ = prxchange('s/\bQQ\b/|/oi', -1, _ins_sus_);
            _ins_sus_ = prxchange('s/\// |/o', -1, _ins_sus_);
            _ins_sus_ = prxchange('s/,/|/o', -1, _ins_sus_);
            _ins_sus_ = prxchange('s/-(?!\s*(?:19|20)\d{2}\b)/|/o', -1, _ins_sus_);
            _ins_sus_ = prxchange('s/\d+\./|/o', -1, _ins_sus_);

            _ins_sus_ = %normalize_spaces(_ins_sus_);
            _n_p_ = countw(_ins_sus_, '|', 'm');
            do _k_ = 1 to _n_p_;
                _part_ = strip(scan(_ins_sus_, _k_, '|', 'm'));
                _part_ = %normalize_spaces(_part_);
                _part_ = prxchange('s/^\s*(?:AND|OR|\/|-)\s*//oi', -1, _part_);
                _part_ = prxchange('s/\s*(?:AND|OR|\/|-)\s*$//oi', -1, _part_);
                _part_ = strip(_part_);
                if lengthn(_part_) > 2 and _cis_n_ < 5 then do;
                    _cis_n_ + 1;
                    _cis_out{_cis_n_} = upcase(_part_);
                end;
            end;

            if _cis_n_ = 0 and lengthn(_ins_sus_) > 2 then do;
                _cis_n_ = 1; clean_insured_1 = upcase(strip(_ins_sus_));
            end;

            _insus_done_: ;
            drop _ins_sus_ _cis_n_ _praw_ _sraw_ _n_p_ _k_ _part_;
        %mend apply_clean_insured_suspend;

        data work.suspend_cleaned;
            set work.suspend_renamed;

            length clean_polis_1   - clean_polis_5   $500
                   clean_slip_1    - clean_slip_5    $500
                   clean_insured_1 - clean_insured_5 $500
                   clean_sertif_1  - clean_sertif_5  $10;

            %apply_clean_polis_suspend(polis_ori);
            %apply_extract_sertif(polis_ori);

            clean_slip_1 = strip(slip_ori);
            clean_slip_2 = ''; clean_slip_3 = ''; clean_slip_4 = ''; clean_slip_5 = '';

            %apply_clean_insured_suspend(insured_ori, polis_ori, slip_ori);

            array _str_s{*} _character_;
            do _sci_ = 1 to dim(_str_s);
                if not missing(_str_s{_sci_}) then
                    %clean_control_chars(_str_s{_sci_});
            end;
            drop _sci_;
        run;

        data work.suspend_clean;
            retain "CEDANT SHRT NAME"n STATUS
                   insured_ori clean_insured_1-clean_insured_5
                   polis_ori clean_polis_1-clean_polis_5 clean_sertif_1-clean_sertif_5
                   slip_ori clean_slip_1-clean_slip_5;
            set work.suspend_cleaned;
        run;
        %put [Phase 1C] Suspend Clean rows: %sysfunc(attrn(%sysfunc(open(work.suspend_clean)), nobs));

        /* Cache staging tables into in-memory CAS casuser */
        %if %sysfunc(libref(casuser)) = 0 %then %do;
            proc casutil incaslib="casuser"; droptable casdata="aca_facul_clean" quiet; run;
            proc casutil incaslib="casuser"; droptable casdata="aca_osbal_clean" quiet; run;
            proc casutil incaslib="casuser"; droptable casdata="aca_suspend_clean" quiet; run;
            data casuser.aca_facul_clean; set work.facul_clean; run;
            data casuser.aca_osbal_clean; set work.osbal_clean; run;
            data casuser.aca_suspend_clean; set work.suspend_clean; run;
            %put [CACHE SAVED] Clean staging tables saved to in-memory CAS casuser.;
        %end;
    %end;
%mend execute_staging_phase;
%execute_staging_phase();

/* --- 4. MULTI-PASS MATCHING ENGINE --- */

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
    _com_date_ = input(put(FAC_COM_DATE, $100.), ?? anydtdte10.);
run;

data work.suspend_clean;
    set work.suspend_clean;
    sus_idx = _n_;
    format _sus_date_ date9. _sus_date_ls_ date9.;
    _sus_date_ = input(put("RECEIPT DATE"n, $100.), ?? anydtdte10.);

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
            _bl_ = input(put(upcase(strip(put(BULAN, $20.))), $bulan_fmt.), 8.);
            _th_ = input(strip(put(TAHUN, $20.)), 8.);
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
        AMOUNT_ORI_NUM = input(_amt_str_, ?? best32.);
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
        AMOUNT_ORI_NUM = input(_amt_str_, ?? best32.);
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

/* --- 6. PERSISTENCE & SYNCHRONIZATION (CAS TARGET LIB & POSTGRESQL FSI_DB) --- */

%macro sync_to_cas();
    %if %sysfunc(libref(casuser)) = 0 %then %do;
        %put [CAS] Promoting 5 production tables to CASLIB &CAS_TARGET_LIB. ...;

        proc casutil incaslib="casuser"; droptable casdata="ACA_OSBAL_CLEAN" quiet; run;
        proc casutil incaslib="casuser"; droptable casdata="ACA_FACUL_CLEAN" quiet; run;
        proc casutil incaslib="casuser"; droptable casdata="ACA_SUSPEND_CLEAN" quiet; run;
        %if %sysfunc(exist(work.osbal_clean)) %then %do;
            proc casutil incaslib="casuser"; droptable casdata="ACA_OSBAL_CLEAN" quiet; run;
            data casuser.ACA_OSBAL_CLEAN; set work.osbal_clean; run;
            proc cas;
                table.dropTable / caslib="&CAS_TARGET_LIB." name="ACA_OSBAL_CLEAN" quiet=true;
                table.promote   / caslib="casuser" name="ACA_OSBAL_CLEAN" targetLib="&CAS_TARGET_LIB.";
                table.save      / caslib="&CAS_TARGET_LIB." table="ACA_OSBAL_CLEAN" name="ACA_OSBAL_CLEAN.sashdat" replace=true;
            quit;
            data casuser.ACA_OSBAL_CLEAN; set work.osbal_clean; run;
        %end;

        %if %sysfunc(exist(work.facul_clean)) %then %do;
            proc casutil incaslib="casuser"; droptable casdata="ACA_FACUL_CLEAN" quiet; run;
            data casuser.ACA_FACUL_CLEAN; set work.facul_clean; run;
            proc cas;
                table.dropTable / caslib="&CAS_TARGET_LIB." name="ACA_FACUL_CLEAN" quiet=true;
                table.promote   / caslib="casuser" name="ACA_FACUL_CLEAN" targetLib="&CAS_TARGET_LIB.";
                table.save      / caslib="&CAS_TARGET_LIB." table="ACA_FACUL_CLEAN" name="ACA_FACUL_CLEAN.sashdat" replace=true;
            quit;
            data casuser.ACA_FACUL_CLEAN; set work.facul_clean; run;
        %end;

        %if %sysfunc(exist(work.suspend_clean)) %then %do;
            proc casutil incaslib="casuser"; droptable casdata="ACA_SUSPEND_CLEAN" quiet; run;
            data casuser.ACA_SUSPEND_CLEAN; set work.suspend_clean; run;
            proc cas;
                table.dropTable / caslib="&CAS_TARGET_LIB." name="ACA_SUSPEND_CLEAN" quiet=true;
                table.promote   / caslib="casuser" name="ACA_SUSPEND_CLEAN" targetLib="&CAS_TARGET_LIB.";
                table.save      / caslib="&CAS_TARGET_LIB." table="ACA_SUSPEND_CLEAN" name="ACA_SUSPEND_CLEAN.sashdat" replace=true;
            quit;
            data casuser.ACA_SUSPEND_CLEAN; set work.suspend_clean; run;
        %end;

        %if %sysfunc(exist(work.aca_matching_v1)) %then %do;
            proc casutil incaslib="casuser"; droptable casdata="ACA_MATCHING_V1" quiet; run;
            data casuser.ACA_MATCHING_V1; set work.aca_matching_v1; run;
            proc cas;
                table.dropTable / caslib="&CAS_TARGET_LIB." name="ACA_MATCHING_V1" quiet=true;
                table.promote   / caslib="casuser" name="ACA_MATCHING_V1" targetLib="&CAS_TARGET_LIB.";
                table.save      / caslib="&CAS_TARGET_LIB." table="ACA_MATCHING_V1" name="ACA_MATCHING_V1.sashdat" replace=true;
            quit;
            data casuser.ACA_MATCHING_V1; set work.aca_matching_v1; run;
            data casuser.SUSPENSE_DATA_SUSPENSE_V1; set work.aca_matching_v1; run;
        %end;

        %if %sysfunc(exist(work.aca_matching_v2)) %then %do;
            proc casutil incaslib="casuser"; droptable casdata="ACA_MATCHING_V2" quiet; run;
            data casuser.ACA_MATCHING_V2; set work.aca_matching_v2; run;
            proc cas;
                table.dropTable / caslib="&CAS_TARGET_LIB." name="ACA_MATCHING_V2" quiet=true;
                table.promote   / caslib="casuser" name="ACA_MATCHING_V2" targetLib="&CAS_TARGET_LIB.";
                table.save      / caslib="&CAS_TARGET_LIB." table="ACA_MATCHING_V2" name="ACA_MATCHING_V2.sashdat" replace=true;
            quit;
            data casuser.ACA_MATCHING_V2; set work.aca_matching_v2; run;
            data casuser.SUSPENSE_DATA_SUSPENSE_V2; set work.aca_matching_v2; run;
        %end;

        %put [CAS] Production tables successfully promoted to &CAS_TARGET_LIB. and cached in casuser!;
    %end;
    %else %do;
        %put [CAS WARNING] CAS session inactive. CAS persistence skipped.;
    %end;
%mend sync_to_cas;
%sync_to_cas();

%macro sync_to_postgresql();
    %local fsidb_ok;
    %let fsidb_ok = 0;

    proc sql noprint;
        select count(*) into :fsidb_ok trimmed
        from dictionary.libnames
        where upcase(libname) = 'FSI_DB';
    quit;

    %if &fsidb_ok. > 0 %then %do;
        %put [POSTGRESQL] Syncing tables to database FSI_DB (100% Identik Python Pipeline)...;

        /* Idempotent Deletes */
        proc sql;
            /* 1. Clean OSBAL */
            delete from FSI_DB."SUSPENSE_DATA_OSBAL_CLEAN"n
            where "CCOS_COMP_NAME"n in ('PT. ASURANSI CENTRAL ASIA', 'PT. ASURANSI CENTRAL ASIA SYARIAH');

            delete from FSI_DB."SUSPENSE_OSBAL_CLEAN_2026"n
            where "CCOS_COMP_NAME"n in ('PT. ASURANSI CENTRAL ASIA', 'PT. ASURANSI CENTRAL ASIA SYARIAH');

            /* 2. Clean Facul */
            delete from FSI_DB."SUSPENSE_DATA_FACUL_CLEAN"n
            where "COMP_NAME"n = 'PT. ASURANSI CENTRAL ASIA';

            delete from FSI_DB."SUSPENSE_FACUL_CLEAN_2026"n
            where "COMP_NAME"n = 'PT. ASURANSI CENTRAL ASIA';

            /* 3. Suspense V1 */
            delete from FSI_DB."SUSPENSE_DATA_SUSPENSE_V1"n
            where "CEDANT NAME"n = 'PT. ASURANSI CENTRAL ASIA' or "CEDANT SHRT NAME"n = 'CENTRAL';

            delete from FSI_DB."SUSPENSE_CLEAN_2026"n
            where "CEDANT NAME"n = 'PT. ASURANSI CENTRAL ASIA' or "CEDANT SHRT NAME"n = 'CENTRAL';

            /* 4. Suspense V2 */
            delete from FSI_DB."SUSPENSE_DATA_SUSPENSE_V2"n
            where "CEDANT NAME"n = 'PT. ASURANSI CENTRAL ASIA' or "CEDANT SHRT NAME"n = 'CENTRAL';
        quit;

        /* Appends (sama persis dengan Python) */
        proc append base=FSI_DB."SUSPENSE_DATA_OSBAL_CLEAN"n data=work.osbal_clean force; run;
        proc append base=FSI_DB."SUSPENSE_OSBAL_CLEAN_2026"n data=work.osbal_clean force; run;

        proc append base=FSI_DB."SUSPENSE_DATA_FACUL_CLEAN"n data=work.facul_clean force; run;
        proc append base=FSI_DB."SUSPENSE_FACUL_CLEAN_2026"n data=work.facul_clean force; run;

        proc append base=FSI_DB."SUSPENSE_DATA_SUSPENSE_V1"n data=work.aca_matching_v1 force; run;
        proc append base=FSI_DB."SUSPENSE_CLEAN_2026"n data=work.aca_matching_v1 force; run;

        proc append base=FSI_DB."SUSPENSE_DATA_SUSPENSE_V2"n data=work.aca_matching_v2 force; run;

        %put [POSTGRESQL] All 5 production tables synced successfully to PostgreSQL FSI_DB!;
    %end;
    %else %do;
        %put [POSTGRESQL WARNING] Libname FSI_DB not connected. Database sync skipped.;
    %end;
%mend sync_to_postgresql;
%sync_to_postgresql();

/* --- 7. EXECUTIVE SUMMARY & RECONCILIATION BREAKDOWN --- */

proc freq data=work.aca_matching_v1;
    tables FLAG_PROD * SKENARIO / nocum nopercent;
    title "ACA Suspense Reconciliation - Output V1 Summary (Row-Level / KEEP_ROWS)";
run;

proc freq data=work.aca_matching_v2;
    tables FLAG_PROD * SKENARIO / nocum nopercent;
    title "ACA Suspense Reconciliation - Output V2 Summary (Aggregated / MERGE_ROWS)";
run;
title;

%put [COMPLETE] =========================================================================;
%put [COMPLETE] ACA Reconciliation Pipeline execution completed successfully!;
%put [COMPLETE] In-Memory Tables : &CAS_TARGET_LIB..ACA_OSBAL_CLEAN;
%put [COMPLETE]                    &CAS_TARGET_LIB..ACA_FACUL_CLEAN;
%put [COMPLETE]                    &CAS_TARGET_LIB..ACA_SUSPEND_CLEAN;
%put [COMPLETE]                    &CAS_TARGET_LIB..ACA_MATCHING_V1;
%put [COMPLETE]                    &CAS_TARGET_LIB..ACA_MATCHING_V2;
%put [COMPLETE] Database Tables  : FSI_DB.SUSPENSE_DATA_OSBAL_CLEAN;
%put [COMPLETE]                    FSI_DB.SUSPENSE_DATA_FACUL_CLEAN;
%put [COMPLETE]                    FSI_DB.SUSPENSE_DATA_SUSPENSE_V1;
%put [COMPLETE]                    FSI_DB.SUSPENSE_CLEAN_2026;
%put [COMPLETE]                    FSI_DB.SUSPENSE_DATA_SUSPENSE_V2;
%put [COMPLETE] =========================================================================;
