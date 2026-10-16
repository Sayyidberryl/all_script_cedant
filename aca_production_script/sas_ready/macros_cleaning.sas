/* =============================================================================
   MACROS_CLEANING.SAS
   Shared macro library untuk semua cleaning script (facul, osbal, suspend).
   Setara dengan fungsi helper bersama di Python (normalize, token, regex patterns).

   CARA PAKAI:
     %include "path/to/macros_cleaning.sas";

   SAS Viya / SAS 9.4 compatible.
   ============================================================================= */


/* =============================================================================
   SECTION 1: COMPILE REGEX PATTERNS (sekali saja, di-reuse via PRXMATCH/PRXCHANGE)
   Setara dengan: re.compile(...) di level modul Python
   ============================================================================= */

%macro compile_cleaning_regex();
    /* --- POLIS exception patterns --- */
    %global RX_POLIS_EXCEPTION RX_POLIS_EXCEPTION_ID;
    %let RX_POLIS_EXCEPTION = /MOP\s*MARINE|(?:LINE\s*SLIP|LINESLIP)|\b(?:P1|P2|P3|P73)\s*CANCEL|\b(?:P1|P2|P3|P73)\b|\bCANCEL\b|PENYELESAIAN|HUTANG\s*PIUTANG/oi;

    /* --- SLIP exception patterns --- */
    %global RX_SLIP_EXCEPTION;
    %let RX_SLIP_EXCEPTION = /\bSUMMARY\b|\bBORDER[OA]\b|\bBORDRO\b|\bSINGGLESHIPMENT\b/oi;

    /* --- Normalize double-space --- */
    %global RX_MULTI_SPACE;
    %let RX_MULTI_SPACE = /\s{2,}/o;

    /* --- SLIP noise words (bulan, tahun, mata uang, dsb) --- */
    %global RX_SLIP_NOISE;
    %let RX_SLIP_NOISE = /\b(?:JANUARI|FEBRUARI|MARET|APRIL|MEI|JUNI|JULI|AGUSTUS|SEPTEMBER|OKTOBER|NOVEMBER|DESEMBER)\b|\b(?:JANUARY|FEBRUARY|MARCH|APRIL|MAY|JUNE|JULY|AUGUST|SEPTEMBER|OCTOBER|NOVEMBER|DECEMBER)\b|\b(?:IDR|USD|ENG)\b|\b(?:P1|P2|P3|P73)\b|\bNEW\b|\bVARIOUS\b|\b20[0-9]{2}\b/oi;

    /* --- SLIP valid token: 7+ alfanumerik --- */
    %global RX_SLIP_TOKEN;
    %let RX_SLIP_TOKEN = /[A-Z0-9][A-Z0-9\-]{6,}/oi;

    /* --- INSURED split pattern --- */
    %global RX_INSURED_SPLIT;
    %let RX_INSURED_SPLIT = /\bAND\s*\/\s*OR\b|\bC\s*\/\s*Q\b|\bQ\s*\/\s*Q\b|\bQ\.?Q\.?\b|,(?![^(]*\))|\/(?![^(]*\))/oi;

    /* --- INSURED prefix/suffix legal entity --- */
    %global RX_INSURED_PREFIX;
    %let RX_INSURED_PREFIX = /^\s*(?:PT\.?|CV\.?|TBK\.?|\(PERSERO\)|\bPERSERO\b|LTD\.?|INC\.?|LLC\.?|UD\.?|PD\.?|NV\.?|BV\.?|GMBH\.?|SDN\s+BHD|BHD\.?)\s*|\s*(?:,?\s*\bTBK\b\s*(?:,?\s*PT\.?)?|,?\s*\bPT\.?\s*$|,?\s*\bCV\.?\s*$|\bPT\.?\b\s*$|\bCV\.?\b\s*$|,?\s*\bLTD\.?\s*$)|^\s*\((?:PERSERO|TBK)\)\s*|^\s*(?:[A-Z]\.){1,}[A-Z]?\s*|\s*,?\s*\b(?:S\.?KOM|S\.?E|S\.?T|S\.?H|S\.?SI|S\.?SOS|S\.?IP|S\.?TP|S\.?PSI|S\.?KED|M\.?M|M\.?B\.?A|M\.?SI|M\.?T|M\.?KN|M\.?H|DRS?|DRA?|IR|PROF|PH\.?D)\.?\b\s*$/oi;

    /* --- INSURED entity anywhere --- */
    %global RX_ENTITY_ANYWHERE;
    %let RX_ENTITY_ANYWHERE = /(?:\b|\.|,)\s*(?:PT|CV|TBK|PERSERO|\(PERSERO\)|LTD|PTE(?:\s+LTD)?|INC|LLC|UD|PD|NV|BV|GMBH|SDN\s+BHD|BHD)\b\.?\s*/oi;

    /* --- INSURED honorifics dan gelar --- */
    %global RX_HONORIFICS;
    %let RX_HONORIFICS = /\b(?:BAPAK|BPK|IBU|NYONYA|NY|MR|MRS|MS|SDR|SAUDARA|SAUDARI|S\.?KOM|S\.?E|S\.?T|S\.?H|S\.?SI|S\.?SOS|S\.?IP|S\.?TP|S\.?PSI|S\.?KED|M\.?M|M\.?B\.?A|M\.?SI|M\.?T|M\.?KN|M\.?H|DRS?|DRA?|IR|PROF|PH\.?D|HJ?)\b\.?\s*/oi;

    /* --- Sertifikat: S/D range --- */
    %global RX_CERT_SD;
    %let RX_CERT_SD = /(\d{1,6})\s*S\/D\s*(\d{1,6})/oi;

    /* --- Sertifikat: suffix digit setelah dash --- */
    %global RX_CERT_FROM_POLIS;
    %let RX_CERT_FROM_POLIS = /[-\s]+\s*(\d{1,6})(?:[^0-9]|$)/o;

    /* --- Sertifikat valid: 1-6 digit numerik --- */
    %global RX_SERTIF_DIGIT;
    %let RX_SERTIF_DIGIT = /^\d{1,6}$/o;

    /* --- Polis base strip: trailing suffix digit setelah dash --- */
    %global RX_POLIS_BASE_STRIP;
    %let RX_POLIS_BASE_STRIP = /^(.+?)-(\d{1,6})$/o;

    /* --- Various / TBA patterns --- */
    %global RX_VARIOUS RX_TBA;
    %let RX_VARIOUS = /\bVARIOUS\b/oi;
    %let RX_TBA     = /\bTBA\b/oi;

%mend compile_cleaning_regex;


options validvarname=any validmemname=extend;

/* =============================================================================
   SECTION 2: UTILITY MACROS
   ============================================================================= */

/* Cek keberadaan variabel di dataset
   Returns: >0 jika variabel ada, 0 jika tidak ada */
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

/* Normalize spaces: collapse multiple whitespace -> single space, strip
   Setara dengan: _normalize_spaces(text) di Python                            */
%macro normalize_spaces(val);
    strip(prxchange('s/\s{2,}/ /', -1, strip(&val)))
%mend normalize_spaces;


/* Normalize uppercase + strip + collapse whitespace
   Setara dengan: _normalize(value) di Python                                   */
%macro normalize(val);
    strip(prxchange('s/\s+/ /', -1, upcase(strip(put(&val, $500.)))))
%mend normalize;


/* Apakah token valid sebagai nomor polis?
   Setara dengan: _is_valid_polis_token(tok)
   True  = panjang >= 5, ada digit, tidak berakhir TBA
   Output: 1 (true) atau 0 (false) -- gunakan dalam IF                          */
%macro is_valid_polis_token(tok);
    (lengthn(strip(&tok)) >= 5
     and anydigit(strip(&tok)) > 0
     and not prxmatch('/TBA$/oi', strip(&tok)))
%mend is_valid_polis_token;


/* Apakah token valid sebagai nomor slip?
   Setara dengan: _is_valid_slip_token(tok)
   True  = panjang >= 7, ada digit, bukan 4-digit tahun                        */
%macro is_valid_slip_token(tok);
    (lengthn(strip(&tok)) >= 7 and anydigit(strip(&tok)) > 0)
%mend is_valid_slip_token;


/* Strip suffix digit pendek setelah dash dari nomor polis base (equiv Python: _strip_polis_base)
   Contoh: '100030817120000016-000149' -> '100030817120000016'                   */
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


/* =============================================================================
   SECTION 3: PROC FORMAT untuk mapping bulan Indonesia -> nomor bulan
   Setara dengan: _BULAN_MAP dict di Python
   Panggil SEKALI saja di awal program.
   ============================================================================= */
%macro define_bulan_format();
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
%mend define_bulan_format;


/* =============================================================================
   SECTION 4: CLEAN POLIS (digunakan oleh facul + osbal -- logika identik)
   Setara dengan: clean_polis(val) di Python

   SAS tidak mendukung fungsi return list, jadi output disimpan ke kolom
   clean_polis_1 s/d clean_polis_5. Panggil macro ini di dalam DATA step.

   Input : _polis_in  (string mentah dari kolom polis_ori)
   Output: clean_polis_1 ... clean_polis_5 (max 5 token)
   ============================================================================= */
/* ===========================================================================
   HELPER: %_extract_polis_tokens_m(text_var)
   Splits by 2+ spaces, then by +/whitespace; validates & strips each token.
   Writes into _cp_out{} / _cp_n_ (must be in scope from caller).
   =========================================================================== */
/* ===========================================================================
   HELPER: %_extract_polis_tokens_m(text_var, pfx=clean_polis, tag=cp)
   Splits by 2+ spaces, then by +/whitespace; validates & strips each token.
   Writes into _&tag._out{} / _cp_n_&tag._ (must be in scope from caller).
   =========================================================================== */
%macro _extract_polis_tokens_m(text_var, pfx=clean_polis, tag=cp);
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
%mend _extract_polis_tokens_m;

/* ===========================================================================
   HELPER: %_expand_plus_suffix_m(text_var, pfx=clean_polis, tag=cp)
   Expands BASE+s1+s2 where BASE is pure digits >=10 and si replaces last digits.
   Sets _eps_ok_=1 and writes into _&tag._out{}/_cp_n_&tag._ if expansion succeeds.
   =========================================================================== */
%macro _expand_plus_suffix_m(text_var, pfx=clean_polis, tag=cp);
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
%mend _expand_plus_suffix_m;

/* ===========================================================================
   apply_clean_polis: 100% faithful to Python cleaning_facul.clean_polis()
   All scenarios: exception, digit+TBA+digit, various/tba-only, dot-suffix,
   S/D base, TBA+content, &, + expansion, double-space multi-token, multi-slash
   with all-short check, single fallback.
   PERF: retain _rxcp_dot_ & _rxcp_sd_ - prxparse() runs ONCE per DATA step.
   Supports custom pfx & tag so multiple calls in one DATA step do not collide.
   =========================================================================== */
%macro apply_clean_polis(in_var, pfx=clean_polis, tag=cp);
    retain _rxcp_dot_ _rxcp_sd_ 0;
    array _&tag._out{5} $500 &pfx._1-&pfx._5;
    do _i_&tag._ = 1 to 5; _&tag._out{_i_&tag._} = ''; end;
    _polis_wrk_ = strip(&in_var);
    _cp_n_&tag._ = 0;

    if missing(_polis_wrk_) or _polis_wrk_ in ('', '0', '-') then goto _pdone_&tag.;

    _polis_up_ = upcase(strip(_polis_wrk_));

    /* 1. POLIS_EXCEPTION_RE -> return normalize(val) */
    if prxmatch('/MOP\s*MARINE|LINE\s*SLIP|LINESLIP|\bP[1-3]\b|\bP73\b|\bCANCEL\b|PENYELESAIAN|HUTANG\s*PIUTANG/oi', _polis_wrk_) then do;
        _cp_n_&tag._ = 1; &pfx._1 = %normalize_spaces(_polis_wrk_); goto _pdone_&tag.;
    end;

    /* 2. digit+TBA+digit -> return normalize(val) */
    if prxmatch('/\d+TBA\d+/oi', _polis_wrk_) then do;
        _cp_n_&tag._ = 1; &pfx._1 = %normalize_spaces(_polis_wrk_); goto _pdone_&tag.;
    end;

    /* 3. VARIOUS / TBA only -> return as-is */
    if prxmatch('/^VARIOUS\s*$/oi', _polis_up_)
       or prxmatch('/^VARIOUS\s*-\s*SEE\s+ATTACH/oi', _polis_up_)
       or prxmatch('/^TBA\s*$/oi', _polis_up_)
    then do;
        _cp_n_&tag._ = 1; &pfx._1 = strip(_polis_wrk_); goto _pdone_&tag.;
    end;

    if _rxcp_dot_ = 0 then _rxcp_dot_ = prxparse('/^(\d[\d\-]+)\.\d{1,6}$/');
    if _rxcp_sd_  = 0 then _rxcp_sd_  = prxparse('/^(\d[\d\-]+?)[-]?\d+S\/D\d+/oi');

    /* 4. Dot-suffix */
    if prxmatch('/^\d[\d\-]+\.\d{1,6}$/o', _polis_wrk_) then do;
        if prxmatch(_rxcp_dot_, _polis_wrk_) then do;
            _base_dot_ = prxposn(_rxcp_dot_, 1, _polis_wrk_);
            if %is_valid_polis_token(_base_dot_) then do;
                _cp_n_&tag._ = 1; &pfx._1 = strip(_base_dot_); goto _pdone_&tag.;
            end;
        end;
    end;

    /* 5. S/D range */
    if prxmatch('/S\/D/oi', _polis_wrk_) then do;
        if prxmatch(_rxcp_sd_, _polis_wrk_) then do;
            _base_sd_ = prxposn(_rxcp_sd_, 1, _polis_wrk_);
            if %is_valid_polis_token(_base_sd_) then do;
                _cp_n_&tag._ = 1; &pfx._1 = strip(_base_sd_); goto _pdone_&tag.;
            end;
        end;
    end;

    /* 6. TBA  content */
    if prxmatch('/^\s*TBA\s{2,}/oi', _polis_wrk_) then do;
        _tba_rest_ = strip(prxchange('s/^\s*TBA\s+//oi', 1, _polis_wrk_));
        %_extract_polis_tokens_m(_tba_rest_, pfx=&pfx., tag=&tag.);
        if _cp_n_&tag._ > 0 then goto _pdone_&tag.;
    end;

    /* 7. & handling */
    if index(_polis_wrk_, '&') > 0 then do;
        %_extract_polis_tokens_m(_polis_wrk_, pfx=&pfx., tag=&tag.);
        if _cp_n_&tag._ > 0 then goto _pdone_&tag.;
    end;

    /* 8. + handling */
    if index(_polis_wrk_, '+') > 0 then do;
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
        %_expand_plus_suffix_m(_polis_wrk_, pfx=&pfx., tag=&tag.);
        if _eps_ok_ then goto _pdone_&tag.;
        %_extract_polis_tokens_m(_polis_wrk_, pfx=&pfx., tag=&tag.);
        if _cp_n_&tag._ > 0 then goto _pdone_&tag.;
    end;

    /* 9. Double-space without + */
    if prxmatch('/\s{2,}/o', _polis_wrk_) then do;
        %_extract_polis_tokens_m(_polis_wrk_, pfx=&pfx., tag=&tag.);
        if _cp_n_&tag._ > 0 then goto _pdone_&tag.;
    end;

    /* 10. Strip VARIOUS/VAR */
    _polis_wrk_ = prxchange('s/\bVARIOUS\b\s*//oi', -1, _polis_wrk_);
    _polis_wrk_ = prxchange('s/\bVAR\b\s*//oi', -1, _polis_wrk_);
    _polis_wrk_ = strip(_polis_wrk_);
    if _polis_wrk_ in ('', '0', '-') then goto _pdone_&tag.;

    /* 11. Multi-slash with all-short check */
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

    /* 12. Fallback single */
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


/* =============================================================================
   SECTION 6: CLEAN SLIP (facul + osbal -- logika identik)
   Setara dengan: clean_slip(val) di Python

   Input : in_var (slip_ori)
   Output: clean_slip_1 ... clean_slip_5
   ============================================================================= */
/* ===========================================================================
   HELPER: %_extract_slip_tokens_m(text_var)
   Splits by 2+ spaces; per block: noise removal -> token extraction with
   dash-split logic and _strip_slip_suffix for single tokens.
   Writes into _csl_out{} / _csl_n_ (must be in scope from caller).
   =========================================================================== */
%macro _extract_slip_tokens_m(text_var, pfx=clean_slip, tag=csl);
    _est_norm_ = prxchange('s/\s{2,}/|/o', -1, strip(&text_var));
    _est_nblk_ = countw(_est_norm_, '|', 'm');
    do _est_bi_ = 1 to _est_nblk_;
        if _csl_n_&tag._ >= 5 then leave;
        _est_blk_   = scan(_est_norm_, _est_bi_, '|', 'm');
        _est_clean_ = prxchange('s/\b(?:JANUARI|FEBRUARI|MARET|APRIL|MEI|JUNI|JULI|AGUSTUS|SEPTEMBER|OKTOBER|NOVEMBER|DESEMBER)\b/ /oi', -1, _est_blk_);
        _est_clean_ = prxchange('s/\b(?:JANUARY|FEBRUARY|MARCH|APRIL|MAY|JUNE|JULY|AUGUST|SEPTEMBER|OCTOBER|NOVEMBER|DECEMBER)\b/ /oi', -1, _est_clean_);
        _est_clean_ = prxchange('s/\b(?:IDR|USD|ENG|P1|P2|P3|P73|NEW|VARIOUS|20\d{2})\b/ /oi', -1, _est_clean_);
        _est_clean_ = prxchange('s/^[\s\-\/+,]+|[\s\-\/+,]+$//o', -1, strip(_est_clean_));
        _est_clean_ = strip(_est_clean_);
        if lengthn(_est_clean_) < 7 then continue;
        _est_tstr_  = prxchange('s/[^A-Za-z0-9\-]/ /oi', -1, _est_clean_);
        _est_tstr_  = prxchange('s/\s{2,}/ /o', -1, _est_tstr_);
        _est_tstr_  = strip(_est_tstr_);
        _est_ntok_  = countw(strip(_est_tstr_), ' ');
        do _est_ti_ = 1 to _est_ntok_;
            if _csl_n_&tag._ >= 5 then leave;
            _est_cand_ = strip(scan(strip(_est_tstr_), _est_ti_, ' '));
            if not prxmatch('/^[A-Za-z0-9][A-Za-z0-9\-]{6,}$/o', _est_cand_) then continue;
            _est_cand_ = prxchange('s/^-+|-+$//o', -1, _est_cand_);
            /* Dash-split: ONE dash + both parts valid + len diff <=2 -> add both */
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
            /* Single token: _strip_slip_suffix then validate */
            if %is_valid_slip_token(_est_cand_) then do;
                if prxmatch('/^.+-\d{1,6}$/o', _est_cand_) then do;
                    _est_g1_ = prxchange('s/^(.+?)-(\d{1,6})$/$1/o', 1, _est_cand_);
                    _est_g2_ = prxchange('s/^(.+?)-(\d{1,6})$/$2/o', 1, _est_cand_);
                    if lengthn(strip(_est_g1_)) > lengthn(strip(_est_g2_)) then
                        _est_cand_ = strip(_est_g1_);
                end;
                if _csl_n_&tag._ < 5 then do; _csl_n_&tag._ + 1; _&tag._out{_csl_n_&tag._} = strip(_est_cand_); end;
            end;
        end;
    end;
%mend _extract_slip_tokens_m;

/* ===========================================================================
   apply_clean_slip: 100% faithful to Python cleaning_facul.clean_slip()
   All scenarios: + with double-space entry blocks, multi-block left check,
   full _extract_slip_tokens with dash-split & _strip_slip_suffix, fallback.
   PERF: No PRXNEXT -- uses helper macro with SCAN.
   Supports custom pfx & tag so multiple calls in one DATA step do not collide.
   =========================================================================== */
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
                    %_extract_slip_tokens_m(_sla_blk_, pfx=&pfx., tag=&tag.);
                end;
            end;
            if _csl_n_&tag._ > 0 then goto _sdone_&tag.;
        end;
    end;

    /* 2. Multi-block left check: if left block invalid, use right blocks only */
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
                %_extract_slip_tokens_m(_slb_rest_, pfx=&pfx., tag=&tag.);
                if _csl_n_&tag._ > 0 then goto _sdone_&tag.;
            end;
        end;
    end;

    /* 3. Main extraction */
    %_extract_slip_tokens_m(_slip_wrk_, pfx=&pfx., tag=&tag.);
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


/* =============================================================================
   SECTION 6b: EXTRACT SERTIFIKAT (facul, osbal, suspend)
   Setara dengan: extract_cert_from_polis() di Python
   ============================================================================= */
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


/* =============================================================================
   SECTION 6c: OSBAL CLSDT VALIDATION & SERTIFIKAT HELPER
   ============================================================================= */
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


/* =============================================================================
   SECTION 7: CLEAN INSURED -- BASE (facul + osbal, bukan suspend)
   Setara dengan: clean_insured(val) dan _clean_insured_name(name) di Python

   Input : in_var (insured_ori)
   Output: clean_insured_1 ... clean_insured_5
   ============================================================================= */
%macro apply_clean_insured(in_var);
    array _cin_out{5} $500 clean_insured_1-clean_insured_5;
    do _i_ = 1 to 5; _cin_out{_i_} = ''; end;
    _ins_in_  = strip(&in_var);
    _cin_n_   = 0;

    if missing(_ins_in_) or _ins_in_ = '' then goto _insured_done_;

    /* Strip outer quotes */
    _ins_in_ = prxchange('s/^[\s"''""]+|[\s"''""]+$//o', -1, _ins_in_);

    /* Proses dengan SCAN atas delimiter yang ditemukan */
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

        /* Skip junk: panjang ≤2, atau pure junk words */
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
            _cin_n_ + 1; _cin_out{_cin_n_} = _pclean_;
        end;
    end;

    /* Fallback: jika semua token di-skip, proses seluruh string */
    if _cin_n_ = 0 then do;
        _fallback_ = _ins_in_;
        _fallback_ = prxchange('s/(?:\b|\.|,)\s*(?:PT|CV|TBK|PERSERO|\(PERSERO\)|LTD|PTE|INC|LLC)\b\.?\s*/ /oi', -1, _fallback_);
        _fallback_ = %normalize_spaces(_fallback_);
        _fallback_ = strip(_fallback_);
        if lengthn(_fallback_) > 2 then do; _cin_n_ = 1; clean_insured_1 = _fallback_; end;
    end;

    /* Jika > 5 token, join menjadi 1 */
    if _cin_n_ > 5 then do;
        _injoin_ = catx(',', of clean_insured_1-clean_insured_5);
        do _i_ = 1 to 5; _cin_out{_i_} = ''; end;
        clean_insured_1 = strip(_injoin_);
    end;

    _insured_done_: ;
    drop _ins_in_ _ins_wrk_ _cin_n_ _part_ _part_up_ _pclean_ _pclean_up_
         _fallback_ _injoin_ _pi_ _n_parts_ _loop_ _i_;
%mend apply_clean_insured;


/* =============================================================================
   SECTION 8: REMOVE ILLEGAL XML CHARS (untuk output Excel yang bersih)
   Setara dengan: df[c].str.replace(r'[\x00-\x08\x0B\x0C\x0E-\x1F]', ...)
   ============================================================================= */
%macro clean_control_chars(var);
    &var = compress(&var, '0001020304050607080B0C0E0F101112131415161718191A1B1C1D1E1F'x);
%mend clean_control_chars;


/* =============================================================================
   SECTION 9: EXPORT TO EXCEL HELPER
   ============================================================================= */
%macro export_excel(dataset, outfile, sheet);
    proc export data=&dataset
        outfile="&outfile"
        dbms=xlsx
        replace;
        sheet="&sheet";
    run;
%mend export_excel;


/* =============================================================================
   END OF macros_cleaning.sas
   ============================================================================= */
