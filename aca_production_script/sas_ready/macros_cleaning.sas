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


/* =============================================================================
   SECTION 2: UTILITY MACROS
   ============================================================================= */

/* Normalize spaces: collapse multiple whitespace → single space, strip
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
   Output: 1 (true) atau 0 (false) — gunakan dalam IF                          */
%macro is_valid_polis_token(tok);
    (lengthn(strip(&tok)) >= 5
     and prxmatch('/\d/', strip(&tok))
     and not prxmatch('/TBA$/oi', strip(&tok)))
%mend is_valid_polis_token;


/* Apakah token valid sebagai nomor slip?
   Setara dengan: _is_valid_slip_token(tok)
   True  = panjang >= 7, ada digit, bukan 4-digit tahun                        */
%macro is_valid_slip_token(tok);
    (lengthn(strip(&tok)) >= 7
     and prxmatch('/\d/', strip(&tok))
     and not prxmatch('/^\d{4}$/', strip(&tok)))
%mend is_valid_slip_token;


/* Strip suffix digit pendek setelah dash dari nomor polis base
   Setara dengan: _strip_polis_base(tok)
   Contoh: '100030817120000016-000149' → '100030817120000016'                   */
%macro strip_polis_base(tok_var, result_var);
    &result_var = strip(&tok_var);
    /* Cek apakah ada pola {base}-{≤6digit} dan base punya ≥10 digit numerik   */
    if prxmatch('/^(.+)-(\d{1,6})$/', strip(&tok_var)) then do;
        _rx_base = prxparse('/^(.+)-(\d{1,6})$/');
        if prxmatch(_rx_base, strip(&tok_var)) then do;
            _base_ = prxposn(_rx_base, 1, strip(&tok_var));
            _suf_  = prxposn(_rx_base, 2, strip(&tok_var));
            /* Hitung digit di base */
            _ndig_ = lengthn(compress(_base_, '', 'kd'));
            if _ndig_ >= 10 and lengthn(_base_) > lengthn(_suf_) then
                &result_var = strip(_base_);
        end;
        call prxfree(_rx_base);
    end;
%mend strip_polis_base;


/* =============================================================================
   SECTION 3: PROC FORMAT untuk mapping bulan Indonesia → nomor bulan
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
   SECTION 4: CLEAN POLIS (digunakan oleh facul + osbal — logika identik)
   Setara dengan: clean_polis(val) di Python

   SAS tidak mendukung fungsi return list, jadi output disimpan ke kolom
   clean_polis_1 s/d clean_polis_5. Panggil macro ini di dalam DATA step.

   Input : _polis_in  (string mentah dari kolom polis_ori)
   Output: clean_polis_1 ... clean_polis_5 (max 5 token)
   ============================================================================= */
%macro apply_clean_polis(in_var);
    /* Reset output kolom */
    array _cp_out{5} $200 clean_polis_1-clean_polis_5;
    do _i_ = 1 to 5; _cp_out{_i_} = ''; end;
    _polis_wrk_ = strip(&in_var);
    _cp_n_ = 0; /* counter token */

    if missing(_polis_wrk_) or _polis_wrk_ = '' then goto _polis_done_;

    /* --- Exception: MOP MARINE, LINESLIP, P1/P2/P3/P73, CANCEL, dll --- */
    if prxmatch('/MOP\s*MARINE|LINE\s*SLIP|LINESLIP|\bP[1-3]\b|\bP73\b|\bCANCEL\b|PENYELESAIAN|HUTANG\s*PIUTANG/oi', _polis_wrk_) then do;
        _cp_n_ = 1;
        clean_polis_1 = %normalize_spaces(_polis_wrk_);
        goto _polis_done_;
    end;

    /* --- Exception: digit TBA digit --- */
    if prxmatch('/\d+TBA\d+/oi', _polis_wrk_) then do;
        _cp_n_ = 1;
        clean_polis_1 = %normalize_spaces(_polis_wrk_);
        goto _polis_done_;
    end;

    /* --- Special words: VARIOUS, TBA saja --- */
    _polis_up_ = upcase(strip(_polis_wrk_));
    if _polis_up_ in ('VARIOUS', 'TBA', 'VARIOUS - SEE ATTACH') then do;
        _cp_n_ = 1; clean_polis_1 = strip(_polis_wrk_); goto _polis_done_;
    end;

    /* --- Dot-suffix: "010114003555.3431" → "010114003555" --- */
    if prxmatch('/^\d[\d\-]+\.\d{1,6}$/', _polis_wrk_) then do;
        _rx_dot_ = prxparse('/^(\d[\d\-]+)\.\d{1,6}$/');
        if prxmatch(_rx_dot_, _polis_wrk_) then do;
            _base_dot_ = prxposn(_rx_dot_, 1, _polis_wrk_);
            if %is_valid_polis_token(_base_dot_) then do;
                _cp_n_ = 1; clean_polis_1 = strip(_base_dot_);
                call prxfree(_rx_dot_); goto _polis_done_;
            end;
        end;
        call prxfree(_rx_dot_);
    end;

    /* --- S/D range: ambil base polis sebelum range --- */
    if prxmatch('/S\/D/oi', _polis_wrk_) then do;
        _rx_sd_ = prxparse('/^(\d[\d\-]+?)[-]?\d+S\/D\d+/oi');
        if prxmatch(_rx_sd_, _polis_wrk_) then do;
            _base_sd_ = prxposn(_rx_sd_, 1, _polis_wrk_);
            _cp_n_ = 1; clean_polis_1 = strip(_base_sd_);
            call prxfree(_rx_sd_); goto _polis_done_;
        end;
        call prxfree(_rx_sd_);
    end;

    /* --- Strip VARIOUS/VAR/¿ dan proses slash-split --- */
    _polis_wrk_ = prxchange('s/\bVARIOUS\b\s*//oi', -1, _polis_wrk_);
    _polis_wrk_ = prxchange('s/\bVAR\b\s*//oi', -1, _polis_wrk_);
    _polis_wrk_ = translate(_polis_wrk_, ' ', '¿');
    _polis_wrk_ = strip(_polis_wrk_);
    if _polis_wrk_ = '' then goto _polis_done_;

    /* Split by slash '/' — ambil token yang valid */
    _slash_n_ = countw(_polis_wrk_, '/', 'm');
    if _slash_n_ > 1 then do;
        do _si_ = 1 to _slash_n_;
            _tok_ = strip(scan(_polis_wrk_, _si_, '/', 'm'));
            if %is_valid_polis_token(_tok_) and _cp_n_ < 5 then do;
                _cp_n_ + 1;
                _cp_out{_cp_n_} = _tok_;
            end;
        end;
    end;

    /* Jika tidak ada token dari slash, ambil langsung */
    if _cp_n_ = 0 then do;
        _tok_ = strip(_polis_wrk_);
        if %is_valid_polis_token(_tok_) then do;
            /* Strip suffix digit jika base punya ≥10 digit numerik */
            _ndig_tok_ = lengthn(compress(_tok_, '', 'kd'));
            if _ndig_tok_ >= 10 then %strip_polis_base(_tok_, _tok_);
            _cp_n_ = 1; clean_polis_1 = _tok_;
        end;
    end;

    _polis_done_: ;
    /* Jika > 5 token, join menjadi 1 comma-separated di kolom 1 */
    if _cp_n_ > 5 then do;
        _joined_ = catx(',', of clean_polis_1-clean_polis_5);
        do _i_ = 1 to 5; _cp_out{_i_} = ''; end;
        clean_polis_1 = strip(_joined_);
    end;
    drop _polis_wrk_ _polis_up_ _cp_n_ _si_ _slash_n_ _tok_ _ndig_tok_
         _base_dot_ _base_sd_ _joined_ _i_;
%mend apply_clean_polis;


/* =============================================================================
   SECTION 5: EXTRACT SERTIF FROM POLIS (facul + suspend)
   Setara dengan: extract_cert_from_polis(val) di Python

   Input : _polis_in (polis_ori)
   Output: clean_sertif_1 ... clean_sertif_5
   ============================================================================= */
%macro apply_extract_sertif(in_var);
    array _cs_out{5} $10 clean_sertif_1-clean_sertif_5;
    do _i_ = 1 to 5; _cs_out{_i_} = ''; end;
    _sertif_wrk_ = strip(&in_var);
    _cs_n_ = 0;

    if missing(_sertif_wrk_) or _sertif_wrk_ = '' then goto _sertif_done_;

    /* Cek S/D range terlebih dahulu */
    _rx_sd2_ = prxparse('/(\d{1,6})\s*S\/D\s*(\d{1,6})/oi');
    if prxmatch(_rx_sd2_, _sertif_wrk_) then do;
        _sd_start_ = input(prxposn(_rx_sd2_, 1, _sertif_wrk_), 8.);
        _sd_end_   = input(prxposn(_rx_sd2_, 2, _sertif_wrk_), 8.);
        if _sd_start_ > _sd_end_ then do;
            _sd_tmp_ = _sd_start_; _sd_start_ = _sd_end_; _sd_end_ = _sd_tmp_;
        end;
        do _sd_i_ = _sd_start_ to min(_sd_end_, _sd_start_ + 4);
            _cs_n_ + 1;
            if _cs_n_ <= 5 then _cs_out{_cs_n_} = put(_sd_i_, z6.);
        end;
        call prxfree(_rx_sd2_); goto _sertif_done_;
    end;
    call prxfree(_rx_sd2_);

    /* Cek suffix digit 1-6 setelah tanda '-' */
    _rx_sfx_ = prxparse('/[-\s]+\s*(\d{1,6})(?:[^0-9]|$)/o');
    if prxmatch(_rx_sfx_, _sertif_wrk_) then do;
        _cert_raw_ = prxposn(_rx_sfx_, 1, _sertif_wrk_);
        if prxmatch('/^\d{1,6}$/', strip(_cert_raw_)) then do;
            _cs_n_ = 1;
            clean_sertif_1 = put(input(strip(_cert_raw_), 8.), z6.);
        end;
    end;
    call prxfree(_rx_sfx_);

    _sertif_done_: ;
    drop _sertif_wrk_ _cs_n_ _sd_start_ _sd_end_ _sd_tmp_ _sd_i_ _cert_raw_ _i_;
%mend apply_extract_sertif;


/* =============================================================================
   SECTION 6: CLEAN SLIP (facul + osbal — logika identik)
   Setara dengan: clean_slip(val) di Python

   Input : in_var (slip_ori)
   Output: clean_slip_1 ... clean_slip_5
   ============================================================================= */
%macro apply_clean_slip(in_var);
    array _csl_out{5} $200 clean_slip_1-clean_slip_5;
    do _i_ = 1 to 5; _csl_out{_i_} = ''; end;
    _slip_wrk_ = strip(&in_var);
    _csl_n_ = 0;

    if missing(_slip_wrk_) or _slip_wrk_ = '' then goto _slip_done_;

    /* Strip noise words (bulan, tahun, currency, dsb) */
    _slip_clean_ = prxchange('s/\b(?:JANUARI|FEBRUARI|MARET|APRIL|MEI|JUNI|JULI|AGUSTUS|SEPTEMBER|OKTOBER|NOVEMBER|DESEMBER)\b/ /oi', -1, _slip_wrk_);
    _slip_clean_ = prxchange('s/\b(?:JANUARY|FEBRUARY|MARCH|APRIL|MAY|JUNE|JULY|AUGUST|SEPTEMBER|OCTOBER|NOVEMBER|DECEMBER)\b/ /oi', -1, _slip_clean_);
    _slip_clean_ = prxchange('s/\b(?:IDR|USD|ENG|P1|P2|P3|P73|NEW|VARIOUS|20\d{2})\b/ /oi', -1, _slip_clean_);
    _slip_clean_ = prxchange('s/^[\s\-\/+,]+|[\s\-\/+,]+$//o', -1, strip(_slip_clean_));
    _slip_clean_ = strip(_slip_clean_);

    /* Extract token alfanumerik panjang ≥7 */
    _rx_sliptok_ = prxparse('/[A-Z0-9][A-Z0-9\-]{6,}/oi');
    _slip_start_ = 1; _slip_stop_ = 0; _slip_pos_ = 1;
    call prxnext(_rx_sliptok_, _slip_start_, length(_slip_clean_), _slip_clean_, _slip_pos_, _slip_len_);
    do while (_slip_pos_ > 0 and _csl_n_ < 5);
        _cand_ = substr(_slip_clean_, _slip_pos_, _slip_len_);
        _cand_ = prxchange('s/^-|-$//o', -1, _cand_); /* strip leading/trailing dash */
        if %is_valid_slip_token(_cand_) then do;
            _csl_n_ + 1; _csl_out{_csl_n_} = strip(_cand_);
        end;
        call prxnext(_rx_sliptok_, _slip_start_, length(_slip_clean_), _slip_clean_, _slip_pos_, _slip_len_);
    end;
    call prxfree(_rx_sliptok_);

    /* Jika tidak ada token, gunakan nilai clean yang ada */
    if _csl_n_ = 0 then do;
        _slip_fb_ = prxchange('s/\bVARIOUS\b\s*//oi', -1, _slip_wrk_);
        _slip_fb_ = prxchange('s/^\s*-\s*//o', -1, strip(_slip_fb_));
        _slip_fb_ = strip(_slip_fb_);
        if _slip_fb_ ne '' then do; _csl_n_ = 1; clean_slip_1 = _slip_fb_; end;
    end;

    /* Jika > 5 token, join menjadi 1 */
    if _csl_n_ > 5 then do;
        _slipjoin_ = catx(',', of clean_slip_1-clean_slip_5);
        do _i_ = 1 to 5; _csl_out{_i_} = ''; end;
        clean_slip_1 = strip(_slipjoin_);
    end;

    _slip_done_: ;
    drop _slip_wrk_ _slip_clean_ _csl_n_ _cand_ _slip_fb_ _slipjoin_
         _rx_sliptok_ _slip_start_ _slip_stop_ _slip_pos_ _slip_len_ _i_;
%mend apply_clean_slip;


/* =============================================================================
   SECTION 7: CLEAN INSURED — BASE (facul + osbal, bukan suspend)
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
    _ins_in_ = prxchange('s/^[\s"''""«»]+|[\s"''""«»]+$//o', -1, _ins_in_);

    /* Split berdasarkan AND/OR, C/Q, Q/Q, QQ, koma, slash */
    _split_rx_  = prxparse('/\bAND\s*\/\s*OR\b|\bC\s*\/\s*Q\b|\bQ\s*\/\s*Q\b|\bQ\.?Q\.?\b|,(?![^(]*\))|\/(?![^(]*\))/oi');
    _ins_start_ = 1; _ins_stop_  = 0; _prev_pos_  = 1;
    _ins_wrk_   = _ins_in_;

    /* Proses dengan SCAN atas delimiter yang ditemukan */
    /* SAS: untuk simplicity, split via scan dengan delimiter khusus */
    /* Gunakan pendekatan: replace delimiter → '|' lalu SCAN */
    _ins_wrk_ = prxchange('s/\bAND\s*\/\s*OR\b/|/oi', -1, _ins_wrk_);
    _ins_wrk_ = prxchange('s/\bC\s*\/\s*Q\b/|/oi',   -1, _ins_wrk_);
    _ins_wrk_ = prxchange('s/\bQ\s*\/\s*Q\b/|/oi',   -1, _ins_wrk_);
    _ins_wrk_ = prxchange('s/\bQ\.?Q\.?\b/|/oi',      -1, _ins_wrk_);
    /* Koma dan slash di luar tanda kurung (simplified: replace semua koma & slash) */
    _ins_wrk_ = prxchange('s/,/|/o', -1, _ins_wrk_);
    _ins_wrk_ = prxchange('s/\//|/o', -1, _ins_wrk_);

    call prxfree(_split_rx_);

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
            _pclean_ = prxchange('s/^[\s"''""«»\-\.,\/;:]+|[\s"''""«»\-\.,\/;:]+$//o', -1, _pclean_);
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
         _fallback_ _injoin_ _pi_ _n_parts_ _loop_ _i_ _split_rx_
         _ins_start_ _ins_stop_ _prev_pos_;
%mend apply_clean_insured;


/* =============================================================================
   SECTION 8: REMOVE ILLEGAL XML CHARS (untuk output Excel yang bersih)
   Setara dengan: df[c].str.replace(r'[\x00-\x08\x0B\x0C\x0E-\x1F]', ...)
   ============================================================================= */
%macro clean_control_chars(var);
    &var = prxchange('s/[\x00-\x08\x0B\x0C\x0E-\x1F]//o', -1, &var);
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
