/* =============================================================================
   MACROS_CLEANING_ASTRA.SAS
   Shared macro library untuk script cleaning cedant PT ASURANSI ASTRA BUANA.
   Kompatibel dengan SAS Viya / SAS 9.4 Studio.
   ============================================================================= */

options validvarname=any validmemname=extend;

/* ── 1. Cek keberadaan variabel di dataset ─────────────────────────────────── */
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


/* ── 2. Format mapping bulan Indonesia ke nomor bulan ──────────────────────── */
%macro define_bulan_format();
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
%mend define_bulan_format;


/* ── 3. Utility Normalisasi String ─────────────────────────────────────────── */
%macro normalize_spaces(val);
    strip(prxchange('s/\s{2,}/ /', -1, strip(&val)))
%mend normalize_spaces;

%macro clean_control_chars(var);
    &var = prxchange('s/[\x00-\x08\x0B\x0C\x0E-\x1F]//o', -1, &var);
%mend clean_control_chars;


/* ── 4. Standarisasi S/D -> SD (Aturan Khusus Astra) ────────────────────────── */
%macro standarisasi_sd(var);
    if index(upcase(&var.), 'S/D') > 0 or index(upcase(&var.), 'S / D') > 0 then
        &var = prxchange('s/\s*S\s*\/\s*D\s*/ SD /oi', -1, &var);
%mend standarisasi_sd;


/* ── 5. Clean Polis Khusus Astra Buana ─────────────────────────────────────── */
/* Aturan:
   - Minimal digit nomor polis Astra = 11 digit (\d{11,})
   - Valid full lengths: 11, 12, 13, 16
   - Hilangkan S/D -> SD
   - Output ke 5 kolom: POLICY_CLEAN_1 s/d POLICY_CLEAN_5
*/
%macro apply_clean_polis_astra(in_var);
    array _pol_out{5} $200 POLICY_CLEAN_1-POLICY_CLEAN_5;
    do _pi_ = 1 to 5; _pol_out{_pi_} = ''; end;
    _pwrk_ = strip(&in_var.);
    _pn_ = 0;

    if missing(_pwrk_) or _pwrk_ = '' then goto _polis_astra_done_;

    /* Ganti S/D ke SD */
    %standarisasi_sd(_pwrk_);

    /* Pola AS PER LIST, LINESLIP, TBA/VAR */
    if prxmatch('/AS\s*PER\s*LIST|ASTRA\s*EX\s*LS|\bTBA\b|\bVARIOUS\b/oi', _pwrk_) then do;
        _pn_ = 1; POLICY_CLEAN_1 = %normalize_spaces(_pwrk_);
        goto _polis_astra_done_;
    end;

    /* Hapus catatan dalam kurung e.g. (LANJUTAN ...) */
    _pwrk_ = prxchange('s/\(\s*LANJUTAN\b[^\)]*\)//oi', -1, _pwrk_);
    _pwrk_ = %normalize_spaces(_pwrk_);

    /* Tokenize slash '/' atau koma ',' */
    _np_parts_ = countw(_pwrk_, '/,', 'm');
    if _np_parts_ > 1 then do;
        do _si_ = 1 to _np_parts_;
            _tok_ = strip(scan(_pwrk_, _si_, '/,', 'm'));
            /* Cek token numerik >= 11 digit atau format polis */
            if prxmatch('/\d{11,}/', _tok_) and _pn_ < 5 then do;
                _pn_ + 1;
                _pol_out{_pn_} = _tok_;
            end;
        end;
    end;

    /* Fallback jika tidak ada hasil dari split */
    if _pn_ = 0 then do;
        if prxmatch('/\d{11,}/', _pwrk_) or lengthn(_pwrk_) >= 6 then do;
            _pn_ = 1; POLICY_CLEAN_1 = _pwrk_;
        end;
    end;

    _polis_astra_done_: ;
    drop _pwrk_ _pn_ _pi_ _si_ _np_parts_ _tok_;
%mend apply_clean_polis_astra;


/* ── 6. Clean Slip Khusus Astra Buana ──────────────────────────────────────── */
/* Aturan:
   - Hapus noise label: bulan, tahun, currency IDR/USD/JPY/SGD/EUR/GBP
   - Hapus kata border: BORDERO, BORDEROUX, SUMMARY
   - Output ke 5 kolom: SLIP_CLEAN_1 s/d SLIP_CLEAN_5
*/
%macro apply_clean_slip_astra(in_var);
    array _slp_out{5} $200 SLIP_CLEAN_1-SLIP_CLEAN_5;
    do _sli_ = 1 to 5; _slp_out{_sli_} = ''; end;
    _swrk_ = strip(&in_var.);
    _sln_ = 0;

    if missing(_swrk_) or _swrk_ = '' then goto _slip_astra_done_;

    %standarisasi_sd(_swrk_);

    /* Hapus prefix tanggal/bulan/tahun/ccy */
    _swrk_ = prxchange('s/^(?:JANUARI|FEBRUARI|MARET|APRIL|MEI|JUNI|JULI|AGUSTUS|SEPTEMBER|OKTOBER|NOVEMBER|DESEMBER|\d{4}|IDR|USD|JPY|EUR|SGD|GBP|\s*[-\/])+//oi', -1, _swrk_);
    /* Hapus noise kata border & summary */
    _swrk_ = prxchange('s/\b(?:BORDEREAUX|BORDEROUX|BORDERO|BORD|SUMMARY|ASTRA|SISA|OSBAL)\b/ /oi', -1, _swrk_);
    _swrk_ = %normalize_spaces(_swrk_);

    /* Extract token kode alfanumerik */
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


/* ── 7. Clean Insured Khusus Astra Buana ────────────────────────────────────── */
/* Aturan:
   - Hapus entitas: PT, CV, TBK, PTE, LTD, PELAYARAN, (PERSERO)
   - Hapus tanda baca titik dan koma
   - Split berdasarkan AND/OR, QQ, C/Q, koma, slash
   - Output ke 5 kolom: INSURED_CLEAN_1 s/d INSURED_CLEAN_5
*/
%macro apply_clean_insured_astra(in_var);
    array _ins_out{5} $500 INSURED_CLEAN_1-INSURED_CLEAN_5;
    do _ii_ = 1 to 5; _ins_out{_ii_} = ''; end;
    _iwrk_ = strip(&in_var.);
    _inn_ = 0;

    if missing(_iwrk_) or _iwrk_ = '' then goto _insured_astra_done_;

    /* Hapus prefix AUTO FACILITY / CARGO FACILITY / LINESLIP */
    _iwrk_ = prxchange('s/^(?:(?:AUTO\s+FACILITY|CARGO\s+FACILITY|LINESLIP)\s*\/\s*)+//oi', -1, _iwrk_);
    /* Hapus entitas hukum */
    _iwrk_ = prxchange('s/\b(?:PT|CV|TBK|PTE|LTD|PELAYARAN|\(PERSERO\)|\bPERSERO\b)\b\.?/ /oi', -1, _iwrk_);
    /* Ganti delimiter AND/OR, QQ, slash, koma ke '|' */
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
