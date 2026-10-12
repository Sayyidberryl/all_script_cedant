/* =========================================================================
   PROGRAM  : lippo_cleaning.sas
   CEDANT   : PT LIPPO GENERAL INSURANCE (LIPPO)
   PLATFORM : SAS Viya LTS 2026.03 (Cloud Analytic Services - In-Memory)
   PURPOSE  : Production Data Cleaning Pipeline (100% Identik Logika Python)
              - Pure Base SAS DATA Step in CAS
              - Tarik data dari CASLIB FSI_DB via PROC FEDSQL (mycas)
              - In-Memory Data Cleaning di CASUSER (Paralel, Super Cepat)
              - Ekstrak Sertifikat & Pembersihan Lengkap (Polis, Slip, Insured)
              - Urutan Kolom Sesuai Persis Skema Database PostgreSQL FSI_DB:
                * SUSPENSE_DATA_OSBAL_CLEAN (42 Kolom Standar)
                * SUSPENSE_DATA_FACUL_CLEAN (36 Kolom Standar)
              - Promosi ke CASLIB Public (Pola Filter & Merge Aman Multi-Cedant)
   ========================================================================= */

options validvarname=any validmemname=extend compress=yes;

/* =============================================================================
   PHASE 0: CAS SESSION INITIALIZATION & LIBREF MAPPING
   ============================================================================= */

/* 0.1. Reset sesi lama & inisialisasi sesi MYCAS sesuai environment Starcore */
cas mycas terminate;
cas mycas;
caslib _all_ assign;

/* 0.2. Validasi Ketersediaan Librefs Wajib */
data _null_;
    length missing_libs $100;
    missing_libs = '';
    if libref('CASUSER') ne 0 then missing_libs = catx(' ', missing_libs, 'CASUSER');
    if libref('PUBLIC')  ne 0 then missing_libs = catx(' ', missing_libs, 'PUBLIC');
    if libref('FSI_DB')  ne 0 then missing_libs = catx(' ', missing_libs, 'FSI_DB');

    if not missing(missing_libs) then do;
        put 'ERROR: [INIT FAIL] Libref berikut tidak terdaftar di SAS session: ' missing_libs;
        abort cancel;
    end;
    else do;
        put '[INIT PASS] Semua Libref wajib (CASUSER, PUBLIC, FSI_DB) berhasil terpetakan ke session MYCAS.';
    end;
run;


/* =============================================================================
   PHASE 1: LOAD & CLEAN DATA FACULTATIVE (SUSPESE_FACUL_ORI_2026)
   ============================================================================= */

/* 1.1. Bersihkan tabel temporary lama & Tarik Data Mentah via PROC FEDSQL */
proc delete data=casuser.facul_raw; run;
proc delete data=casuser.lippo_facul_clean; run;

proc fedsql sessref=mycas;
    create table casuser.facul_raw as
    select * from FSI_DB.SUSPESE_FACUL_ORI_2026
    where "COMP_NAME" like 'PT LIPPO GENERAL INSURANCE%';
quit;

/* 1.2. Cleaning Data Facul In-Memory di CAS (100% Logika Python Lippo) */
data casuser.lippo_facul_clean;
    set casuser.facul_raw;

    length BUSSINESS_PATNER $500
           INSURED_CLEAN_1-INSURED_CLEAN_5 $500
           POLICY_CLEAN_1-POLICY_CLEAN_5 $500
           SERTIF_CLEAN_1-SERTIF_CLEAN_5 $50
           SLIP_CLEAN_1-SLIP_CLEAN_5 $500
           clean_polis_1-clean_polis_5 $500
           clean_slip_1-clean_slip_5 $500
           clean_insured_1-clean_insured_5 $500
           clean_sertif_1-clean_sertif_5 $50
           POLIS_CLN_1-POLIS_CLN_5 $500
           SLIP_NO_CLN_1-SLIP_NO_CLN_5 $500
           INSURED_CLN_1-INSURED_CLN_5 $500
           polis_ori slip_ori insured_ori $500;

    /* Inisialisasi awal */
    array _pclean_{5} $500 POLICY_CLEAN_1-POLICY_CLEAN_5;
    array _sclean_{5} $50  SERTIF_CLEAN_1-SERTIF_CLEAN_5;
    array _slclean_{5} $500 SLIP_CLEAN_1-SLIP_CLEAN_5;
    array _inclean_{5} $500 INSURED_CLEAN_1-INSURED_CLEAN_5;

    do _i_ = 1 to 5;
        _pclean_{_i_} = '';
        _sclean_{_i_} = '';
        _slclean_{_i_} = '';
        _inclean_{_i_} = '';
    end;

    /* Ambil kolom asli sebagai string */
    polis_ori   = strip(cats(FAC_POLICY_NO));
    slip_ori    = strip(cats(FAC_SLIP));
    insured_ori = strip(cats(FAC_INSURED));

    /* A. Business Partner: COMP_NAME2 (jika ada dan bukan DIRECT), fallback ke COMP_NAME */
    _broker_ = strip(cats(COMP_NAME2));
    if missing(_broker_) or upcase(_broker_) in ('DIRECT', '') then
        BUSSINESS_PATNER = strip(cats(COMP_NAME));
    else
        BUSSINESS_PATNER = _broker_;

    /* B. Extract Sertifikat dari Polis (S/D Range atau suffix dash) */
    _p_raw_ = polis_ori;
    retain _rx_sd_f_ _rx_cert_f_ 0;
    if _rx_sd_f_   = 0 then _rx_sd_f_   = prxparse('/(\d{1,6})\s*S\/D\s*(\d{1,6})/oi');
    if _rx_cert_f_ = 0 then _rx_cert_f_ = prxparse('/[-\s]+\s*(\d{1,6})(?:[^0-9]|$)/o');

    if prxmatch(_rx_sd_f_, _p_raw_) then do;
        call prxposn(_rx_sd_f_, 1, _pos1_, _len1_);
        call prxposn(_rx_sd_f_, 2, _pos2_, _len2_);
        if _pos1_ > 0 and _pos2_ > 0 then do;
            _sd1_ = input(substr(_p_raw_, _pos1_, _len1_), 8.);
            _sd2_ = input(substr(_p_raw_, _pos2_, _len2_), 8.);
            if _sd1_ > _sd2_ then do; _tmp_ = _sd1_; _sd1_ = _sd2_; _sd2_ = _tmp_; end;
            if _sd1_ <= _sd2_ then SERTIF_CLEAN_1 = put(_sd1_, z6.);
            if _sd1_ + 1 <= _sd2_ then SERTIF_CLEAN_2 = put(_sd1_ + 1, z6.);
            if _sd1_ + 2 <= _sd2_ then SERTIF_CLEAN_3 = put(_sd1_ + 2, z6.);
            if _sd1_ + 3 <= _sd2_ then SERTIF_CLEAN_4 = put(_sd1_ + 3, z6.);
            if _sd1_ + 4 <= _sd2_ then SERTIF_CLEAN_5 = put(_sd1_ + 4, z6.);
        end;
    end;
    else if prxmatch(_rx_cert_f_, _p_raw_) then do;
        call prxposn(_rx_cert_f_, 1, _pos1_, _len1_);
        if _pos1_ > 0 and _len1_ > 0 then do;
            _cval_ = strip(substr(_p_raw_, _pos1_, _len1_));
            if prxmatch('/^\d{1,6}$/', _cval_) then
                SERTIF_CLEAN_1 = put(input(_cval_, 8.), z6.);
        end;
    end;

    /* C. Clean Polis Lippo (MOP MARINE, LINESLIP, P1-P3/P73, Plus (+) expansion, Dash-suffix expansion) */
    if not missing(_p_raw_) and not (_p_raw_ in ('0', '-')) then do;
        _pwork_ = _p_raw_;

        if prxmatch('/MOP\s*MARINE|LINE\s*SLIP|LINESLIP|\bP[1-3]\b|\bP73\b|\bCANCEL\b|PENYELESAIAN|HUTANG\s*PIUTANG/oi', _pwork_) then do;
            POLICY_CLEAN_1 = strip(prxchange('s/\s{2,}/ /o', -1, strip(_pwork_)));
        end;
        else if upcase(strip(_pwork_)) in ('VARIOUS', 'TBA', 'VARIOUS - SEE ATTACH') then do;
            POLICY_CLEAN_1 = strip(_pwork_);
        end;
        else do;
            /* Pola khusus Lippo: Plus-separated tokens e.g. "1801282000154 + 447 + 047 + 042" */
            if index(_pwork_, '+') > 0 then do;
                _n_plus_ = countw(_pwork_, '+');
                _p_base_prefix_ = '';
                _c_idx_ = 0;
                do _pi_ = 1 to _n_plus_;
                    if _c_idx_ >= 5 then leave;
                    _ptok_ = strip(scan(_pwork_, _pi_, '+'));
                    _ptok_ = prxchange('s/[^\w\-]+//o', -1, _ptok_);
                    if _pi_ = 1 then do;
                        /* Token pertama adalah basis penuh */
                        if lengthn(_ptok_) >= 10 then do;
                            _c_idx_ + 1;
                            _pclean_{_c_idx_} = _ptok_;
                            /* Simpan base prefix (misal ambil prefix di depan digit ekor) */
                            _p_base_prefix_ = substr(_ptok_, 1, lengthn(_ptok_) - 3);
                        end;
                        else if lengthn(_ptok_) >= 6 then do;
                            _c_idx_ + 1;
                            _pclean_{_c_idx_} = _ptok_;
                        end;
                    end;
                    else do;
                        /* Token selanjutnya: jika digit pendek (3-4 digit), gabungkan dengan base prefix */
                        if lengthn(_ptok_) in (2, 3, 4) and not missing(_p_base_prefix_) then do;
                            _c_idx_ + 1;
                            _pclean_{_c_idx_} = cats(_p_base_prefix_, put(input(_ptok_, 8.), z3.));
                        end;
                        else if lengthn(_ptok_) >= 6 then do;
                            _c_idx_ + 1;
                            _pclean_{_c_idx_} = _ptok_;
                        end;
                    end;
                end;
            end;
            else if prxmatch('/^\d{8,}-\d{3}(?:-\d{3})+$/', strip(_pwork_)) then do;
                /* Pola dash berulang Lippo e.g. "15010918003-010-002-035" */
                _first_t_ = scan(strip(_pwork_), 1, '-');
                _p_pfx_   = substr(_first_t_, 1, lengthn(_first_t_) - 3);
                _n_d_     = countw(strip(_pwork_), '-');
                _c_idx_   = 0;
                do _pi_ = 1 to _n_d_;
                    if _c_idx_ >= 5 then leave;
                    _dtok_ = strip(scan(strip(_pwork_), _pi_, '-'));
                    _c_idx_ + 1;
                    if _pi_ = 1 then _pclean_{_c_idx_} = _dtok_;
                    else _pclean_{_c_idx_} = cats(_p_pfx_, _dtok_);
                end;
            end;
            else do;
                /* General tokenization: slash '/', koma ',', titik koma ';' */
                _pwork_ = prxchange('s/[\/\+,;&]+/ /o', -1, _pwork_);
                _pwork_ = prxchange('s/\s{2,}/ /o', -1, strip(_pwork_));
                _ntok_  = countw(strip(_pwork_), ' ');
                _c_idx_ = 0;

                do _i_ = 1 to _ntok_;
                    if _c_idx_ >= 5 then leave;
                    _tok_ = strip(scan(strip(_pwork_), _i_, ' '));
                    _tok_ = prxchange('s/^[\-\/]+|[\-\/]+$//o', -1, _tok_);

                    if lengthn(_tok_) >= 5 and anydigit(_tok_) > 0 and not prxmatch('/TBA$/oi', _tok_) then do;
                        if lengthn(compress(_tok_, '0123456789', 'k')) >= 10 then do;
                            if index(_tok_, '-') > 0 then do;
                                _base_ = scan(_tok_, 1, '-');
                                _suf_  = scan(_tok_, 2, '-');
                                if lengthn(compress(_base_, '0123456789', 'k')) >= 10 and lengthn(compress(_suf_, '0123456789', 'k')) <= 6 then
                                    _tok_ = _base_;
                            end;
                        end;
                        _c_idx_ + 1;
                        _pclean_{_c_idx_} = _tok_;
                    end;
                end;
            end;
            if _c_idx_ = 0 then POLICY_CLEAN_1 = _p_raw_;
        end;
    end;

    /* D. Clean Slip Lippo (Hapus noise kata border, bulan, tahun, IDR/USD, plus-expansion) */
    if not missing(slip_ori) and not (slip_ori in ('0', '-')) then do;
        _swork_ = slip_ori;
        _swork_ = prxchange('s/\b(?:SUMMARY|BORDER[OA]?|BORDRO|SINGGLESHIPMENT|VARIOUS)\b/ /oi', -1, _swork_);
        _swork_ = prxchange('s/\b(?:JANUARI|FEBRUARI|MARET|APRIL|MEI|JUNI|JULI|AGUSTUS|SEPTEMBER|OKTOBER|NOVEMBER|DESEMBER)\b/ /oi', -1, _swork_);
        _swork_ = prxchange('s/\b(?:JANUARY|FEBRUARY|MARCH|APRIL|MAY|JUNE|JULY|AUGUST|SEPTEMBER|OCTOBER|NOVEMBER|DECEMBER)\b/ /oi', -1, _swork_);
        _swork_ = prxchange('s/\b(?:IDR|USD|ENG|P1|P2|P3|P73|NEW|20\d{2})\b/ /oi', -1, _swork_);

        if index(_swork_, '+') > 0 then do;
            _n_spl_ = countw(_swork_, '+');
            _s_base_prefix_ = '';
            _s_idx_ = 0;
            do _sj_ = 1 to _n_spl_;
                if _s_idx_ >= 5 then leave;
                _s_tok_ = strip(scan(_swork_, _sj_, '+'));
                _s_tok_ = prxchange('s/[^\w\-]+//o', -1, _s_tok_);
                if _sj_ = 1 then do;
                    if lengthn(_s_tok_) >= 10 then do;
                        _s_idx_ + 1;
                        _slclean_{_s_idx_} = _s_tok_;
                        _s_base_prefix_ = substr(_s_tok_, 1, lengthn(_s_tok_) - 3);
                    end;
                    else if lengthn(_s_tok_) >= 6 then do;
                        _s_idx_ + 1;
                        _slclean_{_s_idx_} = _s_tok_;
                    end;
                end;
                else do;
                    if lengthn(_s_tok_) in (2, 3, 4) and not missing(_s_base_prefix_) then do;
                        _s_idx_ + 1;
                        _slclean_{_s_idx_} = cats(_s_base_prefix_, put(input(_s_tok_, 8.), z3.));
                    end;
                    else if lengthn(_s_tok_) >= 6 then do;
                        _s_idx_ + 1;
                        _slclean_{_s_idx_} = _s_tok_;
                    end;
                end;
            end;
        end;
        else if prxmatch('/^\d{8,}-\d{3}(?:-\d{3})+$/', strip(_swork_)) then do;
            _s_first_t_ = scan(strip(_swork_), 1, '-');
            _s_pfx_     = substr(_s_first_t_, 1, lengthn(_s_first_t_) - 3);
            _s_nd_      = countw(strip(_swork_), '-');
            _s_idx_     = 0;
            do _sj_ = 1 to _s_nd_;
                if _s_idx_ >= 5 then leave;
                _s_dtok_ = strip(scan(strip(_swork_), _sj_, '-'));
                _s_idx_ + 1;
                if _sj_ = 1 then _slclean_{_s_idx_} = _s_dtok_;
                else _slclean_{_s_idx_} = cats(_s_pfx_, _s_dtok_);
            end;
        end;
        else do;
            _swork_ = prxchange('s/[\/\+,;&]+/ /o', -1, _swork_);
            _swork_ = prxchange('s/\s{2,}/ /o', -1, strip(_swork_));
            _sntok_ = countw(strip(_swork_), ' ');
            _s_idx_ = 0;

            do _j_ = 1 to _sntok_;
                if _s_idx_ >= 5 then leave;
                _stok_ = strip(scan(strip(_swork_), _j_, ' '));
                _stok_ = prxchange('s/^[\-\/]+|[\-\/]+$//o', -1, _stok_);
                if lengthn(_stok_) >= 7 and anydigit(_stok_) > 0 and not prxmatch('/^\d{4}$/', _stok_) then do;
                    if index(_stok_, '-') > 0 then do;
                        _sbase_ = scan(_stok_, 1, '-');
                        _ssuf_  = scan(_stok_, 2, '-');
                        if lengthn(_sbase_) > lengthn(_ssuf_) and lengthn(_ssuf_) <= 6 then
                            _stok_ = _sbase_;
                    end;
                    _s_idx_ + 1;
                    _slclean_{_s_idx_} = _stok_;
                end;
            end;
        end;
        if _s_idx_ = 0 then SLIP_CLEAN_1 = slip_ori;
    end;

    /* E. Clean Insured Lippo (Hapus entitas hukum PT, CV, TBK, PERSERO, LTD, honorifics) */
    if not missing(insured_ori) then do;
        _ins_work_ = insured_ori;
        _ins_work_ = prxchange('s/(?i)\bAS\b\s+(?:THE\s+)?(?:PRINCIPAL|OFF-TAKER|MAINTENANCE|CONTRACTOR).*//o', -1, _ins_work_);
        _ins_work_ = prxchange('s/(?i)\bAND\s+ALL\s+SUBSIDIARI.*//o', -1, _ins_work_);
        _ins_work_ = prxchange('s/(?i)\b(?:AND|AN|OR)\s*\/\s*(?:AND|OR)\b/,/o', -1, _ins_work_);
        _ins_work_ = prxchange('s/(?i)\bAND\s+OR\b/,/o', -1, _ins_work_);
        _ins_work_ = prxchange('s/(?i)\b(?:QQ|C\/Q|Q\/Q)\b/|/o', -1, _ins_work_);
        _ins_work_ = prxchange('s/[\/,\+;]/|/o', -1, _ins_work_);
        _ins_work_ = prxchange('s/(?i)\b(?:PT\.?|CV\.?)\b/|/o', -1, _ins_work_);

        _n_ins_ = countw(strip(_ins_work_), '|');
        _in_idx_ = 0;

        do _k_ = 1 to _n_ins_;
            if _in_idx_ >= 5 then leave;
            _part_ = strip(scan(strip(_ins_work_), _k_, '|'));
            _part_ = prxchange('s/(?i)\b(?:PT|CV|TBK|\(PERSERO\)|PERSERO|LTD|PTE|INC|LLC|UD|PD|NV|BV|GMBH|BHD)\b\.?/ /o', -1, _part_);
            _part_ = prxchange('s/(?i)\b(?:BAPAK|BPK|IBU|NYONYA|NY|MR|MRS|MS|SDR|SAUDARA|SAUDARI)\b\.?/ /o', -1, _part_);
            _part_ = prxchange('s/(?i)\b(?:S\.?KOM|S\.?E|S\.?T|S\.?H|S\.?SI|M\.?M|M\.?B\.?A|DRS?|DRA?|IR|PROF|PH\.?D|HJ?)\b\.?/ /o', -1, _part_);
            _part_ = compress(_part_, '()-.');
            _part_ = prxchange('s/\s{2,}/ /o', -1, strip(_part_));
            _part_ = strip(_part_);
            _up_part_ = upcase(_part_);

            if lengthn(_part_) > 2 and not (_up_part_ in ('AND', 'OR', 'THE', 'OF', 'AS', 'NON FOOD', 'DIV')) then do;
                _in_idx_ + 1;
                _inclean_{_in_idx_} = _up_part_;
            end;
        end;
        if _in_idx_ = 0 then INSURED_CLEAN_1 = upcase(insured_ori);
    end;

    /* Sinkronisasi variabel alias untuk downstream matching */
    clean_polis_1   = POLICY_CLEAN_1; clean_polis_2   = POLICY_CLEAN_2;
    clean_polis_3   = POLICY_CLEAN_3; clean_polis_4   = POLICY_CLEAN_4; clean_polis_5   = POLICY_CLEAN_5;
    clean_sertif_1  = SERTIF_CLEAN_1; clean_sertif_2  = SERTIF_CLEAN_2; clean_sertif_3  = SERTIF_CLEAN_3;
    clean_slip_1    = SLIP_CLEAN_1;   clean_slip_2    = SLIP_CLEAN_2;   clean_slip_3    = SLIP_CLEAN_3;
    clean_slip_4    = SLIP_CLEAN_4;   clean_slip_5    = SLIP_CLEAN_5;
    clean_insured_1 = INSURED_CLEAN_1; clean_insured_2 = INSURED_CLEAN_2; clean_insured_3 = INSURED_CLEAN_3;
    clean_insured_4 = INSURED_CLEAN_4; clean_insured_5 = INSURED_CLEAN_5;

    POLIS_CLN_1 = POLICY_CLEAN_1; POLIS_CLN_2 = POLICY_CLEAN_2; POLIS_CLN_3 = POLICY_CLEAN_3;
    POLIS_CLN_4 = POLICY_CLEAN_4; POLIS_CLN_5 = POLICY_CLEAN_5;
    SLIP_NO_CLN_1 = SLIP_CLEAN_1; SLIP_NO_CLN_2 = SLIP_CLEAN_2; SLIP_NO_CLN_3 = SLIP_CLEAN_3;
    SLIP_NO_CLN_4 = SLIP_CLEAN_4; SLIP_NO_CLN_5 = SLIP_CLEAN_5;
    INSURED_CLN_1 = INSURED_CLEAN_1; INSURED_CLN_2 = INSURED_CLEAN_2; INSURED_CLN_3 = INSURED_CLEAN_3;
    INSURED_CLN_4 = INSURED_CLEAN_4; INSURED_CLN_5 = INSURED_CLEAN_5;

    /* Hapus karakter kontrol XML */
    array _str_vars_{*} _character_;
    do _sci_ = 1 to dim(_str_vars_);
        if not missing(_str_vars_{_sci_}) then
            _str_vars_{_sci_} = compress(_str_vars_{_sci_}, '0001020304050607080B0C0E0F101112131415161718191A1B1C1D1E1F'x);
    end;

    /* SUSUN URUTAN KOLOM SESUAI PERSIS SKEMA POSTGRESQL SUSPENSE_DATA_FACUL_CLEAN (36 KOLOM) */
    retain FAC_CODE FAC_CEDANT COMP_NAME FAC_BROKER COMP_NAME2 BUSSINESS_PATNER
           FAC_INSURED INSURED_CLEAN_1 INSURED_CLEAN_2 INSURED_CLEAN_3 INSURED_CLEAN_4 INSURED_CLEAN_5
           FAC_POLICY_NO POLICY_CLEAN_1 SERTIF_CLEAN_1 POLICY_CLEAN_2 SERTIF_CLEAN_2 POLICY_CLEAN_3 SERTIF_CLEAN_3
           POLICY_CLEAN_4 POLICY_CLEAN_5
           FAC_SLIP SLIP_CLEAN_1 SLIP_CLEAN_2 SLIP_CLEAN_3 SLIP_CLEAN_4 SLIP_CLEAN_5
           FAC_CURRENCY FAC_INP_DATE FAC_COM_DATE FAC_EXP_DATE FAC_SUB_CLASS FAC_RISK FAC_DESC FAC_ACC_STS FAC_STS_SLIP;

    drop _broker_ _p_raw_ _pos1_ _len1_ _pos2_ _len2_ _sd1_ _sd2_ _tmp_ _cval_
         _pwork_ _n_plus_ _p_base_prefix_ _c_idx_ _pi_ _ptok_ _first_t_ _p_pfx_ _n_d_ _dtok_
         _ntok_ _i_ _tok_ _base_ _suf_
         _swork_ _n_spl_ _s_base_prefix_ _s_idx_ _sj_ _s_tok_ _s_first_t_ _s_pfx_ _s_nd_ _s_dtok_
         _sntok_ _j_ _stok_ _sbase_ _ssuf_
         _ins_work_ _n_ins_ _in_idx_ _k_ _part_ _up_part_ _sci_;
run;

/* 1.3. Validasi Hasil Cleaning Facul */
data _null_;
    dsid = open('casuser.lippo_facul_clean');
    if dsid > 0 then do;
        nobs = attrn(dsid, 'nobs');
        nvar = attrn(dsid, 'nvars');
        rc   = close(dsid);
        if nobs > 0 then
            put '[VALIDATION PASS] casuser.lippo_facul_clean: ' nobs comma12. ' rows, ' nvar ' cols.';
        else do;
            put 'ERROR: [VALIDATION FAIL] casuser.lippo_facul_clean kosong (0 rows)!';
            abort cancel;
        end;
    end;
    else do;
        put 'ERROR: [VALIDATION FAIL] casuser.lippo_facul_clean tidak dapat dibuka!';
        abort cancel;
    end;
run;


/* =============================================================================
   PHASE 2: LOAD & CLEAN DATA OUTSTANDING BALANCE (SUSPESE_OSBAL_ORI_2026)
   ============================================================================= */

/* 2.1. Bersihkan tabel temporary lama & Tarik Data Mentah via PROC FEDSQL */
proc delete data=casuser.osbal_raw; run;
proc delete data=casuser.lippo_osbal_clean; run;

proc fedsql sessref=mycas;
    create table casuser.osbal_raw as
    select * from FSI_DB.SUSPESE_OSBAL_ORI_2026
    where "CCOS_COMP_NAME" like 'PT LIPPO GENERAL INSURANCE%';
quit;

/* 2.2. Cleaning Data OSBAL In-Memory di CAS (100% Logika Python Lippo) */
data casuser.lippo_osbal_clean;
    set casuser.osbal_raw;

    length INSURED_CLEAN_1-INSURED_CLEAN_5 $500
           POLICY_CLEAN_1-POLICY_CLEAN_5 $500
           SERTIF_CLEAN_1-SERTIF_CLEAN_5 $50
           SLIP_CLEAN_1-SLIP_CLEAN_5 $500
           clean_polis_1-clean_polis_5 $500
           clean_slip_1-clean_slip_5 $500
           clean_insured_1-clean_insured_5 $500
           clean_sertif_1-clean_sertif_5 $50
           POLIS_CLN_1-POLIS_CLN_5 $500
           SLIP_NO_CLN_1-SLIP_NO_CLN_5 $500
           INSURED_CLN_1-INSURED_CLN_5 $500
           polis_ori slip_ori insured_ori $500;

    /* Inisialisasi awal */
    array _pclean_{5} $500 POLICY_CLEAN_1-POLICY_CLEAN_5;
    array _sclean_{5} $50  SERTIF_CLEAN_1-SERTIF_CLEAN_5;
    array _slclean_{5} $500 SLIP_CLEAN_1-SLIP_CLEAN_5;
    array _inclean_{5} $500 INSURED_CLEAN_1-INSURED_CLEAN_5;

    do _i_ = 1 to 5;
        _pclean_{_i_} = '';
        _sclean_{_i_} = '';
        _slclean_{_i_} = '';
        _inclean_{_i_} = '';
    end;

    /* Ambil kolom asli sebagai string */
    polis_ori   = strip(cats(FAC_POLICY_NO));
    slip_ori    = strip(cats(FAC_SLIP));
    insured_ori = strip(cats(FAC_INSURED));

    /* A. Sertifikat Lippo: Prioritas CLSDT_SERTF_NO, fallback ke range/sertif FAC */
    _clsdt_s_ = strip(cats(CLSDT_SERTF_NO));
    if prxmatch('/\.0$/', _clsdt_s_) then _clsdt_s_ = substr(_clsdt_s_, 1, lengthn(_clsdt_s_) - 2);

    if not missing(_clsdt_s_) and not (_clsdt_s_ in ('', '0', '-', '#')) and prxmatch('/^\d{1,6}$/', _clsdt_s_) then do;
        SERTIF_CLEAN_1 = put(input(_clsdt_s_, 8.), z6.);
    end;
    else do;
        _p_raw_ = polis_ori;
        retain _rx_osb_sd1_ _rx_osb_sd2_ 0;
        if _rx_osb_sd1_ = 0 then _rx_osb_sd1_ = prxparse('/^(\d{10,})\s*-\s*(\d+)\s*(?:S\/D|SD)\s*(\d+)$/oi');
        if _rx_osb_sd2_ = 0 then _rx_osb_sd2_ = prxparse('/^(\d{10,})\s+(\d+)\s*(?:S\/D|SD)\s*(\d+)$/oi');

        if prxmatch(_rx_osb_sd1_, _p_raw_) then do;
            call prxposn(_rx_osb_sd1_, 2, _pos1_, _len1_);
            call prxposn(_rx_osb_sd1_, 3, _pos2_, _len2_);
            if _pos1_ > 0 and _pos2_ > 0 then do;
                _st_ = input(substr(_p_raw_, _pos1_, _len1_), 8.);
                _en_ = input(substr(_p_raw_, _pos2_, _len2_), 8.);
                SERTIF_CLEAN_1 = put(_st_, z6.) || ' SD ' || put(_en_, z6.);
            end;
        end;
        else if prxmatch(_rx_osb_sd2_, _p_raw_) then do;
            call prxposn(_rx_osb_sd2_, 2, _pos1_, _len1_);
            call prxposn(_rx_osb_sd2_, 3, _pos2_, _len2_);
            if _pos1_ > 0 and _pos2_ > 0 then do;
                _st_ = input(substr(_p_raw_, _pos1_, _len1_), 8.);
                _en_ = input(substr(_p_raw_, _pos2_, _len2_), 8.);
                SERTIF_CLEAN_1 = put(_st_, z6.) || ' SD ' || put(_en_, z6.);
            end;
        end;
    end;

    /* B. Clean Polis Lippo: Prioritas CLSDT atas FAC */
    _target_p_ = strip(cats(CLSDT_POLICY_NO));
    _use_clsdt_p_ = 0;
    if not missing(_target_p_) and not (_target_p_ in ('0', '-')) then do;
        if prxmatch('/\d{10,}/', _target_p_) then _use_clsdt_p_ = 1;
        else if prxmatch('/\d{7,}/', _target_p_) and not prxmatch('/TBA$/oi', _target_p_) then _use_clsdt_p_ = 1;
    end;

    if _use_clsdt_p_ = 0 or missing(_target_p_) then
        _target_p_ = polis_ori;

    if not missing(_target_p_) and not (_target_p_ in ('0', '-')) then do;
        _pwork_ = _target_p_;

        if index(_pwork_, '+') > 0 then do;
            _n_plus_ = countw(_pwork_, '+');
            _p_base_prefix_ = '';
            _c_idx_ = 0;
            do _pi_ = 1 to _n_plus_;
                if _c_idx_ >= 5 then leave;
                _ptok_ = strip(scan(_pwork_, _pi_, '+'));
                _ptok_ = prxchange('s/[^\w\-]+//o', -1, _ptok_);
                if _pi_ = 1 then do;
                    if lengthn(_ptok_) >= 10 then do;
                        _c_idx_ + 1;
                        _pclean_{_c_idx_} = _ptok_;
                        _p_base_prefix_ = substr(_ptok_, 1, lengthn(_ptok_) - 3);
                    end;
                    else if lengthn(_ptok_) >= 6 then do;
                        _c_idx_ + 1;
                        _pclean_{_c_idx_} = _ptok_;
                    end;
                end;
                else do;
                    if lengthn(_ptok_) in (2, 3, 4) and not missing(_p_base_prefix_) then do;
                        _c_idx_ + 1;
                        _pclean_{_c_idx_} = cats(_p_base_prefix_, put(input(_ptok_, 8.), z3.));
                    end;
                    else if lengthn(_ptok_) >= 6 then do;
                        _c_idx_ + 1;
                        _pclean_{_c_idx_} = _ptok_;
                    end;
                end;
            end;
        end;
        else if prxmatch('/^\d{8,}-\d{3}(?:-\d{3})+$/', strip(_pwork_)) then do;
            _first_t_ = scan(strip(_pwork_), 1, '-');
            _p_pfx_   = substr(_first_t_, 1, lengthn(_first_t_) - 3);
            _n_d_     = countw(strip(_pwork_), '-');
            _c_idx_   = 0;
            do _pi_ = 1 to _n_d_;
                if _c_idx_ >= 5 then leave;
                _dtok_ = strip(scan(strip(_pwork_), _pi_, '-'));
                _c_idx_ + 1;
                if _pi_ = 1 then _pclean_{_c_idx_} = _dtok_;
                else _pclean_{_c_idx_} = cats(_p_pfx_, _dtok_);
            end;
        end;
        else do;
            _pwork_ = prxchange('s/[\/\+,;&]+/ /o', -1, _pwork_);
            _pwork_ = prxchange('s/\s{2,}/ /o', -1, strip(_pwork_));
            _ntok_  = countw(strip(_pwork_), ' ');
            _c_idx_ = 0;

            do _i_ = 1 to _ntok_;
                if _c_idx_ >= 5 then leave;
                _tok_ = strip(scan(strip(_pwork_), _i_, ' '));
                _tok_ = prxchange('s/^[\-\/]+|[\-\/]+$//o', -1, _tok_);

                if lengthn(_tok_) >= 5 and anydigit(_tok_) > 0 and not prxmatch('/TBA$/oi', _tok_) then do;
                    if lengthn(compress(_tok_, '0123456789', 'k')) >= 10 then do;
                        if index(_tok_, '-') > 0 then do;
                            _base_ = scan(_tok_, 1, '-');
                            _suf_  = scan(_tok_, 2, '-');
                            if lengthn(compress(_base_, '0123456789', 'k')) >= 10 and lengthn(compress(_suf_, '0123456789', 'k')) <= 6 then
                                _tok_ = _base_;
                        end;
                    end;
                    _c_idx_ + 1;
                    _pclean_{_c_idx_} = _tok_;
                end;
            end;
        end;
        if _c_idx_ = 0 then POLICY_CLEAN_1 = _target_p_;
    end;

    /* C. Clean Slip Lippo: Prioritas CLSDT atas FAC */
    _target_s_ = strip(cats(CLSDT_SLIP_NO));
    _use_clsdt_s_ = 0;
    if not missing(_target_s_) and not (_target_s_ in ('0', '-')) then do;
        if prxmatch('/\d{10,}/', _target_s_) then _use_clsdt_s_ = 1;
        else if prxmatch('/\d{7,}/', _target_s_) and not prxmatch('/TBA$/oi', _target_s_) then _use_clsdt_s_ = 1;
    end;

    if _use_clsdt_s_ = 0 or missing(_target_s_) then
        _target_s_ = slip_ori;

    if not missing(_target_s_) and not (_target_s_ in ('0', '-')) then do;
        _swork_ = _target_s_;
        _swork_ = prxchange('s/\b(?:SUMMARY|BORDER[OA]?|BORDRO|SINGGLESHIPMENT|VARIOUS)\b/ /oi', -1, _swork_);
        _swork_ = prxchange('s/\b(?:JANUARI|FEBRUARI|MARET|APRIL|MEI|JUNI|JULI|AGUSTUS|SEPTEMBER|OKTOBER|NOVEMBER|DESEMBER)\b/ /oi', -1, _swork_);
        _swork_ = prxchange('s/\b(?:JANUARY|FEBRUARY|MARCH|APRIL|MAY|JUNE|JULY|AUGUST|SEPTEMBER|OCTOBER|NOVEMBER|DECEMBER)\b/ /oi', -1, _swork_);
        _swork_ = prxchange('s/\b(?:IDR|USD|ENG|P1|P2|P3|P73|NEW|20\d{2})\b/ /oi', -1, _swork_);

        if index(_swork_, '+') > 0 then do;
            _n_spl_ = countw(_swork_, '+');
            _s_base_prefix_ = '';
            _s_idx_ = 0;
            do _sj_ = 1 to _n_spl_;
                if _s_idx_ >= 5 then leave;
                _s_tok_ = strip(scan(_swork_, _sj_, '+'));
                _s_tok_ = prxchange('s/[^\w\-]+//o', -1, _s_tok_);
                if _sj_ = 1 then do;
                    if lengthn(_s_tok_) >= 10 then do;
                        _s_idx_ + 1;
                        _slclean_{_s_idx_} = _s_tok_;
                        _s_base_prefix_ = substr(_s_tok_, 1, lengthn(_s_tok_) - 3);
                    end;
                    else if lengthn(_s_tok_) >= 6 then do;
                        _s_idx_ + 1;
                        _slclean_{_s_idx_} = _s_tok_;
                    end;
                end;
                else do;
                    if lengthn(_s_tok_) in (2, 3, 4) and not missing(_s_base_prefix_) then do;
                        _s_idx_ + 1;
                        _slclean_{_s_idx_} = cats(_s_base_prefix_, put(input(_s_tok_, 8.), z3.));
                    end;
                    else if lengthn(_s_tok_) >= 6 then do;
                        _s_idx_ + 1;
                        _slclean_{_s_idx_} = _s_tok_;
                    end;
                end;
            end;
        end;
        else if prxmatch('/^\d{8,}-\d{3}(?:-\d{3})+$/', strip(_swork_)) then do;
            _s_first_t_ = scan(strip(_swork_), 1, '-');
            _s_pfx_     = substr(_s_first_t_, 1, lengthn(_s_first_t_) - 3);
            _s_nd_      = countw(strip(_swork_), '-');
            _s_idx_     = 0;
            do _sj_ = 1 to _s_nd_;
                if _s_idx_ >= 5 then leave;
                _s_dtok_ = strip(scan(strip(_swork_), _sj_, '-'));
                _s_idx_ + 1;
                if _sj_ = 1 then _slclean_{_s_idx_} = _s_dtok_;
                else _slclean_{_s_idx_} = cats(_s_pfx_, _s_dtok_);
            end;
        end;
        else do;
            _swork_ = prxchange('s/[\/\+,;&]+/ /o', -1, _swork_);
            _swork_ = prxchange('s/\s{2,}/ /o', -1, strip(_swork_));
            _sntok_ = countw(strip(_swork_), ' ');
            _s_idx_ = 0;

            do _j_ = 1 to _sntok_;
                if _s_idx_ >= 5 then leave;
                _stok_ = strip(scan(strip(_swork_), _j_, ' '));
                _stok_ = prxchange('s/^[\-\/]+|[\-\/]+$//o', -1, _stok_);
                if lengthn(_stok_) >= 7 and anydigit(_stok_) > 0 and not prxmatch('/^\d{4}$/', _stok_) then do;
                    if index(_stok_, '-') > 0 then do;
                        _sbase_ = scan(_stok_, 1, '-');
                        _ssuf_  = scan(_stok_, 2, '-');
                        if lengthn(_sbase_) > lengthn(_ssuf_) and lengthn(_ssuf_) <= 6 then
                            _stok_ = _sbase_;
                    end;
                    _s_idx_ + 1;
                    _slclean_{_s_idx_} = _stok_;
                end;
            end;
        end;
        if _s_idx_ = 0 then SLIP_CLEAN_1 = _target_s_;
    end;

    /* D. Clean Insured Lippo: Selalu dari FAC_INSURED */
    if not missing(insured_ori) then do;
        _ins_work_ = insured_ori;
        _ins_work_ = prxchange('s/(?i)\bAS\b\s+(?:THE\s+)?(?:PRINCIPAL|OFF-TAKER|MAINTENANCE|CONTRACTOR).*//o', -1, _ins_work_);
        _ins_work_ = prxchange('s/(?i)\bAND\s+ALL\s+SUBSIDIARI.*//o', -1, _ins_work_);
        _ins_work_ = prxchange('s/(?i)\b(?:AND|AN|OR)\s*\/\s*(?:AND|OR)\b/,/o', -1, _ins_work_);
        _ins_work_ = prxchange('s/(?i)\bAND\s+OR\b/,/o', -1, _ins_work_);
        _ins_work_ = prxchange('s/(?i)\b(?:QQ|C\/Q|Q\/Q)\b/|/o', -1, _ins_work_);
        _ins_work_ = prxchange('s/[\/,\+;]/|/o', -1, _ins_work_);
        _ins_work_ = prxchange('s/(?i)\b(?:PT\.?|CV\.?)\b/|/o', -1, _ins_work_);

        _n_ins_ = countw(strip(_ins_work_), '|');
        _in_idx_ = 0;

        do _k_ = 1 to _n_ins_;
            if _in_idx_ >= 5 then leave;
            _part_ = strip(scan(strip(_ins_work_), _k_, '|'));
            _part_ = prxchange('s/(?i)\b(?:PT|CV|TBK|\(PERSERO\)|PERSERO|LTD|PTE|INC|LLC|UD|PD|NV|BV|GMBH|BHD)\b\.?/ /o', -1, _part_);
            _part_ = prxchange('s/(?i)\b(?:BAPAK|BPK|IBU|NYONYA|NY|MR|MRS|MS|SDR|SAUDARA|SAUDARI)\b\.?/ /o', -1, _part_);
            _part_ = prxchange('s/(?i)\b(?:S\.?KOM|S\.?E|S\.?T|S\.?H|S\.?SI|M\.?M|M\.?B\.?A|DRS?|DRA?|IR|PROF|PH\.?D|HJ?)\b\.?/ /o', -1, _part_);
            _part_ = compress(_part_, '()-.');
            _part_ = prxchange('s/\s{2,}/ /o', -1, strip(_part_));
            _part_ = strip(_part_);
            _up_part_ = upcase(_part_);

            if lengthn(_part_) > 2 and not (_up_part_ in ('AND', 'OR', 'THE', 'OF', 'AS', 'NON FOOD', 'DIV')) then do;
                _in_idx_ + 1;
                _inclean_{_in_idx_} = _up_part_;
            end;
        end;
        if _in_idx_ = 0 then INSURED_CLEAN_1 = upcase(insured_ori);
    end;

    /* Sinkronisasi variabel alias untuk downstream matching */
    clean_polis_1   = POLICY_CLEAN_1; clean_polis_2   = POLICY_CLEAN_2;
    clean_polis_3   = POLICY_CLEAN_3; clean_polis_4   = POLICY_CLEAN_4; clean_polis_5   = POLICY_CLEAN_5;
    clean_sertif_1  = SERTIF_CLEAN_1; clean_sertif_2  = SERTIF_CLEAN_2; clean_sertif_3  = SERTIF_CLEAN_3;
    clean_slip_1    = SLIP_CLEAN_1;   clean_slip_2    = SLIP_CLEAN_2;   clean_slip_3    = SLIP_CLEAN_3;
    clean_slip_4    = SLIP_CLEAN_4;   clean_slip_5    = SLIP_CLEAN_5;
    clean_insured_1 = INSURED_CLEAN_1; clean_insured_2 = INSURED_CLEAN_2; clean_insured_3 = INSURED_CLEAN_3;
    clean_insured_4 = INSURED_CLEAN_4; clean_insured_5 = INSURED_CLEAN_5;

    POLIS_CLN_1 = POLICY_CLEAN_1; POLIS_CLN_2 = POLICY_CLEAN_2; POLIS_CLN_3 = POLICY_CLEAN_3;
    POLIS_CLN_4 = POLICY_CLEAN_4; POLIS_CLN_5 = POLICY_CLEAN_5;
    SLIP_NO_CLN_1 = SLIP_CLEAN_1; SLIP_NO_CLN_2 = SLIP_CLEAN_2; SLIP_NO_CLN_3 = SLIP_CLEAN_3;
    SLIP_NO_CLN_4 = SLIP_CLEAN_4; SLIP_NO_CLN_5 = SLIP_CLEAN_5;
    INSURED_CLN_1 = INSURED_CLEAN_1; INSURED_CLN_2 = INSURED_CLEAN_2; INSURED_CLN_3 = INSURED_CLEAN_3;
    INSURED_CLN_4 = INSURED_CLEAN_4; INSURED_CLN_5 = INSURED_CLEAN_5;

    /* Hapus karakter kontrol XML */
    array _str_vars_o_{*} _character_;
    do _sci_ = 1 to dim(_str_vars_o_);
        if not missing(_str_vars_o_{_sci_}) then
            _str_vars_o_{_sci_} = compress(_str_vars_o_{_sci_}, '0001020304050607080B0C0E0F101112131415161718191A1B1C1D1E1F'x);
    end;

    /* SUSUN URUTAN KOLOM SESUAI PERSIS SKEMA POSTGRESQL SUSPENSE_DATA_OSBAL_CLEAN (42 KOLOM) */
    retain CCOS_DOC_NO CCOS_DATE CCOS_REF_CODE CCOS_COMP CCOS_COMP_NAME CCOS_REF_COMP CCOS_REF_COMP_NAME
           FAC_INSURED INSURED_CLEAN_1 INSURED_CLEAN_2 INSURED_CLEAN_3 INSURED_CLEAN_4 INSURED_CLEAN_5
           CCOS_CURR CCOS_OR_BAL CCOS_BAL_DUE CCOS_OR_BAL_IN_IDR CCOS_BAL_DUE_IN_IDR
           FAC_COM_DATE FAC_EXP_DATE FAC_DUE_DATES FAC_SUB_CLASS
           FAC_POLICY_NO POLICY_CLEAN_1 SERTIF_CLEAN_1 POLICY_CLEAN_2 SERTIF_CLEAN_2 POLICY_CLEAN_3 SERTIF_CLEAN_3
           POLICY_CLEAN_4 POLICY_CLEAN_5
           FAC_SLIP SLIP_CLEAN_1 SLIP_CLEAN_2 SLIP_CLEAN_3 SLIP_CLEAN_4 SLIP_CLEAN_5
           CLASS_CODE CLASS_NAME CLSDT_POLICY_NO CLSDT_SLIP_NO CLSDT_SERTF_NO;

    drop _clsdt_s_ _p_raw_ _pos1_ _len1_ _pos2_ _len2_ _st_ _en_
         _target_p_ _use_clsdt_p_ _pwork_ _n_plus_ _p_base_prefix_ _c_idx_ _pi_ _ptok_ _first_t_ _p_pfx_ _n_d_ _dtok_
         _ntok_ _i_ _tok_ _base_ _suf_
         _target_s_ _use_clsdt_s_ _swork_ _n_spl_ _s_base_prefix_ _s_idx_ _sj_ _s_tok_ _s_first_t_ _s_pfx_ _s_nd_ _s_dtok_
         _sntok_ _j_ _stok_ _sbase_ _ssuf_
         _ins_work_ _n_ins_ _in_idx_ _k_ _part_ _up_part_ _sci_;
run;

/* 2.3. Validasi Hasil Cleaning OSBAL */
data _null_;
    dsid = open('casuser.lippo_osbal_clean');
    if dsid > 0 then do;
        nobs = attrn(dsid, 'nobs');
        nvar = attrn(dsid, 'nvars');
        rc   = close(dsid);
        if nobs > 0 then
            put '[VALIDATION PASS] casuser.lippo_osbal_clean: ' nobs comma12. ' rows, ' nvar ' cols.';
        else do;
            put 'ERROR: [VALIDATION FAIL] casuser.lippo_osbal_clean kosong (0 rows)!';
            abort cancel;
        end;
    end;
    else do;
        put 'ERROR: [VALIDATION FAIL] casuser.lippo_osbal_clean tidak dapat dibuka!';
        abort cancel;
    end;
run;


/* =============================================================================
   PHASE 3: LOAD & CLEAN DATA SUSPENSE (SUSPESE_SUSPENSE_ORI_2026)
   ============================================================================= */

/* 3.1. Bersihkan tabel temporary lama & Tarik Data Mentah via PROC FEDSQL */
proc delete data=casuser.suspend_raw; run;
proc delete data=casuser.lippo_suspend_clean; run;

proc fedsql sessref=mycas;
    create table casuser.suspend_raw as
    select * from FSI_DB.SUSPESE_SUSPENSE_ORI_2026
    where ("CEDANT SHRT NAME" like 'LIPPO%' or "CEDANT NAME" like 'PT LIPPO%') and "STATUS" like 'SUSPENSE%';
quit;

/* 3.2. Cleaning Data Suspend In-Memory di CAS (100% Logika Python Lippo) */
data casuser.lippo_suspend_clean;
    set casuser.suspend_raw;

    length INSURED_CLEAN_1-INSURED_CLEAN_5 $500
           POLICY_CLEAN_1-POLICY_CLEAN_5 $500
           SERTIF_CLEAN_1-SERTIF_CLEAN_5 $50
           SLIP_CLEAN_1-SLIP_CLEAN_5 $500
           clean_polis_1-clean_polis_5 $500
           clean_slip_1-clean_slip_5 $500
           clean_insured_1-clean_insured_5 $500
           clean_sertif_1-clean_sertif_5 $50
           POLIS_CLN_1-POLIS_CLN_5 $500
           SLIP_NO_CLN_1-SLIP_NO_CLN_5 $500
           INSURED_CLN_1-INSURED_CLN_5 $500
           POLIS_CLN $500
           SLIP_NO_CLN $500
           INSURED_CLEAN $500
           polis_ori slip_ori insured_ori $500;

    /* Inisialisasi awal */
    array _pclean_{5} $500 POLICY_CLEAN_1-POLICY_CLEAN_5;
    array _sclean_{5} $50  SERTIF_CLEAN_1-SERTIF_CLEAN_5;
    array _slclean_{5} $500 SLIP_CLEAN_1-SLIP_CLEAN_5;
    array _inclean_{5} $500 INSURED_CLEAN_1-INSURED_CLEAN_5;

    do _i_ = 1 to 5;
        _pclean_{_i_} = '';
        _sclean_{_i_} = '';
        _slclean_{_i_} = '';
        _inclean_{_i_} = '';
    end;

    /* Ambil kolom asli */
    polis_ori   = strip(cats(POLIS));
    slip_ori    = strip(cats("SLIP NO"n));
    insured_ori = strip(cats(INSURED));

    /* A. Extract Sertifikat Suspend Lippo dari POLIS (100% Python extract_cert_from_polis) */
    _p_raw_ = polis_ori;
    retain _rx_sd_s_ _rx_cert_s_ 0;
    if _rx_sd_s_   = 0 then _rx_sd_s_   = prxparse('/(\d{1,6})\s*S\/D\s*(\d{1,6})/oi');
    if _rx_cert_s_ = 0 then _rx_cert_s_ = prxparse('/[-\s]+\s*(\d{1,6})(?:[^0-9]|$)/o');

    if prxmatch(_rx_sd_s_, _p_raw_) then do;
        call prxposn(_rx_sd_s_, 1, _pos1_, _len1_);
        call prxposn(_rx_sd_s_, 2, _pos2_, _len2_);
        if _pos1_ > 0 and _pos2_ > 0 then do;
            _sd1_ = input(substr(_p_raw_, _pos1_, _len1_), 8.);
            _sd2_ = input(substr(_p_raw_, _pos2_, _len2_), 8.);
            if _sd1_ > _sd2_ then do; _tmp_ = _sd1_; _sd1_ = _sd2_; _sd2_ = _tmp_; end;
            if _sd1_ <= _sd2_ then SERTIF_CLEAN_1 = put(_sd1_, z6.);
            if _sd1_ + 1 <= _sd2_ then SERTIF_CLEAN_2 = put(_sd1_ + 1, z6.);
            if _sd1_ + 2 <= _sd2_ then SERTIF_CLEAN_3 = put(_sd1_ + 2, z6.);
        end;
    end;
    else if prxmatch(_rx_cert_s_, _p_raw_) then do;
        call prxposn(_rx_cert_s_, 1, _pos1_, _len1_);
        if _pos1_ > 0 and _len1_ > 0 then do;
            _cval_ = strip(substr(_p_raw_, _pos1_, _len1_));
            if prxmatch('/^\d{1,6}$/', _cval_) then
                SERTIF_CLEAN_1 = put(input(_cval_, 8.), z6.);
        end;
    end;

    /* B. Clean Polis Suspend Lippo: Delimiter dash / strip suffix */
    if not missing(polis_ori) and not (polis_ori in ('0', '-')) then do;
        _psus_ = polis_ori;
        _psus_ = prxchange('s/-EXT\(\d+\)//oi', -1, _psus_);
        _prev_ = '';
        do while (_psus_ ne _prev_);
            _prev_ = _psus_;
            _psus_ = prxchange('s/-(?:\d+(?:\/\d+)?)$//oi', -1, strip(_psus_));
            _psus_ = strip(_psus_);
        end;

        /* Jika ada multiple polis terpisah dash e.g. "1201101700185 - 1201101903855" */
        _n_pd_ = countw(_psus_, '-');
        if _n_pd_ > 1 then do;
            _c_idx_ = 0;
            do _pi_ = 1 to _n_pd_;
                if _c_idx_ >= 5 then leave;
                _dtok_ = strip(scan(_psus_, _pi_, '-'));
                if lengthn(compress(_dtok_, '', 'kd')) >= 8 then do;
                    _c_idx_ + 1;
                    _pclean_{_c_idx_} = _dtok_;
                end;
            end;
            if _c_idx_ = 0 then POLICY_CLEAN_1 = _psus_;
        end;
        else do;
            if _psus_ ne '' then POLICY_CLEAN_1 = _psus_;
            else POLICY_CLEAN_1 = polis_ori;
        end;
    end;

    /* C. Clean Slip Suspend Lippo: Strip suffix - Cancel / -5/3 */
    if not missing(slip_ori) and not (slip_ori in ('0', '-')) then do;
        _sl_s_ = slip_ori;
        _sl_s_ = prxchange('s/\s*-\s*Cancel.*//oi', -1, _sl_s_);
        _sl_s_ = prxchange('s/\s*-\s*\d+\/\d+.*//oi', -1, _sl_s_);
        _sl_s_ = compress(_sl_s_, '.,');
        _sl_s_ = strip(prxchange('s/\s+/ /', -1, _sl_s_));
        if _sl_s_ ne '' then SLIP_CLEAN_1 = _sl_s_;
        else SLIP_CLEAN_1 = slip_ori;
    end;

    /* D. Clean Insured Suspend Lippo: Hapus polis/slip, entitas hukum, honorifics */
    if not missing(insured_ori) then do;
        _ins_s_ = insured_ori;
        if not missing(polis_ori) and polis_ori ne '-' then _ins_s_ = tranwrd(_ins_s_, trim(polis_ori), '');
        if not missing(slip_ori)  and slip_ori  ne '-' then _ins_s_ = tranwrd(_ins_s_, trim(slip_ori), '');

        _ins_s_ = prxchange('s/(?i)\bAS\b\s+(?:THE\s+)?(?:PRINCIPAL|OFF-TAKER|MAINTENANCE|CONTRACTOR).*//o', -1, _ins_s_);
        _ins_s_ = prxchange('s/(?i)\bAND\s+ALL\s+SUBSIDIARI.*//o', -1, _ins_s_);
        _ins_s_ = prxchange('s/(?i)\bINCLUDING\s+ANY\s+SUBSIDIARY.*//o', -1, _ins_s_);
        _ins_s_ = prxchange('s/(?i)\b(?:AND|AN|OR)\s*\/\s*(?:AND|OR)\b/,/o', -1, _ins_s_);
        _ins_s_ = prxchange('s/(?i)\bAND\s+OR\b/,/o', -1, _ins_s_);
        _ins_s_ = prxchange('s/(?i)\b(?:QQ|C\/Q|Q\/Q)\b/|/o', -1, _ins_s_);
        _ins_s_ = prxchange('s/[\/,\+;]/|/o', -1, _ins_s_);
        _ins_s_ = prxchange('s/(?i)\b(?:PT\.?|CV\.?)\b/|/o', -1, _ins_s_);

        _n_ins_ = countw(strip(_ins_s_), '|');
        _in_idx_ = 0;

        do _k_ = 1 to _n_ins_;
            if _in_idx_ >= 5 then leave;
            _part_ = strip(scan(strip(_ins_s_), _k_, '|'));
            _part_ = prxchange('s/(?i)\b(?:PT|CV|TBK|\(PERSERO\)|PERSERO|LTD|PTE|INC|LLC|UD|PD|NV|BV|GMBH|BHD)\b\.?/ /o', -1, _part_);
            _part_ = prxchange('s/(?i)\b(?:BAPAK|BPK|IBU|NYONYA|NY|MR|MRS|MS|SDR|SAUDARA|SAUDARI)\b\.?/ /o', -1, _part_);
            _part_ = prxchange('s/(?i)\b(?:S\.?KOM|S\.?E|S\.?T|S\.?H|S\.?SI|M\.?M|M\.?B\.?A|DRS?|DRA?|IR|PROF|PH\.?D|HJ?)\b\.?/ /o', -1, _part_);
            _part_ = compress(_part_, '()-.');
            _part_ = prxchange('s/\s{2,}/ /o', -1, strip(_part_));
            _part_ = strip(_part_);
            _up_part_ = upcase(_part_);

            if lengthn(_part_) > 2 and not (_up_part_ in ('AND', 'OR', 'THE', 'OF', 'AS', 'NON FOOD', 'DIV')) then do;
                _in_idx_ + 1;
                _inclean_{_in_idx_} = _up_part_;
            end;
        end;
        if _in_idx_ = 0 then INSURED_CLEAN_1 = upcase(insured_ori);
    end;

    /* Sinkronisasi variabel alias untuk kompatibilitas downstream matching */
    clean_polis_1   = POLICY_CLEAN_1; clean_polis_2   = POLICY_CLEAN_2;
    clean_polis_3   = POLICY_CLEAN_3; clean_polis_4   = POLICY_CLEAN_4; clean_polis_5   = POLICY_CLEAN_5;
    clean_sertif_1  = SERTIF_CLEAN_1; clean_sertif_2  = SERTIF_CLEAN_2; clean_sertif_3  = SERTIF_CLEAN_3;
    clean_slip_1    = SLIP_CLEAN_1;   clean_slip_2    = SLIP_CLEAN_2;   clean_slip_3    = SLIP_CLEAN_3;
    clean_slip_4    = SLIP_CLEAN_4;   clean_slip_5    = SLIP_CLEAN_5;
    clean_insured_1 = INSURED_CLEAN_1; clean_insured_2 = INSURED_CLEAN_2; clean_insured_3 = INSURED_CLEAN_3;
    clean_insured_4 = INSURED_CLEAN_4; clean_insured_5 = INSURED_CLEAN_5;

    POLIS_CLN       = POLICY_CLEAN_1;
    SLIP_NO_CLN     = SLIP_CLEAN_1;
    INSURED_CLEAN   = INSURED_CLEAN_1;

    POLIS_CLN_1 = POLICY_CLEAN_1; POLIS_CLN_2 = POLICY_CLEAN_2; POLIS_CLN_3 = POLICY_CLEAN_3;
    POLIS_CLN_4 = POLICY_CLEAN_4; POLIS_CLN_5 = POLICY_CLEAN_5;
    SLIP_NO_CLN_1 = SLIP_CLEAN_1; SLIP_NO_CLN_2 = SLIP_CLEAN_2; SLIP_NO_CLN_3 = SLIP_CLEAN_3;
    SLIP_NO_CLN_4 = SLIP_CLEAN_4; SLIP_NO_CLN_5 = SLIP_CLEAN_5;
    INSURED_CLN_1 = INSURED_CLEAN_1; INSURED_CLN_2 = INSURED_CLEAN_2; INSURED_CLN_3 = INSURED_CLEAN_3;
    INSURED_CLN_4 = INSURED_CLEAN_4; INSURED_CLN_5 = INSURED_CLEAN_5;

    /* Hapus karakter kontrol XML */
    array _str_vars_s_{*} _character_;
    do _sci_ = 1 to dim(_str_vars_s_);
        if not missing(_str_vars_s_{_sci_}) then
            _str_vars_s_{_sci_} = compress(_str_vars_s_{_sci_}, '0001020304050607080B0C0E0F101112131415161718191A1B1C1D1E1F'x);
    end;

    /* SUSUN URUTAN KOLOM SUSPENSE CLEAN */
    retain "RECEIPT NO"n "CREDIT NOTES"n "DETAIL RINCIAN NO"n "RECEIPT DATE"n "CEDANT NAME"n "CEDANT SHRT NAME"n
           INSURED INSURED_CLEAN_1 INSURED_CLEAN_2 INSURED_CLEAN_3 INSURED_CLEAN_4 INSURED_CLEAN_5
           "CURR ORI"n "AMOUNT ORI"n "CURR PAY"n "AMOUNT PAY"n
           POLIS POLICY_CLEAN_1 SERTIF_CLEAN_1 POLICY_CLEAN_2 SERTIF_CLEAN_2 POLICY_CLEAN_3 SERTIF_CLEAN_3
           POLICY_CLEAN_4 POLICY_CLEAN_5
           "SLIP NO"n SLIP_CLEAN_1 SLIP_CLEAN_2 SLIP_CLEAN_3 SLIP_CLEAN_4 SLIP_CLEAN_5
           "DESC 1"n "DESC 2"n "DESC 3"n "DESC 4"n STATUS REC_TYPE "CEK FASE"n;

    drop _p_raw_ _pos1_ _len1_ _pos2_ _len2_ _sd1_ _sd2_ _tmp_ _cval_
         _psus_ _prev_ _n_pd_ _c_idx_ _pi_ _dtok_ _sl_s_
         _ins_s_ _n_ins_ _in_idx_ _k_ _part_ _up_part_ _sci_;
run;

/* 3.3. Validasi Hasil Cleaning Suspend */
data _null_;
    dsid = open('casuser.lippo_suspend_clean');
    if dsid > 0 then do;
        nobs = attrn(dsid, 'nobs');
        nvar = attrn(dsid, 'nvars');
        rc   = close(dsid);
        if nobs > 0 then
            put '[VALIDATION PASS] casuser.lippo_suspend_clean: ' nobs comma12. ' rows, ' nvar ' cols.';
        else do;
            put 'ERROR: [VALIDATION FAIL] casuser.lippo_suspend_clean kosong (0 rows)!';
            abort cancel;
        end;
    end;
    else do;
        put 'ERROR: [VALIDATION FAIL] casuser.lippo_suspend_clean tidak dapat dibuka!';
        abort cancel;
    end;
run;


/* =============================================================================
   PHASE 4: PROMOTE KE CASLIB PUBLIC (SPESIFIK LIPPO & GLOBAL MULTI-CEDANT)
   (Mengikuti Pola Senior Developer: PROC DELETE + DATA STEP promote=yes)
   ============================================================================= */

/* 4.1. Promosikan Tabel Spesifik Lippo */
proc delete data=public.LIPPO_OSBAL_CLEAN; run;
data public.LIPPO_OSBAL_CLEAN (promote=yes);
    set casuser.lippo_osbal_clean;
run;

proc delete data=public.LIPPO_FACUL_CLEAN; run;
data public.LIPPO_FACUL_CLEAN (promote=yes);
    set casuser.lippo_facul_clean;
run;

proc delete data=public.LIPPO_SUSPEND_CLEAN; run;
data public.LIPPO_SUSPEND_CLEAN (promote=yes);
    set casuser.lippo_suspend_clean;
run;

/* 4.2. Promosikan Tabel Global Multi-Cedant (Pola Filter & Merge: Aman Multi-Cedant) */

/* --- A. OSBAL GLOBAL (SUSPENSE_OSBAL_CLEAN) --- */
%if %sysfunc(exist(public.SUSPENSE_OSBAL_CLEAN)) %then %do;
    /* 1. Ambil data cedant lain yang sudah ada di Public (misal ACA, Astra) */
    data casuser._other_osbal;
        set public.SUSPENSE_OSBAL_CLEAN;
        where "CCOS_COMP_NAME"n not like 'PT LIPPO GENERAL INSURANCE%';
    run;

    /* 2. Gabungkan data cedant lain + data bersih Lippo terbaru */
    data casuser._all_osbal;
        set casuser._other_osbal casuser.lippo_osbal_clean;
    run;

    proc delete data=casuser._other_osbal; run;
%end;
%else %do;
    /* Jika tabel global belum ada di Public, inisialisasi dengan data Lippo */
    data casuser._all_osbal;
        set casuser.lippo_osbal_clean;
    run;
%end;

/* 3. Refresh & promote tabel global di Public */
proc delete data=public.SUSPENSE_OSBAL_CLEAN; run;
data public.SUSPENSE_OSBAL_CLEAN (promote=yes);
    set casuser._all_osbal;
run;
proc delete data=casuser._all_osbal; run;


/* --- B. FACUL GLOBAL (SUSPENSE_FACUL_CLEAN) --- */
%if %sysfunc(exist(public.SUSPENSE_FACUL_CLEAN)) %then %do;
    /* 1. Ambil data cedant lain yang sudah ada di Public (misal ACA, Astra) */
    data casuser._other_facul;
        set public.SUSPENSE_FACUL_CLEAN;
        where "COMP_NAME"n not like 'PT LIPPO GENERAL INSURANCE%';
    run;

    /* 2. Gabungkan data cedant lain + data bersih Lippo terbaru */
    data casuser._all_facul;
        set casuser._other_facul casuser.lippo_facul_clean;
    run;

    proc delete data=casuser._other_facul; run;
%end;
%else %do;
    /* Jika tabel global belum ada di Public, inisialisasi dengan data Lippo */
    data casuser._all_facul;
        set casuser.lippo_facul_clean;
    run;
%end;

/* 3. Refresh & promote tabel global di Public */
proc delete data=public.SUSPENSE_FACUL_CLEAN; run;
data public.SUSPENSE_FACUL_CLEAN (promote=yes);
    set casuser._all_facul;
run;
proc delete data=casuser._all_facul; run;

/* 4.3. Validasi Keberhasilan Promosi ke CASLIB Public */
data _null_;
    array tbls[5] $32 _temporary_ ('LIPPO_OSBAL_CLEAN', 'LIPPO_FACUL_CLEAN', 'LIPPO_SUSPEND_CLEAN',
                                   'SUSPENSE_OSBAL_CLEAN', 'SUSPENSE_FACUL_CLEAN');
    do i = 1 to dim(tbls);
        dsname = cats('public.', tbls[i]);
        dsid   = open(dsname);
        if dsid > 0 then do;
            nobs = attrn(dsid, 'nobs');
            rc   = close(dsid);
            put '[PROMOTION PASS] ' dsname $28. ' aktif di Public (' nobs comma12. ' rows).';
        end;
        else do;
            put 'ERROR: [PROMOTION FAIL] ' dsname ' gagal dipromosikan ke CASLIB Public!';
            abort cancel;
        end;
    end;
run;


/* =============================================================================
   PHASE 5: RINGKASAN EKSEKUSI PIPELINE
   ============================================================================= */

%put [COMPLETE] ============================================================;
%put [COMPLETE] Pipeline Lippo Data Cleaning selesai dieksekusi di CAS!;
%put [COMPLETE] Tabel In-Memory Global siap di CASLIB Public:;
%put [COMPLETE]   - Public.LIPPO_OSBAL_CLEAN      (Spesifik Lippo - 42 Kolom DB);
%put [COMPLETE]   - Public.LIPPO_FACUL_CLEAN      (Spesifik Lippo - 36 Kolom DB);
%put [COMPLETE]   - Public.LIPPO_SUSPEND_CLEAN    (Spesifik Lippo);
%put [COMPLETE]   - Public.SUSPENSE_OSBAL_CLEAN   (Global Multi-Cedant Dashboard);
%put [COMPLETE]   - Public.SUSPENSE_FACUL_CLEAN   (Global Multi-Cedant Dashboard);
%put [COMPLETE] ============================================================;
