/* =============================================================================
   CLEANING_SUSPEND.SAS
   Konversi dari: cleaning_suspend.py

   Fungsi:
   - Load data/suspend.xlsx (PERHATIAN: header di baris ke-3, header=2 di pandas)
   - Filter CEDANT SHRT NAME = "CENTRAL" DAN STATUS = "SUSPENSE"
   - clean_polis: strip suffix numerik setelah dash (-01, -02/03, -EXT(n))
   - clean_slip: as-is (tidak ada transformasi)
   - clean_insured: logika BERBEDA dari facul/osbal:
       * Remove nomor polis & slip dari teks insured
       * Split berdasarkan QQ/slash/koma
       * Tidak ada entity strip seperti facul/osbal
   - Ekstrak sertifikat 6-digit dari polis_ori
   - Output Excel: data/suspend_clean_aca.xlsx
   - TIDAK ada export PostgreSQL (sesuai script original)

   Perbedaan kunci vs facul/osbal:
   1. Header Excel di baris ke-3 (datarow=4 di SAS)
   2. Filter ke-2: STATUS = "SUSPENSE"
   3. clean_polis: hanya strip suffix, tidak ada split token
   4. clean_insured: logika custom yang berbeda total
   ============================================================================= */

options mprint mlogic symbolgen;

/* ── 0. Path setup ─────────────────────────────────────────────────────────── */
%let BASE_DIR   = %str(C:\Users\berryl\Desktop\starcore\indore\production_script);
%let DATA_DIR   = %str(&BASE_DIR.\data);
%let INPUT_FILE = %str(&DATA_DIR.\suspend.xlsx);
%let OUT_FILE   = %str(&DATA_DIR.\suspend_clean_aca.xlsx);
%let MACRO_INC  = %str(&BASE_DIR.\sas_ready\macros_cleaning.sas);

/* ── 1. Include shared macros ──────────────────────────────────────────────── */
%include "&MACRO_INC.";

/* ── 2. Konstanta bisnis ──────────────────────────────────────────────────── */
%let CEDANT_COL   = %str(CEDANT SHRT NAME);
%let CEDANT_VALUE = CENTRAL;
%let STATUS_COL   = STATUS;
%let POLIS_COL    = POLIS;
%let SLIP_COL     = %str(SLIP NO);
%let INSURED_COL  = INSURED;
%let TABLE_NAME   = SUSPENSE_DATA_SUSPEND_CLEAN;

%compile_cleaning_regex();


/* =============================================================================
   STEP 1: Baca file Excel
   PENTING: suspend.xlsx punya header di baris ke-3 (0-indexed = header=2 di pandas)
   → di SAS: datarow=4 (baris data mulai dari baris ke-4, header di baris ke-3)
   ============================================================================= */
%put [1/5] Reading: &INPUT_FILE. (header di baris ke-3) ...;

proc import
    datafile="&INPUT_FILE."
    out=work.suspend_raw
    dbms=xlsx
    replace;
    getnames=yes;
    datarow=4;   /* header=2 di pandas = baris ke-3 = datarow=4 di SAS */
run;

%put       Total rows: %sysfunc(attrn(%sysfunc(open(work.suspend_raw)), nobs));


/* =============================================================================
   STEP 2: Filter CEDANT dan STATUS
   Setara dengan:
     df = df[df[CEDANT_FILTER_COL] == "CENTRAL"]
     df = df[df["STATUS"].str.upper() == "SUSPENSE"]
   ============================================================================= */
%put [2/5] Filtering CEDANT = CENTRAL, STATUS = SUSPENSE ...;

data work.suspend_filtered;
    set work.suspend_raw;
    _ced_chk_ = strip("&CEDANT_COL."n);
    _sts_chk_ = upcase(strip(&STATUS_COL.));
    if _ced_chk_ = "&CEDANT_VALUE." and _sts_chk_ = 'SUSPENSE' then output;
    drop _ced_chk_ _sts_chk_;
run;

%put       Rows setelah filter: %sysfunc(attrn(%sysfunc(open(work.suspend_filtered)), nobs));


/* =============================================================================
   STEP 3: Rename kolom utama
   Setara dengan: rename_map = {"INSURED": "insured_ori", "POLIS": "polis_ori", ...}
   ============================================================================= */
data work.suspend_renamed;
    set work.suspend_filtered;
    rename &INSURED_COL. = insured_ori
           &POLIS_COL.   = polis_ori
           "&SLIP_COL."n  = slip_ori;
run;


/* =============================================================================
   STEP 4: Macro CLEAN_POLIS khusus Suspend
   Setara dengan: clean_polis(val) di cleaning_suspend.py

   Logika BERBEDA dari facul/osbal:
   Hanya strip suffix numeric/EXT setelah dash, loop sampai tidak ada lagi.
   Contoh:
     '100030825120000155-001-002' → '100030825120000155'
     '210010421100000031-1/1'    → '210010421100000031'
     '100030825120000155-EXT(3)' → '100030825120000155'
   ============================================================================= */
%macro apply_clean_polis_suspend(in_var);
    length clean_polis_1 $500;
    array _cps_dummy{5} $500 clean_polis_2-clean_polis_5; /* kosong untuk suspend */
    do _i_ = 1 to 5; if _i_ > 1 then _cps_dummy{_i_-1} = ''; end;
    clean_polis_1 = '';

    _polis_sus_ = strip(&in_var.);
    if not missing(_polis_sus_) and _polis_sus_ ne '' then do;
        _prev_sus_ = '';
        do while (_polis_sus_ ne _prev_sus_);
            _prev_sus_ = _polis_sus_;
            /* Strip: -{digit(s)} atau -{digit(s)/digit(s)} atau -EXT({digit}) */
            _polis_sus_ = prxchange(
                's/-(?:\d+(?:\/\d+)?|EXT\(\d+\))$//oi',
                -1, strip(_polis_sus_));
            _polis_sus_ = strip(_polis_sus_);
        end;
        if _polis_sus_ ne '' then clean_polis_1 = _polis_sus_;
    end;
    drop _polis_sus_ _prev_sus_ _i_;
%mend apply_clean_polis_suspend;


/* =============================================================================
   STEP 5: Macro CLEAN_INSURED khusus Suspend
   Setara dengan: clean_insured(val, polis_ori, slip_ori) di cleaning_suspend.py

   Logika BERBEDA:
   1. Hapus nomor polis & slip dari teks insured
   2. Hapus angka dalam kurung, co. ltd, TBK, persero, LTD
   3. Split: QQ, slash, koma, dash (kecuali sebelum tahun 19xx/20xx), digit+titik
   4. Normalisasi tiap bagian (hapus AND/OR di awal/akhir, dsb)
   5. Validasi tiap bagian
   ============================================================================= */
%macro apply_clean_insured_suspend(in_var, polis_in, slip_in);
    length clean_insured_1 - clean_insured_5 $500;
    array _cis_out{5} $500 clean_insured_1-clean_insured_5;
    do _i_ = 1 to 5; _cis_out{_i_} = ''; end;
    _ins_sus_ = strip(&in_var.);
    _cis_n_   = 0;

    if missing(_ins_sus_) or _ins_sus_ = '' then goto _insus_done_;

    /* -- 5a. Hapus polis dan slip dari teks insured ------------------------- */
    /* Hapus polis_ori asli */
    if not missing(&polis_in.) then do;
        _praw_ = strip(put(&polis_in., $500.));
        if _praw_ ne '' and _praw_ ne '-' then
            _ins_sus_ = tranwrd(_ins_sus_, _praw_, '');
        /* Hapus clean_polis_1 juga */
        if clean_polis_1 ne '' and clean_polis_1 ne '-' then
            _ins_sus_ = tranwrd(_ins_sus_, clean_polis_1, '');
    end;

    /* Hapus slip_ori asli */
    if not missing(&slip_in.) then do;
        _sraw_ = strip(put(&slip_in., $500.));
        if _sraw_ ne '' and _sraw_ ne '-' then
            _ins_sus_ = tranwrd(_ins_sus_, _sraw_, '');
        if clean_slip_1 ne '' and clean_slip_1 ne '-' then
            _ins_sus_ = tranwrd(_ins_sus_, clean_slip_1, '');
    end;

    /* -- 5b. Pre-processing ------------------------------------------------- */
    /* Hapus angka dalam kurung */
    _ins_sus_ = prxchange('s/\(\s*[\d\.\/\-]+\s*\)//o', -1, _ins_sus_);
    /* AND OR / AND OR → koma */
    _ins_sus_ = prxchange('s/\b(?:AND|AN|OR)\s*\/\s*(?:AND|OR)\b/,/oi', -1, _ins_sus_);
    _ins_sus_ = prxchange('s/\bAND\s+OR\b/,/oi', -1, _ins_sus_);
    /* CO., LTD. → koma */
    _ins_sus_ = prxchange('s/\bCO\.,?\s*LTD\.?\b/,/oi', -1, _ins_sus_);
    /* Strip TBK, PERSERO, LTD */
    _ins_sus_ = prxchange('s/\bTBK\.?\b//oi',        -1, _ins_sus_);
    _ins_sus_ = prxchange('s/\(PERSERO\)//oi',        -1, _ins_sus_);
    _ins_sus_ = prxchange('s/\bPERSERO\b//oi',        -1, _ins_sus_);
    _ins_sus_ = prxchange('s/\bLTD\.?\b//oi',         -1, _ins_sus_);
    _ins_sus_ = prxchange('s/\(FCI\.\s*I\)//oi',      -1, _ins_sus_);

    /* -- 5c. Split ---------------------------------------------------------- */
    /* Ganti delimiter ke '|' untuk kemudahan scan */
    /* QQ → | */
    _ins_sus_ = prxchange('s/\bQQ\b/|/oi', -1, _ins_sus_);
    /* Slash → | */
    _ins_sus_ = prxchange('s/\// |/o', -1, _ins_sus_);
    /* Koma → | */
    _ins_sus_ = prxchange('s/,/|/o', -1, _ins_sus_);
    /* Dash (kecuali sebelum tahun 19xx/20xx) → | */
    _ins_sus_ = prxchange('s/-(?!\s*(?:19|20)\d{2}\b)/|/o', -1, _ins_sus_);
    /* "1. 2. 3. ..." → | */
    _ins_sus_ = prxchange('s/\d+\./|/o', -1, _ins_sus_);
    /* PT dan CV → | */
    _ins_sus_ = prxchange('s/\bPT\.?\b/|/oi', -1, _ins_sus_);
    _ins_sus_ = prxchange('s/\bCV\.?\b/|/oi', -1, _ins_sus_);
    /* Titik dua, titik koma → | */
    _ins_sus_ = prxchange('s/[:;]/|/o', -1, _ins_sus_);

    /* -- 5d. Proses tiap bagian --------------------------------------------- */
    _n_ins_parts_ = countw(_ins_sus_, '|', 'm');
    do _ipi_ = 1 to _n_ins_parts_;
        _ipart_ = strip(scan(_ins_sus_, _ipi_, '|', 'm'));

        /* Normalisasi: hapus tail patterns */
        _ipart_ = prxchange('s/\bAS\b\s+(?:THE\s+)?(?:PRINCIPAL|OFF-TAKER|MAINTENANCE|CONTRACTOR).*//oi', -1, _ipart_);
        _ipart_ = prxchange('s/\bBEING\s+(?:THE\s+)?(?:PRINCIPAL|OFF-TAKER).*//oi', -1, _ipart_);
        _ipart_ = prxchange('s/\bAND\s+ALL\s+SUBSIDIARI.*//oi', -1, _ipart_);
        _ipart_ = prxchange('s/\bINCLUDING\s+(?:ALL|ANY)\s+SUBSIDIAR.*//oi', -1, _ipart_);
        _ipart_ = prxchange('s/\bCOMPRISING\s+OF.*//oi', -1, _ipart_);
        _ipart_ = prxchange('s/\bINSTALLMENT\b.*//oi', -1, _ipart_);
        _ipart_ = prxchange('s/\bRELATED\s+COMPANY\b.*//oi', -1, _ipart_);
        _ipart_ = prxchange('s/\bPURCHASED\s+OR\s+OTHERWISE\b.*//oi', -1, _ipart_);

        /* Strip KB, A.W. */
        _ipart_ = prxchange('s/\bKB\b//oi', -1, _ipart_);
        _ipart_ = prxchange('s/\bA\.?W\.?\b//oi', -1, _ipart_);

        /* Hapus kurung kosong */
        _ipart_ = prxchange('s/\(\s*\)//o', -1, _ipart_);

        /* Strip AND/OR di awal dan akhir */
        _ipart_ = prxchange('s/^(?:AND|OR)\b\s*//oi', -1, _ipart_);
        _ipart_ = prxchange('s/\s*\b(?:AND|OR)$//oi', -1, _ipart_);

        /* Strip karakter non-alfanumerik di awal dan akhir */
        _ipart_ = prxchange('s/^[^a-zA-Z0-9(]+//o', -1, _ipart_);
        _ipart_ = prxchange('s/[^a-zA-Z0-9)]+$//o', -1, _ipart_);
        _ipart_ = strip(_ipart_);

        /* -- 5e. Validasi bagian -------------------------------------------- */
        if lengthn(_ipart_) <= 2 then continue;
        /* Skip pure digit/slash/dash */
        if prxmatch('/^[\d\/\-\.]+$/', _ipart_) then continue;
        /* Skip junk: nomor bangunan, lantai, jalan */
        if prxmatch('/\b(?:NO\.\s*\d+|BUILDING|FLOOR|ROOM|ROAD|STREET|TOWER|KAV\.?|BLOK)\b/oi', _ipart_) then continue;
        /* Skip: PLTGU, MW, dsb */
        if prxmatch('/\b(?:PLTGU|PLTMH|PLTU|POWER PLANT|COMBINED CYCLE|MW|HYDRO ELECTRIC)\b/oi', _ipart_) then continue;
        /* Skip JUNK WORDS */
        _ipart_up_ = upcase(strip(_ipart_));
        if _ipart_up_ in ('SHANGHAI','PR OF CHINA','CHINA','INDONESIA','JAKARTA',
            'OFFICERS','EMPLOYEES','ALL OTHER CONTRACTORS',
            'SUB-CONTRACTORS','SUB CONTRACTORS',
            'COMPANIES','AFFILIATED','AFFILIATES',
            'SUBSIDIARY','SUBSIDIARIES','ANY SUBSIDIARY COMPANY',
            'RELATED COMPANY','MIGRASI AS400',
            'THE PRINCIPAL','PRINCIPAL','OWNER') then continue;

        /* Simpan bagian yang valid */
        if _cis_n_ < 5 then do;
            _cis_n_ + 1; _cis_out{_cis_n_} = _ipart_;
        end;
    end;

    _insus_done_: ;
    drop _ins_sus_ _cis_n_ _ipart_ _ipart_up_ _ipi_ _n_ins_parts_
         _praw_ _sraw_ _i_;
%mend apply_clean_insured_suspend;


/* =============================================================================
   STEP 6: Jalankan semua cleaning
   ============================================================================= */
%put [3/5] Cleaning (polis, slip, insured, sertif) ...;

data work.suspend_cleaned;
    set work.suspend_renamed;

    length clean_polis_1   - clean_polis_5   $500
           clean_slip_1    - clean_slip_5    $500
           clean_insured_1 - clean_insured_5 $500
           clean_sertif_1  - clean_sertif_5  $10;

    /* -- 6a. Clean POLIS (strip suffix only) -------------------------------- */
    %apply_clean_polis_suspend(polis_ori);

    /* -- 6b. Clean SLIP (as-is) --------------------------------------------- */
    /* Setara dengan: return [val] if val else [] */
    clean_slip_1 = strip(polis_ori);
    do _i_ = 2 to 5; clean_slip_1 = ''; end; /* kolom 2-5 kosong */
    if not missing(slip_ori) then clean_slip_1 = strip(slip_ori);
    /* Kolom 2-5 tetap kosong */
    clean_slip_2 = ''; clean_slip_3 = ''; clean_slip_4 = ''; clean_slip_5 = '';
    drop _i_;

    /* -- 6c. Ekstrak SERTIFIKAT dari polis_ori (format suffix & S/D) -------- */
    %apply_extract_sertif(polis_ori);

    /* -- 6d. Clean INSURED (logika suspend-specific) ------------------------ */
    %apply_clean_insured_suspend(insured_ori, polis_ori, slip_ori);

    /* -- 6e. Clean control chars -------------------------------------------- */
    array _str_sus{*} _character_;
    do _sci_ = 1 to dim(_str_sus);
        if not missing(_str_sus{_sci_}) then
            %clean_control_chars(_str_sus{_sci_});
    end;
    drop _sci_;

run;

%put       Cleaning selesai.;


/* =============================================================================
   STEP 7: Susun urutan kolom output
   ============================================================================= */
%put [4/5] Building output columns ...;

data work.suspend_output;
    retain "&CEDANT_COL."n &STATUS_COL.
           polis_ori    clean_polis_1-clean_polis_5   clean_sertif_1-clean_sertif_5
           slip_ori     clean_slip_1-clean_slip_5
           insured_ori  clean_insured_1-clean_insured_5;
    set work.suspend_cleaned;
run;


/* =============================================================================
   STEP 8: Simpan ke Excel
   ============================================================================= */
%put [5/5] Saving to: &OUT_FILE.;

proc export
    data=work.suspend_output
    outfile="&OUT_FILE."
    dbms=xlsx
    replace;
    sheet="suspend_clean";
run;

%put Done. Output: &OUT_FILE.;
%put [SUSPEND CLEANING] Selesai.;

/* Cleanup */
proc datasets library=work nolist;
    delete suspend_raw suspend_filtered suspend_renamed suspend_cleaned;
run;
quit;
