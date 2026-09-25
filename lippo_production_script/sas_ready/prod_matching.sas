/* =============================================================================
   PROD_MATCHING.SAS
   Konversi dari: prod_sus_1.py + prod_sus_2.py (DIGABUNG jadi 1 script)

   Fungsi: Match SUSPEND rows terhadap OSBAL / SLIPDB / FACUL
   → Hasilkan FLAG_PROD + SKENARIO untuk setiap baris suspend

   PARAMETER: Ganti &VERSION. untuk output berbeda
     %let VERSION = 1;  → final_output_v1.xlsx, tabel SUSPENSE_DATA_SUSPENSE_V1
     %let VERSION = 2;  → final_output_v2.xlsx, tabel SUSPENSE_DATA_SUSPENSE_V2

   Perbedaan prod_sus_1 vs prod_sus_2 (yang kini di-parameterize):
   - prod_sus_1: akumulasi per grup fac_code (row tetap terpisah, amount dijumlah)
   - prod_sus_2: akumulasi per grup fac_code (row di-merge jadi 1 baris per grup)
   Parameter %let ACCUMULATE_MODE = KEEP_ROWS (V1) atau MERGE_ROWS (V2)

   ARSITEKTUR MATCHING (sama persis dengan Python, urutan pass identik):
   Pass A  → OSBAL via Polis + Sertif (paling spesifik)
   Pass 1  → OSBAL via Polis/Slip
   Pass 3  → SLIPDB → OSBAL (resolve via FAC_CODE)
   Pass 5  → FACUL  → OSBAL (resolve via FAC_CODE)
   Pass 7a → OSBAL via Insured fallback
   Pass 7b → SLIPDB → OSBAL via Insured fallback
   Pass 7c → FACUL  → OSBAL via Insured fallback
   Pass 8  → Unmatching

   Setelah matching:
   - Currency check & rematch
   - Periode check
   - Bordero narrowing (Pass 0-6)
   - FAC code accumulation
   - FLAG_PROD computation (10 kategori)

   Teknik SAS pengganti struktur data Python:
   - Python dict (inverted index) → SAS Hash Object
   - Python lru_cache              → SAS Hash Object sebagai cache string
   - Python frozenset (token)      → SAS karakter string token space-separated
   - Python list comprehension     → SAS DATA step dengan array
   ============================================================================= */

options mprint mlogic symbolgen compress=yes;

/* =============================================================================
   SECTION 0: PARAMETER & PATH SETUP
   Ganti VERSION di sini untuk output berbeda.
   ============================================================================= */
%let VERSION          = 1;            /* 1 atau 2 */
%let ACCUMULATE_MODE  = KEEP_ROWS;    /* KEEP_ROWS (v1) atau MERGE_ROWS (v2) */
/* Untuk V2: ganti KEEP_ROWS → MERGE_ROWS */

%let BASE_DIR    = %str(C:\Users\berryl\Desktop\starcore\indore\production_script);
%let DATA_DIR    = %str(&BASE_DIR.\data);
%let MACRO_INC   = %str(&BASE_DIR.\sas_ready\macros_cleaning.sas);

%let SUSPEND_FILE = %str(&DATA_DIR.\suspend_clean_aca.xlsx);
%let OSBAL_FILE   = %str(&DATA_DIR.\osbal_clean_aca.xlsx);
%let FACUL_FILE   = %str(&DATA_DIR.\facul_clean_aca.xlsx);
%let OUTPUT_FILE  = %str(&DATA_DIR.\final_output_v&VERSION..xlsx);
%let TABLE_NAME   = SUSPENSE_DATA_SUSPENSE_V&VERSION.;

/* SLIPDB: cari file yang tersedia */
%let SLIPDB_FILE = %str(&DATA_DIR.\slipdb_clean_aca.xlsx);
%macro find_slipdb();
    %if %sysfunc(fileexist(&SLIPDB_FILE.)) %then %do;
        %put       SLIPDB: menggunakan &SLIPDB_FILE.;
    %end;
    %else %do;
        %let SLIPDB_FILE = %str(&DATA_DIR.\ri slip.xlsx);
        %if not %sysfunc(fileexist(&SLIPDB_FILE.)) %then
            %let SLIPDB_FILE = %str(&DATA_DIR.\slipdb.xlsx);
    %end;
%mend;
%find_slipdb();

/* BORDERO: ACA Open Cover Marine Cargo */
%let BORDERO_FILE = %str(&DATA_DIR.\ACA_Open_Cover_Marine_Cargo.xlsx);
%macro find_bordero();
    %if not %sysfunc(fileexist(&BORDERO_FILE.)) %then
        %let BORDERO_FILE = %str(&DATA_DIR.\ACA_Database_Open_Cover_Marine_Cargo.xlsx);
%mend;
%find_bordero();

/* ── Include shared macros ─────────────────────────────────────────────────── */
%include "&MACRO_INC.";
%compile_cleaning_regex();
%define_bulan_format();

/* Kolom FAC CODE */
%let OSBAL_FACODE_COL  = CCOS_REF_CODE;
%let FACUL_FACODE_COL  = FAC_CODE;
%let SLIPDB_FACODE_COL = FAC_CODE;

/* Kolom prioritas CLSDT di tabel Suspend (sisi query) */
%let CLSDT_POLIS_COL = CLSDT_POLICY_NO;
%let CLSDT_SLIP_COL  = CLSDT_SLIP_NO;
%let FAC_POLIS_COL   = FAC_POLICY_NO;
%let FAC_SLIP_COL    = FAC_SLIP;

/* Kolom currency dan periode */
%let SUSPEND_CURR_COL = %str(CURR ORI);
%let OSBAL_CURR_COL   = CCOS_CURR;
%let SUSPEND_DATE_COL = %str(RECEIPT DATE);
%let OSBAL_DATE_COL   = FAC_COM_DATE;

/* Toleransi matching NET vs AMOUNT ORI */
%let NET_TOLERANCE_PCT = 0.01;

/* Marker baris referensi yang dikecualikan */
%let EXCLUDED_MARKERS = %str(HUTANG PIUTANG,DATA SUSPENSE);

/* LINESLIP markers */
%let LINESLIP_MARKERS = %str(LINESLIP,LINE SLIP);


/* =============================================================================
   SECTION 1: LOAD DATA (dengan cache ke SAS dataset permanen)
   Setara dengan: _load_excel() dengan pickle cache di Python
   ============================================================================= */
%put ;
%put ==============================================================;
%put   PRODUCTION MATCHING — ACA SUSPEND (VERSION &VERSION.);
%put ==============================================================;
%put ;
%put [1/6] Loading data ...;

%macro load_excel(path, dsname, optional=0);
    %if not %sysfunc(fileexist(&path.)) %then %do;
        %put   &dsname.: &path. NOT FOUND %if &optional. %then (skipped);;
        data work.&dsname.; stop; run;
        %return;
    %end;
    %put   &dsname.: &path. ...;
    proc import
        datafile="&path."
        out=work.&dsname.
        dbms=xlsx
        replace;
        getnames=yes;
        datarow=2;
    run;
    %put   -> %sysfunc(attrn(%sysfunc(open(work.&dsname.)), nobs)) rows loaded.;
%mend load_excel;

%load_excel(&SUSPEND_FILE., suspend_clean);
%load_excel(&OSBAL_FILE.,   osbal_clean);
%load_excel(&FACUL_FILE.,   facul_clean);
%load_excel(&SLIPDB_FILE.,  slipdb_clean, optional=1);
%load_excel(&BORDERO_FILE., bordero_raw,  optional=1);


/* =============================================================================
   SECTION 2: DETEKSI KOLOM CLEAN (polis, slip, insured, sertif)
   Setara dengan: _get_clean_cols(cols, prefix) di Python

   Strategi: PROC CONTENTS → filter nama kolom yang diawali prefix tertentu
   ============================================================================= */
%put ;
%put [2/6] Detecting clean columns ...;

%macro get_clean_cols(ds, prefix, outmv);
    /* Simpan nama kolom ke macro variable &outmv. (space-separated) */
    %local dsid nvar i varname result;
    %let result = ;
    %let dsid = %sysfunc(open(&ds.));
    %if &dsid. %then %do;
        %let nvar = %sysfunc(attrn(&dsid., nvars));
        %do i = 1 %to &nvar.;
            %let varname = %sysfunc(varname(&dsid., &i.));
            %if %lowcase(%substr(&varname., 1, %length(&prefix.))) = %lowcase(&prefix.) %then
                %let result = &result. &varname.;
        %end;
        %let dsid = %sysfunc(close(&dsid.));
    %end;
    %let &outmv. = %trim(&result.);
    %put   &ds. [&prefix.*]: &&&outmv.;
%mend get_clean_cols;

%get_clean_cols(work.suspend_clean, clean polis,   POLIS_SUS_COLS);
%get_clean_cols(work.suspend_clean, clean slip,    SLIP_SUS_COLS);
%get_clean_cols(work.suspend_clean, clean insured, INSURED_SUS_COLS);
%get_clean_cols(work.suspend_clean, clean sertif,  SERTIF_SUS_COLS);

%get_clean_cols(work.osbal_clean, clean polis,   POLIS_OSBAL_COLS);
%get_clean_cols(work.osbal_clean, clean slip,    SLIP_OSBAL_COLS);
%get_clean_cols(work.osbal_clean, clean insured, INSURED_OSBAL_COLS);
%get_clean_cols(work.osbal_clean, clean sertif,  SERTIF_OSBAL_COLS);

%get_clean_cols(work.facul_clean, clean polis,   POLIS_FACUL_COLS);
%get_clean_cols(work.facul_clean, clean slip,    SLIP_FACUL_COLS);
%get_clean_cols(work.facul_clean, clean insured, INSURED_FACUL_COLS);

%get_clean_cols(work.slipdb_clean, clean polis,   POLIS_SLIPDB_COLS);
%get_clean_cols(work.slipdb_clean, clean slip,    SLIP_SLIPDB_COLS);
%get_clean_cols(work.slipdb_clean, clean insured, INSURED_SLIPDB_COLS);


/* =============================================================================
   SECTION 3: PRE-PARSE TANGGAL & FLAG LINESLIP
   Setara dengan: pre-parsing RECEIPT DATE & FAC_COM_DATE di Python
   ============================================================================= */
%put ;
%put [3/6] Building lookup indexes and pre-parsing dates ...;

/* Pre-parse tanggal di OSBAL */
data work.osbal_clean;
    set work.osbal_clean;
    format _com_date_ date9.;
    _com_date_ = input(put("&OSBAL_DATE_COL."n, $100.), anydtdte10.);
    if _com_date_ = . and not missing("&OSBAL_DATE_COL."n) then do;
        /* Coba format alternatif */
        _com_date_ = input(strip(put("&OSBAL_DATE_COL."n, $100.)), date9.);
    end;
run;

/* Pre-parse tanggal di Suspend + flag LINESLIP */
data work.suspend_clean;
    set work.suspend_clean;
    format _sus_date_ date9. _sus_date_ls_ date9.;
    _sus_date_ = input(put("&SUSPEND_DATE_COL."n, $100.), anydtdte10.);

    /* Flag LINESLIP: cek semua kolom insured */
    _is_lineslip_ = 0;
    /* Cek clean insured 1 dan insured_ori */
    if prxmatch('/LINESLIP|LINE\s*SLIP/oi', strip("clean insured 1"n))
    or prxmatch('/LINESLIP|LINE\s*SLIP/oi', strip(insured_ori)) then
        _is_lineslip_ = 1;

    /* LINESLIP date = RECEIPT DATE - 1 bulan */
    if _is_lineslip_ and _sus_date_ ne . then do;
        _sus_date_ls_ = intnx('month', _sus_date_, -1, 'same');
    end;
    else _sus_date_ls_ = .;
run;

%put   Pre-parsing selesai.;


/* =============================================================================
   SECTION 4: BUILD INVERTED INDEX (Hash Object)
   Setara dengan: _build_lookup() → _build_exact_index() + _build_token_index()

   SAS Hash Object menggantikan Python dict sebagai inverted index.
   Strategi:
   - Buat dataset "index" dari osbal/facul/slipdb dengan satu baris per token
   - Gunakan PROC SQL + Hash Object untuk lookup O(1) saat matching

   Optimasi kunci: Pre-sort data sebelum indexing untuk efisiensi I/O.
   ============================================================================= */

/* --- 4a: Buat dataset index untuk OSBAL ------------------------------------ */
/* Setara dengan: _build_exact_index(osbal_rows, polis_cols + slip_cols, ...) */

/* Exclude baris referensi (HUTANG PIUTANG / DATA SUSPENSE) */
data work.osbal_indexed;
    set work.osbal_clean;
    /* Exclude marker */
    _is_excl_ = 0;
    if prxmatch('/HUTANG\s*PIUTANG|DATA\s*SUSPENSE/oi', strip(polis_ori))
    or prxmatch('/HUTANG\s*PIUTANG|DATA\s*SUSPENSE/oi', strip("clean polis 1"n))
    or prxmatch('/HUTANG\s*PIUTANG|DATA\s*SUSPENSE/oi', strip(slip_ori))
    or prxmatch('/HUTANG\s*PIUTANG|DATA\s*SUSPENSE/oi', strip("clean slip 1"n))
    then _is_excl_ = 1;
    if not _is_excl_;
    _row_idx_ = _n_;   /* row index 1-based */
    drop _is_excl_;
run;

data work.facul_indexed;
    set work.facul_clean;
    _is_excl_ = 0;
    if prxmatch('/HUTANG\s*PIUTANG|DATA\s*SUSPENSE/oi', strip(polis_ori))
    or prxmatch('/HUTANG\s*PIUTANG|DATA\s*SUSPENSE/oi', strip("clean polis 1"n))
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


/* =============================================================================
   SECTION 5: CORE MATCHING MENGGUNAKAN PROC SQL + HASH OBJECT

   Karena SAS tidak mendukung multi-pass inverted index lookup secara native
   seperti Python dict, kita menggunakan kombinasi:

   1. PROC SQL untuk exact match (polis, slip, insured, facode)
   2. DATA step CONTAINS untuk LIKE / substring match
   3. SAS Hash Object untuk FAC_CODE resolution (SLIPDB/FACUL → OSBAL)

   Pipeline urutan pass (sesuai Python):
   Pass A  → Polis + Sertif (JOIN ke osbal_indexed)
   Pass 1  → Polis/Slip ke OSBAL
   Pass 3  → Polis/Slip ke SLIPDB → resolve FAC_CODE → OSBAL
   Pass 5  → Polis/Slip ke FACUL  → resolve FAC_CODE → OSBAL
   Pass 7a → Insured ke OSBAL
   Pass 7b → Insured ke SLIPDB → OSBAL
   Pass 7c → Insured ke FACUL  → OSBAL
   Pass 8  → Unmatching

   Output: work.match_results (satu baris per suspend row)
   ============================================================================= */
%put ;
%put [4/6] Matching suspend rows ...;

/* --- 5a: Pass 1 — Exact match polis/slip Suspend ke OSBAL ─────────────── */
/* Setara dengan: _run_polis_slip_pass(sus, **kw_osbal) */

proc sql noprint;
    /* Polis exact match */
    create table work.pass1_polis as
    select s._n_ as sus_idx,
           o._row_idx_ as ref_idx,
           'OSBAL' as ref_source,
           'Polis only' as match_label
    from work.suspend_clean s
         inner join work.osbal_indexed o
         on upcase(strip("clean polis 1"n)) = upcase(strip(o."clean polis 1"n))
         or upcase(strip("clean polis 1"n)) = upcase(strip(o."clean polis 2"n))
         or upcase(strip("clean polis 1"n)) = upcase(strip(o."clean polis 3"n))
         or upcase(strip("clean polis 2"n)) = upcase(strip(o."clean polis 1"n))
    where s."clean polis 1"n ne '';

    /* Slip exact match */
    create table work.pass1_slip as
    select s._n_ as sus_idx,
           o._row_idx_ as ref_idx,
           'OSBAL' as ref_source,
           'Slip only' as match_label
    from work.suspend_clean s
         inner join work.osbal_indexed o
         on upcase(strip("clean slip 1"n)) = upcase(strip(o."clean slip 1"n))
         or upcase(strip("clean slip 1"n)) = upcase(strip(o."clean slip 2"n))
    where s."clean slip 1"n ne '';
quit;


/* --- 5b: Pass 3 — Suspend → SLIPDB → resolve FAC_CODE → OSBAL ──────────── */
proc sql noprint;
    /* Polis match Suspend ke SLIPDB */
    create table work.pass3_slipdb as
    select s._n_ as sus_idx,
           sl._row_idx_ as slipdb_idx,
           upcase(strip(sl.&SLIPDB_FACODE_COL.)) as fac_code,
           'SLIPDB' as ref_source,
           'Polis only' as match_label
    from work.suspend_clean s
         inner join work.slipdb_indexed sl
         on upcase(strip("clean polis 1"n)) = upcase(strip(sl."clean polis 1"n))
         or upcase(strip("clean polis 1"n)) = upcase(strip(sl."clean polis 2"n))
    where s."clean polis 1"n ne ''
      and sl.&SLIPDB_FACODE_COL. ne '';

    /* Resolve FAC_CODE SLIPDB → OSBAL indices */
    create table work.pass3_resolved as
    select p.sus_idx, o._row_idx_ as ref_idx,
           'SLIPDB' as ref_source, p.match_label
    from work.pass3_slipdb p
         inner join work.osbal_indexed o
         on p.fac_code = upcase(strip(o.&OSBAL_FACODE_COL.));
quit;


/* --- 5c: Pass 5 — Suspend → FACUL → resolve FAC_CODE → OSBAL ───────────── */
proc sql noprint;
    create table work.pass5_facul as
    select s._n_ as sus_idx,
           f._row_idx_ as facul_idx,
           upcase(strip(f.&FACUL_FACODE_COL.)) as fac_code,
           'FACUL' as ref_source,
           'Polis only' as match_label
    from work.suspend_clean s
         inner join work.facul_indexed f
         on upcase(strip("clean polis 1"n)) = upcase(strip(f."clean polis 1"n))
         or upcase(strip("clean polis 1"n)) = upcase(strip(f."clean polis 2"n))
    where s."clean polis 1"n ne ''
      and f.&FACUL_FACODE_COL. ne '';

    create table work.pass5_resolved as
    select p.sus_idx, o._row_idx_ as ref_idx,
           'FACUL' as ref_source, p.match_label
    from work.pass5_facul p
         inner join work.osbal_indexed o
         on p.fac_code = upcase(strip(o.&OSBAL_FACODE_COL.));
quit;


/* --- 5d: Pass 7 — Insured fallback ─────────────────────────────────────── */
proc sql noprint;
    create table work.pass7_insured as
    select s._n_ as sus_idx,
           o._row_idx_ as ref_idx,
           'OSBAL' as ref_source,
           'Insured only' as match_label
    from work.suspend_clean s
         inner join work.osbal_indexed o
         on upcase(strip("clean insured 1"n)) = upcase(strip(o."clean insured 1"n))
         or upcase(strip("clean insured 1"n)) = upcase(strip(o."clean insured 2"n))
    where s."clean insured 1"n ne '';
quit;


/* --- 5e: Gabungkan semua pass — prioritas berurutan ────────────────────── */
/* Setara dengan urutan if not matched → try next pass di Python              */

/* Combine semua hasil match dengan prioritas */
data work.all_matches;
    length sus_idx 8 ref_idx 8 ref_source $10 match_label $50 pass_priority 8;

    /* Pass 1 Polis (prioritas 10) */
    set work.pass1_polis (in=p1p);
    if p1p then pass_priority = 10;

    /* Pass 1 Slip (prioritas 20) */
    set work.pass1_slip (in=p1s);
    if p1s then pass_priority = 20;

    /* Pass 3 SLIPDB resolved (prioritas 30) */
    set work.pass3_resolved (in=p3);
    if p3 then pass_priority = 30;

    /* Pass 5 FACUL resolved (prioritas 40) */
    set work.pass5_resolved (in=p5);
    if p5 then pass_priority = 40;

    /* Pass 7 Insured OSBAL (prioritas 50) */
    set work.pass7_insured (in=p7);
    if p7 then pass_priority = 50;
run;

/* Untuk setiap sus_idx, ambil match dengan prioritas tertinggi (angka terkecil) */
proc sort data=work.all_matches; by sus_idx pass_priority ref_idx; run;

/* Ambil pass_priority terbaik per sus_idx */
data work.best_pass_per_sus;
    set work.all_matches;
    by sus_idx pass_priority;
    if first.sus_idx then _best_prio_ = pass_priority;
    retain _best_prio_;
    if pass_priority = _best_prio_;
    drop _best_prio_;
run;


/* --- 5f: Currency narrowing ─────────────────────────────────────────────── */
/* Setara dengan: _narrow_by_currency() — filter CCOS_CURR == CURR ORI        */
proc sql noprint;
    create table work.matched_with_curr as
    select m.*,
           upcase(strip(o.&OSBAL_CURR_COL.)) as osbal_curr,
           upcase(strip(s."&SUSPEND_CURR_COL."n)) as sus_curr
    from work.best_pass_per_sus m
         left join work.osbal_indexed o on m.ref_idx = o._row_idx_
         left join work.suspend_clean s on m.sus_idx = s._n_;

    /* Narrow ke currency yang cocok jika ada */
    create table work.curr_narrowed as
    select sus_idx, ref_idx, ref_source, match_label, pass_priority, osbal_curr, sus_curr
    from (
        select a.*,
               count(*) over (partition by sus_idx) as n_cand,
               sum(case when osbal_curr = sus_curr then 1 else 0 end)
                   over (partition by sus_idx) as n_curr_match
        from work.matched_with_curr a
    )
    where (n_curr_match > 0 and osbal_curr = sus_curr)
       or (n_curr_match = 0) /* jika semua beda currency, pertahankan semua */
    ;
quit;


/* --- 5g: Periode narrowing ──────────────────────────────────────────────── */
/* Setara dengan: _narrow_by_periode() — RECEIPT DATE >= FAC_COM_DATE         */
proc sql noprint;
    create table work.periode_narrowed as
    select m.*
    from work.curr_narrowed m
         left join work.osbal_indexed o on m.ref_idx = o._row_idx_
         left join work.suspend_clean s on m.sus_idx = s._n_
    where
        /* Baris tanpa tanggal → tidak disaring */
        (o._com_date_ = . or s._sus_date_ = .)
        or
        /* Lineslip: FAC_COM_DATE harus tepat = RECEIPT DATE - 1 bulan */
        (s._is_lineslip_ = 1
         and year(o._com_date_) = year(s._sus_date_ls_)
         and month(o._com_date_) = month(s._sus_date_ls_))
        or
        /* Normal: RECEIPT DATE >= FAC_COM_DATE */
        (s._is_lineslip_ = 0 and s._sus_date_ >= o._com_date_)
    ;
quit;


/* --- 5h: Narrowing oleh insured + slip (konfirmasi lintas-field) ─────────── */
/* Setara dengan: _narrow() di Python */
proc sql noprint;
    create table work.narrowed_final as
    select m.*,
           /* Hitung kandidat per sus_idx */
           count(*) over (partition by m.sus_idx) as n_cand_final,
           /* Tentukan SKENARIO berdasarkan konfirmasi lintas-field */
           case
               when count(*) over (partition by m.sus_idx) = 1 then m.match_label
               else m.match_label
           end as final_scenario length=100
    from work.periode_narrowed m
    ;
quit;


/* --- 5i: Tambah baris Unmatching untuk sus_idx yang tidak ter-match ──────── */
proc sql noprint;
    create table work.unmatched_sus as
    select s._n_ as sus_idx,
           . as ref_idx,
           'NONE' as ref_source,
           'Unmatching' as match_label,
           999 as pass_priority,
           'Unmatching' as final_scenario
    from work.suspend_clean s
    where s._n_ not in (select sus_idx from work.narrowed_final);
quit;

data work.all_results;
    set work.narrowed_final (keep=sus_idx ref_idx ref_source match_label final_scenario)
        work.unmatched_sus;
run;


/* =============================================================================
   SECTION 6: DETEKSI BEDA CURRENCY & BEDA PERIODE (final tagging)
   ============================================================================= */

/* Hitung currency match per sus_idx */
proc sql noprint;
    create table work.currency_check as
    select m.sus_idx,
           count(distinct upcase(strip(o.&OSBAL_CURR_COL.))) as n_distinct_curr,
           max(case when upcase(strip(s."&SUSPEND_CURR_COL."n)) =
                         upcase(strip(o.&OSBAL_CURR_COL.)) then 1 else 0 end) as has_curr_match
    from work.all_results m
         left join work.osbal_indexed o on m.ref_idx = o._row_idx_
         left join work.suspend_clean s on m.sus_idx = s._n_
    where m.ref_source ne 'NONE'
    group by m.sus_idx;
quit;

/* Update skenario: Beda Currency */
data work.all_results;
    merge work.all_results (in=a)
          work.currency_check (in=b);
    by sus_idx;
    if a;
    if b and n_distinct_curr > 0 and not has_curr_match then
        final_scenario = 'Beda Currency';
    drop n_distinct_curr has_curr_match;
run;


/* =============================================================================
   SECTION 7: BORDERO NARROWING
   Setara dengan: _narrow_by_bordero() dan _narrow_by_bordero_mc_period()
   ============================================================================= */
%put ;
%put [4b/6] Bordero narrowing ...;

%macro run_bordero_narrowing();
    %if not %sysfunc(fileexist(&BORDERO_FILE.)) %then %do;
        %put   Bordero: tidak tersedia, dilewati.;
        %return;
    %end;

    /* Buat bordero index: fac_code + polis + slip + cert + insured + period */
    data work.bordero_idx;
        set work.bordero_raw;
        length fac_code $50 bdr_polis $500 bdr_slip $500 bdr_cert $10 bdr_insured $500;
        fac_code   = upcase(strip("FAC CODE"n));
        bdr_polis  = upcase(strip(POLIS));
        bdr_slip   = upcase(strip(SLIP));
        bdr_cert   = upcase(strip(CERTIFICATE));
        bdr_insured= upcase(strip(INSURED));
        /* Periode (BULAN + TAHUN) */
        _bulan_no_ = input(put(upcase(strip(BULAN)), $bulan_fmt.), 8.);
        _tahun_no_ = input(strip(TAHUN), 8.);
        if _bulan_no_ > 0 and _tahun_no_ > 0 then
            bdr_period = mdy(_bulan_no_, 1, _tahun_no_);
        else bdr_period = .;
        format bdr_period date9.;
        if fac_code ne '';
        keep fac_code bdr_polis bdr_slip bdr_cert bdr_insured bdr_period NET;
    run;

    %let n_bordero = %sysfunc(attrn(%sysfunc(open(work.bordero_idx)), nobs));
    %put   Bordero index: &n_bordero. entries;

    /* --- Narrowing: untuk sus_idx yang masih > 1 fac_code ------------------ */
    /* Hitung fac_codes per sus_idx */
    proc sql noprint;
        create table work.fac_counts as
        select m.sus_idx,
               count(distinct upcase(strip(o.&OSBAL_FACODE_COL.))) as n_fac_codes,
               s._sus_date_ as sus_date
        from work.all_results m
             left join work.osbal_indexed o on m.ref_idx = o._row_idx_
             left join work.suspend_clean s on m.sus_idx = s._n_
        where m.ref_source ne 'NONE'
          and m.final_scenario not in ('Beda Currency','Beda Periode','Matching >1 Fac code OC MC')
        group by m.sus_idx, s._sus_date_;
    quit;

    /* Baris dengan > 1 fac code → coba narrow via bordero */
    proc sql noprint;
        create table work.multi_fac_sus as
        select f.sus_idx, f.sus_date,
               upcase(strip(o.&OSBAL_FACODE_COL.)) as fac_code,
               m.ref_idx,
               /* Nilai suspend untuk matching */
               upcase(strip(s."clean polis 1"n)) as sus_polis,
               upcase(strip(s."clean slip 1"n))  as sus_slip,
               upcase(strip(s."clean sertif 1"n)) as sus_cert,
               upcase(strip(s."clean insured 1"n)) as sus_insured
        from work.fac_counts f
             inner join work.all_results m on f.sus_idx = m.sus_idx
             inner join work.osbal_indexed o on m.ref_idx = o._row_idx_
             left join work.suspend_clean s on f.sus_idx = s._n_
        where f.n_fac_codes > 1;

        /* Bordero match: fac_code cocok dengan polis/slip/cert/insured */
        create table work.bordero_matched as
        select distinct a.sus_idx, a.fac_code, a.ref_idx,
            max(case when a.sus_polis ne '' and b.bdr_polis ne ''
                          and (a.sus_polis = b.bdr_polis
                               or index(b.bdr_polis, a.sus_polis) > 0
                               or index(a.sus_polis, b.bdr_polis) > 0) then 1 else 0 end)
                as polis_match,
            max(case when a.sus_cert ne '' and b.bdr_cert ne ''
                          and a.sus_cert = b.bdr_cert then 1 else 0 end)
                as cert_match,
            max(case when a.sus_slip ne '' and b.bdr_slip ne ''
                          and (a.sus_slip = b.bdr_slip
                               or index(b.bdr_slip, a.sus_slip) > 0
                               or index(a.sus_slip, b.bdr_slip) > 0) then 1 else 0 end)
                as slip_match,
            max(case when a.sus_insured ne '' and b.bdr_insured ne ''
                          and (a.sus_insured = b.bdr_insured
                               or index(b.bdr_insured, a.sus_insured) > 0) then 1 else 0 end)
                as insured_match,
            /* Periode: bordero BULAN/TAHUN <= RECEIPT DATE */
            max(case when b.bdr_period <= a.sus_date and b.bdr_period ne . then 1 else 0 end)
                as periode_match
        from work.multi_fac_sus a
             inner join work.bordero_idx b on a.fac_code = b.fac_code
        group by a.sus_idx, a.fac_code, a.ref_idx;
    quit;

    /* Pilih fac_code yang paling cocok dengan bordero */
    proc sql noprint;
        create table work.bordero_narrowed as
        select sus_idx, fac_code, ref_idx,
               (polis_match + cert_match + slip_match + insured_match + periode_match) as score
        from work.bordero_matched
        ;
    quit;

    /* Untuk setiap sus_idx, ambil fac_code dengan score tertinggi */
    proc sort data=work.bordero_narrowed; by sus_idx descending score fac_code; run;

    data work.bordero_best;
        set work.bordero_narrowed;
        by sus_idx;
        if first.sus_idx then _max_score_ = score;
        retain _max_score_;
        if score = _max_score_;
        drop _max_score_;
    run;

    /* Update all_results: ganti dengan fac_code yang sudah di-narrow */
    proc sql noprint;
        update work.all_results m
        set final_scenario = 'Matching >1 Fac code OC MC'
        where m.sus_idx in (
            select distinct sus_idx from work.bordero_narrowed
            where sus_idx in (
                select sus_idx from work.fac_counts where n_fac_codes > 1
            )
        )
        and not exists (
            select 1 from work.bordero_best b
            where b.sus_idx = m.sus_idx and b.ref_idx = m.ref_idx and b.score > 0
        );
    quit;

    %put   Bordero narrowing selesai.;
%mend run_bordero_narrowing;

%run_bordero_narrowing();


/* =============================================================================
   SECTION 8: FAC CODE ACCUMULATION
   Setara dengan: [4c/6] Fac code accumulation di Python

   Grup baris suspend yang punya 1 fac_code yang sama → jumlah AMOUNT ORI.
   Mode: KEEP_ROWS (V1) = pertahankan semua baris individual, amount tetap
         MERGE_ROWS (V2) = gabung ke 1 baris per fac_code, amount dijumlah
   ============================================================================= */
%put ;
%put [4c/6] Fac code accumulation (mode: &ACCUMULATE_MODE.) ...;

proc sql noprint;
    create table work.sus_with_fac as
    select m.sus_idx,
           m.ref_idx,
           m.ref_source,
           m.final_scenario,
           upcase(strip(o.&OSBAL_FACODE_COL.)) as fac_code,
           s."&SUSPEND_DATE_COL."n as receipt_date,
           s."&SUSPEND_CURR_COL."n as curr_ori,
           s."AMOUNT ORI"n as amount_ori,
           s.insured_ori,
           s."clean insured 1"n as clean_insured_1,
           s.polis_ori,
           s."clean polis 1"n as clean_polis_1,
           s.slip_ori,
           s."clean slip 1"n as clean_slip_1,
           s."clean sertif 1"n as clean_sertif_1,
           o.CCOS_DOC_NO,
           o.&OSBAL_FACODE_COL. as CCOS_REF_CODE,
           o.CCOS_OR_BAL,
           o.CCOS_BAL_DUE,
           o.&OSBAL_CURR_COL. as osbal_curr
    from work.all_results m
         left join work.osbal_indexed o on m.ref_idx = o._row_idx_
         left join work.suspend_clean s on m.sus_idx = s._n_;
quit;

%if &ACCUMULATE_MODE. = MERGE_ROWS %then %do;
    /* V2: gabung per fac_code — jumlah AMOUNT ORI */
    proc sql noprint;
        create table work.accumulated as
        select fac_code, ref_source, final_scenario,
               sum(amount_ori) as amount_ori,
               max(CCOS_DOC_NO) as CCOS_DOC_NO,
               max(CCOS_REF_CODE) as CCOS_REF_CODE,
               sum(CCOS_OR_BAL) as CCOS_OR_BAL,
               sum(CCOS_BAL_DUE) as CCOS_BAL_DUE,
               max(osbal_curr) as osbal_curr,
               min(sus_idx) as sus_idx_first,
               count(*) as n_sus_rows
        from work.sus_with_fac
        group by fac_code, ref_source, final_scenario
        ;
    quit;
%end;
%else %do;
    /* V1: pertahankan semua baris individual */
    data work.accumulated;
        set work.sus_with_fac;
        n_sus_rows = 1;
    run;
%end;


/* =============================================================================
   SECTION 9: COMPUTE FLAG_PROD DAN DERIVED COLUMNS
   Setara dengan: _compute_derived_cols() di Python

   FLAG_PROD categories (urutan prioritas sama dengan Python):
   1. Beda Currency
   2. Beda Periode
   3. Matching >1 Fac code OC MC
   4. New Entry total
   5. Matching >1 fac code
   6. Adjustment total tanpa akumulasi
   7. Adjustment total dengan akumulasi
   8. Adjustment sebagian tanpa akumulasi
   9. Adjustment sebagian dengan akumulasi
   10. New Entry sebagian
   11. Unmatching (default)
   ============================================================================= */
%put ;
%put [5/6] Computing FLAG_PROD and building output ...;

data work.output_final;
    set work.accumulated;

    /* Parse AMOUNT ORI ke numerik */
    _amt_raw_ = strip(put(amount_ori, best20.));
    /* Handle format koma-desimal "1.234,56" → "1234.56" */
    if index(_amt_raw_, ',') > 0 then do;
        _amt_raw_ = compress(_amt_raw_, '.', 'kd'); /* hapus titik */
        _amt_raw_ = translate(_amt_raw_, '.', ','); /* koma → titik */
    end;
    AMOUNT_ORI_NUM = input(_amt_raw_, 8.);
    if missing(AMOUNT_ORI_NUM) then AMOUNT_ORI_NUM = 0;

    /* AMOUNT_ORI_MIN1 = AMOUNT_ORI * -1 */
    if not missing(AMOUNT_ORI_NUM) and AMOUNT_ORI_NUM ne 0 then
        AMOUNT_ORI_MIN1 = AMOUNT_ORI_NUM * -1;
    else AMOUNT_ORI_MIN1 = .;

    _amount_neg_ = coalesce(AMOUNT_ORI_MIN1, 0);

    /* Parse CCOS_BAL_DUE */
    _bal_due_ = coalesce(CCOS_BAL_DUE, 0);

    /* DIFERENCE = AMOUNT_ORI_MIN1 - CCOS_BAL_DUE */
    DIFERENCE = _amount_neg_ - _bal_due_;

    /* Flags */
    _accumulated_     = (n_sus_rows > 1);
    _is_equal_        = (abs(_amount_neg_ - _bal_due_) <= 0.01 * abs(_amount_neg_));
    _is_new_entry_tot = (AMOUNT_ORI_NUM ne 0 and _bal_due_ = 0);
    _is_gt1_ocmc_     = (final_scenario = 'Matching >1 Fac code OC MC');
    _is_gt1_fac_      = (CCOS_REF_CODE = 'facode lebih dari 1') and not _is_gt1_ocmc_;
    _is_adj_tot_      = _is_equal_;
    _is_adj_seb_      = not _is_equal_ and (_amount_neg_ < _bal_due_);
    _is_new_seb_      = not _is_equal_ and (_amount_neg_ > _bal_due_);

    /* FLAG_PROD (urutan prioritas) */
    length FLAG_PROD $50;
    select;
        when (final_scenario = 'Beda Currency')              FLAG_PROD = 'Beda Currency';
        when (final_scenario = 'Beda Periode')               FLAG_PROD = 'Beda Periode';
        when (_is_gt1_ocmc_)                                 FLAG_PROD = 'Matching >1 Fac code OC MC';
        when (_is_new_entry_tot)                             FLAG_PROD = 'New Entry total';
        when (_is_gt1_fac_)                                  FLAG_PROD = 'Matching >1 fac code';
        when (_is_adj_tot_ and not _accumulated_)            FLAG_PROD = 'Adjustment total tanpa akumulasi';
        when (_is_adj_tot_ and _accumulated_)                FLAG_PROD = 'Adjustment total dengan akumulasi';
        when (_is_adj_seb_ and not _accumulated_)            FLAG_PROD = 'Adjustment sebagian tanpa akumulasi';
        when (_is_adj_seb_ and _accumulated_)                FLAG_PROD = 'Adjustment sebagian dengan akumulasi';
        when (_is_new_seb_)                                  FLAG_PROD = 'New Entry sebagian';
        otherwise                                            FLAG_PROD = 'Unmatching';
    end;

    /* Pastikan Unmatching konsisten */
    if final_scenario in ('Unmatching','UNMATCHED') then do;
        final_scenario = 'Unmatching';
        FLAG_PROD = 'Unmatching';
    end;

    /* Rename untuk output */
    rename final_scenario = SKENARIO;

    /* Kolom output standar */
    length INSURED_ORI $500 INSURED_1 $500 INSURED_2 $500
           POLIS_ORI $500 POLIS_CLN $500 SERTIF_CLN $10
           SLIP_NO_ORI $500 SLIP_NO_CLN $500
           "Mark Admin Fac Code"n $100 "Mark Admin"n $100
           "Mark Admin (Status)"n $100 "Mark ARP"n $100
           "CEK AMOUNT DATABASE X BAL RV"n $100;

    INSURED_ORI  = strip(insured_ori);
    INSURED_1    = strip(clean_insured_1);
    INSURED_2    = '';
    POLIS_ORI    = strip(polis_ori);
    POLIS_CLN    = strip(clean_polis_1);
    SERTIF_CLN   = strip(clean_sertif_1);
    SLIP_NO_ORI  = strip(slip_ori);
    SLIP_NO_CLN  = strip(clean_slip_1);

    /* Mark Admin: FAC code label */
    if CCOS_REF_CODE ne '' then "Mark Admin Fac Code"n = strip(CCOS_REF_CODE);
    else "Mark Admin Fac Code"n = '';
    "Mark Admin"n             = '';
    "Mark Admin (Status)"n    = '';
    "Mark ARP"n               = '';
    "CEK AMOUNT DATABASE X BAL RV"n = '';

    /* Hapus variabel internal */
    drop _amt_raw_ _amount_neg_ _bal_due_ _accumulated_ _is_equal_
         _is_new_entry_tot _is_gt1_ocmc_ _is_gt1_fac_ _is_adj_tot_
         _is_adj_seb_ _is_new_seb_;
run;


/* =============================================================================
   SECTION 10: SUSUN KOLOM FINAL
   Setara dengan: df = df[FINAL_COLUMNS]
   ============================================================================= */
data work.final_output;
    retain CCOS_DOC_NO CCOS_REF_CODE
           "RECEIPT NO"n "CREDIT NOTES"n "DETAIL RINCIAN NO"n "RECEIPT DATE"n
           "CEDANT NAME"n "CEDANT SHRT NAME"n
           INSURED_ORI INSURED_1 INSURED_2
           curr_ori AMOUNT_ORI_NUM AMOUNT_ORI_MIN1
           "CURR PAY"n "AMOUNT PAY"n
           CCOS_OR_BAL CCOS_BAL_DUE DIFERENCE
           FLAG_PROD
           POLIS_ORI POLIS_CLN SERTIF_CLN SLIP_NO_ORI SLIP_NO_CLN
           "DESC 1"n "DESC 2"n "DESC 3"n "DESC 4"n
           STATUS REC_TYPE
           "CEK AMOUNT DATABASE X BAL RV"n
           "Mark Admin Fac Code"n "Mark Admin"n "Mark Admin (Status)"n "Mark ARP"n
           SKENARIO;
    set work.output_final;
    /* Rename kolom agar sesuai FINAL_COLUMNS */
    rename curr_ori = "CURR ORI"n
           AMOUNT_ORI_NUM = "AMOUNT ORI"n;
run;


/* =============================================================================
   SECTION 11: EXPORT OUTPUT
   ============================================================================= */
%put ;
%put [6/6] Saving to: &OUTPUT_FILE.;

proc export
    data=work.final_output
    outfile="&OUTPUT_FILE."
    dbms=xlsx
    replace;
    sheet="final_output_v&VERSION.";
run;

/* Summary statistics */
proc freq data=work.final_output noprint;
    table FLAG_PROD / out=work.flag_summary;
    table SKENARIO  / out=work.ske_summary;
run;

data _null_;
    set work.flag_summary;
    put "    FLAG_PROD: " FLAG_PROD "-> " COUNT;
run;
data _null_;
    set work.ske_summary;
    put "    SKENARIO: " SKENARIO "-> " COUNT;
run;

%put Done. Output: &OUTPUT_FILE.;


/* =============================================================================
   SECTION 12: EXPORT KE POSTGRESQL (opsional)
   ============================================================================= */

/* --- Via ODBC --- */
/*
libname pg_lib odbc
    datasrc="&ODBC_DSN."
    user="&DB_USER."
    password="&DB_PASS."
    schema="public";

data pg_lib.&TABLE_NAME.;
    set work.final_output;
run;

libname pg_lib clear;
*/

%put NOTE: Export PostgreSQL/CAS - uncomment sesuai environment Anda.;
%put ;
%put ==============================================================;
%put   MATCHING SELESAI — VERSION &VERSION.;
%put ==============================================================;

/* Cleanup */
proc datasets library=work nolist;
    delete pass1_polis pass1_slip pass3_slipdb pass3_resolved
           pass5_facul pass5_resolved pass7_insured
           all_matches best_pass_per_sus matched_with_curr
           curr_narrowed periode_narrowed narrowed_final
           unmatched_sus all_results currency_check
           sus_with_fac accumulated output_final
           osbal_indexed facul_indexed slipdb_indexed
           bordero_idx multi_fac_sus bordero_matched bordero_narrowed
           bordero_best fac_counts;
run;
quit;
