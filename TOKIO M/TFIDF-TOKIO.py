"""Audit Insured TF-IDF mandiri untuk TOKIO MARINE.
Tidak mengubah data input atau keputusan matching produksi.
Jalankan dari folder project:
    python TFIDF-TOKIO.py
Dependensi: pandas, openpyxl, numpy, scikit-learn.
"""
from pathlib import Path
from collections import defaultdict
import re
import unicodedata
import numpy as np
import pandas as pd
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.neighbors import NearestNeighbors

ROOT = Path(__file__).resolve().parent
OUT = ROOT / 'output'
TOP_N = 10
MIN_SCORE = 0.0

# File matching input (gunakan matching_tokio_new.xlsx jika ada/lebih baru)
MATCH_NEW = ROOT / 'matching_tokio_new.xlsx'
MATCH_BASE = ROOT / 'matching_tokio.xlsx'
MATCH = MATCH_NEW if MATCH_NEW.exists() else MATCH_BASE

OSBAL_NEW = ROOT / 'data_cleaned' / 'tokio_output_osbal_new.xlsx'
OSBAL_BASE = ROOT / 'data_cleaned' / 'tokio_output_osbal.xlsx'
OSBAL = OSBAL_NEW if OSBAL_NEW.exists() else OSBAL_BASE
FACUL = ROOT / 'data_cleaned' / 'tokio_output_facul.xlsx'
BORDERO = ROOT / 'data_cleaned' / 'tokio_output_bordero.xlsx'
RI_SLIP_PATHS = [
    ROOT / 'data_raw' / '5. Database Rekap RI Slip.xlsx',
    ROOT / 'data_raw' / '5.Database Rekap RI Slip.xlsx',
    ROOT / 'data_raw' / 'Database Rekap RI Slip.xlsx',
]
DEST = OUT / 'tokio_insured_tfidf_audit.xlsx'

COMMON = {'PT', 'CV', 'UD', 'TBK', 'PERSERO', 'LTD', 'INC', 'LLC', 'CO', 'CORP', 'CORPORATION', 'COMPANY'}


def txt(value):
    if value is None or (not isinstance(value, (list, tuple)) and pd.isna(value)):
        return ''
    value = str(value).strip()
    return '' if value.upper() in {'NAN', 'NONE', 'NAT'} else value


def norm(value):
    value = unicodedata.normalize('NFKC', txt(value)).upper()
    value = re.sub(r'[^A-Z0-9]+', ' ', value)
    return ' '.join(token for token in value.split() if token not in COMMON)


def col(df, *names):
    lookup = {re.sub(r'[^A-Z0-9]+', '', str(key).upper()): key for key in df.columns}
    for name in names:
        result = lookup.get(re.sub(r'[^A-Z0-9]+', '', name.upper()))
        if result is not None:
            return result
    return None


def insured_clean_columns(df):
    """Mendeteksi berbagai template kolom insured clean:
    INSURED_1, INSURED_CLEAN_1, FAC_INSURED_CLN_1, clean insured 1, dsb.
    """
    result = []
    for column in df.columns:
        name = re.sub(r'[^A-Z0-9]+', '', str(column).upper())
        # Mendukung prefix FAC_ dan varian CLEAN / CLN
        if re.fullmatch(r'(?:FAC)?(?:INSUREDCLEAN|INSUREDCLN|CLEANINSURED|INSURED)[1-5]', name):
            result.append(column)
    return sorted(result, key=lambda c: int(re.search(r'[1-5]$', str(c)).group()))


def name_columns(df, source):
    clean = insured_clean_columns(df)
    if clean:
        return clean
    raw = col(df, 'FAC_INSURED', 'INSURED_ORI', 'INSURED', 'INSURED NAME', 'INSURED_NAME')
    if raw is None:
        raise ValueError(f'{source}: tidak ditemukan kolom insured (clean maupun raw): {list(df.columns)}')
    return [raw]


def variants(row, columns, source):
    for column in columns:
        raw = txt(row[column])
        if not raw:
            continue
        parts = re.split(r'\s*/\s*', raw) if source == 'BORDERO' else [raw]
        for part in parts:
            part = part.strip()
            if norm(part):
                yield str(column), part, norm(part)


def read_required(path, label, header=0):
    if not path.exists():
        raise FileNotFoundError(f'{label} tidak ditemukan: {path}')
    cache = Path(str(path) + '.cache.pkl')
    if cache.exists() and cache.stat().st_mtime >= path.stat().st_mtime:
        try:
            data, cols = pd.read_pickle(cache)
            if isinstance(data, list) and data and isinstance(data[0], dict):
                print(f'[{label}] memuat cepat dari cache pickle ({len(data):,} baris)...')
                return pd.DataFrame(data)
            elif isinstance(data, pd.DataFrame):
                print(f'[{label}] memuat cepat dari cache pickle ({len(data):,} baris)...')
                return data
        except Exception:
            pass
    print(f'[{label}] membaca file Excel: {path.name}...')
    return pd.read_excel(path, header=header, dtype=object)


def read_ri_slip(path):
    """Detect insured in RI Slip at Excel header rows 1-3."""
    for header in (1, 0, 2):
        frame = pd.read_excel(path, header=header, dtype=object)
        if insured_clean_columns(frame) or col(frame, 'FAC_INSURED', 'INSURED_ORI',
                                                'INSURED', 'INSURED NAME', 'INSURED_NAME') is not None:
            if header:
                print(f'[RI SLIP] Header baris {header + 1}: {path.name}')
            return frame
    raise ValueError(f'RI SLIP: tidak ditemukan kolom insured di header baris 1-3: {path}')


def locate_ri():
    for path in RI_SLIP_PATHS:
        if path.exists():
            return path
    raise FileNotFoundError('RI Slip tidak ditemukan; cek RI_SLIP_PATHS: ' + ', '.join(str(x) for x in RI_SLIP_PATHS))


def references(source_df, source):
    namecols = name_columns(source_df, source)
    fc = col(source_df, 'CCOS_REF_CODE', 'FAC_CODE', 'FAC CODE', 'FACODE')
    currency = col(source_df, 'CCOS_CURR', 'FAC_CURRENCY', 'CURRENCY', 'CURR', 'CURR ORI')
    records = {}
    for _, row in source_df.iterrows():
        for column, raw, key in variants(row, namecols, source):
            record = records.setdefault(key, {'source': source, 'name': raw, 'key': key, 'codes': set(),
                                              'currencies': set(), 'count': 0, 'columns': set()})
            record['count'] += 1
            record['columns'].add(column)
            if fc and txt(row[fc]):
                record['codes'].add(txt(row[fc]))
            if currency and txt(row[currency]):
                record['currencies'].add(txt(row[currency]).upper())
    return list(records.values())


def chosen_rows(result):
    scenario_col = col(result, 'SKENARIO')
    flag_col = col(result, 'FLAG_PROD')
    if scenario_col is None and flag_col is None:
        raise ValueError('Output matching tidak punya SKENARIO maupun FLAG_PROD')
    scenarios = result[scenario_col].fillna('').astype(str) if scenario_col else pd.Series('', index=result.index)
    flags = result[flag_col].fillna('').astype(str) if flag_col else pd.Series('', index=result.index)
    # Sama seperti audit Wahana: fokus pada baris Insured only serta Unmatching.
    mask = (scenarios.str.contains('insured', case=False, regex=False) |
            scenarios.str.contains('unmatching|unmatched', case=False, regex=True) |
            flags.str.contains('unmatching|unmatched', case=False, regex=True))
    return result.loc[mask].copy(), scenario_col, flag_col


def main():
    print('=' * 60)
    print('  AUDIT TF-IDF INSURED UNTUK TOKIO MARINE')
    print('=' * 60)
    print(f'Membaca matching Tokio dari: {MATCH.name}')
    result = read_required(MATCH, 'Matching TOKIO')
    data = [
        ('OSBAL', read_required(OSBAL, 'OSBAL')),
        ('FACUL', read_required(FACUL, 'FACUL')),
        ('RI SLIP', read_ri_slip(locate_ri())),
    ]
    if BORDERO.exists():
        data.append(('BORDERO', read_required(BORDERO, 'BORDERO')))
    selected, scenario_col, flag_col = chosen_rows(result)
    print(f'Total baris Suspend Tokio yang diaudit (Insured/Unmatching): {len(selected):,} baris')
    
    qcols = name_columns(result, 'SUSPEND')
    ori_col = col(result, 'INSURED_ORI', 'FAC_INSURED', 'INSURED')
    refs = []
    stats = []
    for source, frame in data:
        group = references(frame, source)
        stats.append({'METRIK': f'Referensi unik {source}', 'NILAI': len(group)})
        refs.extend(group)
    if not refs:
        raise ValueError('Ketiga sumber referensi tidak memiliki nama insured valid')

    print(f'Total nama referensi unik (OSBAL, FACUL, RI SLIP): {len(refs):,}')
    print('Membangun model TF-IDF (char_wb 2-5 gram) dan k-NN...')
    vectorizer = TfidfVectorizer(analyzer='char_wb', ngram_range=(2, 5), min_df=1, sublinear_tf=True)
    matrix = vectorizer.fit_transform([r['key'] for r in refs])
    neighbors = NearestNeighbors(metric='cosine', algorithm='brute', n_neighbors=min(TOP_N, len(refs)))
    neighbors.fit(matrix)

    queries = {}
    details = []
    fac_result = col(result, 'CCOS_REF_CODE', 'FAC_CODE', 'FAC CODE')
    curr_result = col(result, 'CURR ORI', 'CURR')
    receipt = col(result, 'RECEIPT NO', 'RECEIPT_NO')

    for idx, row in selected.iterrows():
        row_variants = list(variants(row, qcols, 'SUSPEND'))
        if not row_variants and ori_col:
            row_variants = list(variants(row, [ori_col], 'SUSPEND'))
        for column, raw, key in row_variants:
            q = queries.setdefault(key, {'name': raw, 'key': key, 'receipts': set(),
                                         'scenarios': set(), 'flags': set(), 'columns': set()})
            q['receipts'].add(idx)
            q['columns'].add(column)
            q['scenarios'].add(txt(row[scenario_col]) if scenario_col else '')
            q['flags'].add(txt(row[flag_col]) if flag_col else '')
            details.append({
                'EXCEL_ROW': idx + 2,
                'RECEIPT_NO': txt(row[receipt]) if receipt else '',
                'INSURED_VARIANT': raw,
                'INSURED_COLUMN': column,
                'INSURED_KEY': key,
                'INSURED_ORI': txt(row[ori_col]) if ori_col else '',
                'SKENARIO': txt(row[scenario_col]) if scenario_col else '',
                'FLAG_PROD': txt(row[flag_col]) if flag_col else '',
                'FAC_CODE_HASIL': txt(row[fac_result]) if fac_result else '',
                'CURR_ORI': txt(row[curr_result]) if curr_result else '',
            })

    print(f'Total nama query unik yang dicari: {len(queries):,}')
    candidates = []
    if queries:
        all_queries = list(queries.values())
        print('Menghitung cosine distance & ranking Top-10...')
        distances, indices = neighbors.kneighbors(vectorizer.transform([q['key'] for q in all_queries]))
        for q, ds, ids in zip(all_queries, distances, indices):
            qt = set(q['key'].split())
            for rank, (distance, ri) in enumerate(zip(ds, ids), 1):
                ref = refs[int(ri)]
                common = qt & set(ref['key'].split())
                score = round(max(0.0, 1.0 - float(distance)) * 100, 2)
                if score < MIN_SCORE:
                    continue
                rt = set(ref['key'].split())
                risks = []
                if (qt ^ rt) & {'SYARIAH', 'SHARIA'}:
                    risks.append('SYARIAH_DIFFERENCE')
                if (qt ^ rt) & {'GROUP', 'GRUP'}:
                    risks.append('GROUP_DIFFERENCE')
                if (qt ^ rt) & {'PALM', 'SAWIT'}:
                    risks.append('BUSINESS_WORD_DIFFERENCE')
                if len(common - {'INDONESIA', 'GROUP', 'GRUP'}) <= 1:
                    risks.append('WEAK_SHARED_TOKENS')
                if len(ref['codes']) > 1:
                    risks.append('REFERENCE_MULTI_FAC_CODE')

                candidates.append({
                    'INSURED_SUSPEND': q['name'],
                    'NORMALIZED_SUSPEND': q['key'],
                    'CLEAN_COLUMNS': ', '.join(sorted(q['columns'])),
                    'JUMLAH_RECEIPT': len(q['receipts']),
                    'SKENARIO_ASAL': ' | '.join(sorted(x for x in q['scenarios'] if x)),
                    'FLAG_ASAL': ' | '.join(sorted(x for x in q['flags'] if x)),
                    'RANK': rank,
                    'TFIDF_SCORE': score,
                    'KANDIDAT_INSURED': ref['name'],
                    'KANDIDAT_NORMALIZED': ref['key'],
                    'SUMBER': ref['source'],
                    'SHARED_TOKENS': ', '.join(sorted(common)),
                    'RISK_FLAGS': ' | '.join(risks),
                    'FAC_CODE_REFERENSI': ', '.join(sorted(ref['codes'])),
                    'JUMLAH_FAC_CODE_REFERENSI': len(ref['codes']),
                    'CURRENCY_REFERENSI': ', '.join(sorted(ref['currencies'])),
                    'JUMLAH_REFERENSI': ref['count'],
                    'KOLOM_REFERENSI': ', '.join(sorted(ref['columns'])),
                    'STATUS_VALIDASI': 'PENDING',
                    'CATATAN_REVIEW': '',
                    'KEPUTUSAN_FAC_CODE': ''
                })

    summary = [
        {'METRIK': 'Perusahaan', 'NILAI': 'TOKIO MARINE'},
        {'METRIK': 'Total baris matching', 'NILAI': len(result)},
        {'METRIK': 'Baris terpilih untuk audit (Insured/Unmatching)', 'NILAI': len(selected)},
        {'METRIK': 'Nama Insured clean unik yang diaudit', 'NILAI': len(queries)},
        *stats,
        {'METRIK': 'Total nama referensi unik seluruh sumber', 'NILAI': len(refs)},
        {'METRIK': 'Jumlah kandidat TF-IDF dihasilkan', 'NILAI': len(candidates)},
        {'METRIK': 'Top N per nama Insured clean', 'NILAI': TOP_N}
    ]

    guide = [
        {'PANDUAN': 'Sumber kandidat: OSBAL, FACUL, RI SLIP (Database Rekap RI Slip), dan BORDERO.'},
        {'PANDUAN': 'Bordero Marine Cargo Tokio dimuat dari tokio_output_bordero.xlsx.'},
        {'PANDUAN': 'Nama Suspend, OSBAL, dan FACUL membaca varian nama clean 1–5; nama raw sebagai fallback.'},
        {'PANDUAN': 'Setiap varian nama insured Suspend mendapat maksimum 10 kandidat gabungan seluruh sumber.'},
        {'PANDUAN': 'TFIDF_SCORE (0–100) adalah ranking kemiripan teks, bukan keputusan matching otomatis.'},
        {'PANDUAN': 'Kolom STATUS_VALIDASI, CATATAN_REVIEW, dan KEPUTUSAN_FAC_CODE disediakan untuk review tim.'},
        {'PANDUAN': 'Audit ini mandiri dan tidak mengubah data matching produksi.'}
    ]

    OUT.mkdir(parents=True, exist_ok=True)
    print(f'Menyimpan output audit ke: {DEST} ...')
    with pd.ExcelWriter(DEST, engine='openpyxl') as writer:
        for sheet, rows in [
            ('AUDIT_SUMMARY', summary),
            ('TFIDF_CANDIDATES', candidates),
            ('RECEIPT_DETAILS', details),
            ('CARA_PAKAI', guide)
        ]:
            frame = pd.DataFrame(rows)
            if frame.empty:
                frame = pd.DataFrame({'INFO': ['Tidak ada data yang memenuhi filter']})
            frame.to_excel(writer, sheet_name=sheet, index=False)
            ws = writer.sheets[sheet]
            ws.freeze_panes = 'A2'
            ws.auto_filter.ref = ws.dimensions
            from openpyxl.styles import Font, PatternFill
            for cell in ws[1]:
                cell.font = Font(color='FFFFFF', bold=True)
                cell.fill = PatternFill('solid', fgColor='1F4E78')
            for col_cells in ws.columns:
                max_len = max(len(str(cell.value or '')) for cell in list(col_cells)[:150]) + 2
                ws.column_dimensions[col_cells[0].column_letter].width = min(52, max(14, max_len))

    print(f'\n[SELESAI] Hasil audit tersimpan di: {DEST}')
    print(f'  Baris diaudit : {len(selected):,}')
    print(f'  Query unik    : {len(queries):,}')
    print(f'  Kandidat      : {len(candidates):,}')


if __name__ == '__main__':
    main()
