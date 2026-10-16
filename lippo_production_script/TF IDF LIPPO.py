"""Audit kandidat nama INSURED, TERPISAH dari Script 3.
Run: python lippo_insured_audit_v1.py
Input: output/final_output_v1.xlsx, lippo_output_osbal_V2.xlsx,
       lippo_output_facul_V3.xlsx
Output: output/lippo_insured_audit_v1.xlsx
Tidak mengubah file input, FAC_CODE, atau hasil matching.
"""
from pathlib import Path
import re
import unicodedata
from collections import Counter, defaultdict
from difflib import SequenceMatcher
import pandas as pd

ROOT = Path(__file__).resolve().parent
OUT = ROOT / 'output'
MATCH_FILE = OUT / 'final_output_v1.xlsx'
OSBAL_FILE = OUT / 'lippo_output_osbal_V2.xlsx'
FACUL_FILE = OUT / 'lippo_output_facul_V3.xlsx'
AUDIT_FILE = OUT / 'lippo_insured_audit_v1.xlsx'
TOP_N = 5
# Skor ini HANYA untuk ranking kandidat; bukan ambang auto-match.
MIN_SCORE = 35

try:
    from rapidfuzz import fuzz
    def ratio(a, b): return float(fuzz.ratio(a, b))
except ImportError:
    def ratio(a, b): return 100.0 * SequenceMatcher(None, a, b).ratio()

COMMON = {'PT', 'CV', 'UD', 'TBK', 'PERSERO', 'LTD', 'INC', 'CO', 'CORP'}
FLAGS = {'SYARIAH', 'SHARIA', 'GROUP', 'GRUP', 'TBK', 'PERSERO', 'PALM', 'SAWIT'}

def clean(x):
    if pd.isna(x): return ''
    x = unicodedata.normalize('NFKC', str(x)).upper()
    x = re.sub(r'[^A-Z0-9]+', ' ', x)
    return ' '.join(x.split())

def core(x):
    return ' '.join(w for w in clean(x).split() if w not in COMMON)

def tokens(x):
    return set(core(x).split())

def flags(x):
    return set(clean(x).split()) & FLAGS

def best_col(df, names):
    lut = {str(c).upper().strip(): c for c in df.columns}
    for n in names:
        if n.upper() in lut: return lut[n.upper()]
    return None

def name_cols(df, preferred):
    cols = []
    for c in preferred:
        match = best_col(df, [c])
        if match is not None and match not in cols: cols.append(match)
    cols += [c for c in df.columns if 'INSURED_CLEAN_' in str(c).upper() and c not in cols]
    return cols

def textval(v):
    if pd.isna(v): return ''
    return str(v).strip()

def score(a, b):
    aa, bb = core(a), core(b)
    if not aa or not bb: return 0.0
    # Gabungan character similarity + overlap token, untuk ranking saja.
    ta, tb = tokens(a), tokens(b)
    overlap = 100 * len(ta & tb) / len(ta | tb) if ta | tb else 0
    return round(0.7 * ratio(aa, bb) + 0.3 * overlap, 1)

def reason(a, b):
    fa, fb = flags(a), flags(b)
    tags = []
    if ('SYARIAH' in fa or 'SHARIA' in fa) != ('SYARIAH' in fb or 'SHARIA' in fb):
        tags.append('SYARIAH_DIFFERENCE')
    if ('GROUP' in fa or 'GRUP' in fa) != ('GROUP' in fb or 'GRUP' in fb):
        tags.append('GROUP_NAME_REVIEW')
    if ('SAWIT' in fa) != ('SAWIT' in fb) or ('PALM' in fa) != ('PALM' in fb):
        tags.append('BUSINESS_NAME_DIFFERENCE')
    if core(a) == core(b): tags.append('SAME_CORE_NAME')
    elif tokens(a) & tokens(b): tags.append('SHARED_NAME_TOKENS')
    else: tags.append('POSSIBLE_TYPO_OR_ALIAS')
    if fa.symmetric_difference(fb) & {'TBK', 'PERSERO'}: tags.append('LEGAL_FLAG_DIFFERENCE')
    return ', '.join(tags)

def read(path):
    if not path.is_file(): raise FileNotFoundError(f'File belum ada: {path}\nJalankan Script 3 terlebih dahulu.')
    return pd.read_excel(path, dtype=object)

def main():
    result = read(MATCH_FILE)
    sources = [('OSBAL', read(OSBAL_FILE), ['INSURED', 'CCOS_INSURED', 'FAC_INSURED', 'INSURED_CLEAN_1']),
               ('FACUL', read(FACUL_FILE), ['FAC_INSURED', 'INSURED', 'INSURED_CLEAN_1'])]
    insured_col = best_col(result, ['INSURED_ORI', 'INSURED', 'INSURED_CLEAN_1'])
    scenario_col = best_col(result, ['SKENARIO'])
    if insured_col is None or scenario_col is None:
        raise ValueError('Kolom INSURED_ORI/INSURED dan SKENARIO tidak ditemukan di final_output_v1.xlsx')
    scenario = result[scenario_col].fillna('').astype(str)
    mask = scenario.str.contains('Unmatching|UNMATCHED|Insured', case=False, regex=True)
    selected = result.loc[mask].copy()
    selected['_audit_name'] = selected[insured_col].map(clean)
    selected = selected[selected['_audit_name'] != '']
    grouped = selected.groupby('_audit_name', dropna=False)
    groups = []
    for key, sub in grouped:
        scenarios = ', '.join(sorted(set(sub[scenario_col].fillna('').astype(str))))
        groups.append({'name_key': key, 'name_original': textval(sub.iloc[0][insured_col]),
                       'receipt_count': len(sub), 'scenarios': scenarios,
                       'priority': 'MATCH_BY_INSURED_REVIEW' if 'insured' in scenarios.lower() and 'unmatching' not in scenarios.lower() else 'UNMATCHING_REVIEW'})
    ref_map = defaultdict(lambda: {'source': set(), 'fac_codes': set(), 'count': 0})
    for src, frame, names in sources:
        candidates = name_cols(frame, names)
        if not candidates: raise ValueError(f'Kolom insured tidak ditemukan pada {src}: {list(frame.columns)[:25]}')
        fac_col = best_col(frame, ['CCOS_REF_CODE', 'FAC_CODE'])
        for _, row in frame.iterrows():
            values = {textval(row[c]) for c in candidates if textval(row[c])}
            # Nama mentah lebih disukai untuk interpretasi, tetapi semua varian ikut diindeks.
            for name in values:
                k = clean(name)
                if not k: continue
                ref = ref_map[(src, k)]
                ref['source'].add(name)
                if fac_col:
                    code = textval(row[fac_col])
                    if code: ref['fac_codes'].add(code)
                ref['count'] += 1
    refs = []
    for (src, k), val in ref_map.items():
        refs.append({'source': src, 'name': sorted(val['source'], key=lambda x: (-len(x), x))[0],
                     'name_key': k, 'fac_codes': ', '.join(sorted(val['fac_codes'])[:15]),
                     'ref_count': val['count']})
    candidates = []
    for group in groups:
        # Blocking: nama dengan token sama diprioritaskan, tetapi typo tetap dicari
        # lewat similarity. Tidak ada keputusan FAC_CODE otomatis.
        ranked = []
        a = group['name_original']
        for ref in refs:
            b = ref['name']
            s = score(a, b)
            if s < MIN_SCORE: continue
            ranked.append((s, ref))
        ranked.sort(key=lambda z: (-z[0], z[1]['source'], z[1]['name']))
        for rank, (s, ref) in enumerate(ranked[:TOP_N], 1):
            candidates.append({'INSURED_SUSPEND': a, 'SKENARIO_ASAL': group['scenarios'],
                               'PRIORITAS': group['priority'], 'JUMLAH_RECEIPT': group['receipt_count'],
                               'RANK': rank, 'KANDIDAT_INSURED': ref['name'], 'SUMBER': ref['source'],
                               'SKOR_KEMIRIPAN': s, 'ALASAN_REVIEW': reason(a, ref['name']),
                               'FAC_CODE_REFERENSI_INFO': ref['fac_codes'],
                               'STATUS_VALIDASI': 'PENDING', 'CATATAN_REVIEW': '',
                               'KEPUTUSAN_FAC_CODE': ''})
        if not ranked:
            candidates.append({'INSURED_SUSPEND': a, 'SKENARIO_ASAL': group['scenarios'],
                               'PRIORITAS': group['priority'], 'JUMLAH_RECEIPT': group['receipt_count'],
                               'RANK': '', 'KANDIDAT_INSURED': '', 'SUMBER': '',
                               'SKOR_KEMIRIPAN': '', 'ALASAN_REVIEW': 'NO_CANDIDATE_ABOVE_THRESHOLD',
                               'FAC_CODE_REFERENSI_INFO': '', 'STATUS_VALIDASI': 'PENDING',
                               'CATATAN_REVIEW': '', 'KEPUTUSAN_FAC_CODE': ''})
    summary = pd.DataFrame([
        {'METRIK': 'Baris hasil Script 3', 'NILAI': len(result)},
        {'METRIK': 'Baris masuk audit', 'NILAI': len(selected)},
        {'METRIK': 'Nama insured unik masuk audit', 'NILAI': len(groups)},
        {'METRIK': 'Nama referensi unik OSBAL/FACUL', 'NILAI': len(refs)},
        {'METRIK': 'Total kandidat untuk review', 'NILAI': len(candidates)},
        {'METRIK': 'CATATAN', 'NILAI': 'Skor bukan keputusan matching; FAC_CODE tidak diubah.'},
    ])
    detail_cols = [c for c in ['RECEIPT NO', 'RECEIPT DATE', 'INSURED_ORI', 'SKENARIO',
                               'CCOS_REF_CODE', 'POLIS', 'SLIP_NO', 'CURR ORI'] if c in selected.columns]
    details = selected[detail_cols].copy()
    details.insert(0, 'INSURED_KEY_AUDIT', selected['_audit_name'].values)
    guide = pd.DataFrame([
        {'LANGKAH': 1, 'INSTRUKSI': 'Jalankan Script 3 sampai output/final_output_v1.xlsx tersedia.'},
        {'LANGKAH': 2, 'INSTRUKSI': 'Jalankan python lippo_insured_audit_v1.py dari folder yang sama.'},
        {'LANGKAH': 3, 'INSTRUKSI': 'Buka INSURED_CANDIDATES; filter PRIORITAS dan lihat kandidat per nama unik.'},
        {'LANGKAH': 4, 'INSTRUKSI': 'Bandingkan dengan RECEIPT_DETAILS dan polis/slip asli bila diperlukan.'},
        {'LANGKAH': 5, 'INSTRUKSI': 'Isi STATUS_VALIDASI: SAME_ENTITY / SAME_GROUP_ONLY / DIFFERENT_ENTITY / UNRESOLVED.'},
        {'LANGKAH': 6, 'INSTRUKSI': 'Isi CATATAN_REVIEW. Jangan isi keputusan FAC_CODE berdasarkan skor semata.'},
        {'LANGKAH': 7, 'INSTRUKSI': 'Simpan salinan audit yang sudah direview; script tidak membacanya kembali otomatis.'},
    ])
    OUT.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(AUDIT_FILE, engine='openpyxl') as writer:
        summary.to_excel(writer, sheet_name='AUDIT_SUMMARY', index=False)
        pd.DataFrame(candidates, columns=['INSURED_SUSPEND','SKENARIO_ASAL','PRIORITAS','JUMLAH_RECEIPT',
            'RANK','KANDIDAT_INSURED','SUMBER','SKOR_KEMIRIPAN','ALASAN_REVIEW',
            'FAC_CODE_REFERENSI_INFO','STATUS_VALIDASI','CATATAN_REVIEW','KEPUTUSAN_FAC_CODE']).to_excel(writer, sheet_name='INSURED_CANDIDATES', index=False)
        details.to_excel(writer, sheet_name='RECEIPT_DETAILS', index=False)
        guide.to_excel(writer, sheet_name='CARA_PAKAI', index=False)
        from openpyxl.styles import Font, PatternFill
        for ws in writer.book:
            ws.freeze_panes = 'A2'
            ws.auto_filter.ref = ws.dimensions
            for cell in ws[1]:
                cell.font = Font(color='FFFFFF', bold=True)
                cell.fill = PatternFill('solid', fgColor='283F64')
            for col in ws.columns:
                letter = col[0].column_letter
                width = min(54, max(14, max(len(str(c.value or '')) for c in list(col)[:150]) + 2))
                ws.column_dimensions[letter].width = width
    print(f'Audit selesai: {AUDIT_FILE}')
    print(f'Baris audit: {len(selected):,}; nama unik: {len(groups):,}; kandidat: {len(candidates):,}')
    print('CATATAN: script ini TIDAK mengubah hasil matching atau FAC_CODE.')

if __name__ == '__main__':
    main()
