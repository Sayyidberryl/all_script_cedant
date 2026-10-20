"""Audit kandidat nama INSURED untuk MAG menggunakan algoritma TF-IDF N-Gram & Hybrid Fuzzy Matching.
Run: python "TF IDF MAG.py"
Input: [16oct2026] output_suspend_mag.xlsx / matching_mag.xlsx, data/mag_output_osbal.xlsx, data/mag_output_facul.xlsx
Output: output/mag_insured_audit_v1.xlsx
Tidak mengubah file input, FAC_CODE, atau hasil matching.
"""
import sys
# Bypass modul opsional yang dapat crash dengan NumPy 2.x
sys.modules['numexpr'] = None
sys.modules['bottleneck'] = None

from pathlib import Path
import re
import unicodedata
from collections import Counter, defaultdict
from difflib import SequenceMatcher
import pandas as pd
import numpy as np

# ─────────────────────────────────────────────────────────────────────────────
# KONFIGURASI FILE & PARAMETER
# ─────────────────────────────────────────────────────────────────────────────
ROOT = Path(__file__).resolve().parent
DATA_DIR = ROOT / 'data'
OUT_DIR = ROOT / 'output'

def find_first_existing(candidates: list) -> Path:
    for c in candidates:
        if c.is_file():
            return c
    return candidates[0]

MATCH_FILE = find_first_existing([
    ROOT / '[16oct2026] output_suspend_mag.xlsx',
    ROOT / 'matching_mag.xlsx',
    OUT_DIR / 'matching_mag.xlsx',
    DATA_DIR / 'mag_output_suspend.xlsx'
])

OSBAL_FILE = find_first_existing([
    DATA_DIR / 'mag_output_osbal.xlsx',
    ROOT / 'mag_output_osbal.xlsx',
    OUT_DIR / 'mag_output_osbal.xlsx'
])

FACUL_FILE = find_first_existing([
    DATA_DIR / 'mag_output_facul.xlsx',
    ROOT / 'mag_output_facul.xlsx',
    OUT_DIR / 'mag_output_facul.xlsx'
])

AUDIT_FILE = OUT_DIR / 'mag_insured_audit_v1.xlsx'

TOP_N = 5
MIN_SCORE = 30.0  # Threshold kandidat masuk daftar audit

# RapidFuzz / Difflib ratio
try:
    from rapidfuzz import fuzz
    def char_ratio(a: str, b: str) -> float:
        return float(fuzz.ratio(a, b))
except ImportError:
    def char_ratio(a: str, b: str) -> float:
        return 100.0 * SequenceMatcher(None, a, b).ratio()

# Scikit-learn TF-IDF Vectorizer
try:
    from sklearn.feature_extraction.text import TfidfVectorizer
    from sklearn.metrics.pairwise import cosine_similarity
    SKLEARN_AVAILABLE = True
except ImportError:
    SKLEARN_AVAILABLE = False

COMMON_ENTITIES = {
    'PT', 'CV', 'UD', 'TBK', 'PERSERO', 'LTD', 'INC', 'CO', 'CORP', 'CORPORATION',
    'NV', 'BV', 'GMBH', 'SDN', 'BHD', 'LLC', 'PTE', 'LIMITED'
}

SPECIAL_FLAGS = {
    'SYARIAH', 'SHARIA', 'GROUP', 'GRUP', 'TBK', 'PERSERO', 'PALM', 'SAWIT',
    'MINING', 'COAL', 'ENERGY', 'LOGISTICS', 'TRANSPORT', 'FINANCE'
}

# ─────────────────────────────────────────────────────────────────────────────
# NORMALISASI & UTILITY
# ─────────────────────────────────────────────────────────────────────────────
def clean_text(x) -> str:
    if pd.isna(x) or x is None:
        return ''
    x = unicodedata.normalize('NFKC', str(x)).upper()
    x = re.sub(r'[^A-Z0-9]+', ' ', x)
    return ' '.join(x.split())

def core_name(x) -> str:
    tokens = clean_text(x).split()
    filtered = [w for w in tokens if w not in COMMON_ENTITIES]
    return ' '.join(filtered) if filtered else clean_text(x)

def name_tokens(x) -> set:
    return set(core_name(x).split())

def flag_tokens(x) -> set:
    return set(clean_text(x).split()) & SPECIAL_FLAGS

def best_col(df: pd.DataFrame, names: list):
    lut = {str(c).upper().strip(): c for c in df.columns}
    for n in names:
        if n.upper() in lut:
            return lut[n.upper()]
    return None

def find_insured_cols(df: pd.DataFrame, preferred: list) -> list:
    cols = []
    for c in preferred:
        match = best_col(df, [c])
        if match is not None and match not in cols:
            cols.append(match)
    for c in df.columns:
        cu = str(c).upper().strip()
        if ('INSURED' in cu or 'CLEAN INSURED' in cu) and c not in cols:
            cols.append(c)
    return cols

def str_val(v) -> str:
    if pd.isna(v) or v is None:
        return ''
    s = str(v).strip()
    if s.endswith('.0'):
        s = s[:-2]
    return s

def calculate_reason(a: str, b: str) -> str:
    fa, fb = flag_tokens(a), flag_tokens(b)
    tags = []
    if ('SYARIAH' in fa or 'SHARIA' in fa) != ('SYARIAH' in fb or 'SHARIA' in fb):
        tags.append('SYARIAH_DIFFERENCE')
    if ('GROUP' in fa or 'GRUP' in fa) != ('GROUP' in fb or 'GRUP' in fb):
        tags.append('GROUP_NAME_REVIEW')
    if ('SAWIT' in fa or 'PALM' in fa) != ('SAWIT' in fb or 'PALM' in fb):
        tags.append('BUSINESS_NAME_DIFFERENCE')
    if fa.symmetric_difference(fb) & {'TBK', 'PERSERO'}:
        tags.append('LEGAL_FLAG_DIFFERENCE')
    
    ca, cb = core_name(a), core_name(b)
    if ca == cb and ca:
        tags.append('SAME_CORE_NAME')
    elif name_tokens(a) & name_tokens(b):
        tags.append('SHARED_NAME_TOKENS')
    else:
        tags.append('POSSIBLE_TYPO_OR_ALIAS')
    return ', '.join(tags)

def read_excel_file(path: Path) -> pd.DataFrame:
    if not path.is_file():
        raise FileNotFoundError(f'File tidak ditemukan: {path}')
    print(f'Membaca: {path.name} ...')
    return pd.read_excel(path, dtype=object)

# ─────────────────────────────────────────────────────────────────────────────
# ENGINE TF-IDF & CANDIDATE RANKING
# ─────────────────────────────────────────────────────────────────────────────
def main():
    print('=' * 80)
    print('   AUDIT INSURED NAME RECONCILIATION VIA TF-IDF (CEDANT: MAG)')
    print('=' * 80)
    
    if not MATCH_FILE.is_file():
        raise FileNotFoundError(f'File suspend matching MAG tidak ditemukan di: {MATCH_FILE}')
    
    result_df = read_excel_file(MATCH_FILE)
    osbal_df  = read_excel_file(OSBAL_FILE)
    facul_df  = read_excel_file(FACUL_FILE)
    
    # Deteksi kolom matching suspend
    insured_col = best_col(result_df, ['INSURED_ORI', 'INSURED', 'clean insured 1', 'INSURED_1'])
    scenario_col = best_col(result_df, ['SKENARIO', 'SCENARIO'])
    
    if insured_col is None or scenario_col is None:
        raise ValueError(f'Kolom insured atau skenario tidak ditemukan di {MATCH_FILE.name}')
    
    # Filter row suspend yang perlu diaudit (Unmatching atau Insured fallback)
    scenario_series = result_df[scenario_col].fillna('').astype(str)
    mask = scenario_series.str.contains('Unmatching|UNMATCHED|Insured', case=False, regex=True)
    selected_df = result_df.loc[mask].copy()
    
    if selected_df.empty:
        print('Tidak ditemukan data unmatching / insured matching yang memerlukan audit.')
        return
    
    selected_df['_audit_name'] = selected_df[insured_col].map(clean_text)
    selected_df = selected_df[selected_df['_audit_name'] != '']
    
    # Agregasi nama suspend unik
    groups = []
    for name_key, sub in selected_df.groupby('_audit_name', dropna=False):
        scenarios = ', '.join(sorted(set(sub[scenario_col].fillna('').astype(str))))
        priority = (
            'MATCH_BY_INSURED_REVIEW'
            if 'insured' in scenarios.lower() and 'unmatching' not in scenarios.lower()
            else 'UNMATCHING_REVIEW'
        )
        groups.append({
            'name_key': name_key,
            'name_original': str_val(sub.iloc[0][insured_col]),
            'receipt_count': len(sub),
            'scenarios': scenarios,
            'priority': priority
        })
    
    print(f'Total baris matching masuk audit: {len(selected_df):,} | Nama unik: {len(groups):,}')
    
    # Bangun referensi nama dari OSBAL dan FACUL
    sources = [
        ('OSBAL', osbal_df, ['insured_ori', 'clean insured 1', 'clean insured 2', 'clean insured 3', 'clean insured 4', 'clean insured 5', 'CCOS_COMP_NAME']),
        ('FACUL', facul_df, ['insured_ori', 'clean insured 1', 'clean insured 2', 'clean insured 3', 'clean insured 4', 'clean insured 5', 'COMP_NAME', 'BUSINESS PARTNERS'])
    ]
    
    ref_map = defaultdict(lambda: {'source': set(), 'fac_codes': set(), 'count': 0})
    for src, frame, pref_names in sources:
        cand_cols = find_insured_cols(frame, pref_names)
        fac_col = best_col(frame, ['CCOS_REF_CODE', 'FAC_CODE', 'FAC. CODE'])
        for _, row in frame.iterrows():
            vals = {str_val(row[c]) for c in cand_cols if str_val(row[c])}
            code = str_val(row[fac_col]) if fac_col else ''
            for v in vals:
                k = clean_text(v)
                if not k or len(k) < 3:
                    continue
                entry = ref_map[(src, k)]
                entry['source'].add(v)
                if code:
                    entry['fac_codes'].add(code)
                entry['count'] += 1
                
    refs = []
    for (src, k), val in ref_map.items():
        primary_name = sorted(val['source'], key=lambda x: (-len(x), x))[0]
        refs.append({
            'source': src,
            'name': primary_name,
            'name_key': k,
            'fac_codes': ', '.join(sorted(val['fac_codes'])[:15]),
            'ref_count': val['count']
        })
        
    print(f'Total entitas referensi unik (OSBAL & FACUL): {len(refs):,}')
    
    # Inisialisasi TF-IDF N-Gram Vectorizer
    tfidf_sim_matrix = None
    if SKLEARN_AVAILABLE and len(refs) > 0 and len(groups) > 0:
        print('Menghitung Matrix TF-IDF Character N-Gram (3, 5)...')
        ref_texts = [r['name_key'] for r in refs]
        query_texts = [g['name_key'] for g in groups]
        
        # Fit vectorizer pada kombinasi query dan referensi
        vectorizer = TfidfVectorizer(analyzer='char_wb', ngram_range=(3, 5), min_df=1)
        ref_matrix = vectorizer.fit_transform(ref_texts)
        query_matrix = vectorizer.transform(query_texts)
        
        # Cosine similarity matrix: queries x refs
        tfidf_sim_matrix = cosine_similarity(query_matrix, ref_matrix)
    else:
        print('Scikit-learn tidak aktif / data kosong; fallback ke token overlap & fuzzy ratio.')
    
    # Scoring & Pemilihan Kandidat Teratas (Vectorized Pre-filtering)
    candidates = []
    total_q = len(groups)
    print(f'Mengevaluasi {total_q} entitas query terhadap top kandidat referensi...')

    for q_idx, group in enumerate(groups):
        a_orig = group['name_original']
        a_clean = group['name_key']
        ta = name_tokens(a_clean)
        ca = core_name(a_clean)
        
        # Pre-select candidate indices using TF-IDF similarity vector
        if tfidf_sim_matrix is not None:
            row_sims = tfidf_sim_matrix[q_idx]
            cand_indices = np.where(row_sims >= 0.15)[0]
            if len(cand_indices) > 50:
                top_indices = np.argpartition(row_sims[cand_indices], -50)[-50:]
                cand_indices = cand_indices[top_indices]
        else:
            cand_indices = [r_idx for r_idx, ref in enumerate(refs) if (ta & name_tokens(ref['name_key']))]

        ranked = []
        for r_idx in cand_indices:
            ref = refs[r_idx]
            b_clean = ref['name_key']
            tb = name_tokens(b_clean)
            cb = core_name(b_clean)
            
            # 1. TF-IDF Score (0 - 100)
            if tfidf_sim_matrix is not None:
                tfidf_score = round(float(tfidf_sim_matrix[q_idx, r_idx]) * 100.0, 1)
            else:
                tfidf_score = 0.0
                
            # 2. Token Jaccard Overlap (0 - 100)
            overlap_score = (100.0 * len(ta & tb) / len(ta | tb)) if (ta | tb) else 0.0
            
            # 3. Fuzzy Levenshtein Ratio (0 - 100)
            fuzz_score = char_ratio(ca, cb) if ca and cb else 0.0
            
            # 4. Hybrid Composite Score
            if tfidf_sim_matrix is not None:
                final_score = round(0.50 * tfidf_score + 0.30 * fuzz_score + 0.20 * overlap_score, 1)
            else:
                final_score = round(0.70 * fuzz_score + 0.30 * overlap_score, 1)
                
            if final_score < MIN_SCORE:
                continue
                
            ranked.append({
                'score': final_score,
                'tfidf': tfidf_score,
                'ref': ref
            })
            
        ranked.sort(key=lambda z: (-z['score'], -z['tfidf'], z['ref']['source'], z['ref']['name']))
        
        # Tentukan Action Recommendation berdasarkan margin & score
        top_candidates = ranked[:TOP_N]
        if top_candidates:
            best_score = top_candidates[0]['score']
            second_score = top_candidates[1]['score'] if len(top_candidates) > 1 else 0.0
            margin = best_score - second_score
            
            for rank, item in enumerate(top_candidates, 1):
                s = item['score']
                tfidf_val = item['tfidf']
                ref = item['ref']
                reasons = calculate_reason(a_orig, ref['name'])
                
                # Rekomendasi aksi
                if rank == 1 and s >= 90.0 and 'SYARIAH_DIFFERENCE' not in reasons:
                    recommendation = 'AUTO_MATCH_INSURED_HIGH'
                elif rank == 1 and s >= 80.0 and margin >= 15.0 and 'SYARIAH_DIFFERENCE' not in reasons:
                    recommendation = 'AUTO_MATCH_INSURED_CONFIDENT'
                else:
                    recommendation = 'REVIEW_MANUAL'
                    
                candidates.append({
                    'INSURED_SUSPEND': a_orig,
                    'SKENARIO_ASAL': group['scenarios'],
                    'PRIORITAS': group['priority'],
                    'JUMLAH_RECEIPT': group['receipt_count'],
                    'RANK': rank,
                    'KANDIDAT_INSURED': ref['name'],
                    'SUMBER': ref['source'],
                    'TFIDF_SCORE': tfidf_val,
                    'SKOR_KOMPOSIT': s,
                    'REKOMENDASI_AKSI': recommendation,
                    'ALASAN_REVIEW': reasons,
                    'FAC_CODE_REFERENSI_INFO': ref['fac_codes'],
                    'STATUS_VALIDASI': 'PENDING',
                    'CATATAN_REVIEW': '',
                    'KEPUTUSAN_FAC_CODE': ''
                })
        else:
            candidates.append({
                'INSURED_SUSPEND': a_orig,
                'SKENARIO_ASAL': group['scenarios'],
                'PRIORITAS': group['priority'],
                'JUMLAH_RECEIPT': group['receipt_count'],
                'RANK': '',
                'KANDIDAT_INSURED': '',
                'SUMBER': '',
                'TFIDF_SCORE': 0.0,
                'SKOR_KOMPOSIT': 0.0,
                'REKOMENDASI_AKSI': 'UNRESOLVED_NO_MATCH',
                'ALASAN_REVIEW': 'NO_CANDIDATE_ABOVE_THRESHOLD',
                'FAC_CODE_REFERENSI_INFO': '',
                'STATUS_VALIDASI': 'PENDING',
                'CATATAN_REVIEW': '',
                'KEPUTUSAN_FAC_CODE': ''
            })
            
    # Sheet 1: AUDIT_SUMMARY
    summary_df = pd.DataFrame([
        {'METRIK': 'Nama File Matching Input', 'NILAI': MATCH_FILE.name},
        {'METRIK': 'Total Baris Hasil Matching', 'NILAI': len(result_df)},
        {'METRIK': 'Baris Suspend Masuk Audit (Unmatching/Insured)', 'NILAI': len(selected_df)},
        {'METRIK': 'Nama Tertanggung Unik Masuk Audit', 'NILAI': len(groups)},
        {'METRIK': 'Total Entitas Referensi OSBAL & FACUL', 'NILAI': len(refs)},
        {'METRIK': 'Total Rekomendasi Kandidat Diekstrak', 'NILAI': len(candidates)},
        {'METRIK': 'Algoritma Scoring', 'NILAI': 'TF-IDF Character N-Gram (3,5) + RapidFuzz Levenshtein + Token Overlap'},
        {'METRIK': 'Catatan Penting', 'NILAI': 'Script ini tidak mengubah hasil matching orisinil MAG secara langsung.'},
    ])
    
    # Sheet 3: RECEIPT_DETAILS
    detail_cols = [c for c in [
        'RECEIPT NO', 'RECEIPT DATE', 'INSURED_ORI', 'SKENARIO',
        'CCOS_REF_CODE', 'POLIS_ORI', 'POLIS_CLN', 'SERTIF_CLN',
        'SLIP_NO_ORI', 'SLIP_NO_CLN', 'CURR ORI', 'AMOUNT ORI'
    ] if c in selected_df.columns]
    details_df = selected_df[detail_cols].copy()
    details_df.insert(0, 'INSURED_KEY_AUDIT', selected_df['_audit_name'].values)
    
    # Sheet 4: CARA_PAKAI
    guide_df = pd.DataFrame([
        {'LANGKAH': 1, 'INSTRUKSI': 'Pastikan matching_mag.py / [16oct2026] output_suspend_mag.xlsx sudah selesai dieksekusi.'},
        {'LANGKAH': 2, 'INSTRUKSI': 'Jalankan python "TF IDF MAG.py" dari folder mag_production_script.'},
        {'LANGKAH': 3, 'INSTRUKSI': 'Buka tab INSURED_CANDIDATES pada output/mag_insured_audit_v1.xlsx.'},
        {'LANGKAH': 4, 'INSTRUKSI': 'Filter REKOMENDASI_AKSI = AUTO_MATCH_INSURED_HIGH / AUTO_MATCH_INSURED_CONFIDENT untuk verifikasi cepat.'},
        {'LANGKAH': 5, 'INSTRUKSI': 'Periksa kolom ALASAN_REVIEW untuk perbedaan badan hukum, flag Syariah, atau group.'},
        {'LANGKAH': 6, 'INSTRUKSI': 'Isi STATUS_VALIDASI (APPROVED / REJECTED) dan KEPUTUSAN_FAC_CODE jika disetujui.'},
        {'LANGKAH': 7, 'INSTRUKSI': 'Gunakan hasil review untuk meng-update tabel database fsi_db atau master matching.'},
    ])
    
    # Simpan ke Excel dengan OpenPyXL styling
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(AUDIT_FILE, engine='openpyxl') as writer:
        summary_df.to_excel(writer, sheet_name='AUDIT_SUMMARY', index=False)
        pd.DataFrame(candidates).to_excel(writer, sheet_name='INSURED_CANDIDATES', index=False)
        details_df.to_excel(writer, sheet_name='RECEIPT_DETAILS', index=False)
        guide_df.to_excel(writer, sheet_name='CARA_PAKAI', index=False)
        
        from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
        thin = Side(border_style='thin', color='CCCCCC')
        cell_border = Border(top=thin, left=thin, right=thin, bottom=thin)
        
        for ws in writer.book:
            ws.freeze_panes = 'A2'
            ws.auto_filter.ref = ws.dimensions
            # Header styling
            for cell in ws[1]:
                cell.font = Font(color='FFFFFF', bold=True, size=11)
                cell.fill = PatternFill('solid', fgColor='283F64')
                cell.alignment = Alignment(horizontal='center', vertical='center')
            # Column auto-width
            for col in ws.columns:
                letter = col[0].column_letter
                max_len = max(len(str(c.value or '')) for c in list(col)[:200])
                ws.column_dimensions[letter].width = min(54, max(14, max_len + 3))
                
    print('=' * 80)
    print(f'Audit Selesai Berhasil! Output disimpan ke: {AUDIT_FILE}')
    print(f'Total kandidat insured diekstrak: {len(candidates):,}')
    print('=' * 80)

if __name__ == '__main__':
    main()
