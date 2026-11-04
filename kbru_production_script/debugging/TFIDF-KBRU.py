"""Audit TF-IDF independen KBRU untuk skenario Insured dan Unmatching.
Jalankan: python kbru_insured_tfidf_audit_3_sumber.py
Dependensi: pandas numpy scikit-learn openpyxl
Hanya baca sumber dan menulis workbook audit; TIDAK mengubah matching/FAC CODE.
"""
from pathlib import Path
from collections import defaultdict
import re
import unicodedata
import numpy as np
import pandas as pd
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.neighbors import NearestNeighbors
from openpyxl.styles import Font, PatternFill

ROOT = Path(__file__).resolve().parent
OUT = ROOT / 'output'
MATCH = OUT / 'final_output_v1.xlsx'
SUSPEND = OUT / 'kbru_output_suspend.xlsx'  # enrichment opsional untuk insured clean 3-5
SOURCES = {
    'OSBAL': [OUT / 'kbru_output_osbal.xlsx'],
    'SLIPDB': [OUT / 'kbru_output_slipdb.xlsx', OUT / 'slipdb_clean_kbru.xlsx', OUT / 'ri slip.xlsx', OUT / 'slipdb.xlsx'],
    'FACUL': [OUT / 'kbru_output_facul.xlsx'],
}
DEST = OUT / 'kbru_insured_tfidf_audit_3_sumber.xlsx'
TOP_N = 10
MIN_SCORE = 0.0
# Sumber SLIPDB memang opsional dalam matching KBRU.
COMMON = {'PT','CV','UD','TBK','PERSERO','LTD','INC','LLC','CO','CORP','CORPORATION','COMPANY'}
DUMMY = {'TBA','VAR','VARIOUS','NONE','NULL','NAN','NA','SELISIH','UNMATCHING','UNMATCHED','0','00','000'}
NON_ENTITY = re.compile(r'SUSPEN|UTANG\s*PIUTANG|TREATY|STRUCTURE|EXCESS\s+OF\s+LOSS|WHOLE\s+ACCOUNT|DATA\s+SUSPENSE', re.I)


def txt(v):
    if v is None:
        return ''
    try:
        if pd.isna(v):
            return ''
    except (TypeError, ValueError):
        pass
    s = str(v).strip()
    return '' if s.upper() in {'NAN','NONE','NAT'} else s


def norm(v):
    s = unicodedata.normalize('NFKC', txt(v)).upper()
    s = re.sub(r'\bSINAR\s+MAS\b', 'SINARMAS', s)
    s = re.sub(r'\bPELINDO\b', 'PELABUHAN INDONESIA', s)
    s = re.sub(r'(?<=\b[A-Z])\.(?=[A-Z]\b)', '', s) if False else s
    # KBRU menggabungkan singkatan seperti S.A.R.I => SARI pada cleaning.
    s = re.sub(r'\b(?:[A-Z]\.){2,}[A-Z]?\.?', lambda m: m.group().replace('.',''), s)
    s = re.sub(r'[^A-Z0-9]+', ' ', s)
    return ' '.join(t for t in s.split() if t not in COMMON)


def valid_name(v):
    s = norm(v)
    return bool(s and len(s.replace(' ', '')) >= 3 and s not in DUMMY and not NON_ENTITY.search(s))


def col(df, *names):
    lookup = {re.sub(r'[^A-Z0-9]+','',str(c).upper()):c for c in df.columns}
    for name in names:
        hit = lookup.get(re.sub(r'[^A-Z0-9]+','',name.upper()))
        if hit is not None:
            return hit
    return None


def clean_cols(df):
    found = []
    for c in df.columns:
        k = re.sub(r'[^A-Z0-9]+','',str(c).upper())
        if re.fullmatch(r'(?:INSUREDCLEAN|INSUREDCLN|CLEANINSURED|FACINSURED|INSURED)[1-5]',k):
            found.append((int(k[-1]), str(c), c))
    return [x[2] for x in sorted(found)]


def insured_cols(df, source):
    clean = clean_cols(df)
    if clean:
        return clean
    raw = col(df,'FAC_INSURED_ORI','FAC_INSURED','INSURED_ORI','INSURED','INSURED NAME','INSURED_NAME','CCOS_INSURED')
    if raw is None:
        raise ValueError(f'{source}: kolom insured tidak tersedia. Kolom: {list(df.columns)}')
    return [raw]


def variants(row, cols):
    # Tidak split string yang sudah di-clean: KBRU punya aturan delimiter kompleks.
    for c in cols:
        name = txt(row[c])
        if valid_name(name):
            yield str(c),name,norm(name)


def read(path,label,required=True):
    if not path or not path.exists():
        if required:
            raise FileNotFoundError(f'{label} tidak ditemukan: {path}')
        return None
    # Beberapa template RI Slip memiliki header bergeser.
    for hdr in (0,1,2):
        df = pd.read_excel(path,header=hdr,dtype=object)
        if label == 'MATCH' or clean_cols(df) or col(df,'FAC_INSURED','INSURED_ORI','INSURED','CCOS_INSURED'):
            return df
    raise ValueError(f'{label}: header/kolom insured tidak ditemukan di {path}')


def find_source(paths):
    return next((p for p in paths if p.exists()),None)


def references(df,source):
    columns = insured_cols(df,source)
    codecol = col(df,'CCOS_REF_CODE','FAC_CODE','FAC CODE','FACODE')
    currcol = col(df,'CCOS_CURR','FAC_CURRENCY','CURR','CURR ORI','CURRENCY')
    records = {}
    for _,row in df.iterrows():
        for c,name,key in variants(row,columns):
            r = records.setdefault(key,dict(source=source,name=name,key=key,codes=set(),currencies=set(),count=0,columns=set()))
            r['count'] += 1
            r['columns'].add(c)
            if codecol and txt(row[codecol]):r['codes'].add(txt(row[codecol]))
            if currcol and txt(row[currcol]):r['currencies'].add(txt(row[currcol]).upper())
    return list(records.values())


def selected_rows(df):
    sce=col(df,'SKENARIO');flag=col(df,'FLAG_PROD')
    if not sce and not flag:raise ValueError('Output matching tidak memiliki SKENARIO/FLAG_PROD')
    ss=df[sce].fillna('').astype(str) if sce else pd.Series('',index=df.index)
    ff=df[flag].fillna('').astype(str) if flag else pd.Series('',index=df.index)
    mask=ss.str.contains('insured|unmatching|unmatched',case=False,regex=True) | ff.str.contains('unmatching|unmatched',case=False,regex=True)
    return df.loc[mask].copy(),sce,flag


def merge_suspend_variants(selected,suspend_df):
    """Tambahkan varian clean 3-5 jika hubungan ke output matching bisa dibuktikan unik.
    Tidak pernah join atas INSURED_ORI semata karena nama bisa berulang.
    """
    if suspend_df is None:
        return {},'Suspend clean tidak tersedia: menggunakan varian dari hasil matching saja.'
    cc=clean_cols(suspend_df)
    if not cc:
        return {},'Tidak ditemukan kolom insured clean pada suspend.'
    keys=['RECEIPT NO','CREDIT NOTES','DETAIL RINCIAN NO']
    pairs=[]
    for k in keys:
        a=col(selected,k);b=col(suspend_df,k)
        if a is not None and b is not None:pairs.append((a,b))
    if not pairs:
        return {},'Tidak ada kunci identitas bersama; enrichment suspend dilewati demi menghindari false join.'
    lookup=defaultdict(list)
    for _,row in suspend_df.iterrows():
        key=tuple(txt(row[b]).upper() for a,b in pairs)
        if any(key):lookup[key].append(row)
    extra={};ambiguous=0;linked=0
    for idx,row in selected.iterrows():
        key=tuple(txt(row[a]).upper() for a,b in pairs)
        matches=lookup.get(key,[]) if any(key) else []
        if len(matches)==1:
            extra[idx]=list(variants(matches[0],cc));linked+=1
        elif len(matches)>1:ambiguous+=1
    return extra,f'Enrichment suspend berhasil untuk {linked} baris; {ambiguous} baris dengan kunci duplikat dilewati.'


def main():
    result=read(MATCH,'MATCH')
    selected,scenario_col,flag_col=selected_rows(result)
    input_frames={};source_paths={};stats=[]
    for source,paths in SOURCES.items():
        path=find_source(paths)
        if path is None:
            if source=='SLIPDB':
                stats.append({'METRIK':'SLIPDB','NILAI':'Tidak ditemukan (opsional sesuai matching KBRU)'})
                continue
            raise FileNotFoundError(f'Sumber wajib {source} tidak ada; dicari: {paths}')
        input_frames[source]=read(path,source)
        source_paths[source]=str(path)
    refs=[]
    for source,df in input_frames.items():
        rr=references(df,source)
        refs.extend(rr)
        stats.append({'METRIK':f'Referensi unik {source}','NILAI':len(rr)})
    if not refs:raise ValueError('Tidak ada nama referensi valid.')
    extra,note=merge_suspend_variants(selected,read(SUSPEND,'SUSPEND',required=False))
    qcols=insured_cols(result,'MATCH')
    ori_col=col(result,'INSURED_ORI','FAC_INSURED','INSURED')
    fac_col=col(result,'CCOS_REF_CODE','FAC_CODE')
    curr_col=col(result,'CURR ORI','CURR')
    receipt_col=col(result,'RECEIPT NO')
    queries={};details=[]
    for idx,row in selected.iterrows():
        allvars=list(variants(row,qcols))+extra.get(idx,[])
        seen=set();allvars=[v for v in allvars if not (v[2] in seen or seen.add(v[2]))]
        if not allvars and ori_col:allvars=list(variants(row,[ori_col]))
        for c,name,key in allvars:
            q=queries.setdefault(key,dict(name=name,key=key,rows=set(),columns=set(),scenarios=set(),flags=set(),matched_codes=set(),currencies=set()))
            q['rows'].add(idx);q['columns'].add(c)
            sce=txt(row[scenario_col]) if scenario_col else ''
            flag=txt(row[flag_col]) if flag_col else ''
            q['scenarios'].add(sce);q['flags'].add(flag)
            if fac_col and txt(row[fac_col]):q['matched_codes'].add(txt(row[fac_col]))
            if curr_col and txt(row[curr_col]):q['currencies'].add(txt(row[curr_col]).upper())
            details.append({'EXCEL_ROW':int(idx)+2,'INSURED_VARIANT':name,'INSURED_COLUMN':c,'NORMALIZED_INSURED':key,
                'INSURED_ORI':txt(row[ori_col]) if ori_col else '', 'SKENARIO':sce,'FLAG_PROD':flag,
                'FAC_CODE_HASIL':txt(row[fac_col]) if fac_col else '',
                'CURR_ORI':txt(row[curr_col]) if curr_col else '',
                'RECEIPT_NO':txt(row[receipt_col]) if receipt_col else ''})
    vectorizer=TfidfVectorizer(analyzer='char_wb',ngram_range=(2,5),min_df=1,sublinear_tf=True)
    matrix=vectorizer.fit_transform([r['key'] for r in refs])
    model=NearestNeighbors(metric='cosine',algorithm='brute',n_neighbors=min(TOP_N,len(refs)))
    model.fit(matrix)
    candidates=[]
    if queries:
        qq=list(queries.values())
        distances,indices=model.kneighbors(vectorizer.transform([q['key'] for q in qq]))
        for q,ds,ids in zip(qq,distances,indices):
            qt=set(q['key'].split())
            for rank,(dist,ri) in enumerate(zip(ds,ids),1):
                r=refs[int(ri)];rt=set(r['key'].split());common=qt&rt
                score=round(max(0,1-float(dist))*100,2)
                if score<MIN_SCORE:continue
                risks=[]
                if (qt^rt)&{'SYARIAH','SHARIA'}:risks.append('SYARIAH_DIFFERENCE')
                if (qt^rt)&{'GROUP','GRUP'}:risks.append('GROUP_DIFFERENCE')
                if (qt^rt)&{'PALM','SAWIT'}:risks.append('BUSINESS_WORD_DIFFERENCE')
                if len(common-{'INDONESIA','GROUP','GRUP'})<=1:risks.append('WEAK_SHARED_TOKENS')
                if len(r['codes'])>1:risks.append('REFERENCE_MULTI_FAC_CODE')
                if q['matched_codes'] and r['codes'] and not (q['matched_codes']&r['codes']):risks.append('FAC_CODE_DIFFERENCE')
                if q['currencies'] and r['currencies'] and not (q['currencies']&r['currencies']):risks.append('CURRENCY_DIFFERENCE')
                candidates.append({'INSURED_SUSPEND':q['name'],'NORMALIZED_SUSPEND':q['key'],
                    'CLEAN_COLUMNS':', '.join(sorted(q['columns'])),'JUMLAH_BARIS':len(q['rows']),
                    'SKENARIO_ASAL':' | '.join(sorted(x for x in q['scenarios'] if x)),
                    'FLAG_ASAL':' | '.join(sorted(x for x in q['flags'] if x)),
                    'FAC_CODE_HASIL_MATCHING':', '.join(sorted(q['matched_codes'])),
                    'RANK':rank,'TFIDF_SCORE':score,'KANDIDAT_INSURED':r['name'],
                    'KANDIDAT_NORMALIZED':r['key'],'SUMBER':r['source'],
                    'SHARED_TOKENS':', '.join(sorted(common)),'RISK_FLAGS':' | '.join(risks),
                    'FAC_CODE_REFERENSI':', '.join(sorted(r['codes'])),
                    'JUMLAH_FAC_CODE_REFERENSI':len(r['codes']),
                    'CURRENCY_REFERENSI':', '.join(sorted(r['currencies'])),
                    'JUMLAH_REFERENSI':r['count'], 'KOLOM_REFERENSI':', '.join(sorted(r['columns'])),
                    'STATUS_VALIDASI':'PENDING','CATATAN_REVIEW':'','KEPUTUSAN_FAC_CODE':''})
    summary=[{'METRIK':'Perusahaan','NILAI':'KBRU'},
        {'METRIK':'Total baris matching','NILAI':len(result)},
        {'METRIK':'Baris terpilih untuk audit','NILAI':len(selected)},
        {'METRIK':'Nama insured unik diaudit','NILAI':len(queries)},
        *stats,{'METRIK':'Total entri referensi per sumber','NILAI':len(refs)},
        {'METRIK':'Jumlah kandidat TF-IDF','NILAI':len(candidates)},
        {'METRIK':'Top N per nama insured','NILAI':TOP_N},
        {'METRIK':'Enrichment suspend','NILAI':note},
        *[{'METRIK':f'File {s}','NILAI':p} for s,p in source_paths.items()]]
    guide=[{'PANDUAN':'Audit hanya SKENARIO Insured dan Unmatching/Unmatched (termasuk FLAG_PROD Unmatching).'},
        {'PANDUAN':'Referensi: OSBAL, FACUL, dan SLIPDB jika tersedia. KBRU tidak memakai Bordero.'},
        {'PANDUAN':'Referensi memakai semua INSURED_CLEAN_1-5 / clean insured 1-5, fallback raw.'},
        {'PANDUAN':'Query matching memakai INSURED_1/INSURED_2 dan tambahan 3-5 dari suspend jika join unik.'},
        {'PANDUAN':'Nama insured clean tidak dipecah ulang berdasarkan slash/comma/QQ; gunakan cleaning KBRU.'},
        {'PANDUAN':'Narasi suspense, treaty, placeholder dan data bukan nama insured diabaikan.'},
        {'PANDUAN':'TFIDF_SCORE bukan probabilitas dan bukan alasan otomatis untuk mengganti FAC CODE.'},
        {'PANDUAN':'FAC_CODE_DIFFERENCE, CURRENCY_DIFFERENCE dan flags lain adalah sinyal review, bukan keputusan final.'},
        {'PANDUAN':'Hasil audit tidak mengubah file matching/cleaning asal.'}]
    OUT.mkdir(parents=True,exist_ok=True)
    with pd.ExcelWriter(DEST,engine='openpyxl') as writer:
        for sheet,rows in [('AUDIT_SUMMARY',summary),('TFIDF_CANDIDATES',candidates),('RECEIPT_DETAILS',details),('CARA_PAKAI',guide)]:
            frame=pd.DataFrame(rows) if rows else pd.DataFrame({'INFO':['Tidak ada data yang memenuhi filter']})
            frame.to_excel(writer,sheet_name=sheet,index=False)
            ws=writer.sheets[sheet];ws.freeze_panes='A2';ws.auto_filter.ref=ws.dimensions
            for cell in ws[1]:
                cell.font=Font(color='FFFFFF',bold=True)
                cell.fill=PatternFill('solid',fgColor='1F4E78')
            for cells in ws.columns:
                ws.column_dimensions[cells[0].column_letter].width=min(52,max(14,max(len(str(c.value or '')) for c in list(cells)[:150])+2))
    print(f'OK: {DEST} | audited={len(selected)} | queries={len(queries)} | candidates={len(candidates)}')

if __name__=='__main__':
    main()
