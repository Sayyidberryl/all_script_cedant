import os, sys, re
sys.modules['numexpr'] = None
sys.modules['bottleneck'] = None
import pandas as pd

df_sus = pd.read_excel("data/kbru_output_suspend.xlsx")
df_osbal = pd.read_excel("data/kbru_output_osbal.xlsx")
df_facul = pd.read_excel("data/kbru_output_facul.xlsx")

def norm(s):
    if pd.isna(s): return ""
    return re.sub(r'\s+', '', str(s).strip().upper())

def get_slip_variants(s):
    if pd.isna(s): return []
    res = []
    val = norm(s)
    if val:
        res.append(val)
        # remove hyphens/dots
        clean_num = re.sub(r'[^A-Z0-9]', '', val)
        if clean_num and clean_num != val:
            res.append(clean_num)
        # remove suffix -NN or -N
        base = re.sub(r'-\d+$', '', val)
        if base and base != val:
            res.append(base)
            base_clean = re.sub(r'[^A-Z0-9]', '', base)
            if base_clean:
                res.append(base_clean)
    return list(set(res))

osbal_slip_index = {}
osbal_slip_cols = ['SLIP_CLEAN_1', 'SLIP_CLEAN_2', 'SLIP_CLEAN_3', 'SLIP_CLEAN_4', 'CLSDT_SLIP_NO', 'FAC_SLIP_ORI']
for idx, r in df_osbal.iterrows():
    for c in osbal_slip_cols:
        for v in get_slip_variants(r.get(c)):
            osbal_slip_index.setdefault(v, set()).add(idx)

facul_slip_index = {}
facul_slip_cols = ['FAC_SLIP_CLEAN_1', 'FAC_SLIP_CLEAN_2', 'FAC_SLIP_CLEAN_3', 'FAC_SLIP_CLEAN_4', 'FAC_SLIP']
for idx, r in df_facul.iterrows():
    for c in facul_slip_cols:
        for v in get_slip_variants(r.get(c)):
            facul_slip_index.setdefault(v, set()).add(idx)

sus_matched_osbal = 0
sus_matched_facul = 0

for idx, r in df_sus.iterrows():
    sus_slips = get_slip_variants(r.get('SLIP_CLEAN_1')) + get_slip_variants(r.get('SLIP_NO_ORI'))
    hit_osbal = False
    for s in sus_slips:
        if s in osbal_slip_index:
            hit_osbal = True
            break
    if hit_osbal:
        sus_matched_osbal += 1
    else:
        hit_facul = False
        for s in sus_slips:
            if s in facul_slip_index:
                hit_facul = True
                break
        if hit_facul:
            sus_matched_facul += 1

print(f"Total Suspend rows: {len(df_sus)}")
print(f"Suspend matched via slip to OSBAL: {sus_matched_osbal}")
print(f"Suspend matched via slip to FACUL: {sus_matched_facul}")
print(f"Total slip matches: {sus_matched_osbal + sus_matched_facul} / {len(df_sus)} ({((sus_matched_osbal + sus_matched_facul)/len(df_sus))*100:.1f}%)")
