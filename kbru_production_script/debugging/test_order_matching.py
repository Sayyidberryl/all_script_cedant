import os, sys, re
sys.modules['numexpr'] = None
sys.modules['bottleneck'] = None
import pandas as pd

df_sus = pd.read_excel("data/kbru_output_suspend.xlsx")
df_osbal = pd.read_excel("data/kbru_output_osbal.xlsx")
df_facul = pd.read_excel("data/kbru_output_facul.xlsx")

def norm(s):
    if pd.isna(s): return ""
    return re.sub(r'\s+', ' ', str(s).strip().upper())

def clean_alphanumeric(s):
    if pd.isna(s): return ""
    return re.sub(r'[^A-Z0-9]', '', str(s).upper())

# Index OSBAL
# OSBAL FAC_CODE
osbal_fac_map = {}
for idx, r in df_osbal.iterrows():
    fac = str(r.get('CCOS_REF_CODE', '')).strip()
    if fac:
        osbal_fac_map.setdefault(fac, []).append(idx)

# OSBAL Slip
osbal_slip_idx = {}
for idx, r in df_osbal.iterrows():
    for c in ['SLIP_CLEAN_1', 'SLIP_CLEAN_2', 'SLIP_CLEAN_3', 'SLIP_CLEAN_4', 'CLSDT_SLIP_NO', 'FAC_SLIP_ORI']:
        v = r.get(c)
        if pd.notna(v):
            v_str = str(v).strip()
            if v_str:
                osbal_slip_idx.setdefault(v_str, []).append(idx)
                c_num = clean_alphanumeric(v_str)
                if c_num:
                    osbal_slip_idx.setdefault(c_num, []).append(idx)

# OSBAL Polis
osbal_polis_idx = {}
for idx, r in df_osbal.iterrows():
    for c in ['POLIS_CLEAN_1', 'POLIS_CLEAN_2', 'CLSDT_POLICY_NO', 'FAC_POLICY_NO_ORI']:
        v = r.get(c)
        if pd.notna(v):
            v_str = str(v).strip()
            if v_str:
                osbal_polis_idx.setdefault(v_str, []).append(idx)
                c_num = clean_alphanumeric(v_str)
                if c_num:
                    osbal_polis_idx.setdefault(c_num, []).append(idx)

# OSBAL Insured
osbal_ins_idx = {}
for idx, r in df_osbal.iterrows():
    for c in ['INSURED_CLEAN_1', 'INSURED_CLEAN_2', 'INSURED_CLEAN_3', 'INSURED_CLEAN_4', 'INSURED_CLEAN_5', 'FAC_INSURED_ORI']:
        v = r.get(c)
        if pd.notna(v):
            v_str = norm(v)
            if v_str:
                osbal_ins_idx.setdefault(v_str, []).append(idx)

# FACUL FAC_CODE map
facul_fac_col = 'FAC_CODE'
facul_slip_idx = {}
facul_polis_idx = {}
facul_ins_idx = {}
for idx, r in df_facul.iterrows():
    fac = str(r.get(facul_fac_col, '')).strip()
    if not fac or fac not in osbal_fac_map:
        continue
    for c in ['FAC_SLIP_CLEAN_1', 'FAC_SLIP_CLEAN_2', 'FAC_SLIP_CLEAN_3', 'FAC_SLIP_CLEAN_4', 'FAC_SLIP']:
        v = r.get(c)
        if pd.notna(v):
            v_str = str(v).strip()
            if v_str:
                facul_slip_idx.setdefault(v_str, []).append(fac)
                c_num = clean_alphanumeric(v_str)
                if c_num:
                    facul_slip_idx.setdefault(c_num, []).append(fac)
    for c in ['POLIS_CLEAN_1', 'POLIS_CLEAN_2', 'POLIS_CLEAN_3', 'FAC_POLICY_NO']:
        v = r.get(c)
        if pd.notna(v):
            v_str = str(v).strip()
            if v_str:
                facul_polis_idx.setdefault(v_str, []).append(fac)
                c_num = clean_alphanumeric(v_str)
                if c_num:
                    facul_polis_idx.setdefault(c_num, []).append(fac)
    for c in ['FAC_INSURED_1', 'FAC_INSURED_2', 'FAC_INSURED_3', 'FAC_INSURED_4', 'FAC_INSURED_5', 'FAC_INSURED_ORI']:
        v = r.get(c)
        if pd.notna(v):
            v_str = norm(v)
            if v_str:
                facul_ins_idx.setdefault(v_str, []).append(fac)

matched_slip = 0
matched_polis = 0
matched_ins = 0
unmatched = 0

for idx, r in df_sus.iterrows():
    # 1. Slip
    sus_slips = []
    for c in ['SLIP_CLEAN_1', 'SLIP_NO_ORI']:
        v = r.get(c)
        if pd.notna(v) and str(v).strip():
            v_str = str(v).strip()
            sus_slips.append(v_str)
            sus_slips.append(clean_alphanumeric(v_str))
    
    hits = set()
    for s in sus_slips:
        if s in osbal_slip_idx:
            hits.update(osbal_slip_idx[s])
        elif s in facul_slip_idx:
            for fac in facul_slip_idx[s]:
                hits.update(osbal_fac_map.get(fac, []))
    if hits:
        matched_slip += 1
        continue
    
    # 2. Polis
    sus_polis = []
    for c in ['POLIS_CLEAN_1', 'POLIS_ORI']:
        v = r.get(c)
        if pd.notna(v) and str(v).strip():
            v_str = str(v).strip()
            sus_polis.append(v_str)
            sus_polis.append(clean_alphanumeric(v_str))
    for p in sus_polis:
        if p in osbal_polis_idx:
            hits.update(osbal_polis_idx[p])
        elif p in facul_polis_idx:
            for fac in facul_polis_idx[p]:
                hits.update(osbal_fac_map.get(fac, []))
    if hits:
        matched_polis += 1
        continue
    
    # 3. Insured
    sus_ins = []
    for c in ['INSURED_1', 'INSURED_ORI']:
        v = r.get(c)
        if pd.notna(v) and str(v).strip():
            sus_ins.append(norm(v))
    for ins in sus_ins:
        if ins in osbal_ins_idx:
            hits.update(osbal_ins_idx[ins])
        elif ins in facul_ins_idx:
            for fac in facul_ins_idx[ins]:
                hits.update(osbal_fac_map.get(fac, []))
    if hits:
        matched_ins += 1
        continue
    
    unmatched += 1

print("--- INITIAL MATCHING STATS ---")
print(f"Matched by Slip:    {matched_slip}")
print(f"Matched by Polis:   {matched_polis}")
print(f"Matched by Insured: {matched_ins}")
print(f"Unmatched:          {unmatched}")
print(f"Total matched:      {matched_slip + matched_polis + matched_ins} / {len(df_sus)}")
