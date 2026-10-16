import os, sys, re
sys.modules['numexpr'] = None
sys.modules['bottleneck'] = None
import pandas as pd

df_sus = pd.read_excel("data/kbru_output_suspend.xlsx")
df_osbal = pd.read_excel("data/kbru_output_osbal.xlsx")
df_facul = pd.read_excel("data/kbru_output_facul.xlsx")

def norm(v):
    if pd.isna(v): return ""
    return re.sub(r'[^A-Z0-9]', '', str(v).upper())

def base_slip(v):
    if pd.isna(v): return ""
    s = str(v).strip().upper()
    # If ends with -NN or -N
    s_base = re.sub(r'-\d+$', '', s)
    return norm(s_base)

sus_slips_base = set(df_sus['SLIP_CLEAN_1'].dropna().map(base_slip)) - {""}

osbal_slips = set()
for c in ['SLIP_CLEAN_1', 'SLIP_CLEAN_2', 'SLIP_CLEAN_3', 'SLIP_CLEAN_4', 'CLSDT_SLIP_NO', 'FAC_SLIP_ORI']:
    if c in df_osbal.columns:
        osbal_slips.update(set(df_osbal[c].dropna().map(norm)) - {""})

facul_slips = set()
for c in ['FAC_SLIP_CLEAN_1', 'FAC_SLIP_CLEAN_2', 'FAC_SLIP_CLEAN_3', 'FAC_SLIP_CLEAN_4', 'FAC_SLIP']:
    if c in df_facul.columns:
        facul_slips.update(set(df_facul[c].dropna().map(norm)) - {""})

print(f"Unique Suspend base slips: {len(sus_slips_base)}")
print(f"Base Slips in OSBAL: {len(sus_slips_base & osbal_slips)}")
print(f"Base Slips in FACUL: {len(sus_slips_base & facul_slips)}")

# Check some samples of Suspend slips matching OSBAL or FACUL
matched_osbal_sample = list(sus_slips_base & osbal_slips)[:5]
print("Matched in OSBAL sample:", matched_osbal_sample)
matched_facul_sample = list(sus_slips_base & facul_slips)[:5]
print("Matched in FACUL sample:", matched_facul_sample)
