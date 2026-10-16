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

# Slips
sus_slips = set(df_sus['SLIP_CLEAN_1'].dropna().map(norm)) - {""}
osbal_slips = set()
for c in ['SLIP_CLEAN_1', 'SLIP_CLEAN_2', 'SLIP_CLEAN_3', 'SLIP_CLEAN_4', 'CLSDT_SLIP_NO', 'FAC_SLIP_ORI']:
    if c in df_osbal.columns:
        osbal_slips.update(set(df_osbal[c].dropna().map(norm)) - {""})

facul_slips = set()
for c in ['FAC_SLIP_CLEAN_1', 'FAC_SLIP_CLEAN_2', 'FAC_SLIP_CLEAN_3', 'FAC_SLIP_CLEAN_4', 'FAC_SLIP']:
    if c in df_facul.columns:
        facul_slips.update(set(df_facul[c].dropna().map(norm)) - {""})

# Polis
sus_polis = set(df_sus['POLIS_CLEAN_1'].dropna().map(norm)) - {""}
osbal_polis = set()
for c in ['POLIS_CLEAN_1', 'POLIS_CLEAN_2', 'CLSDT_POLICY_NO', 'FAC_POLICY_NO_ORI']:
    if c in df_osbal.columns:
        osbal_polis.update(set(df_osbal[c].dropna().map(norm)) - {""})

facul_polis = set()
for c in ['POLIS_CLEAN_1', 'POLIS_CLEAN_2', 'POLIS_CLEAN_3', 'FAC_POLICY_NO']:
    if c in df_facul.columns:
        facul_polis.update(set(df_facul[c].dropna().map(norm)) - {""})

# Insured
sus_ins = set(df_sus['INSURED_1'].dropna().map(lambda x: str(x).strip().upper())) - {""}
osbal_ins = set()
for c in ['INSURED_CLEAN_1', 'INSURED_CLEAN_2', 'INSURED_CLEAN_3', 'INSURED_CLEAN_4', 'INSURED_CLEAN_5']:
    if c in df_osbal.columns:
        osbal_ins.update(set(df_osbal[c].dropna().map(lambda x: str(x).strip().upper())) - {""})

print(f"Total Suspend rows: {len(df_sus)}")
print(f"Unique Suspend slips: {len(sus_slips)}")
print(f"Slips in OSBAL: {len(sus_slips & osbal_slips)} (in FACUL: {len(sus_slips & facul_slips)})")
print(f"Unique Suspend polis: {len(sus_polis)}")
print(f"Polis in OSBAL: {len(sus_polis & osbal_polis)} (in FACUL: {len(sus_polis & facul_polis)})")
print(f"Unique Suspend insured: {len(sus_ins)}")
print(f"Insured in OSBAL: {len(sus_ins & osbal_ins)}")
