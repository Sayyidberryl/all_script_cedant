import os
import sys
sys.modules['numexpr'] = None
sys.modules['bottleneck'] = None

import pandas as pd
import re

DATA_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "data")
df_v1 = pd.read_excel(os.path.join(DATA_DIR, "final_output_v1.xlsx"))
df_facul = pd.read_excel(os.path.join(DATA_DIR, "kbru_output_facul.xlsx"))
df_osbal = pd.read_excel(os.path.join(DATA_DIR, "kbru_output_osbal.xlsx"))

df_ins_gt1 = df_v1[df_v1["SKENARIO"] == "Insured only (> 1 Fac)"]

sample_slips = df_ins_gt1["SLIP_NO"].dropna().unique()[:10]
print("Cek sample slips di FACUL:")
for s in sample_slips:
    clean_s = re.sub(r'[^A-Z0-9]', '', str(s).upper())
    base_s = re.sub(r'-\d+$', '', str(s).strip().upper())
    clean_base = re.sub(r'[^A-Z0-9]', '', base_s)
    
    found = False
    for col in [c for c in df_facul.columns if 'slip' in c.lower()]:
        matches = df_facul[df_facul[col].astype(str).str.contains(clean_base, na=False, regex=False)]
        if len(matches) > 0:
            print(f"  Slip '{s}' ada di FACUL col {col}: {matches['FAC_CODE'].head(2).tolist()}")
            # Apakah FAC_CODE ini ada di OSBAL?
            facs = matches['FAC_CODE'].dropna().unique()
            osb_has = df_osbal[df_osbal['CCOS_REF_CODE'].isin(facs)]
            print(f"    -> FAC_CODE di OSBAL? {len(osb_has)} rows")
            found = True
    if not found:
        print(f"  Slip '{s}' TIDAK ADA di FACUL.")

sample_polis = df_ins_gt1["POLIS"].dropna().unique()[:10]
print("\nCek sample polis di FACUL:")
for p in sample_polis:
    clean_p = re.sub(r'[^A-Z0-9]', '', str(p).upper())
    parts = str(p).split('/')
    core = parts[0] if len(parts) > 0 else p
    found = False
    for col in [c for c in df_facul.columns if 'pol' in c.lower()]:
        matches = df_facul[df_facul[col].astype(str).str.contains(core, na=False, regex=False)]
        if len(matches) > 0:
            print(f"  Polis '{p}' ada di FACUL col {col}: {matches['FAC_CODE'].head(2).tolist()}")
            facs = matches['FAC_CODE'].dropna().unique()
            osb_has = df_osbal[df_osbal['CCOS_REF_CODE'].isin(facs)]
            print(f"    -> FAC_CODE di OSBAL? {len(osb_has)} rows")
            found = True
    if not found:
        print(f"  Polis '{p}' TIDAK ADA di FACUL.")
