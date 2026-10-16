import os
import sys
sys.modules['numexpr'] = None
sys.modules['bottleneck'] = None

import pandas as pd
import re

DATA_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "data")
df_v1 = pd.read_excel(os.path.join(DATA_DIR, "final_output_v1.xlsx"))
df_osbal = pd.read_excel(os.path.join(DATA_DIR, "kbru_output_osbal.xlsx"))
df_facul = pd.read_excel(os.path.join(DATA_DIR, "kbru_output_facul.xlsx"))

df_ins_gt1 = df_v1[df_v1["SKENARIO"] == "Insured only (> 1 Fac)"]

print("Contoh 15 Slip & Polis di 'Insured only (> 1 Fac)':")
for idx, r in df_ins_gt1.head(15).iterrows():
    print(f"Receipt: {r['RECEIPT NO']} | Slip: '{r['SLIP_NO']}' | Polis: '{r['POLIS']}' | Insured: '{r['INSURED_ORI']}'")

# Cek apakah ada slip atau polis yang mirip di OSBAL
print("\nCari kemiripan slip di OSBAL:")
sample_slips = df_ins_gt1["SLIP_NO"].dropna().unique()[:5]
for s in sample_slips:
    clean_s = re.sub(r'[^A-Z0-9]', '', str(s).upper())
    base_s = re.sub(r'-\d+$', '', str(s).strip().upper())
    clean_base = re.sub(r'[^A-Z0-9]', '', base_s)
    
    # Cari di OSBAL
    found_osbal = False
    for col in [c for c in df_osbal.columns if 'slip' in c.lower()]:
        matches = df_osbal[df_osbal[col].astype(str).str.contains(clean_base[:10], na=False, regex=False)]
        if len(matches) > 0:
            print(f"  Slip '{s}' mirip di OSBAL col {col}: {matches[col].head(2).tolist()}")
            found_osbal = True
    if not found_osbal:
        print(f"  Slip '{s}' (base: {base_s}) TIDAK ADA di OSBAL sama sekali.")

print("\nCari kemiripan Polis di OSBAL:")
sample_polis = df_ins_gt1["POLIS"].dropna().unique()[:5]
for p in sample_polis:
    clean_p = re.sub(r'[^A-Z0-9]', '', str(p).upper())
    # ambil potongan tengah misal 00004903 atau 054.0000002
    parts = str(p).split('/')
    core = parts[0] if len(parts) > 0 else p
    found_osbal = False
    for col in [c for c in df_osbal.columns if 'pol' in c.lower()]:
        matches = df_osbal[df_osbal[col].astype(str).str.contains(core, na=False, regex=False)]
        if len(matches) > 0:
            print(f"  Polis '{p}' core '{core}' mirip di OSBAL col {col}: {matches[col].head(2).tolist()}")
            found_osbal = True
    if not found_osbal:
        print(f"  Polis '{p}' core '{core}' TIDAK ADA di OSBAL sama sekali.")
