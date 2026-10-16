import os, sys, re
sys.modules['numexpr'] = None
sys.modules['bottleneck'] = None
import pandas as pd

DATA_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "data")
df_sus = pd.read_excel(os.path.join(DATA_DIR, "kbru_output_suspend.xlsx"))
df_osbal = pd.read_excel(os.path.join(DATA_DIR, "kbru_output_osbal.xlsx"))

# Cari polis 00006702 di suspend dan di osbal
sus_sample = df_sus[df_sus["POLIS_ORI"].astype(str).str.contains("00006702", na=False)]
osb_sample = df_osbal[df_osbal.astype(str).apply(lambda row: row.str.contains("00006702", na=False)).any(axis=1)]

print(f"Suspend rows with '00006702': {len(sus_sample)}")
print(sus_sample[["RECEIPT NO", "SLIP_NO_ORI", "POLIS_ORI", "INSURED_ORI"]].head(3).to_string())

print(f"\nOSBAL rows with '00006702': {len(osb_sample)}")
for c in ["CCOS_REF_CODE", "FAC_SLIP_ORI", "FAC_POLICY_NO_ORI", "CLSDT_SLIP_NO", "CLSDT_POLICY_NO", "FAC_INSURED_ORI"]:
    print(f"  {c}: {osb_sample[c].tolist()[:3]}")
