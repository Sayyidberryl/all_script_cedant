import os, sys
sys.modules['numexpr'] = None
sys.modules['bottleneck'] = None
import pandas as pd

DATA_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "data")
df_sus = pd.read_excel(os.path.join(DATA_DIR, "kbru_output_suspend.xlsx"))
df_osbal = pd.read_excel(os.path.join(DATA_DIR, "kbru_output_osbal.xlsx"))

# Filter First Resources Group
sus_fr = df_sus[df_sus["INSURED_ORI"].astype(str).str.contains("FIRST RESOURCES", case=False, na=False)]
osb_fr = df_osbal[df_osbal["FAC_INSURED_ORI"].astype(str).str.contains("FIRST RESOURCES", case=False, na=False)]

print(f"Suspend rows for FIRST RESOURCES: {len(sus_fr)}")
print(f"OSBAL rows for FIRST RESOURCES: {len(osb_fr)}")

print("\n--- Suspend FIRST RESOURCES samples (Slip, Polis, Receipt Date, Amount) ---")
print(sus_fr[["RECEIPT NO", "SLIP_NO_ORI", "POLIS_ORI", "RECEIPT DATE", "AMOUNT ORI"]].head(10).to_string())

print("\n--- OSBAL FIRST RESOURCES samples (Slip, Polis, Fac Code, Com Date, Balance) ---")
print(osb_fr[["CCOS_REF_CODE", "FAC_SLIP_ORI", "FAC_POLICY_NO_ORI", "CLSDT_SLIP_NO", "CLSDT_POLICY_NO", "FAC_COM_DATE", "CCOS_BAL_DUE"]].head(10).to_string())
