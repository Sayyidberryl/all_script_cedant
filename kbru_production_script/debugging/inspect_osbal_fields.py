import os, sys
sys.modules['numexpr'] = None
sys.modules['bottleneck'] = None
import pandas as pd

DATA_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "data")
df_osbal = pd.read_excel(os.path.join(DATA_DIR, "kbru_output_osbal.xlsx"))

print("Sample 15 baris OSBAL:")
cols = ["CCOS_REF_CODE", "FAC_SLIP_ORI", "FAC_POLICY_NO_ORI", "CLSDT_SLIP_NO", "CLSDT_POLICY_NO", "FAC_INSURED_ORI"]
print(df_osbal[cols].head(15).to_string())

print("\nDistribusi format di FAC_SLIP_ORI vs FAC_POLICY_NO_ORI:")
print("Berapa banyak FAC_SLIP_ORI yang berformat '0000...' (nomor polis KBRU)?")
pol_like_in_slip = df_osbal["FAC_SLIP_ORI"].astype(str).str.contains(r'FPAR|FEQI|FIAR|SPLI|CMHL|SRB|KBRU', regex=True).sum()
print("FAC_SLIP_ORI mirip polis format:", pol_like_in_slip, "dari", len(df_osbal))

slip_like_in_slip = df_osbal["FAC_SLIP_ORI"].astype(str).str.contains(r'^110', regex=True).sum()
print("FAC_SLIP_ORI berawalan '110' (format slip KBRU):", slip_like_in_slip)

clsdt_slip_110 = df_osbal["CLSDT_SLIP_NO"].astype(str).str.contains(r'^110', regex=True).sum()
print("CLSDT_SLIP_NO berawalan '110':", clsdt_slip_110)

clsdt_pol_like = df_osbal["CLSDT_POLICY_NO"].astype(str).str.contains(r'FPAR|FEQI|FIAR|SPLI|CMHL|SRB|KBRU', regex=True).sum()
print("CLSDT_POLICY_NO mirip polis format:", clsdt_pol_like)
