import os, sys
sys.modules['numexpr'] = None
sys.modules['bottleneck'] = None
import pandas as pd

df_sus = pd.read_excel("data/kbru_output_suspend.xlsx")
df_osbal = pd.read_excel("data/kbru_output_osbal.xlsx")
df_facul = pd.read_excel("data/kbru_output_facul.xlsx")

print("--- SUSPEND SLIP SAMPLES ---")
print(df_sus['SLIP_NO_ORI'].dropna().head(10).tolist())
print(df_sus['SLIP_CLEAN_1'].dropna().head(10).tolist())

print("\n--- OSBAL SLIP SAMPLES ---")
print(df_osbal['FAC_SLIP_ORI'].dropna().head(10).tolist())
print(df_osbal['SLIP_CLEAN_1'].dropna().head(10).tolist())

print("\n--- FACUL SLIP SAMPLES ---")
print(df_facul['FAC_SLIP'].dropna().head(10).tolist())
print(df_facul['FAC_SLIP_CLEAN_1'].dropna().head(10).tolist())
