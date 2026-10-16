import os
import sys
sys.modules['numexpr'] = None
sys.modules['bottleneck'] = None

import pandas as pd
import numpy as np
import re

DATA_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "data")
v1_path = os.path.join(DATA_DIR, "final_output_v1.xlsx")
sus_path = os.path.join(DATA_DIR, "kbru_output_suspend.xlsx")
osbal_path = os.path.join(DATA_DIR, "kbru_output_osbal.xlsx")
facul_path = os.path.join(DATA_DIR, "kbru_output_facul.xlsx")

df_v1 = pd.read_excel(v1_path)
df_sus = pd.read_excel(sus_path)
df_osbal = pd.read_excel(osbal_path)
df_facul = pd.read_excel(facul_path)

print("="*60)
print("1. SUMMARY OF FINAL_OUTPUT_V1")
print("="*60)
print("Total rows:", len(df_v1))
print("\nFLAG_PROD value counts:")
print(df_v1["FLAG_PROD"].value_counts())
print("\nSKENARIO value counts:")
print(df_v1["SKENARIO"].value_counts())

print("\n" + "="*60)
print("2. INVESTIGATING 'Insured only (> 1 Fac)' (927 rows)")
print("="*60)
df_ins_gt1 = df_v1[df_v1["SKENARIO"] == "Insured only (> 1 Fac)"]

has_slip = df_ins_gt1["SLIP_NO"].replace("", np.nan).dropna()
has_polis = df_ins_gt1["POLIS"].replace("", np.nan).dropna()
print(f"Dari 927 baris 'Insured only (> 1 Fac)':")
print(f"  - Memiliki SLIP_NO terisi: {len(has_slip)} ({len(has_slip)/len(df_ins_gt1)*100:.1f}%)")
print(f"  - Memiliki POLIS terisi  : {len(has_polis)} ({len(has_polis)/len(df_ins_gt1)*100:.1f}%)")
print(f"  - Keduanya kosong        : {len(df_ins_gt1[df_ins_gt1['SLIP_NO'].fillna('').eq('') & df_ins_gt1['POLIS'].fillna('').eq('')])}")

print("\nSample 10 baris 'Insured only (> 1 Fac)':")
sample_cols = ["RECEIPT NO", "SLIP_NO", "POLIS", "INSURED_ORI", "CCOS_REF_CODE"]
print(df_ins_gt1[sample_cols].head(10).to_string())

# Cek berapa banyak fac codes yang ter-match per baris di Insured only (> 1 Fac)
fac_counts = df_ins_gt1["CCOS_REF_CODE"].apply(lambda x: len(str(x).split(",")) if pd.notna(x) else 0)
print(f"\nDistribusi jumlah FAC code per baris di Insured only (> 1 Fac):")
print(fac_counts.describe())

# Mengapa slip/polis mereka tidak match di OSBAL/FACUL?
# Ambil yang punya slip
slips_in_ins_gt1 = df_ins_gt1[df_ins_gt1["SLIP_NO"].fillna("").ne("")]["SLIP_NO"].unique()
print(f"\nAda {len(slips_in_ins_gt1)} slip unik di baris Insured only (> 1 Fac).")
# Cek apakah slip ini ada di OSBAL atau FACUL
osbal_slips = set(df_osbal["CLSDT_SLIP_NO"].dropna().astype(str).str.strip().str.upper()) if "CLSDT_SLIP_NO" in df_osbal.columns else set()
for c in [col for col in df_osbal.columns if "slip" in col.lower()]:
    osbal_slips.update(df_osbal[c].dropna().astype(str).str.strip().str.upper())

facul_slips = set(df_facul["FAC_SLIP"].dropna().astype(str).str.strip().str.upper()) if "FAC_SLIP" in df_facul.columns else set()
for c in [col for col in df_facul.columns if "slip" in col.lower()]:
    facul_slips.update(df_facul[c].dropna().astype(str).str.strip().str.upper())

found_in_osbal = [s for s in slips_in_ins_gt1 if str(s).strip().upper() in osbal_slips]
found_in_facul = [s for s in slips_in_ins_gt1 if str(s).strip().upper() in facul_slips]
print(f"  - Berapa slip ditemukan di OSBAL? {len(found_in_osbal)}")
print(f"  - Berapa slip ditemukan di FACUL? {len(found_in_facul)}")
if found_in_osbal:
    print("    Contoh slip yg ada di OSBAL tapi malah match Insured only:", found_in_osbal[:5])
if found_in_facul:
    print("    Contoh slip yg ada di FACUL tapi malah match Insured only:", found_in_facul[:5])

print("\n" + "="*60)
print("3. INVESTIGATING 'Slip only (> 1 Fac)' (73 rows)")
print("="*60)
df_slip_gt1 = df_v1[df_v1["SKENARIO"] == "Slip only (> 1 Fac)"]
print(f"Total rows: {len(df_slip_gt1)}")
print("\nSample 10 baris 'Slip only (> 1 Fac)':")
print(df_slip_gt1[["RECEIPT NO", "SLIP_NO", "INSURED_ORI", "CCOS_REF_CODE"]].head(10).to_string())

print("\n" + "="*60)
print("4. INVESTIGATING 'Unmatching' (329 rows)")
print("="*60)
df_unm = df_v1[df_v1["FLAG_PROD"] == "Unmatching"]
has_slip_u = df_unm["SLIP_NO"].replace("", np.nan).dropna()
has_pol_u = df_unm["POLIS"].replace("", np.nan).dropna()
has_ins_u = df_unm["INSURED_ORI"].replace("", np.nan).dropna()
print(f"Dari 329 baris Unmatching:")
print(f"  - Ada SLIP_NO    : {len(has_slip_u)}")
print(f"  - Ada POLIS      : {len(has_pol_u)}")
print(f"  - Ada INSURED_ORI: {len(has_ins_u)}")
print("\nSample 10 baris Unmatching:")
print(df_unm[["RECEIPT NO", "SLIP_NO", "POLIS", "INSURED_ORI", "CURR ORI", "AMOUNT ORI"]].head(10).to_string())

print("\n" + "="*60)
print("5. SLIPDB STATUS")
print("="*60)
slipdb_p = os.path.join(DATA_DIR, "kbru_output_slipdb.xlsx")
print(f"File {slipdb_p} exists? {os.path.exists(slipdb_p)}")
if os.path.exists(slipdb_p):
    try:
        df_sldb = pd.read_excel(slipdb_p)
        print(f"SLIPDB rows: {len(df_sldb)}, cols: {len(df_sldb.columns)}")
    except Exception as e:
        print(f"Error read slipdb: {e}")
