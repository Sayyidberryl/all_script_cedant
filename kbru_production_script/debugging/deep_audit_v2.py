import os
import sys

sys.modules['numexpr'] = None
sys.modules['bottleneck'] = None

import pandas as pd
import numpy as np
import re

DATA_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "data")
v1_file = os.path.join(DATA_DIR, "final_output_v1.xlsx")
sus_file = os.path.join(DATA_DIR, "kbru_output_suspend.xlsx")
osbal_file = os.path.join(DATA_DIR, "kbru_output_osbal.xlsx")

df_v1 = pd.read_excel(v1_file)
df_sus = pd.read_excel(sus_file)
df_osbal = pd.read_excel(osbal_file)

print("="*70)
print("AUDIT 6: MEMBEDAH 100 BARIS 'Slip + Insured' & 13 BARIS 'Slip only'")
print("="*70)
slip_valid_matches = df_v1[df_v1["SKENARIO"].isin(["Slip + Insured", "Slip only"])]
print(f"Total baris: {len(slip_valid_matches)}")
print("Berapa baris yang SLIP_NO == '-'? :", (slip_valid_matches["SLIP_NO"].astype(str).str.strip() == "-").sum())
non_dash_slips = slip_valid_matches[slip_valid_matches["SLIP_NO"].astype(str).str.strip() != "-"]
print(f"Baris dengan SLIP_NO valid: {len(non_dash_slips)}")
print("\nSample 10 baris non-dash Slip match:")
cols_to_show = ["RECEIPT NO", "SLIP_NO", "POLIS", "INSURED_ORI", "CCOS_REF_CODE", "AMOUNT ORI", "CCOS_BAL_DUE", "FLAG_PROD", "SKENARIO"]
print(non_dash_slips[cols_to_show].head(10).to_string())

# Cek kesesuaian slip di OSBAL untuk baris-baris ini
print("\nCek apakah SLIP mereka benar-benar cocok dengan baris OSBAL:")
osbal_by_fac = {}
for idx, r in df_osbal.iterrows():
    fac = str(r.get("CCOS_REF_CODE", "")).strip()
    if fac:
        osbal_by_fac.setdefault(fac, []).append(r)

match_accuracy = []
for idx, r in non_dash_slips.head(20).iterrows():
    s_slip = str(r["SLIP_NO"]).strip()
    s_fac = str(r["CCOS_REF_CODE"]).strip()
    osb_rows = osbal_by_fac.get(s_fac, [])
    found_in_osb = False
    for o in osb_rows:
        for c in ["SLIP_CLEAN_1", "CLSDT_SLIP_NO", "FAC_SLIP_ORI"]:
            if s_slip in str(o.get(c, "")):
                found_in_osb = True
                break
    match_accuracy.append((s_slip, s_fac, r["INSURED_ORI"], found_in_osb))

print("Verifikasi 20 sampel apakah SLIP suspend benar ada di OSBAL row terkait:")
for s_slip, s_fac, ins, ok in match_accuracy[:10]:
    print(f"  Slip: {s_slip} | Fac: {s_fac} | Insured: {ins} -> Cocok di OSBAL: {ok}")

print("\n" + "="*70)
print("AUDIT 7: MEMBEDAH 185 BARIS 'New Entry'")
print("="*70)
new_entries = df_v1[df_v1["FLAG_PROD"] == "New Entry"]
print(f"Total New Entry: {len(new_entries)}")
print("Distribusi SKENARIO pada New Entry:")
print(new_entries["SKENARIO"].value_counts())
print("Distribusi CCOS_BAL_DUE pada New Entry:")
print(new_entries["CCOS_BAL_DUE"].value_counts().head(5))
print("\nSample 10 baris New Entry:")
print(new_entries[cols_to_show].head(10).to_string())

print("\n" + "="*70)
print("AUDIT 8: OVERLAP POLIS SUSPEND DENGAN OSBAL (SUBSTRING / CLEAN MATCH)")
print("="*70)
# Mari kita lihat apakah nomor polis suspend memiliki format spesifik yang bisa dimatch ke OSBAL
# Contoh suspend: 00004903/FEQI/054.0000002-SRB/08/2023
# Bagaimana nomor polis di CLSDT_POLICY_NO OSBAL?
print("Contoh 10 baris CLSDT_POLICY_NO di OSBAL:")
print(df_osbal["CLSDT_POLICY_NO"].dropna().head(10).tolist())
# Coba ekstrak pola nomor depan (8 digit pertama) dari suspend dan OSBAL
sus_pols = df_sus["POLIS_ORI"].dropna().astype(str).tolist()
osb_pols = df_osbal["CLSDT_POLICY_NO"].dropna().astype(str).tolist()

sus_heads = {p.split('/')[0].strip() for p in sus_pols if '/' in p}
osb_heads = {p.split('/')[0].strip() for p in osb_pols if '/' in p}
common_heads = sus_heads & osb_heads
print(f"Unique polis head (sebelum garis miring pertama) di Suspend: {len(sus_heads)}")
print(f"Unique polis head di OSBAL CLSDT_POLICY_NO: {len(osb_heads)}")
print(f"Polis head yang sama persis antara Suspend dan OSBAL: {len(common_heads)}")
if common_heads:
    print("Contoh polis head yang sama:", list(common_heads)[:5])
