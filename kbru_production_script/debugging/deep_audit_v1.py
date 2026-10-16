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
facul_file = os.path.join(DATA_DIR, "kbru_output_facul.xlsx")

df_v1 = pd.read_excel(v1_file)
df_sus = pd.read_excel(sus_file)
df_osbal = pd.read_excel(osbal_file)
df_facul = pd.read_excel(facul_file)

print("="*70)
print("AUDIT 1: CHECKING SLIP = '-' / JUNK VALUES IN MATCHING")
print("="*70)
# Cek berapa baris di final_output_v1 yang SLIP_NO == '-'
dash_slips = df_v1[df_v1["SLIP_NO"].astype(str).str.strip().isin(["-", "--", "---", "", "NAN", "NONE"])]
print(f"Total baris dengan SLIP_NO dummy/dash di final_output_v1: {len(dash_slips)}")
print("Skenario dari baris SLIP_NO dummy/dash:")
print(dash_slips["SKENARIO"].value_counts())
print("\nSample baris dummy SLIP yang matched sebagai Slip only:")
sample_dash = dash_slips[dash_slips["SKENARIO"].str.contains("Slip", na=False)]
print(sample_dash[["RECEIPT NO", "SLIP_NO", "SLIP_NO_CLN", "POLIS", "INSURED_ORI", "CCOS_REF_CODE", "SKENARIO"]].head(10).to_string())

# Cek apakah '-' ada di OSBAL / FACUL clean slip
for name, df in [("OSBAL", df_osbal), ("FACUL", df_facul)]:
    slip_cols = [c for c in df.columns if "slip" in c.lower()]
    for c in slip_cols:
        cnt = df[c].astype(str).str.strip().isin(["-", "--", "NAN", "NONE"]).sum()
        if cnt > 0:
            print(f"[{name}] Kolom {c} memiliki {cnt} baris dengan nilai dummy/dash!")

print("\n" + "="*70)
print("AUDIT 2: MENGAPA HANYA 1 BARIS MATCH VIA POLIS?")
print("="*70)
# Cek ketersediaan polis di Suspend
sus_has_polis = df_sus[df_sus["POLIS_ORI"].dropna().astype(str).str.strip().ne("") & df_sus["POLIS_ORI"].dropna().astype(str).str.strip().ne("-")]
print(f"Jumlah baris Suspend dengan POLIS_ORI terisi: {len(sus_has_polis)} / {len(df_sus)}")
print("Contoh 5 nilai POLIS_ORI di Suspend:")
print(sus_has_polis["POLIS_ORI"].head(5).tolist())
print("Contoh 5 nilai POLIS_CLEAN_1 di Suspend:")
print(sus_has_polis["POLIS_CLEAN_1"].head(5).tolist())

# Cek bagaimana polis disimpan di OSBAL dan FACUL
print("\nKolom Polis di OSBAL:")
osb_pol_cols = [c for c in df_osbal.columns if "pol" in c.lower()]
print(osb_pol_cols)
for c in osb_pol_cols:
    sample_vals = df_osbal[c].dropna().head(3).tolist()
    print(f"  {c} (non-null: {df_osbal[c].notna().sum()}): {sample_vals}")

print("\nKolom Polis di FACUL:")
fac_pol_cols = [c for c in df_facul.columns if "pol" in c.lower()]
print(fac_pol_cols)
for c in fac_pol_cols:
    sample_vals = df_facul[c].dropna().head(3).tolist()
    print(f"  {c} (non-null: {df_facul[c].notna().sum()}): {sample_vals}")

# Coba cari apakah ada pola nomor polis suspend di OSBAL/FACUL
# Format suspend: 00004903/FEQI/054.0000002-SRB/08/2023
# Bersih: 00004903FEQI0540000002SRB082023
# Apakah di OSBAL polis disimpan tanpa broker (-SRB atau -KBRURE)?
print("\nEksperimen substring polis suspend di OSBAL:")
regex_sample = r"00004903"
for c in osb_pol_cols:
    hits = df_osbal[df_osbal[c].astype(str).str.contains(regex_sample, na=False)]
    if len(hits) > 0:
        print(f"  Ketemu {len(hits)} baris di OSBAL col {c} untuk '00004903':")
        print(hits[["CCOS_REF_CODE", c]].head(3).to_string())

print("\n" + "="*70)
print("AUDIT 3: ANALISIS KELAYAKAN 'Insured only' (1,317 baris)")
print("="*70)
df_ins_all = df_v1[df_v1["SKENARIO"].str.startswith("Insured only")]
print(f"Total baris match via Insured only: {len(df_ins_all)}")
print("Breakdown per SKENARIO:")
print(df_ins_all["SKENARIO"].value_counts())
print("\nBreakdown per FLAG_PROD:")
print(df_ins_all["FLAG_PROD"].value_counts())

# Untuk yang Insured only dengan single fac (258 baris), apakah amount-nya cocok?
df_ins_single = df_v1[df_v1["SKENARIO"] == "Insured only"]
exact_amt = (df_ins_single["AMOUNT_ORI_MIN1"].round(2) == df_ins_single["CCOS_BAL_DUE"].round(2)).sum()
print(f"\nDari {len(df_ins_single)} baris 'Insured only' (single fac):")
print(f"  - Amount cocok persis (AMOUNT_ORI_MIN1 == CCOS_BAL_DUE): {exact_amt}")
print(f"  - Amount tidak cocok: {len(df_ins_single) - exact_amt}")
print("Flag prod distribution di Insured only single fac:")
print(df_ins_single["FLAG_PROD"].value_counts())

print("\nSample 10 baris 'Insured only' (single fac):")
sample_cols = ["RECEIPT NO", "INSURED_ORI", "AMOUNT ORI", "CCOS_BAL_DUE", "DIFERENCE", "CCOS_REF_CODE", "FLAG_PROD"]
print(df_ins_single[sample_cols].head(10).to_string())

print("\n" + "="*70)
print("AUDIT 4: MENGAPA SEMUA ADJUSTMENT SEBAGIAN 'DENGAN AKUMULASI'?")
print("="*70)
adj_seb = df_v1[df_v1["FLAG_PROD"].str.contains("Adjustment sebagian", na=False)]
print(f"Jumlah Adjustment sebagian: {len(adj_seb)}")
print("Distribusi SKENARIO pada Adjustment sebagian:")
print(adj_seb["SKENARIO"].value_counts())
print("Distribusi _OSBAL_ROW_COUNT pada Adjustment sebagian (jika ada):")
# Hitung berapa banyak row OSBAL per CCOS_REF_CODE untuk kasus ini
if "CCOS_REF_CODE" in adj_seb.columns:
    osbal_fac_counts = df_osbal["CCOS_REF_CODE"].value_counts()
    facs_in_adj = adj_seb["CCOS_REF_CODE"].dropna()
    row_counts = facs_in_adj.map(osbal_fac_counts)
    print("Jumlah baris OSBAL untuk FAC code pada baris Adjustment sebagian:")
    print(row_counts.value_counts().head(10))

print("\n" + "="*70)
print("AUDIT 5: CROSS-CHECK SLIP VALID (Berapa slip suspend yang benar-benar ada di OSBAL/FACUL?)")
print("="*70)
valid_sus_slips = df_sus[~df_sus["SLIP_NO_ORI"].astype(str).str.strip().isin(["-", "--", "", "nan", "None"])]["SLIP_NO_ORI"].unique()
print(f"Total slip unik valid di Suspend: {len(valid_sus_slips)}")
# Cek di OSBAL
osbal_all_slips = set()
for c in [c for c in df_osbal.columns if "slip" in c.lower()]:
    osbal_all_slips.update(df_osbal[c].dropna().astype(str).str.strip().tolist())
# Cek di FACUL
facul_all_slips = set()
for c in [c for c in df_facul.columns if "slip" in c.lower()]:
    facul_all_slips.update(df_facul[c].dropna().astype(str).str.strip().tolist())

exact_in_osb = [s for s in valid_sus_slips if s in osbal_all_slips]
exact_in_fac = [s for s in valid_sus_slips if s in facul_all_slips]
print(f"Slip valid suspend yang match persis di OSBAL: {len(exact_in_osb)}")
print(f"Slip valid suspend yang match persis di FACUL: {len(exact_in_fac)}")
print(f"Slip valid suspend yang ada di OSBAL ATAU FACUL: {len(set(exact_in_osb) | set(exact_in_fac))}")
