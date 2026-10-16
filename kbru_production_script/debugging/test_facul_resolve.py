import os, sys, re
sys.modules['numexpr'] = None
sys.modules['bottleneck'] = None
import pandas as pd

DATA_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "data")
df_facul = pd.read_excel(os.path.join(DATA_DIR, "kbru_output_facul.xlsx"))
df_osbal = pd.read_excel(os.path.join(DATA_DIR, "kbru_output_osbal.xlsx"))
df_v1 = pd.read_excel(os.path.join(DATA_DIR, "final_output_v1.xlsx"))
df_ins_gt1 = df_v1[df_v1["SKENARIO"] == "Insured only (> 1 Fac)"]

osbal_facs = set(df_osbal["CCOS_REF_CODE"].dropna().astype(str).str.strip())

# Cek polis di FACUL
facul_polis_to_fac = {}
for idx, r in df_facul.iterrows():
    fac = str(r.get("FAC_CODE", "")).strip()
    if not fac or fac not in osbal_facs: continue
    for col in ["FAC_POLICY_NO", "POLIS_CLEAN_1", "POLIS_CLEAN_2", "POLIS_CLEAN_3", "FAC_SLIP"]:
        val = r.get(col)
        if pd.isna(val): continue
        val_str = str(val).strip().upper()
        if val_str in ["-", "NONE", "TBA", ""]: continue
        
        # Ekstrak nomor
        m = re.search(r'0*(\d{4,8})[/.\s]+([A-Z]{3,5})', val_str)
        if m:
            facul_polis_to_fac.setdefault(f"{m.group(1)}{m.group(2)}", set()).add(fac)
        c = re.sub(r'[^A-Z0-9]', '', val_str)
        if len(c) >= 6:
            facul_polis_to_fac.setdefault(c, set()).add(fac)

hits_in_facul = 0
for idx, r in df_ins_gt1.iterrows():
    p = str(r.get("POLIS", "")).strip().upper()
    m = re.search(r'0*(\d{4,8})[/.\s]+([A-Z]{3,5})', p)
    k = f"{m.group(1)}{m.group(2)}" if m else ""
    c = re.sub(r'[^A-Z0-9]', '', p)
    if (k and k in facul_polis_to_fac) or (len(c) >= 6 and c in facul_polis_to_fac):
        hits_in_facul += 1

print(f"Berapa banyak baris Insured only (> 1 Fac) yang bisa di-resolve via FACUL -> OSBAL? {hits_in_facul}")
