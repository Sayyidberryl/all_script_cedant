import sys
sys.modules['numexpr'] = None
sys.modules['bottleneck'] = None

import os
import re
import pandas as pd

DATA_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "data")

v1_path = os.path.join(DATA_DIR, "final_output_v1.xlsx")
df_v1 = pd.read_excel(v1_path)
df_osbal = pd.read_excel(os.path.join(DATA_DIR, "kbru_output_osbal.xlsx"))
df_facul = pd.read_excel(os.path.join(DATA_DIR, "kbru_output_facul.xlsx"))

df_unmatch = df_v1[df_v1["FLAG_PROD"] == "Unmatching"].copy()
print(f"Total Unmatching rows: {len(df_unmatch)}")

# Check against FACUL directly
match_slip_facul = 0
match_polis_facul = 0
match_ins_facul = 0
totally_absent = 0

facul_slips = set(df_facul["FAC_SLIP"].dropna().astype(str).str.strip().str.upper())
facul_polis = set(df_facul["FAC_POLICY_NO"].dropna().astype(str).str.strip().str.upper())
facul_ins = set(df_facul["FAC_INSURED_ORI"].dropna().astype(str).str.strip().str.upper())

# Juga base slip
facul_base_slips = {re.sub(r'-\d+$', '', s) for s in facul_slips}

reasons = {
    "Ada di FACUL (Slip) tapi tidak ada di OSBAL": 0,
    "Ada di FACUL (Polis) tapi tidak ada di OSBAL": 0,
    "Ada di FACUL (Insured) tapi tidak ada di OSBAL": 0,
    "Sama sekali tidak ada di OSBAL maupun FACUL (Data Baru / Unregistered)": 0,
}

details = []
for idx, r in df_unmatch.iterrows():
    s_raw = str(r.get("SLIP_NO", "")).strip().upper()
    s_base = re.sub(r'-\d+$', '', s_raw)
    p_raw = str(r.get("POLIS", "")).strip().upper()
    i_raw = str(r.get("INSURED_ORI", "")).strip().upper()

    has_s = (s_raw in facul_slips) or (s_base in facul_base_slips)
    has_p = (p_raw in facul_polis)
    has_i = any(i_raw in fi or fi in i_raw for fi in facul_ins if len(i_raw) >= 5)

    if has_s:
        reasons["Ada di FACUL (Slip) tapi tidak ada di OSBAL"] += 1
        details.append(("Slip match FACUL only", s_raw, i_raw))
    elif has_p:
        reasons["Ada di FACUL (Polis) tapi tidak ada di OSBAL"] += 1
        details.append(("Polis match FACUL only", p_raw, i_raw))
    elif has_i:
        reasons["Ada di FACUL (Insured) tapi tidak ada di OSBAL"] += 1
        details.append(("Insured match FACUL only", i_raw))
    else:
        reasons["Sama sekali tidak ada di OSBAL maupun FACUL (Data Baru / Unregistered)"] += 1
        details.append(("Completely absent", s_raw, p_raw, i_raw))

print("\nBreakdown 329 Unmatching:")
for k, v in reasons.items():
    print(f"  - {k:<70}: {v:>3} baris ({v/len(df_unmatch)*100:.1f}%)")

print("\nContoh yang Sama Sekali Tidak Ada di OSBAL maupun FACUL (Data Baru):")
absent_samples = [d for d in details if d[0] == "Completely absent"][:10]
for d in absent_samples:
    print("  ", d)
