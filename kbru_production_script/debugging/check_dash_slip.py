import os
import sys
sys.modules['numexpr'] = None
sys.modules['bottleneck'] = None

import pandas as pd
DATA_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "data")
df_sus = pd.read_excel(os.path.join(DATA_DIR, "kbru_output_suspend.xlsx"))
df_osbal = pd.read_excel(os.path.join(DATA_DIR, "kbru_output_osbal.xlsx"))

print("Row 1567 in df_sus:")
row = df_sus.iloc[1567]
for k, v in row.items():
    if pd.notna(v) and str(v).strip():
        print(f"  {k}: {v}")

print("\nCek apakah di OSBAL ada '-' di kolom slip:")
slip_cols = [c for c in df_osbal.columns if 'slip' in c.lower()]
print("OSBAL slip cols:", slip_cols)
for c in slip_cols:
    dash_cnt = (df_osbal[c].astype(str).str.strip() == '-').sum()
    print(f"  {c} with '-': {dash_cnt}")
