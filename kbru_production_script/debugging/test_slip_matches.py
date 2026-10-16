import os, sys
sys.modules['numexpr'] = None
sys.modules['bottleneck'] = None
import pandas as pd

df_sus = pd.read_excel("data/kbru_output_suspend.xlsx")
df_osbal = pd.read_excel("data/kbru_output_osbal.xlsx")

sus_slips = df_sus['SLIP_CLEAN_1'].dropna().tolist()

# Search across all columns of osbal
found_exact = 0
found_sub = 0
found_base = 0

osbal_str = df_osbal.astype(str)

for s in sus_slips[:50]:
    s_clean = str(s).strip()
    base = s_clean.split('-')[0]
    in_any = (osbal_str == s_clean).any().any()
    in_base = (osbal_str == base).any().any()
    in_sub = osbal_str.apply(lambda col: col.str.contains(base, regex=False)).any().any()
    print(f"Suspend Slip: {s_clean} -> exact: {in_any}, base: {in_base}, sub: {in_sub}")
