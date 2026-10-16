import os, sys
sys.modules['numexpr'] = None
sys.modules['bottleneck'] = None
import pandas as pd

df_osbal = pd.read_excel("data/kbru_output_osbal.xlsx")
s = "11055124F00000144"

for col in df_osbal.columns:
    matches = df_osbal[df_osbal[col].astype(str).str.contains(s, regex=False)]
    if len(matches) > 0:
        print(f"Found in col: {col}")
        for idx, r in matches.iterrows():
            print(f"  Row {idx}:")
            print(f"    CCOS_REF_CODE: {r.get('CCOS_REF_CODE')}")
            print(f"    FAC_SLIP_ORI:  {r.get('FAC_SLIP_ORI')}")
            print(f"    SLIP_CLEAN_1:  {r.get('SLIP_CLEAN_1')}")
            print(f"    SLIP_CLEAN_2:  {r.get('SLIP_CLEAN_2')}")
            print(f"    CLSDT_SLIP_NO: {r.get('CLSDT_SLIP_NO')}")
            print(f"    POLIS_CLEAN_1: {r.get('POLIS_CLEAN_1')}")
