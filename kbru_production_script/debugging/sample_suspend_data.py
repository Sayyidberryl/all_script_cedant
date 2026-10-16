import os, sys
sys.modules['numexpr'] = None
sys.modules['bottleneck'] = None
import pandas as pd

df_sus = pd.read_excel("data/kbru_output_suspend.xlsx", nrows=10)
for idx, r in df_sus.iterrows():
    print(f"Row {idx}:")
    print(f"  INSURED_ORI: {r.get('INSURED_ORI')}")
    print(f"  INSURED_1:   {r.get('INSURED_1')}")
    print(f"  POLIS_ORI:   {r.get('POLIS_ORI')}")
    print(f"  POLIS_CLN_1: {r.get('POLIS_CLEAN_1')}")
    print(f"  SLIP_NO_ORI: {r.get('SLIP_NO_ORI')}")
    print(f"  SLIP_CLN_1:  {r.get('SLIP_CLEAN_1')}")
