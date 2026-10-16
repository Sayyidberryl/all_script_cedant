import os, sys
sys.modules['numexpr'] = None
sys.modules['bottleneck'] = None
import pandas as pd

df_facul = pd.read_excel("data/kbru_output_facul.xlsx", nrows=20)
for idx, r in df_facul.iterrows():
    print(f"FACUL {idx}:")
    print(f"  FAC_CODE:   {r.get('FAC_CODE')}")
    print(f"  FAC_POLICY: {r.get('FAC_POLICY_NO')}")
    print(f"  FAC_SLIP:   {r.get('FAC_SLIP')}")
    print(f"  POLIS_CLN1: {r.get('POLIS_CLEAN_1')}")
    print(f"  SLIP_CLN1:  {r.get('FAC_SLIP_CLEAN_1')}")
    print(f"  INSURED:    {r.get('FAC_INSURED_1')}")
