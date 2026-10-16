import os, sys
sys.modules['numexpr'] = None
sys.modules['bottleneck'] = None
import pandas as pd

df_osbal = pd.read_excel("data/kbru_output_osbal.xlsx", nrows=10)
for idx, r in df_osbal.iterrows():
    print(f"OSBAL Row {idx}:")
    print(f"  FAC_INSURED_ORI: {r.get('FAC_INSURED_ORI')}")
    print(f"  INSURED_CLEAN_1: {r.get('INSURED_CLEAN_1')}")
    print(f"  FAC_POLICY_NO_ORI: {r.get('FAC_POLICY_NO_ORI')}")
    print(f"  POLIS_CLEAN_1:   {r.get('POLIS_CLEAN_1')}")
    print(f"  FAC_SLIP_ORI:    {r.get('FAC_SLIP_ORI')}")
    print(f"  SLIP_CLEAN_1:    {r.get('SLIP_CLEAN_1')}")
    print(f"  CLSDT_POLICY_NO: {r.get('CLSDT_POLICY_NO')}")
    print(f"  CLSDT_SLIP_NO:   {r.get('CLSDT_SLIP_NO')}")
    print(f"  CCOS_REF_CODE:   {r.get('CCOS_REF_CODE')}")
