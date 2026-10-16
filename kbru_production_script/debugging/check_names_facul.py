import os, sys
sys.modules['numexpr'] = None
sys.modules['bottleneck'] = None
import pandas as pd

df_sus = pd.read_excel("data/kbru_output_suspend.xlsx")
df_osbal = pd.read_excel("data/kbru_output_osbal.xlsx")
df_facul = pd.read_excel("data/kbru_output_facul.xlsx")

for name in ['CENTRAL PROTEINA PRIMA', 'CITRA BAHARI PERSADA']:
    m_osb = df_osbal[df_osbal.apply(lambda r: any(name in str(v) for v in r.values), axis=1)]
    m_fac = df_facul[df_facul.apply(lambda r: any(name in str(v) for v in r.values), axis=1)]
    print(f"Name: {name} -> in OSBAL: {len(m_osb)}, in FACUL: {len(m_fac)}")
    if len(m_osb) > 0:
        print("  OSBAL sample:", m_osb[['CCOS_REF_CODE', 'INSURED_CLEAN_1', 'POLIS_CLEAN_1', 'SLIP_CLEAN_1']].head(2).to_dict('records'))
    if len(m_fac) > 0:
        print("  FACUL sample:", m_fac[['FAC_CODE', 'FAC_INSURED_1', 'POLIS_CLEAN_1', 'FAC_SLIP_CLEAN_1']].head(2).to_dict('records'))
