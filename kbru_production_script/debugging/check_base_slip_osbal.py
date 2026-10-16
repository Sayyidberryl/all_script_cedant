import os, sys, re
sys.modules['numexpr'] = None
sys.modules['bottleneck'] = None
import pandas as pd

df_sus = pd.read_excel("data/kbru_output_suspend.xlsx")
df_osbal = pd.read_excel("data/kbru_output_osbal.xlsx")

# Check 5 rows from suspend with -NN suffix
for idx in range(10):
    s = str(df_sus.loc[idx, 'SLIP_CLEAN_1']).strip()
    p = str(df_sus.loc[idx, 'POLIS_CLEAN_1']).strip()
    ins = str(df_sus.loc[idx, 'INSURED_1']).strip()
    amt = df_sus.loc[idx, 'AMOUNT ORI']
    base = re.sub(r'-\d+$', '', s)
    
    # search base in osbal
    matches = df_osbal[df_osbal.apply(lambda r: any(base in str(v) for v in r.values), axis=1)]
    print(f"Suspend {idx}: Slip={s} (base={base}), Polis={p}, Insured={ins}, Amt={amt}")
    if len(matches) > 0:
        for m_idx, mr in matches.head(2).iterrows():
            print(f"  -> OSBAL {m_idx}: FAC_CODE={mr.get('CCOS_REF_CODE')}, Doc={mr.get('CCOS_DOC_NO')}, Slip={mr.get('CLSDT_SLIP_NO')}, BalDue={mr.get('CCOS_BAL_DUE')}, Insured={mr.get('INSURED_CLEAN_1')}")
    else:
        print("  -> No match by base slip in OSBAL")
