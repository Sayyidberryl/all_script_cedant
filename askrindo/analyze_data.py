"""Script analisis karakteristik data Askrindo untuk matching."""
import sys
sys.modules['numexpr'] = None
sys.modules['bottleneck'] = None
import pandas as pd
import warnings
warnings.filterwarnings('ignore')

df_sus = pd.read_excel('data/askrindo_output_suspend.xlsx')
df_osbal = pd.read_excel('data/askrindo_output_osbal.xlsx')

print('=== ASKRINDO KARAKTERISTIK ===')
print()
print('1. SUSPEND - Polis null/empty rate:')
polis_empty = df_sus['POLIS'].astype(str).str.strip().isin(['-','nan','','None'])
print(f'   Polis empty/dash: {polis_empty.sum()}/{len(df_sus)} ({100*polis_empty.mean():.0f}%)')
print()

print('2. SUSPEND - Slip validity:')
slip_empty = df_sus['SLIP NO'].isna() | df_sus['SLIP NO'].astype(str).str.strip().isin(['','nan','None'])
print(f'   Slip empty: {slip_empty.sum()}/{len(df_sus)}')
slip_long = df_sus['SLIP_NO_CLEAN'].astype(str).str.len() >= 10
print(f'   Slip clean >= 10 char: {slip_long.sum()}/{len(df_sus)}')
print()

print('3. SUSPEND - INSURED_CLEAN_1 stats:')
ins1_nonempty = (df_sus['INSURED_CLEAN_1'].astype(str).str.strip() != '').sum()
print(f'   Non-empty: {ins1_nonempty}')
print()

print('4. OSBAL - Polis clean stats:')
print(f'   POLIS_CLEAN_1 non-null: {df_osbal["POLIS_CLEAN_1"].notna().sum()}')
print(f'   POLIS_CLEAN_2 non-null: {df_osbal["POLIS_CLEAN_2"].notna().sum()}')
print()

print('5. OSBAL - SLIP clean cols:')
slip_cols = [c for c in df_osbal.columns if 'slip' in c.lower() and ('clean' in c.lower() or 'no' in c.lower())]
print(f'   OSBAL slip clean cols: {slip_cols}')
slip_vals = df_osbal['SLIP_NO_CLEAN'].dropna()
print(f'   SLIP_NO_CLEAN non-null: {len(slip_vals)}')
print(f'   Sample slip values: {slip_vals.head(10).tolist()}')
print()

print('6. Matching analysis - SLIP sus vs OSBAL:')
sus_slips = df_sus['SLIP_NO_CLEAN'].astype(str).str.strip().tolist()
osbal_slips = set(df_osbal['SLIP_NO_CLEAN'].dropna().astype(str).str.strip().tolist())
osbal_clsdt_slip = set(df_osbal['CLSDT_SLIP_NO'].dropna().astype(str).str.strip().tolist())

exact_matches_slip = []
for i, ss in enumerate(sus_slips):
    if ss and ss not in ['-','nan','']:
        if ss in osbal_slips or ss in osbal_clsdt_slip:
            exact_matches_slip.append((i, ss))
            print(f'   EXACT SLIP match row {i}: {ss}')
print(f'   Total EXACT SLIP matches: {len(exact_matches_slip)}/{len(sus_slips)}')
print()

print('7. Matching analysis - INSURED sus vs OSBAL:')
sus_ins = df_sus['INSURED_CLEAN_1'].astype(str).str.strip().tolist()
osbal_ins = set(df_osbal['INSURED_CLEAN_1'].dropna().astype(str).str.strip().tolist())
exact_matches_ins = []
for i, ins in enumerate(sus_ins):
    if ins and ins not in ['nan','']:
        if ins in osbal_ins:
            exact_matches_ins.append((i, ins))
            print(f'   EXACT INS match row {i}: {ins}')
print(f'   Total EXACT INSURED matches: {len(exact_matches_ins)}/{len(sus_ins)}')
print()

print('8. LIKE matching - SLIP (substring):')
like_slip_matches = []
for i, ss in enumerate(sus_slips):
    if ss and ss not in ['-','nan',''] and len(ss) >= 8:
        for os_slip in osbal_slips:
            if os_slip and len(os_slip) >= 8:
                if ss in os_slip or os_slip in ss:
                    like_slip_matches.append((i, ss, os_slip))
                    print(f'   LIKE SLIP row {i}: sus={ss} | osbal={os_slip}')
                    break
print(f'   Total LIKE SLIP matches: {len(like_slip_matches)}/{len(sus_slips)}')
print()

print('9. CLSDT_SLIP_NO in OSBAL matches suspend:')
for i, ss in enumerate(sus_slips):
    if ss and ss not in ['-','nan','']:
        # Try matching against raw CLSDT_SLIP_NO
        for cs in osbal_clsdt_slip:
            if cs and cs not in ['nan','']:
                cs_clean = cs.replace('.', '').replace('-', '').replace('/', '')
                ss_clean = ss.replace('.', '').replace('-', '').replace('/', '')
                if cs_clean == ss_clean or cs_clean in ss_clean or ss_clean in cs_clean:
                    print(f'   CLSDT SLIP match row {i}: sus_clean={ss} | clsdt_raw={cs}')

print()
print('=== FULL SUS DATA ===')
for i, row in df_sus.iterrows():
    print(f'Row {i}: POLIS={str(row["POLIS"]).strip()} | POLIS_CLN={str(row["POLIS_CLEAN"]).strip()} | SLIP_CLN={str(row["SLIP_NO_CLEAN"]).strip()} | INS1={str(row["INSURED_CLEAN_1"]).strip()} | CURR={row["CURR ORI"]}')
