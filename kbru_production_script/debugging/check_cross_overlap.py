import os, sys
sys.modules['numexpr'] = None
sys.modules['bottleneck'] = None
import pandas as pd
import re

DATA_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "data")
df_sus = pd.read_excel(os.path.join(DATA_DIR, "kbru_output_suspend.xlsx"))
df_osbal = pd.read_excel(os.path.join(DATA_DIR, "kbru_output_osbal.xlsx"))

# Kumpulkan semua clean polis dari suspend
sus_polis_clean = set()
for c in [col for col in df_sus.columns if 'polis' in col.lower() or 'policy' in col.lower()]:
    for val in df_sus[c].dropna():
        s = re.sub(r'[^A-Z0-9]', '', str(val).upper())
        if len(s) >= 8:
            sus_polis_clean.add(s)

print(f"Total unique clean polis di Suspend (>=8 char): {len(sus_polis_clean)}")

# Kumpulkan semua clean values dari OSBAL di FAC_SLIP_ORI dan SLIP_CLEAN_*
osbal_slip_clean = set()
for c in ['FAC_SLIP_ORI', 'SLIP_CLEAN_1', 'SLIP_CLEAN_2', 'SLIP_CLEAN_3', 'SLIP_CLEAN_4', 'CLSDT_SLIP_NO']:
    if c in df_osbal.columns:
        for val in df_osbal[c].dropna():
            s = re.sub(r'[^A-Z0-9]', '', str(val).upper())
            if len(s) >= 8:
                osbal_slip_clean.add(s)

print(f"Total unique clean values di OSBAL Slip cols (>=8 char): {len(osbal_slip_clean)}")

# Cek intersection antara Suspend Polis dan OSBAL Slip!
overlap = sus_polis_clean.intersection(osbal_slip_clean)
print(f"Berapa banyak POLIS dari Suspend yang cocok persis dengan SLIP di OSBAL? {len(overlap)}")

# Juga cek prefix / core match (misal 8 digit nomor depan 00004903)
sus_core = {s[:8] for s in sus_polis_clean if s[:4] == '0000'}
osbal_core = {s[:8] for s in osbal_slip_clean if s[:4] == '0000'}
overlap_core = sus_core.intersection(osbal_core)
print(f"Berapa banyak core polis (0000xxxx) dari Suspend yang cocok dengan OSBAL Slip? {len(overlap_core)}")
if overlap:
    print("Contoh yang overlap:", list(overlap)[:5])
if overlap_core:
    print("Contoh core overlap:", list(overlap_core)[:5])
