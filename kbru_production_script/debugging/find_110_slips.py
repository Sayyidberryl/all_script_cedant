import os, sys
sys.modules['numexpr'] = None
sys.modules['bottleneck'] = None
import pandas as pd
import re

DATA_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "data")
df_sus = pd.read_excel(os.path.join(DATA_DIR, "kbru_output_suspend.xlsx"))
df_osbal = pd.read_excel(os.path.join(DATA_DIR, "kbru_output_osbal.xlsx"))
df_facul = pd.read_excel(os.path.join(DATA_DIR, "kbru_output_facul.xlsx"))

# Cari semua file di data dir yang mungkin berisi slipdb
print("Data directory contents:")
for f in os.listdir(DATA_DIR):
    if f.endswith('.xlsx'):
        print(" ", f, f"(size: {os.path.getsize(os.path.join(DATA_DIR, f))} bytes)")

# Kumpulkan semua 110... slips dari suspend
sus_110 = set()
for s in df_sus["SLIP_NO_ORI"].dropna().astype(str):
    m = re.match(r'(110[A-Z0-9]+)', s.strip().upper())
    if m:
        base = re.sub(r'-\d+$', '', m.group(1))
        sus_110.add(base)

print(f"\nTotal unique 110 base slips in Suspend: {len(sus_110)}")

# Cek di OSBAL
osbal_all_text = df_osbal.astype(str).values.flatten()
found_in_osbal = sum(1 for s in sus_110 if any(s in text for text in osbal_all_text))
print(f"Berapa banyak 110 slip dari Suspend ditemukan di MANAPUN di kolom OSBAL? {found_in_osbal}")

# Cek di FACUL
facul_all_text = df_facul.astype(str).values.flatten()
found_in_facul = sum(1 for s in sus_110 if any(s in text for text in facul_all_text))
print(f"Berapa banyak 110 slip dari Suspend ditemukan di MANAPUN di kolom FACUL? {found_in_facul}")
