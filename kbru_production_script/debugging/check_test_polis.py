import os
import sys
sys.modules['numexpr'] = None
sys.modules['bottleneck'] = None
import pandas as pd

DATA_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "data")
df_osbal = pd.read_excel(os.path.join(DATA_DIR, "kbru_output_osbal.xlsx"))
df_facul = pd.read_excel(os.path.join(DATA_DIR, "kbru_output_facul.xlsx"))

test_polis = "00006166/FPAR/038.0000007-KBRURE/11/2025"
test_clean = "00006166FPAR0380000007KBRURE112025"

print("Cek test_polis di OSBAL:")
pol_cols = [c for c in df_osbal.columns if 'pol' in c.lower()]
found = False
for c in pol_cols:
    m = df_osbal[df_osbal[c].astype(str).str.contains("00006166", na=False)]
    if len(m) > 0:
        print(f"  Found in {c}: {len(m)} rows")
        print(m[["CCOS_DOC_NO", "CCOS_REF_CODE", c]].to_string())
        found = True
if not found:
    print("  Tidak ditemukan di OSBAL")

print("\nCek test_polis di FACUL:")
fac_pol_cols = [c for c in df_facul.columns if 'pol' in c.lower()]
found_f = False
for c in fac_pol_cols:
    m = df_facul[df_facul[c].astype(str).str.contains("00006166", na=False)]
    if len(m) > 0:
        print(f"  Found in {c}: {len(m)} rows")
        print(m[["FAC_CODE", c]].to_string())
        found_f = True
if not found_f:
    print("  Tidak ditemukan di FACUL")
