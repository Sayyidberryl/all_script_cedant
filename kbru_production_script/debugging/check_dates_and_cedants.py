import os
import sys
sys.modules['numexpr'] = None
sys.modules['bottleneck'] = None
import pandas as pd

DATA_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "data")
df_sus = pd.read_excel(os.path.join(DATA_DIR, "kbru_output_suspend.xlsx"))
df_osbal = pd.read_excel(os.path.join(DATA_DIR, "kbru_output_osbal.xlsx"))
df_facul = pd.read_excel(os.path.join(DATA_DIR, "kbru_output_facul.xlsx"))

print("=== SUSPEND ===")
print("Total rows:", len(df_sus))
sus_dates = pd.to_datetime(df_sus["RECEIPT DATE"], errors="coerce")
print("Min date:", sus_dates.min(), "Max date:", sus_dates.max())
print("Years distribution in Suspend:")
print(sus_dates.dt.year.value_counts())
print("Cedant names in Suspend:")
print(df_sus["CEDANT NAME"].value_counts().head(5))

print("\n=== OSBAL ===")
print("Total rows:", len(df_osbal))
osb_dates = pd.to_datetime(df_osbal["FAC_COM_DATE"], errors="coerce")
print("Min date:", osb_dates.min(), "Max date:", osb_dates.max())
print("Years distribution in OSBAL:")
print(osb_dates.dt.year.value_counts())
print("Cedant names in OSBAL (CLSDT_CEDANT / CEDANT_NAME):")
for col in [c for c in df_osbal.columns if 'ced' in c.lower()]:
    print(col, df_osbal[col].value_counts().head(3))

print("\n=== FACUL ===")
print("Total rows:", len(df_facul))
fac_dates = pd.to_datetime(df_facul["FAC_COM_DATE"], errors="coerce") if "FAC_COM_DATE" in df_facul.columns else None
if fac_dates is not None:
    print("Min date:", fac_dates.min(), "Max date:", fac_dates.max())
    print("Years distribution in FACUL:")
    print(fac_dates.dt.year.value_counts().head(5))
