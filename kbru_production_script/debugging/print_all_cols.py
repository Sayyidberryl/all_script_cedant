import os, sys
sys.modules['numexpr'] = None
sys.modules['bottleneck'] = None
import pandas as pd

df_facul = pd.read_excel("data/kbru_output_facul.xlsx", nrows=1)
df_osbal = pd.read_excel("data/kbru_output_osbal.xlsx", nrows=1)
df_sus = pd.read_excel("data/kbru_output_suspend.xlsx", nrows=1)

print("FACUL COLS:", df_facul.columns.tolist())
print("OSBAL COLS:", df_osbal.columns.tolist())
print("SUSPEND COLS:", df_sus.columns.tolist())
