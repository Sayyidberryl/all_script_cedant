import os, sys
sys.modules['numexpr'] = None
sys.modules['bottleneck'] = None
import pandas as pd

DATA_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "data")
df_osbal = pd.read_excel(os.path.join(DATA_DIR, "kbru_output_osbal.xlsx"), nrows=5)
print("OSBAL cols:", df_osbal.columns.tolist())

df_facul = pd.read_excel(os.path.join(DATA_DIR, "kbru_output_facul.xlsx"), nrows=5)
print("FACUL cols:", df_facul.columns.tolist())
