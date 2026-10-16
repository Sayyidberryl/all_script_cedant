import os, sys
sys.modules['numexpr'] = None
sys.modules['bottleneck'] = None
import pandas as pd

DATA_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "data")
df_osbal = pd.read_excel(os.path.join(DATA_DIR, "kbru_output_osbal.xlsx"))

print("CCOS_COMP_NAME counts:")
print(df_osbal["CCOS_COMP_NAME"].value_counts().head(5))
print("CCOS_REF_COMP_NAME counts:")
print(df_osbal["CCOS_REF_COMP_NAME"].value_counts().head(5))
