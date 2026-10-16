import os, sys
sys.modules['numexpr'] = None
sys.modules['bottleneck'] = None
import pandas as pd

DATA_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "data")
df_v1 = pd.read_excel(os.path.join(DATA_DIR, "final_output_v1.xlsx"))

df_ins_gt1 = df_v1[df_v1["SKENARIO"] == "Insured only (> 1 Fac)"]
print("Top 20 Insured in 'Insured only (> 1 Fac)':")
print(df_ins_gt1["INSURED_ORI"].value_counts().head(20))
