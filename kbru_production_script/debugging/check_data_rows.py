import os
import sys

sys.modules['numexpr'] = None
sys.modules['bottleneck'] = None

import pandas as pd

for fname in ['kbru_output_facul.xlsx', 'kbru_output_osbal.xlsx', 'kbru_output_suspend.xlsx']:
    p = os.path.join("data", fname)
    df = pd.read_excel(p)
    print(f"{fname}: {df.shape[0]} rows, {df.shape[1]} cols")
