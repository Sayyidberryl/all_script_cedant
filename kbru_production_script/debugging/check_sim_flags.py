import os, sys
sys.modules['numexpr'] = None
sys.modules['bottleneck'] = None
import pandas as pd
import numpy as np

# Import functions from simulate_full_v1_fix logic and test flag_prod
from debugging.simulate_full_v1_fix import raw_results, osbal_rows, facode_osbal_idx, suspend_rows, _compute_derived_cols if hasattr(sys.modules.get('debugging.simulate_full_v1_fix'), '_compute_derived_cols') else (None, None, None, None, None)

