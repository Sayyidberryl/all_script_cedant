import os, sys
sys.modules['numexpr'] = None
sys.modules['bottleneck'] = None
import pandas as pd

df_facul = pd.read_excel("data/kbru_output_facul.xlsx", nrows=1)
df_osbal = pd.read_excel("data/kbru_output_osbal.xlsx", nrows=1)
df_sus = pd.read_excel("data/kbru_output_suspend.xlsx", nrows=1)

def _get_clean_cols(columns, prefix: str) -> list:
    if columns is None or len(columns) == 0:
        return []
    alt_map = {
        "clean polis": ["clean polis", "policy_clean", "polis_clean", "polis_cln", "policy_cln"],
        "clean slip": ["clean slip", "slip_clean", "fac_slip_clean", "slip_no_cln", "slip_cln"],
        "clean insured": ["clean insured", "insured_clean", "fac_insured_", "insured_"],
    }
    prefixes = alt_map.get(prefix.lower(), [prefix.lower()])
    res = []
    for c in columns:
        clow = str(c).lower().strip()
        if any(clow.startswith(p) for p in prefixes):
            res.append(c)
    return res

print("SUSPEND polis:", _get_clean_cols(df_sus.columns, "clean polis"))
print("SUSPEND slip:", _get_clean_cols(df_sus.columns, "clean slip"))
print("SUSPEND insured:", _get_clean_cols(df_sus.columns, "clean insured"))

print("\nOSBAL polis:", _get_clean_cols(df_osbal.columns, "clean polis"))
print("OSBAL slip:", _get_clean_cols(df_osbal.columns, "clean slip"))
print("OSBAL insured:", _get_clean_cols(df_osbal.columns, "clean insured"))

print("\nFACUL polis:", _get_clean_cols(df_facul.columns, "clean polis"))
print("FACUL slip:", _get_clean_cols(df_facul.columns, "clean slip"))
print("FACUL insured:", _get_clean_cols(df_facul.columns, "clean insured"))
