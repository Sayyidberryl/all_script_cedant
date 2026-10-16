import os, sys, re
sys.modules['numexpr'] = None
sys.modules['bottleneck'] = None
import pandas as pd

DATA_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "data")
SUSPEND_FILE = os.path.join(DATA_DIR, "kbru_output_suspend.xlsx")
OSBAL_FILE   = os.path.join(DATA_DIR, "kbru_output_osbal.xlsx")
FACUL_FILE   = os.path.join(DATA_DIR, "kbru_output_facul.xlsx")

df_sus = pd.read_excel(SUSPEND_FILE)
df_osbal = pd.read_excel(OSBAL_FILE)
df_facul = pd.read_excel(FACUL_FILE)

DUMMY_VALUES = frozenset({"-", "--", "---", "0", "00", "000", "NONE", "NULL", "TBA", "NA", "N/A", "SELISIH", "NO", "NON", ""})

def is_valid_token(s: str) -> bool:
    if not s or (isinstance(s, float) and pd.isna(s)): return False
    cleaned = str(s).strip().upper()
    if cleaned in DUMMY_VALUES: return False
    alphanumeric = re.sub(r'[^A-Z0-9]', '', cleaned)
    return len(alphanumeric) >= 3

def clean_alphanumeric(s: str) -> str:
    if not s or (isinstance(s, float) and pd.isna(s)): return ""
    return re.sub(r'[^A-Z0-9]', '', str(s).upper())

def extract_kbru_polis_keys(polis_str: str) -> list:
    if not is_valid_token(polis_str): return []
    s = str(polis_str).strip().upper()
    keys = set()
    c = clean_alphanumeric(s)
    if len(c) >= 6:
        keys.add(c)
        no_broker = re.sub(r'KBRURE|SRB', '', c)
        if len(no_broker) >= 6:
            keys.add(no_broker)
    m = re.search(r'0*(\d{3,8})[/.\s]+([A-Z]{3,5})[/.\s]+0*(\d+)[/.\s]+0*(\d+)', s)
    if m:
        doc_num, cls, p3, p4 = m.groups()
        keys.add(f"{doc_num}{cls}{p3}{p4}")
        keys.add(f"{doc_num}{cls}")
    else:
        m2 = re.search(r'0*(\d{3,8})[/.\s]+([A-Z]{3,5})', s)
        if m2:
            doc_num, cls = m2.groups()
            keys.add(f"{doc_num}{cls}")
    if '+' in s:
        for part in s.split('+'):
            if is_valid_token(part):
                keys.update(extract_kbru_polis_keys(part))
    return [k for k in keys if len(k) >= 3 and k not in DUMMY_VALUES]

def extract_slip_keys(slip_str: str) -> list:
    if not is_valid_token(slip_str): return []
    s = str(slip_str).strip().upper()
    keys = set()
    keys.add(s)
    c = clean_alphanumeric(s)
    if len(c) >= 3:
        keys.add(c)
    base = re.sub(r'-\d+$', '', s)
    if base and base != s and is_valid_token(base):
        keys.add(base)
        c_base = clean_alphanumeric(base)
        if len(c_base) >= 3:
            keys.add(c_base)
    if '+' in s:
        for part in s.split('+'):
            if is_valid_token(part):
                keys.update(extract_slip_keys(part))
    return [k for k in keys if len(k) >= 3 and k not in DUMMY_VALUES]

osbal_fac_map = {}
for idx, r in df_osbal.iterrows():
    fac = str(r.get("CCOS_REF_CODE", "")).strip()
    if fac:
        osbal_fac_map.setdefault(fac, []).append(idx)

# OSBAL
osbal_slip_idx = {}
osbal_polis_idx = {}
for idx, r in df_osbal.iterrows():
    for c in ["CLSDT_SLIP_NO", "SLIP_CLEAN_1", "SLIP_CLEAN_2", "FAC_SLIP_ORI"]:
        val = r.get(c)
        for k in extract_slip_keys(val):
            osbal_slip_idx.setdefault(k, set()).add(idx)
    for c in ["CLSDT_POLICY_NO", "POLIS_CLEAN_1", "POLIS_CLEAN_2", "FAC_POLICY_NO_ORI", "FAC_SLIP_ORI"]:
        val = r.get(c)
        for k in extract_kbru_polis_keys(val):
            osbal_polis_idx.setdefault(k, set()).add(idx)

# FACUL
facul_slip_idx = {}
facul_polis_idx = {}
for idx, r in df_facul.iterrows():
    fac = str(r.get("FAC_CODE", "")).strip()
    if not fac or fac not in osbal_fac_map: continue
    for c in ["FAC_SLIP", "FAC_SLIP_CLEAN_1", "FAC_SLIP_CLEAN_2"]:
        for k in extract_slip_keys(r.get(c)):
            facul_slip_idx.setdefault(k, set()).add(fac)
    for c in ["FAC_POLICY_NO", "POLIS_CLEAN_1", "POLIS_CLEAN_2", "POLIS_CLEAN_3"]:
        for k in extract_kbru_polis_keys(r.get(c)):
            facul_polis_idx.setdefault(k, set()).add(fac)

# Test Suspend Matching order: SLIP -> POLIS -> INSURED
matched_slip = 0
matched_polis = 0
matched_insured = 0
unmatched = 0

for _, r in df_sus.iterrows():
    # 1. Slip
    hits = set()
    for sk in extract_slip_keys(r.get("SLIP_NO_ORI")):
        if sk in osbal_slip_idx:
            hits.update(osbal_slip_idx[sk])
        elif sk in facul_slip_idx:
            for fac in facul_slip_idx[sk]:
                hits.update(osbal_fac_map.get(fac, []))
    if hits:
        matched_slip += 1
        continue
        
    # 2. Polis
    for pk in extract_kbru_polis_keys(r.get("POLIS_ORI")):
        if pk in osbal_polis_idx:
            hits.update(osbal_polis_idx[pk])
        elif pk in facul_polis_idx:
            for fac in facul_polis_idx[pk]:
                hits.update(osbal_fac_map.get(fac, []))
    if hits:
        matched_polis += 1
        continue
        
    unmatched += 1

print(f"Matched by Slip (OSBAL + FACUL):  {matched_slip}")
print(f"Matched by Polis (OSBAL + FACUL): {matched_polis}")
print(f"Total matched by Slip or Polis:   {matched_slip + matched_polis}")
print(f"Remaining for Insured Fallback:   {unmatched}")
