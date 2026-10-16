import os, sys, re, time
sys.modules['numexpr'] = None
sys.modules['bottleneck'] = None
import pandas as pd
import numpy as np

DATA_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "data")
SUSPEND_FILE = os.path.join(DATA_DIR, "kbru_output_suspend.xlsx")
OSBAL_FILE   = os.path.join(DATA_DIR, "kbru_output_osbal.xlsx")
FACUL_FILE   = os.path.join(DATA_DIR, "kbru_output_facul.xlsx")

df_sus = pd.read_excel(SUSPEND_FILE)
df_osbal = pd.read_excel(OSBAL_FILE)
df_facul = pd.read_excel(FACUL_FILE)

# 1. Definisi Token Dummy yang Ditolak
DUMMY_VALUES = {"-", "--", "---", "0", "00", "000", "NONE", "NULL", "TBA", "NA", "N/A", "SELISIH", "NO", "NON", ""}

def is_valid_token(s: str) -> bool:
    if not s: return False
    cleaned = str(s).strip().upper()
    if cleaned in DUMMY_VALUES: return False
    # Harus punya minimal 3 karakter alfanumerik
    alphanumeric = re.sub(r'[^A-Z0-9]', '', cleaned)
    return len(alphanumeric) >= 3

def clean_alphanumeric(s: str) -> str:
    if not s: return ""
    return re.sub(r'[^A-Z0-9]', '', str(s).upper())

# Ekstraksi variasi polis KBRU
def extract_kbru_variants(s: str) -> list:
    if not is_valid_token(s): return []
    text = str(s).strip().upper()
    variants = set()
    
    # 1. String asli bersih
    variants.add(text)
    c_all = clean_alphanumeric(text)
    if len(c_all) >= 3:
        variants.add(c_all)
        
    # 2. Tangani tanda '+' jika multi-polis (misal di OSBAL: '3136 + 3143 + 6702')
    if '+' in text:
        parts = [p.strip() for p in text.split('+')]
        for p in parts:
            if is_valid_token(p):
                variants.update(extract_kbru_variants(p))
                
    # 3. Tangani pola nomor registrasi KBRU: 00003136/FIAR/038.0000007-KBRURE/07/2024
    # Hilangkan leading zero pada nomor dokumen depan: 00003136 -> 3136
    m = re.search(r'0*(\d{3,8})[/.\s]+([A-Z]{3,5})[/.\s]+0*(\d+)[/.\s]+0*(\d+)', text)
    if m:
        doc_num, cls, p3, p4 = m.groups()
        variants.add(f"{doc_num}{cls}{p3}{p4}")
        variants.add(f"{doc_num}{cls}")
        variants.add(doc_num)
        # Juga bentuk stripped KBRURE/SRB
        # Hilangkan kata KBRURE / SRB
        no_broker = re.sub(r'KBRURE|SRB', '', c_all)
        if len(no_broker) >= 6:
            variants.add(no_broker)
            
    # 4. Tangani base slip (strip -01, -02)
    m_base = re.sub(r'-\d+$', '', text)
    if m_base != text and is_valid_token(m_base):
        variants.add(m_base)
        variants.add(clean_alphanumeric(m_base))
        
    return [v for v in variants if len(v) >= 3 and v not in DUMMY_VALUES]

# Bangun Index OSBAL
# KBRU Khusus: Kolom FAC_SLIP_ORI di OSBAL berisi Slip MAUPUN Polis!
osbal_exact_slip = {}
osbal_exact_polis = {}
osbal_exact_insured = {}
osbal_fac_map = {}

for idx, r in df_osbal.iterrows():
    fac = str(r.get("CCOS_REF_CODE", "")).strip()
    if not fac: continue
    osbal_fac_map.setdefault(fac, []).append(idx)
    
    # SLIP columns (termasuk CLSDT_SLIP_NO, FAC_SLIP_ORI)
    for c in ["CLSDT_SLIP_NO", "FAC_SLIP_ORI", "SLIP_CLEAN_1", "SLIP_CLEAN_2"]:
        v = r.get(c)
        for variant in extract_kbru_variants(v):
            osbal_exact_slip.setdefault(variant, set()).add(idx)
            
    # POLIS columns (termasuk CLSDT_POLICY_NO, POLIS_CLEAN_1, dan juga FAC_SLIP_ORI karena format polis KBRU disimpan di situ!)
    for c in ["CLSDT_POLICY_NO", "FAC_POLICY_NO_ORI", "POLIS_CLEAN_1", "POLIS_CLEAN_2", "FAC_SLIP_ORI"]:
        v = r.get(c)
        for variant in extract_kbru_variants(v):
            osbal_exact_polis.setdefault(variant, set()).add(idx)
            
    # INSURED columns
    for c in ["FAC_INSURED_ORI", "INSURED_CLEAN_1", "INSURED_CLEAN_2", "INSURED_CLEAN_3"]:
        v = r.get(c)
        if is_valid_token(v):
            norm_ins = re.sub(r'\s+', ' ', str(v).strip().upper())
            osbal_exact_insured.setdefault(norm_ins, set()).add(idx)

# Bangun Index FACUL -> OSBAL
facul_exact_slip = {}
facul_exact_polis = {}
facul_exact_insured = {}

for idx, r in df_facul.iterrows():
    fac = str(r.get("FAC_CODE", "")).strip()
    if not fac or fac not in osbal_fac_map: continue
    
    for c in ["FAC_SLIP", "FAC_SLIP_CLEAN_1", "FAC_SLIP_CLEAN_2"]:
        v = r.get(c)
        for variant in extract_kbru_variants(v):
            facul_exact_slip.setdefault(variant, set()).add(fac)
            
    for c in ["FAC_POLICY_NO", "POLIS_CLEAN_1", "POLIS_CLEAN_2", "POLIS_CLEAN_3"]:
        v = r.get(c)
        for variant in extract_kbru_variants(v):
            facul_exact_polis.setdefault(variant, set()).add(fac)
            
    for c in ["FAC_INSURED_ORI", "FAC_INSURED_1", "FAC_INSURED_2"]:
        v = r.get(c)
        if is_valid_token(v):
            norm_ins = re.sub(r'\s+', ' ', str(v).strip().upper())
            facul_exact_insured.setdefault(norm_ins, set()).add(fac)

# Run Simulation on Suspend
results = []
for idx, r in df_sus.iterrows():
    matched_idx = set()
    scenario = ""
    source = ""
    
    # 1. Slip Matching
    s_val = r.get("SLIP_NO_ORI")
    s_variants = extract_kbru_variants(s_val)
    for sv in s_variants:
        if sv in osbal_exact_slip:
            matched_idx.update(osbal_exact_slip[sv])
            source = "OSBAL"
            scenario = "Slip only"
            break
        elif sv in facul_exact_slip:
            for fc in facul_exact_slip[sv]:
                matched_idx.update(osbal_fac_map.get(fc, []))
            if matched_idx:
                source = "FACUL"
                scenario = "Slip only"
                break
                
    # 2. Polis Matching (jika belum match slip)
    if not matched_idx:
        p_val = r.get("POLIS_ORI")
        p_variants = extract_kbru_variants(p_val)
        for pv in p_variants:
            if pv in osbal_exact_polis:
                matched_idx.update(osbal_exact_polis[pv])
                source = "OSBAL"
                scenario = "Polis only"
                break
            elif pv in facul_exact_polis:
                for fc in facul_exact_polis[pv]:
                    matched_idx.update(osbal_fac_map.get(fc, []))
                if matched_idx:
                    source = "FACUL"
                    scenario = "Polis only"
                    break

    # 3. Insured Matching (Fallback)
    if not matched_idx:
        i_val = r.get("INSURED_ORI")
        if is_valid_token(i_val):
            norm_i = re.sub(r'\s+', ' ', str(i_val).strip().upper())
            if norm_i in osbal_exact_insured:
                matched_idx.update(osbal_exact_insured[norm_i])
                source = "OSBAL"
                scenario = "Insured only"
            elif norm_i in facul_exact_insured:
                for fc in facul_exact_insured[norm_i]:
                    matched_idx.update(osbal_fac_map.get(fc, []))
                if matched_idx:
                    source = "FACUL"
                    scenario = "Insured only"

    # Evaluasi FAC Codes
    fac_codes = {str(df_osbal.iloc[i].get("CCOS_REF_CODE", "")).strip() for i in matched_idx}
    num_fac = len(fac_codes)
    
    if not matched_idx or not source:
        final_scen = "Unmatching"
        flag = "Unmatching"
    elif num_fac > 1:
        final_scen = f"{scenario} (> 1 Fac)"
        flag = "Matching >1 fac code"
    else:
        final_scen = scenario
        flag = "Single FAC Matched"
        
    results.append({
        "scenario": final_scen,
        "flag": flag,
        "num_fac": num_fac,
        "source": source
    })

df_res = pd.DataFrame(results)
print("=== PERBANDINGAN SEBELUM VS SESUDAH REKOMENDASI ===")
print("\nHASIL SEBELUM REKOMENDASI (BASELINE):")
print("  Matching >1 fac code : 1,000")
print("  Single FAC Matched   :   504 (Adjusment + New Entry + Beda Curr/Per)")
print("  Unmatching           :   329")

print("\nHASIL SIMULASI DENGAN REKOMENDASI IMPLEMENTASI:")
print(df_res["flag"].value_counts())
print("\nRincian Skenario Baru:")
print(df_res["scenario"].value_counts())
