import os, sys, re
sys.modules['numexpr'] = None
sys.modules['bottleneck'] = None
import pandas as pd

DATA_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "data")
df_sus = pd.read_excel(os.path.join(DATA_DIR, "kbru_output_suspend.xlsx"))
df_osbal = pd.read_excel(os.path.join(DATA_DIR, "kbru_output_osbal.xlsx"))

# Fungsi ekstraksi core token polis KBRU
# Contoh polis: 00003136/FIAR/038.0000007-KBRURE/07/2024 -> doc: 3136, class: FIAR, client: 038.0000007
def extract_kbru_polis_keys(polis_str):
    if pd.isna(polis_str): return []
    s = str(polis_str).strip().upper()
    keys = []
    
    # 1. Alphanumeric polos
    c = re.sub(r'[^A-Z0-9]', '', s)
    if len(c) >= 6:
        keys.append(c)
        
    # 2. Pola KBRU: (\d+)[/.\s]+([A-Z]+)[/.\s]+(\d+)[/.\s]+(\d+)
    # Misal 00003136/FIAR/038.0000007
    m = re.search(r'0*(\d{3,8})[/.\s]+([A-Z]{3,5})[/.\s]+0*(\d+)[/.\s]+0*(\d+)', s)
    if m:
        core_num, cls, p3, p4 = m.groups()
        # Key unik: core number + class (misal 3136FIAR387)
        keys.append(f"{core_num}{cls}{p3}{p4}")
        keys.append(f"{core_num}{cls}")
        keys.append(f"{core_num}")
    else:
        # Coba ambil nomor dokumen saja jika diawali 0000...
        m2 = re.search(r'0*(\d{4,8})', s)
        if m2:
            keys.append(m2.group(1))
            
    return keys

# Ekstraksi keys dari OSBAL
# Di OSBAL, cek kolom FAC_SLIP_ORI, POLIS_CLEAN_*, CLSDT_POLICY_NO
osbal_polis_to_fac = {}
for idx, r in df_osbal.iterrows():
    fac = str(r.get("CCOS_REF_CODE", "")).strip()
    if not fac: continue
    
    # Cek seluruh field OSBAL yang mungkin berisi nomor dokumen KBRU
    for col in ["FAC_SLIP_ORI", "POLIS_CLEAN_1", "POLIS_CLEAN_2", "CLSDT_POLICY_NO", "CLSDT_SLIP_NO"]:
        val = r.get(col)
        if pd.isna(val): continue
        val_str = str(val).strip().upper()
        if val_str in ["-", "NONE", "TBA", ""]: continue
        
        # Split jika ada tanda '+'
        subparts = [p.strip() for p in val_str.split('+')]
        for part in subparts:
            keys = extract_kbru_polis_keys(part)
            for k in keys:
                osbal_polis_to_fac.setdefault(k, set()).add(fac)

# Uji terhadap 927 baris 'Insured only (> 1 Fac)'
v1_path = os.path.join(DATA_DIR, "final_output_v1.xlsx")
df_v1 = pd.read_excel(v1_path)
df_ins_gt1 = df_v1[df_v1["SKENARIO"] == "Insured only (> 1 Fac)"]

resolved_single = 0
resolved_multi = 0
not_resolved = 0

for idx, r in df_ins_gt1.iterrows():
    p_ori = r.get("POLIS")
    keys = extract_kbru_polis_keys(p_ori)
    
    # Cari yang match
    hit_facs = set()
    # Prioritaskan key yang paling spesifik (panjang >= 6)
    specific_keys = [k for k in keys if len(k) >= 6]
    for k in specific_keys:
        if k in osbal_polis_to_fac:
            hit_facs.update(osbal_polis_to_fac[k])
            break
            
    if len(hit_facs) == 1:
        resolved_single += 1
    elif len(hit_facs) > 1:
        resolved_multi += 1
    else:
        not_resolved += 1

print("--- HASIL SIMULASI DENGAN ADVANCED POLIS KEY ---")
print(f"Dari 927 baris 'Insured only (> 1 Fac)':")
print(f"  - Berhasil di-resolve ke 1 FAC CODE tunggal : {resolved_single} baris ({resolved_single/927*100:.1f}%)")
print(f"  - Match ke >1 FAC Code spesifik            : {resolved_multi} baris ({resolved_multi/927*100:.1f}%)")
print(f"  - Tetap belum resolve (butuh insured/slip) : {not_resolved} baris ({not_resolved/927*100:.1f}%)")
