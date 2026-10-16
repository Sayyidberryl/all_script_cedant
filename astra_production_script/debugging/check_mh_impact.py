import os
import sys

# Disable numexpr and bottleneck for NumPy 2.x compatibility
sys.modules['numexpr'] = None
sys.modules['bottleneck'] = None

# Ensure current script folder is on sys.path
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, BASE_DIR)

import pandas as pd
from prod_sus_1 import (
    _load_excel, _build_bordero_index, _find_fac_codes_in_bordero,
    _get_effective_sus_cols, _get_clean_cols, _build_lookup,
    _SUS_POLIS_PREFIX, _SUS_SLIP_PREFIX, _SUS_INSURED_PREFIX, _SUS_SERTIF_PREFIX,
    _OSBAL_POLIS_PREFIX, _OSBAL_SLIP_PREFIX, _OSBAL_INSURED_PREFIX, _OSBAL_SERTIF_PREFIX,
    OSBAL_FACODE_COL, FACUL_FACODE_COL,
    SUSPEND_FILE, OSBAL_FILE, FACUL_FILE,
    BORDERO_FIRE_FILE, BORDERO_MC_FILE, BORDERO_MH_FILE
)

print("--- ANALISIS IMPACT BORDERO MARINE HULL (MH) ---")

# 1. Load Bordero files
print("\n[1] Memeriksa file Bordero:")
for label, path in [("FIRE", BORDERO_FIRE_FILE), ("MC", BORDERO_MC_FILE), ("MH", BORDERO_MH_FILE)]:
    rows, cols = _load_excel(path, label, optional=True)
    df = pd.DataFrame(rows)
    facs = df["FAC CODE"].dropna().unique() if "FAC CODE" in df.columns else []
    polis = df["POLIS"].dropna().unique() if "POLIS" in df.columns else []
    print(f"  {label:<5}: {len(rows):,} baris, {len(facs)} unique FAC CODE, {len(polis):,} unique Polis")

# 2. Load Suspend, Osbal, Facul
print("\n[2] Memuat Suspend & OSBAL...")
suspend_rows, suspend_cols = _load_excel(SUSPEND_FILE, "SUSPEND")
osbal_rows, osbal_cols = _load_excel(OSBAL_FILE, "OSBAL")

polis_sus = _get_clean_cols(suspend_cols, _SUS_POLIS_PREFIX)
slip_sus = _get_clean_cols(suspend_cols, _SUS_SLIP_PREFIX)
insured_sus = _get_clean_cols(suspend_cols, _SUS_INSURED_PREFIX)
sertif_sus = _get_clean_cols(suspend_cols, _SUS_SERTIF_PREFIX)

polis_osbal = _get_clean_cols(osbal_cols, _OSBAL_POLIS_PREFIX)
slip_osbal = _get_clean_cols(osbal_cols, _OSBAL_SLIP_PREFIX)
insured_osbal = _get_clean_cols(osbal_cols, _OSBAL_INSURED_PREFIX)

lookup_osbal = _build_lookup(osbal_rows, polis_osbal, slip_osbal, insured_osbal, OSBAL_FACODE_COL, label="OSBAL")
facode_osbal_idx = lookup_osbal[9]

# 3. Build index per file
def get_bordero_item(path, name):
    rows, _ = _load_excel(path, name, optional=True)
    if not rows:
        return None
    idx = _build_bordero_index(rows)
    polis_idx = {}
    for fac, r_list in idx.items():
        for r in r_list:
            p = r.get("polis")
            if p and len(p) >= 5:
                polis_idx.setdefault(p, []).append((fac, r))
    return {"name": name, "rows": rows, "idx": idx, "polis_idx": polis_idx}

b_fire = get_bordero_item(BORDERO_FIRE_FILE, "FIRE")
b_mc = get_bordero_item(BORDERO_MC_FILE, "MC")
b_mh = get_bordero_item(BORDERO_MH_FILE, "MH")

# 4. Cek berapa banyak baris Suspend yang match ke MH secara mandiri (standalone)
print("\n[3] Pengecekan Matching Suspend ke Bordero MH secara Mandiri:")
mh_hits = []
for idx, sus in enumerate(suspend_rows):
    eff_p, eff_s, _, _ = _get_effective_sus_cols(sus, polis_sus, slip_sus)
    bordero_facs, ok = _find_fac_codes_in_bordero(
        sus, b_mh["idx"], eff_p, eff_s,
        bordero_polis_index=b_mh["polis_idx"],
        sertif_sus_cols=sertif_sus
    )
    if ok and bordero_facs:
        # resolve to osbal
        res_osbal = [f for f in bordero_facs if f in facode_osbal_idx]
        mh_hits.append({
            "sus_idx": idx,
            "sus": sus,
            "mh_facs": bordero_facs,
            "osbal_facs": res_osbal,
            "resolved": len(res_osbal) > 0
        })

print(f"  Total baris Suspend yang cocok ke MH: {len(mh_hits)}")
resolved_mh = [h for h in mh_hits if h["resolved"]]
print(f"  Dari {len(mh_hits)} baris tersebut, yang FAC CODE-nya ada di OSBAL: {len(resolved_mh)}")

# 5. Cek tumpang tindih: Apakah baris yang cocok ke MH juga cocok ke FIRE atau MC?
overlap_fire = 0
overlap_mc = 0
only_mh = 0
only_mh_resolved = 0

for h in mh_hits:
    sus = h["sus"]
    eff_p, eff_s, _, _ = _get_effective_sus_cols(sus, polis_sus, slip_sus)
    facs_fire, ok_fire = _find_fac_codes_in_bordero(sus, b_fire["idx"], eff_p, eff_s, bordero_polis_index=b_fire["polis_idx"], sertif_sus_cols=sertif_sus)
    facs_mc, ok_mc = _find_fac_codes_in_bordero(sus, b_mc["idx"], eff_p, eff_s, bordero_polis_index=b_mc["polis_idx"], sertif_sus_cols=sertif_sus)
    
    in_fire = ok_fire and bool(facs_fire)
    in_mc = ok_mc and bool(facs_mc)
    
    if in_fire: overlap_fire += 1
    if in_mc: overlap_mc += 1
    if not in_fire and not in_mc:
        only_mh += 1
        if h["resolved"]:
            only_mh_resolved += 1

print(f"\n[4] Tumpang tindih (Overlap) Bordero MH dengan FIRE & MC:")
print(f"  Cocok juga di FIRE        : {overlap_fire}")
print(f"  Cocok juga di MC          : {overlap_mc}")
print(f"  HANYA cocok di MH         : {only_mh} baris")
print(f"  HANYA cocok di MH & OSBAL : {only_mh_resolved} baris")

# 6. Lihat detail baris unik MH
print(f"\n[5] Detail Baris Suspend yang HANYA Cocok di MH:")
count_shown = 0
for h in mh_hits:
    sus = h["sus"]
    eff_p, eff_s, _, _ = _get_effective_sus_cols(sus, polis_sus, slip_sus)
    facs_fire, ok_fire = _find_fac_codes_in_bordero(sus, b_fire["idx"], eff_p, eff_s, bordero_polis_index=b_fire["polis_idx"], sertif_sus_cols=sertif_sus)
    facs_mc, ok_mc = _find_fac_codes_in_bordero(sus, b_mc["idx"], eff_p, eff_s, bordero_polis_index=b_mc["polis_idx"], sertif_sus_cols=sertif_sus)
    if not (ok_fire and facs_fire) and not (ok_mc and facs_mc):
        count_shown += 1
        print(f"  {count_shown}. Row {h['sus_idx']}: Polis={sus.get('POLIS')}, Slip={sus.get('SLIP NO')}, Amount={sus.get('AMOUNT ORI')}, Curr={sus.get('CURR ORI')}")
        print(f"     MH FAC: {h['mh_facs']} -> OSBAL FAC: {h['osbal_facs']}")
        if count_shown >= 20:
            break
