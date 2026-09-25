import os
import re

import numpy as np
import pandas as pd


# ─────────────────────────────────────────────────────────────────────────────
# PATHS & CONSTANTS
# ─────────────────────────────────────────────────────────────────────────────

SUSPEND_FILE = os.path.join("data", "suspend_clean_aca.xlsx")
OSBAL_FILE   = os.path.join("data", "osbal_clean_aca.xlsx")
FACUL_FILE   = os.path.join("data", "facul_clean_aca.xlsx")
OUTPUT_FILE  = os.path.join("data", "final_output.xlsx")

OSBAL_FACODE_COL = "CCOS_REF_CODE"
FACUL_FACODE_COL = "FAC_CODE"

FACODE_JOIN_SEP = " | "  

FINAL_COLUMNS = [
    "CCOS_DOC_NO", "CCOS_REF_CODE",
    "RECEIPT NO", "CREDIT NOTES", "DETAIL RINCIAN NO", "RECEIPT DATE",
    "CEDANT NAME", "CEDANT SHRT NAME",
    "INSURED_ORI", "INSURED_1", "INSURED_2",
    "CURR ORI", "AMOUNT ORI", "AMOUNT_ORI_MIN1", "CURR PAY", "AMOUNT PAY",
    "CCOS_OR_BAL", "CCOS_BAL_DUE", "DIFERENCE", "FLAG_PROD",
    "POLIS_ORI", "POLIS_CLN", "SLIP_NO_ORI", "SLIP_NO_CLN",
    "DESC 1", "DESC 2", "DESC 3", "DESC 4", "STATUS", "REC_TYPE",
    "CEK AMOUNT DATABASE X BAL RV",
    "Mark Admin Fac Code", "Mark Admin", "Mark Admin (Status)", "Mark ARP",
    "SKENARIO",
]

# Skenario matching berurutan: (label, kolom_prefix_clean, kolom_ori)
SCENARIOS = [
    ("SLIP_CLEAN",    "clean slip",    "slip_ori"),
    ("POLIS_CLEAN",   "clean polis",   "polis_ori"),
    ("INSURED_CLEAN", "clean insured", "insured_ori"),
]
# Skenario "ORI" dipisah karena cuma pakai 1 kolom (tidak butuh explode)
ORI_SCENARIOS = [
    ("SLIP_ORI",    "slip_ori"),
    ("POLIS_ORI",   "polis_ori"),
    ("INSURED_ORI", "insured_ori"),
]


# ─────────────────────────────────────────────────────────────────────────────
# NORMALISASI
# ─────────────────────────────────────────────────────────────────────────────

def _norm(val) -> str:
    if pd.isna(val) or val is None:
        return ""
    return re.sub(r"\s+", " ", str(val).strip().upper())


def _get_clean_cols(df: pd.DataFrame, prefix: str) -> list:
    return [c for c in df.columns if c.lower().startswith(prefix.lower())]


# ─────────────────────────────────────────────────────────────────────────────
# EXPLODE HELPER (khusus untuk MATCHING, bukan untuk output)
# ─────────────────────────────────────────────────────────────────────────────

def _build_match_keys(df: pd.DataFrame, clean_cols: list) -> pd.DataFrame:
    """
    Bangun tabel panjang (long) berisi (row_index, key_value) dari kolom-kolom
    clean_cols. Nilai yang merupakan gabungan koma (>5 item) di-split dulu.
    Ini padanan dari "explode" -- setara PROC TRANSPOSE + split di SAS.
    """
    frames = []
    for col in clean_cols:
        s = df[col].map(_norm)
        s = s[s != ""]
        if s.empty:
            continue
        # split by comma untuk menangani nilai gabungan (>5 item / row)
        exploded = s.str.split(",").explode()
        exploded = exploded.map(lambda x: x.strip())
        exploded = exploded[exploded != ""]
        frames.append(pd.DataFrame({"row_idx": exploded.index, "key": exploded.values}))
    if not frames:
        return pd.DataFrame(columns=["row_idx", "key"])
    out = pd.concat(frames, ignore_index=True).drop_duplicates()
    return out


def _build_ori_keys(df: pd.DataFrame, ori_col: str) -> pd.DataFrame:
    s = df[ori_col].map(_norm)
    s = s[s != ""]
    return pd.DataFrame({"row_idx": s.index, "key": s.values}).drop_duplicates()


# ─────────────────────────────────────────────────────────────────────────────
# NARROWING (verifikasi silang polis/slip pada kandidat hasil merge)
# ─────────────────────────────────────────────────────────────────────────────

def _row_key_set(df: pd.DataFrame, row_idx, clean_cols: list, ori_col: str) -> set:
    """Kumpulkan semua nilai (clean, sudah di-split koma + ori) untuk satu baris."""
    vals = set()
    row = df.loc[row_idx]
    for col in clean_cols:
        v = _norm(row.get(col, ""))
        if v:
            vals.update(x.strip() for x in v.split(",") if x.strip())
    ori = _norm(row.get(ori_col, ""))
    if ori:
        vals.add(ori)
    return vals


def _narrow_pairs(
    pairs: pd.DataFrame,          # kolom: susp_idx, ref_idx
    skenario_label: str,
    df_sus: pd.DataFrame, df_ref: pd.DataFrame,
    sus_polis_cols: list, sus_slip_cols: list,
    ref_polis_cols: list, ref_slip_cols: list,
) -> pd.DataFrame:
    """
    Untuk tiap susp_idx yang match ke >1 ref_idx, sempitkan pakai verifikasi
    silang POLIS/SLIP. Kalau hasil narrowing kosong -> kembalikan pasangan asal
    (persis semantik _narrow_by_pair versi lama).
    """
    if pairs.empty:
        return pairs

    counts = pairs.groupby("susp_idx")["ref_idx"].transform("count")
    single = pairs[counts <= 1]
    multi  = pairs[counts > 1]

    if multi.empty:
        return pairs

    keep_rows = []
    for susp_idx, grp in multi.groupby("susp_idx"):
        sus_polis = _row_key_set(df_sus, susp_idx, sus_polis_cols, "polis_ori")
        sus_slip  = _row_key_set(df_sus, susp_idx, sus_slip_cols,  "slip_ori")

        narrowed = []
        for ref_idx in grp["ref_idx"]:
            ref_polis = _row_key_set(df_ref, ref_idx, ref_polis_cols, "polis_ori")
            ref_slip  = _row_key_set(df_ref, ref_idx, ref_slip_cols,  "slip_ori")
            polis_ok = bool(sus_polis & ref_polis)
            slip_ok  = bool(sus_slip & ref_slip)

            if "SLIP" in skenario_label:
                ok = polis_ok
            elif "POLIS" in skenario_label:
                ok = slip_ok
            else:  # INSURED
                ok = polis_ok or slip_ok

            if ok:
                narrowed.append(ref_idx)

        final_refs = narrowed if narrowed else list(grp["ref_idx"])
        for ref_idx in final_refs:
            keep_rows.append((susp_idx, ref_idx))

    narrowed_multi = pd.DataFrame(keep_rows, columns=["susp_idx", "ref_idx"])
    return pd.concat([single, narrowed_multi], ignore_index=True)


# ─────────────────────────────────────────────────────────────────────────────
# CASCADE MATCH — satu sumber referensi (OSBAL atau FACUL)
# ─────────────────────────────────────────────────────────────────────────────

def _match_against_source(
    df_sus: pd.DataFrame,
    remaining_idx: pd.Index,
    df_ref: pd.DataFrame,
    source_label: str,
) -> tuple:
    """
    Cascade 6 skenario terhadap satu sumber referensi.
    Returns:
      matched_pairs: DataFrame [susp_idx, ref_idx, skenario]
      leftover_idx:  Index baris suspend yang masih belum matched
    """
    sus_polis_cols   = _get_clean_cols(df_sus, "clean polis")
    sus_slip_cols    = _get_clean_cols(df_sus, "clean slip")
    sus_insured_cols = _get_clean_cols(df_sus, "clean insured")

    ref_polis_cols   = _get_clean_cols(df_ref, "clean polis")
    ref_slip_cols    = _get_clean_cols(df_ref, "clean slip")
    ref_insured_cols = _get_clean_cols(df_ref, "clean insured")

    sus_cols_map = {"clean slip": sus_slip_cols, "clean polis": sus_polis_cols, "clean insured": sus_insured_cols}
    ref_cols_map = {"clean slip": ref_slip_cols, "clean polis": ref_polis_cols, "clean insured": ref_insured_cols}

    all_matched = []
    leftover = remaining_idx

    # --- 3 skenario "clean" (dengan explode koma) ---
    for label, prefix, _ori_col in SCENARIOS:
        if leftover.empty:
            break
        sus_keys = _build_match_keys(df_sus.loc[leftover], sus_cols_map[prefix])
        ref_keys = _build_match_keys(df_ref, ref_cols_map[prefix])
        if sus_keys.empty or ref_keys.empty:
            continue

        merged = sus_keys.merge(ref_keys, on="key", suffixes=("_sus", "_ref"))
        if merged.empty:
            continue

        pairs = merged.rename(columns={"row_idx_sus": "susp_idx", "row_idx_ref": "ref_idx"})[["susp_idx", "ref_idx"]].drop_duplicates()
        pairs = _narrow_pairs(pairs, label, df_sus, df_ref, sus_polis_cols, sus_slip_cols, ref_polis_cols, ref_slip_cols)
        pairs["skenario"] = f"{source_label}_{label}"
        all_matched.append(pairs)

        matched_susp = pd.Index(pairs["susp_idx"].unique())
        leftover = leftover.difference(matched_susp)

    # --- 3 skenario "ori" (1 kolom, tanpa explode) ---
    for label, ori_col in ORI_SCENARIOS:
        if leftover.empty:
            break
        sus_keys = _build_ori_keys(df_sus.loc[leftover], ori_col)
        ref_keys = _build_ori_keys(df_ref, ori_col)
        if sus_keys.empty or ref_keys.empty:
            continue

        merged = sus_keys.merge(ref_keys, on="key", suffixes=("_sus", "_ref"))
        if merged.empty:
            continue

        pairs = merged.rename(columns={"row_idx_sus": "susp_idx", "row_idx_ref": "ref_idx"})[["susp_idx", "ref_idx"]].drop_duplicates()
        pairs["skenario"] = f"{source_label}_{label}"
        all_matched.append(pairs)

        matched_susp = pd.Index(pairs["susp_idx"].unique())
        leftover = leftover.difference(matched_susp)

    if all_matched:
        result = pd.concat(all_matched, ignore_index=True)
    else:
        result = pd.DataFrame(columns=["susp_idx", "ref_idx", "skenario"])

    return result, leftover


# ─────────────────────────────────────────────────────────────────────────────
# AGREGASI MULTI-MATCH (fac_code_apn style) + KOLOM TURUNAN
# ─────────────────────────────────────────────────────────────────────────────

def _aggregate_matches(pairs: pd.DataFrame, df_ref: pd.DataFrame, facode_col: str, source_label: str) -> pd.DataFrame:
    """
    Untuk tiap susp_idx: gabungkan semua fac_code (dipisah ' | '), jumlahkan
    CCOS_BAL_DUE / CCOS_OR_BAL, dan flag kalau >1 fac_code berbeda.
    Analog langsung ke bagian fac_code_apn di SAS.
    """
    if pairs.empty:
        return pd.DataFrame(columns=[
            "susp_idx", "skenario", "CCOS_DOC_NO", "CCOS_REF_CODE",
            "CCOS_OR_BAL", "CCOS_BAL_DUE", "Mark Admin Fac Code",
        ])

    ref = df_ref.loc[pairs["ref_idx"]].reset_index(drop=True)
    p = pairs.reset_index(drop=True)
    joined = pd.concat([p, ref], axis=1)

    has_ccos = "CCOS_BAL_DUE" in df_ref.columns

    def agg_group(g: pd.DataFrame) -> pd.Series:
        facodes_unique = list(dict.fromkeys(str(x).strip() for x in g[facode_col] if str(x).strip()))
        facode_apn = FACODE_JOIN_SEP.join(facodes_unique)
        mark = "facode lebih dari 1" if len(facodes_unique) > 1 else (facodes_unique[0] if facodes_unique else "")

        out = {
            "skenario": g["skenario"].iloc[0],
            "Mark Admin Fac Code": mark,
            "CCOS_DOC_NO": ", ".join(filter(None, (str(x) for x in g.get("CCOS_DOC_NO", [])))) if "CCOS_DOC_NO" in g else "",
            "CCOS_REF_CODE": facode_apn if source_label == "OSBAL" else "",
        }
        if has_ccos:
            out["CCOS_OR_BAL"] = pd.to_numeric(g.get("CCOS_OR_BAL", 0), errors="coerce").fillna(0).sum()
            out["CCOS_BAL_DUE"] = pd.to_numeric(g.get("CCOS_BAL_DUE", 0), errors="coerce").fillna(0).sum()
        else:
            out["CCOS_OR_BAL"] = np.nan
            out["CCOS_BAL_DUE"] = np.nan
        return pd.Series(out)

    result = joined.groupby("susp_idx", group_keys=False).apply(agg_group).reset_index()
    return result


# ─────────────────────────────────────────────────────────────────────────────
# KOLOM TURUNAN
# ─────────────────────────────────────────────────────────────────────────────

def _compute_derived_columns(df: pd.DataFrame) -> pd.DataFrame:
    amt_ori  = pd.to_numeric(df.get("AMOUNT ORI",   0), errors="coerce").fillna(0)
    bal_due  = pd.to_numeric(df.get("CCOS_BAL_DUE", 0), errors="coerce").fillna(0)
    facode_s = df.get("Mark Admin Fac Code", pd.Series("", index=df.index)).fillna("")

    df["AMOUNT_ORI_MIN1"] = amt_ori * -1
    df["DIFERENCE"]       = df["AMOUNT_ORI_MIN1"] - bal_due

    a_min1 = df["AMOUNT_ORI_MIN1"]
    df["FLAG_PROD"] = np.select(
        [
            facode_s.str.strip() == "facode lebih dari 1",
            (a_min1 == bal_due) | (a_min1 < bal_due),
            (a_min1 > bal_due)  | ((a_min1 != 0) & (bal_due == 0)),
        ],
        ["Matching >1 fac code", "Adjustment", "New Entry"],
        default="Unmatching",
    )
    return df


# ─────────────────────────────────────────────────────────────────────────────
# LOAD
# ─────────────────────────────────────────────────────────────────────────────

def _load_excel(path: str, label: str) -> pd.DataFrame:
    print(f"  Membaca {label}: {path} ...")
    df = pd.read_excel(path)
    print(f"  -> {len(df):,} baris, {len(df.columns)} kolom")
    return df


# ─────────────────────────────────────────────────────────────────────────────
# BUILD OUTPUT ROWS
# ─────────────────────────────────────────────────────────────────────────────

def _build_output(df_sus: pd.DataFrame, match_info: pd.DataFrame) -> pd.DataFrame:
    """
    match_info: index = susp_idx (subset dari df_sus), berisi skenario,
    CCOS_DOC_NO, CCOS_REF_CODE, CCOS_OR_BAL, CCOS_BAL_DUE, Mark Admin Fac Code.
    Baris suspend yang tidak ada di match_info -> UNMATCHED.
    """
    df = df_sus.copy()
    df = df.join(match_info, how="left")

    df["skenario"] = df["skenario"].fillna("UNMATCHED")
    for c in ["CCOS_DOC_NO", "CCOS_REF_CODE", "Mark Admin Fac Code"]:
        if c not in df.columns:
            df[c] = ""
        df[c] = df[c].fillna("")
    for c in ["CCOS_OR_BAL", "CCOS_BAL_DUE"]:
        if c not in df.columns:
            df[c] = np.nan

    out = pd.DataFrame({
        "CCOS_DOC_NO":  df["CCOS_DOC_NO"],
        "CCOS_REF_CODE": df["CCOS_REF_CODE"],
        "RECEIPT NO":        df.get("RECEIPT NO", ""),
        "CREDIT NOTES":      df.get("CREDIT NOTES", ""),
        "DETAIL RINCIAN NO": df.get("DETAIL RINCIAN NO", ""),
        "RECEIPT DATE":      df.get("RECEIPT DATE", ""),
        "CEDANT NAME":       df.get("CEDANT NAME", ""),
        "CEDANT SHRT NAME":  df.get("CEDANT SHRT NAME", ""),
        "INSURED_ORI": df.get("insured_ori", ""),
        "INSURED_1":   df.get("clean insured 1", ""),
        "INSURED_2":   df.get("clean insured 2", ""),
        "CURR ORI":   df.get("CURR ORI", ""),
        "AMOUNT ORI": df.get("AMOUNT ORI", ""),
        "CURR PAY":   df.get("CURR PAY", ""),
        "AMOUNT PAY": df.get("AMOUNT PAY", ""),
        "CCOS_OR_BAL":  df["CCOS_OR_BAL"],
        "CCOS_BAL_DUE": df["CCOS_BAL_DUE"],
        "POLIS_ORI":   df.get("polis_ori", ""),
        "POLIS_CLN":   df.get("clean polis 1", ""),
        "SLIP_NO_ORI": df.get("slip_ori", ""),
        "SLIP_NO_CLN": df.get("clean slip 1", ""),
        "DESC 1":  df.get("DESC 1", ""),
        "DESC 2":  df.get("DESC 2", ""),
        "DESC 3":  df.get("DESC 3", ""),
        "DESC 4":  df.get("DESC 4", ""),
        "STATUS":   df.get("STATUS", ""),
        "REC_TYPE": df.get("REC_TYPE", ""),
        "CEK AMOUNT DATABASE X BAL RV": "",
        "Mark Admin Fac Code": df["Mark Admin Fac Code"],
        "Mark Admin": "",
        "Mark Admin (Status)": "",
        "Mark ARP": "",
        "SKENARIO": df["skenario"],
    })
    return out.reset_index(drop=True)


# ─────────────────────────────────────────────────────────────────────────────
# MAIN PROCESS
# ─────────────────────────────────────────────────────────────────────────────

def run() -> None:
    print("\n" + "=" * 60)
    print("  PRODUCTION SCRIPT - MATCHING & JOIN (cascade-merge)")
    print("=" * 60)

    print("\n[1/5] Loading data cleaned ...")
    df_sus   = _load_excel(SUSPEND_FILE, "SUSPEND")
    df_osbal = _load_excel(OSBAL_FILE,   "OSBAL")
    df_facul = _load_excel(FACUL_FILE,   "FACUL")

    print(f"\n[2/5] Matching {len(df_sus):,} baris suspend -> OSBAL ...")
    remaining = df_sus.index
    pairs_osbal, remaining = _match_against_source(df_sus, remaining, df_osbal, "OSBAL")
    print(f"  Matched ke OSBAL : {pairs_osbal['susp_idx'].nunique() if not pairs_osbal.empty else 0:,}")
    print(f"  Sisa (belum match): {len(remaining):,}")

    print(f"\n[3/5] Matching sisa -> FACUL ...")
    pairs_facul, remaining = _match_against_source(df_sus, remaining, df_facul, "FACUL")
    print(f"  Matched ke FACUL : {pairs_facul['susp_idx'].nunique() if not pairs_facul.empty else 0:,}")
    print(f"  Sisa akhir (UNMATCHED): {len(remaining):,}")

    print("\n[4/5] Agregasi hasil match (fac_code gabung, CCOS dijumlah) ...")
    agg_osbal = _aggregate_matches(pairs_osbal, df_osbal, OSBAL_FACODE_COL, "OSBAL")
    agg_facul = _aggregate_matches(pairs_facul, df_facul, FACUL_FACODE_COL, "FACUL")

    agg_all = pd.concat([agg_osbal, agg_facul], ignore_index=True).set_index("susp_idx")

    df_out = _build_output(df_sus, agg_all)

    for col in ["AMOUNT ORI", "CCOS_OR_BAL", "CCOS_BAL_DUE"]:
        df_out[col] = pd.to_numeric(df_out[col], errors="coerce")

    df_out = _compute_derived_columns(df_out)
    df_out.loc[df_out["SKENARIO"] == "UNMATCHED", "FLAG_PROD"] = "Unmatching"

    for col in FINAL_COLUMNS:
        if col not in df_out.columns:
            df_out[col] = ""
    df_out = df_out[FINAL_COLUMNS]

    print(f"\n[5/5] Menyimpan hasil ke: {OUTPUT_FILE} ...")
    df_out.to_excel(OUTPUT_FILE, index=False)

    print(f"\n{'=' * 60}")
    print(f"  [OK] SELESAI!")
    print(f"  Total baris output   : {len(df_out):,}")
    print(f"  Total kolom output   : {len(df_out.columns)}")
    print(f"  File disimpan di     : {OUTPUT_FILE}")

    print(f"\n  Ringkasan FLAG_PROD:")
    for flag, cnt in df_out["FLAG_PROD"].value_counts().items():
        print(f"    {flag:<35}: {cnt:,}")

    print(f"\n  Ringkasan SKENARIO:")
    for sce, cnt in df_out["SKENARIO"].value_counts().items():
        print(f"    {sce:<40}: {cnt:,}")

    print("=" * 60)


if __name__ == "__main__":
    run()