"""
sync_facul_osbal.py — Export cleaned FACUL & OSBAL for Cedant IBS to PostgreSQL

Updates tables per-cedant (deletes existing IBS data first, then appends)
without touching other cedants:
  - SUSPENSE_DATA_FACUL_CLEAN & SUSPENSE_FACUL_CLEAN_2026
  - SUSPENSE_DATA_OSBAL_CLEAN & SUSPENSE_OSBAL_CLEAN_2026
"""

import sys
sys.modules['numexpr'] = None
sys.modules['bottleneck'] = None

import os
import time
import pandas as pd
from dotenv import load_dotenv
from sqlalchemy import create_engine, text

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
CLEAN_DIR  = os.path.join(SCRIPT_DIR, "data_cleaned")

FACUL_FILE = os.path.join(CLEAN_DIR, "ibs_output_facul_V2.xlsx")
OSBAL_FILE = os.path.join(CLEAN_DIR, "ibs_output_osbal_V2.xlsx")

FACUL_TARGET_COLS = [
    "FAC_CODE", "FAC_CEDANT", "COMP_NAME", "FAC_BROKER", "COMP_NAME2", "BUSSINESS_PATNER",
    "FAC_INSURED", "INSURED_CLEAN_1", "INSURED_CLEAN_2", "INSURED_CLEAN_3", "INSURED_CLEAN_4", "INSURED_CLEAN_5",
    "FAC_POLICY_NO", "POLICY_CLEAN_1", "SERTIF_CLEAN_1", "POLICY_CLEAN_2", "SERTIF_CLEAN_2", "POLICY_CLEAN_3", "SERTIF_CLEAN_3",
    "POLICY_CLEAN_4", "POLICY_CLEAN_5",
    "FAC_SLIP", "SLIP_CLEAN_1", "SLIP_CLEAN_2", "SLIP_CLEAN_3", "SLIP_CLEAN_4", "SLIP_CLEAN_5",
    "FAC_CURRENCY", "FAC_INP_DATE", "FAC_COM_DATE", "FAC_EXP_DATE", "FAC_SUB_CLASS", "FAC_RISK", "FAC_DESC", "FAC_ACC_STS", "FAC_STS_SLIP"
]

OSBAL_TARGET_COLS = [
    "CCOS_DOC_NO", "CCOS_DATE", "CCOS_REF_CODE", "CCOS_COMP", "CCOS_COMP_NAME", "CCOS_REF_COMP", "CCOS_REF_COMP_NAME",
    "FAC_INSURED", "INSURED_CLEAN_1", "INSURED_CLEAN_2", "INSURED_CLEAN_3", "INSURED_CLEAN_4", "INSURED_CLEAN_5",
    "CCOS_CURR", "CCOS_OR_BAL", "CCOS_BAL_DUE", "CCOS_OR_BAL_IN_IDR", "CCOS_BAL_DUE_IN_IDR",
    "FAC_COM_DATE", "FAC_EXP_DATE", "FAC_DUE_DATES", "FAC_SUB_CLASS",
    "FAC_POLICY_NO", "POLICY_CLEAN_1", "SERTIF_CLEAN_1", "POLICY_CLEAN_2", "SERTIF_CLEAN_2", "POLICY_CLEAN_3", "SERTIF_CLEAN_3", "POLICY_CLEAN_4", "POLICY_CLEAN_5",
    "FAC_SLIP", "SLIP_CLEAN_1", "SLIP_CLEAN_2", "SLIP_CLEAN_3", "SLIP_CLEAN_4", "SLIP_CLEAN_5",
    "CLASS_CODE", "CLASS_NAME", "CLSDT_POLICY_NO", "CLSDT_SLIP_NO", "CLSDT_SERTF_NO"
]

def prepare_facul_df(file_path: str) -> pd.DataFrame:
    print(f"  [FACUL] Reading {file_path} ...", flush=True)
    try:
        df = pd.read_excel(file_path, engine="calamine")
    except Exception:
        df = pd.read_excel(file_path)

    facul_rename = {
        "COMP_NAME.1": "COMP_NAME2",
        "COMP_NAME_1": "COMP_NAME2",
        "BUSINESS_PARTNER": "BUSSINESS_PATNER",
        "FAC_INSURED_ORI": "FAC_INSURED",
    }
    for i in range(1, 6):
        facul_rename[f"INSURED_CLN_{i}"] = f"INSURED_CLEAN_{i}"
        facul_rename[f"POLIS_CLN_{i}"]   = f"POLICY_CLEAN_{i}"
        facul_rename[f"CERTIFICATE_{i}"] = f"SERTIF_CLEAN_{i}"
        facul_rename[f"SLIP_NO_CLN_{i}"] = f"SLIP_CLEAN_{i}"

    df.rename(columns=facul_rename, inplace=True)

    date_cols = ["FAC_INP_DATE", "FAC_COM_DATE", "FAC_EXP_DATE"]
    for col in date_cols:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors="coerce").dt.date
            df[col] = df[col].where(df[col].notnull(), None)

    for col in FACUL_TARGET_COLS:
        if col not in df.columns:
            df[col] = None

    df = df[FACUL_TARGET_COLS]
    for c in df.select_dtypes(include=["object"]).columns:
        if c not in date_cols:
            df[c] = df[c].astype(str).str.strip().replace({"nan": None, "None": None, "": None})

    print(f"  [FACUL] Prepared {len(df):,} rows x {len(df.columns)} cols", flush=True)
    return df

def prepare_osbal_df(file_path: str) -> pd.DataFrame:
    print(f"  [OSBAL] Reading {file_path} ...", flush=True)
    try:
        df = pd.read_excel(file_path, engine="calamine")
    except Exception:
        df = pd.read_excel(file_path)

    osbal_rename = {
        "FAC_INSURED_ORI": "FAC_INSURED",
    }
    for i in range(1, 6):
        osbal_rename[f"INSURED_CLN_{i}"] = f"INSURED_CLEAN_{i}"
        osbal_rename[f"POLIS_CLN_{i}"]   = f"POLICY_CLEAN_{i}"
        osbal_rename[f"CERTIFICATE_{i}"] = f"SERTIF_CLEAN_{i}"
        osbal_rename[f"SLIP_NO_CLN_{i}"] = f"SLIP_CLEAN_{i}"

    df.rename(columns=osbal_rename, inplace=True)

    date_cols = ["CCOS_DATE", "FAC_COM_DATE", "FAC_EXP_DATE", "FAC_DUE_DATES"]
    for col in date_cols:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors="coerce").dt.date
            df[col] = df[col].where(df[col].notnull(), None)

    num_cols = ["CCOS_OR_BAL", "CCOS_BAL_DUE", "CCOS_OR_BAL_IN_IDR", "CCOS_BAL_DUE_IN_IDR"]
    for col in num_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")

    for col in OSBAL_TARGET_COLS:
        if col not in df.columns:
            df[col] = None

    df = df[OSBAL_TARGET_COLS]
    for c in df.select_dtypes(include=["object"]).columns:
        if c not in date_cols:
            df[c] = df[c].astype(str).str.strip().replace({"nan": None, "None": None, "": None})

    print(f"  [OSBAL] Prepared {len(df):,} rows x {len(df.columns)} cols", flush=True)
    return df

def run_sync() -> None:
    t_start = time.perf_counter()
    print("=" * 80)
    print("  SYNC MASTER DATA (FACUL & OSBAL) TO POSTGRESQL — CEDANT IBS")
    print("=" * 80)

    env_path = os.path.join(SCRIPT_DIR, ".env")
    load_dotenv(env_path)
    load_dotenv()

    db_user = os.environ.get("DB_USER", "viy4user")
    db_pass = os.environ.get("DB_PASS", "fild42op61nx")
    db_host = os.environ.get("DB_HOST", "10.10.125.91")
    db_port = os.environ.get("DB_PORT", "5432")
    db_name = os.environ.get("DB_NAME", "fsi_db")
    conn_str = f"postgresql+psycopg2://{db_user}:{db_pass}@{db_host}:{db_port}/{db_name}"

    engine = create_engine(conn_str)

    # 1. FACUL
    if os.path.exists(FACUL_FILE):
        df_facul = prepare_facul_df(FACUL_FILE)
        print("  Exporting FACUL to database ...", flush=True)
        facul_tables = ["SUSPENSE_DATA_FACUL_CLEAN", "SUSPENSE_FACUL_CLEAN_2026"]

        with engine.begin() as con:
            for tbl in facul_tables:
                del_sql = text(f'DELETE FROM "{tbl}" WHERE "COMP_NAME" LIKE \'%IBS%\';')
                con.execute(del_sql)

        for tbl in facul_tables:
            df_facul.to_sql(tbl, con=engine, if_exists="append", index=False)
            print(f"    -> {len(df_facul):,} rows berhasil di-export ke '{tbl}'", flush=True)
    else:
        print(f"  [WARN] File {FACUL_FILE} tidak ditemukan, lewati FACUL sync.", flush=True)

    # 2. OSBAL
    if os.path.exists(OSBAL_FILE):
        df_osbal = prepare_osbal_df(OSBAL_FILE)
        print("  Exporting OSBAL to database ...", flush=True)
        osbal_tables = ["SUSPENSE_DATA_OSBAL_CLEAN", "SUSPENSE_OSBAL_CLEAN_2026"]

        with engine.begin() as con:
            for tbl in osbal_tables:
                del_sql = text(f'DELETE FROM "{tbl}" WHERE "CCOS_COMP_NAME" LIKE \'%IBS%\';')
                con.execute(del_sql)

        for tbl in osbal_tables:
            df_osbal.to_sql(tbl, con=engine, if_exists="append", index=False)
            print(f"    -> {len(df_osbal):,} rows berhasil di-export ke '{tbl}'", flush=True)
    else:
        print(f"  [WARN] File {OSBAL_FILE} tidak ditemukan, lewati OSBAL sync.", flush=True)

    elapsed = time.perf_counter() - t_start
    print(f"\n[OK] Sync Master Data IBS selesai dalam {elapsed:.1f} detik.")
    print("=" * 80)

if __name__ == "__main__":
    run_sync()
