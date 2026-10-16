"""
sync_facul_osbal.py — Export cleaned FACUL & OSBAL for Cedant KBRU to PostgreSQL

Updates tables per-cedant (deletes existing KBRU data first, then appends)
without touching other cedants:
  - SUSPENSE_DATA_FACUL_CLEAN & SUSPENSE_FACUL_CLEAN_2026
  - SUSPENSE_DATA_OSBAL_CLEAN & SUSPENSE_OSBAL_CLEAN_2026
"""

import os
import sys
import time

sys.modules['numexpr'] = None
sys.modules['bottleneck'] = None

import numpy as np
import pandas as pd
from dotenv import load_dotenv
from sqlalchemy import create_engine, text

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR   = os.path.join(SCRIPT_DIR, "data")

FACUL_FILE = os.path.join(DATA_DIR, "kbru_output_facul.xlsx")
OSBAL_FILE = os.path.join(DATA_DIR, "kbru_output_osbal.xlsx")

KBRU_CEDANT_CODE = "9S039"
KBRU_CEDANT_NAME = "PT. KBRU REINSURANCE BROKERS (EX. SIMAS)"

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
    df = pd.read_excel(file_path)
    
    facul_rename = {
        "COMP_NAME.1": "COMP_NAME2",
        "BUSINESS_PARTNER": "BUSSINESS_PATNER",
        "FAC_INSURED_ORI": "FAC_INSURED",
        "FAC_INSURED_1": "INSURED_CLEAN_1",
        "FAC_INSURED_2": "INSURED_CLEAN_2",
        "FAC_INSURED_3": "INSURED_CLEAN_3",
        "FAC_INSURED_4": "INSURED_CLEAN_4",
        "FAC_INSURED_5": "INSURED_CLEAN_5",
        "POLIS_CLEAN_1": "POLICY_CLEAN_1",
        "CERTIFICATE_1": "SERTIF_CLEAN_1",
        "POLIS_CLEAN_2": "POLICY_CLEAN_2",
        "POLIS_CLEAN_3": "POLICY_CLEAN_3",
        "FAC_SLIP_CLEAN_1": "SLIP_CLEAN_1",
        "FAC_SLIP_CLEAN_2": "SLIP_CLEAN_2",
        "FAC_SLIP_CLEAN_3": "SLIP_CLEAN_3",
        "FAC_SLIP_CLEAN_4": "SLIP_CLEAN_4",
    }
    df = df.rename(columns=facul_rename)
    
    # Ensure all target columns exist
    for col in FACUL_TARGET_COLS:
        if col not in df.columns:
            df[col] = None
    df = df[FACUL_TARGET_COLS].copy()
    
    # Format date columns
    date_cols = ["FAC_INP_DATE", "FAC_COM_DATE", "FAC_EXP_DATE"]
    for c in date_cols:
        df[c] = pd.to_datetime(df[c], errors="coerce").dt.date
        df[c] = df[c].where(df[c].notnull(), None)
        
    # Text columns: clean whitespace and nulls
    text_cols = [c for c in FACUL_TARGET_COLS if c not in date_cols]
    for c in text_cols:
        df[c] = df[c].astype(str).str.strip()
        df[c] = df[c].replace({"nan": None, "None": None, "": None})
        
    print(f"  [FACUL] Prepared {len(df):,} rows and {len(df.columns)} columns.", flush=True)
    return df

def prepare_osbal_df(file_path: str) -> pd.DataFrame:
    print(f"  [OSBAL] Reading {file_path} ...", flush=True)
    df = pd.read_excel(file_path)
    
    osbal_rename = {
        "FAC_INSURED_ORI": "FAC_INSURED",
        "FAC_POLICY_NO_ORI": "FAC_POLICY_NO",
        "POLIS_CLEAN_1": "POLICY_CLEAN_1",
        "CERTIFICATE_1": "SERTIF_CLEAN_1",
        "POLIS_CLEAN_2": "POLICY_CLEAN_2",
        "FAC_SLIP_ORI": "FAC_SLIP",
    }
    df = df.rename(columns=osbal_rename)
    
    for col in OSBAL_TARGET_COLS:
        if col not in df.columns:
            df[col] = None
    df = df[OSBAL_TARGET_COLS].copy()
    
    # Format date columns
    date_cols = ["CCOS_DATE", "FAC_COM_DATE", "FAC_EXP_DATE"]
    for c in date_cols:
        df[c] = pd.to_datetime(df[c], errors="coerce").dt.date
        df[c] = df[c].where(df[c].notnull(), None)
        
    # Format numeric columns
    numeric_cols = ["CCOS_OR_BAL", "CCOS_BAL_DUE", "CCOS_OR_BAL_IN_IDR", "CCOS_BAL_DUE_IN_IDR"]
    for c in numeric_cols:
        df[c] = pd.to_numeric(df[c], errors="coerce")
        
    # Text columns: clean whitespace and nulls
    text_cols = [c for c in OSBAL_TARGET_COLS if c not in date_cols and c not in numeric_cols]
    for c in text_cols:
        df[c] = df[c].astype(str).str.strip()
        df[c] = df[c].replace({"nan": None, "None": None, "": None})
        
    print(f"  [OSBAL] Prepared {len(df):,} rows and {len(df.columns)} columns.", flush=True)
    return df

def sync_to_db():
    t_start = time.perf_counter()
    print("=" * 60)
    print("  SYNC MASTER DATA (FACUL & OSBAL) TO POSTGRESQL — CEDANT KBRU")
    print("  Mekanisme: Update per-cedant (aman tidak menimpa cedant lain)")
    print("=" * 60)
    
    load_dotenv()
    db_user = os.environ.get("DB_USER", "postgres")
    db_pass = os.environ.get("DB_PASS", "postgres")
    db_host = os.environ.get("DB_HOST", "localhost")
    db_port = os.environ.get("DB_PORT", "5432")
    db_name = os.environ.get("DB_NAME", "postgres")
    conn_str = f"postgresql+psycopg2://{db_user}:{db_pass}@{db_host}:{db_port}/{db_name}"
    
    engine = create_engine(conn_str)
    
    # 1. FACUL
    df_facul = prepare_facul_df(FACUL_FILE)
    print("  Exporting FACUL to database ...", flush=True)
    facul_tables = ["SUSPENSE_DATA_FACUL_CLEAN", "SUSPENSE_FACUL_CLEAN_2026"]
    
    with engine.begin() as con:
        for tbl in facul_tables:
            del_sql = text(f'DELETE FROM "{tbl}" WHERE "FAC_CEDANT" = :code OR "COMP_NAME" = :name;')
            con.execute(del_sql, {"code": KBRU_CEDANT_CODE, "name": KBRU_CEDANT_NAME})
    
    for tbl in facul_tables:
        df_facul.to_sql(tbl, con=engine, if_exists="append", index=False)
        print(f"    -> {len(df_facul):,} rows berhasil di-export ke '{tbl}'", flush=True)
        
    # 2. OSBAL
    df_osbal = prepare_osbal_df(OSBAL_FILE)
    print("  Exporting OSBAL to database ...", flush=True)
    osbal_tables = ["SUSPENSE_DATA_OSBAL_CLEAN", "SUSPENSE_OSBAL_CLEAN_2026"]
    
    with engine.begin() as con:
        for tbl in osbal_tables:
            del_sql = text(f'DELETE FROM "{tbl}" WHERE "CCOS_COMP" = :code OR "CCOS_COMP_NAME" = :name;')
            con.execute(del_sql, {"code": KBRU_CEDANT_CODE, "name": KBRU_CEDANT_NAME})
            
    for tbl in osbal_tables:
        df_osbal.to_sql(tbl, con=engine, if_exists="append", index=False)
        print(f"    -> {len(df_osbal):,} rows berhasil di-export ke '{tbl}'", flush=True)
        
    elapsed = time.perf_counter() - t_start
    print(f"\n[OK] Selesai sinkronisasi master data FACUL & OSBAL KBRU dalam {elapsed:.1f}s")
    print("=" * 60)

if __name__ == "__main__":
    sync_to_db()
