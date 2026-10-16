import os
import sys

sys.modules['numexpr'] = None
sys.modules['bottleneck'] = None

import pandas as pd
from dotenv import load_dotenv
from sqlalchemy import create_engine, text

load_dotenv()
db_user = os.environ.get("DB_USER", "postgres")
db_pass = os.environ.get("DB_PASS", "postgres")
db_host = os.environ.get("DB_HOST", "localhost")
db_port = os.environ.get("DB_PORT", "5432")
db_name = os.environ.get("DB_NAME", "postgres")

engine = create_engine(f"postgresql+psycopg2://{db_user}:{db_pass}@{db_host}:{db_port}/{db_name}")

print("--- EXCEL FILES CHECK ---")
df_facul = pd.read_excel("data/kbru_output_facul.xlsx", nrows=10)
df_osbal = pd.read_excel("data/kbru_output_osbal.xlsx", nrows=10)
df_sus = pd.read_excel("data/kbru_output_suspend.xlsx", nrows=10)

print("FACUL cols:", df_facul.columns.tolist())
print("FACUL cedant samples:")
for c in ['FAC_CEDANT', 'COMP_NAME', 'FAC_BROKER', 'COMP_NAME.1', 'BUSINESS_PARTNER']:
    if c in df_facul.columns:
        print(f"  {c}:", df_facul[c].unique()[:5])

print("\nOSBAL cols:", df_osbal.columns.tolist())
print("OSBAL comp samples:")
for c in ['CCOS_COMP', 'CCOS_COMP_NAME', 'CCOS_REF_COMP', 'CCOS_REF_COMP_NAME']:
    if c in df_osbal.columns:
        print(f"  {c}:", df_osbal[c].unique()[:5])

print("\nSUSPEND cols:", df_sus.columns.tolist())
print("SUSPEND cedant samples:")
for c in ['CEDANT NAME', 'CEDANT SHRT NAME']:
    if c in df_sus.columns:
        print(f"  {c}:", df_sus[c].unique()[:5])

print("\n--- POSTGRES DB DISTINCT CEDANTS ---")
with engine.connect() as con:
    for tbl, col in [
        ('SUSPENSE_DATA_FACUL_CLEAN', 'FAC_CEDANT'),
        ('SUSPENSE_DATA_FACUL_CLEAN', 'COMP_NAME'),
        ('SUSPENSE_DATA_OSBAL_CLEAN', 'CCOS_COMP'),
        ('SUSPENSE_DATA_OSBAL_CLEAN', 'CCOS_COMP_NAME'),
        ('SUSPENSE_CLEAN_2026', 'CEDANT NAME'),
        ('SUSPENSE_DATA_SUSPENSE_V1', 'CEDANT NAME'),
        ('SUSPENSE_DATA_SUSPENSE_V2', 'CEDANT NAME'),
    ]:
        try:
            res = [r[0] for r in con.execute(text(f'SELECT DISTINCT "{col}" FROM "{tbl}"')).fetchall()]
            print(f"{tbl} -> {col}: {res}")
        except Exception as e:
            print(f"{tbl} -> {col} error: {e}")
