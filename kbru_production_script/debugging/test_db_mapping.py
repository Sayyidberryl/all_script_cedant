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

df_facul = pd.read_excel("data/kbru_output_facul.xlsx")
df_osbal = pd.read_excel("data/kbru_output_osbal.xlsx")

with engine.connect() as con:
    facul_db_cols = [c[0] for c in con.execute(text("SELECT column_name FROM information_schema.columns WHERE table_name = 'SUSPENSE_DATA_FACUL_CLEAN' ORDER BY ordinal_position")).fetchall()]
    osbal_db_cols = [c[0] for c in con.execute(text("SELECT column_name FROM information_schema.columns WHERE table_name = 'SUSPENSE_DATA_OSBAL_CLEAN' ORDER BY ordinal_position")).fetchall()]

print("=== FACUL DB COLS ===")
print("DB cols count:", len(facul_db_cols))
print("Excel cols count:", len(df_facul.columns))

# Check rename mapping for facul
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
df_facul_renamed = df_facul.rename(columns=facul_rename)
for c in facul_db_cols:
    if c not in df_facul_renamed.columns:
        df_facul_renamed[c] = None
df_facul_export = df_facul_renamed[facul_db_cols]
print("FACUL export shape:", df_facul_export.shape)
print("FACUL missing any DB col?", set(facul_db_cols) - set(df_facul_export.columns))

print("\n=== OSBAL DB COLS ===")
print("DB cols count:", len(osbal_db_cols))
print("Excel cols count:", len(df_osbal.columns))

osbal_rename = {
    "FAC_INSURED_ORI": "FAC_INSURED",
    "FAC_POLICY_NO_ORI": "FAC_POLICY_NO",
    "POLIS_CLEAN_1": "POLICY_CLEAN_1",
    "CERTIFICATE_1": "SERTIF_CLEAN_1",
    "POLIS_CLEAN_2": "POLICY_CLEAN_2",
    "FAC_SLIP_ORI": "FAC_SLIP",
}
df_osbal_renamed = df_osbal.rename(columns=osbal_rename)
for c in osbal_db_cols:
    if c not in df_osbal_renamed.columns:
        df_osbal_renamed[c] = None
df_osbal_export = df_osbal_renamed[osbal_db_cols]
print("OSBAL export shape:", df_osbal_export.shape)
print("OSBAL missing any DB col?", set(osbal_db_cols) - set(df_osbal_export.columns))
