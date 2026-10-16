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

with engine.connect() as con:
    res = con.execute(text("SELECT table_name FROM information_schema.tables WHERE table_schema = 'public'")).fetchall()
    tables = [r[0] for r in res]
    print("Tables in DB:", sorted(tables))
    
    # Check tables related to SUSPENSE
    suspense_tables = [t for t in tables if "SUSPENSE" in t.upper() or "FACUL" in t.upper() or "OSBAL" in t.upper()]
    print("\nSuspense/Facul/Osbal Tables:", suspense_tables)
    
    for t in suspense_tables:
        cols = [c[0] for c in con.execute(text(f"SELECT column_name FROM information_schema.columns WHERE table_name = '{t}'")).fetchall()]
        cnt = con.execute(text(f'SELECT count(*) FROM "{t}"')).scalar()
        print(f"\nTable: {t} (Row count: {cnt})")
        # Find cedant-like columns
        ced_cols = [c for c in cols if "CEDANT" in c.upper() or "COMP" in c.upper()]
        print(f"  Cedant-like columns: {ced_cols}")
        if ced_cols:
            first_ced_col = ced_cols[0]
            distinct_ced = [r[0] for r in con.execute(text(f'SELECT DISTINCT "{first_ced_col}" FROM "{t}" LIMIT 10')).fetchall()]
            print(f"  Distinct values in {first_ced_col}: {distinct_ced}")
