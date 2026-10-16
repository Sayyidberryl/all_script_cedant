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
    for t in [
        'SUSPENSE_DATA_FACUL_CLEAN', 'SUSPENSE_FACUL_CLEAN_2026',
        'SUSPENSE_DATA_OSBAL_CLEAN', 'SUSPENSE_OSBAL_CLEAN_2026',
        'SUSPENSE_CLEAN_2026', 'SUSPENSE_DATA_SUSPENSE_V1',
        'SUSPENSE_DATA_SUSPENSE_V2'
    ]:
        cols = [c[0] for c in con.execute(text(f"SELECT column_name FROM information_schema.columns WHERE table_name = '{t}' ORDER BY ordinal_position")).fetchall()]
        cnt = con.execute(text(f'SELECT count(*) FROM "{t}"')).scalar()
        print(f"=== Table: {t} (Rows: {cnt}) ===")
        print("Columns:", cols)
