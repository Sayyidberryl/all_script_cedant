import os
import sys

sys.modules['numexpr'] = None
sys.modules['bottleneck'] = None

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
    res = con.execute(text("SELECT table_name FROM information_schema.tables WHERE table_schema='public'")).fetchall()
    tables = [r[0] for r in res]
    for t in tables:
        if any(k in t.upper() for k in ["SUSPENSE", "SUSP", "FACUL", "OSBAL"]):
            cnt = con.execute(text(f'SELECT count(*) FROM "{t}"')).scalar()
            print(f"{t}: {cnt} rows")
