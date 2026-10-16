import os, sys
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
    for t in ['SUSPENSE_DATA_FACUL_CLEAN', 'SUSPENSE_DATA_OSBAL_CLEAN']:
        res = con.execute(text(f"SELECT column_name, data_type FROM information_schema.columns WHERE table_name = '{t}' ORDER BY ordinal_position")).fetchall()
        print(f"=== {t} types ===")
        for col, dt in res:
            if 'date' in dt or 'time' in dt or 'num' in dt or 'double' in dt or 'int' in dt:
                print(f"  {col}: {dt}")
