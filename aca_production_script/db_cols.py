import os
import pandas as pd
from dotenv import load_dotenv
from sqlalchemy import create_engine

load_dotenv()
engine = create_engine(f'postgresql+psycopg2://{os.environ.get("DB_USER", "postgres")}:{os.environ.get("DB_PASS", "postgres")}@{os.environ.get("DB_HOST", "localhost")}:{os.environ.get("DB_PORT", "5432")}/{os.environ.get("DB_NAME", "postgres")}')
print(pd.read_sql('SELECT * FROM "SUSPENSE_DATA_FACUL_CLEAN" LIMIT 0', engine).columns.tolist())
