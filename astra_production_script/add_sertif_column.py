import os
from dotenv import load_dotenv
from sqlalchemy import create_engine, text

load_dotenv()

db_user = os.environ.get("DB_USER", "postgres")
db_pass = os.environ.get("DB_PASS", "postgres")
db_host = os.environ.get("DB_HOST", "localhost")
db_port = os.environ.get("DB_PORT", "5432")
db_name = os.environ.get("DB_NAME", "postgres")

conn_str = f"postgresql+psycopg2://{db_user}:{db_pass}@{db_host}:{db_port}/{db_name}"
engine = create_engine(conn_str)

queries = [
    'ALTER TABLE "SUSPENSE_DATA_SUSPENSE_V1" ADD COLUMN IF NOT EXISTS "SERTIF_CLN" TEXT;',
    'ALTER TABLE "SUSPENSE_DATA_SUSPENSE_V2" ADD COLUMN IF NOT EXISTS "SERTIF_CLN" TEXT;',
]

with engine.begin() as conn:
    for q in queries:
        try:
            conn.execute(text(q))
            print(f"Successfully executed: {q}")
        except Exception as e:
            print(f"Failed to execute {q}: {e}")

print("Database altered.")
