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

tables = ["SUSPENSE_DATA_OSBAL_CLEAN", "SUSPENSE_DATA_FACUL_CLEAN"]
with engine.begin() as conn:
    for t in tables:
        try:
            conn.execute(text(f'TRUNCATE TABLE "{t}"'))
            print(f"Truncated {t}")
        except Exception as e:
            print(f"Could not truncate {t}: {e}")

print("Done truncating. Now run cleaning_osbal.py and cleaning_facul.py")
