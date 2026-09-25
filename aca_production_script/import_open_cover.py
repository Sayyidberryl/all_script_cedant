import os
import sys

sys.modules['numexpr'] = None
sys.modules['bottleneck'] = None

import pandas as pd
from dotenv import load_dotenv
from sqlalchemy import create_engine, text

# Cari file bordero
def _find_bordero_file() -> str:
    candidates = [
        os.path.join("data", "ACA_Open_Cover_Marine_Cargo.xlsx"),
        os.path.join("data", "ACA_Database_Open_Cover_Marine_Cargo.xlsx"),
        os.path.join("data", "ACA_Database_Open_Cover_Marine_Hull.xlsx"),
    ]
    for p in candidates:
        if os.path.exists(p):
            return p
    data_dir = "data"
    if os.path.exists(data_dir):
        for fname in os.listdir(data_dir):
            if "open_cover" in fname.lower() and fname.endswith(".xlsx"):
                return os.path.join(data_dir, fname)
    return candidates[0]

def import_to_db():
    load_dotenv()
    
    file_path = _find_bordero_file()
    print(f"Membaca file: {file_path}")
    if not os.path.exists(file_path):
        print(f"File tidak ditemukan: {file_path}")
        return
        
    df = pd.read_excel(file_path)
    print(f"Berhasil membaca {len(df)} baris.")
    
    db_user = os.environ.get("DB_USER", "postgres")
    db_pass = os.environ.get("DB_PASS", "postgres")
    db_host = os.environ.get("DB_HOST", "localhost")
    db_port = os.environ.get("DB_PORT", "5432")
    db_name = os.environ.get("DB_NAME", "postgres")
    
    conn_str = f"postgresql+psycopg2://{db_user}:{db_pass}@{db_host}:{db_port}/{db_name}"
    engine = create_engine(conn_str)
    
    table_name = "suspense_open_cover_marine_hull"
    print(f"Memasukkan data ke tabel: {table_name}")
    
    with engine.begin() as conn:
        conn.execute(text(f"DROP TABLE IF EXISTS {table_name} CASCADE"))
        
    df.to_sql(table_name, con=engine, if_exists="append", index=False)
    print("Selesai memasukkan data ke database.")

if __name__ == "__main__":
    import_to_db()
