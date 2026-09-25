"""
main.py — Master Orchestrator Script (Optimized)

Menjalankan pipeline data menggunakan paralelisme:
Phase 1 (Paralel):
  1. cleaning_facul.py
  2. cleaning_osbal.py
  3. cleaning_suspend.py
Phase 2 (Paralel):
  4. prod_sus_1.py
  5. prod_sus_2.py
"""

import os
import sys
import subprocess
import time
from concurrent.futures import ProcessPoolExecutor, as_completed

# Script constants
SUSPEND_FILE = os.path.join("data", "suspend_clean_aca.xlsx")
OSBAL_FILE   = os.path.join("data", "osbal_clean_aca.xlsx")
FACUL_FILE   = os.path.join("data", "facul_clean_aca.xlsx")

PIPELINE_PHASES = [
    {
        "phase_name": "Phase 1: Data Cleaning",
        "steps": [
            {"num": 1, "name": "Cleaning Facul Data", "script": "cleaning_facul.py", "output": FACUL_FILE},
            {"num": 2, "name": "Cleaning OSBAL Data", "script": "cleaning_osbal.py", "output": OSBAL_FILE},
            {"num": 3, "name": "Cleaning Suspend Data", "script": "cleaning_suspend.py", "output": SUSPEND_FILE},
        ]
    },
    {
        "phase_name": "Phase 2: Suspend Matching (v1 & v2)",
        "steps": [
            {"num": 4, "name": "Suspend Matching v1", "script": "prod_sus_1.py", "output": os.path.join("data", "final_output_v1.xlsx")},
            {"num": 5, "name": "Suspend Matching v2", "script": "prod_sus_2.py", "output": os.path.join("data", "final_output_v2.xlsx")},
        ]
    }
]

def execute_step(step_info):
    num = step_info["num"]
    name = step_info["name"]
    script = step_info["script"]
    output_file = step_info["output"]
    log_file = os.path.join("data", f"{script}.log")
    
    print(f"  [START] Step {num}: {name} (Menulis log ke {log_file})", flush=True)
    if not os.path.exists(script):
        return {"step": num, "name": name, "output": output_file, "status": "FAILED (Script missing)", "size_mb": 0.0, "elapsed": 0.0}

    t_start = time.perf_counter()
    try:
        # Menjalankan script dan menangkap outputnya agar tidak bertumpuk di console
        res = subprocess.run([sys.executable, script], check=True, capture_output=True, text=True, encoding='utf-8', errors='replace')
        elapsed = time.perf_counter() - t_start
        
        stdout_str = res.stdout if res.stdout is not None else ""
        with open(log_file, "w", encoding="utf-8", errors="replace") as f:
            f.write(stdout_str)
            
        size_mb = os.path.getsize(output_file) / (1024 * 1024) if os.path.exists(output_file) else 0.0
        print(f"  [OK] Step {num} selesai dalam {elapsed:.1f} detik -> {output_file}", flush=True)
        return {"step": num, "name": name, "output": output_file, "status": "SUCCESS", "size_mb": size_mb, "elapsed": elapsed}
    
    except subprocess.CalledProcessError as e:
        elapsed = time.perf_counter() - t_start
        stdout_str = e.stdout if e.stdout is not None else ""
        stderr_str = e.stderr if e.stderr is not None else ""
        with open(log_file, "w", encoding="utf-8", errors="replace") as f:
            f.write(stdout_str)
            f.write("\n\n--- ERROR ---\n\n")
            f.write(stderr_str)
            
        print(f"  [FAIL] Step {num} GAGAL! Cek {log_file} untuk detail error.", flush=True)
        return {"step": num, "name": name, "output": output_file, "status": f"FAILED ({e.returncode})", "size_mb": 0.0, "elapsed": elapsed}

def run_pipeline():
    t_total_start = time.perf_counter()
    print("=" * 80, flush=True)
    print("  PIPELINE DATA ORCHESTRATOR — INDORE STARCORE (OPTIMIZED)", flush=True)
    print("  Menjalankan eksekusi paralel (Multiprocessing)", flush=True)
    print("=" * 80, flush=True)

    summary_results = []
    
    for phase in PIPELINE_PHASES:
        print(f"\n>> {phase['phase_name']} ...", flush=True)
        print("-" * 80, flush=True)
        
        with ProcessPoolExecutor(max_workers=len(phase["steps"])) as executor:
            futures = [executor.submit(execute_step, step) for step in phase["steps"]]
            for future in as_completed(futures):
                result = future.result()
                summary_results.append(result)
                if "FAIL" in result["status"]:
                    print(f"\n[CRITICAL] Terjadi kegagalan pada {result['name']}. Pipeline dihentikan secara prematur.", flush=True)
                    sys.exit(1)

    t_total = time.perf_counter() - t_total_start

    print("\n" + "=" * 80, flush=True)
    print("  RINGKASAN EKSEKUSI PIPELINE (5 EXCEL & POSTGRESQL OUTPUTS)", flush=True)
    print("=" * 80, flush=True)
    print(f"{'No':<4} {'Nama Step':<35} {'Status':<10} {'Size (MB)':<10} {'Waktu (s)':<10}")
    print("-" * 80, flush=True)

    summary_results.sort(key=lambda x: x["step"])
    for item in summary_results:
        print(f"{item['step']:<4} {item['name']:<35} {item['status']:<10} {item['size_mb']:<10.2f} {item['elapsed']:<10.1f}")

    print("-" * 80, flush=True)
    print(f"Total Waktu Eksekusi: {t_total:.1f} detik ({t_total/60:.1f} menit)", flush=True)
    print("Semua output (Excel & DB PostgreSQL) berhasil diperbarui!", flush=True)
    print("=" * 80, flush=True)

if __name__ == "__main__":
    run_pipeline()
