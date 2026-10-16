"""
main.py — Master Orchestrator Script for Cedant KBRU (Optimized)

Pipeline Eksekusi:
Phase 1:
  1. sync_facul_osbal.py — Sinkronisasi Master Data FACUL & OSBAL ke PostgreSQL per-cedant update
Phase 2 (Paralel):
  2. prod_sus_1.py — Suspend Matching v1 (Row Asli Dipertahankan) -> Excel & DB (SUSPENSE_DATA_SUSPENSE_V1 & SUSPENSE_CLEAN_2026)
  3. prod_sus_2.py — Suspend Matching v2 (Akumulasi per FAC Code Tunggal) -> Excel & DB (SUSPENSE_DATA_SUSPENSE_V2)
"""

import os
import sys
import subprocess
import time
from concurrent.futures import ProcessPoolExecutor, as_completed

# Script constants
FACUL_FILE   = os.path.join("data", "kbru_output_facul.xlsx")
OSBAL_FILE   = os.path.join("data", "kbru_output_osbal.xlsx")
SUSPEND_FILE = os.path.join("data", "kbru_output_suspend.xlsx")

OUTPUT_V1    = os.path.join("data", "final_output_v1.xlsx")
OUTPUT_V2    = os.path.join("data", "final_output_v2.xlsx")

PIPELINE_PHASES = [
    {
        "phase_name": "Phase 1: Master Data Database Synchronization (FACUL & OSBAL)",
        "steps": [
            {"num": 1, "name": "Sync Master Facul & Osbal to DB", "script": "sync_facul_osbal.py", "output": None},
        ]
    },
    {
        "phase_name": "Phase 2: Suspend Matching (v1 & v2)",
        "steps": [
            {"num": 2, "name": "Suspend Matching v1 (Row Asli)", "script": "prod_sus_1.py", "output": OUTPUT_V1},
            {"num": 3, "name": "Suspend Matching v2 (Akumulasi)", "script": "prod_sus_2.py", "output": OUTPUT_V2},
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
        return {"step": num, "name": name, "output": output_file or "-", "status": "FAILED (Script missing)", "size_mb": 0.0, "elapsed": 0.0}

    t_start = time.perf_counter()
    try:
        res = subprocess.run([sys.executable, script], check=True, capture_output=True, text=True, encoding='utf-8', errors='replace')
        elapsed = time.perf_counter() - t_start
        
        stdout_str = res.stdout if res.stdout is not None else ""
        with open(log_file, "w", encoding="utf-8", errors="replace") as f:
            f.write(stdout_str)
            
        size_mb = os.path.getsize(output_file) / (1024 * 1024) if (output_file and os.path.exists(output_file)) else 0.0
        out_display = output_file if output_file else "PostgreSQL DB"
        print(f"  [OK] Step {num} selesai dalam {elapsed:.1f} detik -> {out_display}", flush=True)
        return {"step": num, "name": name, "output": out_display, "status": "SUCCESS", "size_mb": size_mb, "elapsed": elapsed}
    
    except subprocess.CalledProcessError as e:
        elapsed = time.perf_counter() - t_start
        stdout_str = e.stdout if e.stdout is not None else ""
        stderr_str = e.stderr if e.stderr is not None else ""
        with open(log_file, "w", encoding="utf-8", errors="replace") as f:
            f.write(stdout_str)
            f.write("\n\n--- ERROR ---\n\n")
            f.write(stderr_str)
            
        print(f"  [FAIL] Step {num} GAGAL! Cek {log_file} untuk detail error.", flush=True)
        return {"step": num, "name": name, "output": output_file or "-", "status": f"FAILED ({e.returncode})", "size_mb": 0.0, "elapsed": elapsed}

def run_pipeline():
    t_total_start = time.perf_counter()
    print("=" * 80, flush=True)
    print("  PIPELINE DATA ORCHESTRATOR — INDORE STARCORE (CEDANT KBRU)", flush=True)
    print("  Urutan Matching: Slip -> Polis -> Insured | Tanpa Sertifikat | Tanpa Periode 2", flush=True)
    print("  Mekanisme: Update Per-Cedant ke PostgreSQL DB", flush=True)
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
    print("  RINGKASAN EKSEKUSI PIPELINE KBRU (EXCEL & POSTGRESQL)", flush=True)
    print("=" * 80, flush=True)
    print(f"{'No':<4} {'Nama Step':<38} {'Status':<10} {'Size (MB)':<10} {'Waktu (s)':<10}")
    print("-" * 80, flush=True)

    summary_results.sort(key=lambda x: x["step"])
    for item in summary_results:
        print(f"{item['step']:<4} {item['name']:<38} {item['status']:<10} {item['size_mb']:<10.2f} {item['elapsed']:<10.1f}")

    print("-" * 80, flush=True)
    print(f"Total Waktu Eksekusi: {t_total:.1f} detik ({t_total/60:.1f} menit)", flush=True)
    print("Semua output (Excel & DB PostgreSQL) berhasil diperbarui tanpa menimpa cedant lain!", flush=True)
    print("=" * 80, flush=True)

if __name__ == "__main__":
    run_pipeline()
