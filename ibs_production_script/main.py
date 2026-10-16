"""
main.py — Master Orchestrator Pipeline Script for Cedant IBS (End-to-End)

Pipeline Eksekusi:
Phase 1: Data Cleansing
  1. cleansing_data_1_facul.py     -> data_cleaned/ibs_output_facul_V2.xlsx
  2. cleansing_data_2_osbal.py     -> data_cleaned/ibs_output_osbal_V2.xlsx
  3. cleansing_data_3_suspense.py  -> data_cleaned/ibs_output_suspense_V2.xlsx

Phase 2: Master Data Synchronization ke Database PostgreSQL
  4. sync_facul_osbal.py           -> SUSPENSE_DATA_FACUL_CLEAN & SUSPENSE_DATA_OSBAL_CLEAN

Phase 3: Matching & Suspend Synchronization ke Database PostgreSQL
  5. matching_ibs_v1.py            -> final_output_ibs_v1.xlsx &
                                      PostgreSQL (SUSPENSE_DATA_SUSPENSE_IBS_V1,
                                                  SUSPENSE_DATA_SUSPENSE_V1,
                                                  SUSPENSE_CLEAN_2026)
"""

import os
import sys
import subprocess
import time

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

PIPELINE_STEPS = [
    {
        "num": 1,
        "name": "Cleansing Data Facultative (IBS)",
        "script": "cleansing_data_1_facul.py",
        "output": os.path.join(SCRIPT_DIR, "data_cleaned", "ibs_output_facul_V2.xlsx"),
    },
    {
        "num": 2,
        "name": "Cleansing Data OSBAL (IBS)",
        "script": "cleansing_data_2_osbal.py",
        "output": os.path.join(SCRIPT_DIR, "data_cleaned", "ibs_output_osbal_V2.xlsx"),
    },
    {
        "num": 3,
        "name": "Cleansing Data Suspense (IBS)",
        "script": "cleansing_data_3_suspense.py",
        "output": os.path.join(SCRIPT_DIR, "data_cleaned", "ibs_output_suspense_V2.xlsx"),
    },
    {
        "num": 4,
        "name": "Sync Master Data (FACUL & OSBAL) to PostgreSQL DB",
        "script": "sync_facul_osbal.py",
        "output": None,
    },
    {
        "num": 5,
        "name": "Matching IBS & Export Suspend to PostgreSQL DB",
        "script": "matching_ibs_v1.py",
        "output": os.path.join(SCRIPT_DIR, "final_output_ibs_v1.xlsx"),
    },
]

def execute_step(step_info: dict) -> dict:
    num = step_info["num"]
    name = step_info["name"]
    script_name = step_info["script"]
    output_file = step_info["output"]
    script_path = os.path.join(SCRIPT_DIR, script_name)
    log_file = os.path.join(SCRIPT_DIR, f"{script_name}.log")

    print(f"\n[{num}/5] START: {name}", flush=True)
    print(f"      Script: {script_name}", flush=True)

    if not os.path.exists(script_path):
        print(f"      [ERROR] File script {script_path} tidak ditemukan!", flush=True)
        return {"step": num, "name": name, "status": "FAILED (Script missing)", "elapsed": 0.0}

    t_start = time.perf_counter()
    try:
        res = subprocess.run(
            [sys.executable, script_path],
            check=True,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            cwd=SCRIPT_DIR
        )
        elapsed = time.perf_counter() - t_start
        stdout_str = res.stdout if res.stdout else ""

        with open(log_file, "w", encoding="utf-8", errors="replace") as f:
            f.write(stdout_str)

        size_mb = os.path.getsize(output_file) / (1024 * 1024) if (output_file and os.path.exists(output_file)) else 0.0
        out_desc = f"{output_file} ({size_mb:.2f} MB)" if output_file and os.path.exists(output_file) else "PostgreSQL DB"
        print(f"      [OK] Selesai dalam {elapsed:.1f} detik -> {out_desc}", flush=True)

        return {"step": num, "name": name, "status": "SUCCESS", "elapsed": elapsed, "output": out_desc}

    except subprocess.CalledProcessError as e:
        elapsed = time.perf_counter() - t_start
        stdout_str = e.stdout if e.stdout else ""
        stderr_str = e.stderr if e.stderr else ""

        with open(log_file, "w", encoding="utf-8", errors="replace") as f:
            f.write(stdout_str)
            f.write("\n\n--- ERROR STDERR ---\n\n")
            f.write(stderr_str)

        print(f"      [FAIL] Step {num} GAGAL dalam {elapsed:.1f} detik! Cek log di {log_file}", flush=True)
        if stderr_str:
            last_err = stderr_str.strip().splitlines()[-3:]
            print("      Detail Error:\n      " + "\n      ".join(last_err), flush=True)

        return {"step": num, "name": name, "status": f"FAILED ({e.returncode})", "elapsed": elapsed, "output": "-"}

def run_pipeline() -> None:
    t_total_start = time.perf_counter()
    print("=" * 80, flush=True)
    print("  MASTER PIPELINE CLEANSING & MATCHING IBS (END-TO-END)", flush=True)
    print("  Cedant: IBS (PT INTER BENUA REASINDO)", flush=True)
    print("  Tujuan: Cleansing 1, 2, 3 -> Master Sync DB -> Matching IBS -> Suspense Sync DB", flush=True)
    print("=" * 80, flush=True)

    summary = []
    for step in PIPELINE_STEPS:
        res = execute_step(step)
        summary.append(res)
        if "FAIL" in res["status"]:
            print(f"\n[CRITICAL] Terjadi kegagalan pada '{res['name']}'. Pipeline dihentikan.", flush=True)
            sys.exit(1)

    t_total = time.perf_counter() - t_total_start
    print("\n" + "=" * 80, flush=True)
    print("  RINGKASAN PIPELINE END-TO-END IBS BERHASIL", flush=True)
    print("=" * 80, flush=True)
    for r in summary:
        print(f"  Step {r['step']}: {r['name']:<50} | {r['status']:<8} | {r['elapsed']:>6.1f}s | {r['output']}", flush=True)
    print("-" * 80, flush=True)
    print(f"  Total Durasi: {t_total:.1f} detik ({t_total/60:.2f} menit)", flush=True)
    print("=" * 80, flush=True)

if __name__ == "__main__":
    run_pipeline()
