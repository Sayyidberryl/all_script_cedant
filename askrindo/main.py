"""
main.py — Master Orchestrator Script — ASKRINDO

Pipeline data untuk matching suspend cedant ASKRINDO:

Phase 1 (Paralel):
  1. cleaning_facul.py    → facul_clean_askrindo.xlsx
  2. cleaning_osbal.py    → osbal_clean_askrindo.xlsx
  3. cleaning_suspend.py  → suspend_clean_askrindo.xlsx

Phase 2 (Sequential):
  4. prod_sus_askrindo.py → final_output_askrindo.xlsx

Catatan: Phase 1 berjalan paralel jika cleaning scripts tersedia.
         Jika tidak ada, Phase 1 dilewati dan Phase 2 langsung dijalankan
         dengan file output yang sudah ada (askrindo_output_*.xlsx).
"""

import os
import sys
import subprocess
import time
from concurrent.futures import ProcessPoolExecutor, as_completed

# Script constants
ASKRINDO_OUTPUT_FILE = os.path.join("data", "final_output_askrindo.xlsx")

# Cleaning scripts & output paths (opsional — jika tidak ada, dilewati)
CLEANING_STEPS = [
    {
        "num": 1,
        "name": "Cleaning Facul Data",
        "script": "cleaning_facul.py",
        "output": os.path.join("data", "facul_clean_askrindo.xlsx"),
    },
    {
        "num": 2,
        "name": "Cleaning OSBAL Data",
        "script": "cleaning_osbal.py",
        "output": os.path.join("data", "osbal_clean_askrindo.xlsx"),
    },
    {
        "num": 3,
        "name": "Cleaning Suspend Data",
        "script": "cleaning_suspend.py",
        "output": os.path.join("data", "suspend_clean_askrindo.xlsx"),
    },
]

MATCHING_STEPS = [
    {
        "num": 4,
        "name": "Suspend Matching V1 (Row Asli)",
        "script": "prod_sus_1.py",
        "output": os.path.join("data", "final_output_v1.xlsx"),
    },
    {
        "num": 5,
        "name": "Suspend Matching V2 (Akumulasi)",
        "script": "prod_sus_2.py",
        "output": os.path.join("data", "final_output_v2.xlsx"),
    },
]


def execute_step(step_info: dict) -> dict:
    """Jalankan satu step script dan kembalikan ringkasan eksekusi."""
    num         = step_info["num"]
    name        = step_info["name"]
    script      = step_info["script"]
    output_file = step_info["output"]
    log_file    = os.path.join("data", f"{script}.log")

    print(f"  [START] Step {num}: {name} (log: {log_file})", flush=True)

    if not os.path.exists(script):
        msg = f"SKIPPED (script '{script}' tidak ditemukan)"
        print(f"  [SKIP]  Step {num}: {msg}", flush=True)
        return {
            "step": num, "name": name, "output": output_file,
            "status": msg, "size_mb": 0.0, "elapsed": 0.0,
        }

    t_start = time.perf_counter()
    try:
        res = subprocess.run(
            [sys.executable, script],
            check=True, capture_output=True, text=True,
            encoding="utf-8", errors="replace",
        )
        elapsed = time.perf_counter() - t_start

        with open(log_file, "w", encoding="utf-8", errors="replace") as f:
            f.write(res.stdout or "")

        size_mb = (
            os.path.getsize(output_file) / (1024 * 1024)
            if os.path.exists(output_file)
            else 0.0
        )
        print(f"  [OK]    Step {num} selesai dalam {elapsed:.1f}s -> {output_file}", flush=True)
        return {
            "step": num, "name": name, "output": output_file,
            "status": "SUCCESS", "size_mb": size_mb, "elapsed": elapsed,
        }

    except subprocess.CalledProcessError as e:
        elapsed = time.perf_counter() - t_start
        with open(log_file, "w", encoding="utf-8", errors="replace") as f:
            f.write(e.stdout or "")
            f.write("\n\n--- ERROR ---\n\n")
            f.write(e.stderr or "")

        print(f"  [FAIL]  Step {num} GAGAL! Cek {log_file} untuk detail.", flush=True)
        return {
            "step": num, "name": name, "output": output_file,
            "status": f"FAILED ({e.returncode})", "size_mb": 0.0, "elapsed": elapsed,
        }


def run_pipeline() -> None:
    """Jalankan pipeline ASKRINDO dua fase."""
    t_total_start = time.perf_counter()

    print("=" * 80, flush=True)
    print("  PIPELINE DATA ORCHESTRATOR — ASKRINDO — INDORE STARCORE", flush=True)
    print("  Polis -> Slip -> Insured Matching | No Bordero | No Periode-2", flush=True)
    print("=" * 80, flush=True)

    summary_results = []

    # ----------------------------------------------------------------
    # Phase 1: Data Cleaning (paralel, opsional)
    # ----------------------------------------------------------------
    available_cleaning = [s for s in CLEANING_STEPS if os.path.exists(s["script"])]
    if available_cleaning:
        print(f"\n>> Phase 1: Data Cleaning ({len(available_cleaning)} script) ...", flush=True)
        print("-" * 80, flush=True)

        max_workers = min(len(available_cleaning), 3)
        with ProcessPoolExecutor(max_workers=max_workers) as executor:
            futures = [executor.submit(execute_step, step) for step in available_cleaning]
            for future in as_completed(futures):
                result = future.result()
                summary_results.append(result)
                if "FAIL" in result["status"]:
                    print(
                        f"\n[CRITICAL] Kegagalan pada {result['name']}. Pipeline dihentikan.",
                        flush=True,
                    )
                    sys.exit(1)
    else:
        print(
            "\n[INFO] Phase 1 dilewati — cleaning scripts tidak ditemukan.\n"
            "       Menggunakan file cleaned yang sudah ada (askrindo_output_*.xlsx).",
            flush=True,
        )

    # ----------------------------------------------------------------
    # Phase 2: Suspend Matching (sequential)
    # ----------------------------------------------------------------
    print("\n>> Phase 2: Suspend Matching ASKRINDO ...", flush=True)
    print("-" * 80, flush=True)

    for m_step in MATCHING_STEPS:
        res = execute_step(m_step)
        summary_results.append(res)
        if "FAIL" in res["status"]:
            print(f"\n[CRITICAL] Matching {m_step['name']} gagal! Pipeline dihentikan.", flush=True)
            sys.exit(1)

    # ----------------------------------------------------------------
    # Ringkasan
    # ----------------------------------------------------------------
    t_total = time.perf_counter() - t_total_start

    print("\n" + "=" * 80, flush=True)
    print("  RINGKASAN EKSEKUSI PIPELINE — ASKRINDO", flush=True)
    print("=" * 80, flush=True)
    print(f"{'No':<4} {'Nama Step':<35} {'Status':<20} {'Size (MB)':<10} {'Waktu (s)':<10}")
    print("-" * 80, flush=True)

    summary_results.sort(key=lambda x: x["step"])
    for item in summary_results:
        print(
            f"{item['step']:<4} {item['name']:<35} {item['status']:<20} "
            f"{item['size_mb']:<10.2f} {item['elapsed']:<10.1f}"
        )

    print(f"Total Waktu Eksekusi: {t_total:.1f} detik ({t_total/60:.1f} menit)", flush=True)
    print("Output file(s):", flush=True)
    print("  - V1 (Row Asli) : data/final_output_v1.xlsx", flush=True)
    print("  - V2 (Akumulasi): data/final_output_v2.xlsx", flush=True)
    print("=" * 80, flush=True)


if __name__ == "__main__":
    run_pipeline()
