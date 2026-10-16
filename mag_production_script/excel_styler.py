"""
excel_styler.py — Utility to apply light purple background styling to script-added columns across all rows in Excel output files.
"""

import os
import openpyxl
from openpyxl.styles import PatternFill, Font

# Target columns for each type of script output
SCRIPT_ADDED_COLUMNS = {
    "facul_clean": [
        "mitra_bisnis",
        "BUSSINESS_PATNER",
        "COMP_NAME2",
        "prefix:clean ",
        "prefix:policy_clean",
        "prefix:slip_clean",
        "prefix:insured_clean",
        "prefix:sertif_clean",
    ],
    "osbal_clean": [
        "prefix:clean ",
        "prefix:policy_clean",
        "prefix:polis_clean",
        "prefix:slip_clean",
        "prefix:insured_clean",
        "prefix:sertif_clean",
    ],
    "suspend_clean": [
        "prefix:clean ",
        "prefix:sertif_cln",
        "POLIS_CLN",
        "SERTIF_CLN",
        "SLIP_NO_CLN",
        "INSURED_CLEAN",
        "INSURED_1",
        "INSURED_2",
    ],
    "matching": [
        "CCOS_DOC_NO",
        "CCOS_REF_CODE",
        "INSURED_ORI", "INSURED_1", "INSURED_2",
        "AMOUNT_ORI_MIN1",
        "CCOS_OR_BAL", "CCOS_BAL_DUE", "DIFERENCE",
        "FLAG_PROD",
        "POLIS_CLN", "SERTIF_CLN", "SLIP_NO_CLN",
        "POLICY_CLEAN_1", "SERTIF_CLEAN_1",
        "POLICY_CLEAN_2", "SERTIF_CLEAN_2",
        "POLICY_CLEAN_3", "SERTIF_CLEAN_3",
        "POLICY_CLEAN_4", "POLICY_CLEAN_5",
        "CEK AMOUNT DATABASE X BAL RV",
        "Mark Admin Fac Code", "Mark Admin", "Mark Admin (Status)", "Mark ARP",
        "SKENARIO",
        "prefix:clean ",
        "prefix:policy_clean",
        "prefix:slip_clean",
        "prefix:insured_clean",
        "prefix:sertif_clean",
        "prefix:sertif_cln",
    ],
}


def apply_purple_column_style(file_path: str, script_type: str, color_hex: str = "E8DAEF", header_color_hex: str = "D7BDE2") -> None:
    """
    Applies light purple fill (all rows) to columns generated / added by the script.
    - header_color_hex: Slightly darker purple for header cell ('D7BDE2')
    - color_hex: Soft light purple for all data rows ('E8DAEF')
    """
    if not os.path.exists(file_path):
        return

    # Files > 15 MB contain 100,000+ rows; openpyxl full load consumes gigabytes of RAM and corrupts zip structure
    file_size_mb = os.path.getsize(file_path) / (1024 * 1024)
    if file_size_mb > 15.0:
        print(f"  [Excel Styler] File size is {file_size_mb:.1f} MB (>15 MB). Skipping in-place styling to prevent file corruption.", flush=True)
        return

    print(f"  [Excel Styler] Applying light purple fill to script columns in {file_path} ...", flush=True)

    try:
        wb = openpyxl.load_workbook(file_path)
        ws = wb.active
    except Exception as e:
        print(f"  [Excel Styler] Failed to open {file_path}: {e}", flush=True)
        return

    header_fill = PatternFill(start_color=header_color_hex, end_color=header_color_hex, fill_type="solid")
    data_fill   = PatternFill(start_color=color_hex, end_color=color_hex, fill_type="solid")

    target_defs = SCRIPT_ADDED_COLUMNS.get(script_type, SCRIPT_ADDED_COLUMNS["matching"])

    col_indices = []
    for col_idx, cell in enumerate(ws[1], 1):
        val = str(cell.value or "").strip()
        if not val:
            continue
        val_lower = val.lower()
        matched = False
        for target in target_defs:
            if target.startswith("prefix:"):
                prefix = target[7:].lower()
                if val_lower.startswith(prefix):
                    matched = True
                    break
            elif val_lower == target.lower():
                matched = True
                break
        if matched:
            col_indices.append(col_idx)

    if not col_indices:
        wb.close()
        print(f"  [Excel Styler] No matching columns found to style.", flush=True)
        return

    max_row = ws.max_row

    # Openpyxl re-saving files with >100,000 rows corrupts ZIP headers (Bad magic number / Bad CRC-32)
    if max_row > 100000:
        wb.close()
        print(f"  [Excel Styler] File has {max_row:,} rows (>100,000). Skipping in-place styling to prevent file corruption.", flush=True)
        return

    style_all_rows = max_row <= 150000

    for col_idx in col_indices:
        header_cell = ws.cell(row=1, column=col_idx)
        header_cell.fill = header_fill
        header_cell.font = Font(bold=True)
        if style_all_rows:
            for row_idx in range(2, max_row + 1):
                ws.cell(row=row_idx, column=col_idx).fill = data_fill

    try:
        wb.save(file_path)
        row_msg = f"across all {max_row:,} rows" if style_all_rows else f"header row (file has {max_row:,} rows)"
        print(f"  [Excel Styler] Successfully styled {len(col_indices)} column(s) {row_msg}.", flush=True)
    except Exception as e:
        print(f"  [Excel Styler] Error saving styled file: {e}", flush=True)
    finally:
        wb.close()


if __name__ == "__main__":
    import sys
    if len(sys.argv) > 2:
        apply_purple_column_style(sys.argv[1], sys.argv[2])
