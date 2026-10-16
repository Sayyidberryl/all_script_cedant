import os
import sys

sys.modules['numexpr'] = None
sys.modules['bottleneck'] = None

import pandas as pd

df_facul = pd.read_excel("data/kbru_output_facul.xlsx", nrows=2)
df_osbal = pd.read_excel("data/kbru_output_osbal.xlsx", nrows=2)
df_sus = pd.read_excel("data/kbru_output_suspend.xlsx", nrows=2)

FACUL_TARGET_COLS = [
    "FAC_CODE", "FAC_CEDANT", "COMP_NAME", "FAC_BROKER", "COMP_NAME2", "BUSSINESS_PATNER",
    "FAC_INSURED", "INSURED_CLEAN_1", "INSURED_CLEAN_2", "INSURED_CLEAN_3", "INSURED_CLEAN_4", "INSURED_CLEAN_5",
    "FAC_POLICY_NO", "POLICY_CLEAN_1", "SERTIF_CLEAN_1", "POLICY_CLEAN_2", "SERTIF_CLEAN_2", "POLICY_CLEAN_3", "SERTIF_CLEAN_3",
    "POLICY_CLEAN_4", "POLICY_CLEAN_5",
    "FAC_SLIP", "SLIP_CLEAN_1", "SLIP_CLEAN_2", "SLIP_CLEAN_3", "SLIP_CLEAN_4", "SLIP_CLEAN_5",
    "FAC_CURRENCY", "FAC_INP_DATE", "FAC_COM_DATE", "FAC_EXP_DATE", "FAC_SUB_CLASS", "FAC_RISK", "FAC_DESC", "FAC_ACC_STS", "FAC_STS_SLIP"
]

OSBAL_TARGET_COLS = [
    "CCOS_DOC_NO", "CCOS_DATE", "CCOS_REF_CODE", "CCOS_COMP", "CCOS_COMP_NAME", "CCOS_REF_COMP", "CCOS_REF_COMP_NAME",
    "FAC_INSURED", "INSURED_CLEAN_1", "INSURED_CLEAN_2", "INSURED_CLEAN_3", "INSURED_CLEAN_4", "INSURED_CLEAN_5",
    "CCOS_CURR", "CCOS_OR_BAL", "CCOS_BAL_DUE", "CCOS_OR_BAL_IN_IDR", "CCOS_BAL_DUE_IN_IDR",
    "FAC_COM_DATE", "FAC_EXP_DATE", "FAC_DUE_DATES", "FAC_SUB_CLASS",
    "FAC_POLICY_NO", "POLICY_CLEAN_1", "SERTIF_CLEAN_1", "POLICY_CLEAN_2", "SERTIF_CLEAN_2", "POLICY_CLEAN_3", "SERTIF_CLEAN_3", "POLICY_CLEAN_4", "POLICY_CLEAN_5",
    "FAC_SLIP", "SLIP_CLEAN_1", "SLIP_CLEAN_2", "SLIP_CLEAN_3", "SLIP_CLEAN_4", "SLIP_CLEAN_5",
    "CLASS_CODE", "CLASS_NAME", "CLSDT_POLICY_NO", "CLSDT_SLIP_NO", "CLSDT_SERTF_NO"
]

print("=== FACUL MAPPING CHECK ===")
print("Columns in kbru_output_facul.xlsx:")
print(df_facul.columns.tolist())
missing_in_excel = [c for c in FACUL_TARGET_COLS if c not in df_facul.columns]
print("FACUL_TARGET_COLS missing in excel:", missing_in_excel)

print("\n=== OSBAL MAPPING CHECK ===")
print("Columns in kbru_output_osbal.xlsx:")
print(df_osbal.columns.tolist())
missing_in_osbal = [c for c in OSBAL_TARGET_COLS if c not in df_osbal.columns]
print("OSBAL_TARGET_COLS missing in excel:", missing_in_osbal)

print("\n=== SUSPEND COLS ===")
print(df_sus.columns.tolist())
