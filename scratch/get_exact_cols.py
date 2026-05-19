import pandas as pd
import os
import unicodedata

def normalize_str(s):
    if not isinstance(s, str): return str(s)
    s = s.upper().strip()
    s = ''.join(c for c in unicodedata.normalize('NFD', s) if unicodedata.category(c) != 'Mn')
    return s

integ_file = 'd:/ICBF/cost-tracking/data/insumos 28 abril/consolidacion_matriz_integrales_28042026_COPIA_SEGURA.xlsx'
xls = pd.ExcelFile(integ_file)

print("MATRIZ_CALCULADA Columns:")
df_m = pd.read_excel(xls, sheet_name='MATRIZ_CALCULADA', nrows=0)
print(df_m.columns.tolist())

print("\nZONIFICACIÓN- PEDAGOGICO Columns:")
df_z = pd.read_excel(xls, sheet_name='ZONIFICACIÓN- PEDAGOGICO', nrows=0)
print(df_z.columns.tolist())
