import pandas as pd
import os

file_path = 'd:/ICBF/cost-tracking/data/insumos 28 abril/consolidacion_matriz_integrales_28042026_COPIA_SEGURA.xlsx'

if os.path.exists(file_path):
    xls = pd.ExcelFile(file_path)
    print(f"Sheets: {xls.sheet_names}")
    for sheet in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name=sheet, nrows=1)
        print(f"\n--- Sheet: {sheet} ---")
        print(f"Columns: {df.columns.tolist()}")
else:
    print(f"File not found: {file_path}")
