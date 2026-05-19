import pandas as pd
import os

file_path = 'd:/ICBF/cost-tracking/data/insumos 28 abril/UDS_28042026_10AM.xlsx'
if os.path.exists(file_path):
    xls = pd.ExcelFile(file_path)
    print(f"Sheets: {xls.sheet_names}")
    for sheet in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name=sheet, nrows=10)
        print(f"\n--- Sheet: {sheet} ---")
        print(df.head(5))
else:
    print(f"File not found: {file_path}")
