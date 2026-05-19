import pandas as pd
import os
import warnings

warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

file_path = r'D:\ICBF\cost-tracking\data\insumos 28 abril\Elkin\HCB\CALDAS_MATRIZ_ADICIONES_HCB_NIVELACION_CANASTA_30032026V2.2.xlsm'

try:
    xls = pd.ExcelFile(file_path)
    print(f"SHEETS: {xls.sheet_names}")
    
    for sheet in xls.sheet_names:
        if 'ZONIFICAC' in sheet.upper() or 'MATRIZ' in sheet.upper():
            print(f"\n--- Inspection of sheet: {sheet} ---")
            df = pd.read_excel(xls, sheet_name=sheet, nrows=10, header=None)
            print(df)
except Exception as e:
    print(f"Error: {e}")
