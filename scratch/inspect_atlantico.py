import pandas as pd
import warnings
import unicodedata

def normalize_str(s):
    if not isinstance(s, str): return str(s)
    s = s.upper().strip()
    s = ''.join(c for c in unicodedata.normalize('NFD', s) if unicodedata.category(c) != 'Mn')
    return s

warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

import sys
if len(sys.argv) > 1:
    file_path = sys.argv[1]
else:
    file_path = r'D:\ICBF\cost-tracking\data\insumos 28 abril\Elkin\HCB\CALDAS_MATRIZ_ADICIONES_HCB_NIVELACION_CANASTA_30032026V2.2.xlsm'

xls = pd.ExcelFile(file_path)
print(f"Sheets: {xls.sheet_names}")

for sheet in xls.sheet_names:
    if 'MATRIZ' in sheet.upper():
        print(f"\n--- Sheet: {sheet} ---")
        df_temp = pd.read_excel(xls, sheet_name=sheet, nrows=20, header=None)
        h_idx = None
        for idx, row in df_temp.iterrows():
            if any(isinstance(val, str) and 'REGIONAL' == normalize_str(val) for val in row.values):
                h_idx = idx
                break
        
        if h_idx is not None:
            df = pd.read_excel(xls, sheet_name=sheet, header=h_idx)
            print(f"Header index: {h_idx}")
            print(f"Columns: {df.columns.tolist()}")
            # Check for financial columns
            fin_keywords = ['VALOR', 'INICIAL', 'ADICION', 'ACTUAL']
            for col in df.columns:
                if any(kw in normalize_str(col) for kw in fin_keywords):
                    print(f"Found potential fin col: {col}")
        else:
            print("Could not find 'REGIONAL' header row.")
