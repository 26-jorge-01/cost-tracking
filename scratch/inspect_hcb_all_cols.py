import pandas as pd
import unicodedata
import re

def normalize_str(s):
    if not isinstance(s, str):
        return str(s)
    s = s.upper().strip()
    s = ''.join(c for c in unicodedata.normalize('NFD', s) if unicodedata.category(c) != 'Mn')
    s_clean = re.sub(r'[^A-Z0-9]', '', s)
    return s_clean

file_path = r'D:\ICBF\cost-tracking\data\insumos 28 abril\Elkin\HCB\CALDAS_MATRIZ_ADICIONES_HCB_NIVELACION_CANASTA_30032026V2.2.xlsm'

try:
    xls = pd.ExcelFile(file_path)
    sheet_name = 'MATRIZ'
    df_temp = pd.read_excel(xls, sheet_name=sheet_name, nrows=20, header=None)
    h_idx = None
    for idx, row in df_temp.iterrows():
        if any(isinstance(val, str) and 'REGIONAL' == normalize_str(val) for val in row.values):
            h_idx = idx
            break
            
    if h_idx is not None:
        df = pd.read_excel(xls, sheet_name=sheet_name, header=h_idx)
        print("ALL COLUMNS:")
        for col in df.columns:
            print(f"- {col}")
    else:
        print("Could not find header row.")

except Exception as e:
    print(f"Error: {e}")
