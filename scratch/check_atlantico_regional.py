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

file_path = r'D:\ICBF\cost-tracking\data\Matrices_validadas_definitivas\Alba\ATLANTICO\Atlántico_Matriz_Adición_Nivelacion_Canasta_ HCB_14042026_V2 FINAL 1.xlsm'

try:
    xls = pd.ExcelFile(file_path)
    df_temp = pd.read_excel(xls, sheet_name='MATRIZ', nrows=20, header=None)
    h_idx = None
    for idx, row in df_temp.iterrows():
        if any(isinstance(val, str) and 'REGIONAL' == normalize_str(val) for val in row.values):
            h_idx = idx
            break
            
    if h_idx is not None:
        df = pd.read_excel(xls, sheet_name='MATRIZ', header=h_idx)
        print(f"Unique Regionals in Atlantico file:")
        print(df['REGIONAL'].unique())
        print(f"\nFirst 10 rows of REGIONAL column:")
        print(df['REGIONAL'].head(10))
    else:
        print("Header not found.")

except Exception as e:
    print(f"Error: {e}")
