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

file_path = r'D:\ICBF\cost-tracking\data\Matrices_validadas_definitivas\Jhoham\HCB\Antioquia_Matriz_Adición_Nivelacion_Canasta_ HCB_17042026_107CONT.xlsm'

try:
    xls = pd.ExcelFile(file_path)
    df = pd.read_excel(xls, sheet_name='MATRIZ', header=None, nrows=50)
    # Find header
    h_idx = None
    for idx, row in df.iterrows():
        if any(isinstance(val, str) and 'REGIONAL' == normalize_str(val) for val in row.values):
            h_idx = idx
            break
            
    if h_idx is not None:
        df = pd.read_excel(xls, sheet_name='MATRIZ', header=h_idx)
        cupos_col = [c for c in df.columns if 'SUMATORIA' in c.upper() and 'CUPOS' in c.upper() and 'UNICO' in c.upper()]
        if cupos_col:
            print(f"Values in {cupos_col[0]}:")
            print(df[cupos_col[0]].head(20))
            print(f"Sum of column: {df[cupos_col[0]].sum()}")
        else:
            print("Cupos column not found.")
    else:
        print("Header not found.")

except Exception as e:
    print(f"Error: {e}")
