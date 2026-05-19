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

files = [
    r'D:\ICBF\cost-tracking\data\Matrices_validadas_definitivas\Lady\HCB\MATRIZ_ADICIONES_HCB_NIVELACION_CANASTA_14042026_V2.xlsm',
    r'D:\ICBF\cost-tracking\data\Matrices_validadas_definitivas\Lady\HCB\NSantander_Matriz_Adición_Nivelacion_Canasta_ HCB_31032026_V2 (2).xlsm'
]

for f in files:
    try:
        xls = pd.ExcelFile(f)
        df_temp = pd.read_excel(xls, sheet_name='MATRIZ', nrows=20, header=None)
        h_idx = None
        for idx, row in df_temp.iterrows():
            if any(isinstance(val, str) and 'REGIONAL' == normalize_str(val) for val in row.values):
                h_idx = idx
                break
        if h_idx is not None:
            df = pd.read_excel(xls, sheet_name='MATRIZ', header=h_idx)
            print(f"File: {f.split('\\')[-1]}")
            print(f"  Unique Regionals: {df['REGIONAL'].unique()}")
        else:
            print(f"File: {f.split('\\')[-1]} - Header not found.")
    except Exception as e:
        print(f"Error processing {f}: {e}")
