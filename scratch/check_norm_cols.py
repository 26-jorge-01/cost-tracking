import pandas as pd
import os
import re
import unicodedata

def normalize_str(s):
    if not isinstance(s, str): return str(s)
    s = s.upper().strip()
    s = ''.join(c for c in unicodedata.normalize('NFD', s) if unicodedata.category(c) != 'Mn')
    s_clean = re.sub(r'[^A-Z0-9]', '', s)
    return s_clean

base_dir = 'd:/ICBF/cost-tracking/data/insumos 28 abril'
target_file = 'Atlántico_Matriz_Adición_Nivelacion_Canasta_ HCB_14042026_V2.xlsm'
file_path = None
for root, dirs, files in os.walk(base_dir):
    for f in files:
        if 'Atlántico' in f and 'HCB' in f:
            file_path = os.path.join(root, f)
            break
    if file_path: break

xls = pd.ExcelFile(file_path)
df = pd.read_excel(xls, sheet_name='MATRIZ', header=1)

print("Columns in ATLANTICO (Normalized):")
for col in df.columns:
    print(f"'{col}' -> '{normalize_str(col)}'")
