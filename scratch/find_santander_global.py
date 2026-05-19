import os
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

base_dir = r'D:\ICBF\cost-tracking\data\Matrices_validadas_definitivas'
found_files = []

for root, dirs, files in os.walk(base_dir):
    for f in files:
        if f.endswith(('.xlsx', '.xlsm')):
            path = os.path.join(root, f)
            try:
                xls = pd.ExcelFile(path)
                sheet_m = [s for s in xls.sheet_names if 'MATRIZ' == s.upper()]
                if sheet_m:
                    df_temp = pd.read_excel(xls, sheet_name=sheet_m[0], nrows=20, header=None)
                    h_idx = None
                    for idx, row in df_temp.iterrows():
                        if any(isinstance(val, str) and 'REGIONAL' == normalize_str(val) for val in row.values):
                            h_idx = idx
                            break
                    if h_idx is not None:
                        df = pd.read_excel(xls, sheet_name=sheet_m[0], header=h_idx, usecols=['REGIONAL'])
                        unique_regs = df['REGIONAL'].dropna().unique()
                        for r in unique_regs:
                            rn = normalize_str(r)
                            if rn == 'SANTANDER':
                                found_files.append((path, r, len(df[df['REGIONAL'] == r])))
            except Exception:
                pass

print("Files containing 'SANTANDER' (pure):")
for path, reg, count in found_files:
    print(f"- {path} ({reg}: {count} rows)")
