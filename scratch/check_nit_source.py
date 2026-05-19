import pandas as pd
import os

base_dir = 'd:/ICBF/cost-tracking/data/insumos 28 abril'
target_file = 'Antioquia_Matriz_Adicin_Nivelacion_Canasta_ HCB_17042026_107CONT.xlsm'
file_path = None
for root, dirs, files in os.walk(base_dir):
    for f in files:
        if 'Antioquia' in f and 'HCB' in f:
            file_path = os.path.join(root, f)
            break
    if file_path: break

if not file_path:
    print("File not found.")
    exit()

xls = pd.ExcelFile(file_path)
df = pd.read_excel(xls, sheet_name='MATRIZ', header=1)

nit_col = [c for c in df.columns if 'NIT' == str(c).upper().strip()][0]
print(f"NIT column: {nit_col}")
print(f"Non-null NITs: {df[nit_col].notnull().sum()} / {len(df)}")
print(f"Example NITs:\n{df[nit_col].head(20)}")
