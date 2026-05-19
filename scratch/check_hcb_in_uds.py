import pandas as pd
import os

file_path = 'd:/ICBF/cost-tracking/data/insumos 28 abril/UDS_28042026_10AM.xlsx'
df = pd.read_excel(file_path, sheet_name='UDS_28042026_10AM', nrows=100)
print(df.columns.tolist())

serv_col = [c for c in df.columns if 'SERVICIO' in c.upper() or 'MODALIDAD' in c.upper()]
if serv_col:
    print(f"\nValues in {serv_col[0]}:")
    print(df[serv_col[0]].unique())
else:
    print("\nNo servicio/modalidad column found.")
