import pandas as pd
import os

output_file = 'd:/ICBF/cost-tracking/data/insumos 28 abril/consolidacion_matriz_hcb_28042026.xlsx'
if os.path.exists(output_file):
    xls = pd.ExcelFile(output_file)
    df = pd.read_excel(xls, sheet_name='ZONIFICACIÓN- PEDAGOGICO')
    
    print("Columns with data in ZONIFICACIÓN- PEDAGOGICO:")
    non_empty_cols = df.columns[df.notnull().any()].tolist()
    for col in non_empty_cols:
        null_count = df[col].isnull().sum()
        total = len(df)
        print(f"- {col}: {total - null_count} non-null values")
else:
    print(f"File not found: {output_file}")
