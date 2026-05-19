import pandas as pd
import os

file_path = 'd:/ICBF/cost-tracking/data/insumos 28 abril/consolidacion_matriz_integrales_28042026_COPIA_SEGURA.xlsx'
xls = pd.ExcelFile(file_path)
df = pd.read_excel(xls, sheet_name='MATRIZ_CALCULADA')

print("Financial columns with data in Integrales:")
fin_keywords = ['VALOR', 'ADICION', 'REDUCCION', 'INEJECUCION', 'APORTE', 'CONTRAPARTIDA', 'SMLV']
for col in df.columns:
    if any(k in col.upper() for k in fin_keywords):
        null_count = df[col].isnull().sum()
        total = len(df)
        if total - null_count > 0:
            print(f"- {col}: {total - null_count} / {total}")
