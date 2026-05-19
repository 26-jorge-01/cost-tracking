import pandas as pd
import os

file_path = 'd:/ICBF/cost-tracking/data/insumos 28 abril/consolidacion_matriz_integrales_28042026_COPIA_SEGURA.xlsx'
xls = pd.ExcelFile(file_path)
df = pd.read_excel(xls, sheet_name='MATRIZ_CALCULADA')
df = df.dropna(subset=['REFERENCIA / No. CONTRATO SECOP'])
df['REFERENCIA / No. CONTRATO SECOP'] = df['REFERENCIA / No. CONTRATO SECOP'].astype(str)
df = df[df['REFERENCIA / No. CONTRATO SECOP'].str.len() > 2]

dupes = df[df.duplicated(subset=['REFERENCIA / No. CONTRATO SECOP'], keep=False)]
if not dupes.empty:
    cont = dupes['REFERENCIA / No. CONTRATO SECOP'].iloc[0]
    print(f"Rows for contract {cont} in Integrales:")
    print(df[df['REFERENCIA / No. CONTRATO SECOP'] == cont][['REFERENCIA / No. CONTRATO SECOP', 'MUNICIPIO', 'VALOR INICIAL 2026 ']])
else:
    print("No duplicate real contracts found in Integrales.")
