import pandas as pd
import os

output_file = 'd:/ICBF/cost-tracking/data/insumos 28 abril/consolidacion_matriz_hcb_28042026.xlsx'
xls = pd.ExcelFile(output_file)
df = pd.read_excel(xls, sheet_name='MATRIZ_CALCULADA')

df['is_duplicate'] = df.duplicated(subset=['REFERENCIA / No. CONTRATO SECOP'], keep='first')
non_dupe_nulls = df[df['VALOR A ADICIONAR'].isnull() & ~df['is_duplicate']]
print(non_dupe_nulls[['REGIONAL', 'REFERENCIA / No. CONTRATO SECOP', 'Archivo_Origen']])
