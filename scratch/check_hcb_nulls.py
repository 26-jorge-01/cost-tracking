import pandas as pd
import os

output_file = 'd:/ICBF/cost-tracking/data/insumos 28 abril/consolidacion_matriz_hcb_28042026.xlsx'
xls = pd.ExcelFile(output_file)
df = pd.read_excel(xls, sheet_name='MATRIZ_CALCULADA')

# Check if nulls are always the duplicates
df['is_duplicate'] = df.duplicated(subset=['REFERENCIA / No. CONTRATO SECOP'], keep='first')
nulls = df[df['VALOR A ADICIONAR'].isnull()]
print(f"Total nulls: {len(nulls)}")
print(f"Nulls that are duplicates: {len(nulls[nulls['is_duplicate']])}")
print(f"Nulls that are NOT duplicates: {len(nulls[~nulls['is_duplicate']])}")

if len(nulls[~nulls['is_duplicate']]) > 0:
    print("\nExample of Non-Duplicate Null:")
    print(nulls[~nulls['is_duplicate']].iloc[0][['REGIONAL', 'REFERENCIA / No. CONTRATO SECOP', 'Archivo_Origen']])
