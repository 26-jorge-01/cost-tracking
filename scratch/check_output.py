import pandas as pd
import os

output_file = 'd:/ICBF/cost-tracking/data/insumos 28 abril/consolidacion_matriz_hcb_28042026.xlsx'

if os.path.exists(output_file):
    xls = pd.ExcelFile(output_file)
    df = pd.read_excel(xls, sheet_name='MATRIZ_CALCULADA')
    print("Nulls for 'VALOR A ADICIONAR' per Regional:")
    nulls_by_reg = df[df['VALOR A ADICIONAR'].isnull()]['REGIONAL'].value_counts()
    print(nulls_by_reg)
    print("\nTotal rows per Regional:")
    print(df['REGIONAL'].value_counts())
else:
    print(f"File not found: {output_file}")
