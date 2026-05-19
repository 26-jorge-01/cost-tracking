import pandas as pd
import os

output_file = 'd:/ICBF/cost-tracking/data/insumos 28 abril/consolidacion_matriz_hcb_28042026.xlsx'
if os.path.exists(output_file):
    xls = pd.ExcelFile(output_file)
    df_z = pd.read_excel(xls, sheet_name='ZONIFICACIÓN- PEDAGOGICO')
    print("Regional counts in ZONIFICACIÓN:")
    print(df_z['Regional UDS'].value_counts())
    
    df_m = pd.read_excel(xls, sheet_name='MATRIZ_CALCULADA')
    print("\nRegional counts in MATRIZ_CALCULADA:")
    print(df_m['REGIONAL'].value_counts())
else:
    print("File not found.")
