import pandas as pd
import os

output_file = 'd:/ICBF/cost-tracking/data/insumos 28 abril/consolidacion_matriz_hcb_28042026.xlsx'
if os.path.exists(output_file):
    xls = pd.ExcelFile(output_file)
    print("--- MATRIZ_CALCULADA (First 5 rows) ---")
    df_m = pd.read_excel(xls, sheet_name='MATRIZ_CALCULADA')
    print(df_m[['REGIONAL', 'REFERENCIA / No. CONTRATO SECOP', 'Archivo_Origen']].head(10))
    
    print("\n--- ZONIFICACIÓN- PEDAGOGICO (First 10 rows) ---")
    df_z = pd.read_excel(xls, sheet_name='ZONIFICACIÓN- PEDAGOGICO')
    print(df_z[['Regional UDS', 'REFERENCIA / No. CONTRATO SECOP', 'Archivo_Origen']].head(10))
else:
    print("File not found.")
