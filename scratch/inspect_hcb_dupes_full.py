import pandas as pd

output_file = 'd:/ICBF/cost-tracking/data/insumos 28 abril/consolidacion_matriz_hcb_28042026.xlsx'

try:
    df_matriz = pd.read_excel(output_file, sheet_name='MATRIZ_CALCULADA')
    ant = df_matriz[df_matriz['REFERENCIA / No. CONTRATO SECOP'].astype(str).str.contains('5018062024')]
    print("Full details for contract 5018062024 rows:")
    for i, row in ant.iterrows():
        print(f"\nRow {i}:")
        for col in ['REGIONAL', 'REFERENCIA / No. CONTRATO SECOP', 'SERVICIO 2026', 'CUPOS', 'Archivo_Origen']:
            print(f"  {col}: |{row[col]}|")
    
except Exception as e:
    print(f"Error: {e}")
