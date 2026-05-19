import pandas as pd

output_file = 'd:/ICBF/cost-tracking/data/insumos 28 abril/consolidacion_matriz_hcb_28042026.xlsx'

try:
    df_res = pd.read_excel(output_file, sheet_name='RESUMEN_EJECUTIVO')
    sant = df_res[df_res['REGIONAL'].str.contains('SANTANDER', case=False, na=False)]
    print("Summary for Santander regions:")
    print(sant[['REGIONAL', 'No. Contratos', 'Cupos Totales']])
    
    df_matriz = pd.read_excel(output_file, sheet_name='MATRIZ_CALCULADA')
    sant_matriz = df_matriz[df_matriz['REGIONAL'].str.contains('SANTANDER', case=False, na=False)]
    print(f"\nTotal rows for Santander regions in MATRIZ_CALCULADA: {len(sant_matriz)}")
    if len(sant_matriz) > 0:
        print(sant_matriz[['REGIONAL', 'REFERENCIA / No. CONTRATO SECOP', 'CUPOS', 'Archivo_Origen']].head(10))

except Exception as e:
    print(f"Error: {e}")
