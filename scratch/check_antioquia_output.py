import pandas as pd

output_file = 'd:/ICBF/cost-tracking/data/insumos 28 abril/consolidacion_matriz_hcb_28042026.xlsx'

try:
    print("Checking RESUMEN_EJECUTIVO:")
    df_res = pd.read_excel(output_file, sheet_name='RESUMEN_EJECUTIVO')
    print(df_res[df_res['REGIONAL'].str.contains('ANTIOQUIA', case=False, na=False)])
    
    print("\nChecking first 10 rows of Antioquia in MATRIZ_CALCULADA:")
    df_matriz = pd.read_excel(output_file, sheet_name='MATRIZ_CALCULADA')
    ant_matriz = df_matriz[df_matriz['REGIONAL'].str.contains('ANTIOQUIA', case=False, na=False)]
    print(ant_matriz[['REGIONAL', 'REFERENCIA / No. CONTRATO SECOP', 'CUPOS']].head(10))
    print(f"Total Cupos for Antioquia in MATRIZ_CALCULADA: {ant_matriz['CUPOS'].sum()}")

except Exception as e:
    print(f"Error: {e}")
