import pandas as pd

output_file = 'd:/ICBF/cost-tracking/data/insumos 28 abril/consolidacion_matriz_hcb_28042026.xlsx'

try:
    df_res = pd.read_excel(output_file, sheet_name='RESUMEN_EJECUTIVO')
    ant = df_res[df_res['REGIONAL'] == 'ANTIOQUIA']
    print("Full row for ANTIOQUIA in RESUMEN_EJECUTIVO:")
    for col in ant.columns:
        print(f"{col}: {ant[col].values[0]}")

except Exception as e:
    print(f"Error: {e}")
