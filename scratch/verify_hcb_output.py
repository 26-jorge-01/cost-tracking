import pandas as pd

output_file = 'd:/ICBF/cost-tracking/data/insumos 28 abril/consolidacion_matriz_hcb_28042026.xlsx'

try:
    df = pd.read_excel(output_file, sheet_name='MATRIZ_CALCULADA', nrows=5)
    print("Columns in MATRIZ_CALCULADA:")
    print(df.columns.tolist())
    
    df_zoni = pd.read_excel(output_file, sheet_name='ZONIFICACIÓN- PEDAGOGICO', nrows=5)
    print("\nColumns in ZONIFICACIÓN- PEDAGOGICO:")
    print(df_zoni.columns.tolist())
    
except Exception as e:
    print(f"Error: {e}")
