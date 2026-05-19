import pandas as pd

output_file = 'd:/ICBF/cost-tracking/data/insumos 28 abril/consolidacion_matriz_hcb_28042026.xlsx'

try:
    df = pd.read_excel(output_file, sheet_name='MATRIZ_CALCULADA', nrows=5)
    print("MATRIZ_CALCULADA Columns:")
    for i, col in enumerate(df.columns):
        print(f"{i+1}. {col}")
    
except Exception as e:
    print(f"Error: {e}")
