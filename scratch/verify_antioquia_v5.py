import pandas as pd

output_file = 'd:/ICBF/cost-tracking/data/insumos 28 abril/consolidacion_matriz_hcb_28042026.xlsx'

try:
    df_matriz = pd.read_excel(output_file, sheet_name='MATRIZ_CALCULADA')
    ant = df_matriz[df_matriz['REGIONAL'].str.contains('ANTIOQUIA', case=False, na=False)]
    print(f"Total rows for Antioquia in MATRIZ_CALCULADA: {len(ant)}")
    print("\nFirst 10 rows of Antioquia:")
    print(ant[['REGIONAL', 'REFERENCIA / No. CONTRATO SECOP', 'CUPOS']].head(10))
    
    # Check if duplicates have NaN in CUPOS
    first_contract = ant['REFERENCIA / No. CONTRATO SECOP'].iloc[0]
    contract_rows = ant[ant['REFERENCIA / No. CONTRATO SECOP'] == first_contract]
    print(f"\nRows for contract {first_contract}:")
    print(contract_rows[['REGIONAL', 'REFERENCIA / No. CONTRATO SECOP', 'CUPOS']])

except Exception as e:
    print(f"Error: {e}")
