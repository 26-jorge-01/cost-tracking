import pandas as pd
import os

output_file = 'd:/ICBF/cost-tracking/data/insumos 28 abril/consolidacion_matriz_hcb_28042026.xlsx'
xls = pd.ExcelFile(output_file)
df = pd.read_excel(xls, sheet_name='MATRIZ_CALCULADA')

# Find a regional and contract with nulls
null_rows = df[df['VALOR A ADICIONAR'].isnull()]
if not null_rows.empty:
    sample = null_rows.iloc[0]
    reg = sample['REGIONAL']
    cont = sample['REFERENCIA / No. CONTRATO SECOP']
    orig = sample['Archivo_Origen']
    print(f"Sample Null Contract: {cont} in {reg} from {orig}")
    
    # We need to find where this file is
    base_dir = 'd:/ICBF/cost-tracking/data/insumos 28 abril'
    found_path = None
    for root, dirs, files in os.walk(base_dir):
        if orig in files:
            found_path = os.path.join(root, orig)
            break
    
    if found_path:
        print(f"Found source: {found_path}")
        xls_src = pd.ExcelFile(found_path)
        sheet_m = [s for s in xls_src.sheet_names if 'MATRIZ' == s.upper()]
        if sheet_m:
            df_src = pd.read_excel(xls_src, sheet_name=sheet_m[0], header=1) # Try header 1
            # Filter rows for this contract
            # Find the contract column
            cont_col = [c for c in df_src.columns if 'REFERENCIA' in str(c).upper()][0]
            rows = df_src[df_src[cont_col] == cont]
            print(f"Rows for contract {cont}:")
            print(rows)
            # Check for financial columns in source
            fin_cols_src = [c for c in df_src.columns if 'VALOR' in str(c).upper() and 'ADICION' in str(c).upper()]
            print(f"\nFinancial columns found: {fin_cols_src}")
            print("\nValues in these columns:")
            print(rows[fin_cols_src])
else:
    print("No nulls found.")
