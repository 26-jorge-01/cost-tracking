import pandas as pd
import os

file_path = 'd:/ICBF/cost-tracking/data/insumos 28 abril/UDS_28042026_10AM.xlsx'
if os.path.exists(file_path):
    # Just read a few rows to see the content
    df = pd.read_excel(file_path, nrows=100)
    print("Columns in UDS file:")
    print(df.columns.tolist())
    
    # Check for HCB in 'Servicio' or similar column
    # Usually the column is called 'Servicio' or 'Modalidad'
    serv_col = [c for c in df.columns if 'SERVICIO' in c.upper()]
    if serv_col:
        hcb_count = df[df[serv_col[0]].astype(str).str.contains('HCB', case=False, na=False)].shape[0]
        print(f"\nHCB rows in first 100 rows: {hcb_count}")
    else:
        print("\nCould not find Servicio column.")
else:
    print(f"File not found: {file_path}")
