import os
import pandas as pd

def inspect_file(file_path, sheet_name):
    print(f"\n==================================================")
    print(f"FILE:  {os.path.basename(file_path)}")
    print(f"SHEET: {sheet_name}")
    print(f"==================================================")
    
    if not os.path.exists(file_path):
        print("ERROR: File not found.")
        return
        
    try:
        # Read the first 5 rows without header to see what is actually in rows 0, 1, 2...
        df = pd.read_excel(file_path, sheet_name=sheet_name, nrows=6, header=None)
        for idx, row in df.iterrows():
            row_vals = [f"Col {col_idx}: '{val}'" for col_idx, val in enumerate(row) if pd.notna(val)]
            print(f"Row {idx}: {row_vals[:15]}") # Print first 15 columns with content
    except Exception as e:
        print(f"ERROR: {e}")

def main():
    base_dir = r"D:\ICBF\cost-tracking\pilot_automation_v1"
    
    master_file = os.path.join(base_dir, "master_data", "ZonificacIón_Primera_Infancia_Hasta_Oct.xlsx")
    integrales_template = os.path.join(base_dir, "master_data", "MATRIZ SERVICIOS INTEGRALES 2026 REGIONAL BOLIVAR_IP (2).xlsx")
    hcb_template = os.path.join(base_dir, "master_data", "MATRIZ_ADICIONES_HCB_NIVELACION_REGIONAL_TOLIMA_26052025_09CONTV1.xlsx")
    
    inspect_file(master_file, "Integrales")
    inspect_file(master_file, "Comunitarios")
    inspect_file(integrales_template, "Matriz (2)")
    inspect_file(hcb_template, "MATRIZ")

if __name__ == "__main__":
    main()
