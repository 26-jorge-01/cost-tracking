import os
import pandas as pd

def inspect_sheet_to_list(file_path, sheet_name):
    lines = []
    lines.append("==================================================")
    lines.append(f"FILE:  {os.path.basename(file_path)}")
    lines.append(f"SHEET: {sheet_name}")
    lines.append("==================================================")
    
    if not os.path.exists(file_path):
        lines.append("ERROR: File not found.")
        return lines
        
    try:
        df = pd.read_excel(file_path, sheet_name=sheet_name, nrows=6, header=None)
        # Also print the actual DataFrame shape and columns
        lines.append(f"Shape: {df.shape}")
        for idx, row in df.iterrows():
            lines.append(f"\nRow {idx}:")
            for col_idx, val in enumerate(row):
                if pd.notna(val):
                    lines.append(f"  Col {col_idx} ({col_idx_to_letter(col_idx + 1)}): {repr(val)}")
    except Exception as e:
        lines.append(f"ERROR: {e}")
    lines.append("\n\n")
    return lines

def col_idx_to_letter(col):
    result = ""
    while col > 0:
        col, remainder = divmod(col - 1, 26)
        result = chr(65 + remainder) + result
    return result

def main():
    base_dir = r"D:\ICBF\cost-tracking\pilot_automation_v1"
    
    master_file = os.path.join(base_dir, "master_data", "ZonificacIón_Primera_Infancia_Hasta_Oct.xlsx")
    integrales_template = os.path.join(base_dir, "master_data", "MATRIZ SERVICIOS INTEGRALES 2026 REGIONAL BOLIVAR_IP (2).xlsx")
    hcb_template = os.path.join(base_dir, "master_data", "MATRIZ_ADICIONES_HCB_NIVELACION_REGIONAL_TOLIMA_26052025_09CONTV1.xlsx")
    
    all_lines = []
    all_lines.extend(inspect_sheet_to_list(master_file, "Integrales"))
    all_lines.extend(inspect_sheet_to_list(master_file, "Comunitarios"))
    all_lines.extend(inspect_sheet_to_list(integrales_template, "Matriz (2)"))
    all_lines.extend(inspect_sheet_to_list(hcb_template, "MATRIZ"))
    
    output_path = r"D:\ICBF\cost-tracking\scratch\headers_detailed.txt"
    with open(output_path, "w", encoding="utf-8") as f:
        f.write("\n".join(all_lines))
    print(f"Details written to: {output_path}")

if __name__ == "__main__":
    main()
