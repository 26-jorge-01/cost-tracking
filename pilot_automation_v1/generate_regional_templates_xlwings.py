import os
import pandas as pd
import xlwings as xw

BASE_DIR = r"D:\ICBF\cost-tracking\pilot_automation_v1"
MASTER_FILE = os.path.join(BASE_DIR, "master_data", "ZonificacIón_Primera_Infancia_Hasta_Oct.xlsx")
TEMPLATE_FILE = os.path.join(BASE_DIR, "master_data", "MATRIZ SERVICIOS INTEGRALES 2026 REGIONAL BOLIVAR_IP (2).xlsx")
OUTPUT_DIR = os.path.join(BASE_DIR, "shared_folders")

def clean_name(name):
    return str(name).strip().upper().replace('Á', 'A').replace('É', 'E').replace('Í', 'I').replace('Ó', 'O').replace('Ú', 'U').replace('Ñ', 'N')

def generate_templates():
    print("Starting xlwings Native Excel Generation Process...", flush=True)
    xl_master = pd.ExcelFile(MASTER_FILE)
    
    MODALITY_CONFIG = {
        'Integrales': {
            'regional_idx': 1, 'cz_idx': 2, 'mun_idx': 3,
            'uds_code_idx': 4, 'uds_name_idx': 5,
            'modalidad_idx': 8, 'componente_idx': 9, 'cupos_idx': 10,
        },
        'Comunitarios': {
            'regional_idx': 0, 'cz_idx': 1, 'mun_idx': 2,
            'uds_code_idx': 4, 'uds_name_idx': 5,
            'modalidad_idx': 7, 'componente_idx': 6, 'cupos_idx': 10,
        }
    }
    
    # Start Excel app in background
    app = xw.App(visible=False)
    app.display_alerts = False
    app.screen_updating = False
    
    try:
        for sheet_name, cfg in MODALITY_CONFIG.items():
            print(f"\n--- Modality: {sheet_name} ---", flush=True)
            df_master = pd.read_excel(xl_master, sheet_name=sheet_name, header=None)
            reg_col = cfg['regional_idx']
            df_data = df_master.dropna(subset=[reg_col])
            unique_regionals = [r for r in df_data[reg_col].unique() if "REGIONAL" not in str(r).upper() and "TIPO" not in str(r).upper()]
            
            for regional in unique_regionals:
                reg_clean = clean_name(regional)
                if reg_clean in ['', 'NAN']: continue
                
                output_path = os.path.join(OUTPUT_DIR, reg_clean, f"MATRIZ_2026_{reg_clean}_{sheet_name.upper()}.xlsx")
                os.makedirs(os.path.dirname(output_path), exist_ok=True)
                print(f"  > Processing {regional}...", flush=True)
                
                # Filter regional data
                df_reg = df_data[df_data[reg_col] == regional]
                num_rows = len(df_reg)
                
                # Open template
                wb = app.books.open(TEMPLATE_FILE, update_links=False, ignore_read_only_recommended=True)
                
                # We use Matriz (2) which already has 3783 rows, avoiding need for resizing
                ws = wb.sheets['Matriz (2)']
                
                # Write data
                # Prepare a 2D array to write in one go for performance
                # Columns 1 to 14 are A to N.
                col_mapping = {
                    1: cfg['regional_idx'], 2: cfg['cz_idx'], 3: cfg['mun_idx'],
                    4: cfg['componente_idx'], 5: cfg['modalidad_idx'], 6: cfg['uds_name_idx'],
                    7: cfg['uds_code_idx'], 8: cfg['cupos_idx']
                }
                
                start_row = 3
                for i, (_, row_data) in enumerate(df_reg.iterrows()):
                    row_idx = start_row + i
                    # Write mapped columns
                    for col_num, source_idx in col_mapping.items():
                        val = row_data[source_idx]
                        if pd.notna(val):
                            ws.range((row_idx, col_num)).value = val
                        else:
                            ws.range((row_idx, col_num)).value = ""
                    
                    # Clear unmapped diligenciable columns for this row to avoid Bolivar data leakage
                    for col_num in [9, 10, 11, 12, 13, 14, 16, 17, 18, 19, 20, 21, 22]:
                        ws.range((row_idx, col_num)).value = ""
                
                # For remaining unused rows, clear the diligenciable columns
                clear_start = start_row + num_rows
                if clear_start <= 3785:
                    # Clear columns A to N (1 to 14) and P to V (16 to 22)
                    ws.range(f'A{clear_start}:N3785').clear_contents()
                    ws.range(f'P{clear_start}:V3785').clear_contents()
                
                # Hide the original 52-row 'Matriz' and make 'Matriz (2)' visible/renamed if possible
                try:
                    wb.sheets['Matriz'].api.Visible = 2 # xlSheetVeryHidden
                    ws.name = 'Matriz'
                except Exception:
                    pass
                
                wb.save(output_path)
                wb.close()
                print(f"    Saved: {os.path.basename(output_path)}", flush=True)
                
    finally:
        app.quit()

if __name__ == "__main__":
    generate_templates()
