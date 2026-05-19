import os
import pandas as pd
import win32com.client as win32

BASE_DIR = r"D:\ICBF\cost-tracking\pilot_automation_v1"
MASTER_FILE = os.path.join(BASE_DIR, "master_data", "ZonificacIón_Primera_Infancia_Hasta_Oct.xlsx")
TEMPLATE_FILE = os.path.join(BASE_DIR, "master_data", "MATRIZ SERVICIOS INTEGRALES 2026 REGIONAL BOLIVAR_IP (2).xlsx")
OUTPUT_DIR = os.path.join(BASE_DIR, "shared_folders")

def clean_name(name):
    return str(name).strip().upper().replace('Á', 'A').replace('É', 'E').replace('Í', 'I').replace('Ó', 'O').replace('Ú', 'U').replace('Ñ', 'N')

def generate_templates():
    print("Starting COM Native Excel Generation Process...", flush=True)
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
    
    # Initialize Excel COM
    excel = win32.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.AskToUpdateLinks = False
    excel.EnableEvents = False
    
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
                
                df_reg = df_data[df_data[reg_col] == regional]
                num_rows = len(df_reg)
                
                # Open template natively
                wb = excel.Workbooks.Open(TEMPLATE_FILE, UpdateLinks=0, ReadOnly=False)
                
                try:
                    # Unprotect Workbook and Sheets
                    wb.Unprotect("#APOLO1704*")
                    for s in wb.Sheets:
                        try:
                            s.Unprotect("#APOLO1704*")
                        except Exception:
                            pass
                    
                    # We use Matriz (2) which already has 3783 rows
                    ws = wb.Sheets('Matriz (2)')
                    
                    # Make Matriz (2) visible before deleting the original
                    ws.Visible = -1 # xlSheetVisible
                    
                    # Delete original Matriz to completely remove Bolivar data traces
                    try:
                        wb.Sheets('Matriz').Delete()
                    except Exception:
                        pass
                    
                    # Rename Matriz (2) to Matriz
                    try:
                        ws.Name = 'Matriz'
                    except Exception:
                        pass
                    
                    # Clear any existing filters in the tables so the user sees all their data
                    for tbl in ws.ListObjects:
                        try:
                            if tbl.AutoFilter and tbl.AutoFilter.FilterMode:
                                tbl.AutoFilter.ShowAllData()
                        except Exception:
                            pass
                    
                    col_mapping = {
                        1: cfg['regional_idx'], 2: cfg['cz_idx'], 3: cfg['mun_idx'],
                        4: cfg['componente_idx'], 5: cfg['modalidad_idx'], 6: cfg['uds_name_idx'],
                        7: cfg['uds_code_idx'], 8: cfg['cupos_idx']
                    }
                    
                    unmapped_cols = [9, 10, 11, 12, 13, 14, 16, 17, 18, 19, 20, 21, 22]
                    
                    # Prepare 2D array
                    block1_data = []
                    block2_data = []
                    
                    for idx, row_data in df_reg.iterrows():
                        # Block 1: Col 1 to 14
                        b1 = []
                        for c in range(1, 15):
                            if c in col_mapping:
                                val = row_data[col_mapping[c]]
                                b1.append(val if pd.notna(val) else "")
                            else:
                                b1.append("")
                        block1_data.append(b1)
                        
                        # Block 2: Col 16 to 22
                        b2 = []
                        for c in range(16, 23):
                            b2.append("") # All unmapped
                        block2_data.append(b2)
                    
                    start_row = 3
                    end_row = start_row + num_rows - 1
                    
                    # Write blocks natively and FAST
                    ws.Range(ws.Cells(start_row, 1), ws.Cells(end_row, 14)).Value = block1_data
                    ws.Range(ws.Cells(start_row, 16), ws.Cells(end_row, 22)).Value = block2_data
                    
                    # Clear unused rows quickly
                    clear_start = start_row + num_rows
                    if clear_start <= 3785:
                        ws.Range(f'A{clear_start}:N3785').ClearContents()
                        ws.Range(f'P{clear_start}:V3785').ClearContents()
                    
                    # Save and close
                    wb.SaveAs(output_path)
                finally:
                    wb.Close(False)
                
                print(f"    Saved: {os.path.basename(output_path)}", flush=True)
                
    finally:
        excel.Quit()

if __name__ == "__main__":
    generate_templates()
