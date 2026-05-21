import os
import pandas as pd
import win32com.client as win32

# ==============================================================================
# GENERAL PARAMETERS CONFIGURATION
# ==============================================================================
CONFIG = {
    # Directory paths (supports absolute paths or paths relative to BASE_DIR)
    "BASE_DIR": r"D:\ICBF\cost-tracking\pilot_automation_v1",
    "MASTER_FILE": r"master_data\ZonificacIón_Primera_Infancia_Hasta_Oct.xlsx",
    "OUTPUT_DIR": r"shared_folders",
    
    # Sheet protection password (fallback default)
    "PASSWORD": "#APOLO1704*",
}

# ==============================================================================
# EXPLICIT FIELD MAPPING AND CONFIGURATION BY MODALITY
# ==============================================================================
# Each modality specifies its own template file, sheet rules, and column mappings.
# Target column letters (A, B, C...) map to source column indices or column names.
MODALITIES = {
    "Integrales": {
        # File and sheet rules for Integrales
        "template_file": r"master_data\MATRIZ SERVICIOS INTEGRALES 2026 REGIONAL BOLIVAR_IP (2).xlsx",
        "template_sheet": "Matriz (2)",  # The template sheet containing formulas to populate
        "sheet_to_delete": "Matriz",            # Delete old Bolivar data sheet
        "rename_sheet_to": "Matriz",   # Rename populated sheet to 'Matriz'
        "start_row": 3,                         # Row index where data insertion starts (1-indexed)
        "template_max_rows": 3785,              # Max rows of formatted cells in template
        
        # Mapping rules
        "group_column_idx": 1,              # Column index to filter by (e.g. Regional)
        "header_row": None,                 # Row index for headers in master (None if no header row)
        "column_mapping": {
            "A": 0,   # ZONA 2026. <- TIPO (Col 0)
            "B": 1,   # REGIONAL <- Regional UDS (Col 1)
            "C": 2,   # CENTRO ZONAL <- Centro Zonal UDS (Col 2)
            "D": 3,   # MUNICIPIO <- Municipio UDS (Col 3)
            "E": 9,   # Componente para la UDS <- Componente para la UDS (Col 9)
            "F": 8,   # SERVICIO SIM <- SERVICIO 2026 (Col 8)
            "G": 8,   # NOMBRE DEL SERVICIO <- SERVICIO 2026 (Col 8)
            "H": 10,  # ATENCIONES <- Cupos (Col 10)
            "K": "const:1",  # UDS (each row represents 1 UDS)
            "L": "const:",   # Meses (Clear old Bolivar months)
            "M": 7,   # Forma de Contratación <- LITERAL DE CONTRATACION (Col 7)
            "N": "const:",   # NIT CONTRATISTA 2026 (Clear old Bolivar NIT)
            "Q": "const:",   # Fecha_FIn (Clear old Bolivar dates)
            "R": "const:",   # Valor 2025 (Clear)
            "S": "const:",   # Valor 2026 (Clear)
            "T": "const:",   # Aporte ICBF (Clear)
            "U": "const:",   # VALOR INICIAL OTROS CONCEPTOS (Clear)
            "V": "const:",   # CANTIDAD ATENCIONES QUE APLICA TASAS COMPENSATORIAS (Clear)
        }
    },
    "Comunitarios": {
        # File and sheet rules for Comunitarios (HCB)
        "template_file": r"master_data\MATRIZ_ADICIONES_HCB_NIVELACION_REGIONAL_TOLIMA_26052025_09CONTV1.xlsx",
        "template_sheet": "MATRIZ",       # Write directly into the main MATRIZ sheet
        "sheet_to_delete": None,           # No sheet needs deleting
        "rename_sheet_to": None,          # No sheet needs renaming
        "master_sheet": "Comunitarios_agrupado",  # Real sheet name in the master file
        "start_row": 3,                         # Row index where data insertion starts (1-indexed)
        "template_max_rows": 3232,              # Max rows of formatted cells in Tolima HCB template
        
        # Mapping rules
        "group_column_idx": 0,              # Column index to filter by
        "header_row": None,
        "column_mapping": {
            "A": "const:",  # (nueva, vacía por regional)
            "B": 3,         # ZONA <- Col D del maestro
            "C": 0,         # REGIONAL <- Regional UDS (Col 0)
            "D": "const:",  # REFERENCIA / No. CONTRATO SECOP (Clear)
            "E": "const:",  # No. RP  (Clear)
            "F": 1,         # CENTRO ZONAL <- Centro Zonal UDS (Col 1)
            "G": 2,         # MUNICIPIO <- Municipio UDS (Col 2)
            "H": 7,         # NOMBRE DEL SERVICIO SIM (INFORMATIVO) <- Nombre Servicio (Col 7)
            "I": 7,         # NOMBRE DEL SERVICIO <- Nombre Servicio (Col 7)
            "J": "const:",  # CONSECUTIVO CONTRATO  (Clear)
            "K": 10,        # CUPOS POR UNIDAD <- Cupos (Col 10)
            "L": 11,        # CANTIDAD DE UDS <- Madres / Unds (Col 11)
            "M": 11,        # CANTIDAD DE MADRES POR UDS 2024 <- Madres / Unds (Col 11)
            "N": 11,        # CANTIDAD DE MADRES POR UDS 2025-2026 <- Madres / Unds (Col 11)
            "O": 10,        # CUPOS POR UNIDAD <- Cupos (Col 10)
            "P": "const:",  # MOMENTOS EN FAMILIA (Clear)
            "Q": "const:",  # UNIDADES TOTALES POR SERVICIO A ADICIONAR (Clear)
            "S": "const:",  # F PARA CALCULO (Clear)
            "T": "const:",  # TIEMPO INICIAL DEL CONTRATO (Clear)
            "U": "const:",  # TIEMPO A ADICIONAR (Clear)
            "Z": 12,        # BASE ORIGINAL <- Abrev (Col 12)
            "AA": "const:2026", # VIGENCIA
            "AJ": "const:", # FORMA DE CONTRATACIÓN (Clear)
            "AM": "const:", # NIT (Clear)
            "AO": "const:", # TIPO DE ORGANIZACIÓN (Clear)
            "AV": "const:", # FECHA DE FIN DEL CONTRATO (Clear)
            "BA": "const:", # VALOR INICAL DEL SERVICIO 2024-2025-2026 (Clear)
            "BF": "const:", # VALOR INICIAL OTROS CONCEPTOS (Clear)
            "BH": "const:", # VALOR INICIAL CONTRAPARTIDA (Clear)
            "CE": "const:", # CANTIDAD TALENTO... (Clear)
            "CF": "const:", # CANTIDAD TALENTO... (Clear)
            "CG": "const:", # VALOR MOMENTOS... (Clear)
            "CI": "const:", # VALOR ADICIÓN... (Clear)
            "CL": "const:0.05", # % APORTE EAS - VALOR CONTRAPARTIDA
        }
    }
}

# ==============================================================================
# CORE EXECUTION UTILITIES
# ==============================================================================
def clean_name(name):
    """Normalize regional name to create safe, clean directory names."""
    return str(name).strip().upper().replace('Á', 'A').replace('É', 'E').replace('Í', 'I').replace('Ó', 'O').replace('Ú', 'U').replace('Ñ', 'N')

def col_letter_to_index(col):
    """Convert Excel column letter (e.g. 'A', 'AB') to a 1-based column index."""
    if isinstance(col, int):
        return col
    col = col.upper()
    idx = 0
    for char in col:
        idx = idx * 26 + (ord(char) - ord('A') + 1)
    return idx

def run_flexible_generation():
    base_dir = CONFIG["BASE_DIR"]
    master_path = CONFIG["MASTER_FILE"] if os.path.isabs(CONFIG["MASTER_FILE"]) else os.path.join(base_dir, CONFIG["MASTER_FILE"])
    output_base_dir = CONFIG["OUTPUT_DIR"] if os.path.isabs(CONFIG["OUTPUT_DIR"]) else os.path.join(base_dir, CONFIG["OUTPUT_DIR"])
    
    print("Flexible Multi-Template Generator initialized.", flush=True)
    print(f"  Master File:   {master_path}")
    print(f"  Output Base:   {output_base_dir}\n")
    
    if not os.path.exists(master_path):
        print(f"Error: Master file not found at {master_path}")
        return

    # Load master file sheets
    xl_master = pd.ExcelFile(master_path)
    
    # Initialize Excel Application COM object
    print("Initializing Excel COM application...", flush=True)
    excel = win32.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.AskToUpdateLinks = False
    excel.EnableEvents = False
    
    try:
        for sheet_name, mod_cfg in MODALITIES.items():
            master_sheet = mod_cfg.get("master_sheet", sheet_name)
            if master_sheet not in xl_master.sheet_names:
                print(f"Skipping modality '{sheet_name}' because it was not found in the master file.", flush=True)
                continue
            
            # Resolve template file path for this modality
            rel_tmpl_file = mod_cfg["template_file"]
            template_path = rel_tmpl_file if os.path.isabs(rel_tmpl_file) else os.path.join(base_dir, rel_tmpl_file)
            
            print(f"\n--- Modality: {sheet_name} ---", flush=True)
            print(f"  Template File: {template_path}", flush=True)
            
            if not os.path.exists(template_path):
                print(f"  [ERROR] Template file not found: {template_path}. Skipping modality.", flush=True)
                continue
                
            # Read sheet from master file
            header_row = mod_cfg.get("header_row")
            df_master = pd.read_excel(xl_master, sheet_name=master_sheet, header=header_row)
            
            group_col_idx = mod_cfg["group_column_idx"]
            df_data = df_master.dropna(subset=[df_master.columns[group_col_idx]])
            
            # Identify unique groups
            unique_groups = [
                g for g in df_data[df_data.columns[group_col_idx]].unique() 
                if "REGIONAL" not in str(g).upper() and "TIPO" not in str(g).upper()
            ]
            
            # Explicit column mappings to 1-based indices
            mapped_cols_indices = {
                col_letter_to_index(tgt): src for tgt, src in mod_cfg["column_mapping"].items()
            }
            
            for group in unique_groups:
                group_clean = clean_name(group)
                if group_clean in ['', 'NAN']: 
                    continue
                
                output_filename = f"MATRIZ_2026_{group_clean}_{sheet_name.upper()}.xlsx"
                output_path = os.path.join(output_base_dir, group_clean, output_filename)
                os.makedirs(os.path.dirname(output_path), exist_ok=True)
                
                print(f"  > Processing group: {group} -> {output_filename}", flush=True)
                
                # Filter rows belonging to the current group
                df_group = df_data[df_data[df_data.columns[group_col_idx]] == group]
                num_rows = len(df_group)
                
                # Open template natively
                wb = excel.Workbooks.Open(template_path, UpdateLinks=0, ReadOnly=False)
                
                try:
                    # Unprotect Workbook and Sheets
                    pwd = CONFIG.get("PASSWORD")
                    if pwd:
                        try:
                            wb.Unprotect(pwd)
                        except Exception:
                            pass
                        for s in wb.Sheets:
                            try:
                                s.Unprotect(pwd)
                            except Exception:
                                pass
                    
                    # Target sheet selection and renaming
                    temp_sheet_name = mod_cfg.get("template_sheet")
                    if temp_sheet_name:
                        ws = wb.Sheets(temp_sheet_name)
                        ws.Visible = -1 # xlSheetVisible
                    else:
                        ws = wb.ActiveSheet
                    
                    del_sheet_name = mod_cfg.get("sheet_to_delete")
                    if del_sheet_name:
                        try:
                            wb.Sheets(del_sheet_name).Delete()
                        except Exception:
                            pass
                            
                    rename_sheet_name = mod_cfg.get("rename_sheet_to")
                    if rename_sheet_name and temp_sheet_name:
                        try:
                            ws.Name = rename_sheet_name
                        except Exception:
                            pass
                    
                    # Clear filters inside any Tables (ListObjects)
                    for tbl in ws.ListObjects:
                        try:
                            if tbl.AutoFilter and tbl.AutoFilter.FilterMode:
                                tbl.AutoFilter.ShowAllData()
                        except Exception:
                            pass
                    
                    # Write mapped columns
                    start_row = mod_cfg["start_row"]
                    end_row = start_row + num_rows - 1
                    
                    for target_col_idx, source_col in mapped_cols_indices.items():
                        if source_col == "" or source_col is None:
                            col_values = [""] * num_rows
                        elif isinstance(source_col, str) and source_col.startswith("const:"):
                            const_val = source_col[len("const:"):]
                            try:
                                if const_val.isdigit():
                                    const_val = int(const_val)
                                else:
                                    const_val = float(const_val)
                            except ValueError:
                                pass
                            col_values = [const_val] * num_rows
                        elif isinstance(source_col, int):
                            col_values = df_group.iloc[:, source_col].tolist()
                        else:
                            col_values = df_group[source_col].tolist()
                        
                        col_values_2d = [[v if pd.notna(v) else ""] for v in col_values]
                        
                        # Write entire column in one COM call
                        ws.Range(
                            ws.Cells(start_row, target_col_idx), 
                            ws.Cells(end_row, target_col_idx)
                        ).Value = col_values_2d
                    
                    # Clean remaining rows in the template table for mapped columns
                    clear_start = start_row + num_rows
                    template_max_rows = mod_cfg["template_max_rows"]
                    if clear_start <= template_max_rows:
                        for target_col_idx in mapped_cols_indices.keys():
                            ws.Range(
                                ws.Cells(clear_start, target_col_idx), 
                                ws.Cells(template_max_rows, target_col_idx)
                            ).ClearContents()
                    
                    # Save as XLSX (removing any macros automatically)
                    wb.SaveAs(output_path, FileFormat=51)
                finally:
                    wb.Close(False)
                    
    except Exception as e:
        print(f"\n[FATAL ERROR] Generation process failed: {e}", flush=True)
    finally:
        print("\nQuitting Excel Application...", flush=True)
        excel.Quit()
        print("Done.", flush=True)

if __name__ == "__main__":
    run_flexible_generation()
