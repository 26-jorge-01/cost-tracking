import json
import os

notebook_path = r"D:\ICBF\cost-tracking\pilot_automation_v1\ORQUESTADOR_PILOTO_ICBF.ipynb"

if not os.path.exists(notebook_path):
    print(f"Error: Notebook not found at {notebook_path}")
    exit(1)

with open(notebook_path, 'r', encoding='utf-8') as f:
    nb = json.load(f)

# Define cell replacements based on their sequence/content
# Let's write the updated sources for the cells.

# Cell index 2: CONFIGURATION
cell_2_source = """import os
import pandas as pd
import datetime
import warnings

warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

CONFIG = {
    # Directory paths
    'BASE_DIR': r'D:\ICBF\cost-tracking\pilot_automation_v1',
    'MASTER_FILE': r'master_data\ZonificacIón_Primera_Infancia_Hasta_Oct.xlsx',
    'SHARED_FOLDERS': r'shared_folders',
    'OUTPUT_DIR': r'output',
    
    # Sheet protection password (fallback default)
    'PASSWORD': '#APOLO1704*',
}

# Derived paths
MASTER_FILE = os.path.join(CONFIG['BASE_DIR'], CONFIG['MASTER_FILE'])
SHARED_FOLDERS = os.path.join(CONFIG['BASE_DIR'], CONFIG['SHARED_FOLDERS'])
OUTPUT_DIR = os.path.join(CONFIG['BASE_DIR'], CONFIG['OUTPUT_DIR'])
AUDIT_FILE = os.path.join(OUTPUT_DIR, 'PILOT_AUDIT_MASTER.csv')
CONSOLIDATED_FILE_INTEGRALES = os.path.join(OUTPUT_DIR, 'CONSOLIDATED_DATA_INTEGRALES.csv')
CONSOLIDATED_FILE_COMUNITARIOS = os.path.join(OUTPUT_DIR, 'CONSOLIDATED_DATA_COMUNITARIOS.csv')

def clean_name(name):
    return str(name).strip().upper().replace('Á', 'A').replace('É', 'E').replace('Í', 'I').replace('Ó', 'O').replace('Ú', 'U').replace('Ñ', 'N')

print("Configuración general cargada.")"""

# Cell index 4: SETUP ENVIRONMENT
cell_4_source = """def setup_environment():
    if not os.path.exists(SHARED_FOLDERS):
        os.makedirs(SHARED_FOLDERS)
    
    print("Identificando regionales desde el Maestro...")
    xl_master = pd.ExcelFile(MASTER_FILE)
    
    regionals = set()
    if 'Integrales' in xl_master.sheet_names:
        df_int = pd.read_excel(xl_master, sheet_name='Integrales', header=3)
        regionals.update(df_int.iloc[:, 1].dropna().unique())
    if 'Comunitarios' in xl_master.sheet_names:
        df_com = pd.read_excel(xl_master, sheet_name='Comunitarios', header=0)
        regionals.update(df_com.iloc[:, 0].dropna().unique())
        
    count = 0
    for reg in regionals:
        reg_str = str(reg).strip()
        if "REGIONAL" in reg_str.upper() or "TIPO" in reg_str.upper() or reg_str == "":
            continue
        reg_clean = clean_name(reg)
        reg_path = os.path.join(SHARED_FOLDERS, reg_clean)
        if not os.path.exists(reg_path):
            os.makedirs(reg_path)
            count += 1
    
    print(f"Entorno listo. Se crearon {count} nuevas carpetas regionales.")

setup_environment()"""

# Cell index 6: TEMPLATE GENERATOR
cell_6_source = """import pandas as pd
import win32com.client as win32
import os

# Explicit mappings & rules for each modality
MODALITIES = {
    "Integrales": {
        "template_file": r"master_data\MATRIZ SERVICIOS INTEGRALES 2026 REGIONAL BOLIVAR_IP (2).xlsx",
        "template_sheet": "Matriz (2)",
        "sheet_to_delete": "Matriz",
        "rename_sheet_to": "Matriz",
        "start_row": 3,
        "template_max_rows": 3785,
        "group_column_idx": 1,
        "header_row": 3,
        "column_mapping": {
            "A": 1, "B": 2, "C": 3, "D": 9, "E": 8, "F": 8, "G": 0, "H": 10,
            "K": "const:1", "L": "const:", "M": 7, "N": "const:", "Q": "const:",
            "R": "const:", "S": "const:", "T": "const:", "U": "const:", "V": "const:"
        }
    },
    "Comunitarios": {
        "template_file": r"master_data\MATRIZ_ADICIONES_HCB_NIVELACION_REGIONAL_TOLIMA_26052025_09CONTV1.xlsx",
        "template_sheet": "MATRIZ",
        "sheet_to_delete": None,
        "rename_sheet_to": None,
        "start_row": 3,
        "template_max_rows": 3232,
        "group_column_idx": 0,
        "header_row": 0,
        "column_mapping": {
            "A": 0, "B": "const:", "C": "const:", "D": 1, "E": 2, "F": 7, "G": 7,
            "H": "const:", "I": 10, "J": 11, "K": 11, "L": 11, "M": 10, "N": "const:",
            "O": "const:", "Q": "const:", "R": "const:", "S": "const:", "X": 12,
            "Y": "const:2026", "AH": "const:", "AK": "const:", "AM": "const:",
            "AT": "const:", "AY": "const:", "BD": "const:", "BF": "const:",
            "CC": "const:", "CD": "const:", "CE": "const:", "CG": "const:", "CJ": "const:0.05"
        }
    }
}

def col_letter_to_index(col):
    if isinstance(col, int):
        return col
    col = col.upper()
    idx = 0
    for char in col:
        idx = idx * 26 + (ord(char) - ord('A') + 1)
    return idx

def generate_templates():
    print("Iniciando Generación de Plantillas Nativas (COM)...", flush=True)
    
    excel = win32.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.AskToUpdateLinks = False
    excel.EnableEvents = False
    
    xl_master = pd.ExcelFile(MASTER_FILE)
    
    try:
        for sheet_name, mod_cfg in MODALITIES.items():
            if sheet_name not in xl_master.sheet_names:
                print(f"Skipping modality '{sheet_name}' because it was not found in the master file.", flush=True)
                continue
            
            rel_tmpl_file = mod_cfg["template_file"]
            template_path = os.path.join(CONFIG["BASE_DIR"], rel_tmpl_file)
            
            print(f"\\n--- Modality: {sheet_name} ---", flush=True)
            print(f"  Template File: {template_path}", flush=True)
            
            header_row = mod_cfg.get("header_row")
            df_master = pd.read_excel(xl_master, sheet_name=sheet_name, header=header_row)
            
            group_col_idx = mod_cfg["group_column_idx"]
            df_data = df_master.dropna(subset=[df_master.columns[group_col_idx]])
            
            df_data = df_data.copy()
            df_data['Clean_Group'] = df_data.iloc[:, group_col_idx].apply(clean_name)
            
            unique_clean_groups = [
                g for g in df_data['Clean_Group'].unique() 
                if g not in ['', 'NAN'] and "REGIONAL" not in g and "TIPO" not in g
            ]
            
            mapped_cols_indices = {
                col_letter_to_index(tgt): src for tgt, src in mod_cfg["column_mapping"].items()
            }
            
            for group_clean in unique_clean_groups:
                output_filename = f"MATRIZ_2026_{group_clean}_{sheet_name.upper()}.xlsx"
                output_path = os.path.join(SHARED_FOLDERS, group_clean, output_filename)
                os.makedirs(os.path.dirname(output_path), exist_ok=True)
                
                print(f"  > Processing group: {group_clean} -> {output_filename}", flush=True)
                
                df_group = df_data[df_data['Clean_Group'] == group_clean]
                num_rows = len(df_group)
                
                wb = excel.Workbooks.Open(template_path, UpdateLinks=0, ReadOnly=False)
                
                try:
                    pwd = CONFIG.get("PASSWORD")
                    if pwd:
                        try: wb.Unprotect(pwd)
                        except: pass
                        for s in wb.Sheets:
                            try: s.Unprotect(pwd)
                            except: pass
                    
                    temp_sheet_name = mod_cfg.get("template_sheet")
                    if temp_sheet_name:
                        ws = wb.Sheets(temp_sheet_name)
                        ws.Visible = -1
                    else:
                        ws = wb.ActiveSheet
                    
                    del_sheet_name = mod_cfg.get("sheet_to_delete")
                    if del_sheet_name:
                        try: wb.Sheets(del_sheet_name).Delete()
                        except: pass
                            
                    rename_sheet_name = mod_cfg.get("rename_sheet_to")
                    if rename_sheet_name and temp_sheet_name:
                        try: ws.Name = rename_sheet_name
                        except: pass
                    
                    for tbl in ws.ListObjects:
                        try:
                            if tbl.AutoFilter and tbl.AutoFilter.FilterMode:
                                tbl.AutoFilter.ShowAllData()
                        except: pass
                    
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
                        
                        ws.Range(
                            ws.Cells(start_row, target_col_idx), 
                            ws.Cells(end_row, target_col_idx)
                        ).Value = col_values_2d
                    
                    clear_start = start_row + num_rows
                    template_max_rows = mod_cfg["template_max_rows"]
                    if clear_start <= template_max_rows:
                        for target_col_idx in mapped_cols_indices.keys():
                            ws.Range(
                                ws.Cells(clear_start, target_col_idx), 
                                ws.Cells(template_max_rows, target_col_idx)
                            ).ClearContents()
                    
                    wb.SaveAs(output_path, FileFormat=51)
                finally:
                    wb.Close(False)
                    
    finally:
        excel.Quit()
        print("Finalizado con éxito.")

generate_templates()"""

# Cell index 8: BUSINESS RULES
cell_8_source = """def rule_uds_count(df_reg, df_master_reg, modality):
    if len(df_reg) != len(df_master_reg):
        return False, f"Diferencia en UDS: Maestro {len(df_master_reg)} vs Archivo {len(df_reg)}"
    return True, "OK"

def rule_cupo_sum(df_reg, df_master_reg, modality):
    # In both maestro sheets, Col 10 (index 10) is "Cupos"
    master_sum = df_master_reg.iloc[:, 10].sum()
    
    # Identify cupos column in the template output:
    # Integrales Col H (index 7), Comunitarios Col I (index 8)
    cupos_col_idx = 7 if modality == "Integrales" else 8
    file_sum = pd.to_numeric(df_reg.iloc[:, cupos_col_idx], errors='coerce').sum()
    
    if abs(master_sum - file_sum) > 0.1:
        return False, f"Diferencia en Cupos: Maestro {master_sum} vs Archivo {file_sum}"
    return True, "OK"

VALIDATION_RULES = [
    ('Conteo_UDS', rule_uds_count),
    ('Suma_Cupos', rule_cupo_sum)
]

print("Reglas de negocio dinámicas cargadas.")"""

# Cell index 10: ORCHESTRATION & HARVESTER
cell_10_source = """def run_orchestration():
    xl_master = pd.ExcelFile(MASTER_FILE)
    
    df_master_int = pd.read_excel(xl_master, sheet_name='Integrales', header=3)
    df_master_int['Regional_Clean'] = df_master_int.iloc[:, 1].apply(clean_name)
    
    df_master_com = pd.read_excel(xl_master, sheet_name='Comunitarios', header=0)
    df_master_com['Regional_Clean'] = df_master_com.iloc[:, 0].apply(clean_name)
    
    audit_results = []
    consolidated_int = []
    consolidated_com = []
    
    all_regionals = sorted(list(set(df_master_int['Regional_Clean'].unique()) | set(df_master_com['Regional_Clean'].unique())))
    
    for reg in all_regionals:
        if reg in ['', 'NAN']:
            continue
            
        reg_folder = os.path.join(SHARED_FOLDERS, reg)
        
        for modality, df_master_all, sheet_name in [('Integrales', df_master_int, 'Matriz'), ('Comunitarios', df_master_com, 'MATRIZ')]:
            df_m_reg = df_master_all[df_master_all['Regional_Clean'] == reg]
            if len(df_m_reg) == 0:
                continue
                
            status = "PENDING"
            file_found = f"MATRIZ_2026_{reg}_{modality.upper()}.xlsx"
            file_path = os.path.join(reg_folder, file_found)
            val_logs = []
            
            if os.path.exists(file_path):
                try:
                    df_r = pd.read_excel(file_path, sheet_name=sheet_name, header=1)
                    df_r = df_r.dropna(subset=[df_r.columns[0]])
                    
                    all_ok = True
                    for r_name, r_func in VALIDATION_RULES:
                        ok, msg = r_func(df_r, df_m_reg, modality)
                        val_logs.append(f"[{r_name}]: {msg}")
                        if not ok: 
                            all_ok = False
                    
                    status = "VALIDATED" if all_ok else "ERROR"
                    if modality == 'Integrales':
                        consolidated_int.append(df_r)
                    else:
                        consolidated_com.append(df_r)
                except Exception as e:
                    status = f"ERROR: {str(e)}"
            else:
                status = "MISSING"
                
            audit_results.append({
                'Regional': reg,
                'Modalidad': modality,
                'Estado': status,
                'Archivo': file_found,
                'Log_Validacion': " | ".join(val_logs),
                'Fecha_Procesamiento': datetime.datetime.now().strftime('%Y-%m-%d %H:%M')
            })
            
    if not os.path.exists(OUTPUT_DIR):
        os.makedirs(OUTPUT_DIR)
        
    df_audit = pd.DataFrame(audit_results)
    df_audit.to_csv(AUDIT_FILE, index=False)
    
    if consolidated_int:
        pd.concat(consolidated_int, ignore_index=True).to_csv(CONSOLIDATED_FILE_INTEGRALES, index=False)
    if consolidated_com:
        pd.concat(consolidated_com, ignore_index=True).to_csv(CONSOLIDATED_FILE_COMUNITARIOS, index=False)
        
    return df_audit

resumen = run_orchestration()
print("Orquestación y auditoría completada.")"""

# Cell index 12: EXECUTIVE REPORT
cell_12_source = """if 'display' not in globals():
    def display(df):
        print(df.to_string())

print("--- RESUMEN EJECUTIVO DE AVANCES ---")
print(f"Total Evaluaciones: {len(resumen)}")
print(f"Validadas: {len(resumen[resumen['Estado'] == 'VALIDATED'])}")
print(f"Con Errores: {len(resumen[resumen['Estado'].str.contains('ERROR')])}")
print(f"Pendientes (Missing): {len(resumen[resumen['Estado'] == 'MISSING'])}")

print("\\nDesglose por Modalidad:")
print(resumen.groupby(['Modalidad', 'Estado']).size().unstack(fill_value=0))

if len(resumen[resumen['Estado'].str.contains('ERROR')]) > 0:
    print("\\nDetalle de Errores:")
    display(resumen[resumen['Estado'].str.contains('ERROR')][['Regional', 'Modalidad', 'Estado', 'Log_Validacion']])"""

# Helper to format lines for notebook cell source
def to_cell_source(txt):
    return [line + '\n' for line in txt.split('\n')]

# Modify the loaded notebook JSON object
nb['cells'][2]['source'] = to_cell_source(cell_2_source)
nb['cells'][4]['source'] = to_cell_source(cell_4_source)
nb['cells'][6]['source'] = to_cell_source(cell_6_source)
nb['cells'][8]['source'] = to_cell_source(cell_8_source)
nb['cells'][10]['source'] = to_cell_source(cell_10_source)
nb['cells'][12]['source'] = to_cell_source(cell_12_source)

# Write the updated notebook back
with open(notebook_path, 'w', encoding='utf-8') as f:
    json.dump(nb, f, indent=1)

print("Notebook updated successfully.")
