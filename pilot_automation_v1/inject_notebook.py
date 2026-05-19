import json
import os

NOTEBOOK_PATH = "ORQUESTADOR_PILOTO_ICBF.ipynb"

with open(NOTEBOOK_PATH, 'r', encoding='utf-8') as f:
    nb = json.load(f)

mk_cell = {
    'cell_type': 'markdown',
    'metadata': {},
    'source': [
        "## 1.5. Generación Automática de Plantillas (Motor COM)\n",
        "\n",
        "Esta celda toma el archivo maestro de **Zonificación** y la plantilla **MATRIZ SERVICIOS INTEGRALES**, y genera un archivo por cada regional.\n",
        "Utiliza Microsoft Excel de forma nativa (`win32com`) para garantizar que:\n",
        "1. Las contraseñas de protección sean manejadas automáticamente.\n",
        "2. Las hojas y tablas queden completamente desbloqueadas para el usuario.\n",
        "3. Las conexiones a datos externos y formatos se conserven intactos (sin alertas al abrir).\n",
        "4. Los filtros residuales de Bolívar se limpien correctamente."
    ]
}

code_str = '''import pandas as pd
import win32com.client as win32
import os

def generate_templates():
    print("Iniciando Generación de Plantillas Nativas (COM)...", flush=True)
    master_zonificacion = os.path.join(CONFIG['BASE_DIR'], "master_data", "ZonificacIón_Primera_Infancia_Hasta_Oct.xlsx")
    template_file = os.path.join(CONFIG['BASE_DIR'], "master_data", "MATRIZ SERVICIOS INTEGRALES 2026 REGIONAL BOLIVAR_IP (2).xlsx")
    
    xl_master = pd.ExcelFile(master_zonificacion)
    
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
    
    excel = win32.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.AskToUpdateLinks = False
    excel.EnableEvents = False
    
    try:
        for sheet_name, cfg in MODALITY_CONFIG.items():
            print(f"\\n--- Modalidad: {sheet_name} ---", flush=True)
            df_master = pd.read_excel(xl_master, sheet_name=sheet_name, header=None)
            reg_col = cfg['regional_idx']
            df_data = df_master.dropna(subset=[reg_col])
            unique_regionals = [r for r in df_data[reg_col].unique() if "REGIONAL" not in str(r).upper() and "TIPO" not in str(r).upper()]
            
            for regional in unique_regionals:
                reg_clean = clean_name(regional)
                if reg_clean in ['', 'NAN']: continue
                
                output_path = os.path.join(SHARED_FOLDERS, reg_clean, f"MATRIZ_2026_{reg_clean}_{sheet_name.upper()}.xlsx")
                os.makedirs(os.path.dirname(output_path), exist_ok=True)
                print(f"  > Procesando {regional}...", flush=True)
                
                df_reg = df_data[df_data[reg_col] == regional]
                num_rows = len(df_reg)
                
                wb = excel.Workbooks.Open(template_file, UpdateLinks=0, ReadOnly=False)
                
                try:
                    wb.Unprotect("#APOLO1704*")
                    for s in wb.Sheets:
                        try: s.Unprotect("#APOLO1704*")
                        except: pass
                    
                    ws = wb.Sheets('Matriz (2)')
                    ws.Visible = -1
                    try: wb.Sheets('Matriz').Delete()
                    except: pass
                    try: ws.Name = 'Matriz'
                    except: pass
                    
                    for tbl in ws.ListObjects:
                        try:
                            if tbl.AutoFilter and tbl.AutoFilter.FilterMode:
                                tbl.AutoFilter.ShowAllData()
                        except: pass
                    
                    col_mapping = {
                        1: cfg['regional_idx'], 2: cfg['cz_idx'], 3: cfg['mun_idx'],
                        4: cfg['componente_idx'], 5: cfg['modalidad_idx'], 6: cfg['uds_name_idx'],
                        7: cfg['uds_code_idx'], 8: cfg['cupos_idx']
                    }
                    
                    block1_data, block2_data = [], []
                    for idx, row_data in df_reg.iterrows():
                        b1 = [row_data[col_mapping[c]] if c in col_mapping and pd.notna(row_data[col_mapping[c]]) else "" for c in range(1, 15)]
                        block1_data.append(b1)
                        block2_data.append([""] * 7)
                    
                    start_row = 3
                    end_row = start_row + num_rows - 1
                    
                    ws.Range(ws.Cells(start_row, 1), ws.Cells(end_row, 14)).Value = block1_data
                    ws.Range(ws.Cells(start_row, 16), ws.Cells(end_row, 22)).Value = block2_data
                    
                    clear_start = start_row + num_rows
                    if clear_start <= 3785:
                        ws.Range(f'A{clear_start}:N3785').ClearContents()
                        ws.Range(f'P{clear_start}:V3785').ClearContents()
                    
                    wb.SaveAs(output_path)
                finally:
                    wb.Close(False)
                
                print(f"    Guardado: {os.path.basename(output_path)}", flush=True)
                
    finally:
        excel.Quit()

# Ejecutar la generación de plantillas:
generate_templates()'''

code_cell = {
    'cell_type': 'code',
    'execution_count': None,
    'metadata': {},
    'outputs': [],
    'source': [line + '\n' for line in code_str.split('\n')]
}

# The setup_environment function is at index 4
# So we insert after index 4 -> index 5 is mk, 6 is code
nb['cells'].insert(5, mk_cell)
nb['cells'].insert(6, code_cell)

with open(NOTEBOOK_PATH, 'w', encoding='utf-8') as f:
    json.dump(nb, f, indent=1)

print("Inyectado con éxito")
