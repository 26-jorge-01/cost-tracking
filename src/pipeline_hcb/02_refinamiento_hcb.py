import os
import shutil
import pandas as pd
import warnings
import re

warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

# Configuración de rutas
base_dir = 'd:/ICBF/cost-tracking/data/insumos 28 abril'
original_file = os.path.join(base_dir, 'consolidacion_matriz_hcb_28042026.xlsx')
output_file = os.path.join(base_dir, 'consolidacion_matriz_hcb_28042026_COPIA_SEGURA.xlsx')

if not os.path.exists(original_file):
    print(f"Error: No se encuentra el archivo base {original_file}. Ejecuta primero el paso 01.")
    exit()

shutil.copy(original_file, output_file)
print(f"Refinando archivo HCB (Estructura Análoga): {output_file}")

xls_orig = pd.ExcelFile(original_file)
df_matriz_orig = pd.read_excel(xls_orig, sheet_name='MATRIZ_CALCULADA')
df_zoni_orig = pd.read_excel(xls_orig, sheet_name='ZONIFICACIÓN- PEDAGOGICO')

regionales_existentes = set(df_matriz_orig['REGIONAL'].dropna().astype(str).str.upper().str.strip())

def clean_regional(name):
    if not isinstance(name, str): return str(name)
    return name.upper().strip().replace('Á', 'A').replace('É', 'E').replace('Í', 'I').replace('Ó', 'O').replace('Ú', 'U')

regionales_existentes = {clean_regional(r) for r in regionales_existentes}

nuevos_m = []
archivos_ignorados = [os.path.basename(original_file), os.path.basename(output_file)]

for root, dirs, files in os.walk(base_dir):
    for f in files:
        f_lower = f.lower()
        if (f.endswith(('.xlsx', '.xls', '.xlsm')) and not f.startswith('~$') and 
            'hcb' in f_lower and 'alimento' not in f_lower and f not in archivos_ignorados):
            
            path = os.path.join(root, f)
            try:
                xls = pd.ExcelFile(path)
                sheet_m = [s for s in xls.sheet_names if 'MATRIZ' == s.upper()]
                if sheet_m:
                    df_temp = pd.read_excel(xls, sheet_name=sheet_m[0], nrows=10, header=None)
                    h_idx = None
                    for idx, row in df_temp.iterrows():
                        if any(isinstance(val, str) and 'REGIONAL' == val.strip().upper() for val in row.values):
                            h_idx = idx
                            break
                    if h_idx is not None:
                        df = pd.read_excel(xls, sheet_name=sheet_m[0], header=h_idx)
                        col_r = [c for c in df.columns if 'REGIONAL' == str(c).strip().upper()][0]
                        regs_in_file = df[col_r].dropna().unique()
                        regs_to_add = [r for r in regs_in_file if clean_regional(r) not in regionales_existentes]
                        
                        if regs_to_add:
                            print(f"  [NUEVO] Agregando {regs_to_add} desde {f}")
                            df = df[df[col_r].isin(regs_to_add)].copy()
                            df['Archivo_Origen'] = f
                            nuevos_m.append(df)
                            for r in regs_to_add: regionales_existentes.add(clean_regional(r))
            except Exception as e:
                pass

if nuevos_m:
    df_m_final = pd.concat([df_matriz_orig] + nuevos_m, ignore_index=True)
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        df_m_final.to_excel(writer, sheet_name='MATRIZ_CALCULADA', index=False)
        df_zoni_orig.to_excel(writer, sheet_name='ZONIFICACIÓN- PEDAGOGICO', index=False)
    print(f"Proceso completado. Archivo: {output_file}")
else:
    print("No se encontraron nuevas regionales.")
