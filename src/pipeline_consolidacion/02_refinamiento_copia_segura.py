import os
import shutil
import pandas as pd
import warnings
import re

warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

# Configuración de rutas
base_dir = 'd:/ICBF/cost-tracking/data/insumos 28 abril'
original_file = os.path.join(base_dir, 'consolidacion_matriz_integrales_28042026.xlsx')
output_file = os.path.join(base_dir, 'consolidacion_matriz_integrales_28042026_COPIA_SEGURA.xlsx')

if not os.path.exists(original_file):
    print(f"Error: No se encuentra el archivo base {original_file}. Ejecuta primero el paso 01.")
    exit()

# 1. Hacer una copia literal y leer el archivo original
shutil.copy(original_file, output_file)
print(f"Refinando archivo: {output_file}")

xls_orig = pd.ExcelFile(original_file)
df_matriz_orig = pd.read_excel(xls_orig, sheet_name='MATRIZ_CALCULADA') if 'MATRIZ_CALCULADA' in xls_orig.sheet_names else pd.DataFrame()
df_zoni_orig = pd.read_excel(xls_orig, sheet_name='ZONIFICACIÓN- PEDAGOGICO') if 'ZONIFICACIÓN- PEDAGOGICO' in xls_orig.sheet_names else pd.DataFrame()

try:
    df_hoja3_orig = pd.read_excel(xls_orig, sheet_name='Hoja3')
except:
    df_hoja3_orig = None

regionales_existentes = set()
if not df_zoni_orig.empty and 'Regional UDS' in df_zoni_orig.columns:
    regionales_existentes = set(df_zoni_orig['Regional UDS'].dropna().astype(str).str.upper().str.strip())

def clean_regional(name):
    if not isinstance(name, str):
        return str(name)
    name = name.upper().strip()
    name = re.sub(r'[Á]', 'A', name)
    name = re.sub(r'[É]', 'E', name)
    name = re.sub(r'[Í]', 'I', name)
    name = re.sub(r'[Ó]', 'O', name)
    name = re.sub(r'[Ú]', 'U', name)
    return name

regionales_existentes = {clean_regional(r) for r in regionales_existentes}

def encontrar_header(df_temp, keyword):
    for idx, row in df_temp.iterrows():
        if any(isinstance(val, str) and keyword.lower() in val.lower() for val in row.values):
            return idx
    return None

def limpiar_fechas(df):
    for col in df.columns:
        if pd.api.types.is_datetime64_any_dtype(df[col]):
            df[col] = df[col].dt.tz_localize(None).astype(str)
    return df

archivos_ignorados = [os.path.basename(original_file), os.path.basename(output_file), 'UDS_28042026_10AM.xlsx']

nuevos_matriz = []
nuevos_zoni = []
regionales_agregadas = set()

for root, dirs, files in os.walk(base_dir):
    for f in files:
        f_lower = f.lower()
        if (f.endswith(('.xlsx', '.xls', '.xlsm')) and 
            not f.startswith('~$') and 
            f not in archivos_ignorados and
            'hcb' not in f_lower and 
            'integralidad' not in f_lower):
            
            path = os.path.join(root, f)
            try:
                xls = pd.ExcelFile(path)
                
                # Procesar ZONIFICACIÓN
                sheet_match_zoni = [s for s in xls.sheet_names if 'ZONIFICAC' in s.upper()]
                if sheet_match_zoni:
                    sheet_name = sheet_match_zoni[0]
                    df_temp = pd.read_excel(xls, sheet_name=sheet_name, nrows=15, header=None)
                    h_idx = encontrar_header(df_temp, 'Regional UDS')
                    if h_idx is not None:
                        df_z = pd.read_excel(xls, sheet_name=sheet_name, header=h_idx)
                        col_r = [c for c in df_z.columns if 'Regional UDS' in str(c)][0]
                        df_z = df_z.dropna(subset=[col_r])
                        
                        regs_in_file = df_z[col_r].dropna().unique()
                        regs_to_add = [r for r in regs_in_file if clean_regional(r) not in regionales_existentes]
                        
                        if regs_to_add:
                            print(f"  [NUEVO] Agregando {regs_to_add} desde {f}")
                            df_z = df_z[df_z[col_r].isin(regs_to_add)].copy()
                            df_z = limpiar_fechas(df_z)
                            df_z = df_z.rename(columns={col_r: 'Regional UDS'})
                            df_z['Archivo_Origen'] = f
                            nuevos_zoni.append(df_z)
                            for r in regs_to_add:
                                regionales_agregadas.add(clean_regional(r))
                                regionales_existentes.add(clean_regional(r))
                                
                # Procesar MATRIZ_CALCULADA
                sheet_match_mat = [s for s in xls.sheet_names if 'MATRIZ_CALCULADA' in s.upper()]
                if sheet_match_mat:
                    sheet_name = sheet_match_mat[0]
                    df_temp = pd.read_excel(xls, sheet_name=sheet_name, nrows=15, header=None)
                    h_idx = encontrar_header(df_temp, 'REGIONAL')
                    if h_idx is not None:
                        df_m = pd.read_excel(xls, sheet_name=sheet_name, header=h_idx)
                        cols_r = [c for c in df_m.columns if 'REGIONAL' == str(c).strip().upper()]
                        if cols_r:
                            col_r = cols_r[0]
                            df_m = df_m.dropna(subset=[col_r])
                            regs_in_file = df_m[col_r].dropna().unique()
                            # Solo agregamos a la matriz si la regional fue marcada como "nueva" o relevante en este proceso
                            regs_to_add = [r for r in regs_in_file if clean_regional(r) in regionales_agregadas]
                            if regs_to_add:
                                df_m = df_m[df_m[col_r].isin(regs_to_add)].copy()
                                df_m = limpiar_fechas(df_m)
                                df_m['Archivo_Origen'] = f
                                nuevos_matriz.append(df_m)
                                
            except Exception as e:
                print(f"Error procesando {f}: {e}")

# Concatenar y guardar
df_zoni_final = pd.concat([df_zoni_orig] + nuevos_zoni, ignore_index=True) if nuevos_zoni else df_zoni_orig
df_matriz_final = pd.concat([df_matriz_orig] + nuevos_matriz, ignore_index=True) if nuevos_matriz else df_matriz_orig

with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
    if not df_matriz_final.empty:
        df_matriz_final.to_excel(writer, sheet_name='MATRIZ_CALCULADA', index=False)
    if not df_zoni_final.empty:
        df_zoni_final.to_excel(writer, sheet_name='ZONIFICACIÓN- PEDAGOGICO', index=False)
    if df_hoja3_orig is not None:
        df_hoja3_orig.to_excel(writer, sheet_name='Hoja3', index=False)

print(f"\nProceso completado.")
print(f"Nuevas regionales integradas: {regionales_agregadas if regionales_agregadas else 'Ninguna'}")
print(f"Archivo final: {output_file}")
