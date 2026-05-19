import os
import pandas as pd
import warnings
import re

warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

# Configuración de rutas
base_dir = 'd:/ICBF/cost-tracking/data/insumos 28 abril'
output_file = 'd:/ICBF/cost-tracking/data/insumos 28 abril/consolidacion_matriz_integrales_28042026.xlsx'

all_dfs = []
regionales_encontradas = set()

# Lista de todas las regionales del ICBF esperadas (33 regionales)
regionales_icbf = {
    'AMAZONAS', 'ANTIOQUIA', 'ARAUCA', 'ATLANTICO', 'BOGOTA', 'BOGOTA D.C.', 'BOLIVAR', 
    'BOYACA', 'CALDAS', 'CAQUETA', 'CASANARE', 'CAUCA', 'CESAR', 'CHOCO', 'CORDOBA', 
    'CUNDINAMARCA', 'GUAINIA', 'GUAVIARE', 'HUILA', 'LA GUAJIRA', 'MAGDALENA', 'META', 
    'NARIÑO', 'NORTE DE SANTANDER', 'PUTUMAYO', 'QUINDIO', 'RISARALDA', 'SAN ANDRES', 
    'SANTANDER', 'SUCRE', 'TOLIMA', 'VALLE DEL CAUCA', 'VAUPES', 'VICHADA'
}

def clean_regional_name(name):
    if not isinstance(name, str):
        return str(name)
    name = name.upper().strip()
    name = re.sub(r'[Á]', 'A', name)
    name = re.sub(r'[É]', 'E', name)
    name = re.sub(r'[Í]', 'I', name)
    name = re.sub(r'[Ó]', 'O', name)
    name = re.sub(r'[Ú]', 'U', name)
    return name

print(f"Buscando archivos en: {base_dir}")

for root, dirs, files in os.walk(base_dir):
    for f in files:
        f_lower = f.lower()
        # Filtros: que sea excel, que NO sea el consolidado, que NO sea HCB y que NO sea integralidad
        if (f.endswith(('.xlsx', '.xls', '.xlsm')) and 
            not f.startswith('~$') and 
            'consolidacion' not in f_lower and
            'hcb' not in f_lower and 
            'integralidad' not in f_lower):
            
            path = os.path.join(root, f)
            try:
                xls = pd.ExcelFile(path)
                sheet_match = [s for s in xls.sheet_names if 'ZONIFICAC' in s.upper()]
                if sheet_match:
                    sheet_name = sheet_match[0]
                    # Buscar la fila de encabezados
                    df_temp = pd.read_excel(xls, sheet_name=sheet_name, nrows=15, header=None)
                    
                    header_idx = None
                    for idx, row in df_temp.iterrows():
                        if any(isinstance(val, str) and 'Regional UDS' in val for val in row.values):
                            header_idx = idx
                            break
                    
                    if header_idx is not None:
                        df = pd.read_excel(xls, sheet_name=sheet_name, header=header_idx)
                        
                        col_regional = [c for c in df.columns if 'Regional UDS' in str(c)][0]
                        df = df.dropna(subset=[col_regional])
                        
                        # Fix datetime timezone issues to avoid Excel corruption
                        for col in df.columns:
                            if pd.api.types.is_datetime64_any_dtype(df[col]):
                                df[col] = df[col].dt.tz_localize(None).astype(str)
                        
                        df[col_regional] = df[col_regional].astype(str).str.strip()
                        regs = df[col_regional].unique()
                        for r in regs:
                            regionales_encontradas.add(clean_regional_name(r))
                        
                        # Normalize columns
                        df = df.rename(columns={col_regional: 'Regional UDS'})
                        df['Archivo_Origen'] = f
                        all_dfs.append(df)
                        print(f"  [OK] {f} - Regionales: {regs}")
            except Exception as e:
                print(f"  [ERROR] {f}: {e}")

# Escribir el consolidado
if all_dfs:
    df_final = pd.concat(all_dfs, ignore_index=True)
    # Crear carpeta si no existe
    os.makedirs(os.path.dirname(output_file), exist_ok=True)
    df_final.to_excel(output_file, sheet_name='ZONIFICACIÓN- PEDAGOGICO', index=False)
    print("\n¡ARCHIVO CREADO CON ÉXITO!:", output_file)
else:
    print("\nNO SE ENCONTRARON DATOS PARA CONSOLIDAR.")

print("\n--- RESUMEN ---")
print(f"Regionales encontradas: {len(regionales_encontradas)}")
missing = regionales_icbf - {r.replace('BOGOTA D.C.', 'BOGOTA') for r in regionales_encontradas}
if 'BOGOTA' in regionales_encontradas or 'BOGOTA D.C.' in regionales_encontradas:
    missing.discard('BOGOTA D.C.')
    missing.discard('BOGOTA')

if missing:
    print(f"Regionales FALTANTES ({len(missing)}): {sorted(list(missing))}")
else:
    print("¡Todas las regionales están presentes!")
