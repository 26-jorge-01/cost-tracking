import os
import pandas as pd
import warnings
import re
import unicodedata

warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

# Configuración de rutas
base_dir = 'd:/ICBF/cost-tracking/data/Matrices_validadas_definitivas'
output_file = 'd:/ICBF/cost-tracking/data/insumos 28 abril/consolidacion_matriz_hcb_28042026.xlsx'

def normalize_str(s):
    if not isinstance(s, str):
        return str(s)
    s = s.upper().strip()
    s = ''.join(c for c in unicodedata.normalize('NFD', s) if unicodedata.category(c) != 'Mn')
    s_clean = re.sub(r'[^A-Z0-9]', '', s)
    return s_clean

def clean_currency(val):
    if pd.isna(val) or val == '': return 0.0
    if isinstance(val, (int, float)): return float(val)
    # Limpiar string: quitar $, puntos de miles, espacios, etc.
    s = str(val).replace('$', '').replace(' ', '').replace('.', '').replace(',', '.')
    # Si después de limpiar hay varios puntos (ej. 1.200.000 -> 1200000)
    # Ya quitamos los puntos arriba, ahora nos aseguramos que sea convertible
    try:
        return float(s)
    except:
        # Intentar rescatar solo números
        s_only_num = re.sub(r'[^0-9.]', '', s)
        try: return float(s_only_num)
        except: return 0.0

def clean_regional_name(name):
    if not isinstance(name, str): return str(name)
    name = name.upper().strip()
    name = ''.join(c for c in unicodedata.normalize('NFD', name) if unicodedata.category(c) != 'Mn')
    # Casos especiales
    if 'BOGOTA' in name: return 'BOGOTA D.C.'
    return name

# Lista de regionales válidas para filtrado
REGIONALES_VALIDAS = [
    'AMAZONAS', 'ANTIOQUIA', 'ARAUCA', 'ATLANTICO', 'BOGOTA D.C.', 'BOLIVAR', 
    'BOYACA', 'CALDAS', 'CAQUETA', 'CASANARE', 'CAUCA', 'CESAR', 'CHOCO', 
    'CORDOBA', 'CUNDINAMARCA', 'GUAINIA', 'GUAVIARE', 'HUILA', 'LA GUAJIRA', 
    'MAGDALENA', 'META', 'NARINO', 'NORTE DE SANTANDER', 'PUTUMAYO', 'QUINDIO', 
    'RISARALDA', 'SAN ANDRES', 'SANTANDER', 'SUCRE', 'TOLIMA', 'VALLE DEL CAUCA', 
    'VAUPES', 'VICHADA'
]

# Mapeo de alias comunes en nombres de archivos
ALIASES_REGIONALES = {
    'NSANTANDER': 'NORTE DE SANTANDER',
    'NORTEDESANTANDER': 'NORTE DE SANTANDER',
    'N.SANTANDER': 'NORTE DE SANTANDER',
    'SANTANDER': 'SANTANDER',
    'VALLE': 'VALLE DEL CAUCA',
    'BOGOTA': 'BOGOTA D.C.',
    'ATLANTICO': 'ATLANTICO',
    'NARIO': 'NARINO',
    'NARINO': 'NARINO',
    'BOYACA': 'BOYACA',
    'BOLIVAR': 'BOLIVAR',
    'CHOCO': 'CHOCO',
    'CORDOBA': 'CORDOBA',
}

# Columnas finales (Nombres limpios para Excel)
COLS_MATRIZ_HCB = [
    'REGIONAL', 'CENTRO ZONAL', 'MUNICIPIO', 'REFERENCIA / No. CONTRATO SECOP', 
    'No. RP', 'VIGENCIA', 'FORMA CONTRATACION',
    'NIT CONTRATISTA 2026', 'CONTRATISTA 2026', 'SERVICIO 2026', 
    'Componente para la UDS', 'DURACION INICIAL', 'DURACION ADICION',
    'CUPOS', 'CANTIDAD UDS',
    'VALOR UNITARIO MES', 'VALOR CANASTA 2026',
    'VALOR INICIAL 2026', 'VALOR INICIAL 2026 (SMLV)', 'APORTE CONTRAPARTIDA', 
    'VALOR ADICIONES A LA FECHA', 'ADICION OTROS CONCEPTOS', 
    'VALOR REDUCCIONES A LA FECHA', 'INEJECUCIONES', 
    'VALOR ACTUAL', 'VALOR ACTUAL TOTAL DEL CONTRATO (SMLV)', 
    'VALOR FINAL ADICION SERVICIO', 'VALOR FINAL ADICION SERVICIO (SMLV)', 
    'VALOR A ADICIONAR', 'CONTRAPARTIDA ADICION', 
    'VALOR TOTAL DEL CONTRATO (SMLV)', 
    'SANCIONATORIOS', 'ALERTA 1000 SMLMV', 'ALERTA 5000 SMLMV',
    'Archivo_Origen'
]

COLS_ZONI_HCB = [
    'Regional UDS', 'Centro Zonal UDS', 'Municipio UDS', 'REFERENCIA / No. CONTRATO SECOP',
    'No. RP', 'NIT CONTRATISTA 2026', 'CONTRATISTA 2026', 'SERVICIO 2026', 
    'Componente para la UDS', 'Cupos', 'CANTIDAD DE MADRES POR UDS',
    'VALOR UNITARIO MES', 'VALOR ADICION CANASTA (MULTIPLE)', 'VALOR TOTAL ADICION (MULTIPLE)',
    'Archivo_Origen'
]

def map_hcb_columns(df_cols):
    rename_dict = {}
    for col in df_cols:
        c = normalize_str(col)
        if 'REGIONAL' in c: rename_dict[col] = 'REGIONAL'
        elif 'CONTRATOSECOP' in c: rename_dict[col] = 'REFERENCIA / No. CONTRATO SECOP'
        elif 'CENTROZONAL' in c: rename_dict[col] = 'CENTRO ZONAL'
        elif 'MUNICIPIO' in c: rename_dict[col] = 'MUNICIPIO'
        elif 'NOMBREDELSERVICIO' in c: rename_dict[col] = 'SERVICIO 2026'
        elif 'MODALIDADDESERVICIO' in c: rename_dict[col] = 'Componente para la UDS'
        elif c == 'NIT': rename_dict[col] = 'NIT CONTRATISTA 2026'
        elif 'NOMBREEAS' in c: rename_dict[col] = 'CONTRATISTA 2026'
        elif 'CUPOSPORUNIDAD' in c: rename_dict[col] = 'Cupos'
        elif 'MADRESPORUDS' in c and '2025' in c: rename_dict[col] = 'CANTIDAD DE MADRES POR UDS'
        
        # Administrativos y Tiempos
        elif c == 'NORP': rename_dict[col] = 'No. RP'
        elif c == 'VIGENCIA': rename_dict[col] = 'VIGENCIA'
        elif 'FORMADECONTRATACION' in c: rename_dict[col] = 'FORMA CONTRATACION'
        elif 'TIEMPOINICIALDELCONTRATO' in c: rename_dict[col] = 'DURACION INICIAL'
        elif 'TIEMPOAADICIONAR' in c: rename_dict[col] = 'DURACION ADICION'
        elif 'CANTIDADDEUDS' in c: rename_dict[col] = 'CANTIDAD UDS'
        elif 'SANCIONATORIOS' in c: rename_dict[col] = 'SANCIONATORIOS'
        elif 'ALERTADEAVAL' in c and '1000' in c: rename_dict[col] = 'ALERTA 1000 SMLMV'
        elif 'ALERTA' in c and '5000' in c: rename_dict[col] = 'ALERTA 5000 SMLMV'
        
        # Variables de Análisis Presupuestal
        elif 'VALORUNITARIO' in c and 'MES' in c: rename_dict[col] = 'VALOR UNITARIO MES'
        elif 'VALORCANASTA' in c and '2026' in c: rename_dict[col] = 'VALOR CANASTA 2026'
        
        # Pesos
        elif 'VALORTOTALINICIAL' in c and 'APORTEICBF' in c and 'UNICO' in c: rename_dict[col] = 'VALOR INICIAL 2026'
        elif 'VALORINICIALCONTRAPARTIDA' in c: rename_dict[col] = 'APORTE CONTRAPARTIDA'
        elif 'VALORADICIONESHISTORICAS' in c and 'ICBF' in c: rename_dict[col] = 'VALOR ADICIONES A LA FECHA'
        elif 'VALORADICIONOTROSCONCEPTOS' in c: rename_dict[col] = 'ADICION OTROS CONCEPTOS'
        elif 'VALORREDUCCIONESHISTORICAS' in c and 'ICBF' in c: rename_dict[col] = 'VALOR REDUCCIONES A LA FECHA'
        elif 'VALORINEJECUCIONES' in c: rename_dict[col] = 'INEJECUCIONES'
        elif 'VALORACTUALDELCONTRATO' in c and 'ANTES' in c and 'UNICO' in c and 'SML' not in c: rename_dict[col] = 'VALOR ACTUAL'
        elif 'VALORTOTALDELAADICIONCONTRATO' in c and 'UNICO' in c and 'SML' not in c: rename_dict[col] = 'VALOR FINAL ADICION SERVICIO'
        elif 'VALORTOTALDELAADICIONAPORTEICBF' in c and 'UNICO' in c: rename_dict[col] = 'VALOR A ADICIONAR'
        elif 'VALORCONTRAPARTIDAADICION' in c: rename_dict[col] = 'CONTRAPARTIDA ADICION'
        
        # SMLV (Detección por palabra 'SML' o 'SMML')
        elif 'VALORTOTALINICIAL' in c and 'UNICO' in c and ('SML' in c or 'SMML' in c): rename_dict[col] = 'VALOR INICIAL 2026 (SMLV)'
        elif 'VALORACTUAL' in c and 'UNICO' in c and ('SML' in c or 'SMML' in c): rename_dict[col] = 'VALOR ACTUAL TOTAL DEL CONTRATO (SMLV)'
        elif 'VALORTOTALDELAADICION' in c and 'UNICO' in c and ('SML' in c or 'SMML' in c): rename_dict[col] = 'VALOR FINAL ADICION SERVICIO (SMLV)'
        elif 'VALORFINAL' in c and 'UNICO' in c and ('SML' in c or 'SMML' in c): rename_dict[col] = 'VALOR TOTAL DEL CONTRATO (SMLV)'
        
        # Granular
        elif 'VALORADICIONNIVELACIONCANASTA' in c and 'MULTIPLE' in c: rename_dict[col] = 'VALOR ADICION CANASTA (MULTIPLE)'
        elif 'VALORTOTALDELAADICIONAPORTEICBF' in c and 'MULTIPLE' in c: rename_dict[col] = 'VALOR TOTAL ADICION (MULTIPLE)'
        
        # Cupos (Manejo de erratas como 'CONTRAO' y variaciones)
        elif 'CUPOS' in c and ('UNICO' in c or 'CONTRATO' in c or 'CONTRAO' in c or 'SUMATORIA' in c): 
            rename_dict[col] = 'CUPOS'
        elif 'CUPOS' in c and 'UNIDAD' in c: 
            rename_dict[col] = 'Cupos' # Para zonificación
        
    return rename_dict

all_data = []
print(f"Consolidando HCB con Resumen Ejecutivo...")

# Tracking global para evitar duplicados de regionales "fantasma"
processed_pure_regionals = set()

# Columnas que deben ser numéricas
FIN_KEYWORDS = ['VALOR', 'APORTE', 'ADICION', 'REDUCCION', 'INEJECUCION', 'CONTRAPARTIDA', 'SMLV', 'CUPOS', 'CANTIDAD']

for root, dirs, files in os.walk(base_dir):
    for f in files:
        f_lower = f.lower()
        if (f.endswith(('.xlsx', '.xls', '.xlsm')) and not f.startswith('~$') and 
            'hcb' in f_lower and 'alimento' not in f_lower and 'consolidacion' not in f_lower):
            path = os.path.join(root, f)
            try:
                xls = pd.ExcelFile(path)
                sheet_m = [s for s in xls.sheet_names if 'MATRIZ' == s.upper()]
                if sheet_m:
                    df_temp = pd.read_excel(xls, sheet_name=sheet_m[0], nrows=20, header=None)
                    h_idx = None
                    for idx, row in df_temp.iterrows():
                        if any(isinstance(val, str) and 'REGIONAL' == normalize_str(val) for val in row.values):
                            h_idx = idx
                            break
                    if h_idx is not None:
                        df = pd.read_excel(xls, sheet_name=sheet_m[0], header=h_idx)
                        # --- DEDUPLICACIÓN DE COLUMNAS ---
                        df = df.loc[:, ~df.columns.duplicated()].copy()
                        
                        rename_dict = map_hcb_columns(df.columns)
                        df = df.rename(columns=rename_dict)
                        # --- DEDUPLICACIÓN DE COLUMNAS (Después de renombrar) ---
                        df = df.loc[:, ~df.columns.duplicated()].copy()
                        
                        for col in df.columns:
                            if pd.api.types.is_datetime64_any_dtype(df[col]):
                                df[col] = df[col].dt.tz_localize(None).astype(str)
                        if 'REGIONAL' in df.columns or any(normalize_str(c) == 'REGIONAL' for c in df.columns):
                            # Propagación
                            cols_to_ffill = [
                                'REGIONAL', 'CENTRO ZONAL', 'MUNICIPIO', 'REFERENCIA / No. CONTRATO SECOP', 
                                'NIT CONTRATISTA 2026', 'CONTRATISTA 2026', 'SERVICIO 2026', 
                                'Componente para la UDS', 'No. RP', 'VIGENCIA', 'FORMA CONTRATACION',
                                'DURACION INICIAL', 'DURACION ADICION', 'CANTIDAD UDS', 'SANCIONATORIOS',
                                'ALERTA 1000 SMLMV', 'ALERTA 5000 SMLMV', 'CUPOS',
                                'VALOR UNITARIO MES', 'VALOR CANASTA 2026'
                            ]
                            # También ffill de columnas financieras
                            fin_cols_temp = [c for c in df.columns if any(x in str(c).upper() for x in ['VALOR', 'APORTE', 'ADICION', 'REDUCCION', 'INEJECUCION', 'CONTRAPARTIDA'])]
                            for c in cols_to_ffill + fin_cols_temp:
                                if c in df.columns: df[c] = df[c].ffill()
                            if 'REFERENCIA / No. CONTRATO SECOP' in df.columns:
                                df['REFERENCIA / No. CONTRATO SECOP'] = df['REFERENCIA / No. CONTRATO SECOP'].astype(str).str.strip().str.upper()
                            df = df.dropna(subset=['REGIONAL'])
                            df['REGIONAL'] = df['REGIONAL'].apply(clean_regional_name)
                            
                            # Limpieza numérica preventiva
                            for col in df.columns:
                                if any(k in col.upper() for k in FIN_KEYWORDS):
                                    df[col] = df[col].apply(clean_currency)
                            
                            # --- LIMPIEZA DE FILAS DE TOTALES ---
                            # Eliminamos filas que sean resúmenes o totales dentro del Excel regional
                            mask_totals = df.astype(str).apply(lambda x: x.str.contains('TOTAL|SUBTOTAL', case=False, na=False)).any(axis=1)
                            df = df[~mask_totals].copy()
                            
                            # --- FILTRO DE PERTENENCIA REGIONAL ---
                            f_norm = normalize_str(f)
                            root_norm = normalize_str(root)
                            target_regional = None
                            
                            for alias, official in ALIASES_REGIONALES.items():
                                if alias in f_norm or alias in root_norm:
                                    target_regional = official
                                    break
                            if not target_regional:
                                for reg in REGIONALES_VALIDAS:
                                    reg_norm = normalize_str(reg)
                                    if reg_norm in f_norm or reg_norm in root_norm:
                                        target_regional = reg
                                        break
                            if not target_regional:
                                mode_res = df['REGIONAL'].mode()
                                target_regional = mode_res[0] if not mode_res.empty else 'DESCONOCIDO'
                            
                            target_norm = normalize_str(target_regional)
                            df['REGIONAL_NORM'] = df['REGIONAL'].apply(normalize_str)
                            
                            # Lógica de Rescate de Santander
                            mask_target = (df['REGIONAL_NORM'] == target_norm)
                            mask_santander = (df['REGIONAL_NORM'] == 'SANTANDER')
                            
                            if ('SANTANDER' not in processed_pure_regionals) and mask_santander.any():
                                df_clean = df[mask_target | mask_santander].copy()
                                processed_pure_regionals.add('SANTANDER')
                            else:
                                df_clean = df[mask_target].copy()
                                
                            if target_regional not in processed_pure_regionals:
                                processed_pure_regionals.add(target_regional)

                            df_clean = df_clean.drop(columns=['REGIONAL_NORM'])
                            df = df_clean
                            
                            if not df.empty:
                                print(f"  [OK] {f}: Procesada como {target_regional} ({len(df)} filas)")
                                df['Archivo_Origen'] = f
                                all_data.append(df)
            except Exception as e: print(f"  [ERROR] {f}: {e}")

if all_data:
    df_full = pd.concat(all_data, ignore_index=True)
    df_full = df_full.loc[:, ~df_full.columns.duplicated()]

    # --- MATRIZ_CALCULADA ---
    # No agrupamos para mantener la granularidad original (ej. por UDS/Sede)
    df_matriz = df_full.copy()
    
    # --- FALLBACK DE CUPOS (Específico para regionales como Valle que dejan el total en cero) ---
    # Si CUPOS (total contrato) es 0 o NaN, calculamos la suma de Cupos (nivel sede) por contrato
    if 'CUPOS' in df_matriz.columns and 'Cupos' in df_matriz.columns:
        df_matriz['CUPOS'] = pd.to_numeric(df_matriz['CUPOS'], errors='coerce').fillna(0)
        df_matriz['Cupos'] = pd.to_numeric(df_matriz['Cupos'], errors='coerce').fillna(0)
        
        # Calculamos totales por contrato basados en la granularidad de las sedes
        contract_totals = df_matriz.groupby('REFERENCIA / No. CONTRATO SECOP')['Cupos'].transform('sum')
        # Solo aplicamos el fallback donde el total original sea 0
        mask_zero = (df_matriz['CUPOS'] == 0) & (contract_totals > 0)
        df_matriz.loc[mask_zero, 'CUPOS'] = contract_totals[mask_zero]
        
    df_matriz = df_matriz.sort_values(['REGIONAL', 'REFERENCIA / No. CONTRATO SECOP'])
    
    # Para evitar doble contabilización en la suma de Excel:
    # Solo dejamos los valores de contrato en la PRIMERA fila de cada contrato
    fin_cols = [c for c in df_matriz.columns if any(x in c.upper() for x in ['VALOR', 'APORTE', 'ADICION', 'REDUCCION', 'INEJECUCION', 'CONTRAPARTIDA', 'SMLV'])]
    mask_dupe = df_matriz.duplicated(subset=['REFERENCIA / No. CONTRATO SECOP'], keep='first')
    
    for c in fin_cols + ['CUPOS', 'CANTIDAD UDS', 'DURACION INICIAL', 'DURACION ADICION']:
        if c in df_matriz.columns and 'MULTIPLE' not in c.upper():
            df_matriz.loc[mask_dupe, c] = None
    
    # Asegurar columnas finales
    for c in COLS_MATRIZ_HCB:
        if c not in df_matriz.columns: df_matriz[c] = None
    df_matriz = df_matriz[COLS_MATRIZ_HCB]

    # --- ZONIFICACIÓN ---
    df_zoni = df_full.copy().rename(columns={'REGIONAL': 'Regional UDS', 'CENTRO ZONAL': 'Centro Zonal UDS', 'MUNICIPIO': 'Municipio UDS'})
    for c in COLS_ZONI_HCB:
        if c not in df_zoni.columns: df_zoni[c] = None
    df_zoni = df_zoni[COLS_ZONI_HCB]

    # --- RESUMEN_EJECUTIVO ---
    # Usamos df_matriz (sin duplicados) para el resumen para asegurar consistencia
    res_reg = df_matriz.groupby('REGIONAL').agg({
        'REFERENCIA / No. CONTRATO SECOP': 'nunique',
        'CUPOS': 'sum',
        'VALOR INICIAL 2026': 'sum',
        'VALOR A ADICIONAR': 'sum'
    }).rename(columns={'REFERENCIA / No. CONTRATO SECOP': 'No. Contratos', 'CUPOS': 'Cupos Totales', 'VALOR INICIAL 2026': 'Valor Inicial (Total)', 'VALOR A ADICIONAR': 'Valor Adicion (Total)'})
    
    # Madres (sumamos de todas las filas granulares)
    madres_reg = df_full.groupby('REGIONAL')['CANTIDAD DE MADRES POR UDS'].sum()
    res_reg = res_reg.join(madres_reg).rename(columns={'CANTIDAD DE MADRES POR UDS': 'Total Madres'})
    res_reg['Inversion por Cupo'] = res_reg['Valor Adicion (Total)'] / res_reg['Cupos Totales']
    res_reg = res_reg.reset_index()

    # Resumen por Modalidad
    df_contracts = df_matriz.drop_duplicates(subset=['REFERENCIA / No. CONTRATO SECOP'])
    res_mod = df_contracts.groupby('Componente para la UDS').agg({
        'REFERENCIA / No. CONTRATO SECOP': 'nunique',
        'CUPOS': 'sum',
        'VALOR A ADICIONAR': 'sum'
    }).reset_index().rename(columns={'REFERENCIA / No. CONTRATO SECOP': 'No. Contratos', 'CUPOS': 'Cupos', 'VALOR A ADICIONAR': 'Adicion Presupuestal'})

    # --- CONSOLIDADO POR CONTRATO (Estilo Casanare - Super Matriz Ejecutiva) ---
    # Calculamos sumatorias granulares para rescate si el total único falla
    df_full['VALOR TOTAL ADICION (MULTIPLE)'] = pd.to_numeric(df_full['VALOR TOTAL ADICION (MULTIPLE)'], errors='coerce').fillna(0)
    
    agg_contrato = {
        'REGIONAL': 'first',
        'NIT CONTRATISTA 2026': 'first',
        'CONTRATISTA 2026': 'first',
        'SERVICIO 2026': lambda x: ' / '.join(sorted(set(str(v) for v in x if pd.notna(v)))),
        'Componente para la UDS': 'first',
        'No. RP': 'first',
        'VIGENCIA': 'first',
        'FORMA CONTRATACION': 'first',
        'DURACION INICIAL': 'first',
        'DURACION ADICION': 'first',
        'CUPOS': 'first',
        'CANTIDAD DE MADRES POR UDS': 'sum',
        'VALOR UNITARIO MES': 'first',
        'VALOR CANASTA 2026': 'first',
        # Pesos ($)
        'VALOR INICIAL 2026': 'first',
        'APORTE CONTRAPARTIDA': 'first',
        'VALOR ADICIONES A LA FECHA': 'first',
        'ADICION OTROS CONCEPTOS': 'first',
        'VALOR REDUCCIONES A LA FECHA': 'first',
        'INEJECUCIONES': 'first',
        'VALOR ACTUAL': 'first',
        'VALOR FINAL ADICION SERVICIO': 'first',
        'VALOR A ADICIONAR': 'first',
        'CONTRAPARTIDA ADICION': 'first',
        # SMLV
        'VALOR INICIAL 2026 (SMLV)': 'first',
        'VALOR ACTUAL TOTAL DEL CONTRATO (SMLV)': 'first',
        'VALOR FINAL ADICION SERVICIO (SMLV)': 'first',
        'VALOR TOTAL DEL CONTRATO (SMLV)': 'first',
        # Alertas
        'SANCIONATORIOS': 'first',
        'ALERTA 1000 SMLMV': 'first',
        'ALERTA 5000 SMLMV': 'first'
    }
    
    # Ejecutamos la agregación
    df_contrato = df_full.groupby('REFERENCIA / No. CONTRATO SECOP').agg({k: v for k, v in agg_contrato.items() if k in df_full.columns}).reset_index()
    
    # --- FALLBACK FINANCIERO ---
    # Si el valor único es 0, usamos la suma de los múltiples
    df_contrato['VALOR_MULTIPLE_SUM'] = df_full.groupby('REFERENCIA / No. CONTRATO SECOP')['VALOR TOTAL ADICION (MULTIPLE)'].sum().values
    if 'VALOR A ADICIONAR' in df_contrato.columns:
        mask_val_zero = (df_contrato['VALOR A ADICIONAR'] == 0) & (df_contrato['VALOR_MULTIPLE_SUM'] > 0)
        df_contrato.loc[mask_val_zero, 'VALOR A ADICIONAR'] = df_contrato['VALOR_MULTIPLE_SUM']
    
    # Limpiar servicios vacíos
    if 'SERVICIO 2026' in df_contrato.columns:
        df_contrato['SERVICIO 2026'] = df_contrato['SERVICIO 2026'].replace('', 'SERVICIO NO ESPECIFICADO')
    
    # Reordenar columnas al estilo Casanare (Lógica de Bloques)
    # 1. Identificación, 2. Capacidad, 3. Tiempos, 4. Financiero Pesos, 5. Financiero SMLV, 6. Alertas
    cols_priority = [
        'REGIONAL', 'REFERENCIA / No. CONTRATO SECOP', 'NIT CONTRATISTA 2026', 'CONTRATISTA 2026',
        'SERVICIO 2026', 'No. RP', 'VIGENCIA', 'CUPOS', 'CANTIDAD DE MADRES POR UDS',
        'DURACION INICIAL', 'DURACION ADICION',
        'VALOR INICIAL 2026', 'VALOR ACTUAL', 'VALOR A ADICIONAR', 'VALOR FINAL ADICION SERVICIO',
        'VALOR INICIAL 2026 (SMLV)', 'VALOR TOTAL DEL CONTRATO (SMLV)',
        'SANCIONATORIOS', 'ALERTA 1000 SMLMV', 'ALERTA 5000 SMLMV'
    ]
    # Mantener el resto de columnas que existan
    final_cols_contrato = [c for c in cols_priority if c in df_contrato.columns]
    final_cols_contrato += [c for c in df_contrato.columns if c not in final_cols_contrato and c != 'VALOR_MULTIPLE_SUM']
    df_contrato = df_contrato[final_cols_contrato]

    # Guardar
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        res_reg.to_excel(writer, sheet_name='RESUMEN_EJECUTIVO', index=False)
        df_contrato.to_excel(writer, sheet_name='CONSOLIDADO_CONTRATO', index=False)
        df_matriz.to_excel(writer, sheet_name='MATRIZ_CALCULADA', index=False)
        df_zoni.to_excel(writer, sheet_name='ZONIFICACIÓN- PEDAGOGICO', index=False)
        res_mod.to_excel(writer, sheet_name='RESUMEN_MODALIDAD', index=False)

    print(f"\n¡CONSOLIDACIÓN HCB COMPLETADA!")
    print(f"Ubicación: {output_file}")
