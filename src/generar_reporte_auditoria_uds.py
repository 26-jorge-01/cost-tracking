import pandas as pd
import numpy as np
import re
import os

# CONFIGURACIÓN DE RUTAS
BASE_DIR = r'D:\ICBF\cost-tracking'
DATA_DIR = os.path.join(BASE_DIR, 'data', 'insumos matriz 24 abril')

# Archivos Insumo
FILE_UDS = os.path.join(DATA_DIR, '20260414 Contrato_Marzo_2026.xlsx')
FILE_CONSOLIDADO = os.path.join(DATA_DIR, 'Regionales Contratacion Vigente 2026_CONSOLIDADO.xlsx')
FILE_CUE = os.path.join(DATA_DIR, 'ICBFCUEContUnicoxRegxVigxDir.xlsx')

# Archivo Salida
OUTPUT_FILE = os.path.join(BASE_DIR, 'UDS_15042026_Multi_Contrastado.xlsx')

def clean_text(text):
    if pd.isna(text): return ''
    text = str(text).upper()
    text = re.sub(r'[^\w\s]', ' ', text)
    text = ' '.join(text.split())
    return text

def run_pipeline():
    print("--- Iniciando Pipeline de Auditoría UDS ---")
    
    # 1. CARGA Y FILTRADO BASE (UDS)
    print("1/4 Cargando base UDS y filtrando HCB...")
    df_uds = pd.read_excel(FILE_UDS, sheet_name='UDS_15042026')
    df_uds = df_uds[~df_uds['Servicio'].str.contains('HCB', na=False)].copy()
    
    # Cálculo de totales agrupados (Lógica original)
    group_cols = ['NumeroContrato', 'Servicio']
    df_uds['Total_Cupos_Agrupado'] = df_uds.groupby(group_cols)['CuposUDS'].transform('sum')
    df_uds['Is_First_Grp'] = ~df_uds.duplicated(subset=group_cols)
    
    # 2. CRUCE CON CONSOLIDADO REGIONAL
    print("2/4 Cruzando con Consolidado Regional...")
    # Obtener columnas dinámicamente para evitar errores de mayúsculas/minúsculas
    df_cons_cols = pd.read_excel(FILE_CONSOLIDADO, sheet_name='Tabla1', nrows=0).columns.tolist()
    serv_cols = [c for c in df_cons_cols if c.lower().startswith('servicio ')]
    aten_cols = [c for c in df_cons_cols if c.lower().startswith('atenciones_')]
    
    df_cons = pd.read_excel(FILE_CONSOLIDADO, sheet_name='Tabla1', usecols=['Numero Documento Soporte'] + serv_cols + aten_cols)
    
    # Unpivot consolidado
    melted = []
    for s_col, a_col in zip(sorted(serv_cols), sorted(aten_cols)):
        temp = df_cons[['Numero Documento Soporte', s_col, a_col]].copy()
        temp.columns = ['Contrato_Ref', 'Servicio_Ref', 'Atenciones_Ref']
        melted.append(temp)
    
    df_cons_flat = pd.concat(melted).dropna(subset=['Servicio_Ref'])
    df_cons_flat['Servicio_Clean'] = df_cons_flat['Servicio_Ref'].apply(clean_text)
    df_cons_flat['Contrato_Ref'] = df_cons_flat['Contrato_Ref'].astype(str)
    df_cons_flat = df_cons_flat.groupby(['Contrato_Ref', 'Servicio_Clean'])['Atenciones_Ref'].sum().reset_index()

    # Join con Audit
    df_uds['Servicio_Clean'] = df_uds['Servicio'].apply(clean_text)
    df_uds['NumeroContrato_Str'] = df_uds['NumeroContrato'].astype(str)
    
    df_uds = pd.merge(df_uds, df_cons_flat, left_on=['NumeroContrato_Str', 'Servicio_Clean'], right_on=['Contrato_Ref', 'Servicio_Clean'], how='left')
    
    # 3. CRUCE CON CUE (UNIFICADO)
    print("3/4 Cruzando con reporte CUE...")
    df_cue = pd.read_excel(FILE_CUE, usecols=[14, 19]) # 14: Numero contrato, 19: Numero Cupos
    df_cue.columns = ['Contrato_CUE', 'Cupos_CUE']
    df_cue['Contrato_CUE'] = df_cue['Contrato_CUE'].astype(str)
    
    # Totales UDS por Contrato (para diferencia CUE)
    df_uds['Total_Cupos_Contrato_UDS'] = df_uds.groupby('NumeroContrato')['CuposUDS'].transform('sum')
    df_uds = pd.merge(df_uds, df_cue, left_on='NumeroContrato_Str', right_on='Contrato_CUE', how='left')

    # 4. LIMPIEZA FINAL Y EXPORTACIÓN
    print("4/4 Aplicando lógica de visualización y guardando...")
    df_uds['Is_First_Contract'] = ~df_uds.duplicated(subset=['NumeroContrato'])
    
    # Cálculos numéricos antes de convertir a strings
    df_uds['Diferencia_vs_Consolidado'] = df_uds['Total_Cupos_Agrupado'] - df_uds['Atenciones_Ref']
    df_uds['Diferencia_CUE_vs_UDS'] = df_uds['Cupos_CUE'] - df_uds['Total_Cupos_Contrato_UDS']
    
    # Convertir a objeto para permitir celdas vacías
    target_cols = [
        'Total_Cupos_Agrupado', 
        'Atenciones_Ref', 
        'Diferencia_vs_Consolidado', 
        'Cupos_CUE', 
        'Diferencia_CUE_vs_UDS'
    ]
    for col in target_cols:
        df_uds[col] = df_uds[col].astype(object)

    # Aplicar Filtro de Primera Ocurrencia por Grupo (Contrato + Servicio)
    df_uds.loc[~df_uds['Is_First_Grp'], ['Total_Cupos_Agrupado', 'Atenciones_Ref', 'Diferencia_vs_Consolidado']] = ''
    
    # Aplicar Filtro de Primera Ocurrencia por Contrato
    df_uds.loc[~df_uds['Is_First_Contract'], ['Cupos_CUE', 'Diferencia_CUE_vs_UDS']] = ''
    
    # Limpieza de columnas técnicas
    df_uds = df_uds.drop(columns=[
        'Is_First_Grp', 'Is_First_Contract', 'Servicio_Clean', 'NumeroContrato_Str', 
        'Contrato_Ref', 'Contrato_CUE', 'Total_Cupos_Contrato_UDS'
    ])
    
    df_uds = df_uds.rename(columns={
        'Atenciones_Ref': 'Cupos_Consolidado_Reportado',
        'Cupos_CUE': 'Cupos_CUE_Reportado'
    })

    df_uds.to_excel(OUTPUT_FILE, index=False)
    print(f"Éxito! Archivo generado en: {OUTPUT_FILE}")

if __name__ == "__main__":
    run_pipeline()
