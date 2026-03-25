# -*- coding: utf-8 -*-
import pandas as pd
import openpyxl
from pathlib import Path
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side

# Importar configuración
try:
    from . import config
except ImportError:
    import config

def aplicar_estilos_excel(writer, name):
    workbook = writer.book
    worksheet = writer.sheets[name]
    
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    
    # Columnas que deben formatearse como moneda
    cols_currency = ['VALOR_MENSUAL', 'VALOR_EJECUTADO', 'SALDO_DISPONIBLE', 'VALOR', 'VALOR_ACTUAL_CTO', 'TOTAL_APORTE_ICBF', 'VALOR_VALIDACION']
    
    for cell in worksheet[1]:
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal='center')

    for col in worksheet.columns:
        max_length = 0
        column = col[0].column_letter
        header = str(col[0].value).upper() if col[0].value else ""
        
        is_curr = any(c in header for c in cols_currency)
        
        for cell in col:
            if cell.row > 1 and is_curr and cell.value:
                cell.number_format = '$ #,##0'
            
            try:
                if len(str(cell.value)) > max_length: max_length = len(str(cell.value))
            except: pass
        worksheet.column_dimensions[column].width = min(max_length + 2, 60)

def vincular_pdfs_a_contratos(df_contratos, df_datos):
    """
    Vea qué PDFs corresponden a cada contrato basándose en la LLAVE.
    """
    try:
        # Renombrar 'Valor Validacion' a LLAVE si viene de la Fase 1
        if 'Valor Validacion' in df_datos.columns and config.COL_LLAVE not in df_datos.columns:
            df_datos = df_datos.rename(columns={'Valor Validacion': config.COL_LLAVE})
        
        # Renombrar 'Mes Detectado' a config.COL_MES
        if 'Mes Detectado' in df_datos.columns and config.COL_MES not in df_datos.columns:
            df_datos = df_datos.rename(columns={'Mes Detectado': config.COL_MES})
        df_contratos[config.COL_LLAVE] = df_contratos[config.COL_LLAVE].astype(str).str.strip()
        df_datos[config.COL_LLAVE] = df_datos[config.COL_LLAVE].astype(str).str.strip()
        df_datos[config.COL_MES] = df_datos[config.COL_MES].astype(str).str.lower().str.strip()

        # Marcar si es 'No PDF' (para no contar como evidencia real)
        df_evidencia = df_datos[df_datos['Nombre PDF'] != 'No PDF'].copy()

        # Pivotar PDFs por mes
        df_pivot = df_evidencia.drop_duplicates(subset=[config.COL_LLAVE, config.COL_MES])
        df_pivot = df_pivot.pivot(index=config.COL_LLAVE, columns=config.COL_MES, values='Nombre PDF')
        df_pivot = df_pivot.rename(columns=config.MESES_MAP)
        
        columnas_pdf = list(config.MESES_MAP.values())
        for col in columnas_pdf:
            if col not in df_contratos.columns:
                df_contratos[col] = 'N'
            df_contratos[col] = df_contratos[col].astype(object).fillna('N')

        # Mezclar
        cols_p = [c for c in columnas_pdf if c in df_pivot.columns]
        df_final = df_contratos.merge(df_pivot[cols_p], on=config.COL_LLAVE, how='left', suffixes=('_old', ''))
        
        for col in columnas_pdf:
            if col in df_final.columns:
                df_final[col] = df_final[col].fillna('N')
                if f"{col}_old" in df_final.columns: df_final.drop(columns=[f"{col}_old"], inplace=True)
        
        return df_final
    except Exception as e:
        print(f"Error vinculando: {e}")
        return df_contratos

def actualizar_contratos_y_reescribir_limpio():
    try:
        print("--- FASE 2: CALIDAD MAXIMA ---")
        path_celdas = config.BASE_PATH / config.OUTPUT_VALORES_CELDAS
        path_pdfs = config.BASE_PATH / config.OUTPUT_DETALLE_PDFS
        
        if not path_celdas.exists():
            print("No se encontro reporte de Fase 1.")
            return

        df_celdas = pd.read_excel(path_celdas)
        df_pdfs = pd.read_excel(path_pdfs)
        
        # Eliminar filas basura
        df_celdas = df_celdas.dropna(subset=[config.COL_LLAVE])
        df_celdas[config.COL_CONTRATO_ENCONTRADO] = 'SI'

        # Vincular
        df_final = vincular_pdfs_a_contratos(df_celdas, df_pdfs)
        
        # Guardar Maestro
        with pd.ExcelWriter(config.MASTER_FILE_PATH, engine='openpyxl') as writer:
            df_final.to_excel(writer, sheet_name='CONTRATOS', index=False)
            aplicar_estilos_excel(writer, 'CONTRATOS')
            
            df_pdfs.to_excel(writer, sheet_name='DETALLE_PDFS', index=False)
            aplicar_estilos_excel(writer, 'DETALLE_PDFS')

        print(f"EXITO! Reporte generado en: {config.MASTER_FILE_PATH}")

    except Exception as e:
        print(f"Error critico Fase 2: {e}")

if __name__ == "__main__":
    actualizar_contratos_y_reescribir_limpio()