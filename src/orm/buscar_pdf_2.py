# -*- coding: utf-8 -*-
"""
Created on Fri Nov 28 05:59:04 2025

@author: vicga
"""

import pandas as pd

def actualizar_contratos_y_reescribir_limpio(ruta_archivo):
    """
    Compara la columna 'LLAVE' de la hoja 'CONTRATOS' con la columna 'LLAVE'
    de la hoja 'DATOS' y actualiza la hoja 'CONTRATOS' con el número de fila
    (basado en 1) de la coincidencia en 'DATOS' o 'N'.
    Luego reescribe todo el archivo Excel para limpiar posibles corrupciones.
    """
    try:
        
        xls = pd.ExcelFile(ruta_archivo)
        
        
        todas_las_hojas = {}
        for sheet_name in xls.sheet_names:
            todas_las_hojas[sheet_name] = pd.read_excel(xls, sheet_name=sheet_name)

        
        df_contratos_trabajo = todas_las_hojas['CONTRATOS'].copy()
        df_datos = todas_las_hojas['DATOS'].copy()

        
        df_contratos_trabajo['LLAVE'] = df_contratos_trabajo['LLAVE'].astype(str).str.strip()
        df_datos['LLAVE'] = df_datos['LLAVE'].astype(str).str.strip()

        
        df_datos['Mes'] = df_datos['Mes'].astype(str).str.lower().str.strip()

        
        meses_map = {
            'enero': 'PDF ENERO',
            'febrero': 'PDF FEBRERO',
            'marzo': 'PDF MARZO',
            'abril': 'PDF ABRIL',
            'mayo': 'PDF MAYO',
            'junio': 'PDF JUNIO',
            'julio': 'PDF JULIO',
            'agosto': 'PDF AGOSTO',
            'septiembre': 'PDF SEPTIEMBRE',
            'octubre': 'PDF OCTUBRE',
            'noviembre': 'PDF NOVIEMBRE',
            'diciembre': 'PDF DICIEMBRE'
        }

        
        for col_nombre_pdf in meses_map.values():
            if col_nombre_pdf not in df_contratos_trabajo.columns:
                df_contratos_trabajo[col_nombre_pdf] = 'N'
            else:
                df_contratos_trabajo[col_nombre_pdf] = 'N'

        
        for index, row_contratos in df_contratos_trabajo.iterrows():
            llave_contrato = row_contratos['LLAVE']
            contrato_encontrado_val = str(row_contratos['CONTRATO ENCONTRADO']).upper().strip()

            if contrato_encontrado_val == 'SI':
                
                entradas_datos = df_datos[df_datos['LLAVE'] == llave_contrato]

                if not entradas_datos.empty:
                    
                    for original_mes, col_pdf_contratos in meses_map.items():
                        
                        matching_rows_indices = entradas_datos[entradas_datos['Mes'] == original_mes].index

                        if not matching_rows_indices.empty:
                            pandas_index = matching_rows_indices[0]
                            
                            excel_row_number = pandas_index + 2

                            df_contratos_trabajo.loc[index, col_pdf_contratos] = int(excel_row_number)
            
        
        todas_las_hojas['CONTRATOS'] = df_contratos_trabajo

        
        with pd.ExcelWriter(ruta_archivo, engine='openpyxl', mode='w') as writer:
            for sheet_name, df_sheet in todas_las_hojas.items():
                df_sheet.to_excel(writer, sheet_name=sheet_name, index=False)

        print("Proceso completado, engineer. El archivo se ha guardado exitosamente reescribiendo todas las hojas. Esto debería mitigar los errores de corrupción al abrirlo.")

    except FileNotFoundError:
        print(f"Error, engineer: El archivo no se encontró en la ruta especificada: {ruta_archivo}")
    except KeyError as e:
        print(f"Error, engineer: Una columna esperada no se encontró. Por favor, verifica los nombres de las columnas 'LLAVE', 'CONTRATO ENCONTRADO', y 'Mes' en tus hojas. Error: {e}")
    except Exception as e:
        print(f"Ocurrió un error inesperado, engineer: {e}")

# Ruta del archivo Excel
#ruta_del_archivo = r"C:\Users\vicga\OneDrive\Documentos\ICBF\2025\ACTAS DE LEGALIZACION\MODELO BI\TRAER_DATO_PDF_POR_MES.xlsx"
ruta_del_archivo = r"C:\Users\vicga\OneDrive\Documentos\ICBF\2025\ACTAS DE LEGALIZACION\MODELO BI\TRAER_DATO_PDF_POR_MES.xlsx"

# Llamar a la función para ejecutar el proceso
actualizar_contratos_y_reescribir_limpio(ruta_del_archivo)