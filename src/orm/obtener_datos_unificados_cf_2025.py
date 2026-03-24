# -*- coding: utf-8 -*-
"""
Editor de Spyder

Este es un archivo temporal.
"""

import os
import fnmatch
import openpyxl
import pandas as pd
from tqdm import tqdm
import warnings
from datetime import datetime
from pathlib import Path


warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

def encontrar_archivos_xlsm(ruta):
    """
    Encuentra todos los archivos .xlsm dentro de una ruta y sus subcarpetas usando pathlib.
    """
    ruta_path = Path(ruta)
    archivos_xlsm = list(ruta_path.rglob('*.xlsm'))
    return [str(p) for p in archivos_xlsm]

def contar_pdfs_en_ruta(ruta):
    """
    Cuenta el número de archivos .pdf en una ruta específica usando pathlib.
    """
    ruta_path = Path(ruta)
    return len([archivo for archivo in ruta_path.iterdir() if archivo.is_file() and archivo.suffix == '.pdf'])

def listar_pdfs_en_ruta(ruta):
    """
    Lista los nombres de los archivos .pdf en una ruta específica usando pathlib.
    """
    ruta_path = Path(ruta)
    return [archivo.name for archivo in ruta_path.iterdir() if archivo.is_file() and archivo.suffix == '.pdf']

def extraer_mes_del_nombre(nombre_archivo):
    """
    Extrae el nombre del mes de un string (generalmente un nombre de archivo).
    """
    meses_map = {
        'enero': 'Enero', 'febrero': 'Febrero', 'marzo': 'Marzo', 'abril': 'Abril',
        'mayo': 'Mayo', 'junio': 'Junio', 'julio': 'Julio', 'agosto': 'Agosto',
        'septiembre': 'Septiembre', 'octubre': 'Octubre', 'noviembre': 'Noviembre', 'diciembre': 'Diciembre'
    }
    nombre_archivo_lower = nombre_archivo.lower()
    for mes_key, mes_display in meses_map.items():
        if mes_key in nombre_archivo_lower:
            return mes_display
    return 'Mes no encontrado'

# --- Configuración principal ---
# ¡Ruta base actualizada!
#ruta_base = Path("C:/Users/vicga/OneDrive - Instituto Colombiano de Bienestar Familiar/CF_2025/")

ruta_base = Path("C:/Users/jhoham.garzon/Documents/Diciembre_2025/CODIGO_ACTAS/CF_2025_VF3/")



cells_to_extract = ['Q7', 'X7', 'H8', 'U8', 'J9', 'J10', 'J11', 'J12', 'J13', 'E29', 'L29', 'W29', 'F30', 'L30',
                    'Q30', 'X30', 'C32', 'H32', 'M32', 'R32', 'W32', 'C34', 'H34', 'M34', 'R34', 'W34', 'C36',
                    'H36', 'M36', 'R36', 'W36', 'G37', 'U37', 'I38', 'N38', 'U38', 'G39', 'M39', 'U39', 'G40',
                    'U40', 'G41', 'U41', 'G42', 'U42', 'G43', 'U43', 'G44', 'U44', 'G45', 'U45', 'G46', 'U46',
                    'C49', 'C50', 'C51', 'G49', 'G50', 'G51', 'J49', 'J50', 'J51', 'M49', 'M50', 'M51', 'O49',
                    'O50', 'O51', 'Q49', 'Q50', 'Q51', 'S49', 'S50', 'S51', 'U49', 'U50', 'U51', 'W49', 'W50',
                    'W51', 'G54', 'G55', 'P54', 'P55', 'T55', 'C57', 'E62', 'E63', 'E64', 'E65', 'E66', 'E67',
                    'E68', 'E69', 'E70', 'E71', 'E72', 'E73', 'H62', 'H63', 'H64', 'H65', 'H66', 'H67', 'H68',
                    'H69', 'H70', 'H71', 'H72', 'H73', 'N62', 'N63', 'N64', 'N65', 'N66', 'N67', 'N68', 'N69',
                    'N70', 'N71', 'N72', 'N73', 'Q62', 'Q63', 'Q64', 'Q65', 'Q66', 'Q67', 'Q68', 'Q69', 'Q70',
                    'Q71', 'Q72', 'Q73', 'X62', 'X63', 'X64', 'X65', 'X66', 'X67', 'X68', 'X69', 'X70', 'X71',
                    'X72', 'X73', 'C82', 'E85', 'H85', 'K85', 'N85', 'Q85', 'T85', 'W85', 'E87', 'E88', 'E89',
                    'E90', 'E91', 'E92', 'E93', 'E94', 'E95', 'E96', 'E97', 'E98', 'H87', 'H88', 'H89', 'H90',
                    'H91', 'H92', 'H93', 'H94', 'H95', 'H96', 'H97', 'H98', 'N87', 'N88', 'N89', 'N90', 'N91',
                    'N92', 'N93', 'N94', 'N95', 'N96', 'N97', 'N98', 'Q87', 'Q88', 'Q89', 'Q90', 'Q91', 'Q92',
                    'Q93', 'Q94', 'Q95', 'Q96', 'Q97', 'Q98', 'T87', 'T88', 'T89', 'T90', 'T91', 'T92', 'T93',
                    'T94', 'T95', 'T96', 'T97', 'T98', 'W87', 'W88', 'W89', 'W90', 'W91', 'W92', 'W93', 'W94',
                    'W95', 'W96', 'W97', 'W98', 'C102', 'G102', 'L102', 'Q102', 'V102', 'C104', 'U106', 'U107',
                    'U109', 'U110', 'U112', 'U113', 'U114', 'U115', 'U117', 'U118', 'U119', 'U120', 'E121', 'O121',
                    'V121', 'C123', 'F127', 'F128', 'F129', 'F130', 'F131', 'F132', 'F133', 'F134', 'F135', 'F136',
                    'F137', 'F138', 'I127', 'I128', 'I129', 'I130', 'I131', 'I132', 'I133', 'I134', 'I135', 'I136',
                    'I137', 'I138', 'L127', 'L128', 'L129', 'L130', 'L131', 'L132', 'L133', 'L134', 'L135', 'L136',
                    'L137', 'L138', 'C141', 'C150', 'F150', 'H150', 'L151', 'L152', 'L154', 'L155', 'L156', 'O150',
                    'O151', 'O152', 'O154', 'O155', 'O156', 'R150', 'U150', 'X150', 'C161', 'F165', 'F166', 'F167',
                    'F168', 'F169', 'F170', 'F171', 'F172', 'F173', 'F174', 'F175', 'F176', 'I165', 'I166', 'I167',
                    'I168', 'I169', 'I170', 'I171', 'I172', 'I173', 'I174', 'I175', 'I176', 'L165', 'L166', 'L167',
                    'L168', 'L169', 'L170', 'L171', 'L172', 'L173', 'L174', 'L175', 'L176', 'R165', 'C180', 'C188']

lista_xlsm = encontrar_archivos_xlsm(ruta_base)


all_data_celdas = []
all_data_pdfs = []


current_date = datetime.now().strftime("%d/%m/%Y")


for file_path_str in tqdm(lista_xlsm, desc="Procesando archivos XLSM"):
    file_path = Path(file_path_str)
    ruta_directorio = file_path.parent
    ultima_subcarpeta = ruta_directorio.name

    try:
        wb = openpyxl.load_workbook(file_path, data_only=True)
        if 'DICIEMBRE' in wb.sheetnames:
            sheet = wb['DICIEMBRE']

            
            extracted_values_celdas = [str(file_path), current_date]
            num_pdfs = contar_pdfs_en_ruta(ruta_directorio)
            extracted_values_celdas.append(num_pdfs)

            for cell in cells_to_extract:
                cell_value = sheet[cell]
                if isinstance(cell_value, tuple):
                    extracted_values_celdas.append(str(cell_value))
                else:
                    extracted_values_celdas.append(cell_value.value)
            all_data_celdas.append(extracted_values_celdas)

            
            valor_w29 = sheet['W29'].value
            archivos_pdf_en_directorio = listar_pdfs_en_ruta(ruta_directorio)

            if archivos_pdf_en_directorio:
                for pdf_name in archivos_pdf_en_directorio:
                    mes = extraer_mes_del_nombre(pdf_name)
                    extracted_values_pdfs = [str(file_path), current_date, ultima_subcarpeta, pdf_name, mes, valor_w29]
                    all_data_pdfs.append(extracted_values_pdfs)
            else:
                extracted_values_pdfs = [str(file_path), current_date, ultima_subcarpeta, 'No PDF encontrado', 'Mes no encontrado', valor_w29]
                all_data_pdfs.append(extracted_values_pdfs)

        else:
            extracted_values_celdas = [str(file_path), current_date, 0] + ["Hoja 'DICIEMBRE' no encontrada"] * len(cells_to_extract)
            all_data_celdas.append(extracted_values_celdas)

            extracted_values_pdfs = [str(file_path), current_date, ultima_subcarpeta, 'No PDF encontrado', 'Mes no encontrado', 'Hoja "DICIEMBRE" no encontrada']
            all_data_pdfs.append(extracted_values_pdfs)

    except Exception as e:
        print(f"Error procesando {file_path_str}: {e}")
        extracted_values_celdas = [str(file_path), current_date, 0] + ["Archivo dañado/Error"] * len(cells_to_extract)
        all_data_celdas.append(extracted_values_celdas)
        extracted_values_pdfs = [str(file_path), current_date, ultima_subcarpeta, 'Error', 'Mes no encontrado', 'Error']
        all_data_pdfs.append(extracted_values_pdfs)



columns_celdas = ['File Path', 'Fecha', 'Número de PDFs'] + cells_to_extract
df_celdas = pd.DataFrame(all_data_celdas, columns=columns_celdas)

columns_pdfs = ['File Path', 'Fecha', 'Última Subcarpeta', 'Nombre PDF', 'Mes PDF', 'W29']
df_pdfs = pd.DataFrame(all_data_pdfs, columns=columns_pdfs)

ruta_base.mkdir(parents=True, exist_ok=True)

output_path_celdas = ruta_base / 'Extraccion_Valores_Celdas.xlsx'
output_path_pdfs = ruta_base / 'Detalle_PDFs_Encontrados.xlsx'

df_celdas.to_excel(output_path_celdas, index=False)
df_pdfs.to_excel(output_path_pdfs, index=False)

print(f"\nArchivo de valores de celdas guardado en: {output_path_celdas}")
print(f"Archivo de detalle de PDFs guardado en: {output_path_pdfs}")