# -*- coding: utf-8 -*-
from pathlib import Path

# --- CONFIGURACIÓN DE RUTAS ---
# Usamos rutas relativas para que el código funcione en cualquier entorno (Windows/Docker)
ROOT_DIR = Path(__file__).resolve().parent.parent.parent

# 1. Carpeta raíz donde se encuentran todos los subdirectorios de contratos
BASE_PATH = ROOT_DIR / "data" / "Actas"

# 2. Archivo Excel maestro que se desea actualizar con los resultados
MASTER_FILE_PATH = ROOT_DIR / "data" / "TRAER_DATO_PDF_POR_MES.xlsx"

# --- CONFIGURACIÓN DE COLUMNAS (MAPPING) ---
# Aquí puedes asignar nombres legibles a los datos extraídos de las celdas de Excel.
# El sistema usará estos nombres como cabeceras en el reporte final.
MAP_CELDAS = {
    'W29': 'LLAVE',                # N° de Contrato/Convenio (ID único)
    'U46': 'VALOR_MENSUAL',        # Valor Actual del CTO
    'J13': 'NOMBRE_CONTRATISTA',
    'U1': 'FECHA_DOCUMENTO',
    'C32': 'MODALIDAD_ATENCION',
    'C34': 'TIPO_SERVICIO',
    'L25': 'NOMBRE_UDS',
    'L29': 'MUNICIPIO_ATENCION',
    'G102': 'VALOR_EJECUTADO',
    'G112': 'SALDO_DISPONIBLE',
    'W36': 'CONTADOS_C36',         # Cupos contratados
    'X30': 'CUPOS_TOTALES',
}

# Alias para duplicar columnas en el reporte final
COL_ALIASES = {
    'NUMERO_CONTRATO': 'LLAVE',
    'VALIDACION_PDF': 'LLAVE'
}

# --- LISTA DE CELDAS A EXTRAER ---
# Lista exhaustiva basada en los requerimientos del usuario
CELLS_TO_EXTRACT_ALL = [
    'Q7', 'X7', 'J9', 'J10', 'J11', 'J12', 'F30', 'L30', 'Q30', 'X30',
    'C32', 'H32', 'M32', 'R32', 'W32', 'C34', 'H34', 'M34', 'R34', 'W34',
    'C36', 'H36', 'M36', 'R36', 'W36', 'G37', 'U37', 'I38', 'N38', 'U38',
    'G39', 'M39', 'U39', 'G40', 'U40', 'G41', 'U41', 'G42', 'U42', 'G43',
    'U43', 'G44', 'U44', 'G45', 'U45', 'G46', 'U46',
    'C49', 'C50', 'C51', 'G49', 'G50', 'G51', 'J49', 'J50', 'J51',
    'M49', 'M50', 'M51', 'O49', 'O50', 'O51', 'Q49', 'Q50', 'Q51',
    'S49', 'S50', 'S51', 'U49', 'U50', 'U51', 'W49', 'W50',
    'G102', 'L112', 'L114', 'C159', 'L163'
]

# Construcción final de la lista de extracción (Preferencia a las mapeadas)
CELLS_TO_EXTRACT = list(MAP_CELDAS.keys()) + [c for c in CELLS_TO_EXTRACT_ALL if c not in MAP_CELDAS]

# --- CONFIGURACIÓN DE HOJAS ---
SHEET_NAME_CONTRATOS = 'CONTRATOS'
SHEET_NAME_DATOS = 'DATOS'

# --- CONFIGURACIÓN DE COLUMNAS INTERNAS ---
COL_LLAVE = 'LLAVE'  # Este nombre debe coincidir con el valor en MAP_CELDAS
COL_CONTRATO_ENCONTRADO = 'CONTRATO ENCONTRADO'
COL_MES = 'Mes'

# --- CONFIGURACIÓN DE MESES ---
MESES_MAP = {
    'enero': 'PDF ENERO', 'febrero': 'PDF FEBRERO', 'marzo': 'PDF MARZO',
    'abril': 'PDF ABRIL', 'mayo': 'PDF MAYO', 'junio': 'PDF JUNIO',
    'julio': 'PDF JULIO', 'agosto': 'PDF AGOSTO', 'septiembre': 'PDF SEPTIEMBRE',
    'octubre': 'PDF OCTUBRE', 'noviembre': 'PDF NOVIEMBRE', 'diciembre': 'PDF DICIEMBRE'
}

# --- SALIDAS PARCIALES ---
OUTPUT_VALORES_CELDAS = 'Extraccion_Basica_Fase1.xlsx'
OUTPUT_DETALLE_PDFS = 'Detalle_PDFS_Fase1.xlsx'
