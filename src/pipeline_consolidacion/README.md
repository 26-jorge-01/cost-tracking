# Pipeline de Consolidación ICBF

Este directorio contiene los scripts necesarios para generar el archivo maestro de consolidación de regionales.

## Orden de Ejecución

1.  **`01_consolidacion_inicial.py`**:
    *   Escanea la carpeta de insumos.
    *   Filtra archivos relevantes (Modality: Integrales).
    *   Genera el archivo `consolidacion_matriz_integrales_28042026.xlsx`.
    
2.  **`02_refinamiento_copia_segura.py`**:
    *   Toma el archivo del paso anterior.
    *   Busca regionales que hayan llegado tarde o falten en el maestro.
    *   Genera la versión final: `consolidacion_matriz_integrales_28042026_COPIA_SEGURA.xlsx`.
    
3.  **`03_auditoria_calidad.py`**:
    *   Analiza el archivo final en busca de filas de "TOTAL" que causen doble conteo.
    *   Verifica la integridad de los contratos y cupos.

## Requisitos
*   Python 3.x
*   pandas
*   openpyxl
