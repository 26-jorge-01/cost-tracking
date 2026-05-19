# Pipeline de Consolidación HCB (Hogares Comunitarios)

Este pipeline se encarga de consolidar los archivos de la modalidad HCB.

## Diferencias con el Pipeline de Integrales:
- **Filtro de Archivos**: Solo incluye archivos con "HCB" en el nombre.
- **Exclusiones**: Se excluyen archivos de "Integrales" (los que no dicen HCB) y archivos de "Alimentos".
- **Hoja Objetivo**: Utiliza la hoja `MATRIZ`.
- **Encabezados**: Busca dinámicamente la columna `REGIONAL`.

## Orden de Ejecución:
1. `01_consolidacion_hcb.py`
2. `02_refinamiento_hcb.py`
3. `03_auditoria_hcb.py`
