# Estructura de Datos del Repositorio

## `data/` — Insumos y archivos de entrada

El directorio `data/` está en `.gitignore` y contiene los archivos sensibles/pesados.

| Subcarpeta / Archivo | Contenido |
|---|---|
| `insumos metaverso/` | Bases del Metaverso: `Metaverso 2026.xlsx`, `UDS_15052026_3112025.xlsx`, zonificaciones, CONCAT |
| `insumos 28 abril/` | Insumos de planeación recibidos el 28 de abril |
| `insumos 4 junio/` | Insumos de planeación recibidos el 4 de junio |
| `insumos 19 junio/` | Insumos de planeación recibidos el 19 de junio |
| `insumos matriz 24 abril/` | Consolidados regionales y CUE de abril |
| `actas_20_05_2026/` | Actas de cobro recibidas al 20 de mayo |
| `replicacion hcb 5 junio/` | Archivos HCB replicados el 5 de junio, incluye `REGIONALES/` con los archivos partidos por regional |
| `marts/` | Archivos de datos listos para consumo (marts de datos) |
| `Matrices_validadas_definitivas/` | Matrices validadas por las regionales |
| `monitoring/` | Reportes de monitoreo generados (`monitor_hcb.xlsx`) |
| `Plantilla_Matriz_Adición_Nivelacion_V3_...xlsx` | Plantilla maestra nacional |
| `MATRIZ_ADICIONES_HCB_NIVELACION_CANASTA_27032026V2 2.xlsm` | Plantilla HCB |
| `Matriz_Adicion_Nivelacion_Canasta_ HCB_01042026_V2.xlsm` | Plantilla HCB v2 |
| `A3_MATRIZ_4_REGIONAL_ CASANARE ...xlsm` | Matriz de Casanare (ejemplo regional) |
| `REPORTE_ACTAS_20-05-2026.xlsx` | Reporte consolidado de actas |
| `Seguimiento_matrices.xlsx` | Seguimiento de matrices recibidas |

## `artifacts/` — Salidas generadas

Directorio para archivos de salida generados por los pipelines (actualmente vacío, las salidas se guardan dentro de `data/` o en rutas absolutas definidas en cada script).

## `bi/` — Dashboards Power BI

| Archivo | Descripción |
|---|---|
| `Dashboard integrilidad 2026.pbix` | Dashboard ejecutivo de integralidad de datos |
| `seguimiento actas.pbix` | Seguimiento de actas de cobro y soportes PDF |
| `SEGUIMIENTO HCB.pbix` | Dashboard específico de HCB |
| `base_theme.json` | Tema visual corporativo para Power BI |
| `logo.jpg` | Logo institucional ICBF |

## Convención de nombres de archivos generados

| Patrón | Significado |
|---|---|
| `CONCAT_*.xlsx` | Concatenación de múltiples hojas/fuentes |
| `ZONIFICACIONPAIS_*_2026.xlsx` | Tabla maestra de zonificación para la vigencia 2026 |
| `consolidacion_matriz_integrales_*.xlsx` | Consolidado regional de Integrales |
| `MATRIZ_2026_[REGIONAL]_[MODALIDAD].xlsx` | Plantilla individual por regional y modalidad |
| `UDS_[FECHA].xlsx` | Base de UDS al corte de la fecha |
| `*_AUDITADO.xlsx` / `*_AUDITORIA*` | Reportes con marcas de auditoría |
