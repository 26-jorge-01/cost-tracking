# Scripts Utilitarios

## Reporte de Auditoría UDS

### `src/generar_reporte_auditoria_uds.py`
Genera un reporte **multi-contrastado** de UDS cruzando tres fuentes:

1. **Base UDS** (`UDS_15042026.xlsx`) — Cupos reales por UDS y contrato
2. **Consolidado Regional** (`Regionales Contratacion Vigente 2026_CONSOLIDADO.xlsx`) — Cupos reportados por las regionales
3. **Reporte CUE** (`ICBFCUEContUnicoxRegxVigxDir.xlsx`) — Cupos del sistema financiero CUE

**Lógica de auditoría**:
- Calcula diferencia entre cupos UDS y cupos reportados en el consolidado regional.
- Calcula diferencia entre cupos CUE y cupos UDS por contrato.
- Aplica **filtro de primera ocurrencia** para evitar duplicar montos financieros.
- **Salida**: `UDS_15042026_Multi_Contrastado.xlsx`

---

## Vinculación de PDFs a Contratos (Fase 2 del Pilar 3)

### `src/orm/buscar_pdf_2.py`
Toma el reporte de Fase 1 (extracción de celdas de actas) y lo **enriquece con la existencia de PDFs** firmados.

**Proceso**:
1. Carga el reporte de valores de celdas y el detalle de PDFs escaneados.
2. Vincula cada contrato con los meses en que tiene PDF soporte.
3. Aplica pivote: crea columnas por mes (ENERO, FEBRERO, ... DICIEMBRE) con valores "SI" / "N" según exista PDF.
4. Calcula `% CUMPLIMIENTO` = meses con PDF / total meses.
5. Calcula `POSIBLE_INEJECUCION`: si el valor pagado acumulado es menor al valor mensual esperado acumulado del mes anterior.
6. **Salida**: Archivo maestro con hojas `CONTRATOS` y `DETALLE_PDFS`.

### `src/orm/config.py`
Diccionario de configuración con las coordenadas de celdas a extraer de las actas (ej. celda W29 para N° Contrato). Centraliza los parámetros para que si cambia el formato del acta, solo se edite este archivo.

---

## Utilitarios de Cruce y Limpieza

### `src/cruces_plantilla_matriz_contratacion/utils/OpenDataTools.py`
Clase `OpenDataTools` para consultar la **API del SECOP** (Sistema de Contratación Pública de Colombia) a través de Socrata/`www.datos.gov.co`.

**Métodos**:
- `get_contratos_por_cedulas(cedulas)` — Busca contratos por cédula del contratista o representante legal
- `get_contratos_por_nits(nits)` — Busca contratos por NIT de la entidad
- `get_contratos_por_nombre_proveedor(nombres)` — Busca por razón social
- `get_contratos_por_num_contrato(nums_contrato)` — Busca por número de contrato

**Fuentes SECOP**:
| Dataset | ID |
|---|---|
| SECOP I (histórico) | `f789-7hwg` |
| SECOP II (nuevo) | `jbjy-vk9h` |
| SECOP II Procesos | `p6dx-8zbt` |
| TVEC (Consolidado) | `rgxm-mmea` |

### `src/cruces_plantilla_matriz_contratacion/utils/CleanData.py`
Funciones de limpieza de datos: URLs, fechas, textos.

### `src/cruces_plantilla_matriz_contratacion/utils/logging_setup.py`
Configuración centralizada de logging para el módulo de cruces.

---

## Exploración y Análisis (Jupyter Notebooks)

| Notebook | Ubicación | Propósito |
|---|---|---|
| `Analisis_Insumos_Matriz_24_Abril.ipynb` | `src/` | Análisis de insumos recibidos el 24 de abril, incluye lógica NIT Bridge y Fuzzy Matching |
| `orchestrate_pdf_process.ipynb` | `src/` | Orquestación del proceso de extracción de PDFs/actas |
| `process_excel.ipynb` | `src/` | Procesamiento de archivos Excel de actas |
| `split_hcb_by_regional.ipynb` | `src/` | Versión notebook del split HCB (alternativa al .py) |
| `monitor_hcb.ipynb` | `src/monitoring/` | Versión notebook del monitor HCB |
| `reconstruccion_zonificacionpais_2026.ipynb` | `src/metaverso_pipeline/` | Exploración de la reconstrucción de ZonificacionPais |
| `cruce_planilla_secop.ipynb` | `src/cruces_plantilla_matriz_contratacion/` | Cruce de planilla SECOP |
| `zonificacion_vs_oper_directa.ipynb` | `src/cruces_plantilla_matriz_contratacion/` | Comparación zonificación vs. operación directa |
