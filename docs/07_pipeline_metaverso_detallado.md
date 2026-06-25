# Pipeline Detallado del Metaverso (Censo Nacional de UDS)
### Complemento técnico del Pilar 4

## ¿Qué hace este pipeline?

Toma los insumos de planeación 2026 (Comunitarios, Integrales y Alimentos) y los fusiona con la base histórica del Metaverso para generar el censo nacional actualizado de Unidades de Servicio (UDS). El resultado es un archivo Excel con el estado de cada UDS, incluyendo cupos, contratistas, NIT, y fórmulas de validación.

---

## Scripts en orden de ejecución

### 1. `concat_zonificacion_metaverso_2026.py`
**Concatena las hojas Comunitarios e Integrales** de la zonificación en un solo archivo.
- **Entrada**: `zonificación_abastecimiento_servicios_primera_infancia25052026.xlsx` (hojas: `Comunitarios`, `Integrales-Convenios`)
- **Salida**: `CONCAT_ZONIFICACION_METAVERSO_2026.xlsx`
- **Lógica**: Lee cada hoja con `header=None`, asigna nombres de columna fijos, añade columna `Tipo_Modalidad` (COMUNITARIO / INTEGRAL) y concatena ambas.

### 2. `concat_alimentos_2026.py`
**Concatena las hojas ZONIFICACION y ZONIFICACION LA GUAJIRA** del archivo de Alimentos.
- **Entrada**: `CONTRATOS ALIMENTOS ORGANIZACIONES CAMPESINAS.xlsx`
- **Salida**: `CONCAT_ALIMENTOS_2026.xlsx`
- **Lógica**: Lee cada hoja con `header=None`, asigna nombres de columna según diccionario, estandariza `Codigo_UDS` y unifica en un solo DataFrame.

### 3. `enrich_zonificacion_con_uds.py`
**Enriquece el CONCAT con datos de la base UDS_31122025**.
- **Entrada**: `CONCAT_ZONIFICACION_METAVERSO_2026.xlsx`, `UDS_31122025.xlsx`, `Metaverso 2026.xlsx` (template)
- **Salida**: `ZONIFICACION_ENRIQUECIDA_2026.xlsx`
- **Lógica**: Mergea por código UDS, mapea columnas de CONCAT y UDS hacia la estructura del template `ZonificacionPais`, rellenando vacíos de CONCAT con datos UDS. Incluye limpieza de NIT y cálculo de modalidad.

### 4. `rebuild_zonificacionpais_2026.py`
**Reconstruye la tabla ZonificacionPais desde cero**.
- **Entrada**: CONCAT, UDS_15052026 (incluye TH_Transito), UDS_31122025, Metaverso 2026.xlsx
- **Salida**: `ZONIFICACIONPAIS_RECONSTRUIDA_2026.xlsx`
- **Lógica**: Mergea CONCAT con UDS, TH_Transito y UDS_31122025; mapea columnas; calcula modalidad; inyecta fórmulas EXACT.

### 5. `update_zonificacionpais_2026.py`
**Actualiza incrementalmente la ZonificacionPais original** con datos nuevos.
- **Entrada**: CONCAT, UDS_15052026, UDS_31122025, Metaverso 2026.xlsx
- **Salida**: `ZONIFICACIONPAIS_ACTUALIZADA_2026.xlsx`
- **Lógica**: Toma la tabla original, mergea datos nuevos por código UDS, marca filas como `Actualizado`, `No actualizable` o `Nuevo`. Agrega columnas `CuposUDS_31122025`, `CuposServicioUDS_31122025`, `TH_Transito_Cupos` y `Estado_Actualizacion`.

### 6. `update_zonificacionpais_integrales.py`
**Actualiza registros "No actualizable" usando la hoja ZONIFICACION-PEDAGOGICO** de Integrales.
- **Entrada**: `ZONIFICACIONPAIS_ACTUALIZADA_2026.xlsx`, `consolidacion_matriz_integrales_28042026_COPIA_SEGURA.xlsx`
- **Salida**: Sobrescribe `ZONIFICACIONPAIS_ACTUALIZADA_2026.xlsx`
- **Lógica**: Busca coincidencias de código UDS entre los "No actualizable" y la hoja de Integrales; actualiza columnas (Regional, Centro Zonal, Municipio, Servicio, UDS). Marca como `Actualizado x integrales`.

### 7. `update_zonificacionpais_alimentos.py`
**Actualiza registros aún pendientes usando la base de Alimentos**.
- **Entrada**: `ZONIFICACIONPAIS_ACTUALIZADA_2026.xlsx`, `CONCAT_ALIMENTOS_2026.xlsx`, `UDS_15052026_3112025.xlsx`
- **Salida**: Sobrescribe `ZONIFICACIONPAIS_ACTUALIZADA_2026.xlsx` + `REPORTE_NO_ACTUALIZABLES.xlsx` + `NO_ACTUALIZABLES_PERSISTENTES.xlsx`
- **Lógica**: Match de códigos UDS entre los "No actualizable"/"Actualizado x integrales" y la base de Alimentos; actualiza servicio, cupos, NIT, contratista. Marca como `Actualizado x alimentos`. Genera reportes de los que siguen sin coincidencia.

### 8. `_explore_mapping.py` y `_explore_th.py`
Scripts de exploración/analysis (prefijo `_` indica que son auxiliares). El primero explora el mapeo de columnas entre fuentes; el segundo analiza la tabla TH_Transito.

---

## Flujo completo recomendado

```
concat_zonificacion_metaverso_2026.py  (Paso 1: unificar Comunitarios + Integrales)
concat_alimentos_2026.py               (Paso 2: unificar Alimentos)
rebuild_zonificacionpais_2026.py       (Paso 3: reconstrucción base)
update_zonificacionpais_2026.py        (Paso 4: actualización incremental)
update_zonificacionpais_integrales.py  (Paso 5: rescate vía Integrales)
update_zonificacionpais_alimentos.py   (Paso 6: rescate vía Alimentos)
```

## Columnas clave de salida

| Columna | Significado |
|---|---|
| `Codigo Unidad Servicio UDS` | Identificador único de cada sede |
| `Estado_Actualizacion` | Actualizado / No actualizable / Nuevo / Actualizado x integrales / Actualizado x alimentos |
| `Modalidad 2026` | INTEGRAL / COMUNITARIO / FAMILIAR Y COMUNITARIA / ALIMENTOS |
| `Cupos a Programar 2026` | Cupos asignados para la vigencia |
| `NIT CONTRATISTA 2026` | NIT del operador contratado |
| `CONTRATISTA 2026` | Nombre de la entidad contratista |
| `SERVICIO 2026` | Nombre del servicio ofrecido en la UDS |
| `CuposUDS_31122025` | Cupos históricos al cierre de 2025 |
| `TH_Transito_Cupos` | Cupos en tránsito (talento humano) |
| `RIESGO TRANSPARENCIA` | Indicador de riesgo (columna auxiliar para fórmula EXACT) |
