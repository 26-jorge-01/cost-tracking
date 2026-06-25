# División Regional HCB y Sistema de Monitoreo

## 1. Split HCB por Regional (`src/split_hcb_by_regional/`)

### ¿Qué hace?

Toma un archivo Excel maestro de la modalidad **HCB (Hogares Comunitarios)** y lo divide en **archivos individuales por regional**, limpiando hojas innecesarias, desbloqueando tablas y protegiendo las celdas con contraseña.

### Script principal

**`split_regional.py`**

- **Entrada**: Archivo Excel maestro (ej. `MATRIZ_2026_ADICION_COMUNITARIOS_MODFV6.xlsx`)
- **Salida**: Archivos individuales por regional en subcarpetas dentro de `REGIONALES/` (ej. `REGIONALES/ANTIOQUIA/MATRIZ_ANTIOQUIA.xlsx`)
- **Proceso**:
  1. Escanea el archivo y detecta las regionales únicas.
  2. Crea un template limpio (solo las hojas necesarias: `MATRIZ`, `rp`, `COSTEO`, `servicios homologados`, `LISTAS`, `Hoja5`, `INTEGRALIDAD`).
  3. Para cada regional, copia el template, filtra solo sus filas y elimina las filas de otras regionales.
  4. Protege todas las hojas excepto `MATRIZ` con la contraseña `GMYC2026**`.
  5. Convierte nombres a mayúsculas sin acentos (ej. `BOGOTÁ D.C.` → `BOGOTA_D_C`).

### ¿Por qué es necesario?

- Las regionales **no deben ver datos de otras regionales** (confidencialidad).
- Las celdas con fórmulas críticas deben estar **protegidas** para evitar manipulaciones accidentales.
- El archivo original pesa mucho; al dividirlo se reduce el riesgo de corrupción.

---

## 2. Sistema de Monitoreo HCB (`src/monitoring/`)

### ¿Qué hace?

Escanea automáticamente todos los archivos regionales generados y produce un **reporte de calidad** en Excel con indicadores de integridad.

### Scripts

**`scan_hcb_files.py`**
- **Entrada**: Carpeta `REGIONALES/` con los archivos generados por `split_regional.py`
- **Salida**: `monitor_hcb.xlsx` (hoja `DATOS` con métricas + hoja `DICCIONARIO` con descripciones)
- **Métricas por archivo**:
  - `FILAS_DATOS`: cantidad de registros
  - `ERRORES_REF`: celdas con error `#REF!`
  - `VACIOS_CRITICOS`: celdas vacías en columnas clave (contrato, municipio, EAS, RP, servicio)
  - `CONTRATOS_UNICOS`, `MUNICIPIOS`, `EAS_UNICOS`: diversidad de datos
  - `CUPOS_TOTALES`: suma de cupos
  - `VALOR_INICIAL_TOTAL`, `VALOR_ADICIONAR_TOTAL`: montos financieros
  - `ALERTA_1000`, `ALERTA_5000`: contratos que requieren aval por superar 1000/5000 SMLMV
- **Semáforo de estado**:
  - 🟢 **VERDE**: Sin errores
  - 🟡 **AMARILLO**: Celdas vacías críticas
  - 🟠 **NARANJA**: Alertas de aval presentes
  - 🔴 **ROJO**: Errores `#REF!` en fórmulas

**`write_scanner.py`**
- Script auxiliar que extrae el código de `scan_hcb_files.py` y lo escribe en un archivo separado (útil para distribución del scanner como script independiente).
