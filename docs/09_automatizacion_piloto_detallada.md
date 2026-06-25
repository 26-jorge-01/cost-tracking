# Automatización Piloto de Plantillas Regionales
### Complemento técnico del Pilar 2

## ¿Qué hace este subsistema?

Genera **plantillas individuales de Excel** para cada una de las 33 regionales del ICBF, tomando como insumo el archivo maestro de Zonificación. Cada plantilla contiene solo los datos de su regional, con las fórmulas y formatos protegidos.

---

## Scripts en `pilot_automation_v1/`

### 1. `ORQUESTADOR_PILOTO_ICBF.ipynb`
**Cuaderno Jupyter que orquesta todo el proceso piloto**:
1. **Configuración** — Define rutas, contraseña (`#APOLO1704*`) y funciones de limpieza.
2. **Inicialización** — Crea carpetas para las 33 regionales basándose en el maestro.
3. **Generación de plantillas (COM)** — Usa `win32com` (Excel nativo) para:
   - Abrir la plantilla nacional base.
   - Desproteger hojas con contraseña.
   - Insertar datos de cada regional (Integrales y Comunitarios).
   - Limpiar filas sobrantes del template original.
   - Guardar como archivo independiente en la carpeta de cada regional.
4. **Validación** — Aplica reglas de negocio (conteo de UDS, suma de cupos) y genera un reporte de auditoría CSV.
5. **Reporte ejecutivo** — Muestra resumen de estados (VALIDATED, ERROR, MISSING).

### 2. `generate_regional_templates.py`
**Versión openpyxl** (sin dependencia de Excel instalado).
- Usa `openpyxl` para copiar celdas y ajustar fórmulas.
- Recalcula referencias de fórmulas al insertar filas.
- No requiere Microsoft Excel en la máquina.

### 3. `generate_regional_templates_flexible.py`
**Versión COM avanzada** con mapeo explícito de columnas por modalidad.
- Define diccionarios `MODALITIES` con mapeo letra-columna → fuente.
- Soporta valores constantes (`const:1`, `const:2026`, `const:0.05`) para columnas de cálculo.
- Maneja Integrales y Comunitarios con reglas diferentes.
- Más rápido que openpyxl porque usa el motor nativo de Excel.

### 4. `generate_regional_templates_xlwings.py`
Variante que usa la librería `xlwings` como puente a Excel.

### 5. `generate_regional_templates_com.py`
Otra variante COM con ligeras diferencias en el manejo de sheets.

### 6. `inject_notebook.py`
Script que inyecta celdas de código en `ORQUESTADOR_PILOTO_ICBF.ipynb`. Automatiza la actualización del notebook cuando cambia la lógica de generación.

### 7. `update_notebook.py`
Actualiza configuraciones o parámetros dentro del notebook existente.

### 8. `test_formula.py`
Script de prueba para verificar que las fórmulas de Excel se calculan correctamente después de la generación.

---

## Configuración de modalidades

### Integrales
| Columna Excel | Origen en maestro | Descripción |
|---|---|---|
| A (ZONA) | Col 0 (TIPO) | Zona de la UDS |
| B (REGIONAL) | Col 1 | Regional UDS |
| C (CENTRO ZONAL) | Col 2 | Centro zonal |
| D (MUNICIPIO) | Col 3 | Municipio |
| H (ATENCIONES) | Col 10 | Cupos |
| K (UDS) | `const:1` | Cada fila = 1 UDS |

### Comunitarios (HCB)
| Columna Excel | Origen en maestro | Descripción |
|---|---|---|
| C (REGIONAL) | Col 0 | Regional UDS |
| K (CUPOS POR UNIDAD) | Col 10 | Cupos |
| L (CANTIDAD DE UDS) | Col 11 | Madres / Unds |
| AA (VIGENCIA) | `const:2026` | Año 2026 |
| CL (% APORTE EAS) | `const:0.05` | 5% de contrapartida |

---

## Contraseñas usadas

| Propósito | Contraseña |
|---|---|
| Protección de hojas en plantillas generadas | `#APOLO1704*` |
| Archivos HCB partidos por regional | `GMYC2026**` |
| Script PowerShell de distribución | `GMYC2023***` |
