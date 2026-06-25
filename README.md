# ICBF Cost-Tracking — Sistema de Control de Costos y Auditoría

## ¿Qué es este repositorio?

Sistema que automatiza la recolección, consolidación, auditoría y visualización de **cupos y recursos de Primera Infancia del ICBF** (Instituto Colombiano de Bienestar Familiar). Integra datos de planeación de **33 Regionales**, cobertura real (*Cuéntame*), contratos públicos (*SECOP I y II*), inventario físico de sedes (*Metaverso*) y registros financieros (*Compromisos*).

## ¿Para quién es?

- **Directivos ICBF** — Tableros ejecutivos para decisiones sobre cupos y presupuesto.
- **Analistas de contratación** — Cruce y auditoría de contratos vs. ejecución real.
- **Auditores fiscales** — Verificación de actas, soportes PDF y legalidad de pagos.
- **Desarrolladores** — Mantenimiento y mejora de los pipelines de datos.

## Estructura del repositorio

| Carpeta / Archivo | ¿Qué contiene? |
|---|---|
| `docs/` | Documentación detallada de cada pilar del sistema |
| `src/` | Código fuente: pipelines, extractores, cruces |
| `data/` | Insumos y archivos de entrada (git-ignored) |
| `artifacts/` | Salidas generadas por los pipelines |
| `bi/` | Dashboards Power BI y temas visuales |
| `pilot_automation_v1/` | Scripts piloto de automatización de plantillas |
| `Dockerfile` | Imagen Docker para entorno reproducible |
| `make.bat` | Comandos para build, Jupyter, API y más |
| `pyproject.toml` | Dependencias y configuración del proyecto |

## Pilares del sistema

1. **Consolidación Regional** — Unifica archivos de 33 regionales eliminando duplicados.
2. **Plantillas y Distribución** — Genera archivos individuales protegidos por regional.
3. **Extracción de Actas** — Procesa miles de actas de cobro y vincula PDFs soporte.
4. **Censo Metaverso** — Reconstruye el inventario nacional de UDS con semaforización.
5. **Cruce SECOP / NIT Bridge** — Concilia contratos contra bases del Estado.
6. **Dashboards** — Tableros web y Power BI para la alta dirección.

Ver la documentación completa en [`BUSINESS_PROCESS_DOCUMENTATION.md`](BUSINESS_PROCESS_DOCUMENTATION.md).

## Tecnologías principales

- **Python 3.12** — pandas, openpyxl, sodapy, thefuzz, nltk
- **PowerShell** — Script de automatización COM para Excel
- **Docker** — Contenedor reproducible con Jupyter, SSH, Streamlit
- **Power BI** — Tableros corporativos
- **Chart.js** — Dashboard web interactivo

## Cómo empezar

```bash
# Instalar dependencias
uv sync

# O con Docker
make.bat build
make.bat jupyter
```

## Licencia

Proyecto interno del ICBF — Dirección de Primera Infancia.
