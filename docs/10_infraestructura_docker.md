# Infraestructura y Entorno de Desarrollo

## Docker

### `Dockerfile`
Imagen multi-etapa basada en `python:3.12-slim-bookworm`.

**Componentes instalados:**
- Java 17 (OpenJDK) — para compatibilidad con Spark
- `uv` (Astral) — gestor de dependencias rápido
- Dependencias Python desde `pyproject.toml` vía `uv sync`
- Streamlit — para UI de calidad de datos
- PyTorch (CPU o GPU según `USE_CUDA=1`)
- OpenSSH Server — para desarrollo remoto

**Puertos expuestos:**
| Puerto | Servicio |
|---|---|
| 8888 | Jupyter Lab |
| 8000 | API (Uvicorn) |
| 2222 | SSH |
| 8501 | Streamlit |

**Uso:**
```bash
make.bat build      # Construir imagen
make.bat jupyter    # Iniciar Jupyter Lab
make.bat api        # Iniciar API
make.bat ssh        # Iniciar servidor SSH
make.bat quality_tool    # UI Streamlit
```

### `.devcontainer/devcontainer.json`
Configuración para **VS Code Dev Containers**:
- Imagen: `nm-platform-starter`
- Monta el código y `data/` como volúmenes bind
- Puertos: 8888 (Jupyter), 8000 (API), 2222 (SSH)
- Extensiones: Python, Jupyter
- Intérprete predeterminado: `/app/.venv/bin/python`
- Post-start: `uv sync`

---

## Configuración del proyecto

### `pyproject.toml`
- **Nombre**: `nm-platform-starter` v0.1.0
- **Python**: >=3.12, <3.13
- **Dependencias principales**: pandas 2.1.1, openpyxl 3.1.5, numpy 1.26.4, sodapy 2.2.0, nltk, chardet, ipykernel
- **Dev**: pytest, black, flake8, pre-commit, isort

### `make.bat`
Comandos para Windows:
| Comando | Acción |
|---|---|
| `make build` | `docker build` |
| `make rebuild` | `docker build --no-cache` |
| `make jupyter` | Inicia Jupyter Lab en contenedor |
| `make api` | Inicia API Uvicorn |
| `make ssh` | Inicia SSH en contenedor |
| `make quality_tool` | UI Streamlit |
| `make quality_tool-dev` | Streamlit con hot-reload |

### `Generar_Reportes_ICBF.ps1`
Script PowerShell que **parte el archivo maestro por regional** usando Excel COM:
- Detecta automáticamente las 33 regionales.
- Para cada una, abre el Excel, filtra/elimina filas de otras regionales, y guarda como archivo `.xlsm` protegido.
- Contraseña: `GMYC2023***`

### `.gitignore`
Ignora: `data/`, `.env`, `__pycache__/`, `.venv/`, `.vscode/`, archivos `.xlsx`, `.csv`, `.pbix`, logs.
