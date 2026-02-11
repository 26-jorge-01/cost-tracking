@echo off
SETLOCAL

REM ——————————————————————————————
REM Asegura que exista la carpeta local data\
REM ——————————————————————————————
if not exist "%cd%\data" (
    mkdir "%cd%\data"
)

REM ——————————————————————————————
REM BUILD
REM ——————————————————————————————
IF "%1"=="build" (
    docker build --network=host --progress=plain -t nm-platform-starter .
    GOTO :EOF
)

REM ——————————————————————————————
REM BUILD SIN CACHÉ
REM ——————————————————————————————
IF "%1"=="rebuild" (
    docker build --no-cache -t nm-platform-starter .
    GOTO :EOF
)

REM ——————————————————————————————
REM JUPYTER LAB
REM ——————————————————————————————
IF "%1"=="jupyter" (
    docker run -it --rm ^
        -p 8888:8888 ^
        -v "%cd%:/app" ^
        -v "%cd%/data:/app/data" ^
        nm-platform-starter ^
        uv run jupyter lab --ip=0.0.0.0 --port 8888 --no-browser --allow-root
    GOTO :EOF
)

REM ——————————————————————————————
REM API (Uvicorn)
REM ——————————————————————————————
IF "%1"=="api" (
    docker run -it --rm ^
        -p 8000:8000 ^
        -v "%cd%:/app" ^
        -v "%cd%/data:/app/data" ^
        nm-platform-starter ^
        uv run uvicorn src.api.main:app --host 0.0.0.0 --port 8000
    GOTO :EOF
)

REM ——————————————————————————————
REM SSH REMOTO
REM ——————————————————————————————
IF "%1"=="ssh" (
    docker run -d --name nm-platform-ssh ^
        -p 2222:2222 ^
        -v "%cd%:/app" ^
        -v "%cd%/data:/app/data" ^
        nm-platform-starter && ^
    echo Contenedor SSH corriendo en localhost:2222.
    GOTO :EOF
)

REM ——————————————————————————————
REM UI MVP PRODUCCIÓN
REM ——————————————————————————————
IF "%1"=="quality_tool" (
    echo 🧼 Iniciando UI MVP en http://localhost:8501
    docker run --rm ^
        -p 8501:8501 ^
        -v "%cd%:/app" ^
        -v "%cd%/data:/app/data" ^
        nm-platform-starter ^
        streamlit run packages/DataQualityPipelineBuilder/streamlit_app/main.py --server.port=8501 --server.address=0.0.0.0
    GOTO :EOF
)

REM ——————————————————————————————
REM UI MVP DESARROLLO (hot-reload)
REM ——————————————————————————————
IF "%1"=="quality_tool-dev" (
    echo 🧼 UI Streamlit en modo desarrollo (hot-reload)
    docker run -it --rm ^
        -p 8501:8501 ^
        -v "%cd%:/app" ^
        -v "%cd%/data:/app/data" ^
        nm-platform-starter ^
        streamlit run packages/DataQualityPipelineBuilder/streamlit_app/main.py --server.port=8501 --server.address=0.0.0.0
    GOTO :EOF
)

REM ——————————————————————————————
REM REBUILD DE LA IMAGEN
REM ——————————————————————————————

:rebuild
echo 🧼 Rebuild de la imagen
docker build --no-cache -t nm-platform-starter .
goto :eof

echo Comando no reconocido: %1
echo Opciones válidas: build, rebuild, jupyter, api, ssh, quality_tool, quality_tool-dev