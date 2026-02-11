# Dockerfile (multi-stage)
# ==== Etapa base: depende de pyproject.toml, uv.lock ====
FROM python:3.12-slim-bookworm AS base
ENV DEBIAN_FRONTEND=noninteractive

# Java para Spark + utilidades
RUN apt-get update && \
    apt-get install -y --no-install-recommends \
      openjdk-17-jre-headless curl git openssh-server ca-certificates && \
    rm -rf /var/lib/apt/lists/*

# JAVA_HOME (opcional pero útil)
ENV JAVA_HOME=/usr/lib/jvm/java-17-openjdk-amd64
ENV PATH="${JAVA_HOME}/bin:${PATH}"
WORKDIR /app

# --- Instalar uv (Astral) correctamente ---
# (evita el paquete homónimo de PyPI)
RUN curl -LsSf https://astral.sh/uv/install.sh | sh

# uv se instala por defecto en /root/.local/bin
ENV PATH="/root/.local/bin:${PATH}"

# (Comprobación en build; si falla, el build se detiene)
RUN which uv && uv --version

# Copia sólo definiciones de deps para aprovechar caché
COPY pyproject.toml uv.lock* ./

# Copia todo el código del monorepo
COPY . .

# Resolver dependencias del proyecto en .venv del directorio (por defecto de uv)
# --frozen para respetar uv.lock si está presente
RUN uv lock
RUN uv sync --frozen --no-dev

# Asegura que .venv quede primero en PATH para pip/python posteriores
ENV PATH="/app/.venv/bin:${PATH}"

# Instala streamlit dentro del .venv (puedes añadirlo a pyproject y omitir esta línea)
RUN pip install --no-cache-dir streamlit

# Instala torch (CPU o GPU) dentro del .venv según ARG
ARG USE_CUDA=0
RUN if [ "$USE_CUDA" = "1" ]; then \
      pip install --no-cache-dir torch torchvision torchaudio --index-url https://download.pytorch.org/whl/cu118 ; \
    else \
      pip install --no-cache-dir torch torchvision torchaudio --index-url https://download.pytorch.org/whl/cpu ; \
    fi

# Configura SSH para dev remoto
RUN mkdir -p /var/run/sshd && \
    echo 'PermitRootLogin yes' >> /etc/ssh/sshd_config && \
    echo 'PasswordAuthentication yes' >> /etc/ssh/sshd_config && \
    ssh-keygen -A

# Prepara carpeta de datos
RUN mkdir -p /app/data
VOLUME ["/app/data"]

# ==== Etapa prod: código + defaults ====
FROM base AS prod
WORKDIR /app

# Puertos: Jupyter, API, SSH y Streamlit
EXPOSE 8888 8000 2222 8501

# CMD por defecto (Jupyter con uv usando el .venv del proyecto)
CMD ["uv", "run", "jupyter", "lab", "--ip=0.0.0.0", "--port=8888", "--no-browser", "--allow-root"]