import sys
import logging


DEFAULT_FMT = "[%(name)s] - [%(levelname)s] - %(message)s"
DEFAULT_DATEFMT = "%Y-%m-%d %H:%M:%S"


def _normalize_level(level):
    """
    Convierte strings como 'info' o 'DEBUG' en constantes de logging.
    """

    if isinstance(level, int):
        return level

    if isinstance(level, str):
        lvl = getattr(logging, level.upper(), None)

        if isinstance(lvl, int):
            return lvl
    # Fallback
    return logging.INFO


def _is_pytest_capture_handler(handler: logging.Handler) -> bool:
    """
    Detecta el handler de captura de logs de pytest para poder eliminarlo.
    """

    cls = handler.__class__

    return (
        cls.__name__ == "LogCaptureHandler"
        and cls.__module__.startswith("_pytest.")
    )


def init_logging(
    level: str | int = "INFO",
    fmt: str = DEFAULT_FMT,
    datefmt: str = DEFAULT_DATEFMT,
) -> None:
    """
    Configura el logging global:

    - Nivel en el root logger.
    - Si NO hay handlers "reales", crea un StreamHandler(stdout).
    - Si SÍ hay handlers "reales", reutiliza esos y solo actualiza su formatter.
    - Elimina los handlers de pytest del root.
    - Actualiza el nivel de los loggers de RuntimeConfig.* ya existentes.
    """

    root = logging.getLogger()

    # 1) Nivel
    root.setLevel(_normalize_level(level))

    # 2) Formatter
    formatter = logging.Formatter(fmt, datefmt=datefmt)

    # 3) Separa handlers reales de los de pytest
    real_handlers: list[logging.Handler] = []
    pytest_handlers: list[logging.Handler] = []

    for h in root.handlers:
        if _is_pytest_capture_handler(h):
            pytest_handlers.append(h)

        else:
            real_handlers.append(h)

    if real_handlers:
        # Caso B: ya había handlers "reales" -> NO añadimos uno nuevo
        for h in real_handlers:
            h.setFormatter(formatter)

    else:
        # Caso A: solo había handlers de pytest (o ninguno) -> creamos el nuestro
        for h in list(root.handlers):
            root.removeHandler(h)

        stream_handler = logging.StreamHandler(sys.stdout)
        stream_handler.setFormatter(formatter)
        root.addHandler(stream_handler)

    # En cualquier caso, limpiamos handlers de pytest que queden
    for h in pytest_handlers:
        if h in root.handlers:
            root.removeHandler(h)

    # 4) Actualizamos el nivel de loggers ya creados del paquete RuntimeConfig
    for name, logger in logging.Logger.manager.loggerDict.items():
        if not isinstance(logger, logging.Logger):
            continue

        if name.startswith("RuntimeConfig"):
            logger.setLevel(root.level)


def get_logger(name: str) -> logging.Logger:
    """
    Devuelve un logger con el nombre dado, que propaga al root
    para aprovechar la configuración global.
    """

    logger = logging.getLogger(name)
    # Aseguramos que no rompa la propagación al root
    logger.propagate = True
    # No añadimos handlers aquí: solo usamos los del root
    return logger
