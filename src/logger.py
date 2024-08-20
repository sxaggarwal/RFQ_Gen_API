import sys
import loguru

LOGGER = None


def getlogger(name: str = "DefaultName", level="DEBUG"):
    """Initialize logging for app returning logger."""
    global LOGGER  # pylint: disable=global-statement

    if LOGGER is None:
        logobj = loguru.logger
        logobj.remove()
        logger_format = (
            "<green>{time:YYYY-MM-DD HH:mm:ss.SSS}</green> | "
            f"{name} | "
            "<level>{level: <8}</level> | "
            "<cyan>{name}</cyan>:<cyan>{function}</cyan>:<cyan>{line}</cyan> | "
            "<level>{message}</level>"
        )
        logobj.add(
            sys.stderr,
            level=level,
            format=logger_format,
            colorize=None,
            serialize=False,
        )
        LOGGER = logobj

    return LOGGER