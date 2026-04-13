"""
logger_setup.py
---------------
Configures application-wide logging for WorkLog.
Creates a rotating log file in logs/ next to the app.
Also streams INFO+ messages to the console for easy debugging.
"""

import logging
import os
from logging.handlers import RotatingFileHandler

# Resolve the logs/ directory relative to this file, regardless of where
# the user runs the script from.
_APP_DIR = os.path.dirname(os.path.abspath(__file__))
_LOG_DIR = os.path.join(_APP_DIR, "logs")
_LOG_FILE = os.path.join(_LOG_DIR, "worklog.log")


def setup_logging(level=logging.DEBUG):
    """
    Call ONCE at application startup.

    Sets up:
      - A rotating file handler (max 1 MB, keeps 3 backup files)
      - A console handler that shows INFO and above

    Returns the root 'worklog' logger.
    """
    os.makedirs(_LOG_DIR, exist_ok=True)

    logger = logging.getLogger("worklog")
    logger.setLevel(level)

    # Guard against duplicate handlers if called more than once
    if logger.handlers:
        return logger

    formatter = logging.Formatter(
        fmt="%(asctime)s [%(levelname)-8s] %(name)s: %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )

    # --- File handler (DEBUG+) ---
    try:
        file_handler = RotatingFileHandler(
            _LOG_FILE,
            maxBytes=1_000_000,   # 1 MB
            backupCount=3,
            encoding="utf-8",
        )
        file_handler.setLevel(logging.DEBUG)
        file_handler.setFormatter(formatter)
        logger.addHandler(file_handler)
    except Exception as exc:
        # Non-fatal: app can run without a log file
        print(f"[WARNING] Could not create log file at {_LOG_FILE}: {exc}")

    # --- Console handler (INFO+) ---
    console_handler = logging.StreamHandler()
    console_handler.setLevel(logging.INFO)
    console_handler.setFormatter(formatter)
    logger.addHandler(console_handler)

    logger.info("Logging initialized. Log file: %s", _LOG_FILE)
    return logger


def get_logger(name: str) -> logging.Logger:
    """
    Return a child logger under the 'worklog' namespace.

    Usage in every other module:
        from logger_setup import get_logger
        log = get_logger(__name__)
    """
    return logging.getLogger(f"worklog.{name}")
