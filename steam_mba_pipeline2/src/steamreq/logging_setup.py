"""Logging verbose em console + persistente em arquivo (item 9 do projeto).

Regra explicita do item 9: nunca imprimir HTML/JSON gigante. Os helpers deste
modulo reportam tamanhos e contagens, nao conteudo.
"""
from __future__ import annotations

import logging
import os
import sys
import time
from datetime import datetime, timezone

LOGGER_NAME = "steamreq"
_ELAPSED_START = time.monotonic()


class _StageFormatter(logging.Formatter):
    """Formata console de forma legivel; arquivo de forma completa."""

    def __init__(self, *, console: bool) -> None:
        if console:
            super().__init__(fmt="%(message)s")
        else:
            super().__init__(
                fmt="%(asctime)s | %(levelname)-7s | %(name)s | %(message)s",
                datefmt="%Y-%m-%d %H:%M:%S",
            )
        self.console = console

    def format(self, record: logging.LogRecord) -> str:
        msg = super().format(record)
        if self.console and record.levelno >= logging.WARNING:
            return f"{record.levelname:>7}: {msg}"
        return msg


def setup(log_dir: str, *, console_level: str = "INFO",
          file_level: str = "DEBUG", run_id: str | None = None) -> logging.Logger:
    """Configura o logger raiz do pacote. Idempotente."""
    logger = logging.getLogger(LOGGER_NAME)
    if getattr(logger, "_steamreq_configured", False):
        return logger

    logger.setLevel(logging.DEBUG)
    logger.propagate = False

    ch = logging.StreamHandler(stream=sys.stdout)
    ch.setLevel(getattr(logging, console_level.upper(), logging.INFO))
    ch.setFormatter(_StageFormatter(console=True))
    logger.addHandler(ch)

    os.makedirs(log_dir, exist_ok=True)
    run_id = run_id or datetime.now().strftime("%Y%m%d_%H%M%S")
    log_path = os.path.join(log_dir, f"run_{run_id}.log")
    fh = logging.FileHandler(log_path, encoding="utf-8")
    fh.setLevel(getattr(logging, file_level.upper(), logging.DEBUG))
    fh.setFormatter(_StageFormatter(console=False))
    logger.addHandler(fh)

    logger._steamreq_configured = True  # type: ignore[attr-defined]
    logger._steamreq_log_path = log_path  # type: ignore[attr-defined]
    logger.debug("logging inicializado | arquivo=%s", log_path)
    return logger


def get_logger(name: str = "") -> logging.Logger:
    return logging.getLogger(f"{LOGGER_NAME}.{name}" if name else LOGGER_NAME)


def log_path() -> str | None:
    return getattr(logging.getLogger(LOGGER_NAME), "_steamreq_log_path", None)


# --- helpers de formatacao verbose -----------------------------------------

def fmt_elapsed(seconds: float | None = None) -> str:
    """Formata duracao como HH:MM:SS."""
    s = int(time.monotonic() - _ELAPSED_START if seconds is None else seconds)
    return f"{s // 3600:02d}:{(s % 3600) // 60:02d}:{s % 60:02d}"


def fmt_eta(done: int, total: int, elapsed_s: float) -> str:
    """Estimativa de conclusao, apenas quando ha base razoavel (item 9)."""
    if done <= 0 or total <= 0 or done >= total:
        return "n/d"
    per_item = elapsed_s / done
    return f"{fmt_elapsed((total - done) * per_item)} (~{per_item:.2f}s/app)"


def fmt_progress(done: int, total: int) -> str:
    pct = (100.0 * done / total) if total else 0.0
    width = len(str(total))
    return f"{done:0{width}d}/{total} | {pct:.1f}%"


def fmt_bytes(n: int) -> str:
    for unit in ("B", "KB", "MB", "GB"):
        if n < 1024 or unit == "GB":
            return f"{n:.1f} {unit}" if unit != "B" else f"{n} B"
        n /= 1024.0
    return f"{n:.1f} GB"


def utc_now_iso() -> str:
    return datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
