"""
Logging setup with rotation.

Writes to logs/sync.log with 5 MB cap × 5 backups (plain text, no ANSI codes).
Also echoes to stdout — when stdout is a real terminal, output is colorized
(WARNING red, ERROR bright red, customer-match warnings highlighted, etc.).
When stdout is piped (Docker, redirect to file, etc.) colors auto-disable.
"""
from __future__ import annotations

import logging
import re
import sys
from logging.handlers import RotatingFileHandler
from pathlib import Path


# ANSI escape codes — only emitted when the terminal supports them.
_C_RESET   = "\033[0m"
_C_BOLD    = "\033[1m"
_C_RED     = "\033[31m"
_C_BRIGHT_RED = "\033[91m"
_C_YELLOW  = "\033[33m"
_C_GREEN   = "\033[32m"
_C_CYAN    = "\033[36m"
_C_GRAY    = "\033[90m"

# Phrases inside INFO messages worth highlighting in the terminal.
_INFO_HIGHLIGHTS = [
    (re.compile(r"(Fuzzy-matched|Near-matched)"), _C_CYAN),
    (re.compile(r"(\bSweep\b|\bCleanup\b|Authenticating|Loading\b)"), _C_GRAY),
    (re.compile(r"(---- Flow: [^-]+ ----)"), _C_BOLD),
    (re.compile(r"(Progress: \d+/\d+)"), _C_GREEN),
]


class ColoredFormatter(logging.Formatter):
    """Same format as the file formatter, plus ANSI colors based on level."""

    def __init__(self, fmt: str, datefmt: str, use_color: bool):
        super().__init__(fmt, datefmt=datefmt)
        self.use_color = use_color

    def format(self, record: logging.LogRecord) -> str:
        msg = super().format(record)
        if not self.use_color:
            return msg
        if record.levelno >= logging.ERROR:
            return f"{_C_BRIGHT_RED}{_C_BOLD}{msg}{_C_RESET}"
        if record.levelno >= logging.WARNING:
            # Warnings (incl. unmatched customer warnings) in red — caller asked
            # for unmatched lines to be visually obvious.
            return f"{_C_RED}{msg}{_C_RESET}"
        # INFO — colorize specific phrases in-place
        for pattern, color in _INFO_HIGHLIGHTS:
            msg = pattern.sub(lambda m: f"{color}{m.group(0)}{_C_RESET}", msg)
        return msg


def setup_logging(log_dir: Path, level: int = logging.INFO) -> logging.Logger:
    logger = logging.getLogger("automation_worker")
    logger.setLevel(level)

    # Avoid duplicate handlers on re-invocation (e.g. tests)
    if logger.handlers:
        return logger

    plain_fmt = logging.Formatter(
        "%(asctime)s %(levelname)-7s %(name)s — %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )

    # File handler with rotation — plain text always (no ANSI in log files).
    file_handler = RotatingFileHandler(
        log_dir / "sync.log",
        maxBytes=5 * 1024 * 1024,  # 5 MB
        backupCount=5,
        encoding="utf-8",
    )
    file_handler.setFormatter(plain_fmt)
    logger.addHandler(file_handler)

    # Stdout handler — colored when running interactively, plain when piped
    # (Docker logs, `> file.log`, etc.). The isatty check auto-detects.
    use_color = bool(getattr(sys.stdout, "isatty", lambda: False)())
    color_fmt = ColoredFormatter(
        "%(asctime)s %(levelname)-7s %(name)s — %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
        use_color=use_color,
    )
    stream_handler = logging.StreamHandler(sys.stdout)
    stream_handler.setFormatter(color_fmt)
    logger.addHandler(stream_handler)

    return logger
