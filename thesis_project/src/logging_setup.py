# -*- coding: utf-8 -*-
"""Project logging configuration; console output remains backward compatible."""
from __future__ import annotations

import logging


def configure_logging(path: str) -> logging.Logger:
    logger = logging.getLogger("thesis_project")
    logger.setLevel(logging.INFO)
    logger.handlers.clear()
    handler = logging.FileHandler(path, encoding="utf-8")
    handler.setFormatter(logging.Formatter("%(asctime)s %(levelname)s %(message)s"))
    logger.addHandler(handler)
    return logger
