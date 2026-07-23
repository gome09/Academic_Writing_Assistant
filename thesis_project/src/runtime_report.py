# -*- coding: utf-8 -*-
"""Small JSON run-report writer shared by CLI stages."""
from __future__ import annotations

import json
import os
from datetime import datetime, timezone


def write_report(path: str, report: dict) -> str:
    report = {"timestamp": datetime.now(timezone.utc).isoformat(), **report}
    os.makedirs(os.path.dirname(os.path.abspath(path)), exist_ok=True)
    with open(path, "w", encoding="utf-8") as fh:
        json.dump(report, fh, ensure_ascii=False, indent=2)
    return path
