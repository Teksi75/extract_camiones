from __future__ import annotations

import re
from pathlib import Path


def _read_version() -> str:
    """
    Lee la versión desde pyproject.toml y la devuelve con prefijo 'v'.
    Si no la encuentra, devuelve 'v0.0.0'.
    """
    root = Path(__file__).resolve().parents[1]
    pp = root / "pyproject.toml"
    if pp.exists():
        text = pp.read_text(encoding="utf-8")
        m = re.search(r'(?m)^\s*version\s*=\s*["\']([^"\']+)["\']', text)
        if m:
            return f"v{m.group(1)}"
    return "v0.0.0"


APP_VERSION: str = _read_version()
