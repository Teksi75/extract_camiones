from __future__ import annotations

import re
import sys
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
PYPROJECT = ROOT / "pyproject.toml"


def bump_patch() -> None:
    """
    Incrementa en 1 el último componente de la versión declarada en pyproject.toml.

    Ejemplo:
        0.4.1 -> 0.4.2
        1.0   -> 1.1   (también soporta versiones con 2 partes)
    """
    if not PYPROJECT.exists():
        print(f"[ERROR] No se encontró {PYPROJECT}", file=sys.stderr)
        sys.exit(1)

    text = PYPROJECT.read_text(encoding="utf-8")

    # Busca una línea del estilo: version = "0.4.1"
    pattern = r'(^\s*version\s*=\s*["\'])([^"\']+)(["\']\s*$)'
    m = re.search(pattern, text, flags=re.MULTILINE)
    if not m:
        print(
            "[ERROR] No se encontró la línea 'version = \"...\"' en pyproject.toml",
            file=sys.stderr,
        )
        sys.exit(1)

    old_version = m.group(2).strip()
    parts = old_version.split(".")

    try:
        numbers = [int(p) for p in parts]
    except ValueError:
        print(
            f"[ERROR] La versión actual '{old_version}' no es numérica.",
            file=sys.stderr,
        )
        sys.exit(1)

    # Incrementar el último componente
    numbers[-1] += 1
    new_version = ".".join(str(n) for n in numbers)

    # Reemplazar solo el valor de la versión
    new_text = text[: m.start(2)] + new_version + text[m.end(2) :]

    PYPROJECT.write_text(new_text, encoding="utf-8")

    # Evitar caracteres no ASCII para no romper en consolas con cp1252
    print(f"Version actualizada: {old_version} -> {new_version}")


if __name__ == "__main__":
    bump_patch()
