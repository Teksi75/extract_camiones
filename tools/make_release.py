# tools/make_release.py
from __future__ import annotations

import re
import time
from fnmatch import fnmatch
from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED

# --- rutas base ---
ROOT = Path(__file__).resolve().parents[1]   # raíz del proyecto
DIST = ROOT / "tools" / "dist"
DIST.mkdir(parents=True, exist_ok=True)

# --- configuración de exclusiones ---
EXCLUDE_DIRS_BY_NAME = {
    ".git", ".venv", "__pycache__", ".pytest_cache", ".vscode",
    "node_modules", "playwright", "UNKNOWN.egg-info",
}
EXCLUDE_PATTERNS = {
    "*.pyc", "*.pyo", "*.pyd", "*.so", "*.dll",
    "*.log", "*.tmp", "*.bak", "*.whl", "*.msi", "*.exe",
}
EXCLUDE_PATHS: set[Path] = {
    DIST,                   # <-- excluir carpeta de salida completa
}
EXCLUDE_TESTS = True

if EXCLUDE_TESTS:
    EXCLUDE_DIRS_BY_NAME.add("tests")

def project_version() -> str:
    pp = ROOT / "pyproject.toml"
    if pp.exists():
        m = re.search(r'(?m)^\s*version\s*=\s*["\']([^"\']+)["\']', pp.read_text(encoding="utf-8"))
        if m:
            return f"v{m.group(1)}"
    gui = ROOT / "src" / "ui" / "gui.py"
    if gui.exists():
        m = re.search(r'(?m)^\s*APP_VERSION\s*=\s*["\']([^"\']+)["\']', gui.read_text(encoding="utf-8"))
        if m:
            return m.group(1)
    return "v0"

def path_is_under(p: Path, base: Path) -> bool:
    try:
        p.resolve().relative_to(base.resolve())
        return True
    except Exception:
        return False

def should_skip(p: Path) -> bool:
    # excluir por ruta absoluta (carpetas completas)
    if any(path_is_under(p, ex) for ex in EXCLUDE_PATHS):
        return True
    # excluir por nombre de alguna carpeta del camino
    if any(part in EXCLUDE_DIRS_BY_NAME for part in p.parts):
        return True
    # excluir por patrón de nombre de archivo
    if any(fnmatch(p.name, pat) for pat in EXCLUDE_PATTERNS):
        return True
    return False

def build_zip() -> Path:
    ver = project_version()
    ts = time.strftime("%Y%m%d_%H%M%S")
    zip_path = DIST / f"extract_camiones_{ver}_{ts}.zip"

    # 1) Construir lista de candidatos ANTES de abrir el zip (evita auto-inclusión)
    files: list[Path] = []
    for p in ROOT.rglob("*"):
        if p.is_dir():
            continue
        if should_skip(p):
            continue
        files.append(p)

    # 2) Escribir el zip (sin incluirse a sí mismo)
    total_bytes = 0
    with ZipFile(zip_path, "w", compression=ZIP_DEFLATED, compresslevel=9) as z:
        for p in files:
            # seguridad extra: no incluir el zip en caso de coincidencia
            if p.resolve() == zip_path.resolve():
                continue
            arc = p.relative_to(ROOT)
            z.write(p, arc)
            try:
                total_bytes += p.stat().st_size
            except OSError:
                pass

    # Evitar caracteres no ASCII para no romper en consolas con cp1252
    print(f"OK -> {zip_path} ({total_bytes/1048576:.2f} MB, {len(files)} archivos)")
    preview_heaviest(zip_path, top=15)
    return zip_path

def preview_heaviest(zip_path: Path, top: int = 20) -> None:
    from zipfile import ZipFile
    entries = []
    with ZipFile(zip_path, "r") as z:
        for info in z.infolist():
            entries.append((info.filename, info.file_size))
    entries.sort(key=lambda x: x[1], reverse=True)
    print("\nTop archivos por tamaño dentro del ZIP:")
    for name, size in entries[:top]:
        print(f"{size/1048576:7.2f} MB  {name}")

if __name__ == "__main__":
    build_zip()
