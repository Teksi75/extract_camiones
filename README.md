# Extractor de datos MetroWeb -> Excel (INTI)

Version actual: 0.4.14 (segun `pyproject.toml`)
Autor: Pablo J. Siklosi

Aplicacion en Python para extraer automaticamente los datos de Verificacion Previa desde el portal MetroWeb (INTI) y volcarlos en un Excel estructurado. Incluye un flujo GUI con barra de progreso y logs en vivo, y utilidades para versionar y generar releases distribuidos.

## Requisitos
- Windows 10/11
- Python 3.13 o superior
- Dependencias de `requirements.txt` (incluye Playwright, pandas, xlsxwriter, etc.)
- Navegador Playwright: `python -m playwright install chromium`

## Instalacion rapida
1) Clonar el repositorio y ubicarse en la raiz del proyecto.
2) Crear y activar el entorno: `python -m venv .venv` y luego `./.venv/Scripts/Activate.ps1` (o equivalente en tu shell).
3) Instalar dependencias: `pip install -r requirements.txt`.
4) Instalar el navegador de Playwright si aun no se hizo: `python -m playwright install chromium`.

## Uso
### Ejecutar la GUI
- Desde la raiz del proyecto (con el entorno activado):
  ```bash
  python -m src.ui.gui
  ```
- Se abrira la interfaz grafica para lanzar el scraping y generar el Excel usando la plantilla de `assets/`.

### Scripts de mantenimiento
- `python -m tools.bump_version`: incrementa en +1 el ultimo componente de `version` en `pyproject.toml` (por ejemplo, 0.4.14 -> 0.4.15). Ejecutalo siempre desde la raiz del repo.
- `python -m tools.make_release`: genera un ZIP listo para distribuir dentro de `tools/dist/`, excluyendo tests y artefactos temporales. Usa la version de `pyproject.toml` para nombrar el archivo.

### Flujo de release sugerido
1) Actualizar la version: `python -m tools.bump_version`.
2) Opcional: ejecutar pruebas rapidas si aplica.
3) Empaquetar: `python -m tools.make_release` (el ZIP se guarda en `tools/dist/`).

## Estructura del proyecto (resumen)
- `assets/`: recursos graficos y plantilla de Excel.
- `src/cli.py`: entrypoint CLI.
- `src/domain/`: modelos y helpers de direcciones.
- `src/io/`: exportadores a Excel.
- `src/portal/`: scraper MetroWeb basado en Playwright.
- `src/ui/`: GUI y herramientas de merge de Excel.
- `src/version.py`: metadatos de version de la app.
- `tools/`: scripts de soporte (`bump_version.py`, `make_release.py`, `shrink_balanza.py`); `tools/dist/` guarda los ZIP generados.
- `tests/`: pruebas basicas.
- `selectors.yaml`: mapeo de selectores de MetroWeb.
- `pyproject.toml`: metadata del proyecto y versionado.

## Notas
- Ejecuta los comandos desde la raiz del repositorio para que las rutas relativas (assets, selectors) funcionen correctamente.
- Si actualizas Playwright o los navegadores, reinstala con `python -m playwright install chromium` antes de correr la GUI.
