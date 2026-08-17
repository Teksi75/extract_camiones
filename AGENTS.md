# AGENTS.md - extract_camiones

## Project Overview
Python app that extracts verification data from MetroWeb (INTI) portal and exports to Excel. Requires Python 3.13+.

## Key Commands

| Task | Command |
|------|---------|
| Run GUI | `python -m src.ui.gui` |
| Bump version | `python -m tools.bump_version` |
| Make release | `python -m tools.make_release` |
| Run tests | `pytest -q` |
| Install browsers | `python -m playwright install chromium` |

## Environment Setup
```powershell
python -m venv .venv
./.venv/Scripts/Activate.ps1
pip install -r requirements.txt
python -m playwright install chromium
```

## Architecture
- `src/portal/scraper.py` - Playwright scraping logic (sync API)
- `src/io/excel_exporter.py` - Excel export with xlsxwriter
- `src/ui/gui.py` - Tkinter GUI entry point
- `src/ui/excel_merge.py` - Excel merge utilities
- `selectors.yaml` - DOM selector variants (edit before touching scraper)

## Important Conventions

1. **CLI is deprecated** - Use only `python -m src.ui.gui`. Don't implement new CLI features.

2. **Selector changes** - Edit `selectors.yaml` first before modifying `scraper.py` for DOM changes.

3. **Excel columns** - Maintain `COLUMNS_ORDER` in `excel_exporter.py` and `selectors.yaml:export:columnas` in sync.

4. **Tests use mocks** - Playwright is mocked in unit tests. Don't launch real browsers in `pytest`.

5. **Excel locks** - `append_sheet_as_first` creates a copy. Running tests with Excel open causes `PermissionError`.

6. **Headless mode** - Use headless for CI: check the "Ejecutar en modo oculto" checkbox in GUI or pass headless flag.

## Linting & Typecheck
- Ruff: `ruff check .`
- Black: `black --check .`
- Mypy: `mypy src/`

Run in order: `ruff check . && black --check . && mypy src/`

## Release Flow
1. `python -m tools.bump_version` → increments patch version
2. `python -m tools.make_release` → creates ZIP in `tools/dist/`
3. Git commit `pyproject.toml` only
4. Tag: `git tag vX.Y.Z && git push --tags`

## Common Issues
- **OT validation**: Regex `^\d{3}-\d{5}$` in `gui.validar_formato_ot`
- **Date format**: Use `_fecha_castellano` from `excel_exporter.py`
- **NBSP/encoding**: Use `_clean_one_line`, `td_value` helpers in scraper