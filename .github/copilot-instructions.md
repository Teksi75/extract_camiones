---
### Ejemplo: probar el scrapper desde la GUI en modo headless

1. Ejecuta la GUI:
   ```powershell
   python -m src.ui.gui
   ```
2. Ingresa credenciales y OT en el formulario.
3. Marca la casilla "Ejecutar en modo oculto (headless)" para evitar abrir el navegador.
4. Haz clic en "INICIAR EXTRACCIÓN" y verifica el progreso y el log.

Esto permite probar el flujo completo del scrapper sin abrir Chromium, útil para validaciones rápidas y CI.
## Contexto rápido

Proyecto: extractor de datos desde MetroWeb (INTI) → Excel. Componentes principales:
- `src/portal/scraper.py`: lógica de Playwright (sincrónico). Función clave: `extraer_camiones_por_ot` y helpers (`login_y_abrir_ot`, `leer_resumen`, `leer_instrumento`).
- `src/io/excel_exporter.py`: formatea filas en hoja 2-columnas y exporta con `xlsxwriter`. Orden de columnas en `COLUMNS_ORDER`.
- `src/ui/gui.py`: GUI Tkinter (ejecución recomendada: `python -m src.ui.gui`). Usa `extraer_camiones_por_ot` y puede anexar hoja con `src.ui.excel_merge.append_sheet_as_first`.
- `src/ui/excel_merge.py`: mezcla una hoja `datos vpe` como PRIMERA hoja en una COPIA del libro base (usa `openpyxl`).
- `src/domain/*`: modelos de datos (`models.py`) y parsing de domicilios (`address.py`).
- `selectors.yaml`: mapa de variantes de etiquetas del portal (usar/actualizar antes que tocar el scraper).

## Qué necesita saber un agente AI para ser productivo

1. Entradas / puntos de ejecución (IMPORTANTE: CLI descontinuada)
   - GUI (único punto de avance): `python -m src.ui.gui` (Tkinter — muestra dependencias faltantes en arranque). Enfocar cambios y pruebas en la GUI y los helpers que ella usa.
   - Nota: la versión CLI (`src/cli.py`) está descontinuada — no implementar nuevas funcionalidades ni mantener la CLI. Mantener `scraper` y `io` compatibles con la GUI.
   - Tests: `pytest -q` (hay mocks para Playwright en tests). Evitar lanzar navegadores reales en pruebas unitarias.
   - Instalación WebDriver para ejecución real: `python -m playwright install chromium` (solo necesaria para runs integrales, no para tests mockeados).

2. Dependencias relevantes (ver `requirements.txt`)
   - imprescindible: `playwright`, `pandas`, `xlsxwriter`, `openpyxl`, `pillow`, `tk` (Tkinter), `pytest`, `pytest-mock`.
   - Evitar usar la GUI en CI sin headless/browser instalado.

3. Patrones y convenciones del proyecto (detectables en el código)
   - Scraper: usa Playwright sync API y devuelve strings ya limpiados (`_clean_one_line`, `td_value`, `td_values`). Preferir esos helpers para nuevos scrapers.
   - Texto/selector mapping: `selectors.yaml` contiene variantes de etiquetas; modificarlo para soportar cambios del portal antes de tocar `scraper.py`.
   - Fechas: se normalizan en `src/io/excel_exporter.py` a formato castellano con `_fecha_castellano` — procesar fechas antes de exportar.
   - Excel: todas las celdas se escriben como texto para evitar fórmulas; si añades nuevas columnas respeta `COLUMNS_ORDER` o actualiza `selectors.yaml` export.columnas.
   - Merge Excel: `append_sheet_as_first` crea una COPIA; detecta archivo bloqueado y lanza PermissionError — pruebas/edits deben manejar archivo abierto en Excel.
   - Validaciones UI: la validación de OT está en `gui.validar_formato_ot` (regex `^\d{3}-\d{5}$`). Úsala para inputs automáticos.

4. Errores/edge-cases ya tratados por el código
   - Selectores variantes y limpieza de cadenas (acentos, NBSP) para tolerancia de HTML.
   - Fechas en múltiples formatos y heurística para años de 2 dígitos.
   - Archivo Excel abierto: `_is_file_locked` en `excel_merge.py`.

5. Recomendaciones concretas para modificaciones por un agente
   - Si un campo del portal no se encuentra: primero agregar variantes en `selectors.yaml` y reintentar; solo tocar `scraper.py` si el DOM cambió.
   - Al añadir nuevas columnas al Excel: agregar a `COLUMNS_ORDER` y a `selectors.yaml:export:columnas` para consistencia entre scrapper/export y UI.
   - Para cambios UI: reuse `find_project_root` pattern en `gui.py` si se quiere ejecutar módulos 'a pelo'.
   - Para pruebas: usa `pytest` y `pytest-mock` (hay tests que mockean Playwright). Evitar lanzar navegadores reales en pruebas unitarias.

6. Archivos y funciones a inspeccionar para tareas comunes
   - Extraer datos / debug scraper: `src/portal/scraper.py` (`login_y_abrir_ot`, `leer_resumen`, `leer_instrumento`, `extraer_camiones_por_ot`).
   - Exportar Excel: `src/io/excel_exporter.py` (`armar_hoja_verificacion_2columnas`, `exportar_verificacion_2columnas`).
   - Merge Excel: `src/ui/excel_merge.py` (`append_sheet_as_first`).
   - Parse de direcciones: `src/domain/address.py` (`parse_domicilio_fiscal`).
   - Config labels: `selectors.yaml` — prefer editar este archivo para cambios de etiquetas del sitio.

7. Comandos útiles (ejemplos locales)
   - Instalar deps: `pip install -r requirements.txt`
   - Instalar Playwright browsers (solo para runs reales): `python -m playwright install chromium`
   - Ejecutar GUI (flujo de trabajo principal): `python -m src.ui.gui`
   - Test rápido: `pytest tests/test_excel_export.py -q`

Si algo queda ambiguo (p. ej. credenciales en CI, path de Excel base o cómo simular Playwright en integración), dime qué flujo quieres que detalle y lo amplío. Después puedo ajustar el tono/longitud según prefieras.
