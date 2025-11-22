# 🏭 Extractor de datos MetroWeb → Excel (INTI)

**Versión actual:** Alpha 0.4.0
**Autor:** Pablo J. Siklosi  
****  

Aplicación desarrollada en Python para **extraer automáticamente los datos de Verificación Previa** desde el portal **MetroWeb (INTI)** y volcarlos en un archivo **Excel estructurado**.  
Permite obtener información de las **balanzas para camiones/plataforma**, incluyendo detalles del instrumento, modelo, aprobación, fabricante y propietario.

---

## 🚀 Características principales

- ✅ **Extracción automática** desde MetroWeb mediante Playwright (Chromium).  
- 💾 **Exportación directa a Excel** en formato de dos columnas (*Campo | Valor*).  
- 🧩 **Interfaz gráfica (GUI)** moderna con barra de progreso y registro en tiempo real.  
- 🧠 **Procesamiento multi-instrumento:** reconoce múltiples instrumentos dentro de una misma OT.  
- 🧱 **Arquitectura modular:** separa la lógica de scraping, exportación y GUI.  
- 🔒 Compatible con **Windows 10/11** y **Python 3.11–3.13**.

---

## 📂 Estructura del proyecto

extract_camiones/
├── assets/                 # Recursos gráficos usados por la GUI
│   └── balanza.png
├── src/
│   ├── cli.py              # Entrypoint de línea de comandos
│   ├── domain/             # Modelos de dominio y helpers de direcciones
│   ├── io/                 # Exportadores (Excel)
│   ├── portal/             # Scraper de MetroWeb basado en Playwright
│   ├── ui/                 # Interfaz gráfica y herramientas de merge
│   └── version.py          # Datos de versionado de la aplicación
├── tests/                  # Pruebas unitarias e integraciones básicas
├── tools/                  # Utilidades de mantenimiento y build
├── selectors.yaml          # Mapeo de selectores MetroWeb
├── requirements.txt        # Dependencias mínimas
└── pyproject.toml          # Configuración de build
