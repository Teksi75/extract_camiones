# 🧾 Extractor de datos MetroWeb → Excel (Versión 2025)
**Autor:** Pablo J. Siklosi — INTI  
**Rama:** `version-2025`

---

## 📌 Descripción general

La **Versión 2025** del proyecto `extract_camiones` es una edición simplificada y directa del extractor de datos desde MetroWeb (INTI) hacia planillas Excel utilizadas en tareas operativas de Verificación Previa y Control Metrológico.

El objetivo principal de esta versión es ofrecer un **script único, fácil de ejecutar y mantener**, pensado para uso cotidiano del verificador, sin necesidad de instalar una estructura compleja de paquetes.

Esta versión incluye:

- Un ejecutable principal: **`extract_camiones_gui.py`**  
- Plantilla Excel base para completar automáticamente: **`307-xxxxx_para_rellenar.xlsx`**  
- Scripts experimentales y módulos auxiliares organizados en carpetas  
- Notebooks y archivos intermedios usados durante el desarrollo  

> 🔎 La rama `main` mantiene la versión modular “profesional” (con `/src`, `/tools`, etc.).  
> Esta rama `version-2025` prioriza la **simplicidad operativa**.

---

## 🧰 Funcionalidades principales

- ✔ Automatiza la extracción de datos de MetroWeb (OT / VPE).  
- ✔ Procesa balanzas de **camiones / plataforma**.  
- ✔ Completa automáticamente planillas Excel con los datos del instrumento.  
- ✔ Interfaz gráfica simple (botones, selección de archivo, barra de progreso).  
- ✔ Compatible con **Python 3.12 / 3.13 + Playwright (Chromium)**.  

---

## 📂 Estructura del proyecto (rama `version-2025`)

```text
versión-2025/
│
├── extract_camiones_gui.py        ← Script principal (GUI)
├── 307-xxxxx_para_rellenar.xlsx   ← Planilla Excel base
├── requirements.txt               ← Librerías necesarias
├── .gitignore
│
├── maps/                          ← Mapas, imágenes, datos auxiliares
│
├── notebooks/
│   └── extraccion_de_precintos_25.ipynb   ← Notebook de desarrollo
│
├── old_scripts/                   ← Scripts experimentales / archivados
│   ├── extract_camiones.py
│   ├── extract_camiones_1.py
│   ├── extract_camiones_01_cosola.py
│   ├── extract_camiones_01_deepseek.py
│   ├── extract_camiones_02_deepseek.py
│   ├── extract_balanzas_qwen.py
│   ├── fill_identificacion.py
│   └── metroweb_scraper_limpio.py
│
└── outputs/
    └── OT_final.xlsx              ← Ejemplo de salida generada
🚀 Cómo ejecutar el extractor
1. Crear entorno virtual (opcional pero recomendado)
python -m venv .venv


Activar el entorno:

Windows (PowerShell):

.venv\Scripts\Activate.ps1


Windows (cmd):

.venv\Scripts\activate.bat

2. Instalar dependencias
pip install -r requirements.txt

3. Instalar Playwright (si no está instalado)
playwright install chromium

4. Ejecutar la interfaz gráfica
python extract_camiones_gui.py

🧮 Flujo de uso

Ejecutar extract_camiones_gui.py.

Seleccionar el archivo base 307-xxxxx_para_rellenar.xlsx.

Ingresar el número de OT / VPE según corresponda.

El sistema abre MetroWeb (Chromium mediante Playwright), navega la OT/VPE y extrae:

datos del instrumento

modelo

capacidad, división, clase, etc.

datos del propietario

Los datos se vuelcan a la planilla Excel respetando el formato utilizado en INTI.

El archivo resultante puede guardarse junto a las planillas oficiales.

🧪 Scripts y notebooks de apoyo

old_scripts/: contiene versiones previas del extractor y pruebas con distintos enfoques (DeepSeek, Qwen, scraping limpio, scripts de relleno de Excel, etc.).
No forman parte del flujo principal, pero se conservan para referencia y debugging.

notebooks/: incluye notebooks usados para:

prototipar el scraping,

probar selectores de MetroWeb,

experimentar con lógicas de completado de planillas.

🏗️ Relación con la rama main

main: versión modular y empaquetable (src/, tools/, assets/, pyproject.toml).

version-2025: versión simplificada, basada en un script principal y estructura mínima.

La idea futura es:

Unificar lo mejor de ambas ramas:

arquitectura modular de main,

usabilidad y simplicidad de version-2025.

Definir una próxima versión estable (por ejemplo, v0.5.0) que integre ambos enfoques.

✅ Estado actual

Rama version-2025 en uso para trabajo operativo y pruebas.

Rama main como base estable de la versión empaquetable (Alpha v0.4.0).

Próximo paso: evaluar qué flujo se adopta como estándar en INTI y empaquetar una versión distribuible para otros verificadores.