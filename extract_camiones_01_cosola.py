# extract_camiones.py
# -*- coding: utf-8 -*-

"""
Extracción INTI MetroWeb → Excel (Verificación Previa) para balanzas de camiones/plataforma.
Requisitos:
  pip install playwright pandas xlsxwriter colorama
  python -m playwright install chromium

Ejecución:
  python extract_camiones.py --user USR --pass PWD --ot 307-62136 [--headless] [--out RUTA.xlsx]
"""

import re
import sys
import time
import argparse
import getpass
from pathlib import Path
from typing import Dict, List, Tuple, Optional

import pandas as pd
from playwright.sync_api import sync_playwright, BrowserContext, Page

# Colores para terminal
try:
    from colorama import init, Fore, Style, Back
    init(autoreset=True)
    COLORS_AVAILABLE = True
except ImportError:
    COLORS_AVAILABLE = False
    # Fallback sin colores
    class Fore:
        GREEN = BLUE = YELLOW = RED = CYAN = MAGENTA = WHITE = ""
    class Style:
        BRIGHT = RESET_ALL = ""
    class Back:
        BLACK = ""

BASE = "https://app.inti.gob.ar"

# =========================
# Utilidades de UI
# =========================

def print_header():
    """Imprime encabezado visual del script."""
    width = 70
    print("\n" + "=" * width)
    print(f"{Fore.CYAN}{Style.BRIGHT}{'INTI METROWEB - EXTRACTOR DE VERIFICACIÓN PREVIA':^{width}}{Style.RESET_ALL}")
    print(f"{Fore.CYAN}{'Balanzas de Camiones / Plataforma':^{width}}{Style.RESET_ALL}")
    print("=" * width + "\n")

def print_section(title: str):
    """Imprime título de sección."""
    print(f"\n{Fore.YELLOW}{Style.BRIGHT}▶ {title}{Style.RESET_ALL}")
    print(f"{Fore.YELLOW}{'─' * (len(title) + 2)}{Style.RESET_ALL}")

def print_success(msg: str):
    """Mensaje de éxito."""
    print(f"{Fore.GREEN}✓ {msg}{Style.RESET_ALL}")

def print_info(msg: str):
    """Mensaje informativo."""
    print(f"{Fore.CYAN}ℹ {msg}{Style.RESET_ALL}")

def print_warning(msg: str):
    """Mensaje de advertencia."""
    print(f"{Fore.YELLOW}⚠ {msg}{Style.RESET_ALL}")

def print_error(msg: str):
    """Mensaje de error."""
    print(f"{Fore.RED}✗ {msg}{Style.RESET_ALL}")

def print_progress(current: int, total: int, item_name: str = ""):
    """Barra de progreso simple."""
    percentage = (current / total * 100) if total > 0 else 0
    filled = int(percentage / 5)
    bar = "█" * filled + "░" * (20 - filled)
    item_info = f" - {item_name}" if item_name else ""
    print(f"{Fore.BLUE}[{bar}] {percentage:5.1f}% ({current}/{total}){item_info}{Style.RESET_ALL}", end="\r")
    if current == total:
        print()  # Nueva línea al terminar

def print_data_table(data: Dict[str, str], title: str = "Datos extraídos"):
    """Imprime datos en formato tabla."""
    print(f"\n{Fore.MAGENTA}{Style.BRIGHT}{title}:{Style.RESET_ALL}")
    max_key_len = max(len(k) for k in data.keys()) if data else 0
    for key, value in data.items():
        value_display = value[:60] + "." if len(value) > 60 else value
        value_display = value_display or f"{Fore.YELLOW}(vacío){Style.RESET_ALL}"
        print(f"  {Fore.WHITE}{key:<{max_key_len}}{Style.RESET_ALL} : {value_display}")

# =========================
# Utilidades de scraping
# =========================

def _clean_one_line(s: str) -> str:
    s = (s or "").replace("\xa0", " ").strip()
    return re.sub(r"\s+", " ", s)

def only_digits(s: str) -> str:
    return "".join(ch for ch in (s or "") if ch.isdigit())

def td_value(page: Page, label: str, keep_newlines: bool = False, nth: int = 0) -> str:
    """
    Devuelve el texto del <td> siguiente al <td> cuyo contenido CONTIENE 'label'.
    - Usa normalize-space() para ser más robusto.
    - Si 'nth' > 0, devuelve la enésima coincidencia (0-based).
    """
    loc = page.locator(
        f"xpath=(//td[contains(normalize-space(.), '{label}')]/following-sibling::td[1])[{nth+1}]"
    )
    try:
        if loc and loc.count():
            txt = loc.inner_text(timeout=10_000)
            if keep_newlines:
                txt = txt.replace("\r", "\n")
                lines = [ln.strip() for ln in txt.split("\n") if ln.strip()]
                return "\n".join(lines)
            return _clean_one_line(txt)
    except Exception:
        pass
    return ""

def td_values(page: Page, label: str, keep_newlines: bool = False) -> List[str]:
    """
    Devuelve TODAS las coincidencias del <td> siguiente al que contiene label.
    """
    loc = page.locator(
        f"xpath=//td[contains(normalize-space(.), '{label}')]/following-sibling::td[1]"
    )
    out = []
    try:
        n = loc.count()
        for i in range(n):
            txt = loc.nth(i).inner_text(timeout=10_000)
            if keep_newlines:
                txt = txt.replace("\r", "\n")
                lines = [ln.strip() for ln in txt.split("\n") if ln.strip()]
                out.append("\n".join(lines))
            else:
                out.append(_clean_one_line(txt))
    except Exception:
        pass
    return out

def td_value_any(page: Page, labels: List[str], keep_newlines: bool = False) -> str:
    for lb in labels:
        v = td_value(page, lb, keep_newlines=keep_newlines)
        if v:
            return v
    return ""

def split_domicilio(block_text: str) -> Tuple[str, str, str]:
    """Devuelve (domicilio, localidad, provincia) a partir del bloque multilínea del sitio."""
    if not block_text:
        return "", "", ""
    parts = [p.strip() for p in block_text.replace("\r", "\n").split("\n") if p.strip()]
    dom = parts[0] if len(parts) > 0 else ""
    loc = parts[1] if len(parts) > 1 else ""
    prov = parts[2] if len(parts) > 2 else ""
    return dom, loc, prov

# =========================
# Login + navegación a la OT
# =========================

def login_y_abrir_ot(context: BrowserContext, usuario: str, password: str, ot: str) -> Tuple[Page, Dict[str, str], List[str]]:
    """
    Inicia sesión, navega a la OT y devuelve:
      - página ya dentro del VPE
      - meta {ot, vpe, empresa_solicitante, usuario_representado}
      - lista de hrefs a cada instrumento
    """
    page = context.new_page()
    page.set_default_timeout(60_000)

    print_info("Conectando con MetroWeb.")
    # Login
    page.goto(f"{BASE}/MetroWeb/pages/ingreso.jsp")

    # Encontrar usuario
    if page.locator('input[name="usuario"]').count():
        page.fill('input[name="usuario"]', usuario)
    elif page.locator('input[id="usuario"]').count():
        page.fill('input[id="usuario"]', usuario)
    else:
        page.fill('xpath=(//input[@type="text"])[1]', usuario)

    # Password
    if page.locator('input[name="contrasena"]').count():
        page.fill('input[name="contrasena"]', password)
    elif page.locator('input[name="password"]').count():
        page.fill('input[name="password"]', password)
    else:
        page.fill('xpath=(//input[@type="password"])[1]', password)

    if page.locator('input[value="Ingresar"]').count():
        page.click('input[value="Ingresar"]')
    elif page.locator('input[type="submit"]').count():
        page.click('input[type="submit"]')
    else:
        page.keyboard.press("Enter")

    print_info("Autenticando credenciales.")
    page.wait_for_load_state("networkidle")
    print_success("Sesión iniciada correctamente")

    # Buscar OT
    print_info(f"Buscando OT: {ot}")
    page.goto(f"{BASE}/MetroWeb/entrarPML.do")

    if page.locator('input[name="numeroOT"]').count():
        page.fill('input[name="numeroOT"]', ot)
    elif page.locator('input[name="nroOT"]').count():
        page.fill('input[name="nroOT"]', ot)
    else:
        caja = page.locator("xpath=//*[contains(normalize-space(.),'Número OT') or contains(normalize-space(.),'Nmero OT')]/following::input[1]")
        caja.fill(ot)

    if page.locator('input[value="Buscar"]').count():
        page.click('input[value="Buscar"]')
    else:
        page.keyboard.press("Enter")
    page.wait_for_load_state("networkidle")

    # Abrir primer trámite VPE
    link_vpe = page.locator('a[href*="tramiteVPE"]').first
    if not link_vpe or not link_vpe.count():
        print_error("No se encontró enlace de trámite VPE para esa OT")
        return page, {"ot": ot, "vpe": "", "empresa_solicitante": "", "usuario_representado": ""}, []

    vpe_text = _clean_one_line(link_vpe.inner_text())
    vpe_num = only_digits(vpe_text)
    print_success(f"VPE encontrado: {vpe_num}")
    link_vpe.click()
    page.wait_for_load_state("networkidle")

    # Ir a resumen.jsp
    print_info("Accediendo a datos del trámite.")
    page.goto(f"{BASE}/MetroWeb/pages/tramiteVPE/resumen.jsp")
    page.wait_for_load_state("networkidle")
    time.sleep(0.5)

    meta = leer_resumen(page)
    if not meta.get("ot"):
        meta["ot"] = ot
    if not meta.get("vpe"):
        meta["vpe"] = vpe_num

    # Enlaces a instrumentos
    instrument_links = []
    enlaces = page.locator('a[href*="instrumentoDetalle.do"]')
    for i in range(enlaces.count()):
        href = enlaces.nth(i).get_attribute("href") or ""
        if "instrumentoDetalle.do" in href:
            if not href.startswith("http"):
                href = BASE + href
            instrument_links.append(href)

    return page, meta, instrument_links

# =========================
# Lecturas de páginas
# =========================

def leer_resumen(page: Page) -> Dict[str, str]:
    """
    Lee campos desde resumen.jsp:
      - Nro OT
      - Número: vpeXXXXX → solo dígitos
      - Empresa Solicitante
      - Usuario Representado
    """
    meta = {
        "ot": "",
        "vpe": "",
        "empresa_solicitante": "",
        "usuario_representado": ""
    }

    # OT
    ot_val = td_value(page, "Nro OT") or td_value(page, "N° OT") or td_value(page, "Número de O.T.") or ""
    meta["ot"] = _clean_one_line(ot_val)

    # VPE
    try:
        html = page.content()
        m = re.search(r"vpe\s*0*?(\d+)", html, re.IGNORECASE)
        if m:
            meta["vpe"] = m.group(1)
    except Exception:
        pass
    if not meta["vpe"]:
        vpe_inline = td_value(page, "Número:") or ""
        meta["vpe"] = only_digits(vpe_inline)

    # Empresa Solicitante
    meta["empresa_solicitante"] = td_value(page, "Empresa Solicitante")

    # Usuario Representado
    meta["usuario_representado"] = td_value(page, "Usuario Representado")

    return meta

def leer_detalle_vpe(context: BrowserContext) -> Dict[str, str]:
    """
    Abre pages/tramiteVPE/detalle.jsp (misma sesión) y extrae:
      - 'Nombre del Usuario del Instrumento'  → razón social (propietario)
      - 'Dirección Legal'                     → domicilio fiscal
    """
    datos = {"nombre_usuario_instr": "", "direccion_legal": ""}

    page = context.new_page()
    page.set_default_timeout(60_000)
    try:
        page.goto(f"{BASE}/MetroWeb/pages/tramiteVPE/detalle.jsp")
        page.wait_for_load_state("networkidle")
        time.sleep(0.3)

        datos["nombre_usuario_instr"] = td_value_any(
            page,
            [
                "Nombre del Usuario del Instrumento",
                "Nombre del Usuario del instrumento",
                "Nombre del Usuario del equipo",
                "Nombre del Usuario",
            ],
        )

        datos["direccion_legal"] = td_value_any(
            page,
            [
                "Dirección Legal",
                "Dirección legal",
                "Direccion Legal",
                "Direccion legal",
            ],
            keep_newlines=False,
        )

    finally:
        try:
            page.close()
        except Exception:
            pass

    return datos

def leer_modelo_detalle(context: BrowserContext, href: str) -> Dict[str, str]:
    """
    Abre modeloDetalle.do y devuelve datos generales (+ características metrológicas si existen).
    """
    datos = {
        "modelo": "",
        "marca": "",
        "fabricante": "",
        "origen": "",
        "n_aprob": "",
        "fecha_aprob": "",
        "tipo_instr": "",
        "max": "",
        "min": "",
        "e": "",
        "dd_dt": "",
        "clase": "",
        "codigo_aprobacion": ""
    }

    if not href:
        return datos

    page = context.new_page()
    page.set_default_timeout(60_000)

    try:
        page.goto(href)
        page.wait_for_load_state("networkidle")
        time.sleep(0.3)

        # Datos generales
        datos["modelo"] = td_value_any(page, ["Modelo Aprobado", "Modelo"])
        datos["fabricante"] = td_value_any(page, ["Fabricante/Importador", "Fabricante", "Importador"])
        datos["marca"] = td_value(page, "Marca")
        datos["origen"] = td_value_any(page, ["País Origen", "País de Origen", "País  Origen", "Origen"])
        datos["n_aprob"] = td_value_any(page, [
            "Nº Disposicion", "N° Disposicion",
            "Nº Disposición", "N° Disposición",
            "Nº Disposici", "N° Disposici",
            "N° de Aprobación", "Nº de Aprobación"
        ])
        datos["fecha_aprob"] = td_value_any(page, ["Fecha Aprobación", "Fecha de Aprobación"])
        datos["tipo_instr"] = td_value_any(page, ["Tipo Instrumento", "Tipo de Instrumento"])

        # Características metrológicas
        datos["max"]   = td_value_any(page, ["Máximo", "Capacidad Máx.", "Capacidad máxima"])
        datos["min"]   = td_value_any(page, ["Mínimo", "Capacidad Mín.", "Capacidad mínima"])
        datos["e"]     = td_value(page, "e")
        datos["dd_dt"] = td_value_any(page, ["dd=dt", "dt", "dd", "d"])
        datos["clase"] = td_value(page, "Clase") or "III"

        # Posible código en el modelo
        datos["codigo_aprobacion"] = td_value_any(page, ["Código Aprobación", "Codigo Aprobación", "Codigo Aprobacion"])

    finally:
        try:
            page.close()
        except Exception:
            pass

    return datos

def leer_instrumento(context: BrowserContext, id_instrumento: str) -> Dict[str, any]:
    """
    Abre instrumentoDetalle.do?idInstrumento=. y extrae:
      - Domicilio (para lugar de instalación)
      - Receptor: href modelo, código de aprobación, nro de serie
      - Indicador: href modelo, código de aprobación, nro de serie
    """
    url = f"{BASE}/MetroWeb/instrumentoDetalle.do?idInstrumento={id_instrumento}"
    page = context.new_page()
    page.set_default_timeout(60_000)

    data = {
        "inst_dom": "",
        "inst_loc": "",
        "inst_prov": "",
        "receptor": {"href": "", "code": "", "serie": ""},
        "indicador": {"href": "", "code": "", "serie": ""}
    }

    try:
        page.goto(url)
        page.wait_for_load_state("networkidle")
        time.sleep(0.3)

        # Ubicación
        dom_block = td_value(page, "Domicilio", keep_newlines=True)
        dom, loc, prov = split_domicilio(dom_block)
        data["inst_dom"], data["inst_loc"], data["inst_prov"] = dom, loc, prov

        # Links a modelos
        links = page.locator("a[href*='modeloDetalle.do']")
        hrefs = []
        for i in range(links.count()):
            h = links.nth(i).get_attribute("href") or ""
            if "modeloDetalle.do" in h:
                hrefs.append(h if h.startswith("http") else BASE + h)

        # Códigos de aprobación (texto en página)
        codes = td_values(page, "Código de Aprobación de Modelo") or td_values(page, "Código de Aprobación")

        # Series
        series = td_values(page, "Nro de serie")

        # Mapear: 0=receptor, 1=indicador
        if len(hrefs) >= 1:
            data["receptor"]["href"] = hrefs[0]
        if len(hrefs) >= 2:
            data["indicador"]["href"] = hrefs[1]

        if len(codes) >= 1:
            data["receptor"]["code"] = codes[0]
        if len(codes) >= 2:
            data["indicador"]["code"] = codes[1]

        if len(series) >= 1:
            data["receptor"]["serie"] = series[0]
        if len(series) >= 2:
            data["indicador"]["serie"] = series[1]

    finally:
        try:
            page.close()
        except Exception:
            pass

    return data

# =========================
# Extracción principal
# =========================

def extraer_camiones_por_ot(ot: str, user: str, pwd: str, mostrar_navegador: bool = True) -> List[Dict[str, str]]:
    """
    Extrae todos los instrumentos de la OT (camiones/plataforma) y devuelve una lista de filas (dict)
    con las 33 columnas solicitadas.
    """
    filas: List[Dict[str, str]] = []

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=not mostrar_navegador, slow_mo=0)
        context = browser.new_context()
        try:
            # 1) Login + entrar a VPE + resumen
            page, meta, instrument_links = login_y_abrir_ot(context, user, pwd, ot)

            # 1.b) detalle.jsp
            print_info("Extrayendo datos del propietario.")
            det = leer_detalle_vpe(context)
            nombre_usuario_det = det.get("nombre_usuario_instr", "").strip()
            direccion_legal_det = det.get("direccion_legal", "").strip()

            nro_ot = meta.get("ot", "").strip()
            vpe_num = meta.get("vpe", "").strip()
            empresa = meta.get("empresa_solicitante", "").strip()
            usuario_rep = nombre_usuario_det or meta.get("usuario_representado", "").strip()

            print_section("INFORMACIÓN DEL TRÁMITE")
            print_data_table({
                "Número de O.T.": nro_ot,
                "VPE Nº": vpe_num,
                "Empresa Solicitante": empresa,
                "Razón Social (Propietario)": usuario_rep,
                "Domicilio Fiscal": direccion_legal_det or "(no disponible)"
            })

            if not instrument_links:
                print_warning("No se detectaron instrumentos en el VPE")
                return []

            print_section(f"PROCESANDO INSTRUMENTOS ({len(instrument_links)} encontrados)")

            # 2) Recorrer instrumentos
            for idx, href in enumerate(instrument_links, start=1):
                print_progress(idx, len(instrument_links), f"Instrumento {idx}")

                # idInstrumento
                m = re.search(r"idInstrumento=(\d+)", href)
                id_instrumento = m.group(1) if m else ""

                inst = leer_instrumento(context, id_instrumento) if id_instrumento else {
                    "inst_dom": "", "inst_loc": "", "inst_prov": "",
                    "receptor": {"href": "", "code": "", "serie": ""},
                    "indicador": {"href": "", "code": "", "serie": ""},
                }

                # Modelos
                rec = inst["receptor"]
                ind = inst["indicador"]

                rec_model = leer_modelo_detalle(context, rec.get("href", "")) if rec.get("href") else {}
                ind_model = leer_modelo_detalle(context, ind.get("href", "")) if ind.get("href") else {}

                # Resolver campos receptor
                fab_rec   = (rec_model.get("fabricante") or "").strip()
                marca_rec = (rec_model.get("marca") or "").strip()
                modelo_rec= (rec_model.get("modelo") or "").strip()
                serie_rec = (rec.get("serie") or "").strip()
                codap_rec = (rec_model.get("codigo_aprobacion") or rec.get("code") or "").strip()
                origen_rec= (rec_model.get("origen") or "").strip()
                e_rec     = (rec_model.get("e") or "").strip()
                max_rec   = (rec_model.get("max") or "").strip()
                min_rec   = (rec_model.get("min") or "").strip()
                dd_dt_rec = (rec_model.get("dd_dt") or "").strip()
                clase_rec = (rec_model.get("clase") or "").strip()
                naprob_rec= (rec_model.get("n_aprob") or "").strip()
                faprob_rec= (rec_model.get("fecha_aprob") or "").strip()

                # Indicador
                fab_ind   = (ind_model.get("fabricante") or "").strip()
                marca_ind = (ind_model.get("marca") or "").strip()
                modelo_ind= (ind_model.get("modelo") or "").strip()
                serie_ind = (ind.get("serie") or "").strip()
                codap_ind = (ind_model.get("codigo_aprobacion") or ind.get("code") or "").strip()
                origen_ind= (ind_model.get("origen") or "").strip()
                naprob_ind= (ind_model.get("n_aprob") or "").strip()
                faprob_ind= (ind_model.get("fecha_aprob") or "").strip()

                # Ubicación
                lugar_dom  = inst.get("inst_dom", "")
                lugar_loc  = inst.get("inst_loc", "")
                lugar_prov = inst.get("inst_prov", "")

                fila = {
                    "Número de O.T.": nro_ot,
                    "VPE Nº": vpe_num,
                    "Empresa solicitante": empresa,
                    "Razón social (Propietario)": usuario_rep,
                    "Domicilio (Fiscal)": direccion_legal_det,
                    "Localidad (Fiscal)": "",
                    "Provincia (Fiscal)": "",
                    "Lugar propio de instalación - Domicilio": lugar_dom,
                    "Lugar propio de instalación - Localidad": lugar_loc,
                    "Lugar propio de instalación - Provincia": lugar_prov,
                    "Instrumento verificado": "Balanza para pesar camiones",

                    # Receptor
                    "Fabricante receptor": fab_rec,
                    "Marca Receptor": marca_rec,
                    "Modelo Receptor": modelo_rec,
                    "N° de serie Receptor": serie_rec,
                    "Cód ap. mod. Receptor": codap_rec,
                    "Origen Receptor": origen_rec,
                    "e": e_rec,
                    "máx": max_rec,
                    "mín": min_rec,
                    "dd=dt": dd_dt_rec,
                    "clase": clase_rec,
                    "N° de Aprobación Modelo (Receptor)": naprob_rec,
                    "Fecha de Aprobación Modelo (Receptor)": faprob_rec,

                    # Indicador
                    "Tipo (Indicador)": "electrónica",
                    "Fabricante Indicador": fab_ind,
                    "Marca Indicador": marca_ind,
                    "Modelo Indicador": modelo_ind,
                    "N° de serie Indicador": serie_ind,
                    "Código Aprobación (Indicador)": codap_ind,
                    "Origen Indicador": origen_ind,
                    "N° de Aprobación Modelo (Indicador)": naprob_ind,
                    "Fecha de Aprobación Modelo (Indicador)": faprob_ind,
                }

                filas.append(fila)

            print()  # Nueva línea después de la barra de progreso
            print_success(f"Se procesaron {len(filas)} instrumento(s) correctamente")

        finally:
            try:
                browser.close()
            except Exception:
                pass

    return filas

# =========================
# Excel
# =========================

COLUMNS_ORDER = [
    "Número de O.T.",
    "VPE Nº",
    "Empresa solicitante",
    "Razón social (Propietario)",
    "Domicilio (Fiscal)",
    "Localidad (Fiscal)",
    "Provincia (Fiscal)",
    "Lugar propio de instalación - Domicilio",
    "Lugar propio de instalación - Localidad",
    "Lugar propio de instalación - Provincia",
    "Instrumento verificado",

    # Receptor
    "Fabricante receptor",
    "Marca Receptor",
    "Modelo Receptor",
    "N° de serie Receptor",
    "Cód ap. mod. Receptor",
    "Origen Receptor",
    "e",
    "máx",
    "mín",
    "dd=dt",
    "clase",
    "N° de Aprobación Modelo (Receptor)",
    "Fecha de Aprobación Modelo (Receptor)",

    # Indicador
    "Tipo (Indicador)",
    "Fabricante Indicador",
    "Marca Indicador",
    "Modelo Indicador",
    "N° de serie Indicador",
    "Código Aprobación (Indicador)",
    "Origen Indicador",
    "N° de Aprobación Modelo (Indicador)",
    "Fecha de Aprobación Modelo (Indicador)"
]

def armar_hoja_verificacion(filas: List[Dict[str, str]]) -> pd.DataFrame:
    df = pd.DataFrame(filas)
    # Asegurar columnas, en orden exacto
    for col in COLUMNS_ORDER:
        if col not in df.columns:
            df[col] = ""
    df = df[COLUMNS_ORDER]
    return df

def exportar_verificacion(df: pd.DataFrame, ruta: Path) -> Path:
    print_section("GENERANDO ARCHIVO EXCEL")
    print_info(f"Creando archivo: {ruta}")

    ruta = Path(ruta)
    ruta.parent.mkdir(parents=True, exist_ok=True)

    with pd.ExcelWriter(ruta, engine="xlsxwriter") as writer:
        df.to_excel(writer, sheet_name="Verificación", index=False)
        wb = writer.book
        ws = writer.sheets["Verificación"]

        # Formatos
        fmt_header = wb.add_format({
            "bold": True,
            "bg_color": "#4472C4",
            "font_color": "white",
            "border": 1,
            "align": "center",
            "valign": "vcenter"
        })

        fmt_data = wb.add_format({
            "border": 1,
            "valign": "top"
        })

        # Cabeceras
        for i, col in enumerate(df.columns):
            ws.write(0, i, col, fmt_header)

        # Datos con formato
        for row_num in range(1, len(df) + 1):
            for col_num in range(len(df.columns)):
                ws.write(row_num, col_num, df.iloc[row_num - 1, col_num], fmt_data)

        # Anchos de columna
        for i, col in enumerate(df.columns):
            try:
                max_len = int(df[col].astype(str).map(len).max() if not df.empty else 10)
            except Exception:
                max_len = 10
            max_len = max(15, min(50, max_len + 2))
            ws.set_column(i, i, max_len)

        ws.freeze_panes(1, 0)
        ws.set_row(0, 30)  # Altura de cabecera

    return ruta

# =========================
# CLI
# =========================

def solicitar_ruta_salida(ot: str, num_instrumentos: int = 0) -> str:
    """
    Solicita al usuario la ruta y nombre del archivo de salida.
    """
    print_section("CONFIGURACIÓN DE ARCHIVO DE SALIDA")

    if num_instrumentos > 0:
        print_success(f"Se extrajeron {num_instrumentos} instrumento(s) exitosamente")
        print()

    nombre_sugerido = f"OT_{ot.replace('-', '_')}_VERIFICACION_PREVIA.xlsx"

    print_info(f"Nombre sugerido: {Fore.WHITE}{nombre_sugerido}{Style.RESET_ALL}")
    print()

    # Preguntar si usar nombre sugerido
    usar_sugerido = input(f"{Fore.CYAN}¿Desea usar el nombre sugerido? (s/n): {Style.RESET_ALL}").strip().lower()
    if usar_sugerido in ['s', 'si', 'sí', 'y', 'yes', '']:
        ruta_completa = Path.cwd() / nombre_sugerido
    else:
        # Solicitar nombre
        nombre_archivo = input(f"{Fore.CYAN}Nombre del archivo .xlsx (ENTER para usar sugerido): {Style.RESET_ALL}").strip() or nombre_sugerido

        # Asegurar extensión .xlsx
        if not nombre_archivo.endswith('.xlsx'):
            nombre_archivo += '.xlsx'

        # Elegir carpeta
        print()
        print_info("Seleccione una carpeta de destino (o ENTER para usar la carpeta actual).")
        print_info(f"Carpeta actual: {Fore.WHITE}{Path.cwd()}{Style.RESET_ALL}")
        carpeta_txt = input(f"{Fore.CYAN}Ruta de carpeta (ENTER=actual): {Style.RESET_ALL}").strip()
        if not carpeta_txt:
            ruta_completa = Path.cwd() / nombre_archivo
        else:
            try:
                carpeta = Path(carpeta_txt).expanduser()
                if not carpeta.exists():
                    print_warning("La carpeta no existe.")
                    crear = input(f"{Fore.YELLOW}¿Desea crearla? (s/n): {Style.RESET_ALL}").strip().lower()
                    if crear in ['s', 'si', 'sí', 'y', 'yes', '']:
                        carpeta.mkdir(parents=True, exist_ok=True)
                        print_success(f"Carpeta creada: {carpeta}")
                    else:
                        print_warning("Usando carpeta actual")
                        carpeta = Path.cwd()
                ruta_completa = carpeta / nombre_archivo

            except Exception as e:
                print_error(f"Ruta inválida: {e}")
                print_warning("Usando carpeta actual")
                ruta_completa = Path.cwd() / nombre_archivo

    print()
    # Verificar si el archivo ya existe
    if ruta_completa.exists():
        print_warning(f"El archivo ya existe: {ruta_completa}")
        sobrescribir = input(f"{Fore.YELLOW}¿Desea sobrescribirlo? (s/n): {Style.RESET_ALL}").strip().lower()
        if sobrescribir not in ['s', 'si', 'sí', 'y', 'yes', '']:
            # Generar nombre alternativo con timestamp
            timestamp = time.strftime("%Y%m%d_%H%M%S")
            nombre_base = ruta_completa.stem
            ruta_completa = ruta_completa.parent / f"{nombre_base}_{timestamp}.xlsx"
            print_info(f"Se usará un nombre alternativo: {ruta_completa.name}")

    print()
    print_success(f"Archivo se guardará en:")
    print(f"  {Fore.WHITE}{Style.BRIGHT}{ruta_completa.resolve()}{Style.RESET_ALL}")
    print()

    return str(ruta_completa)

def main():
    print_header()

    parser = argparse.ArgumentParser(
        description="Extracción MetroWeb → Excel Verificación Previa (balanzas de camiones/plataforma).",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=f"""
{Fore.CYAN}Ejemplos de uso:{Style.RESET_ALL}
  python extract_camiones.py --user miusuario --pass mipass --ot 307-62136
  python extract_camiones.py --user miusuario --pass mipass --ot 307-62136 --headless
  python extract_camiones.py --ot 307-62136 --out /ruta/resultado.xlsx
        """
    )
    parser.add_argument("--user", dest="user", help="Usuario MetroWeb")
    parser.add_argument("--pass", dest="pwd", help="Contraseña MetroWeb")
    parser.add_argument("--ot", dest="ot", help="Número de OT (ej. 307-62136)")
    parser.add_argument("--headless", action="store_true", help="Ejecutar sin mostrar navegador")
    parser.add_argument("--out", dest="out", help="Ruta de salida .xlsx (opcional, si no se especifica se preguntará)")

    args = parser.parse_args()

    # Solicitar datos faltantes interactivamente
    print_section("CONFIGURACIÓN DE ACCESO")

    user = args.user
    if not user:
        user = input(f"{Fore.CYAN}Usuario MetroWeb: {Style.RESET_ALL}").strip()
    else:
        print_info(f"Usuario: {user}")

    pwd = args.pwd
    if not pwd:
        pwd = getpass.getpass(f"{Fore.CYAN}Contraseña MetroWeb: {Style.RESET_ALL}").strip()
    else:
        print_info("Contraseña: ********")

    ot = args.ot
    if not ot:
        ot = input(f"{Fore.CYAN}Número de OT (ej. 307-62136): {Style.RESET_ALL}").strip()
    else:
        print_info(f"Número de OT: {ot}")

    headless = args.headless

    print()

    # Validación de formato OT
    if not re.match(r"^\d{3}-\d{5}$", ot):
        print_warning(f"Formato de OT no estándar. Se esperaba: XXX-XXXXX (ej. 307-62136)")
        continuar = input(f"{Fore.YELLOW}¿Desea continuar de todas formas? (s/n): {Style.RESET_ALL}").lower()
        if continuar not in ['s', 'si', 'sí', 'y', 'yes']:
            print_error("Operación cancelada por el usuario")
            sys.exit(1)
        print()

    # Solicitar ruta de salida (si no se pasó por argumento)
    print_info(f"Modo navegador: {'headless (oculto)' if headless else 'visible'}")
    print()

    # Confirmación antes de continuar
    input(f"{Fore.GREEN}{Style.BRIGHT}Presione ENTER para iniciar la extracción.{Style.RESET_ALL}")
    print()

    # Inicio del proceso
    print_section("INICIANDO EXTRACCIÓN")
    start_time = time.time()

    try:
        filas = extraer_camiones_por_ot(ot=ot, user=user, pwd=pwd, mostrar_navegador=not headless)

        if not filas:
            print_warning("No se generaron filas. Verifique la OT o las credenciales")
            print_info("Se creará un archivo Excel vacío con la estructura correcta")

        # Solicitar ruta de salida DESPUÉS de la extracción
        if args.out:
            out = args.out
            print_section("CONFIGURACIÓN DE ARCHIVO DE SALIDA")
            print_info(f"Archivo de salida (especificado por argumento): {out}")
            print()
        else:
            print()
            out = solicitar_ruta_salida(ot, len(filas))

        df = armar_hoja_verificacion(filas)
        ruta = exportar_verificacion(df, Path(out))

        elapsed_time = time.time() - start_time

        # Resumen final
        print()
        print("=" * 70)
        print(f"{Fore.GREEN}{Style.BRIGHT}{'✓ EXTRACCIÓN COMPLETADA EXITOSAMENTE':^70}{Style.RESET_ALL}")
        print("=" * 70)
        print()
        print_data_table({
            "Archivo generado": str(ruta.resolve()),
            "Instrumentos procesados": str(len(filas)),
            "Tiempo de ejecución": f"{elapsed_time:.2f} segundos",
            "Tamaño del archivo": f"{ruta.stat().st_size / 1024:.2f} KB"
        }, "Resumen de la operación")
        print()
        print(f"{Fore.CYAN}💡 Tip: Puede abrir el archivo con Excel, LibreOffice o Google Sheets{Style.RESET_ALL}")
        print(f"{Fore.WHITE}   Ubicación: {Style.BRIGHT}{ruta.parent}{Style.RESET_ALL}")
        print()

    except KeyboardInterrupt:
        print()
        print_error("Operación cancelada por el usuario (Ctrl+C)")
        sys.exit(1)
    except Exception as e:
        print()
        print_error(f"Error durante la extracción: {str(e)}")
        print()
        print(f"{Fore.RED}Detalles del error:{Style.RESET_ALL}")
        import traceback
        traceback.print_exc()
        sys.exit(1)

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print()
        print_error(f"Error fatal: {str(e)}")
        sys.exit(1)
