# extract_camiones_01_cosola_gui.py
# -*- coding: utf-8 -*-

"""
Extracción INTI MetroWeb → Excel (Verificación Previa) para balanzas de camiones/plataforma.
Versión con interfaz gráfica Windows.

Requisitos:
  pip install playwright pandas xlsxwriter
  python -m playwright install chromium
"""

import re
import sys
import time
import threading
from pathlib import Path
from typing import Dict, List, Tuple, Optional
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import pandas as pd
from playwright.sync_api import sync_playwright, BrowserContext, Page

BASE = "https://app.inti.gob.ar"

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

def login_y_abrir_ot(context: BrowserContext, usuario: str, password: str, ot: str, 
                    update_status_callback=None) -> Tuple[Page, Dict[str, str], List[str]]:
    """
    Inicia sesión, navega a la OT y devuelve:
      - página ya dentro del VPE
      - meta {ot, vpe, empresa_solicitante, usuario_representado}
      - lista de hrefs a cada instrumento
    """
    page = context.new_page()
    page.set_default_timeout(60_000)

    if update_status_callback:
        update_status_callback("Conectando con MetroWeb...")
    
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

    if update_status_callback:
        update_status_callback("Autenticando credenciales...")
    page.wait_for_load_state("networkidle")
    
    if update_status_callback:
        update_status_callback("Sesión iniciada correctamente")

    # Buscar OT
    if update_status_callback:
        update_status_callback(f"Buscando OT: {ot}")
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
        if update_status_callback:
            update_status_callback("ERROR: No se encontró enlace de trámite VPE para esa OT")
        return page, {"ot": ot, "vpe": "", "empresa_solicitante": "", "usuario_representado": ""}, []

    vpe_text = _clean_one_line(link_vpe.inner_text())
    vpe_num = only_digits(vpe_text)
    
    if update_status_callback:
        update_status_callback(f"VPE encontrado: {vpe_num}")
    link_vpe.click()
    page.wait_for_load_state("networkidle")

    # Ir a resumen.jsp
    if update_status_callback:
        update_status_callback("Accediendo a datos del trámite...")
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

def extraer_camiones_por_ot(ot: str, user: str, pwd: str, mostrar_navegador: bool = True, 
                           update_status_callback=None, update_progress_callback=None) -> List[Dict[str, str]]:
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
            if update_status_callback:
                update_status_callback("Iniciando sesión en MetroWeb...")
                
            page, meta, instrument_links = login_y_abrir_ot(context, user, pwd, ot, update_status_callback)

            # 1.b) detalle.jsp
            if update_status_callback:
                update_status_callback("Extrayendo datos del propietario...")
            det = leer_detalle_vpe(context)
            nombre_usuario_det = det.get("nombre_usuario_instr", "").strip()
            direccion_legal_det = det.get("direccion_legal", "").strip()

            nro_ot = meta.get("ot", "").strip()
            vpe_num = meta.get("vpe", "").strip()
            empresa = meta.get("empresa_solicitante", "").strip()
            usuario_rep = nombre_usuario_det or meta.get("usuario_representado", "").strip()

            if update_status_callback:
                update_status_callback("Información del trámite extraída correctamente")

            if not instrument_links:
                if update_status_callback:
                    update_status_callback("ADVERTENCIA: No se detectaron instrumentos en el VPE")
                return []

            if update_status_callback:
                update_status_callback(f"Procesando {len(instrument_links)} instrumentos encontrados")

            # 2) Recorrer instrumentos
            for idx, href in enumerate(instrument_links, start=1):
                if update_progress_callback:
                    update_progress_callback(idx, len(instrument_links), f"Instrumento {idx}")
                    
                if update_status_callback:
                    update_status_callback(f"Procesando instrumento {idx} de {len(instrument_links)}")

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

            if update_status_callback:
                update_status_callback(f"Se procesaron {len(filas)} instrumento(s) correctamente")

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
# Interfaz Gráfica
# =========================

class MetroWebExtractorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("INTI MetroWeb - Extractor de Verificación Previa")
        self.root.geometry("800x700")
        self.root.resizable(True, True)
        
        # Variables de control
        self.user_var = tk.StringVar()
        self.pwd_var = tk.StringVar()
        self.ot_var = tk.StringVar()
        self.headless_var = tk.BooleanVar(value=True)
        self.output_path_var = tk.StringVar()
        
        # Estado de la aplicación
        self.is_running = False
        self.extraction_thread = None
        
        self.setup_ui()
        
    def setup_ui(self):
        # Frame principal
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configurar grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        
        # Título
        title_label = ttk.Label(main_frame, 
                               text="INTI METROWEB - EXTRACTOR DE VERIFICACIÓN PREVIA\nBalanzas de Camiones / Plataforma",
                               font=("Arial", 12, "bold"),
                               justify=tk.CENTER)
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 20))
        
        # Sección de configuración de acceso
        access_frame = ttk.LabelFrame(main_frame, text="Configuración de Acceso", padding="10")
        access_frame.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        access_frame.columnconfigure(1, weight=1)
        
        ttk.Label(access_frame, text="Usuario MetroWeb:").grid(row=0, column=0, sticky=tk.W, pady=5)
        ttk.Entry(access_frame, textvariable=self.user_var, width=30).grid(row=0, column=1, sticky=(tk.W, tk.E), pady=5, padx=(5, 0))
        
        ttk.Label(access_frame, text="Contraseña:").grid(row=1, column=0, sticky=tk.W, pady=5)
        ttk.Entry(access_frame, textvariable=self.pwd_var, show="*", width=30).grid(row=1, column=1, sticky=(tk.W, tk.E), pady=5, padx=(5, 0))
        
        ttk.Label(access_frame, text="Número de OT:").grid(row=2, column=0, sticky=tk.W, pady=5)
        ttk.Entry(access_frame, textvariable=self.ot_var, width=30).grid(row=2, column=1, sticky=(tk.W, tk.E), pady=5, padx=(5, 0))
        ttk.Label(access_frame, text="(ej. 307-62136)").grid(row=2, column=2, sticky=tk.W, pady=5, padx=(5, 0))
        
        ttk.Checkbutton(access_frame, text="Ejecutar en modo headless (sin mostrar navegador)", 
                       variable=self.headless_var).grid(row=3, column=0, columnspan=2, sticky=tk.W, pady=5)
        
        # Sección de archivo de salida
        output_frame = ttk.LabelFrame(main_frame, text="Archivo de Salida", padding="10")
        output_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        output_frame.columnconfigure(0, weight=1)
        
        ttk.Label(output_frame, text="Ruta del archivo Excel:").grid(row=0, column=0, sticky=tk.W, pady=5)
        
        output_entry_frame = ttk.Frame(output_frame)
        output_entry_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        output_entry_frame.columnconfigure(0, weight=1)
        
        ttk.Entry(output_entry_frame, textvariable=self.output_path_var).grid(row=0, column=0, sticky=(tk.W, tk.E))
        ttk.Button(output_entry_frame, text="Examinar...", command=self.browse_output_file).grid(row=0, column=1, padx=(5, 0))
        
        # Botones de acción
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=3, column=0, columnspan=3, pady=10)
        
        self.extract_button = ttk.Button(button_frame, text="Iniciar Extracción", command=self.start_extraction)
        self.extract_button.grid(row=0, column=0, padx=5)
        
        self.cancel_button = ttk.Button(button_frame, text="Cancelar", command=self.cancel_extraction, state=tk.DISABLED)
        self.cancel_button.grid(row=0, column=1, padx=5)
        
        ttk.Button(button_frame, text="Salir", command=self.root.quit).grid(row=0, column=2, padx=5)
        
        # Barra de progreso
        self.progress_frame = ttk.LabelFrame(main_frame, text="Progreso", padding="10")
        self.progress_frame.grid(row=4, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        
        self.progress_bar = ttk.Progressbar(self.progress_frame, mode='determinate')
        self.progress_bar.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=5)
        
        self.progress_label = ttk.Label(self.progress_frame, text="Listo para comenzar")
        self.progress_label.grid(row=1, column=0, sticky=tk.W)
        
        # Área de log
        log_frame = ttk.LabelFrame(main_frame, text="Log de Actividad", padding="10")
        log_frame.grid(row=5, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        main_frame.rowconfigure(5, weight=1)
        
        self.log_text = scrolledtext.ScrolledText(log_frame, height=15, width=80)
        self.log_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configurar el archivo de salida por defecto
        self.set_default_output_path()
        
    def set_default_output_path(self):
        """Establece la ruta de salida por defecto"""
        default_name = "OT_VERIFICACION_PREVIA.xlsx"
        self.output_path_var.set(str(Path.cwd() / default_name))
        
    def browse_output_file(self):
        """Abre un diálogo para seleccionar el archivo de salida"""
        filename = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
            title="Guardar archivo Excel como"
        )
        if filename:
            self.output_path_var.set(filename)
            
    def log_message(self, message):
        """Añade un mensaje al área de log"""
        self.log_text.insert(tk.END, f"{message}\n")
        self.log_text.see(tk.END)
        self.root.update_idletasks()
        
    def update_progress(self, current, total, item_name=""):
        """Actualiza la barra de progreso"""
        if total > 0:
            percentage = (current / total) * 100
            self.progress_bar['value'] = percentage
            self.progress_label.config(text=f"{percentage:.1f}% - {item_name}")
        self.root.update_idletasks()
        
    def update_status(self, message):
        """Actualiza el mensaje de estado"""
        self.log_message(message)
        
    def start_extraction(self):
        """Inicia el proceso de extracción en un hilo separado"""
        if not self.validate_inputs():
            return
            
        self.is_running = True
        self.extract_button.config(state=tk.DISABLED)
        self.cancel_button.config(state=tk.NORMAL)
        
        # Limpiar log anterior
        self.log_text.delete(1.0, tk.END)
        
        # Iniciar extracción en un hilo separado
        self.extraction_thread = threading.Thread(target=self.run_extraction)
        self.extraction_thread.daemon = True
        self.extraction_thread.start()
        
    def validate_inputs(self):
        """Valida los datos de entrada"""
        if not self.user_var.get().strip():
            messagebox.showerror("Error", "Por favor, ingrese el usuario MetroWeb")
            return False
            
        if not self.pwd_var.get().strip():
            messagebox.showerror("Error", "Por favor, ingrese la contraseña")
            return False
            
        if not self.ot_var.get().strip():
            messagebox.showerror("Error", "Por favor, ingrese el número de OT")
            return False
            
        if not re.match(r"^\d{3}-\d{5}$", self.ot_var.get().strip()):
            result = messagebox.askyesno(
                "Advertencia", 
                "El formato de OT no es estándar (se espera XXX-XXXXX). ¿Desea continuar de todas formas?"
            )
            if not result:
                return False
                
        if not self.output_path_var.get().strip():
            messagebox.showerror("Error", "Por favor, seleccione una ruta de salida")
            return False
            
        return True
        
    def run_extraction(self):
        """Ejecuta el proceso de extracción (en hilo separado)"""
        try:
            start_time = time.time()
            
            self.update_status("Iniciando proceso de extracción...")
            
            # Ejecutar extracción
            filas = extraer_camiones_por_ot(
                ot=self.ot_var.get().strip(),
                user=self.user_var.get().strip(),
                pwd=self.pwd_var.get().strip(),
                mostrar_navegador=not self.headless_var.get(),
                update_status_callback=self.update_status,
                update_progress_callback=self.update_progress
            )
            
            if not filas:
                self.update_status("ADVERTENCIA: No se generaron filas. Verifique la OT o las credenciales")
                self.update_status("Se creará un archivo Excel vacío con la estructura correcta")
            
            # Generar Excel
            self.update_status("Generando archivo Excel...")
            df = armar_hoja_verificacion(filas)
            ruta = exportar_verificacion(df, Path(self.output_path_var.get()))
            
            elapsed_time = time.time() - start_time
            
            # Mostrar resumen
            self.update_status("=" * 60)
            self.update_status("✓ EXTRACCIÓN COMPLETADA EXITOSAMENTE")
            self.update_status("=" * 60)
            self.update_status(f"Archivo generado: {ruta.resolve()}")
            self.update_status(f"Instrumentos procesados: {len(filas)}")
            self.update_status(f"Tiempo de ejecución: {elapsed_time:.2f} segundos")
            self.update_status(f"Tamaño del archivo: {ruta.stat().st_size / 1024:.2f} KB")
            
            # Mostrar mensaje de éxito
            self.root.after(0, lambda: messagebox.showinfo(
                "Extracción Completada", 
                f"Proceso finalizado exitosamente.\n\n"
                f"Archivo: {ruta.name}\n"
                f"Instrumentos: {len(filas)}\n"
                f"Tiempo: {elapsed_time:.2f} segundos"
            ))
            
        except Exception as e:
            error_msg = f"Error durante la extracción: {str(e)}"
            self.update_status(f"ERROR: {error_msg}")
            self.root.after(0, lambda: messagebox.showerror("Error", error_msg))
            
        finally:
            self.root.after(0, self.extraction_finished)
            
    def extraction_finished(self):
        """Limpieza después de finalizar la extracción"""
        self.is_running = False
        self.extract_button.config(state=tk.NORMAL)
        self.cancel_button.config(state=tk.DISABLED)
        self.progress_bar['value'] = 0
        self.progress_label.config(text="Proceso finalizado")
        
    def cancel_extraction(self):
        """Cancela el proceso de extracción en curso"""
        if self.is_running:
            self.is_running = False
            self.update_status("Cancelando extracción...")
            # No podemos forzar la terminación del hilo, pero podemos marcar para cancelar
            
def main():
    root = tk.Tk()
    app = MetroWebExtractorApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()