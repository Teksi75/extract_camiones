# extract_camiones_gui.py
# -*- coding: utf-8 -*-
# type: ignore

"""
Extracción INTI MetroWeb → Excel (Verificación Previa) para balanzas de camiones/plataforma.
VERSIÓN CON INTERFAZ GRÁFICA (GUI)

Requisitos:
  pip install playwright pandas xlsxwriter
  python -m playwright install chromium

Ejecución:
  python extract_camiones_gui.py
"""

import re
import sys
import time
import threading
from pathlib import Path
from typing import Dict, List, Tuple, Optional
from datetime import datetime

import pandas as pd
from playwright.sync_api import sync_playwright, BrowserContext, Page

# GUI imports
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
from tkinter.font import Font

BASE = "https://app.inti.gob.ar"

# =========================
# Utilidades de scraping (sin cambios)
# =========================

def _clean_one_line(s: str) -> str:
    s = (s or "").replace("\xa0", " ").strip()
    return re.sub(r"\s+", " ", s)

def only_digits(s: str) -> str:
    return "".join(ch for ch in (s or "") if ch.isdigit())

def td_value(page: Page, label: str, keep_newlines: bool = False, nth: int = 0) -> str:
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
    if not block_text:
        return "", "", ""
    parts = [p.strip() for p in block_text.replace("\r", "\n").split("\n") if p.strip()]
    dom = parts[0] if len(parts) > 0 else ""
    loc = parts[1] if len(parts) > 1 else ""
    prov = parts[2] if len(parts) > 2 else ""
    return dom, loc, prov

# =========================
# Login + navegación
# =========================

def login_y_abrir_ot(context: BrowserContext, usuario: str, password: str, ot: str, log_callback=None) -> Tuple[Page, Dict[str, str], List[str]]:
    page = context.new_page()
    page.set_default_timeout(60_000)

    if log_callback:
        log_callback("🔗 Conectando con MetroWeb...")
    
    page.goto(f"{BASE}/MetroWeb/pages/ingreso.jsp")

    if page.locator('input[name="usuario"]').count():
        page.fill('input[name="usuario"]', usuario)
    elif page.locator('input[id="usuario"]').count():
        page.fill('input[id="usuario"]', usuario)
    else:
        page.fill('xpath=(//input[@type="text"])[1]', usuario)

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

    if log_callback:
        log_callback("🔐 Autenticando credenciales...")
    
    page.wait_for_load_state("networkidle")
    
    if log_callback:
        log_callback("✅ Sesión iniciada correctamente")

    if log_callback:
        log_callback(f"🔍 Buscando OT: {ot}")
    
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

    link_vpe = page.locator('a[href*="tramiteVPE"]').first
    if not link_vpe or not link_vpe.count():
        raise Exception("No se encontró enlace de trámite VPE para esa OT")

    vpe_text = _clean_one_line(link_vpe.inner_text())
    vpe_num = only_digits(vpe_text)
    
    if log_callback:
        log_callback(f"✅ VPE encontrado: {vpe_num}")
    
    link_vpe.click()
    page.wait_for_load_state("networkidle")

    if log_callback:
        log_callback("📄 Accediendo a datos del trámite...")
    
    page.goto(f"{BASE}/MetroWeb/pages/tramiteVPE/resumen.jsp")
    page.wait_for_load_state("networkidle")
    time.sleep(0.5)

    meta = leer_resumen(page)
    if not meta.get("ot"):
        meta["ot"] = ot
    if not meta.get("vpe"):
        meta["vpe"] = vpe_num

    instrument_links = []
    enlaces = page.locator('a[href*="instrumentoDetalle.do"]')
    for i in range(enlaces.count()):
        href = enlaces.nth(i).get_attribute("href") or ""
        if "instrumentoDetalle.do" in href:
            if not href.startswith("http"):
                href = BASE + href
            instrument_links.append(href)

    return page, meta, instrument_links

def leer_resumen(page: Page) -> Dict[str, str]:
    meta = {
        "ot": "",
        "vpe": "",
        "empresa_solicitante": "",
        "usuario_representado": ""
    }

    ot_val = td_value(page, "Nro OT") or td_value(page, "N° OT") or td_value(page, "Número de O.T.") or ""
    meta["ot"] = _clean_one_line(ot_val)

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

    meta["empresa_solicitante"] = td_value(page, "Empresa Solicitante")
    meta["usuario_representado"] = td_value(page, "Usuario Representado")

    return meta

def leer_detalle_vpe(context: BrowserContext) -> Dict[str, str]:
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

        datos["max"]   = td_value_any(page, ["Máximo", "Capacidad Máx.", "Capacidad máxima"])
        datos["min"]   = td_value_any(page, ["Mínimo", "Capacidad Mín.", "Capacidad mínima"])
        datos["e"]     = td_value(page, "e")
        datos["dd_dt"] = td_value_any(page, ["dd=dt", "dt", "dd", "d"])
        datos["clase"] = td_value(page, "Clase") or "III"

        datos["codigo_aprobacion"] = td_value_any(page, ["Código Aprobación", "Codigo Aprobación", "Codigo Aprobacion"])

    finally:
        try:
            page.close()
        except Exception:
            pass

    return datos

def leer_instrumento(context: BrowserContext, id_instrumento: str) -> Dict[str, any]:
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

        dom_block = td_value(page, "Domicilio", keep_newlines=True)
        dom, loc, prov = split_domicilio(dom_block)
        data["inst_dom"], data["inst_loc"], data["inst_prov"] = dom, loc, prov

        links = page.locator("a[href*='modeloDetalle.do']")
        hrefs = []
        for i in range(links.count()):
            h = links.nth(i).get_attribute("href") or ""
            if "modeloDetalle.do" in h:
                hrefs.append(h if h.startswith("http") else BASE + h)

        codes = td_values(page, "Código de Aprobación de Modelo") or td_values(page, "Código de Aprobación")
        series = td_values(page, "Nro de serie")

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

def extraer_camiones_por_ot(ot: str, user: str, pwd: str, mostrar_navegador: bool = False, 
                           log_callback=None, progress_callback=None) -> List[Dict[str, str]]:
    filas: List[Dict[str, str]] = []

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=not mostrar_navegador, slow_mo=0)
        context = browser.new_context()
        try:
            page, meta, instrument_links = login_y_abrir_ot(context, user, pwd, ot, log_callback)

            if log_callback:
                log_callback("📊 Extrayendo datos del propietario...")
            
            det = leer_detalle_vpe(context)
            nombre_usuario_det = det.get("nombre_usuario_instr", "").strip()
            direccion_legal_det = det.get("direccion_legal", "").strip()

            nro_ot = meta.get("ot", "").strip()
            vpe_num = meta.get("vpe", "").strip()
            empresa = meta.get("empresa_solicitante", "").strip()
            usuario_rep = nombre_usuario_det or meta.get("usuario_representado", "").strip()

            if log_callback:
                log_callback(f"\n📋 INFORMACIÓN DEL TRÁMITE:")
                log_callback(f"   • Número de O.T.: {nro_ot}")
                log_callback(f"   • VPE Nº: {vpe_num}")
                log_callback(f"   • Empresa: {empresa}")
                log_callback(f"   • Propietario: {usuario_rep}")
                log_callback("")

            if not instrument_links:
                if log_callback:
                    log_callback("⚠️ No se detectaron instrumentos en el VPE")
                return []

            total = len(instrument_links)
            if log_callback:
                log_callback(f"🔧 Procesando {total} instrumento(s)...\n")

            for idx, href in enumerate(instrument_links, start=1):
                if progress_callback:
                    progress_callback(idx, total)
                
                if log_callback:
                    log_callback(f"   [{idx}/{total}] Procesando instrumento...")

                m = re.search(r"idInstrumento=(\d+)", href)
                id_instrumento = m.group(1) if m else ""

                inst = leer_instrumento(context, id_instrumento) if id_instrumento else {
                    "inst_dom": "", "inst_loc": "", "inst_prov": "",
                    "receptor": {"href": "", "code": "", "serie": ""},
                    "indicador": {"href": "", "code": "", "serie": ""},
                }

                rec = inst["receptor"]
                ind = inst["indicador"]

                rec_model = leer_modelo_detalle(context, rec.get("href", "")) if rec.get("href") else {}
                ind_model = leer_modelo_detalle(context, ind.get("href", "")) if ind.get("href") else {}

                fab_rec   = (rec_model.get("fabricante") or "").strip()
                marca_rec = (rec_model.get("marca") or "").strip()
                modelo_rec= (rec_model.get("modelo") or "").strip()
                serie_rec = (rec.get("serie") or "").strip()
                codap_rec = (rec_model.get("codigo_aprobacion") or rec.get("code") or "").strip()
                origen_rec= (rec_model.get("origen") or "").strip()
                
                # CORRECCIÓN: e debe ser igual a dd=dt
                dd_dt_rec = (rec_model.get("dd_dt") or "").strip()
                e_rec     = dd_dt_rec  # e = dd=dt
                
                max_rec   = (rec_model.get("max") or "").strip()
                min_rec   = (rec_model.get("min") or "").strip()
                clase_rec = (rec_model.get("clase") or "").strip()
                naprob_rec= (rec_model.get("n_aprob") or "").strip()
                faprob_rec= (rec_model.get("fecha_aprob") or "").strip()

                fab_ind   = (ind_model.get("fabricante") or "").strip()
                marca_ind = (ind_model.get("marca") or "").strip()
                modelo_ind= (ind_model.get("modelo") or "").strip()
                serie_ind = (ind.get("serie") or "").strip()
                codap_ind = (ind_model.get("codigo_aprobacion") or ind.get("code") or "").strip()
                origen_ind= (ind_model.get("origen") or "").strip()
                naprob_ind= (ind_model.get("n_aprob") or "").strip()
                faprob_ind= (ind_model.get("fecha_aprob") or "").strip()

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

            if log_callback:
                log_callback(f"\n✅ Se procesaron {len(filas)} instrumento(s) correctamente")

        finally:
            try:
                browser.close()
            except Exception:
                pass

    return filas

# =========================
# Excel (FORMATO 2 COLUMNAS)
# =========================

COLUMNS_ORDER = [
    "Número de O.T.", "VPE Nº", "Empresa solicitante", "Razón social (Propietario)",
    "Domicilio (Fiscal)", "Localidad (Fiscal)", "Provincia (Fiscal)",
    "Lugar propio de instalación - Domicilio", "Lugar propio de instalación - Localidad",
    "Lugar propio de instalación - Provincia", "Instrumento verificado",
    "Fabricante receptor", "Marca Receptor", "Modelo Receptor", "N° de serie Receptor",
    "Cód ap. mod. Receptor", "Origen Receptor", "e", "máx", "mín", "dd=dt", "clase",
    "N° de Aprobación Modelo (Receptor)", "Fecha de Aprobación Modelo (Receptor)",
    "Tipo (Indicador)", "Fabricante Indicador", "Marca Indicador", "Modelo Indicador",
    "N° de serie Indicador", "Código Aprobación (Indicador)", "Origen Indicador",
    "N° de Aprobación Modelo (Indicador)", "Fecha de Aprobación Modelo (Indicador)"
]

def armar_hoja_verificacion_2columnas(filas: List[Dict[str, str]]) -> pd.DataFrame:
    """
    Arma el DataFrame en formato de 2 columnas (Campo | Valor)
    Si hay múltiples instrumentos, los separa con filas vacías
    """
    if not filas:
        return pd.DataFrame(columns=["Campo", "Valor"])
    
    data_final = []
    
    for idx, fila in enumerate(filas, start=1):
        if idx > 1:
            # Separador entre instrumentos
            data_final.append({"Campo": "", "Valor": ""})
            data_final.append({"Campo": f"=== INSTRUMENTO {idx} ===", "Valor": ""})
        
        # Agregar cada campo como una fila
        for col in COLUMNS_ORDER:
            valor = fila.get(col, "")
            data_final.append({
                "Campo": col,
                "Valor": valor
            })
    
    df = pd.DataFrame(data_final)
    return df

def exportar_verificacion_2columnas(df: pd.DataFrame, ruta: Path) -> Path:
    """
    Exporta el DataFrame en formato de 2 columnas con formato Excel
    """
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
            "valign": "vcenter",
            "font_size": 11
        })

        fmt_campo = wb.add_format({
            "bold": True,
            "bg_color": "#4472C4",
            "font_color": "white",
            "border": 1,
            "align": "left",
            "valign": "vcenter",
            "text_wrap": True
        })

        fmt_valor = wb.add_format({
            "border": 1,
            "align": "left",
            "valign": "top",
            "text_wrap": True
        })
        
        fmt_separador = wb.add_format({
            "bold": True,
            "bg_color": "#FFC000",
            "font_color": "#000000",
            "border": 1,
            "align": "center",
            "valign": "vcenter"
        })

        # Escribir cabeceras
        ws.write(0, 0, "Campo", fmt_header)
        ws.write(0, 1, "Valor", fmt_header)

        # Escribir datos con formato
        for row_num in range(1, len(df) + 1):
            campo = df.iloc[row_num - 1, 0]
            valor = df.iloc[row_num - 1, 1]
            
            # Detectar separadores
            if campo.startswith("==="):
                ws.write(row_num, 0, campo, fmt_separador)
                ws.write(row_num, 1, valor, fmt_separador)
            elif campo == "":
                ws.write(row_num, 0, "", fmt_valor)
                ws.write(row_num, 1, "", fmt_valor)
            else:
                ws.write(row_num, 0, campo, fmt_campo)
                ws.write(row_num, 1, valor, fmt_valor)

        # Ajustar anchos de columna
        ws.set_column(0, 0, 45)  # Columna Campo
        ws.set_column(1, 1, 60)  # Columna Valor

        # Congelar primera fila
        ws.freeze_panes(1, 0)
        ws.set_row(0, 25)  # Altura de cabecera

    return ruta

# =========================
# Utilidad para limpiar nombre de archivo
# =========================

def limpiar_nombre_archivo(texto: str) -> str:
    """
    Limpia un texto para que sea válido como nombre de archivo
    """
    # Reemplazar caracteres inválidos
    invalidos = ['<', '>', ':', '"', '/', '\\', '|', '?', '*']
    for char in invalidos:
        texto = texto.replace(char, '_')
    
    # Limitar longitud y eliminar espacios al inicio/final
    texto = texto.strip()[:100]
    
    return texto

# =========================
# INTERFAZ GRÁFICA (GUI)
# =========================

class ModernButton(tk.Canvas):
    """Botón moderno personalizado con efectos hover"""
    def __init__(self, parent, text, command, bg_color="#4472C4", fg_color="white", 
                 hover_color="#365a9b", width=200, height=40, **kwargs):
        super().__init__(parent, width=width, height=height, highlightthickness=0, **kwargs)
        
        self.command = command
        self.bg_color = bg_color
        self.hover_color = hover_color
        self.fg_color = fg_color
        self.text = text
        
        self.configure(bg=parent['bg'])
        
        # Dibujar botón
        self.rect = self.create_rectangle(2, 2, width-2, height-2, 
                                          fill=bg_color, outline="", width=0)
        self.text_id = self.create_text(width//2, height//2, 
                                        text=text, fill=fg_color, 
                                        font=("Segoe UI", 10, "bold"))
        
        # Bindings
        self.bind("<Enter>", self._on_enter)
        self.bind("<Leave>", self._on_leave)
        self.bind("<Button-1>", self._on_click)
    
    def _on_enter(self, e):
        self.itemconfig(self.rect, fill=self.hover_color)
        self.configure(cursor="hand2")
    
    def _on_leave(self, e):
        self.itemconfig(self.rect, fill=self.bg_color)
        self.configure(cursor="")
    
    def _on_click(self, e):
        if self.command:
            self.command()
    
    def config_state(self, state):
        if state == "disabled":
            self.itemconfig(self.rect, fill="#cccccc")
            self.unbind("<Enter>")
            self.unbind("<Leave>")
            self.unbind("<Button-1>")
        else:
            self.itemconfig(self.rect, fill=self.bg_color)
            self.bind("<Enter>", self._on_enter)
            self.bind("<Leave>", self._on_leave)
            self.bind("<Button-1>", self._on_click)


class ExtractorGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("INTI MetroWeb - Extractor de Verificación Previa")
        self.root.geometry("800x700")
        self.root.resizable(False, False)
        # Colores modernos
        self.bg_color = "#f0f0f0"
        self.card_color = "#ffffff"
        self.primary_color = "#4472C4"
        self.success_color = "#28a745"
        self.warning_color = "#ffc107"
        self.error_color = "#dc3545"
        
        self.root.configure(bg=self.bg_color)
        
        # Variables
        self.usuario_var = tk.StringVar()
        self.password_var = tk.StringVar()
        self.ot_var = tk.StringVar()
        self.headless_var = tk.BooleanVar(value=True)
        self.filas_extraidas = []
        self.razon_social = ""  # Guardar razón social para el nombre del archivo
        
        self.crear_interfaz()
    
    def crear_interfaz(self):
        # Header
        header_frame = tk.Frame(self.root, bg=self.primary_color, height=80)
        header_frame.pack(fill="x", pady=(0, 20))
        header_frame.pack_propagate(False)
        
        title_label = tk.Label(header_frame, 
                              text="🏭 INTI METROWEB",
                              font=("Segoe UI", 20, "bold"),
                              bg=self.primary_color,
                              fg="white")
        title_label.pack(pady=(10, 0))
        
        subtitle_label = tk.Label(header_frame,
                                 text="Extractor de Verificación Previa - Balanzas de Camiones",
                                 font=("Segoe UI", 10),
                                 bg=self.primary_color,
                                 fg="white")
        subtitle_label.pack()
        
        # Main container
        main_container = tk.Frame(self.root, bg=self.bg_color)
        main_container.pack(fill="both", expand=True, padx=20, pady=10)
        
        # Card de credenciales
        cred_card = self.crear_card(main_container, "🔐 Credenciales de Acceso")
        cred_card.pack(fill="x", pady=(0, 15))
        
        # Usuario
        self.crear_campo(cred_card, "Usuario MetroWeb:", self.usuario_var)
        
        # Password
        self.crear_campo(cred_card, "Contraseña:", self.password_var, show="*")
        
        # Card de OT
        ot_card = self.crear_card(main_container, "📋 Orden de Trabajo")
        ot_card.pack(fill="x", pady=(0, 15))
        
        self.crear_campo(ot_card, "Número de OT (ej. 307-62136):", self.ot_var)
        
        # Checkbox navegador
        check_frame = tk.Frame(ot_card, bg=self.card_color)
        check_frame.pack(fill="x", padx=15, pady=(5, 10))
        
        check = tk.Checkbutton(check_frame,
                              text="Ejecutar en modo oculto (headless)",
                              variable=self.headless_var,
                              bg=self.card_color,
                              font=("Segoe UI", 9),
                              activebackground=self.card_color)
        check.pack(anchor="w")
        
        # Botón de extracción
        btn_frame = tk.Frame(main_container, bg=self.bg_color)
        btn_frame.pack(pady=10)
        
        self.extract_btn = ModernButton(btn_frame,
                                        text="🚀 INICIAR EXTRACCIÓN",
                                        command=self.iniciar_extraccion,
                                        bg_color=self.success_color,
                                        hover_color="#218838",
                                        width=250,
                                        height=50)
        self.extract_btn.pack()
        
        # Card de progreso
        progress_card = self.crear_card(main_container, "📊 Progreso de Extracción")
        progress_card.pack(fill="both", expand=True, pady=(0, 10))
        
        # Barra de progreso
        progress_frame = tk.Frame(progress_card, bg=self.card_color)
        progress_frame.pack(fill="x", padx=15, pady=(10, 5))
        
        self.progress = ttk.Progressbar(progress_frame,
                                       mode='determinate',
                                       length=700)
        self.progress.pack(fill="x")
        
        self.progress_label = tk.Label(progress_frame,
                                       text="Esperando inicio...",
                                       font=("Segoe UI", 9),
                                       bg=self.card_color,
                                       fg="#666666")
        self.progress_label.pack(pady=(5, 0))
        
        # Log de eventos
        log_frame = tk.Frame(progress_card, bg=self.card_color)
        log_frame.pack(fill="both", expand=True, padx=15, pady=(5, 15))
        
        log_label = tk.Label(log_frame,
                            text="Registro de eventos:",
                            font=("Segoe UI", 9, "bold"),
                            bg=self.card_color,
                            anchor="w")
        log_label.pack(anchor="w", pady=(0, 5))
        
        self.log_text = scrolledtext.ScrolledText(log_frame,
                                                  height=10,
                                                  font=("Consolas", 9),
                                                  bg="#1e1e1e",
                                                  fg="#d4d4d4",
                                                  insertbackground="white",
                                                  relief="flat",
                                                  borderwidth=0)
        self.log_text.pack(fill="both", expand=True)
        self.log_text.config(state="disabled")
        
        # Footer
        footer = tk.Label(self.root,
                         text="INTI - Instituto Nacional de Tecnología Industrial",
                         font=("Segoe UI", 8),
                         bg=self.bg_color,
                         fg="#666666")
        footer.pack(side="bottom", pady=(0, 10))
    
    def crear_card(self, parent, titulo):
        """Crea una tarjeta con título"""
        card = tk.Frame(parent, bg=self.card_color, relief="flat", borderwidth=1)
        card.configure(highlightbackground="#dddddd", highlightthickness=1)
        
        title_frame = tk.Frame(card, bg=self.card_color)
        title_frame.pack(fill="x", padx=15, pady=(15, 10))
        
        title_label = tk.Label(title_frame,
                              text=titulo,
                              font=("Segoe UI", 11, "bold"),
                              bg=self.card_color,
                              fg="#333333")
        title_label.pack(anchor="w")
        
        separator = tk.Frame(card, height=1, bg="#e0e0e0")
        separator.pack(fill="x", padx=15)
        
        return card
    
    def crear_campo(self, parent, label_text, variable, show=None):
        """Crea un campo de entrada con etiqueta"""
        frame = tk.Frame(parent, bg=self.card_color)
        frame.pack(fill="x", padx=15, pady=8)
        
        label = tk.Label(frame,
                        text=label_text,
                        font=("Segoe UI", 9),
                        bg=self.card_color,
                        fg="#555555",
                        anchor="w")
        label.pack(anchor="w", pady=(0, 5))
        
        entry_frame = tk.Frame(frame, bg="white", relief="solid", borderwidth=1)
        entry_frame.pack(fill="x")
        
        entry = tk.Entry(entry_frame,
                        textvariable=variable,
                        font=("Segoe UI", 10),
                        relief="flat",
                        bg="white",
                        show=show)
        entry.pack(fill="x", padx=8, pady=8)
        
        return entry
    
    def log(self, mensaje):
        """Agrega mensaje al log"""
        self.log_text.config(state="normal")
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.insert("end", f"[{timestamp}] {mensaje}\n")
        self.log_text.see("end")
        self.log_text.config(state="disabled")
        self.root.update()
    
    def actualizar_progreso(self, actual, total):
        """Actualiza barra de progreso"""
        porcentaje = (actual / total * 100) if total > 0 else 0
        self.progress['value'] = porcentaje
        self.progress_label.config(text=f"Procesando: {actual} de {total} instrumentos ({porcentaje:.1f}%)")
        self.root.update()
    
    def validar_campos(self):
        """Valida que los campos estén completos"""
        if not self.usuario_var.get().strip():
            messagebox.showerror("Error", "Debe ingresar el usuario")
            return False
        
        if not self.password_var.get().strip():
            messagebox.showerror("Error", "Debe ingresar la contraseña")
            return False
        
        if not self.ot_var.get().strip():
            messagebox.showerror("Error", "Debe ingresar el número de OT")
            return False
        
        # Validar formato OT
        ot = self.ot_var.get().strip()
        if not re.match(r"^\d{3}-\d{5}$", ot):
            respuesta = messagebox.askyesno(
                "Formato no estándar",
                f"El formato de OT '{ot}' no es estándar.\n"
                "Se esperaba: XXX-XXXXX (ej. 307-62136)\n\n"
                "¿Desea continuar de todas formas?"
            )
            if not respuesta:
                return False
        
        return True
    
    def iniciar_extraccion(self):
        """Inicia el proceso de extracción en un hilo separado"""
        if not self.validar_campos():
            return
        
        # Limpiar log
        self.log_text.config(state="normal")
        self.log_text.delete(1.0, "end")
        self.log_text.config(state="disabled")
        
        # Resetear progreso
        self.progress['value'] = 0
        self.progress_label.config(text="Iniciando extracción...")
        
        # Deshabilitar botón
        self.extract_btn.config_state("disabled")
        
        # Iniciar en hilo separado
        thread = threading.Thread(target=self.ejecutar_extraccion, daemon=True)
        thread.start()
    
    def ejecutar_extraccion(self):
        """Ejecuta la extracción (corre en hilo separado)"""
        try:
            usuario = self.usuario_var.get().strip()
            password = self.password_var.get().strip()
            ot = self.ot_var.get().strip()
            headless = self.headless_var.get()
            
            self.log("=" * 50)
            self.log("🚀 INICIANDO EXTRACCIÓN")
            self.log("=" * 50)
            self.log("")
            
            # Extraer datos
            self.filas_extraidas = extraer_camiones_por_ot(
                ot=ot,
                user=usuario,
                pwd=password,
                mostrar_navegador=not headless,
                log_callback=self.log,
                progress_callback=self.actualizar_progreso
            )
            
            if not self.filas_extraidas:
                self.log("")
                self.log("⚠️ No se encontraron datos para extraer")
                messagebox.showwarning(
                    "Sin datos",
                    "No se encontraron instrumentos en la OT especificada.\n"
                    "Verifique el número de OT y las credenciales."
                )
                self.extract_btn.config_state("normal")
                return
            
            # Guardar razón social para el nombre del archivo
            self.razon_social = self.filas_extraidas[0].get("Razón social (Propietario)", "")
            
            self.log("")
            self.log("=" * 50)
            self.log(f"✅ EXTRACCIÓN COMPLETADA: {len(self.filas_extraidas)} instrumento(s)")
            self.log("=" * 50)
            self.log("")
            
            # Solicitar ubicación de guardado
            self.root.after(500, self.solicitar_guardado)
            
        except Exception as e:
            self.log("")
            self.log(f"❌ ERROR: {str(e)}")
            messagebox.showerror("Error", f"Error durante la extracción:\n\n{str(e)}")
            self.extract_btn.config_state("normal")
    
    def solicitar_guardado(self):
        """Solicita ubicación para guardar el archivo"""
        ot = self.ot_var.get().strip()
        
        # Limpiar razón social para nombre de archivo
        razon_limpia = limpiar_nombre_archivo(self.razon_social) if self.razon_social else "SIN_RAZON"
        
        # Formato: OT_307-63160_RAZON_SOCIAL.xlsx
        nombre_sugerido = f"OT_{ot}_{razon_limpia}.xlsx"
        
        archivo = filedialog.asksaveasfilename(
            title="Guardar archivo Excel",
            defaultextension=".xlsx",
            initialfile=nombre_sugerido,
            filetypes=[
                ("Archivos Excel", "*.xlsx"),
                ("Todos los archivos", "*.*")
            ]
        )
        
        if not archivo:
            self.log("⚠️ Guardado cancelado por el usuario")
            messagebox.showinfo(
                "Cancelado",
                "No se guardó el archivo.\n"
                "Los datos extraídos se perderán."
            )
            self.extract_btn.config_state("normal")
            return
        
        # Guardar archivo
        try:
            self.log("")
            self.log("💾 Generando archivo Excel (formato 2 columnas)...")
            
            # Usar función de 2 columnas
            df = armar_hoja_verificacion_2columnas(self.filas_extraidas)
            ruta = exportar_verificacion_2columnas(df, Path(archivo))
            
            tamaño_kb = ruta.stat().st_size / 1024
            
            self.log(f"✅ Archivo guardado exitosamente")
            self.log(f"   📁 Ubicación: {ruta.resolve()}")
            self.log(f"   📊 Tamaño: {tamaño_kb:.2f} KB")
            self.log(f"   📝 Instrumentos: {len(self.filas_extraidas)}")
            self.log(f"   📋 Formato: 2 columnas (Campo | Valor)")
            self.log("")
            self.log("=" * 50)
            self.log("🎉 PROCESO COMPLETADO")
            self.log("=" * 50)
            
            # Mensaje de éxito
            respuesta = messagebox.askyesno(
                "✅ Éxito",
                f"Archivo generado correctamente:\n\n"
                f"📁 {ruta.name}\n"
                f"📊 {len(self.filas_extraidas)} instrumento(s) procesados\n"
                f"💾 {tamaño_kb:.2f} KB\n"
                f"📋 Formato: 2 columnas\n\n"
                f"¿Desea abrir la carpeta donde se guardó?"
            )
            
            if respuesta:
                import os
                import platform
                
                carpeta = ruta.parent
                if platform.system() == 'Windows':
                    os.startfile(carpeta)
                elif platform.system() == 'Darwin':  # macOS
                    os.system(f'open "{carpeta}"')
                else:  # Linux
                    os.system(f'xdg-open "{carpeta}"')
            
        except Exception as e:
            self.log(f"❌ ERROR al guardar: {str(e)}")
            messagebox.showerror("Error", f"No se pudo guardar el archivo:\n\n{str(e)}")
        
        finally:
            self.extract_btn.config_state("normal")
            self.progress['value'] = 0
            self.progress_label.config(text="Proceso completado - Listo para nueva extracción")


# =========================
# PUNTO DE ENTRADA
# =========================

def main():
    root = tk.Tk()
    
    # Configurar ícono (si existe)
    try:
        # Puedes agregar un ícono .ico aquí si lo tienes
        # root.iconbitmap("icono.ico")
        pass
    except:
        pass
    
    app = ExtractorGUI(root)
    
    # Centrar ventana
    root.update_idletasks()
    width = root.winfo_width()
    height = root.winfo_height()
    x = (root.winfo_screenwidth() // 2) - (width // 2)
    y = (root.winfo_screenheight() // 2) - (height // 2)
    root.geometry(f'{width}x{height}+{x}+{y}')
    
    root.mainloop()


if __name__ == "__main__":
    main()