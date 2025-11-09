# -*- coding: utf-8 -*-
"""
Capa portal (scraper): INTI MetroWeb → datos estructurados
Implementa extraer_camiones_por_ot(...) y helpers compatibles con la GUI.

Requisitos:
  pip install playwright
  python -m playwright install chromium
"""

from __future__ import annotations

import re
import time
from datetime import datetime
from pathlib import Path
from typing import Callable, Dict, List, Optional, Tuple

from playwright.sync_api import BrowserContext, Page, sync_playwright

BASE = "https://app.inti.gob.ar"

# ------------------------------------------------------------------------------
# Helpers de texto / extracción de celdas
# ------------------------------------------------------------------------------

def _clean_one_line(s: str) -> str:
    s = (s or "").replace("\xa0", " ").strip()
    return re.sub(r"\s+", " ", s)

def only_digits(s: str) -> str:
    return "".join(ch for ch in (s or "") if ch.isdigit())

def td_value(page: Page, label: str, keep_newlines: bool = False, nth: int = 0) -> str:
    """
    Devuelve el texto de la TD siguiente a la TD que contiene 'label'.
    Si existen varias coincidencias, usa el índice 'nth'.
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
    Devuelve todas las celdas TD (sibling inmediato) que siguen a una TD que contenga 'label'.
    """
    loc = page.locator(
        f"xpath=//td[contains(normalize-space(.), '{label}')]/following-sibling::td[1]"
    )
    out: List[str] = []
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
    """
    Prueba varias etiquetas alternativas y devuelve el primer valor no vacío.
    """
    for lb in labels:
        v = td_value(page, lb, keep_newlines=keep_newlines)
        if v:
            return v
    return ""

def split_domicilio(block_text: str) -> Tuple[str, str, str]:
    """
    Divide un bloque multilinea en (domicilio, localidad, provincia).
    """
    if not block_text:
        return "", "", ""
    parts = [p.strip() for p in block_text.replace("\r", "\n").split("\n") if p.strip()]
    dom = parts[0] if len(parts) > 0 else ""
    loc = parts[1] if len(parts) > 1 else ""
    prov = parts[2] if len(parts) > 2 else ""
    return dom, loc, prov

# ------------------------------------------------------------------------------
# Navegación MetroWeb
# ------------------------------------------------------------------------------

def login_y_abrir_ot(
    context: BrowserContext,
    usuario: str,
    password: str,
    ot: str,
    log_callback: Optional[Callable[[str], None]] = None,
) -> Tuple[Page, Dict[str, str], List[str]]:
    page = context.new_page()
    page.set_default_timeout(60_000)

    if log_callback:
        log_callback("🔗 Conectando con MetroWeb...")

    page.goto(f"{BASE}/MetroWeb/pages/ingreso.jsp")

    # Usuario
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

    # Enviar
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
        log_callback(f"🔍 Buscando OT: {ot}")

    page.goto(f"{BASE}/MetroWeb/entrarPML.do")

    if page.locator('input[name="numeroOT"]').count():
        page.fill('input[name="numeroOT"]', ot)
    elif page.locator('input[name="nroOT"]').count():
        page.fill('input[name="nroOT"]', ot)
    else:
        caja = page.locator(
            "xpath=//*[contains(normalize-space(.),'Número OT') or contains(normalize-space(.),'Nmero OT')]/following::input[1]"
        )
        caja.fill(ot)

    if page.locator('input[value="Buscar"]').count():
        page.click('input[value="Buscar"]')
    else:
        page.keyboard.press("Enter")
    page.wait_for_load_state("networkidle")

    link_vpe = page.locator('a[href*="tramiteVPE"]').first
    if not link_vpe or not link_vpe.count():
        raise RuntimeError("No se encontró enlace de trámite VPE para esa OT")

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

    instrument_links: List[str] = []
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
        "usuario_representado": "",
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
    datos: Dict[str, str] = {
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
        "codigo_aprobacion": "",
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
        datos["n_aprob"] = td_value_any(
            page,
            [
                "Nº Disposicion",
                "N° Disposicion",
                "Nº Disposición",
                "N° Disposición",
                "Nº Disposici",
                "N° Disposici",
                "N° de Aprobación",
                "Nº de Aprobación",
            ],
        )
        datos["fecha_aprob"] = td_value_any(page, ["Fecha Aprobación", "Fecha de Aprobación"])
        datos["tipo_instr"] = td_value_any(page, ["Tipo Instrumento", "Tipo de Instrumento"])

        datos["max"] = td_value_any(page, ["Máximo", "Capacidad Máx.", "Capacidad máxima"])
        datos["min"] = td_value_any(page, ["Mínimo", "Capacidad Mín.", "Capacidad mínima"])
        datos["e"] = td_value(page, "e")
        datos["dd_dt"] = td_value_any(page, ["dd=dt", "dt", "dd", "d"])
        datos["clase"] = td_value(page, "Clase") or "III"
        datos["codigo_aprobacion"] = td_value_any(
            page, ["Código Aprobación", "Codigo Aprobación", "Codigo Aprobacion"]
        )
    finally:
        try:
            page.close()
        except Exception:
            pass

    return datos

def leer_instrumento(context: BrowserContext, id_instrumento: str) -> Dict[str, object]:
    url = f"{BASE}/MetroWeb/instrumentoDetalle.do?idInstrumento={id_instrumento}"
    page = context.new_page()
    page.set_default_timeout(60_000)

    data: Dict[str, object] = {
        "inst_dom": "",
        "inst_loc": "",
        "inst_prov": "",
        "receptor": {"href": "", "code": "", "serie": ""},
        "indicador": {"href": "", "code": "", "serie": ""},
    }

    try:
        page.goto(url)
        page.wait_for_load_state("networkidle")
        time.sleep(0.3)

        dom_block = td_value(page, "Domicilio", keep_newlines=True)
        dom, loc, prov = split_domicilio(dom_block)
        data["inst_dom"], data["inst_loc"], data["inst_prov"] = dom, loc, prov

        links = page.locator("a[href*='modeloDetalle.do']")
        hrefs: List[str] = []
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

# ------------------------------------------------------------------------------
# Punto principal llamado por la GUI
# ------------------------------------------------------------------------------

def extraer_camiones_por_ot(
    ot: str,
    user: str,
    pwd: str,
    mostrar_navegador: bool = False,
    log_callback: Optional[Callable[[str], None]] = None,
    progress_callback: Optional[Callable[[int, int], None]] = None,
) -> List[Dict[str, str]]:
    """
    Retorna una lista de filas (dict campo->valor) para exportar a Excel.
    """
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
                log_callback("\n📋 INFORMACIÓN DEL TRÁMITE:")
                log_callback(f"   • Número de O.T.: {nro_ot}")
                log_callback(f"   • VPE Nº: {vpe_num}")
                log_callback(f"   • Empresa: {empresa}")
                log_callback(f"   • Propietario: {usuario_rep}\n")

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

                inst = (
                    leer_instrumento(context, id_instrumento)
                    if id_instrumento
                    else {"inst_dom": "", "inst_loc": "", "inst_prov": "", "receptor": {}, "indicador": {}}
                )

                rec = inst["receptor"]  # type: ignore[index]
                ind = inst["indicador"]  # type: ignore[index]

                rec_model = leer_modelo_detalle(context, rec.get("href", "")) if rec.get("href") else {}
                ind_model = leer_modelo_detalle(context, ind.get("href", "")) if ind.get("href") else {}

                fab_rec = (rec_model.get("fabricante") or "").strip()
                marca_rec = (rec_model.get("marca") or "").strip()
                modelo_rec = (rec_model.get("modelo") or "").strip()
                serie_rec = (rec.get("serie") or "").strip()
                codap_rec = (rec_model.get("codigo_aprobacion") or rec.get("code") or "").strip()
                origen_rec = (rec_model.get("origen") or "").strip()

                # e = dd=dt (corrección usada en tu script original)
                dd_dt_rec = (rec_model.get("dd_dt") or "").strip()
                e_rec = dd_dt_rec

                max_rec = (rec_model.get("max") or "").strip()
                min_rec = (rec_model.get("min") or "").strip()
                clase_rec = (rec_model.get("clase") or "").strip()
                naprob_rec = (rec_model.get("n_aprob") or "").strip()
                faprob_rec = (rec_model.get("fecha_aprob") or "").strip()

                fab_ind = (ind_model.get("fabricante") or "").strip()
                marca_ind = (ind_model.get("marca") or "").strip()
                modelo_ind = (ind_model.get("modelo") or "").strip()
                serie_ind = (ind.get("serie") or "").strip()
                codap_ind = (ind_model.get("codigo_aprobacion") or ind.get("code") or "").strip()
                origen_ind = (ind_model.get("origen") or "").strip()
                naprob_ind = (ind_model.get("n_aprob") or "").strip()
                faprob_ind = (ind_model.get("fecha_aprob") or "").strip()

                lugar_dom = inst.get("inst_dom", "")  # type: ignore[assignment]
                lugar_loc = inst.get("inst_loc", "")
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
