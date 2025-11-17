# extract_camiones.py
# -*- coding: utf-8 -*-

"""
Extracción INTI MetroWeb → Excel (Verificación Previa) para balanzas de camiones/plataforma.
Requisitos:
  pip install playwright pandas xlsxwriter
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


BASE = "https://app.inti.gob.ar"


# =========================
# Utilidades de scraping
# =========================

def _clean_one_line(s: str) -> str:
    s = (s or "").replace("\xa0", " ").strip()
    return re.sub(r"\s+", " ", s)

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

def td_values(page: Page, label: str, keep_newlines: bool = False, max_items: int = 10) -> List[str]:
    out = []
    for i in range(max_items):
        v = td_value(page, label, keep_newlines=keep_newlines, nth=i)
        if not v:
            break
        out.append(v)
    return out

def td_value_any(page: Page, labels: List[str], keep_newlines: bool = False, nth: int = 0) -> str:
    for lb in labels:
        v = td_value(page, lb, keep_newlines=keep_newlines, nth=nth)
        if v:
            return v
    return ""

def split_domicilio(block_text: str) -> Tuple[str, str, str]:
    """
    Devuelve (domicilio, localidad, provincia) a partir del bloque multilínea del sitio.
    """
    if not block_text:
        return "", "", ""
    parts = [p.strip() for p in block_text.replace("\r", "\n").split("\n") if p.strip()]
    dom = parts[0] if len(parts) > 0 else ""
    loc = parts[1] if len(parts) > 1 else ""
    prov = parts[2] if len(parts) > 2 else ""
    return dom, loc, prov

def only_digits(s: str) -> str:
    return "".join(re.findall(r"\d+", s or ""))


# =========================
# Navegación / Login / OT
# =========================

def login_y_abrir_ot(context: BrowserContext, usuario: str, password: str, ot: str) -> Tuple[Page, Dict[str, str], List[str]]:
    """
    Inicia sesión, busca la OT y entra al VPE.
    Luego abre resumen.jsp y devuelve:
      - page (posicionada en resumen.jsp)
      - meta: dict con 'ot', 'vpe', 'empresa_solicitante', 'usuario_representado'
      - instrument_links: lista de hrefs a instrumentoDetalle.do
    """
    page = context.new_page()
    page.set_default_timeout(60_000)

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

    print("🔐 Iniciando sesión…")
    page.wait_for_load_state("networkidle")

    # Buscar OT
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
        print("⚠️ No se encontró enlace de trámite VPE para esa OT.")
        return page, {"ot": ot, "vpe": "", "empresa_solicitante": "", "usuario_representado": ""}, []

    vpe_text = _clean_one_line(link_vpe.inner_text())
    vpe_num = only_digits(vpe_text)
    link_vpe.click()
    page.wait_for_load_state("networkidle")

    # Ir a resumen.jsp (fuente principal de datos generales)
    page.goto(f"{BASE}/MetroWeb/pages/tramiteVPE/resumen.jsp")
    page.wait_for_load_state("networkidle")
    time.sleep(0.5)

    meta = leer_resumen(page)
    # Asegurar OT y VPE desde lo que ya sabemos
    if not meta.get("ot"):
        meta["ot"] = ot
    if not meta.get("vpe"):
        meta["vpe"] = vpe_num

    # Enlaces a instrumentos desde el resumen (generalmente están allí)
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

    # OT (Nro OT)
    ot_val = td_value(page, "Nro OT") or td_value(page, "N° OT") or td_value(page, "Número de O.T.") or ""
    meta["ot"] = _clean_one_line(ot_val)

    # VPE (buscar texto 'vpeXXXXX' en la página)
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

    # Empresa Solicitante (sale de resumen.jsp)
    meta["empresa_solicitante"] = td_value(page, "Empresa Solicitante")

    # Usuario Representado (igual a Razón social Propietario si no hay en detalle.jsp)
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
    if not href.startswith("http"):
        href = BASE + href

    page = context.new_page()
    page.set_default_timeout(60_000)
    try:
        page.goto(href)
        page.wait_for_load_state("networkidle")
        time.sleep(0.3)

        # Datos Generales
        datos["modelo"]         = td_value(page, "Modelo")
        datos["marca"]          = td_value(page, "Marca")
        datos["fabricante"]     = td_value_any(page, ["Fabricante", "Fabricante/Importador", "Importador"])
        datos["origen"]         = td_value_any(page, ["País Origen", "País de Origen", "Origen"])
        datos["n_aprob"]        = td_value_any(page, [
                                    "Nº Disposicion", "N° Disposicion", "Nº Disposición", "N° Disposición",
                                    "N° de Aprobación", "Nº de Aprobación", "Nº Disp", "N° Disp"
                                ])
        datos["fecha_aprob"]    = td_value_any(page, ["Fecha Aprobación", "Fecha de Aprobación"])
        datos["tipo_instr"]     = td_value(page, "Tipo Instrumento")
        datos["codigo_aprobacion"] = td_value_any(page, ["Código Aprobación", "Código de Aprobación"])

        # Características metrológicas (si aplica; relevante para receptor)
        datos["max"]   = td_value_any(page, ["Máximo", "Capacidad Máx.", "Capacidad máxima"])
        datos["min"]   = td_value_any(page, ["Mínimo", "Capacidad Mín.", "Capacidad mínima"])
        datos["e"]     = td_value(page, "e")
        datos["dd_dt"] = td_value_any(page, ["dd=dt", "dt", "dd", "d"])
        datos["clase"] = td_value(page, "Clase") or "III"

    finally:
        try:
            page.close()
        except Exception:
            pass

    return datos

def leer_instrumento(context: BrowserContext, id_instrumento: str) -> Dict[str, any]:
    """
    Abre instrumentoDetalle.do?idInstrumento=... y extrae:
      - Domicilio (para lugar de instalación)
      - Receptor: href modelo, código de aprobación, nro de serie
      - Indicador: href modelo, código de aprobación, nro de serie
    Heurística: se toma el primer bloque como receptor y el segundo como indicador.
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

        # Ubicación (Lugar propio de instalación)
        dom_block = td_value(page, "Domicilio", keep_newlines=True)
        dom, loc, prov = split_domicilio(dom_block)
        data["inst_dom"], data["inst_loc"], data["inst_prov"] = dom, loc, prov

        # Links a modelos dentro del detalle del instrumento
        links = page.locator("a[href*='modeloDetalle.do']")
        hrefs = []
        for i in range(links.count()):
            h = links.nth(i).get_attribute("href") or ""
            if "modeloDetalle.do" in h:
                hrefs.append(h if h.startswith("http") else BASE + h)

        # Códigos de aprobación de modelo (dos posibles bloques)
        codes = td_values(page, "Código de Aprobación de Modelo") or td_values(page, "Código de Aprobación")

        # Series (dos posibles bloques)
        series = td_values(page, "Nro de serie")

        # Mapear por posición: 0=receptor, 1=indicador
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
    con las 30 columnas solicitadas (completadas o vacías).
    """
    filas: List[Dict[str, str]] = []

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=not mostrar_navegador, slow_mo=0)
        context = browser.new_context()
        try:
            # 1) Login + entrar a VPE + resumen.jsp
            page, meta, instrument_links = login_y_abrir_ot(context, user, pwd, ot)

            # 1.b) detalle.jsp → Nombre del Usuario del Instrumento / Dirección Legal
            det = leer_detalle_vpe(context)
            nombre_usuario_det = det.get("nombre_usuario_instr", "").strip()
            direccion_legal_det = det.get("direccion_legal", "").strip()

            # Validaciones mínimas desde resumen.jsp
            nro_ot = meta.get("ot", "").strip()
            vpe_num = meta.get("vpe", "").strip()
            empresa = meta.get("empresa_solicitante", "").strip()

            # Razón social (Propietario): priorizar detalle.jsp; si no, usar 'Usuario Representado' de resumen.jsp
            usuario_rep = nombre_usuario_det or meta.get("usuario_representado", "").strip()

            if not instrument_links:
                print("⚠️ No se detectaron instrumentos en el VPE.")
                instrument_links = []  # seguirá vacío, generará 0 filas

            print(f"📄 OT: {nro_ot} | VPE: {vpe_num} | Empresa Solicitante: {empresa} | Usuario (detalle.jsp): {nombre_usuario_det or 'N/D'}")
            print(f"🔗 Instrumentos encontrados: {len(instrument_links)}")

            # 2) Recorrer instrumentos
            for idx, href in enumerate(instrument_links, start=1):
                # idInstrumento
                m = re.search(r"idInstrumento=(\d+)", href)
                id_instrumento = m.group(1) if m else ""

                inst = leer_instrumento(context, id_instrumento) if id_instrumento else {
                    "inst_dom": "", "inst_loc": "", "inst_prov": "",
                    "receptor": {"href": "", "code": "", "serie": ""},
                    "indicador": {"href": "", "code": "", "serie": ""},
                }

                # Abrir modelos de receptor e indicador
                rec = inst["receptor"]
                ind = inst["indicador"]

                rec_model = leer_modelo_detalle(context, rec.get("href", "")) if rec.get("href") else {}
                ind_model = leer_modelo_detalle(context, ind.get("href", "")) if ind.get("href") else {}

                # Resolver campos receptor (modeloDetalle > instrumentoDetalle)
                fab_rec  = (rec_model.get("fabricante") or "").strip()
                marca_rec= (rec_model.get("marca") or "").strip()
                modelo_rec=(rec_model.get("modelo") or "").strip()
                serie_rec= (rec.get("serie") or "").strip()
                codap_rec= (rec_model.get("codigo_aprobacion") or rec.get("code") or "").strip()
                origen_rec=(rec_model.get("origen") or "").strip()
                e_rec    = (rec_model.get("e") or "").strip()
                max_rec  = (rec_model.get("max") or "").strip()
                min_rec  = (rec_model.get("min") or "").strip()
                dd_dt_rec= (rec_model.get("dd_dt") or "").strip()
                clase_rec= (rec_model.get("clase") or "").strip()

                # Indicador (modeloDetalle > instrumentoDetalle)
                marca_ind = (ind_model.get("marca") or "").strip()
                modelo_ind= (ind_model.get("modelo") or "").strip()
                serie_ind = (ind.get("serie") or "").strip()
                codap_ind = (ind_model.get("codigo_aprobacion") or ind.get("code") or "").strip()
                origen_ind= (ind_model.get("origen") or "").strip()
                naprob_ind= (ind_model.get("n_aprob") or "").strip()
                faprob_ind= (ind_model.get("fecha_aprob") or "").strip()

                # Ubicación de instalación (desde instrumento)
                lugar_dom = inst.get("inst_dom", "")
                lugar_loc = inst.get("inst_loc", "")
                lugar_prov= inst.get("inst_prov", "")

                fila = {
                    # 1-10
                    "Número de O.T.": nro_ot,
                    "VPE Nº": vpe_num,
                    "Empresa solicitante": empresa,
                    "Razón social (Propietario)": usuario_rep,
                    "Domicilio (Fiscal)": direccion_legal_det,  # <- tomado de detalle.jsp
                    "Localidad (Fiscal)": "",
                    "Provincia (Fiscal)": "",
                    "Lugar propio de instalación - Domicilio": lugar_dom,
                    "Lugar propio de instalación - Localidad": lugar_loc,
                    "Lugar propio de instalación - Provincia": lugar_prov,
                    # 11
                    "Instrumento verificado": "Balanza para pesar camiones",
                    # 12-22 Receptor
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
                    # 23-30 Indicador
                    "Tipo (Indicador)": "electrónica",
                    "Marca Indicador": marca_ind,
                    "Modelo Indicador": modelo_ind,
                    "N° de serie Indicador": serie_ind,
                    "Código Aprobación (Indicador)": codap_ind,
                    "Origen Indicador": origen_ind,
                    "N° de Aprobación Modelo (Indicador)": naprob_ind,
                    "Fecha de Aprobación Modelo (Indicador)": faprob_ind,
                }

                filas.append(fila)
                print(f"✅ Instrumento {idx} procesado → Receptor: {modelo_rec or 'N/D'} | Indicador: {modelo_ind or 'N/D'}")

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
    "Tipo (Indicador)",
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
    # Asegurar las 30 columnas, en orden exacto
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

        # Cabeceras en negrita
        fmt_header = wb.add_format({"bold": True})
        for i, col in enumerate(df.columns):
            ws.write(0, i, col, fmt_header)

        # Anchos de columna aproximados
        for i, col in enumerate(df.columns):
            try:
                max_len = int(df[col].astype(str).map(len).max() if not df.empty else 10)
            except Exception:
                max_len = 10
            max_len = max(10, min(50, max_len + 2))
            ws.set_column(i, i, max_len)

        ws.freeze_panes(1, 0)

    return ruta


# =========================
# CLI
# =========================

def main():
    parser = argparse.ArgumentParser(description="Extracción MetroWeb → Excel Verificación Previa (balanzas de camiones/plataforma).")
    parser.add_argument("--user", dest="user", help="Usuario MetroWeb")
    parser.add_argument("--pass", dest="pwd", help="Contraseña MetroWeb")
    parser.add_argument("--ot", dest="ot", help="Número de OT (ej. 307-62136)")
    parser.add_argument("--headless", action="store_true", help="Ejecutar sin mostrar navegador")
    parser.add_argument("--out", dest="out", help="Ruta de salida .xlsx (opcional)")

    args = parser.parse_args()

    user = args.user or input("Usuario MetroWeb: ").strip()
    pwd = args.pwd or getpass.getpass("Contraseña MetroWeb: ").strip()
    ot  = args.ot or input("Número de OT (ej. 307-62136): ").strip()
    headless = args.headless
    out = args.out

    if not re.match(r"^\d{3}-\d{5}$", ot):
        print("⚠️ Formato de OT esperado: 307-XXXXX (5 dígitos).")
        # no aborta: permite continuar igualmente

    if not out:
        out = f"OT_{ot}_VERIFICACION_PREVIA.xlsx"

    print(f"➡️  Iniciando extracción | OT={ot} | headless={headless}")
    filas = extraer_camiones_por_ot(ot=ot, user=user, pwd=pwd, mostrar_navegador=not headless)

    if not filas:
        print("⚠️ No se generaron filas. Verifique la OT o credenciales.")
        # Aún así crea un Excel vacío con estructura
    df = armar_hoja_verificacion(filas)
    ruta = exportar_verificacion(df, Path(out))
    print(f"✅ Archivo generado: {ruta.resolve()}")

if __name__ == "__main__":
    main()
