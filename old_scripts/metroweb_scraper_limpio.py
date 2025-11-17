# metroweb_scraper_limpio.py
# -*- coding: utf-8 -*-
# pylint: disable=too-many-statements, too-many-branches, too-many-locals, line-too-long

"""
Scraper de MetroWeb (INTI) con Playwright.
- Extrae cabecera del VPE (detalle.jsp).
- Lee "Usuario Representado" y cabecera (resumen.jsp).
- Recorre instrumentos desde inputs ocultos (instrumentoDetalle.do?idInstrumento=...).
- Abre detalle del modelo (modeloDetalle.do) para capacidades, clase, N° de disposición, etc.
- Exporta Excel con detalle y una hoja de resumen por modelo.
"""

from __future__ import annotations

import os
import re
from typing import Dict, Tuple, Optional, List

# Stdlib primero
import tkinter as tk
from tkinter import filedialog, simpledialog, messagebox

# Terceros
import pandas as pd
from playwright.sync_api import sync_playwright

# Constantes
DEFAULT_WAIT_MS = 60_000
DEFAULT_SLOW_MO_MS = 0  # subir a 300 si querés ver paso a paso


# =========================
# Helpers de scraping
# =========================

def td_value(page, label: str, keep_newlines: bool = False) -> str:
    """Devuelve el texto de la <td> siguiente a la celda cuya etiqueta CONTENGA 'label'."""
    loc = page.locator(
        f"xpath=//td[contains(normalize-space(.), '{label}')]/following-sibling::td[1]"
    )
    if not loc.count():
        return ""
    txt = loc.first.inner_text()
    if not keep_newlines:
        txt = txt.replace("\r", "").replace("\n", " ")
    return txt.strip()


def td_value_exact(page, labels: List[str], keep_newlines: bool = False) -> str:
    """Coincidencia EXACTA de la etiqueta. Útil para 'Máximo', 'Mínimo', 'e', etc."""
    for lbl in labels:
        loc = page.locator(
            f"xpath=//td[normalize-space(.)='{lbl}']/following-sibling::td[1]"
        )
        if loc.count():
            txt = loc.first.inner_text()
            if not keep_newlines:
                txt = txt.replace("\r", "").replace("\n", " ")
            return txt.strip()
    return ""


def td_value_any(page, labels: List[str], keep_newlines: bool = False) -> str:
    """Intenta múltiples variantes/sinónimos de etiqueta."""
    for lbl in labels:
        v = td_value(page, lbl, keep_newlines=keep_newlines)
        if v:
            return v
    return ""


def split_domicilio(domicilio_multiline: str) -> Tuple[str, str, str]:
    """Separa domicilio multilínea (3 líneas) → (domicilio, localidad, provincia)."""
    if not domicilio_multiline:
        return ("", "", "")
    lines = [ln.strip(" \t\r") for ln in domicilio_multiline.replace("\r", "").split("\n")]
    lines = [ln for ln in lines if ln.strip()]
    dom = lines[0] if len(lines) > 0 else ""
    loc = lines[1] if len(lines) > 1 else ""
    prov = lines[2] if len(lines) > 2 else ""
    return (dom, loc, prov)


def parse_modelo_tipo_marca_fabricante(texto: str) -> Dict[str, str]:
    """
    'MINI - Balanza - TREBOL - BASCULAS TREBOL SAIC'
      → {'Modelo': 'MINI', 'Balanza tipo 1': 'Balanza', 'Marca': 'TREBOL', 'Fabricante/Importador': 'BASCULAS TREBOL SAIC'}
    """
    out = {"Modelo": "", "Balanza tipo 1": "", "Marca": "", "Fabricante/Importador": ""}
    if not texto:
        return out
    partes = [p.strip() for p in texto.split(" - ") if p.strip()]
    if len(partes) >= 1:
        out["Modelo"] = partes[0]
    if len(partes) >= 2:
        out["Balanza tipo 1"] = partes[1]
    if len(partes) >= 3:
        out["Marca"] = partes[2]
    if len(partes) >= 4:
        out["Fabricante/Importador"] = " - ".join(partes[3:])
    return out


def _resumen_header_table(page):
    """Tabla inmediatamente debajo del texto 'Detalle Historia Adjuntos'."""
    root = page.locator(
        "xpath=//td[contains(normalize-space(.),'Detalle Historia Adjuntos')]/ancestor::table[1]/following-sibling::table[1]"
    )
    return root if root.count() else page


def _td_value_in(root, labels):
    """Busca dentro de 'root' (no en toda la página) y toma el primer match válido."""
    for lbl in labels:
        loc = root.locator(
            f"xpath=.//td[.//strong[contains(normalize-space(.), '{lbl}')]]/following-sibling::td[1]"
        )
        if not loc.count():
            loc = root.locator(
                f"xpath=.//td[contains(normalize-space(.), '{lbl}')]/following-sibling::td[1]"
            )
        if loc.count():
            txt = loc.first.inner_text().replace("\r", "").replace("\n", " ").strip()
            if txt:
                return txt
    return ""


# =========================
# Lectores de páginas
# =========================

def leer_vpe_cabecera(page) -> Dict[str, str]:
    """
    Lee la sección 'Usuario del equipo' del detalle del VPE (detalle.jsp).
    Usa etiquetas y, si falta, un fallback estructural por filas.
    """
    v = {
        "Nombre del Usuario del instrumento": td_value_any(
            page,
            [
                "Nombre del Usuario del instrumento",
                "Nombre del Usuario",
                "Usuario del instrumento",
                "Usuario del equipo",
            ],
        ),
        "Nº de CUIT": td_value_any(
            page,
            ["Nº de CUIT", "N° de CUIT", "Nro de CUIT", "N&ordm; de CUIT", "CUIT"],
        ),
        "Dirección Legal": td_value_any(
            page,
            [
                "Dirección Legal",
                "Direcci\u00f3n Legal",
                "Direcci&oacute;n Legal",
                "Domicilio Legal",
            ],
        ),
    }

    # Fallback estructural si algo quedó vacío (la tabla inmediata a "Usuario del equipo")
    if any(not v[k] for k in v):
        sec = page.locator(
            "xpath=//td[contains(@class,'screenform')]//*[contains(normalize-space(.),'Usuario del equipo')]"
        )
        if sec.count():
            tabla = sec.locator("xpath=../../following-sibling::tr[1]//table").first
            filas = tabla.locator("xpath=.//tr")
            n = filas.count()
            for i in range(n):
                celdas = filas.nth(i).locator("xpath=./td")
                if celdas.count() < 2:
                    continue
                etiqueta = (
                    celdas.nth(0)
                    .inner_text()
                    .replace("\r", "")
                    .replace("\n", " ")
                    .strip()
                    .lower()
                )
                valor = (
                    celdas.nth(1)
                    .inner_text()
                    .replace("\r", "")
                    .replace("\n", " ")
                    .strip()
                )

                if (
                    "nombre del usuario" in etiqueta
                    or "usuario del equipo" in etiqueta
                ) and not v["Nombre del Usuario del instrumento"]:
                    v["Nombre del Usuario del instrumento"] = valor
                elif "cuit" in etiqueta and not v["Nº de CUIT"]:
                    v["Nº de CUIT"] = valor
                elif (
                    "dirección legal" in etiqueta
                    or "direccion legal" in etiqueta
                    or "domicilio legal" in etiqueta
                ) and not v["Dirección Legal"]:
                    v["Dirección Legal"] = valor

    return v


def leer_vpe_resumen(context, page_detalle, wait_ms: int = DEFAULT_WAIT_MS) -> Dict[str, str]:
    """
    Abre resumen.jsp y retorna campos útiles.
    La lectura de 'Número/Estado/Nro OT/Inicio/Aceptado' se hace
    restringida a la caja 'Detalle Historia Adjuntos' para evitar choques con el menú 'Inicio'.
    """
    out: Dict[str, str] = {}
    link = page_detalle.locator("a[href*='tramiteVPE.do?acc=resumen']").first
    if not link.count():
        return out

    href = link.get_attribute("href") or ""
    if not href:
        return out

    p = context.new_page()
    try:
        p.goto(
            href if href.startswith("http") else "https://app.inti.gob.ar" + href,
            timeout=wait_ms,
        )
        p.wait_for_load_state("networkidle", timeout=wait_ms)

        header = _resumen_header_table(p)
        out["Número VPE"] = _td_value_in(header, ["Número", "N\u00famero"])
        out["Estado"] = _td_value_in(header, ["Estado"])
        out["Nro OT (encabezado)"] = _td_value_in(header, ["Nro OT", "N° OT", "Nº OT"])
        out["Inicio"] = _td_value_in(header, ["Inicio"])
        out["Aceptado"] = _td_value_in(header, ["Aceptado"])

        out["Empresa Solicitante"] = td_value_any(p, ["Empresa Solicitante"])
        out["Usuario Representado"] = td_value_any(
            p, ["Usuario Representado", "Usuario  Representado"]
        )

    finally:
        try:
            p.close()
        except Exception:  # pylint: disable=broad-exception-caught
            pass

    return out


def asegurar_detalle_vpe(page, vpe: str, wait_ms: int):
    """
    Si el click al VPE nos dejó en resumen.jsp (o acc=resumen), navega al detalle.jsp (acc=detalle).
    Usa el idTramite derivado del texto 'vpeNNNNN' si hace falta.
    NO espera anchors (en detalle no hay). Solo navega.
    """
    # ¿Ya estamos en detalle? (aparece la sección 'Usuario del equipo' en detalle)
    if page.locator(
        "xpath=//td[contains(@class,'screenform')]//*[contains(normalize-space(.),'Usuario del equipo')]"
    ).count():
        return

    url_actual = page.url
    url_detalle = url_actual.replace("acc=resumen", "acc=detalle").replace(
        "/resumen.jsp", "/detalle.jsp"
    )

    if url_detalle != url_actual:
        page.goto(url_detalle, timeout=wait_ms)
    else:
        # Derivar idTramite desde 'vpe12345'
        m_obj = re.search(r"vpe\s*?(\d+)", vpe, flags=re.I)
        if m_obj:
            idt = m_obj.group(1)
            url_detalle = (
                "https://app.inti.gob.ar/MetroWeb/tramiteVPE.do"
                f"?acc=detalle&idTramite={idt}"
            )
            page.goto(url_detalle, timeout=wait_ms)
        else:
            # Último intento: ir directo a /pages/tramiteVPE/detalle.jsp
            page.goto(
                "https://app.inti.gob.ar/MetroWeb/pages/tramiteVPE/detalle.jsp",
                timeout=wait_ms,
            )


def leer_instrumento(page) -> Dict[str, str]:
    """Lee los datos visibles en la tabla del Instrumento dentro del detalle del VPE."""
    datos: Dict[str, str] = {}

    # Código Aprobación de Modelo
    datos["Código Ap. de Modelo"] = td_value_any(
        page,
        [
            "Código de Aprobación de Modelo",
            "C\u00f3digo de Aprobaci\u00f3n de Modelo",
            "Codigo Aprobación",
        ],
    )

    # Línea combinada: Modelo - Tipo - Marca - Fabricante
    linea_modelo = td_value_any(
        page, ["Modelo - Tipo de Instrumento - Marca - Fabricante", "Modelo - Tipo - Marca - Fabricante"]
    )
    parsed = parse_modelo_tipo_marca_fabricante(linea_modelo)
    datos.update(parsed)

    # N° de Serie
    datos["N° de Serie"] = td_value_any(
        page, ["Nro de serie", "N° de serie", "Número de serie", "Nº de serie", "N&ordm; de serie"]
    )

    # N° Verificación (OT)
    datos["N° Verificación (OT)"] = td_value_any(
        page, ["Nro Verificación (OT)", "N° Verificación (OT)", "Nº Verificación (OT)", "Nro Verificaci\u00f3n (OT)"]
    )

    # Domicilio multilínea (desde página de instrumento o cabecera del VPE)
    domicilio_full = td_value_any(
        page, ["Domicilio", "Domicilio donde están", "Domicilio donde est\u00e1n"], keep_newlines=True
    )
    dom, loc, prov = split_domicilio(domicilio_full)
    datos["Domicilio"] = dom
    datos["Localidad instalación"] = loc
    datos["Provincia"] = prov

    # Fechas
    datos["Fecha verificación"] = td_value_any(
        page,
        [
            "Fecha última Verificación",
            "Fecha &uacute;ltima Verificaci\u00f3n",
            "Fecha de Verificación",
            "Fecha verificaci\u00f3n",
        ],
    )
    datos["Fecha de Precintos"] = td_value_any(page, ["Fecha de Precintos"])

    # Default
    datos.setdefault("Clase", "III")
    return datos


def leer_modelo(context, page, modelos_cache: Dict[str, Dict[str, str]]) -> Dict[str, str]:
    """Lee datos del modelo desde 'modeloDetalle.do' (si está) y los cachea."""
    # Clave de cache
    codigo = td_value_any(
        page,
        [
            "Código de Aprobación de Modelo",
            "C\u00f3digo de Aprobaci\u00f3n de Modelo",
            "Codigo Aprobación",
        ],
    )
    linea_modelo = td_value_any(
        page, ["Modelo - Tipo de Instrumento - Marca - Fabricante", "Modelo - Tipo - Marca - Fabricante"]
    )
    parsed = parse_modelo_tipo_marca_fabricante(linea_modelo)
    cache_key = f"{codigo}|{parsed.get('Modelo','')}|{parsed.get('Marca','')}"

    if cache_key in modelos_cache:
        return modelos_cache[cache_key]

    datos = {
        "Balanza tipo 1": parsed.get("Balanza tipo 1", ""),
        "Capacidad Máx.": "",
        "Capacidad Mín.": "",
        "División e": "",
        "Clase": "III",
        "Fabricante/Importador": parsed.get("Fabricante/Importador", ""),
        "Marca": parsed.get("Marca", ""),
        "Modelo": parsed.get("Modelo", ""),
        "Código Ap. de Modelo": codigo or "",
        "Origen": "",
        "N° Aprob. Modelo": "",
        "Fecha Aprob. Modelo": "",
    }

    # Intentar abrir detalle del modelo
    link = page.locator('a[href*="modeloDetalle.do"]').first
    if not link.count():
        link = page.locator('a[href*="modelo"]').first

    try:
        if link.count():
            href = link.get_attribute("href")
            if href:
                p2 = context.new_page()
                try:
                    p2.goto(
                        href if href.startswith("http") else "https://app.inti.gob.ar" + href,
                        timeout=DEFAULT_WAIT_MS,
                    )
                    p2.wait_for_load_state("networkidle", timeout=DEFAULT_WAIT_MS)

                    # Etiquetas exactas en 'modeloDetalle'
                    datos["Capacidad Máx."] = td_value_exact(p2, ["Máximo", "Maximo"]) or datos["Capacidad Máx."]
                    datos["Capacidad Mín."] = td_value_exact(p2, ["Mínimo", "Minimo"]) or datos["Capacidad Mín."]
                    datos["División e"] = td_value_exact(p2, ["e"]) or datos["División e"]
                    datos["Clase"] = td_value_exact(p2, ["Clase"]) or datos["Clase"]

                    datos["Origen"] = td_value_any(
                        p2,
                        ["País Origen", "Pais Origen", "País de origen", "Pa\u00eds Origen"],
                    ) or datos["Origen"]

                    # N° Aprob. Modelo aparece como 'Nº Disposición'
                    datos["N° Aprob. Modelo"] = td_value_any(
                        p2,
                        [
                            "Nº Disposición",
                            "N° Disposición",
                            "Nº Disposicion",
                            "N° Disposicion",
                            "Nro Disposición",
                            "Nro Disposicion",
                        ],
                    ) or datos["N° Aprob. Modelo"]

                    datos["Fecha Aprob. Modelo"] = td_value_any(
                        p2, ["Fecha Aprobación", "Fecha de Aprobación", "Fecha Aprob. Modelo"]
                    ) or datos["Fecha Aprob. Modelo"]

                finally:
                    try:
                        p2.close()
                    except Exception:  # pylint: disable=broad-exception-caught
                        pass
    except Exception:  # pylint: disable=broad-exception-caught
        pass

    modelos_cache[cache_key] = datos
    return datos


# =========================
# Excel helpers
# =========================

def _formatear_hoja_excel(ws, df: pd.DataFrame, writer, wrap_cols: List[str], default_width: int = 22):
    workbook = writer.book
    wrap = workbook.add_format({"text_wrap": True, "valign": "top"})
    header = workbook.add_format({"bold": True, "bg_color": "#E9EEF7", "border": 1})

    # Encabezados
    for colnum, colname in enumerate(df.columns):
        ws.write(0, colnum, colname, header)

    # Filtro y freeze
    ws.autofilter(0, 0, max(0, len(df.index)), max(0, len(df.columns) - 1))
    ws.freeze_panes(1, 0)

    # Anchos
    for colnum, colname in enumerate(df.columns):
        if colname in wrap_cols:
            ws.set_column(colnum, colnum, 35, cell_format=wrap)
        else:
            ws.set_column(colnum, colnum, default_width)


def guardar_excel(df_detalle: pd.DataFrame, df_resumen: pd.DataFrame, ruta_salida: str) -> str:
    """Crea un Excel con 'Detalle' y 'Resumen' y formatea columnas largas."""
    if not ruta_salida.lower().endswith(".xlsx"):
        ruta_salida += ".xlsx"

    with pd.ExcelWriter(ruta_salida, engine="xlsxwriter") as writer:
        # Detalle
        df_detalle.to_excel(writer, sheet_name="Detalle", index=False)
        ws1 = writer.sheets["Detalle"]
        _formatear_hoja_excel(ws1, df_detalle, writer, wrap_cols=["Domicilio", "Dirección Legal", "Números_de_serie"])

        # Resumen
        df_resumen.to_excel(writer, sheet_name="Resumen", index=False)
        ws2 = writer.sheets["Resumen"]
        _formatear_hoja_excel(ws2, df_resumen, writer, wrap_cols=["Domicilio", "Dirección Legal", "Números_de_serie"])

    return ruta_salida


# =========================
# GUI
# =========================

def pedir_ot_y_destino(pred_ot: str = "") -> Tuple[Optional[str], Optional[str]]:
    root = tk.Tk()
    root.withdraw()

    ot = simpledialog.askstring("INTI | OT a consultar", "Ingresá el número de OT:", initialvalue=pred_ot)
    if not ot:
        return None, None

    ruta = filedialog.asksaveasfilename(
        title="Guardar Excel",
        defaultextension=".xlsx",
        filetypes=[("Excel", "*.xlsx")],
        initialfile=f"OT_{ot}_metroweb.xlsx",
    )
    if not ruta:
        return None, None

    return ot.strip(), ruta


# =========================
# Flujo principal
# =========================

def extraer_datos_metroweb(
    ot_numero: str, usuario: str, password: str, mostrar_navegador: bool = True
) -> pd.DataFrame:
    filas: List[Dict[str, str]] = []
    domicilio_cache: Optional[Tuple[str, str, str]] = None
    modelos_cache: Dict[str, Dict[str, str]] = {}

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=not mostrar_navegador, slow_mo=DEFAULT_SLOW_MO_MS)
        context = browser.new_context()
        page = context.new_page()

        try:
            # --- LOGIN ---
            page.goto("https://app.inti.gob.ar/MetroWeb/pages/ingreso.jsp", timeout=DEFAULT_WAIT_MS)
            page.wait_for_load_state("domcontentloaded", timeout=DEFAULT_WAIT_MS)

            # Campos usuario/clave tolerantes
            campo_usuario = page.locator('input[name="usuario"]')
            if not campo_usuario.count():
                campo_usuario = page.locator("input#usuario")
            if not campo_usuario.count():
                campo_usuario = page.locator('input[type="text"]').first

            campo_password = page.locator('input[name="contrasena"]')
            if not campo_password.count():
                campo_password = page.locator('input[name="password"]')
            if not campo_password.count():
                campo_password = page.locator('input[type="password"]').first

            if not (campo_usuario and campo_password and campo_usuario.count() and campo_password.count()):
                raise RuntimeError("No se encontraron los campos de login.")

            campo_usuario.fill(usuario)
            campo_password.fill(password)

            if page.locator('input[value="Ingresar"]').count():
                page.click('input[value="Ingresar"]')
            elif page.locator('input[type="submit"]').count():
                page.click('input[type="submit"]')
            else:
                page.keyboard.press("Enter")

            page.wait_for_load_state("networkidle", timeout=DEFAULT_WAIT_MS)

            # --- BUSCAR OT ---
            page.goto("https://app.inti.gob.ar/MetroWeb/entrarPML.do", timeout=DEFAULT_WAIT_MS)
            page.wait_for_selector('input[name="numeroOT"]', timeout=DEFAULT_WAIT_MS)
            page.fill('input[name="numeroOT"]', ot_numero)

            if page.locator('input[value="Buscar"]').count():
                page.click('input[value="Buscar"]')
            else:
                page.keyboard.press("Enter")

            page.wait_for_load_state("networkidle", timeout=DEFAULT_WAIT_MS)

            # --- ABRIR VPE ---
            link_vpe = page.locator('a[href*="tramiteVPE"]').first
            if not link_vpe or not link_vpe.count():
                raise RuntimeError(f"No se encontró el trámite VPE para la OT {ot_numero}.")
            vpe = link_vpe.inner_text().strip()
            link_vpe.click()

            page.wait_for_load_state("networkidle", timeout=DEFAULT_WAIT_MS)

            # Si caímos en RESUMEN, redirigir a DETALLE
            asegurar_detalle_vpe(page, vpe, DEFAULT_WAIT_MS)

            # Asegurar que cargó 'Usuario del equipo' (detalle.jsp)
            page.wait_for_selector(
                "xpath=//td[contains(@class,'screenform')]//*[contains(normalize-space(.),'Usuario del equipo')]",
                timeout=DEFAULT_WAIT_MS,
            )

            # --- CABECERA VPE (detalle.jsp) ---
            datos_vpe = leer_vpe_cabecera(page)

            # --- RESUMEN (resumen.jsp) ---
            datos_resumen = leer_vpe_resumen(context, page, wait_ms=DEFAULT_WAIT_MS)

            # Completar nombre con 'Usuario Representado' si faltó
            if not datos_vpe.get("Nombre del Usuario del instrumento") and datos_resumen.get("Usuario Representado"):
                datos_vpe["Nombre del Usuario del instrumento"] = datos_resumen["Usuario Representado"]

            # Domicilio cabecera (fallback para instrumentos sin domicilio)
            dom_cab, loc_cab, prov_cab = split_domicilio(
                td_value_any(page, ["Domicilio donde están", "Domicilio donde est\u00e1n"], keep_newlines=True)
            )

            # --- LISTA DE INSTRUMENTOS (inputs ocultos + fallback regex) ---
            ids: List[str] = []
            try:
                # Intento normal por selector
                page.wait_for_selector('input[name^="instrumentos"][name$=".idInstrumento"]', timeout=15_000)
                ids = page.eval_on_selector_all(
                    'input[name^="instrumentos"][name$=".idInstrumento"]',
                    "els => els.map(e => e.value).filter(Boolean)",
                )
            except Exception:  # pylint: disable=broad-exception-caught
                ids = []

            if not ids:
                # Fallback robusto: parsear el HTML
                html = page.content()
                ids = re.findall(r'name="instrumentos\[\d+\]\.idInstrumento"\s+value="(\d+)"', html)

            if not ids:
                raise RuntimeError("No se encontraron instrumentos en el detalle del VPE.")

            for i, id_inst in enumerate(ids, start=1):
                href = f"https://app.inti.gob.ar/MetroWeb/instrumentoDetalle.do?idInstrumento={id_inst}"

                detalle = context.new_page()
                try:
                    detalle.goto(href, timeout=DEFAULT_WAIT_MS)
                    detalle.wait_for_load_state("networkidle", timeout=DEFAULT_WAIT_MS)

                    # Instrumento
                    di = leer_instrumento(detalle)

                    # Cache domicilio (primera vez)
                    if domicilio_cache is None and (
                        di.get("Domicilio") or di.get("Localidad instalación") or di.get("Provincia")
                    ):
                        domicilio_cache = (
                            di.get("Domicilio", ""),
                            di.get("Localidad instalación", ""),
                            di.get("Provincia", ""),
                        )

                    # Fallbacks de domicilio
                    if domicilio_cache is not None:
                        di["Domicilio"] = di.get("Domicilio") or domicilio_cache[0]
                        di["Localidad instalación"] = di.get("Localidad instalación") or domicilio_cache[1]
                        di["Provincia"] = di.get("Provincia") or domicilio_cache[2]

                    if dom_cab or loc_cab or prov_cab:
                        di["Domicilio"] = di.get("Domicilio") or dom_cab
                        di["Localidad instalación"] = di.get("Localidad instalación") or loc_cab
                        di["Provincia"] = di.get("Provincia") or prov_cab

                    # Modelo (con cache)
                    dm = leer_modelo(context, detalle, modelos_cache)

                    fila = {
                        "OT": ot_numero,
                        "VPE": vpe,
                        # Cabecera del VPE
                        "Nombre del Usuario del instrumento": datos_vpe.get("Nombre del Usuario del instrumento", ""),
                        "Usuario Representado": datos_resumen.get("Usuario Representado", ""),
                        "Nº de CUIT": datos_vpe.get("Nº de CUIT", ""),
                        "Dirección Legal": datos_vpe.get("Dirección Legal", ""),
                        # Instrumento
                        "Fecha verificación": di.get("Fecha verificación", ""),
                        "Domicilio": di.get("Domicilio", ""),
                        "Localidad instalación": di.get("Localidad instalación", ""),
                        "Provincia": di.get("Provincia", ""),
                        # Modelo
                        "Balanza tipo 1": dm.get("Balanza tipo 1", ""),
                        "Capacidad Máx.": dm.get("Capacidad Máx.", ""),
                        "Capacidad Mín.": dm.get("Capacidad Mín.", ""),
                        "División e": dm.get("División e", ""),
                        "Clase": dm.get("Clase", "III"),
                        "Fabricante/Importador": dm.get("Fabricante/Importador", ""),
                        "Marca": dm.get("Marca", ""),
                        "Modelo": dm.get("Modelo", ""),
                        "Código Ap. de Modelo": di.get("Código Ap. de Modelo", "") or dm.get("Código Ap. de Modelo", ""),
                        "Origen": dm.get("Origen", ""),
                        "N° Aprob. Modelo": dm.get("N° Aprob. Modelo", ""),
                        "Fecha Aprob. Modelo": dm.get("Fecha Aprob. Modelo", ""),
                        # Identificadores
                        "N° de Serie": di.get("N° de Serie", ""),
                        "N° Verificación (OT)": di.get("N° Verificación (OT)", ""),
                    }

                    filas.append(fila)
                    print(f"✅ Instrumento {i} capturado ({fila.get('Modelo') or 'Modelo N/D'})")

                finally:
                    try:
                        detalle.close()
                    except Exception:  # pylint: disable=broad-exception-caught
                        pass

        except Exception as exc:  # pylint: disable=broad-exception-caught
            root = tk.Tk()
            root.withdraw()
            messagebox.showerror("INTI | Error", str(exc))
            return pd.DataFrame()
        finally:
            try:
                browser.close()
            except Exception:  # pylint: disable=broad-exception-caught
                pass

    # Orden y columnas
    cols = [
        "OT",
        "VPE",
        "Nombre del Usuario del instrumento",
        "Usuario Representado",
        "Nº de CUIT",
        "Dirección Legal",
        "Fecha verificación",
        "Domicilio",
        "Localidad instalación",
        "Provincia",
        "Balanza tipo 1",
        "Capacidad Máx.",
        "Capacidad Mín.",
        "División e",
        "Clase",
        "Fabricante/Importador",
        "Marca",
        "Modelo",
        "Código Ap. de Modelo",
        "Origen",
        "N° Aprob. Modelo",
        "Fecha Aprob. Modelo",
        "N° de Serie",
        "N° Verificación (OT)",
    ]

    df = pd.DataFrame(filas)
    for c in cols:
        if c not in df.columns:
            df[c] = ""
    return df[cols]


def resumir_por_modelo(df: pd.DataFrame, juntar_con_saltos: bool = True) -> pd.DataFrame:
    """Resumen agrupado por características del modelo."""
    if df.empty:
        return df

    keys = [
        "Fabricante/Importador",
        "Marca",
        "Modelo",
        "Código Ap. de Modelo",
        "Origen",
        "N° Aprob. Modelo",
        "Fecha Aprob. Modelo",
        "Balanza tipo 1",
        "Capacidad Máx.",
        "Capacidad Mín.",
        "División e",
        "Clase",
    ]

    for c in keys + [
        "OT",
        "VPE",
        "Domicilio",
        "Localidad instalación",
        "Provincia",
        "N° de Serie",
        "Nombre del Usuario del instrumento",
        "Usuario Representado",
        "Nº de CUIT",
        "Dirección Legal",
    ]:
        if c not in df.columns:
            df[c] = ""
        df[c] = df[c].fillna("")

    sep = "\n" if juntar_con_saltos else " - "

    def _concat_series(series_):
        vistos, out = set(), []
        for x in series_:
            s = str(x).strip()
            if s and s not in vistos:
                vistos.add(s)
                out.append(s)
        return sep.join(out)

    g = df.sort_values(keys + ["N° de Serie"]).groupby(keys, dropna=False)
    resumen = g.agg(
        Cantidad=("N° de Serie", "size"),
        **{"Números_de_serie": ("N° de Serie", _concat_series)},
        OT=("OT", "first"),
        VPE=("VPE", "first"),
        **{"Nombre del Usuario del instrumento": ("Nombre del Usuario del instrumento", "first")},
        **{"Usuario Representado": ("Usuario Representado", "first")},
        **{"Nº de CUIT": ("Nº de CUIT", "first")},
        **{"Dirección Legal": ("Dirección Legal", "first")},
        Domicilio=("Domicilio", "first"),
        **{"Localidad instalación": ("Localidad instalación", "first")},
        Provincia=("Provincia", "first"),
    ).reset_index()

    cols = [
        "OT",
        "VPE",
        "Nombre del Usuario del instrumento",
        "Usuario Representado",
        "Nº de CUIT",
        "Dirección Legal",
        "Domicilio",
        "Localidad instalación",
        "Provincia",
        "Fabricante/Importador",
        "Marca",
        "Modelo",
        "Código Ap. de Modelo",
        "Origen",
        "N° Aprob. Modelo",
        "Fecha Aprob. Modelo",
        "Balanza tipo 1",
        "Capacidad Máx.",
        "Capacidad Mín.",
        "División e",
        "Clase",
        "Cantidad",
        "Números_de_serie",
    ]
    for c in cols:
        if c not in resumen.columns:
            resumen[c] = ""
    return resumen[cols]


# =========================
# Main
# =========================

def main():
    root = tk.Tk()
    root.withdraw()
    usuario = os.getenv("METROWEB_USER") or simpledialog.askstring("INTI | Usuario", "Usuario:")
    password = os.getenv("METROWEB_PASS") or simpledialog.askstring("INTI | Contraseña", "Contraseña:", show="*")
    if not usuario or not password:
        messagebox.showwarning("INTI", "Se canceló el ingreso de credenciales.")
        return

    ot, ruta_excel = pedir_ot_y_destino(pred_ot="")
    if not ot or not ruta_excel:
        messagebox.showwarning("INTI", "Operación cancelada por el usuario.")
        return

    df_detalle = extraer_datos_metroweb(ot, usuario, password, mostrar_navegador=True)
    if df_detalle.empty:
        return

    df_resumen = resumir_por_modelo(df_detalle, juntar_con_saltos=True)

    try:
        ruta_final = guardar_excel(df_detalle, df_resumen, ruta_excel)
        messagebox.showinfo("INTI | Listo", f"Se creó el Excel con éxito.\n\nArchivo:\n{ruta_final}")
    except Exception as exc:  # pylint: disable=broad-exception-caught
        messagebox.showerror("INTI | Error guardando", f"Ocurrió un error al guardar el Excel:\n{exc}")
        raise


if __name__ == "__main__":
    main()
