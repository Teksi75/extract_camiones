# -*- coding: utf-8 -*-
"""
Extracción MetroWeb (balanzas de camiones/plataforma) → hoja de verificación XLSX.

Requisitos (una sola vez):
    pip install playwright pandas xlsxwriter
    python -m playwright install chromium

Uso:
    python extract_camiones.py               # pide user/pass/ot
    python extract_camiones.py --ot 307-62136 --headless
    python extract_camiones.py --user USR --pass PWD --ot 307-62136 --out salida.xlsx
"""

from __future__ import annotations
import argparse
import getpass
import re
from typing import Dict, List, Tuple, Optional
import pandas as pd
from playwright.sync_api import sync_playwright

DEFAULT_WAIT_MS = 60_000
DEFAULT_SLOW_MO_MS = 0


# ==============================
# Helpers
# ==============================
def _clean(s: Optional[str]) -> str:
    if s is None:
        return ""
    return str(s).replace("\r", " ").replace("\n", " ").strip()


def _solo_digitos(s: str) -> str:
    return "".join(re.findall(r"\d+", s or ""))


def td_after_exact(page, label: str, keep_newlines: bool = False) -> str:
    """Valor de la celda a la derecha de un TD cuyo texto coincida EXACTO con label."""
    xp = f"xpath=//td[normalize-space(.)='{label}']/following-sibling::td[1]"
    loc = page.locator(xp)
    if loc.count():
        txt = loc.first.inner_text()
        if not keep_newlines:
            txt = txt.replace("\r", " ").replace("\n", " ")
        return txt.strip()
    return ""


def td_after_any(page, labels: List[str], keep_newlines: bool = False) -> str:
    for lbl in labels:
        val = td_after_exact(page, lbl, keep_newlines=keep_newlines)
        if val:
            return val
    return ""


def td_exact(page, lbls: List[str]) -> str:
    for lbl in lbls:
        loc = page.locator(f"xpath=//td[normalize-space(.)='{lbl}']/following-sibling::td[1]")
        if loc.count():
            return _clean(loc.first.inner_text())
    return ""


def _split_domicilio_3lineas(txt: str) -> Tuple[str, str, str]:
    if not txt:
        return ("", "", "")
    lines = [l.strip() for l in txt.replace("\r", "").split("\n") if l.strip()]
    dom = lines[0] if len(lines) > 0 else ""
    loc = lines[1] if len(lines) > 1 else ""
    prov = lines[2] if len(lines) > 2 else ""
    return (dom, loc, prov)


def _parse_modelo_tipo_marca_fabricante(texto: str) -> Dict[str, str]:
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
        out["Fabricante/Importador"] = partes[3]
    return out


def _sec_row_value(page, section_title: str, row_label_exact: str) -> str:
    xp = (
        f"xpath=//td[normalize-space(.)='{section_title}']"
        "/ancestor::table[1]//tr[td[1][normalize-space(.)='{row_label_exact}']]/td[2]"
    )
    loc = page.locator(xp)
    return _clean(loc.first.inner_text()) if loc.count() else ""


def _sec_row_link(page, section_title: str, row_label_exact: str) -> str:
    xp = (
        f"xpath=//td[normalize-space(.)='{section_title}']"
        "/ancestor::table[1]//tr[td[1][normalize-space(.)='{row_label_exact}']]/td[2]//a"
    )
    a = page.locator(xp)
    if a.count():
        href = a.first.get_attribute("href") or ""
        if href and not href.startswith("http"):
            href = "https://app.inti.gob.ar" + href
        return href
    return ""


# ==============================
# 1-2) Login y abrir OT/VPE
# ==============================
def login_y_abrir_ot(context, usuario: str, password: str, ot_numero: str):
    page = context.new_page()
    page.goto("https://app.inti.gob.ar/MetroWeb/pages/ingreso.jsp", timeout=DEFAULT_WAIT_MS)
    page.wait_for_load_state("domcontentloaded", timeout=DEFAULT_WAIT_MS)

    usr = page.locator('input[name="usuario"]') or page.locator("#usuario")
    if not usr.count():
        usr = page.locator('input[type="text"]').first
    pwd = page.locator('input[name="contrasena"]') or page.locator('input[name="password"]')
    if not pwd.count():
        pwd = page.locator('input[type="password"]').first

    usr.fill(usuario)
    pwd.fill(password)
    if page.locator('input[value="Ingresar"]').count():
        page.click('input[value="Ingresar"]')
    else:
        page.keyboard.press("Enter")
    page.wait_for_load_state("networkidle", timeout=DEFAULT_WAIT_MS)

    page.goto("https://app.inti.gob.ar/MetroWeb/entrarPML.do", timeout=DEFAULT_WAIT_MS)
    page.wait_for_selector('input[name="numeroOT"]', timeout=DEFAULT_WAIT_MS)
    page.fill('input[name="numeroOT"]', ot_numero)
    if page.locator('input[value="Buscar"]').count():
        page.click('input[value="Buscar"]')
    else:
        page.keyboard.press("Enter")
    page.wait_for_load_state("networkidle", timeout=DEFAULT_WAIT_MS)

    link_vpe = page.locator('a[href*="tramiteVPE"]').first
    if not link_vpe.count():
        raise RuntimeError(f"No se encontró VPE para la OT {ot_numero}.")
    vpe_text = _clean(link_vpe.inner_text())
    link_vpe.click()
    page.wait_for_load_state("networkidle", timeout=DEFAULT_WAIT_MS)

    if "acc=resumen" in page.url or page.url.endswith("/resumen.jsp"):
        page.goto(
            page.url.replace("acc=resumen", "acc=detalle").replace("/resumen.jsp", "/detalle.jsp"),
            timeout=DEFAULT_WAIT_MS,
        )

    return page, vpe_text


# ==============================
# 3) Datos generales — ***ESENCIAL: detalle.jsp***
# ==============================
def _goto_resumen(page):
    if "acc=detalle" in page.url:
        page.goto(page.url.replace("acc=detalle", "acc=resumen"), timeout=DEFAULT_WAIT_MS)
    elif page.url.endswith("/detalle.jsp"):
        page.goto(page.url.replace("/detalle.jsp", "/resumen.jsp"), timeout=DEFAULT_WAIT_MS)


def leer_generales(page) -> Dict[str, str]:
    """
    Fuente principal:
      - detalle.jsp → "Nombre del Usuario del Instrumento" (Razón social Propietario)
      - detalle.jsp → "Dirección Legal" (Domicilio Fiscal)
      - detalle.jsp → "Nº de CUIT"
    Fallback:
      - resumen.jsp → "Usuario Representado" y/o "Dirección Legal" si no estuvieran.
    """
    out = {}

    # 1) detalle.jsp (PRINCIPAL)
    razon_det = td_after_any(page, [
        "Nombre del Usuario del Instrumento",  # exacto como en tu captura
        "Nombre del Usuario",                   # alias por si cambia
        "Usuario del instrumento",
        "Usuario del equipo",
    ])
    cuit_det = td_after_any(page, ["Nº de CUIT", "N° de CUIT", "Nro de CUIT", "N&ordm; de CUIT", "CUIT"])
    dir_legal_det = td_after_any(
        page,
        ["Dirección Legal", "Direcci\u00f3n Legal", "Direcci&oacute;n Legal", "Domicilio Legal"],
        keep_newlines=False,  # pedido: llevarla tal cual a "Domicilio (Fiscal)"
    )

    out["Razon_social_propietario"] = razon_det
    out["CUIT"] = cuit_det
    out["Direccion_legal"] = dir_legal_det  # *** esto va directo a "Domicilio (Fiscal)"

    # 2) Fallback a resumen.jsp solo si falta algo
    if not out["Razon_social_propietario"] or not out["Direccion_legal"] or not out["CUIT"]:
        _goto_resumen(page)
        if not out["Razon_social_propietario"]:
            out["Razon_social_propietario"] = td_after_any(page, ["Usuario Representado", "Usuario  Representado"])
        if not out["Direccion_legal"]:
            out["Direccion_legal"] = td_after_any(page, ["Dirección Legal", "Domicilio Legal"], keep_newlines=False)
        if not out["CUIT"]:
            out["CUIT"] = td_after_any(page, ["Nº de CUIT", "N° de CUIT", "CUIT"])
        # regresar a detalle
        if "acc=resumen" in page.url or page.url.endswith("/resumen.jsp"):
            page.goto(
                page.url.replace("acc=resumen", "acc=detalle").replace("/resumen.jsp", "/detalle.jsp"),
                timeout=DEFAULT_WAIT_MS,
            )

    # 3) Lugar de instalación (puede estar en 3 líneas; lo dejamos dividido)
    dom_ml = td_after_any(page, ["Domicilio donde están", "Domicilio donde est\u00e1n", "Domicilio"], keep_newlines=True)
    dom, loc, prov = _split_domicilio_3lineas(dom_ml)
    out["Instalacion_domicilio"] = dom
    out["Instalacion_localidad"] = loc
    out["Instalacion_provincia"] = prov

    out["Fecha_verificacion"] = td_after_any(
        page,
        [
            "Fecha última Verificación",
            "Fecha &uacute;ltima Verificaci\u00f3n",
            "Fecha de Verificación",
            "Fecha verificaci\u00f3n",
        ],
    )
    out["Tipo_verificacion"] = td_after_any(page, ["Tipo de Verificación", "Tipo Verificación", "Tipo verificación"])
    out["Tolerancia"] = td_after_any(page, ["Tolerancia"])

    # (Por pedido actual, NO partimos Dirección Legal en localidad/provincia fiscal)
    out["Fiscal_domicilio"] = out["Direccion_legal"]
    out["Fiscal_localidad"] = ""
    out["Fiscal_provincia"] = ""
    return out


# ==============================
# 4-5) modeloDetalle (opcional, se mantiene por si ya lo usabas)
# ==============================
def leer_modelo_detalle(context, href_modelo: str) -> Dict[str, str]:
    out = {
        "Modelo": "",
        "CodigoAprob": "",
        "Fabricante": "",
        "Marca": "",
        "PaisOrigen": "",
        "FechaAprob": "",
        "NroDisposicion": "",
        "Max": "",
        "Min": "",
        "e": "",
        "dd_dt": "",
        "Clase": "",
    }
    if not href_modelo:
        return out
    p2 = context.new_page()
    try:
        p2.goto(href_modelo, timeout=DEFAULT_WAIT_MS)
        p2.wait_for_load_state("networkidle", timeout=DEFAULT_WAIT_MS)

        out["Modelo"] = td_exact(p2, ["Modelo"])
        out["CodigoAprob"] = td_exact(p2, ["Codigo Aprobación", "Código Aprobación", "Codigo Aprobaci\u00f3n"])
        out["Fabricante"] = td_exact(p2, ["Fabricante"])
        out["Marca"] = td_exact(p2, ["Marca"])
        out["PaisOrigen"] = td_exact(p2, ["País Origen", "Pais Origen"])
        out["FechaAprob"] = td_exact(p2, ["Fecha Aprobación", "Fecha Aprobaci\u00f3n"])
        out["NroDisposicion"] = td_exact(
            p2, ["Nº Disposición", "N° Disposición", "Nro Disposición", "Nº Disposicion", "N° Disposicion"]
        )

        out["Max"] = td_exact(p2, ["Máximo", "Máximo (Máximo)", "Maximo"])
        out["Min"] = td_exact(p2, ["Mínimo", "Minimo"])
        out["e"] = td_exact(p2, ["e"])
        out["dd_dt"] = td_exact(p2, ["dd", "d", "dt", "dd=dt", "dd = dt"]) or out["e"]
        out["Clase"] = td_exact(p2, ["Clase"])
    finally:
        try:
            p2.close()
        except Exception:
            pass
    return out


# ==============================
# 6) Extracción integral por OT
# ==============================
def extraer_camiones_por_ot(ot_numero: str, usuario: str, password: str, mostrar_navegador: bool = True) -> pd.DataFrame:
    filas: List[Dict[str, str]] = []

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=not mostrar_navegador, slow_mo=DEFAULT_SLOW_MO_MS)
        context = browser.new_context()
        page, vpe_text = login_y_abrir_ot(context, usuario, password, ot_numero)

        html = page.content()
        ids = re.findall(r'name="instrumentos\[\d+\]\.idInstrumento"\s+value="(\d+)"', html)
        if not ids:
            raise RuntimeError("No se encontraron instrumentos en el detalle del VPE.")

        for idx, id_inst in enumerate(ids, start=1):
            det = context.new_page()
            try:
                det.goto(
                    f"https://app.inti.gob.ar/MetroWeb/instrumentoDetalle.do?idInstrumento={id_inst}",
                    timeout=DEFAULT_WAIT_MS,
                )
                det.wait_for_load_state("networkidle", timeout=DEFAULT_WAIT_MS)

                gen = leer_generales(det)

                # Receptor (línea y link a modelo por si sirve luego)
                rec_linea = _sec_row_value(
                    det, "Parte base del Instrumento", "Modelo - Tipo de Instrumento - Marca - Fabricante"
                )
                rec_href_modelo = _sec_row_link(
                    det, "Parte base del Instrumento", "Modelo - Tipo de Instrumento - Marca - Fabricante"
                )
                rec_cod_ap = _sec_row_value(det, "Parte base del Instrumento", "Código de Aprobación de Modelo")
                rec_serie = _sec_row_value(det, "Parte base del Instrumento", "Nro de serie")
                parsed_rec = _parse_modelo_tipo_marca_fabricante(rec_linea)
                md_rec = leer_modelo_detalle(context, rec_href_modelo)

                # Indicador
                ind_linea = _sec_row_value(
                    det, "Indicador Electrónico", "Modelo - Tipo de Instrumento - Marca - Fabricante"
                )
                ind_href_modelo = _sec_row_link(
                    det, "Indicador Electrónico", "Modelo - Tipo de Instrumento - Marca - Fabricante"
                )
                ind_cod_ap = _sec_row_value(det, "Indicador Electrónico", "Código de Aprobación de Modelo")
                ind_serie = _sec_row_value(det, "Indicador Electrónico", "Nro de serie")
                parsed_ind = _parse_modelo_tipo_marca_fabricante(ind_linea)
                md_ind = leer_modelo_detalle(context, ind_href_modelo)

                fila = {
                    # Generales (lo que pediste corregir)
                    "OT": ot_numero,
                    "VPE": _solo_digitos(vpe_text),
                    "Fecha verificación": gen["Fecha_verificacion"],
                    "Razón Social (Propietario)": gen["Razon_social_propietario"],  # <-- Nombre del Usuario del Instrumento
                    "CUIT": gen["CUIT"],
                    "Dirección Legal (Fiscal)": gen["Direccion_legal"],            # <-- Dirección Legal (texto completo)
                    "Fiscal - Domicilio": gen["Fiscal_domicilio"],                 # (igual a Dirección Legal por ahora)
                    "Fiscal - Localidad": gen["Fiscal_localidad"],                 # (vacío por pedido)
                    "Fiscal - Provincia": gen["Fiscal_provincia"],                 # (vacío por pedido)

                    # Instalación
                    "Instalación - Domicilio": gen["Instalacion_domicilio"],
                    "Instalación - Localidad": gen["Instalacion_localidad"],
                    "Instalación - Provincia": gen["Instalacion_provincia"],
                    "Tipo de Verificación": gen["Tipo_verificacion"],
                    "Tolerancia": gen["Tolerancia"],

                    # Receptor
                    "Fabricante receptor": md_rec["Fabricante"] or parsed_rec.get("Fabricante/Importador", ""),
                    "Marca Receptor": md_rec["Marca"] or parsed_rec.get("Marca", ""),
                    "Modelo Receptor": md_rec["Modelo"] or parsed_rec.get("Modelo", ""),
                    "N° de serie Receptor": rec_serie,
                    "Cód ap. mod. Receptor": rec_cod_ap or md_rec["CodigoAprob"],
                    "Origen Receptor": md_rec["PaisOrigen"],

                    # Metrológicas (receptor)
                    "e": md_rec["e"],
                    "máx": md_rec["Max"],
                    "mín": md_rec["Min"],
                    "dd=dt": md_rec["dd_dt"],
                    "clase": md_rec["Clase"],

                    # Indicador
                    "Marca Indicador": md_ind["Marca"] or parsed_ind.get("Marca", ""),
                    "Modelo Indicador": md_ind["Modelo"] or parsed_ind.get("Modelo", ""),
                    "N° de serie Indicador": ind_serie,
                    "Código ap. mod. Indicador": ind_cod_ap or md_ind["CodigoAprob"],
                    "Origen Indicador": md_ind["PaisOrigen"],
                    "N° Aprobación Modelo (Ind.)": md_ind["NroDisposicion"],
                    "Fecha Aprobación Modelo (Ind.)": md_ind["FechaAprob"],
                }

                filas.append(fila)
                print(
                    f"✅ Instrumento {idx} capturado (Receptor {fila['Marca Receptor'] or '-'} / Indicador {fila['Marca Indicador'] or '-'})"
                )
            finally:
                try:
                    det.close()
                except Exception:
                    pass

        browser.close()

    cols = [
        "OT",
        "VPE",
        "Fecha verificación",
        "Razón Social (Propietario)",
        "CUIT",
        "Dirección Legal (Fiscal)",
        "Fiscal - Domicilio",
        "Fiscal - Localidad",
        "Fiscal - Provincia",
        "Instalación - Domicilio",
        "Instalación - Localidad",
        "Instalación - Provincia",
        "Tipo de Verificación",
        "Tolerancia",
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
        "Marca Indicador",
        "Modelo Indicador",
        "N° de serie Indicador",
        "Código ap. mod. Indicador",
        "Origen Indicador",
        "N° Aprobación Modelo (Ind.)",
        "Fecha Aprobación Modelo (Ind.)",
    ]
    df = pd.DataFrame(filas)
    for c in cols:
        if c not in df.columns:
            df[c] = ""
    return df[cols]


# ==============================
# 7) Hoja de verificación (sin columnas día/mes/año)
# ==============================
def armar_hoja_verificacion(df: pd.DataFrame) -> pd.DataFrame:
    out = pd.DataFrame(
        {
            "Número de O.T.": df["OT"],
            "VPE Nº": df["VPE"],
            # (a pedido: sin día/mes/año; van manual en el Excel final)
            "Razón social (Propietario)": df["Razón Social (Propietario)"],
            "Domicilio (Fiscal)": df["Dirección Legal (Fiscal)"],  # <-- tal cual Dirección Legal
            "Localidad (Fiscal)": df["Fiscal - Localidad"],
            "Provincia (Fiscal)": df["Fiscal - Provincia"],
            "Lugar propio de instalación - Domicilio": df["Instalación - Domicilio"],
            "Lugar propio de instalación - Localidad": df["Instalación - Localidad"],
            "Lugar propio de instalación - Provincia": df["Instalación - Provincia"],
            "Instrumento verificado": "Balanza para pesar camiones",
            "Fabricante receptor": df["Fabricante receptor"],
            "Marca Receptor": df["Marca Receptor"],
            "Modelo Receptor": df["Modelo Receptor"],
            "N° de serie Receptor": df["N° de serie Receptor"],
            "Cód ap. mod. Receptor": df["Cód ap. mod. Receptor"],
            "Origen Receptor": df["Origen Receptor"],
            "e": df["e"],
            "máx": df["máx"],
            "mín": df["mín"],
            "dd=dt": df["dd=dt"],
            "clase": df["clase"],
            "Tipo (Indicador)": "electrónica",
            "Marca Indicador": df["Marca Indicador"],
            "Modelo Indicador": df["Modelo Indicador"],
            "N° de serie Indicador": df["N° de serie Indicador"],
            "Código Aprobación (Indicador)": df["Código ap. mod. Indicador"],
            "Origen Indicador": df["Origen Indicador"],
            "N° de Aprobación Modelo (Indicador)": df["N° Aprobación Modelo (Ind.)"],
            "Fecha de Aprobación Modelo (Indicador)": df["Fecha Aprobación Modelo (Ind.)"],
            "Tipo de Verificación": df["Tipo de Verificación"],
            "Tolerancia": df["Tolerancia"],
            "CUIT del solicitante": df["CUIT"],
        }
    )
    return out


def exportar_verificacion(df_verif: pd.DataFrame, ruta_xlsx: str) -> str:
    if not ruta_xlsx.lower().endswith(".xlsx"):
        ruta_xlsx += ".xlsx"
    with pd.ExcelWriter(ruta_xlsx, engine="xlsxwriter") as writer:
        df_verif.to_excel(writer, sheet_name="Verificación", index=False)
        ws = writer.sheets["Verificación"]
        for i, _ in enumerate(df_verif.columns):
            ws.set_column(i, i, 28)
    return ruta_xlsx


# ==============================
# CLI
# ==============================
def main():
    ap = argparse.ArgumentParser(description="Extracción MetroWeb (camiones/plataforma) → hoja de verificación XLSX")
    ap.add_argument("--user", help="Usuario MetroWeb")
    ap.add_argument("--pass", dest="pwd", help="Contraseña MetroWeb")
    ap.add_argument("--ot", help="Número de OT (ej. 307-62136)")
    ap.add_argument("--headless", action="store_true", help="No mostrar el navegador (Chromium headless)")
    ap.add_argument("--out", default=None, help="Ruta de salida .xlsx (por defecto: OT_<OT>_VERIFICACION_PREVIA.xlsx)")
    args = ap.parse_args()

    usuario = args.user or input("Usuario INTI: ").strip()
    password = args.pwd or getpass.getpass("Contraseña INTI: ").strip()
    ot_numero = args.ot or input("N° de OT (formato 307-xxxxx): ").strip()

    df_cam = extraer_camiones_por_ot(ot_numero, usuario, password, mostrar_navegador=not args.headless)
    df_ver = armar_hoja_verificacion(df_cam)

    out_path = args.out or f"OT_{ot_numero}_VERIFICACION_PREVIA.xlsx"
    ruta = exportar_verificacion(df_ver, out_path)
    print("Archivo generado:", ruta)


if __name__ == "__main__":
    main()
