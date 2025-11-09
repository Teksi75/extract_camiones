# -*- coding: utf-8 -*-
"""
Exportación a Excel (formato 2 columnas: Campo | Valor)

Provee:
- armar_hoja_verificacion_2columnas(...)
- exportar_verificacion_2columnas(...)

Además, da formato castellano a:
- "Fecha de Aprobación Modelo (Receptor)"
- "Fecha de Aprobación Modelo (Indicador)"

Ej.: "22/04/1997" -> "22 de abril de 1997"
"""

from __future__ import annotations

from pathlib import Path
from typing import Dict, List, Optional
from datetime import datetime

import pandas as pd

# Orden estándar de columnas para construir la hoja
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

# Campos de fecha a normalizar en castellano
FIELD_FECHA_RECEPTOR = "Fecha de Aprobación Modelo (Receptor)"
FIELD_FECHA_INDICADOR = "Fecha de Aprobación Modelo (Indicador)"
DATE_FIELDS = {FIELD_FECHA_RECEPTOR, FIELD_FECHA_INDICADOR}

# Mapa de meses en castellano (en minúsculas)
_MESES = {
    1: "enero", 2: "febrero", 3: "marzo", 4: "abril",
    5: "mayo", 6: "junio", 7: "julio", 8: "agosto",
    9: "septiembre", 10: "octubre", 11: "noviembre", 12: "diciembre",
}

def _es_formato_castellano(s: str) -> bool:
    """Detecta si ya está en formato '22 de abril de 1997'."""
    return " de " in s and any(m in s.lower() for m in _MESES.values())

def _parse_date(value: str) -> Optional[datetime]:
    """
    Intenta parsear fechas típicas del portal:
    - 22/04/1997
    - 22-04-1997
    - 22.04.1997
    - 22/4/1997 (sin cero a la izquierda)
    - 1997-04-22
    Si no puede, devuelve None.
    """
    if not value:
        return None
    value = value.strip()
    # Casos ya “bien” (castellano), no parseamos.
    if _es_formato_castellano(value):
        return None

    # Intentos más comunes
    formatos = [
        "%d/%m/%Y",
        "%d-%m-%Y",
        "%d.%m.%Y",
        "%Y-%m-%d",
        "%d/%m/%y",
        "%d-%m-%y",
    ]
    for fmt in formatos:
        try:
            return datetime.strptime(value, fmt)
        except ValueError:
            continue

    # Intento flexible: normalizar separadores a "/"
    import re
    dig = re.findall(r"\d+", value)
    # Esperamos 3 grupos: d, m, Y
    if len(dig) == 3:
        d, m, y = dig
        # Corrige años tipo "97" -> "1997" (heurística simple)
        if len(y) == 2:
            y = ("19" if int(y) >= 50 else "20") + y
        try:
            return datetime(int(y), int(m), int(d))
        except ValueError:
            return None
    return None

def _fecha_castellano(value: str) -> str:
    """
    Devuelve la fecha en formato '22 de abril de 1997'.
    Si no se puede parsear, devuelve el valor original.
    """
    if not value:
        return value
    if _es_formato_castellano(value):
        # Ya está en el formato deseado.
        return value

    dt = _parse_date(value)
    if not dt:
        return value

    mes = _MESES.get(dt.month, "")
    # Día sin cero a la izquierda
    return f"{dt.day} de {mes} de {dt.year}"

def armar_hoja_verificacion_2columnas(filas: List[Dict[str, str]]) -> pd.DataFrame:
    """
    Convierte las filas (un dict por instrumento) a formato 2 columnas (Campo|Valor).
    Aplica formato castellano a los campos de fecha definidos.
    """
    if not filas:
        return pd.DataFrame(columns=["Campo", "Valor"])

    data_final: List[Dict[str, str]] = []
    for idx, fila in enumerate(filas, start=1):
        if idx > 1:
            data_final.append({"Campo": "", "Valor": ""})
            data_final.append({"Campo": f"=== INSTRUMENTO {idx} ===", "Valor": ""})

        for col in COLUMNS_ORDER:
            val = fila.get(col, "")

            # Normaliza fechas específicas al castellano
            if col in DATE_FIELDS:
                val = _fecha_castellano(str(val))

            data_final.append({"Campo": col, "Valor": val})

    return pd.DataFrame(data_final)

def exportar_verificacion_2columnas(df: pd.DataFrame, ruta: Path) -> Path:
    """
    Exporta el DataFrame formateado con estilos básicos y anchos adecuados.
    Se conserva el tipo TEXTO de las fechas (no aplica formato numérico).
    """
    ruta = Path(ruta)
    ruta.parent.mkdir(parents=True, exist_ok=True)

    with pd.ExcelWriter(ruta, engine="xlsxwriter") as writer:
        df.to_excel(writer, sheet_name="Verificación", index=False)
        wb = writer.book
        ws = writer.sheets["Verificación"]

        fmt_header = wb.add_format(
            {"bold": True, "bg_color": "#4472C4", "font_color": "white", "border": 1,
             "align": "center", "valign": "vcenter", "font_size": 11}
        )
        fmt_campo = wb.add_format(
            {"bold": True, "bg_color": "#4472C4", "font_color": "white", "border": 1,
             "align": "left", "valign": "vcenter", "text_wrap": True}
        )
        fmt_valor = wb.add_format({"border": 1, "align": "left", "valign": "top", "text_wrap": True})
        fmt_sep = wb.add_format({"bold": True, "bg_color": "#FFC000", "font_color": "#000000",
                                 "border": 1, "align": "center", "valign": "vcenter"})

        # Encabezados
        ws.write(0, 0, "Campo", fmt_header)
        ws.write(0, 1, "Valor", fmt_header)

        # Cuerpo
        for row_num in range(1, len(df) + 1):
            campo = df.iloc[row_num - 1, 0]
            valor = df.iloc[row_num - 1, 1]

            # Aseguramos string seguro para Excel (evitar fórmulas accidentales)
            campo_str = "" if campo is None else str(campo)
            valor_str = "" if valor is None else str(valor)

            if isinstance(campo_str, str) and campo_str.startswith("==="):
                # Separador entre instrumentos: escribir como TEXTO (no fórmula)
                ws.write_string(row_num, 0, campo_str, fmt_sep)
                ws.write_string(row_num, 1, valor_str, fmt_sep)
            elif campo_str == "":
                ws.write_string(row_num, 0, "", fmt_valor)
                ws.write_string(row_num, 1, "", fmt_valor)
            else:
                # Campo/Valor normales: siempre como string para evitar fórmulas
                ws.write_string(row_num, 0, campo_str, fmt_campo)
                ws.write_string(row_num, 1, valor_str, fmt_valor)


        # Ajustes de presentación
        ws.set_column(0, 0, 45)  # Campo
        ws.set_column(1, 1, 60)  # Valor
        ws.freeze_panes(1, 0)
        ws.set_row(0, 25)

    return ruta
