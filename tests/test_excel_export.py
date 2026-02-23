import pandas as pd

from src.io.excel_exporter import (
    armar_hoja_verificacion_2columnas,
    exportar_verificacion_2columnas,
)


def test_exportar_verificacion_2columnas_crea_archivo(workspace_tmp_path):
    df = pd.DataFrame({"Campo": ["A"], "Valor": ["B"]})
    ruta = workspace_tmp_path / "salida.xlsx"
    result = exportar_verificacion_2columnas(df, ruta)
    assert result.exists()


def test_armar_hoja_verificacion_2columnas_agrega_separador():
    filas = [
        {"Número de O.T.": "307-11111", "VPE Nº": "123"},
        {"Número de O.T.": "307-11111", "VPE Nº": "124"},
    ]

    df = armar_hoja_verificacion_2columnas(filas)
    assert "=== INSTRUMENTO 2 ===" in df["Campo"].tolist()
