import pandas as pd

from src.ui.gui import exportar_excel_detalle_instrumentos


def test_exportar_excel_detalle_instrumentos_omite_archivo_si_hay_un_instrumento(
    workspace_tmp_path,
):
    ruta_principal = workspace_tmp_path / "resultado.xlsx"
    ruta_principal.touch()
    filas = [{"Número de O.T.": "307-11111", "VPE Nº": "123"}]

    result = exportar_excel_detalle_instrumentos(filas, ruta_principal)

    assert result is None


def test_exportar_excel_detalle_instrumentos_crea_archivo_adicional(
    workspace_tmp_path,
):
    ruta_principal = workspace_tmp_path / "resultado.xlsx"
    ruta_principal.touch()
    filas = [
        {"Número de O.T.": "307-11111", "VPE Nº": "123", "Modelo Receptor": "M1"},
        {"Número de O.T.": "307-11111", "VPE Nº": "124", "Modelo Receptor": "M2"},
    ]

    result = exportar_excel_detalle_instrumentos(filas, ruta_principal)

    assert result is not None
    assert result.exists()
    assert result.name == "resultado_instrumentos.xlsx"

    df = pd.read_excel(result, sheet_name="Verificación")
    assert "=== INSTRUMENTO 2 ===" in df["Campo"].astype(str).tolist()
