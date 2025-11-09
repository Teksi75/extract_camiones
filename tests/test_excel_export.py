import pandas as pd
from pathlib import Path
from src.io.excel_exporter import exportar

def test_exportar_crea_archivo(tmp_path):
    df = pd.DataFrame({"Campo": ["A"], "Valor": ["B"]})
    ruta = tmp_path / "salida.xlsx"
    result = exportar(df, ruta)
    assert result.exists()
