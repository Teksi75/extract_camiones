# tests/test_address.py
from src.domain.address import parse_domicilio_fiscal

def test_parse_mendoza():
    d, l, p = parse_domicilio_fiscal("Suyai 2632 Luján de Cuyo Mendoza")
    assert l == "Luján de Cuyo"
    assert p == "Mendoza"

def test_parse_caba_largo():
    d, l, p = parse_domicilio_fiscal("Av. Rivadavia 12345 Caballito Ciudad Autónoma de Buenos Aires")
    assert l == "Caballito"
    assert "Buenos Aires" in p
